import os
import logging
import imaplib
import smtplib
import email
from email.header import decode_header
from email.mime.text import MIMEText

from telegram import Update
from telegram.ext import (
    Application,
    CommandHandler,
    MessageHandler,
    ContextTypes,
    filters,
)
from anthropic import Anthropic

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)

TELEGRAM_TOKEN = os.environ["TELEGRAM_TOKEN"]
ANTHROPIC_API_KEY = os.environ["ANTHROPIC_API_KEY"]
ALLOWED_USER_ID = int(os.environ["ALLOWED_USER_ID"])

GMAIL_ADDRESS = os.environ.get("GMAIL_ADDRESS")
GMAIL_APP_PASSWORD = os.environ.get("GMAIL_APP_PASSWORD")
MAIL_CHECK_INTERVAL = int(os.environ.get("MAIL_CHECK_INTERVAL", "120"))

MODEL_NAME = os.environ.get("MODEL_NAME", "claude-haiku-4-5-20251001")

client = Anthropic(api_key=ANTHROPIC_API_KEY)

SYSTEM_PROMPT = (
    "당신은 사용자의 개인 AI 비서입니다. 한국어로 친절하고 간결하게 답변하세요. "
    "불필요하게 길게 설명하지 말고, 핵심 위주로 답하세요."
)

conversations: dict[int, list[dict]] = {}
MAX_HISTORY = 20

# 메일 감시용: 마지막으로 확인한 IMAP UID (재시작하면 초기화됨)
last_uid_seen: int | None = None


def is_allowed(update: Update) -> bool:
    user = update.effective_user
    return user is not None and user.id == ALLOWED_USER_ID


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_allowed(update):
        await update.message.reply_text("죄송해요, 이 봇은 개인용이라 사용할 수 없어요.")
        return
    await update.message.reply_text(
        "안녕하세요! AI 비서예요. 👋\n\n"
        "그냥 편하게 메시지를 보내면 대화할 수 있어요.\n\n"
        "사용 가능한 명령어:\n"
        "/remind <분> <내용> - 알림 예약 (예: /remind 30 회의 참석)\n"
        "/mail <받는사람이메일> - 메일 작성해서 보내기\n"
        "/reset - 지금까지의 대화 기억 지우기\n\n"
        "새 이메일이 오면 자동으로 요약해서 알려드려요. 📬"
    )


async def reset(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_allowed(update):
        return
    conversations.pop(update.effective_chat.id, None)
    await update.message.reply_text("대화 기록을 초기화했어요.")


async def remind(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_allowed(update):
        return

    try:
        minutes = float(context.args[0])
        text = " ".join(context.args[1:]).strip()
        if not text:
            raise ValueError
    except (IndexError, ValueError):
        await update.message.reply_text(
            "사용법: /remind <분> <내용>\n예: /remind 30 회의 참석"
        )
        return

    chat_id = update.effective_chat.id
    context.job_queue.run_once(
        send_reminder, when=minutes * 60, chat_id=chat_id, data=text
    )
    await update.message.reply_text(f"⏰ {minutes}분 후에 알려드릴게요: {text}")


async def send_reminder(context: ContextTypes.DEFAULT_TYPE) -> None:
    job = context.job
    await context.bot.send_message(chat_id=job.chat_id, text=f"⏰ 알림: {job.data}")


def _send_email(to_addr: str, subject: str, body: str) -> None:
    msg = MIMEText(body)
    msg["Subject"] = subject
    msg["From"] = GMAIL_ADDRESS
    msg["To"] = to_addr

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(GMAIL_ADDRESS, GMAIL_APP_PASSWORD)
        server.sendmail(GMAIL_ADDRESS, [to_addr], msg.as_string())


async def mail_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_allowed(update):
        return

    if not GMAIL_ADDRESS or not GMAIL_APP_PASSWORD:
        await update.message.reply_text("메일 발송 기능이 설정되어 있지 않아요.")
        return

    if not context.args:
        await update.message.reply_text(
            "사용법: /mail <받는사람이메일>\n예: /mail friend@example.com"
        )
        return

    to_addr = context.args[0]
    context.user_data["mail_draft"] = {"to": to_addr, "state": "subject"}
    await update.message.reply_text(f"받는사람: {to_addr}\n메일 제목을 입력해주세요.")


async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_allowed(update):
        await update.message.reply_text("죄송해요, 이 봇은 개인용이라 사용할 수 없어요.")
        return

    chat_id = update.effective_chat.id
    user_text = update.message.text

    # 메일 작성 중이면(제목/내용/확인 단계) 일반 대화 대신 메일 작성 흐름으로 처리
    draft = context.user_data.get("mail_draft")
    if draft:
        state = draft["state"]

        if state == "subject":
            draft["subject"] = user_text
            draft["state"] = "body"
            await update.message.reply_text("메일 내용을 입력해주세요.")
            return

        if state == "body":
            draft["body"] = user_text
            draft["state"] = "confirm"
            preview = (
                f"받는사람: {draft['to']}\n"
                f"제목: {draft['subject']}\n\n"
                f"{draft['body']}\n\n"
                "이대로 보낼까요? '네' 또는 '아니오'로 답해주세요."
            )
            await update.message.reply_text(preview)
            return

        if state == "confirm":
            context.user_data.pop("mail_draft", None)
            if user_text.strip() in ("네", "예", "ㅇㅇ", "y", "yes", "Y"):
                try:
                    _send_email(draft["to"], draft["subject"], draft["body"])
                    await update.message.reply_text("✅ 메일을 보냈어요!")
                except Exception:
                    logger.exception("메일 발송 중 오류")
                    await update.message.reply_text("❌ 메일 발송에 실패했어요. 잠시 후 다시 시도해주세요.")
            else:
                await update.message.reply_text("메일 발송을 취소했어요.")
            return

    history = conversations.setdefault(chat_id, [])
    history.append({"role": "user", "content": user_text})
    history = history[-MAX_HISTORY:]

    await context.bot.send_chat_action(chat_id=chat_id, action="typing")

    try:
        response = client.messages.create(
            model=MODEL_NAME,
            max_tokens=1024,
            system=SYSTEM_PROMPT,
            messages=history,
        )
        reply_text = response.content[0].text
    except Exception:
        logger.exception("Claude API 호출 중 오류")
        reply_text = "죄송해요, 답변을 만드는 중에 오류가 발생했어요. 잠시 후 다시 시도해주세요."

    history.append({"role": "assistant", "content": reply_text})
    conversations[chat_id] = history

    await update.message.reply_text(reply_text)


def _decode_mime_words(s: str) -> str:
    if not s:
        return ""
    decoded = decode_header(s)
    return "".join(
        (t.decode(enc or "utf-8", errors="ignore") if isinstance(t, bytes) else t)
        for t, enc in decoded
    )


def _get_email_body(msg: email.message.Message) -> str:
    if msg.is_multipart():
        for part in msg.walk():
            ctype = part.get_content_type()
            disp = str(part.get("Content-Disposition"))
            if ctype == "text/plain" and "attachment" not in disp:
                charset = part.get_content_charset() or "utf-8"
                try:
                    return part.get_payload(decode=True).decode(charset, errors="ignore")
                except Exception:
                    return ""
        return ""
    else:
        charset = msg.get_content_charset() or "utf-8"
        try:
            return msg.get_payload(decode=True).decode(charset, errors="ignore")
        except Exception:
            return ""


async def check_new_mail(context: ContextTypes.DEFAULT_TYPE) -> None:
    global last_uid_seen

    if not GMAIL_ADDRESS or not GMAIL_APP_PASSWORD:
        return

    try:
        imap = imaplib.IMAP4_SSL("imap.gmail.com")
        imap.login(GMAIL_ADDRESS, GMAIL_APP_PASSWORD)
        imap.select("INBOX")

        status, data = imap.uid("search", None, "ALL")
        uids = data[0].split()
        if not uids:
            imap.logout()
            return

        latest_uid = int(uids[-1])

        if last_uid_seen is None:
            # 처음 실행될 때는 지금 시점만 기준으로 잡고, 과거 메일은 알리지 않음
            last_uid_seen = latest_uid
            imap.logout()
            return

        new_uids = [uid for uid in uids if int(uid) > last_uid_seen]

        for uid in new_uids:
            status, msg_data = imap.uid("fetch", uid, "(RFC822)")
            if not msg_data or msg_data[0] is None:
                continue
            raw_email = msg_data[0][1]
            msg = email.message_from_bytes(raw_email)

            subject = _decode_mime_words(msg.get("Subject", "(제목 없음)"))
            sender = _decode_mime_words(msg.get("From", "(발신자 알 수 없음)"))
            body = _get_email_body(msg)[:2000]

            try:
                response = client.messages.create(
                    model=MODEL_NAME,
                    max_tokens=300,
                    system="이메일 내용을 한국어로 3줄 이내로 간결하게 요약해줘. 핵심만 전달해.",
                    messages=[
                        {
                            "role": "user",
                            "content": f"보낸사람: {sender}\n제목: {subject}\n본문:\n{body}",
                        }
                    ],
                )
                summary = response.content[0].text
            except Exception:
                logger.exception("메일 요약 중 오류")
                summary = "(요약 생성 실패)"

            text = f"📬 새 메일 도착\n\n보낸사람: {sender}\n제목: {subject}\n\n요약:\n{summary}"
            await context.bot.send_message(chat_id=ALLOWED_USER_ID, text=text)

        last_uid_seen = latest_uid
        imap.logout()
    except Exception:
        logger.exception("메일 확인 중 오류")


def main() -> None:
    app = Application.builder().token(TELEGRAM_TOKEN).build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("reset", reset))
    app.add_handler(CommandHandler("remind", remind))
    app.add_handler(CommandHandler("mail", mail_command))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))

    if GMAIL_ADDRESS and GMAIL_APP_PASSWORD:
        app.job_queue.run_repeating(
            check_new_mail, interval=MAIL_CHECK_INTERVAL, first=10
        )
        logger.info("이메일 확인 작업이 등록되었습니다 (%d초 간격).", MAIL_CHECK_INTERVAL)
    else:
        logger.info("GMAIL_ADDRESS/GMAIL_APP_PASSWORD가 없어 이메일 확인 기능은 꺼져 있습니다.")

    logger.info("봇을 시작합니다...")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
