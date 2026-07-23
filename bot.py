import os
import logging

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

MODEL_NAME = os.environ.get("MODEL_NAME", "claude-haiku-4-5-20251001")

client = Anthropic(api_key=ANTHROPIC_API_KEY)

SYSTEM_PROMPT = (
    "당신은 사용자의 개인 AI 비서입니다. 한국어로 친절하고 간결하게 답변하세요. "
    "불필요하게 길게 설명하지 말고, 핵심 위주로 답하세요."
)

conversations: dict[int, list[dict]] = {}
MAX_HISTORY = 20


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
        "/reset - 지금까지의 대화 기억 지우기"
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


async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_allowed(update):
        await update.message.reply_text("죄송해요, 이 봇은 개인용이라 사용할 수 없어요.")
        return

    chat_id = update.effective_chat.id
    user_text = update.message.text

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


def main() -> None:
    app = Application.builder().token(TELEGRAM_TOKEN).build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("reset", reset))
    app.add_handler(CommandHandler("remind", remind))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))

    logger.info("봇을 시작합니다...")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
