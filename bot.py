import os
import json
import asyncio
import logging
import imaplib
import smtplib
import email
from email.header import decode_header
from email.mime.text import MIMEText

import httpx
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

# 메일 보낼 때 "보낸사람"으로 표시할 주소. 지정 안 하면 GMAIL_ADDRESS 그대로 사용.
# (구글 계정에 "다른 이메일 주소로 보내기" 별칭으로 인증된 주소여야 정상 발송돼요)
MAIL_FROM_ADDRESS = os.environ.get("MAIL_FROM_ADDRESS", GMAIL_ADDRESS)

MODEL_NAME = os.environ.get("MODEL_NAME", "claude-haiku-4-5-20251001")

# 카카오 REST API 키 (카카오맵 대중교통/도보/자전거 길찾기 + 카카오모빌리티 자동차 길찾기 공용)
KAKAO_REST_API_KEY = os.environ.get("KAKAO_REST_API_KEY")

client = Anthropic(api_key=ANTHROPIC_API_KEY)

# 집/회사 등 저장된 위치를 담아두는 파일 (재시작해도 남아있도록 디스크에 저장)
PLACES_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "places.json")

SYSTEM_PROMPT = (
    "당신은 사용자의 개인 AI 비서입니다. 한국어로 친절하고 간결하게 답변하세요. "
    "불필요하게 길게 설명하지 말고, 핵심 위주로 답하세요. "
    "날씨, 최신 뉴스, 맛집, 가격 등 최신 정보가 필요한 질문은 웹 검색 도구를 적극 활용하세요."
)

CHAT_TOOLS = [
    {"type": "web_search_20250305", "name": "web_search", "max_uses": 3}
]

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
        "/route - 두 지점 사이 길찾기 (자동차/대중교통/도보/자전거 비교)\n"
        "/sethome - 집 위치 등록\n"
        "/setwork - 회사 위치 등록\n"
        "/towork - 집→회사 길찾기 (등록 필요)\n"
        "/tohome - 회사→집 길찾기 (등록 필요)\n"
        "/reset - 지금까지의 대화 기억 지우기\n\n"
        "길찾기는 명령어 없이 그냥 '강남역까지 얼마나 걸려?', '홍대에서 여의도까지 어떻게 가?'처럼 물어보셔도 알아들어요.\n\n"
        "새 이메일이 오면 자동으로 요약해서 알려드려요. 📬\n"
        "날씨, 최신 뉴스, 맛집 등도 그냥 물어보시면 웹 검색해서 답해드려요.\n"
        "📎(첨부) 버튼으로 '위치'를 공유해주시면, 그 위치 기준으로 근처 맛집도 찾아드려요."
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


async def _reverse_geocode(lat: float, lon: float) -> str:
    url = "https://nominatim.openstreetmap.org/reverse"
    params = {"lat": lat, "lon": lon, "format": "json", "accept-language": "ko"}
    headers = {"User-Agent": "yujin-ai-bot/1.0"}
    async with httpx.AsyncClient(timeout=10) as http_client:
        resp = await http_client.get(url, params=params, headers=headers)
        resp.raise_for_status()
        data = resp.json()
    return data.get("display_name") or f"위도 {lat}, 경도 {lon}"


def _load_places() -> dict:
    try:
        with open(PLACES_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return {}


def _save_places(places: dict) -> None:
    with open(PLACES_FILE, "w", encoding="utf-8") as f:
        json.dump(places, f, ensure_ascii=False, indent=2)


async def _kakao_geocode_keyword(query: str) -> dict | None:
    """주소나 장소명을 좌표로 변환 (카카오맵 키워드 검색)"""
    if not KAKAO_REST_API_KEY:
        return None
    url = "https://dapi.kakao.com/v2/local/search/keyword.json"
    headers = {"Authorization": f"KakaoAK {KAKAO_REST_API_KEY}"}
    params = {"query": query, "size": 1}
    async with httpx.AsyncClient(timeout=10) as http_client:
        resp = await http_client.get(url, headers=headers, params=params)
        resp.raise_for_status()
        data = resp.json()
    docs = data.get("documents") or []
    if not docs:
        return None
    d = docs[0]
    return {
        "name": d.get("place_name") or d.get("road_address_name") or d.get("address_name"),
        "address": d.get("road_address_name") or d.get("address_name"),
        "x": d["x"],
        "y": d["y"],
    }


async def _extract_place_name_llm(text: str) -> str | None:
    """문장에서 장소명/주소만 뽑아냄 ('현재 위치는 강남역이야' -> '강남역')"""
    system = (
        "사용자 문장에서 언급된 장소명이나 주소만 정확히 추출해서 그 텍스트만 답하세요. "
        "다른 설명, 문장부호, 따옴표 없이 장소명만 출력하세요. "
        "장소를 찾을 수 없으면 정확히 NONE 이라고만 답하세요."
    )
    try:
        response = client.messages.create(
            model=MODEL_NAME,
            max_tokens=30,
            system=system,
            messages=[{"role": "user", "content": text}],
        )
        raw = "".join(b.text for b in response.content if b.type == "text").strip()
        if not raw or raw.upper() == "NONE":
            return None
        return raw
    except Exception:
        logger.exception("장소명 추출 중 오류")
        return None


async def _resolve_place(text: str, places: dict) -> dict | None:
    """사용자가 입력한 텍스트를 위치 정보로 변환. '집'/'회사'는 저장된 위치 사용"""
    t = text.strip()
    if t in ("집", "우리집", "집으로", "자택"):
        return places.get("home")
    if t in ("회사", "직장", "사무실", "회사로"):
        return places.get("work")
    try:
        place = await _kakao_geocode_keyword(t)
        if place:
            return place
        # 문장 형태로 입력한 경우("현재 위치는 강남역이야") 장소명만 뽑아서 재시도
        extracted = await _extract_place_name_llm(t)
        if extracted and extracted != t:
            return await _kakao_geocode_keyword(extracted)
        return None
    except Exception:
        logger.exception("장소 검색 중 오류")
        return None


async def _kakao_driving_route(sx: str, sy: str, ex: str, ey: str) -> dict | None:
    if not KAKAO_REST_API_KEY:
        return None
    url = "https://apis-navi.kakaomobility.com/v1/directions"
    headers = {"Authorization": f"KakaoAK {KAKAO_REST_API_KEY}"}
    params = {
        "origin": f"{sx},{sy}",
        "destination": f"{ex},{ey}",
        "priority": "RECOMMEND",
    }
    try:
        async with httpx.AsyncClient(timeout=10) as http_client:
            resp = await http_client.get(url, headers=headers, params=params)
            resp.raise_for_status()
            data = resp.json()
        routes = data.get("routes") or []
        if not routes or routes[0].get("result_code") != 0:
            return None
        return routes[0].get("summary")
    except Exception:
        logger.exception("자동차 길찾기 조회 중 오류")
        return None


async def _kakao_transit_route(sx: str, sy: str, ex: str, ey: str) -> dict | None:
    if not KAKAO_REST_API_KEY:
        return None
    url = "https://dapi.kakao.com/v2/routing/publictraffic"
    headers = {"Authorization": f"KakaoAK {KAKAO_REST_API_KEY}"}
    params = {"start_x": sx, "start_y": sy, "end_x": ex, "end_y": ey}
    try:
        async with httpx.AsyncClient(timeout=10) as http_client:
            resp = await http_client.get(url, headers=headers, params=params)
            resp.raise_for_status()
            return resp.json()
    except Exception:
        logger.exception("대중교통 길찾기 조회 중 오류")
        return None


async def _kakao_walk_or_bicycle_route(sx: str, sy: str, ex: str, ey: str, mode: str) -> dict | None:
    if not KAKAO_REST_API_KEY:
        return None
    path = "walk" if mode == "walk" else "bicycle"
    url = f"https://dapi.kakao.com/v2/routing/{path}"
    headers = {"Authorization": f"KakaoAK {KAKAO_REST_API_KEY}"}
    params = {"start_x": sx, "start_y": sy, "end_x": ex, "end_y": ey}
    try:
        async with httpx.AsyncClient(timeout=10) as http_client:
            resp = await http_client.get(url, headers=headers, params=params)
            resp.raise_for_status()
            return resp.json()
    except Exception:
        logger.exception("도보/자전거 길찾기 조회 중 오류")
        return None


def _format_driving(summary: dict | None) -> str:
    if not summary:
        return "🚗 자동차: 정보 없음 (승인 대기 중이거나 경로를 찾지 못했어요)"
    minutes = round(summary.get("duration", 0) / 60)
    km = round(summary.get("distance", 0) / 1000, 1)
    text = f"🚗 자동차: 약 {minutes}분 ({km}km)"
    toll = (summary.get("fare") or {}).get("toll")
    if toll:
        text += f", 통행료 {toll}원"
    return text


def _format_transit(data: dict | None) -> str:
    if not data or data.get("status") != "OK" or not data.get("routes"):
        return "🚌 대중교통: 정보 없음"
    route = data["routes"][0]
    props = route["properties"]
    minutes = round(props.get("totalTime", 0) / 60)
    line = f"🚌 대중교통: 약 {minutes}분"
    if props.get("transfers") is not None:
        line += f", 환승 {props['transfers']}회"
    fare = (props.get("fare") or {}).get("value")
    if fare:
        line += f", 요금 {fare}원"

    vehicle_names = []
    for step in route.get("steps", []):
        sp = step.get("properties", {})
        if sp.get("type") in ("BUS", "SUBWAY"):
            kind = "버스" if sp["type"] == "BUS" else "지하철"
            for v in sp.get("vehicles", []):
                label = f"{kind} {v.get('name', '')}".strip()
                if label not in vehicle_names:
                    vehicle_names.append(label)
    if vehicle_names:
        line += "\n   이용 노선: " + ", ".join(vehicle_names)
    return line


def _format_walk_or_bicycle(data: dict | None, label: str, emoji: str) -> str:
    if not data or data.get("status") != "OK":
        return f"{emoji} {label}: 정보 없음"
    props = data["route"]["properties"]
    minutes = round(props.get("totalTime", 0) / 60)
    km = round(props.get("totalDistance", 0) / 1000, 1)
    return f"{emoji} {label}: 약 {minutes}분 ({km}km)"


async def _reply_route_comparison(update: Update, start: dict, end: dict) -> None:
    start_label = start.get("name") or start.get("address") or "출발지"
    end_label = end.get("name") or end.get("address") or "도착지"

    if not KAKAO_REST_API_KEY:
        await update.message.reply_text("길찾기 기능이 아직 설정되지 않았어요.")
        return

    await update.message.reply_text(f"🔎 '{start_label}' → '{end_label}' 경로를 찾고 있어요...")

    sx, sy, ex, ey = start["x"], start["y"], end["x"], end["y"]

    driving, transit, walk, bicycle = await asyncio.gather(
        _kakao_driving_route(sx, sy, ex, ey),
        _kakao_transit_route(sx, sy, ex, ey),
        _kakao_walk_or_bicycle_route(sx, sy, ex, ey, "walk"),
        _kakao_walk_or_bicycle_route(sx, sy, ex, ey, "bicycle"),
    )

    lines = [
        f"📍 {start_label} → {end_label}",
        "",
        _format_driving(driving),
        _format_transit(transit),
        _format_walk_or_bicycle(walk, "도보", "🚶"),
        _format_walk_or_bicycle(bicycle, "자전거", "🚴"),
    ]
    await update.message.reply_text("\n".join(lines))


async def _advance_route_query(update: Update, context: ContextTypes.DEFAULT_TYPE, place: dict) -> None:
    """진행 중인 route_query 상태에 place를 반영하고 다음 단계로 진행 (텍스트 입력/위치 공유 공통 처리)"""
    route_query = context.user_data["route_query"]
    state = route_query["state"]
    label = place.get("name") or place.get("address") or "위치"

    if state == "await_current_location":
        dest_text = route_query.pop("pending_destination_text", None)
        if dest_text:
            places = _load_places()
            dest_place = await _resolve_place(dest_text, places)
            context.user_data.pop("route_query", None)
            if not dest_place:
                await update.message.reply_text(
                    f"현재 위치는 확인했는데, '{dest_text}'는 찾지 못했어요. 목적지를 다시 알려주세요."
                )
                context.user_data["route_query"] = {"state": "to", "from": place}
                return
            await _reply_route_comparison(update, place, dest_place)
        else:
            route_query["from"] = place
            route_query["state"] = "to"
            await update.message.reply_text(f"현재 위치: {label}\n\n도착지는 어디인가요?")
        return

    if state == "from":
        route_query["from"] = place
        route_query["state"] = "to"
        await update.message.reply_text(f"출발지: {label}\n\n도착지는 어디인가요?")
        return

    # state == "to"
    context.user_data.pop("route_query", None)
    await _reply_route_comparison(update, route_query["from"], place)


ROUTE_KEYWORDS = (
    "길찾", "가는 길", "가는길", "가는법", "가는 방법", "어떻게 가", "어떻게가",
    "얼마나 걸려", "얼마나 걸리", "출근길", "퇴근길", "이동시간", "차로 가", "대중교통으로",
    "버스로 가", "도보로 가", "걸어서 가",
)


async def _extract_route_intent(text: str) -> dict | None:
    """자유롭게 쓴 문장에서 길찾기 요청인지, 출발지/목적지가 무엇인지 뽑아냄"""
    system = (
        "사용자 문장이 길찾기(이동 경로) 요청인지 판단하고, 아래 JSON 형식으로만 답하세요. "
        "설명이나 다른 텍스트는 절대 포함하지 마세요.\n"
        '{"is_route": true 또는 false, '
        '"origin": "current"(현재 위치를 말하거나 출발지가 명시되지 않은 경우) 또는 "home" 또는 "work" 또는 구체적 장소명 문자열 또는 null, '
        '"destination": 구체적 장소명 문자열 또는 null(불명확한 경우)}\n'
        "길찾기 요청이 아니면 is_route를 false로 하세요."
    )
    try:
        response = client.messages.create(
            model=MODEL_NAME,
            max_tokens=200,
            system=system,
            messages=[{"role": "user", "content": text}],
        )
        raw = "".join(b.text for b in response.content if b.type == "text").strip()
        raw = raw.strip("`")
        if raw.lower().startswith("json"):
            raw = raw[4:].strip()
        data = json.loads(raw)
        if not isinstance(data, dict):
            return None
        return data
    except Exception:
        logger.exception("길찾기 의도 분석 중 오류")
        return None


async def _handle_natural_route_request(update: Update, context: ContextTypes.DEFAULT_TYPE, user_text: str) -> bool:
    """자연어 길찾기 요청을 감지해서 처리. 처리했으면 True 반환"""
    if not KAKAO_REST_API_KEY:
        return False
    if not any(kw in user_text for kw in ROUTE_KEYWORDS):
        return False

    intent = await _extract_route_intent(user_text)
    if not intent or not intent.get("is_route"):
        return False

    places = _load_places()
    origin = intent.get("origin")
    destination_text = intent.get("destination")

    # 출발지가 "현재 위치"를 가리키는 경우: 최근 공유된 위치가 있으면 재사용, 없으면 위치 공유 요청
    if origin in (None, "current"):
        recent = context.user_data.get("location")
        if isinstance(recent, dict) and recent.get("x") and recent.get("y"):
            from_place = {
                "name": recent.get("address"),
                "address": recent.get("address"),
                "x": recent["x"],
                "y": recent["y"],
            }
        else:
            context.user_data["route_query"] = {
                "state": "await_current_location",
                "pending_destination_text": destination_text,
            }
            extra = f" (목적지: {destination_text})" if destination_text else ""
            await update.message.reply_text(
                f"현재 위치를 확인할게요.{extra}\n📎(첨부) 버튼으로 위치를 공유해주세요."
            )
            return True
    else:
        from_place = await _resolve_place(origin, places)
        if not from_place:
            context.user_data["route_query"] = {"state": "from"}
            await update.message.reply_text(
                f"'{origin}' 위치를 찾지 못했어요. 출발지를 다시 알려주세요."
            )
            return True

    if not destination_text:
        context.user_data["route_query"] = {"state": "to", "from": from_place}
        await update.message.reply_text("도착지는 어디인가요?")
        return True

    to_place = await _resolve_place(destination_text, places)
    if not to_place:
        context.user_data["route_query"] = {"state": "to", "from": from_place}
        await update.message.reply_text(
            f"'{destination_text}'를 찾지 못했어요. 도착지를 다시 알려주세요."
        )
        return True

    await _reply_route_comparison(update, from_place, to_place)
    return True


async def route_start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_allowed(update):
        return
    if not KAKAO_REST_API_KEY:
        await update.message.reply_text("길찾기 기능이 아직 설정되지 않았어요.")
        return
    context.user_data["route_query"] = {"state": "from"}
    await update.message.reply_text(
        "출발지를 알려주세요.\n"
        "주소나 장소명을 입력하거나, 📎(첨부)로 위치를 공유하거나, '집' 또는 '회사'라고 입력해주세요."
    )


async def set_home(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_allowed(update):
        return
    context.user_data["saving_place"] = "home"
    await update.message.reply_text("🏠 집 위치를 설정할게요. 위치를 공유하거나 주소/건물명을 입력해주세요.")


async def set_work(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_allowed(update):
        return
    context.user_data["saving_place"] = "work"
    await update.message.reply_text("🏢 회사 위치를 설정할게요. 위치를 공유하거나 주소/건물명을 입력해주세요.")


async def commute_to_work(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_allowed(update):
        return
    places = _load_places()
    if not places.get("home") or not places.get("work"):
        await update.message.reply_text(
            "먼저 /sethome 과 /setwork 으로 집과 회사 위치를 등록해주세요."
        )
        return
    await _reply_route_comparison(update, places["home"], places["work"])


async def commute_to_home(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_allowed(update):
        return
    places = _load_places()
    if not places.get("home") or not places.get("work"):
        await update.message.reply_text(
            "먼저 /sethome 과 /setwork 으로 집과 회사 위치를 등록해주세요."
        )
        return
    await _reply_route_comparison(update, places["work"], places["home"])


async def handle_location(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_allowed(update):
        return

    loc = update.message.location
    try:
        address = await _reverse_geocode(loc.latitude, loc.longitude)
    except Exception:
        logger.exception("위치 확인 중 오류")
        address = f"위도 {loc.latitude}, 경도 {loc.longitude}"

    place = {"name": address, "address": address, "x": str(loc.longitude), "y": str(loc.latitude)}

    # 집/회사 위치 저장 중이면 그쪽으로 처리
    saving = context.user_data.get("saving_place")
    if saving:
        places = _load_places()
        places[saving] = place
        _save_places(places)
        context.user_data.pop("saving_place", None)
        label = "집" if saving == "home" else "회사"
        await update.message.reply_text(f"✅ {label} 위치를 저장했어요: {address}")
        return

    # 길찾기 진행 중이면 그쪽으로 처리
    if context.user_data.get("route_query"):
        await _advance_route_query(update, context, place)
        return

    context.user_data["location"] = {"address": address, "x": str(loc.longitude), "y": str(loc.latitude)}
    await update.message.reply_text(
        f"📍 위치를 받았어요: {address}\n"
        "이제 '근처 맛집 추천해줘'처럼 물어보시면 이 위치를 기준으로 찾아드릴게요."
    )


def _send_email(to_addr: str, subject: str, body: str) -> None:
    msg = MIMEText(body)
    msg["Subject"] = subject
    msg["From"] = MAIL_FROM_ADDRESS
    msg["To"] = to_addr

    # 로그인은 항상 GMAIL_ADDRESS 계정으로 하고, 보낸사람 표시만 MAIL_FROM_ADDRESS로 바뀜
    # (구글의 "다른 이메일 주소로 보내기" 별칭 인증이 되어 있어야 함)
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
    await update.message.reply_text(
        f"보낸사람: {MAIL_FROM_ADDRESS}\n받는사람: {to_addr}\n메일 제목을 입력해주세요."
    )


async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_allowed(update):
        await update.message.reply_text("죄송해요, 이 봇은 개인용이라 사용할 수 없어요.")
        return

    chat_id = update.effective_chat.id
    user_text = update.message.text

    # 집/회사 위치를 텍스트(주소/장소명)로 저장하려는 중이면 그쪽으로 처리
    saving = context.user_data.get("saving_place")
    if saving:
        place = await _resolve_place(user_text, {})  # '집'/'회사' 문구는 여기선 의미 없으니 빈 places
        if not place:
            await update.message.reply_text("위치를 찾지 못했어요. 다른 이름으로 다시 입력하거나 위치를 공유해주세요.")
            return
        places = _load_places()
        places[saving] = place
        _save_places(places)
        context.user_data.pop("saving_place", None)
        label = "집" if saving == "home" else "회사"
        await update.message.reply_text(f"✅ {label} 위치를 저장했어요: {place.get('name')}")
        return

    # 길찾기(출발지/도착지) 입력 중이면 그쪽으로 처리
    if context.user_data.get("route_query"):
        state = context.user_data["route_query"]["state"]
        place = await _resolve_place(user_text, _load_places())
        if not place:
            if state == "await_current_location":
                await update.message.reply_text(
                    "위치를 찾지 못했어요. 📎(첨부) 버튼으로 정확한 현재 위치를 공유해주시거나, "
                    "계신 곳 이름을 다시 알려주세요."
                )
            else:
                await update.message.reply_text("위치를 찾지 못했어요. 다른 이름으로 다시 입력하거나 위치를 공유해주세요.")
            return
        await _advance_route_query(update, context, place)
        return

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
                f"보낸사람: {MAIL_FROM_ADDRESS}\n"
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

    # 자연스러운 문장으로 길찾기를 요청한 경우 감지해서 처리
    if await _handle_natural_route_request(update, context, user_text):
        return

    history = conversations.setdefault(chat_id, [])
    history.append({"role": "user", "content": user_text})
    history = history[-MAX_HISTORY:]

    await context.bot.send_chat_action(chat_id=chat_id, action="typing")

    system_prompt = SYSTEM_PROMPT
    loc = context.user_data.get("location")
    if loc:
        loc_address = loc.get("address") if isinstance(loc, dict) else loc
        system_prompt += (
            f"\n\n참고: 사용자의 최근 공유 위치는 '{loc_address}'입니다. "
            "근처 맛집/장소 등 위치 기반 질문에는 이 정보를 활용해 웹 검색하세요."
        )

    try:
        response = client.messages.create(
            model=MODEL_NAME,
            max_tokens=1024,
            system=system_prompt,
            messages=history,
            tools=CHAT_TOOLS,
        )
        # 도구 사용(검색) 결과가 섞여 있을 수 있어 텍스트 블록만 모아서 답변으로 사용
        reply_text = "".join(
            block.text for block in response.content if block.type == "text"
        ).strip()
        if not reply_text:
            reply_text = "죄송해요, 답변을 생성하지 못했어요. 다시 한번 물어봐주세요."
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
    app.add_handler(CommandHandler("route", route_start))
    app.add_handler(CommandHandler("sethome", set_home))
    app.add_handler(CommandHandler("setwork", set_work))
    app.add_handler(CommandHandler("towork", commute_to_work))
    app.add_handler(CommandHandler("tohome", commute_to_home))
    app.add_handler(MessageHandler(filters.LOCATION, handle_location))
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
