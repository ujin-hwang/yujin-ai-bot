import os
import io
import re
import json
import asyncio
import logging
import imaplib
import smtplib
import email
from email.header import decode_header
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

import httpx
import openpyxl
from openpyxl.formula.translate import Translator
from openpyxl.cell.cell import MergedCell
from PIL import Image, ImageDraw, ImageFont
from reportlab.pdfgen import canvas as pdf_canvas
from reportlab.lib.utils import ImageReader
from pypdf import PdfReader, PdfWriter
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

# 가입증명서 자동 생성용 자산(폰트/템플릿) 경로
# assets 폴더가 있으면 그 안에서, 없으면(루트에 바로 올린 경우) bot.py와 같은 위치에서 찾음
_BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def _find_asset(filename: str) -> str:
    for candidate in (os.path.join(_BASE_DIR, "assets", filename), os.path.join(_BASE_DIR, filename)):
        if os.path.exists(candidate):
            return candidate
    return os.path.join(_BASE_DIR, filename)


CERT_FONT_PATH = _find_asset("NotoSerifKR.ttf")
CERT_TEMPLATE_PATH = _find_asset("cert_template.pdf")
CERT_PAGE_W, CERT_PAGE_H = 595.2, 841.92
CERT_FONT_SCALE = 4  # 텍스트를 이미지로 그릴 때 선명하게 보이도록 확대 비율

# 브랜드별 정산 통합파일을 서버에 계속 보관/갱신하는 폴더
MASTERS_DIR = os.path.join(_BASE_DIR, "masters")
# 브랜드별 증권번호를 저장해두는 파일 (파일명에서 못 찾을 때 사용)
POLICY_NUMBERS_FILE = os.path.join(_BASE_DIR, "policy_numbers.json")

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
        "/mail <받는사람이메일> - 메일 작성해서 보내기 (파일 첨부 가능)\n"
        "/route - 두 지점 사이 길찾기 (자동차/대중교통/도보/자전거 비교)\n"
        "/sethome - 집 위치 등록\n"
        "/setwork - 회사 위치 등록\n"
        "/towork - 집→회사 길찾기 (등록 필요)\n"
        "/tohome - 회사→집 길찾기 (등록 필요)\n"
        "/reset - 지금까지의 대화 기억 지우기\n"
        "/setpolicy <브랜드명> <증권번호> - 브랜드별 증권번호 등록 (가입증명서에 사용)\n"
        "/brands - 등록된 브랜드(통합파일) 목록 확인\n"
        "/resetbrand <브랜드명> - 해당 브랜드 통합파일 삭제하고 처음부터 다시 등록\n\n"
        "길찾기는 명령어 없이 그냥 '강남역까지 얼마나 걸려?', '홍대에서 여의도까지 어떻게 가?'처럼 물어보셔도 알아들어요.\n\n"
        "새 이메일이 오면 자동으로 요약해서 알려드려요. 📬\n"
        "날씨, 최신 뉴스, 맛집 등도 그냥 물어보시면 웹 검색해서 답해드려요.\n"
        "📎(첨부) 버튼으로 '위치'를 공유해주시면, 그 위치 기준으로 근처 맛집도 찾아드려요.\n\n"
        "정산양식 엑셀 파일을 보내주시면, 브랜드별 통합파일과 비교해서 새로 추가된 매장을 찾아 가입증명서 PDF를 자동으로 만들어드리고, 갱신된 통합파일도 함께 보내드려요. 📄"
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


async def _kakao_address_search(query: str) -> dict | None:
    """정확한 주소 문자열을 좌표로 변환 (지번/도로명 주소 전용, 상세 동/호수는 제외 권장)"""
    if not KAKAO_REST_API_KEY:
        return None
    url = "https://dapi.kakao.com/v2/local/search/address.json"
    headers = {"Authorization": f"KakaoAK {KAKAO_REST_API_KEY}"}
    params = {"query": query}
    async with httpx.AsyncClient(timeout=10) as http_client:
        resp = await http_client.get(url, headers=headers, params=params)
        resp.raise_for_status()
        data = resp.json()
    docs = data.get("documents") or []
    if not docs:
        return None
    d = docs[0]
    return {
        "name": d.get("address_name"),
        "address": d.get("address_name"),
        "x": d["x"],
        "y": d["y"],
    }


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
        # 정확한 주소(도로명/지번)는 주소 검색 API가 훨씬 정확함. 안 되면 장소명 검색으로 재시도
        place = await _kakao_address_search(t)
        if place:
            return place
        place = await _kakao_geocode_keyword(t)
        if place:
            return place
        # 문장 형태로 입력한 경우("현재 위치는 강남역이야") 장소명만 뽑아서 재시도
        extracted = await _extract_place_name_llm(t)
        if extracted and extracted != t:
            place = await _kakao_address_search(extracted)
            if place:
                return place
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
    "길찾", "가는길", "가는법", "가는방법", "어떻게가", "어떻게가나",
    "얼마나걸려", "얼마나걸리", "출근길", "퇴근길", "이동시간", "차로가", "대중교통으로",
    "버스로가", "도보로가", "걸어서가", "경로",
)


def _has_route_keyword(text: str) -> bool:
    normalized = text.replace(" ", "")
    return any(kw in normalized for kw in ROUTE_KEYWORDS)


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
    if not _has_route_keyword(user_text):
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


# ===================== 가입증명서 자동 생성 =====================
# 정산양식 엑셀(신규매장/폐점매장 시트)을 브랜드별 통합파일과 비교해 새로 추가된 매장을
# 찾아, 예시 PDF와 같은 양식의 가입증명서를 자동으로 만들어주는 기능. 원본 서식(배경/직인
# 이미지)은 그대로 두고 매장별로 달라지는 값(점포명/주소/보험기간/보험료 등)만 흰색으로
# 덮은 뒤 새로 그려넣는 방식.

# 브랜드에 등록된 증권번호가 없을 때 예시 양식과 동일하게 쓸 기본 증권번호
DEFAULT_POLICY_NO = "82509565736000"


def _render_text_image(text: str, font_size_pt: float):
    px_size = int(font_size_pt * CERT_FONT_SCALE)
    font = ImageFont.truetype(CERT_FONT_PATH, px_size)
    ascent, descent = font.getmetrics()
    bbox = font.getbbox(text)
    text_w = bbox[2] - bbox[0]
    img_h = ascent + descent
    img = Image.new("RGBA", (max(text_w, 1) + 4, img_h), (255, 255, 255, 0))
    draw = ImageDraw.Draw(img)
    draw.text((-bbox[0] + 2, 0), text, font=font, fill=(0, 0, 0, 255))
    return img, img.width / CERT_FONT_SCALE, img.height / CERT_FONT_SCALE, descent / CERT_FONT_SCALE


def _build_certificate_pdf(data: dict) -> bytes:
    """data keys: policy_no, start_date, end_date, store_name, address,
    stock_amt(int), facility_amt(int), premium(int)"""
    buf = io.BytesIO()
    c = pdf_canvas.Canvas(buf, pagesize=(CERT_PAGE_W, CERT_PAGE_H))

    def y_of(bottom):
        return CERT_PAGE_H - bottom

    def cover(x0, top, bottom, x1=545):
        c.setFillColorRGB(1, 1, 1)
        c.rect(x0 - 2, CERT_PAGE_H - bottom - 2, (x1 - x0) + 4, (bottom - top) + 4, fill=1, stroke=0)

    def draw_text(x0, bottom, s, size=12):
        img, w_pt, h_pt, descent_pt = _render_text_image(s, size)
        baseline_y = y_of(bottom)
        c.drawImage(ImageReader(img), x0, baseline_y - descent_pt, width=w_pt, height=h_pt, mask="auto")

    cover(122.7, 178.2, 190.2, x1=210)
    draw_text(122.7, 190.15, data["policy_no"])

    cover(122.7, 201.4, 213.4, x1=183)
    draw_text(122.7, 213.35, data["start_date"])
    cover(254.7, 201.4, 213.4, x1=315)
    draw_text(254.7, 213.35, data["end_date"])

    cover(122.7, 295.3, 307.3)
    draw_text(122.7, 307.25, data["store_name"])

    cover(122.7, 313.3, 325.3)
    draw_text(122.7, 325.25, data["address"])

    cover(218.7, 367.6, 379.6, x1=305)
    draw_text(218.7, 379.55, f'{data["stock_amt"]:,}원,')
    cover(368.8, 367.6, 379.6, x1=445)
    draw_text(368.8, 379.55, f'{data["facility_amt"]:,}원')

    cover(122.7, 439.8, 451.8, x1=160)
    draw_text(122.7, 451.75, f'{data["premium"]:,}')

    c.save()
    buf.seek(0)

    overlay_reader = PdfReader(buf)
    template_reader = PdfReader(CERT_TEMPLATE_PATH)
    writer = PdfWriter()
    page = template_reader.pages[0]
    page.merge_page(overlay_reader.pages[0])
    writer.add_page(page)

    out = io.BytesIO()
    writer.write(out)
    return out.getvalue()


def _norm_header(s) -> str:
    if s is None:
        return ""
    s = str(s).replace("\n", "").strip()
    s = re.sub(r"\([^)]*\)", "", s)  # 괄호 안 요율/단가 등은 파일마다 달라서 제거
    s = s.replace(" ", "")
    return s


def _find_header_row(ws, search_upto: int = 6) -> int | None:
    """'매장명' 라벨이 있는 행(하위 헤더 행)을 찾음. 파일마다 헤더가 2~3행 또는 3~4행일 수 있어 위치를 직접 스캔."""
    for r in range(1, min(ws.max_row, search_upto) + 1):
        row_vals = [_norm_header(c.value) for c in ws[r]]
        if "매장명" in row_vals:
            return r
    return None


def _build_header_map(ws) -> tuple[dict, int]:
    """헤더 행을 자동으로 찾아 '정규화된 헤더명 -> 0-based 컬럼 인덱스' 매핑과 데이터 시작 행을 반환.
    '매장명' 등 주요 라벨이 있는 행(main_row) 바로 아래 행에 병합된 그룹(재물부문 등)의
    하위 항목명(재고/시설/비품/건물 등)이 있으므로 두 행을 합쳐서 매핑을 만들고,
    실제 데이터는 그 두 행 다음부터 시작함."""
    main_row = _find_header_row(ws)
    if main_row is None:
        return {}, 4

    sub_row = main_row + 1
    row_main = [c.value for c in ws[main_row]]
    row_sub = [c.value for c in ws[sub_row]] if sub_row <= ws.max_row else []
    max_col = max(len(row_main), len(row_sub))

    header_map = {}
    for idx in range(max_col):
        vmain = row_main[idx] if idx < len(row_main) else None
        vsub = row_sub[idx] if idx < len(row_sub) else None
        # 병합된 그룹 칸(예: '재물부문(요율 0.0665%)')은 main_row에, 그 아래 세부 항목명
        # (재고/시설/비품/건물 등)은 sub_row에 있으므로 더 구체적인 sub_row를 우선함
        label = vsub if vsub not in (None, "") else vmain
        norm = _norm_header(label)
        if norm and norm not in header_map:
            header_map[norm] = idx
    return header_map, sub_row + 1


def _row_values(ws, header_map: dict, r: int) -> dict:
    row_cells = ws[r]
    vals = {}
    for key, idx in header_map.items():
        vals[key] = row_cells[idx].value if idx < len(row_cells) else None
    return vals


def _extract_data_rows(ws, header_map: dict, min_row: int) -> list:
    out = []
    for r in range(min_row, ws.max_row + 1):
        vals = _row_values(ws, header_map, r)
        if not vals.get("매장명"):
            continue
        vals["_excel_row"] = r
        out.append(vals)
    return out


def _dedup_key(vals: dict) -> tuple:
    code = vals.get("매장코드")
    if code:
        return ("code", str(code).strip())
    name = str(vals.get("매장명") or "").strip()
    date = vals.get("접수일자")
    date_key = date.isoformat() if hasattr(date, "isoformat") else str(date)
    return ("namedate", name, date_key)


def _get_brand_name(wb) -> str | None:
    """첫 시트 A1 셀이 '*. 브랜드명' 형식이면 브랜드명을 반환"""
    ws = wb.worksheets[0]
    a1 = ws["A1"].value
    if not a1:
        return None
    m = re.match(r"\*\.\s*(.+)", str(a1).strip())
    return m.group(1).strip() if m else None


def _find_type_sheets(wb, keyword: str) -> dict:
    """시트 이름에 keyword(예: '신규매장')가 포함된 시트를 '정상'/'상설' 등으로 분류"""
    result = {}
    for name in wb.sheetnames:
        if keyword not in name:
            continue
        if "정상" in name:
            result["정상"] = wb[name]
        elif "상설" in name:
            result["상설"] = wb[name]
        else:
            result[name] = wb[name]
    return result


SKIP_WRITE_FIELDS = {"월별", "순번"}


def _find_target_row(ws, header_map: dict, min_row: int) -> int:
    """매장명이 비어있는 첫 번째(수식이 이미 준비된) 행을 찾음. 없으면 max_row+1"""
    name_idx = header_map.get("매장명")
    if name_idx is None:
        return ws.max_row + 1
    for r in range(min_row, ws.max_row + 1):
        row_cells = ws[r]
        if name_idx >= len(row_cells) or not row_cells[name_idx].value:
            return r
    return ws.max_row + 1


def _find_template_row(ws, header_map: dict, min_row: int, target_row: int) -> int:
    """target_row 바로 위에서 실제 데이터(매장명)가 채워진 가장 가까운 행을 찾아 수식 기준 행으로 씀.
    (일부 파일은 빈 줄마다 수식이 끝까지 미리 채워져 있지 않고 중간에 끊겨 있어서,
    바로 위 빈 줄이 아니라 '가장 최근 실제로 채워진 행'을 기준으로 삼아야 수식이 안전하게 이어짐)"""
    name_idx = header_map.get("매장명")
    r = target_row - 1
    while r >= min_row:
        if name_idx is None or ws.cell(row=r, column=name_idx + 1).value:
            return r
        r -= 1
    return target_row


def _write_row(ws, header_map: dict, target_row: int, values: dict, template_row: int) -> None:
    """각 컬럼마다 '기준 행(template_row)'의 셀이 수식이면 그 수식을 target_row로 복사(번역)하고,
    수식이 아니면(원본 입력값 칸이면) values의 값을 그대로 채워 넣음. 필드명을 하드코딩해서
    구분하지 않고, 실제로 그 칸이 수식인지 아닌지를 셀 단위로 직접 확인하기 때문에 파일마다
    수식이 있는 칸이 달라도(예: 폐점매장의 매장명이 수식인 경우 등) 안전하게 동작함."""
    for key, idx in header_map.items():
        if key in SKIP_WRITE_FIELDS:
            continue
        col = idx + 1
        target_cell = ws.cell(row=target_row, column=col)
        if isinstance(target_cell, MergedCell):
            continue
        template_cell = ws.cell(row=template_row, column=col)
        template_val = template_cell.value
        is_formula = isinstance(template_val, str) and template_val.startswith("=")

        if is_formula:
            if target_row != template_row:
                target_cell.value = Translator(
                    template_val, origin=template_cell.coordinate
                ).translate_formula(target_cell.coordinate)
            continue

        if key in values and values[key] not in (None, ""):
            target_cell.value = values[key]

    # 새로 만든 행(기존 서식 범위를 벗어난 경우)엔 순번을 이어서 채워줌
    if target_row != template_row:
        seq_idx = header_map.get("순번")
        if seq_idx is not None:
            prev = ws.cell(row=template_row, column=seq_idx + 1).value
            if isinstance(prev, (int, float)):
                ws.cell(row=target_row, column=seq_idx + 1).value = int(prev) + 1


def _extract_rate(ws, header_map: dict, field: str, pattern: str, default: float, min_row: int) -> float:
    idx = header_map.get(field)
    if idx is None:
        return default
    for r in range(min_row, min(ws.max_row, min_row + 50) + 1):
        cell = ws.cell(row=r, column=idx + 1)
        if isinstance(cell.value, str) and cell.value.startswith("="):
            m = re.search(pattern, cell.value)
            if m:
                return float(m.group(1))
    return default


def _compute_new_store_cert_values(vals: dict, rate1_pct: float, rate2: float) -> dict:
    stock = float(vals.get("재고") or 0)
    facility = float(vals.get("시설/비품") or 0)
    building = float(vals.get("건물") or 0)
    total = stock + facility + building
    property_premium = total * rate1_pct / 100
    pyeong = float(vals.get("평수") or 0)
    liability_premium = pyeong * rate2
    start = vals.get("보험시작일")
    end = vals.get("보험종기일")
    if hasattr(start, "date"):
        start_d = start
    else:
        start_d = None
    if hasattr(end, "date"):
        end_d = end
    else:
        end_d = None
    if start_d and end_d:
        days = (end_d - start_d).days
        premium = (property_premium + liability_premium) * days / 365
    else:
        premium = 0
    return {
        "stock_amt": int(stock),
        "facility_amt": int(facility),
        "premium": round(premium),
        "start_date": start_d.strftime("%Y.%m.%d") if start_d else str(start or ""),
        "start_date_yymmdd": start_d.strftime("%y%m%d") if start_d else "",
        "end_date": end_d.strftime("%Y.%m.%d") if end_d else str(end or ""),
    }


def _load_policy_numbers() -> dict:
    try:
        with open(POLICY_NUMBERS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return {}


def _save_policy_numbers(policy_numbers: dict) -> None:
    os.makedirs(os.path.dirname(POLICY_NUMBERS_FILE), exist_ok=True)
    with open(POLICY_NUMBERS_FILE, "w", encoding="utf-8") as f:
        json.dump(policy_numbers, f, ensure_ascii=False, indent=2)


async def set_policy_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_allowed(update):
        return
    if len(context.args) < 2:
        await update.message.reply_text(
            "사용법: /setpolicy <브랜드명> <증권번호>\n"
            "예: /setpolicy 트레몰로 82509565736000\n\n"
            "브랜드명은 엑셀 시트 A1 셀에 적힌 이름(예: '트레몰로', '월메이드')과 정확히 같아야 해요."
        )
        return
    brand = context.args[0]
    policy_no = context.args[1]
    policy_numbers = _load_policy_numbers()
    policy_numbers[brand] = policy_no
    _save_policy_numbers(policy_numbers)
    await update.message.reply_text(f"✅ '{brand}' 브랜드의 증권번호를 등록했어요: {policy_no}")


async def list_brands_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_allowed(update):
        return
    if not os.path.isdir(MASTERS_DIR):
        await update.message.reply_text("등록된 브랜드가 아직 없어요.")
        return
    brands = sorted(f[:-5] for f in os.listdir(MASTERS_DIR) if f.endswith(".xlsx"))
    if not brands:
        await update.message.reply_text("등록된 브랜드가 아직 없어요.")
        return
    policy_numbers = _load_policy_numbers()
    lines = [f"- {b} ({'증권번호 등록됨' if b in policy_numbers else '⚠️ 증권번호 미등록'})" for b in brands]
    await update.message.reply_text("📋 등록된 브랜드 목록:\n" + "\n".join(lines))


async def reset_brand_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_allowed(update):
        return
    if not context.args:
        await update.message.reply_text(
            "사용법: /resetbrand <브랜드명>\n예: /resetbrand 월메이드\n\n"
            "해당 브랜드의 통합파일을 삭제해요. 다음에 그 브랜드 엑셀을 올리면 그 파일을 새 기준으로 다시 등록해요.\n"
            "정확한 브랜드명은 /brands 로 확인할 수 있어요."
        )
        return
    brand = " ".join(context.args)
    master_path = os.path.join(MASTERS_DIR, f"{brand}.xlsx")
    if os.path.exists(master_path):
        os.remove(master_path)
        await update.message.reply_text(f"🗑️ '{brand}' 통합파일을 삭제했어요. 다음에 이 브랜드 엑셀을 올리면 그 파일을 새 기준으로 등록할게요.")
    else:
        await update.message.reply_text(f"'{brand}' 통합파일을 찾지 못했어요. /brands 로 정확한 브랜드명을 확인해주세요.")


def _sync_brand_excel(file_bytes: bytes) -> dict | None:
    """엑셀을 읽어 브랜드를 판별하고, 저장된 통합파일과 비교해 신규/폐점 매장을 반영.
    반환: {"brand", "cold_start", "new_stores": [...], "closed_count": int, "master_bytes": bytes}
    브랜드를 못 찾으면 None."""
    try:
        input_wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    except Exception:
        logger.exception("엑셀 파일을 여는 중 오류")
        return None

    brand = _get_brand_name(input_wb)
    if not brand:
        return None

    os.makedirs(MASTERS_DIR, exist_ok=True)
    master_path = os.path.join(MASTERS_DIR, f"{brand}.xlsx")

    if not os.path.exists(master_path):
        with open(master_path, "wb") as f:
            f.write(file_bytes)
        return {"brand": brand, "cold_start": True, "new_stores": [], "closed_count": 0, "master_bytes": file_bytes}

    # 쓰기용(수식 보존)과 읽기용(중복 확인은 계산된 값 기준) 두 벌로 마스터 파일을 엶.
    # 폐점매장 시트처럼 '매장명' 칸 자체가 VLOOKUP 수식인 행이 실제로 존재하기 때문에,
    # 중복 판별은 반드시 계산된 값(data_only=True) 기준으로 해야 정확함.
    master_wb = openpyxl.load_workbook(master_path, data_only=False)
    master_wb_values = openpyxl.load_workbook(master_path, data_only=True)

    policy_numbers = _load_policy_numbers()
    policy_no = policy_numbers.get(brand, "")

    new_stores = []
    closed_count = 0

    input_new_sheets = _find_type_sheets(input_wb, "신규매장")
    master_new_sheets = _find_type_sheets(master_wb, "신규매장")
    master_new_sheets_values = _find_type_sheets(master_wb_values, "신규매장")

    for sub_type, input_ws in input_new_sheets.items():
        master_ws = master_new_sheets.get(sub_type)
        master_ws_values = master_new_sheets_values.get(sub_type)
        if master_ws is None or master_ws_values is None:
            continue
        input_header, input_min_row = _build_header_map(input_ws)
        master_header, master_min_row = _build_header_map(master_ws)
        if not input_header or not master_header:
            continue
        existing_keys = {_dedup_key(v) for v in _extract_data_rows(master_ws_values, master_header, master_min_row)}
        rate1 = _extract_rate(master_ws, master_header, "연간재물보험료", r"\*([\d.]+)%", 0.0665, master_min_row)
        rate2 = _extract_rate(master_ws, master_header, "연간영업배상보험료", r"\*([\d.]+)", 1793, master_min_row)

        for vals in _extract_data_rows(input_ws, input_header, input_min_row):
            key = _dedup_key(vals)
            if key in existing_keys:
                continue
            existing_keys.add(key)

            target_row = _find_target_row(master_ws, master_header, master_min_row)
            template_row = _find_template_row(master_ws, master_header, master_min_row, target_row)
            _write_row(master_ws, master_header, target_row, vals, template_row)

            cert_vals = _compute_new_store_cert_values(vals, rate1, rate2)
            address = str(vals.get("매장주소") or "").strip()
            address = re.sub(r"(?<=\S)\(", " (", address)
            new_stores.append({
                "policy_no": policy_no or DEFAULT_POLICY_NO,
                "store_code": str(vals.get("매장코드") or "").strip(),
                "store_name": str(vals.get("매장명") or "").strip(),
                "address": address,
                **cert_vals,
            })

    input_closed_sheets = _find_type_sheets(input_wb, "폐점매장")
    master_closed_sheets = _find_type_sheets(master_wb, "폐점매장")
    master_closed_sheets_values = _find_type_sheets(master_wb_values, "폐점매장")

    for sub_type, input_ws in input_closed_sheets.items():
        master_ws = master_closed_sheets.get(sub_type)
        master_ws_values = master_closed_sheets_values.get(sub_type)
        if master_ws is None or master_ws_values is None:
            continue
        input_header, input_min_row = _build_header_map(input_ws)
        master_header, master_min_row = _build_header_map(master_ws)
        if not input_header or not master_header:
            continue
        existing_keys = {_dedup_key(v) for v in _extract_data_rows(master_ws_values, master_header, master_min_row)}

        for vals in _extract_data_rows(input_ws, input_header, input_min_row):
            key = _dedup_key(vals)
            if key in existing_keys:
                continue
            existing_keys.add(key)

            target_row = _find_target_row(master_ws, master_header, master_min_row)
            template_row = _find_template_row(master_ws, master_header, master_min_row, target_row)
            _write_row(master_ws, master_header, target_row, vals, template_row)
            closed_count += 1

    if not new_stores and not closed_count:
        return {"brand": brand, "cold_start": False, "new_stores": [], "closed_count": 0, "master_bytes": None}

    out = io.BytesIO()
    master_wb.save(out)
    master_bytes = out.getvalue()
    with open(master_path, "wb") as f:
        f.write(master_bytes)

    return {
        "brand": brand,
        "cold_start": False,
        "new_stores": new_stores,
        "closed_count": closed_count,
        "master_bytes": master_bytes,
        "has_policy_no": bool(policy_no),
    }


async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_allowed(update):
        return

    draft = context.user_data.get("mail_draft")
    if not (draft and draft.get("state") == "attach"):
        await update.message.reply_text(
            "사진은 메일 작성 중 '첨부할 파일' 단계에서만 첨부할 수 있어요. /mail로 메일 작성을 시작해보세요."
        )
        return

    photo = update.message.photo[-1]  # 가장 고화질
    try:
        tg_file = await photo.get_file()
        file_bytes = bytes(await tg_file.download_as_bytearray())
    except Exception:
        logger.exception("사진 다운로드 중 오류")
        await update.message.reply_text("사진을 받아오는 중 오류가 발생했어요. 다시 보내주세요.")
        return

    n = len(draft.get("attachments", [])) + 1
    filename = f"사진_{n}.jpg"
    draft.setdefault("attachments", []).append((filename, file_bytes))
    await update.message.reply_text(
        f"📎 첨부됨: {filename} (총 {n}개)\n계속 보내시거나, 다 되셨으면 '없음'/'완료'라고 입력해주세요."
    )


async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_allowed(update):
        return

    doc = update.message.document
    filename = doc.file_name or "attachment"

    # 메일 작성 중 첨부파일 단계면, 파일 종류 상관없이 메일 첨부로 처리
    draft = context.user_data.get("mail_draft")
    if draft and draft.get("state") == "attach":
        try:
            tg_file = await doc.get_file()
            file_bytes = bytes(await tg_file.download_as_bytearray())
        except Exception:
            logger.exception("첨부파일 다운로드 중 오류")
            await update.message.reply_text("파일을 받아오는 중 오류가 발생했어요. 다시 보내주세요.")
            return
        draft.setdefault("attachments", []).append((filename, file_bytes))
        n = len(draft["attachments"])
        await update.message.reply_text(
            f"📎 첨부됨: {filename} (총 {n}개)\n계속 보내시거나, 다 되셨으면 '없음'/'완료'라고 입력해주세요."
        )
        return

    if not filename.lower().endswith(".xlsx"):
        await update.message.reply_text(
            "죄송해요, 지금은 정산양식 엑셀(.xlsx) 파일만 처리할 수 있어요."
        )
        return

    if not os.path.exists(CERT_TEMPLATE_PATH) or not os.path.exists(CERT_FONT_PATH):
        await update.message.reply_text("가입증명서 생성 기능이 아직 설정되지 않았어요.")
        return

    await update.message.reply_text("엑셀을 확인하고 있어요...")

    try:
        tg_file = await doc.get_file()
        file_bytes = bytes(await tg_file.download_as_bytearray())
    except Exception:
        logger.exception("엑셀 파일 다운로드 중 오류")
        await update.message.reply_text("파일을 받아오는 중 오류가 발생했어요.")
        return

    try:
        result = _sync_brand_excel(file_bytes)
    except Exception:
        logger.exception("엑셀 동기화 중 오류")
        await update.message.reply_text("⚠️ 엑셀을 처리하는 중 오류가 발생했어요.")
        return

    if result is None:
        await update.message.reply_text(
            "이 엑셀 형식은 알아보지 못했어요. 각 시트 A1 셀에 '*. 브랜드명'이 적힌 정산양식인지 확인해주세요."
        )
        return

    brand = result["brand"]

    if result["cold_start"]:
        await update.message.reply_text(
            f"📁 '{brand}' 통합파일을 처음 등록했어요. 앞으로 이 파일을 기준으로 신규/폐점 매장을 비교할게요."
        )
        return

    if not result["new_stores"] and not result["closed_count"]:
        await update.message.reply_text(f"'{brand}' 기준으로 새로운 신규/폐점 매장이 없어요.")
        return

    if not result.get("has_policy_no"):
        await update.message.reply_text(
            f"⚠️ '{brand}'의 증권번호가 등록되어 있지 않아 기본 증권번호로 가입증명서를 만들어요."
        )

    for store in result["new_stores"]:
        try:
            pdf_bytes = _build_certificate_pdf(store)
        except Exception:
            logger.exception("가입증명서 생성 중 오류")
            await update.message.reply_text(f"⚠️ '{store['store_name']}' 가입증명서 생성에 실패했어요.")
            continue

        out_name = f"{store['store_name']}_{store['store_code']}_{store['start_date_yymmdd']}.pdf"
        await update.message.reply_document(
            document=io.BytesIO(pdf_bytes),
            filename=out_name,
            caption=f"📄 {store['store_name']} 가입증명서",
        )

    summary = f"✅ '{brand}' 통합파일을 갱신했어요.\n신규 매장 {len(result['new_stores'])}곳"
    if result["closed_count"]:
        summary += f", 폐점 매장 {result['closed_count']}곳"
    await update.message.reply_text(summary)

    if result.get("master_bytes"):
        await update.message.reply_document(
            document=io.BytesIO(result["master_bytes"]),
            filename=f"{brand}.xlsx",
            caption=f"📎 갱신된 '{brand}' 통합파일이에요.",
        )


def _send_email(to_addr: str, subject: str, body: str, attachments: list | None = None) -> None:
    """attachments: [(파일명, 파일바이트), ...]"""
    if attachments:
        msg = MIMEMultipart()
        msg.attach(MIMEText(body))
        for filename, file_bytes in attachments:
            part = MIMEApplication(file_bytes, Name=filename)
            part["Content-Disposition"] = f'attachment; filename="{filename}"'
            msg.attach(part)
    else:
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
            draft["state"] = "attach"
            draft["attachments"] = []
            await update.message.reply_text(
                "첨부할 파일이 있으면 지금 보내주세요 (여러 개 가능해요, 엑셀/PDF/워드/한글 파일, 사진 등).\n"
                "다 보내셨으면 '없음' 또는 '완료'라고 입력해주세요."
            )
            return

        if state == "attach":
            if user_text.strip() in ("없음", "완료", "no", "done", "skip"):
                draft["state"] = "confirm"
                n = len(draft.get("attachments", []))
                attach_line = f"첨부파일: {n}개\n" if n else ""
                preview = (
                    f"보낸사람: {MAIL_FROM_ADDRESS}\n"
                    f"받는사람: {draft['to']}\n"
                    f"제목: {draft['subject']}\n"
                    f"{attach_line}\n"
                    f"{draft['body']}\n\n"
                    "이대로 보낼까요? '네' 또는 '아니오'로 답해주세요."
                )
                await update.message.reply_text(preview)
            else:
                await update.message.reply_text(
                    "파일을 보내시거나, 다 되셨으면 '없음'/'완료'라고 입력해주세요."
                )
            return

        if state == "confirm":
            context.user_data.pop("mail_draft", None)
            if user_text.strip() in ("네", "예", "ㅇㅇ", "y", "yes", "Y"):
                try:
                    _send_email(
                        draft["to"], draft["subject"], draft["body"],
                        attachments=draft.get("attachments") or None,
                    )
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
    app.add_handler(CommandHandler("setpolicy", set_policy_command))
    app.add_handler(CommandHandler("brands", list_brands_command))
    app.add_handler(CommandHandler("resetbrand", reset_brand_command))
    app.add_handler(MessageHandler(filters.LOCATION, handle_location))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(MessageHandler(filters.PHOTO, handle_photo))
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
