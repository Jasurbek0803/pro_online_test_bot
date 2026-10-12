from __future__ import annotations

from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.colors import HexColor, black
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
import asyncio
import contextlib
import html
import json
import logging
import math
import sys
import time
import re
import os
import tempfile
import random
import string
from datetime import datetime, timezone, timedelta, time as dtime
from collections import defaultdict, deque

import asyncpg
import base64
import hashlib
import hmac
from urllib.parse import parse_qsl
import aiohttp
from aiohttp import web
from dotenv import load_dotenv
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

from aiogram import Bot, Dispatcher, Router, F
from aiogram.client.default import DefaultBotProperties
from aiogram.enums import ParseMode
from aiogram.filters import CommandStart, StateFilter, Command
from aiogram.types import (
    Message, CallbackQuery,
    ReplyKeyboardMarkup, KeyboardButton,
    InlineKeyboardMarkup, InlineKeyboardButton,
    FSInputFile, BufferedInputFile, BotCommand, BotCommandScopeDefault, BotCommandScopeChat, ErrorEvent,
    WebAppInfo
)
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.base import BaseStorage, StorageKey, StateType
from aiogram.fsm.storage.memory import SimpleEventIsolation
from aiogram.dispatcher.middlewares.base import BaseMiddleware
from aiogram.exceptions import TelegramRetryAfter, TelegramForbiddenError, TelegramBadRequest
from aiogram.utils.callback_answer import CallbackAnswerMiddleware
from aiogram.client.session.middlewares.base import BaseRequestMiddleware
from aiogram.methods import AnswerCallbackQuery

# ===================== CONFIG =====================
load_dotenv()


def _env(name: str, default: str | None = None, required: bool = False) -> str:
    val = os.environ.get(name, default)
    if required and not val:
        raise RuntimeError(f"Environment variable {name} is required (see .env.example)")
    return val


BOT_TOKEN = _env("BOT_TOKEN", required=True)
ADMIN_ID = int(_env("ADMIN_ID", required=True))
DATABASE_URL = _env("DATABASE_URL", required=True)

ACADEMY_NAME = _env("ACADEMY_NAME", "Smart Edu Center")
CERT_LOGO_PATH = _env("CERT_LOGO_PATH", "assets/logo.png")
CERT_SIGN_PATH = _env("CERT_SIGN_PATH", "assets/sign.png")
CERT_MIN_PERCENT = 50.0  # Barcha testlar uchun threshold

PAYMENT_CARD = _env("PAYMENT_CARD", required=True)
PAYMENT_OWNER = _env("PAYMENT_OWNER", required=True)
PAYMENT_PHONE = _env("PAYMENT_PHONE", required=True)
PAYMENT_NOTE = "To'lov izohiga test kodini yozing (masalan: MS-001)."

# ✅ Ma'lumotlar bazasidan avtomatik zaxira nusxa (backup) — necha soatda bir
# marta ADMIN_ID'ga JSON fayl qilib yuboriladi. Bu hosting ishonchliligidan
# mustaqil qo'shimcha xavfsizlik: hosting/server muammo qilsa ham, adminning
# Telegram xabarlar tarixida oxirgi holatning zaxirasi qoladi. 0 = o'chirilgan.
DB_BACKUP_INTERVAL_HOURS = int(_env("DB_BACKUP_INTERVAL_HOURS", "24") or "24")

# ✅ Mini App (veb-profil): botni Telegram ichida ochiladigan veb-sahifa bilan
# to'ldiradi (statistika/natijalar/sertifikatlar). WEBAPP_URL bo'sh bo'lsa,
# tugma umuman ko'rsatilmaydi — Telegram web_app tugmasi faqat HTTPS manzil
# bilan ishlaydi, shuning uchun buni faqat public HTTPS orqasida (masalan
# domenli VPS yoki vaqtincha ngrok) sozlab, keyin WEBAPP_URL'ga yozing.
WEBAPP_URL = _env("WEBAPP_URL", "").rstrip("/")
WEBAPP_HOST = _env("WEBAPP_HOST", "0.0.0.0")
WEBAPP_PORT = int(_env("WEBAPP_PORT", "8080") or "8080")

# ✅ IELTS Writing/Speaking'ni test yaratuvchi xohlasa AI (Gemini) avtomatik
# baholaydi, xohlamasa qo'lda (ustoz) tekshiradi — buni har bir bo'lim uchun
# alohida tanlash mumkin. Kalit bo'sh bo'lsa, AI tanlovi umuman ko'rsatilmaydi
# va hammasi avvalgidek qo'lda tekshiriladi.
GEMINI_API_KEY = _env("GEMINI_API_KEY", "")
GEMINI_MODEL = _env("GEMINI_MODEL", "gemini-3.6-flash")

TOP_N = 10

# Asia/Tashkent fixed offset
TZ = timezone(timedelta(hours=5))  # +05:00

# Anti-spam / security
MSG_WINDOW_SEC = 15
MSG_LIMIT = 10
CB_WINDOW_SEC = 15
CB_LIMIT = 15
SAME_CALLBACK_COOLDOWN_SEC = 2
MUTE_STEPS = [120, 300, 900, 3600, 86400]  # 2m, 5m, 15m, 1h, 1d
STRIKE_RESET_AFTER_SEC = 24 * 3600

# Countdown (reverse timer)
COUNTDOWN_TICK_SEC = 60  # 60s

# Global connection pool (created in main() before polling starts)
POOL: asyncpg.Pool | None = None

# ===================== UI TEXT =====================
BTN_FREE = "🆓 Tekin testlar"
BTN_PAID = "🧾 Pullik testlar"
BTN_MY = "📦 Pullik testlarim"
BTN_RATING = "🏆 Reyting"
BTN_RESULTS = "📄 Natijalarim"
BTN_HELP = "ℹ️ Yordam"

BTN_ADD_TEST = "➕ Test Yaratish"
BTN_MY_CREATED_TESTS = "🛠 Yaratgan testlarim"

BTN_ADMIN_PENDING = "🟡 Pending to'lovlar (Admin)"
BTN_ADMIN_STATS = "📊 Statistika (Admin)"
BTN_ADMIN_CHANNELS = "📣 Majburiy kanallar (Admin)"
BTN_ADMIN_DEL_TEST = "🗑 Testni o'chirish (Admin)"
BTN_ADMIN_ADD_ADMIN = "👮 Adminlarni boshqarish"
BTN_PROFILE = "👤 Profilim"

BTN_HOME = "🏠 Bosh menyu"
# ✅ Bu tugma haqiqatda BIR QADAM ORQAGA emas — butun jarayonni bekor qilib
# bosh menyuga qaytaradi (xuddi "Bekor qilish"/"Bosh menyu" kabi). Avval
# "⬅️ Ortga" deb nomlangan edi — bu "oldingi qadamga qaytadi" degan noto'g'ri
# taassurot qoldirardi, uzun (masalan bosqichli IELTS/CEFR) sehrgarlarda
# foydalanuvchi bitta xatoni tuzatmoqchi bo'lib bosganda BUTUN progressini
# yo'qotib qo'yardi. Nomi shu voqelikka mos qilib o'zgartirildi.
BTN_BACK = "🔄 Boshidan boshlash"
BTN_CANCEL = "Bekor qilish"

TOP_MENU_BTNS = {
    BTN_FREE, BTN_PAID, BTN_MY, BTN_RATING, BTN_RESULTS, BTN_HELP,
    BTN_ADD_TEST, BTN_MY_CREATED_TESTS,
    BTN_ADMIN_PENDING, BTN_ADMIN_STATS, BTN_ADMIN_CHANNELS,
    BTN_ADMIN_DEL_TEST, BTN_ADMIN_ADD_ADMIN,
    BTN_HOME, BTN_BACK, BTN_CANCEL
}

# Section-switch buttons only (excludes nav buttons, which have their own dedicated handlers)
SECTION_BTNS = TOP_MENU_BTNS - {BTN_HOME, BTN_BACK, BTN_CANCEL}


# ===================== STATES =====================
class RegistrationState(StatesGroup):
    full_name = State()
    phone = State()


class EditProfileState(StatesGroup):
    waiting_name = State()
    waiting_phone = State()
    waiting_age = State()
    waiting_school = State()
    waiting_grade = State()
    waiting_payment_card = State()
    waiting_payment_owner = State()
    waiting_payment_phone = State()


UZ_REGIONS = [
    "Toshkent shahri", "Toshkent viloyati", "Andijon", "Farg'ona", "Namangan",
    "Samarqand", "Buxoro", "Navoiy", "Qashqadaryo", "Surxondaryo",
    "Jizzax", "Sirdaryo", "Xorazm", "Qoraqalpog'iston Respublikasi",
]


class PaymentState(StatesGroup):
    waiting_for_screenshot = State()


class AdminAddTestState(StatesGroup):
    subject = State()
    title = State()
    is_free = State()
    price = State()
    duration = State()
    questions_count = State()
    exam_type = State()
    schedule_mode = State()
    schedule_date = State()
    schedule_time = State()
    file = State()
    answers = State()
    open_ref_answers = State()


class TestProcessState(StatesGroup):
    solving = State()
    awaiting_free_text = State()  # variantsiz savolga matn javob kutilmoqda


class StagedBuildState(StatesGroup):
    """Bosqichli imtihon (DTM 5-fan / IELTS 4-ko'nikma) yaratish sehrgari —
    navbatdagi har bir bosqich uchun ketma-ket so'raladi. Qaysi maydon
    (nomi/davomiyligi/fayli/soni/kaliti/matni) so'ralayotgani FSM
    data'dagi "stage_field" orqali kuzatiladi."""
    collecting = State()


class StagedSolvingState(StatesGroup):
    """O'quvchi bosqichli imtihonni yechayotgan holat — qaysi bosqichda
    ekanligi va uning turi (mcq/open_text/open_voice) FSM data'da saqlanadi."""
    in_section = State()
    awaiting_free_text = State()  # MCQ bosqichida variantsiz savolga matn kutilmoqda


class StagedEditState(StatesGroup):
    """Mavjud bosqichli testning bitta bosqichini (masalan Reading yoki
    fan2) qayta sozlash holati."""
    collecting = State()


class StagedGradingState(StatesGroup):
    """Ustoz Writing/Speaking uchun band + fikr-mulohaza kiritayotgan holat."""
    writing_band = State()
    writing_feedback = State()
    speaking_band = State()
    speaking_feedback = State()


class AdminRejectState(StatesGroup):
    waiting_reason = State()


class SearchByCodeState(StatesGroup):
    waiting_code = State()
    scope = State()  # paid | free | my


class AdminChannelState(StatesGroup):
    waiting_channel_ref = State()
    waiting_join_url = State()
    edit_id = State()
    waiting_edit_payload = State()


class AdminManageState(StatesGroup):
    waiting_test_code_to_del = State()
    waiting_admin_id = State()


class AdminBroadcastState(StatesGroup):
    waiting_text = State()


class AdminUserSearchState(StatesGroup):
    waiting_query = State()


class AdminEditTestState(StatesGroup):
    waiting_code = State()
    waiting_field = State()
    waiting_value = State()


class AskAdminState(StatesGroup):
    waiting_text = State()


class CertConfigState(StatesGroup):
    waiting_academy = State()
    waiting_signee = State()
    waiting_template = State()


class RestoreState(StatesGroup):
    waiting_file = State()


class ReopenScheduleState(StatesGroup):
    waiting_datetime = State()


# ===================== HELPERS =====================
def now_str_local():
    return datetime.now(TZ).strftime("%Y-%m-%d %H:%M")


def now_ts() -> int:
    return int(time.time())


def fmt_seconds(seconds: int) -> str:
    seconds = max(0, int(seconds))
    if seconds < 60:
        return f"{seconds} soniya"
    m, s = divmod(seconds, 60)
    if m < 60:
        return f"{m} daqiqa {s} soniya"
    h, m = divmod(m, 60)
    if h < 24:
        return f"{h} soat {m} daqiqa"
    d, h = divmod(h, 24)
    return f"{d} kun {h} soat"


def normalize_code(code: str) -> str:
    code = (code or "").strip().upper()
    code = re.sub(r"\s+", "", code)
    return code


def is_valid_code(code: str) -> bool:
    return bool(re.fullmatch(r"[A-Z0-9\-_]{3,20}", code))


def normalize_answers(ans: str) -> str:
    ans = (ans or "").strip().lower()
    ans = re.sub(r"\s+", "", ans)
    return ans


def is_answer_string_valid(ans: str) -> bool:
    # ✅ Interaktiv test tugmalari faqat A..F (6 ta variant) beradi — javob
    # kalitida shundan tashqari harf (masalan xato bosilgan "g") bo'lsa, o'sha
    # savolni tugma orqali hech kim to'g'ri bosa olmaydi (doim xato hisoblanadi).
    # Shu sabab kalit ham, foydalanuvchi javobi ham a..f bilan cheklanadi.
    # "x" — variantsiz (ochiq) savol belgisi: javob kalitida bo'lsa, o'sha savol
    # avtomatik baholashdan chetlab o'tiladi (pastda score_skip_open'ga qarang);
    # foydalanuvchi javobida bo'lsa, u shu savolga tugma o'rniga matn kiritgan.
    return bool(re.fullmatch(r"[a-f\*x]+", ans))  # * belgisi o'tkazib yuborilganlarga ruxsat


def score_skip_open(user_ans: str, key: str, qcount: int) -> tuple[int, int]:
    """To'g'ri javoblar sonini va HISOBGA OLINADIGAN savollar sonini qaytaradi.
    Javob kalitida "x" turgan pozitsiyalar (variantsiz/ochiq savol — ustoz
    alohida tekshiradi) na to'g'ri, na xato hisoblanmaydi, umuman
    hisoblashdan chetlab o'tiladi."""
    correct = 0
    scored_total = 0
    for i in range(qcount):
        if key[i] == "x":
            continue
        scored_total += 1
        if user_ans[i] == key[i]:
            correct += 1
    return correct, scored_total


def _h(s) -> str:
    """Foydalanuvchi/tashqi manbadan kelgan matnni HTML xabar ichida xavfsiz
    ko'rsatish uchun ekranlaydi (Telegram HTML rejimida "<" yoki "&" xabarni
    butunlay buzib, yuborilmay qolishiga sabab bo'lardi)."""
    return html.escape(str(s if s is not None else ""), quote=False)


def _load_json_dict(raw) -> dict:
    if not raw:
        return {}
    if isinstance(raw, dict):
        return raw
    try:
        val = json.loads(raw)
    except (json.JSONDecodeError, TypeError):
        return {}
    return val if isinstance(val, dict) else {}


def open_answer_has_content(v) -> bool:
    text_val, photo_id = _open_answer_preview(v)
    return bool((text_val or "").strip() or photo_id)


def _norm_open_answer(s: str) -> str:
    s = (s or "").strip().lower()
    for ch in ("ʻ", "ʼ", "’", "‘", "`", "´"):
        s = s.replace(ch, "'")
    s = re.sub(r"\s+", "", s)
    s = s.rstrip(".!;")
    if re.fullmatch(r"[-+]?\d+,\d+", s):
        s = s.replace(",", ".")
    return s


def auto_grade_open_answer(answer_text: str, ref: str) -> int:
    """O'quvchining ochiq savolga yozgan javobini ustoz test yaratishda
    kiritgan kalit (kutilgan javob) bilan avtomatik solishtiradi: katta-kichik
    harf, bo'sh joylar, apostrof turlari va 0,5/0.5 kabi farqlar hisobga
    olinmaydi. Kalitda bir nechta to'g'ri variant "|" bilan berilishi mumkin
    (masalan "Parij|Paris"). 1 = to'g'ri, 0 = xato."""
    a = _norm_open_answer(answer_text)
    if not a:
        return 0
    for variant in str(ref or "").split("|"):
        r = _norm_open_answer(variant)
        if not r:
            continue
        if a == r:
            return 1
        try:
            if abs(float(a) - float(r)) < 1e-9:
                return 1
        except ValueError:
            pass
    return 0


def initial_open_grades(key: str, open_answers: dict | None, open_refs: dict | None = None,
                        offset: int = 0) -> dict:
    """Ochiq (x) savollar uchun boshlang'ich baholar lug'ati:
      - o'quvchi umuman javob yozmagan — darhol 0 (xato);
      - matnli javob va ustoz kiritgan kalit bor — AVTOMATIK solishtiriladi
        (ustoz kalitni test yaratishda allaqachon bergan, qayta tekshirishi
        shart emas);
      - faqat rasm yuborilgan yoki kalit yo'q — lug'atga kirmaydi, ya'ni
        ustoz qo'lda tekshirishi kutiladi."""
    open_answers = open_answers or {}
    open_refs = open_refs or {}
    grades = {}
    for i, c in enumerate(key):
        if c != "x":
            continue
        q = str(offset + i + 1)
        raw = open_answers.get(q)
        if not open_answer_has_content(raw):
            grades[q] = 0
            continue
        text_val, _photo = _open_answer_preview(raw)
        ref = open_refs.get(str(i + 1)) if offset else open_refs.get(q)
        if (text_val or "").strip() and ref:
            grades[q] = auto_grade_open_answer(text_val, ref)
    return grades


def score_with_open(user_ans: str, key: str, qcount: int, open_grades: dict | None = None,
                    offset: int = 0) -> tuple[int, int, int]:
    """(to'g'ri javoblar, hisobga olingan savollar, tekshirilmagan ochiq savollar).
    Ochiq (x) savol ustoz baholagan bo'lsa (1/0) oddiy savol kabi hisobga
    kiradi; hali baholanmagan bo'lsa — vaqtincha chetda turadi."""
    open_grades = open_grades or {}
    correct = scored = pending = 0
    for i in range(qcount):
        if i >= len(key):
            break
        if key[i] == "x":
            g = open_grades.get(str(offset + i + 1))
            if g is None:
                pending += 1
                continue
            scored += 1
            if int(g) == 1:
                correct += 1
            continue
        scored += 1
        if i < len(user_ans) and user_ans[i] == key[i]:
            correct += 1
    return correct, scored, pending


def open_question_positions(answer_key: str) -> list[int]:
    """Javob kalitida "x" (variantsiz/ochiq savol) turgan pozitsiyalarni
    1-asosli savol raqamlari sifatida qaytaradi."""
    return [i + 1 for i, c in enumerate(answer_key) if c == "x"]


def _open_answer_preview(v) -> tuple[str, str | None]:
    """O'quvchining ochiq savolga bergan javobini (matn, rasm_id) qilib
    normallashtiradi — eski format (faqat matn qatori) bilan ham, yangi
    format ({"text":..., "photo_id":...} lug'at) bilan ham ishlaydi."""
    if isinstance(v, dict):
        return (v.get("text") or ""), v.get("photo_id")
    return (v or ""), None


def _open_ref_prompt_text(positions: list[int]) -> str:
    nums = ", ".join(str(p) for p in positions)
    return (
        f"✏️ <b>{len(positions)} ta savol</b> variantsiz (ochiq) deb belgilandi: "
        f"<b>{nums}</b>-savol(lar).\n\n"
        f"Har biri uchun <b>aniq (kutilgan) javobni</b> kiriting — bu keyinchalik "
        f"o'quvchi javobini tekshirishda sizga (yoki AI'ga) tayanch bo'ladi.\n\n"
        f"Har bir javobni ALOHIDA QATORDA, savollar tartibida yuboring "
        f"(jami <b>{len(positions)} qator</b>). Masalan:\n"
        f"<code>Parij\n42\nfotosintez</code>"
    )


def parse_open_ref_answers(text: str, positions: list[int]) -> dict | None:
    """Ustoz yuborgan ko'p qatorli matnni {savol_raqami: javob} lug'atiga
    aylantiradi. Qatorlar soni pozitsiyalar soniga ANIQ mos kelishi kerak —
    aks holda None qaytaradi (chaqiruvchi qayta so'raydi)."""
    lines = [ln.strip() for ln in (text or "").split("\n") if ln.strip()]
    if len(lines) != len(positions):
        return None
    return {str(pos): sanitize_plain_text(line, max_len=300) for pos, line in zip(positions, lines)}


def smart_title(full_name: str) -> str:
    """Ism-familiyani so'z boshini bosh harf qiladi, lekin apostrofdan (o'g'li kabi) keyin emas."""
    words = (full_name or "").split()
    out = []
    for word in words:
        if word:
            out.append(word[0].upper() + word[1:].lower())
    return " ".join(out) if out else "Noma'lum"


def sanitize_plain_text(s: str, max_len: int = 200) -> str:
    """
    Foydalanuvchi kiritgan erkin matnni (ism, test nomi, fan nomi va h.k.)
    xavfsizlantiradi. Bu matnlar keyinchalik o'nlab joyda parse_mode=HTML
    bilan ko'rsatiladi — agar ichida "<", ">" yoki "&" bo'lsa, Telegram
    xabarni yubormay, o'sha testni/foydalanuvchini ko'rsatishga urinilgan
    HAR BIR joyda xatolik chiqarardi. Shu belgilarni kirish nuqtasida olib
    tashlash — har bir chiqish joyini alohida ekranlashdan ko'ra ancha
    ishonchli va oson.
    """
    cleaned = re.sub(r"[<>&]", "", (s or "")).strip()
    return cleaned[:max_len]


def parse_date(date_str: str):
    try:
        return datetime.strptime(date_str.strip(), "%Y-%m-%d").date()
    except Exception:
        return None


def parse_time(time_str: str) -> dtime | None:
    try:
        return datetime.strptime(time_str.strip(), "%H:%M").time()
    except Exception:
        return None


def to_start_ts(date_str: str, time_str: str) -> tuple[int, str] | None:
    """
    Returns: (start_ts_utc, start_at_string_local) or None if invalid
    """
    d = parse_date(date_str)
    t = parse_time(time_str)
    if not d or not t:
        return None
    dt_local = datetime(d.year, d.month, d.day, t.hour, t.minute, tzinfo=TZ)
    start_at = dt_local.strftime("%Y-%m-%d %H:%M")
    return int(dt_local.timestamp()), start_at


def normalize_phone_input(text: str | None) -> str | None:
    """Qo'lda yozilgan telefon raqamini tekshiradi: 7..15 ta raqam bo'lishi
    shart (masalan "+998 90 123 45 67"). Aks holda None — avval istalgan
    matn (masalan "salom") telefon sifatida saqlanib qolardi va admin bunday
    o'quvchi bilan bog'lana olmasdi."""
    s = sanitize_plain_text(text or "", max_len=30)
    digits = re.sub(r"\D", "", s)
    if not (7 <= len(digits) <= 15):
        return None
    return s


def schedule_window_problem(start_ts_val: int, duration_min: int | None) -> str | None:
    """Rejalashtirilgan boshlanish vaqtini tekshiradi. Butun imtihon oynasi
    (boshlanish + davomiylik) allaqachon o'tib ketgan bo'lsa — bunday testni
    hech kim yecha olmaydi, shuning uchun rad etiladi."""
    dur = int(duration_min or 0) * 60
    if int(start_ts_val) + dur <= now_ts():
        return ("⚠️ Bu vaqt allaqachon o'tib ketgan — test hech kim yecha olmaydigan holatda yaratilib "
                "qolardi. Kelajakdagi (yoki hozirgi) sana va vaqtni kiriting:")
    return None


def is_certificate_eligible(exam_type: str, percent: float, grade: str | None) -> bool:
    """
    Oddiy/DTM testlar uchun: foiz CERT_MIN_PERCENT dan yuqori bo'lsa yetarli.
    Rasch testlar uchun: rasmiy daraja (grade) borligi shart, chunki
    ularda foiz va daraja jadvali bir xil bo'lmasligi mumkin (masalan, foiz
    50% dan yuqori bo'lib, lekin rasmiy daraja "yo'q" bo'lishi mumkin edi).
    """
    et = (exam_type or "simple")
    if et == "rasch":
        return grade is not None
    return float(percent or 0.0) >= float(CERT_MIN_PERCENT)


async def get_user_full_name(user_id_db: int) -> str:
    row = await POOL.fetchrow("SELECT full_name FROM users WHERE id=$1", user_id_db)
    return (row["full_name"] if row and row["full_name"] else "Noma'lum")


async def _resolve_cert_settings(cert_academy_name, cert_signee_name, cert_template, creator_id):
    """
    Test uchun saqlangan sertifikat sozlamalarini haqiqiy qiymatlarga aylantiradi.
    Agar admin/creator "mas'ul shaxs"ni maxsus kiritmagan bo'lsa — test
    yaratuvchisining ismi ishlatiladi (u ham bo'lmasa make_certificate_pdf
    o'zining standart qiymatini qo'llaydi).
    """
    academy = cert_academy_name
    signee = cert_signee_name
    if not signee and creator_id:
        row = await POOL.fetchrow("SELECT full_name FROM users WHERE telegram_id=$1", creator_id)
        if row and row["full_name"]:
            signee = row["full_name"]
    template = cert_template or DEFAULT_CERT_TEMPLATE
    return academy, signee, template


async def get_test_meta_for_certificate(test_id: int):
    return await POOL.fetchrow("""
        SELECT code, title, subject, start_mode, start_ts, duration, exam_type
        FROM tests WHERE id=$1
    """, test_id)


async def is_admin_user(telegram_id: int) -> bool:
    if telegram_id == ADMIN_ID:
        return True
    row = await POOL.fetchrow("SELECT role FROM users WHERE telegram_id=$1", telegram_id)
    return bool(row and row["role"] == "admin")


async def get_user_db_id(telegram_id: int):
    row = await POOL.fetchrow("SELECT id FROM users WHERE telegram_id=$1", telegram_id)
    return row["id"] if row else None


async def test_by_code(code: str):
    code = normalize_code(code)
    return await POOL.fetchrow("""
        SELECT id, code, title, is_free, price, duration, questions_count, file_id, answers,
               start_mode, start_ts, start_at
        FROM tests
        WHERE code=$1
    """, code)


async def test_code_exists(code: str) -> bool:
    code = normalize_code(code)
    row = await POOL.fetchrow("SELECT 1 FROM tests WHERE code=$1 LIMIT 1", code)
    return bool(row)


def is_scheduled_not_started(start_mode: str, start_ts_val: int | None) -> tuple[bool, int]:
    """
    Returns (blocked, seconds_left)
    """
    if (start_mode or "normal") != "scheduled":
        return (False, 0)
    if not start_ts_val:
        return (False, 0)
    diff = int(start_ts_val) - now_ts()
    if diff > 0:
        return (True, diff)
    return (False, 0)


def is_scheduled_expired(start_mode: str, start_ts_val: int | None, duration_min: int | None) -> tuple[bool, int]:
    """
    Returns (expired, seconds_after_end)
    Faqat scheduled test uchun ishlaydi.
    """
    if (start_mode or "normal") != "scheduled":
        return (False, 0)
    if not start_ts_val:
        return (False, 0)
    dur = int(duration_min or 0) * 60
    if dur <= 0:
        return (False, 0)

    end_ts = int(start_ts_val) + dur
    diff = now_ts() - end_ts
    if diff > 60:  # 60 sekund grace
        return (True, diff)
    return (False, 0)


# ===================== CERTIFICATE (premium PDF) =====================
NAVY = HexColor("#1a2a6c")
GOLD = HexColor("#c9a227")
GOLD_LIGHT = HexColor("#e6c866")
GOLD_DARK = HexColor("#9c7a17")
MAROON = HexColor("#7a0c1e")
CREAM = HexColor("#fdfaf1")
INK = HexColor("#2b2b2b")
GREEN = HexColor("#0f6b3a")

# ✅ Sertifikat ko'rinishlari (test yaratuvchisi tanlaydi, tests.cert_template
# ustunida saqlanadi). Har biri: ramka/sarlavha rangi (primary), oltin rangdagi
# bezaklar (trim — barcha ko'rinishlarda saqlanib qoladi, chunki oltin bezak
# har qanday rang bilan uyg'un ko'rinadi), va ism/muhr rangi (secondary).
CERT_TEMPLATES = {
    "gold": {
        "label": "🟡 Oltin (klassik)",
        "primary": HexColor("#1a2a6c"),
        "trim": HexColor("#c9a227"),
        "trim_light": HexColor("#e6c866"),
        "trim_dark": HexColor("#9c7a17"),
        "secondary": HexColor("#7a0c1e"),
    },
    "green": {
        "label": "🟢 Yashil",
        "primary": HexColor("#0d6b4f"),
        "trim": HexColor("#c9a227"),
        "trim_light": HexColor("#e6c866"),
        "trim_dark": HexColor("#9c7a17"),
        "secondary": HexColor("#6b1029"),
    },
    "blue": {
        "label": "🔵 Ko'k",
        "primary": HexColor("#1a4a8a"),
        "trim": HexColor("#c9a227"),
        "trim_light": HexColor("#e6c866"),
        "trim_dark": HexColor("#9c7a17"),
        "secondary": HexColor("#12235c"),
    },
}
DEFAULT_CERT_TEMPLATE = "gold"


def _laurel_leaf(c, x, y, angle_deg, length, width, color):
    c.saveState()
    c.translate(x, y)
    c.rotate(angle_deg)
    c.setFillColor(color)
    p = c.beginPath()
    p.moveTo(0, 0)
    p.curveTo(length * 0.25, width, length * 0.75, width, length, 0)
    p.curveTo(length * 0.75, -width, length * 0.25, -width, 0, 0)
    c.drawPath(p, fill=1, stroke=0)
    c.restoreState()


def _draw_laurel_branch(c, cx, cy, side=1, count=6, color=GOLD, height=50):
    """side=1 -> branch curves to the right, side=-1 mirrors it."""
    c.saveState()
    c.translate(cx, cy)
    c.scale(side, 1)
    c.setStrokeColor(color)
    c.setLineWidth(1.3)
    p = c.beginPath()
    p.moveTo(0, 0)
    p.curveTo(height * 0.18, height * 0.23, height * 0.26, height * 0.59, height * 0.18, height)
    c.drawPath(p, fill=0, stroke=1)
    for i in range(count):
        t = i / (count - 1)
        px = height * 0.18 * t * t * (3 - 2 * t) * 0.9 + t * 2.4
        py = height * t
        ang = 35 + 40 * t
        _laurel_leaf(c, px, py, ang, (height * 0.19) * (1 - 0.25 * t), 3.6 - t, color)
        _laurel_leaf(c, px, py, -(ang) + 10, -((height * 0.19) * (1 - 0.25 * t)), 3.6 - t, color)
    c.restoreState()


def _draw_seal(c, cx, cy, r, grade_text, ribbon_color=MAROON, gold=GOLD, tails=True):
    """Mukofot-medal uslubidagi muhr: oltin tasbeh-shaklidagi ramka, ichida
    sertifikat rangiga mos (bordo) doira, va pastda ikkita 'qaldirg'och dumi'
    lenta — mashhur 'BEST AWARD' nishonlar uslubiga o'xshash."""
    c.saveState()

    if tails:
        tail_w = r * 0.34
        top_y = cy - r * 0.62
        bot_y = cy - r * 1.55
        notch_y = cy - r * 1.28
        for sign in (-1, 1):
            p = c.beginPath()
            p.moveTo(cx + sign * tail_w * 0.15, top_y)
            p.lineTo(cx + sign * tail_w * 1.05, top_y + r * 0.05)
            p.lineTo(cx + sign * tail_w * 1.05, bot_y)
            p.lineTo(cx + sign * tail_w * 0.55, notch_y)
            p.lineTo(cx + sign * tail_w * 0.15, bot_y + r * 0.12)
            p.close()
            c.setFillColor(ribbon_color)
            c.drawPath(p, fill=1, stroke=0)
            c.setStrokeColor(gold)
            c.setLineWidth(1)
            c.drawPath(p, fill=0, stroke=1)

    # tashqi oltin "tasbeh" halqasi
    n = 24
    for i in range(n):
        ang = 2 * math.pi * i / n
        sx = cx + r * math.cos(ang)
        sy = cy + r * math.sin(ang)
        c.setFillColor(gold)
        c.circle(sx, sy, r * 0.135, fill=1, stroke=0)
    c.setStrokeColor(GOLD_DARK)
    c.setLineWidth(0.6)
    c.circle(cx, cy, r * 1.02, fill=0, stroke=1)

    # ingichka oltin halqa
    c.setFillColor(gold)
    c.circle(cx, cy, r * 0.86, fill=1, stroke=0)

    # markaziy (bordo) yuzasi
    c.setFillColor(ribbon_color)
    c.circle(cx, cy, r * 0.74, fill=1, stroke=0)
    c.setStrokeColor(gold)
    c.setLineWidth(1.3)
    c.circle(cx, cy, r * 0.74, fill=0, stroke=1)
    c.setStrokeColor(GOLD_LIGHT)
    c.setLineWidth(0.6)
    c.circle(cx, cy, r * 0.64, fill=0, stroke=1)

    c.setFillColor(gold)
    if grade_text:
        c.setFont("Times-Bold", r * 0.62)
        c.drawCentredString(cx, cy - r * 0.24, grade_text)
    else:
        c.setFont("Times-Bold", r * 0.42)
        c.drawCentredString(cx, cy - r * 0.16, "★")

    c.setFont("Times-Bold", r * 0.16)
    c.drawCentredString(cx - r * 0.5, cy + r * 0.34, "★")
    c.drawCentredString(cx + r * 0.5, cy + r * 0.34, "★")

    c.restoreState()


def _draw_corner_ornament(c, x, y, quad, size=34, color=GOLD):
    sx, sy = quad
    c.saveState()
    c.translate(x, y)
    c.scale(sx, sy)
    c.setStrokeColor(color)
    c.setLineWidth(1.6)
    c.line(0, 0, size, 0)
    c.line(0, 0, 0, size)
    p = c.beginPath()
    p.moveTo(size * 0.15, 0)
    p.curveTo(size * 0.6, size * 0.05, size * 0.55, size * 0.55, size * 0.65, size * 0.75)
    c.drawPath(p, fill=0, stroke=1)
    c.setFillColor(color)
    c.circle(size * 0.15, 0, 2.6, fill=1, stroke=0)
    c.circle(0, size * 0.15, 2.6, fill=1, stroke=0)
    c.circle(size * 0.65, size * 0.75, 3.2, fill=1, stroke=0)
    c.restoreState()


def make_certificate_pdf(
        full_name: str,
        subject: str,
        test_code: str,
        test_title: str,
        exam_date_str: str,
        percent: float = 0.0,
        correct_count: int = 0,
        total_questions: int = 0,
        rasch_ball: float | None = None,
        grade: str | None = None,
        out_path: str = "certificate.pdf",
        academy_name: str | None = None,
        signee_name: str | None = None,
        template: str = DEFAULT_CERT_TEMPLATE,
):
    """Tantanavor, albom yo'nalishdagi (landscape) premium sertifikat.

    academy_name/signee_name/template — test yaratuvchisi tanlagan (yoki
    standart) sertifikat sozlamalari; har bir test o'zining sozlamalari bilan
    boshqa ko'rinishda sertifikat chiqarishi mumkin.
    """
    tpl = CERT_TEMPLATES.get(template) or CERT_TEMPLATES[DEFAULT_CERT_TEMPLATE]
    PRIMARY = tpl["primary"]
    TRIM = tpl["trim"]
    TRIM_LIGHT = tpl["trim_light"]
    TRIM_DARK = tpl["trim_dark"]
    SECONDARY = tpl["secondary"]

    academy = (academy_name or ACADEMY_NAME).strip()
    signee = (signee_name or "Administratsiya").strip()

    page = landscape(A4)
    c = canvas.Canvas(out_path, pagesize=page)
    w, h = page

    # ---------- Background ----------
    c.setFillColor(CREAM)
    c.rect(0, 0, w, h, fill=1, stroke=0)

    # faint repeating watermark
    c.saveState()
    c.setFillColor(TRIM)
    c.setFillAlpha(0.032)
    c.setFont("Times-Bold", 22)
    label = academy.upper()
    c.rotate(22)
    step_x, step_y = 300, 130
    for row in range(-5, 8):
        for col in range(-2, 6):
            c.drawCentredString(col * step_x, row * step_y, label)
    c.restoreState()

    # ---------- Frames ----------
    c.setStrokeColor(TRIM)
    c.setLineWidth(3)
    c.rect(24, 24, w - 48, h - 48)

    c.setStrokeColor(PRIMARY)
    c.setLineWidth(1.6)
    c.rect(34, 34, w - 68, h - 68)

    c.setStrokeColor(TRIM)
    c.setLineWidth(0.8)
    c.rect(40, 40, w - 80, h - 80)

    m = 44
    _draw_corner_ornament(c, m, m, (1, 1), color=TRIM)
    _draw_corner_ornament(c, w - m, m, (-1, 1), color=TRIM)

    # ---------- Logo (top-left) ----------
    try:
        if CERT_LOGO_PATH and os.path.exists(CERT_LOGO_PATH):
            logo = ImageReader(CERT_LOGO_PATH)
            c.drawImage(logo, 68, h - 148, width=100, height=100, mask="auto", preserveAspectRatio=True)
    except Exception:
        pass

    # ---------- Title ----------
    c.setFont("Times-Bold", 46)
    c.setFillColor(PRIMARY)
    c.drawCentredString(w / 2, h - 108, "S E R T I F I K A T")

    ry = h - 128
    c.setStrokeColor(TRIM)
    c.setLineWidth(1.2)
    c.line(w / 2 - 150, ry, w / 2 - 14, ry)
    c.line(w / 2 + 14, ry, w / 2 + 150, ry)
    c.saveState()
    c.translate(w / 2, ry)
    c.rotate(45)
    c.setFillColor(TRIM)
    c.rect(-5, -5, 10, 10, fill=1, stroke=0)
    c.restoreState()

    c.setFont("Times-Italic", 15)
    c.setFillColor(HexColor("#5a5a5a"))
    c.drawCentredString(w / 2, h - 150, f"Ushbu hujjat {academy} tomonidan taqdim etiladi")

    # ---------- Name with laurel ----------
    name_y = h - 205
    display_name = smart_title(full_name)
    name_font_size = 38
    max_name_w = w - 340
    while name_font_size > 20 and c.stringWidth(display_name, "Times-BoldItalic", name_font_size) > max_name_w:
        name_font_size -= 2
    c.setFont("Times-BoldItalic", name_font_size)
    c.setFillColor(SECONDARY)
    c.drawCentredString(w / 2, name_y, display_name)

    name_w = c.stringWidth(display_name, "Times-BoldItalic", name_font_size)
    laurel_offset = min(max(name_w / 2 + 16, 120), 260)
    _draw_laurel_branch(c, w / 2 - laurel_offset, name_y - 12, side=-1, color=TRIM)
    _draw_laurel_branch(c, w / 2 + laurel_offset, name_y - 12, side=1, color=TRIM)

    c.setStrokeColor(TRIM)
    c.setLineWidth(1)
    p = c.beginPath()
    p.moveTo(w / 2 - 130, name_y - 16)
    p.curveTo(w / 2 - 40, name_y - 24, w / 2 + 40, name_y - 24, w / 2 + 130, name_y - 16)
    c.drawPath(p, fill=0, stroke=1)

    # ---------- Body text ----------
    c.setFont("Times-Roman", 15)
    c.setFillColor(INK)
    c.drawCentredString(w / 2, h - 245, "ga o'quv dasturi doirasidagi test sinovlarida muvaffaqiyatli qatnashib,")
    c.drawCentredString(w / 2, h - 267, "qayd etgan yuqori natijasi va ilm yo'lidagi intiluvchanligi uchun taqdim etiladi")

    # ---------- Details block ----------
    box_top = h - 300
    box_h = 132
    c.setFillColor(HexColor("#f4ecd8"))
    c.roundRect(w / 2 - 230, box_top - box_h, 460, box_h, 6, fill=1, stroke=0)
    c.setStrokeColor(TRIM)
    c.setLineWidth(0.7)
    c.roundRect(w / 2 - 230, box_top - box_h, 460, box_h, 6, fill=0, stroke=1)

    y = box_top - 24
    lh = 25
    rows = [
        ("Fan:", subject or "Noma'lum"),
        ("Test nomi:", test_title),
        ("Sana:", exam_date_str),
    ]
    for label, val in rows:
        c.setFont("Times-Bold", 13.5)
        c.setFillColor(PRIMARY)
        c.drawRightString(w / 2 - 8, y, label)
        c.setFont("Times-Roman", 13.5)
        c.setFillColor(INK)
        c.drawString(w / 2 + 8, y, str(val))
        y -= lh

    y -= 8
    c.setFont("Times-Bold", 15)
    c.setFillColor(GREEN)
    c.drawCentredString(w / 2, y, f"NATIJA:  {correct_count} / {total_questions}   —   {percent:.1f}%")

    # ---------- Seal / grade medallion ----------
    seal_cx = w - 148
    seal_cy = 118
    seal_r = 34
    _draw_seal(c, seal_cx, seal_cy, seal_r, grade if grade else None, ribbon_color=SECONDARY, gold=TRIM)
    if rasch_ball is not None:
        c.setFont("Times-Italic", 10.5)
        c.setFillColor(HexColor("#5a5a5a"))
        c.drawCentredString(seal_cx, seal_cy - seal_r - 16, f"Rasch ball: {rasch_ball:.1f}")

    # ---------- Signature ----------
    sig_x = 95
    sig_y = 96
    c.setFont("Times-Italic", 13)
    c.setFillColor(INK)
    c.drawString(sig_x, sig_y, "Mas'ul shaxs:")
    c.setFont("Times-Bold", 14)
    c.drawString(sig_x + 90, sig_y, signee)
    c.setStrokeColor(black)
    c.setLineWidth(0.8)
    c.line(sig_x, sig_y - 8, sig_x + 220, sig_y - 8)
    c.setFont("Times-Roman", 9)
    c.setFillColor(HexColor("#7a7a7a"))
    c.drawCentredString(sig_x + 110, sig_y - 20, "imzo")

    try:
        if CERT_SIGN_PATH and os.path.exists(CERT_SIGN_PATH):
            sign = ImageReader(CERT_SIGN_PATH)
            c.drawImage(sign, sig_x + 40, sig_y - 5, width=140, height=55, mask="auto", preserveAspectRatio=True)
    except Exception:
        pass

    # ---------- Footer ----------
    c.setFont("Times-Roman", 9.5)
    c.setFillColor(HexColor("#6b6b6b"))
    c.drawCentredString(w / 2, 56, f"© 2026 {academy} • Raqamli tizim orqali tasdiqlangan")
    c.setFont("Times-Bold", 9.5)
    c.drawCentredString(w / 2, 44, f"Sertifikat kodi: {test_code}-{(full_name or 'X').strip()[:1].upper()}")

    c.showPage()
    c.save()


# ===================== RUSH MODULE (Finalize after exam end) =====================

# p smoothing (Bayes) parametrlari: p = (correct + A) / (total + A + B) — bu
# silliqlash tufayli Rasch ENDI GURUH KATTALIGIDAN QAT'I NAZAR (1 kishidan
# minglab kishigacha) bir xil formula bilan ishonchli ishlaydi, alohida
# "kichik guruh" rejimi yoki alohida "Maxsus" moduli shart emas.
RUSH_SMOOTH_A = 1.0
RUSH_SMOOTH_B = 1.0

# ✅ Savol qiyinligiga qarab OG'IRLIKLI baholash. Har bir savol uchun guruhning
# necha foizi to'g'ri topgani (p) hisoblanadi:
#   - to'g'ri javob uchun ball = 1 + REWARD_K*(1-p)  -> kam kishi yechgan (qiyin)
#     savolni to'g'ri qilsangiz KO'PROQ ball beradi;
#   - xato javob uchun jarima = PENALTY_K*p           -> ko'pchilik yechgan (oson)
#     savolni XATO qilsangiz ball yo'qotasiz.
# Diqqat: klassik Rasch/1PL modelida xom ball bir xil bo'lsa, QAYSI aniq savolda
# xato qilingani natijaga umuman ta'sir qilmaydi (bu "raw score — yetarli
# statistika" degan rasman isbotlangan xossa). Bu yerda esa ATAYLAB shundan
# chetga chiqamiz. MUHIM: REWARD_K va PENALTY_K bir xil bo'lsa, ikkalasi
# matematik jihatdan bir-birini aniq bekor qiladi (sinovdan o'tkazilgan) — shu
# sabab REWARD_K ataylab PENALTY_K'dan katta tanlangan, shundagina xuddi shu
# xom balga ega ikki kishidan qiyin savolni topib, oson savolni yo'qotgani —
# oson savolni topib, qiyin savolni yo'qotganidan yuqoriroq ball oladi.
RUSH_REWARD_K = 2.5
RUSH_PENALTY_K = 1.0


# Milliy sertifikat Rasch ball chegaralari (BBA/UZBMB e'loniga mos)
# 70+ A+, 65-69.9 A, 60-64.9 B+, 55-59.9 B, 50-54.9 C+, 46-49.9 C
# 46 dan past -> sertifikat/daraja yo'q
def grade_from_rasch_ball(ball: float) -> str | None:
    ball = float(ball or 0.0)
    if ball >= 70:
        return "A+"
    if ball >= 65:
        return "A"
    if ball >= 60:
        return "B+"
    if ball >= 55:
        return "B"
    if ball >= 50:
        return "C+"
    if ball >= 46:
        return "C"
    return None


def _clamp(x: float, a: float, b: float) -> float:
    return a if x < a else b if x > b else x


def item_difficulty_p(correct: int, total: int, a: float = RUSH_SMOOTH_A, b: float = RUSH_SMOOTH_B) -> float:
    """Guruhning shu savolni to'g'ri topgan ulushi (kichik guruhlarda ham
    ishonchli bo'lishi uchun Bayes silliqlash bilan: p = (to'g'ri+A)/(jami+A+B))."""
    total = int(total or 0)
    correct = int(correct or 0)
    if total <= 0:
        return 0.5
    p = (correct + float(a)) / (total + float(a) + float(b))
    return _clamp(p, 0.01, 0.99)


def weighted_item_points(
    is_correct: bool, p_i: float, reward_k: float = RUSH_REWARD_K, penalty_k: float = RUSH_PENALTY_K
) -> float:
    p_i = _clamp(float(p_i), 0.0, 1.0)
    if is_correct:
        return 1.0 + reward_k * (1.0 - p_i)
    return -penalty_k * p_i


def weighted_max_points(p_i: float, reward_k: float = RUSH_REWARD_K) -> float:
    return 1.0 + reward_k * (1.0 - _clamp(float(p_i), 0.0, 1.0))


async def _gemini_check_rasch_sanity(
    qcount: int, n_solvers: int, p_list: list[float], score_pairs: list[tuple[int, float]]
) -> dict:
    """Guruhning vaznli Rasch ballarini o'quvchilarga E'LON QILISHDAN OLDIN
    AI orqali "haqiqatga yaqinmi" deb so'nggi tekshiruvdan o'tkazadi.

    Muhim dizayn qarori: AI hech qachon aniq sonlarni o'zi o'ylab topib
    yozmaydi (bu determinizm/adolat kafolatini buzardi — bir xil xom balga
    ega ikki kishi turli sessiyada turlicha yakuniy ball olib qolishi
    mumkin edi). AI faqat ikkita ALLAQACHON TEKSHIRILGAN, deterministik
    formuladan qay birini ishlatishni tanlaydi: (1) standart vaznli
    (item-difficulty) formula, yoki (2) anomaliya aniqlansa — oddiy
    chiziqli foiz->ball formulasiga qaytish. Gemini mavjud bo'lmasa yoki
    xato bersa, har doim standart formula ishonchli deb qabul qilinadi —
    natija hech qachon AI mavjudligiga bog'liq bo'lib qolmaydi."""
    default = {"ok": True, "reason": ""}
    if not GEMINI_API_KEY or not score_pairs:
        return default

    if len(score_pairs) <= 40:
        sample = score_pairs
    else:
        step = len(score_pairs) / 40.0
        sample = [score_pairs[int(i * step)] for i in range(40)]
    diff_summary = (
        f"eng oson (p) = {max(p_list):.2f}, eng qiyin (p) = {min(p_list):.2f}, "
        f"o'rtacha p = {sum(p_list)/len(p_list):.2f}"
    ) if p_list else "yo'q"
    sample_lines = "\n".join(f"  {raw}/{qcount} xom to'g'ri javob -> {ball:.1f} ball (0-75 shkala)"
                             for raw, ball in sample)

    prompt = (
        "Siz psixometrika (Rasch/IRT) mutaxassisisiz. Milliy sertifikat uslubidagi bitta "
        "guruh imtihonining vaznli (savol qiyinligiga qarab) hisoblangan ballari statistikasi "
        "berilgan — natijalar o'quvchilarga E'LON QILINISHIDAN OLDIN oxirgi tekshiruv.\n\n"
        f"Savollar soni: {qcount}, yechganlar soni: {n_solvers}\n"
        f"Savol qiyinligi (p = to'g'ri topganlar ulushi): {diff_summary}\n"
        f"Xom to'g'ri javoblar soni -> hisoblangan ball namunasi:\n{sample_lines}\n\n"
        "MUHIM KONTEKST — bular XATO EMAS, ataylab shunday qilingan: "
        "(a) qiyin savolni to'g'ri topgan ko'proq mukofotlanadi, oson savolni xato qilgan "
        "ko'proq jarimalanadi — shu sabab BIR XIL xom balga ega ikki kishi turli yakuniy ball "
        "olishi mumkin, bu normal; (b) ball 0-75 oralig'ida.\n\n"
        "Faqat quyidagi HAQIQIY muammolarni qidiring: (1) ko'proq xom to'g'ri javob berganlar "
        "TIZIMLI ravishda kamroq yechganlardan past ball olayotgan bo'lsa (kuchli teskari "
        "monotonlik buzilishi, tasodifiy emas), (2) deyarli barcha savol qiyinligi 0 yoki 1 ga "
        "juda yaqin bo'lib javob kaliti butunlay noto'g'ri kiritilganidan darak bersa, "
        "(3) ballar 0-75 oralig'idan chiqib ketgan yoki mantiqsiz bo'lsa.\n\n"
        'Faqat shu JSON formatda javob bering, boshqa hech narsa yozmang: '
        '{"ok": true yoki false, "reason": "<agar ok=false bo\'lsa qisqa o\'zbekcha sabab, aks holda bo\'sh satr>"}'
    )

    # ✅ Kalit URL'da emas, header'da — aiohttp xatolari (masalan JSON o'rniga
    # HTML qaytsa) to'liq URL'ni logga yozadi, kalit server logiga tushardi.
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{GEMINI_MODEL}:generateContent"
    payload = {"contents": [{"parts": [{"text": prompt}]}]}
    try:
        async with aiohttp.ClientSession(timeout=aiohttp.ClientTimeout(total=30)) as session:
            async with session.post(url, json=payload, headers={"x-goog-api-key": GEMINI_API_KEY}) as resp:
                if resp.status != 200:
                    logging.error(f"Gemini Rasch tekshiruvi xatosi: {resp.status} {await resp.text()}")
                    return default
                data = await resp.json()
        raw_text = data["candidates"][0]["content"]["parts"][0]["text"]
        obj = _extract_json_object(raw_text)
        if not obj or "ok" not in obj:
            return default
        return {
            "ok": bool(obj.get("ok", True)),
            "reason": sanitize_plain_text(str(obj.get("reason") or ""), max_len=500),
        }
    except Exception:
        logging.exception("Gemini Rasch sanity tekshiruvida xatolik")
        return default


def fallback_rasch_ball_from_raw_percent(raw_percent_0_100: float) -> float:
    """Juda kichik N bo'lsa vaznli baholashni ishlatmaymiz, foizni chiziqli 0..75 shkalaga o'tkazamiz."""
    p = max(0.0, min(100.0, float(raw_percent_0_100 or 0.0)))
    return (p / 100.0) * 75.0


async def get_test_end_ts(test_id: int) -> int:
    """
    Testning umumiy tugash vaqtini hisoblaydi:
    - scheduled bo'lsa: start_ts + duration*60
    - normal bo'lsa: earliest finished_ts + duration*60 (fallback, faqat "simple/dtm" uchun amalda ishlatiladi)
    """
    t = await POOL.fetchrow("SELECT start_mode, start_ts, duration, is_closed FROM tests WHERE id=$1", test_id)
    if not t:
        return 0

    # ✅ Admin/creator testni qo'lda "yakunlagan" bo'lsa — asl jadvalini
    # (start_mode/start_ts/duration) buzmasdan, hisoblashni DARHOL ruxsat
    # beramiz (Rasch kutayotgan bo'lsa ham). Test qayta ochilsa,
    # asl belgilangan vaqti butunligicha qoladi.
    if int(t["is_closed"] or 0) == 1:
        return now_ts()

    start_mode, start_ts_val, duration_min = t["start_mode"], t["start_ts"], t["duration"]
    duration_sec = int(duration_min or 0) * 60

    if (start_mode or "normal") == "scheduled" and int(start_ts_val or 0) > 0:
        return int(start_ts_val) + duration_sec

    r = await POOL.fetchrow("""
        SELECT MIN(finished_ts) AS m
        FROM results
        WHERE test_id=$1 AND finished_ts>0
    """, test_id)
    first_finish = int((r["m"] or 0) if r else 0)
    if first_finish <= 0:
        return 0
    return first_finish + duration_sec


async def remaining_time_text(test_id: int) -> str:
    """Guruh natijasi (Rasch) hali chiqmagan bo'lsa, "qancha vaqt
    qoldi" degan tushunarli matn qaytaradi — foydalanuvchi noaniqlikda
    kutib qolmasligi uchun."""
    end_ts = await get_test_end_ts(test_id)
    if end_ts <= 0:
        return "hali aniq emas — test hali boshlanmagan"
    remain = end_ts - now_ts()
    if remain <= 0:
        pending = await POOL.fetchval(
            "SELECT COUNT(*) FROM results WHERE test_id=$1 AND COALESCE(open_pending,0)>0", test_id)
        if pending:
            return "test vaqti tugagan — ustoz ochiq (yozma) javoblarni tekshirgach e'lon qilinadi"
        return "tez orada (hisoblash navbatda)"
    return f"{fmt_seconds(remain)} qoldi"


OPEN_GRADE_REMIND_EVERY_SEC = 6 * 3600


async def _remind_open_grading(bot: "Bot", test_id: int, pending_n: int) -> None:
    """Rasch natijalari ochiq javoblar tekshirilmagani sabab kutib turganda
    test yaratuvchisiga (eng ko'pi bilan har 6 soatda bir marta) eslatma."""
    t = await POOL.fetchrow(
        "SELECT code, title, creator_id, COALESCE(open_grade_reminded_ts,0) AS rts FROM tests WHERE id=$1", test_id)
    if not t or now_ts() - int(t["rts"]) < OPEN_GRADE_REMIND_EVERY_SEC:
        return
    await POOL.execute("UPDATE tests SET open_grade_reminded_ts=$1 WHERE id=$2", now_ts(), test_id)
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="✍️ Javoblarni tekshirish", callback_data=f"ogl_{test_id}")],
        [InlineKeyboardButton(text="⚡ Tekshirmasdan hisoblash", callback_data=f"ogforce_{test_id}")],
    ])
    try:
        await bot.send_message(
            t["creator_id"] or ADMIN_ID,
            f"⏳ <b>{_h(t['title'])}</b> (<code>{t['code']}</code>) — test vaqti tugadi.\n\n"
            f"Rasch natijalari e'lon qilinishi uchun <b>{pending_n}</b> ta o'quvchining ochiq (yozma) "
            f"javoblarini tekshirish kerak. Tekshirib bo'lgach, natijalar avtomatik hisoblanadi.",
            parse_mode=ParseMode.HTML, reply_markup=kb
        )
    except Exception:
        logging.exception(f"Ochiq javob eslatmasini yuborib bo'lmadi (test={test_id})")


RASCH_SUBMIT_GRACE_SEC = 60  # process_test_submission'dagi muhlat bilan bir xil


async def finalize_rush_for_test_if_ready(test_id: int, bot: "Bot | None" = None, force: bool = False) -> bool:
    end_ts = await get_test_end_ts(test_id)
    if end_ts <= 0:
        return False
    # ✅ Test vaqti tugagach o'quvchilarga yana 60 soniya topshirish muhlati
    # beriladi — hisoblash shu muhlat tugagach boshlanadi. Avval natijalar
    # rasmiy tugash lahzasidayoq hisoblanib, muhlat ichida topshirganlarning
    # Rasch bali hech qachon hisoblanmay qolardi.
    is_closed = await POOL.fetchval("SELECT COALESCE(is_closed,0) FROM tests WHERE id=$1", test_id)
    grace = 0 if int(is_closed or 0) == 1 else RASCH_SUBMIT_GRACE_SEC
    if now_ts() < end_ts + grace:
        return False

    async with POOL.acquire() as db:
        row = await db.fetchrow("SELECT rush_finalized FROM tests WHERE id=$1", test_id)
        if row and int(row["rush_finalized"] or 0) == 1:
            return True

        t = await db.fetchrow("SELECT questions_count, answers FROM tests WHERE id=$1", test_id)
        if not t:
            return False

        qcount = int(t["questions_count"] or 0)
        correct_key = normalize_answers(t["answers"] or "")
        qcount = min(qcount, len(correct_key))
        if qcount <= 0:
            return False

        # ✅ Ochiq (x) savollar ham Rasch hisobiga kiradi — lekin buning uchun
        # ustoz ularni tekshirib bo'lgan bo'lishi kerak. Tekshirilmagan javob
        # bor ekan, natijalar e'lon qilinmaydi (aks holda ochiq savollar
        # hisobsiz qolib, ballar noto'g'ri chiqardi). force=True — ustoz
        # "tekshirmasdan hisoblash"ni tanlagan: tekshirilmaganlar 0 hisoblanadi.
        pending_n = await db.fetchval(
            "SELECT COUNT(*) FROM results WHERE test_id=$1 AND COALESCE(open_pending,0)>0", test_id)
        if pending_n and not force:
            if bot is not None:
                await _remind_open_grading(bot, test_id, int(pending_n))
            return False

        rows = await db.fetch("""
            SELECT id, user_answers, open_grades
            FROM results
            WHERE test_id=$1 AND user_answers IS NOT NULL AND LENGTH(user_answers) >= $2
        """, test_id, qcount)

        if not rows:
            return False

        # har bir yechuvchi uchun savollar bo'yicha to'g'ri/xato vektori
        def _correctness(ans: str, grades: dict) -> list[bool]:
            out = []
            for i in range(qcount):
                if correct_key[i] == "x":
                    g = grades.get(str(i + 1))
                    out.append(g is not None and int(g) == 1)
                else:
                    out.append(ans[i] == correct_key[i])
            return out

        solver_rows = []  # (row_id, correctness, grades)
        for r in rows:
            ans = normalize_answers(r["user_answers"] or "")
            if len(ans) < qcount:
                continue
            grades = _load_json_dict(r["open_grades"])
            if force:
                for i in range(qcount):
                    if correct_key[i] == "x":
                        grades.setdefault(str(i + 1), 0)
            solver_rows.append((int(r["id"]), _correctness(ans, grades), grades))

        n_solvers = len(solver_rows)

        # ✅ Guruh kattaligidan qat'i nazar (1 kishidan minglab kishigacha)
        # BIR XIL vaznli (savol qiyinligiga qarab) formula ishlatiladi — alohida
        # "kichik guruh" chiziqli fallback endi shart emas, chunki
        # item_difficulty_p'dagi Bayes silliqlash (A=B=1) kam ma'lumotli
        # holatlarda ham p'ni 0.5'ga yaqin ushlab, haddan tashqari keskin
        # baholarning oldini avtomatik oladi (N=1 bo'lsa ham p 0.33-0.67
        # oralig'ida qoladi, hech qachon 0 yoki 1 ga sakramaydi).
        #
        # --- savol qiyinligiga qarab vaznli baholash. Ochiq (x) savollar ham
        # oddiy savol kabi qatnashadi — ularning to'g'ri/xatoligini ustoz
        # belgilagan (open_grades) ---
        correct_counts = [0] * qcount
        total_counts = [0] * qcount

        for _rid, corr, _g in solver_rows:
            for i in range(qcount):
                total_counts[i] += 1
                if corr[i]:
                    correct_counts[i] += 1

        p_list = [item_difficulty_p(correct_counts[i], total_counts[i]) for i in range(qcount)]
        max_points = sum(weighted_max_points(p_list[i]) for i in range(qcount))

        # ✅ Hali bazaga yozmasdan — har bir natija uchun XOM to'g'ri javoblar
        # soni va vaznli ballni oldindan hisoblab qo'yamiz, shu holat asosida
        # natijalarni e'lon qilishdan oldin AI'dan yakuniy sanity-tekshiruv
        # so'raymiz (pastda).
        candidates = []  # (row_id, raw_correct, rasch_ball, rasch_percent, grade)
        for row_id, corr, _g in solver_rows:
            raw_correct = sum(1 for c in corr if c)
            earned = sum(weighted_item_points(corr[i], p_list[i]) for i in range(qcount))
            ratio = _clamp(earned / max_points, 0.0, 1.0) if max_points > 0 else 0.0
            rasch_ball = ratio * 75.0
            rasch_percent = ratio * 100.0
            grade = grade_from_rasch_ball(rasch_ball)
            candidates.append((row_id, raw_correct, rasch_ball, rasch_percent, grade))

        # ✅ MUHIM: natijalar o'quvchilarga E'LON QILINISHIDAN OLDIN AI orqali
        # oxirgi tekshiruv — hisoblangan ballar guruh statistikasiga (xom
        # to'g'ri javoblar soniga) mos, "haqiqatga yaqin"mi? Shubhali
        # anomaliya topilsa, shu test uchun oddiy va shaffof chiziqli
        # foiz->ball formulasiga qaytamiz va adminga xabar beramiz. AI hech
        # qachon sonlarni o'zi o'ylab topmaydi — faqat ikkita allaqachon
        # tekshirilgan deterministik formuladan birini tanlaydi.
        score_pairs = sorted({(rc, rb) for _, rc, rb, _, _ in candidates})
        sanity = await _gemini_check_rasch_sanity(qcount, n_solvers, p_list, score_pairs)

        if not sanity.get("ok", True):
            logging.warning(f"Rasch AI sanity: test={test_id} anomaliya -> {sanity.get('reason')}")
            candidates = []
            for row_id, corr, _g in solver_rows:
                score = sum(1 for c in corr if c)
                raw_percent = (score / qcount * 100.0) if qcount > 0 else 0.0
                rasch_ball = fallback_rasch_ball_from_raw_percent(raw_percent)
                rasch_percent = (rasch_ball / 75.0) * 100.0
                grade = grade_from_rasch_ball(rasch_ball)
                candidates.append((row_id, score, rasch_ball, rasch_percent, grade))

            if bot is not None:
                try:
                    await bot.send_message(
                        ADMIN_ID,
                        f"⚠️ <b>Rasch AI tekshiruvi anomaliya topdi</b> (test_id={test_id})\n\n"
                        f"Sabab: {sanity.get('reason') or '-'}\n\n"
                        f"Xavfsizlik uchun bu test natijalari vaznli formula o'rniga "
                        f"oddiy chiziqli foiz→ball formulasi bilan hisoblandi.",
                        parse_mode=ParseMode.HTML
                    )
                except Exception:
                    logging.exception("Rasch anomaliya xabarini adminga yuborib bo'lmadi")

        grades_by_row = {rid: g for rid, _c, g in solver_rows}
        async with db.transaction():
            for row_id, raw, rasch_ball, rasch_percent, grade in candidates:
                await db.execute("""
                    UPDATE results
                    SET rasch_ball=$1, rasch_percent=$2, grade=$3,
                        score=$4, percent=$5, open_grades=$6, open_pending=0
                    WHERE id=$7
                """, float(rasch_ball), float(rasch_percent), grade,
                    int(raw), (raw / qcount * 100.0) if qcount else 0.0,
                    json.dumps(grades_by_row.get(row_id) or {}, ensure_ascii=False), row_id)

            await db.execute(
                "UPDATE tests SET rush_finalized=1, rush_finalized_ts=$1 WHERE id=$2",
                now_ts(), test_id
            )
        return True


# ===================== CEFR / IELTS / SAT MODULE (approximate conversion) =====================
# ✅ Bu uchala format ham (rasch'dan farqli) guruh natijasiga bog'liq emas —
# har bir o'quvchining darajasi FAQAT o'zining foizidan darhol hisoblanadi,
# shuning uchun "scheduled" rejim shart emas va natija/sertifikat darhol
# (dtm kabi) chiqariladi. Konvertatsiya taxminiy — rasmiy CEFR/IELTS/SAT
# ballari emas, buni foydalanuvchiga ko'rsatiladigan har bir matnda alohida
# ta'kidlaymiz.
def cefr_level_from_percent(percent: float) -> str:
    """Chegaralar Cambridge/IELTS rasman e'lon qilgan IELTS band<->CEFR
    moslik jadvalidan olingan (taxminan: band<3.0->A1, 3.0-3.5->A2,
    4.0-5.0->B1, 5.5-6.5->B2, 7.0-8.0->C1, 8.5+->C2 — manba: ielts.org va
    Cambridge English), so'ng shu bandlar botning o'zidagi IELTS_READING_TABLE
    orqali foizga o'tkazilgan — shunday qilib bot ichidagi CEFR va IELTS
    (tezkor) natijalari BIR XIL foizda BIR-BIRIGA ZID kelmaydi (avval, masalan,
    70% da CEFR "C1" deb, IELTS esa mos keladigan 6.0 bandni rasmiy jadvalga
    ko'ra "B2" deb ko'rsatardi — endi ikkalasi ham "B2" deydi)."""
    p = float(percent or 0.0)
    if p < 20:
        return "A1"
    if p < 32.5:
        return "A2"
    if p < 57.5:
        return "B1"
    if p < 80:
        return "B2"
    if p < 92.5:
        return "C1"
    return "C2"


# ✅ IELTS Reading/Listening uchun ommaviy nashr etilgan (Cambridge/British
# Council tayyorlov materiallarida keng qo'llaniladigan) taxminiy xom
# ball -> band konvertatsiya jadvallari. Ikkalasi 40 savolga mo'ljallangan;
# boshqa sonli test uchun natija proportsional 40 talik shkalaga o'tkaziladi.
# Rasmiy IELTS har bir test shakli uchun alohida tenglashtiradi (buni ochiq
# takrorlab bo'lmaydi), shuning uchun bu — eng yaqin, umumiy qabul qilingan
# taxmin, aniq rasmiy natija emas.
IELTS_LISTENING_TABLE = [
    (39, 9.0), (37, 8.5), (35, 8.0), (32, 7.5), (30, 7.0),
    (26, 6.5), (23, 6.0), (18, 5.5), (16, 5.0), (13, 4.5),
    (10, 4.0), (8, 3.5), (6, 3.0), (4, 2.5), (0, 2.0),
]
IELTS_READING_TABLE = [
    (39, 9.0), (37, 8.5), (36, 8.0), (34, 7.5), (32, 7.0),
    (30, 6.5), (27, 6.0), (23, 5.5), (19, 5.0), (15, 4.5),
    (13, 4.0), (10, 3.5), (8, 3.0), (6, 2.5), (0, 2.0),
]


def _ielts_band_from_table(correct: int, total: int, table: list[tuple[int, float]]) -> float:
    total = max(1, int(total or 0))
    correct = max(0, int(correct or 0))
    scaled = correct * 40.0 / total  # jadval 40 savolga mo'ljallangan
    for threshold, band in table:
        if scaled >= threshold:
            return band
    return table[-1][1]


def ielts_reading_band(correct: int, total: int) -> float:
    return _ielts_band_from_table(correct, total, IELTS_READING_TABLE)


def ielts_listening_band(correct: int, total: int) -> float:
    return _ielts_band_from_table(correct, total, IELTS_LISTENING_TABLE)


def ielts_band_from_percent(percent: float) -> str:
    """Tezkor (bitta baholash, bosqichsiz) IELTS rejimi uchun — foizni 40
    savolga teng ekvivalent to'g'ri javobga aylantirib, xuddi shu Reading
    jadvalidan foydalanadi."""
    p = max(0.0, min(100.0, float(percent or 0.0)))
    correct_equiv = round(p / 100.0 * 40)
    band = _ielts_band_from_table(correct_equiv, 40, IELTS_READING_TABLE)
    return f"{band:.1f}"


# ✅ Rasmiy SAT konvertatsiyasi har bir test shakli/sanasi uchun alohida
# tenglashtiriladi va ochiq e'lon qilinmaydi (digital SAT — moslashuvchan),
# shuning uchun buni aniq takrorlab bo'lmaydi. Quyidagi tayanch nuqtalar
# (foiz -> normallashtirilgan 0..1 ball) College Board'ning ilgari e'lon
# qilingan Math bo'limi xom ball->shkala ball jadvalidan olingan (masalan
# 54/54->800, 49/54->710, 40/54->630, 25-26/54->500, 10/54->370, 5/54->310,
# 0/54->220 — manba: PrepScholar'da keltirilgan rasmiy jadval). Bu haqiqiy
# egri chiziq O'RTADA taxminan CHIZIQLI, lekin ikkala UCHIDA (ayniqsa
# yuqorida) SEZILARLI TIKROQ ekanini ko'rsatadi — eng yuqori ballar
# oralig'ida har bir qo'shimcha xato o'rtachaga qaraganda ko'proq ball
# yo'qotadi. Bu oddiy simmetrik sigmoiddan (ikkala uchida ham tekislanadigan)
# tubdan farqli, shuning uchun to'g'ridan-to'g'ri shu tayanch nuqtalar
# orqali chiziqli interpolyatsiya qilinadi.
_SAT_CURVE_POINTS = [
    (0.00, 0.03), (0.09, 0.18), (0.19, 0.28), (0.47, 0.50),
    (0.74, 0.72), (0.91, 0.85), (1.00, 1.00),
]


def _sat_normalized_score(percent: float) -> float:
    p = max(0.0, min(100.0, float(percent or 0.0))) / 100.0
    pts = _SAT_CURVE_POINTS
    norm = pts[-1][1]
    for i in range(len(pts) - 1):
        x0, y0 = pts[i]
        x1, y1 = pts[i + 1]
        if x0 <= p <= x1:
            frac = (p - x0) / (x1 - x0) if x1 > x0 else 0.0
            norm = y0 + frac * (y1 - y0)
            break
    return norm


def sat_score_from_percent(percent: float) -> int:
    """Tezkor (bitta baholash) SAT rejimi uchun — bitta aralash foizdan
    to'g'ridan-to'g'ri 400-1600 kompozit ballga o'tkazadi."""
    score = 400 + _sat_normalized_score(percent) * 1200
    return int(round(score / 10.0) * 10)


def sat_section_score_from_percent(percent: float) -> int:
    """Bosqichli SAT (Reading&Writing yoki Math, har biri alohida) uchun —
    xuddi shu egri chiziq shaklidan foydalanib, bitta bo'lim ballini rasmiy
    SAT bo'lim shkalasi 200-800 oralig'iga o'tkazadi. Ikkala bo'lim balli
    qo'shilsa, rasmiy SAT kabi 400-1600 kompozit hosil bo'ladi."""
    score = 200 + _sat_normalized_score(percent) * 600
    return int(round(score / 10.0) * 10)


def grade_value_for_exam_type(exam_type: str, percent: float) -> str | None:
    et = (exam_type or "simple")
    if et == "cefr":
        return cefr_level_from_percent(percent)
    if et == "ielts":
        return ielts_band_from_percent(percent)
    if et == "sat":
        return str(sat_score_from_percent(percent))
    return None


# ===================== REQUIRED CHANNELS (FORCED SUBSCRIPTION) =====================
def _normalize_channel_ref(text: str) -> str:
    s = (text or "").strip()
    if s.startswith("@"):
        s = s[1:]
    return s


def _is_valid_join_url(url: str) -> bool:
    """Telegram inline tugmasi qabul qiladigan havola. Avval faqat "http" bilan
    boshlanishi tekshirilardi — "http://" yoki ichida bo'sh joy bo'lgan havola
    "A'zo bo'ling" tugmasini Telegram rad etishiga va a'zo bo'lmagan HAR BIR
    foydalanuvchi "Kutilmagan xatolik" olishiga olib kelardi."""
    return bool(re.fullmatch(r"https?://[A-Za-z0-9.-]+\.[A-Za-z]{2,}(/[^\s<>\"']*)?", url or ""))


def _looks_like_int(s: str) -> bool:
    try:
        int(s)
        return True
    except Exception:
        return False


async def get_required_channels(active_only: bool = True):
    q = "SELECT id, chat_id, username, title, join_url, is_active FROM required_channels"
    if active_only:
        q += " WHERE is_active=1"
    q += " ORDER BY id ASC"
    return await POOL.fetch(q)


def build_join_kb(channels_rows):
    kb = InlineKeyboardMarkup(inline_keyboard=[])
    for row in channels_rows:
        chat_id, username, title, join_url = row["chat_id"], row["username"], row["title"], row["join_url"]
        url = (join_url or "").strip()
        if not url:
            if username:
                url = f"https://t.me/{username}"
            else:
                url = ""
        btn_text = f"➕ {title or (('@' + username) if username else str(chat_id))}"
        if url:
            kb.inline_keyboard.append([InlineKeyboardButton(text=btn_text, url=url)])
    kb.inline_keyboard.append([InlineKeyboardButton(text="✅ A’zo bo‘ldim — Tekshirish", callback_data="chk_sub")])
    return kb


# ✅ Obunani tekshirish keshi.
#
# MUAMMO EDI: avvalgi versiyada har bir xabar/tugma bosilganda barcha majburiy
# kanallar uchun Telegram'ning get_chat_member so'rovi YANGIDAN yuborilardi.
# Faol foydalanuvchilar ko'p bo'lganda bu Telegram'ning so'rov chegarasiga
# (flood limit) tez urilib qolardi — natijada Telegram vaqtincha xato qaytarardi
# va bot buni "a'zo emas" deb noto'g'ri talqin qilib, HAQIQATDA A'ZO bo'lgan
# odamlarni ham bloklab qo'yardi. Endi natija bir necha daqiqa keshlanadi, shu
# bilan ortiqcha so'rovlar oldini olinadi va vaqtinchalik Telegram xatolarida
# oldingi (muvaffaqiyatli) holat saqlab qolinadi.
SUBSCRIPTION_CACHE_TTL_OK = 600   # a'zo deb tasdiqlangan holat 10 daqiqa ishonchli hisoblanadi
SUBSCRIPTION_CACHE_TTL_FAIL = 20  # a'zo emas/aniqlanmagan holat tezroq qayta tekshiriladi
_subscription_cache: dict[int, tuple[bool, float]] = {}


async def check_user_subscribed(bot: Bot, user_id: int, force: bool = False) -> tuple[bool, list]:
    rows = await get_required_channels(active_only=True)
    if not rows:
        _subscription_cache.pop(user_id, None)
        return True, []

    if not force:
        cached = _subscription_cache.get(user_id)
        if cached is not None:
            ok_cached, checked_at = cached
            ttl = SUBSCRIPTION_CACHE_TTL_OK if ok_cached else SUBSCRIPTION_CACHE_TTL_FAIL
            if (time.time() - checked_at) < ttl and ok_cached:
                return True, []

    not_joined = []
    had_infra_error = False

    for row in rows:
        chat_id, username, title = row["chat_id"], row["username"], row["title"]
        _id = row["id"]

        resolved_chat_id = chat_id
        if not resolved_chat_id and username:
            try:
                ch = await bot.get_chat(f"@{username}")
                resolved_chat_id = ch.id
                await POOL.execute("UPDATE required_channels SET chat_id=$1 WHERE id=$2", int(resolved_chat_id), _id)
            except Exception as e:
                logging.error(f"Kanal topilmadi: @{username} (id={_id}). Xato: {e}")
                resolved_chat_id = None

        if not resolved_chat_id:
            not_joined.append(row)
            had_infra_error = True
            continue

        try:
            member = await bot.get_chat_member(chat_id=int(resolved_chat_id), user_id=user_id)
            if member.status in ("left", "kicked"):
                not_joined.append(row)
        except TelegramForbiddenError:
            # Bot bu kanal/guruhga a'zo emas yoki admin huquqi yo'q — bu sozlash
            # xatosi, foydalanuvchining aybi emas. Adminlar panelida ko'rinadi.
            logging.error(
                f"Majburiy kanal id={_id} ('{title}'): bot u yerda ADMIN emas, "
                f"a'zolikni tekshira olmayapti."
            )
            not_joined.append(row)
            had_infra_error = True
        except TelegramRetryAfter as e:
            logging.warning(f"Telegram so'rov chegarasi (flood): {e.retry_after}s. Keshdan foydalanamiz.")
            not_joined.append(row)
            had_infra_error = True
        except Exception as e:
            logging.error(f"A'zolikni tekshirishda xatolik (kanal id={_id}): {e}")
            not_joined.append(row)
            had_infra_error = True

    ok = (len(not_joined) == 0)

    if not ok and had_infra_error:
        # Vaqtinchalik Telegram/konfiguratsiya xatosi tufayli avval tasdiqlangan
        # foydalanuvchini bloklamaymiz — oxirgi ma'lum holatini saqlab qolamiz.
        cached = _subscription_cache.get(user_id)
        if cached and cached[0] and (time.time() - cached[1]) < SUBSCRIPTION_CACHE_TTL_OK:
            return True, []

    _subscription_cache[user_id] = (ok, time.time())
    return ok, not_joined


async def ensure_subscribed_message(message: Message, bot: Bot) -> bool:
    if await is_admin_user(message.from_user.id):
        return True
    ok, not_joined = await check_user_subscribed(bot, message.from_user.id)
    if ok:
        return True
    kb = build_join_kb(not_joined)
    await message.answer(
        "🔒 <b>Botdan foydalanish uchun quyidagi kanal/guruhlarga a'zo bo'ling.</b>\n\n"
        "A'zo bo'lgach, pastdagi <b>«✅ A'zo bo'ldim — Tekshirish»</b> tugmasini bosing.",
        parse_mode=ParseMode.HTML,
        reply_markup=kb
    )
    return False


async def ensure_subscribed_callback(callback: CallbackQuery, bot: Bot) -> bool:
    if await is_admin_user(callback.from_user.id):
        return True
    ok, not_joined = await check_user_subscribed(bot, callback.from_user.id)
    if ok:
        return True
    kb = build_join_kb(not_joined)
    await callback.message.answer(
        "🔒 <b>Davom etish uchun quyidagi kanal/guruhlarga a'zo bo'ling.</b>\n\n"
        "A'zo bo'lgach, pastdagi <b>«✅ A'zo bo'ldim — Tekshirish»</b> tugmasini bosing.",
        parse_mode=ParseMode.HTML,
        reply_markup=kb
    )
    return False


class SubscriptionMessageMiddleware(BaseMiddleware):
    async def __call__(self, handler, event: Message, data):
        if not event.from_user:
            return await handler(event, data)

        bot: Bot = data.get("bot")
        if not bot:
            return await handler(event, data)

        if await is_admin_user(event.from_user.id):
            return await handler(event, data)

        ok = await ensure_subscribed_message(event, bot)
        if not ok:
            return
        return await handler(event, data)


class SubscriptionCallbackMiddleware(BaseMiddleware):
    async def __call__(self, handler, event: CallbackQuery, data):
        if not event.from_user:
            return await handler(event, data)

        if (event.data or "") == "chk_sub":
            return await handler(event, data)

        bot: Bot = data.get("bot")
        if not bot:
            return await handler(event, data)

        if await is_admin_user(event.from_user.id):
            return await handler(event, data)

        ok = await ensure_subscribed_callback(event, bot)
        if not ok:
            return
        return await handler(event, data)


# ===================== DATABASE INIT (PostgreSQL) =====================
async def _ensure_column(db, table: str, column: str, col_def: str):
    await db.execute(f"ALTER TABLE {table} ADD COLUMN IF NOT EXISTS {column} {col_def}")


async def init_db(db):
    await db.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id SERIAL PRIMARY KEY,
            telegram_id BIGINT UNIQUE,
            full_name TEXT,
            phone TEXT,
            role TEXT DEFAULT 'student'
        )
    """)

    await db.execute("""
        CREATE TABLE IF NOT EXISTS tests (
            id SERIAL PRIMARY KEY,
            code TEXT UNIQUE,
            title TEXT,
            is_free INTEGER DEFAULT 0,
            price INTEGER DEFAULT 0,
            duration INTEGER DEFAULT 0,
            questions_count INTEGER DEFAULT 30,
            file_id TEXT,
            answers TEXT,

            start_mode TEXT DEFAULT 'normal',
            start_ts BIGINT DEFAULT 0,
            start_at TEXT,

            rush_finalized INTEGER DEFAULT 0,
            rush_finalized_ts BIGINT DEFAULT 0
        )
    """)

    await db.execute("""
        CREATE TABLE IF NOT EXISTS payments (
            id SERIAL PRIMARY KEY,
            user_id INTEGER REFERENCES users(id),
            test_id INTEGER REFERENCES tests(id),
            status TEXT DEFAULT 'pending',
            screenshot_id TEXT,
            created_at TEXT,
            reject_reason TEXT,
            started_at TEXT,
            started_ts BIGINT
        )
    """)

    await db.execute("""
        CREATE TABLE IF NOT EXISTS results (
            id SERIAL PRIMARY KEY,
            user_id INTEGER REFERENCES users(id),
            test_id INTEGER,
            score INTEGER,
            total_questions INTEGER,
            percent DOUBLE PRECISION,
            date TEXT
        )
    """)

    await db.execute("""
        CREATE TABLE IF NOT EXISTS sessions (
            id SERIAL PRIMARY KEY,
            user_id INTEGER,
            test_id INTEGER,
            mode TEXT,
            start_ts BIGINT,
            duration_sec INTEGER,
            created_at TEXT,
            UNIQUE(user_id, test_id, mode)
        )
    """)

    # ✅ Bosqichli imtihonlar (DTM 5-fan / IELTS 4-ko'nikma) uchun. Har bir
    # bosqich (fan yoki ko'nikma) sozlamasi tests.skills_cfg (JSON) da,
    # o'quvchining topshirgan javoblari/natijalari shu jadvalda saqlanadi.
    await db.execute("""
        CREATE TABLE IF NOT EXISTS staged_attempts (
            id SERIAL PRIMARY KEY,
            user_id INTEGER REFERENCES users(id),
            test_id INTEGER REFERENCES tests(id),
            status TEXT DEFAULT 'in_progress',
            section_idx INTEGER DEFAULT 0,
            sections_json TEXT,
            overall_score DOUBLE PRECISION,
            overall_grade TEXT,
            graded_by BIGINT,
            graded_at BIGINT,
            started_at BIGINT,
            finished_at BIGINT,
            created_ts BIGINT
        )
    """)
    # ✅ Avval bitta (user, test) juftligiga faqat bitta urinishga ruxsat
    # berilardi. Endi qayta yechish imkoni bo'lgani uchun (eski natija
    # yo'qolmasdan) bu cheklovni olib tashlaymiz — attempt_no orqali
    # tartiblanadi.
    await db.execute("ALTER TABLE staged_attempts DROP CONSTRAINT IF EXISTS staged_attempts_user_id_test_id_key")
    await _ensure_column(db, "staged_attempts", "attempt_no", "INTEGER DEFAULT 1")

    # ✅ FSM holati (kim, qaysi qadamda, qanday ma'lumot bilan) — bot qayta
    # ishga tushganda ham foydalanuvchi test yechish/test yaratish jarayonini
    # yo'qotib qo'ymasligi uchun.
    await db.execute("""
        CREATE TABLE IF NOT EXISTS fsm_storage (
            storage_key TEXT PRIMARY KEY,
            state TEXT,
            data TEXT,
            updated_ts BIGINT DEFAULT 0
        )
    """)

    await db.execute("""
        CREATE TABLE IF NOT EXISTS user_limits (
            user_tg_id BIGINT PRIMARY KEY,
            strikes INTEGER DEFAULT 0,
            blocked_until_ts BIGINT DEFAULT 0,
            last_violation_ts BIGINT DEFAULT 0
        )
    """)

    await db.execute("""
        CREATE TABLE IF NOT EXISTS required_channels (
            id SERIAL PRIMARY KEY,
            chat_id BIGINT UNIQUE,
            username TEXT,
            title TEXT,
            join_url TEXT,
            is_active INTEGER DEFAULT 1
        )
    """)

    # ✅ Creator ID for tests (BIGINT: bu Telegram user ID, INTEGER ga sig'maydi)
    await _ensure_column(db, "tests", "creator_id", "BIGINT")
    await _ensure_column(db, "tests", "created_ts", "BIGINT DEFAULT 0")

    await _ensure_column(db, "tests", "rush_finalized", "INTEGER DEFAULT 0")
    await _ensure_column(db, "tests", "rush_finalized_ts", "BIGINT DEFAULT 0")
    await _ensure_column(db, "tests", "exam_type", "TEXT DEFAULT 'simple'")
    await _ensure_column(db, "tests", "dtm_cfg", "TEXT")

    # ✅ "Maxsus" rejimi butunlay olib tashlandi (Rasch endi guruh
    # kattaligidan qat'i nazar bir xil ishlaydi, alohida modul shart emas) —
    # eski, endi ishlatilmaydigan ustunlarni tozalaymiz.
    await db.execute("ALTER TABLE tests DROP COLUMN IF EXISTS maxsus_finalized")
    await db.execute("ALTER TABLE tests DROP COLUMN IF EXISTS maxsus_finalized_ts")

    await _ensure_column(db, "results", "earned_points", "DOUBLE PRECISION")
    await _ensure_column(db, "results", "max_points", "DOUBLE PRECISION")
    await _ensure_column(db, "results", "details", "TEXT")
    await _ensure_column(db, "results", "user_answers", "TEXT")
    await _ensure_column(db, "results", "open_text_answers", "TEXT")  # variantsiz savollarga matn javoblar (JSON)
    await _ensure_column(db, "results", "grade", "TEXT")
    await _ensure_column(db, "results", "finished_ts", "BIGINT DEFAULT 0")

    await _ensure_column(db, "tests", "is_free", "INTEGER DEFAULT 0")
    await _ensure_column(db, "tests", "code", "TEXT")
    await _ensure_column(db, "tests", "questions_count", "INTEGER DEFAULT 30")

    await _ensure_column(db, "results", "rasch_ball", "DOUBLE PRECISION")
    await _ensure_column(db, "results", "rasch_percent", "DOUBLE PRECISION")

    await _ensure_column(db, "tests", "start_mode", "TEXT DEFAULT 'normal'")
    await _ensure_column(db, "tests", "start_ts", "BIGINT DEFAULT 0")
    await _ensure_column(db, "tests", "start_at", "TEXT")
    await _ensure_column(db, "tests", "subject", "TEXT")

    await _ensure_column(db, "results", "certificate_sent", "INTEGER DEFAULT 0")
    await _ensure_column(db, "results", "certificate_file_id", "TEXT")
    await _ensure_column(db, "results", "certificate_sent_ts", "BIGINT DEFAULT 0")

    await db.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_tests_code_unique ON tests(code)")

    # ✅ Pending to'lovlar poyga holatini (race condition) oldini olish uchun:
    # bir user bitta testga faqat bitta "pending" to'lovga ega bo'la oladi.
    await db.execute("""
        CREATE UNIQUE INDEX IF NOT EXISTS idx_payments_pending_unique
        ON payments(user_id, test_id) WHERE status='pending'
    """)

    # ✅ Bitta foydalanuvchi bitta testni faqat 1 marta topshira oladi — bu
    # ilgari faqat ilova darajasida (SELECT keyin INSERT) tekshirilardi, ya'ni
    # ikkita so'rov bir vaqtda kelsa (masalan tugma ikki marta bosilsa) ikkita
    # natija qatori yozilib qolishi mumkin edi. Endi DB darajasida ham kafolatlanadi.
    await db.execute("""
        CREATE UNIQUE INDEX IF NOT EXISTS idx_results_user_test_unique
        ON results(user_id, test_id)
    """)

    # ✅ Admin panelidagi yangi imkoniyatlar uchun:
    await _ensure_column(db, "users", "created_ts", "BIGINT DEFAULT 0")  # ro'yxatdan o'tgan vaqt

    # ✅ Profildagi ixtiyoriy ma'lumotlar — hech biri majburiy emas, faqat
    # xohlaganlar to'ldiradi (masalan o'qituvchi o'quvchilarini yaxshiroq
    # tanishi uchun).
    await _ensure_column(db, "users", "region", "TEXT")
    await _ensure_column(db, "users", "age", "INTEGER")
    await _ensure_column(db, "users", "school", "TEXT")
    await _ensure_column(db, "users", "grade", "TEXT")

    # ✅ Har bir ADMIN o'zining shaxsiy to'lov kartasini biriktirishi mumkin —
    # shu admin yaratgan pullik testni sotib olmoqchi bo'lgan o'quvchiga
    # aynan shu karta ko'rsatiladi va to'lov screenshoti to'g'ridan-to'g'ri
    # o'sha adminga (asosiy ADMIN_ID'ga emas) boradi. Oddiy foydalanuvchi
    # (admin bo'lmagan) yaratgan pullik testlarda esa umumiy standart karta
    # ishlatiladi — bu holda asosiy admin to'lovni qabul qilib, tegishli
    # ulushni test yaratuvchisiga o'zi o'tkazib berishi kerak bo'ladi.
    await _ensure_column(db, "users", "payment_card", "TEXT")
    await _ensure_column(db, "users", "payment_owner", "TEXT")
    await _ensure_column(db, "users", "payment_phone", "TEXT")
    await _ensure_column(db, "tests", "is_hidden", "INTEGER DEFAULT 0")  # vaqtincha yashirilgan test
    await _ensure_column(db, "tests", "is_closed", "INTEGER DEFAULT 0")  # admin/creator tomonidan yakunlangan test

    # ✅ Har bir test o'z sertifikat sozlamalariga ega bo'lishi mumkin (bo'sh
    # bo'lsa standart qiymatlar ishlatiladi — pastda make_certificate_pdf'ga
    # uzatilganda hal qilinadi).
    await _ensure_column(db, "tests", "cert_academy_name", "TEXT")
    await _ensure_column(db, "tests", "cert_signee_name", "TEXT")
    await _ensure_column(db, "tests", "cert_template", "TEXT DEFAULT 'gold'")

    # ✅ Bosqichli imtihon (DTM 5-fan / IELTS 4-ko'nikma) sozlamalari — JSON:
    # {"kind": "dtm_staged"|"ielts4", "order": [...], "sections": {key: {...}}}
    await _ensure_column(db, "tests", "skills_cfg", "TEXT")

    # ✅ Javob kalitida "x" (variantsiz/ochiq savol) bo'lgan pozitsiyalar
    # uchun test yaratuvchi kiritgan KUTILGAN/TO'G'RI javob — JSON:
    # {"6": "Parij", "9": "42"} (1-asosli savol raqami -> matn). Ustozga
    # o'quvchi javobini tekshirishda tayanch sifatida ko'rsatiladi.
    await _ensure_column(db, "tests", "open_ref_answers", "TEXT")

    # ✅ Ochiq (x) savollarni ustoz tekshirishi: har bir natija uchun
    # {"savol_raqami": 1|0} (1 = to'g'ri) va hali tekshirilmagan javoblar soni.
    await _ensure_column(db, "results", "open_grades", "TEXT")
    await _ensure_column(db, "results", "open_pending", "INTEGER DEFAULT 0")
    await _ensure_column(db, "tests", "open_grade_reminded_ts", "BIGINT DEFAULT 0")

    # ✅ O'quvchi testdan menyuga chiqib ketsa, belgilagan javoblari shu yerda
    # saqlanadi — testga qaytganda (vaqt qolgan bo'lsa) davom ettiradi.
    await _ensure_column(db, "sessions", "saved_answers", "TEXT")

    # ✅ Qayta yechish ruxsati: ustoz "qayta yechishni taklif qilish"ni
    # bosganda har bir o'quvchiga BITTA ruxsat beriladi. Avval taklif xabari
    # (yoki retakeok_ tugmasi) cheksiz ishlatilib, "1 urinish" qoidasini
    # istalgancha aylanib o'tish mumkin edi.
    await db.execute("""
        CREATE TABLE IF NOT EXISTS retake_grants (
            user_id INTEGER,
            test_id INTEGER,
            granted_ts BIGINT DEFAULT 0,
            used_ts BIGINT DEFAULT 0,
            PRIMARY KEY (user_id, test_id)
        )
    """)

    # ✅ To'lov summasi va tasdiqlangan vaqti alohida saqlanadi — moliyaviy
    # hisobot testning HOZIRGI narxiga emas, haqiqatda to'langan summaga va
    # tasdiqlangan kunga qarab hisoblanadi.
    await _ensure_column(db, "payments", "amount", "INTEGER")
    await _ensure_column(db, "payments", "approved_at", "TEXT")

    # ✅ Tezlik uchun indekslar — faqat qidiruvni tezlashtiradi, hech qanday
    # ma'lumotni o'zgartirmaydi yoki o'chirmaydi.
    await db.execute("CREATE INDEX IF NOT EXISTS idx_results_test_id ON results(test_id)")
    await db.execute("CREATE INDEX IF NOT EXISTS idx_payments_user_test ON payments(user_id, test_id)")
    await db.execute("CREATE INDEX IF NOT EXISTS idx_staged_attempts_user_test ON staged_attempts(user_id, test_id)")


# ===================== SECURITY =====================
class SecurityManager:
    def __init__(self):
        self.msg_hits: dict[int, deque] = defaultdict(deque)
        self.cb_hits: dict[int, deque] = defaultdict(deque)
        self.last_cb: dict[int, tuple[str, float]] = {}

    async def _get_limit_row(self, user_id: int):
        row = await POOL.fetchrow(
            "SELECT strikes, blocked_until_ts, last_violation_ts FROM user_limits WHERE user_tg_id=$1",
            user_id,
        )
        if not row:
            await POOL.execute(
                "INSERT INTO user_limits (user_tg_id, strikes, blocked_until_ts, last_violation_ts) "
                "VALUES ($1, 0, 0, 0) ON CONFLICT (user_tg_id) DO NOTHING",
                user_id,
            )
            return (0, 0, 0)
        return (row["strikes"], row["blocked_until_ts"], row["last_violation_ts"])

    async def is_blocked(self, user_id: int) -> tuple[bool, int]:
        _, blocked_until_ts, _ = await self._get_limit_row(user_id)
        now_ = int(time.time())
        if blocked_until_ts and blocked_until_ts > now_:
            return True, blocked_until_ts - now_
        return False, 0

    async def _apply_violation(self, user_id: int):
        now_ = int(time.time())
        strikes, _, last_violation_ts = await self._get_limit_row(user_id)

        if last_violation_ts and (now_ - last_violation_ts) > STRIKE_RESET_AFTER_SEC:
            strikes = 0

        strikes += 1
        step_index = min(strikes - 1, len(MUTE_STEPS) - 1)
        mute_sec = MUTE_STEPS[step_index]
        new_blocked_until = now_ + mute_sec

        await POOL.execute(
            "UPDATE user_limits SET strikes=$1, blocked_until_ts=$2, last_violation_ts=$3 WHERE user_tg_id=$4",
            strikes, new_blocked_until, now_, user_id,
        )

        return strikes, mute_sec

    def _trim_deque(self, dq: deque, window_sec: int, now_ts_: float):
        while dq and (now_ts_ - dq[0]) > window_sec:
            dq.popleft()

    async def check_message(self, user_id: int) -> tuple[bool, str]:
        blocked, remain = await self.is_blocked(user_id)
        if blocked:
            return False, f"⛔️ Siz vaqtincha bloklandingiz.\n⏳ Qolgan vaqt: {fmt_seconds(remain)}"

        now_ts_ = time.time()
        dq = self.msg_hits[user_id]
        dq.append(now_ts_)
        self._trim_deque(dq, MSG_WINDOW_SEC, now_ts_)

        if len(dq) > MSG_LIMIT:
            strikes, mute_sec = await self._apply_violation(user_id)
            return False, (
                "🚫 Juda ko‘p so‘rov yuboryapsiz.\n"
                f"🔒 Vaqtincha blok: {fmt_seconds(mute_sec)}\n"
                f"⚠️ Ogohlantirishlar: {strikes}"
            )
        return True, ""

    async def check_callback(self, user_id: int, data: str) -> tuple[bool, str, bool]:
        blocked, remain = await self.is_blocked(user_id)
        if blocked:
            return False, f"⛔️ Siz vaqtincha bloklandingiz.\n⏳ Qolgan vaqt: {fmt_seconds(remain)}", False

        now_ts_ = time.time()
        last = self.last_cb.get(user_id)
        if last and last[0] == data and (now_ts_ - last[1]) < SAME_CALLBACK_COOLDOWN_SEC:
            return False, "⏳ Sekinroq bosing (2 soniya).", True

        self.last_cb[user_id] = (data, now_ts_)

        dq = self.cb_hits[user_id]
        dq.append(now_ts_)
        self._trim_deque(dq, CB_WINDOW_SEC, now_ts_)

        if len(dq) > CB_LIMIT:
            strikes, mute_sec = await self._apply_violation(user_id)
            return False, (
                "🚫 Tugmalarni juda tez bosyapsiz.\n"
                f"🔒 Vaqtincha blok: {fmt_seconds(mute_sec)}\n"
                f"⚠️ Ogohlantirishlar: {strikes}"
            ), False

        return True, "", False

    async def check_exam_callback(self, user_id: int) -> tuple[bool, str]:
        """Test yechish tugmalari (A-F, Oldingi/Keyingi, Yakunlash) uchun
        alohida, ancha yumshoq chegara. Oddiy chegara (15 soniyada 15 bosish
        va bir xil tugmaga 2 soniyalik pauza) imtihon paytida tez ishlayotgan
        o'quvchini BLOKLAB qo'yardi (har savolga "javob + Keyingi" = 2 bosish),
        uning javoblari yo'qolar va taymer esa to'xtamasdan ketaverardi."""
        blocked, remain = await self.is_blocked(user_id)
        if blocked:
            return False, f"⛔️ Siz vaqtincha bloklandingiz.\n⏳ Qolgan vaqt: {fmt_seconds(remain)}"
        now_ts_ = time.time()
        dq = self.exam_cb_hits[user_id]
        dq.append(now_ts_)
        self._trim_deque(dq, EXAM_CB_WINDOW_SEC, now_ts_)
        if len(dq) > EXAM_CB_LIMIT:
            strikes, mute_sec = await self._apply_violation(user_id)
            return False, (
                "🚫 Tugmalarni haddan tashqari tez bosyapsiz.\n"
                f"🔒 Vaqtincha blok: {fmt_seconds(mute_sec)}\n"
                f"⚠️ Ogohlantirishlar: {strikes}"
            )
        return True, ""


# Imtihon tugmalari: insoniy tezlikda (sekundiga 1-3 bosish) hech qachon
# chegaraga yetmaydi, faqat avtomatik skript/bot bunchalik tez bosa oladi.
EXAM_CB_PREFIXES = ("t_", "sq_")
EXAM_CB_WINDOW_SEC = 15
EXAM_CB_LIMIT = 90

SECURITY = SecurityManager()
SECURITY.exam_cb_hits = defaultdict(deque)


class MessageSecurityMiddleware(BaseMiddleware):
    async def __call__(self, handler, event: Message, data):
        user_id = event.from_user.id if event.from_user else 0
        if not user_id:
            return await handler(event, data)
        # ✅ Adminlar spam-himoyadan ozod — bu himoya o'quvchilarning suiiste'molini
        # cheklash uchun, admin panelidagi tez-tez tugma bosishlar tufayli
        # adminning o'zini bloklab qo'yishi noto'g'ri bo'lardi.
        if await is_admin_user(user_id):
            return await handler(event, data)
        allowed, msg = await SECURITY.check_message(user_id)
        if not allowed:
            try:
                await event.answer(msg)
            except Exception:
                pass
            return
        return await handler(event, data)


class CallbackSecurityMiddleware(BaseMiddleware):
    async def __call__(self, handler, event: CallbackQuery, data):
        user_id = event.from_user.id if event.from_user else 0
        if not user_id:
            return await handler(event, data)
        if await is_admin_user(user_id):
            return await handler(event, data)
        if (event.data or "").startswith(EXAM_CB_PREFIXES):
            allowed, msg = await SECURITY.check_exam_callback(user_id)
            if not allowed:
                try:
                    await event.answer(msg, show_alert=True)
                except Exception:
                    pass
                return
            return await handler(event, data)
        allowed, msg, silent = await SECURITY.check_callback(user_id, event.data or "")
        if not allowed:
            try:
                await event.answer(msg, show_alert=not silent)
            except Exception:
                pass
            return
        return await handler(event, data)


# ===================== KEYBOARDS =====================
def main_menu_kb(is_admin=False):
    buttons = [
        [KeyboardButton(text=BTN_FREE), KeyboardButton(text=BTN_PAID)],
        [KeyboardButton(text=BTN_MY), KeyboardButton(text=BTN_RATING)],
        [KeyboardButton(text=BTN_RESULTS), KeyboardButton(text=BTN_HELP)],
        [KeyboardButton(text=BTN_ADD_TEST), KeyboardButton(text=BTN_MY_CREATED_TESTS)],
        [KeyboardButton(text=BTN_PROFILE)]
    ]
    if is_admin:
        buttons.append([KeyboardButton(text=BTN_ADMIN_PENDING), KeyboardButton(text=BTN_ADMIN_STATS)])
        buttons.append([KeyboardButton(text=BTN_ADMIN_CHANNELS), KeyboardButton(text=BTN_ADMIN_DEL_TEST)])
        buttons.append([KeyboardButton(text=BTN_ADMIN_ADD_ADMIN)])
    return ReplyKeyboardMarkup(keyboard=buttons, resize_keyboard=True)


def phone_kb():
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="📞 Kontaktni ulashish", request_contact=True)],
            [KeyboardButton(text=BTN_CANCEL), KeyboardButton(text=BTN_HOME)],
        ],
        resize_keyboard=True
    )


def nav_kb(include_back=True):
    row = [KeyboardButton(text=BTN_HOME)]
    if include_back:
        row.insert(0, KeyboardButton(text=BTN_BACK))
    return ReplyKeyboardMarkup(keyboard=[row, [KeyboardButton(text=BTN_CANCEL)]], resize_keyboard=True)


def search_kb(scope: str):
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="🔎 Kod orqali topish", callback_data=f"askcode_{scope}")]
    ])


# ===================== COUNTDOWN =====================
ACTIVE_TIMERS: dict[tuple[int, int], asyncio.Task] = {}


async def start_countdown(bot: Bot, chat_id: int, test_id: int, end_ts: int, msg_id: int):
    key = (chat_id, test_id)
    try:
        while True:
            remain = end_ts - now_ts()
            if remain <= 0:
                try:
                    await bot.edit_message_text(
                        chat_id=chat_id,
                        message_id=msg_id,
                        text="⏱ <b>Vaqt tugadi!</b>",
                        parse_mode=ParseMode.HTML
                    )
                except Exception:
                    pass
                # o'zimizni ro'yxatdan chiqaramiz — aks holda topshirish jarayonidagi
                # cancel_timer() aynan shu vazifani bekor qilib, uni yarim yo'lda to'xtatardi
                if ACTIVE_TIMERS.get(key) is asyncio.current_task():
                    ACTIVE_TIMERS.pop(key, None)
                try:
                    await _auto_submit_on_timeout(bot, chat_id, test_id, end_ts)
                except Exception:
                    logging.exception(f"Vaqt tugaganda avtomatik topshirib bo'lmadi (user={chat_id}, test={test_id})")
                break

            text = f"⏳ <b>Qolgan vaqt:</b> {fmt_seconds(remain)}"
            try:
                await bot.edit_message_text(
                    chat_id=chat_id,
                    message_id=msg_id,
                    text=text,
                    parse_mode=ParseMode.HTML
                )
            except Exception:
                pass

            await asyncio.sleep(min(COUNTDOWN_TICK_SEC, max(1, end_ts - now_ts())))
    finally:
        # ✅ Faqat O'ZIMIZNI o'chiramiz. Avval har doim o'chirilardi: eski taymer
        # bekor qilinib, o'rniga yangisi (masalan keyingi bosqich taymeri)
        # yozilgach, eskisining finally bloki YANGISINI ro'yxatdan o'chirib
        # yuborardi — yangi taymerni keyin hech kim to'xtata olmasdi.
        if ACTIVE_TIMERS.get(key) is asyncio.current_task():
            ACTIVE_TIMERS.pop(key, None)


def cancel_timer(user_tg_id: int, test_id: int):
    key = (user_tg_id, test_id)
    task = ACTIVE_TIMERS.pop(key, None)
    if task and not task.done():
        task.cancel()


async def _resume_active_timers(bot: Bot):
    """
    ACTIVE_TIMERS xotirada saqlanadi — bot qayta ishga tushsa (masalan
    xatolik yoki server qayta yuklanishi tufayli), eski hisoblagich
    xabarlari "muzlab qoladi" (endi yangilanmaydi), garchi imtihonning
    haqiqiy tugash vaqti bazada saqlanib, to'g'ri hisoblanaversa ham.

    Bu funksiya bot ishga tushganda hali vaqti tugamagan sessiyalarni topib,
    har biriga YANGI hisoblagich xabari yuboradi — shunda foydalanuvchi
    "vaqtim yo'qolibdimi" deb xavotirlanmaydi va testni qayta ochib
    (hisoblagich qayta boshlanmasdan) davom ettirishi mumkinligini biladi.
    """
    now_ = now_ts()
    try:
        rows = await POOL.fetch("""
            SELECT s.user_id, s.test_id, s.start_ts, s.duration_sec, u.telegram_id
            FROM sessions s
            JOIN users u ON u.id = s.user_id
            WHERE s.start_ts + s.duration_sec > $1
        """, now_)
    except Exception:
        logging.exception("_resume_active_timers: sessiyalarni o'qib bo'lmadi")
        return

    for r in rows:
        end_ts_ = int(r["start_ts"]) + int(r["duration_sec"])
        tg_id = r["telegram_id"]
        test_id = r["test_id"]
        try:
            timer_msg = await bot.send_message(
                tg_id,
                "🔄 <b>Bot qayta ishga tushdi</b> — xavotir olmang, vaqtingiz yo'qolmagan.\n\n"
                "⏳ Qolgan vaqt hisoblanmoqda...\n"
                "Davom etish uchun shu testni qayta oching (tugma orqali «Boshlash»ni bosing) — "
                "hisoblagich qaytadan boshlanmaydi, qolgan vaqtdan davom etadi.",
                parse_mode=ParseMode.HTML
            )
            cancel_timer(tg_id, test_id)
            ACTIVE_TIMERS[(tg_id, test_id)] = asyncio.create_task(
                start_countdown(bot, tg_id, test_id, end_ts_, timer_msg.message_id)
            )
        except Exception:
            logging.warning(f"Timer tiklanmadi: user={tg_id} test={test_id}")


FSM_STORAGE = None  # main() ichida o'rnatiladi (taymer avtomatik topshirishi uchun)
EVENTS_ISOLATION = None  # main() ichida o'rnatiladi (foydalanuvchi bo'yicha qulf)


class _DetachedState:
    """process_test_submission uchun minimal FSMContext o'rnini bosuvchi —
    foydalanuvchining hozirgi (boshqa) holatini buzmasdan, saqlangan
    javoblarni topshirish uchun."""

    def __init__(self, data: dict):
        self._data = dict(data)

    async def get_data(self):
        return dict(self._data)

    async def clear(self):
        self._data = {}

    async def set_state(self, *_a, **_k):
        pass

    async def update_data(self, **kw):
        self._data.update(kw)


async def _auto_submit_on_timeout(bot: "Bot", user_tg_id: int, test_id: int, end_ts: int) -> None:
    # ✅ Taymer foydalanuvchining o'z tugmalari bilan BIR NAVBATDA ishlaydi —
    # aks holda oxirgi soniyada "Yakunlash" bosilib, taymer ham bir vaqtda
    # ishga tushsa, bosqichli testda bitta bosqich ikki marta topshirilib,
    # keyingi bosqich javobsiz yakunlanib ketishi mumkin edi.
    if EVENTS_ISOLATION is None:
        return await _auto_submit_on_timeout_unlocked(bot, user_tg_id, test_id, end_ts)
    key = StorageKey(bot_id=bot.id, chat_id=user_tg_id, user_id=user_tg_id)
    async with EVENTS_ISOLATION.lock(key):
        await _auto_submit_on_timeout_unlocked(bot, user_tg_id, test_id, end_ts)


async def _auto_submit_on_timeout_unlocked(bot: "Bot", user_tg_id: int, test_id: int, end_ts: int) -> None:
    """✅ Vaqt tugaganda o'quvchi belgilagan javoblarni AVTOMATIK topshiradi.
    Avval vaqt tugasa, "Yakunlash"ni o'zi 60 soniya ichida bosmagan o'quvchining
    butun ishi yo'qolardi ("Vaqt tugadi! Javoblar qabul qilinmadi"). Bosqichli
    imtihonlarda esa vaqt umuman majburiy emas edi — taymer faqat ko'rsatardi."""
    if FSM_STORAGE is None:
        return
    user_db_id = await get_user_db_id(user_tg_id)
    if not user_db_id:
        return
    key = StorageKey(bot_id=bot.id, chat_id=user_tg_id, user_id=user_tg_id)
    state = FSMContext(storage=FSM_STORAGE, key=key)
    st = await state.get_state()
    data = await state.get_data()

    def _same_session(sess) -> bool:
        return bool(sess) and abs(int(sess["start_ts"]) + int(sess["duration_sec"]) - int(end_ts)) <= 5

    # 1) oddiy test — o'quvchi hali test oynasida
    if st in (TestProcessState.solving.state, TestProcessState.awaiting_free_text.state) \
            and int(data.get("test_id") or 0) == test_id:
        mode = data.get("mode") or "free"
        if not _same_session(await get_session(user_db_id, test_id, mode)):
            return
        qcount = int(data.get("qcount") or 0)
        answers = data.get("test_answers", {}) or {}
        user_ans = "".join(answers.get(str(i), "*") for i in range(1, qcount + 1))
        await bot.send_message(user_tg_id, "⏱ <b>Vaqt tugadi</b> — belgilagan javoblaringiz avtomatik topshirildi.",
                               parse_mode=ParseMode.HTML)
        await process_test_submission(user_tg_id, bot, state, user_ans, data.get("test_open_answers") or {})
        return

    # 2) bosqichli test — joriy bosqich
    if st in (StagedSolvingState.in_section.state, StagedSolvingState.awaiting_free_text.state) \
            and data.get("staged_test_id") == test_id:
        skey = data.get("stage_current_key")
        if not skey or not _same_session(await get_session(user_db_id, test_id, f"stage_{skey}")):
            return
        await state.set_state(StagedSolvingState.in_section)
        stype = data.get("stage_current_type")
        await bot.send_message(user_tg_id, "⏱ <b>Bosqich vaqti tugadi</b> — javoblaringiz avtomatik topshirildi.",
                               parse_mode=ParseMode.HTML)
        if stype in (STAGE_TYPE_MCQ, STAGE_TYPE_MCQ_BAND):
            qcount = int(data.get("stage_current_qcount") or 0)
            answers = data.get("stage_mcq_answers", {}) or {}
            user_ans = "".join(answers.get(str(i), "*") for i in range(1, qcount + 1))
            await _complete_current_stage_mcq(user_tg_id, bot, state, user_ans)
        else:
            await delete_session(user_db_id, test_id, f"stage_{skey}")
            result = {"text": None, "photo_id": None, "voice_id": None, "band": None, "feedback": None}
            await _save_stage_result_and_advance(user_tg_id, bot, state, skey, result)
        return

    # 3) o'quvchi testdan chiqib ketgan — saqlangan javoblar bilan topshiramiz
    for mode in ("free", "paid"):
        sess = await POOL.fetchrow(
            "SELECT start_ts, duration_sec, saved_answers FROM sessions WHERE user_id=$1 AND test_id=$2 AND mode=$3",
            user_db_id, test_id, mode)
        if not _same_session(sess):
            continue
        saved = _load_json_dict(sess["saved_answers"])
        qcount = int(await POOL.fetchval("SELECT questions_count FROM tests WHERE id=$1", test_id) or 0)
        answers = saved.get("answers") or {}
        user_ans = "".join(answers.get(str(i), "*") for i in range(1, qcount + 1))
        tmp = _DetachedState({"test_id": test_id, "mode": mode})
        await bot.send_message(user_tg_id, "⏱ <b>Vaqt tugadi</b> — belgilagan javoblaringiz avtomatik topshirildi.",
                               parse_mode=ParseMode.HTML)
        await process_test_submission(user_tg_id, bot, tmp, user_ans, saved.get("open_answers") or {})
        return


async def _stash_test_progress(state: FSMContext, user_tg_id: int) -> None:
    """O'quvchi test oynasidan menyuga chiqib ketayotgan bo'lsa — belgilagan
    javoblarini sessiyaga saqlab qo'yadi (holat tozalanishidan oldin)."""
    try:
        st = await state.get_state()
        if st not in (TestProcessState.solving.state, TestProcessState.awaiting_free_text.state):
            return
        data = await state.get_data()
        test_id = int(data.get("test_id") or 0)
        mode = data.get("mode") or "free"
        user_db_id = await get_user_db_id(user_tg_id)
        if not test_id or not user_db_id:
            return
        payload = {"answers": data.get("test_answers") or {}, "open_answers": data.get("test_open_answers") or {},
                   "current_q": data.get("current_q", 1)}
        await POOL.execute("UPDATE sessions SET saved_answers=$1 WHERE user_id=$2 AND test_id=$3 AND mode=$4",
                           json.dumps(payload, ensure_ascii=False), user_db_id, test_id, mode)
    except Exception:
        logging.exception("Test javoblarini saqlab qo'yib bo'lmadi")


async def _restore_test_progress(user_db_id: int, test_id: int, mode: str) -> dict:
    raw = await POOL.fetchval(
        "SELECT saved_answers FROM sessions WHERE user_id=$1 AND test_id=$2 AND mode=$3", user_db_id, test_id, mode)
    return _load_json_dict(raw)


async def upsert_session(user_id: int, test_id: int, mode: str, start_ts_: int, duration_sec: int):
    await POOL.execute("""
        INSERT INTO sessions (user_id, test_id, mode, start_ts, duration_sec, created_at)
        VALUES ($1, $2, $3, $4, $5, $6)
        ON CONFLICT(user_id, test_id, mode) DO UPDATE SET
            start_ts=excluded.start_ts,
            duration_sec=excluded.duration_sec,
            created_at=excluded.created_at
    """, user_id, test_id, mode, start_ts_, duration_sec, now_str_local())


async def get_session(user_id: int, test_id: int, mode: str):
    return await POOL.fetchrow("""
        SELECT start_ts, duration_sec
        FROM sessions
        WHERE user_id=$1 AND test_id=$2 AND mode=$3
    """, user_id, test_id, mode)


async def delete_session(user_id: int, test_id: int, mode: str):
    await POOL.execute("DELETE FROM sessions WHERE user_id=$1 AND test_id=$2 AND mode=$3", user_id, test_id, mode)


# ===================== ROUTER + MIDDLEWARE =====================
class CallbackAnswerDedupMiddleware(BaseRequestMiddleware):
    """Telegram bitta tugma bosilishiga (callback query) faqat BIR MARTA javob
    qabul qiladi — ikkinchi answerCallbackQuery "query is too old ... or query
    ID is invalid" xatosi bilan qaytadi. Ko'p handlerlar o'zi callback.answer()
    chaqiradi, keyin CallbackAnswerMiddleware ham avtomatik yana javob berardi
    — natijada xato global ushlagichga tushib, foydalanuvchiga har safar
    "⚠️ Kutilmagan xatolik" xabari chiqardi. Bu so'rov darajasidagi
    middleware har bir query uchun faqat birinchi javobni yuboradi."""

    def __init__(self):
        self._answered: dict[str, float] = {}

    async def __call__(self, make_request, bot, method):
        if isinstance(method, AnswerCallbackQuery):
            qid = method.callback_query_id
            now_ = time.time()
            if len(self._answered) > 20000:
                cutoff = now_ - 600
                self._answered = {k: v for k, v in self._answered.items() if v > cutoff}
            if qid in self._answered:
                return True
            self._answered[qid] = now_
            try:
                return await make_request(bot, method)
            except TelegramBadRequest as e:
                if "query is too old" in str(e) or "query ID is invalid" in str(e):
                    return True
                raise
        return await make_request(bot, method)


def _install_bot_session_middlewares(bot: Bot) -> None:
    bot.session.middleware(CallbackAnswerDedupMiddleware())


router = Router()

# ✅ Ko'p callback handlerlar javobni tugatgach `callback.answer()` chaqirishni
# unutgan edi — natijada Telegram tugmada "yuklanmoqda" aylanmasini bir muncha
# vaqt ko'rsatib turadi (foydalanuvchiga "osilib qoldi" bo'lib tuyuladi).
# Bu middleware har bir callback handler tugagach (agar hali javob berilmagan
# bo'lsa) avtomatik `callback.answer()` chaqiradi. Eng birinchi bo'lib
# ro'yxatdan o'tkazilgan — shunda quyidagi middlewarelar (masalan obuna
# tekshiruvi) erta to'xtab qolsa ham, tugma baribir javob oladi.
router.callback_query.middleware(CallbackAnswerMiddleware())

router.message.middleware(MessageSecurityMiddleware())
router.callback_query.middleware(CallbackSecurityMiddleware())

# ✅ Forced subscription as GLOBAL gate
router.message.middleware(SubscriptionMessageMiddleware())
router.callback_query.middleware(SubscriptionCallbackMiddleware())


# ===================== REQUIRED SUB CHECK BUTTON =====================
@router.callback_query(F.data == "chk_sub")
async def check_subscribed_cb(callback: CallbackQuery, bot: Bot, state: FSMContext):
    # ✅ force=True: bu yerda foydalanuvchi aynan "Tekshirish" tugmasini bosgan,
    # shuning uchun keshni chetlab o'tib, Telegram'dan darhol yangi holatni so'raymiz.
    ok, not_joined = await check_user_subscribed(bot, callback.from_user.id, force=True)
    if not ok:
        kb = build_join_kb(not_joined)
        await callback.answer("Hali barcha kanal/guruhlarga a'zo bo'lmagansiz.", show_alert=True)
        await callback.message.answer(
            "⚠️ <b>Hali a'zo bo'lmagan kanal/guruhlar bor:</b>\n\n"
            "Har biriga a'zo bo'lgach, «Tekshirish» tugmasini qayta bosing.\n"
            "<i>Agar a'zo bo'lsangiz-u baribir shu xabar chiqsa, bir necha soniyadan so'ng qayta urinib ko'ring — "
            "Telegram tomonidan vaqtinchalik tekshiruv kechikayotgan bo'lishi mumkin.</i>",
            parse_mode=ParseMode.HTML,
            reply_markup=kb
        )
        return

    await callback.answer("✅ Tasdiqlandi, xush kelibsiz!", show_alert=False)
    await state.clear()
    admin_flag = await is_admin_user(callback.from_user.id)
    await callback.message.answer("✅ Kirish ruxsat berildi.", reply_markup=main_menu_kb(admin_flag))


# ===================== GLOBAL NAVIGATION (STATE RESET) =====================
async def go_home(message: Message, state: FSMContext):
    await _stash_test_progress(state, message.from_user.id)
    await state.clear()
    admin_flag = await is_admin_user(message.from_user.id)
    await message.answer("🏠 Bosh menyu", reply_markup=main_menu_kb(admin_flag))


@router.message(StateFilter("*"), F.text == BTN_HOME)
async def nav_home(message: Message, state: FSMContext):
    await go_home(message, state)


# ===================== ORQAGA QAYTISH TARIXI (test yaratish sehrgarlari) =====================
# ✅ Test yaratish (oddiy ham, bosqichli ham) ko'p qadamli — ba'zan 15-20
# qadamgacha. Har bir "oldinga" o'tishda o'sha qadam (node) shu stekka
# qo'shiladi; "🔄 Boshidan boshlash" tugmasi endi bu stekni tekshiradi —
# agar tarix bo'sh bo'lmasa, BUTUN jarayonni bekor qilish o'rniga oxirgi
# qadamga qaytaradi (foydalanuvchi bitta xatoni tuzatish uchun hammasini
# qayta boshlashga majbur bo'lmaydi). Tarix bo'sh bo'lsa (birinchi qadamda
# ekan) — odatdagidek bosh menyuga qaytadi.
async def _nav_push(state: FSMContext, w: str, node: str, idx: int | None = None) -> None:
    data = await state.get_data()
    stack = list(data.get("nav_stack") or [])
    stack.append({"w": w, "node": node, "idx": idx})
    await state.update_data(nav_stack=stack)


async def _nav_pop(state: FSMContext) -> dict | None:
    data = await state.get_data()
    stack = list(data.get("nav_stack") or [])
    if not stack:
        return None
    entry = stack.pop()
    await state.update_data(nav_stack=stack)
    return entry


@router.message(StateFilter("*"), F.text == BTN_BACK)
async def nav_back(message: Message, state: FSMContext):
    entry = await _nav_pop(state)
    if entry:
        if entry["w"] == "base":
            handler = BASE_NAV_HANDLERS.get(entry["node"])
            if handler:
                await handler(message, state)
                return
        elif entry["w"] == "staged":
            ok = await _staged_nav_replay(message, state, entry)
            if ok:
                return
    await go_home(message, state)


@router.message(StateFilter("*"), F.text.in_(SECTION_BTNS))
async def nav_any_section(message: Message, state: FSMContext, bot: Bot):
    if not await ensure_subscribed_message(message, bot):
        return

    await _stash_test_progress(state, message.from_user.id)
    await state.clear()

    if message.text == BTN_FREE:
        return await show_free_tests(message)
    if message.text == BTN_PAID:
        return await show_paid_tests(message)
    if message.text == BTN_MY:
        return await my_tests(message)
    if message.text == BTN_RATING:
        return await rating_menu(message)
    if message.text == BTN_RESULTS:
        return await my_results(message)
    if message.text == BTN_HELP:
        return await help_menu(message)
    if message.text == BTN_ADD_TEST:
        return await admin_add_test(message, state)
    if message.text == BTN_MY_CREATED_TESTS:
        return await list_my_created_tests(message)
    if message.text == BTN_ADMIN_PENDING:
        return await admin_pending(message)
    if message.text == BTN_ADMIN_STATS:
        return await admin_stats(message)
    if message.text == BTN_ADMIN_CHANNELS:
        return await admin_channels_panel(message, state, bot)
    if message.text == BTN_ADMIN_DEL_TEST:
        return await admin_ask_del_test(message, state)
    if message.text == BTN_ADMIN_ADD_ADMIN:
        return await admin_ask_manage_admin(message, state)
    if message.text == BTN_PROFILE:
        return await profile_menu(message, state)


@router.message(StateFilter("*"), F.text == BTN_CANCEL)
async def cancel_action(message: Message, state: FSMContext):
    await _stash_test_progress(state, message.from_user.id)
    await state.clear()
    admin_flag = await is_admin_user(message.from_user.id)
    await message.answer("Bekor qilindi.", reply_markup=main_menu_kb(admin_flag))


# ===================== /start + REGISTRATION =====================
@router.message(CommandStart())
async def command_start(message: Message, state: FSMContext, bot: Bot):
    if not await ensure_subscribed_message(message, bot):
        return

    user = await POOL.fetchrow("SELECT * FROM users WHERE telegram_id=$1", message.from_user.id)

    # ✅ Foydalanuvchi bazada bo'lishi mumkin, lekin hali ro'yxatdan o'tmagan
    # (masalan asosiy admin uni /addadmin orqali oldindan admin qilib qo'ygan —
    # u holda faqat telegram_id va rol yozilgan, ism/telefon yo'q). Avval
    # bunday odamga "Xush kelibsiz, None!" deyilib, ro'yxatdan o'tish butunlay
    # o'tkazib yuborilardi.
    if user and (user["full_name"] or "").strip():
        admin_flag = await is_admin_user(message.from_user.id)
        await message.answer(f"👋 Xush kelibsiz, <b>{_h(user['full_name'])}</b>!", parse_mode=ParseMode.HTML,
                             reply_markup=main_menu_kb(admin_flag))
        return

    await message.answer(
        f"👋 Assalomu alaykum!\n\n"
        f"<b>{ACADEMY_NAME}</b> test tizimiga xush kelibsiz. Foydalanishni boshlash uchun avval qisqacha "
        f"ro'yxatdan o'tamiz.\n\n"
        f"Ism va familiyangizni kiriting:",
        parse_mode=ParseMode.HTML,
        reply_markup=nav_kb(include_back=False)
    )
    await state.set_state(RegistrationState.full_name)


@router.message(RegistrationState.full_name)
async def process_name(message: Message, state: FSMContext):
    if (message.text or "") in TOP_MENU_BTNS:
        return
    full_name = sanitize_plain_text(message.text, max_len=100)
    if len(full_name) < 3:
        return await message.answer("⚠️ Ism-familiya juda qisqa yoki noto'g'ri belgilar bor. Qayta kiriting:")
    await state.update_data(full_name=full_name)
    await message.answer(
        "📱 Endi telefon raqamingizni yuboring.\n"
        "Pastdagi <b>«📞 Kontaktni ulashish»</b> tugmasini bosing (yoki qo'lda yozib yuboring):",
        parse_mode=ParseMode.HTML,
        reply_markup=phone_kb())
    await state.set_state(RegistrationState.phone)


@router.message(RegistrationState.phone)
async def process_phone(message: Message, state: FSMContext):
    if (message.text or "") in TOP_MENU_BTNS:
        return
    if message.contact:
        contact = message.contact.phone_number
    else:
        contact = normalize_phone_input(message.text)
        if not contact:
            return await message.answer(
                "⚠️ Telefon raqami noto'g'ri. «📞 Kontaktni ulashish» tugmasini bosing yoki raqamni "
                "to'liq yozing (masalan: <code>+998901234567</code>):",
                parse_mode=ParseMode.HTML, reply_markup=phone_kb())
    data = await state.get_data()
    role = 'admin' if message.from_user.id == ADMIN_ID else 'student'

    # ✅ Qator oldindan mavjud bo'lsa (masalan admin qilib qo'yilgan, lekin
    # ro'yxatdan o'tmagan) — ism/telefonni YANGILAYMIZ, rolini esa saqlaymiz.
    # Avval "DO NOTHING" bo'lgani uchun bunday odamning ismi umuman saqlanmasdi.
    await POOL.execute(
        "INSERT INTO users (telegram_id, full_name, phone, role, created_ts) VALUES ($1, $2, $3, $4, $5) "
        "ON CONFLICT (telegram_id) DO UPDATE SET full_name=EXCLUDED.full_name, phone=EXCLUDED.phone, "
        "created_ts=CASE WHEN COALESCE(users.created_ts,0)=0 THEN EXCLUDED.created_ts ELSE users.created_ts END",
        message.from_user.id, data.get("full_name", ""), contact, role, now_ts()
    )

    admin_flag = await is_admin_user(message.from_user.id)
    await message.answer(
        "✅ <b>Ro'yxatdan muvaffaqiyatli o'tdingiz!</b>\n\n"
        "Endi pastdagi menyudan tekin yoki pullik testlarni tanlab, ishlashni boshlashingiz mumkin.\n\n"
        f"ℹ️ Savol yoki takliflar bo'lsa: {PAYMENT_PHONE}",
        parse_mode=ParseMode.HTML,
        reply_markup=main_menu_kb(admin_flag))
    await state.clear()


# ===================== SEARCH BY CODE =====================
@router.callback_query(F.data.startswith("askcode_"))
async def ask_code(callback: CallbackQuery, state: FSMContext, bot: Bot):
    if not await ensure_subscribed_callback(callback, bot):
        return

    scope = callback.data.split("_", 1)[1]  # paid/free/my
    await state.update_data(scope=scope)
    await callback.message.answer(
        "🔎 <b>Test kodini kiriting</b>\nMasalan: <code>DTM24A</code> yoki <code>MS-001</code>",
        parse_mode=ParseMode.HTML,
        reply_markup=nav_kb(include_back=True)
    )
    await state.set_state(SearchByCodeState.waiting_code)


@router.message(SearchByCodeState.waiting_code)
async def handle_code_search(message: Message, state: FSMContext):
    if (message.text or "") in TOP_MENU_BTNS:
        return

    user_db_id = await get_user_db_id(message.from_user.id)
    if not user_db_id:
        await message.answer("Avval /start orqali ro‘yxatdan o‘ting.")
        await state.clear()
        return

    st = await state.get_data()
    scope = st.get("scope", "paid")

    code = normalize_code(message.text)
    if not is_valid_code(code):
        await message.answer("⚠️ Kod formati noto‘g‘ri. Qayta kiriting:")
        return

    t = await test_by_code(code)
    if not t:
        await message.answer("❌ Bunday kodli test topilmadi. Qayta urinib ko‘ring.")
        return

    (test_id, code, title, is_free, price, duration, qcount, file_id, answers,
     start_mode, start_ts_val, start_at) = (
        t["id"], t["code"], t["title"], t["is_free"], t["price"], t["duration"], t["questions_count"],
        t["file_id"], t["answers"], t["start_mode"], t["start_ts"], t["start_at"]
    )

    scheduled_blocked, left = is_scheduled_not_started(start_mode, start_ts_val)

    status = None
    payment_id = None
    started_at_paid = None
    reject_reason = None

    prow = await POOL.fetchrow("""
        SELECT p.id, p.status, p.started_at, p.reject_reason
        FROM payments p
        WHERE p.user_id=$1 AND p.test_id=$2
        ORDER BY p.id DESC LIMIT 1
    """, user_db_id, test_id)
    if prow:
        payment_id, status, started_at_paid, reject_reason = prow["id"], prow["status"], prow["started_at"], prow["reject_reason"]

    has_result = bool(await POOL.fetchrow(
        "SELECT 1 FROM results WHERE user_id=$1 AND test_id=$2 LIMIT 1", user_db_id, test_id
    ))

    # ✅ Ijtimoiy dalil: nechta kishi yechgan, o'rtacha natija qancha
    social = await POOL.fetchrow(
        "SELECT COUNT(*) AS cnt, AVG(percent) AS avgp FROM results WHERE test_id=$1", test_id
    )

    kind = "🆓 Tekin" if is_free == 1 else "🧾 Pullik"
    text = (
        f"✅ <b>Test topildi</b>\n\n"
        f"🔑 Kod: <code>{code}</code>\n"
        f"📚 Nomi: <b>{title}</b>\n"
        f"🏷 Turi: <b>{kind}</b>\n"
        f"🧮 Savollar: <b>{qcount}</b> ta\n"
        f"⏳ Vaqt: <b>{duration} daqiqa</b>\n"
    )
    if is_free == 0:
        text += f"💰 Narx: <b>{price} so'm</b>\n"
    if social and social["cnt"]:
        text += f"👥 {social['cnt']} kishi yechgan, o'rtacha natija {float(social['avgp'] or 0):.1f}%\n"

    if (start_mode or "normal") == "scheduled":
        text += f"📅 Boshlanish: <b>{start_at or 'Noma’lum'}</b>\n"
        if scheduled_blocked:
            text += f"⏳ Qolgan: <b>{fmt_seconds(left)}</b>\n"

    if is_free == 0:
        if status == "approved":
            text += "✅ Holat: Sizda bor\n"
        elif status == "pending":
            text += "🟡 Holat: To‘lov kutilmoqda\n"
        elif status == "rejected":
            text += "❌ Holat: Rad etilgan\n"
            if reject_reason:
                text += f"📝 Sabab: <i>{_h(reject_reason)}</i>\n"
        else:
            text += "ℹ️ Holat: Sotib olinmagan\n"
    else:
        if has_result:
            text += "⚠️ Siz bu tekin testni allaqachon topshirgansiz (1-urinish).\n"

    kb = InlineKeyboardMarkup(inline_keyboard=[])

    if scope == "free":
        if is_free == 1:
            if has_result:
                kb.inline_keyboard.append(
                    [InlineKeyboardButton(text="📄 Natijani ko‘rish", callback_data=f"viewres_{test_id}")])
                kb.inline_keyboard.append(
                    [InlineKeyboardButton(text="🏆 Reytingni ko‘rish", callback_data="rt_overall")])
            else:
                kb.inline_keyboard.append([InlineKeyboardButton(text="🚀 Boshlash", callback_data=f"fstart_{test_id}")])
        else:
            text += "\n⚠️ Bu test pullik. '🧾 Pullik testlar' bo‘limidan sotib oling."
    elif scope == "paid":
        if is_free == 0:
            if status == "approved":
                kb.inline_keyboard.append(
                    [InlineKeyboardButton(text="📦 Pullik testlarim", callback_data="open_mytests")])
            elif status == "pending":
                kb.inline_keyboard.append(
                    [InlineKeyboardButton(text="🟡 Kutilmoqda", callback_data=f"payinfo_{test_id}")])
            elif status == "rejected":
                kb.inline_keyboard.append(
                    [InlineKeyboardButton(text="♻️ Qayta to‘lov", callback_data=f"rebuy_{test_id}")])
            else:
                kb.inline_keyboard.append([InlineKeyboardButton(text="Sotib olish", callback_data=f"buy_{test_id}")])
        else:
            text += "\n⚠️ Bu test tekin. '🆓 Tekin testlar' bo‘limidan boshlang."
    else:  # my
        if is_free == 0:
            if status == "approved":
                if has_result:
                    kb.inline_keyboard.append(
                        [InlineKeyboardButton(text="📄 Natijani ko‘rish", callback_data=f"viewres_{test_id}")])
                else:
                    kb.inline_keyboard.append(
                        [InlineKeyboardButton(text="🚀 Testni boshlash", callback_data=f"begin_{payment_id}")])
                if started_at_paid:
                    kb.inline_keyboard.append(
                        [InlineKeyboardButton(text="📄 PDF-ni ko‘rish", callback_data=f"pdf_{test_id}")])
            elif status == "pending":
                kb.inline_keyboard.append([InlineKeyboardButton(text="🟡 Pending", callback_data=f"payinfo_{test_id}")])
            elif status == "rejected":
                kb.inline_keyboard.append(
                    [InlineKeyboardButton(text="♻️ Qayta to‘lov yuborish", callback_data=f"rebuy_{test_id}")])
            else:
                kb.inline_keyboard.append([InlineKeyboardButton(text="Sotib olish", callback_data=f"buy_{test_id}")])
        else:
            text += "\nℹ️ Bu test tekin — pastdagi «🆓 Tekin testlar» bo‘limidan boshlang."

    # ✅ Mini App sozlangan bo'lsa — testni interaktiv oynada ochish tugmasi
    if WEBAPP_URL:
        kb.inline_keyboard.append([InlineKeyboardButton(
            text="📱 Mini App'da ochish", web_app=WebAppInfo(url=f"{WEBAPP_URL}/webapp?code={code}"))])

    await message.answer(text, parse_mode=ParseMode.HTML, reply_markup=kb if kb.inline_keyboard else None)
    await state.clear()


@router.callback_query(F.data == "noop")
async def noop(callback: CallbackQuery):
    await callback.answer("OK", show_alert=False)


@router.callback_query(F.data == "dismiss_msg")
async def dismiss_msg(callback: CallbackQuery):
    await callback.answer("Bekor qilindi.")
    try:
        await callback.message.edit_text("❌ Bekor qilindi.")
    except Exception:
        try:
            await callback.message.delete()
        except Exception:
            pass


# ===================== INTERACTIVE TEST UI (A, B, C, D, E, F) =====================
def _is_open_question(correct_key: str, current_q: int) -> bool:
    """Javob kalitida shu savol pozitsiyasida "x" (variantsiz/ochiq savol)
    turganini tekshiradi."""
    idx = current_q - 1
    return bool(correct_key) and 0 <= idx < len(correct_key) and correct_key[idx] == "x"


def question_status_note(correct_key: str, current_q: int) -> str:
    """Savol yozuvi ustiga qo'shiladigan ogohlantirish — ochiq savol bo'lsa
    o'quvchiga variant emas, matn kutilayotgani darhol aniq bo'lsin."""
    if _is_open_question(correct_key, current_q):
        return "❗️ <b>Bu savolga variant emas — matn bilan javob yozishingiz kerak.</b>\n"
    return ""


def test_ui_kb(current_q: int, qcount: int, answers_dict: dict, correct_key: str = ""):
    kb = InlineKeyboardMarkup(inline_keyboard=[])
    current_ans = answers_dict.get(str(current_q))

    if _is_open_question(correct_key, current_q):
        # ✅ Bu savol javob kalitida "x" — variantsiz (ochiq) deb belgilangan.
        # A-F tugmalarini ko'rsatish mantiqsiz (variant tanlash mumkin emas)
        # va chalkashtiradi, shuning uchun FAQAT matn yozish tugmasi chiqadi.
        mark = "✅ " if current_ans == "x" else "✍️ "
        kb.inline_keyboard.append(
            [InlineKeyboardButton(text=f"{mark}Matn bilan javob yozish", callback_data="t_freetext")])
    else:
        row1 = []
        for letter in ['A', 'B', 'C']:
            mark = "🟢 " if current_ans == letter.lower() else ""
            row1.append(InlineKeyboardButton(text=f"{mark}{letter}", callback_data=f"t_ans_{letter.lower()}"))
        kb.inline_keyboard.append(row1)

        row2 = []
        for letter in ['D', 'E', 'F']:
            mark = "🟢 " if current_ans == letter.lower() else ""
            row2.append(InlineKeyboardButton(text=f"{mark}{letter}", callback_data=f"t_ans_{letter.lower()}"))
        kb.inline_keyboard.append(row2)

    nav = []
    if current_q > 1:
        nav.append(InlineKeyboardButton(text="⬅️ Oldingi", callback_data="t_nav_prev"))

    nav.append(InlineKeyboardButton(text=f"🔄 {current_q}/{qcount}", callback_data="t_nav_noop"))

    if current_q < qcount:
        nav.append(InlineKeyboardButton(text="Keyingi ➡️", callback_data="t_nav_next"))

    kb.inline_keyboard.append(nav)

    answered_count = len(answers_dict)
    kb.inline_keyboard.append(
        [InlineKeyboardButton(text=f"✅ Yakunlash ({answered_count}/{qcount})", callback_data="t_submit")])
    return kb


# ===================== FREE TESTS =====================
# Foydalanuvchiga ko'rinadigan ro'yxatlar sahifalab chiqariladi — testlar soni
# 50+ ga yetganda bitta ulkan, chalkash xabar+tugmalar to'plami chiqmasligi uchun.
USER_LIST_PAGE_SIZE = 8


def _user_pagination_nav_row(list_key: str, page: int, has_more: bool):
    row = []
    if page > 0:
        row.append(InlineKeyboardButton(text="⬅️ Oldingi", callback_data=f"ulp_{list_key}_{page - 1}"))
    if has_more:
        row.append(InlineKeyboardButton(text="Keyingi ➡️", callback_data=f"ulp_{list_key}_{page + 1}"))
    return row


@router.callback_query(F.data.startswith("ulp_"))
async def user_list_page_nav(callback: CallbackQuery):
    _, list_key, page_s = callback.data.split("_", 2)
    page = int(page_s)
    renderer = _USER_PAGINATED_RENDERERS.get(list_key)
    if renderer:
        await renderer(callback.message, callback.from_user.id, page)


async def _render_free_tests(answerable, user_tg_id: int, page: int = 0):
    user_db_id = await get_user_db_id(user_tg_id)
    if not user_db_id:
        return await answerable.answer("Avval /start orqali ro‘yxatdan o‘ting.")

    offset = max(0, page) * USER_LIST_PAGE_SIZE
    tests = await POOL.fetch("""
        SELECT id, code, title, duration, questions_count, start_mode, start_ts, start_at
        FROM tests
        WHERE is_free=1 AND COALESCE(is_hidden, 0)=0
        ORDER BY id DESC
        LIMIT $1 OFFSET $2
    """, USER_LIST_PAGE_SIZE + 1, offset)
    has_more = len(tests) > USER_LIST_PAGE_SIZE
    tests = tests[:USER_LIST_PAGE_SIZE]

    if not tests:
        if page == 0:
            return await answerable.answer("Hozircha tekin testlar yo‘q.", reply_markup=search_kb("free"))
        return await answerable.answer("Boshqa tekin test yo'q.")

    text = f"<b>🆓 Tekin testlar</b> ({page + 1}-sahifa)\n\n"
    kb = InlineKeyboardMarkup(inline_keyboard=[])
    kb.inline_keyboard.append([InlineKeyboardButton(text="🔎 Kod orqali topish", callback_data="askcode_free")])

    for row in tests:
        tid, code, title, duration, qcount = row["id"], row["code"], row["title"], row["duration"], row["questions_count"]
        start_mode, start_ts_val, start_at = row["start_mode"], row["start_ts"], row["start_at"]
        sched = (start_mode or "normal") == "scheduled"
        blocked, left = is_scheduled_not_started(start_mode, start_ts_val)

        line = f"🔑 <code>{code}</code> — <b>{title}</b> | 🧮 {qcount} ta | ⏳ {duration} daqiqa"
        if sched:
            line += f"\n   📅 Boshlanish: <b>{start_at}</b>"
            if blocked:
                line += f" | ⏳ {fmt_seconds(left)}"
        text += line + "\n\n"

        kb.inline_keyboard.append([InlineKeyboardButton(text=f"🚀 Boshlash: {code}", callback_data=f"fstart_{tid}")])

    nav = _user_pagination_nav_row("free", page, has_more)
    if nav:
        kb.inline_keyboard.append(nav)

    await answerable.answer(text, parse_mode=ParseMode.HTML, reply_markup=kb)


@router.message(F.text == BTN_FREE)
async def show_free_tests(message: Message):
    await _render_free_tests(message, message.from_user.id, 0)


@router.callback_query(F.data.startswith("fstart_"))
async def free_test_start(callback: CallbackQuery, state: FSMContext, bot: Bot):
    if not await ensure_subscribed_callback(callback, bot):
        return

    test_id = int(callback.data.split("_")[1])

    user_db_id = await get_user_db_id(callback.from_user.id)
    if not user_db_id:
        return await callback.message.answer("Avval /start orqali ro‘yxatdan o‘ting.")

    # ✅ Bosqichli testlar (DTM 5-fan / IELTS 4-ko'nikma) butunlay boshqa
    # oqimga ega (bitta fayl/javob kaliti emas, ketma-ket bosqichlar) —
    # shuning uchun ularni darhol o'z motoriga yo'naltiramiz.
    et_row = await POOL.fetchrow("SELECT exam_type FROM tests WHERE id=$1 AND is_free=1", test_id)
    if et_row and (et_row["exam_type"] or "") in STAGED_EXAM_TYPES:
        return await start_staged_attempt(callback.from_user.id, test_id, callback.message, bot, state)

    if await POOL.fetchrow("SELECT 1 FROM results WHERE user_id=$1 AND test_id=$2 LIMIT 1", user_db_id, test_id):
        return await callback.message.answer(
            "⚠️ Siz bu testni allaqachon topshirgansiz. Qayta topshirib bo‘lmaydi.")

    row = await POOL.fetchrow("""
        SELECT code, title, file_id, duration, questions_count, answers, start_mode, start_ts, start_at,
               is_hidden, is_closed
        FROM tests
        WHERE id=$1 AND is_free=1
    """, test_id)

    if not row:
        return await callback.message.answer("Test topilmadi (yoki tekin emas).")

    if int(row["is_hidden"] or 0) == 1:
        return await callback.message.answer("⏸ Bu test admin tomonidan vaqtincha to'xtatilgan. Keyinroq urinib ko'ring.")
    if int(row["is_closed"] or 0) == 1:
        return await callback.message.answer("🛑 Bu test yakunlangan, endi yangi urinish qabul qilinmaydi.")

    code, title, file_id, duration_min, qcount, answers, start_mode, start_ts_val, start_at = (
        row["code"], row["title"], row["file_id"], row["duration"], row["questions_count"],
        row["answers"], row["start_mode"], row["start_ts"], row["start_at"]
    )

    blocked, left = is_scheduled_not_started(start_mode, start_ts_val)
    if blocked:
        kb = InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text="🔁 Qayta tekshirish", callback_data=f"chkfree_{test_id}")],
        ])
        return await callback.message.answer(
            f"⏳ <b>Test hali boshlanmagan</b>\n\n"
            f"🔑 Kod: <code>{code}</code>\n"
            f"📅 Boshlanish: <b>{start_at}</b>\n"
            f"⏳ Qolgan: <b>{fmt_seconds(left)}</b>\n\n"
            f"Iltimos, kuting.",
            parse_mode=ParseMode.HTML,
            reply_markup=kb
        )

    answer_key = normalize_answers(answers)
    if len(answer_key) != int(qcount):
        return await callback.message.answer(
            "⚠️ Admin testni xato qo‘shgan: javob kaliti uzunligi savol soniga teng emas.")

    cancel_timer(callback.from_user.id, test_id)

    # ✅ Agar foydalanuvchida bu test uchun sessiya allaqachon mavjud bo'lsa
    # (masalan "Ortga" bosgan yoki bot qayta ishga tushgan bo'lsa), hisoblagichni
    # QAYTA BOSHLAMAYMIZ — aks holda "Boshlash"ni qayta-qayta bosib vaqtni
    # cheksiz uzaytirib olish mumkin bo'lardi. Buning o'rniga eski vaqt
    # o'lchagichdan davom ettiramiz.
    existing_session = await get_session(user_db_id, test_id, "free")
    if existing_session:
        start_ts_ = int(existing_session["start_ts"])
        duration_sec = int(existing_session["duration_sec"])
    else:
        start_ts_ = int(start_ts_val) if (start_mode or "normal") == "scheduled" and int(
            start_ts_val or 0) > 0 else now_ts()
        duration_sec = int(duration_min) * 60
    end_ts_ = start_ts_ + duration_sec

    if now_ts() > end_ts_ + 60:
        await delete_session(user_db_id, test_id, "free")
        return await callback.message.answer("❌ Bu test vaqti tugab bo‘lgan.")

    await upsert_session(user_db_id, test_id, "free", start_ts_, duration_sec)

    await bot.send_document(
        chat_id=callback.from_user.id,
        document=file_id,
        caption=(
            f"📄 <b>{title}</b>\n"
            f"🔑 Kod: <code>{code}</code>\n"
            f"🧮 Savollar: <b>{qcount}</b>\n\n"
            f"🆓 Tekin test. Vaqt ketdi ✅"
        ),
        parse_mode=ParseMode.HTML
    )

    timer_msg = await bot.send_message(
        chat_id=callback.from_user.id,
        text="⏳ <b>Qolgan vaqt:</b> ...",
        parse_mode=ParseMode.HTML
    )

    ACTIVE_TIMERS[(callback.from_user.id, test_id)] = asyncio.create_task(
        start_countdown(bot, callback.from_user.id, test_id, end_ts_, timer_msg.message_id)
    )

    # ✅ O'quvchi avval shu testni boshlab, menyuga chiqib ketgan bo'lsa —
    # belgilagan javoblari tiklanadi (yordam bo'limida va'da qilinganidek).
    # ✅ O'quvchi test yechayotganda eski xabardagi "Boshlash"ni qayta bossa,
    # joriy (hali saqlanmagan) javoblari avval sessiyaga yoziladi — aks holda
    # holat ustidan yozilib, belgilagan javoblari o'chib ketardi.
    await _stash_test_progress(state, callback.from_user.id)
    saved = await _restore_test_progress(user_db_id, test_id, "free")
    saved_answers = saved.get("answers") or {}
    await state.update_data(
        test_id=test_id, mode="free", qcount=qcount, current_q=1, test_answers=saved_answers,
        test_open_answers=saved.get("open_answers") or {}, test_answers_key=answer_key
    )
    resumed_note = (f"♻️ Avvalgi <b>{len(saved_answers)}</b> ta javobingiz tiklandi.\n"
                    if saved_answers else "")
    await bot.send_message(
        chat_id=callback.from_user.id,
        text=(
            f"📝 <b>Test boshlandi!</b>\n\n"
            f"{resumed_note}"
            f"{question_status_note(answer_key, 1)}"
            f"Savol: <b>1 / {qcount}</b>\n"
            f"Pastdagi tugmalar orqali javoblarni interaktiv belgilang.\n"
            f"<i>(Yoki xohlasangiz odatdagidek abcd... formatida birdaniga yozib yuboring)</i>"
        ),
        parse_mode=ParseMode.HTML,
        reply_markup=test_ui_kb(1, int(qcount), saved_answers, answer_key)
    )
    await state.set_state(TestProcessState.solving)


@router.callback_query(F.data.startswith("chkfree_"))
async def check_free_start(callback: CallbackQuery):
    test_id = int(callback.data.split("_")[1])
    row = await POOL.fetchrow("SELECT code, start_mode, start_ts, start_at FROM tests WHERE id=$1", test_id)
    if not row:
        return await callback.message.answer("Test topilmadi.")
    code, start_mode, start_ts_val, start_at = row["code"], row["start_mode"], row["start_ts"], row["start_at"]
    blocked, left = is_scheduled_not_started(start_mode, start_ts_val)
    if blocked:
        return await callback.message.answer(
            f"⏳ Hali erta.\n🔑 <code>{code}</code>\n📅 {start_at}\n⏳ Qolgan: {fmt_seconds(left)}",
            parse_mode=ParseMode.HTML
        )
    await callback.message.answer("✅ Boshlash vaqti keldi! Endi 'Boshlash' tugmasini bosing.")


# ===================== PAID TESTS =====================
async def _render_paid_tests(answerable, user_tg_id: int, page: int = 0):
    user_db_id = await get_user_db_id(user_tg_id)
    if not user_db_id:
        return await answerable.answer("Avval /start orqali ro‘yxatdan o‘ting.")

    offset = max(0, page) * USER_LIST_PAGE_SIZE
    tests = await POOL.fetch("""
        SELECT id, code, title, price, duration, questions_count, start_mode, start_ts, start_at
        FROM tests
        WHERE is_free=0 AND COALESCE(is_hidden, 0)=0
        ORDER BY id DESC
        LIMIT $1 OFFSET $2
    """, USER_LIST_PAGE_SIZE + 1, offset)
    has_more = len(tests) > USER_LIST_PAGE_SIZE
    tests = tests[:USER_LIST_PAGE_SIZE]

    pay_rows = await POOL.fetch("""
        SELECT p.test_id, p.status
        FROM payments p
        WHERE p.user_id=$1
        ORDER BY p.id DESC
    """, user_db_id)

    if not tests:
        if page == 0:
            return await answerable.answer("Hozircha pullik testlar yo‘q.", reply_markup=search_kb("paid"))
        return await answerable.answer("Boshqa pullik test yo'q.")

    latest_status = {}
    for r in pay_rows:
        if r["test_id"] not in latest_status:
            latest_status[r["test_id"]] = r["status"]

    text = f"<b>🧾 Pullik testlar</b> ({page + 1}-sahifa)\n\n"
    kb = InlineKeyboardMarkup(inline_keyboard=[])
    kb.inline_keyboard.append([InlineKeyboardButton(text="🔎 Kod orqali topish", callback_data="askcode_paid")])

    for row in tests:
        tid, code, title, price, duration, qcount = (
            row["id"], row["code"], row["title"], row["price"], row["duration"], row["questions_count"]
        )
        start_mode, start_ts_val, start_at = row["start_mode"], row["start_ts"], row["start_at"]
        st = latest_status.get(tid)
        st_txt = ""
        if st == "approved":
            st_txt = "✅ Sizda bor"
        elif st == "pending":
            st_txt = "🟡 Kutilmoqda"
        elif st == "rejected":
            st_txt = "❌ Rad"

        sched = (start_mode or "normal") == "scheduled"
        blocked, left = is_scheduled_not_started(start_mode, start_ts_val)

        text += f"🔑 <code>{code}</code> — <b>{title}</b>\n"
        text += f"🧮 {qcount} ta | 💰 {price} | ⏳ {duration} daq"
        if st_txt:
            text += f" | {st_txt}"
        text += "\n"
        if sched:
            text += f"📅 Boshlanish: <b>{start_at}</b>"
            if blocked:
                text += f" | ⏳ {fmt_seconds(left)}"
            text += "\n"
        text += "\n"

        if st == "approved":
            kb.inline_keyboard.append(
                [InlineKeyboardButton(text=f"📦 Pullik testlarim: {code}", callback_data="open_mytests")])
        elif st == "pending":
            kb.inline_keyboard.append([InlineKeyboardButton(text=f"🟡 Pending: {code}", callback_data=f"payinfo_{tid}")])
        elif st == "rejected":
            kb.inline_keyboard.append(
                [InlineKeyboardButton(text=f"♻️ Qayta to'lov: {code}", callback_data=f"rebuy_{tid}")])
        else:
            kb.inline_keyboard.append([InlineKeyboardButton(text=f"Sotib olish: {code}", callback_data=f"buy_{tid}")])

    nav = _user_pagination_nav_row("paid", page, has_more)
    if nav:
        kb.inline_keyboard.append(nav)

    await answerable.answer(text, parse_mode=ParseMode.HTML, reply_markup=kb)


async def _resolve_payment_routing(test_id: int) -> dict:
    """Test uchun ko'rsatiladigan to'lov kartasini va to'lov screenshoti
    kimga borishini aniqlaydi:
      - Test yaratuvchisi ADMIN bo'lib, o'zining shaxsiy kartasini sozlagan
        bo'lsa — O'SHA karta ko'rsatiladi, chek TO'G'RIDAN-TO'G'RI o'ziga
        boradi (u o'zi tasdiqlaydi, komissiya ajratish shart emas).
      - Aks holda (oddiy foydalanuvchi yaratgan yoki admin karta
        sozlamagan) — umumiy standart karta ko'rsatiladi, chek asosiy
        ADMIN_ID'ga boradi. Agar aynan ODDIY foydalanuvchi yaratgan bo'lsa,
        qo'shimcha "share_reminder" bayrog'i qo'yiladi — bu orqali asosiy
        admin tegishli ulushni test yaratuvchisiga o'tkazib berishi
        kerakligi haqida eslatma oladi."""
    row = await POOL.fetchrow("""
        SELECT t.creator_id, u.full_name, u.payment_card, u.payment_owner, u.payment_phone
        FROM tests t
        LEFT JOIN users u ON u.telegram_id = t.creator_id
        WHERE t.id=$1
    """, test_id)

    creator_id = row["creator_id"] if row else None
    creator_name = (row["full_name"] if row else None) or "Noma'lum"
    # ✅ is_admin_user() ADMIN_ID uchun DB'dagi role='admin' bo'lishini talab
    # qilmaydi (maxsus bypass bor) — shuning uchun xom "role" ustunini emas,
    # aynan shu funksiyani chaqirish kerak, aks holda asosiy admin o'zining
    # profilida karta sozlash imkoniyatini ko'ra olmay qolishi mumkin edi.
    creator_is_admin = bool(creator_id) and await is_admin_user(creator_id)

    if row and creator_is_admin and row["payment_card"]:
        return {
            "card": row["payment_card"],
            "owner": row["payment_owner"] or PAYMENT_OWNER,
            "phone": row["payment_phone"] or PAYMENT_PHONE,
            "notify_tg_id": creator_id,
            "share_reminder": False,
            "creator_name": creator_name,
        }

    return {
        "card": PAYMENT_CARD,
        "owner": PAYMENT_OWNER,
        "phone": PAYMENT_PHONE,
        "notify_tg_id": ADMIN_ID,
        "share_reminder": bool(creator_id) and not creator_is_admin,
        "creator_name": creator_name,
    }


@router.message(F.text == BTN_PAID)
async def show_paid_tests(message: Message):
    await _render_paid_tests(message, message.from_user.id, 0)


@router.callback_query(F.data == "open_mytests")
async def open_mytests_cb(callback: CallbackQuery):
    await callback.message.answer("📦 Pullik testlarim bo‘limiga o‘ting (menyudan).")


@router.callback_query(F.data.startswith("payinfo_"))
async def payinfo(callback: CallbackQuery):
    await callback.message.answer("🟡 Bu test bo‘yicha to‘lov admin tasdiqlashini kutyapti.")


@router.callback_query(F.data.startswith("buy_"))
async def start_payment(callback: CallbackQuery, state: FSMContext, bot: Bot):
    if not await ensure_subscribed_callback(callback, bot):
        return

    user_db_id = await get_user_db_id(callback.from_user.id)
    if not user_db_id:
        return await callback.message.answer("Avval /start orqali ro‘yxatdan o‘ting.")

    test_id = int(callback.data.split("_")[1])
    tr = await POOL.fetchrow("""
        SELECT code, title, start_mode, start_ts, start_at, duration, is_hidden, is_closed
        FROM tests
        WHERE id=$1 AND is_free=0
    """, test_id)

    if not tr:
        return await callback.message.answer("Test topilmadi.")

    if int(tr["is_hidden"] or 0) == 1:
        return await callback.message.answer("⏸ Bu test admin tomonidan vaqtincha to'xtatilgan. Keyinroq urinib ko'ring.")
    if int(tr["is_closed"] or 0) == 1:
        return await callback.message.answer("🛑 Bu test yakunlangan, endi yangi urinish qabul qilinmaydi.")

    tcode, ttitle, tmode, tstart_ts, tstart_at, tdur = (
        tr["code"], tr["title"], tr["start_mode"], tr["start_ts"], tr["start_at"], tr["duration"]
    )

    expired, after = is_scheduled_expired(tmode, tstart_ts, tdur)
    if expired:
        return await callback.message.answer(
            f"❌ <b>Bu pullik testning vaqti tugab bo‘lgan.</b>\n\n"
            f"🔑 Kod: <code>{tcode}</code>\n"
            f"📚 Test: <b>{ttitle}</b>\n"
            f"📅 Boshlanish: <b>{tstart_at or 'Noma’lum'}</b>\n\n"
            f"✅ To‘lov qabul qilinmaydi.",
            parse_mode=ParseMode.HTML
        )

    if await POOL.fetchrow("""
        SELECT 1 FROM payments
        WHERE user_id=$1 AND test_id=$2 AND status='pending'
        LIMIT 1
    """, user_db_id, test_id):
        return await callback.message.answer(
            "🟡 Sizda shu test bo‘yicha pending to‘lov bor. Admin javobini kuting.")

    # ✅ Allaqachon sotib olingan test uchun qayta to'lov qabul qilinmaydi
    # (eski "Sotib olish" tugmasi bosilsa, o'quvchi ikki marta to'lab qo'yardi).
    approved = await POOL.fetchrow(
        "SELECT id FROM payments WHERE user_id=$1 AND test_id=$2 AND status='approved' ORDER BY id DESC LIMIT 1",
        user_db_id, test_id)
    if approved:
        return await callback.message.answer(
            "✅ Siz bu testni allaqachon sotib olgansiz — qayta to'lash shart emas.",
            reply_markup=InlineKeyboardMarkup(inline_keyboard=[[
                InlineKeyboardButton(text="🚀 Testni boshlash", callback_data=f"begin_{approved['id']}")
            ]]))

    routing = await _resolve_payment_routing(test_id)
    await state.update_data(test_id=test_id)
    await callback.message.answer(
        "💳 <b>To'lov ma'lumotlari:</b>\n\n"
        f"💳 Karta: <code>{routing['card']}</code>\n"
        f"👤 Ega: <b>{routing['owner']}</b>\n"
        f"📞 Telefon: <code>{routing['phone']}</code>\n"
        f"📝 Izoh: <i>{PAYMENT_NOTE}</i>\n\n"
        "✅ To'lov qilib bo‘lgach, chek (skrinshot) yuboring. Va tezroq tasdiqlanishi uchun berilgan raqamga telefon qiling",
        parse_mode=ParseMode.HTML,
        reply_markup=nav_kb(include_back=True)
    )
    await state.set_state(PaymentState.waiting_for_screenshot)


@router.callback_query(F.data.startswith("rebuy_"))
async def rebuy(callback: CallbackQuery, state: FSMContext, bot: Bot):
    if not await ensure_subscribed_callback(callback, bot):
        return

    user_db_id = await get_user_db_id(callback.from_user.id)
    if not user_db_id:
        return await callback.message.answer("Avval /start orqali ro‘yxatdan o‘ting.")

    test_id = int(callback.data.split("_")[1])
    tr = await POOL.fetchrow("""
        SELECT code, title, start_mode, start_ts, start_at, duration, is_hidden, is_closed
        FROM tests
        WHERE id=$1 AND is_free=0
    """, test_id)

    if not tr:
        return await callback.message.answer("Test topilmadi.")

    if int(tr["is_hidden"] or 0) == 1:
        return await callback.message.answer("⏸ Bu test admin tomonidan vaqtincha to'xtatilgan. Keyinroq urinib ko'ring.")
    if int(tr["is_closed"] or 0) == 1:
        return await callback.message.answer("🛑 Bu test yakunlangan, endi yangi urinish qabul qilinmaydi.")

    tcode, ttitle, tmode, tstart_ts, tstart_at, tdur = (
        tr["code"], tr["title"], tr["start_mode"], tr["start_ts"], tr["start_at"], tr["duration"]
    )

    expired, after = is_scheduled_expired(tmode, tstart_ts, tdur)
    if expired:
        return await callback.message.answer(
            f"❌ <b>Bu pullik testning vaqti tugab bo‘lgan.</b>\n\n"
            f"🔑 Kod: <code>{tcode}</code>\n"
            f"📚 Test: <b>{ttitle}</b>\n"
            f"📅 Boshlanish: <b>{tstart_at or 'Noma’lum'}</b>\n\n"
            f"✅ To‘lov qabul qilinmaydi.",
            parse_mode=ParseMode.HTML
        )

    if await POOL.fetchrow("""
        SELECT 1 FROM payments
        WHERE user_id=$1 AND test_id=$2 AND status='pending'
        LIMIT 1
    """, user_db_id, test_id):
        return await callback.message.answer(
            "🟡 Sizda shu test bo‘yicha pending to‘lov bor. Admin javobini kuting.")

    # ✅ Allaqachon sotib olingan test uchun qayta to'lov qabul qilinmaydi
    # (eski "Sotib olish" tugmasi bosilsa, o'quvchi ikki marta to'lab qo'yardi).
    approved = await POOL.fetchrow(
        "SELECT id FROM payments WHERE user_id=$1 AND test_id=$2 AND status='approved' ORDER BY id DESC LIMIT 1",
        user_db_id, test_id)
    if approved:
        return await callback.message.answer(
            "✅ Siz bu testni allaqachon sotib olgansiz — qayta to'lash shart emas.",
            reply_markup=InlineKeyboardMarkup(inline_keyboard=[[
                InlineKeyboardButton(text="🚀 Testni boshlash", callback_data=f"begin_{approved['id']}")
            ]]))

    routing = await _resolve_payment_routing(test_id)
    await state.update_data(test_id=test_id)
    await callback.message.answer(
        "💳 <b>To'lov ma'lumotlari:</b>\n\n"
        f"💳 Karta: <code>{routing['card']}</code>\n"
        f"👤 Ega: <b>{routing['owner']}</b>\n"
        f"📞 Telefon: <code>{routing['phone']}</code>\n"
        f"📝 Izoh: <i>{PAYMENT_NOTE}</i>\n\n"
        "✅ To'lov qilib bo‘lgach, chek (skrinshot) yuboring. Va tezroq tasdiqlanishi uchun berilgan raqamga telefon qiling",
        parse_mode=ParseMode.HTML,
        reply_markup=nav_kb(include_back=True)
    )
    await state.set_state(PaymentState.waiting_for_screenshot)


@router.message(PaymentState.waiting_for_screenshot, F.photo)
async def process_screenshot(message: Message, state: FSMContext, bot: Bot):
    data = await state.get_data()
    test_id = int(data.get("test_id") or 0)
    photo_id = message.photo[-1].file_id

    user_db_id = await get_user_db_id(message.from_user.id)
    if not user_db_id:
        await message.answer("Avval /start orqali ro‘yxatdan o‘ting.")
        await state.clear()
        return

    tr = await POOL.fetchrow("""
        SELECT code, title, start_mode, start_ts, start_at, duration
        FROM tests
        WHERE id=$1 AND is_free=0
    """, test_id)

    if not tr:
        await message.answer("Test topilmadi.")
        await state.clear()
        return

    tcode, ttitle, tmode, tstart_ts, tstart_at, tdur = (
        tr["code"], tr["title"], tr["start_mode"], tr["start_ts"], tr["start_at"], tr["duration"]
    )
    expired, after = is_scheduled_expired(tmode, tstart_ts, tdur)
    if expired:
        await message.answer(
            f"❌ <b>Bu testning vaqti tugab bo‘lgan.</b>\n\n"
            f"🔑 Kod: <code>{tcode}</code>\n"
            f"📚 Test: <b>{ttitle}</b>\n"
            f"📅 Boshlanish: <b>{tstart_at or 'Noma’lum'}</b>\n\n"
            f"✅ Chek qabul qilinmaydi.",
            parse_mode=ParseMode.HTML
        )
        await state.clear()
        return

    t = await POOL.fetchrow("SELECT code, title, price FROM tests WHERE id=$1", test_id)
    if not t:
        await message.answer("Test topilmadi.")
        await state.clear()
        return
    code, title, price = t["code"], t["title"], t["price"]

    if await POOL.fetchval(
            "SELECT 1 FROM payments WHERE user_id=$1 AND test_id=$2 AND status='approved'", user_db_id, test_id):
        await message.answer("✅ Siz bu testni allaqachon sotib olgansiz — chek qabul qilinmadi.")
        await state.clear()
        return

    try:
        payment_id = await POOL.fetchval("""
            INSERT INTO payments (user_id, test_id, screenshot_id, status, created_at, amount)
            VALUES ($1, $2, $3, 'pending', $4, (SELECT price FROM tests WHERE id=$2))
            RETURNING id
        """, user_db_id, test_id, photo_id, now_str_local())
    except asyncpg.UniqueViolationError:
        await message.answer("🟡 Sizda shu test bo‘yicha pending to‘lov bor.")
        await state.clear()
        return

    uname_row = await POOL.fetchrow("SELECT full_name FROM users WHERE id=$1", user_db_id)
    uname = uname_row["full_name"] if uname_row else "Noma'lum"

    admin_kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="✅ Tasdiqlash", callback_data=f"appr_{payment_id}")],
        [InlineKeyboardButton(text="❌ Rad etish (Sabab)", callback_data=f"rejask_{payment_id}")]
    ])

    routing = await _resolve_payment_routing(test_id)
    share_note = ""
    if routing["share_reminder"]:
        share_note = (
            f"\n\n💡 <i>Bu test oddiy foydalanuvchi — {routing['creator_name']} — tomonidan yaratilgan. "
            f"Tasdiqlasangiz, tegishli ulushni o'ziga o'tkazib berishni unutmang.</i>"
        )

    # ✅ Xabar oluvchi (test egasi-admin) botni bloklagan bo'lsa, avval o'quvchi
    # "Kutilmagan xatolik" olardi va chek kutish holatida qotib qolardi
    # (to'lov esa bazaga yozilib bo'lgan). Endi asosiy adminga zaxira yuboriladi.
    caption = (
        f"🟡 <b>Yangi to'lov (PENDING)</b>\n\n"
        f"👤 User: {_h(uname)}\n"
        f"📚 Test: {_h(title)}\n"
        f"🔑 Kod: <code>{code}</code>\n"
        f"💰 Summa: {price} so'm\n"
        f"🆔 Payment ID: <code>{payment_id}</code>"
        f"{share_note}"
    )
    for notify_id in dict.fromkeys([routing["notify_tg_id"], ADMIN_ID]):
        try:
            await bot.send_photo(chat_id=notify_id, photo=photo_id, caption=caption,
                                 parse_mode=ParseMode.HTML, reply_markup=admin_kb)
            break
        except (TelegramForbiddenError, TelegramBadRequest):
            logging.warning(f"To'lov xabarini yuborib bo'lmadi (tg_id={notify_id}, payment={payment_id})")

    admin_flag = await is_admin_user(message.from_user.id)
    await message.answer("✅ Chek qabul qilindi. Admin tasdiqlashini kuting.", reply_markup=main_menu_kb(admin_flag))
    await state.clear()


# ===================== MY TESTS =====================
async def _render_my_tests(answerable, user_tg_id: int, page: int = 0):
    user_db_id = await get_user_db_id(user_tg_id)
    if not user_db_id:
        return await answerable.answer("Avval /start orqali ro‘yxatdan o‘ting.")

    # Har bir sahifa uchun ko'proq qator olamiz, chunki bir test uchun bir necha
    # to'lov urinishi bo'lishi mumkin (masalan rad etilib, qayta to'langan) —
    # keyin ular test_id bo'yicha bittaga birlashtiriladi.
    fetch_n = USER_LIST_PAGE_SIZE * 3 + 1
    offset = max(0, page) * USER_LIST_PAGE_SIZE * 3
    rows = await POOL.fetch("""
        SELECT
            p.id AS payment_id,
            p.status,
            p.test_id,
            p.reject_reason,
            p.started_at,
            t.code,
            t.title,
            t.duration,
            t.price,
            t.questions_count,
            t.start_mode,
            t.start_ts,
            t.start_at,
            (
                SELECT r.percent
                FROM results r
                WHERE r.user_id=p.user_id AND r.test_id=p.test_id
                ORDER BY r.id ASC LIMIT 1
            ) AS first_percent
        FROM payments p
        JOIN tests t ON t.id=p.test_id
        WHERE p.user_id=$1
        ORDER BY p.id DESC
        LIMIT $2 OFFSET $3
    """, user_db_id, fetch_n, offset)

    actions = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="🔎 Kod orqali topish", callback_data="askcode_my")]
    ])

    if not rows:
        if page == 0:
            admin_flag = await is_admin_user(user_tg_id)
            await answerable.answer("Siz hali pullik test sotib olmagansiz.", reply_markup=main_menu_kb(admin_flag))
            await answerable.answer("Kod orqali qidirish:", reply_markup=actions)
        else:
            await answerable.answer("Boshqa test yo'q.")
        return

    seen = set()
    items = []
    for r in rows:
        if r["test_id"] in seen:
            continue
        seen.add(r["test_id"])
        items.append(r)
    has_more = len(items) > USER_LIST_PAGE_SIZE
    items = items[:USER_LIST_PAGE_SIZE]

    text = f"<b>📦 Mening pullik testlarim</b> ({page + 1}-sahifa)\n\n"
    for r in items:
        payment_id, status, test_id, reject_reason, started_at_paid = (
            r["payment_id"], r["status"], r["test_id"], r["reject_reason"], r["started_at"]
        )
        code, title, duration, price, qcount = r["code"], r["title"], r["duration"], r["price"], r["questions_count"]
        start_mode, start_ts_val, start_at, first_percent = r["start_mode"], r["start_ts"], r["start_at"], r["first_percent"]

        sched = (start_mode or "normal") == "scheduled"
        blocked, left = is_scheduled_not_started(start_mode, start_ts_val)

        text += f"🔑 <code>{code}</code> — <b>{title}</b>\n"
        text += f"🧮 {qcount} ta | 💰 {price} | ⏳ {duration} daq\n"
        if sched:
            text += f"📅 Boshlanish: <b>{start_at}</b>"
            if blocked:
                text += f" | ⏳ {fmt_seconds(left)}"
            text += "\n"

        if status == "approved":
            if first_percent is not None:
                text += f"✅ Natija: <b>{float(first_percent):.1f}%</b> (1-urinish)\n\n"
                actions.inline_keyboard.append(
                    [InlineKeyboardButton(text=f"📄 Natija: {code}", callback_data=f"viewres_{test_id}")])
                if started_at_paid:
                    actions.inline_keyboard.append(
                        [InlineKeyboardButton(text=f"📄 PDF: {code}", callback_data=f"pdf_{test_id}")])
            else:
                text += "✅ Tasdiqlangan (hali yechilmagan)\n\n"
                actions.inline_keyboard.append(
                    [InlineKeyboardButton(text=f"🚀 Boshlash: {code}", callback_data=f"begin_{payment_id}")])

        elif status == "pending":
            text += "🟡 Pending\n\n"
        elif status == "rejected":
            text += "❌ Rad etilgan"
            if reject_reason:
                text += f" — {_h(reject_reason)}"
            text += "\n\n"
            actions.inline_keyboard.append(
                [InlineKeyboardButton(text=f"♻️ Qayta to‘lov: {code}", callback_data=f"rebuy_{test_id}")])
        else:
            text += "ℹ️ Noma’lum holat\n\n"

    nav = _user_pagination_nav_row("mytests", page, has_more)
    if nav:
        actions.inline_keyboard.append(nav)

    await answerable.answer(text, parse_mode=ParseMode.HTML, reply_markup=actions)


@router.message(F.text == BTN_MY)
async def my_tests(message: Message):
    await _render_my_tests(message, message.from_user.id, 0)


@router.callback_query(F.data.startswith("pdf_"))
async def send_paid_pdf(callback: CallbackQuery, bot: Bot):
    test_id = int(callback.data.split("_")[1])
    user_db_id = await get_user_db_id(callback.from_user.id)
    if not user_db_id:
        return await callback.message.answer("Avval /start orqali ro‘yxatdan o‘ting.")

    row = await POOL.fetchrow("""
        SELECT t.code, t.title, t.file_id, p.started_at
        FROM payments p
        JOIN tests t ON t.id=p.test_id
        WHERE p.user_id=$1 AND p.test_id=$2 AND p.status='approved'
        ORDER BY p.id DESC LIMIT 1
    """, user_db_id, test_id)

    if not row:
        return await callback.message.answer("Bu test sizda topilmadi yoki tasdiqlanmagan.")

    code, title, file_id, started_at = row["code"], row["title"], row["file_id"], row["started_at"]
    if not started_at:
        return await callback.message.answer("⚠️ PDFni ko‘rish uchun avval testni 'Boshlash' tugmasi orqali boshlang.")

    await bot.send_document(
        chat_id=callback.from_user.id,
        document=file_id,
        caption=f"📄 <b>{title}</b>\n🔑 Kod: <code>{code}</code>",
        parse_mode=ParseMode.HTML
    )


def _render_staged_result_text(row) -> str:
    """Bosqichli test natijasini o'quvchiga ko'rsatish uchun matn (results.details
    ichidagi bosqichlar va tests.skills_cfg asosida)."""
    cfg = _load_skills_cfg(row["skills_cfg"])
    kind = cfg.get("kind")
    sections = (_load_json_dict(row["details"]).get("stages") or {})
    lines = []
    for k in cfg.get("order", []):
        sec = cfg["sections"].get(k, {})
        v = sections.get(k) or {}
        name = _h(sec.get("name", k))
        if not v:
            continue
        if sec.get("type") == STAGE_TYPE_MCQ:
            lines.append(f"• {name}: {v.get('score', 0)}/{v.get('total', 0)} "
                         f"({float(v.get('earned') or 0):.1f}/{float(v.get('max') or 0):.1f} ball)")
        elif sec.get("type") == STAGE_TYPE_MCQ_BAND:
            lines.append(f"• {name}: <b>{v.get('band', '-')}</b> ({v.get('score', 0)}/{v.get('total', 0)})")
        else:
            fb = f"\n   💬 {_h(v['feedback'])}" if v.get("feedback") else ""
            lines.append(f"• {name}: <b>{v.get('band', '-')}</b>{fb}")
    if kind == "dtm_staged":
        overall = (f"🏆 <b>Umumiy ball:</b> {float(row['earned_points'] or 0):.1f}/{float(row['max_points'] or 0):.1f} "
                   f"({float(row['percent'] or 0):.1f}%)")
    elif kind == "cefr4":
        overall = f"🏆 <b>Umumiy daraja (CEFR):</b> {row['grade'] or '-'}"
    elif kind == "sat_staged":
        overall = f"🏆 <b>Umumiy SAT ball:</b> {row['grade'] or '-'} / 1600"
    else:
        overall = f"🏆 <b>Overall band (IELTS):</b> {row['grade'] or '-'}"
    return (
        f"📄 <b>Natija</b>\n\n🔑 Kod: <code>{row['code']}</code>\n📚 Test: <b>{_h(row['title'])}</b>\n\n"
        + "\n".join(lines) + f"\n\n{overall}\n🗓 Sana: {row['date']}"
    )


@router.callback_query(F.data.startswith("viewres_"))
async def view_result(callback: CallbackQuery, bot: Bot):
    test_id = int(callback.data.split("_")[1])
    user_db_id = await get_user_db_id(callback.from_user.id)
    if not user_db_id:
        return await callback.message.answer("Avval /start orqali ro‘yxatdan o‘ting.")

    row = await POOL.fetchrow("""
        SELECT t.code, t.title, t.subject, t.is_free, t.price, t.exam_type, t.answers, t.skills_cfg,
               r.score, r.total_questions, r.percent,
               r.earned_points, r.max_points, r.details,
               r.rasch_ball, r.rasch_percent, r.grade, r.date,
               r.open_grades, COALESCE(r.open_pending,0) AS open_pending
        FROM results r
        JOIN tests t ON t.id=r.test_id
        WHERE r.user_id=$1 AND r.test_id=$2
        ORDER BY r.id ASC LIMIT 1
    """, user_db_id, test_id)

    if not row:
        return await callback.message.answer("Bu test bo‘yicha natija topilmadi.")

    # ✅ Bosqichli testlar (IELTS/CEFR/SAT/DTM 5-bosqich) — avval bu yerda
    # ularning asosiy natijasi (band/daraja/ball) umuman ko'rsatilmasdi.
    if (row["exam_type"] or "") in STAGED_EXAM_TYPES:
        return await callback.message.answer(
            _render_staged_result_text(row), parse_mode=ParseMode.HTML,
            reply_markup=InlineKeyboardMarkup(inline_keyboard=[
                [InlineKeyboardButton(text="🔎 Batafsil (bosqichlar bo'yicha)", callback_data=f"mistakes_{test_id}")],
                [InlineKeyboardButton(text="📤 Natijani ulashish", callback_data=f"share_{test_id}")],
            ]))

    (code, title, subject, is_free, price, exam_type,
     score, total, percent,
     earned_points, max_points, details_json,
     rasch_ball, rasch_percent, grade, date_) = (
        row["code"], row["title"], row["subject"], row["is_free"], row["price"], row["exam_type"],
        row["score"], row["total_questions"], row["percent"],
        row["earned_points"], row["max_points"], row["details"],
        row["rasch_ball"], row["rasch_percent"], row["grade"], row["date"]
    )

    exam_type = (exam_type or "simple")

    if exam_type == "rasch":
        await finalize_rush_for_test_if_ready(test_id, bot=bot)

        rr = await POOL.fetchrow("""
            SELECT rasch_ball, rasch_percent, grade
            FROM results
            WHERE user_id=$1 AND test_id=$2
            ORDER BY id ASC LIMIT 1
        """, user_db_id, test_id)
        if rr:
            rasch_ball, rasch_percent, grade = rr["rasch_ball"], rr["rasch_percent"], rr["grade"]

    pay_line = ""
    if int(is_free or 0) == 0:
        pay_line = f"💳 To‘lov: <b>{int(price or 0)} so‘m</b>\n"

    key_str = normalize_answers(row["answers"] or "")
    open_total = sum(1 for c in key_str if c == "x")
    open_pending = int(row["open_pending"] or 0)
    open_line = ""
    if open_total:
        if open_pending:
            open_line = (f"✍️ Ochiq savollar: {open_total} ta — {open_pending} tasi ustoz tekshiruvida "
                         f"(natija tekshirilgach yangilanadi)\n")
        else:
            og = _load_json_dict(row["open_grades"])
            open_ok = sum(1 for v in og.values() if int(v) == 1)
            open_line = f"✍️ Ochiq savollar: {open_ok}/{open_total} to'g'ri\n"
    wrong = max(0, int(total or 0) - int(score or 0) - open_pending)

    base_block = (
        f"📄 <b>Natija (1-urinish)</b>\n\n"
        f"🔑 Kod: <code>{code}</code>\n"
        f"📚 Fan: <b>{_h(subject or 'Noma’lum')}</b>\n"
        f"📚 Test: <b>{_h(title)}</b>\n"
        f"{pay_line}"
        f"✅ To‘g‘ri: {score}/{total}\n"
        f"❌ Xato: {wrong}\n"
        f"{open_line}"
        f"📊 Oddiy foiz: <b>{float(percent):.1f}%</b>\n"
    )

    dtm_block = ""
    if exam_type == "dtm":
        dtm_block = (
            f"\n🏁 Ball: <b>{float(earned_points or 0):.1f}/{float(max_points or 0):.1f}</b>\n"
            f"📌 Ball foiz: <b>{float(percent):.1f}%</b>\n"
        )

    rasch_block = ""
    if exam_type == "rasch":
        if rasch_ball is None:
            remain_text = await remaining_time_text(test_id)
            rasch_block = (
                "\n⚡ Rasch natija: <b>Hali hisoblanmadi (test tugagach chiqadi)</b>\n"
                "🎯 Daraja: <b>Test tugagach chiqadi</b>\n"
                f"🕒 Qolgan vaqt: <b>{remain_text}</b>\n"
            )
        else:
            if grade is None:
                rasch_block = (
                    f"\n⚡ Rasch ball: <b>{float(rasch_ball):.1f} / 75</b>\n"
                    f"⚡ Rasch foiz: <b>{float(rasch_percent or 0):.1f}%</b>\n"
                    f"🎯 Daraja: <b>Sertifikat yo‘q</b>\n"
                )
            else:
                rasch_block = (
                    f"\n⚡ Rasch ball: <b>{float(rasch_ball):.1f} / 75</b>\n"
                    f"⚡ Rasch foiz: <b>{float(rasch_percent or 0):.1f}%</b>\n"
                    f"🎯 Daraja: <b>{grade}</b>\n"
                )

    intl_block = ""
    if exam_type == "cefr":
        intl_block = (
            f"\n🌍 CEFR daraja (taxminiy): <b>{grade}</b>\n"
            f"<i>Bu rasmiy CEFR sertifikati emas — taxminiy konvertatsiya.</i>\n"
        )
    elif exam_type == "ielts":
        intl_block = (
            f"\n🇬🇧 IELTS band (taxminiy): <b>{grade}</b>\n"
            f"<i>Bu rasmiy IELTS balli emas — taxminiy konvertatsiya.</i>\n"
        )
    elif exam_type == "sat":
        intl_block = (
            f"\n📝 SAT ball (taxminiy): <b>{grade} / 1600</b>\n"
            f"<i>Bu rasmiy SAT balli emas — taxminiy konvertatsiya.</i>\n"
        )

    # ✅ Reytingdagi o'rni: shu testni yechgan boshqalarning 1-urinish natijalari orasida
    rank_row = await POOL.fetchrow("""
        SELECT COUNT(*) AS total,
               SUM(CASE WHEN r.percent > $2 THEN 1 ELSE 0 END) AS better
        FROM results r
        JOIN (
            SELECT user_id, MIN(id) AS min_id FROM results WHERE test_id=$1 GROUP BY user_id
        ) x ON x.min_id = r.id
    """, test_id, float(percent))
    rank_block = ""
    if rank_row and rank_row["total"]:
        rank_block = f"\n🏅 Reytingda: <b>{int(rank_row['better'] or 0) + 1} / {rank_row['total']}</b> o'rin\n"

    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="🔎 Xatolarni ko'rish", callback_data=f"mistakes_{test_id}")],
        [InlineKeyboardButton(text="📤 Natijani ulashish", callback_data=f"share_{test_id}")],
    ])

    await callback.message.answer(
        base_block + dtm_block + rasch_block + intl_block + rank_block + f"\n🗓 Sana: {date_}",
        parse_mode=ParseMode.HTML,
        reply_markup=kb
    )


@router.callback_query(F.data.startswith("mistakes_"))
async def view_mistakes(callback: CallbackQuery):
    test_id = int(callback.data.split("_")[1])
    user_db_id = await get_user_db_id(callback.from_user.id)
    if not user_db_id:
        return await callback.message.answer("Avval /start orqali ro‘yxatdan o‘ting.")

    row = await POOL.fetchrow("""
        SELECT r.user_answers, r.open_text_answers, r.open_grades, r.details, t.answers, t.questions_count,
               t.exam_type, t.open_ref_answers, t.skills_cfg
        FROM results r JOIN tests t ON t.id=r.test_id
        WHERE r.user_id=$1 AND r.test_id=$2
        ORDER BY r.id ASC LIMIT 1
    """, user_db_id, test_id)

    if row and (row["exam_type"] or "") in STAGED_EXAM_TYPES:
        return await _send_staged_mistakes(callback.message, row)

    if not row or not row["user_answers"]:
        return await callback.message.answer("❌ Bu test bo'yicha javoblaringiz topilmadi.")

    # ✅ Rasch — guruh testi tugamaguncha xato tahlilini ko'rsatmaymiz,
    # aks holda erta tugatgan o'quvchi hali yechayotganlarga javob "sizib
    # chiqarishi" mumkin edi.
    exam_type = row["exam_type"] or "simple"
    if exam_type == "rasch":
        end_ts = await get_test_end_ts(test_id)
        if end_ts <= 0 or now_ts() < end_ts:
            remain_text = await remaining_time_text(test_id)
            return await callback.message.answer(
                f"⏳ Guruh xavfsizligi uchun batafsil xato tahlili faqat test vaqti "
                f"tugagach ko'rsatiladi.\n🕒 <b>{remain_text}</b>",
                parse_mode=ParseMode.HTML
            )

    user_ans = normalize_answers(row["user_answers"])
    correct_key = normalize_answers(row["answers"])
    qcount = min(int(row["questions_count"] or 0), len(correct_key), len(user_ans))

    open_answers = {}
    if row["open_text_answers"]:
        try:
            open_answers = json.loads(row["open_text_answers"])
        except json.JSONDecodeError:
            open_answers = {}

    open_refs = {}
    if row["open_ref_answers"]:
        try:
            open_refs = json.loads(row["open_ref_answers"])
        except json.JSONDecodeError:
            open_refs = {}

    open_grades = _load_json_dict(row["open_grades"])
    mistakes = []
    open_lines = []
    open_photos = []
    for i in range(qcount):
        if correct_key[i] == "x":
            raw = open_answers.get(str(i + 1))
            g = open_grades.get(str(i + 1))
            verdict = "⏳ tekshirilmoqda" if g is None else ("✅ to'g'ri" if int(g) == 1 else "❌ xato")
            if raw:
                text_val, photo_id = _open_answer_preview(raw)
                if photo_id:
                    open_photos.append((i + 1, photo_id))
                line = f"{i + 1}-savol (ochiq): <i>{_h(text_val) if text_val else '[rasm yuborilgan]'}</i> — {verdict}"
            else:
                line = f"{i + 1}-savol (ochiq): javob yozilmagan — ❌"
            ref = open_refs.get(str(i + 1))
            if ref:
                line += f"\n   🔑 To'g'ri javob: <b>{_h(ref)}</b>"
            open_lines.append(line)
            continue
        if user_ans[i] != correct_key[i]:
            miss_char = "belgilamagansiz" if user_ans[i] == "*" else f"siz: {user_ans[i].upper()}"
            mistakes.append(f"{i + 1}-savol: {miss_char}, to'g'ri javob: <b>{correct_key[i].upper()}</b>")

    if not mistakes and not open_lines:
        return await callback.message.answer("🎉 Xatoyingiz yo'q — barcha savollarga to'g'ri javob bergansiz!")

    text = ""
    if mistakes:
        text += f"🔎 <b>Xatolaringiz</b> ({len(mistakes)} ta):\n\n" + "\n".join(mistakes[:40])
        if len(mistakes) > 40:
            text += f"\n... va yana {len(mistakes) - 40} ta"
    if open_lines:
        text += ("\n\n" if text else "") + "✍️ <b>Ochiq savol javoblaringiz:</b>\n\n" + "\n".join(open_lines[:20])

    await callback.message.answer(text, parse_mode=ParseMode.HTML)
    for qnum, photo_id in open_photos[:20]:
        await callback.message.answer_photo(photo_id, caption=f"{qnum}-savol javobi (rasm)")


async def _send_staged_mistakes(answerable, row) -> None:
    """O'quvchi uchun bosqichli test bo'yicha batafsil: har bir test-bosqichdagi
    xatolar va Writing/Speaking fikr-mulohazasi."""
    cfg = _load_skills_cfg(row["skills_cfg"])
    sections = (_load_json_dict(row["details"]).get("stages") or {})
    parts = []
    for k in cfg.get("order", []):
        sec = cfg["sections"].get(k, {})
        v = sections.get(k) or {}
        name = _h(sec.get("name", k))
        if sec.get("type") in (STAGE_TYPE_MCQ, STAGE_TYPE_MCQ_BAND):
            ck = normalize_answers(sec.get("answers") or "")
            ua = normalize_answers(v.get("user_answers") or "")
            grades = v.get("open_grades") or {}
            errs = []
            for i in range(min(len(ck), len(ua))):
                if ck[i] == "x":
                    g = grades.get(str(i + 1))
                    if g is not None and int(g) == 0:
                        ref = (sec.get("open_ref_answers") or {}).get(str(i + 1), "-")
                        errs.append(f"{i + 1}-savol (ochiq): to'g'ri javob — {_h(ref)}")
                    continue
                if ua[i] != ck[i]:
                    mine = "belgilanmagan" if ua[i] == "*" else ua[i].upper()
                    errs.append(f"{i + 1}-savol: siz {mine}, to'g'ri {ck[i].upper()}")
            body = "\n".join(errs[:30]) if errs else "Xatolar yo'q ✅"
            if len(errs) > 30:
                body += f"\n... yana {len(errs) - 30} ta"
            parts.append(f"📖 <b>{name}</b>\n{body}")
        elif v:
            fb = _h(v.get("feedback") or "-")
            parts.append(f"✍️ <b>{name}</b>: {v.get('band', '-')}\n💬 {fb}")
    text = "\n\n".join(parts) if parts else "Ma'lumot topilmadi."
    await answerable.answer(text[:4000], parse_mode=ParseMode.HTML)


@router.callback_query(F.data.startswith("share_"))
async def share_result(callback: CallbackQuery, bot: Bot):
    test_id = int(callback.data.split("_")[1])
    user_db_id = await get_user_db_id(callback.from_user.id)
    if not user_db_id:
        return await callback.message.answer("Avval /start orqali ro‘yxatdan o‘ting.")

    row = await POOL.fetchrow("""
        SELECT t.title, r.score, r.total_questions, r.percent
        FROM results r JOIN tests t ON t.id=r.test_id
        WHERE r.user_id=$1 AND r.test_id=$2
        ORDER BY r.id ASC LIMIT 1
    """, user_db_id, test_id)
    if not row:
        return await callback.message.answer("Natija topilmadi.")

    me = await bot.get_me()
    bot_link = f"https://t.me/{me.username}" if me.username else ""

    share_text = (
        f"🏆 Men \"{row['title']}\" testida {row['score']}/{row['total_questions']} "
        f"(<b>{float(row['percent']):.1f}%</b>) natija ko'rsatdim!\n\n"
        f"📚 {ACADEMY_NAME} test tizimida siz ham sinab ko'ring 👇\n{bot_link}"
    )
    await callback.message.answer(
        "📤 <b>Do'stlaringizga ulashish uchun matn tayyor</b> — pastdagi xabarni forward qiling yoki nusxa oling:",
        parse_mode=ParseMode.HTML
    )
    await callback.message.answer(share_text, parse_mode=ParseMode.HTML)


# ===================== PAID TEST BEGIN =====================
@router.callback_query(F.data.startswith("begin_"))
async def begin_test(callback: CallbackQuery, state: FSMContext, bot: Bot):
    if not await ensure_subscribed_callback(callback, bot):
        return

    payment_id = int(callback.data.split("_")[1])
    user_db_id = await get_user_db_id(callback.from_user.id)
    if not user_db_id:
        return await callback.message.answer("Avval /start orqali ro‘yxatdan o‘ting.")

    row = await POOL.fetchrow("""
        SELECT p.status, p.started_at, p.started_ts, p.test_id,
               t.code, t.title, t.file_id, t.duration, t.questions_count, t.answers,
               t.start_mode, t.start_ts, t.start_at, t.exam_type, t.is_closed
        FROM payments p
        JOIN tests t ON t.id=p.test_id
        WHERE p.id=$1 AND p.user_id=$2
    """, payment_id, user_db_id)

    if not row:
        return await callback.message.answer("Xatolik: payment topilmadi.")

    (status, started_at_paid, started_ts_paid, test_id,
     code, title, file_id, duration_min, qcount, answers,
     start_mode, start_ts_val, start_at, exam_type) = (
        row["status"], row["started_at"], row["started_ts"], row["test_id"],
        row["code"], row["title"], row["file_id"], row["duration"], row["questions_count"], row["answers"],
        row["start_mode"], row["start_ts"], row["start_at"], row["exam_type"]
    )

    if status != "approved":
        return await callback.message.answer("Bu test hali tasdiqlanmagan.")
    if int(row["is_closed"] or 0) == 1 and not await POOL.fetchval(
            "SELECT 1 FROM sessions WHERE user_id=$1 AND test_id=$2", user_db_id, test_id):
        return await callback.message.answer("🛑 Bu test yakunlangan, endi yangi urinish qabul qilinmaydi.")

    # ✅ Bosqichli testlar (DTM 5-fan / IELTS 4-ko'nikma) — o'z alohida
    # motoriga yo'naltiramiz (to'lov allaqachon tasdiqlangan, davom etamiz).
    if (exam_type or "") in STAGED_EXAM_TYPES:
        # to'lovi tasdiqlangan xaridor yashirilgan testni ham davom ettira oladi
        return await start_staged_attempt(callback.from_user.id, test_id, callback.message, bot, state,
                                          allow_hidden=True)

    blocked, left = is_scheduled_not_started(start_mode, start_ts_val)
    if blocked:
        kb = InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text="🔁 Qayta tekshirish", callback_data=f"chkpaid_{payment_id}")]
        ])
        return await callback.message.answer(
            f"⏳ <b>Test hali boshlanmagan</b>\n\n"
            f"🔑 Kod: <code>{code}</code>\n"
            f"📅 Boshlanish: <b>{start_at}</b>\n"
            f"⏳ Qolgan: <b>{fmt_seconds(left)}</b>\n\n"
            f"Iltimos, kuting.",
            parse_mode=ParseMode.HTML,
            reply_markup=kb
        )

    if await POOL.fetchrow("SELECT 1 FROM results WHERE user_id=$1 AND test_id=$2 LIMIT 1", user_db_id, test_id):
        return await callback.message.answer(
            "⚠️ Siz bu testni allaqachon topshirgansiz. Qayta topshirib bo‘lmaydi.")

    answer_key = normalize_answers(answers)
    if len(answer_key) != int(qcount):
        return await callback.message.answer(
            "⚠️ Admin testni xato qo‘shgan: javob kaliti uzunligi savol soniga teng emas.")

    cancel_timer(callback.from_user.id, test_id)

    # ✅ Agar test "oddiy" rejimda bo'lsa, hisoblagich birinchi marta "Boshlash"
    # bosilgan vaqtdan (payments.started_ts) davom etadi — qayta bosilganda
    # QAYTA BOSHLANMAYDI. Bu ham himoya (vaqtni cheksiz uzaytirib bo'lmaydi),
    # ham qulaylik (uzilib qolgan testni xavfsiz davom ettirish mumkin).
    if (start_mode or "normal") == "scheduled" and int(start_ts_val or 0) > 0:
        exam_start_ts = int(start_ts_val)
    elif started_ts_paid:
        exam_start_ts = int(started_ts_paid)
    else:
        exam_start_ts = now_ts()

    duration_sec = int(duration_min) * 60
    end_ts_ = exam_start_ts + duration_sec

    if now_ts() > end_ts_ + 60:
        return await callback.message.answer("❌ Bu test vaqti tugab bo‘lgan.")

    if not started_ts_paid:
        await POOL.execute("UPDATE payments SET started_at=$1, started_ts=$2 WHERE id=$3",
                            now_str_local(), exam_start_ts, payment_id)

    await upsert_session(user_db_id, test_id, "paid", exam_start_ts, duration_sec)

    await bot.send_document(
        chat_id=callback.from_user.id,
        document=file_id,
        caption=(
            f"📄 <b>{title}</b>\n"
            f"🔑 Kod: <code>{code}</code>\n"
            f"🧮 Savollar: <b>{qcount}</b>\n\n"
            f"✅ Test boshlandi. Vaqt ketdi!"
        ),
        parse_mode=ParseMode.HTML
    )

    timer_msg = await bot.send_message(
        chat_id=callback.from_user.id,
        text="⏳ <b>Qolgan vaqt:</b> ...",
        parse_mode=ParseMode.HTML
    )

    ACTIVE_TIMERS[(callback.from_user.id, test_id)] = asyncio.create_task(
        start_countdown(bot, callback.from_user.id, test_id, end_ts_, timer_msg.message_id)
    )

    # ✅ O'quvchi avval shu testni boshlab, menyuga chiqib ketgan bo'lsa —
    # belgilagan javoblari tiklanadi (yordam bo'limida va'da qilinganidek).
    # ✅ O'quvchi test yechayotganda eski xabardagi "Boshlash"ni qayta bossa,
    # joriy (hali saqlanmagan) javoblari avval sessiyaga yoziladi — aks holda
    # holat ustidan yozilib, belgilagan javoblari o'chib ketardi.
    await _stash_test_progress(state, callback.from_user.id)
    saved = await _restore_test_progress(user_db_id, test_id, "paid")
    saved_answers = saved.get("answers") or {}
    await state.update_data(
        test_id=test_id, mode="paid", qcount=qcount, current_q=1, test_answers=saved_answers,
        test_open_answers=saved.get("open_answers") or {}, test_answers_key=answer_key
    )
    resumed_note = (f"♻️ Avvalgi <b>{len(saved_answers)}</b> ta javobingiz tiklandi.\n"
                    if saved_answers else "")
    await bot.send_message(
        chat_id=callback.from_user.id,
        text=(
            f"📝 <b>Test boshlandi!</b>\n\n"
            f"{resumed_note}"
            f"{question_status_note(answer_key, 1)}"
            f"Savol: <b>1 / {qcount}</b>\n"
            f"Pastdagi tugmalar orqali javoblarni interaktiv belgilang.\n"
            f"<i>(Yoki xohlasangiz odatdagidek abcd... formatida birdaniga yozib yuboring)</i>"
        ),
        parse_mode=ParseMode.HTML,
        reply_markup=test_ui_kb(1, int(qcount), saved_answers, answer_key)
    )
    await state.set_state(TestProcessState.solving)


@router.callback_query(F.data.startswith("chkpaid_"))
async def check_paid_start(callback: CallbackQuery):
    payment_id = int(callback.data.split("_")[1])
    user_db_id = await get_user_db_id(callback.from_user.id)
    if not user_db_id:
        return await callback.message.answer("Avval /start orqali ro‘yxatdan o‘ting.")

    row = await POOL.fetchrow("""
        SELECT t.code, t.start_mode, t.start_ts, t.start_at
        FROM payments p
        JOIN tests t ON t.id=p.test_id
        WHERE p.id=$1 AND p.user_id=$2
    """, payment_id, user_db_id)

    if not row:
        return await callback.message.answer("Topilmadi.")

    code, start_mode, start_ts_val, start_at = row["code"], row["start_mode"], row["start_ts"], row["start_at"]
    blocked, left = is_scheduled_not_started(start_mode, start_ts_val)
    if blocked:
        return await callback.message.answer(
            f"⏳ Hali erta.\n🔑 <code>{code}</code>\n📅 {start_at}\n⏳ Qolgan: {fmt_seconds(left)}",
            parse_mode=ParseMode.HTML
        )
    await callback.message.answer("✅ Boshlash vaqti keldi! Endi 'Boshlash' tugmasini bosing.")


# ===================== DTM CONFIG =====================
DTM_DEFAULT_CFG = {
    "subjects": [
        {"name": "1-fan", "q": 30, "w": 3.1},
        {"name": "2-fan", "q": 30, "w": 2.1},
        {"name": "3-fan", "q": 10, "w": 1.1},
        {"name": "4-fan", "q": 10, "w": 1.1},
        {"name": "5-fan", "q": 10, "w": 1.1},
    ]
}


def dtm_total_questions(cfg: dict) -> int:
    return sum(int(s["q"]) for s in cfg["subjects"])


def dtm_score_points(user_ans: str, key: str, cfg_json: str,
                     open_grades: dict | None = None) -> tuple[float, float, dict]:
    cfg = json.loads(cfg_json) if cfg_json else DTM_DEFAULT_CFG
    idx = 0
    earned = 0.0
    maxp = 0.0
    details = {"subjects": []}

    for s in cfg["subjects"]:
        q = int(s["q"])
        w = float(s["w"])
        part_user = user_ans[idx:idx + q]
        part_key = key[idx:idx + q]

        correct, scored_q, _pending = score_with_open(part_user, part_key, q, open_grades, offset=idx)

        earned_s = correct * w
        max_s = scored_q * w

        earned += earned_s
        maxp += max_s

        details["subjects"].append({
            "name": s.get("name"),
            "q": q,
            "w": w,
            "correct": correct,
            "earned": earned_s,
            "max": max_s
        })
        idx += q

    return earned, maxp, details


# ===================== STAGED EXAM MODULE (DTM 5-fan / IELTS 4-ko'nikma) =====================
# ✅ Ikkala talab ("DTM'ni 5 bosqichga bo'lish" va "IELTS'ning 4 ko'nikmasi
# ketma-ket kelsin") bitta umumiy motorga qurilgan: har test N ta "bosqich"dan
# iborat, har biri o'z vaqt chegarasiga ega, ketma-ket o'tiladi. Bosqich turi:
#   - mcq        : oddiy variantli test, DTM fan og'irligi bilan baholanadi
#   - mcq_band   : variantli test, lekin natija haqiqiy IELTS band jadvaliga
#                  o'tkaziladi (Reading/Listening)
#   - open_text  : matn/rasm bilan javob — ustoz baholaydi (Writing)
#   - open_voice : ovozli xabar bilan javob — ustoz baholaydi (Speaking)
STAGE_TYPE_MCQ = "mcq"
STAGE_TYPE_MCQ_BAND = "mcq_band"
STAGE_TYPE_TEXT = "open_text"
STAGE_TYPE_VOICE = "open_voice"

# Haqiqiy DTM tuzilishi: 2 ta ixtisoslik (asosiy) fan 30 savoldan (3.1 va 2.1
# ball), 3 ta majburiy fan (Tarix, Ona tili, Matematika) 10 savoldan (1.1 ball).
DTM_STAGE_SPEC = [
    {"key": "fan1", "toggle": True, "ask_name": True, "type": STAGE_TYPE_MCQ,
     "qcount": 30, "weight": 3.1, "default_label": "1-fan (ixtisoslik, asosiy)"},
    {"key": "fan2", "toggle": True, "ask_name": True, "type": STAGE_TYPE_MCQ,
     "qcount": 30, "weight": 2.1, "default_label": "2-fan (ixtisoslik)"},
    {"key": "fan3", "toggle": True, "ask_name": False, "type": STAGE_TYPE_MCQ,
     "qcount": 10, "weight": 1.1, "default_label": "Tarix"},
    {"key": "fan4", "toggle": True, "ask_name": False, "type": STAGE_TYPE_MCQ,
     "qcount": 10, "weight": 1.1, "default_label": "Ona tili"},
    {"key": "fan5", "toggle": True, "ask_name": False, "type": STAGE_TYPE_MCQ,
     "qcount": 10, "weight": 1.1, "default_label": "Matematika"},
]

# Foydalanuvchi tasvirlagan tartib: Reading -> Writing -> Speaking -> Listening.
IELTS4_STAGE_SPEC = [
    {"key": "reading", "toggle": True, "ask_name": False, "type": STAGE_TYPE_MCQ_BAND, "default_label": "📖 Reading"},
    {"key": "writing", "toggle": True, "ask_name": False, "type": STAGE_TYPE_TEXT, "default_label": "✍️ Writing"},
    {"key": "speaking", "toggle": True, "ask_name": False, "type": STAGE_TYPE_VOICE, "default_label": "🎙 Speaking"},
    {"key": "listening", "toggle": True, "ask_name": False, "type": STAGE_TYPE_MCQ_BAND, "default_label": "🎧 Listening"},
]

CEFR4_STAGE_SPEC = [
    {"key": "reading", "toggle": True, "ask_name": False, "type": STAGE_TYPE_MCQ_BAND, "default_label": "📖 Reading"},
    {"key": "writing", "toggle": True, "ask_name": False, "type": STAGE_TYPE_TEXT, "default_label": "✍️ Writing"},
    {"key": "speaking", "toggle": True, "ask_name": False, "type": STAGE_TYPE_VOICE, "default_label": "🎙 Speaking"},
    {"key": "listening", "toggle": True, "ask_name": False, "type": STAGE_TYPE_MCQ_BAND, "default_label": "🎧 Listening"},
]

# ✅ Rasmiy (raqamli) SAT tuzilishi: ikkita majburiy bo'lim — Reading&Writing
# va Math — har biri o'z vaqtida, har biri o'z 200-800 shkala balliga ega,
# ikkalasi qo'shilib 400-1600 kompozit ball hosil qiladi. Yozma/og'zaki
# qism yo'q (zamonaviy raqamli SAT'da esse ixtiyoriy/bekor qilingan),
# shuning uchun ikkalasi ham to'liq avtomatik baholanadigan mcq_band turida.
SAT_STAGE_SPEC = [
    {"key": "reading_writing", "toggle": True, "ask_name": False, "type": STAGE_TYPE_MCQ_BAND,
     "default_label": "📖 Reading & Writing"},
    {"key": "math", "toggle": True, "ask_name": False, "type": STAGE_TYPE_MCQ_BAND,
     "default_label": "🧮 Math"},
]

STAGE_SPEC_BY_KIND = {
    "dtm_staged": DTM_STAGE_SPEC, "ielts4": IELTS4_STAGE_SPEC, "cefr4": CEFR4_STAGE_SPEC,
    "sat_staged": SAT_STAGE_SPEC,
}
STAGED_EXAM_TYPES = ("dtm_staged", "ielts4", "cefr4", "sat_staged")

CEFR_LEVEL_ORDER = ["A1", "A2", "B1", "B2", "C1", "C2"]


def _cefr_level_to_ordinal(level: str) -> int:
    return CEFR_LEVEL_ORDER.index(level) if level in CEFR_LEVEL_ORDER else 0


def _cefr_ordinal_to_level(ordinal: float) -> str:
    # round() Python'da "bank yaxlitlashi" (2.5 -> 2, 3.5 -> 4) — bir xil
    # holat turlicha yaxlitlanardi. Har doim yarimni yuqoriga yaxlitlaymiz.
    idx = max(0, min(len(CEFR_LEVEL_ORDER) - 1, int(math.floor(ordinal + 0.5))))
    return CEFR_LEVEL_ORDER[idx]


def _stage_band_for_mcq(kind: str, key: str, score: int, total: int):
    """MCQ (Reading/Listening/Math) bosqichi uchun xom balldan darajaga
    o'tkazadi — IELTS uchun raqamli band (haqiqiy jadval), CEFR uchun A1-C2
    daraja, SAT uchun 200-800 bo'lim balli."""
    if kind == "cefr4":
        pct = (score / total * 100) if total else 0.0
        return cefr_level_from_percent(pct)
    if kind == "sat_staged":
        pct = (score / total * 100) if total else 0.0
        return sat_section_score_from_percent(pct)
    if key == "listening":
        return ielts_listening_band(score, total)
    return ielts_reading_band(score, total)


def _stage_ai_scale(kind: str) -> str:
    return "cefr" if kind == "cefr4" else "ielts"


def _compute_overall_stage_grade(kind: str, sections: dict, order: list):
    """Barcha bosqichlarning bandi/darajasidan umumiy natijani hisoblaydi.
    IELTS uchun — o'rtacha raqamli band (0.5 qadamli). CEFR uchun — har bir
    darajani ordinal songa aylantirib o'rtachasini olib, yana darajaga
    qaytaradi (masalan B1+B2 o'rtachasi ~B1/B2 oralig'ida bo'ladi). SAT
    uchun — rasmiy SAT kabi ikkala bo'lim balli QO'SHILADI (o'rtacha emas),
    400-1600 kompozit hosil qilish uchun."""
    values = [sections[k]["band"] for k in order if sections.get(k, {}).get("band") is not None]
    if kind == "cefr4":
        if not values:
            return "A1"
        avg_ordinal = sum(_cefr_level_to_ordinal(v) for v in values) / len(values)
        return _cefr_ordinal_to_level(avg_ordinal)
    if kind == "sat_staged":
        return int(sum(int(v) for v in values)) if values else 0
    if not values:
        return 0.0
    nums = [float(v) for v in values]
    # ✅ Rasmiy IELTS qoidasi: o'rtacha .25 -> .5 ga, .75 -> keyingi butunga
    # yaxlitlanadi. Python round() (bank yaxlitlashi) 6.25 ni 6.0 qilardi.
    return math.floor((sum(nums) / len(nums)) * 2 + 0.5) / 2


def _stage_spec_list(kind: str) -> list[dict]:
    return STAGE_SPEC_BY_KIND[kind]


def _load_skills_cfg(raw: str | None) -> dict:
    if not raw:
        return {"kind": None, "order": [], "sections": {}}
    try:
        return json.loads(raw)
    except (json.JSONDecodeError, TypeError):
        return {"kind": None, "order": [], "sections": {}}


# ----- Yaratish sehrgari (admin/o'qituvchi) -----
async def start_staged_wizard(callback: CallbackQuery, state: FSMContext, kind: str) -> None:
    await state.update_data(
        stage_kind=kind, stage_spec_idx=0, stage_cfg={}, stage_creator_tg_id=callback.from_user.id
    )
    await _ask_next_stage(callback.message, state)


async def _ask_next_stage(answerable, state: FSMContext) -> None:
    data = await state.get_data()
    kind = data["stage_kind"]
    idx = data["stage_spec_idx"]
    spec_list = _stage_spec_list(kind)

    if idx >= len(spec_list):
        await _finalize_staged_test(answerable, state)
        return

    spec = spec_list[idx]
    progress = f"[{idx + 1}/{len(spec_list)}] "
    if spec["toggle"]:
        kb = InlineKeyboardMarkup(inline_keyboard=[[
            InlineKeyboardButton(text="✅ Ha, bo'lsin", callback_data="stg_on"),
            InlineKeyboardButton(text="➖ Yo'q, o'tkazib yuborish", callback_data="stg_off"),
        ]])
        await answerable.answer(
            f"{progress}<b>{spec['default_label']}</b> ushbu testda bo'lsinmi?",
            parse_mode=ParseMode.HTML, reply_markup=kb
        )
        return

    await _ask_stage_name_or_duration(answerable, state, spec, progress)


async def _ask_stage_name_or_duration(answerable, state: FSMContext, spec: dict, progress: str = "") -> None:
    """Bosqich yoqilgach (toggle bo'lsin, majburiy bo'lsin) navbatdagi qadamga
    o'tkazadi — agar fanning nomi so'ralishi kerak bo'lsa (masalan DTM'ning
    ixtisoslik fanlari) avval shuni, aks holda to'g'ridan-to'g'ri davomiylikni
    so'raydi. Bitta joyda saqlanadi — toggle "Ha" bosilganda ham, majburiy
    bosqichda ham AYNAN bir xil ishlashi kerak (avval ikkalasi alohida yozilib,
    toggle yo'li ask_name'ni chetlab o'tib qo'yishi mumkin edi)."""
    if spec.get("ask_name"):
        await state.update_data(stage_field="name")
        await answerable.answer(
            f"{progress}<b>{spec['default_label']}</b>"
            + (f" — {spec['qcount']} ta savol, og'irligi {spec['weight']}." if spec.get("qcount") else ".")
            + "\nBu fanning nomini kiriting (masalan: <i>Fizika</i>):",
            parse_mode=ParseMode.HTML
        )
        await state.set_state(StagedBuildState.collecting)
        return

    await state.update_data(stage_current_name=spec["default_label"])
    await _ask_stage_duration(answerable, state, spec)


async def _ask_stage_duration(answerable, state: FSMContext, spec: dict) -> None:
    data = await state.get_data()
    label = data.get("stage_current_name", spec["default_label"])
    await answerable.answer(
        f"<b>{label}</b> uchun davomiylikni daqiqada kiriting (masalan: 40):",
        parse_mode=ParseMode.HTML
    )
    await state.update_data(stage_field="duration")
    await state.set_state(StagedBuildState.collecting)


async def _ask_stage_prompt_or_file(answerable, state: FSMContext, spec: dict) -> None:
    data = await state.get_data()
    label = data.get("stage_current_name", spec["default_label"])
    if spec["type"] in (STAGE_TYPE_MCQ, STAGE_TYPE_MCQ_BAND):
        if spec["key"] == "listening":
            # ✅ Listening'da IKKITA alohida fayl kerak: avval audio/video
            # yozuvning o'zi, keyin savollar bitilgan PDF (Reading'dan farqli —
            # u yerda bitta PDF yetarli, chunki matn ham, savollar ham o'shanda).
            await answerable.answer(
                f"<b>{label}</b> uchun AUDIO yoki VIDEO faylni yuboring "
                f"(o'quvchi shuni tinglaydi):",
                parse_mode=ParseMode.HTML
            )
            await state.update_data(stage_field="listening_media")
        else:
            await answerable.answer(
                f"<b>{label}</b> uchun savollar bo'lgan PDF faylni yuboring:",
                parse_mode=ParseMode.HTML
            )
            await state.update_data(stage_field="file")
    elif spec["type"] == STAGE_TYPE_TEXT:
        await answerable.answer(
            f"<b>{label}</b> uchun topshiriq matnini kiriting (masalan: Writing Task 2 mavzusi):",
            parse_mode=ParseMode.HTML
        )
        await state.update_data(stage_field="prompt")
    else:  # STAGE_TYPE_VOICE
        await answerable.answer(
            f"<b>{label}</b> uchun mavzu/savol matnini kiriting (o'quvchi shu mavzuda ovozli javob beradi):",
            parse_mode=ParseMode.HTML
        )
        await state.update_data(stage_field="prompt")
    await state.set_state(StagedBuildState.collecting)


async def _ask_stage_qcount(answerable, state: FSMContext, spec: dict) -> None:
    data = await state.get_data()
    label = data.get("stage_current_name", spec["default_label"])
    await answerable.answer(
        f"<b>{label}</b> uchun savollar sonini kiriting (masalan: 40):",
        parse_mode=ParseMode.HTML
    )
    await state.update_data(stage_field="qcount")
    await state.set_state(StagedBuildState.collecting)


async def _ask_stage_answers(answerable, state: FSMContext, spec: dict, qcount: int) -> None:
    data = await state.get_data()
    label = data.get("stage_current_name", spec["default_label"])
    await answerable.answer(
        f"<b>{label}</b> uchun javob kalitini kiriting — aynan <b>{qcount} ta</b> harf "
        f"(A-F), bo'shliqsiz. Variantsiz (ochiq) savol bo'lsa, o'rniga <code>x</code> qo'ying:",
        parse_mode=ParseMode.HTML
    )
    await state.update_data(stage_field="answers")
    await state.set_state(StagedBuildState.collecting)


async def _save_current_stage_and_advance(answerable, state: FSMContext, section_data: dict) -> None:
    data = await state.get_data()
    kind = data["stage_kind"]
    idx = data["stage_spec_idx"]
    spec = _stage_spec_list(kind)[idx]

    stage_cfg = dict(data.get("stage_cfg") or {})
    stage_cfg[spec["key"]] = section_data
    await state.update_data(stage_cfg=stage_cfg, stage_spec_idx=idx + 1, stage_current_name=None)
    await _ask_next_stage(answerable, state)


@router.callback_query(F.data.in_({"stg_on", "stg_off"}))
async def staged_toggle_cb(callback: CallbackQuery, state: FSMContext) -> None:
    data = await state.get_data()
    kind = data.get("stage_kind")
    idx = data.get("stage_spec_idx")
    if kind is None or idx is None:
        await callback.answer()
        return
    spec = _stage_spec_list(kind)[idx]

    if callback.data == "stg_off":
        await _nav_push(state, "staged", "toggle", idx)
        await state.update_data(stage_spec_idx=idx + 1)
        await callback.answer("O'tkazib yuborildi")
        await _ask_next_stage(callback.message, state)
        return

    await callback.answer()
    await _nav_push(state, "staged", "toggle", idx)
    progress = f"[{idx + 1}/{len(_stage_spec_list(kind))}] "
    await _ask_stage_name_or_duration(callback.message, state, spec, progress)


@router.message(StagedBuildState.collecting, F.document | F.audio | F.video)
async def staged_build_file(message: Message, state: FSMContext) -> None:
    data = await state.get_data()
    field = data.get("stage_field")
    if field not in ("file", "listening_media", "listening_pdf"):
        await message.answer("⚠️ Hozir fayl emas, yuqoridagi savolga javob (matn yoki tugma) kutilmoqda.")
        return
    kind = data["stage_kind"]
    idx = data["stage_spec_idx"]
    spec = _stage_spec_list(kind)[idx]

    if message.document:
        file_id, file_kind = message.document.file_id, "document"
    elif message.audio:
        file_id, file_kind = message.audio.file_id, "audio"
    else:
        file_id, file_kind = message.video.file_id, "video"

    if field == "listening_media":
        # ✅ Audio/video qabul qilindi — endi alohida savollar PDF'ini so'raymiz.
        await state.update_data(stage_pending_file_id=file_id, stage_pending_file_kind=file_kind)
        await message.answer("✅ Audio/video qabul qilindi.\n\nEndi savollar bo'lgan PDF faylni yuboring:")
        await state.update_data(stage_field="listening_pdf")
        return

    if field == "listening_pdf":
        await state.update_data(stage_pending_question_file_id=file_id)
    else:
        await state.update_data(stage_pending_file_id=file_id, stage_pending_file_kind=file_kind)

    await _nav_push(state, "staged", "file", idx)
    if spec.get("qcount"):
        await _finish_mcq_stage(message, state, spec, int(spec["qcount"]))
    else:
        await _ask_stage_qcount(message, state, spec)


async def _finish_mcq_stage(answerable, state: FSMContext, spec: dict, qcount: int) -> None:
    await state.update_data(stage_pending_qcount=qcount)
    await _ask_stage_answers(answerable, state, spec, qcount)


@router.message(StagedBuildState.collecting)
async def staged_build_text(message: Message, state: FSMContext) -> None:
    if (message.text or "") in TOP_MENU_BTNS:
        return
    if not message.text:
        return

    data = await state.get_data()
    field = data.get("stage_field")
    kind = data["stage_kind"]
    idx = data["stage_spec_idx"]
    spec = _stage_spec_list(kind)[idx]
    text = message.text.strip()

    if field == "name":
        name = sanitize_plain_text(text, max_len=80)
        if len(name) < 2:
            return await message.answer("⚠️ Fan nomi juda qisqa. Qayta kiriting:")
        await state.update_data(stage_current_name=name)
        await _nav_push(state, "staged", "name", idx)
        await _ask_stage_duration(message, state, spec)
        return

    if field == "duration":
        if not text.isdigit() or not (1 <= int(text) <= 300):
            return await message.answer("⚠️ Faqat 1..300 oralig'idagi raqam kiriting:")
        await state.update_data(stage_pending_duration=int(text))
        await _nav_push(state, "staged", "duration", idx)
        await _ask_stage_prompt_or_file(message, state, spec)
        return

    if field == "qcount":
        if not text.isdigit() or not (1 <= int(text) <= 300):
            return await message.answer("⚠️ Faqat 1..300 oralig'idagi raqam kiriting:")
        await _nav_push(state, "staged", "qcount", idx)
        await _finish_mcq_stage(message, state, spec, int(text))
        return

    if field == "answers":
        qcount = int((data.get("stage_pending_qcount")) or spec.get("qcount") or 0)
        answers = normalize_answers(text)
        if not is_answer_string_valid(answers):
            return await message.answer(
                "⚠️ Javob kaliti faqat A-F harflaridan iborat bo'lishi kerak. Qayta kiriting:")
        if len(answers) != qcount:
            return await message.answer(
                f"⚠️ Javoblar soni {qcount} ta bo'lishi kerak, siz {len(answers)} ta yubordingiz. Qayta kiriting:")

        section_data = {
            "name": data.get("stage_current_name", spec["default_label"]),
            "type": spec["type"],
            "duration_min": int(data.get("stage_pending_duration") or 0),
            "qcount": qcount,
            "answers": answers,
            "file_id": data.get("stage_pending_file_id"),
            "file_kind": data.get("stage_pending_file_kind"),
        }
        if data.get("stage_pending_question_file_id"):
            section_data["question_file_id"] = data.get("stage_pending_question_file_id")
        if spec["type"] == STAGE_TYPE_MCQ:
            section_data["weight"] = spec["weight"]

        open_positions = open_question_positions(answers)
        if open_positions:
            await state.update_data(
                stage_pending_section=section_data, stage_pending_open_positions=open_positions,
                stage_field="open_ref"
            )
            await _nav_push(state, "staged", "answers", idx)
            await message.answer(_open_ref_prompt_text(open_positions), parse_mode=ParseMode.HTML)
            return

        await _nav_push(state, "staged", "answers", idx)
        await _save_current_stage_and_advance(message, state, section_data)
        return

    if field == "open_ref":
        positions = data.get("stage_pending_open_positions") or []
        ref_map = parse_open_ref_answers(text, positions)
        if ref_map is None:
            await message.answer(
                f"⚠️ Qatorlar soni mos kelmadi — aynan <b>{len(positions)} qator</b> yuboring "
                f"(har birida bitta javob, savollar tartibida):",
                parse_mode=ParseMode.HTML
            )
            return
        section_data = data.get("stage_pending_section") or {}
        section_data["open_ref_answers"] = ref_map
        await _nav_push(state, "staged", "open_ref", idx)
        await _save_current_stage_and_advance(message, state, section_data)
        return

    if field == "prompt":
        if len(text) < 3:
            return await message.answer("⚠️ Matn juda qisqa. Qayta kiriting:")
        section_data = {
            "name": data.get("stage_current_name", spec["default_label"]),
            "type": spec["type"],
            "duration_min": int(data.get("stage_pending_duration") or 0),
            "prompt": text,
        }
        if GEMINI_API_KEY:
            await state.update_data(stage_pending_section=section_data)
            await _nav_push(state, "staged", "prompt", idx)
            kb = InlineKeyboardMarkup(inline_keyboard=[[
                InlineKeyboardButton(text="🤖 AI avtomatik tekshirsin", callback_data="stg_grader_ai"),
                InlineKeyboardButton(text="👤 O'zim/ustoz qo'lda", callback_data="stg_grader_manual"),
            ]])
            await message.answer(
                f"<b>{section_data['name']}</b> javoblarini kim baholasin?",
                parse_mode=ParseMode.HTML, reply_markup=kb
            )
            return
        section_data["grader"] = "manual"
        await _nav_push(state, "staged", "prompt", idx)
        await _save_current_stage_and_advance(message, state, section_data)
        return

    # ✅ Fayl kutilayotgan qadamda matn yuborilsa — avval bot jim qolardi.
    if field in ("file", "listening_pdf"):
        await message.answer("📄 Iltimos, savollar yozilgan <b>PDF faylni</b> yuboring.", parse_mode=ParseMode.HTML)
    elif field == "listening_media":
        await message.answer("🎧 Iltimos, <b>audio yoki video</b> faylni yuboring.", parse_mode=ParseMode.HTML)
    else:
        await message.answer("👆 Yuqoridagi savolga javob bering yoki tugmalardan birini tanlang.")


@router.callback_query(F.data.startswith("stg_grader_"))
async def staged_grader_choice(callback: CallbackQuery, state: FSMContext) -> None:
    choice = callback.data.split("_")[2]
    data = await state.get_data()
    section_data = data.get("stage_pending_section")
    if not section_data:
        await callback.answer()
        return
    section_data["grader"] = "ai" if choice == "ai" else "manual"
    await callback.answer("🤖 AI tanlandi" if choice == "ai" else "👤 Qo'lda tekshirish tanlandi")
    await _nav_push(state, "staged", "grader", data["stage_spec_idx"])
    await _save_current_stage_and_advance(callback.message, state, section_data)


async def _staged_nav_replay(answerable, state: FSMContext, entry: dict) -> bool:
    """Bosqichli sehrgarda "🔄 Boshidan boshlash" bosilganda — entry['idx']
    ko'rsatgan bosqichga qaytadi. Shu bosqichdan BOSHLAB (o'zi va undan
    keyingi barchasi) hali tasdiqlanmagan deb hisoblanadi — shu sabab ular
    uchun stage_cfg'da avval saqlangan (agar bo'lsa) ma'lumotlar tozalanadi.
    Aks holda foydalanuvchi masalan avval "Ha, bo'lsin" deb keyin fikridan
    qaytib "Yo'q"ni tanlasa, eski to'liq ma'lumot "arvoh" bo'lib yakuniy
    testda qolib ketishi mumkin edi."""
    data = await state.get_data()
    kind = data.get("stage_kind")
    idx = entry.get("idx")
    if kind is None or idx is None:
        return False
    spec_list = _stage_spec_list(kind)
    if idx >= len(spec_list):
        return False
    spec = spec_list[idx]

    stage_cfg = dict(data.get("stage_cfg") or {})
    stale_keys = {spec_list[i]["key"] for i in range(idx, len(spec_list))}
    stage_cfg = {k: v for k, v in stage_cfg.items() if k not in stale_keys}
    await state.update_data(stage_spec_idx=idx, stage_cfg=stage_cfg)

    node = entry["node"]
    progress = f"[{idx + 1}/{len(spec_list)}] "

    if node == "toggle":
        kb = InlineKeyboardMarkup(inline_keyboard=[[
            InlineKeyboardButton(text="✅ Ha, bo'lsin", callback_data="stg_on"),
            InlineKeyboardButton(text="➖ Yo'q, o'tkazib yuborish", callback_data="stg_off"),
        ]])
        await answerable.answer(
            f"{progress}<b>{spec['default_label']}</b> ushbu testda bo'lsinmi?",
            parse_mode=ParseMode.HTML, reply_markup=kb
        )
        return True

    if node == "name":
        await state.update_data(stage_field="name")
        await answerable.answer(
            f"{progress}<b>{spec['default_label']}</b>"
            + (f" — {spec['qcount']} ta savol, og'irligi {spec['weight']}." if spec.get("qcount") else ".")
            + "\nBu fanning nomini kiriting (masalan: <i>Fizika</i>):",
            parse_mode=ParseMode.HTML
        )
        await state.set_state(StagedBuildState.collecting)
        return True

    if node == "duration":
        await _ask_stage_duration(answerable, state, spec)
        return True

    if node in ("file", "prompt"):
        await _ask_stage_prompt_or_file(answerable, state, spec)
        return True

    if node == "qcount":
        await _ask_stage_qcount(answerable, state, spec)
        return True

    if node == "answers":
        fresh = await state.get_data()
        qcount = int(fresh.get("stage_pending_qcount") or spec.get("qcount") or 0)
        await _ask_stage_answers(answerable, state, spec, qcount)
        return True

    if node == "open_ref":
        fresh = await state.get_data()
        positions = fresh.get("stage_pending_open_positions") or []
        if not positions:
            return False
        await state.update_data(stage_field="open_ref")
        await answerable.answer(_open_ref_prompt_text(positions), parse_mode=ParseMode.HTML)
        await state.set_state(StagedBuildState.collecting)
        return True

    if node == "grader":
        fresh = await state.get_data()
        section_data = fresh.get("stage_pending_section")
        if not section_data:
            return False
        kb = InlineKeyboardMarkup(inline_keyboard=[[
            InlineKeyboardButton(text="🤖 AI avtomatik tekshirsin", callback_data="stg_grader_ai"),
            InlineKeyboardButton(text="👤 O'zim/ustoz qo'lda", callback_data="stg_grader_manual"),
        ]])
        await answerable.answer(
            f"<b>{section_data['name']}</b> javoblarini kim baholasin?",
            parse_mode=ParseMode.HTML, reply_markup=kb
        )
        return True

    return False


async def _finalize_staged_test(answerable, state: FSMContext) -> None:
    data = await state.get_data()
    kind = data["stage_kind"]
    stage_cfg = data.get("stage_cfg") or {}
    spec_list = _stage_spec_list(kind)
    order = [s["key"] for s in spec_list if s["key"] in stage_cfg]

    if not order:
        await answerable.answer(
            "⚠️ Kamida bitta bosqich yoqilishi kerak edi. Test yaratilmadi — "
            "\"➕ Test Yaratish\" orqali qaytadan boshlang."
        )
        await state.clear()
        return

    total_duration = sum(int(stage_cfg[k].get("duration_min") or 0) for k in order)
    total_qcount = sum(int(stage_cfg[k].get("qcount") or 0) for k in order)

    is_free = int(data.get("is_free") or 0)
    price = 0 if is_free else int(data.get("price") or 0)
    creator_tg_id = data.get("stage_creator_tg_id")

    skills_cfg_json = json.dumps({"kind": kind, "order": order, "sections": stage_cfg}, ensure_ascii=False)

    new_test_id = await POOL.fetchval("""
        INSERT INTO tests (
            code, title, subject, price, duration, questions_count, is_free,
            start_mode, start_ts, start_at,
            exam_type, skills_cfg, creator_id, created_ts
        )
        VALUES ($1, $2, $3, $4, $5, $6, $7, 'normal', 0, NULL, $8, $9, $10, $11)
        RETURNING id
    """,
        data.get("code"), data.get("title"), data.get("subject"), price,
        total_duration, total_qcount, is_free, kind, skills_cfg_json, creator_tg_id, now_ts()
    )

    section_lines = "\n".join(
        f"• <b>{stage_cfg[k]['name']}</b> ({STAGE_TYPE_LABELS.get(stage_cfg[k]['type'], stage_cfg[k]['type'])}, "
        f"{stage_cfg[k].get('duration_min', 0)} daq)"
        for k in order
    )
    kind_label = {"dtm_staged": "DTM (5 bosqich)", "ielts4": "IELTS (4 ko'nikma)",
                  "cefr4": "CEFR (4 ko'nikma)", "sat_staged": "SAT (2 bosqich)"}.get(kind, kind)
    await answerable.answer(
        f"✅ <b>{kind_label}</b> testi qo'shildi!\n🔑 Kod: <code>{data.get('code')}</code>\n\n"
        f"<b>Bosqichlar:</b>\n{section_lines}",
        parse_mode=ParseMode.HTML,
        reply_markup=main_menu_kb(await is_admin_user(creator_tg_id))
    )
    await state.clear()


STAGE_TYPE_LABELS = {
    STAGE_TYPE_MCQ: "test",
    STAGE_TYPE_MCQ_BAND: "test + band",
    STAGE_TYPE_TEXT: "matn/rasm, ustoz baholaydi",
    STAGE_TYPE_VOICE: "ovozli, ustoz baholaydi",
}


# ----- O'quvchi: bosqichlarni ketma-ket yechish -----
def stage_mcq_kb(current_q: int, qcount: int, answers_dict: dict, correct_key: str = "") -> InlineKeyboardMarkup:
    kb = InlineKeyboardMarkup(inline_keyboard=[])
    current_ans = answers_dict.get(str(current_q))

    if _is_open_question(correct_key, current_q):
        # ✅ Bu savol javob kalitida "x" — variantsiz (ochiq) deb belgilangan,
        # shuning uchun A-F o'rniga FAQAT matn yozish tugmasi ko'rsatiladi.
        mark = "✅ " if current_ans == "x" else "✍️ "
        kb.inline_keyboard.append(
            [InlineKeyboardButton(text=f"{mark}Matn bilan javob yozish", callback_data="sq_freetext")])
    else:
        row1, row2 = [], []
        for letter in ["A", "B", "C"]:
            mark = "🟢 " if current_ans == letter.lower() else ""
            row1.append(InlineKeyboardButton(text=f"{mark}{letter}", callback_data=f"sq_ans_{letter.lower()}"))
        for letter in ["D", "E", "F"]:
            mark = "🟢 " if current_ans == letter.lower() else ""
            row2.append(InlineKeyboardButton(text=f"{mark}{letter}", callback_data=f"sq_ans_{letter.lower()}"))
        kb.inline_keyboard.append(row1)
        kb.inline_keyboard.append(row2)

    nav = []
    if current_q > 1:
        nav.append(InlineKeyboardButton(text="⬅️ Oldingi", callback_data="sq_nav_prev"))
    nav.append(InlineKeyboardButton(text=f"🔄 {current_q}/{qcount}", callback_data="sq_nav_noop"))
    if current_q < qcount:
        nav.append(InlineKeyboardButton(text="Keyingi ➡️", callback_data="sq_nav_next"))
    kb.inline_keyboard.append(nav)

    answered = len(answers_dict)
    kb.inline_keyboard.append(
        [InlineKeyboardButton(text=f"✅ Bosqichni yakunlash ({answered}/{qcount})", callback_data="sq_submit")])
    return kb


async def start_staged_attempt(user_tg_id: int, test_id: int, answerable, bot: Bot, state: FSMContext,
                               allow_hidden: bool = False) -> None:
    user_db_id = await get_user_db_id(user_tg_id)
    if not user_db_id:
        await answerable.answer("Avval /start orqali ro'yxatdan o'ting.")
        return

    t = await POOL.fetchrow("SELECT skills_cfg, is_hidden, is_closed FROM tests WHERE id=$1", test_id)
    if not t:
        await answerable.answer("Test topilmadi.")
        return
    if int(t["is_hidden"] or 0) == 1 and not allow_hidden:
        await answerable.answer("⏸ Bu test vaqtincha to'xtatilgan.")
        return
    if int(t["is_closed"] or 0) == 1:
        await answerable.answer("🛑 Bu test yakunlangan.")
        return

    cfg = _load_skills_cfg(t["skills_cfg"])
    order = cfg.get("order") or []
    if not order:
        await answerable.answer("⚠️ Bu test hali to'liq sozlanmagan.")
        return

    # ✅ Endi bir (user, test) juftligi bo'yicha bir nechta urinishga ruxsat
    # beriladi — eng oxirgisi (ORDER BY id DESC) holatni belgilaydi. Avvalgi
    # urinish(lar) o'chirilmaydi, to'liq tarix sifatida saqlanadi.
    existing = await POOL.fetchrow(
        "SELECT id, status, section_idx, overall_grade FROM staged_attempts "
        "WHERE user_id=$1 AND test_id=$2 ORDER BY id DESC LIMIT 1",
        user_db_id, test_id
    )

    if existing:
        if existing["status"] == "in_progress":
            await state.update_data(
                staged_attempt_id=existing["id"], staged_test_id=test_id,
                staged_order=order, staged_pos=int(existing["section_idx"] or 0)
            )
            await state.set_state(StagedSolvingState.in_section)
            await _send_stage_section(user_tg_id, bot, state)
            return
        if existing["status"] == "pending_review":
            await answerable.answer("⏳ Siz bu testni allaqachon topshirgansiz — ustoz hozircha baholayapti.")
            return

        # status == 'graded' — qayta yechish taklif qilinadi, eski natija saqlanib qoladi
        kb = InlineKeyboardMarkup(inline_keyboard=[[
            InlineKeyboardButton(text="🔄 Ha, qaytadan boshlayman", callback_data=f"stagedretake_{test_id}"),
            InlineKeyboardButton(text="❌ Yo'q", callback_data="dismiss_msg"),
        ]])
        await answerable.answer(
            f"✅ Siz bu testni allaqachon topshirgansiz (natija: <b>{existing['overall_grade'] or '-'}</b>).\n\n"
            f"Qayta yechmoqchimisiz? Avvalgi natijangiz o'chmaydi, saqlanib qoladi — "
            f"ustoz ikkalasini ham ko'radi.",
            parse_mode=ParseMode.HTML,
            reply_markup=kb
        )
        return

    await _create_new_staged_attempt(user_tg_id, test_id, order, 1, answerable, bot, state)


async def _create_new_staged_attempt(
    user_tg_id: int, test_id: int, order: list, attempt_no: int, answerable, bot: Bot, state: FSMContext
) -> None:
    user_db_id = await get_user_db_id(user_tg_id)
    attempt_id = await POOL.fetchval("""
        INSERT INTO staged_attempts (
            user_id, test_id, status, section_idx, sections_json, attempt_no, started_at, created_ts
        )
        VALUES ($1, $2, 'in_progress', 0, '{}', $3, $4, $5)
        RETURNING id
    """, user_db_id, test_id, attempt_no, now_ts(), now_ts())

    await state.update_data(staged_attempt_id=attempt_id, staged_test_id=test_id, staged_order=order, staged_pos=0)
    await state.set_state(StagedSolvingState.in_section)
    await _send_stage_section(user_tg_id, bot, state)


@router.callback_query(F.data.startswith("stagedretake_"))
async def staged_retake_confirm(callback: CallbackQuery, state: FSMContext, bot: Bot) -> None:
    test_id = int(callback.data.split("_")[1])
    user_db_id = await get_user_db_id(callback.from_user.id)
    if not user_db_id:
        await callback.message.answer("Avval /start orqali ro'yxatdan o'ting.")
        return

    t = await POOL.fetchrow("SELECT skills_cfg, is_hidden, is_closed FROM tests WHERE id=$1", test_id)
    if not t or int(t["is_hidden"] or 0) == 1 or int(t["is_closed"] or 0) == 1:
        await callback.message.answer("⏸ Bu test hozircha mavjud emas.")
        return

    cfg = _load_skills_cfg(t["skills_cfg"])
    order = cfg.get("order") or []

    last_attempt_no = await POOL.fetchval(
        "SELECT COALESCE(MAX(attempt_no), 0) FROM staged_attempts WHERE user_id=$1 AND test_id=$2",
        user_db_id, test_id
    )
    await callback.answer()
    await _create_new_staged_attempt(
        callback.from_user.id, test_id, order, int(last_attempt_no) + 1, callback.message, bot, state
    )


async def _send_stage_section(user_tg_id: int, bot: Bot, state: FSMContext) -> None:
    data = await state.get_data()
    test_id = data["staged_test_id"]
    order = data["staged_order"]
    pos = data["staged_pos"]
    user_db_id = await get_user_db_id(user_tg_id)

    t = await POOL.fetchrow("SELECT skills_cfg FROM tests WHERE id=$1", test_id)
    cfg = _load_skills_cfg(t["skills_cfg"])
    key = order[pos]
    section = cfg["sections"][key]
    label = section["name"]
    duration_min = int(section.get("duration_min") or 0)
    stype = section["type"]

    # ✅ Bu bosqich AVVAL ham boshlanganmi (masalan foydalanuvchi "🏠 Bosh
    # menyu"ga chiqmasdan shu testga qaytib "Boshlash"ni bosdi, yoki server
    # qayta ishga tushgandan keyin ham holat saqlanib qolgani uchun) — agar
    # shunday bo'lsa, taymerni QAYTA BOSHLAMAYMIZ va MCQ javoblarini
    # O'CHIRIB YUBORMAYMIZ (aks holda foydalanuvchi "Boshlash"ni qayta bosib
    # vaqtni cheksiz uzaytirishi yoki belgilagan javoblarini yo'qotishi
    # mumkin edi). Yangi bosqichga o'tishda esa (kalit farq qiladi) hammasi
    # toza boshlanadi.
    prev_key = data.get("stage_current_key")
    existing_session = await get_session(user_db_id, test_id, f"stage_{key}")
    is_resume = existing_session is not None and prev_key == key

    await state.update_data(stage_current_key=key, stage_current_type=stype)
    cancel_timer(user_tg_id, test_id)

    if existing_session:
        start_ts_ = int(existing_session["start_ts"])
        duration_sec = int(existing_session["duration_sec"])
    else:
        start_ts_ = now_ts()
        duration_sec = duration_min * 60
        await upsert_session(user_db_id, test_id, f"stage_{key}", start_ts_, duration_sec)
    end_ts_ = start_ts_ + duration_sec

    if now_ts() > end_ts_ + 60:
        # ✅ Bu bosqichning vaqti (masalan bot uzoq vaqt o'chiq turgani
        # sababli) allaqachon tugab bo'lgan — endi kutish ma'nosiz, mavjud
        # javoblar bilan shu bosqichni avtomatik yakunlaymiz.
        await bot.send_message(
            user_tg_id,
            f"⌛ <b>{label}</b> vaqti tugagan edi — mavjud javoblaringiz bilan yakunlanmoqda.",
            parse_mode=ParseMode.HTML
        )
        await delete_session(user_db_id, test_id, f"stage_{key}")
        if stype in (STAGE_TYPE_MCQ, STAGE_TYPE_MCQ_BAND):
            qcount = int(section.get("qcount") or data.get("stage_current_qcount") or 0)
            answers = data.get("stage_mcq_answers", {}) if is_resume else {}
            user_ans = "".join(answers.get(str(i), "*") for i in range(1, qcount + 1))
            await _complete_current_stage_mcq(user_tg_id, bot, state, user_ans)
        else:
            result = {"text": None, "photo_id": None, "voice_id": None, "band": None, "feedback": None}
            await _save_stage_result_and_advance(user_tg_id, bot, state, key, result)
        return

    intro = f"➡️ <b>{label}</b> ({pos + 1}/{len(order)}-bosqich)\n⏳ Vaqt: {duration_min} daqiqa"

    if stype in (STAGE_TYPE_MCQ, STAGE_TYPE_MCQ_BAND):
        qcount = int(section["qcount"])
        answer_key = normalize_answers(section.get("answers") or "")
        await state.update_data(stage_answer_key=answer_key)

        if is_resume:
            current_q = data.get("stage_mcq_q", 1)
            answers = data.get("stage_mcq_answers", {})
            kb_msg = await bot.send_message(
                user_tg_id,
                f"▶️ <b>{label}</b> bosqichini davom ettiryapsiz — avvalgi javoblaringiz saqlangan.\n\n"
                f"{question_status_note(answer_key, current_q)}"
                f"📝 <b>{current_q}-savol</b> / {qcount}",
                parse_mode=ParseMode.HTML,
                reply_markup=stage_mcq_kb(current_q, qcount, answers, answer_key)
            )
            await state.update_data(stage_kb_mid=kb_msg.message_id)
        else:
            # ✅ stage_mcq_open_answers ham tozalanadi — avval oldingi bosqichning
            # ochiq javoblari keyingi bosqich natijasiga "oqib" o'tib qolardi.
            await state.update_data(stage_current_qcount=qcount, stage_mcq_q=1, stage_mcq_answers={},
                                    stage_mcq_open_answers={})

            if section.get("file_id"):
                caption = f"{label} — {qcount} ta savol"
                fk = section.get("file_kind")
                try:
                    if fk == "audio":
                        await bot.send_audio(user_tg_id, section["file_id"], caption=caption)
                    elif fk == "video":
                        await bot.send_video(user_tg_id, section["file_id"], caption=caption)
                    else:
                        await bot.send_document(user_tg_id, section["file_id"], caption=caption)
                except Exception:
                    logging.exception(f"Bosqich faylini yuborib bo'lmadi (test={test_id}, key={key})")

            if section.get("question_file_id"):
                try:
                    await bot.send_document(
                        user_tg_id, section["question_file_id"],
                        caption=f"{label} — savollar"
                    )
                except Exception:
                    logging.exception(f"Bosqich savollar faylini yuborib bo'lmadi (test={test_id}, key={key})")

            kb_msg = await bot.send_message(
                user_tg_id,
                f"📝 {question_status_note(answer_key, 1)}<b>1-savol</b> / {qcount}\n"
                f"Pastdagi tugmalar orqali javobni belgilang:",
                parse_mode=ParseMode.HTML,
                reply_markup=stage_mcq_kb(1, qcount, {}, answer_key)
            )
            await state.update_data(stage_kb_mid=kb_msg.message_id)

        timer_msg = await bot.send_message(user_tg_id, intro + "\n⏳ Qolgan vaqt: ...", parse_mode=ParseMode.HTML)
        ACTIVE_TIMERS[(user_tg_id, test_id)] = asyncio.create_task(
            start_countdown(bot, user_tg_id, test_id, end_ts_, timer_msg.message_id)
        )
    else:
        timer_msg = await bot.send_message(user_tg_id, intro + "\n⏳ Qolgan vaqt: ...", parse_mode=ParseMode.HTML)
        ACTIVE_TIMERS[(user_tg_id, test_id)] = asyncio.create_task(
            start_countdown(bot, user_tg_id, test_id, end_ts_, timer_msg.message_id)
        )
        prompt_text = section.get("prompt", "")
        if stype == STAGE_TYPE_TEXT:
            await bot.send_message(
                user_tg_id,
                f"✍️ <b>Topshiriq:</b>\n{_h(prompt_text)}\n\n"
                f"Javobingizni <b>matn</b> yoki <b>rasm</b> (masalan qo'lda yozilgan esse fotosi) "
                f"qilib yuboring.",
                parse_mode=ParseMode.HTML
            )
        else:
            await bot.send_message(
                user_tg_id,
                f"🎙 <b>Mavzu:</b>\n{_h(prompt_text)}\n\n"
                f"Javobingizni <b>ovozli xabar</b> qilib yuboring.",
                parse_mode=ParseMode.HTML
            )


@router.callback_query(StateFilter(StagedSolvingState.in_section, StagedSolvingState.awaiting_free_text),
                       F.data.startswith("sq_"))
async def staged_mcq_handler(callback: CallbackQuery, state: FSMContext, bot: Bot) -> None:
    data = await state.get_data()
    # ✅ Faqat AMALDAGI bosqich klaviaturasi qabul qilinadi. Avval oldingi
    # bosqich xabaridagi tugma (yoki "Bosqichni yakunlash"ning ikkinchi
    # bosilishi) KEYINGI bosqichga tegib, uni javobsiz topshirib yuborardi.
    kb_mid = data.get("stage_kb_mid")
    if kb_mid and callback.message and callback.message.message_id != kb_mid:
        await callback.answer("⏱ Bu tugma oldingi bosqichga tegishli — eng so'nggi xabardagi tugmalardan "
                              "foydalaning.", show_alert=True)
        return
    if callback.data != "sq_freetext" and await state.get_state() == StagedSolvingState.awaiting_free_text.state:
        await state.set_state(StagedSolvingState.in_section)
    current_q = data.get("stage_mcq_q", 1)
    qcount = int(data.get("stage_current_qcount", 0))
    answers = data.get("stage_mcq_answers", {})
    correct_key = data.get("stage_answer_key", "")
    action = callback.data

    if action == "sq_nav_noop":
        await callback.answer(f"Siz hozir {current_q}-savoldasiz")
        return

    if action.startswith("sq_ans_"):
        letter = action.split("_")[2]
        answers[str(current_q)] = letter
        await state.update_data(stage_mcq_answers=answers)
        try:
            await callback.message.edit_reply_markup(
                reply_markup=stage_mcq_kb(current_q, qcount, answers, correct_key))
        except Exception:
            pass
        await callback.answer(f"✅ {letter.upper()} belgilandi")
        return

    if action == "sq_nav_prev":
        if current_q > 1:
            current_q -= 1
            await state.update_data(stage_mcq_q=current_q)
            try:
                await callback.message.edit_text(
                    text=f"📝 {question_status_note(correct_key, current_q)}<b>{current_q}-savol</b> / {qcount}\n"
                         f"Pastdagi tugmalar orqali javobni belgilang:",
                    parse_mode=ParseMode.HTML,
                    reply_markup=stage_mcq_kb(current_q, qcount, answers, correct_key)
                )
            except Exception:
                pass
        await callback.answer(f"⬅️ {current_q}-savolga o'tdingiz")
        return

    if action == "sq_nav_next":
        if current_q < qcount:
            current_q += 1
            await state.update_data(stage_mcq_q=current_q)
            try:
                await callback.message.edit_text(
                    text=f"📝 {question_status_note(correct_key, current_q)}<b>{current_q}-savol</b> / {qcount}\n"
                         f"Pastdagi tugmalar orqali javobni belgilang:",
                    parse_mode=ParseMode.HTML,
                    reply_markup=stage_mcq_kb(current_q, qcount, answers, correct_key)
                )
            except Exception:
                pass
        await callback.answer(f"➡️ {current_q}-savolga o'tdingiz")
        return

    if action == "sq_freetext":
        await callback.answer()
        await callback.message.answer(
            f"✏️ <b>{current_q}-savol</b> uchun javobingizni matn qilib yozing "
            f"(variantsiz/ochiq savol — ustoz keyinroq tekshiradi):",
            parse_mode=ParseMode.HTML
        )
        await state.set_state(StagedSolvingState.awaiting_free_text)
        return

    if action == "sq_submit":
        user_ans = "".join(answers.get(str(i), "*") for i in range(1, qcount + 1))
        try:
            await callback.message.delete_reply_markup()
        except Exception:
            pass
        await callback.answer("Qabul qilindi, hisoblanmoqda...")
        await _complete_current_stage_mcq(callback.from_user.id, bot, state, user_ans)


@router.message(StagedSolvingState.awaiting_free_text)
async def staged_collect_free_text(message: Message, state: FSMContext) -> None:
    if (message.text or "") in TOP_MENU_BTNS:
        return

    photo_id = message.photo[-1].file_id if message.photo else None
    text_val = sanitize_plain_text(message.text or message.caption or "", max_len=500)
    if not photo_id and len(text_val) < 1:
        await message.answer("⚠️ Javob bo'sh bo'lmasin. Matn yozing yoki rasm yuboring:")
        return

    data = await state.get_data()
    current_q = data.get("stage_mcq_q", 1)
    qcount = int(data.get("stage_current_qcount", 0))
    answers = data.get("stage_mcq_answers", {})
    open_answers = data.get("stage_mcq_open_answers", {})
    correct_key = data.get("stage_answer_key", "")

    answers[str(current_q)] = "x"
    open_answers[str(current_q)] = {"text": text_val or None, "photo_id": photo_id}
    await state.update_data(stage_mcq_answers=answers, stage_mcq_open_answers=open_answers)
    await state.set_state(StagedSolvingState.in_section)

    ack = "rasm" if photo_id else "matn"
    kb_msg = await message.answer(
        f"✅ {current_q}-savol uchun {ack} javobingiz qabul qilindi.\n\n"
        f"📝 <b>{current_q}-savol</b> / {qcount}\nPastdagi tugmalar orqali davom eting:",
        parse_mode=ParseMode.HTML,
        reply_markup=stage_mcq_kb(current_q, qcount, answers, correct_key)
    )
    await state.update_data(stage_kb_mid=kb_msg.message_id)


def _build_stage_mcq_result(cfg: dict, key: str, user_ans: str, open_answers: dict) -> dict:
    """Bosqichli testning variantli bo'limi natijasi (chat ham, Mini App ham
    aynan shu funksiya orqali hisoblaydi — natija har ikkalasida bir xil)."""
    section = cfg["sections"][key]
    qcount = int(section["qcount"])
    correct_key = normalize_answers(section["answers"])

    qn = min(qcount, len(correct_key), len(user_ans))
    # ✅ Ochiq (x) savollar kalit bilan avtomatik tekshiriladi; faqat rasm
    # ko'rinishidagi yoki kalitsiz javoblar ustoz tekshiruviga qoladi.
    open_answers = open_answers or {}
    open_grades = initial_open_grades(correct_key[:qn], open_answers, section.get("open_ref_answers") or {})
    score, scored_total, open_pending = score_with_open(user_ans, correct_key, qn, open_grades)
    result = {"score": score, "total": scored_total, "user_answers": user_ans[:qn],
              "open_grades": open_grades, "open_pending": open_pending}
    if open_answers:
        result["open_answers"] = open_answers

    if section["type"] == STAGE_TYPE_MCQ_BAND:
        result["band"] = _stage_band_for_mcq(cfg["kind"], key, score, scored_total)
    else:
        weight = float(section.get("weight") or 1.0)
        result["weight"] = weight
        result["earned"] = score * weight
        result["max"] = scored_total * weight
    return result


async def _complete_current_stage_mcq(user_tg_id: int, bot: Bot, state: FSMContext, user_ans: str) -> None:
    data = await state.get_data()
    test_id = data["staged_test_id"]
    order = data["staged_order"]
    pos = data["staged_pos"]
    user_db_id = await get_user_db_id(user_tg_id)

    t = await POOL.fetchrow("SELECT skills_cfg FROM tests WHERE id=$1", test_id)
    cfg = _load_skills_cfg(t["skills_cfg"])
    key = order[pos]
    result = _build_stage_mcq_result(cfg, key, user_ans, data.get("stage_mcq_open_answers") or {})

    await delete_session(user_db_id, test_id, f"stage_{key}")
    cancel_timer(user_tg_id, test_id)

    await _save_stage_result_and_advance(user_tg_id, bot, state, key, result)


@router.message(StagedSolvingState.in_section, F.text | F.photo)
async def staged_open_text_handler(message: Message, state: FSMContext, bot: Bot) -> None:
    if (message.text or "") in TOP_MENU_BTNS:
        return
    data = await state.get_data()
    stype = data.get("stage_current_type")
    if stype != STAGE_TYPE_TEXT:
        # ✅ Avval bu yerda bot jim qolardi — o'quvchi nima qilishini bilmasdi.
        if stype == STAGE_TYPE_VOICE:
            await message.answer("🎙 Bu bosqichda javobni <b>ovozli xabar</b> qilib yuboring.",
                                 parse_mode=ParseMode.HTML)
        else:
            await message.answer("👆 Bu bosqichda javoblarni yuqoridagi tugmalar (A–F) orqali belgilang, "
                                 "so'ng «✅ Bosqichni yakunlash»ni bosing.")
        return

    text_val = sanitize_plain_text(message.text or message.caption or "", max_len=4000)
    photo_id = message.photo[-1].file_id if message.photo else None
    if not text_val and not photo_id:
        await message.answer("⚠️ Matn yoki rasm yuboring.")
        return

    key = data["stage_current_key"]
    test_id = data["staged_test_id"]
    user_db_id = await get_user_db_id(message.from_user.id)
    await delete_session(user_db_id, test_id, f"stage_{key}")
    cancel_timer(message.from_user.id, test_id)

    result = {"text": text_val or None, "photo_id": photo_id, "band": None, "feedback": None}
    await _save_stage_result_and_advance(message.from_user.id, bot, state, key, result)


@router.message(StagedSolvingState.in_section, F.voice)
async def staged_open_voice_handler(message: Message, state: FSMContext, bot: Bot) -> None:
    data = await state.get_data()
    stype = data.get("stage_current_type")
    if stype != STAGE_TYPE_VOICE:
        if stype == STAGE_TYPE_TEXT:
            await message.answer("✍️ Bu bosqichda javobni <b>matn</b> yoki <b>rasm</b> qilib yuboring "
                                 "(ovozli xabar emas).", parse_mode=ParseMode.HTML)
        else:
            await message.answer("👆 Bu bosqichda javoblarni yuqoridagi tugmalar (A–F) orqali belgilang.")
        return

    key = data["stage_current_key"]
    test_id = data["staged_test_id"]
    user_db_id = await get_user_db_id(message.from_user.id)
    await delete_session(user_db_id, test_id, f"stage_{key}")
    cancel_timer(message.from_user.id, test_id)

    result = {"voice_id": message.voice.file_id, "band": None, "feedback": None}
    await _save_stage_result_and_advance(message.from_user.id, bot, state, key, result)


async def _save_stage_result_and_advance(
    user_tg_id: int, bot: Bot, state: FSMContext, key: str, result: dict
) -> None:
    data = await state.get_data()
    attempt_id = data["staged_attempt_id"]
    test_id = data["staged_test_id"]
    order = data["staged_order"]
    pos = data["staged_pos"] + 1

    row = await POOL.fetchrow("SELECT sections_json FROM staged_attempts WHERE id=$1", attempt_id)
    sections = json.loads(row["sections_json"]) if row and row["sections_json"] else {}
    sections[key] = result
    await POOL.execute(
        "UPDATE staged_attempts SET sections_json=$1, section_idx=$2 WHERE id=$3",
        json.dumps(sections, ensure_ascii=False), pos, attempt_id
    )

    if pos < len(order):
        await state.update_data(staged_pos=pos)
        await _send_stage_section(user_tg_id, bot, state)
    else:
        await _finalize_staged_attempt(user_tg_id, bot, state, attempt_id, test_id, sections)


async def _complete_dtm_staged_attempt(bot: Bot, attempt_id: int, graded_by: int | None = None) -> None:
    """DTM (5 bosqich) urinishini yakunlaydi: umumiy ball, natija (results),
    o'quvchiga xabar va sertifikat. Ochiq javoblar ustoz tomonidan
    tekshirilgandan keyin ham shu funksiya chaqiriladi."""
    a = await POOL.fetchrow("SELECT user_id, test_id, sections_json FROM staged_attempts WHERE id=$1", attempt_id)
    if not a:
        return
    test_id, user_db_id = a["test_id"], a["user_id"]
    sections = _load_json_dict(a["sections_json"])
    t = await POOL.fetchrow("""
        SELECT skills_cfg, title, code, subject, cert_academy_name, cert_signee_name, cert_template, creator_id
        FROM tests WHERE id=$1
    """, test_id)
    if not t:
        return
    cfg = _load_skills_cfg(t["skills_cfg"])
    user_tg_id = await POOL.fetchval("SELECT telegram_id FROM users WHERE id=$1", user_db_id)
    admin_flag = await is_admin_user(user_tg_id)

    earned = sum(float(sections[k].get("earned") or 0) for k in cfg["order"] if k in sections)
    maxp = sum(float(sections[k].get("max") or 0) for k in cfg["order"] if k in sections)
    percent = (earned / maxp * 100.0) if maxp > 0 else 0.0
    raw_score = sum(int(sections[k].get("score") or 0) for k in cfg["order"] if k in sections)
    raw_total = sum(int(sections[k].get("total") or 0) for k in cfg["order"] if k in sections)

    await POOL.execute(
        "UPDATE staged_attempts SET status='graded', overall_score=$1, finished_at=COALESCE(finished_at, $2), "
        "graded_by=$3, graded_at=$2 WHERE id=$4",
        percent, now_ts(), graded_by, attempt_id
    )

    # ✅ Qayta yechish mumkin bo'lgani uchun "results" jadvalidagi yagona
    # qatorni ENG SO'NGGI urinish bilan yangilaymiz (upsert) — reyting,
    # Excel eksport va h.k. har doim eng oxirgi natijani ko'radi. To'liq
    # urinishlar tarixi esa staged_attempts'da saqlanib qoladi.
    await POOL.execute("""
        INSERT INTO results (
            user_id, test_id, score, total_questions, percent,
            earned_points, max_points, details, finished_ts, date
        )
        VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10)
        ON CONFLICT (user_id, test_id) DO UPDATE SET
            score=EXCLUDED.score, total_questions=EXCLUDED.total_questions,
            percent=EXCLUDED.percent, earned_points=EXCLUDED.earned_points,
            max_points=EXCLUDED.max_points, details=EXCLUDED.details,
            finished_ts=EXCLUDED.finished_ts, date=EXCLUDED.date,
            certificate_sent=0, certificate_file_id=NULL, certificate_sent_ts=0
    """,
        user_db_id, test_id, raw_score, raw_total, float(percent),
        float(earned), float(maxp),
        json.dumps({"stages": sections}, ensure_ascii=False),
        now_ts(), now_str_local()
    )

    lines = "\n".join(
        f"• {cfg['sections'][k]['name']}: {sections[k].get('score', 0)}/{sections[k].get('total', 0)} "
        f"({sections[k].get('earned', 0):.1f}/{sections[k].get('max', 0):.1f} ball)"
        for k in cfg["order"] if k in sections
    )
    # ✅ Ustoz baholaganda chaqiriladi — o'quvchi botni bloklagan bo'lsa,
    # avval ustozning tugmasi "Kutilmagan xatolik" bilan yiqilardi.
    try:
        await bot.send_message(
            user_tg_id,
            f"🏁 <b>DTM (5 bosqich) yakunlandi!</b>\n\n{lines}\n\n"
            f"🏆 <b>Umumiy ball:</b> {earned:.1f}/{maxp:.1f} ({percent:.1f}%)",
            parse_mode=ParseMode.HTML,
            reply_markup=main_menu_kb(admin_flag)
        )
    except (TelegramForbiddenError, TelegramBadRequest):
        logging.warning(f"DTM yakuniy natijasi o'quvchiga yetmadi (attempt={attempt_id})")
        return

    if percent >= CERT_MIN_PERCENT:
        try:
            cert_academy, cert_signee, cert_tpl = await _resolve_cert_settings(
                t["cert_academy_name"], t["cert_signee_name"], t["cert_template"], t["creator_id"])
            user_full_name = await get_user_full_name(user_db_id)
            tmp_dir = tempfile.gettempdir()
            safe_code = re.sub(r"[^A-Z0-9\-_]+", "_", str(t["code"]).upper())
            pdf_path = os.path.join(tmp_dir, f"cert_{safe_code}_{user_db_id}.pdf")
            make_certificate_pdf(
                full_name=user_full_name, subject=t["subject"] or "DTM",
                test_code=t["code"], test_title=t["title"], exam_date_str=now_str_local(),
                percent=float(percent), correct_count=raw_score, total_questions=raw_total,
                out_path=pdf_path, academy_name=cert_academy, signee_name=cert_signee, template=cert_tpl,
            )
            sent = await bot.send_document(user_tg_id, FSInputFile(pdf_path), caption="🎁 Sertifikatingiz tayyor!")
            cert_file_id = sent.document.file_id if sent and sent.document else None
            await POOL.execute(
                "UPDATE results SET certificate_sent=1, certificate_file_id=$1, certificate_sent_ts=$2 "
                "WHERE user_id=$3 AND test_id=$4",
                cert_file_id, now_ts(), user_db_id, test_id
            )
        except Exception:
            logging.exception(f"DTM bosqichli sertifikat yuborilmadi (user={user_db_id}, test={test_id})")


async def _finalize_staged_attempt(
    user_tg_id: int, bot: Bot, state: FSMContext, attempt_id: int, test_id: int, sections: dict
) -> None:
    await state.clear()
    t = await POOL.fetchrow("""
        SELECT skills_cfg, title, code, subject, cert_academy_name, cert_signee_name, cert_template, creator_id
        FROM tests WHERE id=$1
    """, test_id)
    cfg = _load_skills_cfg(t["skills_cfg"])
    kind = cfg["kind"]
    user_db_id = await get_user_db_id(user_tg_id)
    admin_flag = await is_admin_user(user_tg_id)

    open_pending_total = sum(int(sections.get(k, {}).get("open_pending") or 0) for k in cfg["order"])

    if kind == "dtm_staged":
        if open_pending_total > 0:
            # ✅ Rasm ko'rinishidagi ochiq javoblar bor — ustoz tekshirmaguncha
            # yakuniy ball e'lon qilinmaydi.
            await POOL.execute(
                "UPDATE staged_attempts SET status='pending_review', finished_at=$1 WHERE id=$2",
                now_ts(), attempt_id
            )
            await bot.send_message(
                user_tg_id,
                "🏁 <b>Barcha bosqichlar yuborildi!</b>\n\n"
                "✍️ Rasm ko'rinishida yuborilgan ochiq javoblaringizni ustoz tekshirgach, yakuniy ball "
                "(va loyiq bo'lsa sertifikat) shu yerga yuboriladi.",
                parse_mode=ParseMode.HTML, reply_markup=main_menu_kb(admin_flag)
            )
            # ✅ Ustoz botni bloklagan bo'lsa ham o'quvchining oqimi buzilmaydi —
            # urinish "✍️ Tekshirilmagan ochiq javoblar" panelida ko'rinib turadi.
            try:
                await _send_staged_open_for_review(bot, t["creator_id"] or ADMIN_ID, attempt_id)
            except (TelegramForbiddenError, TelegramBadRequest):
                logging.warning(f"Ochiq javoblarni ustozga yuborib bo'lmadi (attempt={attempt_id})")
            return
        await _complete_dtm_staged_attempt(bot, attempt_id)
        return

    # ---- IELTS4: avval AI (agar shu bo'lim uchun yoqilgan bo'lsa) baholab
    # ko'radi, faqat qolganini ustozga yuboradi ----
    order = cfg["order"]
    section_cfg = cfg["sections"]
    ai_graded_any = False
    ai_failed_sections = []  # [(label, reason_key), ...] — ustozga SABABINI aytish uchun

    # ✅ AI (Gemini) so'rovi bir necha soniya (ba'zan sekinroq) davom etishi
    # mumkin, shu payt o'quvchiga hech narsa yuborilmasa "bot qotib qoldi"
    # degan taassurot qoldirardi. Shuning uchun AI chaqirilishidan OLDIN,
    # kutish kerakligini ANIQ aytamiz.
    pending_ai = any(
        section_cfg.get(k, {}).get("type") in (STAGE_TYPE_TEXT, STAGE_TYPE_VOICE)
        and section_cfg.get(k, {}).get("grader") == "ai"
        and sections.get(k, {}).get("band") is None
        for k in order
    )
    if pending_ai:
        try:
            await bot.send_message(
                user_tg_id,
                "⏳ <b>Barcha javoblaringiz qabul qilindi!</b>\n\n"
                "🤖 Writing/Speaking javobingizni sun'iy intellekt tekshirmoqda — "
                "bu odatda bir necha soniya davom etadi.\n\n"
                "<b>Iltimos, biroz sabr qiling</b> — hech narsa qotib qolgani yo'q, "
                "natijangiz tayyor bo'lishi bilan shu yerga keladi.",
                parse_mode=ParseMode.HTML
            )
        except Exception:
            logging.exception(f"AI kutish xabarini yuborib bo'lmadi (user_tg_id={user_tg_id})")

    for k in order:
        sec_cfg = section_cfg.get(k, {})
        if sec_cfg.get("type") not in (STAGE_TYPE_TEXT, STAGE_TYPE_VOICE) or sec_cfg.get("grader") != "ai":
            continue
        v = sections.get(k, {})
        if v.get("band") is not None:
            continue

        ai_scale = _stage_ai_scale(kind)
        if sec_cfg["type"] == STAGE_TYPE_TEXT and v.get("text"):
            result = await _gemini_grade(sec_cfg.get("prompt", ""), essay_text=v["text"], skill="writing", scale=ai_scale)
        elif sec_cfg["type"] == STAGE_TYPE_VOICE and v.get("voice_id"):
            result = await _gemini_grade(
                sec_cfg.get("prompt", ""), voice_file_id=v["voice_id"], bot=bot, skill="speaking", scale=ai_scale)
        elif sec_cfg["type"] == STAGE_TYPE_TEXT and v.get("photo_id"):
            # ✅ Writing rasm (fotosurat) qilib yuborilgan bo'lsa, AI hozircha
            # o'qiy olmaydi — bu ham "muvaffaqiyatsiz" deb hisoblanadi, shunda
            # ustozga sababi aniq aytiladi (avval sababsiz jim qo'lda qolardi).
            result = {"ok": False, "reason": "photo_unsupported"}
        else:
            continue

        if result["ok"]:
            v["band"] = result["band"]
            v["feedback"] = result["feedback"]
            v["graded_by"] = "ai"
            sections[k] = v
            ai_graded_any = True
        else:
            v["ai_failed_reason"] = result["reason"]
            sections[k] = v
            ai_failed_sections.append((sec_cfg["name"], result["reason"]))
            if result["reason"] == "quota":
                await _maybe_alert_admin_ai_quota(bot)

    if ai_graded_any or ai_failed_sections:
        await POOL.execute(
            "UPDATE staged_attempts SET sections_json=$1 WHERE id=$2",
            json.dumps(sections, ensure_ascii=False), attempt_id
        )

    needs_manual = open_pending_total > 0 or any(
        section_cfg.get(k, {}).get("type") in (STAGE_TYPE_TEXT, STAGE_TYPE_VOICE)
        and sections.get(k, {}).get("band") is None
        for k in order
    )

    if not needs_manual:
        # ✅ Hamma narsa (AI yoki avtomatik band jadvali orqali) allaqachon
        # baholangan — ustozni bezovta qilmasdan darhol yakunlaymiz.
        overall_band = await _apply_staged_grading_result(
            bot, attempt_id, test_id, user_db_id, sections, order, section_cfg, graded_by=None, kind=kind
        )
        await bot.send_message(
            user_tg_id,
            "🏁 <b>Test yakunlandi — barchasi avtomatik baholandi!</b>\n"
            "Batafsil natija va fikr-mulohaza yuqorida yuborildi.",
            parse_mode=ParseMode.HTML,
            reply_markup=main_menu_kb(admin_flag)
        )
        return

    await POOL.execute(
        "UPDATE staged_attempts SET status='pending_review', finished_at=$1 WHERE id=$2",
        now_ts(), attempt_id
    )
    await bot.send_message(
        user_tg_id,
        "🏁 <b>Barcha qismlar yuborildi!</b>\n\n"
        "📖 Reading/🎧 Listening natijasi avtomatik hisoblandi, "
        "✍️ Writing va/yoki 🎙 Speaking'ni ustoz tekshiradi. Baholangach, umumiy natija "
        "va fikr-mulohaza shu yerga yuboriladi.",
        parse_mode=ParseMode.HTML,
        reply_markup=main_menu_kb(admin_flag)
    )

    creator_id = t["creator_id"] or ADMIN_ID
    student = await POOL.fetchrow("SELECT full_name FROM users WHERE id=$1", user_db_id)

    summary_lines = []
    for k in order:
        v = sections.get(k, {})
        label = section_cfg[k]["name"]
        stype = section_cfg[k]["type"]
        ai_fail_reason = v.get("ai_failed_reason")
        if stype in (STAGE_TYPE_MCQ_BAND, STAGE_TYPE_MCQ):
            line = f"• {label}: {v.get('score', 0)}/{v.get('total', 0)} — band {v.get('band', '-')}"
            if int(v.get("open_pending") or 0):
                line += f" (+{v['open_pending']} ta rasmli ochiq javob — tekshiring)"
            summary_lines.append(line)
        elif v.get("band") is not None:
            summary_lines.append(f"• {label}: 🤖 AI baholadi — band {v['band']} (qayta ko'rib chiqish shart emas)")
        elif ai_fail_reason:
            reason_text = AI_FAIL_REASON_LABELS.get(ai_fail_reason, ai_fail_reason)
            summary_lines.append(
                f"• {label}: ⚠️ AI baholay olmadi ({reason_text}) — <b>hozircha qo'lda tekshiring</b>"
            )
        elif stype == STAGE_TYPE_TEXT:
            kind_txt = "rasm" if v.get("photo_id") else f"{len(v.get('text') or '')} belgi"
            summary_lines.append(f"• {label}: matn/rasm yuborilgan ({kind_txt}) — qo'lda tekshiring")
        else:
            summary_lines.append(f"• {label}: ovozli javob yuborilgan — qo'lda tekshiring")

    ai_notice = ""
    if ai_failed_sections:
        ai_notice = (
            "\n\n⚠️ <b>Diqqat:</b> yuqoridagi bo'lim(lar) uchun AI avtomatik baholash yoqilgan edi, "
            "lekin hozir ishlamadi. AI qayta ishlay boshlaguncha (masalan, limit tiklanguncha) "
            "shu turdagi topshiriqlarni <b>qo'lda</b> tekshirib turing."
        )

    kb = InlineKeyboardMarkup(inline_keyboard=[[
        InlineKeyboardButton(text="✅ Baholashni boshlash", callback_data=f"skg_start_{attempt_id}")
    ]])
    kind_title = "CEFR" if kind == "cefr4" else "IELTS"
    try:
        await bot.send_message(
            creator_id,
            f"🎓 <b>Yangi {kind_title} topshiriq — baholash kerak</b>\n\n"
            f"👤 {student['full_name'] if student else 'Noma’lum'}\n"
            f"📚 {t['title']} (<code>{t['code']}</code>)\n\n" + "\n".join(summary_lines) + ai_notice,
            parse_mode=ParseMode.HTML,
            reply_markup=kb
        )
    except Exception:
        logging.exception(f"Baholash bildirishnomasini yuborib bo'lmadi (creator={creator_id})")


# ----- AI (Gemini) orqali avtomatik baholash -----
def _extract_json_object(text: str) -> dict | None:
    match = re.search(r"\{.*\}", text or "", re.DOTALL)
    if not match:
        return None
    try:
        return json.loads(match.group(0))
    except json.JSONDecodeError:
        return None


# ✅ AI baholash muvaffaqiyatsiz bo'lganda ustozga (va zarur bo'lsa adminga)
# ANIQ sababni aytish uchun — avval har qanday xato jim "None" qaytarib,
# ustoz nega qo'lda tekshirayotganini bilmay qolardi.
AI_FAIL_REASON_LABELS = {
    "no_key": "AI ulanmagan",
    "quota": "AI kunlik/oylik limiti tugagan",
    "api_error": "AI xizmatida vaqtinchalik xatolik",
    "timeout": "AI javob berishga ulgurmadi (vaqt tugadi)",
    "network": "AI xizmatiga ulanishda xatolik",
    "parse_error": "AI javobini o'qib bo'lmadi",
    "voice_download": "Ovozli faylni yuklab bo'lmadi",
    "photo_unsupported": "AI hozircha rasm (fotosurat) shaklidagi javoblarni o'qiy olmaydi",
}

# ✅ AI kvotasi tugashi BUTUN tizim uchun umumiy muammo (bitta API kalit) —
# har bir muvaffaqiyatsiz baholashda alohida xabar yuborsak, admin bir zumda
# o'nlab bir xil xabarga to'lib ketardi. Shuning uchun shu turdagi
# ogohlantirish faqat cooldown oralig'ida BIR MARTA yuboriladi.
_LAST_AI_QUOTA_ALERT_TS = 0
AI_QUOTA_ALERT_COOLDOWN_SEC = 3600


async def _maybe_alert_admin_ai_quota(bot: Bot) -> None:
    global _LAST_AI_QUOTA_ALERT_TS
    now = now_ts()
    if now - _LAST_AI_QUOTA_ALERT_TS < AI_QUOTA_ALERT_COOLDOWN_SEC:
        return
    _LAST_AI_QUOTA_ALERT_TS = now
    try:
        await bot.send_message(
            ADMIN_ID,
            "⚠️ <b>Gemini AI limiti tugagan ko'rinadi</b>\n\n"
            "Writing/Speaking'ni avtomatik (AI) baholash hozircha ishlamayapti — "
            "tegishli o'qituvchilarga qo'lda tekshirish so'ralmoqda.\n\n"
            "Google AI Studio hisobingizda kredit/limitni tekshiring. "
            f"(Keyingi shunga o'xshash ogohlantirish {AI_QUOTA_ALERT_COOLDOWN_SEC // 60} daqiqadan keyin yuboriladi.)",
            parse_mode=ParseMode.HTML
        )
    except Exception:
        logging.exception("AI kvota ogohlantirishini adminga yuborib bo'lmadi")


async def _gemini_grade(
    prompt_text: str, *, essay_text: str | None = None,
    voice_file_id: str | None = None, bot: Bot | None = None, skill: str = "writing",
    scale: str = "ielts"
) -> dict:
    """Gemini API orqali Writing (matn) yoki Speaking (ovozli xabar)
    javobini baholaydi. Har doim dict qaytaradi:
      - muvaffaqiyatli: {"ok": True, "band": <0-9 yoki A1-C2>, "feedback": "..."}
      - muvaffaqiyatsiz: {"ok": False, "reason": "<AI_FAIL_REASON_LABELS kaliti>"}
    Muvaffaqiyatsiz bo'lganda chaqiruvchi qo'lda (ustoz) tekshirishga
    qaytishi VA sababni ustozga aytishi kerak — natija hech qachon
    "yo'qolib" qolmasligi, sababi ham hech qachon yashirin qolmasligi uchun.
    """
    if not GEMINI_API_KEY:
        return {"ok": False, "reason": "no_key"}

    skill_label = "Writing" if skill == "writing" else "Speaking"
    if scale == "cefr":
        instructions = (
            f"Siz tajribali CEFR {skill_label} baholovchisisiz. Quyidagi topshiriq/mavzu va "
            f"o'quvchi javobini rasmiy CEFR (A1-C2) mezonlari asosida baholang.\n\n"
            f"Topshiriq/mavzu: {prompt_text}\n\n"
            f"Faqat quyidagi JSON formatda javob bering, boshqa hech qanday matn, izoh yoki "
            f"markdown belgisi qo'shmang:\n"
            f'{{"level": "<A1, A2, B1, B2, C1 yoki C2>", '
            f'"feedback": "<o\'zbek tilida 2-4 jumlali aniq, foydali fikr-mulohaza>"}}'
        )
    else:
        instructions = (
            f"Siz tajribali IELTS {skill_label} baholovchisisiz. Quyidagi topshiriq/mavzu va "
            f"o'quvchi javobini rasmiy IELTS {skill_label} mezonlari asosida baholang.\n\n"
            f"Topshiriq/mavzu: {prompt_text}\n\n"
            f"Faqat quyidagi JSON formatda javob bering, boshqa hech qanday matn, izoh yoki "
            f"markdown belgisi qo'shmang:\n"
            f'{{"band": <0.0 dan 9.0 gacha, 0.5 qadamli raqam>, '
            f'"feedback": "<o\'zbek tilida 2-4 jumlali aniq, foydali fikr-mulohaza>"}}'
        )

    parts = [{"text": instructions}]
    if essay_text:
        parts.append({"text": f"O'quvchi javobi:\n{essay_text}"})
    if voice_file_id and bot:
        try:
            tg_file = await bot.get_file(voice_file_id)
            buf = await bot.download_file(tg_file.file_path)
            audio_b64 = base64.b64encode(buf.read()).decode("ascii")
            parts.append({"inline_data": {"mime_type": "audio/ogg", "data": audio_b64}})
        except Exception:
            logging.exception("Gemini uchun ovozli faylni yuklab bo'lmadi")
            return {"ok": False, "reason": "voice_download"}

    url = f"https://generativelanguage.googleapis.com/v1beta/models/{GEMINI_MODEL}:generateContent"
    payload = {"contents": [{"parts": parts}]}

    try:
        async with aiohttp.ClientSession(timeout=aiohttp.ClientTimeout(total=60)) as session:
            async with session.post(url, json=payload, headers={"x-goog-api-key": GEMINI_API_KEY}) as resp:
                if resp.status == 429:
                    logging.error(f"Gemini API kvota tugadi: {resp.status} {await resp.text()}")
                    return {"ok": False, "reason": "quota"}
                if resp.status != 200:
                    logging.error(f"Gemini API xatosi: {resp.status} {await resp.text()}")
                    return {"ok": False, "reason": "api_error"}
                data = await resp.json()

        raw_text = data["candidates"][0]["content"]["parts"][0]["text"]
        obj = _extract_json_object(raw_text)
        if not obj:
            return {"ok": False, "reason": "parse_error"}

        feedback = sanitize_plain_text(str(obj.get("feedback") or ""), max_len=1000)
        if not feedback:
            feedback = "🤖 AI tomonidan baholandi."

        if scale == "cefr":
            level = str(obj.get("level") or "").strip().upper()
            if level not in CEFR_LEVEL_ORDER:
                return {"ok": False, "reason": "parse_error"}
            return {"ok": True, "band": level, "feedback": feedback}

        if "band" not in obj:
            return {"ok": False, "reason": "parse_error"}
        band = round(_clamp(float(obj["band"]), 0.0, 9.0) * 2) / 2
        return {"ok": True, "band": band, "feedback": feedback}
    except asyncio.TimeoutError:
        logging.exception("Gemini API so'rovi vaqti tugadi")
        return {"ok": False, "reason": "timeout"}
    except Exception:
        logging.exception("Gemini API so'rovida xatolik")
        return {"ok": False, "reason": "network"}


# ----- Ustoz: bosqichli testdagi ochiq (x) javoblarni baholash -----
def _staged_open_kb(attempt_id: int, key: str, qnum: str, current) -> InlineKeyboardMarkup:
    ok_mark = "✔️ " if current is not None and int(current) == 1 else ""
    no_mark = "✔️ " if current is not None and int(current) == 0 else ""
    return InlineKeyboardMarkup(inline_keyboard=[[
        InlineKeyboardButton(text=f"{ok_mark}✅ To'g'ri", callback_data=f"sog_{attempt_id}_{qnum}_1_{key}"),
        InlineKeyboardButton(text=f"{no_mark}❌ Noto'g'ri", callback_data=f"sog_{attempt_id}_{qnum}_0_{key}"),
    ]])


async def _send_staged_open_for_review(bot: Bot, chat_id: int, attempt_id: int) -> int:
    """Bosqichli urinishdagi avtomatik tekshirib bo'lmagan (rasm/kalitsiz)
    ochiq javoblarni ustozga ✅/❌ tugmalari bilan yuboradi."""
    row = await POOL.fetchrow("""
        SELECT sa.sections_json, t.skills_cfg, t.title, t.code, u.full_name
        FROM staged_attempts sa JOIN tests t ON t.id=sa.test_id JOIN users u ON u.id=sa.user_id
        WHERE sa.id=$1
    """, attempt_id)
    if not row:
        return 0
    cfg = _load_skills_cfg(row["skills_cfg"])
    sections = _load_json_dict(row["sections_json"])
    items = []
    for k in cfg.get("order", []):
        sec_cfg = cfg["sections"].get(k, {})
        v = sections.get(k) or {}
        if sec_cfg.get("type") not in (STAGE_TYPE_MCQ, STAGE_TYPE_MCQ_BAND):
            continue
        grades = v.get("open_grades") or {}
        answers = v.get("open_answers") or {}
        refs = sec_cfg.get("open_ref_answers") or {}
        ck = normalize_answers(sec_cfg.get("answers") or "")
        for i, c in enumerate(ck):
            q = str(i + 1)
            if c == "x" and grades.get(q) is None and open_answer_has_content(answers.get(q)):
                items.append((k, sec_cfg.get("name", k), q, answers.get(q), refs.get(q)))
    if not items:
        return 0
    await bot.send_message(
        chat_id,
        f"📝 <b>Ochiq javoblarni tekshirish</b>\n\n👤 {_h(row['full_name'])}\n"
        f"📚 {_h(row['title'])} (<code>{row['code']}</code>)\n\nHar bir javob ostidagi ✅/❌ tugmasini bosing.",
        parse_mode=ParseMode.HTML
    )
    for k, label, q, raw, ref in items:
        text_val, photo_id = _open_answer_preview(raw)
        body = f"✏️ <b>{_h(label)} — {q}-savol:</b>\n{_h(text_val) if text_val else '[rasm]'}"
        if ref:
            body += f"\n\n🔑 Kutilgan javob: <b>{_h(ref)}</b>"
        kb = _staged_open_kb(attempt_id, k, q, None)
        try:
            if photo_id:
                await bot.send_photo(chat_id, photo_id, caption=body[:1000], parse_mode=ParseMode.HTML, reply_markup=kb)
            else:
                await bot.send_message(chat_id, body[:3900], parse_mode=ParseMode.HTML, reply_markup=kb)
        except Exception:
            logging.exception(f"Bosqichli ochiq javobni yuborib bo'lmadi (attempt={attempt_id})")
    return len(items)


def _recompute_stage_section(kind: str, key: str, sec_cfg: dict, v: dict) -> dict:
    ck = normalize_answers(sec_cfg.get("answers") or "")
    ua = normalize_answers(v.get("user_answers") or "")
    qn = min(len(ck), len(ua))
    score, total, pending = score_with_open(ua, ck, qn, v.get("open_grades") or {})
    v["score"], v["total"], v["open_pending"] = score, total, pending
    if sec_cfg.get("type") == STAGE_TYPE_MCQ_BAND:
        v["band"] = _stage_band_for_mcq(kind, key, score, total)
    else:
        w = float(v.get("weight") or sec_cfg.get("weight") or 1.0)
        v["earned"], v["max"] = score * w, total * w
    return v


@router.callback_query(F.data.startswith("sog_"))
async def staged_open_grade_cb(callback: CallbackQuery, state: FSMContext, bot: Bot) -> None:
    parts = callback.data.split("_", 4)
    if len(parts) != 5 or not parts[1].isdigit():
        await callback.answer("Noto'g'ri tugma.", show_alert=True)
        return
    attempt_id, qnum, val, key = int(parts[1]), parts[2], (1 if parts[3] == "1" else 0), parts[4]
    row = await POOL.fetchrow("""
        SELECT sa.test_id, sa.user_id, sa.status, sa.sections_json, t.skills_cfg, t.creator_id
        FROM staged_attempts sa JOIN tests t ON t.id=sa.test_id WHERE sa.id=$1
    """, attempt_id)
    if not row:
        await callback.answer("Topilmadi.", show_alert=True)
        return
    if not (await is_admin_user(callback.from_user.id) or row["creator_id"] == callback.from_user.id):
        await callback.answer("Sizda ruxsat yo'q!", show_alert=True)
        return
    if row["status"] != "pending_review":
        await callback.answer("Bu topshiriq allaqachon yakunlangan.", show_alert=True)
        return
    cfg = _load_skills_cfg(row["skills_cfg"])
    kind = cfg["kind"]
    sections = _load_json_dict(row["sections_json"])
    sec_cfg = cfg["sections"].get(key)
    v = sections.get(key)
    if not sec_cfg or v is None:
        await callback.answer("Bosqich topilmadi.", show_alert=True)
        return
    grades = dict(v.get("open_grades") or {})
    grades[qnum] = val
    v["open_grades"] = grades
    sections[key] = _recompute_stage_section(kind, key, sec_cfg, v)
    await POOL.execute("UPDATE staged_attempts SET sections_json=$1 WHERE id=$2",
                       json.dumps(sections, ensure_ascii=False), attempt_id)

    fsm = await state.get_data()
    in_grading = fsm.get("grading_attempt_id") == attempt_id
    if in_grading:
        await state.update_data(grading_sections=sections)

    try:
        await callback.message.edit_reply_markup(reply_markup=_staged_open_kb(attempt_id, key, qnum, val))
    except Exception:
        pass
    await callback.answer("✅ To'g'ri deb belgilandi" if val else "❌ Noto'g'ri deb belgilandi")

    if any(int((sections.get(k) or {}).get("open_pending") or 0) for k in cfg["order"]):
        return

    if kind == "dtm_staged":
        await _complete_dtm_staged_attempt(bot, attempt_id, graded_by=callback.from_user.id)
        if in_grading:
            await state.clear()
        await callback.message.answer("✅ Barcha ochiq javoblar baholandi — natija o'quvchiga yuborildi.")
        return

    if in_grading:
        await _ask_next_grading_field(callback.message, state, bot)
        return
    ws_pending = any(
        cfg["sections"].get(k, {}).get("type") in (STAGE_TYPE_TEXT, STAGE_TYPE_VOICE)
        and (sections.get(k) or {}).get("band") is None for k in cfg["order"]
    )
    if ws_pending:
        kb = InlineKeyboardMarkup(inline_keyboard=[[
            InlineKeyboardButton(text="✅ Writing/Speaking'ni baholash", callback_data=f"skg_start_{attempt_id}")
        ]])
        await callback.message.answer("✅ Ochiq javoblar baholandi. Endi Writing/Speaking'ni baholang:",
                                      reply_markup=kb)
        return
    overall = await _apply_staged_grading_result(
        bot, attempt_id, row["test_id"], row["user_id"], sections, cfg["order"], cfg["sections"],
        graded_by=callback.from_user.id, kind=kind
    )
    await callback.message.answer(f"✅ Baholandi va o'quvchiga yuborildi. Natija: {overall}")


# ----- Ustoz: Writing/Speaking'ni baholash -----
def _parse_band(text: str) -> float | None:
    try:
        v = float((text or "").strip().replace(",", "."))
    except ValueError:
        return None
    if v < 0 or v > 9:
        return None
    return round(v * 2) / 2


@router.callback_query(F.data.startswith("skg_pending_"))
async def staged_grading_pending_list(callback: CallbackQuery) -> None:
    test_id = int(callback.data.split("_")[2])
    t = await POOL.fetchrow("SELECT creator_id, title, code FROM tests WHERE id=$1", test_id)
    if not t:
        return await callback.answer("Test topilmadi.", show_alert=True)
    if not (await is_admin_user(callback.from_user.id) or t["creator_id"] == callback.from_user.id):
        return await callback.answer("Sizda ruxsat yo'q!", show_alert=True)

    rows = await POOL.fetch("""
        SELECT sa.id, u.full_name, sa.finished_at
        FROM staged_attempts sa
        JOIN users u ON u.id = sa.user_id
        WHERE sa.test_id=$1 AND sa.status='pending_review'
        ORDER BY sa.finished_at ASC
        LIMIT 30
    """, test_id)

    await callback.answer()
    if not rows:
        await callback.message.answer("✅ Hozircha baholash kutayotgan topshiriq yo'q.")
        return

    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(
            text=f"{r['full_name']} — {datetime.fromtimestamp(r['finished_at'], TZ).strftime('%d-%m %H:%M') if r['finished_at'] else '?'}",
            callback_data=f"skg_start_{r['id']}"
        )]
        for r in rows
    ])
    await callback.message.answer(
        f"🎓 <b>{t['title']}</b> (<code>{t['code']}</code>) — baholash kutayotganlar:",
        parse_mode=ParseMode.HTML, reply_markup=kb
    )


@router.callback_query(F.data.startswith("skg_start_"))
async def staged_grading_start(callback: CallbackQuery, state: FSMContext, bot: Bot) -> None:
    attempt_id = int(callback.data.split("_")[2])
    row = await POOL.fetchrow("""
        SELECT sa.test_id, sa.user_id, sa.status, sa.sections_json, t.skills_cfg, t.creator_id
        FROM staged_attempts sa
        JOIN tests t ON t.id = sa.test_id
        WHERE sa.id=$1
    """, attempt_id)
    if not row:
        await callback.answer("Topilmadi.", show_alert=True)
        return
    if not (await is_admin_user(callback.from_user.id) or row["creator_id"] == callback.from_user.id):
        await callback.answer("Sizda ruxsat yo'q!", show_alert=True)
        return
    if row["status"] != "pending_review":
        await callback.answer("Bu topshiriq allaqachon baholangan yoki hali tugallanmagan.", show_alert=True)
        return

    cfg = _load_skills_cfg(row["skills_cfg"])
    sections = json.loads(row["sections_json"] or "{}")

    await state.update_data(
        grading_attempt_id=attempt_id, grading_test_id=row["test_id"], grading_user_db_id=row["user_id"],
        grading_sections=sections, grading_order=cfg["order"], grading_cfg=cfg["sections"],
        grading_teacher_tg_id=callback.from_user.id, grading_kind=cfg["kind"],
    )

    if "writing" in sections:
        w = sections["writing"]
        if w.get("photo_id"):
            await callback.message.answer_photo(w["photo_id"], caption="✍️ Writing javobi (rasm)")
        if w.get("text"):
            await callback.message.answer(f"✍️ <b>Writing matni:</b>\n\n{_h(w['text'])}", parse_mode=ParseMode.HTML)
    if "speaking" in sections and sections["speaking"].get("voice_id"):
        await callback.message.answer_voice(sections["speaking"]["voice_id"], caption="🎙 Speaking javobi")

    for k in ("reading", "listening"):
        if k in sections:
            v = sections[k]
            emoji = "📖" if k == "reading" else "🎧"
            await callback.message.answer(
                f"{emoji} <b>{k.title()}:</b> {v.get('score', 0)}/{v.get('total', 0)} — band {v.get('band', '-')}",
                parse_mode=ParseMode.HTML
            )

    await callback.answer()
    await _send_staged_open_for_review(bot, callback.from_user.id, attempt_id)
    await _ask_next_grading_field(callback.message, state, bot)


def _parse_grade_input(text: str, kind: str) -> float | str | None:
    if kind == "cefr4":
        level = (text or "").strip().upper()
        return level if level in CEFR_LEVEL_ORDER else None
    return _parse_band(text)


def _grade_input_hint(kind: str) -> str:
    if kind == "cefr4":
        return "A1, A2, B1, B2, C1 yoki C2 darajasini kiriting:"
    return "0 dan 9 gacha, 0.5 qadamli raqam kiriting (masalan: 6.5):"


async def _ask_next_grading_field(answerable, state: FSMContext, bot: Bot) -> None:
    data = await state.get_data()
    sections = data["grading_sections"]
    kind = data.get("grading_kind", "ielts4")
    hint = _grade_input_hint(kind)

    if any(int((sections.get(k) or {}).get("open_pending") or 0) for k in data.get("grading_order", [])):
        await answerable.answer(
            "👆 Avval yuqoridagi ochiq javoblarni ✅/❌ tugmalari bilan baholang — shundan so'ng "
            "davom etamiz.")
        return

    if "writing" in sections and sections["writing"].get("band") is None:
        await answerable.answer(f"✍️ <b>Writing</b> uchun baho bering — {hint}", parse_mode=ParseMode.HTML)
        await state.set_state(StagedGradingState.writing_band)
        return
    if "speaking" in sections and sections["speaking"].get("band") is None:
        await answerable.answer(f"🎙 <b>Speaking</b> uchun baho bering — {hint}", parse_mode=ParseMode.HTML)
        await state.set_state(StagedGradingState.speaking_band)
        return

    await _finalize_staged_grading(answerable, state, bot)


@router.message(StagedGradingState.writing_band)
async def grading_writing_band(message: Message, state: FSMContext, bot: Bot) -> None:
    data = await state.get_data()
    kind = data.get("grading_kind", "ielts4")
    band = _parse_grade_input(message.text or "", kind)
    if band is None:
        await message.answer(f"⚠️ {_grade_input_hint(kind)}")
        return
    await state.update_data(pending_writing_band=band)
    await message.answer("✍️ Writing uchun qisqa fikr-mulohaza (feedback) yozing:")
    await state.set_state(StagedGradingState.writing_feedback)


@router.message(StagedGradingState.writing_feedback)
async def grading_writing_feedback(message: Message, state: FSMContext, bot: Bot) -> None:
    feedback = sanitize_plain_text(message.text or "", max_len=1000)
    if len(feedback) < 2:
        await message.answer("⚠️ Fikr-mulohaza juda qisqa. Qayta yozing:")
        return
    data = await state.get_data()
    sections = data["grading_sections"]
    sections["writing"]["band"] = data["pending_writing_band"]
    sections["writing"]["feedback"] = feedback
    await state.update_data(grading_sections=sections)
    await _ask_next_grading_field(message, state, bot)


@router.message(StagedGradingState.speaking_band)
async def grading_speaking_band(message: Message, state: FSMContext, bot: Bot) -> None:
    data = await state.get_data()
    kind = data.get("grading_kind", "ielts4")
    band = _parse_grade_input(message.text or "", kind)
    if band is None:
        await message.answer(f"⚠️ {_grade_input_hint(kind)}")
        return
    await state.update_data(pending_speaking_band=band)
    await message.answer("🎙 Speaking uchun qisqa fikr-mulohaza (feedback) yozing:")
    await state.set_state(StagedGradingState.speaking_feedback)


@router.message(StagedGradingState.speaking_feedback)
async def grading_speaking_feedback(message: Message, state: FSMContext, bot: Bot) -> None:
    feedback = sanitize_plain_text(message.text or "", max_len=1000)
    if len(feedback) < 2:
        await message.answer("⚠️ Fikr-mulohaza juda qisqa. Qayta yozing:")
        return
    data = await state.get_data()
    sections = data["grading_sections"]
    sections["speaking"]["band"] = data["pending_speaking_band"]
    sections["speaking"]["feedback"] = feedback
    await state.update_data(grading_sections=sections)
    await _ask_next_grading_field(message, state, bot)


async def _finalize_staged_grading(answerable, state: FSMContext, bot: Bot) -> None:
    data = await state.get_data()
    attempt_id = data["grading_attempt_id"]
    test_id = data["grading_test_id"]
    user_db_id = data["grading_user_db_id"]
    sections = data["grading_sections"]
    order = data["grading_order"]
    section_cfg = data["grading_cfg"]
    teacher_tg_id = data["grading_teacher_tg_id"]
    kind = data.get("grading_kind", "ielts4")

    if kind == "dtm_staged":
        await _complete_dtm_staged_attempt(bot, attempt_id, graded_by=teacher_tg_id)
        await answerable.answer("✅ Baholandi va o'quvchiga yuborildi.")
        await state.clear()
        return

    overall = await _apply_staged_grading_result(
        bot, attempt_id, test_id, user_db_id, sections, order, section_cfg, graded_by=teacher_tg_id, kind=kind
    )

    await answerable.answer(f"✅ Baholandi va o'quvchiga yuborildi. Natija: {overall}")
    await state.clear()


def _stage_overall_percent(kind: str, overall_value) -> float:
    if kind == "cefr4":
        ordinal = _cefr_level_to_ordinal(str(overall_value))
        return ordinal / (len(CEFR_LEVEL_ORDER) - 1) * 100.0
    if kind == "sat_staged":
        return (float(overall_value) - 400.0) / 1200.0 * 100.0
    return float(overall_value) / 9.0 * 100.0


def _stage_cert_eligible(kind: str, overall_value) -> bool:
    if kind == "cefr4":
        return _cefr_level_to_ordinal(str(overall_value)) >= 2  # B1 va undan yuqori
    if kind == "sat_staged":
        return float(overall_value) >= 1000.0  # taxminan o'rtachadan yuqori (~1050 o'rtacha)
    return float(overall_value) >= 5.0


async def _apply_staged_grading_result(
    bot: Bot, attempt_id: int, test_id: int, user_db_id: int,
    sections: dict, order: list, section_cfg: dict, graded_by: int | None, kind: str = "ielts4"
):
    """sections'dagi barcha kerakli bo'limlar allaqachon band/daraja+feedback
    bilan to'liq (ustoz qo'lda kiritgan bo'lsin, AI hisoblagan bo'lsin —
    farqi yo'q) — buni yakunlab, natijani saqlaydi, sertifikat va yakuniy
    xabarni yuboradi. graded_by=None bo'lsa (hammasi AI/avtomatik
    hisoblangan bo'lsa) shunga mos qayd etiladi. Hisoblangan umumiy
    natijani (IELTS uchun raqam, CEFR uchun daraja) qaytaradi."""
    overall = _compute_overall_stage_grade(kind, sections, order)
    if kind == "cefr4":
        overall_label = "Umumiy daraja"
    elif kind == "sat_staged":
        overall_label = "Umumiy SAT ball (400-1600)"
    else:
        overall_label = "Overall band"
    overall_percent = _stage_overall_percent(kind, overall)

    await POOL.execute("""
        UPDATE staged_attempts
        SET status='graded', overall_score=$1, overall_grade=$2, sections_json=$3, graded_by=$4, graded_at=$5
        WHERE id=$6
    """,
        overall_percent, str(overall), json.dumps(sections, ensure_ascii=False),
        graded_by, now_ts(), attempt_id
    )

    t = await POOL.fetchrow("""
        SELECT title, code, subject, cert_academy_name, cert_signee_name, cert_template, creator_id
        FROM tests WHERE id=$1
    """, test_id)
    u = await POOL.fetchrow("SELECT telegram_id, full_name FROM users WHERE id=$1", user_db_id)
    student_tg_id = u["telegram_id"]

    # ✅ Umumlashtirilgan: "reading"/"listening" kabi qattiq kodlangan kalitlar
    # o'rniga — MCQ_BAND turidagi (avtomatik baholanadigan) barcha bosqichlar
    # bo'yicha xom to'g'ri javoblar/jami savollarni yig'adi (IELTS/CEFR uchun
    # Reading+Listening, SAT uchun Reading&Writing+Math).
    raw_score = sum(
        int(sections.get(k, {}).get("score") or 0)
        for k in order if section_cfg.get(k, {}).get("type") == STAGE_TYPE_MCQ_BAND
    )
    raw_total = sum(
        int(sections.get(k, {}).get("total") or 0)
        for k in order if section_cfg.get(k, {}).get("type") == STAGE_TYPE_MCQ_BAND
    )

    # ✅ Qayta yechish mumkin — "results"dagi yagona qatorni eng so'nggi
    # urinish bilan yangilaymiz (upsert), to'liq tarix staged_attempts'da qoladi.
    await POOL.execute("""
        INSERT INTO results (
            user_id, test_id, score, total_questions, percent, grade, details, finished_ts, date
        )
        VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9)
        ON CONFLICT (user_id, test_id) DO UPDATE SET
            score=EXCLUDED.score, total_questions=EXCLUDED.total_questions,
            percent=EXCLUDED.percent, grade=EXCLUDED.grade, details=EXCLUDED.details,
            finished_ts=EXCLUDED.finished_ts, date=EXCLUDED.date,
            certificate_sent=0, certificate_file_id=NULL, certificate_sent_ts=0
    """,
        user_db_id, test_id, raw_score, raw_total,
        overall_percent, str(overall),
        json.dumps({"stages": sections}, ensure_ascii=False),
        now_ts(), now_str_local()
    )

    lines = []
    for k in order:
        v = sections.get(k, {})
        label = section_cfg[k]["name"]
        stype = section_cfg[k]["type"]
        if stype == STAGE_TYPE_MCQ_BAND:
            lines.append(f"• {label}: <b>{v.get('band', '-')}</b> ({v.get('score', 0)}/{v.get('total', 0)})")
        else:
            fb = _h(v.get("feedback") or "-")
            lines.append(f"• {_h(label)}: <b>{v.get('band', '-')}</b>\n   💬 {fb}")

    admin_flag = await is_admin_user(student_tg_id)
    try:
        await bot.send_message(
            student_tg_id,
            "🎓 <b>Natijangiz baholandi!</b>\n\n" + "\n".join(lines) +
            f"\n\n🏆 <b>{overall_label}: {overall}</b>",
            parse_mode=ParseMode.HTML,
            reply_markup=main_menu_kb(admin_flag)
        )
    except Exception:
        logging.exception(f"Natijani studentga yuborib bo'lmadi (tg={student_tg_id})")

    if _stage_cert_eligible(kind, overall):
        try:
            cert_academy, cert_signee, cert_tpl = await _resolve_cert_settings(
                t["cert_academy_name"], t["cert_signee_name"], t["cert_template"], t["creator_id"])
            tmp_dir = tempfile.gettempdir()
            safe_code = re.sub(r"[^A-Z0-9\-_]+", "_", str(t["code"]).upper())
            pdf_path = os.path.join(tmp_dir, f"cert_{safe_code}_{user_db_id}.pdf")
            cert_subject_fallback = "CEFR" if kind == "cefr4" else ("SAT" if kind == "sat_staged" else "IELTS")
            make_certificate_pdf(
                full_name=u["full_name"] or "Noma'lum", subject=t["subject"] or cert_subject_fallback,
                test_code=t["code"], test_title=t["title"], exam_date_str=now_str_local(),
                percent=overall_percent, correct_count=raw_score,
                total_questions=max(raw_total, 1), grade=str(overall), out_path=pdf_path,
                academy_name=cert_academy, signee_name=cert_signee, template=cert_tpl,
            )
            sent = await bot.send_document(student_tg_id, FSInputFile(pdf_path), caption="🎁 Sertifikatingiz tayyor!")
            cert_file_id = sent.document.file_id if sent and sent.document else None
            await POOL.execute(
                "UPDATE results SET certificate_sent=1, certificate_file_id=$1, certificate_sent_ts=$2 "
                "WHERE user_id=$3 AND test_id=$4",
                cert_file_id, now_ts(), user_db_id, test_id
            )
        except Exception:
            logging.exception(f"Sertifikat yuborilmadi (user={user_db_id}, test={test_id})")

    return overall


# ===================== INTERACTIVE & TEXT SUBMISSION LOGIC =====================
# ✅ OCHIQ (x) SAVOLLARNI BAHOLASH. Ustoz test yaratishda har bir ochiq savol
# uchun kutilgan javobni (kalit) kiritadi — shuning uchun o'quvchining matnli
# javobi shu kalit bilan AVTOMATIK solishtiriladi va ustozdan hech narsa
# talab qilinmaydi. Faqat avtomatik tekshirib bo'lmaydigan javoblar (o'quvchi
# javobni rasm qilib yuborgan yoki kalit kiritilmagan eski testlar) ustozga
# ✅/❌ tugmalari bilan yuboriladi. Ustoz avtomatik bahoni ham xohlasa
# o'zgartirishi mumkin. Avval o'quvchiga "ustoz tekshiradi" deyilardi, lekin
# ustozda baholash imkoniyati umuman yo'q edi — ochiq savollar hech qachon
# hisobga kirmasdi.
def _open_grade_kb(result_id: int, qnum: str, current) -> InlineKeyboardMarkup:
    ok_mark = "✔️ " if current is not None and int(current) == 1 else ""
    no_mark = "✔️ " if current is not None and int(current) == 0 else ""
    return InlineKeyboardMarkup(inline_keyboard=[[
        InlineKeyboardButton(text=f"{ok_mark}✅ To'g'ri", callback_data=f"og_{result_id}_{qnum}_1"),
        InlineKeyboardButton(text=f"{no_mark}❌ Noto'g'ri", callback_data=f"og_{result_id}_{qnum}_0"),
    ]])


async def _send_open_answers_for_review(bot: Bot, chat_id: int, result_id: int, only_pending: bool = True) -> int:
    """Bitta natijadagi ochiq javoblarni ✅/❌ tugmalari bilan ustozga yuboradi.
    only_pending=True — faqat avtomatik tekshirib bo'lmaganlari (rasm/kalitsiz)."""
    row = await POOL.fetchrow("""
        SELECT r.open_text_answers, r.open_grades, u.full_name, t.title, t.code, t.answers, t.open_ref_answers
        FROM results r JOIN users u ON u.id=r.user_id JOIN tests t ON t.id=r.test_id
        WHERE r.id=$1
    """, result_id)
    if not row:
        return 0
    answers = _load_json_dict(row["open_text_answers"])
    grades = _load_json_dict(row["open_grades"])
    refs = _load_json_dict(row["open_ref_answers"])
    key = normalize_answers(row["answers"] or "")
    qnums = [str(i + 1) for i, c in enumerate(key) if c == "x"]
    to_send = [q for q in qnums
               if open_answer_has_content(answers.get(q)) and (not only_pending or grades.get(q) is None)]
    if not to_send:
        return 0

    await bot.send_message(
        chat_id,
        f"📝 <b>Ochiq javoblarni tekshirish</b>\n\n"
        f"👤 {_h(row['full_name'])}\n"
        f"📚 {_h(row['title'])} (<code>{row['code']}</code>)\n\n"
        f"Har bir javob ostidagi ✅/❌ tugmasini bosing — natija avtomatik qayta hisoblanadi.",
        parse_mode=ParseMode.HTML
    )
    for q in sorted(to_send, key=int):
        text_val, photo_id = _open_answer_preview(answers.get(q))
        body = f"✏️ <b>{q}-savol javobi:</b>\n{_h(text_val) if text_val else '[rasm]'}"
        if refs.get(q):
            body += f"\n\n🔑 Kutilgan javob: <b>{_h(refs[q])}</b>"
        if grades.get(q) is not None:
            body += "\n🤖 Avtomatik baho: " + ("✅ to'g'ri" if int(grades[q]) == 1 else "❌ xato")
        kb = _open_grade_kb(result_id, q, grades.get(q))
        try:
            if photo_id:
                await bot.send_photo(chat_id, photo_id, caption=body[:1000], parse_mode=ParseMode.HTML, reply_markup=kb)
            else:
                await bot.send_message(chat_id, body[:3900], parse_mode=ParseMode.HTML, reply_markup=kb)
        except Exception:
            logging.exception(f"Ochiq javobni tekshirishga yuborib bo'lmadi (result={result_id}, q={q})")
    return len(to_send)


async def _recompute_result_scores(result_id: int, grades: dict | None = None) -> dict | None:
    """Natijaning ball/foiz/darajasini ochiq savol baholari bilan qayta
    hisoblaydi va bazaga yozadi."""
    row = await POOL.fetchrow("""
        SELECT r.id, r.user_id, r.test_id, r.user_answers, r.open_grades, r.details, r.grade,
               t.answers, t.questions_count, t.exam_type, t.dtm_cfg
        FROM results r JOIN tests t ON t.id=r.test_id WHERE r.id=$1
    """, result_id)
    if not row:
        return None
    key = normalize_answers(row["answers"] or "")
    ua = normalize_answers(row["user_answers"] or "")
    qcount = min(int(row["questions_count"] or 0), len(key))
    if grades is None:
        grades = _load_json_dict(row["open_grades"])

    score, scored, pending = score_with_open(ua, key, qcount, grades)
    percent = (score / scored * 100.0) if scored > 0 else 0.0
    earned, maxp, details = float(score), float(scored), row["details"]
    et = row["exam_type"] or "simple"
    percent_points = percent
    if et == "dtm":
        earned, maxp, det = dtm_score_points(ua, key, row["dtm_cfg"] or "", grades)
        details = json.dumps(det, ensure_ascii=False)
        percent_points = (earned / maxp * 100.0) if maxp > 0 else 0.0
    grade = row["grade"]
    if et in ("cefr", "ielts", "sat"):
        grade = grade_value_for_exam_type(et, percent_points)

    await POOL.execute("""
        UPDATE results SET score=$1, percent=$2, earned_points=$3, max_points=$4, details=$5, grade=$6,
               open_grades=$7, open_pending=$8
        WHERE id=$9
    """, int(score), float(percent_points), float(earned), float(maxp), details, grade,
        json.dumps(grades, ensure_ascii=False), int(pending), result_id)
    return {"score": score, "scored": scored, "pending": pending, "percent": percent_points, "grade": grade,
            "exam_type": et, "test_id": row["test_id"], "user_id": row["user_id"], "qcount": qcount}


def _merge_regraded_open(key: str, answers: dict, refs: dict, old: dict) -> dict:
    """Yangi kalit bo'yicha ochiq javoblarni qayta avtomatik baholaydi, lekin
    ustoz qo'lda qo'ygan baholarni (rasm/kalitsiz javoblar) saqlab qoladi."""
    new = initial_open_grades(key, answers, refs)
    for q, g in (old or {}).items():
        qi = int(q) - 1 if str(q).isdigit() else -1
        if q not in new and 0 <= qi < len(key) and key[qi] == "x" and open_answer_has_content(answers.get(q)):
            new[q] = g
    return new


async def _regrade_test_results(test_id: int) -> int:
    """✅ Ustoz javob kalitini tuzatsa — shu testning BARCHA mavjud natijalari
    yangi kalit bo'yicha qayta hisoblanadi. Avval kalit o'zgartirilsa ham
    avvalgi (xato kalit bilan hisoblangan) ballar o'zgarishsiz qolardi."""
    t = await POOL.fetchrow(
        "SELECT answers, open_ref_answers, questions_count, exam_type FROM tests WHERE id=$1", test_id)
    if not t:
        return 0
    key = normalize_answers(t["answers"] or "")
    qcount = min(int(t["questions_count"] or 0), len(key))
    refs = _load_json_dict(t["open_ref_answers"])
    rows = await POOL.fetch(
        "SELECT id, open_text_answers, open_grades FROM results WHERE test_id=$1 AND user_answers IS NOT NULL",
        test_id)
    for r in rows:
        grades = _merge_regraded_open(key[:qcount], _load_json_dict(r["open_text_answers"]), refs,
                                      _load_json_dict(r["open_grades"]))
        await _recompute_result_scores(r["id"], grades)
    # sertifikatga "loyiq emas" deb belgilanganlar qayta ko'rib chiqiladi
    await POOL.execute("UPDATE results SET certificate_sent=0 WHERE test_id=$1 AND certificate_sent=2", test_id)
    if (t["exam_type"] or "") == "rasch" and rows:
        await POOL.execute("UPDATE tests SET rush_finalized=0 WHERE id=$1", test_id)
        await POOL.execute("UPDATE results SET rasch_ball=NULL, rasch_percent=NULL, grade=NULL WHERE test_id=$1",
                           test_id)
    return len(rows)


async def _refresh_staged_scores(attempt_id: int) -> None:
    """Baholangan bosqichli urinishning umumiy natijasini (va agar u eng so'nggi
    urinish bo'lsa — results qatorini) jim qayta hisoblaydi."""
    a = await POOL.fetchrow("SELECT user_id, test_id, sections_json FROM staged_attempts WHERE id=$1", attempt_id)
    if not a:
        return
    cfg = _load_skills_cfg(await POOL.fetchval("SELECT skills_cfg FROM tests WHERE id=$1", a["test_id"]))
    kind, order, section_cfg = cfg.get("kind"), cfg.get("order", []), cfg.get("sections", {})
    sections = _load_json_dict(a["sections_json"])
    latest = await POOL.fetchval("SELECT MAX(id) FROM staged_attempts WHERE user_id=$1 AND test_id=$2",
                                 a["user_id"], a["test_id"])
    details = json.dumps({"stages": sections}, ensure_ascii=False)
    if kind == "dtm_staged":
        earned = sum(float(sections.get(k, {}).get("earned") or 0) for k in order)
        maxp = sum(float(sections.get(k, {}).get("max") or 0) for k in order)
        percent = (earned / maxp * 100.0) if maxp > 0 else 0.0
        raw_score = sum(int(sections.get(k, {}).get("score") or 0) for k in order)
        raw_total = sum(int(sections.get(k, {}).get("total") or 0) for k in order)
        await POOL.execute("UPDATE staged_attempts SET overall_score=$1 WHERE id=$2", percent, attempt_id)
        if latest == attempt_id:
            await POOL.execute("""
                UPDATE results SET score=$1, total_questions=$2, percent=$3, earned_points=$4, max_points=$5,
                       details=$6 WHERE user_id=$7 AND test_id=$8
            """, raw_score, raw_total, percent, earned, maxp, details, a["user_id"], a["test_id"])
    else:
        overall = _compute_overall_stage_grade(kind, sections, order)
        percent = _stage_overall_percent(kind, overall)
        raw_score = sum(int(sections.get(k, {}).get("score") or 0)
                        for k in order if section_cfg.get(k, {}).get("type") == STAGE_TYPE_MCQ_BAND)
        raw_total = sum(int(sections.get(k, {}).get("total") or 0)
                        for k in order if section_cfg.get(k, {}).get("type") == STAGE_TYPE_MCQ_BAND)
        await POOL.execute("UPDATE staged_attempts SET overall_score=$1, overall_grade=$2 WHERE id=$3",
                           percent, str(overall), attempt_id)
        if latest == attempt_id:
            await POOL.execute("""
                UPDATE results SET score=$1, total_questions=$2, percent=$3, grade=$4, details=$5
                WHERE user_id=$6 AND test_id=$7
            """, raw_score, raw_total, percent, str(overall), details, a["user_id"], a["test_id"])
    if latest == attempt_id:
        await POOL.execute("UPDATE results SET certificate_sent=0 WHERE user_id=$1 AND test_id=$2 AND certificate_sent=2",
                           a["user_id"], a["test_id"])


async def _regrade_staged_section(test_id: int, key: str) -> int:
    """Bosqichli testda bitta bosqichning kaliti o'zgarsa — shu bosqichni
    yechgan barcha urinishlar qayta hisoblanadi."""
    cfg = _load_skills_cfg(await POOL.fetchval("SELECT skills_cfg FROM tests WHERE id=$1", test_id))
    sec_cfg = cfg.get("sections", {}).get(key)
    if not sec_cfg:
        return 0
    kind = cfg.get("kind")
    ck = normalize_answers(sec_cfg.get("answers") or "")
    refs = sec_cfg.get("open_ref_answers") or {}
    n = 0
    for a in await POOL.fetch("SELECT id, status, sections_json FROM staged_attempts WHERE test_id=$1", test_id):
        sections = _load_json_dict(a["sections_json"])
        v = sections.get(key)
        if not v or "user_answers" not in v:
            continue
        ua = normalize_answers(v.get("user_answers") or "")
        qn = min(len(ck), len(ua))
        v["open_grades"] = _merge_regraded_open(ck[:qn], v.get("open_answers") or {}, refs, v.get("open_grades") or {})
        sections[key] = _recompute_stage_section(kind, key, sec_cfg, v)
        await POOL.execute("UPDATE staged_attempts SET sections_json=$1 WHERE id=$2",
                           json.dumps(sections, ensure_ascii=False), a["id"])
        if a["status"] == "graded":
            await _refresh_staged_scores(a["id"])
        n += 1
    return n


def _cert_decision(exam_type: str, percent, grade, rasch_finalized: bool, open_pending: int):
    """True = sertifikat beriladi, False = berilmaydi (yakuniy), None = hali
    aniq emas (Rasch hisoblanmagan yoki ochiq javoblar tekshirilmoqda)."""
    if int(open_pending or 0) > 0:
        return None
    et = exam_type or "simple"
    if et == "rasch":
        if not rasch_finalized:
            return None
        return grade is not None
    if et in ("ielts4", "cefr4", "sat_staged"):
        if grade is None:
            return False
        try:
            return _stage_cert_eligible(et, grade)
        except (ValueError, TypeError):
            return False
    return float(percent or 0.0) >= float(CERT_MIN_PERCENT)


async def _send_result_certificate(bot: Bot, result_id: int) -> bool:
    """Bitta natija uchun sertifikat PDF yaratib yuboradi va belgilab qo'yadi."""
    row = await POOL.fetchrow("""
        SELECT r.user_id, r.test_id, r.percent, r.rasch_percent, r.grade, r.date, r.score, r.total_questions,
               u.telegram_id, u.full_name,
               t.code, t.title, t.subject, t.exam_type,
               t.cert_academy_name, t.cert_signee_name, t.cert_template, t.creator_id
        FROM results r JOIN users u ON u.id=r.user_id JOIN tests t ON t.id=r.test_id
        WHERE r.id=$1
    """, result_id)
    if not row:
        return False
    et = row["exam_type"] or "simple"
    pct = row["rasch_percent"] if et == "rasch" else row["percent"]
    cert_academy, cert_signee, cert_tpl = await _resolve_cert_settings(
        row["cert_academy_name"], row["cert_signee_name"], row["cert_template"], row["creator_id"])
    safe_code = re.sub(r"[^A-Z0-9\-_]+", "_", str(row["code"]).upper())
    pdf_path = os.path.join(tempfile.gettempdir(), f"cert_{safe_code}_{row['user_id']}.pdf")
    make_certificate_pdf(
        full_name=row["full_name"] or "Noma'lum",
        subject=(row["subject"] or "Noma'lum"),
        test_code=row["code"],
        test_title=row["title"],
        exam_date_str=(row["date"] or now_str_local()),
        percent=float(pct or 0.0),
        correct_count=int(row["score"] or 0),
        total_questions=int(row["total_questions"] or 0),
        grade=row["grade"],
        out_path=pdf_path,
        academy_name=cert_academy,
        signee_name=cert_signee,
        template=cert_tpl,
    )
    sent = await bot.send_document(
        chat_id=int(row["telegram_id"]),
        document=FSInputFile(pdf_path),
        caption="🎁 Sertifikatingiz tayyor! (PDF)"
    )
    cert_file_id = sent.document.file_id if sent and sent.document else None
    await POOL.execute("""
        UPDATE results SET certificate_sent=1, certificate_file_id=$1, certificate_sent_ts=$2 WHERE id=$3
    """, cert_file_id, now_ts(), result_id)
    return True


async def _notify_student_open_graded(bot: Bot, result_id: int) -> None:
    row = await POOL.fetchrow("""
        SELECT r.open_grades, r.score, r.total_questions, r.percent, r.grade, r.earned_points, r.max_points,
               u.telegram_id, t.title, t.code, t.exam_type, t.answers
        FROM results r JOIN users u ON u.id=r.user_id JOIN tests t ON t.id=r.test_id WHERE r.id=$1
    """, result_id)
    if not row:
        return
    grades = _load_json_dict(row["open_grades"])
    key = normalize_answers(row["answers"] or "")
    lines = []
    for i, c in enumerate(key):
        if c != "x":
            continue
        g = grades.get(str(i + 1))
        if g is None:
            continue
        lines.append(f"{'✅' if int(g) == 1 else '❌'} {i + 1}-savol: {'to‘g‘ri' if int(g) == 1 else 'noto‘g‘ri'}")
    et = row["exam_type"] or "simple"
    text = (
        f"✍️ <b>Ochiq savollaringiz tekshirildi</b>\n"
        f"📚 {_h(row['title'])} (<code>{row['code']}</code>)\n\n" + "\n".join(lines)
    )
    if et == "rasch":
        text += "\n\n⚡ Rasch ball va daraja test yakunlangach hammaga bir vaqtda e'lon qilinadi."
    else:
        text += (f"\n\n🏁 <b>Yakuniy natija:</b> {row['score']}/{row['total_questions']} "
                 f"({float(row['percent'] or 0):.1f}%)")
        if et == "dtm":
            text += f"\n🏁 Ball: <b>{float(row['earned_points'] or 0):.1f}/{float(row['max_points'] or 0):.1f}</b>"
        if et in ("cefr", "ielts", "sat") and row["grade"]:
            text += f"\n🎯 Daraja (taxminiy): <b>{row['grade']}</b>"
    try:
        await bot.send_message(row["telegram_id"], text, parse_mode=ParseMode.HTML)
    except Exception:
        logging.exception(f"O'quvchiga ochiq javob natijasini yuborib bo'lmadi (result={result_id})")


async def _after_open_grading_complete(bot: Bot, result_id: int, info: dict) -> None:
    """Natijaning barcha ochiq javoblari baholangach: o'quvchiga xabar,
    sertifikat (loyiq bo'lsa), Rasch bo'lsa — guruh hisobini ishga tushirish."""
    await _notify_student_open_graded(bot, result_id)
    et = info["exam_type"]
    if et == "rasch":
        await finalize_rush_for_test_if_ready(info["test_id"], bot=bot)
        return
    r = await POOL.fetchrow("SELECT certificate_sent, percent, grade, open_pending FROM results WHERE id=$1", result_id)
    if r and int(r["certificate_sent"] or 0) == 0:
        decision = _cert_decision(et, r["percent"], r["grade"], False, r["open_pending"])
        if decision:
            try:
                await _send_result_certificate(bot, result_id)
            except Exception:
                logging.exception(f"Ochiq javoblardan keyin sertifikat yuborilmadi (result={result_id})")


async def process_test_submission(
    user_tg_id: int, bot: Bot, state: FSMContext, user_ans: str, open_answers: dict | None = None
):
    user_db_id = await get_user_db_id(user_tg_id)
    if not user_db_id:
        await bot.send_message(user_tg_id, "Avval /start orqali ro‘yxatdan o‘ting.")
        await state.clear()
        return

    st = await state.get_data()
    test_id = int(st.get("test_id") or 0)
    mode = st.get("mode") or "free"

    sess = await get_session(user_db_id, test_id, mode)
    if not sess:
        await bot.send_message(user_tg_id, "⚠️ Sessiya topilmadi. Testni qayta boshlang.")
        cancel_timer(user_tg_id, test_id)
        await state.clear()
        return

    start_ts_, duration_sec = sess["start_ts"], sess["duration_sec"]
    end_ts_ = int(start_ts_) + int(duration_sec)
    if now_ts() > end_ts_ + 60:
        await bot.send_message(user_tg_id, "❌ Vaqt tugadi! Javoblar qabul qilinmadi.")
        await delete_session(user_db_id, test_id, mode)
        cancel_timer(user_tg_id, test_id)
        await state.clear()
        return

    if await POOL.fetchrow("SELECT 1 FROM results WHERE user_id=$1 AND test_id=$2 LIMIT 1", user_db_id, test_id):
        await bot.send_message(user_tg_id,
                               "⚠️ Siz bu testni allaqachon topshirgansiz. Qayta topshirib bo‘lmaydi.")
        await delete_session(user_db_id, test_id, mode)
        cancel_timer(user_tg_id, test_id)
        await state.clear()
        return

    t = await POOL.fetchrow("""
        SELECT code, title, subject, questions_count, answers, exam_type, dtm_cfg, start_mode, start_ts, duration,
               cert_academy_name, cert_signee_name, cert_template, creator_id, open_ref_answers
        FROM tests
        WHERE id=$1
    """, test_id)

    if not t:
        await bot.send_message(user_tg_id, "Test topilmadi.")
        await delete_session(user_db_id, test_id, mode)
        cancel_timer(user_tg_id, test_id)
        await state.clear()
        return

    code, title, subject, qcount, answer_key_raw, exam_type, dtm_cfg, start_mode, start_ts_val, duration_min = (
        t["code"], t["title"], t["subject"], t["questions_count"], t["answers"], t["exam_type"], t["dtm_cfg"],
        t["start_mode"], t["start_ts"], t["duration"]
    )
    qcount = int(qcount)
    correct_key = normalize_answers(answer_key_raw)

    if len(user_ans) != qcount:
        await bot.send_message(
            user_tg_id,
            f"⚠️ Javoblar soni noto‘g‘ri.\n"
            f"Kerakli: <b>{qcount}</b>, Siz yubordingiz: <b>{len(user_ans)}</b>.\n\n"
            f"Iltimos, aniq <b>{qcount} ta</b> javob yuboring.",
            parse_mode=ParseMode.HTML
        )
        return

    # ✅ Ochiq (x) savollar: kalit bilan avtomatik solishtiriladi; avtomatik
    # tekshirib bo'lmaydiganlari (rasm/kalitsiz) ustoz tekshiruviga qoladi.
    open_answers = open_answers or {}
    open_refs = _load_json_dict(t["open_ref_answers"])
    open_grades = initial_open_grades(correct_key[:qcount], open_answers, open_refs)
    score, scored_qcount, open_pending = score_with_open(user_ans, correct_key, qcount, open_grades)
    open_count = sum(1 for c in correct_key[:qcount] if c == "x")

    mistakes = []
    for i in range(qcount):
        if correct_key[i] == "x":
            g = open_grades.get(str(i + 1))
            if g is not None and int(g) == 0:
                mistakes.append(f"{i + 1}-savol (ochiq): javobingiz kalitga mos kelmadi | to‘g‘ri={_h(open_refs.get(str(i + 1), '-'))}")
            continue
        if user_ans[i] != correct_key[i]:
            miss_char = "Belgilangan: yo'q" if user_ans[i] == "*" else f"Siz: {user_ans[i]}"
            mistakes.append(f"{i + 1}-savol: {miss_char} | to‘g‘ri={correct_key[i]}")

    percent = (score / scored_qcount * 100) if scored_qcount > 0 else 0.0

    earned_points = float(score)
    max_points = float(scored_qcount)
    details = None

    if (exam_type or "simple") == "dtm":
        earned_points, max_points, details_obj = dtm_score_points(user_ans, correct_key, dtm_cfg or "", open_grades)
        details = json.dumps(details_obj, ensure_ascii=False)
        percent_points = (earned_points / max_points) * 100 if max_points > 0 else 0.0
    else:
        percent_points = percent

    # ✅ CEFR/IELTS/SAT — guruhga bog'liq emas, darajasi/balli darhol foizdan
    # hisoblanadi (rasch'dagi kabi test tugashini kutish shart emas).
    grade_value = grade_value_for_exam_type(exam_type, percent_points)

    open_answers_json = json.dumps(open_answers, ensure_ascii=False) if open_answers else None

    try:
        result_id = await POOL.fetchval("""
            INSERT INTO results (
                user_id, test_id, score, total_questions, percent,
                user_answers, open_text_answers,
                earned_points, max_points, details,
                rasch_ball, rasch_percent, grade,
                finished_ts, date, open_grades, open_pending
            )
            VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, NULL, NULL, $11, $12, $13, $14, $15)
            RETURNING id
        """,
            user_db_id, test_id,
            score, qcount,
            float(percent_points),
            user_ans, open_answers_json,
            float(earned_points), float(max_points), details,
            grade_value,
            now_ts(),
            now_str_local(),
            json.dumps(open_grades, ensure_ascii=False), int(open_pending)
        )
    except asyncpg.UniqueViolationError:
        # Ikkita so'rov bir vaqtda kelib, ikkalasi ham shu yergacha yetib kelgan —
        # DB darajasidagi himoya ishladi, foydalanuvchiga tushunarli xabar beramiz.
        await bot.send_message(user_tg_id, "⚠️ Siz bu testni allaqachon topshirgansiz. Qayta topshirib bo‘lmaydi.")
        await delete_session(user_db_id, test_id, mode)
        cancel_timer(user_tg_id, test_id)
        await state.clear()
        return

    et = (exam_type or "simple")
    if et == "rasch":
        # ✅ Rasch allaqachon hisoblangan bo'lsa (masalan test yopilganda kimdir
        # hali yechayotgan edi) — yangi natija ham guruh hisobiga kirishi uchun
        # hisob qaytadan bajariladi. Aks holda bu natija abadiy "hisoblanmagan"
        # holatda qolardi.
        if await POOL.fetchval("SELECT COALESCE(rush_finalized,0) FROM tests WHERE id=$1", test_id) == 1:
            await POOL.execute("UPDATE tests SET rush_finalized=0 WHERE id=$1", test_id)
            await POOL.execute(
                "UPDATE results SET certificate_sent=0 WHERE test_id=$1 AND certificate_sent=2", test_id)
    can_immediate = ((start_mode or "normal") == "normal") and (et in ("simple", "dtm", "cefr", "ielts", "sat"))
    # ✅ Ochiq javoblar hali tekshirilmagan bo'lsa sertifikatni kutamiz —
    # yakuniy foiz ustoz tekshirgach aniq bo'ladi.
    if can_immediate and open_pending == 0:
        if is_certificate_eligible(et, percent_points, None):
            try:
                await _send_result_certificate(bot, result_id)
            except Exception:
                # Sertifikat darhol yuborilmasa ham imtihon natijasini buzmaymiz —
                # certificate_daemon uni keyinroq certificate_sent=0 bo'lgani uchun
                # baribir topib yuboradi. Sababini kuzatish uchun log qoldiramiz.
                logging.exception(
                    f"Sertifikatni darhol yuborishda xatolik (user_db_id={user_db_id}, test_id={test_id})"
                )

    await delete_session(user_db_id, test_id, mode)
    cancel_timer(user_tg_id, test_id)
    await state.clear()

    if open_pending:
        creator_id = t["creator_id"] or ADMIN_ID
        asyncio.create_task(_send_open_answers_for_review(bot, creator_id, result_id))

    exam_type = (exam_type or "simple")

    # ✅ Rasch — guruh testi ODATDA hammaga bir vaqtda (scheduled)
    # o'tkaziladi, shuning uchun kim qaysi savolda xato qilganini DARHOL
    # ko'rsatish xavfli: erta tugatgan o'quvchi hali yechayotgan sinfdoshlariga
    # to'g'ri javoblarni "sizib chiqarishi" mumkin edi. Shu sabab bu turda
    # batafsil xato tahlili test VAQTI tugagunicha yashiriladi; boshqa barcha
    # turlarda (simple/dtm/cefr/ielts/sat) xatolar avvalgidek darhol chiqadi.
    if exam_type == "rasch":
        remain_text = await remaining_time_text(test_id)
        mistakes_text = (
            f"⏳ Guruh xavfsizligi uchun batafsil xato tahlili test vaqti "
            f"tugagach ko'rsatiladi (\"🔎 Xatolarni ko'rish\" tugmasi orqali).\n"
            f"🕒 <b>{remain_text}</b>"
        )
    else:
        mistakes_text = "\n".join(mistakes[:25])
        if len(mistakes) > 25:
            mistakes_text += "\n..."
        if not mistakes_text:
            mistakes_text = "Xatolar yo‘q ✅"

    open_note = ""
    if open_count:
        if open_pending:
            open_note = (f"✍️ Ochiq savollar: <b>{open_count}</b> ta — <b>{open_pending}</b> tasini ustoz "
                         f"tekshiradi (rasm ko'rinishidagi javoblar). Tekshirilgach yakuniy natija "
                         f"shu yerga yuboriladi.\n")
        elif exam_type != "rasch":
            n_ok = sum(1 for g in open_grades.values() if int(g) == 1)
            open_note = f"✍️ Ochiq savollar: <b>{n_ok}/{open_count}</b> to'g'ri (kalit bilan avtomatik tekshirildi).\n"
        else:
            open_note = f"✍️ Ochiq savollar: <b>{open_count}</b> ta — kalit bilan avtomatik tekshirildi.\n"
    extra_points_line = ""
    if exam_type == "dtm":
        extra_points_line = (
            f"🏁 Ball: <b>{earned_points:.1f}/{max_points:.1f}</b>\n"
            f"📌 Ball foiz: <b>{percent_points:.1f}%</b>\n"
        )

    rasch_pending_block = ""
    if exam_type == "rasch":
        rasch_pending_block = (
            "⚡ Rasch foiz: <b>Test tugagach avtomatik hisoblanadi</b>\n"
            "⚡ Rasch ball: <b>Test tugagach avtomatik hisoblanadi</b>\n"
            "🎯 Daraja: <b>Test tugagach chiqadi</b>\n"
            f"🕒 Qolgan vaqt: <b>{remain_text}</b>\n"
            "📌 Rasch ball ko‘rish uchun: <b>“📄 Natijani ko‘rish”</b> tugmasini bosing.\n"
        )

    intl_block = ""
    if exam_type == "cefr":
        intl_block = (
            f"🌍 CEFR daraja (taxminiy): <b>{grade_value}</b>\n"
            f"<i>Bu rasmiy CEFR sertifikati emas — taxminiy konvertatsiya.</i>\n"
        )
    elif exam_type == "ielts":
        intl_block = (
            f"🇬🇧 IELTS band (taxminiy): <b>{grade_value}</b>\n"
            f"<i>Bu rasmiy IELTS balli emas — taxminiy konvertatsiya.</i>\n"
        )
    elif exam_type == "sat":
        intl_block = (
            f"📝 SAT ball (taxminiy): <b>{grade_value} / 1600</b>\n"
            f"<i>Bu rasmiy SAT balli emas — taxminiy konvertatsiya.</i>\n"
        )

    if exam_type == "rasch":
        score_lines = "📊 Javoblaringiz qabul qilindi.\n"
    else:
        score_lines = (
            f"✅ To‘g‘ri: <b>{score}/{scored_qcount}</b>\n"
            f"❌ Xato: <b>{scored_qcount - score}</b>\n"
            f"📊 Foiz: <b>{percent:.1f}%</b>\n"
        )

    result_text = (
        f"🏁 <b>Test yakunlandi!</b>\n\n"
        f"🔑 Kod: <code>{code}</code>\n"
        f"📚 Test: <b>{_h(title)}</b>\n"
        f"🧮 Savollar: {qcount}\n\n"
        f"{score_lines}"
        f"{open_note}"
        f"{extra_points_line}"
        f"{rasch_pending_block}"
        f"{intl_block}\n"
        f"🔎 <b>Xatolar:</b>\n"
        f"{mistakes_text}\n\n"
        f"ℹ️ Bu test bo‘yicha natija faqat <b>1-urinish</b> hisoblanadi."
    )
    # ✅ Har bir topshiruv uchun adminga alohida xabar YUBORILMAYDI (ko'p
    # o'quvchi bir vaqtda test yechsa, admin xabarlarga to'lib ketardi).
    # Buning o'rniga admin/test yaratuvchisi istalgan vaqt "📊 Statistika" →
    # "🛑 Testlarni yakunlash/qayta ochish" yoki "⏱ Vaqt oralig'ida natijalar"
    # orqali test holatini o'zi tekshiradi va boshqaradi.

    res_kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="📄 Natijani ko‘rish", callback_data=f"viewres_{test_id}")]
    ])

    admin_flag = await is_admin_user(user_tg_id)
    await bot.send_message(user_tg_id, result_text, parse_mode=ParseMode.HTML, reply_markup=res_kb)
    await bot.send_message(user_tg_id, "🏠 Bosh menyu", reply_markup=main_menu_kb(admin_flag))


# ----- Ustoz: ochiq javoblarni baholash tugmalari -----
@router.callback_query(F.data.startswith("og_"))
async def open_grade_cb(callback: CallbackQuery, bot: Bot):
    parts = callback.data.split("_")
    if len(parts) != 4 or not parts[1].isdigit() or not parts[2].isdigit():
        return await callback.answer("Noto'g'ri tugma.", show_alert=True)
    result_id, qnum, val = int(parts[1]), parts[2], (1 if parts[3] == "1" else 0)

    row = await POOL.fetchrow("""
        SELECT r.open_grades, COALESCE(r.open_pending,0) AS open_pending, r.test_id,
               t.creator_id, t.exam_type, COALESCE(t.rush_finalized,0) AS rush_finalized, t.answers
        FROM results r JOIN tests t ON t.id=r.test_id WHERE r.id=$1
    """, result_id)
    if not row:
        await callback.answer("Bu natija endi mavjud emas (o'chirilgan yoki qayta yechilgan).", show_alert=True)
        try:
            await callback.message.edit_reply_markup(reply_markup=None)
        except Exception:
            pass
        return
    if not (await is_admin_user(callback.from_user.id) or row["creator_id"] == callback.from_user.id):
        return await callback.answer("Sizda ruxsat yo'q!", show_alert=True)
    if (row["exam_type"] or "") == "rasch" and int(row["rush_finalized"]) == 1:
        return await callback.answer(
            "Rasch natijalari allaqachon e'lon qilingan — bahoni endi o'zgartirib bo'lmaydi.", show_alert=True)

    key = normalize_answers(row["answers"] or "")
    qi = int(qnum) - 1
    if not (0 <= qi < len(key)) or key[qi] != "x":
        return await callback.answer("Bu savol ochiq savol emas.", show_alert=True)

    grades = _load_json_dict(row["open_grades"])
    old = grades.get(qnum)
    grades[qnum] = val
    info = await _recompute_result_scores(result_id, grades)

    try:
        await callback.message.edit_reply_markup(reply_markup=_open_grade_kb(result_id, qnum, val))
    except Exception:
        pass
    await callback.answer("✅ To'g'ri deb belgilandi" if val else "❌ Noto'g'ri deb belgilandi")

    if info and info["pending"] == 0 and (int(row["open_pending"]) > 0 or old != val):
        await _after_open_grading_complete(bot, result_id, info)


async def _render_open_pending_list(answerable, test_id: int) -> None:
    t = await POOL.fetchrow("SELECT code, title FROM tests WHERE id=$1", test_id)
    if not t:
        await answerable.answer("Test topilmadi.")
        return
    rows = await POOL.fetch("""
        SELECT r.id, u.full_name, r.open_pending FROM results r JOIN users u ON u.id=r.user_id
        WHERE r.test_id=$1 AND COALESCE(r.open_pending,0)>0 ORDER BY r.id ASC LIMIT 40
    """, test_id)
    if not rows:
        await answerable.answer("✅ Bu testda tekshirilmagan ochiq javob yo'q.")
        return
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=f"👤 {r['full_name'] or 'Noma’lum'} — {r['open_pending']} ta"[:60],
                              callback_data=f"ogr_{r['id']}")]
        for r in rows
    ])
    await answerable.answer(
        f"✍️ <b>{_h(t['title'])}</b> (<code>{t['code']}</code>) — tekshirilmagan ochiq javoblar:",
        parse_mode=ParseMode.HTML, reply_markup=kb)


@router.callback_query(F.data.startswith("ogl_"))
async def open_pending_list_cb(callback: CallbackQuery):
    test_id = int(callback.data.split("_")[1])
    if not await _is_admin_or_creator(callback.from_user.id, test_id):
        return await callback.answer("Sizda ruxsat yo'q!", show_alert=True)
    await _render_open_pending_list(callback.message, test_id)


@router.callback_query(F.data.startswith("ogr_"))
async def open_pending_result_cb(callback: CallbackQuery, bot: Bot):
    result_id = int(callback.data.split("_")[1])
    test_id = await POOL.fetchval("SELECT test_id FROM results WHERE id=$1", result_id)
    if not test_id:
        return await callback.answer("Natija topilmadi.", show_alert=True)
    if not await _is_admin_or_creator(callback.from_user.id, test_id):
        return await callback.answer("Sizda ruxsat yo'q!", show_alert=True)
    n = await _send_open_answers_for_review(bot, callback.from_user.id, result_id, only_pending=True)
    if not n:
        await callback.message.answer("✅ Bu o'quvchining barcha ochiq javoblari allaqachon baholangan.")


@router.callback_query(F.data.startswith("oga_"))
async def open_all_list_cb(callback: CallbackQuery):
    """Ixtiyoriy: avtomatik tekshirilganlari bilan birga BARCHA ochiq javoblar —
    ustoz xohlasa avtomatik bahoni ko'rib, o'zgartirishi mumkin."""
    test_id = int(callback.data.split("_")[1])
    if not await _is_admin_or_creator(callback.from_user.id, test_id):
        return await callback.answer("Sizda ruxsat yo'q!", show_alert=True)
    rows = await POOL.fetch("""
        SELECT r.id, u.full_name FROM results r JOIN users u ON u.id=r.user_id
        WHERE r.test_id=$1 AND r.open_text_answers IS NOT NULL ORDER BY r.id DESC LIMIT 40
    """, test_id)
    if not rows:
        return await callback.message.answer("Bu testda hali ochiq javob yozganlar yo'q.")
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=f"👤 {r['full_name'] or 'Noma’lum'}"[:60], callback_data=f"ogra_{r['id']}")]
        for r in rows
    ])
    await callback.message.answer("🔍 Kimning ochiq javoblarini ko'rmoqchisiz? (bahoni o'zgartirish ham mumkin)",
                                  reply_markup=kb)


@router.callback_query(F.data.startswith("ogra_"))
async def open_all_result_cb(callback: CallbackQuery, bot: Bot):
    result_id = int(callback.data.split("_")[1])
    test_id = await POOL.fetchval("SELECT test_id FROM results WHERE id=$1", result_id)
    if not test_id:
        return await callback.answer("Natija topilmadi.", show_alert=True)
    if not await _is_admin_or_creator(callback.from_user.id, test_id):
        return await callback.answer("Sizda ruxsat yo'q!", show_alert=True)
    n = await _send_open_answers_for_review(bot, callback.from_user.id, result_id, only_pending=False)
    if not n:
        await callback.message.answer("Bu o'quvchi ochiq savollarga javob yozmagan.")


@router.callback_query(F.data == "ogtests")
async def open_pending_tests_cb(callback: CallbackQuery):
    is_admin = await is_admin_user(callback.from_user.id)
    rows = await POOL.fetch("""
        SELECT t.id, t.code, t.title, COUNT(*) AS n
        FROM results r JOIN tests t ON t.id=r.test_id
        WHERE COALESCE(r.open_pending,0)>0 AND ($1 OR t.creator_id=$2)
        GROUP BY t.id, t.code, t.title ORDER BY t.id DESC LIMIT 40
    """, is_admin, callback.from_user.id)
    if not rows:
        return await callback.message.answer("✅ Tekshirilishi kerak bo'lgan ochiq javoblar yo'q.")
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=f"✍️ {r['code']} — {r['title']} ({r['n']})"[:60], callback_data=f"ogl_{r['id']}")]
        for r in rows
    ])
    await callback.message.answer("✍️ <b>Ochiq javoblari tekshirilmagan testlar:</b>",
                                  parse_mode=ParseMode.HTML, reply_markup=kb)


@router.callback_query(F.data.startswith("ogforce_"))
async def open_force_ask(callback: CallbackQuery):
    test_id = int(callback.data.split("_")[1])
    if not await _is_admin_or_creator(callback.from_user.id, test_id):
        return await callback.answer("Sizda ruxsat yo'q!", show_alert=True)
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="⚡ Ha, hisoblansin", callback_data=f"ogforceok_{test_id}")],
        [InlineKeyboardButton(text="❌ Bekor qilish", callback_data="dismiss_msg")],
    ])
    await callback.message.answer(
        "⚠️ Tekshirilmagan ochiq javoblar <b>0 (xato)</b> deb hisoblanadi va Rasch natijalari darhol "
        "e'lon qilinadi. Bu amalni ortga qaytarib bo'lmaydi. Davom etasizmi?",
        parse_mode=ParseMode.HTML, reply_markup=kb)


@router.callback_query(F.data.startswith("ogforceok_"))
async def open_force_confirm(callback: CallbackQuery, bot: Bot):
    test_id = int(callback.data.split("_")[1])
    if not await _is_admin_or_creator(callback.from_user.id, test_id):
        return await callback.answer("Sizda ruxsat yo'q!", show_alert=True)
    ok = await finalize_rush_for_test_if_ready(test_id, bot=bot, force=True)
    if ok:
        await callback.message.answer("✅ Rasch natijalari hisoblandi va e'lon qilindi.")
    else:
        await callback.message.answer("⏳ Test vaqti hali tugamagan — natijalar test tugagach hisoblanadi.")


# Eski matn (Text) qabul qilish formati uchun
@router.message(TestProcessState.solving)
async def submit_answers_text(message: Message, state: FSMContext, bot: Bot):
    if (message.text or "") in TOP_MENU_BTNS: return
    user_ans = normalize_answers(message.text)

    if not is_answer_string_valid(user_ans):
        await message.answer(
            "⚠️ Javoblar faqat <b>A, B, C, D, E, F</b> harflaridan iborat bo'lsin. Masalan: <code>abcdabcd...</code>",
            parse_mode=ParseMode.HTML)
        return

    # ✅ Testda ochiq (variantsiz) savollar bo'lsa, ularning javobini matn
    # qatori orqali berib bo'lmaydi — aks holda ular hech qanday javobsiz
    # (avtomatik xato) qolib ketardi. O'quvchini tugmalarga yo'naltiramiz.
    st = await state.get_data()
    key = st.get("test_answers_key") or ""
    if "x" in key:
        nums = ", ".join(str(p) for p in open_question_positions(key))
        await message.answer(
            f"✍️ Bu testda <b>ochiq savollar</b> bor ({nums}-savol) — ularga javobni matn bilan yozish kerak.\n"
            f"Iltimos, javoblarni yuqoridagi tugmalar orqali belgilang (ochiq savolda «✍️ Matn bilan javob "
            f"yozish» tugmasi chiqadi), so'ng «✅ Yakunlash»ni bosing.",
            parse_mode=ParseMode.HTML)
        return

    await process_test_submission(message.from_user.id, bot, state, user_ans)


# Yangi interaktiv (Tugma) formati uchun
# ✅ awaiting_free_text holati ham qo'shildi: o'quvchi «✍️ Matn bilan javob
# yozish»ni bosib, keyin fikridan qaytib «Keyingi ➡️» yoki «✅ Yakunlash»ni
# bossa, avval bu tugmalar umuman ishlamasdi (hech qanday handler mos
# kelmasdi, tugma "osilib" qolardi).
@router.callback_query(StateFilter(TestProcessState.solving, TestProcessState.awaiting_free_text),
                       F.data.startswith("t_"))
async def interactive_test_handler(callback: CallbackQuery, state: FSMContext, bot: Bot):
    if callback.data != "t_freetext" and await state.get_state() == TestProcessState.awaiting_free_text.state:
        await state.set_state(TestProcessState.solving)
    data = await state.get_data()
    current_q = data.get("current_q", 1)
    qcount = int(data.get("qcount", 30))
    answers = data.get("test_answers", {})
    correct_key = data.get("test_answers_key", "")

    action = callback.data

    if action == "t_nav_noop":
        return await callback.answer(f"Siz hozir {current_q}-savoldasiz", show_alert=False)

    elif action.startswith("t_ans_"):
        # ✅ Avval javob belgilash bilan bir vaqtda AVTOMATIK keyingi savolga
        # o'tkazardi — bu shu qadar sezilmas edi-ki, foydalanuvchi "3-savolga
        # o'tib ketibman"ni payqamay qolardi. Endi tugma bosilganda FAQAT javob
        # belgilanadi (yashil ✅ bilan) va aniq bildirishnoma (toast) chiqadi;
        # keyingi savolga esa faqat "Keyingi ➡️" tugmasi orqali, foydalanuvchi
        # o'zi bosib o'tadi.
        ans_letter = action.split("_")[2]
        answers[str(current_q)] = ans_letter
        await state.update_data(test_answers=answers)
        try:
            await callback.message.edit_reply_markup(reply_markup=test_ui_kb(current_q, qcount, answers, correct_key))
        except Exception:
            pass
        await callback.answer(f"✅ {ans_letter.upper()} belgilandi", show_alert=False)

    elif action == "t_nav_prev":
        if current_q > 1:
            current_q -= 1
            await state.update_data(current_q=current_q)
            try:
                await callback.message.edit_text(
                    text=f"📝 {question_status_note(correct_key, current_q)}<b>{current_q}-savol</b> / {qcount}\n"
                         f"Pastdagi tugmalar orqali javobni belgilang:",
                    parse_mode=ParseMode.HTML,
                    reply_markup=test_ui_kb(current_q, qcount, answers, correct_key)
                )
            except Exception:
                pass
        await callback.answer(f"⬅️ {current_q}-savolga o'tdingiz", show_alert=False)

    elif action == "t_nav_next":
        if current_q < qcount:
            current_q += 1
            await state.update_data(current_q=current_q)
            try:
                await callback.message.edit_text(
                    text=f"📝 {question_status_note(correct_key, current_q)}<b>{current_q}-savol</b> / {qcount}\n"
                         f"Pastdagi tugmalar orqali javobni belgilang:",
                    parse_mode=ParseMode.HTML,
                    reply_markup=test_ui_kb(current_q, qcount, answers, correct_key)
                )
            except Exception:
                pass
        await callback.answer(f"➡️ {current_q}-savolga o'tdingiz", show_alert=False)

    elif action == "t_freetext":
        await callback.answer()
        await callback.message.answer(
            f"✏️ <b>{current_q}-savol</b> uchun javobingizni matn qilib yozing "
            f"(variantsiz/ochiq savollar uchun — ustoz keyinroq tekshiradi):",
            parse_mode=ParseMode.HTML
        )
        await state.set_state(TestProcessState.awaiting_free_text)

    elif action == "t_submit":
        user_ans = ""
        for i in range(1, qcount + 1):
            user_ans += answers.get(str(i), "*")

        try:
            await callback.message.delete_reply_markup()
        except TelegramBadRequest:
            pass  # klaviatura allaqachon olingan / xabar eski — topshirish to'xtamasligi kerak
        await callback.answer("Javoblaringiz qabul qilindi, baholanmoqda...", show_alert=False)
        open_answers = data.get("test_open_answers", {})
        await process_test_submission(callback.from_user.id, bot, state, user_ans, open_answers)


@router.message(TestProcessState.awaiting_free_text)
async def collect_free_text_answer(message: Message, state: FSMContext) -> None:
    if (message.text or "") in TOP_MENU_BTNS:
        return

    photo_id = message.photo[-1].file_id if message.photo else None
    text_val = sanitize_plain_text(message.text or message.caption or "", max_len=500)
    if not photo_id and len(text_val) < 1:
        await message.answer("⚠️ Javob bo'sh bo'lmasin. Matn yozing yoki rasm (masalan qo'lda yechilgan misol) yuboring:")
        return

    data = await state.get_data()
    current_q = data.get("current_q", 1)
    qcount = int(data.get("qcount", 30))
    answers = data.get("test_answers", {})
    open_answers = data.get("test_open_answers", {})
    correct_key = data.get("test_answers_key", "")

    answers[str(current_q)] = "x"
    open_answers[str(current_q)] = {"text": text_val or None, "photo_id": photo_id}
    await state.update_data(test_answers=answers, test_open_answers=open_answers)
    await state.set_state(TestProcessState.solving)

    ack = "rasm" if photo_id else "matn"
    await message.answer(
        f"✅ {current_q}-savol uchun {ack} javobingiz qabul qilindi.\n\n"
        f"📝 <b>{current_q}-savol</b> / {qcount}\nPastdagi tugmalar orqali davom eting:",
        parse_mode=ParseMode.HTML,
        reply_markup=test_ui_kb(current_q, qcount, answers, correct_key)
    )


# ===================== RESULTS MENU =====================
@router.message(F.text == BTN_RESULTS)
async def my_results(message: Message):
    user_db_id = await get_user_db_id(message.from_user.id)
    if not user_db_id:
        return await message.answer("Avval /start orqali ro‘yxatdan o‘ting.")

    rows = await POOL.fetch("""
        SELECT t.id, t.code, t.title, r.score, r.total_questions, r.percent, r.date
        FROM results r
        JOIN tests t ON t.id=r.test_id
        WHERE r.user_id=$1
        ORDER BY r.id DESC
        LIMIT 10
    """, user_db_id)

    if not rows:
        return await message.answer("Sizda hali natijalar yo‘q.")

    # ✅ Shaxsiy statistika: nechta test yechgan, o'rtacha va eng yaxshi natija
    summary = await POOL.fetchrow("""
        SELECT COUNT(*) AS cnt, AVG(percent) AS avgp, MAX(percent) AS maxp
        FROM results WHERE user_id=$1
    """, user_db_id)

    text = (
        "<b>📊 Mening statistikam</b>\n"
        f"🧮 Jami topshirilgan: <b>{summary['cnt']}</b> ta\n"
        f"📈 O'rtacha natija: <b>{float(summary['avgp'] or 0):.1f}%</b>\n"
        f"🏅 Eng yaxshi natija: <b>{float(summary['maxp'] or 0):.1f}%</b>\n\n"
        "<b>📄 So'nggi natijalar (oxirgi 10)</b>\n\n"
    )
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="📜 Sertifikatlarim", callback_data="my_certs")]
    ])

    for r in rows:
        test_id, code, title, score, total, pct, dt = (
            r["id"], r["code"], r["title"], r["score"], r["total_questions"], r["percent"], r["date"]
        )
        text += (
            f"🔹 <code>{code}</code> {title}\n"
            f"   ✅ {score}/{total} | 📊 <b>{float(pct):.1f}%</b> | 🗓 {dt}\n\n"
        )
        kb.inline_keyboard.append([
            InlineKeyboardButton(text=f"📄 Ochish: {code}", callback_data=f"viewres_{test_id}")
        ])

    await message.answer(text, parse_mode=ParseMode.HTML, reply_markup=kb)


@router.callback_query(F.data == "my_certs")
async def my_certificates(callback: CallbackQuery):
    user_db_id = await get_user_db_id(callback.from_user.id)
    if not user_db_id:
        return await callback.message.answer("Avval /start orqali ro‘yxatdan o‘ting.")

    rows = await POOL.fetch("""
        SELECT r.id, t.code, t.title, r.percent, r.date
        FROM results r JOIN tests t ON t.id=r.test_id
        WHERE r.user_id=$1 AND r.certificate_sent=1 AND r.certificate_file_id IS NOT NULL
        ORDER BY r.id DESC
        LIMIT 30
    """, user_db_id)

    if not rows:
        return await callback.message.answer(
            "📜 Sizda hali sertifikat yo'q. Testdan 50% va undan yuqori natija bilan o'tsangiz, "
            "sertifikat avtomatik beriladi."
        )

    kb = InlineKeyboardMarkup(inline_keyboard=[])
    for r in rows:
        kb.inline_keyboard.append([InlineKeyboardButton(
            text=f"📜 {r['code']} — {float(r['percent']):.1f}%", callback_data=f"certget_{r['id']}")])
    await callback.message.answer("📜 <b>Sertifikatlaringiz</b>\nBirini tanlang — qayta yuboraman:",
                                  parse_mode=ParseMode.HTML, reply_markup=kb)


@router.callback_query(F.data.startswith("certget_"))
async def my_certificate_resend(callback: CallbackQuery, bot: Bot):
    user_db_id = await get_user_db_id(callback.from_user.id)
    if not user_db_id:
        return await callback.message.answer("Avval /start orqali ro‘yxatdan o‘ting.")

    result_id = int(callback.data.split("_")[1])
    row = await POOL.fetchrow("""
        SELECT certificate_file_id, t.title
        FROM results r JOIN tests t ON t.id=r.test_id
        WHERE r.id=$1 AND r.user_id=$2
    """, result_id, user_db_id)

    if not row or not row["certificate_file_id"]:
        return await callback.message.answer("❌ Sertifikat topilmadi.")

    await bot.send_document(
        chat_id=callback.from_user.id,
        document=row["certificate_file_id"],
        caption=f"📜 {row['title']}"
    )


# ===================== RATING =====================
@router.message(F.text == BTN_RATING)
async def rating_menu(message: Message):
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="🏆 Umumiy reyting (TOP)", callback_data="rt_overall")],
        [InlineKeyboardButton(text="📅 Haftalik TOP (7 kun)", callback_data="rt_weekly")],
        [InlineKeyboardButton(text="📚 Test bo‘yicha reyting", callback_data="rt_tests")]
    ])
    await message.answer("🏆 <b>Reyting</b>\nBo‘limni tanlang:", parse_mode=ParseMode.HTML, reply_markup=kb)


@router.callback_query(F.data == "rt_overall")
async def rating_overall(callback: CallbackQuery):
    user_db_id = await get_user_db_id(callback.from_user.id)
    if not user_db_id:
        return await callback.message.answer("Avval /start orqali ro‘yxatdan o‘ting.")

    # ✅ Hisoblash SQL tomonda (RANK() OVER ...) bajariladi — 10 000+
    # foydalanuvchi/natija bo'lganda ham butun jadval Python xotirasiga
    # yuklanmaydi, faqat TOP_N + so'rovchining o'z qatori qaytariladi.
    rows = await POOL.fetch("""
        WITH agg AS (
            SELECT user_id, AVG(percent) AS avgp, COUNT(*) AS cnt
            FROM results
            GROUP BY user_id
        ), ranked AS (
            SELECT user_id, avgp, cnt, RANK() OVER (ORDER BY avgp DESC) AS rnk
            FROM agg
        )
        SELECT rk.user_id, rk.avgp, rk.cnt, rk.rnk, COALESCE(u.full_name, 'Noma''lum') AS full_name
        FROM ranked rk
        JOIN users u ON u.id = rk.user_id
        WHERE rk.rnk <= $1 OR rk.user_id = $2
        ORDER BY rk.rnk ASC
    """, TOP_N, user_db_id)

    if not rows:
        return await callback.message.answer("Hozircha reyting uchun natijalar yo‘q.")

    text = "<b>🏆 Umumiy reyting (TOP)</b>\n(natijalar bo‘yicha o‘rtacha %)\n\n"
    my_row = None
    for r in rows:
        if r["rnk"] <= TOP_N:
            text += f"{r['rnk']}) {r['full_name']} — <b>{float(r['avgp']):.1f}%</b> (testlar: {r['cnt']})\n"
        if r["user_id"] == user_db_id:
            my_row = r

    if my_row:
        if my_row["rnk"] > TOP_N:
            text += f"\n<b>Sizning o‘rningiz:</b> {my_row['rnk']}-o‘rin — {float(my_row['avgp']):.1f}%"
    else:
        text += "\nSiz hali reytingga kirmagansiz."

    await callback.message.answer(text, parse_mode=ParseMode.HTML)


@router.callback_query(F.data == "rt_weekly")
async def rating_weekly(callback: CallbackQuery):
    cutoff = (datetime.now(TZ) - timedelta(days=7)).strftime("%Y-%m-%d %H:%M")
    rows = await POOL.fetch("""
        SELECT COALESCE(u.full_name, 'Noma''lum') AS full_name, MAX(r.percent) AS bestp
        FROM results r
        JOIN users u ON u.id = r.user_id
        WHERE r.date >= $1
        GROUP BY u.id, u.full_name
        ORDER BY bestp DESC
        LIMIT $2
    """, cutoff, TOP_N)

    if not rows:
        return await callback.message.answer("Oxirgi 7 kunda reyting uchun natija yo‘q.")

    text = "<b>📅 Haftalik TOP (7 kun)</b>\n(eng yaxshi natija)\n\n"
    for i, r in enumerate(rows, start=1):
        text += f"{i}) {r['full_name']} — <b>{float(r['bestp']):.1f}%</b>\n"

    await callback.message.answer(text, parse_mode=ParseMode.HTML)


async def _render_rating_tests_list(answerable, user_tg_id: int, page: int = 0):
    offset = max(0, page) * USER_LIST_PAGE_SIZE
    tests = await POOL.fetch(
        "SELECT id, code, title FROM tests ORDER BY id DESC LIMIT $1 OFFSET $2",
        USER_LIST_PAGE_SIZE + 1, offset
    )
    has_more = len(tests) > USER_LIST_PAGE_SIZE
    tests = tests[:USER_LIST_PAGE_SIZE]

    if not tests:
        if page == 0:
            return await answerable.answer("Hozircha testlar yo‘q.")
        return await answerable.answer("Boshqa test yo'q.")

    kb = InlineKeyboardMarkup(inline_keyboard=[])
    for r in tests:
        kb.inline_keyboard.append([InlineKeyboardButton(text=f"🔑 {r['code']} — {r['title']}", callback_data=f"rt_test_{r['id']}")])

    nav = _user_pagination_nav_row("rttests", page, has_more)
    if nav:
        kb.inline_keyboard.append(nav)

    await answerable.answer(f"📚 <b>Test bo‘yicha reyting</b> ({page + 1}-sahifa)\nTestni tanlang:",
                            parse_mode=ParseMode.HTML, reply_markup=kb)


@router.callback_query(F.data == "rt_tests")
async def rating_tests_list(callback: CallbackQuery):
    await _render_rating_tests_list(callback.message, callback.from_user.id, 0)


@router.callback_query(F.data.startswith("rt_test_"))
async def rating_per_test(callback: CallbackQuery, bot: Bot):
    test_id = int(callback.data.split("_")[2])

    t = await POOL.fetchrow("SELECT code, title, exam_type FROM tests WHERE id=$1", test_id)
    if not t:
        return await callback.message.answer("Test topilmadi.")

    et = (t["exam_type"] or "simple")
    if et == "rasch":
        await finalize_rush_for_test_if_ready(test_id, bot=bot)

    rows = await POOL.fetch("""
        SELECT COALESCE(u.full_name, 'Noma''lum') AS full_name, r.percent
        FROM results r
        JOIN users u ON u.id = r.user_id
        WHERE r.test_id=$1
        ORDER BY r.percent DESC
        LIMIT $2
    """, test_id, TOP_N)

    if not rows:
        return await callback.message.answer("Bu test bo‘yicha hali natija yo‘q.")

    text = f"<b>🏆 Reyting: {t['title']}</b>\n🔑 <code>{t['code']}</code>\n\n"
    for i, r in enumerate(rows, start=1):
        text += f"{i}) {r['full_name']} — <b>{float(r['percent']):.1f}%</b>\n"

    await callback.message.answer(text, parse_mode=ParseMode.HTML)


# ===================== ADMIN: PENDING =====================
@router.message(F.text == BTN_ADMIN_PENDING)
async def admin_pending(message: Message):
    if not await is_admin_user(message.from_user.id):
        return

    rows = await POOL.fetch("""
        SELECT p.id, u.full_name, t.code, t.title, t.price, p.created_at
        FROM payments p
        JOIN users u ON u.id=p.user_id
        JOIN tests t ON t.id=p.test_id
        WHERE p.status='pending'
        ORDER BY p.id DESC
        LIMIT 30
    """)

    if not rows:
        return await message.answer("🟡 Pending to‘lovlar yo‘q.")

    text = "<b>🟡 Pending to‘lovlar (oxirgi 30 ta)</b>\n\n"
    kb = InlineKeyboardMarkup(inline_keyboard=[])
    for r in rows:
        pid, uname, code, title, price, created_at = (
            r["id"], r["full_name"], r["code"], r["title"], r["price"], r["created_at"]
        )
        text += (
            f"🆔 <code>{pid}</code>\n"
            f"👤 {_h(uname)}\n"
            f"🔑 <code>{code}</code> — <b>{_h(title)}</b>\n"
            f"💰 {price} so'm | 🕒 {created_at}\n\n"
        )
        kb.inline_keyboard.append([InlineKeyboardButton(text=f"Ko‘rish: {pid}", callback_data=f"pview_{pid}")])

    await message.answer(text, parse_mode=ParseMode.HTML, reply_markup=kb)


@router.callback_query(F.data.startswith("pview_"))
async def admin_pending_view(callback: CallbackQuery, bot: Bot):
    if not await is_admin_user(callback.from_user.id):
        return

    pid = int(callback.data.split("_")[1])

    row = await POOL.fetchrow("""
        SELECT p.id, p.screenshot_id, p.created_at, u.full_name, t.code, t.title, t.price
        FROM payments p
        JOIN users u ON u.id=p.user_id
        JOIN tests t ON t.id=p.test_id
        WHERE p.id=$1
    """, pid)

    if not row:
        return await callback.message.answer("Payment topilmadi.")

    pid, screenshot_id, created_at, uname, code, title, price = (
        row["id"], row["screenshot_id"], row["created_at"], row["full_name"], row["code"], row["title"], row["price"]
    )
    admin_kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="✅ Tasdiqlash", callback_data=f"appr_{pid}")],
        [InlineKeyboardButton(text="❌ Rad etish (Sabab)", callback_data=f"rejask_{pid}")]
    ])

    await bot.send_photo(
        chat_id=callback.from_user.id,
        photo=screenshot_id,
        caption=(
            f"🟡 <b>PENDING to‘lov</b>\n\n"
            f"🆔 Payment ID: <code>{pid}</code>\n"
            f"👤 User: {_h(uname)}\n"
            f"📚 Test: {_h(title)}\n"
            f"🔑 Kod: <code>{code}</code>\n"
            f"💰 Summa: {price} so'm\n"
            f"🕒 Sana: {created_at}"
        ),
        parse_mode=ParseMode.HTML,
        reply_markup=admin_kb
    )


@router.callback_query(F.data.startswith("appr_"))
async def approve_payment(callback: CallbackQuery, bot: Bot):
    if not await is_admin_user(callback.from_user.id):
        return

    payment_id = int(callback.data.split("_")[1])

    row = await POOL.fetchrow("""
        SELECT p.status, u.telegram_id, t.code, t.title, t.duration, t.questions_count,
               t.start_mode, t.start_ts, t.start_at
        FROM payments p
        JOIN users u ON u.id=p.user_id
        JOIN tests t ON t.id=p.test_id
        WHERE p.id=$1
    """, payment_id)

    if not row:
        return await callback.message.answer("Xatolik: payment topilmadi.")
    if row["status"] == "approved":
        return await callback.answer("Bu to'lov allaqachon tasdiqlangan.", show_alert=True)

    user_tg_id, code, title, duration, qcount, start_mode, start_ts_val, start_at = (
        row["telegram_id"], row["code"], row["title"], row["duration"], row["questions_count"],
        row["start_mode"], row["start_ts"], row["start_at"]
    )
    # ✅ Vaqt tekshiruvi tasdiqlashdan OLDIN — avval to'lov "approved" qilib
    # qo'yilib, keyin "tasdiqlab bo'lmaydi" deyilardi (holat ziddiyatli qolardi).
    expired, after = is_scheduled_expired(start_mode, start_ts_val, duration)
    if expired:
        await callback.message.answer("❌ Bu test vaqti tugagan. Tasdiqlab bo‘lmaydi.")
        return

    # ✅ Atomik: ikki admin (yoki tugmani ikki marta bosish) bir vaqtda
    # tasdiqlasa, o'quvchiga ikki marta "tasdiqlandi" borardi. Shuningdek,
    # o'quvchining shu testga boshqa tasdiqlangan to'lovi bo'lsa (eski rad
    # etilgan chek keyin tasdiqlansa) — ikkinchi marta pul hisoblanmaydi.
    upd = await POOL.execute("""
        UPDATE payments SET status='approved', reject_reason=NULL, approved_at=$2
        WHERE id=$1 AND status<>'approved'
          AND NOT EXISTS (SELECT 1 FROM payments p2 WHERE p2.user_id=payments.user_id
                          AND p2.test_id=payments.test_id AND p2.status='approved')
    """, payment_id, now_str_local())
    if upd.endswith(" 0"):
        return await callback.answer("Bu to'lov (yoki shu o'quvchining shu testga boshqa to'lovi) "
                                     "allaqachon tasdiqlangan.", show_alert=True)

    start_kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="🚀 Testni boshlash", callback_data=f"begin_{payment_id}")]
    ])

    extra = ""
    if (start_mode or "normal") == "scheduled":
        extra = f"\n📅 Boshlanish: <b>{start_at}</b>\n⚠️ PDF faqat vaqt yetganda beriladi."

    try:
        await bot.send_message(
            chat_id=user_tg_id,
            text=(
                f"✅ <b>To‘lov tasdiqlandi!</b>\n\n"
                f"📚 Test: <b>{title}</b>\n"
                f"🔑 Kod: <code>{code}</code>\n"
                f"🧮 Savollar: <b>{qcount}</b>\n"
                f"⏳ Vaqt: {duration} daqiqa\n"
                f"{extra}\n\n"
                f"✅ Boshlash uchun tugmani bosing."
            ),
            parse_mode=ParseMode.HTML,
            reply_markup=start_kb
        )
    except (TelegramForbiddenError, TelegramBadRequest):
        logging.warning(f"Tasdiq xabari o'quvchiga yetmadi (payment={payment_id})")
        return await callback.message.answer(
            f"✅ Tasdiqlandi. Payment ID: {payment_id}\n"
            f"⚠️ O'quvchiga xabar yetkazilmadi (botni bloklagan bo'lishi mumkin). Test uning "
            f"«📦 Pullik testlarim» bo'limida ochiq turibdi.")

    await callback.message.answer(f"✅ Tasdiqlandi. Payment ID: {payment_id}")


@router.callback_query(F.data.startswith("rejask_"))
async def reject_ask_reason(callback: CallbackQuery, state: FSMContext):
    if not await is_admin_user(callback.from_user.id):
        return
    payment_id = int(callback.data.split("_")[1])
    cur = await POOL.fetchval("SELECT status FROM payments WHERE id=$1", payment_id)
    if cur == "approved":
        return await callback.answer(
            "Bu to'lov allaqachon TASDIQLANGAN — o'quvchi testdan foydalanmoqda, rad etib bo'lmaydi.",
            show_alert=True)
    await state.update_data(payment_id=payment_id)
    await callback.message.answer(
        "❌ <b>Rad etish sababi</b>ni yozing (1-2 jumla):",
        parse_mode=ParseMode.HTML,
        reply_markup=nav_kb(include_back=True)
    )
    await state.set_state(AdminRejectState.waiting_reason)


@router.message(AdminRejectState.waiting_reason)
async def reject_with_reason(message: Message, state: FSMContext, bot: Bot):
    if not await is_admin_user(message.from_user.id):
        await state.clear()
        return
    if (message.text or "") in TOP_MENU_BTNS:
        return

    data = await state.get_data()
    payment_id = int(data.get("payment_id") or 0)
    # ✅ Sabab keyinchalik o'quvchiga HTML xabarlar ichida ko'rsatiladi — avval
    # tozalanmagani uchun admin "<" yoki "&" yozsa, o'quvchiga xabar umuman
    # bormasdi VA uning "📦 Pullik testlarim" bo'limi doimiy ishlamay qolardi.
    reason = sanitize_plain_text(message.text or "", max_len=300)
    if not reason:
        return await message.answer("Sabab bo‘sh bo‘lmasin. Qayta yozing:")

    upd = await POOL.execute(
        "UPDATE payments SET status='rejected', reject_reason=$1 WHERE id=$2 AND status<>'approved'",
        reason, payment_id)
    if upd.endswith(" 0"):
        await state.clear()
        return await message.answer("⚠️ Bu to'lov allaqachon tasdiqlangan — rad etilmadi.")

    # ✅ Holat yuborishdan OLDIN tozalanadi — avval o'quvchi botni bloklagan
    # bo'lsa xato tufayli admin "sabab kutish" holatida qolib ketardi va uning
    # keyingi har bir xabari yana shu to'lovga "sabab" bo'lib yozilardi.
    await state.clear()
    row = await POOL.fetchrow("""
        SELECT u.telegram_id, t.code, t.title, t.price
        FROM payments p
        JOIN users u ON u.id=p.user_id
        JOIN tests t ON t.id=p.test_id
        WHERE p.id=$1
    """, payment_id)

    delivered = True
    if row:
        user_tg_id, code, title, price = row["telegram_id"], row["code"], row["title"], row["price"]
        try:
            await bot.send_message(
                chat_id=user_tg_id,
                text=(
                    f"❌ <b>To‘lov rad etildi</b>\n\n"
                    f"📚 Test: <b>{title}</b>\n"
                    f"🔑 Kod: <code>{code}</code>\n"
                    f"💰 Summa: {price} so'm\n"
                    f"📝 Sabab: <i>{_h(reason)}</i>\n\n"
                    f"📦 'Pullik testlarim' bo‘limidan qayta to‘lov yuborishingiz mumkin."
                ),
                parse_mode=ParseMode.HTML
            )
        except (TelegramForbiddenError, TelegramBadRequest):
            delivered = False
            logging.warning(f"Rad etish xabari o'quvchiga yetmadi (payment={payment_id})")

    await message.answer(f"❌ Rad etildi. Payment ID: {payment_id}"
                         + ("" if delivered else "\n⚠️ O'quvchiga xabar yetkazilmadi (botni bloklagan bo'lishi mumkin)."))


# ===================== ADMIN: STATS + EXCEL EXPORT =====================
@router.message(F.text == BTN_ADMIN_STATS)
async def admin_stats(message: Message):
    if not await is_admin_user(message.from_user.id):
        return

    users_count = await POOL.fetchval("SELECT COUNT(*) FROM users")
    tests_count = await POOL.fetchval("SELECT COUNT(*) FROM tests")
    free_count = await POOL.fetchval("SELECT COUNT(*) FROM tests WHERE is_free=1")
    paid_count = await POOL.fetchval("SELECT COUNT(*) FROM tests WHERE is_free=0")

    pending_count = await POOL.fetchval("SELECT COUNT(*) FROM payments WHERE status='pending'")
    approved_count = await POOL.fetchval("SELECT COUNT(*) FROM payments WHERE status='approved'")
    rejected_count = await POOL.fetchval("SELECT COUNT(*) FROM payments WHERE status='rejected'")

    overall_avg = float(await POOL.fetchval("SELECT COALESCE(AVG(percent), 0) FROM results") or 0.0)

    best = await POOL.fetchrow("""
        SELECT u.full_name, t.code, t.title, r.percent, r.date
        FROM results r
        JOIN users u ON u.id=r.user_id
        JOIN tests t ON t.id=r.test_id
        ORDER BY r.percent DESC
        LIMIT 1
    """)

    total_revenue = await POOL.fetchval("""
        SELECT COALESCE(SUM(COALESCE(p.amount, t.price)), 0)
        FROM payments p JOIN tests t ON t.id=p.test_id WHERE p.status='approved'
    """)

    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="📤 Excel eksport (test bo‘yicha)", callback_data="xl_menu"),
         InlineKeyboardButton(text="⚡ Rasch hisobot", callback_data="rasch_menu")],
        [InlineKeyboardButton(text="⏱ Vaqt oralig'ida natijalar", callback_data="trange_menu"),
         InlineKeyboardButton(text="🟢 Hozir faol foydalanuvchilar", callback_data="live_sessions")],
        [InlineKeyboardButton(text="👥 Foydalanuvchilar", callback_data="users_menu"),
         InlineKeyboardButton(text="🔍 Foydalanuvchi qidirish", callback_data="user_search")],
        [InlineKeyboardButton(text="📢 Xabar yuborish (barchaga)", callback_data="broadcast_start"),
         InlineKeyboardButton(text="💰 Moliyaviy hisobot", callback_data="finance_report")],
        [InlineKeyboardButton(text="✏️ Testni tahrirlash", callback_data="edittest_start"),
         InlineKeyboardButton(text="⏸ Testni yashirish/qaytarish", callback_data="hidetest_menu")],
        [InlineKeyboardButton(text="📚 Barcha testlar (kim yaratgan)", callback_data="alltests_menu")],
        [InlineKeyboardButton(text="📅 Bugungi hisobot", callback_data="today_digest"),
         InlineKeyboardButton(text="🏆 Testlar reytingi", callback_data="tests_ranking")],
        [InlineKeyboardButton(text="🛑 Testlarni yakunlash/qayta ochish", callback_data="closetest_menu")],
        [InlineKeyboardButton(text="✍️ Tekshirilmagan ochiq javoblar", callback_data="ogtests")],
        [InlineKeyboardButton(text="💾 Zaxira nusxa olish (backup)", callback_data="backup_now"),
         InlineKeyboardButton(text="📥 Zaxiradan tiklash", callback_data="restore_start")],
    ])

    text = (
        f"<b>📊 Admin Statistika Paneli</b>\n\n"
        f"👥 Userlar: <b>{users_count}</b>\n"
        f"📚 Testlar: <b>{tests_count}</b> (🆓 {free_count} / 🧾 {paid_count})\n\n"
        f"🟡 Pending: <b>{pending_count}</b>\n"
        f"✅ Approved (sotilgan): <b>{approved_count}</b>\n"
        f"❌ Rejected: <b>{rejected_count}</b>\n\n"
        f"💰 Umumiy tushum: <b>{int(total_revenue):,}".replace(",", " ") + " so'm</b>\n"
        f"📈 Umumiy o‘rtacha o‘zlashtirish: <b>{overall_avg:.1f}%</b>\n"
    )

    if best:
        uname, code, title, pct, dt = best["full_name"], best["code"], best["title"], best["percent"], best["date"]
        text += f"\n🏅 Eng yuqori natija: <b>{uname}</b>\n🔑 <code>{code}</code> {title} — <b>{float(pct):.1f}%</b> ({dt})\n"

    await message.answer(text, parse_mode=ParseMode.HTML, reply_markup=kb)


@router.callback_query(F.data == "rasch_menu")
async def rasch_menu(callback: CallbackQuery):
    if not await is_admin_user(callback.from_user.id):
        return

    tests = await POOL.fetch("""
        SELECT id, code, title
        FROM tests
        WHERE exam_type='rasch'
        ORDER BY id DESC
        LIMIT 50
    """)

    if not tests:
        return await callback.message.answer("Rasch testlar yo‘q.")

    kb = InlineKeyboardMarkup(inline_keyboard=[])
    for r in tests:
        kb.inline_keyboard.append([
            InlineKeyboardButton(text=f"⚡ {r['code']} — {r['title']}", callback_data=f"raschrep_{r['id']}")
        ])

    await callback.message.answer(
        "⚡ <b>Rasch hisobot</b>\nTestni tanlang:",
        parse_mode=ParseMode.HTML,
        reply_markup=kb
    )


async def _render_xl_list(answerable, page: int = 0):
    tests, has_more = await _fetch_tests_page(page)
    if not tests and page == 0:
        return await answerable.answer("Testlar yo‘q.")

    kb = InlineKeyboardMarkup(inline_keyboard=[])
    for r in tests:
        kb.inline_keyboard.append([InlineKeyboardButton(text=f"📤 {r['code']} — {r['title']}", callback_data=f"xl_{r['id']}")])
    nav = _pagination_nav_row("xl", page, has_more)
    if nav:
        kb.inline_keyboard.append(nav)

    await answerable.answer(f"📤 <b>Excel eksport</b> ({page + 1}-sahifa)\nQaysi test bo‘yicha?",
                            parse_mode=ParseMode.HTML, reply_markup=kb)


@router.callback_query(F.data == "xl_menu")
async def xl_menu(callback: CallbackQuery):
    if not await is_admin_user(callback.from_user.id):
        return
    await _render_xl_list(callback.message, 0)


def _xlsx_safe(val):
    """
    Excel/Google Sheets formula-injection'dan himoya: agar foydalanuvchi
    kiritgan matn (masalan ism-familiya) "=", "+", "-", "@" bilan boshlansa,
    Excel buni formula deb o'qib, faylni ochgan admin kompyuterida buyruq
    ishga tushirishi mumkin. Oldini olish uchun bunday qatorlar oldiga
    apostrof qo'yamiz — Excel uni oddiy matn sifatida ko'rsatadi.
    """
    if val is None:
        return ""
    s = str(val)
    if s and s[0] in ("=", "+", "-", "@"):
        return "'" + s
    return s


@router.callback_query(F.data.startswith("xl_"))
async def xl_export(callback: CallbackQuery, bot: Bot):
    test_id = int(callback.data.split("_")[1])
    user_tg_id = callback.from_user.id
    is_admin = await is_admin_user(user_tg_id)

    c_row = await POOL.fetchrow("SELECT creator_id FROM tests WHERE id=$1", test_id)
    is_creator = c_row and c_row["creator_id"] == user_tg_id

    if not (is_admin or is_creator):
        return await callback.answer("Sizda ruxsat yo'q!", show_alert=True)

    row = await POOL.fetchrow("SELECT exam_type FROM tests WHERE id=$1", test_id)
    if row:
        et = (row["exam_type"] or "simple")
        if et == "rasch":
            await finalize_rush_for_test_if_ready(test_id, bot=bot)

    t = await POOL.fetchrow("""
        SELECT code, title, is_free, price, duration, questions_count, start_mode, start_at, exam_type,
               open_ref_answers
        FROM tests WHERE id=$1
    """, test_id)

    if not t:
        return await callback.message.answer("Test topilmadi.")

    code, title, is_free, price, duration, qcount, start_mode, start_at, et2 = (
        t["code"], t["title"], t["is_free"], t["price"], t["duration"], t["questions_count"],
        t["start_mode"], t["start_at"], t["exam_type"]
    )

    results = await POOL.fetch("""
        SELECT u.full_name, u.phone, u.telegram_id,
               r.score, r.total_questions, r.percent,
               r.rasch_ball, r.rasch_percent, r.grade,
               r.open_text_answers, r.date
        FROM results r
        JOIN users u ON u.id=r.user_id
        WHERE r.test_id=$1
        ORDER BY r.percent DESC, r.id ASC
    """, test_id)

    pays = await POOL.fetch("""
        SELECT u.full_name, u.phone, u.telegram_id,
               p.status, p.created_at, p.reject_reason, p.started_at
        FROM payments p
        JOIN users u ON u.id=p.user_id
        WHERE p.test_id=$1
        ORDER BY p.id DESC
    """, test_id)

    wb = Workbook()
    ws_info = wb.active
    ws_info.title = "TestInfo"

    ws_info.append(["Field", "Value"])
    ws_info.append(["Code", code])
    ws_info.append(["Title", _xlsx_safe(title)])
    ws_info.append(["Type", "FREE" if int(is_free) == 1 else "PAID"])
    ws_info.append(["Price", int(price) if price is not None else 0])
    ws_info.append(["Duration(min)", int(duration)])
    ws_info.append(["Questions", int(qcount)])
    ws_info.append(["ExamType", (et2 or "simple")])
    ws_info.append(["StartMode", start_mode or "normal"])
    ws_info.append(["StartAt", start_at or ""])

    ws_res = wb.create_sheet("Results")
    ws_res.append([
        "FullName", "Phone", "TelegramID", "Score", "Total", "Percent", "RaschBall", "RaschPercent", "Grade",
        "OpenAnswers", "Date"
    ])

    test_open_refs = {}
    if t["open_ref_answers"]:
        try:
            test_open_refs = json.loads(t["open_ref_answers"])
        except json.JSONDecodeError:
            test_open_refs = {}

    for r in results:
        open_txt = ""
        if r["open_text_answers"]:
            try:
                open_map = json.loads(r["open_text_answers"])
                parts = []
                for k, v in open_map.items():
                    text_val, photo_id = _open_answer_preview(v)
                    shown = text_val if text_val else ("[rasm — botda ko'ring]" if photo_id else "-")
                    parts.append(f"Q{k}: {shown} (kutilgan: {test_open_refs.get(k, '-')})")
                open_txt = " | ".join(parts)
            except json.JSONDecodeError:
                open_txt = ""
        ws_res.append([_xlsx_safe(r["full_name"]), _xlsx_safe(r["phone"]), r["telegram_id"], r["score"],
                        r["total_questions"], r["percent"], r["rasch_ball"], r["rasch_percent"], r["grade"],
                        _xlsx_safe(open_txt), r["date"]])

    ws_pay = wb.create_sheet("Payments")
    ws_pay.append(["FullName", "Phone", "TelegramID", "Status", "CreatedAt", "RejectReason", "StartedAt"])
    for p in pays:
        ws_pay.append([_xlsx_safe(p["full_name"]), _xlsx_safe(p["phone"]), p["telegram_id"], p["status"],
                        p["created_at"], _xlsx_safe(p["reject_reason"]), p["started_at"]])

    for ws in [ws_info, ws_res, ws_pay]:
        for col in range(1, ws.max_column + 1):
            max_len = 0
            col_letter = get_column_letter(col)
            for cell in ws[col_letter]:
                val = "" if cell.value is None else str(cell.value)
                max_len = max(max_len, len(val))
            ws.column_dimensions[col_letter].width = min(max_len + 2, 40)

    safe_code = re.sub(r"[^A-Z0-9\-_]+", "_", code.upper())
    filename = f"test_stats_{safe_code}.xlsx"
    tmp_dir = tempfile.gettempdir()
    path = os.path.join(tmp_dir, filename)
    await asyncio.to_thread(wb.save, path)  # katta faylda bot "qotib" qolmasligi uchun

    await bot.send_document(
        chat_id=callback.from_user.id,
        document=FSInputFile(path),
        caption=f"📤 Excel eksport tayyor.\n🔑 <code>{code}</code> — <b>{title}</b>",
        parse_mode=ParseMode.HTML
    )


@router.callback_query(F.data.startswith("raschrep_"))
async def rasch_report_text(callback: CallbackQuery, bot: Bot):
    test_id = int(callback.data.split("_")[1])
    user_tg_id = callback.from_user.id
    is_admin = await is_admin_user(user_tg_id)

    c_row = await POOL.fetchrow("SELECT creator_id FROM tests WHERE id=$1", test_id)
    is_creator = c_row and c_row["creator_id"] == user_tg_id

    if not (is_admin or is_creator):
        return await callback.answer("Sizda ruxsat yo'q!", show_alert=True)

    await finalize_rush_for_test_if_ready(test_id, bot=bot)

    t = await POOL.fetchrow("SELECT code, title FROM tests WHERE id=$1", test_id)
    if not t:
        return await callback.message.answer("Test topilmadi.")
    code, title = t["code"], t["title"]

    agg = await POOL.fetchrow("""
        SELECT
            COUNT(*) AS total_results,
            SUM(CASE WHEN rasch_ball IS NOT NULL THEN 1 ELSE 0 END) AS computed,
            AVG(rasch_ball) AS avg_ball,
            AVG(rasch_percent) AS avg_pct,
            SUM(CASE WHEN grade IS NOT NULL THEN 1 ELSE 0 END) AS graded
        FROM results
        WHERE test_id=$1
    """, test_id)

    if not agg:
        total_results, computed, avg_ball, avg_pct, graded = 0, 0, 0, 0, 0
    else:
        total_results, computed, avg_ball, avg_pct, graded = (
            agg["total_results"], agg["computed"], agg["avg_ball"], agg["avg_pct"], agg["graded"]
        )

    text = (
        f"⚡ <b>Rasch Hisobot</b>\n\n"
        f"🔑 <code>{code}</code>\n"
        f"📚 <b>{title}</b>\n\n"
        f"👥 Qatnashganlar: <b>{int(total_results or 0)}</b>\n"
        f"🧮 Hisoblanganlar: <b>{int(computed or 0)}</b>\n"
        f"📈 O‘rtacha Rasch ball: <b>{float(avg_ball or 0):.1f} / 75</b>\n"
        f"📈 O‘rtacha Rasch foiz: <b>{float(avg_pct or 0):.1f}%</b>\n"
        f"🏅 Daraja olganlar: <b>{int(graded or 0)}</b>\n"
    )

    await callback.message.answer(text, parse_mode=ParseMode.HTML)


def _fmt_money(v) -> str:
    return f"{int(v or 0):,}".replace(",", " ") + " so'm"


# ===================== ADMIN: VAQT ORALIG'IDA NATIJALAR =====================
TRANGE_OPTIONS = [("1 soat", 1), ("3 soat", 3), ("6 soat", 6), ("24 soat", 24), ("7 kun", 168), ("Boshidan", 0)]
_TRANGE_LABELS = {hours: label for label, hours in TRANGE_OPTIONS}


async def _render_trange_list(answerable, page: int = 0):
    tests, has_more = await _fetch_tests_page(page)
    if not tests and page == 0:
        return await answerable.answer("Testlar yo'q.")
    kb = InlineKeyboardMarkup(inline_keyboard=[])
    for r in tests:
        kb.inline_keyboard.append(
            [InlineKeyboardButton(text=f"⏱ {r['code']} — {r['title']}", callback_data=f"trange_t_{r['id']}")])
    nav = _pagination_nav_row("trange", page, has_more)
    if nav:
        kb.inline_keyboard.append(nav)
    await answerable.answer(f"⏱ <b>Vaqt oralig'ida natijalar</b> ({page + 1}-sahifa)\nQaysi test bo'yicha ko'rmoqchisiz?",
                            parse_mode=ParseMode.HTML, reply_markup=kb)


@router.callback_query(F.data == "trange_menu")
async def trange_menu(callback: CallbackQuery):
    if not await is_admin_user(callback.from_user.id):
        return
    await _render_trange_list(callback.message, 0)


@router.callback_query(F.data.startswith("trange_t_"))
async def trange_pick_range(callback: CallbackQuery):
    if not await is_admin_user(callback.from_user.id):
        return
    test_id = int(callback.data.split("_")[2])
    kb = InlineKeyboardMarkup(inline_keyboard=[])
    row = []
    for label, hours in TRANGE_OPTIONS:
        row.append(InlineKeyboardButton(text=label, callback_data=f"trange_r_{test_id}_{hours}"))
        if len(row) == 2:
            kb.inline_keyboard.append(row)
            row = []
    if row:
        kb.inline_keyboard.append(row)
    await callback.message.answer("Qaysi vaqt oralig'ini ko'rmoqchisiz?", reply_markup=kb)


@router.callback_query(F.data.startswith("trange_r_"))
async def trange_show(callback: CallbackQuery):
    if not await is_admin_user(callback.from_user.id):
        return
    _, _, test_id_s, hours_s = callback.data.split("_")
    test_id, hours = int(test_id_s), int(hours_s)

    t = await POOL.fetchrow("SELECT code, title FROM tests WHERE id=$1", test_id)
    if not t:
        return await callback.message.answer("Test topilmadi.")

    if hours > 0:
        cutoff = now_ts() - hours * 3600
        rows = await POOL.fetch("""
            SELECT u.full_name, r.percent, r.score, r.total_questions, r.finished_ts
            FROM results r JOIN users u ON u.id=r.user_id
            WHERE r.test_id=$1 AND r.finished_ts>=$2
            ORDER BY r.finished_ts DESC
            LIMIT 40
        """, test_id, cutoff)
    else:
        rows = await POOL.fetch("""
            SELECT u.full_name, r.percent, r.score, r.total_questions, r.finished_ts
            FROM results r JOIN users u ON u.id=r.user_id
            WHERE r.test_id=$1
            ORDER BY r.finished_ts DESC
            LIMIT 40
        """, test_id)

    range_label = _TRANGE_LABELS.get(hours, f"{hours} soat")
    text = f"⏱ <b>{t['title']}</b> (<code>{t['code']}</code>) — {range_label} ichida:\n\n"
    if not rows:
        text += "Hech kim topshirmagan."
    else:
        text += f"👥 Topilgan: <b>{len(rows)}</b> ta natija (so'nggi 40 tasi ko'rsatiladi)\n\n"
        for r in rows:
            when = datetime.fromtimestamp(r["finished_ts"], TZ).strftime("%d-%m, %H:%M") if r["finished_ts"] else "?"
            text += f"• {r['full_name']} — {r['score']}/{r['total_questions']} ({float(r['percent']):.1f}%) — {when}\n"

    await callback.message.answer(text, parse_mode=ParseMode.HTML)


# ===================== ADMIN: HOZIR FAOL FOYDALANUVCHILAR =====================
@router.callback_query(F.data == "live_sessions")
async def live_sessions(callback: CallbackQuery):
    if not await is_admin_user(callback.from_user.id):
        return
    now_ = now_ts()
    rows = await POOL.fetch("""
        SELECT u.full_name, t.code, t.title, s.start_ts, s.duration_sec
        FROM sessions s
        JOIN users u ON u.id=s.user_id
        JOIN tests t ON t.id=s.test_id
        WHERE s.start_ts + s.duration_sec > $1
        ORDER BY s.start_ts DESC
        LIMIT 50
    """, now_)
    if not rows:
        return await callback.message.answer("🟢 Hozirda hech kim test yechayotgani yo'q.")
    text = f"🟢 <b>Hozir faol: {len(rows)} kishi</b>\n\n"
    for r in rows:
        remain = int(r["start_ts"]) + int(r["duration_sec"]) - now_
        text += f"• {r['full_name']} — <code>{r['code']}</code> {r['title']} (qoldi: {fmt_seconds(remain)})\n"
    await callback.message.answer(text, parse_mode=ParseMode.HTML)


# ===================== ADMIN: FOYDALANUVCHILAR =====================
@router.callback_query(F.data == "users_menu")
async def users_menu(callback: CallbackQuery):
    if not await is_admin_user(callback.from_user.id):
        return
    total = await POOL.fetchval("SELECT COUNT(*) FROM users")
    recent = await POOL.fetch("SELECT full_name, telegram_id, created_ts FROM users ORDER BY id DESC LIMIT 10")

    text = f"👥 <b>Foydalanuvchilar</b>\n\nJami: <b>{total}</b> ta\n\n<b>Oxirgi qo'shilganlar:</b>\n"
    for r in recent:
        when = datetime.fromtimestamp(r["created_ts"], TZ).strftime("%Y-%m-%d %H:%M") if r["created_ts"] else "noma'lum"
        text += f"• {r['full_name']} — <code>{r['telegram_id']}</code> ({when})\n"

    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="📤 Excel eksport (barcha foydalanuvchilar)", callback_data="users_export")]
    ])
    await callback.message.answer(text, parse_mode=ParseMode.HTML, reply_markup=kb)


@router.callback_query(F.data == "users_export")
async def users_export(callback: CallbackQuery, bot: Bot):
    if not await is_admin_user(callback.from_user.id):
        return

    rows = await POOL.fetch("""
        SELECT u.id, u.full_name, u.phone, u.telegram_id, u.role, u.created_ts,
               (SELECT COUNT(*) FROM results r WHERE r.user_id=u.id) AS results_count,
               (SELECT COUNT(*) FROM payments p WHERE p.user_id=u.id AND p.status='approved') AS approved_payments
        FROM users u
        ORDER BY u.id ASC
    """)

    wb = Workbook()
    ws = wb.active
    ws.title = "Users"
    ws.append(["FullName", "Phone", "TelegramID", "Role", "RegisteredAt", "ResultsCount", "ApprovedPayments"])
    for r in rows:
        reg_at = datetime.fromtimestamp(r["created_ts"], TZ).strftime("%Y-%m-%d %H:%M") if r["created_ts"] else ""
        ws.append([_xlsx_safe(r["full_name"]), _xlsx_safe(r["phone"]), r["telegram_id"], r["role"], reg_at,
                   r["results_count"], r["approved_payments"]])

    for col in range(1, ws.max_column + 1):
        max_len = 0
        col_letter = get_column_letter(col)
        for cell in ws[col_letter]:
            val = "" if cell.value is None else str(cell.value)
            max_len = max(max_len, len(val))
        ws.column_dimensions[col_letter].width = min(max_len + 2, 40)

    tmp_dir = tempfile.gettempdir()
    path = os.path.join(tmp_dir, "barcha_foydalanuvchilar.xlsx")
    await asyncio.to_thread(wb.save, path)  # katta faylda bot "qotib" qolmasligi uchun

    await bot.send_document(
        chat_id=callback.from_user.id,
        document=FSInputFile(path),
        caption=f"📤 Barcha foydalanuvchilar ({len(rows)} ta), Telegram ID bilan."
    )


@router.callback_query(F.data == "user_search")
async def user_search_start(callback: CallbackQuery, state: FSMContext):
    if not await is_admin_user(callback.from_user.id):
        return
    await callback.message.answer(
        "🔍 Qidirish uchun Telegram ID, ism yoki telefon raqamining bir qismini yuboring:"
    )
    await state.set_state(AdminUserSearchState.waiting_query)


@router.message(AdminUserSearchState.waiting_query)
async def user_search_run(message: Message, state: FSMContext):
    if not await is_admin_user(message.from_user.id):
        await state.clear()
        return
    if (message.text or "") in TOP_MENU_BTNS:
        await state.clear()
        return

    q = (message.text or "").strip()
    if _looks_like_int(q.lstrip("+")):
        digits = _digits_only(q)
        rows = await POOL.fetch("""
            SELECT id, full_name, telegram_id FROM users
            WHERE telegram_id=$1 OR regexp_replace(COALESCE(phone,''), '\\D', '', 'g') LIKE '%' || $2 || '%'
            ORDER BY id DESC LIMIT 15
        """, int(digits) if digits else 0, digits[-9:] if len(digits) >= 7 else digits)
    else:
        like = f"%{q}%"
        rows = await POOL.fetch("""
            SELECT id, full_name, telegram_id FROM users
            WHERE full_name ILIKE $1 OR phone ILIKE $1
            ORDER BY id DESC LIMIT 15
        """, like)

    if not rows:
        await message.answer("❌ Hech kim topilmadi. Boshqa so'z bilan qayta urinib ko'ring:")
        return

    kb = InlineKeyboardMarkup(inline_keyboard=[])
    for r in rows:
        kb.inline_keyboard.append([InlineKeyboardButton(
            text=f"{r['full_name']} — {r['telegram_id']}", callback_data=f"userprofile_{r['id']}")])
    await message.answer(f"🔍 Topildi: {len(rows)} ta", reply_markup=kb)
    await state.clear()


@router.callback_query(F.data.startswith("userprofile_"))
async def user_profile_view(callback: CallbackQuery, bot: Bot):
    if not await is_admin_user(callback.from_user.id):
        return
    uid = int(callback.data.split("_")[1])
    u = await POOL.fetchrow("SELECT * FROM users WHERE id=$1", uid)
    if not u:
        return await callback.message.answer("Foydalanuvchi topilmadi.")

    # ✅ Username'ni bazada saqlamaymiz (u tez-tez o'zgarishi mumkin), shuning
    # uchun profil ochilganda Telegram'dan jonli so'raymiz — admin shu odam
    # bilan to'g'ridan-to'g'ri bog'lana olishi uchun.
    username_line = ""
    try:
        chat = await bot.get_chat(u["telegram_id"])
        if chat.username:
            username_line = f"🔗 Username: @{chat.username}\n"
    except Exception:
        pass

    results_count = await POOL.fetchval("SELECT COUNT(*) FROM results WHERE user_id=$1", uid)
    avg_pct = await POOL.fetchval("SELECT AVG(percent) FROM results WHERE user_id=$1", uid)
    approved = await POOL.fetchval("SELECT COUNT(*) FROM payments WHERE user_id=$1 AND status='approved'", uid)
    pending = await POOL.fetchval("SELECT COUNT(*) FROM payments WHERE user_id=$1 AND status='pending'", uid)
    spent = await POOL.fetchval("""
        SELECT COALESCE(SUM(COALESCE(p.amount, t.price)),0) FROM payments p JOIN tests t ON t.id=p.test_id
        WHERE p.user_id=$1 AND p.status='approved'
    """, uid)
    reg_at = datetime.fromtimestamp(u["created_ts"], TZ).strftime("%Y-%m-%d %H:%M") if u["created_ts"] else "noma'lum"

    text = (
        f"👤 <b>{u['full_name']}</b>\n\n"
        f"🆔 Telegram ID: <code>{u['telegram_id']}</code>\n"
        f"{username_line}"
        f"📞 Telefon: {u['phone'] or '-'}\n"
        f"📍 Viloyat: {u['region'] or '-'}\n"
        f"🎂 Yosh: {u['age'] if u['age'] is not None else '-'}\n"
        f"🏫 Maktab/muassasa: {u['school'] or '-'}\n"
        f"📚 Sinf/kurs: {u['grade'] or '-'}\n"
        f"🎓 Rol: {u['role']}\n"
        f"🗓 Ro'yxatdan o'tgan: {reg_at}\n\n"
        f"📄 Topshirgan testlar: <b>{results_count}</b> (o'rtacha {float(avg_pct or 0):.1f}%)\n"
        f"💳 To'lovlar: ✅ {approved} | 🟡 {pending}\n"
        f"💰 Jami to'lagan: <b>{_fmt_money(spent)}</b>"
    )
    await callback.message.answer(text, parse_mode=ParseMode.HTML)


# ===================== ADMIN: BARCHAGA XABAR YUBORISH (REKLAMA) =====================
# ✅ bot.copy_message() ishlatiladi — bu Telegram'ning har qanday xabar turini
# (matn, rasm, video, hujjat/fayl, ovozli xabar, audio, GIF, so'rovnoma va h.k.)
# formatini/media-faylini o'zgartirmasdan boshqa chatga nusxalab yuboradigan
# tayyor usuli. Shu sababli admin nima yuborsa (yoki bir nechta turini
# birlashtirib yuborsa ham) — buni alohida-alohida qo'lda aniqlashning
# hojati yo'q, xabar aynan qanday yuborilgan bo'lsa, xuddi shunday tarqatiladi.
@router.callback_query(F.data == "broadcast_start")
async def broadcast_start(callback: CallbackQuery, state: FSMContext):
    if not await is_admin_user(callback.from_user.id):
        return
    await callback.message.answer(
        "📢 <b>Barchaga xabar (reklama) yuborish</b>\n\n"
        "Istalgan turdagi xabarni yuboring — <b>matn, rasm, video, hujjat/fayl, "
        "ovozli xabar, audio yoki GIF</b> — nima yuborsangiz, xuddi shu ko'rinishda "
        "barcha foydalanuvchilarga yetkaziladi.\n\n"
        "<i>Eslatma: yuborishdan oldin sizdan albatta tasdiq so'rayman.</i>",
        parse_mode=ParseMode.HTML,
        reply_markup=nav_kb(include_back=True)
    )
    await state.set_state(AdminBroadcastState.waiting_text)


@router.message(AdminBroadcastState.waiting_text)
async def broadcast_preview(message: Message, state: FSMContext):
    if not await is_admin_user(message.from_user.id):
        await state.clear()
        return
    if (message.text or "") in TOP_MENU_BTNS:
        return

    has_content = bool(
        message.text or message.caption or message.photo or message.video or message.document or
        message.audio or message.voice or message.animation or message.video_note or
        message.sticker or message.poll
    )
    if not has_content:
        return await message.answer("⚠️ Bo'sh xabar. Matn, rasm, video, fayl yoki boshqa turdagi xabar yuboring:")

    await state.update_data(broadcast_chat_id=message.chat.id, broadcast_message_id=message.message_id)
    total = await POOL.fetchval("SELECT COUNT(*) FROM users")

    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=f"✅ Ha, {total} kishiga yuborilsin", callback_data="broadcast_confirm")],
        [InlineKeyboardButton(text="❌ Bekor qilish", callback_data="broadcast_cancel")]
    ])
    await message.answer(
        f"👆 <b>Yuqoridagi xabaringiz</b> — aynan shu ko'rinishda <b>{total}</b> ta ro'yxatdan o'tgan "
        f"foydalanuvchiga yuboriladi.\n\n❗️Tasdiqlaysizmi?",
        parse_mode=ParseMode.HTML, reply_markup=kb
    )


@router.callback_query(AdminBroadcastState.waiting_text, F.data == "broadcast_cancel")
async def broadcast_cancel(callback: CallbackQuery, state: FSMContext):
    await state.clear()
    await callback.message.answer("Bekor qilindi.")


@router.callback_query(AdminBroadcastState.waiting_text, F.data == "broadcast_confirm")
async def broadcast_confirm(callback: CallbackQuery, state: FSMContext, bot: Bot):
    if not await is_admin_user(callback.from_user.id):
        return
    data = await state.get_data()
    src_chat_id = data.get("broadcast_chat_id")
    src_message_id = data.get("broadcast_message_id")
    await state.clear()
    if not src_chat_id or not src_message_id:
        await callback.message.answer("⚠️ Xatolik: yubormoqchi bo'lgan xabar topilmadi. Qaytadan boshlang.")
        return

    await callback.message.answer("🚀 Yuborish boshlandi — tayyor bo'lgach sizga alohida xabar yuboraman.")
    asyncio.create_task(_run_broadcast(bot, callback.from_user.id, src_chat_id, src_message_id))


async def _run_broadcast(bot: Bot, admin_tg_id: int, src_chat_id: int, src_message_id: int):
    rows = await POOL.fetch("SELECT telegram_id FROM users")
    sent = 0
    failed = 0
    for r in rows:
        # ✅ Telegram 429 (RetryAfter) qaytarsa, ko'rsatilgan vaqt kutib, shu
        # foydalanuvchiga qayta urinamiz — avval u xabarni umuman olmasdi.
        for attempt in range(3):
            try:
                await bot.copy_message(chat_id=r["telegram_id"], from_chat_id=src_chat_id,
                                       message_id=src_message_id)
                sent += 1
                break
            except TelegramRetryAfter as e:
                if attempt == 2:
                    failed += 1
                    break
                await asyncio.sleep(e.retry_after + 1)
            except Exception:
                failed += 1
                break
        await asyncio.sleep(0.05)  # Telegram flood-limitiga urilmaslik uchun
    try:
        await bot.send_message(
            admin_tg_id,
            f"✅ <b>Xabar yuborish yakunlandi.</b>\n📤 Yetkazildi: {sent}\n❌ Yetkazilmadi: {failed}",
            parse_mode=ParseMode.HTML
        )
    except Exception:
        pass


# ===================== ADMIN: MOLIYAVIY HISOBOT =====================
@router.callback_query(F.data == "finance_report")
async def finance_report(callback: CallbackQuery):
    if not await is_admin_user(callback.from_user.id):
        return

    # ✅ Haqiqatda to'langan summa (p.amount) va TASDIQLANGAN kun bo'yicha.
    # Avval testning hozirgi narxi va chek yuborilgan kun ishlatilardi — narx
    # o'zgarsa, butun o'tmishdagi tushum ham "o'zgarib" ketardi.
    total = await POOL.fetchval("""
        SELECT COALESCE(SUM(COALESCE(p.amount, t.price)),0) FROM payments p JOIN tests t ON t.id=p.test_id
        WHERE p.status='approved'
    """)
    today_start = (datetime.now(TZ).strftime("%Y-%m-%d") + " 00:00")
    today = await POOL.fetchval("""
        SELECT COALESCE(SUM(COALESCE(p.amount, t.price)),0) FROM payments p JOIN tests t ON t.id=p.test_id
        WHERE p.status='approved' AND COALESCE(p.approved_at, p.created_at) >= $1
    """, today_start)
    week_start = (datetime.now(TZ) - timedelta(days=7)).strftime("%Y-%m-%d %H:%M")
    week = await POOL.fetchval("""
        SELECT COALESCE(SUM(COALESCE(p.amount, t.price)),0) FROM payments p JOIN tests t ON t.id=p.test_id
        WHERE p.status='approved' AND COALESCE(p.approved_at, p.created_at) >= $1
    """, week_start)

    top_tests = await POOL.fetch("""
        SELECT t.code, t.title, COUNT(*) AS cnt, SUM(COALESCE(p.amount, t.price)) AS revenue
        FROM payments p JOIN tests t ON t.id=p.test_id
        WHERE p.status='approved'
        GROUP BY t.id, t.code, t.title
        ORDER BY cnt DESC
        LIMIT 5
    """)

    text = (
        f"💰 <b>Moliyaviy hisobot</b>\n\n"
        f"📊 Umumiy tushum: <b>{_fmt_money(total)}</b>\n"
        f"📅 Bugun: <b>{_fmt_money(today)}</b>\n"
        f"🗓 So'nggi 7 kun: <b>{_fmt_money(week)}</b>\n"
    )
    if top_tests:
        text += "\n🏅 <b>Eng ko'p sotilgan testlar:</b>\n"
        for r in top_tests:
            text += f"• <code>{r['code']}</code> {_h(r['title'])} — {r['cnt']} ta ({_fmt_money(r['revenue'])})\n"

    await callback.message.answer(text, parse_mode=ParseMode.HTML)


# ===================== TESTLAR RO'YXATI — SAHIFALASH (PAGINATION) =====================
# ✅ Admin tugmalar orqali testlar ro'yxatini ko'radigan barcha joylar (tahrirlash,
# o'chirish, yashirish/qaytarish, yakunlash/qayta ochish, Excel eksport) avval
# faqat "oxirgi 50 ta" bilan cheklangan edi — 50 tadan ortiq test bo'lsa, eski
# testlarga tugma orqali umuman kirib bo'lmasdi. Endi hammasi sahifalab
# ko'rsatiladi, "⬅️ Oldingi/Keyingi ➡️" tugmalari bilan.
LIST_PAGE_SIZE = 20


async def _fetch_tests_page(page: int, extra_where: str = "", extra_args: list | None = None):
    extra_args = extra_args or []
    offset = max(0, page) * LIST_PAGE_SIZE
    rows = await POOL.fetch(
        f"SELECT id, code, title, is_free, COALESCE(is_hidden,0) AS is_hidden, "
        f"COALESCE(is_closed,0) AS is_closed FROM tests {extra_where} "
        f"ORDER BY id DESC LIMIT {LIST_PAGE_SIZE + 1} OFFSET {offset}",
        *extra_args
    )
    has_more = len(rows) > LIST_PAGE_SIZE
    return rows[:LIST_PAGE_SIZE], has_more


def _pagination_nav_row(list_key: str, page: int, has_more: bool):
    row = []
    if page > 0:
        row.append(InlineKeyboardButton(text="⬅️ Oldingi", callback_data=f"tlp_{list_key}_{page - 1}"))
    if has_more:
        row.append(InlineKeyboardButton(text="Keyingi ➡️", callback_data=f"tlp_{list_key}_{page + 1}"))
    return row


@router.callback_query(F.data.startswith("tlp_"))
async def test_list_page_nav(callback: CallbackQuery):
    if not await is_admin_user(callback.from_user.id):
        return
    _, list_key, page_s = callback.data.split("_")
    page = int(page_s)
    renderer = _PAGINATED_LIST_RENDERERS.get(list_key)
    if renderer:
        await renderer(callback.message, page)


# ===================== ADMIN: TESTNI TAHRIRLASH =====================
async def _render_edittest_list(answerable, page: int = 0):
    tests, has_more = await _fetch_tests_page(page)
    if not tests and page == 0:
        return await answerable.answer("Testlar yo'q.")

    kb = InlineKeyboardMarkup(inline_keyboard=[])
    for r in tests:
        kb.inline_keyboard.append(
            [InlineKeyboardButton(text=f"✏️ {r['code']} — {r['title']}", callback_data=f"admedit_pick_{r['id']}")])
    nav = _pagination_nav_row("edit", page, has_more)
    if nav:
        kb.inline_keyboard.append(nav)

    await answerable.answer(
        f"✏️ <b>Testni tahrirlash</b> ({page + 1}-sahifa)\nQaysi testni tahrirlamoqchisiz?",
        parse_mode=ParseMode.HTML, reply_markup=kb
    )


@router.callback_query(F.data == "edittest_start")
async def edittest_start(callback: CallbackQuery, state: FSMContext):
    if not await is_admin_user(callback.from_user.id):
        return
    await _render_edittest_list(callback.message, 0)


@router.callback_query(F.data.startswith("admedit_pick_"))
async def admedit_pick(callback: CallbackQuery, state: FSMContext):
    test_id = int(callback.data.split("_")[2])
    if not await _is_admin_or_creator(callback.from_user.id, test_id):
        return await callback.answer("Bu testga huquqingiz yo'q!", show_alert=True)

    t = await POOL.fetchrow(
        "SELECT id, code, title, subject, price, duration, is_free, questions_count, exam_type, skills_cfg "
        "FROM tests WHERE id=$1", test_id)
    if not t:
        return await callback.answer("Test topilmadi.", show_alert=True)

    # ✅ Bosqichli testlar (DTM 5-fan / IELTS 4-ko'nikma) butunlay boshqa
    # tuzilishga ega (bitta fayl/javob kaliti o'rniga har bosqichda alohida) —
    # ular uchun alohida, bosqich tanlab tahrirlaydigan oqim ishlatiladi.
    if (t["exam_type"] or "") in STAGED_EXAM_TYPES:
        await _send_staged_edit_picker(callback.message, state, t)
        return

    await _send_edit_field_picker(callback.message, state, t)


async def _send_edit_field_picker(answerable, state: FSMContext, t):
    await state.update_data(edit_test_id=t["id"])
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=f"📚 Fan: {t['subject'] or '-'}", callback_data="edtf_subject")],
        [InlineKeyboardButton(text=f"📝 Sarlavha: {t['title']}", callback_data="edtf_title")],
        [InlineKeyboardButton(text=f"⏳ Davomiylik: {t['duration']} daq", callback_data="edtf_duration")],
        [InlineKeyboardButton(text="📄 Faylni almashtirish", callback_data="edtf_file")],
    ])
    if int(t["is_free"] or 0) == 0:
        kb.inline_keyboard.append(
            [InlineKeyboardButton(text=f"💰 Narx: {t['price']} so'm", callback_data="edtf_price")])
    if "questions_count" in t.keys():
        kb.inline_keyboard.append(
            [InlineKeyboardButton(text=f"🧮 Savollar soni: {t['questions_count']}", callback_data="edtf_questions_count")])
        kb.inline_keyboard.append(
            [InlineKeyboardButton(text="🔑 Javob kaliti", callback_data="edtf_answers")])

    await answerable.answer(
        f"✏️ <b>{t['title']}</b> (<code>{t['code']}</code>)\nQaysi maydonni o'zgartirmoqchisiz?\n\n"
        f"<i>⚠️ Savollar sonini yoki javob kalitini o'zgartirsangiz, ular bir-biriga mos kelishiga "
        f"e'tibor bering.</i>",
        parse_mode=ParseMode.HTML, reply_markup=kb
    )
    await state.set_state(AdminEditTestState.waiting_field)


_EDIT_FIELD_LABELS = {
    "subject": "Fan nomini",
    "title": "Sarlavhani",
    "duration": "Davomiylikni (daqiqada, raqam)",
    "price": "Narxni (so'mda, raqam)",
    "questions_count": "Savollar sonini (raqam)",
    "answers": "Javob kalitini (A-F harflarida, bo'shliqsiz — variantsiz/ochiq savol uchun x qo'ying)",
    "file": "Yangi faylni (PDF/audio/video)",
}


@router.callback_query(AdminEditTestState.waiting_field, F.data.startswith("edtf_"))
async def edittest_pick_field(callback: CallbackQuery, state: FSMContext):
    field = callback.data.split("_", 1)[1]
    if field not in _EDIT_FIELD_LABELS:
        return
    await state.update_data(edit_field=field)

    await state.set_state(AdminEditTestState.waiting_value)
    if field == "file":
        await callback.message.answer(
            "📄 Yangi faylni yuboring (PDF hujjat, yoki audio/video — Listening kabi holatlar uchun):"
        )
        return
    await callback.message.answer(f"Yangi qiymatni kiriting — {_EDIT_FIELD_LABELS[field]}:")


@router.message(AdminEditTestState.waiting_value, F.document | F.audio | F.video)
async def edittest_save_file(message: Message, state: FSMContext):
    data = await state.get_data()
    if data.get("edit_field") != "file":
        return
    test_id = data.get("edit_test_id")
    if not test_id:
        await state.clear()
        return await message.answer("⚠️ Xatolik yuz berdi. Qaytadan boshlang.")

    admin_flag = await is_admin_user(message.from_user.id)
    if not admin_flag:
        owner = await POOL.fetchval("SELECT creator_id FROM tests WHERE id=$1", test_id)
        if owner != message.from_user.id:
            await state.clear()
            return await message.answer("⚠️ Bu testni tahrirlashga ruxsatingiz yo'q.")

    if message.document:
        file_id = message.document.file_id
    elif message.audio:
        file_id = message.audio.file_id
    else:
        file_id = message.video.file_id

    await POOL.execute("UPDATE tests SET file_id=$1 WHERE id=$2", file_id, test_id)
    await state.clear()
    await message.answer("✅ Fayl yangilandi.", reply_markup=main_menu_kb(admin_flag))


@router.message(AdminEditTestState.waiting_value)
async def edittest_save(message: Message, state: FSMContext):
    if (message.text or "") in TOP_MENU_BTNS:
        return

    data = await state.get_data()
    test_id = data.get("edit_test_id")
    field = data.get("edit_field")
    value_raw = (message.text or "").strip()
    regraded = 0

    if (field not in _EDIT_FIELD_LABELS and field != "open_ref_answers") or not test_id:
        await state.clear()
        return await message.answer("⚠️ Xatolik yuz berdi. Qaytadan boshlang.")

    # ✅ Bu holatga ikkita yo'l bilan kelish mumkin: admin (istalgan testni) yoki
    # testning o'zi yaratgan egasi (faqat o'zinikini) tahrirlaydi.
    admin_flag = await is_admin_user(message.from_user.id)
    if not admin_flag:
        owner = await POOL.fetchval("SELECT creator_id FROM tests WHERE id=$1", test_id)
        if owner != message.from_user.id:
            await state.clear()
            return await message.answer("⚠️ Bu testni tahrirlashga ruxsatingiz yo'q.")

    if field == "file":
        return await message.answer("⚠️ Fayl (hujjat/audio/video) yuboring, matn emas:")

    if field in ("duration", "price", "questions_count"):
        if not value_raw.isdigit():
            return await message.answer("⚠️ Faqat raqam kiriting:")
        value = int(value_raw)
        if field == "duration" and (value < 1 or value > 600):
            return await message.answer("⚠️ Davomiylik 1..600 daqiqa oralig'ida bo'lsin:")
        if field == "questions_count" and (value < 1 or value > 300):
            return await message.answer("⚠️ Savollar soni 1..300 oralig'ida bo'lsin:")
        if field == "questions_count":
            # ✅ Savollar soni va javob kaliti DOIM birga o'zgarishi kerak. Avval faqat
            # son o'zgartirilib, kalit eskicha qolardi — test hech kim yecha olmaydigan
            # ("javob kaliti uzunligi mos emas") holatga tushardi. Natijalari bor
            # testda esa umuman o'zgartirib bo'lmaydi (o'quvchilar javoblari eski
            # savollar soniga mos).
            has_results = await POOL.fetchval("SELECT COUNT(*) FROM results WHERE test_id=$1", test_id)
            if has_results:
                await state.clear()
                return await message.answer(
                    f"⚠️ Bu testni allaqachon <b>{has_results}</b> kishi topshirgan — savollar sonini o'zgartirib "
                    f"bo'lmaydi (natijalar buzilib ketadi). Kerak bo'lsa, yangi test yarating.",
                    parse_mode=ParseMode.HTML, reply_markup=main_menu_kb(admin_flag))
            await state.update_data(edit_field="answers", pending_qcount=value)
            return await message.answer(
                f"✅ Savollar soni: <b>{value}</b>. Endi shunga mos <b>yangi javob kalitini</b> kiriting — "
                f"aynan {value} ta harf (A-F, ochiq savol uchun x):",
                parse_mode=ParseMode.HTML)
        # field ("duration","price") bilan cheklangani uchun xavfsiz
        await POOL.execute(f"UPDATE tests SET {field}=$1 WHERE id=$2", value, test_id)
    elif field == "answers":
        val = normalize_answers(value_raw)
        if not is_answer_string_valid(val):
            return await message.answer(
                "⚠️ Javob kaliti faqat <b>A, B, C, D, E, F</b> harflaridan iborat bo'lishi kerak:",
                parse_mode=ParseMode.HTML)
        current_q = data.get("pending_qcount") or await POOL.fetchval(
            "SELECT questions_count FROM tests WHERE id=$1", test_id)
        if len(val) != int(current_q or 0):
            return await message.answer(
                f"⚠️ Javoblar soni <b>{current_q}</b> ta bo'lishi kerak, siz {len(val)} ta yubordingiz. "
                f"Qayta kiriting:", parse_mode=ParseMode.HTML)

        open_positions = open_question_positions(val)
        if open_positions:
            await state.update_data(
                pending_answers=val, pending_open_positions=open_positions, edit_field="open_ref_answers"
            )
            await message.answer(_open_ref_prompt_text(open_positions), parse_mode=ParseMode.HTML)
            return
        await POOL.execute("UPDATE tests SET answers=$1, open_ref_answers=NULL, questions_count=$3 WHERE id=$2",
                           val, test_id, len(val))
        regraded = await _regrade_test_results(test_id)

    elif field == "open_ref_answers":
        positions = data.get("pending_open_positions") or []
        ref_map = parse_open_ref_answers(value_raw, positions)
        if ref_map is None:
            return await message.answer(
                f"⚠️ Qatorlar soni mos kelmadi — aynan <b>{len(positions)} qator</b> yuboring "
                f"(har birida bitta javob, savollar tartibida):",
                parse_mode=ParseMode.HTML
            )
        await POOL.execute(
            "UPDATE tests SET answers=$1, open_ref_answers=$2, questions_count=$4 WHERE id=$3",
            data.get("pending_answers"), json.dumps(ref_map, ensure_ascii=False), test_id,
            len(data.get("pending_answers") or "")
        )
        regraded = await _regrade_test_results(test_id)
    else:
        max_len = 150 if field == "title" else 80
        value_clean = sanitize_plain_text(value_raw, max_len=max_len)
        if len(value_clean) < 2:
            return await message.answer("⚠️ Juda qisqa yoki noto'g'ri belgilar bor. Qayta kiriting:")
        await POOL.execute(f"UPDATE tests SET {field}=$1 WHERE id=$2", value_clean, test_id)

    await state.clear()
    extra = ""
    if field in ("answers", "open_ref_answers") and regraded:
        extra = (f"\n♻️ Yangi kalit bo'yicha <b>{regraded}</b> ta mavjud natija qayta hisoblandi "
                 f"(sertifikatga loyiq bo'lib qolganlarga sertifikat avtomatik yuboriladi).")
    await message.answer(f"✅ Test yangilandi.{extra}", parse_mode=ParseMode.HTML,
                         reply_markup=main_menu_kb(admin_flag))


# ----- Bosqichli test (DTM 5-fan / IELTS 4-ko'nikma) tahrirlash -----
async def _send_staged_edit_picker(answerable, state: FSMContext, t) -> None:
    await state.update_data(edit_test_id=t["id"])
    cfg = _load_skills_cfg(t["skills_cfg"])
    order = cfg.get("order") or []

    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=f"📚 Fan: {t['subject'] or '-'}", callback_data="edtf_subject")],
        [InlineKeyboardButton(text=f"📝 Sarlavha: {t['title']}", callback_data="edtf_title")],
    ])
    if int(t["is_free"] or 0) == 0:
        kb.inline_keyboard.append(
            [InlineKeyboardButton(text=f"💰 Narx: {t['price']} so'm", callback_data="edtf_price")])

    for key in order:
        sec = cfg["sections"].get(key, {})
        kb.inline_keyboard.append(
            [InlineKeyboardButton(text=f"🧩 {sec.get('name', key)}", callback_data=f"edtstage_sec_{t['id']}_{key}")]
        )

    await answerable.answer(
        f"✏️ <b>{t['title']}</b> (<code>{t['code']}</code>) — bosqichli test.\n"
        f"Umumiy ma'lumotni yoki muayyan bosqichni tanlang:",
        parse_mode=ParseMode.HTML, reply_markup=kb
    )
    await state.set_state(AdminEditTestState.waiting_field)


@router.callback_query(F.data.startswith("edtstage_sec_"))
async def staged_edit_pick_section(callback: CallbackQuery, state: FSMContext) -> None:
    parts = callback.data.split("_")
    test_id, key = int(parts[2]), parts[3]
    if not await _is_admin_or_creator(callback.from_user.id, test_id):
        return await callback.answer("Sizda ruxsat yo'q!", show_alert=True)

    t = await POOL.fetchrow("SELECT skills_cfg FROM tests WHERE id=$1", test_id)
    cfg = _load_skills_cfg(t["skills_cfg"])
    section = cfg["sections"].get(key)
    if not section:
        return await callback.answer("Bosqich topilmadi.", show_alert=True)

    await state.update_data(
        stagedit_test_id=test_id, stagedit_key=key, stagedit_cfg=cfg, stagedit_field="duration"
    )
    await callback.answer()
    await callback.message.answer(
        f"✏️ <b>{section['name']}</b> — hozirgi davomiylik: {section.get('duration_min', 0)} daq.\n"
        f"Yangi davomiylikni daqiqada kiriting:"
    )
    await state.set_state(StagedEditState.collecting)


async def _staged_edit_ask_qcount(answerable, state: FSMContext) -> None:
    data = await state.get_data()
    cfg = data["stagedit_cfg"]
    key = data["stagedit_key"]
    current_q = cfg["sections"][key].get("qcount", 0)
    await answerable.answer(
        f"Savollar sonini kiriting (hozirgi: {current_q}). O'zgartirmaslik uchun <code>-</code> yuboring:",
        parse_mode=ParseMode.HTML
    )
    await state.update_data(stagedit_field="qcount")
    await state.set_state(StagedEditState.collecting)


async def _staged_edit_ask_answers(answerable, state: FSMContext) -> None:
    data = await state.get_data()
    cfg = data["stagedit_cfg"]
    key = data["stagedit_key"]
    qcount = cfg["sections"][key]["qcount"]
    await answerable.answer(
        f"Yangi javob kalitini kiriting — aynan <b>{qcount} ta</b> harf "
        f"(A-F, variantsiz savol uchun x):",
        parse_mode=ParseMode.HTML
    )
    await state.update_data(stagedit_field="answers")
    await state.set_state(StagedEditState.collecting)


async def _staged_edit_ask_question_file(answerable, state: FSMContext) -> None:
    await answerable.answer(
        "Endi savollar bo'lgan PDF faylni yuboring. Eskisini saqlab qolish uchun <code>-</code> yozing:",
        parse_mode=ParseMode.HTML
    )
    await state.update_data(stagedit_field="question_file")
    await state.set_state(StagedEditState.collecting)


@router.message(StagedEditState.collecting, F.document | F.audio | F.video)
async def staged_edit_file(message: Message, state: FSMContext) -> None:
    data = await state.get_data()
    field = data.get("stagedit_field")
    if field not in ("file", "question_file"):
        return

    cfg = data["stagedit_cfg"]
    key = data["stagedit_key"]

    if field == "question_file":
        file_id = message.document.file_id if message.document else (
            message.audio.file_id if message.audio else message.video.file_id
        )
        cfg["sections"][key]["question_file_id"] = file_id
        await state.update_data(stagedit_cfg=cfg)
        await _staged_edit_ask_qcount(message, state)
        return

    if message.document:
        file_id, file_kind = message.document.file_id, "document"
    elif message.audio:
        file_id, file_kind = message.audio.file_id, "audio"
    else:
        file_id, file_kind = message.video.file_id, "video"

    cfg["sections"][key]["file_id"] = file_id
    cfg["sections"][key]["file_kind"] = file_kind
    await state.update_data(stagedit_cfg=cfg)
    if key == "listening":
        await _staged_edit_ask_question_file(message, state)
    else:
        await _staged_edit_ask_qcount(message, state)


@router.message(StagedEditState.collecting)
async def staged_edit_text(message: Message, state: FSMContext) -> None:
    if (message.text or "") in TOP_MENU_BTNS:
        return
    if not message.text:
        return

    data = await state.get_data()
    field = data.get("stagedit_field")
    cfg = data["stagedit_cfg"]
    key = data["stagedit_key"]
    test_id = data["stagedit_test_id"]
    section = cfg["sections"][key]
    text = message.text.strip()

    if field == "duration":
        if not text.isdigit() or not (1 <= int(text) <= 300):
            return await message.answer("⚠️ 1..300 oralig'idagi raqam kiriting:")
        section["duration_min"] = int(text)
        cfg["sections"][key] = section
        await state.update_data(stagedit_cfg=cfg)
        if section["type"] in (STAGE_TYPE_MCQ, STAGE_TYPE_MCQ_BAND):
            if key == "listening":
                await message.answer(
                    "Yangi AUDIO/VIDEO faylni yuboring. Eskisini saqlab qolish uchun <code>-</code> yozing:",
                    parse_mode=ParseMode.HTML
                )
            else:
                await message.answer(
                    "Yangi faylni yuboring. Eskisini saqlab qolish uchun <code>-</code> yozing:",
                    parse_mode=ParseMode.HTML
                )
            await state.update_data(stagedit_field="file")
        else:
            await message.answer(
                f"Hozirgi topshiriq/mavzu matni:\n<i>{_h(section.get('prompt', '-'))}</i>\n\n"
                f"Yangisini kiriting:",
                parse_mode=ParseMode.HTML
            )
            await state.update_data(stagedit_field="prompt")
        return

    if field == "file":
        if text != "-":
            return await message.answer("⚠️ Fayl yuboring yoki eskisini saqlash uchun \"-\" yozing:")
        if key == "listening":
            await _staged_edit_ask_question_file(message, state)
        else:
            await _staged_edit_ask_qcount(message, state)
        return

    if field == "question_file":
        if text != "-":
            return await message.answer("⚠️ Fayl yuboring yoki eskisini saqlash uchun \"-\" yozing:")
        await _staged_edit_ask_qcount(message, state)
        return

    if field == "qcount":
        if text != "-":
            if not text.isdigit() or not (1 <= int(text) <= 300):
                return await message.answer("⚠️ 1..300 oralig'idagi raqam yoki \"-\" kiriting:")
            if int(text) != int(section.get("qcount") or 0) and await POOL.fetchval(
                    "SELECT 1 FROM staged_attempts WHERE test_id=$1 AND sections_json LIKE $2 LIMIT 1",
                    test_id, f'%"{key}"%'):
                return await message.answer(
                    "⚠️ Bu bosqichni allaqachon yechganlar bor — savollar sonini o'zgartirib bo'lmaydi. "
                    "O'zgartirmaslik uchun \"-\" yuboring:")
            section["qcount"] = int(text)
            cfg["sections"][key] = section
            await state.update_data(stagedit_cfg=cfg)
        await _staged_edit_ask_answers(message, state)
        return

    if field == "answers":
        val = normalize_answers(text)
        qcount = int(section.get("qcount") or 0)
        if not is_answer_string_valid(val):
            return await message.answer("⚠️ Faqat A-F yoki x harflaridan iborat bo'lsin:")
        if len(val) != qcount:
            return await message.answer(
                f"⚠️ Aynan {qcount} ta belgi bo'lishi kerak, siz {len(val)} ta yubordingiz:")

        open_positions = open_question_positions(val)
        if open_positions:
            await state.update_data(
                stagedit_pending_answers=val, stagedit_pending_open_positions=open_positions,
                stagedit_field="open_ref"
            )
            await message.answer(_open_ref_prompt_text(open_positions), parse_mode=ParseMode.HTML)
            return

        section["answers"] = val
        section.pop("open_ref_answers", None)
        cfg["sections"][key] = section
        await POOL.execute(
            "UPDATE tests SET skills_cfg=$1 WHERE id=$2", json.dumps(cfg, ensure_ascii=False), test_id
        )
        regraded = await _regrade_staged_section(test_id, key)
        admin_flag = await is_admin_user(message.from_user.id)
        await state.clear()
        await message.answer(
            f"✅ <b>{_h(section['name'])}</b> bosqichi yangilandi."
            + (f"\n♻️ {regraded} ta urinish yangi kalit bo'yicha qayta hisoblandi." if regraded else ""),
            parse_mode=ParseMode.HTML, reply_markup=main_menu_kb(admin_flag)
        )
        return

    if field == "open_ref":
        positions = data.get("stagedit_pending_open_positions") or []
        ref_map = parse_open_ref_answers(text, positions)
        if ref_map is None:
            await message.answer(
                f"⚠️ Qatorlar soni mos kelmadi — aynan <b>{len(positions)} qator</b> yuboring "
                f"(har birida bitta javob, savollar tartibida):",
                parse_mode=ParseMode.HTML
            )
            return
        section["answers"] = data.get("stagedit_pending_answers")
        section["open_ref_answers"] = ref_map
        cfg["sections"][key] = section
        await POOL.execute(
            "UPDATE tests SET skills_cfg=$1 WHERE id=$2", json.dumps(cfg, ensure_ascii=False), test_id
        )
        regraded = await _regrade_staged_section(test_id, key)
        admin_flag = await is_admin_user(message.from_user.id)
        await state.clear()
        await message.answer(
            f"✅ <b>{_h(section['name'])}</b> bosqichi yangilandi."
            + (f"\n♻️ {regraded} ta urinish yangi kalit bo'yicha qayta hisoblandi." if regraded else ""),
            parse_mode=ParseMode.HTML, reply_markup=main_menu_kb(admin_flag)
        )
        return

    if field == "prompt":
        if len(text) < 3:
            return await message.answer("⚠️ Juda qisqa. Qayta kiriting:")
        section["prompt"] = text
        cfg["sections"][key] = section
        await state.update_data(stagedit_cfg=cfg)

        if GEMINI_API_KEY:
            kb = InlineKeyboardMarkup(inline_keyboard=[[
                InlineKeyboardButton(text="🤖 AI avtomatik tekshirsin", callback_data="edtg_grader_ai"),
                InlineKeyboardButton(text="👤 O'zim/ustoz qo'lda", callback_data="edtg_grader_manual"),
            ]])
            await message.answer(f"<b>{section['name']}</b> javoblarini kim baholasin?",
                                 parse_mode=ParseMode.HTML, reply_markup=kb)
            return

        section["grader"] = "manual"
        cfg["sections"][key] = section
        await POOL.execute(
            "UPDATE tests SET skills_cfg=$1 WHERE id=$2", json.dumps(cfg, ensure_ascii=False), test_id
        )
        admin_flag = await is_admin_user(message.from_user.id)
        await state.clear()
        await message.answer(
            f"✅ <b>{section['name']}</b> bosqichi yangilandi.",
            parse_mode=ParseMode.HTML, reply_markup=main_menu_kb(admin_flag)
        )
        return


@router.callback_query(StagedEditState.collecting, F.data.startswith("edtg_grader_"))
async def staged_edit_grader_choice(callback: CallbackQuery, state: FSMContext) -> None:
    choice = callback.data.split("_")[2]
    data = await state.get_data()
    cfg = data["stagedit_cfg"]
    key = data["stagedit_key"]
    test_id = data["stagedit_test_id"]
    section = cfg["sections"][key]
    section["grader"] = "ai" if choice == "ai" else "manual"
    cfg["sections"][key] = section

    await POOL.execute("UPDATE tests SET skills_cfg=$1 WHERE id=$2", json.dumps(cfg, ensure_ascii=False), test_id)
    admin_flag = await is_admin_user(callback.from_user.id)
    await state.clear()
    await callback.answer("🤖 AI tanlandi" if choice == "ai" else "👤 Qo'lda tekshirish tanlandi")
    await callback.message.answer(
        f"✅ <b>{section['name']}</b> bosqichi yangilandi.",
        parse_mode=ParseMode.HTML, reply_markup=main_menu_kb(admin_flag)
    )


# ===================== ADMIN: TESTNI YASHIRISH/QAYTARISH =====================
async def _render_hidetest_list(answerable, page: int = 0):
    tests, has_more = await _fetch_tests_page(page)
    if not tests and page == 0:
        return await answerable.answer("Testlar yo'q.")

    kb = InlineKeyboardMarkup(inline_keyboard=[])
    for r in tests:
        mark = "⏸ Yashirilgan" if r["is_hidden"] else "✅ Faol"
        kb.inline_keyboard.append([InlineKeyboardButton(
            text=f"{mark} — {r['code']}", callback_data=f"hidetest_toggle_{r['id']}")])
    nav = _pagination_nav_row("hide", page, has_more)
    if nav:
        kb.inline_keyboard.append(nav)

    await answerable.answer(
        f"⏸ <b>Testlarni yashirish/qaytarish</b> ({page + 1}-sahifa)\n"
        "Yashirilgan test ro'yxatlarda ko'rinmaydi va uni yangi topshiruvchi sotib ololmaydi "
        "(lekin allaqachon tasdiqlangan xaridorlar davom eta oladi). Bosilganda holati almashadi:",
        parse_mode=ParseMode.HTML, reply_markup=kb
    )


@router.callback_query(F.data == "hidetest_menu")
async def hidetest_menu(callback: CallbackQuery):
    if not await is_admin_user(callback.from_user.id):
        return
    await _render_hidetest_list(callback.message, 0)


@router.callback_query(F.data.startswith("hidetest_toggle_"))
async def hidetest_toggle(callback: CallbackQuery):
    if not await is_admin_user(callback.from_user.id):
        return
    test_id = int(callback.data.split("_")[2])
    row = await POOL.fetchrow("SELECT is_hidden, code FROM tests WHERE id=$1", test_id)
    if not row:
        return await callback.answer("Topilmadi", show_alert=True)

    new_val = 0 if int(row["is_hidden"] or 0) == 1 else 1
    await POOL.execute("UPDATE tests SET is_hidden=$1 WHERE id=$2", new_val, test_id)
    status = "yashirildi ⏸" if new_val == 1 else "qaytarildi ✅"
    await callback.answer(f"{row['code']} {status}", show_alert=False)


# ===================== ADMIN: BUGUNGI HISOBOT =====================
@router.callback_query(F.data == "today_digest")
async def today_digest(callback: CallbackQuery):
    if not await is_admin_user(callback.from_user.id):
        return

    today_epoch = int(datetime.now(TZ).replace(hour=0, minute=0, second=0, microsecond=0).timestamp())
    today_str = datetime.now(TZ).strftime("%Y-%m-%d") + " 00:00"

    new_users = await POOL.fetchval("SELECT COUNT(*) FROM users WHERE created_ts >= $1", today_epoch)
    results_today = await POOL.fetchval("SELECT COUNT(*) FROM results WHERE date >= $1", today_str)
    payments_today = await POOL.fetchval("SELECT COUNT(*) FROM payments WHERE created_at >= $1", today_str)
    revenue_today = await POOL.fetchval("""
        SELECT COALESCE(SUM(COALESCE(p.amount, t.price)),0) FROM payments p JOIN tests t ON t.id=p.test_id
        WHERE p.status='approved' AND COALESCE(p.approved_at, p.created_at) >= $1
    """, today_str)

    text = (
        f"📅 <b>Bugungi hisobot</b> ({datetime.now(TZ).strftime('%Y-%m-%d')})\n\n"
        f"🆕 Yangi ro'yxatdan o'tganlar: <b>{new_users}</b>\n"
        f"📄 Topshirilgan testlar: <b>{results_today}</b>\n"
        f"💳 Kelgan to'lov so'rovlari: <b>{payments_today}</b>\n"
        f"💰 Bugungi tushum: <b>{_fmt_money(revenue_today)}</b>"
    )
    await callback.message.answer(text, parse_mode=ParseMode.HTML)


# ===================== ADMIN: TESTLAR REYTINGI =====================
@router.callback_query(F.data == "tests_ranking")
async def tests_ranking(callback: CallbackQuery):
    if not await is_admin_user(callback.from_user.id):
        return

    rows = await POOL.fetch("""
        SELECT t.code, t.title, COUNT(DISTINCT r.user_id) AS participants, AVG(r.percent) AS avgp
        FROM tests t LEFT JOIN results r ON r.test_id=t.id
        GROUP BY t.id, t.code, t.title
        ORDER BY participants DESC
        LIMIT 10
    """)

    text = "🏆 <b>Eng mashhur testlar</b> (ishtirokchilar soni bo'yicha)\n\n"
    shown = 0
    for i, r in enumerate(rows, 1):
        if r["participants"] == 0:
            continue
        shown += 1
        text += f"{shown}) <code>{r['code']}</code> {r['title']} — {r['participants']} kishi (o'rtacha {float(r['avgp'] or 0):.1f}%)\n"

    if shown == 0:
        text += "Hozircha hech kim test topshirmagan."

    await callback.message.answer(text, parse_mode=ParseMode.HTML)


# ===================== ADMIN: CHANNELS (ADD/EDIT/DELETE/TOGGLE) =====================
async def _bot_admin_status_in_chat(bot: Bot, chat_id: int) -> bool | None:
    """
    Bot shu kanal/guruhda administrator/creator ekanini tekshiradi.
    True/False = aniq javob, None = tekshirib bo'lmadi (masalan bot u yerda umuman yo'q).
    Bu tekshiruv MUHIM: agar bot administrator bo'lmasa, u a'zolarning
    holatini ko'ra olmaydi va HAQIQIY a'zolarni ham "a'zo emas" deb ko'rsatadi.
    """
    try:
        me = await bot.get_me()
        member = await bot.get_chat_member(chat_id=chat_id, user_id=me.id)
        return member.status in ("administrator", "creator")
    except Exception as e:
        logging.warning(f"Bot holatini tekshirib bo'lmadi (chat_id={chat_id}): {e}")
        return None


async def _save_required_channel(chat_id: int, username: str | None, title: str | None, join_url: str | None):
    await POOL.execute("""
        INSERT INTO required_channels (chat_id, username, title, join_url, is_active)
        VALUES ($1, $2, $3, $4, 1)
        ON CONFLICT(chat_id) DO UPDATE SET
            username=excluded.username,
            title=excluded.title,
            join_url=excluded.join_url,
            is_active=1
    """, chat_id, username, title, join_url or None)


@router.message(F.text == BTN_ADMIN_CHANNELS)
async def admin_channels_panel(message: Message, state: FSMContext, bot: Bot):
    if not await is_admin_user(message.from_user.id):
        return

    rows = await get_required_channels(active_only=False)
    text = (
        "<b>📣 Majburiy obuna — kanal/guruhlar</b>\n\n"
        "<i>Bu yerga qo'shilgan kanal/guruhlarga a'zo bo'lmagan foydalanuvchi botdan foydalana olmaydi.</i>\n\n"
    )
    if not rows:
        text += "Hozircha hech qanday kanal/guruh qo'shilmagan."
    else:
        warn_count = 0
        for row in rows:
            _id, chat_id, username, title, join_url, is_active = (
                row["id"], row["chat_id"], row["username"], row["title"], row["join_url"], row["is_active"]
            )
            st = "✅ Yoqilgan" if int(is_active or 0) == 1 else "⛔ O'chirilgan"
            ref = f"@{username}" if username else (str(chat_id) if chat_id else "N/A")
            link = (join_url or "").strip()

            admin_ok = await _bot_admin_status_in_chat(bot, int(chat_id)) if chat_id else None
            warn_line = ""
            if admin_ok is False or admin_ok is None:
                warn_count += 1
                warn_line = (
                    "\n⚠️ <b>Bot bu yerda administrator emas!</b> A'zolikni tekshira olmayapti — "
                    "haqiqiy a'zolarni ham \"a'zo emas\" deb ko'rsatishi mumkin."
                )

            text += (
                f"\n🆔 <code>{_id}</code> | {st}\n"
                f"<b>{_h(title or 'Nomsiz')}</b> | <code>{_h(ref)}</code>"
                f"{warn_line}"
            )
            if link:
                text += f"\n🔗 {_h(link)}"
            text += "\n"

        if warn_count:
            text += f"\n⚠️ <b>{warn_count} ta kanal/guruhda bot admin emas — tuzatish kerak.</b>"

    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="➕ Kanal/guruh qo'shish", callback_data="ch_add")],
        [InlineKeyboardButton(text="🧹 Tahrirlash / o'chirish", callback_data="ch_manage")],
    ])
    await message.answer(text, parse_mode=ParseMode.HTML, reply_markup=kb)
    await state.clear()


@router.callback_query(F.data == "ch_add")
async def ch_add_start(callback: CallbackQuery, state: FSMContext):
    if not await is_admin_user(callback.from_user.id):
        return
    await callback.message.answer(
        "➕ <b>Majburiy kanal/guruh qo'shish</b>\n\n"
        "<b>Diqqat — bu 2 ta bosqichda amalga oshadi:</b>\n"
        "1️⃣ Avval kanal/guruhga botni <b>administrator</b> qilib qo'shing "
        "(oddiy a'zo emas — aynan admin, bo'lmasa bot a'zolikni tekshira olmaydi).\n"
        "2️⃣ Shundan keyin shu yerga kanal/guruh manzilini yuboring:\n\n"
        "• Ochiq (public) bo'lsa — username: <code>@mening_kanalim</code>\n"
        "• Yopiq (private) bo'lsa — chat ID: <code>-1001234567890</code>\n"
        "  <i>(ID ni bilish uchun kanalga istalgan xabarni forward qiling @userinfobot ga)</i>\n\n"
        "Yuboring:",
        parse_mode=ParseMode.HTML,
        reply_markup=nav_kb(include_back=True)
    )
    await state.set_state(AdminChannelState.waiting_channel_ref)


@router.message(AdminChannelState.waiting_channel_ref)
async def ch_add_got_ref(message: Message, state: FSMContext, bot: Bot):
    if not await is_admin_user(message.from_user.id):
        await state.clear()
        return
    if (message.text or "") in TOP_MENU_BTNS:
        return

    ref_raw = (message.text or "").strip()
    ref = _normalize_channel_ref(ref_raw)

    try:
        if _looks_like_int(ref):
            chat = await bot.get_chat(int(ref))
        else:
            chat = await bot.get_chat(f"@{ref}")
    except Exception:
        await message.answer(
            "❌ Bunday kanal/guruh topilmadi.\n\n"
            "Tekshiring: username to'g'rimi, yoki bot shu kanal/guruhga umuman qo'shilganmi? "
            "Qayta yuboring:"
        )
        return

    chat_id = int(chat.id)
    username = chat.username or None
    title = chat.title or "Nomsiz"

    # ✅ Oddiy foydalanuvchi (shaxsiy chat) majburiy "kanal" bo'la olmaydi — unga
    # hech kim "a'zo" bo'la olmaydi va bot BARCHA o'quvchilarni bloklab qo'yardi.
    if str(getattr(chat.type, "value", chat.type)) not in ("channel", "group", "supergroup"):
        await message.answer(
            "❌ Bu kanal yoki guruh emas (oddiy foydalanuvchi akkaunti). Majburiy obuna uchun faqat "
            "kanal yoki guruh qo'shish mumkin. Boshqa manzil yuboring:")
        return

    admin_ok = await _bot_admin_status_in_chat(bot, chat_id)
    if admin_ok is not True:
        await message.answer(
            f"⚠️ <b>Diqqat:</b> bot <b>{_h(title)}</b> ichida hozircha administrator emas "
            f"(yoki umuman a'zo emas).\n\n"
            f"Iltimos, botni o'sha yerga <b>administrator</b> qilib qo'shing — aks holda "
            f"a'zolikni tekshira olmaydi va haqiqiy a'zolarni ham \"a'zo emas\" deb ko'rsatadi.\n\n"
            f"Buni keyinroq ham tuzatishingiz mumkin, hozircha davom etamiz.",
            parse_mode=ParseMode.HTML
        )

    # ✅ Avtomatik join-link: ochiq kanal bo'lsa username orqali, yopiq bo'lsa
    # (bot admin bo'lsa) taklif havolasini o'zi yaratadi — adminга qo'lda
    # havola qidirib, joylashtirish shart bo'lmaydi.
    auto_join_url = ""
    if username:
        auto_join_url = f"https://t.me/{username}"
    elif admin_ok is True:
        try:
            auto_join_url = await bot.export_chat_invite_link(chat_id)
        except Exception as e:
            logging.warning(f"Invite link avtomatik yaratilmadi (chat_id={chat_id}): {e}")
            auto_join_url = ""

    if auto_join_url:
        await _save_required_channel(chat_id, username, title, auto_join_url)
        await state.clear()
        await message.answer(
            f"✅ Qo'shildi: <b>{_h(title)}</b>\n🔗 Havola avtomatik aniqlandi: {auto_join_url}",
            parse_mode=ParseMode.HTML,
            reply_markup=main_menu_kb(True)
        )
        return

    # Yopiq guruh/kanal va havolani o'zimiz yarata olmadik — admindan so'raymiz.
    await state.update_data(chat_id=chat_id, username=username, title=title)
    await message.answer(
        "🔗 <b>Taklif havolasini yuboring</b>\n\n"
        "Bu yopiq kanal/guruh bo'lgani uchun havolani avtomatik topa olmadim.\n"
        "Guruh/kanal sozlamalaridan invite link nusxalab yuboring "
        "(yoki havolasiz davom etish uchun <code>-</code> yuboring):",
        parse_mode=ParseMode.HTML,
        reply_markup=nav_kb(include_back=True)
    )
    await state.set_state(AdminChannelState.waiting_join_url)


@router.message(AdminChannelState.waiting_join_url)
async def ch_add_save(message: Message, state: FSMContext):
    if not await is_admin_user(message.from_user.id):
        await state.clear()
        return
    if (message.text or "") in TOP_MENU_BTNS:
        return

    join_url = (message.text or "").strip()
    if join_url == "-":
        join_url = ""

    if join_url and not _is_valid_join_url(join_url):
        await message.answer("⚠️ Havola noto'g'ri ko'rinishda. Masalan: https://t.me/+AbCdEfGh. Qayta yuboring:")
        return

    data = await state.get_data()
    chat_id = int(data["chat_id"])
    username = data.get("username")
    title = data.get("title")

    await _save_required_channel(chat_id, username, title, join_url or None)

    await state.clear()
    await message.answer(f"✅ Qo'shildi: <b>{_h(title)}</b>", parse_mode=ParseMode.HTML, reply_markup=main_menu_kb(True))


@router.callback_query(F.data == "ch_manage")
async def ch_manage(callback: CallbackQuery):
    if not await is_admin_user(callback.from_user.id):
        return

    rows = await get_required_channels(active_only=False)
    if not rows:
        return await callback.message.answer("Ro‘yxat bo‘sh.")

    kb = InlineKeyboardMarkup(inline_keyboard=[])
    for row in rows:
        _id, chat_id, username, title, join_url, is_active = (
            row["id"], row["chat_id"], row["username"], row["title"], row["join_url"], row["is_active"]
        )
        st = "ON✅" if int(is_active or 0) == 1 else "OFF⛔"
        name = title or (f"@{username}" if username else str(chat_id))
        kb.inline_keyboard.append([
            InlineKeyboardButton(text=f"{st} {name}", callback_data=f"ch_toggle_{_id}"),
            InlineKeyboardButton(text="✏️", callback_data=f"ch_edit_{_id}"),
            InlineKeyboardButton(text="🗑", callback_data=f"ch_del_{_id}")
        ])

    await callback.message.answer("🧹 <b>Tahrirlash</b>\nON/OFF, ✏️ tahrirlash yoki 🗑 o‘chirish:",
                                  parse_mode=ParseMode.HTML, reply_markup=kb)


@router.callback_query(F.data.startswith("ch_toggle_"))
async def ch_toggle(callback: CallbackQuery):
    if not await is_admin_user(callback.from_user.id):
        return
    _id = int(callback.data.split("_")[2])

    row = await POOL.fetchrow("SELECT is_active FROM required_channels WHERE id=$1", _id)
    if not row:
        return await callback.answer("Topilmadi", show_alert=True)

    new_val = 0 if int(row["is_active"] or 0) == 1 else 1
    await POOL.execute("UPDATE required_channels SET is_active=$1 WHERE id=$2", new_val, _id)

    await callback.answer("✅ Yangilandi", show_alert=False)


@router.callback_query(F.data.startswith("ch_del_"))
async def ch_delete(callback: CallbackQuery):
    if not await is_admin_user(callback.from_user.id):
        return
    _id = int(callback.data.split("_")[2])

    await POOL.execute("DELETE FROM required_channels WHERE id=$1", _id)

    await callback.answer("🗑 O‘chirildi", show_alert=False)


@router.callback_query(F.data.startswith("ch_edit_"))
async def ch_edit_start(callback: CallbackQuery, state: FSMContext):
    if not await is_admin_user(callback.from_user.id):
        return
    _id = int(callback.data.split("_")[2])
    await state.update_data(edit_id=_id)
    await callback.message.answer(
        "✏️ <b>Tahrirlash</b>\n\n"
        "2 qatorda yuboring:\n"
        "1) Yangi title (o‘zgartirmasangiz <code>-</code>)\n"
        "2) Yangi join link:\n"
        "   • o‘zgartirmasangiz <code>-</code>\n"
        "   • tozalash (o‘chirish) uchun <code>clear</code>\n\n"
        "Misol (faqat linkni tozalash):\n"
        "<code>-</code>\n"
        "<code>clear</code>\n\n"
        "Misol (title+link yangilash):\n"
        "<code>DTM Kanal</code>\n"
        "<code>https://t.me/dtm_demo</code>",
        parse_mode=ParseMode.HTML,
        reply_markup=nav_kb(include_back=True)
    )
    await state.set_state(AdminChannelState.waiting_edit_payload)


@router.message(AdminChannelState.waiting_edit_payload)
async def ch_edit_save(message: Message, state: FSMContext):
    if not await is_admin_user(message.from_user.id):
        await state.clear()
        return
    if (message.text or "") in TOP_MENU_BTNS:
        return

    lines = [x.strip() for x in (message.text or "").splitlines() if x.strip()]
    if len(lines) < 2:
        await message.answer("⚠️ 2 qatorda yuboring: (title) va (link). Qayta yuboring:")
        return

    new_title_raw = lines[0]
    new_link_raw = lines[1]

    new_title = None if new_title_raw == "-" else new_title_raw

    keep_link = False
    clear_link = False
    new_link = None

    if new_link_raw == "-":
        keep_link = True
    elif new_link_raw.lower() == "clear":
        clear_link = True
    else:
        if not _is_valid_join_url(new_link_raw):
            await message.answer("⚠️ Link noto‘g‘ri. https://... bo‘lishi kerak. Qayta yuboring:")
            return
        new_link = new_link_raw

    data = await state.get_data()
    _id = int(data.get("edit_id") or 0)

    row = await POOL.fetchrow("SELECT title, join_url FROM required_channels WHERE id=$1", _id)
    if not row:
        await state.clear()
        await message.answer("Topilmadi.")
        return

    old_title, old_link = row["title"], row["join_url"]

    final_title = new_title if new_title is not None else old_title

    if keep_link:
        final_link = old_link
    elif clear_link:
        final_link = ""
    else:
        final_link = new_link

    await POOL.execute(
        "UPDATE required_channels SET title=$1, join_url=$2 WHERE id=$3",
        final_title, final_link, _id
    )

    await state.clear()
    await message.answer("✅ Tahrirlandi.", reply_markup=main_menu_kb(True))


# ===================== ALL USERS: ADD TEST =====================
async def _nav_show_subject(answerable, state: FSMContext) -> None:
    await answerable.answer(
        "🆕 <b>Yangi test yaratish</b>\n\n"
        "1-qadam: bu test qaysi fanga tegishli?\n"
        "Masalan: <i>Matematika, Ingliz tili, Tarix, Ona tili...</i>",
        parse_mode=ParseMode.HTML,
        reply_markup=nav_kb(include_back=True)
    )
    await state.set_state(AdminAddTestState.subject)


@router.message(F.text == BTN_ADD_TEST)
async def admin_add_test(message: Message, state: FSMContext):
    is_admin = await is_admin_user(message.from_user.id)

    if not is_admin:
        today_start = int(datetime.now(TZ).replace(hour=0, minute=0, second=0, microsecond=0).timestamp())
        count = await POOL.fetchval(
            "SELECT COUNT(*) FROM tests WHERE creator_id=$1 AND created_ts>=$2",
            message.from_user.id, today_start
        )
        if count >= 10:
            return await message.answer(
                "⚠️ Bugungi kunlik limitingiz (10 ta test) tugadi.\nErtaga yana test yaratishingiz mumkin bo'ladi.")

    await state.update_data(nav_stack=[])
    await _nav_show_subject(message, state)


async def generate_unique_code():
    while True:
        code = "T-" + "".join(random.choices(string.ascii_uppercase + string.digits, k=5))
        if not await test_code_exists(code):
            return code


async def _nav_show_title(answerable, state: FSMContext) -> None:
    await answerable.answer(
        "2-qadam: test nomini kiriting.\nMasalan: <i>DTM Blok 2024</i>",
        parse_mode=ParseMode.HTML,
        reply_markup=nav_kb(include_back=True)
    )
    await state.set_state(AdminAddTestState.title)


@router.message(AdminAddTestState.subject)
async def set_subject(message: Message, state: FSMContext):
    if (message.text or "") in TOP_MENU_BTNS:
        return
    subject = sanitize_plain_text(message.text, max_len=80)
    if len(subject) < 2:
        return await message.answer("⚠️ Fan nomi juda qisqa yoki noto'g'ri belgilar bor. Qayta kiriting:")
    await state.update_data(subject=subject)
    await _nav_push(state, "base", "subject")
    await _nav_show_title(message, state)


async def _nav_show_is_free(answerable, state: FSMContext) -> None:
    data = await state.get_data()
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="🆓 Tekin test", callback_data="setfree_1")],
        [InlineKeyboardButton(text="🧾 Pullik test", callback_data="setfree_0")]
    ])
    await answerable.answer(
        f"✅ Test kodi: <code>{data.get('code', '')}</code>\n\n"
        f"3-qadam: test turi qanday bo'lsin?",
        parse_mode=ParseMode.HTML, reply_markup=kb)
    await state.set_state(AdminAddTestState.is_free)


@router.message(AdminAddTestState.title)
async def set_title(message: Message, state: FSMContext):
    if (message.text or "") in TOP_MENU_BTNS:
        return
    title = sanitize_plain_text(message.text, max_len=150)
    if len(title) < 2:
        return await message.answer("⚠️ Test nomi juda qisqa yoki noto'g'ri belgilar bor. Qayta kiriting:")
    await state.update_data(title=title)

    # ✅ Test kodini admin qo'lda o'ylab topmaydi — takrorlanmas kod avtomatik yaratiladi.
    new_code = await generate_unique_code()
    await state.update_data(code=new_code)

    await _nav_push(state, "base", "title")
    await _nav_show_is_free(message, state)


async def _nav_show_price(answerable, state: FSMContext) -> None:
    await answerable.answer(
        "💰 Narxni so'mda kiriting (masalan: 25000):",
        reply_markup=nav_kb(include_back=True)
    )
    await state.set_state(AdminAddTestState.price)


@router.callback_query(AdminAddTestState.is_free, F.data.startswith("setfree_"))
async def set_is_free(callback: CallbackQuery, state: FSMContext):
    is_free = int(callback.data.split("_")[1])
    await state.update_data(is_free=is_free)
    await _nav_push(state, "base", "is_free")

    if is_free == 1:
        await callback.message.answer("✅ Tekin test tanlandi.")
        await _ask_exam_type(callback.message, state)
    else:
        await callback.message.answer("✅ Pullik test tanlandi.")
        await _nav_show_price(callback.message, state)


@router.message(AdminAddTestState.price)
async def set_price(message: Message, state: FSMContext):
    if (message.text or "") in TOP_MENU_BTNS:
        return
    if not (message.text or "").isdigit():
        return await message.answer("⚠️ Narxni faqat raqamda kiriting (masalan: 25000):")
    await state.update_data(price=int(message.text))

    if not await is_admin_user(message.from_user.id):
        await message.answer(
            "💡 <i>Diqqat: to'lov administratorning umumiy kartasiga tushadi. "
            "Tasdiqlangan to'lovdan sizga tegishli ulushni admin o'zi o'tkazib beradi.</i>",
            parse_mode=ParseMode.HTML
        )

    await _nav_push(state, "base", "price")
    await _ask_exam_type(message, state)


# ✅ AVTOMATLASHTIRISH: avval "davomiylik" va "savollar soni" ikkita alohida
# savol sifatida so'ralardi. Bu ikkalasi deyarli har doim birga kiritilgani
# uchun endi bitta qadamga birlashtirildi — tez-tez ishlatiladigan
# kombinatsiyalar tugma orqali 1 marta bosib tanlanadi, xohlasa admin
# o'zi ham "daqiqa savol_soni" shaklida yozishi mumkin.
DURATION_QCOUNT_PRESETS = [(20, 20), (30, 30), (45, 45), (60, 60)]


def duration_qcount_kb():
    kb = InlineKeyboardMarkup(inline_keyboard=[])
    row = []
    for d, q in DURATION_QCOUNT_PRESETS:
        row.append(InlineKeyboardButton(text=f"{d} daq / {q} ta", callback_data=f"qdq_{d}_{q}"))
        if len(row) == 2:
            kb.inline_keyboard.append(row)
            row = []
    if row:
        kb.inline_keyboard.append(row)
    return kb


async def _ask_duration_and_qcount(answerable, state: FSMContext):
    await answerable.answer(
        "5-qadam: <b>test davomiyligi va savollar soni</b>\n\n"
        "Tayyor variantlardan birini tanlang,\n"
        "yoki o'zingiz <b>«daqiqa savol_soni»</b> shaklida yozing.\n"
        "Masalan: <code>40 25</code> — 40 daqiqa, 25 ta savol.",
        parse_mode=ParseMode.HTML,
        reply_markup=duration_qcount_kb()
    )
    await state.set_state(AdminAddTestState.duration)


async def _proceed_to_schedule_step(answerable, state: FSMContext) -> None:
    data = await state.get_data()
    et = data.get("exam_type") or "simple"

    # ✅ MUHIM: Rasch testlarda barcha qatnashuvchilarning natijasi BIR VAQTDA
    # (guruh bo'yicha) hisoblanadi, shuning uchun ular albatta "rejalashtirilgan"
    # (hammaga bir xil boshlanish/tugash vaqti) rejimda bo'lishi SHART.
    # Aks holda kech boshlagan foydalanuvchilarning natijasi hech qachon
    # Rasch balga aylantirilmay qolib ketardi (chunki hisoblash faqat 1 marta,
    # birinchi topshirgan kishining tugash vaqtiga qarab ishga tushadi).
    if et == "rasch":
        await state.update_data(start_mode="scheduled")
        await answerable.answer(
            "ℹ️ Rasch testlarda barcha o'quvchilarning natijasi bir vaqtda (guruh bo'yicha) "
            "hisoblanadi, shuning uchun bu turdagi testlar faqat <b>rejalashtirilgan</b> "
            "(hammaga bir xil boshlanish vaqti) rejimda ishlaydi — bu qadam avtomatik tanlandi.",
            parse_mode=ParseMode.HTML
        )
        await _ask_schedule_datetime(answerable, state)
        return

    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="🚀 Oddiy (o'quvchi istalgan vaqtda boshlaydi)", callback_data="sch_normal")],
        [InlineKeyboardButton(text="📅 Rejalashtirilgan (hammaga bir xil vaqt)", callback_data="sch_scheduled")]
    ])
    await answerable.answer("6-qadam: <b>test qachon boshlanadi?</b>", parse_mode=ParseMode.HTML, reply_markup=kb)
    await state.set_state(AdminAddTestState.schedule_mode)


async def _save_duration_and_advance(answerable, state: FSMContext, duration_val: int, qcount_val: int) -> None:
    data = await state.get_data()
    if data.get("exam_type") == "dtm" and data.get("dtm_cfg"):
        qcount_val = dtm_total_questions(json.loads(data["dtm_cfg"]))
    await state.update_data(duration=duration_val, questions_count=qcount_val)
    await _nav_push(state, "base", "duration")
    await _proceed_to_schedule_step(answerable, state)


@router.callback_query(AdminAddTestState.duration, F.data.startswith("qdq_"))
async def pick_duration_and_qcount(callback: CallbackQuery, state: FSMContext):
    _, d, q = callback.data.split("_")
    await callback.answer()
    await _save_duration_and_advance(callback.message, state, int(d), int(q))


@router.message(AdminAddTestState.duration)
async def set_duration_and_qcount(message: Message, state: FSMContext):
    if (message.text or "") in TOP_MENU_BTNS:
        return
    parts = (message.text or "").split()
    if len(parts) != 2 or not all(p.isdigit() for p in parts):
        return await message.answer(
            "⚠️ Format noto'g'ri. Ikkita raqamni bo'sh joy bilan ajratib yuboring.\n"
            "Masalan: <code>40 25</code> (40 daqiqa, 25 ta savol)",
            parse_mode=ParseMode.HTML
        )
    duration_val, qcount_val = int(parts[0]), int(parts[1])
    if duration_val < 1 or duration_val > 600:
        return await message.answer("⚠️ Davomiylik 1..600 daqiqa oralig'ida bo'lsin.")
    if qcount_val < 1 or qcount_val > 300:
        return await message.answer("⚠️ Savollar soni 1..300 oralig'ida bo'lsin.")

    await _save_duration_and_advance(message, state, duration_val, qcount_val)


async def _ask_exam_type(answerable, state: FSMContext):
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="📝 Oddiy test", callback_data="etype_simple")],
        [InlineKeyboardButton(text="⚡ Rasch (MS)", callback_data="etype_rasch")],
        [InlineKeyboardButton(text="📚 DTM (5 fan, 1 sessiya)", callback_data="etype_dtm")],
        [InlineKeyboardButton(text="🎯 DTM (5 bosqich, alohida vaqt)", callback_data="etype_dtmstaged")],
        [InlineKeyboardButton(text="🎓 CEFR (ko'nikmalar, alohida vaqt)", callback_data="etype_cefr4")],
        [InlineKeyboardButton(text="🎧 IELTS (ko'nikmalar, ustoz/AI baholaydi)", callback_data="etype_ielts4")],
        [InlineKeyboardButton(text="📝 SAT (tezkor, 1 ball)", callback_data="etype_sat")],
        [InlineKeyboardButton(text="🧮 SAT (2 bosqich: R&W + Math)", callback_data="etype_satstaged")],
    ])
    await answerable.answer(
        "4-qadam: <b>baholash rejimini tanlang</b>\n\n"
        "📝 <b>Oddiy</b> — har savol 1 ball, natija oddiy foizda.\n"
        "⚡ <b>Rasch (MS)</b> — natija guruhga nisbatan statistik baholanadi "
        "(milliy sertifikat uslubiga yaqin); istalgan guruh kattaligida (1 kishidan "
        "minglab kishigacha) ishlaydi.\n"
        "📚 <b>DTM (5 fan, 1 sessiya)</b> — 90 savol bitta yaxlit vaqt ichida.\n"
        "🎯 <b>DTM (5 bosqich)</b> — xuddi shu 5 fan, lekin har biri O'Z ALOHIDA "
        "vaqtida ketma-ket topshiriladi (haqiqiy DTM tartibiga yaqin); xohlagan fan(lar)ni "
        "o'tkazib yuborish ham mumkin.\n"
        "🎓 <b>CEFR</b> — Reading/Writing/Speaking/Listening'dan xohlaganlarini tanlaysiz "
        "(bittasini tansangiz ham bo'ladi) — har biri o'z vaqtida, natija A1–C2 darajasida.\n"
        "🎧 <b>IELTS</b> — xuddi CEFR kabi, lekin natija raqamli band (0–9) ko'rinishida; "
        "Writing/Speaking'ni ustoz (yoki yoqilgan bo'lsa AI) baholaydi.\n"
        "📝 <b>SAT (tezkor)</b> — bitta test, natija darhol taxminiy 400–1600 ballga o'tkaziladi.\n"
        "🧮 <b>SAT (2 bosqich)</b> — rasmiy SAT tuzilishiga mos: Reading&Writing va Math "
        "ALOHIDA vaqtda topshiriladi (xohlasangiz faqat bittasi ham bo'ladi), har biri "
        "200-800 shkalada baholanadi, ikkalasi qo'shilib 400-1600 kompozit ball hosil qiladi.\n\n"
        "<i>CEFR/IELTS/SAT konvertatsiyasi taxminiy (foizdan hisoblanadi) — rasmiy sertifikat emas.</i>",
        parse_mode=ParseMode.HTML,
        reply_markup=kb
    )
    await state.set_state(AdminAddTestState.exam_type)


@router.callback_query(AdminAddTestState.exam_type, F.data.startswith("etype_"))
async def set_exam_type(callback: CallbackQuery, state: FSMContext):
    et = callback.data.split("_", 1)[1]  # simple/rasch/dtm/cefr/ielts/sat
    await _nav_push(state, "base", "exam_type")

    # ✅ Bosqichli rejimlar (DTM 5-bosqich, IELTS 4-ko'nikma) butunlay boshqa
    # oqim: fayl/javob kaliti bitta emas, har bosqichda alohida. Shu sabab
    # keyingi (schedule_mode/file/answers) qadamlarni chetlab o'tib, o'z
    # sehrgariga o'tkazamiz — u yakunida testni o'zi saqlaydi.
    if et == "dtmstaged":
        await callback.message.answer(
            "✅ <b>DTM (5 bosqich)</b> tanlandi. Har bir fan uchun alohida so'rayman.",
            parse_mode=ParseMode.HTML
        )
        await start_staged_wizard(callback, state, "dtm_staged")
        return
    if et == "ielts4":
        await callback.message.answer(
            "✅ <b>IELTS (4 ko'nikma)</b> tanlandi. Har bir ko'nikma uchun alohida so'rayman.",
            parse_mode=ParseMode.HTML
        )
        await start_staged_wizard(callback, state, "ielts4")
        return
    if et == "cefr4":
        await callback.message.answer(
            "✅ <b>CEFR (4 ko'nikma)</b> tanlandi. Har bir ko'nikma uchun alohida so'rayman.",
            parse_mode=ParseMode.HTML
        )
        await start_staged_wizard(callback, state, "cefr4")
        return
    if et == "satstaged":
        await callback.message.answer(
            "✅ <b>SAT (2 bosqich)</b> tanlandi. Reading&Writing va Math uchun alohida so'rayman.",
            parse_mode=ParseMode.HTML
        )
        await start_staged_wizard(callback, state, "sat_staged")
        return

    if et == "dtm":
        cfg = DTM_DEFAULT_CFG
        await state.update_data(exam_type="dtm", dtm_cfg=json.dumps(cfg, ensure_ascii=False))
        await state.update_data(questions_count=dtm_total_questions(cfg))

        await callback.message.answer(
            "✅ <b>DTM (5 fan)</b> rejimi tanlandi.\n"
            "1-2-fan 30 tadan, 3-4-5-fan 10 tadan savol — jami "
            f"<b>{dtm_total_questions(cfg)} ta</b> savol (ilgari kiritilgan savollar soni "
            "bu rejim uchun avtomatik almashtirildi).\n"
            "Ball og'irliklari: 3.1 / 2.1 / 1.1 / 1.1 / 1.1",
            parse_mode=ParseMode.HTML
        )

    elif et == "rasch":
        await state.update_data(exam_type="rasch", dtm_cfg=None)
        await callback.message.answer("✅ <b>Rasch (MS)</b> rejimi tanlandi.", parse_mode=ParseMode.HTML)
    elif et == "sat":
        await state.update_data(exam_type="sat", dtm_cfg=None)
        await callback.message.answer(
            "✅ <b>SAT (tezkor)</b> rejimi tanlandi. Natija testni yakunlagach darhol "
            "(taxminiy SAT ko'rinishida) chiqadi.",
            parse_mode=ParseMode.HTML
        )
    else:
        await state.update_data(exam_type="simple", dtm_cfg=None)
        await callback.message.answer("✅ <b>Oddiy</b> rejim tanlandi.", parse_mode=ParseMode.HTML)

    await _ask_duration_and_qcount(callback.message, state)


async def _nav_show_file(answerable, state: FSMContext) -> None:
    await answerable.answer(
        "7-qadam: test faylini (PDF) yuboring:",
        parse_mode=ParseMode.HTML,
        reply_markup=nav_kb(include_back=True)
    )
    await state.set_state(AdminAddTestState.file)


@router.callback_query(AdminAddTestState.schedule_mode, F.data == "sch_normal")
async def schedule_normal(callback: CallbackQuery, state: FSMContext):
    await state.update_data(start_mode="normal", start_ts=0, start_at=None)
    await _nav_push(state, "base", "schedule_mode")
    await _nav_show_file(callback.message, state)


@router.callback_query(AdminAddTestState.schedule_mode, F.data == "sch_scheduled")
async def schedule_scheduled(callback: CallbackQuery, state: FSMContext):
    await state.update_data(start_mode="scheduled")
    await _nav_push(state, "base", "schedule_mode")
    await _ask_schedule_datetime(callback.message, state)


async def _ask_schedule_datetime(answerable, state: FSMContext):
    await answerable.answer(
        "📅 <b>Boshlanish sanasi va vaqtini kiriting</b>\n\n"
        "Format: <code>YYYY-MM-DD HH:MM</code>\n"
        "Masalan: <code>2026-02-15 16:00</code>",
        parse_mode=ParseMode.HTML,
        reply_markup=nav_kb(include_back=True)
    )
    await state.set_state(AdminAddTestState.schedule_date)


@router.message(AdminAddTestState.schedule_date)
async def set_schedule_datetime(message: Message, state: FSMContext):
    if (message.text or "") in TOP_MENU_BTNS:
        return

    parts = (message.text or "").strip().split(maxsplit=1)
    if len(parts) != 2:
        return await message.answer(
            "⚠️ Format noto'g'ri. Sana va vaqtni bo'sh joy bilan ajratib, bitta xabarda yuboring.\n"
            "Masalan: <code>2026-02-15 16:00</code>",
            parse_mode=ParseMode.HTML
        )

    date_str, time_str = parts[0], parts[1]
    res = to_start_ts(date_str, time_str)
    if not res:
        return await message.answer(
            "⚠️ Sana yoki vaqt formati noto'g'ri.\n"
            "Sana: <code>YYYY-MM-DD</code>, vaqt: <code>HH:MM</code>.\n"
            "Masalan: <code>2026-02-15 16:00</code>",
            parse_mode=ParseMode.HTML
        )

    start_ts_val, start_at = res
    data = await state.get_data()
    problem = schedule_window_problem(start_ts_val, data.get("duration"))
    if problem:
        return await message.answer(problem)
    await state.update_data(start_ts=start_ts_val, start_at=start_at)

    already = " (⚠️ bu vaqt allaqachon boshlangan — o'quvchilar qolgan vaqtda yechadi)" if start_ts_val < now_ts() else ""
    await message.answer(
        f"✅ Rejalashtirildi.\n📅 Boshlanish: <b>{start_at}</b>{already}",
        parse_mode=ParseMode.HTML
    )
    await _nav_push(state, "base", "schedule_date")
    await _nav_show_file(message, state)


async def _nav_show_answers(answerable, state: FSMContext) -> None:
    data = await state.get_data()
    qcount = int(data.get("questions_count") or 30)
    await answerable.answer(
        f"8-qadam (oxirgi): <b>javoblar kalitini kiriting.</b>\n"
        f"Aynan <b>{qcount} ta</b> harf, bo'shliqsiz, savollar tartibida.\n"
        f"Faqat <b>A, B, C, D, E, F</b> harflaridan foydalaning (masalan: <code>abcdabcd...</code>)\n\n"
        f"<i>Variantsiz (ochiq) savol bo'lsa, o'sha o'rniga <code>x</code> harfini qo'ying — "
        f"o'quvchi shu savolga matn bilan javob yozadi, avtomatik baholanmaydi, "
        f"siz alohida ko'rib chiqasiz.</i>",
        parse_mode=ParseMode.HTML,
        reply_markup=nav_kb(include_back=True)
    )
    await state.set_state(AdminAddTestState.answers)


@router.message(AdminAddTestState.file, F.document)
async def set_file(message: Message, state: FSMContext):
    await state.update_data(file_id=message.document.file_id)
    await message.answer("✅ Fayl qabul qilindi.")
    await _nav_push(state, "base", "file")
    await _nav_show_answers(message, state)


@router.message(AdminAddTestState.answers)
async def save_test(message: Message, state: FSMContext):
    if (message.text or "") in TOP_MENU_BTNS:
        return

    data = await state.get_data()
    answers = normalize_answers(message.text)
    qcount = int(data.get("questions_count") or 30)

    if not is_answer_string_valid(answers):
        await message.answer(
            "⚠️ Javob kaliti faqat <b>A, B, C, D, E, F</b> harflaridan iborat bo'lishi kerak "
            "(interaktiv tugmalar shulardan iborat). Qayta yuboring:",
            parse_mode=ParseMode.HTML)
        return

    if len(answers) != qcount:
        await message.answer(
            f"⚠️ Javoblar soni noto‘g‘ri.\n"
            f"Kerakli: <b>{qcount}</b>, Siz yubordingiz: <b>{len(answers)}</b>.\n\n"
            f"Iltimos, aniq <b>{qcount} ta</b> javob yuboring:",
            parse_mode=ParseMode.HTML
        )
        return

    open_positions = open_question_positions(answers)
    if open_positions:
        await state.update_data(pending_answers=answers, pending_open_positions=open_positions)
        await _nav_push(state, "base", "answers")
        await message.answer(_open_ref_prompt_text(open_positions), parse_mode=ParseMode.HTML)
        await state.set_state(AdminAddTestState.open_ref_answers)
        return

    await _finish_base_test_creation(message, state, answers, None)


@router.message(AdminAddTestState.open_ref_answers)
async def save_test_open_refs(message: Message, state: FSMContext):
    if (message.text or "") in TOP_MENU_BTNS:
        return

    data = await state.get_data()
    positions = data.get("pending_open_positions") or []
    ref_map = parse_open_ref_answers(message.text or "", positions)
    if ref_map is None:
        await message.answer(
            f"⚠️ Qatorlar soni mos kelmadi — aynan <b>{len(positions)} qator</b> yuboring "
            f"(har birida bitta javob, savollar tartibida):",
            parse_mode=ParseMode.HTML
        )
        return

    await _finish_base_test_creation(message, state, data.get("pending_answers"), json.dumps(ref_map, ensure_ascii=False))


async def _finish_base_test_creation(
    message: Message, state: FSMContext, answers: str, open_ref_json: str | None
) -> None:
    data = await state.get_data()
    qcount = int(data.get("questions_count") or 30)
    is_free = int(data.get("is_free") or 0)
    price = 0 if is_free else int(data.get("price") or 0)

    start_mode = data.get("start_mode") or "normal"
    start_ts_val = int(data.get("start_ts") or 0)
    start_at = data.get("start_at")

    exam_type = data.get("exam_type") or "simple"
    dtm_cfg = data.get("dtm_cfg")

    new_test_id = await POOL.fetchval("""
        INSERT INTO tests (
            code, title, subject, price, duration, questions_count, file_id, answers, is_free,
            start_mode, start_ts, start_at,
            exam_type, dtm_cfg, creator_id, created_ts, open_ref_answers
        )
        VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13, $14, $15, $16, $17)
        RETURNING id
    """,
        data.get("code"),
        data.get("title"),
        data.get("subject"),
        price,
        int(data.get("duration") or 0),
        qcount,
        data.get("file_id"),
        answers,
        is_free,
        start_mode,
        start_ts_val,
        start_at,
        exam_type,
        dtm_cfg,
        message.from_user.id,
        now_ts(),
        open_ref_json,
    )

    extra = ""
    if start_mode == "scheduled":
        extra = f"\n📅 Boshlanish: <b>{start_at}</b>"

    admin_flag = await is_admin_user(message.from_user.id)
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="🎓 Sertifikat sozlamalari (ixtiyoriy)", callback_data=f"certcfg_start_{new_test_id}")]
    ])
    await message.answer(
        f"✅ Test qo‘shildi!\n🔑 Kod: <code>{data.get('code')}</code>\n🧮 Savollar: <b>{qcount}</b>{extra}\n\n"
        f"<i>Xohlasangiz, bu test uchun sertifikat ko'rinishini (o'quv markazi nomi, mas'ul shaxs, rangi) "
        f"sozlashingiz mumkin — aks holda standart sozlamalar ishlatiladi.</i>",
        parse_mode=ParseMode.HTML,
        reply_markup=kb
    )
    await message.answer("🏠 Bosh menyu", reply_markup=main_menu_kb(admin_flag))
    await state.clear()


# ✅ Har bir node "shu qadamni qayta ko'rsat" funksiyasiga ishora qiladi —
# nav_back (yuqorida) shu jadval orqali "🔄" bosilganda qaysi promptni qayta
# chizishni topadi. Ba'zilari maxsus "_nav_show_*" funksiyalari, ba'zilari
# esa allaqachon mavjud "_ask_*" funksiyalarni qayta ishlatadi.
BASE_NAV_HANDLERS = {
    "subject": _nav_show_subject,
    "title": _nav_show_title,
    "is_free": _nav_show_is_free,
    "price": _nav_show_price,
    "exam_type": _ask_exam_type,
    "duration": _ask_duration_and_qcount,
    "schedule_mode": _proceed_to_schedule_step,
    "schedule_date": _ask_schedule_datetime,
    "file": _nav_show_file,
    "answers": _nav_show_answers,
}


# ===================== SERTIFIKAT SOZLAMALARI (test yaratuvchisi/admin) =====================
async def _cert_settings_entry_check(user_tg_id: int, test_id: int) -> bool:
    t = await POOL.fetchrow("SELECT creator_id FROM tests WHERE id=$1", test_id)
    if not t:
        return False
    return t["creator_id"] == user_tg_id or await is_admin_user(user_tg_id)


async def _nav_show_certcfg_academy(answerable, state: FSMContext) -> None:
    await answerable.answer(
        f"🎓 <b>1/3 — O'quv markazi nomi</b>\n\n"
        f"Sertifikatda ko'rinadigan nomni kiriting.\n"
        f"O'zgartirmaslik uchun <code>-</code> yuboring (standart: <b>{ACADEMY_NAME}</b>):",
        parse_mode=ParseMode.HTML,
        reply_markup=nav_kb(include_back=True)
    )
    await state.set_state(CertConfigState.waiting_academy)


@router.callback_query(F.data.startswith("certcfg_start_"))
async def certcfg_start(callback: CallbackQuery, state: FSMContext):
    test_id = int(callback.data.split("_")[2])
    if not await _cert_settings_entry_check(callback.from_user.id, test_id):
        return await callback.answer("Bu testga huquqingiz yo'q!", show_alert=True)

    await state.update_data(certcfg_test_id=test_id, nav_stack=[])
    await _nav_show_certcfg_academy(callback.message, state)


async def _nav_show_certcfg_signee(answerable, state: FSMContext) -> None:
    await answerable.answer(
        "🎓 <b>2/3 — Mas'ul shaxs</b>\n\n"
        "Sertifikatda «Mas'ul shaxs» sifatida yoziladigan ism-familiyani kiriting.\n"
        "O'zgartirmaslik uchun <code>-</code> yuboring (standart: sizning ismingiz):",
        parse_mode=ParseMode.HTML,
        reply_markup=nav_kb(include_back=True)
    )
    await state.set_state(CertConfigState.waiting_signee)


# ✅ BASE_NAV_HANDLERS yuqorida (test yaratish bo'limi oxirida) allaqachon
# e'lon qilingan — bu ikkalasi o'sha jadvalga QO'SHILADI, chunki ularning
# funksiyalari o'sha jadval yozilgan paytda hali mavjud emas edi (Python
# modul darajasidagi lug'at o'z qatorida DARHOL baholanadi, funksiyalarni
# keyinroq e'lon qilib bo'lmaydi).
BASE_NAV_HANDLERS["certcfg_academy"] = _nav_show_certcfg_academy
BASE_NAV_HANDLERS["certcfg_signee"] = _nav_show_certcfg_signee


@router.message(CertConfigState.waiting_academy)
async def certcfg_got_academy(message: Message, state: FSMContext):
    if (message.text or "") in TOP_MENU_BTNS:
        return
    val = sanitize_plain_text(message.text, max_len=100)
    if val and val != "-":
        await state.update_data(certcfg_academy=val)

    await _nav_push(state, "base", "certcfg_academy")
    await _nav_show_certcfg_signee(message, state)


@router.message(CertConfigState.waiting_signee)
async def certcfg_got_signee(message: Message, state: FSMContext, bot: Bot):
    if (message.text or "") in TOP_MENU_BTNS:
        return
    val = sanitize_plain_text(message.text, max_len=100)
    if val and val != "-":
        await state.update_data(certcfg_signee=val)
    await _nav_push(state, "base", "certcfg_signee")

    data = await state.get_data()
    test_id = data.get("certcfg_test_id")
    academy = data.get("certcfg_academy")
    signee = data.get("certcfg_signee")
    if not signee:
        my_uid = await get_user_db_id(message.from_user.id)
        signee = await get_user_full_name(my_uid) if my_uid else None

    await message.answer("⏳ <b>3/3 — Namunalar tayyorlanmoqda...</b>", parse_mode=ParseMode.HTML)

    tmp_dir = tempfile.gettempdir()
    for tpl_key, tpl_meta in CERT_TEMPLATES.items():
        pdf_path = os.path.join(tmp_dir, f"certpreview_{test_id}_{tpl_key}.pdf")
        make_certificate_pdf(
            full_name="Ism Familiya",
            subject="Namunaviy fan",
            test_code="NAMUNA-01",
            test_title="Namunaviy test",
            exam_date_str=now_str_local(),
            percent=92.0, correct_count=23, total_questions=25,
            rasch_ball=64.5, grade="A",
            academy_name=academy, signee_name=signee, template=tpl_key,
            out_path=pdf_path,
        )
        kb = InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text=f"✅ {tpl_meta['label']} — shuni tanlash",
                                  callback_data=f"certcfg_pick_{test_id}_{tpl_key}")]
        ])
        await bot.send_document(
            chat_id=message.from_user.id,
            document=FSInputFile(pdf_path),
            caption=tpl_meta["label"],
            reply_markup=kb
        )

    await message.answer("👆 Yoqqan ko'rinishni yuqoridagi 3 ta namunadan tanlang.")
    await state.set_state(CertConfigState.waiting_template)


@router.callback_query(CertConfigState.waiting_template, F.data.startswith("certcfg_pick_"))
async def certcfg_pick(callback: CallbackQuery, state: FSMContext):
    parts = callback.data.split("_")
    test_id, tpl_key = int(parts[2]), parts[3]
    if not await _cert_settings_entry_check(callback.from_user.id, test_id):
        return await callback.answer("Bu testga huquqingiz yo'q!", show_alert=True)

    data = await state.get_data()
    academy = data.get("certcfg_academy")
    signee = data.get("certcfg_signee")

    await POOL.execute(
        "UPDATE tests SET cert_academy_name=$1, cert_signee_name=$2, cert_template=$3 WHERE id=$4",
        academy, signee, tpl_key, test_id
    )
    await state.clear()
    await callback.answer("✅ Saqlandi!", show_alert=False)
    await callback.message.answer(
        f"✅ <b>Sertifikat sozlamalari saqlandi:</b> {CERT_TEMPLATES[tpl_key]['label']}.\n"
        f"Bu testni yechganlarning barchasiga shu ko'rinishdagi sertifikat beriladi.",
        parse_mode=ParseMode.HTML
    )


# ===================== ALL USERS: MY CREATED TESTS =====================
async def _render_created_tests(answerable, user_tg_id: int, page: int = 0):
    offset = max(0, page) * USER_LIST_PAGE_SIZE
    rows = await POOL.fetch("""
        SELECT id, code, title, is_free, price, created_ts
        FROM tests
        WHERE creator_id=$1
        ORDER BY id DESC LIMIT $2 OFFSET $3
    """, user_tg_id, USER_LIST_PAGE_SIZE + 1, offset)
    has_more = len(rows) > USER_LIST_PAGE_SIZE
    rows = rows[:USER_LIST_PAGE_SIZE]

    if not rows:
        if page == 0:
            return await answerable.answer("Siz hali test yaratmagansiz.")
        return await answerable.answer("Boshqa test yo'q.")

    text = f"🛠 <b>Siz yaratgan testlar</b> ({page + 1}-sahifa):\n\n"
    kb = InlineKeyboardMarkup(inline_keyboard=[])

    for r in rows:
        tid, code, title, is_free, price = r["id"], r["code"], r["title"], r["is_free"], r["price"]
        kind = "🆓 Tekin" if is_free else f"🧾 {price} so'm"
        text += f"🔑 <code>{code}</code> — <b>{title}</b> ({kind})\n\n"
        kb.inline_keyboard.append(
            [InlineKeyboardButton(text=f"📊 Statistika: {code}", callback_data=f"mycreatestat_{tid}")])

    nav = _user_pagination_nav_row("created", page, has_more)
    if nav:
        kb.inline_keyboard.append(nav)

    await answerable.answer(text, parse_mode=ParseMode.HTML, reply_markup=kb)


@router.message(F.text == BTN_MY_CREATED_TESTS)
async def list_my_created_tests(message: Message):
    await _render_created_tests(message, message.from_user.id, 0)


@router.callback_query(F.data.startswith("mycreatestat_"))
async def creator_test_stat_cb(callback: CallbackQuery):
    test_id = int(callback.data.split("_")[1])
    user_tg_id = callback.from_user.id

    t = await POOL.fetchrow("SELECT code, title, exam_type, creator_id FROM tests WHERE id=$1", test_id)

    if not t or t["creator_id"] != user_tg_id:
        return await callback.answer("Bu test sizga tegishli emas!", show_alert=True)

    code, title, exam_type = t["code"], t["title"], t["exam_type"]

    res = await POOL.fetchrow(
        "SELECT COUNT(DISTINCT user_id) AS c, AVG(percent) AS avgp, AVG(score) AS avgs FROM results WHERE test_id=$1",
        test_id
    )

    participants = res["c"] if res and res["c"] else 0
    avg_pct = res["avgp"] if res and res["avgp"] else 0.0
    avg_score = res["avgs"] if res and res["avgs"] else 0.0

    text = (
        f"📊 <b>Test statistikasi</b>\n\n"
        f"🔑 Kod: <code>{code}</code>\n"
        f"📚 Nomi: <b>{title}</b>\n\n"
        f"👥 Qatnashuvchilar: <b>{participants}</b>\n"
        f"📈 O'rtacha foiz: <b>{avg_pct:.1f}%</b>\n"
        f"🎯 O'rtacha to'g'ri javob: <b>{avg_score:.1f}</b>\n"
    )

    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="📤 Excel eksport", callback_data=f"xl_{test_id}")]
    ])
    if exam_type == 'rasch':
        kb.inline_keyboard.append([InlineKeyboardButton(text="⚡ Rasch hisobot", callback_data=f"raschrep_{test_id}")])
    open_pending_n = await POOL.fetchval(
        "SELECT COUNT(*) FROM results WHERE test_id=$1 AND COALESCE(open_pending,0)>0", test_id)
    if open_pending_n:
        kb.inline_keyboard.append([
            InlineKeyboardButton(text=f"✍️ Ochiq javoblarni tekshirish ({open_pending_n})",
                                 callback_data=f"ogl_{test_id}")
        ])
    has_open_q = "x" in normalize_answers(
        await POOL.fetchval("SELECT answers FROM tests WHERE id=$1", test_id) or "")
    if has_open_q:
        kb.inline_keyboard.append([
            InlineKeyboardButton(text="🔍 Ochiq javoblarni ko'rish (avto-baho)", callback_data=f"oga_{test_id}")
        ])
    if exam_type in STAGED_EXAM_TYPES:
        pending_n = await POOL.fetchval(
            "SELECT COUNT(*) FROM staged_attempts WHERE test_id=$1 AND status='pending_review'", test_id
        )
        kb.inline_keyboard.append([
            InlineKeyboardButton(text=f"🎓 Baholash kutilayotganlar ({pending_n})", callback_data=f"skg_pending_{test_id}")
        ])
    if exam_type in STAGED_EXAM_TYPES:
        kb.inline_keyboard.append([
            InlineKeyboardButton(text="📋 Ishtirokchilar (batafsil)", callback_data=f"stagepart_{test_id}_0")
        ])
    kb.inline_keyboard.append([
        InlineKeyboardButton(text="✏️ Tahrirlash", callback_data=f"admedit_pick_{test_id}"),
        InlineKeyboardButton(text="🗑 O'chirish", callback_data=f"admdel_ask_{test_id}"),
    ])
    kb.inline_keyboard.append([
        InlineKeyboardButton(text="🛑 Yakunlash/Qayta ochish", callback_data=f"closetest_toggle_{test_id}"),
    ])
    kb.inline_keyboard.append([
        InlineKeyboardButton(text="🎓 Sertifikat sozlamalari", callback_data=f"certcfg_start_{test_id}"),
    ])

    await callback.message.answer(text, parse_mode=ParseMode.HTML, reply_markup=kb)
    await callback.answer()


# ----- Bosqichli testlar: ishtirokchilar ro'yxati va batafsil ko'rish -----
STAGE_STATUS_LABELS = {
    "in_progress": "🟢 Yechmoqda",
    "pending_review": "🟡 Baholanmoqda",
    "graded": "✅ Tugatgan",
}
STAGE_PARTICIPANTS_PAGE_SIZE = 10


async def _render_stage_participants(answerable, test_id: int, page: int = 0) -> None:
    offset = max(0, page) * STAGE_PARTICIPANTS_PAGE_SIZE
    rows = await POOL.fetch("""
        SELECT sa.id, sa.status, sa.attempt_no, sa.overall_grade, u.full_name
        FROM staged_attempts sa
        JOIN users u ON u.id = sa.user_id
        WHERE sa.test_id=$1
        ORDER BY sa.id DESC
        LIMIT $2 OFFSET $3
    """, test_id, STAGE_PARTICIPANTS_PAGE_SIZE + 1, offset)
    has_more = len(rows) > STAGE_PARTICIPANTS_PAGE_SIZE
    rows = rows[:STAGE_PARTICIPANTS_PAGE_SIZE]

    t = await POOL.fetchrow("SELECT title, code FROM tests WHERE id=$1", test_id)
    if not t:
        await answerable.answer("Test topilmadi.")
        return
    if not rows:
        await answerable.answer("Hozircha hech kim bu testni boshlamagan." if page == 0 else "Boshqa ishtirokchi yo'q.")
        return

    counts = await POOL.fetch(
        "SELECT status, COUNT(*) AS c FROM staged_attempts WHERE test_id=$1 GROUP BY status", test_id
    )
    count_map = {r["status"]: r["c"] for r in counts}

    text = (
        f"📋 <b>{t['title']}</b> (<code>{t['code']}</code>) — ishtirokchilar ({page + 1}-sahifa)\n\n"
        f"🟢 Yechmoqda: <b>{count_map.get('in_progress', 0)}</b> | "
        f"🟡 Baholanmoqda: <b>{count_map.get('pending_review', 0)}</b> | "
        f"✅ Tugatgan: <b>{count_map.get('graded', 0)}</b>\n\n"
        f"Batafsil ko'rish uchun kimnidir tanlang:"
    )

    kb = InlineKeyboardMarkup(inline_keyboard=[])
    for r in rows:
        status_label = STAGE_STATUS_LABELS.get(r["status"], r["status"])
        attempt_suffix = f" (#{r['attempt_no']})" if r["attempt_no"] and r["attempt_no"] > 1 else ""
        grade_part = f" — {r['overall_grade']}" if r["status"] == "graded" and r["overall_grade"] else ""
        btn_text = f"{status_label} · {r['full_name']}{attempt_suffix}{grade_part}"
        kb.inline_keyboard.append([InlineKeyboardButton(text=btn_text[:64], callback_data=f"stagedetail_{r['id']}")])

    nav = []
    if page > 0:
        nav.append(InlineKeyboardButton(text="⬅️ Oldingi", callback_data=f"stagepart_{test_id}_{page - 1}"))
    if has_more:
        nav.append(InlineKeyboardButton(text="Keyingi ➡️", callback_data=f"stagepart_{test_id}_{page + 1}"))
    if nav:
        kb.inline_keyboard.append(nav)

    await answerable.answer(text, parse_mode=ParseMode.HTML, reply_markup=kb)


@router.callback_query(F.data.startswith("stagepart_"))
async def stage_participants_nav(callback: CallbackQuery) -> None:
    parts = callback.data.split("_")
    test_id, page = int(parts[1]), int(parts[2])
    if not await _is_admin_or_creator(callback.from_user.id, test_id):
        await callback.answer("Sizda ruxsat yo'q!", show_alert=True)
        return
    await callback.answer()
    await _render_stage_participants(callback.message, test_id, page)


@router.callback_query(F.data.startswith("stagedetail_"))
async def stage_attempt_detail(callback: CallbackQuery) -> None:
    attempt_id = int(callback.data.split("_")[1])
    row = await POOL.fetchrow("""
        SELECT sa.status, sa.attempt_no, sa.overall_grade, sa.sections_json,
               t.skills_cfg, t.code, t.title, t.creator_id, u.full_name
        FROM staged_attempts sa
        JOIN tests t ON t.id = sa.test_id
        JOIN users u ON u.id = sa.user_id
        WHERE sa.id=$1
    """, attempt_id)
    if not row:
        await callback.answer("Topilmadi.", show_alert=True)
        return
    if not (await is_admin_user(callback.from_user.id) or row["creator_id"] == callback.from_user.id):
        await callback.answer("Sizda ruxsat yo'q!", show_alert=True)
        return

    cfg = _load_skills_cfg(row["skills_cfg"])
    sections = json.loads(row["sections_json"] or "{}")
    order = cfg.get("order", [])
    section_cfg = cfg.get("sections", {})

    status_label = STAGE_STATUS_LABELS.get(row["status"], row["status"])
    attempt_suffix = f" (#{row['attempt_no']})" if row["attempt_no"] and row["attempt_no"] > 1 else ""

    header = (
        f"📋 <b>{row['full_name']}</b>{attempt_suffix}\n"
        f"📚 {row['title']} (<code>{row['code']}</code>)\n"
        f"Holat: {status_label}\n"
    )
    if row["overall_grade"]:
        header += f"🏆 Umumiy natija: <b>{row['overall_grade']}</b>\n"

    await callback.answer()
    await callback.message.answer(header, parse_mode=ParseMode.HTML)

    for k in order:
        sec = section_cfg.get(k, {})
        v = sections.get(k, {})
        label = sec.get("name", k)
        stype = sec.get("type")

        if not v:
            await callback.message.answer(f"⏳ <b>{label}</b> — hali yechilmagan.", parse_mode=ParseMode.HTML)
            continue

        if stype in (STAGE_TYPE_MCQ, STAGE_TYPE_MCQ_BAND):
            key_str = normalize_answers(sec.get("answers", ""))
            user_ans = normalize_answers(v.get("user_answers", ""))
            open_refs = sec.get("open_ref_answers") or {}
            open_student_answers = v.get("open_answers") or {}
            mistakes = []
            open_lines = []
            open_photos = []
            qn = min(len(key_str), len(user_ans))
            for i in range(qn):
                if key_str[i] == "x":
                    raw = open_student_answers.get(str(i + 1))
                    if raw:
                        text_val, photo_id = _open_answer_preview(raw)
                        if photo_id:
                            open_photos.append((i + 1, photo_id))
                        line = f"{i + 1}-savol (ochiq): <i>{text_val or '[rasm yuborilgan]'}</i>"
                        ref = open_refs.get(str(i + 1))
                        if ref:
                            line += f"\n   ✅ Kutilgan javob: <b>{ref}</b>"
                        open_lines.append(line)
                    continue
                if user_ans[i] != key_str[i]:
                    miss = "belgilamagan" if user_ans[i] == "*" else user_ans[i].upper()
                    mistakes.append(f"{i + 1}-savol: siz {miss}, to'g'ri {key_str[i].upper()}")

            msg = f"📖 <b>{label}</b>: {v.get('score', 0)}/{v.get('total', 0)} — <b>{v.get('band', '-')}</b>\n\n"
            if mistakes:
                msg += "\n".join(mistakes[:30])
                if len(mistakes) > 30:
                    msg += f"\n... yana {len(mistakes) - 30} ta"
            else:
                msg += "Xatolar yo'q ✅"
            if open_lines:
                msg += "\n\n✍️ <b>Ochiq savol javoblari:</b>\n" + "\n".join(open_lines[:20])
            await callback.message.answer(msg, parse_mode=ParseMode.HTML)
            for qnum, photo_id in open_photos[:20]:
                await callback.message.answer_photo(photo_id, caption=f"{qnum}-savol javobi (rasm)")

        elif stype == STAGE_TYPE_TEXT:
            msg = f"✍️ <b>{label}</b>"
            if v.get("band") is not None:
                msg += f" — band <b>{v['band']}</b>"
            msg += "\n"
            if v.get("feedback"):
                msg += f"💬 {_h(v['feedback'])}\n"
            await callback.message.answer(msg, parse_mode=ParseMode.HTML)
            if v.get("photo_id"):
                await callback.message.answer_photo(v["photo_id"], caption=f"{label} — rasm javob")
            if v.get("text"):
                await callback.message.answer(f"<i>{_h(v['text'])}</i>", parse_mode=ParseMode.HTML)

        elif stype == STAGE_TYPE_VOICE:
            msg = f"🎙 <b>{label}</b>"
            if v.get("band") is not None:
                msg += f" — band <b>{v['band']}</b>"
            msg += "\n"
            if v.get("feedback"):
                msg += f"💬 {_h(v['feedback'])}\n"
            await callback.message.answer(msg, parse_mode=ParseMode.HTML)
            if v.get("voice_id"):
                await callback.message.answer_voice(v["voice_id"])


async def _send_delete_confirm(answerable, test_id: int):
    """
    O'chirish tugmasi bosilganda ko'rsatiladigan ogohlantirish + tasdiqlash
    ekrani. Ikki turdagi xavfni alohida ko'rsatadi:
    1) hozir shu testni faol yechayotganlar bo'lsa (sessions jadvali)
    2) avval topshirgan/natija qoldirganlar bo'lsa (results jadvali)
    Har ikkala holatda ham o'chirish FAQAT aniq tasdiqlashdan keyin bajariladi.
    """
    t = await POOL.fetchrow("SELECT code, title FROM tests WHERE id=$1", test_id)
    if not t:
        return await answerable.answer("Test topilmadi.")

    participants = await POOL.fetchval("SELECT COUNT(DISTINCT user_id) FROM results WHERE test_id=$1", test_id)
    active_now = await POOL.fetchval(
        "SELECT COUNT(*) FROM sessions WHERE test_id=$1 AND start_ts + duration_sec > $2",
        test_id, now_ts()
    )

    warn = ""
    if active_now:
        warn += (f"\n\n🔴 <b>Diqqat: hozir {active_now} kishi shu testni yechayapti!</b> "
                 f"O'chirsangiz ularning jarayoni to'xtab qoladi.")
    if participants:
        warn += (f"\n\n⚠️ Bu testni allaqachon <b>{participants}</b> kishi topshirgan — "
                 f"ularning natijalari va sertifikatlari ham butunlay o'chib ketadi!")

    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="✅ Ha, butunlay o'chirilsin", callback_data=f"admdel_confirm_{test_id}")],
        [InlineKeyboardButton(text="❌ Bekor qilish", callback_data="dismiss_msg")],
    ])
    await answerable.answer(
        f"🗑 <b>{t['title']}</b> (<code>{t['code']}</code>) butunlay o'chiriladi.{warn}\n\n"
        f"Bu amalni ortga qaytarib bo'lmaydi. Tasdiqlaysizmi?",
        parse_mode=ParseMode.HTML, reply_markup=kb
    )


@router.callback_query(F.data.startswith("admdel_ask_"))
async def admdel_ask(callback: CallbackQuery):
    test_id = int(callback.data.split("_")[2])
    if not await _is_admin_or_creator(callback.from_user.id, test_id):
        return await callback.answer("Bu testga huquqingiz yo'q!", show_alert=True)
    await _send_delete_confirm(callback.message, test_id)


@router.callback_query(F.data.startswith("admdel_confirm_"))
async def admdel_confirm(callback: CallbackQuery):
    test_id = int(callback.data.split("_")[2])
    if not await _is_admin_or_creator(callback.from_user.id, test_id):
        return await callback.answer("Bu testga huquqingiz yo'q!", show_alert=True)

    t = await POOL.fetchrow("SELECT code, title FROM tests WHERE id=$1", test_id)
    if not t:
        return await callback.answer("Test topilmadi", show_alert=True)

    async with POOL.acquire() as db:
        async with db.transaction():
            await db.execute("DELETE FROM results WHERE test_id=$1", test_id)
            await db.execute("DELETE FROM payments WHERE test_id=$1", test_id)
            await db.execute("DELETE FROM sessions WHERE test_id=$1", test_id)
            await db.execute("DELETE FROM staged_attempts WHERE test_id=$1", test_id)
            await db.execute("DELETE FROM tests WHERE id=$1", test_id)

    await callback.message.answer(f"✅ <code>{t['code']}</code> — <b>{t['title']}</b> o'chirildi.", parse_mode=ParseMode.HTML)


# ===================== SLASH (/) BUYRUQLAR MENYUSI =====================
# ✅ Telegram'dagi "/" tugmasi bosilganda chiqadigan ro'yxat — set_my_commands
# orqali o'rnatiladi. Oddiy foydalanuvchilarga faqat umumiy buyruqlar, admin
# bo'lganlarga esa (har bir admin uchun alohida, chat-scope orqali) qo'shimcha
# boshqaruv buyruqlari ham ko'rsatiladi.
DEFAULT_BOT_COMMANDS = [
    BotCommand(command="start", description="🏠 Botni ishga tushirish / bosh menyu"),
    BotCommand(command="help", description="ℹ️ Yordam va foydalanish qo'llanmasi"),
]

ADMIN_EXTRA_BOT_COMMANDS = [
    BotCommand(command="addadmin", description="👮 Yangi admin tayinlash (ID, @username yoki telefon)"),
    BotCommand(command="deladmin", description="🚫 Adminlik huquqini bekor qilish"),
    BotCommand(command="deltest", description="🗑 Test kodi bo'yicha testni o'chirish (tasdiqlash bilan)"),
]


async def _sync_admin_commands(bot: Bot):
    """Har bir adminning shaxsiy chatida qo'shimcha buyruqlar ko'rinishini yangilaydi."""
    try:
        rows = await POOL.fetch("SELECT telegram_id FROM users WHERE role='admin'")
        admin_ids = {ADMIN_ID} | {r["telegram_id"] for r in rows}
        for tg_id in admin_ids:
            try:
                await bot.set_my_commands(
                    DEFAULT_BOT_COMMANDS + ADMIN_EXTRA_BOT_COMMANDS,
                    scope=BotCommandScopeChat(chat_id=tg_id)
                )
            except Exception as e:
                logging.warning(f"Admin buyruqlarini o'rnatib bo'lmadi (tg_id={tg_id}): {e}")
    except Exception:
        logging.exception("_sync_admin_commands xatolik")


@router.message(Command("help"))
async def help_command(message: Message):
    await help_menu(message)


# ===================== ADMIN: ADMIN VA TESTLARNI BOSHQARISH =====================
def _digits_only(s: str) -> str:
    return re.sub(r"\D", "", s or "")


async def _find_user_by_phone(phone_input: str):
    tail = _digits_only(phone_input)[-9:]
    if len(tail) < 7:
        return None
    return await POOL.fetchrow("""
        SELECT telegram_id, full_name FROM users
        WHERE regexp_replace(COALESCE(phone,''), '\\D', '', 'g') LIKE '%' || $1
        LIMIT 1
    """, tail)


async def resolve_user_identifier(bot: Bot, text: str):
    """
    Foydalanuvchini Telegram ID, @username yoki telefon raqami (+ bilan) orqali topadi.
    Qaytaradi: (telegram_id, ism_yoki_None) yoki topilmasa None.

    Telefon raqam faqat bizning bazamizda (avval botga /start bosgan
    foydalanuvchilar orasida) qidiriladi — oddiy botlar Telegram'ning o'zidan
    "bu raqam kimniki" deb so'ray olmaydi (Bot API bunga imkon bermaydi).
    """
    t = (text or "").strip()
    if not t:
        return None

    if t.startswith("+"):
        row = await _find_user_by_phone(t)
        return (int(row["telegram_id"]), row["full_name"]) if row else None

    if _looks_like_int(t):
        return int(t), None

    username = _normalize_channel_ref(t)
    try:
        chat = await bot.get_chat(f"@{username}")
        name = " ".join(filter(None, [getattr(chat, "first_name", None), getattr(chat, "last_name", None)]))
        return int(chat.id), (name or username)
    except Exception:
        return None


@router.message(Command("addadmin"))
async def add_admin_command(message: Message, bot: Bot):
    if message.from_user.id != ADMIN_ID:
        return await message.answer("Bunga faqat asosiy admin huquqli!")

    parts = message.text.split(maxsplit=1)
    if len(parts) < 2:
        return await message.answer(
            "Format: <code>/addadmin ID | @username | +998901234567</code>\n"
            "Misol: <code>/addadmin 123456789</code>",
            parse_mode=ParseMode.HTML)

    resolved = await resolve_user_identifier(bot, parts[1])
    if not resolved:
        return await message.answer(
            "❌ Foydalanuvchi topilmadi. ID, @username yoki + bilan telefon raqam kiriting.")

    target_id, name = resolved
    await POOL.execute("""
        INSERT INTO users (telegram_id, role) VALUES ($1, 'admin')
        ON CONFLICT (telegram_id) DO UPDATE SET role='admin'
    """, target_id)
    await _sync_admin_commands(bot)

    who = f"{_h(name)} (<code>{target_id}</code>)" if name else f"<code>{target_id}</code>"
    await message.answer(f"✅ {who} admin qilib tayinlandi!", parse_mode=ParseMode.HTML)


@router.message(Command("deladmin"))
async def del_admin_command(message: Message, bot: Bot):
    if message.from_user.id != ADMIN_ID:
        return await message.answer("Bunga faqat asosiy admin huquqli!")

    parts = message.text.split(maxsplit=1)
    if len(parts) < 2:
        return await message.answer(
            "Format: <code>/deladmin ID | @username | +998901234567</code>", parse_mode=ParseMode.HTML)

    resolved = await resolve_user_identifier(bot, parts[1])
    if not resolved:
        return await message.answer(
            "❌ Foydalanuvchi topilmadi. ID, @username yoki + bilan telefon raqam kiriting.")

    target_id, name = resolved
    await POOL.execute("UPDATE users SET role='student' WHERE telegram_id=$1", target_id)
    await _sync_admin_commands(bot)

    who = f"{_h(name)} (<code>{target_id}</code>)" if name else f"<code>{target_id}</code>"
    await message.answer(f"✅ {who} dan adminlik olib tashlandi!", parse_mode=ParseMode.HTML)


@router.message(Command("deltest"))
async def del_test_command(message: Message):
    if not await is_admin_user(message.from_user.id):
        return

    parts = message.text.split()
    if len(parts) < 2:
        return await message.answer("Format: <code>/deltest TEST_KODI</code>\nMisol: /deltest DTM24A",
                                    parse_mode=ParseMode.HTML)

    code = normalize_code(parts[1])
    t = await POOL.fetchrow("SELECT id FROM tests WHERE code=$1", code)
    if not t:
        return await message.answer(f"❌ <code>{code}</code> kodli test topilmadi.", parse_mode=ParseMode.HTML)

    # ✅ Endi darhol o'chirmaydi — avval ogohlantirish va tasdiqlash ko'rsatiladi
    # (hozir kimdir yechayotgan bo'lsa yoki avval topshirganlar bo'lsa alohida
    # ogohlantiradi), xuddi tugma orqali o'chirishdagi kabi.
    await _send_delete_confirm(message, t["id"])


# ===================== ADMIN: ADMIN VA TESTLARNI BOSHQARISH (TUGMALAR) =====================

async def _render_admdel_list(answerable, page: int = 0):
    tests, has_more = await _fetch_tests_page(page)
    if not tests and page == 0:
        return await answerable.answer("Testlar yo'q.")

    kb = InlineKeyboardMarkup(inline_keyboard=[])
    for r in tests:
        kind = "🆓" if r["is_free"] else "🧾"
        kb.inline_keyboard.append(
            [InlineKeyboardButton(text=f"🗑 {kind} {r['code']} — {r['title']}", callback_data=f"admdel_ask_{r['id']}")])
    nav = _pagination_nav_row("del", page, has_more)
    if nav:
        kb.inline_keyboard.append(nav)

    await answerable.answer(
        f"🗑 <b>Testni o'chirish</b> ({page + 1}-sahifa)\n\nO'chirmoqchi bo'lgan testni tanlang "
        "(agar hozir kimdir yechayotgan bo'lsa yoki avval natija bo'lsa, tanlagach ogohlantiraman):",
        parse_mode=ParseMode.HTML, reply_markup=kb
    )


async def admin_ask_del_test(message: Message, state: FSMContext):
    if not await is_admin_user(message.from_user.id): return
    await _render_admdel_list(message, 0)


async def admin_ask_manage_admin(message: Message, state: FSMContext):
    if message.from_user.id != ADMIN_ID:
        return await message.answer("Bunga faqat asosiy admin huquqli!")

    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="👀 Adminlar ro'yxati", callback_data="admin_list")],
        [InlineKeyboardButton(text="➕ Admin qo'shish", callback_data="admin_add")],
        [InlineKeyboardButton(text="➖ Adminni o'chirish", callback_data="admin_del")]
    ])
    await message.answer("👮 <b>Adminlarni boshqarish</b>\n\nQuyidagilardan birini tanlang:", parse_mode=ParseMode.HTML,
                         reply_markup=kb)


@router.callback_query(F.data == "admin_list")
async def show_admin_list(callback: CallbackQuery):
    if callback.from_user.id != ADMIN_ID: return

    admins = await POOL.fetch("SELECT telegram_id, full_name, phone FROM users WHERE role='admin'")

    text = "👮 <b>Adminlar ro'yxati:</b>\n\n"
    text += f"👑 <b>Asosiy Admin (Siz):</b> <code>{ADMIN_ID}</code>\n\n"

    if not admins:
        text += "<i>Boshqa adminlar yo'q.</i>"
    else:
        for r in admins:
            tg_id, fname, phone = r["telegram_id"], r["full_name"], r["phone"]
            name_str = fname if fname else "Noma'lum"
            phone_str = phone if phone else "Noma'lum"
            text += f"👤 Ism: <b>{name_str}</b>\n🆔 ID: <code>{tg_id}</code>\n📞 Tel: {phone_str}\n\n"

    await callback.message.answer(text, parse_mode=ParseMode.HTML)
    await callback.answer()


@router.callback_query(F.data.in_({"admin_add", "admin_del"}))
async def admin_manage_cb(callback: CallbackQuery, state: FSMContext):
    if callback.from_user.id != ADMIN_ID: return
    action = callback.data
    await state.update_data(admin_action=action)

    act_text = "qo'shish" if action == "admin_add" else "o'chirish"
    await callback.message.answer(
        f"Adminni {act_text} uchun foydalanuvchini kiriting:\n\n"
        f"• Telegram ID: <code>123456789</code>\n"
        f"• Username: <code>@username</code>\n"
        f"• Telefon (+ bilan): <code>+998901234567</code> "
        f"<i>(faqat botga avval /start bosgan bo'lsa topiladi)</i>",
        parse_mode=ParseMode.HTML,
        reply_markup=nav_kb(include_back=True)
    )
    await state.set_state(AdminManageState.waiting_admin_id)


@router.message(AdminManageState.waiting_admin_id)
async def process_manage_admin(message: Message, state: FSMContext, bot: Bot):
    if (message.text or "") in TOP_MENU_BTNS: return

    resolved = await resolve_user_identifier(bot, message.text)
    if not resolved:
        return await message.answer(
            "❌ Foydalanuvchi topilmadi. ID, @username yoki + bilan telefon raqam kiriting:")

    target_id, resolved_name = resolved
    data = await state.get_data()
    action = data.get("admin_action")
    who = f"{_h(resolved_name)} (<code>{target_id}</code>)" if resolved_name else f"<code>{target_id}</code>"

    if action == "admin_add":
        await POOL.execute("""
            INSERT INTO users (telegram_id, role) VALUES ($1, 'admin')
            ON CONFLICT (telegram_id) DO UPDATE SET role='admin'
        """, target_id)
        text = f"✅ {who} admin qilib tayinlandi!"
    else:
        await POOL.execute("UPDATE users SET role='student' WHERE telegram_id=$1", target_id)
        text = f"✅ {who} dan adminlik olib tashlandi!"

    await _sync_admin_commands(bot)
    await message.answer(text, parse_mode=ParseMode.HTML, reply_markup=main_menu_kb(True))
    await state.clear()


# ===================== HELP =====================
@router.message(F.text == BTN_HELP)
async def help_menu(message: Message):
    is_admin = await is_admin_user(message.from_user.id)

    help_text = (
        f"📚 <b>{ACADEMY_NAME} — botdan foydalanish qo'llanmasi</b>\n\n"

        "👨‍🎓 <b>1. Test ishlash</b>\n"
        "• <b>🆓 Tekin testlar</b> — barcha uchun ochiq.\n"
        "• <b>🧾 Pullik testlar</b> — to'lov tasdiqlangach ochiladi.\n"
        "<i>Testni tezroq topish uchun «🔎 Kod orqali topish» tugmasidan foydalaning — u yerda "
        "testni necha kishi yechgani va o'rtacha natija ham ko'rsatiladi.</i>\n\n"

        "<b>2. Test qoidalari</b>\n"
        "• Har bir testga faqat <b>1 marta</b> urinish beriladi — ehtiyot bo'lib javob bering.\n"
        "• Ochiq (variantsiz) savolga «✍️ Matn bilan javob yozish» tugmasi orqali javob yozasiz — "
        "javobingiz ustoz kiritgan kalit bilan avtomatik tekshiriladi.\n"
        "• Javobni tugmalar (A–F) orqali belgilang, yoki hammasini birdaniga matn "
        "shaklida yuboring (masalan: <code>abcdabcd...</code>). Javoblar soni savollar soniga teng bo'lishi shart.\n"
        "• Aloqa uzilib qolsa yoki «🏠 Bosh menyu» bosib chiqib ketsangiz ham xavotir olmang — "
        "shu testni qayta ochsangiz, vaqt hisoblagichi boshidan boshlanmay, qolgan vaqtdan davom etadi "
        "va belgilagan javoblaringiz tiklanadi.\n"
        "• Vaqt tugasa, shu paytgacha belgilagan javoblaringiz <b>avtomatik topshiriladi</b>.\n"
        "• 50% va undan yuqori natija bilan avtomatik <b>Sertifikat (PDF)</b> beriladi 🎁.\n\n"

        "<b>3. Natijadan keyin nima qilsa bo'ladi</b>\n"
        "• <b>🔎 Xatolarni ko'rish</b> — qaysi savolda xato qilganingizni istalgan vaqt qayta ko'rish.\n"
        "• <b>🏅 Reytingdagi o'rningiz</b> — shu testni yechganlar orasida qaysi o'rindasiz.\n"
        "• <b>📤 Natijani ulashish</b> — do'stlaringizga yuborish uchun tayyor matn.\n"
        "• <b>📜 Sertifikatlarim</b> («📄 Natijalarim» ichida) — oldin olgan sertifikatlaringizni qayta yuklab olish.\n\n"

        "<b>4. Pullik test uchun to'lov</b>\n"
        "1️⃣ Testni tanlab «Sotib olish»ni bosing.\n"
        "2️⃣ Ko'rsatilgan karta raqamiga to'lov qiling.\n"
        "3️⃣ To'lov chekini (skrinshot) botga yuboring.\n"
        "4️⃣ Admin tasdiqlagach, test <b>«📦 Pullik testlarim»</b> bo'limida ochiladi.\n\n"

        "<b>5. O'zingiz test yaratish</b> (barcha ro'yxatdan o'tgan foydalanuvchilar uchun)\n"
        "• <b>➕ Test yaratish</b> — o'zingiz test tuzib ulashishingiz mumkin, kuniga 10 tagacha "
        "(tekin yoki pullik — pullik bo'lsa to'lov admin kartasiga tushadi, u sizga tegishli "
        "ulushni o'zi o'tkazib beradi).\n"
        "• <b>🛠 Yaratgan testlarim</b> — har bir testingiz uchun bitta joyda: statistikasi, "
        "<b>✏️ tahrirlash</b>, <b>🛑 vaqtincha yakunlash/qayta ochish</b>, <b>🎓 sertifikat ko'rinishini "
        "sozlash</b> (o'quv markazi nomi, mas'ul shaxs, rangi) va kerak bo'lsa <b>🗑 o'chirish</b>.\n\n"

        "<b>6. Reyting va profil</b>\n"
        "• <b>📄 Natijalarim</b> — shaxsiy statistikangiz va topshirgan testlaringiz tarixi.\n"
        "• <b>🏆 Reyting</b> — umumiy va test bo'yicha eng yaxshi natijalar.\n"
        "• <b>👤 Profilim</b> — ism-familiya va telefon raqamini o'zgartirish.\n\n"

        "🔁 <b>Navigatsiya:</b> istalgan vaqtda «🏠 Bosh menyu» tugmasidan foydalaning; test yaratishda "
        "«🔄 Boshidan boshlash» bitta qadam orqaga qaytaradi.\n\n"

        "📞 <b>Aloqa:</b> pastdagi «✉️ Admin bilan bog'lanish» tugmasi orqali to'g'ridan-to'g'ri yozishingiz "
        "mumkin, yoki:\n"
        f"👤 {PAYMENT_OWNER}\n"
        f"📞 {PAYMENT_PHONE}"
    )

    if is_admin:
        help_text += (
            "\n\n➖➖➖➖➖➖➖➖\n"
            "👮 <b>Admin uchun qo'shimcha bo'limlar:</b>\n"
            "• <b>🟡 Pending to'lovlar</b> — tasdiq kutayotgan to'lovlarni ko'rish/tasdiqlash.\n"
            "• <b>📊 Statistika</b> — bitta panelda: vaqt oralig'ida natijalar, hozir faol "
            "foydalanuvchilar, foydalanuvchilar ro'yxati/qidiruv, <b>📢 barchaga xabar (reklama) yuborish</b> "
            "(matn, rasm, video, fayl — nima yuborsangiz shuni), moliyaviy hisobot, testni tahrirlash, "
            "vaqtincha yashirish/qaytarish, testni yakunlash/qayta ochish, bugungi hisobot, testlar reytingi "
            "va <b>💾 zaxira nusxa olish</b> (bazaning to'liq holatini Telegram orqali yuklab olish).\n"
            "• <b>📣 Majburiy kanallar</b> — obuna majburiy kanal/guruhlarni boshqarish.\n"
            "• <b>🗑 Testni o'chirish</b> — ro'yxatdan tanlab, ogohlantirish bilan o'chirish "
            "(hozir kimdir yechayotgan bo'lsa alohida ogohlantiriladi).\n"
            "• <b>👮 Adminlarni boshqarish</b> — ID, @username yoki telefon raqami orqali "
            "admin tayinlash/olib tashlash.\n\n"
            "💾 <i>Ma'lumotlar bazasi zaxirasi har "
            f"{DB_BACKUP_INTERVAL_HOURS} soatda avtomatik shu chatga yuboriladi.</i>"
        )

    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="✉️ Admin bilan bog'lanish", callback_data="ask_admin_start")]
    ])
    await message.answer(help_text, parse_mode=ParseMode.HTML, reply_markup=kb)


@router.callback_query(F.data == "ask_admin_start")
async def ask_admin_start(callback: CallbackQuery, state: FSMContext):
    await callback.message.answer(
        "✉️ <b>Adminga xabar</b>\n\nSavolingiz yoki taklifingizni yozing — to'g'ridan-to'g'ri adminga yetkazaman:",
        parse_mode=ParseMode.HTML,
        reply_markup=nav_kb(include_back=True)
    )
    await state.set_state(AskAdminState.waiting_text)


@router.message(AskAdminState.waiting_text)
async def ask_admin_send(message: Message, state: FSMContext, bot: Bot):
    if (message.text or "") in TOP_MENU_BTNS:
        return
    text = (message.text or "").strip()
    if len(text) < 2:
        return await message.answer("⚠️ Xabar juda qisqa. Qayta yozing:")

    user_row = await POOL.fetchrow("SELECT full_name, phone FROM users WHERE telegram_id=$1", message.from_user.id)
    full_name = (user_row["full_name"] if user_row and user_row["full_name"] else message.from_user.full_name)
    phone = user_row["phone"] if user_row else None

    # ✅ Admin foydalanuvchi bilan bevosita bog'lana olishi uchun: Telegram
    # username bo'lsa @nickname, bo'lmasa ro'yxatdan o'tishda qoldirgan
    # telefon raqami ko'rsatiladi.
    if message.from_user.username:
        contact_line = f"🔗 Username: @{message.from_user.username}"
    elif phone:
        contact_line = f"📞 Telefon: {phone}"
    else:
        contact_line = "📞 Telefon/username: noma'lum"

    try:
        await bot.send_message(
            ADMIN_ID,
            f"✉️ <b>Yangi xabar</b>\n\n"
            f"👤 {_h(full_name)}\n"
            f"{contact_line}\n"
            f"🆔 <code>{message.from_user.id}</code>\n\n"
            f"{message.html_text}",
            parse_mode=ParseMode.HTML
        )
        admin_flag = await is_admin_user(message.from_user.id)
        await message.answer("✅ Xabaringiz adminga yuborildi. Tez orada javob berishadi.",
                             reply_markup=main_menu_kb(admin_flag))
    except Exception:
        logging.exception("ask_admin_send: xabarni adminga yuborib bo'lmadi")
        await message.answer("⚠️ Xabarni yuborishda xatolik yuz berdi. Birozdan so'ng qayta urinib ko'ring.")

    await state.clear()


CERT_CHECK_EVERY_SEC = 30


async def certificate_daemon(bot: Bot):
    """Kechiktirilgan sertifikatlarni yuboradi (Rasch — test yakunlangach,
    ochiq javoblar tekshirilgach, yoki darhol yuborishda xato bo'lgan hollar).

    ✅ Avval har safar faqat birinchi 20 ta "yuborilmagan" natija olinar edi.
    Sertifikatga loyiq bo'lmagan (masalan 50% dan past) natijalar ham
    certificate_sent=0 bo'lib qolaverardi — ular 20 tadan oshgach, navbatni
    abadiy to'sib qo'yib, keyingi LOYIQ o'quvchilarga sertifikat umuman
    yuborilmay qolardi. Endi: loyiq emasligi aniq bo'lgan natijalar
    certificate_sent=2 deb belgilanadi va har aylanishda butun ro'yxat id
    bo'yicha bo'laklab ko'rib chiqiladi."""
    while True:
        try:
            last_id = 0
            while True:
                rows = await POOL.fetch("""
                    SELECT r.id, r.percent, r.rasch_percent, r.grade, COALESCE(r.open_pending,0) AS open_pending,
                           r.test_id, t.exam_type, COALESCE(t.rush_finalized,0) AS rush_finalized
                    FROM results r
                    JOIN tests t ON t.id=r.test_id
                    WHERE (r.certificate_sent IS NULL OR r.certificate_sent=0) AND r.id > $1
                    ORDER BY r.id ASC
                    LIMIT 50
                """, last_id)
                if not rows:
                    break
                for row in rows:
                    last_id = row["id"]
                    et = row["exam_type"] or "simple"
                    finalized = int(row["rush_finalized"]) == 1
                    if et == "rasch" and not finalized:
                        if await finalize_rush_for_test_if_ready(row["test_id"], bot=bot):
                            finalized = True
                            fresh = await POOL.fetchrow(
                                "SELECT rasch_percent, grade, COALESCE(open_pending,0) AS op FROM results WHERE id=$1",
                                row["id"])
                            if not fresh:
                                continue
                            decision = _cert_decision(et, fresh["rasch_percent"], fresh["grade"], True, fresh["op"])
                        else:
                            continue
                    else:
                        decision = _cert_decision(et, row["percent"], row["grade"], finalized, row["open_pending"])
                    if decision is None:
                        continue
                    if decision is False:
                        await POOL.execute("UPDATE results SET certificate_sent=2 WHERE id=$1", row["id"])
                        continue
                    try:
                        await _send_result_certificate(bot, row["id"])
                    except TelegramForbiddenError:
                        # o'quvchi botni bloklagan — har 30 soniyada qayta urinib, logni to'ldirmaymiz
                        await POOL.execute("UPDATE results SET certificate_sent=3 WHERE id=$1", row["id"])
                    except Exception:
                        logging.exception(f"certificate_daemon: sertifikat yuborilmadi (result={row['id']})")
        except Exception:
            logging.exception("certificate_daemon iteration failed")

        await asyncio.sleep(CERT_CHECK_EVERY_SEC)


# ===================== MA'LUMOTLAR BAZASI ZAXIRA NUSXASI (BACKUP) =====================
BACKUP_TABLES = [
    "users", "tests", "payments", "results", "sessions", "required_channels", "user_limits",
    "staged_attempts", "retake_grants",
]
# Eski zaxira fayllarida bo'lmagan jadvallar (ular bo'lmasa ham tiklash ishlaydi)
BACKUP_OPTIONAL_TABLES = {"retake_grants"}


async def _run_db_backup(bot: Bot, requester_tg_id: int | None = None):
    """
    Barcha jadvallarni JSON qilib ADMIN_ID'ga (yoki so'ragan adminga) hujjat
    sifatida yuboradi. pg_dump kabi tashqi dasturga bog'liq emas — faqat
    asyncpg orqali, shuning uchun har qanday hosting muhitida ishlaydi.
    """
    dump = {"created_at": now_str_local(), "tables": {}}
    for table in BACKUP_TABLES:
        rows = await POOL.fetch(f"SELECT * FROM {table}")
        dump["tables"][table] = [dict(r) for r in rows]

    tmp_dir = tempfile.gettempdir()
    fname = f"backup_{datetime.now(TZ).strftime('%Y%m%d_%H%M%S')}.json"
    path = os.path.join(tmp_dir, fname)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(dump, f, ensure_ascii=False, default=str)

    total_rows = sum(len(v) for v in dump["tables"].values())
    caption = (
        f"💾 <b>Zaxira nusxa</b>\n"
        f"🗓 {dump['created_at']}\n"
        f"📊 Jami: {total_rows} qator, {len(BACKUP_TABLES)} jadval"
    )

    targets = {requester_tg_id} if requester_tg_id else set()
    targets.add(ADMIN_ID)
    for tg_id in targets:
        try:
            await bot.send_document(chat_id=tg_id, document=FSInputFile(path), caption=caption,
                                    parse_mode=ParseMode.HTML)
        except Exception:
            logging.exception(f"Zaxira nusxani yuborib bo'lmadi (tg_id={tg_id})")

    try:
        os.remove(path)
    except Exception:
        pass


async def db_backup_daemon(bot: Bot):
    if DB_BACKUP_INTERVAL_HOURS <= 0:
        return
    while True:
        await asyncio.sleep(DB_BACKUP_INTERVAL_HOURS * 3600)
        try:
            await _run_db_backup(bot)
        except Exception:
            logging.exception("db_backup_daemon iteration failed")


@router.callback_query(F.data == "backup_now")
async def backup_now(callback: CallbackQuery, bot: Bot):
    if not await is_admin_user(callback.from_user.id):
        return
    await callback.message.answer("⏳ Zaxira nusxa tayyorlanmoqda...")
    try:
        await _run_db_backup(bot, requester_tg_id=callback.from_user.id)
    except Exception:
        logging.exception("backup_now failed")
        await callback.message.answer("⚠️ Zaxira nusxa olishda xatolik yuz berdi.")


# ===================== ZAXIRADAN TIKLASH (RESTORE) =====================
# ✅ _run_db_backup bilan juftlashadi: falokat holatida admin oldin olingan
# .json faylni yuborib, bazani o'sha holatga qaytarishi mumkin. Juda xavfli
# amal bo'lgani uchun har doim aniq tasdiqlash bosqichidan o'tadi.
_TABLES_WITH_SERIAL_ID = [
    "users", "tests", "payments", "results", "sessions", "required_channels", "staged_attempts",
]


async def _run_db_restore(tables: dict):
    async with POOL.acquire() as db:
        async with db.transaction():
            # 1) hammasini bo'shatamiz (bog'liqlikka teskari tartibda)
            for t in reversed(BACKUP_TABLES):
                await db.execute(f"TRUNCATE TABLE {t} RESTART IDENTITY CASCADE")

            # 2) zaxiradagi qatorlarni bog'liqlik tartibida qayta yozamiz
            for t in BACKUP_TABLES:
                rows = tables.get(t) or []
                if not rows:
                    continue
                # ✅ Faqat hozirgi sxemada MAVJUD ustunlar yoziladi — eski zaxirada
                # keyinchalik olib tashlangan ustun bo'lsa, tiklash butunlay yiqilardi.
                existing_cols = {r["column_name"] for r in await db.fetch(
                    "SELECT column_name FROM information_schema.columns "
                    "WHERE table_schema=current_schema() AND table_name=$1", t)}
                cols = [c for c in rows[0].keys() if c in existing_cols]
                col_list = ", ".join(cols)
                placeholders = ", ".join(f"${i + 1}" for i in range(len(cols)))
                insert_sql = f"INSERT INTO {t} ({col_list}) VALUES ({placeholders})"
                for row in rows:
                    await db.execute(insert_sql, *[row.get(c) for c in cols])

            # 3) SERIAL id ketma-ketligini eng katta id'dan davom etadigan qilib qo'yamiz —
            # aks holda keyingi oddiy INSERT eski (band) id'ga urilib xato beradi.
            for t in _TABLES_WITH_SERIAL_ID:
                await db.execute(f"""
                    SELECT setval(pg_get_serial_sequence('{t}', 'id'), COALESCE((SELECT MAX(id) FROM {t}), 1))
                """)


@router.callback_query(F.data == "restore_start")
async def restore_start(callback: CallbackQuery, state: FSMContext):
    if not await is_admin_user(callback.from_user.id):
        return
    await callback.message.answer(
        "📥 <b>Zaxiradan tiklash</b>\n\n"
        "🔴 <b>DIQQAT — bu og'ir amal:</b> tasdiqlasangiz, hozirgi bazadagi <b>barcha</b> "
        "foydalanuvchilar, testlar, natijalar va to'lovlar zaxira faylidagi holat bilan "
        "<b>butunlay almashtiriladi</b>. Zaxira olingandan keyin qo'shilgan har qanday yangi "
        "ma'lumot yo'qoladi. Bu amalni ortga qaytarib bo'lmaydi.\n\n"
        "Faqat haqiqiy falokat holatida ishlating. Davom etish uchun «💾 Zaxira nusxa olish» "
        "orqali olingan <code>.json</code> faylni shu yerga yuboring:",
        parse_mode=ParseMode.HTML
    )
    await state.set_state(RestoreState.waiting_file)


@router.message(RestoreState.waiting_file, F.document)
async def restore_got_file(message: Message, state: FSMContext, bot: Bot):
    if not await is_admin_user(message.from_user.id):
        await state.clear()
        return

    fname = message.document.file_name or ""
    if not fname.lower().endswith(".json"):
        return await message.answer("⚠️ Bu .json fayl emasga o'xshaydi. To'g'ri zaxira faylini yuboring:")

    tg_file = await bot.get_file(message.document.file_id)
    tmp_path = os.path.join(tempfile.gettempdir(), f"restore_{now_ts()}.json")
    await bot.download_file(tg_file.file_path, destination=tmp_path)

    try:
        with open(tmp_path, encoding="utf-8") as f:
            payload = json.load(f)
        tables = payload.get("tables")
        if not tables or not all(t in tables for t in BACKUP_TABLES if t not in BACKUP_OPTIONAL_TABLES):
            raise ValueError("noto'liq fayl")
    except Exception:
        try:
            os.remove(tmp_path)
        except Exception:
            pass
        return await message.answer(
            "⚠️ Faylni o'qib bo'lmadi yoki format noto'g'ri. Bu botning o'zi yuborgan zaxira "
            "faylimi, tekshirib qayta yuboring:")

    await state.update_data(restore_path=tmp_path)
    counts_text = "\n".join(f"• {t}: {len(rows)} qator" for t, rows in tables.items())

    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="🔴 Ha, HAMMASINI shu holat bilan almashtir", callback_data="restore_confirm")],
        [InlineKeyboardButton(text="❌ Bekor qilish", callback_data="restore_cancel")],
    ])
    await message.answer(
        f"📦 <b>Zaxira topildi</b> ({payload.get('created_at', 'noma’lum vaqt')}):\n{counts_text}\n\n"
        f"❗️Tasdiqlasangiz, hozirgi bazadagi barcha ma'lumotlar o'chib, shu zaxira bilan almashtiriladi. "
        f"Aniq ishonchingiz komilmi?",
        parse_mode=ParseMode.HTML, reply_markup=kb
    )


@router.callback_query(RestoreState.waiting_file, F.data == "restore_cancel")
async def restore_cancel(callback: CallbackQuery, state: FSMContext):
    data = await state.get_data()
    path = data.get("restore_path")
    if path and os.path.exists(path):
        try:
            os.remove(path)
        except Exception:
            pass
    await state.clear()
    await callback.message.answer("Bekor qilindi.")


@router.callback_query(RestoreState.waiting_file, F.data == "restore_confirm")
async def restore_confirm(callback: CallbackQuery, state: FSMContext):
    if not await is_admin_user(callback.from_user.id):
        return
    data = await state.get_data()
    path = data.get("restore_path")
    await state.clear()

    if not path or not os.path.exists(path):
        return await callback.message.answer("⚠️ Fayl topilmadi, qaytadan boshlang.")

    await callback.message.answer("⏳ Tiklanmoqda, biroz kuting...")
    try:
        with open(path, encoding="utf-8") as f:
            payload = json.load(f)
        await _run_db_restore(payload["tables"])
        await callback.message.answer("✅ Ma'lumotlar bazasi zaxiradan muvaffaqiyatli tiklandi.")
    except Exception:
        logging.exception("restore_confirm failed")
        await callback.message.answer(
            "⚠️ Tiklashda xatolik yuz berdi. Amal bekor qilindi (tranzaksiya ichida bo'lgani uchun "
            "baza o'zgarmagan bo'lishi kerak) — server logini tekshiring.")
    finally:
        try:
            os.remove(path)
        except Exception:
            pass


# ===================== TESTNI YAKUNLASH / QAYTA OCHISH =====================
async def _is_admin_or_creator(user_tg_id: int, test_id: int) -> bool:
    if await is_admin_user(user_tg_id):
        return True
    row = await POOL.fetchrow("SELECT creator_id FROM tests WHERE id=$1", test_id)
    return bool(row and row["creator_id"] == user_tg_id)


async def _render_closetest_list(answerable, page: int = 0):
    tests, has_more = await _fetch_tests_page(page)
    if not tests and page == 0:
        return await answerable.answer("Testlar yo'q.")

    kb = InlineKeyboardMarkup(inline_keyboard=[])
    for r in tests:
        mark = "🛑 Yakunlangan" if r["is_closed"] else "🟢 Faol"
        kb.inline_keyboard.append([InlineKeyboardButton(
            text=f"{mark} — {r['code']}", callback_data=f"closetest_toggle_{r['id']}")])
    nav = _pagination_nav_row("close", page, has_more)
    if nav:
        kb.inline_keyboard.append(nav)

    await answerable.answer(
        f"🛑 <b>Testlarni yakunlash / qayta ochish</b> ({page + 1}-sahifa)\n\n"
        "Yakunlangan testga yangi urinish qabul qilinmaydi (Rasch bo'lsa, natijalar darhol "
        "hisoblab chiqiladi). Qayta ochsangiz, oldin yechganlarga qayta yechish taklif qilinadi.\n\n"
        "Bosilganda holati almashadi:",
        parse_mode=ParseMode.HTML, reply_markup=kb
    )


@router.callback_query(F.data == "closetest_menu")
async def closetest_menu(callback: CallbackQuery):
    if not await is_admin_user(callback.from_user.id):
        return
    await _render_closetest_list(callback.message, 0)


def _retake_offer_kb(test_id: int) -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="✅ Ha, taklif qilinsin", callback_data=f"retakeoffer_{test_id}")],
        [InlineKeyboardButton(text="❌ Yo'q, kerak emas", callback_data="dismiss_msg")],
    ])


@router.callback_query(F.data.startswith("closetest_toggle_"))
async def closetest_toggle(callback: CallbackQuery, state: FSMContext, bot: Bot):
    test_id = int(callback.data.split("_")[2])
    if not await _is_admin_or_creator(callback.from_user.id, test_id):
        return await callback.answer("Bu testga huquqingiz yo'q!", show_alert=True)

    row = await POOL.fetchrow(
        "SELECT is_closed, code, exam_type, start_mode, start_ts, duration FROM tests WHERE id=$1", test_id)
    if not row:
        return await callback.answer("Topilmadi", show_alert=True)

    new_val = 0 if int(row["is_closed"] or 0) == 1 else 1
    await POOL.execute("UPDATE tests SET is_closed=$1 WHERE id=$2", new_val, test_id)
    et = row["exam_type"] or "simple"

    if new_val == 1:
        active_now = await POOL.fetchval(
            "SELECT COUNT(*) FROM sessions WHERE test_id=$1 AND start_ts + duration_sec > $2", test_id, now_ts())
        if active_now:
            await callback.message.answer(
                f"ℹ️ Hozir <b>{active_now}</b> kishi bu testni yechayapti — ular boshlagan urinishini yakunlay "
                f"oladi (yangi urinish qabul qilinmaydi)."
                + (" Ular topshirgach Rasch natijalari avtomatik qayta hisoblanadi." if et == "rasch" else ""),
                parse_mode=ParseMode.HTML)
        if et == "rasch":
            done = await finalize_rush_for_test_if_ready(test_id, bot=bot)
            if not done:
                pending = await POOL.fetchval(
                    "SELECT COUNT(*) FROM results WHERE test_id=$1 AND COALESCE(open_pending,0)>0", test_id)
                if pending:
                    await callback.message.answer(
                        f"🛑 Test yakunlandi. Rasch natijalari <b>{pending}</b> ta o'quvchining rasm "
                        f"ko'rinishidagi ochiq javoblari tekshirilgach e'lon qilinadi.",
                        parse_mode=ParseMode.HTML,
                        reply_markup=InlineKeyboardMarkup(inline_keyboard=[
                            [InlineKeyboardButton(text="✍️ Javoblarni tekshirish", callback_data=f"ogl_{test_id}")],
                            [InlineKeyboardButton(text="⚡ Tekshirmasdan hisoblash",
                                                  callback_data=f"ogforce_{test_id}")],
                        ]))
        await callback.answer(f"🛑 {row['code']} yakunlandi.", show_alert=True)
        return

    # ---- qayta ochish ----
    if et == "rasch":
        # ✅ Qayta ochilgan Rasch testda yangi (qayta yechgan) natijalar ham
        # guruh hisobiga kirishi uchun hisob qaytadan bajariladi. Avval
        # rush_finalized=1 qolib, keyingi natijalar hech qachon Rasch ball
        # olmasdi.
        await POOL.execute("UPDATE tests SET rush_finalized=0, open_grade_reminded_ts=0 WHERE id=$1", test_id)
        await POOL.execute(
            "UPDATE results SET rasch_ball=NULL, rasch_percent=NULL, grade=NULL WHERE test_id=$1", test_id)
    await callback.answer(f"🟢 {row['code']} qayta ochildi.", show_alert=True)

    expired, _ = is_scheduled_expired(row["start_mode"], row["start_ts"], row["duration"])
    if expired:
        # ✅ Rejalashtirilgan testning vaqti allaqachon o'tgan — yangi vaqtsiz
        # qayta ochish ma'nosiz (hech kim boshlay olmaydi). Yangi vaqt so'raymiz.
        await state.set_state(ReopenScheduleState.waiting_datetime)
        await state.update_data(reopen_test_id=test_id)
        await callback.message.answer(
            "📅 Bu testning belgilangan vaqti allaqachon o'tib ketgan. Qayta yechish mumkin bo'lishi uchun "
            "<b>yangi boshlanish sana va vaqtini</b> kiriting.\nFormat: <code>YYYY-MM-DD HH:MM</code>",
            parse_mode=ParseMode.HTML, reply_markup=nav_kb(include_back=False)
        )
        return

    await callback.message.answer(
        "Bu testni oldin yechgan foydalanuvchilarga qayta yechish imkoniyati taklif qilinsinmi?",
        reply_markup=_retake_offer_kb(test_id)
    )


@router.message(ReopenScheduleState.waiting_datetime)
async def reopen_set_datetime(message: Message, state: FSMContext):
    if (message.text or "") in TOP_MENU_BTNS:
        return
    data = await state.get_data()
    test_id = data.get("reopen_test_id")
    if not test_id or not await _is_admin_or_creator(message.from_user.id, test_id):
        await state.clear()
        return
    parts = (message.text or "").strip().split(maxsplit=1)
    res = to_start_ts(parts[0], parts[1]) if len(parts) == 2 else None
    if not res:
        return await message.answer("⚠️ Format noto'g'ri. Masalan: <code>2026-02-15 16:00</code>",
                                    parse_mode=ParseMode.HTML)
    start_ts_val, start_at = res
    duration = await POOL.fetchval("SELECT duration FROM tests WHERE id=$1", test_id)
    problem = schedule_window_problem(start_ts_val, duration)
    if problem:
        return await message.answer(problem)
    await POOL.execute(
        "UPDATE tests SET start_mode='scheduled', start_ts=$1, start_at=$2, is_closed=0, rush_finalized=0 WHERE id=$3",
        start_ts_val, start_at, test_id)
    await state.clear()
    admin_flag = await is_admin_user(message.from_user.id)
    await message.answer(f"✅ Yangi vaqt saqlandi: <b>{start_at}</b>", parse_mode=ParseMode.HTML,
                         reply_markup=main_menu_kb(admin_flag))
    await message.answer("Bu testni oldin yechgan foydalanuvchilarga qayta yechish imkoniyati taklif qilinsinmi?",
                         reply_markup=_retake_offer_kb(test_id))


@router.callback_query(F.data.startswith("retakeoffer_"))
async def retake_offer_send(callback: CallbackQuery, bot: Bot):
    test_id = int(callback.data.split("_")[1])
    if not await _is_admin_or_creator(callback.from_user.id, test_id):
        return await callback.answer("Bu testga huquqingiz yo'q!", show_alert=True)

    await callback.message.answer("🚀 Taklif yuborilmoqda...")
    asyncio.create_task(_run_retake_offer(bot, callback.from_user.id, test_id))


async def _run_retake_offer(bot: Bot, requester_tg_id: int, test_id: int):
    t = await POOL.fetchrow("SELECT code, title FROM tests WHERE id=$1", test_id)
    if not t:
        return
    rows = await POOL.fetch("""
        SELECT DISTINCT u.telegram_id FROM results r JOIN users u ON u.id=r.user_id WHERE r.test_id=$1
    """, test_id)

    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="🔄 Qayta yechish", callback_data=f"retake_{test_id}")]
    ])
    sent = 0
    for r in rows:
        # har bir o'quvchiga shu taklif uchun BITTA qayta yechish ruxsati
        await POOL.execute("""
            INSERT INTO retake_grants (user_id, test_id, granted_ts, used_ts)
            SELECT id, $2, $3, 0 FROM users WHERE telegram_id=$1
            ON CONFLICT (user_id, test_id) DO UPDATE SET granted_ts=EXCLUDED.granted_ts, used_ts=0
        """, r["telegram_id"], test_id, now_ts())
        try:
            await bot.send_message(
                r["telegram_id"],
                f"🔄 <b>{t['title']}</b> (<code>{t['code']}</code>) testi qayta ochildi!\n"
                f"Xohlasangiz qayta yechib, natijangizni yangilashingiz mumkin.",
                parse_mode=ParseMode.HTML, reply_markup=kb
            )
            sent += 1
        except Exception:
            pass
        await asyncio.sleep(0.05)

    try:
        await bot.send_message(requester_tg_id, f"✅ Qayta yechish taklifi {sent} kishiga yuborildi.")
    except Exception:
        pass


async def _has_retake_grant(user_db_id: int, test_id: int) -> bool:
    return bool(await POOL.fetchval(
        "SELECT 1 FROM retake_grants WHERE user_id=$1 AND test_id=$2 AND COALESCE(used_ts,0)=0",
        user_db_id, test_id))


@router.callback_query(F.data.startswith("retake_"))
async def retake_ask(callback: CallbackQuery):
    test_id = int(callback.data.split("_")[1])
    user_db_id = await get_user_db_id(callback.from_user.id)
    if not user_db_id:
        return await callback.message.answer("Avval /start orqali ro'yxatdan o'ting.")

    has_result = await POOL.fetchrow(
        "SELECT 1 FROM results WHERE user_id=$1 AND test_id=$2", user_db_id, test_id)
    if not has_result:
        return await callback.message.answer("Sizda bu test bo'yicha oldingi natija topilmadi.")
    if not await _has_retake_grant(user_db_id, test_id):
        return await callback.message.answer(
            "ℹ️ Bu taklif allaqachon ishlatilgan yoki muddati o'tgan. Har bir testga bitta urinish beriladi — "
            "qayta yechish faqat ustoz yangi taklif yuborsa mumkin.")

    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="✅ Ha, qaytadan boshlayman", callback_data=f"retakeok_{test_id}")],
        [InlineKeyboardButton(text="❌ Yo'q", callback_data="dismiss_msg")],
    ])
    await callback.message.answer(
        "⚠️ <b>Diqqat:</b> eski natijangiz butunlay o'chiriladi va testni <b>boshidan</b> boshlaysiz. "
        "Davom etasizmi?",
        parse_mode=ParseMode.HTML, reply_markup=kb
    )


@router.callback_query(F.data.startswith("retakeok_"))
async def retake_confirm(callback: CallbackQuery):
    test_id = int(callback.data.split("_")[1])
    user_db_id = await get_user_db_id(callback.from_user.id)
    if not user_db_id:
        return await callback.message.answer("Avval /start orqali ro'yxatdan o'ting.")

    t = await POOL.fetchrow(
        "SELECT is_free, is_hidden, is_closed, start_mode, start_ts, duration, exam_type FROM tests WHERE id=$1",
        test_id)
    if not t:
        return await callback.message.answer("Test topilmadi.")
    if (t["exam_type"] or "") not in STAGED_EXAM_TYPES and not await _has_retake_grant(user_db_id, test_id):
        return await callback.message.answer(
            "ℹ️ Qayta yechish uchun ruxsat yo'q (taklif ishlatilgan yoki yuborilmagan). "
            "Avvalgi natijangiz saqlanib qoldi.")
    if int(t["is_hidden"] or 0) == 1 or int(t["is_closed"] or 0) == 1:
        return await callback.message.answer("⏸ Bu test hozircha yopiq, birozdan so'ng urinib ko'ring.")

    # ✅ MUHIM: eski natija FAQAT qayta boshlash haqiqatan mumkin bo'lsa
    # o'chiriladi. Avval natija darhol o'chirilib, keyin test vaqti o'tgani
    # (rejalashtirilgan/Rasch test) yoki pullik testda eski boshlanish vaqti
    # qolib ketgani sabab "vaqti tugab bo'lgan" chiqib — o'quvchi natijasini
    # butunlay yo'qotardi.
    expired, _ = is_scheduled_expired(t["start_mode"], t["start_ts"], t["duration"])
    if expired:
        return await callback.message.answer(
            "⏳ Bu testning belgilangan vaqti tugagan — ustoz yangi vaqt belgilagach qayta yechishingiz "
            "mumkin bo'ladi. Avvalgi natijangiz saqlanib qoldi.")

    payment = None
    if int(t["is_free"]) != 1:
        payment = await POOL.fetchrow(
            "SELECT id FROM payments WHERE user_id=$1 AND test_id=$2 AND status='approved' ORDER BY id DESC LIMIT 1",
            user_db_id, test_id)
        if not payment:
            return await callback.message.answer(
                "⚠️ Bu test uchun tasdiqlangan to'lovingiz topilmadi — avvalgi natijangiz saqlanib qoldi.")

    if (t["exam_type"] or "") not in STAGED_EXAM_TYPES:
        await POOL.execute("DELETE FROM results WHERE user_id=$1 AND test_id=$2", user_db_id, test_id)
        await POOL.execute("DELETE FROM sessions WHERE user_id=$1 AND test_id=$2", user_db_id, test_id)
        await POOL.execute("UPDATE retake_grants SET used_ts=$1 WHERE user_id=$2 AND test_id=$3",
                           now_ts(), user_db_id, test_id)
        if payment:
            # pullik testda vaqt birinchi "Boshlash"dan hisoblanadi — qayta
            # yechish uchun uni nolga qaytaramiz, aks holda vaqt allaqachon
            # tugagan deb hisoblanardi.
            await POOL.execute("UPDATE payments SET started_at=NULL, started_ts=NULL WHERE id=$1", payment["id"])

    if payment is None:
        kb = InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text="🚀 Boshlash", callback_data=f"fstart_{test_id}")]
        ])
    else:
        kb = InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text="🚀 Boshlash", callback_data=f"begin_{payment['id']}")]
        ])

    await callback.message.answer(
        "✅ Eski natijangiz o'chirildi. Qaytadan boshlash uchun tugmani bosing:",
        reply_markup=kb
    )


# ===================== PROFILNI TAHRIRLASH =====================
@router.message(F.text == BTN_PROFILE)
async def profile_menu(message: Message, state: FSMContext):
    user_db_id = await get_user_db_id(message.from_user.id)
    if not user_db_id:
        return await message.answer("Avval /start orqali ro'yxatdan o'ting.")

    row = await POOL.fetchrow(
        "SELECT full_name, phone, region, age, school, grade, payment_card "
        "FROM users WHERE telegram_id=$1",
        message.from_user.id
    )

    if not row: return
    fname, phone = row["full_name"], row["phone"]
    region = row["region"] or "belgilanmagan"
    age = row["age"] if row["age"] is not None else "belgilanmagan"
    school = row["school"] or "belgilanmagan"
    grade = row["grade"] or "belgilanmagan"
    is_admin = await is_admin_user(message.from_user.id)

    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="✏️ Ism-familiya", callback_data="edit_prof_name"),
         InlineKeyboardButton(text="📱 Telefon", callback_data="edit_prof_phone")],
        [InlineKeyboardButton(text="📍 Viloyat", callback_data="edit_prof_region"),
         InlineKeyboardButton(text="🎂 Yosh", callback_data="edit_prof_age")],
        [InlineKeyboardButton(text="🏫 Maktab/muassasa", callback_data="edit_prof_school"),
         InlineKeyboardButton(text="📚 Sinf/kurs", callback_data="edit_prof_grade")],
    ])

    card_block = ""
    if is_admin:
        card_status = row["payment_card"] or "sozlanmagan (umumiy karta ishlatiladi)"
        card_block = f"\n💳 To'lov kartangiz: <b>{card_status}</b>\n"
        kb.inline_keyboard.append(
            [InlineKeyboardButton(text="💳 To'lov kartamni sozlash", callback_data="edit_prof_paycard")])

    await message.answer(
        f"👤 <b>Sizning profilingiz:</b>\n\n"
        f"Ism: <b>{fname}</b>\n"
        f"Telefon: <b>{phone}</b>\n"
        f"{card_block}\n"
        f"<i>Quyidagilar ixtiyoriy — xohlasangiz to'ldiring:</i>\n"
        f"Viloyat: <b>{region}</b>\n"
        f"Yosh: <b>{age}</b>\n"
        f"Maktab/muassasa: <b>{school}</b>\n"
        f"Sinf/kurs: <b>{grade}</b>\n\n"
        f"O'zgartirish uchun pastdagi tugmalardan birini tanlang:",
        parse_mode=ParseMode.HTML,
        reply_markup=kb
    )


def _region_picker_kb() -> InlineKeyboardMarkup:
    kb = InlineKeyboardMarkup(inline_keyboard=[])
    row = []
    for i, region in enumerate(UZ_REGIONS):
        row.append(InlineKeyboardButton(text=region, callback_data=f"profregion_{i}"))
        if len(row) == 2:
            kb.inline_keyboard.append(row)
            row = []
    if row:
        kb.inline_keyboard.append(row)
    return kb


@router.callback_query(F.data.startswith("edit_prof_"))
async def edit_profile_cb(callback: CallbackQuery, state: FSMContext):
    action = callback.data.split("_", 2)[2]
    if action == "name":
        await callback.message.answer("✏️ Yangi ism-familiyangizni kiriting:", reply_markup=nav_kb(include_back=True))
        await state.set_state(EditProfileState.waiting_name)
    elif action == "phone":
        await callback.message.answer("📱 Yangi telefon raqamingizni yuboring:", reply_markup=phone_kb())
        await state.set_state(EditProfileState.waiting_phone)
    elif action == "region":
        await callback.message.answer("📍 Viloyatingizni tanlang:", reply_markup=_region_picker_kb())
    elif action == "age":
        await callback.message.answer(
            "🎂 Yoshingizni raqamda kiriting (masalan: 16):", reply_markup=nav_kb(include_back=True))
        await state.set_state(EditProfileState.waiting_age)
    elif action == "school":
        await callback.message.answer(
            "🏫 Maktab/muassasa nomini kiriting (masalan: 20-maktab, yoki TATU):",
            reply_markup=nav_kb(include_back=True))
        await state.set_state(EditProfileState.waiting_school)
    elif action == "grade":
        await callback.message.answer(
            "📚 Sinf yoki kursingizni kiriting (masalan: 9-sinf, yoki 2-kurs):",
            reply_markup=nav_kb(include_back=True))
        await state.set_state(EditProfileState.waiting_grade)
    elif action == "paycard":
        if not await is_admin_user(callback.from_user.id):
            await callback.answer("Bu faqat adminlar uchun.", show_alert=True)
            return
        await state.update_data(nav_stack=[])
        await _nav_show_paycard_number(callback.message, state)
    await callback.answer()


@router.callback_query(F.data.startswith("profregion_"))
async def edit_profile_region_pick(callback: CallbackQuery, state: FSMContext):
    idx = int(callback.data.split("_")[1])
    if not (0 <= idx < len(UZ_REGIONS)):
        return await callback.answer()
    region = UZ_REGIONS[idx]
    await POOL.execute("UPDATE users SET region=$1 WHERE telegram_id=$2", region, callback.from_user.id)
    admin_flag = await is_admin_user(callback.from_user.id)
    await callback.answer(f"✅ {region}")
    await callback.message.answer(f"✅ Viloyat <b>{region}</b> qilib saqlandi.", parse_mode=ParseMode.HTML,
                                  reply_markup=main_menu_kb(admin_flag))


@router.message(EditProfileState.waiting_age)
async def edit_age_save(message: Message, state: FSMContext):
    if (message.text or "") in TOP_MENU_BTNS: return
    text = (message.text or "").strip()
    if not text.isdigit() or not (5 <= int(text) <= 100):
        return await message.answer("⚠️ Yoshni 5 dan 100 gacha raqamda kiriting:")
    await POOL.execute("UPDATE users SET age=$1 WHERE telegram_id=$2", int(text), message.from_user.id)
    admin_flag = await is_admin_user(message.from_user.id)
    await message.answer("✅ Yoshingiz saqlandi!", reply_markup=main_menu_kb(admin_flag))
    await state.clear()


@router.message(EditProfileState.waiting_school)
async def edit_school_save(message: Message, state: FSMContext):
    if (message.text or "") in TOP_MENU_BTNS: return
    value = sanitize_plain_text(message.text, max_len=150)
    if len(value) < 1:
        return await message.answer("⚠️ Juda qisqa. Qayta kiriting:")
    await POOL.execute("UPDATE users SET school=$1 WHERE telegram_id=$2", value, message.from_user.id)
    admin_flag = await is_admin_user(message.from_user.id)
    await message.answer("✅ Maktab/muassasa saqlandi!", reply_markup=main_menu_kb(admin_flag))
    await state.clear()


@router.message(EditProfileState.waiting_grade)
async def edit_grade_save(message: Message, state: FSMContext):
    if (message.text or "") in TOP_MENU_BTNS: return
    value = sanitize_plain_text(message.text, max_len=50)
    if len(value) < 1:
        return await message.answer("⚠️ Juda qisqa. Qayta kiriting:")
    await POOL.execute("UPDATE users SET grade=$1 WHERE telegram_id=$2", value, message.from_user.id)
    admin_flag = await is_admin_user(message.from_user.id)
    await message.answer("✅ Sinf/kurs saqlandi!", reply_markup=main_menu_kb(admin_flag))
    await state.clear()


async def _nav_show_paycard_number(answerable, state: FSMContext) -> None:
    await answerable.answer(
        "💳 <b>To'lov kartangizni sozlash</b>\n\n"
        "Bundan keyin siz yaratgan PULLIK testlarni sotib olmoqchi bo'lgan o'quvchilarga "
        "aynan shu karta ko'rsatiladi va to'lov cheklari to'g'ridan-to'g'ri sizga keladi.\n\n"
        "Karta raqamini kiriting (masalan: 8600 1234 5678 9012):",
        parse_mode=ParseMode.HTML, reply_markup=nav_kb(include_back=True)
    )
    await state.set_state(EditProfileState.waiting_payment_card)


async def _nav_show_paycard_owner(answerable, state: FSMContext) -> None:
    await answerable.answer(
        "👤 Karta egasining F.I.Sh. kiriting (masalan: Jasurbek Aktamov):",
        reply_markup=nav_kb(include_back=True)
    )
    await state.set_state(EditProfileState.waiting_payment_owner)


async def _nav_show_paycard_phone(answerable, state: FSMContext) -> None:
    await answerable.answer(
        "📞 Aloqa uchun telefon raqamingizni kiriting (masalan: +998901234567):",
        reply_markup=nav_kb(include_back=True)
    )
    await state.set_state(EditProfileState.waiting_payment_phone)


BASE_NAV_HANDLERS["paycard_number"] = _nav_show_paycard_number
BASE_NAV_HANDLERS["paycard_owner"] = _nav_show_paycard_owner


@router.message(EditProfileState.waiting_payment_card)
async def edit_payment_card_save(message: Message, state: FSMContext):
    if (message.text or "") in TOP_MENU_BTNS:
        return
    if not await is_admin_user(message.from_user.id):
        await state.clear()
        return
    card = sanitize_plain_text(message.text or "", max_len=40)
    if len(card) < 8:
        return await message.answer("⚠️ Karta raqami juda qisqa. Qayta kiriting:")
    await state.update_data(pending_payment_card=card)
    await _nav_push(state, "base", "paycard_number")
    await _nav_show_paycard_owner(message, state)


@router.message(EditProfileState.waiting_payment_owner)
async def edit_payment_owner_save(message: Message, state: FSMContext):
    if (message.text or "") in TOP_MENU_BTNS:
        return
    if not await is_admin_user(message.from_user.id):
        await state.clear()
        return
    owner = sanitize_plain_text(message.text or "", max_len=100)
    if len(owner) < 3:
        return await message.answer("⚠️ Juda qisqa. Qayta kiriting:")
    await state.update_data(pending_payment_owner=owner)
    await _nav_push(state, "base", "paycard_owner")
    await _nav_show_paycard_phone(message, state)


@router.message(EditProfileState.waiting_payment_phone)
async def edit_payment_phone_save(message: Message, state: FSMContext):
    if (message.text or "") in TOP_MENU_BTNS:
        return
    if not await is_admin_user(message.from_user.id):
        await state.clear()
        return
    phone = sanitize_plain_text(message.text or "", max_len=30)
    if len(phone) < 7:
        return await message.answer("⚠️ Juda qisqa. Qayta kiriting:")

    data = await state.get_data()
    card = data.get("pending_payment_card")
    owner = data.get("pending_payment_owner")
    await POOL.execute(
        "UPDATE users SET payment_card=$1, payment_owner=$2, payment_phone=$3 WHERE telegram_id=$4",
        card, owner, phone, message.from_user.id
    )
    admin_flag = await is_admin_user(message.from_user.id)
    await message.answer(
        f"✅ <b>To'lov kartangiz saqlandi!</b>\n\n"
        f"💳 <code>{card}</code>\n👤 {owner}\n📞 {phone}\n\n"
        f"Bundan buyon siz yaratgan pullik testlarni sotib olganlar aynan shu kartaga to'laydi.",
        parse_mode=ParseMode.HTML, reply_markup=main_menu_kb(admin_flag)
    )
    await state.clear()


@router.message(EditProfileState.waiting_name)
async def edit_name_save(message: Message, state: FSMContext):
    if (message.text or "") in TOP_MENU_BTNS: return
    new_name = sanitize_plain_text(message.text, max_len=100)
    if len(new_name) < 3:
        return await message.answer("⚠️ Ism-familiya juda qisqa yoki noto'g'ri belgilar bor. Qayta kiriting:")
    await POOL.execute("UPDATE users SET full_name=$1 WHERE telegram_id=$2", new_name, message.from_user.id)
    admin_flag = await is_admin_user(message.from_user.id)
    await message.answer("✅ Ismingiz muvaffaqiyatli o'zgartirildi!", reply_markup=main_menu_kb(admin_flag))
    await state.clear()


@router.message(EditProfileState.waiting_phone)
async def edit_phone_save(message: Message, state: FSMContext):
    if (message.text or "") in TOP_MENU_BTNS: return
    contact = message.contact.phone_number if message.contact else normalize_phone_input(message.text)
    if not contact:
        return await message.answer(
            "⚠️ Telefon raqami noto'g'ri. Masalan: <code>+998901234567</code> — qayta yuboring:",
            parse_mode=ParseMode.HTML)
    await POOL.execute("UPDATE users SET phone=$1 WHERE telegram_id=$2", contact, message.from_user.id)
    admin_flag = await is_admin_user(message.from_user.id)
    await message.answer("✅ Raqamingiz muvaffaqiyatli o'zgartirildi!", reply_markup=main_menu_kb(admin_flag))
    await state.clear()


# ===================== UMUMIY XATO USHLAGICH =====================
async def global_error_handler(event: ErrorEvent, bot: Bot):
    """
    Har qanday handler ichida kutilmagan xatolik chiqsa (biz payqamagan bug,
    Telegram API xatosi va h.k.), avval bu shunchaki jim o'tib ketardi —
    foydalanuvchi tugmani bossa hech qanday javob ko'rmasdi. Endi bunday
    holatda ham foydalanuvchiga tushunarli xabar chiqadi, ham to'liq xato
    server logiga yoziladi (diagnostika uchun).
    """
    logging.exception(f"Ushlanmagan xatolik: {event.exception}")

    chat_id = None
    try:
        upd = event.update
        if upd.message:
            chat_id = upd.message.chat.id
        elif upd.callback_query and upd.callback_query.message:
            chat_id = upd.callback_query.message.chat.id
    except Exception:
        pass

    if chat_id:
        try:
            await bot.send_message(
                chat_id,
                "⚠️ Kutilmagan xatolik yuz berdi. Iltimos, qayta urinib ko'ring "
                "yoki 🏠 Bosh menyu tugmasini bosing."
            )
        except Exception:
            pass

    return True


# ✅ Sahifalash tugmalari (tlp_<list_key>_<page>) bosilganda qaysi ro'yxatni
# qayta chizish kerakligini shu lug'at orqali topamiz.
ALLTESTS_PAGE_SIZE = 10


async def _render_alltests_list(answerable, page: int = 0):
    offset = max(0, page) * ALLTESTS_PAGE_SIZE
    rows = await POOL.fetch("""
        SELECT t.id, t.code, t.title, t.is_free, t.exam_type, t.creator_id, u.full_name AS creator_name
        FROM tests t
        LEFT JOIN users u ON u.telegram_id = t.creator_id
        ORDER BY t.id DESC
        LIMIT $1 OFFSET $2
    """, ALLTESTS_PAGE_SIZE + 1, offset)
    has_more = len(rows) > ALLTESTS_PAGE_SIZE
    rows = rows[:ALLTESTS_PAGE_SIZE]

    if not rows and page == 0:
        return await answerable.answer("Hozircha testlar yo'q.")
    if not rows:
        return await answerable.answer("Boshqa test yo'q.")

    text = f"📚 <b>Barcha testlar</b> ({page + 1}-sahifa)\n\n"
    for r in rows:
        kind = "🆓" if r["is_free"] else "🧾"
        creator = _h(r["creator_name"] or "noma'lum")
        creator_id_part = f" · <code>{r['creator_id']}</code>" if r["creator_id"] else " · admin"
        text += (
            f"{kind} <code>{r['code']}</code> — <b>{_h(r['title'])}</b>\n"
            f"   👤 {creator}{creator_id_part} · {r['exam_type'] or 'simple'}\n\n"
        )

    kb = InlineKeyboardMarkup(inline_keyboard=[])
    nav = _pagination_nav_row("alltests", page, has_more)
    if nav:
        kb.inline_keyboard.append(nav)

    await answerable.answer(text, parse_mode=ParseMode.HTML, reply_markup=kb if kb.inline_keyboard else None)


@router.callback_query(F.data == "alltests_menu")
async def alltests_menu(callback: CallbackQuery):
    if not await is_admin_user(callback.from_user.id):
        return
    await _render_alltests_list(callback.message, 0)


# ===================== KUTILMAGAN XABAR / ESKIRGAN TUGMA (eng oxirida) =====================
# ✅ Bu ikki handler ATAYLAB eng oxirida ro'yxatdan o'tkazilgan — boshqa
# hech qaysi handler mos kelmagandagina ishlaydi. Avval bunday hollarda bot
# butunlay jim qolardi (masalan to'lov chekini kutayotganda matn yuborilsa,
# test fayli o'rniga rasm yuborilsa, tugma kutilganda matn yozilsa, yoki
# eski xabardagi tugma bosilsa — tugma "osilib" qolardi).
_STATE_HINTS = {
    PaymentState.waiting_for_screenshot.state:
        "📸 To'lov chekini <b>rasm (skrinshot)</b> ko'rinishida yuboring.",
    AdminAddTestState.file.state: "📄 Test faylini <b>PDF hujjat</b> ko'rinishida yuboring.",
    AdminAddTestState.is_free.state: "👆 Yuqoridagi tugmalardan birini tanlang (tekin yoki pullik).",
    AdminAddTestState.exam_type.state: "👆 Yuqoridagi tugmalardan baholash rejimini tanlang.",
    AdminAddTestState.schedule_mode.state: "👆 Yuqoridagi tugmalardan boshlanish rejimini tanlang.",
    CertConfigState.waiting_template.state: "👆 Yuqoridagi 3 ta namunadan birini tugma orqali tanlang.",
    RestoreState.waiting_file.state: "📥 Zaxira <b>.json</b> faylini hujjat sifatida yuboring.",
    AdminEditTestState.waiting_field.state: "👆 Qaysi maydonni o'zgartirishni tugma orqali tanlang.",
    TestProcessState.solving.state: "👆 Javoblarni tugmalar orqali belgilang, so'ng «✅ Yakunlash»ni bosing.",
    StagedSolvingState.in_section.state: "👆 Ko'rsatilgan topshiriq turida javob bering.",
}


@router.message(StateFilter("*"))
async def fallback_message(message: Message, state: FSMContext):
    st = await state.get_state()
    if st is None:
        admin_flag = await is_admin_user(message.from_user.id)
        await message.answer("🤖 Kerakli bo'limni pastdagi menyudan tanlang 👇",
                             reply_markup=main_menu_kb(admin_flag))
        return
    hint = _STATE_HINTS.get(st, "⚠️ Bu turdagi xabar hozir kutilmayapti. Yuqoridagi ko'rsatmaga amal qiling "
                                "yoki «🏠 Bosh menyu»ni bosing.")
    await message.answer(hint, parse_mode=ParseMode.HTML)


@router.callback_query()
async def fallback_callback(callback: CallbackQuery):
    data = callback.data or ""
    if data.startswith(("t_", "sq_")):
        text = "⏱ Bu test sessiyasi yakunlangan yoki eskirgan. Testni ro'yxatdan qayta oching."
    else:
        text = "⚠️ Bu tugma eskirgan yoki hozir ishlamaydi. Menyudan qaytadan oching."
    await callback.answer(text, show_alert=True)


_PAGINATED_LIST_RENDERERS = {
    "edit": _render_edittest_list,
    "del": _render_admdel_list,
    "hide": _render_hidetest_list,
    "close": _render_closetest_list,
    "xl": _render_xl_list,
    "trange": _render_trange_list,
    "alltests": _render_alltests_list,
}

# ✅ Foydalanuvchiga ko'rinadigan ro'yxatlar uchun xuddi shunday, lekin
# ulp_<list_key>_<page> orqali (ustidan qarang: user_list_page_nav).
_USER_PAGINATED_RENDERERS = {
    "free": _render_free_tests,
    "paid": _render_paid_tests,
    "mytests": _render_my_tests,
    "created": _render_created_tests,
    "rttests": _render_rating_tests_list,
}


# ===================== MINI APP (veb-ilova: profil + interaktiv test yechish) =====================
# ✅ Telegram Mini App — bot menyusidagi "📱 Mini App" tugmasi (yoki test
# kodi topilganda chiqadigan tugma) orqali Telegram ichida ochiladigan
# veb-ilova. Ikki qismdan iborat:
#   1) 📝 Test yechish — o'quvchi test kodini kiritadi, ilova test turiga
#      moslashgan interaktiv oyna ochadi: DTM javoblar varaqasi (doirachalar),
#      SAT uchun kalkulyator/grafik/formulalar, ochiq savollar uchun yozish
#      maydoni (+ rasm), bosqichli imtihonlarda bo'limma-bo'lim o'tish,
#      Writing uchun matn muharriri, Listening uchun audio pleyer.
#   2) 👤 Profil — statistika, natijalar, sertifikatlar.
#
# MUHIM: ilova botning O'ZIDAGI qoidalar va funksiyalardan foydalanadi —
# sessiyalar (sessions), taymer, to'lov holati, baholash
# (process_test_submission, _finalize_staged_attempt), sertifikat. Natija
# chatga ham avvalgidek yuboriladi. Ilovada boshlangan test chatda (yoki
# aksincha) davom ettirilsa ham javoblar yo'qolmaydi.
#
# Xavfsizlik: har bir so'rov Telegram "initData" imzosi bilan tekshiriladi
# (validate_webapp_init_data); javob kalitlari hech qachon brauzerga
# yuborilmaydi (faqat qaysi savol ochiq ekani ko'rsatiladi — botda ham shunday).
def validate_webapp_init_data(init_data: str, bot_token: str, max_age_sec: int = 86400) -> dict | None:
    if not init_data:
        return None
    try:
        parsed = dict(parse_qsl(init_data, strict_parsing=True))
    except ValueError:
        return None

    received_hash = parsed.pop("hash", None)
    if not received_hash:
        return None

    auth_date = parsed.get("auth_date")
    if not auth_date:
        return None
    try:
        if now_ts() - int(auth_date) > max_age_sec:
            return None
    except ValueError:
        return None

    data_check_string = "\n".join(f"{k}={v}" for k, v in sorted(parsed.items()))
    secret_key = hmac.new(b"WebAppData", bot_token.encode(), hashlib.sha256).digest()
    computed_hash = hmac.new(secret_key, data_check_string.encode(), hashlib.sha256).hexdigest()
    if not hmac.compare_digest(computed_hash, received_hash):
        return None

    user_json = parsed.get("user")
    if not user_json:
        return None
    try:
        return json.loads(user_json)
    except json.JSONDecodeError:
        return None


async def _webapp_auth_user_row(request: web.Request):
    """initData'ni header'dan (fetch so'rovlari) yoki query'dan (oddiy <a href>
    havolalar — brauzer havola bosilganda maxsus header qo'sha olmaydi)
    o'qiydi, tekshiradi va shu Telegram ID'ga tegishli users qatorini qaytaradi."""
    init_data = request.headers.get("X-Telegram-Init-Data") or request.query.get("initData", "")
    tg_user = validate_webapp_init_data(init_data, BOT_TOKEN)
    if not tg_user or "id" not in tg_user:
        return None
    row = await POOL.fetchrow("SELECT id, full_name FROM users WHERE telegram_id=$1", int(tg_user["id"]))
    return row


WEB_EXAM_LABELS = {
    "simple": "Oddiy test", "rasch": "Rasch (MS)", "dtm": "DTM (5 fan)", "cefr": "CEFR", "ielts": "IELTS",
    "sat": "SAT", "dtm_staged": "DTM (5 bosqich)", "ielts4": "IELTS (4 ko'nikma)",
    "cefr4": "CEFR (4 ko'nikma)", "sat_staged": "SAT (R&W + Math)",
}
WEB_DEADLINE_TASKS: dict[tuple[int, int], asyncio.Task] = {}
WEB_DEADLINE_GRACE_SEC = 20   # ilova o'zi topshirishga ulgurmasa (internet uzilgan) — server topshiradi
WEB_FILE_CACHE_DIR = os.path.join(tempfile.gettempdir(), "edutest_webapp_files")
WEB_MAX_PHOTO_BYTES = 10 * 1024 * 1024
WEB_RATE_WINDOW_SEC = 60
WEB_RATE_LIMIT = 240          # avtomatik saqlash bilan birga ham bemalol yetadi
WEB_OPEN_TEXT_MAX = 500       # botdagi ochiq savol javobi chegarasi bilan bir xil
WEB_WRITING_TEXT_MAX = 4000   # botdagi Writing chegarasi bilan bir xil
_web_hits: dict[int, deque] = defaultdict(deque)
_web_file_locks: dict[str, asyncio.Lock] = defaultdict(asyncio.Lock)


class WebError(Exception):
    def __init__(self, code: str, message: str, status: int = 400, **extra):
        super().__init__(message)
        self.code = code
        self.message = message
        self.status = status
        self.extra = extra


def _web_handler(fn):
    async def wrapper(request: web.Request):
        try:
            return await fn(request)
        except WebError as e:
            return web.json_response({"ok": False, "error": e.code, "message": e.message, **e.extra},
                                     status=e.status)
        except Exception:
            logging.exception(f"Mini App API xatosi: {request.method} {request.path}")
            return web.json_response(
                {"ok": False, "error": "server", "message": "Serverda kutilmagan xatolik. Qayta urinib ko'ring."},
                status=500)
    wrapper.__name__ = fn.__name__
    return wrapper


def _web_rate_ok(tg_id: int) -> bool:
    now_ = time.time()
    dq = _web_hits[tg_id]
    while dq and now_ - dq[0] > WEB_RATE_WINDOW_SEC:
        dq.popleft()
    if len(dq) >= WEB_RATE_LIMIT:
        return False
    dq.append(now_)
    return True


async def _web_auth(request: web.Request, need_registered: bool = True) -> dict:
    init_data = request.headers.get("X-Telegram-Init-Data") or request.query.get("initData", "")
    tg_user = validate_webapp_init_data(init_data, BOT_TOKEN)
    if not tg_user or "id" not in tg_user:
        raise WebError("unauthorized", "Ilovani Telegram ichidan, bot menyusi orqali oching.", 401)
    tg_id = int(tg_user["id"])
    if not _web_rate_ok(tg_id):
        raise WebError("rate", "Juda ko'p so'rov yuborildi. Bir oz kutib, qayta urinib ko'ring.", 429)
    blocked, remain = await SECURITY.is_blocked(tg_id)
    if blocked:
        raise WebError("blocked", f"Siz vaqtincha bloklangansiz. Qolgan vaqt: {fmt_seconds(remain)}", 403)
    row = await POOL.fetchrow("SELECT id, full_name FROM users WHERE telegram_id=$1", tg_id)
    registered = bool(row and (row["full_name"] or "").strip())
    if need_registered and not registered:
        raise WebError("not_registered", "Avval botda ro'yxatdan o'ting: botga /start yuboring.", 403)
    return {"tg_id": tg_id, "user_db_id": row["id"] if row else None,
            "full_name": row["full_name"] if row else None, "registered": registered}


async def _web_body(request: web.Request) -> dict:
    try:
        data = await request.json()
    except Exception:
        raise WebError("bad_request", "So'rov formati noto'g'ri.")
    if not isinstance(data, dict):
        raise WebError("bad_request", "So'rov formati noto'g'ri.")
    return data


def _web_int(value, name: str = "id") -> int:
    try:
        return int(value)
    except (TypeError, ValueError):
        raise WebError("bad_request", f"Noto'g'ri {name}.")


def _web_lock(bot: Bot, tg_id: int):
    """Botdagi tugmalar bilan BIR NAVBATDA ishlash uchun foydalanuvchi qulfi
    (SimpleEventIsolation) — chat va ilova bir vaqtda bir testni buzmasligi uchun."""
    if EVENTS_ISOLATION is None:
        return contextlib.nullcontext()
    return EVENTS_ISOLATION.lock(StorageKey(bot_id=bot.id, chat_id=tg_id, user_id=tg_id))


def _web_fsm(bot: Bot, tg_id: int) -> FSMContext | None:
    if FSM_STORAGE is None:
        return None
    return FSMContext(storage=FSM_STORAGE, key=StorageKey(bot_id=bot.id, chat_id=tg_id, user_id=tg_id))


async def _web_bot_username(bot: Bot) -> str:
    try:
        me = await bot.me()
        return me.username or ""
    except Exception:
        return ""


async def _web_require_subscribed(bot: Bot, tg_id: int) -> None:
    if await is_admin_user(tg_id):
        return
    ok, not_joined = await check_user_subscribed(bot, tg_id)
    if ok:
        return
    channels = []
    for r in not_joined:
        url = (r["join_url"] or "").strip() or (f"https://t.me/{r['username']}" if r["username"] else "")
        channels.append({"title": r["title"] or (("@" + r["username"]) if r["username"] else "Kanal"), "url": url})
    raise WebError("not_subscribed", "Davom etish uchun quyidagi kanal/guruhlarga a'zo bo'ling.", 403,
                   channels=channels)


def _web_n_opts(key: str) -> int:
    """Variantlar soni — kalitdagi eng katta harfga qarab (kamida A-D).
    Butun test uchun bitta son: qaysi savolda qaysi javob ekanini oshkor qilmaydi."""
    letters = [ord(c) - 96 for c in (key or "") if c in "abcdef"]
    return max(4, max(letters) if letters else 4)


def _web_layout(exam_type: str, dtm_cfg_raw: str | None, qcount: int) -> list[dict]:
    if (exam_type or "") == "dtm":
        try:
            cfg = json.loads(dtm_cfg_raw) if dtm_cfg_raw else DTM_DEFAULT_CFG
            blocks, start = [], 1
            for s in cfg.get("subjects", []):
                q = int(s["q"])
                blocks.append({"name": str(s.get("name") or ""), "from": start, "to": start + q - 1,
                               "w": float(s.get("w") or 1)})
                start += q
            if start - 1 == qcount and blocks:
                return blocks
        except (ValueError, KeyError, TypeError, json.JSONDecodeError):
            pass
    return [{"name": "", "from": 1, "to": qcount}]


def _web_tools(exam_type: str, section_key: str | None = None) -> dict:
    et = exam_type or "simple"
    sat_math = et == "sat" or (et == "sat_staged" and section_key == "math")
    sat_any = et in ("sat", "sat_staged")
    return {"calculator": sat_math, "reference": sat_math, "eliminator": sat_any,
            "default_view": "card" if sat_any else "sheet"}


def _web_question_payload(key: str, qcount: int, saved: dict) -> dict:
    key = key[:qcount]
    open_pos = [i + 1 for i, c in enumerate(key) if c == "x"]
    return {
        "qcount": qcount, "n_opts": _web_n_opts(key), "open": open_pos,
        "answers": {k: v for k, v in (saved.get("answers") or {}).items() if isinstance(v, str)},
        "open_answers": {k: {"text": (v or {}).get("text"), "photo_id": (v or {}).get("photo_id")}
                         for k, v in (saved.get("open_answers") or {}).items() if isinstance(v, dict)},
        "flags": [int(x) for x in (saved.get("flags") or []) if str(x).isdigit()],
        "current_q": int(saved.get("current_q") or 1),
        "updated": int(saved.get("updated") or 0),
    }


def _web_clean_answers(raw_answers, raw_open, key: str, qcount: int, allowed_photos: set) -> tuple[dict, dict]:
    """Brauzerdan kelgan javoblarni bot formatiga keltiradi: variantli savolga
    faqat ruxsat etilgan harf, ochiq (x) savolga matn yoki SHU foydalanuvchi
    yuklagan rasm. Ochiq savolga javob bo'lsa answers[q] = "x" (botdagidek)."""
    raw_answers = raw_answers if isinstance(raw_answers, dict) else {}
    raw_open = raw_open if isinstance(raw_open, dict) else {}
    letters = "abcdef"[:_web_n_opts(key)]
    answers, open_answers = {}, {}
    for q in range(1, qcount + 1):
        sq = str(q)
        if q - 1 < len(key) and key[q - 1] == "x":
            o = raw_open.get(sq) if isinstance(raw_open.get(sq), dict) else {}
            text = sanitize_plain_text(str(o.get("text") or ""), max_len=WEB_OPEN_TEXT_MAX) or None
            pid = o.get("photo_id")
            pid = pid if isinstance(pid, str) and pid in allowed_photos else None
            if text or pid:
                open_answers[sq] = {"text": text, "photo_id": pid}
                answers[sq] = "x"
        else:
            v = str(raw_answers.get(sq) or "").strip().lower()
            if v in letters and len(v) == 1:
                answers[sq] = v
    return answers, open_answers


def _web_clean_flags(raw, qcount: int) -> list[int]:
    out = []
    for x in raw if isinstance(raw, list) else []:
        try:
            n = int(x)
        except (TypeError, ValueError):
            continue
        if 1 <= n <= qcount and n not in out:
            out.append(n)
    return out[:qcount]


async def _web_load_saved(user_db_id: int, test_id: int, mode: str) -> dict:
    return _load_json_dict(await POOL.fetchval(
        "SELECT saved_answers FROM sessions WHERE user_id=$1 AND test_id=$2 AND mode=$3", user_db_id, test_id, mode))


async def _web_store_saved(user_db_id: int, test_id: int, mode: str, payload: dict) -> None:
    payload["updated"] = now_ts()
    await POOL.execute("UPDATE sessions SET saved_answers=$1 WHERE user_id=$2 AND test_id=$3 AND mode=$4",
                       json.dumps(payload, ensure_ascii=False), user_db_id, test_id, mode)


async def _web_test_row(code: str | None = None, test_id: int | None = None):
    q = """
        SELECT id, code, title, subject, is_free, price, duration, questions_count, file_id, answers,
               start_mode, start_ts, start_at, exam_type, dtm_cfg, skills_cfg,
               COALESCE(is_hidden,0) AS is_hidden, COALESCE(is_closed,0) AS is_closed,
               COALESCE(rush_finalized,0) AS rush_finalized, open_ref_answers
        FROM tests WHERE """
    if test_id is not None:
        return await POOL.fetchrow(q + "id=$1", test_id)
    return await POOL.fetchrow(q + "code=$1", normalize_code(code or ""))


def _web_test_info(t) -> dict:
    et = t["exam_type"] or "simple"
    info = {
        "id": t["id"], "code": t["code"], "title": t["title"], "subject": t["subject"],
        "exam_type": et, "exam_label": WEB_EXAM_LABELS.get(et, et), "is_free": int(t["is_free"] or 0) == 1,
        "price": int(t["price"] or 0), "duration": int(t["duration"] or 0),
        "qcount": int(t["questions_count"] or 0), "scheduled": (t["start_mode"] or "normal") == "scheduled",
        "start_at": t["start_at"], "staged": et in STAGED_EXAM_TYPES,
    }
    if info["staged"]:
        cfg = _load_skills_cfg(t["skills_cfg"])
        info["sections"] = [
            {"key": k, "name": cfg["sections"][k].get("name") or k, "type": cfg["sections"][k].get("type"),
             "type_label": STAGE_TYPE_LABELS.get(cfg["sections"][k].get("type"), ""),
             "duration_min": int(cfg["sections"][k].get("duration_min") or 0),
             "qcount": int(cfg["sections"][k].get("qcount") or 0)}
            for k in cfg.get("order") or [] if k in cfg.get("sections", {})
        ]
    return info


async def _web_latest_payment(user_db_id: int, test_id: int):
    return await POOL.fetchrow(
        "SELECT id, status, started_ts, reject_reason FROM payments WHERE user_id=$1 AND test_id=$2 "
        "ORDER BY id DESC LIMIT 1", user_db_id, test_id)


async def _web_latest_attempt(user_db_id: int, test_id: int):
    return await POOL.fetchrow(
        "SELECT id, status, section_idx, sections_json, overall_grade, overall_score, attempt_no "
        "FROM staged_attempts WHERE user_id=$1 AND test_id=$2 ORDER BY id DESC LIMIT 1", user_db_id, test_id)


async def _web_access(user_db_id: int, t) -> dict:
    """Bot qoidalari bilan bir xil: shu o'quvchi shu testni hozir boshlay
    oladimi (free_test_start / begin_test / start_staged_attempt mantig'i)."""
    et = t["exam_type"] or "simple"
    staged = et in STAGED_EXAM_TYPES
    is_free = int(t["is_free"] or 0) == 1
    payment = None
    if is_free:
        if int(t["is_hidden"]) == 1:
            return {"state": "hidden"}
    else:
        payment = await _web_latest_payment(user_db_id, t["id"])
        if not payment or payment["status"] not in ("pending", "rejected", "approved"):
            return {"state": "need_payment"}
        if payment["status"] == "pending":
            return {"state": "payment_pending"}
        if payment["status"] == "rejected":
            return {"state": "payment_rejected", "reason": payment["reject_reason"]}

    if staged:
        if int(t["is_closed"]) == 1:
            return {"state": "closed"}
        a = await _web_latest_attempt(user_db_id, t["id"])
        if a:
            if a["status"] == "in_progress":
                return {"state": "resume", "attempt_id": a["id"]}
            if a["status"] == "pending_review":
                return {"state": "pending_review"}
            # status == 'graded' — botdagidek qayta yechish mumkin (eski natija saqlanadi)
            can_retake = int(t["is_hidden"]) == 0
            return {"state": "staged_done", "can_retake": can_retake, "grade": a["overall_grade"]}
        return {"state": "ready"}

    mode = "free" if is_free else "paid"
    session = await get_session(user_db_id, t["id"], mode)
    if int(t["is_closed"]) == 1 and (is_free or not session):
        return {"state": "closed"}
    blocked, left = is_scheduled_not_started(t["start_mode"], t["start_ts"])
    if blocked:
        return {"state": "not_started", "left": left}
    if await POOL.fetchval("SELECT 1 FROM results WHERE user_id=$1 AND test_id=$2 LIMIT 1", user_db_id, t["id"]):
        return {"state": "done"}
    key = normalize_answers(t["answers"] or "")
    if len(key) != int(t["questions_count"] or 0):
        return {"state": "broken"}

    if session:
        start_ts_, dur = int(session["start_ts"]), int(session["duration_sec"])
    else:
        dur = int(t["duration"] or 0) * 60
        if (t["start_mode"] or "normal") == "scheduled" and int(t["start_ts"] or 0) > 0:
            start_ts_ = int(t["start_ts"])
        elif payment and payment["started_ts"]:
            start_ts_ = int(payment["started_ts"])
        else:
            start_ts_ = None
    if start_ts_ is not None and now_ts() > start_ts_ + dur + 60:
        return {"state": "expired"}
    if session:
        return {"state": "resume", "end_ts": start_ts_ + dur}
    return {"state": "ready"}


WEB_STATE_MESSAGES = {
    "ready": "Test boshlashga tayyor.",
    "resume": "Bu testni avval boshlagansiz — qolgan vaqtdan davom etasiz, javoblaringiz saqlangan.",
    "done": "Siz bu testni allaqachon topshirgansiz.",
    "hidden": "Bu test vaqtincha to'xtatilgan. Keyinroq urinib ko'ring.",
    "closed": "Bu test yakunlangan — yangi urinish qabul qilinmaydi.",
    "not_started": "Test hali boshlanmagan.",
    "expired": "Bu testning vaqti tugagan.",
    "broken": "Test noto'g'ri sozlangan (javob kaliti savollar soniga mos emas). Ustozga xabar bering.",
    "need_payment": "Bu test pullik. To'lov bot orqali qilinadi: «🧾 Pullik testlar» bo'limi.",
    "payment_pending": "To'lov chekingiz tekshirilmoqda. Tasdiqlangach shu yerda boshlay olasiz.",
    "payment_rejected": "To'lovingiz rad etilgan. Botdagi «📦 Pullik testlarim» orqali qayta yuboring.",
    "pending_review": "Javoblaringiz yuborilgan — ustoz baholayapti. Natija botga keladi.",
    "staged_done": "Siz bu imtihonni topshirgansiz. Xohlasangiz qaytadan yechishingiz mumkin (eski natija saqlanadi).",
}


def _web_schedule_deadline(bot: Bot, tg_id: int, test_id: int, end_ts: int) -> None:
    key = (tg_id, test_id)
    old = WEB_DEADLINE_TASKS.pop(key, None)
    if old and not old.done():
        old.cancel()
    WEB_DEADLINE_TASKS[key] = asyncio.create_task(_web_deadline(bot, tg_id, test_id, end_ts))


def _web_cancel_deadline(tg_id: int, test_id: int) -> None:
    task = WEB_DEADLINE_TASKS.pop((tg_id, test_id), None)
    if task and not task.done() and task is not asyncio.current_task():
        task.cancel()


async def _web_deadline(bot: Bot, tg_id: int, test_id: int, end_ts: int) -> None:
    """Ilovada boshlangan test/bo'lim vaqti tugaganda (ilova yopilgan yoki
    internet uzilgan bo'lsa ham) oxirgi saqlangan javoblar bilan topshiradi."""
    key = (tg_id, test_id)
    try:
        await asyncio.sleep(max(0, end_ts - now_ts()) + WEB_DEADLINE_GRACE_SEC)
        if WEB_DEADLINE_TASKS.get(key) is asyncio.current_task():
            WEB_DEADLINE_TASKS.pop(key, None)
        async with _web_lock(bot, tg_id):
            et = await POOL.fetchval("SELECT exam_type FROM tests WHERE id=$1", test_id)
            if (et or "") in STAGED_EXAM_TYPES:
                await _web_staged_timeout(bot, tg_id, test_id, end_ts)
            else:
                await _auto_submit_on_timeout_unlocked(bot, tg_id, test_id, end_ts)
    except asyncio.CancelledError:
        raise
    except Exception:
        logging.exception(f"Mini App: vaqt tugaganda topshirib bo'lmadi (user={tg_id}, test={test_id})")
    finally:
        if WEB_DEADLINE_TASKS.get(key) is asyncio.current_task():
            WEB_DEADLINE_TASKS.pop(key, None)


# ----- Fayllar (test PDF, Listening audio) -----
def _web_sniff_type(head: bytes, kind: str | None) -> str:
    if head.startswith(b"%PDF"):
        return "application/pdf"
    if head.startswith(b"RIFF") and head[8:12] == b"WAVE":
        return "audio/wav"
    if head.startswith(b"OggS"):
        return "audio/ogg"
    if head.startswith(b"ID3") or head[:2] in (b"\xff\xfb", b"\xff\xf3", b"\xff\xf2"):
        return "audio/mpeg"
    if head[4:8] == b"ftyp":
        return "audio/mp4" if kind == "audio" else "video/mp4"
    if head.startswith(b"\x89PNG"):
        return "image/png"
    if head.startswith(b"\xff\xd8"):
        return "image/jpeg"
    return "application/octet-stream"


async def _web_fetch_file(bot: Bot, file_id: str) -> str:
    os.makedirs(WEB_FILE_CACHE_DIR, exist_ok=True)
    name = hashlib.sha256(file_id.encode()).hexdigest()[:48]
    path = os.path.join(WEB_FILE_CACHE_DIR, name)
    async with _web_file_locks[name]:
        if not os.path.exists(path):
            try:
                tg_file = await bot.get_file(file_id)
            except TelegramBadRequest as e:
                if "too big" in str(e).lower():
                    raise WebError("file_too_big", "Fayl juda katta (20 MB dan ortiq) — uni botdan oling.", 413)
                raise
            tmp = path + ".part"
            await bot.download_file(tg_file.file_path, destination=tmp)
            os.replace(tmp, path)
    return path


# ----- Bitta sessiyali testlar (oddiy, DTM, Rasch, SAT, CEFR, IELTS) -----
async def _web_take_over_chat_single(bot: Bot, tg_id: int, test_id: int) -> None:
    """Shu test chatda ochiq turgan bo'lsa — chatdagi javoblarni sessiyaga
    yozib, chat holatini yopamiz (endi test ilovada davom etadi)."""
    fsm = _web_fsm(bot, tg_id)
    if fsm is None:
        return
    st = await fsm.get_state()
    data = await fsm.get_data()
    if st in (TestProcessState.solving.state, TestProcessState.awaiting_free_text.state) \
            and int(data.get("test_id") or 0) == test_id:
        await _stash_test_progress(fsm, tg_id)
        await fsm.clear()
        cancel_timer(tg_id, test_id)


async def _web_start_single(bot: Bot, ctx: dict, t, access: dict) -> dict:
    user_db_id, tg_id, test_id = ctx["user_db_id"], ctx["tg_id"], t["id"]
    is_free = int(t["is_free"] or 0) == 1
    mode = "free" if is_free else "paid"
    session = await get_session(user_db_id, test_id, mode)
    if session:
        start_ts_, dur = int(session["start_ts"]), int(session["duration_sec"])
    else:
        dur = int(t["duration"] or 0) * 60
        payment = None if is_free else await _web_latest_payment(user_db_id, test_id)
        if (t["start_mode"] or "normal") == "scheduled" and int(t["start_ts"] or 0) > 0:
            start_ts_ = int(t["start_ts"])
        elif payment and payment["started_ts"]:
            start_ts_ = int(payment["started_ts"])
        else:
            start_ts_ = now_ts()
        if payment and not payment["started_ts"]:
            await POOL.execute("UPDATE payments SET started_at=$1, started_ts=$2 WHERE id=$3",
                               now_str_local(), start_ts_, payment["id"])
        await upsert_session(user_db_id, test_id, mode, start_ts_, dur)
    end_ts_ = start_ts_ + dur

    await _web_take_over_chat_single(bot, tg_id, test_id)
    saved = await _web_load_saved(user_db_id, test_id, mode)
    _web_schedule_deadline(bot, tg_id, test_id, end_ts_)

    key = normalize_answers(t["answers"] or "")
    qcount = int(t["questions_count"] or 0)
    et = t["exam_type"] or "simple"
    files = [{"slot": "main", "kind": "document", "name": "Savollar"}] if t["file_id"] else []
    return {
        "kind": "single", "test": _web_test_info(t), "mode": mode, "end_ts": end_ts_, "server_now": now_ts(),
        "layout": _web_layout(et, t["dtm_cfg"], qcount), "tools": _web_tools(et), "files": files,
        **_web_question_payload(key, qcount, saved),
    }


async def _web_result_view(user_db_id: int, test_id: int) -> dict:
    t = await _web_test_row(test_id=test_id)
    if not t:
        raise WebError("not_found", "Test topilmadi.", 404)
    et = t["exam_type"] or "simple"
    base = {"test": _web_test_info(t)}
    if et in STAGED_EXAM_TYPES:
        a = await POOL.fetchrow(
            "SELECT id, status, sections_json, overall_grade, overall_score FROM staged_attempts "
            "WHERE user_id=$1 AND test_id=$2 AND status<>'in_progress' ORDER BY id DESC LIMIT 1", user_db_id, test_id)
        if not a:
            return {**base, "state": "none"}
        cfg = _load_skills_cfg(t["skills_cfg"])
        sections = _load_json_dict(a["sections_json"])
        out = []
        for k in cfg.get("order") or []:
            s = sections.get(k) or {}
            sec_cfg = cfg["sections"].get(k, {})
            out.append({
                "key": k, "name": sec_cfg.get("name") or k, "type": sec_cfg.get("type"),
                "score": s.get("score"), "total": s.get("total"), "band": s.get("band"),
                "earned": s.get("earned"), "max": s.get("max"), "feedback": s.get("feedback"),
                "open_pending": int(s.get("open_pending") or 0), "submitted": k in sections,
            })
        r = await POOL.fetchrow("SELECT id, percent, certificate_sent FROM results WHERE user_id=$1 AND test_id=$2",
                                user_db_id, test_id)
        return {**base, "state": "staged", "status": a["status"], "overall_grade": a["overall_grade"],
                "overall_score": a["overall_score"], "sections": out,
                "percent": float(r["percent"]) if r and r["percent"] is not None else None,
                "certificate": bool(r and int(r["certificate_sent"] or 0) == 1), "result_id": r["id"] if r else None}

    r = await POOL.fetchrow("""
        SELECT id, score, total_questions, percent, earned_points, max_points, details, grade, rasch_ball,
               rasch_percent, user_answers, open_text_answers, open_grades, COALESCE(open_pending,0) AS open_pending,
               COALESCE(certificate_sent,0) AS certificate_sent
        FROM results WHERE user_id=$1 AND test_id=$2
    """, user_db_id, test_id)
    if not r:
        return {**base, "state": "none"}
    key = normalize_answers(t["answers"] or "")
    view = {
        **base, "state": "single", "result_id": r["id"], "score": r["score"], "total": r["total_questions"],
        "percent": float(r["percent"] or 0), "earned": r["earned_points"], "max": r["max_points"],
        "grade": r["grade"], "open_pending": int(r["open_pending"]), "certificate": int(r["certificate_sent"]) == 1,
    }
    open_grades = _load_json_dict(r["open_grades"])
    scored_total = sum(1 for i, c in enumerate(key[:int(t["questions_count"] or 0)])
                       if c != "x" or open_grades.get(str(i + 1)) is not None)
    view["scored_total"] = scored_total
    if et == "dtm":
        view["subjects"] = (_load_json_dict(r["details"]) or {}).get("subjects") or []
    review_allowed = True
    if et == "rasch":
        finalized = int(t["rush_finalized"]) == 1
        view["rasch"] = {"finalized": finalized, "ball": r["rasch_ball"], "percent": r["rasch_percent"],
                         "grade": r["grade"], "remain_text": None if finalized else await remaining_time_text(test_id)}
        end_ts = await get_test_end_ts(test_id)
        review_allowed = end_ts > 0 and now_ts() >= end_ts
    view["review_allowed"] = review_allowed
    if review_allowed and r["user_answers"]:
        user_ans = normalize_answers(r["user_answers"])
        open_answers = _load_json_dict(r["open_text_answers"])
        refs = _load_json_dict(t["open_ref_answers"])
        items = []
        for i in range(min(len(key), len(user_ans))):
            n = str(i + 1)
            if key[i] == "x":
                g = open_grades.get(n)
                oa = open_answers.get(n) or {}
                items.append({"n": i + 1, "open": True, "ok": None if g is None else int(g) == 1,
                              "you": (oa.get("text") if isinstance(oa, dict) else None) or
                                     ("📷 rasm" if isinstance(oa, dict) and oa.get("photo_id") else ""),
                              "key": refs.get(n) or ""})
            else:
                items.append({"n": i + 1, "open": False, "ok": user_ans[i] == key[i],
                              "you": "" if user_ans[i] == "*" else user_ans[i].upper(), "key": key[i].upper()})
        view["review"] = items
    return view


# ----- Bosqichli imtihonlar (DTM 5 bosqich, IELTS/CEFR 4 ko'nikma, SAT 2 bo'lim) -----
def _web_staged_overview(t, attempt) -> dict:
    cfg = _load_skills_cfg(t["skills_cfg"])
    order = cfg.get("order") or []
    sections = _load_json_dict(attempt["sections_json"]) if attempt else {}
    pos = int(attempt["section_idx"] or 0) if attempt else 0
    items = []
    for i, k in enumerate(order):
        s = cfg["sections"].get(k, {})
        items.append({
            "key": k, "name": s.get("name") or k, "type": s.get("type"),
            "type_label": STAGE_TYPE_LABELS.get(s.get("type"), ""),
            "duration_min": int(s.get("duration_min") or 0), "qcount": int(s.get("qcount") or 0),
            "status": "done" if k in sections else ("current" if i == pos else "locked"),
        })
    return {"kind": "staged", "test": _web_test_info(t), "attempt_id": attempt["id"] if attempt else None,
            "pos": pos, "sections": items, "server_now": now_ts()}


async def _web_staged_attempt(user_db_id: int, test_id: int):
    return await POOL.fetchrow(
        "SELECT id, status, section_idx, sections_json FROM staged_attempts "
        "WHERE user_id=$1 AND test_id=$2 AND status='in_progress' ORDER BY id DESC LIMIT 1", user_db_id, test_id)


async def _web_take_over_chat_staged(bot: Bot, tg_id: int, test_id: int, user_db_id: int, key: str) -> None:
    """Shu bosqich chatda ochiq bo'lsa — chatdagi belgilangan javoblarni
    sessiyaga ko'chirib, chat holatini yopamiz."""
    fsm = _web_fsm(bot, tg_id)
    if fsm is None:
        return
    st = await fsm.get_state()
    data = await fsm.get_data()
    if st not in (StagedSolvingState.in_section.state, StagedSolvingState.awaiting_free_text.state) \
            or data.get("staged_test_id") != test_id:
        return
    if data.get("stage_current_key") == key:
        payload = {"answers": data.get("stage_mcq_answers") or {},
                   "open_answers": data.get("stage_mcq_open_answers") or {},
                   "current_q": data.get("stage_mcq_q", 1)}
        if await get_session(user_db_id, test_id, f"stage_{key}"):
            await _web_store_saved(user_db_id, test_id, f"stage_{key}", payload)
    await fsm.clear()
    cancel_timer(tg_id, test_id)


async def _web_staged_finish_section(bot: Bot, tg_id: int, user_db_id: int, t, attempt_id: int,
                                     key: str, raw: dict) -> dict:
    """Bo'limni yakunlaydi (botdagi _complete_current_stage_mcq /
    _save_stage_result_and_advance bilan bir xil natija), keyingisiga o'tadi
    yoki butun imtihonni yakunlab baholaydi. Chatga keyingi bo'lim YUBORILMAYDI."""
    test_id = t["id"]
    cfg = _load_skills_cfg(t["skills_cfg"])
    order = cfg.get("order") or []
    sec = cfg["sections"][key]
    stype = sec.get("type")
    saved = await _web_load_saved(user_db_id, test_id, f"stage_{key}")
    allowed_photos = set(saved.get("uploads") or [])

    if stype in (STAGE_TYPE_MCQ, STAGE_TYPE_MCQ_BAND):
        skey = normalize_answers(sec.get("answers") or "")
        qcount = int(sec.get("qcount") or 0)
        answers, open_answers = _web_clean_answers(raw.get("answers"), raw.get("open_answers"), skey, qcount,
                                                   allowed_photos)
        user_ans = "".join(answers.get(str(i), "*") for i in range(1, qcount + 1))
        result = _build_stage_mcq_result(cfg, key, user_ans, open_answers)
    else:
        text = sanitize_plain_text(str(raw.get("text") or ""), max_len=WEB_WRITING_TEXT_MAX) or None
        pid = raw.get("photo_id")
        pid = pid if isinstance(pid, str) and pid in allowed_photos else None
        result = {"text": text, "photo_id": pid, "band": None, "feedback": None}

    await delete_session(user_db_id, test_id, f"stage_{key}")
    _web_cancel_deadline(tg_id, test_id)
    cancel_timer(tg_id, test_id)

    row = await POOL.fetchrow("SELECT sections_json FROM staged_attempts WHERE id=$1", attempt_id)
    sections = _load_json_dict(row["sections_json"]) if row else {}
    sections[key] = result
    pos = order.index(key) + 1
    await POOL.execute("UPDATE staged_attempts SET sections_json=$1, section_idx=$2 WHERE id=$3",
                       json.dumps(sections, ensure_ascii=False), pos, attempt_id)

    if pos < len(order):
        attempt = await POOL.fetchrow(
            "SELECT id, status, section_idx, sections_json FROM staged_attempts WHERE id=$1", attempt_id)
        return {"finished": False, "overview": _web_staged_overview(t, attempt)}

    await _finalize_staged_attempt(tg_id, bot, _DetachedState({}), attempt_id, test_id, sections)
    return {"finished": True, "result": await _web_result_view(user_db_id, test_id)}


async def _web_staged_timeout(bot: Bot, tg_id: int, test_id: int, end_ts: int) -> None:
    user_db_id = await get_user_db_id(tg_id)
    if not user_db_id:
        return
    a = await _web_staged_attempt(user_db_id, test_id)
    t = await _web_test_row(test_id=test_id)
    if not a or not t:
        return
    cfg = _load_skills_cfg(t["skills_cfg"])
    order = cfg.get("order") or []
    pos = int(a["section_idx"] or 0)
    if pos >= len(order):
        return
    key = order[pos]
    sess = await get_session(user_db_id, test_id, f"stage_{key}")
    if not sess or abs(int(sess["start_ts"]) + int(sess["duration_sec"]) - int(end_ts)) > 5:
        return
    saved = await _web_load_saved(user_db_id, test_id, f"stage_{key}")
    try:
        await bot.send_message(tg_id, "⏱ <b>Bo'lim vaqti tugadi</b> — Mini App'da belgilagan javoblaringiz "
                                      "avtomatik topshirildi.", parse_mode=ParseMode.HTML)
    except (TelegramForbiddenError, TelegramBadRequest):
        pass
    await _web_staged_finish_section(bot, tg_id, user_db_id, t, a["id"], key, saved)


async def _web_staged_section_payload(bot: Bot, ctx: dict, t, attempt) -> dict:
    user_db_id, tg_id, test_id = ctx["user_db_id"], ctx["tg_id"], t["id"]
    cfg = _load_skills_cfg(t["skills_cfg"])
    order = cfg.get("order") or []
    pos = int(attempt["section_idx"] or 0)
    if pos >= len(order):
        raise WebError("no_section", "Barcha bo'limlar topshirilgan.")
    key = order[pos]
    sec = cfg["sections"][key]
    stype = sec.get("type")
    et = t["exam_type"]
    base = {"kind": "section", "test": _web_test_info(t), "key": key, "name": sec.get("name") or key,
            "type": stype, "pos": pos, "total": len(order), "duration_min": int(sec.get("duration_min") or 0)}
    if stype == STAGE_TYPE_VOICE:
        return {**base, "handoff": True, "prompt": sec.get("prompt") or "",
                "bot_username": await _web_bot_username(bot)}

    await _web_take_over_chat_staged(bot, tg_id, test_id, user_db_id, key)
    sess = await get_session(user_db_id, test_id, f"stage_{key}")
    if sess:
        start_ts_, dur = int(sess["start_ts"]), int(sess["duration_sec"])
    else:
        start_ts_, dur = now_ts(), int(sec.get("duration_min") or 0) * 60
        await upsert_session(user_db_id, test_id, f"stage_{key}", start_ts_, dur)
    end_ts_ = start_ts_ + dur
    saved = await _web_load_saved(user_db_id, test_id, f"stage_{key}")

    if now_ts() > end_ts_ + 60:
        # ✅ Bo'lim vaqti (masalan ilova yopiq turganda) tugab bo'lgan —
        # botdagidek oxirgi saqlangan javoblar bilan yakunlaymiz.
        res = await _web_staged_finish_section(bot, tg_id, user_db_id, t, attempt["id"], key, saved)
        return {"kind": "expired", "section_name": base["name"], **res}

    _web_schedule_deadline(bot, tg_id, test_id, end_ts_)
    files = []
    if sec.get("file_id"):
        fk = sec.get("file_kind") or "document"
        files.append({"slot": f"s_{key}", "kind": fk, "name": "Audio" if fk == "audio" else
                      ("Video" if fk == "video" else "Savollar")})
    if sec.get("question_file_id"):
        files.append({"slot": f"sq_{key}", "kind": "document", "name": "Savollar"})
    payload = {**base, "end_ts": end_ts_, "server_now": now_ts(), "tools": _web_tools(et, key), "files": files}
    if stype in (STAGE_TYPE_MCQ, STAGE_TYPE_MCQ_BAND):
        skey = normalize_answers(sec.get("answers") or "")
        qcount = int(sec.get("qcount") or 0)
        payload.update(_web_question_payload(skey, qcount, saved))
        payload["layout"] = [{"name": "", "from": 1, "to": qcount}]
    else:
        payload.update({"prompt": sec.get("prompt") or "", "text": saved.get("text") or "",
                        "photo_id": saved.get("photo_id"), "text_max": WEB_WRITING_TEXT_MAX,
                        "updated": int(saved.get("updated") or 0)})
    return payload


# ----- HTTP endpointlar -----
@_web_handler
async def webapp_api_me(request: web.Request):
    ctx = await _web_auth(request, need_registered=False)
    bot: Bot = request.app["bot"]
    out = {"ok": True, "registered": ctx["registered"], "full_name": ctx["full_name"],
           "bot_username": await _web_bot_username(bot), "academy": ACADEMY_NAME}
    if ctx["registered"]:
        try:
            await _web_require_subscribed(bot, ctx["tg_id"])
            out["subscribed"] = True
        except WebError as e:
            out["subscribed"] = False
            out["channels"] = e.extra.get("channels", [])
        out["active"] = await _web_active_tests(ctx["user_db_id"])
    return web.json_response(out)


async def _web_active_tests(user_db_id: int) -> list[dict]:
    """Bosh sahifadagi "Davom ettirilayotgan testlar" — o'quvchi boshlagan,
    vaqti hali tugamagan testlar va yakunlanmagan bosqichli imtihonlar."""
    out = []
    rows = await POOL.fetch("""
        SELECT t.code, t.title, s.start_ts + s.duration_sec AS end_ts
        FROM sessions s JOIN tests t ON t.id = s.test_id
        WHERE s.user_id=$1 AND s.mode IN ('free', 'paid') AND s.start_ts + s.duration_sec > $2
        ORDER BY end_ts
    """, user_db_id, now_ts())
    for r in rows:
        out.append({"code": r["code"], "title": r["title"], "end_ts": int(r["end_ts"]), "staged": False})
    staged = await POOL.fetch("""
        SELECT DISTINCT ON (a.test_id) t.code, t.title
        FROM staged_attempts a JOIN tests t ON t.id = a.test_id
        WHERE a.user_id=$1 AND a.status='in_progress'
        ORDER BY a.test_id, a.id DESC
    """, user_db_id)
    for r in staged:
        out.append({"code": r["code"], "title": r["title"], "end_ts": None, "staged": True})
    return out[:10]


@_web_handler
async def webapp_api_lookup(request: web.Request):
    ctx = await _web_auth(request)
    bot: Bot = request.app["bot"]
    await _web_require_subscribed(bot, ctx["tg_id"])
    body = await _web_body(request)
    code = normalize_code(str(body.get("code") or ""))
    if not is_valid_code(code):
        raise WebError("bad_code", "Kod formati noto'g'ri. Masalan: MS-001")
    t = await _web_test_row(code=code)
    if not t:
        raise WebError("not_found", "Bunday kodli test topilmadi. Kodni tekshirib qayta kiriting.", 404)
    access = await _web_access(ctx["user_db_id"], t)
    msg = WEB_STATE_MESSAGES.get(access["state"], "")
    if access["state"] == "payment_rejected" and access.get("reason"):
        msg += f" Sabab: {access['reason']}"
    return web.json_response({"ok": True, "test": _web_test_info(t), "access": access, "message": msg,
                              "bot_username": await _web_bot_username(bot), "server_now": now_ts()})


@_web_handler
async def webapp_api_start(request: web.Request):
    ctx = await _web_auth(request)
    bot: Bot = request.app["bot"]
    await _web_require_subscribed(bot, ctx["tg_id"])
    body = await _web_body(request)
    test_id = _web_int(body.get("test_id"), "test")
    async with _web_lock(bot, ctx["tg_id"]):
        t = await _web_test_row(test_id=test_id)
        if not t:
            raise WebError("not_found", "Test topilmadi.", 404)
        access = await _web_access(ctx["user_db_id"], t)
        st = access["state"]
        et = t["exam_type"] or "simple"
        if et in STAGED_EXAM_TYPES:
            if st == "staged_done" and body.get("retake") and access.get("can_retake"):
                last_no = await POOL.fetchval(
                    "SELECT COALESCE(MAX(attempt_no), 0) FROM staged_attempts WHERE user_id=$1 AND test_id=$2",
                    ctx["user_db_id"], test_id)
                await POOL.execute("""
                    INSERT INTO staged_attempts (user_id, test_id, status, section_idx, sections_json, attempt_no,
                                                 started_at, created_ts)
                    VALUES ($1, $2, 'in_progress', 0, '{}', $3, $4, $4)
                """, ctx["user_db_id"], test_id, int(last_no) + 1, now_ts())
            elif st == "ready":
                await POOL.execute("""
                    INSERT INTO staged_attempts (user_id, test_id, status, section_idx, sections_json, attempt_no,
                                                 started_at, created_ts)
                    VALUES ($1, $2, 'in_progress', 0, '{}', 1, $3, $3)
                """, ctx["user_db_id"], test_id, now_ts())
            elif st != "resume":
                raise WebError(st, WEB_STATE_MESSAGES.get(st, "Testni boshlab bo'lmaydi."), 409)
            attempt = await _web_staged_attempt(ctx["user_db_id"], test_id)
            return web.json_response({"ok": True, **_web_staged_overview(t, attempt)})

        if st not in ("ready", "resume"):
            raise WebError(st, WEB_STATE_MESSAGES.get(st, "Testni boshlab bo'lmaydi."), 409)
        payload = await _web_start_single(bot, ctx, t, access)
    return web.json_response({"ok": True, **payload})


@_web_handler
async def webapp_api_section(request: web.Request):
    ctx = await _web_auth(request)
    bot: Bot = request.app["bot"]
    await _web_require_subscribed(bot, ctx["tg_id"])
    body = await _web_body(request)
    test_id = _web_int(body.get("test_id"), "test")
    async with _web_lock(bot, ctx["tg_id"]):
        t = await _web_test_row(test_id=test_id)
        if not t or (t["exam_type"] or "") not in STAGED_EXAM_TYPES:
            raise WebError("not_found", "Bosqichli test topilmadi.", 404)
        attempt = await _web_staged_attempt(ctx["user_db_id"], test_id)
        if not attempt:
            raise WebError("no_attempt", "Faol urinish topilmadi — testni qaytadan oching.", 409)
        payload = await _web_staged_section_payload(bot, ctx, t, attempt)
    return web.json_response({"ok": True, **payload})


async def _web_resolve_mode(ctx: dict, t, section_key: str | None) -> tuple[str, str, int]:
    """Saqlash/topshirish uchun sessiya rejimi, javob kaliti va savollar soni."""
    et = t["exam_type"] or "simple"
    if et in STAGED_EXAM_TYPES:
        a = await _web_staged_attempt(ctx["user_db_id"], t["id"])
        cfg = _load_skills_cfg(t["skills_cfg"])
        order = cfg.get("order") or []
        pos = int(a["section_idx"] or 0) if a else -1
        if not a or pos >= len(order) or order[pos] != section_key:
            raise WebError("stale", "Bu bo'lim allaqachon yakunlangan yoki eskirgan.", 409)
        sec = cfg["sections"][section_key]
        return f"stage_{section_key}", normalize_answers(sec.get("answers") or ""), int(sec.get("qcount") or 0)
    mode = "free" if int(t["is_free"] or 0) == 1 else "paid"
    return mode, normalize_answers(t["answers"] or ""), int(t["questions_count"] or 0)


@_web_handler
async def webapp_api_save(request: web.Request):
    ctx = await _web_auth(request)
    bot: Bot = request.app["bot"]
    body = await _web_body(request)
    test_id = _web_int(body.get("test_id"), "test")
    section_key = body.get("section_key") or None
    async with _web_lock(bot, ctx["tg_id"]):
        t = await _web_test_row(test_id=test_id)
        if not t:
            raise WebError("not_found", "Test topilmadi.", 404)
        mode, key, qcount = await _web_resolve_mode(ctx, t, section_key)
        sess = await get_session(ctx["user_db_id"], test_id, mode)
        if not sess:
            raise WebError("no_session", "Test sessiyasi yakunlangan.", 409)
        end_ts_ = int(sess["start_ts"]) + int(sess["duration_sec"])
        if now_ts() > end_ts_ + 60:
            raise WebError("time_over", "Vaqt tugagan.", 409)
        saved = await _web_load_saved(ctx["user_db_id"], test_id, mode)
        allowed = set(saved.get("uploads") or [])
        payload = {"uploads": list(allowed)}
        if mode.startswith("stage_") and not qcount:
            payload["text"] = sanitize_plain_text(str(body.get("text") or ""), max_len=WEB_WRITING_TEXT_MAX)
            pid = body.get("photo_id")
            payload["photo_id"] = pid if isinstance(pid, str) and pid in allowed else None
        else:
            answers, open_answers = _web_clean_answers(body.get("answers"), body.get("open_answers"), key, qcount,
                                                       allowed)
            payload.update({"answers": answers, "open_answers": open_answers,
                            "flags": _web_clean_flags(body.get("flags"), qcount),
                            "current_q": max(1, min(qcount or 1, _web_int(body.get("current_q") or 1, "savol")))})
        await _web_store_saved(ctx["user_db_id"], test_id, mode, payload)
    return web.json_response({"ok": True, "server_now": now_ts(), "end_ts": end_ts_})


@_web_handler
async def webapp_api_submit(request: web.Request):
    ctx = await _web_auth(request)
    bot: Bot = request.app["bot"]
    body = await _web_body(request)
    test_id = _web_int(body.get("test_id"), "test")
    async with _web_lock(bot, ctx["tg_id"]):
        t = await _web_test_row(test_id=test_id)
        if not t:
            raise WebError("not_found", "Test topilmadi.", 404)
        if (t["exam_type"] or "") in STAGED_EXAM_TYPES:
            raise WebError("bad_request", "Bosqichli test bo'limlarini alohida topshiring.")
        mode, key, qcount = await _web_resolve_mode(ctx, t, None)
        if not await get_session(ctx["user_db_id"], test_id, mode):
            # vaqt tugab server allaqachon topshirgan bo'lishi mumkin — natijani ko'rsatamiz
            view = await _web_result_view(ctx["user_db_id"], test_id)
            if view.get("state") != "none":
                return web.json_response({"ok": True, "result": view})
            raise WebError("no_session", "Test sessiyasi topilmadi — testni qaytadan oching.", 409)
        saved = await _web_load_saved(ctx["user_db_id"], test_id, mode)
        answers, open_answers = _web_clean_answers(body.get("answers"), body.get("open_answers"), key, qcount,
                                                   set(saved.get("uploads") or []))
        user_ans = "".join(answers.get(str(i), "*") for i in range(1, qcount + 1))
        _web_cancel_deadline(ctx["tg_id"], test_id)
        await process_test_submission(ctx["tg_id"], bot, _DetachedState({"test_id": test_id, "mode": mode}),
                                      user_ans, open_answers)
        view = await _web_result_view(ctx["user_db_id"], test_id)
        if view.get("state") == "none":
            raise WebError("time_over", "Vaqt tugagani uchun javoblar qabul qilinmadi.", 409)
    return web.json_response({"ok": True, "result": view})


@_web_handler
async def webapp_api_section_submit(request: web.Request):
    ctx = await _web_auth(request)
    bot: Bot = request.app["bot"]
    body = await _web_body(request)
    test_id = _web_int(body.get("test_id"), "test")
    section_key = str(body.get("section_key") or "")
    async with _web_lock(bot, ctx["tg_id"]):
        t = await _web_test_row(test_id=test_id)
        if not t or (t["exam_type"] or "") not in STAGED_EXAM_TYPES:
            raise WebError("not_found", "Bosqichli test topilmadi.", 404)
        await _web_resolve_mode(ctx, t, section_key)   # eskirgan bo'limni rad etadi
        attempt = await _web_staged_attempt(ctx["user_db_id"], test_id)
        if not await get_session(ctx["user_db_id"], test_id, f"stage_{section_key}"):
            raise WebError("no_session", "Bo'lim sessiyasi topilmadi — imtihonni qaytadan oching.", 409)
        res = await _web_staged_finish_section(bot, ctx["tg_id"], ctx["user_db_id"], t, attempt["id"],
                                               section_key, body)
    return web.json_response({"ok": True, **res})


@_web_handler
async def webapp_api_handoff(request: web.Request):
    """🎙 Speaking (ovozli) bo'limi — Telegram'ning o'z ovozli xabari eng
    qulay va ishonchli usul, shuning uchun shu bo'limdan boshlab imtihon
    chatda davom etadi (bot keyingi qadamlarni o'zi yuboradi)."""
    ctx = await _web_auth(request)
    bot: Bot = request.app["bot"]
    body = await _web_body(request)
    test_id = _web_int(body.get("test_id"), "test")
    async with _web_lock(bot, ctx["tg_id"]):
        t = await _web_test_row(test_id=test_id)
        attempt = await _web_staged_attempt(ctx["user_db_id"], test_id) if t else None
        if not t or not attempt:
            raise WebError("no_attempt", "Faol urinish topilmadi.", 409)
        cfg = _load_skills_cfg(t["skills_cfg"])
        order = cfg.get("order") or []
        fsm = _web_fsm(bot, ctx["tg_id"])
        if fsm is None:
            raise WebError("server", "Bot holati tayyor emas.", 500)
        _web_cancel_deadline(ctx["tg_id"], test_id)
        await fsm.clear()
        await fsm.set_state(StagedSolvingState.in_section)
        await fsm.update_data(staged_attempt_id=attempt["id"], staged_test_id=test_id, staged_order=order,
                              staged_pos=int(attempt["section_idx"] or 0))
        await _send_stage_section(ctx["tg_id"], bot, fsm)
    return web.json_response({"ok": True, "bot_username": await _web_bot_username(bot)})


@_web_handler
async def webapp_api_upload_photo(request: web.Request):
    ctx = await _web_auth(request)
    bot: Bot = request.app["bot"]
    if request.content_length and request.content_length > WEB_MAX_PHOTO_BYTES + 64 * 1024:
        raise WebError("too_big", "Rasm juda katta (10 MB gacha).", 413)
    reader = await request.multipart()
    fields, data, ctype = {}, None, ""
    while True:
        part = await reader.next()
        if part is None:
            break
        if part.name == "file":
            ctype = (part.headers.get("Content-Type") or "").lower()
            data = bytearray()
            while True:
                chunk = await part.read_chunk(64 * 1024)
                if not chunk:
                    break
                data.extend(chunk)
                if len(data) > WEB_MAX_PHOTO_BYTES:
                    raise WebError("too_big", "Rasm juda katta (10 MB gacha).", 413)
        else:
            fields[part.name] = (await part.text())[:64]
    if not data or not ctype.startswith("image/"):
        raise WebError("bad_file", "Faqat rasm (JPG/PNG) yuklash mumkin.")
    test_id = _web_int(fields.get("test_id"), "test")
    section_key = fields.get("section_key") or None
    q = fields.get("q") or ""
    async with _web_lock(bot, ctx["tg_id"]):
        t = await _web_test_row(test_id=test_id)
        if not t:
            raise WebError("not_found", "Test topilmadi.", 404)
        mode, _key, _qc = await _web_resolve_mode(ctx, t, section_key)
        if not await get_session(ctx["user_db_id"], test_id, mode):
            raise WebError("no_session", "Test sessiyasi yakunlangan.", 409)
        caption = f"📎 {q}-savol uchun javob (Mini App)" if q.isdigit() else "📎 Javob rasmi (Mini App)"
        try:
            sent = await bot.send_photo(ctx["tg_id"], BufferedInputFile(bytes(data), filename="javob.jpg"),
                                        caption=caption)
        except TelegramBadRequest:
            raise WebError("bad_file", "Rasmni qabul qilib bo'lmadi — boshqa rasm tanlang.")
        photo_id = sent.photo[-1].file_id
        saved = await _web_load_saved(ctx["user_db_id"], test_id, mode)
        saved["uploads"] = (list(saved.get("uploads") or []) + [photo_id])[-50:]
        await _web_store_saved(ctx["user_db_id"], test_id, mode, saved)
    return web.json_response({"ok": True, "photo_id": photo_id})


@_web_handler
async def webapp_api_file(request: web.Request):
    ctx = await _web_auth(request)
    bot: Bot = request.app["bot"]
    test_id = _web_int(request.match_info["test_id"], "test")
    slot = request.match_info["slot"]
    t = await _web_test_row(test_id=test_id)
    if not t:
        raise WebError("not_found", "Test topilmadi.", 404)
    file_id, kind = None, "document"
    if slot == "main":
        mode = "free" if int(t["is_free"] or 0) == 1 else "paid"
        has_session = await get_session(ctx["user_db_id"], test_id, mode)
        has_result = await POOL.fetchval("SELECT 1 FROM results WHERE user_id=$1 AND test_id=$2",
                                         ctx["user_db_id"], test_id)
        if not (has_session or has_result):
            raise WebError("forbidden", "Fayl faqat test boshlangandan keyin ochiladi.", 403)
        file_id = t["file_id"]
    elif slot.startswith(("s_", "sq_")):
        key = slot.split("_", 1)[1]
        cfg = _load_skills_cfg(t["skills_cfg"])
        sec = cfg.get("sections", {}).get(key)
        if not sec or not await get_session(ctx["user_db_id"], test_id, f"stage_{key}"):
            raise WebError("forbidden", "Bo'lim fayli faqat shu bo'lim davomida ochiladi.", 403)
        if slot.startswith("sq_"):
            file_id = sec.get("question_file_id")
        else:
            file_id, kind = sec.get("file_id"), sec.get("file_kind") or "document"
    if not file_id:
        raise WebError("not_found", "Fayl topilmadi.", 404)
    path = await _web_fetch_file(bot, file_id)
    with open(path, "rb") as fh:
        head = fh.read(16)
    return web.FileResponse(path, headers={"Content-Type": _web_sniff_type(head, kind),
                                           "Cache-Control": "private, max-age=3600"})


@_web_handler
async def webapp_api_result(request: web.Request):
    ctx = await _web_auth(request)
    test_id = _web_int(request.match_info["test_id"], "test")
    return web.json_response({"ok": True, "result": await _web_result_view(ctx["user_db_id"], test_id)})


async def webapp_index(request: web.Request):
    return web.Response(text=WEBAPP_HTML, content_type="text/html",
                        headers={"Cache-Control": "no-cache"})


async def webapp_api_profile(request: web.Request):
    urow = await _webapp_auth_user_row(request)
    if not urow:
        return web.json_response({"error": "unauthorized"}, status=401)

    user_db_id = urow["id"]
    summary = await POOL.fetchrow(
        "SELECT COUNT(*) AS cnt, AVG(percent) AS avgp, MAX(percent) AS maxp FROM results WHERE user_id=$1",
        user_db_id
    )
    results = await POOL.fetch("""
        SELECT r.id, r.test_id, t.code, t.title, t.exam_type, r.score, r.total_questions, r.percent, r.grade, r.date,
               r.certificate_sent
        FROM results r JOIN tests t ON t.id=r.test_id
        WHERE r.user_id=$1
        ORDER BY r.id DESC
        LIMIT 50
    """, user_db_id)

    return web.json_response({
        "full_name": urow["full_name"],
        "summary": {
            "count": summary["cnt"] or 0,
            "avg_percent": float(summary["avgp"] or 0),
            "best_percent": float(summary["maxp"] or 0),
        },
        "results": [
            {
                "id": r["id"], "test_id": r["test_id"], "code": r["code"], "title": r["title"],
                "exam_label": WEB_EXAM_LABELS.get(r["exam_type"] or "simple", ""),
                "score": r["score"], "total": r["total_questions"],
                "percent": float(r["percent"] or 0), "grade": r["grade"],
                "date": r["date"], "has_certificate": int(r["certificate_sent"] or 0) == 1,
            }
            for r in results
        ],
    })


async def webapp_api_certificate(request: web.Request):
    urow = await _webapp_auth_user_row(request)
    if not urow:
        return web.json_response({"error": "unauthorized"}, status=401)

    try:
        result_id = int(request.match_info["result_id"])
    except ValueError:
        return web.json_response({"error": "bad_request"}, status=400)

    row = await POOL.fetchrow("""
        SELECT r.user_id, r.score, r.total_questions, r.percent, r.rasch_percent, r.grade, r.date,
               r.certificate_sent,
               t.code, t.title, t.subject, t.exam_type, t.cert_academy_name, t.cert_signee_name,
               t.cert_template, t.creator_id
        FROM results r JOIN tests t ON t.id=r.test_id
        WHERE r.id=$1
    """, result_id)

    if not row or row["user_id"] != urow["id"]:
        return web.json_response({"error": "not_found"}, status=404)
    # ✅ Faqat haqiqatan berilgan sertifikatni yuklab olish mumkin — avval bu
    # manzil orqali istalgan (masalan 0% li) natija uchun ham sertifikat
    # yasab olish mumkin edi.
    if int(row["certificate_sent"] or 0) != 1:
        return web.json_response({"error": "no_certificate"}, status=404)

    cert_academy, cert_signee, cert_tpl = await _resolve_cert_settings(
        row["cert_academy_name"], row["cert_signee_name"], row["cert_template"], row["creator_id"])

    tmp_dir = tempfile.gettempdir()
    pdf_path = os.path.join(tmp_dir, f"webcert_{result_id}.pdf")

    await asyncio.to_thread(
        make_certificate_pdf,
        full_name=urow["full_name"] or "Noma'lum",
        subject=row["subject"] or "Noma'lum",
        test_code=row["code"],
        test_title=row["title"],
        exam_date_str=row["date"] or now_str_local(),
        percent=float((row["rasch_percent"] if (row["exam_type"] or "") == "rasch" else row["percent"]) or 0.0),
        correct_count=int(row["score"] or 0),
        total_questions=int(row["total_questions"] or 0),
        grade=row["grade"],
        out_path=pdf_path,
        academy_name=cert_academy,
        signee_name=cert_signee,
        template=cert_tpl,
    )

    return web.FileResponse(pdf_path, headers={
        "Content-Disposition": f'attachment; filename="certificate_{row["code"]}.pdf"'
    })


def build_webapp_application(bot: Bot | None = None) -> web.Application:
    app = web.Application(client_max_size=WEB_MAX_PHOTO_BYTES + 1024 * 1024)
    app["bot"] = bot
    app.router.add_get("/webapp", webapp_index)
    app.router.add_get("/", webapp_index)   # BotFather'da manzil "/webapp"siz kiritilsa ham ochiladi
    app.router.add_get("/api/me", webapp_api_me)
    app.router.add_get("/api/profile", webapp_api_profile)
    app.router.add_get("/api/certificate/{result_id}", webapp_api_certificate)
    app.router.add_post("/api/test/lookup", webapp_api_lookup)
    app.router.add_post("/api/test/start", webapp_api_start)
    app.router.add_post("/api/test/save", webapp_api_save)
    app.router.add_post("/api/test/submit", webapp_api_submit)
    app.router.add_post("/api/staged/section", webapp_api_section)
    app.router.add_post("/api/staged/submit", webapp_api_section_submit)
    app.router.add_post("/api/staged/handoff", webapp_api_handoff)
    app.router.add_post("/api/upload_photo", webapp_api_upload_photo)
    app.router.add_get("/api/file/{test_id}/{slot}", webapp_api_file)
    app.router.add_get("/api/result/{test_id}", webapp_api_result)
    return app


async def run_webapp_server(bot: Bot | None = None):
    app = build_webapp_application(bot)
    runner = web.AppRunner(app)
    await runner.setup()
    site = web.TCPSite(runner, WEBAPP_HOST, WEBAPP_PORT)
    await site.start()
    logging.info(f"Mini App server ishga tushdi: http://{WEBAPP_HOST}:{WEBAPP_PORT}/webapp")


# --- MINI APP HTML (avtomatik joylangan; tahrirlash uchun scratchpad/miniapp.html) ---
WEBAPP_HTML = r"""<!DOCTYPE html>
<html lang="uz">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, viewport-fit=cover">
<title>Test yechish</title>
<script src="https://telegram.org/js/telegram-web-app.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js"></script>
<style>
:root{
  --page: var(--tg-theme-secondary-bg-color, #eef1f5);
  --card: var(--tg-theme-section-bg-color, var(--tg-theme-bg-color, #ffffff));
  --text: var(--tg-theme-text-color, #15202b);
  --hint: var(--tg-theme-hint-color, #6b7680);
  --accent: var(--tg-theme-button-color, #2f6fed);
  --accent-text: var(--tg-theme-button-text-color, #ffffff);
  --link: var(--tg-theme-link-color, #2f6fed);
  --line: rgba(127,127,127,.22);
  --soft: rgba(127,127,127,.10);
  --ok: #1f9d55; --bad: #d64545; --warn: #e08a00; --flag: #f08c00;
  --radius: 14px;
}
*{box-sizing:border-box;-webkit-tap-highlight-color:transparent}
html,body{margin:0;padding:0}
body{background:var(--page);color:var(--text);font:16px/1.45 -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;min-height:100vh;-webkit-font-smoothing:antialiased}
button{font:inherit;color:inherit;cursor:pointer}
input,textarea{font:inherit;color:var(--text)}
.hidden{display:none!important}
.screen{max-width:760px;margin:0 auto;padding:16px 16px 96px}
h1{font-size:22px;margin:4px 0 4px}
h2{font-size:18px;margin:0 0 8px}
.muted{color:var(--hint)}
.small{font-size:13px}
.card{background:var(--card);border-radius:var(--radius);padding:16px;margin:12px 0;box-shadow:0 1px 2px rgba(0,0,0,.04)}
.btn{display:inline-flex;align-items:center;justify-content:center;gap:6px;border:0;border-radius:12px;padding:12px 16px;background:var(--accent);color:var(--accent-text);font-weight:600;min-height:46px}
.btn.block{display:flex;width:100%}
.btn.ghost{background:var(--soft);color:var(--text)}
.btn.danger{background:var(--bad);color:#fff}
.btn:disabled{opacity:.55}
.btn + .btn{margin-top:8px}
.row-btns{display:flex;gap:8px}.row-btns .btn{flex:1;margin-top:0}
.input{width:100%;border:1.5px solid var(--line);background:var(--page);border-radius:12px;padding:13px 14px;font-size:18px;letter-spacing:.06em;text-transform:uppercase;outline:none}
.input:focus{border-color:var(--accent)}
.input::placeholder{text-transform:none;letter-spacing:0;color:var(--hint);opacity:.8}
.resume-item{display:flex;align-items:center;gap:12px;padding:12px;border-radius:12px;background:var(--soft);margin-top:8px;cursor:pointer;border:0;width:100%;text-align:left}
.resume-item b{display:block}
.resume-item .go{margin-left:auto;font-weight:700;color:var(--accent);white-space:nowrap}
.chips{display:flex;flex-wrap:wrap;gap:6px;margin:8px 0}
.chip{background:var(--soft);border-radius:999px;padding:4px 10px;font-size:13px;white-space:nowrap}
.chip.acc{background:var(--accent);color:var(--accent-text)}
.status{border-radius:12px;padding:12px 14px;margin:12px 0;font-size:14px;background:var(--soft)}
.status.ok{background:rgba(31,157,85,.12)}.status.bad{background:rgba(214,69,69,.12)}.status.warn{background:rgba(224,138,0,.13)}
.feat{display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-top:4px}
.feat div{background:var(--soft);border-radius:12px;padding:10px;font-size:13px}
.tabs{position:fixed;left:0;right:0;bottom:0;display:flex;background:var(--card);border-top:1px solid var(--line);padding-bottom:env(safe-area-inset-bottom);z-index:20}
.tabs button{flex:1;border:0;background:none;padding:10px 0 12px;font-size:12px;color:var(--hint)}
.tabs button b{display:block;font-size:20px;line-height:1.2}
.tabs button.on{color:var(--accent)}
.list-item{display:flex;justify-content:space-between;align-items:center;gap:8px;padding:12px 0;border-bottom:1px solid var(--line)}
.list-item:last-child{border-bottom:0}
.badge{background:var(--accent);color:var(--accent-text);border-radius:8px;padding:3px 8px;font-size:13px;font-weight:600;white-space:nowrap}
.stats{display:grid;grid-template-columns:repeat(3,1fr);gap:8px}
.stat{background:var(--soft);border-radius:12px;padding:10px 6px;text-align:center}
.stat b{display:block;font-size:19px}.stat span{font-size:12px;color:var(--hint)}
.sec-item{display:flex;gap:12px;align-items:center;padding:12px;border-radius:12px;background:var(--soft);margin-bottom:8px}
.sec-item .ico{font-size:22px;width:30px;text-align:center}
.sec-item.current{outline:2px solid var(--accent)}
.sec-item.locked{opacity:.55}
/* ---- solve ---- */
.solve{position:fixed;inset:0;display:flex;flex-direction:column;background:var(--page);z-index:10}
.sbar{display:flex;align-items:center;gap:6px;padding:8px 10px;background:var(--card);border-bottom:1px solid var(--line)}
.timer{font-variant-numeric:tabular-nums;font-weight:700;font-size:17px;padding:5px 10px;border-radius:10px;background:var(--soft);min-width:84px;text-align:center}
.timer.warn{background:rgba(224,138,0,.18);color:var(--warn)}
.timer.crit{background:rgba(214,69,69,.18);color:var(--bad);animation:pulse 1s infinite}
@keyframes pulse{50%{opacity:.55}}
.prog{flex:1;min-width:0}
.prog .t{font-size:12px;color:var(--hint);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.prog .bar{height:5px;border-radius:5px;background:var(--soft);overflow:hidden;margin-top:3px}
.prog .bar i{display:block;height:100%;background:var(--accent);width:0;transition:width .25s}
.icon-btn{border:0;background:var(--soft);border-radius:10px;width:38px;height:38px;font-size:18px;display:inline-flex;align-items:center;justify-content:center;flex:none}
.icon-btn.on{background:var(--accent);color:var(--accent-text)}
.savest{font-size:12px;width:18px;text-align:center;flex:none}
.seg{display:flex;gap:6px;padding:8px 10px;background:var(--card);border-bottom:1px solid var(--line)}
.seg button{flex:1;border:0;border-radius:10px;padding:8px;background:var(--soft);font-weight:600;font-size:14px}
.seg button.on{background:var(--accent);color:var(--accent-text)}
.panes{flex:1;display:flex;min-height:0}
.pane{flex:1;overflow:auto;min-width:0;-webkit-overflow-scrolling:touch}
@media (max-width:859px){.panes .pane{display:none}.panes .pane.active{display:block}}
@media (min-width:860px){.seg{display:none}.pane.pdfpane{flex:1.3;border-right:1px solid var(--line)}}
.sfoot{display:flex;gap:8px;padding:10px 12px calc(10px + env(safe-area-inset-bottom));background:var(--card);border-top:1px solid var(--line)}
.sfoot .btn{flex:1;margin:0}
.pane-head{position:sticky;top:0;z-index:3;display:flex;align-items:center;justify-content:space-between;gap:8px;padding:0 12px;height:48px;background:var(--page);border-bottom:1px solid var(--line);font-size:13px;color:var(--hint)}
.pdfwrap{padding:8px;display:flex;flex-direction:column;align-items:flex-start;gap:8px}
.pdfpage{background:#fff;box-shadow:0 1px 3px rgba(0,0,0,.18);margin:0 auto;flex:none}
.pdfpage canvas{display:block;width:100%;height:auto}
.media{padding:10px 12px}.media audio,.media video{width:100%}
/* bubble sheet */
.block-h{position:sticky;top:48px;z-index:2;background:var(--page);padding:10px 14px 8px;font-weight:700;font-size:15px;display:flex;justify-content:space-between;align-items:baseline;border-bottom:1px solid var(--line)}
.block-h span{font-weight:500;color:var(--hint);font-size:12px}
.qrow{display:flex;align-items:center;gap:10px;padding:6px 10px 6px 12px;border-bottom:1px solid var(--line);background:var(--card)}
.qrow.alt{background-image:linear-gradient(rgba(127,127,127,.045),rgba(127,127,127,.045))}
.qrow.cur{box-shadow:inset 4px 0 0 var(--accent)}
.qn{width:30px;flex:none;text-align:right;font-weight:700;font-size:15px;color:var(--hint);background:none;border:0;padding:0;font-variant-numeric:tabular-nums}
.qrow.done .qn{color:var(--accent)}
.bubs{display:flex;gap:10px;flex:1}
.bub{width:40px;height:40px;border-radius:50%;border:2px solid rgba(127,127,127,.55);background:transparent;font-weight:700;font-size:15px;padding:0;transition:transform .08s,background .12s}
.bubs.many{gap:6px}.bubs.many .bub{width:34px;height:34px;font-size:14px}
.bub:active{transform:scale(.88)}
.bub.on{background:var(--accent);border-color:var(--accent);color:var(--accent-text);box-shadow:0 2px 6px rgba(0,0,0,.18)}
.flagb{border:0;background:none;font-size:18px;opacity:.3;width:36px;height:36px;border-radius:50%;flex:none;padding:0}
.flagb.on{opacity:1;color:var(--flag);background:rgba(240,140,0,.12)}
.openbtn{flex:1;border:1.5px dashed rgba(127,127,127,.6);background:transparent;border-radius:10px;padding:9px 12px;text-align:left;font-size:14px;min-height:40px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.openbtn.has{border-style:solid;border-color:var(--accent);background:rgba(47,111,237,.08);background:color-mix(in srgb,var(--accent) 10%,transparent)}
/* card mode */
.qcard{padding:14px}
.qcard .ch{display:flex;justify-content:space-between;align-items:center;margin-bottom:12px}
.qcard .ch b{font-size:20px}
.opt{display:flex;align-items:center;gap:12px;padding:12px;border:1.5px solid var(--line);border-radius:12px;margin-bottom:8px;background:var(--card)}
.opt.on{border-color:var(--accent);background:rgba(47,111,237,.08);background:color-mix(in srgb,var(--accent) 10%,transparent)}
.opt{min-height:58px;cursor:pointer;font-size:17px}
.opt.elim{opacity:.4}.opt.elim .ol{text-decoration:line-through}
.opt .ol{width:34px;height:34px;border-radius:50%;border:1.7px solid var(--hint);display:flex;align-items:center;justify-content:center;font-weight:700;flex:none}
.opt.on .ol{background:var(--accent);border-color:var(--accent);color:var(--accent-text)}
.opt .lbl{flex:1}
.opt .strike{border:0;background:var(--soft);border-radius:8px;width:32px;height:32px;flex:none}
.cnav{display:flex;gap:8px;margin-top:12px}.cnav .btn{flex:1;margin:0}
textarea.ta{width:100%;min-height:120px;border:1.5px solid var(--line);border-radius:12px;padding:12px;background:var(--page);resize:vertical;outline:none;font-size:16px}
textarea.ta:focus{border-color:var(--accent)}
textarea.ta.big{min-height:45vh}
.thumb{max-width:140px;max-height:140px;border-radius:10px;display:block;margin-top:8px}
/* modal / sheet */
.mback{position:fixed;inset:0;background:rgba(0,0,0,.45);display:flex;align-items:flex-end;justify-content:center;z-index:60;animation:fade .15s}
@keyframes fade{from{opacity:0}}
.msheet{background:var(--card);width:100%;max-width:560px;border-radius:18px 18px 0 0;padding:18px 16px calc(16px + env(safe-area-inset-bottom));max-height:88vh;overflow:auto}
@media (min-width:860px){.mback{align-items:center}.msheet{border-radius:18px}}
.navgrid{display:grid;grid-template-columns:repeat(auto-fill,minmax(42px,1fr));gap:6px;margin:12px 0}
.navgrid button{height:42px;border-radius:10px;border:1.5px solid var(--line);background:transparent;font-weight:700;position:relative}
.navgrid button.ans{background:var(--accent);color:var(--accent-text);border-color:var(--accent)}
.navgrid button.flag{box-shadow:0 0 0 2px var(--flag)}
.navgrid button.ok{background:rgba(31,157,85,.16);border-color:var(--ok);color:var(--ok)}
.navgrid button.bad{background:rgba(214,69,69,.14);border-color:var(--bad);color:var(--bad)}
.navgrid button.pend{border-style:dashed}
.legend{display:flex;flex-wrap:wrap;gap:10px;font-size:12px;color:var(--hint)}
.legend i{display:inline-block;width:12px;height:12px;border-radius:4px;vertical-align:-2px;margin-right:4px;border:1.5px solid var(--line)}
.toast{position:fixed;left:50%;bottom:calc(84px + env(safe-area-inset-bottom));transform:translateX(-50%);background:rgba(20,24,30,.92);color:#fff;padding:10px 16px;border-radius:12px;font-size:14px;z-index:80;max-width:92vw;text-align:center;animation:fade .2s}
.overlay{position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:90;display:flex;align-items:center;justify-content:center;color:#fff;flex-direction:column;gap:14px;text-align:center;padding:20px}
.spin{width:42px;height:42px;border-radius:50%;border:4px solid rgba(255,255,255,.3);border-top-color:#fff;animation:sp 1s linear infinite}
.spin.dark{border-color:var(--soft);border-top-color:var(--accent)}
@keyframes sp{to{transform:rotate(360deg)}}
.center{display:flex;flex-direction:column;align-items:center;justify-content:center;min-height:60vh;text-align:center;gap:10px}
.ring{width:150px;height:150px;margin:6px auto}
.ring text{font-weight:800}
.kv{display:flex;justify-content:space-between;padding:8px 0;border-bottom:1px solid var(--line)}.kv:last-child{border-bottom:0}
table.tb{width:100%;border-collapse:collapse;font-size:14px}table.tb td,table.tb th{padding:7px 4px;border-bottom:1px solid var(--line);text-align:left}table.tb th{color:var(--hint);font-weight:600;font-size:12px}
/* calculator */
.calc{position:fixed;top:64px;right:12px;width:330px;max-width:calc(100vw - 24px);background:var(--card);border-radius:16px;box-shadow:0 8px 30px rgba(0,0,0,.28);z-index:50;overflow:hidden;border:1px solid var(--line)}
.calc .ch{display:flex;align-items:center;gap:6px;padding:8px 10px;background:var(--soft);cursor:move;touch-action:none;user-select:none}
.calc .ch .tb2{flex:1;display:flex;gap:4px}
.calc .ch .tb2 button{border:0;border-radius:8px;padding:5px 9px;background:transparent;font-size:13px;font-weight:600}
.calc .ch .tb2 button.on{background:var(--card)}
.calc .disp{padding:10px 12px}
.calc .expr{width:100%;border:0;background:transparent;font-size:20px;text-align:right;outline:none;font-variant-numeric:tabular-nums}
.calc .res{text-align:right;color:var(--hint);min-height:22px;font-size:16px;font-variant-numeric:tabular-nums}
.keys{display:grid;grid-template-columns:repeat(5,1fr);gap:5px;padding:0 10px 10px}
.keys button{border:0;border-radius:9px;height:40px;background:var(--soft);font-weight:600;font-size:15px}
.keys button.op{background:rgba(47,111,237,.13);background:color-mix(in srgb,var(--accent) 15%,transparent);color:var(--accent)}
.keys button.eq{background:var(--accent);color:var(--accent-text)}
.keys button.fn{font-size:13px}
.hist{max-height:70px;overflow:auto;padding:0 12px 6px;font-size:12px;color:var(--hint);text-align:right}
.graph{padding:8px 10px 10px}
.graph canvas{width:100%;height:230px;background:var(--page);border-radius:10px;touch-action:none;display:block}
.fnrow{display:flex;align-items:center;gap:6px;margin-bottom:6px}
.fnrow i{width:10px;height:10px;border-radius:50%;flex:none}
.fnrow input{flex:1;border:1.5px solid var(--line);border-radius:8px;padding:6px 8px;background:var(--page);outline:none;font-size:14px}
.ref h3{font-size:15px;margin:14px 0 6px}
.ref .f{display:flex;justify-content:space-between;gap:10px;padding:6px 0;border-bottom:1px solid var(--line);font-size:14px}
.ref .f b{font-family:"Times New Roman",serif;font-style:italic;font-size:16px;white-space:nowrap}
.words{font-size:12px;color:var(--hint);text-align:right;margin-top:4px}
.prompt{white-space:pre-wrap}
</style>
</head>
<body>
<div id="view"></div>
<nav class="tabs hidden" id="tabs">
  <button data-tab="home" class="on"><b>📝</b>Test yechish</button>
  <button data-tab="profile"><b>👤</b>Profil</button>
</nav>
<div id="calc" class="calc hidden"></div>

<script>
"use strict";
const tg = window.Telegram && window.Telegram.WebApp;
// ✅ Telegram ma'lumoti odatda telegram-web-app.js orqali olinadi; skript
// yuklanmasa yoki bo'sh qaytarsa — Telegram uni sahifa manzilining #...
// qismiga ham yozadi (tgWebAppData), shu yerdan o'qiymiz.
function readInitData() {
  if (tg && tg.initData) return tg.initData;
  try { const h = new URLSearchParams(location.hash.slice(1)).get("tgWebAppData"); if (h) return h; } catch (e) {}
  try { const p = JSON.parse(sessionStorage.getItem("__telegram__initParams") || "{}"); if (p.tgWebAppData) return p.tgWebAppData; } catch (e) {}
  return "";
}
const initData = readInitData();
function readUid() {
  if (tg && tg.initDataUnsafe && tg.initDataUnsafe.user && tg.initDataUnsafe.user.id) return tg.initDataUnsafe.user.id;
  try { return JSON.parse(new URLSearchParams(initData).get("user") || "{}").id || 0; } catch (e) { return 0; }
}
const S = { me: null, uid: readUid(),
            back: null, tick: null, st: null, tab: "home" };
const LETTERS = "ABCDEF";

/* ------------ yordamchilar ------------ */
function el(tag, attrs, ...kids) {
  const e = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs || {})) {
    if (v == null || v === false) continue;
    if (k === "class") e.className = v;
    else if (k.startsWith("on") && typeof v === "function") e.addEventListener(k.slice(2), v);
    else if (k === "html") e.innerHTML = v;
    else e.setAttribute(k, v === true ? "" : v);
  }
  for (const k of kids.flat()) {
    if (k == null || k === false) continue;
    e.append(k instanceof Node ? k : document.createTextNode(String(k)));
  }
  return e;
}
const $view = () => document.getElementById("view");
function clean(nodes) { return nodes.flat().filter(n => n != null && n !== false); }
function mount(...nodes) { const v = $view(); v.replaceChildren(...clean(nodes)); window.scrollTo(0, 0); }
function toast(msg, ms) {
  document.querySelectorAll(".toast").forEach(t => t.remove());
  const t = el("div", {class: "toast"}, msg); document.body.append(t);
  setTimeout(() => t.remove(), ms || 2600);
}
function haptic(kind) {
  try {
    if (!tg || !tg.HapticFeedback) return;
    if (kind === "sel") tg.HapticFeedback.selectionChanged();
    else if (kind === "ok" || kind === "error" || kind === "warning") tg.HapticFeedback.notificationOccurred(kind === "ok" ? "success" : kind);
    else tg.HapticFeedback.impactOccurred(kind || "light");
  } catch (e) {}
}
function overlay(text) {
  const o = el("div", {class: "overlay"}, el("div", {class: "spin"}), el("div", {}, text));
  document.body.append(o); return () => o.remove();
}
function fmtTime(sec) {
  sec = Math.max(0, Math.floor(sec));
  const h = Math.floor(sec / 3600), m = Math.floor(sec % 3600 / 60), s = sec % 60;
  const p = n => String(n).padStart(2, "0");
  return h ? `${h}:${p(m)}:${p(s)}` : `${p(m)}:${p(s)}`;
}
function plural(n, w) { return `${n} ta ${w}`; }
function lsGet(k) { try { return JSON.parse(localStorage.getItem(k) || "null"); } catch (e) { return null; } }
function lsSet(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) {} }
function lsDel(k) { try { localStorage.removeItem(k); } catch (e) {} }
function stopTick() { if (S.tick) { clearInterval(S.tick); S.tick = null; } }

async function api(path, opts) {
  opts = opts || {};
  const o = {method: opts.method || (opts.body || opts.form ? "POST" : "GET"), headers: {"X-Telegram-Init-Data": initData}};
  if (opts.body !== undefined) { o.headers["Content-Type"] = "application/json"; o.body = JSON.stringify(opts.body); }
  if (opts.form) o.body = opts.form;
  if (opts.keepalive) o.keepalive = true;
  let res;
  try { res = await fetch(path, o); }
  catch (e) { throw {network: true, message: "Internet aloqasi yo'q. Qayta urinib ko'ring."}; }
  let data = {};
  try { data = await res.json(); } catch (e) {}
  if (!res.ok || data.ok === false) throw Object.assign({status: res.status}, data, {message: data.message || "Xatolik yuz berdi."});
  return data;
}
function handleErr(e) {
  if (e && e.error === "not_subscribed") return renderGate("sub", e.channels);
  if (e && e.error === "not_registered") return renderGate("reg");
  if (e && e.error === "unauthorized") return renderGate("tg");
  toast((e && e.message) || "Xatolik yuz berdi.", 3500);
  haptic("error");
}
function openBot() {
  const u = S.me && S.me.bot_username;
  if (tg && u) { tg.openTelegramLink("https://t.me/" + u); setTimeout(() => tg.close(), 300); }
  else if (tg) tg.close();
}
function setBack(fn) {
  S.back = fn;
  if (!tg || !tg.BackButton) return;
  if (fn) tg.BackButton.show(); else tg.BackButton.hide();
}
let confirmOpen = false;
function confirmBox(title, text, okText, danger) {
  // tugma tez ikki marta bosilsa ikkinchi oyna ochilmaydi
  if (confirmOpen) return Promise.resolve(false);
  confirmOpen = true;
  return new Promise(resolve => {
    const close = v => { confirmOpen = false; back.remove(); resolve(v); };
    const back = el("div", {class: "mback", onclick: e => { if (e.target === back) close(false); }},
      el("div", {class: "msheet"},
        el("h2", {}, title), text ? el("p", {class: "muted"}, text) : null,
        el("button", {class: "btn block" + (danger ? " danger" : ""), onclick: () => close(true)}, okText || "Ha"),
        el("button", {class: "btn block ghost", onclick: () => close(false)}, "Bekor qilish")));
    document.body.append(back);
  });
}
function sheet(...kids) {
  const back = el("div", {class: "mback"});
  const box = el("div", {class: "msheet"}, ...kids);
  back.append(box);
  back.addEventListener("click", e => { if (e.target === back) back.remove(); });
  document.body.append(back);
  return back;
}
function showTabs(on) { document.getElementById("tabs").classList.toggle("hidden", !on); }

/* ------------ kirish ------------ */
// Bir martalik sozlash. Avval boot() ichida edi — "A'zo bo'ldim" bosilib boot()
// qayta chaqirilganda tugma tutqichlari ikki marta ulanardi (orqaga tugmasi
// tasdiqlash oynasini 2 marta ochardi).
function initOnce() {
  if (tg) {
    try { tg.ready(); tg.expand(); } catch (e) {}
    try { tg.BackButton.onClick(() => S.back && S.back()); } catch (e) {}
    try { tg.setHeaderColor("secondary_bg_color"); tg.setBackgroundColor("secondary_bg_color"); } catch (e) {}
    try { if (tg.disableVerticalSwipes) tg.disableVerticalSwipes(); } catch (e) {}
  }
  document.querySelectorAll("#tabs button").forEach(b => b.addEventListener("click", () => {
    const t = b.dataset.tab; if (t === "home") renderHome(); else renderProfile();
  }));
}
async function boot() {
  if (!initData) return renderGate("tg");
  mount(el("div", {class: "center"}, el("div", {class: "spin dark"}), el("div", {class: "muted"}, "Yuklanmoqda...")));
  try { S.me = await api("/api/me"); S.meTs = Date.now(); }
  catch (e) { return handleErr(e); }
  if (!S.me.registered) return renderGate("reg");
  if (!S.me.subscribed) return renderGate("sub", S.me.channels);
  const params = new URLSearchParams(location.search);
  const code = params.get("code") || (tg && tg.initDataUnsafe && tg.initDataUnsafe.start_param) || "";
  renderHome();
  if (code) lookup(code);
}

function renderGate(kind, channels) {
  stopTick(); showTabs(false); setBack(null);
  if (kind === "tg") {
    const diag = `${tg ? "sdk" : "nosdk"} · ${(tg && tg.platform) || "-"} · v${(tg && tg.version) || "-"} · ${initData ? "data" : "nodata"}`;
    return mount(el("div", {class: "screen center"}, el("div", {style: "font-size:48px"}, "📱"),
      el("h2", {}, "Ilovani Telegram ichida oching"),
      el("p", {class: "muted"}, initData
        ? "Telegram ma'lumotlari tasdiqlanmadi. Ilovani yopib, bot chatidagi «Test yechish» menyu tugmasi orqali qayta oching."
        : "Bu sahifa bot chatidagi «Test yechish» menyu tugmasi orqali ochiladi."),
      el("div", {class: "muted small", style: "opacity:.6"}, diag)));
  }
  if (kind === "reg") {
    return mount(el("div", {class: "screen center"}, el("div", {style: "font-size:48px"}, "📝"),
      el("h2", {}, "Avval ro'yxatdan o'ting"),
      el("p", {class: "muted"}, "Botga /start yuboring, ism va telefon raqamingizni kiriting. So'ng ilovani qayta oching."),
      el("button", {class: "btn", onclick: openBot}, "Botga o'tish")));
  }
  const box = el("div", {class: "card"});
  (channels || []).forEach(c => box.append(el("div", {class: "list-item"}, el("b", {}, c.title),
    c.url ? el("button", {class: "btn", onclick: () => tg ? tg.openTelegramLink(c.url) : window.open(c.url)}, "A'zo bo'lish") : null)));
  mount(el("div", {class: "screen"}, el("div", {class: "center", style: "min-height:auto;padding-top:20px"},
      el("div", {style: "font-size:48px"}, "🔒"), el("h2", {}, "Kanallarga a'zo bo'ling"),
      el("p", {class: "muted"}, "Testlardan foydalanish uchun quyidagi kanal/guruhlarga a'zo bo'ling.")), box,
    el("button", {class: "btn block", onclick: () => boot()}, "✅ A'zo bo'ldim — tekshirish")));
}

/* ------------ bosh sahifa: kod kiritish ------------ */
function renderHome() {
  stopTick(); closeCalc(); showTabs(true); setBack(null); setTab("home");
  if (tg) { try { tg.disableClosingConfirmation(); } catch (e) {} }
  const inp = el("input", {class: "input", id: "code", placeholder: "Masalan: MS-001", autocomplete: "off",
    autocapitalize: "characters", spellcheck: "false", maxlength: "20", enterkeyhint: "go"});
  inp.addEventListener("keydown", e => { if (e.key === "Enter") go(); });
  const go = () => { const v = inp.value.trim(); if (!v) { inp.focus(); return toast("Test kodini kiriting"); } lookup(v); };
  const name = (S.me && S.me.full_name) || "";
  const activeBox = el("div", {});
  const drawActive = (list) => {
    if (!list || !list.length) { activeBox.replaceChildren(); return; }
    activeBox.replaceChildren(el("div", {class: "card"}, el("h2", {}, "▶️ Davom ettirilayotgan testlar"),
      ...list.map(a => el("button", {class: "resume-item", onclick: () => lookup(a.code)},
        el("div", {}, el("b", {}, a.title || a.code),
          el("div", {class: "muted small"}, a.staged ? `${a.code} · bosqichli imtihon` :
            `${a.code} · ${Math.max(0, Math.round((a.end_ts - Date.now() / 1000) / 60))} daqiqa qoldi`)),
        el("span", {class: "go"}, "Davom ›")))));
  };
  drawActive(S.me && S.me.active);
  if (!S.meTs || Date.now() - S.meTs > 3000) {
    S.meTs = Date.now();
    api("/api/me").then(me => { S.me = Object.assign(S.me || {}, me); if (activeBox.isConnected) drawActive(me.active); }).catch(() => {});
  }
  mount(el("div", {class: "screen"},
    el("div", {class: "muted small"}, (S.me && S.me.academy) || ""),
    el("h1", {}, "Assalomu alaykum" + (name ? ", " + name.split(" ")[0] : "") + "! 👋"),
    el("p", {class: "muted"}, "Ustozingiz bergan test kodini kiriting — test turiga mos interaktiv oyna ochiladi."),
    el("div", {class: "card"},
      el("label", {class: "muted small", for: "code"}, "Test kodi"), el("div", {style: "height:6px"}), inp,
      el("div", {style: "height:12px"}), el("button", {class: "btn block", onclick: go}, "🔎 Testni ochish")),
    activeBox,
    el("div", {class: "card"}, el("h2", {}, "Imkoniyatlar"),
      el("div", {class: "feat"},
        el("div", {}, "🔵 DTM javoblar varaqasi — doirachani bosib belgilang"),
        el("div", {}, "🧮 SAT: kalkulyator, grafik va formulalar"),
        el("div", {}, "✍️ Ochiq savollarga matn yoki rasm"),
        el("div", {}, "💾 Javoblar avtomatik saqlanadi"),
        el("div", {}, "⏱ Taymer: vaqt tugasa avtomatik topshiriladi"),
        el("div", {}, "🎧 Listening audio va 📄 PDF shu yerda")))));
}
function setTab(t) {
  S.tab = t;
  document.querySelectorAll("#tabs button").forEach(b => b.classList.toggle("on", b.dataset.tab === t));
}

/* ------------ test kartasi ------------ */
async function lookup(code) {
  const done = overlay("Test qidirilmoqda...");
  let r;
  try { r = await api("/api/test/lookup", {body: {code}}); }
  catch (e) { done(); return handleErr(e); }
  done(); renderInfo(r);
}
const STATE_STYLE = {ready: "ok", resume: "ok", done: "", staged_done: "", not_started: "warn", payment_pending: "warn",
  pending_review: "warn", need_payment: "warn", payment_rejected: "bad", closed: "bad", hidden: "bad", expired: "bad", broken: "bad"};

function renderInfo(r) {
  stopTick(); showTabs(false); setBack(renderHome);
  const t = r.test, a = r.access, st = a.state;
  const chips = el("div", {class: "chips"},
    el("span", {class: "chip acc"}, t.exam_label), el("span", {class: "chip"}, t.is_free ? "🆓 Tekin" : `💰 ${t.price.toLocaleString("ru-RU")} so'm`),
    t.staged ? el("span", {class: "chip"}, `🧩 ${t.sections.length} bo'lim`) : el("span", {class: "chip"}, `🧮 ${t.qcount} savol`),
    el("span", {class: "chip"}, `⏳ ${t.duration} daqiqa`), t.subject ? el("span", {class: "chip"}, "📚 " + t.subject) : null);
  const statusBox = el("div", {class: "status " + (STATE_STYLE[st] || "")}, r.message);
  const body = [];
  if (t.staged) {
    const list = el("div", {});
    t.sections.forEach((s, i) => list.append(el("div", {class: "sec-item"}, el("div", {class: "ico"}, secIcon(s)),
      el("div", {style: "flex:1"}, el("b", {}, `${i + 1}. ${s.name}`),
        el("div", {class: "muted small"}, `${s.type_label}${s.qcount ? " · " + s.qcount + " savol" : ""} · ${s.duration_min} daqiqa`)))));
    body.push(el("div", {class: "card"}, el("h2", {}, "Bo'limlar"), list));
  }
  const btns = [];
  const start = async (retake) => {
    const txt = t.staged ? "Bo'limlar ketma-ket topshiriladi, har birining o'z vaqti bor." :
      `Vaqt ${t.duration} daqiqa. Boshlangach taymer to'xtamaydi — ilovani yopsangiz ham vaqt ketaveradi, javoblaringiz saqlanadi.`;
    if (st === "ready" || retake) { if (!await confirmBox("Testni boshlaysizmi?", txt, "🚀 Boshlash")) return; }
    startTest(t, retake);
  };
  if (st === "ready") btns.push(el("button", {class: "btn block", onclick: () => start(false)}, "🚀 Testni boshlash"));
  if (st === "resume") btns.push(el("button", {class: "btn block", onclick: () => start(false)}, "▶️ Davom ettirish"));
  if (st === "done" || st === "staged_done" || st === "pending_review") btns.push(el("button", {class: "btn block", onclick: () => showResult(t.id)}, "📊 Natijani ko'rish"));
  if (st === "staged_done" && a.can_retake) btns.push(el("button", {class: "btn block ghost", onclick: () => start(true)}, "🔄 Qaytadan yechish"));
  if (st === "need_payment" || st === "payment_rejected" || st === "payment_pending") btns.push(el("button", {class: "btn block", onclick: openBot}, "💳 Botga o'tish"));
  btns.push(el("button", {class: "btn block ghost", onclick: renderHome}, "⬅️ Boshqa kod"));
  if (st === "not_started" && a.left) {
    const cd = el("b", {}, fmtTime(a.left));
    statusBox.append(el("div", {style: "margin-top:6px"}, "Boshlanishiga: ", cd));
    const until = Date.now() / 1000 + a.left;
    S.tick = setInterval(() => {
      const left = until - Date.now() / 1000; cd.textContent = fmtTime(left);
      if (left <= 0) { stopTick(); lookup(t.code); }
    }, 1000);
  }
  mount(el("div", {class: "screen"},
    el("div", {class: "muted small"}, "Kod: " + t.code), el("h1", {}, t.title || "Test"), chips,
    t.scheduled && t.start_at ? el("div", {class: "muted small"}, "📅 Boshlanish: " + t.start_at) : null,
    statusBox, ...body, el("div", {class: "card"}, ...btns)));
}
function secIcon(s) {
  const byKey = {listening: "🎧", reading: "📖", reading_writing: "📖", math: "🧮", writing: "✍️", speaking: "🎙"};
  return byKey[s.key] || {mcq: "📝", mcq_band: "📖", open_text: "✍️", open_voice: "🎙"}[s.type] || "📝";
}

async function startTest(t, retake) {
  const done = overlay("Tayyorlanmoqda...");
  let p;
  try { p = await api("/api/test/start", {body: {test_id: t.id, retake: !!retake}}); }
  catch (e) { done(); return handleErr(e); }
  done();
  if (p.kind === "staged") renderOverview(p); else renderSolve(p, null);
}

/* ------------ bosqichli imtihon xaritasi ------------ */
function renderOverview(ov) {
  stopTick(); closeCalc(); showTabs(false); setBack(renderHome);
  if (tg) { try { tg.disableClosingConfirmation(); } catch (e) {} }
  const t = ov.test;
  const list = el("div", {});
  let current = null;
  ov.sections.forEach((s, i) => {
    if (s.status === "current") current = s;
    const mark = s.status === "done" ? "✅" : s.status === "current" ? "▶️" : "🔒";
    list.append(el("div", {class: "sec-item " + s.status}, el("div", {class: "ico"}, secIcon(s)),
      el("div", {style: "flex:1"}, el("b", {}, `${i + 1}. ${s.name}`),
        el("div", {class: "muted small"}, `${s.type_label}${s.qcount ? " · " + s.qcount + " savol" : ""} · ${s.duration_min} daqiqa`)),
      el("div", {style: "font-size:20px"}, mark)));
  });
  const actions = el("div", {class: "card"});
  if (current) {
    if (current.type === "open_voice") {
      actions.append(el("p", {}, el("b", {}, current.name), " bo'limi Telegram'ning ovozli xabari orqali topshiriladi — bu eng qulay va ishonchli usul."),
        el("p", {class: "muted small"}, "Tugmani bosganingizda bot chatda topshiriqni yuboradi va vaqt boshlanadi. Ovozli javobdan keyin qolgan bo'limlarni chatda yoki shu ilovada (kodni qayta kiritib) davom ettirishingiz mumkin."),
        el("button", {class: "btn block", onclick: () => handoff(t)}, "🎙 Botda ovozli javob berish"));
    } else {
      actions.append(el("button", {class: "btn block", onclick: async () => {
        if (!await confirmBox(`${current.name} bo'limini boshlaysizmi?`, `Vaqt: ${current.duration_min} daqiqa. Boshlangach taymer to'xtamaydi.`, "▶️ Boshlash")) return;
        openSection(t);
      }}, `▶️ ${current.name} bo'limini boshlash`));
    }
  }
  actions.append(el("button", {class: "btn block ghost", onclick: renderHome}, "🏠 Bosh sahifa"));
  mount(el("div", {class: "screen"}, el("div", {class: "muted small"}, "Kod: " + t.code), el("h1", {}, t.title),
    el("div", {class: "chips"}, el("span", {class: "chip acc"}, t.exam_label)),
    el("div", {class: "card"}, el("h2", {}, "Imtihon xaritasi"), list), actions));
}
async function openSection(t) {
  const done = overlay("Bo'lim ochilmoqda...");
  let p;
  try { p = await api("/api/staged/section", {body: {test_id: t.id}}); }
  catch (e) { done(); return handleErr(e); }
  done();
  if (p.kind === "expired") {
    toast(`⏱ «${p.section_name}» vaqti tugagan edi — saqlangan javoblar bilan topshirildi.`, 4000);
    return p.finished ? renderResult(p.result) : renderOverview(p.overview);
  }
  if (p.handoff) {
    try { return renderOverview(await api("/api/test/start", {body: {test_id: t.id}})); }
    catch (e) { return handleErr(e); }
  }
  if (p.type === "open_text") return renderWriting(p);
  renderSolve(p, p.key);
}
async function handoff(t) {
  const done = overlay("Bot chatiga yuborilmoqda...");
  try { await api("/api/staged/handoff", {body: {test_id: t.id}}); }
  catch (e) { done(); return handleErr(e); }
  done();
  mount(el("div", {class: "screen center"}, el("div", {style: "font-size:52px"}, "🎙"), el("h2", {}, "Topshiriq bot chatiga yuborildi"),
    el("p", {class: "muted"}, "Chatga qayting, topshiriqni o'qing va javobingizni ovozli xabar qilib yuboring. Vaqt boshlandi."),
    el("button", {class: "btn", onclick: openBot}, "💬 Chatga qaytish")));
  setBack(renderHome);
}

/* ------------ test yechish oynasi ------------ */
function lsKeyFor(st) { return `et:${S.uid}:${st.testId}:${st.sectionKey || "main"}:${st.endTs}`; }

function renderSolve(p, sectionKey) {
  stopTick(); showTabs(false);
  const st = {
    testId: p.test.id, sectionKey, test: p.test, name: p.name || p.test.title, qcount: p.qcount, nOpts: p.n_opts,
    open: new Set(p.open || []), answers: Object.assign({}, p.answers), openAnswers: Object.assign({}, p.open_answers),
    flags: new Set(p.flags || []), elim: {}, current: Math.min(Math.max(p.current_q || 1, 1), p.qcount),
    view: (p.tools && p.tools.default_view) || "sheet", endTs: p.end_ts, offset: p.server_now - Date.now() / 1000,
    layout: p.layout && p.layout.length ? p.layout : [{name: "", from: 1, to: p.qcount}], files: p.files || [], tools: p.tools || {},
    saveTimer: null, saving: false, again: false, submitted: false, pane: "ans", photos: {}, warned: {}
  };
  // qurilmada saqlangan (internet uzilgan paytdagi) javoblar serverdagidan yangiroq bo'lsa — ularni olamiz
  const local = lsGet(lsKeyFor(st));
  if (local && local.ts > (p.updated || 0)) {
    st.answers = local.answers || {}; st.openAnswers = local.openAnswers || {}; st.flags = new Set(local.flags || []);
    setTimeout(() => scheduleSave(st), 300);
  }
  S.st = st;
  if (tg) { try { tg.enableClosingConfirmation(); } catch (e) {} }
  setBack(async () => {
    if (await confirmBox("Testdan chiqasizmi?", "Javoblaringiz saqlangan, vaqt esa ketaveradi. Kodni qayta kiritib davom ettirishingiz mumkin.", "Chiqish")) {
      await flushSave(st); renderHome();
    }
  });

  const hasFiles = st.files.length > 0;
  const timer = el("div", {class: "timer"}, "--:--");
  const progT = el("div", {class: "t"}), progBar = el("i");
  const saveSt = el("span", {class: "savest", title: "Saqlash holati"}, "");
  const tools = el("div", {style: "display:flex;gap:6px"});
  if (st.tools.calculator) tools.append(el("button", {class: "icon-btn", title: "Kalkulyator", onclick: () => toggleCalc()}, "🧮"));
  if (st.tools.reference) tools.append(el("button", {class: "icon-btn", title: "Formulalar", onclick: showReference}, "📐"));
  tools.append(el("button", {class: "icon-btn", title: "Savollar xaritasi", onclick: () => showNavigator(st)}, "☰"));
  const bar = el("div", {class: "sbar"}, timer, el("div", {class: "prog"}, progT, el("div", {class: "bar"}, progBar)), saveSt, tools);

  const ansPane = el("div", {class: "pane anspane active"});
  const pdfPane = el("div", {class: "pane pdfpane"});
  const seg = el("div", {class: "seg" + (hasFiles ? "" : " hidden")});
  const segBtn = (key, txt) => el("button", {class: key === "ans" ? "on" : "", onclick: (ev) => {
    st.pane = key; seg.querySelectorAll("button").forEach(b => b.classList.toggle("on", b === ev.currentTarget));
    ansPane.classList.toggle("active", key === "ans"); pdfPane.classList.toggle("active", key === "pdf");
    if (key === "pdf") ensureFiles(st, pdfPane);
  }}, txt);
  seg.append(segBtn("pdf", "📄 Savollar"), segBtn("ans", "📝 Javoblar"));
  const submitBtn = el("button", {class: "btn", onclick: () => askSubmit(st)}, "Yakunlash");
  const foot = el("div", {class: "sfoot"}, submitBtn);
  const panes = el("div", {class: "panes"}, hasFiles ? pdfPane : null, ansPane);
  const root = el("div", {class: "solve"}, bar, seg, panes, foot);
  mount(root);
  st.ui = {timer, progT, progBar, saveSt, ansPane, pdfPane, submitBtn};
  if (hasFiles && window.matchMedia("(min-width: 860px)").matches) ensureFiles(st, pdfPane);
  renderAnswers(st); updateProgress(st);

  const tick = () => {
    const remain = st.endTs - (Date.now() / 1000 + st.offset);
    timer.textContent = "⏱ " + fmtTime(remain);
    timer.className = "timer" + (remain <= 60 ? " crit" : remain <= 300 ? " warn" : "");
    if (remain <= 300 && remain > 60 && !st.warned[5]) { st.warned[5] = 1; toast("⚠️ 5 daqiqa qoldi"); haptic("warning"); }
    if (remain <= 60 && !st.warned[1]) { st.warned[1] = 1; toast("⚠️ 1 daqiqa qoldi!"); haptic("warning"); }
    if (remain <= 0 && !st.submitted) { stopTick(); toast("⏱ Vaqt tugadi — javoblaringiz topshirilmoqda", 3000); doSubmit(st, true); }
  };
  tick(); S.tick = setInterval(tick, 500);
}

function isAnswered(st, q) { return !!st.answers[String(q)]; }
function answeredCount(st) { let n = 0; for (let q = 1; q <= st.qcount; q++) if (isAnswered(st, q)) n++; return n; }
function updateProgress(st) {
  const n = answeredCount(st);
  st.ui.progT.textContent = `${st.name} · ${n}/${st.qcount} javob`;
  st.ui.progBar.style.width = (st.qcount ? n / st.qcount * 100 : 0) + "%";
  st.ui.submitBtn.textContent = st.sectionKey ? `Bo'limni yakunlash (${n}/${st.qcount})` : `Yakunlash (${n}/${st.qcount})`;
}

function renderAnswers(st) {
  const pane = st.ui.ansPane;
  const head = el("div", {class: "pane-head"},
    el("span", {}, st.view === "sheet" ? "Javoblar varaqasi" : "Bittalab ko'rinish"),
    el("button", {class: "btn ghost", style: "min-height:32px;padding:4px 10px;font-size:13px", onclick: () => {
      st.view = st.view === "sheet" ? "card" : "sheet"; renderAnswers(st);
    }}, st.view === "sheet" ? "▭ Bittalab" : "▦ Varaq"));
  pane.replaceChildren(head, st.view === "sheet" ? buildSheet(st) : buildCard(st));
}

function buildSheet(st) {
  const wrap = el("div", {});
  st.layout.forEach(b => {
    if (b.name || st.layout.length > 1) {
      wrap.append(el("div", {class: "block-h"}, el("div", {}, b.name || "Savollar"),
        el("span", {}, `${b.from}–${b.to}` + (b.w ? ` · har biri ${b.w} ball` : ""))));
    }
    for (let q = b.from; q <= b.to; q++) wrap.append(buildRow(st, q));
  });
  wrap.addEventListener("click", e => {
    const row = e.target.closest(".qrow"); if (!row) return;
    const q = +row.dataset.q;
    if (e.target.closest(".bub")) {
      const l = e.target.closest(".bub").dataset.l;
      setAnswer(st, q, st.answers[String(q)] === l ? null : l);
    } else if (e.target.closest(".flagb")) toggleFlag(st, q);
    else if (e.target.closest(".openbtn")) openEditor(st, q);
    else if (e.target.closest(".qn")) { st.current = q; st.view = "card"; renderAnswers(st); }
  });
  return wrap;
}
function buildRow(st, q) {
  const row = el("div", {class: "qrow" + (q % 2 === 0 ? " alt" : "") + (isAnswered(st, q) ? " done" : ""), "data-q": q, id: "row" + q});
  row.append(el("button", {class: "qn", title: "Bittalab ko'rinishda ochish"}, q));
  if (st.open.has(q)) {
    const oa = st.openAnswers[String(q)];
    const label = oa && (oa.text || oa.photo_id) ? ("✅ " + (oa.text ? oa.text : "📷 Rasm biriktirilgan")) : "✍️ Yozma javob (bosing)";
    row.append(el("button", {class: "openbtn" + (oa && (oa.text || oa.photo_id) ? " has" : "")}, label));
  } else {
    const bubs = el("div", {class: "bubs" + (st.nOpts > 4 ? " many" : "")});
    for (let i = 0; i < st.nOpts; i++) {
      const l = LETTERS[i].toLowerCase();
      bubs.append(el("button", {class: "bub" + (st.answers[String(q)] === l ? " on" : ""), "data-l": l, "aria-label": `${q}-savol ${LETTERS[i]}`}, LETTERS[i]));
    }
    row.append(bubs);
  }
  row.append(el("button", {class: "flagb" + (st.flags.has(q) ? " on" : ""), title: "Keyinroq qaytish uchun belgilash"}, "⚑"));
  return row;
}
function refreshRow(st, q) {
  const old = document.getElementById("row" + q);
  if (old) old.replaceWith(buildRow(st, q));
}

function buildCard(st) {
  const q = st.current, sq = String(q);
  const box = el("div", {class: "qcard"});
  const flagBtn = el("button", {class: "btn ghost", style: "min-height:34px;padding:4px 10px;font-size:13px" + (st.flags.has(q) ? ";color:var(--flag)" : ""),
    onclick: () => { toggleFlag(st, q); renderAnswers(st); }}, st.flags.has(q) ? "⚑ Belgilangan" : "⚑ Belgilash");
  box.append(el("div", {class: "ch"}, el("div", {}, el("b", {}, q + "-savol"), el("span", {class: "muted"}, " / " + st.qcount)), flagBtn));
  if (st.open.has(q)) {
    const oa = st.openAnswers[sq] || {};
    const ta = el("textarea", {class: "ta", maxlength: "500", placeholder: "Javobingizni shu yerga yozing..."});
    ta.value = oa.text || "";
    ta.addEventListener("input", () => setOpen(st, q, ta.value, (st.openAnswers[sq] || {}).photo_id || null, true));
    box.append(el("div", {class: "status warn small"}, "❗️ Bu savol variantsiz — javobni matn bilan yozing (yoki yechimingiz rasmini yuboring)."),
      ta, photoControl(st, q));
  } else {
    const elim = st.elim[q] || (st.elim[q] = new Set());
    for (let i = 0; i < st.nOpts; i++) {
      const l = LETTERS[i].toLowerCase();
      const on = st.answers[sq] === l, crossed = elim.has(l);
      const opt = el("div", {class: "opt" + (on ? " on" : "") + (crossed ? " elim" : "")},
        el("div", {class: "ol"}, LETTERS[i]), el("div", {class: "lbl"}, `${LETTERS[i]} variant`),
        st.tools.eliminator ? el("button", {class: "strike", title: "Variantni chiqarib tashlash", onclick: e => {
          e.stopPropagation(); if (elim.has(l)) elim.delete(l); else { elim.add(l); if (st.answers[sq] === l) setAnswer(st, q, null); }
          renderAnswers(st);
        }}, crossed ? "↺" : "✕") : null);
      opt.addEventListener("click", () => { setAnswer(st, q, on ? null : l); renderAnswers(st); });
      box.append(opt);
    }
  }
  box.append(el("div", {class: "cnav"},
    el("button", {class: "btn ghost", disabled: q <= 1, onclick: () => { st.current = q - 1; renderAnswers(st); scheduleSave(st); }}, "⬅️ Oldingi"),
    q < st.qcount ? el("button", {class: "btn", onclick: () => { st.current = q + 1; renderAnswers(st); scheduleSave(st); }}, "Keyingi ➡️")
      : el("button", {class: "btn", onclick: () => askSubmit(st)}, "✅ Yakunlash")));
  return box;
}

function setAnswer(st, q, l) {
  const sq = String(q);
  if (l) st.answers[sq] = l; else delete st.answers[sq];
  haptic("sel");
  if (st.view === "sheet") refreshRow(st, q);
  updateProgress(st); scheduleSave(st);
}
function toggleFlag(st, q) {
  if (st.flags.has(q)) st.flags.delete(q); else st.flags.add(q);
  haptic("sel");
  if (st.view === "sheet") refreshRow(st, q);
  scheduleSave(st);
}
function setOpen(st, q, text, photoId, quiet) {
  const sq = String(q);
  text = (text || "").slice(0, 500);
  if (text.trim() || photoId) { st.openAnswers[sq] = {text: text.trim() ? text : null, photo_id: photoId || null}; st.answers[sq] = "x"; }
  else { delete st.openAnswers[sq]; delete st.answers[sq]; }
  if (!quiet && st.view === "sheet") refreshRow(st, q);
  updateProgress(st); scheduleSave(st);
}
function photoControl(st, q) {
  const sq = String(q);
  const wrap = el("div", {style: "margin-top:10px"});
  const draw = () => {
    const oa = st.openAnswers[sq] || {};
    const file = el("input", {type: "file", accept: "image/*", class: "hidden"});
    file.addEventListener("change", async () => {
      if (!file.files || !file.files[0]) return;
      const done = overlay("Rasm yuklanmoqda...");
      try {
        const blob = await compressImage(file.files[0]);
        const fd = new FormData();
        fd.append("test_id", st.testId); if (st.sectionKey) fd.append("section_key", st.sectionKey);
        fd.append("q", String(q)); fd.append("file", blob, "javob.jpg");
        const r = await api("/api/upload_photo", {form: fd});
        st.photos[sq] = URL.createObjectURL(blob);
        setOpen(st, q, (st.openAnswers[sq] || {}).text || "", r.photo_id);
        toast("📷 Rasm biriktirildi"); haptic("ok");
      } catch (e) { handleErr(e); }
      done(); draw();
    });
    const kids = [file, el("button", {class: "btn ghost", onclick: () => file.click()}, oa.photo_id ? "📷 Rasmni almashtirish" : "📷 Rasm biriktirish")];
    if (oa.photo_id) {
      kids.push(el("button", {class: "btn ghost", style: "margin-left:8px", onclick: () => { setOpen(st, q, oa.text || "", null); draw(); }}, "🗑"));
      if (st.photos[sq]) kids.push(el("img", {class: "thumb", src: st.photos[sq], alt: "javob rasmi"}));
      else kids.push(el("div", {class: "muted small", style: "margin-top:6px"}, "✅ Rasm biriktirilgan"));
    }
    wrap.replaceChildren(...kids);
  };
  draw();
  return wrap;
}
async function compressImage(file) {
  try {
    const bmp = await createImageBitmap(file);
    const k = Math.min(1, 1600 / Math.max(bmp.width, bmp.height));
    const c = document.createElement("canvas");
    c.width = Math.round(bmp.width * k); c.height = Math.round(bmp.height * k);
    c.getContext("2d").drawImage(bmp, 0, 0, c.width, c.height);
    return await new Promise(res => c.toBlob(b => res(b || file), "image/jpeg", 0.85));
  } catch (e) { return file; }
}
function openEditor(st, q) {
  const sq = String(q);
  const ta = el("textarea", {class: "ta", maxlength: "500", placeholder: "Javobingizni yozing..."});
  ta.value = (st.openAnswers[sq] || {}).text || "";
  const cnt = el("div", {class: "words"});
  const upd = () => { cnt.textContent = `${ta.value.length}/500`; };
  ta.addEventListener("input", () => { upd(); setOpen(st, q, ta.value, (st.openAnswers[sq] || {}).photo_id || null, true); });
  upd();
  const s = sheet(el("h2", {}, `${q}-savol — yozma javob`), el("p", {class: "muted small"}, "Kalit bilan avtomatik tekshiriladi. Rasm yuborsangiz, ustoz tekshiradi."),
    ta, cnt, photoControl(st, q),
    el("div", {style: "height:12px"}),
    el("button", {class: "btn block", onclick: () => { s.remove(); refreshRow(st, q); }}, "✅ Saqlash"));
  setTimeout(() => ta.focus(), 120);
  s.addEventListener("click", e => { if (e.target === s) refreshRow(st, q); });
}

function goToQuestion(st, q) {
  st.current = q;
  // ✅ Avval "Javoblar" panelini ochamiz — yashirin (display:none) panelda
  // kerakli qatorga aylantirib bo'lmaydi.
  if (st.pane !== "ans") { const b = document.querySelector(".seg button:last-child"); if (b) b.click(); }
  if (st.view === "card") return renderAnswers(st);
  requestAnimationFrame(() => {
    const r = document.getElementById("row" + q);
    if (r) { r.scrollIntoView({block: "center", behavior: "smooth"}); r.classList.add("cur"); setTimeout(() => r.classList.remove("cur"), 1600); }
  });
}
function unansweredList(st) { const out = []; for (let q = 1; q <= st.qcount; q++) if (!isAnswered(st, q)) out.push(q); return out; }
function showNavigator(st) {
  const grid = el("div", {class: "navgrid"});
  for (let q = 1; q <= st.qcount; q++) {
    const cls = (isAnswered(st, q) ? "ans " : "") + (st.flags.has(q) ? "flag" : "");
    grid.append(el("button", {class: cls, onclick: () => { s.remove(); goToQuestion(st, q); }}, st.open.has(q) ? q + "✍" : q));
  }
  const n = answeredCount(st), miss = unansweredList(st);
  const firstFlag = Array.from(st.flags).sort((a, b) => a - b)[0];
  const s = sheet(el("h2", {}, "Savollar xaritasi"), el("div", {class: "muted small"}, `Javob berilgan: ${n}/${st.qcount} · Belgilangan: ${st.flags.size}`),
    el("div", {class: "row-btns", style: "margin-top:10px"},
      miss.length ? el("button", {class: "btn ghost", onclick: () => { s.remove(); goToQuestion(st, miss[0]); }}, `➡️ Javobsiz: ${miss[0]}`) : null,
      firstFlag ? el("button", {class: "btn ghost", onclick: () => { s.remove(); goToQuestion(st, firstFlag); }}, `⚑ Belgilangan: ${firstFlag}`) : null),
    grid, el("div", {class: "legend"}, el("span", {}, el("i", {style: "background:var(--accent);border-color:var(--accent)"}), "javob berilgan"),
      el("span", {}, el("i", {style: "box-shadow:0 0 0 2px var(--flag)"}), "belgilangan"), el("span", {}, el("i", {}), "javobsiz"), el("span", {}, "✍ — ochiq savol")));
}

/* ------------ avtomatik saqlash ------------ */
function savePayload(st) {
  return {test_id: st.testId, section_key: st.sectionKey, answers: st.answers, open_answers: st.openAnswers,
          flags: Array.from(st.flags), current_q: st.current};
}
function setSaveState(st, s) {
  if (!st.ui) return;
  st.ui.saveSt.textContent = {saving: "⏳", ok: "✓", offline: "⚠️", "": ""}[s] || "";
  st.ui.saveSt.title = {saving: "Saqlanmoqda...", ok: "Saqlandi", offline: "Internet yo'q — javoblar qurilmada saqlandi"}[s] || "";
}
function scheduleSave(st) {
  if (st.submitted) return;
  lsSet(lsKeyFor(st), {answers: st.answers, openAnswers: st.openAnswers, flags: Array.from(st.flags), ts: Math.floor(Date.now() / 1000 + st.offset)});
  setSaveState(st, "saving");
  clearTimeout(st.saveTimer);
  st.saveTimer = setTimeout(() => flushSave(st), 1200);
}
async function flushSave(st) {
  clearTimeout(st.saveTimer);
  if (st.submitted) return;
  if (st.saving) { st.again = true; return; }
  st.saving = true;
  try {
    const r = await api("/api/test/save", {body: savePayload(st)});
    if (r.end_ts) st.endTs = r.end_ts;
    setSaveState(st, "ok");
  } catch (e) {
    if (e.network) { setSaveState(st, "offline"); st.saveTimer = setTimeout(() => flushSave(st), 5000); }
    else if (e.error === "time_over" || e.error === "no_session" || e.error === "stale") { setSaveState(st, ""); }
    else setSaveState(st, "offline");
  }
  st.saving = false;
  if (st.again) { st.again = false; flushSave(st); }
}
document.addEventListener("visibilitychange", () => {
  const st = S.st;
  if (document.visibilityState === "hidden" && st && !st.submitted && st.ui) {
    fetch(st.writing ? "/api/test/save" : "/api/test/save", {method: "POST", keepalive: true,
      headers: {"X-Telegram-Init-Data": initData, "Content-Type": "application/json"},
      body: JSON.stringify(st.writing ? writingPayload(st) : savePayload(st))}).catch(() => {});
  }
});

/* ------------ topshirish ------------ */
async function askSubmit(st) {
  const n = answeredCount(st), missList = unansweredList(st), miss = missList.length;
  const lines = [];
  if (miss) lines.push(`${miss} ta savol javobsiz qolgan: ${missList.slice(0, 12).join(", ")}${miss > 12 ? " va boshqalar." : "."}`);
  if (st.flags.size) lines.push(`${st.flags.size} ta savol "keyinroq qaytish" uchun belgilangan.`);
  lines.push(st.sectionKey ? "Bo'lim yakunlangach unga qaytib bo'lmaydi." : "Topshirgandan keyin javoblarni o'zgartirib bo'lmaydi.");
  if (await confirmBox(st.sectionKey ? "Bo'limni yakunlaysizmi?" : "Testni yakunlaysizmi?", lines.join(" "), "✅ Ha, topshirish")) doSubmit(st, false);
}
async function doSubmit(st, auto) {
  if (st.submitted) return;
  st.submitted = true; stopTick(); clearTimeout(st.saveTimer);
  const done = overlay(auto ? "Vaqt tugadi — topshirilmoqda..." : "Javoblar baholanmoqda...");
  try {
    let r;
    if (st.sectionKey) r = await api("/api/staged/submit", {body: savePayload(st)});
    else r = await api("/api/test/submit", {body: {test_id: st.testId, answers: st.answers, open_answers: st.openAnswers}});
    lsDel(lsKeyFor(st));
    done(); haptic("ok"); closeCalc();
    if (tg) { try { tg.disableClosingConfirmation(); } catch (e) {} }
    S.st = null;
    if (st.sectionKey && !r.finished) { toast(`✅ «${st.name}» topshirildi`); return renderOverview(r.overview); }
    renderResult(r.result);
  } catch (e) {
    done();
    if (e.network) {
      st.submitted = false; toast("Internet yo'q — javoblar saqlangan. Aloqa tiklangach qayta bosing.", 4000);
      if (!auto) return;
      setTimeout(() => doSubmit(st, true), 4000); return;
    }
    if (e.error === "stale" || e.error === "no_session") {
      S.st = null; toast(e.message, 3500);
      try { const r = await api("/api/result/" + st.testId); if (r.result.state !== "none") return renderResult(r.result); } catch (x) {}
      return renderHome();
    }
    st.submitted = false; handleErr(e);
  }
}

/* ------------ Writing (matnli bo'lim) ------------ */
function writingPayload(st) {
  return {test_id: st.testId, section_key: st.sectionKey, text: st.text, photo_id: st.photoId};
}
function renderWriting(p) {
  stopTick(); showTabs(false);
  const st = {writing: true, testId: p.test.id, sectionKey: p.key, name: p.name, text: p.text || "", photoId: p.photo_id || null,
    endTs: p.end_ts, offset: p.server_now - Date.now() / 1000, submitted: false, saveTimer: null, warned: {}, qcount: 0};
  const local = lsGet(`et:${S.uid}:${st.testId}:${st.sectionKey}:${st.endTs}`);
  if (local && local.ts > (p.updated || 0) && typeof local.text === "string") st.text = local.text;
  S.st = st;
  if (tg) { try { tg.enableClosingConfirmation(); } catch (e) {} }
  setBack(async () => { if (await confirmBox("Bo'limdan chiqasizmi?", "Matningiz saqlangan, vaqt ketaveradi.", "Chiqish")) { await saveWriting(st); renderHome(); } });
  const timer = el("div", {class: "timer"}, "--:--");
  const saveSt = el("span", {class: "savest"});
  const ta = el("textarea", {class: "ta big", maxlength: String(p.text_max || 4000), placeholder: "Javobingizni shu yerga yozing..."});
  ta.value = st.text;
  const words = el("div", {class: "words"});
  const upd = () => { const w = ta.value.trim() ? ta.value.trim().split(/\s+/).length : 0; words.textContent = `${w} so'z · ${ta.value.length}/${p.text_max || 4000} belgi`; };
  ta.addEventListener("input", () => {
    st.text = ta.value; upd(); saveSt.textContent = "⏳";
    lsSet(`et:${S.uid}:${st.testId}:${st.sectionKey}:${st.endTs}`, {text: st.text, ts: Math.floor(Date.now() / 1000 + st.offset)});
    clearTimeout(st.saveTimer); st.saveTimer = setTimeout(() => saveWriting(st).then(ok => { saveSt.textContent = ok ? "✓" : "⚠️"; }), 1500);
  });
  upd();
  const photoBox = el("div", {});
  const drawPhoto = () => {
    const file = el("input", {type: "file", accept: "image/*", class: "hidden"});
    file.addEventListener("change", async () => {
      if (!file.files || !file.files[0]) return;
      const done = overlay("Rasm yuklanmoqda...");
      try {
        const blob = await compressImage(file.files[0]);
        const fd = new FormData(); fd.append("test_id", st.testId); fd.append("section_key", st.sectionKey); fd.append("file", blob, "javob.jpg");
        const r = await api("/api/upload_photo", {form: fd});
        st.photoId = r.photo_id; st.photoUrl = URL.createObjectURL(blob); await saveWriting(st); toast("📷 Rasm biriktirildi");
      } catch (e) { handleErr(e); }
      done(); drawPhoto();
    });
    photoBox.replaceChildren(...clean([file, el("button", {class: "btn ghost", onclick: () => file.click()}, st.photoId ? "📷 Rasmni almashtirish" : "📷 Qo'lda yozilgan javob rasmini biriktirish"),
      st.photoId ? (st.photoUrl ? el("img", {class: "thumb", src: st.photoUrl}) : el("div", {class: "muted small"}, "✅ Rasm biriktirilgan")) : null]));
  };
  drawPhoto();
  const submit = async () => {
    if (!st.text.trim() && !st.photoId) return toast("Avval javob yozing yoki rasm biriktiring");
    if (!await confirmBox("Bo'limni yakunlaysizmi?", "Topshirgandan keyin matnni o'zgartirib bo'lmaydi.", "✅ Topshirish")) return;
    submitWriting(st);
  };
  mount(el("div", {class: "solve"},
    el("div", {class: "sbar"}, timer, el("div", {class: "prog"}, el("div", {class: "t"}, p.name)), saveSt),
    el("div", {class: "pane active", style: "padding:12px 14px"},
      el("div", {class: "card", style: "margin-top:0"}, el("div", {class: "muted small"}, "✍️ Topshiriq"), el("div", {class: "prompt"}, p.prompt)),
      ta, words, el("div", {style: "height:10px"}), photoBox),
    el("div", {class: "sfoot"}, el("button", {class: "btn", onclick: submit}, "✅ Bo'limni topshirish"))));
  const tick = () => {
    const remain = st.endTs - (Date.now() / 1000 + st.offset);
    timer.textContent = "⏱ " + fmtTime(remain);
    timer.className = "timer" + (remain <= 60 ? " crit" : remain <= 300 ? " warn" : "");
    if (remain <= 0 && !st.submitted) { stopTick(); toast("⏱ Vaqt tugadi — matningiz topshirilmoqda"); submitWriting(st, true); }
  };
  tick(); S.tick = setInterval(tick, 500);
}
async function saveWriting(st) {
  if (st.submitted) return true;
  try { await api("/api/test/save", {body: writingPayload(st)}); return true; } catch (e) { return false; }
}
async function submitWriting(st, auto) {
  if (st.submitted) return;
  st.submitted = true; stopTick(); clearTimeout(st.saveTimer);
  const done = overlay("Topshirilmoqda... (AI tekshirsa bir necha soniya davom etadi)");
  try {
    const r = await api("/api/staged/submit", {body: writingPayload(st)});
    lsDel(`et:${S.uid}:${st.testId}:${st.sectionKey}:${st.endTs}`);
    done(); haptic("ok"); S.st = null;
    if (tg) { try { tg.disableClosingConfirmation(); } catch (e) {} }
    if (!r.finished) { toast(`✅ «${st.name}» topshirildi`); return renderOverview(r.overview); }
    renderResult(r.result);
  } catch (e) {
    done(); st.submitted = false;
    if (e.network && auto) { setTimeout(() => submitWriting(st, true), 4000); return; }
    handleErr(e);
  }
}

/* ------------ fayllar: PDF / audio ------------ */
async function ensureFiles(st, pane) {
  if (st.filesLoaded) return;
  st.filesLoaded = true;
  pane.replaceChildren(el("div", {class: "center", style: "min-height:40vh"}, el("div", {class: "spin dark"}), el("div", {class: "muted"}, "Fayl yuklanmoqda...")));
  const parts = [];
  for (const f of st.files) {
    const box = el("div", {});
    parts.push(box);
    try {
      const res = await fetch(`/api/file/${st.testId}/${f.slot}`, {headers: {"X-Telegram-Init-Data": initData}});
      if (!res.ok) { let m = "Faylni yuklab bo'lmadi."; try { m = (await res.json()).message || m; } catch (e) {} throw new Error(m); }
      const ct = res.headers.get("Content-Type") || "";
      const blob = await res.blob();
      if (ct.includes("pdf")) await renderPdf(await blob.arrayBuffer(), box, pane);
      else if (ct.startsWith("audio")) box.append(el("div", {class: "media"}, el("div", {class: "muted small"}, "🎧 Audio (Listening)"), el("audio", {controls: true, src: URL.createObjectURL(blob), preload: "auto"})));
      else if (ct.startsWith("video")) box.append(el("div", {class: "media"}, el("video", {controls: true, src: URL.createObjectURL(blob), playsinline: true})));
      else if (ct.startsWith("image")) box.append(el("div", {class: "pdfwrap"}, el("img", {src: URL.createObjectURL(blob), class: "pdfpage", style: "width:100%"})));
      else box.append(el("div", {class: "media"}, el("a", {href: URL.createObjectURL(blob), download: "savollar"}, "📥 Faylni yuklab olish")));
    } catch (e) {
      box.append(el("div", {class: "status bad", style: "margin:12px"}, "⚠️ " + (e.message || "Faylni yuklab bo'lmadi.")));
    }
  }
  // audio har doim tepada tursin
  parts.sort((a, b) => (b.querySelector("audio") ? 1 : 0) - (a.querySelector("audio") ? 1 : 0));
  pane.replaceChildren(...parts);
}
async function renderPdf(buf, box, pane) {
  if (!window.pdfjsLib) { box.append(el("div", {class: "status bad", style: "margin:12px"}, "PDF ko'ruvchini yuklab bo'lmadi.")); return; }
  pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
  const pdf = await pdfjsLib.getDocument({data: buf}).promise;
  let zoom = 1;
  const pages = el("div", {class: "pdfwrap"});
  const info = el("span", {}, `${pdf.numPages} sahifa`);
  const z = d => { zoom = Math.min(3, Math.max(0.6, zoom + d)); zl.textContent = Math.round(zoom * 100) + "%"; rerender(); };
  const zl = el("span", {}, "100%");
  const head = el("div", {class: "pane-head"}, info, el("div", {style: "display:flex;gap:6px;align-items:center"},
    el("button", {class: "icon-btn", onclick: () => z(-0.2)}, "−"), zl, el("button", {class: "icon-btn", onclick: () => z(0.2)}, "+")));
  box.append(head, pages);
  const slots = [];
  const io = new IntersectionObserver(ents => ents.forEach(en => { if (en.isIntersecting) draw(en.target); }), {root: pane, rootMargin: "600px 0px"});
  for (let i = 1; i <= pdf.numPages; i++) {
    const d = el("div", {class: "pdfpage", "data-p": i, style: "width:100%;min-height:300px"});
    slots.push(d); pages.append(d); io.observe(d);
  }
  async function draw(d) {
    const want = zoom + ":" + pane.clientWidth;
    if (d.dataset.done === want || d.dataset.busy) return;
    d.dataset.busy = "1";
    try {
      const page = await pdf.getPage(+d.dataset.p);
      const avail = Math.max(240, (pane.clientWidth || window.innerWidth) - 16);
      const base = page.getViewport({scale: 1});
      const cssW = avail * zoom, scale = cssW / base.width, dpr = Math.min(window.devicePixelRatio || 1, 2.5);
      const vp = page.getViewport({scale: scale * dpr});
      const c = document.createElement("canvas"); c.width = vp.width; c.height = vp.height;
      await page.render({canvasContext: c.getContext("2d"), viewport: vp}).promise;
      d.style.width = cssW + "px"; d.style.minHeight = ""; d.replaceChildren(c); d.dataset.done = want;
    } catch (e) { d.textContent = "Sahifani chizib bo'lmadi"; }
    delete d.dataset.busy;
  }
  function rerender() { slots.forEach(d => { const r = d.getBoundingClientRect(), pr = pane.getBoundingClientRect(); if (r.bottom > pr.top - 600 && r.top < pr.bottom + 600) draw(d); }); }
  window.addEventListener("resize", () => rerender());
}

/* ------------ natija ------------ */
async function showResult(testId) {
  const done = overlay("Natija yuklanmoqda...");
  try { const r = await api("/api/result/" + testId); done(); renderResult(r.result); }
  catch (e) { done(); handleErr(e); }
}
function ring(pct, label) {
  const p = Math.max(0, Math.min(100, pct || 0)), R = 62, C = 2 * Math.PI * R;
  const color = p >= 70 ? "var(--ok)" : p >= 50 ? "var(--warn)" : "var(--bad)";
  const svg = `<svg class="ring" viewBox="0 0 150 150"><circle cx="75" cy="75" r="${R}" fill="none" stroke="rgba(127,127,127,.18)" stroke-width="12"/>
    <circle cx="75" cy="75" r="${R}" fill="none" stroke="${color}" stroke-width="12" stroke-linecap="round" stroke-dasharray="${C}" stroke-dashoffset="${C * (1 - p / 100)}" transform="rotate(-90 75 75)"/>
    <text x="75" y="${label ? 76 : 84}" text-anchor="middle" font-size="30" fill="currentColor">${label ? "" : (Number.isInteger(p) ? p : p.toFixed(1)) + "%"}</text></svg>`;
  const d = el("div", {html: svg});
  if (label) d.querySelector("text").textContent = label;
  return d;
}
function renderResult(v) {
  stopTick(); closeCalc(); showTabs(false); setBack(renderHome); S.st = null;
  if (tg) { try { tg.disableClosingConfirmation(); } catch (e) {} }
  const t = v.test;
  const kids = [el("div", {class: "muted small"}, `${t.exam_label} · ${t.code}`), el("h1", {}, t.title)];
  if (v.state === "none") {
    kids.push(el("div", {class: "status warn"}, "Bu test bo'yicha natija hali yo'q."));
  } else if (v.state === "staged") {
    const pend = v.status === "pending_review";
    const big = v.overall_grade != null && !pend ? String(v.overall_grade) : "⏳";
    kids.push(el("div", {class: "card", style: "text-align:center"}, ring(pend ? 0 : (v.percent != null ? v.percent : 0), big),
      el("div", {class: "muted"}, pend ? "Ustoz (yoki AI) baholayapti — yakuniy natija botga keladi." : "Umumiy natija")));
    const tb = el("table", {class: "tb"}, el("tr", {}, el("th", {}, "Bo'lim"), el("th", {}, "Natija"), el("th", {}, "Daraja/ball")));
    v.sections.forEach(s => tb.append(el("tr", {}, el("td", {}, s.name),
      el("td", {}, s.total != null ? `${s.score}/${s.total}` : (s.submitted ? "topshirildi" : "—")),
      el("td", {}, s.band != null ? String(s.band) : (s.earned != null ? `${(+s.earned).toFixed(1)}/${(+s.max).toFixed(1)}` : (s.submitted ? "⏳" : "—"))))));
    kids.push(el("div", {class: "card"}, el("h2", {}, "Bo'limlar"), tb));
    v.sections.filter(s => s.feedback).forEach(s => kids.push(el("div", {class: "card"}, el("b", {}, s.name + " — fikr-mulohaza"), el("p", {class: "prompt"}, s.feedback))));
  } else {
    const rasch = v.rasch;
    let big;
    if (rasch && !rasch.finalized) kids.push(el("div", {class: "card", style: "text-align:center"}, ring(0, "⏳"),
      el("div", {}, "⚡ Rasch ball test yakunlangach hammaga bir vaqtda e'lon qilinadi."), el("div", {class: "muted small"}, rasch.remain_text || "")));
    else if (rasch) kids.push(el("div", {class: "card", style: "text-align:center"}, ring(rasch.percent, (rasch.ball != null ? (+rasch.ball).toFixed(1) : "-")),
      el("div", {}, "Rasch ball (0–75)"), rasch.grade ? el("div", {class: "chips", style: "justify-content:center"}, el("span", {class: "chip acc"}, "Daraja: " + rasch.grade)) : null));
    else kids.push(el("div", {class: "card", style: "text-align:center"}, ring(v.percent),
      v.grade ? el("div", {class: "chips", style: "justify-content:center"}, el("span", {class: "chip acc"}, gradeLabel(t.exam_type) + v.grade)) : null));
    if (!rasch || rasch.finalized) {
      const wrong = (v.scored_total != null ? v.scored_total : v.total) - v.score;
      kids.push(el("div", {class: "stats"}, el("div", {class: "stat"}, el("b", {style: "color:var(--ok)"}, v.score), el("span", {}, "to'g'ri")),
        el("div", {class: "stat"}, el("b", {style: "color:var(--bad)"}, Math.max(0, wrong)), el("span", {}, "xato/javobsiz")),
        el("div", {class: "stat"}, el("b", {}, (v.percent || 0).toFixed(1) + "%"), el("span", {}, "foiz"))));
    } else kids.push(el("div", {class: "status"}, "📊 Javoblaringiz qabul qilindi."));
    if (t.exam_type === "dtm" && v.max) kids.push(el("div", {class: "card"}, el("div", {class: "kv"}, el("span", {}, "🏁 DTM ball"), el("b", {}, `${(+v.earned).toFixed(1)} / ${(+v.max).toFixed(1)}`)),
      ...(v.subjects || []).map(s => el("div", {class: "kv"}, el("span", {}, `${s.name} (${s.q} savol × ${s.w})`), el("span", {}, `${s.correct} ta · ${(+s.earned).toFixed(1)} ball`)))));
    if (v.open_pending) kids.push(el("div", {class: "status warn"}, `✍️ ${v.open_pending} ta ochiq javobingizni ustoz tekshiradi — yakuniy natija botga keladi.`));
    if (v.review) {
      const grid = el("div", {class: "navgrid"});
      v.review.forEach(it => grid.append(el("button", {class: it.ok == null ? "pend" : it.ok ? "ok" : "bad", onclick: () => {
        toast(it.open ? `${it.n}-savol (ochiq): siz — «${it.you || "javobsiz"}», kalit — «${it.key || "-"}»` :
          `${it.n}-savol: siz — ${it.you || "javobsiz"}, to'g'ri — ${it.key}`, 3500);
      }}, it.n)));
      kids.push(el("div", {class: "card"}, el("h2", {}, "Javoblar tahlili"), el("div", {class: "muted small"}, "Raqamni bosing — javobingiz va to'g'ri javob ko'rinadi."), grid,
        el("div", {class: "legend"}, el("span", {}, el("i", {style: "background:rgba(31,157,85,.3);border-color:var(--ok)"}), "to'g'ri"),
          el("span", {}, el("i", {style: "background:rgba(214,69,69,.25);border-color:var(--bad)"}), "xato"), el("span", {}, el("i", {style: "border-style:dashed"}), "tekshirilmoqda"))));
    } else if (v.review_allowed === false) kids.push(el("div", {class: "status"}, "🔒 Guruh xavfsizligi uchun batafsil tahlil test vaqti tugagach ochiladi."));
  }
  if (v.certificate && v.result_id) kids.push(el("div", {class: "card"}, el("b", {}, "🎓 Sertifikat"), el("p", {class: "muted small"}, "Sertifikatingiz botga ham yuborilgan."),
    el("button", {class: "btn block", onclick: () => downloadCert(v.result_id, t.code)}, "📥 Sertifikatni yuklab olish")));
  kids.push(el("div", {class: "card"}, el("p", {class: "muted small"}, "Batafsil natija bot chatiga ham yuborildi."),
    el("button", {class: "btn block", onclick: renderHome}, "🏠 Bosh sahifa"), el("button", {class: "btn block ghost", onclick: openBot}, "💬 Botga qaytish")));
  mount(el("div", {class: "screen"}, ...kids));
}
// Telegram WebView oddiy <a href> orqali faylni yuklab ololmaydi (ayniqsa
// iOS va Desktop) — Telegram 8.0+ da o'zining yuklab olish oynasi bor,
// eski versiyalarda havola tashqi brauzerda ochiladi.
function downloadCert(rid, code) {
  const url = `${location.origin}/api/certificate/${rid}?initData=${encodeURIComponent(initData)}`;
  try {
    if (tg && tg.downloadFile && tg.isVersionAtLeast && tg.isVersionAtLeast("8.0")) {
      return tg.downloadFile({url, file_name: `sertifikat_${code || rid}.pdf`});
    }
    if (tg && tg.openLink) return tg.openLink(url);
  } catch (e) {}
  window.open(url, "_blank");
}
function gradeLabel(et) { return {cefr: "CEFR: ", ielts: "IELTS band: ", sat: "SAT: ", rasch: "Daraja: "}[et] || "Daraja: "; }

/* ------------ profil ------------ */
async function renderProfile() {
  stopTick(); closeCalc(); showTabs(true); setBack(renderHome); setTab("profile");
  mount(el("div", {class: "center"}, el("div", {class: "spin dark"})));
  let d;
  try { d = await api("/api/profile"); } catch (e) { return handleErr(e); }
  if (S.tab !== "profile") return;
  const list = el("div", {});
  if (!d.results.length) list.append(el("div", {class: "muted", style: "text-align:center;padding:24px 0"}, "Hali natijalar yo'q"));
  d.results.forEach(r => list.append(el("div", {class: "list-item", style: "cursor:pointer", onclick: () => showResult(r.test_id)},
    el("div", {}, el("b", {}, r.title), el("div", {class: "muted small"}, `${r.code} · ${r.exam_label} · ${r.date || ""}`),
      r.has_certificate ? el("div", {class: "small"}, "🎓 sertifikat bor") : null),
    el("span", {class: "badge"}, r.percent.toFixed(1) + "%"))));
  mount(el("div", {class: "screen"}, el("h1", {}, d.full_name || "Profil"), el("div", {class: "muted"}, "Statistika va natijalar"),
    el("div", {class: "card"}, el("div", {class: "stats"},
      el("div", {class: "stat"}, el("b", {}, d.summary.count), el("span", {}, "jami test")),
      el("div", {class: "stat"}, el("b", {}, d.summary.avg_percent.toFixed(1) + "%"), el("span", {}, "o'rtacha")),
      el("div", {class: "stat"}, el("b", {}, d.summary.best_percent.toFixed(1) + "%"), el("span", {}, "eng yaxshi")))),
    el("div", {class: "card"}, el("h2", {}, "So'nggi natijalar"), list)));
}

/* ------------ SAT: kalkulyator ------------ */
const Calc = {deg: false, ans: 0, mode: "calc", hist: [], fns: ["x^2-2x-3", "", ""], view: {x0: -10, x1: 10, y0: -7, y1: 7}};
function tokenize(src) {
  const s = src.replace(/×/g, "*").replace(/÷/g, "/").replace(/−/g, "-").replace(/π/g, "pi").replace(/√/g, "sqrt").replace(/²/g, "^2");
  const out = []; let i = 0;
  while (i < s.length) {
    const c = s[i];
    if (/\s/.test(c)) { i++; continue; }
    const m = /^(\d+\.?\d*|\.\d+)(e[+-]?\d+)?/i.exec(s.slice(i));
    if (m) { out.push({t: "n", v: parseFloat(m[0])}); i += m[0].length; continue; }
    const w = /^[a-z]+/i.exec(s.slice(i));
    if (w) { out.push({t: "id", v: w[0].toLowerCase()}); i += w[0].length; continue; }
    if ("+-*/^(),!%".includes(c)) { out.push({t: "op", v: c}); i++; continue; }
    throw new Error("Noto'g'ri belgi: " + c);
  }
  return out;
}
const FN = {sin: 1, cos: 1, tan: 1, asin: 1, acos: 1, atan: 1, sqrt: 1, cbrt: 1, ln: 1, log: 1, abs: 1, exp: 1};
function parseExpr(tokens) {
  let p = 0;
  const peek = () => tokens[p], next = () => tokens[p++];
  const isOp = (v) => peek() && peek().t === "op" && peek().v === v;
  function expr() { let n = term(); while (isOp("+") || isOp("-")) { const o = next().v, r = term(), l = n; n = o === "+" ? (x) => l(x) + r(x) : (x) => l(x) - r(x); } return n; }
  function startsPrimary() { const t = peek(); return t && (t.t === "n" || t.t === "id" || (t.t === "op" && t.v === "(")); }
  function term() {
    let n = unary();
    for (;;) {
      if (isOp("*") || isOp("/")) { const o = next().v, r = unary(), l = n; n = o === "*" ? (x) => l(x) * r(x) : (x) => l(x) / r(x); }
      else if (startsPrimary()) { const r = power(), l = n; n = (x) => l(x) * r(x); }
      else return n;
    }
  }
  function unary() { if (isOp("-")) { next(); const u = unary(); return (x) => -u(x); } if (isOp("+")) { next(); return unary(); } return power(); }
  function power() { const b = postfix(); if (isOp("^")) { next(); const e = unary(); return (x) => Math.pow(b(x), e(x)); } return b; }
  function postfix() {
    let n = primary();
    for (;;) {
      if (isOp("!")) { next(); const l = n; n = (x) => fact(l(x)); }
      else if (isOp("%")) { next(); const l = n; n = (x) => l(x) / 100; }
      else return n;
    }
  }
  function primary() {
    const t = next();
    if (!t) throw new Error("Ifoda tugallanmagan");
    if (t.t === "n") return () => t.v;
    if (t.t === "op" && t.v === "(") { const n = expr(); if (!isOp(")")) throw new Error("Qavs yopilmagan"); next(); return n; }
    if (t.t === "id") {
      if (t.v === "pi") return () => Math.PI;
      if (t.v === "e") return () => Math.E;
      if (t.v === "x") return (x) => x;
      if (t.v === "ans") return () => Calc.ans;
      if (FN[t.v]) {
        let a;
        if (isOp("(")) { next(); a = expr(); if (!isOp(")")) throw new Error("Qavs yopilmagan"); next(); }
        else a = power();
        const f = t.v, rad = (v) => Calc.deg ? v * Math.PI / 180 : v, inv = (v) => Calc.deg ? v * 180 / Math.PI : v;
        return (x) => { const v = a(x);
          switch (f) { case "sin": return Math.sin(rad(v)); case "cos": return Math.cos(rad(v)); case "tan": return Math.tan(rad(v));
            case "asin": return inv(Math.asin(v)); case "acos": return inv(Math.acos(v)); case "atan": return inv(Math.atan(v));
            case "sqrt": return Math.sqrt(v); case "cbrt": return Math.cbrt(v); case "ln": return Math.log(v); case "log": return Math.log10(v);
            case "abs": return Math.abs(v); case "exp": return Math.exp(v); } };
      }
      throw new Error("Noma'lum: " + t.v);
    }
    throw new Error("Kutilmagan: " + t.v);
  }
  const f = expr();
  if (p < tokens.length) throw new Error("Ortiqcha: " + tokens[p].v);
  return f;
}
function fact(n) { if (n < 0 || !Number.isInteger(n) || n > 170) return NaN; let r = 1; for (let i = 2; i <= n; i++) r *= i; return r; }
function fmtNum(v) {
  if (!isFinite(v)) return isNaN(v) ? "Xato" : (v > 0 ? "∞" : "-∞");
  if (Math.abs(v) >= 1e12 || (Math.abs(v) < 1e-7 && v !== 0)) return v.toExponential(6).replace(/\.?0+e/, "e");
  return String(+v.toPrecision(12));
}
function closeCalc() { const c = document.getElementById("calc"); if (c) c.classList.add("hidden"); }
function toggleCalc() {
  const c = document.getElementById("calc");
  if (!c.dataset.built) buildCalc(c);
  c.classList.toggle("hidden");
  if (!c.classList.contains("hidden") && Calc.mode === "graph") drawGraph();
}
function buildCalc(c) {
  c.dataset.built = "1";
  const tabs = el("div", {class: "tb2"});
  const tabBtn = (m, t) => el("button", {class: Calc.mode === m ? "on" : "", onclick: () => { Calc.mode = m; renderCalcBody(); tabs.querySelectorAll("button").forEach(b => b.classList.toggle("on", b.textContent === t)); }}, t);
  tabs.append(tabBtn("calc", "Kalkulyator"), tabBtn("graph", "Grafik"));
  const head = el("div", {class: "ch"}, el("span", {}, "🧮"), tabs, el("button", {class: "icon-btn", style: "width:30px;height:30px", onclick: closeCalc}, "✕"));
  const body = el("div", {});
  c.append(head, body);
  // sudrab ko'chirish
  let drag = null;
  head.addEventListener("pointerdown", e => { if (e.target.closest("button")) return; const r = c.getBoundingClientRect(); drag = {dx: e.clientX - r.left, dy: e.clientY - r.top}; head.setPointerCapture(e.pointerId); });
  head.addEventListener("pointermove", e => { if (!drag) return; c.style.left = Math.max(0, Math.min(window.innerWidth - 80, e.clientX - drag.dx)) + "px"; c.style.top = Math.max(0, Math.min(window.innerHeight - 60, e.clientY - drag.dy)) + "px"; c.style.right = "auto"; });
  head.addEventListener("pointerup", () => { drag = null; });
  function renderCalcBody() { body.replaceChildren(Calc.mode === "calc" ? calcPanel() : graphPanel()); if (Calc.mode === "graph") drawGraph(); }
  renderCalcBody();
}
function calcPanel() {
  const expr = el("input", {class: "expr", inputmode: window.matchMedia("(pointer:coarse)").matches ? "none" : "text", spellcheck: "false", placeholder: "0"});
  expr.value = Calc.expr || "";
  const res = el("div", {class: "res"});
  const hist = el("div", {class: "hist"});
  const drawHist = () => hist.replaceChildren(...Calc.hist.slice(-4).map(h => el("div", {}, h)));
  const live = () => { Calc.expr = expr.value; try { if (!expr.value.trim()) { res.textContent = ""; return; } res.textContent = "= " + fmtNum(parseExpr(tokenize(expr.value))(0)); } catch (e) { res.textContent = ""; } };
  const ins = (s) => { const a = expr.selectionStart ?? expr.value.length, b = expr.selectionEnd ?? expr.value.length; expr.value = expr.value.slice(0, a) + s + expr.value.slice(b); const p = a + s.length; expr.setSelectionRange(p, p); live(); };
  const evalNow = () => {
    try { const v = parseExpr(tokenize(expr.value))(0); Calc.hist.push(`${expr.value} = ${fmtNum(v)}`); Calc.ans = v; expr.value = fmtNum(v); res.textContent = ""; drawHist(); }
    catch (e) { res.textContent = e.message; }
  };
  expr.addEventListener("input", live);
  expr.addEventListener("keydown", e => { if (e.key === "Enter") { e.preventDefault(); evalNow(); } });
  const degBtn = el("button", {class: "fn"}, Calc.deg ? "DEG" : "RAD");
  degBtn.addEventListener("click", () => { Calc.deg = !Calc.deg; degBtn.textContent = Calc.deg ? "DEG" : "RAD"; live(); });
  const K = (label, cls, fn) => el("button", {class: cls || "", onclick: () => { fn(); haptic("sel"); }}, label);
  const keys = el("div", {class: "keys"},
    degBtn, K("(", "", () => ins("(")), K(")", "", () => ins(")")), K("⌫", "", () => { const a = expr.selectionStart ?? expr.value.length; if (a > 0) { expr.value = expr.value.slice(0, a - 1) + expr.value.slice(a); expr.setSelectionRange(a - 1, a - 1); } live(); }), K("AC", "", () => { expr.value = ""; live(); }),
    K("sin", "fn", () => ins("sin(")), K("cos", "fn", () => ins("cos(")), K("tan", "fn", () => ins("tan(")), K("π", "", () => ins("π")), K("e", "", () => ins("e")),
    K("ln", "fn", () => ins("ln(")), K("log", "fn", () => ins("log(")), K("√", "", () => ins("√(")), K("x²", "", () => ins("^2")), K("^", "op", () => ins("^")),
    K("7", "", () => ins("7")), K("8", "", () => ins("8")), K("9", "", () => ins("9")), K("÷", "op", () => ins("÷")), K("n!", "fn", () => ins("!")),
    K("4", "", () => ins("4")), K("5", "", () => ins("5")), K("6", "", () => ins("6")), K("×", "op", () => ins("×")), K("%", "", () => ins("%")),
    K("1", "", () => ins("1")), K("2", "", () => ins("2")), K("3", "", () => ins("3")), K("−", "op", () => ins("−")), K("Ans", "fn", () => ins("ans")),
    K("0", "", () => ins("0")), K(".", "", () => ins(".")), K("1/x", "fn", () => { expr.value = "1/(" + (expr.value || "0") + ")"; live(); }), K("+", "op", () => ins("+")), K("=", "eq", evalNow));
  setTimeout(drawHist, 0);
  return el("div", {}, el("div", {class: "disp"}, expr, res), hist, keys);
}
const GCOL = ["#2f6fed", "#d64545", "#1f9d55"];
function graphPanel() {
  const rows = Calc.fns.map((f, i) => {
    const inp = el("input", {value: f, placeholder: i ? "ixtiyoriy, masalan 2x+1" : "masalan x^2-2x-3", spellcheck: "false"});
    inp.addEventListener("input", () => { Calc.fns[i] = inp.value; drawGraph(); });
    return el("div", {class: "fnrow"}, el("i", {style: "background:" + GCOL[i]}), el("span", {class: "small"}, "y ="), inp);
  });
  const cv = el("canvas", {id: "gcv"});
  const coord = el("div", {class: "muted small", id: "gcoord", style: "text-align:center;min-height:18px"}, "Grafikni sudrang, +/− bilan kattalashtiring");
  const zoom = k => { const v = Calc.view, cx = (v.x0 + v.x1) / 2, cy = (v.y0 + v.y1) / 2; v.x0 = cx - (cx - v.x0) * k; v.x1 = cx + (v.x1 - cx) * k; v.y0 = cy - (cy - v.y0) * k; v.y1 = cy + (v.y1 - cy) * k; drawGraph(); };
  let drag = null;
  cv.addEventListener("pointerdown", e => { drag = {x: e.clientX, y: e.clientY}; cv.setPointerCapture(e.pointerId); });
  cv.addEventListener("pointermove", e => {
    const r = cv.getBoundingClientRect(), v = Calc.view;
    const gx = v.x0 + (e.clientX - r.left) / r.width * (v.x1 - v.x0);
    if (drag) { const dx = (e.clientX - drag.x) / r.width * (v.x1 - v.x0), dy = (e.clientY - drag.y) / r.height * (v.y1 - v.y0); v.x0 -= dx; v.x1 -= dx; v.y0 += dy; v.y1 += dy; drag = {x: e.clientX, y: e.clientY}; drawGraph(); }
    let txt = `x = ${fmtNum(+gx.toFixed(3))}`;
    Calc.fns.forEach((f, i) => { if (!f.trim()) return; try { txt += `  ·  y${i + 1} = ${fmtNum(+parseExpr(tokenize(f))(gx).toFixed(4))}`; } catch (er) {} });
    coord.textContent = txt;
  });
  cv.addEventListener("pointerup", () => { drag = null; });
  return el("div", {class: "graph"}, ...rows, cv, coord, el("div", {class: "row-btns", style: "margin-top:6px"},
    el("button", {class: "btn ghost", onclick: () => zoom(0.7)}, "+"), el("button", {class: "btn ghost", onclick: () => zoom(1 / 0.7)}, "−"),
    el("button", {class: "btn ghost", onclick: () => { Calc.view = {x0: -10, x1: 10, y0: -7, y1: 7}; drawGraph(); }}, "⟲")));
}
function drawGraph() {
  const cv = document.getElementById("gcv"); if (!cv) return;
  const dpr = window.devicePixelRatio || 1, W = cv.clientWidth || 300, H = cv.clientHeight || 230;
  cv.width = W * dpr; cv.height = H * dpr;
  const g = cv.getContext("2d"); g.setTransform(dpr, 0, 0, dpr, 0, 0); g.clearRect(0, 0, W, H);
  const v = Calc.view, X = x => (x - v.x0) / (v.x1 - v.x0) * W, Y = y => H - (y - v.y0) / (v.y1 - v.y0) * H;
  const css = getComputedStyle(document.body), txt = css.color;
  const step = s => { const raw = s / 8, p = Math.pow(10, Math.floor(Math.log10(raw))), m = raw / p; return (m < 2 ? 1 : m < 5 ? 2 : 5) * p; };
  const sx = step(v.x1 - v.x0), sy = step(v.y1 - v.y0);
  g.lineWidth = 1; g.strokeStyle = "rgba(127,127,127,.18)"; g.fillStyle = txt; g.font = "10px sans-serif";
  for (let x = Math.ceil(v.x0 / sx) * sx; x <= v.x1; x += sx) { g.beginPath(); g.moveTo(X(x), 0); g.lineTo(X(x), H); g.stroke(); if (Math.abs(x) > 1e-9) g.fillText(fmtNum(+x.toFixed(6)), X(x) + 2, Math.min(H - 3, Math.max(10, Y(0) + 11))); }
  for (let y = Math.ceil(v.y0 / sy) * sy; y <= v.y1; y += sy) { g.beginPath(); g.moveTo(0, Y(y)); g.lineTo(W, Y(y)); g.stroke(); if (Math.abs(y) > 1e-9) g.fillText(fmtNum(+y.toFixed(6)), Math.min(W - 24, Math.max(2, X(0) + 3)), Y(y) - 2); }
  g.strokeStyle = "rgba(127,127,127,.75)"; g.lineWidth = 1.3;
  g.beginPath(); g.moveTo(0, Y(0)); g.lineTo(W, Y(0)); g.moveTo(X(0), 0); g.lineTo(X(0), H); g.stroke();
  Calc.fns.forEach((f, i) => {
    if (!f.trim()) return;
    let fn; try { fn = parseExpr(tokenize(f)); } catch (e) { return; }
    g.strokeStyle = GCOL[i]; g.lineWidth = 2.2; g.beginPath();
    let pen = false, prev = null;
    for (let px = 0; px <= W; px += 1) {
      const x = v.x0 + px / W * (v.x1 - v.x0), y = fn(x);
      if (!isFinite(y) || (prev != null && Math.abs(Y(y) - Y(prev)) > H * 2)) { pen = false; prev = isFinite(y) ? y : null; continue; }
      if (pen) g.lineTo(px, Y(y)); else { g.moveTo(px, Y(y)); pen = true; }
      prev = y;
    }
    g.stroke();
  });
}
function showReference() {
  const F = (name, f) => el("div", {class: "f"}, el("span", {}, name), el("b", {}, f));
  sheet(el("div", {class: "ref"}, el("h2", {}, "📐 SAT Math ma'lumotnomasi"),
    el("div", {class: "muted small"}, "Imtihonda beriladigan formulalar"),
    el("h3", {}, "Yuzalar"), F("Doira yuzasi", "A = πr²"), F("Aylana uzunligi", "C = 2πr"), F("To'g'ri to'rtburchak", "A = ℓw"), F("Uchburchak", "A = ½·b·h"),
    el("h3", {}, "Uchburchaklar"), F("Pifagor teoremasi", "c² = a² + b²"), F("30°–60°–90°", "x,  x√3,  2x"), F("45°–45°–90°", "s,  s,  s√2"),
    el("h3", {}, "Hajmlar"), F("To'g'ri burchakli parallelepiped", "V = ℓwh"), F("Silindr", "V = πr²h"), F("Shar", "V = ⁴⁄₃·πr³"), F("Konus", "V = ⅓·πr²h"), F("Piramida", "V = ⅓·ℓwh"),
    el("h3", {}, "Faktlar"), F("Aylanadagi yoy", "360° = 2π radian"), F("Uchburchak burchaklari yig'indisi", "180°")));
}

initOnce();
boot();
</script>
</body>
</html>
"""
# --- MINI APP HTML OXIRI ---


# ===================== FSM HOLATINI SAQLASH (Postgres) =====================
# ✅ aiogram standart bo'yicha FSM holatini (foydalanuvchi qaysi savolda
# turgani, belgilagan javoblari, admin sehrgarining qaysi qadamida ekani va
# h.k.) faqat OPERATIV XOTIRADA saqlaydi. Bot qayta ishga tushganda
# (server yangilanishi, xatolik tufayli qayta ko'tarilishi, deploy) bu
# xotira BUTUNLAY o'chib ketadi — o'sha payt test yechayotgan har bir odam
# hozirgacha belgilagan javoblarini yo'qotib qo'yardi (garchi taymer
# to'g'ri davom etsa ham). Bu klass xuddi shu holatni Postgres'ga yozadi,
# shuning uchun qayta ishga tushirilgandan keyin ham davom etadi. Tezlik
# uchun operativ xotirada nusxa (cache) saqlanadi — o'qish bazaga
# murojaat qilmaydi, faqat yozish bazaga ham yoziladi (write-through);
# ishga tushganda esa bazadagi hammasi xotiraga oldindan yuklanadi.
class PostgresStorage(BaseStorage):
    def __init__(self):
        self._cache: dict[str, dict] = {}

    @staticmethod
    def _key(key: StorageKey) -> str:
        return f"{key.bot_id}:{key.chat_id}:{key.user_id}:{key.thread_id}:{key.destiny}"

    async def preload(self) -> None:
        rows = await POOL.fetch("SELECT storage_key, state, data FROM fsm_storage")
        for r in rows:
            try:
                data = json.loads(r["data"]) if r["data"] else {}
            except json.JSONDecodeError:
                data = {}
            self._cache[r["storage_key"]] = {"state": r["state"], "data": data}
        logging.info(f"FSM holati Postgres'dan tiklandi: {len(self._cache)} ta faol sessiya")

    async def _persist(self, k: str, entry: dict) -> None:
        if entry["state"] is None and not entry["data"]:
            await POOL.execute("DELETE FROM fsm_storage WHERE storage_key=$1", k)
            self._cache.pop(k, None)
            return
        await POOL.execute("""
            INSERT INTO fsm_storage (storage_key, state, data, updated_ts)
            VALUES ($1, $2, $3, $4)
            ON CONFLICT (storage_key) DO UPDATE SET state=$2, data=$3, updated_ts=$4
        """, k, entry["state"], json.dumps(entry["data"], ensure_ascii=False), now_ts())

    async def set_state(self, key: StorageKey, state: StateType = None) -> None:
        state_str = state.state if isinstance(state, State) else state
        k = self._key(key)
        entry = self._cache.setdefault(k, {"state": None, "data": {}})
        entry["state"] = state_str
        await self._persist(k, entry)

    async def get_state(self, key: StorageKey) -> str | None:
        entry = self._cache.get(self._key(key))
        return entry["state"] if entry else None

    async def set_data(self, key: StorageKey, data: dict) -> None:
        k = self._key(key)
        entry = self._cache.setdefault(k, {"state": None, "data": {}})
        entry["data"] = dict(data)
        await self._persist(k, entry)

    async def get_data(self, key: StorageKey) -> dict:
        entry = self._cache.get(self._key(key))
        return dict(entry["data"]) if entry else {}

    async def close(self) -> None:
        pass


# ===================== MAIN =====================
async def main():
    global POOL
    logging.basicConfig(level=logging.INFO, stream=sys.stdout)

    POOL = await asyncpg.create_pool(dsn=DATABASE_URL, min_size=2, max_size=10)
    async with POOL.acquire() as db:
        await init_db(db)

    bot = Bot(token=BOT_TOKEN, default=DefaultBotProperties(parse_mode=ParseMode.HTML))
    _install_bot_session_middlewares(bot)

    await bot.set_my_commands(DEFAULT_BOT_COMMANDS, scope=BotCommandScopeDefault())
    await _sync_admin_commands(bot)

    asyncio.create_task(certificate_daemon(bot))
    asyncio.create_task(db_backup_daemon(bot))
    # ✅ WEBAPP_URL bo'sh bo'lsa tugma ko'rinmaydi — portni ham ochib qo'yish shart emas.
    if WEBAPP_URL:
        asyncio.create_task(run_webapp_server(bot))

    global FSM_STORAGE
    fsm_storage = PostgresStorage()
    await fsm_storage.preload()
    FSM_STORAGE = fsm_storage

    # ✅ Bitta foydalanuvchining update'lari KETMA-KET ishlanadi. aiogram
    # standart bo'yicha har bir update'ni alohida task'da parallel ishlaydi —
    # tugma ikki marta tez bosilsa (masalan "Bosqichni yakunlash", to'lovni
    # tasdiqlash, reklamani tasdiqlash) ikkala nusxa bir vaqtda bir xil eski
    # holatni o'qib, amalni ikki marta bajarardi.
    global EVENTS_ISOLATION
    EVENTS_ISOLATION = SimpleEventIsolation()
    dp = Dispatcher(storage=fsm_storage, events_isolation=EVENTS_ISOLATION)
    dp.include_router(router)
    dp.errors.register(global_error_handler)

    await _resume_active_timers(bot)

    try:
        await dp.start_polling(bot)
    finally:
        await POOL.close()


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print("Bot to'xtatildi")
