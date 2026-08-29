from __future__ import annotations

from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.colors import HexColor, black
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
import asyncio
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
from datetime import datetime, timezone, timedelta, date, time as dtime
from collections import defaultdict, deque

import asyncpg
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
    FSInputFile, BotCommand, BotCommandScopeDefault, BotCommandScopeChat, ErrorEvent
)
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.dispatcher.middlewares.base import BaseMiddleware
from aiogram.exceptions import TelegramRetryAfter, TelegramForbiddenError, TelegramBadRequest
from aiogram.utils.callback_answer import CallbackAnswerMiddleware

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

PAYMENT_CARD = _env("PAYMENT_CARD", "9860 0121 0681 0950")
PAYMENT_OWNER = _env("PAYMENT_OWNER", "Jasurbek Aktamov")
PAYMENT_PHONE = _env("PAYMENT_PHONE", "+998 93 224 47 30")
PAYMENT_NOTE = "To'lov izohiga test kodini yozing (masalan: MS-001)."

# ✅ Ma'lumotlar bazasidan avtomatik zaxira nusxa (backup) — necha soatda bir
# marta ADMIN_ID'ga JSON fayl qilib yuboriladi. Bu hosting ishonchliligidan
# mustaqil qo'shimcha xavfsizlik: hosting/server muammo qilsa ham, adminning
# Telegram xabarlar tarixida oxirgi holatning zaxirasi qoladi. 0 = o'chirilgan.
DB_BACKUP_INTERVAL_HOURS = int(_env("DB_BACKUP_INTERVAL_HOURS", "24") or "24")

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
BTN_MY = "📦 Mening testlarim"
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
BTN_BACK = "⬅️ Ortga"
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


class TestProcessState(StatesGroup):
    solving = State()


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
    return bool(re.fullmatch(r"[a-f\*]+", ans))  # * belgisi o'tkazib yuborilganlarga ruxsat


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


def is_certificate_eligible(exam_type: str, percent: float, grade: str | None) -> bool:
    """
    Oddiy/DTM testlar uchun: foiz CERT_MIN_PERCENT dan yuqori bo'lsa yetarli.
    Rasch/Maxsus testlar uchun: rasmiy daraja (grade) borligi shart, chunki
    ularda foiz va daraja jadvali bir xil bo'lmasligi mumkin (masalan, foiz
    50% dan yuqori bo'lib, lekin rasmiy daraja "yo'q" bo'lishi mumkin edi).
    """
    et = (exam_type or "simple")
    if et in ("rasch", "maxsus"):
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

# Rasch stabilizatsiya sozlamalari
RUSH_SMALL_N = 20  # <=20 bo'lsa "kichik guruh" rejimi
RUSH_LARGE_N = 50  # >=50 bo'lsa "katta guruh" rejimi (percentile ishlaydi)
RUSH_HARD_FALLBACK_N = 6  # <= 6 qatnashuvchi bo'lsa Rasch emas, oddiy foizdan ball

# p smoothing (Bayes) parametrlari: p = (correct + A) / (total + A + B)
RUSH_SMOOTH_A = 1.0
RUSH_SMOOTH_B = 1.0

# kichik guruhda theta->ball rescale uchun barqaror diapazon
RUSH_FIXED_LO = -3.0
RUSH_FIXED_HI = 3.0

# b_list ni haddan tashqari ketib qolmasligi uchun clamp
B_MIN = -4.0
B_MAX = 4.0

THETA_MIN = -6.0
THETA_MAX = 6.0


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


def _sigmoid(x: float) -> float:
    if x >= 35:
        return 1.0
    if x <= -35:
        return 0.0
    return 1.0 / (1.0 + math.exp(-x))


def _clamp(x: float, a: float, b: float) -> float:
    return a if x < a else b if x > b else x


def rasch_difficulty_from_p(p: float) -> float:
    p = _clamp(float(p), 0.01, 0.99)
    return math.log((1.0 - p) / p)


def rasch_difficulty_from_counts(correct: int, total: int, a: float = RUSH_SMOOTH_A, b: float = RUSH_SMOOTH_B) -> float:
    total = int(total or 0)
    correct = int(correct or 0)
    if total <= 0:
        return 0.0
    p = (correct + float(a)) / (total + float(a) + float(b))
    bi = rasch_difficulty_from_p(p)
    return _clamp(bi, B_MIN, B_MAX)


def choose_rescale_bounds(thetas: list[float]) -> tuple[float, float]:
    n = len(thetas)
    if n <= 0:
        return (RUSH_FIXED_LO, RUSH_FIXED_HI)

    thetas_sorted = sorted(thetas)

    def _pct(arr, q):
        idx = int(round((len(arr) - 1) * q))
        idx = max(0, min(len(arr) - 1, idx))
        return float(arr[idx])

    if n <= RUSH_SMALL_N:
        return (RUSH_FIXED_LO, RUSH_FIXED_HI)

    if n >= RUSH_LARGE_N:
        lo = _pct(thetas_sorted, 0.05)
        hi = _pct(thetas_sorted, 0.95)
        if hi <= lo + 1e-9:
            return (RUSH_FIXED_LO, RUSH_FIXED_HI)
        return (lo, hi)

    mean = sum(thetas) / n
    var = sum((x - mean) ** 2 for x in thetas) / max(1, n - 1)
    std = math.sqrt(var)
    lo = mean - 2.0 * std
    hi = mean + 2.0 * std
    lo = _clamp(lo, THETA_MIN, THETA_MAX)
    hi = _clamp(hi, THETA_MIN, THETA_MAX)
    if hi <= lo + 1e-9:
        return (RUSH_FIXED_LO, RUSH_FIXED_HI)
    return (lo, hi)


def estimate_theta_rasch(ans: str, key: str, b_list: list[float], iters: int = 25) -> float:
    theta = 0.0
    n = len(b_list)

    for _ in range(iters):
        g = 0.0
        h = 0.0
        for i in range(n):
            x = 1.0 if ans[i] == key[i] else 0.0
            p = _sigmoid(theta - b_list[i])
            g += (x - p)
            h -= (p * (1.0 - p))

        if abs(h) < 1e-9:
            break

        step = g / h
        theta -= step
        theta = _clamp(theta, THETA_MIN, THETA_MAX)

        if abs(step) < 1e-4:
            break

    return theta


def rescale_theta_to_75(theta: float, lo: float, hi: float) -> float:
    if hi <= lo + 1e-9:
        return 37.5
    x = (theta - lo) / (hi - lo)
    x = _clamp(x, 0.0, 1.0)
    return x * 75.0


def rasch_ball_from_weighted_ratio(ratio_0_1: float) -> float:
    ratio_0_1 = max(0.0, min(1.0, float(ratio_0_1 or 0.0)))
    return ratio_0_1 * 75.0


def fallback_rasch_ball_from_raw_percent(raw_percent_0_100: float) -> float:
    """Juda kichik N bo'lsa Raschni ishlatmaymiz, foizni chiziqli 0..75 shkalaga o'tkazamiz."""
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
    # beramiz (Rasch/Maxsus kutayotgan bo'lsa ham). Test qayta ochilsa,
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


async def finalize_rush_for_test_if_ready(test_id: int) -> bool:
    end_ts = await get_test_end_ts(test_id)
    if end_ts <= 0 or now_ts() < end_ts:
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

        rows = await db.fetch("""
            SELECT id, user_answers
            FROM results
            WHERE test_id=$1 AND user_answers IS NOT NULL AND LENGTH(user_answers) >= $2
        """, test_id, qcount)

        if not rows:
            return False

        solvers = []
        for r in rows:
            ans = normalize_answers(r["user_answers"] or "")
            if len(ans) >= qcount:
                solvers.append(ans)

        n_solvers = len(solvers)

        async with db.transaction():
            # === HARD FALLBACK ===
            if n_solvers <= RUSH_HARD_FALLBACK_N:
                for r in rows:
                    ans = normalize_answers(r["user_answers"] or "")
                    if len(ans) < qcount:
                        continue

                    score = 0
                    for i in range(qcount):
                        if ans[i] == correct_key[i]:
                            score += 1

                    raw_percent = (score / qcount) * 100.0
                    rasch_ball = fallback_rasch_ball_from_raw_percent(raw_percent)
                    rasch_percent = (rasch_ball / 75.0) * 100.0
                    grade = grade_from_rasch_ball(rasch_ball)

                    await db.execute("""
                        UPDATE results
                        SET rasch_ball=$1, rasch_percent=$2, grade=$3
                        WHERE id=$4
                    """, float(rasch_ball), float(rasch_percent), grade, int(r["id"]))

                await db.execute(
                    "UPDATE tests SET rush_finalized=1, rush_finalized_ts=$1 WHERE id=$2",
                    now_ts(), test_id
                )
                return True

            # --- item difficulty ---
            correct_counts = [0] * qcount
            total_counts = [0] * qcount

            for ans in solvers:
                for i in range(qcount):
                    total_counts[i] += 1
                    if ans[i] == correct_key[i]:
                        correct_counts[i] += 1

            b_list = []
            for i in range(qcount):
                if total_counts[i] == 0:
                    b_list.append(0.0)
                    continue
                if n_solvers <= RUSH_SMALL_N:
                    b_list.append(rasch_difficulty_from_counts(correct_counts[i], total_counts[i]))
                else:
                    p = correct_counts[i] / total_counts[i]
                    b_list.append(_clamp(rasch_difficulty_from_p(p), B_MIN, B_MAX))

            thetas = [estimate_theta_rasch(ans, correct_key, b_list) for ans in solvers]
            lo, hi = choose_rescale_bounds(thetas)

            for r in rows:
                ans = normalize_answers(r["user_answers"] or "")
                if len(ans) < qcount:
                    continue

                theta = estimate_theta_rasch(ans, correct_key, b_list)
                rasch_ball = rescale_theta_to_75(theta, lo, hi)
                rasch_percent = (rasch_ball / 75.0) * 100.0
                grade = grade_from_rasch_ball(rasch_ball)

                await db.execute("""
                    UPDATE results
                    SET rasch_ball=$1, rasch_percent=$2, grade=$3
                    WHERE id=$4
                """, float(rasch_ball), float(rasch_percent), grade, int(r["id"]))

            await db.execute(
                "UPDATE tests SET rush_finalized=1, rush_finalized_ts=$1 WHERE id=$2",
                now_ts(), test_id
            )
            return True


# ===================== MAXSUS MODULE (Rasch-like, but local) =====================

MAXSUS_RELIABILITY_M = 3  # n kichik bo'lsa natijani 50% ga yaqinlashtiradi


def _round_half_up(x: float) -> int:
    return int(math.floor(x + 0.5))


def maxsus_grade_from_k_any(n: int, k: int) -> str | None:
    n = int(n or 0)
    k = int(k or 0)
    if n <= 0:
        return None

    if n == 30:
        if 25 <= k <= 30: return "A+"
        if 22 <= k <= 24: return "A"
        if 18 <= k <= 21: return "B+"
        if 15 <= k <= 17: return "B"
        if 12 <= k <= 14: return "C+"
        if 9 <= k <= 11: return "C"
        return None

    if n == 35:
        if 29 <= k <= 35: return "A+"
        if 25 <= k <= 28: return "A"
        if 21 <= k <= 24: return "B+"
        if 18 <= k <= 20: return "B"
        if 14 <= k <= 17: return "C+"
        if 11 <= k <= 13: return "C"
        return None

    if n == 55:
        if 45 <= k <= 55: return "A+"
        if 40 <= k <= 44: return "A"
        if 34 <= k <= 39: return "B+"
        if 28 <= k <= 33: return "B"
        if 22 <= k <= 27: return "C+"
        if 17 <= k <= 21: return "C"
        return None

    c_min = _round_half_up(0.30 * n)
    cp_min = _round_half_up(0.40 * n)
    b_min = _round_half_up(0.50 * n)
    bp_min = _round_half_up((0.62 if n >= 50 else 0.60) * n)
    a_min = _round_half_up(0.72 * n)
    ap_min = _round_half_up(0.82 * n)

    cp_min = max(cp_min, c_min + 1)
    b_min = max(b_min, cp_min + 1)
    bp_min = max(bp_min, b_min + 1)
    a_min = max(a_min, bp_min + 1)
    ap_min = max(ap_min, a_min + 1)

    if k >= ap_min: return "A+"
    if k >= a_min:  return "A"
    if k >= bp_min: return "B+"
    if k >= b_min:  return "B"
    if k >= cp_min: return "C+"
    if k >= c_min:  return "C"
    return None


def maxsus_theta_from_score(k: int, n: int) -> float:
    k = int(k)
    n = int(n)
    if n <= 0:
        return 0.0
    return math.log((k + 0.5) / ((n - k) + 0.5))


def maxsus_rasch_p(theta: float) -> float:
    return _sigmoid(theta)


def maxsus_reliability(n: int, m: int = MAXSUS_RELIABILITY_M) -> float:
    n = int(n)
    m = int(m)
    if n <= 0:
        return 0.0
    return n / (n + m)


def maxsus_ball_percent_from_score(k: int, n: int) -> tuple[float, float, float]:
    theta = maxsus_theta_from_score(k, n)
    p = maxsus_rasch_p(theta)
    r = maxsus_reliability(n)
    p_final = 0.5 + r * (p - 0.5)
    p_final = max(0.0, min(1.0, p_final))

    rasch_ball = 75.0 * p_final
    rasch_percent = 100.0 * p_final
    return rasch_ball, rasch_percent, theta


async def finalize_maxsus_for_test_if_ready(test_id: int) -> bool:
    end_ts = await get_test_end_ts(test_id)
    if end_ts <= 0 or now_ts() < end_ts:
        return False

    async with POOL.acquire() as db:
        row = await db.fetchrow("SELECT maxsus_finalized FROM tests WHERE id=$1", test_id)
        if row and int(row["maxsus_finalized"] or 0) == 1:
            return True

        t = await db.fetchrow("SELECT questions_count, answers FROM tests WHERE id=$1", test_id)
        if not t:
            return False

        qcount = int(t["questions_count"] or 0)
        correct_key = normalize_answers(t["answers"] or "")
        qcount = min(qcount, len(correct_key))
        if qcount <= 0:
            return False

        rows = await db.fetch("""
            SELECT id, user_answers
            FROM results
            WHERE test_id=$1 AND user_answers IS NOT NULL AND LENGTH(user_answers) >= $2
        """, test_id, qcount)

        if not rows:
            return False

        async with db.transaction():
            for r in rows:
                ans = normalize_answers(r["user_answers"] or "")
                if len(ans) < qcount:
                    continue

                k = 0
                for i in range(qcount):
                    if ans[i] == correct_key[i]:
                        k += 1

                rasch_ball, rasch_percent, _theta = maxsus_ball_percent_from_score(k, qcount)
                grade = maxsus_grade_from_k_any(qcount, k)

                await db.execute("""
                    UPDATE results
                    SET rasch_ball=$1, rasch_percent=$2, grade=$3
                    WHERE id=$4
                """, float(rasch_ball), float(rasch_percent), grade, int(r["id"]))

            await db.execute(
                "UPDATE tests SET maxsus_finalized=1, maxsus_finalized_ts=$1 WHERE id=$2",
                now_ts(), test_id
            )

    return True


# ===================== REQUIRED CHANNELS (FORCED SUBSCRIPTION) =====================
def _normalize_channel_ref(text: str) -> str:
    s = (text or "").strip()
    if s.startswith("@"):
        s = s[1:]
    return s


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
    await _ensure_column(db, "tests", "maxsus_finalized", "INTEGER DEFAULT 0")
    await _ensure_column(db, "tests", "maxsus_finalized_ts", "BIGINT DEFAULT 0")
    await _ensure_column(db, "tests", "exam_type", "TEXT DEFAULT 'simple'")
    await _ensure_column(db, "tests", "dtm_cfg", "TEXT")

    await _ensure_column(db, "results", "earned_points", "DOUBLE PRECISION")
    await _ensure_column(db, "results", "max_points", "DOUBLE PRECISION")
    await _ensure_column(db, "results", "details", "TEXT")
    await _ensure_column(db, "results", "user_answers", "TEXT")
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
    await _ensure_column(db, "tests", "is_hidden", "INTEGER DEFAULT 0")  # vaqtincha yashirilgan test
    await _ensure_column(db, "tests", "is_closed", "INTEGER DEFAULT 0")  # admin/creator tomonidan yakunlangan test

    # ✅ Har bir test o'z sertifikat sozlamalariga ega bo'lishi mumkin (bo'sh
    # bo'lsa standart qiymatlar ishlatiladi — pastda make_certificate_pdf'ga
    # uzatilganda hal qilinadi).
    await _ensure_column(db, "tests", "cert_academy_name", "TEXT")
    await _ensure_column(db, "tests", "cert_signee_name", "TEXT")
    await _ensure_column(db, "tests", "cert_template", "TEXT DEFAULT 'gold'")


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


SECURITY = SecurityManager()


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
                        text="⏱ <b>Vaqt tugadi!</b>\nJavob yuborilmasa qabul qilinmaydi.",
                        parse_mode=ParseMode.HTML
                    )
                except Exception:
                    pass
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

            await asyncio.sleep(COUNTDOWN_TICK_SEC)
    finally:
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
    await state.clear()
    admin_flag = await is_admin_user(message.from_user.id)
    await message.answer("🏠 Bosh menyu", reply_markup=main_menu_kb(admin_flag))


@router.message(StateFilter("*"), F.text == BTN_HOME)
async def nav_home(message: Message, state: FSMContext):
    await go_home(message, state)


@router.message(StateFilter("*"), F.text == BTN_BACK)
async def nav_back(message: Message, state: FSMContext):
    await go_home(message, state)


@router.message(StateFilter("*"), F.text.in_(SECTION_BTNS))
async def nav_any_section(message: Message, state: FSMContext, bot: Bot):
    if not await ensure_subscribed_message(message, bot):
        return

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
    await state.clear()
    admin_flag = await is_admin_user(message.from_user.id)
    await message.answer("Bekor qilindi.", reply_markup=main_menu_kb(admin_flag))


# ===================== /start + REGISTRATION =====================
@router.message(CommandStart())
async def command_start(message: Message, state: FSMContext, bot: Bot):
    if not await ensure_subscribed_message(message, bot):
        return

    user = await POOL.fetchrow("SELECT * FROM users WHERE telegram_id=$1", message.from_user.id)

    if user:
        admin_flag = await is_admin_user(message.from_user.id)
        await message.answer(f"👋 Xush kelibsiz, <b>{user['full_name']}</b>!", parse_mode=ParseMode.HTML,
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
    contact = message.contact.phone_number if message.contact else sanitize_plain_text(message.text, max_len=30)
    data = await state.get_data()
    role = 'admin' if message.from_user.id == ADMIN_ID else 'student'

    await POOL.execute(
        "INSERT INTO users (telegram_id, full_name, phone, role, created_ts) VALUES ($1, $2, $3, $4, $5) "
        "ON CONFLICT (telegram_id) DO NOTHING",
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
                text += f"📝 Sabab: <i>{reject_reason}</i>\n"
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
                    [InlineKeyboardButton(text="📦 Mening testlarim", callback_data="open_mytests")])
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
def test_ui_kb(current_q: int, qcount: int, answers_dict: dict):
    kb = InlineKeyboardMarkup(inline_keyboard=[])
    current_ans = answers_dict.get(str(current_q))

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

    await state.update_data(test_id=test_id, mode="free", qcount=qcount, current_q=1, test_answers={})
    await bot.send_message(
        chat_id=callback.from_user.id,
        text=(
            f"📝 <b>Test boshlandi!</b>\n\n"
            f"Savol: <b>1 / {qcount}</b>\n"
            f"Pastdagi tugmalar orqali javoblarni interaktiv belgilang.\n"
            f"<i>(Yoki xohlasangiz odatdagidek abcd... formatida birdaniga yozib yuboring)</i>"
        ),
        parse_mode=ParseMode.HTML,
        reply_markup=test_ui_kb(1, int(qcount), {})
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
                [InlineKeyboardButton(text=f"📦 Mening testlarim: {code}", callback_data="open_mytests")])
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


@router.message(F.text == BTN_PAID)
async def show_paid_tests(message: Message):
    await _render_paid_tests(message, message.from_user.id, 0)


@router.callback_query(F.data == "open_mytests")
async def open_mytests_cb(callback: CallbackQuery):
    await callback.message.answer("📦 Mening testlarim bo‘limiga o‘ting (menyudan).")


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

    await state.update_data(test_id=test_id)
    await callback.message.answer(
        "💳 <b>To'lov ma'lumotlari:</b>\n\n"
        f"💳 Karta: <code>{PAYMENT_CARD}</code>\n"
        f"👤 Ega: <b>{PAYMENT_OWNER}</b>\n"
        f"📞 Telefon: <code>{PAYMENT_PHONE}</code>\n"
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

    await state.update_data(test_id=test_id)
    await callback.message.answer(
        "💳 <b>To'lov ma'lumotlari:</b>\n\n"
        f"💳 Karta: <code>{PAYMENT_CARD}</code>\n"
        f"👤 Ega: <b>{PAYMENT_OWNER}</b>\n"
        f"📞 Telefon: <code>{PAYMENT_PHONE}</code>\n"
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

    try:
        payment_id = await POOL.fetchval("""
            INSERT INTO payments (user_id, test_id, screenshot_id, status, created_at)
            VALUES ($1, $2, $3, 'pending', $4)
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

    await bot.send_photo(
        chat_id=ADMIN_ID,
        photo=photo_id,
        caption=(
            f"🟡 <b>Yangi to'lov (PENDING)</b>\n\n"
            f"👤 User: {uname}\n"
            f"📚 Test: {title}\n"
            f"🔑 Kod: <code>{code}</code>\n"
            f"💰 Summa: {price} so'm\n"
            f"🆔 Payment ID: <code>{payment_id}</code>"
        ),
        parse_mode=ParseMode.HTML,
        reply_markup=admin_kb
    )

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
                text += f" — {reject_reason}"
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


@router.callback_query(F.data.startswith("viewres_"))
async def view_result(callback: CallbackQuery):
    test_id = int(callback.data.split("_")[1])
    user_db_id = await get_user_db_id(callback.from_user.id)
    if not user_db_id:
        return await callback.message.answer("Avval /start orqali ro‘yxatdan o‘ting.")

    row = await POOL.fetchrow("""
        SELECT t.code, t.title, t.subject, t.is_free, t.price, t.exam_type,
               r.score, r.total_questions, r.percent,
               r.earned_points, r.max_points, r.details,
               r.rasch_ball, r.rasch_percent, r.grade, r.date
        FROM results r
        JOIN tests t ON t.id=r.test_id
        WHERE r.user_id=$1 AND r.test_id=$2
        ORDER BY r.id ASC LIMIT 1
    """, user_db_id, test_id)

    if not row:
        return await callback.message.answer("Bu test bo‘yicha natija topilmadi.")

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

    if exam_type in ("rasch", "maxsus"):
        if exam_type == "rasch":
            await finalize_rush_for_test_if_ready(test_id)
        else:
            await finalize_maxsus_for_test_if_ready(test_id)

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

    base_block = (
        f"📄 <b>Natija (1-urinish)</b>\n\n"
        f"🔑 Kod: <code>{code}</code>\n"
        f"📚 Fan: <b>{subject or 'Noma’lum'}</b>\n"
        f"📚 Test: <b>{title}</b>\n"
        f"{pay_line}"
        f"✅ To‘g‘ri: {score}/{total}\n"
        f"❌ Xato: {int(total or 0) - int(score or 0)}\n"
        f"📊 Oddiy foiz: <b>{float(percent):.1f}%</b>\n"
    )

    dtm_block = ""
    if exam_type == "dtm":
        dtm_block = (
            f"\n🏁 Ball: <b>{float(earned_points or 0):.1f}/{float(max_points or 0):.1f}</b>\n"
            f"📌 Ball foiz: <b>{float(percent):.1f}%</b>\n"
        )

    rasch_block = ""
    if exam_type in ("rasch", "maxsus"):
        if rasch_ball is None:
            rasch_block = (
                "\n⚡ Rasch natija: <b>Hali hisoblanmadi (test tugagach chiqadi)</b>\n"
                "🎯 Daraja: <b>Test tugagach chiqadi</b>\n"
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
        base_block + dtm_block + rasch_block + rank_block + f"\n🗓 Sana: {date_}",
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
        SELECT r.user_answers, t.answers, t.questions_count
        FROM results r JOIN tests t ON t.id=r.test_id
        WHERE r.user_id=$1 AND r.test_id=$2
        ORDER BY r.id ASC LIMIT 1
    """, user_db_id, test_id)

    if not row or not row["user_answers"]:
        return await callback.message.answer("❌ Bu test bo'yicha javoblaringiz topilmadi.")

    user_ans = normalize_answers(row["user_answers"])
    correct_key = normalize_answers(row["answers"])
    qcount = min(int(row["questions_count"] or 0), len(correct_key), len(user_ans))

    mistakes = []
    for i in range(qcount):
        if user_ans[i] != correct_key[i]:
            miss_char = "belgilamagansiz" if user_ans[i] == "*" else f"siz: {user_ans[i].upper()}"
            mistakes.append(f"{i + 1}-savol: {miss_char}, to'g'ri javob: <b>{correct_key[i].upper()}</b>")

    if not mistakes:
        return await callback.message.answer("🎉 Xatoyingiz yo'q — barcha savollarga to'g'ri javob bergansiz!")

    text = f"🔎 <b>Xatolaringiz</b> ({len(mistakes)} ta):\n\n" + "\n".join(mistakes[:40])
    if len(mistakes) > 40:
        text += f"\n... va yana {len(mistakes) - 40} ta"
    await callback.message.answer(text, parse_mode=ParseMode.HTML)


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
               t.start_mode, t.start_ts, t.start_at
        FROM payments p
        JOIN tests t ON t.id=p.test_id
        WHERE p.id=$1 AND p.user_id=$2
    """, payment_id, user_db_id)

    if not row:
        return await callback.message.answer("Xatolik: payment topilmadi.")

    (status, started_at_paid, started_ts_paid, test_id,
     code, title, file_id, duration_min, qcount, answers,
     start_mode, start_ts_val, start_at) = (
        row["status"], row["started_at"], row["started_ts"], row["test_id"],
        row["code"], row["title"], row["file_id"], row["duration"], row["questions_count"], row["answers"],
        row["start_mode"], row["start_ts"], row["start_at"]
    )

    if status != "approved":
        return await callback.message.answer("Bu test hali tasdiqlanmagan.")

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

    await state.update_data(test_id=test_id, mode="paid", qcount=qcount, current_q=1, test_answers={})
    await bot.send_message(
        chat_id=callback.from_user.id,
        text=(
            f"📝 <b>Test boshlandi!</b>\n\n"
            f"Savol: <b>1 / {qcount}</b>\n"
            f"Pastdagi tugmalar orqali javoblarni interaktiv belgilang.\n"
            f"<i>(Yoki xohlasangiz odatdagidek abcd... formatida birdaniga yozib yuboring)</i>"
        ),
        parse_mode=ParseMode.HTML,
        reply_markup=test_ui_kb(1, int(qcount), {})
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


def dtm_score_points(user_ans: str, key: str, cfg_json: str) -> tuple[float, float, dict]:
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

        correct = 0
        for i in range(q):
            if part_user[i] == part_key[i]:
                correct += 1

        earned_s = correct * w
        max_s = q * w

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


# ===================== INTERACTIVE & TEXT SUBMISSION LOGIC =====================
async def process_test_submission(user_tg_id: int, bot: Bot, state: FSMContext, user_ans: str):
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
               cert_academy_name, cert_signee_name, cert_template, creator_id
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

    score = 0
    mistakes = []
    for i in range(qcount):
        if user_ans[i] == correct_key[i]:
            score += 1
        else:
            miss_char = "Belgilangan: yo'q" if user_ans[i] == "*" else f"Siz: {user_ans[i]}"
            mistakes.append(f"{i + 1}-savol: {miss_char} | to‘g‘ri={correct_key[i]}")

    percent = (score / qcount) * 100

    earned_points = float(score)
    max_points = float(qcount)
    details = None

    if (exam_type or "simple") == "dtm":
        earned_points, max_points, details_obj = dtm_score_points(user_ans, correct_key, dtm_cfg or "")
        details = json.dumps(details_obj, ensure_ascii=False)
        percent_points = (earned_points / max_points) * 100 if max_points > 0 else 0.0
    else:
        percent_points = percent

    try:
        await POOL.execute("""
            INSERT INTO results (
                user_id, test_id, score, total_questions, percent,
                user_answers,
                earned_points, max_points, details,
                rasch_ball, rasch_percent, grade,
                finished_ts, date
            )
            VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, NULL, NULL, NULL, $10, $11)
        """,
            user_db_id, test_id,
            score, qcount,
            float(percent_points),
            user_ans,
            float(earned_points), float(max_points), details,
            now_ts(),
            now_str_local()
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
    can_immediate = ((start_mode or "normal") == "normal") and (et in ("simple", "dtm"))
    if can_immediate:
        eligible = is_certificate_eligible(et, percent_points, None)
        if eligible:
            try:
                user_full_name = await get_user_full_name(user_db_id)
                test_code = code
                test_title = title
                test_subject = subject or "Noma'lum"
                exam_date = now_str_local()

                tmp_dir = tempfile.gettempdir()
                safe_code = re.sub(r"[^A-Z0-9\-_]+", "_", str(test_code).upper())
                pdf_path = os.path.join(tmp_dir, f"cert_{safe_code}_{user_db_id}.pdf")

                cert_academy, cert_signee, cert_tpl = await _resolve_cert_settings(
                    t["cert_academy_name"], t["cert_signee_name"], t["cert_template"], t["creator_id"])

                make_certificate_pdf(
                    full_name=user_full_name,
                    subject=test_subject,
                    test_code=test_code,
                    test_title=test_title,
                    exam_date_str=exam_date,
                    percent=float(percent_points),
                    correct_count=score,
                    total_questions=qcount,
                    out_path=pdf_path,
                    academy_name=cert_academy,
                    signee_name=cert_signee,
                    template=cert_tpl,
                )

                sent = await bot.send_document(
                    chat_id=user_tg_id,
                    document=FSInputFile(pdf_path),
                    caption="🎁 Sertifikatingiz tayyor! (PDF)"
                )
                cert_file_id = sent.document.file_id if sent and sent.document else None
                await POOL.execute("""
                    UPDATE results
                    SET certificate_sent=1, certificate_file_id=$1, certificate_sent_ts=$2
                    WHERE user_id=$3 AND test_id=$4
                """, cert_file_id, now_ts(), user_db_id, test_id)
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

    mistakes_text = "\n".join(mistakes[:25])
    if len(mistakes) > 25:
        mistakes_text += "\n..."

    exam_type = (exam_type or "simple")
    extra_points_line = ""
    if exam_type == "dtm":
        extra_points_line = (
            f"🏁 Ball: <b>{earned_points:.1f}/{max_points:.1f}</b>\n"
            f"📌 Ball foiz: <b>{percent_points:.1f}%</b>\n"
        )

    rasch_pending_block = ""
    if exam_type in ("rasch", "maxsus"):
        rasch_pending_block = (
            "⚡ Rasch foiz: <b>Test tugagach avtomatik hisoblanadi</b>\n"
            "⚡ Rasch ball: <b>Test tugagach avtomatik hisoblanadi</b>\n"
            "🎯 Daraja: <b>Test tugagach chiqadi</b>\n"
            "📌 Rasch ball ko‘rish uchun: <b>“📄 Natijani ko‘rish”</b> tugmasini bosing.\n"
        )

    result_text = (
        f"🏁 <b>Test yakunlandi!</b>\n\n"
        f"🔑 Kod: <code>{code}</code>\n"
        f"📚 Test: <b>{title}</b>\n"
        f"🧮 Savollar: {qcount}\n\n"
        f"✅ To‘g‘ri: <b>{score}/{qcount}</b>\n"
        f"❌ Xato: <b>{qcount - score}</b>\n"
        f"📊 Oddiy foiz: <b>{percent:.1f}%</b>\n"
        f"{extra_points_line}"
        f"{rasch_pending_block}\n"
        f"🔎 <b>Xatolar (qisqa):</b>\n"
        f"{mistakes_text if mistakes_text else 'Xatolar yo‘q ✅'}\n\n"
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

    await process_test_submission(message.from_user.id, bot, state, user_ans)


# Yangi interaktiv (Tugma) formati uchun
@router.callback_query(TestProcessState.solving, F.data.startswith("t_"))
async def interactive_test_handler(callback: CallbackQuery, state: FSMContext, bot: Bot):
    data = await state.get_data()
    current_q = data.get("current_q", 1)
    qcount = int(data.get("qcount", 30))
    answers = data.get("test_answers", {})

    action = callback.data

    if action == "t_nav_noop":
        return await callback.answer(f"Siz hozir {current_q}-savoldasiz", show_alert=False)

    elif action.startswith("t_ans_"):
        ans_letter = action.split("_")[2]
        answers[str(current_q)] = ans_letter
        await state.update_data(test_answers=answers)

        if current_q < qcount:
            current_q += 1
            await state.update_data(current_q=current_q)
            try:
                await callback.message.edit_text(
                    text=f"📝 <b>Test davom etmoqda...</b>\n\nSavol: <b>{current_q} / {qcount}</b>\nPastdagi tugmalar orqali javobni belgilang:",
                    parse_mode=ParseMode.HTML,
                    reply_markup=test_ui_kb(current_q, qcount, answers)
                )
            except Exception:
                pass
        else:
            try:
                await callback.message.edit_reply_markup(reply_markup=test_ui_kb(current_q, qcount, answers))
            except Exception:
                pass

        await callback.answer()

    elif action == "t_nav_prev":
        if current_q > 1:
            current_q -= 1
            await state.update_data(current_q=current_q)
            try:
                await callback.message.edit_text(
                    text=f"📝 <b>Test davom etmoqda...</b>\n\nSavol: <b>{current_q} / {qcount}</b>\nPastdagi tugmalar orqali javobni belgilang:",
                    parse_mode=ParseMode.HTML,
                    reply_markup=test_ui_kb(current_q, qcount, answers)
                )
            except Exception:
                pass
        await callback.answer()

    elif action == "t_nav_next":
        if current_q < qcount:
            current_q += 1
            await state.update_data(current_q=current_q)
            try:
                await callback.message.edit_text(
                    text=f"📝 <b>Test davom etmoqda...</b>\n\nSavol: <b>{current_q} / {qcount}</b>\nPastdagi tugmalar orqali javobni belgilang:",
                    parse_mode=ParseMode.HTML,
                    reply_markup=test_ui_kb(current_q, qcount, answers)
                )
            except Exception:
                pass
        await callback.answer()

    elif action == "t_submit":
        user_ans = ""
        for i in range(1, qcount + 1):
            user_ans += answers.get(str(i), "*")

        await callback.message.delete_reply_markup()
        await callback.answer("Javoblaringiz qabul qilindi, baholanmoqda...", show_alert=False)
        await process_test_submission(callback.from_user.id, bot, state, user_ans)


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
        SELECT rk.user_id, rk.avgp, rk.cnt, rk.rnk, u.full_name
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
        SELECT u.full_name, MAX(r.percent) AS bestp
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
async def rating_per_test(callback: CallbackQuery):
    test_id = int(callback.data.split("_")[2])

    t = await POOL.fetchrow("SELECT code, title, exam_type FROM tests WHERE id=$1", test_id)
    if not t:
        return await callback.message.answer("Test topilmadi.")

    et = (t["exam_type"] or "simple")
    if et == "rasch":
        await finalize_rush_for_test_if_ready(test_id)
    elif et == "maxsus":
        await finalize_maxsus_for_test_if_ready(test_id)

    rows = await POOL.fetch("""
        SELECT u.full_name, r.percent
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
            f"👤 {uname}\n"
            f"🔑 <code>{code}</code> — <b>{title}</b>\n"
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
            f"👤 User: {uname}\n"
            f"📚 Test: {title}\n"
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

    await POOL.execute("UPDATE payments SET status='approved', reject_reason=NULL WHERE id=$1", payment_id)

    row = await POOL.fetchrow("""
        SELECT u.telegram_id, t.code, t.title, t.duration, t.questions_count, t.start_mode, t.start_ts, t.start_at
        FROM payments p
        JOIN users u ON u.id=p.user_id
        JOIN tests t ON t.id=p.test_id
        WHERE p.id=$1
    """, payment_id)

    if not row:
        return await callback.message.answer("Xatolik: payment topilmadi.")

    user_tg_id, code, title, duration, qcount, start_mode, start_ts_val, start_at = (
        row["telegram_id"], row["code"], row["title"], row["duration"], row["questions_count"],
        row["start_mode"], row["start_ts"], row["start_at"]
    )
    expired, after = is_scheduled_expired(start_mode, start_ts_val, duration)
    if expired:
        await callback.message.answer("❌ Bu test vaqti tugagan. Tasdiqlab bo‘lmaydi.")
        return

    start_kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="🚀 Testni boshlash", callback_data=f"begin_{payment_id}")]
    ])

    extra = ""
    if (start_mode or "normal") == "scheduled":
        extra = f"\n📅 Boshlanish: <b>{start_at}</b>\n⚠️ PDF faqat vaqt yetganda beriladi."

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

    await callback.message.answer(f"✅ Tasdiqlandi. Payment ID: {payment_id}")


@router.callback_query(F.data.startswith("rejask_"))
async def reject_ask_reason(callback: CallbackQuery, state: FSMContext):
    if not await is_admin_user(callback.from_user.id):
        return
    payment_id = int(callback.data.split("_")[1])
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
    reason = (message.text or "").strip()
    if not reason:
        return await message.answer("Sabab bo‘sh bo‘lmasin. Qayta yozing:")

    await POOL.execute("UPDATE payments SET status='rejected', reject_reason=$1 WHERE id=$2", reason, payment_id)

    row = await POOL.fetchrow("""
        SELECT u.telegram_id, t.code, t.title, t.price
        FROM payments p
        JOIN users u ON u.id=p.user_id
        JOIN tests t ON t.id=p.test_id
        WHERE p.id=$1
    """, payment_id)

    if row:
        user_tg_id, code, title, price = row["telegram_id"], row["code"], row["title"], row["price"]
        await bot.send_message(
            chat_id=user_tg_id,
            text=(
                f"❌ <b>To‘lov rad etildi</b>\n\n"
                f"📚 Test: <b>{title}</b>\n"
                f"🔑 Kod: <code>{code}</code>\n"
                f"💰 Summa: {price} so'm\n"
                f"📝 Sabab: <i>{reason}</i>\n\n"
                f"📦 'Mening testlarim' bo‘limidan qayta to‘lov yuborishingiz mumkin."
            ),
            parse_mode=ParseMode.HTML
        )

    await message.answer(f"❌ Rad etildi. Payment ID: {payment_id}")
    await state.clear()


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

    perc_rows = await POOL.fetch("SELECT percent FROM results")
    perc = [float(x["percent"]) for x in perc_rows]
    overall_avg = (sum(perc) / len(perc)) if perc else 0.0

    best = await POOL.fetchrow("""
        SELECT u.full_name, t.code, t.title, r.percent, r.date
        FROM results r
        JOIN users u ON u.id=r.user_id
        JOIN tests t ON t.id=r.test_id
        ORDER BY r.percent DESC
        LIMIT 1
    """)

    total_revenue = await POOL.fetchval("""
        SELECT COALESCE(SUM(t.price), 0) FROM payments p JOIN tests t ON t.id=p.test_id WHERE p.status='approved'
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
        [InlineKeyboardButton(text="📅 Bugungi hisobot", callback_data="today_digest"),
         InlineKeyboardButton(text="🏆 Testlar reytingi", callback_data="tests_ranking")],
        [InlineKeyboardButton(text="🛑 Testlarni yakunlash/qayta ochish", callback_data="closetest_menu")],
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


@router.callback_query(F.data == "xl_menu")
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
            await finalize_rush_for_test_if_ready(test_id)
        elif et == "maxsus":
            await finalize_maxsus_for_test_if_ready(test_id)

    t = await POOL.fetchrow("""
        SELECT code, title, is_free, price, duration, questions_count, start_mode, start_at, exam_type
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
               r.date
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
    ws_res.append(
        ["FullName", "Phone", "TelegramID", "Score", "Total", "Percent", "RaschBall", "RaschPercent", "Grade", "Date"])

    for r in results:
        ws_res.append([_xlsx_safe(r["full_name"]), _xlsx_safe(r["phone"]), r["telegram_id"], r["score"],
                        r["total_questions"], r["percent"], r["rasch_ball"], r["rasch_percent"], r["grade"], r["date"]])

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
    wb.save(path)

    await bot.send_document(
        chat_id=callback.from_user.id,
        document=FSInputFile(path),
        caption=f"📤 Excel eksport tayyor.\n🔑 <code>{code}</code> — <b>{title}</b>",
        parse_mode=ParseMode.HTML
    )


@router.callback_query(F.data.startswith("raschrep_"))
async def rasch_report_text(callback: CallbackQuery):
    test_id = int(callback.data.split("_")[1])
    user_tg_id = callback.from_user.id
    is_admin = await is_admin_user(user_tg_id)

    c_row = await POOL.fetchrow("SELECT creator_id FROM tests WHERE id=$1", test_id)
    is_creator = c_row and c_row["creator_id"] == user_tg_id

    if not (is_admin or is_creator):
        return await callback.answer("Sizda ruxsat yo'q!", show_alert=True)

    await finalize_rush_for_test_if_ready(test_id)

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


@router.callback_query(F.data == "trange_menu")
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
    wb.save(path)

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
    if _looks_like_int(q):
        rows = await POOL.fetch("SELECT id, full_name, telegram_id FROM users WHERE telegram_id=$1", int(q))
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
        SELECT COALESCE(SUM(t.price),0) FROM payments p JOIN tests t ON t.id=p.test_id
        WHERE p.user_id=$1 AND p.status='approved'
    """, uid)
    reg_at = datetime.fromtimestamp(u["created_ts"], TZ).strftime("%Y-%m-%d %H:%M") if u["created_ts"] else "noma'lum"

    text = (
        f"👤 <b>{u['full_name']}</b>\n\n"
        f"🆔 Telegram ID: <code>{u['telegram_id']}</code>\n"
        f"{username_line}"
        f"📞 Telefon: {u['phone'] or '-'}\n"
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


@router.callback_query(F.data == "broadcast_cancel")
async def broadcast_cancel(callback: CallbackQuery, state: FSMContext):
    await state.clear()
    await callback.message.answer("Bekor qilindi.")


@router.callback_query(F.data == "broadcast_confirm")
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
        try:
            await bot.copy_message(chat_id=r["telegram_id"], from_chat_id=src_chat_id, message_id=src_message_id)
            sent += 1
        except Exception:
            failed += 1
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

    total = await POOL.fetchval("""
        SELECT COALESCE(SUM(t.price),0) FROM payments p JOIN tests t ON t.id=p.test_id WHERE p.status='approved'
    """)
    today_start = (datetime.now(TZ).strftime("%Y-%m-%d") + " 00:00")
    today = await POOL.fetchval("""
        SELECT COALESCE(SUM(t.price),0) FROM payments p JOIN tests t ON t.id=p.test_id
        WHERE p.status='approved' AND p.created_at >= $1
    """, today_start)
    week_start = (datetime.now(TZ) - timedelta(days=7)).strftime("%Y-%m-%d %H:%M")
    week = await POOL.fetchval("""
        SELECT COALESCE(SUM(t.price),0) FROM payments p JOIN tests t ON t.id=p.test_id
        WHERE p.status='approved' AND p.created_at >= $1
    """, week_start)

    top_tests = await POOL.fetch("""
        SELECT t.code, t.title, COUNT(*) AS cnt, t.price
        FROM payments p JOIN tests t ON t.id=p.test_id
        WHERE p.status='approved'
        GROUP BY t.id, t.code, t.title, t.price
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
            text += f"• <code>{r['code']}</code> {r['title']} — {r['cnt']} ta ({_fmt_money(r['cnt'] * r['price'])})\n"

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
        "SELECT id, code, title, subject, price, duration, is_free, questions_count FROM tests WHERE id=$1", test_id)
    if not t:
        return await callback.answer("Test topilmadi.", show_alert=True)

    await _send_edit_field_picker(callback.message, state, t)


async def _send_edit_field_picker(answerable, state: FSMContext, t):
    await state.update_data(edit_test_id=t["id"])
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=f"📚 Fan: {t['subject'] or '-'}", callback_data="edtf_subject")],
        [InlineKeyboardButton(text=f"📝 Sarlavha: {t['title']}", callback_data="edtf_title")],
        [InlineKeyboardButton(text=f"⏳ Davomiylik: {t['duration']} daq", callback_data="edtf_duration")],
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
    "answers": "Javob kalitini (A-F harflarida, bo'shliqsiz)",
}


@router.callback_query(AdminEditTestState.waiting_field, F.data.startswith("edtf_"))
async def edittest_pick_field(callback: CallbackQuery, state: FSMContext):
    field = callback.data.split("_", 1)[1]
    if field not in _EDIT_FIELD_LABELS:
        return
    await state.update_data(edit_field=field)
    await callback.message.answer(f"Yangi qiymatni kiriting — {_EDIT_FIELD_LABELS[field]}:")
    await state.set_state(AdminEditTestState.waiting_value)


@router.message(AdminEditTestState.waiting_value)
async def edittest_save(message: Message, state: FSMContext):
    if (message.text or "") in TOP_MENU_BTNS:
        return

    data = await state.get_data()
    test_id = data.get("edit_test_id")
    field = data.get("edit_field")
    value_raw = (message.text or "").strip()

    if field not in _EDIT_FIELD_LABELS or not test_id:
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

    if field in ("duration", "price", "questions_count"):
        if not value_raw.isdigit():
            return await message.answer("⚠️ Faqat raqam kiriting:")
        value = int(value_raw)
        if field == "duration" and (value < 1 or value > 600):
            return await message.answer("⚠️ Davomiylik 1..600 daqiqa oralig'ida bo'lsin:")
        if field == "questions_count" and (value < 1 or value > 300):
            return await message.answer("⚠️ Savollar soni 1..300 oralig'ida bo'lsin:")
        # field allaqachon ("duration","price","questions_count") bilan cheklangani uchun xavfsiz
        await POOL.execute(f"UPDATE tests SET {field}=$1 WHERE id=$2", value, test_id)
        if field == "questions_count":
            await message.answer(
                "⚠️ <b>Diqqat:</b> savollar sonini o'zgartirdingiz — eski javob kaliti endi uzunligi "
                "mos kelmasligi mumkin. Zarur bo'lsa «✏️ Javob kaliti»ni ham shu yerdan yangilang.",
                parse_mode=ParseMode.HTML
            )
    elif field == "answers":
        val = normalize_answers(value_raw)
        if not is_answer_string_valid(val):
            return await message.answer(
                "⚠️ Javob kaliti faqat <b>A, B, C, D, E, F</b> harflaridan iborat bo'lishi kerak:",
                parse_mode=ParseMode.HTML)
        current_q = await POOL.fetchval("SELECT questions_count FROM tests WHERE id=$1", test_id)
        if len(val) != int(current_q or 0):
            return await message.answer(
                f"⚠️ Javoblar soni <b>{current_q}</b> ta bo'lishi kerak, siz {len(val)} ta yubordingiz. "
                f"Qayta kiriting:", parse_mode=ParseMode.HTML)
        await POOL.execute("UPDATE tests SET answers=$1 WHERE id=$2", val, test_id)
    else:
        max_len = 150 if field == "title" else 80
        value_clean = sanitize_plain_text(value_raw, max_len=max_len)
        if len(value_clean) < 2:
            return await message.answer("⚠️ Juda qisqa yoki noto'g'ri belgilar bor. Qayta kiriting:")
        await POOL.execute(f"UPDATE tests SET {field}=$1 WHERE id=$2", value_clean, test_id)

    await state.clear()
    await message.answer("✅ Test yangilandi.", reply_markup=main_menu_kb(admin_flag))


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
        SELECT COALESCE(SUM(t.price),0) FROM payments p JOIN tests t ON t.id=p.test_id
        WHERE p.status='approved' AND p.created_at >= $1
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
                f"<b>{title or 'Nomsiz'}</b> | <code>{ref}</code>"
                f"{warn_line}"
            )
            if link:
                text += f"\n🔗 {link}"
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

    admin_ok = await _bot_admin_status_in_chat(bot, chat_id)
    if admin_ok is not True:
        await message.answer(
            f"⚠️ <b>Diqqat:</b> bot <b>{title}</b> ichida hozircha administrator emas "
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
            f"✅ Qo'shildi: <b>{title}</b>\n🔗 Havola avtomatik aniqlandi: {auto_join_url}",
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

    if join_url and not join_url.startswith("http"):
        await message.answer("⚠️ Havola noto'g'ri ko'rinishda. Masalan: https://t.me/+AbCdEfGh. Qayta yuboring:")
        return

    data = await state.get_data()
    chat_id = int(data["chat_id"])
    username = data.get("username")
    title = data.get("title")

    await _save_required_channel(chat_id, username, title, join_url or None)

    await state.clear()
    await message.answer(f"✅ Qo'shildi: <b>{title}</b>", parse_mode=ParseMode.HTML, reply_markup=main_menu_kb(True))


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
        if not new_link_raw.startswith("http"):
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

    await message.answer(
        "🆕 <b>Yangi test yaratish</b>\n\n"
        "1-qadam: bu test qaysi fanga tegishli?\n"
        "Masalan: <i>Matematika, Ingliz tili, Tarix, Ona tili...</i>",
        parse_mode=ParseMode.HTML,
        reply_markup=nav_kb(include_back=True)
    )
    await state.set_state(AdminAddTestState.subject)


async def generate_unique_code():
    while True:
        code = "T-" + "".join(random.choices(string.ascii_uppercase + string.digits, k=5))
        if not await test_code_exists(code):
            return code


@router.message(AdminAddTestState.subject)
async def set_subject(message: Message, state: FSMContext):
    if (message.text or "") in TOP_MENU_BTNS:
        return
    subject = sanitize_plain_text(message.text, max_len=80)
    if len(subject) < 2:
        return await message.answer("⚠️ Fan nomi juda qisqa yoki noto'g'ri belgilar bor. Qayta kiriting:")
    await state.update_data(subject=subject)

    await message.answer(
        "2-qadam: test nomini kiriting.\nMasalan: <i>DTM Blok 2024</i>",
        parse_mode=ParseMode.HTML,
        reply_markup=nav_kb(include_back=True)
    )
    await state.set_state(AdminAddTestState.title)


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

    is_admin = await is_admin_user(message.from_user.id)
    if is_admin:
        kb = InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text="🆓 Tekin test", callback_data="setfree_1")],
            [InlineKeyboardButton(text="🧾 Pullik test", callback_data="setfree_0")]
        ])
        await message.answer(
            f"✅ Test kodi avtomatik yaratildi: <code>{new_code}</code>\n\n"
            f"3-qadam: test turi qanday bo'lsin?",
            parse_mode=ParseMode.HTML, reply_markup=kb)
        await state.set_state(AdminAddTestState.is_free)
    else:
        await state.update_data(is_free=1)
        await message.answer(
            f"✅ Test kodi avtomatik yaratildi: <code>{new_code}</code>\n"
            f"<i>(Oddiy foydalanuvchi sifatida faqat tekin test yaratishingiz mumkin.)</i>",
            parse_mode=ParseMode.HTML, reply_markup=nav_kb(include_back=True))
        await _ask_duration_and_qcount(message, state)


@router.callback_query(F.data.startswith("setfree_"))
async def set_is_free(callback: CallbackQuery, state: FSMContext):
    is_free = int(callback.data.split("_")[1])
    await state.update_data(is_free=is_free)

    if is_free == 1:
        await callback.message.answer("✅ Tekin test tanlandi.")
        await _ask_duration_and_qcount(callback.message, state)
    else:
        await callback.message.answer("✅ Pullik test tanlandi.\n💰 Narxni so'mda kiriting (masalan: 25000):",
                                      reply_markup=nav_kb(include_back=True))
        await state.set_state(AdminAddTestState.price)


@router.message(AdminAddTestState.price)
async def set_price(message: Message, state: FSMContext):
    if (message.text or "") in TOP_MENU_BTNS:
        return
    if not (message.text or "").isdigit():
        return await message.answer("⚠️ Narxni faqat raqamda kiriting (masalan: 25000):")
    await state.update_data(price=int(message.text))
    await _ask_duration_and_qcount(message, state)


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
        "4-qadam: <b>test davomiyligi va savollar soni</b>\n\n"
        "Tayyor variantlardan birini tanlang,\n"
        "yoki o'zingiz <b>«daqiqa savol_soni»</b> shaklida yozing.\n"
        "Masalan: <code>40 25</code> — 40 daqiqa, 25 ta savol.",
        parse_mode=ParseMode.HTML,
        reply_markup=duration_qcount_kb()
    )
    await state.set_state(AdminAddTestState.duration)


@router.callback_query(AdminAddTestState.duration, F.data.startswith("qdq_"))
async def pick_duration_and_qcount(callback: CallbackQuery, state: FSMContext):
    _, d, q = callback.data.split("_")
    await state.update_data(duration=int(d), questions_count=int(q))
    await callback.answer()
    await _ask_exam_type(callback.message, state)


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

    await state.update_data(duration=duration_val, questions_count=qcount_val)
    await _ask_exam_type(message, state)


async def _ask_exam_type(answerable, state: FSMContext):
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="📝 Oddiy test", callback_data="etype_simple")],
        [InlineKeyboardButton(text="⚡ Rasch (MS)", callback_data="etype_rasch")],
        [InlineKeyboardButton(text="⭐ Maxsus", callback_data="etype_maxsus")],
        [InlineKeyboardButton(text="📚 DTM (5 fan)", callback_data="etype_dtm")],
    ])
    await answerable.answer(
        "5-qadam: <b>baholash rejimini tanlang</b>\n\n"
        "📝 <b>Oddiy</b> — har savol 1 ball, natija oddiy foizda.\n"
        "⚡ <b>Rasch (MS)</b> — natija guruhga nisbatan statistik baholanadi "
        "(milliy sertifikat uslubiga yaqin).\n"
        "⭐ <b>Maxsus</b> — Rasch'ga o'xshash, kichik guruhlar uchun ham mos.\n"
        "📚 <b>DTM (5 fan)</b> — 5 ta fan turlicha og'irlikda baholanadi "
        "(savollar soni avtomatik 90 taga sozlanadi).",
        parse_mode=ParseMode.HTML,
        reply_markup=kb
    )
    await state.set_state(AdminAddTestState.exam_type)


@router.callback_query(F.data.startswith("etype_"))
async def set_exam_type(callback: CallbackQuery, state: FSMContext):
    et = callback.data.split("_", 1)[1]  # simple/rasch/dtm/maxsus

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
    elif et == "maxsus":
        await state.update_data(exam_type="maxsus", dtm_cfg=None)
        await callback.message.answer(
            "✅ <b>Maxsus</b> rejim tanlandi. Natija test yakunlangach avtomatik hisoblab chiqiladi.",
            parse_mode=ParseMode.HTML
        )
    else:
        await state.update_data(exam_type="simple", dtm_cfg=None)
        await callback.message.answer("✅ <b>Oddiy</b> rejim tanlandi.", parse_mode=ParseMode.HTML)

    # ✅ MUHIM: Rasch/Maxsus testlarda barcha qatnashuvchilarning natijasi BIR VAQTDA
    # (guruh bo'yicha) hisoblanadi, shuning uchun ular albatta "rejalashtirilgan"
    # (hammaga bir xil boshlanish/tugash vaqti) rejimda bo'lishi SHART.
    # Aks holda kech boshlagan foydalanuvchilarning natijasi hech qachon
    # Rasch balga aylantirilmay qolib ketardi (chunki hisoblash faqat 1 marta,
    # birinchi topshirgan kishining tugash vaqtiga qarab ishga tushadi).
    if et in ("rasch", "maxsus"):
        await state.update_data(start_mode="scheduled")
        await callback.message.answer(
            "ℹ️ Rasch/Maxsus testlarda barcha o'quvchilarning natijasi bir vaqtda (guruh bo'yicha) "
            "hisoblanadi, shuning uchun bu turdagi testlar faqat <b>rejalashtirilgan</b> "
            "(hammaga bir xil boshlanish vaqti) rejimda ishlaydi — bu qadam avtomatik tanlandi.",
            parse_mode=ParseMode.HTML
        )
        await _ask_schedule_datetime(callback.message, state)
        return

    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="🚀 Oddiy (o'quvchi istalgan vaqtda boshlaydi)", callback_data="sch_normal")],
        [InlineKeyboardButton(text="📅 Rejalashtirilgan (hammaga bir xil vaqt)", callback_data="sch_scheduled")]
    ])
    await callback.message.answer("6-qadam: <b>test qachon boshlanadi?</b>", parse_mode=ParseMode.HTML, reply_markup=kb)
    await state.set_state(AdminAddTestState.schedule_mode)


@router.callback_query(F.data == "sch_normal")
async def schedule_normal(callback: CallbackQuery, state: FSMContext):
    await state.update_data(start_mode="normal", start_ts=0, start_at=None)
    await callback.message.answer(
        "7-qadam: test faylini (PDF) yuboring:",
        parse_mode=ParseMode.HTML,
        reply_markup=nav_kb(include_back=True)
    )
    await state.set_state(AdminAddTestState.file)


@router.callback_query(F.data == "sch_scheduled")
async def schedule_scheduled(callback: CallbackQuery, state: FSMContext):
    await state.update_data(start_mode="scheduled")
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
    await state.update_data(start_ts=start_ts_val, start_at=start_at)

    await message.answer(
        f"✅ Rejalashtirildi.\n📅 Boshlanish: <b>{start_at}</b>\n\n"
        f"7-qadam: endi test faylini (PDF) yuboring:",
        parse_mode=ParseMode.HTML,
        reply_markup=nav_kb(include_back=True)
    )
    await state.set_state(AdminAddTestState.file)


@router.message(AdminAddTestState.file, F.document)
async def set_file(message: Message, state: FSMContext):
    await state.update_data(file_id=message.document.file_id)
    data = await state.get_data()
    qcount = int(data.get("questions_count") or 30)

    await message.answer(
        f"✅ Fayl qabul qilindi.\n\n"
        f"8-qadam (oxirgi): <b>javoblar kalitini kiriting.</b>\n"
        f"Aynan <b>{qcount} ta</b> harf, bo'shliqsiz, savollar tartibida.\n"
        f"Faqat <b>A, B, C, D, E, F</b> harflaridan foydalaning (masalan: <code>abcdabcd...</code>)",
        parse_mode=ParseMode.HTML,
        reply_markup=nav_kb(include_back=True)
    )
    await state.set_state(AdminAddTestState.answers)


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
            exam_type, dtm_cfg, creator_id, created_ts
        )
        VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13, $14, $15, $16)
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
        now_ts()
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


# ===================== SERTIFIKAT SOZLAMALARI (test yaratuvchisi/admin) =====================
async def _cert_settings_entry_check(user_tg_id: int, test_id: int) -> bool:
    t = await POOL.fetchrow("SELECT creator_id FROM tests WHERE id=$1", test_id)
    if not t:
        return False
    return t["creator_id"] == user_tg_id or await is_admin_user(user_tg_id)


@router.callback_query(F.data.startswith("certcfg_start_"))
async def certcfg_start(callback: CallbackQuery, state: FSMContext):
    test_id = int(callback.data.split("_")[2])
    if not await _cert_settings_entry_check(callback.from_user.id, test_id):
        return await callback.answer("Bu testga huquqingiz yo'q!", show_alert=True)

    await state.update_data(certcfg_test_id=test_id)
    await callback.message.answer(
        f"🎓 <b>1/3 — O'quv markazi nomi</b>\n\n"
        f"Sertifikatda ko'rinadigan nomni kiriting.\n"
        f"O'zgartirmaslik uchun <code>-</code> yuboring (standart: <b>{ACADEMY_NAME}</b>):",
        parse_mode=ParseMode.HTML,
        reply_markup=nav_kb(include_back=True)
    )
    await state.set_state(CertConfigState.waiting_academy)


@router.message(CertConfigState.waiting_academy)
async def certcfg_got_academy(message: Message, state: FSMContext):
    if (message.text or "") in TOP_MENU_BTNS:
        return
    val = sanitize_plain_text(message.text, max_len=100)
    if val and val != "-":
        await state.update_data(certcfg_academy=val)

    await message.answer(
        "🎓 <b>2/3 — Mas'ul shaxs</b>\n\n"
        "Sertifikatda «Mas'ul shaxs» sifatida yoziladigan ism-familiyani kiriting.\n"
        "O'zgartirmaslik uchun <code>-</code> yuboring (standart: sizning ismingiz):",
        parse_mode=ParseMode.HTML,
        reply_markup=nav_kb(include_back=True)
    )
    await state.set_state(CertConfigState.waiting_signee)


@router.message(CertConfigState.waiting_signee)
async def certcfg_got_signee(message: Message, state: FSMContext, bot: Bot):
    if (message.text or "") in TOP_MENU_BTNS:
        return
    val = sanitize_plain_text(message.text, max_len=100)
    if val and val != "-":
        await state.update_data(certcfg_signee=val)

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
    if exam_type in ('rasch', 'maxsus'):
        kb.inline_keyboard.append([InlineKeyboardButton(text="⚡ Rasch hisobot", callback_data=f"raschrep_{test_id}")])
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

    who = f"{name} (<code>{target_id}</code>)" if name else f"<code>{target_id}</code>"
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

    who = f"{name} (<code>{target_id}</code>)" if name else f"<code>{target_id}</code>"
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
    who = f"{resolved_name} (<code>{target_id}</code>)" if resolved_name else f"<code>{target_id}</code>"

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
        "• Javobni tugmalar (A–F) orqali belgilang, yoki hammasini birdaniga matn "
        "shaklida yuboring (masalan: <code>abcdabcd...</code>). Javoblar soni savollar soniga teng bo'lishi shart.\n"
        "• Aloqa uzilib qolsa yoki «⬅️ Ortga» bosib chiqib ketsangiz ham xavotir olmang — "
        "shu testni qayta ochsangiz, vaqt hisoblagichi boshidan boshlanmay, qolgan vaqtdan davom etadi.\n"
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
        "4️⃣ Admin tasdiqlagach, test <b>«📦 Mening testlarim»</b> bo'limida ochiladi.\n\n"

        "<b>5. O'zingiz test yaratish</b> (barcha ro'yxatdan o'tgan foydalanuvchilar uchun)\n"
        "• <b>➕ Test yaratish</b> — o'zingiz test tuzib ulashishingiz mumkin "
        "(kuniga 10 tagacha, faqat tekin turda).\n"
        "• <b>🛠 Yaratgan testlarim</b> — har bir testingiz uchun bitta joyda: statistikasi, "
        "<b>✏️ tahrirlash</b>, <b>🛑 vaqtincha yakunlash/qayta ochish</b>, <b>🎓 sertifikat ko'rinishini "
        "sozlash</b> (o'quv markazi nomi, mas'ul shaxs, rangi) va kerak bo'lsa <b>🗑 o'chirish</b>.\n\n"

        "<b>6. Reyting va profil</b>\n"
        "• <b>📄 Natijalarim</b> — shaxsiy statistikangiz va topshirgan testlaringiz tarixi.\n"
        "• <b>🏆 Reyting</b> — umumiy va test bo'yicha eng yaxshi natijalar.\n"
        "• <b>👤 Profilim</b> — ism-familiya va telefon raqamini o'zgartirish.\n\n"

        "🔁 <b>Navigatsiya:</b> istalgan vaqtda «🏠 Bosh menyu» yoki «⬅️ Ortga» tugmasidan foydalaning.\n\n"

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
            f"👤 {full_name}\n"
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
    while True:
        try:
            rows = await POOL.fetch("""
                SELECT r.user_id, r.test_id, r.percent, r.grade, r.date,
                       u.telegram_id, u.full_name,
                       t.code, t.title, t.subject, t.exam_type,
                       t.cert_academy_name, t.cert_signee_name, t.cert_template, t.creator_id,
                       r.score, r.total_questions
                FROM results r
                JOIN users u ON u.id=r.user_id
                JOIN tests t ON t.id=r.test_id
                WHERE (r.certificate_sent IS NULL OR r.certificate_sent=0)
                ORDER BY r.id ASC
                LIMIT 20
            """)

            for row in rows:
                (user_id_db, test_id, pct, grade, date_str,
                 user_tg_id, full_name,
                 tcode, ttitle, tsubject, exam_type, score, total_q) = (
                    row["user_id"], row["test_id"], row["percent"], row["grade"], row["date"],
                    row["telegram_id"], row["full_name"],
                    row["code"], row["title"], row["subject"], row["exam_type"], row["score"], row["total_questions"]
                )
                cert_academy, cert_signee, cert_tpl = await _resolve_cert_settings(
                    row["cert_academy_name"], row["cert_signee_name"], row["cert_template"], row["creator_id"])

                et = (exam_type or "simple")

                if et == "rasch":
                    await finalize_rush_for_test_if_ready(test_id)
                    rr = await POOL.fetchrow("""
                        SELECT rasch_percent, grade
                        FROM results
                        WHERE user_id=$1 AND test_id=$2
                        ORDER BY id ASC LIMIT 1
                    """, user_id_db, test_id)
                    if rr:
                        pct = rr["rasch_percent"]
                        grade = rr["grade"]

                elif et == "maxsus":
                    await finalize_maxsus_for_test_if_ready(test_id)
                    rr = await POOL.fetchrow("""
                        SELECT rasch_percent, grade
                        FROM results
                        WHERE user_id=$1 AND test_id=$2
                        ORDER BY id ASC LIMIT 1
                    """, user_id_db, test_id)
                    if rr:
                        pct = rr["rasch_percent"]
                        grade = rr["grade"]

                if not is_certificate_eligible(et, float(pct or 0.0), grade):
                    continue

                tmp_dir = tempfile.gettempdir()
                safe_code = re.sub(r"[^A-Z0-9\-_]+", "_", str(tcode).upper())
                pdf_path = os.path.join(tmp_dir, f"cert_{safe_code}_{user_id_db}.pdf")

                make_certificate_pdf(
                    full_name=full_name or "Noma'lum",
                    subject=(tsubject or "Noma'lum"),
                    test_code=tcode,
                    test_title=ttitle,
                    exam_date_str=(date_str or now_str_local()),
                    percent=float(pct or 0.0),
                    correct_count=int(score or 0),
                    total_questions=int(total_q or 0),
                    grade=grade,
                    out_path=pdf_path,
                    academy_name=cert_academy,
                    signee_name=cert_signee,
                    template=cert_tpl,
                )

                sent = await bot.send_document(
                    chat_id=int(user_tg_id),
                    document=FSInputFile(pdf_path),
                    caption="🎁 Sertifikatingiz tayyor! (PDF)"
                )
                cert_file_id = sent.document.file_id if sent and sent.document else None

                await POOL.execute("""
                    UPDATE results
                    SET certificate_sent=1, certificate_file_id=$1, certificate_sent_ts=$2
                    WHERE user_id=$3 AND test_id=$4
                """, cert_file_id, now_ts(), user_id_db, test_id)

        except Exception:
            logging.exception("certificate_daemon iteration failed")

        await asyncio.sleep(CERT_CHECK_EVERY_SEC)


# ===================== MA'LUMOTLAR BAZASI ZAXIRA NUSXASI (BACKUP) =====================
BACKUP_TABLES = ["users", "tests", "payments", "results", "sessions", "required_channels", "user_limits"]


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
_TABLES_WITH_SERIAL_ID = ["users", "tests", "payments", "results", "sessions", "required_channels"]


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
                cols = list(rows[0].keys())
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
        if not tables or not all(t in tables for t in BACKUP_TABLES):
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


@router.callback_query(F.data == "restore_cancel")
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


@router.callback_query(F.data == "restore_confirm")
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


@router.callback_query(F.data == "closetest_menu")
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
        "Yakunlangan testga yangi urinish qabul qilinmaydi (Rasch/Maxsus bo'lsa, natijalar darhol "
        "hisoblab chiqiladi). Qayta ochsangiz, oldin yechganlarga qayta yechish taklif qilinadi.\n\n"
        "Bosilganda holati almashadi:",
        parse_mode=ParseMode.HTML, reply_markup=kb
    )


async def closetest_menu(callback: CallbackQuery):
    if not await is_admin_user(callback.from_user.id):
        return
    await _render_closetest_list(callback.message, 0)


@router.callback_query(F.data.startswith("closetest_toggle_"))
async def closetest_toggle(callback: CallbackQuery):
    test_id = int(callback.data.split("_")[2])
    if not await _is_admin_or_creator(callback.from_user.id, test_id):
        return await callback.answer("Bu testga huquqingiz yo'q!", show_alert=True)

    row = await POOL.fetchrow("SELECT is_closed, code, exam_type FROM tests WHERE id=$1", test_id)
    if not row:
        return await callback.answer("Topilmadi", show_alert=True)

    new_val = 0 if int(row["is_closed"] or 0) == 1 else 1
    await POOL.execute("UPDATE tests SET is_closed=$1 WHERE id=$2", new_val, test_id)

    if new_val == 1:
        et = row["exam_type"] or "simple"
        if et == "rasch":
            await finalize_rush_for_test_if_ready(test_id)
        elif et == "maxsus":
            await finalize_maxsus_for_test_if_ready(test_id)
        await callback.answer(f"🛑 {row['code']} yakunlandi.", show_alert=True)
    else:
        await callback.answer(f"🟢 {row['code']} qayta ochildi.", show_alert=True)
        kb = InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text="✅ Ha, taklif qilinsin", callback_data=f"retakeoffer_{test_id}")],
            [InlineKeyboardButton(text="❌ Yo'q, kerak emas", callback_data="dismiss_msg")],
        ])
        await callback.message.answer(
            "Bu testni oldin yechgan foydalanuvchilarga qayta yechish imkoniyati taklif qilinsinmi?",
            reply_markup=kb
        )


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

    t = await POOL.fetchrow("SELECT is_free, is_hidden, is_closed FROM tests WHERE id=$1", test_id)
    if not t:
        return await callback.message.answer("Test topilmadi.")
    if int(t["is_hidden"] or 0) == 1 or int(t["is_closed"] or 0) == 1:
        return await callback.message.answer("⏸ Bu test hozircha yopiq, birozdan so'ng urinib ko'ring.")

    await POOL.execute("DELETE FROM results WHERE user_id=$1 AND test_id=$2", user_db_id, test_id)
    await POOL.execute("DELETE FROM sessions WHERE user_id=$1 AND test_id=$2", user_db_id, test_id)

    if int(t["is_free"]) == 1:
        kb = InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text="🚀 Boshlash", callback_data=f"fstart_{test_id}")]
        ])
    else:
        payment = await POOL.fetchrow(
            "SELECT id FROM payments WHERE user_id=$1 AND test_id=$2 AND status='approved' ORDER BY id DESC LIMIT 1",
            user_db_id, test_id)
        if not payment:
            return await callback.message.answer(
                "✅ Eski natijangiz o'chirildi, lekin to'lovingiz topilmadi. "
                "Qaytadan boshlash uchun testni qaytadan sotib oling.")
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

    row = await POOL.fetchrow("SELECT full_name, phone FROM users WHERE telegram_id=$1", message.from_user.id)

    if not row: return
    fname, phone = row["full_name"], row["phone"]

    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="✏️ Ism-familiyani o'zgartirish", callback_data="edit_prof_name")],
        [InlineKeyboardButton(text="📱 Telefon raqamni o'zgartirish", callback_data="edit_prof_phone")]
    ])

    await message.answer(
        f"👤 <b>Sizning profilingiz:</b>\n\n"
        f"Ism: <b>{fname}</b>\n"
        f"Telefon: <b>{phone}</b>\n\n"
        f"O'zgartirish uchun pastdagi tugmalardan birini tanlang:",
        parse_mode=ParseMode.HTML,
        reply_markup=kb
    )


@router.callback_query(F.data.startswith("edit_prof_"))
async def edit_profile_cb(callback: CallbackQuery, state: FSMContext):
    action = callback.data.split("_")[2]
    if action == "name":
        await callback.message.answer("✏️ Yangi ism-familiyangizni kiriting:", reply_markup=nav_kb(include_back=True))
        await state.set_state(EditProfileState.waiting_name)
    else:
        await callback.message.answer("📱 Yangi telefon raqamingizni yuboring:", reply_markup=phone_kb())
        await state.set_state(EditProfileState.waiting_phone)
    await callback.answer()


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
    contact = message.contact.phone_number if message.contact else sanitize_plain_text(message.text, max_len=30)
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
_PAGINATED_LIST_RENDERERS = {
    "edit": _render_edittest_list,
    "del": _render_admdel_list,
    "hide": _render_hidetest_list,
    "close": _render_closetest_list,
    "xl": _render_xl_list,
    "trange": _render_trange_list,
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


# ===================== MAIN =====================
async def main():
    global POOL
    logging.basicConfig(level=logging.INFO, stream=sys.stdout)

    POOL = await asyncpg.create_pool(dsn=DATABASE_URL, min_size=2, max_size=10)
    async with POOL.acquire() as db:
        await init_db(db)

    bot = Bot(token=BOT_TOKEN, default=DefaultBotProperties(parse_mode=ParseMode.HTML))

    await bot.set_my_commands(DEFAULT_BOT_COMMANDS, scope=BotCommandScopeDefault())
    await _sync_admin_commands(bot)

    asyncio.create_task(certificate_daemon(bot))
    asyncio.create_task(db_backup_daemon(bot))
    dp = Dispatcher()
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
