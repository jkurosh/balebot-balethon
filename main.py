import asyncio
from balethon import Client
from balethon.conditions import private, is_joined
from balethon.objects import InlineKeyboard
import khayyam
from datetime import timedelta
import os
import re
import sqlite3
from contextlib import contextmanager
import openpyxl
import random
from datetime import datetime

bot = Client("bot-token")

ADMIN_ID = 123456789


user_states = {}

SCREEN_DIR = "path"
os.makedirs(SCREEN_DIR, exist_ok=True)

DB_PATH = "appointments.db"


@contextmanager
def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    try:
        yield conn
    finally:
        conn.close()


def init_db():
    with get_db() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS appointments (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                name TEXT NOT NULL,
                phone TEXT NOT NULL,
                date TEXT NOT NULL,
                time TEXT NOT NULL,
                screen_path TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                chat_id INTEGER UNIQUE NOT NULL,
                phone_number TEXT NOT NULL,
                name TEXT NOT NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        conn.execute('''
                CREATE TABLE IF NOT EXISTS user_answers (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    user_id INTEGER NOT NULL,
                    question_id INTEGER NOT NULL,
                    answer TEXT NOT NULL,
                    answered_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    UNIQUE(user_id, question_id)
                )
            ''')
        conn.execute('''
                CREATE TABLE IF NOT EXISTS questions (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    text TEXT NOT NULL,
                    option_a TEXT NOT NULL,
                    option_b TEXT NOT NULL,
                    option_c TEXT NOT NULL,
                    option_d TEXT NOT NULL,
                    correct_answer TEXT NOT NULL,
                    explanation TEXT
                )
            ''')

        # جدول وضعیت سوال روز
        conn.execute('''
                CREATE TABLE IF NOT EXISTS question_of_day_state (
                    id INTEGER PRIMARY KEY DEFAULT 1,
                    current_question_id INTEGER DEFAULT 0,
                    last_update_date TEXT
                )
            ''')

        conn.commit()
        cursor = conn.execute("SELECT COUNT(*) FROM question_of_day_state")
        if cursor.fetchone()[0] == 0:
            today = datetime.now().date().isoformat()
            conn.execute(
                "INSERT INTO question_of_day_state (id, current_question_id, last_update_date) VALUES (1, 0, ?)",
                (today,)
            )

        conn.commit()

def cleanup_old_slots():
    today = khayyam.JalaliDatetime.now()
    today_str = today.strftime("%Y/%m/%d")

    with get_db() as conn:
        conn.execute("DELETE FROM appointments WHERE date < ?", (today_str,))
        conn.commit()

def is_user_registered(chat_id: int) -> bool:
    """بررسی اینکه کاربر قبلا ثبت‌نام کرده یا نه"""
    with get_db() as conn:
        result = conn.execute(
            "SELECT 1 FROM users WHERE chat_id = ?",
            (chat_id,)
        ).fetchone()
        return result is not None


def save_user(chat_id: int, phone_number: str, name: str):
    """ذخیره اطلاعات کاربر جدید"""
    with get_db() as conn:
        conn.execute(
            "INSERT INTO users (chat_id, phone_number, name) VALUES (?, ?, ?)",
            (chat_id, phone_number, name)
        )
        conn.commit()


def has_user_answered(user_id: int, question_id: int) -> bool:
    """چک می‌کند آیا کاربر قبلاً به این سوال پاسخ داده"""
    with get_db() as conn:
        result = conn.execute(
            "SELECT id FROM user_answers WHERE user_id = ? AND question_id = ?",
            (user_id, question_id)
        ).fetchone()
        return result is not None

def is_slot_booked(date: str, time: str) -> bool:
    with get_db() as conn:
        result = conn.execute(
            "SELECT 1 FROM appointments WHERE date = ? AND time = ?",
            (date, time)
        ).fetchone()
        return result is not None


def save_appointment(user_id: int, name: str, phone: str, date: str, time: str, screen_path: str):
    with get_db() as conn:
        conn.execute(
            "INSERT INTO appointments (user_id, name, phone, date, time, screen_path) VALUES (?, ?, ?, ?, ?, ?)",
            (user_id, name, phone, date, time, screen_path)
        )
        conn.commit()

EXAM_CATEGORIES = {
    "navle": "NAVLE",
    "bcse": "BCSE"
}

CATEGORY_MAP = {
    "dam_kochik": "دام کوچک",
    "dam_bozorg": "دام بزرگ",
    "surg": "جراحی",
    "radio": "رادیولوژی",
    "clinical": "کلینیکال پاتولوژی",
    "pharma": "فارماکولوژی",
    "birds": "پرندگان",
    "others": "دیگر",
    "short_course":"دوره کوتاه"
}

KONKOR_CATEGORIES = {
    "surgery": "جراحی",
    "mamaii": "مامایی",
    "dam_bozorg_internal": "داخلی دام بزرگ",
    "dam_kochik_internal": "داخلی دام کوچک",
    "radiology": "رادیولوژی",
    "clinical_path": "کلینیکال پاتولوژی"
}


def parse_konkor(file_path: str) -> dict:
    """پارس کردن فایل کنکور و دسته‌بندی بر اساس رشته"""
    konkor_data = {
        "surgery": [],
        "mamaii": [],
        "dam_bozorg_internal": [],
        "dam_kochik_internal": [],
        "radiology": [],
        "clinical_path": []
    }

    with open(file_path, "r", encoding="utf-8") as f:
        content = f.read()

    entries = content.split("ـــــــــــــــ")

    for entry in entries:
        entry = entry.strip()
        if not entry:
            continue

        lines = entry.split("\n")
        if len(lines) < 2:
            continue

        title = lines[0].strip()
        link = ""

        for line in lines[1:]:
            if line.strip().startswith("دانلود:"):
                link = line.replace("دانلود:", "").strip()
                break

        if not link:
            continue

        if "جراحی" in title:
            konkor_data["surgery"].append({"title": title, "link": link})
        elif "مامائی" in title or "مامایی" in title:
            konkor_data["mamaii"].append({"title": title, "link": link})
        elif "داخلی دام بزرگ" in title:
            konkor_data["dam_bozorg_internal"].append({"title": title, "link": link})
        elif "داخلی دام کوچک" in title:
            konkor_data["dam_kochik_internal"].append({"title": title, "link": link})
        elif "رادیولوژی" in title:
            konkor_data["radiology"].append({"title": title, "link": link})
        elif "کلینیکال پاتولوژی" in title:
            konkor_data["clinical_path"].append({"title": title, "link": link})

    return konkor_data


ALL_KONKOR = parse_konkor("konkour.txt")


def get_current_question_id():
    """دریافت شماره سوال فعلی از دیتابیس"""
    with get_db() as conn:
        row = conn.execute(
            "SELECT current_question_id, last_update_date FROM question_of_day_state WHERE id = 1"
        ).fetchone()

        if not row:
            return 0

        current_id = row[0]
        last_update = datetime.fromisoformat(row[1]).date()
        today = datetime.now().date()

        # اگه روز عوض شده، سوال رو آپدیت کن
        if today > last_update:
            questions = load_questions()
            if questions:
                new_id = (current_id + 1) % len(questions)
                conn.execute(
                    "UPDATE question_of_day_state SET current_question_id = ?, last_update_date = ? WHERE id = 1",
                    (new_id, today.isoformat())
                )
                conn.commit()
                return new_id

        return current_id


def get_current_question():
    """دریافت سوال روز فعلی"""
    questions = load_questions()
    if not questions:
        return None

    question_id = get_current_question_id()
    return questions[question_id]


def has_user_answered_today(user_id: int) -> bool:
    """چک می‌کند آیا کاربر امروز به سوال پاسخ داده"""
    question_id = get_current_question_id()
    return has_user_answered(user_id, question_id)

def get_konkor_by_category(category_key: str) -> list:
    """دریافت لیست نمونه سوالات بر اساس رشته"""
    return ALL_KONKOR.get(category_key, [])


def format_konkor(k: dict) -> str:
    download_link = k.get('link', '-')
    # تبدیل لینک به فرمت کلیک‌شو
    if download_link and download_link != '-':
        download_text = f"[دانلود]({download_link})"
    else:
        download_text = "لینک موجود نیست"

    return (
        f"📝 {k.get('title', '-')}\n"
        f"🔗 {download_text}"
    )


def build_konkor_page(items: list, page: int, category_key: str):
    """ساخت صفحه نمونه سوالات با صفحه‌بندی"""
    per_page = 10
    start = page * per_page
    end = start + per_page
    page_items = items[start:end]
    total_pages = (len(items) + per_page - 1) // per_page

    text = f"📖 {KONKOR_CATEGORIES[category_key]} | صفحه {page + 1} از {total_pages}\n"
    text += "━" * 25 + "\n\n"
    text += ("\n\n" + "━" * 25 + "\n\n").join(format_konkor(k) for k in page_items)

    nav = []
    if page > 0:
        nav.append((f"◀️ قبلی", f"kpage_{category_key}_{page - 1}"))
    if end < len(items):
        nav.append((f"بعدی ▶️", f"kpage_{category_key}_{page + 1}"))

    rows = []
    if nav:
        rows.append(nav)
    rows.append([("بازگشت به کنکور", "konkor"), ("منوی اصلی", "main_menu")])

    return text, InlineKeyboard(*rows)


def load_questions():
    """بارگذاری سوالات از فایل questions.xlsx"""
    try:
        wb = openpyxl.load_workbook("questions.xlsx")
        sheet = wb.active
        questions = []

        for row in sheet.iter_rows(min_row=2, values_only=True):  # از ردیف 2 شروع (ردیف 1 هدر)
            if row[0]:  # اگر سوال وجود داشت
                questions.append({
                    "question": row[0],
                    "option1": row[1],
                    "option2": row[2],
                    "option3": row[3],
                    "option4": row[4],
                    "correct": row[5],  # مثلا "option 1"
                    "description": row[6]
                })

        return questions
    except Exception as e:
        print(f"خطا در خواندن فایل سوالات: {e}")
        return []


# ذخیره وضعیت سوال روز

def shuffle_options(question):
    """گزینه‌ها رو رندوم می‌کنه و نقشه صحیح رو برمی‌گردونه"""
    options = [
        ("option1", question["option1"]),
        ("option2", question["option2"]),
        ("option3", question["option3"]),
        ("option4", question["option4"])
    ]

    random.shuffle(options)

    # پیدا کردن گزینه صحیح بعد از shuffle
    correct_key = question["correct"].replace(" ", "")  # "option 1" -> "option1"

    shuffled_map = {}
    for idx, (original_key, text) in enumerate(options, 1):
        shuffled_map[f"opt{idx}"] = {
            "text": text,
            "is_correct": original_key == correct_key
        }

    return options, shuffled_map

def parse_handouts(file_path: str) -> list:
    handouts = []
    current = {}

    with open(file_path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("---"):
                if current:
                    handouts.append(current)
                    current = {}
                continue
            for key in ["درس", "کتگوری", "دانشگاه", "اسم استاد", "توضیحات", "لینک دانلود جزوه"]:
                if line.startswith(f"{key}:"):
                    value = line.replace(f"{key}:", "").strip()
                    current[key] = value
                    break
            else:
                if current and "لینک دانلود جزوه" in current and not current["لینک دانلود جزوه"]:
                    current["لینک دانلود جزوه"] = line

    if current:
        handouts.append(current)

    return handouts

def parse_videos(file_path: str) -> list:
    videos = []
    current = {}
    waiting_for_download_link = False

    with open(file_path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()

            # جداکننده هر آیتم
            if line.startswith("----------------"):
                if current:
                    videos.append(current)
                    current = {}
                waiting_for_download_link = False
                continue

            if not line:
                continue

            # عنوان
            if line.startswith("عنوان:"):
                current["title"] = line.replace("عنوان:", "").strip()

            # دسته بندی
            elif line.startswith("دسته بندی:"):
                current["category"] = line.replace("دسته بندی:", "").strip()

            # خط دانلود
            elif line.startswith("دانلود:"):
                waiting_for_download_link = True

            # لینک دانلود در خط بعد
            elif waiting_for_download_link:
                current["download"] = line
                waiting_for_download_link = False

    # آخرین آیتم
    if current:
        videos.append(current)

    return videos

def parse_exams(file_path: str) -> list:
    """پارس کردن فایل آزمون‌ها (NAVLE/BCSE)"""
    exams = []
    current = {}

    with open(file_path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()

            if line.startswith("----------------"):
                if current:
                    exams.append(current)
                    current = {}
                continue

            if not line:
                continue

            if line.startswith("عنوان:"):
                current["title"] = line.replace("عنوان:", "").strip()
            elif line.startswith("دسته بندی:"):
                current["category"] = line.replace("دسته بندی:", "").strip()
            elif line.startswith("دانلود:"):
                pass  # خط بعدی لینک هست
            elif "http" in line or "my.files.ir" in line or "my.uupload.ir" in line:
                current["download"] = line

    if current:
        exams.append(current)

    return exams

ALL_EXAMS = parse_exams("navle_bcse_files copy.txt")


ALL_HANDOUTS = parse_handouts("jozve.txt")

ALL_VIDEOS = parse_videos("video_bot.txt")

def get_handouts_by_category(category_key: str) -> list:
    category_name = CATEGORY_MAP.get(category_key, "")
    return [h for h in ALL_HANDOUTS if h.get("کتگوری") == category_name]

def get_videos_by_category(category_key: str) -> list:
    category_name = CATEGORY_MAP.get(category_key, "")
    return [v for v in ALL_VIDEOS if v.get("category") == category_name]

def get_exams_by_category(category_key: str) -> list:
    """دریافت آزمون‌ها بر اساس دسته (navle/bcse)"""
    return [e for e in ALL_EXAMS if e.get("category") == category_key]


def format_handout(h: dict) -> str:
    download_link = h.get('لینک دانلود جزوه', '-')
    # تبدیل لینک به فرمت کلیک‌شو
    if download_link and download_link != '-':
        download_text = f"[دانلود]({download_link})"
    else:
        download_text = "لینک موجود نیست"

    return (
        f"📚 عنوان درس: {h.get('درس', '-')}\n"
        f"🏫 دانشگاه: {h.get('دانشگاه', '-')}\n"
        f"👨‍🏫 نام استاد: {h.get('اسم استاد', '-')}\n"
        f"📝 توضیحات: {h.get('توضیحات', '-')}\n"
        f"🔗 {download_text}"
    )

def format_video(v: dict) -> str:
    download_link = v.get("download", "")

    if download_link:
        download_text = f"[دانلود]({download_link})"
    else:
        download_text = "لینک موجود نیست"

    return (
        f"🎥 عنوان: {v.get('title', '-')}\n"
        f"🔗 {download_text}"
    )

def format_exam(e: dict) -> str:
    """فرمت کردن یک آیتم آزمون"""
    download_link = e.get("download", "")

    if download_link:
        download_text = f"[دانلود]({download_link})"
    else:
        download_text = "لینک موجود نیست"

    return (
        f"📝 {e.get('title', '-')}\n"
        f"🔗 {download_text}"
    )


def build_handouts_page(items: list, page: int, category_key: str):
    per_page = 10
    start = page * per_page
    end = start + per_page
    page_items = items[start:end]
    total_pages = (len(items) + per_page - 1) // per_page

    text = f"📖 {CATEGORY_MAP[category_key]} | صفحه {page + 1} از {total_pages}\n"
    text += "━" * 25 + "\n\n"
    text += ("\n\n" + "━" * 25 + "\n\n").join(format_handout(h) for h in page_items)

    rows = []

    # ساخت ردیف ناوبری
    nav_row = []
    if page > 0:
        nav_row.append((f"◀️ قبلی", f"hpage_{category_key}_{page - 1}"))
    if end < len(items):
        nav_row.append((f"بعدی ▶️", f"hpage_{category_key}_{page + 1}"))

    if nav_row:
        rows.append(nav_row)

    rows.append([("بازگشت به جزوات", "handouts"), ("منوی اصلی", "main_menu")])

    return text, InlineKeyboard(*rows)

def build_videos_page(items: list, page: int, category_key: str):
    per_page = 10

    start = page * per_page
    end = start + per_page

    page_items = items[start:end]

    total_pages = (len(items) + per_page - 1) // per_page

    text = f"🎥 {CATEGORY_MAP[category_key]} | صفحه {page + 1} از {total_pages}\n"
    text += "━" * 25 + "\n\n"

    text += ("\n\n" + "━" * 25 + "\n\n").join(
        format_video(v) for v in page_items
    )

    rows = []

    nav_row = []

    if page > 0:
        nav_row.append(("◀️ قبلی", f"vpage_{category_key}_{page - 1}"))

    if end < len(items):
        nav_row.append(("بعدی ▶️", f"vpage_{category_key}_{page + 1}"))

    if nav_row:
        rows.append(nav_row)

    rows.append([
        ("بازگشت به ویدیوها", "vid"),
        ("منوی اصلی", "main_menu")
    ])

    return text, InlineKeyboard(*rows)

def build_exams_page(items: list, page: int, category_key: str):
    """ساخت صفحه آزمون‌ها با صفحه‌بندی"""
    per_page = 10
    start = page * per_page
    end = start + per_page
    page_items = items[start:end]
    total_pages = (len(items) + per_page - 1) // per_page

    text = f"📖 {EXAM_CATEGORIES[category_key]} | صفحه {page + 1} از {total_pages}\n"
    text += "━" * 25 + "\n\n"
    text += ("\n\n" + "━" * 25 + "\n\n").join(format_exam(e) for e in page_items)

    rows = []

    nav_row = []
    if page > 0:
        nav_row.append(("◀️ قبلی", f"epage_{category_key}_{page - 1}"))
    if end < len(items):
        nav_row.append(("بعدی ▶️", f"epage_{category_key}_{page + 1}"))

    if nav_row:
        rows.append(nav_row)

    rows.append([("بازگشت به آزمون‌ها", "soalat"), ("منوی اصلی", "main_menu")])

    return text, InlineKeyboard(*rows)

def build_konkor_page(items: list, page: int, category_key: str):
    per_page = 10
    start = page * per_page
    end = start + per_page
    page_items = items[start:end]
    total_pages = (len(items) + per_page - 1) // per_page

    text = f"📖 {KONKOR_CATEGORIES[category_key]} | صفحه {page + 1} از {total_pages}\n"
    text += "━" * 25 + "\n\n"
    text += ("\n\n" + "━" * 25 + "\n\n").join(format_konkor(k) for k in page_items)

    rows = []

    # ساخت ردیف ناوبری
    nav_row = []
    if page > 0:
        nav_row.append((f"◀️ قبلی", f"kpage_{category_key}_{page - 1}"))
    if end < len(items):
        nav_row.append((f"بعدی ▶️", f"kpage_{category_key}_{page + 1}"))

    if nav_row:
        rows.append(nav_row)

    rows.append([("بازگشت به کنکور", "konkor"), ("منوی اصلی", "main_menu")])

    return text, InlineKeyboard(*rows)


def normalize_persian_digits(text: str) -> str:
    persian_digits = '۰۱۲۳۴۵۶۷۸۹'
    english_digits = '0123456789'
    return text.translate(str.maketrans(persian_digits, english_digits))


def validate_phone(number: str) -> bool:
    number = normalize_persian_digits(number)
    return bool(re.match(r'^(09[0-9]{9}|\+98[0-9]{10}|0098[0-9]{10})$', number))


MAIN_MENU_KB = InlineKeyboard(
    [("ویدیو های آموزشی", "vid"), ("آزمون‌ها", "soalat")],
    [("فروشگاه", "shop"), ("تطبیق مدرک", "tatbigh")],
    [("جزوات ایرانی", "handouts"), ("رزرو وقت مشاوره", "appointment")],
    #[("سوال روز 📝", "question_of_day")]  # دکمه جدید
)

@bot.on_message(private)
async def answer_message(message):
    user_id = message.author.id

    if user_id in user_states:
        state = user_states[user_id]

        if state["step"] == "waiting_name":
            user_states[user_id]["name"] = message.text
            user_states[user_id]["step"] = "waiting_phone"
            await message.reply(
                "لطفاً شماره تماس خود را وارد کنید:\n(مثال: 09123456789)",
                InlineKeyboard([("بازگشت به منوی اصلی", "main_menu")])
            )
            return

        elif state["step"] == "waiting_phone":
            if validate_phone(message.text):
                user_states[user_id]["phone"] = normalize_persian_digits(message.text)
                user_states[user_id]["step"] = "waiting_payment"
                await message.reply(
                    "✅ اطلاعات شما ثبت شد.\n\n"
                    "لطفا مبلغ ۱.۵۰۰.۰۰۰ تومان را به شماره کارت زیر واریز کنید و اسکرین شات انتقال وجه را در ادامه ارسال بفرمایید.\n\n5022-2913-1061-4589\n\nم د\n\nسپس منتظر باشید که توسط ادمین این نوبت تائید شود.\n\nاین فرایند ممکن است کمی طول بکشد. لطفا صبور باشید.\n\nنتیجه ثبت نام از طریق بات به شما اعلام می‌شود.",
                    InlineKeyboard([("بازگشت به منوی اصلی", "main_menu")])
                )

            else:
                await message.reply(
                    "❌ شماره تماس نامعتبر است.\nفرمت‌های قابل قبول:\n09123456789\n+989123456789",
                    InlineKeyboard([("بازگشت به منوی اصلی", "main_menu")])
                )
            return
        # ثبت‌نام برای سوال روز
        if state["step"] == "qod_waiting_name":
            user_states[user_id]["qod_name"] = message.text
            user_states[user_id]["step"] = "qod_waiting_phone"
            await message.reply(
                "لطفاً شماره تماس خود را وارد کنید:\n(مثال: 09123456789)",
                InlineKeyboard([("بازگشت به منوی اصلی", "main_menu")])
            )
            return

        elif state["step"] == "qod_waiting_phone":
            if validate_phone(message.text):
                phone = normalize_persian_digits(message.text)
                name = user_states[user_id]["qod_name"]

                save_user(user_id, phone, name)
                del user_states[user_id]

                await message.reply(
                    "✅ ثبت‌نام شما با موفقیت انجام شد!\n\nحالا می‌توانید سوال روز را مشاهده کنید.",
                    InlineKeyboard([("مشاهده سوال روز 📝", "question_of_day")])
                )
            else:
                await message.reply(
                    "❌ شماره تماس نامعتبر است.\nفرمت‌های قابل قبول:\n09123456789\n+989123456789",
                    InlineKeyboard([("بازگشت به منوی اصلی", "main_menu")])
                )
            return
        elif state["step"] == "waiting_payment":
            if message.photo:
                downloading = await message.reply("در حال دریافت رسید...")
                photo = message.photo[-1]
                response = await bot.download(photo.id)
                safe_date = state['date'].replace("/", "-")
                safe_time = state['time'].replace("-", "_")
                file_path = os.path.join(SCREEN_DIR, f"{user_id}_{safe_date}_{safe_time}.jpg")
                with open(file_path, "wb") as file:
                    file.write(response)
                await downloading.delete()

                user_states[user_id]["screen_path"] = file_path

                try:
                    await bot.send_message(ADMIN_ID,
                                            f"🔔 مشاوره جدید:\n\n"
                                            f"👤 نام: {state['name']}\n"
                                            f"📞 شماره: {state['phone']}\n"
                                            f"📅 تاریخ: {state['date']}\n"
                                            f"🕐 ساعت: {state['time']}\n"
                                            f"🆔 User ID: {user_id}"
                                            )
                    await asyncio.sleep(1)
                    await bot.send_photo(
                        ADMIN_ID,
                        file_path,
                        caption="رسید پرداخت:",
                        reply_markup=InlineKeyboard(
                            [("✅ تایید و ثبت", f"confirm_{user_id}_{safe_date}_{safe_t ime}"),
                            ("❌ رد کردن", f"reject_{user_id}_{safe_date}_{safe_time}")]
                        )
                    )
                except Exception as e:
                    print(f"خطا در ارسال به ادمین: {e}")

                await message.reply(
                    "✅ رسید پرداخت شما دریافت شد.\nپس از تایید ادمین، زمان شما رزرو خواهد شد. 🙏",
                    InlineKeyboard([("بازگشت به منوی اصلی", "main_menu")])
                )
            else:
                await message.reply(
                    "لطفاً تصویر رسید پرداخت را ارسال کنید.",
                    InlineKeyboard([("بازگشت به منوی اصلی", "main_menu")])
                )
            return

    await message.reply("👨🏻‍⚕️👩🏻‍⚕️ سلام دکتر عزیز"
            "\n\n🔰 این بات در مجموعه‌ی آموزشی @ برای دسترسی راحت و جامع به اطلاعات دامپزشکی در زمان قطع اینترنت بین‌الملل برای دامپزشکان ساخته شده. شما می‌توانید با کلیک کردن روی هر کدام از دکمه‌های زیر به اطلاعات دسته‌بندی مربوطه دسترسی پیدا کنید. 🙏🏻"
            "\n\nهمچنین می‌توانید با معرفی این بات به دیگران از بات ما حمایت کنید 😍"
        "\n\n🔸 در صورت نیاز به پشتیبانی و یا پیشنهاد ارائه فایل‌ها جدید به آیدی @@ در پیامرسان «بله» پیام بدهید.", MAIN_MENU_KB)


@bot.on_callback_query()
async def answer_callback_query(callback_query):
    user_id = callback_query.author.id
    data = callback_query.data

    # ویدیوها - صفحه‌بندی
    if data.startswith("vpage_"):
        parts = data.split("_")

        if len(parts) >= 3:
            page_str = parts[-1]
            category_key = "_".join(parts[1:-1])

            try:
                page = int(page_str)

                items = get_videos_by_category(category_key)

                text, kb = build_videos_page(items, page, category_key)

                await callback_query.message.edit_text(text,reply_markup=kb)

            except (ValueError, KeyError) as e:
                await callback_query.answer(
                    f"خطا: {str(e)}",
                    show_alert=True
                )

        return

    # آزمون‌ها - صفحه‌بندی
    if data.startswith("epage_"):
        parts = data.split("_")
        if len(parts) >= 3:
            page_str = parts[-1]
            category_key = "_".join(parts[1:-1])

            try:
                page = int(page_str)
                items = get_exams_by_category(category_key)
                text, kb = build_exams_page(items, page, category_key)
                await callback_query.message.edit_text(text, reply_markup=kb)
            except (ValueError, KeyError) as e:
                await callback_query.answer(f"خطا: {str(e)}", show_alert=True)
        return


    # جزوات - صفحه‌بندی (باید قبل از CATEGORY_MAP باشه)
    if data.startswith("hpage_"):
        parts = data.split("_")  # بدون محدودیت split
        if len(parts) >= 3:
            # parts = ["hpage", "dam", "kochik", "1"] یا ["hpage", "radio", "1"]
            page_str = parts[-1]  # آخرین عنصر = شماره صفحه
            category_key = "_".join(parts[1:-1])  # همه چیز بین hpage و شماره صفحه

            try:
                page = int(page_str)
                items = get_handouts_by_category(category_key)
                text, kb = build_handouts_page(items, page, category_key)
                await callback_query.message.edit_text(text, reply_markup=kb)
            except (ValueError, KeyError) as e:
                await callback_query.answer(f"خطا: {str(e)}", show_alert=True)
                text, kb = build_handouts_page(items, page, category_key)
                await callback_query.message.edit_text(text, reply_markup=kb, parse_mode="markdown")
        return

    # کنکور - صفحه‌بندی
    if data.startswith("kpage_"):
        parts = data.split("_")
        if len(parts) >= 3:
            page_str = parts[-1]
            category_key = "_".join(parts[1:-1])

            try:
                page = int(page_str)
                items = get_konkor_by_category(category_key)
                text, kb = build_konkor_page(items, page, category_key)
                await callback_query.message.edit_text(text, reply_markup=kb)
            except (ValueError, KeyError) as e:
                await callback_query.answer(f"خطا: {str(e)}", show_alert=True)
                text, kb = build_konkor_page(items, page, category_key)
                await callback_query.message.edit_text(text, reply_markup=kb, parse_mode="markdown")
        return

    if data == "check_membership":
        member = await bot.get_chat_member(user_id)
        if member and member.status in ["creator", "administrator", "member"]:
            await callback_query.message.edit_text("👨🏻‍⚕️👩🏻‍⚕️ سلام دکتر عزیز"
            "\n\n🔰 این بات در مجموعه‌ی آموزشی @ برای دسترسی راحت و جامع به اطلاعات دامپزشکی در زمان قطع اینترنت بین‌الملل برای دامپزشکان ساخته شده. شما می‌توانید با کلیک کردن روی هر کدام از دکمه‌های زیر به اطلاعات دسته‌بندی مربوطه دسترسی پیدا کنید. 🙏🏻"
            "\n\nهمچنین می‌توانید با معرفی این بات به دیگران از بات ما حمایت کنید 😍"
        "\n\n🔸 در صورت نیاز به پشتیبانی و یا پیشنهاد ارائه فایل‌ها جدید به آیدی @@ در پیامرسان «بله» پیام بدهید.", MAIN_MENU_KB)
        else:
            await callback_query.answer("هنوز عضو کانال نشدید ❌", show_alert=True)
        return

    if data.startswith("confirm_"):
        if user_id != ADMIN_ID:
            return

        parts = data.split("_", 3 )
        target_user = int(parts[1])
        date = parts[2].replace("-", "/")
        time = parts[3].replace("_", "-")

        state = user_states.get(target_user)
        if not state:
            await callback_query.answer("خطا: اطلاعات کاربر یافت نشد", show_alert=True)
            return

        # ذخیره رزرو
        save_appointment(
            target_user,
            state["name"],
            state["phone"],
            date,
            time,
            state["screen_path"]
        )

        del user_states[target_user]

        # اطلاع به کاربر
        await bot.send_message(
            target_user,
            f"✅ رزرو شما برای تاریخ {date} ساعت {time} تایید شد.\n"
            "در زمان مقرر آنلاین باشید 🙏",
            reply_markup=InlineKeyboard([("بازگشت به منوی اصلی", "main_menu")])
        )

        # ویرایش caption عکس
        try:
            await callback_query.message.edit_caption(
                f"✅ تایید شده\n\n"
                f"👤 نام: {state['name']}\n"
                f"📞 شماره: {state['phone']}\n"
                f"📅 تاریخ: {date}\n"
                f"🕐 ساعت: {time}\n"
                f"🆔 User ID: {target_user}",
                reply_markup=None
            )
        except Exception as e:
            print(f"خطا در ویرایش: {e}")

        await callback_query.answer("رزرو تایید و ثبت شد ✅")
        return

    if data.startswith("reject_"):
        if user_id != ADMIN_ID:
            return

        parts = data.split("_")
        target_user = int(parts[1])
        date = parts[2].replace("-", "/")
        time = parts[3].replace("_", "-")

        state = user_states.get(target_user)

        if target_user in user_states:
            del user_states[target_user]

        await bot.send_message(
            target_user,
            "❌ متاسفانه رزرو شما تایید نشد.\nلطفاً دوباره تلاش کنید یا با پشتیبانی تماس بگیرید.",
            reply_markup=InlineKeyboard([("بازگشت به منوی اصلی", "main_menu")])
        )

        # ویرایش caption عکس
        try:
            name = state['name'] if state else "نامشخص"
            phone = state['phone'] if state else "نامشخص"

            await callback_query.message.edit_caption(
                f"❌ رد شده\n\n"
                f"👤 نام: {name}\n"
                f"📞 شماره: {phone}\n"
                f"📅 تاریخ: {date}\n"
                f"🕐 ساعت: {time}\n"
                f"🆔 User ID: {target_user}",
                reply_markup=None
            )
        except Exception as e:
            print(f"خطا در ویرایش: {e}")

        await callback_query.answer("رزرو رد شد ❌")
        return

    if data == "main_menu":
        if user_id in user_states:
            step = user_states[user_id].get("step","")
            if step not in ["waiting_payment"]:
                del user_states[user_id]
        await callback_query.message.edit_text(
            "👨🏻‍⚕️👩🏻‍⚕️ سلام دکتر عزیز"
            "\n\n🔰 این بات در مجموعه‌ی آموزشی @ برای دسترسی راحت و جامع به اطلاعات دامپزشکی در زمان قطع اینترنت بین‌الملل برای دامپزشکان ساخته شده. شما می‌توانید با کلیک کردن روی هر کدام از دکمه‌های زیر به اطلاعات دسته‌بندی مربوطه دسترسی پیدا کنید. 🙏🏻"
            "\n\nهمچنین می‌توانید با معرفی این بات به دیگران از بات ما حمایت کنید 😍"
        "\n\n🔸 در صورت نیاز به پشتیبانی و یا پیشنهاد ارائه فایل‌ها جدید به آیدی @@ در پیامرسان «بله» پیام بدهید.",
            reply_markup=MAIN_MENU_KB
        )
        return
    if data == "Zuku":
        await callback_query.message.edit_text(
            "",


        )
        if data == "vetprep":
            await callback_query.message.edit_text(
                "",



            )
    if data == "vid":
        await callback_query.message.edit_text(
            "♦️ در هر دسته بندی می‌توانید به بانک ویدئو مورد نظرتان دسترسی دشته باشید. برای غنی شدن منبع ویدئویی می‌توانید پیشنهادات یا فایل‌های خود را با آیدی تلگرامی @@ در «بله» به اشتراک بذارید.",
            reply_markup=InlineKeyboard(
                [("دام بزرگ", "video_dam_bozorg"), ("دام کوچک", "video_dam_kochik")],
                [("رادیولوژی", "video_radio"), ("بیهوشی و جراحی", "video_surg")],
                [("فارماکولوژی", "video_pharma"), ("کلینیکال پاتولوژی", "video_clinical")],
                [("بازگشت به منوی اصلی", "main_menu"),("دوره کوتاه","video_short_course")]
            )
        )
        return

    if data == "soalat":
        await callback_query.message.edit_text(
            "🔸 در هر دسته بندی می‌توانید به بانک سوالات مورد نظرتان دسترسی دشته باشید.",
            reply_markup=InlineKeyboard(
                [("NAVLE", "NAVLE"), ("BCSE", "BCSE")],
                [("کنکور ایران", "konkor"), ("بازگشت به منوی اصلی", "main_menu")],
            )
        )
        return

    # انتخاب NAVLE یا BCSE
    if data in ["NAVLE", "BCSE"]:
        category_key = data.lower()
        items = get_exams_by_category(category_key)

        if not items:
            await callback_query.message.edit_text(
                f"❌ فایلی برای {data} یافت نشد.",
                reply_markup=InlineKeyboard([("بازگشت", "soalat")])
            )
            return

        text, kb = build_exams_page(items, 0, category_key)
        await callback_query.message.edit_text(text, reply_markup=kb)
        return

    if data == "question_of_day":
        if not is_user_registered(user_id):
            user_states[user_id] = {"step": "qod_waiting_name"}
            await callback_query.message.edit_text(
                "👋 برای دسترسی به سوال روز، لطفاً ابتدا ثبت‌نام کنید.\n\n"
                "لطفاً نام و نام خانوادگی خود را وارد کنید:",
                reply_markup=InlineKeyboard([("بازگشت به منوی اصلی", "main_menu")])
            )
            return

        question = get_current_question()
        if not question:
            await callback_query.message.edit_text(
                "❌ خطا در بارگذاری سوال روز",
                reply_markup=InlineKeyboard([("بازگشت", "main_menu")])
            )
            return

        question_id = get_current_question_id()

        if has_user_answered(user_id, question_id):
            await callback_query.answer(
                "شما قبلاً به این سوال پاسخ داده‌اید!",
                show_alert=True
            )
            return

        shuffled_options, option_map = shuffle_options(question)

        if user_id not in user_states:
            user_states[user_id] = {}
        user_states[user_id]["qod_map"] = option_map
        user_states[user_id]["qod_description"] = question["description"]
        user_states[user_id]["qod_question_id"] = question_id

        keyboard_rows = []
        for idx, (_, text) in enumerate(shuffled_options, 1):
            keyboard_rows.append([(text, f"qod_answer_{idx}")])
        keyboard_rows.append([("بازگشت", "main_menu")])

        await callback_query.message.edit_text(
            f"📝 سوال روز:\n\n{question['question']}",
            reply_markup=InlineKeyboard(*keyboard_rows)
        )
        return

    if data.startswith("qod_answer_"):
        question_id = get_current_question_id()

        if has_user_answered(user_id, question_id):
            await callback_query.answer("شما قبلاً پاسخ داده‌اید!", show_alert=True)
            return

        selected = data.replace("qod_answer_", "")
        state = user_states.get(user_id, {})
        option_map = state.get("qod_map", {})
        description = state.get("qod_description", "")

        if not option_map:
            await callback_query.answer("خطا در پردازش پاسخ", show_alert=True)
            return

        selected_option = option_map.get(f"opt{selected}")
        if not selected_option:
            await callback_query.answer("خطا در پردازش پاسخ", show_alert=True)
            return

        is_correct = selected_option["is_correct"]

        with get_db() as conn:
            try:
                conn.execute(
                    "INSERT INTO user_answers (user_id, question_id, answer) VALUES (?, ?, ?)",
                    (user_id, question_id, selected)
                )
                conn.commit()
            except sqlite3.IntegrityError:
                pass

        if user_id in user_states:
            user_states[user_id].pop("qod_map", None)
            user_states[user_id].pop("qod_description", None)
            user_states[user_id].pop("qod_question_id", None)

        result_text = "✅ درست" if is_correct else "❌ غلط"

        await callback_query.message.edit_text(
            f"{result_text}\n\n"
            f"📖 توضیحات:\n{description}\n\n"
            f"فردا برای سوال جدید برگردید! 🌟",
            reply_markup=InlineKeyboard([("بازگشت به منوی اصلی", "main_menu")])
        )
        return

    if data == "konkor":
        await callback_query.message.edit_text(
            "📝 نمونه سوالات کنکور ایران\n\nرشته مورد نظر خود را انتخاب کنید:",
            reply_markup=InlineKeyboard(
                [("جراحی", "konkor_surgery"), ("مامایی", "konkor_mamaii")],
                [("داخلی دام بزرگ", "konkor_dam_bozorg_internal"), ("داخلی دام کوچک", "konkor_dam_kochik_internal")],
                [("رادیولوژی", "konkor_radiology"), ("کلینیکال پاتولوژی", "konkor_clinical_path")],
                [("بازگشت", "soalat")]
            )
        )
        return

    if data.startswith("konkor_"):
        category_key = data.replace("konkor_", "")

        if category_key in KONKOR_CATEGORIES:
            items = get_konkor_by_category(category_key)

            if not items:
                await callback_query.message.edit_text(
                    "❌ سوالی برای این رشته یافت نشد.",
                    reply_markup=InlineKeyboard([("بازگشت", "konkor")])
                )
                return

            text, kb = build_konkor_page(items, 0, category_key)
            await callback_query.message.edit_text(text, reply_markup=kb)
            return

    if data.startswith("video_"):

        category_key = data.replace("video_", "")

        items = get_videos_by_category(category_key)

        if not items:
            await callback_query.message.edit_text(
                "❌ ویدیویی برای این دسته یافت نشد.",
                reply_markup=InlineKeyboard(
                    [("بازگشت", "vid")]
                )
            )
            return

        text, kb = build_videos_page(items, 0, category_key)

        await callback_query.message.edit_text(text,reply_markup=kb)

        return

    if data == "shop":
        await callback_query.message.edit_text(
            "مجموعه‌ی @ اولین و تنها مجموعه‌ی آموزش دامپزشکی در ایران است که دسترسی حرفه‌ای به اپلیکیشن‌های حرفه‌ای تطبیق مدرک دامپزشکی را مطابق با نسخه‌ها اصلی برای دامپزشکان فراهم کرده. 🤩✌🏻\n\nشما می‌توانید در گزینه‌های زیر ویژگی‌های هر کدام از این اپ‌ها را بررسی کنید.",
            reply_markup=InlineKeyboard(
                [("زوکو", "https://@.com/product/%d8%a7%d9%be%d9%84%db%8c%da%a9%db%8c%d8%b4%d9%86-%d8%b2%d9%88%da%a9%d9%88-%d9%86%d8%b1%d9%85%d8%a7%d9%81%d8%b2%d8%a7%d8%b1-%d8%ac%d8%a7%d9%85%d8%b9-%d8%a2%d9%85%d8%a7%d8%af%da%af%db%8c/","https://@.com/product/%d8%a7%d9%be%d9%84%db%8c%da%a9%db%8c%d8%b4%d9%86-%d8%b2%d9%88%da%a9%d9%88-%d9%86%d8%b1%d9%85%d8%a7%d9%81%d8%b2%d8%a7%d8%b1-%d8%ac%d8%a7%d9%85%d8%b9-%d8%a2%d9%85%d8%a7%d8%af%da%af%db%8c/")],
                [("وت پرپ", "https://@.com/product/%d8%a7%d9%be%d9%84%db%8c%da%a9%db%8c%d8%b4%d9%86-%d9%88%d8%aa-%d9%be%d8%b1%d9%be-%d8%a7%d9%be%d9%84%db%8c%da%a9%db%8c%d8%b4%d9%86-%d8%a2%d8%b2%d9%85%d9%88%d9%86%d9%87%d8%a7%db%8c-%d8%aa/","https://@.com/product/%d8%a7%d9%be%d9%84%db%8c%da%a9%db%8c%d8%b4%d9%86-%d9%88%d8%aa-%d9%be%d8%b1%d9%be-%d8%a7%d9%be%d9%84%db%8c%da%a9%db%8c%d8%b4%d9%86-%d8%a2%d8%b2%d9%85%d9%88%d9%86%d9%87%d8%a7%db%8c-%d8%aa/")],
                [("بازگشت به منوی اصلی", "main_menu")]
            )
        )
        return

    if data == "tatbigh":
        await callback_query.message.edit_text(
            "📋 تطبیق مدرک دامپزشکی\n\nبرای اطلاعات بیشتر به وبلاگ ما مراجعه کنید:",
            reply_markup=InlineKeyboard(
                [("وبلاگ @ 🔗", "https://@.com/blog/", "https://@.com/blog/")],
                [("بازگشت به منوی اصلی", "main_menu")]
            )
        )
        return

    if data == "handouts":
        await callback_query.message.edit_text(
            " 📚 شما در هر دسته بندی به تعداد زیادی فایل جزوه که در دانشکده‌های دامپزشکی مختلف در ایران هستند می‌توانید دسترسی پیدا کنید. این جزوات از طریق کانال‌های مختلف تلگرامی به دست آمده‌اند. برای تکمیل شدن این اطلاعات می‌توانید جزوات خود را به آيدی @@ در «بله» ارسال کنید تا بتوانیم هرچه بیشتر منابع مورد نیاز دامپزشکان را در اختیارشان قرار بدهیم.🙏🏻🌱",
            reply_markup=InlineKeyboard(
                [("دام کوچک", "dam_kochik"), ("دام بزرگ", "dam_bozorg")],
                [("بیهوشی و جراحی", "surg"), ("رادیولوژی", "radio")],
                [("کلینیکال پاتولوژی", "clinical"), ("فارماکولوژی", "pharma")],
                [("پرندگان", "birds"), ("دیگر", "others")],
                [("بازگشت به منوی اصلی", "main_menu")]
            )
        )
        return

    if data in CATEGORY_MAP:
        items = get_handouts_by_category(data)
        if not items:
            await callback_query.message.edit_text(
                "❌ جزوه‌ای برای این دسته یافت نشد.",
                reply_markup=InlineKeyboard([("بازگشت", "handouts")])
            )
            return
        text, kb = build_handouts_page(items, 0, data)
        await callback_query.message.edit_text(text, reply_markup=kb)
        return

    if data == "appointment":
        cleanup_old_slots()

        today = khayyam.JalaliDatetime.now()
        dates = []
        for i in range(4):
            d = today + timedelta(days=i)
            dates.append(d.strftime("%Y/%m/%d"))

        keyboard = InlineKeyboard(
            *[[(d, f"date_{d}")] for d in dates],
            [("بازگشت به منوی اصلی", "main_menu")]
        )

        await callback_query.message.edit_text(
            "🔘 شما می‌توانید برای دریافت مشاوره‌های مهاجرت از طریق معادلسازی مدرک دامپزشکی به کشورهای آمریکا، کانادا، استرالیا، انگلستان، هلند، آلمان، آفریقای جنوبی وقت مشاوره خود را دریافت کنید و مسیر مناسب خود برای معادلسازی مدرک را به صورت اختصاصی دریافت کنید.\n\n📲 این مشاوره به صورت ویدئو کال خواهد بود و به تمام سوال شما در این زمینه پاسخ داده خواهد شد.\n\n🔴 نکته: تمام عزیزانی که از مشاوره‌ی ما استفاده کنند، در صورتی که تمایل به تهیه اپلیکیشن‌ها تطبیق مدرک را داشته باشند در نظر داشته باشند که هزینه‌ی مشاوره آن‌ها از هزینه‌ی تهیه‌ی اپلیکیشن کسر می‌شود.\n\n🗓️ انتحاب تاریخ مورد نظر:",
            reply_markup=keyboard
        )
        return

    if data.startswith("date_"):
        selected_date = data.replace("date_", "")
        times = ["10-11", "16-17", "21-22"]

        available_times = [t for t in times if not is_slot_booked(selected_date, t)]

        if not available_times:
            await callback_query.message.edit_text(
                f"❌ متاسفانه برای تاریخ {selected_date} زمانی خالی نیست.\n\nلطفاً تاریخ دیگری انتخاب کنید.",
                reply_markup=InlineKeyboard([("بازگشت", "appointment")])
            )
            return

        rows = [[(t, f"time_{selected_date}_{t}")] for t in available_times]
        rows.append([("بازگشت", "appointment")])

        await callback_query.message.edit_text(
            f"📅 تاریخ: {selected_date}\nلطفاً ساعت را انتخاب کنید:",
            reply_markup=InlineKeyboard(*rows)
        )
        return

    if data.startswith("time_"):
        _, date, time = data.split("_", 2)

        if is_slot_booked(date, time):
            await callback_query.answer("این زمان قبلاً رزرو شده است", show_alert=True)
            return

        user_states[user_id] = {
            "step": "waiting_name",
            "date": date,
            "time": time
        }

        await callback_query.message.edit_text(
            f"📅 {date} | 🕐 {time}\n\nلطفاً نام و نام خانوادگی خود را وارد کنید:",
            reply_markup=InlineKeyboard([("بازگشت به منوی اصلی", "main_menu")])
        )
        return

if __name__ == "__main__":
    init_db()
    bot.run()
