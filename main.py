import asyncio
import html
import logging
import os
import sqlite3
from datetime import datetime, timedelta
from pathlib import Path

from aiohttp import web
from aiogram import Bot, Dispatcher, F
from aiogram.filters import Command, CommandStart
from aiogram.types import (
    CallbackQuery,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    LabeledPrice,
    Message,
    PreCheckoutQuery,
)
from google import genai
from google.genai import types
from pypdf import PdfReader
from docx import Document


# ============================================================
# CONFIG
# ============================================================

BOT_TOKEN = os.getenv("BOT_TOKEN")
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
ADMIN_ID = int(os.getenv("ADMIN_ID", "0"))

# Основная бесплатная модель
GEMINI_MODEL = os.getenv(
    "GEMINI_MODEL",
    "gemini-3.6-flash"
)

# Резервная бесплатная модель
GEMINI_FALLBACK_MODEL = os.getenv(
    "GEMINI_FALLBACK_MODEL",
    "gemini-3.1-flash-lite"
)

PORT = int(os.getenv("PORT", "10000"))

DB_PATH = "zolog_ai.db"
FILES_DIR = Path("user_files")
FILES_DIR.mkdir(exist_ok=True)

START_GENERATIONS = 10
REFERRAL_REWARD = 1

# Повторные попытки
RETRY_DELAYS = [3, 6, 12, 20]

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s"
)

logger = logging.getLogger("zolog-ai")


# ============================================================
# CHECK CONFIG
# ============================================================

if not BOT_TOKEN:
    raise RuntimeError("BOT_TOKEN не найден в Environment Variables")

if not GEMINI_API_KEY:
    raise RuntimeError("GEMINI_API_KEY не найден в Environment Variables")


# ============================================================
# BOT / AI
# ============================================================

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()

gemini = genai.Client(
    api_key=GEMINI_API_KEY
)


# ============================================================
# DATABASE
# ============================================================

db = sqlite3.connect(
    DB_PATH,
    check_same_thread=False
)

db.row_factory = sqlite3.Row


def init_db():
    cursor = db.cursor()

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS users (
            telegram_id INTEGER PRIMARY KEY,
            username TEXT,
            first_name TEXT,
            language TEXT DEFAULT 'ru',
            generations INTEGER DEFAULT 10,
            referrals_count INTEGER DEFAULT 0,
            referred_by INTEGER,
            is_blocked INTEGER DEFAULT 0,
            created_at TEXT
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS materials (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            telegram_id INTEGER,
            material_type TEXT,
            title TEXT,
            content TEXT,
            created_at TEXT
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS books (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            telegram_id INTEGER,
            filename TEXT,
            file_path TEXT,
            extracted_text TEXT,
            created_at TEXT
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS payments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            telegram_id INTEGER,
            package TEXT,
            stars INTEGER,
            generations INTEGER,
            telegram_charge_id TEXT,
            created_at TEXT
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS referrals (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            inviter_id INTEGER,
            invited_id INTEGER,
            created_at TEXT
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            telegram_id INTEGER,
            action TEXT,
            details TEXT,
            created_at TEXT
        )
    """)

    db.commit()


init_db()


# ============================================================
# HELPERS
# ============================================================

def now():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def log_action(
    telegram_id: int,
    action: str,
    details: str = ""
):
    try:
        db.execute(
            """
            INSERT INTO logs
            (telegram_id, action, details, created_at)
            VALUES (?, ?, ?, ?)
            """,
            (
                telegram_id,
                action,
                details,
                now()
            )
        )
        db.commit()
    except Exception as e:
        logger.error(f"Ошибка логирования: {e}")


def get_user(telegram_id: int):
    return db.execute(
        "SELECT * FROM users WHERE telegram_id = ?",
        (telegram_id,)
    ).fetchone()


def create_user(message: Message):
    user = message.from_user

    existing = get_user(user.id)

    if existing:
        db.execute(
            """
            UPDATE users
            SET username = ?, first_name = ?
            WHERE telegram_id = ?
            """,
            (
                user.username,
                user.first_name,
                user.id
            )
        )
        db.commit()
        return existing

    db.execute(
        """
        INSERT INTO users
        (
            telegram_id,
            username,
            first_name,
            language,
            generations,
            created_at
        )
        VALUES (?, ?, ?, ?, ?, ?)
        """,
        (
            user.id,
            user.username,
            user.first_name,
            "ru",
            START_GENERATIONS,
            now()
        )
    )

    db.commit()

    return get_user(user.id)


def add_generations(
    telegram_id: int,
    amount: int
):
    db.execute(
        """
        UPDATE users
        SET generations = generations + ?
        WHERE telegram_id = ?
        """,
        (
            amount,
            telegram_id
        )
    )

    db.commit()


def remove_generation(telegram_id: int):
    user = get_user(telegram_id)

    if not user:
        return False

    if user["generations"] <= 0:
        return False

    db.execute(
        """
        UPDATE users
        SET generations = generations - 1
        WHERE telegram_id = ?
        """,
        (telegram_id,)
    )

    db.commit()

    return True


def user_is_blocked(telegram_id: int):
    user = get_user(telegram_id)

    if not user:
        return False

    return bool(user["is_blocked"])


# ============================================================
# AI ERROR DETECTION
# ============================================================

def is_rate_limit_error(error_text: str):
    text = error_text.lower()

    keywords = [
        "429",
        "resource_exhausted",
        "quota",
        "rate limit",
        "too many requests",
        "requests per minute",
        "requests per day",
        "quota exceeded",
    ]

    return any(
        keyword in text
        for keyword in keywords
    )


def is_temporary_error(error_text: str):
    text = error_text.lower()

    keywords = [
        "503",
        "unavailable",
        "high demand",
        "temporarily unavailable",
        "internal server error",
        "deadline exceeded",
        "timeout",
        "timed out",
    ]

    return any(
        keyword in text
        for keyword in keywords
    )


# ============================================================
# GEMINI AI
# ============================================================

async def ask_ai(prompt: str):
    """
    Система AI:

    1. Основная модель.
    2. Если 429/лимит -> резервная модель.
    3. Если 503/перегрузка -> повтор.
    4. Повторные попытки:
       3 -> 6 -> 12 -> 20 секунд.
    5. Если основная модель не справилась,
       переходим к резервной.
    """

    models = [
        GEMINI_MODEL,
        GEMINI_FALLBACK_MODEL
    ]

    # Не используем одну и ту же модель дважды
    unique_models = []

    for model in models:
        if model and model not in unique_models:
            unique_models.append(model)

    last_error = None

    for model_index, model in enumerate(unique_models):

        logger.info(
            f"🤖 Используем модель: {model}"
        )

        attempt = 0

        while attempt <= len(RETRY_DELAYS):

            try:
                response = await asyncio.to_thread(
                    gemini.models.generate_content,
                    model=model,
                    contents=prompt,
                    config=types.GenerateContentConfig(
                        temperature=0.7
                    )
                )

                if response is None:
                    raise RuntimeError(
                        "Gemini вернул пустой ответ"
                    )

                text = getattr(
                    response,
                    "text",
                    None
                )

                if text:
                    logger.info(
                        f"✅ Ответ получен через {model}"
                    )

                    return text

                raise RuntimeError(
                    "Gemini вернул ответ без текста"
                )

            except Exception as e:

                error_text = str(e)
                last_error = error_text

                logger.error(
                    f"❌ Ошибка {model}: {error_text}"
                )

                # --------------------------------------------
                # 429 / QUOTA
                # --------------------------------------------

                if is_rate_limit_error(error_text):

                    logger.warning(
                        f"⚠️ Лимит модели {model}"
                    )

                    # Если есть резервная модель
                    if model_index < len(unique_models) - 1:

                        fallback = unique_models[
                            model_index + 1
                        ]

                        logger.warning(
                            f"🔄 Переключение "
                            f"{model} -> {fallback}"
                        )

                        break

                    # Резервная тоже получила лимит.
                    # Делаем повторные попытки.
                    if attempt < len(RETRY_DELAYS):

                        delay = RETRY_DELAYS[attempt]

                        logger.info(
                            f"⏳ Повтор через {delay} сек."
                        )

                        await asyncio.sleep(delay)

                        attempt += 1

                        continue

                    break

                # --------------------------------------------
                # 503 / TEMPORARY
                # --------------------------------------------

                if is_temporary_error(error_text):

                    if attempt < len(RETRY_DELAYS):

                        delay = RETRY_DELAYS[attempt]

                        logger.warning(
                            f"⏳ Временная ошибка. "
                            f"Повтор через {delay} сек."
                        )

                        await asyncio.sleep(delay)

                        attempt += 1

                        continue

                    # Если основная модель не отвечает,
                    # переключаемся на резервную.
                    if model_index < len(unique_models) - 1:

                        fallback = unique_models[
                            model_index + 1
                        ]

                        logger.warning(
                            f"🔄 {model} недоступна. "
                            f"Переходим на {fallback}"
                        )

                        break

                    break

                # --------------------------------------------
                # ДРУГАЯ ОШИБКА
                # --------------------------------------------

                logger.error(
                    f"❌ Неповторяемая ошибка: {error_text}"
                )

                # На всякий случай пробуем резервную
                if model_index < len(unique_models) - 1:

                    logger.warning(
                        f"🔄 Пробуем резервную модель "
                        f"{unique_models[model_index + 1]}"
                    )

                    break

                break

    logger.error(
        f"❌ Все AI-модели недоступны: {last_error}"
    )

    return None


# ============================================================
# KEYBOARDS
# ============================================================

def main_menu():
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(
                    text="📝 Создать материал",
                    callback_data="create_material"
                )
            ],
            [
                InlineKeyboardButton(
                    text="🧠 Спросить AI",
                    callback_data="ask_ai"
                )
            ],
            [
                InlineKeyboardButton(
                    text="📖 Загрузить книгу",
                    callback_data="upload_book"
                )
            ],
            [
                InlineKeyboardButton(
                    text="📚 Моя библиотека",
                    callback_data="library"
                )
            ],
            [
                InlineKeyboardButton(
                    text="👤 Профиль",
                    callback_data="profile"
                )
            ],
            [
                InlineKeyboardButton(
                    text="🎁 Пригласить друга",
                    callback_data="referral"
                )
            ],
            [
                InlineKeyboardButton(
                    text="⭐ Получить генерации",
                    callback_data="buy_generations"
                )
            ],
            [
                InlineKeyboardButton(
                    text="🌐 Язык",
                    callback_data="language"
                ),
                InlineKeyboardButton(
                    text="⚙️ Настройки",
                    callback_data="settings"
                )
            ],
            [
                InlineKeyboardButton(
                    text="ℹ️ Помощь",
                    callback_data="help"
                )
            ]
        ]
    )


def material_menu():
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(
                    text="📄 Реферат",
                    callback_data="material_ref"
                ),
                InlineKeyboardButton(
                    text="📚 Курсовая",
                    callback_data="material_course"
                )
            ],
            [
                InlineKeyboardButton(
                    text="🎓 Дипломная",
                    callback_data="material_diploma"
                ),
                InlineKeyboardButton(
                    text="📖 Конспект",
                    callback_data="material_notes"
                )
            ],
            [
                InlineKeyboardButton(
                    text="📊 Презентация",
                    callback_data="material_presentation"
                ),
                InlineKeyboardButton(
                    text="📝 Эссе",
                    callback_data="material_essay"
                )
            ],
            [
                InlineKeyboardButton(
                    text="🔬 Доклад",
                    callback_data="material_report"
                ),
                InlineKeyboardButton(
                    text="🧠 Свой запрос",
                    callback_data="material_custom"
                )
            ],
            [
                InlineKeyboardButton(
                    text="◀️ Назад",
                    callback_data="back_main"
                )
            ]
        ]
    )


def language_menu():
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(
                    text="🇺🇦 Українська",
                    callback_data="lang_uk"
                )
            ],
            [
                InlineKeyboardButton(
                    text="🇷🇺 Русский",
                    callback_data="lang_ru"
                )
            ],
            [
                InlineKeyboardButton(
                    text="🇬🇧 English",
                    callback_data="lang_en"
                )
            ],
            [
                InlineKeyboardButton(
                    text="◀️ Назад",
                    callback_data="back_main"
                )
            ]
        ]
    )


def payment_menu():
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(
                    text="⭐ 50 → 100 генераций",
                    callback_data="buy_basic"
                )
            ],
            [
                InlineKeyboardButton(
                    text="⭐ 150 → 500 генераций",
                    callback_data="buy_pro"
                )
            ],
            [
                InlineKeyboardButton(
                    text="⭐ 350 → 1500 генераций",
                    callback_data="buy_premium"
                )
            ],
            [
                InlineKeyboardButton(
                    text="◀️ Назад",
                    callback_data="back_main"
                )
            ]
        ]
    )


def admin_menu():
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(
                    text="📊 Dashboard",
                    callback_data="admin_dashboard"
                )
            ],
            [
                InlineKeyboardButton(
                    text="👥 Пользователи",
                    callback_data="admin_users"
                )
            ],
            [
                InlineKeyboardButton(
                    text="🗄️ Материалы",
                    callback_data="admin_materials"
                )
            ],
            [
                InlineKeyboardButton(
                    text="📚 Книги",
                    callback_data="admin_books"
                )
            ],
            [
                InlineKeyboardButton(
                    text="🤖 AI",
                    callback_data="admin_ai"
                )
            ],
            [
                InlineKeyboardButton(
                    text="🎁 Рефералы",
                    callback_data="admin_referrals"
                )
            ],
            [
                InlineKeyboardButton(
                    text="💳 Оплаты",
                    callback_data="admin_payments"
                )
            ],
            [
                InlineKeyboardButton(
                    text="📈 Статистика оплат",
                    callback_data="admin_payment_stats"
                )
            ],
            [
                InlineKeyboardButton(
                    text="📢 Рассылки",
                    callback_data="admin_broadcast"
                )
            ],
            [
                InlineKeyboardButton(
                    text="🛡️ Модерация",
                    callback_data="admin_moderation"
                )
            ],
            [
                InlineKeyboardButton(
                    text="📝 Логи",
                    callback_data="admin_logs"
                )
            ],
            [
                InlineKeyboardButton(
                    text="🧪 Test Mode",
                    callback_data="admin_test"
                )
            ],
            [
                InlineKeyboardButton(
                    text="⚙️ Настройки",
                    callback_data="admin_settings"
                )
            ],
            [
                InlineKeyboardButton(
                    text="◀️ В главное меню",
                    callback_data="back_main"
                )
            ]
        ]
    )


# ============================================================
# USER STATES
# ============================================================

# Так как проект в одном файле,
# состояние пользователя храним здесь.
user_states = {}


def set_state(
    telegram_id: int,
    state: str
):
    user_states[telegram_id] = state


def get_state(telegram_id: int):
    return user_states.get(telegram_id)


def clear_state(telegram_id: int):
    user_states.pop(telegram_id, None)


# ============================================================
# /START
# ============================================================

@dp.message(CommandStart())
async def start_handler(message: Message):

    user = create_user(message)

    telegram_id = message.from_user.id

    # --------------------------------------------
    # REFERRAL
    # --------------------------------------------

    args = message.text.split(maxsplit=1)

    if len(args) > 1:

        referral_code = args[1]

        if referral_code.startswith("ref_"):

            try:
                inviter_id = int(
                    referral_code.replace(
                        "ref_",
                        "",
                        1
                    )
                )

                # Нельзя пригласить самого себя
                if inviter_id != telegram_id:

                    current_user = get_user(
                        telegram_id
                    )

                    # Только если пользователь новый
                    # и ещё не имеет пригласившего
                    if (
                        current_user
                        and current_user["referred_by"] is None
                    ):

                        inviter = get_user(
                            inviter_id
                        )

                        if inviter:

                            db.execute(
                                """
                                UPDATE users
                                SET referred_by = ?
                                WHERE telegram_id = ?
                                """,
                                (
                                    inviter_id,
                                    telegram_id
                                )
                            )

                            db.execute(
                                """
                                UPDATE users
                                SET
                                    referrals_count =
                                    referrals_count + 1,
                                    generations =
                                    generations + ?
                                WHERE telegram_id = ?
                                """,
                                (
                                    REFERRAL_REWARD,
                                    inviter_id
                                )
                            )

                            db.execute(
                                """
                                INSERT INTO referrals
                                (
                                    inviter_id,
                                    invited_id,
                                    created_at
                                )
                                VALUES (?, ?, ?)
                                """,
                                (
                                    inviter_id,
                                    telegram_id,
                                    now()
                                )
                            )

                            db.commit()

                            log_action(
                                inviter_id,
                                "referral",
                                f"Приглашён пользователь {telegram_id}"
                            )

                            try:
                                await bot.send_message(
                                    inviter_id,
                                    "🎉 Новый пользователь "
                                    "перешёл по вашей ссылке!\n\n"
                                    f"Вам начислена "
                                    f"+{REFERRAL_REWARD} генерация."
                                )
                            except Exception:
                                pass

            except ValueError:
                pass

    clear_state(telegram_id)

    await message.answer(
        "🤖 <b>Zolog AI</b>\n\n"
        "Добро пожаловать!\n\n"
        "Я помогу создавать учебные материалы, "
        "работать с книгами и отвечать на вопросы.\n\n"
        f"🎁 Ваш баланс: "
        f"<b>{user['generations']}</b> генераций",
        reply_markup=main_menu(),
        parse_mode="HTML"
    )


# ============================================================
# TEXT MAIN COMMAND
# ============================================================

@dp.message(Command("menu"))
async def menu_command(message: Message):

    create_user(message)

    clear_state(
        message.from_user.id
    )

    await message.answer(
        "🏠 Главное меню",
        reply_markup=main_menu()
    )


# ============================================================
# PROFILE
# ============================================================

async def show_profile(
    telegram_id: int,
    target: Message | CallbackQuery
):

    user = get_user(telegram_id)

    if not user:
        return

    username = (
        f"@{user['username']}"
        if user["username"]
        else "не указан"
    )

    text = (
        "👤 <b>Профиль</b>\n\n"
        f"Имя: <b>{html.escape(user['first_name'] or '')}</b>\n"
        f"Username: {username}\n"
        f"ID: <code>{telegram_id}</code>\n\n"
        f"⭐ Генераций: <b>{user['generations']}</b>\n"
        f"🎁 Приглашено: <b>{user['referrals_count']}</b>\n"
        f"🌐 Язык: <b>{user['language']}</b>"
    )

    keyboard = InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(
                    text="◀️ Назад",
                    callback_data="back_main"
                )
            ]
        ]
    )

    if isinstance(target, CallbackQuery):

        await target.message.edit_text(
            text,
            reply_markup=keyboard,
            parse_mode="HTML"
        )

    else:

        await target.answer(
            text,
            reply_markup=keyboard,
            parse_mode="HTML"
        )


# ============================================================
# REFERRAL
# ============================================================

async def show_referral(
    telegram_id: int,
    target: CallbackQuery
):

    user = get_user(telegram_id)

    if not user:
        return

    bot_info = await bot.get_me()

    link = (
        f"https://t.me/"
        f"{bot_info.username}"
        f"?start=ref_{telegram_id}"
    )

    text = (
        "🎁 <b>Реферальная система</b>\n\n"
        "Приглашай друзей в Zolog AI.\n\n"
        f"За каждого нового пользователя "
        f"ты получаешь <b>+{REFERRAL_REWARD}</b> генерацию.\n\n"
        f"👥 Приглашено: "
        f"<b>{user['referrals_count']}</b>\n\n"
        "🔗 Твоя ссылка:\n"
        f"<code>{link}</code>"
    )

    keyboard = InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(
                    text="◀️ Назад",
                    callback_data="back_main"
                )
            ]
        ]
    )

    await target.message.edit_text(
        text,
        reply_markup=keyboard,
        parse_mode="HTML"
    )


# ============================================================
# MATERIAL GENERATION
# ============================================================

MATERIAL_NAMES = {
    "material_ref": "Реферат",
    "material_course": "Курсовая работа",
    "material_diploma": "Дипломная работа",
    "material_notes": "Конспект",
    "material_presentation": "Презентация",
    "material_essay": "Эссе",
    "material_report": "Доклад",
}


@dp.callback_query(F.data == "create_material")
async def create_material_callback(
    callback: CallbackQuery
):

    await callback.answer()

    clear_state(
        callback.from_user.id
    )

    await callback.message.edit_text(
        "📝 <b>Выберите тип материала:</b>",
        reply_markup=material_menu(),
        parse_mode="HTML"
    )


@dp.callback_query(
    F.data.in_(list(MATERIAL_NAMES.keys()))
)
async def material_type_callback(
    callback: CallbackQuery
):

    await callback.answer()

    material_type = MATERIAL_NAMES[
        callback.data
    ]

    telegram_id = callback.from_user.id

    set_state(
        telegram_id,
        f"material:{material_type}"
    )

    await callback.message.edit_text(
        f"📝 Вы выбрали: <b>{material_type}</b>\n\n"
        "Теперь отправьте мне тему.\n\n"
        "Например:\n"
        "<i>Физическая терапия после инсульта</i>",
        parse_mode="HTML"
    )


@dp.callback_query(F.data == "material_custom")
async def material_custom_callback(
    callback: CallbackQuery
):

    await callback.answer()

    set_state(
        callback.from_user.id,
        "material:Свой запрос"
    )

    await callback.message.edit_text(
        "🧠 <b>Свой запрос</b>\n\n"
        "Напишите подробно, что должен создать Zolog AI.",
        parse_mode="HTML"
    )


# ============================================================
# ASK AI
# ============================================================

@dp.callback_query(F.data == "ask_ai")
async def ask_ai_callback(
    callback: CallbackQuery
):

    await callback.answer()

    set_state(
        callback.from_user.id,
        "ask_ai"
    )

    await callback.message.edit_text(
        "🧠 <b>Задайте вопрос AI</b>\n\n"
        "Напишите свой вопрос следующим сообщением.",
        parse_mode="HTML"
    )


# ============================================================
# MESSAGE PROCESSOR
# ============================================================

@dp.message(F.text)
async def text_handler(message: Message):

    telegram_id = message.from_user.id

    if user_is_blocked(telegram_id):

        await message.answer(
            "🚫 Ваш аккаунт заблокирован."
        )

        return

    create_user(message)

    state = get_state(telegram_id)

    if not state:

        await message.answer(
            "Выберите действие в меню:",
            reply_markup=main_menu()
        )

        return

    # --------------------------------------------
    # ASK AI
    # --------------------------------------------

    if state == "ask_ai":

        if not remove_generation(
            telegram_id
        ):

            await message.answer(
                "❌ У вас закончились генерации.\n\n"
                "Получить новые можно через "
                "⭐ «Получить генерации»."
            )

            clear_state(telegram_id)

            return

        prompt = (
            "Ты — AI-помощник Zolog AI.\n"
            "Отвечай понятно, подробно и по существу.\n"
            "Если вопрос учебный — объясняй структурировано.\n\n"
            f"Вопрос пользователя:\n{message.text}"
        )

        await message.answer(
            "🤖 Думаю над ответом..."
        )

        result = await ask_ai(prompt)

        if result is None:

            add_generations(
                telegram_id,
                1
            )

            await message.answer(
                "⚠️ Сейчас AI временно недоступен.\n\n"
                "Генерация возвращена на баланс."
            )

        else:

            await message.answer(
                result
            )

        log_action(
            telegram_id,
            "ask_ai",
            message.text[:500]
        )

        clear_state(telegram_id)

        return

    # --------------------------------------------
    # MATERIAL
    # --------------------------------------------

    if state.startswith("material:"):

        material_type = state.split(
            ":",
            1
        )[1]

        topic = message.text.strip()

        if not remove_generation(
            telegram_id
        ):

            await message.answer(
                "❌ У вас закончились генерации.\n\n"
                "Получите новые генерации через меню."
            )

            clear_state(telegram_id)

            return

        await message.answer(
            "🤖 Начинаю подготовку материала...\n\n"
            "Это может занять некоторое время."
        )

        prompt = f"""
Ты — профессиональный AI-сервис Zolog AI.

Создай учебный материал.

Тип:
{material_type}

Тема:
{topic}

Требования:
- писать на русском языке;
- структурировать материал;
- использовать заголовки;
- раскрыть тему содержательно;
- не выдумывать конкретные источники;
- если это учебная работа, добавить введение,
  основную часть и выводы;
- текст должен быть пригоден для дальнейшего
  редактирования и оформления.

Материал:
"""

        result = await ask_ai(prompt)

        if result is None:

            add_generations(
                telegram_id,
                1
            )

            await message.answer(
                "⚠️ Не удалось получить материал.\n\n"
                "Генерация возвращена."
            )

        else:

            title = topic[:200]

            db.execute(
                """
                INSERT INTO materials
                (
                    telegram_id,
                    material_type,
                    title,
                    content,
                    created_at
                )
                VALUES (?, ?, ?, ?, ?)
                """,
                (
                    telegram_id,
                    material_type,
                    title,
                    result,
                    now()
                )
            )

            db.commit()

            await message.answer(
                "✅ <b>Материал готов!</b>\n\n"
                f"<b>{html.escape(topic)}</b>\n\n"
                f"{result}",
                parse_mode="HTML"
            )

            log_action(
                telegram_id,
                "material_created",
                material_type
            )

        clear_state(telegram_id)

        return


# ============================================================
# BOOK UPLOAD
# ============================================================

@dp.callback_query(F.data == "upload_book")
async def upload_book_callback(
    callback: CallbackQuery
):

    await callback.answer()

    set_state(
        callback.from_user.id,
        "upload_book"
    )

    await callback.message.edit_text(
        "📖 <b>Загрузка книги</b>\n\n"
        "Отправьте PDF, DOCX или TXT-файл.\n\n"
        "После загрузки книга появится в вашей "
        "личной библиотеке.",
        parse_mode="HTML"
    )


@dp.message(F.document)
async def document_handler(message: Message):

    telegram_id = message.from_user.id

    create_user(message)

    state = get_state(telegram_id)

    if state != "upload_book":

        await message.answer(
            "📖 Если хотите добавить этот файл "
            "в библиотеку, сначала нажмите "
            "«📖 Загрузить книгу»."
        )

        return

    document = message.document

    filename = document.file_name or "file"

    extension = (
        Path(filename)
        .suffix
        .lower()
    )

    allowed = [
        ".pdf",
        ".docx",
        ".txt"
    ]

    if extension not in allowed:

        await message.answer(
            "❌ Поддерживаются только:\n"
            "PDF, DOCX и TXT."
        )

        return

    await message.answer(
        "📥 Загружаю и анализирую файл..."
    )

    user_dir = FILES_DIR / str(
        telegram_id
    )

    user_dir.mkdir(
        parents=True,
        exist_ok=True
    )

    safe_filename = (
        f"{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        f"_{filename}"
    )

    file_path = user_dir / safe_filename

    try:

        telegram_file = await bot.get_file(
            document.file_id
        )

        await bot.download_file(
            telegram_file.file_path,
            destination=file_path
        )

        extracted_text = ""

        # --------------------------------------------
        # PDF
        # --------------------------------------------

        if extension == ".pdf":

            reader = PdfReader(
                str(file_path)
            )

            pages = []

            for page in reader.pages:

                try:
                    text = page.extract_text()

                    if text:
                        pages.append(text)

                except Exception:
                    continue

            extracted_text = "\n".join(
                pages
            )

        # --------------------------------------------
        # DOCX
        # --------------------------------------------

        elif extension == ".docx":

            doc = Document(
                str(file_path)
            )

            extracted_text = "\n".join(
                paragraph.text
                for paragraph in doc.paragraphs
                if paragraph.text.strip()
            )

        # --------------------------------------------
        # TXT
        # --------------------------------------------

        elif extension == ".txt":

            extracted_text = file_path.read_text(
                encoding="utf-8",
                errors="ignore"
            )

        db.execute(
            """
            INSERT INTO books
            (
                telegram_id,
                filename,
                file_path,
                extracted_text,
                created_at
            )
            VALUES (?, ?, ?, ?, ?)
            """,
            (
                telegram_id,
                filename,
                str(file_path),
                extracted_text,
                now()
            )
        )

        db.commit()

        clear_state(telegram_id)

        log_action(
            telegram_id,
            "book_uploaded",
            filename
        )

        await message.answer(
            "✅ <b>Книга добавлена!</b>\n\n"
            f"📖 {html.escape(filename)}\n"
            f"📄 Символов извлечено: "
            f"<b>{len(extracted_text)}</b>\n\n"
            "Теперь она доступна в "
            "«📚 Моей библиотеке».",
            parse_mode="HTML",
            reply_markup=main_menu()
        )

    except Exception as e:

        logger.exception(
            "Ошибка обработки книги"
        )

        await message.answer(
            "❌ Не удалось обработать файл.\n\n"
            f"Ошибка: {str(e)[:500]}"
        )


# ============================================================
# LIBRARY
# ============================================================

@dp.callback_query(F.data == "library")
async def library_callback(
    callback: CallbackQuery
):

    await callback.answer()

    telegram_id = callback.from_user.id

    books = db.execute(
        """
        SELECT *
        FROM books
        WHERE telegram_id = ?
        ORDER BY id DESC
        LIMIT 10
        """,
        (telegram_id,)
    ).fetchall()

    materials = db.execute(
        """
        SELECT *
        FROM materials
        WHERE telegram_id = ?
        ORDER BY id DESC
        LIMIT 10
        """,
        (telegram_id,)
    ).fetchall()

    text = "📚 <b>Моя библиотека</b>\n\n"

    text += "📖 <b>Книги:</b>\n"

    if books:

        for book in books:

            text += (
                f"• {html.escape(book['filename'])}\n"
            )

    else:

        text += "Пока нет загруженных книг.\n"

    text += "\n📝 <b>Материалы:</b>\n"

    if materials:

        for material in materials:

            text += (
                f"• {html.escape(material['title'])} "
                f"— {html.escape(material['material_type'])}\n"
            )

    else:

        text += "Пока нет созданных материалов.\n"

    keyboard = InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(
                    text="◀️ Назад",
                    callback_data="back_main"
                )
            ]
        ]
    )

    await callback.message.edit_text(
        text,
        reply_markup=keyboard,
        parse_mode="HTML"
    )


# ============================================================
# PROFILE CALLBACK
# ============================================================

@dp.callback_query(F.data == "profile")
async def profile_callback(
    callback: CallbackQuery
):

    await callback.answer()

    await show_profile(
        callback.from_user.id,
        callback
    )


# ============================================================
# REFERRAL CALLBACK
# ============================================================

@dp.callback_query(F.data == "referral")
async def referral_callback(
    callback: CallbackQuery
):

    await callback.answer()

    await show_referral(
        callback.from_user.id,
        callback
    )


# ============================================================
# LANGUAGE
# ============================================================

@dp.callback_query(F.data == "language")
async def language_callback(
    callback: CallbackQuery
):

    await callback.answer()

    await callback.message.edit_text(
        "🌐 <b>Выберите язык:</b>",
        reply_markup=language_menu(),
        parse_mode="HTML"
    )


@dp.callback_query(
    F.data.in_(
        [
            "lang_uk",
            "lang_ru",
            "lang_en"
        ]
    )
)
async def language_change(
    callback: CallbackQuery
):

    languages = {
        "lang_uk": "uk",
        "lang_ru": "ru",
        "lang_en": "en"
    }

    language = languages[
        callback.data
    ]

    db.execute(
        """
        UPDATE users
        SET language = ?
        WHERE telegram_id = ?
        """,
        (
            language,
            callback.from_user.id
        )
    )

    db.commit()

    await callback.answer(
        "Язык сохранён!"
    )

    await callback.message.edit_text(
        "✅ Язык успешно изменён.",
        reply_markup=InlineKeyboardMarkup(
            inline_keyboard=[
                [
                    InlineKeyboardButton(
                        text="◀️ Главное меню",
                        callback_data="back_main"
                    )
                ]
            ]
        )
    )


# ============================================================
# SETTINGS
# ============================================================

@dp.callback_query(F.data == "settings")
async def settings_callback(
    callback: CallbackQuery
):

    await callback.answer()

    await callback.message.edit_text(
        "⚙️ <b>Настройки</b>\n\n"
        "Здесь будут доступны настройки "
        "Zolog AI.\n\n"
        "Раздел подготовлен для дальнейшего "
        "расширения.",
        reply_markup=InlineKeyboardMarkup(
            inline_keyboard=[
                [
                    InlineKeyboardButton(
                        text="◀️ Назад",
                        callback_data="back_main"
                    )
                ]
            ]
        ),
        parse_mode="HTML"
    )


# ============================================================
# HELP
# ============================================================

@dp.callback_query(F.data == "help")
async def help_callback(
    callback: CallbackQuery
):

    await callback.answer()

    await callback.message.edit_text(
        "ℹ️ <b>Помощь Zolog AI</b>\n\n"
        "📝 Создать материал — создание "
        "рефератов, курсовых, конспектов и т.д.\n\n"
        "🧠 Спросить AI — задать обычный вопрос.\n\n"
        "📖 Загрузить книгу — добавить PDF, DOCX "
        "или TXT в библиотеку.\n\n"
        "📚 Библиотека — ваши книги и материалы.\n\n"
        "🎁 Пригласить друга — получить генерации "
        "за приглашённых пользователей.\n\n"
        "⭐ Получить генерации — купить генерации "
        "за Telegram Stars.",
        reply_markup=InlineKeyboardMarkup(
            inline_keyboard=[
                [
                    InlineKeyboardButton(
                        text="◀️ Назад",
                        callback_data="back_main"
                    )
                ]
            ]
        ),
        parse_mode="HTML"
    )


# ============================================================
# PAYMENTS
# ============================================================

PACKAGES = {
    "basic": {
        "name": "BASIC",
        "stars": 50,
        "generations": 100
    },
    "pro": {
        "name": "PRO",
        "stars": 150,
        "generations": 500
    },
    "premium": {
        "name": "PREMIUM",
        "stars": 350,
        "generations": 1500
    }
}


@dp.callback_query(F.data == "buy_generations")
async def buy_generations_callback(
    callback: CallbackQuery
):

    await callback.answer()

    await callback.message.edit_text(
        "⭐ <b>Получить генерации</b>\n\n"
        "Выберите пакет:",
        reply_markup=payment_menu(),
        parse_mode="HTML"
    )


async def send_package_invoice(
    telegram_id: int,
    package_key: str
):

    package = PACKAGES[
        package_key
    ]

    payload = (
        f"zolog_{package_key}_"
        f"{telegram_id}_"
        f"{datetime.now().timestamp()}"
    )

    await bot.send_invoice(
        chat_id=telegram_id,
        title=f"Zolog AI — {package['name']}",
        description=(
            f"{package['generations']} "
            "генераций Zolog AI"
        ),
        payload=payload,
        currency="XTR",
        prices=[
            LabeledPrice(
                label=f"{package['generations']} генераций",
                amount=package["stars"]
            )
        ]
    )


@dp.callback_query(
    F.data.in_(
        [
            "buy_basic",
            "buy_pro",
            "buy_premium"
        ]
    )
)
async def package_callback(
    callback: CallbackQuery
):

    await callback.answer()

    mapping = {
        "buy_basic": "basic",
        "buy_pro": "pro",
        "buy_premium": "premium"
    }

    package_key = mapping[
        callback.data
    ]

    try:

        await send_package_invoice(
            callback.from_user.id,
            package_key
        )

    except Exception as e:

        logger.exception(
            "Ошибка создания платежа"
        )

        await callback.message.answer(
            "❌ Не удалось создать платёж.\n\n"
            f"{str(e)[:500]}"
        )


# ============================================================
# PRE-CHECKOUT
# ============================================================

@dp.pre_checkout_query()
async def pre_checkout_handler(
    query: PreCheckoutQuery
):

    await query.answer(
        ok=True
    )


# ============================================================
# SUCCESSFUL PAYMENT
# ============================================================

@dp.message(
    F.successful_payment
)
async def successful_payment_handler(
    message: Message
):

    payment = message.successful_payment

    telegram_id = message.from_user.id

    charge_id = (
        payment.telegram_payment_charge_id
    )

    # Защита от повторного начисления
    existing = db.execute(
        """
        SELECT id
        FROM payments
        WHERE telegram_charge_id = ?
        """,
        (charge_id,)
    ).fetchone()

    if existing:

        await message.answer(
            "⚠️ Этот платёж уже был обработан."
        )

        return

    package_key = None

    for key, package in PACKAGES.items():

        if package["stars"] == payment.total_amount:

            package_key = key
            break

    if package_key is None:

        await message.answer(
            "⚠️ Платёж получен, "
            "но пакет не удалось определить.\n"
            "Обратитесь к администратору."
        )

        return

    package = PACKAGES[
        package_key
    ]

    db.execute(
        """
        INSERT INTO payments
        (
            telegram_id,
            package,
            stars,
            generations,
            telegram_charge_id,
            created_at
        )
        VALUES (?, ?, ?, ?, ?, ?)
        """,
        (
            telegram_id,
            package["name"],
            package["stars"],
            package["generations"],
            charge_id,
            now()
        )
    )

    db.commit()

    add_generations(
        telegram_id,
        package["generations"]
    )

    log_action(
        telegram_id,
        "payment",
        f"{package['name']} +{package['generations']}"
    )

    user = get_user(
        telegram_id
    )

    await message.answer(
        "✅ <b>Оплата прошла успешно!</b>\n\n"
        f"⭐ Пакет: <b>{package['name']}</b>\n"
        f"➕ Начислено: "
        f"<b>{package['generations']}</b> генераций\n"
        f"💰 Текущий баланс: "
        f"<b>{user['generations']}</b>",
        parse_mode="HTML",
        reply_markup=main_menu()
    )


# ============================================================
# ADMIN CHECK
# ============================================================

def is_admin(telegram_id: int):
    return (
        ADMIN_ID != 0
        and telegram_id == ADMIN_ID
    )


async def admin_required(
    callback: CallbackQuery
):

    if not is_admin(
        callback.from_user.id
    ):

        await callback.answer(
            "⛔ Нет доступа.",
            show_alert=True
        )

        return False

    return True


# ============================================================
# ADMIN COMMAND
# ============================================================

@dp.message(Command("admin"))
async def admin_command(
    message: Message
):

    create_user(message)

    if not is_admin(
        message.from_user.id
    ):

        await message.answer(
            "⛔ Доступ запрещён."
        )

        return

    await message.answer(
        "🛡️ <b>Админ-панель Zolog AI</b>",
        reply_markup=admin_menu(),
        parse_mode="HTML"
    )


# ============================================================
# ADMIN DASHBOARD
# ============================================================

@dp.callback_query(
    F.data == "admin_dashboard"
)
async def admin_dashboard(
    callback: CallbackQuery
):

    if not await admin_required(
        callback
    ):
        return

    users = db.execute(
        "SELECT COUNT(*) AS c FROM users"
    ).fetchone()["c"]

    materials = db.execute(
        "SELECT COUNT(*) AS c FROM materials"
    ).fetchone()["c"]

    books = db.execute(
        "SELECT COUNT(*) AS c FROM books"
    ).fetchone()["c"]

    referrals = db.execute(
        "SELECT COUNT(*) AS c FROM referrals"
    ).fetchone()["c"]

    payments = db.execute(
        "SELECT COUNT(*) AS c FROM payments"
    ).fetchone()["c"]

    stars = db.execute(
        "SELECT COALESCE(SUM(stars), 0) AS s FROM payments"
    ).fetchone()["s"]

    text = (
        "📊 <b>Dashboard</b>\n\n"
        f"👥 Пользователей: <b>{users}</b>\n"
        f"📝 Материалов: <b>{materials}</b>\n"
        f"📚 Книг: <b>{books}</b>\n"
        f"🎁 Рефералов: <b>{referrals}</b>\n"
        f"💳 Платежей: <b>{payments}</b>\n"
        f"⭐ Получено Stars: <b>{stars}</b>"
    )

    await callback.message.edit_text(
        text,
        reply_markup=admin_menu(),
        parse_mode="HTML"
    )


# ============================================================
# ADMIN USERS
# ============================================================

@dp.callback_query(
    F.data == "admin_users"
)
async def admin_users(
    callback: CallbackQuery
):

    if not await admin_required(
        callback
    ):
        return

    total = db.execute(
        "SELECT COUNT(*) AS c FROM users"
    ).fetchone()["c"]

    active = db.execute(
        """
        SELECT COUNT(*) AS c
        FROM users
        WHERE is_blocked = 0
        """
    ).fetchone()["c"]

    blocked = db.execute(
        """
        SELECT COUNT(*) AS c
        FROM users
        WHERE is_blocked = 1
        """
    ).fetchone()["c"]

    text = (
        "👥 <b>Пользователи</b>\n\n"
        f"Всего: <b>{total}</b>\n"
        f"Активных: <b>{active}</b>\n"
        f"Заблокированных: <b>{blocked}</b>"
    )

    await callback.message.edit_text(
        text,
        reply_markup=admin_menu(),
        parse_mode="HTML"
    )


# ============================================================
# ADMIN MATERIALS
# ============================================================

@dp.callback_query(
    F.data == "admin_materials"
)
async def admin_materials(
    callback: CallbackQuery
):

    if not await admin_required(
        callback
    ):
        return

    total = db.execute(
        "SELECT COUNT(*) AS c FROM materials"
    ).fetchone()["c"]

    text = (
        "🗄️ <b>Материалы</b>\n\n"
        f"Всего создано: <b>{total}</b>"
    )

    await callback.message.edit_text(
        text,
        reply_markup=admin_menu(),
        parse_mode="HTML"
    )


# ============================================================
# ADMIN BOOKS
# ============================================================

@dp.callback_query(
    F.data == "admin_books"
)
async def admin_books(
    callback: CallbackQuery
):

    if not await admin_required(
        callback
    ):
        return

    total = db.execute(
        "SELECT COUNT(*) AS c FROM books"
    ).fetchone()["c"]

    text = (
        "📚 <b>Книги</b>\n\n"
        f"Загружено книг: <b>{total}</b>"
    )

    await callback.message.edit_text(
        text,
        reply_markup=admin_menu(),
        parse_mode="HTML"
    )


# ============================================================
# ADMIN AI
# ============================================================

@dp.callback_query(
    F.data == "admin_ai"
)
async def admin_ai(
    callback: CallbackQuery
):

    if not await admin_required(
        callback
    ):
        return

    text = (
        "🤖 <b>AI</b>\n\n"
        f"Основная модель:\n"
        f"<code>{html.escape(GEMINI_MODEL)}</code>\n\n"
        f"Резервная модель:\n"
        f"<code>{html.escape(GEMINI_FALLBACK_MODEL)}</code>\n\n"
        "🔄 Повторы при временных ошибках: "
        "<b>включены</b>\n\n"
        "🔁 Переключение при 429/лимите: "
        "<b>включено</b>\n\n"
        "⏱ Задержки:\n"
        "3 → 6 → 12 → 20 секунд"
    )

    await callback.message.edit_text(
        text,
        reply_markup=admin_menu(),
        parse_mode="HTML"
    )


# ============================================================
# ADMIN REFERRALS
# ============================================================

@dp.callback_query(
    F.data == "admin_referrals"
)
async def admin_referrals(
    callback: CallbackQuery
):

    if not await admin_required(
        callback
    ):
        return

    total = db.execute(
        "SELECT COUNT(*) AS c FROM referrals"
    ).fetchone()["c"]

    text = (
        "🎁 <b>Реферальная система</b>\n\n"
        f"Всего приглашений: <b>{total}</b>\n"
        f"Награда: <b>+{REFERRAL_REWARD}</b> генерация"
    )

    await callback.message.edit_text(
        text,
        reply_markup=admin_menu(),
        parse_mode="HTML"
    )


# ============================================================
# ADMIN PAYMENTS
# ============================================================

@dp.callback_query(
    F.data == "admin_payments"
)
async def admin_payments(
    callback: CallbackQuery
):

    if not await admin_required(
        callback
    ):
        return

    payments = db.execute(
        "SELECT COUNT(*) AS c FROM payments"
    ).fetchone()["c"]

    stars = db.execute(
        """
        SELECT COALESCE(SUM(stars), 0) AS s
        FROM payments
        """
    ).fetchone()["s"]

    generations = db.execute(
        """
        SELECT COALESCE(SUM(generations), 0) AS g
        FROM payments
        """
    ).fetchone()["g"]

    text = (
        "💳 <b>Оплаты</b>\n\n"
        f"Платежей: <b>{payments}</b>\n"
        f"⭐ Stars: <b>{stars}</b>\n"
        f"➕ Начислено генераций: "
        f"<b>{generations}</b>"
    )

    await callback.message.edit_text(
        text,
        reply_markup=admin_menu(),
        parse_mode="HTML"
    )


# ============================================================
# ADMIN PAYMENT STATS
# ============================================================

@dp.callback_query(
    F.data == "admin_payment_stats"
)
async def admin_payment_stats(
    callback: CallbackQuery
):

    if not await admin_required(
        callback
    ):
        return

    total_stars = db.execute(
        """
        SELECT COALESCE(SUM(stars), 0) AS s
        FROM payments
        """
    ).fetchone()["s"]

    total_payments = db.execute(
        """
        SELECT COUNT(*) AS c
        FROM payments
        """
    ).fetchone()["c"]

    paying_users = db.execute(
        """
        SELECT COUNT(DISTINCT telegram_id) AS c
        FROM payments
        """
    ).fetchone()["c"]

    average = (
        round(
            total_stars / total_payments,
            2
        )
        if total_payments
        else 0
    )

    today = (
        datetime.now()
        .replace(
            hour=0,
            minute=0,
            second=0,
            microsecond=0
        )
        .strftime("%Y-%m-%d %H:%M:%S")
    )

    week = (
        datetime.now() - timedelta(days=7)
    ).strftime(
        "%Y-%m-%d %H:%M:%S"
    )

    month = (
        datetime.now() - timedelta(days=30)
    ).strftime(
        "%Y-%m-%d %H:%M:%S"
    )

    today_stars = db.execute(
        """
        SELECT COALESCE(SUM(stars), 0) AS s
        FROM payments
        WHERE created_at >= ?
        """,
        (today,)
    ).fetchone()["s"]

    week_stars = db.execute(
        """
        SELECT COALESCE(SUM(stars), 0) AS s
        FROM payments
        WHERE created_at >= ?
        """,
        (week,)
    ).fetchone()["s"]

    month_stars = db.execute(
        """
        SELECT COALESCE(SUM(stars), 0) AS s
        FROM payments
        WHERE created_at >= ?
        """,
        (month,)
    ).fetchone()["s"]

    popular = db.execute(
        """
        SELECT package, COUNT(*) AS c
        FROM payments
        GROUP BY package
        ORDER BY c DESC
        LIMIT 1
        """
    ).fetchone()

    popular_package = (
        popular["package"]
        if popular
        else "нет данных"
    )

    text = (
        "📈 <b>Статистика оплат</b>\n\n"
        f"⭐ Всего Stars: <b>{total_stars}</b>\n"
        f"💳 Всего платежей: <b>{total_payments}</b>\n"
        f"👥 Платящих пользователей: "
        f"<b>{paying_users}</b>\n"
        f"📊 Средний платёж: "
        f"<b>{average} ⭐</b>\n\n"
        f"Сегодня: <b>{today_stars} ⭐</b>\n"
        f"7 дней: <b>{week_stars} ⭐</b>\n"
        f"30 дней: <b>{month_stars} ⭐</b>\n\n"
        f"🏆 Популярный пакет: "
        f"<b>{popular_package}</b>"
    )

    await callback.message.edit_text(
        text,
        reply_markup=admin_menu(),
        parse_mode="HTML"
    )


# ============================================================
# ADMIN BROADCAST
# ============================================================

@dp.callback_query(
    F.data == "admin_broadcast"
)
async def admin_broadcast(
    callback: CallbackQuery
):

    if not await admin_required(
        callback
    ):
        return

    await callback.message.edit_text(
        "📢 <b>Рассылки</b>\n\n"
        "Система массовых рассылок подготовлена "
        "для дальнейшего расширения.",
        reply_markup=admin_menu(),
        parse_mode="HTML"
    )


# ============================================================
# ADMIN MODERATION
# ============================================================

@dp.callback_query(
    F.data == "admin_moderation"
)
async def admin_moderation(
    callback: CallbackQuery
):

    if not await admin_required(
        callback
    ):
        return

    await callback.message.edit_text(
        "🛡️ <b>Модерация</b>\n\n"
        "Раздел модерации подготовлен "
        "для дальнейшего расширения.",
        reply_markup=admin_menu(),
        parse_mode="HTML"
    )


# ============================================================
# ADMIN LOGS
# ============================================================

@dp.callback_query(
    F.data == "admin_logs"
)
async def admin_logs(
    callback: CallbackQuery
):

    if not await admin_required(
        callback
    ):
        return

    logs = db.execute(
        """
        SELECT *
        FROM logs
        ORDER BY id DESC
        LIMIT 10
        """
    ).fetchall()

    text = "📝 <b>Последние логи</b>\n\n"

    if not logs:

        text += "Логов пока нет."

    else:

        for item in logs:

            text += (
                f"• {html.escape(item['action'])} "
                f"| {item['telegram_id']}\n"
            )

    await callback.message.edit_text(
        text,
        reply_markup=admin_menu(),
        parse_mode="HTML"
    )


# ============================================================
# ADMIN TEST MODE
# ============================================================

@dp.callback_query(
    F.data == "admin_test"
)
async def admin_test(
    callback: CallbackQuery
):

    if not await admin_required(
        callback
    ):
        return

    await callback.message.edit_text(
        "🧪 <b>Test Mode</b>\n\n"
        "Режим тестирования подготовлен "
        "для дальнейшего расширения.",
        reply_markup=admin_menu(),
        parse_mode="HTML"
    )


# ============================================================
# ADMIN SETTINGS
# ============================================================

@dp.callback_query(
    F.data == "admin_settings"
)
async def admin_settings(
    callback: CallbackQuery
):

    if not await admin_required(
        callback
    ):
        return

    text = (
        "⚙️ <b>Настройки Zolog AI</b>\n\n"
        f"Основная модель:\n"
        f"<code>{html.escape(GEMINI_MODEL)}</code>\n\n"
        f"Резервная модель:\n"
        f"<code>{html.escape(GEMINI_FALLBACK_MODEL)}</code>\n\n"
        f"Стартовые генерации: "
        f"<b>{START_GENERATIONS}</b>\n"
        f"Реферальная награда: "
        f"<b>{REFERRAL_REWARD}</b>"
    )

    await callback.message.edit_text(
        text,
        reply_markup=admin_menu(),
        parse_mode="HTML"
    )


# ============================================================
# BACK MAIN
# ============================================================

@dp.callback_query(
    F.data == "back_main"
)
async def back_main(
    callback: CallbackQuery
):

    await callback.answer()

    clear_state(
        callback.from_user.id
    )

    user = get_user(
        callback.from_user.id
    )

    if user:

        await callback.message.edit_text(
            "🏠 <b>Главное меню Zolog AI</b>\n\n"
            f"⭐ Ваш баланс: "
            f"<b>{user['generations']}</b> генераций",
            reply_markup=main_menu(),
            parse_mode="HTML"
        )

    else:

        await callback.message.edit_text(
            "🏠 Главное меню",
            reply_markup=main_menu()
        )


# ============================================================
# HEALTH SERVER FOR RENDER
# ============================================================

async def health(request):
    return web.Response(
        text="Zolog AI is running"
    )


async def root(request):
    return web.Response(
        text="Zolog AI"
    )


async def start_web_server():

    app = web.Application()

    app.router.add_get(
        "/",
        root
    )

    app.router.add_get(
        "/health",
        health
    )

    runner = web.AppRunner(app)

    await runner.setup()

    site = web.TCPSite(
        runner,
        "0.0.0.0",
        PORT
    )

    await site.start()

    logger.info(
        f"🌐 Web server started on port {PORT}"
    )


# ============================================================
# MAIN
# ============================================================

async def main():

    logger.info(
        "======================================"
    )

    logger.info(
        "🚀 Zolog AI запускается..."
    )

    logger.info(
        f"🤖 Main model: {GEMINI_MODEL}"
    )

    logger.info(
        f"🔄 Fallback model: {GEMINI_FALLBACK_MODEL}"
    )

    logger.info(
        "🔁 Retry system: ENABLED"
    )

    logger.info(
        "======================================"
    )

    # Если раньше был webhook,
    # удаляем его перед polling.
    await bot.delete_webhook(
        drop_pending_updates=True
    )

    await start_web_server()

    try:

        await dp.start_polling(
            bot
        )

    finally:

        await bot.session.close()


if __name__ == "__main__":

    try:

        asyncio.run(
            main()
        )

    except KeyboardInterrupt:

        logger.info(
            "Zolog AI остановлен."
        )
