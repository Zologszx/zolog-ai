import os
import asyncio
import logging
import sqlite3

from aiohttp import web
from aiogram import Bot, Dispatcher, types
from aiogram.filters import Command
from google import genai


# =========================
# НАСТРОЙКИ
# =========================

BOT_TOKEN = os.getenv("BOT_TOKEN")
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
ADMIN_ID = os.getenv("ADMIN_ID")

GEMINI_MODEL = os.getenv(
    "GEMINI_MODEL",
    "gemini-3.6-flash"
)

if not BOT_TOKEN:
    raise RuntimeError("BOT_TOKEN не найден")

if not GEMINI_API_KEY:
    raise RuntimeError("GEMINI_API_KEY не найден")


# =========================
# ЛОГИ
# =========================

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s"
)

logger = logging.getLogger("zolog-ai")


# =========================
# TELEGRAM + GEMINI
# =========================

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()

gemini = genai.Client(
    api_key=GEMINI_API_KEY
)


# =========================
# БАЗА ДАННЫХ
# =========================

DB_NAME = "zolog_ai.db"


def init_db():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            telegram_id INTEGER UNIQUE NOT NULL,
            username TEXT,
            first_name TEXT,
            requests_count INTEGER DEFAULT 0,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS materials (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            telegram_id INTEGER NOT NULL,
            material_type TEXT,
            title TEXT,
            content TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS payments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            telegram_id INTEGER NOT NULL,
            stars INTEGER DEFAULT 0,
            status TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)

    conn.commit()
    conn.close()


def add_user(user: types.User):
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()

    cursor.execute("""
        INSERT OR IGNORE INTO users
        (telegram_id, username, first_name)
        VALUES (?, ?, ?)
    """, (
        user.id,
        user.username,
        user.first_name
    ))

    cursor.execute("""
        UPDATE users
        SET username = ?, first_name = ?
        WHERE telegram_id = ?
    """, (
        user.username,
        user.first_name,
        user.id
    ))

    conn.commit()
    conn.close()


def save_material(
    telegram_id: int,
    material_type: str,
    title: str,
    content: str
):
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()

    cursor.execute("""
        INSERT INTO materials
        (telegram_id, material_type, title, content)
        VALUES (?, ?, ?, ?)
    """, (
        telegram_id,
        material_type,
        title,
        content
    ))

    cursor.execute("""
        UPDATE users
        SET requests_count = requests_count + 1
        WHERE telegram_id = ?
    """, (telegram_id,))

    conn.commit()
    conn.close()


# =========================
# GEMINI
# =========================

async def ask_ai(prompt: str) -> str:
    try:
        response = await asyncio.to_thread(
            gemini.models.generate_content,
            model=GEMINI_MODEL,
            contents=prompt
        )

        if not response or not response.text:
            return "❌ Gemini не вернул текстовый ответ."

        return response.text

    except Exception as error:
        logger.exception("Ошибка Gemini: %s", error)

        return (
            "❌ Произошла ошибка при обращении к AI.\n\n"
            f"Тип ошибки: {type(error).__name__}\n"
            f"Описание: {error}"
        )


# =========================
# /START
# =========================

@dp.message(Command("start"))
async def start_handler(message: types.Message):
    add_user(message.from_user)

    await message.answer(
        "🤖 <b>Zolog AI</b>\n\n"
        "Привет! Я твой AI-помощник.\n\n"
        "Я умею:\n"
        "📚 работать с учебными материалами\n"
        "📝 создавать тексты\n"
        "🎓 помогать с рефератами и курсовыми\n"
        "🧠 отвечать на вопросы\n"
        "📊 анализировать информацию\n\n"
        "Просто напиши мне свой запрос.",
        parse_mode="HTML"
    )


# =========================
# /HELP
# =========================

@dp.message(Command("help"))
async def help_handler(message: types.Message):
    await message.answer(
        "📖 <b>Помощь Zolog AI</b>\n\n"
        "Примеры запросов:\n\n"
        "• Объясни, что такое ХОЗЛ\n"
        "• Создай план реферата по физической терапии\n"
        "• Составь конспект по теме\n"
        "• Объясни тему простыми словами\n\n"
        "В будущем здесь появятся:\n"
        "📚 библиотека файлов\n"
        "📄 DOCX/PDF\n"
        "📊 презентации\n"
        "🎥 анализ видео\n"
        "💳 Telegram Stars\n"
        "👑 полноценная админ-панель"
    )


# =========================
# /ADMIN
# =========================

def is_admin(user_id: int) -> bool:
    if not ADMIN_ID:
        return False

    return str(user_id) == str(ADMIN_ID)


@dp.message(Command("admin"))
async def admin_handler(message: types.Message):
    if not is_admin(message.from_user.id):
        await message.answer("⛔ У вас нет доступа.")
        return

    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()

    cursor.execute("SELECT COUNT(*) FROM users")
    users_count = cursor.fetchone()[0]

    cursor.execute("SELECT COUNT(*) FROM materials")
    materials_count = cursor.fetchone()[0]

    cursor.execute("SELECT COUNT(*) FROM payments")
    payments_count = cursor.fetchone()[0]

    conn.close()

    await message.answer(
        "👑 <b>Zolog AI — Admin Panel</b>\n\n"
        f"👥 Пользователей: <b>{users_count}</b>\n"
        f"📚 Материалов: <b>{materials_count}</b>\n"
        f"💳 Платежей: <b>{payments_count}</b>\n\n"
        "Команды:\n"
        "/stats — статистика\n"
        "/users — пользователи\n"
        "/materials — материалы",
        parse_mode="HTML"
    )


# =========================
# /STATS
# =========================

@dp.message(Command("stats"))
async def stats_handler(message: types.Message):
    if not is_admin(message.from_user.id):
        await message.answer("⛔ У вас нет доступа.")
        return

    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()

    cursor.execute("SELECT COUNT(*) FROM users")
    users = cursor.fetchone()[0]

    cursor.execute("SELECT COUNT(*) FROM materials")
    materials = cursor.fetchone()[0]

    cursor.execute(
        "SELECT COALESCE(SUM(requests_count), 0) FROM users"
    )
    requests = cursor.fetchone()[0]

    conn.close()

    await message.answer(
        "📊 <b>Статистика Zolog AI</b>\n\n"
        f"👥 Пользователей: {users}\n"
        f"📝 AI-запросов: {requests}\n"
        f"📚 Материалов: {materials}",
        parse_mode="HTML"
    )


# =========================
# /USERS
# =========================

@dp.message(Command("users"))
async def users_handler(message: types.Message):
    if not is_admin(message.from_user.id):
        await message.answer("⛔ У вас нет доступа.")
        return

    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()

    cursor.execute("""
        SELECT telegram_id, username, first_name, requests_count
        FROM users
        ORDER BY id DESC
        LIMIT 20
    """)

    users = cursor.fetchall()
    conn.close()

    if not users:
        await message.answer("👥 Пользователей пока нет.")
        return

    text = "👥 <b>Последние пользователи</b>\n\n"

    for telegram_id, username, first_name, requests in users:
        name = first_name or "Без имени"
        username_text = f"@{username}" if username else "без username"

        text += (
            f"👤 {name}\n"
            f"   {username_text}\n"
            f"   ID: <code>{telegram_id}</code>\n"
            f"   Запросов: {requests}\n\n"
        )

    await message.answer(
        text,
        parse_mode="HTML"
    )


# =========================
# /MATERIALS
# =========================

@dp.message(Command("materials"))
async def materials_handler(message: types.Message):
    if not is_admin(message.from_user.id):
        await message.answer("⛔ У вас нет доступа.")
        return

    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()

    cursor.execute("""
        SELECT id, telegram_id, material_type, title
        FROM materials
        ORDER BY id DESC
        LIMIT 20
    """)

    materials = cursor.fetchall()
    conn.close()

    if not materials:
        await message.answer("📚 Материалов пока нет.")
        return

    text = "📚 <b>Последние материалы</b>\n\n"

    for material_id, telegram_id, material_type, title in materials:
        text += (
            f"#{material_id} — {material_type}\n"
            f"👤 ID: <code>{telegram_id}</code>\n"
            f"📄 {title}\n\n"
        )

    await message.answer(
        text,
        parse_mode="HTML"
    )


# =========================
# ОБЫЧНЫЕ СООБЩЕНИЯ
# =========================

@dp.message()
async def message_handler(message: types.Message):
    add_user(message.from_user)

    prompt = (
        "Ты — Zolog AI, интеллектуальный помощник.\n"
        "Отвечай понятно, структурированно и по существу.\n"
        "Если вопрос учебный — объясняй материал так, "
        "чтобы студент мог его понять и использовать в учебе.\n\n"
        f"Запрос пользователя:\n{message.text}"
    )

    answer = await ask_ai(prompt)

    save_material(
        telegram_id=message.from_user.id,
        material_type="ai_answer",
        title=message.text[:100],
        content=answer
    )

    await message.answer(answer)


# =========================
# HEALTH CHECK ДЛЯ RENDER
# =========================

async def health(request):
    return web.Response(
        text="Zolog AI is running"
    )


async def start_web_server():
    app = web.Application()

    app.router.add_get(
        "/",
        health
    )

    app.router.add_get(
        "/health",
        health
    )

    runner = web.AppRunner(app)
    await runner.setup()

    port = int(
        os.getenv("PORT", "10000")
    )

    site = web.TCPSite(
        runner,
        "0.0.0.0",
        port
    )

    await site.start()

    logger.info(
        "Health server started on port %s",
        port
    )


# =========================
# ЗАПУСК
# =========================

async def main():
    logger.info("Запуск Zolog AI...")

    init_db()

    await start_web_server()

    await bot.delete_webhook(
        drop_pending_updates=True
    )

    logger.info(
        "Bot started. Model: %s",
        GEMINI_MODEL
    )

    await dp.start_polling(bot)


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        logger.info("Bot stopped")
