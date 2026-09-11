import asyncio
import os
import sqlite3
from datetime import datetime
from pathlib import Path

from aiohttp import web
from aiogram import Bot, Dispatcher, F, Router
from aiogram.filters import Command, CommandStart
from aiogram.types import Message, FSInputFile

from google import genai
from google.genai import types

from pypdf import PdfReader
from docx import Document


# =========================================================
# CONFIG
# =========================================================

BOT_TOKEN = os.getenv("BOT_TOKEN", "")
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "")
ADMIN_ID = int(os.getenv("ADMIN_ID", "0"))

GEMINI_MODEL = os.getenv("GEMINI_MODEL", "gemini-2.5-flash")

MAX_FILE_SIZE = 20 * 1024 * 1024

BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "data"
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"

DATA_DIR.mkdir(exist_ok=True)
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

DB_PATH = DATA_DIR / "zolog.db"


# =========================================================
# CHECK CONFIG
# =========================================================

if not BOT_TOKEN:
    raise RuntimeError("BOT_TOKEN is not configured")

if not GEMINI_API_KEY:
    raise RuntimeError("GEMINI_API_KEY is not configured")


# =========================================================
# CLIENTS
# =========================================================

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()
router = Router()

gemini = genai.Client(api_key=GEMINI_API_KEY)


# =========================================================
# DATABASE
# =========================================================

def db():
    return sqlite3.connect(DB_PATH)


def init_db():
    connection = db()
    cursor = connection.cursor()

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY,
            username TEXT,
            first_name TEXT,
            created_at TEXT,
            requests INTEGER DEFAULT 0
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS files (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER,
            filename TEXT,
            path TEXT,
            created_at TEXT
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS materials (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER,
            type TEXT,
            prompt TEXT,
            content TEXT,
            created_at TEXT
        )
    """)

    connection.commit()
    connection.close()


def save_user(message: Message):
    user = message.from_user

    connection = db()
    cursor = connection.cursor()

    cursor.execute("""
        INSERT INTO users (
            id,
            username,
            first_name,
            created_at
        )
        VALUES (?, ?, ?, ?)
        ON CONFLICT(id) DO UPDATE SET
            username=excluded.username,
            first_name=excluded.first_name
    """, (
        user.id,
        user.username,
        user.first_name,
        datetime.now().isoformat()
    ))

    connection.commit()
    connection.close()


def save_material(user_id, material_type, prompt, content):
    connection = db()
    cursor = connection.cursor()

    cursor.execute("""
        INSERT INTO materials (
            user_id,
            type,
            prompt,
            content,
            created_at
        )
        VALUES (?, ?, ?, ?, ?)
    """, (
        user_id,
        material_type,
        prompt,
        content,
        datetime.now().isoformat()
    ))

    cursor.execute("""
        UPDATE users
        SET requests = requests + 1
        WHERE id = ?
    """, (user_id,))

    connection.commit()
    connection.close()


# =========================================================
# ADMIN
# =========================================================

def is_admin(user_id: int) -> bool:
    return user_id == ADMIN_ID


# =========================================================
# GEMINI
# =========================================================

async def ask_gemini(prompt: str) -> str:

    response = await asyncio.to_thread(
        gemini.models.generate_content,
        model=GEMINI_MODEL,
        contents=prompt,
        config=types.GenerateContentConfig(
            temperature=0.7,
            max_output_tokens=8192
        )
    )

    if not response.text:
        raise RuntimeError("Gemini returned an empty response")

    return response.text


# =========================================================
# FILE TEXT EXTRACTION
# =========================================================

def extract_pdf(path: str) -> str:

    reader = PdfReader(path)

    text = []

    for page in reader.pages:
        try:
            page_text = page.extract_text() or ""
            text.append(page_text)
        except Exception:
            pass

    return "\n".join(text)


def extract_docx(path: str) -> str:

    document = Document(path)

    paragraphs = []

    for paragraph in document.paragraphs:
        if paragraph.text.strip():
            paragraphs.append(paragraph.text)

    return "\n".join(paragraphs)


def extract_text_file(path: str) -> str:

    return Path(path).read_text(
        encoding="utf-8",
        errors="ignore"
    )


def extract_text(path: str) -> str:

    extension = Path(path).suffix.lower()

    if extension == ".pdf":
        return extract_pdf(path)

    if extension == ".docx":
        return extract_docx(path)

    if extension in [".txt", ".md"]:
        return extract_text_file(path)

    return ""


# =========================================================
# START
# =========================================================

@router.message(CommandStart())
async def start(message: Message):

    save_user(message)

    admin_text = ""

    if is_admin(message.from_user.id):
        admin_text = "\n\n👑 Ты вошёл как администратор.\n/admin — админ-панель"

    await message.answer(
        "🤖 <b>Zolog AI</b>\n\n"
        "Привет! Я AI-помощник для учебы.\n\n"
        "Что я умею:\n"
        "📝 Рефераты\n"
        "📚 Курсовые\n"
        "📋 Конспекты\n"
        "❓ Ответы на вопросы\n"
        "📖 Работа с загруженными книгами и файлами\n"
        "🧠 Обычный AI-помощник\n\n"
        "Просто напиши мне свой запрос."
        + admin_text,
        parse_mode="HTML"
    )


# =========================================================
# HELP
# =========================================================

@router.message(Command("help"))
async def help_command(message: Message):

    await message.answer(
        "🆘 <b>Zolog AI — помощь</b>\n\n"
        "Напиши обычный запрос, например:\n\n"
        "• Объясни туннельный синдром простыми словами\n"
        "• Напиши реферат на тему ХОЗЛ\n"
        "• Сделай план курсовой по физической терапии\n"
        "• Составь конспект по теме\n\n"
        "Также можно отправить PDF, DOCX или TXT файл.",
        parse_mode="HTML"
    )


# =========================================================
# ADMIN PANEL
# =========================================================

@router.message(Command("admin"))
async def admin_command(message: Message):

    if not is_admin(message.from_user.id):
        await message.answer("⛔ У тебя нет доступа к админ-панели.")
        return

    connection = db()
    cursor = connection.cursor()

    cursor.execute("SELECT COUNT(*) FROM users")
    users = cursor.fetchone()[0]

    cursor.execute("SELECT COUNT(*) FROM materials")
    materials = cursor.fetchone()[0]

    cursor.execute("SELECT COUNT(*) FROM files")
    files = cursor.fetchone()[0]

    connection.close()

    await message.answer(
        "👑 <b>ZOLOG AI — ADMIN PANEL</b>\n\n"
        f"👥 Пользователи: <b>{users}</b>\n"
        f"📝 Материалы: <b>{materials}</b>\n"
        f"📚 Файлы: <b>{files}</b>\n\n"
        "Команды:\n"
        "/users — пользователи\n"
        "/materials — материалы\n"
        "/files — загруженные файлы",
        parse_mode="HTML"
    )


# =========================================================
# ADMIN USERS
# =========================================================

@router.message(Command("users"))
async def admin_users(message: Message):

    if not is_admin(message.from_user.id):
        return

    connection = db()
    cursor = connection.cursor()

    cursor.execute("""
        SELECT id, username, first_name, requests
        FROM users
        ORDER BY created_at DESC
        LIMIT 20
    """)

    rows = cursor.fetchall()

    connection.close()

    if not rows:
        await message.answer("Пользователей пока нет.")
        return

    text = "👥 <b>Последние пользователи</b>\n\n"

    for row in rows:
        user_id, username, first_name, requests = row

        text += (
            f"👤 {first_name or 'Без имени'}\n"
            f"ID: <code>{user_id}</code>\n"
            f"Username: @{username or 'нет'}\n"
            f"Запросов: {requests}\n\n"
        )

    await message.answer(text, parse_mode="HTML")


# =========================================================
# ADMIN MATERIALS
# =========================================================

@router.message(Command("materials"))
async def admin_materials(message: Message):

    if not is_admin(message.from_user.id):
        return

    connection = db()
    cursor = connection.cursor()

    cursor.execute("""
        SELECT id, user_id, type, created_at
        FROM materials
        ORDER BY id DESC
        LIMIT 20
    """)

    rows = cursor.fetchall()

    connection.close()

    if not rows:
        await message.answer("Материалов пока нет.")
        return

    text = "📝 <b>Последние материалы</b>\n\n"

    for material_id, user_id, material_type, created_at in rows:

        text += (
            f"#{material_id} — {material_type}\n"
            f"👤 User ID: <code>{user_id}</code>\n"
            f"🕐 {created_at}\n\n"
        )

    await message.answer(text, parse_mode="HTML")


# =========================================================
# ADMIN FILES
# =========================================================

@router.message(Command("files"))
async def admin_files(message: Message):

    if not is_admin(message.from_user.id):
        return

    connection = db()
    cursor = connection.cursor()

    cursor.execute("""
        SELECT id, user_id, filename, created_at
        FROM files
        ORDER BY id DESC
        LIMIT 20
    """)

    rows = cursor.fetchall()

    connection.close()

    if not rows:
        await message.answer("Файлов пока нет.")
        return

    text = "📚 <b>Последние файлы</b>\n\n"

    for file_id, user_id, filename, created_at in rows:

        text += (
            f"#{file_id} — {filename}\n"
            f"👤 User ID: <code>{user_id}</code>\n"
            f"🕐 {created_at}\n\n"
        )

    await message.answer(text, parse_mode="HTML")


# =========================================================
# FILE UPLOAD
# =========================================================

@router.message(F.document)
async def receive_document(message: Message):

    save_user(message)

    document = message.document

    if document.file_size and document.file_size > MAX_FILE_SIZE:
        await message.answer(
            "❌ Файл слишком большой.\n"
            "Максимальный размер: 20 МБ."
        )
        return

    filename = document.file_name or "file"

    extension = Path(filename).suffix.lower()

    allowed = [".pdf", ".docx", ".txt", ".md"]

    if extension not in allowed:
        await message.answer(
            "❌ Я пока поддерживаю только:\n\n"
            "📕 PDF\n"
            "📘 DOCX\n"
            "📄 TXT\n"
            "📝 MD"
        )
        return

    user_dir = UPLOAD_DIR / str(message.from_user.id)
    user_dir.mkdir(parents=True, exist_ok=True)

    destination = user_dir / filename

    try:

        telegram_file = await bot.get_file(document.file_id)

        await bot.download_file(
            telegram_file.file_path,
            destination
        )

        text = extract_text(str(destination))

        if not text.strip():
            await message.answer(
                "⚠️ Файл загружен, но текст извлечь не удалось."
            )
            return

        connection = db()
        cursor = connection.cursor()

        cursor.execute("""
            INSERT INTO files (
                user_id,
                filename,
                path,
                created_at
            )
            VALUES (?, ?, ?, ?)
        """, (
            message.from_user.id,
            filename,
            str(destination),
            datetime.now().isoformat()
        ))

        connection.commit()
        connection.close()

        # Сохраняем текст рядом с файлом
        text_path = destination.with_suffix(".txt")

        text_path.write_text(
            text,
            encoding="utf-8"
        )

        await message.answer(
            f"✅ Файл <b>{filename}</b> загружен.\n\n"
            f"Извлечено символов: <b>{len(text):,}</b>\n\n"
            "Теперь можешь написать, например:\n"
            "«Сделай конспект по этому файлу»\n"
            "или\n"
            "«Напиши реферат, используя этот материал».",
            parse_mode="HTML"
        )

    except Exception as error:

        print("FILE ERROR:", repr(error))

        await message.answer(
            "❌ Произошла ошибка при обработке файла."
        )


# =========================================================
# TEXT REQUEST
# =========================================================

@router.message(F.text)
async def text_request(message: Message):

    save_user(message)

    user_text = message.text.strip()

    if not user_text:
        return

    # Не обрабатываем команды
    if user_text.startswith("/"):
        return

    await message.answer("🧠 Думаю над ответом...")

    # Ищем последний загруженный файл пользователя
    connection = db()
    cursor = connection.cursor()

    cursor.execute("""
        SELECT path
        FROM files
        WHERE user_id = ?
        ORDER BY id DESC
        LIMIT 1
    """, (message.from_user.id,))

    row = cursor.fetchone()

    connection.close()

    source_text = ""

    if row:
        file_path = Path(row[0])

        text_path = file_path.with_suffix(".txt")

        if text_path.exists():
            try:
                source_text = text_path.read_text(
                    encoding="utf-8",
                    errors="ignore"
                )
            except Exception:
                source_text = ""

    if len(source_text) > 60000:
        source_text = source_text[:60000]

    prompt = f"""
Ты — Zolog AI, интеллектуальный учебный помощник.

Пользователь написал:
{user_text}

Если ниже есть материал из загруженного пользователем файла,
используй его как основной источник информации.

Материал пользователя:
--------------------
{source_text}
--------------------

Правила:
1. Отвечай на языке пользователя.
2. Не выдумывай факты, если вопрос требует конкретной информации.
3. Если пользователь просит реферат — создай структурированный реферат.
4. Если просит курсовую — создай подробный план и академический текст.
5. Если просит конспект — сделай структурированный конспект.
6. Используй загруженный материал, когда он релевантен.
7. Пиши понятно и грамотно.
"""

    try:

        answer = await ask_gemini(prompt)

        save_material(
            message.from_user.id,
            "ai_request",
            user_text,
            answer
        )

        # Telegram ограничивает длину одного сообщения
        max_length = 4000

        for i in range(0, len(answer), max_length):

            await message.answer(
                answer[i:i + max_length]
            )

        except Exception as error:

        print("================================")
        print("GEMINI ERROR")
        print("================================")
        print(type(error).__name__)
        print(str(error))
        print(repr(error))
        print("================================")

        await message.answer(
            "❌ Ошибка Gemini.\n\n"
            "Я записал подробности ошибки в логи Render."
        )


# =========================================================
# HEALTH SERVER FOR RENDER
# =========================================================

async def health(request):
    return web.Response(
        text="Zolog AI is running"
    )


async def start_web_server():

    app = web.Application()

    app.router.add_get("/", health)
    app.router.add_get("/health", health)

    port = int(os.getenv("PORT", "10000"))

    runner = web.AppRunner(app)

    await runner.setup()

    site = web.TCPSite(
        runner,
        "0.0.0.0",
        port
    )

    await site.start()

    print(f"Health server started on port {port}")


# =========================================================
# BOT START
# =========================================================

async def main():

    init_db()

    dp.include_router(router)

    await start_web_server()

    print("================================")
    print("ZOLOG AI STARTED")
    print("================================")
    print("Model:", GEMINI_MODEL)
    print("Admin ID:", ADMIN_ID)

    await bot.delete_webhook(
        drop_pending_updates=True
    )

    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
