import asyncio
import logging
import os
import sqlite3

from aiogram import Bot, Dispatcher, F
from aiogram.filters import Command, CommandStart
from aiogram.types import Message, InlineKeyboardMarkup, InlineKeyboardButton, CallbackQuery
from google import genai

logging.basicConfig(level=logging.INFO)

BOT_TOKEN = os.getenv("BOT_TOKEN")
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
GEMINI_MODEL = os.getenv("GEMINI_MODEL", "gemini-2.5-flash")

if not BOT_TOKEN:
    raise RuntimeError("Не задан BOT_TOKEN")
if not GEMINI_API_KEY:
    raise RuntimeError("Не задан GEMINI_API_KEY")

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()
client = genai.Client(api_key=GEMINI_API_KEY)

DB_PATH = "stories.db"
user_states = {}


def init_db():
    with sqlite3.connect(DB_PATH) as db:
        db.execute("""
            CREATE TABLE IF NOT EXISTS stories (
                user_id INTEGER PRIMARY KEY,
                universe TEXT,
                character TEXT,
                story TEXT
            )
        """)
        db.commit()


def get_story(user_id: int):
    with sqlite3.connect(DB_PATH) as db:
        row = db.execute(
            "SELECT universe, character, story FROM stories WHERE user_id = ?",
            (user_id,)
        ).fetchone()
    return row


def save_story(user_id: int, universe: str, character: str, story: str):
    with sqlite3.connect(DB_PATH) as db:
        db.execute("""
            INSERT INTO stories(user_id, universe, character, story)
            VALUES (?, ?, ?, ?)
            ON CONFLICT(user_id) DO UPDATE SET
                universe=excluded.universe,
                character=excluded.character,
                story=excluded.story
        """, (user_id, universe, character, story))
        db.commit()


def menu():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="🌎 Создать историю", callback_data="new_story")],
        [InlineKeyboardButton(text="📖 Продолжить историю", callback_data="continue_story")],
        [InlineKeyboardButton(text="👤 Мой персонаж", callback_data="character")],
        [InlineKeyboardButton(text="🗑️ Новая история", callback_data="reset")]
    ])


def universe_menu():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="🍥 Naruto", callback_data="u_naruto")],
        [InlineKeyboardButton(text="⚡ Harry Potter", callback_data="u_hp")],
        [InlineKeyboardButton(text="🦸 Marvel", callback_data="u_marvel")],
        [InlineKeyboardButton(text="🎭 Своя вселенная", callback_data="u_custom")]
    ])


def character_menu():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="🍥 Naruto", callback_data="c_naruto")],
        [InlineKeyboardButton(text="⚡ Harry Potter", callback_data="c_harry")],
        [InlineKeyboardButton(text="👁️ Kakashi", callback_data="c_kakashi")],
        [InlineKeyboardButton(text="✏️ Свой персонаж", callback_data="c_custom")]
    ])


@dp.message(CommandStart())
async def start(message: Message):
    await message.answer(
        "🎭 <b>Zolog Story AI</b>\n\n"
        "Интерактивные фанфики, где ты управляешь своим персонажем.\n\n"
        "Бот управляет миром и NPC, а ты сам решаешь, "
        "что говорит и делает твой герой.",
        reply_markup=menu(),
        parse_mode="HTML"
    )


@dp.callback_query(F.data == "new_story")
async def new_story(callback: CallbackQuery):
    user_states[callback.from_user.id] = {"step": "universe"}
    await callback.message.answer("🌎 Выбери вселенную:", reply_markup=universe_menu())
    await callback.answer()


@dp.callback_query(F.data.startswith("u_"))
async def choose_universe(callback: CallbackQuery):
    uid = callback.from_user.id
    state = user_states.setdefault(uid, {})
    values = {
        "u_naruto": "Naruto",
        "u_hp": "Harry Potter",
        "u_marvel": "Marvel",
        "u_custom": "Своя вселенная"
    }
    state["universe"] = values[callback.data]
    state["step"] = "character"

    await callback.message.answer(
        f"🌎 Вселенная: <b>{state['universe']}</b>\n\n"
        "👤 За кого ты будешь играть?",
        reply_markup=character_menu(),
        parse_mode="HTML"
    )
    await callback.answer()


@dp.callback_query(F.data.startswith("c_"))
async def choose_character(callback: CallbackQuery):
    uid = callback.from_user.id
    state = user_states.setdefault(uid, {})
    values = {
        "c_naruto": "Naruto",
        "c_harry": "Harry Potter",
        "c_kakashi": "Kakashi",
        "c_custom": None
    }

    character = values[callback.data]

    if character is None:
        state["step"] = "custom_character"
        await callback.message.answer("✏️ Напиши имя своего персонажа.")
        await callback.answer()
        return

    state["character"] = character
    await begin_story(callback.message, uid)
    await callback.answer()


async def begin_story(message: Message, uid: int):
    state = user_states[uid]
    universe = state["universe"]
    character = state["character"]

    prompt = f"""
Ты ведущий интерактивного фанфика.

Вселенная: {universe}
Игрок управляет персонажем: {character}

Правила:
1. Игрок полностью управляет своим персонажем.
2. Никогда не придумывай за игрока его реплики, мысли, решения или действия.
3. Ты управляешь всеми NPC, окружающим миром и событиями.
4. После каждого события остановись и дай игроку возможность ответить.
5. Сохраняй характеры канонических персонажей.
6. Не пересказывай весь сюжет. Продвигай сцену постепенно.
7. Если игрок делает неканоничный поступок, принимай его как часть альтернативной истории.
8. Пиши живо, как интерактивную RPG.
9. Не говори пользователю, что ты ИИ или ведущий.

Начни новую историю с интересной сцены.
Не делай действие за персонажа игрока.
Закончись на моменте, когда игрок должен сделать выбор или ответить.
"""

    try:
        response = await asyncio.to_thread(
            client.models.generate_content,
            model=GEMINI_MODEL,
            contents=prompt
        )
        text = response.text or "Не удалось начать историю."
    except Exception:
        logging.exception("Gemini error")
        await message.answer("❌ Ошибка AI. Проверь GEMINI_API_KEY и GEMINI_MODEL.")
        return

    save_story(uid, universe, character, text)
    await message.answer(
        f"📖 <b>{universe}</b>\n👤 Ты играешь за: <b>{character}</b>\n\n{text}",
        parse_mode="HTML"
    )


@dp.message()
async def handle_message(message: Message):
    uid = message.from_user.id
    state = user_states.get(uid, {})

    if state.get("step") == "custom_character":
        state["character"] = message.text.strip()
        state["step"] = "story"
        await begin_story(message, uid)
        return

    story = get_story(uid)

    if not story:
        await message.answer(
            "Сначала создай историю через меню.",
            reply_markup=menu()
        )
        return

    universe, character, previous_story = story

    prompt = f"""
Ты продолжаешь интерактивный фанфик.

Вселенная: {universe}
Персонаж игрока: {character}

Последняя часть истории:
{previous_story[-12000:]}

Новое сообщение игрока:
{message.text}

СТРОГИЕ ПРАВИЛА:
- Не управляй персонажем игрока.
- Не пиши за него действия.
- Не пиши за него мысли.
- Не придумывай его реплики.
- Не отменяй его решение.
- Управляй только NPC, миром и последствиями.
- Сохраняй непрерывность событий.
- Продолжи сцену и снова остановись, чтобы игрок мог ответить.
"""

    try:
        response = await asyncio.to_thread(
            client.models.generate_content,
            model=GEMINI_MODEL,
            contents=prompt
        )
        text = response.text or "Продолжение не удалось сгенерировать."
    except Exception:
        logging.exception("Gemini error")
        await message.answer("❌ Ошибка AI при продолжении истории.")
        return

    new_story = previous_story + "\n\nИгрок: " + message.text + "\n\nБот: " + text

    # Ограничиваем размер памяти, чтобы запросы не становились бесконечными.
    new_story = new_story[-30000:]

    save_story(uid, universe, character, new_story)
    await message.answer(text)


@dp.callback_query(F.data == "continue_story")
async def continue_story(callback: CallbackQuery):
    story = get_story(callback.from_user.id)

    if not story:
        await callback.message.answer("📭 У тебя пока нет сохранённой истории.")
    else:
        universe, character, text = story
        await callback.message.answer(
            f"📖 <b>{universe}</b>\n👤 {character}\n\n"
            "Последняя сцена:\n" + text[-5000:],
            parse_mode="HTML"
        )

    await callback.answer()


@dp.callback_query(F.data == "character")
async def show_character(callback: CallbackQuery):
    story = get_story(callback.from_user.id)

    if not story:
        await callback.message.answer("Персонаж ещё не создан.")
    else:
        universe, character, _ = story
        await callback.message.answer(
            f"👤 <b>Персонаж</b>\n\n"
            f"🌎 Вселенная: {universe}\n"
            f"🎭 Персонаж: {character}",
            parse_mode="HTML"
        )

    await callback.answer()


@dp.callback_query(F.data == "reset")
async def reset_story(callback: CallbackQuery):
    with sqlite3.connect(DB_PATH) as db:
        db.execute("DELETE FROM stories WHERE user_id = ?", (callback.from_user.id,))
        db.commit()

    user_states.pop(callback.from_user.id, None)

    await callback.message.answer(
        "🗑️ История удалена.\n\nСоздадим новую?",
        reply_markup=universe_menu()
    )
    await callback.answer()


async def main():
    init_db()
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
