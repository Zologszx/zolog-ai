import asyncio
import html
import logging
import os
import sqlite3

from aiohttp import web
from aiogram import Bot, Dispatcher, F
from aiogram.filters import CommandStart
from aiogram.types import Message, InlineKeyboardMarkup, InlineKeyboardButton, CallbackQuery
from google import genai


# =========================================================
# НАСТРОЙКИ
# =========================================================

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(name)s | %(message)s"
)

BOT_TOKEN = os.getenv("BOT_TOKEN")
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
GEMINI_MODEL = os.getenv("GEMINI_MODEL", "gemini-3.6-flash")

if not BOT_TOKEN:
    raise RuntimeError("Не задан BOT_TOKEN")

if not GEMINI_API_KEY:
    raise RuntimeError("Не задан GEMINI_API_KEY")


# =========================================================
# TELEGRAM + GEMINI
# =========================================================

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()

client = genai.Client(
    api_key=GEMINI_API_KEY
)


# =========================================================
# БАЗА ДАННЫХ
# =========================================================

DB_PATH = "stories.db"

# Временные состояния пользователей.
# Например:
# {"step": "universe"}
# {"step": "custom_character"}
user_states = {}


def init_db():
    with sqlite3.connect(DB_PATH) as db:
        db.execute(
            """
            CREATE TABLE IF NOT EXISTS stories (
                user_id INTEGER PRIMARY KEY,
                universe TEXT NOT NULL,
                character TEXT NOT NULL,
                story TEXT NOT NULL
            )
            """
        )

        db.commit()


def get_story(user_id: int):
    with sqlite3.connect(DB_PATH) as db:
        return db.execute(
            """
            SELECT universe, character, story
            FROM stories
            WHERE user_id = ?
            """,
            (user_id,)
        ).fetchone()


def save_story(
    user_id: int,
    universe: str,
    character: str,
    story: str
):
    with sqlite3.connect(DB_PATH) as db:
        db.execute(
            """
            INSERT INTO stories (
                user_id,
                universe,
                character,
                story
            )
            VALUES (?, ?, ?, ?)

            ON CONFLICT(user_id)
            DO UPDATE SET
                universe = excluded.universe,
                character = excluded.character,
                story = excluded.story
            """,
            (
                user_id,
                universe,
                character,
                story
            )
        )

        db.commit()


def delete_story(user_id: int):
    with sqlite3.connect(DB_PATH) as db:
        db.execute(
            """
            DELETE FROM stories
            WHERE user_id = ?
            """,
            (user_id,)
        )

        db.commit()


# =========================================================
# КЛАВИАТУРЫ
# =========================================================

def main_menu():
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(
                    text="🌎 Создать историю",
                    callback_data="new_story"
                )
            ],
            [
                InlineKeyboardButton(
                    text="📖 Продолжить историю",
                    callback_data="continue_story"
                )
            ],
            [
                InlineKeyboardButton(
                    text="👤 Мой персонаж",
                    callback_data="character"
                )
            ],
            [
                InlineKeyboardButton(
                    text="🗑️ Новая история",
                    callback_data="reset"
                )
            ]
        ]
    )


def universe_menu():
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(
                    text="🍥 Naruto",
                    callback_data="u_naruto"
                )
            ],
            [
                InlineKeyboardButton(
                    text="⚡ Harry Potter",
                    callback_data="u_hp"
                )
            ],
            [
                InlineKeyboardButton(
                    text="🦸 Marvel",
                    callback_data="u_marvel"
                )
            ],
            [
                InlineKeyboardButton(
                    text="🎭 Своя вселенная",
                    callback_data="u_custom"
                )
            ]
        ]
    )


def character_menu():
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(
                    text="🍥 Naruto",
                    callback_data="c_naruto"
                )
            ],
            [
                InlineKeyboardButton(
                    text="⚡ Harry Potter",
                    callback_data="c_harry"
                )
            ],
            [
                InlineKeyboardButton(
                    text="👁️ Kakashi",
                    callback_data="c_kakashi"
                )
            ],
            [
                InlineKeyboardButton(
                    text="✏️ Свой персонаж",
                    callback_data="c_custom"
                )
            ]
        ]
    )


# =========================================================
# GEMINI
# =========================================================

async def generate_ai(prompt: str) -> str:
    """
    Запускает Gemini в отдельном потоке,
    чтобы не блокировать Telegram-бота.
    """

    response = await asyncio.to_thread(
        client.models.generate_content,
        model=GEMINI_MODEL,
        contents=prompt
    )

    text = getattr(response, "text", None)

    if not text:
        return ""

    return text.strip()


# =========================================================
# START
# =========================================================

@dp.message(CommandStart())
async def start(message: Message):

    user_states[message.from_user.id] = {
        "step": "idle"
    }

    await message.answer(
        "🎭 <b>Zolog Story AI</b>\n\n"
        "Интерактивные фанфики и истории.\n\n"
        "Ты полностью управляешь своим персонажем.\n"
        "Бот управляет миром, NPC и событиями.\n\n"
        "Ты сам решаешь, что говорит и делает твой герой.",
        reply_markup=main_menu(),
        parse_mode="HTML"
    )


# =========================================================
# СОЗДАНИЕ ИСТОРИИ
# =========================================================

@dp.callback_query(F.data == "new_story")
async def new_story(callback: CallbackQuery):

    uid = callback.from_user.id

    user_states[uid] = {
        "step": "universe"
    }

    await callback.message.answer(
        "🌎 <b>Выбери вселенную</b>",
        reply_markup=universe_menu(),
        parse_mode="HTML"
    )

    await callback.answer()


# =========================================================
# ВЫБОР ВСЕЛЕННОЙ
# =========================================================

@dp.callback_query(F.data.startswith("u_"))
async def choose_universe(callback: CallbackQuery):

    uid = callback.from_user.id

    state = user_states.setdefault(
        uid,
        {}
    )

    universes = {
        "u_naruto": "Naruto",
        "u_hp": "Harry Potter",
        "u_marvel": "Marvel"
    }

    # Пользователь хочет свою вселенную
    if callback.data == "u_custom":

        state["step"] = "custom_universe"

        await callback.message.answer(
            "✏️ Напиши название своей вселенной.\n\n"
            "Например:\n"
            "• Своя школа магии\n"
            "• Альтернативный мир Naruto\n"
            "• Постапокалипсис\n"
            "• Любая придуманная вселенная"
        )

        await callback.answer()

        return

    universe = universes.get(callback.data)

    if not universe:

        await callback.answer(
            "Неизвестная вселенная.",
            show_alert=True
        )

        return

    state["universe"] = universe
    state["step"] = "character"

    await callback.message.answer(
        f"🌎 Вселенная: <b>{html.escape(universe)}</b>\n\n"
        "👤 <b>За кого ты будешь играть?</b>",
        reply_markup=character_menu(),
        parse_mode="HTML"
    )

    await callback.answer()


# =========================================================
# ВЫБОР ПЕРСОНАЖА
# =========================================================

@dp.callback_query(F.data.startswith("c_"))
async def choose_character(callback: CallbackQuery):

    uid = callback.from_user.id

    state = user_states.setdefault(
        uid,
        {}
    )

    characters = {
        "c_naruto": "Naruto",
        "c_harry": "Harry Potter",
        "c_kakashi": "Kakashi"
    }

    # Собственный персонаж
    if callback.data == "c_custom":

        state["step"] = "custom_character"

        await callback.message.answer(
            "✏️ Напиши имя своего персонажа."
        )

        await callback.answer()

        return

    character = characters.get(callback.data)

    if not character:

        await callback.answer(
            "Неизвестный персонаж.",
            show_alert=True
        )

        return

    state["character"] = character
    state["step"] = "playing"

    await begin_story(
        callback.message,
        uid
    )

    await callback.answer()


# =========================================================
# НАЧАЛО ИСТОРИИ
# =========================================================

async def begin_story(
    message: Message,
    uid: int
):

    state = user_states.get(uid)

    if not state:
        await message.answer(
            "Произошла ошибка состояния. Создай историю заново.",
            reply_markup=main_menu()
        )

        return

    universe = state.get("universe")
    character = state.get("character")

    if not universe or not character:

        await message.answer(
            "Не удалось определить вселенную или персонажа.",
            reply_markup=main_menu()
        )

        return

    prompt = f"""
Ты — движок интерактивного сюжетного фанфика.

ВСЕЛЕННАЯ:
{universe}

ПЕРСОНАЖ ИГРОКА:
{character}

ГЛАВНОЕ ПРАВИЛО:
Игрок полностью управляет своим персонажем.

Ты НИКОГДА не должен придумывать за персонажа игрока:

- действия;
- движения;
- мысли;
- эмоции;
- решения;
- реплики;
- ответы;
- реакции.

Ты управляешь только:

- NPC;
- каноническими персонажами;
- окружающим миром;
- событиями;
- последствиями действий игрока;
- диалогами NPC;
- описанием окружающей обстановки.

Если игрок написал действие своего персонажа,
считай его действие фактом истории и продолжай события
с учётом этого действия.

Не исправляй решения игрока.

Не заставляй персонажа игрока делать то,
чего игрок не писал.

Сохраняй характеры канонических персонажей.

Не пересказывай всю историю.

Развивай сюжет постепенно.

Каждый ответ должен быть продолжением текущей сцены.

В конце ответа оставляй возможность игроку самому решить,
что делать дальше.

Не говори пользователю, что ты искусственный интеллект.

НАЧАЛО:

Создай интересную первую сцену этой истории.

Не выполняй никаких действий за персонажа игрока.

Остановись перед первым решением или действием игрока.
"""

    try:

        text = await generate_ai(prompt)

    except Exception as error:

        logging.exception(
            "Ошибка Gemini при создании истории"
        )

        await message.answer(
            "❌ <b>Ошибка AI</b>\n\n"
            f"Модель: <code>{html.escape(GEMINI_MODEL)}</code>\n\n"
            "Проверь GEMINI_API_KEY и доступность модели.",
            parse_mode="HTML"
        )

        return

    if not text:

        text = "Не удалось начать историю. Попробуй ещё раз."

    save_story(
        uid,
        universe,
        character,
        text
    )

    await message.answer(
        f"📖 <b>{html.escape(universe)}</b>\n"
        f"👤 Ты играешь за: <b>{html.escape(character)}</b>\n\n"
        f"{html.escape(text)}",
        parse_mode="HTML"
    )


# =========================================================
# ОБРАБОТКА СООБЩЕНИЙ
# =========================================================

@dp.message()
async def handle_message(message: Message):

    uid = message.from_user.id

    if not message.text:

        await message.answer(
            "Пожалуйста, отправь текстовое сообщение."
        )

        return

    state = user_states.get(
        uid,
        {}
    )

    # -----------------------------------------------------
    # СОЗДАНИЕ СОБСТВЕННОЙ ВСЕЛЕННОЙ
    # -----------------------------------------------------

    if state.get("step") == "custom_universe":

        universe = message.text.strip()

        if not universe:

            await message.answer(
                "Напиши название своей вселенной."
            )

            return

        state["universe"] = universe
        state["step"] = "character"

        await message.answer(
            f"🌎 Вселенная: <b>{html.escape(universe)}</b>\n\n"
            "👤 <b>За кого ты будешь играть?</b>",
            reply_markup=character_menu(),
            parse_mode="HTML"
        )

        return

    # -----------------------------------------------------
    # СОЗДАНИЕ СОБСТВЕННОГО ПЕРСОНАЖА
    # -----------------------------------------------------

    if state.get("step") == "custom_character":

        character = message.text.strip()

        if not character:

            await message.answer(
                "Напиши имя своего персонажа."
            )

            return

        state["character"] = character
        state["step"] = "playing"

        await begin_story(
            message,
            uid
        )

        return

    # -----------------------------------------------------
    # ПРОВЕРЯЕМ СОХРАНЁННУЮ ИСТОРИЮ
    # -----------------------------------------------------

    story = get_story(uid)

    if not story:

        await message.answer(
            "📭 У тебя пока нет истории.\n\n"
            "Создай её через меню.",
            reply_markup=main_menu()
        )

        return

    universe, character, previous_story = story

    # -----------------------------------------------------
    # ПРОДОЛЖЕНИЕ ИСТОРИИ
    # -----------------------------------------------------

    prompt = f"""
Ты — движок интерактивного сюжетного фанфика.

ВСЕЛЕННАЯ:
{universe}

ПЕРСОНАЖ ИГРОКА:
{character}

ИСТОРИЯ ДО ЭТОГО:
{previous_story[-18000:]}

НОВОЕ СООБЩЕНИЕ ИГРОКА:
{message.text}

СТРОГИЕ ПРАВИЛА:

1. Игрок полностью управляет своим персонажем.

2. Никогда не придумывай за персонажа игрока
его действия.

3. Никогда не придумывай за него мысли.

4. Никогда не придумывай за него эмоции.

5. Никогда не придумывай за него решения.

6. Никогда не придумывай за него реплики.

7. Никогда не заставляй его отвечать.

8. Никогда не отменяй решение игрока.

9. Если игрок написал действие,
считай его действие совершённым.

10. Ты управляешь только NPC,
каноническими персонажами, миром,
событиями и последствиями.

11. Сохраняй характеры персонажей.

12. Соблюдай непрерывность истории.

13. Не пересказывай предыдущую историю целиком.

14. Продолжай именно текущую сцену.

15. Не завершай историю без причины.

16. После своего ответа снова оставь игроку
возможность самому решить, что делать.

17. Не говори, что ты ИИ.

Продолжи историю прямо с момента,
который следует из сообщения игрока.

НЕ ПИШИ НИКАКИХ ДЕЙСТВИЙ,
МЫСЛЕЙ ИЛИ РЕПЛИК ЗА ПЕРСОНАЖА ИГРОКА.
"""

    try:

        text = await generate_ai(prompt)

    except Exception:

        logging.exception(
            "Ошибка Gemini при продолжении истории"
        )

        await message.answer(
            "❌ <b>Ошибка AI при продолжении истории.</b>\n\n"
            f"Модель: <code>{html.escape(GEMINI_MODEL)}</code>",
            parse_mode="HTML"
        )

        return

    if not text:

        text = "AI не смог продолжить сцену. Попробуй ещё раз."

    # -----------------------------------------------------
    # СОХРАНЕНИЕ
    # -----------------------------------------------------

    new_story = (
        previous_story
        + "\n\n"
        + "Игрок: "
        + message.text
        + "\n\n"
        + "Бот: "
        + text
    )

    # Чтобы запросы Gemini не становились бесконечно большими.
    new_story = new_story[-30000:]

    save_story(
        uid,
        universe,
        character,
        new_story
    )

    await message.answer(
        html.escape(text),
        parse_mode="HTML"
    )


# =========================================================
# ПРОДОЛЖИТЬ ИСТОРИЮ
# =========================================================

@dp.callback_query(F.data == "continue_story")
async def continue_story(callback: CallbackQuery):

    uid = callback.from_user.id

    story = get_story(uid)

    if not story:

        await callback.message.answer(
            "📭 У тебя пока нет сохранённой истории.\n\n"
            "Создай новую историю.",
            reply_markup=main_menu()
        )

    else:

        # ВАЖНО:
        # После перезапуска Render состояние пользователя
        # восстанавливается в режим игры.
        user_states[uid] = {
            "step": "playing"
        }

        universe, character, text = story

        await callback.message.answer(
            f"📖 <b>{html.escape(universe)}</b>\n"
            f"👤 <b>{html.escape(character)}</b>\n\n"
            "Последняя часть истории:\n\n"
            f"{html.escape(text[-5000:])}\n\n"
            "Продолжай историю своим сообщением.",
            parse_mode="HTML"
        )

    await callback.answer()


# =========================================================
# ИНФОРМАЦИЯ О ПЕРСОНАЖЕ
# =========================================================

@dp.callback_query(F.data == "character")
async def show_character(callback: CallbackQuery):

    uid = callback.from_user.id

    story = get_story(uid)

    if not story:

        await callback.message.answer(
            "👤 Персонаж ещё не создан."
        )

    else:

        universe, character, _ = story

        await callback.message.answer(
            f"👤 <b>Твой персонаж</b>\n\n"
            f"🌎 Вселенная: <b>{html.escape(universe)}</b>\n"
            f"🎭 Персонаж: <b>{html.escape(character)}</b>",
            parse_mode="HTML"
        )

    await callback.answer()


# =========================================================
# НОВАЯ ИСТОРИЯ / УДАЛЕНИЕ СТАРОЙ
# =========================================================

@dp.callback_query(F.data == "reset")
async def reset_story(callback: CallbackQuery):

    uid = callback.from_user.id

    delete_story(uid)

    user_states.pop(
        uid,
        None
    )

    await callback.message.answer(
        "🗑️ Старая история удалена.\n\n"
        "🌎 Выбери новую вселенную:",
        reply_markup=universe_menu()
    )

    await callback.answer()


# =========================================================
# WEB SERVER ДЛЯ RENDER
# =========================================================

async def health(request):
    return web.Response(
        text="Zolog Story AI is running"
    )


async def run_web_server():

    app = web.Application()

    app.router.add_get(
        "/",
        health
    )

    app.router.add_get(
        "/health",
        health
    )

    runner = web.AppRunner(
        app
    )

    await runner.setup()

    port = int(
        os.getenv(
            "PORT",
            "10000"
        )
    )

    site = web.TCPSite(
        runner,
        "0.0.0.0",
        port
    )

    await site.start()

    logging.info(
        "Web server started on port %s",
        port
    )


# =========================================================
# MAIN
# =========================================================

async def main():

    logging.info(
        "Starting Zolog Story AI..."
    )

    logging.info(
        "Gemini model: %s",
        GEMINI_MODEL
    )

    init_db()

    # HTTP-сервер необходим для Render Web Service.
    await run_web_server()

    # Удаляем webhook перед использованием polling.
    # Это предотвращает конфликт webhook/polling.
    await bot.delete_webhook(
        drop_pending_updates=False
    )

    logging.info(
        "Telegram polling started."
    )

    await dp.start_polling(
        bot
    )


if __name__ == "__main__":

    try:
        asyncio.run(
            main()
        )

    except KeyboardInterrupt:

        logging.info(
            "Bot stopped."
        )
