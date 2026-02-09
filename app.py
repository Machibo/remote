from __future__ import annotations

import asyncio
import logging
import os
import re
from datetime import datetime, timezone
from typing import Optional

from telegram import InlineKeyboardButton, InlineKeyboardMarkup, Update
from telegram.ext import (
    Application,
    CallbackQueryHandler,
    CommandHandler,
    ContextTypes,
)

from asgard.ai import generate_post
from asgard.config import get_env, load_config, load_env, load_sources
from asgard.db import (
    fetch_by_status,
    fetch_due_posts,
    fetch_prepared_unnotified,
    get_connection,
    get_setting,
    init_db,
    mark_notified,
    mark_posted,
    mark_prepared,
    set_setting,
    set_status,
    upsert_item,
)
from asgard.image import generate_image
from asgard.sources import fetch_all_sources
from asgard.utils import (
    detect_explain,
    ensure_hashtag,
    normalize_text,
    next_schedule_time,
    select_category,
)


logging.basicConfig(level=logging.INFO)
LOGGER = logging.getLogger("asgard")

CONFIG = load_config()
PROMPT_NEWS_PATH = CONFIG.ai.get("prompt_news", "prompt_news.txt")
PROMPT_EXPLAIN_PATH = CONFIG.ai.get("prompt_explain", "prompt_explain.txt")


def read_prompt(path: str) -> str:
    with open(path, "r", encoding="utf-8") as handle:
        return handle.read()


PROMPT_NEWS = read_prompt(PROMPT_NEWS_PATH)
PROMPT_EXPLAIN = read_prompt(PROMPT_EXPLAIN_PATH)


def get_db():
    conn = get_connection()
    init_db(conn)
    return conn


def get_admin_chat_id(conn) -> Optional[int]:
    env_admin = get_env("ADMIN_CHAT_ID")
    if env_admin:
        return int(env_admin)
    value = get_setting(conn, "admin_chat_id")
    if value:
        return int(value)
    return None


def extract_title_for_image(text: str) -> str:
    first_line = text.splitlines()[0].strip()
    first_line = re.sub(r"^[^A-Za-zА-Яа-я0-9]+", "", first_line)
    max_chars = int(CONFIG.app.get("max_title_chars", 90))
    return first_line[:max_chars].strip()


def build_post_text(raw: str, category_hashtag: Optional[str], source_url: Optional[str]) -> str:
    normalized = normalize_text(raw)
    if CONFIG.app.get("include_source_link", False) and source_url:
        normalized = f"{normalized}\n\nИсточник: {source_url}"
    hashtag = category_hashtag or CONFIG.app.get("default_hashtag", "#майнинг")
    return ensure_hashtag(normalized, hashtag)


def schedule_next_slot(now_utc: datetime) -> str:
    if now_utc.tzinfo is None:
        now_utc = now_utc.replace(tzinfo=timezone.utc)
    tz_name = CONFIG.app.get("timezone", "Europe/Moscow")
    tz = timezone.utc
    try:
        from zoneinfo import ZoneInfo

        tz = ZoneInfo(tz_name)
    except Exception:
        tz = timezone.utc
    schedule_times = CONFIG.app.get("schedule_times", [])
    next_local = next_schedule_time(now_utc, schedule_times, tz)
    next_utc = next_local.astimezone(timezone.utc).replace(tzinfo=None)
    return next_utc.isoformat(timespec="seconds")


async def fetch_and_prepare() -> None:
    sources = load_sources()
    items = await fetch_all_sources(sources)
    conn = get_db()
    for item in items:
        upsert_item(conn, item)
    await asyncio.to_thread(prepare_new_items)


def prepare_new_items() -> None:
    conn = get_db()
    limit = int(CONFIG.app.get("review_batch_size", 8)) * 2
    new_items = fetch_by_status(conn, "new", limit=limit)
    if not new_items:
        return
    model = get_env("OLLAMA_MODEL") or CONFIG.ai.get("model")
    temperature = float(CONFIG.ai.get("temperature", 0.2))
    max_tokens = int(CONFIG.ai.get("max_tokens", 800))
    explain_keywords = CONFIG.ai.get("explain_keywords", [])
    for item in new_items:
        content = f"{item['title']}\n\n{item['content'] or ''}".strip()
        if not content:
            continue
        use_explain = detect_explain(content, explain_keywords)
        template = PROMPT_EXPLAIN if use_explain else PROMPT_NEWS
        generated = generate_post(template, content, model, temperature, max_tokens)
        if not generated or len(generated) < 200:
            continue
        category_key, category_hashtag = select_category(
            generated + " " + item["title"], CONFIG.categories
        )
        post_text = build_post_text(generated, category_hashtag, item["url"])

        image_path = None
        image_cfg = CONFIG.image
        if image_cfg.get("enabled", True):
            title = extract_title_for_image(post_text)
            image_path = generate_image(
                title=title,
                template_path=image_cfg.get("template_path", "templates/base.png"),
                output_dir=image_cfg.get("output_dir", "output"),
                size=tuple(image_cfg.get("size", [1080, 1080])),
                title_box=tuple(image_cfg.get("title_box", [120, 150, 840, 260])),
                title_color=image_cfg.get("title_color", "#22C55E"),
                title_font_size=int(image_cfg.get("title_font_size", 72)),
                title_font_paths=image_cfg.get("title_font_paths", []),
                stroke_width=int(image_cfg.get("stroke_width", 2)),
                stroke_fill=image_cfg.get("stroke_fill", "#000000"),
                background_color=image_cfg.get("background_color", "#FFFFFF"),
                uppercase=bool(CONFIG.app.get("title_uppercase", True)),
            )

        mark_prepared(
            conn,
            item_id=item["id"],
            summary=post_text,
            image_path=image_path,
            category=category_key,
        )


def build_review_keyboard(item_id: int) -> InlineKeyboardMarkup:
    keyboard = [
        [
            InlineKeyboardButton("✅ Утвердить", callback_data=f"approve:{item_id}"),
            InlineKeyboardButton("❌ Отклонить", callback_data=f"reject:{item_id}"),
        ]
    ]
    return InlineKeyboardMarkup(keyboard)


async def send_for_review(context: ContextTypes.DEFAULT_TYPE) -> None:
    conn = get_db()
    admin_chat_id = get_admin_chat_id(conn)
    if not admin_chat_id:
        return
    pending = fetch_prepared_unnotified(
        conn, limit=int(CONFIG.app.get("review_batch_size", 8))
    )
    if not pending:
        return
    for item in pending:
        keyboard = build_review_keyboard(item["id"])
        text = f"ID {item['id']} | {item['source']}\n\n{item['summary']}"
        message_id = None
        if item["image_path"] and os.path.exists(item["image_path"]):
            with open(item["image_path"], "rb") as image:
                await context.bot.send_photo(
                    chat_id=admin_chat_id,
                    photo=image,
                )
        msg = await context.bot.send_message(
            chat_id=admin_chat_id,
            text=text,
            reply_markup=keyboard,
        )
        message_id = str(msg.message_id)
        mark_notified(conn, item["id"], message_id)


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    conn = get_db()
    set_setting(conn, "admin_chat_id", str(update.effective_chat.id))
    await update.message.reply_text("Админ-чат сохранен. Бот готов к работе.")


async def queue(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await send_for_review(context)
    await update.message.reply_text("Отправил новые черновики на проверку.")


async def stats(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    conn = get_db()
    prepared = fetch_by_status(conn, "prepared", limit=9999)
    approved = fetch_by_status(conn, "approved", limit=9999)
    posted = fetch_by_status(conn, "posted", limit=9999)
    await update.message.reply_text(
        f"Готово к проверке: {len(prepared)}\n"
        f"Утверждено: {len(approved)}\n"
        f"Опубликовано: {len(posted)}"
    )


async def handle_callback(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    query = update.callback_query
    if not query:
        return
    await query.answer()
    action, item_id_str = query.data.split(":")
    item_id = int(item_id_str)
    conn = get_db()
    if action == "approve":
        scheduled_at = None
        if CONFIG.app.get("auto_schedule", True):
            scheduled_at = schedule_next_slot(datetime.now(timezone.utc))
        set_status(conn, item_id, "approved", scheduled_at=scheduled_at)
        note = "Утверждено"
        if not CONFIG.app.get("publish_enabled", True):
            note += " (публикация отключена)"
        if scheduled_at:
            note += f" → публикация {scheduled_at} UTC"
        base_text = query.message.caption or query.message.text or ""
        updated = f"{base_text}\n\n✅ {note}"
        if query.message.caption:
            await query.edit_message_caption(caption=updated[:1024])
        else:
            await query.edit_message_text(updated[:4096])
    elif action == "reject":
        set_status(conn, item_id, "rejected")
        if query.message.caption:
            await query.edit_message_caption(
                caption=f"{query.message.caption}\n\n❌ Отклонено"[:1024]
            )
        else:
            await query.edit_message_text(f"{query.message.text}\n\n❌ Отклонено")


async def post_due(context: ContextTypes.DEFAULT_TYPE) -> None:
    if not CONFIG.app.get("publish_enabled", True):
        return
    conn = get_db()
    now_iso = datetime.utcnow().replace(tzinfo=None).isoformat(timespec="seconds")
    due = fetch_due_posts(conn, now_iso, limit=10)
    if not due:
        return
    channel_id = get_env("CHANNEL_ID")
    if not channel_id:
        return
    for item in due:
        message_id = None
        if item["image_path"] and os.path.exists(item["image_path"]):
            with open(item["image_path"], "rb") as image:
                if len(item["summary"]) <= 1024:
                    msg = await context.bot.send_photo(
                        chat_id=channel_id,
                        photo=image,
                        caption=item["summary"],
                    )
                    message_id = str(msg.message_id)
                else:
                    msg = await context.bot.send_photo(
                        chat_id=channel_id,
                        photo=image,
                    )
                    await context.bot.send_message(
                        chat_id=channel_id,
                        text=item["summary"],
                    )
                    message_id = str(msg.message_id)
        else:
            msg = await context.bot.send_message(
                chat_id=channel_id,
                text=item["summary"],
            )
            message_id = str(msg.message_id)
        mark_posted(conn, item["id"], message_id)


async def fetch_job(context: ContextTypes.DEFAULT_TYPE) -> None:
    await fetch_and_prepare()


def main() -> None:
    load_env()
    token = get_env("BOT_TOKEN")
    if not token:
        raise SystemExit("BOT_TOKEN not set in .env")

    application = Application.builder().token(token).build()
    application.add_handler(CommandHandler("start", start))
    application.add_handler(CommandHandler("queue", queue))
    application.add_handler(CommandHandler("stats", stats))
    application.add_handler(CallbackQueryHandler(handle_callback, pattern="^(approve|reject):"))

    application.job_queue.run_repeating(
        fetch_job,
        interval=int(CONFIG.app.get("fetch_interval_minutes", 30)) * 60,
        first=5,
    )
    application.job_queue.run_repeating(send_for_review, interval=60, first=10)
    if CONFIG.app.get("publish_enabled", True):
        application.job_queue.run_repeating(post_due, interval=60, first=15)
    else:
        LOGGER.info("Publishing disabled: posts will not be sent to channel.")

    LOGGER.info("Asgard bot started.")
    application.run_polling()


if __name__ == "__main__":
    main()
