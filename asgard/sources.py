from __future__ import annotations

import asyncio
from datetime import datetime
from typing import Any, Dict, List, Optional
from urllib.parse import urljoin, urlparse

import feedparser
import requests
from bs4 import BeautifulSoup
from telethon import TelegramClient

from .config import get_env


def _iso(dt: Optional[datetime]) -> Optional[str]:
    if not dt:
        return None
    return dt.isoformat()


def _extract_meta(url: str, timeout: int = 15) -> Dict[str, str]:
    try:
        response = requests.get(url, timeout=timeout, headers={"User-Agent": "Mozilla/5.0"})
        response.raise_for_status()
    except requests.RequestException:
        return {}
    soup = BeautifulSoup(response.text, "html.parser")
    meta = {}
    for key in ("og:title", "og:description", "description"):
        tag = soup.find("meta", attrs={"property": key}) or soup.find("meta", attrs={"name": key})
        if tag and tag.get("content"):
            meta[key] = tag["content"].strip()
    if "og:title" not in meta and soup.title:
        meta["og:title"] = soup.title.get_text(strip=True)
    return meta


def fetch_rss(source: Dict[str, Any]) -> List[Dict[str, Any]]:
    feed = feedparser.parse(source["url"])
    items: List[Dict[str, Any]] = []
    for entry in feed.entries[: source.get("limit", 20)]:
        url = entry.get("link")
        if not url:
            continue
        title = entry.get("title", "").strip()
        summary = entry.get("summary", "").strip()
        published = entry.get("published_parsed") or entry.get("updated_parsed")
        published_at = None
        if published:
            published_at = datetime(*published[:6]).isoformat()
        items.append(
            {
                "source": source["name"],
                "source_type": "rss",
                "url": url,
                "title": title,
                "content": summary,
                "language": source.get("language"),
                "published_at": published_at,
            }
        )
    return items


def _extract_items_from_html(source: Dict[str, Any]) -> List[Dict[str, Any]]:
    try:
        response = requests.get(source["url"], timeout=20, headers={"User-Agent": "Mozilla/5.0"})
        response.raise_for_status()
    except requests.RequestException:
        return []

    soup = BeautifulSoup(response.text, "html.parser")
    item_selector = source.get("item_selector")
    link_pattern = source.get("link_pattern")
    base_url = source.get("base_url") or source["url"]
    items: List[Dict[str, Any]] = []

    elements = soup.select(item_selector) if item_selector else []
    if not elements:
        elements = soup.find_all("a")

    for element in elements:
        link_el = None
        if source.get("link_selector"):
            link_el = element.select_one(source["link_selector"])
        if not link_el and element.name == "a":
            link_el = element
        if not link_el:
            continue

        href = link_el.get("href")
        if not href:
            continue
        if link_pattern and link_pattern not in href:
            continue

        url = urljoin(base_url, href)
        title = ""
        if source.get("title_selector"):
            title_el = element.select_one(source["title_selector"])
            if title_el:
                title = title_el.get_text(strip=True)
        if not title:
            title = link_el.get_text(strip=True)

        if not title or len(title) < 6:
            continue

        items.append(
            {
                "source": source["name"],
                "source_type": "html",
                "url": url,
                "title": title,
                "content": "",
                "language": source.get("language"),
                "published_at": None,
            }
        )
        if len(items) >= source.get("limit", 20):
            break

    return items


def fetch_html(source: Dict[str, Any]) -> List[Dict[str, Any]]:
    items = _extract_items_from_html(source)
    for item in items:
        meta = _extract_meta(item["url"])
        if meta.get("og:title") and len(meta["og:title"]) > len(item["title"]):
            item["title"] = meta["og:title"]
        if meta.get("og:description"):
            item["content"] = meta["og:description"]
        elif meta.get("description"):
            item["content"] = meta["description"]
    return items


async def fetch_telegram(source: Dict[str, Any]) -> List[Dict[str, Any]]:
    api_id = get_env("TELEGRAM_API_ID")
    api_hash = get_env("TELEGRAM_API_HASH")
    phone = get_env("TELEGRAM_PHONE")
    if not api_id or not api_hash or not phone:
        return []

    session_name = source.get("session", "telethon")
    client = TelegramClient(session_name, int(api_id), api_hash)
    await client.start(phone=phone)
    channel = source["channel"].lstrip("@")
    messages = await client.get_messages(channel, limit=source.get("limit", 20))
    items: List[Dict[str, Any]] = []
    for message in messages:
        if not message.message:
            continue
        url = f"https://t.me/{channel}/{message.id}"
        items.append(
            {
                "source": source["name"],
                "source_type": "telegram",
                "url": url,
                "title": (message.message[:120] or "").split("\n")[0],
                "content": message.message,
                "language": "ru",
                "published_at": _iso(message.date),
            }
        )
    await client.disconnect()
    return items


async def fetch_all_sources(sources: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    items: List[Dict[str, Any]] = []
    for source in sources:
        source_type = source.get("type")
        if source_type == "rss":
            items.extend(await asyncio.to_thread(fetch_rss, source))
        elif source_type == "html":
            items.extend(await asyncio.to_thread(fetch_html, source))
        elif source_type == "telegram":
            items.extend(await fetch_telegram(source))
    return items
