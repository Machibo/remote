from __future__ import annotations

import re
from datetime import datetime, timedelta, time
from typing import Dict, Iterable, List, Optional, Tuple
from zoneinfo import ZoneInfo


def normalize_text(text: str) -> str:
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    text = re.sub(r"🧩\s*(\S)", r"🧩 \1", text)
    text = re.sub(r"📌\s*(\S)", r"📌 \1", text)
    return text.strip()


def ensure_hashtag(text: str, hashtag: str) -> str:
    if "#" in text:
        return text
    return f"{text}\n\n{hashtag}"


def select_category(text: str, categories: Dict[str, Dict[str, Iterable[str]]]) -> Tuple[Optional[str], Optional[str]]:
    lowered = text.lower()
    for name, data in categories.items():
        for keyword in data.get("keywords", []):
            if keyword.lower() in lowered:
                return name, data.get("hashtag")
    return None, None


def detect_explain(text: str, keywords: Iterable[str]) -> bool:
    lowered = text.lower()
    for keyword in keywords:
        if keyword.lower() in lowered:
            return True
    return False


def next_schedule_time(
    now: datetime,
    schedule_times: List[str],
    tz: ZoneInfo,
) -> datetime:
    if not schedule_times:
        return now.astimezone(tz) + timedelta(minutes=5)
    local_now = now.astimezone(tz)
    today = local_now.date()
    times: List[time] = []
    for item in schedule_times:
        hours, minutes = item.split(":")
        times.append(time(int(hours), int(minutes)))
    for slot in times:
        candidate = datetime.combine(today, slot, tzinfo=tz)
        if candidate > local_now:
            return candidate
    tomorrow = today + timedelta(days=1)
    return datetime.combine(tomorrow, times[0], tzinfo=tz)
