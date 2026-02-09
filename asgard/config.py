from __future__ import annotations

import os
from dataclasses import dataclass
from typing import Any, Dict, List

import yaml
from dotenv import load_dotenv


@dataclass
class AppConfig:
    data: Dict[str, Any]

    @property
    def app(self) -> Dict[str, Any]:
        return self.data.get("app", {})

    @property
    def ai(self) -> Dict[str, Any]:
        return self.data.get("ai", {})

    @property
    def image(self) -> Dict[str, Any]:
        return self.data.get("image", {})

    @property
    def categories(self) -> Dict[str, Any]:
        return self.data.get("categories", {})


def load_yaml(path: str) -> Dict[str, Any]:
    with open(path, "r", encoding="utf-8") as handle:
        return yaml.safe_load(handle) or {}


def load_config(config_path: str = "config.yaml") -> AppConfig:
    return AppConfig(load_yaml(config_path))


def load_sources(sources_path: str = "sources.yaml") -> List[Dict[str, Any]]:
    data = load_yaml(sources_path)
    return data.get("sources", [])


def load_env(env_path: str = ".env") -> None:
    if os.path.exists(env_path):
        load_dotenv(env_path)


def get_env(key: str, default: str | None = None) -> str | None:
    value = os.environ.get(key)
    if value:
        return value
    return default
