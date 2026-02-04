# Asgard Autopost (Telegram)

Автоматизация для Telegram-канала с подтверждением в боте:

- сбор новостей из источников (включая Telegram-каналы),
- перевод/сжатие через локальный Ollama,
- авто‑картинки на каждый пост,
- утверждение кнопками ✅/❌,
- публикация по расписанию.

## Быстрый старт (Mac)

### 1) Установить зависимости

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

### 2) Установить Ollama и модель

```bash
brew install ollama
ollama serve
ollama pull llama3.1:8b
```

### 3) Настроить .env

```bash
cp .env.example .env
```

Заполните:

- `BOT_TOKEN` — токен от BotFather
- `CHANNEL_ID` — например `@asgardchanel`
- `TELEGRAM_API_ID`, `TELEGRAM_API_HASH`, `TELEGRAM_PHONE` — для чтения DeCenter

### 4) Добавить шаблон картинки

Положите шаблон в `templates/base.png`.

Если шаблона нет — скрипт создаст простой белый фон с заголовком.

### 5) Запустить

```bash
python app.py
```

В боте:
- `/start` — сохранить админ‑чат
- `/queue` — отправить черновики на проверку
- `/stats` — статистика

## Как работает

1) Каждые 30 минут сбор источников из `sources.yaml`  
2) LLM готовит текст по `prompt_news.txt` / `prompt_explain.txt`  
3) Генерируется картинка (шаблон + заголовок)  
4) Бот шлёт черновики на подтверждение  
5) Публикация по расписанию из `config.yaml`  

## Настройка

Все параметры — в `config.yaml`:
- расписание постов,
- ключевые слова рубрик,
- включение ссылки на источник,
- параметры генерации изображений.

Источники — в `sources.yaml`.
