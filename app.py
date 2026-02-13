"""
DOCX Translator — Перевод Word-документов с сохранением форматирования.
Использует Groq API для быстрого перевода через LLM.

Это точка входа приложения. Создаёт Flask-инстанс, регистрирует маршруты
и запускает сервер.

Структура проекта:
    config.py              — Настройки (env-переменные, пути, логирование)
    translation_engine.py  — Движок перевода через Groq API
    docx_translator.py     — Обход и перевод .docx документа
    routes.py              — Flask-маршруты (API endpoints)
    app.py                 — Точка входа (этот файл)
"""

from flask import Flask

from config import SECRET_KEY, PORT, FLASK_DEBUG
from routes import register_routes

# ---------------------------------------------------------------------------
# Создание и настройка Flask-приложения
# ---------------------------------------------------------------------------

app = Flask(__name__)
app.secret_key = SECRET_KEY

# Регистрируем все HTTP-маршруты из routes.py
register_routes(app)

# ---------------------------------------------------------------------------
# Запуск сервера
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=PORT, debug=FLASK_DEBUG)
