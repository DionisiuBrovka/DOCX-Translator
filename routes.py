"""
Flask-маршруты (API endpoints) для DOCX Translator.

Маршруты:
    GET  /                  — Отдаёт веб-интерфейс (SPA).
    POST /translate         — Загрузка и перевод документа.
    GET  /progress/<id>     — Получение текущего прогресса задачи.
    GET  /download/<id>     — Скачивание переведённого файла.

Хранение состояния:
    Прогресс и метаданные задач хранятся в словаре translation_progress
    прямо в памяти процесса. При перезапуске сервера всё теряется.
    Для продакшена нужно использовать Redis, БД или брокер задач (Celery/RQ).
"""

import uuid
import logging
from pathlib import Path

from flask import render_template, request, send_file, jsonify

from config import UPLOAD_FOLDER, OUTPUT_FOLDER, GROQ_API_KEY, GROQ_MODEL
from translation_engine import TranslationEngine
from docx_translator import DocxTranslator

logger = logging.getLogger("docx_translator")

# ---------------------------------------------------------------------------
# In-memory хранилище прогресса задач
# ---------------------------------------------------------------------------
# Ключ — task_id (str), значение — dict с полями:
#   progress (float)   — процент выполнения (0–100)
#   status (str)       — "processing" | "done" | "error"
#   output_name (str)  — имя файла для скачивания
#   output_path (str)  — путь к переведённому файлу на диске
#   stats (dict)       — статистика перевода (после завершения)
#   error (str)        — текст ошибки (при status="error")
translation_progress: dict[str, dict] = {}


def register_routes(app):
    """
    Регистрирует все маршруты на переданном Flask-приложении.

    Вызывается из app.py при инициализации. Такой подход позволяет
    держать маршруты отдельно от создания Flask-инстанса.
    """

    @app.route("/")
    def index():
        """Главная страница — отдаёт SPA-интерфейс с формой загрузки."""
        return render_template("index.html", groq_model=GROQ_MODEL)

    @app.route("/translate", methods=["POST"])
    def translate():
        """
        Загрузка и перевод документа.

        Ожидаемые поля формы (multipart/form-data):
            file         — .docx файл для перевода
            api_key      — (опционально) Groq API ключ, если не задан в env
            model        — (опционально) модель LLM
            source_lang  — (опционально) язык оригинала (по умолчанию Russian)
            target_lang  — (опционально) целевой язык (по умолчанию English)

        Возвращает JSON:
            Успех: { task_id, status, stats, download_url }
            Ошибка: { error } с HTTP 400 или 500
        """
        # --- Валидация API-ключа ---
        api_key = request.form.get("api_key", "").strip() or GROQ_API_KEY
        if not api_key:
            return jsonify({"error": "Groq API key is required"}), 400

        # --- Валидация файла ---
        file = request.files.get("file")
        if not file or not file.filename:
            return jsonify({"error": "No file uploaded"}), 400

        if not file.filename.lower().endswith(".docx"):
            return jsonify({"error": "Only .docx files are supported"}), 400

        # --- Параметры перевода ---
        model = request.form.get("model", GROQ_MODEL).strip()
        source_lang = request.form.get("source_lang", "Russian").strip()
        target_lang = request.form.get("target_lang", "English").strip()

        # --- Сохранение загруженного файла ---
        task_id = str(uuid.uuid4())[:8]
        original_name = Path(file.filename).stem
        input_path = UPLOAD_FOLDER / f"{task_id}_{file.filename}"
        output_name = f"{original_name}_translated.docx"
        output_path = OUTPUT_FOLDER / f"{task_id}_{output_name}"

        file.save(str(input_path))

        # Инициализируем запись прогресса
        translation_progress[task_id] = {
            "progress": 0,
            "status": "processing",
            "output_name": output_name,
            "output_path": str(output_path),
        }

        # --- Перевод (синхронный) ---
        # ВНИМАНИЕ: блокирует поток на всё время перевода.
        # Для продакшена нужен Celery/RQ для асинхронной обработки.
        try:
            engine = TranslationEngine(
                api_key=api_key,
                model=model,
                source_lang=source_lang,
                target_lang=target_lang,
            )
            translator = DocxTranslator(engine)

            # Callback для обновления прогресса в реальном времени
            def on_progress(pct):
                translation_progress[task_id]["progress"] = round(pct, 1)

            stats = translator.translate_document(
                str(input_path), str(output_path), progress_callback=on_progress
            )

            # Обновляем статус — перевод успешно завершён
            translation_progress[task_id].update({
                "status": "done",
                "progress": 100,
                "stats": stats,
            })

            return jsonify({
                "task_id": task_id,
                "status": "done",
                "stats": stats,
                "download_url": f"/download/{task_id}",
            })

        except Exception as e:
            logger.exception("Ошибка при переводе")
            translation_progress[task_id].update({
                "status": "error",
                "error": str(e),
            })
            return jsonify({"error": str(e)}), 500

        finally:
            # Всегда удаляем загруженный файл — он больше не нужен
            try:
                input_path.unlink()
            except OSError:
                pass

    @app.route("/progress/<task_id>")
    def progress(task_id):
        """
        Получение текущего прогресса перевода.

        Фронтенд опрашивает этот endpoint с интервалом, чтобы
        обновлять прогресс-бар в реальном времени.
        """
        info = translation_progress.get(task_id)
        if not info:
            return jsonify({"error": "Task not found"}), 404
        return jsonify(info)

    @app.route("/download/<task_id>")
    def download(task_id):
        """
        Скачивание переведённого документа.

        Доступен только для задач со статусом "done".
        Отдаёт файл с правильным MIME-типом для .docx.
        """
        info = translation_progress.get(task_id)
        if not info or info["status"] != "done":
            return jsonify({"error": "File not ready"}), 404

        output_path = info["output_path"]
        if not Path(output_path).exists():
            return jsonify({"error": "File not found"}), 404

        return send_file(
            output_path,
            as_attachment=True,
            download_name=info["output_name"],
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
