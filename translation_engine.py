"""
Движок перевода текста через Groq API.

Основная идея — перевод на уровне «ранов» (runs). В формате .docx каждый
параграф состоит из ранов, где каждый ран хранит свой набор форматирования
(жирный, курсив, размер шрифта и т.д.). Чтобы сохранить это форматирование
при переводе, мы переводим раны как структурированный JSON-массив, чтобы
LLM вернула ровно столько же фрагментов, сколько было на входе.

Два режима перевода:
  - Простой (_translate_simple): один фрагмент текста → один перевод.
  - Структурный (_translate_structured): массив фрагментов → массив переводов,
    с сохранением количества и порядка элементов.

Если структурный перевод не удался (LLM вернула невалидный JSON или неверное
количество элементов), срабатывает fallback: весь параграф переводится целиком
и помещается в первый ран.
"""

import json
import logging

from groq import Groq

from config import GROQ_MODEL

logger = logging.getLogger("docx_translator")


class TranslationEngine:
    """Переводчик текста через Groq API с сохранением структуры ранов."""

    def __init__(self, api_key: str, model: str = GROQ_MODEL,
                 source_lang: str = "Russian", target_lang: str = "English"):
        """
        Инициализация движка перевода.

        Args:
            api_key: Ключ доступа к Groq API.
            model: Идентификатор LLM-модели (по умолчанию из конфига).
            source_lang: Язык оригинала (для системного промпта LLM).
            target_lang: Целевой язык перевода.
        """
        self.client = Groq(api_key=api_key)
        self.model = model
        self.source_lang = source_lang
        self.target_lang = target_lang

        # Счётчики для статистики: сколько запросов отправлено и токенов потрачено
        self.request_count = 0
        self.total_tokens = 0

    # -----------------------------------------------------------------
    # Публичный метод — точка входа
    # -----------------------------------------------------------------

    def translate_runs(self, runs_text: list[str]) -> list[str]:
        """
        Переводит список текстовых фрагментов (ранов), сохраняя их количество
        и соответствие. Это ключ к сохранению inline-форматирования.

        Пример:
            Вход:  ["Привет, ", "мир!"]
            Выход: ["Hello, ", "world!"]

        Логика выбора стратегии:
            - Пустой текст → возвращаем как есть.
            - Один ран → простой перевод (_translate_simple).
            - Несколько ранов → структурный перевод (_translate_structured).
        """
        # Склеиваем все раны, чтобы проверить, есть ли вообще текст
        combined = "".join(runs_text).strip()
        if not combined:
            return runs_text

        # Один фрагмент — нет смысла в JSON-массиве, переводим напрямую
        if len(runs_text) == 1:
            translated = self._translate_simple(runs_text[0])
            return [translated]

        # Несколько фрагментов — структурный перевод через JSON-массив
        return self._translate_structured(runs_text)

    # -----------------------------------------------------------------
    # Простой перевод (один фрагмент)
    # -----------------------------------------------------------------

    def _translate_simple(self, text: str) -> str:
        """
        Переводит одиночный текстовый фрагмент.

        Важно: сохраняем ведущие и завершающие пробелы — они могут быть
        значимы для форматирования документа (например, пробел между
        ранами с разным форматированием).
        """
        if not text.strip():
            return text

        # Запоминаем пробелы, чтобы вернуть их после перевода
        leading = text[:len(text) - len(text.lstrip())]
        trailing = text[len(text.rstrip()):]

        response = self.client.chat.completions.create(
            model=self.model,
            messages=[
                {
                    "role": "system",
                    "content": (
                        f"You are a professional translator from {self.source_lang} "
                        f"to {self.target_lang}. Translate the user's text accurately "
                        f"and naturally. Return ONLY the translated text, nothing else. "
                        f"Do not add explanations, notes, or quotes around the text. "
                        f"Preserve any leading/trailing whitespace exactly."
                    )
                },
                {"role": "user", "content": text.strip()}
            ],
            # Температура 0.3 — чуть выше нуля для естественности,
            # но достаточно низкая для стабильных результатов
            temperature=0.3,
            max_tokens=4096,
        )

        # Обновляем счётчики статистики
        self.request_count += 1
        self.total_tokens += response.usage.total_tokens if response.usage else 0

        result = response.choices[0].message.content.strip()
        return leading + result + trailing

    # -----------------------------------------------------------------
    # Структурный перевод (несколько фрагментов)
    # -----------------------------------------------------------------

    def _translate_structured(self, runs_text: list[str]) -> list[str]:
        """
        Переводит несколько фрагментов одновременно, сохраняя их количество
        и порядок. Фрагменты отправляются как JSON-массив, и LLM должна
        вернуть JSON-массив той же длины.

        Алгоритм:
            1. Отфильтровать пустые раны (пробелы, пустые строки).
            2. Отправить непустые фрагменты как JSON-массив в Groq API.
            3. Распарсить JSON-ответ.
            4. Подставить переводы на места оригинальных непустых ранов.

        Fallback: если JSON не парсится или длина массива не совпадает,
        переводим весь параграф целиком простым методом и кладём в первый ран.
        """
        # Разделяем пустые и непустые раны, запоминая их позиции
        indexed = [(i, t) for i, t in enumerate(runs_text)]
        non_empty = [(i, t) for i, t in indexed if t.strip()]

        if not non_empty:
            return runs_text

        # Формируем JSON-массив только из непустых фрагментов
        fragments = [t for _, t in non_empty]
        fragments_json = json.dumps(fragments, ensure_ascii=False)

        response = self.client.chat.completions.create(
            model=self.model,
            messages=[
                {
                    "role": "system",
                    "content": (
                        f"You are a professional translator from {self.source_lang} "
                        f"to {self.target_lang}.\n\n"
                        f"TASK: Translate a JSON array of text fragments. These fragments "
                        f"form a single paragraph when concatenated. Each fragment has "
                        f"its own formatting (bold, italic, etc.) that must be preserved "
                        f"through maintaining the array structure.\n\n"
                        f"RULES:\n"
                        f"1. Return a valid JSON array with EXACTLY {len(fragments)} elements\n"
                        f"2. Each element corresponds to the same-index input element\n"
                        f"3. Preserve leading/trailing spaces in each element\n"
                        f"4. The concatenation of all elements must read as natural "
                        f"{self.target_lang}\n"
                        f"5. Return ONLY the JSON array, no explanation or markdown\n"
                        f"6. If a fragment is just whitespace or punctuation, keep it as-is"
                    )
                },
                {"role": "user", "content": fragments_json}
            ],
            # Температура 0.2 — ниже, чем для простого перевода,
            # т.к. нужен строгий JSON-формат на выходе
            temperature=0.2,
            max_tokens=4096,
        )

        self.request_count += 1
        self.total_tokens += response.usage.total_tokens if response.usage else 0

        raw = response.choices[0].message.content.strip()

        # --- Парсинг JSON-ответа ---
        try:
            # LLM иногда оборачивает JSON в markdown-блок ```json ... ```
            if raw.startswith("```"):
                raw = raw.split("\n", 1)[1] if "\n" in raw else raw[3:]
                if raw.endswith("```"):
                    raw = raw[:-3]
                raw = raw.strip()

            translated = json.loads(raw)

            # Проверяем, что ответ — массив нужной длины
            if isinstance(translated, list) and len(translated) == len(fragments):
                # Собираем результат: копируем оригинал, затем подставляем
                # переводы на позиции непустых ранов
                result = list(runs_text)
                for (orig_idx, _), trans in zip(non_empty, translated):
                    result[orig_idx] = str(trans)
                return result

        except (json.JSONDecodeError, ValueError) as e:
            logger.warning(f"Не удалось распарсить структурный ответ, fallback: {e}")

        # --- Fallback ---
        # Если структурный перевод не удался, переводим параграф целиком
        # и помещаем результат в первый ран, остальные обнуляем.
        full_text = "".join(runs_text)
        translated_full = self._translate_simple(full_text)
        result = [""] * len(runs_text)
        result[0] = translated_full
        return result
