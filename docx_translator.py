"""
Обход и перевод .docx документа с сохранением форматирования.

Этот модуль отвечает за навигацию по структуре Word-документа:
  - Параграфы тела документа
  - Таблицы (включая вложенные)
  - Колонтитулы (headers/footers) всех типов

Сам перевод делегируется TranslationEngine — этот модуль лишь
определяет, ЧТО переводить и в каком порядке.

Ключевые особенности:
  - Кириллический фильтр: параграфы без кириллицы пропускаются
    (считаем, что они уже на целевом языке).
  - Прогресс-callback: вызывается после каждого обработанного элемента,
    чтобы фронтенд мог показывать прогресс-бар.
  - Статистика: собирается по ходу обработки (параграфы, ячейки,
    колонтитулы, пропущенные элементы, запросы к API, токены).
"""

import logging

from docx import Document

from translation_engine import TranslationEngine

logger = logging.getLogger("docx_translator")


class DocxTranslator:
    """Обходит .docx документ и переводит весь текст, сохраняя форматирование."""

    def __init__(self, engine: TranslationEngine):
        """
        Args:
            engine: Экземпляр TranslationEngine для выполнения переводов.
        """
        self.engine = engine

        # Статистика обработки — заполняется в процессе перевода
        self.stats = {
            "paragraphs": 0,       # Обработано параграфов в теле
            "table_cells": 0,      # Обработано ячеек таблиц
            "headers_footers": 0,  # Обработано элементов колонтитулов
            "skipped_empty": 0,    # Пропущено пустых/некириллических параграфов
        }

    # -----------------------------------------------------------------
    # Основной метод
    # -----------------------------------------------------------------

    def translate_document(self, input_path: str, output_path: str,
                           progress_callback=None) -> dict:
        """
        Главный метод: читает .docx, переводит все текстовые элементы,
        сохраняет результат.

        Args:
            input_path: Путь к исходному .docx файлу.
            output_path: Путь для сохранения переведённого файла.
            progress_callback: Функция(процент), вызывается при обновлении прогресса.

        Returns:
            Словарь со статистикой: количество обработанных элементов,
            API-запросов и потраченных токенов.
        """
        doc = Document(input_path)

        # Считаем общее количество элементов заранее,
        # чтобы корректно вычислять процент прогресса
        total = self._count_elements(doc)
        current = 0

        def update_progress(increment=1):
            """Внутренний хелпер: обновляет прогресс и вызывает callback."""
            nonlocal current
            current += increment
            if progress_callback and total > 0:
                progress_callback(min(current / total * 100, 100))

        # --- Этап 1: Параграфы тела документа ---
        logger.info("Перевод параграфов тела документа...")
        for para in doc.paragraphs:
            self._translate_paragraph(para)
            self.stats["paragraphs"] += 1
            update_progress()

        # --- Этап 2: Таблицы ---
        logger.info("Перевод таблиц...")
        for table in doc.tables:
            self._translate_table(table, update_progress)

        # --- Этап 3: Колонтитулы (headers и footers) ---
        # В .docx у каждой секции может быть до 3 типов header и footer:
        #   - default (основной)
        #   - first_page (для первой страницы секции)
        #   - even_page (для чётных страниц)
        logger.info("Перевод колонтитулов...")
        for section in doc.sections:
            # Хедеры
            for header in [section.header, section.first_page_header,
                           section.even_page_header]:
                if header and header.is_linked_to_previous is False:
                    for para in header.paragraphs:
                        self._translate_paragraph(para)
                        self.stats["headers_footers"] += 1
                        update_progress()
                    for table in header.tables:
                        self._translate_table(table, update_progress)

            # Футеры
            for footer in [section.footer, section.first_page_footer,
                           section.even_page_footer]:
                if footer and footer.is_linked_to_previous is False:
                    for para in footer.paragraphs:
                        self._translate_paragraph(para)
                        self.stats["headers_footers"] += 1
                        update_progress()
                    for table in footer.tables:
                        self._translate_table(table, update_progress)

        # Сохраняем документ. Картинки и OLE-объекты сохраняются
        # автоматически — python-docx не трогает бинарные части архива.
        doc.save(output_path)
        logger.info(f"Перевод завершён: {output_path}")

        return {
            **self.stats,
            "api_requests": self.engine.request_count,
            "total_tokens": self.engine.total_tokens,
        }

    # -----------------------------------------------------------------
    # Перевод параграфа
    # -----------------------------------------------------------------

    def _translate_paragraph(self, para):
        """
        Переводит один параграф, работая на уровне ранов.

        Пропускает параграф, если:
          - В нём нет ранов (пустой или содержит только спец-элементы).
          - Текст пустой после объединения всех ранов.
          - Текст не содержит кириллицу (вероятно, уже на целевом языке).
        """
        runs = para.runs
        if not runs:
            self.stats["skipped_empty"] += 1
            return

        runs_text = [run.text for run in runs]
        combined = "".join(runs_text).strip()

        if not combined:
            self.stats["skipped_empty"] += 1
            return

        # Проверка на наличие кириллицы (Unicode-диапазон \u0400–\u04FF).
        # Если кириллицы нет — считаем, что текст уже на целевом языке.
        if not any('\u0400' <= c <= '\u04FF' for c in combined):
            return

        # Переводим через движок и записываем результат обратно в раны.
        # Форматирование (bold, italic, font и т.д.) остаётся нетронутым,
        # т.к. мы меняем только свойство .text у каждого рана.
        translated = self.engine.translate_runs(runs_text)

        for run, new_text in zip(runs, translated):
            run.text = new_text

    # -----------------------------------------------------------------
    # Перевод таблиц
    # -----------------------------------------------------------------

    def _translate_table(self, table, update_progress):
        """
        Переводит все ячейки таблицы, включая вложенные таблицы.

        Word-таблицы могут содержать вложенные таблицы внутри ячеек —
        обрабатываем их рекурсивно.
        """
        for row in table.rows:
            for cell in row.cells:
                # Параграфы внутри ячейки
                for para in cell.paragraphs:
                    self._translate_paragraph(para)
                    self.stats["table_cells"] += 1
                    update_progress()

                # Рекурсивно обрабатываем вложенные таблицы
                for nested_table in cell.tables:
                    self._translate_table(nested_table, update_progress)

    # -----------------------------------------------------------------
    # Подсчёт элементов для прогресс-бара
    # -----------------------------------------------------------------

    def _count_elements(self, doc) -> int:
        """
        Считает общее количество переводимых элементов в документе.
        Используется для расчёта процента прогресса.

        Учитывает: параграфы тела, ячейки таблиц, элементы колонтитулов.
        """
        count = len(doc.paragraphs)

        for table in doc.tables:
            count += self._count_table_elements(table)

        for section in doc.sections:
            for hf in [section.header, section.footer,
                       section.first_page_header, section.first_page_footer,
                       section.even_page_header, section.even_page_footer]:
                if hf:
                    count += len(hf.paragraphs)
                    for table in hf.tables:
                        count += self._count_table_elements(table)

        return count

    def _count_table_elements(self, table) -> int:
        """Рекурсивно считает параграфы в таблице (включая вложенные)."""
        count = 0
        for row in table.rows:
            for cell in row.cells:
                count += len(cell.paragraphs)
                for nested in cell.tables:
                    count += self._count_table_elements(nested)
        return count
