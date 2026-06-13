from __future__ import annotations

import shutil
from pathlib import Path

from build_semantic_tone_report import build_report as build_semantic_tone_report
from build_tone_word_report import build_report as build_tone_word_report
from recount_semantic_codes import build_outputs
from rebuild_itog_report import build_report


ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "output"


def main() -> None:
    if OUT.exists():
        shutil.rmtree(OUT)
    OUT.mkdir(parents=True, exist_ok=True)

    readme = OUT / "README.txt"
    readme.write_text(
        "Содержимое output\n\n"
        "01_Отчет/\n"
        "  01_Итоговый_отчет.docx — основной аналитический Word-отчёт по полному корпусу, построенный на 7 укрупнённых смысловых кодах.\n"
        "  02_Отчет_по_тональности.docx — отдельный аналитический Word-отчёт по тональности: 4 тональные категории, таблицы, диаграммы и кросс-анализ со смысловыми кодами.\n\n"
        "02_Пересчет_смысловых_кодов/\n"
        "  01_Матрица_смысловых_кодов.xlsx — полная кодировочная матрица по 7 укрупнённым смысловым кодам для атомарных высказываний.\n"
        "  02_Пояснительная_записка_по_смысловым_кодам.docx — пояснение логики укрупнения кодов, состава категорий и итогов пересчёта.\n"
        "  03_Итоги_смысловых_кодов.txt — краткая текстовая сводка итогов пересчёта по смысловым кодам.\n"
        "  04_Отчет_по_тональности.xlsx — Excel-отчёт: смысловые категории × тональность, сводные таблицы и диаграммы.\n"
        "  05_PDF_check_сегментации.txt — диагностический отчёт гибридного режима: DOCX как основной источник, PDF как дополнительная проверка качества сегментации.\n\n"
        "03_Исходные_данные/\n"
        "  01_ТЗ_контент.docx — основной исходный документ, из которого извлекается корпус ответов.\n"
        "  02_ТЗ_контент.pdf — PDF-копия исходного документа.\n"
        "  03_Методика_контент_анализа.md — методическое описание подхода, на котором основаны пересчёт и интерпретация результатов.\n",
        encoding="utf-8",
    )

    src_dir = OUT / "03_Исходные_данные"
    src_dir.mkdir(parents=True, exist_ok=True)
    root = Path(__file__).resolve().parents[1]
    for source, target_name in [
        (root / "ТЗ контент.docx", "01_ТЗ_контент.docx"),
        (root / "ТЗ контент.pdf", "02_ТЗ_контент.pdf"),
        (root / "Вот первое задание. Есть в социологии метод контен.md", "03_Методика_контент_анализа.md"),
    ]:
        if source.exists():
            shutil.copy2(source, src_dir / target_name)

    semantic = build_outputs()
    report = build_report()
    semantic_tone = build_semantic_tone_report()
    tone_word = build_tone_word_report()

    print("Пересборка завершена без архивов:")
    print(f"- {semantic['xlsx_path']}")
    print(f"- {semantic['docx_path']}")
    print(f"- {semantic['txt_path']}")
    print(f"- {semantic['pdf_check_path']}")
    print(f"- {semantic_tone}")
    print(f"- {tone_word}")
    print(f"- {report}")


if __name__ == "__main__":
    main()
