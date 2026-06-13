from __future__ import annotations

import shutil
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
ARCHIVE = ROOT / "Архив"


def move_if_exists(src: Path, dst: Path) -> None:
    if not src.exists():
        return
    dst.parent.mkdir(parents=True, exist_ok=True)
    if dst.exists():
        if dst.is_file():
            dst.unlink()
        else:
            shutil.rmtree(dst)
    shutil.move(str(src), str(dst))


def main() -> None:
    moves = [
        (ROOT / "Итог.docx", ARCHIVE / "Корень_проекта" / "Итог.docx"),
        (ROOT / "Итог.md", ARCHIVE / "Корень_проекта" / "Итог.md"),
        (ROOT / "Корректировка.docx", ARCHIVE / "Корень_проекта" / "Корректировка.docx"),
        (ROOT / "itog_utf8_check.txt", ARCHIVE / "Служебные_проверки" / "itog_utf8_check.txt"),
        (ROOT / "output" / "Пересчет_смысловых_кодов", ARCHIVE / "Старые_структуры_output" / "Пересчет_смысловых_кодов"),
        (ROOT / "output" / "Контент-анализ_по_ТЗ", ARCHIVE / "Старые_структуры_output" / "Контент-анализ_по_ТЗ"),
    ]
    for src, dst in moves:
        move_if_exists(src, dst)

    print(f"Архив подготовлен: {ARCHIVE}")


if __name__ == "__main__":
    main()
