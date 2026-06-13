from __future__ import annotations

import shutil
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
ARCHIVE = ROOT / "Архив" / "Старые_скрипты"
SCRIPTS = ROOT / "scripts"


def move_if_exists(src: Path, dst: Path) -> None:
    if not src.exists():
        return
    dst.parent.mkdir(parents=True, exist_ok=True)
    if dst.exists():
        dst.unlink()
    shutil.move(str(src), str(dst))


def main() -> None:
    files = [
        "build_content_analysis.py",
        "build_detailed_report.py",
        "md_to_docx.py",
        "update_itog_with_semantic_table.py",
    ]
    for name in files:
        move_if_exists(SCRIPTS / name, ARCHIVE / name)
    print(f"Старые скрипты перенесены в: {ARCHIVE}")


if __name__ == "__main__":
    main()
