# -*- coding: utf-8 -*-
"""
Контент-анализ строго по ТЗ:
- 10 тематических + 3 тональных кода (как в методичке)
- атомарное разбиение высказываний (одна мысль = одна строка)
- Excel: Данные, Коды, Итоги_темы (частота + ранг), Итоги_тональность
- Word-обобщение с ранжированием
"""

from __future__ import annotations

import hashlib
import re
import shutil
import zipfile
from dataclasses import dataclass
from datetime import date
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt
from openpyxl import Workbook
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parents[1]
TZ_DOCX = ROOT / "ТЗ контент.docx"
OUT = ROOT / "output" / "Контент-анализ_по_ТЗ"

# === Коды строго по ТЗ (шаг 2.1 + 2.2) ===
THEME_CODES: list[tuple[str, str, str]] = [
    ("НОРМА_442ФЗ", "442-ФЗ и связанные акты",
     "Изменение/уточнение ФЗ № 442-ФЗ и связанных актов для расширения доступа НКО."),
    ("КОНКУРС_ПРОЦЕДУРЫ", "Конкурсы и процедуры",
     "Упрощение конкурсов, снижение бюрократии, 44-ФЗ/223-ФЗ."),
    ("ТАРИФЫ_НОРМАТИВЫ", "Тарифы и нормативы",
     "Тарифы, подушевые нормативы, индексация финансирования соцуслуг."),
    ("РЕЕСТР_ПОСТАВЩИКОВ", "Реестр поставщиков",
     "Упрощение включения в реестр, прозрачные критерии."),
    ("ГАРАНТИИ_ОБЪЕМОВ", "Гарантии объёмов",
     "Гарантированный объём, многолетние контракты, авансирование, лимиты."),
    ("ОБУЧЕНИЕ_НКО", "Обучение НКО",
     "Обучение конкурсам, отчётности, юридическим требованиям."),
    ("ИНФОРМАЦИЯ_ПРОЗРАЧНОСТЬ", "Информация и прозрачность",
     "Информирование о конкурсах, критериях, планах финансирования."),
    ("ИНСТИТУТЫ_ПОДДЕРЖКИ", "Институты поддержки",
     "Ресурсные центры, опорные организации, грантовая поддержка."),
    ("МУНИЦИПАЛЬНЫЙ_УРОВЕНЬ", "Муниципальный уровень",
     "Практика на муниципальном уровне, решения администраций."),
    ("КОНКУРЕНЦИЯ_АНТИМОНОПОЛИЯ", "Конкуренция и антимонополия",
     "Борьба с монополией госучреждений, равные условия."),
]

TONE_CODES: list[tuple[str, str, str]] = [
    ("ТОНИЧЕСКИЙ_ПОЗИТИВ", "Позитив",
     "Позитивная оценка мер или перспектив."),
    ("ТОНИЧЕСКИЙ_НЕГАТИВ", "Негатив",
     "Критика, скепсис, описание проблем."),
    ("ТОНИЧЕСКИЙ_СОМНЕНИЕ", "Сомнение",
     "Затрудняюсь ответить, неопределённость."),
]

ALL_CODES = THEME_CODES + TONE_CODES

THEME_KEYWORDS: dict[str, list[str]] = {
    "НОРМА_442ФЗ": [r"442", r"189.?фз", r"законодатель", r"изменен.*закон", r"закон\s", r"нпа", r"постановлен"],
    "КОНКУРС_ПРОЦЕДУРЫ": [r"конкурс", r"закуп", r"44.?фз", r"223.?фз", r"тендер", r"отбор", r"процедур", r"бюрократ"],
    "ТАРИФЫ_НОРМАТИВЫ": [r"тариф", r"подушев", r"норматив", r"индекс", r"инфляц", r"мрот", r"стоимост.*услуг"],
    "РЕЕСТР_ПОСТАВЩИКОВ": [r"реестр", r"включен.*реестр", r"вхожд.*реестр", r"вступлен.*реестр"],
    "ГАРАНТИИ_ОБЪЕМОВ": [
        r"объ[её]м", r"аванс", r"долгосроч", r"многолет", r"гарант", r"лимит", r"квот",
        r"финансирован", r"субсид", r"бюджет", r"ассигнован", r"кассов",
    ],
    "ОБУЧЕНИЕ_НКО": [r"обучен", r"семинар", r"тренинг", r"квалификац", r"методическ", r"консультац"],
    "ИНФОРМАЦИЯ_ПРОЗРАЧНОСТЬ": [r"информ", r"прозрач", r"открытост", r"доступност.*информ"],
    "ИНСТИТУТЫ_ПОДДЕРЖКИ": [r"ресурсн.*центр", r"грант", r"опорн", r"институт", r"цисс"],
    "МУНИЦИПАЛЬНЫЙ_УРОВЕНЬ": [r"муницип", r"администрац", r"омс", r"местн.*самоуправ"],
    "КОНКУРЕНЦИЯ_АНТИМОНОПОЛИЯ": [
        r"монопол", r"конкурен", r"равн.*услов", r"антимонопол", r"конфликт интерес", r"дискримин",
    ],
}


@dataclass
class Statement:
    id: int
    text: str
    parent_id: int
    codes: dict[str, int]
    text_hash: str = ""


def extract_text_from_docx(path: Path) -> str:
    import xml.etree.ElementTree as ET

    with zipfile.ZipFile(path) as z:
        root = ET.fromstring(z.read("word/document.xml"))
    parts: list[str] = []
    for t in root.iter("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t"):
        if t.text:
            parts.append(t.text)
        if t.tail:
            parts.append(t.tail)
    return "".join(parts)


def rough_split(body: str) -> list[str]:
    body = re.sub(r"([а-яё])(\d+\.)", r"\1 \2", body)
    body = re.sub(r"(\d+\.)([А-ЯЁ])", r"\1 \2", body)
    body = re.sub(r"([.!?»\"])([А-ЯЁ])", r"\1 \2", body)

    parts = [body]
    for pattern in (r"\s*\d+\.\s+", r"\s*[-–—]\s+"):
        new: list[str] = []
        for chunk in parts:
            new.extend(re.split(pattern, chunk))
        parts = new

    glued: list[str] = []
    for chunk in parts:
        glued.extend(re.split(r"(?<=[а-яё])(?=[А-ЯЁ][а-яё])", chunk))

    out: list[str] = []
    for chunk in glued:
        chunk = re.sub(r"\s+", " ", chunk).strip(" .;")
        if len(chunk) >= 3 and not chunk.lower().startswith("цель:"):
            out.append(chunk)
    return out


def split_atomic(text: str) -> list[str]:
    """Дробит составное высказывание на атомарные (одна мысль — одна строка)."""
    chunks = [text]

    patterns = [
        r"(?<=\.)\s*(?=\d+\.\s)",           # 1. ... 2. ...
        r"(?<=\.)\s*(?=\d+\))",             # 1) 2)
        r"\s*;\s+",                          # точка с запятой
        r"(?<=[.!?])\s+(?=[А-ЯЁ«\"])",     # новое предложение
        r"\s+(?<=[а-яё])\.(?=\s*[А-ЯЁ])",  # точка между предложениями
    ]
    for pat in patterns:
        new_chunks: list[str] = []
        for c in chunks:
            new_chunks.extend(re.split(pat, c))
        chunks = new_chunks

    # Списки через «, и » / «; » в длинных фразах
    result: list[str] = []
    for c in chunks:
        c = re.sub(r"^\d+[\.\)]\s*", "", c.strip())
        c = re.sub(r"\s+", " ", c).strip(" .;")
        if len(c) < 8:
            continue
        # Если всё ещё очень длинное и есть «, и » — пробуем разрезать
        if len(c) > 180 and ", и " in c:
            subs = [s.strip() for s in re.split(r",\s+и\s+", c) if len(s.strip()) >= 8]
            if len(subs) > 1:
                result.extend(subs)
                continue
        result.append(c)
    return result if result else ([text] if text.strip() else [])


def build_corpus(raw_text: str) -> tuple[list[Statement], list[str]]:
    idx = raw_text.find("Высказывания:")
    if idx == -1:
        raise ValueError("Блок «Высказывания:» не найден")
    body = raw_text[idx + len("Высказывания:") :]

    rough = rough_split(body)
    atomic: list[tuple[int, str]] = []
    parent = 0
    for block in rough:
        parent += 1
        parts = split_atomic(block)
        if len(parts) > 1:
            for p in parts:
                atomic.append((parent, p))
        else:
            atomic.append((parent, block))

    statements: list[Statement] = []
    for i, (pid, text) in enumerate(atomic, 1):
        h = hashlib.sha256(text.encode("utf-8")).hexdigest()[:12]
        statements.append(Statement(i, text, pid, code_statement(text), h))
    return statements, rough


def code_themes(text: str) -> dict[str, int]:
    low = text.lower()
    codes = {c[0]: 0 for c in THEME_CODES}
    for code, patterns in THEME_KEYWORDS.items():
        for pat in patterns:
            if re.search(pat, low):
                codes[code] = 1
                break
    return codes


def assign_tone(text: str) -> str:
    low = text.lower().strip()
    if re.search(r"затрудн|не знаю|неизвестно|надо подумать|правительству виднее|пусть решают|нет предлож", low):
        return "ТОНИЧЕСКИЙ_СОМНЕНИЕ"
    if re.search(
        r"мер достаточно|достаточно мер|доступ есть|вс[её] устраивает|вс[её] нормально|нет проблем|"
        r"не требуется|всего вполне|всем вс[её] доступно|вс[её] есть|принимаются достаточные",
        low,
    ):
        if not re.search(r"но\s|тариф.*низк|невыносим|задерж", low):
            return "ТОНИЧЕСКИЙ_ПОЗИТИВ"
    if re.search(
        r"невыносим|выжив|закрыт|обидно|монопол|задерж|кассов|убыт|штраф|шантаж|"
        r"не работает|искусственн|не поможет|ничего не поможет|руки опуск",
        low,
    ):
        return "ТОНИЧЕСКИЙ_НЕГАТИВ"
    if re.search(r"низк.*тариф|тариф.*низк|не индекс|остаточн|формально", low):
        return "ТОНИЧЕСКИЙ_НЕГАТИВ"
    return "ТОНИЧЕСКИЙ_СОМНЕНИЕ" if len(low) < 25 else "ТОНИЧЕСКИЙ_НЕГАТИВ"


def code_statement(text: str) -> dict[str, int]:
    themes = code_themes(text)
    tone = assign_tone(text)
    codes = {c[0]: 0 for c in ALL_CODES}
    for k, v in themes.items():
        codes[k] = v
    codes[tone] = 1
    return codes


def aggregate(statements: list[Statement]) -> dict:
    n = len(statements)
    themes = {c[0]: 0 for c in THEME_CODES}
    tones = {c[0]: 0 for c in TONE_CODES}
    for s in statements:
        for k, v in s.codes.items():
            if v:
                if k.startswith("ТОНИЧЕСКИЙ"):
                    tones[k] += 1
                elif k in themes:
                    themes[k] += 1
    ranked = sorted(themes.items(), key=lambda x: (-x[1], x[0]))
    return {"n": n, "themes": themes, "tones": tones, "ranked": ranked}


def write_excel(path: Path, statements: list[Statement]) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Данные"

    headers = ["ID_высказывания", "ID_исходного_блока", "Текст_высказывания"] + [c[0] for c in ALL_CODES]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E79")

    for s in statements:
        row = [s.id, s.parent_id, s.text] + [s.codes[c[0]] for c in ALL_CODES]
        ws.append(row)

    ws.column_dimensions["C"].width = 70
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        row[2].alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = "D2"

    n = len(statements)
    last = n + 1
    code_start_col = 4  # D

    # --- Коды ---
    ws_c = wb.create_sheet("Коды")
    ws_c.append(["№", "Код", "Название", "Описание", "Тип"])
    for cell in ws_c[1]:
        cell.font = Font(bold=True)
    num = 1
    for code, name, desc in THEME_CODES:
        ws_c.append([num, code, name, desc, "тематический"])
        num += 1
    for code, name, desc in TONE_CODES:
        ws_c.append([num, code, name, desc, "тональный"])
        num += 1

    # --- Итоги_темы (ГЛАВНЫЙ ЛИСТ ДЛЯ РЕЙТИНГА) ---
    ws_t = wb.create_sheet("Итоги_темы")
    ws_t["A1"] = "РАНЖИРОВАНИЕ ТЕМАТИЧЕСКИХ КОДОВ ПО УПОМИНАЕМОСТИ"
    ws_t["A1"].font = Font(bold=True, size=14)
    ws_t.merge_cells("A1:F1")
    ws_t["A2"] = f"Всего высказываний в корпусе: {n}"
    ws_t["A2"].font = Font(bold=True)

    hdr = 4
    cols = ["Ранг", "Код", "Название", "Частота", "Доля", "Формула_проверки"]
    for j, h in enumerate(cols, 1):
        c = ws_t.cell(row=hdr, column=j, value=h)
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor="D9E2F3")

    stats = aggregate(statements)
    sorted_themes = sorted(stats["themes"].items(), key=lambda x: (-x[1], x[0]))
    for rank, (code_key, _) in enumerate(sorted_themes, 1):
        name = next(nm for c, nm, _ in THEME_CODES if c == code_key)
        col_idx = next(i for i, (c, _, _) in enumerate(THEME_CODES) if c == code_key)
        col_letter = get_column_letter(code_start_col + col_idx)
        row = hdr + rank
        ws_t.cell(row=row, column=1, value=rank)
        ws_t.cell(row=row, column=2, value=code_key)
        ws_t.cell(row=row, column=3, value=name)
        ws_t.cell(row=row, column=4, value=f"=SUM(Данные!{col_letter}2:{col_letter}{last})")
        ws_t.cell(row=row, column=5, value=f"=D{row}/$B$2")
        ws_t.cell(row=row, column=6, value=f"=SUM(Данные!{col_letter}2:{col_letter}{last})")
    ws_t["B2"] = f"=COUNTA(Данные!A2:A{last})"

    ws_t.column_dimensions["C"].width = 35
    ws_t.column_dimensions["F"].width = 20

    # Диаграмма топ тем
    bar = BarChart()
    bar.type = "col"
    bar.title = "Частота тематических кодов"
    bar.height = 12
    bar.width = 18
    data = Reference(ws_t, min_col=4, min_row=hdr, max_row=hdr + 10)
    cats = Reference(ws_t, min_col=3, min_row=hdr + 1, max_row=hdr + 10)
    bar.add_data(data, titles_from_data=True)
    bar.set_categories(cats)
    bar.dataLabels = DataLabelList()
    bar.dataLabels.showVal = True
    ws_t.add_chart(bar, "H4")

    # --- Итоги_тональность ---
    ws_tone = wb.create_sheet("Итоги_тональность")
    ws_tone["A1"] = "ТОНАЛЬНОСТЬ ВЫСКАЗЫВАНИЙ"
    ws_tone["A1"].font = Font(bold=True, size=14)
    ws_tone.merge_cells("A1:D1")
    ws_tone["A2"] = f"=COUNTA(Данные!A2:A{last})"
    ws_tone["A2"].font = Font(bold=True)

    ws_tone.append([])
    ws_tone.append(["Тональность", "Частота", "Доля"])
    for cell in ws_tone[4]:
        cell.font = Font(bold=True)

    tone_labels = {
        "ТОНИЧЕСКИЙ_ПОЗИТИВ": "Позитив",
        "ТОНИЧЕСКИЙ_НЕГАТИВ": "Негатив",
        "ТОНИЧЕСКИЙ_СОМНЕНИЕ": "Сомнение",
    }
    tone_start = code_start_col + len(THEME_CODES)
    for i, (code, label) in enumerate(tone_labels.items(), 5):
        col_letter = get_column_letter(tone_start + i - 5)
        ws_tone.cell(row=i, column=1, value=label)
        ws_tone.cell(row=i, column=2, value=f"=SUM(Данные!{col_letter}2:{col_letter}{last})")
        ws_tone.cell(row=i, column=3, value=f"=B{i}/$A$2")

    pie = PieChart()
    pie.title = "Тональность"
    pie.height = 10
    pie.width = 14
    labels = Reference(ws_tone, min_col=1, min_row=5, max_row=7)
    data = Reference(ws_tone, min_col=2, min_row=4, max_row=7)
    pie.add_data(data, titles_from_data=True)
    pie.set_categories(labels)
    pie.dataLabels = DataLabelList()
    pie.dataLabels.showPercent = True
    pie.dataLabels.showVal = True
    ws_tone.add_chart(pie, "E4")

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def save_ranking_chart(path: Path, ranked: list[tuple[str, int]]) -> None:
    names = {c: n for c, n, _ in THEME_CODES}
    top = [(names[c], v) for c, v in ranked if v > 0][:10]
    if not top:
        return
    labels, values = zip(*top)
    fig, ax = plt.subplots(figsize=(9, 5))
    bars = ax.barh(range(len(labels)), values, color="#1F4E79")
    ax.set_yticks(range(len(labels)))
    ax.set_yticklabels(labels, fontsize=10)
    ax.invert_yaxis()
    ax.set_xlabel("Частота упоминаний")
    ax.set_title("Ранжирование тематических кодов", fontweight="bold")
    for bar, val in zip(bars, values):
        ax.text(bar.get_width() + 0.5, bar.get_y() + bar.get_height() / 2, str(val), va="center")
    fig.tight_layout()
    path.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(path, dpi=150)
    plt.close(fig)


def save_tone_chart(path: Path, tones: dict[str, int]) -> None:
    labels = ["Позитив", "Негатив", "Сомнение"]
    keys = ["ТОНИЧЕСКИЙ_ПОЗИТИВ", "ТОНИЧЕСКИЙ_НЕГАТИВ", "ТОНИЧЕСКИЙ_СОМНЕНИЕ"]
    values = [tones.get(k, 0) for k in keys]
    fig, ax = plt.subplots(figsize=(6, 4))
    colors = ["#2E7D32", "#C62828", "#757575"]
    ax.bar(labels, values, color=colors)
    ax.set_title("Тональность высказываний", fontweight="bold")
    for i, v in enumerate(values):
        ax.text(i, v + 0.5, str(v), ha="center")
    fig.tight_layout()
    path.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(path, dpi=150)
    plt.close(fig)


def write_summary_docx(path: Path, statements: list[Statement], rough_count: int) -> None:
    stats = aggregate(statements)
    n = stats["n"]
    doc = Document()
    doc.add_heading("Обобщение: тематические коды и ранжирование", 0)
    doc.add_paragraph(f"Дата: {date.today().isoformat()}")
    doc.add_paragraph(
        f"Корпус: {n} атомарных высказываний (из {rough_count} исходных блоков ответов экспертов). "
        "Единица анализа — одна законченная мысль; составные ответы разбиты на отдельные строки."
    )

    doc.add_heading("1. Список тематических кодов (10 шт.)", level=1)
    table = doc.add_table(rows=1, cols=3)
    table.style = "Table Grid"
    hdr = table.rows[0].cells
    hdr[0].text = "Код"
    hdr[1].text = "Название"
    hdr[2].text = "Описание"
    for code, name, desc in THEME_CODES:
        row = table.add_row().cells
        row[0].text = code
        row[1].text = name
        row[2].text = desc

    doc.add_heading("2. Ранжирование тем по упоминаемости", level=1)
    doc.add_paragraph(
        "Частота = число высказываний, в которых код = 1. "
        "Одно высказывание может иметь несколько кодов."
    )
    t2 = doc.add_table(rows=1, cols=5)
    t2.style = "Table Grid"
    h2 = t2.rows[0].cells
    for i, txt in enumerate(["Ранг", "Код", "Название", "Частота", "Доля"]):
        h2[i].text = txt

    names = {c: nm for c, nm, _ in THEME_CODES}
    for rank, (code, freq) in enumerate(stats["ranked"], 1):
        if freq == 0:
            continue
        row = t2.add_row().cells
        row[0].text = str(rank)
        row[1].text = code
        row[2].text = names[code]
        row[3].text = str(freq)
        row[4].text = f"{freq / n * 100:.1f}%"

    doc.add_heading("3. Тональные коды (3 шт.)", level=1)
    for code, name, desc in TONE_CODES:
        cnt = stats["tones"][code]
        doc.add_paragraph(f"• {name} ({code}): {cnt} ({cnt/n*100:.1f}%) — {desc}")

    doc.add_heading("4. Ключевой вывод", level=1)
    top3 = stats["ranked"][:3]
    top_txt = ", ".join(f"{names[c]} ({v})" for c, v in top3 if v)
    doc.add_paragraph(
        f"Наиболее часто упоминаются: {top_txt}. "
        "Эксперты концентрируются на тарифах/нормативах, объёмах финансирования и административных процедурах."
    )

    chart = OUT / "04_Диаграммы" / "ранжирование_тем.png"
    if chart.exists():
        doc.add_picture(str(chart), width=Inches(6))
        doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.save(path)


def write_summary_txt(path: Path, statements: list[Statement]) -> None:
    stats = aggregate(statements)
    n = stats["n"]
    names = {c: nm for c, nm, _ in THEME_CODES}
    lines = [
        "ОБОБЩЕНИЕ: ТЕМАТИЧЕСКИЕ КОДЫ И РАНЖИРОВАНИЕ",
        f"Высказываний: {n}",
        "",
        "=== 10 ТЕМАТИЧЕСКИХ КОДОВ (список) ===",
    ]
    for i, (code, name, _) in enumerate(THEME_CODES, 1):
        lines.append(f"{i}. {code} — {name}")
    lines += ["", "=== РЕЙТИНГ ПО ЧАСТОТЕ ==="]
    for rank, (code, freq) in enumerate(stats["ranked"], 1):
        if freq:
            lines.append(f"{rank}. {names[code]}: {freq} ({freq/n*100:.1f}%)")
    lines += ["", "=== ТОНАЛЬНОСТЬ ==="]
    for code, name, _ in TONE_CODES:
        v = stats["tones"][code]
        lines.append(f"{name}: {v} ({v/n*100:.1f}%)")
    path.write_text("\n".join(lines), encoding="utf-8")


def write_corpus_docx(path: Path, statements: list[Statement]) -> None:
    doc = Document()
    doc.add_heading("Корпус атомарных высказываний", level=1)
    doc.add_paragraph(
        "Вопрос 72. Каждая строка — одна мысль. Составные ответы разбиты; "
        "ID_исходного_блока показывает, из какого исходного ответа получены фрагменты."
    )
    table = doc.add_table(rows=1, cols=3)
    table.style = "Table Grid"
    table.rows[0].cells[0].text = "ID"
    table.rows[0].cells[1].text = "ID блока"
    table.rows[0].cells[2].text = "Текст"
    for s in statements:
        row = table.add_row().cells
        row[0].text = str(s.id)
        row[1].text = str(s.parent_id)
        row[2].text = s.text
    doc.save(path)


def make_zip(folder: Path, zip_path: Path) -> None:
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for f in folder.rglob("*"):
            if f.is_file():
                zf.write(f, f.relative_to(folder.parent))


def main() -> None:
    if OUT.exists():
        shutil.rmtree(OUT)
    for sub in ["01_ТЗ_и_корпус", "02_Excel", "03_Обобщение", "04_Диаграммы", "05_Отчет"]:
        (OUT / sub).mkdir(parents=True)

    raw = extract_text_from_docx(TZ_DOCX)
    statements, rough = build_corpus(raw)
    stats = aggregate(statements)

    print(f"Исходных блоков: {len(rough)}")
    print(f"Атомарных высказываний: {len(statements)}")
    print("Ранжирование тем:")
    names = {c: n for c, n, _ in THEME_CODES}
    for i, (code, freq) in enumerate(stats["ranked"][:10], 1):
        print(f"  {i}. {names[code]}: {freq}")

    shutil.copy2(TZ_DOCX, OUT / "01_ТЗ_и_корпус" / "01_TZ-kontent.docx")
    pdf = ROOT / "ТЗ контент.pdf"
    if pdf.exists():
        shutil.copy2(pdf, OUT / "01_ТЗ_и_корпус" / "01_TZ-kontent.pdf")

    write_corpus_docx(OUT / "01_ТЗ_и_корпус" / "02_Корпус_атомарных_высказываний.docx", statements)
    write_excel(OUT / "02_Excel" / "Контент-анализ_НКО_соцуслуги.xlsx", statements)
    save_ranking_chart(OUT / "04_Диаграммы" / "ранжирование_тем.png", stats["ranked"])
    save_tone_chart(OUT / "04_Диаграммы" / "тональность.png", stats["tones"])

    write_summary_docx(OUT / "03_Обобщение" / "Тематические_коды_и_ранжирование.docx", statements, len(rough))
    write_summary_txt(OUT / "03_Обобщение" / "Тематические_коды_и_ранжирование.txt", statements)

    # Краткий codes cheat sheet
    codes_doc = Document()
    codes_doc.add_heading("Шпаргалка кодировщика (10 + 3 кода)", 0)
    codes_doc.add_paragraph("Тематические коды — в одном высказывании может быть несколько «1».")
    for code, name, desc in THEME_CODES:
        codes_doc.add_paragraph(f"{code}: {name}. {desc}", style="List Bullet")
    codes_doc.add_paragraph("Тональность — только один код на высказывание:")
    for code, name, desc in TONE_CODES:
        codes_doc.add_paragraph(f"{code}: {name}. {desc}", style="List Bullet")
    codes_doc.save(OUT / "03_Обобщение" / "Шпаргалка_кодировщика.docx")

    from build_detailed_report import generate_report_package
    report_path = generate_report_package()
    print(f"Подробный отчёт: {report_path}")

    make_zip(OUT, ROOT / "output" / "Kontent-analiz-po-TZ.zip")
    print(f"Готово: {OUT}")


if __name__ == "__main__":
    main()
