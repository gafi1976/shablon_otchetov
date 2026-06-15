# generator_spisan.py — Генератор Word отчётов для СПИСАНИЯ оборудования

from docx import Document
from docx.shared import Pt, RGBColor, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from datetime import datetime
import os
import copy


# ─────────────────────────────────────────────
#  Вспомогательные функции форматирования
# ─────────────────────────────────────────────

def set_cell_border(cell, **kwargs):
    """Устанавливает рамку ячейки таблицы."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        tag = f'w:{edge}'
        border = OxmlElement(tag)
        val = kwargs.get(edge, 'single')
        border.set(qn('w:val'), val)
        border.set(qn('w:sz'), '6')
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), '2E4057')
        tcBorders.append(border)
    tcPr.append(tcBorders)


def set_cell_bg(cell, hex_color):
    """Устанавливает фоновый цвет ячейки."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)


def add_paragraph_with_style(doc, text, bold=False, size=12, align=WD_ALIGN_PARAGRAPH.LEFT,
                              color=None, space_before=0, space_after=6, italic=False):
    """Добавляет абзац с заданными стилями."""
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = Pt(14)
    run = p.add_run(text)
    run.bold = bold
    run.italic = italic
    run.font.size = Pt(size)
    run.font.name = 'Times New Roman'
    if color:
        run.font.color.rgb = RGBColor(*color)
    return p


def set_doc_margins(doc, top=2.0, bottom=2.0, left=2.5, right=1.5):
    """Устанавливает поля документа."""
    section = doc.sections[0]
    section.top_margin = Cm(top)
    section.bottom_margin = Cm(bottom)
    section.left_margin = Cm(left)
    section.right_margin = Cm(right)


# ─────────────────────────────────────────────
#  Генерация одного документа СПИСАНИЯ
# ─────────────────────────────────────────────

def create_spisan_doc(data: dict) -> Document:
    """
    Создаёт Word документ для акта списания.

    data = {
        'org_name': str,         # Название организации
        'region': str,           # Регион
        'commission_head': str,  # Председатель комиссии
        'member1': str,          # Член комиссии 1
        'member2': str,          # Член комиссии 2
        'doc_number': str,       # Номер акта
        'doc_date': str,         # Дата (дд.мм.гггг)
        'items': [               # Список оборудования
            {
                'num': int,
                'name': str,          # Наименование
                'inv_number': str,    # Инвентарный номер
                'year': str,          # Год приобретения
                'initial_cost': str,  # Первоначальная стоимость
                'residual_cost': str, # Остаточная стоимость
                'reason': str,        # Причина списания
                'note': str,          # Примечание
            },
            ...
        ]
    }
    """
    doc = Document()
    set_doc_margins(doc)

    # ── Шрифт по умолчанию ──
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(12)

    # ══════════════════════════════════════════
    #  ШАПКА ДОКУМЕНТА
    # ══════════════════════════════════════════

    # Гриф утверждения (правый верхний угол)
    p_approve = doc.add_paragraph()
    p_approve.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_approve.paragraph_format.space_before = Pt(0)
    p_approve.paragraph_format.space_after = Pt(0)

    r = p_approve.add_run('ТАСДИҚЛАНДИ / УТВЕРЖДАЮ\n')
    r.bold = True
    r.font.size = Pt(11)
    r.font.name = 'Times New Roman'

    r2 = p_approve.add_run(f"{data.get('org_name', '')}\n")
    r2.font.size = Pt(11)
    r2.font.name = 'Times New Roman'

    r3 = p_approve.add_run('Раҳбар / Руководитель ____________\n')
    r3.font.size = Pt(11)
    r3.font.name = 'Times New Roman'

    r4 = p_approve.add_run(f"{data.get('commission_head', '')}\n")
    r4.bold = True
    r4.font.size = Pt(11)
    r4.font.name = 'Times New Roman'

    r5 = p_approve.add_run(f"\"___\" ____________ {datetime.now().year} й.")
    r5.font.size = Pt(11)
    r5.font.name = 'Times New Roman'

    doc.add_paragraph()

    # ── Горизонтальная линия ──
    p_line = doc.add_paragraph()
    p_line.paragraph_format.space_before = Pt(0)
    p_line.paragraph_format.space_after = Pt(4)
    pPr = p_line._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '12')
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), '2E4057')
    pBdr.append(bottom)
    pPr.append(pBdr)

    # ── Заголовок ──
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_p.paragraph_format.space_before = Pt(8)
    title_p.paragraph_format.space_after = Pt(4)
    tr = title_p.add_run('АКТ СПИСАНИЯ\nОСНОВНЫХ СРЕДСТВ')
    tr.bold = True
    tr.font.size = Pt(16)
    tr.font.name = 'Times New Roman'
    tr.font.color.rgb = RGBColor(0x2E, 0x40, 0x57)

    # ── Номер и дата ──
    num_p = doc.add_paragraph()
    num_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    num_p.paragraph_format.space_before = Pt(2)
    num_p.paragraph_format.space_after = Pt(8)
    nr = num_p.add_run(f"№ {data.get('doc_number', '__')}     от  {data.get('doc_date', '__.__.__')}")
    nr.font.size = Pt(12)
    nr.font.name = 'Times New Roman'
    nr.bold = True

    # ── Организация и регион ──
    org_p = doc.add_paragraph()
    org_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    org_p.paragraph_format.space_before = Pt(0)
    org_p.paragraph_format.space_after = Pt(10)
    or_ = org_p.add_run(f"{data.get('org_name', '')}  |  {data.get('region', '')}")
    or_.font.size = Pt(11)
    or_.font.name = 'Times New Roman'
    or_.italic = True

    # ══════════════════════════════════════════
    #  СОСТАВ КОМИССИИ
    # ══════════════════════════════════════════
    add_paragraph_with_style(doc, 'СОСТАВ КОМИССИИ:', bold=True, size=12,
                             color=(0x2E, 0x40, 0x57), space_before=4, space_after=3)

    members = [
        ('Председатель комиссии:', data.get('commission_head', '—')),
        ('Член комиссии:', data.get('member1', '—')),
        ('Член комиссии:', data.get('member2', '—')),
    ]
    for label, name in members:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = Pt(14)
        rl = p.add_run(f'  {label}  ')
        rl.font.size = Pt(11)
        rl.font.name = 'Times New Roman'
        rn = p.add_run(name)
        rn.bold = True
        rn.font.size = Pt(11)
        rn.font.name = 'Times New Roman'

    # ══════════════════════════════════════════
    #  ВВОДНЫЙ ТЕКСТ
    # ══════════════════════════════════════════
    doc.add_paragraph()
    intro_p = doc.add_paragraph()
    intro_p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    intro_p.paragraph_format.space_before = Pt(0)
    intro_p.paragraph_format.space_after = Pt(8)
    intro_p.paragraph_format.line_spacing = Pt(16)
    ir = intro_p.add_run(
        f'\tНастоящий акт составлен комиссией в составе: Председатель — {data.get("commission_head", "___")}, '
        f'члены комиссии: {data.get("member1", "___")}, {data.get("member2", "___")}, '
        f'в том, что нижеперечисленное оборудование подлежит списанию с баланса организации '
        f'{data.get("org_name", "___")} по следующим основаниям:'
    )
    ir.font.size = Pt(12)
    ir.font.name = 'Times New Roman'

    # ══════════════════════════════════════════
    #  ТАБЛИЦА ОБОРУДОВАНИЯ
    # ══════════════════════════════════════════
    items = data.get('items', [])
    col_widths = [Cm(0.9), Cm(4.5), Cm(2.8), Cm(2.0), Cm(2.5), Cm(2.5), Cm(3.8), Cm(2.0)]
    headers = ['№', 'Наименование\nоборудования', 'Инвентарный\nномер',
               'Год\nввода', 'Первонач.\nстоимость\n(сум)', 'Остат.\nстоимость\n(сум)',
               'Причина\nсписания', 'Примечание']

    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = 'Table Grid'

    # Заголовок таблицы
    hdr_row = table.rows[0]
    for i, (cell, width, hdr) in enumerate(zip(hdr_row.cells, col_widths, headers)):
        cell.width = width
        set_cell_bg(cell, '2E4057')
        set_cell_border(cell)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = Pt(13)
        run = p.add_run(hdr)
        run.bold = True
        run.font.size = Pt(9)
        run.font.name = 'Times New Roman'
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    # Строки данных
    for i, item in enumerate(items):
        row = table.add_row()
        bg = 'F0F4F8' if i % 2 == 0 else 'FFFFFF'
        values = [
            str(item.get('num', i + 1)),
            item.get('name', ''),
            item.get('inv_number', ''),
            item.get('year', ''),
            item.get('initial_cost', ''),
            item.get('residual_cost', ''),
            item.get('reason', ''),
            item.get('note', ''),
        ]
        for j, (cell, val) in enumerate(zip(row.cells, values)):
            set_cell_bg(cell, bg)
            set_cell_border(cell)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if j in (0, 2, 3, 4, 5) else WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.line_spacing = Pt(13)
            run = p.add_run(val)
            run.font.size = Pt(10)
            run.font.name = 'Times New Roman'
            # Подсветка красным если причина содержит "эскирган" или "списание"
            if j == 6 and val:
                run.font.color.rgb = RGBColor(0xC0, 0x39, 0x2B)

    doc.add_paragraph()

    # ══════════════════════════════════════════
    #  ЗАКЛЮЧЕНИЕ
    # ══════════════════════════════════════════
    add_paragraph_with_style(doc, 'ЗАКЛЮЧЕНИЕ КОМИССИИ:', bold=True, size=12,
                             color=(0x2E, 0x40, 0x57), space_before=6, space_after=3)

    conclusion_p = doc.add_paragraph()
    conclusion_p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    conclusion_p.paragraph_format.space_before = Pt(0)
    conclusion_p.paragraph_format.space_after = Pt(6)
    conclusion_p.paragraph_format.line_spacing = Pt(16)
    cr = conclusion_p.add_run(
        f'\tКомиссия установила, что перечисленные выше основные средства в количестве '
        f'{len(items)} ({"  ".join([item.get("name", "") for item in items[:2]])}'
        f'{"и др." if len(items) > 2 else ""}) морально и физически устарели, '
        f'восстановление нецелесообразно. Комиссия рекомендует произвести списание '
        f'указанного имущества с баланса {data.get("org_name", "___")}.'
    )
    cr.font.size = Pt(12)
    cr.font.name = 'Times New Roman'

    doc.add_paragraph()
    doc.add_paragraph()

    # ══════════════════════════════════════════
    #  ПОДПИСИ
    # ══════════════════════════════════════════
    add_paragraph_with_style(doc, 'ПОДПИСИ ЧЛЕНОВ КОМИССИИ:', bold=True, size=11,
                             color=(0x2E, 0x40, 0x57), space_before=4, space_after=6)

    sig_table = doc.add_table(rows=3, cols=3)
    sig_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    sig_labels = [
        ('Председатель:', data.get('commission_head', '_______________')),
        ('Член комиссии:', data.get('member1', '_______________')),
        ('Член комиссии:', data.get('member2', '_______________')),
    ]
    for i, (label, name) in enumerate(sig_labels):
        row = sig_table.rows[i]
        row.cells[0].width = Cm(4)
        row.cells[1].width = Cm(6)
        row.cells[2].width = Cm(6)

        p0 = row.cells[0].paragraphs[0]
        p0.paragraph_format.space_before = Pt(0)
        p0.paragraph_format.space_after = Pt(6)
        r0 = p0.add_run(label)
        r0.bold = True
        r0.font.size = Pt(11)
        r0.font.name = 'Times New Roman'

        p1 = row.cells[1].paragraphs[0]
        p1.paragraph_format.space_before = Pt(0)
        p1.paragraph_format.space_after = Pt(6)
        r1 = p1.add_run('_____________________')
        r1.font.size = Pt(11)
        r1.font.name = 'Times New Roman'

        p2 = row.cells[2].paragraphs[0]
        p2.paragraph_format.space_before = Pt(0)
        p2.paragraph_format.space_after = Pt(6)
        r2 = p2.add_run(name)
        r2.font.size = Pt(11)
        r2.font.name = 'Times New Roman'

    # ── Нижняя линия ──
    doc.add_paragraph()
    p_footer = doc.add_paragraph()
    p_footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_footer.paragraph_format.space_before = Pt(10)
    p_footer.paragraph_format.space_after = Pt(0)
    pPr2 = p_footer._p.get_or_add_pPr()
    pBdr2 = OxmlElement('w:pBdr')
    top2 = OxmlElement('w:top')
    top2.set(qn('w:val'), 'single')
    top2.set(qn('w:sz'), '12')
    top2.set(qn('w:space'), '1')
    top2.set(qn('w:color'), '2E4057')
    pBdr2.append(top2)
    pPr2.append(pBdr2)
    fr = p_footer.add_run(
        f"М.П.   Документ составлен: {data.get('doc_date', datetime.now().strftime('%d.%m.%Y'))}"
    )
    fr.font.size = Pt(9)
    fr.font.name = 'Times New Roman'
    fr.italic = True
    fr.font.color.rgb = RGBColor(0x88, 0x88, 0x88)

    return doc


# ─────────────────────────────────────────────
#  Сохранение — один файл на документ
# ─────────────────────────────────────────────

def save_single_files(data_list: list, output_dir: str) -> list:
    """
    Сохраняет каждый акт списания в отдельный файл.
    Возвращает список созданных файлов.
    """
    os.makedirs(output_dir, exist_ok=True)
    created = []
    for data in data_list:
        doc = create_spisan_doc(data)
        safe = data.get('doc_number', 'N').replace('/', '_').replace('\\', '_')
        name = f"Akt_Spisaniya_{safe}.docx"
        path = os.path.join(output_dir, name)
        doc.save(path)
        created.append(path)
    return created


# ─────────────────────────────────────────────
#  Сохранение — все акты в одном файле
# ─────────────────────────────────────────────

def save_all_in_one(data_list: list, output_path: str) -> str:
    """
    Объединяет все акты списания в один документ.
    Возвращает путь к файлу.
    """
    from docx.oxml import OxmlElement as OE
    combined = Document()
    set_doc_margins(combined)
    combined.styles['Normal'].font.name = 'Times New Roman'
    combined.styles['Normal'].font.size = Pt(12)

    for idx, data in enumerate(data_list):
        single_doc = create_spisan_doc(data)

        # Добавляем все параграфы и таблицы из single_doc
        for element in single_doc.element.body:
            combined.element.body.append(copy.deepcopy(element))

        # Разрыв страницы между актами (кроме последнего)
        if idx < len(data_list) - 1:
            pb_p = OE('w:p')
            pb_r = OE('w:r')
            pb_br = OE('w:br')
            pb_br.set(qn('w:type'), 'page')
            pb_r.append(pb_br)
            pb_p.append(pb_r)
            combined.element.body.append(pb_p)

    combined.save(output_path)
    return output_path
