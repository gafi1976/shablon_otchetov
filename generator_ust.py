# generator_ust.py — Генератор Word отчётов для УСТАНОВКИ оборудования

from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from datetime import datetime
import os
import copy


# ─────────────────────────────────────────────
#  Вспомогательные функции (общие)
# ─────────────────────────────────────────────

def set_cell_border(cell, color='1A5276'):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for edge in ('top', 'left', 'bottom', 'right'):
        tag = f'w:{edge}'
        border = OxmlElement(tag)
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), '6')
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), color)
        tcBorders.append(border)
    tcPr.append(tcBorders)


def set_cell_bg(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)


def set_doc_margins(doc, top=2.0, bottom=2.0, left=2.5, right=1.5):
    section = doc.sections[0]
    section.top_margin = Cm(top)
    section.bottom_margin = Cm(bottom)
    section.left_margin = Cm(left)
    section.right_margin = Cm(right)


def add_horizontal_line(doc, color='1A5276', thickness='12', position='bottom'):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(4)
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    edge = OxmlElement(f'w:{position}')
    edge.set(qn('w:val'), 'single')
    edge.set(qn('w:sz'), thickness)
    edge.set(qn('w:space'), '1')
    edge.set(qn('w:color'), color)
    pBdr.append(edge)
    pPr.append(pBdr)
    return p


# ─────────────────────────────────────────────
#  Генерация одного документа УСТАНОВКИ
# ─────────────────────────────────────────────

def create_ust_doc(data: dict) -> Document:
    """
    Создаёт Word документ акта ввода в эксплуатацию (установки).

    data = {
        'org_name': str,            # Название организации
        'region': str,              # Регион
        'commission_head': str,     # Председатель комиссии
        'member1': str,             # Член комиссии 1
        'member2': str,             # Член комиссии 2
        'doc_number': str,          # Номер акта
        'doc_date': str,            # Дата (дд.мм.гггг)
        'location': str,            # Место установки
        'items': [
            {
                'num': int,
                'name': str,              # Наименование оборудования
                'model': str,             # Модель / марка
                'serial_number': str,     # Серийный номер
                'inv_number': str,        # Инвентарный номер
                'year': str,              # Год выпуска
                'cost': str,              # Стоимость (сум)
                'install_date': str,      # Дата установки
                'condition': str,         # Состояние (Янги / Ishlatilgan)
                'note': str,              # Примечание
            },
            ...
        ]
    }
    """
    doc = Document()
    set_doc_margins(doc)

    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(12)

    # ══════════════════════════════════════════
    #  ШАПКА — ГРИФ УТВЕРЖДЕНИЯ
    # ══════════════════════════════════════════
    p_approve = doc.add_paragraph()
    p_approve.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_approve.paragraph_format.space_before = Pt(0)
    p_approve.paragraph_format.space_after = Pt(0)

    runs = [
        ('ТАСДИҚЛАНДИ / УТВЕРЖДАЮ\n', True, 11, None),
        (f"{data.get('org_name', '')}\n", False, 11, None),
        ('Раҳбар / Руководитель ____________\n', False, 11, None),
        (f"{data.get('commission_head', '')}\n", True, 11, None),
        (f'"___" ____________ {datetime.now().year} й.', False, 11, None),
    ]
    for text, bold, size, color in runs:
        r = p_approve.add_run(text)
        r.bold = bold
        r.font.size = Pt(size)
        r.font.name = 'Times New Roman'

    doc.add_paragraph()
    add_horizontal_line(doc, color='1A5276')

    # ══════════════════════════════════════════
    #  ЗАГОЛОВОК
    # ══════════════════════════════════════════
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_p.paragraph_format.space_before = Pt(8)
    title_p.paragraph_format.space_after = Pt(4)
    tr = title_p.add_run('АКТ ВВОДА В ЭКСПЛУАТАЦИЮ\n(УСТАНОВКИ ОБОРУДОВАНИЯ)')
    tr.bold = True
    tr.font.size = Pt(16)
    tr.font.name = 'Times New Roman'
    tr.font.color.rgb = RGBColor(0x1A, 0x52, 0x76)

    num_p = doc.add_paragraph()
    num_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    num_p.paragraph_format.space_before = Pt(2)
    num_p.paragraph_format.space_after = Pt(4)
    nr = num_p.add_run(f"№ {data.get('doc_number', '__')}     от  {data.get('doc_date', '__.__.__')}")
    nr.bold = True
    nr.font.size = Pt(12)
    nr.font.name = 'Times New Roman'

    org_p = doc.add_paragraph()
    org_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    org_p.paragraph_format.space_before = Pt(0)
    org_p.paragraph_format.space_after = Pt(8)
    or_ = org_p.add_run(f"{data.get('org_name', '')}  |  {data.get('region', '')}")
    or_.font.size = Pt(11)
    or_.font.name = 'Times New Roman'
    or_.italic = True

    # ── Место установки ──
    loc_p = doc.add_paragraph()
    loc_p.paragraph_format.space_before = Pt(0)
    loc_p.paragraph_format.space_after = Pt(6)
    loc_label = loc_p.add_run('Место установки:  ')
    loc_label.bold = True
    loc_label.font.size = Pt(11)
    loc_label.font.name = 'Times New Roman'
    loc_val = loc_p.add_run(data.get('location', '___________________________'))
    loc_val.font.size = Pt(11)
    loc_val.font.name = 'Times New Roman'

    # ══════════════════════════════════════════
    #  СОСТАВ КОМИССИИ
    # ══════════════════════════════════════════
    ch_p = doc.add_paragraph()
    ch_p.paragraph_format.space_before = Pt(4)
    ch_p.paragraph_format.space_after = Pt(3)
    ch_r = ch_p.add_run('СОСТАВ КОМИССИИ:')
    ch_r.bold = True
    ch_r.font.size = Pt(12)
    ch_r.font.name = 'Times New Roman'
    ch_r.font.color.rgb = RGBColor(0x1A, 0x52, 0x76)

    for label, name in [
        ('Председатель комиссии:', data.get('commission_head', '—')),
        ('Член комиссии:', data.get('member1', '—')),
        ('Член комиссии:', data.get('member2', '—')),
    ]:
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
        f'в том, что нижеперечисленное оборудование принято в эксплуатацию и установлено '
        f'по адресу: {data.get("location", "___")}. '
        f'Оборудование проверено, находится в рабочем состоянии и соответствует техническим требованиям.'
    )
    ir.font.size = Pt(12)
    ir.font.name = 'Times New Roman'

    # ══════════════════════════════════════════
    #  ТАБЛИЦА ОБОРУДОВАНИЯ
    # ══════════════════════════════════════════
    items = data.get('items', [])

    headers = [
        '№',
        'Наименование\nоборудования',
        'Модель /\nМарка',
        'Серийный\nномер',
        'Инв.\nномер',
        'Год\nвыпуска',
        'Стоимость\n(сум)',
        'Дата\nустановки',
        'Состояние',
        'Примечание',
    ]
    col_widths = [
        Cm(0.8), Cm(3.8), Cm(2.5), Cm(2.8),
        Cm(2.3), Cm(1.5), Cm(2.5), Cm(2.3),
        Cm(2.0), Cm(2.0),
    ]

    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = 'Table Grid'

    # Заголовок таблицы
    hdr_row = table.rows[0]
    for i, (cell, width, hdr) in enumerate(zip(hdr_row.cells, col_widths, headers)):
        cell.width = width
        set_cell_bg(cell, '1A5276')
        set_cell_border(cell)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = Pt(12)
        run = p.add_run(hdr)
        run.bold = True
        run.font.size = Pt(8)
        run.font.name = 'Times New Roman'
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    # Строки данных
    for i, item in enumerate(items):
        row = table.add_row()
        bg = 'EBF5FB' if i % 2 == 0 else 'FFFFFF'
        condition = item.get('condition', '')
        values = [
            str(item.get('num', i + 1)),
            item.get('name', ''),
            item.get('model', ''),
            item.get('serial_number', ''),
            item.get('inv_number', ''),
            item.get('year', ''),
            item.get('cost', ''),
            item.get('install_date', ''),
            condition,
            item.get('note', ''),
        ]
        for j, (cell, val) in enumerate(zip(row.cells, values)):
            set_cell_bg(cell, bg)
            set_cell_border(cell)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if j in (0, 3, 4, 5, 6, 7) else WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.line_spacing = Pt(12)
            run = p.add_run(val)
            run.font.size = Pt(9)
            run.font.name = 'Times New Roman'
            # Цветовая метка состояния
            if j == 8:
                if 'янги' in val.lower() or 'новый' in val.lower() or 'new' in val.lower():
                    run.font.color.rgb = RGBColor(0x1E, 0x8B, 0x4C)
                    run.bold = True
                elif 'ишлатилган' in val.lower() or 'б/у' in val.lower() or 'used' in val.lower():
                    run.font.color.rgb = RGBColor(0xD3, 0x7A, 0x00)
                    run.bold = True

    doc.add_paragraph()

    # ══════════════════════════════════════════
    #  ЗАКЛЮЧЕНИЕ
    # ══════════════════════════════════════════
    concl_label = doc.add_paragraph()
    concl_label.paragraph_format.space_before = Pt(6)
    concl_label.paragraph_format.space_after = Pt(3)
    clr = concl_label.add_run('ЗАКЛЮЧЕНИЕ КОМИССИИ:')
    clr.bold = True
    clr.font.size = Pt(12)
    clr.font.name = 'Times New Roman'
    clr.font.color.rgb = RGBColor(0x1A, 0x52, 0x76)

    concl_p = doc.add_paragraph()
    concl_p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    concl_p.paragraph_format.space_before = Pt(0)
    concl_p.paragraph_format.space_after = Pt(6)
    concl_p.paragraph_format.line_spacing = Pt(16)
    crr = concl_p.add_run(
        f'\tКомиссия установила, что перечисленное оборудование в количестве {len(items)} ед. '
        f'установлено, проверено и введено в эксплуатацию. Оборудование находится '
        f'в технически исправном состоянии, соответствует техническим требованиям и готово к работе. '
        f'Комиссия рекомендует принять оборудование на баланс {data.get("org_name", "___")}.'
    )
    crr.font.size = Pt(12)
    crr.font.name = 'Times New Roman'

    doc.add_paragraph()
    doc.add_paragraph()

    # ══════════════════════════════════════════
    #  ПОДПИСИ
    # ══════════════════════════════════════════
    sig_label = doc.add_paragraph()
    sig_label.paragraph_format.space_before = Pt(4)
    sig_label.paragraph_format.space_after = Pt(6)
    slr = sig_label.add_run('ПОДПИСИ ЧЛЕНОВ КОМИССИИ:')
    slr.bold = True
    slr.font.size = Pt(11)
    slr.font.name = 'Times New Roman'
    slr.font.color.rgb = RGBColor(0x1A, 0x52, 0x76)

    sig_table = doc.add_table(rows=3, cols=3)
    sig_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    sig_entries = [
        ('Председатель:', data.get('commission_head', '_______________')),
        ('Член комиссии:', data.get('member1', '_______________')),
        ('Член комиссии:', data.get('member2', '_______________')),
    ]
    for i, (label, name) in enumerate(sig_entries):
        row = sig_table.rows[i]
        row.cells[0].width = Cm(4)
        row.cells[1].width = Cm(6)
        row.cells[2].width = Cm(6)
        for ci, (cell, text, bold) in enumerate(zip(row.cells,
                                                     [label, '_____________________', name],
                                                     [True, False, False])):
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(6)
            r = p.add_run(text)
            r.bold = bold
            r.font.size = Pt(11)
            r.font.name = 'Times New Roman'

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
    top2.set(qn('w:color'), '1A5276')
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
#  Сохранение — отдельный файл на каждый акт
# ─────────────────────────────────────────────

def save_single_files(data_list: list, output_dir: str) -> list:
    os.makedirs(output_dir, exist_ok=True)
    created = []
    for data in data_list:
        doc = create_ust_doc(data)
        safe = data.get('doc_number', 'N').replace('/', '_').replace('\\', '_')
        name = f"Akt_Ustanovki_{safe}.docx"
        path = os.path.join(output_dir, name)
        doc.save(path)
        created.append(path)
    return created


# ─────────────────────────────────────────────
#  Сохранение — все акты в одном файле
# ─────────────────────────────────────────────

def save_all_in_one(data_list: list, output_path: str) -> str:
    combined = Document()
    set_doc_margins(combined)
    combined.styles['Normal'].font.name = 'Times New Roman'
    combined.styles['Normal'].font.size = Pt(12)

    for idx, data in enumerate(data_list):
        single_doc = create_ust_doc(data)
        for element in single_doc.element.body:
            combined.element.body.append(copy.deepcopy(element))
        if idx < len(data_list) - 1:
            pb_p = OxmlElement('w:p')
            pb_r = OxmlElement('w:r')
            pb_br = OxmlElement('w:br')
            pb_br.set(qn('w:type'), 'page')
            pb_r.append(pb_br)
            pb_p.append(pb_r)
            combined.element.body.append(pb_p)

    combined.save(output_path)
    return output_path
