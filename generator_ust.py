# generator_ust.py
# Генератор Word актов УСТАНОВКИ
# Поддерживает два алфавита: lang='latin' (по умолчанию) или lang='cyrillic'

import os
import copy
from datetime import datetime
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from transliterate import TEXTS_UST, MONTHS, convert_to, detect_script


def _set_cell_bg(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)


def _set_cell_border(cell, color='1A5276'):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for edge in ('top', 'left', 'bottom', 'right'):
        b = OxmlElement(f'w:{edge}')
        b.set(qn('w:val'), 'single')
        b.set(qn('w:sz'), '6')
        b.set(qn('w:space'), '0')
        b.set(qn('w:color'), color)
        tcBorders.append(b)
    tcPr.append(tcBorders)


def _set_margins(doc):
    sec = doc.sections[0]
    sec.top_margin    = Cm(2.0)
    sec.bottom_margin = Cm(2.0)
    sec.left_margin   = Cm(2.5)
    sec.right_margin  = Cm(1.5)


def _p(doc, text='', bold=False, size=12, align=WD_ALIGN_PARAGRAPH.LEFT,
       color=None, before=0, after=4, italic=False):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after  = Pt(after)
    p.paragraph_format.line_spacing = Pt(15)
    if text:
        run = p.add_run(text)
        run.bold   = bold
        run.italic = italic
        run.font.size = Pt(size)
        run.font.name = 'Times New Roman'
        if color:
            run.font.color.rgb = RGBColor(*color)
    return p


def _hr(doc, color='1A5276'):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(4)
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    e = OxmlElement('w:bottom')
    e.set(qn('w:val'), 'single')
    e.set(qn('w:sz'), '12')
    e.set(qn('w:space'), '1')
    e.set(qn('w:color'), color)
    pBdr.append(e)
    pPr.append(pBdr)


def _auto_lang(data: dict) -> str:
    """Автоопределение алфавита по данным."""
    for key in ('org_name', 'commission_head', 'member1', 'member2', 'address'):
        val = data.get(key, '').strip()
        if val:
            script = detect_script(val)
            if script in ('latin', 'cyrillic'):
                return script
    items = data.get('items', [])
    if items:
        name = items[0].get('name', '').strip()
        if name:
            script = detect_script(name)
            if script in ('latin', 'cyrillic'):
                return script
    return 'latin'


def create_ust_doc(data: dict, doc_number: str = '1',
                   lang: str = 'auto') -> Document:
    """
    Создаёт Word акт установки оборудования.
    lang: 'latin' | 'cyrillic' | 'auto'
    """
    if lang == 'auto':
        lang = _auto_lang(data)

    # Конвертируем пользовательские данные
    org_name = convert_to(data.get('org_name', ''),        lang)
    address  = convert_to(data.get('address', data.get('region', '')), lang)
    boss     = convert_to(data.get('commission_head', ''), lang)
    eng1     = convert_to(data.get('member1', ''),         lang)
    job1     = convert_to(data.get('member1_title', ''),   lang) or ('Yetakchi muhandis' if lang == 'latin' else 'Етакчи муҳандис')
    eng2     = convert_to(data.get('member2', ''),         lang)
    job2     = convert_to(data.get('member2_title', ''),   lang) or ('Yetakchi muhandis' if lang == 'latin' else 'Етакчи муҳандис')
    num_otch = data.get('doc_number', doc_number)

    # Конвертируем items
    items = []
    for item in data.get('items', []):
        items.append({
            'num':           item.get('num', ''),
            'name':          convert_to(item.get('name', item.get('model', '')), lang),
            'serial_number': item.get('serial_number', ''),
            'install_date':  item.get('install_date', ''),
            'location':      convert_to(item.get('location', item.get('note', '')), lang),
        })

    T = TEXTS_UST[lang]
    M = MONTHS[lang]

    today    = datetime.now()
    date_str = data.get('doc_date', today.strftime('%d.%m.%Y'))
    try:
        dt    = datetime.strptime(date_str, '%d.%m.%Y')
        day   = str(dt.day)
        month = M[dt.month]
        year  = str(dt.year)
    except Exception:
        day   = str(today.day)
        month = M[today.month]
        year  = str(today.year)

    doc = Document()
    _set_margins(doc)
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(12)

    # ══════ ШАПКА — правый блок с таблицей ══════
    # Используем таблицу 1x2: левая колонка пустая, правая — текст шапки
    hdr_tbl = doc.add_table(rows=1, cols=2)
    hdr_tbl.alignment = WD_TABLE_ALIGNMENT.RIGHT
    # Убираем видимые границы таблицы
    for cell in hdr_tbl.rows[0].cells:
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        tcBorders = OxmlElement('w:tcBorders')
        for edge in ('top','left','bottom','right'):
            b = OxmlElement(f'w:{edge}')
            b.set(qn('w:val'), 'none')
            b.set(qn('w:sz'), '0')
            b.set(qn('w:space'), '0')
            b.set(qn('w:color'), 'auto')
            tcBorders.append(b)
        tcPr.append(tcBorders)

    hdr_tbl.rows[0].cells[0].width = Cm(5)   # пустая левая колонка
    hdr_tbl.rows[0].cells[1].width = Cm(12)  # правая колонка с шапкой

    right_cell = hdr_tbl.rows[0].cells[1]

    def hdr_p(text, bold=False, size=11, underline=False):
        """Добавляет строку в правую ячейку шапки."""
        p = right_cell.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(1)
        p.paragraph_format.line_spacing = Pt(14)
        r = p.add_run(text)
        r.bold = bold
        r.underline = underline
        r.font.size = Pt(size)
        r.font.name = 'Times New Roman'
        return p

    # Удаляем пустой параграф который Word добавляет в ячейку по умолчанию
    for p in right_cell.paragraphs:
        p._element.getparent().remove(p._element)

    # «TASDIQLAYMAN» — жирный
    hdr_p(T['tasdiq'], bold=True, size=11)

    # Организация + rahbari — перенос если длинная строка
    org_rahbar = f'{org_name} {T["rahbar"]}'
    if len(org_rahbar) > 40:
        words = org_rahbar.split()
        mid   = len(words) // 2
        hdr_p(' '.join(words[:mid]), size=10)
        hdr_p(' '.join(words[mid:]), size=10)
    else:
        hdr_p(org_rahbar, size=10)

    # Подпись: линия + ФИО
    sign_p = right_cell.add_paragraph()
    sign_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sign_p.paragraph_format.space_before = Pt(2)
    sign_p.paragraph_format.space_after  = Pt(1)
    sign_p.paragraph_format.line_spacing = Pt(14)
    r_line = sign_p.add_run('________________  ')
    r_line.font.size = Pt(10)
    r_line.font.name = 'Times New Roman'
    r_boss = sign_p.add_run(boss)
    r_boss.font.size = Pt(10)
    r_boss.font.name = 'Times New Roman'

    # Дата
    hdr_p(f'«{day}» {month} {year} {T["yil"]}', size=10)

    doc.add_paragraph().paragraph_format.space_after = Pt(4)
    _hr(doc)

    # ══════ ОРГАНИЗАЦИЯ ══════
    _p(doc, org_name, bold=True, size=13,
       align=WD_ALIGN_PARAGRAPH.CENTER,
       color=(0x1A, 0x52, 0x76), before=6, after=2)
    _p(doc, f'{T["manzil"]} {address}',
       size=10, align=WD_ALIGN_PARAGRAPH.CENTER,
       italic=True, after=8, color=(0x55, 0x55, 0x55))

    # ══════ ЗАГОЛОВОК ══════
    _p(doc, T['title'], bold=True, size=15,
       align=WD_ALIGN_PARAGRAPH.CENTER,
       color=(0x1A, 0x52, 0x76), before=4, after=3)
    _p(doc, f'{T["dalolatnoma"]}{num_otch}',
       bold=True, size=12,
       align=WD_ALIGN_PARAGRAPH.CENTER, before=0, after=8)

    # ══════ ВВОДНЫЙ ТЕКСТ ══════
    intro = doc.add_paragraph()
    intro.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    intro.paragraph_format.space_before = Pt(0)
    intro.paragraph_format.space_after  = Pt(8)
    intro.paragraph_format.line_spacing = Pt(16)
    ir = intro.add_run('\t' + T['intro'])
    ir.font.size = Pt(12)
    ir.font.name = 'Times New Roman'

    # ══════ ТАБЛИЦА ══════
    COL_WIDTHS   = [Cm(1.2), Cm(2.8), Cm(5.0), Cm(3.5), Cm(5.0)]
    HEADER_COLOR = '1A5276'
    EVEN_BG      = 'EBF5FB'
    ODD_BG       = 'FFFFFF'

    table = doc.add_table(rows=1, cols=5)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = 'Table Grid'

    hdr_row = table.rows[0]
    hdr_row.height = Pt(28)
    hdr_labels = [T['col_num'], T['col_sana'], T['col_uskuna'],
                  T['col_serial'], T['col_qayer']]
    for ci, (cell, label, w) in enumerate(zip(hdr_row.cells, hdr_labels, COL_WIDTHS)):
        cell.width = w
        _set_cell_bg(cell, HEADER_COLOR)
        _set_cell_border(cell, HEADER_COLOR)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after  = Pt(2)
        p.paragraph_format.line_spacing = Pt(12)
        run = p.add_run(label)
        run.bold = True
        run.font.size = Pt(9)
        run.font.name = 'Times New Roman'
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    for i, item in enumerate(items):
        row = table.add_row()
        row.height = Pt(22)
        bg = EVEN_BG if i % 2 == 0 else ODD_BG
        values = [
            item.get('num', str(i + 1)),
            item.get('install_date', date_str) or date_str,
            item.get('name', ''),
            item.get('serial_number', ''),
            item.get('location', ''),
        ]
        for ci, (cell, val) in enumerate(zip(row.cells, values)):
            _set_cell_bg(cell, bg)
            _set_cell_border(cell)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            p.alignment = (WD_ALIGN_PARAGRAPH.CENTER if ci in (0, 1, 3)
                           else WD_ALIGN_PARAGRAPH.LEFT)
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after  = Pt(2)
            p.paragraph_format.line_spacing = Pt(12)
            run = p.add_run(str(val))
            run.font.size = Pt(10)
            run.font.name = 'Times New Roman'

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # ══════ ЗАКЛЮЧЕНИЕ ══════
    outro = doc.add_paragraph()
    outro.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    outro.paragraph_format.space_before = Pt(0)
    outro.paragraph_format.space_after  = Pt(10)
    outro.paragraph_format.line_spacing = Pt(16)
    or_ = outro.add_run('\t' + T['outro'])
    or_.font.size = Pt(12)
    or_.font.name = 'Times New Roman'

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # ══════ ПОДПИСИ ══════
    sig_tbl = doc.add_table(rows=2, cols=3)
    sig_tbl.alignment = WD_TABLE_ALIGNMENT.LEFT
    for i, (job, engineer) in enumerate([(job1, eng1), (job2, eng2)]):
        row = sig_tbl.rows[i]
        row.cells[0].width = Cm(6)
        row.cells[1].width = Cm(6)
        row.cells[2].width = Cm(5)

        def sc(cell, text, bold=False):
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after  = Pt(8)
            r = p.add_run(text)
            r.bold = bold
            r.font.size = Pt(11)
            r.font.name = 'Times New Roman'

        sc(row.cells[0], f'{T["masul"]}\n{job}:', bold=True)
        sc(row.cells[1], '________________________')
        sc(row.cells[2], engineer)

    # Нижняя линия
    doc.add_paragraph().paragraph_format.space_after = Pt(4)
    p_foot = doc.add_paragraph()
    p_foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pPr = p_foot._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    top_e = OxmlElement('w:top')
    top_e.set(qn('w:val'), 'single')
    top_e.set(qn('w:sz'), '12')
    top_e.set(qn('w:space'), '1')
    top_e.set(qn('w:color'), '1A5276')
    pBdr.append(top_e)
    pPr.append(pBdr)
    fr = p_foot.add_run(
        f"{T['mp']}   {T['dalolatnoma']}{num_otch}   |   {T['sana']} {date_str}"
    )
    fr.font.size = Pt(9)
    fr.font.name = 'Times New Roman'
    fr.italic = True
    fr.font.color.rgb = RGBColor(0x88, 0x88, 0x88)

    return doc


def save_single_files(data_list: list, output_dir: str,
                      lang: str = 'auto') -> list:
    """Один .docx на каждый item оборудования."""
    os.makedirs(output_dir, exist_ok=True)
    created = []
    for data in data_list:
        items = data.get('items', [])
        if not items:
            doc  = create_ust_doc(data, doc_number='1', lang=lang)
            path = os.path.join(output_dir, 'Dalolatnoma_1.docx')
            doc.save(path)
            created.append(path)
            continue
        for item in items:
            single_data = {**data, 'items': [item]}
            num      = str(item.get('num', '')).strip()
            name_raw = item.get('name', item.get('model', '')).strip()
            safe     = ''.join(c if c.isalnum() or c in (' ', '-', '_') else '_'
                               for c in name_raw).strip()[:40]
            fname = f'Dalolatnoma_{num}_{safe}.docx' if safe else f'Dalolatnoma_{num}.docx'
            path  = os.path.join(output_dir, fname)
            doc   = create_ust_doc(single_data,
                                   doc_number=num or str(len(created) + 1),
                                   lang=lang)
            doc.save(path)
            created.append(path)
    return created


def save_all_in_one(data_list: list, output_path: str,
                    lang: str = 'auto') -> str:
    """Все акты в одном файле."""
    combined = Document()
    combined.sections[0].top_margin    = Cm(2.0)
    combined.sections[0].bottom_margin = Cm(2.0)
    combined.sections[0].left_margin   = Cm(2.5)
    combined.sections[0].right_margin  = Cm(1.5)
    combined.styles['Normal'].font.name = 'Times New Roman'
    combined.styles['Normal'].font.size = Pt(12)

    for idx, data in enumerate(data_list, start=1):
        single = create_ust_doc(data, doc_number=str(idx), lang=lang)
        for element in single.element.body:
            combined.element.body.append(copy.deepcopy(element))
        if idx < len(data_list):
            pb = OxmlElement('w:p')
            pb_r = OxmlElement('w:r')
            pb_br = OxmlElement('w:br')
            pb_br.set(qn('w:type'), 'page')
            pb_r.append(pb_br)
            pb.append(pb_r)
            combined.element.body.append(pb)

    combined.save(output_path)
    return output_path
