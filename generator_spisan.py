# generator_spisan.py
# Генератор Word актов СПИСАНИЯ
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

from transliterate import TEXTS_SPISAN, MONTHS, convert_to, detect_script

# ─────────────────────────────────────────────
#  Вспомогательные функции
# ─────────────────────────────────────────────

def _set_cell_bg(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)

def _set_cell_border(cell, color='2E4057'):
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

def _set_margins(doc, top=2.0, bottom=2.0, left=2.5, right=1.5):
    sec = doc.sections[0]
    sec.top_margin    = Cm(top)
    sec.bottom_margin = Cm(bottom)
    sec.left_margin   = Cm(left)
    sec.right_margin  = Cm(right)

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

def _hr(doc, color='2E4057'):
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

def _auto_lang(group: dict) -> str:
    """
    Автоматически определяет язык по тексту в данных группы.
    Смотрит org_name, commission_head, member1 — берёт первый непустой.
    """
    for key in ('org_name', 'commission_head', 'member1', 'member2'):
        val = group.get(key, '').strip()
        if val:
            script = detect_script(val)
            if script in ('latin', 'cyrillic'):
                return script
    # Смотрим названия устройств
    for dev in group.get('devices', []):
        name = dev.get('name', '').strip()
        if name:
            script = detect_script(name)
            if script in ('latin', 'cyrillic'):
                return script
    return 'latin'  # по умолчанию


# ─────────────────────────────────────────────
#  Основной генератор
# ─────────────────────────────────────────────

def create_spisan_doc(group: dict, doc_number: str = '1',
                      lang: str = 'auto') -> Document:
    """
    Создаёт Word акт технического заключения (списания).

    group = {
        'inv_number':      str,
        'devices': [{'name': str, 'parts': [{'part_name','condition','defect'}]}],
        'org_name':        str,
        'region':          str,
        'commission_head': str,
        'member1':         str,
        'member2':         str,
        'doc_date':        str,   # дд.мм.гггг
    }
    lang: 'latin' | 'cyrillic' | 'auto'
    """
    # Определяем язык
    if lang == 'auto':
        lang = _auto_lang(group)

    # Конвертируем пользовательские данные если нужно
    org_name = convert_to(group.get('org_name', ''),        lang)
    region   = convert_to(group.get('region', ''),          lang)
    boss     = convert_to(group.get('commission_head', ''), lang)
    eng1     = convert_to(group.get('member1', ''),         lang)
    eng2     = convert_to(group.get('member2', ''),         lang)
    inv      = group.get('inv_number', '')

    # Устройства и компоненты — названия НЕ конвертируем (оставляем как есть)
    # Конвертируем только condition/defect если они на кириллице/латинице
    devices = []
    for dev in group.get('devices', []):
        parts = []
        for p in dev.get('parts', []):
            parts.append({
                'part_name': p.get('part_name', ''),   # не конвертируем
                'condition': p.get('condition', ''),   # не конвертируем
                'defect':    p.get('defect', ''),      # не конвертируем
            })
        devices.append({
            'name':  dev.get('name', ''),   # не конвертируем
            'parts': parts,
        })

    # Системные тексты по выбранному алфавиту
    T = TEXTS_SPISAN[lang]
    M = MONTHS[lang]

    # Дата
    today    = datetime.now()
    date_str = group.get('doc_date', today.strftime('%d.%m.%Y'))
    try:
        dt    = datetime.strptime(date_str, '%d.%m.%Y')
        day   = str(dt.day)
        month = M[dt.month]
        year  = str(dt.year)
    except Exception:
        day   = str(today.day)
        month = M[today.month]
        year  = str(today.year)

    # ── Создаём документ ──
    doc = Document()
    _set_margins(doc)
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(12)

    # ══════════════════════════════════
    #  ШАПКА — таблица: пустая левая | текст справа
    # ══════════════════════════════════
    hdr_tbl = doc.add_table(rows=1, cols=2)
    hdr_tbl.alignment = WD_TABLE_ALIGNMENT.RIGHT
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

    hdr_tbl.rows[0].cells[0].width = Cm(5)
    hdr_tbl.rows[0].cells[1].width = Cm(12)
    right_cell = hdr_tbl.rows[0].cells[1]

    # Удаляем пустой параграф по умолчанию (безопасно — копируем список)
    for p in list(right_cell.paragraphs):
        p._element.getparent().remove(p._element)

    def hdr_p(text, bold=False, size=11):
        p = right_cell.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(1)
        p.paragraph_format.line_spacing = Pt(14)
        r = p.add_run(text)
        r.bold = bold
        r.font.size = Pt(size)
        r.font.name = 'Times New Roman'
        return p

    # «TASDIQLAYMAN»
    hdr_p(T['tasdiq'], bold=True, size=11)

    # Организация + должность — разбиваем на 2 строки если длинно
    # В TEXTS_SPISAN нет 'rahbar' — используем должность из данных или пустую строку
    rahbar_title = T.get('rahbar', '')
    org_rahbar = f'{org_name} {rahbar_title}'.strip() if rahbar_title else org_name
    if len(org_rahbar) > 40:
        words = org_rahbar.split()
        mid   = len(words) // 2
        hdr_p(' '.join(words[:mid]), size=10)
        hdr_p(' '.join(words[mid:]), size=10)
    else:
        hdr_p(org_rahbar, size=10)

    # Линия подписи + ФИО
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

    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    _hr(doc)

    # ══════════════════════════════════
    #  ЗАГОЛОВОК
    # ══════════════════════════════════
    _p(doc, f'{T["title"]} {doc_number}',
       bold=True, size=15, align=WD_ALIGN_PARAGRAPH.CENTER,
       color=(0x2E, 0x40, 0x57), before=6, after=6)

    _p(doc, f'{org_name}  |  {region}',
       size=10, align=WD_ALIGN_PARAGRAPH.CENTER,
       italic=True, after=8, color=(0x55, 0x55, 0x55))

    # ══════════════════════════════════
    #  ВВОДНЫЙ ТЕКСТ
    # ══════════════════════════════════
    intro = doc.add_paragraph()
    intro.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    intro.paragraph_format.space_before = Pt(0)
    intro.paragraph_format.space_after  = Pt(8)
    intro.paragraph_format.line_spacing = Pt(16)
    intro_text = '\t' + T['intro'].format(org=org_name, region=region,
                                          eng1=eng1, eng2=eng2)
    ir = intro.add_run(intro_text)
    ir.font.size = Pt(12)
    ir.font.name = 'Times New Roman'

    # ══════════════════════════════════
    #  ТАБЛИЦА
    # ══════════════════════════════════
    COL_WIDTHS   = [Cm(6.0), Cm(4.5), Cm(7.0)]
    HEADER_COLOR = '2E4057'
    EVEN_BG      = 'F0F4F8'
    ODD_BG       = 'FFFFFF'
    BAD_COLOR    = (0xC0, 0x39, 0x2B)
    GOOD_COLOR   = (0x1E, 0x8B, 0x4C)

    table = doc.add_table(rows=0, cols=3)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = 'Table Grid'

    for dev in devices:
        dev_name = dev.get('name', '').strip()
        parts    = dev.get('parts', [])

        # Объединённая строка: Qurilma nomi + Inventar raqami
        title_row = table.add_row()
        title_row.cells[0].merge(title_row.cells[1])
        title_row.cells[0].merge(title_row.cells[2])
        cell = title_row.cells[0]
        _set_cell_bg(cell, 'D6EAF8')
        _set_cell_border(cell, '2E4057')
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p_t = cell.paragraphs[0]
        p_t.paragraph_format.space_before = Pt(3)
        p_t.paragraph_format.space_after  = Pt(3)

        for txt, bold, color in [
            (f'{T["qurilma"]}  ', True,  (0x1A, 0x52, 0x76)),
            (dev_name,           True,  None),
            (f'     {T["inventar"]}  ', True, (0x1A, 0x52, 0x76)),
            (inv,                True,  None),
        ]:
            r = p_t.add_run(txt)
            r.bold = bold
            r.font.size = Pt(11)
            r.font.name = 'Times New Roman'
            if color:
                r.font.color.rgb = RGBColor(*color)

        # Заголовок колонок
        hdr_row = table.add_row()
        hdr_row.height = Pt(22)
        for ci, (cell, label, w) in enumerate(zip(
                hdr_row.cells,
                [T['col_qism'], T['col_yarоq'], T['col_nosoz']],
                COL_WIDTHS)):
            cell.width = w
            _set_cell_bg(cell, HEADER_COLOR)
            _set_cell_border(cell, HEADER_COLOR)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after  = Pt(2)
            p.paragraph_format.line_spacing = Pt(13)
            run = p.add_run(label)
            run.bold = True
            run.font.size = Pt(9)
            run.font.name = 'Times New Roman'
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

        # Строки компонентов
        for pi, part in enumerate(parts):
            part_name = part.get('part_name', '')
            condition = part.get('condition', '')
            defect    = part.get('defect', '')
            is_bad    = any(w in condition.lower() for w in
                           ('yaroqsiz', 'яроқсиз', 'yaramsiz', 'нет', 'неисправен'))

            data_row = table.add_row()
            data_row.height = Pt(20)
            bg = EVEN_BG if pi % 2 == 0 else ODD_BG

            for ci, (cell, val) in enumerate(zip(data_row.cells,
                                                  [part_name, condition, defect])):
                _set_cell_bg(cell, bg)
                _set_cell_border(cell)
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                p = cell.paragraphs[0]
                p.alignment = (WD_ALIGN_PARAGRAPH.CENTER if ci == 1
                               else WD_ALIGN_PARAGRAPH.LEFT)
                p.paragraph_format.space_before = Pt(2)
                p.paragraph_format.space_after  = Pt(2)
                p.paragraph_format.line_spacing = Pt(13)
                run = p.add_run(val)
                run.font.size = Pt(10)
                run.font.name = 'Times New Roman'
                if ci == 1:
                    if is_bad:
                        run.font.color.rgb = RGBColor(*BAD_COLOR)
                        run.bold = True
                    else:
                        run.font.color.rgb = RGBColor(*GOOD_COLOR)

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # ══════════════════════════════════
    #  XULOSA / ХУЛОСА
    # ══════════════════════════════════
    _p(doc, T['xulosa'], bold=True, size=13,
       align=WD_ALIGN_PARAGRAPH.CENTER,
       color=(0x2E, 0x40, 0x57), before=6, after=4)

    dev_names = [d.get('name', '').strip() for d in devices if d.get('name', '').strip()]
    names_str = ', '.join(f'«{n}»' for n in dev_names)

    for txt_key in ('xulosa1', 'xulosa2'):
        xp = doc.add_paragraph()
        xp.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        xp.paragraph_format.space_before = Pt(0)
        xp.paragraph_format.space_after  = Pt(4)
        xp.paragraph_format.line_spacing = Pt(16)
        text = '\t' + T[txt_key].format(names=names_str)
        xr = xp.add_run(text)
        xr.font.size = Pt(12)
        xr.font.name = 'Times New Roman'

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # ══════════════════════════════════
    #  ПОДПИСИ
    # ══════════════════════════════════
    sig_tbl = doc.add_table(rows=2, cols=3)
    sig_tbl.alignment = WD_TABLE_ALIGNMENT.LEFT
    for i, engineer in enumerate([eng1, eng2]):
        row = sig_tbl.rows[i]
        row.cells[0].width = Cm(5)
        row.cells[1].width = Cm(7)
        row.cells[2].width = Cm(5)

        def sc(cell, text, bold=False):
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after  = Pt(8)
            r = p.add_run(text)
            r.bold = bold
            r.font.size = Pt(11)
            r.font.name = 'Times New Roman'

        sc(row.cells[0], T['muhandis'], bold=True)
        sc(row.cells[1], '______________________')
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
    top_e.set(qn('w:color'), '2E4057')
    pBdr.append(top_e)
    pPr.append(pBdr)
    fr = p_foot.add_run(
        f"{T['mp']}   {T['inv_label']} {inv}   |   {T['sana']} {date_str}"
    )
    fr.font.size = Pt(9)
    fr.font.name = 'Times New Roman'
    fr.italic = True
    fr.font.color.rgb = RGBColor(0x88, 0x88, 0x88)

    return doc


# ─────────────────────────────────────────────
#  Сохранение
# ─────────────────────────────────────────────

def _safe_filename(text: str) -> str:
    """Убирает все символы недопустимые в имени файла Windows/Linux."""
    # Заменяем переносы строк, табуляции и пробелы на подчёркивание
    text = text.replace('\n', '_').replace('\r', '').replace('\t', '_')
    # Убираем недопустимые символы Windows: \ / : * ? " < > |
    for ch in r'\/:*?"<>|':
        text = text.replace(ch, '-')
    # Убираем лишние пробелы и подчёркивания
    text = '_'.join(text.split())
    return text[:60] or 'doc'  # максимум 60 символов


def save_single_files(groups: list, output_dir: str,
                      lang: str = 'auto') -> list:
    """Один .docx на каждую группу (инвентарный номер)."""
    os.makedirs(output_dir, exist_ok=True)
    created = []
    for i, group in enumerate(groups, start=1):
        inv  = _safe_filename(group.get('inv_number', str(i)))
        doc  = create_spisan_doc(group, doc_number=str(i), lang=lang)
        name = f'Texnik_Xulosa_{inv}.docx'
        path = os.path.join(output_dir, name)
        doc.save(path)
        created.append(path)
    return created


def save_all_in_one(groups: list, output_path: str,
                    lang: str = 'auto') -> str:
    """Все акты в одном файле."""
    combined = Document()
    combined.sections[0].top_margin    = Cm(2.0)
    combined.sections[0].bottom_margin = Cm(2.0)
    combined.sections[0].left_margin   = Cm(2.5)
    combined.sections[0].right_margin  = Cm(1.5)
    combined.styles['Normal'].font.name = 'Times New Roman'
    combined.styles['Normal'].font.size = Pt(12)

    for idx, group in enumerate(groups, start=1):
        single = create_spisan_doc(group, doc_number=str(idx), lang=lang)
        for element in single.element.body:
            combined.element.body.append(copy.deepcopy(element))
        if idx < len(groups):
            pb = OxmlElement('w:p')
            pb_r = OxmlElement('w:r')
            pb_br = OxmlElement('w:br')
            pb_br.set(qn('w:type'), 'page')
            pb_r.append(pb_br)
            pb.append(pb_r)
            combined.element.body.append(pb)

    combined.save(output_path)
    return output_path
