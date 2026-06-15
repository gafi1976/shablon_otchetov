# generator_spisan.py
# Генератор Word актов СПИСАНИЯ по шаблону shablon_spisan.docx
#
# Структура выходного документа (один файл = одна группа по инвентарному номеру):
#
# «TASDIQLAYMAN» {Tashilot nomi}
# __________ {boss name} «{day}»__{month}__ {year} yil
#
# TEXNIK XULOSA № {num}
#
# Biz, quyida imzo chekuvchilar — ABM MChJ {region} bo'lim {enginer1} va {enginer2} ...
#
# [ТАБЛИЦА]:
#   Qurilma nomi: {oborud}  Inventar raqami: {inv_num}
#   Qismlar nomi | Foydalanishga yaroqliligi | Nosozlik belgilari
#   <строки компонентов>
#
# XULOSA
# <текст заключения>
#
# Yetakchi muhandis: ____  {enginer1}
# Yetakchi muhandis: ____  {enginer2}

import os
import copy
from datetime import datetime
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ─────────────────────────────────────────────
#  Вспомогательные функции
# ─────────────────────────────────────────────

MONTHS_UZ = {
    1:'yanvar', 2:'fevral', 3:'mart', 4:'aprel',
    5:'may', 6:'iyun', 7:'iyul', 8:'avgust',
    9:'sentabr', 10:'oktabr', 11:'noyabr', 12:'dekabr'
}

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
    for edge in ('top','left','bottom','right'):
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
       color=None, before=0, after=4, italic=False, underline=False):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after  = Pt(after)
    p.paragraph_format.line_spacing = Pt(15)
    if text:
        run = p.add_run(text)
        run.bold      = bold
        run.italic    = italic
        run.underline = underline
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
    for edge in ('bottom',):
        e = OxmlElement(f'w:{edge}')
        e.set(qn('w:val'), 'single')
        e.set(qn('w:sz'), '12')
        e.set(qn('w:space'), '1')
        e.set(qn('w:color'), color)
        pBdr.append(e)
    pPr.append(pBdr)


# ─────────────────────────────────────────────
#  Основной генератор — один документ на
#  один инвентарный номер (группа устройств)
# ─────────────────────────────────────────────

def create_spisan_doc(group: dict, doc_number: str = '1') -> Document:
    """
    group = {
        'inv_number':      str,   # инвентарный номер — имя файла
        'devices': [              # список устройств внутри группы
            {
                'name':  str,     # Qurilma nomi
                'parts': [        # компоненты
                    {'part_name': str, 'condition': str, 'defect': str}
                ]
            }
        ],
        'org_name':        str,   # Tashilot nomi
        'region':          str,
        'commission_head': str,   # boss name
        'member1':         str,   # enginer1
        'member2':         str,   # enginer2
        'doc_date':        str,   # дд.мм.гггг
    }
    """
    doc = Document()
    _set_margins(doc)

    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(12)

    today = datetime.now()
    date_str = group.get('doc_date', today.strftime('%d.%m.%Y'))
    try:
        dt = datetime.strptime(date_str, '%d.%m.%Y')
        day   = str(dt.day)
        month = MONTHS_UZ[dt.month]
        year  = str(dt.year)
    except Exception:
        day   = str(today.day)
        month = MONTHS_UZ[today.month]
        year  = str(today.year)

    org_name = group.get('org_name', '')
    region   = group.get('region', '')
    boss     = group.get('commission_head', '')
    eng1     = group.get('member1', '')
    eng2     = group.get('member2', '')
    inv      = group.get('inv_number', '')
    devices  = group.get('devices', [])

    # ══════════════════════════════════
    #  ШАПКА — «TASDIQLAYMAN»
    # ══════════════════════════════════
    p_hdr = doc.add_paragraph()
    p_hdr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_hdr.paragraph_format.space_before = Pt(0)
    p_hdr.paragraph_format.space_after  = Pt(0)

    def rr(p, text, bold=False, size=11):
        r = p.add_run(text)
        r.bold = bold
        r.font.size = Pt(size)
        r.font.name = 'Times New Roman'

    rr(p_hdr, '«TASDIQLAYMAN»\n', bold=True, size=11)
    rr(p_hdr, f'{org_name}\n', size=10)
    rr(p_hdr, f'__________ {boss}\n', size=10)
    rr(p_hdr, f'«{day}» {month} {year} yil', size=10)

    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    _hr(doc)

    # ══════════════════════════════════
    #  ЗАГОЛОВОК
    # ══════════════════════════════════
    _p(doc, f'TEXNIK XULOSA № {doc_number}',
       bold=True, size=15, align=WD_ALIGN_PARAGRAPH.CENTER,
       color=(0x2E, 0x40, 0x57), before=6, after=6)

    # Регион + организация
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
    ir = intro.add_run(
        f"\tBiz, quyida imzo chekuvchilar — {org_name} {region} bo'lim "
        f"yetakchi muhandisi {eng1} va yetakchi muhandis {eng2} lar, "
        f"quyida keltirilgan qurilmalarni texnik ko'rikdan o'tkazdik:"
    )
    ir.font.size = Pt(12)
    ir.font.name = 'Times New Roman'

    # ══════════════════════════════════
    #  ТАБЛИЦА — все устройства группы
    # ══════════════════════════════════
    # Одна таблица на всю группу (инвентарный номер)
    # Структура как в шаблоне:
    #   [merged] Qurilma nomi: X   Inventar raqami: INV
    #   Qismlar nomi | Yaroqliligi | Nosozlik belgilari
    #   <части>
    #   ... повторяем для каждого устройства в группе

    COL_WIDTHS = [Cm(6.0), Cm(4.5), Cm(7.0)]
    HEADER_COLOR = '2E4057'
    HDR_FONT     = (0xFF, 0xFF, 0xFF)
    EVEN_BG      = 'F0F4F8'
    ODD_BG       = 'FFFFFF'
    BAD_COLOR    = (0xC0, 0x39, 0x2B)
    GOOD_COLOR   = (0x1E, 0x8B, 0x4C)

    table = doc.add_table(rows=0, cols=3)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = 'Table Grid'

    for dev in devices:
        dev_name  = dev.get('name', '').strip()
        parts     = dev.get('parts', [])

        # ── Строка: Qurilma nomi + Inventar raqami (merged) ──
        title_row = table.add_row()
        # Объединяем все 3 ячейки
        title_row.cells[0].merge(title_row.cells[1])
        title_row.cells[0].merge(title_row.cells[2])
        cell = title_row.cells[0]
        _set_cell_bg(cell, 'D6EAF8')
        _set_cell_border(cell, '2E4057')
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p_title = cell.paragraphs[0]
        p_title.paragraph_format.space_before = Pt(3)
        p_title.paragraph_format.space_after  = Pt(3)

        r1 = p_title.add_run('Qurilma nomi:  ')
        r1.bold = True
        r1.font.size = Pt(11)
        r1.font.name = 'Times New Roman'
        r1.font.color.rgb = RGBColor(0x1A, 0x52, 0x76)

        r2 = p_title.add_run(dev_name)
        r2.bold = True
        r2.font.size = Pt(11)
        r2.font.name = 'Times New Roman'

        r3 = p_title.add_run('     Inventar raqami:  ')
        r3.bold = True
        r3.font.size = Pt(11)
        r3.font.name = 'Times New Roman'
        r3.font.color.rgb = RGBColor(0x1A, 0x52, 0x76)

        r4 = p_title.add_run(inv)
        r4.bold = True
        r4.font.size = Pt(11)
        r4.font.name = 'Times New Roman'

        # ── Строка заголовков колонок ──
        hdr_row = table.add_row()
        hdr_row.height = Pt(22)
        hdr_labels = ['Qismlar nomi', 'Foydalanishga\nyaroqliligi', 'Nosozlik belgilari']
        for ci, (cell, label, w) in enumerate(zip(hdr_row.cells, hdr_labels, COL_WIDTHS)):
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
            run.font.color.rgb = RGBColor(*HDR_FONT)

        # ── Строки компонентов ──
        for pi, part in enumerate(parts):
            part_name = part.get('part_name', '')
            condition = part.get('condition', '')
            defect    = part.get('defect', '')
            is_bad    = condition.lower() in ('yaroqsiz', 'яроқсиз', 'нет', 'неисправен')

            data_row = table.add_row()
            data_row.height = Pt(20)
            bg = EVEN_BG if pi % 2 == 0 else ODD_BG

            for ci, (cell, val) in enumerate(zip(data_row.cells,
                                                  [part_name, condition, defect])):
                _set_cell_bg(cell, bg)
                _set_cell_border(cell)
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                p = cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER if ci == 1 else WD_ALIGN_PARAGRAPH.LEFT
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
    #  XULOSA (Заключение)
    # ══════════════════════════════════
    _p(doc, 'XULOSA', bold=True, size=13,
       align=WD_ALIGN_PARAGRAPH.CENTER,
       color=(0x2E, 0x40, 0x57), before=6, after=4)

    # Список устройств для заключения
    dev_names = [d.get('name','').strip() for d in devices if d.get('name','').strip()]
    names_str = ', '.join(f'«{n}»' for n in dev_names)

    xulosa1 = doc.add_paragraph()
    xulosa1.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    xulosa1.paragraph_format.space_before = Pt(0)
    xulosa1.paragraph_format.space_after  = Pt(4)
    xulosa1.paragraph_format.line_spacing = Pt(16)
    xr1 = xulosa1.add_run(
        f"\t{names_str} zamonaviy talablarga javob bermaydi, ma'naviy eskirgan. "
        f"Ta'mirlash uchun zarur bo'lgan ehtiyot qismlar texnologik jarayondan "
        f"olib tashlanganligini hisobga olib, uni tiklash maqsadga muvofiq emas."
    )
    xr1.font.size = Pt(12)
    xr1.font.name = 'Times New Roman'

    xulosa2 = doc.add_paragraph()
    xulosa2.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    xulosa2.paragraph_format.space_before = Pt(0)
    xulosa2.paragraph_format.space_after  = Pt(6)
    xulosa2.paragraph_format.line_spacing = Pt(16)
    xr2 = xulosa2.add_run(
        f"\tYuqorida keltirib o'tilgan jadvaldagi qurilmalarni xolatlaridan "
        f"kelib chiqib, komissiya asosiy vositalar ro'yxatidan chiqarilsin "
        f"degan xulosaga keldi."
    )
    xr2.font.size = Pt(12)
    xr2.font.name = 'Times New Roman'

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # ══════════════════════════════════
    #  ПОДПИСИ
    # ══════════════════════════════════
    sig_tbl = doc.add_table(rows=2, cols=3)
    sig_tbl.alignment = WD_TABLE_ALIGNMENT.LEFT

    for i, (label, engineer) in enumerate([
        ('Yetakchi muhandis:', eng1),
        ('Yetakchi muhandis:', eng2),
    ]):
        row = sig_tbl.rows[i]
        row.cells[0].width = Cm(5)
        row.cells[1].width = Cm(7)
        row.cells[2].width = Cm(5)

        def sig_cell(cell, text, bold=False):
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after  = Pt(8)
            r = p.add_run(text)
            r.bold = bold
            r.font.size = Pt(11)
            r.font.name = 'Times New Roman'

        sig_cell(row.cells[0], label, bold=True)
        sig_cell(row.cells[1], '______________________')
        sig_cell(row.cells[2], engineer)

    # ── Нижняя линия ──
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
        f"М.П.   Inventar raqami: {inv}   |   Sana: {date_str}"
    )
    fr.font.size = Pt(9)
    fr.font.name = 'Times New Roman'
    fr.italic = True
    fr.font.color.rgb = RGBColor(0x88, 0x88, 0x88)

    return doc


# ─────────────────────────────────────────────
#  Сохранение файлов
# ─────────────────────────────────────────────

def save_single_files(groups: list, output_dir: str) -> list:
    """
    Один .docx файл на каждую группу (инвентарный номер).
    Имя файла: Texnik_Xulosa_{inv_number}.docx
    """
    os.makedirs(output_dir, exist_ok=True)
    created = []
    for i, group in enumerate(groups, start=1):
        inv  = group.get('inv_number', str(i)).replace('/', '-').replace('\\', '-')
        doc  = create_spisan_doc(group, doc_number=str(i))
        name = f'Texnik_Xulosa_{inv}.docx'
        path = os.path.join(output_dir, name)
        doc.save(path)
        created.append(path)
    return created


def save_all_in_one(groups: list, output_path: str) -> str:
    """Все акты в одном файле, разделённые разрывом страницы."""
    combined = Document()
    combined.sections[0].top_margin    = Cm(2.0)
    combined.sections[0].bottom_margin = Cm(2.0)
    combined.sections[0].left_margin   = Cm(2.5)
    combined.sections[0].right_margin  = Cm(1.5)
    combined.styles['Normal'].font.name = 'Times New Roman'
    combined.styles['Normal'].font.size = Pt(12)

    for idx, group in enumerate(groups, start=1):
        single = create_spisan_doc(group, doc_number=str(idx))
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
