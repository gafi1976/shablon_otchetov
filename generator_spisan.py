# generator_spisan.py
# Альбомный акт списания основных средств
# По образцу AKTOC2026N20.docx

import os
import copy
from datetime import datetime
from docx import Document
from docx.shared import Pt, RGBColor, Cm, Twips
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from transliterate import MONTHS, convert_to, detect_script


# ════════════════════════════════════════════════════
#  ТЕКСТЫ ДОКУМЕНТА
# ════════════════════════════════════════════════════

TEXTS = {
    'cyrillic': {
        'tasdiq':      'Тасдиқлайман:',
        'rahbar':      'раҳбари',
        'title1':      'АСОСИЙ ВОСИТАЛАРНИ РЎЙХАТДАН ЧИҚАРИШ БЎЙИЧА',
        'title2':      'ДАЛОЛАТНОМА',
        'komissiya':   'АБМ {org} томонидан {date} йил {buyruq} сонли буйруғи билан тайинланган комиссия қуйидагиларга асосан',
        'ish_chiq':    '{year} йилда ишлаб чиқилган.',
        'kelgan':      '{year} йилда корхонага келган.',
        'ishga':       '{year} йилда ишга туширилган.',
        'kapital':     "Капитал таъмирлаш сони {soni}, унга {summa} сўм харажат қилинган.",
        'texnik':      'Техник ҳолати: маънавий эскирган, ишга яроқсиз ҳолга келган, тузатиш мақсадга мувофиқ эмас.',
        'xulosa':      'Комиссия хулосаси: {org} да ишлатилган {names} {date} йил «{day}» {month} "Техник хулоса" га асосан асосий воситалар рўйхатидан чиқарилсин.',
        'moddiy':      'Бўлим бўйича моддий жавобгар шахс:',
        'komrais':     'Комиссия раиси {org}:',
        'komazolari':  'Комиссия аъзолари:',
        # Таблица
        'bulim':       'Бўлим',
        'debet':       'Дебет\nҲисобварақ\nкоди',
        'kredit':      'Кредит\nҲисобварақ\nкоди',
        'nomi':        'Номи',
        'narxi':       'Нархи\n(сўм)',
        'inv':         'Инв.№',
        'amort':       'Амортизация\nчегирмалари\n(сўм)',
        'norma':       'Норма\n(%)',
        'qoldiq':      'Қолдиқ\nнархи\n(сўм)',
    },
    'latin': {
        'tasdiq':      "Tasdiqlayman:",
        'rahbar':      "rahbari",
        'title1':      "ASOSIY VOSITALARNI RO'YXATDAN CHIQARISH BO'YICHA",
        'title2':      'DALOLATNOMA',
        'komissiya':   "ABM {org} tomonidan {date} yil {buyruq} sonli buyrug'i bilan tayinlangan komissiya quyidagilarga asosan",
        'ish_chiq':    '{year} yilda ishlab chiqilgan.',
        'kelgan':      '{year} yilda korxonaga kelgan.',
        'ishga':       '{year} yilda ishga tushirilgan.',
        'kapital':     "Kapital ta'mirlash soni {soni}, unga {summa} so'm xarajat qilingan.",
        'texnik':      "Texnik holati: ma'naviy eskirgan, ishga yaroqsiz holga kelgan, tuzatish maqsadga muvofiq emas.",
        'xulosa':      'Komissiya xulosasi: {org} da ishlatilingan {names} {date} yil «{day}» {month} "Texnik xulosa" ga asosan asosiy vositalar ro\'yxatidan chiqarilsin.',
        'moddiy':      "Bo'lim bo'yicha moddiy javobgar shaxs:",
        'komrais':     "Komissiya raisi {org}:",
        'komazolari':  "Komissiya a'zolari:",
        'bulim':       "Bo'lim",
        'debet':       'Debet\nHisob\nraqami',
        'kredit':      'Kredit\nHisob\nraqami',
        'nomi':        'Nomi',
        'narxi':       "Narxi\n(so'm)",
        'inv':         'Inv.№',
        'amort':       'Amortizatsiya\negirmalari\n(sum)',
        'norma':       'Norma\n(%)',
        'qoldiq':      'Qoldiq\nnarxi\n(sum)',
    },
}


# ════════════════════════════════════════════════════
#  ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# ════════════════════════════════════════════════════

def _auto_lang(group: dict) -> str:
    for key in ('org_name', 'direktor', 'boshliq'):
        val = group.get(key, '').strip()
        if val:
            s = detect_script(val)
            if s in ('latin', 'cyrillic'):
                return s
    azolar = group.get('azolar', [])
    if azolar:
        s = detect_script(azolar[0].get('fio', ''))
        if s in ('latin', 'cyrillic'):
            return s
    return 'latin'


def _set_landscape(doc):
    section = doc.sections[0]
    section.page_width    = Cm(29.7)
    section.page_height   = Cm(21.0)
    section.top_margin    = Cm(0.8)
    section.bottom_margin = Cm(0.8)
    section.left_margin   = Cm(2.0)
    section.right_margin  = Cm(2.0)
    sectPr = section._sectPr
    pgSz = sectPr.find(qn('w:pgSz'))
    if pgSz is None:
        pgSz = OxmlElement('w:pgSz')
        sectPr.append(pgSz)
    pgSz.set(qn('w:orient'), 'landscape')


def _set_cell_bg(cell, hex_color):
    tc = cell._tc; tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)


def _set_cell_border(cell, color='555555'):
    tc = cell._tc; tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for edge in ('top', 'left', 'bottom', 'right'):
        b = OxmlElement(f'w:{edge}')
        b.set(qn('w:val'), 'single')
        b.set(qn('w:sz'), '4')
        b.set(qn('w:space'), '0')
        b.set(qn('w:color'), color)
        tcBorders.append(b)
    tcPr.append(tcBorders)


def _p(doc, text='', bold=False, size=10,
       align=WD_ALIGN_PARAGRAPH.LEFT,
       color=None, before=0, after=2, italic=False):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after  = Pt(after)
    p.paragraph_format.line_spacing = Pt(13)
    if text:
        r = p.add_run(text)
        r.bold = bold; r.italic = italic
        r.font.size = Pt(size); r.font.name = 'Times New Roman'
        if color:
            r.font.color.rgb = RGBColor(*color)
    return p


def _cell_p(cell, text='', bold=False, size=9,
            align=WD_ALIGN_PARAGRAPH.CENTER, color=None):
    for p in list(cell.paragraphs):
        p._element.getparent().remove(p._element)
    p = cell.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(0)
    p.paragraph_format.line_spacing = Pt(11)
    if text:
        r = p.add_run(str(text))
        r.bold = bold; r.font.size = Pt(size); r.font.name = 'Times New Roman'
        if color:
            r.font.color.rgb = RGBColor(*color)
    return p


def _safe_filename(text: str) -> str:
    text = text.replace('\n', '_').replace('\r', '').replace('\t', '_')
    for ch in r'\/:*?"<>|':
        text = text.replace(ch, '-')
    return '_'.join(text.split())[:60] or 'doc'


# ════════════════════════════════════════════════════
#  ПОСТРОИТЕЛЬ ДОКУМЕНТА
# ════════════════════════════════════════════════════

def _build_doc(groups: list, doc_number: str, lang: str) -> Document:
    """
    Строит один альбомный документ для списка групп.
    Одна группа → одна строка в таблице + общее заключение.
    """
    T = TEXTS['cyrillic' if lang == 'cyrillic' else 'latin']
    M = MONTHS[lang]

    first    = groups[0]
    org_name = convert_to(first.get('org_name', ''),  lang)
    direktor = convert_to(first.get('direktor', ''),  lang)
    boshliq  = convert_to(first.get('boshliq', ''),   lang)
    buyruq   = first.get('buyruq', '___')

    today    = datetime.now()
    date_str = first.get('doc_date', today.strftime('%d.%m.%Y'))
    try:
        dt = datetime.strptime(date_str, '%d.%m.%Y')
        day = str(dt.day); month = M[dt.month]; year = str(dt.year)
    except Exception:
        day = str(today.day); month = M[today.month]; year = str(today.year)

    doc = Document()
    _set_landscape(doc)
    doc.styles['Normal'].font.name = 'Times New Roman'
    doc.styles['Normal'].font.size = Pt(10)

    # ══════════════════════
    #  ШАПКА — справа
    # ══════════════════════
    p_hdr = doc.add_paragraph()
    p_hdr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_hdr.paragraph_format.space_before = Pt(0)
    p_hdr.paragraph_format.space_after  = Pt(0)

    def rr(p, text, bold=False, size=10):
        r = p.add_run(text)
        r.bold = bold; r.font.size = Pt(size)
        r.font.name = 'Times New Roman'

    rr(p_hdr, f'{T["tasdiq"]}\n', bold=True)
    # Директор или начальник отдела
    if direktor:
        rr(p_hdr, f'{org_name} {T["rahbar"]}\n')
        rr(p_hdr, f'___________ {direktor}\n')
    elif boshliq:
        rr(p_hdr, f'{org_name}\n')
        rr(p_hdr, f'___________ {boshliq}\n')
    else:
        rr(p_hdr, f'{org_name}\n')
        rr(p_hdr, '___________ _______________\n')
    rr(p_hdr, f'{year} й. «_____»  ____________')

    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    # ══════════════════════
    #  ЗАГОЛОВОК
    # ══════════════════════
    _p(doc, T['title1'], bold=True, size=12,
       align=WD_ALIGN_PARAGRAPH.CENTER, before=2, after=0)
    _p(doc, f'№ {doc_number} - СОН {T["title2"]}',
       bold=True, size=12,
       align=WD_ALIGN_PARAGRAPH.CENTER, before=0, after=4)

    # ══════════════════════
    #  ТАБЛИЦА
    # ══════════════════════
    col_widths = [Cm(3.0), Cm(2.0), Cm(2.0), Cm(6.5), Cm(3.0),
                  Cm(2.8), Cm(3.2), Cm(2.2), Cm(2.8)]
    col_keys   = ['bulim','debet','kredit','nomi','narxi',
                  'inv','amort','norma','qoldiq']
    HDR_FILL   = '2E4057'

    table = doc.add_table(rows=1, cols=len(col_widths))
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Заголовок таблицы
    hdr_row = table.rows[0]
    hdr_row.height = Twips(900)
    for ci, (w, key) in enumerate(zip(col_widths, col_keys)):
        cell = hdr_row.cells[ci]
        cell.width = w
        _set_cell_bg(cell, HDR_FILL)
        _set_cell_border(cell, HDR_FILL)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        _cell_p(cell, T[key], bold=True, size=8,
                align=WD_ALIGN_PARAGRAPH.CENTER, color=(255, 255, 255))

    # Строки данных — по одной на каждую группу
    all_names  = []
    all_groups = []

    for group in groups:
        inv      = group.get('inv_number', '')
        devices  = group.get('devices', [])
        dev_names = [d.get('name', '').strip() for d in devices if d.get('name', '').strip()]
        names_str = ', '.join(dev_names)
        if names_str:
            all_names.append(names_str)
        all_groups.append(group)

        bulim_v = convert_to(group.get('bulim', org_name), lang)

        data_row = table.add_row()
        data_row.height = Twips(400)
        row_values = [
            bulim_v,
            group.get('debet', '9210'),
            group.get('kredit', '0150'),
            names_str,
            group.get('narxi', ''),
            inv,
            group.get('amort', ''),
            group.get('norma', ''),
            group.get('qoldiq', ''),
        ]
        for ci, (w, val) in enumerate(zip(col_widths, row_values)):
            cell = data_row.cells[ci]
            cell.width = w
            _set_cell_border(cell)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            _cell_p(cell, val, size=9,
                    align=WD_ALIGN_PARAGRAPH.LEFT if ci in (0, 3)
                    else WD_ALIGN_PARAGRAPH.CENTER)

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # ══════════════════════
    #  ЗАКЛЮЧЕНИЕ
    # ══════════════════════
    # Комиссия
    comm = T['komissiya'].format(
        org=org_name, date=year, buyruq=buyruq or '___')
    _p(doc, comm, size=10, before=0, after=3)

    # История — берём из первой группы
    g0           = all_groups[0]
    ish_chiq_yil = g0.get('ish_chiq_yil', '')
    kelgan_yil   = g0.get('kelgan_yil', '')
    ishga_yil    = g0.get('ishga_yil', '')
    tamir_soni   = g0.get('tamir_soni', '_____')
    tamir_summa  = g0.get('tamir_summa', '_________')

    for yr, key in [
        (ish_chiq_yil, 'ish_chiq'),
        (kelgan_yil,   'kelgan'),
        (ishga_yil,    'ishga'),
    ]:
        _p(doc, T[key].format(year=yr or '____'), size=10, before=0, after=1)

    _p(doc, T['kapital'].format(
        soni=tamir_soni or '_____',
        summa=tamir_summa or '_________'),
       size=10, before=0, after=1)
    _p(doc, T['texnik'], size=10, before=0, after=2)

    # Хулоса
    xulosa = T['xulosa'].format(
        org=org_name,
        names='; '.join(all_names),
        date=year, day=day, month=month)
    _p(doc, xulosa, size=10, before=0, after=6)

    # ══════════════════════
    #  ПОДПИСИ
    # ══════════════════════
    # Моддий жавобгар
    p_mod = doc.add_paragraph()
    p_mod.paragraph_format.space_before = Pt(0)
    p_mod.paragraph_format.space_after  = Pt(4)
    r1 = p_mod.add_run(f'{T["moddiy"]}  _______________ ')
    r1.font.size = Pt(10); r1.font.name = 'Times New Roman'
    r2 = p_mod.add_run(boshliq or direktor or '')
    r2.bold = True; r2.font.size = Pt(10); r2.font.name = 'Times New Roman'

    # Комиссия раиси (директор)
    p_rais = doc.add_paragraph()
    p_rais.paragraph_format.space_before = Pt(0)
    p_rais.paragraph_format.space_after  = Pt(3)
    rais_name = direktor or boshliq or ''
    r1 = p_rais.add_run(T['komrais'].format(org=org_name))
    r1.font.size = Pt(10); r1.font.name = 'Times New Roman'
    r2 = p_rais.add_run(f'                    _____________ {rais_name}')
    r2.font.size = Pt(10); r2.font.name = 'Times New Roman'

    # Члены комиссии — динамически из azolar
    azolar = g0.get('azolar', [])
    if azolar:
        p_az = doc.add_paragraph()
        p_az.paragraph_format.space_before = Pt(0)
        p_az.paragraph_format.space_after  = Pt(2)
        p_az.add_run(T['komazolari']).font.name = 'Times New Roman'

        for azo in azolar:
            fio     = convert_to(azo.get('fio', ''),     lang)
            lavozim = convert_to(azo.get('lavozim', ''), lang)
            if not fio:
                continue
            p_e = doc.add_paragraph()
            p_e.paragraph_format.space_before = Pt(0)
            p_e.paragraph_format.space_after  = Pt(2)
            # Должность — если есть
            label = f'{lavozim}' if lavozim else ''
            r_lbl = p_e.add_run(f'{label:<40}')
            r_lbl.font.size = Pt(10); r_lbl.font.name = 'Times New Roman'
            r_line = p_e.add_run('____________  ')
            r_line.font.size = Pt(10); r_line.font.name = 'Times New Roman'
            r_name = p_e.add_run(fio)
            r_name.font.size = Pt(10); r_name.font.name = 'Times New Roman'

    return doc


# ════════════════════════════════════════════════════
#  ПУБЛИЧНЫЕ ФУНКЦИИ СОХРАНЕНИЯ
# ════════════════════════════════════════════════════

def create_spisan_doc(group: dict, doc_number: str = '1',
                      lang: str = 'auto') -> Document:
    """Создаёт один документ для одной группы."""
    if lang == 'auto':
        lang = _auto_lang(group)
    return _build_doc([group], doc_number, lang)


def save_single_files(groups: list, output_dir: str,
                      lang: str = 'auto') -> list:
    """Один .docx на каждую группу (инвентарный номер)."""
    os.makedirs(output_dir, exist_ok=True)
    created = []
    for i, group in enumerate(groups, start=1):
        effective_lang = lang if lang != 'auto' else _auto_lang(group)
        inv  = _safe_filename(group.get('inv_number', str(i)))
        doc  = _build_doc([group], str(i), effective_lang)
        name = f'Akt_Spisaniya_{inv}.docx'
        path = os.path.join(output_dir, name)
        doc.save(path)
        created.append(path)
    return created


def save_all_in_one(groups: list, output_path: str,
                    lang: str = 'auto') -> str:
    """Все группы в одном документе — единая таблица и заключение."""
    if not groups:
        return output_path
    effective_lang = lang if lang != 'auto' else _auto_lang(groups[0])
    doc = _build_doc(groups, '1', effective_lang)
    doc.save(output_path)
    return output_path
