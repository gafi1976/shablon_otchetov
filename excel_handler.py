# excel_handler.py
# Один файл shablon.xlsx с двумя листами:
#   Лист 1: "Spisaniye"  — данные для акта СПИСАНИЯ
#   Лист 2: "Ustanovka"  — данные для акта УСТАНОВКИ
#
# ════════════════════════════════════════════════════
#  ЛИСТ "Spisaniye" — структура (строка 1 = заголовки)
# ════════════════════════════════════════════════════
#  A  Tashkilot       — Название организации
#  B  Rahbar          — Руководитель (ФИО)
#  C  Muhandis1       — Инженер 1 (ФИО)
#  D  Muhandis2       — Инженер 2 (ФИО)
#  E  Inventar        — Инвентарный номер  (ключ группировки)
#  F  Qurilma         — Название оборудования
#  G  Qism            — Компонент
#  H  Holat           — Состояние: Yaroqli / Yaroqsiz
#  I  Sabab           — Причина / неисправность
#
# ════════════════════════════════════════════════════
#  ЛИСТ "Ustanovka" — структура (строка 1 = заголовки)
# ════════════════════════════════════════════════════
#  A  Tashkilot       — Название организации
#  B  Manzil          — Адрес организации
#  C  Rahbar          — Руководитель (ФИО)
#  D  Muhandis1       — Инженер 1 (ФИО)
#  E  Lavozim1        — Должность инженера 1
#  F  Muhandis2       — Инженер 2 (ФИО)
#  G  Lavozim2        — Должность инженера 2
#  H  Sana            — Дата установки (дд.мм.гггг)
#  I  Uskuna          — Название / модель оборудования
#  J  Seriya          — Серийный номер
#  K  Joyi            — Место установки

import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime, date, timedelta

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation


# ─────────────────────────────────────────────────
#  Стили
# ─────────────────────────────────────────────────

_FONT_HDR  = Font(name='Times New Roman', bold=True, color='FFFFFF', size=11)
_FONT_DATA = Font(name='Times New Roman', size=10)
_FILL_S    = PatternFill('solid', fgColor='2E4057')   # списание — тёмно-синий
_FILL_U    = PatternFill('solid', fgColor='1A5276')   # установка — синий
_FILL_ES   = PatternFill('solid', fgColor='EBF5FB')
_FILL_ODD  = PatternFill('solid', fgColor='FFFFFF')
_ALIGN_C   = Alignment(horizontal='center', vertical='center', wrap_text=True)
_ALIGN_L   = Alignment(horizontal='left',   vertical='center', wrap_text=True)


def _border():
    s = Side(style='thin', color='AAAAAA')
    return Border(left=s, right=s, top=s, bottom=s)


def _make_sheet(ws, title_text, title_color, headers_widths, hdr_fill):
    """
    Создаёт лист с:
      Строка 1 — цветная строка-заголовок листа (название)
      Строка 2 — заголовки колонок
      Строки 3+ — пустые строки для данных
    """
    ws.sheet_view.showGridLines = False

    # Строка 1 — заголовок листа
    ncols = len(headers_widths)
    ws.merge_cells(start_row=1, start_column=1,
                   end_row=1,   end_column=ncols)
    cell = ws.cell(row=1, column=1, value=title_text)
    cell.font      = Font(name='Times New Roman', bold=True, size=14, color='FFFFFF')
    cell.fill      = PatternFill('solid', fgColor=title_color)
    cell.alignment = _ALIGN_C
    ws.row_dimensions[1].height = 28

    # Строка 2 — заголовки колонок
    ws.row_dimensions[2].height = 38
    for ci, (header, width) in enumerate(headers_widths, start=1):
        from openpyxl.utils import get_column_letter
        ws.column_dimensions[get_column_letter(ci)].width = width
        cell = ws.cell(row=2, column=ci, value=header)
        cell.font      = _FONT_HDR
        cell.fill      = hdr_fill
        cell.alignment = _ALIGN_C
        cell.border    = _border()

    # Строки 3–52 — пустые для ввода данных (50 строк)
    for ri in range(3, 53):
        ws.row_dimensions[ri].height = 20
        fill = _FILL_ES if ri % 2 == 0 else _FILL_ODD
        for ci in range(1, ncols + 1):
            cell = ws.cell(row=ri, column=ci)
            cell.font      = _FONT_DATA
            cell.fill      = fill
            cell.border    = _border()
            cell.alignment = _ALIGN_C if ci in (5, 8) else _ALIGN_L

    ws.freeze_panes = 'A3'


# ─────────────────────────────────────────────────
#  СОЗДАНИЕ shablon.xlsx
# ─────────────────────────────────────────────────

SHABLON_NAME = 'shablon.xlsx'

# Заголовки листа Spisaniye (9 колонок)
HEADERS_S = [
    ('Tashkilot\n(Организация)',          28),
    ('Rahbar\n(Руководитель)',             22),
    ('Muhandis1\n(Инженер 1)',             22),
    ('Muhandis2\n(Инженер 2)',             22),
    ('Inventar\n(Инв. номер)',             16),
    ('Qurilma\n(Оборудование)',            26),
    ('Qism\n(Компонент)',                  22),
    ('Holat\n(Yaroqli / Yaroqsiz)',        16),
    ('Sabab\n(Причина / неисправность)',   28),
]

# Заголовки листа Ustanovka (11 колонок)
HEADERS_U = [
    ('Tashkilot\n(Организация)',           28),
    ('Manzil\n(Адрес)',                    22),
    ('Rahbar\n(Руководитель)',             22),
    ('Muhandis1\n(Инженер 1)',             22),
    ('Lavozim1\n(Должность 1)',            20),
    ('Muhandis2\n(Инженер 2)',             22),
    ('Lavozim2\n(Должность 2)',            20),
    ('Sana\n(Дата дд.мм.гггг)',            16),
    ('Uskuna\n(Оборудование / Модель)',    26),
    ('Seriya\n(Серийный номер)',           18),
    ("Joyi\n(Место установки)",           22),
]


def create_template(path: str) -> str:
    """
    Создаёт shablon.xlsx с двумя листами:
      Лист "Spisaniye" — для акта СПИСАНИЯ
      Лист "Ustanovka" — для акта УСТАНОВКИ
    Только заголовки, пустые строки для заполнения.
    """
    wb = Workbook()

    # Лист 1 — Spisaniye
    ws_s = wb.active
    ws_s.title = 'Spisaniye'
    _make_sheet(ws_s,
                'АКТ СПИСАНИЯ — заполните данные начиная с строки 3',
                '2E4057', HEADERS_S, _FILL_S)

    # Валидация колонки H (Holat)
    dv = DataValidation(type='list',
                        formula1='"Yaroqli,Yaroqsiz"',
                        allow_blank=True, showDropDown=False)
    dv.sqref = 'H3:H200'
    ws_s.add_data_validation(dv)

    # Лист 2 — Ustanovka
    ws_u = wb.create_sheet(title='Ustanovka')
    _make_sheet(ws_u,
                'АКТ УСТАНОВКИ — заполните данные начиная с строки 3',
                '1A5276', HEADERS_U, _FILL_U)

    wb.save(path)
    return path


# ─────────────────────────────────────────────────
#  Для обратной совместимости (вызывается из app.py)
# ─────────────────────────────────────────────────

def create_spisan_template(path: str) -> str:
    """Создаёт полный shablon.xlsx (оба листа) в папке path."""
    import os
    folder = os.path.dirname(path)
    out    = os.path.join(folder, SHABLON_NAME)
    return create_template(out)


def create_ust_template(path: str) -> str:
    """Создаёт полный shablon.xlsx (оба листа) в папке path."""
    import os
    folder = os.path.dirname(path)
    out    = os.path.join(folder, SHABLON_NAME)
    return create_template(out)


# ─────────────────────────────────────────────────
#  Низкоуровневый читатель xlsx
# ─────────────────────────────────────────────────

def _read_sheet(path: str, sheet_index: int = 0) -> list:
    """
    Читает указанный лист xlsx-файла (по индексу).
    Возвращает list строк; строка 0 = строка 1 Excel.
    Каждая строка — list строковых значений ячеек.
    """
    with zipfile.ZipFile(path) as z:
        names = z.namelist()

        # Shared strings
        strings = []
        if 'xl/sharedStrings.xml' in names:
            ns = {'n': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
            for si in ET.parse(z.open('xl/sharedStrings.xml')).findall('.//n:si', ns):
                strings.append(''.join(t.text or '' for t in si.findall('.//n:t', ns)))

        # Все листы
        sheets = sorted(f for f in names if f.startswith('xl/worksheets/sheet'))
        if sheet_index >= len(sheets):
            return []

        ns = {'n': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        all_rows = ET.parse(z.open(sheets[sheet_index])).findall('.//n:row', ns)
        if not all_rows:
            return []

        def coln(ref: str) -> int:
            col = ''.join(c for c in ref if c.isalpha())
            n = 0
            for ch in col.upper():
                n = n * 26 + (ord(ch) - ord('A') + 1)
            return n

        max_row = max(int(r.get('r', 0)) for r in all_rows)
        result  = [[] for _ in range(max_row + 1)]

        for row_el in all_rows:
            rn    = int(row_el.get('r', 0))
            cells = row_el.findall('n:c', ns)
            if not cells:
                continue
            mc  = max(coln(c.get('r', 'A')) for c in cells)
            row = [''] * mc
            for c in cells:
                ci = coln(c.get('r', 'A1')) - 1
                t  = c.get('t', '')
                v  = c.find('n:v', ns)
                if v is not None and v.text is not None:
                    row[ci] = strings[int(v.text)] if t == 's' else v.text
            result[rn] = row

        return result[1:]   # индекс 0 пустой; result[0] = строка 1 Excel


def _g(row: list, idx: int) -> str:
    """Безопасно возвращает ячейку как строку, убирает переносы строк."""
    if idx >= len(row):
        return ''
    v = row[idx]
    if not v:
        return ''
    # Убираем переносы строк из значений ячеек
    return str(v).strip().replace('\n', ' ').replace('\r', '').replace('\t', ' ')


def _to_date(val: str) -> str:
    """Excel serial number → дд.мм.гггг; строку оставляет как есть."""
    if not val:
        return ''
    val = val.strip()
    # Уже в нужном формате
    if len(val) == 10 and val[2] == '.' and val[5] == '.':
        return val
    # Excel serial number
    try:
        n = float(val)
        if n > 1000:
            return (date(1899, 12, 30) + timedelta(days=int(n))).strftime('%d.%m.%Y')
    except ValueError:
        pass
    return val


def _get_sheet_index(path: str, sheet_name: str) -> int:
    """Возвращает индекс листа по его имени (из workbook.xml)."""
    try:
        with zipfile.ZipFile(path) as z:
            ns = {'n': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
            tree = ET.parse(z.open('xl/workbook.xml'))
            sheets = tree.findall('.//n:sheet', ns)
            for i, s in enumerate(sheets):
                name = s.get('name', '') or s.get(
                    '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}name', '')
                # openpyxl сохраняет имя в атрибуте 'name'
                if name.lower() == sheet_name.lower():
                    return i
    except Exception:
        pass
    return 0


# ─────────────────────────────────────────────────
#  ЧТЕНИЕ ЛИСТА "Spisaniye"
# ─────────────────────────────────────────────────

def read_spisan_excel(path: str) -> list:
    """
    Читает лист "Spisaniye" из файла path.
    Строка 1 = заголовки, данные с строки 2.
    Колонки: A=Tashkilot, B=Rahbar, C=Muhandis1, D=Muhandis2,
             E=Inventar, F=Qurilma, G=Qism, H=Holat, I=Sabab.
    Возвращает список групп (по колонке E Inventar).
    """
    idx  = _get_sheet_index(path, 'Spisaniye')
    rows = _read_sheet(path, idx)
    if not rows:
        raise ValueError(f'Лист "Spisaniye" пуст или не найден: {path}')

    inv_data  = {}
    inv_order = []

    for row in rows[1:]:          # строка 0 = заголовки → пропускаем
        if not row or not any(row):
            continue

        tashkilot = _g(row, 0)   # A
        rahbar    = _g(row, 1)   # B
        muh1      = _g(row, 2)   # C
        muh2      = _g(row, 3)   # D
        inventar  = _g(row, 4)   # E  ← ключ
        qurilma   = _g(row, 5)   # F
        qism      = _g(row, 6)   # G
        holat     = _g(row, 7)   # H
        sabab     = _g(row, 8)   # I

        if not inventar:
            continue

        if inventar not in inv_data:
            inv_data[inventar] = {
                'inv_number':      inventar,
                'org_name':        tashkilot,
                'region':          '',
                'doc_date':        datetime.now().strftime('%d.%m.%Y'),
                'commission_head': rahbar,
                'member1':         muh1,
                'member2':         muh2,
                'devices':         {},
                'dev_order':       [],
            }
            inv_order.append(inventar)

        g = inv_data[inventar]
        if not g['org_name']        and tashkilot: g['org_name']        = tashkilot
        if not g['commission_head'] and rahbar:    g['commission_head'] = rahbar
        if not g['member1']         and muh1:      g['member1']         = muh1
        if not g['member2']         and muh2:      g['member2']         = muh2

        if qurilma:
            if qurilma not in g['devices']:
                g['devices'][qurilma] = []
                g['dev_order'].append(qurilma)
            if qism:
                g['devices'][qurilma].append({
                    'part_name': qism,
                    'condition': holat,
                    'defect':    sabab,
                })

    result = []
    for inv in inv_order:
        g = inv_data[inv]
        result.append({
            'inv_number':      g['inv_number'],
            'org_name':        g['org_name'],
            'region':          '',
            'doc_date':        g['doc_date'],
            'commission_head': g['commission_head'],
            'member1':         g['member1'],
            'member2':         g['member2'],
            'devices': [
                {'name': dn, 'parts': g['devices'][dn]}
                for dn in g['dev_order']
            ],
        })
    return result


# ─────────────────────────────────────────────────
#  ЧТЕНИЕ ЛИСТА "Ustanovka"
# ─────────────────────────────────────────────────

def read_ust_excel(path: str) -> dict:
    """
    Читает лист "Ustanovka" из файла path.
    Строка 1 = заголовки, данные с строки 2.
    Колонки: A=Tashkilot, B=Manzil, C=Rahbar, D=Muhandis1, E=Lavozim1,
             F=Muhandis2, G=Lavozim2, H=Sana, I=Uskuna, J=Seriya, K=Joyi.
    Возвращает dict с полем items (список оборудования).
    """
    idx  = _get_sheet_index(path, 'Ustanovka')
    rows = _read_sheet(path, idx)
    if not rows:
        raise ValueError(f'Лист "Ustanovka" пуст или не найден: {path}')

    org = ''; addr = ''; boss = ''
    e1  = ''; j1   = ''; e2   = ''; j2 = ''
    items = []

    for row in rows[1:]:          # строка 0 = заголовки → пропускаем
        if not row or not any(row):
            continue

        tashkilot = _g(row,  0)   # A
        manzil    = _g(row,  1)   # B
        rahbar    = _g(row,  2)   # C
        muh1      = _g(row,  3)   # D
        lav1      = _g(row,  4)   # E
        muh2      = _g(row,  5)   # F
        lav2      = _g(row,  6)   # G
        sana      = _g(row,  7)   # H
        uskuna    = _g(row,  8)   # I
        seriya    = _g(row,  9)   # J
        joyi      = _g(row, 10)   # K

        if not org  and tashkilot: org  = tashkilot
        if not addr and manzil:    addr = manzil
        if not boss and rahbar:    boss = rahbar
        if not e1   and muh1:      e1   = muh1; j1 = lav1
        if not e2   and muh2:      e2   = muh2; j2 = lav2

        if not uskuna:
            continue

        items.append({
            'num':           str(len(items) + 1),
            'name':          uskuna,
            'model':         uskuna,
            'serial_number': seriya,
            'inv_number':    '',
            'install_date':  _to_date(sana) or datetime.now().strftime('%d.%m.%Y'),
            'location':      joyi,
            'cost':          '',
            'condition':     'Yangi',
            'note':          joyi,
        })

    doc_date = items[0]['install_date'] if items else datetime.now().strftime('%d.%m.%Y')

    return {
        'org_name':        org,
        'region':          org,
        'address':         addr,
        'doc_number':      '1',
        'doc_date':        doc_date,
        'commission_head': boss,
        'member1':         e1,
        'member1_title':   j1 or 'Yetakchi muhandis',
        'member2':         e2,
        'member2_title':   j2 or 'Yetakchi muhandis',
        'items':           items,
    }
