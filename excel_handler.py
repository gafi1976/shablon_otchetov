# excel_handler.py
# Создание и чтение Excel шаблонов для актов СПИСАНИЯ и УСТАНОВКИ
#
# ════════════════════════════════════════════════════════════
#  ФОРМАТ ШАБЛОНА СПИСАНИЯ (shablon_spisan.xlsx)
# ════════════════════════════════════════════════════════════
#  Строка 1 — ЗАГОЛОВКИ:
#   A: Tashkilot        — Организация
#   B: Rahbar           — Руководитель (ФИО)
#   C: Muhandis1        — Инженер 1 (ФИО)
#   D: Muhandis2        — Инженер 2 (ФИО)
#   E: Inventar         — Инвентарный номер  ← ключ группировки
#   F: Qurilma          — Название оборудования
#   G: Qism             — Компонент
#   H: Holat            — Состояние (Yaroqli / Yaroqsiz)
#   I: Sabab            — Причина / неисправность
#
#  Строки 2+ — данные. Пустые строки игнорируются.
#  Группировка по E (Inventar) → один Word файл на инв. номер.
#
# ════════════════════════════════════════════════════════════
#  ФОРМАТ ШАБЛОНА УСТАНОВКИ (shablom_ust.xlsx)
# ════════════════════════════════════════════════════════════
#  Строка 1 — ЗАГОЛОВКИ:
#   A: Tashkilot        — Организация
#   B: Manzil           — Адрес
#   C: Rahbar           — Руководитель (ФИО)
#   D: Muhandis1        — Инженер 1 (ФИО)
#   E: Lavozim1         — Должность инженера 1
#   F: Muhandis2        — Инженер 2 (ФИО)
#   G: Lavozim2         — Должность инженера 2
#   H: Sana             — Дата установки (дд.мм.гггг)
#   I: Nom              — Наименование оборудования / модель
#   J: Seriya           — Серийный номер
#   K: Joyi             — Место установки
#
#  Строки 2+ — данные. Каждая строка = одна единица оборудования.
# ════════════════════════════════════════════════════════════

import os
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime, date, timedelta

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation


# ─── Стили ──────────────────────────────────────────────────

_HDR_S   = PatternFill('solid', fgColor='2E4057')
_HDR_U   = PatternFill('solid', fgColor='1A5276')
_HFONT   = Font(name='Times New Roman', bold=True, color='FFFFFF', size=10)
_DFONT   = Font(name='Times New Roman', size=10)
_RFONT_S = Font(name='Times New Roman', size=10, color='C0392B', bold=True)
_RFONT_G = Font(name='Times New Roman', size=10, color='1E8B4C', bold=True)
_CTR     = Alignment(horizontal='center', vertical='center', wrap_text=True)
_LFT     = Alignment(horizontal='left',   vertical='center', wrap_text=True)
_ES      = PatternFill('solid', fgColor='F0F4F8')
_EU      = PatternFill('solid', fgColor='EBF5FB')
_ODD     = PatternFill('solid', fgColor='FFFFFF')


def _brd(color='2E4057'):
    s = Side(style='thin', color=color)
    return Border(left=s, right=s, top=s, bottom=s)


def _title_row(ws, text, ncols, fill_color):
    """Строка-заголовок листа (строка 1)."""
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    cell = ws.cell(row=1, column=1, value=text)
    cell.font      = Font(name='Times New Roman', bold=True, size=13, color='FFFFFF')
    cell.fill      = PatternFill('solid', fgColor=fill_color)
    cell.alignment = _CTR
    ws.row_dimensions[1].height = 24


def _header_row(ws, headers, hdr_fill, row=2):
    """Строка заголовков колонок."""
    ws.row_dimensions[row].height = 36
    for ci, (text, width) in enumerate(headers, start=1):
        from openpyxl.utils import get_column_letter
        ws.column_dimensions[get_column_letter(ci)].width = width
        cell = ws.cell(row=row, column=ci, value=text)
        cell.font = _HFONT; cell.fill = hdr_fill
        cell.alignment = _CTR; cell.border = _brd()


# ════════════════════════════════════════════════════════════
#  СОЗДАНИЕ ШАБЛОНА СПИСАНИЯ
# ════════════════════════════════════════════════════════════

def create_spisan_template(path: str) -> str:
    wb = Workbook()
    ws = wb.active
    ws.title = 'Spisaniye'
    ws.sheet_view.showGridLines = False

    _title_row(ws, 'АКТ СПИСАНИЯ — шаблон данных', 9, '2E4057')

    headers = [
        ('Tashkilot\n(Организация)',          28),
        ('Rahbar\n(Руководитель)',             22),
        ('Muhandis 1\n(Инженер 1)',            22),
        ('Muhandis 2\n(Инженер 2)',            22),
        ('Inventar ★\n(Инв. номер)',           16),
        ('Qurilma nomi\n(Оборудование)',       26),
        ('Qism nomi\n(Компонент)',             22),
        ('Holat\n(Yaroqli/Yaroqsiz)',          16),
        ('Nosozlik sababi\n(Причина)',         26),
    ]
    _header_row(ws, headers, _HDR_S)

    # Валидация для колонки H
    dv = DataValidation(type='list', formula1='"Yaroqli,Yaroqsiz"',
                        allow_blank=True, showDropDown=False)
    dv.sqref = 'H3:H500'
    ws.add_data_validation(dv)

    # Примеры данных (строки 3+)
    ORG = 'ABM Sirdaryo viloyati MChJ'
    B   = 'Palonov P.A'
    E1  = 'Raxmatov V.A'
    E2  = 'Xolbekov G.T'
    rows = [
        (ORG, B, E1, E2, '26-0006039', 'Server HP ML350',        'Asosiy plata', 'Yaroqli',  "Ma'naviy eskirgan"),
        (ORG, B, E1, E2, '26-0006039', 'Server HP ML350',        'Tok manbayi',  'Yaroqli',  "Ma'naviy eskirgan"),
        (ORG, B, E1, E2, '26-0006039', 'Server HP ML350',        'Qattiq disk',  'Yaroqsiz', 'BAD bloklar mavjud'),
        (ORG, B, E1, E2, '24-0006162', 'Konditsioner ART-12HI',  'Tok manbai',   'Yaroqsiz', "Tok sarfi ko'p"),
        (ORG, B, E1, E2, '24-0006162', 'Konditsioner ART-12HI',  'Kompressor',   'Yaroqsiz', 'Shovqin mavjud'),
        (ORG, B, E1, E2, '24-0006162', 'Konditsioner ART-12HI',  'Radiator',     'Yaroqsiz', 'Yamalgan'),
        (ORG, B, E1, E2, '26-0003436', "Monitor LCD 19''",       'Asosiy plata', 'Yaroqsiz', 'Korpus singan'),
        (ORG, B, E1, E2, '26-0003436', "Monitor LCD 19''",       'Tok manbayi',  'Yaroqsiz', 'Kuygan'),
    ]

    for ri, row in enumerate(rows):
        rn   = ri + 3
        fill = _ES if ri % 2 == 0 else _ODD
        ws.row_dimensions[rn].height = 18
        for ci, val in enumerate(row, start=1):
            cell = ws.cell(row=rn, column=ci, value=val)
            cell.fill = fill; cell.border = _brd()
            cell.alignment = _CTR if ci in (5, 8) else _LFT
            if ci == 8:
                cell.font = _RFONT_S if val == 'Yaroqsiz' else _RFONT_G
            else:
                cell.font = _DFONT

    # 20 пустых строк
    for ri in range(len(rows), len(rows) + 20):
        rn   = ri + 3
        fill = _ES if ri % 2 == 0 else _ODD
        ws.row_dimensions[rn].height = 18
        for ci in range(1, 10):
            cell = ws.cell(row=rn, column=ci)
            cell.font = _DFONT; cell.fill = fill
            cell.border = _brd(); cell.alignment = _LFT

    # Инструкция
    ir = len(rows) + 24
    ws.merge_cells(start_row=ir, start_column=1, end_row=ir, end_column=9)
    c = ws.cell(row=ir, column=1,
        value='★ Каждая строка = один компонент. '
              'Группировка в Word-файл по колонке E (Inventar). '
              'Holat: Yaroqli = исправен, Yaroqsiz = неисправен.')
    c.font = Font(name='Times New Roman', italic=True, size=9, color='888888')
    c.alignment = _LFT
    ws.row_dimensions[ir].height = 16

    ws.freeze_panes = 'A3'
    wb.save(path)
    return path


# ════════════════════════════════════════════════════════════
#  СОЗДАНИЕ ШАБЛОНА УСТАНОВКИ
# ════════════════════════════════════════════════════════════

def create_ust_template(path: str) -> str:
    wb = Workbook()
    ws = wb.active
    ws.title = 'Ustanovka'
    ws.sheet_view.showGridLines = False

    _title_row(ws, 'АКТ УСТАНОВКИ — шаблон данных', 11, '1A5276')

    today = datetime.now().strftime('%d.%m.%Y')
    headers = [
        ('Tashkilot\n(Организация)',             28),
        ('Manzil\n(Адрес)',                       22),
        ('Rahbar\n(Руководитель)',                22),
        ('Muhandis 1\n(Инженер 1)',               22),
        ('Lavozim 1\n(Должность 1)',              18),
        ('Muhandis 2\n(Инженер 2)',               22),
        ('Lavozim 2\n(Должность 2)',              18),
        ('Sana\n(Дата  дд.мм.гггг)',              16),
        ('Uskuna nomi\n(Оборудование / Модель)',  26),
        ('Seriya raqami\n(Серийный номер)',        18),
        ("O'rnatish joyi\n(Место установки)",     22),
    ]
    _header_row(ws, headers, _HDR_U)

    ORG  = 'ABM Sirdaryo viloyati MChJ'
    ADDR = "Guliston sh., O'zbekiston ko'ch."
    B    = 'Palonov P.A'
    E1   = 'Raxmatov V.A';  L1 = 'Yetakchi muhandis'
    E2   = 'Xolbekov G.T';  L2 = 'Yetakchi muhandis'
    rows = [
        (ORG, ADDR, B, E1, L1, E2, L2, today, 'Maipu MP1900X-22',   'SN-001', 'Guliston, 1-bino'),
        (ORG, ADDR, B, E1, L1, E2, L2, today, 'Maipu MP1900X-22',   'SN-002', 'Guliston, 2-bino'),
        (ORG, ADDR, B, E1, L1, E2, L2, today, 'Switch Cisco SG350', 'SN-003', 'Guliston, 3-bino'),
        (ORG, ADDR, B, E1, L1, E2, L2, today, 'UPS APC 1500VA',     'SN-004', 'Server xona'),
    ]

    for ri, row in enumerate(rows):
        rn   = ri + 3
        fill = _EU if ri % 2 == 0 else _ODD
        ws.row_dimensions[rn].height = 20
        for ci, val in enumerate(row, start=1):
            cell = ws.cell(row=rn, column=ci, value=val)
            cell.font = _DFONT; cell.fill = fill
            cell.border = _brd('1A5276')
            cell.alignment = _CTR if ci == 8 else _LFT

    for ri in range(len(rows), len(rows) + 20):
        rn   = ri + 3
        fill = _EU if ri % 2 == 0 else _ODD
        ws.row_dimensions[rn].height = 20
        for ci in range(1, 12):
            cell = ws.cell(row=rn, column=ci)
            cell.font = _DFONT; cell.fill = fill
            cell.border = _brd('1A5276'); cell.alignment = _LFT

    ir = len(rows) + 24
    ws.merge_cells(start_row=ir, start_column=1, end_row=ir, end_column=11)
    c = ws.cell(row=ir, column=1,
        value='Каждая строка = одна единица оборудования. '
              'Общие поля (организация, руководитель, инженеры) '
              'одинаковы для всех строк. Дата: дд.мм.гггг.')
    c.font = Font(name='Times New Roman', italic=True, size=9, color='888888')
    c.alignment = _LFT
    ws.row_dimensions[ir].height = 16

    ws.freeze_panes = 'A3'
    wb.save(path)
    return path


# ════════════════════════════════════════════════════════════
#  НИЗКОУРОВНЕВЫЙ ЧИТАТЕЛЬ xlsx (stdlib, без pandas)
# ════════════════════════════════════════════════════════════

def _read_rows(path: str) -> list:
    """Возвращает список строк начиная с индекса 0 = строка 1 Excel."""
    with zipfile.ZipFile(path) as z:
        strings = []
        if 'xl/sharedStrings.xml' in z.namelist():
            ns = {'n': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
            for si in ET.parse(z.open('xl/sharedStrings.xml')).findall('.//n:si', ns):
                strings.append(''.join(t.text or '' for t in si.findall('.//n:t', ns)))

        sheets = sorted(f for f in z.namelist() if f.startswith('xl/worksheets/sheet'))
        if not sheets:
            return []

        ns = {'n': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        all_rows = ET.parse(z.open(sheets[0])).findall('.//n:row', ns)
        if not all_rows:
            return []

        def coln(ref):
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

        return result[1:]   # result[0] пустой, возвращаем с row=1


def _g(row, idx):
    """Безопасно возвращает строку из строки по индексу."""
    v = row[idx] if idx < len(row) else ''
    return str(v).strip() if v else ''


def _to_date(val: str) -> str:
    """Excel serial → дд.мм.гггг или оставляет строку как есть."""
    if not val:
        return ''
    val = val.strip()
    if len(val) == 10 and val[2] == '.' and val[5] == '.':
        return val
    try:
        n = float(val)
        if n > 1000:
            return (date(1899, 12, 30) + timedelta(days=int(n))).strftime('%d.%m.%Y')
    except ValueError:
        pass
    return val


# ════════════════════════════════════════════════════════════
#  ЧТЕНИЕ ШАБЛОНА СПИСАНИЯ
# ════════════════════════════════════════════════════════════

def read_spisan_excel(path: str) -> list:
    """
    Читает shablon_spisan.xlsx.
    Колонки: A=Tashkilot, B=Rahbar, C=Muhandis1, D=Muhandis2,
             E=Inventar★, F=Qurilma, G=Qism, H=Holat, I=Sabab
    Строка 1 = заголовки, данные с строки 2.
    Возвращает список групп (dict) по инвентарному номеру.
    """
    rows = _read_rows(path)
    if not rows:
        raise ValueError(f'Файл пуст: {path}')

    inv_data   = {}   # inv → dict
    inv_order  = []

    for row in rows[1:]:          # строка 0 = заголовки, данные с индекса 1
        if not row or not any(row):
            continue

        org   = _g(row, 0)   # A
        boss  = _g(row, 1)   # B
        eng1  = _g(row, 2)   # C
        eng2  = _g(row, 3)   # D
        inv   = _g(row, 4)   # E  ← ключ
        dev   = _g(row, 5)   # F
        part  = _g(row, 6)   # G
        holat = _g(row, 7)   # H
        sabab = _g(row, 8)   # I

        if not inv:
            continue

        if inv not in inv_data:
            inv_data[inv] = {
                'inv_number':      inv,
                'org_name':        org,
                'region':          '',
                'doc_date':        datetime.now().strftime('%d.%m.%Y'),
                'commission_head': boss,
                'member1':         eng1,
                'member2':         eng2,
                'devices':         {},
                'dev_order':       [],
            }
            inv_order.append(inv)

        g = inv_data[inv]
        # Обновляем общие поля если пустые
        if not g['org_name']        and org:  g['org_name']        = org
        if not g['commission_head'] and boss: g['commission_head'] = boss
        if not g['member1']         and eng1: g['member1']         = eng1
        if not g['member2']         and eng2: g['member2']         = eng2

        if dev:
            if dev not in g['devices']:
                g['devices'][dev] = []
                g['dev_order'].append(dev)
            if part:
                g['devices'][dev].append({
                    'part_name': part,
                    'condition': holat,
                    'defect':    sabab,
                })

    # Преобразуем в финальный формат
    result = []
    for inv in inv_order:
        g = inv_data[inv]
        devices = [
            {'name': dn, 'parts': g['devices'][dn]}
            for dn in g['dev_order']
        ]
        result.append({
            'inv_number':      g['inv_number'],
            'org_name':        g['org_name'],
            'region':          g['region'],
            'doc_date':        g['doc_date'],
            'commission_head': g['commission_head'],
            'member1':         g['member1'],
            'member2':         g['member2'],
            'devices':         devices,
        })
    return result


# ════════════════════════════════════════════════════════════
#  ЧТЕНИЕ ШАБЛОНА УСТАНОВКИ
# ════════════════════════════════════════════════════════════

def read_ust_excel(path: str) -> dict:
    """
    Читает shablom_ust.xlsx.
    Колонки: A=Tashkilot, B=Manzil, C=Rahbar, D=Muhandis1, E=Lavozim1,
             F=Muhandis2, G=Lavozim2, H=Sana, I=Uskuna nomi,
             J=Seriya, K=O'rnatish joyi
    Строка 1 = заголовки, данные с строки 2.
    Возвращает один dict со списком items.
    """
    rows = _read_rows(path)
    if not rows:
        raise ValueError(f'Файл пуст: {path}')

    org   = ''; addr  = ''; boss  = ''
    eng1  = ''; job1  = ''; eng2  = ''; job2  = ''
    items = []

    for row in rows[1:]:          # данные с индекса 1
        if not row or not any(row):
            continue

        org_v  = _g(row,  0)   # A
        addr_v = _g(row,  1)   # B
        boss_v = _g(row,  2)   # C
        e1_v   = _g(row,  3)   # D
        j1_v   = _g(row,  4)   # E
        e2_v   = _g(row,  5)   # F
        j2_v   = _g(row,  6)   # G
        sana   = _g(row,  7)   # H  — дата установки
        nom    = _g(row,  8)   # I  — наименование оборудования
        seriya = _g(row,  9)   # J
        joyi   = _g(row, 10)   # K

        if not org   and org_v:  org  = org_v
        if not addr  and addr_v: addr = addr_v
        if not boss  and boss_v: boss = boss_v
        if not eng1  and e1_v:   eng1 = e1_v;  job1 = j1_v
        if not eng2  and e2_v:   eng2 = e2_v;  job2 = j2_v

        if not nom:
            continue

        items.append({
            'num':           str(len(items) + 1),
            'name':          nom,
            'model':         nom,
            'serial_number': seriya,
            'inv_number':    '',
            'install_date':  _to_date(sana) or datetime.now().strftime('%d.%m.%Y'),
            'location':      joyi,
            'cost':          '',
            'condition':     'Yangi',
            'note':          joyi,
        })

    doc_date = (items[0]['install_date']
                if items else datetime.now().strftime('%d.%m.%Y'))

    return {
        'org_name':        org,
        'region':          org,
        'address':         addr,
        'doc_number':      '1',
        'doc_date':        doc_date,
        'commission_head': boss,
        'member1':         eng1,
        'member1_title':   job1 or 'Yetakchi muhandis',
        'member2':         eng2,
        'member2_title':   job2 or 'Yetakchi muhandis',
        'items':           items,
    }
