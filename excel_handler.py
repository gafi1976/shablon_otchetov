# excel_handler.py
# Чтение реальных Excel шаблонов и создание новых шаблонов
#
# ═══════════════════════════════════════════════════
#  СТРУКТУРА shablon_spisan.xlsx / shablon.xlsx
# ═══════════════════════════════════════════════════
#  Строка 1 — заголовки:
#    A: Qurilma nomi          — название оборудования
#    B: Qismlar nomi          — компонент
#    C: Foydalanishga yaroqliligi — состояние
#    D: Nosozlik belgilari    — неисправность
#    E: Inventar raqami       — инвентарный номер  ← ключ группировки
#    F: Tashilot nomi         — организация
#    G: {boss name}           — руководитель
#    H: {enginer1}            — инженер 1
#    I: {enginer2}            — инженер 2
#  Пустые строки между группами — разделители.
#
# ═══════════════════════════════════════════════════
#  СТРУКТУРА shablom_ust.xlsx
# ═══════════════════════════════════════════════════
#  Строка 1 — заголовки:
#    A: Data                  — дата
#    B: {num_otch}            — № строки
#    C: {oborud name}         — модель оборудования
#    D: serial num            — серийный номер
#    E: (пусто / inv)
#    F: {where oborud}        — место установки
#    G: {Organizasiya}        — организация
#    H: {adress}              — адрес
#    I: {boss name}           — руководитель
#    J: {engine1}             — инженер 1
#    K: {job title1}          — должность 1
#    L: {engine2}             — инженер 2
#    M: {job title2}          — должность 2
# ═══════════════════════════════════════════════════

import os
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime, date, timedelta
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation


# ──────────────────────────────────────────────
#  Низкоуровневый читатель xlsx через ZIP+XML
#  (работает без pandas, только stdlib)
# ──────────────────────────────────────────────

def _read_xlsx_rows(path: str) -> list:
    """
    Читает первый лист xlsx и возвращает список строк (list of list of str).
    Индекс 0 = строка 1 в Excel.
    Пустые строки возвращаются как [].
    """
    with zipfile.ZipFile(path) as z:
        names = z.namelist()

        # Shared strings
        strings = []
        if 'xl/sharedStrings.xml' in names:
            ns = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
            tree = ET.parse(z.open('xl/sharedStrings.xml'))
            for si in tree.findall('.//ns:si', ns):
                t = ''.join(n.text or '' for n in si.findall('.//ns:t', ns))
                strings.append(t)

        # Первый лист
        sheets = sorted(f for f in names if f.startswith('xl/worksheets/sheet'))
        if not sheets:
            return []

        ns = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        tree = ET.parse(z.open(sheets[0]))
        all_rows = tree.findall('.//ns:row', ns)

        def col_num(ref: str) -> int:
            col = ''.join(c for c in ref if c.isalpha())
            n = 0
            for ch in col.upper():
                n = n * 26 + (ord(ch) - ord('A') + 1)
            return n

        # Максимальный номер строки
        max_row = max((int(r.get('r', 0)) for r in all_rows), default=0)
        result = [[] for _ in range(max_row + 1)]

        for row_el in all_rows:
            row_num = int(row_el.get('r', 0))
            cells = row_el.findall('ns:c', ns)
            if not cells:
                result[row_num] = []
                continue

            max_col = max(col_num(c.get('r', 'A')) for c in cells)
            row_vals = [''] * max_col

            for c in cells:
                ref = c.get('r', 'A1')
                ci  = col_num(ref) - 1
                t   = c.get('t', '')
                v_el = c.find('ns:v', ns)
                if v_el is not None and v_el.text is not None:
                    if t == 's':
                        idx = int(v_el.text)
                        row_vals[ci] = strings[idx] if idx < len(strings) else ''
                    else:
                        row_vals[ci] = v_el.text
                else:
                    row_vals[ci] = ''

            result[row_num] = row_vals

        return result[1:]   # убираем нулевой индекс, возвращаем с 0 = строка 1


def _s(row: list, idx: int) -> str:
    """Безопасно получает строковое значение ячейки."""
    return str(row[idx]).strip() if idx < len(row) and row[idx] else ''


def _excel_date(val: str) -> str:
    """Конвертирует числовую дату Excel в строку дд.мм.гггг."""
    try:
        n = float(val)
        d = date(1899, 12, 30) + timedelta(days=int(n))
        return d.strftime('%d.%m.%Y')
    except Exception:
        return str(val).strip()


# ══════════════════════════════════════════════
#  ЧТЕНИЕ ШАБЛОНА СПИСАНИЯ
#  → список групп, сгруппированных по Inventar raqami
#  → каждая группа = один Word файл
# ══════════════════════════════════════════════

def read_spisan_excel(path: str) -> list:
    """
    Читает shablon_spisan.xlsx / shablon.xlsx.

    Возвращает список групп (list of dict), где каждая группа
    соответствует одному инвентарному номеру и передаётся в
    generator_spisan.create_spisan_doc().

    Структура одной группы:
    {
        'inv_number':      str,
        'org_name':        str,
        'region':          str,   # пустая — задаётся вручную
        'doc_date':        str,   # дд.мм.гггг (сегодня)
        'commission_head': str,
        'member1':         str,
        'member2':         str,
        'devices': [
            {
                'name':  str,
                'parts': [
                    {'part_name': str, 'condition': str, 'defect': str}
                ]
            }
        ]
    }
    """
    rows = _read_xlsx_rows(path)
    if not rows:
        raise ValueError(f'Файл пуст или не читается: {path}')

    # Строка 0 — заголовки, пропускаем
    # Группируем: inv_number → {org, boss, eng1, eng2, devices: {name → parts}}

    inv_groups   = {}   # {inv_number: dict}
    inv_order    = []   # сохраняем порядок инвентарных номеров

    for row in rows[1:]:            # с 1-го индекса = 2-я строка Excel
        if not row or not any(row):
            continue                # пустая строка — разделитель, пропускаем

        dev_name  = _s(row, 0)
        part_name = _s(row, 1)
        condition = _s(row, 2)
        defect    = _s(row, 3)
        inv       = _s(row, 4)
        org       = _s(row, 5)
        boss      = _s(row, 6)
        eng1      = _s(row, 7)
        eng2      = _s(row, 8)

        if not inv:
            continue   # строка без инвентарного номера — пропускаем

        # Создаём группу если ещё нет
        if inv not in inv_groups:
            inv_groups[inv] = {
                'inv_number':      inv,
                'org_name':        org,
                'region':          '',
                'doc_date':        datetime.now().strftime('%d.%m.%Y'),
                'commission_head': boss,
                'member1':         eng1,
                'member2':         eng2,
                'devices':         {},   # {dev_name: [parts]}
                'device_order':    [],
            }
            inv_order.append(inv)

        grp = inv_groups[inv]

        # Обновляем общие поля если они пустые
        if not grp['org_name'] and org:
            grp['org_name'] = org
        if not grp['commission_head'] and boss:
            grp['commission_head'] = boss
        if not grp['member1'] and eng1:
            grp['member1'] = eng1
        if not grp['member2'] and eng2:
            grp['member2'] = eng2

        # Добавляем устройство
        if dev_name:
            if dev_name not in grp['devices']:
                grp['devices'][dev_name] = []
                grp['device_order'].append(dev_name)
            if part_name:
                grp['devices'][dev_name].append({
                    'part_name': part_name,
                    'condition': condition,
                    'defect':    defect,
                })

    # Преобразуем в финальный формат (devices как список)
    result = []
    for inv in inv_order:
        grp = inv_groups[inv]
        devices = []
        for dev_name in grp['device_order']:
            devices.append({
                'name':  dev_name,
                'parts': grp['devices'][dev_name],
            })
        result.append({
            'inv_number':      grp['inv_number'],
            'org_name':        grp['org_name'],
            'region':          grp['region'],
            'doc_date':        grp['doc_date'],
            'commission_head': grp['commission_head'],
            'member1':         grp['member1'],
            'member2':         grp['member2'],
            'devices':         devices,
        })

    return result


# ══════════════════════════════════════════════
#  ЧТЕНИЕ ШАБЛОНА УСТАНОВКИ
#  → один dict с полным списком items
# ══════════════════════════════════════════════

def read_ust_excel(path: str) -> dict:
    """
    Читает shablom_ust.xlsx.

    Возвращает один dict, совместимый с
    generator_ust.create_ust_doc().

    {
        'org_name':        str,
        'region':          str,
        'address':         str,
        'doc_number':      str,
        'doc_date':        str,
        'commission_head': str,
        'member1':         str,
        'member1_title':   str,
        'member2':         str,
        'member2_title':   str,
        'items': [
            {
                'num':           str,
                'name':          str,
                'serial_number': str,
                'inv_number':    str,
                'install_date':  str,
                'location':      str,
                'model':         str,
                'cost':          str,
                'condition':     str,
                'note':          str,
            }
        ]
    }
    """
    rows = _read_xlsx_rows(path)
    if not rows:
        raise ValueError(f'Файл пуст или не читается: {path}')

    org_name  = ''
    address   = ''
    boss      = ''
    eng1      = ''
    job1      = ''
    eng2      = ''
    job2      = ''
    doc_date  = datetime.now().strftime('%d.%m.%Y')
    items     = []

    for row in rows[1:]:        # с 1-го индекса = 2-я строка Excel
        if not row or not any(row):
            continue

        raw_date  = _s(row, 0)
        num       = _s(row, 1)
        oborud    = _s(row, 2)
        serial    = _s(row, 3)
        inv       = _s(row, 4)
        where     = _s(row, 5)
        org       = _s(row, 6)
        addr      = _s(row, 7)
        bss       = _s(row, 8)
        e1        = _s(row, 9)
        j1        = _s(row, 10)
        e2        = _s(row, 11)
        j2        = _s(row, 12)

        # Общие поля берём из первой строки с данными
        if not org_name and org:
            org_name = org
        if not address and addr:
            address = addr
        if not boss and bss:
            boss = bss
        if not eng1 and e1:
            eng1 = e1
            job1 = j1
        if not eng2 and e2:
            eng2 = e2
            job2 = j2

        # Дата
        if raw_date and doc_date == datetime.now().strftime('%d.%m.%Y'):
            if '.' not in raw_date and raw_date.replace('.','').isdigit():
                doc_date = _excel_date(raw_date)
            else:
                doc_date = raw_date

        if not oborud:
            continue

        items.append({
            'num':           num or str(len(items) + 1),
            'name':          oborud,
            'model':         oborud,
            'serial_number': serial,
            'inv_number':    inv,
            'install_date':  doc_date,
            'location':      where,
            'cost':          '',
            'condition':     'Yangi',
            'note':          where,
        })

    return {
        'org_name':        org_name,
        'region':          org_name,
        'address':         address,
        'doc_number':      '1',
        'doc_date':        doc_date,
        'commission_head': boss,
        'member1':         eng1,
        'member1_title':   job1 or 'Yetakchi muhandis',
        'member2':         eng2,
        'member2_title':   job2 or 'Yetakchi muhandis',
        'items':           items,
    }


# ══════════════════════════════════════════════
#  СТИЛИ ДЛЯ СОЗДАНИЯ ШАБЛОНОВ
# ══════════════════════════════════════════════

_HDR_FILL_S = PatternFill('solid', fgColor='2E4057')
_HDR_FILL_U = PatternFill('solid', fgColor='1A5276')
_HDR_FONT   = Font(name='Times New Roman', bold=True, color='FFFFFF', size=10)
_DATA_FONT  = Font(name='Times New Roman', size=10)
_CENTER     = Alignment(horizontal='center', vertical='center', wrap_text=True)
_LEFT       = Alignment(horizontal='left',   vertical='center', wrap_text=True)
_EVEN_S     = PatternFill('solid', fgColor='F0F4F8')
_EVEN_U     = PatternFill('solid', fgColor='EBF5FB')
_ODD        = PatternFill('solid', fgColor='FFFFFF')


def _border(color='2E4057'):
    s = Side(style='thin', color=color)
    return Border(left=s, right=s, top=s, bottom=s)


def _hdr_cell(ws, row, col, text, fill, width=None):
    cell = ws.cell(row=row, column=col, value=text)
    cell.font      = _HDR_FONT
    cell.fill      = fill
    cell.alignment = _CENTER
    cell.border    = _border()
    if width:
        from openpyxl.utils import get_column_letter
        ws.column_dimensions[get_column_letter(col)].width = width
    return cell


# ══════════════════════════════════════════════
#  СОЗДАНИЕ ШАБЛОНА EXCEL — СПИСАНИЕ
#  Структура = shablon_spisan.xlsx
# ══════════════════════════════════════════════

def create_spisan_template(path: str) -> str:
    """
    Создаёт Excel шаблон для данных СПИСАНИЯ.
    Структура совпадает с shablon_spisan.xlsx:
      A=Qurilma nomi, B=Qismlar nomi, C=Foydalanishga yaroqliligi,
      D=Nosozlik belgilari, E=Inventar raqami,
      F=Tashilot nomi, G={boss name}, H={enginer1}, I={enginer2}
    """
    wb = Workbook()
    ws = wb.active
    ws.title = 'Spisaniye'
    ws.sheet_view.showGridLines = False
    ws.row_dimensions[1].height = 38

    headers = [
        ('Qurilma nomi\n(Название оборудования)', 28),
        ('Qismlar nomi\n(Компонент)',              22),
        ('Foydalanishga yaroqliligi\n(Состояние)', 20),
        ('Nosozlik belgilari\n(Неисправность)',    26),
        ('Inventar raqami\n(Инв. номер) *',        18),
        ('Tashilot nomi\n(Организация)',            32),
        ('{boss name}\n(Руководитель)',             22),
        ('{enginer1}\n(Инженер 1)',                 22),
        ('{enginer2}\n(Инженер 2)',                 22),
    ]
    for ci, (text, width) in enumerate(headers, start=1):
        _hdr_cell(ws, 1, ci, text, _HDR_FILL_S, width)

    # Валидация состояния
    dv = DataValidation(type='list', formula1='"Yaroqli,Yaroqsiz"',
                        allow_blank=True, showDropDown=False)
    dv.sqref = 'C2:C500'
    ws.add_data_validation(dv)

    # Примеры данных
    examples = [
        ('Server HP ML 350',   'Asosiy plata', 'Yaroqli',  "ma'naviy eskirgan", '26-0006039', 'ABM Sirdaryo viloyati MChJ', 'Palonov P.A', 'Raxmatov V.A', 'Xolbekov G.T'),
        ('Server HP ML 350',   'Tok manbayi',  'Yaroqli',  "ma'naviy eskirgan", '26-0006039', 'ABM Sirdaryo viloyati MChJ', 'Palonov P.A', 'Raxmatov V.A', 'Xolbekov G.T'),
        ('Server HP ML 350',   'Qattiq disk',  'Yaroqsiz', 'BAD bloklar mavjud','26-0006039', 'ABM Sirdaryo viloyati MChJ', 'Palonov P.A', 'Raxmatov V.A', 'Xolbekov G.T'),
        ('',)*9,
        ('Кондиционер ART-12HI','Tok manbai',  'Yaroqsiz', "tok sarfi ko'p",    '24-0006162', 'ABM Sirdaryo viloyati MChJ', 'Palonov P.A', 'Raxmatov V.A', 'Xolbekov G.T'),
        ('Кондиционер ART-12HI','Kompressor',  'Yaroqsiz', 'shovqin mavjud',    '24-0006162', 'ABM Sirdaryo viloyati MChJ', 'Palonov P.A', 'Raxmatov V.A', 'Xolbekov G.T'),
        ('Кондиционер ART-12HI','Radiator',    'Yaroqsiz', 'yamalgan',          '24-0006162', 'ABM Sirdaryo viloyati MChJ', 'Palonov P.A', 'Raxmatov V.A', 'Xolbekov G.T'),
        ('',)*9,
        ("Монитор LCD 19''",   'Asosiy plata', 'Yaroqsiz', 'korpus singan',     '26-0003436', 'ABM Sirdaryo viloyati MChJ', 'Palonov P.A', 'Raxmatov V.A', 'Xolbekov G.T'),
        ("Монитор LCD 19''",   'Tok manbayi',  'Yaroqsiz', 'kuygan',            '26-0003436', 'ABM Sirdaryo viloyati MChJ', 'Palonov P.A', 'Raxmatov V.A', 'Xolbekov G.T'),
    ]

    for ri, row_data in enumerate(examples):
        rn   = ri + 2
        ws.row_dimensions[rn].height = 18
        empty = not any(row_data)
        fill  = _EVEN_S if (ri % 2 == 0 and not empty) else _ODD
        for ci, val in enumerate(row_data, start=1):
            cell = ws.cell(row=rn, column=ci, value=val or None)
            cell.font      = _DATA_FONT
            cell.border    = _border()
            cell.alignment = _CENTER if ci == 3 else _LEFT
            if not empty:
                cell.fill = fill
                if ci == 3 and val == 'Yaroqsiz':
                    cell.font = Font(name='Times New Roman', size=10,
                                     color='C0392B', bold=True)

    # Пустые строки для ввода
    for ri in range(len(examples), len(examples) + 15):
        rn = ri + 2
        ws.row_dimensions[rn].height = 18
        fill = _EVEN_S if ri % 2 == 0 else _ODD
        for ci in range(1, 10):
            cell = ws.cell(row=rn, column=ci)
            cell.font   = _DATA_FONT
            cell.fill   = fill
            cell.border = _border()
            cell.alignment = _LEFT

    # Инструкция
    ir = len(examples) + 17
    ws.merge_cells(start_row=ir, start_column=1, end_row=ir, end_column=9)
    ws.cell(row=ir, column=1).value = (
        '* ИНСТРУКЦИЯ: Каждая строка — один компонент оборудования. '
        'Группировка в один Word файл идёт по полю "Inventar raqami". '
        'Пустая строка = визуальный разделитель (необязательно). '
        'Yaroqli = исправен, Yaroqsiz = неисправен.'
    )
    ws.cell(row=ir, column=1).font = Font(name='Times New Roman',
                                          italic=True, size=9, color='888888')
    ws.cell(row=ir, column=1).alignment = _LEFT
    ws.row_dimensions[ir].height = 22

    ws.freeze_panes = 'A2'
    wb.save(path)
    return path


# ══════════════════════════════════════════════
#  СОЗДАНИЕ ШАБЛОНА EXCEL — УСТАНОВКА
#  Структура = shablom_ust.xlsx
# ══════════════════════════════════════════════

def create_ust_template(path: str) -> str:
    """
    Создаёт Excel шаблон для данных УСТАНОВКИ.
    Структура совпадает с shablom_ust.xlsx:
      A=Data, B={num_otch}, C={oborud name}, D=serial num, E=inv,
      F={where oborud}, G={Organizasiya}, H={adress},
      I={boss name}, J={engine1}, K={job title1}, L={engine2}, M={job title2}
    """
    wb = Workbook()
    ws = wb.active
    ws.title = 'Ustanovka'
    ws.sheet_view.showGridLines = False
    ws.row_dimensions[1].height = 38

    headers = [
        ('Data\n(Дата)',                     14),
        ('{num_otch}\n(№)',                   6),
        ('{oborud name}\n(Модель/Марка)',     26),
        ('Serial num\n(Серийный №)',          18),
        ('Inventar raqami\n(Инв. номер)',     16),
        ('{where oborud}\n(Место установки)', 24),
        ('{Organizasiya}\n(Организация)',     28),
        ('{adress}\n(Адрес)',                 22),
        ('{boss name}\n(Руководитель)',       22),
        ('{engine1}\n(Инженер 1)',            20),
        ('{job title1}\n(Должность 1)',       18),
        ('{engine2}\n(Инженер 2)',            20),
        ('{job title2}\n(Должность 2)',       18),
    ]
    for ci, (text, width) in enumerate(headers, start=1):
        _hdr_cell(ws, 1, ci, text, _HDR_FILL_U, width)

    today = datetime.now().strftime('%d.%m.%Y')
    examples = [
        (today, '1', 'Maipu MP1900X-22',   'SN-001', 'INV-0001', 'Guliston, 1-bino', 'ABM Sirdaryo viloyati MChJ', "Guliston sh., O'zbekiston k.", 'Palonov P.A', 'Raxmatov V.A', 'Yetakchi muhandis', 'Xolbekov G.T', 'Yetakchi muhandis'),
        (today, '2', 'Maipu MP1900X-22',   'SN-002', 'INV-0002', 'Guliston, 2-bino', 'ABM Sirdaryo viloyati MChJ', "Guliston sh., O'zbekiston k.", 'Palonov P.A', 'Raxmatov V.A', 'Yetakchi muhandis', 'Xolbekov G.T', 'Yetakchi muhandis'),
        (today, '3', 'Switch Cisco SG350', 'SN-003', 'INV-0003', 'Guliston, 3-bino', 'ABM Sirdaryo viloyati MChJ', "Guliston sh., O'zbekiston k.", 'Palonov P.A', 'Raxmatov V.A', 'Yetakchi muhandis', 'Xolbekov G.T', 'Yetakchi muhandis'),
    ]

    for ri, row_data in enumerate(examples):
        rn   = ri + 2
        ws.row_dimensions[rn].height = 20
        fill = _EVEN_U if ri % 2 == 0 else _ODD
        for ci, val in enumerate(row_data, start=1):
            cell = ws.cell(row=rn, column=ci, value=val)
            cell.font      = _DATA_FONT
            cell.fill      = fill
            cell.border    = _border('1A5276')
            cell.alignment = _CENTER if ci in (1, 2, 4, 5) else _LEFT

    for ri in range(len(examples), len(examples) + 15):
        rn = ri + 2
        ws.row_dimensions[rn].height = 20
        fill = _EVEN_U if ri % 2 == 0 else _ODD
        for ci in range(1, 14):
            cell = ws.cell(row=rn, column=ci)
            cell.font   = _DATA_FONT
            cell.fill   = fill
            cell.border = _border('1A5276')
            cell.alignment = _LEFT

    ir = len(examples) + 17
    ws.merge_cells(start_row=ir, start_column=1, end_row=ir, end_column=13)
    ws.cell(row=ir, column=1).value = (
        'ИНСТРУКЦИЯ: Каждая строка — одна единица оборудования. '
        'Поля организации, руководителя и инженеров одинаковы для всех строк. '
        'Дата вводится в формате дд.мм.гггг.'
    )
    ws.cell(row=ir, column=1).font = Font(name='Times New Roman',
                                          italic=True, size=9, color='888888')
    ws.cell(row=ir, column=1).alignment = _LEFT
    ws.row_dimensions[ir].height = 22

    ws.freeze_panes = 'A2'
    wb.save(path)
    return path
