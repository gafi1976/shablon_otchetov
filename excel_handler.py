# excel_handler.py — Чтение реальных шаблонов из GitHub и создание новых шаблонов
#
# ═══════════════════════════════════════════════════════════════
#  СТРУКТУРА ОРИГИНАЛЬНЫХ ШАБЛОНОВ (shablon_spisan.xlsx / shablom_ust.xlsx)
# ═══════════════════════════════════════════════════════════════
#
#  shablon_spisan.xlsx / shablon.xlsx — шаблон СПИСАНИЯ:
#  Строка 1 — заголовки:
#    A: Qurilma nomi         — название оборудования (группировка)
#    B: Qismlar nomi         — название компонента
#    C: Foydalanishga yaroqliligi — состояние (Yaroqli / Yaroqsiz)
#    D: Nosozlik belgilari   — признаки неисправности
#    E: Inventar raqami / {inv_num} — инвентарный номер
#    F: Tashilot nomi        — название организации
#    G: {boss name}          — руководитель (подписант)
#    H: {enginer1}           — инженер 1
#    I: {enginer2}           — инженер 2
#  Строки 2+ — данные. Пустые строки между группами оборудования — разделители.
#
#  shablom_ust.xlsx — шаблон УСТАНОВКИ:
#  Строка 1 — заголовки:
#    A: Data                 — дата (число Excel или строка)
#    B: {num_otch}           — порядковый номер строки
#    C: {oborud name}        — наименование оборудования (модель)
#    D: (серийный номер — пусто в примере)
#    E: (пусто)
#    F: {where oborud}       — место установки
#    G: {Organizasiya}       — адрес организации
#    H: {adress}             — адрес
#    I: {boss name}          — руководитель
#    J: {engine1}            — инженер 1
#    K: {job title1}         — должность инженера 1
#    L: {engine2}            — инженер 2
#    M: {job title2}         — должность инженера 2
#  Строки 2+ — данные оборудования.
# ═══════════════════════════════════════════════════════════════

import os
import zipfile
import xml.etree.ElementTree as ET
from openpyxl import Workbook
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side
)
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from datetime import datetime, date


# ══════════════════════════════════════════════════════
#  НИЗКОУРОВНЕВЫЙ ЧИТАТЕЛЬ xlsx (через ZIP + XML)
#  Работает без pandas, совместим с любой версией Python
# ══════════════════════════════════════════════════════

def _read_xlsx_rows(path: str) -> list:
    """
    Читает первый лист xlsx-файла и возвращает список строк.
    Каждая строка — список значений ячеек (str).
    Пустые строки возвращаются как пустые списки [].
    """
    with zipfile.ZipFile(path) as z:
        names = z.namelist()

        # Общие строки (sharedStrings)
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
        all_rows_el = tree.findall('.//ns:row', ns)

        # Определяем максимальный номер строки
        max_row = 0
        for row_el in all_rows_el:
            r = int(row_el.get('r', 0))
            if r > max_row:
                max_row = r

        result = [[] for _ in range(max_row + 1)]  # индекс = номер строки

        for row_el in all_rows_el:
            row_num = int(row_el.get('r', 0))
            row_data = []
            # Собираем ячейки с учётом пропусков (пустые ячейки)
            cells = row_el.findall('ns:c', ns)
            if not cells:
                result[row_num] = []
                continue

            # Определяем макс. колонку в строке
            def col_letter_to_num(ref):
                col = ''.join(c for c in ref if c.isalpha())
                num = 0
                for ch in col:
                    num = num * 26 + (ord(ch.upper()) - ord('A') + 1)
                return num

            max_col = max(col_letter_to_num(c.get('r', 'A')) for c in cells)

            row_vals = [''] * max_col
            for c in cells:
                ref = c.get('r', 'A1')
                col_num = col_letter_to_num(ref) - 1
                t = c.get('t', '')
                v_el = c.find('ns:v', ns)
                if v_el is not None and v_el.text is not None:
                    if t == 's':
                        idx = int(v_el.text)
                        row_vals[col_num] = strings[idx] if idx < len(strings) else ''
                    elif t == 'str' or t == 'inlineStr':
                        row_vals[col_num] = v_el.text
                    else:
                        row_vals[col_num] = v_el.text
                else:
                    row_vals[col_num] = ''

            result[row_num] = row_vals

        # Убираем нулевой элемент (строки нумеруются с 1)
        return result[1:]


def _excel_date_to_str(val: str) -> str:
    """Конвертирует числовую дату Excel в строку дд.мм.гггг."""
    try:
        n = float(val)
        # Excel serial date: 1 = 01.01.1900
        d = date(1899, 12, 30)
        from datetime import timedelta
        d = d + timedelta(days=int(n))
        return d.strftime('%d.%m.%Y')
    except Exception:
        return str(val)


def _safe(val) -> str:
    if val is None:
        return ''
    return str(val).strip()


# ══════════════════════════════════════════════════════
#  ЧТЕНИЕ ШАБЛОНА СПИСАНИЯ
#  Формат: shablon_spisan.xlsx / shablon.xlsx
# ══════════════════════════════════════════════════════

def read_spisan_excel(path: str) -> dict:
    """
    Читает Excel шаблон списания (shablon_spisan.xlsx или shablon.xlsx).

    Структура файла:
      Строка 1: заголовки
        A=Qurilma nomi, B=Qismlar nomi, C=Foydalanishga yaroqliligi,
        D=Nosozlik belgilari, E=Inventar raqami, F=Tashilot nomi,
        G=boss name (руководитель), H=enginer1, I=enginer2
      Строки 2+: данные (пустые строки — разделители между группами)

    Возвращает dict совместимый с generator_spisan.create_spisan_doc():
      {
        'org_name': str,
        'region': str,           # пусто — не хранится в шаблоне
        'doc_number': str,       # пусто — задаётся вручную
        'doc_date': str,
        'commission_head': str,
        'member1': str,
        'member2': str,
        'items': [
          {
            'num': str,
            'name': str,          # Qurilma nomi
            'inv_number': str,    # Inventar raqami
            'year': str,
            'initial_cost': str,
            'residual_cost': str,
            'reason': str,        # Nosozlik belgilari
            'note': str,          # Foydalanishga yaroqliligi
            'parts': [            # Qismlar — детали оборудования
              {'part_name': str, 'condition': str, 'defect': str}
            ]
          }
        ]
      }
    """
    rows = _read_xlsx_rows(path)
    if not rows:
        raise ValueError(f'Файл пуст или не читается: {path}')

    # Строка 1 — заголовки, пропускаем
    # Извлекаем общие данные из первой строки с данными (строка 2)
    org_name        = ''
    commission_head = ''
    member1         = ''
    member2         = ''

    # Группируем строки по оборудованию (Qurilma nomi)
    # Каждая группа — отдельное оборудование
    devices = {}   # {device_name: {'inv': str, 'org': str, 'boss': str, 'eng1': str, 'eng2': str, 'parts': []}}
    device_order = []  # сохраняем порядок

    for i, row in enumerate(rows[1:], start=2):  # с 2-й строки
        if not row or not any(row):
            continue  # пустая строка — разделитель

        def g(idx):
            return _safe(row[idx]) if idx < len(row) else ''

        device_name = g(0)
        part_name   = g(1)
        condition   = g(2)
        defect      = g(3)
        inv_number  = g(4)
        org         = g(5)
        boss        = g(6)
        eng1        = g(7)
        eng2        = g(8)

        if not device_name and not part_name:
            continue

        # Извлекаем общие данные из первой встреченной строки
        if not org_name and org:
            org_name = org
        if not commission_head and boss:
            commission_head = boss
        if not member1 and eng1:
            member1 = eng1
        if not member2 and eng2:
            member2 = eng2

        # Группировка по названию оборудования
        if device_name not in devices:
            devices[device_name] = {
                'inv': inv_number,
                'org': org,
                'boss': boss,
                'eng1': eng1,
                'eng2': eng2,
                'parts': []
            }
            device_order.append(device_name)

        if part_name:
            devices[device_name]['parts'].append({
                'part_name': part_name,
                'condition': condition,
                'defect':    defect,
            })

    # Формируем список items (одно устройство = одна строка в акте)
    items = []
    for i, dev_name in enumerate(device_order, start=1):
        dev = devices[dev_name]
        # Определяем причину списания из компонентов
        defects = [p['defect'] for p in dev['parts'] if p['defect']]
        bad_parts = [p['part_name'] for p in dev['parts'] if p['condition'].lower() in ('yaroqsiz', 'яроқсиз')]
        reason = '; '.join(sorted(set(defects))) if defects else 'Eskirgan'
        note   = f"Yaroqsiz qismlar: {', '.join(bad_parts)}" if bad_parts else ''

        items.append({
            'num':           str(i),
            'name':          dev_name.strip(),
            'inv_number':    dev['inv'],
            'year':          '',
            'initial_cost':  '',
            'residual_cost': '',
            'reason':        reason,
            'note':          note,
            # Передаём детали для расширенного отчёта
            'parts':         dev['parts'],
        })

    return {
        'org_name':        org_name,
        'region':          '',
        'doc_number':      '',
        'doc_date':        datetime.now().strftime('%d.%m.%Y'),
        'commission_head': commission_head,
        'member1':         member1,
        'member2':         member2,
        'items':           items,
    }


# ══════════════════════════════════════════════════════
#  ЧТЕНИЕ ШАБЛОНА УСТАНОВКИ
#  Формат: shablom_ust.xlsx
# ══════════════════════════════════════════════════════

def read_ust_excel(path: str) -> dict:
    """
    Читает Excel шаблон установки (shablom_ust.xlsx).

    Структура файла:
      Строка 1: заголовки
        A=Data (дата), B=num_otch (№), C=oborud name (модель оборудования),
        D=серийный номер (пусто), E=пусто,
        F=where oborud (место установки), G=Organizasiya, H=adress,
        I=boss name, J=engine1, K=job title1, L=engine2, M=job title2
      Строки 2+: данные

    Возвращает dict совместимый с generator_ust.create_ust_doc().
    """
    rows = _read_xlsx_rows(path)
    if not rows:
        raise ValueError(f'Файл пуст или не читается: {path}')

    org_name        = ''
    location        = ''
    address         = ''
    commission_head = ''
    member1         = ''
    member1_title   = ''
    member2         = ''
    member2_title   = ''
    doc_date        = datetime.now().strftime('%d.%m.%Y')

    items = []

    for i, row in enumerate(rows[1:], start=2):
        if not row or not any(row):
            continue

        def g(idx):
            return _safe(row[idx]) if idx < len(row) else ''

        raw_date    = g(0)
        num         = g(1)
        oborud_name = g(2)
        serial_num  = g(3)
        # g(4) — пусто
        where       = g(5)
        org         = g(6)
        addr        = g(7)
        boss        = g(8)
        eng1        = g(9)
        job1        = g(10)
        eng2        = g(11)
        job2        = g(12)

        # Извлекаем общие данные из первой строки
        if not org_name and org:
            org_name = org
        if not location and where:
            location = where
        if not address and addr:
            address = addr
        if not commission_head and boss:
            commission_head = boss
        if not member1 and eng1:
            member1 = eng1
            member1_title = job1
        if not member2 and eng2:
            member2 = eng2
            member2_title = job2

        # Дата из первой строки
        if raw_date and doc_date == datetime.now().strftime('%d.%m.%Y'):
            # Если это число — конвертируем из Excel serial date
            if raw_date.replace('.', '').isdigit() and '.' not in raw_date:
                doc_date = _excel_date_to_str(raw_date)
            elif raw_date:
                doc_date = raw_date

        if not oborud_name:
            continue

        items.append({
            'num':           num or str(len(items) + 1),
            'name':          oborud_name,
            'model':         oborud_name,   # модель = название
            'serial_number': serial_num,
            'inv_number':    '',
            'year':          '',
            'cost':          '',
            'install_date':  doc_date,
            'condition':     'Янги (новый)',
            'note':          where,
        })

    return {
        'org_name':        org_name,
        'region':          address,
        'doc_number':      '',
        'doc_date':        doc_date,
        'location':        location,
        'commission_head': commission_head,
        'member1':         member1,
        'member2':         member2,
        'items':           items,
    }


# ══════════════════════════════════════════════════════
#  СОЗДАНИЕ ШАБЛОНА EXCEL — СПИСАНИЕ
#  Структура совпадает с shablon_spisan.xlsx / shablon.xlsx
# ══════════════════════════════════════════════════════

HEADER_FILL_SPISAN = PatternFill('solid', fgColor='2E4057')
HEADER_FILL_UST    = PatternFill('solid', fgColor='1A5276')
HEADER_FONT        = Font(name='Times New Roman', bold=True, color='FFFFFF', size=10)
DATA_FONT          = Font(name='Times New Roman', size=10)
CENTER             = Alignment(horizontal='center', vertical='center', wrap_text=True)
LEFT               = Alignment(horizontal='left',   vertical='center', wrap_text=True)
ROW_EVEN_SPISAN    = PatternFill('solid', fgColor='F0F4F8')
ROW_EVEN_UST       = PatternFill('solid', fgColor='EBF5FB')
ROW_ODD            = PatternFill('solid', fgColor='FFFFFF')


def _thin_border(color='2E4057'):
    s = Side(style='thin', color=color)
    return Border(left=s, right=s, top=s, bottom=s)


def create_spisan_template(path: str) -> str:
    """
    Создаёт Excel шаблон для данных СПИСАНИЯ.

    Структура (совпадает с shablon_spisan.xlsx):
      Строка 1: заголовки
        A=Qurilma nomi, B=Qismlar nomi, C=Foydalanishga yaroqliligi,
        D=Nosozlik belgilari, E=Inventar raqami,
        F=Tashilot nomi, G={boss name}, H={enginer1}, I={enginer2}
      Строки 2+: данные оборудования
    """
    wb = Workbook()
    ws = wb.active
    ws.title = 'Spisaniye'
    ws.sheet_view.showGridLines = False

    # ── Строка заголовка ──
    headers = [
        ('A', 'Qurilma nomi\n(Название оборудования)', 28),
        ('B', 'Qismlar nomi\n(Компонент)',              22),
        ('C', 'Foydalanishga yaroqliligi\n(Состояние)', 20),
        ('D', 'Nosozlik belgilari\n(Неисправность)',    26),
        ('E', 'Inventar raqami\n(Инв. номер)',          18),
        ('F', 'Tashilot nomi\n(Организация)',            32),
        ('G', '{boss name}\n(Руководитель)',             22),
        ('H', '{enginer1}\n(Инженер 1)',                 22),
        ('I', '{enginer2}\n(Инженер 2)',                 22),
    ]

    ws.row_dimensions[1].height = 38
    for col_letter, header, width in headers:
        col_idx = ord(col_letter) - ord('A') + 1
        cell = ws.cell(row=1, column=col_idx)
        cell.value = header
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL_SPISAN
        cell.alignment = CENTER
        cell.border = _thin_border()
        ws.column_dimensions[col_letter].width = width

    # ── Пример данных (3 устройства, по 2-3 компонента) ──
    example_rows = [
        # Qurilma nomi,              Qismlar nomi,  Yaroqliligi,  Nosozlik,                     Inv,           Tashilot,                                     Boss,          Eng1,          Eng2
        ('Server HP ML 350',         'Asosiy plata','Yaroqli',    "ma'naviy eskirgan",           '26-0006039',  'ABM Sirdaryo viloyati axborot markazi MChJ', 'Palonov P.A', 'Raxmatov V.A','Xolbekov G.T'),
        ('Server HP ML 350',         'Tok manbayi', 'Yaroqli',    "ma'naviy eskirgan",           '26-0006039',  'ABM Sirdaryo viloyati axborot markazi MChJ', 'Palonov P.A', 'Raxmatov V.A','Xolbekov G.T'),
        ('Server HP ML 350',         'Qattiq disk', 'Yaroqsiz',   'BAD bloklar mavjud',          '26-0006039',  'ABM Sirdaryo viloyati axborot markazi MChJ', 'Palonov P.A', 'Raxmatov V.A','Xolbekov G.T'),
        ('',                         '',            '',           '',                            '',            '',                                           '',            '',            ''),
        ('Кондиционер ART-12HI',     'Tok manbai',  'Yaroqsiz',   "tok sarfi ko'p",              '24-0006162',  'ABM Sirdaryo viloyati axborot markazi MChJ', 'Palonov P.A', 'Raxmatov V.A','Xolbekov G.T'),
        ('Кондиционер ART-12HI',     'Kompressor',  'Yaroqsiz',   'shovqin mavjud',              '24-0006162',  'ABM Sirdaryo viloyati axborot markazi MChJ', 'Palonov P.A', 'Raxmatov V.A','Xolbekov G.T'),
        ('Кондиционер ART-12HI',     'Radiator',    'Yaroqsiz',   'yamalgan',                    '24-0006162',  'ABM Sirdaryo viloyati axborot markazi MChJ', 'Palonov P.A', 'Raxmatov V.A','Xolbekov G.T'),
        ('',                         '',            '',           '',                            '',            '',                                           '',            '',            ''),
        ("Монитор LCD 19''",         'Asosiy plata','Yaroqsiz',   'Yaroqsiz',                    '26-0003436',  'ABM Sirdaryo viloyati axborot markazi MChJ', 'Palonov P.A', 'Raxmatov V.A','Xolbekov G.T'),
        ("Монитор LCD 19''",         'Tok manbayi', 'Yaroqsiz',   'Yaroqsiz',                    '26-0003436',  'ABM Sirdaryo viloyati axborot markazi MChJ', 'Palonov P.A', 'Raxmatov V.A','Xolbekov G.T'),
        ("Монитор LCD 19''",         'Displey',     'Yaroqsiz',   'Yaroqsiz',                    '26-0003436',  'ABM Sirdaryo viloyati axborot markazi MChJ', 'Palonov P.A', 'Raxmatov V.A','Xolbekov G.T'),
    ]

    # Валидация состояния
    dv = DataValidation(
        type='list',
        formula1='"Yaroqli,Yaroqsiz"',
        allow_blank=True,
        showDropDown=False,
    )
    dv.sqref = f'C2:C200'
    ws.add_data_validation(dv)

    for i, row_data in enumerate(example_rows):
        row_num = i + 2
        ws.row_dimensions[row_num].height = 18
        is_empty = not any(row_data)
        fill = ROW_EVEN_SPISAN if (i % 2 == 0 and not is_empty) else ROW_ODD

        for col_idx, val in enumerate(row_data, start=1):
            cell = ws.cell(row=row_num, column=col_idx)
            cell.value = val if val else None
            cell.font = DATA_FONT
            cell.border = _thin_border()
            cell.alignment = CENTER if col_idx in (3,) else LEFT
            if not is_empty:
                cell.fill = fill
                # Красный цвет для Yaroqsiz
                if col_idx == 3 and val == 'Yaroqsiz':
                    cell.font = Font(name='Times New Roman', size=10,
                                     color='C0392B', bold=True)

    # ── 10 пустых строк для ввода данных ──
    for i in range(len(example_rows), len(example_rows) + 10):
        row_num = i + 2
        ws.row_dimensions[row_num].height = 18
        fill = ROW_EVEN_SPISAN if i % 2 == 0 else ROW_ODD
        for col_idx in range(1, 10):
            cell = ws.cell(row=row_num, column=col_idx)
            cell.font = DATA_FONT
            cell.fill = fill
            cell.border = _thin_border()
            cell.alignment = LEFT

    # ── Инструкция ──
    instr_row = len(example_rows) + 12 + 2
    ws.merge_cells(start_row=instr_row, start_column=1, end_row=instr_row, end_column=9)
    ws.cell(row=instr_row, column=1).value = (
        'ИНСТРУКЦИЯ: Каждая строка — один компонент оборудования. '
        'Оборудование группируется по полю "Qurilma nomi". '
        'Пустая строка = разделитель между разными устройствами. '
        'Yaroqli = исправен, Yaroqsiz = неисправен.'
    )
    ws.cell(row=instr_row, column=1).font = Font(name='Times New Roman', italic=True,
                                                  size=9, color='888888')
    ws.cell(row=instr_row, column=1).alignment = LEFT
    ws.row_dimensions[instr_row].height = 22

    ws.freeze_panes = 'A2'
    wb.save(path)
    return path


# ══════════════════════════════════════════════════════
#  СОЗДАНИЕ ШАБЛОНА EXCEL — УСТАНОВКА
#  Структура совпадает с shablom_ust.xlsx
# ══════════════════════════════════════════════════════

def create_ust_template(path: str) -> str:
    """
    Создаёт Excel шаблон для данных УСТАНОВКИ.

    Структура (совпадает с shablom_ust.xlsx):
      Строка 1: заголовки
        A=Data, B={num_otch}, C={oborud name}, D=serial num,
        E=inv_number, F={where oborud}, G={Organizasiya}, H={adress},
        I={boss name}, J={engine1}, K={job title1}, L={engine2}, M={job title2}
      Строки 2+: данные
    """
    wb = Workbook()
    ws = wb.active
    ws.title = 'Ustanovka'
    ws.sheet_view.showGridLines = False

    headers = [
        ('A', 'Data\n(Дата)',                    14),
        ('B', '{num_otch}\n(№)',                   6),
        ('C', '{oborud name}\n(Модель/Марка)',     28),
        ('D', 'Serial num\n(Серийный №)',          18),
        ('E', 'Inventar raqami\n(Инв. номер)',     16),
        ('F', '{where oborud}\n(Место установки)', 26),
        ('G', '{Organizasiya}\n(Организация)',     30),
        ('H', '{adress}\n(Адрес)',                 22),
        ('I', '{boss name}\n(Руководитель)',        22),
        ('J', '{engine1}\n(Инженер 1)',             20),
        ('K', '{job title1}\n(Должность 1)',        18),
        ('L', '{engine2}\n(Инженер 2)',             20),
        ('M', '{job title2}\n(Должность 2)',        18),
    ]

    ws.row_dimensions[1].height = 38
    for col_letter, header, width in headers:
        col_idx = ord(col_letter) - ord('A') + 1
        cell = ws.cell(row=1, column=col_idx)
        cell.value = header
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL_UST
        cell.alignment = CENTER
        cell.border = _thin_border('1A5276')
        ws.column_dimensions[col_letter].width = width

    today_str = datetime.now().strftime('%d.%m.%Y')

    example_rows = [
        (today_str, '1', 'Maipu MP1900X-22', 'SN-001-2024', 'INV-001', "Guliston sh., 1-bino", "Guliston sh. O'zbekiston k.", 'ABM Sirdaryo viloyati MChJ', 'Palonov P.A', 'Raxmatov V.A', 'Etakchi muhandis', 'Xolbekov G.T', 'Etakchi muhandis'),
        (today_str, '2', 'Maipu MP1900X-22', 'SN-002-2024', 'INV-002', "Guliston sh., 2-bino", "Guliston sh. O'zbekiston k.", 'ABM Sirdaryo viloyati MChJ', 'Palonov P.A', 'Raxmatov V.A', 'Etakchi muhandis', 'Xolbekov G.T', 'Etakchi muhandis'),
        (today_str, '3', 'Switch Cisco SG350', 'SN-003-2024','INV-003', "Guliston sh., 3-bino", "Guliston sh. O'zbekiston k.", 'ABM Sirdaryo viloyati MChJ', 'Palonov P.A', 'Raxmatov V.A', 'Etakchi muhandis', 'Xolbekov G.T', 'Etakchi muhandis'),
    ]

    for i, row_data in enumerate(example_rows):
        row_num = i + 2
        ws.row_dimensions[row_num].height = 20
        fill = ROW_EVEN_UST if i % 2 == 0 else ROW_ODD
        for col_idx, val in enumerate(row_data, start=1):
            cell = ws.cell(row=row_num, column=col_idx)
            cell.value = val
            cell.font = DATA_FONT
            cell.fill = fill
            cell.border = _thin_border('1A5276')
            cell.alignment = CENTER if col_idx in (1, 2, 4, 5) else LEFT

    # 10 пустых строк
    for i in range(len(example_rows), len(example_rows) + 10):
        row_num = i + 2
        ws.row_dimensions[row_num].height = 20
        fill = ROW_EVEN_UST if i % 2 == 0 else ROW_ODD
        for col_idx in range(1, 14):
            cell = ws.cell(row=row_num, column=col_idx)
            cell.font = DATA_FONT
            cell.fill = fill
            cell.border = _thin_border('1A5276')
            cell.alignment = LEFT

    # Инструкция
    instr_row = len(example_rows) + 12 + 2
    ws.merge_cells(start_row=instr_row, start_column=1, end_row=instr_row, end_column=13)
    ws.cell(row=instr_row, column=1).value = (
        'ИНСТРУКЦИЯ: Каждая строка — одна единица оборудования. '
        'Поля организации, руководителя и инженеров одинаковы для всех строк. '
        'Дата вводится в формате дд.мм.гггг.'
    )
    ws.cell(row=instr_row, column=1).font = Font(name='Times New Roman', italic=True,
                                                  size=9, color='888888')
    ws.cell(row=instr_row, column=1).alignment = LEFT
    ws.row_dimensions[instr_row].height = 22

    ws.freeze_panes = 'A2'
    wb.save(path)
    return path
