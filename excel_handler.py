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

    for row in rows[2:]:            # строка 1=title, строка 2=заголовки, данные с строки 3
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

    ТОЧНАЯ структура колонок (единая для оригинала и сгенерированного шаблона):
      A(0)  = Data              — дата документа
      B(1)  = {num_otch}        — порядковый номер
      C(2)  = installation date — дата установки (если пусто → берём из A)
      D(3)  = {oborud name}     — наименование/модель оборудования
      E(4)  = {serial num}      — серийный номер
      F(5)  = {where oborud}    — место установки
      G(6)  = {Organizasiya}    — организация
      H(7)  = {adress}          — адрес
      I(8)  = пустая колонка    — намеренный пропуск (как в оригинале)
      J(9)  = {boss name}       — руководитель
      K(10) = {engine1}         — инженер 1
      L(11) = {job title1}      — должность 1
      M(12) = {engine2}         — инженер 2
      N(13) = {job title2}      — должность 2
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
    doc_date  = ''
    items     = []

    def _to_date(val: str) -> str:
        if not val:
            return ''
        val = val.strip()
        if len(val) == 10 and val[2] == '.' and val[5] == '.':
            return val
        if val.replace('.', '').isdigit() and '.' not in val:
            return _excel_date(val)
        return val

    for row in rows[2:]:        # строка 1=title, строка 2=заголовки, данные с строки 3
        if not row or not any(row):
            continue

        raw_doc_date  = _s(row,  0)   # A: дата документа
        num           = _s(row,  1)   # B: №
        raw_inst_date = _s(row,  2)   # C: дата установки
        oborud        = _s(row,  3)   # D: модель оборудования
        serial        = _s(row,  4)   # E: серийный номер
        where         = _s(row,  5)   # F: место установки
        org           = _s(row,  6)   # G: организация
        addr          = _s(row,  7)   # H: адрес
        # I(8) = пустая колонка (в обоих файлах)
        bss           = _s(row,  9)   # J: руководитель
        e1            = _s(row, 10)   # K: инженер 1
        j1            = _s(row, 11)   # L: должность 1
        e2            = _s(row, 12)   # M: инженер 2
        j2            = _s(row, 13)   # N: должность 2

        if not doc_date and raw_doc_date:
            doc_date = _to_date(raw_doc_date)

        install_date = (_to_date(raw_inst_date) if raw_inst_date
                        else _to_date(raw_doc_date) if raw_doc_date
                        else datetime.now().strftime('%d.%m.%Y'))

        if not org_name and org:  org_name = org
        if not address  and addr: address  = addr
        if not boss     and bss:  boss     = bss
        if not eng1     and e1:   eng1 = e1; job1 = j1
        if not eng2     and e2:   eng2 = e2; job2 = j2

        if not oborud:
            continue

        items.append({
            'num':           num or str(len(items) + 1),
            'name':          oborud,
            'model':         oborud,
            'serial_number': serial,
            'inv_number':    '',
            'install_date':  install_date,
            'location':      where,
            'cost':          '',
            'condition':     'Yangi',
            'note':          where,
        })

    if not doc_date:
        doc_date = datetime.now().strftime('%d.%m.%Y')

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
    """Универсальный Excel шаблон для акта СПИСАНИЯ."""
    wb = Workbook()
    ws = wb.active
    ws.title = 'Spisaniye'
    ws.sheet_view.showGridLines = False

    # Заголовок листа (строка 1)
    ws.merge_cells('A1:I1')
    ws['A1'] = 'АКТ СПИСАНИЯ — ШАБЛОН ДАННЫХ'
    ws['A1'].font      = Font(name='Times New Roman', bold=True, size=13, color='FFFFFF')
    ws['A1'].fill      = PatternFill('solid', fgColor='2E4057')
    ws['A1'].alignment = _CENTER
    ws.row_dimensions[1].height = 26

    # Заголовки колонок (строка 2)
    cols_s = [
        ('A', 'Qurilma nomi\n(Оборудование)',          26),
        ('B', 'Qismlar nomi\n(Компонент)',              20),
        ('C', 'Holati\n(Yaroqli / Yaroqsiz)',           16),
        ('D', 'Nosozlik sababi\n(Причина неисправн.)',  26),
        ('E', 'Inventar raqami\n(Инв. номер)  ★',      18),
        ('F', 'Tashkilot nomi\n(Организация)',          28),
        ('G', 'Rahbar\n(Руководитель)',                  22),
        ('H', 'Muhandis 1\n(Инженер 1)',                22),
        ('I', 'Muhandis 2\n(Инженер 2)',                22),
    ]
    for ci, (col, text, width) in enumerate(cols_s, start=1):
        cell = ws.cell(row=2, column=ci, value=text)
        cell.font = _HDR_FONT
        cell.fill = _HDR_FILL_S
        cell.alignment = _CENTER
        cell.border = _border()
        ws.column_dimensions[col].width = width
    ws.row_dimensions[2].height = 34

    # Валидация состояния
    dv = DataValidation(type='list', formula1='"Yaroqli,Yaroqsiz"',
                        allow_blank=True, showDropDown=False)
    dv.sqref = 'C3:C500'
    ws.add_data_validation(dv)

    # Примеры данных (данные начинаются с строки 3)
    ORG = 'ABM Sirdaryo viloyati MChJ'
    B = 'Palonov P.A'; E1 = 'Raxmatov V.A'; E2 = 'Xolbekov G.T'
    ex_s = [
        ('Server HP ML 350',     'Asosiy plata', 'Yaroqli',  "Ma'naviy eskirgan", '26-0006039', ORG, B, E1, E2),
        ('Server HP ML 350',     'Tok manbayi',  'Yaroqli',  "Ma'naviy eskirgan", '26-0006039', ORG, B, E1, E2),
        ('Server HP ML 350',     'Qattiq disk',  'Yaroqsiz', 'BAD bloklar mavjud','26-0006039', ORG, B, E1, E2),
        (None,)*9,
        ('Konditsioner ART-12HI','Tok manbai',   'Yaroqsiz', "Tok sarfi ko'p",    '24-0006162', ORG, B, E1, E2),
        ('Konditsioner ART-12HI','Kompressor',   'Yaroqsiz', 'Shovqin mavjud',    '24-0006162', ORG, B, E1, E2),
        ('Konditsioner ART-12HI','Radiator',     'Yaroqsiz', 'Yamalgan',          '24-0006162', ORG, B, E1, E2),
        (None,)*9,
        ("Monitor LCD 19''",     'Asosiy plata', 'Yaroqsiz', 'Korpus singan',     '26-0003436', ORG, B, E1, E2),
        ("Monitor LCD 19''",     'Tok manbayi',  'Yaroqsiz', 'Kuygan',            '26-0003436', ORG, B, E1, E2),
    ]

    for ri, row_data in enumerate(ex_s):
        rn = ri + 3
        empty = not any(v for v in row_data if v)
        fill  = _EVEN_S if (ri % 2 == 0 and not empty) else _ODD
        ws.row_dimensions[rn].height = 18
        for ci, val in enumerate(row_data, start=1):
            cell = ws.cell(row=rn, column=ci, value=val)
            cell.font = _DATA_FONT; cell.fill = fill
            cell.border = _border(); cell.alignment = _CENTER if ci == 3 else _LEFT
            if ci == 3 and val == 'Yaroqsiz':
                cell.font = Font(name='Times New Roman', size=10, color='C0392B', bold=True)
            if ci == 3 and val == 'Yaroqli':
                cell.font = Font(name='Times New Roman', size=10, color='1E8B4C', bold=True)

    for ri in range(len(ex_s), len(ex_s) + 20):
        rn = ri + 3
        fill = _EVEN_S if ri % 2 == 0 else _ODD
        ws.row_dimensions[rn].height = 18
        for ci in range(1, 10):
            cell = ws.cell(row=rn, column=ci)
            cell.font = _DATA_FONT; cell.fill = fill
            cell.border = _border(); cell.alignment = _LEFT

    ir = len(ex_s) + 24
    ws.merge_cells(start_row=ir, start_column=1, end_row=ir, end_column=9)
    ws.cell(row=ir, column=1).value = (
        '★ Каждая строка = один компонент оборудования.  '
        'Группировка в Word-файл по колонке E (Inventar raqami).  '
        'Пустая строка — разделитель.  Holati: Yaroqli = исправен, Yaroqsiz = неисправен.')
    ws.cell(row=ir, column=1).font = Font(name='Times New Roman', italic=True, size=9, color='888888')
    ws.cell(row=ir, column=1).alignment = _LEFT
    ws.row_dimensions[ir].height = 18
    ws.freeze_panes = 'A3'
    wb.save(path)
    return path


# ══════════════════════════════════════════════
#  СОЗДАНИЕ ШАБЛОНА EXCEL — УСТАНОВКА
#  Структура = shablom_ust.xlsx
# ══════════════════════════════════════════════


def create_ust_template(path: str) -> str:
    """Универсальный Excel шаблон для акта УСТАНОВКИ."""
    wb = Workbook()
    ws = wb.active
    ws.title = 'Ustanovka'
    ws.sheet_view.showGridLines = False

    # Заголовок листа (строка 1)
    ws.merge_cells('A1:N1')
    ws['A1'] = 'АКТ УСТАНОВКИ — ШАБЛОН ДАННЫХ'
    ws['A1'].font      = Font(name='Times New Roman', bold=True, size=13, color='FFFFFF')
    ws['A1'].fill      = PatternFill('solid', fgColor='1A5276')
    ws['A1'].alignment = _CENTER
    ws.row_dimensions[1].height = 26

    # Заголовки колонок (строка 2)
    cols_u = [
        ('A', 'Sana (akt)\n(Дата акта)',             14),
        ('B', '№',                                     4),
        ('C', "O'rnatish sanasi\n(Вр. установки)",   16),
        ('D', 'Uskuna nomi\n(Оборудование / Модель)', 26),
        ('E', 'Seriya raqami\n(Серийный номер)',       18),
        ('F', "O'rnatish joyi\n(Место установки)",    24),
        ('G', 'Tashkilot nomi\n(Организация)',         28),
        ('H', 'Manzil\n(Адрес)',                       22),
        ('I', '',                                       4),
        ('J', 'Rahbar\n(Руководитель)',                22),
        ('K', 'Muhandis 1\n(Инженер 1)',               20),
        ('L', "Muhandis 1\nlavozimi (Должность)",      18),
        ('M', 'Muhandis 2\n(Инженер 2)',               20),
        ('N', "Muhandis 2\nlavozimi (Должность)",      18),
    ]
    for ci, (col, text, width) in enumerate(cols_u, start=1):
        ws.column_dimensions[col].width = width
        if text:
            cell = ws.cell(row=2, column=ci, value=text)
            cell.font = _HDR_FONT; cell.fill = _HDR_FILL_U
            cell.alignment = _CENTER; cell.border = _border('1A5276')
    ws.row_dimensions[2].height = 34

    # Примеры данных (с строки 3)
    today = datetime.now().strftime('%d.%m.%Y')
    ORG = 'ABM Sirdaryo viloyati MChJ'
    ADDR = "Guliston sh., O'zbekiston ko'ch."
    B = 'Palonov P.A'; E1 = 'Raxmatov V.A'; L1 = 'Yetakchi muhandis'
    E2 = 'Xolbekov G.T'; L2 = 'Yetakchi muhandis'
    ex_u = [
        (today,'1',today,'Maipu MP1900X-22',  'SN-001','Guliston, 1-bino',ORG,ADDR,'',B,E1,L1,E2,L2),
        (today,'2',today,'Maipu MP1900X-22',  'SN-002','Guliston, 2-bino',ORG,ADDR,'',B,E1,L1,E2,L2),
        (today,'3',today,'Switch Cisco SG350','SN-003','Guliston, 3-bino',ORG,ADDR,'',B,E1,L1,E2,L2),
        (today,'4',today,'UPS APC 1500VA',    'SN-004','Server xona',     ORG,ADDR,'',B,E1,L1,E2,L2),
    ]

    for ri, row_data in enumerate(ex_u):
        rn = ri + 3
        fill = _EVEN_U if ri % 2 == 0 else _ODD
        ws.row_dimensions[rn].height = 20
        for ci, val in enumerate(row_data, start=1):
            cell = ws.cell(row=rn, column=ci, value=val if val else None)
            cell.font = _DATA_FONT; cell.fill = fill
            cell.border = _border('1A5276')
            cell.alignment = _CENTER if ci in (1,2,3,9) else _LEFT

    for ri in range(len(ex_u), len(ex_u) + 20):
        rn = ri + 3
        fill = _EVEN_U if ri % 2 == 0 else _ODD
        ws.row_dimensions[rn].height = 20
        for ci in range(1, 15):
            cell = ws.cell(row=rn, column=ci)
            cell.font = _DATA_FONT; cell.fill = fill
            cell.border = _border('1A5276'); cell.alignment = _LEFT

    ir = len(ex_u) + 24
    ws.merge_cells(start_row=ir, start_column=1, end_row=ir, end_column=14)
    ws.cell(row=ir, column=1).value = (
        'A=Дата акта,  C=Вр. установки (дд.мм.гггг),  D=Модель,  '
        'E=Серийный №,  F=Место установки,  G=Организация,  H=Адрес,  '
        'J=Руководитель,  K/M=Инженеры,  L/N=Должности.  Колонку I не заполнять!')
    ws.cell(row=ir, column=1).font = Font(name='Times New Roman', italic=True, size=9, color='888888')
    ws.cell(row=ir, column=1).alignment = _LEFT
    ws.row_dimensions[ir].height = 18
    ws.freeze_panes = 'A3'
    wb.save(path)
    return path
