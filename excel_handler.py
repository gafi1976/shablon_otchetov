# excel_handler.py — Чтение данных из Excel шаблонов и сохранение шаблонов

import os
from openpyxl import Workbook, load_workbook
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, GradientFill
)
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation


# ══════════════════════════════════════════════════════
#  СТИЛИ
# ══════════════════════════════════════════════════════

HEADER_FILL_SPISAN = PatternFill('solid', fgColor='2E4057')
HEADER_FILL_UST    = PatternFill('solid', fgColor='1A5276')
HEADER_FONT        = Font(name='Times New Roman', bold=True, color='FFFFFF', size=10)
DATA_FONT          = Font(name='Times New Roman', size=10)
TITLE_FONT_SPISAN  = Font(name='Times New Roman', bold=True, size=14, color='2E4057')
TITLE_FONT_UST     = Font(name='Times New Roman', bold=True, size=14, color='1A5276')
CENTER             = Alignment(horizontal='center', vertical='center', wrap_text=True)
LEFT               = Alignment(horizontal='left',   vertical='center', wrap_text=True)

ROW_EVEN_SPISAN    = PatternFill('solid', fgColor='F0F4F8')
ROW_EVEN_UST       = PatternFill('solid', fgColor='EBF5FB')
ROW_ODD            = PatternFill('solid', fgColor='FFFFFF')

RED_FILL   = PatternFill('solid', fgColor='FADBD8')
RED_FONT   = Font(name='Times New Roman', color='C0392B', bold=True, size=10)
GREEN_FILL = PatternFill('solid', fgColor='D5F5E3')
GREEN_FONT = Font(name='Times New Roman', color='1E8B4C', bold=True, size=10)


def thin_border(color='2E4057'):
    side = Side(style='thin', color=color)
    return Border(left=side, right=side, top=side, bottom=side)


# ══════════════════════════════════════════════════════
#  СОЗДАНИЕ ШАБЛОНА EXCEL — СПИСАНИЕ
# ══════════════════════════════════════════════════════

def create_spisan_template(path: str):
    """Создаёт красивый Excel шаблон для ввода данных акта СПИСАНИЯ."""
    wb = Workbook()
    ws = wb.active
    ws.title = 'Данные для списания'
    ws.sheet_view.showGridLines = False

    # ── Заголовок листа ──
    ws.merge_cells('A1:J1')
    ws['A1'] = 'АКТ СПИСАНИЯ ОСНОВНЫХ СРЕДСТВ — ШАБЛОН ДАННЫХ'
    ws['A1'].font = TITLE_FONT_SPISAN
    ws['A1'].alignment = CENTER
    ws['A1'].fill = PatternFill('solid', fgColor='D6EAF8')
    ws.row_dimensions[1].height = 28

    # ── Блок мета-данных ──
    meta_labels = [
        ('B3', 'Организация:'),
        ('B4', 'Регион:'),
        ('B5', 'Номер акта:'),
        ('B6', 'Дата акта:'),
        ('B7', 'Председатель комиссии:'),
        ('B8', 'Член комиссии 1:'),
        ('B9', 'Член комиссии 2:'),
    ]
    meta_fill = PatternFill('solid', fgColor='EBF5FB')
    for cell_addr, label in meta_labels:
        ws[cell_addr] = label
        ws[cell_addr].font = Font(name='Times New Roman', bold=True, size=10, color='1A5276')
        ws[cell_addr].alignment = LEFT

        # Поле для ввода
        row = ws[cell_addr].row
        val_cell = ws.cell(row=row, column=3)
        ws.merge_cells(start_row=row, start_column=3, end_row=row, end_column=7)
        val_cell.fill = meta_fill
        val_cell.font = DATA_FONT
        val_cell.alignment = LEFT
        val_cell.border = thin_border('1A5276')
        ws.row_dimensions[row].height = 18

    ws.row_dimensions[2].height = 8
    ws.row_dimensions[10].height = 10

    # ── Заголовок таблицы оборудования ──
    columns_spisan = [
        ('A', '№',                   5),
        ('B', 'Наименование\nоборудования', 28),
        ('C', 'Инвентарный\nномер',   16),
        ('D', 'Год\nввода',           8),
        ('E', 'Перв. стоимость\n(сум)', 16),
        ('F', 'Ост. стоимость\n(сум)', 16),
        ('G', 'Причина\nсписания',    28),
        ('H', 'Примечание',           18),
    ]

    header_row = 11
    ws.row_dimensions[header_row].height = 32

    for col_letter, header, width in columns_spisan:
        col_idx = ord(col_letter) - ord('A') + 1
        cell = ws.cell(row=header_row, column=col_idx)
        cell.value = header
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL_SPISAN
        cell.alignment = CENTER
        cell.border = thin_border('2E4057')
        ws.column_dimensions[col_letter].width = width

    # ── 20 пустых строк для данных ──
    for i in range(1, 21):
        row_num = header_row + i
        ws.row_dimensions[row_num].height = 18
        fill = ROW_EVEN_SPISAN if i % 2 == 0 else ROW_ODD
        for col_idx in range(1, 9):
            cell = ws.cell(row=row_num, column=col_idx)
            cell.font = DATA_FONT
            cell.fill = fill
            cell.border = thin_border()
            cell.alignment = CENTER if col_idx in (1, 3, 4, 5, 6) else LEFT
            if col_idx == 1:
                cell.value = i

    # ── Инструкция ──
    ws.row_dimensions[header_row + 21].height = 10
    instr_row = header_row + 22
    ws.merge_cells(start_row=instr_row, start_column=1, end_row=instr_row, end_column=8)
    ws.cell(row=instr_row, column=1).value = (
        '⚠ Заполните данные организации (строки 3–9), '
        'затем внесите данные оборудования начиная со строки 12.'
    )
    ws.cell(row=instr_row, column=1).font = Font(name='Times New Roman', italic=True, size=9, color='888888')
    ws.cell(row=instr_row, column=1).alignment = LEFT
    ws.row_dimensions[instr_row].height = 16

    # ── Заморозить первые строки ──
    ws.freeze_panes = 'A12'

    wb.save(path)
    return path


# ══════════════════════════════════════════════════════
#  СОЗДАНИЕ ШАБЛОНА EXCEL — УСТАНОВКА
# ══════════════════════════════════════════════════════

def create_ust_template(path: str):
    """Создаёт красивый Excel шаблон для ввода данных акта УСТАНОВКИ."""
    wb = Workbook()
    ws = wb.active
    ws.title = 'Данные для установки'
    ws.sheet_view.showGridLines = False

    # ── Заголовок ──
    ws.merge_cells('A1:K1')
    ws['A1'] = 'АКТ ВВОДА В ЭКСПЛУАТАЦИЮ — ШАБЛОН ДАННЫХ'
    ws['A1'].font = TITLE_FONT_UST
    ws['A1'].alignment = CENTER
    ws['A1'].fill = PatternFill('solid', fgColor='D6EAF8')
    ws.row_dimensions[1].height = 28

    # ── Мета-данные ──
    meta_labels_ust = [
        ('B3',  'Организация:'),
        ('B4',  'Регион:'),
        ('B5',  'Номер акта:'),
        ('B6',  'Дата акта:'),
        ('B7',  'Место установки:'),
        ('B8',  'Председатель комиссии:'),
        ('B9',  'Член комиссии 1:'),
        ('B10', 'Член комиссии 2:'),
    ]
    meta_fill = PatternFill('solid', fgColor='EBF5FB')
    for cell_addr, label in meta_labels_ust:
        ws[cell_addr] = label
        ws[cell_addr].font = Font(name='Times New Roman', bold=True, size=10, color='1A5276')
        ws[cell_addr].alignment = LEFT
        row = ws[cell_addr].row
        val_cell = ws.cell(row=row, column=3)
        ws.merge_cells(start_row=row, start_column=3, end_row=row, end_column=8)
        val_cell.fill = meta_fill
        val_cell.font = DATA_FONT
        val_cell.alignment = LEFT
        val_cell.border = thin_border('1A5276')
        ws.row_dimensions[row].height = 18

    ws.row_dimensions[2].height = 8
    ws.row_dimensions[11].height = 10

    # ── Заголовок таблицы ──
    columns_ust = [
        ('A', '№',                5),
        ('B', 'Наименование\nоборудования', 25),
        ('C', 'Модель /\nМарка',  16),
        ('D', 'Серийный\nномер',  18),
        ('E', 'Инв.\nномер',      14),
        ('F', 'Год\nвыпуска',      8),
        ('G', 'Стоимость\n(сум)', 16),
        ('H', 'Дата\nустановки',  14),
        ('I', 'Состояние',        14),
        ('J', 'Примечание',       16),
    ]

    header_row = 12
    ws.row_dimensions[header_row].height = 32

    for col_letter, header, width in columns_ust:
        col_idx = ord(col_letter) - ord('A') + 1
        cell = ws.cell(row=header_row, column=col_idx)
        cell.value = header
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL_UST
        cell.alignment = CENTER
        cell.border = thin_border('1A5276')
        ws.column_dimensions[col_letter].width = width

    # ── Валидация для колонки "Состояние" ──
    dv = DataValidation(
        type='list',
        formula1='"Янги (новый),Ишлатилган (б/у),Яроқсиз (неисправный)"',
        allow_blank=True,
        showDropDown=False,
    )
    dv.sqref = f'I{header_row + 1}:I{header_row + 20}'
    ws.add_data_validation(dv)

    # ── 20 строк для данных ──
    for i in range(1, 21):
        row_num = header_row + i
        ws.row_dimensions[row_num].height = 18
        fill = ROW_EVEN_UST if i % 2 == 0 else ROW_ODD
        for col_idx in range(1, 11):
            cell = ws.cell(row=row_num, column=col_idx)
            cell.font = DATA_FONT
            cell.fill = fill
            cell.border = thin_border('1A5276')
            cell.alignment = CENTER if col_idx in (1, 4, 5, 6, 7, 8) else LEFT
            if col_idx == 1:
                cell.value = i

    # ── Инструкция ──
    ws.row_dimensions[header_row + 21].height = 10
    instr_row = header_row + 22
    ws.merge_cells(start_row=instr_row, start_column=1, end_row=instr_row, end_column=10)
    ws.cell(row=instr_row, column=1).value = (
        '⚠ Заполните данные организации (строки 3–10), '
        'затем внесите данные оборудования начиная со строки 13. '
        'Для колонки "Состояние" используйте выпадающий список.'
    )
    ws.cell(row=instr_row, column=1).font = Font(name='Times New Roman', italic=True, size=9, color='888888')
    ws.cell(row=instr_row, column=1).alignment = LEFT
    ws.row_dimensions[instr_row].height = 16

    ws.freeze_panes = 'A13'
    wb.save(path)
    return path


# ══════════════════════════════════════════════════════
#  ЧТЕНИЕ ДАННЫХ ИЗ EXCEL — СПИСАНИЕ
# ══════════════════════════════════════════════════════

def read_spisan_excel(path: str) -> dict:
    """
    Читает данные из Excel шаблона списания.
    Возвращает словарь data, совместимый с generator_spisan.create_spisan_doc().
    """
    wb = load_workbook(path, data_only=True)
    ws = wb.active

    def val(row, col):
        v = ws.cell(row=row, column=col).value
        return str(v).strip() if v is not None else ''

    data = {
        'org_name':        val(3, 3),
        'region':          val(4, 3),
        'doc_number':      val(5, 3),
        'doc_date':        val(6, 3),
        'commission_head': val(7, 3),
        'member1':         val(8, 3),
        'member2':         val(9, 3),
        'items': [],
    }

    # Данные начинаются со строки 12 (header=11, data=12+)
    for row_num in range(12, ws.max_row + 1):
        name = val(row_num, 2)
        if not name:
            continue
        data['items'].append({
            'num':           val(row_num, 1) or str(len(data['items']) + 1),
            'name':          name,
            'inv_number':    val(row_num, 3),
            'year':          val(row_num, 4),
            'initial_cost':  val(row_num, 5),
            'residual_cost': val(row_num, 6),
            'reason':        val(row_num, 7),
            'note':          val(row_num, 8),
        })

    return data


# ══════════════════════════════════════════════════════
#  ЧТЕНИЕ ДАННЫХ ИЗ EXCEL — УСТАНОВКА
# ══════════════════════════════════════════════════════

def read_ust_excel(path: str) -> dict:
    """
    Читает данные из Excel шаблона установки.
    Возвращает словарь data, совместимый с generator_ust.create_ust_doc().
    """
    wb = load_workbook(path, data_only=True)
    ws = wb.active

    def val(row, col):
        v = ws.cell(row=row, column=col).value
        return str(v).strip() if v is not None else ''

    data = {
        'org_name':        val(3, 3),
        'region':          val(4, 3),
        'doc_number':      val(5, 3),
        'doc_date':        val(6, 3),
        'location':        val(7, 3),
        'commission_head': val(8, 3),
        'member1':         val(9, 3),
        'member2':         val(10, 3),
        'items': [],
    }

    for row_num in range(13, ws.max_row + 1):
        name = val(row_num, 2)
        if not name:
            continue
        data['items'].append({
            'num':           val(row_num, 1) or str(len(data['items']) + 1),
            'name':          name,
            'model':         val(row_num, 3),
            'serial_number': val(row_num, 4),
            'inv_number':    val(row_num, 5),
            'year':          val(row_num, 6),
            'cost':          val(row_num, 7),
            'install_date':  val(row_num, 8),
            'condition':     val(row_num, 9),
            'note':          val(row_num, 10),
        })

    return data
