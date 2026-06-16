"""
app.py — Главное приложение: Генератор актов списания и установки оборудования
Графическая оболочка на Tkinter
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import sys
import threading
from datetime import datetime

# ── Импорт модулей генерации ──
try:
    import generator_spisan
    import generator_ust
    import excel_handler
except ImportError as e:
    print(f"Ошибка импорта: {e}")
    sys.exit(1)


# ══════════════════════════════════════════════════════
#  ЦВЕТОВАЯ ПАЛИТРА И КОНСТАНТЫ
# ══════════════════════════════════════════════════════

COLORS = {
    'bg':           '#F5F7FA',
    'sidebar':      '#1A2B45',
    'sidebar_text': '#FFFFFF',
    'sidebar_hover':'#2E4A6E',
    'accent_red':   '#2E4057',    # цвет списания
    'accent_blue':  '#1A5276',    # цвет установки
    'accent_green': '#1E8B4C',
    'card_bg':      '#FFFFFF',
    'border':       '#DDE3EE',
    'text':         '#1A2B45',
    'text_muted':   '#7F8C9A',
    'entry_bg':     '#FFFFFF',
    'entry_border': '#B0BEC5',
    'btn_primary':  '#1A5276',
    'btn_danger':   '#2E4057',
    'btn_success':  '#1E8B4C',
    'btn_text':     '#FFFFFF',
    'header_bg':    '#1A2B45',
    'header_text':  '#FFFFFF',
    'tab_active':   '#1A5276',
    'highlight':    '#EBF5FB',
}

FONT_TITLE    = ('Segoe UI', 22, 'bold')
FONT_SUBTITLE = ('Segoe UI', 13)
FONT_HEADER   = ('Segoe UI', 11, 'bold')
FONT_BODY     = ('Segoe UI', 10)
FONT_SMALL    = ('Segoe UI', 9)
FONT_MONO     = ('Consolas', 9)
FONT_BTN      = ('Segoe UI', 10, 'bold')

WINDOW_W = 1180
WINDOW_H = 760


# ══════════════════════════════════════════════════════
#  ВСПОМОГАТЕЛЬНЫЕ ВИДЖЕТЫ
# ══════════════════════════════════════════════════════

class RoundedButton(tk.Button):
    """Стилизованная кнопка (совместима со всеми версиями Tkinter/Windows)."""

    def __init__(self, parent, text='', command=None, bg='#1A5276',
                 fg='#FFFFFF', width=160, height=36, radius=8,
                 font=FONT_BTN, **kwargs):
        # Вычисляем примерную ширину в символах из пикселей
        char_width = max(8, width // 8)

        def _get_parent_bg(p):
            try:
                return p['bg']
            except Exception:
                return '#F5F7FA'

        super().__init__(
            parent,
            text=text,
            command=command,
            bg=bg,
            fg=fg,
            font=font,
            relief='flat',
            bd=0,
            cursor='hand2',
            activebackground=self._darken(bg),
            activeforeground=fg,
            padx=8,
            pady=4,
        )
        self._bg = bg
        self._hover_bg = self._darken(bg)

        self.bind('<Enter>', lambda e: self.config(bg=self._hover_bg))
        self.bind('<Leave>', lambda e: self.config(bg=self._bg))

    @staticmethod
    def _darken(hex_color):
        try:
            r = int(hex_color[1:3], 16)
            g = int(hex_color[3:5], 16)
            b = int(hex_color[5:7], 16)
            factor = 0.82
            return f'#{int(r*factor):02x}{int(g*factor):02x}{int(b*factor):02x}'
        except Exception:
            return hex_color

    def config_text(self, text):
        self.config(text=text)


class LabeledEntry(tk.Frame):
    """Поле ввода с подписью сверху."""

    def __init__(self, parent, label='', width=30, bg=None, **kwargs):
        bg = bg or COLORS['bg']
        super().__init__(parent, bg=bg)
        tk.Label(self, text=label, bg=bg, fg=COLORS['text_muted'],
                 font=FONT_SMALL).pack(anchor='w', pady=(0, 2))
        self.var = tk.StringVar()
        self.entry = tk.Entry(self, textvariable=self.var, width=width,
                              bg=COLORS['entry_bg'], fg=COLORS['text'],
                              relief='flat', font=FONT_BODY,
                              highlightthickness=1,
                              highlightbackground=COLORS['entry_border'],
                              highlightcolor=COLORS['tab_active'],
                              **kwargs)
        self.entry.pack(fill='x', ipady=5)

    def get(self):
        return self.var.get().strip()

    def set(self, val):
        self.var.set(val)


class SectionCard(tk.Frame):
    """Карточка-секция с заголовком и отступами."""

    def __init__(self, parent, title='', accent='#1A5276', bg=None):
        bg = bg or COLORS['card_bg']
        super().__init__(parent, bg=bg, bd=0,
                         highlightthickness=1,
                         highlightbackground=COLORS['border'])
        # Полоска-акцент сверху
        tk.Frame(self, bg=accent, height=4).pack(fill='x')
        if title:
            tk.Label(self, text=title, bg=bg, fg=accent,
                     font=FONT_HEADER).pack(anchor='w', padx=12, pady=(8, 4))
        self.inner = tk.Frame(self, bg=bg)
        self.inner.pack(fill='both', expand=True, padx=12, pady=(0, 12))


class StatusBar(tk.Frame):
    """Нижняя строка статуса."""

    def __init__(self, parent):
        super().__init__(parent, bg=COLORS['header_bg'], height=28)
        self._var = tk.StringVar(value='Готов к работе')
        tk.Label(self, textvariable=self._var, bg=COLORS['header_bg'],
                 fg='#A8C7E8', font=FONT_SMALL, anchor='w').pack(
            side='left', padx=12)
        self._clock = tk.Label(self, bg=COLORS['header_bg'],
                               fg='#A8C7E8', font=FONT_SMALL)
        self._clock.pack(side='right', padx=12)
        self._tick()

    def _tick(self):
        self._clock.config(text=datetime.now().strftime('%d.%m.%Y  %H:%M:%S'))
        self.after(1000, self._tick)

    def set(self, msg, color='#A8C7E8'):
        self._var.set(msg)


# ══════════════════════════════════════════════════════
#  ВКЛАДКА — ДАННЫЕ ОРГАНИЗАЦИИ (общая форма)
# ══════════════════════════════════════════════════════

class OrgForm(tk.Frame):
    """Блок ввода данных организации."""

    def __init__(self, parent, accent='#1A5276', bg=None):
        bg = bg or COLORS['bg']
        super().__init__(parent, bg=bg)
        card = SectionCard(self, title='Данные организации', accent=accent, bg=bg)
        card.pack(fill='x', pady=(0, 8))

        row1 = tk.Frame(card.inner, bg=bg)
        row1.pack(fill='x', pady=2)
        row2 = tk.Frame(card.inner, bg=bg)
        row2.pack(fill='x', pady=2)
        row3 = tk.Frame(card.inner, bg=bg)
        row3.pack(fill='x', pady=2)

        self.org_name = LabeledEntry(row1, 'Название организации', width=40, bg=bg)
        self.org_name.pack(side='left', padx=(0, 12), fill='x', expand=True)

        self.region = LabeledEntry(row1, 'Регион', width=28, bg=bg)
        self.region.pack(side='left', fill='x', expand=True)

        self.doc_number = LabeledEntry(row2, 'Номер акта', width=20, bg=bg)
        self.doc_number.pack(side='left', padx=(0, 12))
        self.doc_number.set(f'{datetime.now().year}-001')

        self.doc_date = LabeledEntry(row2, 'Дата акта (дд.мм.гггг)', width=20, bg=bg)
        self.doc_date.pack(side='left', padx=(0, 12))
        self.doc_date.set(datetime.now().strftime('%d.%m.%Y'))

        self.comm_head = LabeledEntry(row3, 'Председатель комиссии', width=36, bg=bg)
        self.comm_head.pack(side='left', padx=(0, 12), fill='x', expand=True)

        self.member1 = LabeledEntry(row3, 'Член комиссии 1', width=28, bg=bg)
        self.member1.pack(side='left', padx=(0, 12), fill='x', expand=True)

        self.member2 = LabeledEntry(row3, 'Член комиссии 2', width=28, bg=bg)
        self.member2.pack(side='left', fill='x', expand=True)

    def get_data(self):
        return {
            'org_name':        self.org_name.get(),
            'region':          self.region.get(),
            'doc_number':      self.doc_number.get(),
            'doc_date':        self.doc_date.get(),
            'commission_head': self.comm_head.get(),
            'member1':         self.member1.get(),
            'member2':         self.member2.get(),
        }

    def set_data(self, d: dict):
        self.org_name.set(d.get('org_name', ''))
        self.region.set(d.get('region', ''))
        self.doc_number.set(d.get('doc_number', ''))
        self.doc_date.set(d.get('doc_date', ''))
        self.comm_head.set(d.get('commission_head', ''))
        self.member1.set(d.get('member1', ''))
        self.member2.set(d.get('member2', ''))


# ══════════════════════════════════════════════════════
#  ТАБЛИЦА РЕДАКТИРОВАНИЯ (общая)
# ══════════════════════════════════════════════════════

class ItemsTable(tk.Frame):
    """Таблица для ввода/редактирования строк оборудования."""

    def __init__(self, parent, columns: list, accent='#1A5276', bg=None):
        """
        columns: список (key, header, width) — ключ, заголовок, ширина колонки
        """
        bg = bg or COLORS['bg']
        super().__init__(parent, bg=bg)
        self._columns = columns
        self._accent = accent
        self._bg = bg
        self._rows_data = []   # список StringVar строк

        card = SectionCard(self, title='Оборудование', accent=accent, bg=bg)
        card.pack(fill='both', expand=True)

        # Toolbar
        toolbar = tk.Frame(card.inner, bg=bg)
        toolbar.pack(fill='x', pady=(0, 6))

        RoundedButton(toolbar, text='+ Добавить строку',
                      command=self._add_row, bg=accent,
                      width=150, height=30).pack(side='left', padx=(0, 8))
        RoundedButton(toolbar, text='✕ Удалить выбранную',
                      command=self._del_row, bg='#A93226',
                      width=160, height=30).pack(side='left', padx=(0, 8))
        RoundedButton(toolbar, text='↺ Очистить всё',
                      command=self._clear_all, bg='#5D6D7E',
                      width=140, height=30).pack(side='left')

        # Treeview + Scrollbars
        frame_tv = tk.Frame(card.inner, bg=bg)
        frame_tv.pack(fill='both', expand=True)

        col_keys = [c[0] for c in columns]
        col_heads = {c[0]: c[1] for c in columns}
        col_widths = {c[0]: c[2] for c in columns}

        self.tv = ttk.Treeview(frame_tv, columns=col_keys, show='headings',
                               height=10, selectmode='browse')
        for key in col_keys:
            self.tv.heading(key, text=col_heads[key])
            self.tv.column(key, width=col_widths[key], minwidth=40, anchor='center')

        vsb = ttk.Scrollbar(frame_tv, orient='vertical', command=self.tv.yview)
        hsb = ttk.Scrollbar(frame_tv, orient='horizontal', command=self.tv.xview)
        self.tv.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tv.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        frame_tv.grid_rowconfigure(0, weight=1)
        frame_tv.grid_columnconfigure(0, weight=1)

        self.tv.bind('<Double-1>', self._on_double_click)
        self._style_treeview()

        # Счётчик строк
        self._count_var = tk.StringVar(value='Строк: 0')
        tk.Label(card.inner, textvariable=self._count_var,
                 bg=bg, fg=COLORS['text_muted'], font=FONT_SMALL).pack(
            anchor='e', pady=(4, 0))

    def _style_treeview(self):
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('Treeview',
                         background=COLORS['card_bg'],
                         foreground=COLORS['text'],
                         fieldbackground=COLORS['card_bg'],
                         rowheight=24,
                         font=FONT_BODY)
        style.configure('Treeview.Heading',
                         background=self._accent,
                         foreground='white',
                         font=FONT_HEADER,
                         relief='flat')
        style.map('Treeview',
                  background=[('selected', self._accent)],
                  foreground=[('selected', 'white')])
        style.map('Treeview.Heading',
                  background=[('active', self._accent)])

    def _add_row(self, values=None):
        """Открывает диалог добавления/редактирования строки."""
        self._open_edit_dialog(values, editing=False)

    def _del_row(self):
        sel = self.tv.selection()
        if not sel:
            messagebox.showinfo('Информация', 'Выберите строку для удаления')
            return
        for item in sel:
            self.tv.delete(item)
        self._reindex()
        self._update_count()

    def _clear_all(self):
        if messagebox.askyesno('Подтверждение', 'Очистить все строки?'):
            for item in self.tv.get_children():
                self.tv.delete(item)
            self._update_count()

    def _on_double_click(self, event):
        region = self.tv.identify_region(event.x, event.y)
        if region != 'cell':
            return
        item = self.tv.selection()
        if not item:
            return
        values = self.tv.item(item[0], 'values')
        self._open_edit_dialog(list(values), editing=True, item_id=item[0])

    def _reindex(self):
        for i, item in enumerate(self.tv.get_children(), 1):
            vals = list(self.tv.item(item, 'values'))
            vals[0] = str(i)
            self.tv.item(item, values=vals)

    def _update_count(self):
        n = len(self.tv.get_children())
        self._count_var.set(f'Строк: {n}')

    def _open_edit_dialog(self, values=None, editing=False, item_id=None):
        """Диалоговое окно для добавления/редактирования строки."""
        dlg = tk.Toplevel()
        dlg.title('Редактировать строку' if editing else 'Добавить строку')
        dlg.configure(bg=COLORS['bg'])
        dlg.grab_set()
        dlg.resizable(True, True)

        # Центрирование
        dlg.update_idletasks()
        w, h = 540, 420
        x = dlg.winfo_screenwidth() // 2 - w // 2
        y = dlg.winfo_screenheight() // 2 - h // 2
        dlg.geometry(f'{w}x{h}+{x}+{y}')

        tk.Label(dlg, text='Редактировать строку' if editing else 'Новая строка',
                 bg=COLORS['bg'], fg=COLORS['text'],
                 font=FONT_HEADER).pack(pady=(14, 6))

        frame = tk.Frame(dlg, bg=COLORS['bg'])
        frame.pack(fill='both', expand=True, padx=16, pady=4)

        entries = []
        cols = self._columns
        skip_first = True  # первая колонка — №, автоматически

        row_idx = 0
        col_idx_gui = 0
        for i, (key, header, width) in enumerate(cols):
            if i == 0:
                entries.append(None)   # placeholder для №
                continue
            r = row_idx
            c = col_idx_gui % 2
            sub = tk.Frame(frame, bg=COLORS['bg'])
            sub.grid(row=r, column=c, padx=6, pady=4, sticky='ew')
            tk.Label(sub, text=header.replace('\n', ' '), bg=COLORS['bg'],
                     fg=COLORS['text_muted'], font=FONT_SMALL).pack(anchor='w')
            var = tk.StringVar()
            if values and i < len(values):
                var.set(values[i])
            entry = tk.Entry(sub, textvariable=var, font=FONT_BODY,
                             bg=COLORS['entry_bg'], fg=COLORS['text'],
                             relief='flat', highlightthickness=1,
                             highlightbackground=COLORS['entry_border'],
                             highlightcolor=self._accent)
            entry.pack(fill='x', ipady=4)
            entries.append((var, entry))
            col_idx_gui += 1
            if col_idx_gui % 2 == 0:
                row_idx += 1

        frame.grid_columnconfigure(0, weight=1)
        frame.grid_columnconfigure(1, weight=1)

        def on_save():
            row_vals = []
            n = len(self.tv.get_children()) + (0 if editing else 1)
            row_vals.append(str(n))
            for e in entries[1:]:
                if e is None:
                    row_vals.append('')
                else:
                    row_vals.append(e[0].get().strip())

            if editing and item_id:
                self.tv.item(item_id, values=row_vals)
            else:
                tag = 'even' if n % 2 == 0 else 'odd'
                self.tv.insert('', 'end', values=row_vals, tags=(tag,))
                self.tv.tag_configure('even', background=COLORS['highlight'])
                self.tv.tag_configure('odd', background=COLORS['card_bg'])
            self._update_count()
            dlg.destroy()

        btn_frame = tk.Frame(dlg, bg=COLORS['bg'])
        btn_frame.pack(pady=10)
        RoundedButton(btn_frame, text='Сохранить', command=on_save,
                      bg=self._accent, width=120, height=32).pack(side='left', padx=6)
        RoundedButton(btn_frame, text='Отмена', command=dlg.destroy,
                      bg='#5D6D7E', width=100, height=32).pack(side='left', padx=6)

        # Фокус на первом поле
        for e in entries[1:]:
            if e:
                e[1].focus_set()
                break

    def get_items(self) -> list:
        """Возвращает список словарей по строкам таблицы."""
        items = []
        col_keys = [c[0] for c in self._columns]
        for item in self.tv.get_children():
            vals = self.tv.item(item, 'values')
            d = {k: (vals[i] if i < len(vals) else '') for i, k in enumerate(col_keys)}
            items.append(d)
        return items

    def load_items(self, items: list):
        """Загружает список данных в таблицу."""
        for item in self.tv.get_children():
            self.tv.delete(item)
        col_keys = [c[0] for c in self._columns]
        for i, d in enumerate(items):
            vals = [d.get(k, '') for k in col_keys]
            tag = 'even' if i % 2 == 0 else 'odd'
            self.tv.insert('', 'end', values=vals, tags=(tag,))
            self.tv.tag_configure('even', background=COLORS['highlight'])
            self.tv.tag_configure('odd', background=COLORS['card_bg'])
        self._update_count()


# ══════════════════════════════════════════════════════
#  ВКЛАДКА СПИСАНИЯ
# ══════════════════════════════════════════════════════

SPISAN_COLS = [
    ('num',           '№',                   42),
    ('name',          'Наименование',        180),
    ('inv_number',    'Инв. номер',           90),
    ('year',          'Год',                  52),
    ('initial_cost',  'Перв. ст-сть',        100),
    ('residual_cost', 'Ост. ст-сть',         100),
    ('reason',        'Причина',             160),
    ('note',          'Примечание',          110),
]

UST_COLS = [
    ('num',           '№',              42),
    ('name',          'Наименование',  160),
    ('model',         'Модель',        110),
    ('serial_number', 'Серийный №',    110),
    ('inv_number',    'Инв. номер',     80),
    ('year',          'Год',            52),
    ('cost',          'Стоимость',     100),
    ('install_date',  'Дата уст-ки',    90),
    ('condition',     'Состояние',      90),
    ('note',          'Примечание',    100),
]


class SpisanTab(tk.Frame):
    """Вкладка для генерации актов СПИСАНИЯ."""

    ACCENT = COLORS['accent_red']

    def __init__(self, parent, status_bar):
        super().__init__(parent, bg=COLORS['bg'])
        self._status = status_bar
        self._build()

    def _build(self):
        bg = COLORS['bg']

        # ── Заголовок вкладки ──
        hdr = tk.Frame(self, bg=self.ACCENT, height=52)
        hdr.pack(fill='x')
        hdr.pack_propagate(False)
        tk.Label(hdr, text='  📋  АКТ СПИСАНИЯ ОСНОВНЫХ СРЕДСТВ',
                 bg=self.ACCENT, fg='white', font=FONT_TITLE).pack(
            side='left', padx=16, pady=10)

        # ── Скролл контейнер ──
        canvas = tk.Canvas(self, bg=bg, highlightthickness=0)
        vsb = ttk.Scrollbar(self, orient='vertical', command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side='right', fill='y')
        canvas.pack(fill='both', expand=True)

        self._scroll_frame = tk.Frame(canvas, bg=bg)
        self._win = canvas.create_window((0, 0), window=self._scroll_frame, anchor='nw')
        self._scroll_frame.bind('<Configure>',
                                lambda e: canvas.configure(
                                    scrollregion=canvas.bbox('all')))
        canvas.bind('<Configure>',
                    lambda e: canvas.itemconfig(self._win, width=e.width))
        canvas.bind_all('<MouseWheel>',
                        lambda e: canvas.yview_scroll(-1*(e.delta//120), 'units'))

        inner = self._scroll_frame

        # ── Форма организации ──
        self.org_form = OrgForm(inner, accent=self.ACCENT, bg=bg)
        self.org_form.pack(fill='x', padx=16, pady=(12, 0))

        # ── Таблица оборудования ──
        self.items_table = ItemsTable(inner, SPISAN_COLS,
                                      accent=self.ACCENT, bg=bg)
        self.items_table.pack(fill='both', expand=True, padx=16, pady=(8, 0))

        # ── Параметры сохранения ──
        self._build_save_options(inner, bg)

        # ── Лог ──
        self._build_log(inner, bg)

    def _build_save_options(self, parent, bg):
        card = SectionCard(parent, title='Параметры сохранения', accent=self.ACCENT, bg=bg)
        card.pack(fill='x', padx=16, pady=(8, 0))

        row = tk.Frame(card.inner, bg=bg)
        row.pack(fill='x', pady=4)

        # Режим
        tk.Label(row, text='Режим сохранения:', bg=bg, fg=COLORS['text'],
                 font=FONT_BODY).pack(side='left', padx=(0, 8))
        self._save_mode = tk.StringVar(value='single')
        tk.Radiobutton(row, text='Каждый акт — отдельный файл',
                       variable=self._save_mode, value='single',
                       bg=bg, fg=COLORS['text'], font=FONT_BODY,
                       selectcolor=bg, activebackground=bg).pack(side='left', padx=(0, 16))
        tk.Radiobutton(row, text='Все акты в одном файле',
                       variable=self._save_mode, value='all',
                       bg=bg, fg=COLORS['text'], font=FONT_BODY,
                       selectcolor=bg, activebackground=bg).pack(side='left')

        row2 = tk.Frame(card.inner, bg=bg)
        row2.pack(fill='x', pady=(6, 0))

        self._out_dir = LabeledEntry(row2, 'Папка для сохранения', width=50, bg=bg)
        self._out_dir.pack(side='left', padx=(0, 8), fill='x', expand=True)
        self._out_dir.set(os.path.join(os.path.expanduser('~'), 'Desktop'))

        RoundedButton(row2, text='Обзор...', command=self._browse_dir,
                      bg='#5D6D7E', width=90, height=30).pack(side='left', anchor='s', pady=2)

        row3 = tk.Frame(card.inner, bg=bg)
        row3.pack(fill='x', pady=(8, 4))

        RoundedButton(row3, text='📥  Загрузить из Excel',
                      command=self._load_excel,
                      bg='#1E8B4C', width=180, height=36).pack(side='left', padx=(0, 10))
        RoundedButton(row3, text='📄  Создать шаблон Excel',
                      command=self._create_template,
                      bg='#5D6D7E', width=190, height=36).pack(side='left', padx=(0, 10))
        RoundedButton(row3, text='🖨  Сгенерировать акт(ы)',
                      command=self._generate,
                      bg=self.ACCENT, width=195, height=36).pack(side='left')

    def _build_log(self, parent, bg):
        card = SectionCard(parent, title='Журнал событий', accent=self.ACCENT, bg=bg)
        card.pack(fill='x', padx=16, pady=(8, 16))

        self._log = scrolledtext.ScrolledText(
            card.inner, height=6, font=FONT_MONO,
            bg='#0D1B2A', fg='#A8D8EA', state='disabled',
            relief='flat', bd=0, wrap='word')
        self._log.pack(fill='x')

    def _log_msg(self, msg, color='#A8D8EA'):
        self._log.config(state='normal')
        ts = datetime.now().strftime('%H:%M:%S')
        self._log.insert('end', f'[{ts}]  {msg}\n')
        self._log.see('end')
        self._log.config(state='disabled')

    def _browse_dir(self):
        d = filedialog.askdirectory(title='Выберите папку для сохранения')
        if d:
            self._out_dir.set(d)

    def _create_template(self):
        path = filedialog.asksaveasfilename(
            title='Сохранить шаблон Excel',
            defaultextension='.xlsx',
            filetypes=[('Excel файлы', '*.xlsx')],
            initialfile='shablon_spisaniya.xlsx')
        if not path:
            return
        try:
            excel_handler.create_spisan_template(path)
            self._log_msg(f'✅ Шаблон Excel создан: {path}')
            self._status.set(f'Шаблон сохранён: {os.path.basename(path)}')
            messagebox.showinfo('Успех', f'Шаблон Excel создан:\n{path}')
        except Exception as e:
            self._log_msg(f'❌ Ошибка: {e}')
            messagebox.showerror('Ошибка', str(e))

    def _load_excel(self):
        path = filedialog.askopenfilename(
            title='Открыть Excel шаблон списания (shablon_spisan.xlsx)',
            filetypes=[('Excel файлы', '*.xlsx *.xls')])
        if not path:
            return
        try:
            # read_spisan_excel возвращает список групп по инвентарному номеру
            groups = excel_handler.read_spisan_excel(path)
            self._loaded_groups = groups   # сохраняем для генерации

            if groups:
                # Показываем данные первой группы в форме
                first = groups[0]
                self.org_form.set_data(first)
                # В таблице показываем все компоненты всех устройств
                flat_items = []
                for g in groups:
                    for dev in g.get('devices', []):
                        for p in dev.get('parts', []):
                            flat_items.append({
                                'num':           '',
                                'name':          dev['name'],
                                'inv_number':    g['inv_number'],
                                'year':          '',
                                'initial_cost':  '',
                                'residual_cost': '',
                                'reason':        p.get('defect', ''),
                                'note':          p.get('condition', ''),
                            })
                self.items_table.load_items(flat_items)

            total_devs = sum(len(g.get('devices', [])) for g in groups)
            self._log_msg(
                f'✅ Загружено из Excel: {os.path.basename(path)} — '
                f'{len(groups)} инв. номеров, {total_devs} устройств'
            )
            self._status.set(f'Загружено: {os.path.basename(path)} ({len(groups)} групп)')
        except Exception as e:
            self._log_msg(f'❌ Ошибка загрузки: {e}')
            messagebox.showerror('Ошибка загрузки', str(e))

    def _generate(self):
        out_dir = self._out_dir.get() or '.'
        mode    = self._save_mode.get()

        # Если данные загружены из Excel — используем их
        groups = getattr(self, '_loaded_groups', None)

        if not groups:
            # Если нет загруженных групп — берём из формы вручную
            org   = self.org_form.get_data()
            items = self.items_table.get_items()
            if not org['org_name']:
                messagebox.showwarning('Предупреждение', 'Укажите название организации')
                return
            if not items:
                messagebox.showwarning('Предупреждение', 'Добавьте хотя бы одну строку оборудования')
                return
            # Группируем по inv_number
            inv_map = {}
            inv_order = []
            for item in items:
                inv = item.get('inv_number', 'N/A') or 'N/A'
                if inv not in inv_map:
                    inv_map[inv] = {'name': item.get('name',''), 'parts': []}
                    inv_order.append(inv)
                inv_map[inv]['parts'].append({
                    'part_name': item.get('name', ''),
                    'condition': item.get('note', ''),
                    'defect':    item.get('reason', ''),
                })
            groups = []
            for inv in inv_order:
                groups.append({
                    **org,
                    'inv_number': inv,
                    'devices': [{'name': inv_map[inv]['name'],
                                 'parts': inv_map[inv]['parts']}],
                })

        def do_gen():
            try:
                self._status.set('Генерация актов списания...')
                if mode == 'single':
                    paths = generator_spisan.save_single_files(groups, out_dir)
                    for p in paths:
                        self._log_msg(f'✅ Создан: {os.path.basename(p)}')
                    msg = f'Создано файлов: {len(paths)}\nПапка: {out_dir}'
                else:
                    fname    = f'Texnik_Xulosa_vse_{datetime.now().strftime("%Y%m%d_%H%M%S")}.docx'
                    out_path = os.path.join(out_dir, fname)
                    generator_spisan.save_all_in_one(groups, out_path)
                    self._log_msg(f'✅ Создан общий файл: {fname}')
                    msg = f'Создан файл:\n{out_path}'

                self._status.set(f'Готово ✓  ({len(groups)} актов)')
                messagebox.showinfo('Готово', msg)
            except Exception as e:
                self._log_msg(f'❌ Ошибка генерации: {e}')
                self._status.set('Ошибка!')
                messagebox.showerror('Ошибка', str(e))

        threading.Thread(target=do_gen, daemon=True).start()


# ══════════════════════════════════════════════════════
#  ВКЛАДКА УСТАНОВКИ
# ══════════════════════════════════════════════════════

class UstTab(tk.Frame):
    """Вкладка для генерации актов УСТАНОВКИ."""

    ACCENT = COLORS['accent_blue']

    def __init__(self, parent, status_bar):
        super().__init__(parent, bg=COLORS['bg'])
        self._status = status_bar
        self._build()

    def _build(self):
        bg = COLORS['bg']

        hdr = tk.Frame(self, bg=self.ACCENT, height=52)
        hdr.pack(fill='x')
        hdr.pack_propagate(False)
        tk.Label(hdr, text='  🔧  АКТ ВВОДА В ЭКСПЛУАТАЦИЮ (УСТАНОВКИ)',
                 bg=self.ACCENT, fg='white', font=FONT_TITLE).pack(
            side='left', padx=16, pady=10)

        canvas = tk.Canvas(self, bg=bg, highlightthickness=0)
        vsb = ttk.Scrollbar(self, orient='vertical', command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side='right', fill='y')
        canvas.pack(fill='both', expand=True)

        self._scroll_frame = tk.Frame(canvas, bg=bg)
        self._win = canvas.create_window((0, 0), window=self._scroll_frame, anchor='nw')
        self._scroll_frame.bind('<Configure>',
                                lambda e: canvas.configure(
                                    scrollregion=canvas.bbox('all')))
        canvas.bind('<Configure>',
                    lambda e: canvas.itemconfig(self._win, width=e.width))
        canvas.bind_all('<MouseWheel>',
                        lambda e: canvas.yview_scroll(-1*(e.delta//120), 'units'))

        inner = self._scroll_frame

        # Форма организации + поле "Место установки"
        self.org_form = OrgForm(inner, accent=self.ACCENT, bg=bg)
        self.org_form.pack(fill='x', padx=16, pady=(12, 0))

        # Дополнительное поле "Место установки"
        loc_card = SectionCard(inner, title='Место установки', accent=self.ACCENT, bg=bg)
        loc_card.pack(fill='x', padx=16, pady=(8, 0))
        self.location = LabeledEntry(loc_card.inner, 'Адрес / кабинет / место установки',
                                     width=60, bg=bg)
        self.location.pack(fill='x')

        self.items_table = ItemsTable(inner, UST_COLS, accent=self.ACCENT, bg=bg)
        self.items_table.pack(fill='both', expand=True, padx=16, pady=(8, 0))

        self._build_save_options(inner, bg)
        self._build_log(inner, bg)

    def _build_save_options(self, parent, bg):
        card = SectionCard(parent, title='Параметры сохранения', accent=self.ACCENT, bg=bg)
        card.pack(fill='x', padx=16, pady=(8, 0))

        row = tk.Frame(card.inner, bg=bg)
        row.pack(fill='x', pady=4)
        tk.Label(row, text='Режим сохранения:', bg=bg, fg=COLORS['text'],
                 font=FONT_BODY).pack(side='left', padx=(0, 8))
        self._save_mode = tk.StringVar(value='single')
        tk.Radiobutton(row, text='Каждый акт — отдельный файл',
                       variable=self._save_mode, value='single',
                       bg=bg, fg=COLORS['text'], font=FONT_BODY,
                       selectcolor=bg, activebackground=bg).pack(side='left', padx=(0, 16))
        tk.Radiobutton(row, text='Все акты в одном файле',
                       variable=self._save_mode, value='all',
                       bg=bg, fg=COLORS['text'], font=FONT_BODY,
                       selectcolor=bg, activebackground=bg).pack(side='left')

        row2 = tk.Frame(card.inner, bg=bg)
        row2.pack(fill='x', pady=(6, 0))
        self._out_dir = LabeledEntry(row2, 'Папка для сохранения', width=50, bg=bg)
        self._out_dir.pack(side='left', padx=(0, 8), fill='x', expand=True)
        self._out_dir.set(os.path.join(os.path.expanduser('~'), 'Desktop'))
        RoundedButton(row2, text='Обзор...', command=self._browse_dir,
                      bg='#5D6D7E', width=90, height=30).pack(side='left', anchor='s', pady=2)

        row3 = tk.Frame(card.inner, bg=bg)
        row3.pack(fill='x', pady=(8, 4))
        RoundedButton(row3, text='📥  Загрузить из Excel',
                      command=self._load_excel,
                      bg='#1E8B4C', width=180, height=36).pack(side='left', padx=(0, 10))
        RoundedButton(row3, text='📄  Создать шаблон Excel',
                      command=self._create_template,
                      bg='#5D6D7E', width=190, height=36).pack(side='left', padx=(0, 10))
        RoundedButton(row3, text='🖨  Сгенерировать акт(ы)',
                      command=self._generate,
                      bg=self.ACCENT, width=195, height=36).pack(side='left')

    def _build_log(self, parent, bg):
        card = SectionCard(parent, title='Журнал событий', accent=self.ACCENT, bg=bg)
        card.pack(fill='x', padx=16, pady=(8, 16))
        self._log = scrolledtext.ScrolledText(
            card.inner, height=6, font=FONT_MONO,
            bg='#0D1B2A', fg='#A8D8EA', state='disabled',
            relief='flat', bd=0, wrap='word')
        self._log.pack(fill='x')

    def _log_msg(self, msg):
        self._log.config(state='normal')
        ts = datetime.now().strftime('%H:%M:%S')
        self._log.insert('end', f'[{ts}]  {msg}\n')
        self._log.see('end')
        self._log.config(state='disabled')

    def _browse_dir(self):
        d = filedialog.askdirectory(title='Выберите папку для сохранения')
        if d:
            self._out_dir.set(d)

    def _create_template(self):
        path = filedialog.asksaveasfilename(
            title='Сохранить шаблон Excel',
            defaultextension='.xlsx',
            filetypes=[('Excel файлы', '*.xlsx')],
            initialfile='shablon_ustanovki.xlsx')
        if not path:
            return
        try:
            excel_handler.create_ust_template(path)
            self._log_msg(f'✅ Шаблон Excel создан: {path}')
            self._status.set(f'Шаблон сохранён: {os.path.basename(path)}')
            messagebox.showinfo('Успех', f'Шаблон Excel создан:\n{path}')
        except Exception as e:
            self._log_msg(f'❌ Ошибка: {e}')
            messagebox.showerror('Ошибка', str(e))

    def _load_excel(self):
        path = filedialog.askopenfilename(
            title='Открыть Excel шаблон установки (shablom_ust.xlsx)',
            filetypes=[('Excel файлы', '*.xlsx *.xls')])
        if not path:
            return
        try:
            data = excel_handler.read_ust_excel(path)
            self._loaded_data = data   # сохраняем для генерации
            self.org_form.set_data(data)
            self.location.set(data.get('location', data.get('address', '')))
            self.items_table.load_items(data.get('items', []))
            self._log_msg(
                f'✅ Загружено из Excel: {os.path.basename(path)} '
                f'({len(data.get("items", []))} строк)'
            )
            self._status.set(f'Загружено: {os.path.basename(path)}')
        except Exception as e:
            self._log_msg(f'❌ Ошибка загрузки: {e}')
            messagebox.showerror('Ошибка загрузки', str(e))

    def _generate(self):
        out_dir = self._out_dir.get() or '.'
        mode    = self._save_mode.get()

        # Предпочитаем данные загруженные из Excel
        loaded = getattr(self, '_loaded_data', None)
        if loaded:
            data = loaded
        else:
            org   = self.org_form.get_data()
            items = self.items_table.get_items()
            if not org['org_name']:
                messagebox.showwarning('Предупреждение', 'Укажите название организации')
                return
            if not items:
                messagebox.showwarning('Предупреждение', 'Добавьте хотя бы одну строку оборудования')
                return
            data = {
                **org,
                'address':       self.location.get(),
                'location':      self.location.get(),
                'member1_title': 'Yetakchi muhandis',
                'member2_title': 'Yetakchi muhandis',
                'items':         items,
            }

        def do_gen():
            try:
                self._status.set('Генерация акта установки...')
                if mode == 'single':
                    paths = generator_ust.save_single_files([data], out_dir)
                    for p in paths:
                        self._log_msg(f'✅ Создан: {os.path.basename(p)}')
                    msg = f'Создано файлов: {len(paths)}\nПапка: {out_dir}'
                else:
                    fname    = f'Dalolatnoma_vse_{datetime.now().strftime("%Y%m%d_%H%M%S")}.docx'
                    out_path = os.path.join(out_dir, fname)
                    generator_ust.save_all_in_one([data], out_path)
                    self._log_msg(f'✅ Создан общий файл: {fname}')
                    msg = f'Создан файл:\n{out_path}'

                self._status.set(f'Готово ✓')
                messagebox.showinfo('Готово', msg)
            except Exception as e:
                self._log_msg(f'❌ Ошибка генерации: {e}')
                self._status.set('Ошибка!')
                messagebox.showerror('Ошибка', str(e))

        threading.Thread(target=do_gen, daemon=True).start()


# ══════════════════════════════════════════════════════
#  ВКЛАДКА — ПАКЕТНАЯ ГЕНЕРАЦИЯ (несколько актов сразу)
# ══════════════════════════════════════════════════════

class BatchTab(tk.Frame):
    """Вкладка пакетной генерации из нескольких Excel файлов."""

    def __init__(self, parent, status_bar):
        super().__init__(parent, bg=COLORS['bg'])
        self._status = status_bar
        self._files = []
        self._build()

    def _build(self):
        bg = COLORS['bg']
        ACCENT = '#5D3FD3'

        hdr = tk.Frame(self, bg=ACCENT, height=52)
        hdr.pack(fill='x')
        hdr.pack_propagate(False)
        tk.Label(hdr, text='  📦  ПАКЕТНАЯ ГЕНЕРАЦИЯ',
                 bg=ACCENT, fg='white', font=FONT_TITLE).pack(
            side='left', padx=16, pady=10)

        inner = tk.Frame(self, bg=bg)
        inner.pack(fill='both', expand=True, padx=16, pady=16)

        # ── Описание ──
        desc = tk.Label(inner,
                        text='Добавьте несколько Excel файлов (шаблонов списания или установки) '
                             'и сгенерируйте все акты за один раз.',
                        bg=bg, fg=COLORS['text_muted'], font=FONT_BODY,
                        wraplength=700, justify='left')
        desc.pack(anchor='w', pady=(0, 12))

        # ── Тип акта ──
        type_frame = tk.Frame(inner, bg=bg)
        type_frame.pack(fill='x', pady=(0, 10))
        tk.Label(type_frame, text='Тип акта:', bg=bg, fg=COLORS['text'],
                 font=FONT_HEADER).pack(side='left', padx=(0, 10))
        self._doc_type = tk.StringVar(value='spisan')
        tk.Radiobutton(type_frame, text='Списание', variable=self._doc_type,
                       value='spisan', bg=bg, fg=COLORS['text'], font=FONT_BODY,
                       selectcolor=bg, activebackground=bg).pack(side='left', padx=(0, 16))
        tk.Radiobutton(type_frame, text='Установка', variable=self._doc_type,
                       value='ust', bg=bg, fg=COLORS['text'], font=FONT_BODY,
                       selectcolor=bg, activebackground=bg).pack(side='left')

        # ── Список файлов ──
        list_card = SectionCard(inner, title='Файлы Excel', accent=ACCENT, bg=bg)
        list_card.pack(fill='both', expand=True, pady=(0, 10))

        toolbar = tk.Frame(list_card.inner, bg=bg)
        toolbar.pack(fill='x', pady=(0, 6))
        RoundedButton(toolbar, text='+ Добавить файлы', command=self._add_files,
                      bg=ACCENT, width=150, height=30).pack(side='left', padx=(0, 8))
        RoundedButton(toolbar, text='✕ Удалить выбранный', command=self._remove_file,
                      bg='#A93226', width=160, height=30).pack(side='left', padx=(0, 8))
        RoundedButton(toolbar, text='Очистить список', command=self._clear_files,
                      bg='#5D6D7E', width=140, height=30).pack(side='left')

        self.file_listbox = tk.Listbox(
            list_card.inner, height=8, font=FONT_BODY,
            bg=COLORS['card_bg'], fg=COLORS['text'],
            selectbackground=ACCENT, relief='flat',
            highlightthickness=1, highlightbackground=COLORS['border'])
        self.file_listbox.pack(fill='both', expand=True)

        # ── Параметры сохранения ──
        save_card = SectionCard(inner, title='Параметры', accent=ACCENT, bg=bg)
        save_card.pack(fill='x', pady=(0, 10))

        row = tk.Frame(save_card.inner, bg=bg)
        row.pack(fill='x', pady=4)
        tk.Label(row, text='Режим:', bg=bg, fg=COLORS['text'],
                 font=FONT_BODY).pack(side='left', padx=(0, 8))
        self._save_mode = tk.StringVar(value='single')
        tk.Radiobutton(row, text='Каждый акт — отдельный файл',
                       variable=self._save_mode, value='single',
                       bg=bg, fg=COLORS['text'], font=FONT_BODY,
                       selectcolor=bg, activebackground=bg).pack(side='left', padx=(0, 16))
        tk.Radiobutton(row, text='Все акты в одном файле',
                       variable=self._save_mode, value='all',
                       bg=bg, fg=COLORS['text'], font=FONT_BODY,
                       selectcolor=bg, activebackground=bg).pack(side='left')

        row2 = tk.Frame(save_card.inner, bg=bg)
        row2.pack(fill='x', pady=(6, 4))
        self._out_dir = LabeledEntry(row2, 'Папка для сохранения', width=50, bg=bg)
        self._out_dir.pack(side='left', padx=(0, 8), fill='x', expand=True)
        self._out_dir.set(os.path.join(os.path.expanduser('~'), 'Desktop'))
        RoundedButton(row2, text='Обзор...', command=self._browse,
                      bg='#5D6D7E', width=90, height=30).pack(side='left', anchor='s', pady=2)

        RoundedButton(save_card.inner, text='🖨  Запустить пакетную генерацию',
                      command=self._generate, bg=ACCENT,
                      width=260, height=36).pack(pady=(8, 4))

        # ── Лог ──
        log_card = SectionCard(inner, title='Журнал', accent=ACCENT, bg=bg)
        log_card.pack(fill='x')
        self._log = scrolledtext.ScrolledText(
            log_card.inner, height=5, font=FONT_MONO,
            bg='#0D1B2A', fg='#A8D8EA', state='disabled',
            relief='flat', bd=0, wrap='word')
        self._log.pack(fill='x')

    def _log_msg(self, msg):
        self._log.config(state='normal')
        ts = datetime.now().strftime('%H:%M:%S')
        self._log.insert('end', f'[{ts}]  {msg}\n')
        self._log.see('end')
        self._log.config(state='disabled')

    def _add_files(self):
        paths = filedialog.askopenfilenames(
            title='Выберите Excel файлы',
            filetypes=[('Excel файлы', '*.xlsx *.xls')])
        for p in paths:
            if p not in self._files:
                self._files.append(p)
                self.file_listbox.insert('end', os.path.basename(p))

    def _remove_file(self):
        sel = self.file_listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        self.file_listbox.delete(idx)
        self._files.pop(idx)

    def _clear_files(self):
        self.file_listbox.delete(0, 'end')
        self._files.clear()

    def _browse(self):
        d = filedialog.askdirectory()
        if d:
            self._out_dir.set(d)

    def _generate(self):
        if not self._files:
            messagebox.showwarning('Предупреждение', 'Добавьте хотя бы один Excel файл')
            return
        doc_type = self._doc_type.get()
        out_dir = self._out_dir.get() or '.'
        mode = self._save_mode.get()

        def do_gen():
            data_list = []
            for path in self._files:
                try:
                    if doc_type == 'spisan':
                        groups = excel_handler.read_spisan_excel(path)
                        data_list.extend(groups)
                        self._log_msg(f'📂 {os.path.basename(path)}: {len(groups)} инв. номеров')
                    else:
                        data = excel_handler.read_ust_excel(path)
                        data_list.append(data)
                        self._log_msg(f'📂 {os.path.basename(path)}: {len(data.get("items", []))} строк')
                except Exception as e:
                    self._log_msg(f'❌ Ошибка чтения {os.path.basename(path)}: {e}')

            if not data_list:
                self._status.set('Нет данных для генерации')
                return

            try:
                self._status.set('Генерация...')
                if mode == 'single':
                    if doc_type == 'spisan':
                        paths = generator_spisan.save_single_files(data_list, out_dir)
                    else:
                        paths = generator_ust.save_single_files(data_list, out_dir)
                    for p in paths:
                        self._log_msg(f'✅ Создан: {os.path.basename(p)}')
                else:
                    ts    = datetime.now().strftime('%Y%m%d_%H%M%S')
                    label = 'Xulosa' if doc_type == 'spisan' else 'Dalolatnoma'
                    fname = f'{label}_vse_{ts}.docx'
                    out_path = os.path.join(out_dir, fname)
                    if doc_type == 'spisan':
                        generator_spisan.save_all_in_one(data_list, out_path)
                    else:
                        generator_ust.save_all_in_one(data_list, out_path)
                    self._log_msg(f'✅ Создан общий файл: {fname}')

                self._status.set(f'Готово ✓  Создано актов: {len(data_list)}')
                messagebox.showinfo('Готово',
                                    f'Генерация завершена!\nСоздано актов: {len(data_list)}\n'
                                    f'Папка: {out_dir}')
            except Exception as e:
                self._log_msg(f'❌ Ошибка: {e}')
                self._status.set('Ошибка!')
                messagebox.showerror('Ошибка', str(e))

        threading.Thread(target=do_gen, daemon=True).start()


# ══════════════════════════════════════════════════════
#  ГЛАВНОЕ ОКНО
# ══════════════════════════════════════════════════════

class App(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title('Генератор актов | Списание и Установка оборудования')
        self.configure(bg=COLORS['header_bg'])
        self._center_window(WINDOW_W, WINDOW_H)
        self.minsize(900, 600)

        self._build_header()
        self._build_body()
        self._build_statusbar()

    def _center_window(self, w, h):
        self.update_idletasks()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        x = (sw - w) // 2
        y = (sh - h) // 2
        self.geometry(f'{w}x{h}+{x}+{y}')

    def _build_header(self):
        header = tk.Frame(self, bg=COLORS['header_bg'], height=60)
        header.pack(fill='x')
        header.pack_propagate(False)

        tk.Label(header,
                 text='⚙  ГЕНЕРАТОР АКТОВ ОБОРУДОВАНИЯ',
                 bg=COLORS['header_bg'], fg='white',
                 font=('Segoe UI', 15, 'bold')).pack(side='left', padx=20, pady=12)

        tk.Label(header,
                 text='ABM MChJ  |  Sirdaryo viloyati',
                 bg=COLORS['header_bg'], fg='#A8C7E8',
                 font=FONT_SMALL).pack(side='right', padx=20)

        # Разделитель
        tk.Frame(self, bg=COLORS['tab_active'], height=3).pack(fill='x')

    def _build_body(self):
        # Notebook (вкладки)
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('Custom.TNotebook',
                         background=COLORS['bg'],
                         tabmargins=[2, 4, 2, 0])
        style.configure('Custom.TNotebook.Tab',
                         background=COLORS['sidebar'],
                         foreground='#A8C7E8',
                         font=FONT_HEADER,
                         padding=[16, 8],
                         borderwidth=0)
        style.map('Custom.TNotebook.Tab',
                  background=[('selected', COLORS['bg']),
                               ('active', COLORS['sidebar_hover'])],
                  foreground=[('selected', COLORS['text']),
                               ('active', 'white')])

        self.notebook = ttk.Notebook(self, style='Custom.TNotebook')
        self.notebook.pack(fill='both', expand=True)

        # Создаём статус-бар заранее
        self.status_bar = StatusBar(self)

        # Вкладки
        self.tab_spisan = SpisanTab(self.notebook, self.status_bar)
        self.tab_ust    = UstTab(self.notebook, self.status_bar)
        self.tab_batch  = BatchTab(self.notebook, self.status_bar)

        self.notebook.add(self.tab_spisan, text='  📋  СПИСАНИЕ  ')
        self.notebook.add(self.tab_ust,    text='  🔧  УСТАНОВКА  ')
        self.notebook.add(self.tab_batch,  text='  📦  ПАКЕТНАЯ ГЕНЕРАЦИЯ  ')

    def _build_statusbar(self):
        self.status_bar.pack(fill='x', side='bottom')


# ══════════════════════════════════════════════════════
#  ТОЧКА ВХОДА
# ══════════════════════════════════════════════════════

if __name__ == '__main__':
    app = App()
    app.mainloop()
