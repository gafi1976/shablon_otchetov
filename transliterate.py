# transliterate.py
# Конвертация узбекского текста: латиница ↔ кириллица
# Поддерживает оба направления + автоопределение алфавита

# ══════════════════════════════════════════════════════
#  ТАБЛИЦЫ КОНВЕРТАЦИИ
# ══════════════════════════════════════════════════════

# Латиница → Кириллица (узбекский, новая орфография)
# Порядок важен: сначала диграфы (sh, ch, ng, o', g', ...), потом одиночные
_LAT_TO_CYR = [
    # Диграфы и апостроф — сначала!
    ("O'",  'О'),  ("o'",  'о'),
    ("G'",  'Ғ'),  ("g'",  'ғ'),
    ('Sh',  'Ш'),  ('sh',  'ш'),  ('SH', 'Ш'),
    ('Ch',  'Ч'),  ('ch',  'ч'),  ('CH', 'Ч'),
    ('Ng',  'Нг'), ('ng',  'нг'), ('NG', 'НГ'),
    ('Yo',  'Ё'),  ('yo',  'ё'),  ('YO', 'Ё'),
    ('Yu',  'Ю'),  ('yu',  'ю'),  ('YU', 'Ю'),
    ('Ya',  'Я'),  ('ya',  'я'),  ('YA', 'Я'),
    ('Ts',  'Ц'),  ('ts',  'ц'),  ('TS', 'Ц'),
    # Одиночные
    ('A',   'А'),  ('a',   'а'),
    ('B',   'Б'),  ('b',   'б'),
    ('D',   'Д'),  ('d',   'д'),
    ('E',   'Е'),  ('e',   'е'),
    ('F',   'Ф'),  ('f',   'ф'),
    ('G',   'Г'),  ('g',   'г'),
    ('H',   'Ҳ'),  ('h',   'ҳ'),
    ('I',   'И'),  ('i',   'и'),
    ('J',   'Ж'),  ('j',   'ж'),
    ('K',   'К'),  ('k',   'к'),
    ('L',   'Л'),  ('l',   'л'),
    ('M',   'М'),  ('m',   'м'),
    ('N',   'Н'),  ('n',   'н'),
    ('O',   'О'),  ('o',   'о'),
    ('P',   'П'),  ('p',   'п'),
    ('Q',   'Қ'),  ('q',   'қ'),
    ('R',   'Р'),  ('r',   'р'),
    ('S',   'С'),  ('s',   'с'),
    ('T',   'Т'),  ('t',   'т'),
    ('U',   'У'),  ('u',   'у'),
    ('V',   'В'),  ('v',   'в'),
    ('W',   'В'),  ('w',   'в'),
    ('X',   'Х'),  ('x',   'х'),
    ('Y',   'Й'),  ('y',   'й'),
    ('Z',   'З'),  ('z',   'з'),
    ("\u2019", "'"),  # правая кавычка → апостроф
]

# Кириллица → Латиница
_CYR_TO_LAT = [
    ('Ё',  'Yo'),  ('ё',  'yo'),
    ('Ю',  'Yu'),  ('ю',  'yu'),
    ('Я',  'Ya'),  ('я',  'ya'),
    ('Ш',  'Sh'),  ('ш',  'sh'),
    ('Ч',  'Ch'),  ('ч',  'ch'),
    ('Ц',  'Ts'),  ('ц',  'ts'),
    ('НГ', 'Ng'),  ('Нг', 'Ng'),  ('нг', 'ng'),
    ('Ғ',  "G'"), ('ғ',  "g'"),
    ('Қ',  'Q'),   ('қ',  'q'),
    ('Ҳ',  'H'),   ('ҳ',  'h'),
    ('Ў',  "O'"), ('ў',  "o'"),
    ('А',  'A'),   ('а',  'a'),
    ('Б',  'B'),   ('б',  'b'),
    ('В',  'V'),   ('в',  'v'),
    ('Г',  'G'),   ('г',  'g'),
    ('Д',  'D'),   ('д',  'd'),
    ('Е',  'E'),   ('е',  'e'),
    ('Ж',  'J'),   ('ж',  'j'),
    ('З',  'Z'),   ('з',  'z'),
    ('И',  'I'),   ('и',  'i'),
    ('Й',  'Y'),   ('й',  'y'),
    ('К',  'K'),   ('к',  'k'),
    ('Л',  'L'),   ('л',  'l'),
    ('М',  'M'),   ('м',  'm'),
    ('Н',  'N'),   ('н',  'n'),
    ('О',  'O'),   ('о',  'o'),
    ('П',  'P'),   ('п',  'p'),
    ('Р',  'R'),   ('р',  'r'),
    ('С',  'S'),   ('с',  's'),
    ('Т',  'T'),   ('т',  't'),
    ('У',  'U'),   ('у',  'u'),
    ('Ф',  'F'),   ('ф',  'f'),
    ('Х',  'X'),   ('х',  'x'),
    ('Ъ',  "'"),   ('ъ',  "'"),
    ('Ь',  ''),    ('ь',  ''),
    ('Э',  'E'),   ('э',  'e'),
    ('Ю',  'Yu'),  ('ю',  'yu'),
    ('Я',  'Ya'),  ('я',  'ya'),
]


# ══════════════════════════════════════════════════════
#  ФУНКЦИИ КОНВЕРТАЦИИ
# ══════════════════════════════════════════════════════

def lat_to_cyr(text: str) -> str:
    """Конвертирует узбекский текст из латиницы в кириллицу."""
    if not text:
        return text
    result = text
    for lat, cyr in _LAT_TO_CYR:
        result = result.replace(lat, cyr)
    return result


def cyr_to_lat(text: str) -> str:
    """Конвертирует узбекский текст из кириллицы в латиницу."""
    if not text:
        return text
    result = text
    for cyr, lat in _CYR_TO_LAT:
        result = result.replace(cyr, lat)
    return result


def detect_script(text: str) -> str:
    """
    Автоматически определяет алфавит текста.
    Возвращает: 'latin', 'cyrillic' или 'unknown'
    """
    if not text:
        return 'unknown'

    cyr_chars = sum(1 for c in text if '\u0400' <= c <= '\u04FF')
    lat_chars  = sum(1 for c in text if c.isalpha() and ord(c) < 128)

    if cyr_chars == 0 and lat_chars == 0:
        return 'unknown'
    if cyr_chars > lat_chars:
        return 'cyrillic'
    return 'latin'


def convert_to(text: str, target: str) -> str:
    """
    Конвертирует текст в нужный алфавит.
    target: 'latin' или 'cyrillic'
    Если текст уже в нужном алфавите — возвращает без изменений.
    """
    if not text or not text.strip():
        return text

    script = detect_script(text)

    if target == 'cyrillic':
        if script == 'latin':
            return lat_to_cyr(text)
        return text   # уже кириллица или смешанный

    elif target == 'latin':
        if script == 'cyrillic':
            return cyr_to_lat(text)
        return text   # уже латиница

    return text


def convert_dict(data: dict, target: str, keys: list = None) -> dict:
    """
    Конвертирует строковые значения словаря в нужный алфавит.
    keys — список ключей для конвертации (если None — все строковые поля).
    """
    result = dict(data)
    for k, v in result.items():
        if keys and k not in keys:
            continue
        if isinstance(v, str):
            result[k] = convert_to(v, target)
        elif isinstance(v, list):
            result[k] = [
                convert_dict(item, target, keys) if isinstance(item, dict)
                else convert_to(item, target) if isinstance(item, str)
                else item
                for item in v
            ]
    return result


# ══════════════════════════════════════════════════════
#  ФИКСИРОВАННЫЕ ТЕКСТЫ ДОКУМЕНТОВ
#  (системные фразы — не из Excel, а встроенные)
# ══════════════════════════════════════════════════════

# Тексты для акта СПИСАНИЯ
TEXTS_SPISAN = {
    'latin': {
        'tasdiq':        "«TASDIQLAYMAN»",
        'yil':           "yil",
        'title':         "TEXNIK XULOSA №",
        'intro':         "Biz, quyida imzo chekuvchilar — {org} {region} bo'lim "
                         "yetakchi muhandisi {eng1} va yetakchi muhandis {eng2} lar, "
                         "quyida keltirilgan qurilmalarni texnik ko'rikdan o'tkazdik:",
        'qurilma':       "Qurilma nomi:",
        'inventar':      "Inventar raqami:",
        'col_qism':      "Qismlar nomi",
        'col_yarоq':     "Foydalanishga\nyaroqliligi",
        'col_nosoz':     "Nosozlik belgilari",
        'xulosa':        "XULOSA",
        'xulosa1':       "{names} zamonaviy talablarga javob bermaydi, ma'naviy eskirgan. "
                         "Ta'mirlash uchun zarur bo'lgan ehtiyot qismlar texnologik jarayondan "
                         "olib tashlanganligini hisobga olib, uni tiklash maqsadga muvofiq emas.",
        'xulosa2':       "Yuqorida keltirib o'tilgan jadvaldagi qurilmalarni xolatlaridan "
                         "kelib chiqib, komissiya asosiy vositalar ro'yxatidan chiqarilsin "
                         "degan xulosaga keldi.",
        'muhandis':      "Yetakchi muhandis:",
        'mp':            "M.P.",
        'sana':          "Sana:",
        'inv_label':     "Inventar raqami:",
    },
    'cyrillic': {
        'tasdiq':        "«ТАСДИҚЛАНДИ»",
        'yil':           "йил",
        'title':         "ТЕХНИК ХУЛОСА №",
        'intro':         "{org} {region} бўлим етакчи муҳандиси {eng1} ва "
                         "етакчи муҳандис {eng2} лар, қуйида келтирилган "
                         "қурилмаларни техник кўрикдан ўтказдик:",
        'qurilma':       "Қурилма номи:",
        'inventar':      "Инвентар рақами:",
        'col_qism':      "Қисмлар номи",
        'col_yarоq':     "Фойдаланишга\nяроқлилиги",
        'col_nosoz':     "Носозлик белгилари",
        'xulosa':        "ХУЛОСА",
        'xulosa1':       "{names} замонавий талабларга жавоб бермайди, маънавий эскирган. "
                         "Таъмирлаш учун зарур бўлган эҳтиёт қисмлар технологик жараёндан "
                         "олиб ташланганлигини ҳисобга олиб, уни тиклаш мақсадга мувофиқ эмас.",
        'xulosa2':       "Юқорида келтириб ўтилган жадвалдаги қурилмаларни ҳолатларидан "
                         "келиб чиқиб, комиссия асосий воситалар рўйхатидан чиқарилсин "
                         "деган хулосага келди.",
        'muhandis':      "Етакчи муҳандис:",
        'mp':            "М.П.",
        'sana':          "Сана:",
        'inv_label':     "Инвентар рақами:",
    },
}

# Тексты для акта УСТАНОВКИ
TEXTS_UST = {
    'latin': {
        'tasdiq':        "«TASDIQLAYMAN»",
        'rahbar':        "rahbari",
        'yil':           "yil",
        'manzil':        "Manzil:",
        'title':         "USKUNANI O\u2019RNATISH DALOLATNOMASI",
        'dalolatnoma':   "Dalolatnoma №",
        'intro':         "Mazkur dalolatnoma texnologik jarayonlarni modernizatsiya qilish "
                         "doirasida quyidagi uskunalar o\u2019rnatilganligini tasdiqlash "
                         "maqsadida tuzildi.",
        'col_num':       "№",
        'col_sana':      "O\u2019rnatilgan\nsana",
        'col_uskuna':    "Uskuna\n(Model)",
        'col_serial':    "Seriya\nraqami",
        'col_qayer':     "Qayerga\no\u2019rnatilgan",
        'outro':         "Yuqorida ko\u2019rsatilgan uskunalar belgilangan tartibda "
                         "o\u2019rnatildi, tekshirildi va ekspluatatsiyaga topshirildi.",
        'masul':         "O\u2019rnatish uchun mas\u2019ul shaxs",
        'mp':            "M.P.",
        'sana':          "Sana:",
    },
    'cyrillic': {
        'tasdiq':        "«ТАСДИҚЛАНДИ»",
        'rahbar':        "раҳбари",
        'yil':           "йил",
        'manzil':        "Манзил:",
        'title':         "УСКУНАНИ ЎРНАТИШ ДАЛОЛАТНОМАСИ",
        'dalolatnoma':   "Далолатнома №",
        'intro':         "Мазкур далолатнома технологик жараёнларни модернизация қилиш "
                         "доирасида қуйидаги ускуналар ўрнатилганлигини тасдиқлаш "
                         "мақсадида тузилди.",
        'col_num':       "№",
        'col_sana':      "Ўрнатилган\nsана",
        'col_uskuna':    "Ускуна\n(Модел)",
        'col_serial':    "Серия\nрақами",
        'col_qayer':     "Қаерга\nўрнатилган",
        'outro':         "Юқорида кўрсатилган ускуналар белгиланган тартибда "
                         "ўрнатилди, текширилди ва эксплуатацияга топширилди.",
        'masul':         "Ўрнатиш учун масъул шахс",
        'mp':            "М.П.",
        'sana':          "Сана:",
    },
}

# Названия месяцев
MONTHS = {
    'latin': {
        1:'yanvar', 2:'fevral', 3:'mart', 4:'aprel',
        5:'may', 6:'iyun', 7:'iyul', 8:'avgust',
        9:'sentabr', 10:'oktabr', 11:'noyabr', 12:'dekabr'
    },
    'cyrillic': {
        1:'январ', 2:'феврал', 3:'март', 4:'апрел',
        5:'май', 6:'июн', 7:'июл', 8:'август',
        9:'сентябр', 10:'октябр', 11:'ноябр', 12:'декабр'
    },
}
