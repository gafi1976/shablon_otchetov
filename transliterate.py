# -*- coding: utf-8 -*-
# transliterate.py
# Правильная транслитерация узбекского текста латиница ↔ кириллица
# Алгоритм взят из lat_cur.py (оригинальный файл пользователя)

import re

# ══════════════════════════════════════════════════════
#  ТАБЛИЦЫ КОНВЕРТАЦИИ
# ══════════════════════════════════════════════════════

LATIN_TO_CYRILLIC = {
    'a': 'а', 'A': 'А',
    'b': 'б', 'B': 'Б',
    'd': 'д', 'D': 'Д',
    'e': 'е', 'E': 'Е',
    'f': 'ф', 'F': 'Ф',
    'g': 'г', 'G': 'Г',
    'h': 'ҳ', 'H': 'Ҳ',
    'i': 'и', 'I': 'И',
    'j': 'ж', 'J': 'Ж',
    'k': 'к', 'K': 'К',
    'l': 'л', 'L': 'Л',
    'm': 'м', 'M': 'М',
    'n': 'н', 'N': 'Н',
    'o': 'о', 'O': 'О',
    'p': 'п', 'P': 'П',
    'q': 'қ', 'Q': 'Қ',
    'r': 'р', 'R': 'Р',
    's': 'с', 'S': 'С',
    't': 'т', 'T': 'Т',
    'u': 'у', 'U': 'У',
    'v': 'в', 'V': 'В',
    'x': 'х', 'X': 'Х',
    'y': 'й', 'Y': 'Й',
    'z': 'з', 'Z': 'З',
    '\u2018': 'ъ',
}

LATIN_VOWELS = (
    'a', 'A', 'e', 'E', 'i', 'I', 'o', 'O', 'u', 'U', 'o\u2018', 'O\u2018'
)

CYRILLIC_TO_LATIN = {
    'а': 'a',  'А': 'A',
    'б': 'b',  'Б': 'B',
    'в': 'v',  'В': 'V',
    'г': 'g',  'Г': 'G',
    'д': 'd',  'Д': 'D',
    'е': 'e',  'Е': 'E',
    'ё': 'yo', 'Ё': 'Yo',
    'ж': 'j',  'Ж': 'J',
    'з': 'z',  'З': 'Z',
    'и': 'i',  'И': 'I',
    'й': 'y',  'Й': 'Y',
    'к': 'k',  'К': 'K',
    'л': 'l',  'Л': 'L',
    'м': 'm',  'М': 'M',
    'н': 'n',  'Н': 'N',
    'о': 'o',  'О': 'O',
    'п': 'p',  'П': 'P',
    'р': 'r',  'Р': 'R',
    'с': 's',  'С': 'S',
    'т': 't',  'Т': 'T',
    'у': 'u',  'У': 'U',
    'ф': 'f',  'Ф': 'F',
    'х': 'x',  'Х': 'X',
    'ц': 's',  'Ц': 'S',
    'ч': 'ch', 'Ч': 'Ch',
    'ш': 'sh', 'Ш': 'Sh',
    'ъ': '\u02bc', 'Ъ': '\u02bc',
    'ь': '',   'Ь': '',
    'э': 'e',  'Э': 'E',
    'ю': 'yu', 'Ю': 'Yu',
    'я': 'ya', 'Я': 'Ya',
    'ў': 'o\u02bb', 'Ў': 'O\u02bb',
    'қ': 'q',  'Қ': 'Q',
    'ғ': 'g\u02bb', 'Ғ': 'G\u02bb',
    'ҳ': 'h',  'Ҳ': 'H',
}

CYRILLIC_VOWELS = (
    'а', 'А', 'е', 'Е', 'ё', 'Ё', 'и', 'И', 'о', 'О',
    'у', 'У', 'э', 'Э', 'ю', 'Ю', 'я', 'Я', 'ў', 'Ў'
)


# ══════════════════════════════════════════════════════
#  ЛАТИНИЦА → КИРИЛЛИЦА
# ══════════════════════════════════════════════════════

def to_cyrillic(text: str) -> str:
    """
    Транслитерирует узбекский текст из латиницы в кириллицу.
    Правила:
      1. ch → ч,  sh → ш  (до одиночных букв)
      2. yo' → йў  (до yo)
      3. yo → ё,  yu → ю,  ya → я,  ye → е
      4. o' / oʻ → ў,  g' / gʻ → ғ
      5. В начале слова: ye → е,  e → э
      6. После гласной:  ye → е,  e → э
      7. Одиночные буквы по таблице LATIN_TO_CYRILLIC
    """
    # Нормализация апострофов (windows vs mac)
    text = text.replace('\u02bb', '\u2018')
    text = text.replace('\u2019', '\u2018')

    # 1. Диграфы — первый проход (ch, sh, yo')
    compounds_first = {
        'ch': 'ч', 'Ch': 'Ч', 'CH': 'Ч',
        'sh': 'ш', 'Sh': 'Ш', 'SH': 'Ш',
        'yo\u2018': 'йў', 'Yo\u2018': 'Йў', 'YO\u2018': 'ЙЎ',
    }
    text = re.sub(
        r'(%s)' % '|'.join(re.escape(k) for k in compounds_first),
        lambda m: compounds_first[m.group(1)],
        text
    )

    # 2. Диграфы — второй проход (yo, yu, ya, ye, o', g')
    compounds_second = {
        'yo': 'ё', 'Yo': 'Ё', 'YO': 'Ё',
        'yu': 'ю', 'Yu': 'Ю', 'YU': 'Ю',
        'ya': 'я', 'Ya': 'Я', 'YA': 'Я',
        'ye': 'е', 'Ye': 'Е', 'YE': 'Е',
        'o\u2018': 'ў', 'O\u2018': 'Ў',
        'o\u02bb': 'ў', 'O\u02bb': 'Ў',
        'g\u2018': 'ғ', 'G\u2018': 'Ғ',
        'g\u02bb': 'ғ', 'G\u02bb': 'Ғ',
    }
    text = re.sub(
        r'(%s)' % '|'.join(re.escape(k) for k in compounds_second),
        lambda m: compounds_second[m.group(1)],
        text
    )

    # 3. В начале слова: ye → е,  e → э
    beginning_rules = {
        'ye': 'е', 'Ye': 'Е', 'YE': 'Е',
        'e': 'э',  'E': 'Э',
    }
    text = re.sub(
        r'\b(%s)' % '|'.join(re.escape(k) for k in beginning_rules),
        lambda m: beginning_rules[m.group(1)],
        text
    )

    # 4. После гласной: ye → е,  e → э
    after_vowel_rules = {
        'ye': 'е', 'Ye': 'Е', 'YE': 'Е',
        'e': 'э',  'E': 'Э',
    }
    text = re.sub(
        r'(%s)(%s)' % (
            '|'.join(re.escape(v) for v in LATIN_VOWELS),
            '|'.join(re.escape(k) for k in after_vowel_rules)
        ),
        lambda m: m.group(1) + after_vowel_rules[m.group(2)],
        text
    )

    # 5. Одиночные буквы
    text = re.sub(
        r'(%s)' % '|'.join(re.escape(k) for k in LATIN_TO_CYRILLIC),
        lambda m: LATIN_TO_CYRILLIC[m.group(1)],
        text
    )

    return text


# ══════════════════════════════════════════════════════
#  КИРИЛЛИЦА → ЛАТИНИЦА
# ══════════════════════════════════════════════════════

def to_latin(text: str) -> str:
    """
    Транслитерирует узбекский текст из кириллицы в латиницу.
    Правила:
      1. Сентябр → Sentabr,  Октябр → Oktabr  (исключения)
      2. В начале слова: ц → s,  е → ye
      3. После гласной:  ц → ts, е → ye
      4. Остальные по таблице CYRILLIC_TO_LATIN
    """
    # Исключения: сентябр / октябр
    text = re.sub(
        r'(сент|окт)([яЯ])(бр)',
        lambda m: m.group(1) + ('a' if m.group(2) == 'я' else 'A') + m.group(3),
        text,
        flags=re.IGNORECASE
    )

    # В начале слова
    beginning_rules = {
        'ц': 's', 'Ц': 'S',
        'е': 'ye', 'Е': 'Ye',
    }
    text = re.sub(
        r'\b(%s)' % '|'.join(re.escape(k) for k in beginning_rules),
        lambda m: beginning_rules[m.group(1)],
        text
    )

    # После гласной
    after_vowel_rules = {
        'ц': 'ts', 'Ц': 'Ts',
        'е': 'ye', 'Е': 'Ye',
    }
    text = re.sub(
        r'(%s)(%s)' % (
            '|'.join(re.escape(v) for v in CYRILLIC_VOWELS),
            '|'.join(re.escape(k) for k in after_vowel_rules)
        ),
        lambda m: m.group(1) + after_vowel_rules[m.group(2)],
        text
    )

    # Все остальные символы
    text = re.sub(
        r'(%s)' % '|'.join(re.escape(k) for k in CYRILLIC_TO_LATIN),
        lambda m: CYRILLIC_TO_LATIN[m.group(1)],
        text
    )

    return text


# ══════════════════════════════════════════════════════
#  АВТООПРЕДЕЛЕНИЕ АЛФАВИТА
# ══════════════════════════════════════════════════════

def detect_script(text: str) -> str:
    """
    Определяет алфавит текста.
    Возвращает: 'latin', 'cyrillic' или 'unknown'
    """
    if not text:
        return 'unknown'
    cyr = sum(1 for c in text if '\u0400' <= c <= '\u04FF')
    lat = sum(1 for c in text if c.isalpha() and ord(c) < 128)
    if cyr == 0 and lat == 0:
        return 'unknown'
    return 'cyrillic' if cyr > lat else 'latin'


def convert_to(text: str, target: str) -> str:
    """
    Конвертирует текст в нужный алфавит.
    target: 'latin' | 'cyrillic'
    Если текст уже в нужном алфавите — возвращает без изменений.
    """
    if not text or not text.strip():
        return text
    script = detect_script(text)
    if target == 'cyrillic' and script == 'latin':
        return to_cyrillic(text)
    if target == 'latin' and script == 'cyrillic':
        return to_latin(text)
    return text


def convert_dict(data: dict, target: str, keys: list = None) -> dict:
    """Конвертирует строковые значения словаря."""
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
# ══════════════════════════════════════════════════════

TEXTS_SPISAN = {
    'latin': {
        'tasdiq':    "«TASDIQLAYMAN»",
        'yil':       "yil",
        'title':     "TEXNIK XULOSA №",
        'intro':     "Biz, quyida imzo chekuvchilar — {org} {region} bo\u02bblim "
                     "yetakchi muhandisi {eng1} va yetakchi muhandis {eng2} lar, "
                     "quyida keltirilgan qurilmalarni texnik ko\u02bbrikdan o\u02bbtkazdik:",
        'qurilma':   "Qurilma nomi:",
        'inventar':  "Inventar raqami:",
        'col_qism':  "Qismlar nomi",
        'col_yarоq': "Foydalanishga\nyaroqliligi",
        'col_nosoz': "Nosozlik belgilari",
        'xulosa':    "XULOSA",
        'xulosa1':   "{names} zamonaviy talablarga javob bermaydi, ma\u02bbnaviy eskirgan. "
                     "Ta\u02bbmirlash uchun zarur bo\u02bblgan ehtiyot qismlar texnologik jarayondan "
                     "olib tashlanganligini hisobga olib, uni tiklash maqsadga muvofiq emas.",
        'xulosa2':   "Yuqorida keltirib o\u02bbtilgan jadvaldagi qurilmalarni xolatlaridan "
                     "kelib chiqib, komissiya asosiy vositalar ro\u02bbyhxatidan chiqarilsin "
                     "degan xulosaga keldi.",
        'muhandis':  "Yetakchi muhandis:",
        'mp':        "M.P.",
        'sana':      "Sana:",
        'inv_label': "Inventar raqami:",
    },
    'cyrillic': {
        'tasdiq':    "«ТАСДИҚЛАНДИ»",
        'yil':       "йил",
        'title':     "ТЕХНИК ХУЛОСА №",
        'intro':     "Биз, қуйида имзо чекувчилар — {org} {region} бўлим "
                     "етакчи муҳандиси {eng1} ва етакчи муҳандис {eng2} лар, "
                     "қуйида келтирилган қурилмаларни техник кўрикдан ўтказдик:",
        'qurilma':   "Қурилма номи:",
        'inventar':  "Инвентар рақами:",
        'col_qism':  "Қисмлар номи",
        'col_yarоq': "Фойдаланишга\nяроқлилиги",
        'col_nosoz': "Носозлик белгилари",
        'xulosa':    "ХУЛОСА",
        'xulosa1':   "{names} замонавий талабларга жавоб бермайди, маънавий эскирган. "
                     "Таъмирлаш учун зарур бўлган эҳтиёт қисмлар технологик жараёндан "
                     "олиб ташланганлигини ҳисобга олиб, уни тиклаш мақсадга мувофиқ эмас.",
        'xulosa2':   "Юқорида келтириб ўтилган жадвалдаги қурилмаларни ҳолатларидан "
                     "келиб чиқиб, комиссия асосий воситалар рўйхатидан чиқарилсин "
                     "деган хулосага келди.",
        'muhandis':  "Етакчи муҳандис:",
        'mp':        "М.П.",
        'sana':      "Сана:",
        'inv_label': "Инвентар рақами:",
    },
}

TEXTS_UST = {
    'latin': {
        'tasdiq':      "«TASDIQLAYMAN»",
        'rahbar':      "rahbari",
        'yil':         "yil",
        'manzil':      "Manzil:",
        'title':       "USKUNANI O\u02bbRNATISH DALOLATNOMASI",
        'dalolatnoma': "Dalolatnoma №",
        'intro':       "Mazkur dalolatnoma texnologik jarayonlarni modernizatsiya qilish "
                       "doirasida quyidagi uskunalar o\u02bbnatilganligini tasdiqlash "
                       "maqsadida tuzildi.",
        'col_num':     "№",
        'col_sana':    "O\u02bbnatilgan\nsana",
        'col_uskuna':  "Uskuna\n(Model)",
        'col_serial':  "Seriya\nraqami",
        'col_qayer':   "Qayerga\no\u02bbnatilgan",
        'outro':       "Yuqorida ko\u02bbsatilgan uskunalar belgilangan tartibda "
                       "o\u02bbnatildi, tekshirildi va ekspluatatsiyaga topshirildi.",
        'masul':       "O\u02bbnatish uchun mas\u02bbul shaxs",
        'mp':          "M.P.",
        'sana':        "Sana:",
    },
    'cyrillic': {
        'tasdiq':      "«ТАСДИҚЛАНДИ»",
        'rahbar':      "раҳбари",
        'yil':         "йил",
        'manzil':      "Манзил:",
        'title':       "УСКУНАНИ ЎРНАТИШ ДАЛОЛАТНОМАСИ",
        'dalolatnoma': "Далолатнома №",
        'intro':       "Мазкур далолатнома технологик жараёнларни модернизация қилиш "
                       "доирасида қуйидаги ускуналар ўрнатилганлигини тасдиқлаш "
                       "мақсадида тузилди.",
        'col_num':     "№",
        'col_sana':    "Ўрнатилган\nсана",
        'col_uskuna':  "Ускуна\n(Модел)",
        'col_serial':  "Серия\nрақами",
        'col_qayer':   "Қаерга\nўрнатилган",
        'outro':       "Юқорида кўрсатилган ускуналар белгиланган тартибда "
                       "ўрнатилди, текширилди ва эксплуатацияга топширилди.",
        'masul':       "Ўрнатиш учун масъул шахс",
        'mp':          "М.П.",
        'sana':        "Сана:",
    },
}

MONTHS = {
    'latin': {
        1: 'yanvar',  2: 'fevral', 3: 'mart',    4: 'aprel',
        5: 'may',     6: 'iyun',   7: 'iyul',    8: 'avgust',
        9: 'sentabr', 10: 'oktabr', 11: 'noyabr', 12: 'dekabr',
    },
    'cyrillic': {
        1: 'январ',   2: 'феврал', 3: 'март',    4: 'апрел',
        5: 'май',     6: 'июн',    7: 'июл',     8: 'август',
        9: 'сентябр', 10: 'октябр', 11: 'ноябр',  12: 'декабр',
    },
}
