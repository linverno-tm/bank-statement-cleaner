# -*- coding: utf-8 -*-
"""
SoddaHisobot — bank hisobotlarini soddalashtiruvchi ilovaning BUTUN
mantiqi va interfeysi.

Bu fayl GitHub'da turadi va launcher (.exe) uni har ishga tushganda
yuklab olib bajaradi. Ya'ni tuzatish yoki yangi funksiya chiqarish uchun
shu faylni GitHub'ga push qilish kifoya — foydalanuvchiga hech narsa
qayta o'rnatish kerak emas, .exe o'zgarmaydi.

Shuning uchun ilova ichida hech qanday "avtomatik yangilanish" kodi
yo'q: yangilanish launcher darajasida, faylni yuklab olish orqali
sodir bo'ladi.

Tuzilishi:
  1-QISM  — tasniflash: xom hisobotni o'qish, kontragentni guruhga ajratish
  2-QISM  — hisobot yasash: soddalashtirilgan .xlsx chiqarish
  3-QISM  — interfeys (Tkinter GUI) va main()
"""
import datetime
import json
import os
import platform
import queue
import re
import subprocess
import sys
import tempfile
import threading
import traceback
import urllib.request
from collections import defaultdict
from copy import copy
from decimal import Decimal

import openpyxl
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter

try:
    import winreg
except ImportError:            # Windows emas (sayt serveri) - Excel qidirilmaydi
    winreg = None

# Oyna kutubxonalari faqat desktop ilovaga kerak. Sayt (bank ko'chirmasini
# brauzerda soddalashtirish) core.py ning faqat hisob qismini ishlatadi va
# serverda tkinter bo'lmasligi mumkin.
try:
    import tkinter as tk
    from tkinter import font as tkfont
    from tkinter import ttk, filedialog, messagebox

    import sv_ttk
except ImportError:
    class _NoTk:
        Tk = Toplevel = object
    tk = _NoTk()
    tkfont = ttk = filedialog = messagebox = sv_ttk = None

# Ilovaning joriy versiyasi. Launcher .exe o'zgarmaydi, shuning uchun
# foydalanuvchi ko'radigan versiya aynan shu fayldan olinadi.
CORE_VERSION = "2.2.3"

# Kodni qaysi shoxobchadan olganini launcher.py exec() dan oldin shu
# nom bilan uzatadi. To'g'ridan-to'g'ri `python core.py` bilan ishga
# tushirilganda hech kim uzatmaydi — o'shanda master deb hisoblanadi.
SOURCE_BRANCH = globals().get("SOURCE_BRANCH", "master")

# Ilova ochilganda faqat kompyuter nomi va versiyani yuboradi — bu
# "kimda qaysi kod ishlab turibdi" degan savolga javob berish uchun.
_PING_URL = "https://soddahisobot-telemetry.tasks-bot.workers.dev/ping"
_ERROR_URL = "https://soddahisobot-telemetry.tasks-bot.workers.dev/xato"
_USER_AGENT = "SoddaHisobot-Core"


def send_error(xato):
    """Yuz bergan xatoni serverga xabar qiladi. Alohida oqimda chaqiriladi.

    Nima uchun: telemetriya kim qaysi versiyada ekanini aytadi, lekin
    ilova foydalanuvchida yiqilsa bu hech qayerda ko'rinmasdi —
    "menda ishlamadi" degan gapni telefonda tekshirishga to'g'ri kelardi.

    Faqat xato matni yuboriladi, fayl yo'llari va hisobot mazmuni emas."""
    try:
        payload = json.dumps({
            "host": platform.node() or "noma'lum",
            "version": CORE_VERSION,
            "xato": str(xato)[:1000],
        }).encode("utf-8")
        req = urllib.request.Request(
            _ERROR_URL,
            data=payload,
            headers={"Content-Type": "application/json", "User-Agent": _USER_AGENT},
            method="POST",
        )
        urllib.request.urlopen(req, timeout=5).close()
    except Exception:
        pass


def report_error(xato):
    """Xabarni fon oqimida yuboradi — interfeys kutib qolmasin."""
    threading.Thread(target=send_error, args=(xato,), daemon=True).start()


def send_ping():
    """Ochilganini xabar qiladi. To'liq "ovozsiz": internet yo'q bo'lsa
    ilova ishlashiga ta'sir qilmaydi. Alohida oqimda chaqirilishi kerak."""
    try:
        payload = json.dumps({
            "host": platform.node() or "noma'lum",
            "version": CORE_VERSION,
        }).encode("utf-8")
        req = urllib.request.Request(
            _PING_URL,
            data=payload,
            headers={"Content-Type": "application/json", "User-Agent": _USER_AGENT},
            method="POST",
        )
        urllib.request.urlopen(req, timeout=5).close()
    except Exception:
        pass


# ======================================================================
# 1-QISM — TASNIFLASH
# ======================================================================

TOTAL_ROW_MARKERS = ("итоговый оборот", "итого", "jami", "всего")


def is_total_row(first_cell):
    if first_cell is None:
        return False
    return str(first_cell).strip().lower().startswith(TOTAL_ROW_MARKERS)


# Sana ustuni "faqat sana"dan iboratmi? Hisobot oxirida bank qo'shimcha
# qatorlar qo'shadi: "Обороты по дебету: кол-во 8", "Исходящий остаток на
# 31.08.2026", bank nomi/manzili va h.k. Ular operatsiya emas, lekin
# ustunlarida summa turgani uchun avval oddiy qator deb o'qilib, "?"
# nomsiz kontragent sifatida so'ralib qolardi.
#
# Kalit so'z ro'yxatiga tayanmaymiz (har bankda har xil yoziladi) —
# o'rniga tuzilmaga tayanamiz: HAQIQIY operatsiya qatorida sana ustuni
# TOZA sana bo'ladi. "Исходящий остаток на 31.08.2026" ichida sana bor,
# lekin butun katak sana emas — shuning uchun to'liq moslik talab qilinadi.
_DATE_ONLY_RX = re.compile(
    r"^\s*\d{1,2}[./-]\d{1,2}[./-]\d{2,4}"      # 03.08.2026 / 3-8-26
    r"(?:[\s,]+\d{1,2}:\d{2}(?::\d{2})?)?"       # ixtiyoriy vaqt
    r"(?:\.\d+)?\s*$"                             # ixtiyoriy mikrosoniya
)


def looks_like_date(value):
    """Katak butunlay sanadan iboratmi (matn ichidagi sana emas)."""
    if value is None or value == "":
        return False
    if isinstance(value, (datetime.datetime, datetime.date)):
        return True
    return bool(_DATE_ONLY_RX.match(str(value)))


# Bitta hisob raqam (masalan G'aznachilik / Молия вазирлиги hisobvarag'i)
# orqali butunlay boshqa-boshqa maqsaddagi to'lovlar o'tadi: foyda solig'i,
# QQS, elektr uchun to'lov... Shuning uchun bu qoidalar hisob raqam bo'yicha
# o'rgatilgan guruhdan ham USTUN turadi — to'lov maqsadi (Назначение
# платежа) matni bu yerda hal qiluvchi hisoblanadi.
# G'aznachilik / Moliya vazirligi hisobvarag'i orqali butunlay har xil
# maqsaddagi to'lovlar o'tadi: foyda solig'i, QQS, elektr, ijtimoiy soliq...
# Shuning uchun bunday kontragentni HECH QACHON hisob raqam yoki nom
# bo'yicha "o'rganib" qo'ymaymiz — aks holda bir marta "QQS" deb
# belgilangach, keyingi barcha to'lovlar ham QQS bo'lib ketardi.
#
# Bunday qatorlar uchun tartib: avval to'lov maqsadi qoidalari (soliq
# turi, elektr...) ishlaydi; ular topa olmasa — foydalanuvchidan
# SO'RALADI, va javob faqat shu faylga tegishli bo'ladi.
ALWAYS_ASK_NAME_RX = re.compile(
    r"молия\s*вазирлиг|газначилиг|ягона\s*газна|казначейств",
    re.I,
)

# Foydalanuvchi shu ish davomida bergan javoblar: (hisob, maqsad izi) -> guruh.
# Diskka YOZILMAYDI — "doim so'ralsin" degani shu.
SESSION_OVERRIDES = {}


def purpose_signature(text):
    """To'lov maqsadidan barqaror "iz" yasaydi: raqamlar, hujjat nomerlari
    va ajratuvchilar olib tashlanadi, faqat ma'noli so'zlar qoladi.
    Shunda "...~32 Фойда солиги учун олдиндан тулов" va boshqa hujjat
    raqamli xuddi shu to'lov bitta savol sifatida ko'rinadi."""
    t = re.sub(r"[\d~№/\\.,:;()-]+", " ", str(text or ""))
    return " ".join(t.split()).lower()[:70]


def is_always_ask(name, account=None):
    """Kontragent ko'p maqsadli (har safar so'ralishi kerak) mi?"""
    return bool(ALWAYS_ASK_NAME_RX.search(str(name or "")))


def set_session_override(account, purpose, category):
    """Foydalanuvchi javobini shu ish uchun eslab qoladi (diskka emas)."""
    SESSION_OVERRIDES[(str(account or ""), purpose_signature(purpose))] = category


def clear_session_overrides():
    SESSION_OVERRIDES.clear()


PURPOSE_FIRST_RULES = [
    # Soliqdan QAYTGAN summa — bu soliq to'lovi emas, qaytimi. Ro'yxatning
    # eng boshida turadi, chunki matnda "ПНФЛ"/"даромад солиги" ham
    # uchraydi va quyidagi qoidalar uni oddiy soliq to'lovi deb olib
    # qo'yardi. Faqat "возврат" so'ziga qaramaymiz: yetkazib beruvchidan
    # tovar qaytarilganda ham shu so'z ishlatiladi, shuning uchun soliq
    # belgisi (ГНИ, ПНФЛ, солик, бюджет) ham talab qilinadi.
    (
        re.compile(
            r"(?=.*возврат)(?=.*(?:пнфл|гни|нало[гж]|соли[кқгғ]|бюджет))",
            re.I | re.S,
        ),
        "ДСИ",
    ),
    (re.compile(r"фойда\s*соли[гғ]и", re.I), "солик фойда"),
    # DIQQAT: bu yerda faqat aniq "Кушилган киймат солиги" iborasi tekshiriladi.
    # Umumiy "НДС" so'zi ATAYLAB kiritilmagan — oddiy tovar to'lovlarida ham
    # "Сумма ... В т.ч. НДС (12%) ..." deb yoziladi, va u paytda bu soliq
    # to'lovi emas, balki narxning tarkibiy qismi. Umumiy "НДС" qoidasi
    # quyida, past darajali TEXT_RULES ichida qoldirilgan.
    (re.compile(r"[кқ]ушилган\s*[кқ]иймат\s*соли[гғ]и", re.I), "солик КҚС"),
    (re.compile(r"ижтимоий\s*соли[кқ]", re.I), "солик ижтимоий"),
    (re.compile(r"даромадидан\s*олинадиган\s*соли[кқ]|даромад\s*соли[гғ]и", re.I), "солик даромад"),
    (re.compile(r"пенсия\s*бадалига", re.I), "солик пенсия"),
    (re.compile(r"сув\s*таъминоти|ичимлик\s*сув", re.I), "коммунал"),
    (re.compile(r"табиий\s*газ|газ\s*учун", re.I), "коммунал"),
    # Elektr alohida guruh — kommunalga qo'shilmaydi.
    (re.compile(r"фойдаланилган\s*электр|электр\s*учун|электр\s*энергия", re.I), "электр"),
    (re.compile(r"ижара\s*ту[лл]ови", re.I), "ижара"),
]


# Ordered keyword rules: (regex, category, confidence)
# confidence 'high' = auto-apply, 'guess' = apply but flag for review
TEXT_RULES = [
    (re.compile(r"начисленные\s*%%", re.I), "банк хизмати", "high"),  # naименование check done separately too
    (re.compile(r"smartvista|humo\s*\(|возмещение клиенту по покупкам тсп", re.I), "Терминал", "high"),
    (re.compile(r"tbc\s*fin\s*service|tbc\s*bnpl", re.I), "ф пайми", "high"),
    (re.compile(r"variant\s*retail\s*finance", re.I), "ф вариант", "high"),
    (re.compile(r"куриклаш", re.I), "куриклаш", "high"),
    (re.compile(r"зарплата|иш\s*ха[кқ]и", re.I), "иш хаки ПК", "high"),
    (re.compile(r"ижтимоий\s*соли[кқ]", re.I), "солик ижтимоий", "high"),
    (re.compile(r"даромадидан\s*олинадиган\s*соли[кқ]|даромад\s*соли[гғ]и", re.I), "солик даромад", "high"),
    (re.compile(r"пенсия\s*бадалига", re.I), "солик пенсия", "high"),
    (re.compile(r"сорж", re.I), "СОРЖ", "high"),
    # JNS LABS faqat shu xizmatni ko'rsatadi — umumiy "хизмат" guruhiga
    # qo'shilmasin, alohida qator bo'lib chiqsin.
    (re.compile(r"ароматизац|jns\s*labs", re.I), "ароматизация", "high"),
    # Lower confidence guesses (new patterns not seen in the DDD reference file yet).
    # Finance-partner style wording ("Публичная оферта" + "ген соглашение"-like BNPL
    # contracts) is checked BEFORE the generic "НДС" substring rule, since a BNPL
    # settlement text often mentions VAT only incidentally as a line item.
    (re.compile(r"оплата\s*100\s*%.*по\s*договору\s*публичная\s*оферта", re.I), "ф (аникланмаган)", "guess"),
    (re.compile(r"ижара\s*тулови|ижара\s*ту[лл]ови", re.I), "ижара", "guess"),
    (re.compile(r"фойдаланилган\s*электр|электр\s*учун", re.I), "электр", "guess"),
    (re.compile(r"консалтинг|konsalting", re.I), "хизмат", "guess"),
    (re.compile(r"фойда\s*соли[гғ]и", re.I), "солик фойда", "guess"),
    (re.compile(r"кушилган\s*[кқ]иймат\s*соли[гғ]и|\bндс\b", re.I), "солик КҚС", "guess"),
]

GOODS_PURCHASE_HINT = re.compile(
    r"маиший техника|maishiy texnika|телефон|планшет|товар|жихоз|асбоб|"
    r"курилиш махсулот|дастурий таъминот|болалар уйинчок|мобил алока воситалари|"
    r"бытовой техник|бытавой техник|посуд",
    re.I,
)

# Vendor (INN) -> category, confirmed from a real "ДДД Август" file the boss
# already produced by hand. Same suppliers keep recurring across different
# companies' statements (same Andijon retail/appliance trade network), so
# this dictionary is the highest-confidence signal available and should grow
# over time (persist_dictionary.json).
_BUILTIN_VENDOR_INN = {
    "303389344": "МЕБ",       # Andijon Ravnaqi МЧЖ
    "207180749": "МЕБ",       # Goods And Services Impex
    "306776074": "ф пайми",   # TBC FIN SERVICE
    "312422124": "ф пайми",   # TBC BNPL
    "307490921": "ф вариант", # VARIANT RETAIL FINANCE
    "200292692": "солик пенсия",  # tuman ДСИ
    "200237592": "куриклаш",  # ИИБ КОШИДАГИ КУРИКЛАШ БУЛИМИ
    "310692639": "МЕБ",       # BIG ELECTRONICA МЧЖ
    "311019672": "МЕБ",       # ELEKTROMAX МЧЖ
    "309018562": "МЕБ",       # TOSHIBA AND MCHJ
    "310890749": "МЕБ",       # AMIR GROUP DU
}
KNOWN_VENDOR_INN = dict(_BUILTIN_VENDOR_INN)



def _base_dir():
    # When bundled by PyInstaller, __file__ points into a temp extraction
    # folder, so use the actual .exe location instead.
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


DICT_FILE = "persist_dictionary.json"


def _data_dir():
    """O'rganilgan guruhlar saqlanadigan papka.

    .exe yonida EMAS: .exe boshqa papkaga ko'chirilsa yoki qaytadan
    yuklab olinsa, yillar davomida to'plangan lug'at eski papkada qolib
    ketardi va ilova hammasini unutgandek bo'lardi. %LOCALAPPDATA%
    esa .exe qayerda turishidan qat'i nazar o'zgarmaydi — kesh,
    kalit va branch.txt ham o'sha yerda.

    SODDA_DATA_DIR - sayt o'z papkasini beradi (desktop ilovada yo'q)."""
    sayt = os.environ.get("SODDA_DATA_DIR")
    if sayt:
        os.makedirs(sayt, exist_ok=True)
        return sayt
    lokal = os.environ.get("LOCALAPPDATA")
    if not lokal:
        return _base_dir()
    yol = os.path.join(lokal, "SoddaHisobot")
    try:
        os.makedirs(yol, exist_ok=True)
    except OSError:
        return _base_dir()
    return yol


_DICT_PATH = os.path.join(_data_dir(), DICT_FILE)


def _migrate_dictionary():
    """Eski joydagi (.exe yonidagi) lug'atni yangi joyga ko'chiradi.

    Nusxalanadi, ko'chirilmaydi: eski fayl zaxira bo'lib joyida qoladi.
    Yangi joyda lug'at paydo bo'lgach bu funksiya hech narsa qilmaydi,
    ya'ni keyingi o'zgarishlar eski fayl bilan qayta yozilmaydi."""
    if os.path.exists(_DICT_PATH):
        return
    eski = os.path.join(_base_dir(), DICT_FILE)
    if os.path.abspath(eski) == os.path.abspath(_DICT_PATH):
        return
    try:
        with open(eski, encoding="utf-8") as f:
            mazmun = f.read()
    except OSError:
        return
    try:
        with open(_DICT_PATH, "w", encoding="utf-8") as f:
            f.write(mazmun)
    except OSError:
        pass


_migrate_dictionary()

# Name-keyed learned mappings, for counterparties with no ИНН. Stored in the
# same persist_dictionary.json file, namespaced with a "NAME::" key prefix so
# it never collides with the (purely numeric) ИНН keys.
KNOWN_VENDOR_NAME = {}
# Foydalanuvchi o'zi kiritgan ИНН/hisob kalitlari (kodga yozilganlar emas).
LEARNED_VENDOR_KEYS = set()
_NAME_KEY_PREFIX = "NAME::"

# Bitta kontragent yo'nalishiga qarab boshqa guruhga tushishi mumkin:
# masalan chiqimda "Авто тулов", tushumda esa butunlay boshqa nom.
# Kalit oxiriga "@D" (debet/chiqim) yoki "@K" (kredit/tushum) qo'shiladi.
# Belgisiz kalit avvalgidek ikkala yo'nalish uchun ishlaydi, ya'ni eski
# yozuvlar o'zgarishsiz qolaveradi.
DIR_DEBIT = "D"
DIR_CREDIT = "K"
_DIR_SEP = "@"
KNOWN_VENDOR_INN_DIR = {}    # (kalit, yo'nalish) -> guruh
KNOWN_VENDOR_NAME_DIR = {}   # (nom, yo'nalish) -> guruh
KNOWN_PURPOSE_TEXT_DIR = {}  # (matn, yo'nalish) -> guruh


def split_direction(kalit):
    """Kalitdan yo'nalish belgisini ajratadi: ("2020...", "K")."""
    if len(kalit) > 2 and kalit[-2] == _DIR_SEP and kalit[-1] in (DIR_DEBIT, DIR_CREDIT):
        return kalit[:-2], kalit[-1]
    return kalit, None


def with_direction(kalit, yonalish):
    return f"{kalit}{_DIR_SEP}{yonalish}" if yonalish else kalit


def row_direction(debit, credit):
    """Qator qaysi yo'nalishda — chiqim (debet) yoki tushum (kredit)."""
    if debit:
        return DIR_DEBIT
    if credit:
        return DIR_CREDIT
    return None
# Foydalanuvchi o'zi qo'shgan "to'lov maqsadi matni -> guruh" qoidalari.
# Bitta hisob raqam orqali turli maqsaddagi to'lovlar o'tganda, ilovada
# tayyor qoida bo'lmasa, foydalanuvchi shu yerga o'zi qoida qo'sha oladi.
KNOWN_PURPOSE_TEXT = {}
_TEXT_KEY_PREFIX = "TEXT::"


# Ilova avval ba'zi guruh nomlarini lotin harflari bilan yozardi
# ("солик QQS", "Avto to'lov"). Hisobot butunlay kirillcha bo'lishi kerak,
# shuning uchun nomlar o'zgartirildi. Foydalanuvchining eski yozuvlari
# yo'qolib qolmasligi uchun ular ochilishda avtomatik ko'chiriladi.
_LEGACY_GROUP_RENAMES = {
    "солик QQS": "солик КҚС",
    "ф (aniqlanmagan)": "ф (аникланмаган)",
    "ARALASH": "АРАЛАШ",
    "Click": "Клик",
    "Avto to'lov": "Авто тулов",
}


def _migrate_legacy_groups():
    """Diskdagi lug'atda eski (lotincha) guruh nomlari qolgan bo'lsa,
    ularni yangi kirillcha nomlarga almashtiradi."""
    try:
        with open(_DICT_PATH, "r", encoding="utf-8") as f:
            raw = json.load(f)
    except (OSError, ValueError):
        return
    yangi = {k: _LEGACY_GROUP_RENAMES.get(v, v) for k, v in raw.items()}
    if yangi != raw:
        try:
            with open(_DICT_PATH, "w", encoding="utf-8") as f:
                json.dump(yangi, f, ensure_ascii=False, indent=2, sort_keys=True)
        except OSError:
            pass


def _load_persist_dict():
    try:
        with open(_DICT_PATH, "r", encoding="utf-8") as f:
            raw = json.load(f)
    except FileNotFoundError:
        raw = {}
    for k, v in raw.items():
        if not isinstance(k, str):
            continue
        toza, yonalish = split_direction(k)
        if toza.startswith(_TEXT_KEY_PREFIX):
            matn = toza[len(_TEXT_KEY_PREFIX):]
            if yonalish:
                KNOWN_PURPOSE_TEXT_DIR[(matn, yonalish)] = v
            else:
                KNOWN_PURPOSE_TEXT[matn] = v
        elif toza.startswith(_NAME_KEY_PREFIX):
            nomi = toza[len(_NAME_KEY_PREFIX):]
            if yonalish:
                KNOWN_VENDOR_NAME_DIR[(nomi, yonalish)] = v
            else:
                KNOWN_VENDOR_NAME[nomi] = v
        else:
            if yonalish:
                KNOWN_VENDOR_INN_DIR[(toza, yonalish)] = v
            else:
                KNOWN_VENDOR_INN[toza] = v
            # Kodga yozilganlardan ajratib qo'yamiz: "noldan boshlash"
            # rejimida faqat foydalanuvchi o'zi kiritganlari ishlaydi.
            LEARNED_VENDOR_KEYS.add(toza)


_migrate_legacy_groups()
_load_persist_dict()


def save_learned_category(identifier, name, category, yonalish=None):
    """Persist a user-supplied category for a previously-unresolved ("?")
    counterparty, so future reports auto-classify it. Keyed by Xisob raqam
    (Счет) when available (most reliable, and stable per counterparty);
    falls back to the exact counterparty name otherwise. `identifier` may
    also be an ИНН, for backward compatibility with entries learned by
    older versions of this app. Updates both the on-disk dictionary and the
    in-memory maps so the rest of the current run benefits immediately."""
    identifier = str(identifier).strip() if identifier else ""
    name = (name or "").strip()
    if not identifier and not name:
        return
    try:
        with open(_DICT_PATH, "r", encoding="utf-8") as f:
            raw = json.load(f)
    except FileNotFoundError:
        raw = {}
    if identifier.startswith(_TEXT_KEY_PREFIX):
        phrase = identifier[len(_TEXT_KEY_PREFIX):]
        raw[with_direction(identifier, yonalish)] = category
        if yonalish:
            KNOWN_PURPOSE_TEXT_DIR[(phrase, yonalish)] = category
        else:
            KNOWN_PURPOSE_TEXT[phrase] = category
    elif identifier:
        raw[with_direction(identifier, yonalish)] = category
        if yonalish:
            KNOWN_VENDOR_INN_DIR[(identifier, yonalish)] = category
        else:
            KNOWN_VENDOR_INN[identifier] = category
        LEARNED_VENDOR_KEYS.add(identifier)
    else:
        raw[with_direction(f"{_NAME_KEY_PREFIX}{name}", yonalish)] = category
        if yonalish:
            KNOWN_VENDOR_NAME_DIR[(name, yonalish)] = category
        else:
            KNOWN_VENDOR_NAME[name] = category
    with open(_DICT_PATH, "w", encoding="utf-8") as f:
        json.dump(raw, f, ensure_ascii=False, indent=2, sort_keys=True)


def load_all_mappings():
    """Return the raw on-disk mapping ({ИНН or 'NAME::<nomi>': kategoriya}),
    for a UI to list/edit directly. Built-in (hardcoded) vendor mappings are
    NOT included here — only what the user has explicitly taught the app."""
    try:
        with open(_DICT_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        return {}


def save_all_mappings(mapping):
    """Overwrite persist_dictionary.json wholesale with `mapping` (as
    returned/edited from load_all_mappings) and refresh the in-memory maps
    so a Guruhlar-manager UI's changes apply immediately, without discarding
    the built-in defaults."""
    with open(_DICT_PATH, "w", encoding="utf-8") as f:
        json.dump(mapping, f, ensure_ascii=False, indent=2, sort_keys=True)
    KNOWN_VENDOR_INN.clear()
    KNOWN_VENDOR_INN.update(_BUILTIN_VENDOR_INN)
    KNOWN_VENDOR_NAME.clear()
    KNOWN_PURPOSE_TEXT.clear()
    for k, v in mapping.items():
        if isinstance(k, str) and k.startswith(_TEXT_KEY_PREFIX):
            KNOWN_PURPOSE_TEXT[k[len(_TEXT_KEY_PREFIX):]] = v
        elif isinstance(k, str) and k.startswith(_NAME_KEY_PREFIX):
            KNOWN_VENDOR_NAME[k[len(_NAME_KEY_PREFIX):]] = v
        else:
            KNOWN_VENDOR_INN[k] = v


def unique_counterparties(rows):
    """Return every unique counterparty found in a raw statement (keyed by
    Xisob raqam / Счет, or by name when the account number is missing),
    regardless of whether it was already classified. Used by the
    Guruhlar-manager UI to let the user pick a real counterparty out of an
    actual file (like browsing the statement itself) instead of typing an
    identifier by hand."""
    parties = {}
    for r in rows:
        account = str(r["account"]).strip() if r["account"] else ""
        inn = str(r["inn"]).strip() if r["inn"] else ""
        name = str(r["name"] or "").strip()
        if not account and not name:
            continue
        key = account if account else f"{_NAME_KEY_PREFIX}{name}"
        if key not in parties:
            parties[key] = {
                "account": account,
                "inn": inn,
                "name": name,
                "mfo": str(r.get("mfo") or "").strip(),
                "sample": str(r["purpose"] or "")[:200],
            }
    return parties


def find_unresolved(rows):
    """Scan rows and return unique unclassified ("?") counterparties, keyed
    by Xisob raqam / Счет (or by name when the account number is missing),
    for the caller to ask the user about before generating the report."""
    unresolved = {}
    for r in rows:
        cat, conf = classify_row(
            r["op"], r["name"], r["purpose"], r["inn"], r["account"],
            row_direction(r["debit"], r["credit"]),
        )
        if conf != "review":
            continue
        account = str(r["account"]).strip() if r["account"] else ""
        name = str(r["name"] or "").strip()
        koʻp_maqsadli = is_always_ask(name, account)
        if koʻp_maqsadli:
            # Har bir maqsad alohida savol bo'lsin: bitta hisob raqamda
            # soliq ham, elektr ham bo'lishi mumkin.
            key = f"{account}|{purpose_signature(r['purpose'])}"
        else:
            key = account if account else f"{_NAME_KEY_PREFIX}{name}"
        entry = unresolved.setdefault(key, {
            "account": account,
            "inn": str(r["inn"]).strip() if r["inn"] else "",
            "name": name,
            "mfo": str(r.get("mfo") or "").strip(),
            "sample": str(r["purpose"] or "")[:200],
            "always_ask": koʻp_maqsadli,
            "purpose": str(r["purpose"] or ""),
            "count": 0,
        })
        entry["count"] += 1
    return unresolved


def classify_row(op, name, text, inn, account=None, yonalish=None):
    """Classify a SINGLE row by its own text/name/Xisob raqam. Never looks
    at other rows sharing the same raw account number.

    Matching is done primarily by the counterparty's Xisob raqam (Счет).
    ИНН is kept as a secondary fallback so categories learned by older
    versions of this app (persist_dictionary.json entries keyed by ИНН)
    keep working after this update.

    ISTISNO: PURPOSE_FIRST_RULES — bitta hisob raqam orqali turli xil
    to'lovlar o'tadigan holatlar (G'aznachilik hisobvarag'i: foyda solig'i,
    QQS, elektr uchun to'lov) uchun to'lov maqsadi matni hisob raqamdan
    ustun turadi, aks holda hammasi bitta guruhga tushib qolardi."""
    name = name or ""
    text = text or ""

    # "Faqat o'zim" rejimida kodga yozilgan qoidalar chetlab o'tiladi:
    # ular bir kishining atamalariga moslangan. Foydalanuvchining o'zi
    # kiritgan yozuvlari esa ishlayveradi.
    faqat_ozim = scratch_mode()

    if not faqat_ozim and re.search(r"начисленные\s*%%", name, re.I):
        return "банк хизмати", "high"

    account_str = str(account).strip() if account else ""

    # Foydalanuvchi shu ish davomida aynan shu to'lov uchun javob bergan
    # bo'lsa — o'shani ishlatamiz.
    sess = SESSION_OVERRIDES.get((account_str, purpose_signature(text)))
    if sess:
        return sess, "high"

    if not faqat_ozim:
        for rx, cat in PURPOSE_FIRST_RULES:
            if rx.search(text):
                return cat, "high"

    # Foydalanuvchi o'zi qo'shgan matn qoidalari — hisob raqamdan ustun
    # turadi, chunki ular aynan shunday "bitta hisob raqam, ko'p maqsad"
    # holatlarini qo'lda ajratish uchun kiritilgan.
    text_low = text.lower()
    # Yo'nalishga bog'langan qoida umumiysidan ustun turadi: u aniqroq
    # holat uchun yozilgan.
    if yonalish:
        for (phrase, yon), cat in KNOWN_PURPOSE_TEXT_DIR.items():
            if yon == yonalish and phrase.lower() in text_low:
                return cat, "high"
    for phrase, cat in KNOWN_PURPOSE_TEXT.items():
        if phrase.lower() in text_low:
            return cat, "high"

    # Ko'p maqsadli kontragent (G'aznachilik) — hisob/nom bo'yicha
    # o'rganilgan guruh QO'LLANMAYDI, chunki u har safar boshqa maqsadda
    # bo'lishi mumkin. Yuqoridagi maqsad qoidalari ishlamagan bo'lsa,
    # pastda "?" qaytadi va ilova foydalanuvchidan so'raydi.
    if is_always_ask(name, account):
        return "?", "review"

    account = account_str
    inn = str(inn).strip() if inn else ""
    name_key = name.strip()

    if yonalish:
        for kalit in (account, inn):
            if kalit and (kalit, yonalish) in KNOWN_VENDOR_INN_DIR:
                return KNOWN_VENDOR_INN_DIR[(kalit, yonalish)], "high"
        if (name_key, yonalish) in KNOWN_VENDOR_NAME_DIR:
            return KNOWN_VENDOR_NAME_DIR[(name_key, yonalish)], "high"

    for kalit in (account, inn):
        if not kalit or kalit not in KNOWN_VENDOR_INN:
            continue
        if faqat_ozim and kalit not in LEARNED_VENDOR_KEYS:
            continue  # kodga yozilgan kontragent — bu rejimda hisobga olinmaydi
        return KNOWN_VENDOR_INN[kalit], "high"

    if name_key in KNOWN_VENDOR_NAME:
        return KNOWN_VENDOR_NAME[name_key], "high"

    if not faqat_ozim:
        for rx, cat, conf in TEXT_RULES:
            if rx.search(text) or rx.search(name):
                return cat, conf

        if op == 1 and GOODS_PURCHASE_HINT.search(text):
            return "МЕБ", "guess"

    return "?", "review"


def propose_category_for_group(op_values, names, texts, inns, accounts=None):
    """Given all rows sharing one raw account number, classify each row on
    its own merits, then only collapse to a single group-level category if
    every row agrees. A raw account is NOT trusted as a category by itself
    (e.g. a generic treasury account can carry profit tax, VAT AND utility
    payments; conversely two different accounts that both happen to mention
    HUMO/SmartVista text should NOT be assumed identical without agreeing)."""
    if accounts is None:
        accounts = [None] * len(op_values)
    per_row = [classify_row(op, n, t, i, a) for op, n, t, i, a in zip(op_values, names, texts, inns, accounts)]
    cats = {c for c, _ in per_row}
    if len(cats) == 1:
        cat, conf = per_row[0]
        return cat, conf, per_row
    return "АРАЛАШ", "mixed", per_row


# Ustunlarni JOYLASHUVI bo'yicha emas, SARLAVHA NOMI bo'yicha aniqlaymiz —
# chunki banklar bir necha xil ko'rinishda hisobot beradi:
#   A) "CBreport63": Дата | Счет | ИНН | Наименование | № док-та | Оп | МФО | ...
#   B) "Сведения о работе счета": ИНН va Наименование ustunlari YO'Q, hisob
#      raqam katagida nom yangi qatordan keyin turadi, summalar esa matn
#      ko'rinishida ("1 170 682,76")
#   C) "Sheet1": Дата проводки | Номер документа | МФО корресп. | Счет
#      корреспондента | Наименование корресп. | ИНН | Детали | Дебет | Кредит
#   D) "Справка о работе счета": ustunlar IKKI TOMONGA bo'lingan — mijoz
#      (bizning hisobimiz) va korrespondent (kontragent). Bizga faqat
#      korrespondent kerak, shuning uchun "ИНН клиента" kabi nomlar
#      ataylab ro'yxatga kiritilmagan: ular hech qanday maydonga
#      tushmaydi va e'tiborsiz qoladi. Bank interfeysi tiliga qarab shu
#      formatning o'zbekcha kirill ko'rinishi ham keladi ("Сана",
#      "Номи корреспондент", "СТИР", "Ҳисоб", "Тўлов мақсади") —
#      ustun nomlari boshqa, tuzilishi bir xil.
# Shu tarzda kelajakda yana bir variant chiqsa ham kod ishlayveradi.
_COLUMN_ALIASES = {
    "date": ("дата проводки", "дата", "сана"),
    "account": ("счет корреспондента", "счет корресп.", "счет",
                "ҳисоб корреспондент"),
    "inn": ("инн корреспондента", "инн", "стир корреспондент"),
    "name": ("наименование корреспондента", "наименование корресп.",
             "наименование корресп", "наименование", "номи корреспондент"),
    "doc_no": ("№ док-та", "номер документа", "№ док", "№ ҳуж-та"),
    "op": ("во", "оп"),
    "mfo": ("мфо корреспондента", "мфо корресп.", "мфо корресп", "мфо",
            "мфо корреспондент"),
    "debit": ("сумма дебета", "оборот дебет", "дебет", "дебет суммаси"),
    "credit": ("сумма кредита", "оборот кредит", "кредит", "кредит суммаси"),
    "purpose": ("назначение платежа", "детали", "тўлов мақсади"),
}


def _norm_header(value):
    """Sarlavha matnini solishtirishga tayyorlaydi: kichik harf, yangi
    qator va ortiqcha bo'shliqlar bitta bo'shliqqa aylanadi.

    "ё" harfi "е" ga aylantiriladi: bir bank "Счет", boshqasi "Счёт" deb
    yozadi va ular boshqa-boshqa belgi — busiz ustun umuman
    tanilmay qolardi."""
    return " ".join(str(value or "").split()).strip().lower().replace("ё", "е")


def _detect_header(ws, max_scan=30):
    """Sarlavha qatorini va ustunlar xaritasini topadi. Kamida "дата",
    "счет" va "назначение платежа"/"детали" bo'lishi shart, aks holda bu
    qator sarlavha emas."""
    for row in ws.iter_rows(min_row=1, max_row=min(max_scan, ws.max_row)):
        found = {}
        for cell in row:
            h = _norm_header(cell.value)
            if not h:
                continue
            for field, aliases in _COLUMN_ALIASES.items():
                if field in found:
                    continue
                if h in aliases:
                    found[field] = cell.column - 1  # 0-asosli indeks
                    break
        if "date" in found and "account" in found and "purpose" in found:
            return row[0].row, found
    return None, {}


def _parse_amount(value):
    """Summani songa aylantiradi. Ba'zi hisobotlarda summa matn bo'lib
    keladi: "1 170 682,76" yoki "0,00" — bo'sh joylar ajratuvchi, vergul
    esa kasr belgisi."""
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return value or None
    t = str(value).replace("\xa0", " ").replace(" ", "").replace(",", ".")
    if not t or t in ("-", "."):
        return None
    try:
        num = float(t)
    except ValueError:
        return None
    return num or None


def _parse_op(value):
    """Operatsiya turi ba'zi hisobotlarda matn ("4"), ba'zilarida son
    bo'lib keladi, C formatida esa bu ustun umuman yo'q."""
    if value is None or value == "":
        return None
    try:
        return int(str(value).strip())
    except ValueError:
        return None


def _split_account_cell(value):
    """B formatida hisob raqam katagi "23508000007102718500\\nYANGI..."
    ko'rinishida keladi — raqam birinchi qatorda, kontragent nomi esa
    keyingi qatorlarda. Ikkalasini ajratib qaytaradi."""
    parts = [p.strip() for p in str(value or "").splitlines() if p.strip()]
    if not parts:
        return "", ""
    return parts[0], " ".join(parts[1:])


_XLS_CONVERT_PS = (
    "$e = New-Object -ComObject Excel.Application; "
    "$e.Visible = $false; $e.DisplayAlerts = $false; "
    "$wb = $e.Workbooks.Open('{kirish}'); "
    "$wb.SaveAs('{chiqish}', 51); "
    "$wb.Close($false); $e.Quit(); "
    "[System.Runtime.InteropServices.Marshal]::ReleaseComObject($e) | Out-Null"
)


_HTML_TOKEN = re.compile(r"<!--.*?-->|<(/?)([a-zA-Z][a-zA-Z0-9]*)([^>]*)>|([^<]+)", re.S)
_HTML_ENTITIES = {"nbsp": " ", "amp": "&", "lt": "<", "gt": ">", "quot": '"', "apos": "'", "laquo": "«", "raquo": "»"}
_HTML_ENTITY = re.compile(r"&(#x[0-9a-fA-F]+|#\d+|[a-zA-Z]+);")
_INT_RX = re.compile(r"^-?\d{1,15}$")
_MONEY_RX = re.compile(r"^-?\d{1,3}(?:[ \u00a0]\d{3})*(?:[.,]\d+)?$|^-?\d+[.,]\d+$")


def _html_unescape(t):
    def one(m):
        k = m.group(1)
        try:
            if k.startswith("#x"):
                return chr(int(k[2:], 16))
            if k.startswith("#"):
                return chr(int(k[1:]))
        except (ValueError, OverflowError):
            return m.group(0)
        return _HTML_ENTITIES.get(k.lower(), m.group(0))
    return _HTML_ENTITY.sub(one, t)


def _html_cell_value(text, as_text):
    """Katak qiymati - Excel HTML ni ochgandagidek: class="txt" matn bo'lib
    qoladi (hisob raqam, MFO "00083"), qolgan son ko'rinishidagilar son."""
    t = _html_unescape(text).replace("\u00a0", " ")
    t = "\n".join(" ".join(line.split()) for line in t.split("\n")).strip()
    if not t:
        return None
    if as_text:
        return t
    if _INT_RX.match(t):
        return int(t)
    if _MONEY_RX.match(t):
        try:
            return float(t.replace(" ", "").replace("\u00a0", "").replace(",", "."))
        except ValueError:
            pass
    return t


def html_table_rows(raw):
    """HTML jadval(lar)ini qatorlarga ajratadi. Bank fayli yopilmagan
    <td>/<th> teglarini ham ishlatadi (<th>Счет<th>ИНН) - yangi katak
    oldingisini yopadi. html.parser ishlatilmaydi: u .exe ichida yo'q."""
    head = raw[:4096].decode("ascii", "ignore").lower()
    m = re.search(r"charset=[\"']?([\w-]+)", head)
    encs = ([m.group(1)] if m else []) + ["utf-8", "cp1251"]
    text = None
    for enc in encs:
        try:
            text = raw.decode(enc)
            break
        except (LookupError, UnicodeDecodeError):
            continue
    if text is None:
        text = raw.decode("cp1251", "replace")

    rows, row, cell = [], None, None
    skip = 0

    def close_cell():
        nonlocal cell
        if cell is not None and row is not None:
            row.append(_html_cell_value("".join(cell["t"]), cell["txt"]))
            row.extend([None] * (cell["span"] - 1))
        cell = None

    def close_row():
        nonlocal row
        close_cell()
        if row is not None and any(v is not None for v in row):
            rows.append(row)
        row = None

    for mt in _HTML_TOKEN.finditer(text):
        end, tag, attrs, data = mt.group(1), (mt.group(2) or "").lower(), mt.group(3) or "", mt.group(4)
        if data is not None:
            if cell is not None and not skip:
                cell["t"].append(data.replace("\r", "").replace("\n", " "))
            continue
        if tag in ("script", "style", "head", "title"):
            skip = max(0, skip - 1) if end else skip + 1
        elif skip:
            continue
        elif tag == "tr":
            close_row()
            if not end:
                row = []
        elif tag in ("td", "th"):
            close_cell()
            if not end:
                if row is None:
                    row = []
                sp = re.search(r"colspan\s*=\s*[\"']?(\d+)", attrs, re.I)
                cell = {"t": [], "span": max(1, min(int(sp.group(1)), 50)) if sp else 1,
                        "txt": bool(re.search(r"class\s*=\s*[\"']?[^>]*\btxt\b", attrs, re.I))}
        elif tag == "br" and cell is not None:
            cell["t"].append("\n")
        elif tag in ("table", "tbody", "thead") and end:
            close_row()
    close_row()
    return rows


def _html_to_xlsx(path):
    """HTML ni (bank ".xls" deb beradigan) .xlsx ga o'giradi - Excel'siz."""
    with open(path, "rb") as f:
        rows = html_table_rows(f.read())
    if not rows:
        raise ValueError("Fayl ichida jadval topilmadi.")
    wb = openpyxl.Workbook()
    ws = wb.active
    # Varaq nomi - Excel o'girgandagidek fayl nomi (natijada 2-varaq shu nomda)
    nom = re.sub(r"[\\/*?:\[\]]", "", os.path.splitext(os.path.basename(path))[0]).strip("' ")[:31]
    ws.title = nom or "Sheet1"
    for r in rows:
        ws.append(r)
    ish = tempfile.mkdtemp(prefix="soddahisobot_")
    chiqish = os.path.join(ish, os.path.splitext(os.path.basename(path))[0] + ".xlsx")
    wb.save(chiqish)
    return chiqish


def _is_html_file(path):
    try:
        with open(path, "rb") as f:
            bosh = f.read(1024).lstrip(b"\xef\xbb\xbf \t\r\n").lower()
    except OSError:
        return False
    return bosh.startswith(b"<") and (b"<html" in bosh or b"<table" in bosh or b"<!doctype" in bosh)


def ensure_xlsx(path):
    """Eski .xls fayl bo'lsa, uni vaqtinchalik .xlsx ga o'giradi.

    openpyxl faqat .xlsx ni o'qiydi, bank esa ba'zan eski .xls beradi va
    ilova "does not support the old .xls file format" deb to'xtab qolardi.

    O'girish Excel'ning o'zi orqali qilinadi: ilovaga yangi kutubxona
    qo'shilsa, uni .exe ichiga kiritish va .exe ni hamma foydalanuvchiga
    qaytadan tarqatish kerak bo'lardi. Excel esa bu kompyuterlarda
    allaqachon bor — hisobot baribir unda ochiladi.

    Bank (IABS/Client-Bank) ko'pincha ".xls" nomli HTML beradi - uni
    Excel'siz, o'zimiz o'qiymiz: boshliq kompyuterida Excel o'girishi
    ishlamay, fayl "tanilmay" qolardi (2.2.3).

    Qaytaradi: o'qish uchun yo'l. .xlsx bo'lsa — o'zgarishsiz."""
    if _is_html_file(path):
        return _html_to_xlsx(path)
    if not path.lower().endswith(".xls"):
        return path

    ish = tempfile.mkdtemp(prefix="soddahisobot_")
    chiqish = os.path.join(ish, os.path.basename(path)[:-4] + ".xlsx")
    buyruq = _XLS_CONVERT_PS.format(
        kirish=os.path.abspath(path).replace("'", "''"),
        chiqish=chiqish.replace("'", "''"),
    )
    try:
        subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-Command", buyruq],
            check=True, timeout=120,
            stdout=subprocess.DEVNULL, stderr=subprocess.PIPE,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
    except Exception as e:
        raise ValueError(
            "Eski .xls formatidagi faylni o'girib bo'lmadi. Faylni Excel'da "
            "ochib, \"Сохранить как\" orqali .xlsx sifatida saqlang va qaytadan "
            f"tanlang. ({type(e).__name__})"
        ) from e

    if not os.path.exists(chiqish):
        raise ValueError(
            "Eski .xls fayl .xlsx ga o'girilmadi. Faylni Excel'da ochib, "
            ".xlsx sifatida saqlang va qaytadan tanlang."
        )
    return chiqish


def load_raw_rows(path, sheet_name=None):
    """Xom hisobotni o'qiydi. Qaytaradi: (wb, ws, layout, rows), bunda
    `layout` — {"header_row": int, "cols": {maydon: ustun indeksi}}.
    Chaqiruvchi hisobotga yozishda `layout["cols"]["account"]` dan
    foydalanishi kerak, chunki hisob raqam ustuni har formatda har xil
    joyda turadi."""
    wb = openpyxl.load_workbook(ensure_xlsx(path), data_only=True)
    ws = wb[sheet_name] if sheet_name else wb[wb.sheetnames[0]]

    header_row, cols = _detect_header(ws)
    if header_row is None:
        raise ValueError(
            "Bu faylda tanish sarlavha qatori topilmadi "
            "(kamida \"Дата\", \"Счет\" va \"Назначение платежа\" ustunlari kerak)."
        )

    def get(vals, field):
        idx = cols.get(field)
        if idx is None or idx >= len(vals):
            return None
        return vals[idx]

    rows = []
    for row in ws.iter_rows(min_row=header_row + 1, max_row=ws.max_row):
        vals = [c.value for c in row]
        if not vals or all(v is None for v in vals):
            continue
        date_val = get(vals, "date")
        if is_total_row(date_val):
            continue
        # Sana ustuni toza sana bo'lmasa — bu jadval oxiridagi yakuniy
        # ma'lumot qatori (jami aylanma, chiqish qoldig'i, bank manzili).
        if not looks_like_date(date_val):
            continue

        account_raw = get(vals, "account")
        account, embedded_name = _split_account_cell(account_raw)
        name = get(vals, "name")
        if not name and embedded_name:
            # B formatida alohida "Наименование" ustuni yo'q — nom hisob
            # raqam katagining ichida keladi.
            name = embedded_name

        rows.append({
            "excel_row": row[0].row,
            "date": get(vals, "date"),
            "account": account,
            "inn": get(vals, "inn"),
            "name": name,
            "doc_no": get(vals, "doc_no"),
            "op": _parse_op(get(vals, "op")),
            "mfo": get(vals, "mfo"),
            "debit": _parse_amount(get(vals, "debit")),
            "credit": _parse_amount(get(vals, "credit")),
            # To'lov izohidagi yangi qator belgilari bo'shliqqa
            # aylantiriladi. "Справка о работе счета" formatida matn
            # katak ichida bir necha qatorga bo'lingan holda keladi va
            # "Возмещение клиенту по покупкам" va "ТСП" so'zlari
            # orasiga qator uzilishi tushib qolgani uchun Терминал
            # qoidasi mos kelmay qolardi — yirik summalar "?" ga
            # tushib ketardi.
            "purpose": " ".join((get(vals, "purpose") or "").split()),
        })
    return wb, ws, {"header_row": header_row, "cols": cols}, rows


def build_account_proposals(rows):
    groups = defaultdict(lambda: {"op": [], "names": [], "texts": [], "inns": [], "rows": [],
                                   "debit": Decimal(0), "credit": Decimal(0)})
    for r in rows:
        g = groups[r["account"]]
        g["op"].append(r["op"])
        g["names"].append(str(r["name"] or ""))
        g["texts"].append(str(r["purpose"] or ""))
        g["inns"].append(r["inn"])
        g["rows"].append(r)
        if r["debit"]:
            g["debit"] += Decimal(str(r["debit"]))
        if r["credit"]:
            g["credit"] += Decimal(str(r["credit"]))

    proposals = {}
    for acct, g in groups.items():
        accounts = [acct] * len(g["rows"])
        cat, conf, per_row = propose_category_for_group(g["op"], g["names"], g["texts"], g["inns"], accounts)
        proposals[acct] = {
            "category": cat,
            "confidence": conf,
            "count": len(g["rows"]),
            "debit_sum": g["debit"],
            "credit_sum": g["credit"],
            "sample_text": g["texts"][0][:160],
            "sample_name": g["names"][0],
            "rows": g["rows"],
            "per_row": per_row,
        }
    return proposals


if __name__ == "__main__":
    path = sys.argv[1]
    wb, ws, header_row, rows = load_raw_rows(path)
    proposals = build_account_proposals(rows)

    out_lines = []
    out_lines.append(f"Jami qatorlar: {len(rows)}, noyob hisob raqamlari: {len(proposals)}\n")
    for acct, info in sorted(proposals.items(), key=lambda kv: -kv[1]["count"]):
        if info["confidence"] == "mixed":
            out_lines.append(
                f"{acct!r:28} -> ARALASH (har qator alohida)  soni={info['count']:4}\n"
            )
            for r, (cat, conf) in zip(info["rows"], info["per_row"]):
                out_lines.append(
                    f"    [{cat:16}/{conf:6}] {r['name']} | {str(r['purpose'])[:140]}\n"
                )
            continue
        flag = "" if info["confidence"] == "high" else f"  <== TEKSHIRISH ({info['confidence']})"
        out_lines.append(
            f"{acct!r:28} -> {info['category']:20} soni={info['count']:4} "
            f"debet={info['debit_sum']:>15} kredit={info['credit_sum']:>15}{flag}\n"
            f"    namuna: {info['sample_name']} | {info['sample_text']}\n"
        )
    with open("account_proposals.txt", "w", encoding="utf-8") as f:
        f.writelines(out_lines)
    print("done, rows=", len(rows), "accounts=", len(proposals))


# ======================================================================
# 2-QISM — HISOBOT YASASH
# ======================================================================

# Yakuniy hisobotdagi yagona shrift (buyurtmachi so'roviga ko'ra)
FONT_NAME = "Times New Roman"
FONT_SIZE = 14


def _writable_cell(ws, row, col):
    """Katakka yozish uchun tayyor obyekt qaytaradi. Ba'zi hisobotlarda
    ustunlar birlashtirilgan (merged) bo'ladi — bunday katakka to'g'ridan
    to'g'ri yozib bo'lmaydi, faqat birlashma boshidagi katakka yoziladi.
    Shu boshlang'ich katakni topib beradi."""
    cell = ws.cell(row=row, column=col)
    if not isinstance(cell, MergedCell):
        return cell
    for rng in ws.merged_cells.ranges:
        if (rng.min_row <= row <= rng.max_row) and (rng.min_col <= col <= rng.max_col):
            return ws.cell(row=rng.min_row, column=rng.min_col)
    return None


# Boshlang'ich/yakuniy qoldiq turli banklarda turlicha yoziladi:
#   "Остаток на начало периода: 511 258,38"   -> raqam matn ichida
#   "Входящий остаток на 01.08.2026:  135 491 685,73"
#   "Исходящий остаток на 31.08.2026"          -> raqam ALOHIDA katakda
# Ustun/qator raqamiga tayanib bo'lmaydi (A4, B4, F4, H25 ... hammasi
# uchraydi), shuning uchun kalit ibora bo'yicha qatorni topamiz.
_OPEN_BALANCE_WORDS = ("входящий остаток", "остаток на начало", "начальный остаток",
                       "давр бошидаги қолдиқ")
_CLOSE_BALANCE_WORDS = ("исходящий остаток", "остаток на конец", "конечный остаток",
                        "давр охиридаги қолдиқ")


def _parse_balance_number(text):
    """Matndan pul summasini ajratadi. FAQAT ikki nuqta (":") dan keyingi
    qismga qaraydi — aks holda "Исходящий остаток на 31.08.2026" dagi
    sana raqam deb o'qilib ketardi."""
    if ":" not in text:
        return None
    tail = text.rsplit(":", 1)[1]
    m = re.search(r"([\d\s\xa0]+[.,]\d+|[\d\s\xa0]+)\s*$", tail)
    if not m:
        return None
    # Bo'shliqlar (oddiy va uzilmas) — ming ajratuvchi, olib tashlanadi.
    # Vergul — kasr ajratuvchi, NUQTAGA aylantiriladi (o'chirilmaydi:
    # "511 258,38" -> 511258.38, aks holda 51125838 bo'lib ketardi).
    t = m.group(1).replace(" ", "").replace("\xa0", "").replace(",", ".")
    try:
        return Decimal(t)
    except Exception:
        return None


def _find_balances(ws):
    """Varaqdan boshlang'ich va yakuniy qoldiqni topadi.

    Kataklar birma-bir tekshiriladi (qator emas), chunki ba'zi
    hisobotlarda IKKALA qoldiq ham BITTA qatorda turadi (A4 da
    boshlang'ich, F4 da yakuniy), boshqalarida esa har xil qatorda va
    raqam butunlay boshqa katakda bo'ladi (A25 da yozuv, H25 da son)."""
    open_bal = close_bal = None
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
        for cell in row:
            low = str(cell.value or "").strip().lower()
            if not low:
                continue
            if any(w in low for w in _OPEN_BALANCE_WORDS):
                kind = "open"
            elif any(w in low for w in _CLOSE_BALANCE_WORDS):
                kind = "close"
            else:
                continue
            if (kind == "open" and open_bal is not None) or \
               (kind == "close" and close_bal is not None):
                continue

            # 1) Raqam yozuvning o'zida bo'lishi mumkin
            value = _parse_balance_number(str(cell.value))
            # 2) Bo'lmasa — shu qatordagi alohida raqamli katakdan
            if value is None:
                for other in row:
                    if isinstance(other.value, (int, float)):
                        value = Decimal(str(other.value))
                        break
            if value is None:
                continue

            if kind == "open":
                open_bal = value
            else:
                close_bal = value
        if open_bal is not None and close_bal is not None:
            break
    return open_bal, close_bal


def build_simplified_report(src_path, out_path):
    wb, ws, layout, rows = load_raw_rows(src_path)

    # classify every row independently, then sanity-check by raw account:
    # if one raw account ends up split across >1 category, surface it so it
    # gets extra attention (still applied per-row, never silently collapsed).
    by_account = defaultdict(set)

    # Guruh nomi shu foydalanuvchidagi atamaga almashtiriladi. Bitta
    # joyda qilinadi, shunda xom varaqdagi ustun ham, Лист1 ham, jami
    # ham bir xil nomni ko'radi.
    nomlar = load_name_map()

    results = []  # (row_dict, category, confidence)
    for r in rows:
        cat, conf = classify_row(
            r["op"], r["name"], r["purpose"], r["inn"], r["account"],
            row_direction(r["debit"], r["credit"]),
        )
        cat = nomlar.get(cat, cat)
        results.append((r, cat, conf))
        by_account[r["account"]].add(cat)

    # 1) "Счет" ustunini joyida kategoriya nomiga almashtiramiz (qolgan
    #    bezaklar, ustun kengliklari o'zgarmaydi). Ustun indeksi hisobot formatiga qarab har xil (B, D ...), shuning
    # uchun load_raw_rows aniqlagan joylashuvdan olamiz.
    account_col = layout["cols"]["account"] + 1  # openpyxl 1-asosli
    for r, cat, conf in results:
        cell = _writable_cell(ws, r["excel_row"], account_col)
        if cell is None:
            continue
        cell.value = cat
        cell.number_format = "@"

    # 2) build the Лист1 summary sheet (SUMIF over the report sheet)
    report_sheet_name = ws.title
    totals = defaultdict(lambda: [Decimal(0), Decimal(0)])  # cat -> [debet, kredit]
    for r, cat, conf in results:
        if r["debit"]:
            totals[cat][0] += Decimal(str(r["debit"]))
        if r["credit"]:
            totals[cat][1] += Decimal(str(r["credit"]))

    open_bal, close_bal = _find_balances(ws)

    if "Лист1" in wb.sheetnames:
        del wb["Лист1"]
    summary = wb.create_sheet("Лист1", 0)

    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    acc_fmt = '_-* #,##0_-;\\-* #,##0_-;_-* "-"??_-;_-@_-'

    # Faqat fayl nomi ("АБ М Август"), to'liq yo'l emas. os.path.basename
    # ishlatiladi, chunki fayl tanlash oynasi yo'lni "/" bilan qaytaradi va
    # oldingi "\\" bo'yicha ajratish butun yo'lni sarlavhaga yozib qo'yardi.
    title = os.path.splitext(os.path.basename(src_path))[0]
    sarlavhali = titled_style()
    if sarlavhali:
        # Fayl nomi yuqorida alohida sarlavha bo'lib turadi, 2-qatorda
        # esa uning o'rniga qoldiq yozuvi qoladi. Pastki chiziqlar
        # bo'shliqqa aylantiriladi: fayl nomida ular ko'p uchraydi.
        summary.merge_cells(start_row=1, start_column=1, end_row=1, end_column=4)
        bosh = summary.cell(row=1, column=1, value=title.replace("_", " "))
        bosh.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        summary.row_dimensions[1].height = 40
        summary["B2"] = "Колдик"
    else:
        summary["B2"] = title
    summary["B2"].font = Font(bold=True)
    if open_bal is not None:
        summary["C2"] = float(open_bal)
    if close_bal is not None:
        summary["D2"] = float(close_bal)
    for col in ("B2", "C2", "D2"):
        summary[col].font = Font(bold=True)
        summary[col].border = border
        summary[col].number_format = acc_fmt

    qizil = Font(color="FFFF0000", bold=True)

    headers = ["№", "Названия строк", " Дебет", " Кредит"]
    for i, h in enumerate(headers):
        c = summary.cell(row=3, column=1 + i, value=h)
        c.border = border
        c.number_format = acc_fmt
        if sarlavhali:
            c.font = qizil

    row_i = 4
    total_debet = Decimal(0)
    total_kredit = Decimal(0)
    review_rows = []
    for idx, (cat, (deb, kred)) in enumerate(sorted(totals.items()), start=1):
        summary.cell(row=row_i, column=1, value=idx).border = border
        c_name = summary.cell(row=row_i, column=2, value=cat)
        c_name.border = border
        if deb:
            c = summary.cell(row=row_i, column=3, value=float(deb))
            c.border = border
            c.number_format = acc_fmt
            total_debet += deb
        if kred:
            c = summary.cell(row=row_i, column=4, value=float(kred))
            c.border = border
            c.number_format = acc_fmt
            total_kredit += kred
        for col in (1, 2, 3, 4):
            summary.cell(row=row_i, column=col).number_format = acc_fmt
            summary.cell(row=row_i, column=col).border = border
        row_i += 1

    summary.cell(row=row_i, column=2, value="Общий итог").border = border
    c = summary.cell(row=row_i, column=3, value=float(total_debet))
    c.number_format = acc_fmt
    c.border = border
    c = summary.cell(row=row_i, column=4, value=float(total_kredit))
    c.number_format = acc_fmt
    c.border = border
    summary.cell(row=row_i, column=1).border = border
    if sarlavhali:
        for col in (1, 2, 3, 4):
            summary.cell(row=row_i, column=col).font = qizil

    summary.column_dimensions["A"].width = 5.2
    summary.column_dimensions["B"].width = 24
    summary.column_dimensions["C"].width = 16
    summary.column_dimensions["D"].width = 16

    # Butun kitob bo'ylab yagona shrift: Times New Roman, 14. Qalinlik
    # (bold) va boshqa bezaklar qayerda bo'lsa, o'sha holicha saqlanadi.
    for sheet in wb.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                old = cell.font
                cell.font = Font(
                    name=FONT_NAME,
                    size=FONT_SIZE,
                    bold=old.bold,
                    italic=old.italic,
                    underline=old.underline,
                    color=old.color,
                )

    if sarlavhali:
        # Yuqoridagi tenglashtirish hamma katakni 14 o'lchamga keltiradi;
        # sarlavha undan keyin kattalashtiriladi, aks holda yo'qolardi.
        summary["A1"].font = Font(name=FONT_NAME, size=16, bold=True)

    wb.save(out_path)

    # 3) Past ishonch bilan belgilangan qatorlar soni. Avval bu ro'yxat
    #    alohida "_tekshirish.txt" fayliga yozilardi, lekin buyurtmachiga
    #    faqat Excel fayl kerak — endi bu son ilova jurnalida ko'rsatiladi.
    review_count = sum(1 for _r, _cat, conf in results if conf in ("guess", "review"))

    return {
        "out_path": out_path,
        "review_count": review_count,
        "total_debet": total_debet,
        "total_kredit": total_kredit,
        "balans_farqi": _balance_mismatch(open_bal, close_bal, total_debet, total_kredit),
    }


def _balance_mismatch(open_bal, close_bal, total_debet, total_kredit):
    """Bank o'zi yozgan qoldiqlar bilan hisoblangan aylanma mos keladimi.

    Tenglik: boshlang'ich + kredit - debet = yakuniy. Bu bankning o'z
    raqamlari bilan bizning o'qiganimizni solishtiradi, ya'ni qatorlar
    tushib qolgani yoki ikki marta hisoblangani darhol bilinadi.

    Nima uchun kerak: bunday xato hisobotni jimgina buzadi — fayl
    chiroyli ko'rinadi, jami chiqadi, lekin raqamlar noto'g'ri bo'ladi.
    Ko'z bilan sezib bo'lmaydi.

    Qaytaradi: farq (Decimal) yoki None — qoldiqlar topilmasa tekshirib
    bo'lmaydi, bu xato emas (ba'zi hisobotlarda ular umuman yo'q)."""
    if open_bal is None or close_bal is None:
        return None
    kutilgan = Decimal(str(open_bal)) + total_kredit - total_debet
    farq = kutilgan - Decimal(str(close_bal))
    # Tiyin darajasidagi yaxlitlash farqi xato emas.
    return farq if abs(farq) > Decimal("0.01") else Decimal(0)



# ======================================================================
# 3-QISM — INTERFEYS
# ======================================================================

def find_excel_exe():
    """Locate the real Excel.exe via the Windows "App Paths" registry,
    bypassing whatever program .xlsx happens to be (mis)associated with on
    this machine (a common support issue: someone once chose "Open with ->
    Notepad -> always use this app" for .xlsx files)."""
    if winreg is None:
        return None
    candidates = [
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\EXCEL.EXE"),
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\EXCEL.EXE"),
        (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\EXCEL.EXE"),
    ]
    for hive, subkey in candidates:
        try:
            with winreg.OpenKey(hive, subkey) as key:
                path, _ = winreg.QueryValueEx(key, "")
        except OSError:
            continue
        if path and os.path.exists(path):
            return path
    return None


def resource_base_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


APP_DIR = resource_base_dir()


class FileRow:
    def __init__(self, path):
        self.path = path
        self.status = "Kutmoqda"
        self.error = None
        self.out_path = None


# ------------------------------------------------------------ AI yordamchi
# Nomlanmagan kontragent uchun guruh TAKLIF qiladi — hech qachon o'zi yozib
# qo'ymaydi. Sabab: noto'g'ri kategoriya hisobotni jimgina buzadi va uni
# keyin topish qiyin, shuning uchun oxirgi qaror foydalanuvchida qoladi.
#
# Ayniqsa "Сведения о работе счета" formatida kerak: unda ИНН ustuni yo'q,
# ya'ni ИНН ga tayangan qoidalar umuman ishlamaydi va yirik summalar "?"
# bo'lib qoladi. Bunday hisobotda yagona ishonchli belgi — to'lov maqsadi
# matni, uni esa qoida bilan emas, ma'no bilan tushunish kerak.
#
# Tarmoq yo'q yoki javob buzuq bo'lsa — bo'sh natija qaytadi va ilova
# avvalgidek (qo'lda kiritish bilan) ishlayveradi.
#
# So'rov to'g'ridan-to'g'ri Gemini'ga emas, o'zimizning Worker orqali
# ketadi. Sabab: to'g'ridan-to'g'ri murojaat uchun API kalit har bir
# foydalanuvchining kompyuterida turishi kerak edi, uni u yerdan ko'chirib
# olish esa oson — hisob bitta bo'lgani uchun kalit tarqalsa so'rovlar
# bizning nomimizdan ketaverardi. Endi kalit faqat serverda, ilovada
# umuman yo'q. Worker'da kunlik chegara ham bor.
AI_MODEL = "gemini-3-flash-preview"
_AI_URL = "https://soddahisobot-telemetry.tasks-bot.workers.dev/ai"
_AI_RULES_FILE = "ai_qollanma.txt"

# Worker qaytaradigan xato kodlari uchun foydalanuvchiga ko'rsatiladigan
# matn. Avval har qanday nosozlik "AI mos taklif topa olmadi" bo'lib
# ko'rinardi — ya'ni AI buzilganini bilib bo'lmasdi.
_AI_ERRORS = {
    "limit": "AI kunlik chegaraga yetdi — ertaga qayta ishlaydi.",
    "config": "AI sozlamasida xato (server tomonida kalit yo'q).",
    "upstream": "AI xizmati javob bermadi.",
    "empty": "AI bo'sh javob qaytardi.",
    "bad_request": "AI so'rovi qabul qilinmadi.",
}

# Foydalanuvchi ai_qollanma.txt ni birinchi marta ochganda nima yozishni
# bilishi uchun namuna. Fayl yo'q bo'lsa shu mazmun bilan yaratiladi.
_AI_RULES_TEMPLATE = """\
# AI uchun ko'rsatmalar. Har qatorga bitta qoida, oddiy tilda yozing.
# Ilova har ochilganda shu fayl o'qiladi — qayta o'rnatish shart emas.
# "#" bilan boshlangan qatorlar e'tiborga olinmaydi.
#
# Namunalar (kerak bo'lsa o'chirib, o'zingiznikini yozing):
# ORZUIM MCHJ har doim хизмат guruhiga kirsin, tovar emas.
# Matnda "аренда банкомата" bo'lsa - ижара.
# Nomida "МИБ" bo'lsa hech qachon МЕБ dema.
"""


SETTINGS_FILE = "sozlamalar.json"
AI_ENABLED_KEY = "ai_yoqilgan"


def _settings_path():
    return os.path.join(_data_dir(), SETTINGS_FILE)


def get_setting(nom, standart=None):
    try:
        with open(_settings_path(), encoding="utf-8") as f:
            return json.load(f).get(nom, standart)
    except (OSError, ValueError):
        return standart


def set_setting(nom, qiymat):
    try:
        with open(_settings_path(), encoding="utf-8") as f:
            sozlamalar = json.load(f)
        if not isinstance(sozlamalar, dict):
            sozlamalar = {}
    except (OSError, ValueError):
        sozlamalar = {}
    sozlamalar[nom] = qiymat
    try:
        with open(_settings_path(), "w", encoding="utf-8") as f:
            json.dump(sozlamalar, f, ensure_ascii=False, indent=2)
    except OSError:
        pass


SCRATCH_KEY = "faqat_ozim"
STYLE_KEY = "sarlavhali_korinish"
NAMES_FILE = "nomlar.json"


def titled_style():
    """Hisobot yuqorisida sarlavha, "Колдик" qatori va qizil yig'indi.

    Foydalanuvchilar hisobotni turlicha ko'rishni xohlashadi, shuning
    uchun bu tanlov — standart holatda o'chiq, ya'ni mavjud
    foydalanuvchilarning odatlangan ko'rinishi o'zgarmaydi."""
    return bool(get_setting(STYLE_KEY, False))


_SCRATCH_CACHE = None


def set_scratch_mode(yoqilgan):
    global _SCRATCH_CACHE
    set_setting(SCRATCH_KEY, bool(yoqilgan))
    _SCRATCH_CACHE = bool(yoqilgan)


def scratch_mode():
    """"Faqat o'zim kiritgan guruhlar" rejimi.

    Kodga yozilgan qoidalar va kontragentlar bir kishining ish uslubiga
    moslangan. Boshqa foydalanuvchi butunlay boshqa nomlar bilan
    ishlashi mumkin — bu rejimda tayyor qoidalar umuman qo'llanmaydi,
    har bir kontragent "?" bo'lib so'raladi va u o'z ro'yxatini noldan
    tuzadi. Standart holatda o'chiq: mavjud foydalanuvchilarning ishi
    o'zgarmasin."""
    # Har qator uchun fayldan o'qimaymiz: classify_row minglab marta
    # chaqiriladi. O'zgartirish set_scratch_mode() orqali ketadi.
    global _SCRATCH_CACHE
    if _SCRATCH_CACHE is None:
        _SCRATCH_CACHE = bool(get_setting(SCRATCH_KEY, False))
    return _SCRATCH_CACHE


def _names_path():
    return os.path.join(_data_dir(), NAMES_FILE)


def load_name_map():
    """Guruhlarning shu foydalanuvchidagi nomi: {"МЕБ": "tovar"}.

    Qoidalar avvalgidek ishlaydi, faqat hisobotga yoziladigan nom
    almashadi — ya'ni tayyor mantiqni yo'qotmasdan o'z atamalaringizni
    ishlatish mumkin."""
    try:
        with open(_names_path(), encoding="utf-8") as f:
            xarita = json.load(f)
        return {str(k): str(v) for k, v in xarita.items() if str(v).strip()}
    except (OSError, ValueError, AttributeError):
        return {}


def save_name_map(xarita):
    try:
        with open(_names_path(), "w", encoding="utf-8") as f:
            json.dump(xarita, f, ensure_ascii=False, indent=2, sort_keys=True)
    except OSError:
        pass


def ai_enabled():
    """AI takliflari yoqilganmi. Standart holatda yoqilgan.

    O'chirgich kerak, chunki kalit endi serverda va AI hammada o'z-o'zidan
    ishlaydi — noto'g'ri taklif bera boshlasa, foydalanuvchida uni to'xtatish
    yo'li bo'lishi shart."""
    return bool(get_setting(AI_ENABLED_KEY, True))


def ai_rules_paths():
    return [os.path.join(_base_dir(), _AI_RULES_FILE),
            os.path.join(_data_dir(), _AI_RULES_FILE)]


AI_CORRECTIONS_FILE = "ai_tuzatishlar.json"
AI_CORRECTIONS_LIMIT = 40


def _corrections_path():
    return os.path.join(_data_dir(), AI_CORRECTIONS_FILE)


def load_ai_corrections():
    try:
        with open(_corrections_path(), encoding="utf-8") as f:
            yozuvlar = json.load(f)
        return yozuvlar if isinstance(yozuvlar, list) else []
    except (OSError, ValueError):
        return []


def record_ai_correction(nomi, matn, ai_taklifi, togri):
    """Foydalanuvchi AI taklifini tuzatganini yozib qo'yadi.

    Nima uchun: aynan shu kontragent keyingi safar lug'atdan tanilib
    ketadi, ya'ni AI gacha yetib bormaydi. Lekin XATO TURI takrorlanadi —
    masalan AI har safar IT firmasini tovar deb o'ylayveradi. Tuzatishlar
    keyingi so'rovlarga misol bo'lib qo'shiladi va AI shu xatoni qayta
    qilmaydi.

    Faqat oxirgi bir nechtasi saqlanadi: so'rov cheksiz o'smasin."""
    nomi = (nomi or "").strip()[:60]
    if not nomi or not togri or ai_taklifi == togri:
        return
    yozuvlar = [
        y for y in load_ai_corrections()
        if y.get("nomi") != nomi
    ]
    yozuvlar.append({
        "nomi": nomi,
        "matn": (matn or "").strip()[:120],
        "ai": ai_taklifi or "",
        "togri": togri,
    })
    try:
        with open(_corrections_path(), "w", encoding="utf-8") as f:
            json.dump(yozuvlar[-AI_CORRECTIONS_LIMIT:], f, ensure_ascii=False, indent=2)
    except OSError:
        pass


def ai_extra_rules():
    """Foydalanuvchi qo'lda yozgan ko'rsatmalar.

    Nima uchun alohida fayl: guruhga oid har bir yangi holat uchun kodni
    o'zgartirib, qayta chiqarish kerak bo'lmasin. Foydalanuvchi o'zi
    ko'rgan xatoni darhol yozib qo'yadi va keyingi ochilishda AI shuni
    hisobga oladi.

    Qaytaradi: (matn, qoidalar_soni)."""
    for yol in ai_rules_paths():
        try:
            with open(yol, encoding="utf-8-sig") as f:
                xom = f.read()
        except OSError:
            continue
        qatorlar = [
            q.strip() for q in xom.splitlines()
            if q.strip() and not q.strip().startswith("#")
        ]
        if qatorlar:
            return "\n".join(f"- {q}" for q in qatorlar), len(qatorlar)
    return "", 0


def ensure_ai_rules_file():
    """Qo'llanma fayli yo'q bo'lsa, namuna bilan yaratadi — foydalanuvchi
    uni topib, to'ldirishi uchun. Qaytaradi: fayl yo'li."""
    yol = ai_rules_paths()[-1]
    if not os.path.exists(yol):
        try:
            os.makedirs(os.path.dirname(yol), exist_ok=True)
            with open(yol, "w", encoding="utf-8") as f:
                f.write(_AI_RULES_TEMPLATE)
        except OSError:
            pass
    return yol

# Guruh nomlari qisqartma bo'lgani uchun AI ularni o'zicha tushunmaydi.
# Sinov shuni ko'rsatdi: izohsiz "Шахрихон туман МИБ" (Majburiy Ijro
# Byurosi) "МЕБ" (tovar) guruhiga qo'shib yuborilgan edi. Izohlar
# qo'shilgach xato yo'qoldi — shuning uchun bu ro'yxat majburiy.
AI_GROUP_HINTS = {
    "МЕБ": "tovar/mahsulot sotib olish (maishiy texnika, telefon, mototsikl, jihoz, xo'jalik mollari)",
    "Терминал": "HUMO / SmartVista terminal orqali tushum",
    "ф пайми": "moliya-hamkor (BNPL, nasiya) bilan hisob-kitob — TBC Fin Service, TBC BNPL",
    "ф вариант": "moliya-hamkor (BNPL, mikromoliya) — Variant Retail Finance, Uzum Nasiya",
    "банк хизмати": "bank komissiyasi, hisoblangan foizlar",
    "иш хаки ПК": "ish haqi, oylik to'lovi",
    "солик даромад": "jismoniy shaxs daromad solig'i",
    "солик КҚС": "qo'shilgan qiymat solig'i to'lovi",
    "солик фойда": "foyda solig'i to'lovi",
    "солик ижтимоий": "ijtimoiy soliq",
    "солик пенсия": "pensiya jamg'armasiga badal",
    "куриклаш": "qo'riqlash xizmati (IIB qoshidagi bo'lim)",
    "коммунал": "suv, gaz uchun to'lov",
    "электр": "elektr energiyasi uchun to'lov",
    "ижара": "ijara to'lovi",
    "хизмат": "xizmat ko'rsatish (IT, konsalting, dasturiy ta'minot, boshqa xizmatlar)",
    "СОРЖ": "SORJ",
    "ДСИ": "soliqdan QAYTGAN summa (ГНИ xulosasiga ko'ra qaytarish) — to'lov emas, qaytim",
    "ароматизация": "havoni aromatizatsiya qilish xizmati",
}


def _ai_prompt(items, groups):
    hints = {g: AI_GROUP_HINTS[g] for g in groups if g in AI_GROUP_HINTS}
    rows = [
        {"id": i, "nomi": str(it.get("name") or "")[:60], "matn": str(it.get("sample") or "")[:200]}
        for i, it in enumerate(items)
    ]
    qollanma, _ = ai_extra_rules()
    # Foydalanuvchi ko'rsatmalari eng oxirida va "ustun turadi" deb
    # beriladi: ular aynan AI xato qilgan holatlar uchun yozilgan,
    # shuning uchun umumiy ta'riflardan kuchliroq bo'lishi kerak.
    qollanma_blok = (
        f"\nFOYDALANUVCHI KO'RSATMALARI — bular yuqoridagi umumiy "
        f"ta'riflardan USTUN turadi:\n{qollanma}\n"
        if qollanma else ""
    )

    # Avval qilingan xatolar misol sifatida beriladi. Aynan o'sha
    # kontragent keyingi safar lug'atdan tanilib ketadi, lekin xato TURI
    # takrorlanadi — masalan IT firmasini har safar tovar deb o'ylash.
    tuzatishlar = load_ai_corrections()
    tuzatish_blok = ""
    if tuzatishlar:
        qatorlar = "\n".join(
            f"- \"{t['nomi']}\" ({t['matn']}) -> to'g'risi: {t['togri']}"
            + (f", sen xato qilib \"{t['ai']}\" degan eding" if t.get("ai") else "")
            for t in tuzatishlar
        )
        tuzatish_blok = (
            "\nAVVAL QILGAN XATOLARING — shunga o'xshash holatlarda "
            f"takrorlama:\n{qatorlar}\n"
        )
    return (
        "Sen O'zbekiston buxgalteriyasida bank ko'chirmalarini guruhlarga ajratasan.\n\n"
        f"MAVJUD GURUHLAR: {sorted(groups)}\n\n"
        f"GURUHLAR NIMANI ANGLATADI:\n{json.dumps(hints, ensure_ascii=False, indent=1)}\n\n"
        "QOIDALAR:\n"
        "- Faqat yuqoridagi ro'yxatdan tanla, yangi nom o'ylab topma.\n"
        "- O'xshash qisqartmalarni chalkashtirma: \"МИБ\" (Majburiy Ijro Byurosi) bu \"МЕБ\" EMAS.\n"
        "- Tovar sotib olish (texnika, transport, aloqa vositasi, xo'jalik mollari) -> МЕБ\n"
        "- Ishonching past bo'lsa yoki mos guruh bo'lmasa \"?\" yoz. "
        "Noto'g'ri taxmindan ko'ra \"?\" yaxshiroq.\n"
        f"{qollanma_blok}{tuzatish_blok}\n"
        "Javobni JSON massiv sifatida qaytar:\n"
        '[{"id":0,"guruh":"...","ishonch":"yuqori|past","sabab":"qisqa izoh"}]\n\n'
        f"Qatorlar:\n{json.dumps(rows, ensure_ascii=False, indent=1)}"
    )


def ai_suggest(items, groups, timeout=60):
    """Har bir kontragent uchun guruh taklif qiladi.

    items  — find_unresolved() qaytargan yozuvlar ro'yxati
    groups — ruxsat etilgan guruh nomlari

    Qaytaradi: (takliflar, xato), bunda takliflar —
    {indeks: {"guruh", "ishonch", "sabab"}}, xato esa None yoki
    foydalanuvchiga ko'rsatiladigan qisqa sabab.

    Xato matni alohida qaytariladi, chunki avval har qanday nosozlik —
    tarmoq uzilishi, chegara, server xatosi — bir xil "taklif topilmadi"
    bo'lib ko'rinardi va AI buzilganini bilib bo'lmasdi."""
    if not items or not groups:
        return {}, None
    try:
        body = json.dumps({
            "host": platform.node() or "noma'lum",
            "model": AI_MODEL,
            "prompt": _ai_prompt(items, groups),
        }).encode("utf-8")
        req = urllib.request.Request(
            _AI_URL,
            data=body,
            # User-Agent shart: Cloudflare urllib'ning standart nomini
            # bloklaydi va 403 (xato 1010) qaytaradi.
            headers={"Content-Type": "application/json", "User-Agent": _USER_AGENT},
        )
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            javob = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        # Worker xato sababini javob tanasida beradi (chegara, sozlama,
        # yuqori oqim) — o'shani o'qib, tushunarli xabarga aylantiramiz.
        try:
            kod = json.loads(e.read().decode("utf-8")).get("error", "")
        except Exception:
            kod = ""
        return {}, _AI_ERRORS.get(kod, f"AI xizmatida xato ({e.code}).")
    except Exception:
        return {}, "AI xizmatiga ulanib bo'lmadi — internetni tekshiring."

    if not javob.get("ok"):
        return {}, _AI_ERRORS.get(javob.get("error", ""), "AI javob bermadi.")

    try:
        answers = json.loads(javob["text"])
    except (KeyError, ValueError):
        return {}, "AI javobini o'qib bo'lmadi."

    allowed = set(groups)
    out = {}
    for a in answers if isinstance(answers, list) else []:
        try:
            idx = int(a["id"])
            guruh = str(a.get("guruh") or "").strip()
        except (KeyError, TypeError, ValueError):
            continue
        # AI o'ylab topgan nomlarni qabul qilmaymiz — faqat mavjud guruhlar.
        if guruh not in allowed:
            continue
        out[idx] = {
            "guruh": guruh,
            "ishonch": str(a.get("ishonch") or "").strip(),
            "sabab": str(a.get("sabab") or "").strip(),
        }
    return out, None


def ai_known_groups():
    """AI tanlashi mumkin bo'lgan guruhlar — qoidalarda uchraydiganlarning
    hammasi."""
    return (
        set(KNOWN_VENDOR_INN.values())
        | {cat for _, cat in PURPOSE_FIRST_RULES}
        | {cat for _, cat, _ in TEXT_RULES}
    )


def enable_mousewheel(toplevel, canvas):
    """Sichqoncha g'ildiragi bilan aylantirishni yoqadi.

    tk.Canvas g'ildirak hodisasini O'ZI eshitmaydi — qo'lda ulash kerak,
    aks holda faqat yon tarafdagi chiziqni sichqoncha bilan tortish
    qoladi.

    Bog'lash TOPLEVEL ga qilinadi, canvas'ga emas: Tk'da hodisa
    widget -> klass -> toplevel -> "all" zanjiri bo'ylab tarqaladi,
    shuning uchun kursor ichki widget (yorliq, kiritish maydoni) ustida
    turganda ham ishlaydi. Canvas'ning o'ziga bog'lansa, kursor biror
    yorliq ustiga tushishi bilan g'ildirak ishlamay qolardi."""

    def on_wheel(event):
        # Windows'da event.delta 120 ning karrali (bir "tirqish" = 120).
        canvas.yview_scroll(-int(event.delta / 120), "units")
        return "break"

    def on_x11_wheel(yonalish):
        def handler(_event):
            canvas.yview_scroll(yonalish, "units")
            return "break"
        return handler

    toplevel.bind("<MouseWheel>", on_wheel)         # Windows / macOS
    toplevel.bind("<Button-4>", on_x11_wheel(-1))   # X11: yuqoriga
    toplevel.bind("<Button-5>", on_x11_wheel(1))    # X11: pastga


class UnresolvedDialog(tk.Toplevel):
    """Modal oyna: hisobotlarda kategoriyasi aniqlanmagan (nomlanmagan)
    kontragentlar ro'yxatini ko'rsatadi va har biri uchun nom/kategoriya
    kiritishni so'raydi. Kiritilgan qiymatlar persist_dictionary.json ga
    saqlanadi va keyingi barcha loyihalarda avtomatik tanilib qoladi."""

    def __init__(self, parent, unresolved):
        super().__init__(parent)
        self.title("Nomlanmagan kontragentlar topildi")
        # Qatorlarda hisob raqam, МФО va to'lov izohi bir yo'lda keladi —
        # tor oynada ular o'ngdan kesilib qolardi. Ekran ruxsat berganicha
        # keng ochamiz, lekin ekrandan chiqib ketmasin.
        kengligi = min(1180, max(760, self.winfo_screenwidth() - 160))
        balandligi = min(780, max(520, self.winfo_screenheight() - 200))
        self.geometry(f"{kengligi}x{balandligi}")
        self.minsize(680, 460)
        self.transient(parent)
        self.grab_set()
        self.result_entries = {}  # key -> (info, tk.StringVar)
        self._hint_labels = {}    # key -> AI taklifi ko'rsatiladigan yorliq
        self._row_order = []      # AI javobidagi indeks -> key moslashuvi
        self._ai_taklif = {}      # key -> AI aytgan guruh (tuzatishni bilish uchun)
        self._wrap_labels = []    # oyna kengligiga qarab o'raladigan yozuvlar
        self._last_wrap = 0
        self.confirmed = False

        header = ttk.Frame(self, padding=10)
        header.pack(fill="x")
        if scratch_mode():
            ttk.Label(
                header, foreground="#c62828", wraplength=1000, justify="left",
                text=("\"Faqat o'zim kiritgan guruhlar\" rejimi yoqilgan - shuning uchun "
                      "tayyor guruhlar ishlamayapti va deyarli hamma kontragent "
                      "so'ralyapti. Guruhlarni boshqarish oynasidan o'chirsa bo'ladi."),
            ).pack(anchor="w", pady=(0, 8))
        ttk.Label(
            header,
            text=(
                f"{len(unresolved)} ta kontragent hech qanday kategoriyaga to'g'ri kelmadi.\n"
                "Har biri uchun kategoriya nomini kiriting (bo'sh qoldirsangiz \"?\" bo'lib qoladi). "
                "Kiritganlaringiz keyingi loyihalar uchun ham eslab qolinadi."
            ),
            wraplength=720,
            justify="left",
        ).pack(anchor="w")

        canvas = tk.Canvas(self, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        inner = ttk.Frame(canvas)
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        inner_id = canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True, padx=(10, 0), pady=(0, 10))
        scrollbar.pack(side="left", fill="y", pady=(0, 10))
        enable_mousewheel(self, canvas)
        canvas.bind("<Configure>", lambda e: self._fit_width(canvas, inner_id, e.width))

        for key, info in sorted(unresolved.items(), key=lambda kv: -kv[1]["count"]):
            row = ttk.Frame(inner, padding=6, relief="groove", borderwidth=1)
            row.pack(fill="x", padx=6, pady=4)

            label_bits = []
            if info["name"]:
                label_bits.append(info["name"])
            if info.get("account"):
                label_bits.append(f"Xisob raqam {info['account']}")
            elif info.get("inn"):
                label_bits.append(f"ИНН {info['inn']}")
            if info.get("mfo"):
                label_bits.append(f"МФО {info['mfo']}")
            label_bits.append(f"{info['count']} qatorda uchraydi")
            if info.get("always_ask"):
                label_bits.append("HAR SAFAR SO'RALADI (ko'p maqsadli hisob)")
            sarlavha = ttk.Label(row, text="  |  ".join(label_bits), font=("", 9, "bold"),
                                 justify="left")
            sarlavha.pack(anchor="w", fill="x")
            self._wrap_labels.append(sarlavha)
            if info["sample"]:
                izoh = ttk.Label(row, text=info["sample"], foreground="#555", justify="left")
                izoh.pack(anchor="w", fill="x")
                self._wrap_labels.append(izoh)

            entry_row = ttk.Frame(row)
            entry_row.pack(fill="x", pady=(4, 0))
            ttk.Label(entry_row, text="Kategoriya nomi:").pack(side="left")
            var = tk.StringVar(value="")
            ttk.Entry(entry_row, textvariable=var, width=30).pack(side="left", padx=(6, 0))
            # AI taklifi shu yerda ko'rinadi. Taklif faqat maslahat —
            # yozilgan qiymatni istagancha o'zgartirish mumkin.
            hint = ttk.Label(entry_row, text="", foreground="#0a66c2")
            hint.pack(side="left", padx=(10, 0))
            self.result_entries[key] = (info, var)
            self._hint_labels[key] = hint
            self._row_order.append(key)

        # before=... bilan footer paketlash tartibida kengayuvchi
        # qismdan OLDIN turadi — shunda oyna kichraytirilganda tugmalar
        # kesilmaydi, o'rniga ro'yxat qisqaradi.
        footer = ttk.Frame(self, padding=10)
        footer.pack(side="bottom", fill="x", before=canvas)
        ttk.Button(footer, text="Saqlash va davom etish", command=self._on_confirm).pack(side="right")
        ttk.Button(footer, text="Bekor qilish", command=self._on_cancel).pack(side="right", padx=(0, 8))

        # O'chirgich shu yerda turadi — AI aynan shu oynada ishlaydi,
        # noto'g'ri taklif ko'rgan odam uni darhol shu yerdan to'xtata
        # olishi kerak, sozlamalar ichidan qidirmasdan.
        self.ai_var = tk.BooleanVar(value=ai_enabled())
        ttk.Checkbutton(
            footer, text="AI takliflari", variable=self.ai_var, command=self._toggle_ai
        ).pack(side="left", padx=(0, 14))

        self._ai_status = ttk.Label(footer, text="", foreground="#666666")
        self._ai_status.pack(side="left")
        self._start_ai()

        self.protocol("WM_DELETE_WINDOW", self._on_cancel)

    def _fit_width(self, canvas, inner_id, kenglik):
        """Ro'yxatni oyna kengligiga moslaydi.

        Ikki ish qilinadi: ichki freym canvas kengligiga cho'ziladi (aks
        holda u o'z tabiiy kengligida qolib, o'ng tomoni ko'rinmay
        qolardi) va uzun yozuvlar shu kenglikda o'raladi — kesilmaydi.

        Kenglik sezilarli o'zgarmasa tegilmaydi: wraplength o'zgarishi
        yozuv balandligini o'zgartiradi, u esa yangi hodisa keltirib
        chiqaradi — oyna cho'zilayotganda bu halqa sekinlashtirardi."""
        canvas.itemconfigure(inner_id, width=kenglik)
        if abs(kenglik - self._last_wrap) < 8:
            return
        self._last_wrap = kenglik
        wrap = max(320, kenglik - 48)
        for lbl in self._wrap_labels:
            lbl.configure(wraplength=wrap)

    # -------------------------------------------------- AI takliflari
    def _toggle_ai(self):
        yoqilgan = bool(self.ai_var.get())
        set_setting(AI_ENABLED_KEY, yoqilgan)
        if yoqilgan:
            self._start_ai()
        else:
            self._ai_status.configure(text="AI takliflari o'chirilgan.")

    def _start_ai(self):
        """Fon oqimida guruh takliflarini so'raydi. Bu faqat yordam —
        javob kelmasa ham oyna avvalgidek ishlayveradi."""
        if not self._row_order:
            return
        if not self.ai_var.get():
            self._ai_status.configure(text="AI takliflari o'chirilgan.")
            return

        # Fayl birinchi ochilishda namuna bilan yaratiladi, aks holda
        # foydalanuvchi uni qayerga yozishni bilmaydi.
        ensure_ai_rules_file()
        _, self._qollanma_soni = ai_extra_rules()

        self._ai_status.configure(text="AI takliflari so'ralmoqda...")
        items = [self.result_entries[k][0] for k in self._row_order]
        guruhlar = ai_known_groups()

        # Javob navbat orqali qaytariladi, fon oqimidan to'g'ridan-to'g'ri
        # widget'ga tegilmaydi: Tkinter faqat asosiy oqimdan chaqirilishi
        # kerak, aks holda "main thread is not in main loop" xatosi chiqadi
        # yoki interfeys tushunarsiz buziladi.
        self._ai_queue = queue.Queue()
        threading.Thread(
            target=lambda: self._ai_queue.put(ai_suggest(items, guruhlar)),
            daemon=True,
        ).start()
        self.after(150, self._poll_ai)

    def _poll_ai(self):
        if not self.winfo_exists():
            return
        try:
            takliflar, xato = self._ai_queue.get_nowait()
        except queue.Empty:
            self.after(150, self._poll_ai)
            return
        if xato:
            self._ai_status.configure(text=xato)
            return
        self._apply_ai(takliflar)

    def _apply_ai(self, takliflar):
        """Takliflarni maydonlarga yozadi. Foydalanuvchi allaqachon biror
        narsa yozgan bo'lsa — tegilmaydi."""
        if not self.winfo_exists():
            return
        qollandi = 0
        for idx, data in takliflar.items():
            if idx >= len(self._row_order):
                continue
            key = self._row_order[idx]
            _info, var = self.result_entries[key]
            if var.get().strip():
                continue
            var.set(data["guruh"])
            self._ai_taklif[key] = data["guruh"]
            belgi = "AI" if data.get("ishonch") == "yuqori" else "AI (ishonch past)"
            self._hint_labels[key].configure(text=f"{belgi} · {data.get('sabab', '')[:60]}")
            qollandi += 1
        # Qo'llanmadagi qoidalar soni ko'rsatiladi — foydalanuvchi yozgan
        # ko'rsatma haqiqatan o'qilganini shundan biladi.
        qollanma = (
            f"  |  qo'llanma: {self._qollanma_soni} ta qoida"
            if getattr(self, "_qollanma_soni", 0) else
            f"  |  qo'llanma bo'sh ({os.path.basename(ai_rules_paths()[-1])})"
        )
        if qollandi:
            self._ai_status.configure(
                text=f"AI {qollandi} ta taklif berdi — tekshirib, kerak bo'lsa o'zgartiring.{qollanma}"
            )
        else:
            self._ai_status.configure(text=f"AI mos taklif topa olmadi.{qollanma}")

    def _on_confirm(self):
        # AI taklifi tuzatilgan bo'lsa yozib qo'yamiz — keyingi so'rovlarda
        # misol bo'lib beriladi va AI shu xatoni takrorlamaydi.
        for key, (info, var) in self.result_entries.items():
            # Ko'p maqsadli kontragentlar (G'aznachilik) tashlab
            # yuboriladi: ularning guruhi har safar boshqacha, misol
            # sifatida saqlash AI ni faqat adashtiradi.
            if info.get("always_ask"):
                continue
            record_ai_correction(
                info.get("name"), info.get("sample"),
                self._ai_taklif.get(key, ""), var.get().strip(),
            )
        self.confirmed = True
        self.destroy()

    def _on_cancel(self):
        self.confirmed = False
        self.destroy()

    def get_assignments(self):
        """Foydalanuvchi to'ldirgan qatorlarni qaytaradi.

        Har bir element: (info, guruh). `info` ichida "always_ask" bayrog'i
        bor — G'aznachilik kabi ko'p maqsadli kontragentlar uchun javob
        diskka yozilmaydi, faqat shu faylga qo'llaniladi."""
        out = []
        for key, (info, var) in self.result_entries.items():
            cat = var.get().strip()
            if cat:
                out.append((info, cat))
        return out


class NameMapDialog(tk.Toplevel):
    """Guruh nomlarini shu foydalanuvchi uchun qayta nomlash.

    Qoidalar o'zgarmaydi — faqat hisobotga yoziladigan so'z almashadi.
    Ya'ni tayyor mantiqni yo'qotmasdan o'z atamalaringizda ishlash
    mumkin: masalan "МЕБ" o'rniga "товар"."""

    def __init__(self, parent):
        super().__init__(parent)
        self.title("Guruh nomlarim")
        self.geometry("560x620")
        self.minsize(460, 400)
        self.transient(parent)
        self.grab_set()

        self.xarita = load_name_map()
        self.vars = {}

        ttk.Label(
            self, padding=14,
            text=("Har bir guruh sizning hisobotingizda qanday nom bilan chiqishini "
                  "yozing. Bo'sh qoldirsangiz, standart nom ishlatiladi."),
            wraplength=520, justify="left", foreground="#666666",
        ).pack(anchor="w")

        canvas = tk.Canvas(self, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        inner = ttk.Frame(canvas)
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        inner_id = canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True, padx=(14, 0), pady=(0, 10))
        scrollbar.pack(side="left", fill="y", pady=(0, 10))
        enable_mousewheel(self, canvas)
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure(inner_id, width=e.width))

        footer = ttk.Frame(self, padding=12)
        footer.pack(side="bottom", fill="x", before=canvas)
        ttk.Button(footer, text="Saqlash", style="Accent.TButton",
                   command=self._save).pack(side="right")
        ttk.Button(footer, text="Yopish", command=self.destroy).pack(side="right", padx=(0, 8))

        # Ro'yxatga qoidalardagi guruhlar ham, foydalanuvchi o'zi kiritgan
        # guruhlar ham kiradi — ikkalasini ham qayta nomlash mumkin.
        guruhlar = sorted(ai_known_groups() | set(load_all_mappings().values()))
        for guruh in guruhlar:
            qator = ttk.Frame(inner, padding=(6, 4))
            qator.pack(fill="x")
            ttk.Label(qator, text=guruh, width=22).pack(side="left")
            ttk.Label(qator, text="->", foreground="#999999").pack(side="left", padx=(0, 8))
            var = tk.StringVar(value=self.xarita.get(guruh, ""))
            ttk.Entry(qator, textvariable=var, width=24).pack(side="left", fill="x", expand=True)
            self.vars[guruh] = var

    def _save(self):
        yangi = {}
        for guruh, var in self.vars.items():
            nom = var.get().strip()
            if nom and nom != guruh:
                yangi[guruh] = nom
        save_name_map(yangi)
        messagebox.showinfo(
            "Saqlandi",
            f"{len(yangi)} ta guruh qayta nomlandi." if yangi
            else "Barcha guruhlar standart nomda qoldi.",
        )
        self.destroy()


class GroupsManagerDialog(tk.Toplevel):
    """Guruhlar (kategoriyalar) boshqaruv oynasi. Foydalanuvchi ilovani
    birinchi marta ishga tushirganda — yoki istalgan payt — ma'lum hisob
    raqam (ИНН) yoki kompaniya nomi uchun guruh (kategoriya) belgilab
    qo'yishi mumkin, fayl tashlanishini kutmasdan. Barcha yozuvlar
    persist_dictionary.json ga saqlanadi va shu zahoti ishlatila boshlaydi."""

    def __init__(self, parent):
        super().__init__(parent)
        self.title("Guruhlarni boshqarish")
        self.geometry("720x560")
        self.minsize(600, 420)
        self.transient(parent)
        self.grab_set()

        self.mapping = dict(load_all_mappings())
        self.rows = {}  # key -> (row_frame, id_var, cat_var)

        header = ttk.Frame(self, padding=14)
        header.pack(fill="x")
        ttk.Label(header, text="Guruhlarni boshqarish", font=("Segoe UI Semibold", 13)).pack(anchor="w")
        ttk.Label(
            header,
            text=(
                "Bu yerda hisob raqam yoki kompaniya nomi qaysi guruhga tegishli ekanini oldindan "
                "belgilab qo'yishingiz mumkin. Bitta hisob raqam orqali turli maqsaddagi to'lovlar "
                "o'tsa (masalan G'aznachilik: soliq, elektr...), \"To'lov maqsadi matni bo'yicha\" "
                "variantini tanlab, matndan bir bo'lakni (masalan \"электр учун\") kiriting — "
                "shunda ilova aynan shu matnli qatorlarni alohida guruhga ajratadi."
            ),
            wraplength=680, foreground="#666666", justify="left",
        ).pack(anchor="w", pady=(4, 0))

        add_box = ttk.LabelFrame(self, text="Yangi guruh qo'shish", padding=10)
        add_box.pack(fill="x", padx=14, pady=(10, 0))

        self.new_type = tk.StringVar(value="account")
        type_row = ttk.Frame(add_box)
        type_row.pack(fill="x")
        ttk.Radiobutton(type_row, text="Xisob raqam bo'yicha", variable=self.new_type, value="account").pack(side="left")
        ttk.Radiobutton(type_row, text="Nomi bo'yicha", variable=self.new_type, value="name").pack(side="left", padx=(12, 0))
        ttk.Radiobutton(
            type_row, text="To'lov maqsadi matni bo'yicha", variable=self.new_type, value="text"
        ).pack(side="left", padx=(12, 0))

        fields_row = ttk.Frame(add_box)
        fields_row.pack(fill="x", pady=(8, 0))
        ttk.Label(fields_row, text="Xisob raqam / Nomi:").pack(side="left")
        self.new_id_var = tk.StringVar()
        ttk.Entry(fields_row, textvariable=self.new_id_var, width=22).pack(side="left", padx=(6, 16))
        ttk.Label(fields_row, text="Guruh nomi:").pack(side="left")
        self.new_cat_var = tk.StringVar()
        ttk.Entry(fields_row, textvariable=self.new_cat_var, width=22).pack(side="left", padx=(6, 16))
        ttk.Button(fields_row, text="+ Qo'shish", command=self._add_row).pack(side="left")

        # Bitta kontragent yo'nalishiga qarab boshqa guruhga tushishi
        # mumkin — masalan chiqimda to'lov, tushumda esa boshqa narsa.
        dir_row = ttk.Frame(add_box)
        dir_row.pack(fill="x", pady=(8, 0))
        ttk.Label(dir_row, text="Qaysi holatda:").pack(side="left")
        self.new_dir = tk.StringVar(value="")
        ttk.Radiobutton(dir_row, text="Har doim", variable=self.new_dir,
                        value="").pack(side="left", padx=(8, 0))
        ttk.Radiobutton(dir_row, text="Faqat chiqimda (дебет)", variable=self.new_dir,
                        value=DIR_DEBIT).pack(side="left", padx=(12, 0))
        ttk.Radiobutton(dir_row, text="Faqat tushumda (кредит)", variable=self.new_dir,
                        value=DIR_CREDIT).pack(side="left", padx=(12, 0))

        # Har bir foydalanuvchining o'z atamalari bo'lishi mumkin, shuning
        # uchun sozlamalar aynan shu oynada — guruhlar bilan bir joyda.
        moslash = ttk.LabelFrame(self, text="Meniki uchun moslash", padding=10)
        moslash.pack(fill="x", padx=14, pady=(10, 0))

        self.scratch_var = tk.BooleanVar(value=scratch_mode())
        ttk.Checkbutton(
            moslash,
            text="Faqat o'zim kiritgan guruhlar ishlatilsin (tayyor qoidalar o'chiriladi)",
            variable=self.scratch_var, command=self._toggle_scratch,
        ).pack(anchor="w")
        ttk.Label(
            moslash,
            text=("Yoqilsa, ilovaga oldindan yozilgan guruhlar qo'llanmaydi — har bir "
                  "kontragent so'raladi va siz ro'yxatni o'zingiz tuzasiz. "
                  "Tayyor qoidalar kerak, lekin nomlari boshqacha bo'lsa, "
                  "buni yoqmasdan \"Guruh nomlarim\" dan foydalaning."),
            wraplength=660, foreground="#666666", justify="left",
        ).pack(anchor="w", pady=(4, 8))
        self.style_var = tk.BooleanVar(value=titled_style())
        ttk.Checkbutton(
            moslash,
            text="Hisobot yuqorisida sarlavha va qizil yig'indi bo'lsin",
            variable=self.style_var, command=self._toggle_style,
        ).pack(anchor="w", pady=(0, 8))

        ttk.Button(moslash, text="✎ Guruh nomlarim...", command=self._open_names).pack(anchor="w")

        pick_row = ttk.Frame(add_box)
        pick_row.pack(fill="x", pady=(8, 0))
        ttk.Button(pick_row, text="📄 Excel'dan tanlash...", command=self._pick_from_excel).pack(side="left")
        ttk.Label(
            pick_row,
            text="— Xisob raqamni qo'lda yozish o'rniga, haqiqiy hisobot faylini ochib, kontragentlarni belgilab tanlang.",
            foreground="#888888",
        ).pack(side="left", padx=(8, 0))

        list_label = ttk.Label(self, text="Mavjud guruhlar:", font=("Segoe UI Semibold", 10))
        list_label.pack(anchor="w", padx=14, pady=(14, 4))

        canvas_frame = ttk.Frame(self)
        canvas_frame.pack(fill="both", expand=True, padx=14)
        self.canvas = tk.Canvas(canvas_frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(canvas_frame, orient="vertical", command=self.canvas.yview)
        self.inner = ttk.Frame(self.canvas)
        self.inner.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.canvas.configure(yscrollcommand=scrollbar.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="left", fill="y")
        enable_mousewheel(self, self.canvas)

        self._render_rows()

        # before=... bilan footer paketlash tartibida kengayuvchi
        # qismdan OLDIN turadi — shunda oyna kichraytirilganda tugmalar
        # kesilmaydi, o'rniga ro'yxat qisqaradi.
        footer = ttk.Frame(self, padding=14)
        footer.pack(side="bottom", fill="x", before=canvas_frame)
        ttk.Button(footer, text="Saqlash", style="Accent.TButton", command=self._save).pack(side="right")
        ttk.Button(footer, text="Yopish", command=self.destroy).pack(side="right", padx=(0, 8))

    def _toggle_style(self):
        set_setting(STYLE_KEY, self.style_var.get())

    def _toggle_scratch(self):
        set_scratch_mode(self.scratch_var.get())

    def _open_names(self):
        dlg = NameMapDialog(self)
        self.wait_window(dlg)

    def _render_rows(self):
        for child in self.inner.winfo_children():
            child.destroy()
        self.rows = {}
        if not self.mapping:
            ttk.Label(self.inner, text="Hali hech qanday guruh belgilanmagan.", foreground="#888888").pack(
                anchor="w", pady=10
            )
            return
        for key, cat in sorted(self.mapping.items()):
            self._add_row_widget(key, cat)

    def _add_row_widget(self, key, category):
        toza, yonalish = split_direction(key)
        if toza.startswith(_TEXT_KEY_PREFIX):
            display_id = toza[len(_TEXT_KEY_PREFIX):]
            label_prefix = "Matn: "
        elif toza.startswith(_NAME_KEY_PREFIX):
            display_id = toza[len(_NAME_KEY_PREFIX):]
            label_prefix = "Nomi: "
        else:
            display_id = toza
            label_prefix = "Xisob raqam: "

        row = ttk.Frame(self.inner, padding=6, relief="groove", borderwidth=1)
        row.pack(fill="x", pady=3)
        ttk.Label(row, text=f"{label_prefix}{display_id}", width=32, anchor="w").pack(side="left")
        if yonalish:
            ttk.Label(
                row, foreground="#0a66c2",
                text="chiqimda" if yonalish == DIR_DEBIT else "tushumda",
            ).pack(side="left", padx=(0, 6))
        cat_var = tk.StringVar(value=category)
        ttk.Entry(row, textvariable=cat_var, width=22).pack(side="left", padx=(6, 6))
        ttk.Button(row, text="Saqlash", command=lambda k=key, v=cat_var: self._update_row(k, v)).pack(side="left")
        ttk.Button(row, text="O'chirish", command=lambda k=key: self._delete_row(k)).pack(side="left", padx=(6, 0))
        self.rows[key] = (row, cat_var)

    def _update_row(self, key, var):
        new_cat = var.get().strip()
        if not new_cat:
            messagebox.showwarning("Diqqat", "Guruh nomi bo'sh bo'lishi mumkin emas.")
            return
        self.mapping[key] = new_cat
        save_all_mappings(self.mapping)
        messagebox.showinfo("Saqlandi", "Guruh yangilandi.")

    def _delete_row(self, key):
        if not messagebox.askyesno("Tasdiqlash", "Ushbu guruhni o'chirishni tasdiqlaysizmi?"):
            return
        self.mapping.pop(key, None)
        save_all_mappings(self.mapping)
        self._render_rows()

    def _pick_from_excel(self):
        path = filedialog.askopenfilename(
            title="Hisobot faylini tanlang",
            filetypes=[("Excel fayllar", "*.xlsx *.xls"), ("Barcha fayllar", "*.*")],
        )
        if not path:
            return
        try:
            _wb, _ws, _hdr, rows = load_raw_rows(path)
        except Exception as e:
            messagebox.showerror("Xato", f"Faylni o'qib bo'lmadi:\n{e}")
            return
        parties = unique_counterparties(rows)
        if not parties:
            messagebox.showinfo("Diqqat", "Faylda kontragentlar topilmadi.")
            return
        picker = CounterpartyPickerDialog(self, parties)
        self.wait_window(picker)
        if picker.applied:
            self.mapping = dict(load_all_mappings())
            self._render_rows()

    def _add_row(self):
        ident = self.new_id_var.get().strip()
        cat = self.new_cat_var.get().strip()
        if not ident or not cat:
            messagebox.showwarning("Diqqat", "Xisob raqam / Nomi / Matn va Guruh nomini kiriting.")
            return
        kind = self.new_type.get()
        if kind == "account":
            key = ident
        elif kind == "text":
            key = f"{_TEXT_KEY_PREFIX}{ident}"
        else:
            key = f"{_NAME_KEY_PREFIX}{ident}"
        key = with_direction(key, self.new_dir.get())
        self.mapping[key] = cat
        save_all_mappings(self.mapping)
        self.new_id_var.set("")
        self.new_cat_var.set("")
        self._render_rows()

    def _save(self):
        save_all_mappings(self.mapping)
        self.destroy()


class CounterpartyPickerDialog(tk.Toplevel):
    """Xom hisobot faylini jadval (Excel'ga o'xshash) ko'rinishida ochib,
    foydalanuvchi bir nechta kontragentni belgilab (check qilib), bittasiga
    guruh nomi berib bir yo'la qo'sha oladigan oyna. INN'ni qo'lda yozish
    o'rniga, haqiqiy fayldan tanlab olish uchun."""

    def __init__(self, parent, parties):
        super().__init__(parent)
        self.title("Fayldan kontragent tanlash")
        self.geometry("860x560")
        self.minsize(680, 420)
        self.transient(parent)
        self.grab_set()
        self.applied = False
        self.parties = parties
        self.checked = set()

        header = ttk.Frame(self, padding=(14, 14, 14, 6))
        header.pack(fill="x")
        ttk.Label(header, text="Fayldan kontragent tanlash", font=("Segoe UI Semibold", 13)).pack(anchor="w")
        ttk.Label(
            header,
            text="Kerakli qatorlarni belgilang (katakchani bosing), so'ng pastda guruh nomini kiritib qo'shing.",
            foreground="#666666", wraplength=820,
        ).pack(anchor="w", pady=(2, 0))

        tree_frame = ttk.Frame(self)
        tree_frame.pack(fill="both", expand=True, padx=14, pady=(8, 0))

        columns = ("check", "account", "mfo", "name", "sample")
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings", selectmode="none", height=6)
        self.tree.heading("check", text="✓")
        self.tree.heading("account", text="Xisob raqam")
        self.tree.heading("mfo", text="МФО")
        self.tree.heading("name", text="Nomi")
        self.tree.heading("sample", text="Namuna matn")
        self.tree.column("check", width=36, anchor="center")
        self.tree.column("account", width=140, anchor="w")
        self.tree.column("mfo", width=80, anchor="w")
        self.tree.column("name", width=200, anchor="w")
        self.tree.column("sample", width=360, anchor="w")
        self.tree.pack(side="left", fill="both", expand=True)
        self.tree.bind("<Button-1>", self._on_click)

        scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        scroll.pack(side="left", fill="y")
        self.tree.configure(yscrollcommand=scroll.set)

        for key, info in sorted(parties.items(), key=lambda kv: (kv[1]["name"] or kv[1]["account"]).lower()):
            self.tree.insert(
                "", "end", iid=key,
                values=("☐", info["account"], info.get("mfo", ""), info["name"], info["sample"]),
            )

        # before=... bilan footer paketlash tartibida kengayuvchi
        # qismdan OLDIN turadi — shunda oyna kichraytirilganda tugmalar
        # kesilmaydi, o'rniga ro'yxat qisqaradi.
        footer = ttk.Frame(self, padding=14)
        footer.pack(side="bottom", fill="x", before=tree_frame)
        ttk.Label(footer, text="Guruh nomi:").pack(side="left")
        self.cat_var = tk.StringVar()
        ttk.Entry(footer, textvariable=self.cat_var, width=22).pack(side="left", padx=(6, 14))
        self.count_label = ttk.Label(footer, text="0 ta belgilandi")
        self.count_label.pack(side="left")
        ttk.Button(footer, text="Guruhga qo'shish", style="Accent.TButton", command=self._apply).pack(side="right")
        ttk.Button(footer, text="Yopish", command=self.destroy).pack(side="right", padx=(0, 8))

    def _on_click(self, event):
        if self.tree.identify_region(event.x, event.y) != "cell":
            return
        row = self.tree.identify_row(event.y)
        if not row or self.tree.identify_column(event.x) != "#1":
            return
        if row in self.checked:
            self.checked.remove(row)
            self.tree.set(row, "check", "☐")
        else:
            self.checked.add(row)
            self.tree.set(row, "check", "☑")
        self.count_label.config(text=f"{len(self.checked)} ta belgilandi")

    def _apply(self):
        if not self.checked:
            messagebox.showwarning("Diqqat", "Kamida bitta qatorni belgilang.")
            return
        cat = self.cat_var.get().strip()
        if not cat:
            messagebox.showwarning("Diqqat", "Guruh nomini kiriting.")
            return
        for key in self.checked:
            info = self.parties[key]
            if info["account"]:
                save_learned_category(info["account"], "", cat)
            else:
                save_learned_category("", info["name"], cat)
        self.applied = True
        messagebox.showinfo("Saqlandi", f"{len(self.checked)} ta kontragent \"{cat}\" guruhiga qo'shildi.")
        self.destroy()


class App(tk.Tk):
    STATUS_COLORS = {
        "Kutmoqda": "#666666",
        "Ishlanmoqda...": "#0a66c2",
        "Tayyor": "#1a7f37",
        "Xato": "#c62828",
    }

    def __init__(self):
        super().__init__()
        # Sinov nusxasini ishchi nusxa deb o'ylab qolmaslik uchun
        # shoxobcha nomi sarlavhada turadi.
        self.title(
            "SoddaHisobot"
            if SOURCE_BRANCH == "master"
            else f"SoddaHisobot  —  SINOV REJIMI: {SOURCE_BRANCH}"
        )
        self._apply_dpi_scaling()
        self.geometry("960x680")
        # Kichraytirilganda ham barcha tugmalar ko'rinib turadigan eng kichik
        # o'lcham (ro'yxat va jurnal qisqaradi, tugmalar kesilmaydi).
        self.minsize(780, 520)

        sv_ttk.set_theme("light")
        # Oyna kattalashtirilganda Windows yangi ochilgan bo'sh joyni
        # widget'lar chizilgunicha oyna fonи bilan to'ldiradi. Tk'ning
        # standart foni tema fonidan farq qilgani uchun shu lahzada qora
        # yoki kulrang yo'l ko'rinib qolardi — fonni temaga tenglashtiramiz.
        self.configure(background="#fafafa")
        self._setup_fonts()
        self._setup_styles()

        self.files = []  # list[FileRow]
        self.out_dir = tk.StringVar(value="")
        self.is_running = False
        self.ui_queue = queue.Queue()
        self._last_wrap_width = 0

        # Tkinter ushlanmagan xatoni odatda konsolga yozadi, grafik ilovada
        # esa konsol yo'q — xato butunlay ko'rinmay ketardi.
        self.report_callback_exception = self._on_unhandled

        self._build_ui()
        self._last_state = self.state()
        self.bind("<Configure>", self._on_root_configure)
        # Kompyuterda .exe ning bir necha nusxasi qolib ketishi mumkin
        # (qayta-qayta qo'lda yuklanganda). Qaysi nusxa ochilganini bilish
        # uchun joylashuvni jurnalga yozamiz — aks holda "menda eski versiya
        # ko'rinyapti" degan holatni tekshirib bo'lmaydi.
        self._log(
            f"Versiya v{CORE_VERSION}  |  Shoxobcha: {SOURCE_BRANCH}"
            f"  |  Joylashuvi: {os.path.abspath(sys.executable)}"
        )
        # Bu rejim tayyor qoidalarning hammasini o'chiradi. Jimgina
        # yoqilib qolsa "ilova hech narsani tanimay qoldi" degan
        # tushunmovchilik chiqadi, shuning uchun har ochilishda aytamiz.
        if scratch_mode():
            self._log(
                "DIQQAT: \"Faqat o'zim kiritgan guruhlar\" rejimi YOQILGAN - "
                "tayyor qoidalar ishlamaydi. O'chirish: Guruhlarni boshqarish."
            )
        self.after(100, self._poll_queue)
        threading.Thread(target=send_ping, daemon=True).start()

    def _on_unhandled(self, tur, qiymat, iz):
        matn = "".join(traceback.format_exception(tur, qiymat, iz))[-900:]
        report_error(matn)
        self._log(f"KUTILMAGAN XATO: {qiymat}")
        try:
            messagebox.showerror(
                "Xato",
                f"Kutilmagan xato yuz berdi:\n\n{qiymat}\n\n"
                "Xabar dasturchiga yuborildi. Ilovani qayta ochib ko'ring.",
            )
        except Exception:
            pass

    def _on_root_configure(self, event):
        """Windowsda oyna maximize/restore qilinganda ba'zi ttk widget'lar
        (ayniqsa sv_ttk kabi rasm asosida chiziladigan zamonaviy temalar)
        darhol qayta chizilmay, qora "yamalgan" joylar ko'rinib qolishi
        mumkin — bu Tk/DWM darajasidagi tanish nuqson, ilova mantig'iga
        aloqasi yo'q. Oyna holati (normal/zoomed) chindan o'zgarganda butun
        widget daraxtini majburan qayta chizib, shu nuqsonni oldini olamiz."""
        if event.widget is not self:
            return
        try:
            state = self.state()
        except tk.TclError:
            return
        if state != self._last_state:
            self._last_state = state
            # Darhol chizamiz. Avval bu 50 ms kechikib bajarilardi va aynan
            # o'sha kechikish ko'zga "qora yamoq" bo'lib tashlanardi: oyna
            # kattayib bo'lgan, lekin widget'lar hali chizilmagan bo'lardi.
            self._force_full_redraw()
            # DWM oynani bir necha kadr davomida cho'zib ko'rsatadi, shuning
            # uchun keyingi bo'sh lahzada yana bir marta chizamiz — birinchi
            # chizish animatsiya tugashidan oldin bo'lib qolsa ham yamoq
            # qolib ketmasin.
            self.after_idle(self._force_full_redraw)

    def _force_full_redraw(self):
        # update_idletasks butun ilovadagi kutayotgan chizish ishlarini
        # bajaradi — qaysi widget'dan chaqirilishidan qat'i nazar. Avval
        # bu butun daraxt bo'ylab rekursiv chaqirilardi, ya'ni yuzlab
        # marta takrorlanardi va maximize paytida ko'zga ko'rinarli
        # qotish berardi. Bitta chaqiruv yetarli.
        try:
            self.update_idletasks()
        except tk.TclError:
            pass

    # ---------------------------------------------------------- UI layout
    def _apply_dpi_scaling(self):
        """_enable_dpi_awareness() Windows'ga haqiqiy piksel o'lchamlarini
        ko'rsatishga majbur qiladi; shu real DPI qiymatiga qarab Tk'ning
        ichki masshtabini (scaling) moslaymiz — aks holda widget'lar
        yuqori-DPI monitorlarda juda mayda yoki noto'g'ri o'lchamda
        chizilib, oyna o'lchami o'zgarganda joylashuv buzilib qolishi
        mumkin edi."""
        try:
            dpi = self.winfo_fpixels("1i")
            if dpi > 0:
                self.tk.call("tk", "scaling", dpi / 72.0)
        except Exception:
            pass

    def _setup_fonts(self):
        base = tkfont.nametofont("TkDefaultFont")
        base.configure(family="Segoe UI", size=10)
        self.option_add("*Font", base)
        self.heading_font = tkfont.Font(family="Segoe UI Semibold", size=21)
        self.subtitle_font = tkfont.Font(family="Segoe UI", size=10)
        self.step_font = tkfont.Font(family="Segoe UI Semibold", size=11)
        self.section_font = tkfont.Font(family="Segoe UI Semibold", size=10)
        self.mono_font = tkfont.Font(family="Consolas", size=9)

    def _setup_styles(self):
        """Mavzu ustidan bir nechta o'lcham tuzatishi.

        sv_ttk almashtirilmaydi — faqat sarlavhalar ajralib tursin va
        ro'yxat qatorlari zich bo'lmasin. Mayda, siqilgan interfeysda
        ko'z qayerga qarashni bilmaydi; qatorlar orasidagi bo'sh joy
        chiroylilikdan ko'ra o'qishga yordam beradi."""
        style = ttk.Style()
        style.configure("TLabelframe.Label", font=self.section_font, foreground="#5a5a5a")
        style.configure("TLabelframe", borderwidth=1)
        style.configure("Treeview", rowheight=28)
        style.configure("Treeview.Heading", font=self.section_font)

    def _on_root_resize(self, event):
        # Tavsif matni oyna torayganda so'zma-so'z pastga tushib
        # ("wrap" bo'lib) yozilsin — bitta uzun qatorda kesilib qolmasin.
        #
        # wraplength o'zgarishi label balandligini o'zgartiradi, u esa
        # yangi <Configure> hodisasini keltirib chiqaradi. Kenglik
        # o'zgarmagan bo'lsa hech narsa qilmaymiz — aks holda oyna
        # cho'zilayotganda shu halqa sekundiga o'nlab marta aylanib,
        # ilova "qotgandek" sekinlashardi.
        if abs(event.width - self._last_wrap_width) < 8:
            return
        self._last_wrap_width = event.width
        try:
            self.subtitle_label.configure(wraplength=max(300, event.width - 28))
        except Exception:
            pass

    def _step_frame(self, parent, title):
        """A labeled 'card' section used to break the workflow into clear,
        numbered steps so a non-technical user always knows what's next."""
        frame = ttk.LabelFrame(parent, text=title, padding=16)
        return frame

    def _build_ui(self):
        pad = 18
        root = ttk.Frame(self, padding=pad)
        root.pack(fill="both", expand=True)

        header = ttk.Frame(root)
        header.pack(fill="x", pady=(0, 6))
        title_row = ttk.Frame(header)
        title_row.pack(fill="x")
        title_row.columnconfigure(0, weight=1)
        title_group = ttk.Frame(title_row)
        title_group.grid(row=0, column=0, sticky="w")
        ttk.Label(title_group, text="SoddaHisobot", font=self.heading_font).pack(side="left")
        ttk.Label(
            title_group, text=f"  v{CORE_VERSION}", font=self.subtitle_font, foreground="#999999"
        ).pack(side="left", anchor="s", pady=(0, 3))
        self.subtitle_label = ttk.Label(
            header,
            text="Xom Hamkorbank hisobotlarini tanlang — har biri alohida Excel faylga aylantiriladi.",
            font=self.subtitle_font,
            foreground="#666666",
        )
        self.subtitle_label.pack(anchor="w")
        root.bind("<Configure>", self._on_root_resize)

        # Sarlavhani ish maydonidan ajratib turadigan ingichka chiziq —
        # qadamlar shu chiziqdan pastda boshlanadi.
        ttk.Separator(root, orient="horizontal").pack(fill="x", pady=(12, 16))

        step1 = self._step_frame(root, "1-qadam · Fayllarni tanlang")
        step1.pack(fill="x", pady=(0, 14))
        row1 = ttk.Frame(step1)
        row1.pack(fill="x")
        ttk.Button(row1, text="📂 Fayllarni tanlash...", command=self.pick_files).pack(side="left")
        ttk.Button(row1, text="Ro'yxatni tozalash", command=self.clear_files).pack(side="left", padx=(8, 0))
        ttk.Button(row1, text="🏷 Guruhlarni boshqarish...", command=self.open_groups_manager).pack(side="left", padx=(8, 0))
        self.count_label = ttk.Label(row1, text="0 ta fayl tanlangan", font=self.step_font)
        self.count_label.pack(side="left", padx=(16, 0))

        step2 = self._step_frame(root, "2-qadam · Natijalarni qayerga saqlash")
        step2.pack(fill="x", pady=(0, 14))
        out_frame = ttk.Frame(step2)
        out_frame.pack(fill="x")
        self.out_entry = ttk.Entry(out_frame, textvariable=self.out_dir)
        self.out_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))
        ttk.Button(out_frame, text="💾 Papka tanlash...", command=self.pick_out_dir).pack(side="left")

        # Jurnal va 3-qadam ichidagi tugmalar "pastdan" joylashtiriladi:
        # pack side="bottom" bo'lgan elementlar o'z joyini birinchi bo'lib
        # egallaydi, shuning uchun oyna kichraytirilganda ular kesilib
        # qolmaydi — o'rniga yuqoridagi ro'yxat qisqaradi.
        log_frame = ttk.LabelFrame(root, text="Jurnal", padding=(10, 6))
        log_frame.pack(side="bottom", fill="x", pady=(14, 0))
        self.log = tk.Text(
            log_frame, height=4, wrap="word", state="disabled",
            font=self.mono_font, relief="flat", borderwidth=0,
            background="#f5f5f5" if sv_ttk.get_theme() == "light" else "#1e1e1e",
        )
        self.log.pack(fill="both", expand=True)

        step3 = self._step_frame(root, "3-qadam · Boshlang va kuzating")
        step3.pack(fill="both", expand=True)

        action_row = ttk.Frame(step3)
        action_row.pack(side="bottom", fill="x")
        self.start_btn = ttk.Button(action_row, text="▶  Boshlash", style="Accent.TButton", command=self.start_processing)
        self.start_btn.pack(side="left", ipadx=6)
        self.open_out_btn = ttk.Button(action_row, text="📁 Papkani ochish", command=self.open_out_dir)
        self.open_out_btn.pack(side="left", padx=(8, 0))
        self.open_excel_btn = ttk.Button(action_row, text="📊 Excelda ochish", command=self.open_selected_in_excel)
        self.open_excel_btn.pack(side="left", padx=(8, 0))
        self.status_label = ttk.Label(action_row, text="Tayyor", font=self.step_font)
        self.status_label.pack(side="right")

        self.progress = ttk.Progressbar(step3, orient="horizontal", mode="determinate")
        self.progress.pack(side="bottom", fill="x", pady=(0, 10))

        ttk.Label(
            step3, text="Tayyor bo'lgan faylni ochish uchun ustiga ikki marta bosing.",
            foreground="#888888",
        ).pack(side="bottom", anchor="w", pady=(0, 8))

        list_frame = ttk.Frame(step3)
        list_frame.pack(fill="both", expand=True, pady=(0, 10))

        columns = ("file", "status")
        self.tree = ttk.Treeview(list_frame, columns=columns, show="headings", selectmode="extended", height=4)
        # Ustun sarlavhasi tagidagi matn bilan bir chiziqda tursin —
        # standart holatda sarlavha markazda, matn chapda bo'lib,
        # ro'yxat qiyshiq ko'rinardi.
        self.tree.heading("file", text="Fayl", anchor="w")
        self.tree.heading("status", text="Holati", anchor="w")
        self.tree.column("file", width=600, anchor="w", stretch=True)
        self.tree.column("status", width=160, anchor="w", stretch=False)
        self.tree.pack(side="left", fill="both", expand=True)
        self.tree.bind("<Double-1>", self._on_tree_double_click)

        scroll = ttk.Scrollbar(list_frame, orient="vertical", command=self.tree.yview)
        scroll.pack(side="left", fill="y")
        self.tree.configure(yscrollcommand=scroll.set)

        self.tree.tag_configure("Kutmoqda", foreground=self.STATUS_COLORS["Kutmoqda"])
        self.tree.tag_configure("Ishlanmoqda...", foreground=self.STATUS_COLORS["Ishlanmoqda..."])
        self.tree.tag_configure("Tayyor", foreground=self.STATUS_COLORS["Tayyor"])
        self.tree.tag_configure("Xato", foreground=self.STATUS_COLORS["Xato"])

    # ------------------------------------------------------------ actions
    def pick_files(self):
        paths = filedialog.askopenfilenames(
            title="Xom hisobot fayllarini tanlang",
            filetypes=[("Excel fayllar", "*.xlsx"), ("Barcha fayllar", "*.*")],
        )
        if not paths:
            return
        existing = {f.path for f in self.files}
        for p in paths:
            if p not in existing:
                self.files.append(FileRow(p))
        self._refresh_tree()

    def clear_files(self):
        if self.is_running:
            return
        self.files = []
        self._refresh_tree()

    def open_groups_manager(self):
        dialog = GroupsManagerDialog(self)
        self.wait_window(dialog)

    def pick_out_dir(self):
        d = filedialog.askdirectory(title="Natijalarni saqlash papkasini tanlang")
        if d:
            self.out_dir.set(d)

    def open_out_dir(self):
        d = self.out_dir.get().strip()
        if d and os.path.isdir(d):
            os.startfile(d)
        else:
            messagebox.showinfo("Diqqat", "Avval saqlash papkasini tanlang.")

    def _on_tree_double_click(self, _event):
        sel = self.tree.selection()
        if not sel:
            return
        idx = int(sel[0])
        f = self.files[idx]
        if f.out_path and os.path.exists(f.out_path):
            self._open_path_in_excel(f.out_path)

    def open_selected_in_excel(self):
        sel = self.tree.selection()
        if sel:
            targets = [self.files[int(i)] for i in sel]
        else:
            targets = self.files
        paths = [f.out_path for f in targets if f.out_path and os.path.exists(f.out_path)]
        if not paths:
            messagebox.showinfo(
                "Diqqat",
                "Ochish uchun tayyor natija fayli topilmadi. Avval \"Boshlash\" bilan qayta ishlashni yakunlang.",
            )
            return
        excel_exe = find_excel_exe()
        if not excel_exe:
            messagebox.showwarning(
                "Excel topilmadi",
                "Bu kompyuterda Microsoft Excel ro'yxatdan o'tmagan.\n\n"
                "Fayllar to'g'ri .xlsx formatida saqlangan, lekin ularni ochish uchun "
                "Excel (yoki LibreOffice Calc, WPS Office kabi mos dastur) o'rnatilgan bo'lishi kerak.\n\n"
                "Agar Excel o'rnatilgan bo'lsa-yu, fayl baribir Notepad'da ochilsa: fayl ustida "
                "o'ng tugmani bosing -> \"Open with\" -> Excel -> \"Always use this app\".",
            )
            return
        for p in paths:
            try:
                subprocess.Popen([excel_exe, p])
            except Exception as e:
                self._log(f"XATO: Excelda ochib bo'lmadi ({os.path.basename(p)}): {e}")

    def _open_path_in_excel(self, path):
        excel_exe = find_excel_exe()
        if excel_exe:
            try:
                subprocess.Popen([excel_exe, path])
                return
            except Exception as e:
                self._log(f"XATO: Excelda ochib bo'lmadi ({os.path.basename(path)}): {e}")
        os.startfile(path)

    def _refresh_tree(self):
        self.tree.delete(*self.tree.get_children())
        for i, f in enumerate(self.files):
            self.tree.insert("", "end", iid=str(i), values=(f.path, f.status), tags=(f.status,))
        self.count_label.config(text=f"{len(self.files)} ta fayl tanlangan")

    def _log(self, msg):
        self.log.configure(state="normal")
        self.log.insert("end", msg + "\n")
        # Jurnal cheksiz o'smasin: bir necha marta ishlov berilgandan
        # keyin minglab qator to'planadi va Text widget har yangi qatorda
        # sezilarli sekinlashadi. Faqat oxirgi qatorlarni saqlaymiz —
        # foydalanuvchi baribir oxirini o'qiydi.
        ortiqcha = int(self.log.index("end-1c").split(".")[0]) - 500
        if ortiqcha > 0:
            self.log.delete("1.0", f"{ortiqcha + 1}.0")
        self.log.see("end")
        self.log.configure(state="disabled")

    def start_processing(self):
        if self.is_running:
            return
        if not self.files:
            messagebox.showwarning("Diqqat", "Avval kamida bitta fayl tanlang.")
            return
        out_dir = self.out_dir.get().strip()
        if not out_dir:
            messagebox.showwarning("Diqqat", "Avval saqlash papkasini tanlang.")
            return
        if not os.path.isdir(out_dir):
            try:
                os.makedirs(out_dir, exist_ok=True)
            except Exception as e:
                messagebox.showerror("Xato", f"Papka yaratib bo'lmadi:\n{e}")
                return

        for f in self.files:
            f.status = "Kutmoqda"
            f.error = None
        self._refresh_tree()

        # Oldingi fayllar uchun berilgan "faqat shu safar" javoblari
        # yangi ishga o'tib ketmasin.
        clear_session_overrides()
        self.is_running = True
        self.start_btn.configure(state="disabled")
        self.status_label.configure(text="Tekshirilmoqda...")
        self._log(f"--- Skanerlash boshlandi: {len(self.files)} ta fayl ---")

        t = threading.Thread(target=self._scan_worker, args=(out_dir,), daemon=True)
        t.start()

    def _scan_worker(self, out_dir):
        """Ishlov berishdan oldin: barcha tanlangan fayllarni o'qib, hech
        qanday kategoriyaga to'g'ri kelmagan ("?") kontragentlarni yig'ib
        chiqadi, shu bilan foydalanuvchidan bir marta so'rab olib bo'lgach
        haqiqiy qayta ishlash boshlanadi."""
        unresolved = {}
        for f in self.files:
            try:
                _wb, _ws, _hdr, rows = load_raw_rows(f.path)
            except Exception as e:
                self.ui_queue.put(("log", None, f"Skanerlashda xato ({os.path.basename(f.path)}): {e}", None))
                continue
            for key, info in find_unresolved(rows).items():
                if key not in unresolved:
                    unresolved[key] = dict(info)
                else:
                    unresolved[key]["count"] += info["count"]
        self.ui_queue.put(("scan_done", out_dir, unresolved, None))

    def _handle_scan_done(self, out_dir, unresolved):
        if unresolved:
            self._log(f"{len(unresolved)} ta nomlanmagan/kategoriyalanmagan kontragent topildi.")
            dialog = UnresolvedDialog(self, unresolved)
            self.wait_window(dialog)
            if not dialog.confirmed:
                self._log("Bekor qilindi.")
                self.is_running = False
                self.start_btn.configure(state="normal")
                self.status_label.configure(text="Tayyor")
                return
            assignments = dialog.get_assignments()
            for info, cat in assignments:
                nomi = info.get("name") or info.get("account") or ""
                if info.get("always_ask"):
                    # G'aznachilik kabi kontragent: bugun soliq, ertaga
                    # elektr bo'lishi mumkin. Shuning uchun diskka
                    # yozmaymiz — javob faqat shu faylga tegishli va
                    # keyingi safar yana so'raladi.
                    set_session_override(
                        info.get("account") or "", info.get("purpose") or "", cat
                    )
                    self._log(f"Shu fayl uchun: {nomi[:40]} -> {cat}")
                else:
                    save_learned_category(
                        info.get("account") or "", info.get("name") or "", cat
                    )
                    self._log(f"Saqlandi: {nomi[:40]} -> {cat}")
        else:
            self._log("Nomlanmagan kontragent topilmadi.")

        self.progress.configure(maximum=len(self.files), value=0)
        self.status_label.configure(text="Ishlanmoqda...")
        self._log(f"--- Qayta ishlash boshlandi: {len(self.files)} ta fayl ---")
        t = threading.Thread(target=self._worker, args=(out_dir,), daemon=True)
        t.start()

    def _worker(self, out_dir):
        done_ok = 0
        done_err = 0
        balans_xato = []
        for idx, f in enumerate(self.files):
            self.ui_queue.put(("status", idx, "Ishlanmoqda...", None))
            base = os.path.splitext(os.path.basename(f.path))[0]
            out_path = os.path.join(out_dir, f"{base} - soddalashtirilgan.xlsx")
            try:
                info = build_simplified_report(f.path, out_path)
                f.out_path = out_path
                self.ui_queue.put(("status", idx, "Tayyor", None))
                self.ui_queue.put((
                    "log", None,
                    f"OK: {os.path.basename(f.path)} -> {os.path.basename(out_path)} "
                    f"({info['review_count']} ta tekshirish bandi)",
                    None,
                ))
                farq = info.get("balans_farqi")
                if farq:
                    # Bankning o'z qoldig'i bilan hisoblangan aylanma mos
                    # kelmadi — demak qatorlar tushib qolgan yoki ikki
                    # marta hisoblangan. Jimgina o'tkazib yubormaymiz.
                    balans_xato.append(os.path.basename(f.path))
                    self.ui_queue.put((
                        "log", None,
                        f"DIQQAT: {os.path.basename(f.path)} — bank qoldig'i bilan "
                        f"aylanma mos kelmadi, farq {farq}",
                        None,
                    ))
                done_ok += 1
            except Exception as e:
                err = "".join(traceback.format_exception_only(type(e), e)).strip()
                self.ui_queue.put(("status", idx, "Xato", err))
                self.ui_queue.put(("log", None, f"XATO: {os.path.basename(f.path)} -> {err}", None))
                report_error(f"fayl qayta ishlashda: {err}")
                done_err += 1
            self.ui_queue.put(("progress", idx + 1, None, None))
        self.ui_queue.put(("done", done_ok, done_err, balans_xato))

    def _poll_queue(self):
        # Navbatni oxirigacha bo'shatmaymiz. Ishlov paytida ishchi oqim
        # yuzlab xabar yuboradi (har fayl uchun holat, foiz, jurnal
        # qatorlari) — hammasini bitta tikda chizish asosiy oqimni uzoq
        # band qiladi va ilova qisqa vaqtga qotib qolgandek ko'rinadi.
        # Shuning uchun bir tikda cheklangan miqdorda ishlaymiz; xabar
        # qolgan bo'lsa keyingi tikni darrov rejalashtiramiz, shunda
        # ekran ham yangilanib turadi, ham tez to'ladi.
        navbatda_bor = True
        try:
            for _ in range(50):
                kind, a, b, c = self.ui_queue.get_nowait()
                if kind == "status":
                    idx, status = a, b
                    self.files[idx].status = status
                    self.tree.item(str(idx), values=(self.files[idx].path, status), tags=(status,))
                elif kind == "progress":
                    self.progress.configure(value=a)
                elif kind == "log":
                    self._log(b)
                elif kind == "scan_done":
                    out_dir, unresolved = a, b
                    self._handle_scan_done(out_dir, unresolved)
                elif kind == "done":
                    ok, err = a, b
                    self.is_running = False
                    self.start_btn.configure(state="normal")
                    self.status_label.configure(text="Tayyor")
                    self._log(f"--- Jarayon tugadi: {ok} ta muvaffaqiyatli, {err} ta xato ---")
                    tip = (
                        "\n\nFaylni ochish uchun ustiga ikki marta bosing yoki "
                        "\"Excelda ochish\" tugmasini bosing (agar oddiy ikki marta "
                        "bosish Notepad'da ochsa, bu tugma majburan Excel bilan ochadi)."
                        if ok else ""
                    )
                    if err:
                        messagebox.showwarning(
                            "Tugadi",
                            f"{ok} ta fayl tayyor bo'ldi, {err} ta faylda xato yuz berdi.\nJurnalni tekshiring.{tip}",
                        )
                    else:
                        messagebox.showinfo("Tugadi", f"Barcha {ok} ta fayl muvaffaqiyatli qayta ishlandi.{tip}")
                    # Balans nazorati alohida ogohlantiriladi: fayl "tayyor"
                    # bo'lib ko'rinadi, lekin raqamlariga ishonib bo'lmaydi.
                    if c:
                        messagebox.showwarning(
                            "Balans mos kelmadi",
                            "Quyidagi fayllarda bank ko'rsatgan qoldiq bilan hisoblangan "
                            "aylanma mos kelmadi:\n\n" + "\n".join(f"  • {n}" for n in c) +
                            "\n\nBu odatda ba'zi qatorlar o'qilmaganini bildiradi. "
                            "Natijani ishlatishdan oldin tekshiring.",
                        )
        except queue.Empty:
            navbatda_bor = False
        self.after(10 if navbatda_bor else 100, self._poll_queue)


def _enable_dpi_awareness():
    """Windows monitor masshtabi (125%, 150% va h.k.) turli bo'lganda yoki
    oyna boshqa monitorga ko'chirilganda/kattalashtirilganda elementlar
    joyidan siljib, ustma-ust tushib qolishining asosiy sababi — dastur
    Windows'ga "men DPI-ga moslashaman" deb aytmagani. Shuni tuzatamiz;
    muvaffaqiyatsiz bo'lsa (masalan eski Windows yoki boshqa OS) jim
    tarzda o'tkazib yuboriladi — ilova baribir ishlayveradi."""
    try:
        windll.shcore.SetProcessDpiAwareness(1)  # PROCESS_SYSTEM_DPI_AWARE
    except Exception:
        try:
            windll.user32.SetProcessDPIAware()
        except Exception:
            pass


def _ensure_desktop_shortcut():
    """Ish stolidagi yorliqni HAR SAFAR joriy .exe ga yo'naltiradi.

    Faqat "yo'q bo'lsa yaratish" yetarli emas edi: eski nusxa o'chirilsa
    yoki ilova boshqa papkaga ko'chirilsa, yorliq ishlamay qolardi yoki
    eski faylni ochib, "menda eski versiya" degan chalkashlik berardi.
    Xato chiqsa ilova ishlashda davom etadi (bu shunchaki qulaylik)."""
    if not getattr(sys, "frozen", False):
        return
    try:
        desktop = os.path.join(os.environ.get("USERPROFILE", ""), "Desktop")
        shortcut_path = os.path.join(desktop, "SoddaHisobot.lnk")
        if not os.path.isdir(desktop):
            return
        target = os.path.abspath(sys.executable)
        workdir = os.path.dirname(target)
        ps_script = (
            "$WshShell = New-Object -ComObject WScript.Shell; "
            f'$Shortcut = $WshShell.CreateShortcut("{shortcut_path}"); '
            f'$Shortcut.TargetPath = "{target}"; '
            f'$Shortcut.WorkingDirectory = "{workdir}"; '
            f'$Shortcut.IconLocation = "{target}"; '
            "$Shortcut.Save()"
        )
        subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-Command", ps_script],
            creationflags=subprocess.CREATE_NO_WINDOW,
            timeout=10,
            check=False,
        )
    except Exception:
        pass


def main():
    _enable_dpi_awareness()
    os.chdir(APP_DIR)
    _ensure_desktop_shortcut()
    app = App()
    app.mainloop()
