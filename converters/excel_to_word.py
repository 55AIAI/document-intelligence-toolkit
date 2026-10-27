"""
УНИВЕРСАЛЬНЫЙ ГЕНЕРАТОР ОТЧЁТА: Excel схема → Word документ  (v4.1)
====================================================================
Исправления v4.1 (vs v4):
  1. CLASSIFY: 2-й проход корректно определяет contractor/subcontractor
     без зависимости от порядка обработки (ищет org_pending, not only pending).
  2. link_subs: primary contractor = самый верхний (min from_row), не левый.
  3. find_contract: порог сужен ±8 строк (было ±14), уменьшает захват соседей.
  4. extract_org_data: start_idx инициализируется явно во всех ветках.
  5. _looks_russian: добавлен порог кириллицы >=30% символов.
  6. emit_org: depth-параметр — уровень 2+ получает сокращённый блок.
  7. recs_parsed: ограничение длины пункта 300 символов.
  8. parse_amount_to_tenge: убран abs() — sign сохраняется для detect наценки.
  9. subtitle2 FMT: bold для названия проекта, italic только для авто-темы.
  10. read_shapes: ранняя проверка — предупреждение если shapes пустой.
"""

import sys, os, re, io, json, zipfile, requests, traceback
import xml.etree.ElementTree as ET
from collections import Counter

# Конфиг Qwen/Ollama берётся из корневого config.py (единый источник).
# Fallback — если файл запущен напрямую и config.py не виден в sys.path.
try:
    import os as _os, sys as _sys
    _sys.path.insert(0, _os.path.dirname(_os.path.dirname(_os.path.abspath(__file__))))
    from config import OLLAMA_URL, MODEL
except Exception:
    OLLAMA_URL = "http://localhost:11434"
    MODEL      = "qwen2.5"

NS = {
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "a":   "http://schemas.openxmlformats.org/drawingml/2006/main",
}

# Казахские буквы — добавляем во все классы букв, иначе ФИО/названия с ними
# (Ғ, Қ, Ұ, Ө, Ә, Һ, І, Ң, Ү) теряются при распознавании имён/текста.
KZ  = "әғқңөұүһіӘҒҚҢӨҰҮҺІ"
LET = "а-яёА-ЯЁ" + KZ + "A-Za-z"     # использовать внутри [ ... ]


# ══════════════════════════════════════════════════════════════════════
# ЧАСТЬ 0: НОРМАЛИЗАЦИЯ ТЕКСТА
# ══════════════════════════════════════════════════════════════════════

def _normalize(s):
    if not s: return s
    for c in ('\u00AD','\u200B','\u200C','\u200D','\uFEFF','\u2060'):
        s = s.replace(c, '')
    s = s.replace('\u00A0', ' ').replace('\u202F', ' ')
    s = s.replace('\u2011', '-').replace('\u2012', '-')
    s = s.replace('«', '"').replace('»', '"')
    s = s.replace('\u201C', '"').replace('\u201D', '"')
    s = s.replace('\u2018', "'").replace('\u2019', "'")
    return s

def _clean(s):
    if s is None: return ""
    s = _normalize(str(s))
    for ch in ["\t","\r","\n","\u2003","\u2002"]:
        s = s.replace(ch, " ")
    return re.sub(r"\s{2,}", " ", s).strip()


# ══════════════════════════════════════════════════════════════════════
# ЧАСТЬ 1: ЧТЕНИЕ ВСЕХ ФИГУР
# ══════════════════════════════════════════════════════════════════════

def read_shapes(xlsx_path):
    with open(xlsx_path, "rb") as f:
        data = f.read()
    with zipfile.ZipFile(io.BytesIO(data)) as zf:
        dfiles = sorted([
            n for n in zf.namelist()
            if "drawing" in n.lower() and n.endswith(".xml") and "_rels" not in n
        ])
        if not dfiles:
            raise ValueError("Фигуры не найдены в Excel файле")
        print(f"  Drawing файлов: {len(dfiles)}")
        xml_parts = [zf.read(df).decode("utf-8") for df in dfiles]

    shapes, connectors, sid = [], [], 0
    for xml_data in xml_parts:
        root = ET.fromstring(xml_data)
        for anc in root.findall("xdr:twoCellAnchor", NS):
            try:
                fc = int(anc.find("xdr:from/xdr:col", NS).text)
                fr = int(anc.find("xdr:from/xdr:row", NS).text)
                tc = int(anc.find("xdr:to/xdr:col",   NS).text)
                tr = int(anc.find("xdr:to/xdr:row",   NS).text)
            except (AttributeError, TypeError):
                continue
            sp = anc.find("xdr:sp", NS)
            if sp is not None:
                sid   += 1
                # Собираем параграфы: каждый <a:p> = одна строка
                # Внутри параграфа объединяем все <a:t> (runs) в одну строку
                # Это решает проблему "540" + " млн.тг" в разных runs
                para_ns = "http://schemas.openxmlformats.org/drawingml/2006/main"
                paras   = sp.findall(f".//{{{para_ns}}}p")
                if paras:
                    texts = []
                    for p_el in paras:
                        runs = p_el.findall(f".//{{{para_ns}}}t")
                        joined = "".join(
                            _normalize(r.text) for r in runs if r.text
                        ).strip()
                        if joined:
                            texts.append(joined)
                else:
                    texts = [_normalize(t.text)
                             for t in sp.findall(".//a:t", NS) if t.text]
                texts = [t for t in texts if t and t.strip()]
                bg_el  = sp.find(".//a:solidFill/a:srgbClr", NS)
                bg     = (bg_el.get("val") or "").upper() if bg_el is not None else ""
                ln_el  = sp.find(".//a:ln", NS)
                dash   = (sp.find(".//a:prstDash", NS) is not None or
                          (ln_el is not None and
                           "dash" in ET.tostring(ln_el, encoding="unicode").lower()))
                xfrm   = sp.find(".//a:xfrm", NS)
                rot    = int(xfrm.get("rot","0")) if xfrm is not None else 0
                shapes.append({
                    "id": sid, "text": "\n".join(texts), "lines": [t for t in texts if t.strip()],
                    "from_col": fc, "from_row": fr, "to_col": tc, "to_row": tr,
                    "center_col": (fc+tc)/2, "center_row": (fr+tr)/2,
                    "bg": bg, "dash": dash, "rot": rot, "type": None,
                })
            if anc.find("xdr:cxnSp", NS) is not None:
                connectors.append({"from_col": fc,"from_row": fr,"to_col": tc,"to_row": tr})
            # Картинка (фото руководителя/лого) — сохраняем как ФОТО-рамку:
            # текста нет, но позиция важна (зарезервированное место под фото).
            if anc.find("xdr:pic", NS) is not None:
                sid += 1
                shapes.append({
                    "id": sid, "type": "photo", "text": "", "lines": [],
                    "from_col": fc, "from_row": fr, "to_col": tc, "to_row": tr,
                    "center_col": (fc+tc)/2, "center_row": (fr+tr)/2,
                    "bg": "", "dash": False, "rot": 0,
                })

    print(f"  Фигур: {len(shapes)}, стрелок: {len(connectors)}")
    # FIX 10: ранняя проверка пустого результата
    if not shapes:
        print("  ⚠️  Фигуры не найдены — возможно файл содержит только изображения.")
    return shapes, connectors


# ══════════════════════════════════════════════════════════════════════
# ЧАСТЬ 1b: НОРМАЛИЗАЦИЯ ОРИЕНТАЦИИ (схема может быть снизу-вверх)
# ══════════════════════════════════════════════════════════════════════

# Маркеры организации-заказчика (гос. орган). Направление схемы определяем
# по положению заказчика, а не по «кто выше» — потому что схему могут
# рисовать и сверху-вниз, и снизу-вверх (дерево, сходящееся в корень).
_GOV_MARKERS = re.compile(
    r'аппарат\s+акима|акимат|министерств|ведомств|управлени|департамент|'
    r'\bГУ\b|\bКГУ\b|\bРГУ\b|\bГККП\b|\bКГП\b|\bРГП\b|заказчик',
    re.I)

def _is_org_shape(s):
    t = s.get("text", "")
    return bool("БИН" in t or
                re.match(r'^(ТОО|ГУ|АО|ИП|ООО|ЗАО|НКО|КГП|РГП|ГККП)[\s"«]', t[:6]))

def normalize_orientation(shapes, connectors):
    """
    Если схема нарисована снизу-вверх (заказчик/корень внизу, ветки идут
    вверх), переворачиваем координаты по вертикали. После этого заказчик
    оказывается сверху и вся остальная логика (classify/link_subs, которая
    исходит из «родитель выше ребёнка») работает без изменений.

    Консервативно: переворачиваем ТОЛЬКО при явных признаках инверсии, чтобы
    не сломать обычный случай сверху-вниз (он в большинстве данных).
    """
    orgs = [s for s in shapes if _is_org_shape(s)]
    if len(orgs) < 2:
        return shapes, connectors
    marked = [s for s in orgs if _GOV_MARKERS.search(s.get("text", ""))]
    if not marked:
        return shapes, connectors      # нет надёжного маркера заказчика — не рискуем
    cust = max(marked, key=lambda s: s["from_row"])   # самый нижний гос-орган
    rows = [s["from_row"] for s in orgs]
    rmin, rmax = min(rows), max(rows)
    if rmax - rmin < 5:
        return shapes, connectors      # схема плоская — переворачивать нечего
    frac       = (cust["from_row"] - rmin) / (rmax - rmin)
    orgs_above = [s for s in orgs if s["from_row"] < cust["from_row"] - 2]
    if frac > 0.6 and orgs_above:      # заказчик внизу, а над ним есть организации
        MAX = max(s["to_row"] for s in shapes)
        for s in shapes:
            nfr, ntr = MAX - s["to_row"], MAX - s["from_row"]
            s["from_row"], s["to_row"] = nfr, ntr
            s["center_row"] = (nfr + ntr) / 2
        for c in connectors:
            c["from_row"], c["to_row"] = MAX - c["to_row"], MAX - c["from_row"]
        print("  🔄 Схема снизу-вверх — нормализована к виду сверху-вниз")
    return shapes, connectors


# ══════════════════════════════════════════════════════════════════════
# ЧАСТЬ 2: КЛАССИФИКАЦИЯ
# ══════════════════════════════════════════════════════════════════════

def classify(shapes, connectors=None):
    # org = фигура с БИН или начинается с типа организации
    org_shapes = [s for s in shapes if (
        "БИН" in s["text"]
        or re.match(r'^(ТОО|ГУ|АО|ИП|ООО|ЗАО|НКО|КГП|РГП|ГККП)[\s"«]', s["text"][:6])
    ) and s["lines"]]
    cust_shape = min(org_shapes, key=lambda s: s["from_row"]) if org_shapes else None
    cust_id    = cust_shape["id"] if cust_shape else None
    cust_row   = cust_shape["from_row"] if cust_shape else 0

    RED_COLORS    = {"FF0000","FFE4E1","FFCCCC","FFC7CE","FF4444","FF6B6B","FFAAAA","FFB3B3","FF9999","FF8080","FF5555"}
    # "Общая сумма: " без числа → НЕ contract. Требуем цифру перед единицей
    CONTRACT_KEYS = re.compile(
        r'\d[\d\s]*(?:[.,]\d+)?\s*(?:млрд\.?|млн\.?|тыс\.?|тг|тенге)',
        re.I)

    for s in shapes:
        if s.get("type") == "photo":     # фото-рамка (xdr:pic) — не классифицируем
            continue
        t, bg, fr = s["text"], s.get("bg") or "", s["from_row"]
        is_header_text = ("ФИО" in t or "ИИН" in t or re.search(r'руководитель|директор|ф\.и\.о', t, re.I))

        # ПРИОРИТЕТ:
        # 1. header (руководитель)
        # 2. title (выше заказчика, любой фон — включая красный заголовок проекта)
        # 3. contract (есть сумма — даже если красный фон, это метка договора!)
        # 4. violation (красный фон, нет суммы)
        # 5. customer / org_pending / unknown
        if s["dash"] or is_header_text:
            s["type"] = "header"

        # title = строго выше заказчика И без БИН (любой цвет)
        elif fr < cust_row and "БИН" not in t:
            s["type"] = "title"

        # contract: повёрнутый блок ИЛИ содержит сумму/описание договора
        # Исключаем: нарушения, акты, тексты без суммы
        elif (s.get("rot", 0) != 0  # повёрнутый блок = договорная метка
              or (CONTRACT_KEYS.search(t)
                  and "БИН" not in t
                  and not re.match(r'^(ТОО|ГУ|АО|ИП|ООО|ЗАО|НКО|КГП|РГП|ГККП)[\s"«]', t[:6])
                  and "ФИО" not in t and "ИИН" not in t
                  and not re.search(r'отсутствует|нарушен|не соответ|согласно акт', t, re.I))):
            s["type"] = "contract"

        # violation = красный фон, НЕТ суммы, ниже заказчика
        elif (bg in RED_COLORS or (bg.startswith("FF") and len(bg)==6
              and int(bg[2:4],16)<80 and int(bg[4:6],16)<80)) and fr > cust_row:
            s["type"] = "violation"

        elif s["id"] == (cust_id or -1):
            # Заказчик — всегда customer, даже если в блоке есть сумма
            s["type"] = "customer"
        elif ("БИН" in t
              or re.match(r'^(ТОО|ГУ|АО|ИП|ООО|ЗАО|НКО|КГП|РГП|ГККП)[\s"«]', t[:6])):
            s["type"] = "org_pending"
        else:
            s["type"] = "unknown"

    # 2-й проход — contractor vs subcontractor
    # Приоритет: если есть стрелка ОТ ЗАКАЗЧИКА к этой org → contractor
    # Иначе: смотрим на позицию (org выше в той же колонке → subcontractor)
    pending     = [s for s in shapes if s["type"] == "org_pending"]
    pending_ids = {s["id"] for s in pending}
    shape_by_id_tmp = {s["id"]: s for s in shapes}

    # Кто получает стрелку напрямую от заказчика?
    direct_from_customer = set()
    if connectors and cust_shape:
        for conn in connectors:
            # Стрелка начинается рядом с заказчиком
            cust_r = cust_shape["from_row"]
            cust_b = cust_shape["to_row"]
            cust_c = cust_shape["center_col"]
            if cust_r - 3 <= conn["from_row"] <= cust_b + 8:
                # Ищем org_pending куда указывает стрелка
                for o in pending:
                    if abs(o["from_row"] - conn["to_row"]) <= 8:
                        direct_from_customer.add(o["id"])

    for o in pending:
        if o["id"] in direct_from_customer:
            # Прямая стрелка от заказчика → это contractor
            o["type"] = "contractor"
        else:
            # Нет стрелки — смотрим на позицию
            above_same_col = [p for p in pending
                              if p["id"] != o["id"]
                              and p["from_row"] < o["from_row"] - 1
                              and abs(p["center_col"] - o["center_col"]) <= 8]
            above_customer = (cust_shape is not None
                              and abs(cust_shape["center_col"] - o["center_col"]) <= 8
                              and cust_shape["from_row"] < o["from_row"] - 1)
            o["type"] = "subcontractor" if (above_same_col or above_customer) else "contractor"

    return shapes


# ══════════════════════════════════════════════════════════════════════
# ЧАСТЬ 3: ИЗВЛЕЧЕНИЕ ДАННЫХ
# ══════════════════════════════════════════════════════════════════════

def _is_label(s):
    """Строка вида «Название поля: ...» — начало нового поля (не продолжение)."""
    return bool(re.match(r'^[^\d:]{2,40}:', (s or "").strip()))

def extract_org_data(lines):
    data = {"name": "", "bin": None, "activity": None, "legal_address": None}
    if not lines: return data

    ORG_TYPES = ["ТОО","ГУ","АО","ИП","ООО","ЗАО","НКО","КГП","РГП","ГККП"]
    SKIP_KW   = ["БИН","ФИО","ИИН","ТЕЛ","АДР","ВИД","ЮР","РУК"]
    BIN_RE    = re.compile(r'БИН[:\s]*(\d{9,12})', re.I)
    BIN_BARE  = re.compile(r'^\d{9,12}$')
    AMT_RE    = re.compile(r'\d[\d\s]*(?:[.,]\d+)?\s*(?:млрд|млн|тыс|тг|тенге)', re.I)

    # ── Объединяем разбитое имя: ['ТОО"', 'Demo Company"', ...] → 'ТОО"Demo Company"' ──
    merged = list(lines)
    if len(merged) >= 2:
        f0 = merged[0].strip()
        f1 = merged[1].strip()
        is_org_start = re.match(r'^(ТОО|ГУ|АО|ИП|ООО|ЗАО|НКО|КГП|РГП|ГККП)[\s"«]', f0)
        no_close_q   = not re.search(r'["\u00bb]\s*$', f0)
        not_skip     = not any(k in f1.upper() for k in SKIP_KW)
        not_bin      = not BIN_RE.search(f1)
        not_amount   = not AMT_RE.search(f1)
        if is_org_start and no_close_q and not_skip and not_bin and not_amount:
            merged = [f0 + f1] + merged[2:]

    first       = merged[0].strip()
    first_clean = re.sub(r'["\u201c\u201d\u00ab\u00bb\u2018\u2019]+', '', first).strip()
    first_key   = first_clean if first_clean in ORG_TYPES else first
    start_idx   = 1

    if first_key in ORG_TYPES:
        org_type   = re.sub(r'["\u2019\u201c\u201d\u00ab\u00bb\u2018]+$', '', first.strip())
        name_parts = []
        for j in range(1, min(6, len(merged))):
            nxt = merged[j].strip()
            if not nxt: continue
            if any(k in nxt.upper() for k in SKIP_KW) or BIN_RE.search(nxt) or AMT_RE.search(nxt):
                start_idx = j; break
            part = re.sub(r'^["\u2019\u201c\u201d\u00ab\u00bb\u2018]+|["\u2019\u201c\u201d\u00ab\u00bb\u2018]+$', '', nxt).strip()
            if part: name_parts.append(part)
            start_idx = j + 1
        data["name"] = f'{org_type} "{" ".join(name_parts)}"' if name_parts else org_type
    elif re.match(r'^(ТОО|ГУ|АО|ИП|ООО|ЗАО|НКО|КГП|РГП|ГККП)', first):
        data["name"] = re.split(r'\s+БИН', first, flags=re.I)[0].strip()
        m = BIN_RE.search(first)
        if m: data["bin"] = m.group(1)
        start_idx = 1
    else:
        data["name"] = first
        start_idx = 1

    act_parts, addr_parts, extra = [], [], []
    i = start_idx
    # Используем merged (не lines)
    work_lines = merged
    while i < len(work_lines):
        line = work_lines[i].strip(); low = line.lower(); i += 1
        if not line: continue
        # Пропускаем строки с суммами — они относятся к договору, не к org
        if AMT_RE.search(line) and "БИН" not in line: continue
        if "БИН" in line or "бин" in low:
            m = BIN_RE.search(line)
            if m: data["bin"] = m.group(1)
            elif i < len(work_lines) and BIN_BARE.match(work_lines[i].strip()):
                data["bin"] = work_lines[i].strip(); i += 1
        elif "виды деятельности" in low or re.search(r'деятельност', low):
            val = line.split(":", 1)[-1].strip() if ":" in line else ""
            if val: act_parts.append(val)
            while i < len(work_lines):
                nxt = work_lines[i].strip()
                if not nxt or any(k in nxt.upper() for k in SKIP_KW) or _is_label(nxt): break
                act_parts.append(nxt); i += 1
        elif ("юр.адрес" in low or "юридический адрес" in low or
              ("адрес" in low and "прописки" not in low and "руководитель" not in low)):
            val = line.split(":", 1)[-1].strip() if ":" in line else ""
            if val: addr_parts.append(val)
            while i < len(work_lines):
                nxt = work_lines[i].strip()
                if (not nxt or any(k in nxt.upper() for k in ["БИН","ФИО","ИИН","ТЕЛ","ВИД"])
                        or _is_label(nxt)): break
                addr_parts.append(nxt); i += 1
        else:
            # Любая прочая строка блока — доп. информация (email, дата
            # регистрации, учредитель и т.п.). Не теряем её. Отсекаем только
            # строки-метки руководителя/БИН по НАЧАЛУ строки (не по подстроке,
            # иначе «учредиТЕЛь» ложно ловится на «ТЕЛ»).
            if len(line) >= 3 and not re.match(
                    r'^\s*(ФИО|ИИН|тел|руковод|адрес\s+прописк|бин)\b', line, re.I):
                extra.append(line)

    if act_parts:  data["activity"]      = ", ".join(act_parts)
    if addr_parts: data["legal_address"] = " ".join(addr_parts)
    if extra:      data["extra"]         = extra
    data["name"] = re.sub(r'\s+БИН\s*$', '', data["name"], flags=re.I).strip()
    data["name"] = re.sub(r'\s+', ' ', data["name"]).strip()
    return data


def extract_header_data(lines):
    """
    Универсальное извлечение данных руководителя. Понимает разные форматы:
      • «ФИО: Иванов И.И.» и просто «Иванов Иван Иванович» (без метки);
      • «ИИН: 12цифр» и «ИИН 12цифр» (без двоеточия);
      • заголовки «Руководитель», «Руководитель (с ДАТА):», должности «… по …:»;
      • казахские буквы в ФИО (Ғ, Қ, Ұ, Ө, Ә, Һ, І, Ң, Ү).
    """
    data = {"fio": None, "iin": None, "phone": None, "address": None}
    # порядок важен: iin/phone/address ищем раньше fio (у 'фио' короткий ключ)
    FIELD_KEYS = [
        ("iin",     ["иин", "и.и.н"]),
        ("phone",   ["тел.номер", "тел номер", "телефон", "тел.", "тел:", "тел "]),
        ("address", ["адрес прописки", "адрес проживания", "юр.адрес", "адрес", "мекен"]),
        ("fio",     ["фио", "ф.и.о"]),
    ]
    def is_phone(s):
        c = re.sub(r"[\s\-\(\)\+]", "", s)
        return bool(re.match(r"^[78]\d{9,10}$", c)) or bool(re.search(r"\+7|8\s*\(", s))
    def is_iin(s): return bool(re.fullmatch(r"\d{12}", s.strip()))
    def is_addr(s): return any(w in s.lower() for w in
        ["г.", "ул.", "пр.", "мкр", "город", "область", "район", "дом", "кв", "уч.", "мекен", "көше", "қала"])
    def is_name(s):
        w = s.strip().split()
        return (2 <= len(w) <= 4
                and all(re.match(rf"^[{LET}][{LET}\.\-']*$", x) for x in w)
                and not re.search(r'руковод|директ|тел|адрес|иин|фио|отсут|нару|соглас', s, re.I))

    expecting = None
    for raw in lines:
        s = raw.strip(); low = s.lower()
        if not s: continue

        # заголовок блока (не значение): «Руководитель…», должность, роль «…:»
        is_heading = bool(re.match(
            r'^(руководитель|директор|и\.?о\.?\b|заместитель|главный|начальник|представитель|учредитель)', low)) \
            or (s.endswith(':') and not re.search(r'\d', s)
                and not any(v in low for _, vs in FIELD_KEYS for v in vs))

        # какое поле упомянуто в строке
        field = kw = None
        for f, vs in FIELD_KEYS:
            for v in vs:
                if v in low: field, kw = f, v; break
            if field: break

        if field:
            idx = low.find(kw) + len(kw)
            val = s[idx:].lstrip(" :.-–—").strip()
            if field == "iin":
                m = re.search(r'\d{12}', s); val = m.group(0) if m else val
            if val and len(val) > 1 and not (field == "iin" and not re.search(r'\d', val)):
                if data[field] is None: data[field] = val
                expecting = None
            else:
                expecting = field
            continue

        if is_heading:
            expecting = None; continue

        if expecting:
            if data[expecting] is None: data[expecting] = s
            expecting = None; continue

        # строка без метки — эвристики по содержимому
        if len(s) <= 1 or s in (":", "—", "-", "нет", "Нет"): continue
        if   data["iin"]   is None and is_iin(s):   data["iin"]   = s
        elif data["phone"] is None and is_phone(s): data["phone"] = s
        elif data["fio"]   is None and is_name(s):  data["fio"]   = s
        elif data["address"] is None and is_addr(s):data["address"] = s
    return data


def extract_contract_data(lines):
    """
    Разделяет строки на description (ТРУ) и amount (сумма).
    Обрабатывает случаи когда число и единица в разных строках
    ('540' + 'млн.тг - 2024г.' → '540 млн.тг - 2024г.').
    """
    AMT_RE  = re.compile(
        r'([~≈]?\s*-?\s*\d[\d\s]*(?:[.,]\d+)?\s*(?:млрд\.?|млн\.?|тыс\.?|тг|тенге)\b.*)',
        re.I)
    UNIT_RE = re.compile(r'^\s*(?:млрд\.?|млн\.?|тыс\.?|тг|тенге)', re.I)
    NUM_RE  = re.compile(r'^\s*\d[\d\s,.]*\s*$')
    YEAR_RE = re.compile(r'\b(20\d{2})\b')

    # Объединяем строки: только число + единица → одна строка
    merged = []
    i = 0
    clean = [l.strip() for l in lines if l.strip()]
    while i < len(clean):
        line = clean[i]
        if NUM_RE.match(line) and i+1 < len(clean) and UNIT_RE.match(clean[i+1]):
            merged.append(line + " " + clean[i+1])
            i += 2
        else:
            merged.append(line)
            i += 1

    data = {"description": "", "amount": None, "year": None}
    desc_parts = []
    amt_parts  = []

    for line in merged:
        m = AMT_RE.search(line)
        if m and "БИН" not in line:
            before = line[:m.start()].strip().rstrip(',;:- ')
            amount_part = line[m.start():].strip()
            if before:
                desc_parts.append(before)
            amt_parts.append(amount_part)
            yr = YEAR_RE.search(amount_part)
            if yr: data["year"] = yr.group(1)
        else:
            desc_parts.append(line)

    raw_dsc = " ".join(desc_parts).strip()
    raw_dsc = re.sub(r'^(Общая\s+сумма|Сумма)\s*:?\s*', '', raw_dsc)
    # Убираем хвосты: "за сумму", "на сумму", " за " в конце
    raw_dsc = re.sub(r'\s+(за\s+сумму|на\s+сумму|за\s*$)', '', raw_dsc, flags=re.I)
    raw_dsc = raw_dsc.strip().rstrip(' .,;:-')
    data["description"] = raw_dsc
    if amt_parts:
        data["amount"] = " ".join(amt_parts)
    return data


def extract_violation_text(lines):
    parts = [re.sub(r'[⚠️✕✖✗✘×]+','',l).strip() for l in lines]
    return " ".join(p for p in parts if p)


# ══════════════════════════════════════════════════════════════════════
# ЧАСТЬ 4: ПРИВЯЗКИ
# ══════════════════════════════════════════════════════════════════════

def link_headers(shapes):
    orgs    = [s for s in shapes if s["type"] in ("customer","contractor","subcontractor")]
    headers = [s for s in shapes if s["type"] == "header"]
    links   = {}
    for h in headers:
        cands = [(o, h["from_row"]-o["to_row"], abs(o["center_col"]-h["center_col"]))
                 for o in orgs
                 if -5 <= h["from_row"]-o["to_row"] <= 70
                 and abs(o["center_col"]-h["center_col"]) < 18]
        if cands:
            # Приоритет: организация, чья нижняя граница ПРЯМО над header
            # (мин. вертикальный зазор), затем — ближайшая по колонке.
            # Иначе в одной колонке header уводит верхняя org с тем же центром.
            links[h["id"]] = min(cands, key=lambda x:(abs(x[1]), x[2]))[0]["id"]
    return links


def link_subs(shapes, connectors):
    orgs        = [s for s in shapes if s["type"] in ("contractor","subcontractor")]
    subs        = [s for s in shapes if s["type"] == "subcontractor"]
    contractors = [s for s in shapes if s["type"] == "contractor"]

    def match(col, row, pool, prefer_bottom):
        if not pool: return None
        def d(s):
            ref = s["to_row"] if prefer_bottom else s["from_row"]
            return abs(s["center_col"]-col)*2 + abs(ref-row)
        best = min(pool, key=d)
        return best if abs(best["center_col"]-col) <= 12 else None

    arrow_parent = {}
    for conn in connectors:
        src = match(conn["from_col"],conn["from_row"],orgs,True)
        dst = match(conn["to_col"],  conn["to_row"],  subs,False)
        if src and dst and src["id"]!=dst["id"] and src["from_row"]<dst["from_row"]:
            arrow_parent[dst["id"]] = src["id"]

    shape_by_id = {s["id"]: s for s in shapes}
    links = {}

    # Стрелки: ищем src/dst по пересечению строк (не по center_row)
    # Стрелка row46→58: src = org у которой to_row >= 46-8 (т.е. Demo Company to_row=39 — близко)
    for conn in connectors:
        fr, tr = conn["from_row"], conn["to_row"]
        fc, tc = conn["from_col"], conn["to_col"]
        cc = (fc + tc) / 2  # центральная колонка стрелки

        # Источник: org чья нижняя граница (to_row) ближайшая к fr стрелки
        src_cands = [o for o in orgs
                     if abs(o["to_row"] - fr) <= 12
                     and abs(o["center_col"] - cc) <= 14]
        # Получатель: org чья верхняя граница (from_row) ближайшая к tr стрелки
        dst_cands = [o for o in subs
                     if abs(o["from_row"] - tr) <= 10
                     and abs(o["center_col"] - cc) <= 14]

        if src_cands and dst_cands:
            src = min(src_cands, key=lambda o: abs(o["to_row"] - fr))
            dst = min(dst_cands, key=lambda o: abs(o["from_row"] - tr))
            if src["id"] != dst["id"]:
                arrow_parent[dst["id"]] = src["id"]

    # Остальные по вертикальной близости (сверху вниз)
    for sub in sorted(subs, key=lambda s: s["from_row"]):
        if sub["id"] in links:
            continue
        col   = sub["center_col"]
        above = [o for o in orgs
                 if o["id"] != sub["id"]
                 and o["from_row"] < sub["from_row"] - 1
                 and abs(o["center_col"] - col) <= 12]
        if above:
            nearest = max(above, key=lambda o: o["from_row"])
            links[sub["id"]] = nearest["id"]
        elif contractors:
            primary = min(contractors, key=lambda c: (c["from_row"], c["center_col"]))
            links[sub["id"]] = primary["id"]

    return links


def link_violations(shapes):
    orgs  = [s for s in shapes if s["type"] in ("customer","contractor","subcontractor")]
    viols = [s for s in shapes if s["type"] == "violation"]
    links = {}
    for v in viols:
        cands = [o for o in orgs
                 if abs(o["from_row"]-v["from_row"]) < 80
                 and abs(o["center_col"]-v["center_col"]) < 28]
        if cands:
            best = min(cands, key=lambda o: abs(o["center_col"]-v["center_col"])*3+abs(o["from_row"]-v["from_row"]))
            links[v["id"]] = best["id"]
    return links


def link_contracts(shapes):
    """
    Привязывает договорной блок к исполнителю.
    Два вида меток:
    1. Повёрнутые (rot!=0) — боковые метки для левых/правых исполнителей.
       Привязываются по ГОРИЗОНТАЛЬНОЙ близости к ближайшей org.
    2. Вертикальные (rot==0) — метки над/под org в цепочке.
       Привязываются по ВЕРТИКАЛЬНОЙ близости.
    """
    executors = [s for s in shapes if s["type"] in ("contractor","subcontractor")]
    contracts = [s for s in shapes if s["type"] == "contract"]
    links     = {}

    for ct in contracts:
        is_rotated = ct.get("rot", 0) != 0

        if is_rotated:
            # Повёрнутая метка — привязываем к ближайшей org по горизонтали
            # (метка слева от org или справа)
            cands = [e for e in executors
                     if abs(e["center_row"] - ct["center_row"]) < 30  # примерно та же высота
                     ]
            if cands:
                best = min(cands, key=lambda e: abs(e["center_col"] - ct["center_col"]))
                links[ct["id"]] = best["id"]
        else:
            # Вертикальный блок — привязываем к ближайшей org выше по вертикали
            # (контракт стоит МЕЖДУ двумя org — относится к верхней)
            col_tol = 10
            above = [e for e in executors
                     if abs(e["center_col"] - ct["center_col"]) < col_tol
                     and e["to_row"] <= ct["from_row"] + 5]  # org заканчивается до контракта
            if above:
                # Ближайшая org ВЫШЕ
                best = max(above, key=lambda e: e["to_row"])
                links[ct["id"]] = best["id"]
            else:
                # Если ничего выше — берём ближайшую ниже
                below = [e for e in executors
                         if abs(e["center_col"] - ct["center_col"]) < col_tol
                         and e["from_row"] >= ct["to_row"] - 5]
                if below:
                    best = min(below, key=lambda e: e["from_row"])
                    links[ct["id"]] = best["id"]

    return links


# ══════════════════════════════════════════════════════════════════════
# ЧАСТЬ 5: СБОРКА СТРУКТУРЫ
# ══════════════════════════════════════════════════════════════════════

def build_full_structure(shapes, connectors):
    by_type        = {}
    for s in shapes: by_type.setdefault(s["type"],[]).append(s)
    header_links   = link_headers(shapes)
    sub_links      = link_subs(shapes, connectors)
    viol_links     = link_violations(shapes)
    contract_links = link_contracts(shapes)

    def reverse(d):
        rev = {}
        for cid,pid in d.items(): rev.setdefault(pid,[]).append(cid)
        return rev

    org_to_headers   = reverse(header_links)
    org_to_viols     = reverse(viol_links)
    org_to_contracts = reverse(contract_links)
    shape_by_id      = {s["id"]:s for s in shapes}
    get_shape        = lambda sid: shape_by_id.get(sid)

    edges = []  # (from_shape_id, to_shape_id, edge_type) — раньше нигде не сохранялось

    def geom(shape):
        return {"id": shape["id"], "from_col": shape["from_col"], "from_row": shape["from_row"],
                "to_col": shape["to_col"], "to_row": shape["to_row"]}

    titles  = by_type.get("title",[])
    project = " ".join(s["text"].strip() for s in sorted(titles,key=lambda s:s["from_row"])) if titles else ""
    # Убираем префикс «Проект» (он добавляется при рендере)
    project = re.sub(r'^(Проект\s*)', '', project, flags=re.I).strip()
    # Если осталась пустая строка — название не найдено (нормально)
    project = project.strip()

    customers = by_type.get("customer",[])
    if not customers:
        all_orgs = [s for s in shapes if "БИН" in s["text"]]
        if all_orgs: customers=[min(all_orgs,key=lambda s:s["from_row"])]

    customer_data = None
    if customers:
        cs = customers[0]
        customer_data = {**extract_org_data(cs["lines"]),"header":None,"_geom":geom(cs)}
        for hid in org_to_headers.get(cs["id"],[]):
            hs = get_shape(hid)
            if hs: customer_data["header"]=extract_header_data(hs["lines"]); break

    # children[parent_id] = [sub_shapes]
    # Каждый субподрядчик привязан к своему прямому родителю (не root!)
    children        = {}
    _used_contracts = set()
    for ss in by_type.get("subcontractor", []):
        pid = sub_links.get(ss["id"])
        if pid is not None:
            children.setdefault(pid, []).append(ss)
            edges.append((pid, ss["id"], "subcontract"))

    def find_contract(shape):
        for cid in org_to_contracts.get(shape["id"],[]):
            if cid in _used_contracts: continue
            cs = get_shape(cid)
            if cs: _used_contracts.add(cid); return extract_contract_data(cs["lines"])
        # FIX 3: порог сужен ±8 строк (было ±14)
        cands = [s for s in by_type.get("contract",[])
                 if s["id"] not in _used_contracts
                 and shape["from_row"]-8 <= s["from_row"] <= shape["to_row"]+8
                 and abs(s["center_col"]-shape["center_col"]) < 12]
        if cands:
            best = min(cands,key=lambda s:(abs(s["center_col"]-shape["center_col"]),abs(s["from_row"]-shape["from_row"])))
            _used_contracts.add(best["id"])
            return extract_contract_data(best["lines"])
        return None

    def build_node(shape):
        org = extract_org_data(shape["lines"])
        hdr = None
        for hid in org_to_headers.get(shape["id"],[]):
            hs = get_shape(hid)
            if hs: hdr=extract_header_data(hs["lines"]); break
        ct  = find_contract(shape)
        vls = []
        for vid in org_to_viols.get(shape["id"],[]):
            vs = get_shape(vid)
            if vs:
                tx = extract_violation_text(vs["lines"])
                if tx: vls.append(tx)
        kids = sorted(children.get(shape["id"],[]),key=lambda s:s["from_row"])
        return {**org,"header":hdr,"contract":ct,"violations":vls,
                "_geom":geom(shape), "_id":shape["id"],
                "subcontractors":[build_node(k) for k in kids]}

    contractor_shapes = sorted(by_type.get("contractor",[]),key=lambda s:s["center_col"])
    contractors_data  = [build_node(cs) for cs in contractor_shapes]
    if customer_data is not None:
        for cs in contractor_shapes:
            edges.append((customer_data["_geom"]["id"], cs["id"], "direct"))

    global_viols = []
    for vs in by_type.get("violation",[]):
        if vs["id"] not in viol_links:
            tx = extract_violation_text(vs["lines"])
            if tx: global_viols.append(tx)

    return {"project_name":project,"customer":customer_data,
            "contractors":contractors_data,"violations":global_viols,
            "_edges":edges}


def build_editor_state(shapes, connectors, project=""):
    """
    Строит состояние редактора (blocks+edges) НАПРЯМУЮ из фигур и коннекторов,
    без сведения в дерево. Поэтому сохраняется ВСЁ: несколько договоров и
    несколько людей на организацию, связи человек↔человек, точные позиции.
    Возвращает dict вида {"project_name":..., "_editor":{"blocks":[...],"edges":[...]}}
    — его понимает writers.graph_html.build_html (грузит DATA._editor как есть).
    """
    PX_C, PX_R = 70, 34          # крупнее шаг → блоки не наезжают друг на друга
    AMT = re.compile(r'\d[\d\s.,]*\s*(?:млрд|млн|тыс|тг|тенге)', re.I)
    ROLE = {"customer":"cust","contractor":"con","subcontractor":"sub",
            "header":"person","contract":"contract","violation":"violation"}
    blocks, kept = [], set()
    for s in shapes:
        if s["type"] == "title":
            continue
        x, y = s["from_col"]*PX_C, s["from_row"]*PX_R
        w = max(90, (s["to_col"]-s["from_col"])*PX_C)
        h = max(100, (s["to_row"]-s["from_row"])*PX_R)
        if s["type"] == "photo":                      # фото-рамка → слот под фото
            blocks.append({"id": f"s{s['id']}", "type": "photo", "role": "photo",
                           "x": x, "y": y, "w": max(120, w), "h": max(130, h), "photo": None})
            kept.add(s["id"]); continue
        lines = [l for l in (s.get("lines") or []) if l and l.strip()]
        if s["type"] in (None, "unknown") or not lines:
            continue
        role = ROLE.get(s["type"], "shape")
        w = max(170, w)
        b = {"id": f"s{s['id']}", "x": x, "y": y, "w": w, "role": role}
        if s["type"] in ("customer","contractor","subcontractor","header"):
            b["type"] = "org"; b["minh"] = h
            b["lines"] = [{"k": ("name" if i==0 else
                                 ("sub" if re.match(r'^(руковод|директ)', l, re.I) else "fld")),
                           "t": l} for i, l in enumerate(lines)]
        elif s["type"] == "contract":
            amt = next((l for l in lines if AMT.search(l)), "")
            desc = " ".join(l for l in lines if l != amt)
            b["type"] = "contract"; b["fields"] = {"desc": desc, "amt": amt}
        else:  # violation и прочее
            b["type"] = "text"; b["minh"] = h; b["fields"] = {"text": " ".join(lines)}
        blocks.append(b); kept.add(s["id"])

    # связи: для каждого коннектора — блоки, ближайшие к его концам (углам bbox)
    by_id = {s["id"]: s for s in shapes}
    def nearest(col, row):
        best = None
        for sid in kept:
            s = by_id[sid]
            cx, cy = (s["from_col"]+s["to_col"])/2, (s["from_row"]+s["to_row"])/2
            d = (cx-col)**2 + (cy-row)**2
            if best is None or d < best[0]:
                best = (d, sid)
        return best[1] if best else None
    edges = []
    for i, c in enumerate(connectors):
        a = nearest(c["from_col"], c["from_row"])
        z = nearest(c["to_col"],   c["to_row"])
        if a is not None and z is not None and a != z:
            edges.append({"id": f"e{i}", "from": f"s{a}", "to": f"s{z}"})

    # Связи по близости: в схеме многое соединено расположением, а не стрелкой.
    # Договор/лицо/нарушение без связи → ближайшая организация; фото → ближайшее лицо.
    linked = set()
    for e in edges:
        linked.add(e["from"]); linked.add(e["to"])
    pos = {b["id"]: (b["x"] + b.get("w", 170) / 2,
                     b["y"] + (b.get("h") or b.get("minh") or 100) / 2) for b in blocks}
    def near_role(bid, roles):
        pool = [b["id"] for b in blocks if b["role"] in roles and b["id"] != bid]
        if not pool: return None
        bx, by = pos[bid]
        return min(pool, key=lambda p: (pos[p][0]-bx)**2 + (pos[p][1]-by)**2)
    eid = len(connectors) + 1
    for b in blocks:
        if b["id"] in linked:
            continue
        tgt = None
        if b["role"] in ("contract", "person", "violation"):
            tgt = near_role(b["id"], ("cust", "con", "sub"))
        elif b["role"] == "photo":
            tgt = near_role(b["id"], ("person", "con", "sub", "cust"))
        if tgt:
            edges.append({"id": f"e{eid}", "from": tgt, "to": b["id"]}); eid += 1
            linked.add(b["id"])
    return {"project_name": project, "_editor": {"blocks": blocks, "edges": edges}}


def flatten_graph(structure):
    """
    Строит плоскую графовую модель (nodes+edges с явными geometry) из вложенной
    структуры build_full_structure() — именно эта модель встраивается в .docx,
    чтобы word_to_excel мог восстановить схему по координатам, а не заново
    угадывать расположение из текста.
    """
    nodes = []

    def add_node(org, kind):
        g = org.get("_geom") or {}
        nodes.append({
            "id": str(g.get("id", id(org))),
            "type": kind,
            "name": org.get("name"),
            "bin": org.get("bin"),
            "activity": org.get("activity"),
            "address": org.get("address"),
            "amount": (org.get("contract") or {}).get("amount") if kind != "customer" else None,
            "year": (org.get("contract") or {}).get("year") if kind != "customer" else None,
            "description": (org.get("contract") or {}).get("description") if kind != "customer" else None,
            "director": org.get("header"),
            "violations": org.get("violations") or [],
            "geometry": {k: v for k, v in g.items() if k != "id"},
        })
        for kid in org.get("subcontractors") or []:
            add_node(kid, "subcontractor")

    cust = structure.get("customer")
    if cust:
        add_node(cust, "customer")
    for con in structure.get("contractors") or []:
        add_node(con, "contractor")

    edges = [{"from": str(a), "to": str(b), "type": t} for a, b, t in (structure.get("_edges") or [])]

    return {"project_name": structure.get("project_name"), "nodes": nodes, "edges": edges}


# ══════════════════════════════════════════════════════════════════════
# ЧАСТЬ 6: ВЫВОД В ТЕРМИНАЛ
# ══════════════════════════════════════════════════════════════════════

def print_full_structure(st):
    print(f"\n  Проект: {st['project_name'] or '(не найден)'}")
    if st.get("customer"):
        c=st["customer"]; print(f"  ЗАКАЗЧИК: {c['name']}  БИН={c['bin'] or '—'}")
    def show(o,ind,lbl):
        print(f"{ind}{lbl}: {o['name']}  БИН={o['bin'] or '—'}")
        ct=o.get("contract") or {}
        if ct.get("amount"): print(f"{ind}   сумма: {ct['amount']}")
        for v in o.get("violations",[]): print(f"{ind}   ⚠ {v[:90]}")
        for s in o.get("subcontractors",[]): show(s,ind+"   ","СУБ")
    for i,c in enumerate(st.get("contractors",[]),1): show(c,"  ",f"ИСПОЛНИТЕЛЬ {i}")
    if st.get("violations"):
        for v in st["violations"]: print(f"  ⚠ ГЛОБАЛЬНОЕ: {v[:90]}")


# ══════════════════════════════════════════════════════════════════════
# ЧАСТЬ 7: LLM
# ══════════════════════════════════════════════════════════════════════

def parse_amount_to_tenge(s):
    # FIX 8: убран abs() — отрицательные значения сохраняют знак
    if not s: return None
    s = str(s).lower().replace(",",".").replace("~","").replace("≈","")
    sign = -1 if s.strip().startswith("-") else 1
    m = re.search(r'\d[\d\s]*(?:\.\d+)?', s)
    if not m: return None
    try: num = float(m.group(0).replace(" ",""))
    except ValueError: return None
    if "млрд" in s:   num *= 1_000_000_000
    elif "млн" in s:  num *= 1_000_000
    elif "тыс" in s:  num *= 1_000
    return sign * num

def human_tenge(n):
    if n is None: return None
    return f"{int(round(abs(n))):,}".replace(",", " ")

def _role_name(kind):
    return {"customer":"заказчика","contractor":"исполнителя (поставщика)",
            "subcontractor":"субподрядной организации"}.get(kind,"участника")

def _count_subs(conts):
    n=0
    def w(o):
        nonlocal n
        for s in o.get("subcontractors") or []: n+=1; w(s)
    for c in conts: w(c)
    return n

def _count_viols(st):
    n=len(st.get("violations") or [])
    def w(o):
        nonlocal n
        n+=len(o.get("violations") or [])
        for s in o.get("subcontractors") or []: w(s)
    for c in st.get("contractors") or []: w(c)
    return n

def _all_orgs(st):
    out=[]
    if st.get("customer"): out.append(("customer",st["customer"]))
    def w(o,kind): out.append((kind,o)); [w(s,"subcontractor") for s in o.get("subcontractors") or []]
    for c in st.get("contractors") or []: w(c,"contractor")
    return out

def analyze_cost_changes(st):
    findings=[]
    def w(parent):
        top=parse_amount_to_tenge((parent.get("contract") or {}).get("amount"))
        for sub in parent.get("subcontractors") or []:
            bot=parse_amount_to_tenge((sub.get("contract") or {}).get("amount"))
            if top and bot:
                diff=top-bot  # FIX 8: без abs, знак важен
                pct=diff/top*100 if top else 0
                if abs(pct)>=5:
                    direction="снижение" if diff>0 else "УВЕЛИЧЕНИЕ"
                    findings.append({
                        "parent":parent.get("name","—"),"child":sub.get("name","—"),
                        "parent_amount":top,"child_amount":bot,
                        "diff":abs(diff),"pct":abs(pct),"direction":direction,
                        "suspicious":abs(pct)>=20,
                    })
            w(sub)
    for c in st.get("contractors") or []: w(c)
    return findings

_INVESTIGATOR_STYLE = """Ты — старший аналитик Агентства по финансовому мониторингу Республики Казахстан. По схеме движения средств и связей между организациями ты готовишь аналитическую справку для служебного пользования.

КАК ПИСАТЬ (стиль):
- Пиши связным, читаемым текстом опытного аналитика, а не сухим перечислением. Предложения соединяй логически (поскольку, в связи с этим, при этом, вместе с тем, как следствие), меняй их построение — не начинай каждое одинаково.
- Тон — официально-деловой и сдержанный, но человеческий и понятный. Формулируй так, чтобы читатель сразу видел суть и логику вывода.
- Приводи конкретику: наименования организаций, суммы, цепочку, роли — своими словами, а не копируя формулировки из данных.

ОБЯЗАТЕЛЬНЫЕ ПРАВИЛА:
- ПИШИ ИСКЛЮЧИТЕЛЬНО НА РУССКОМ ЯЗЫКЕ. Ни одного слова на английском или другом языке; нет русского термина — опиши по-русски.
- Используй ТОЛЬКО факты из предоставленных данных. Ничего не додумывай, не добавляй сумм, имён и обстоятельств, которых нет.
- Не повторяй одни и те же факты дословно в разных разделах.
- Не делай юридических выводов о виновности. Запрещены слова «сговор», «мошенничество», «коррупция» — используй «признаки», «обстоятельства, требующие проверки»."""

def llm_available():
    from llm.qwen import is_available
    return is_available()


def _looks_russian(txt):
    # FIX 5: проверяем % кириллицы + стоп-слова
    if not txt: return False
    total_alpha = len(re.findall(rf'[a-zA-Zа-яёА-ЯЁ{KZ}]', txt))
    if total_alpha == 0: return False
    cyr_count = len(re.findall(rf'[а-яёА-ЯЁ{KZ}]', txt))
    cyr_pct   = cyr_count / total_alpha
    if cyr_pct < 0.3: return False  # менее 30% кириллицы — отбраковываем
    EN_STOP = {"the","is","are","was","were","however","company","may","certain","tasks","ensure",
               "have","that","noted","note","engaged","subcontracting","execution","project",
               "increase","volume","provided","chain","work","services"}
    low_lat = re.findall(r'\b[a-z]{3,}\b', txt.lower())
    # FIX5: даже одно стоп-слово в русском тексте — признак английского дрейфа
    en_stops_found = sum(1 for w in low_lat if w in EN_STOP)
    if en_stops_found >= 2: return False
    return True

def _llm(prompt, num_predict=500, temperature=0.25):
    for temp in [temperature, 0.0]:
        try:
            from llm.qwen import generate
            txt = generate(prompt, num_predict=num_predict, temperature=temp, timeout=180)
        except: return None
        txt=re.sub(r'\*\*(.+?)\*\*',r'\1',txt)
        txt=re.sub(r'[#*_`]','',txt)
        txt=_clean(txt)
        if txt and _looks_russian(txt): return txt
    return None

def llm_paragraph(facts_str, task, num_predict=450, temperature=0.25):
    prompt=(f"{_INVESTIGATOR_STYLE}\n\nФАКТЫ:\n{facts_str}\n\n"
            f"ЗАДАЧА: {task}\nОтветь ОДНИМ связным абзацем строго на русском языке.")
    return _llm(prompt, num_predict, temperature)

def llm_investigator_conclusion(st, cost_findings):
    cust=st.get("customer") or {}; conts=st.get("contractors") or []; proj=st.get("project_name") or "—"
    lines=[f"ПРОЕКТ: {proj}",
           f"ЗАКАЗЧИК: {cust.get('name','—')}  БИН: {cust.get('bin','—')}",
           f"Исполнителей: {len(conts)}, субподрядчиков: {_count_subs(conts)}, нарушений: {_count_viols(st)}"]
    for kind,org in _all_orgs(st):
        if kind=="customer": continue
        ct=org.get("contract") or {}; h=org.get("header") or {}
        s=f"[{kind.upper()}] {org.get('name','—')}  БИН: {org.get('bin','—')}"
        if ct.get("amount"): s+=f"  сумма: {ct['amount']}"
        if org.get("activity"): s+=f"  деятельность: {org['activity']}"
        if h.get("fio"): s+=f"  руководитель: {h['fio']}"
        lines.append(s)
        for v in org.get("violations") or []: lines.append(f"  НАРУШЕНИЕ: {v}")
    for f in cost_findings:
        lines.append(f"ОТКЛОНЕНИЕ: {f['parent']} → {f['child']}: "
                     f"{human_tenge(f['parent_amount'])} → {human_tenge(f['child_amount'])} тг "
                     f"({f['direction']} {f['pct']:.1f}%)" + (" [ПОДОЗРИТЕЛЬНО]" if f["suspicious"] else ""))
    for v in st.get("violations") or []: lines.append(f"ГЛОБАЛЬНОЕ НАРУШЕНИЕ: {v}")
    prompt=(f"{_INVESTIGATOR_STYLE}\n\nДАННЫЕ СХЕМЫ:\n" + "\n".join(lines) + "\n\n"
            "ЗАДАЧА: Напиши раздел «ОБЩИЙ ВЫВОД» справки — единым связным текстом (не списком, не по пунктам), 4–7 предложений.\n"
            "Раскрой по смыслу, естественно переходя от одного к другому:\n"
            "— что установлено (кто заказчик, кто исполнители, суммы, цепочка субподряда);\n"
            "— какие обстоятельства вызывают вопросы и почему именно (отклонение сумм по цепочке, нарушения, отсутствие подтверждающих документов);\n"
            "— общий смысл: что требует дополнительной проверки.\n"
            "Пиши по-русски, живо и по делу, без канцелярских штампов-повторов.")
    return _llm(prompt, num_predict=650, temperature=0.3)

def llm_investigator_recs(st, cost_findings):
    cust=st.get("customer") or {}; conts=st.get("contractors") or []; proj=st.get("project_name") or "—"
    lines=[f"ПРОЕКТ: {proj}", f"ЗАКАЗЧИК: {cust.get('name','—')}"]
    for kind,org in _all_orgs(st):
        if kind=="customer": continue
        ct=org.get("contract") or {}
        s=f"[{kind.upper()}] {org.get('name','—')}"
        if ct.get("description"): s+=f"  предмет: {ct['description']}"
        if ct.get("amount"):      s+=f"  сумма: {ct['amount']}"
        lines.append(s)
        for v in org.get("violations") or []: lines.append(f"  НАРУШЕНИЕ: {v}")
    for f in cost_findings:
        lines.append(f"ОТКЛОНЕНИЕ: {f['parent']} → {f['child']}  {f['pct']:.0f}% "
                     + ("(ПОДОЗРИТЕЛЬНО)" if f["suspicious"] else ""))
    for v in st.get("violations") or []: lines.append(f"ГЛОБАЛЬНОЕ: {v}")
    prompt=(f"{_INVESTIGATOR_STYLE}\n\nДАННЫЕ СХЕМЫ:\n" + "\n".join(lines) + "\n\n"
            "ЗАДАЧА: Напиши раздел «РЕКОМЕНДАЦИИ» справки — что конкретно проверить дальше.\n"
            "- Каждый пункт с новой строки, начинается с номера и точки (1. 2. 3. ...).\n"
            "- 5–8 пунктов, каждый — законченное конкретное действие ИМЕННО по этой схеме "
            "(какую организацию, по какому договору/сумме, что запросить или сверить).\n"
            "- Формулируй по-деловому, но конкретно; не повторяй одинаковые пункты разными словами.\n"
            "- Строго на русском.")
    return _llm(prompt, num_predict=750, temperature=0.25)


# ══════════════════════════════════════════════════════════════════════
# ЧАСТЬ 8: ДЕТЕРМИНИРОВАННЫЕ ТЕКСТЫ (fallback)
# ══════════════════════════════════════════════════════════════════════

def _det_general(st):
    cust=st.get("customer") or {}; conts=st.get("contractors") or []; proj=st.get("project_name") or "проекта"
    names=[c["name"] for c in conts if c.get("name")]
    parts=[f"В рамках анализа проекта «{proj}» установлено, что заказчиком выступает "
           f"{cust.get('name') or 'организация-заказчик'}"
           + (f" (БИН {cust['bin']})" if cust.get("bin") else "") + "."]
    if len(names)==1: parts.append(f"Для исполнения обязательств привлечена организация {names[0]}.")
    elif names: parts.append("Для исполнения обязательств привлечены: " + ", ".join(names) + ".")
    total=sum(v for c in conts for v in [parse_amount_to_tenge((c.get("contract") or {}).get("amount"))] if v)
    if total: parts.append(f"Общая стоимость договорных обязательств — {human_tenge(total)} тенге.")
    return " ".join(parts)

def _det_org_paragraph(org, kind):
    name=org.get("name") or "Организация"
    s=name
    if org.get("bin"): s+=f" (БИН {org['bin']})"
    s+=f" выступает в роли {_role_name(kind)}."
    if org.get("activity"): s+=f" Вид деятельности: {org['activity']}."
    ct=org.get("contract") or {}
    if ct.get("description") and ct.get("amount"): s+=f" Предмет: «{ct['description']}», сумма {ct['amount']}."
    elif ct.get("amount"): s+=f" Сумма договора: {ct['amount']}."
    elif ct.get("description"): s+=f" Предмет: «{ct['description']}»."
    return s

def _det_conclusion(st, cost_findings):
    cust=st.get("customer") or {}; conts=st.get("contractors") or []
    n_sub=_count_subs(conts); n_viol=_count_viols(st); proj=st.get("project_name") or "проекта"
    parts=[f"По результатам анализа проекта «{proj}» задействовано {1+len(conts)+n_sub} организаций: "
           f"заказчик, {len(conts)} исполнител{'ь' if len(conts)==1 else 'ей'}"
           + (f" и {n_sub} субподрядчик{'а' if n_sub<5 else 'ов'}" if n_sub else "") + "."]
    if n_viol: parts.append(f"Выявлено {n_viol} обстоятельств, требующих проверки.")
    for f in cost_findings:
        if f["suspicious"]:
            parts.append(f"По цепочке {f['parent']} → {f['child']} стоимость изменилась на {f['pct']:.0f}% — выявлены признаки отклонения.")
    parts.append("Материалы требуют дополнительной проверки.")
    return " ".join(parts)

def _det_recs(st, cost_findings):
    """Рекомендации по требованиям из ТЗ — 8 базовых + контекстные."""
    conts = st.get("contractors") or []
    any_v = _count_viols(st) > 0
    has_s = _count_subs(conts) > 0
    recs = [
        "Провести дополнительную аудиторскую проверку исполнения договорных обязательств.",
        "Запросить подтверждающие документы о фактическом выполнении работ и поставке товаров.",
        "Провести сверку финансовых документов и платёжных операций.",
        "Истребовать сведения о движении денежных средств в рамках исполнения договора.",
        "Провести проверку обоснованности формирования стоимости закупки.",
        "Рассмотреть вопрос о передаче материалов в уполномоченный орган государственного аудита.",
        "Направить материалы для дополнительного изучения в Департамент внутреннего государственного аудита.",
        "Организовать повторный контроль исполнения обязательств поставщиком.",
    ]
    if has_s:
        recs.append("Провести анализ аффилированности участников закупки и проверить обоснованность привлечения субподрядных организаций.")
    if any(f["suspicious"] for f in cost_findings):
        recs.append("Провести независимую оценку рыночной стоимости закупленных товаров, работ и услуг.")
    if any_v:
        recs.append("Рассмотреть вопрос о применении мер ответственности к исполнителям, допустившим нарушения условий договора.")
    return recs


# ══════════════════════════════════════════════════════════════════════
# ЧАСТЬ 9: СБОРКА БЛОКОВ ДОКУМЕНТА
# ══════════════════════════════════════════════════════════════════════

def build_report_blocks(st, use_llm=True):
    blocks=[]; add=lambda k,t="": blocks.append((k,t))
    cust=st.get("customer") or {}; conts=st.get("contractors") or []
    proj=st.get("project_name") or ""
    use_llm=use_llm and llm_available()
    cost=analyze_cost_changes(st)

    # ── Шапка ──────────────────────────────────────────────────────────
    add("title","СПРАВКА")

    # Если есть название проекта — пишем подзаголовок
    if proj:
        add("subtitle2", f"по результатам анализа проекта «{proj}»")
    # Если нет — просто СПРАВКА и сразу к делу (без лишней строки)
    add("blank")

    section=1

    # ── 1. ОБЩИЕ СВЕДЕНИЯ ──────────────────────────────────────────────
    add("h1",f"{section}. ОБЩИЕ СВЕДЕНИЯ"); section+=1
    det_gen=_det_general(st)
    gen=None
    if use_llm:
        gen=llm_paragraph(det_gen,"Перепиши как вводный абзац официальной справки. Сохрани все цифры и названия точно. Не добавляй лишнего.",num_predict=300)
    add("body",gen or det_gen)
    add("blank")

    # ── 2. УЧАСТНИКИ — нумерованный список кратко (заказчик + исполнители) ──
    if conts:
        add("h1",f"{section}. СВЕДЕНИЯ ОБ УЧАСТНИКАХ ПРОЕКТА"); section+=1
        counter=[1]
        def emit_participant(o, is_sub, level=0):
            ct  = o.get("contract") or {}
            amt = ct.get("amount") or o.get("amount") or ""
            amt = re.sub(r'^(Общая\s+сумма|Сумма)\s*:?\s*','',amt,flags=re.I).strip()
            tru = ct.get("description") or o.get("activity") or ""
            # Убираем суммы и хвосты типа "за сумму", "на 2024 г." из ТРУ
            tru = re.sub(r'\s*\d[\d\s,.]*(?:млрд|млн\.?|тыс\.?|тг|тенге)[\w.\s\-]*г?\.?','',tru,flags=re.I)
            tru = re.sub(r'\s*(за\s+сумму|на\s+\d{4}\s*г\.?)\s*$','',tru,flags=re.I)
            tru = tru.strip().rstrip(' .,-')
            # Если ТРУ длиннее 120 символов — обрезаем (избегаем текст актов)
            if len(tru) > 120:
                tru = tru[:120].rsplit(' ', 1)[0] + "..."
            tag    = " (субподрядчик)" if is_sub else ""
            prefix = "  " * level
            line   = f"{prefix}{counter[0]}. {o.get('name','')}{tag}"
            if amt: line += f" — {amt}"
            if tru: line += f". ТРУ: {tru}"
            add("num", line + ".")
            counter[0] += 1
            for sub in o.get("subcontractors") or []:
                emit_participant(sub, True, level+1)
        for con in conts:
            emit_participant(con, False)
        add("blank")

    # ── 3. АНАЛИЗ ОРГАНИЗАЦИЙ ──────────────────────────────────────────
    add("h1",f"{section}. АНАЛИЗ ОРГАНИЗАЦИЙ"); section+=1

    # FIX 6: depth — уровень 2+ получает сокращённый блок без отдельного LLM-запроса
    def emit_org(org, kind, depth=0):
        add("h2", org.get("name") or "Организация")
        add("spravochno","Справочно")
        h=org.get("header") or {}
        ref_parts=[]
        if org.get("name"):         ref_parts.append(org["name"])
        if org.get("bin"):          ref_parts.append(f"БИН {org['bin']}")
        if org.get("legal_address"):ref_parts.append(f"адрес: {org['legal_address']}")
        if org.get("activity"):     ref_parts.append(f"деятельность: {org['activity']}")
        if h.get("fio"):
            s=f"руководитель: {h['fio']}"
            if h.get("iin"):   s+=f", ИИН {h['iin']}"
            if h.get("phone"): s+=f", тел. {h['phone']}"
            ref_parts.append(s)
        for ex in org.get("extra") or []:
            ref_parts.append(ex)
        if ref_parts: add("spravochno_body","; ".join(ref_parts)+".")
        add("blank")

        subs=org.get("subcontractors") or []
        viols=org.get("violations") or []
        ct=org.get("contract") or {}
        ana_facts=[]
        if ct.get("description"): ana_facts.append(f"Предмет договора: {ct['description']}")
        if ct.get("amount"):      ana_facts.append(f"Сумма договора: {ct['amount']}")
        for sub in subs:
            sc=sub.get("contract") or {}
            ana_facts.append(f"Привлечён субподрядчик: {sub.get('name','—')}" + (f" на {sc['amount']}" if sc.get("amount") else ""))
        for v in viols: ana_facts.append(f"Нарушение: {v}")

        body=None
        # FIX 6: LLM только для глубины 0 и 1, уровень 2+ — детерминированный
        if use_llm and ana_facts and depth < 2:
            body=llm_paragraph("\n".join(ana_facts),
                f"Аналитический абзац по {org.get('name')} в роли {_role_name(kind)}. "
                "Не повторяй справочные данные (БИН, адрес, ФИО). Фокус: договор, субподряд, нарушения.",
                num_predict=350)
        if not body: body=_det_org_paragraph(org,kind)
        add("body",body)

        if subs:
            for sub in subs:
                sc=sub.get("contract") or {}
                amt=re.sub(r'^(Общая\s+сумма|Сумма)\s*:?\s*','',sc.get("amount") or "",flags=re.I).strip()
                s=(f"Для исполнения обязательств {org.get('name','исполнитель')} "
                   f"привлекло {sub.get('name','субподрядную организацию')} в качестве субподрядчика.")
                if amt: s+=f" Стоимость: {amt}."
                add("body",s)

        if viols:
            add("body","В ходе анализа выявлены следующие обстоятельства: "
                + "; ".join(_clean(v) for v in viols)+".")

        name_ = org.get("name","организации")
        ct_   = org.get("contract") or {}
        amt_  = ct_.get("amount") or ""
        if viols and subs:
            concl = (f"Вывод. По результатам анализа {name_} установлено наличие обстоятельств, "
                     f"требующих дополнительной проверки, а также факт передачи части обязательств "
                     f"субподрядным организациям.")
        elif viols:
            concl = (f"Вывод. По результатам анализа {name_} выявлены обстоятельства, "
                     f"требующие дополнительной проверки полноты исполнения договорных обязательств.")
        elif subs:
            concl = (f"Вывод. По результатам анализа {name_} установлен факт привлечения "
                     f"субподрядных организаций к исполнению обязательств по договору"
                     + (f" на сумму {amt_}" if amt_ else "") + ".")
        else:
            concl = (f"Вывод. Сведения по {name_} соответствуют материалам анализа. "
                     f"Дополнительных обстоятельств, требующих проверки, не выявлено.")
        add("body",concl)
        add("blank")

    def walk_emit(o, kind, depth=0):
        emit_org(o, kind, depth)
        for sub in o.get("subcontractors") or []:
            walk_emit(sub,"subcontractor", depth+1)

    for c in conts: walk_emit(c,"contractor",0)

    # ── 4. АНАЛИЗ СТОИМОСТИ ─────────────────────────────────────────────
    if cost:
        add("h1",f"{section}. АНАЛИЗ ИЗМЕНЕНИЯ СТОИМОСТИ"); section+=1
        for f in cost:
            direction="снизилась" if f["direction"]=="снижение" else "ВОЗРОСЛА"
            txt=(f"По цепочке {f['parent']} → {f['child']}: стоимость {direction} "
                 f"с {human_tenge(f['parent_amount'])} до {human_tenge(f['child_amount'])} тг "
                 f"({f['pct']:.1f}%, разница {human_tenge(f['diff'])} тг).")
            if f["suspicious"]: txt+=" Выявлены признаки существенного отклонения стоимости."
            add("body",txt)
        add("blank")

    # ── 5. ОБЩИЙ ВЫВОД ───────────────────────────────────────────────────
    add("h1",f"{section}. ОБЩИЙ ВЫВОД"); section+=1
    concl=None
    if use_llm: concl=llm_investigator_conclusion(st,cost)
    add("body",concl or _det_conclusion(st,cost))
    add("blank")

    # ── 6. РЕКОМЕНДАЦИИ ──────────────────────────────────────────────────
    add("h1",f"{section}. РЕКОМЕНДАЦИИ")
    recs_raw=None
    if use_llm: recs_raw=llm_investigator_recs(st,cost)
    recs_parsed=[]
    if recs_raw:
        for m in re.finditer(r'\d+\.\s+(.+?)(?=\s*\d+\.|$)',recs_raw,re.DOTALL):
            item=re.sub(r'\s+'," ",m.group(1)).strip().rstrip(".")
            # FIX 7: ограничение длины пункта 300 символов
            if 20 < len(item) <= 300: recs_parsed.append(item+".")
    if not recs_parsed: recs_parsed=_det_recs(st,cost)
    for i,rec in enumerate(recs_parsed,1): add("num",f"{i}. {rec}")
    add("blank")
    add("signature","Исполнитель: _______________________")
    add("signature","Дата:        ___________________________")
    return blocks


# ══════════════════════════════════════════════════════════════════════
# ЧАСТЬ 10: РЕНДЕР В WORD
# ══════════════════════════════════════════════════════════════════════

def embed_graph_in_docx(docx_path, graph):
    """
    Встраивает графовую JSON-модель (nodes+edges с координатами) в уже
    сохранённый .docx как скрытую часть customXml/item1.xml — невидимую
    пользователю. Это даёт word_to_excel возможность точно восстановить
    схему по координатам, а не заново угадывать раскладку из текста.
    .docx — обычный zip, поэтому части добавляются напрямую в архив.
    """
    import base64, shutil, zipfile

    payload = base64.b64encode(json.dumps(graph, ensure_ascii=False).encode("utf-8")).decode("ascii")

    tmp_path = docx_path + ".tmp"
    with zipfile.ZipFile(docx_path, "r") as zin:
        names = zin.namelist()

        # python-docx уже кладёт свои customXml/item1.xml (itemProps) в шаблон по
        # умолчанию — нельзя слепо использовать "item1.xml", иначе будет дубликат
        # имени в zip и Word покажет «требуется восстановление файла».
        existing_items = [int(m.group(1)) for n in names
                           for m in [re.match(r'customXml/item(\d+)\.xml$', n)] if m]
        item_n = max(existing_items, default=0) + 1
        part_name = f"customXml/item{item_n}.xml"
        item_xml = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><graphData encoding="base64">{payload}</graphData>'

        content_types = zin.read("[Content_Types].xml").decode("utf-8")
        rels_path = "word/_rels/document.xml.rels"
        rels = zin.read(rels_path).decode("utf-8") if rels_path in names else \
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'

        if f"/{part_name}" not in content_types:
            content_types = content_types.replace(
                "</Types>",
                f'<Override PartName="/{part_name}" ContentType="application/xml"/></Types>'
            )

        existing_ids = re.findall(r'Id="(rId\d+)"', rels)
        next_n = max([int(i[3:]) for i in existing_ids], default=0) + 1
        rid = f"rId{next_n}"
        rels = rels.replace(
            "</Relationships>",
            f'<Relationship Id="{rid}" '
            f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" '
            f'Target="../{part_name}"/></Relationships>'
        )

        with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename in ("[Content_Types].xml", rels_path):
                    continue
                zout.writestr(item, zin.read(item.filename))
            zout.writestr("[Content_Types].xml", content_types)
            zout.writestr(rels_path, rels)
            zout.writestr(part_name, item_xml)

    shutil.move(tmp_path, docx_path)


def render_docx(blocks, output_path):
    from docx import Document
    from docx.shared import Pt, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn

    doc=Document()
    sec=doc.sections[0]
    sec.page_width=Cm(21.0); sec.page_height=Cm(29.7)
    sec.left_margin=Cm(3.0); sec.right_margin=Cm(1.5)
    sec.top_margin=Cm(2.0);  sec.bottom_margin=Cm(2.0)

    def style_run(run, size=12, bold=False, italic=False):
        run.bold=bold; run.italic=italic
        run.font.name="Times New Roman"; run.font.size=Pt(size)
        try:
            run._element.rPr.rFonts.set(qn("w:eastAsia"),"Times New Roman")
            run._element.rPr.rFonts.set(qn("w:cs"),"Times New Roman")
        except: pass

    A=WD_ALIGN_PARAGRAPH
    FMT={
        "title":          (16,True, False,A.CENTER, False,0,  6),
        "subtitle":       (14,True, False,A.CENTER, False,0,  6),
        # FIX 9: subtitle_proj = bold (название проекта), subtitle2 = italic (авто-тема)
        "subtitle2":      (12,False,False,A.CENTER, False,0, 18),
        "h1":             (12,True, False,A.LEFT,   False,14, 6),
        "h2":             (12,True, False,A.LEFT,   False,12, 4),
        "spravochno":     (12,True, True, A.LEFT,   False,4,  2),
        "spravochno_body":(12,False,True, A.LEFT,   True, 0,  6),
        "body":           (12,False,False,A.LEFT,   True, 0,  6),
        "num":            (12,False,False,A.LEFT,   False,1,  3),
        "signature":      (12,False,False,A.LEFT,   False,0,  4),
    }

    for kind,text in blocks:
        if kind=="blank": doc.add_paragraph(); continue
        size,bold,italic,align,indent,spb,spa=FMT.get(kind,FMT["body"])
        txt=_clean(text)
        if not txt: doc.add_paragraph(); continue
        p=doc.add_paragraph(); p.alignment=align
        pf=p.paragraph_format
        if indent: pf.first_line_indent=Cm(1.25)
        pf.space_before=Pt(spb); pf.space_after=Pt(spa); pf.line_spacing=Pt(18)
        style_run(p.add_run(txt),size=size,bold=bold,italic=italic)

    doc.save(output_path)

def generate_report(st, output_path, use_llm=True):
    blocks=build_report_blocks(st, use_llm=use_llm)
    render_docx(blocks, output_path)
    try:
        graph = flatten_graph(st)
        embed_graph_in_docx(output_path, graph)
    except Exception as e:
        print(f"  ⚠️  Не удалось встроить графовую модель в docx: {e}")
    return output_path


# ══════════════════════════════════════════════════════════════════════
# ГЛАВНАЯ ФУНКЦИЯ
# ══════════════════════════════════════════════════════════════════════

def sep(title): print(f"\n{'='*58}\n  {title}\n{'='*58}")

def run(xlsx_path, use_llm=True, anonymize=False, graph=False, html=False):
    base=os.path.splitext(os.path.basename(xlsx_path))[0]
    if anonymize:
        base += "_ANON"
    out_dir=os.path.dirname(os.path.abspath(xlsx_path))
    json_out=os.path.join(out_dir,f"{base}_data.json")
    docx_out=os.path.join(out_dir,f"{base}_spravka.docx")

    sep("ШАГ 1: ЧИТАЕМ EXCEL"); print(f"  {xlsx_path}")
    try: shapes,connectors=read_shapes(xlsx_path)
    except Exception as e: print(f"  ❌ {e}"); sys.exit(1)

    sep("ШАГ 2: КЛАССИФИКАЦИЯ")
    shapes, connectors = normalize_orientation(shapes, connectors)
    shapes=classify(shapes, connectors)
    cnt=Counter(s["type"] for s in shapes)
    icons={"customer":"👤","contractor":"🏢","subcontractor":"🔗","contract":"💰",
           "violation":"⚠️ ","header":"👔","title":"📌","unknown":"❓","org_pending":"⏳"}
    for t,n in sorted(cnt.items()): print(f"  {icons.get(t,'·')} {t:18s}: {n}")

    sep("ШАГ 3: ИЗВЛЕКАЕМ ДАННЫЕ")
    try: structure=build_full_structure(shapes,connectors)
    except Exception as e: print(f"  ❌ {e}"); traceback.print_exc(); sys.exit(1)
    if anonymize:
        from anon import anonymize_data
        structure = anonymize_data(structure)
        print("  🕶️  Режим обезличивания: все значения заменены на заглушки")
    print_full_structure(structure)

    sep("ШАГ 4: СОХРАНЯЕМ JSON")
    def _jsc(obj):
        if isinstance(obj,dict): return {k:_jsc(v) for k,v in obj.items() if not k.startswith("_")}
        if isinstance(obj,list): return [_jsc(i) for i in obj]
        return obj
    with open(json_out,"w",encoding="utf-8") as f: json.dump(_jsc(structure),f,ensure_ascii=False,indent=2)
    print(f"  ✅ {json_out}")

    sep("ШАГ 5: СОЗДАЁМ СПРАВКУ")
    if use_llm:
        if llm_available(): print("  ✅ Qwen2.5 — режим следователя активен")
        else: print("  ⚠️  Ollama недоступен — шаблонный текст")
    else: print("  ▶️  Режим --no-llm")
    try:
        generate_report(structure,docx_out,use_llm=use_llm)
        print(f"  ✅ {docx_out}  ({os.path.getsize(docx_out)//1024} КБ)")
        print(f"  ✅ Графовая модель (координаты+связи) встроена в docx скрыто")
    except Exception as e: print(f"  ❌ {e}"); traceback.print_exc(); sys.exit(1)

    if graph:
        from writers.graph_svg import build_svg
        svg_out = os.path.join(out_dir, f"{base}_graph.svg")
        with open(svg_out, "w", encoding="utf-8") as f:
            f.write(build_svg(structure))
        print(f"  ✅ Граф связей (SVG): {svg_out}")

    if html:
        from writers.graph_html import build_html
        # Редактор строим НАПРЯМУЮ из фигур+связей → сохраняются все договоры,
        # все люди на организацию и связи человек↔человек (не теряем в дереве).
        state = build_editor_state(shapes, connectors, structure.get("project_name") or "")
        html_out = os.path.join(out_dir, f"{base}_editor.html")
        with open(html_out, "w", encoding="utf-8") as f:
            f.write(build_html(state))
        print(f"  ✅ Редактор (HTML):   {html_out}  "
              f"(блоков: {len(state['_editor']['blocks'])}, связей: {len(state['_editor']['edges'])})")

    sep("ГОТОВО")
    print(f"  📊 Excel:   {xlsx_path}")
    print(f"  📋 Данные:  {json_out}")
    print(f"  📄 Справка: {docx_out}\n")

if __name__ == "__main__":
    print("\nExcel схема → Официальная Справка (v5 — с встроенной графовой моделью)")
    args=list(sys.argv[1:])
    use_llm="--no-llm" not in args
    if not use_llm: args.remove("--no-llm")
    anon = "--anonymize" in args
    if anon: args.remove("--anonymize")
    graph = "--graph" in args
    if graph: args.remove("--graph")
    html = "--html" in args
    if html: args.remove("--html")
    if not args:
        print("\n  Использование:")
        print("    python excel_to_word_v5.py схема.xlsx")
        print("    python excel_to_word_v5.py схема.xlsx --no-llm")
        print("    python excel_to_word_v5.py схема.xlsx --anonymize")
        print("    python excel_to_word_v5.py схема.xlsx --graph")
        print("    python excel_to_word_v5.py схема.xlsx --html\n")
        sys.exit(1)
    path=args[0]
    if not os.path.exists(path): print(f"\n  ❌ Не найден: {path}"); sys.exit(1)
    missing=[]
    for pkg,imp in [("python-docx","docx"),("requests","requests")]:
        try: __import__(imp)
        except ImportError: missing.append(pkg)
    if missing: print(f"\n  ❌ Установите: pip install {chr(32).join(missing)}"); sys.exit(1)
    run(path,use_llm=use_llm,anonymize=anon,graph=graph,html=html)