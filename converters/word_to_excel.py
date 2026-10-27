"""
Word Справка → Excel схема (v5 — точный формат пример_1.xlsx)
=============================================================
Воспроизводит схему ТОЧНО как в пример_1.xlsx:
  - Заголовок: белый фон, красный текст, Times New Roman 28pt
  - Org блоки: <a:ln/> + style accent2 (серая рамка Excel)
  - Руководитель: отдельный TextBox под блоком
  - Стрелки: straightConnector1, bw=28575, без цвета
  - Договорные метки: rot=20194346/-20194346, Arial 16pt красный
  - Нарушения: FF0000 фон, белый текст, Arial 16pt
  - Вертикальный договор: Times New Roman 12pt, по центру

Запуск:
  python word_to_excel_v5.py справка.docx
  python word_to_excel_v5.py справка.docx --no-llm
"""

import sys, os, re, io, json, zipfile, requests, traceback
import xml.etree.ElementTree as ET
import xml.sax.saxutils as SX

# Конфиг Qwen/Ollama берётся из корневого config.py (единый источник).
# Fallback — если файл запущен напрямую и config.py не виден в sys.path.
try:
    import os as _os, sys as _sys
    _sys.path.insert(0, _os.path.dirname(_os.path.dirname(_os.path.abspath(__file__))))
    from config import OLLAMA_URL, MODEL
except Exception:
    OLLAMA_URL = "http://localhost:11434"
    MODEL      = "qwen2.5"

# ── EMU константы из пример_1.xlsx ────────────────────────────────────
# Измерено из Synthetic reference Demo Company: col18 off=16499 → col24 off=609402, cx=4270381
# COL = (4270381 - (609402-16499)) / (24-18) = 612913
# ROW = (1252737 - (91108-0)) / (39-33) = 193605
COL = 612913
ROW = 193605


# ══════════════════════════════════════════════════════
# ЧАСТЬ 1: ЧТЕНИЕ WORD
# ══════════════════════════════════════════════════════

def read_docx(path):
    from docx import Document
    doc   = Document(path)
    parts = []
    for p in doc.paragraphs:
        if p.text.strip():
            parts.append(p.text.strip())
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if cell.text.strip():
                    parts.append(cell.text.strip())
    return "\n".join(parts)


def extract_graph_from_docx(path):
    """
    Ищет встроенную графовую модель (customXml/itemN.xml с корнем <graphData>),
    которую кладёт excel_to_word_v5.py. Если найдена — это ТОЧНАЯ копия
    исходной схемы (координаты, связи), и Excel можно восстановить без LLM
    и без пересчёта раскладки. Если не найдена — документ либо создан другой
    системой, либо старой версией скрипта; тогда используется обычный путь
    (текст → LLM/regex → tree-layout), см. extract_llm()/extract_fallback().
    """
    import base64
    try:
        with zipfile.ZipFile(path) as zf:
            candidates = [n for n in zf.namelist() if re.match(r'customXml/item\d+\.xml$', n)]
            for name in candidates:
                try:
                    raw = zf.read(name).decode("utf-8")
                    root = ET.fromstring(raw)
                    if root.tag != "graphData":
                        continue
                    payload = base64.b64decode(root.text or "").decode("utf-8")
                    graph = json.loads(payload)
                    if isinstance(graph, dict) and "nodes" in graph and "edges" in graph:
                        return graph
                except Exception:
                    continue
    except Exception:
        pass
    return None


def graph_to_tree(graph):
    """
    Переводит плоскую графовую модель (nodes+edges) обратно во вложенную
    структуру {"project_name","customer","contractors":[...]}, которую
    понимают Layout/OrgNode/render() — но с сохранением ИСХОДНЫХ координат
    в поле "_geom", чтобы render_from_geometry() мог нарисовать схему
    пиксель-в-пиксель как в оригинале, а не пересчитывать layout заново.
    """
    nodes = graph.get("nodes", [])
    if len(nodes) > 1000 or not all(isinstance(n, dict) and "id" in n for n in nodes):
        raise ValueError("Invalid embedded graph nodes")
    nodes_by_id = {n["id"]: n for n in nodes}
    if len(nodes_by_id) != len(nodes):
        raise ValueError("Duplicate embedded graph node IDs")
    children_of = {}
    parent_of = {}
    for e in graph.get("edges", []):
        children_of.setdefault(e["from"], []).append(e["to"])
        parent_of[e["to"]] = e["from"]

    def to_org(node, trail=()):
        if node["id"] in trail or len(trail) > 32:
            raise ValueError("Embedded graph contains a cycle or exceeds 32 levels")
        trail = trail + (node["id"],)
        return {
            "name": node.get("name"), "bin": node.get("bin"),
            "activity": node.get("activity"), "address": node.get("address"),
            "amount": node.get("amount"), "year": node.get("year"),
            "description": node.get("description"),
            "director": node.get("director"),
            "violations": node.get("violations") or [],
            "_geom": node.get("geometry"),
            "subcontractors": [to_org(nodes_by_id[cid], trail) for cid in children_of.get(node["id"], [])
                                if cid in nodes_by_id],
        }

    customer_node = next((n for n in graph.get("nodes", []) if n["type"] == "customer"), None)
    if not customer_node:
        return None

    contractor_ids = children_of.get(customer_node["id"], [])
    data = {
        "project_name": graph.get("project_name"),
        "customer": to_org(customer_node),
        "contractors": [to_org(nodes_by_id[cid]) for cid in contractor_ids if cid in nodes_by_id],
    }
    return data


# ══════════════════════════════════════════════════════
# ЧАСТЬ 2: ИЗВЛЕЧЕНИЕ ЧЕРЕЗ LLM
# ══════════════════════════════════════════════════════

PROMPT = """Ты аналитик. Из текста справки извлеки ВСЕХ участников и ВСЕ данные.
Верни ТОЛЬКО валидный JSON без ``` и пояснений. Пустые поля = null.

ВАЖНО:
- Не теряй ни одной организации, ни одной суммы, ни одного нарушения.
- Сохраняй отдельные договоры и разные связи одной организации. Одна компания
  может встречаться у разных заказчиков или с разными суммами/предметами договора.
  Объединяй только точные повторы одного договора у одного родителя.
- Субподряд может иметь ЛЮБУЮ глубину: субподрядчик может сам иметь своих
  субподрядчиков ("subcontractors" внутри subcontractors — рекурсивно, а не
  только один уровень). Строй цепочку строго по факту "кто кого привлёк"
  из текста, не ограничивайся двумя уровнями.

Формат узла организации (одинаковый на любом уровне вложенности):
{
  "name": "полное название",
  "bin": "12 цифр или null",
  "activity": "вид деятельности или null",
  "address": "адрес или null",
  "amount": "сумма ТОЧНО как в тексте или null (null для заказчика)",
  "year": "4 цифры или null",
  "description": "предмет договора/ТРУ или null",
  "director": {"fio": null, "iin": null, "phone": null, "address": null},
  "violations": ["текст нарушения"],
  "subcontractors": [ /* узлы такого же формата, рекурсивно, может быть пусто */ ]
}

Верни:
{
  "project_name": "название проекта без слова Проект, или null",
  "customer": { ...узел организации выше, без amount/description... },
  "contractors": [ ...узлы организации выше (прямые исполнители заказчика)... ]
}

Текст справки:
{TEXT}"""


def llm_ok():
    from llm.qwen import is_available
    return is_available()


def extract_llm(text):
    from llm.qwen import generate
    from validation import validate_document, unsupported_facts
    if len(text) > 7000:
        print("  LLM context limit exceeded; using full-text deterministic extraction")
        return None
    raw = generate(PROMPT.replace("{TEXT}", text), num_predict=3000,
                   temperature=0.05, timeout=180, json_mode=True)
    fence = chr(96) * 3
    raw = re.sub("^" + fence + r"(?:json)?\s*|\s*" + fence + "$", "", raw.strip())
    try:
        data = validate_document(json.loads(raw))
        if unsupported_facts(data, text):
            print("  Unsupported LLM facts rejected; using deterministic extraction")
            return None
        return data
    except (ValueError, TypeError):
        return None


# ── Универсальное извлечение без LLM (любой формат документа) ──────────
_KZ  = "әғқңөұүһіӘҒҚҢӨҰҮҺІ"
_LET = "A-Za-zА-Яа-яЁё" + _KZ
_UP  = "A-ZА-ЯЁӘҒҚҢӨҰҮҺІ"
_SHORT = r'ТОО|ГУ|АО|НАО|ИП|ООО|ЗАО|КГУ|РГУ|КГП|РГП|ГККП|НКО|ПК|ЖШС|АҚ|ЖК|КХ|ТД'
_LONG_PATS = [
    (r'товариществ\w*\s+с\s+ограниченной\s+ответственност\w*', 'ТОО'),
    (r'жауапкершілігі\s+шектеулі\s+серіктестік\w*', 'ЖШС'),
    (r'государственн\w*\s+учрежден\w*', 'ГУ'),
    (r'мемлекеттік\s+мекеме\w*', 'ММ'),
    (r'акционерн\w*\s+обществ\w*', 'АО'),
    (r'индивидуальн\w*\s+предпринимател\w*', 'ИП'),
    (r'коммунальн\w*\s+государственн\w*\s+предприят\w*', 'КГП'),
    (r'аппарат\w*\s+акима\w*', 'Аппарат акима'),
    (r'акимат\w*', 'Акимат'), (r'әкімдіг\w*|әкімдік\w*|әкімшілік\w*', 'Әкімдік'),
    (r'управлени\w+', 'Управление'), (r'басқарма\w*', 'Басқарма'),
    (r'министерств\w+', 'Министерство'), (r'министрлі\w+', 'Министрлік'),
    (r'департамент\w*', 'Департамент'), (r'комитет\w*', 'Комитет'),
]
_GOV_RE = re.compile(r'акимат|аппарат\s+акима|управлени|министерств|департамент|комитет|'
                     r'әкім\w*|басқарма|министрлі|тапсырыс\s+беруш|'
                     r'\bГУ\b|\bКГУ\b|\bРГУ\b|\bММ\b|заказчик', re.I)
_STOP = re.compile(r'^(?:БИН|БСН|СУММ\w*|СОМА\w*|НА|ЗА|В|С|И|ПО|КАЧЕСТВ\w*|ПРИВЛЕК\w*|ПЕРЕДА\w*|'
                   r'РУКОВОДИТ\w*|ДИРЕКТОР\w*|БАСШЫ\w*|ДОГОВОР\w*|СТОИМОСТ\w*|ИИН|ЖСН|ФИО|ТЕЛ\w*|'
                   r'АДРЕС|МЕКЕН\w*|НАРУШЕН\w*|СОГЛАСНО|СУБПОДРЯД\w*|ТЕҢГЕ|МЫҢ)$', re.I)
_AMT = re.compile(r'(\d[\d .,]{0,15}?)\s*(млрд|млн|тыс|мың|миллиард\w*|миллион\w*)\s*(?:тг|тенге|теңге)?', re.I)
_BIN = re.compile(r'(?:БИН|БСН)\s*:?\s*№?\s*(\d{12})', re.I)
_IIN = re.compile(r'(?:ИИН|ЖСН)\s*:?\s*(\d{12})', re.I)
_TEL = re.compile(r'(?:тел\.?\s*номер|телефон|тел\.?)\s*:?\s*([+\d][\d\s\-\(\)]{6,15})', re.I)
_ADR = re.compile(r'(?:адрес|мекен-?жай\w*)[^:]*:?\s*([^\n,;]{5,80})', re.I)
_FIO1 = re.compile(rf'ФИО\s*:?\s*([{_UP}][{_LET}\.\-]+(?:\s+[{_UP}][{_LET}\.\-]+){{1,3}})')
_FIO2 = re.compile(rf'(?:руководител[ья]|директор[аы]?|басшы\w*|бас\s+директор\w*)\s*[—:\-]?\s*'
                   rf'([{_UP}][{_LET}\.\-]+(?:\s+[{_UP}][{_LET}\.\-]+){{1,3}})', re.I)
_SUB = re.compile(r'субподряд|привлек\w*|переда\w*\s+част|қосалқы\s+мердігер|тартт\w*', re.I)

def _fb_norm(s):
    return re.sub(r'\s+', ' ', re.sub(r'["«»]', '', s or '')).upper().strip()

def _fb_grab_name(text, end, typ):
    rest = text[end:end+90]
    q = re.match(r'\s*["«]([^"»]{2,70})["»]', rest)
    if q:
        return f'{typ} «{q.group(1).strip()}»'
    seg = re.split(r'[\n(,;:.]', rest, maxsplit=1)[0]      # имя — до разделителя/переноса
    keep = []
    for tk in seg.split():
        if _STOP.match(tk) or re.search(r'\d', tk):
            break
        keep.append(tk)
        if len(keep) >= 4: break
    nm = " ".join(keep).strip(" -")
    return f'{typ} {nm}'.strip() if (nm or typ != 'Акимат') else None

def _fb_sentences(text):
    return [(m.start(), m.group()) for m in re.finditer(r'[^.;\n]+[.;\n]?', text)]

def extract_fallback(text):
    """
    Универсальное извлечение без LLM. Не привязано к одному шаблону:
    организации ловятся по типу (ТОО/ГУ/АО/… или «товарищество с …»,
    «государственное учреждение», «акимат», «управление» и т.п.) — с кавычками
    и без; БИН — по метке и без метки (таблицы); суммы, директора (в т.ч. без
    метки «ФИО» и с казахскими буквами), нарушения — из сегмента организации;
    роли: заказчик = гос-орган, субподряд = по фразе «привлёк/субподряд».
    """
    data = {"project_name": None, "customer": None, "contractors": []}
    m = re.search(r'проект[а-я]*\s*[«"]([^»"]+)[»"]', text, re.I)
    if m: data["project_name"] = m.group(1).strip()

    ment = []
    for mm in re.finditer(rf'\b({_SHORT})\b', text):
        nm = _fb_grab_name(text, mm.end(), mm.group(1))
        if nm: ment.append([mm.start(), nm])
    for pat, canon in _LONG_PATS:
        for mm in re.finditer(pat, text, re.I):
            nm = _fb_grab_name(text, mm.end(), canon)
            if nm: ment.append([mm.start(), nm])
    ment.sort()
    dedup = []                       # убираем пересекающиеся упоминания одной org
    for pos, name in ment:
        drop = False
        for j, (p2, n2) in enumerate(dedup):
            if abs(pos - p2) < 50:
                a, b = _fb_norm(name), _fb_norm(n2)
                if a and b and (a in b or b in a):
                    if len(a) > len(b): dedup[j] = [pos, name]
                    drop = True; break
        if not drop: dedup.append([pos, name])
    ment = dedup
    if not ment:
        return data

    positions = [p for p, _ in ment]
    seen, orgs = {}, []
    for idx, (pos, name) in enumerate(ment):
        key = _fb_norm(name)
        nxt = positions[idx+1] if idx+1 < len(positions) else len(text)
        seg = text[pos:nxt]
        if key in seen:
            seen[key]["seg"] += " " + seg; continue
        o = {"name": name, "pos": pos, "seg": seg, "key": key}
        seen[key] = o; orgs.append(o)

    def fld(seg, rgx):
        r = rgx.search(seg); return r.group(1).strip() if r else None
    def director(seg):
        m1 = _FIO1.search(seg) or _FIO2.search(seg)
        fio = m1.group(1).strip() if m1 else None
        iin, tel, adr = fld(seg, _IIN), fld(seg, _TEL), fld(seg, _ADR)
        return {"fio": fio, "iin": iin, "phone": tel, "address": adr} if any([fio, iin, tel, adr]) else None
    def viols(seg):
        return [s.strip()[:200] for s in re.split(r'(?<=[.;])\s+', seg)
                if re.search(r'нарушени|не\s+соответ|отсутству|согласно\s+акт', s, re.I)]

    for o in orgs:
        o["bin"] = fld(o["seg"], _BIN)
        if not o["bin"]:
            for mm in re.finditer(r'\b\d{12}\b', o["seg"]):
                pre = o["seg"][max(0, mm.start()-6):mm.start()].upper()
                if 'ИИН' not in pre and 'ЖСН' not in pre:
                    o["bin"] = mm.group(0); break
        a = _AMT.search(o["seg"]); o["amount"] = a.group(0).strip() if a else None
        o["director"] = director(o["seg"])
        o["violations"] = viols(o["seg"])
        o["is_gov"] = bool(_GOV_RE.search(o["name"]) or _GOV_RE.search(o["seg"][:50]))

    gov = [o for o in orgs if o["is_gov"]]
    cust = gov[0] if gov else orgs[0]

    sents = _fb_sentences(text)
    def sentence_at(pos):
        cur = None
        for st, s in sents:
            if st <= pos < st + len(s): return st, s
            if st > pos: break
            cur = (st, s)
        return cur if cur else (0, text)

    parent_of = {}
    for o in orgs:
        if o is cust: continue
        st, s = sentence_at(o["pos"])
        if _SUB.search(s):
            prev = [x for x in orgs if x is not cust and st <= x["pos"] < o["pos"]]
            if not prev:
                prev = [x for x in orgs if x is not cust and x["pos"] < o["pos"]]
            if prev: parent_of[o["key"]] = prev[-1]["key"]

    def node(o):
        return {"name": o["name"], "bin": o["bin"], "activity": None, "address": None,
                "amount": o["amount"], "year": None, "description": None,
                "director": o["director"], "violations": o["violations"], "subcontractors": []}
    nodes = {o["key"]: node(o) for o in orgs if o is not cust}
    roots = []
    for o in orgs:
        if o is cust: continue
        p = parent_of.get(o["key"])
        if p and p in nodes and p != o["key"]:
            nodes[p]["subcontractors"].append(nodes[o["key"]])
        else:
            roots.append(nodes[o["key"]])

    data["customer"] = {"name": cust["name"], "bin": cust["bin"], "activity": None,
                        "address": None, "director": cust["director"]}
    data["contractors"] = roots
    return data


# ══════════════════════════════════════════════════════
# ЧАСТЬ 2b: ВАЛИДАЦИЯ И ДЕДУПЛИКАЦИЯ ГРАФА
# ══════════════════════════════════════════════════════
#
# LLM извлекает смысл, но не обязана сама следить за целостностью графа —
# поэтому дубли/пустые узлы отлавливаются здесь детерминированным кодом,
# а не повторным обращением к модели.

def _norm_name(name):
    """Нормализация названия организации для сравнения (без кавычек/регистра/пробелов)."""
    if not name:
        return ""
    s = str(name).upper()
    s = re.sub(r'["\u00ab\u00bb\u2018\u2019\u201c\u201d]', '', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s


def _node_key(node):
    """Ключ идентичности организации: БИН приоритетнее названия."""
    b = (node.get("bin") or "").strip()
    if b:
        return f"bin:{b}"
    return f"name:{_norm_name(node.get('name'))}"


def validate_and_dedup(data, warnings=None):
    """Merge exact sibling occurrences; preserve distinct contracts and parents."""
    from validation import validate_document
    validate_document(data)
    warnings = [] if warnings is None else warnings

    def clean(nodes):
        seen = {}
        kept = []
        for node in nodes:
            if not (node.get("name") or "").strip():
                warnings.append("Пропущен пустой узел без названия")
                continue
            node["subcontractors"] = clean(node.get("subcontractors") or [])
            details = {k: v for k, v in node.items()
                       if k not in {"name", "bin", "subcontractors", "violations"}}
            key = (_node_key(node), json.dumps(details, ensure_ascii=False, sort_keys=True))
            if key not in seen:
                seen[key] = node
                kept.append(node)
                continue
            target = seen[key]
            target["violations"] = list(dict.fromkeys(
                (target.get("violations") or []) + (node.get("violations") or [])))
            target["subcontractors"] = clean(target["subcontractors"] + node["subcontractors"])
            warnings.append("Объединён точный повтор одной организации и договора у одного родителя")
        return kept

    data["contractors"] = clean(data["contractors"])
    if data.get("customer"):
        data["customer"]["subcontractors"] = clean(data["customer"].get("subcontractors") or [])
    return data, warnings


# ══════════════════════════════════════════════════════
# ЧАСТЬ 3: LAYOUT ENGINE
# ══════════════════════════════════════════════════════

def node_size(org):
    """
    Автоматический размер блока организации по объёму текста —
    вместо фиксированного ORG_COLS/ORG_ROWS для всех узлов подряд.
    Эвристика откалибрована на средние пропорции блоков пример_1.xlsx
    (там 1 колонка ≈ 9.2 символа шрифтом 12pt).
    """
    name = (org.get("name") or "").strip()
    cols = 7
    if len(name) > 18:
        cols = min(18, 7 + (len(name) - 18) // 6)

    lines = 1 if name else 0                      # строка названия
    if org.get("bin"):      lines += 1
    if org.get("activity"): lines += 1 + len(str(org["activity"])) // (cols * 9)
    if org.get("address"):  lines += 1 + len(str(org["address"]))  // (cols * 9)
    rows = max(8, 4 + lines * 2)
    return cols, rows


class OrgNode:
    """
    Узел дерева граф-модели схемы: заказчик → исполнители → (рекурсивно)
    субподрядчики любой глубины. Каждый узел сам знает свой размер и, после
    compute_width()/assign_cols()/assign_rows(), свою позицию — раскладка
    полностью детерминирована и без пересечений по построению (ширина
    поддерева = сумма ширин детей, а не константа), в отличие от прежнего
    Layout, который понимал только один уровень subcontractors.
    """

    GAP_C = 2
    GAP_R = 3
    HDR_COLS = 6
    HDR_ROWS = 7
    VIO_COLS = 10
    VIO_ROWS = 5
    CON_COLS = 8
    CON_ROWS = 9

    def __init__(self, data, kind="subcontractor"):
        self.data = data
        self.kind = kind
        self.cols, self.rows = node_size(data)
        # Ширина, которую ветка реально резервирует у соседей, должна учитывать
        # не только сам блок организации, но и блок руководителя (рисуется правее
        # на HDR_COLS) и блоки нарушений (рисуются сбоку на VIO_COLS) — иначе
        # соседняя ветка может встать поверх них (визуальный баг, который и
        # проявлялся как «фигуры стоят неправильно»).
        footprint = max(self.cols, 1 + self.HDR_COLS)
        n_viol = len(data.get("violations") or [])
        if n_viol:
            footprint += 2 * (self.VIO_COLS + 2)  # запас с обеих сторон (сторона определится позже)
        self.footprint_cols = footprint
        self.children = [OrgNode(c, "subcontractor")
                          for c in (data.get("subcontractors") or [])]
        self.col = 0
        self.row = 0
        self.subtree_width = self.footprint_cols
        self.side = "left"   # унаследуется от корневого исполнителя (для нарушений/меток)

    def compute_width(self):
        if not self.children:
            self.subtree_width = self.footprint_cols
        else:
            used = sum(c.compute_width() for c in self.children) + self.GAP_C * (len(self.children) - 1)
            self.subtree_width = max(self.footprint_cols, used)
        return self.subtree_width

    def assign_cols(self, left_edge):
        if not self.children:
            self.col = left_edge + (self.subtree_width - self.cols) // 2
            return
        used = sum(c.subtree_width for c in self.children) + self.GAP_C * (len(self.children) - 1)
        cursor = left_edge + max(0, (self.subtree_width - used) // 2)
        centers = []
        for c in self.children:
            c.side = self.side
            c.assign_cols(cursor)
            centers.append(c.col + c.cols / 2)
            cursor += c.subtree_width + self.GAP_C
        self.col = int(round(sum(centers) / len(centers) - self.cols / 2))

    def assign_rows(self, top_row):
        self.row = top_row
        next_row = top_row + self.rows + self.HDR_ROWS + self.GAP_R + 2
        for c in self.children:
            c.assign_rows(next_row)

    def walk(self):
        yield self
        for c in self.children:
            yield from c.walk()


class Layout:
    """
    Собирает общую раскладку: заказчик в центре сверху, исполнители —
    рекурентные поддеревья (OrgNode) слева/справа от него, каждое со своей
    произвольной глубиной субподряда. Пересечения невозможны по построению,
    т.к. соседние ветки резервируют себе ровно subtree_width, а не
    фиксированную константу.
    """

    GAP_C = 2
    GAP_R = 3

    def __init__(self, data):
        self.data  = data
        self.cust  = data.get("customer") or {}
        self.cons  = data.get("contractors") or []
        self._exact = self._geometry_available()
        if self._exact:
            g = self.cust["_geom"]
            self.cust_cols = g["to_col"] - g["from_col"]
            self.cust_rows = g["to_row"] - g["from_row"]
        else:
            self.cust_cols, self.cust_rows = node_size(self.cust) if self.cust else (7, 8)
        self.ORG_COLS = self.cust_cols   # для обратной совместимости старого render()
        self.ORG_ROWS = self.cust_rows
        self.HDR_COLS = OrgNode.HDR_COLS
        self.HDR_ROWS = OrgNode.HDR_ROWS
        self.CON_COLS = OrgNode.CON_COLS
        self.CON_ROWS = OrgNode.CON_ROWS
        self.VIO_COLS = OrgNode.VIO_COLS
        self.VIO_ROWS = OrgNode.VIO_ROWS
        self.GAP_R = OrgNode.GAP_R
        self._calc()

    def _geometry_available(self):
        """
        True, если у заказчика и у ВСЕХ исполнителей/субподрядчиков (любой
        глубины) сохранена исходная геометрия ("_geom" — встроена
        excel_to_word_v5.py). В этом случае Excel строится ПОЛНОСТЬЮ ТОЧНО
        по оригинальным координатам, без пересчёта раскладки и без LLM.
        """
        def has_geom(o):
            if not o.get("_geom"):
                return False
            return all(has_geom(k) for k in (o.get("subcontractors") or []))
        if not self.cust.get("_geom"):
            return False
        return all(has_geom(c) for c in self.cons)

    def _calc(self):
        self.trees = [OrgNode(c, "contractor") for c in self.cons]
        if self._exact:
            self._calc_exact()
        else:
            self._calc_layout()
        self._calc_extents()

    def _calc_exact(self):
        """Точное восстановление позиций из встроенной геометрии — без layout-алгоритма."""
        g = self.cust["_geom"]
        self.cust_col = g["from_col"]
        self.cust_row = g["from_row"]
        cust_center = self.cust_col + self.cust_cols / 2

        def assign(node, side):
            ng = node.data.get("_geom") or {}
            node.col  = ng.get("from_col", node.col)
            node.row  = ng.get("from_row", node.row)
            node.cols = ng.get("to_col", node.col + node.cols) - node.col
            node.rows = ng.get("to_row", node.row + node.rows) - node.row
            node.side = side
            for c in node.children:
                assign(c, side)

        for t in self.trees:
            tg = t.data["_geom"]
            center = (tg["from_col"] + tg["to_col"]) / 2
            side = "left" if center < cust_center else "right"
            assign(t, side)

        self.cont_cols = [t.col for t in self.trees]
        self.n_left    = sum(1 for t in self.trees if t.side == "left")
        self.cont_row  = min([t.row for t in self.trees], default=self.cust_row + self.cust_rows)

    def _calc_layout(self):
        """Резервный путь: считает раскладку с нуля (нет исходной геометрии — LLM/regex-извлечение)."""
        GC = self.GAP_C
        for t in self.trees:
            t.compute_width()

        n       = len(self.trees)
        n_left  = (n + 1) // 2
        n_right = n // 2
        left_trees  = self.trees[:n_left]
        right_trees = self.trees[n_left:]

        left_w = sum(t.subtree_width for t in left_trees) + GC * max(0, len(left_trees) - 1)
        self.cust_col = left_w + GC
        self.cust_row = 18   # старт совпадает с пример_1, дальше всё считается от размеров

        self.cont_row = self.cust_row + self.cust_rows + self.GAP_R + 3

        # Левые ветки — размещаем справа налево от заказчика (ближайшая слева)
        cursor = self.cust_col - GC
        for t in reversed(left_trees):
            t.side = "left"
            t.assign_cols(cursor - t.subtree_width)
            t.assign_rows(self.cont_row)
            cursor -= (t.subtree_width + GC)

        # Правые ветки
        cursor = self.cust_col + self.cust_cols + GC
        for t in right_trees:
            t.side = "right"
            t.assign_cols(cursor)
            t.assign_rows(self.cont_row)
            cursor += t.subtree_width + GC

        self.cont_cols = [t.col for t in self.trees]
        self.n_left    = n_left

    def _calc_extents(self):
        max_row = self.cont_row
        max_col = self.cust_col + self.cust_cols
        min_col = self.cust_col
        for t in self.trees:
            for node in t.walk():
                max_row = max(max_row, node.row + node.rows + self.HDR_ROWS + 4)
                max_col = max(max_col, node.col + node.cols)
                min_col = min(min_col, node.col)

        self.total_cols = max_col + 10 - min(0, min_col)
        self.total_rows = max_row + 15


# ══════════════════════════════════════════════════════
# ЧАСТЬ 4: XML СТРОИТЕЛИ (точный формат пример_1)
# ══════════════════════════════════════════════════════

_ID = [1]

def nid():
    i = _ID[0]; _ID[0] += 1; return i


def xe(s):
    """Escape для XML атрибутов."""
    if not s: return ""
    s = str(s)
    # Убираем invalid XML chars
    s = re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]', '', s)
    return SX.escape(s).replace('"', '&quot;').replace("'", '&apos;')


def xt(s):
    """Escape для XML текста."""
    if not s: return ""
    s = str(s)
    s = re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]', '', s)
    return SX.escape(s)


def _emu(col, row, col_off=0, row_off=0):
    """Позиция в EMU."""
    return col * COL + col_off, row * ROW + row_off


def _anchor(fc, fr, fc_off, fr_off, tc, tr, tc_off, tr_off, inner):
    """twoCellAnchor (без атрибута editAs — как в оригинале)."""
    return (
        f'<xdr:twoCellAnchor>'
        f'<xdr:from>'
        f'<xdr:col>{fc}</xdr:col><xdr:colOff>{fc_off}</xdr:colOff>'
        f'<xdr:row>{fr}</xdr:row><xdr:rowOff>{fr_off}</xdr:rowOff>'
        f'</xdr:from>'
        f'<xdr:to>'
        f'<xdr:col>{tc}</xdr:col><xdr:colOff>{tc_off}</xdr:colOff>'
        f'<xdr:row>{tr}</xdr:row><xdr:rowOff>{tr_off}</xdr:rowOff>'
        f'</xdr:to>'
        f'{inner}'
        f'<xdr:clientData/>'
        f'</xdr:twoCellAnchor>'
    )


def _style_accent2():
    """xdr:style как у org-блоков (серая рамка Excel themes)."""
    return (
        '<xdr:style>'
        '<a:lnRef idx="2"><a:schemeClr val="accent2"/></a:lnRef>'
        '<a:fillRef idx="1"><a:schemeClr val="lt1"/></a:fillRef>'
        '<a:effectRef idx="0"><a:schemeClr val="accent2"/></a:effectRef>'
        '<a:fontRef idx="minor"><a:schemeClr val="dk1"/></a:fontRef>'
        '</xdr:style>'
    )


def _para(text, bold=True, sz=1200, color=None, align="ctr",
          font="Times New Roman"):
    """Параграф текста."""
    if not text:
        return (f'<a:p><a:pPr algn="{align}"/>'
                f'<a:endParaRPr lang="ru-RU" sz="{sz}" b="{1 if bold else 0}" dirty="0">'
                f'<a:latin typeface="{font}"/>'
                f'<a:cs typeface="{font}"/>'
                f'</a:endParaRPr></a:p>')
    col_xml = f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>' if color else ''
    return (
        f'<a:p><a:pPr algn="{align}"/><a:r>'
        f'<a:rPr lang="ru-RU" sz="{sz}" b="{1 if bold else 0}" dirty="0">'
        f'{col_xml}'
        f'<a:latin typeface="{font}"/>'
        f'<a:cs typeface="{font}"/>'
        f'</a:rPr>'
        f'<a:t>{xt(text)}</a:t>'
        f'</a:r></a:p>'
    )


# ── Блок организации (TextBox с <a:ln/> + style accent2) ───────────────

def org_box(name, fc, fr, tc, tr, lines,
            fc_off=16499, fr_off=0,
            tc_off=609402, tr_off=91108):
    """
    Точный формат из пример_1: <a:ln/> пустой + xdr:style accent2.
    Именно эта комбинация даёт серую рамку Excel без явного цвета.
    """
    sid  = nid()
    ox, oy = _emu(fc, fr, fc_off, fr_off)
    tx, ty = _emu(tc, tr, tc_off, tr_off)
    cx   = abs(tx - ox)
    cy   = abs(ty - oy)
    ps   = "".join(lines)
    nm   = xe(name)

    sp = (
        f'<xdr:sp macro="" textlink="">'
        f'<xdr:nvSpPr>'
        f'<xdr:cNvPr id="{sid}" name="{nm}"/>'
        f'<xdr:cNvSpPr txBox="1"/>'
        f'</xdr:nvSpPr>'
        f'<xdr:spPr>'
        f'<a:xfrm><a:off x="{ox}" y="{oy}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'<a:ln/>'
        f'</xdr:spPr>'
        f'{_style_accent2()}'
        f'<xdr:txBody>'
        f'<a:bodyPr vertOverflow="clip" horzOverflow="clip" wrap="square" rtlCol="0" anchor="ctr"/>'
        f'<a:lstStyle/>'
        f'{ps}'
        f'</xdr:txBody>'
        f'</xdr:sp>'
    )
    return _anchor(fc, fr, fc_off, fr_off, tc, tr, tc_off, tr_off, sp)


# ── Прямоугольник с заливкой (заголовок, нарушения, метки) ─────────────

def rect_fill(name, fc, fr, tc, tr, lines, bg_color,
              rot=0,
              fc_off=76404, fr_off=60307,
              tc_off=564776, tr_off=59070):
    """
    Прямоугольник с заливкой — точный формат заголовка/нарушений.
    Без wrap в bodyPr (как в оригинале).
    """
    sid  = nid()
    ox, oy = _emu(fc, fr, fc_off, fr_off)
    tx, ty = _emu(tc, tr, tc_off, tr_off)
    cx   = abs(tx - ox)
    cy   = abs(ty - oy)
    ra   = f'rot="{rot}"' if rot else ""
    ps   = "".join(lines)
    nm   = xe(name)

    sp = (
        f'<xdr:sp macro="" textlink="">'
        f'<xdr:nvSpPr>'
        f'<xdr:cNvPr id="{sid}" name="{nm}"/>'
        f'<xdr:cNvSpPr/>'
        f'</xdr:nvSpPr>'
        f'<xdr:spPr>'
        f'<a:xfrm {ra}>'
        f'<a:off x="{ox}" y="{oy}"/><a:ext cx="{cx}" cy="{cy}"/>'
        f'</a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'<a:solidFill><a:srgbClr val="{bg_color}"/></a:solidFill>'
        f'</xdr:spPr>'
        f'{_style_accent2()}'
        f'<xdr:txBody>'
        f'<a:bodyPr vertOverflow="clip" horzOverflow="clip" rtlCol="0" anchor="ctr"/>'
        f'<a:lstStyle/>'
        f'{ps}'
        f'</xdr:txBody>'
        f'</xdr:sp>'
    )
    return _anchor(fc, fr, fc_off, fr_off, tc, tr, tc_off, tr_off, sp)


# ── Фото руководителя: плейсхолдер + вставка реального изображения ──────

def photo_box(name, fc, fr, tc, tr):
    """
    Пустой плейсхолдер под фото руководителя (серая рамка + подпись).
    Человек вставляет сюда картинку вручную; рамка задаёт размер и место.
    """
    return rect_fill(
        name, fc, fr, tc, tr,
        [_para("ФОТО",         bold=True,  sz=1100, color="A6A6A6", align="ctr"),
         _para("руководителя", bold=False, sz=900,  color="A6A6A6", align="ctr")],
        bg_color="F2F2F2",
        fc_off=20000, fr_off=20000, tc_off=560000, tr_off=170000,
    )


def pic_anchor(rid, fc, fr, tc, tr,
               fc_off=20000, fr_off=20000, tc_off=560000, tr_off=170000):
    """
    Реальное фото — <xdr:pic> (формат как в пример_1.xlsx, без поворота).
    r:embed ссылается на media через drawing1.xml.rels (см. pack_xlsx).
    """
    sid = nid()
    ox, oy = _emu(fc, fr, fc_off, fr_off)
    tx, ty = _emu(tc, tr, tc_off, tr_off)
    cx, cy = abs(tx - ox), abs(ty - oy)
    pic = (
        f'<xdr:pic>'
        f'<xdr:nvPicPr>'
        f'<xdr:cNvPr id="{sid}" name="Фото {sid}"/>'
        f'<xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr>'
        f'</xdr:nvPicPr>'
        f'<xdr:blipFill>'
        f'<a:blip r:embed="{rid}"/>'
        f'<a:stretch><a:fillRect/></a:stretch>'
        f'</xdr:blipFill>'
        f'<xdr:spPr>'
        f'<a:xfrm><a:off x="{ox}" y="{oy}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'<a:ln w="28575"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:ln>'
        f'</xdr:spPr>'
        f'</xdr:pic>'
    )
    return _anchor(fc, fr, fc_off, fr_off, tc, tr, tc_off, tr_off, pic)


def load_photos(base_dir):
    """
    Сканирует папку photos/ рядом с исходником. Ключ — нормализованное имя
    файла (без расширения): под него кладут фото по ФИО / ИИН / названию ТОО.
    Возвращает {ключ: {'ext':..,'data':bytes}}. Нет папки — пустой словарь.
    """
    out = {}
    pdir = os.path.join(base_dir, "photos")
    if not os.path.isdir(pdir):
        return out
    for fn in os.listdir(pdir):
        ext = os.path.splitext(fn)[1].lower().lstrip(".")
        if ext not in ("png", "jpg", "jpeg"):
            continue
        stem = os.path.splitext(fn)[0]
        try:
            with open(os.path.join(pdir, fn), "rb") as f:
                out[_norm_name(stem)] = {"ext": ext, "data": f.read()}
        except Exception:
            pass
    return out


def match_photo(photos, node_data):
    """Ищет фото для узла по ФИО руководителя, ИИН или названию организации."""
    if not photos:
        return None
    d    = node_data or {}
    dir_ = d.get("director") or {}
    keys = [dir_.get("fio"), dir_.get("iin"), d.get("name")]
    norm = [_norm_name(k) for k in keys if k]
    for k in norm:
        if k in photos:
            return photos[k]
    # частичное совпадение (имя файла — часть ключа или наоборот)
    for k in norm:
        for pk, val in photos.items():
            if k and pk and (k in pk or pk in k):
                return val
    return None


# ── Стрелка (коннектор) ─────────────────────────────────────────────────

def arrow(name, fc, fr, tc, tr,
          fc_off=10114, fr_off=19406,
          tc_off=22412, tr_off=11207):
    """
    Стрелка из пример_1: straightConnector1, bw=28575, без цвета, tailEnd=arrow.
    Из CXN#7: bw=28575, head=none, tail=none (стрелка через style темы).
    """
    sid  = nid()
    ox, oy = _emu(fc, fr, fc_off, fr_off)
    tx, ty = _emu(tc, tr, tc_off, tr_off)
    cx   = abs(tx - ox) or 914
    cy   = abs(ty - oy) or 914
    nm   = xe(name)

    # Правильный xfrm для стрелки:
    # off = левый верхний угол bounding box
    # ext = ширина и высота bounding box (всегда положительные)
    left_x = min(ox, tx)
    top_y  = min(oy, ty)
    width  = abs(tx - ox) or 914
    height = abs(ty - oy) or 914

    cxn = (
        f'<xdr:cxnSp macro="">'
        f'<xdr:nvCxnSpPr>'
        f'<xdr:cNvPr id="{sid}" name="{nm}"/>'
        f'<xdr:cNvCxnSpPr/>'
        f'</xdr:nvCxnSpPr>'
        f'<xdr:spPr>'
        f'<a:xfrm>'
        f'<a:off x="{left_x}" y="{top_y}"/>'
        f'<a:ext cx="{width}" cy="{height}"/>'
        f'</a:xfrm>'
        f'<a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>'
        # ЯВНАЯ чёрная заливка линии — без неё Excel не рисует коннектор
        # (в пример_1 цвет шёл из <xdr:style>, ниже он тоже добавлен).
        f'<a:ln w="28575">'
        f'<a:solidFill><a:srgbClr val="000000"/></a:solidFill>'
        f'<a:tailEnd type="arrow" w="med" len="med"/>'
        f'</a:ln>'
        f'</xdr:spPr>'
        # Стиль коннектора как в пример_1.xlsx (даёт линии/стрелке цвет темы).
        f'<xdr:style>'
        f'<a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef>'
        f'<a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef>'
        f'<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>'
        f'<a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef>'
        f'</xdr:style>'
        f'</xdr:cxnSp>'
    )
    return _anchor(fc, fr, fc_off, fr_off, tc, tr, tc_off, tr_off, cxn)


# ── Параграфы для org блока ─────────────────────────────────────────────

def org_lines(org, is_customer=False):
    """Строки внутри блока организации."""
    ps = []
    name = (org.get("name") or "").strip()
    if name:
        sz = 1400 if is_customer else 1200
        ps.append(_para(name, bold=True, sz=sz, align="ctr"))
    if org.get("bin"):
        ps.append(_para(f'БИН {org["bin"]}', bold=False, sz=1000, align="ctr"))
    if org.get("activity"):
        ps.append(_para(f'Виды деятельности: {org["activity"]}',
                        bold=False, sz=950, align="l"))
    if org.get("address"):
        ps.append(_para(f'Юр.адрес: {org["address"]}',
                        bold=False, sz=950, align="l"))
    if not ps:
        ps.append(_para(name or "Организация", bold=True, sz=1200, align="ctr"))
    return ps


def hdr_lines(director):
    """Блок руководителя — все поля как в пример_1."""
    d   = director or {}
    fio = d.get("fio") or ""
    iin = d.get("iin") or ""
    tel = d.get("phone") or ""
    adr = d.get("address") or ""
    return [
        _para("Руководитель",                              bold=True,  sz=1000, align="l"),
        _para("",                                          bold=False, sz=900),
        _para(f"ФИО: {fio}"           if fio else "ФИО:", bold=False, sz=950, align="l"),
        _para(f"ИИН: {iin}"           if iin else "ИИН:",  bold=False, sz=950, align="l"),
        _para(f"Тел.номер: {tel}"     if tel else "Тел.номер:",  bold=False, sz=950, align="l"),
        _para(f"Адрес прописки: {adr}" if adr else "Адрес прописки:", bold=False, sz=950, align="l"),
    ]


def contract_label_lines(con):
    """Повёрнутая договорная метка — Arial 16pt красный (как в пример_1)."""
    parts = []
    if con.get("description"):
        parts.append(str(con["description"]))
    if con.get("amount"):
        amt = re.sub(r'^(Общая\s+сумма|Сумма)\s*:?\s*', '',
                     str(con["amount"]), flags=re.I).strip()
        parts.append(f"за сумму {amt}")
    if con.get("year"):
        parts.append(f"на {con['year']} г.")
    text = " ".join(parts)
    if not text:
        return None
    return [_para(text, bold=True, sz=1600, color="FF0000",
                  align="ctr", font="Arial")]


def vert_contract_lines(con):
    """Вертикальный договорной блок (между org и sub) — Times New Roman 12pt."""
    parts = []
    if con.get("description"):
        parts.append(str(con["description"]))
    if con.get("amount"):
        amt = re.sub(r'^(Общая\s+сумма|Сумма)\s*:?\s*', '',
                     str(con["amount"]), flags=re.I).strip()
        parts.append(amt)
    if con.get("year"):
        parts.append(f"- {con['year']}г.")
    text = " ".join(parts)
    if not text:
        return None
    return [_para(text, bold=False, sz=1200, align="ctr")]


# ══════════════════════════════════════════════════════
# ЧАСТЬ 5: РЕНДЕР СХЕМЫ
# ══════════════════════════════════════════════════════

def render(data, photos=None):
    """Строит drawing XML по Layout. Работает рекурсивно на любую глубину субподряда.

    photos — словарь из load_photos(); если для узла есть фото, вставляется
    реальная картинка, иначе рисуется плейсхолдер «ФОТО». Возвращает
    (drawing_xml, media) — media нужен pack_xlsx для вставки картинок в архив.
    """
    _ID[0] = 1
    parts  = []
    media  = []          # [{'rid','idx','ext','data'}] — реальные фото
    lay    = Layout(data)
    cust   = lay.cust
    proj   = (data.get("project_name") or "").strip()

    L = lay
    PHOTO_COLS = 5       # ширина фото-слота в ячейках (портрет ~5×высота блока)

    def add(xml):
        parts.append(xml)

    def emit_photo(key_name, node_data, fc, fr, tc, tr):
        """Реальное фото (если найдено в photos/) либо плейсхолдер под фото."""
        ph = match_photo(photos, node_data)
        if ph:
            idx = len(media) + 1
            rid = f"rId{idx}"
            media.append({"rid": rid, "idx": idx, "ext": ph["ext"], "data": ph["data"]})
            add(pic_anchor(rid, fc, fr, tc, tr))
        else:
            add(photo_box(f"photo_{key_name}", fc, fr, tc, tr))

    # ── Заголовок проекта ─────────────────────────────────────────────
    title_fc = L.cust_col - 2
    title_tc = L.cust_col + L.cust_cols + 2
    title_text = f"Проект {proj}" if proj else "Проект"
    add(rect_fill(
        "Заголовок", title_fc, 1, title_tc, 4,
        [_para(title_text, bold=True, sz=2800, color="FF0000", align="ctr")],
        bg_color="FFFFFF",
        fc_off=76404, fr_off=60307,
        tc_off=564776, tr_off=59070
    ))

    # ── Заказчик ──────────────────────────────────────────────────────
    cc, cr = L.cust_col, L.cust_row
    add(org_box(
        "Заказчик",
        cc, cr, cc + L.cust_cols, cr + L.cust_rows,
        org_lines(cust, is_customer=True),
        fc_off=319927, fr_off=26894,
        tc_off=295435, tr_off=32229
    ))

    cust_center = cc + L.cust_cols // 2
    add(arrow("arr_title_cust",
              cust_center, 4, cust_center, cr,
              fc_off=10114, fr_off=19406,
              tc_off=22412, tr_off=11207))

    # ── Рекурсивная отрисовка ветки (исполнитель + его субподряд любой глубины) ──
    def draw_node(node, is_root_contractor, idx=None):
        col, row = node.col, node.row
        cols, rows = node.cols, node.rows
        is_right = node.side == "right"
        center = col + cols // 2
        node_id = f"{'con' if is_root_contractor else 'sub'}_{idx if idx is not None else id(node)}"

        if is_root_contractor:
            # Договорная метка (повёрнутая) — между заказчиком и прямым исполнителем
            lbl = contract_label_lines(node.data)
            if lbl:
                rot = -20194346 if is_right else 20194346
                if is_right:
                    lbl_fc = cc + L.cust_cols + 1
                    lbl_tc = lbl_fc + L.CON_COLS
                else:
                    lbl_tc = cc - 1
                    lbl_fc = max(1, lbl_tc - L.CON_COLS)
                lbl_fr = cr + 2
                lbl_tr = lbl_fr + L.CON_ROWS
                add(rect_fill(
                    f"lbl_{node_id}", lbl_fc, lbl_fr, lbl_tc, lbl_tr,
                    lbl, bg_color="FFFFFF", rot=rot,
                    fc_off=519213, fr_off=132617,
                    tc_off=169713, tr_off=163136
                ))
            add(org_box(
                f"con_{node_id}_{node.data.get('name','')[:15]}",
                col, row, col + cols, row + rows,
                org_lines(node.data),
                fc_off=602425, fr_off=82882,
                tc_off=589190, tr_off=173990
            ))
            add(arrow(f"arr_cust_{node_id}", cust_center, cr, center, row,
                      fc_off=291008, fr_off=29561, tc_off=319927, tr_off=82882))
        else:
            # Вертикальный договорной блок над субподрядчиком (описание/сумма)
            vct = vert_contract_lines(node.data)
            parent_bottom = node._parent_bottom
            if vct:
                ct_fr = parent_bottom + 1
                ct_tr = ct_fr + 5
                add(org_box(f"contract_{node_id}", col, ct_fr, col + cols, ct_tr, vct,
                             fc_off=78441, fr_off=133349, tc_off=224117, tr_off=89646))
            add(arrow(f"arr_down_{node_id}", center, parent_bottom, center, row,
                      fc_off=10114, fr_off=164224, tc_off=14653, tr_off=14654))
            add(org_box(
                f"sub_{node_id}_{node.data.get('name','')[:15]}",
                col, row, col + cols, row + rows,
                org_lines(node.data),
                fc_off=16499, fr_off=0, tc_off=609402, tr_off=91108
            ))

        hdr_row = row + rows + 1
        add(org_box(
            f"hdr_{node_id}", col + 1, hdr_row, col + 1 + L.HDR_COLS, hdr_row + L.HDR_ROWS,
            hdr_lines(node.data.get("director")),
            fc_off=555866, fr_off=82881, tc_off=589189, tr_off=168606
        ))

        # Нарушения — на внешней стороне от центральной оси (left/right ветки)
        for vi, viol in enumerate(node.data.get("violations") or []):
            vc = (col + cols + 1) if is_right else max(1, col - L.VIO_COLS - 1)
            vr = row + vi * (L.VIO_ROWS + 1)
            add(rect_fill(
                f"vio_{node_id}_{vi}", vc, vr, vc + L.VIO_COLS, vr + L.VIO_ROWS,
                [_para(str(viol), bold=True, sz=1600, color="FFFFFF", align="ctr", font="Arial")],
                bg_color="FF0000",
                fc_off=497276, fr_off=123700, tc_off=381000, tr_off=1236
            ))
            add(arrow(f"arr_vio_{node_id}_{vi}",
                      vc + L.VIO_COLS // 2, vr + L.VIO_ROWS // 2, col, row + rows // 2,
                      fc_off=291008, fr_off=29561, tc_off=319927, tr_off=82882))

        # Фото руководителя — справа от блока (портрет высотой с org-блок).
        # Если справа есть нарушения (right-ветка) — сдвигаем фото за них.
        ph_fc = col + cols + 1
        if is_right and (node.data.get("violations")):
            ph_fc += L.VIO_COLS + 2
        emit_photo(node_id, node.data, ph_fc, row, ph_fc + PHOTO_COLS, row + rows)

        node_bottom = hdr_row + L.HDR_ROWS + 2
        for child in node.children:
            child._parent_bottom = node_bottom
            draw_node(child, is_root_contractor=False)

    for idx, tree in enumerate(L.trees):
        draw_node(tree, is_root_contractor=True, idx=idx)

    # ── Сборка XML ─────────────────────────────────────────────────────
    NS = (
        'xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
    )
    drawing = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<xdr:wsDr {NS}>'
        + "".join(parts)
        + '</xdr:wsDr>'
    )
    ET.fromstring(drawing)  # валидация
    return drawing, media


# ══════════════════════════════════════════════════════
# ЧАСТЬ 6: УПАКОВКА В XLSX
# ══════════════════════════════════════════════════════

def pack_xlsx(drawing_xml, total_cols, total_rows, media=None):
    from openpyxl import Workbook
    from openpyxl.utils import get_column_letter
    media = media or []

    wb = Workbook()
    ws = wb.active
    ws.title = "Схема"

    # Настройки страницы как в пример_1 (scale=27, landscape)
    ws.sheet_view.zoomScale   = 60
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize   = 9    # A4

    # Размеры колонок/строк соответствующие EMU константам
    col_chars = COL / (7 * 9525)   # ~9.2 символа
    row_pts   = ROW / 12700         # ~15.2 pt

    for col in range(1, total_cols + 8):
        ws.column_dimensions[get_column_letter(col)].width = col_chars
    for row in range(1, total_rows + 10):
        ws.row_dimensions[row].height = row_pts

    buf = io.BytesIO()
    wb.save(buf)

    base = {}
    with zipfile.ZipFile(io.BytesIO(buf.getvalue())) as zf:
        for name in zf.namelist():
            base[name] = zf.read(name)

    # Патчим sheet1.xml — добавляем xmlns:r и drawing ref
    sx = base['xl/worksheets/sheet1.xml'].decode('utf-8')
    if 'xmlns:r' not in sx:
        sx = sx.replace(
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
            ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        )
    if '<drawing' not in sx:
        sx = sx.replace('</worksheet>', '<drawing r:id="rId1"/></worksheet>')
    base['xl/worksheets/sheet1.xml'] = sx.encode('utf-8')

    # Content_Types
    cx = base['[Content_Types].xml'].decode('utf-8')
    if 'drawing1.xml' not in cx:
        cx = cx.replace(
            '</Types>',
            '<Override PartName="/xl/drawings/drawing1.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>'
            '</Types>'
        )
    # Default-типы для картинок (нужны, если есть реальные фото)
    exts = {m['ext'] for m in media}
    for ext, ctype in (("png", "image/png"), ("jpg", "image/jpeg"),
                       ("jpeg", "image/jpeg")):
        if ext in exts and f'Extension="{ext}"' not in cx:
            cx = cx.replace(
                '</Types>',
                f'<Default Extension="{ext}" ContentType="{ctype}"/></Types>'
            )
    base['[Content_Types].xml'] = cx.encode('utf-8')

    # Rels
    rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" '
        'Target="../drawings/drawing1.xml"/>'
        '</Relationships>'
    )
    # drawing.rels — ОБЯЗАТЕЛЕН (без него Excel ломается).
    # Если есть реальные фото — добавляем связи на media/imageN.
    rel_items = ""
    for m in media:
        fn = f"image{m['idx']}.{m['ext']}"
        base[f"xl/media/{fn}"] = m["data"]
        rel_items += (
            f'<Relationship Id="{m["rid"]}" '
            f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
            f'Target="../media/{fn}"/>'
        )
    drawing_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'{rel_items}</Relationships>'
    )
    base['xl/drawings/drawing1.xml']            = drawing_xml.encode('utf-8')
    base['xl/worksheets/_rels/sheet1.xml.rels'] = rels.encode('utf-8')
    base['xl/drawings/_rels/drawing1.xml.rels'] = drawing_rels.encode('utf-8')

    out = io.BytesIO()
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zf:
        for name, content in base.items():
            zf.writestr(name, content)
    return out.getvalue()


# ══════════════════════════════════════════════════════
# ГЛАВНАЯ ФУНКЦИЯ
# ══════════════════════════════════════════════════════

def sep(t): print(f"\n{'═'*55}\n  {t}\n{'═'*55}")


def run(docx_path, use_llm=True, anonymize=False, graph=False, html=False):
    base     = os.path.splitext(os.path.basename(docx_path))[0]
    if anonymize:
        base += "_ANON"
    out_dir  = os.path.dirname(os.path.abspath(docx_path))
    json_out = os.path.join(out_dir, f"{base}_entities.json")
    xlsx_out = os.path.join(out_dir, f"{base}_scheme.xlsx")

    sep("ШАГ 1: ЧИТАЕМ WORD")
    print(f"  {docx_path}")
    try:
        text = read_docx(docx_path)
        print(f"  ✅ {len(text.split())} слов")
    except Exception as e:
        print(f"  ❌ {e}"); sys.exit(1)

    sep("ШАГ 2: ИЗВЛЕКАЕМ СТРУКТУРУ")

    sep("ШАГ 2a: ПОИСК ВСТРОЕННОЙ ГРАФ-МОДЕЛИ")
    embedded_graph = extract_graph_from_docx(docx_path)
    exact_mode = False
    data = None
    if embedded_graph:
        try:
            rebuilt = graph_to_tree(embedded_graph)
        except (ValueError, KeyError, TypeError):
            rebuilt = None
        if rebuilt and rebuilt.get("customer"):
            data = rebuilt
            exact_mode = True
            print("  ✅ Найдена встроенная графовая модель — точное восстановление "
                  "координат из оригинала (LLM не используется)")
        else:
            print("  ⚠️  Граф найден, но неполный — переходим к LLM/regex-извлечению")
    else:
        print("  ▶️  Встроенной графовой модели нет (документ создан вручную "
              "или старой версией скрипта) — извлекаем из текста")

    if not exact_mode:
        if use_llm:
            if llm_ok():
                print("  ✅ Qwen2.5 доступен")
                try:
                    data = extract_llm(text)
                except Exception as e:
                    print(f"  ⚠️  {e} — fallback")
            else:
                print("  ⚠️  Ollama недоступна — простое извлечение")

        if not data or not data.get("contractors"):
            if data is not None:
                print("  ⚠️  Мало данных из LLM — простое извлечение")
            data = extract_fallback(text)

    from validation import validate_document
    validate_document(data)
    sep("ШАГ 2b: ПРОВЕРКА ГРАФА (дедупликация, пустые узлы)")
    if exact_mode:
        print("  ▶️  Пропущено — узлы взяты из исходных фигур Excel, дублей быть не может")
    else:
        data, _warn = validate_and_dedup(data)

    def count_subs(nodes):
        n = 0
        for o in nodes:
            kids = o.get("subcontractors") or []
            n += len(kids) + count_subs(kids)
        return n

    def max_depth(nodes, d=0):
        if not nodes: return d
        return max(max_depth(o.get("subcontractors") or [], d + 1) for o in nodes)

    cust = data.get("customer") or {}
    cons = data.get("contractors") or []
    subs = count_subs(cons)
    print(f"  Проект:           {data.get('project_name') or '(не найден)'}")
    print(f"  Заказчик:         {cust.get('name','—')[:50]}")
    print(f"  Исполнителей:     {len(cons)}")
    print(f"  Субподрядчиков:   {subs} (макс. глубина цепочки: {max_depth(cons)})")

    if anonymize:
        from anon import anonymize_data
        data = anonymize_data(data)
        print("  🕶️  Режим обезличивания: все значения заменены на заглушки")

    sep("ШАГ 3: СОХРАНЯЕМ JSON")
    with open(json_out, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"  ✅ {json_out}")

    sep("ШАГ 4: ГЕНЕРИРУЕМ СХЕМУ")
    photos = {} if anonymize else load_photos(out_dir)
    if photos:
        print(f"  📷 Папка photos/: найдено {len(photos)} файлов — вставлю реальные фото")
    else:
        print("  📷 Папка photos/ не найдена — рисую плейсхолдеры «ФОТО»")
    try:
        drawing_xml, media = render(data, photos)
        lay         = Layout(data)
        xlsx_bytes  = pack_xlsx(drawing_xml, lay.total_cols, lay.total_rows, media)

        with open(xlsx_out, "wb") as f:
            f.write(xlsx_bytes)

        with zipfile.ZipFile(io.BytesIO(xlsx_bytes)) as zf:
            dr       = zf.read('xl/drawings/drawing1.xml').decode('utf-8')
            has_rels = 'xl/drawings/_rels/drawing1.xml.rels' in zf.namelist()
            n_sp     = dr.count('<xdr:sp ')
            n_cxn    = dr.count('<xdr:cxnSp ')
            n_pic    = dr.count('<xdr:pic>')
            n_media  = sum(1 for n in zf.namelist() if n.startswith('xl/media/'))

        print(f"  ✅ {xlsx_out}  ({os.path.getsize(xlsx_out)//1024} КБ)")
        print(f"     Фигур: {n_sp}, Стрелок: {n_cxn}, Фото/слотов: {n_pic}, "
              f"встроено картинок: {n_media}")
        print(f"     drawing.rels: {'✅' if has_rels else '❌'}")

    except ET.ParseError as e:
        print(f"  ❌ XML ошибка: {e}"); traceback.print_exc(); sys.exit(1)
    except Exception as e:
        print(f"  ❌ {e}"); traceback.print_exc(); sys.exit(1)

    if graph:
        from writers.graph_svg import build_svg
        svg_out = os.path.join(out_dir, f"{base}_graph.svg")
        with open(svg_out, "w", encoding="utf-8") as f:
            f.write(build_svg(data))
        print(f"  ✅ Граф связей (SVG): {svg_out}")

    if html:
        from writers.graph_html import build_html
        html_out = os.path.join(out_dir, f"{base}_editor.html")
        with open(html_out, "w", encoding="utf-8") as f:
            f.write(build_html(data))
        print(f"  ✅ Редактор (HTML):   {html_out}")

    sep("ГОТОВО")
    print(f"\n  📄 Word:   {docx_path}")
    print(f"  📋 JSON:   {json_out}")
    print(f"  📊 Схема:  {xlsx_out}\n")


if __name__ == "__main__":
    print()
    print("╔══════════════════════════════════════════════════════════╗")
    print("║  Word Справка → Excel схема  (v5 — точный формат)       ║")
    print("╚══════════════════════════════════════════════════════════╝")

    args    = list(sys.argv[1:])
    use_llm = "--no-llm" not in args
    if not use_llm: args.remove("--no-llm")
    anon = "--anonymize" in args
    if anon: args.remove("--anonymize")
    graph = "--graph" in args
    if graph: args.remove("--graph")
    html = "--html" in args
    if html: args.remove("--html")

    if not args:
        print("\n  Использование:")
        print("    python word_to_excel_v5.py справка.docx")
        print("    python word_to_excel_v5.py справка.docx --no-llm")
        print("    python word_to_excel_v5.py справка.docx --anonymize")
        print("    python word_to_excel_v5.py справка.docx --graph   # статичный SVG")
        print("    python word_to_excel_v5.py справка.docx --html    # интерактивный редактор\n")
        sys.exit(1)

    path = args[0]
    if not os.path.exists(path):
        print(f"\n  ❌ Не найден: {path}"); sys.exit(1)

    missing = []
    for pkg, imp in [("python-docx","docx"),("openpyxl","openpyxl"),("requests","requests")]:
        try: __import__(imp)
        except ImportError: missing.append(pkg)
    if missing:
        print(f"\n  ❌ pip install {' '.join(missing)}"); sys.exit(1)

    run(path, use_llm=use_llm, anonymize=anon, graph=graph, html=html)