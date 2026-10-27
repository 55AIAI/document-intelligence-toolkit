"""
Третий тип вывода: ГРАФ СВЯЗЕЙ в SVG — в стиле пример_1.xlsx.

Каждая организация = ДВА блока, как в оригинальной схеме:
  1) блок организации: название, БИН, Виды деятельности, Юр.адрес;
  2) блок «Руководитель» с метками: ФИО / ИИН / Тел.номер / Адрес прописки.
Рядом с руководителем — слот под фото (человек вставит после скачивания).
Договор — подпись с суммой на стрелке. Нарушения — красные блоки сбоку.

Связи берутся из МОДЕЛИ ДАННЫХ (дерево заказчик → исполнители → субподряд),
поэтому стрелки всегда соответствуют реальным связям, а не вычисляются по
координатам, как в Excel.

Вход — тот же dict, что у конвертеров (customer/contractors, header|director).
"""

# ── Геометрия ───────────────────────────────────────────────────────────
BOX_W   = 250
ORG_H   = 78                 # высота блока организации
DIR_H   = 96                 # высота блока руководителя
DGAP    = 6                  # зазор между org и руководителем
NODE_H  = ORG_H + DGAP + DIR_H
PHOTO_W = 58                 # слот под фото справа от руководителя
VIO_W, VIO_H = 132, 42       # блок нарушения (слева)
H_GAP   = 46
V_GAP   = 64                 # вертикальный зазор (место под подпись договора)
MARGIN  = 60
ROW     = NODE_H + V_GAP
# ширина ячейки одного листа: [нарушения] [org] [фото] + зазор
UNIT    = VIO_W + 12 + BOX_W + 8 + PHOTO_W + H_GAP
LEFT    = VIO_W + 12          # отступ слева внутри ячейки до org-блока

ROLE_STYLE = {
    "customer":      ("#2f5597", "#dfe8f5"),
    "contractor":    ("#548235", "#eaf3ea"),
    "subcontractor": ("#bf8f00", "#fdf4e3"),
}
ROLE_RU = {"customer": "Заказчик", "contractor": "Исполнитель",
           "subcontractor": "Субподрядчик"}


def _esc(s):
    return (str(s).replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;").replace('"', "&quot;"))


def _trunc(s, n):
    s = str(s)
    return s if len(s) <= n else s[:n - 1].rstrip() + "…"


def _wrap(text, width, max_lines):
    words, lines, cur = str(text).split(), [], ""
    for w in words:
        if len(cur) + len(w) + 1 <= width:
            cur = (cur + " " + w).strip()
        else:
            lines.append(cur); cur = w
            if len(lines) == max_lines:
                break
    if cur and len(lines) < max_lines:
        lines.append(cur)
    if len(lines) == max_lines and len(" ".join(words)) > len(" ".join(lines)):
        lines[-1] = lines[-1][:width - 1].rstrip() + "…"
    return lines or [""]


def _txt(x, y, s, size=11, bold=False, fill="#222", anchor="start", italic=False):
    weight = 'font-weight="bold" ' if bold else ""
    style = 'font-style="italic" ' if italic else ""
    return (f'<text x="{x:.0f}" y="{y:.0f}" font-size="{size}" '
            f'fill="{fill}" text-anchor="{anchor}" '
            f'{weight}{style}font-family="Arial">{_esc(s)}</text>')


# ── Дерево и раскладка ──────────────────────────────────────────────────
def _build_tree(data):
    cust = dict(data.get("customer") or {"name": "Заказчик"})
    cust["_role"] = "customer"
    cust["subcontractors"] = list(data.get("contractors") or [])

    def mark(n, role):
        n["_role"] = role
        for k in n.get("subcontractors") or []:
            mark(k, "subcontractor")
    for c in cust["subcontractors"]:
        mark(c, "contractor")
    return cust


def _layout(node, depth, leaf):
    node["_y"] = MARGIN + depth * ROW
    kids = node.get("subcontractors") or []
    if not kids:
        node["_cx"] = MARGIN + LEFT + leaf[0] * UNIT + BOX_W / 2
        leaf[0] += 1
    else:
        for k in kids:
            _layout(k, depth + 1, leaf)
        node["_cx"] = (kids[0]["_cx"] + kids[-1]["_cx"]) / 2
    d = depth
    for k in kids:
        d = max(d, k["_maxd"])
    node["_maxd"] = d
    return node


# ── Отрисовка узла ──────────────────────────────────────────────────────
def _node_svg(node):
    role = node.get("_role", "subcontractor")
    stroke, fill = ROLE_STYLE.get(role, ROLE_STYLE["subcontractor"])
    x  = node["_cx"] - BOX_W / 2
    y  = node["_y"]
    out = []

    # ── блок организации ──
    out.append(f'<rect x="{x:.0f}" y="{y:.0f}" width="{BOX_W}" height="{ORG_H}" '
               f'rx="4" fill="{fill}" stroke="{stroke}" stroke-width="2"/>')
    out.append(f'<rect x="{x:.0f}" y="{y:.0f}" width="{BOX_W}" height="18" '
               f'fill="{stroke}"/>')
    out.append(_txt(x + 7, y + 13, ROLE_RU.get(role, role), 10, fill="#fff"))
    ty = y + 33
    for ln in _wrap(node.get("name") or "—", 30, 2):
        out.append(_txt(x + 7, ty, ln, 12.5, bold=True)); ty += 15
    out.append(_txt(x + 7, ty, "БИН " + (_trunc(node.get("bin"), 18) if node.get("bin") else ""), 10, fill="#333")); ty += 13
    out.append(_txt(x + 7, ty, "Виды деятельности: " + _trunc(node.get("activity") or "", 24), 9.5, fill="#333")); ty += 12
    adr = node.get("legal_address") or node.get("address") or ""
    out.append(_txt(x + 7, ty, "Юр.адрес: " + _trunc(adr, 30), 9.5, fill="#333"))

    # ── блок руководителя ──
    dy = y + ORG_H + DGAP
    hdr = node.get("header") or node.get("director") or {}
    out.append(f'<rect x="{x:.0f}" y="{dy:.0f}" width="{BOX_W}" height="{DIR_H}" '
               f'rx="4" fill="#ffffff" stroke="{stroke}" stroke-width="1.4"/>')
    out.append(_txt(x + 7, dy + 15, "Руководитель", 10.5, bold=True))
    fields = [("ФИО: ", hdr.get("fio")), ("ИИН: ", hdr.get("iin")),
              ("Тел.номер: ", hdr.get("phone")), ("Адрес прописки: ", hdr.get("address"))]
    fy = dy + 32
    for lbl, val in fields:
        out.append(_txt(x + 7, fy, lbl + (_trunc(val, 22) if val else ""), 9.5, fill="#333"))
        fy += 15

    # ── слот под фото (справа от руководителя) ──
    px = x + BOX_W + 6
    out.append(f'<rect x="{px:.0f}" y="{dy:.0f}" width="{PHOTO_W}" height="{DIR_H}" '
               f'fill="#f2f2f2" stroke="#bbb" stroke-dasharray="4 3"/>')
    out.append(_txt(px + PHOTO_W / 2, dy + DIR_H / 2, "ФОТО", 9, fill="#999", anchor="middle"))

    # ── нарушения (красные блоки слева) ──
    for i, v in enumerate(node.get("violations") or []):
        vx = x - VIO_W - 10
        vy = y + i * (VIO_H + 6)
        out.append(f'<rect x="{vx:.0f}" y="{vy:.0f}" width="{VIO_W}" height="{VIO_H}" '
                   f'rx="3" fill="#c00000"/>')
        for j, ln in enumerate(_wrap(v, 22, 2)):
            out.append(_txt(vx + VIO_W / 2, vy + 17 + j * 13, ln, 9, fill="#fff", anchor="middle"))
        out.append(f'<line x1="{vx+VIO_W:.0f}" y1="{vy+VIO_H/2:.0f}" x2="{x:.0f}" '
                   f'y2="{y+ORG_H/2:.0f}" stroke="#c00000" stroke-width="1" stroke-dasharray="3 2"/>')
    return "".join(out)


def _edge_svg(parent, child):
    x1, y1 = parent["_cx"], parent["_y"] + NODE_H
    x2, y2 = child["_cx"], child["_y"]
    out = [f'<path d="M {x1:.0f} {y1:.0f} L {x2:.0f} {y2:.0f}" '
           f'stroke="#555" stroke-width="1.6" fill="none" marker-end="url(#arr)"/>']
    ct = child.get("contract") or {}
    desc = ct.get("description") or ""
    amt  = ct.get("amount") or child.get("amount") or ""
    if desc and amt:
        label = f"{_trunc(desc, 20)} — {_trunc(amt, 20)}"
    elif amt:
        label = _trunc(amt, 28)
    elif desc:
        label = _trunc(desc, 28)
    else:
        label = "субподряд" if child.get("_role") == "subcontractor" else "договор"
    mx, my = (x1 + x2) / 2, (y1 + y2) / 2
    w = int(6.2 * len(label)) + 12
    out.append(f'<rect x="{mx-w/2:.0f}" y="{my-9:.0f}" width="{w}" height="16" '
               f'rx="3" fill="#fff" stroke="#e0b4b4"/>')
    out.append(_txt(mx, my + 3, label, 9.5, fill="#c00000", anchor="middle", italic=True))
    return "".join(out)


def build_svg(data):
    root = _build_tree(data)
    leaf = [0]
    _layout(root, 0, leaf)
    width  = 2 * MARGIN + max(leaf[0], 1) * UNIT
    height = 2 * MARGIN + (root["_maxd"] + 1) * ROW

    parts = [
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{width:.0f}" '
        f'height="{height:.0f}" viewBox="0 0 {width:.0f} {height:.0f}" font-family="Arial">',
        '<defs><marker id="arr" markerWidth="10" markerHeight="10" refX="8" refY="3" '
        'orient="auto"><path d="M0,0 L8,3 L0,6 Z" fill="#555"/></marker></defs>',
        f'<rect width="{width:.0f}" height="{height:.0f}" fill="#ffffff"/>',
    ]
    proj = (data.get("project_name") or "").strip()
    if proj:
        parts.append(_txt(width / 2, 34, _trunc(proj, 90), 16, bold=True,
                          fill="#c00000", anchor="middle"))

    def walk_edges(n):
        for k in n.get("subcontractors") or []:
            parts.append(_edge_svg(n, k)); walk_edges(k)

    def walk_nodes(n):
        parts.append(_node_svg(n))
        for k in n.get("subcontractors") or []:
            walk_nodes(k)

    walk_edges(root)
    walk_nodes(root)
    parts.append("</svg>")
    return "\n".join(parts)
