"""
Обезличивание данных перед показом (мне/наружу).

Заменяет все чувствительные значения (названия, БИН, ИИН, ФИО, суммы,
адреса, нарушения, доп. поля) на нейтральные заглушки, НО сохраняет
структуру: число организаций, вложенность субподряда, связи, наличие
полей, длину чисел. Схема, собранная из обезличенных данных, выглядит
и «строится» так же, как настоящая — но не раскрывает информацию.

Работает и с деревом word_to_excel (ключи address/director),
и со структурой excel_to_word (ключи legal_address/header).
"""

import re


def _mask_digits(s):
    """Заменяет цифры на #, сохраняя длину и единицы (БИН, суммы)."""
    return re.sub(r"\d", "#", str(s)) if s else s


def _anon_org(node, kind, counter):
    if not isinstance(node, dict):
        return node

    # Название — по роли + номер
    if kind == "customer":
        label = "Заказчик"
    else:
        counter["n"] += 1
        label = f"Организация {counter['n']}"
    if node.get("name"):
        node["name"] = label

    for k in ("bin", "iin"):
        if node.get(k):
            node[k] = _mask_digits(node[k])
    for key in list(node):
        if key.startswith("_"):
            node.pop(key)
    for key in ("amount", "year"):
        if node.get(key):
            node[key] = _mask_digits(node[key])
    if node.get("description"):
        node["description"] = "предмет договора"
    if node.get("activity"):
        node["activity"] = "вид деятельности"
    for k in ("legal_address", "address"):
        if node.get(k):
            node[k] = "адрес"
    if node.get("extra"):
        node["extra"] = [f"доп. поле {i+1}" for i in range(len(node["extra"]))]

    # Руководитель (header / director)
    for hk in ("header", "director"):
        h = node.get(hk)
        if isinstance(h, dict):
            if h.get("fio"):     h["fio"] = "Ф.И.О."
            if h.get("iin"):     h["iin"] = _mask_digits(h["iin"])
            if h.get("phone"):   h["phone"] = _mask_digits(h["phone"])
            if h.get("address"): h["address"] = "адрес"

    # Договор
    c = node.get("contract")
    if isinstance(c, dict):
        if c.get("description"):
            c["description"] = "предмет договора"
        if c.get("amount"):
            # маскируем цифры, единицы (млн/тг/год) сохраняем
            c["amount"] = _mask_digits(c["amount"])

        if c.get("year"):
            c["year"] = _mask_digits(c["year"])

    # Нарушения
    if node.get("violations"):
        node["violations"] = [f"нарушение {i+1}" for i in range(len(node["violations"]))]

    # Субподрядчики — рекурсивно
    for sub in node.get("subcontractors") or []:
        _anon_org(sub, "subcontractor", counter)

    return node


def anonymize_data(data):
    """Обезличивает дерево/структуру на месте и возвращает его же."""
    if not isinstance(data, dict):
        return data
    # Canvas state may contain raw text and base64 photos; rebuild from masked data.
    for key in list(data):
        if key.startswith("_"):
            data.pop(key)
    counter = {"n": 0}
    if data.get("project_name"):
        data["project_name"] = "Проект"
    if isinstance(data.get("customer"), dict):
        _anon_org(data["customer"], "customer", counter)
    for con in data.get("contractors") or []:
        _anon_org(con, "contractor", counter)
    if data.get("violations"):
        data["violations"] = [f"нарушение {i+1}" for i in range(len(data["violations"]))]
    return data
