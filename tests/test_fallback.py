#!/usr/bin/env python3
"""
Регрессионный тест универсального извлечения (Word → структура) БЕЗ LLM.

Генерит синтетические Word-документы РАЗНЫХ форматов (не только шаблонная
справка) и проверяет, что extract_fallback корректно вытаскивает заказчика,
исполнителей, субподряд, БИН, суммы, руководителей, нарушения. Реальных данных
не использует — всё синтетическое.

Запуск:  python tests/test_fallback.py
"""

import os
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from converters.word_to_excel import read_docx, extract_fallback


def _make_docs(d):
    from docx import Document
    # A: шаблонная справка (кавычки, метки)
    doc = Document()
    doc.add_heading("Справка по проекту «Цифровизация»", 0)
    doc.add_paragraph('Заказчик: ГУ «Аппарат акима района» БИН 010101010101.')
    doc.add_paragraph('Исполнитель: ТОО «Альфа» БИН 111111111111, сумма 500 млн.тг. '
                      'Руководитель ФИО: Иванов Иван Иванович, ИИН: 800101300123, тел: +77010000000.')
    doc.add_paragraph('ТОО «Альфа» привлекло ТОО «Бета» БИН 222222222222 в качестве субподрядчика '
                      'на 200 млн.тг. Руководитель ФИО: Петров Пётр, ИИН: 900202400234.')
    doc.add_paragraph('Исполнитель: ТОО «Гамма» БИН 333333333333, сумма 150 млн.тг. '
                      'Нарушение: не соответствует ТЗ.')
    doc.save(os.path.join(d, "A.docx"))
    # B: проза без кавычек, полные слова
    doc = Document()
    doc.add_paragraph("Заказчиком выступает Государственное учреждение Управление образования, БИН 020202020202.")
    doc.add_paragraph("Договор заключён с Товариществом с ограниченной ответственностью СтройИнвест "
                      "(БИН 444444444444) на сумму 1,2 млрд тенге. Директор — Сидоров Сергей Сергеевич, ИИН 850303500345.")
    doc.add_paragraph("СтройИнвест передал часть работ субподрядной организации ИП Каиров "
                      "(БИН 555555555555) стоимостью 300 миллионов тенге.")
    doc.save(os.path.join(d, "B.docx"))
    # C: таблица
    doc = Document()
    t = doc.add_table(rows=1, cols=4); t.style = 'Table Grid'
    h = t.rows[0].cells
    h[0].text, h[1].text, h[2].text, h[3].text = "Организация", "БИН", "Сумма", "Руководитель"
    for nm, bn, su, ru in [("АО НацКомпания", "606060606060", "5 млрд тг", "Абаев А.А."),
                           ("ТОО ТехноПром", "707070707070", "800 млн тг", "Жумаев Ж.Ж.")]:
        r = t.add_row().cells; r[0].text, r[1].text, r[2].text, r[3].text = nm, bn, su, ru
    doc.save(os.path.join(d, "C.docx"))
    # D: минимальный список
    doc = Document()
    for line in ["1. Заказчик — Акимат области", "2. ТОО Вектор, 90 млн тг", "3. ТОО Горизонт, 40 млн тг"]:
        doc.add_paragraph(line)
    doc.save(os.path.join(d, "D.docx"))
    # E: документ на казахском (ЖШС, БСН/ЖСН, басшы, әкімдік, теңге)
    doc = Document()
    doc.add_paragraph("Тапсырыс беруші: Ақтөбе қаласының әкімдігі, БСН 191919191919.")
    doc.add_paragraph("Орындаушы: ЖШС «Құрылыс» БСН 202020202020, сомасы 400 млн теңге. "
                      "Басшысы: Ахметов Асхат Серікұлы, ЖСН 900101300111.")
    doc.add_paragraph("ЖШС «Құрылыс» ЖШС «Жөндеу» БСН 212121212121 қосалқы мердігер ретінде тартты, "
                      "150 млн теңге.")
    doc.save(os.path.join(d, "E.docx"))
    # F: несколько заказчиков (гос-органов)
    doc = Document()
    doc.add_paragraph("Заказчики: ГУ «Управление финансов» БИН 232323232323 и "
                      "ГУ «Управление строительства» БИН 242424242424.")
    doc.add_paragraph("Исполнитель: ТОО «Генподрядчик» БИН 252525252525, сумма 800 млн тг.")
    doc.save(os.path.join(d, "F.docx"))


def _count_subs(nodes):
    return sum(len(o.get("subcontractors") or []) + _count_subs(o.get("subcontractors") or []) for o in nodes)


CHECKS = {
    "A": lambda d: (d["customer"] and d["customer"]["bin"] == "010101010101"
                    and len(d["contractors"]) == 2 and _count_subs(d["contractors"]) == 1),
    "B": lambda d: (d["customer"] and "образовани" in (d["customer"]["name"] or "").lower()
                    and len(d["contractors"]) == 1 and _count_subs(d["contractors"]) == 1),
    "C": lambda d: (d["customer"] and len(d["contractors"]) >= 1
                    and any(c.get("bin") for c in d["contractors"])),
    "D": lambda d: (d["customer"] and "аким" in (d["customer"]["name"] or "").lower()
                    and len(d["contractors"]) == 2),
    # E (казахский): заказчик-әкімдік, ЖШС Құрылыс → субподряд Жөндеу, БСН и басшы извлечены
    "E": lambda d: (d["customer"] and "әкім" in (d["customer"]["name"] or "").lower()
                    and _count_subs(d["contractors"]) == 1
                    and any(c.get("bin") for c in d["contractors"])
                    and any((c.get("director") or {}).get("fio") for c in d["contractors"])),
    # F (несколько заказчиков): все 3 организации извлечены, заказчик — гос-орган
    "F": lambda d: (d["customer"] and "управлени" in (d["customer"]["name"] or "").lower()
                    and (len(d["contractors"]) + _count_subs(d["contractors"])) == 2),
}


def main():
    with tempfile.TemporaryDirectory() as d:
        _make_docs(d)
        ok = True
        for name in ["A", "B", "C", "D", "E", "F"]:
            data = extract_fallback(read_docx(os.path.join(d, name + ".docx")))
            passed = False
            try:
                passed = bool(CHECKS[name](data))
            except Exception as e:
                print(f"  формат {name}: ОШИБКА {e}")
            cust = (data.get("customer") or {}).get("name")
            print(f"  формат {name}: {'✅' if passed else '❌'}  "
                  f"заказчик={cust!r} исполнителей={len(data['contractors'])} "
                  f"субподряд={_count_subs(data['contractors'])}")
            ok = ok and passed
        print("\nИТОГ:", "ВСЕ ФОРМАТЫ OK ✅" if ok else "ЕСТЬ ПРОВАЛЫ ❌")
        sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
