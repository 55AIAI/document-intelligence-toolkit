#!/usr/bin/env python3
"""
Универсальный конвертер: Excel-схема ⇄ Word-справка.

Направление выбирается автоматически по расширению файла:
    .xlsx / .xlsm  →  Word-справка (.docx)
    .docx          →  Excel-схема  (.xlsx)
    .json          →  пересборка граф SVG + Excel-схема из отредактированных данных

Запуск:
    python main.py путь/к/файлу.xlsx
    python main.py путь/к/файлу.docx
    python main.py путь/к/файлу.docx --no-llm       # без Qwen, только парсер
    python main.py путь/к/файлу.docx --anonymize    # обезличить всё (для показа наружу)
    python main.py путь/к/файлу.docx --graph         # + граф связей в SVG (третий формат)

Модель Qwen 2.5 настраивается в config.py.
Данные никуда не отправляются — всё считается локально,
LLM работает через локальный Ollama.
"""

import os
import sys

from config import OLLAMA_URL, MODEL


def main():
    import argparse
    parser = argparse.ArgumentParser(description="Convert Word reports and Excel relationship diagrams.")
    parser.add_argument("file", help="Input .docx, .xlsx, .xlsm or structured .json")
    parser.add_argument("--no-llm", action="store_true", help="Use deterministic extraction only")
    parser.add_argument("--anonymize", action="store_true", help="Mask supported fields and discard saved canvas data")
    parser.add_argument("--graph", action="store_true", help="Also export an SVG graph")
    parser.add_argument("--html", action="store_true", help="Also export an interactive HTML editor")
    parser.add_argument("--version", action="version", version="Document Intelligence Toolkit 0.2.0")
    args = parser.parse_args()
    path = args.file
    use_llm, anonymize, graph, html = not args.no_llm, args.anonymize, args.graph, args.html
    if not os.path.isfile(path):
        parser.error("Input file does not exist")

    ext = os.path.splitext(path)[1].lower()

    print(f"⚙️  Модель: {MODEL} @ {OLLAMA_URL}  (LLM: {'вкл' if use_llm else 'выкл'})")

    if ext in (".xlsx", ".xlsm"):
        print("➡️  Направление: Excel-схема → Word-справка\n")
        from converters.excel_to_word import run
        run(path, use_llm=use_llm, anonymize=anonymize, graph=graph, html=html)

    elif ext == ".docx":
        print("➡️  Направление: Word-справка → Excel-схема\n")
        from converters.word_to_excel import run
        run(path, use_llm=use_llm, anonymize=anonymize, graph=graph, html=html)

    elif ext == ".json":
        # Пересборка из отредактированных данных: правишь *_entities.json —
        # получаешь обновлённые граф (SVG) и Excel-схему, без разбора документа.
        print("➡️  Пересборка из данных (JSON) → редактор + граф SVG + Excel-схема\n")
        rebuild_from_json(path, anonymize=anonymize)

    else:
        print(f"❌ Не понимаю расширение '{ext}'. Ожидаю .xlsx, .docx или .json")
        sys.exit(1)


def rebuild_from_json(path, anonymize=False):
    import json
    out_dir = os.path.dirname(os.path.abspath(path))
    base    = os.path.splitext(os.path.basename(path))[0]
    for suf in ("_entities", "_data"):          # чистим служебные хвосты в имени
        if base.endswith(suf):
            base = base[:-len(suf)]
    with open(path, encoding="utf-8") as f:
        data = json.load(f)
    if isinstance(data, dict) and "_editor" in data and "contractors" not in data:
        if anonymize:
            raise ValueError("Canvas-only state cannot be reliably anonymized; use structured source JSON")
        from writers.graph_html import build_html
        html_out = os.path.join(out_dir, f"{base}_editor.html")
        with open(html_out, "w", encoding="utf-8") as output:
            output.write(build_html(data))
        print(f"Editor restored: {html_out}. Export SVG/PNG/JPEG using its toolbar.")
        return
    from converters.word_to_excel import validate_and_dedup
    data, _warnings = validate_and_dedup(data)
    if anonymize:
        from anon import anonymize_data
        data = anonymize_data(data)
        base += "_ANON"

    # 1a) интерактивный редактор (двигать блоки, размер, фото)
    from writers.graph_html import build_html
    html_out = os.path.join(out_dir, f"{base}_editor.html")
    with open(html_out, "w", encoding="utf-8") as f:
        f.write(build_html(data))
    print(f"  ✅ Редактор (HTML):    {html_out}")

    # 1b) граф связей (статичный SVG)
    from writers.graph_svg import build_svg
    svg_out = os.path.join(out_dir, f"{base}_graph.svg")
    with open(svg_out, "w", encoding="utf-8") as f:
        f.write(build_svg(data))
    print(f"  ✅ Граф связей (SVG):  {svg_out}")

    # 2) Excel-схема
    try:
        from converters.word_to_excel import render, pack_xlsx, Layout
        drawing_xml, media = render(data)
        lay = Layout(data)
        xlsx = pack_xlsx(drawing_xml, lay.total_cols, lay.total_rows, media)
        xlsx_out = os.path.join(out_dir, f"{base}_scheme.xlsx")
        with open(xlsx_out, "wb") as f:
            f.write(xlsx)
        print(f"  ✅ Excel-схема:        {xlsx_out}")
    except Exception as e:
        print(f"  ⚠️  Excel-схему собрать не удалось: {e} (граф SVG готов)")


if __name__ == "__main__":
    try:
        main()
    except (ValueError, OSError) as exc:
        print(f"Conversion failed: {exc}", file=sys.stderr)
        sys.exit(1)
