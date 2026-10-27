# Architecture

The CLI chooses a converter from the file extension. DOCX input first checks for
an embedded graph produced by this toolkit; otherwise text is extracted from
paragraphs and tables. Optional Ollama extraction returns JSON, which must pass
shape and literal-support checks. A rule-based parser is the fallback.

The current interchange model is a tree of organization **occurrences**:
customer, contractors, and recursively nested subcontractors. Each occurrence
describes its relationship to its parent and may have its own amount, year and
contract description. Two appearances of the same legal entity are therefore
valid. Only exact sibling occurrences are merged; different contracts and
different parents remain separate. It is not a general-purpose entity database.

JSON is validated before rebuilding outputs. Renderers produce Excel drawing XML,
SVG, and self-contained HTML. The HTML editor supports moving/editing nodes,
link creation and export. Graph cycles are reported rather than traversed forever.
Document text is escaped when embedded in scripts and inserted into HTML.

All Ollama calls go through llm/qwen.py. Core conversion needs python-docx,
openpyxl and requests; no database, cloud service or LLM account is required.
