"""Build a draft extraction table for the qualitative document analysis of curricula.

Retrieval step: finds candidate passages about programming tools (PTs) in the
page-marked text extracts (txt/txt1) and writes them, together with context
hints and the existing manual document-level codes, to an Excel workbook.
The workbook is meant to be verified and completed manually against the PDFs.

It also writes passages.json (passages + full page text + document metadata),
which 04_coding_viewer/build_viewer_data.py turns into data for viewer.html.

Run from anywhere:  python Curricula/03_extraction_table/build_extraction_table.py
"""

from __future__ import annotations

import hashlib
import json
from pathlib import Path

import pandas as pd
import regex as re
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

BASE = Path(__file__).resolve().parent.parent  # .../Curricula
TXT_DIR = BASE / "01_text_extraction" / "txt1"
PDF_DIR = BASE / "00_sources" / "pdfs"
MANUAL_DIR = BASE / "legacy" / "ai_coding_attempt" / "res" / "manual"  # earlier document-level coding (reference only)
INCLUDE_TABLE = BASE / "00_sources" / "corpus_include.csv"
OUT_DIR = Path(__file__).resolve().parent
OUT_FILE = OUT_DIR / "Extraktionstabelle_Entwurf.xlsx"
OUT_JSON = OUT_DIR / "passages.json"

PAGE_MARKER = re.compile(r"\[\[SEITE\s+(\d+)\]\]")

# ------------------------------------------------------------------
# Retrieval keywords (adapted from keyword-analysis.ipynb)
# Anchors and explicit tool names are searched everywhere; tool-type
# terms are specific enough to be searched everywhere as well.
# ------------------------------------------------------------------
RETRIEVAL = {
    # anchors
    "Programmierwerkzeug": r"Programmier-?\s?werkzeug\w*",
    "Programmiersprache": r"Programmier-?\s?sprache\w*",
    "Programmierumgebung": r"Programmier-?\s?umgebung\w*",
    "Entwicklungsumgebung": r"Entwicklungs-?\s?umgebung\w*",
    "Programmiersystem": r"Programmier-?\s?system\w*",
    "Skriptsprache": r"S[kc]ript(ing)?-?\s?sprache\w*",
    # tool types
    "blockbasiert": r"block-?\s?(basiert|orientiert)\w*|Bausteinprogrammier\w*",
    "Blöcke (Wort)": r"\bBlöcke\w*",  # recall only; often unrelated (e.g. cipher blocks)
    "textbasiert": r"text-?\s?basiert\w*|textuelle\w*\s+Programmier\w*",
    "visuell/grafisch": r"(visuelle|grafische|graphische)\w*\s+(Programmier\w*|Entwicklungsumgebung\w*|Sprache\w*)",
    # explicit tools
    "Scratch": r"\bScratch\b", "Snap!": r"\bSnap!", "Blockly": r"\bBlockly\b",
    "mBlock": r"\bmBlock\b", "Open Roberta": r"\bOpen\s*Roberta\b", "MakeCode": r"\bMakeCode\b",
    "NEPO": r"\bNEPO\b", "App Inventor": r"\bApp\s*Inventor\b",
    "Python": r"\bPython\b", "Java": r"\bJava\b(?!\s*-?Script)", "JavaScript": r"\bJava\s*-?Script\b",
    "C#": r"\bC\s?#", "C++": r"\bC\+\+", "Delphi": r"\bDelphi\b", "Pascal": r"\bPascal\b",
    "Haskell": r"\bHaskell\b", "Prolog": r"\bProlog\b", "PHP": r"\bPHP\b", "Lisp": r"\bLisp\b",
    "Scheme": r"\bScheme\b", "Visual Basic": r"\bVisual\s*Basic\b",
    "Robot Karol": r"\bRobot\s*Karol\b", "Kara": r"\bKara\b", "Logo": r"\bLogo\b(?!s)",
    "Hamster": r"\bHamster\b", "Greenfoot": r"\bGreenfoot\b", "BlueJ": r"\bBlueJ\b",
    "EOS": r"\bEOS\b", "Thonny": r"\bThonny\b", "Jupyter": r"\bJupyter\b",
    "Arduino": r"\bArduino\b", "Calliope": r"\bCalliope\b", "micro:bit": r"\bmicro\s*[:\-]?\s*bit\b",
    "Raspberry Pi": r"\bRaspberry\s*Pi\b", "Lego": r"\bLego\b", "Ozobot": r"\bOzobot\b",
}
CASE_SENSITIVE = {"EOS", "NEPO", "PHP", "Logo"}
# hits on these alone (without tool names, tool types or attribute words) get low priority
GENERIC_ANCHORS = {"Programmiersprache", "Programmierumgebung", "Entwicklungsumgebung", "Programmiersystem",
                   "Blöcke (Wort)"}
RETRIEVAL_RE = {
    k: re.compile(v, 0 if k in CASE_SENSITIVE else re.IGNORECASE) for k, v in RETRIEVAL.items()
}

# Hints shown to the verifier (never used as codes)
ATTRIBUTE_HINTS = {
    "geeignet": r"\bgeeignet\w*|\bangemessen\w*",
    "didaktisch reduziert": r"didaktisch\w*",
    "einfach": r"\beinfach\w*",
    "altersangemessen": r"alters(gerecht|angemessen)\w*|schülergerecht\w*|kindgerecht\w*",
    "aktuell": r"\baktuell\w*|\bmodern\w*",
    "Mittel zum Zweck": r"Produktschulung|kein Selbstzweck|nicht Selbstzweck|Mittel zum Zweck",
    "OOP/Paradigma": r"objekt-?\s?orientiert\w*|Paradigm\w*|funktional\w*|deklarativ\w*|imperativ\w*|logisch\w* Programmier\w*",
    "Motivation": r"motivier\w*|Motivation\w*",
    "Syntaxfehler": r"Syntaxfehler\w*|syntaktisch\w*",
    "Open Source": r"Open[\s-]?Source|freie\w* Software|quelloffen\w*",
}
BINDING_HINTS = {
    "verbindlich": r"verbindlich\w*|verpflichtend\w*|\bmuss\b|\bmüssen\b|\bist\s+zu\s+\w+|\bsind\s+zu\s+\w+",
    "Empfehlung": r"empfohlen|empfiehlt|\bsollte\w*|bietet\s+sich\s+an|bieten\s+sich\s+an",
    "Beispiel": r"z\.\s?B\.|beispielsweise|\betwa\b|\bwie\s+\w+\s+oder\b|exemplarisch|\bu\.\s?a\.",
    "Möglichkeit": r"\bkann\b|\bkönnen\b|möglich\w*",
}
CONTEXT_HINTS = {
    "Klassenstufe": r"(Klassen?(stufen?)?|Jahrgangsstufen?|Jahrgänge?|Jg\.)\s*\d{1,2}(\s*[-–/]\s*\d{1,2})?",
    "Sek I/II": r"Sekundarstufe\s+(I{1,2}|1|2)\b|Sek\.?\s*(I{1,2}|1|2)\b",
    "Oberstufenphase": r"Einführungsphase|Qualifikationsphase|Kursstufe|Kurshalbjahr\w*|Q[12]\b|E[12]\b",
    "Kursart": r"Wahlpflicht\w*|Wahlfach\w*|Leistungskurs\w*|Grundkurs\w*|Profil\w*|Aufbaukurs\w*|Basiskurs\w*",
    "Niveau": r"erhöhte\w*\s+Anforderungsniveau\w*|grundlegende\w*\s+Anforderungsniveau\w*|eA\b|gA\b",
}
ATTR_RE = {k: re.compile(v, re.IGNORECASE) for k, v in ATTRIBUTE_HINTS.items()}
# attribute words that raise priority when they occur in the hit sentence itself
PRIORITY_ATTRS = ["geeignet", "didaktisch reduziert", "einfach", "altersangemessen", "aktuell",
                  "Mittel zum Zweck", "Syntaxfehler", "Open Source"]
BIND_RE = {k: re.compile(v, re.IGNORECASE) for k, v in BINDING_HINTS.items()}
CTX_RE = {k: re.compile(v) for k, v in CONTEXT_HINTS.items()}

SENTENCE_END = re.compile(r"(?<=[.!?;:])\s+(?=[A-ZÄÖÜ•\-–])|\s*•\s*")
PASSAGE_RADIUS = 1  # sentences before/after the hit sentence


# ------------------------------------------------------------------
def normalize(page_text: str) -> str:
    """Join hyphenated line breaks and collapse whitespace."""
    t = re.sub(r"(\w)-\n(?=[a-zäöüß])", r"\1", page_text)
    t = re.sub(r"\n(?=\s*•)", " • ", t)
    t = re.sub(r"\s*\n\s*", " ", t)
    return re.sub(r"\s{2,}", " ", t).strip()


def parse_pages(text: str) -> list[tuple[int, str]]:
    markers = list(PAGE_MARKER.finditer(text))
    if not markers:
        return [(1, text)]
    pages = []
    for i, m in enumerate(markers):
        end = markers[i + 1].start() if i + 1 < len(markers) else len(text)
        pages.append((int(m.group(1)), text[m.end():end]))
    return pages


def sentence_spans(text: str) -> list[tuple[int, int]]:
    spans, start = [], 0
    for m in SENTENCE_END.finditer(text):
        if m.start() > start:
            spans.append((start, m.start()))
        start = m.end()
    if start < len(text):
        spans.append((start, len(text)))
    return spans


def find_passages(page_text: str) -> list[dict]:
    """Return merged passages (sentence windows) containing retrieval hits."""
    spans = sentence_spans(page_text)
    hit_sentences: dict[int, set[str]] = {}
    for name, rx in RETRIEVAL_RE.items():
        for m in rx.finditer(page_text):
            idx = next((i for i, (a, b) in enumerate(spans) if a <= m.start() < b), None)
            if idx is not None:
                hit_sentences.setdefault(idx, set()).add(name)
    if not hit_sentences:
        return []
    # build windows and merge overlapping ones
    windows = sorted(
        (max(0, i - PASSAGE_RADIUS), min(len(spans) - 1, i + PASSAGE_RADIUS), kws)
        for i, kws in hit_sentences.items()
    )
    merged: list[list] = []
    for lo, hi, kws in windows:
        if merged and lo <= merged[-1][1] + 1:
            merged[-1][1] = max(merged[-1][1], hi)
            merged[-1][2] |= kws
        else:
            merged.append([lo, hi, set(kws)])
    out = []
    for lo, hi, kws in merged:
        a, b = spans[lo][0], spans[hi][1]
        core = " ".join(page_text[spans[i][0]:spans[i][1]] for i in hit_sentences if lo <= i <= hi)
        out.append({"text": page_text[a:b].strip(), "core": core, "keywords": sorted(kws)})
    return out


def hints(regexes: dict, text: str) -> str:
    return "; ".join(k for k, rx in regexes.items() if rx.search(text))


def context_cues(page_text: str) -> str:
    cues = []
    for name, rx in CTX_RE.items():
        found = sorted({m.group(0).strip() for m in rx.finditer(page_text)})
        if found:
            cues.append(f"{name}: " + ", ".join(found[:6]))
    return " | ".join(cues)


def doc_meta(rel: Path) -> dict:
    parts = rel.parts
    if parts[0] == "Lehrpläne":
        state = parts[1]
        school = " / ".join(parts[2:-1])
    else:
        state, school = parts[0], ""
    return {"Gruppe": parts[0], "Land": state, "Schulform/Ordner": school, "Dokument": rel.stem}


def load_manual(rel: Path) -> str:
    f = MANUAL_DIR / rel.with_suffix(".json")
    if not f.exists():
        return "(keine manuelle Codierung)"
    try:
        data = json.loads(f.read_text(encoding="utf-8") or "[]")
    except json.JSONDecodeError:
        return "(Datei leer/ungültig)"
    if not data:
        return "–"
    return "; ".join(
        f"{a.get('code')} [{a.get('binding')}, Kl. {a.get('grade')}"
        + (", eA" if a.get("higer_standards") else "")
        + (", Wahl" if a.get("elective") else "")
        + "]"
        for a in data
    )


def load_include() -> dict[str, bool]:
    if not INCLUDE_TABLE.exists():
        return {}
    df = pd.read_csv(INCLUDE_TABLE)
    return {str(Path(r).with_suffix("")).replace("\\", "/"): bool(i) for r, i in zip(df.rel_path, df.include)}


# ------------------------------------------------------------------
def passage_id(doc_key: str, page_no: int, text: str) -> str:
    """Stable ID from content, so codings survive re-runs with changed keywords."""
    h = hashlib.sha1(f"{doc_key}|{page_no}|{text[:120]}".encode("utf-8")).hexdigest()[:8]
    return f"P-{h}"


def build():
    include = load_include()
    rows, docs = [], []
    json_docs, json_passages = {}, []
    for txt in sorted(TXT_DIR.rglob("*.txt")):
        rel = txt.relative_to(TXT_DIR)
        meta = doc_meta(rel)
        key = str(rel.with_suffix("")).replace("\\", "/")
        pdf = PDF_DIR / rel.with_suffix(".pdf")
        raw = txt.read_text(encoding="utf-8", errors="replace")
        pages = parse_pages(raw)
        n_pass = 0
        page_texts = {}
        for page_no, page_raw in pages:
            page = normalize(page_raw)
            for p in find_passages(page):
                n_pass += 1
                attr = hints(ATTR_RE, p["text"])
                specific = set(p["keywords"]) - GENERIC_ANCHORS
                core_attr = {k for k in PRIORITY_ATTRS if ATTR_RE[k].search(p["core"])}
                pid = passage_id(key, page_no, p["text"])
                page_texts[page_no] = page
                json_passages.append({
                    "id": pid, "doc": key, "page": page_no,
                    "priority": "hoch" if specific or core_attr else "niedrig",
                    "keywords": p["keywords"], "text": p["text"],
                    "hint_binding": hints(BIND_RE, p["text"]), "hint_attributes": attr,
                    "context_cues": context_cues(page),
                })
                rows.append({
                    "ID": pid,
                    **meta,
                    "Im Korpus": include.get(key, "n/a"),
                    "Seite": page_no,
                    "Priorität": "hoch" if specific or core_attr else "niedrig",
                    "Treffer": ", ".join(p["keywords"]),
                    "Passage (aus Textextrakt)": p["text"][:1500],
                    "Hinweis Verbindlichkeit (automatisch)": hints(BIND_RE, p["text"]),
                    "Hinweis Attribute (automatisch)": attr,
                    "Kontext-Hinweise auf der Seite (automatisch)": context_cues(page),
                    # columns for manual verification
                    "Relevant?": "",
                    "Aussage Nr.": "",
                    "Code(s)": "",
                    "Stufe": "",
                    "Klassen/Phase (wörtlich)": "",
                    "Kursart": "",
                    "Anforderungsniveau": "",
                    "Verbindlichkeit": "",
                    "Markierung im Dokument (wörtlich)": "",
                    "Begründung im Dokument": "",
                    "Zitat geprüft (PDF)": "",
                    "Notiz": "",
                    "Zweitprüfung": "",
                })
        docs.append({
            **meta,
            "Pfad": key,
            "PDF vorhanden": pdf.exists(),
            "Im Korpus (include_table)": include.get(key, "n/a"),
            "Seiten": len(pages),
            "Kandidatenpassagen": n_pass,
            "Manuelle Codierung (bisher)": load_manual(rel),
            "Dokumentstatus (verbindlich/Entwurf/veraltet)": "",
            "Gültig für Stufe/Schulform": "",
            "Vollständig gelesen?": "",
            "Notiz": "",
        })
        json_docs[key] = {
            **meta, "key": key, "in_corpus": include.get(key, None),
            "pdf": str(rel.with_suffix(".pdf")).replace("\\", "/") if pdf.exists() else None,
            "n_pages": len(pages), "n_passages": n_pass,
            "manual_coding_v1": load_manual(rel),
            "pages": {str(k): v for k, v in page_texts.items()},
        }
    ids = [r["ID"] for r in rows]
    dupes = {i for i in ids if ids.count(i) > 1}
    if dupes:
        raise SystemExit(f"duplicate passage IDs: {sorted(dupes)[:5]}")
    write_workbook(pd.DataFrame(rows), pd.DataFrame(docs))
    OUT_JSON.write_text(json.dumps({"documents": json_docs, "passages": json_passages}, ensure_ascii=False),
                        encoding="utf-8")
    print(f"{len(docs)} documents, {len(rows)} candidate passages -> {OUT_FILE.name}, {OUT_JSON.name}")


# ------------------------------------------------------------------
INSTRUCTIONS = [
    ("Zweck", "Entwurf einer Extraktionstabelle für die qualitative Dokumentenanalyse der Curricula (RQ1a). "
              "Die Passagen wurden automatisch per Schlagwortsuche gefunden (Retrieval). Alle Kontextangaben "
              "werden manuell anhand der PDFs ergänzt."),
    ("Vorgehen", "1) Für jede Passage im PDF (Seite siehe Spalte 'Seite') prüfen, ob sie eine Vorgabe/Empfehlung zu "
                 "Programmierwerkzeugen enthält ('Relevant?'). 2) Falls ja: Code(s), Stufe, Kursart und Verbindlichkeit "
                 "eintragen und die Markierung, aus der die Verbindlichkeit folgt, wörtlich notieren. 3) Zitat gegen das "
                 "PDF prüfen (der Textextrakt kann Fehler enthalten, z. B. rückwärts gelesene Randbeschriftungen). "
                 "4) Auf dem Blatt 'Dokumente' Status und Geltungsbereich jedes Dokuments festhalten. "
                 "5) Stichprobe von Dokumenten vollständig lesen, um übersehene Passagen zu finden (Recall-Prüfung); "
                 "fehlende Passagen als neue Zeilen ergänzen (ID mit 'M' beginnen). "
                 "Alternativ (empfohlen) im Codier-Viewer 04_coding_viewer/viewer.html codieren."),
    ("Seite", "Seitenzahl im PDF-Viewer (Index der PDF-Datei), nicht die im Dokument gedruckte Seitenzahl."),
    ("Priorität", "hoch = Passage enthält ein konkretes Werkzeug, einen Werkzeugtyp (block-/textbasiert, visuell, Skriptsprache) "
                  "oder im Treffersatz selbst ein Attributwort (geeignet, einfach, altersangemessen, ...); niedrig = nur ein allgemeiner Anker wie 'Programmiersprache'. "
                  "Empfehlung: zuerst alle 'hoch'-Zeilen prüfen, 'niedrig'-Zeilen danach zügig durchsehen (meist 'nein'). "
                  "Auch 'niedrig'-Zeilen können relevant sein (z. B. 'eine geeignete Programmiersprache' ohne Attributwort im Fenster)."),
    ("Mehrspaltige Tabellen", "Tabellen werden im Textextrakt zeilenweise ineinander geschoben (Spalten wiederholt/vermischt). "
                              "Stufe und Verbindlichkeit in solchen Fällen immer aus dem PDF bestimmen."),
    ("Spalten mit (automatisch)", "Nur Lesehilfen! Sie zeigen, ob bestimmte Signalwörter in der Passage bzw. auf der Seite "
                                  "vorkommen. Sie sind keine Codierung und dürfen nicht ungeprüft übernommen werden."),
    ("Relevant?", "ja = Vorgabe/Empfehlung/Beispiel zu Programmierwerkzeugen (Sprachen, Umgebungen); "
                  "nein = falsch positiver Treffer (z. B. Programmiersprache als Unterrichtsgegenstand ohne Werkzeugbezug, "
                  "Logo als Bildmarke); unklar = in der Diskussion klären."),
    ("Aussagen", "Codiereinheit ist die AUSSAGE, nicht die Passage. Enthält eine Passage mehrere Vorgaben, die sich in Stufe, "
                 "Kursart, Anforderungsniveau oder Verbindlichkeit unterscheiden (z. B. grundlegend: grafisch; erhöht: textuell), "
                 "wird die Zeile kopiert und je Vorgabe eine Zeile codiert ('Aussage Nr.' 1, 2, …; gleiche ID). "
                 "NICHT aufteilen, wenn nur mehrere Werkzeuge unter denselben Bedingungen genannt werden."),
    ("Code(s)", "Mehrfachnennung mit ';'. Codebuch: blockbasiert | visuell | textbasiert | geeignet/didaktisch reduziert | "
                "einfach | altersangemessen | aktuell | Mittel zum Zweck | Inhalt als Proxy (OOP/Paradigma) | "
                "freie/Open-Source-Software | Übergang/Progression (Aussage ist Teil einer Abfolge, z. B. 'zunächst grafisch, "
                "später textuell' – beide Aussagen erhalten den Code) | konkretes Werkzeug: <Name>. "
                "Neue Codes nur nach Rücksprache und im Codebuch ergänzen."),
    ("Stufe", "Sek I | Sek II | Sek I+II | unklar (grob halten; Details in 'Klassen/Phase (wörtlich)')."),
    ("Kursart", "regulär | Wahlpflicht/Wahlfach | Leistungskurs | alle | unklar (Art des Kurses)."),
    ("Anforderungsniveau", "grundlegend | erhöht | nicht differenziert (Niveau innerhalb eines Kurses bzw. Standards, "
                           "z. B. grundlegendes/erhöhtes Anforderungsniveau der Bildungsstandards)."),
    ("Verbindlichkeit", "verbindlich | Empfehlung | Beispiel | nicht spezifiziert. Nur aus Markierungen des Dokuments selbst "
                        "ableiten (z. B. 'verbindlich', Abschnitt 'Hinweise', 'z. B.'); sonst 'nicht spezifiziert'."),
    ("Zweitprüfung", "Für die zweite Person: ok | Änderung (in 'Notiz' begründen). Unstimmigkeiten werden im Konsens geklärt."),
]

DROPDOWNS = {
    "Relevant?": "ja,nein,unklar",
    "Stufe": "Sek I,Sek II,Sek I+II,unklar",
    "Kursart": "regulär,Wahlpflicht/Wahlfach,Leistungskurs,alle,unklar",
    "Anforderungsniveau": "grundlegend,erhöht,nicht differenziert",
    "Verbindlichkeit": "verbindlich,Empfehlung,Beispiel,nicht spezifiziert",
    "Zitat geprüft (PDF)": "ja,korrigiert,nein",
    "Zweitprüfung": "ok,Änderung",
}
WIDTHS = {
    "ID": 8, "Gruppe": 10, "Land": 16, "Schulform/Ordner": 18, "Dokument": 24, "Im Korpus": 9, "Seite": 7, "Priorität": 9, "Treffer": 18,
    "Passage (aus Textextrakt)": 80, "Hinweis Verbindlichkeit (automatisch)": 16,
    "Hinweis Attribute (automatisch)": 18, "Kontext-Hinweise auf der Seite (automatisch)": 30,
}
AUTO_FILL = PatternFill("solid", fgColor="EDEDED")
MANUAL_FILL = PatternFill("solid", fgColor="FFF2CC")
HEADER_FONT = Font(bold=True)


def write_sheet(ws, df: pd.DataFrame, manual_from: str | None):
    ws.append(list(df.columns))
    for row in df.itertuples(index=False):
        ws.append(list(row))
    manual_start = list(df.columns).index(manual_from) + 1 if manual_from else None
    for j, col in enumerate(df.columns, 1):
        letter = get_column_letter(j)
        ws.column_dimensions[letter].width = WIDTHS.get(col, 16)
        head = ws.cell(row=1, column=j)
        head.font = HEADER_FONT
        head.alignment = Alignment(wrap_text=True, vertical="top")
        is_manual = manual_start and j >= manual_start
        head.fill = MANUAL_FILL if is_manual else AUTO_FILL
        for i in range(2, ws.max_row + 1):
            ws.cell(row=i, column=j).alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = ws.cell(row=2, column=6)
    ws.auto_filter.ref = ws.dimensions


def write_workbook(passages: pd.DataFrame, docs: pd.DataFrame):
    wb = Workbook()
    ws = wb.active
    ws.title = "Anleitung"
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 120
    for k, v in INSTRUCTIONS:
        ws.append([k, v])
    for row in ws.iter_rows():
        row[0].font = HEADER_FONT
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.append([])
    ws.append(["Farben", "Grau = automatisch erzeugt, Gelb = manuell auszufüllen."])

    wp = wb.create_sheet("Passagen")
    write_sheet(wp, passages, "Relevant?")
    cols = list(passages.columns)
    for col, options in DROPDOWNS.items():
        letter = get_column_letter(cols.index(col) + 1)
        dv = DataValidation(type="list", formula1=f'"{options}"', allow_blank=True)
        dv.add(f"{letter}2:{letter}{len(passages) + 1}")
        wp.add_data_validation(dv)

    wd = wb.create_sheet("Dokumente")
    write_sheet(wd, docs, "Dokumentstatus (verbindlich/Entwurf/veraltet)")
    wd.freeze_panes = "E2"
    wb.save(OUT_FILE)


if __name__ == "__main__":
    build()
