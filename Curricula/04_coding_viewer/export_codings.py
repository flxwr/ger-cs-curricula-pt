"""
Collects the coding exports from viewer.html (codings/coding_*.json) and writes
codings_export.xlsx:
  - one sheet per coder: one row per STATEMENT (passage ID + statement number),
    joined with document metadata
  - 'Dokumente': document-level fields and defaults per coder
  - for >= 2 coders (first two are compared):
      'Übereinstimmung': agreement figures
      'Vergleich_Passagen': step 1 -- relevance and number of statements per passage
      'Vergleich_Aussagen': step 2 -- statements paired by signature, for passages
                            where both coders have the same number of statements

Older exports (schema v1: one coding per passage) are converted on the fly the
same way viewer.html does it.

Usage:  python Curricula/04_coding_viewer/export_codings.py [coding_A.json coding_B.json]
(without arguments, all codings/coding_*.json are used)
"""
import json
import sys
from itertools import permutations
from pathlib import Path

import pandas as pd

HERE = Path(__file__).resolve().parent
CORPUS = json.loads((HERE.parent / "03_extraction_table" / "passages.json").read_text(encoding="utf-8"))
DOCS = CORPUS["documents"]
PASSAGES = {p["id"]: p for p in CORPUS["passages"]}
OUT = HERE / "codings_export.xlsx"

DEFAULTABLE = ["level", "grades", "course", "niveau", "binding", "marker"]
SIGNATURE_FIELDS = ["level", "course", "niveau", "binding"]
V1_FIELDS = ["codes", "tools", "level", "grades", "course", "binding", "marker", "reason", "quote", "quote_checked", "own"]


# ---------------------------------------------------------------------------
def migrate(r: dict) -> dict:
    """v1 passage record -> v2 with statements (mirrors migrateRecord in viewer.html)."""
    if "statements" in r:
        return r
    s = {f: r.get(f) for f in V1_FIELDS if f in r}
    s.setdefault("codes", {})
    s.setdefault("own", {f: bool(s.get(f)) for f in DEFAULTABLE})
    if s.get("course") == "erhöhtes Niveau/LK":
        s["course"], s["niveau"] = "Leistungskurs", "erhöht"
        s["_migrated"] = "Kursart migriert"
    has_content = any(v for k, v in s.items() if k not in ("codes", "own")) or any(s["codes"].values())
    r["statements"] = [s] if (has_content or r.get("relevant") == "ja") else []
    return r


def load(path: Path) -> dict:
    d = json.loads(path.read_text(encoding="utf-8-sig"))  # -sig: tolerate a BOM from editors
    d.setdefault("coder", path.stem.replace("coding_", ""))
    for r in d.get("passages", {}).values():
        migrate(r)
    return d


def codes_of(s: dict) -> list[str]:
    return sorted(k for k, v in (s.get("codes") or {}).items() if v)


def signature(s: dict) -> str:
    return "+".join(c for c in codes_of(s) if c != "werkzeug") + " | " + " / ".join(s.get(f) or "–" for f in SIGNATURE_FIELDS)


def in_corpus(c: dict, doc: str) -> bool:
    """Coder's decision in the viewer overrides 00_sources/corpus_include.csv (mirrors inCorpus in viewer.html)."""
    decision = (c.get("documents", {}).get(doc) or {}).get("in_corpus", "")
    if decision in ("ja", "nein"):
        return decision == "ja"
    return DOCS.get(doc, {}).get("in_corpus") is not False


def passage_meta(c: dict, pid: str, r: dict) -> dict:
    p = PASSAGES.get(pid) or {**c.get("added", {}).get(pid, {}), "id": pid}
    doc = p.get("doc") or r.get("_doc")
    m = DOCS.get(doc, {})
    return {"ID": pid, "Land": m.get("Land"), "Dokument": m.get("Dokument"), "Schulform/Ordner": m.get("Schulform/Ordner"),
            "Im Korpus": in_corpus(c, doc), "Seite": p.get("page") or r.get("_page"), "ergänzt": pid.startswith("M-")}


# ---------------------------------------------------------------------------
def statement_rows(c: dict) -> pd.DataFrame:
    rows = []
    for pid, r in c.get("passages", {}).items():
        if not r.get("relevant"):
            continue
        meta = passage_meta(c, pid, r)
        ss = r["statements"] if r["relevant"] == "ja" else []
        if not ss:  # nein / unklar (or ja without statement): one row for the passage
            rows.append({**meta, "Aussage": None, "Anzahl Aussagen": 0, "Relevant": r["relevant"],
                         "Notiz (Passage)": r.get("note")})
            continue
        for i, s in enumerate(ss, 1):
            own = s.get("own") or {}
            rows.append({
                **meta, "Aussage": i, "Anzahl Aussagen": len(ss), "Relevant": r["relevant"],
                "Codes": "; ".join(codes_of(s)), **{f"code_{k}": bool(v) for k, v in (s.get("codes") or {}).items()},
                "Werkzeuge": s.get("tools"), "Stufe": s.get("level"), "Klassen/Phase": s.get("grades"),
                "Kursart": s.get("course"), "Anforderungsniveau": s.get("niveau"), "Verbindlichkeit": s.get("binding"),
                "Markierung": s.get("marker"), "Begründung": s.get("reason"), "Zitat": s.get("quote"),
                "Zitat geprüft": s.get("quote_checked"),
                "aus Dokument-Standard": "; ".join(f for f in DEFAULTABLE if s.get(f) and not own.get(f, True)),
                "migriert/prüfen": s.get("_migrated", ""), "Notiz (Passage)": r.get("note"), "geändert": r.get("updated_at"),
            })
    df = pd.DataFrame(rows)
    if not df.empty:
        df["_rank"] = df["Relevant"].map({"ja": 0, "unklar": 1, "nein": 2})
        df = df.sort_values(["_rank", "Land", "Dokument", "Seite", "Aussage"]).drop(columns="_rank")
    return df


def doc_rows(coders: list[dict]) -> pd.DataFrame:
    rows = []
    for c in coders:
        for key, r in c.get("documents", {}).items():
            defaults = r.get("defaults") or {}
            if any(r.get(k) for k in ("status", "scope", "read_fully", "note", "in_corpus", "corpus_reason")) or any(defaults.values()):
                m = DOCS.get(key, {})
                rows.append({"Codierer:in": c["coder"], "Land": m.get("Land"), "Dokument": m.get("Dokument"),
                             "Im Korpus (CSV)": m.get("in_corpus") is not False,
                             "Im Korpus (Entscheidung)": r.get("in_corpus") or "wie CSV",
                             "Im Korpus (wirksam)": in_corpus(c, key), "Grund Ein-/Ausschluss": r.get("corpus_reason"),
                             "Status": r.get("status"), "Gültig für": r.get("scope"),
                             "Vollständig gelesen": r.get("read_fully"), "Notiz": r.get("note"),
                             **{f"Standard {k}": v for k, v in defaults.items()}})
    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
def kappa(a: pd.Series, b: pd.Series):
    a, b = a.fillna("").astype(str), b.fillna("").astype(str)
    if len(a) == 0:
        return None
    po = (a == b).mean()
    pe = sum((a == k).mean() * (b == k).mean() for k in set(a) | set(b))
    return None if pe == 1 else round((po - pe) / (1 - pe), 3)


def pair_statements(sa: list[dict], sb: list[dict]) -> list[tuple[dict, dict]]:
    """Pair statements of equal-length lists so that the number of matching fields is maximal."""
    def score(x, y):
        return sum((x.get(f) or "") == (y.get(f) or "") for f in SIGNATURE_FIELDS) + len(set(codes_of(x)) & set(codes_of(y)))
    if len(sa) > 6:  # avoid factorial blow-up; keep given order
        return list(zip(sa, sb))
    best = max(permutations(range(len(sb))), key=lambda perm: sum(score(sa[i], sb[j]) for i, j in enumerate(perm)))
    return [(sa[i], sb[j]) for i, j in enumerate(best)]


def comparison(c1: dict, c2: dict):
    n1, n2 = c1["coder"], c2["coder"]
    ids = sorted(set(c1["passages"]) | set(c2["passages"]))
    prow, srow, paired = [], [], []
    for pid in ids:
        a, b = c1["passages"].get(pid, {}), c2["passages"].get(pid, {})
        ra, rb = a.get("relevant", ""), b.get("relevant", "")
        if not (ra or rb):
            continue
        sa = a.get("statements", []) if ra == "ja" else []
        sb = b.get("statements", []) if rb == "ja" else []
        meta = passage_meta(c1, pid, a)
        same_n = ra == rb == "ja" and len(sa) == len(sb)
        sig_equal = same_n and sorted(map(signature, sa)) == sorted(map(signature, sb))
        prow.append({**{k: meta[k] for k in ("ID", "Land", "Dokument", "Seite")},
                     f"relevant ({n1})": ra, f"relevant ({n2})": rb,
                     f"Aussagen ({n1})": len(sa), f"Aussagen ({n2})": len(sb),
                     f"Signaturen ({n1})": " || ".join(map(signature, sa)), f"Signaturen ({n2})": " || ".join(map(signature, sb)),
                     "Abweichung": not (ra == rb and (ra != "ja" or sig_equal))})
        if same_n:
            for i, (x, y) in enumerate(pair_statements(sa, sb), 1):
                paired.append((x, y))
                row = {"ID": pid, "Dokument": meta["Dokument"], "Seite": meta["Seite"], "Paar": i}
                for f in SIGNATURE_FIELDS:
                    row[f"{f} ({n1})"], row[f"{f} ({n2})"] = x.get(f, ""), y.get(f, "")
                row[f"codes ({n1})"], row[f"codes ({n2})"] = "; ".join(codes_of(x)), "; ".join(codes_of(y))
                row["Abweichung"] = signature(x) != signature(y)
                srow.append(row)

    rel_both = [(c1["passages"][i].get("relevant", ""), c2["passages"][i].get("relevant", ""))
                for i in ids if c1["passages"].get(i, {}).get("relevant") and c2["passages"].get(i, {}).get("relevant")]
    yes_both = [p for p in prow if p[f"relevant ({n1})"] == p[f"relevant ({n2})"] == "ja"]
    ser = lambda xs, k: pd.Series([x[k] for x in xs])
    stats = [
        {"Schritt": "1", "Merkmal": "relevant (beide codiert)", "n": len(rel_both),
         "Übereinstimmung": round(sum(a == b for a, b in rel_both) / len(rel_both), 3) if rel_both else None,
         "Cohens kappa": kappa(pd.Series([a for a, _ in rel_both]), pd.Series([b for _, b in rel_both]))},
        {"Schritt": "1", "Merkmal": "gleiche Anzahl Aussagen (beide 'ja')", "n": len(yes_both),
         "Übereinstimmung": round(sum(p[f"Aussagen ({n1})"] == p[f"Aussagen ({n2})"] for p in yes_both) / len(yes_both), 3) if yes_both else None,
         "Cohens kappa": None},
    ]
    for f in SIGNATURE_FIELDS:
        a, b = pd.Series([x.get(f, "") for x, _ in paired]), pd.Series([y.get(f, "") for _, y in paired])
        stats.append({"Schritt": "2", "Merkmal": f"{f} (gepaarte Aussagen)", "n": len(paired),
                      "Übereinstimmung": round((a == b).mean(), 3) if paired else None, "Cohens kappa": kappa(a, b)})
    all_codes = sorted({k for x, y in paired for k in list((x.get("codes") or {})) + list((y.get("codes") or {}))})
    for k in all_codes:
        a = pd.Series([bool((x.get("codes") or {}).get(k)) for x, _ in paired])
        b = pd.Series([bool((y.get("codes") or {}).get(k)) for _, y in paired])
        stats.append({"Schritt": "2", "Merkmal": f"Code {k} (gepaarte Aussagen)", "n": len(paired),
                      "Übereinstimmung": round((a == b).mean(), 3) if paired else None, "Cohens kappa": kappa(a, b)})
    return pd.DataFrame(stats), pd.DataFrame(prow), pd.DataFrame(srow)


def main():
    files = [Path(a) for a in sys.argv[1:]] or sorted((HERE / "codings").glob("coding_*.json"))
    if not files:
        sys.exit("No coding files found (export from viewer.html into codings/).")
    coders = [load(f) for f in files]
    with pd.ExcelWriter(OUT) as xw:
        for c in coders:
            statement_rows(c).to_excel(xw, sheet_name=str(c["coder"])[:31], index=False)
        doc_rows(coders).to_excel(xw, sheet_name="Dokumente", index=False)
        if len(coders) >= 2:
            stats, prow, srow = comparison(coders[0], coders[1])
            stats.to_excel(xw, sheet_name="Übereinstimmung", index=False)
            prow.to_excel(xw, sheet_name="Vergleich_Passagen", index=False)
            srow.to_excel(xw, sheet_name="Vergleich_Aussagen", index=False)
    print(f"{len(coders)} coding file(s) -> {OUT.name}")


if __name__ == "__main__":
    main()
