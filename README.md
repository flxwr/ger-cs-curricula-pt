# Programming Tools in German CS Curricula — Code and Data

Code, coding scheme, codings and results of a qualitative document analysis of
German computer-science curricula (state curricula, KMK and GI
recommendations): which programming tools (languages, environments) the
curricula require, recommend or give as examples, for which grades, and how
bindingly.

## What is (not) in this repository

| Included | Not included |
| --- | --- |
| Pipeline code (`Curricula/01_text_extraction` – `05_synthesis`) | The curriculum documents (PDFs) |
| Corpus register with inclusion decisions (`00_sources/corpus_include.csv`) | Full text extracted from them |
| Source list with URLs and checksums (`00_sources/sources.csv`) | |
| Coded passages, i.e. the retrieval windows (`03_extraction_table/passages.json`) | Raw keyword-hit context tables |
| Codings incl. verbatim quotes (`04_coding_viewer/codings/`) | |
| Browser-based coding tool (`04_coding_viewer/viewer.html`) | |
| Aggregate keyword statistics and sensitivity analysis (`02_keyword_retrieval/keyword_results`) | |
| Synthesis tables and figures (`05_synthesis/output`) | |

## Reproducing the results

```
pip install -r requirements.txt
cd Curricula/05_synthesis
jupyter nbconvert --to notebook --execute evaluation.ipynb
```

`evaluation.ipynb` needs only `03_extraction_table/passages.json`,
`04_coding_viewer/codings/coding_rater1.json` and
`00_sources/corpus_include.csv`. All analytic choices (corpus overrides,
grade ranges assumed for Sek I/II, tool categories) are set in its section 1.
The unit of aggregation is the state; GI/KMK recommendations are reported
separately.

To browse the codings, open `Curricula/04_coding_viewer/viewer.html` in a
browser and load `codings/coding_rater1.json`.

## Rebuilding the passages

Steps 01–03 need the curriculum documents as PDFs in
`Curricula/00_sources/pdfs/`, at the paths in the `file` column of
`00_sources/sources.csv`. That file gives a download URL and the SHA-256 of
the coded version for every document. The documents were collected in
late 2024 and early 2025, and several have been revised or removed since.
Column `verified` records what was found online in November 2026:

- `live`: the identical file (same SHA-256) is available at `url`.
- `wayback`: the identical file is available only as an Internet Archive snapshot.
- `mirror`: the identical file is available only on a third-party server.
- `equivalent`: the file differs, but its text is the same.
- `render`: a PDF generated from a web curriculum (Bavarian LehrplanPLUS,
  Baden-Württemberg Bildungspläne). Re-generating it gives the current web version.
- `changed`: only another version is available; `note` says which.
- `manual`: no direct file link; `url` points to where the document is listed.
- `no-pdf`: no separate file; the text was taken from the document named in `note`.
- `none`: nothing found.

Then run:

1. `01_text_extraction/analyse.ipynb` — PDF → text with page markers
2. `02_keyword_retrieval/keyword-analysis.ipynb` — keyword hits, context windows, sensitivity analysis
3. `python Curricula/03_extraction_table/build_extraction_table.py` — candidate passages
4. `python Curricula/04_coding_viewer/build_viewer_data.py` — data for the viewer

Passage ids (`P-xxxxxxxx`) are derived from document, page and passage text,
so codings stay attached to their passages as long as the text is unchanged.

## Coding scheme

The unit of coding is the *statement*. A passage is only the retrieval
window; if it contains requirements that differ in level, course type, level
of demand or bindingness, each becomes its own statement.

Per passage: relevant (ja / nein / unklar), note.
Per statement: codes (blockbasiert, visuell/grafisch, textbasiert,
geeignet/didaktisch reduziert, einfach, altersangemessen, aktuell/modern,
Mittel zum Zweck, Inhalt als Proxy, freie/Open-Source-Software,
Übergang/Progression, konkretes Werkzeug), named tools, Stufe (Sek I / Sek II),
grades, Kursart, Anforderungsniveau, Verbindlichkeit (only from the
document's own markers), reason, verbatim quote.

## License

- **Code** (`*.py`, `*.ipynb`, `viewer.html`): [MIT License](LICENSE).
- **Data and documentation** (corpus register, codings, results):
  [CC BY 4.0](https://creativecommons.org/licenses/by/4.0/), see
  [LICENSE-CC-BY-4.0](LICENSE-CC-BY-4.0).

Passage texts and quotes are excerpts from the curriculum documents and
remain subject to the terms of their publishers.
