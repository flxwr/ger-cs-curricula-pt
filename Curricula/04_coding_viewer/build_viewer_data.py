"""
Converts ../03_extraction_table/passages.json into viewer_data.js, which
viewer.html loads via a plain <script src="..."> tag. A <script> tag (unlike
fetch()) works under the file:// protocol, so viewer.html can be opened by
double-clicking it -- no local server needed.

Run after build_extraction_table.py:
    python Curricula/04_coding_viewer/build_viewer_data.py
"""
import json
from pathlib import Path

OUT_DIR = Path(__file__).resolve().parent
IN_JSON = OUT_DIR.parent / "03_extraction_table" / "passages.json"
OUT_JS = OUT_DIR / "viewer_data.js"


def main():
    data = json.loads(IN_JSON.read_text(encoding="utf-8"))
    payload = json.dumps(data, ensure_ascii=False, separators=(",", ":"))
    OUT_JS.write_text(f"const CORPUS = {payload};\n", encoding="utf-8")
    print(f"wrote {OUT_JS.name}: {len(data['documents'])} documents, "
          f"{len(data['passages'])} passages, {OUT_JS.stat().st_size / 1e6:.1f} MB")


if __name__ == "__main__":
    main()
