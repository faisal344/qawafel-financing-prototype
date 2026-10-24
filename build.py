"""Inline data/sample-data.js into app.html -> index.html (one self-contained file you can double-click)."""
from pathlib import Path
root = Path(__file__).resolve().parent.parent
html = (root / "app.html").read_text()
data = (root / "data" / "sample-data.js").read_text()
(root / "index.html").write_text(html.replace("<!--DATA-->", "<script>\n" + data + "</script>"))
print("wrote index.html", round(len(html + data) / 1e6, 2), "MB")
