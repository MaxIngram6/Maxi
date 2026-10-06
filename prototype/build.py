"""Inline data.js into app.html to make one file that opens anywhere (phone, email attachment, Teams)."""
import pathlib
here = pathlib.Path(__file__).parent
html = (here / "app.html").read_text()
data = (here / "data.js").read_text()
tag = '<script src="data.js"></script>'
assert tag in html
out = here / "Eurogold-Brickwork-Payments.html"
out.write_text(html.replace(tag, "<script>\n" + data + "\n</script>"))
print(out, out.stat().st_size // 1024, "KB")
