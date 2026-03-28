# Windows Business Automation Tool (URL Checker)

A standalone, installation-free Windows desktop application designed to perform bulk URL health checks and export the results to Excel instantly.

✔ Saves hours of manual clicking by validating thousands of URLs in seconds
✔ Requires absolutely zero technical knowledge to use—just double-click the .EXE
✔ Provides instant, actionable reports via clean Excel/CSV exports and color-coded UI

## Use Cases
- **SEO Auditing:** Allow non-technical marketing teams to instantly check thousands of backlinks for 404 errors.
- **IT Operations:** Quickly verify the uptime of internal web services and dashboards without writing custom scripts.
- **Data Cleaning:** Validate massive lists of web directories or leads before running expensive marketing campaigns.

---

## Tech Stack

- **Python + tkinter** — native cross-platform GUI
- **httpx** — async-capable HTTP client with redirect tracking
- **openpyxl** — Excel export
- **PyInstaller** — packages everything into a single .exe

---

## Run from Source

```bash
git clone https://github.com/bck-stack/python-desktop-tool
cd python-desktop-tool
pip install -r requirements.txt
python app.py
```

---

## Build .EXE

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name "URLHealthChecker" app.py
# Output: dist/URLHealthChecker.exe
```

---

## Usage

1. Paste URLs into the text box (one per line)
2. Or click **Load from file** to import a .txt list
3. Click **Check URLs**
4. Review color-coded results
5. Export with **Export Excel** or **Export CSV**

---

## Example Output

| URL | Status | Time (ms) | Final URL |
|-----|--------|-----------|-----------|
| https://example.com | 200 | 142.3 | https://example.com/ |
| https://httpbin.org/status/404 | 404 | 310.1 | https://httpbin.org/status/404 |
| https://broken-site.xyz | CONN_ERR | -1 | — |

---

## Screenshot

![Preview](screenshots/preview.png)

---

## License

MIT
