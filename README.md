# HTML to Editable PPTX Converter

`html_ppt.py` converts slide-oriented HTML documents into editable PowerPoint (`.pptx`) decks.  
The focus is on recreating text, shapes, tables, charts, and images as native PowerPoint objects, keeping fonts, colors, and layout as close as possible to the source HTML/CSS.

## Features
- Parses `.slide`, `.slide-container`, or `[data-slide]` sections as individual slides.
- Converts headings, paragraphs, lists, and inline styling (color, bold, italics) into editable text boxes.
- Recreates HTML tables with cell-level styles (backgrounds, borders, alignment, fonts).
- Imports images with correct sizing and relative paths.
- Converts CSS-based shapes, gradients, and simple SVG graphics (polylines, circles, ellipses) into PowerPoint shapes.
- Supports CSS `conic-gradient` pie charts by approximating each slice with a PowerPoint freeform wedge.
- Automatically scales each HTML slide to fit the 16:9 PPT canvas without distorting aspect ratios.

## Requirements
- Python 3.9+
- [python-pptx](https://python-pptx.readthedocs.io/)
- [beautifulsoup4](https://www.crummy.com/software/BeautifulSoup/bs4/doc/)
- [pillow](https://pillow.readthedocs.io/)
- [playwright](https://playwright.dev/python/) (for the browser engine)

Install dependencies using `requirements.txt`:

```bash
pip install -r requirements.txt
playwright install chromium
```

## Usage

```bash
python html_ppt.py input.html output.pptx \
    --engine browser \
    --selector ".slide, .slide-container, [data-slide]" \
    --viewport 1920x1080 \
    --dpr 2
```

Arguments:
- `input.html` – path to the slide-like HTML file.
- `output.pptx` – destination PowerPoint file.
- `--engine` – `browser` (default) uses Playwright for full CSS layout; `static` uses BeautifulSoup (fallback, less accurate).
- `--selector` – CSS selector to identify slide sections; defaults to `.slide, .slide-container, [data-slide]`.
- `--viewport` – browser viewport size when using the `browser` engine.
- `--dpr` – device pixel ratio for Playwright screenshots (affects image/render quality).

If `--engine browser` is unavailable (e.g., Playwright not installed), run with `--engine static` for a best-effort conversion that does not require a headless browser.

## Notes
- Remote images referenced via HTTP/HTTPS are not downloaded; place assets locally or ensure paths are accessible.
- For best fidelity, keep each slide within a fixed-size container (e.g., `1280x720`).
- The converter prioritizes editable content; highly customized CSS layouts may need manual post-processing in PowerPoint.

## Troubleshooting
- If you see Playwright errors, confirm `playwright install chromium` has been run and that the `playwright` Python package is installed.
- Missing fonts may change layout in PowerPoint; install the same fonts locally for accurate rendering.
- For debugging, rerun with `--engine static` to rule out browser-specific issues.

## License
This project is provided as-is; adapt it to your workflow as needed.
