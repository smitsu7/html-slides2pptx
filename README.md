# HTML to Editable PPTX Converter

`html_ppt.py` converts slide-oriented HTML documents into editable PowerPoint (`.pptx`) decks.  
The focus is on recreating text, shapes, tables, charts, and images as native PowerPoint objects, keeping fonts, colors, and layout as close as possible to the source HTML/CSS.

## Features
- Parses `.slide`, `.slide-container`, or `[data-slide]` sections as individual slides.
- Converts headings, paragraphs, lists, and inline styling (color, bold, italics) into editable text boxes.
- Recreates HTML tables with cell-level styles (backgrounds, borders, alignment, fonts).
- Carries HTML table cell padding into PowerPoint cell margins for denser table layouts.
- Imports images with correct sizing and relative paths.
- Converts CSS-based shapes, gradients, and simple SVG graphics (polylines, circles, ellipses) into PowerPoint shapes.
- Supports CSS `conic-gradient` pie charts by approximating each slice with a PowerPoint freeform wedge.
- Automatically scales each HTML slide to fit the selected PPT canvas (`16:9` or A4 landscape) without distorting aspect ratios.
- Can auto-fallback complex slides to page images in `browser` mode for higher visual fidelity.

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
    --dpr 2 \
    --page-size A4 \
    --image-map IMAGE_URL_1=./assets/cover.png
```

Arguments:
- `input.html` – path to the slide-like HTML file.
- `output.pptx` – destination PowerPoint file.
- `--engine` – `browser` (default) uses Playwright for full CSS layout; `static` uses BeautifulSoup (fallback, less accurate).
- `--selector` – CSS selector to identify slide sections; defaults to `.slide, .slide-container, [data-slide]`.
- `--viewport` – browser viewport size when using the `browser` engine.
- `--dpr` – device pixel ratio for Playwright screenshots (affects image/render quality).
- `--page-size` – output slide size; `16:9` (default) or `A4` (landscape).
- `--rasterize-slides` – `none` (default), `auto`, or a 1-based list/range such as `3-6,8`; use only when you explicitly prefer visual fidelity over editability for selected slides.
- `--image-map` – replace placeholder image URLs such as `IMAGE_URL_1` with real local files. Repeatable.

Example with multiple placeholder mappings:

```bash
python html_ppt.py input.html output.pptx \
    --image-map IMAGE_URL_1=./assets/hero.png \
    --image-map IMAGE_URL_2=~/Downloads/chart.jpg
```

If `--engine browser` is unavailable (e.g., Playwright not installed), run with `--engine static` for a best-effort conversion that does not require a headless browser.

## Notes
- Remote images referenced via HTTP/HTTPS are not downloaded; place assets locally or ensure paths are accessible.
- Placeholder strings like `IMAGE_URL_1` can be mapped explicitly with `--image-map`; relative mapped paths are resolved from the input HTML directory.
- For best fidelity, keep each slide within a fixed-size container (e.g., `1280x720`).
- The converter prioritizes editable content by default. If needed, `--rasterize-slides` can be used explicitly for selected slides as a fidelity fallback.

## Troubleshooting
- If you see Playwright errors, confirm `playwright install chromium` has been run and that the `playwright` Python package is installed.
- Missing fonts may change layout in PowerPoint; install the same fonts locally for accurate rendering.
- For debugging, rerun with `--engine static` to rule out browser-specific issues.

## License
This project is provided as-is; adapt it to your workflow as needed.
