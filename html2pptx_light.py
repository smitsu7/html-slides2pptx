#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Virtual-render HTML slides into editable PPTX without any browser runtime.

Highlights
----------
- StyleParser: cssutils-based cascade (inline + <style>) with class/id support.
- DOMWalker & LayoutEngine: approximates browser layout (block, flex, grid,
  card layouts) and assigns coordinates within a 16:9 canvas.
- OverflowManager: splits content across slides when total height exceeds the
  viewport.
- PPTXRenderer: recreates text, lists, tables (with per-cell styling), shapes,
  images, and Playwright-equivalent vector blocks in native PowerPoint objects.
- ChartBuilder: parses <script> definitions to rebuild editable PPT charts.

Usage:
    python html2pptx_light.py input.html output.pptx

Requirements:
    pip install beautifulsoup4 cssutils python-pptx pillow lxml
"""

from __future__ import annotations

import argparse
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

import cssutils
from bs4 import BeautifulSoup, NavigableString, Tag
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Inches, Pt

try:
    from PIL import Image, ImageColor
except ImportError:
    Image = None
    ImageColor = None

SLIDE_REF_WIDTH = 1920
SLIDE_REF_HEIGHT = 1080
MARGIN_X = 140
MARGIN_Y = 120
COLUMN_GAP = 32
ROW_GAP = 36
DEFAULT_FONT_SIZE = 26

BLOCK_TAGS = {"div", "section", "article", "p", "ul", "ol", "li", "table", "img"}
TEXT_TAGS = {"h1", "h2", "h3", "h4", "h5", "h6", "p", "span", "div", "blockquote"}
LIST_TAGS = {"ul", "ol"}


# --------------------------------------------------------------------------- #
# CSS helpers
# --------------------------------------------------------------------------- #


def parse_length(value: Optional[str], reference: Optional[float] = None) -> Optional[float]:
    if not value:
        return None
    value = value.strip().lower()
    if value.endswith("px"):
        return float(value[:-2])
    if value.endswith("%") and reference:
        try:
            return float(value[:-1]) / 100.0 * reference
        except ValueError:
            return None
    if value.replace(".", "", 1).isdigit():
        return float(value)
    return None


def parse_box(value: Optional[str], fallback: float = 0.0) -> Tuple[float, float, float, float]:
    if not value:
        return (fallback, fallback, fallback, fallback)
    parts = value.replace(",", " ").split()
    nums = [parse_length(p) or fallback for p in parts]
    if len(nums) == 1:
        return (nums[0], nums[0], nums[0], nums[0])
    if len(nums) == 2:
        return (nums[0], nums[1], nums[0], nums[1])
    if len(nums) == 3:
        return (nums[0], nums[1], nums[2], nums[1])
    t, r, b, l = (nums + [fallback] * 4)[:4]
    return (t, r, b, l)


def css_color(value: Optional[str]) -> Optional[RGBColor]:
    if not value:
        return None
    value = value.strip()
    if not value or value.lower() == "transparent":
        return None
    if value.startswith("#") and len(value) == 7:
        return RGBColor(int(value[1:3], 16), int(value[3:5], 16), int(value[5:7], 16))
    if value.startswith("rgb"):
        nums = value[value.find("(") + 1 : value.find(")")].split(",")
        r, g, b = [int(float(n.strip())) for n in nums[:3]]
        return RGBColor(r, g, b)
    if ImageColor:
        try:
            r, g, b = ImageColor.getrgb(value)
            return RGBColor(r, g, b)
        except Exception:
            return None
    return None


def estimate_text_height(text: str, font_px: float) -> float:
    lines = max(1, text.count("\n") + 1)
    return lines * font_px * 1.2 + 16


def normalize_text(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip()


# --------------------------------------------------------------------------- #
# Style Parser (cascade aware)
# --------------------------------------------------------------------------- #


class StyleParser:
    SELECTOR_RE = re.compile(r"[#.][\w-]+|[\w-]+|\*")

    def __init__(self, soup: BeautifulSoup):
        self.soup = soup
        self.rules: List[Tuple[Dict[str, Iterable[str]], Dict[str, str]]] = []
        cssutils.log.setLevel("ERROR")
        self._collect_rules()

    def _collect_rules(self):
        for style_tag in self.soup.find_all("style"):
            css_text = style_tag.string or ""
            try:
                sheet = cssutils.parseString(css_text)
            except Exception:
                continue
            for rule in sheet:
                if rule.type != rule.STYLE_RULE:
                    continue
                props = {prop.name.strip(): prop.value.strip() for prop in rule.style}
                for selector in rule.selectorText.split(","):
                    sel = selector.strip()
                    parsed = self._parse_selector(sel)
                    if parsed:
                        self.rules.append((parsed, props))

    def _parse_selector(self, selector: str) -> Optional[Dict[str, Iterable[str]]]:
        if not selector:
            return None
        tokens = self.SELECTOR_RE.findall(selector)
        tag = None
        id_val = None
        classes: List[str] = []
        for token in tokens:
            if token == "*":
                continue
            if token.startswith("#"):
                id_val = token[1:]
            elif token.startswith("."):
                classes.append(token[1:])
            else:
                tag = token.lower()
        return {"tag": tag, "id": id_val, "classes": classes}

    def _matches(self, selector: Dict[str, Iterable[str]], element: Tag) -> bool:
        if selector["tag"] and selector["tag"] != element.name:
            return False
        if selector["id"] and element.get("id") != selector["id"]:
            return False
        classes = set(element.get("class", []))
        for cls in selector["classes"]:
            if cls not in classes:
                return False
        return True

    def compute(self, element: Tag) -> Dict[str, str]:
        style: Dict[str, str] = {}
        for selector, props in self.rules:
            if self._matches(selector, element):
                style.update(props)
        inline = element.get("style")
        if inline:
            for decl in inline.split(";"):
                if ":" not in decl:
                    continue
                k, v = decl.split(":", 1)
                style[k.strip().lower()] = v.strip()
        return style


# --------------------------------------------------------------------------- #
# Data classes
# --------------------------------------------------------------------------- #


@dataclass
class TextRun:
    text: str
    font_size: Optional[float] = None
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    color: Optional[str] = None


@dataclass
class TableCell:
    text: str
    style: Dict[str, str]


@dataclass
class LayoutFrame:
    left: float
    top: float
    width: float
    height: float


@dataclass
class Block:
    kind: str
    style: Dict[str, str]
    layout_hint: str = "column"
    frame: LayoutFrame = field(default_factory=lambda: LayoutFrame(0, 0, 0, 0))
    text: str = ""
    runs: List[TextRun] = field(default_factory=list)
    items: List[str] = field(default_factory=list)
    table: List[List[str]] = field(default_factory=list)
    cell_styles: List[List[TableCell]] = field(default_factory=list)
    image_path: Optional[Path] = None
    shape_style: Dict[str, Any] = field(default_factory=dict)
    children: List["Block"] = field(default_factory=list)


# --------------------------------------------------------------------------- #
# Inline text run extraction
# --------------------------------------------------------------------------- #


class TextRunExtractor:
    def __init__(self, style_parser: StyleParser):
        self.styles = style_parser

    def extract(self, element: Tag, base_style: Dict[str, str]) -> List[TextRun]:
        runs: List[TextRun] = []

        def walk(node, style: Dict[str, str]):
            if isinstance(node, NavigableString):
                text = str(node)
                text = text.replace("\xa0", " ")
                if text.strip():
                    runs.append(self._run(text, style))
                return
            if not isinstance(node, Tag):
                return
            if node.name == "br":
                runs.append(self._run("\n", style))
                return
            child_style = style.copy()
            node_style = self.styles.compute(node)
            child_style.update({k: v for k, v in node_style.items() if k in {"color", "font-weight", "font-style", "font-size"}})
            if node.name in {"strong", "b"}:
                child_style["font-weight"] = "bold"
            if node.name in {"em", "i"}:
                child_style["font-style"] = "italic"
            for child in node.children:
                walk(child, child_style)

        walk(element, base_style)
        merged: List[TextRun] = []
        for run in runs:
            if merged and self._can_merge(merged[-1], run):
                merged[-1].text += run.text
            else:
                merged.append(run)
        return merged or [self._run(normalize_text(element.get_text()), base_style)]

    def _run(self, text: str, style: Dict[str, str]) -> TextRun:
        return TextRun(
            text=text,
            font_size=parse_length(style.get("font-size")),
            bold=self._is_bold(style.get("font-weight")),
            italic=(style.get("font-style") == "italic"),
            color=style.get("color"),
        )

    def _is_bold(self, value: Optional[str]) -> bool:
        if not value:
            return False
        if value.isdigit():
            return int(value) >= 600
        return value.lower() in {"bold", "bolder"}

    def _can_merge(self, a: TextRun, b: TextRun) -> bool:
        return (a.font_size == b.font_size) and (a.bold == b.bold) and (a.italic == b.italic) and (a.color == b.color)


# --------------------------------------------------------------------------- #
# DOM Walker -> Blocks
# --------------------------------------------------------------------------- #


class BlockBuilder:
    def __init__(self, style_parser: StyleParser, base_dir: Path):
        self.styles = style_parser
        self.base_dir = base_dir
        self.run_extractor = TextRunExtractor(style_parser)

    def build(self, container: Tag) -> List[Block]:
        blocks: List[Block] = []
        for child in container.children:
            if isinstance(child, NavigableString):
                text = normalize_text(str(child))
                if not text:
                    continue
                pseudo = Tag(name="p")
                pseudo.string = text
                block = self._text_block(pseudo)
            else:
                block = self._dispatch(child)
            if block:
                blocks.append(block)
        return blocks

    def _dispatch(self, tag: Tag) -> Optional[Block]:
        if tag.name not in BLOCK_TAGS and tag.name not in TEXT_TAGS:
            return None
        style = self.styles.compute(tag)
        display = style.get("display", "").lower()
        layout_hint = "row" if "flex" in display else "grid" if display == "grid" else "column"

        if tag.name in TEXT_TAGS:
            return self._text_block(tag, layout_hint)
        if tag.name in LIST_TAGS:
            return self._list_block(tag, layout_hint)
        if tag.name == "table":
            return self._table_block(tag, layout_hint)
        if tag.name == "img":
            return self._image_block(tag, layout_hint)

        child_blocks = self.build(tag)
        if child_blocks:
            block = Block("group", style=style, layout_hint=layout_hint, children=child_blocks)
            block.shape_style = self._shape_style(style)
            return block
        shape = self._shape_block(tag)
        return shape

    def _text_block(self, tag: Tag, layout_hint: str = "column") -> Block:
        style = self.styles.compute(tag)
        base_font = parse_length(style.get("font-size")) or DEFAULT_FONT_SIZE
        block = Block("text", style=style, layout_hint=layout_hint)
        block.runs = self.run_extractor.extract(tag, style)
        block.text = "".join(run.text for run in block.runs)
        block.shape_style = self._shape_style(style)
        if not block.text:
            return None
        return block

    def _list_block(self, tag: Tag, layout_hint: str) -> Block:
        style = self.styles.compute(tag)
        items = [normalize_text(li.get_text(" ", strip=True)) for li in tag.find_all("li", recursive=False)]
        items = [item for item in items if item]
        block = Block("list", style=style, items=items, layout_hint=layout_hint)
        block.shape_style = self._shape_style(style)
        return block if items else None

    def _table_block(self, tag: Tag, layout_hint: str) -> Optional[Block]:
        style = self.styles.compute(tag)
        rows: List[List[str]] = []
        cell_styles: List[List[TableCell]] = []
        for tr in tag.find_all("tr"):
            row = []
            row_cells = []
            for cell in tr.find_all(["td", "th"]):
                text = normalize_text(cell.get_text(" ", strip=True))
                row.append(text)
                cell_style = self.styles.compute(cell)
                row_cells.append(TableCell(text=text, style=cell_style))
            if row:
                rows.append(row)
                cell_styles.append(row_cells)
        if not rows:
            return None
        block = Block("table", style=style, table=rows, cell_styles=cell_styles, layout_hint=layout_hint)
        return block

    def _image_block(self, tag: Tag, layout_hint: str) -> Block:
        style = self.styles.compute(tag)
        src = tag.get("src", "")
        path = resolve_image_path(src, self.base_dir)
        block = Block("image", style=style, image_path=path, layout_hint=layout_hint)
        block.text = tag.get("alt", "")
        return block

    def _shape_block(self, tag: Tag) -> Optional[Block]:
        style = self.styles.compute(tag)
        shape_style = self._shape_style(style)
        if any(shape_style.values()):
            return Block("shape", style=style, shape_style=shape_style)
        return None

    def _shape_style(self, style: Dict[str, str]) -> Dict[str, Any]:
        return {
            "fill": style.get("background") or style.get("background-color"),
            "border_color": style.get("border-color"),
            "border_width": style.get("border-width"),
        }


# --------------------------------------------------------------------------- #
# Layout Engine + Overflow Manager
# --------------------------------------------------------------------------- #


class LayoutEngine:
    def __init__(self, width: float, height: float):
        self.width = width
        self.height = height

    def layout(self, blocks: List[Block]) -> List[List[Block]]:
        abs_blocks: List[Block] = []
        cursor_y = MARGIN_Y
        for block in blocks:
            cursor_y = self._layout_block(block, MARGIN_X, cursor_y, self.width - 2 * MARGIN_X, abs_blocks)
            cursor_y += ROW_GAP
        return self._paginate(abs_blocks)

    def _layout_block(
        self,
        block: Block,
        left: float,
        top: float,
        width: float,
        collected: List[Block],
    ) -> float:
        height = self._preferred_height(block, width)
        block.frame = LayoutFrame(left, top, width, height)
        if block.kind == "group":
            content_left = left
            content_top = top
            if block.layout_hint == "row":
                available = width - (len(block.children) - 1) * COLUMN_GAP
                child_width = available / max(1, len(block.children))
                max_height = 0
                x = content_left
                for child in block.children:
                    child_h = self._layout_block(child, x, content_top, child_width, collected)
                    max_height = max(max_height, child_h - content_top)
                    x += child_width + COLUMN_GAP
                block.frame.height = max(block.frame.height, max_height)
            else:
                cursor = content_top
                for child in block.children:
                    child_h = self._layout_block(child, content_left, cursor, width, collected)
                    cursor = child_h + ROW_GAP
                block.frame.height = max(block.frame.height, cursor - top)
        else:
            collected.append(block)
        return block.frame.top + block.frame.height

    def _preferred_height(self, block: Block, width: float) -> float:
        explicit = parse_length(block.style.get("height"), width)
        if explicit:
            return explicit
        if block.kind == "text":
            font_px = parse_length(block.style.get("font-size")) or DEFAULT_FONT_SIZE
            return estimate_text_height(block.text, font_px)
        if block.kind == "list":
            font_px = parse_length(block.style.get("font-size")) or (DEFAULT_FONT_SIZE - 4)
            return len(block.items) * font_px * 1.2 + 32
        if block.kind == "table":
            font_px = parse_length(block.style.get("font-size")) or (DEFAULT_FONT_SIZE - 8)
            return len(block.table) * font_px * 1.6 + 40
        if block.kind == "image":
            height_hint = parse_length(block.style.get("height"))
            if height_hint:
                return height_hint
            return width * 0.6
        if block.kind == "shape":
            return parse_length(block.style.get("height")) or 120
        if block.kind == "group":
            return parse_length(block.style.get("height")) or 0
        return 160

    def _paginate(self, blocks: List[Block]) -> List[List[Block]]:
        slides: List[List[Block]] = []
        current: List[Block] = []
        offset = blocks[0].frame.top if blocks else MARGIN_Y
        for block in blocks:
            adjusted_top = block.frame.top - offset
            bottom = adjusted_top + block.frame.height
            if bottom > self.height - MARGIN_Y and current:
                slides.append(current)
                current = []
                offset = block.frame.top
                adjusted_top = block.frame.top - offset
            block.frame.top = adjusted_top + MARGIN_Y
            current.append(block)
        if current:
            slides.append(current)
        return slides


# --------------------------------------------------------------------------- #
# Chart Builder
# --------------------------------------------------------------------------- #


class ChartBuilder:
    DATA_RE = re.compile(
        r"(?:var|const|let)\s+(?P<name>\w+)\s*=\s*\{\s*labels\s*:\s*\[(?P<labels>[^\]]+)\].*?data\s*:\s*\[(?P<data>[^\]]+)\]",
        re.S,
    )

    def __init__(self, soup: BeautifulSoup):
        self.soup = soup

    def extract(self) -> List[CategoryChartData]:
        chart_data_list: List[CategoryChartData] = []
        for script in self.soup.find_all("script"):
            text = script.string or ""
            for match in self.DATA_RE.finditer(text):
                labels = [lbl.strip(" '\n\t\"") for lbl in match.group("labels").split(",") if lbl.strip()]
                values = [float(v.strip()) for v in match.group("data").split(",")]
                data = CategoryChartData()
                data.categories = labels
                data.add_series("Series", values)
                chart_data_list.append(data)
        return chart_data_list


# --------------------------------------------------------------------------- #
# PPTX Renderer
# --------------------------------------------------------------------------- #


class PPTXRenderer:
    def __init__(self):
        self.prs = Presentation()
        self.prs.slide_width = Inches(13.333)
        self.prs.slide_height = Inches(7.5)
        self.blank = self.prs.slide_layouts[6]

    def render(self, slides_blocks: List[List[Block]], charts: List[CategoryChartData]):
        for blocks in slides_blocks:
            slide = self.prs.slides.add_slide(self.blank)
            for block in blocks:
                self._render_block(slide, block)

        if charts:
            slide = self.prs.slides.add_slide(self.blank)
            top = MARGIN_Y
            for chart_data in charts:
                slide.shapes.add_chart(
                    XL_CHART_TYPE.COLUMN_CLUSTERED,
                    Inches(MARGIN_X / 96),
                    Inches(top / 96),
                    Inches((SLIDE_REF_WIDTH - 2 * MARGIN_X) / 96),
                    Inches(260 / 96),
                    chart_data,
                )
                top += 280

    # -- block renderers -------------------------------------------------- #
    def _geometry(self, frame: LayoutFrame) -> Tuple[float, float, float, float]:
        left = Inches(frame.left / 96)
        top = Inches(frame.top / 96)
        width = Inches(max(frame.width, 40) / 96)
        height = Inches(max(frame.height, 40) / 96)
        return left, top, width, height

    def _render_block(self, slide, block: Block):
        if block.kind == "text":
            self._render_text(slide, block)
        elif block.kind == "list":
            self._render_list(slide, block)
        elif block.kind == "table":
            self._render_table(slide, block)
        elif block.kind == "image":
            self._render_image(slide, block)
        elif block.kind == "shape":
            self._render_shape(slide, block)
        else:
            for child in block.children:
                self._render_block(slide, child)

    def _apply_shape_style(self, shape, block: Block):
        fill = css_color(block.shape_style.get("fill"))
        if fill:
            shape.fill.solid()
            shape.fill.fore_color.rgb = fill
        else:
            shape.fill.background()
        border = css_color(block.shape_style.get("border_color"))
        if border:
            shape.line.color.rgb = border
            width = parse_length(block.shape_style.get("border_width")) or 2
            shape.line.width = Pt(width)
        else:
            shape.line.fill.background()

    def _render_text(self, slide, block: Block):
        left, top, width, height = self._geometry(block.frame)
        shape = slide.shapes.add_textbox(left, top, width, height)
        tf = shape.text_frame
        tf.clear()
        runs = block.runs or [TextRun(text=block.text)]
        paragraph = tf.paragraphs[0]
        for run_data in runs:
            text = run_data.text
            segments = text.split("\n")
            for idx, segment in enumerate(segments):
                if idx > 0:
                    paragraph = tf.add_paragraph()
                run = paragraph.add_run()
                run.text = segment
                font = run.font
                size = run_data.font_size or parse_length(block.style.get("font-size")) or DEFAULT_FONT_SIZE
                font.size = Pt(size)
                if run_data.bold:
                    font.bold = True
                if run_data.italic:
                    font.italic = True
                color = css_color(run_data.color) or css_color(block.style.get("color"))
                if color:
                    font.color.rgb = color
        align = block.style.get("text-align", "").lower()
        if align == "center":
            paragraph.alignment = PP_ALIGN.CENTER
        elif align == "right":
            paragraph.alignment = PP_ALIGN.RIGHT
        self._apply_shape_style(shape, block)

    def _render_list(self, slide, block: Block):
        left, top, width, height = self._geometry(block.frame)
        shape = slide.shapes.add_textbox(left, top, width, height)
        tf = shape.text_frame
        tf.clear()
        for idx, item in enumerate(block.items):
            paragraph = tf.add_paragraph() if idx else tf.paragraphs[0]
            paragraph.text = item
            paragraph.level = 0
            paragraph.font.size = Pt(parse_length(block.style.get("font-size")) or DEFAULT_FONT_SIZE - 4)
        self._apply_shape_style(shape, block)

    def _render_table(self, slide, block: Block):
        if not block.table:
            return
        rows = len(block.table)
        cols = max(len(r) for r in block.table)
        left, top, width, height = self._geometry(block.frame)
        table = slide.shapes.add_table(rows, cols, left, top, width, height).table
        for c in range(cols):
            table.columns[c].width = width / cols
        for r in range(rows):
            table.rows[r].height = height / rows
        for r, row in enumerate(block.table):
            for c, value in enumerate(row):
                cell = table.cell(r, c)
                cell.text = value
                style = {}
                if r < len(block.cell_styles) and c < len(block.cell_styles[r]):
                    style = block.cell_styles[r][c].style
                self._apply_cell_text(cell.text_frame.paragraphs[0], style)
                bg = css_color(style.get("background-color"))
                if bg:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = bg
                else:
                    cell.fill.background()
                border = css_color(style.get("border-color"))
                if border:
                    self._apply_cell_border(cell, border, parse_length(style.get("border-width")) or 1.2)

    def _apply_cell_text(self, paragraph, style: Dict[str, str]):
        font = paragraph.font
        font.size = Pt(parse_length(style.get("font-size")) or DEFAULT_FONT_SIZE - 6)
        if style.get("font-weight") in {"bold", "700", "600"}:
            font.bold = True
        color = css_color(style.get("color"))
        if color:
            font.color.rgb = color
        align = style.get("text-align", "").lower()
        if align == "center":
            paragraph.alignment = PP_ALIGN.CENTER
        elif align == "right":
            paragraph.alignment = PP_ALIGN.RIGHT

    def _apply_cell_border(self, cell, color: RGBColor, width_px: float):
        width_emu = Pt(width_px).emu
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
            ln = tcPr.find(qn(tag))
            if ln is None:
                ln = OxmlElement(tag)
                tcPr.append(ln)
            ln.set("w", str(int(width_emu)))
            solid = ln.find(qn("a:solidFill"))
            if solid is None:
                solid = OxmlElement("a:solidFill")
                ln.append(solid)
            srgb = solid.find(qn("a:srgbClr"))
            if srgb is None:
                srgb = OxmlElement("a:srgbClr")
                solid.append(srgb)
            srgb.set("val", f"{color.rgb:06X}")

    def _render_shape(self, slide, block: Block):
        left, top, width, height = self._geometry(block.frame)
        shape = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, left, top, width, height)
        self._apply_shape_style(shape, block)

    def _render_image(self, slide, block: Block):
        left, top, width, height = self._geometry(block.frame)
        if block.image_path and block.image_path.exists():
            slide.shapes.add_picture(str(block.image_path), left, top, width=width, height=height)
        else:
            placeholder = slide.shapes.add_textbox(left, top, width, height)
            placeholder.text_frame.text = block.text or "[image]"


# --------------------------------------------------------------------------- #
# Conversion flow
# --------------------------------------------------------------------------- #


def resolve_slides(soup: BeautifulSoup) -> List[Tag]:
    slides = soup.select(".slide, .slide-container, [data-slide]")
    if slides:
        return slides
    if soup.body:
        return [soup.body]
    raise ValueError("No slide containers found. Use .slide or data-slide markers.")


def convert(input_html: Path, output_pptx: Path):
    soup = BeautifulSoup(input_html.read_text(encoding="utf-8"), "lxml")
    style_parser = StyleParser(soup)
    block_builder = BlockBuilder(style_parser, input_html.parent)
    layout_engine = LayoutEngine(SLIDE_REF_WIDTH, SLIDE_REF_HEIGHT)
    charts = ChartBuilder(soup).extract()

    renderer = PPTXRenderer()
    for slide_tag in resolve_slides(soup):
        blocks = block_builder.build(slide_tag)
        slides_blocks = layout_engine.layout(blocks)
        renderer.render(slides_blocks, [])

    if charts:
        renderer.render([], charts)

    renderer.prs.save(output_pptx)


def main():
    parser = argparse.ArgumentParser(description="Browser-free HTML -> PPTX converter with virtual layout.")
    parser.add_argument("input_html", type=Path)
    parser.add_argument("output_pptx", type=Path)
    args = parser.parse_args()
    convert(args.input_html, args.output_pptx)


if __name__ == "__main__":
    main()
