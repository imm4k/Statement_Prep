from __future__ import annotations

import sqlite3
from datetime import datetime
from typing import Dict

import config

from pptx.presentation import Presentation
from pptx.shapes.base import BaseShape
from pptx.slide import Slide
from lxml import etree

from ppt_objects import ObjectUpdater, UpdateContext
from ppt_text_replace import replace_tokens_in_shape


def _update_cover_title(slide: Slide, shape: BaseShape, prs: Presentation, ctx: UpdateContext) -> None:
    token_map = {"[T1]": ctx.t1_str}
    count = replace_tokens_in_shape(shape, token_map)
    print(f"cover_title replacements applied: {count}")


def _update_cover_subtitle(slide: Slide, shape: BaseShape, prs: Presentation, ctx: UpdateContext) -> None:
    token_map = {"[Investor]": ctx.investor}
    count = replace_tokens_in_shape(shape, token_map)
    print(f"cover_subtitle replacements applied: {count}")

def _set_cell_text_preserve_cell_format(cell, text: str) -> None:
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN
    from pptx.util import Pt

    ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}

    txBody = cell._tc.txBody
    p = txBody.find("a:p", ns)
    if p is None:
        cell.text_frame.text = text
        return

    pPr = p.find("a:pPr", ns)
    endParaRPr = p.find("a:endParaRPr", ns)

    algn = None
    if pPr is not None:
        algn = pPr.get("algn")

    typeface = None
    sz = None
    b = None
    i = None
    u = None
    color_val = None

    if endParaRPr is not None:
        sz = endParaRPr.get("sz")
        b = endParaRPr.get("b")
        i = endParaRPr.get("i")
        u = endParaRPr.get("u")

        latin = endParaRPr.find("a:latin", ns)
        if latin is not None:
            typeface = latin.get("typeface")

        srgb = endParaRPr.find(".//a:srgbClr", ns)
        if srgb is not None:
            color_val = srgb.get("val")

    cell.text_frame.text = text

    p0 = cell.text_frame.paragraphs[0]
    if algn == "ctr":
        p0.alignment = PP_ALIGN.CENTER
    elif algn == "l":
        p0.alignment = PP_ALIGN.LEFT
    elif algn == "r":
        p0.alignment = PP_ALIGN.RIGHT
    elif algn == "just":
        p0.alignment = PP_ALIGN.JUSTIFY

    if p0.runs:
        r0 = p0.runs[0]
        if typeface:
            r0.font.name = typeface
        if sz and str(sz).isdigit():
            r0.font.size = Pt(int(sz) / 100)
        if b is not None:
            r0.font.bold = (str(b) == "1")
        if i is not None:
            r0.font.italic = (str(i) == "1")
        if u is not None:
            r0.font.underline = (str(u).lower() != "none")
        if color_val and len(color_val) == 6:
            r0.font.color.rgb = RGBColor(int(color_val[0:2], 16), int(color_val[2:4], 16), int(color_val[4:6], 16))

def _update_summary_table(slide: Slide, shape: BaseShape, prs: Presentation, ctx: UpdateContext) -> None:
    if not hasattr(shape, "table"):
        return

    tbl = shape.table

    # column headers on row index 1
    duration_col = None
    for c in range(len(tbl.columns)):
        if tbl.cell(1, c).text.strip() == "Duration\n(Months)":
            duration_col = c
            break

    if duration_col is None:
        return

    sql = """
        SELECT property, MIN(acquired)
        FROM gl_agg
        WHERE investor = ?
          AND acquired IS NOT NULL
        GROUP BY property
    """

    con = sqlite3.connect(str(config.SQLITE_PATH))
    rows = con.execute(sql, (ctx.investor,)).fetchall()
    con.close()

    durations = {}

    for prop, acquired_val in rows:
        acquired_dt = datetime.strptime(str(acquired_val)[:10], "%Y-%m-%d")

        months = (
            (ctx.statement_thru_date_dt.year - acquired_dt.year) * 12
            + (ctx.statement_thru_date_dt.month - acquired_dt.month)
        )

        if ctx.statement_thru_date_dt.day < acquired_dt.day:
            months -= 1

        durations[str(prop)] = float(months)

    hits = 0

    # rows start at index 2, property in col 0
    for r in range(2, len(tbl.rows)):
        prop_name = tbl.cell(r, 0).text.strip()

        if prop_name not in durations:
            continue

        _set_cell_text_preserve_cell_format(tbl.cell(r, duration_col), str(durations[prop_name]))
        hits += 1

    print(f"summary_table updated rows: {hits}")


OBJECT_UPDATERS: Dict[str, ObjectUpdater] = {
    "cover_title": _update_cover_title,
    "cover_subtitle": _update_cover_subtitle,
    "summary_table": _update_summary_table,
}
