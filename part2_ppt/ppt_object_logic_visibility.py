from __future__ import annotations

import config

from pptx.presentation import Presentation
from pptx.shapes.base import BaseShape
from pptx.slide import Slide

from ppt_objects import UpdateContext

_P_NS = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main"}

def _set_shape_hidden_via_selection_pane(shape: BaseShape, hide: bool) -> None:
    """
    Properly hide or unhide a shape so PowerPoint actually renders it invisible.
    """
    try:
        # 1) Set the cNvPr hidden attribute (legacy method)
        cNvPr = shape._element.xpath(".//p:cNvPr", namespaces=_P_NS)
        if cNvPr:
            node = cNvPr[0]
            if hide:
                node.set("hidden", "1")
            elif "hidden" in node.attrib:
                del node.attrib["hidden"]

        # 2) Set shape properties noShowAsBullet/noClick so PPT UI respects it
        spPr = shape._element.xpath(".//p:spPr", namespaces=_P_NS)
        if spPr:
            spPr = spPr[0]
            if hide:
                spPr.set("noShowAsBullet", "1")
                spPr.set("noClick", "1")
            else:
                if "noShowAsBullet" in spPr.attrib:
                    del spPr.attrib["noShowAsBullet"]
                if "noClick" in spPr.attrib:
                    del spPr.attrib["noClick"]
    except Exception:
        return

def update_partial_ownership_visibility(slide: Slide, shape: BaseShape, prs: Presentation, ctx: UpdateContext) -> None:
    rules = getattr(config, "PARTIAL_OWNERSHIP_VISIBILITY_RULES", {}) or {}
    rule = rules.get(getattr(shape, "name", ""), None)

    is_partial = float(ctx.ownership_pct or 0.0) < 100.0

    if rule is None:
        if getattr(shape, "name", "") == "pct_owner_note":
            should_show = is_partial
        else:
            return
    else:
        should_show = bool(rule.get("partial" if is_partial else "full", True))

    _set_shape_hidden_via_selection_pane(shape, hide=(not should_show))
    print(f"visibility updated. shape: {shape.name}. visible: {should_show}. ownership_pct: {ctx.ownership_pct}%")
