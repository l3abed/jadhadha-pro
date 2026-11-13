# -*- coding: utf-8 -*-
"""
styles.py — طبقة أنماط للجذاذات الناتجة: تغيّر الألوان والحدود والخطوط فقط،
ولا تمسّ بنية القالب ولا نصوصه. تعريفات الأنماط مضمّنة في هذا الملف.
"""
import copy
import json
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

_HERE = Path(__file__).resolve().parent
_WHITE = {"ffffff", "auto", ""}


STYLES = {'classic': {'label': 'الأصلي', 'note': 'بلا تغيير: يبقى تنسيق القالب كما هو'}, 'emerald': {'label': 'زمردي ملوّن', 'header_fill': '1B5E20', 'header_text': 'FFFFFF', 'zebra_fill': 'F3F8F1', 'border': '1B5E20', 'text': '1E2A22', 'font': None}, 'modern': {'label': 'عصري', 'header_fill': '2E4A7D', 'header_text': 'FFFFFF', 'zebra_fill': 'EEF3FA', 'border': '8FA3C7', 'text': '1B1F2A', 'font': 'Arial'}, 'print_bw': {'label': 'أبيض وأسود (للطباعة)', 'header_fill': None, 'header_text': '000000', 'zebra_fill': None, 'border': '000000', 'text': '000000', 'font': None, 'bold_header': True}}


def load_styles() -> dict:
    """تعريفات الأنماط مضمّنة في الكود (لا تعتمد على ملف خارجي يُفقد أثناء التغليف)."""
    return copy.deepcopy(STYLES)


def _set_child(parent, tag, attrs):
    el = parent.find(qn(tag))
    if el is None:
        el = OxmlElement(tag)
        parent.append(el)
    for k, v in attrs.items():
        el.set(qn(k), v)
    return el


def _recolor_borders(container, color):
    for b in container.iter():
        if b.tag.endswith("}top") or b.tag.endswith("}left") or b.tag.endswith("}bottom") \
           or b.tag.endswith("}right") or b.tag.endswith("}insideH") or b.tag.endswith("}insideV"):
            if b.get(qn("w:val")) not in (None, "nil", "none"):
                b.set(qn("w:color"), color)


def _is_header_cell(tc, first_row: bool) -> bool:
    if first_row:
        return True
    shd = tc.find(qn("w:tcPr") + "/" + qn("w:shd"))
    if shd is None:
        return False
    fill = (shd.get(qn("w:fill")) or "").lower()
    return fill not in _WHITE and fill != ""


def _style_table(tbl, st: dict, depth: int = 0):
    rows = tbl.findall(qn("w:tr"))
    for ri, tr in enumerate(rows):
        for tc in tr.findall(qn("w:tc")):
            tcPr = tc.find(qn("w:tcPr"))
            if tcPr is None:
                tcPr = OxmlElement("w:tcPr")
                tc.insert(0, tcPr)
            header = _is_header_cell(tc, ri == 0)
            fill = st.get("header_fill") if header else (st.get("zebra_fill") if ri % 2 == 0 else None)
            shd = tcPr.find(qn("w:shd"))
            if fill:
                if shd is None:
                    shd = OxmlElement("w:shd")
                    tcPr.append(shd)
                shd.set(qn("w:val"), "clear")
                shd.set(qn("w:color"), "auto")
                shd.set(qn("w:fill"), fill)
            elif shd is not None:
                tcPr.remove(shd)
            color = st.get("header_text") if header else st.get("text")
            for r in tc.iter(qn("w:r")):
                rPr = r.find(qn("w:rPr"))
                if rPr is None:
                    rPr = OxmlElement("w:rPr")
                    r.insert(0, rPr)
                if color:
                    _set_child(rPr, "w:color", {"w:val": color})
                if st.get("font"):
                    rf = _set_child(rPr, "w:rFonts", {})
                    for a in ("w:ascii", "w:hAnsi", "w:cs"):
                        rf.set(qn(a), st["font"])
                if header and st.get("bold_header"):
                    if rPr.find(qn("w:b")) is None:
                        rPr.append(OxmlElement("w:b"))
    tblPr = tbl.find(qn("w:tblPr"))
    if st.get("border") and tblPr is not None:
        _recolor_borders(tblPr, st["border"])
    for tc in tbl.iter(qn("w:tc")):
        tcb = tc.find(qn("w:tcPr") + "/" + qn("w:tcBorders"))
        if tcb is not None and st.get("border"):
            _recolor_borders(tcb, st["border"])


def apply_style(docx_path, style_name: str) -> bool:
    """يطبّق النمط على الملف في مكانه. يعيد False إن كان النمط «classic» أو غير معروف (لا تغيير)."""
    styles = load_styles()
    key = style_name or "classic"
    if key == "classic":
        return False
    st = styles.get(key)
    if not st:
        raise ValueError(f"نمط غير معروف: {key}")
    doc = Document(str(docx_path))
    for tbl in doc.element.body.iter(qn("w:tbl")):
        _style_table(tbl, st)
    doc.save(str(docx_path))
    return True
