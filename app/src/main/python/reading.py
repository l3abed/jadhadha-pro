#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
jadhadha_build.py — نسخة معدلة بالأسماء الافتراضية
تدعم الجداول الفرعية وإطارات النصوص الداخلية.

الذكاء الاصطناعي (اختياري): يملأ الخانات الفارغة فقط (التقويم، المحتوى، أنشطة المدرس، القدرات)
وأهداف الكفايات الناقصة في الترويسة؛ ولا يمسّ أي خانة فيها نص من المصدر.

مفتاح Gemini لا يُكتب داخل السكريبت. ضعه في أحد الموضعين:
  • متغير بيئة:  GEMINI_API_KEY        (ويندوز:  setx GEMINI_API_KEY "مفتاحك")
  • ملف .env بجوار السكريبت أو في المجلد الحالي، فيه سطر:  GEMINI_API_KEY=مفتاحك
وبدون مفتاح (أو مع --no-ai) يعمل السكريبت كما كان دون ذكاء اصطناعي.
"""
import argparse
import copy
import json
import math
import os
import random
import re
import sys
import time
import urllib.error
import urllib.request
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.table import _Cell

# ── الأسماء الافتراضية للملفات ──
DEFAULT_MODEL = "المنفى.docx"
DEFAULT_SOURCE = "جذاذة--- --(مسيرة امرأة)المختار-3إع.docx"
DEFAULT_OUT = "جذاذة مسيرة امرأة.docx"
OUT_DIR = "الجذاذات_المنجزة"
DEFAULT_AI_MODEL = os.environ.get("GEMINI_MODEL", "gemini-3.5-flash")
DEFAULT_API_KEY = ""   # لا مفتاح داخل الكود: يُمرَّر من إعدادات التطبيق أو GEMINI_API_KEY

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
COL = {"time": 0, "eval": 1, "ind": 2, "content": 3, "learner": 4, "teacher": 5,
       "stage": 6, "step": 7, "abil": 8}


class ShapeError(SystemExit):
    """تخطيط الملف المعطى غير المتوقع."""


def open_docx(path, what):
    try:
        return Document(path)
    except Exception as e:
        raise ShapeError(f"تعذّر فتح {what} كملف Word (.docx) صالح ({type(e).__name__}). ")


# ═══════════════════════════ تنظيف النصوص ═══════════════════════════
TYPOS = [("شفوياُ", "شفويا"), ("تكوني", "تكويني"), ("المتعلون", "المتعلمون"),
         ("فرضيامت وضوع", "فرضيات موضوع"), ("ّإ", "إ"), ("القراء والكتابة", "القراءة والكتابة"),
         ("المحورية النص", "المحورية للنص"), ("قمية", "قيمة")]


def clean(s):
    s = s.replace("\u0640", "").replace("\u00a0", " ")
    s = re.sub(r"[ \t]+", " ", s).strip()
    s = re.sub(r"\s+([:،؛؟\.])", r"\1", s)
    for a, b in TYPOS:
        s = s.replace(a, b)
    s = re.sub(r"\(\s+", "(", s)
    s = re.sub(r"\s+\)", ")", s)
    s = re.sub(r'"([^"]+)"', r"«\1»", s)
    s = re.sub(r"،(?=\S)", "، ", s)
    s = re.sub(r":(?=[(«])", ": ", s)
    s = re.sub(r"«\s+", "«", s)
    s = re.sub(r"\s+»", "»", s)
    return s


TERMINAL = (".", "؟", "!", ":", ")", "»", "…", "؛")
MARKER = re.compile(r"^\s*(?:[+\-*•]|=+|\d+\s*[-–.)]|[أ-ي]\s*[-.])\s*")


def has_marker(s):
    return bool(MARKER.match(s))


def strip_marker(s):
    prev = None
    while prev != s:
        prev = s
        s = MARKER.sub("", s, count=1)
    return s.strip()


def end(t):
    t = t.strip()
    return t if (not t or t[-1] in ".؟!:)»…،؛") else t + "."


def key(t):
    return re.sub(r"\s+", " ", strip_marker(t)).strip()


def join_items(raw):
    out, blank = [], True
    for ln in raw:
        if not ln:
            blank = True
            continue
        if (blank or not out or has_marker(ln) or ln.startswith("{") or ln.startswith("⟦")
                or out[-1].startswith("⟦") or out[-1].rstrip().endswith(TERMINAL)):
            out.append(ln)
        else:
            out[-1] = out[-1].rstrip() + " " + ln
        blank = False
    return out
# ═══════════════════════════ تنسيق البنود ═══════════════════════════
def T(t):
    return "+ " + end(strip_marker(t))

def D(t):
    return "- " + end(strip_marker(t))

def H(t):
    return f"<u>{t}</u>"

def split_label(t, max_words=6):
    if ":" in t:
        lab, rest = t.split(":", 1)
        if lab.strip() and len(lab.split()) <= max_words:
            return lab.strip(), rest.strip()
    return None

def label_item(t, bullet="- "):
    t = strip_marker(t)
    sp = split_label(t)
    if sp:
        lab, rest = sp
        rest = re.sub(r"^\s*-\s+", "", rest).strip()
        return f"{bullet}<u>{lab}</u>:" + (f" {rest}" if rest else "")
    return bullet + t

def label_only(t):
    return t.endswith(":") and ":" not in t[:-1]

class ContentFmt:
    def __init__(self, num_style="heading", number_headings=False):
        self.num_style = num_style
        self.number_headings = number_headings
        self.k = 0
        self.child = False

    def heading(self, text):
        text = text.strip()
        if "الحقول المعجمية" in text:
            text = "معجم النص (" + text.rstrip(":").strip() + "):"
        self.k += 1
        return H(f"{self.k}- {text}")

    def fmt(self, items):
        out = []
        for idx, it in enumerate(items):
            t = it.strip()
            if TBL_RE.search(t):
                out.append(t)
                self.child = False
                continue
            m_num = re.match(r"^(\d+)\s*[-–.)]\s*(.*)$", t)
            m_let = re.match(r"^([أ-ي])\s*[-.]\s*(.*)$", t)
            m_eq = re.match(r"^=+\s*(.*)$", t)
            m_bul = re.match(r"^([+\-*•])\s*(.*)$", t)
            if m_num:
                rest = m_num.group(2).strip()
                if self.num_style == "item":
                    out.append(label_item(rest, "- "))
                    self.child = label_only(rest)
                elif label_only(rest):
                    out.append(self.heading(rest))
                else:
                    sp = split_label(rest, 8)
                    self.k += 1
                    if sp:
                        out.append(f"<u>{self.k}- {sp[0]}:</u>" + (f" {sp[1]}" if sp[1] else ""))
                    else:
                        out.append(f"{self.k}- {rest}")
            elif m_let:
                rest = m_let.group(2).strip()
                sp = split_label(rest)
                if sp:
                    out.append(f"{m_let.group(1)}- <u>{sp[0]}</u>:" + (f" {sp[1]}" if sp[1] else ""))
                else:
                    out.append(f"{m_let.group(1)}- {rest}")
                self.child = False
            elif m_eq:
                out.append(m_eq.group(1).strip())
                self.child = False
            elif m_bul:
                ch, rest = m_bul.group(1), m_bul.group(2).strip()
                parts = re.split(r"\s+(?=\+\s)", t) if ch == "+" else [t]
                for part in parts:
                    ch2 = "+" if (ch == "+" or (self.child and ch == "-")) else "-"
                    out.append(label_item(part, f"{ch2} "))
            else:
                if idx == 0 and label_only(t):
                    if self.number_headings:
                        out.append(self.heading(t))
                    else:
                        out.append(H(t))
                else:
                    sp = split_label(t, 4)
                    out.append(f"<u>{sp[0]}:</u> {sp[1]}" if sp and sp[1] else t)
                self.child = False
        return out


# ═══════════════════════════ عناصر OOXML ═══════════════════════════
def segs(markup):
    out, pos = [], 0
    for m in re.finditer(r"<u>(.*?)</u>", markup, re.S):
        if m.start() > pos:
            out.append((markup[pos:m.start()], False))
        out.append((m.group(1), True))
        pos = m.end()
    if pos < len(markup):
        out.append((markup[pos:], False))
    return out

def _fonts(rpr):
    f = OxmlElement("w:rFonts")
    for a in ("w:ascii", "w:hAnsi", "w:cs"):
        f.set(qn(a), "Sakkal Majalla")
    rpr.append(f)
    rpr.append(OxmlElement("w:b"))
    rpr.append(OxmlElement("w:bCs"))
    
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), "18")
    rpr.append(sz)
    
    szcs = OxmlElement("w:szCs")
    szcs.set(qn("w:val"), "18")
    rpr.append(szcs)

def mk_run(text, u=False):
    r = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    _fonts(rpr)
    if u:
        uu = OxmlElement("w:u"); uu.set(qn("w:val"), "single"); rpr.append(uu)
    rpr.append(OxmlElement("w:rtl"))
    r.append(rpr)
    t = OxmlElement("w:t"); t.set(qn("xml:space"), "preserve"); t.text = text
    r.append(t)
    return r

def mk_par(markup, kind):
    p = OxmlElement("w:p")
    ppr = OxmlElement("w:pPr")
    ppr.append(OxmlElement("w:bidi"))
    
    # 1. فرض المحاذاة في الوسط (Center) على جميع الخانات والمحتويات
    jc = OxmlElement("w:jc")
    jc.set(qn("w:val"), "center")
    ppr.append(jc)
    
    # 2. إزالة الهوامش الداخلية تماماً لتوفير المساحة وعدم هدرها
    ind = OxmlElement("w:ind")
    ind.set(qn("w:right"), "0")
    ind.set(qn("w:left"), "0")
    ind.set(qn("w:firstLine"), "0")
    ppr.append(ind)
    
    mark = OxmlElement("w:rPr")
    _fonts(mark)
    ppr.append(mark)
    p.append(ppr)
    
    for text, u in segs(markup or ""):
        if text:
            p.append(mk_run(text, u))
    return p



# ═══════════════════════════ الجداول المتداخلة: تُحفظ جداول لا نصوصاً ═══════════════════════════
NESTED = {}                                   # رقم ← نسخة الجدول الأصلي
TBL_RE = re.compile(r"⟦جدول (\d+)⟧")


def is_content_table(tbl):
    """جدول محتوى (خليتان مملوءتان فأكثر) لا حاوية نص من خلية واحدة."""
    cells = [tc for tr in tbl.findall(W + "tr") for tc in tr.findall(W + "tc")]
    filled = [tc for tc in cells if "".join(t.text or "" for t in tc.iter(W + "t")).strip()]
    return len(cells) >= 2 and len(filled) >= 2


_NESTED_IDX = {}                              # بصمة الجدول ← رقمه (فلا يتكرر لو قُرئت الخانة مرتين)


def reset_nested():
    NESTED.clear()
    _NESTED_IDX.clear()


def register_table(el):
    """يحفظ نسخة الجدول ويعيد سطر علامة موضعه (ينتهي بنقطتين فلا تُضاف له نقطة)."""
    import hashlib
    from lxml import etree
    sig = hashlib.md5(etree.tostring(el)).hexdigest()
    k = _NESTED_IDX.get(sig)
    if k is None:
        k = len(NESTED)
        NESTED[k] = copy.deepcopy(el)
        _NESTED_IDX[sig] = k
    return f"⟦جدول {k}⟧:"


_TBLPR_ORDER = ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize", "tblStyleColBandSize",
                "tblW", "jc", "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook",
                "tblCaption", "tblDescription"]
_RPR_AFTER_SZ = ["highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang",
                 "eastAsianLayout", "specVanish", "oMath"]


def _local(el):
    return el.tag.split("}")[1] if "}" in el.tag else el.tag


def _sort_children(parent, order):
    kids = list(parent)
    kids.sort(key=lambda e: order.index(_local(e)) if _local(e) in order else len(order))
    for k in kids:
        parent.remove(k)
    for k in kids:
        parent.append(k)


def cell_width(tc, default=3000):
    w = tc.find(qn("w:tcPr") + "/" + qn("w:tcW"))
    try:
        if w is not None and w.get(qn("w:type"), "dxa") == "dxa":
            return int(w.get(qn("w:w")))
    except (TypeError, ValueError):
        pass
    return default


def _fix_run_props(r, size):
    rpr = r.find(qn("w:rPr"))
    if rpr is None:
        rpr = OxmlElement("w:rPr")
        r.insert(0, rpr)
    for tag in ("w:rFonts", "w:sz", "w:szCs"):
        for x in rpr.findall(qn(tag)):
            rpr.remove(x)
    f = OxmlElement("w:rFonts")
    for a in ("w:ascii", "w:hAnsi", "w:cs"):
        f.set(qn(a), "Sakkal Majalla")
    rpr.insert(1 if rpr.find(qn("w:rStyle")) is not None else 0, f)
    pos = next((i for i, c in enumerate(rpr) if _local(c) in _RPR_AFTER_SZ), len(rpr))
    for tag in ("w:szCs", "w:sz"):
        z = OxmlElement(tag)
        z.set(qn("w:val"), str(size))
        rpr.insert(pos, z)


def materialize(el, width):
    """نسخة من الجدول الأصلي مضبوطة العرض والخط لتلائم خانة القالب (دون تغيير محتواه)."""
    t = copy.deepcopy(el)
    width = max(1200, width - 200)
    tblPr = t.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        t.insert(0, tblPr)
    had_borders = tblPr.find(qn("w:tblBorders")) is not None or any(True for _ in t.iter(qn("w:tcBorders")))
    for tag in ("w:tblpPr", "w:tblInd", "w:tblW", "w:tblLayout", "w:jc", "w:tblCaption", "w:tblDescription"):
        for x in tblPr.findall(qn(tag)):
            tblPr.remove(x)
    tw = OxmlElement("w:tblW"); tw.set(qn("w:w"), str(width)); tw.set(qn("w:type"), "dxa"); tblPr.append(tw)
    jc = OxmlElement("w:jc"); jc.set(qn("w:val"), "center"); tblPr.append(jc)
    lay = OxmlElement("w:tblLayout"); lay.set(qn("w:type"), "fixed"); tblPr.append(lay)
    if not had_borders:                                   # جدول بلا حدود: خطوط رفيعة ليبقى جدولاً ظاهراً
        b = OxmlElement("w:tblBorders")
        for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
            e = OxmlElement("w:" + side)
            e.set(qn("w:val"), "single"); e.set(qn("w:sz"), "4"); e.set(qn("w:space"), "0"); e.set(qn("w:color"), "7F7F7F")
            b.append(e)
        tblPr.append(b)
    _sort_children(tblPr, _TBLPR_ORDER)

    rows = t.findall(qn("w:tr"))
    grid = t.find(qn("w:tblGrid"))
    old = []
    if grid is not None:
        for g in grid.findall(qn("w:gridCol")):
            try:
                old.append(int(g.get(qn("w:w"))))
            except (TypeError, ValueError):
                old.append(0)
    ncols = len(old) or max((sum(int((tc.find(qn("w:tcPr") + "/" + qn("w:gridSpan")).get(qn("w:val"))) if tc.find(qn("w:tcPr") + "/" + qn("w:gridSpan")) is not None else 1) for tc in r.findall(qn("w:tc"))) for r in rows), default=1)
    if not old or sum(old) <= 0:
        old = [1] * ncols
    tot = float(sum(old))
    new = [max(200, int(round(w * width / tot))) for w in old]
    if grid is None:
        grid = OxmlElement("w:tblGrid")
        tblPr.addnext(grid)
        for _ in new:
            grid.append(OxmlElement("w:gridCol"))
    for g, w in zip(grid.findall(qn("w:gridCol")), new):
        g.set(qn("w:w"), str(w))

    size = 16 if ncols >= 3 else 18
    for tr in rows:
        trPr = tr.find(qn("w:trPr"))
        if trPr is not None:
            for x in trPr.findall(qn("w:trHeight")):
                trPr.remove(x)
        ci = 0
        for tc in tr.findall(qn("w:tc")):
            tcPr = tc.find(qn("w:tcPr"))
            if tcPr is None:
                tcPr = OxmlElement("w:tcPr")
                tc.insert(0, tcPr)
            gs = tcPr.find(qn("w:gridSpan"))
            span = int(gs.get(qn("w:val"))) if gs is not None else 1
            for x in tcPr.findall(qn("w:tcW")):
                tcPr.remove(x)
            cw = OxmlElement("w:tcW")
            cw.set(qn("w:w"), str(sum(new[ci:ci + span]) or new[-1])); cw.set(qn("w:type"), "dxa")
            tcPr.insert(1 if tcPr.find(qn("w:cnfStyle")) is not None else 0, cw)
            ci += span
            for p in tc.iter(qn("w:p")):
                old_ppr = p.find(qn("w:pPr"))
                jc_old = old_ppr.find(qn("w:jc")) if old_ppr is not None else None
                mark = old_ppr.find(qn("w:rPr")) if old_ppr is not None else None
                if old_ppr is not None:
                    p.remove(old_ppr)
                ppr = OxmlElement("w:pPr")
                ppr.append(OxmlElement("w:bidi"))
                sp = OxmlElement("w:spacing"); sp.set(qn("w:before"), "0"); sp.set(qn("w:after"), "0"); ppr.append(sp)
                ind = OxmlElement("w:ind"); ind.set(qn("w:left"), "0"); ind.set(qn("w:right"), "0"); ind.set(qn("w:firstLine"), "0"); ppr.append(ind)
                j = OxmlElement("w:jc"); j.set(qn("w:val"), jc_old.get(qn("w:val")) if jc_old is not None else "center"); ppr.append(j)
                if mark is not None:
                    ppr.append(mark)
                p.insert(0, ppr)
            for r in tc.iter(qn("w:r")):
                _fix_run_props(r, size)
    return t


def spacer_par():
    p = OxmlElement("w:p")
    ppr = OxmlElement("w:pPr")
    ppr.append(OxmlElement("w:bidi"))
    sp = OxmlElement("w:spacing"); sp.set(qn("w:before"), "0"); sp.set(qn("w:after"), "0"); sp.set(qn("w:line"), "40"); sp.set(qn("w:lineRule"), "exact")
    ppr.append(sp)
    rpr = OxmlElement("w:rPr")
    z = OxmlElement("w:sz"); z.set(qn("w:val"), "2"); rpr.append(z)
    ppr.append(rpr)
    p.append(ppr)
    return p


def fill(cell, specs, kind):
    tc = cell._tc
    for p in tc.findall(qn("w:p")):
        tc.remove(p)
    for tb in tc.findall(qn("w:tbl")):
        tc.remove(tb)
    width = cell_width(tc)
    last_table = False
    for s in (specs or [""]):
        m = TBL_RE.search(s or "")
        if m and int(m.group(1)) in NESTED:
            tc.append(materialize(NESTED[int(m.group(1))], width))
            tc.append(spacer_par())
            last_table = False
        else:
            tc.append(mk_par(s, kind))

def tcell(tbl, r, c):
    return _Cell(tbl.rows[r]._tr.findall(qn("w:tc"))[c], tbl)

def set_cell_keep_format(cell, texts):
    tc = cell._tc
    ps = tc.findall(qn("w:p"))
    p0 = ps[0]
    runs = p0.findall(qn("w:r"))
    def rtext(r): return "".join(t.text or "" for t in r.findall(qn("w:t")))
    base = max(runs, key=lambda r: len(rtext(r))) if runs else None
    rpr = None
    if base is not None and base.find(qn("w:rPr")) is not None:
        rpr = copy.deepcopy(base.find(qn("w:rPr")))
        if rpr.find(qn("w:rtl")) is None:
            rtl = OxmlElement("w:rtl")
            lang = rpr.find(qn("w:lang"))
            lang.addprevious(rtl) if lang is not None else rpr.append(rtl)
    ppr = copy.deepcopy(p0.find(qn("w:pPr"))) if p0.find(qn("w:pPr")) is not None else None
    for p in ps:
        tc.remove(p)
    for txt in texts:
        p = OxmlElement("w:p")
        if ppr is not None: p.append(copy.deepcopy(ppr))
        r = OxmlElement("w:r")
        if rpr is not None: r.append(copy.deepcopy(rpr))
        t = OxmlElement("w:t"); t.set(qn("xml:space"), "preserve"); t.text = txt
        r.append(t)
        p.append(r)
        tc.append(p)
# ═══════════════════════════ قراءة المصدر ═══════════════════════════
class Box:
    def __init__(self, lines, anchor):
        self.lines, self.anchor, self.kind = lines, anchor, None

class Source:
    def __init__(self, path):
        reset_nested()
        self.doc = open_docx(path, "الجذاذة المصدر")
        if not self.doc.tables: raise ShapeError("لم أجد جدولًا في الجذاذة المصدر.")
        self.tbl = self.doc.tables[0]
        if len(self.tbl.rows) != 11 or len(self.tbl.rows[1].cells) < 9:
            raise ShapeError("تخطيط مختلف عن المختار 3 إع.")
        self.boxes = self._read_boxes()

    def raw(self, row, col):
        cells = self.tbl.rows[row].cells
        idx = COL[col]
        c = cells[idx]
        for k in range(idx):
            if cells[k]._tc is c._tc:
                return []
        
        def get_lines(element):
            if element.tag == W + "p":
                direct_text = []
                for r in element.findall(W + "r"):
                    for child in r:
                        if child.tag == W + "t" and child.text:
                            direct_text.append(child.text)
                        elif child.tag == W + "tab":
                            direct_text.append(" ")
                        elif child.tag in (W + "drawing", W + "pict"):
                            if direct_text:
                                yield "".join(direct_text)
                                direct_text = []
                            for txbx in child.findall(f".//{W}txbxContent"):
                                for block in txbx: yield from get_lines(block)
                if direct_text: yield "".join(direct_text)
            elif element.tag == W + "tbl" and is_content_table(element):
                yield register_table(element)
            elif element.tag == W + "tbl":
                for tr in element.findall(W + "tr"):
                    row_texts = []
                    for tc in tr.findall(W + "tc"):
                        tc_lines = []
                        for block in tc: tc_lines.extend(list(get_lines(block)))
                        txt = " ".join(tc_lines).strip()
                        if txt: row_texts.append(txt)
                    if row_texts:
                        if len(row_texts) > 1: yield " : ".join(row_texts)
                        else: yield row_texts[0]
            elif element.tag == W + "sdt":
                sdt_content = element.find(f".//{W}sdtContent")
                if sdt_content is not None:
                    for block in sdt_content: yield from get_lines(block)

        out = []
        for child in c._tc:
            for line in get_lines(child):
                for ln in line.split("\n"):
                    if ln.strip(): out.append(clean(ln))
        return out

    def items(self, row, col):
        return join_items(self.raw(row, col))

    def _read_boxes(self):
        boxes = []
        for tb in self.doc.element.body.iter(W + "txbxContent"):
            curr = tb
            inside_tbl = False
            while curr is not None:
                if curr.tag == W + "tbl" and curr is self.tbl._tbl:
                    inside_tbl = True
                    break
                curr = curr.getparent()
            if inside_tbl: continue
            lines = [clean("".join(t.text or "" for t in p.iter(W + "t"))) for p in tb.iter(W + "p")]
            lines = [l for l in lines if l]
            b = Box(lines, None)
            b.kind = self._classify(b)
            boxes.append(b)
        return boxes

    @staticmethod
    def _classify(b):
        if not b.lines: return "ignore"
        full_text = re.sub(r"[\u064b-\u065f]", "", " ".join(b.lines))
        if "الفئة المستهدفة" in full_text: return "info"
        if "كفاية" in full_text and ("تواصلية" in full_text or "منهجية" in full_text): return "comps"
        if re.search(r"الحصة\s+الا?ول[ىي]", full_text): return "s1"
        if re.search(r"الحصة\s+الثانية", full_text): return "s2"
        if "ينتقى منها" in full_text: return "note"
        if any(w in full_text for w in ["جذاذة", "الهدف"]): return "ignore"
        return "extra"

    def box(self, kind):
        for b in self.boxes:
            if b.kind == kind: return b
        
        # حل ذكي: إذا لم يجد مربع النص، يعوضه تلقائياً بدل إيقاف البرنامج
        if kind == "s1": return Box(["الحصة الأولى"], None)
        if kind == "s2": return Box(["الحصة الثانية"], None)
        
        raise ShapeError(f"لم أجد في المصدر مربع النص المتوقع ({kind}).")

class Report:
    def __init__(self):
        self.dropped, self.notes = [], []
    def drop(self, where, text):
        self.dropped.append((where, text))

FLUFF = [r"^يحفز المتعلمين", r"يجلب اهتمامهم", r"^يوجه المتعلمين إلى ملاحظة",
         r"^يساعد المتعلمين على فهم النص عبر الحوار", r"^يذلل صعوبات", r"^يساعد المتعلمين على تحديد رسالة النص"]
KEEP_LEARNER = [r"الكلمات الغامضة"]
TEACHER_VERB = re.compile(r"^ي[\u0621-\u064a]{2,}")

def is_fluff(t): return any(re.search(p, t) for p in FLUFF)

def act_lines(items, rep, where, skip=None):
    out = []
    for it in items:
        base = strip_marker(it)
        if skip and re.search(skip, base): continue
        if is_fluff(base):
            rep.drop(where, base)
            continue
        ch = (re.match(r"^\s*([+\-*•])", it) or [None, ""])[1]
        numbered = bool(re.match(r"^\s*\d+\s*[-–.)]", it))
        if ch == "-":
            for part in re.split(r"\s+(?=-\s)", it.strip()): out.append(D(part))
        elif numbered: out.append(D(it))
        elif TEACHER_VERB.match(base): out.append(T(base))
        else: out.append(D(base))
    return out

def eval_lines(items):
    out, prev_d = [], False
    for i, it in enumerate(items):
        ch = (re.match(r"^\s*([+\-*•])", it) or [None, ""])[1]
        if i > 0 and (ch == "-" or (ch == "" and prev_d)):
            out.append(D(it))
            prev_d = True
        else:
            out.append(T(it))
            prev_d = False
    return out

def find(items, pattern, name, col, start=0):
    for i in range(start, len(items)):
        if re.search(pattern, items[i]): return i
    raise ShapeError(f"لم أجد علامة «{name}» في عمود «{col}».")

def uniq_items(src, rows, col):
    seen, out = set(), []
    for r in rows:
        for it in src.items(r, col):
            if key(it) not in seen:
                seen.add(key(it))
                out.append(it)
    return out
# ═══════════════════════════ الذكاء الاصطناعي: ملء الفراغات ═══════════════════════════
class AIError(Exception):
    pass


class AICfg:
    def __init__(self, model, key, retries=4, timeout=120):
        self.model, self.key, self.retries, self.timeout = model, key, retries, timeout


def resolve_api_key(explicit=""):
    """المفتاح من: الوسيط --api-key ← متغير البيئة ← ملف .env ← المفتاح الافتراضي في أعلى الملف."""
    if explicit and explicit.strip():
        return explicit.strip()
    names = ("GEMINI_API_KEY", "GOOGLE_API_KEY")
    for name in names:
        if os.environ.get(name, "").strip():
            return os.environ[name].strip()
    for base in dict.fromkeys((Path.cwd(), Path(__file__).resolve().parent, Path("/storage/emulated/0"))):
        f = base / ".env"
        try:
            if not f.is_file():
                continue
            for line in f.read_text(encoding="utf-8", errors="ignore").splitlines():
                k, sep, v = line.strip().partition("=")
                if sep and k.strip() in names and v.strip():
                    return v.strip().strip("'\"")
        except OSError:
            continue
    return DEFAULT_API_KEY


def _candidate_text(data):
    cands = data.get("candidates") or []
    if not cands:
        reason = (data.get("promptFeedback") or {}).get("blockReason", "")
        raise AIError("استجابة بلا نتيجة" + (f" ({reason})" if reason else ""))
    parts = (cands[0].get("content") or {}).get("parts") or []
    text = "".join(p.get("text", "") for p in parts if isinstance(p, dict)).strip()
    if not text:
        raise AIError("استجابة فارغة")
    return text


def _parse_model_json(text):
    text = re.sub(r"^```(?:json)?\s*|\s*```$", "", text.strip())
    i, j = text.find("{"), text.rfind("}")
    if i < 0 or j <= i:
        raise AIError("لا يوجد JSON في الاستجابة")
    obj = json.loads(text[i:j + 1])
    if not isinstance(obj, dict) or not isinstance(obj.get("rows"), list):
        raise AIError("بنية JSON غير متوقعة")
    return obj


def _http_error_message(e):
    try:
        return str(json.loads(e.read().decode("utf-8", "ignore")).get("error", {}).get("message", ""))[:200]
    except Exception:
        return ""


def call_gemini(prompt, ai):
    """يستدعي Gemini مع إعادة محاولة بتأخير متزايد للأخطاء العابرة (503/429/انقطاع الشبكة/JSON تالف)."""
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{ai.model}:generateContent"
    body = json.dumps({"contents": [{"parts": [{"text": prompt}]}],
                       "generationConfig": {"responseMimeType": "application/json", "temperature": 0.3}}
                      ).encode("utf-8")
    headers = {"Content-Type": "application/json", "x-goog-api-key": ai.key}
    last = ""
    for attempt in range(1, ai.retries + 1):
        wait = None
        try:
            req = urllib.request.Request(url, data=body, method="POST", headers=headers)
            with urllib.request.urlopen(req, timeout=ai.timeout) as resp:
                data = json.loads(resp.read().decode("utf-8"))
            return _parse_model_json(_candidate_text(data))
        except urllib.error.HTTPError as e:
            last = f"HTTP {e.code} {_http_error_message(e)}".strip()
            if e.code not in (408, 429, 500, 502, 503, 504):
                raise AIError(last)
            ra = e.headers.get("Retry-After") if e.headers else None
            wait = float(ra) if ra and ra.replace(".", "", 1).isdigit() else None
        except (urllib.error.URLError, OSError) as e:
            last = f"تعذّر الاتصال: {getattr(e, 'reason', e)}"
        except (ValueError, AIError) as e:
            last = str(e)
        if attempt < ai.retries:
            delay = min(wait if wait is not None else 2 ** attempt, 60) + random.uniform(0, 1)
            print(f"  {last} — إعادة المحاولة بعد {delay:.0f}ث ({attempt}/{ai.retries - 1})", flush=True)
            time.sleep(delay)
    raise AIError(last or "فشل الاتصال")


# أعمدة الجدول الرئيسي القابلة للملء: (المفتاح، اسمها، رقمها في الصف)
AI_FIELDS = (("eval", "التقويم والمؤشرات", 0), ("content", "المحتوى", 1),
             ("activities", "أنشطة المدرس", 2), ("abil", "القدرات", 3))
# صفوف الجدول الرئيسي بعد حذف صف «التمهيد»: (رقم الصف، اسم المرحلة، ما يُنتظر في عمود المحتوى)
AI_ROWS = (
    (2, "تشخيص المكتسبات", "الإجابات المتوقعة للمتعلمين عن أسئلة التشخيص"),
    (3, "ملاحظة النص", "ما يُتوقع أن يلاحظه المتعلمون في مؤشرات النص الخارجية (العنوان، الصاحب، المصدر…) وفرضيات القراءة"),
    (4, "فهم النص", "خلاصة موجزة لمعنى النص وأفكاره الكبرى"),
    (7, "القراءة التحليلية", "عناصر التحليل المتوقعة (القيم، البنية، الأسلوب…) بإيجاز"),
    (8, "التركيب", "خلاصة تركيبية موجزة للدرس"),
    (9, "الاستثمار", "تمرين أو وضعية استثمار بسيطة مع نموذج جواب"),
)
AI_MAX_LINES, AI_MAX_CHARS = 6, 240


def tc_plain(tc):
    return "\n".join("".join(t.text or "" for t in p.iter(W + "t")) for p in tc.findall(W + "p")).strip()


def is_blank(txt):
    return not re.sub(r"[\s+\-–•*.:：،]+", "", txt)


def cell_state(tc):
    """'empty': لا شيء فيها | 'heading': عنوان مسطّر فقط بلا متن (مثل «التركيب:») | 'ok': فيها متن."""
    if is_blank(tc_plain(tc)):
        return "empty"
    body = []
    for r in tc.iter(W + "r"):
        rpr = r.find(W + "rPr")
        u = rpr.find(W + "u") if rpr is not None else None
        if u is not None and u.get(qn("w:val")) not in (None, "none"):
            continue                                    # نص مسطّر = عنوان
        body.append("".join(t.text or "" for t in r.iter(W + "t")))
    return "heading" if is_blank("".join(body)) else "ok"


def put_specs(cell, specs, mode):
    """mode='replace' يستبدل محتوى الخانة الفارغة، و'append' يُبقي العنوان ويضيف المتن بعده. يعيد الفقرات الجديدة."""
    tc = cell._tc
    if mode == "replace":
        for p in tc.findall(qn("w:p")):
            tc.remove(p)
    new = [mk_par(sp, "c") for sp in specs]
    for p in new:
        tc.append(p)
    return new


def ai_lines(val):
    """يقبل نصاً أو قائمة، وينظّف الأسطر ويحدّ عددها وطولها."""
    items = [val] if isinstance(val, str) else [str(x) for x in val] if isinstance(val, list) else []
    out = []
    for it in items:
        for ln in it.splitlines():
            ln = clean(ln)
            if strip_marker(ln):
                out.append(ln[:AI_MAX_CHARS])
    return out[:AI_MAX_LINES]


def ai_format(field, lines):
    """يحوّل أسطر الذكاء الاصطناعي إلى بنود بتنسيق الجذاذة (+ / - / تسطير)."""
    if field == "eval":
        return eval_lines(lines)
    if field == "activities":
        return act_lines(lines, Report(), "ذكاء اصطناعي")
    return [D(x) for x in lines]            # المحتوى والقدرات: بنود فرعية "- "


def ai_fill_blanks(MT, H0, cells_of, comps, wanted, topic, level, component, ai):
    """يملأ الخانات الفارغة فقط (لا يلمس أي خانة فيها نص) ويكمل أهداف الكفايات الناقصة."""
    jobs = []
    for r, label, hint in AI_ROWS:
        try:
            cells = cells_of(r)
        except IndexError:
            continue
        if len(cells) < 4:
            continue
        texts = {k: tc_plain(cells[i]._tc) for k, _, i in AI_FIELDS}
        states = {k: cell_state(cells[i]._tc) for k, _, i in AI_FIELDS}
        todo = [k for k, st in states.items() if st != "ok"]
        if todo:
            jobs.append({"row": r, "stage": label, "content_hint": hint, "fill": todo,
                         "heading_only": [k for k in todo if states[k] == "heading"],
                         "existing": {k: texts[k][:900] for k in texts if states[k] != "empty"},
                         "_states": states})
    need_obj = [w for w in wanted if w not in comps]
    added = []                                  # الفقرات التي أضافها الذكاء الاصطناعي (لتقليصها عند ضيق الصفحة)
    if not jobs and not need_obj:
        print("الذكاء الاصطناعي: لا توجد فراغات تحتاج ملئاً.")
        return added

    payload = {"lesson": {"topic": topic, "level": level, "component": component},
               "rows": [{k: v for k, v in j.items() if not k.startswith("_")} for j in jobs],
               "objectives_needed": need_obj, "source_objectives": comps}
    prompt = (
        "أنت مفتش تربوي خبير بالجذاذات المغربية لمادة اللغة العربية (مكوّن القراءة) في التعليم الثانوي الإعدادي.\n"
        f"المطلوب: ملء الخانات الفارغة فقط في جذاذة درس القراءة «{topic}» ({level}).\n\n"
        "القواعد الصارمة:\n"
        "1. أعد JSON فقط، وفق الصيغة المبيَّنة أدناه.\n"
        "2. لكل صف في rows املأ الحقول المذكورة في fill وحدها، ولا تُرجع غيرها ولا تُعد كتابة existing.\n"
        "   الحقول المذكورة في heading_only فيها عنوان مسطّر موجود (انظر existing) بلا متن: اكتب متنها فقط دون إعادة العنوان.\n"
        "3. استعن بـ existing وبعنوان الدرس لتبقى الخانات المملوءة منسجمة مع باقي الجذاذة.\n"
        "4. eval (التقويم والمؤشرات): سؤال أو مؤشر تقويمي مناسب للمرحلة.\n"
        "5. content (المحتوى): اتبع content_hint لكل صف.\n"
        "6. activities (أنشطة المدرس): أفعال بصيغة الغائب تبدأ بـ «ي» (يطرح، يوجّه، يناقش…)، "
        "وكل تفصيل فرعي يبدأ بـ «- ».\n"
        "7. abil (القدرات): قدرة أو قدرتان مناسبتان، كل واحدة في سطر مستقل.\n"
        f"8. كن موجزاً جداً (من سطرين إلى {AI_MAX_LINES} أسطر للحقل، والسطر لا يتجاوز نحو 25 كلمة) "
        "لأن الجذاذة كلها محصورة في صفحتين؛ "
        "لا تضع ترقيماً ولا رموز تنسيق ولا Markdown.\n"
        "9. لا تنسب إلى النص أقوالاً حرفية أو معلومات لم ترد في المدخلات؛ "
        "وإن لم تكفِ المدخلات فاكتب صياغة تربوية عامة تناسب المرحلة.\n"
        "10. objectives: اكتب فقط الكفايات المذكورة في objectives_needed (هدف موجز واحد لكل كفاية، بالمفتاح نفسه)، "
        "وإلا فاتركه كائناً فارغاً.\n"
        'الصيغة: {"rows":[{"row":2,"eval":[],"content":[],"activities":[],"abil":[]}],'
        '"objectives":{"تواصلية":"","منهجية":"","ثقافية":""}}\n\n'
        "المدخلات:\n" + json.dumps(payload, ensure_ascii=False)
    )
    print(f"جاري ملء الفراغات بـ {ai.model}…", flush=True)
    try:
        result = call_gemini(prompt, ai)
    except AIError as e:
        print(f"⚠ تعذّر ملء الفراغات بالذكاء الاصطناعي ({e}) — بقيت الخانات الفارغة كما هي.")
        return added

    by_row = {}
    for item in result["rows"]:
        if isinstance(item, dict):
            try:
                by_row[int(item.get("row"))] = item
            except (TypeError, ValueError):
                pass
    names = {k: n for k, n, _ in AI_FIELDS}
    col = {k: i for k, _, i in AI_FIELDS}
    filled = []
    for job in jobs:
        item = by_row.get(job["row"])
        if item is None:
            print(f"⚠ لم يُرجِع الذكاء الاصطناعي الصف «{job['stage']}»؛ بقي فارغاً.")
            continue
        cells = cells_of(job["row"])
        for k in job["fill"]:                       # الحقول المطلوبة وحدها
            lines = ai_lines(item.get(k))
            if lines:
                mode = "append" if job["_states"][k] == "heading" else "replace"
                added += put_specs(cells[col[k]], ai_format(k, lines), mode)
                filled.append(f"{job['stage']}: {names[k]}" + (" (بعد العنوان)" if mode == "append" else ""))

    obj = result.get("objectives") if isinstance(result.get("objectives"), dict) else {}
    pres = {"تواصلية": "", "منهجية": "- ", "ثقافية": "- "}
    rows_h = {"تواصلية": 3, "منهجية": 4, "ثقافية": 5}
    for w in need_obj:
        text = " ".join(ai_lines(obj.get(w)))
        if text:
            set_cell_keep_format(tcell(H0, rows_h[w], 0), [f"{pres[w]}كفاية {w}: {text}"])
            filled.append(f"الأهداف: كفاية {w}")
    if filled:
        print(f"✔ الذكاء الاصطناعي ملأ {len(filled)} خانة (راجعها قبل الاعتماد):")
        for f in filled:
            print("   •", f)
    else:
        print("• الذكاء الاصطناعي لم يُرجع ما يصلح للملء.")
    return added


# ═══════════════════════════ ضبط الصفحات (الحدّ الأقصى صفحتان) ═══════════════════════════
# تقدير ارتفاع الصفحة دون Word: الثوابت أدناه معايَرة تقريباً على خط Sakkal Majalla الغامق 9pt.
CHAR_EM = 0.28          # متوسط عرض الحرف العربي (نسبة إلى حجم الخط)
LINE_EM = 1.25          # ارتفاع السطر (نسبة إلى حجم الخط)
CELL_PAD_PT = 11.0      # الهامش الأفقي الداخلي للخانة (يمين + يسار)
DEFAULT_SZ = 22         # حجم الخط الافتراضي بنصف النقطة إن لم يُحدَّد
FIT_SAFETY = 0.98       # هامش أمان من ارتفاع الصفحة المتاح
FONT_START, FONT_SOFT_FLOOR, FONT_HARD_FLOOR = 18, 16, 14   # 9pt ← 8pt ← 7pt


def _int(el, attr, default=0):
    try:
        return int(el.get(qn(attr)))
    except (TypeError, ValueError, AttributeError):
        return default


def par_size(p):
    sizes = []
    for r in p.iter(W + "r"):
        if not "".join(t.text or "" for t in r.iter(W + "t")).strip():
            continue
        rpr = r.find(W + "rPr")
        sz = rpr.find(W + "sz") if rpr is not None else None
        if sz is not None:
            sizes.append(_int(sz, "w:val", DEFAULT_SZ))
    if sizes:
        return max(sizes)
    ppr = p.find(W + "pPr")
    mark = ppr.find(W + "rPr") if ppr is not None else None
    sz = mark.find(W + "sz") if mark is not None else None
    return _int(sz, "w:val", DEFAULT_SZ) if sz is not None else DEFAULT_SZ


def par_height(p, width_pt):
    chars = len("".join(t.text or "" for t in p.iter(W + "t")))
    sz = par_size(p) / 2.0
    avail = max(width_pt - CELL_PAD_PT, 20.0)
    lines = max(1, math.ceil(chars * CHAR_EM * sz * 1.04 / avail))
    line_h, before, after = sz * LINE_EM, 0.0, 0.0
    ppr = p.find(W + "pPr")
    sp = ppr.find(W + "spacing") if ppr is not None else None
    if sp is not None:
        before, after = _int(sp, "w:before") / 20.0, _int(sp, "w:after") / 20.0
        line, rule = _int(sp, "w:line"), sp.get(qn("w:lineRule"), "auto")
        if line:
            line_h = sz * LINE_EM * line / 240.0 if rule == "auto" else (line / 20.0 if rule == "exact" else max(line_h, line / 20.0))
    return lines * line_h + before + after


def grid_widths(tbl):
    g = tbl.find(qn("w:tblGrid"))
    return [_int(c, "w:w") / 20.0 for c in g.findall(qn("w:gridCol"))] if g is not None else []


def cell_heights(tr, grid):
    """[(tc, ارتفاعه التقديري بالنقطة)] لكل خانة في الصف (تُهمل الخانات التابعة لدمج رأسي)."""
    out, col = [], 0
    for tc in tr.findall(qn("w:tc")):
        pr = tc.find(qn("w:tcPr"))
        span, w, cont = 1, 0.0, False
        if pr is not None:
            gs = pr.find(qn("w:gridSpan"))
            span = _int(gs, "w:val", 1) if gs is not None else 1
            tw = pr.find(qn("w:tcW"))
            if tw is not None and tw.get(qn("w:type"), "dxa") == "dxa":
                w = _int(tw, "w:w") / 20.0
            vm = pr.find(qn("w:vMerge"))
            cont = vm is not None and vm.get(qn("w:val"), "continue") == "continue"
        if not w:
            w = sum(grid[col:col + span]) if grid else 0.0
        col += span
        if cont:
            continue
        h = sum(par_height(p, w or 100.0) for p in tc.findall(qn("w:p")))
        h += sum(table_height(t) for t in tc.findall(qn("w:tbl")))        # جداول متداخلة داخل الخانة
        out.append((tc, h))
    return out


NESTED_SAFETY = 1.35                        # أعمدة الجداول المتداخلة ضيقة فيلتفّ نصها أكثر مما يُقدَّر


def table_height(tbl):
    g = grid_widths(tbl)
    return sum(row_height(tr, g) for tr in tbl.findall(qn("w:tr"))) * NESTED_SAFETY + 6.0


def row_height(tr, grid):
    best = max((h for _, h in cell_heights(tr, grid)), default=0.0)
    trpr = tr.find(qn("w:trPr"))
    th = trpr.find(qn("w:trHeight")) if trpr is not None else None
    if th is not None:
        v = _int(th, "w:val") / 20.0
        best = v if th.get(qn("w:hRule")) == "exact" else max(best, v)
    return best + 3.0


def shrink_to(trs, sz):
    for tr in trs:
        for tag in ("w:sz", "w:szCs"):
            for el in tr.iter(qn(tag)):
                if _int(el, "w:val") > sz:
                    el.set(qn("w:val"), str(sz))


def trim_ai(trs, grid, added, removed):
    """يحذف آخر فقرة أضافها الذكاء الاصطناعي من الخانة التي تحدّد ارتفاع أطول صفّ (وتبقى فقرة واحدة على الأقل)."""
    for tr in sorted(trs, key=lambda t: -row_height(t, grid)):
        cells = cell_heights(tr, grid)
        if not cells:
            continue
        tc = max(cells, key=lambda c: c[1])[0]
        if len(tc.findall(qn("w:p"))) < 2:
            continue
        for p in reversed(added):
            if p.getparent() is tc:
                tc.remove(p)
                added.remove(p)
                removed.append((tc, p))
                return True
    return False


def fit_group(trs, grid, extra, limit, added):
    """يُدخل صفوف صفحة واحدة في المساحة المتاحة: تصغير الخط إلى 8pt ← حذف أسطر الذكاء الاصطناعي ← 7pt."""
    total = lambda: extra + sum(row_height(t, grid) for t in trs)
    if total() <= limit:
        return {"fit": True, "sz": FONT_START, "trimmed": 0, "est": total()}
    sz, removed = FONT_START, []
    while total() > limit and sz > FONT_SOFT_FLOOR:
        sz -= 1
        shrink_to(trs, sz)
    while total() > limit and trim_ai(trs, grid, added, removed):
        pass
    while total() > limit and sz > FONT_HARD_FLOOR:
        sz -= 1
        shrink_to(trs, sz)
    ok = total() <= limit
    if not ok:                                   # الحذف لم يُجدِ: أعد أسطر الذكاء الاصطناعي كما كانت
        for tc, p in reversed(removed):
            tc.append(p)
            added.append(p)
        removed = []
    return {"fit": ok, "sz": sz, "trimmed": len(removed), "est": total()}


def fit_pages(doc, H0, MT, added):
    """الصفحة 1 = الترويسة + الصفوف 0–4، والصفحة 2 = الصفوف 5–9 (يفصلهما pageBreakBefore)."""
    sec = doc.sections[0]
    page_h = (sec.page_height or 0) / 12700.0 or 842.0          # EMU → نقطة (A4 إن لم يُحدَّد)
    usable = page_h - (sec.top_margin or 0) / 12700.0 - (sec.bottom_margin or 0) / 12700.0
    limit = usable * FIT_SAFETY
    g0, g1 = grid_widths(H0._tbl), grid_widths(MT._tbl)
    head = sum(row_height(t, g0) for t in H0._tbl.findall(qn("w:tr")))
    trs = MT._tbl.findall(qn("w:tr"))
    for name, group, extra in (("الأولى", trs[:5], head + 14.0), ("الثانية", trs[5:], 6.0)):
        r = fit_group(group, g1, extra, limit, added)
        if r["sz"] != FONT_START or r["trimmed"]:
            print(f"ضبط الصفحة {name}: الخط {r['sz'] / 2:g}pt" +
                  (f"، وحُذف {r['trimmed']} سطر/أسطر من إضافات الذكاء الاصطناعي" if r["trimmed"] else ""))
        if not r["fit"]:
            print(f"⚠ الصفحة {name} قد تتجاوز المساحة رغم التصغير؛ راجعها في Word واختصر المصدر إن لزم.")


# ═══════════════════════════ بناء الجذاذة ═══════════════════════════
def build(model_path, source_path, out_name, ai=None, out_root=Path(OUT_DIR), fit=True):
    src = Source(source_path)
    rep = Report()
    doc = open_docx(model_path, "الجذاذة النموذج")
    if len(doc.tables) != 2 or len(doc.tables[1].rows) != 11:
        raise ShapeError("النموذج المتوقع بتخطيط «المنفى»")
    H0, MT = doc.tables[0], doc.tables[1]

    info = src.box("info").lines
    def info_value(prefix):
        for ln in info:
            if ln.startswith(prefix): return ln.split(":", 1)[1].strip()
        return ""

    level = info_value("الفئة المستهدفة")
    unit_lines = [l for l in info if l.startswith("الوحدة")]
    unit_line = unit_lines[0] if unit_lines else ""
    m = re.match(r"الوحدة\s+(\S+)\s*:", unit_line)
    unit_val = m.group(1) if m else ""
    domain_val = re.sub(r"^المجال\s+", "", unit_line.split(":", 1)[1].strip()) if ":" in unit_line else ""
    component = info_value("المكون")
    topic = re.sub(r"\s*/\s*", " / ", info_value("الموضوع"))
    topic = re.sub(r"\s+ص\s*(\d+)\s*$", r" – ص: \1", topic)
    
    di = next((i for i, l in enumerate(info) if l.startswith("الحصيص الزمني")), -1)
    duration = re.sub(r"\s+", " ", (info_value("الحصيص الزمني") + " " + (info[di + 1] if di != -1 and di+1 < len(info) else "")).strip()).replace(" )", ")")
    
    mi = next((i for i, l in enumerate(info) if l.startswith("الوسائل")), -1)
    if mi != -1:
        mat_raw = info[mi].split(":", 1)[1] if ":" in info[mi] else ""
        if mi + 1 < len(info): mat_raw += " " + info[mi + 1]
        materials = " – ".join(x.strip() for x in re.split(r"\s*\+\s*", mat_raw) if x.strip())
        materials = re.sub(r"(\S)\(", r"\1 (", materials)
    else:
        materials = ""

    set_cell_keep_format(tcell(H0, 0, 0), [level])
    set_cell_keep_format(tcell(H0, 0, 4), [unit_val])
    set_cell_keep_format(tcell(H0, 1, 0), [materials])
    set_cell_keep_format(tcell(H0, 1, 4), [domain_val])
    set_cell_keep_format(tcell(H0, 2, 2), [component])
    set_cell_keep_format(tcell(H0, 3, 2), [topic])
    set_cell_keep_format(tcell(H0, 4, 2), [duration])

    comps = {}
    for ln in src.box("comps").lines:
        mm = re.match(r"^\*?\s*(?:ال)?كفاية\s+(\S+)\s*:\s*(.+)$", ln)
        if mm: comps[re.sub(r"^ال", "", mm.group(1))] = mm.group(2).strip()
    wanted = ["تواصلية", "منهجية", "ثقافية"]
    for r, w, pre in zip((3, 4, 5), wanted, ("", "- ", "- ")):
        if w in comps: set_cell_keep_format(tcell(H0, r, 0), [f"{pre}كفاية {w}: {comps[w]}"])

    rows = MT.rows
    def cells_of(r): return [_Cell(tc, MT) for tc in rows[r]._tr.findall(qn("w:tc"))]

    def clear_row(ri, title):
        """يمسح محتوى صف القالب عند غياب مرحلته في المصدر، بدل إبقاء مثال قديم."""
        c = cells_of(ri)
        for cell in c[:4]:
            fill(cell, [""], "c")
        fill(c[4], [title], "c")

    fill(cells_of(0)[0], [src.box("s1").lines[0]], "c")
    fill(cells_of(5)[0], [src.box("s2").lines[0]], "c")
    tmhid = rows[7]._tr
    MT._tbl.remove(tmhid)

    groups = []
    for r in range(2, 10):
        k = key(" ".join(src.items(r, "step")))
        if groups and groups[-1][0] == k: groups[-1][1].append(r)
        else: groups.append((k, [r]))
    
    g_diag = groups[0][1] if len(groups) > 0 else []
    g_obs = groups[1][1] if len(groups) > 1 else []
    g_und = groups[2][1] if len(groups) > 2 else []

    # تشخيص
    if g_diag:
        r = g_diag[0]
        stage = " ".join(src.items(r, "stage"))
        ind = src.items(r, "ind")
        c0 = [f"+ {stage}: {end(ind[0])}" if ind else f"+ {stage}."]
        c0 += [end(it) for it in src.items(r, "eval")]
        content = src.items(r, "content")
        note = [b for b in src.boxes if b.kind == "note"]
        if note and not any(note[0].lines[0] in c for c in content):
            pos = next((i for i, it in enumerate(content) if it.startswith("(")), len(content))
            content = content[:pos] + [note[0].lines[0]] + content[pos:]
        teacher = [re.sub(r"،?\s*ليطلع عليها\.?$", "", it) for it in src.items(r, "teacher")]
        for it in src.items(r, "learner"): rep.drop("تشخيص", strip_marker(it))
        c = cells_of(2)
        fill(c[0], c0, "c"); fill(c[1], content, "c"); fill(c[2], act_lines(teacher, rep, "تشخيص/المدرس"), "c")
        fill(c[3], [end(strip_marker(x)) for x in src.items(r, "abil")], "c"); fill(c[4], ["- تشخيص المكتسبات"], "c")
    else:
        clear_row(2, "- تشخيص المكتسبات")
        rep.notes.append("لا مرحلة تشخيص في المصدر؛ مُسح مثال القالب من صف التشخيص.")

    # ملاحظة
    if g_obs:
        c0, c1, c2, c3 = [], [], [], []
        cf = ContentFmt(num_style="item", number_headings=True)
        for n, r in enumerate(g_obs):
            if n: c0.append(""); c2.append(""); c3.append("")
            c0 += [T(x) for x in src.items(r, "eval") + src.items(r, "ind")]
            c1 += cf.fmt(src.items(r, "content"))
            c2.append(H(f"{n + 1}- {' '.join(src.items(r, 'stage'))}:"))
            c2 += act_lines(src.items(r, "teacher"), rep, "ملاحظة/المدرس")
            c3 += ["- " + end(strip_marker(x)) for x in src.items(r, "abil")]
        c = cells_of(3)
        fill(c[0], c0, "c"); fill(c[1], c1, "r"); fill(c[2], c2, "c"); fill(c[3], c3, "c"); fill(c[4], ["- ملاحظة النص"], "c")
    else:
        clear_row(3, "- ملاحظة النص")
        rep.notes.append("لا مرحلة ملاحظة في المصدر؛ مُسح مثال القالب من صف الملاحظة.")

    # فهم
    if g_und:
        seen_content, c1 = set(), []
        cf = ContentFmt(num_style="heading", number_headings=False)
        spec_done, read_spec = False, next((strip_marker(it) for r in g_und for it in src.items(r, "teacher") if re.match(r"^يقرأ الأستاذ النص", strip_marker(it))), None)
        for r in g_und:
            its = src.items(r, "content")
            fresh = [it for it in its if key(it) not in seen_content]
            seen_content.update(key(it) for it in its)
            lines = cf.fmt(fresh) if fresh else []
            if lines and not spec_done and read_spec:
                c1 += [lines[0], D(read_spec)] + lines[1:]
                spec_done = True
            else: c1 += lines
        c0 = eval_lines(uniq_items(src, g_und, "eval")) + [T(x) for x in uniq_items(src, g_und, "ind")]
        c3 = ["+ " + end(strip_marker(x)) for x in uniq_items(src, g_und, "abil")]
        blocks, c2 = [], []
        for r in g_und:
            stage = " ".join(src.items(r, "stage"))
            if blocks and key(blocks[-1][0]) == key(stage): blocks[-1][1].append(r)
            else: blocks.append((stage, [r]))
        for n, (stage, brs) in enumerate(blocks):
            body = act_lines(uniq_items(src, brs, "teacher"), rep, "فهم/المدرس", skip=r"^يقرأ الأستاذ النص")
            for it in uniq_items(src, brs, "learner"):
                if any(re.search(p, strip_marker(it)) for p in KEEP_LEARNER): body.append(D(strip_marker(it)))
                else: rep.drop("فهم/المتعلم", strip_marker(it))
            if n: c2.append("")
            c2.append(H(f"{n + 1}- {stage}" + (":" if body else "."))); c2 += body
        c = cells_of(4)
        fill(c[0], c0, "c"); fill(c[1], c1, "r"); fill(c[2], c2, "c"); fill(c[3], c3, "c"); fill(c[4], ["فهم النص"], "c")
    else:
        clear_row(4, "فهم النص")
        rep.notes.append("لا مرحلة فهم في المصدر؛ مُسح مثال القالب من صف الفهم.")

    # الحصة الثانية
    R = 10
    cont, ev, ind_, tch, ab = (src.items(R, x) for x in ("content", "eval", "ind", "teacher", "abil"))
    stg = src.items(R, "stage")
    
    # فك دمج القدرات والمهارات
    new_ab = []
    for a in ab:
        for s in re.split(r"(?=\b[ا]?لقدرة\b|\bقدرة\b)", a):
            if s.strip(): new_ab.append(s.strip())
    ab = new_ab

    # دالة ذكية للبحث تتجاهل العناوين المدمجة
    def get_idx(items, pat, start=0, exclude=None):
        for i in range(start, len(items)):
            if exclude and re.search(exclude, items[i]): continue
            if re.search(pat, items[i]): return i
        return len(items)

    ci_s = get_idx(cont, r"\bالتركيب\b|خلاصة تركيبية", exclude=r"التحليلية والتركيبية")
    if ci_s == len(cont): ci_s = get_idx(cont, r"نوعية النص")
    ci_i = get_idx(cont, r"الاستثمار", start=ci_s)

    ei_s = get_idx(ev, r"مهارة|إنتاج|كتابة")
    ei_i = get_idx(ev, r"دعم|حرية|امتداد", start=ei_s)
    ii_s = get_idx(ind_, r"تركيب|خلاصة")
    
    ti_s = get_idx(tch, r"تجميع|التركيب", exclude=r"التحليلية والتركيبية")
    ti_i = get_idx(tch, r"الاستثمار|بحث", start=ti_s)
    
    ai_s = get_idx(ab, r"تحويل|وضعية|التركيب", exclude=r"التحليلية والتركيبية")
    ai_i = get_idx(ab, r"توظيف|الاستثمار", start=ai_s)
    
    si_s = get_idx(stg, r"الربط|التركيب", exclude=r"التحليلية والتركيبية")
    si_i = get_idx(stg, r"استثمار", start=si_s)

    def clean_title(idx, default_title):
        if idx < len(stg) and len(stg[idx]) < 40: return stg[idx]
        return default_title

    stage_s = clean_title(si_s, "التركيب")
    stage_i = clean_title(si_i, "الاستثمار")

    # ───── القراءة التحليلية ─────
    analysis_t = tch[:ti_s]
    rasd = [t for t in analysis_t if re.match(r"^\W*يقوم برصد", t)]
    c0 = eval_lines(ev[:ei_s]) + [T(x) for x in ind_[:ii_s]] + ["+ " + end(re.sub(r"^يقوم برصد", "رصد", strip_marker(x))) for x in rasd]
    
    cf = ContentFmt(num_style="heading", number_headings=False)
    raw_analysis = cont[:ci_s]
    if raw_analysis and re.search(r"التحليلية والتركيبية", raw_analysis[0]):
        raw_analysis.pop(0) 
        
    # <-- التعديل هنا: إزالة التمحيص ونوعية النص من محتوى التحليل لتسليط الضوء على القيم -->
    raw_analysis = [x for x in raw_analysis if not re.search(r"تمحيص|نوعية النص", x)]
    c1 = cf.fmt(raw_analysis)

    ext_extra, acts = [], []
    for it in analysis_t:
        # <-- التعديل هنا: تجاهل أسطر التمحيص في أنشطة المدرس -->
        if it in rasd or re.search(r"تمحيص|نوعية النص", it): continue
        base = strip_marker(it)
        mm = re.match(r"^يساعد المتعلمين على استخراج (.+?)\.?$", base)
        if mm and not base.endswith(":"): ext_extra.append(mm.group(1))
        else: acts += act_lines([it], rep, "تحليل/المدرس")
        
    if ext_extra:
        first_lead = next((i for i, a in enumerate(acts) if a.startswith("+ ") and a.endswith(":")), None)
        if first_lead is not None:
            j = first_lead + 1
            while j < len(acts) and acts[j].startswith("- "): j += 1
            for e in reversed(ext_extra): acts.insert(j, D(e))
        else: acts += [D(e) for e in ext_extra]
    acts = [("- معجم النص: " + a[2:]) if (a.startswith("- ") and "الحقول المعجمية" in a and "معجم النص" not in a) else a for a in acts]
    
    c = cells_of(7)
    fill(c[0], c0, "c"); fill(c[1], c1, "r"); fill(c[2], acts, "c"); fill(c[3], ["- " + end(strip_marker(x)) for x in ab[:ai_s]], "c"); fill(c[4], ["القراءة التحليلية"], "c")

    # ───── التركيب ─────
    c = cells_of(8)
    fill(c[0], [T(x) for x in ev[ei_s:ei_i] + ind_[ii_s:]], "c")
    fill(c[1], ContentFmt().fmt(cont[ci_s:ci_i]), "r")
    fill(c[2], [H(stage_s + ":")] + [T(x) for x in tch[ti_s:ti_i]], "c")
    fill(c[3], ["- " + end(strip_marker(x)) for x in ab[ai_s:ai_i]], "c")
    fill(c[4], ["التركيب"], "c")
    # ───── الاستثمار ─────
    seg = [re.sub(r"^\{\s*(.*?)\s*\}$", r"(\1)", s) for s in ContentFmt().fmt(cont[ci_i:])]
    c = cells_of(9)
    fill(c[0], [T(x) for x in ev[ei_i:]], "c"); fill(c[1], seg, "r")
    fill(c[2], [H(stage_i + ":")] + [T(x) for x in tch[ti_i:]], "c")
    fill(c[3], ["- " + end(strip_marker(x)) for x in ab[ai_i:]], "c")
    fill(c[4], ["الاستثمار"], "c")

    # ───── الذكاء الاصطناعي: ملء الفراغات فقط ─────
    ai_added = []
    if ai:
        ai_added = ai_fill_blanks(MT, H0, cells_of, comps, wanted, topic, level, component, ai) or []

    for b in src.boxes:
        if b.kind == "extra": rep.notes.append("مربع نص لم يُنقل: " + " ".join(b.lines)[:70])
        
    # ──────── فصل الحصتين في صفحتين ────────
    tc_h2 = MT.rows[5].cells[0]._tc
    p_h2 = tc_h2.find(qn("w:p"))
    if p_h2 is not None:
        pPr = p_h2.find(qn("w:pPr"))
        if pPr is None:
            pPr = OxmlElement("w:pPr")
            p_h2.insert(0, pPr)
        pPr.append(OxmlElement("w:pageBreakBefore"))
    # ───────────────────────────────────────
    
    # ──────── تصغير الهوامش لتتسع الصفحة ────────
    for section in doc.sections:
        sectPr = section._sectPr
        pgMar = sectPr.find(qn("w:pgMar"))
        if pgMar is not None:
            pgMar.set(qn("w:top"), "720")
            pgMar.set(qn("w:bottom"), "720")
            pgMar.set(qn("w:left"), "720")
            pgMar.set(qn("w:right"), "720")
    # ──────────────────────────────────────────

    if fit:
        fit_pages(doc, H0, MT, ai_added)

    out_path = save_safely(doc, output_path(out_root, out_name, component, unit_val))
    print("تم الحفظ بنجاح في:", out_path)
    return rep


# ═══════════════════════════ البحث والتشغيل ═══════════════════════════
USED_OUTPUTS = set()                    # الملفات الناتجة في هذا التشغيل


def output_path(root, name, component, unit):
    """الجذاذات_المنجزة / المكون / الوحدة / اسم الملف (مثل سكريبت اللغة)."""
    def folder(v, default):
        v = re.sub(r'[<>:"/\\|?*\r\n\t]+', "", v or "").strip(" .")
        return v or default
    target = Path(root) / folder(component, "القراءة") / folder(unit, "بدون وحدة")
    target.mkdir(parents=True, exist_ok=True)
    path, stem, n = target / name, Path(name).stem, 2
    while path in USED_OUTPUTS:                 # جذاذتان بالاسم نفسه في هذا التشغيل: لا نكتب فوق الأولى
        path = target / f"{stem}_{n}{Path(name).suffix}"
        n += 1
    USED_OUTPUTS.add(path)
    return path


def save_safely(doc, path):
    """يحفظ الملف؛ وإن كان مفتوحاً في Word يحفظ بنسخة مرقّمة بدل أن ينهار."""
    try:
        doc.save(str(path))
        return path
    except PermissionError:
        for i in range(2, 100):
            alt = path.with_name(f"{path.stem} ({i}){path.suffix}")
            try:
                doc.save(str(alt))
                return alt
            except PermissionError:
                continue
        raise


def find_file(filename):
    p = Path(filename)
    if p.exists(): 
        return str(p)
    
    print(f"جاري البحث عن «{filename}» في المجلدات...")
    # البحث في جميع المجلدات المتفرعة
    for f in Path(".").rglob(p.name):
        if f.is_file():
            print(f"تم العثور على الملف داخل: {f.parent}")
            return str(f)
            
    raise ShapeError(f"لم أجد الملف «{filename}» في أي مجلد.")

def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("model", nargs="?", default=DEFAULT_MODEL)
    ap.add_argument("source", nargs="?", default=DEFAULT_SOURCE)
    ap.add_argument("out", nargs="?", default=DEFAULT_OUT)
    ap.add_argument("--no-ai", action="store_true", help="تعطيل ملء الفراغات بالذكاء الاصطناعي")
    ap.add_argument("--ai-model", default=DEFAULT_AI_MODEL, help="اسم نموذج Gemini")
    ap.add_argument("--no-fit", action="store_true", help="تعطيل ضبط الجذاذة في صفحتين")
    ap.add_argument("--api-key", default="", help="مفتاح Gemini (الأفضل متغير البيئة GEMINI_API_KEY)")
    a = ap.parse_args(argv)

    ai = None
    if not a.no_ai:
        key = resolve_api_key(a.api_key)
        if key:
            ai = AICfg(a.ai_model, key)
            print(f"الذكاء الاصطناعي: مفعّل ({a.ai_model})")
        else:
            print("الذكاء الاصطناعي: غير مفعّل — لا يوجد مفتاح GEMINI_API_KEY (راجع أعلى الملف).")
    
    # البحث التلقائي عن النموذج والمصدر
    real_model = find_file(a.model)
    real_source = find_file(a.source)
    
    # المسار النهائي يُحدَّد داخل build من المكون والوحدة في المصدر
    build(real_model, real_source, a.out, ai, Path(OUT_DIR), fit=not a.no_fit)
    return 0

if __name__ == "__main__":
    sys.exit(main())
