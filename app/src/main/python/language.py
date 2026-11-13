#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_jadhadha.py — الإصدار 2
=============================
ينقل محتوى جذاذة (المصدر) إلى نموذج جذاذة جاهز بتنسيق موحَّد، مع تنقيح بيداغوجي اختياري بالذكاء الاصطناعي.

الاستعمال
---------
    python build_jadhadha.py                                  # الأسماء الافتراضية في المجلد الحالي
    python build_jadhadha.py النموذج.docx جذاذة.docx
    python build_jadhadha.py النموذج.docx جذاذة1.docx جذاذة2.docx --no-ai     # دفعة من الجذاذات
    python build_jadhadha.py النموذج.docx جذاذة.docx --no-reflow               # دون دمج الفقرات المتتالية
    python build_jadhadha.py --help

مفتاح الذكاء الاصطناعي (اختياري)
--------------------------------
لا يُكتب المفتاح داخل السكريبت أبداً. ضعه في أحد الموضعين:
  • متغير بيئة:  GEMINI_API_KEY        (ويندوز:  setx GEMINI_API_KEY "مفتاحك")
  • ملف .env بجوار السكريبت أو في المجلد الحالي، فيه سطر:  GEMINI_API_KEY=مفتاحك
وبدون مفتاح يعمل السكريبت كاملاً دون ذكاء اصطناعي.

المتطلبات:  pip install python-docx

ما الجديد عن الإصدار السابق
---------------------------
• إصلاح فقدان النص: كانت خلايا الجداول المتداخلة تحتفظ بآخر فقرة فقط.
• إعادة تجميع الأسطر: المصدر يضع كل كلمة في فقرة مستقلة لتضييق الخلايا، فتُجمَّع في جمل سليمة،
  وتُحذف الأسطر المكررة (مثل «5 د» و«المقطع التمهيدي» حين تتكرر).
• الترقيم التلقائي (I. II. …) يُحوَّل إلى نص صريح، لأن أرقام القوائم تتغير عند النقل بين الملفين.
• الأهداف التعلمية في ترويسة النموذج تُحدَّث من المصدر (كانت تبقى نصاً قديماً من درس آخر).
• تحديد الجداول والأعمدة بالمضمون لا بالموضع؛ والأعمدة الاختيارية الناقصة لا توقف العمل.
• الصور والترقيم والروابط لا تُفسد الملف الناتج (تُنقل الصور، ويُحذف ما لا يمكن نقله مع تنبيه).
• ترتيب عناصر OOXML صحيح، فلا يشتكي Word من «محتوى غير مقروء».
• المفتاح خارج الكود + إعادة محاولة ذكية + التحقق من ردّ الذكاء الاصطناعي قبل اعتماده.
• تقرير ختامي، وفحص آلي يتأكد من أن نصوص المصدر وصلت كلها إلى الملف الناتج.
• فصل «التقويم التشخيصي» عن «ملاحظة الأمثلة ووصفها» في صفّين، وتوزيع القدرات والأنشطة والتقويم
  والمحتوى بينهما بحسب كتل المصدر (--no-split للإبقاء على صف واحد).
• حذف «أنشطة المتعلم» والإبقاء على أنشطة المدرس (--with-learner لاسترجاعها).
• معالجة عدة جذاذات دفعة واحدة، وتسمية الملف الناتج من عنوان الدرس.
"""
from __future__ import annotations

import argparse
import copy
import io
import json
import os
import random
import re
import sys
import time
import traceback
import unicodedata
import urllib.error
import urllib.request
from dataclasses import dataclass, field
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Twips

# ════════════════════════════════ الإعدادات ════════════════════════════════
DEFAULT_MODEL = "اسم الفاعل وعمله (صيغ المبالغة).docx"
DEFAULT_SOURCE = "جذاذة --(أسلوب الاختصاص)3اع- المختار.docx"
OUT_DIR = "الجذاذات_المنجزة"
DEFAULT_AI_MODEL = "gemini-3.5-flash"

LABEL_TEACHER = "أنشطة المدرس:"
LABEL_LEARNER = "أنشطة المتعلم:"

# ما يُنقل إلى ترويسة النموذج: (التسمية في النموذج ← المعلومة المستخرجة من المصدر)
HEADER_FIELDS = (
    ("الوحدة", "unit"),
    ("المجال", "domain"),
    ("المكون", "component"),
    ("الموضوع", "topic"),
    ("مدة الإنجاز", "duration"),
)


@dataclass
class Settings:
    font: str = "Sakkal Majalla"
    size_pt: float = 10.0
    bold: bool = True
    margin_twips: int = 400
    keep_colors: bool = False          # الاحتفاظ بألوان المصدر (التسطير والمائل يُحفظان دائماً)
    update_objectives: bool = True     # تحديث «الأهداف التعلمية» في الترويسة من المصدر
    split_diagnostic: bool = True      # فصل «التقويم التشخيصي» عن «ملاحظة الأمثلة ووصفها» في صفّين
    learner_activities: bool = True    # True: تبقى جمل المتعلم كلها (يُحذف العنوان فقط)
    activity_labels: bool = False      # عناوين «أنشطة المدرس:/المتعلم:» (لا تُكتب افتراضياً؛ --labels لإظهارها)
    header_rows: int = 2               # صفوف رأس جدول المصدر (يُكتشف العدد الصحيح تلقائياً)
    model_header_rows: int = 1         # صفوف رأس الجدول الرئيسي في النموذج
    nested_proto_index: int = 1        # أي جدول متداخل في النموذج يُؤخذ منه تنسيق الجداول
    use_ai: bool = True
    ai_model: str = DEFAULT_AI_MODEL
    api_key: str = ""
    ai_retries: int = 4
    ai_timeout: int = 120
    out_dir: Path = Path(OUT_DIR)
    out_name: str = ""
    verbose: bool = False
    used_outputs: set = field(default_factory=set)   # أسماء الملفات الناتجة في هذا التشغيل


class FatalError(Exception):
    """خطأ يُعرض للمستخدم برسالة واضحة دون تتبّع تقني."""


# ═════════════════════════════ ثوابت OOXML ═════════════════════════════
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
MC = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

# الترتيب الإلزامي لأبناء العناصر في مخطط OOXML (مخالفته تُظهر رسالة «محتوى غير مقروء» في Word)
PPR_ORDER = ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr",
             "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap",
             "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
             "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc",
             "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr",
             "sectPr", "pPrChange"]
RPR_ORDER = ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline",
             "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing",
             "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd", "fitText",
             "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath"]
TBLPR_ORDER = ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize", "tblStyleColBandSize",
               "tblW", "jc", "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout", "tblCellMar",
               "tblLook", "tblCaption", "tblDescription"]
TCPR_ORDER = ["cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd", "noWrap", "tcMar",
              "textDirection", "tcFitText", "vAlign", "hideMark"]


# ═══════════════════════════════ التقرير ═══════════════════════════════
class Report:
    """يجمع ملاحظات التنفيذ ويطبعها في ملخص واحد في النهاية."""
    ICONS = {"ok": "✔", "info": "•", "warn": "⚠", "error": "✖"}

    def __init__(self) -> None:
        self.items: list[tuple[str, str]] = []

    def ok(self, msg: str) -> None:
        self.items.append(("ok", msg))

    def info(self, msg: str) -> None:
        self.items.append(("info", msg))

    def warn(self, msg: str) -> None:
        self.items.append(("warn", msg))

    def error(self, msg: str) -> None:
        self.items.append(("error", msg))

    def show(self) -> None:
        print("─" * 48)
        for kind, msg in self.items:
            print(f"{self.ICONS[kind]} {msg}")


def say(msg: str) -> None:
    print(msg, flush=True)


# ═════════════════════════════ أدوات النصوص ═════════════════════════════
_MARKS = re.compile("[\u0640\u064B-\u065F\u0670]")   # التطويل + التشكيل


def clean_text(s: str) -> str:
    """تنظيف سطر: تطويل، مسافات زائدة، وضبط المسافات حول علامات الترقيم العربية."""
    s = s.replace("\u0640", "").replace("\u00a0", " ")
    s = re.sub(r"[ \t\r\f\v]+", " ", s).strip()
    s = re.sub(r"\s+([:،؛؟.])", r"\1", s)
    s = re.sub(r"(?<=[\u0621-\u064a])«", " «", s)
    s = re.sub(r'،(?=[^\s)\]»"”])', "، ", s)
    return s


def norm(s: str) -> str:
    """تطبيع للمقارنة: حذف التشكيل والتطويل، توحيد الألف والياء والتاء المربوطة."""
    s = _MARKS.sub("", s or "")
    s = re.sub("[أإآٱ]", "ا", s).replace("ى", "ي").replace("ة", "ه")
    s = re.sub("[\u200e\u200f\u061c]", "", s)
    return re.sub(r"\s+", " ", s).strip()


def compact(s: str) -> str:
    """صيغة مضغوطة للمقارنة: بلا مسافات ولا ترقيم (تتحمّل كلمات مقطّعة بمسافات)."""
    return re.sub(r"[^\w]+", "", norm(s))


def p_text(p) -> str:
    """نص الفقرة: tab ← فراغ، br ← سطر جديد؛ ويتجاهل مربعات النص المتداخلة والنص المحذوف."""
    out: list[str] = []

    def walk(node) -> None:
        for ch in node:
            tag = ch.tag
            if not isinstance(tag, str):
                continue
            if tag == W + "t":
                out.append(ch.text or "")
            elif tag == W + "tab":
                out.append(" ")
            elif tag in (W + "br", W + "cr"):
                out.append("\n")
            elif tag == W + "noBreakHyphen":
                out.append("-")
            elif tag in (W + "txbxContent", W + "delText", W + "instrText", W + "pPr", MC + "Fallback"):
                continue
            else:
                walk(ch)

    walk(p)
    return "".join(out)


def tc_text(tc) -> str:
    return " ".join(clean_text(p_text(p).replace("\n", " ")) for p in tc.findall(W + "p")).strip()


def iter_paragraphs(root):
    """كل الفقرات (ومنها فقرات مربعات النص) مع تجاهل النسخ الاحتياطية mc:Fallback."""
    for p in root.iter(W + "p"):
        if any(a.tag == MC + "Fallback" for a in p.iterancestors()):
            continue
        yield p


def all_texts(doc) -> list[str]:
    seen: set[str] = set()
    out: list[str] = []
    for p in iter_paragraphs(doc.element.body):
        t = clean_text(p_text(p).replace("\n", " "))
        if t and t not in seen:
            seen.add(t)
            out.append(t)
    return out


# ── إعادة تجميع الأسطر ──
_TERMINAL = (".", ":", "؟", "?", "!", "؛", ";", "…")
_BULLET = re.compile(r"^(?:[-–—•*+·▪●◦■□]|\(?\d{1,2}\s*[-–.)/]|\(?[ا-ي]\s*[-–.)]\s)")
_ORDINAL = re.compile(r"^(?:اولا|ثانيا|ثالثا|رابعا|خامسا|سادسا|سابعا|ثامنا|تاسعا|عاشرا)(?![ء-ي])")
_PUNCT_ONLY = re.compile(r"^[.،:؛؟!…]+$")
JOIN_FRAGMENTS = True      # False (عبر --no-reflow): كل فقرة غير فارغة سطر مستقل كما في المصدر


def _starts_item(t: str) -> bool:
    n = norm(t)
    return bool(_BULLET.match(n) or _ORDINAL.match(n))


def reflow_blocks(paragraphs) -> list[list[str]]:
    """يعيد تجميع فقرات المصدر المقطّعة قسراً (كلمة في كل فقرة) إلى أسطر منطقية، مقسَّمةً إلى كتل.
    • الفقرة الفارغة فاصل بين كتلتين (وهي ما يحدّد حدّ المرحلتين حين يضمّ الصفّ مرحلتين).
    • داخل الكتلة يُلصق السطر بما قبله، ما لم يبدأ ببند (-، 1-، أ.، أولا…) أو ينتهِ السابق بعلامة ختام.
    • علامة ترقيم منفردة («.») تُلصق بالسطر السابق."""
    blocks: list[list[str]] = []
    after_blank = True
    for raw in paragraphs:
        t = clean_text(raw)
        if not t:
            after_blank = True
            continue
        if _PUNCT_ONLY.match(t):
            if blocks:
                blocks[-1][-1] = blocks[-1][-1].rstrip() + t
            after_blank = False
            continue
        if after_blank or not blocks:
            blocks.append([t])
        elif not JOIN_FRAGMENTS or _starts_item(t) or blocks[-1][-1].rstrip().endswith(_TERMINAL):
            blocks[-1].append(t)
        else:
            blocks[-1][-1] = blocks[-1][-1].rstrip() + " " + t
        after_blank = False
    return blocks


def reflow(paragraphs) -> list[str]:
    return [ln for blk in reflow_blocks(paragraphs) for ln in blk]


def dedupe(lines: list[str]) -> list[str]:
    seen: set[str] = set()
    out: list[str] = []
    for ln in lines:
        k = compact(ln)
        if k and k in seen:
            continue
        seen.add(k)
        out.append(ln)
    return out


# ═════════════════ الترقيم التلقائي في المصدر ← نص صريح ═════════════════
def _roman(n: int) -> str:
    out = ""
    for v, s in ((1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"), (50, "L"),
                 (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")):
        while n >= v:
            out += s
            n -= v
    return out or str(n)


def _letters(n: int) -> str:
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s or str(n)


_AR_ALPHA = "أبتثجحخدذرزسشصضطظعغفقكلمنهوي"
_AR_ABJAD = ["أ", "ب", "ج", "د", "هـ", "و", "ز", "ح", "ط", "ي", "ك", "ل", "م", "ن",
             "س", "ع", "ف", "ص", "ق", "ر", "ش", "ت", "ث", "خ", "ذ", "ض", "ظ", "غ"]
_AR_INDIC = str.maketrans("0123456789", "٠١٢٣٤٥٦٧٨٩")


def format_number(n: int, fmt: str) -> str:
    if fmt == "decimalZero":
        return f"{n:02d}"
    if fmt == "upperRoman":
        return _roman(n)
    if fmt == "lowerRoman":
        return _roman(n).lower()
    if fmt == "upperLetter":
        return _letters(n)
    if fmt == "lowerLetter":
        return _letters(n).lower()
    if fmt == "arabicAlpha":
        return _AR_ALPHA[(n - 1) % len(_AR_ALPHA)]
    if fmt == "arabicAbjad":
        return _AR_ABJAD[(n - 1) % len(_AR_ABJAD)]
    if fmt in ("hindiNumbers", "hindi1"):
        return str(n).translate(_AR_INDIC)
    return str(n)


def _val(parent, tag: str, default: str = "") -> str:
    e = parent.find(qn(tag))
    v = e.get(qn("w:val")) if e is not None else None
    return v if v is not None else default


class SourceNumbering:
    """يحسب نصوص الترقيم التلقائي لفقرات المصدر (I. / 1- / -) لتُكتب نصاً صريحاً في الناتج.
    السبب: معرّفات الترقيم (numId) تختلف بين الملفين، فيظهر الترقيم في الناتج بشكل مغاير للمصدر."""

    def __init__(self, doc) -> None:
        self.levels: dict[str, dict[int, tuple]] = {}
        self.num_abs: dict[str, str] = {}
        self.starts: dict[tuple, int] = {}
        self.labels: dict = {}
        try:
            root = doc.part.numbering_part.element
        except Exception:
            return
        for an in root.findall(qn("w:abstractNum")):
            lv: dict[int, tuple] = {}
            for lvl in an.findall(qn("w:lvl")):
                lv[int(lvl.get(qn("w:ilvl")))] = (_val(lvl, "w:numFmt", "decimal"),
                                                   _val(lvl, "w:lvlText", ""),
                                                   int(_val(lvl, "w:start", "1") or 1))
            self.levels[an.get(qn("w:abstractNumId"))] = lv
        for n in root.findall(qn("w:num")):
            nid = n.get(qn("w:numId"))
            ab = n.find(qn("w:abstractNumId"))
            if ab is not None:
                self.num_abs[nid] = ab.get(qn("w:val"))
            for ov in n.findall(qn("w:lvlOverride")):
                so = ov.find(qn("w:startOverride"))
                if so is not None:
                    self.starts[(nid, int(ov.get(qn("w:ilvl"))))] = int(so.get(qn("w:val")))

    def prepare(self, root) -> None:
        counters: dict[str, dict[int, int]] = {}
        for p in root.iter(W + "p"):
            ppr = p.find(W + "pPr")
            numpr = ppr.find(W + "numPr") if ppr is not None else None
            if numpr is None:
                continue
            nid = _val(numpr, "w:numId", "0")
            ilvl = int(_val(numpr, "w:ilvl", "0") or 0)
            levels = self.levels.get(self.num_abs.get(nid, ""), {})
            if nid == "0" or ilvl not in levels:
                continue
            cnt = counters.setdefault(nid, {})
            start = self.starts.get((nid, ilvl), levels[ilvl][2])
            cnt[ilvl] = cnt.get(ilvl, start - 1) + 1
            for deeper in [k for k in cnt if k > ilvl]:
                del cnt[deeper]
            fmt, txt, _ = levels[ilvl]
            if fmt == "none":
                continue
            if fmt == "bullet":
                ch = txt.strip()
                label = ch if ch and ch.isprintable() and ord(ch[0]) < 0xE000 else "-"
            else:
                def repl(m, levels=levels, cnt=cnt) -> str:
                    lv = int(m.group(1)) - 1
                    info = levels.get(lv, ("decimal", "", 1))
                    return format_number(cnt.get(lv, info[2]), info[0])
                label = re.sub(r"%(\d)", repl, txt)
            self.labels[p] = label


def para_text(p, numbering: SourceNumbering | None = None) -> str:
    t = p_text(p)
    if numbering is not None and t.strip():
        lab = numbering.labels.get(p)
        if lab:
            t = f"{lab} {t.lstrip()}"
    return t


def tc_blocks(tc, numbering: SourceNumbering | None = None) -> list[list[str]]:
    frags: list[str] = []
    for p in tc.findall(W + "p"):
        frags.extend(para_text(p, numbering).split("\n"))
    return reflow_blocks(frags)


def tc_lines(tc, numbering: SourceNumbering | None = None) -> list[str]:
    return [ln for blk in tc_blocks(tc, numbering) for ln in blk]


# ═══════════════════════════ بناء عناصر XML ═══════════════════════════
def el(tag: str, **attrs):
    e = OxmlElement(tag)
    for k, v in attrs.items():
        e.set(qn("w:" + k), str(v))
    return e


def put(parent, child, order: list[str]):
    """يضع child في موضعه الصحيح وفق ترتيب المخطط، ويستبدل أي عنصر سابق من النوع نفسه."""
    for old in parent.findall(child.tag):
        parent.remove(old)
    name = child.tag.split("}")[1]
    later = {W + n for n in order[order.index(name) + 1:]}
    for sib in parent:
        if sib.tag in later:
            sib.addprevious(child)
            return child
    parent.append(child)
    return child


def _flag_on(rpr, tag: str) -> bool:
    e = rpr.find(qn(tag))
    return e is not None and e.get(qn("w:val")) not in ("0", "false", "off")


_NEUTRAL = {"auto", "000000", "000", "0d0d0d", "262626", "404040"}


def make_rpr(cfg: Settings, old=None, underline: bool = False):
    """خصائص الخط الموحَّدة؛ مع الاحتفاظ بالمائل والتسطير (ولو بالألوان عند الطلب) من المصدر."""
    rpr = OxmlElement("w:rPr")
    put(rpr, el("w:rFonts", ascii=cfg.font, hAnsi=cfg.font, cs=cfg.font, eastAsia=cfg.font), RPR_ORDER)
    if cfg.bold:
        put(rpr, el("w:b"), RPR_ORDER)
        put(rpr, el("w:bCs"), RPR_ORDER)
    u_val = "single" if underline else None
    if old is not None:
        if _flag_on(old, "w:i"):
            put(rpr, el("w:i"), RPR_ORDER)
            put(rpr, el("w:iCs"), RPR_ORDER)
        u = old.find(qn("w:u"))
        if u is not None and u.get(qn("w:val")) != "none":
            u_val = u.get(qn("w:val")) or "single"
        if cfg.keep_colors:
            c = old.find(qn("w:color"))
            if c is not None and (c.get(qn("w:val")) or "auto").lower() not in _NEUTRAL:
                put(rpr, copy.deepcopy(c), RPR_ORDER)
    half = int(round(cfg.size_pt * 2))
    put(rpr, el("w:sz", val=half), RPR_ORDER)
    put(rpr, el("w:szCs", val=half), RPR_ORDER)
    if u_val:
        put(rpr, el("w:u", val=u_val), RPR_ORDER)
    put(rpr, el("w:rtl"), RPR_ORDER)
    put(rpr, el("w:lang", bidi="ar-SA"), RPR_ORDER)
    return rpr


def make_ppr(jc: str = "center"):
    ppr = OxmlElement("w:pPr")
    put(ppr, el("w:bidi"), PPR_ORDER)
    put(ppr, el("w:spacing", before=0, after=0, line=240, lineRule="auto"), PPR_ORDER)
    put(ppr, el("w:jc", val=jc), PPR_ORDER)
    return ppr


def make_run(text: str, cfg: Settings, underline: bool = False):
    r = OxmlElement("w:r")
    r.append(make_rpr(cfg, None, underline))
    t = OxmlElement("w:t")
    t.set(XML_SPACE, "preserve")
    t.text = re.sub(r"[\t\r\n]+", " ", text)
    r.append(t)
    return r


def make_p(text: str, cfg: Settings, underline: bool = False):
    p = OxmlElement("w:p")
    p.append(make_ppr())
    if text:
        p.append(make_run(text, cfg, underline))
    return p


def make_spacer(cfg: Settings):
    """فقرة فاصلة دقيقة (4pt): لازمة بعد الجدول المتداخل وبين جدولين متجاورين."""
    p = OxmlElement("w:p")
    ppr = make_ppr()
    mark = OxmlElement("w:rPr")
    put(mark, el("w:sz", val=4), RPR_ORDER)
    put(mark, el("w:szCs", val=4), RPR_ORDER)
    put(ppr, mark, PPR_ORDER)
    p.append(ppr)
    return p


def clear_cell(tc) -> None:
    for ch in list(tc):
        if ch.tag != W + "tcPr":
            tc.remove(ch)


def ensure_trailing_paragraph(tc, cfg: Settings) -> None:
    if len(tc) == 0 or tc[-1].tag != W + "p":
        tc.append(make_spacer(cfg))


def fill_cell(tc, lines: list[str], cfg: Settings) -> None:
    clear_cell(tc)
    for line in lines:
        line = line.strip()
        if line:
            tc.append(make_p(line, cfg, underline=line in (LABEL_TEACHER, LABEL_LEARNER)))
    ensure_trailing_paragraph(tc, cfg)


# ═══════════════════ نسخ المحتوى (فقرات وجداول متداخلة) ═══════════════════
@dataclass
class Ctx:
    cfg: Settings
    report: Report
    numbering: SourceNumbering
    src_part: object
    dst_doc: object
    nested_proto: object = None
    expected: list = field(default_factory=list)   # أسطر محتوى المصدر المتوقَّع ظهورها حرفياً
    next_docpr: int = 5000


_DROP = ("proofErr", "bookmarkStart", "bookmarkEnd", "commentRangeStart", "commentRangeEnd",
         "lastRenderedPageBreak", "permStart", "permEnd", "pPr", "del")
_UNWRAP = ("smartTag", "ins", "hyperlink")


def _detach(e) -> None:
    parent = e.getparent()
    if parent is not None:
        parent.remove(e)


def _unwrap(e) -> None:
    parent = e.getparent()
    if parent is None:
        return
    idx = parent.index(e)
    for k, ch in enumerate(list(e)):
        parent.insert(idx + k, ch)
    parent.remove(e)


def carry_media(p, ctx: Ctx) -> None:
    """ينقل الصور المضمّنة إلى حزمة النموذج؛ وما لا يمكن نقله يُحذف (وإلا صار الملف تالفاً في Word)."""
    for drawing in list(p.iter(W + "drawing")):
        ok = True
        for blip in drawing.iter(A + "blip"):
            try:
                blob = ctx.src_part.related_parts[blip.get(R + "embed")].blob
                new_rid, _ = ctx.dst_doc.part.get_or_add_image(io.BytesIO(blob))
                blip.set(R + "embed", new_rid)
                if blip.get(R + "link") is not None:
                    del blip.attrib[R + "link"]
            except Exception:
                ok = False
                break
        if ok:
            for dp in drawing.iter(WP + "docPr"):
                dp.set("id", str(ctx.next_docpr))
                ctx.next_docpr += 1
        else:
            _detach(drawing.getparent() if drawing.getparent().tag == W + "r" else drawing)
            ctx.report.warn("حُذفت صورة من المحتوى لتعذّر نقلها.")
    # أي علاقة أخرى (VML، كائنات، روابط) لا يمكن نقلها بين الملفين
    for e in list(p.iter()):
        if not isinstance(e.tag, str) or e.tag == A + "blip":
            continue
        if any(k.startswith(R) for k in e.attrib):
            run = next((a for a in e.iterancestors() if a.tag == W + "r"), None)
            if any(a.tag == W + "drawing" for a in e.iterancestors()) and e.tag in (A + "hlinkClick", A + "hlinkHover"):
                _detach(e)
            else:
                _detach(run if run is not None else e)
                ctx.report.warn("حُذف عنصر مرتبط بملف خارجي (صورة قديمة/كائن/رابط) لتعذّر نقله.")


def copy_paragraph(src_p, ctx: Ctx):
    """يستنسخ فقرة محتوى بعد تنظيفها وتوحيد خطها، مع الإبقاء على التسطير والمائل. يُرجع None للفقرة الفارغة."""
    cfg = ctx.cfg
    p = copy.deepcopy(src_p)
    label = ctx.numbering.labels.get(src_p)
    for tag in _DROP:
        for e in list(p.iter(W + tag)):
            _detach(e)
    for tag in _UNWRAP:
        for e in list(p.iter(W + tag)):
            _unwrap(e)
    for tag in ("commentReference", "footnoteReference", "endnoteReference"):
        for e in list(p.iter(W + tag)):
            _detach(e.getparent() if e.getparent().tag == W + "r" else e)
    for br in list(p.iter(W + "br")):
        if br.get(W + "type") in ("page", "column"):
            _detach(br)
    for tab in list(p.iter(W + "tab")):
        t = OxmlElement("w:t")
        t.text = " "
        tab.getparent().replace(tab, t)
    carry_media(p, ctx)

    texts = list(p.iter(W + "t"))
    prev = None
    for t in texts:
        if t.text:
            s = re.sub(r"[ \t]{2,}", " ", t.text.replace("\u0640", "").replace("\u00a0", " "))
            t.text = re.sub(r"\s+([:،؛؟.])", r"\1", s)               # «أولا :» ← «أولا:» كبقية الأعمدة
            if prev is not None and t.text[:1] in (":", "،", "؛", "؟", ".") and prev.text.endswith(" "):
                prev.text = prev.text.rstrip()                          # الفراغ في نهاية الجزء السابق
            prev = t
        t.set(XML_SPACE, "preserve")
    filled = [t for t in texts if (t.text or "").strip()]
    if not filled and next(p.iter(W + "drawing"), None) is None:
        return None
    if filled:                                  # لا فراغات في طرفَي الفقرة (حتى لو كانت في أجزاء مستقلة)
        first, last = filled[0], filled[-1]
        for t in texts:
            if t is first:
                break
            t.text = ""
        for t in reversed(texts):
            if t is last:
                break
            t.text = ""
        first.text = first.text.lstrip()
        last.text = last.text.rstrip()

    for r in p.iter(W + "r"):
        old = r.find(W + "rPr")
        if old is not None:
            r.remove(old)
        r.insert(0, make_rpr(cfg, old))
    p.insert(0, make_ppr())
    if label:
        p.insert(1, make_run(label + " ", cfg))
    return p


def table_spec(tbl, numbering: SourceNumbering | None) -> dict:
    """يقرأ جدولاً متداخلاً من المصدر: عرض الأعمدة، الدمج، وأسطر كل خلية."""
    g = tbl.find(W + "tblGrid")
    grid = [int(c.get(W + "w") or 0) for c in g] if g is not None else []
    rows = []
    for tr in tbl.findall(W + "tr"):
        row = []
        for tc in tr.findall(W + "tc"):
            pr = tc.find(W + "tcPr")
            span, vm = 1, None
            if pr is not None:
                gs = pr.find(W + "gridSpan")
                if gs is not None:
                    span = int(gs.get(W + "val") or 1)
                v = pr.find(W + "vMerge")
                if v is not None:
                    vm = v.get(W + "val") or "continue"
            row.append({"lines": tc_lines(tc, numbering), "span": span, "vm": vm})
        rows.append(row)
    ncols = sum(c["span"] for c in rows[0]) if rows else 0
    if len(grid) != ncols or sum(grid) == 0:
        grid = [1] * ncols
    return {"grid": grid, "rows": rows}


def make_table(proto, spec: dict, total: int, cfg: Settings):
    """يبني جدولاً متداخلاً بتنسيق نموذج الجدول في القالب (رأس الجدول وصفوفه) وبيانات المصدر."""
    tbl = copy.deepcopy(proto)
    pr = tbl.find(W + "tblPr")
    if pr is None:
        pr = OxmlElement("w:tblPr")
        tbl.insert(0, pr)
    for tag in ("tblpPr", "tblOverlap"):          # الجداول العائمة تتراكب داخل الخلايا
        for e in pr.findall(W + tag):
            pr.remove(e)
    put(pr, el("w:tblW", w=total, type="dxa"), TBLPR_ORDER)
    put(pr, el("w:bidiVisual"), TBLPR_ORDER)


    protos = proto.findall(W + "tr")
    if not protos or not spec["rows"]:
        return tbl
    hdr_tr = protos[0]
    body_tr = protos[1] if len(protos) > 1 else protos[0]
    for tr in tbl.findall(W + "tr"):
        tbl.remove(tr)

    g = spec["grid"] or [1]
    widths = [max(1, round(total * x / sum(g))) for x in g]
    widths[-1] += total - sum(widths)
    grid_el = tbl.find(W + "tblGrid")
    if grid_el is not None:
        for gc in list(grid_el):
            grid_el.remove(gc)
        for w_ in widths:
            grid_el.append(el("w:gridCol", w=w_))

    for ri, row in enumerate(spec["rows"]):
        is_hdr = ri == 0 and len(spec["rows"]) > 1
        tr = copy.deepcopy(hdr_tr if is_hdr else body_tr)
        pcells = tr.findall(W + "tc")
        for tc in pcells:
            tr.remove(tc)
        col = 0
        for ci, cell in enumerate(row):
            ptc = copy.deepcopy(pcells[min(ci, len(pcells) - 1)]) if pcells else OxmlElement("w:tc")
            span = cell["span"]
            tcpr = ptc.find(W + "tcPr")
            if tcpr is None:
                tcpr = OxmlElement("w:tcPr")
                ptc.insert(0, tcpr)
            put(tcpr, el("w:tcW", w=sum(widths[col:col + span]), type="dxa"), TCPR_ORDER)
            for tag in ("gridSpan", "vMerge"):
                for e in tcpr.findall(W + tag):
                    tcpr.remove(e)
            if span > 1:
                put(tcpr, el("w:gridSpan", val=span), TCPR_ORDER)
            if cell["vm"]:
                put(tcpr, el("w:vMerge", val="restart") if cell["vm"] == "restart" else el("w:vMerge"), TCPR_ORDER)
            for e in ptc.findall(W + "p") + ptc.findall(W + "tbl"):
                ptc.remove(e)
            for ln in (cell["lines"] or [""]):
                ptc.append(make_p(ln, cfg))           # كل سطر في فقرته (الإصدار السابق كان يُبقي آخر سطر فقط)
            tr.append(ptc)
            col += span
        tbl.append(tr)
    return tbl


def copy_table_fallback(tbl, ctx: Ctx):
    """عند غياب جدول متداخل نموذجي في القالب: استنساخ الجدول كما هو مع تنظيف فقراته."""
    new = copy.deepcopy(tbl)
    pr = new.find(W + "tblPr")
    if pr is not None:
        for tag in ("tblpPr", "tblOverlap"):
            for e in pr.findall(W + tag):
                pr.remove(e)
            put(pr, el("w:bidiVisual"), TBLPR_ORDER)

    for tc in new.iter(W + "tc"):
        for p in list(tc.findall(W + "p")):
            q = copy_paragraph(p, ctx)
            tc.replace(p, q if q is not None else make_spacer(ctx.cfg))
    return new


def cell_inner_width(tc, fallback: int = 7700) -> int:
    pr = tc.find(W + "tcPr")
    w = pr.find(W + "tcW") if pr is not None else None
    try:
        return max(1500, int(w.get(W + "w")) - 100)
    except (AttributeError, TypeError, ValueError):
        return fallback


def convert_content(elements: list, ctx: Ctx, width: int) -> list:
    """يحوّل عناصر محتوى المصدر (فقرات/جداول) إلى عناصر جاهزة للإدراج في خلية النموذج."""
    out: list = []
    prev_table = False
    for e in elements:
        if e.tag == W + "tbl":
            spec = table_spec(e, ctx.numbering)
            new = make_table(ctx.nested_proto, spec, width, ctx.cfg) if ctx.nested_proto is not None \
                else copy_table_fallback(e, ctx)
            if prev_table:
                out.append(make_spacer(ctx.cfg))     # جدولان متجاوران يندمجان في Word دون فقرة بينهما
            out.append(new)
            for row in spec["rows"]:
                for c in row:
                    ctx.expected.extend(c["lines"])
            prev_table = True
        elif e.tag == W + "p":
            p = copy_paragraph(e, ctx)
            if p is None:
                continue
            out.append(p)
            ctx.expected.append(para_text(e, ctx.numbering))
            prev_table = False
    return out


# ═══════════════════════════ تحليل النموذج (القالب) ═══════════════════════════
MODEL_COLS = (("stage", "المراحل"), ("capabilities", "القدرات"), ("activities", "الانشطه"),
              ("content", "المحتوي"), ("evaluation", "التقويم"))


@dataclass
class Model:
    doc: object
    header: object           # جدول الترويسة
    main: object             # الجدول الرئيسي
    cols: dict               # فهرس كل عمود (بترتيب خلايا XML)
    nested_proto: object     # جدول متداخل نموذجي يُستنسخ منه تنسيق الجداول


def analyze_model(doc, cfg: Settings) -> Model:
    """يحدّد جدولَي النموذج بمضمونهما (لا بترتيبهما)، وأعمدة الجدول الرئيسي من عناوينها."""
    tables = doc.element.body.findall(W + "tbl")
    header = main = None
    for t in tables:
        head = " ".join(compact(tc_text(tc)) for tr in t.findall(W + "tr")[:3] for tc in tr.findall(W + "tc"))
        if header is None and "الوحده" in head and "المجال" in head:
            header = t
        elif main is None and "المراحل" in head and "المحتوي" in head:
            main = t
    if header is None and tables:
        header = tables[0]
    if main is None and len(tables) > 1:
        main = tables[1]
    if header is None or main is None or header is main:
        raise FatalError("لم أتعرف على جدولَي النموذج (الترويسة والجدول الرئيسي).")
    trs = main.findall(W + "tr")
    if len(trs) <= cfg.model_header_rows:
        raise FatalError("الجدول الرئيسي في النموذج لا يحوي صفاً نموذجياً بعد صف العناوين.")
    cols: dict[str, int] = {}
    for i, tc in enumerate(trs[0].findall(W + "tc")):
        ct = compact(tc_text(tc))
        for key, kw in MODEL_COLS:
            if key not in cols and kw in ct:
                cols[key] = i
                break
    for i, (key, _) in enumerate(MODEL_COLS):
        cols.setdefault(key, i)
    if max(cols.values()) >= len(trs[cfg.model_header_rows].findall(W + "tc")):
        raise FatalError("خلايا الصف النموذجي في النموذج لا تطابق أعمدة الجدول الرئيسي.")
    nested = [t for t in main.iter(W + "tbl") if t is not main]
    proto = None
    if nested:
        pick = nested[cfg.nested_proto_index] if len(nested) > cfg.nested_proto_index else nested[0]
        proto = copy.deepcopy(pick)
    return Model(doc, header, main, cols, proto)


# ═══════════════════════════ تحليل المصدر ═══════════════════════════
COL_RULES = (   # (العمود، الكلمات الدالة عليه بترتيب الأولوية) — بعد التطبيع والضغط
    ("time", ("الحصيص", "الزمني")),
    ("eval", ("التقويم", "الدعم")),
    ("content", ("محتوي", "التعلمات", "المخرجات", "المضامين", "المضمون")),
    ("learner", ("المتعلم", "انجازالنشاط", "التلميذ")),
    ("teacher", ("المدرس", "الاستاذ", "المدخلات")),
    ("stage", ("المراحل", "المهارات")),
    ("segment", ("مقاطع", "المقطع")),
    ("cap", ("القدرات", "الكفايات", "الاهداف")),
)
OPTIONAL_COLS = {"time": "الحصيص الزمني", "eval": "التقويم والدعم", "stage": "المراحل",
                 "segment": "مقاطع الدرس", "cap": "القدرات"}
READ_COLS = ("time", "eval", "content", "learner", "teacher", "stage", "segment", "cap", "activities")


def build_grid(tbl) -> list[list]:
    """شبكة الجدول: لكل صف قائمة خلايا بعدد أعمدة الشبكة (الخلية المدموجة أفقياً تتكرر،
    والمدموجة رأسياً تشير إلى خلية الصف الذي فوقها)."""
    rows: list[list] = []
    for tr in tbl.findall(W + "tr"):
        trpr = tr.find(W + "trPr")
        gb = trpr.find(W + "gridBefore") if trpr is not None else None
        row: list = [None] * int(gb.get(W + "val") or 0) if gb is not None else []
        for tc in tr.findall(W + "tc"):
            pr = tc.find(W + "tcPr")
            span, vm = 1, None
            if pr is not None:
                gs = pr.find(W + "gridSpan")
                if gs is not None:
                    span = int(gs.get(W + "val") or 1)
                v = pr.find(W + "vMerge")
                if v is not None:
                    vm = v.get(W + "val") or "continue"
            for _ in range(span):
                if vm == "continue" and rows and len(row) < len(rows[-1]):
                    row.append(rows[-1][len(row)])
                else:
                    row.append(tc)
        rows.append(row)
    return rows


def detect_columns(grid: list[list], header_rows: int) -> dict[str, int]:
    ncols = max(len(r) for r in grid)
    heads: list[str] = []
    for c in range(ncols):
        parts, seen = [], set()
        for r in range(min(header_rows, len(grid))):
            tc = grid[r][c] if c < len(grid[r]) else None
            if tc is not None and tc not in seen:
                seen.add(tc)
                parts.append(compact(tc_text(tc)))
        heads.append(" ".join(parts))
    cols: dict[str, int] = {}
    for key, kws in COL_RULES:
        for kw in kws:
            hit = next((i for i, h in enumerate(heads) if kw in h and i not in cols.values()), None)
            if hit is not None:
                cols[key] = hit
                break
    if "teacher" not in cols and "learner" not in cols:        # عمود أنشطة واحد غير مقسوم
        hit = next((i for i, h in enumerate(heads) if "الانشطه" in h and i not in cols.values()), None)
        if hit is not None:
            cols["activities"] = hit
    elif "teacher" not in cols:                                 # «أنشطة التعليم والتعلم» بجانب عمود المتعلم
        hit = next((i for i, h in enumerate(heads) if ("الانشطه" in h or "التعليم" in h) and i not in cols.values()), None)
        if hit is not None:
            cols["teacher"] = hit
    return cols


def _row_looks_like_data(grid, ri: int, limit: int = 45) -> bool:
    """صف تبدو خلاياه جملاً طويلة (أكثر من limit حرفاً) فلا يصلح صف عناوين."""
    seen = []
    for tc in grid[ri]:
        if tc is None or tc in seen:
            continue
        seen.append(tc)
        if len(compact(tc_text(tc))) > limit:
            return True
    return False


def analyze_source(doc, cfg: Settings):
    """يختار جدول الجذاذة وعدد صفوف رأسه وأعمدته بالمضمون (الأوفر أعمدةً مكتشَفة)."""
    best = None
    for tbl in doc.element.body.findall(W + "tbl"):
        grid = build_grid(tbl)
        if len(grid) < 2:
            continue
        for h in dict.fromkeys((cfg.header_rows, 1, 2, 3)):
            if h >= len(grid):
                continue
            if h > 1 and _row_looks_like_data(grid, h - 1):    # صف «العناوين» الإضافي فيه جمل طويلة: هو بيانات لا عناوين
                continue
            cols = detect_columns(grid, h)
            if not ({"teacher", "learner", "activities"} & set(cols)) or not ({"content", "stage"} & set(cols)):
                continue
            score = (len(cols), len(grid), -h)            # عند التعادل: أقل صفوف رأس (فلا يُبتلع صف التمهيد)
            if best is None or score > best[0]:
                best = (score, grid, h, cols)
    if best is None:
        raise FatalError("لم أجد في المصدر جدول الجذاذة (أعمدة المراحل والأنشطة).")
    return best[1], best[2], best[3]


@dataclass
class Row:
    idx: int
    stage: list = field(default_factory=list)
    capabilities: list = field(default_factory=list)
    activities: list = field(default_factory=list)
    evaluation: list = field(default_factory=list)
    content: list = field(default_factory=list)       # عناصر XML (فقرات وجداول) تُنسخ كما هي
    diagnostic: list = field(default_factory=list)    # إجابات متوقعة يقترحها الذكاء الاصطناعي
    content_preview: str = ""
    orig_lines: list = field(default_factory=list)    # أسطر المصدر الأصلية للتحقق الآلي
    key: str = ""                                     # المرحلة القانونية: diag / obs / concl / apply

    @property
    def is_diagnostic(self) -> bool:
        return any("تشخيص" in norm(s) for s in self.stage)


def _has_visible(els: list) -> bool:
    for e in els:
        if e.tag == W + "tbl" or next(e.iter(W + "drawing"), None) is not None:
            return True
        if e.tag == W + "p" and clean_text(p_text(e).replace("\n", " ")):
            return True
    return False


def content_preview(els: list, numbering: SourceNumbering) -> str:
    parts = []
    for e in els:
        if e.tag == W + "p":
            t = clean_text(para_text(e, numbering).replace("\n", " "))
        elif e.tag == W + "tbl":
            t = " | ".join(clean_text(p_text(p)) for p in e.iter(W + "p"))
        else:
            continue
        if t:
            parts.append(t)
    return " ".join(parts)


def _x_intervals(tbl):
    """الامتداد الأفقي لكل خلية (من عرضها tcW): ({id(خلية): (من، إلى)}, [خلايا كل صف مع امتدادها]) أو None."""
    allmap, rows = {}, []
    for tr in tbl.findall(W + "tr"):
        x, items = 0, []
        trpr = tr.find(W + "trPr")
        wb = trpr.find(W + "wBefore") if trpr is not None else None
        if wb is not None:
            try:
                x += int(wb.get(W + "w") or 0)
            except ValueError:
                pass
        for tc in tr.findall(W + "tc"):
            pr = tc.find(W + "tcPr")
            w = pr.find(W + "tcW") if pr is not None else None
            try:
                val = int(w.get(W + "w")) if w is not None and (w.get(W + "type") or "dxa") == "dxa" else 0
            except (TypeError, ValueError):
                val = 0
            if val <= 0:
                return None
            allmap[id(tc)] = (x, x + val)
            items.append((tc, x, x + val))
            x += val
        rows.append(items)
    return allmap, rows


def _merge_cells(cells: list):
    """خلية افتراضية تجمع محتوى عدة خلايا متجاورة بترتيبها (فقرات وجداول)."""
    tc = OxmlElement("w:tc")
    for c in cells:
        for e in c:
            if e.tag in (W + "p", W + "tbl"):
                tc.append(copy.deepcopy(e))
    return tc


def _tables_of(tc) -> list:
    """الجداول المتداخلة المباشرة في خلية (تُهمَلها tc_blocks فننقلها إلى المحتوى)."""
    return [e for e in tc if e.tag == W + "tbl"] if tc is not None else []


def _flat(blocks: list) -> list[str]:
    return [ln for blk in blocks for ln in blk]


def _first_block_split(blocks: list) -> tuple[list[str], list[str]]:
    """(حصة التقويم التشخيصي، حصة ملاحظة الأمثلة): الكتلة الأولى للتشخيص والباقي للملاحظة.
    وإن وُجدت كتلة واحدة فهي للتشخيص إن ذكرت استحضار المكتسبات، وإلا فللملاحظة."""
    if len(blocks) >= 2:
        return list(blocks[0]), _flat(blocks[1:])
    only = _flat(blocks)
    if only and any(w in norm(" ".join(only)) for w in ("استحضار", "تشخيص")):
        return only, []
    return [], only


def split_content(elements: list, numbering: SourceNumbering, next_stage: str) -> tuple[list, list]:
    """يقسم عناصر المحتوى عند عنوان المرحلة الثانية («I. الملاحظة والوصف:»):
    ما قبله (إجابات المتعلمين المقوَّمة) للتشخيص، وما بعده (الأمثلة والجدول) للملاحظة."""
    words = (next_stage or "").split()
    stem = compact(words[0]) if words else ""

    def heading(e) -> bool:
        if e.tag != W + "p":
            return False
        t = clean_text(p_text(e).replace("\n", " "))
        return bool(t) and len(t) <= 60 and t.endswith(":")

    picks = (lambda e: bool(stem) and heading(e) and stem in compact(p_text(e)),
             lambda e: heading(e) and bool(numbering.labels.get(e)))
    for pick in picks:
        for i, e in enumerate(elements):
            if pick(e):
                return elements[:i], elements[i:]
    return [], elements


def _activities(teacher: list[str], learner: list[str], generic: list[str], cfg: Settings) -> list[str]:
    """أسطر المدرس ثم سطر فارغ ثم أسطر المتعلم، بلا أي عنوان (إلا بخيار --labels)."""
    keep_learner = cfg.learner_activities and bool(learner)
    acts: list[str] = []
    if teacher:
        if cfg.activity_labels:
            acts.append(LABEL_TEACHER)
        acts += teacher
    if keep_learner:
        if teacher:
            acts.append("")
        if cfg.activity_labels:
            acts.append(LABEL_LEARNER)
        acts += learner
    return acts or generic


def _make_row(stage, caps, teacher, learner, generic, ev, content, numbering, cfg: Settings) -> Row:
    learner = learner if cfg.learner_activities else []
    used = (teacher + learner) or generic
    return Row(idx=0, stage=stage, capabilities=caps, activities=_activities(teacher, learner, generic, cfg),
               evaluation=ev, content=content, content_preview=content_preview(content, numbering),
               orig_lines=stage + caps + used + ev)


def _is_combined_intro(stage: list[str]) -> bool:
    """هل يجمع الصف مرحلتين: «تقويم تشخيصي» أولاً ثم مرحلة أخرى (ملاحظة الظاهرة ووصفها)؟"""
    return (len(stage) >= 2 and "تشخيص" in norm(stage[0])
            and not any("تشخيص" in norm(s) for s in stage[1:]))


def extract_rows(grid: list[list], cols: dict, header_rows: int, numbering: SourceNumbering,
                 cfg: Settings, report: Report) -> list[Row]:
    exploded = explode_list_rows(grid, cols, header_rows, numbering, cfg, report)
    if exploded is not None:
        return exploded
    rows: list[Row] = []
    seen: set = set()
    dropped = splits = 0
    # جدول غير منتظم (صفوف بعدد أعمدة شبكة مختلف): نطابق الأعمدة بالموضع الأفقي لا بالرقم
    first_tc = next((c for c in grid[0] if c is not None), None)
    xinfo = None
    if first_tc is not None and len({len(r) for r in grid}) > 1:
        xinfo = _x_intervals(first_tc.getparent().getparent())
    colint = {}
    if xinfo:
        for key, idx in cols.items():
            hc = grid[header_rows - 1][idx] if idx < len(grid[header_rows - 1]) else None
            if hc is not None and id(hc) in xinfo[0]:
                colint[key] = xinfo[0][id(hc)]
        rep_note = "جدول المصدر غير منتظم الأعمدة؛ طابقتُ أعمدته بمواضعها الأفقية"
        report.info(rep_note)
    for ri in range(header_rows, len(grid)):
        cells = grid[ri]

        def take(key: str, ri=ri, cells=cells):
            """الخلية (أو الخلايا) الواقعة تحت عنوان العمود؛ عنوان ممتد على عمودين يجمع خلايا الصف تحته."""
            idx = cols.get(key)
            if idx is None:
                return None
            hrow = grid[header_rows - 1]
            hc = hrow[idx] if idx < len(hrow) else None
            if xinfo and key in colint:
                x0, x1 = colint[key]
                scored = [(min(b, x1) - max(a, x0), a, b, tc) for tc, a, b in xinfo[1][ri]]
                cand = [tc for ov, a, b, tc in scored if ov > 0 and ov >= 0.5 * (b - a)]
                if not cand:
                    best = max(scored, key=lambda t: t[0], default=None)
                    cand = [best[3]] if best and best[0] > 0 else []
            else:
                span = [j for j, c in enumerate(hrow) if hc is not None and c is hc] or [idx]
                cand = [cells[j] for j in span if j < len(cells) and cells[j] is not None]
            uniq: list = []
            for tc in cand:
                if tc not in uniq and tc not in seen:        # الخلية المدموجة رأسياً تُؤخذ مرة واحدة
                    uniq.append(tc)
            if not uniq:
                return None
            seen.update(uniq)
            return uniq[0] if len(uniq) == 1 else _merge_cells(uniq)

        # الأنشطة قبل المحتوى: الخلية الممتدة أفقياً عبر عمود المدرس والمحتوى تُعدّ نشاطاً لا محتوى
        t = {}
        for k in [c for c in READ_COLS if c != "content"] + ["content"]:
            t[k] = take(k)

        def blocks(k: str) -> list[list[str]]:
            return tc_blocks(t[k], numbering) if t.get(k) is not None else []

        stage = dedupe(_flat(blocks("stage")))
        tail = dedupe(_flat(blocks("segment")) + _flat(blocks("time")))       # اسم المقطع والحصيص
        caps, ev = blocks("cap"), blocks("eval")
        teacher, learner, generic = blocks("teacher"), blocks("learner"), blocks("activities")
        content = [e for e in t["content"] if e.tag in (W + "p", W + "tbl")] if t["content"] is not None else []
        for k in ("teacher", "learner", "activities", "eval", "cap", "stage"):
            content = content + _tables_of(t.get(k))
        if not (stage or tail or caps or ev or teacher or learner or generic or _has_visible(content)):
            continue
        if not cfg.learner_activities:
            dropped += len(_flat(learner))
        if cfg.split_diagnostic and _is_combined_intro(stage):
            a_caps, b_caps = _first_block_split(caps)
            a_tea, b_tea = _first_block_split(teacher)
            a_lea, b_lea = _first_block_split(learner)
            a_gen, b_gen = _first_block_split(generic)
            a_ev, b_ev = _first_block_split(ev)
            a_cont, b_cont = split_content(content, numbering, stage[1])
            rows.append(_make_row(dedupe(stage[:1] + tail), a_caps, a_tea, a_lea, a_gen, a_ev, a_cont,
                                  numbering, cfg))
            rows.append(_make_row(stage[1:], b_caps, b_tea, b_lea, b_gen, b_ev, b_cont, numbering, cfg))
            splits += 1
        else:
            rows.append(_make_row(dedupe(stage + tail), _flat(caps), _flat(teacher), _flat(learner),
                                  _flat(generic), _flat(ev), content, numbering, cfg))
    for i, r in enumerate(rows):
        r.idx = i
    if splits:
        report.ok("فُصلت مرحلة «التقويم التشخيصي» عن المرحلة التي تليها في صفّين مستقلين")
    if dropped:
        report.ok(f"حُذفت أنشطة المتعلم (عدد الأسطر: {dropped})، وبقيت أنشطة المدرس")
    return rows


# ═══════════════════ ترويسة الجذاذة (الوحدة، المجال، الأهداف…) ═══════════════════
_LEAD = re.compile(r"^[\s*•·+\-–—]+")
_FIELD_KEYS = {
    "domain": ("المجال",),
    "component": ("المكون",),
    "topic": ("موضوعالدرس", "الموضوع"),
    "duration": ("الحصيصالزمني", "مدهالانجاز", "المده", "الحصيص"),
    "target": ("الفئهالمستهدفه",),
    "specific": ("القدرهالخاصهبالدرس",),
}


def parse_header(texts: list[str]) -> dict:
    """يستخرج بيانات الترويسة من نصوص المصدر (بما فيها مربعات النص)."""
    info: dict = {"competencies": {}}
    for raw in texts:
        label, sep, value = _LEAD.sub("", raw).partition(":")
        label, value = label.strip(), value.strip()
        words = label.split()
        if not sep or not words:
            continue
        key = compact(label)
        if compact(words[0]) == "الوحده" and "unit" not in info:      # «الوحدة الثالثة: المجال الحضاري»
            if len(words) > 1:
                info["unit"] = " ".join(words[1:])
                parts = value.split(None, 1)
                if parts and compact(parts[0]) == "المجال":           # لا نكرّر كلمة «المجال» في قيمته
                    dom = parts[1] if len(parts) > 1 else ""
                else:
                    dom = value
                if dom and "domain" not in info:
                    info["domain"] = dom
            elif value:
                info["unit"] = value
            continue
        if key.startswith("الكفايه") and len(key) > 7 and value:
            kind = key[7:]
            info["competencies"].setdefault(kind[2:] if kind.startswith("ال") else kind, value)
            continue
        for k, aliases in _FIELD_KEYS.items():
            if k not in info and key in aliases and value:
                info[k] = value
                break
    return info


def topic_from_filename(path: Path) -> str:
    s = re.sub(r"[_\-.\d()]+", " ", path.stem)
    s = re.sub(r"\b(?:جذاذة|جذاذه|درس|تربوية|رقم)\b", " ", s)
    return re.sub(r"\s+", " ", s).strip() or "الدرس"


def value_cell(tcs: list, i: int):
    """خلية قيمة الحقل: أقرب خلية سابقة غير فارغة (نتخطى خلية النقطتين «:»)."""
    j = i - 1
    while j >= 0 and not compact(tc_text(tcs[j])):
        j -= 1
    return tcs[j] if j >= 0 else None


_LEVEL = re.compile(r"(الاولي|الثانيه|الثالثه)")


def _level(s: str):
    m = _LEVEL.search(norm(s))
    return m.group(1) if m else None


def fill_header(model: Model, hdr: dict, cfg: Settings, report: Report) -> None:
    wanted = {compact(lbl): (lbl, key) for lbl, key in HEADER_FIELDS}
    filled, missing = [], []
    for tr in model.header.findall(W + "tr"):
        tcs = tr.findall(W + "tc")
        for i, tc in enumerate(tcs):
            ck = compact(tc_text(tc))
            if ck == "الفئهالمستهدفه":                      # لا نغيّرها، لكن ننبّه إن اختلف المستوى
                vtc = value_cell(tcs, i)
                a = _level(tc_text(vtc)) if vtc is not None else None
                b = _level(hdr.get("target", ""))
                if a and b and a != b:
                    report.warn(f"مستوى المصدر ({hdr['target']}) يخالف مستوى النموذج؛ راجع «الفئة المستهدفة».")
                continue
            if ck not in wanted:
                continue
            lbl, key = wanted[ck]
            vtc = value_cell(tcs, i)
            if vtc is None:
                continue
            val = clean_text(hdr.get(key, ""))
            if val:
                fill_cell(vtc, [val], cfg)
                filled.append(f"{lbl}={val}")
            else:
                missing.append(lbl)
    if filled:
        report.ok("الترويسة: " + " | ".join(filled))
    if missing:
        report.warn("لم أجد في المصدر: " + "، ".join(missing) + " — بقيت قيمها من النموذج فراجعها.")


_AI_KEYS = {"تواصليه": "communicative", "منهجيه": "methodological", "معرفيه": "cognitive"}


def _objective_value(kind: str, hdr: dict, ai_obj):
    ai_key = _AI_KEYS.get(kind)
    if ai_obj and ai_key and ai_obj.get(ai_key):
        return ai_obj[ai_key]
    comps = hdr["competencies"]
    if comps.get(kind):
        return comps[kind]
    if kind == "معرفيه":                                   # أقرب مقابل في المصدر
        return hdr.get("specific") or comps.get("ثقافيه")
    return None


def fill_objectives(model: Model, hdr: dict, ai_obj, cfg: Settings, report: Report) -> None:
    """يحدّث أسطر «الأهداف التعلمية» في ترويسة النموذج (كانت تبقى نصاً قديماً من درس آخر)."""
    trs = model.header.findall(W + "tr")
    start = next((ri for ri, tr in enumerate(trs) for tc in tr.findall(W + "tc")
                  if compact(tc_text(tc)).startswith("الاهداف")), None)
    if start is None:
        return
    done, kept = 0, []
    for tr in trs[start + 1:]:
        first = tr.find(W + "tc")
        txt = tc_text(first) if first is not None else ""
        head = txt.split(":", 1)[0]
        ck = compact(head)
        if ":" not in txt or not ck.startswith("كفايه"):
            break
        label = head.strip(" -–•*\u00a0")
        value = _objective_value(ck[5:], hdr, ai_obj)
        if value:
            fill_cell(first, [f"- {label}: {clean_text(value)}"], cfg)
            done += 1
        else:
            kept.append(label)
    if done:
        report.ok(f"الأهداف التعلمية: حُدِّثت من المصدر (عددها {done})" + (" بصياغة الذكاء الاصطناعي" if ai_obj else ""))
    elif not kept:
        report.warn("لم أتعرف على أسطر «الأهداف التعلمية» في النموذج؛ بقيت كما هي فراجعها.")
    for label in kept:
        report.warn(f"بقي هدف «{label}» نصاً قديماً من النموذج (لا مقابل له في المصدر)؛ حدّثه يدوياً.")


# ═══════════════════════════ التنقيح بالذكاء الاصطناعي ═══════════════════════════
class AIError(Exception):
    pass


def resolve_api_key(cfg: Settings) -> str:
    """المفتاح من: الوسيط --api-key ← متغير البيئة ← ملف .env. لا يُخزَّن في الكود أبداً."""
    if cfg.api_key:
        return cfg.api_key.strip()
    for name in ("GEMINI_API_KEY", "GOOGLE_API_KEY"):
        if os.environ.get(name, "").strip():
            return os.environ[name].strip()
    for base in dict.fromkeys((Path.cwd(), Path(__file__).resolve().parent, Path("/storage/emulated/0"))):
        f = base / ".env"
        if not f.is_file():
            continue
        for line in f.read_text(encoding="utf-8", errors="ignore").splitlines():
            k, sep, v = line.strip().partition("=")
            if sep and k.strip() in ("GEMINI_API_KEY", "GOOGLE_API_KEY") and v.strip():
                return v.strip().strip("'\"")
    return ""


def _candidate_text(data: dict) -> str:
    cands = data.get("candidates") or []
    if not cands:
        reason = (data.get("promptFeedback") or {}).get("blockReason", "")
        raise AIError("استجابة بلا نتيجة" + (f" ({reason})" if reason else ""))
    parts = (cands[0].get("content") or {}).get("parts") or []
    text = "".join(p.get("text", "") for p in parts if isinstance(p, dict)).strip()
    if not text:
        raise AIError("استجابة فارغة")
    return text


def _parse_model_json(text: str, required: str = "rows") -> dict:
    text = re.sub(r"^```(?:json)?\s*|\s*```$", "", text.strip())
    i, j = text.find("{"), text.rfind("}")
    if i < 0 or j <= i:
        raise AIError("لا يوجد JSON في الاستجابة")
    obj = json.loads(text[i:j + 1])
    if not isinstance(obj, dict) or not isinstance(obj.get(required), list):
        raise AIError("بنية JSON غير متوقعة")
    return obj


def _http_error_message(e: urllib.error.HTTPError) -> str:
    try:
        return str(json.loads(e.read().decode("utf-8", "ignore")).get("error", {}).get("message", ""))[:200]
    except Exception:
        return ""


def call_gemini(prompt: str, cfg: Settings, required: str = "rows") -> dict:
    """يستدعي Gemini مع إعادة محاولة بتأخير متزايد للأخطاء العابرة (503/429/انقطاع الشبكة/JSON تالف)."""
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{cfg.ai_model}:generateContent"
    body = json.dumps({"contents": [{"parts": [{"text": prompt}]}],
                       "generationConfig": {"responseMimeType": "application/json", "temperature": 0.2}}
                      ).encode("utf-8")
    headers = {"Content-Type": "application/json", "x-goog-api-key": cfg.api_key}
    last = ""
    for attempt in range(1, cfg.ai_retries + 1):
        wait = None
        try:
            req = urllib.request.Request(url, data=body, method="POST", headers=headers)
            with urllib.request.urlopen(req, timeout=cfg.ai_timeout) as resp:
                data = json.loads(resp.read().decode("utf-8"))
            return _parse_model_json(_candidate_text(data), required)
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
        if attempt < cfg.ai_retries:
            delay = min(wait if wait is not None else 2 ** attempt, 60) + random.uniform(0, 1)
            say(f"  {last} — إعادة المحاولة بعد {delay:.0f}ث ({attempt}/{cfg.ai_retries - 1})")
            time.sleep(delay)
    raise AIError(last or "فشل الاتصال")


def as_lines(val) -> list[str]:
    """يقبل نصاً أو قائمة، وينظّف الأسطر ويحدّ عددها."""
    if isinstance(val, str):
        items = [val]
    elif isinstance(val, list):
        items = [x if isinstance(x, str) else str(x) for x in val]
    else:
        return []
    out = [clean_text(ln) for it in items for ln in it.splitlines()]
    return [ln for ln in out if ln][:40]


def _clean_ai_activities(lines: list[str], cfg: Settings) -> list[str]:
    """يضمن اتساق ردّ الذكاء الاصطناعي مع الإعدادات: بلا أنشطة المتعلم، وبعنوان المدرس أو دونه."""
    out: list[str] = []
    for ln in lines:
        key = compact(ln)
        if key == compact(LABEL_LEARNER) and not cfg.learner_activities:
            break                                    # كل ما بعد هذا العنوان أنشطة متعلم
        if key in (compact(LABEL_TEACHER), compact(LABEL_LEARNER)) and not cfg.activity_labels:
            continue                                 # لا عناوين: تبقى الجمل وحدها
        out.append(ln)
    if out and cfg.activity_labels and not cfg.learner_activities and compact(out[0]) != compact(LABEL_TEACHER):
        out.insert(0, LABEL_TEACHER)
    return out


def ai_refine(rows: list[Row], hdr: dict, cfg: Settings, report: Report):
    """يعرض الجذاذة على Gemini للتنقيح. يُرجع (أهداف مقترحة | None، هل نجح الاستدعاء؟).
    لا يمسّ العمود «المراحل» ولا «المحتوى»؛ وكل ردّ يُتحقق منه قبل اعتماده."""
    comps = hdr["competencies"]

    def full_content(r: Row) -> str:
        """نص محتوى الصف كاملاً (فقرات وجداول) لا معاينة مقتطعة، حتى يرى الذكاء الاصطناعي القاعدة والأمثلة."""
        parts = []
        for e in r.content:
            if e.tag == W + "p":
                t = p_text(e).strip()
                if t:
                    parts.append(t)
            elif e.tag == W + "tbl":
                cells = [" ".join(p_text(p) for p in tc.iter(W + "p")).strip() for tc in e.iter(W + "tc")]
                parts.append(" | ".join(c for c in cells if c))
        return "\n".join(parts)[:2500]

    def rule_text(r: Row) -> str:
        return "\n".join([full_content(r)] + list(r.activities) + list(r.capabilities))

    concl = [r for r in rows if r.key == "concl"]
    lesson_rule = "\n".join(rule_text(r) for r in concl).strip()
    if not lesson_rule:                              # القاعدة قد ترد في صف الملاحظة أو التطبيق
        lesson_rule = "\n".join(line for r in rows for line in rule_text(r).split("\n")
                                 if _CONCL_LINE.search(norm(line)))
    lesson_rule = lesson_rule[:1200]
    payload = {
        "lesson_rule_from_source": lesson_rule,
        "rows": [{"row_index": r.idx, "stage": " / ".join(r.stage), "key": r.key,
                  "capabilities": "\n".join(r.capabilities), "activities": "\n".join(r.activities),
                  "evaluation": "\n".join(r.evaluation), "content": full_content(r)}
                 for r in rows],
        "source_competencies": {"communicative": comps.get("تواصليه", ""),
                                "methodological": comps.get("منهجيه", ""),
                                "lesson_specific": hdr.get("specific", "")},
    }
    if cfg.learner_activities and not cfg.activity_labels:
        rule4 = ("4. activities: أسطر المدرس ثم سطر فارغ ثم أسطر المتعلم، بلا أي عنوان (لا تكتب «أنشطة المدرس» "
                 "ولا «أنشطة المتعلم»)، وكل نشاط في سطر مستقل، ولا تحذف أي جملة.\n")
    elif cfg.learner_activities:
        rule4 = ("4. activities: أبقِ العنوانين «أنشطة المدرس:» ثم «أنشطة المتعلم:» "
                 "وكل نشاط في سطر مستقل يبدأ بـ «- ».\n")
    else:
        head = "ابدأ بالعنوان «أنشطة المدرس:» ثم " if cfg.activity_labels else ""
        rule4 = ("4. activities: أنشطة المدرس وحدها (لا تكتب أنشطة المتعلم إطلاقاً)؛ " + head +
                 "كل نشاط في سطر مستقل يبدأ بـ «- ».\n")
    prompt = (
    "أنت مفتش تربوي خبير بالجذاذات المغربية لمادة اللغة العربية في التعليم الثانوي الإعدادي.\n"
    f"المطلوب: تنقيح جذاذة الدرس اللغوي «{hdr.get('topic', '')}» ({hdr.get('target', '')}) "
    "وإرجاعها بنفس البنية.\n\n"
    "القواعد الصارمة:\n"
    "1. أعد JSON فقط، بنفس عدد الصفوف وبنفس قيم row_index.\n"
    "2. احتفظ بكل أفكار المصدر وأمثلته وصحّح أخطاءها؛ كن موجزاً جداً لأن الجذاذة تُطبع في صفحة واحدة.\n"
    "3. capabilities (القدرات): إذا كانت فارغة، فاستنتج قدرة مناسبة للأنشطة، وإلا فأوجزها.\n"
    + rule4 +
    "5. evaluation (التقويم): إذا كانت فارغة، فاقترح سؤالاً تقويمياً مناسباً (مثل ركب جملة، أعرب..)، وإلا فأوجزها.\n"
    "6. diagnostic_content: هذا الحقل مخصص لملء الفراغات في عمود 'المحتوى'. \n"
    "   - في صف التقويم التشخيصي: اكتب الإجابات المتوقعة.\n"
    "   - في صف الخلاصة/الاستنتاج (إذا كان content_preview فارغاً): اكتب خلاصة موجزة للدرس.\n"
    "   - في صف التطبيق (إذا كان content_preview فارغاً): اكتب تمريناً تطبيقياً بسيطاً مع حله.\n"
    "   - إذا كان المحتوى موجوداً بالفعل، اتركه فارغاً.\n"
    "7. objectives: ثلاثة أهداف موجزة خاصة بهذا الدرس: communicative وmethodological وcognitive.\n"
    "8. الاتساق (إلزامي): كل ما تكتبه في الخانات الفارغة يجب أن يتقيّد بما في الجذاذة:\n"
    "   - التمرين والأسئلة يستعملان القاعدة والأمثلة الواردة في lesson_rule_from_source والصفوف الأخرى، لا قاعدة من عندك.\n"
    "   - التقويم يقيس الهدف نفسه الذي تقيسه الأنشطة في الصف، لا موضوعاً آخر.\n"
    "   - القدرات تصف ما تُنجزه الأنشطة فعلاً في الصف.\n"
    "   - الأهداف تتّسق مع القدرات والكفايات الواردة في المدخلات، ولا تكررها حرفياً.\n"
    "   - لا تُعِد كتابة ما هو موجود؛ املأ الفارغ فقط.\n"
    'الصيغة: {"rows":[{"row_index":0,"capabilities":"","activities":"","evaluation":"",'
    '"diagnostic_content":""}],"objectives":{"communicative":"","methodological":"","cognitive":""}}\n\n'
    "المدخلات:\n" + json.dumps(payload, ensure_ascii=False)
    )
    if not cfg.api_key:
        report.warn("لا يوجد مفتاح Gemini؛ اعتُمد نص المصدر دون تنقيح بالذكاء الاصطناعي.")
        return None, False
    say(f"جاري عرض الجذاذة على {cfg.ai_model} للتنقيح البيداغوجي…")
    try:
        result = call_gemini(prompt, cfg)
    except Exception as e:                         # أي فشل في الاتصال أو الرد لا يوقف التفريغ
        report.warn(f"تعذّر التنقيح بالذكاء الاصطناعي ({e}) — اعتُمد نص المصدر كما هو.")
        return None, False

    by_idx: dict[int, dict] = {}
    for item in result["rows"]:
        if isinstance(item, dict):
            try:
                by_idx[int(item.get("row_index"))] = item
            except (TypeError, ValueError):
                pass
    changed = 0
    for r in rows:
        item = by_idx.get(r.idx)
        if item is None:
            report.warn(f"لم يُرجِع الذكاء الاصطناعي الصف {r.idx + 1}؛ بقي نصه الأصلي.")
            continue
        for f in ("capabilities", "activities", "evaluation"):
            new = as_lines(item.get(f))
            if f == "activities":
                new = _clean_ai_activities(new, cfg)
            if new and not getattr(r, f):             # يملأ الخانات الفارغة فقط
                setattr(r, f, new)
                changed += 1
        diag = as_lines(item.get("diagnostic_content"))
        if diag and not r.diagnostic:
            r.diagnostic = diag
            changed += 1

    obj = result.get("objectives")
    ai_obj = None
    if isinstance(obj, dict):
        ai_obj = {k: " ".join(as_lines(obj.get(k))) for k in ("communicative", "methodological", "cognitive")}
        ai_obj = {k: v for k, v in ai_obj.items() if v} or None
    report.ok(f"الذكاء الاصطناعي ({cfg.ai_model}): عدد الحقول المعدَّلة {changed}")
    return ai_obj, True


# ═══════════════════════════ بناء الجدول الرئيسي والتحقق ═══════════════════════════
def insert_diagnostic(new_els: list, lines: list[str], cfg: Settings) -> None:
    """يضع الإجابات المتوقعة بعد سطر «إجابات المتعلمين…» إن وُجد، وإلا في أول الخلية."""
    paras = [make_p(ln, cfg) for ln in lines]
    pos = 0
    for i, e in enumerate(new_els[:3]):
        if e.tag == W + "p" and "اجابات" in compact(p_text(e)):
            pos = i + 1
            break
    new_els[pos:pos] = paras



# ═══════════════ مراحل الدرس اللغوي القانونية: تمهيد ← ملاحظة وتحليل ← استنتاج ← تطبيق ═══════════════
STAGE_KEYS = ("diag", "obs", "concl", "apply")
STAGE_NAMES = {"diag": "التقويم التشخيصي", "obs": "ملاحظة الأمثلة ووصفها وتحليلها",
               "concl": "الاستنتاج", "apply": "التطبيق والاستثمار"}
_CUE = {
    "diag": re.compile(r"تشخيص|تمهيد|استحضار|تهيئ|تهيي"),
    "obs": re.compile(r"ملاحظ|وصف|تحليل|فهم|اكتشاف|الامثله|مناقشه|تقويم تكويني|عرض الامثله|عرض الظاهره"),
    "concl": re.compile(r"قاعده|قواعد|الاستنتاج|استنتاج عام|خلاصه|تركيب|استخلاص"),
    "apply": re.compile(r"تطبيق|استثمار|توظيف|اعراب|تمارين|تمرين|تقويم اجمالي|تقويم ختامي|معالج"),
}
_CONCL_LINE = re.compile(r"قاعده|قواعد|خلاصه|الاستنتاج العام|استنتاج عام|تركيب|^الاستنتاج")
_TIMING = re.compile(r"مقطع|^\d+\s*(?:د|دقيق)")


def stage_cues(text: str) -> list[str]:
    """المراحل القانونية المذكورة في نص (عنوان مرحلة أو سطر) بترتيبها القانوني."""
    n = norm(text)
    partial_re = re.compile(r"(?:ال)?استنتاج\w*\s+(?:ال)?جزئي\w*")   # «الاستنتاج الجزئي» من الملاحظة لا من الاستنتاج العام
    partial = bool(partial_re.search(n))
    n = partial_re.sub(" ", n)
    keys = [k for k in STAGE_KEYS if _CUE[k].search(n)]
    if partial and "obs" not in keys:
        keys.insert(0 if "diag" not in keys else 1, "obs")
    if not keys and re.fullmatch(r"(?:ال)?تقويم", n.strip(" .:")):
        keys = ["apply"]                                  # «التقويم» وحده في آخر الجذاذة
    return keys


def _row_stage_text(r: Row) -> str:
    return " ".join(x for x in r.stage if not _TIMING.search(norm(x)))


def _clone_row(r: Row, key: str, **kw) -> Row:
    n = Row(idx=0, stage=list(r.stage), key=key)
    for k, v in kw.items():
        setattr(n, k, v)
    n.orig_lines = [x for x in (n.capabilities + n.activities + n.evaluation) if x]
    n.content_preview = ""
    return n


def _split_obs_concl(r: Row) -> list[Row]:
    """صف يجمع «الملاحظة» و«الاستنتاج»: ما فيه قاعدة/خلاصة يُنقل إلى الاستنتاج، والباقي يبقى في الملاحظة."""
    def part(lines):
        """ما قبل أول سطر فيه مؤشر الاستنتاج للملاحظة، وما بعده (بسطره) للاستنتاج."""
        i = next((k for k, ln in enumerate(lines) if ln and _CONCL_LINE.search(norm(ln))), None)
        return (list(lines), []) if i is None else (lines[:i], lines[i:])
    a_act, b_act = part(r.activities)
    a_cap, b_cap = part(r.capabilities)
    a_ev, b_ev = part(r.evaluation)
    idx = next((i for i, e in enumerate(r.content)
                if e.tag == W + "p" and _CONCL_LINE.search(norm(p_text(e)))), None)
    a_c, b_c = (r.content[:idx], r.content[idx:]) if idx is not None else (list(r.content), [])
    return [_clone_row(r, "obs", capabilities=a_cap, activities=a_act, evaluation=a_ev, content=a_c),
            _clone_row(r, "concl", capabilities=b_cap, activities=b_act, evaluation=b_ev, content=b_c)]


def merge_rows(rows: list[Row]) -> Row:
    m = Row(idx=0)
    for r in rows:
        m.capabilities += [x for x in r.capabilities if x not in m.capabilities]
        m.activities += r.activities
        m.evaluation += [x for x in r.evaluation if x not in m.evaluation]
        m.content += r.content
        m.diagnostic += r.diagnostic
        stage_set = set(r.stage)
        m.orig_lines += [x for x in r.orig_lines if x not in stage_set]
        m.content_preview = (m.content_preview + " " + r.content_preview).strip()
    return m


def canonicalize_rows(rows: list[Row], cfg: Settings, report: Report) -> list[Row]:
    """يوزّع صفوف المصدر على المراحل الأربع القانونية دون خلطها، ويدمج ما كان من مرحلة واحدة."""
    items, prev, shown = [], None, []
    for r in rows:
        keys = stage_cues(_row_stage_text(r))
        if not keys:
            keys = [prev or "obs"]
        label = clean_text(_row_stage_text(r)) or "(بلا عنوان)"
        if len(keys) == 2 and keys == ["obs", "concl"]:
            a, b = _split_obs_concl(r)
            items += [("obs", a), ("concl", b)]
            shown.append(f"«{label[:30]}» ← ملاحظة + استنتاج (فُصل)")
        else:
            items.append((keys[0], r))
            shown.append(f"«{label[:30]}» ← {STAGE_NAMES[keys[0]]}")
        prev = items[-1][0]
    by: dict[str, list[Row]] = {}
    for k, r in items:
        by.setdefault(k, []).append(r)
    out = []
    for k in STAGE_KEYS:
        if k in by:
            m = merge_rows(by[k])
            m.key, m.stage = k, [STAGE_NAMES[k]]
            out.append(m)
    for i, r in enumerate(out):
        r.idx = i
    report.info("توزيع المراحل: " + " | ".join(shown))
    return out


def _cue_of(text: str):
    ks = stage_cues(text)
    return ks[0] if ks else None


def explode_list_rows(grid, cols, header_rows, numbering, cfg, report):
    """جذاذة بقوائم متوازية داخل خلايا قليلة (المراحل أسطر في خلية، وأنشطتها أسطر في خلية أخرى)."""
    body = list(range(header_rows, len(grid)))
    if not body or len(body) > 2 or "stage" not in cols:
        return None
    seen: set = set()
    data = {k: [] for k in READ_COLS}
    content: list = []
    for ri in body:
        cells = grid[ri]
        for k in READ_COLS:
            idx = cols.get(k)
            if idx is None or idx >= len(cells):
                continue
            tc = cells[idx]
            if tc is None or tc in seen:
                continue
            seen.add(tc)
            if k == "content":
                content += [e for e in tc if e.tag in (W + "p", W + "tbl")]
            else:
                data[k] += _flat(tc_blocks(tc, numbering))
                content += _tables_of(tc)
    labels = [x for x in data["stage"] if x.strip()]
    if len(labels) < 3 or any(len(x) > 70 for x in labels) or sum(1 for x in labels if stage_cues(x)) < 2:
        return None
    n = len(labels)
    lk = [set(stage_cues(x)) for x in labels]                 # عنوان مركّب («ملاحظة واستنتاج») يحمل مؤشرين
    EQ = {"obs": ("diag",), "concl": ("obs",), "apply": ("concl",), "diag": ()}
    num_re = re.compile(r"^\W*(\d+)\s*[-–.)]")

    def walk(texts):
        idxs, cur, prev_num = [], 0, None
        for t in texts:
            c = _cue_of(t)
            m = num_re.match(t or "")
            num = int(m.group(1)) if m else None
            moved = False
            if c:
                for want in (c,) + EQ.get(c, ()):
                    j = next((i for i in range(cur, n) if want in lk[i]), None)
                    if j is not None:
                        moved = j > cur
                        cur = j
                        break
            if not moved and num is not None and prev_num is not None and num <= prev_num and cur < n - 1:
                cur += 1                                       # إعادة الترقيم دون تقدّم بالمؤشر: كتلة جديدة (تمارين مثلاً)
            if num is not None:
                prev_num = num
            idxs.append(cur)
        return idxs

    approx, assign = False, {}
    for key in ("teacher", "activities", "cap", "eval", "learner"):
        items = data[key]
        if not items:
            assign[key] = []
        elif len(items) == n:
            assign[key] = list(range(n))
        elif key == "learner" and assign.get("teacher") and len(items) == len(assign["teacher"]):
            assign[key] = list(assign["teacher"])
        else:
            approx = True
            assign[key] = walk(items)
    cidx = walk([p_text(e) if e.tag == W + "p" else "" for e in content]) if content else []
    rows = []
    for i, lab in enumerate(labels):
        def pick(key, i=i):
            return [it for it, j in zip(data[key], assign[key]) if j == i]
        c = [e for e, j in zip(content, cidx) if j == i]
        rows.append(_make_row([lab], pick("cap"), pick("teacher"), pick("learner"), pick("activities"),
                              pick("eval"), c, numbering, cfg))
    for i, r in enumerate(rows):
        r.idx = i
    report.info("المصدر قوائم متوازية داخل خلايا قليلة" + ("؛ وُزّعت الأنشطة على المراحل تقريباً فراجعها" if approx else ""))
    return rows


def parse_header_general(texts: list[str], source_path=None) -> dict:
    """ترويسة الجذاذة: القارئ الأصلي ثم تقويمه بقارئ يفهم أزواج (القيمة | : | العنوان) وعدة عناوين في سطر."""
    info = parse_header(texts)
    try:
        import reading_general as _G
        if source_path is not None:
            _src = _G.Source(str(source_path))
            gi, gc = _src.info, _src.comps
        else:
            gi, gc = _G.parse_pieces(texts)
    except Exception:
        return info
    for gk, hk in (("unit", "unit"), ("domain", "domain"), ("component", "component"), ("topic", "topic"),
                   ("duration", "duration"), ("level", "target")):
        if gi.get(gk):
            info[hk] = gi[gk]
    for k, v in gc.items():
        ck = compact(k)
        ck = ck[2:] if ck.startswith("ال") else ck
        info["competencies"].setdefault(ck, v)
    return info

def fill_main_table(model: Model, rows: list[Row], ctx: Ctx) -> None:
    """يملأ صفوف النموذج نفسها بحسب مراحلها (يبقى عمود المراحل بأسماء النموذج)."""
    cfg = ctx.cfg
    main, c = model.main, model.cols
    trs = main.findall(W + "tr")
    body = trs[cfg.model_header_rows:]

    def first_line(tr) -> str:
        tc = tr.findall(W + "tc")[c["stage"]]
        for p in iter_paragraphs(tc):
            t = clean_text(p_text(p).replace("\n", " "))
            if t:
                return t
        return ""

    mkeys = [(stage_cues(first_line(tr)) or [None])[0] for tr in body]
    if None in mkeys or len(set(mkeys)) != len(mkeys):          # أسماء غير واضحة: نعتمد الترتيب
        mkeys = [STAGE_KEYS[min(i, len(STAGE_KEYS) - 1)] for i in range(len(body))]
    by: dict[str, list[Row]] = {}
    for r in rows:
        k = r.key if r.key in mkeys else next((m for m in reversed(mkeys[:max(1, STAGE_KEYS.index(r.key) + 1)]) if m), mkeys[0])
        by.setdefault(k, []).append(r)
    for tr, key in zip(body, mkeys):
        tcs = tr.findall(W + "tc")
        row = merge_rows(by[key]) if key in by else None
        if row is None:
            ctx.report.warn(f"لا توجد في المصدر مرحلة «{STAGE_NAMES[key]}»؛ بقي صفها فارغاً.")
            row = Row(idx=0)
        fill_cell(tcs[c["capabilities"]], row.capabilities, cfg)
        fill_cell(tcs[c["activities"]], row.activities, cfg)
        fill_cell(tcs[c["evaluation"]], row.evaluation, cfg)
        ctc = tcs[c["content"]]
        clear_cell(ctc)
        new_els = convert_content(row.content, ctx, cell_inner_width(ctc))
        if row.diagnostic:
            insert_diagnostic(new_els, row.diagnostic, cfg)
        for e in new_els:
            ctc.append(e)
        ensure_trailing_paragraph(ctc, cfg)


def verify(model: Model, rows: list[Row], ctx: Ctx, ai_used: bool, report: Report) -> None:
    """فحص آلي: هل وصلت نصوص المصدر كلها إلى الجدول الناتج؟"""
    blob = compact(" ".join(clean_text(p_text(p).replace("\n", " ")) for p in iter_paragraphs(model.main)))

    def missing(lines: list[str]) -> list[str]:
        return [ln for ln in lines if compact(ln) and compact(ln) not in blob]

    miss_content = missing(ctx.expected)
    narrow = [ln for r in rows for ln in r.orig_lines]
    miss_narrow = missing(narrow)
    if miss_content:
        report.warn(f"بعض أسطر «المحتوى» لم تظهر في الناتج (عددها {len(miss_content)})، أولها: «{miss_content[0][:40]}»")
    if ai_used:
        report.info(f"أُعيدت صياغة {len(miss_narrow)} من أصل {len(narrow)} سطراً (القدرات/الأنشطة/التقويم) بالذكاء الاصطناعي")
    elif miss_narrow:
        report.warn(f"بعض أسطر الأعمدة الضيقة لم تظهر في الناتج (عددها {len(miss_narrow)})، أولها: «{miss_narrow[0][:40]}»")
    if not miss_content and (ai_used or not miss_narrow):
        report.ok("التحقق: " + ("المحتوى الأصلي كله موجود في الناتج" if ai_used
                               else "كل نصوص المصدر موجودة في الناتج"))


# ═══════════════════════════ الملفات والتشغيل ═══════════════════════════
def find_file(name: str) -> Path:
    """يجد الملف بمساره، أو بالاسم في المجلد الحالي/مجلد السكريبت (يتسامح مع المسافات والشرطات)."""
    p = Path(name).expanduser()
    if p.is_file():
        return p

    def key(s: str) -> str:
        return compact(re.sub(r"[_\-\s()]+", "", unicodedata.normalize("NFC", s)))

    target = key(p.name)
    for base in dict.fromkeys((Path.cwd(), Path(__file__).resolve().parent)):
        direct = base / p.name
        if direct.is_file():
            return direct
        depth = len(base.parts)
        for root, dirs, files in os.walk(base):
                dirs[:] = [d for d in dirs if not d.startswith(".") and d not in (OUT_DIR, "node_modules", "__pycache__", "Android")]
                
                for f in files:
                    if f.lower().endswith(".docx") and not f.startswith("~$") and key(f) == target:
                    	return Path(root, f)
    raise FatalError(f"لم أجد الملف «{name}». ضعه في المجلد الحالي أو مرّر مساره كاملاً "
                     "(وإن أردت تسمية الملف الناتج فاستعمل --out-name).")


def output_path(cfg: Settings, hdr: dict) -> Path:
    topic = hdr.get("topic", "الدرس")
    name = cfg.out_name or f"جذاذة_{topic}_الاحترافية.docx"
    name = re.sub(r'[<>:"/\\|?*\r\n\t]+', " ", name).strip(" .")
    name = re.sub(r"\s+", "_", name)
    if not name.lower().endswith(".docx"):
        name += ".docx"
    
    # استخراج المكون والوحدة من الترويسة لإنشاء المجلدات
    comp = hdr.get("component", "الدرس اللغوي").strip() or "الدرس اللغوي"
    unit = hdr.get("unit", "بدون وحدة").strip() or "بدون وحدة"
    
    # تنظيف أسماء المجلدات من الرموز التي يمنعها النظام
    comp = re.sub(r'[<>:"/\\|?*\r\n\t]+', "", comp).strip()
    unit = re.sub(r'[<>:"/\\|?*\r\n\t]+', "", unit).strip()
    
    # المسار الجديد: الجذاذات_المنجزة / المكون / الوحدة
    target_dir = cfg.out_dir / comp / unit
    target_dir.mkdir(parents=True, exist_ok=True)
    
    path, stem, n = target_dir / name, Path(name).stem, 2
    while path in cfg.used_outputs:                    # مصدران بالموضوع نفسه: لا نكتب فوق الأول
        path = target_dir / f"{stem}_{n}.docx"
        n += 1
    cfg.used_outputs.add(path)
    return path


def save_safely(doc, path: Path) -> Path:
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


def build(model_path: Path, source_path: Path, cfg: Settings, report: Report) -> Path:
    say(f"النموذج: {model_path}\nالمصدر : {source_path}")
    doc, src = Document(str(model_path)), Document(str(source_path))
    model = analyze_model(doc, cfg)

    numbering = SourceNumbering(src)
    numbering.prepare(src.element.body)
    hdr = parse_header_general(all_texts(src), source_path)
    if not hdr.get("topic"):
        hdr["topic"] = topic_from_filename(source_path)
        report.warn(f"لم أجد «موضوع الدرس» في المصدر؛ اعتُمد من اسم الملف: {hdr['topic']}")

    grid, hrows, cols = analyze_source(src, cfg)
    absent = [name for key, name in OPTIONAL_COLS.items() if key not in cols]
    if absent:
        report.info("أعمدة غير موجودة في المصدر (تُترك فارغة): " + "، ".join(absent))
    rows = extract_rows(grid, cols, hrows, numbering, cfg, report)
    if not rows:
        raise FatalError("لم أجد صفوفاً تحتوي بيانات في جدول المصدر.")
    rows = canonicalize_rows(rows, cfg, report)

    ai_obj, ai_used = ai_refine(rows, hdr, cfg, report) if cfg.use_ai else (None, False)

    ctx = Ctx(cfg, report, numbering, src.part, doc, model.nested_proto)
    fill_header(model, hdr, cfg, report)
    if cfg.update_objectives:
        fill_objectives(model, hdr, ai_obj, cfg, report)
    fill_main_table(model, rows, ctx)
    report.ok(f"الجدول الرئيسي: عدد الصفوف {len(rows)}")

    for section in doc.sections:
        section.top_margin = section.bottom_margin = Twips(cfg.margin_twips)
        section.left_margin = section.right_margin = Twips(cfg.margin_twips)
    doc.core_properties.title = f"جذاذة {hdr['topic']}"

    verify(model, rows, ctx, ai_used, report)
    out = save_safely(doc, output_path(cfg, hdr))

    report.ok(f"الملف: {out}")
    return out


def parse_args(argv):
    ap = argparse.ArgumentParser(
        prog="build_jadhadha.py",
        description="ينقل جذاذة (المصدر) إلى نموذج جذاذة جاهز بتنسيق موحَّد، مع تنقيح اختياري بالذكاء الاصطناعي.")
    ap.add_argument("model", nargs="?", default=DEFAULT_MODEL, help=f"ملف النموذج (الافتراضي: {DEFAULT_MODEL})")
    ap.add_argument("sources", nargs="*", help=f"ملف أو أكثر من جذاذات المصدر (الافتراضي: {DEFAULT_SOURCE})")
    ap.add_argument("-o", "--out-dir", default=OUT_DIR, help="مجلد الإخراج")
    ap.add_argument("--out-name", default="", help="اسم الملف الناتج (لمصدر واحد فقط)")
    ap.add_argument("--no-ai", action="store_true", help="تعطيل التنقيح بالذكاء الاصطناعي")
    ap.add_argument("--ai-model", default=os.environ.get("GEMINI_MODEL", DEFAULT_AI_MODEL), help="اسم نموذج Gemini")
    ap.add_argument("--api-key", default="", help="مفتاح Gemini (الأفضل متغير البيئة GEMINI_API_KEY)")
    ap.add_argument("--font", default="Sakkal Majalla", help="اسم الخط")
    ap.add_argument("--size", type=float, default=10.0, help="حجم الخط بالنقاط")
    ap.add_argument("--margin", type=int, default=400, help="الهوامش بوحدة twips (20 twip = 1 نقطة)")
    ap.add_argument("--keep-colors", action="store_true", help="الاحتفاظ بألوان نص المصدر")
    ap.add_argument("--no-split", action="store_true",
                    help="عدم فصل «التقويم التشخيصي» عن «ملاحظة الأمثلة ووصفها» (يبقيان في صف واحد)")
    ap.add_argument("--drop-learner", action="store_true",
                    help="حذف جمل «أنشطة المتعلم» (الافتراضي: تبقى كلها)")
    ap.add_argument("--labels", action="store_true",
                    help="كتابة عناوين «أنشطة المدرس:/المتعلم:» (الافتراضي: بلا عناوين)")
    ap.add_argument("--with-learner", action="store_true", help=argparse.SUPPRESS)          # قديم: بلا أثر
    ap.add_argument("--no-activity-labels", action="store_true", help=argparse.SUPPRESS)   # قديم: بلا أثر
    ap.add_argument("--no-reflow", action="store_true",
                    help="عدم دمج الفقرات المتتالية: كل فقرة سطر (لمصادر لا تقطّع الجمل على فقرات)")
    ap.add_argument("--keep-model-objectives", action="store_true", help="إبقاء «الأهداف التعلمية» كما في النموذج")
    ap.add_argument("-v", "--verbose", action="store_true", help="عرض تفاصيل الأخطاء")
    return ap.parse_args(argv)


def main(argv=None) -> int:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass
    args = parse_args(argv)
    global JOIN_FRAGMENTS
    JOIN_FRAGMENTS = not args.no_reflow
    cfg = Settings(font=args.font, size_pt=args.size, margin_twips=args.margin, keep_colors=args.keep_colors,
                   update_objectives=not args.keep_model_objectives, use_ai=not args.no_ai,
                   ai_model=args.ai_model, api_key=args.api_key, out_dir=Path(args.out_dir),
                   out_name=args.out_name, verbose=args.verbose,
                   split_diagnostic=not args.no_split, learner_activities=not args.drop_learner,
                   activity_labels=args.labels)
    try:
        model_path = find_file(args.model)
        sources = [find_file(s) for s in (args.sources or [DEFAULT_SOURCE])]
    except FatalError as e:
        print(f"✖ {e}", file=sys.stderr)
        return 1
    if len(sources) > 1 and cfg.out_name:
        print("تنبيه: --out-name يُتجاهل مع أكثر من مصدر.")
        cfg.out_name = ""


    if cfg.use_ai:
        cfg.api_key = resolve_api_key(cfg)
        if cfg.api_key:
            say(f"الذكاء الاصطناعي: مفعّل ({cfg.ai_model})")
        else:
            say("الذكاء الاصطناعي: غير مفعّل — لا يوجد مفتاح GEMINI_API_KEY (راجع أعلى الملف).")
            cfg.use_ai = False

    failures = 0
    for src in sources:
        report = Report()
        try:
            build(model_path, src, cfg, report)
        except FatalError as e:
            report.error(str(e))
            failures += 1
        except Exception as e:                       # خطأ غير متوقع: رسالة موجزة، والتفاصيل مع -v
            if cfg.verbose:
                traceback.print_exc()
            report.error(f"خطأ غير متوقع ({type(e).__name__}: {e}) — أعد التشغيل مع -v لمعرفة التفاصيل.")
            failures += 1
        report.show()
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
