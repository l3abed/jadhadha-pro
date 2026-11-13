#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_insha.py — جذاذة التعبير والإنشاء (مع الذكاء الاصطناعي)
===========================================================
يحوّل جذاذة مصدر إلى نموذج التعبير والإنشاء ذي الحصص الأربع،
مع ملء الفراغات وتنقيح الأنشطة أوتوماتيكياً بواسطة Gemini.
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
DEFAULT_MODEL = "التدرب على كتابة اليوميات (1).docx"
DEFAULT_SOURCE = "مهارة -كتابة حكاية عجيبة أو من الخيال العلمي---المختار 3إع.docx"
OUT_DIR = "الجذاذات_المنجزة"

LABEL_TEACHER = "أنشطة المدرس:"
LABEL_LEARNER = "أنشطة المتعلم:"
MEANS_FIELD = "reference"          
OBJECTIVE_SESSIONS = (1, 2, 3)     

@dataclass
class Settings:
    font: str = "Sakkal Majalla"
    size_pt: float = 10.0
    bold: bool = True
    margin_twips: int = 0              
    keep_colors: bool = False
    join_fragments: bool = True
    labels: bool = True                
    learner: bool = True               
    nested_proto_index: int = 0
    out_dir: Path = Path(OUT_DIR)
    out_name: str = ""
    verbose: bool = False
    used_outputs: set = field(default_factory=set)
    
    # --- إعدادات الذكاء الاصطناعي ---
    use_ai: bool = True
    ai_model: str = "gemini-3.5-flash"
    api_key: str = ""
    ai_retries: int = 4
    ai_timeout: int = 120

class FatalError(Exception): pass
class AIError(Exception): pass

# ═════════════════════════════ دوال الذكاء الاصطناعي ═════════════════════════════
def _candidate_text(data: dict) -> str:
    cands = data.get("candidates") or []
    if not cands: raise AIError("استجابة بلا نتيجة")
    parts = (cands[0].get("content") or {}).get("parts") or []
    text = "".join(p.get("text", "") for p in parts if isinstance(p, dict)).strip()
    if not text: raise AIError("استجابة فارغة")
    return text

def _parse_model_json(text: str) -> dict:
    text = re.sub(r"^```(?:json)?\s*|\s*```$", "", text.strip())
    i, j = text.find("{"), text.rfind("}")
    if i < 0 or j <= i: raise AIError("لا يوجد JSON في الاستجابة")
    obj = json.loads(text[i:j + 1])
    if not isinstance(obj, dict) or not isinstance(obj.get("rows"), list): raise AIError("بنية JSON غير متوقعة")
    return obj

def call_gemini(prompt: str, cfg: Settings) -> dict:
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{cfg.ai_model}:generateContent"
    body = json.dumps({"contents": [{"parts": [{"text": prompt}]}],
                       "generationConfig": {"responseMimeType": "application/json", "temperature": 0.2}}).encode("utf-8")
    headers = {"Content-Type": "application/json", "x-goog-api-key": cfg.api_key}
    last = ""
    for attempt in range(1, cfg.ai_retries + 1):
        try:
            req = urllib.request.Request(url, data=body, method="POST", headers=headers)
            with urllib.request.urlopen(req, timeout=cfg.ai_timeout) as resp:
                return _parse_model_json(_candidate_text(json.loads(resp.read().decode("utf-8"))))
        except urllib.error.HTTPError as e:
            last = f"HTTP {e.code}"
            if e.code not in (429, 500, 502, 503, 504): raise AIError(last)
        except Exception as e:
            last = str(e)
        if attempt < cfg.ai_retries:
            time.sleep(min(2 ** attempt, 60) + random.uniform(0, 1))
    raise AIError(last or "فشل الاتصال")

def as_lines(val) -> list[str]:
    if isinstance(val, str): items = [val]
    elif isinstance(val, list): items = [x if isinstance(x, str) else str(x) for x in val]
    else: return []
    out = [clean_text(ln) for it in items for ln in it.splitlines()]
    return [ln for ln in out if ln][:40]

def ai_refine_insha(rows: list[SrcRow], hdr: dict, cfg: Settings, report: Report):
    def full_text(r: SrcRow) -> str:
        """محتوى الصف كاملاً (فقرات ومربعات وجداول) لا معاينة مقتطعة."""
        return "\n".join(tbox_text(e) for e in r.content if tbox_text(e))[:2500]

    payload = {
        "lesson_context": {"stages": [" / ".join(r.stage) for r in rows if r.stage]},
        "rows": [{"row_index": r.idx, "stage": " / ".join(r.stage),
                  "capabilities": "\n".join(r.caps), "activities": "\n".join(r.teacher + r.learner + r.generic),
                  "evaluation": "\n".join(r.ev),
                  "content": full_text(r)}
                 for r in rows]
    }
    prompt = (
        "أنت مفتش تربوي خبير بالجذاذات المغربية، مكون التعبير والإنشاء.\n"
        f"المطلوب: تنقيح جذاذة مهارة «{hdr.get('topic', '')}» وإرجاعها بنفس البنية.\n\n"
        "القواعد الصارمة:\n"
        "1. أعد JSON فقط، بنفس عدد الصفوف وبنفس قيم row_index.\n"
        "2. احتفظ بكل أفكار المصدر؛ كن موجزاً جداً.\n"
        "3. capabilities (القدرات): إذا كانت فارغة، فاستنتج قدرة مناسبة للحصة وإلا فأوجزها.\n"
        "4. activities (الأنشطة): نسّق الأنشطة في أسطر تبدأ بـ «- ».\n"
        "5. evaluation (التقويم): إذا كانت فارغة، فاقترح تقويماً تكوينياً مناسباً.\n"
        "6. diagnostic_content: مخصص لملء الفراغات في عمود 'المحتوى' إذا كان شبه فارغ:\n"
        "   - في حصة الاكتساب: اقترح خطوات المهارة أو استنتاجاً.\n"
        "   - في التطبيق والإنتاج: اقترح 'نص الموضوع' مناسباً للتدرب.\n"
        "   - في التصحيح: اقترح شبكة تصحيح مبسطة.\n"
        "   - اتركه فارغاً إذا كان المحتوى الأصلي كافياً.\n"
        "7. الاتساق (إلزامي): كل ما تكتبه في الخانات الفارغة يتقيّد بما في الجذاذة نفسها:\n"
        "   - نص الموضوع والتمرين يستعملان المعايير والأمثلة الواردة في المحتوى والصفوف الأخرى، لا معايير من عندك.\n"
        "   - التقويم يقيس المهارة نفسها التي تدرّبها أنشطة الصف، لا موضوعاً آخر.\n"
        "   - القدرات تصف ما تُنجزه الأنشطة فعلاً في الصف.\n"
        "   - لا تعِد كتابة ما هو موجود؛ املأ الفارغ فقط.\n"
        'الصيغة: {"rows":[{"row_index":0,"capabilities":"","activities":"","evaluation":"",'
        '"diagnostic_content":""}]}\n\n'
        "المدخلات:\n" + json.dumps(payload, ensure_ascii=False)
    )
    say(f"جاري عرض الجذاذة على {cfg.ai_model} للتنقيح وملء الفراغات…")
    try:
        result = call_gemini(prompt, cfg)
    except AIError as e:
        report.warn(f"تعذّر التنقيح بالذكاء الاصطناعي ({e}) — اعتُمد نص المصدر كما هو.")
        return False

    by_idx = {int(item.get("row_index")): item for item in result["rows"] if isinstance(item, dict) and "row_index" in item}
    changed = 0
    for r in rows:
        item = by_idx.get(r.idx)
        if not item: continue
        for f, attr in (("capabilities", "caps"), ("activities", "teacher"), ("evaluation", "ev")):
            new = as_lines(item.get(f))
            if new and not getattr(r, attr):          # الذكاء الاصطناعي يملأ الخانات الفارغة فقط
                setattr(r, attr, new)
                changed += 1
        diag = as_lines(item.get("diagnostic_content"))
        if diag and not r.diagnostic:
            r.diagnostic = diag
            changed += 1
    report.ok(f"الذكاء الاصطناعي ({cfg.ai_model}): تم تنقيح وتعبئة {changed} حقل.")
    return True
    # ═════════════════════════════ ثوابت OOXML ═════════════════════════════
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
MC = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

PPR_ORDER = ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc", "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange"]
RPR_ORDER = ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing", "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath"]
TBLPR_ORDER = ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize", "tblStyleColBandSize", "tblW", "jc", "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook", "tblCaption", "tblDescription"]
TCPR_ORDER = ["cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd", "noWrap", "tcMar", "textDirection", "tcFitText", "vAlign", "hideMark"]

class Report:
    ICONS = {"ok": "✔", "info": "•", "warn": "⚠", "error": "✖"}
    def __init__(self) -> None: self.items: list[tuple[str, str]] = []
    def ok(self, msg: str) -> None: self.items.append(("ok", msg))
    def info(self, msg: str) -> None: self.items.append(("info", msg))
    def warn(self, msg: str) -> None: self.items.append(("warn", msg))
    def error(self, msg: str) -> None: self.items.append(("error", msg))
    def show(self) -> None:
        print("─" * 48)
        for kind, msg in self.items: print(f"{self.ICONS[kind]} {msg}")

def say(msg: str) -> None: print(msg, flush=True)
# ═════════════════════════════ أدوات النصوص ═════════════════════════════
_MARKS = re.compile("[\u0640\u064B-\u065F\u0670]")   

def clean_text(s: str) -> str:
    s = s.replace("\u0640", "").replace("\u00a0", " ")
    s = re.sub(r"[ \t\r\f\v]+", " ", s).strip()
    s = re.sub(r"^\*+\s*", "+ ", s) if re.sub(r"^\*+\s*", "", s) else s     # نجمة بداية السطر رمز نقطة في المصدر
    s = re.sub(r"\s+([:،؛؟.])", r"\1", s)
    s = re.sub(r"(?<=[\u0621-\u064a])«", " «", s)
    s = re.sub(r'،(?=[^\s)\]»"”])', "، ", s)
    return s

def norm(s: str) -> str:
    s = _MARKS.sub("", s or "")
    s = re.sub("[أإآٱ]", "ا", s).replace("ى", "ي").replace("ة", "ه")
    s = re.sub("[\u200e\u200f\u061c]", "", s)
    return re.sub(r"\s+", " ", s).strip()

def compact(s: str) -> str:
    return re.sub(r"[^\w]+", "", norm(s))

def p_text(p) -> str:
    out: list[str] = []
    def walk(node) -> None:
        for ch in node:
            tag = ch.tag
            if not isinstance(tag, str): continue
            if tag == W + "t": out.append(ch.text or "")
            elif tag == W + "tab": out.append(" ")
            elif tag in (W + "br", W + "cr"): out.append("\n")
            elif tag == W + "noBreakHyphen": out.append("-")
            elif tag in (W + "txbxContent", W + "delText", W + "instrText", W + "pPr", MC + "Fallback"): continue
            else: walk(ch)
    walk(p)
    return "".join(out)

def tc_text(tc) -> str:
    return " ".join(clean_text(p_text(p).replace("\n", " ")) for p in tc.findall(W + "p")).strip()

def iter_paragraphs(root):
    for p in root.iter(W + "p"):
        if any(a.tag == MC + "Fallback" for a in p.iterancestors()): continue
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

_TERMINAL = (".", ":", "؟", "?", "!", "؛", ";", "…")
_BULLET = re.compile(r"^(?:[-–—•*+·▪●◦■□]|\(?\d{1,2}\s*[-–.)/]|\(?[ا-ي]\s*[-–.)]\s)")
_ORDINAL = re.compile(r"^(?:اولا|ثانيا|ثالثا|رابعا|خامسا|سادسا|سابعا|ثامنا|تاسعا|عاشرا)(?![ء-ي])")
_PUNCT_ONLY = re.compile(r"^[.،:؛؟!…]+$")

def _starts_item(t: str) -> bool:
    n = norm(t)
    return bool(_BULLET.match(n) or _ORDINAL.match(n))

def reflow(paragraphs, join_fragments: bool) -> list[str]:
    lines: list[str] = []
    after_blank = True
    for raw in paragraphs:
        t = clean_text(raw)
        if not t:
            after_blank = True
            continue
        if _PUNCT_ONLY.match(t):
            if lines: lines[-1] = lines[-1].rstrip() + t
            after_blank = False
            continue
        if (not join_fragments or not lines or after_blank or _starts_item(t)
                or lines[-1].rstrip().endswith(_TERMINAL)):
            lines.append(t)
        else:
            lines[-1] = lines[-1].rstrip() + " " + t
        after_blank = False
    return lines

class SourceNumbering:
    def __init__(self, doc) -> None:
        self.levels, self.num_abs, self.starts, self.labels = {}, {}, {}, {}
        try: root = doc.part.numbering_part.element
        except Exception: return
        for an in root.findall(qn("w:abstractNum")):
            lv = {}
            for lvl in an.findall(qn("w:lvl")):
                lv[int(lvl.get(qn("w:ilvl")))] = (_val(lvl, "w:numFmt", "decimal"), _val(lvl, "w:lvlText", ""), int(_val(lvl, "w:start", "1") or 1))
            self.levels[an.get(qn("w:abstractNumId"))] = lv
        for n in root.findall(qn("w:num")):
            nid = n.get(qn("w:numId"))
            ab = n.find(qn("w:abstractNumId"))
            if ab is not None: self.num_abs[nid] = ab.get(qn("w:val"))
            for ov in n.findall(qn("w:lvlOverride")):
                so = ov.find(qn("w:startOverride"))
                if so is not None: self.starts[(nid, int(ov.get(qn("w:ilvl"))))] = int(so.get(qn("w:val")))

    def prepare(self, root) -> None:
        counters = {}
        for p in root.iter(W + "p"):
            ppr = p.find(W + "pPr")
            numpr = ppr.find(W + "numPr") if ppr is not None else None
            if numpr is None: continue
            nid = _val(numpr, "w:numId", "0")
            ilvl = int(_val(numpr, "w:ilvl", "0") or 0)
            levels = self.levels.get(self.num_abs.get(nid, ""), {})
            if nid == "0" or ilvl not in levels: continue
            cnt = counters.setdefault(nid, {})
            start = self.starts.get((nid, ilvl), levels[ilvl][2])
            cnt[ilvl] = cnt.get(ilvl, start - 1) + 1
            for deeper in [k for k in cnt if k > ilvl]: del cnt[deeper]
            fmt, txt, _ = levels[ilvl]
            if fmt == "none": continue
            if fmt == "bullet":
                ch = txt.strip()
                label = ch if ch and ch.isprintable() and ord(ch[0]) < 0xE000 else "-"
            else:
                label = re.sub(r"%(\d)", lambda m: str(cnt.get(int(m.group(1)) - 1, 1)), txt)
            self.labels[p] = label

def _val(parent, tag: str, default: str = "") -> str:
    e = parent.find(qn(tag))
    v = e.get(qn("w:val")) if e is not None else None
    return v if v is not None else default

def para_text(p, numbering: SourceNumbering | None = None) -> str:
    t = p_text(p)
    if numbering is not None and t.strip():
        lab = numbering.labels.get(p)
        if lab: t = f"{lab} {t.lstrip()}"
    return t

def tc_lines(tc, numbering: SourceNumbering | None, join_fragments: bool) -> list[str]:
    frags = []
    for p in tc.findall(W + "p"): frags.extend(para_text(p, numbering).split("\n"))
    return reflow(frags, join_fragments)
    # ═══════════════════════════ بناء عناصر XML ═══════════════════════════
def el(tag: str, **attrs):
    e = OxmlElement(tag)
    for k, v in attrs.items(): e.set(qn("w:" + k), str(v))
    return e

def put(parent, child, order: list[str]):
    for old in parent.findall(child.tag): parent.remove(old)
    name = child.tag.split("}")[1]
    later = {W + n for n in order[order.index(name) + 1:]}
    for sib in parent:
        if sib.tag in later:
            sib.addprevious(child)
            return child
    parent.append(child)
    return child

def make_rpr(cfg: Settings, old=None, underline: bool = False):
    rpr = OxmlElement("w:rPr")
    put(rpr, el("w:rFonts", ascii=cfg.font, hAnsi=cfg.font, cs=cfg.font, eastAsia=cfg.font), RPR_ORDER)
    if cfg.bold:
        put(rpr, el("w:b"), RPR_ORDER)
        put(rpr, el("w:bCs"), RPR_ORDER)
    u_val = "single" if underline else None
    half = int(round(cfg.size_pt * 2))
    put(rpr, el("w:sz", val=half), RPR_ORDER)
    put(rpr, el("w:szCs", val=half), RPR_ORDER)
    if u_val: put(rpr, el("w:u", val=u_val), RPR_ORDER)
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
    if text: p.append(make_run(text, cfg, underline))
    return p

def make_spacer(cfg: Settings):
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
        if ch.tag != W + "tcPr": tc.remove(ch)

def ensure_trailing_paragraph(tc, cfg: Settings) -> None:
    if len(tc) == 0 or tc[-1].tag != W + "p": tc.append(make_spacer(cfg))

def fill_cell(tc, lines: list[str], cfg: Settings) -> None:
    clear_cell(tc)
    for line in lines:
        line = line.strip()
        if line: tc.append(make_p(line, cfg, underline=line in (LABEL_TEACHER, LABEL_LEARNER)))
    ensure_trailing_paragraph(tc, cfg)

@dataclass
class Ctx:
    cfg: Settings
    report: Report
    numbering: SourceNumbering
    src_part: object
    dst_doc: object
    nested_proto: object = None
    next_docpr: int = 5000

def _detach(e):
    p = e.getparent()
    if p is not None: p.remove(e)

def _unwrap(e):
    p = e.getparent()
    if p is None: return
    idx = p.index(e)
    for k, ch in enumerate(list(e)): p.insert(idx + k, ch)
    p.remove(e)

def copy_paragraph(src_p, ctx: Ctx):
    cfg = ctx.cfg
    p = copy.deepcopy(src_p)
    label = ctx.numbering.labels.get(src_p)
    
    # 1. إزالة العناصر المزعجة والروابط والصور التي تتلف الملف
    for tag in ("proofErr", "bookmarkStart", "bookmarkEnd", "commentRangeStart", "commentRangeEnd", "lastRenderedPageBreak", "permStart", "permEnd", "pPr", "del"):
        for e in list(p.iter(W + tag)): _detach(e)
    
    # 2. إزالة غلاف الروابط التشعبية مع الاحتفاظ بالنص المكتوب
    for tag in ("smartTag", "ins", "hyperlink"):
        for e in list(p.iter(W + tag)): _unwrap(e)
        
    # 3. حذف الصور والكائنات لأن نقلها بدون تحديث العلاقات يتلف الوورد
    for tag in ("drawing", "pict", "object"):
        for e in list(p.iter(W + tag)):
            run = next((a for a in e.iterancestors() if a.tag == W + "r"), None)
            _detach(run if run is not None else e)

    for tab in list(p.iter(W + "tab")):
        t = OxmlElement("w:t")
        t.text = " "
        tab.getparent().replace(tab, t)

    texts = list(p.iter(W + "t"))
    prev = None
    for t in texts:
        if t.text:
            s = re.sub(r"[ \t]{2,}", " ", t.text.replace("\u0640", "").replace("\u00a0", " "))
            t.text = re.sub(r"\s+([:،؛؟.])", r"\1", s)               
            if prev is not None and t.text[:1] in (":", "،", "؛", "؟", ".") and prev.text.endswith(" "): prev.text = prev.text.rstrip()                          
            prev = t
        t.set(XML_SPACE, "preserve")
        
    filled = [t for t in texts if (t.text or "").strip()]
    if filled and re.fullmatch(r"\*+\s*", filled[0].text or ""):
        filled[0].text = ""
        filled = filled[1:]
    elif filled:
        filled[0].text = re.sub(r"^\*+\s*", "", filled[0].text)
    if not filled: return None
    
    first, last = filled[0], filled[-1]
    for t in texts:
        if t is first: break
        t.text = ""
    for t in reversed(texts):
        if t is last: break
        t.text = ""
    first.text = first.text.lstrip()
    last.text = last.text.rstrip()

    for r in p.iter(W + "r"):
        old = r.find(W + "rPr")
        if old is not None: r.remove(old)
        r.insert(0, make_rpr(cfg, old))
    p.insert(0, make_ppr())
    if label: p.insert(1, make_run(label + " ", cfg))
    return p


def table_spec(tbl, numbering: SourceNumbering | None, cfg: Settings) -> dict:
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
                if gs is not None: span = int(gs.get(W + "val") or 1)
                v = pr.find(W + "vMerge")
                if v is not None: vm = v.get(W + "val") or "continue"
            row.append({"lines": tc_lines(tc, numbering, cfg.join_fragments), "span": span, "vm": vm})
        rows.append(row)
    ncols = sum(c["span"] for c in rows[0]) if rows else 0
    if len(grid) != ncols or sum(grid) == 0: grid = [1] * ncols
    return {"grid": grid, "rows": rows}

def make_table(proto, spec: dict, total: int, cfg: Settings):
    tbl = copy.deepcopy(proto)
    pr = tbl.find(W + "tblPr")
    if pr is None:
        pr = OxmlElement("w:tblPr")
        tbl.insert(0, pr)
    for tag in ("tblpPr", "tblOverlap"):          
        for e in pr.findall(W + tag): pr.remove(e)
    
    put(pr, el("w:tblW", w=total, type="dxa"), TBLPR_ORDER)
    put(pr, el("w:bidiVisual"), TBLPR_ORDER) # فرض الاتجاه RTL

    protos = proto.findall(W + "tr")
    if not protos or not spec["rows"]: return tbl
    hdr_tr = protos[0]
    body_tr = protos[1] if len(protos) > 1 else protos[0]
    for tr in tbl.findall(W + "tr"): tbl.remove(tr)

    g = spec["grid"] or [1]
    widths = [max(1, round(total * x / sum(g))) for x in g]
    widths[-1] += total - sum(widths)
    grid_el = tbl.find(W + "tblGrid")
    if grid_el is not None:
        for gc in list(grid_el): grid_el.remove(gc)
        for w_ in widths: grid_el.append(el("w:gridCol", w=w_))

    for ri, row in enumerate(spec["rows"]):
        is_hdr = ri == 0 and len(spec["rows"]) > 1
        tr = copy.deepcopy(hdr_tr if is_hdr else body_tr)
        pcells = tr.findall(W + "tc")
        for tc in pcells: tr.remove(tc)
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
                for e in tcpr.findall(W + tag): tcpr.remove(e)
            if span > 1: put(tcpr, el("w:gridSpan", val=span), TCPR_ORDER)
            if cell["vm"]: put(tcpr, el("w:vMerge", val="restart") if cell["vm"] == "restart" else el("w:vMerge"), TCPR_ORDER)
            for e in ptc.findall(W + "p") + ptc.findall(W + "tbl"): ptc.remove(e)
            for ln in (cell["lines"] or [""]): ptc.append(make_p(ln, cfg))           
            tr.append(ptc)
            col += span
        tbl.append(tr)
    return tbl

def copy_table_fallback(tbl, ctx: Ctx):
    new = copy.deepcopy(tbl)
    for old_p, new_p in zip(tbl.iter(W + "p"), new.iter(W + "p")):
        lab = ctx.numbering.labels.get(old_p)
        if lab: ctx.numbering.labels[new_p] = lab
    pr = new.find(W + "tblPr")
    if pr is not None:
        for tag in ("tblpPr", "tblOverlap"):
            for e in pr.findall(W + tag): pr.remove(e)
        put(pr, el("w:bidiVisual"), TBLPR_ORDER) # فرض الاتجاه RTL
    for tc in list(new.iter(W + "tc")):
        for p in list(tc.findall(W + "p")):
            q = copy_paragraph(p, ctx)
            tc.replace(p, q if q is not None else make_spacer(ctx.cfg))
    return new

def cell_inner_width(tc, fallback: int = 7700) -> int:
    pr = tc.find(W + "tcPr")
    w = pr.find(W + "tcW") if pr is not None else None
    try:
        if w is None or (w.get(W + "type") or "dxa") != "dxa": return fallback
        return max(1500, int(w.get(W + "w")) - 100)
    except (AttributeError, TypeError, ValueError): return fallback

def textbox_paragraphs(src_p) -> list:
    """فقرات المربعات النصية العائمة داخل فقرة (مثل «ملحوظة»)، حتى لا تضيع عند حذف الكائنات."""
    out, seen = [], set()
    for tb in src_p.iter(W + "txbxContent"):
        if any(a.tag == MC + "Fallback" for a in tb.iterancestors()): continue
        for q in tb.findall(W + "p"):
            txt = clean_text(p_text(q).replace("\n", " "))
            if not txt or is_title_line(txt) or txt in seen: continue
            seen.add(txt)
            out.append(q)
    return out

def convert_content(elements: list, ctx: Ctx, width: int) -> list:
    out: list = []
    prev_table = False
    for e in elements:
        if e.tag == W + "tbl":
            spec = table_spec(e, ctx.numbering, ctx.cfg)
            new = make_table(ctx.nested_proto, spec, width, ctx.cfg) if ctx.nested_proto is not None else copy_table_fallback(e, ctx)
            if prev_table: out.append(make_spacer(ctx.cfg))     
            out.append(new)
            prev_table = True
        elif e.tag == W + "p":
            p = copy_paragraph(e, ctx)
            if p is not None:
                out.append(p)
                prev_table = False
            for q in textbox_paragraphs(e):
                tp = copy_paragraph(q, ctx)
                if tp is not None:
                    out.append(tp)
                    prev_table = False
    return out
    # ═════════════════════════ الحصص واستخراج البيانات ═════════════════════════
COL_RULES = (("time", ("الحصيص", "الزمني")), ("eval", ("التقويم", "الدعم")), ("content", ("محتوي", "التعلمات", "المخرجات")), ("learner", ("المتعلم",)), ("teacher", ("المدرس", "الاستاذ", "المدخلات")), ("stage", ("المراحل", "المهارات")), ("segment", ("مقاطع", "المقطع")), ("cap", ("القدرات", "الكفايات", "الاهداف")))
READ_COLS = ("time", "eval", "content", "learner", "teacher", "stage", "segment", "cap", "activities")

def build_grid(tbl) -> list[list]:
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
                if gs is not None: span = int(gs.get(W + "val") or 1)
                v = pr.find(W + "vMerge")
                if v is not None: vm = v.get(W + "val") or "continue"
            for _ in range(span):
                if vm == "continue" and rows and len(row) < len(rows[-1]): row.append(rows[-1][len(row)])
                else: row.append(tc)
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
    if "teacher" not in cols and "learner" not in cols:        
        hit = next((i for i, h in enumerate(heads) if "الانشطه" in h and i not in cols.values()), None)
        if hit is not None: cols["activities"] = hit
    return cols

def _guess_header_rows(grid: list[list]) -> int:
    """عدد أسطر الرأس الفعلية.

    لا يكفي أن تكون الخلايا قصيرة: صفوف البيانات القصيرة كانت تُحسب رأساً
    فتُهمل أول صفوف المصدر. نختار أول عدد أسطر تُنتج أعمدة معروفة فعلاً.
    """
    for h in (1, 2, 3):
        if h >= len(grid):
            break
        cols = detect_columns(grid, h)
        if "content" in cols and ({"stage", "cap", "eval", "teacher", "learner", "activities"} & set(cols)):
            return h
    return 1

def _has_visible(els: list) -> bool:
    for e in els:
        if e.tag == W + "tbl" or next(e.iter(W + "drawing"), None) is not None: return True
        if e.tag == W + "p" and clean_text(p_text(e).replace("\n", " ")): return True
    return False

_ORD = {"الاولي": 1, "الثانيه": 2, "الثالثه": 3, "الرابعه": 4}
_TITLE = re.compile(r"الحصه(الاولي|الثانيه|الثالثه|الرابعه)انشطه")
SESSION_WORDS = (("اكتساب", 1), ("تطبيق", 2), ("انتاج", 3), ("تصحيح", 4))
SESSION_LABELS = {1: "الاكتساب", 2: "التطبيق", 3: "الإنتاج", 4: "التصحيح"}

def title_session(s: str):
    m = _TITLE.search(compact(s or ""))
    return _ORD[m.group(1)] if m else None

def is_title_line(s: str) -> bool:
    return title_session(s) is not None or norm(s) == "الاكتساب"

def kind_from_text(s: str):
    n = norm(s)
    hits = [(n.find(w), k) for w, k in SESSION_WORDS if w in n]
    return min(hits)[1] if hits else None

def _ordinal_sessions(text: str) -> set:
    """الحصص المذكورة بعناوين «الحصة الأولى/الثانية…» في النص (تدل على جذاذة متعددة الأنشطة)."""
    return {_ORD[m] for m in re.findall(r"الحصه\s*(الاولي|الثانيه|الثالثه|الرابعه)", norm(text)) if m in _ORD}


def source_kind(src_doc, source_path, rows) -> int | None:
    """نشاط الجذاذة الواحدة (1 اكتساب، 2 تطبيق، 3 إنتاج، 4 تصحيح)، أو None إن كانت متعددة الأنشطة.

    لا نستنتج «نشاطاً واحداً» من مجرد ورود كلمة في النص: هذا كان يحذف بقية
    الصفوف بصمت. الاعتماد على علامة صريحة (عنوان حصة أو اسم الملف) فقط،
    أو على صف واحد لا لبس فيه.
    """
    markers = {r.marker for r in rows if r.marker}
    if len(markers) >= 2:
        return None
    body_text = " ".join(clean_text(x) for x in all_texts(src_doc))
    if len(_ordinal_sessions(body_text)) >= 2:
        return None
    name_kind = kind_from_text(Path(source_path).stem)
    if name_kind:
        return name_kind
    if len(markers) == 1:
        return next(iter(markers))
    # ملف بصف واحد فقط: يمكن نسبته إلى نشاطه من عنوان المرحلة.
    if len(rows) == 1:
        only = rows[0]
        return kind_from_text(" ".join(only.stage) + " " + " ".join(only.teacher) + " " + " ".join(only.generic))
    return None


def role(stage: str) -> str:
    n = norm(stage)
    if "اخطاء" in n: return "other"
    if "تشخيص" in n or "تمهيد" in n: return "diag"
    if "فهم المطلوب" in n or "تقديم الدرس" in n: return "understand"
    if n.startswith("الانجاز") or n.startswith("انجاز"): return "do"
    return "other"

def tbox_text(e) -> str:
    return "".join(x.text or "" for x in e.iter(W + "t"))

def raw_lines(tc, numbering) -> list[str]:
    out: list[str] = []
    for p in tc.findall(W + "p"):
        for ln in para_text(p, numbering).split("\n"):
            ln = clean_text(ln)
            if ln and not is_title_line(ln): out.append(ln)
    return out

def content_els(tc) -> list:
    out = []
    for e in tc:
        if e.tag == W + "p":
            lines = [clean_text(x) for x in p_text(e).split("\n") if clean_text(x)]
            if lines and all(is_title_line(x) for x in lines): continue
            if not lines and is_title_line(tbox_text(e)): continue      
            out.append(e)
        elif e.tag == W + "tbl": out.append(e)
    return out

@dataclass
class SrcRow:
    idx: int = 0
    stage: list = field(default_factory=list)
    time: list = field(default_factory=list)
    caps: list = field(default_factory=list)
    teacher: list = field(default_factory=list)
    learner: list = field(default_factory=list)
    generic: list = field(default_factory=list)
    ev: list = field(default_factory=list)
    content: list = field(default_factory=list)
    diagnostic: list = field(default_factory=list)
    marker: int | None = None
    simple: bool = False
    first: bool = False

@dataclass
class Row:
    stage: list = field(default_factory=list)
    caps: list = field(default_factory=list)
    acts: list = field(default_factory=list)
    ev: list = field(default_factory=list)
    content: list = field(default_factory=list)
    diag: list = field(default_factory=list)

def _split_caps(lines: list[str]):
    groups: list[list[str]] = []
    for ln in lines:
        if re.match(r"^(?:ال)?قدره\b", norm(ln)) or not groups: groups.append([ln])
        else: groups[-1].append(ln)
    return (groups[0], groups[1]) if len(groups) == 2 else (lines, [])

def _split_run(lines: list[str]):
    n = 0
    while n < len(lines) and _starts_item(lines[n]): n += 1
    while 0 < n < len(lines) and not lines[n - 1].rstrip().endswith(_TERMINAL): n += 1
    return (lines[:n], lines[n:]) if 0 < n < len(lines) else (lines, [])

_SPLIT = re.compile(r"(?:تقويم\s+)?تشخيص\s+المكتسبات|تقويم\s+تشخيصي|(?<!\S)التمهيد")

def split_merged(sr: SrcRow, report: Report) -> list[SrcRow]:
    s = clean_text(" ".join(sr.stage))
    m = _SPLIT.search(s)
    if not m or not s[:m.start()].strip(): return [sr]
    a, b = copy.copy(sr), SrcRow()
    a.stage, b.stage = [s[:m.start()].strip()], [s[m.start():].strip()]
    a.time, b.time = sr.time[:1], sr.time[-1:] if len(sr.time) > 1 else []
    a.caps, b.caps = _split_caps(sr.caps)
    a.teacher, b.teacher = _split_run(sr.teacher)
    a.learner, b.learner = _split_run(sr.learner)
    a.marker, b.marker = None, sr.marker      # علامة «الحصة ..» تخص الحصة التي يبدؤها الجزء الثاني (التمهيد)
    report.warn(f"فُصل صف مدمج («{a.stage[0]}» + «{b.stage[0]}»)؛ راجع توزيع نصّيه.")
    return [a, b]

def _header_span(grid, hrows, idx):
    """أعمدة الشبكة التي يغطيها عنوان العمود idx (عنوان ممتد أفقياً يغطي عدة أعمدة)."""
    if idx is None or hrows < 1 or hrows - 1 >= len(grid): return [idx]
    hrow = grid[hrows - 1]
    hc = hrow[idx] if idx < len(hrow) else None
    if hc is None: return [idx]
    span = [j for j, c in enumerate(hrow) if c is hc]
    return span or [idx]


def _union_tc(cands: list):
    """خلية تجمع فقرات وجداول عدة خلايا متجاورة تحت عنوان واحد (بترتيبها)."""
    from docx.oxml import OxmlElement as _OE
    tc = _OE("w:tc")
    for c in cands:
        for e in c:
            if e.tag in (W + "p", W + "tbl"):
                tc.append(copy.deepcopy(e))
    return tc


def extract_full(grid, cols, hrows, numbering) -> list[SrcRow]:
    out, seen = [], set()
    spans = {k: _header_span(grid, hrows, v) for k, v in cols.items()}
    for ri in range(hrows, len(grid)):
        cells = grid[ri]
        def take(key: str):
            idx = cols.get(key)
            if idx is None: return None
            cand = []
            for j in spans.get(key, [idx]):
                if j >= len(cells): continue
                tc = cells[j]
                if tc is None or tc in seen or tc in cand: continue
                cand.append(tc)
            if not cand: return None
            seen.update(cand)
            return cand[0] if len(cand) == 1 else _union_tc(cand)
        t = {k: take(k) for k in READ_COLS}
        def rl(k: str) -> list[str]: return raw_lines(t[k], numbering) if t.get(k) is not None else []
        sr = SrcRow(stage=rl("stage"), time=rl("time"), caps=rl("cap"), teacher=rl("teacher"), learner=rl("learner"),
                    generic=rl("activities"), ev=rl("eval"), content=content_els(t["content"]) if t["content"] is not None else [])
        sr.marker = next((m for tc in dict.fromkeys(c for c in cells if c is not None) for m in [title_session(tbox_text(tc))] if m), None)
        if sr.stage or sr.caps or sr.teacher or sr.learner or sr.generic or sr.ev or _has_visible(sr.content): out.append(sr)
    return out

def extract_simple(tbl, numbering) -> list[SrcRow]:
    out, marker = [], None
    for ri, tr in enumerate(tbl.findall(W + "tr")):
        tcs = tr.findall(W + "tc")
        if len(tcs) < 2: continue
        m = title_session(" ".join(tbox_text(t) for t in tcs))
        if m: marker = m
        if ri == 0 and (m or "مراحلالدرس" in compact(tc_text(tcs[0]))): continue
        els = content_els(tcs[1])
        ti = next((i for i, e in enumerate(els) if e.tag == W + "tbl"), None)
        if ti is not None:
            cut = ti
            if cut > 0 and els[cut - 1].tag == W + "p":
                prev = clean_text(p_text(els[cut - 1]).replace("\n", " "))
                if len(prev) < 60 and not _starts_item(prev): cut -= 1
        else:
            first = norm(clean_text(p_text(els[0]).replace("\n", " "))) if els and els[0].tag == W + "p" else ""
            cut = 0 if ("المقترحه" in first or "الدعم" in first) else len(els)
        acts = [ln for e in els[:cut] if e.tag == W + "p" for ln in [clean_text(x) for x in para_text(e, numbering).split("\n")] if ln]
        sr = SrcRow(stage=raw_lines(tcs[0], numbering), teacher=acts, content=els[cut:], simple=True, first=not out, marker=marker if not out else None)
        out.append(sr)
    return out

def read_source_rows(src, numbering, report: Report) -> list[SrcRow]:
    rows: list[SrcRow] = []
    for tbl in src.element.body.findall(W + "tbl"):
        grid = build_grid(tbl)
        if len(grid) < 2: continue
        found = None
        for h in dict.fromkeys((_guess_header_rows(grid), 2, 1, 3)):
            if h >= len(grid): continue
            cols = detect_columns(grid, h)
            if "content" in cols and ({"teacher", "learner", "activities"} & set(cols)):
                found = (h, cols); break
        if found:
            for sr in extract_full(grid, found[1], found[0], numbering): rows.extend(split_merged(sr, report))
        elif max(len(r) for r in grid) <= 3:
            rows.extend(extract_simple(tbl, numbering))
        else: report.warn("جدول في المصدر لم أتعرف على أعمدته فتجاوزته.")
    if not rows: raise FatalError("لم أجد صفوفاً تحتوي بيانات في جداول المصدر.")
    
    for i, r in enumerate(rows): r.idx = i
    return rows

def assign_sessions(rows: list[SrcRow]) -> dict[int, list[SrcRow]]:
    out: dict[int, list[SrcRow]] = {1: [], 2: [], 3: [], 4: []}
    cur, has_rows, has_do = 0, set(), set()
    for sr in rows:
        rl = role(clean_text(" ".join(sr.stage)))
        had_rows = cur in has_rows
        had_do = cur in has_do
        if sr.marker:
            cur = sr.marker
        elif cur == 0:
            cur = 4 if sr.simple else 1
        elif sr.simple and sr.first:
            cur = min(cur + 1, 4)
        elif rl == "diag" and had_rows:
            cur = min(cur + 1, 4)
        elif rl == "understand" and had_do:
            cur = min(cur + 1, 4)
        has_rows.add(cur)
        if rl == "do":
            has_do.add(cur)
        out[cur].append(sr)
    return out

def make_row(sr: SrcRow, cfg: Settings) -> Row:
    st = re.sub(r"^\d+\s*[-–.)]\s*", "", clean_text(" ".join(sr.stage)))
    if norm(st) == "تشخيص المكتسبات": st = "تقويم تشخيصي"
    stage = [st] if st else []
    if role(st) == "diag" and "تمهيد" not in norm(st): stage.append("(تمهيد)")
    stage += [clean_text(x) for x in sr.time]
    teacher = reflow(sr.teacher, cfg.join_fragments)
    learner = reflow(sr.learner, cfg.join_fragments) if cfg.learner else []
    acts: list[str] = []
    if teacher: acts += ([LABEL_TEACHER] if cfg.labels else []) + teacher
    if learner: acts += ([LABEL_LEARNER] if cfg.labels else []) + learner
    if not acts: acts = reflow(sr.generic, cfg.join_fragments)
    
    diag: list[str] = sr.diagnostic.copy()
    if role(st) == "diag" and not diag: acts, diag = _split_diagnostic_answers(acts)
    
    return Row(stage=stage, caps=reflow(sr.caps, cfg.join_fragments), acts=acts, ev=reflow(sr.ev, cfg.join_fragments),
               content=sr.content, diag=diag)

def _split_diagnostic_answers(activities: list[str]):
    for i, line in enumerate(activities):
        n = norm(line)
        if "اجاب" in n and ("متعلم" in n or "متوقع" in n): return activities[:i], activities[i:]
    return activities, []
    # ═════════════════════════ ترويسة المصدر (فقرات خارج الجداول) ═════════════════════════
_LEAD = re.compile(r"^[\s*•·+\-–—]+")

def source_lines(doc) -> list[str]:
    out = []
    for p in iter_paragraphs(doc.element.body):
        if any(a.tag == W + "tbl" for a in p.iterancestors()): continue
        for ln in p_text(p).split("\n"):
            ln = clean_text(ln)
            if ln: out.append(ln)
    return out

def parse_header_lines(lines: list[str]) -> dict:
    info: dict = {"competencies": {}}
    last = None
    for raw in lines:
        s = _LEAD.sub("", raw)
        label, sep, value = s.partition(":")
        label, value, key = label.strip(), value.strip(), compact(label)
        words = label.split()
        if sep and key:
            last = None
            if compact(words[0]) == "الوحده":
                if len(words) > 1:
                    info.setdefault("unit", " ".join(words[1:]))
                    if value: info.setdefault("domain", re.sub(r"^مجال\s+", "", value))
                elif value: info.setdefault("unit", value)
            elif key.startswith("الكفايه") and len(key) > 7 and value:
                kind = key[7:]
                info["competencies"].setdefault(kind[2:] if kind.startswith("ال") else kind, value)
            elif key == "الفئهالمستهدفه": info.setdefault("target", value)
            elif key == "المكون": info.setdefault("component", value)
            elif key == "الموضوع":
                info.setdefault("theme", value); last = "theme"
            elif key == "المدهالزمنيهللحصه": info.setdefault("duration", value)
            elif key == "الحصيصالزمني": info.setdefault("total", value)
            elif key == "القدرهالخاصهبالدرس": info.setdefault("specific", value)
            elif key.startswith("الوسائل"):
                info.setdefault("means", value); last = "means"
        elif last == "means" and raw.lstrip().startswith(("+", "-", "•")):
            info["means"] = (info.get("means", "") + " " + raw.strip()).strip()
        elif last == "theme" and "title" not in info:
            info["title"] = re.sub(r"^\s*\d+\s*[-–.)]\s*", "", s).strip(); last = None
        else: last = None
    info["topic"] = info.get("title") or info.get("theme") or ""
    info["duration"] = info.get("duration") or info.get("total") or ""
    if info.get("means"):
        m = clean_text(re.sub(r"^[+\-•\s]+", "", info["means"]))
        if m.count("(") > m.count(")"):                       # قوس مفتوح بلا إغلاق في المصدر
            i = m.find(" + ", m.find("("))
            m = (m[:i] + ")" + m[i:]) if i != -1 else m.rstrip(".") + ")."
        info["means"] = m
    return info

def topic_from_filename(path: Path) -> str:
    s = re.sub(r"[_\-.\d()]+", " ", path.stem)
    s = re.sub(r"\b(?:جذاذة|جذاذه|درس|تربوية|رقم|مهارة)\b", " ", s)
    return re.sub(r"\s+", " ", s).strip() or "الدرس"

# ═════════════════════════ النموذج والتعبئة ═════════════════════════
MODEL_COLS = (("stage", "المراحل"), ("capabilities", "القدرات"), ("activities", "الانشطه"), ("content", "المحتوي"), ("evaluation", "التقويم"))

@dataclass
class Block:
    kind: int
    header: object
    main: object
    cols: dict

@dataclass
class Model:
    blocks: list
    nested_proto: object

def _suffix_kind(header, idx: int) -> int:
    for tr in header.findall(W + "tr"):
        tcs = tr.findall(W + "tc")
        for i, tc in enumerate(tcs):
            if compact(tc_text(tc)) == "المكون":
                v = value_cell(tcs, i)
                k = kind_from_text(tc_text(v)) if v is not None else None
                if k: return k
    return idx + 1

def analyze_model(doc, cfg: Settings) -> Model:
    headers, mains = [], []
    for t in doc.element.body.findall(W + "tbl"):
        head = " ".join(compact(tc_text(tc)) for tr in t.findall(W + "tr")[:3] for tc in tr.findall(W + "tc"))
        if "الوحده" in head and "المجال" in head: headers.append(t)
        elif "المراحل" in head and "المحتوي" in head: mains.append(t)
    if not headers or len(headers) != len(mains): raise FatalError("لم أتعرف على أزواج الترويسة في النموذج.")
    blocks = []
    for i, (h, m) in enumerate(zip(headers, mains)):
        trs = m.findall(W + "tr")
        if len(trs) < 2: raise FatalError("الجدول الرئيسي في النموذج لا يحوي صفاً نموذجياً.")
        cols: dict[str, int] = {}
        for ci, tc in enumerate(trs[0].findall(W + "tc")):
            ct = compact(tc_text(tc))
            for key, kw in MODEL_COLS:
                if key not in cols and kw in ct:
                    cols[key] = ci
                    break
        ncells = len(trs[1].findall(W + "tc"))
        for ci, (key, _) in enumerate(MODEL_COLS):
            if key not in cols: cols[key] = ci if ci not in cols.values() else next((k for k in range(ncells) if k not in cols.values()), ci)
        if len(set(cols.values())) < len(cols) or max(cols.values()) >= ncells: raise FatalError("تعذّر ربط أعمدة النموذج.")
        blocks.append(Block(_suffix_kind(h, i), h, m, cols))
    nested = [t for b in blocks for t in b.main.iter(W + "tbl") if t is not b.main]
    proto = copy.deepcopy(nested[min(cfg.nested_proto_index, len(nested) - 1)]) if nested else None
    return Model(blocks, proto)

def unfloat(tbl) -> None:
    pr = tbl.find(W + "tblPr")
    if pr is None or pr.find(W + "tblpPr") is None: return
    for tag in ("tblpPr", "tblOverlap"):
        for e in pr.findall(W + tag): pr.remove(e)
    if pr.find(W + "jc") is None: put(pr, el("w:jc", val="center"), TBLPR_ORDER)

def detect_size(doc, cfg: Settings) -> None:
    for t in doc.element.body.findall(W + "tbl"):
        for sz in t.iter(W + "sz"):
            v = sz.get(W + "val")
            if v and v.isdigit() and int(v) >= 14:
                cfg.size_pt = int(v) / 2
                return

def value_cell(tcs: list, i: int):
    j = i - 1
    while j >= 0 and tc_text(tcs[j]).strip() in (":", "：", "؛"): j -= 1
    return tcs[j] if j >= 0 else None

def fill_keep(tc, text: str, cfg: Settings) -> None:
    p0 = tc.find(W + "p")
    ppr = rpr = None
    if p0 is not None:
        ppr = p0.find(W + "pPr")
        runs = list(p0.iter(W + "r"))
        r0 = next((r for r in runs if "".join(t.text or "" for t in r.iter(W + "t")).strip()), runs[0] if runs else None)
        rpr = r0.find(W + "rPr") if r0 is not None else None
    clear_cell(tc)
    p = OxmlElement("w:p")
    p.append(copy.deepcopy(ppr) if ppr is not None else make_ppr())
    r = OxmlElement("w:r")
    r.append(copy.deepcopy(rpr) if rpr is not None else make_rpr(cfg))
    t = OxmlElement("w:t")
    t.set(XML_SPACE, "preserve")
    t.text = text
    r.append(t)
    p.append(r)
    tc.append(p)

_LEVEL = re.compile(r"(الاولي|الثانيه|الثالثه)")
def _level(s: str):
    m = _LEVEL.search(norm(s))
    return m.group(1) if m else None

def fill_header(blk: Block, hdr: dict, cfg: Settings, report: Report, announce: bool) -> None:
    vals = {"الوحده": hdr.get("unit"), "المجال": hdr.get("domain"), "الموضوع": hdr.get("topic"), "مدهالانجاز": hdr.get("duration")}
    if MEANS_FIELD == "means": vals["الوسائلالتعليميه"] = hdr.get("means")
    else: vals["المرجع"] = hdr.get("means")
    filled, missing = [], []
    for tr in blk.header.findall(W + "tr"):
        tcs = tr.findall(W + "tc")
        for i, tc in enumerate(tcs):
            ck = compact(tc_text(tc))
            vtc = value_cell(tcs, i) if ck else None
            if vtc is None: continue
            if ck == "الفئهالمستهدفه":
                a, b = _level(tc_text(vtc)), _level(hdr.get("target", ""))
                if a and b and a != b: report.warn(f"مستوى المصدر ({b}) يخالف النموذج ({a}).")
            elif ck == "المكون":
                comp = clean_text(hdr.get("component", ""))
                m = re.search(r"حصة\s+(\S+)", tc_text(vtc))
                suffix = m.group(1) if m else SESSION_LABELS[blk.kind]
                if comp: fill_keep(vtc, f"{comp} – حصة {suffix}", cfg)
            elif ck in vals:
                val = clean_text(vals[ck] or "")
                if val:
                    fill_keep(vtc, val, cfg)
                    filled.append(f"{ck}={val}")
                else: missing.append(ck)
    if announce:
        if filled: report.ok("الترويسة: " + " | ".join(filled))
        if missing: report.warn("لم أجد في المصدر: " + "، ".join(missing))

def fill_blank_labels(blocks: list) -> None:
    """يملأ خانة عنوان فارغة في ترويسة حصة بنسخها من خانة الموضع نفسه في الحصة الأولى."""
    if len(blocks) < 2: return
    ref = blocks[0].header.findall(W + "tr")
    for blk in blocks[1:]:
        for ri, tr in enumerate(blk.header.findall(W + "tr")):
            if ri >= len(ref): break
            tcs, rtcs = tr.findall(W + "tc"), ref[ri].findall(W + "tc")
            if len(tcs) != len(rtcs): continue
            for i, (tc, rtc) in enumerate(zip(tcs, rtcs)):
                if tc_text(tc).strip() or not tc_text(rtc).strip(): continue
                if i == 0 or tc_text(tcs[i - 1]).strip() not in (":", "：", "؛"): continue
                clear_cell(tc)
                for pp in rtc.findall(W + "p"): tc.append(copy.deepcopy(pp))

def _objective_value(kind: str, hdr: dict):
    comps = hdr["competencies"]
    if comps.get(kind): return comps[kind]
    if kind == "معرفيه": return hdr.get("specific") or comps.get("ثقافيه")
    return None

def fill_objectives(blk: Block, hdr: dict, cfg: Settings, report: Report, announce: bool) -> None:
    trs = blk.header.findall(W + "tr")
    start = next((ri for ri, tr in enumerate(trs) for tc in tr.findall(W + "tc") if compact(tc_text(tc)).startswith("الاهداف")), None)
    if start is None: return
    done = 0
    for tr in trs[start + 1:]:
        first = tr.find(W + "tc")
        txt = tc_text(first) if first is not None else ""
        head = txt.split(":", 1)[0]
        ck = compact(head)
        if ":" not in txt or not ck.startswith("كفايه"): break
        value = _objective_value(ck[5:], hdr)
        if value:
            fill_keep(first, f"- {head.strip(' -–•*')}: {clean_text(value)}", cfg)
            done += 1
    if done and announce: report.ok(f"الأهداف التعلمية: حُدّث {done} من المصدر (الحصص {', '.join(map(str, OBJECTIVE_SESSIONS))})")

def insert_diagnostic(new_els: list, lines: list[str], cfg: Settings) -> None:
    new_els[:0] = [make_p(ln, cfg) for ln in lines]

def fill_main(blk: Block, rows: list[Row], ctx: Ctx) -> None:
    cfg, c = ctx.cfg, blk.cols
    trs = blk.main.findall(W + "tr")
    template = copy.deepcopy(trs[1])
    for tr in trs[1:]: blk.main.remove(tr)
    for row in rows:
        tr = copy.deepcopy(template)
        blk.main.append(tr)
        tcs = tr.findall(W + "tc")
        fill_cell(tcs[c["stage"]], row.stage, cfg)
        fill_cell(tcs[c["capabilities"]], row.caps, cfg)
        fill_cell(tcs[c["activities"]], row.acts, cfg)
        fill_cell(tcs[c["evaluation"]], row.ev, cfg)
        ctc = tcs[c["content"]]
        clear_cell(ctc)
        new_els = convert_content(row.content, ctx, cell_inner_width(ctc))
        if row.diag: insert_diagnostic(new_els, row.diag, cfg)
        for e in new_els: ctc.append(e)
        ensure_trailing_paragraph(ctc, cfg)

def remove_block(body, blk: Block, stop) -> None:
    cur = blk.header
    while cur is not None and cur is not stop and cur.tag != W + "sectPr":
        nxt = cur.getnext()
        body.remove(cur)
        cur = nxt

def trim_trailing_breaks(body) -> None:
    kids = [e for e in body if e.tag != W + "sectPr"]
    while kids and kids[-1].tag == W + "p" and not p_text(kids[-1]).strip() and any(b.get(W + "type") == "page" for b in kids[-1].iter(W + "br")):
        body.remove(kids.pop())
        # ═════════════════════════ التشغيل والحفظ ═════════════════════════
def find_file(name: str) -> Path:
    p = Path(name).expanduser()
    if p.is_file(): return p
    def key(s: str) -> str: return compact(re.sub(r"[_\-\s()]+", "", unicodedata.normalize("NFC", s)))
    target = key(p.name)
    print(f"جاري البحث عن «{name}»...")
    search_bases = [Path.cwd(), Path(__file__).resolve().parent, Path("/storage/emulated/0")]
    for base in search_bases:
        if not base.exists(): continue
        direct = base / p.name
        if direct.is_file(): return direct
        depth = len(base.parts)
        for root, dirs, files in os.walk(base):
            dirs[:] = [d for d in dirs if not d.startswith(".") and d not in (OUT_DIR, "node_modules", "__pycache__", "Android")]
            for f in files:
                if f.lower().endswith(".docx") and not f.startswith("~$") and key(f) == target: return Path(root, f)
    raise FatalError(f"لم أجد الملف «{name}».")

def output_path(cfg: Settings, hdr: dict) -> Path:
    topic = hdr.get("topic", "الدرس")
    name = cfg.out_name or f"جذاذة_{topic}_الاحترافية.docx"
    name = re.sub(r'[<>:"/\\|?*\r\n\t]+', " ", name).strip(" .")
    name = re.sub(r"\s+", "_", name)
    if not name.lower().endswith(".docx"): name += ".docx"
    
    comp = hdr.get("component", "مكون غير محدد").strip() or "مكون غير محدد"
    unit = hdr.get("unit", "وحدة غير محددة").strip() or "وحدة غير محددة"
    comp = re.sub(r'[<>:"/\\|?*\r\n\t]+', "", comp).strip()
    unit = re.sub(r'[<>:"/\\|?*\r\n\t]+', "", unit).strip()
    
    target_dir = cfg.out_dir / comp / unit
    target_dir.mkdir(parents=True, exist_ok=True)
    
    path, stem, n = target_dir / name, Path(name).stem, 2
    while path in cfg.used_outputs:
        path = target_dir / f"{stem}_{n}.docx"
        n += 1
    cfg.used_outputs.add(path)
    return path

def save_safely(doc, path: Path) -> Path:
    try: doc.save(str(path)); return path
    except PermissionError:
        for i in range(2, 100):
            alt = path.with_name(f"{path.stem} ({i}){path.suffix}")
            try: doc.save(str(alt)); return alt
            except PermissionError: continue
        raise

def build(model_path: Path, source_path: Path, cfg: Settings, report: Report) -> Path:
    say(f"النموذج: {model_path}\nالمصدر : {source_path}")
    doc, src = Document(str(model_path)), Document(str(source_path))
    model = analyze_model(doc, cfg)
    detect_size(doc, cfg)
    for blk in model.blocks:
        unfloat(blk.header)
        unfloat(blk.main)

    numbering = SourceNumbering(src)
    numbering.prepare(src.element.body)
    hdr = parse_header_lines(source_lines(src))
    if not hdr.get("topic"): hdr["topic"] = topic_from_filename(source_path)

    rows = read_source_rows(src, numbering, report)
    
    if cfg.use_ai and cfg.api_key:
        ai_refine_insha(rows, hdr, cfg, report)
    elif cfg.use_ai:
        say("تنبيه: الذكاء الاصطناعي معطّل لعدم توفر المفتاح.")
        
    sessions = assign_sessions(rows)
    kind = source_kind(src, source_path, rows)
    if kind is not None:
        sessions = {1: [], 2: [], 3: [], 4: []}
        sessions[kind] = rows                        # جذاذة نشاط واحد: تُعبأ حصته وحدها، وتُحذف بقية الحصص من النموذج
        report.info(f"الجذاذة نشاط واحد: {SESSION_LABELS[kind]}")
    ctx = Ctx(cfg, report, numbering, src.part, doc, model.nested_proto)

    body, drop, first = doc.element.body, [], True
    for i, blk in enumerate(model.blocks):
        srows = [r for r in (make_row(s, cfg) for s in sessions.get(blk.kind, [])) if r.stage or r.caps or r.acts or r.ev or r.content]
        if not srows:
            report.warn(f"لا توجد في المصدر حصة «{SESSION_LABELS[blk.kind]}»؛ حُذف قسمها من الناتج.")
            drop.append((blk, model.blocks[i + 1].header if i + 1 < len(model.blocks) else None))
            continue
        fill_header(blk, hdr, cfg, report, first)
        if blk.kind in OBJECTIVE_SESSIONS: fill_objectives(blk, hdr, cfg, report, first)
        first = False
        fill_main(blk, srows, ctx)
        report.info(f"حصة {SESSION_LABELS[blk.kind]}: {len(srows)} صفوف")
    fill_blank_labels(model.blocks)
    for blk, stop in drop: remove_block(body, blk, stop)
    if drop: trim_trailing_breaks(body)
    extra = [k for k in sessions if sessions[k] and k not in {b.kind for b in model.blocks}]
    if extra: report.warn("حصص في المصدر بلا قسم مقابل في النموذج: " + "، ".join(SESSION_LABELS[k] for k in extra))

    if cfg.margin_twips:
        for section in doc.sections:
            section.top_margin = section.bottom_margin = section.left_margin = section.right_margin = Twips(cfg.margin_twips)
    doc.core_properties.title = f"جذاذة {hdr['topic']}"
    out = save_safely(doc, output_path(cfg, hdr))

    report.ok(f"تم إنجاز الجذاذة بنجاح! الملف: {out}")
    return out

def parse_args(argv):
    ap = argparse.ArgumentParser()
    ap.add_argument("model", nargs="?", default=DEFAULT_MODEL)
    ap.add_argument("sources", nargs="*", default=[DEFAULT_SOURCE])
    ap.add_argument("-o", "--out-dir", default=OUT_DIR)
    ap.add_argument("--out-name", default="")
    ap.add_argument("--font", default="Sakkal Majalla")
    ap.add_argument("--margin", type=int, default=0)
    ap.add_argument("--no-reflow", action="store_true")
    ap.add_argument("--no-labels", action="store_true")
    ap.add_argument("--no-learner", action="store_true")
    ap.add_argument("-v", "--verbose", action="store_true")
    return ap.parse_args(argv)

def main(argv=None) -> int:
    args = parse_args(argv)
    cfg = Settings(font=args.font, margin_twips=args.margin, join_fragments=not args.no_reflow,
                   labels=not args.no_labels, learner=not args.no_learner,
                   out_dir=Path(args.out_dir), out_name=args.out_name, verbose=args.verbose)
    cfg.api_key = os.environ.get("GEMINI_API_KEY", "")   # المفتاح من البيئة فقط
    try: model_path = find_file(args.model)
    except FatalError as e:
        print(f"✖ {e}"); return 1
    failures, sources = 0, []
    for s in (args.sources or [DEFAULT_SOURCE]):
        try: sources.append(find_file(s))
        except FatalError as e:
            print(f"✖ {e}"); failures += 1
    for src in sources:
        report = Report()
        try: build(model_path, src, cfg, report)
        except Exception as e:
            report.error(f"خطأ: {e}")
            if cfg.verbose: traceback.print_exc()
            failures += 1
        report.show()
    return 1 if failures else 0

if __name__ == "__main__":
    sys.exit(main())