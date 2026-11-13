# -*- coding: utf-8 -*-
"""
reading_general.py — قارئ عام لجذاذات «القراءة» بأي تخطيط (أي عدد أعمدة، أو بلا جداول).

الفكرة: (1) قراءة المصدر إلى بنية وسيطة (ترويسة + مراحل)  (2) تعبئة قالب «المنفى» منها.
  - الأعمدة تُصنَّف من عناوينها لا من مواضعها (مرادفات: المحتوى/المضامين/المادة…).
  - المراحل تُصنَّف قانونياً (تشخيص، تأطير، فرضية، قراءة، فهم، تحليل، تركيب، استثمار).
  - لا يُحذف نص من المصدر؛ ما لا موضع له يُلحق بأقرب مرحلة ويُذكر في التقرير.
  - الذكاء الاصطناعي اختياري ويملأ الفراغات فقط (كما في reading.py).
"""
import copy
import re
from pathlib import Path

from docx.oxml.ns import qn

import reading as R

W = R.W
ROLES = ("time", "ability", "teacher", "learner", "activities", "content", "eval", "support", "indic", "notes")


# ═══════════════════════════ تطبيع عربي ═══════════════════════════
def plain(s):
    """نص بلا تطويل/تشكيل/مسافات زائدة (للمقارنة والبحث عن العناوين)."""
    s = (s or "").replace("\u00a0", " ")
    s = re.sub("[\u0640\u064B-\u0652\u0670]", "", s)
    return re.sub(r"[ \t\r\f\v]+", " ", s).strip()


def norm(s):
    """تطبيع أشدّ: همزات/ياء/تاء مربوطة وحذف الترقيم."""
    s = plain(s)
    s = re.sub("[أإآٱ]", "ا", s).replace("ى", "ي").replace("ة", "ه")
    return re.sub(r"\s+", " ", re.sub(r"[^\w\s]", " ", s)).strip()


# ═══════════════════════════ قراءة العناصر إلى أسطر ═══════════════════════════
def element_lines(el, keep_tables=False):
    """أسطر نص عنصر (فقرة/جدول متداخل/صندوق محتوى) مع نصوص المربعات العائمة داخلها."""
    if el.tag == W + "p":
        buf = []
        for r in el.iter(W + "r"):
            if any(a.tag in (W + "txbxContent",) for a in r.iterancestors()):
                continue
            for ch in r:
                if ch.tag == W + "t" and ch.text:
                    buf.append(ch.text)
                elif ch.tag == W + "tab":
                    buf.append(" ")
                elif ch.tag in (W + "br", W + "cr"):
                    buf.append("\n")
                elif ch.tag in (W + "drawing", W + "pict"):
                    if buf:
                        yield "".join(buf)
                        buf = []
                    seen = set()
                    for tb in ch.iter(W + "txbxContent"):
                        if any(a.tag.endswith("}Fallback") for a in tb.iterancestors()):
                            continue
                        for blk in tb:
                            for ln in element_lines(blk, keep_tables):
                                if ln not in seen:
                                    seen.add(ln)
                                    yield ln
        if buf:
            yield "".join(buf)
    elif el.tag == W + "tbl" and keep_tables and R.is_content_table(el):
        yield R.register_table(el)
    elif el.tag == W + "tbl":
        for tr in el.findall(W + "tr"):
            texts = []
            for tc in tr.findall(W + "tc"):
                ls = [x for blk in tc for x in element_lines(blk, keep_tables)]
                t = " ".join(ls).strip()
                if t:
                    texts.append(t)
            if texts:
                yield " : ".join(texts) if len(texts) > 1 else texts[0]
    elif el.tag == W + "sdt":
        sc = el.find(f".//{W}sdtContent")
        if sc is not None:
            for blk in sc:
                yield from element_lines(blk, keep_tables)


def cell_lines(tc):
    out = []
    for blk in tc:
        for ln in element_lines(blk, keep_tables=True):
            for x in ln.split("\n"):
                x = R.clean(plain(x))
                if x:
                    out.append(x)
    return out


def body_lines(body, skip_tbls=()):
    out = []
    for blk in body:
        if blk.tag == W + "tbl" and any(blk is t for t in skip_tbls):
            continue
        for ln in element_lines(blk):
            for x in ln.split("\n"):
                x = plain(x)
                if x:
                    out.append(x)
    return out


# ═══════════════════════════ شبكة الجدول ═══════════════════════════
class GCell:
    __slots__ = ("lines", "r0", "r1", "c0", "c1")

    def __init__(self, lines, r0, c0, c1):
        self.lines, self.r0, self.r1, self.c0, self.c1 = lines, r0, r0, c0, c1


class Grid:
    def __init__(self, tbl_el):
        self.el = tbl_el
        self.cells, open_vm, ri = [], {}, -1
        for ri, tr in enumerate(tbl_el.findall(W + "tr")):
            ci = 0
            trPr = tr.find(W + "trPr")
            if trPr is not None:
                gb = trPr.find(W + "gridBefore")
                if gb is not None:
                    ci += int(gb.get(W + "val", "0") or 0)
            for tc in tr.findall(W + "tc"):
                tcPr, span, vm = tc.find(W + "tcPr"), 1, None
                if tcPr is not None:
                    gs = tcPr.find(W + "gridSpan")
                    if gs is not None:
                        span = int(gs.get(W + "val", "1") or 1)
                    v = tcPr.find(W + "vMerge")
                    if v is not None:
                        vm = v.get(W + "val", "continue")
                if vm == "continue" and ci in open_vm:
                    open_vm[ci].r1 = ri
                    ci += span
                    continue
                c = GCell(cell_lines(tc), ri, ci, ci + span - 1)
                for k in range(ci, ci + span):
                    if vm == "restart":
                        open_vm[k] = c
                    else:
                        open_vm.pop(k, None)
                self.cells.append(c)
                ci += span
        self.nrows = ri + 1
        self.ncols = (max(c.c1 for c in self.cells) + 1) if self.cells else 0

    def at(self, r, c):
        for x in self.cells:
            if x.r0 <= r <= x.r1 and x.c0 <= c <= x.c1:
                return x
        return None

    def row_cells(self, r):
        return [x for x in self.cells if x.r0 == r]

    def text(self):
        return " ".join(" ".join(c.lines) for c in self.cells)


# ═══════════════════════════ تصنيف الأعمدة ═══════════════════════════
def classify(label):
    """دور العمود من عنوانه، أو None."""
    n = norm(label)
    if not n or len(n) > 70:
        return None
    if "وسائل" in n or "فئه" in n or "مجال" in n:
        return None
    if "مؤشر" in n:
        return "indic"
    if "تقويم" in n or "تقييم" in n:
        return "eval"
    if "دعم" in n:
        return "support"
    if re.search(r"محتو|مضامين|مضمون", n):
        return "content"
    if re.search(r"ماده (ال)?علمي", n):
        return "teacher"
    if re.search(r"اهداف|قدرات|كفايات|مهارات محقق", n):
        return "ability"
    # «الأنشطة التعليمية» و«الأنشطة التعليمية التعلمية» عمود أنشطة، وليس عمود مدرس.
    if re.search(r"انشط", n):
        if re.search(r"متعلم|تلميذ", n):
            return "learner"
        if re.search(r"مدرس|استاذ|معلم", n):
            return "teacher"
        return "activities"
    if re.search(r"مدرس|استاذ|معلم", n):
        return "teacher"
    if re.search(r"متعلم|تلميذ", n):
        return "learner"
    if re.search(r"مراحل|مقاطع|خطوات|محطات|فقرات", n):
        return "stage"
    if re.search(r"زمن|حصيص|مده|وقت", n):
        return "time"
    if re.search(r"ملحوظ|ملاحظ", n):
        return "notes"
    return None


def _header_row_score(grid, r):
    """عدد الأعمدة التي يحمل عنوانها دوراً معروفاً في هذا الصف."""
    return sum(1 for c in grid.row_cells(r) if c.lines and classify(" ".join(c.lines)))


def _is_header_like(grid, r):
    """صف عناوين فعلية: كل خلاياه غير الفارغة قصيرة، وليست عناوين مراحل/بيانات."""
    cells = [c for c in grid.row_cells(r) if c.lines]
    if not cells:
        return False
    if any(len(" ".join(c.lines)) >= 30 for c in cells):
        return False
    if any(canon(" ".join(c.lines)) for c in cells):
        return False
    return True


def find_header(grid, max_scan=6):
    """(سطر بداية العناوين، عدد سطور العناوين، {عمود: دور}) أو None.

    القاعدة: صف العناوين هو الصف الأكثر احتواءً على أدوار معروفة في أول الصفوف.
    لا نضمّ صفوف البيانات إلى الرأس إلا إذا كان الرأس ناقصاً فعلاً (خلايا مدمجة)
    وكان الصف التالي صفَّ عناوين خالصاً (قصير، بلا عنوان مرحلة).
    """
    scan = range(min(max_scan, grid.nrows))
    scores = [(r, _header_row_score(grid, r)) for r in scan]
    start, best = max(scores, key=lambda x: (x[1], -x[0]), default=(None, 0))
    if start is None or best < 2:
        return None
    roles = {c: None for c in range(grid.ncols)}
    for c in range(grid.ncols):
        cell = grid.at(start, c)
        if cell and cell.lines:
            roles[c] = classify(" ".join(cell.lines))

    # رأس من سطرين فقط عند وجود أعمدة بلا دور، وبشرط أن يكون السطر التالي عناوين خالصة.
    n = 1
    while start + n < grid.nrows and n < 3:
        missing = any(role is None for role in roles.values())
        nxt = start + n
        if missing and _is_header_like(grid, nxt) and _header_row_score(grid, nxt) >= 2:
            for c in range(grid.ncols):
                cell = grid.at(nxt, c)
                if cell and cell.lines:
                    role = classify(" ".join(cell.lines))
                    if role:
                        roles[c] = role
            n += 1
        else:
            break
    return start, n, roles


# ═══════════════════════════ المراحل القانونية ═══════════════════════════
STAGE_ORDER = ["diag", "frame", "hyp", "read", "und", "ana", "syn", "inv"]
STAGE_TITLE = {"diag": "تشخيص المكتسبات", "frame": "تأطير النص", "hyp": "بناء الفرضية", "read": "القراءة",
               "und": "الفهم", "ana": "التحليل", "syn": "التركيب", "inv": "الاستثمار"}
STAGE_CUES = [
    ("hyp", r"فرضي"),
    ("diag", r"تشخيص|تمهيد|ما قبل القراءه|استدعاء|تهيئ|تهييء|مراقبه الاعداد|الاعداد القبلي"),
    ("frame", r"تاطير|ملاحظه النص|عتبات|الاستكشافيه|ملاحظه"),
    ("read", r"انشطه القراءه|القراءه النموذجيه|القراءه المسترسله|^قراءه|شرح المفردات|الايضاح اللغوي|المعجم"),
    ("und", r"فهم|الفكره (العامه|المحوريه)|الحدث المحوري|توجيهيه"),
    ("ana", r"تحليل|خصائص|الحقول|الاساليب|البنيه"),
    ("syn", r"تركيب|خلاصه|استنتاج|تجميع"),
    ("inv", r"استثمار|توظيف|انجاز|تطبيق|الدرس المقبل|تقويم الدرس|امتداد"),
]


def canon(label):
    n = norm(label)
    if not n:
        return None
    for key, pat in STAGE_CUES:
        if re.search(pat, n):
            return key
    return None


def item_cue(text, role=None):
    """ترجيح مرحلة لبند (للمحاذاة في جذاذات «القوائم داخل خلية واحدة»)؛ يُنظر إلى مطلع البند/عنوانه فقط."""
    t = R.strip_marker(text)
    lab = t.split(":", 1)[0] if ":" in t and len(t.split(":", 1)[0].split()) <= 6 else " ".join(t.split()[:7])
    n = norm(lab)
    probes = [
        ("hyp", r"فرضي|نفترض"),
        ("diag", r"الاعداد القبلي|مراقبه الاعداد|استدعاء|تشخيص"),
        ("frame", r"عتبات|صاحب النص|مصدر النص|دلاله العنوان|دراسه العنوان|دلاله الصوره|قراءه الصوره|نوعيه النص|تاطير|تعريف صاحب|ملاحظه"),
        ("read", r"قراءه (الاستاذ|الاستاذه|التلاميذ|المتعلمين|المتعلمون|نموذجيه)|بقراءه النص|النص المسترسل|الشروح|شرح (مفرداته|المفردات)|المفردات|الايضاح اللغوي"),
        ("und", r"الفكره (العامه|المحوريه)|الفهم|فهمهم|صحه الفرضيه|^الحدث|القضايا المعالجه|استخلاص"),
        ("ana", r"خصائص|استخراج|تحليل|القيم|الشخصيات|الحقول|الاحداث|الزمان|المكان|الامكنه|علاقه"),
        ("syn", r"تركيب|خلاصه|استنتاج"),
        ("inv", r"الدرس المقبل|استثمار|انجاز انشطه|توظيف|انشطه النص|بحث"),
    ]
    for key, pat in probes:
        if re.search(pat, n):
            return key
    return None


# ═══════════════════════════ الوحدة المرحلية ═══════════════════════════
class Unit:
    def __init__(self, label="", key=None):
        self.label, self.key = label, key
        self.f = {r: [] for r in ROLES}
        self.order = []

    def add(self, role, items):
        seen = {R.key(x) for x in self.f[role]}
        for it in items:
            if it and R.key(it) not in seen:
                self.f[role].append(it)
                seen.add(R.key(it))

    def size(self):
        return sum(len(v) for v in self.f.values())


# ═══════════════════════════ الترويسة ═══════════════════════════
FIELD_LABELS = [
    ("level", r"الفئة المستهدفة|الفئة|المستوى"),
    ("duration", r"الحصيص الزمني|المدة الزمنية|مدة الانجاز|مدة الإنجاز|المدة|الزمن"),
    ("materials", r"الوسائل الديداكتيكية|الوسائل الديدكتيكية|الوسائل التعليمية|الوسائل الديداكتيكيه|الوسائل|الدعامات"),
    ("reference", r"المرجع|المراجع"),
    ("domain", r"المجال"),
    ("component", r"المكون"),
    ("topic", r"الموضوع|عنوان الدرس|النص"),
    ("subject", r"المادة"),
    ("objective", r"الهدف التعلمي|الهدف العام|الهدف|الأهداف|الاهداف"),
    ("teacher", r"الأستاذة|الأستاذ|الاستاذ|الاستاذة"),
    ("school", r"المؤسسة"),
    ("_sep", r"الكفايات المستهدفة|الكفايات المتوخاة|الكفايات|رقم الجذاذة|حكمة الأسبوع|الحصة|الأنشطة"),
]
COMP_WORDS = r"تواصلية|التواصلية|منهجية|المنهجية|ثقافية|الثقافية|استراتيجية|الاستراتيجية|إستراتيجية|الإستراتيجية|تكنولوجية|التكنولوجية|معرفية|المعرفية"


def _label_re():
    parts = []
    for fld, pat in FIELD_LABELS:
        parts.append(f"(?P<{fld}>(?:^|(?<=[\\s\\-–—ـ]))(?:{pat})(?=\\s*[:：]))")
    parts.append(f"(?P<comp>(?:^|(?<=[\\s\\-–—ـ*]))(?:(?:ال)?كفاية\\s+)?(?:{COMP_WORDS})(?=\\s*[:：]))")
    return re.compile("|".join(parts))


LABEL_RE = _label_re()


def looks_reversed(s):
    toks = ["المستهدفة", "الكفايات", "الوحدة", "المجال", "الموضوع", "الفئة"]
    rev = sum(1 for t in toks if t[::-1] in s)
    fwd = sum(1 for t in toks if t in s)
    return rev > fwd


_STOP = r"(?=\s+(?:المكون|مكون|الموضوع|الحصيص|المدة|مدة|الوسائل|الفئة|المستوى|الأستاذ|الاستاذ|الأستاذة|المؤسسة|المادة|رقم|الكفايات|الهدف)\b|\s{2,}|$)"
_UNIT_RES = [
    re.compile(r"الوحدة\s*(?:الدراسية|التعلمية)?\s*[:：]\s*(\S+?)\s*[–\-:：]?\s*(?:مجال\s+(.+?))?" + _STOP),
    re.compile(r"الوحدة\s+(\S+?)\s*[:：–\-]\s*(?:مجال\s+)?(.+?)" + _STOP),
]
_ORD = r"(?:ال)?(?:اولى|أولى|الأولى|ثانية|ثالثة|رابعة|خامسة|سادسة|\d+)"


def extract_unit(s):
    for rx in _UNIT_RES:
        m = rx.search(s)
        if m and re.fullmatch(_ORD, m.group(1).strip("–-: ")):
            dom = (m.group(2) or "").strip(" .–-:")
            cut = LABEL_RE.search(dom)
            if cut:
                dom = dom[:cut.start()].strip(" .–-:")
            return m.group(1).strip("–-: "), re.sub(r"^مجال\s+", "", dom), (s[:m.start()] + " " + s[m.end():]).strip()
    return None, None, s


def _tidy(fld, val):
    val = re.sub(r"\s+\.$", "", val).strip(" –-:")
    if fld in ("level", "duration", "topic", "materials"):
        val = val.rstrip(". ")
    if fld == "component":
        n = norm(val)
        if "نصوص" in n or "قراءه" in n:
            return "القراءة"
    return val


def parse_pieces(pieces):
    """قاموس الحقول من نصوص حرّة: «عنوان: قيمة» مع عدة عناوين في السطر الواحد."""
    info, comps = {}, {}
    for raw in pieces:
        if not raw or looks_reversed(raw):
            continue
        s = plain(raw)
        u, d, s = extract_unit(s)
        if u and "unit" not in info:
            info["unit"] = u
            if d:
                info.setdefault("domain", d)
        # «مكون القراءة: (نص مسترسل)»
        m = re.search(r"(?:^|[\s\-–ـ])مكون\s+([^\s:：]+)\s*[:：]?\s*(\([^)]*\))?", s)
        if m and "component" not in info:
            info["component"] = (m.group(1) + (" " + m.group(2) if m.group(2) else "")).strip()
        matches = list(LABEL_RE.finditer(s))
        for i, mt in enumerate(matches):
            fld = mt.lastgroup
            end = mt.end()
            nxt = matches[i + 1].start() if i + 1 < len(matches) else len(s)
            val = re.sub(r"^[\s:：\-–ـ]+", "", s[end:nxt])
            val = re.sub(r"[\s\-–ـ]+$", "", val)
            val = re.sub(r"\s{2,}", " ", val)
            if not val:
                continue
            if fld == "_sep":
                continue
            if fld == "comp":
                k = re.sub(r"^ال", "", re.sub(r"^(?:ال)?كفاية\s+", "", mt.group(0)).strip())
                k = re.sub(r"^ال", "", k)
                k = "استراتيجية" if k.startswith("إستراتيجية") else k
                comps.setdefault(k, val)
            else:
                val = _tidy(fld, val)
                if val:
                    info.setdefault(fld, val)
    return info, comps


# ═══════════════════════════ قراءة المصدر ═══════════════════════════
class Source:
    def __init__(self, path):
        R.reset_nested()
        self.doc = R.open_docx(path, "الجذاذة المصدر")
        self.body = self.doc.element.body
        self.grids = [Grid(t) for t in self.body.findall(W + "tbl")]
        self.layout = ""
        self.notes = []
        self.units = []
        self._main = self._pick_main()
        pieces = self._header_pieces()
        self.info, self.comps = parse_pieces(pieces)
        self._read_stages()

    # ----- الجدول الرئيسي -----
    def _pick_main(self):
        best = None
        for g in self.grids:
            h = find_header(g)
            if not h:
                continue
            start, n, roles = h
            kinds = {r for r in roles.values() if r}
            core = kinds & {"stage", "content", "teacher", "learner", "activities"}
            if len(kinds) < 3 or not core:
                continue
            score = len(kinds) * 10 + len(g.text()) / 300.0
            if best is None or score > best[0]:
                best = (score, g, h)
        if best:
            return best[1], best[2]
        return None

    def _header_pieces(self):
        main = self._main[0] if self._main else None
        pieces = []
        for g in self.grids:
            if g is main:
                st = self._main[1][0]
                for c in g.cells:
                    if c.r1 < st and c.lines:                 # سطور تسبق عناوين الأعمدة
                        pieces.extend(c.lines)
                        pieces.append(" ".join(c.lines))
                continue
            if self._main and g.ncols == main.ncols and find_header(g) is None and g.nrows <= 12 and main is not None \
                    and any(c.lines for c in g.cells) and g.ncols >= 4 and g is not self.grids[0]:
                continue                                        # تتمة الجدول الرئيسي
            pieces += self._table_pieces(g)
        pieces += body_lines(self.body, skip_tbls=[g.el for g in self.grids])
        # مربعات النص خارج الجداول
        seen = set()
        for tb in self.body.iter(W + "txbxContent"):
            if any(a.tag.endswith("}Fallback") for a in tb.iterancestors()):
                continue
            if any(a is g.el for a in tb.iterancestors() for g in self.grids):
                continue
            ls = []
            for p in tb.iter(W + "p"):
                t = plain("".join(x.text or "" for x in p.iter(W + "t")))
                if t:
                    ls.append(t)
            joined = " ".join(ls)
            if joined and joined not in seen:
                seen.add(joined)
                pieces.extend(ls)
                pieces.append(joined)
        out, seen2 = [], set()
        for x in pieces:
            if x and x not in seen2:
                seen2.add(x)
                out.append(x)
        return out

    def _table_pieces(self, g):
        out = []
        labels = {n for n in (norm(p) for _, p in [(f, x) for f, ps in FIELD_LABELS for x in ps.split("|")])} | {norm("الوحدة")}
        for r in range(g.nrows):
            cs = g.row_cells(r)
            texts = [" ".join(c.lines) for c in cs]
            for i, t in enumerate(texts):
                if norm(t) in labels:
                    j = i - 1
                    while j >= 0 and (texts[j].strip() in (":", "：", "") or norm(texts[j]) in labels):
                        j -= 1
                    if j >= 0:
                        out.append(f"{t.strip()}: {texts[j].strip()}")
                else:
                    out.extend(cs[i].lines)
        return out

    # ----- المراحل -----
    def _read_stages(self):
        if not self._main:
            self._read_free_text()
            return
        grid, (hs, hn, roles) = self._main
        body_start = hs + hn
        parts = [(grid, body_start)]
        for g in self.grids:                                     # جداول متتابعة بنفس الأعمدة ولا عناوين لها
            if g is grid:
                continue
            if g.ncols == grid.ncols and find_header(g) is None and self.grids.index(g) > self.grids.index(grid):
                parts.append((g, 0))
        rows = []
        for g, st in parts:
            for r in range(st, g.nrows):
                rec = {k: [] for k in ROLES + ("stage",)}
                for c in range(g.ncols):
                    role = roles.get(c)
                    cell = g.at(r, c)
                    if cell is None or not role:
                        continue
                    if cell.c0 != c:
                        continue
                    if cell.r0 != r and role != "stage":
                        continue
                    rec[role] += cell.lines
                if any(rec[k] for k in rec):
                    rows.append(rec)
        self.layout = f"جدول من {grid.ncols} أعمدة"
        # صفّ واحد/صفّان يحوي المراحل كقائمة؟
        stage_lines = [x for rec in rows for x in rec["stage"]]
        if len(rows) <= 2 and len(stage_lines) >= 3 and all(len(x) <= 45 for x in stage_lines) \
                and sum(1 for x in stage_lines if canon(x)) >= 2:
            self.layout = "قوائم متوازية داخل خلية واحدة"
            self._explode(rows)
        else:
            self._rows_to_units(rows)

    def _rows_to_units(self, rows):
        units, last_key = [], None
        for rec in rows:
            label = " ".join(rec["stage"]).strip()
            k = canon(label)
            if k is None and not label:
                k = last_key
            if k is None:
                # خمّن من المحتوى/أنشطة المدرس
                probe = " ".join(rec["teacher"][:2] + rec["activities"][:2] + rec["content"][:2])
                k = item_cue(probe, "teacher") or last_key
            if units and units[-1].key == k and (units[-1].label == label or not label or True):
                u = units[-1]
            else:
                u = Unit(label, k)
                units.append(u)
            last_key = k
            self._fill_unit(u, rec)
        self.units = units

    def _fill_unit(self, u, rec):
        for role in ROLES:
            if rec.get(role):
                u.add(role, R.join_items(rec[role]))

    def _explode(self, rows):
        labels = [x for rec in rows for x in rec["stage"]]
        n = len(labels)
        units = [Unit(lab, canon(lab)) for lab in labels]
        merged = {k: [] for k in ROLES}
        for rec in rows:
            for k in ROLES:
                merged[k] += rec[k]
        approx, assign = False, {}
        EQUIV = {"hyp": ("frame",), "read": ("und", "frame"), "frame": ("diag",), "und": ("read", "frame"),
                 "ana": ("und",), "syn": ("ana",), "inv": ("syn",)}
        order = ["teacher", "activities", "content", "ability", "eval", "support", "indic", "notes", "learner", "time"]
        for role in order:
            items = R.join_items(merged[role])
            if not items:
                continue
            if len(items) == n:
                idxs = list(range(n))
            elif role == "learner" and "teacher" in assign and len(items) == len(assign["teacher"]):
                idxs = list(assign["teacher"])                  # قائمتان متوازيتان: المدرس/المتعلم
            else:
                approx = True
                idxs, cur = [], 0
                for it in items:
                    cue = item_cue(it, role)
                    if cue:
                        for c in (cue,) + EQUIV.get(cue, ()):
                            idx = next((i for i in range(cur, n) if units[i].key == c), None)
                            if idx is not None:
                                cur = idx
                                break
                    idxs.append(cur)
            assign[role] = idxs
            for it, i in zip(items, idxs):
                units[i].add(role, [it])
        if approx:
            self.notes.append("حُوذيت أنشطة بعض الأعمدة مع المراحل تقريباً (قوائم داخل خلية واحدة)؛ يُستحسن مراجعتها.")
        self.units = units

    def _read_free_text(self):
        """جذاذة بلا جدول: عناوين المراحل وسطور «أنشطة المدرس:» ونحوها."""
        self.layout = "نص حرّ بلا أعمدة"
        lines = [x for x in body_lines(self.body) if x]
        role_pat = [("teacher", r"^(?:أنشطة|انشطة)\s*(?:ال)?(?:مدرس|أستاذ|استاذ)|^أنشطة التعليم"),
                    ("learner", r"^(?:أنشطة|انشطة)\s*(?:ال)?(?:متعلم|تلميذ)|^أنشطة التعلم"),
                    ("content", r"^(?:ال)?محتوى|^المضامين"),
                    ("eval", r"^(?:ال)?تقويم|^الدعم"),
                    ("ability", r"^(?:ال)?قدرات|^(?:ال)?أهداف|^الاهداف")]
        units, cur, role = [], None, "content"
        for ln in lines:
            nl = norm(ln)
            head = len(ln) <= 60 and canon(ln) and (re.match(r"^\W*(?:[ا-ي]|\d+|[IVX]+)\s*[\-–.)]", ln) or ln.rstrip().endswith(":") or len(ln.split()) <= 4)
            hit = next((r for r, p in role_pat if re.search(p, ln)), None)
            if head and not hit:
                cur = Unit(re.sub(r"[:：]\s*$", "", ln), canon(ln))
                units.append(cur)
                role = "content"
                continue
            if cur is None:
                continue
            if hit:
                role = hit
                rest = re.sub(r"^[^:：]*[:：]", "", ln).strip()
                if rest:
                    cur.add(role, [rest])
                continue
            cur.add(role, [ln])
        self.units = [u for u in units if u.size()]
        if not self.units:
            self.notes.append("لم أتعرّف على مراحل في هذا النص الحرّ.")


# ═══════════════════════════ التعبئة ═══════════════════════════
LEARNER_RE = re.compile(r"^\W*(?:ال)?(?:متعلمون|متعلمات|تلاميذ|متعلمين)\b|^\W*(?:يجيب|يقرأ|يستخرج|يصوغ|يشرح|يتبادل|يبدي|يعبر|ينجز|يحدد|يلاحظ|يتعرف|يلخص|يستنتج)\s*(?:ال)?(?:متعلم|تلميذ)")


def split_activities(items):
    t, l = [], []
    for it in items:
        (l if LEARNER_RE.search(R.strip_marker(it)) else t).append(it)
    return t, l


def dedupe(lines):
    seen, out = set(), []
    for x in lines:
        k = R.key(x) if x else None
        if x and k in seen:
            continue
        if k:
            seen.add(k)
        out.append(x)
    return out


def clean_label(lbl):
    s = re.sub(r"\s*\d+\s*(?:د|دقيقة|دقائق)\b", "", lbl)
    s = re.sub(r"^\W*(?:[IVX]+|\d+|[أ-ي])\s*[\-–.)]\s*", "", s).strip()
    return re.sub(r"[:：\s]+$", "", s)


def unit_activity_lines(u):
    teacher = list(u.f["teacher"])
    learner = list(u.f["learner"])
    if u.f["activities"]:
        t, l = split_activities(u.f["activities"])
        teacher += t
        learner += l
    return [R.T(x) for x in teacher] + [R.D(x) for x in learner]


def render_row(units, numbered):
    c0, c1, c2, c3 = [], [], [], []
    cf = R.ContentFmt(num_style="heading", number_headings=False)
    for n, u in enumerate(units):
        if n:
            c0 and c0.append("")
            c2.append("")
            c3 and c3.append("")
        for k in ("eval", "indic", "support", "notes"):
            c0 += [R.T(x) for x in u.f[k]]
        c1 += cf.fmt(u.f["content"])
        acts = unit_activity_lines(u)
        if numbered and len(units) > 1:
            c2.append(R.H(f"{n + 1}- {clean_label(u.label) or STAGE_TITLE.get(u.key, '')}:"))
        c2 += acts
        c3 += ["- " + R.end(R.strip_marker(x)) for x in u.f["ability"]]
    return dedupe(c0), c1, dedupe(c2), dedupe(c3)


ROW_MAP = [  # (صف القالب بعد حذف «تمهيد»، مفاتيح المراحل، عنوان المقطع)
    (2, ("diag",), "تشخيص المكتسبات"),
    (3, ("frame", "hyp"), "ملاحظة النص"),
    (4, ("read", "und"), "فهم النص"),
    (7, ("ana",), "القراءة التحليلية"),
    (8, ("syn",), "التركيب"),
    (9, ("inv",), "الاستثمار"),
]


def build_general(model_path, source_path, out_name, ai=None, out_root=None, fit=True):
    out_root = Path(out_root) if out_root else Path(R.OUT_DIR)
    src = Source(source_path)
    rep = R.Report()
    for n in src.notes:
        rep.notes.append(n)
    rep.notes.append("تخطيط المصدر: " + (src.layout or "غير معروف") + " (قارئ عام)")

    doc = R.open_docx(model_path, "الجذاذة النموذج")
    if len(doc.tables) != 2 or len(doc.tables[1].rows) != 11:
        raise R.ShapeError("النموذج المتوقع بتخطيط «المنفى»")
    H0, MT = doc.tables[0], doc.tables[1]
    info = src.info

    level = info.get("level", "")
    unit_val = re.sub(r"^الوحدة\s+", "", info.get("unit", ""))
    domain = re.sub(r"^مجال\s+", "", info.get("domain", ""))
    component = info.get("component", "") or "القراءة"
    topic = re.sub(r"\s*/\s*", " / ", info.get("topic", ""))
    topic = re.sub(r"\s+ص\s*[:：]?\s*(\d+)\s*$", r" – ص: \1", topic)
    duration = info.get("duration", "")
    materials = info.get("materials", "")
    materials = " – ".join(x.strip() for x in re.split(r"\s*[+،\-–]\s+", materials) if x.strip()) if materials else ""

    R.set_cell_keep_format(R.tcell(H0, 0, 0), [level])
    R.set_cell_keep_format(R.tcell(H0, 0, 4), [unit_val])
    R.set_cell_keep_format(R.tcell(H0, 1, 0), [materials])
    R.set_cell_keep_format(R.tcell(H0, 1, 4), [domain])
    R.set_cell_keep_format(R.tcell(H0, 2, 2), [component])
    R.set_cell_keep_format(R.tcell(H0, 3, 2), [topic])
    R.set_cell_keep_format(R.tcell(H0, 4, 2), [duration])
    wanted = ["تواصلية", "منهجية", "ثقافية"]
    comps = dict(src.comps)
    for r, w, pre in zip((3, 4, 5), wanted, ("", "- ", "- ")):
        if w in comps:
            R.set_cell_keep_format(R.tcell(H0, r, 0), [f"{pre}كفاية {w}: {comps[w]}"])
    for k in ("level", "unit", "domain", "topic", "duration"):
        if not info.get(k):
            rep.notes.append(f"لم أجد في المصدر حقل الترويسة: {k}")

    rows = MT.rows

    def cells_of(r):
        return [R._Cell(tc, MT) for tc in rows[r]._tr.findall(qn("w:tc"))]

    R.fill(cells_of(0)[0], ["الحصة الأولى: التقويم التشخيصي، ملاحظة النص وفهمه"], "c")
    R.fill(cells_of(5)[0], ["الحصة الثانية: القراءة التحليلية والتركيبية والاستثمار"], "c")
    MT._tbl.remove(rows[7]._tr)

    by = {}
    for u in src.units:
        by.setdefault(u.key, []).append(u)
    placed = set()
    for ri, keys, title in ROW_MAP:
        us = [u for k in keys for u in by.get(k, [])]
        for u in us:
            placed.add(id(u))
        c = cells_of(ri)
        if not us:
            for cell in c[:4]:
                R.fill(cell, [""], "c")
            R.fill(c[4], [title], "c")
            rep.notes.append(f"لا مرحلة في المصدر تقابل «{title}»؛ مُسح محتوى القالب القديم.")
            continue
        c0, c1, c2, c3 = render_row(us, numbered=ri in (3, 4))
        R.fill(c[0], c0, "c")
        R.fill(c[1], c1, "r")
        R.fill(c[2], c2, "c")
        R.fill(c[3], c3, "c")
        R.fill(c[4], [("- " + title) if ri in (2, 3) else title], "c")
    # وحدات بلا مرحلة معروفة: تُلحق بأقرب صف سابق في القالب
    for i, u in enumerate(src.units):
        if id(u) in placed:
            continue
        prev = next((x for x in reversed(src.units[:i]) if id(x) in placed), None)
        ri = 2
        if prev is not None:
            ri = next((r for r, ks, _ in ROW_MAP if prev.key in ks), 2)
        c = cells_of(ri)
        c0, c1, c2, c3 = render_row([u], numbered=False)
        extra_head = [R.H(clean_label(u.label) + ":")] if u.label else []
        for cell, extra, kind in ((c[0], c0, "c"), (c[1], c1, "r"), (c[2], extra_head + c2, "c"), (c[3], c3, "c")):
            if extra:
                old = [p for p in R_cell_texts(cell)]
                R.fill(cell, old + [""] + extra, kind)
        rep.notes.append(f"مرحلة «{u.label}» أُلحقت بالصف «{ROW_MAP[[r for r, _, _ in ROW_MAP].index(ri)][2]}»")

    ai_added = []
    if ai:
        ai_added = R.ai_fill_blanks(MT, H0, cells_of, comps, wanted, topic, level, component, ai) or []

    # فصل الحصتين في صفحتين + هوامش (كما في reading.py)
    tc_h2 = MT.rows[5].cells[0]._tc
    p_h2 = tc_h2.find(qn("w:p"))
    if p_h2 is not None:
        pPr = p_h2.find(qn("w:pPr"))
        if pPr is None:
            from docx.oxml import OxmlElement
            pPr = OxmlElement("w:pPr")
            p_h2.insert(0, pPr)
        from docx.oxml import OxmlElement
        pPr.append(OxmlElement("w:pageBreakBefore"))
    for section in doc.sections:
        pgMar = section._sectPr.find(qn("w:pgMar"))
        if pgMar is not None:
            for side in ("top", "bottom", "left", "right"):
                pgMar.set(qn("w:" + side), "720")
    if fit:
        R.fit_pages(doc, H0, MT, ai_added)

    out_path = R.save_safely(doc, R.output_path(out_root, out_name, component, unit_val))
    print("تم الحفظ بنجاح في:", out_path)
    return rep


def R_cell_texts(cell):
    """نصوص فقرات خلية القالب بعد التعبئة (للإلحاق بها)."""
    out = []
    for p in cell._tc.findall(qn("w:p")):
        t = "".join(x.text or "" for x in p.iter(qn("w:t")))
        out.append(t)
    while out and not out[-1].strip():
        out.pop()
    return out


# ═══════════════════════════ نقطة الدخول: القديم أولاً ثم العام ═══════════════════════════
def build_any(model_path, source_path, out_name, ai=None, out_root=None, fit=True):
    """يجرّب قارئ «المختار 3 إع» الدقيق، فإن اختلف التخطيط استعمل القارئ العام."""
    out_root = Path(out_root) if out_root else Path(R.OUT_DIR)
    try:
        return R.build(model_path, source_path, out_name, ai, out_root, fit)
    except R.ShapeError as e:                # تخطيط مختلف عن «المنفى» ← القارئ العام
        fallback_reason = str(e) or type(e).__name__
    rep = build_general(model_path, source_path, out_name, ai, out_root, fit)
    rep.notes.append("استُعمل القارئ العام بعد تعذّر القارئ الدقيق: " + fallback_reason)
    return rep
