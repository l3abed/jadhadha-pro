# -*- coding: utf-8 -*-
"""
linkimport.py — استيراد نص التحضير من رابط (أو من نص ملصوق).

- الروابط: https فقط، ويُحترم ملف robots.txt، ومهلة وحد أقصى للحجم، ومنع الشبكات الداخلية.
- الصفحة بيانات لا تعليمات: النص يُدرج كما هو، ولا يُنفَّذ.
- المنصات التي تتطلب تسجيل الدخول أو تمنع الجلب تُعاد برسالة واضحة.
"""
import ipaddress
import re
import socket
import urllib.error
import urllib.parse
import urllib.request
import urllib.robotparser
from html.parser import HTMLParser

UA = "JadhadhaApp/2.7 (personal lesson preparation)"
MAX_BYTES = 2_000_000
TIMEOUT = 15
_SKIP = {"script", "style", "nav", "header", "footer", "aside", "form", "noscript", "svg", "button"}
_BLOCK = {"p", "li", "h1", "h2", "h3", "h4", "div", "br", "tr", "section", "article"}


class LinkError(Exception):
    pass


def _validate_url(url: str) -> str:
    """رابط HTTPS عام فقط: لا localhost ولا شبكات داخلية ولا إعادة توجيه إليها."""
    url = (url or "").strip()
    p = urllib.parse.urlsplit(url)
    if p.scheme != "https" or not p.netloc:
        raise LinkError("الرابط غير صالح: يجب أن يبدأ بـ https://")
    host = (p.hostname or "").strip().lower().rstrip(".")
    if not host or host in ("localhost", "localhost.localdomain") or host.endswith(".local"):
        raise LinkError("لا يُسمح بجلب العناوين المحلية.")
    try:
        port = p.port or 443
    except ValueError:
        raise LinkError("الرابط غير صالح: منفذ غير صحيح.")
    try:
        infos = socket.getaddrinfo(host, port, type=socket.SOCK_STREAM)
    except OSError:
        raise LinkError("تعذّر تحليل اسم النطاق.")
    for info in infos:
        ip = ipaddress.ip_address(info[4][0])
        if (ip.is_private or ip.is_loopback or ip.is_link_local or ip.is_reserved
                or ip.is_multicast or ip.is_unspecified):
            raise LinkError("لا يُسمح بالوصول إلى عناوين الشبكة الداخلية.")
    return url


class _SafeRedirect(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        _validate_url(newurl)
        return super().redirect_request(req, fp, code, msg, headers, newurl)


_OPENER = urllib.request.build_opener(_SafeRedirect())


def _open(req, timeout=TIMEOUT):
    return _OPENER.open(req, timeout=timeout)


class _Text(HTMLParser):
    def __init__(self):
        super().__init__()
        self.parts, self.skip, self.title = [], 0, ""
        self._in_title = False

    def handle_starttag(self, tag, attrs):
        if tag == "title":
            self._in_title = True
        if tag in _SKIP:
            self.skip += 1
        if tag in _BLOCK:
            self.parts.append("\n")

    def handle_endtag(self, tag):
        if tag == "title":
            self._in_title = False
        if tag in _SKIP and self.skip:
            self.skip -= 1
        if tag in _BLOCK:
            self.parts.append("\n")

    def handle_data(self, data):
        if self._in_title:
            self.title += data
        elif not self.skip:
            self.parts.append(data)


def clean_text(raw: str) -> str:
    lines = [re.sub(r"[ \t\u00a0]+", " ", ln).strip() for ln in raw.split("\n")]
    out, blank = [], False
    for ln in lines:
        if not ln:
            if not blank and out:
                out.append("")
            blank = True
            continue
        blank = False
        out.append(ln)
    return "\n".join(out).strip()


def _robots_ok(url: str) -> bool:
    p = urllib.parse.urlsplit(url)
    rp = urllib.robotparser.RobotFileParser()
    try:
        req = urllib.request.Request(f"https://{p.netloc}/robots.txt", headers={"User-Agent": UA})
        with _open(req, timeout=TIMEOUT) as r:
            rp.parse(r.read().decode("utf-8", "replace").splitlines())
    except urllib.error.HTTPError as e:
        if e.code in (401, 403):
            return False                           # المنصة تمنع الجلب الآلي
        return True                                # لا يوجد robots.txt
    except Exception:
        return True
    return rp.can_fetch(UA, url)


def fetch_text(url: str) -> dict:
    url = _validate_url(url)
    p = urllib.parse.urlsplit(url)
    if not _robots_ok(url):
        raise LinkError("هذه المنصة تمنع الجلب الآلي لصفحاتها. انسخ النص والصقه في خانة اللصق.")
    req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept": "text/html,text/plain"})
    try:
        with _open(req, timeout=TIMEOUT) as r:
            ctype = r.headers.get("Content-Type", "")
            if "html" not in ctype and "text" not in ctype:
                raise LinkError("الرابط ليس صفحة نصية.")
            raw = r.read(MAX_BYTES + 1)
    except urllib.error.HTTPError as e:
        if e.code in (401, 403):
            raise LinkError("الصفحة تتطلب تسجيل الدخول أو تمنع الوصول. انسخ النص والصقه.")
        raise LinkError(f"تعذّر فتح الرابط (رمز {e.code}).")
    except (urllib.error.URLError, TimeoutError) as e:
        raise LinkError(f"تعذّر الاتصال بالرابط: {e}")
    if len(raw) > MAX_BYTES:
        raise LinkError("الصفحة كبيرة جداً.")
    html = raw.decode("utf-8", "replace")
    parser = _Text()
    parser.feed(html)
    text = clean_text("".join(parser.parts))
    if len(text) < 40:
        raise LinkError("لم أجد نصاً كافياً في الصفحة (قد تكون محمية بالجافاسكربت). انسخ النص والصقه.")
    return {"title": parser.title.strip(), "text": text[:20000], "source": url}


def text_from_paste(text: str) -> str:
    t = clean_text(text or "")
    if not t:
        raise LinkError("الخانة فارغة.")
    return t[:20000]


def _norm_ar(s: str) -> str:
    s = re.sub("[\u064B-\u0652\u0640]", "", s or "")
    s = re.sub("[أإآ]", "ا", s).replace("ى", "ي").replace("ة", "ه")
    return re.sub(r"\s+", " ", s).strip()


def insert_into_output(docx_path, text: str, stage_key: str = "") -> int:
    """يُدرج النص في خانة «المحتوى» من صف المرحلة المطلوبة (أو أول صف). يعيد عدد الفقرات المُدرجة."""
    import copy
    from docx import Document
    from docx.oxml.ns import qn
    doc = Document(str(docx_path))
    main, hrow = None, None
    for t in doc.tables:
        for hi, r in enumerate(t.rows[:4]):
            heads = [_norm_ar(c.text) for c in r.cells]
            if any("المحتوي" in h for h in heads) and any(("المراحل" in h or "المقاط" in h) for h in heads):
                main, hrow = t, hi
                break
        if main is not None:
            break
    if main is None:
        raise LinkError("لم أجد جدول المحتوى في الجذاذة.")
    heads = [_norm_ar(c.text) for c in main.rows[hrow].cells]
    ci = next(i for i, h in enumerate(heads) if "المحتوي" in h)
    si = next(i for i, h in enumerate(heads) if ("المراحل" in h or "المقاط" in h))
    body = main.rows[hrow + 1:]
    row = None
    if stage_key.strip():
        key = _norm_ar(stage_key)
        row = next((r for r in body if key and key in _norm_ar(r.cells[si].text)), None)
        if row is None:
            raise LinkError(f"لم أجد مرحلة باسم «{stage_key}» في الجذاذة.")
    else:
        row = body[0] if body else None
    if row is None:
        raise LinkError("الجذاذة لا تحتوي صفوف مراحل.")
    cell = row.cells[ci]._tc
    ref_p = cell.find(qn("w:p"))
    added = 0
    for line in [ln for ln in text.split("\n")]:
        p = copy.deepcopy(ref_p) if ref_p is not None else None
        if p is None:
            from docx.oxml import OxmlElement
            p = OxmlElement("w:p")
        for r in p.findall(qn("w:r")):
            p.remove(r)
        for h in p.findall(qn("w:hyperlink")):
            p.remove(h)
        from docx.oxml import OxmlElement
        r = OxmlElement("w:r")
        if ref_p is not None:
            src_r = ref_p.find(".//" + qn("w:r"))
            if src_r is not None and src_r.find(qn("w:rPr")) is not None:
                r.append(copy.deepcopy(src_r.find(qn("w:rPr"))))
        t_el = OxmlElement("w:t")
        t_el.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        t_el.text = line.strip()
        r.append(t_el)
        p.append(r)
        cell.append(p)
        added += 1
    doc.save(str(docx_path))
    return added
