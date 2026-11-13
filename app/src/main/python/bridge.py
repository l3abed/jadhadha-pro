# -*- coding: utf-8 -*-
"""
bridge.py — واجهة موحّدة لسكريبتات الجذاذات الثلاثة (قراءة / لغة / تعبير).

يستدعيها التطبيق (Chaquopy) هكذا:
    run("reading" | "language" | "expression", model_path, source_path, out_dir, api_key="", out_name="", header_json="")
وتعيد نص JSON:  {"ok": bool, "path": "...", "log": "...", "error": "...", "target": {"component": "..", "unit": ".."}}

لا مفتاح داخل الكود: إن مُرِّر api_key فُعّل التنقيح بالذكاء الاصطناعي، وإلا عمل السكريبت دون إنترنت.
"""
import contextlib
import importlib
import io
import json
import os
import re
import sys
import traceback
from pathlib import Path

_HERE = Path(__file__).resolve().parent
if str(_HERE) not in sys.path:
    sys.path.insert(0, str(_HERE))

COMPONENTS = ("reading", "language", "expression")
_PATH_RE = re.compile(r"(?:الملف|تم الحفظ بنجاح في)\s*:\s*(.+?\.docx)\s*$", re.M)


@contextlib.contextmanager
def _env_key(api_key):
    old = os.environ.get("GEMINI_API_KEY")
    if api_key:
        os.environ["GEMINI_API_KEY"] = api_key
    else:
        os.environ.pop("GEMINI_API_KEY", None)
    try:
        yield
    finally:
        if old is None:
            os.environ.pop("GEMINI_API_KEY", None)
        else:
            os.environ["GEMINI_API_KEY"] = old


_BAD_NAME = re.compile(r'[<>:"/\\|?*\r\n\t]+')


def _safe_out_name(out_name, source):
    out_name = _BAD_NAME.sub("", out_name or "").strip(" .")
    if not out_name:
        return ""
    if not out_name.lower().endswith(".docx"):
        out_name += ".docx"
    return out_name


def _run_reading(model, source, out_dir, api_key, out_name=""):
    mod = importlib.import_module("reading")
    gen = importlib.import_module("reading_general")
    mod.USED_OUTPUTS.clear()
    ai = mod.AICfg(mod.DEFAULT_AI_MODEL, api_key) if api_key else None
    out_name = _safe_out_name(out_name, source) or ("جذاذة_" + Path(source).stem + ".docx")
    gen.build_any(str(model), str(source), out_name, ai, Path(out_dir), True)


def _run_cli(component, model, source, out_dir, api_key, out_name=""):
    mod = importlib.import_module(component)
    argv = [str(model), str(source), "-o", str(out_dir)]
    name = _safe_out_name(out_name, source)
    if name:
        argv += ["--out-name", name]
    if component == "language":
        argv += ["--api-key", api_key] if api_key else ["--no-ai"]
    code = mod.main(argv)
    if code not in (0, None):
        raise RuntimeError("انتهى السكريبت برمز خطأ %s" % code)


# ───────── التعرّف على المكوّن والوحدة من ترويسة الجذاذة (لاختيار مجلد التخزين) ─────────
COMPONENT_FOLDERS = {"reading": "القراءة", "language": "الدرس اللغوي", "expression": "التعبير والإنشاء"}
_UNITS = ("الأولى", "الثانية", "الثالثة", "الرابعة", "الخامسة", "السادسة")


def _n(s):
    s = re.sub("[\u064B-\u0652\u0640]", "", s or "")
    s = re.sub("[أإآٱ]", "ا", s).replace("ى", "ي").replace("ة", "ه")
    return re.sub(r"\s+", " ", re.sub(r"[^\w\s]", " ", s)).strip()


def _header_value(table, label):
    """قيمة الخانة الواقعة قبل عنوان معيّن (الترويسة: القيمة | : | العنوان)."""
    for row in table.rows:
        seen, cells = set(), []
        for c in row.cells:
            if id(c._tc) in seen:
                continue
            seen.add(id(c._tc))
            cells.append(re.sub(r"\s+", " ", c.text).strip())
        for i, t in enumerate(cells):
            if _n(t) != label:
                continue
            j = i - 1
            while j >= 0 and (cells[j] in (":", "：", "؛") or not cells[j]):
                j -= 1
            if j >= 0:
                return cells[j]
    return ""


def _paragraph_header(doc):
    """ترويسة مكتوبة كفقرات: «العنوان: القيمة» أو «القيمة : العنوان»."""
    out = {}
    for p in doc.paragraphs:
        t = re.sub(r"\s+", " ", p.text).strip()
        if not t:
            continue
        m = re.match(r"^(.+?)\s*[:：]\s*(.+)$", t)
        if not m:
            continue
        left, right = m.group(1).strip(), m.group(2).strip()
        for key, labels in _HEADER_LABELS.items():
            if _n(left) in labels and _n(right) not in labels:
                out.setdefault(key, right)
            elif _n(right) in labels and _n(left) not in labels:
                out.setdefault(key, left)
    return out



def detect_target(docx_path, fallback_component=""):
    """يعيد {"component": مجلد المكوّن, "unit": مجلد الوحدة} اعتماداً على ترويسة الملف."""
    from docx import Document
    comp, unit, src = "", "", ""
    try:
        doc = Document(str(docx_path))
        v_comp = v_unit = ""
        if doc.tables:
            t = doc.tables[0]
            v_comp, v_unit = _header_value(t, "المكون"), _header_value(t, "الوحده")
        if not v_comp or not v_unit:
            para = _paragraph_header(doc)
            v_comp = v_comp or para.get("component", "")
            v_unit = v_unit or para.get("unit", "")
        n = _n(v_comp)
        if "قراءه" in n:
            comp = COMPONENT_FOLDERS["reading"]
        elif "تعبير" in n or "انشاء" in n:
            comp = COMPONENT_FOLDERS["expression"]
        elif "لغ" in n or "قواعد" in n or "صرف" in n or "نحو" in n:
            comp = COMPONENT_FOLDERS["language"]
        nu = _n(v_unit)
        for i, w in enumerate(_UNITS, 1):
            if _n(w) in nu or re.search(r"(?<!\d)%d(?!\d)" % i, nu):
                unit = "الوحدة " + w
                break
        src = "الترويسة"
    except Exception:
        pass
    if not comp:
        comp, src = COMPONENT_FOLDERS.get(fallback_component, ""), "التبويب"
    return {"component": comp, "unit": unit, "from": src}


# ───────── تعديل الترويسة (الأستاذ، المؤسسة، الفئة…) من إعدادات التطبيق ─────────
_HEADER_LABELS = {
    "teacher": ("الاستاذ", "الاستاذه"),
    "school": ("المؤسسه",),
    "level": ("الفئه المستهدفه",),
    "unit": ("الوحده",),
    "component": ("المكون",),
}


def _set_cell_text(cell, text, label_cell=None):
    p = cell.paragraphs[0]
    if p.runs:
        p.runs[0].text = text
        for r in p.runs[1:]:
            r.text = ""
    else:
        run = p.add_run(text)
        if label_cell is not None and label_cell.paragraphs and label_cell.paragraphs[0].runs:
            src = label_cell.paragraphs[0].runs[0]._r.find(
                "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr")
            if src is not None:
                import copy as _copy
                run._r.insert(0, _copy.deepcopy(src))


def apply_header(docx_path, overrides):
    """يضع قيم الإعدادات غير الفارغة في خانات الترويسة (القيمة تسبق عنوانها بنقطتين). يعيد عدد الخانات المعدَّلة."""
    ov = {k: (v or "").strip() for k, v in (overrides or {}).items() if k in _HEADER_LABELS and (v or "").strip()}
    if not ov:
        return 0
    from docx import Document
    doc = Document(str(docx_path))
    all_labels = sum(_HEADER_LABELS.values(), ())
    done = 0
    for tbl in doc.tables:
        for row in tbl.rows:
            seen, cells = set(), []
            for c in row.cells:
                if id(c._tc) not in seen:
                    seen.add(id(c._tc))
                    cells.append(c)
            for i, c in enumerate(cells):
                n = _n(c.text)
                for key, labels in _HEADER_LABELS.items():
                    if key in ov and n in labels:
                        j = i - 1
                        while j >= 0 and cells[j].text.strip() in (":", "：", "؛"):
                            j -= 1
                        if j >= 0 and _n(cells[j].text) not in all_labels:
                            _set_cell_text(cells[j], ov[key], c)
                            done += 1
    if done:
        doc.save(str(docx_path))
    return done


def apply_header_json(docx_path, overrides_json):
    """واجهة للتطبيق: تعدّل الترويسة من نص JSON وتعيد عدد الخانات المعدَّلة."""
    try:
        return apply_header(docx_path, json.loads(overrides_json or "{}"))
    except Exception:
        return 0


def run(component, model, source, out_dir, api_key="", out_name="", header_json="", style="classic"):
    res = {"ok": False, "path": "", "log": "", "error": "", "target": {}}
    buf = io.StringIO()
    try:
        if component not in COMPONENTS:
            raise ValueError("مكوّن غير معروف: %s" % component)
        Path(out_dir).mkdir(parents=True, exist_ok=True)
        with _env_key(api_key), contextlib.redirect_stdout(buf), contextlib.redirect_stderr(buf):
            if component == "reading":
                _run_reading(model, source, out_dir, api_key, out_name)
            else:
                _run_cli(component, model, source, out_dir, api_key, out_name)
        log = buf.getvalue()
        m = _PATH_RE.findall(log)
        res["path"] = m[-1].strip() if m else ""
        res["ok"] = bool(res["path"]) and Path(res["path"]).exists()
        if not res["ok"]:
            res["error"] = ("لم يطبع السكريبت مسار الملف الناتج؛ أُلغي التفريغ بدل اختيار ملف قديم."
                            if not res["path"] else "الملف الناتج غير موجود: " + res["path"])
        else:
            try:
                ov = json.loads(header_json) if header_json else {}
            except ValueError:
                ov = {}
            apply_header(res["path"], ov)
            try:
                import styles as _styles
                _styles.apply_style(res["path"], style)
            except Exception as _e:                      # النمط تحسين شكلي: لا يُفشل التفريغ
                res["log"] += f"\n⚠ تعذّر تطبيق النمط: {_e}"
            res["target"] = detect_target(res["path"], component)
    except SystemExit as e:                      # ShapeError في سكريبت القراءة يرث SystemExit
        res["error"] = str(e) or "تخطيط المصدر غير مدعوم."
    except Exception as e:
        res["error"] = "%s: %s" % (type(e).__name__, e)
        buf.write(traceback.format_exc())
    res["log"] = buf.getvalue()
    return json.dumps(res, ensure_ascii=False)



# ───────── واجهات الجولة الثانية: توليد الجذاذة، واستيراد نص التحضير ─────────
def generate_lesson(component, lesson, subject, level, activity, api_key, out_dir, model_path, style="classic", header_json=""):
    """يولّد جذاذة بالذكاء الاصطناعي ثم يفرّغها في القالب المختار. يعيد JSON مثل run()."""
    import generate as G
    res = {"ok": False, "path": "", "log": "", "error": "", "target": {}}
    buf = io.StringIO()
    try:
        with contextlib.redirect_stdout(buf), contextlib.redirect_stderr(buf):
            src = G.generate(component, lesson, subject, level, activity, api_key, Path(out_dir) / "_مصدر")
            out = run(component, model_path, str(src), out_dir, "", "", header_json, style)   # الجذاذة مكتملة: لا تنقيح ثانٍ
        r = json.loads(out)
        res.update({"ok": r["ok"], "path": r["path"], "target": r["target"], "error": r["error"]})
    except G.GenerationError as e:
        res["error"] = str(e)
    except Exception as e:
        res["error"] = "%s: %s" % (type(e).__name__, e)
        buf.write(traceback.format_exc())
    res["log"] = buf.getvalue()
    return json.dumps(res, ensure_ascii=False)


def fetch_link(url):
    import linkimport as K
    try:
        r = K.fetch_text(url)
        return json.dumps({"ok": True, **r}, ensure_ascii=False)
    except K.LinkError as e:
        return json.dumps({"ok": False, "error": str(e)}, ensure_ascii=False)


def transcribe_text(component, text, activity, api_key, out_dir, model_path, style="classic", header_json=""):
    """يحوّل النص الملصوق أو المجلوب من الرابط إلى مصدر، ثم يفرّغه في القالب (بلا تنقيح ثانٍ)."""
    import generate as G
    import linkimport as K
    res = {"ok": False, "path": "", "log": "", "error": "", "target": {}}
    buf = io.StringIO()
    try:
        body = K.text_from_paste(text)
        with contextlib.redirect_stdout(buf), contextlib.redirect_stderr(buf):
            src = G.structure_text(component, body, activity, api_key, Path(out_dir) / "_مصدر")
            out = run(component, model_path, str(src), out_dir, "", "", header_json, style)
        r = json.loads(out)
        res.update({"ok": r["ok"], "path": r["path"], "target": r["target"], "error": r["error"]})
    except (G.GenerationError, K.LinkError) as e:
        res["error"] = str(e)
    except Exception as e:
        res["error"] = "%s: %s" % (type(e).__name__, e)
        buf.write(traceback.format_exc())
    res["log"] = buf.getvalue()
    return json.dumps(res, ensure_ascii=False)


def insert_imported(path, text, stage="", source=""):
    """يُدرج النص الملصوق أو المجلوب في خانة المحتوى من المرحلة المختارة. مع المصدر سطر توثيق في التقرير."""
    import linkimport as K
    try:
        body = K.text_from_paste(text)
        if source:
            body = body + "\n(مصدر: " + source + ")"
        n = K.insert_into_output(path, body, stage)
        return json.dumps({"ok": True, "added": n}, ensure_ascii=False)
    except K.LinkError as e:
        return json.dumps({"ok": False, "error": str(e)}, ensure_ascii=False)

if __name__ == "__main__":
    if len(sys.argv) < 5:
        print("الاستعمال: python bridge.py <reading|language|expression> <model.docx> <source.docx> <out_dir> [api_key]")
        sys.exit(2)
    print(run(*sys.argv[1:5],
              api_key=(sys.argv[5] if len(sys.argv) > 5 else ""),
              out_name=(sys.argv[6] if len(sys.argv) > 6 else "")))
