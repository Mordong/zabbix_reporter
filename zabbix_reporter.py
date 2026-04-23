"""
Zabbix Reporter v3.0  —  Windows 10/11
Получение репортов из Zabbix 7.x через JSON-RPC API
Экспорт: PDF, Excel (.xlsx), CSV
Данные: Проблемы/Алерты, Графики метрик, Статус хостов, История событий
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json
import urllib.request
import urllib.error
import csv
import os
import threading
import datetime
import ssl
import io

# ── Опциональные зависимости ──────────────────────────────────────────────────
try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.chart import BarChart, LineChart, Reference
    EXCEL_OK = True
except ImportError:
    EXCEL_OK = False

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import mm, cm
    from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                     TableStyle, PageBreak, HRFlowable)
    from reportlab.platypus import Image as RLImage
    from reportlab.graphics.charts.barcharts import VerticalBarChart
    from reportlab.graphics.charts.piecharts import Pie
    from reportlab.graphics.shapes import Drawing
    from reportlab.lib.colors import HexColor
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont

    # ── Регистрация шрифта с поддержкой кириллицы ────────────────────────────
    # Ищем подходящий TTF в системных папках Windows / Linux
    def _find_cyrillic_font():
        candidates = [
            # Windows
            r"C:\Windows\Fontsrial.ttf",
            r"C:\Windows\Fonts\calibri.ttf",
            r"C:\Windows\Fonts\segoeui.ttf",
            r"C:\Windows\Fonts	ahoma.ttf",
            r"C:\Windows\Fontserdana.ttf",
            # Linux fallback
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
        ]
        for p in candidates:
            if os.path.isfile(p):
                return p
        return None

    _CYR_TTF      = _find_cyrillic_font()
    _CYR_TTF_BOLD = None

    if _CYR_TTF:
        # Bold variant — same folder, try common naming
        _bold_try = [
            _CYR_TTF.replace("arial.ttf","arialbd.ttf"),
            _CYR_TTF.replace("calibri.ttf","calibrib.ttf"),
            _CYR_TTF.replace("segoeui.ttf","segoeuib.ttf"),
            _CYR_TTF.replace("tahoma.ttf","tahomabd.ttf"),
            _CYR_TTF.replace("verdana.ttf","verdanab.ttf"),
            _CYR_TTF.replace("-Regular","-Bold"),
            _CYR_TTF.replace("Sans.ttf","Sans-Bold.ttf"),
        ]
        for p in _bold_try:
            if os.path.isfile(p):
                _CYR_TTF_BOLD = p
                break

        pdfmetrics.registerFont(TTFont("CyrFont",     _CYR_TTF))
        pdfmetrics.registerFont(TTFont("CyrFont-Bold",
                                        _CYR_TTF_BOLD or _CYR_TTF))
        from reportlab.pdfbase.pdfmetrics import registerFontFamily
        registerFontFamily("CyrFont",
                            normal="CyrFont", bold="CyrFont-Bold",
                            italic="CyrFont", boldItalic="CyrFont-Bold")
        _PDF_FONT      = "CyrFont"
        _PDF_FONT_BOLD = "CyrFont-Bold"
    else:
        # Кириллица не будет отображаться, но хотя бы не упадёт
        _PDF_FONT      = "Helvetica"
        _PDF_FONT_BOLD = "Helvetica-Bold"

    PDF_OK = True
except ImportError:
    PDF_OK = False
    _PDF_FONT      = "Helvetica"
    _PDF_FONT_BOLD = "Helvetica-Bold"
    _CYR_TTF       = None

try:
    import matplotlib
    # Do NOT call matplotlib.use("Agg") here — it breaks FigureCanvasTkAgg.
    # The Agg backend is set only inside PDF export functions when needed.
    import matplotlib.pyplot as plt
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    MPL_OK    = True
    MPL_TK_OK = True
except ImportError:
    MPL_OK    = False
    MPL_TK_OK = False


# ══════════════════════════════════════════════════════════════════════════════
#  Константы
# ══════════════════════════════════════════════════════════════════════════════
SEVERITY_NAMES = {
    "0": "Not classified", "1": "Information", "2": "Warning",
    "3": "Average",        "4": "High",         "5": "Disaster",
}
SEVERITY_HEX = {
    "0": "#97AAB3", "1": "#7499FF", "2": "#FFC859",
    "3": "#FFA059", "4": "#E97659", "5": "#E45959",
}
HOST_STATUS = {"0": "Enabled", "1": "Disabled"}
HOST_AVAIL  = {"0": "Unknown", "1": "Available", "2": "Unavailable"}
IFACE_TYPE  = {"1": "Agent", "2": "SNMP", "3": "IPMI", "4": "JMX"}
IFACE_AVAIL = {"0": "Unknown", "1": "Available", "2": "Unavailable"}
IFACE_AVAIL_COLOR = {"0": "#f9e2af", "1": "#a6e3a1", "2": "#f38ba8"}

def _iface_avail_str(ifaces):
    """
    Возвращает (avail_summary, color, iface_str) для строки хоста.
    avail_summary — наихудший статус из всех интерфейсов.
    iface_str — «Agent·✓ SNMP·✗» и т.п.
    """
    if not ifaces:
        return "Unknown", "#f9e2af", ""
    parts = []
    # Priority: Unavailable > Unknown > Available
    worst = "0"  # Unknown
    for i in ifaces:
        av  = str(i.get("available", "0"))
        typ = IFACE_TYPE.get(str(i.get("type", "1")), "?")
        sym = {"0": "·?", "1": "·✓", "2": "·✗"}.get(av, "")
        main = i.get("main", "0") == "1"
        s = f"{typ}{sym}"
        if main:
            parts.insert(0, s)   # главный — первым
        else:
            parts.append(s)
        # worst: 2 > 0 > 1
        if av == "2":
            worst = "2"
        elif av == "0" and worst != "2":
            worst = "0"
        elif av == "1" and worst == "1":
            worst = "1"
        elif worst not in ("2", "0"):
            worst = "1"
    # Fix: если хотя бы один Available и нет Unavailable → Available
    avs = [str(i.get("available","0")) for i in ifaces]
    if "2" in avs:
        worst = "2"
    elif all(a == "1" for a in avs):
        worst = "1"
    elif "1" in avs:
        worst = "1"   # хотя бы один Available
    else:
        worst = "0"
    color = IFACE_AVAIL_COLOR.get(worst, "#f9e2af")
    return IFACE_AVAIL.get(worst, "Unknown"), color, "  ".join(parts)


# ══════════════════════════════════════════════════════════════════════════════
#  Утилиты
# ══════════════════════════════════════════════════════════════════════════════
def ts2str(ts) -> str:
    try:    return datetime.datetime.fromtimestamp(int(ts)).strftime("%Y-%m-%d %H:%M:%S")
    except: return str(ts)

def dur_str(ts) -> str:
    try:
        d = datetime.datetime.now() - datetime.datetime.fromtimestamp(int(ts))
        h, r = divmod(int(d.total_seconds()), 3600)
        m, s = divmod(r, 60)
        return f"{h}h {m}m {s}s"
    except: return ""

def dur_str_from(start_ts, end_ts) -> str:
    """Длительность между двумя timestamps."""
    try:
        s = max(0, int(end_ts) - int(start_ts))
        d, rem = divmod(s, 86400)
        h, rem = divmod(rem, 3600)
        m, ss  = divmod(rem, 60)
        if d: return f"{d}d {h}h {m}m"
        if h: return f"{h}h {m}m {ss}s"
        if m: return f"{m}m {ss}s"
        return f"{ss}s"
    except Exception:
        return ""



# ══════════════════════════════════════════════════════════════════════════════
#  Сравнение шаблонов — вспомогательные функции
# ══════════════════════════════════════════════════════════════════════════════
import difflib as _difflib

def _name_similarity(a: str, b: str) -> float:
    return _difflib.SequenceMatcher(None, a.lower(), b.lower()).ratio()

def _field_diff(obj_a: dict, obj_b: dict, fields: list) -> str:
    diffs = []
    for f in fields:
        va = str(obj_a.get(f, "")).strip()
        vb = str(obj_b.get(f, "")).strip()
        if va != vb:
            diffs.append(f"{f}: [{va[:35]}] vs [{vb[:35]}]")
    return "; ".join(diffs)

COMPARE_FIELDS = {
    "items":       ["key_","type","value_type","delay","units","status"],
    "triggers":    ["expression","priority","status","recovery_mode"],
    "drules":      ["key_","type","delay","status"],
    "item_protos": ["key_","type","value_type","delay","units","status"],
    "trig_protos": ["expression","priority","status","recovery_mode"],
}

def _compare_section(list_a, list_b, name_field, key_field, diff_fields,
                     similarity_threshold=0.72):
    rows = []
    by_name_a = {o[name_field]: o for o in list_a}
    by_name_b = {o[name_field]: o for o in list_b}
    by_key_a  = {o.get(key_field,"__"): o for o in list_a if o.get(key_field)}
    by_key_b  = {o.get(key_field,"__"): o for o in list_b if o.get(key_field)}
    matched_a = set(); matched_b = set()

    # 1. Exact name match
    for name, oa in by_name_a.items():
        if name in by_name_b:
            ob = by_name_b[name]
            diff = _field_diff(oa, ob, diff_fields)
            rows.append(("Совпадает" if not diff else "Различия",
                         name, name, oa.get(key_field,""), ob.get(key_field,""), diff))
            matched_a.add(name); matched_b.add(name)

    # 2. Key match
    if key_field:
        for key, oa in by_key_a.items():
            if oa[name_field] in matched_a: continue
            if key in by_key_b:
                ob = by_key_b[key]
                if ob[name_field] in matched_b: continue
                diff = _field_diff(oa, ob, diff_fields)
                rows.append(("Различия" if diff else "Совпадает",
                             oa[name_field], ob[name_field],
                             oa.get(key_field,""), ob.get(key_field,""),
                             diff or "Ключи совпадают, имена различаются"))
                matched_a.add(oa[name_field]); matched_b.add(ob[name_field])

    # 3. Fuzzy name match
    unmatched_a = [o for o in list_a if o[name_field] not in matched_a]
    unmatched_b = [o for o in list_b if o[name_field] not in matched_b]
    used_b = set()
    for oa in unmatched_a:
        best_score = 0; best_ob = None
        for ob in unmatched_b:
            if ob[name_field] in used_b: continue
            score = _name_similarity(oa[name_field], ob[name_field])
            if score > best_score:
                best_score = score; best_ob = ob
        if best_ob and best_score >= similarity_threshold:
            diff = _field_diff(oa, best_ob, diff_fields)
            key_diff = "Key различается" if oa.get(key_field,"") != best_ob.get(key_field,"") else ""
            combined = "; ".join(filter(None, [key_diff, diff])) or f"Схожесть: {best_score:.0%}"
            rows.append(("Похожие", oa[name_field], best_ob[name_field],
                         oa.get(key_field,""), best_ob.get(key_field,""), combined))
            matched_a.add(oa[name_field]); used_b.add(best_ob[name_field])

    # 4/5. Only A / Only B
    for oa in list_a:
        if oa[name_field] not in matched_a:
            rows.append(("Только A", oa[name_field], "",
                         oa.get(key_field,""), "", "Отсутствует в шаблоне B"))
    for ob in list_b:
        if ob[name_field] not in matched_b and ob[name_field] not in used_b:
            rows.append(("Только B", "", ob[name_field],
                         "", ob.get(key_field,""), "Отсутствует в шаблоне A"))

    order = {"Различия":0,"Похожие":1,"Только A":2,"Только B":3,"Совпадает":4}
    rows.sort(key=lambda r: order.get(r[0], 9))
    return rows

def _compare_templates(data_a: dict, data_b: dict) -> dict:
    return {
        "items":       _compare_section(data_a["items"],       data_b["items"],       "name",        "key_",       COMPARE_FIELDS["items"]),
        "triggers":    _compare_section(data_a["triggers"],    data_b["triggers"],    "description", "expression", COMPARE_FIELDS["triggers"]),
        "drules":      _compare_section(data_a["drules"],      data_b["drules"],      "name",        "key_",       COMPARE_FIELDS["drules"]),
        "item_protos": _compare_section(data_a["item_protos"], data_b["item_protos"], "name",        "key_",       COMPARE_FIELDS["item_protos"]),
        "trig_protos": _compare_section(data_a["trig_protos"], data_b["trig_protos"], "description", "expression", COMPARE_FIELDS["trig_protos"]),
    }

def _compare_to_csv(folder: str, result: dict):
    import csv, os
    a_name = result.get("_a_name","A"); b_name = result.get("_b_name","B")
    headers = ["Статус", f"Имя ({a_name})", f"Имя ({b_name})",
               f"Key ({a_name})", f"Key ({b_name})", "Различия"]
    labels = {"items":"Items","triggers":"Triggers","drules":"Discovery_Rules",
              "item_protos":"Item_Prototypes","trig_protos":"Trigger_Prototypes"}
    ts = datetime.date.today().isoformat()
    for key, label in labels.items():
        rows = result.get(key, [])
        if not rows: continue
        with open(os.path.join(folder, f"compare_{label}_{ts}.csv"), "w",
                  newline="", encoding="utf-8-sig") as f:
            csv.writer(f).writerows([headers] + list(rows))

def _compare_to_pdf(path: str, result: dict):
    from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer,
                                    Table, TableStyle, PageBreak, HRFlowable)
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.units import mm
    from reportlab.lib import colors
    from reportlab.lib.colors import HexColor
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

    a_name = result.get("_a_name","Шаблон A")
    b_name = result.get("_b_name","Шаблон B")
    doc = SimpleDocTemplate(path, pagesize=landscape(A4),
        topMargin=18*mm, bottomMargin=14*mm, leftMargin=12*mm, rightMargin=12*mm)
    W = landscape(A4)[0] - 24*mm
    ss = getSampleStyleSheet()
    kw      = {"parent": ss["Normal"]}   # no fontName here — set per style
    kw_body = {**kw, "fontName": _PDF_FONT}
    kw_bold = {**kw, "fontName": _PDF_FONT_BOLD}
    S = {
        "title": ParagraphStyle("CT",   **kw_bold, fontSize=16, leading=20, textColor=HexColor("#0F3460")),
        "h2":    ParagraphStyle("CH2",  **kw_bold, fontSize=11, leading=14, textColor=HexColor("#0F3460"), spaceBefore=8),
        "body":  ParagraphStyle("CB",   **kw_body, fontSize=7,  leading=9,  textColor=HexColor("#333333"), wordWrap="CJK"),
        "hdr":   ParagraphStyle("CHdr", **kw_bold, fontSize=8,  leading=10, textColor=colors.white),
    }
    STC = {"Только A": HexColor("#FFCDD2"), "Только B": HexColor("#C8E6C9"),
           "Различия": HexColor("#FFF9C4"), "Похожие":  HexColor("#FFE0B2"),
           "Совпадает":HexColor("#F5F5F5")}
    STX = {"Только A": HexColor("#C62828"), "Только B": HexColor("#2E7D32"),
           "Различия": HexColor("#F57F17"), "Похожие":  HexColor("#E65100"),
           "Совпадает":HexColor("#757575")}

    st = [Paragraph("Сравнение шаблонов Zabbix", S["title"]),
          Paragraph(f"A: {a_name}   vs   B: {b_name}", S["body"]),
          Paragraph(f"Сформирован: {datetime.datetime.now():%Y-%m-%d %H:%M}", S["body"]),
          HRFlowable(width=W, color=HexColor("#E94560"), thickness=2),
          Spacer(1, 4*mm)]

    # Легенда
    legend = [("Только A","Присутствует только в A"),("Только B","Присутствует только в B"),
              ("Различия","Одинаковое имя, но отличаются параметры"),
              ("Похожие","Схожие имена, нечёткое совпадение"),("Совпадает","Полное совпадение")]
    lt = Table([[Paragraph(s,S["body"]),Paragraph(d,S["body"])] for s,d in legend],
               colWidths=[W*0.12, W*0.55])
    lt.setStyle(TableStyle([("GRID",(0,0),(-1,-1),0.3,HexColor("#CCC")),
                             ("TOPPADDING",(0,0),(-1,-1),3),("BOTTOMPADDING",(0,0),(-1,-1),3),
                             ("LEFTPADDING",(0,0),(-1,-1),5)]))
    st += [Paragraph("Легенда", S["h2"]), lt, Spacer(1, 5*mm)]

    LABELS = {"items":"Items","triggers":"Triggers","drules":"Discovery Rules",
              "item_protos":"Item Prototypes","trig_protos":"Trigger Prototypes"}
    COL_W = [9, 22, 22, 19, 19, 29]

    for key, label in LABELS.items():
        rows = result.get(key, [])
        diff_cnt = sum(1 for r in rows if r[0] in ("Только A","Только B","Различия","Похожие"))
        st.append(Paragraph(f"{label}  —  всего: {len(rows)}, отличий: {diff_cnt}", S["h2"]))
        if not rows:
            st += [Paragraph("(нет данных)", S["body"]), Spacer(1,3*mm)]; continue
        hdr = [Paragraph(h, S["hdr"]) for h in ("Статус","Имя A","Имя B","Key A","Key B","Различия")]
        td = [hdr] + [[Paragraph(str(v), S["body"]) for v in r] for r in rows]
        cw = [W*w/sum(COL_W) for w in COL_W]
        ts_ = TableStyle([
            ("BACKGROUND",(0,0),(-1,0),HexColor("#0F3460")),
            ("FONTNAME",(0,0),(-1,0),_PDF_FONT_BOLD),("FONTSIZE",(0,0),(-1,0),8),
            ("ALIGN",(0,0),(-1,0),"CENTER"),
            ("GRID",(0,0),(-1,-1),0.3,HexColor("#CCC")),
            ("TOPPADDING",(0,0),(-1,-1),3),("BOTTOMPADDING",(0,0),(-1,-1),3),
            ("LEFTPADDING",(0,0),(-1,-1),4),("VALIGN",(0,0),(-1,-1),"TOP"),
            ("FONTSIZE",(0,1),(-1,-1),7)])
        for i, r in enumerate(rows):
            if r[0] in STC: ts_.add("BACKGROUND",(0,i+1),(-1,i+1),STC[r[0]])
            if r[0] in STX: ts_.add("TEXTCOLOR",(0,i+1),(0,i+1),STX[r[0]])
        tbl = Table(td, colWidths=cw, repeatRows=1)
        tbl.setStyle(ts_)
        st += [tbl, Spacer(1,4*mm)]
        if key != "trig_protos": st.append(PageBreak())

    def _p(cv, doc):
        cv.saveState()
        w,h = landscape(A4)
        cv.setFillColor(HexColor("#0F3460")); cv.rect(0,h-16*mm,w,16*mm,fill=1,stroke=0)
        cv.setFillColor(colors.white); cv.setFont(_PDF_FONT_BOLD,10)
        cv.drawString(12*mm,h-10*mm,"Zabbix Reporter v3.0 — Сравнение шаблонов")
        cv.setFont(_PDF_FONT,8)
        cv.drawRightString(w-12*mm,h-10*mm,datetime.datetime.now().strftime("%Y-%m-%d %H:%M"))
        cv.setFillColor(HexColor("#0F3460")); cv.rect(0,0,w,8*mm,fill=1,stroke=0)
        cv.setFillColor(colors.white); cv.setFont(_PDF_FONT,7)
        cv.drawCentredString(w/2,2.5*mm,f"Страница {doc.page}")
        cv.restoreState()
    doc.build(st, onFirstPage=_p, onLaterPages=_p)


def _auth_to_pdf(path: str, auth_data: dict, user_dirs: list):
    """
    Генерирует PDF-отчёт по настройкам аутентификации Zabbix.
    Содержит: глобальные настройки, детали каждого LDAP/SAML сервера,
    JIT provisioning (group mapping + media type mapping).
    """
    from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer,
                                    Table, TableStyle, PageBreak, HRFlowable)
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.lib import colors
    from reportlab.lib.colors import HexColor
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

    doc = SimpleDocTemplate(
        path, pagesize=A4,
        topMargin=22*mm, bottomMargin=16*mm,
        leftMargin=15*mm, rightMargin=15*mm)
    W = A4[0] - 30*mm

    ss = getSampleStyleSheet()
    kw_body = {"parent": ss["Normal"], "fontName": _PDF_FONT}
    kw_bold = {"parent": ss["Normal"], "fontName": _PDF_FONT_BOLD}
    S = {
        "title": ParagraphStyle("AT",   **kw_bold, fontSize=18, leading=22,
                                textColor=HexColor("#0F3460")),
        "h1":    ParagraphStyle("AH1",  **kw_bold, fontSize=13, leading=16,
                                textColor=HexColor("#0F3460"), spaceBefore=10),
        "h2":    ParagraphStyle("AH2",  **kw_bold, fontSize=10, leading=13,
                                textColor=HexColor("#E94560"), spaceBefore=6),
        "body":  ParagraphStyle("AB",   **kw_body, fontSize=9,  leading=12,
                                textColor=HexColor("#333333"), wordWrap="CJK"),
        "hdr":   ParagraphStyle("AHdr", **kw_bold, fontSize=9,  leading=11,
                                textColor=colors.white),
        "dim":   ParagraphStyle("ADim", **kw_body, fontSize=8,  leading=10,
                                textColor=HexColor("#888888")),
    }

    YESNO = lambda v: "Да" if str(v) == "1" else "Нет"
    AUTH_TYPE = {"0": "Internal", "1": "LDAP", "2": "HTTP"}
    IDP_TYPE  = {"1": "LDAP",    "2": "SAML"}
    GC_TYPE   = {"1": "memberOf", "2": "groupOfNames"}

    def _tbl(headers, rows, weights):
        total = sum(weights)
        cw    = [W * w / total for w in weights]
        hrow  = [Paragraph(h, S["hdr"]) for h in headers]
        brows = [[Paragraph(str(v) if v is not None else "—", S["body"])
                  for v in row] for row in rows]
        ts = TableStyle([
            ("BACKGROUND",    (0,0),(-1,0), HexColor("#0F3460")),
            ("GRID",          (0,0),(-1,-1), 0.3, HexColor("#CCCCCC")),
            ("ROWBACKGROUNDS",(0,1),(-1,-1),
             [HexColor("#F5F8FF"), HexColor("#FFFFFF")]),
            ("TOPPADDING",    (0,0),(-1,-1), 4),
            ("BOTTOMPADDING", (0,0),(-1,-1), 4),
            ("LEFTPADDING",   (0,0),(-1,-1), 6),
            ("VALIGN",        (0,0),(-1,-1), "MIDDLE"),
            ("FONTSIZE",      (0,1),(-1,-1), 8),
        ])
        t = Table([hrow]+brows, colWidths=cw, repeatRows=1)
        t.setStyle(ts)
        return t

    # ── Сборка документа ──────────────────────────────────────────────────
    st = []

    st += [
        Paragraph("Отчёт: Настройки аутентификации Zabbix", S["title"]),
        Spacer(1, 2*mm),
        Paragraph(f"Сформирован: {datetime.datetime.now():%Y-%m-%d %H:%M:%S}",
                  S["dim"]),
        HRFlowable(width=W, color=HexColor("#E94560"), thickness=2),
        Spacer(1, 5*mm),
    ]

    # ── Глобальные настройки ──────────────────────────────────────────────
    at = auth_data.get("authentication_type","0")
    rules = int(auth_data.get("passwd_check_rules","0"))
    rule_list = ", ".join(filter(None, [
        "Заглавные" if rules & 1 else "",
        "Строчные"  if rules & 2 else "",
        "Цифры"     if rules & 4 else "",
        "Спецсимволы" if rules & 8 else "",
    ])) or "—"

    glob_rows = [
        ("Default authentication",  AUTH_TYPE.get(at, at)),
        ("LDAP auth enabled",       YESNO(auth_data.get("ldap_auth_enabled","0"))),
        ("LDAP JIT provisioning",   YESNO(auth_data.get("ldap_jit_status","0"))),
        ("LDAP JIT interval",       auth_data.get("jit_provision_interval","—")),
        ("LDAP case sensitive",     YESNO(auth_data.get("ldap_case_sensitive","1"))),
        ("HTTP auth enabled",       YESNO(auth_data.get("http_auth_enabled","0"))),
        ("HTTP strip domains",      auth_data.get("http_strip_domains","") or "—"),
        ("HTTP case sensitive",     YESNO(auth_data.get("http_case_sensitive","1"))),
        ("SAML auth enabled",       YESNO(auth_data.get("saml_auth_enabled","0"))),
        ("SAML JIT status",         YESNO(auth_data.get("saml_jit_status","0"))),
        ("SAML case sensitive",     YESNO(auth_data.get("saml_case_sensitive","0"))),
        ("MFA enabled",             YESNO(auth_data.get("mfa_status","0"))),
        ("Password min length",     auth_data.get("passwd_min_length","8")),
        ("Password complexity",     rule_list),
    ]
    st += [
        Paragraph("1. Глобальные настройки аутентификации", S["h1"]),
        Spacer(1, 2*mm),
        _tbl(["Параметр", "Значение"], glob_rows, [55, 35]),
        Spacer(1, 6*mm),
    ]

    if not user_dirs:
        st.append(Paragraph("LDAP / SAML серверы не настроены.", S["body"]))
    else:
        st.append(Paragraph("2. LDAP / SAML серверы", S["h1"]))
        st.append(Spacer(1, 2*mm))
        for idx, d in enumerate(user_dirs, 1):
            idp  = IDP_TYPE.get(str(d.get("idp_type","1")), "?")
            name = d.get("name","") or f"[{idp}]"
            jit  = d.get("provision_status","0") == "1"
            st.append(Paragraph(f"2.{idx}  {idp}: {name}", S["h2"]))

            if str(d.get("idp_type","1")) == "1":  # LDAP
                srv_rows = [
                    ("Host",              d.get("host","")),
                    ("Port",              d.get("port","389")),
                    ("Base DN",           d.get("base_dn","")),
                    ("Bind DN",           d.get("bind_dn","") or "(anonymous)"),
                    ("Bind password",     "***" if d.get("bind_password","") else "(не задан)"),
                    ("StartTLS",          YESNO(d.get("start_tls","0"))),
                    ("Search attribute",  d.get("search_attribute","")),
                    ("JIT Provisioning",  "Включён" if jit else "Отключён"),
                ]
                if jit:
                    gc = GC_TYPE.get(str(d.get("group_configuration","1")),"?")
                    srv_rows += [
                        ("Group configuration", gc),
                        ("Group base DN",       d.get("group_base_dn","") or "—"),
                        ("Group name attr",     d.get("group_name","") or "—"),
                        ("Group member attr",   d.get("group_member","") or "—"),
                        ("User username attr",  d.get("user_username","") or "—"),
                        ("User lastname attr",  d.get("user_lastname","") or "—"),
                        ("User ref attr",       d.get("user_ref_attr","") or "—"),
                        ("Group filter",        d.get("group_filter","") or "—"),
                    ]
            else:  # SAML
                srv_rows = [
                    ("IDP Entity ID",          d.get("idp_entityid","")),
                    ("SSO URL",                d.get("sso_url","")),
                    ("SLO URL",                d.get("slo_url","") or "—"),
                    ("Username attribute",     d.get("username_attribute","")),
                    ("SP Entity ID",           d.get("sp_entityid","")),
                    ("NameID format",          d.get("nameid_format","") or "—"),
                    ("SCIM enabled",           YESNO(d.get("scim_status","0"))),
                    ("JIT Provisioning",       "Включён" if jit else "Отключён"),
                    ("Sign messages",          YESNO(d.get("sign_messages","0"))),
                    ("Sign assertions",        YESNO(d.get("sign_assertions","0"))),
                    ("Sign authn requests",    YESNO(d.get("sign_authn_requests","0"))),
                    ("Sign logout requests",   YESNO(d.get("sign_logout_requests","0"))),
                    ("Sign logout responses",  YESNO(d.get("sign_logout_responses","0"))),
                    ("Encrypt NameID",         YESNO(d.get("encrypt_nameid","0"))),
                    ("Encrypt assertions",     YESNO(d.get("encrypt_assertions","0"))),
                ]

            if d.get("description"):
                srv_rows.append(("Description", d["description"][:120]))

            st += [
                _tbl(["Параметр", "Значение"], srv_rows, [45, 45]),
                Spacer(1, 3*mm),
            ]

            # JIT: Group mapping
            pgs = d.get("provision_groups", [])
            if pgs:
                st.append(Paragraph("User Group Mapping (JIT)", S["h2"]))
                pg_rows = []
                for pg in pgs:
                    role = pg.get("_role_name","") or pg.get("roleid","?")
                    grps = ", ".join(
                        ug.get("_grp_name","") or ug.get("usrgrpid","?")
                        for ug in pg.get("user_groups",[]))
                    pg_rows.append((pg.get("name","*"), role, grps or "—"))
                st += [
                    _tbl(["LDAP group pattern", "Zabbix role", "Zabbix user groups"],
                         pg_rows, [35, 25, 30]),
                    Spacer(1, 3*mm),
                ]

            # JIT: Media type mapping
            pms = d.get("provision_media", [])
            if pms:
                st.append(Paragraph("Media Type Mapping (JIT)", S["h2"]))
                pm_rows = []
                for pm in pms:
                    mt = pm.get("_mt_name","") or pm.get("mediatypeid","?")
                    pm_rows.append((
                        pm.get("name",""), mt,
                        pm.get("attribute",""),
                        YESNO(pm.get("active","1")),
                        pm.get("period","1-7,00:00-24:00"),
                    ))
                st += [
                    _tbl(["Имя", "Media type", "Атрибут LDAP", "Активно", "Период"],
                         pm_rows, [20, 22, 26, 10, 22]),
                    Spacer(1, 4*mm),
                ]

            if idx < len(user_dirs):
                st.append(PageBreak())

    def _on_page(canvas, doc):
        canvas.saveState()
        w, h = A4
        canvas.setFillColor(HexColor("#0F3460"))
        canvas.rect(0, h-18*mm, w, 18*mm, fill=1, stroke=0)
        canvas.setFillColor(colors.white)
        canvas.setFont(_PDF_FONT_BOLD, 10)
        canvas.drawString(15*mm, h-11*mm, "Zabbix Reporter v3.0 — Аутентификация")
        canvas.setFont(_PDF_FONT, 8)
        canvas.drawRightString(w-15*mm, h-11*mm,
                               datetime.datetime.now().strftime("%Y-%m-%d %H:%M"))
        canvas.setFillColor(HexColor("#0F3460"))
        canvas.rect(0, 0, w, 9*mm, fill=1, stroke=0)
        canvas.setFillColor(colors.white)
        canvas.setFont(_PDF_FONT, 7)
        canvas.drawCentredString(w/2, 3*mm,
                                 f"Страница {doc.page}  |  Zabbix Reporter v3.0")
        canvas.restoreState()

    doc.build(st, onFirstPage=_on_page, onLaterPages=_on_page)



def _dupes_to_csv(path: str, groups: list):
    """Сохраняет все группы дубликатов в один CSV файл."""
    import csv
    headers = ["Тип дубликата", "Причина", "Host name", "Видимое имя",
               "Группы", "IP / интерфейсы", "Статус", "Доступность", "Host ID"]
    TYPE_LABELS = {"hostname": "По hostname", "ip": "По IP"}
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(headers)
        for tk_, g in groups:
            label  = TYPE_LABELS.get(tk_, tk_)
            reason = g["reason"]
            for h in g["hosts"]:
                ifaces = h.get("interfaces", [])
                ip_parts = []
                for iface in ifaces:
                    ip = iface.get("ip","").strip()
                    if ip and ip not in ("","0.0.0.0"):
                        itype = IFACE_TYPE.get(str(iface.get("type","1")),"?")
                        ip_parts.append(f"{itype}:{ip}")
                grps     = ", ".join(g2.get("name","")
                                     for g2 in h.get("groups", h.get("hostgroups",[])))
                status   = HOST_STATUS.get(str(h.get("status","0")),"?")
                avail, _, _ = _iface_avail_str(ifaces)
                w.writerow([label, reason,
                            h.get("host",""), h.get("name",""),
                            grps, "  ".join(ip_parts) or "—",
                            status, avail, h.get("hostid","")])


def _dupes_to_pdf(path: str, groups: list, total_hosts: int):
    """Генерирует PDF-отчёт по найденным дубликатам хостов."""
    from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer,
                                    Table, TableStyle, PageBreak, HRFlowable)
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.lib import colors
    from reportlab.lib.colors import HexColor
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

    doc = SimpleDocTemplate(
        path, pagesize=A4,
        topMargin=22*mm, bottomMargin=16*mm,
        leftMargin=14*mm, rightMargin=14*mm)
    W = A4[0] - 28*mm

    ss  = getSampleStyleSheet()
    kb  = {"parent": ss["Normal"], "fontName": _PDF_FONT}
    kbd = {"parent": ss["Normal"], "fontName": _PDF_FONT_BOLD}
    S = {
        "title": ParagraphStyle("DT",  **kbd, fontSize=18, leading=22,
                                textColor=HexColor("#0F3460")),
        "sub":   ParagraphStyle("DS",  **kb,  fontSize=9,  leading=12,
                                textColor=HexColor("#888888")),
        "h1":    ParagraphStyle("DH1", **kbd, fontSize=12, leading=15,
                                textColor=HexColor("#0F3460"), spaceBefore=8),
        "h2":    ParagraphStyle("DH2", **kbd, fontSize=10, leading=13,
                                textColor=HexColor("#E94560"), spaceBefore=5),
        "body":  ParagraphStyle("DB",  **kb,  fontSize=8,  leading=10,
                                textColor=HexColor("#333333"), wordWrap="CJK"),
        "hdr":   ParagraphStyle("DHdr",**kbd, fontSize=8,  leading=10,
                                textColor=colors.white),
        "warn":  ParagraphStyle("DW",  **kb,  fontSize=8,  leading=10,
                                textColor=HexColor("#E65100")),
    }

    TYPE_LABELS  = {"hostname": "По hostname", "ip": "По IP адресу"}
    TYPE_COLORS  = {
        "hostname": HexColor("#FFF9C4"),
        "ip":       HexColor("#FFCDD2"),
    }
    TYPE_HDR = {
        "hostname": HexColor("#F9A825"),
        "ip":       HexColor("#C62828"),
    }

    def _tbl(headers, rows, weights, hdr_color=None):
        total = sum(weights)
        cw    = [W * w / total for w in weights]
        hc    = hdr_color or HexColor("#0F3460")
        hrow  = [Paragraph(h, S["hdr"]) for h in headers]
        brows = [[Paragraph(str(v) if v is not None else "—", S["body"])
                  for v in row] for row in rows]
        ts = TableStyle([
            ("BACKGROUND",    (0,0),(-1,0), hc),
            ("GRID",          (0,0),(-1,-1), 0.3, HexColor("#CCCCCC")),
            ("ROWBACKGROUNDS",(0,1),(-1,-1),
             [HexColor("#FAFAFA"), HexColor("#FFFFFF")]),
            ("TOPPADDING",    (0,0),(-1,-1), 3),
            ("BOTTOMPADDING", (0,0),(-1,-1), 3),
            ("LEFTPADDING",   (0,0),(-1,-1), 5),
            ("VALIGN",        (0,0),(-1,-1), "MIDDLE"),
            ("FONTSIZE",      (0,1),(-1,-1), 7),
        ])
        t = Table([hrow]+brows, colWidths=cw, repeatRows=1)
        t.setStyle(ts)
        return t

    st = []
    by_type   = {}
    for tk_, g in groups:
        by_type.setdefault(tk_, []).append(g)

    total_groups = len(groups)
    total_dup    = sum(len(g["hosts"]) for _, g in groups)

    # ── Титул ──────────────────────────────────────────────────────────────
    st += [
        Paragraph("Отчёт: Дубликаты хостов Zabbix", S["title"]),
        Spacer(1, 2*mm),
        Paragraph(f"Сформирован: {datetime.datetime.now():%Y-%m-%d %H:%M:%S}",
                  S["sub"]),
        Paragraph(f"Всего хостов проанализировано: {total_hosts}  |  "
                  f"Групп дубликатов: {total_groups}  |  "
                  f"Хостов в дубликатах: {total_dup}", S["sub"]),
        HRFlowable(width=W, color=HexColor("#E94560"), thickness=2),
        Spacer(1, 5*mm),
    ]

    # ── Сводка ─────────────────────────────────────────────────────────────
    sum_rows = []
    for tk_, label in TYPE_LABELS.items():
        grps = by_type.get(tk_, [])
        hosts_cnt = sum(len(g["hosts"]) for g in grps)
        sum_rows.append((label, str(len(grps)), str(hosts_cnt)))
    st += [
        Paragraph("Сводка по типам дубликатов", S["h1"]),
        Spacer(1, 2*mm),
        _tbl(["Тип дубликата", "Групп", "Хостов затронуто"],
             sum_rows, [50, 15, 20]),
        Spacer(1, 6*mm),
    ]

    # ── По каждому типу ────────────────────────────────────────────────────
    for tk_, label in TYPE_LABELS.items():
        grp_list = by_type.get(tk_, [])
        if not grp_list:
            continue
        st += [
            PageBreak(),
            Paragraph(f"Тип: {label}", S["h1"]),
            Spacer(1, 3*mm),
        ]
        hdr_c = TYPE_HDR.get(tk_, HexColor("#0F3460"))
        for g in grp_list:
            st.append(Paragraph(f"⚠  {g['reason']}", S["h2"]))
            rows = []
            for h in g["hosts"]:
                ifaces   = h.get("interfaces", [])
                ip_parts = []
                for iface in ifaces:
                    ip = iface.get("ip","").strip()
                    if ip and ip not in ("","0.0.0.0"):
                        itype = IFACE_TYPE.get(str(iface.get("type","1")),"?")
                        ip_parts.append(f"{itype}:{ip}")
                grps_str = ", ".join(
                    g2.get("name","") for g2 in
                    h.get("groups", h.get("hostgroups",[])))
                status   = HOST_STATUS.get(str(h.get("status","0")),"?")
                avail, _, _ = _iface_avail_str(ifaces)
                rows.append((
                    h.get("host",""),
                    h.get("name",""),
                    grps_str[:60],
                    "  ".join(ip_parts) or "—",
                    status,
                    avail,
                ))
            st += [
                _tbl(["Host name","Visible name","Группы",
                      "IP / интерфейсы","Статус","Доступн."],
                     rows, [20, 18, 20, 22, 9, 10], hdr_color=hdr_c),
                Spacer(1, 5*mm),
            ]

    def _on_page(canvas, doc):
        canvas.saveState()
        w, h = A4
        canvas.setFillColor(HexColor("#0F3460"))
        canvas.rect(0, h-18*mm, w, 18*mm, fill=1, stroke=0)
        canvas.setFillColor(colors.white)
        canvas.setFont(_PDF_FONT_BOLD, 10)
        canvas.drawString(14*mm, h-11*mm, "Zabbix Reporter v3.0 — Дубликаты хостов")
        canvas.setFont(_PDF_FONT, 8)
        canvas.drawRightString(w-14*mm, h-11*mm,
                               datetime.datetime.now().strftime("%Y-%m-%d %H:%M"))
        canvas.setFillColor(HexColor("#0F3460"))
        canvas.rect(0, 0, w, 9*mm, fill=1, stroke=0)
        canvas.setFillColor(colors.white)
        canvas.setFont(_PDF_FONT, 7)
        canvas.drawCentredString(w/2, 3*mm,
                                 f"Страница {doc.page}  |  Zabbix Reporter v3.0")
        canvas.restoreState()

    doc.build(st, onFirstPage=_on_page, onLaterPages=_on_page)

# ══════════════════════════════════════════════════════════════════════════════
#  Поиск дубликатов хостов
# ══════════════════════════════════════════════════════════════════════════════
import re as _re

def _normalize_hostname(name: str) -> str:
    """Приводит имя к нижнему регистру для сравнения."""
    return name.strip().lower()

def _short_name(name: str) -> str:
    """Возвращает короткое имя (до первой точки) в нижнем регистре."""
    return name.strip().lower().split(".")[0]

def _is_fqdn(name: str) -> bool:
    """Проверяет, является ли имя FQDN (содержит точку, кроме trailing dot)."""
    return "." in name.strip().rstrip(".")

def find_duplicates(hosts: list) -> dict:
    """
    Находит дубликаты хостов по двум категориям.
    Возвращает:
      {
        "by_hostname": [ {"reason", "key", "hosts": [...]}, ... ],
        "by_ip":       [ ... ],
      }
    Категория by_hostname объединяет дубликаты:
      - имена совпадают без учёта регистра
      - одинаковое short-name, но разный формат (short vs FQDN)
    """
    # ── Дубликаты по имени хоста ──────────────────────────────────────────
    # Ключ группировки: short-name в нижнем регистре.
    # Все хосты с одинаковым short-name — потенциальные дубликаты.
    short_groups: dict = {}   # short_lower -> [host, ...]
    for h in hosts:
        raw = h.get("host", "").strip()
        if not raw:
            continue
        short_groups.setdefault(_short_name(raw), []).append(h)

    by_hostname = []
    for short, group in short_groups.items():
        if len(group) < 2:
            continue
        # Собрать детали группы
        names_exact    = [_normalize_hostname(h["host"]) for h in group]
        has_case_dup   = len(set(names_exact)) < len(names_exact)  # точные дубликаты по регистру
        unique_names   = set(names_exact)
        has_short_fqdn = (
            any(not _is_fqdn(h["host"]) for h in group) and
            any(_is_fqdn(h["host"])      for h in group)
        )
        reasons = []
        if has_case_dup:
            reasons.append("разный регистр")
        if has_short_fqdn:
            reasons.append("short vs FQDN")
        if len(unique_names) > 1 and not reasons:
            reasons.append("разные варианты написания")
        if not reasons and len(unique_names) == 1:
            reasons.append("точные дубликаты")
        by_hostname.append({
            "type":   "По hostname",
            "reason": f"Совпадает short-name «{short}» ({', '.join(reasons)})",
            "key":    short,
            "hosts":  group,
        })

    # ── Дубликаты по IP адресу ────────────────────────────────────────────
    # Группируем по (ip, тип_интерфейса). Пустые IP пропускаем.
    ip_groups: dict = {}
    for h in hosts:
        for iface in h.get("interfaces", []):
            ip = (iface.get("ip") or "").strip()
            # Фильтруем ТОЛЬКО пустые и 0.0.0.0; 127.0.0.1 — валидный
            # IP для локальных агентов и встречается у многих хостов
            if not ip or ip == "0.0.0.0":
                continue
            itype = str(iface.get("type", "1"))
            key   = (ip, itype)
            ip_groups.setdefault(key, []).append((h, iface))

    by_ip = []
    for (ip, itype), pairs in ip_groups.items():
        # Уникальные хосты в группе
        unique_host_ids = []
        seen_ids = set()
        unique_hosts = []
        for h, _ in pairs:
            if h["hostid"] not in seen_ids:
                seen_ids.add(h["hostid"])
                unique_host_ids.append(h["hostid"])
                unique_hosts.append(h)
        if len(unique_hosts) < 2:
            continue
        type_name = IFACE_TYPE.get(itype, f"Type{itype}")
        by_ip.append({
            "type":   "По IP",
            "reason": f"IP {ip} в интерфейсе {type_name}",
            "key":    ip,
            "itype":  itype,
            "hosts":  unique_hosts,
        })

    return {"by_hostname": by_hostname, "by_ip": by_ip}

# ══════════════════════════════════════════════════════════════════════════════
#  Zabbix JSON-RPC API клиент
# ══════════════════════════════════════════════════════════════════════════════
class ZabbixAPI:
    def __init__(self, url: str, verify_ssl: bool = True):
        self.url = url.rstrip("/") + "/api_jsonrpc.php"
        self.auth_token = None
        self.verify_ssl = verify_ssl
        self._rid = 1
        self._api_version = None   # tuple e.g. (7, 4) — set after version()
        self._log_cb = None        # callable(msg, level) for GUI logging

    def _call(self, method: str, params, _no_auth: bool = False):
        """
        _no_auth=True: запрос без авторизации (apiinfo.version, user.login).
        Zabbix 6.0+: токен передаётся в заголовке Authorization: Bearer.
        Zabbix 5.x-: токен передаётся в поле "auth" тела запроса.
        """
        body = {"jsonrpc": "2.0", "method": method,
                "params": params, "id": self._rid}
        self._rid += 1
        headers = {"Content-Type": "application/json-rpc"}
        if not _no_auth and self.auth_token:
            if self._api_version and self._api_version >= (6, 0):
                headers["Authorization"] = f"Bearer {self.auth_token}"
            else:
                body["auth"] = self.auth_token
        if self._log_cb:
            auth_info = ("Bearer ***" if headers.get("Authorization")
                         else ("body.auth" if "auth" in body else "none"))
            self._log_cb(
                f"→ {method}  auth={auth_info}  "
                f"params_keys={list(params.keys()) if isinstance(params, dict) else type(params).__name__}",
                "req"
            )
        req = urllib.request.Request(
            self.url, data=json.dumps(body).encode(),
            headers=headers)
        ctx = ssl.create_default_context()
        if not self.verify_ssl:
            ctx.check_hostname = False
            ctx.verify_mode = ssl.CERT_NONE
        try:
            with urllib.request.urlopen(req, context=ctx, timeout=30) as r:
                res = json.loads(r.read().decode())
        except urllib.error.URLError as e:
            if self._log_cb:
                self._log_cb(f"  ✗ {method} → URLError: {e.reason}", "err")
            raise ConnectionError(f"Ошибка подключения: {e.reason}") from e
        if "error" in res:
            err = res["error"]
            raise RuntimeError(f"API [{err['code']}]: {err['data']}")
        # Краткий лог успешного ответа
        result = res["result"]
        if self._log_cb:
            if isinstance(result, list):
                self._log_cb(f"  ✓ {method} → {len(result)} записей", "ok")
            elif isinstance(result, str):
                self._log_cb(f"  ✓ {method} → {repr(result)[:80]}", "ok")
            else:
                self._log_cb(f"  ✓ {method} → OK", "ok")
        return result

    def login(self, user, pwd):
        # apiinfo.version и user.login вызываются БЕЗ токена авторизации
        ver_str = self._call("apiinfo.version", {}, _no_auth=True)
        try:
            parts = ver_str.split(".")
            self._api_version = tuple(int(x) for x in parts[:2])
        except Exception:
            self._api_version = (6, 0)
        self.auth_token = self._call("user.login",
                                     {"username": user, "password": pwd},
                                     _no_auth=True)

    def logout(self):
        if self.auth_token:
            try: self._call("user.logout", [])
            except: pass
            self.auth_token = None

    def version(self) -> str:
        # apiinfo.version не принимает заголовок Authorization
        ver_str = self._call("apiinfo.version", {}, _no_auth=True)
        try:
            parts = ver_str.split(".")
            self._api_version = tuple(int(x) for x in parts[:2])
        except Exception:
            pass
        return ver_str

    def get_problems(self, sev_min=0, limit=500):
        """
        Zabbix 6.0+: problem.get не поддерживает selectHosts.
        Имя хоста получаем через trigger.get по objectid (triggerid).
        """
        problems = self._call("problem.get", {
            "output": "extend",
            "selectAcknowledges": ["acknowledgeid", "userid", "clock", "message", "action"],
            "selectTags": "extend",
            "severities": list(range(sev_min, 6)),
            "recent": True,
            "sortfield": ["eventid"], "sortorder": "DESC",
            "limit": limit,
        })
        if not problems:
            return problems

        # Получить имена хостов через триггеры (objectid = triggerid для событий типа TRIGGER)
        trigger_ids = list({p["objectid"] for p in problems
                            if p.get("object") in ("0", 0, None)})
        if trigger_ids:
            triggers = self._call("trigger.get", {
                "output": ["triggerid"],
                "selectHosts": ["hostid", "host", "name"],
                "triggerids": trigger_ids[:500],
            })
            trig_map = {t["triggerid"]: t.get("hosts", []) for t in triggers}
            for p in problems:
                p["hosts"] = trig_map.get(p.get("objectid", ""), [])
        return problems

    def get_hosts(self):
        # Try Zabbix 6.2+ selectHostGroups, fall back to selectGroups
        base_params = {
            "output": ["hostid", "host", "name", "status",
                       "description", "maintenance_status", "maintenanceid",
                       "maintenance_type"],
            # In Zabbix 6.0+ availability is per-interface, not per-host.
            # We request "available" on each interface object.
            "selectInterfaces":  ["ip", "dns", "type", "main", "port", "available"],
            "selectTriggers":    "count",
            "sortfield": "name",
        }
        try:
            hosts = self._call("host.get", {**base_params,
                                             "selectHostGroups": ["name"]})
            for h in hosts:
                if "hostgroups" in h and "groups" not in h:
                    h["groups"] = h["hostgroups"]
                elif "groups" not in h:
                    h["groups"] = []
        except Exception as e:
            if "selectHostGroups" in str(e) or "unexpected parameter" in str(e).lower():
                hosts = self._call("host.get", {**base_params,
                                                 "selectGroups": ["name"]})
            else:
                raise

        # Resolve maintenance names if any host is in maintenance
        mids = list({h["maintenanceid"] for h in hosts
                     if h.get("maintenance_status") == "1"
                     and h.get("maintenanceid","0") != "0"})
        mname_map = {}
        if mids:
            try:
                ms = self._call("maintenance.get", {
                    "output": ["maintenanceid", "name"],
                    "maintenanceids": mids,
                })
                mname_map = {m["maintenanceid"]: m["name"] for m in ms}
            except Exception:
                pass
        for h in hosts:
            h["_maint_name"] = (mname_map.get(h.get("maintenanceid",""), "")
                                if h.get("maintenance_status") == "1" else "")
        return hosts

    def get_hosts_detail(self, hostids=None, groupids=None, templateids=None):
        """
        Расширенные данные хостов для вкладки «Объекты мониторинга».
        Возвращает хосты со всеми нужными полями.
        """
        params = {
            "output": ["hostid", "host", "name", "status", "description",
                       "proxy_hostid"],
            "selectInterfaces":      ["interfaceid", "type", "ip", "dns",
                                      "port", "useip", "main"],
            "selectHostGroups":      ["groupid", "name"],   # Zabbix 6.2+
            "selectParentTemplates": ["templateid", "name", "host"],
            "selectTags":            ["tag", "value"],
            "selectMacros":          ["macro", "value", "description", "type"],
            "sortfield": "name",
        }
        if hostids:
            params["hostids"] = hostids
        if groupids:
            params["groupids"] = groupids
        if templateids:
            params["templateids"] = templateids

        try:
            hosts = self._call("host.get", params)
        except Exception as e:
            # Zabbix <6.2 does not support selectHostGroups — fall back
            if "selectHostGroups" in str(e) or "unexpected parameter" in str(e).lower():
                params["selectGroups"] = params.pop("selectHostGroups")
                hosts = self._call("host.get", params)
            else:
                raise
        if not hosts:
            return hosts

        # Normalize: Zabbix 6.2+ returns field "hostgroups", older returns "groups"
        for h in hosts:
            if "hostgroups" in h and "groups" not in h:
                h["groups"] = h["hostgroups"]
            elif "groups" not in h:
                h["groups"] = []

        # Agent version — из item system.uname или agent.version per host
        hids = [h["hostid"] for h in hosts]
        try:
            agent_items = self._call("item.get", {
                "output":  ["hostid", "lastvalue", "lastclock"],
                "hostids": hids,
                "search":  {"key_": "agent.version"},
                "searchWildcardsEnabled": False,
            })
            av_map = {i["hostid"]: i.get("lastvalue","") for i in agent_items}
        except Exception:
            av_map = {}

        # Proxy names
        proxy_ids = list({h["proxy_hostid"] for h in hosts
                          if h.get("proxy_hostid") and h["proxy_hostid"] != "0"})
        proxy_map = {}
        if proxy_ids:
            try:
                proxies = self._call("proxy.get", {
                    "output": ["proxyid", "name"],
                    "proxyids": proxy_ids,
                })
                proxy_map = {p["proxyid"]: p["name"] for p in proxies}
            except Exception:
                # Zabbix <7 uses host.get with proxy_hostid
                try:
                    prx = self._call("host.get", {
                        "output": ["hostid","host"],
                        "hostids": proxy_ids,
                    })
                    proxy_map = {p["hostid"]: p["host"] for p in prx}
                except Exception:
                    pass

        for h in hosts:
            h["_agent_version"] = av_map.get(h["hostid"], "")
            h["_proxy_name"]    = proxy_map.get(h.get("proxy_hostid","0"), "Zabbix Server")

        return hosts

    def get_host_groups_list(self):
        return self._call("hostgroup.get", {
            "output": ["groupid","name"], "sortfield": "name",
        })

    def get_templates_list(self):
        return self._call("template.get", {
            "output": ["templateid","name","host"], "sortfield": "name",
        })

    def get_authentication(self):
        """Глобальные настройки аутентификации Zabbix."""
        return self._call("authentication.get", {"output": "extend"})

    def get_userdirectories(self):
        """
        Все LDAP/SAML user directories с полными данными,
        включая JIT provisioning (provision_groups, provision_media).
        Автоматически резолвит имена ролей, групп и типов медиа.
        """
        try:
            dirs = self._call("userdirectory.get", {
                "output":                "extend",
                "selectProvisionMedia":  "extend",
                "selectProvisionGroups": "extend",
            })
        except Exception as e:
            # Zabbix < 6.4 — provisioning не поддерживается
            if "selectProvision" in str(e) or "unexpected parameter" in str(e).lower():
                dirs = self._call("userdirectory.get", {"output": "extend"})
                for d in dirs:
                    d.setdefault("provision_groups", [])
                    d.setdefault("provision_media",  [])
            else:
                raise

        # Собрать ID для резолвинга имён
        role_ids = set()
        grp_ids  = set()
        mt_ids   = set()
        for d in dirs:
            for pg in d.get("provision_groups", []):
                if pg.get("roleid"):
                    role_ids.add(pg["roleid"])
                for ug in pg.get("user_groups", []):
                    if ug.get("usrgrpid"):
                        grp_ids.add(ug["usrgrpid"])
            for pm in d.get("provision_media", []):
                if pm.get("mediatypeid"):
                    mt_ids.add(pm["mediatypeid"])

        role_map = {}
        if role_ids:
            try:
                roles = self._call("role.get", {
                    "output": ["roleid", "name"], "roleids": list(role_ids)})
                role_map = {r["roleid"]: r["name"] for r in roles}
            except Exception:
                pass

        grp_map = {}
        if grp_ids:
            try:
                grps = self._call("usergroup.get", {
                    "output": ["usrgrpid", "name"], "usrgrpids": list(grp_ids)})
                grp_map = {g["usrgrpid"]: g["name"] for g in grps}
            except Exception:
                pass

        mt_map = {}
        if mt_ids:
            try:
                mts = self._call("mediatype.get", {
                    "output": ["mediatypeid", "name"], "mediatypeids": list(mt_ids)})
                mt_map = {m["mediatypeid"]: m["name"] for m in mts}
            except Exception:
                pass

        # Встроить имена в структуру
        for d in dirs:
            for pg in d.get("provision_groups", []):
                pg["_role_name"] = role_map.get(pg.get("roleid", ""), "")
                for ug in pg.get("user_groups", []):
                    ug["_grp_name"] = grp_map.get(ug.get("usrgrpid", ""), "")
            for pm in d.get("provision_media", []):
                pm["_mt_name"] = mt_map.get(pm.get("mediatypeid", ""), "")

        return dirs

    def get_events(self, time_from, time_till, limit=1000,
                   severities=None, hostids=None):
        """
        Zabbix 6.0+: event.get не поддерживает selectHosts напрямую.
        Имена хостов резолвим через trigger.get по objectid.
        Время восстановления: event.get value:1 не содержит r_clock;
        берём recovery события (value:0) и сопоставляем по r_eventid.
        """
        params = {
            "output": "extend",
            "time_from": time_from, "time_till": time_till,
            "sortfield": ["clock", "eventid"], "sortorder": "DESC",
            "limit": limit, "value": 1,
            "source": 0,
            "selectRelatedObject": ["triggerid"],
        }
        if severities:
            params["severities"] = severities
        if hostids:
            params["hostids"] = hostids
        events = self._call("event.get", params)
        if not events:
            return events

        # Получаем recovery-события чтобы знать r_clock
        event_ids = [e["eventid"] for e in events]
        try:
            recoveries = self._call("event.get", {
                "output":      ["eventid", "clock", "r_eventid"],
                "eventids":    [e.get("r_eventid","0") for e in events
                                if e.get("r_eventid","0") not in ("","0")],
                "source":      0,
            })
            # r_eventid в problem указывает на recovery event;
            # recovery event имеет clock — время восстановления
            rec_map = {r["eventid"]: r.get("clock","0") for r in recoveries}
            for e in events:
                rid = e.get("r_eventid","0")
                if rid and rid != "0" and rid in rec_map:
                    e["r_clock"] = rec_map[rid]
                else:
                    e["r_clock"] = e.get("r_clock","0") or "0"
        except Exception:
            for e in events:
                if "r_clock" not in e:
                    e["r_clock"] = "0"

        # Резолвим хосты через триггеры
        trigger_ids = list({e["objectid"] for e in events if e.get("objectid")})
        if trigger_ids:
            triggers = self._call("trigger.get", {
                "output": ["triggerid"],
                "selectHosts": ["hostid", "host", "name"],
                "triggerids": trigger_ids[:500],
            })
            trig_map = {t["triggerid"]: t.get("hosts", []) for t in triggers}
            for e in events:
                e["hosts"] = trig_map.get(e.get("objectid", ""), [])
        return events

    def get_items(self, host_ids, search="", limit=200):
        params = {
            "output": ["itemid", "name", "key_", "units",
                       "lastvalue", "lastclock", "value_type", "hostid"],
            "hostids": host_ids, "monitored": True,
            "sortfield": "name", "limit": limit,
        }
        if search:
            params["search"] = {"name": search}
        return self._call("item.get", params)

    def get_history(self, item_ids, time_from, time_till,
                    history=0, limit=500):
        return self._call("history.get", {
            "output": "extend", "itemids": item_ids,
            "history": history,
            "time_from": time_from, "time_till": time_till,
            "sortfield": "clock", "sortorder": "ASC",
            "limit": limit,
        })

    def get_template_full(self, templateid: str) -> dict:
        """Полные данные шаблона: items, triggers, discovery rules с прототипами."""
        # Items
        items = self._call("item.get", {
            "output": ["itemid","name","key_","type","value_type",
                       "delay","units","description","status"],
            "templateids": [templateid], "inherited": False,
        })
        # Triggers
        triggers = self._call("trigger.get", {
            "output": ["triggerid","description","expression","priority",
                       "status","recovery_mode","recovery_expression","comments"],
            "templateids": [templateid], "inherited": False,
        })
        # Discovery rules
        drules = self._call("discoveryrule.get", {
            "output": ["itemid","name","key_","type","delay","status","description"],
            "templateids": [templateid], "inherited": False,
        })
        drule_ids = [d["itemid"] for d in drules]
        item_protos = []
        trig_protos = []
        if drule_ids:
            item_protos = self._call("itemprototype.get", {
                "output": ["itemid","name","key_","type","value_type",
                           "delay","units","description","status"],
                "discoveryids": drule_ids, "inherited": False,
            })
            trig_protos = self._call("triggerprototype.get", {
                "output": ["triggerid","description","expression","priority",
                           "status","recovery_mode","recovery_expression"],
                "discoveryids": drule_ids, "inherited": False,
            })
        return {
            "items":       items,
            "triggers":    triggers,
            "drules":      drules,
            "item_protos": item_protos,
            "trig_protos": trig_protos,
        }

    def export_template_yaml(self, templateid: str) -> str:
        """Экспортирует один шаблон в формате YAML (строка)."""
        try:
            return self._call("configuration.export", {
                "format":  "yaml",
                "options": {"templates": [templateid]},
            })
        except Exception as e:
            if "yaml" in str(e).lower() or "format" in str(e).lower():
                # Старые версии Zabbix не поддерживают yaml — fallback на XML
                return self._call("configuration.export", {
                    "format":  "xml",
                    "options": {"templates": [templateid]},
                })
            raise

    def diagnose(self) -> dict:
        """Диагностика: что видит текущий пользователь через API."""
        result = {}
        # Кто я?
        try:
            users = self._call("user.get", {
                "output": ["userid", "username", "roleid"],
                "selectRole": ["roleid", "name", "type"],
                "selectUsrgrps": ["usrgrpid", "name", "gui_access", "users_status"],
                "getAccess": True,
            })
            result["current_user"] = users[0] if users else {}
        except Exception as e:
            result["current_user_error"] = str(e)

        # Группы хостов — видит ли пользователь хоть что-нибудь?
        try:
            hgroups = self._call("hostgroup.get", {
                "output": ["groupid", "name"],
                "sortfield": "name",
                "limit": 50,
            })
            result["host_groups"] = hgroups
        except Exception as e:
            result["host_groups_error"] = str(e)

        # Хосты без фильтров (минимальный запрос)
        try:
            hosts_raw = self._call("host.get", {
                "output": ["hostid", "host", "status"],
                "limit": 10,
            })
            result["hosts_sample"] = hosts_raw
        except Exception as e:
            result["hosts_raw_error"] = str(e)

        # Проблемы без фильтра серьёзности
        try:
            prob_raw = self._call("problem.get", {
                "output": ["eventid", "name", "severity"],
                "recent": True,
                "limit": 10,
            })
            result["problems_sample"] = prob_raw
        except Exception as e:
            result["problems_raw_error"] = str(e)

        # Общее число хостов через host.get count
        try:
            cnt = self._call("host.get", {
                "countOutput": True,
            })
            result["hosts_total_count"] = cnt
        except Exception as e:
            result["hosts_count_error"] = str(e)

        return result


# ══════════════════════════════════════════════════════════════════════════════
#  PDF генератор (ReportLab)
# ══════════════════════════════════════════════════════════════════════════════
class PDFReporter:
    def __init__(self, path):
        self.path = path
        self.styles = getSampleStyleSheet()
        self._add_styles()

    def _add_styles(self):
        s = self.styles
        kw = {"parent": s["Normal"]}
        s.add(ParagraphStyle("ZTitle",    **kw, fontSize=22,
                              leading=28,
                              textColor=HexColor("#0F3460"),
                              spaceBefore=0, spaceAfter=6,
                              fontName=_PDF_FONT_BOLD))
        s.add(ParagraphStyle("ZSub",      **kw, fontSize=11,
                              textColor=HexColor("#555555"), spaceAfter=16,
                              fontName=_PDF_FONT))
        s.add(ParagraphStyle("ZH1",       **kw, fontSize=15,
                              textColor=HexColor("#0F3460"),
                              fontName=_PDF_FONT_BOLD,
                              spaceBefore=16, spaceAfter=7))
        s.add(ParagraphStyle("ZH2",       **kw, fontSize=11,
                              textColor=HexColor("#E94560"),
                              fontName=_PDF_FONT_BOLD,
                              spaceBefore=9, spaceAfter=4))
        s.add(ParagraphStyle("ZBody",     **kw, fontSize=8, leading=11,
                              textColor=HexColor("#222222"), fontName=_PDF_FONT,
                              wordWrap="CJK"))
        s.add(ParagraphStyle("ZCaption",  **kw, fontSize=8,
                              textColor=HexColor("#888888"), alignment=1,
                              fontName=_PDF_FONT))
        s.add(ParagraphStyle("ZBodyHdr",  **kw, fontSize=9,
                              textColor=colors.white,
                              fontName=_PDF_FONT_BOLD))

    def _table(self, headers, rows, col_weights, row_fills=None, page_w=None):
        """
        col_weights: список весов (любые числа > 0), сумма нормируется к page_w.
        page_w: ширина страницы в pt (по умолчанию A4 - поля).
        """
        W = page_w or (A4[0] - 30*mm)
        total = sum(col_weights)
        cw = [W * w / total for w in col_weights]

        hrow = [Paragraph(h, self.styles["ZBodyHdr"]) for h in headers]
        brows = [[Paragraph(str(v) if v is not None else "—", self.styles["ZBody"])
                  for v in row] for row in rows]
        data = [hrow] + brows
        ts = TableStyle([
            ("BACKGROUND",    (0, 0), (-1, 0), HexColor("#0F3460")),
            ("TEXTCOLOR",     (0, 0), (-1, 0), colors.white),
            ("FONTNAME",      (0, 0), (-1, 0), _PDF_FONT_BOLD),
            ("FONTSIZE",      (0, 0), (-1, 0), 9),
            ("ALIGN",         (0, 0), (-1, 0), "CENTER"),
            ("ROWBACKGROUNDS",(0, 1), (-1, -1),
             [HexColor("#F5F8FF"), HexColor("#FFFFFF")]),
            ("GRID",          (0, 0), (-1, -1), 0.35, HexColor("#CCCCCC")),
            ("TOPPADDING",    (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ("LEFTPADDING",   (0, 0), (-1, -1), 5),
            ("RIGHTPADDING",  (0, 0), (-1, -1), 5),
            ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
            ("FONTSIZE",      (0, 1), (-1, -1), 8),
            ("WORDWRAP",      (0, 0), (-1, -1), True),
        ])
        if row_fills:
            for ri, clr in row_fills.items():
                ts.add("BACKGROUND", (0, ri+1), (-1, ri+1), clr)
                ts.add("TEXTCOLOR",  (0, ri+1), (-1, ri+1), colors.white)
        t = Table(data, colWidths=cw, repeatRows=1)
        t.setStyle(ts)
        return t

    def _sev_pie(self, counts):
        items = [(k, v) for k, v in counts.items() if v > 0]
        if not items: return None
        d = Drawing(270, 180)
        pie = Pie(); pie.x = 40; pie.y = 10
        pie.width = pie.height = 150
        pie.data   = [v for _, v in items]
        pie.labels = [f"{SEVERITY_NAMES.get(k,k)} ({v})" for k, v in items]
        pie.sideLabels = True
        for i, (k, _) in enumerate(items):
            c_hex = SEVERITY_HEX.get(k, "#97AAB3")
            pie.slices[i].fillColor = HexColor(c_hex)
        d.add(pie)
        return d

    def _avail_bar(self, counts):
        d = Drawing(260, 150)
        bc = VerticalBarChart()
        bc.x = 35; bc.y = 20; bc.width = 210; bc.height = 110
        labels = list(counts.keys())
        bc.data = [list(counts.values())]
        bc.categoryAxis.categoryNames = labels
        bc.bars[0].fillColor = HexColor("#0F3460")
        bc.valueAxis.valueMin = 0
        d.add(bc)
        return d

    def _metrics_img(self, hist_data, title):
        if not MPL_OK or not hist_data: return None
        import matplotlib
        _prev_backend = matplotlib.get_backend()
        matplotlib.use("Agg")   # switch to non-interactive for file render
        import matplotlib.pyplot as _plt_pdf
        fig, ax = _plt_pdf.subplots(figsize=(7, 2.8))
        fig.patch.set_facecolor("#F8FAFF")
        ax.set_facecolor("#F4F6FF")
        palette = ["#0F3460","#E94560","#457B9D","#2A9D8F","#E9C46A"]
        for i, (name, pts) in enumerate(hist_data.items()):
            if not pts: continue
            xs = [datetime.datetime.fromtimestamp(int(p["clock"])) for p in pts]
            ys = []
            for p in pts:
                try:    ys.append(float(p["value"]))
                except: ys.append(0.0)
            ax.plot(xs, ys, color=palette[i % len(palette)],
                    linewidth=1.8, label=name[:40])
        ax.set_title(title, fontsize=10, color="#0F3460", pad=5)
        ax.tick_params(labelsize=7, colors="#666")
        ax.spines[["top","right"]].set_visible(False)
        ax.spines[["left","bottom"]].set_color("#CCCCCC")
        ax.grid(axis="y", linestyle="--", alpha=0.5, color="#CCCCCC")
        if len(hist_data) > 1:
            ax.legend(fontsize=7, framealpha=0.5)
        buf = io.BytesIO()
        fig.savefig(buf, format="png", dpi=130, bbox_inches="tight")
        _plt_pdf.close(fig)
        try:
            matplotlib.use(_prev_backend)
        except Exception:
            pass
        buf.seek(0)
        return RLImage(buf, width=16*cm, height=6.5*cm)

    @staticmethod
    def _on_page(canvas, doc):
        canvas.saveState()
        w, h = A4
        canvas.setFillColor(HexColor("#0F3460"))
        canvas.rect(0, h - 22*mm, w, 22*mm, fill=1, stroke=0)
        canvas.setFillColor(colors.white)
        canvas.setFont(_PDF_FONT_BOLD, 11)
        canvas.drawString(15*mm, h - 14*mm, "ZABBIX REPORTER v3.0")
        canvas.setFont(_PDF_FONT, 9)
        canvas.drawRightString(w - 15*mm, h - 14*mm,
                               datetime.datetime.now().strftime("%Y-%m-%d %H:%M"))
        canvas.setFillColor(HexColor("#0F3460"))
        canvas.rect(0, 0, w, 10*mm, fill=1, stroke=0)
        canvas.setFillColor(colors.white)
        canvas.setFont(_PDF_FONT, 8)
        canvas.drawCentredString(w/2, 3.5*mm,
                                 f"Страница {doc.page}  |  Zabbix Reporter v3.0")
        canvas.restoreState()

    def build(self, problems, hosts, events, metrics=None, sections=None, hosts_detail=None, auth_data=None, user_dirs=None):
        if sections is None: sections = {'problems','hosts','events','metrics','hosts_detail','auth'}
        doc = SimpleDocTemplate(
            self.path, pagesize=A4,
            topMargin=28*mm, bottomMargin=18*mm,
            leftMargin=15*mm, rightMargin=15*mm,
        )
        s  = self.styles
        W  = A4[0] - 30*mm
        st = []

        # ── Титульная страница ────────────────────────────────────────────────
        gen_ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
        st += [Spacer(1, 10*mm),
               Paragraph("Отчёт мониторинга", s["ZTitle"]),
               Spacer(1, 2*mm),
               Paragraph(f"Сформирован: {gen_ts}", s["ZSub"]),
               HRFlowable(width=W, color=HexColor("#E94560"), thickness=2),
               Spacer(1, 6*mm)]

        sev_cnt = {}
        for p in problems:
            k = str(p.get("severity","0"))
            sev_cnt[k] = sev_cnt.get(k, 0) + 1
        av_cnt = {"Available": 0, "Unavailable": 0, "Unknown": 0}
        for h in hosts:
            av_cnt[HOST_AVAIL.get(str(h.get("available","0")),"Unknown")] += 1

        # Плитки сводки
        tile_data = [
            ["Всего проблем", "Хостов доступно", "Хостов недоступно", "Событий"],
            [str(len(problems)), str(av_cnt["Available"]),
             str(av_cnt["Unavailable"]), str(len(events))],
        ]
        # Tile table: plain strings so TEXTCOLOR works on both rows
        tile_style = TableStyle([
            ("BACKGROUND",   (0,0), (-1,0), HexColor("#E94560")),
            ("TEXTCOLOR",    (0,0), (-1,0), colors.white),
            ("FONTNAME",     (0,0), (-1,0), _PDF_FONT_BOLD),
            ("FONTSIZE",     (0,0), (-1,0), 10),
            ("TOPPADDING",   (0,0), (-1,0), 12),
            ("BOTTOMPADDING",(0,0), (-1,0), 12),
            ("BACKGROUND",   (0,1), (-1,1), HexColor("#0F3460")),
            ("TEXTCOLOR",    (0,1), (-1,1), colors.white),
            ("FONTNAME",     (0,1), (-1,1), _PDF_FONT_BOLD),
            ("FONTSIZE",     (0,1), (-1,1), 26),
            ("TOPPADDING",   (0,1), (-1,1), 14),
            ("BOTTOMPADDING",(0,1), (-1,1), 14),
            ("ALIGN",        (0,0), (-1,-1), "CENTER"),
            ("VALIGN",       (0,0), (-1,-1), "MIDDLE"),
            ("GRID",         (0,0), (-1,-1), 1.5, colors.white),
            ("ROWHEIGHT",    (0,1), (-1,1), 18*mm),
        ])
        tile_t = Table(tile_data, colWidths=[W/4]*4,
                       rowHeights=[None, 18*mm])
        tile_t.setStyle(tile_style)
        st += [tile_t, Spacer(1, 8*mm)]

        pie = self._sev_pie(sev_cnt)
        bar = self._avail_bar(av_cnt)
        if pie and bar:
            chart_row = Table([[pie, bar]], colWidths=[W*0.55, W*0.45])
            chart_row.setStyle(TableStyle([("VALIGN",(0,0),(-1,-1),"TOP")]))
            st += [Paragraph("Серьёзность проблем / Доступность хостов",
                              s["ZH2"]), chart_row, Spacer(1, 4*mm)]
        # ── Нумерация разделов ────────────────────────────────────────────────
        _sec_num = [0]
        def _sn(): _sec_num[0] += 1; return _sec_num[0]

        if 'problems' in sections:
            st.append(PageBreak())
            n = _sn()
            st += [Paragraph(f"{n}. Активные проблемы", s["ZH1"]),
               Paragraph(f"Всего: <b>{len(problems)}</b>", s["ZBody"]),
                   Spacer(1, 3*mm)]
            if problems:
                rfills = {}; rows = []
                for i, p in enumerate(problems[:200]):
                    sid = str(p.get("severity","0"))
                    hl  = p.get("hosts",[{}])
                    hname = hl[0].get("host","") if hl else ""
                    rows.append([SEVERITY_NAMES.get(sid,sid), p.get("name",""),
                                  hname, ts2str(p.get("clock",0)),
                                  dur_str(p.get("clock",0)),
                                  len(p.get("acknowledges",[]))])
                    if sid in ("4","5"):
                        rfills[i] = HexColor(SEVERITY_HEX[sid])
                st.append(self._table(
                    ["Серьёзность","Проблема","Хост","Начало","Длит.","Подт."],
                    rows, [12, 38, 18, 12, 10, 6], rfills))
                if len(problems) > 200:
                    st.append(Paragraph(
                        f"... ещё {len(problems)-200} записей", s["ZCaption"]))

        if 'hosts' in sections:
            st.append(PageBreak())
            n = _sn()
            st += [Paragraph(f"{n}. Статус хостов", s["ZH1"]),
                   Paragraph(f"Всего: <b>{len(hosts)}</b>", s["ZBody"]),
                   Spacer(1, 3*mm)]
            if hosts:
                rows2 = []; rfills2 = {}
                for i, h in enumerate(hosts[:300]):
                    ifaces   = h.get("interfaces",[])
                    avail_str, _, ifaces_str = _iface_avail_str(ifaces)
                    if not ifaces:
                        ac = str(h.get("available","0"))
                        avail_str = HOST_AVAIL.get(ac,"Unknown")
                    in_m  = h.get("maintenance_status","0") == "1"
                    mstr  = h.get("_maint_name","") or ("Да" if in_m else "")
                    rows2.append([h.get("host",""), h.get("name",""),
                                   ", ".join(g["name"] for g in h.get("groups",[])),
                                   HOST_STATUS.get(str(h.get("status","0")),""),
                                   avail_str,
                                   ifaces_str,
                                   str(h.get("triggers","0")),
                                   mstr])
                    if avail_str == "Unavailable":
                        rfills2[i] = HexColor("#E57373")
                    elif in_m:
                        rfills2[i] = HexColor("#F4A261")
                st.append(self._table(
                    ["Хост","Имя","Группа","Статус","Доступн.","Интерфейсы","Тригг.","Обслуж."],
                    rows2, [13, 13, 13, 7, 9, 14, 5, 12], rfills2))

        if 'events' in sections:
            st.append(PageBreak())
            n = _sn()
            st += [Paragraph(f"{n}. История событий", s["ZH1"]),
                   Paragraph(f"Всего: <b>{len(events)}</b>", s["ZBody"]),
                   Spacer(1, 3*mm)]
            if events:
                rows3 = []
                for e in events[:300]:
                    sid = str(e.get("severity","0"))
                    rc  = e.get("r_clock","0")
                    eh  = e.get("hosts",[{}])
                    rows3.append([e.get("eventid",""),
                                   SEVERITY_NAMES.get(sid,sid),
                                   e.get("name",""),
                                   eh[0].get("host","") if eh else "",
                                   ts2str(e.get("clock",0)),
                                   ts2str(rc) if rc not in ("","0") else "-"])
                st.append(self._table(
                    ["ID","Серьёзность","Событие","Хост","Время","Восст."],
                    rows3, [7, 11, 34, 16, 14, 14]))

        if 'hosts_detail' in sections and hosts_detail:
            st.append(PageBreak())
            n = _sn()
            st += [Paragraph(f"{n}. Объекты мониторинга", s["ZH1"]),
                   Paragraph(f"Всего: <b>{len(hosts_detail)}</b>", s["ZBody"]),
                   Spacer(1, 3*mm)]
            ITYPE = {"1":"Agent","2":"SNMP","3":"IPMI","4":"JMX"}
            rows_hd = []
            for h in hosts_detail[:300]:
                ifaces = h.get("interfaces",[])
                main_if = next((i for i in ifaces if i.get("main")=="1"),
                               ifaces[0] if ifaces else {})
                addr    = main_if.get("ip","") or main_if.get("dns","")
                itype   = ITYPE.get(str(main_if.get("type","1")),"")
                av      = h.get("_agent_version","") if itype=="Agent" else ""
                tpls    = "; ".join(t.get("name","") for t in h.get("parentTemplates",[])[:3])
                grps    = "; ".join(g["name"] for g in h.get("groups",[])[:3])
                tags    = "; ".join(
                    f"{tg['tag']}={tg['value']}" if tg.get("value") else tg['tag']
                    for tg in h.get("tags",[])[:4])
                status  = "Да" if str(h.get("status","0"))=="0" else "Нет"
                rows_hd.append([
                    h.get("host",""), h.get("name",""),
                    grps[:60], tpls[:60],
                    f"{itype} {addr}"[:40], av[:20],
                    status, tags[:60],
                    h.get("_proxy_name","Zabbix Server"),
                ])
            st.append(self._table(
                ["Host name","Visible name","Группы","Шаблоны",
                 "Interface","Agent ver","Enabled","Tags","Monitored by"],
                rows_hd,
                [13, 13, 14, 16, 14, 9, 6, 14, 11]))

        if 'auth' in sections and auth_data:
            st.append(PageBreak())
            n = _sn()
            st += [Paragraph(f"{n}. Аутентификация и LDAP", s["ZH1"]),
                   Spacer(1, 3*mm)]
            # Глобальные настройки
            AUTH_TYPE_P = {"0": "Internal", "1": "LDAP", "2": "HTTP"}
            YESNO_P = lambda v: "Да" if str(v)=="1" else "Нет"
            at = auth_data.get("authentication_type","0")
            global_rows = [
                ["Default authentication", AUTH_TYPE_P.get(at, at)],
                ["LDAP JIT provisioning",  YESNO_P(auth_data.get("ldap_jit_status","0"))],
                ["JIT provision interval", auth_data.get("jit_provision_interval","—")],
                ["LDAP case sensitive",    YESNO_P(auth_data.get("ldap_case_sensitive","1"))],
                ["HTTP auth enabled",      YESNO_P(auth_data.get("http_auth_enabled","0"))],
                ["SAML auth enabled",      YESNO_P(auth_data.get("saml_auth_enabled","0"))],
                ["MFA enabled",            YESNO_P(auth_data.get("mfa_status","0"))],
                ["Password min length",    auth_data.get("passwd_min_length","8")],
            ]
            st += [Paragraph("Глобальные настройки", s["ZH2"]),
                   self._table(["Параметр","Значение"], global_rows, [55, 35]),
                   Spacer(1, 5*mm)]
            # LDAP / SAML серверы
            IDP_P = {"1":"LDAP","2":"SAML"}
            for d in (user_dirs or []):
                idp = IDP_P.get(str(d.get("idp_type","1")),"?")
                name = d.get("name","") or f"[{idp}]"
                st += [Paragraph(f"{idp}: {name}", s["ZH2"]), Spacer(1,2*mm)]
                srv_rows = []
                if d.get("idp_type","1") == "1":  # LDAP
                    srv_rows += [
                        ["Host",             d.get("host","")],
                        ["Port",             d.get("port","")],
                        ["Base DN",          d.get("base_dn","")],
                        ["Bind DN",          d.get("bind_dn","") or "(anonymous)"],
                        ["Bind password",    "***" if d.get("bind_password","") else "(не задан)"],
                        ["StartTLS",         YESNO_P(d.get("start_tls","0"))],
                        ["Search attribute", d.get("search_attribute","")],
                        ["JIT Provisioning", "Включён" if d.get("provision_status","0")=="1" else "Отключён"],
                    ]
                    if d.get("provision_status","0") == "1":
                        gc = {"1":"memberOf","2":"groupOfNames"}.get(str(d.get("group_configuration","1")),"?")
                        srv_rows += [
                            ["Group configuration",  gc],
                            ["Group base DN",        d.get("group_base_dn","") or "—"],
                            ["Group name attr",      d.get("group_name","") or "—"],
                            ["Group member attr",    d.get("group_member","") or "—"],
                            ["User username attr",   d.get("user_username","") or "—"],
                            ["User lastname attr",   d.get("user_lastname","") or "—"],
                        ]
                else:  # SAML
                    srv_rows += [
                        ["IDP Entity ID",   d.get("idp_entityid","")],
                        ["SSO URL",         d.get("sso_url","")],
                        ["Username attr",   d.get("username_attribute","")],
                        ["SP Entity ID",    d.get("sp_entityid","")],
                        ["SCIM enabled",    YESNO_P(d.get("scim_status","0"))],
                        ["JIT Provisioning","Включён" if d.get("provision_status","0")=="1" else "Отключён"],
                    ]
                if srv_rows:
                    st.append(self._table(["Параметр","Значение"], srv_rows, [45, 45]))
                # Group mappings
                pgs = d.get("provision_groups", [])
                if pgs:
                    st += [Spacer(1,3*mm), Paragraph("User Group Mapping (JIT)", s["ZH2"])]
                    pg_rows = []
                    for pg in pgs:
                        role = pg.get("_role_name","") or pg.get("roleid","?")
                        grps = ", ".join(ug.get("_grp_name","") or ug.get("usrgrpid","?")
                                         for ug in pg.get("user_groups",[]))
                        pg_rows.append([pg.get("name","*"), role, grps])
                    st.append(self._table(
                        ["LDAP group pattern","Role","Zabbix groups"],
                        pg_rows, [30, 20, 40]))
                # Media mappings
                pms = d.get("provision_media", [])
                if pms:
                    st += [Spacer(1,3*mm), Paragraph("Media Type Mapping (JIT)", s["ZH2"])]
                    pm_rows = []
                    for pm in pms:
                        mt = pm.get("_mt_name","") or pm.get("mediatypeid","?")
                        pm_rows.append([pm.get("name",""), mt,
                                        pm.get("attribute",""),
                                        YESNO_P(pm.get("active","1")),
                                        pm.get("period","1-7,00:00-24:00")])
                    st.append(self._table(
                        ["Имя","Media type","Атрибут LDAP","Активно","Период"],
                        pm_rows, [18, 18, 22, 10, 22]))
                st.append(Spacer(1, 6*mm))

        if 'metrics' in sections and metrics:
            st.append(PageBreak())
            n = _sn()
            st += [Paragraph(f"{n}. Графики метрик", s["ZH1"]), Spacer(1, 3*mm)]
            for title, hdata in metrics.items():
                img = self._metrics_img(hdata, title)
                if img:
                    st += [Paragraph(title, s["ZH2"]), img, Spacer(1, 5*mm)]

        doc.build(st, onFirstPage=self._on_page, onLaterPages=self._on_page)


# ══════════════════════════════════════════════════════════════════════════════
#  Excel генератор (openpyxl)
# ══════════════════════════════════════════════════════════════════════════════
class ExcelReporter:
    _HF  = Font(bold=True, color="FFFFFF", size=10, name="Segoe UI")
    _HFL = PatternFill("solid", fgColor="0F3460")
    _TH  = Side(style="thin", color="CCCCCC")
    _BRD = Border(**{s: Side(style="thin", color="CCCCCC")
                     for s in ("left","right","top","bottom")})
    _ALT = [PatternFill("solid", fgColor="F5F8FF"),
            PatternFill("solid", fgColor="FFFFFF")]
    _SEV = {
        "0": PatternFill("solid", fgColor="ECEFF1"),
        "1": PatternFill("solid", fgColor="BBDEFB"),
        "2": PatternFill("solid", fgColor="FFF9C4"),
        "3": PatternFill("solid", fgColor="FFE0B2"),
        "4": PatternFill("solid", fgColor="FFCCBC"),
        "5": PatternFill("solid", fgColor="EF9A9A"),
    }

    def __init__(self, path):
        self.path = path
        self.wb   = openpyxl.Workbook()

    def _sheet(self, ws, headers, rows, widths, sev_col=None):
        ws.row_dimensions[1].height = 22
        for ci, (h, w) in enumerate(zip(headers, widths), 1):
            c = ws.cell(1, ci, h)
            c.font = self._HF; c.fill = self._HFL
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = self._BRD
            ws.column_dimensions[get_column_letter(ci)].width = w
        for ri, row in enumerate(rows, 2):
            sid = None
            if sev_col is not None:
                raw = str(row[sev_col])
                for k, v in SEVERITY_NAMES.items():
                    if v == raw or k == raw: sid = k; break
            fill = (self._SEV.get(sid) if sid else self._ALT[(ri-2) % 2])
            font = Font(size=9, name="Segoe UI",
                        bold=(sid in ("4","5") if sid else False))
            for ci, val in enumerate(row, 1):
                c = ws.cell(ri, ci, str(val) if val is not None else "")
                c.fill = fill; c.font = font; c.border = self._BRD
                c.alignment = Alignment(
                    horizontal="center" if ci == 1 else "left",
                    vertical="center", wrap_text=True)
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

    def build(self, problems, hosts, events, metrics=None, sections=None, hosts_detail=None, auth_data=None, user_dirs=None):
        if sections is None: sections = {'problems','hosts','events','metrics','hosts_detail','auth'}
        if hosts_detail is None: hosts_detail = []
        if user_dirs    is None: user_dirs    = []
        # ── Сводка ───────────────────────────────────────────────────────────
        ws0 = self.wb.active; ws0.title = "Сводка"
        ws0.sheet_view.showGridLines = False
        ws0["A1"].value = "ZABBIX REPORTER — Сводный отчёт"
        ws0["A1"].font  = Font(bold=True, size=16, color="0F3460", name="Segoe UI")
        ws0.merge_cells("A1:D1")
        ws0["A2"].value = f"Сформирован: {datetime.datetime.now():%Y-%m-%d %H:%M:%S}"
        ws0["A2"].font  = Font(size=9, italic=True, color="888888", name="Segoe UI")
        ws0.merge_cells("A2:D2")
        sev_cnt = {}
        for p in problems:
            k = str(p.get("severity","0")); sev_cnt[k] = sev_cnt.get(k,0)+1
        av_cnt = {"Available":0,"Unavailable":0,"Unknown":0}
        for h in hosts: av_cnt[HOST_AVAIL.get(str(h.get("available","0")),"Unknown")] += 1
        stats = ([("Всего проблем", len(problems)), ("","")] +
                 [(f"  {SEVERITY_NAMES.get(k,'?')}", v)
                  for k, v in sorted(sev_cnt.items())] +
                 [("",""),("Всего хостов", len(hosts))] +
                 [(f"  {k}", v) for k, v in av_cnt.items()] +
                 [("",""),("Событий в периоде", len(events))])
        for ri, (lbl, val) in enumerate(stats, 4):
            ca = ws0.cell(ri, 1, lbl); cb = ws0.cell(ri, 2, val)
            bold = bool(lbl and not lbl.startswith(" "))
            fill = PatternFill("solid", fgColor=("0F3460" if bold else
                               ("F0F4FF" if ri%2==0 else "FFFFFF")))
            fc   = "FFFFFF" if bold else "333333"
            for c in (ca, cb):
                c.font = Font(bold=bold, size=10 if bold else 9,
                              color=fc, name="Segoe UI")
                c.fill = fill
        ws0.column_dimensions["A"].width = 28
        ws0.column_dimensions["B"].width = 14

        # ── Проблемы ─────────────────────────────────────────────────────────
        if 'problems' in sections:
            ws1 = self.wb.create_sheet("Проблемы")
            self._sheet(ws1,
                ["Серьёзность","Проблема","Хост","Начало","Длительность","Подтверждений"],
                [(SEVERITY_NAMES.get(str(p.get("severity","0")),""),
                  p.get("name",""),
                  (p.get("hosts",[{}])[0].get("host","") if p.get("hosts") else ""),
                  ts2str(p.get("clock",0)), dur_str(p.get("clock",0)),
                  len(p.get("acknowledges",[]))) for p in problems],
                [20,50,26,22,18,14], sev_col=0)
            # Диаграмма серьёзностей
            if sev_cnt:
                ws_c = self.wb.create_sheet("График серьёзностей")
                ws_c["A1"].value = "Серьёзность"; ws_c["B1"].value = "Кол-во"
                for i, (k, v) in enumerate(sorted(sev_cnt.items()), 2):
                    ws_c.cell(i, 1, SEVERITY_NAMES.get(k, k)); ws_c.cell(i, 2, v)
                ch = BarChart(); ch.type = "col"; ch.grouping = "clustered"
                ch.title = "Проблемы по серьёзности"
                ch.y_axis.title = "Кол-во"; ch.x_axis.title = "Серьёзность"
                ch.width = 18; ch.height = 12
                dr = Reference(ws_c, min_col=2, min_row=1, max_row=len(sev_cnt)+1)
                ct = Reference(ws_c, min_col=1, min_row=2, max_row=len(sev_cnt)+1)
                ch.add_data(dr, titles_from_data=True); ch.set_categories(ct)
                ws_c.add_chart(ch, "D2")

        # ── Хосты ────────────────────────────────────────────────────────────
        if 'hosts' in sections:
            ws2 = self.wb.create_sheet("Хосты")
            def _xl_host_row(h):
                ifaces    = h.get("interfaces",[])
                avail_str, _, ifaces_str = _iface_avail_str(ifaces)
                if not ifaces:
                    ac = str(h.get("available","0"))
                    avail_str = HOST_AVAIL.get(ac,"Unknown")
                in_m  = h.get("maintenance_status","0") == "1"
                mstr  = h.get("_maint_name","") or ("Да" if in_m else "")
                grps  = h.get("groups", h.get("hostgroups", []))
                return (h.get("host",""), h.get("name",""),
                        ", ".join(g["name"] for g in grps),
                        HOST_STATUS.get(str(h.get("status","0")),""),
                        avail_str, ifaces_str,
                        h.get("triggers","0"), mstr)
            self._sheet(ws2,
                ["Хост","Имя","Группа","Статус","Доступность",
                 "Интерфейсы","Триггеры","Обслуживание"],
                [_xl_host_row(h) for h in hosts],
                [22,24,24,14,16,28,10,20])

        # ── События ───────────────────────────────────────────────────────────
        if 'events' in sections:
            ws3 = self.wb.create_sheet("События")
            self._sheet(ws3,
                ["Event ID","Серьёзность","Событие","Хост","Время","Восстановлено"],
                [(e.get("eventid",""),
                  SEVERITY_NAMES.get(str(e.get("severity","0")),""),
                  e.get("name",""),
                  (e.get("hosts",[{}])[0].get("host","") if e.get("hosts") else ""),
                  ts2str(e.get("clock",0)),
                  (ts2str(e.get("r_clock","0"))
                   if e.get("r_clock","0") not in ("","0") else "-"))
                 for e in events],
                [14,20,56,26,22,22], sev_col=1)

        # ── Объекты мониторинга ─────────────────────────────────────────────
        if 'hosts_detail' in sections and hosts_detail:
            ws_hd = self.wb.create_sheet("Объекты мониторинга")
            ITYPE = {"1":"Agent","2":"SNMP","3":"IPMI","4":"JMX"}
            hd_rows = []
            for h in hosts_detail:
                ifaces = h.get("interfaces",[])
                main_if = next((i for i in ifaces if i.get("main")=="1"),
                               ifaces[0] if ifaces else {})
                addr   = main_if.get("ip","") or main_if.get("dns","")
                port   = main_if.get("port","")
                itype  = ITYPE.get(str(main_if.get("type","1")),"")
                all_ifaces = "; ".join(
                    f"{ITYPE.get(str(i.get('type','1')),'?')} "
                    f"{i.get('ip','') or i.get('dns','')}:{i.get('port','')}"
                    for i in ifaces)
                av     = h.get("_agent_version","") if itype=="Agent" else ""
                tpls   = "; ".join(t.get("name","") for t in h.get("parentTemplates",[]))
                grps   = "; ".join(g["name"] for g in h.get("groups",[]))
                tags   = "; ".join(
                    f"{tg['tag']}={tg['value']}" if tg.get("value") else tg['tag']
                    for tg in h.get("tags",[]))
                macros = "; ".join(
                    f"{m['macro']}={'***' if str(m.get('type','0'))=='1' else m.get('value','')}"
                    for m in h.get("macros",[]))
                status = "Да" if str(h.get("status","0"))=="0" else "Нет"
                hd_rows.append((
                    h.get("host",""), h.get("name",""),
                    grps, tpls, all_ifaces, av,
                    status, h.get("_proxy_name","Zabbix Server"),
                    tags, macros,
                ))
            self._sheet(ws_hd,
                ["Host name","Visible name","Host groups","Templates",
                 "Interfaces","Agent version","Enabled","Monitored by",
                 "Tags","Macros"],
                hd_rows,
                [24,24,32,36,40,20,10,24,32,36])

        # ── Аутентификация ────────────────────────────────────────────────────
        if 'auth' in sections and auth_data:
            YESNO_XL = lambda v: "Да" if str(v)=="1" else "Нет"
            AUTH_TYPE_XL = {"0":"Internal","1":"LDAP","2":"HTTP"}
            IDP_XL = {"1":"LDAP","2":"SAML"}

            # Лист: глобальные настройки
            ws_auth = self.wb.create_sheet("Аутентификация")
            at = auth_data.get("authentication_type","0")
            auth_global_rows = [
                ("Default authentication", AUTH_TYPE_XL.get(at,at)),
                ("LDAP JIT provisioning",  YESNO_XL(auth_data.get("ldap_jit_status","0"))),
                ("JIT provision interval", auth_data.get("jit_provision_interval","—")),
                ("LDAP case sensitive",    YESNO_XL(auth_data.get("ldap_case_sensitive","1"))),
                ("HTTP auth enabled",      YESNO_XL(auth_data.get("http_auth_enabled","0"))),
                ("HTTP strip domains",     auth_data.get("http_strip_domains","") or "—"),
                ("HTTP case sensitive",    YESNO_XL(auth_data.get("http_case_sensitive","1"))),
                ("SAML auth enabled",      YESNO_XL(auth_data.get("saml_auth_enabled","0"))),
                ("SAML JIT status",        YESNO_XL(auth_data.get("saml_jit_status","0"))),
                ("SAML case sensitive",    YESNO_XL(auth_data.get("saml_case_sensitive","0"))),
                ("MFA enabled",            YESNO_XL(auth_data.get("mfa_status","0"))),
                ("Password min length",    auth_data.get("passwd_min_length","8")),
                ("Password check rules",   auth_data.get("passwd_check_rules","0")),
            ]
            self._sheet(ws_auth, ["Параметр","Значение"], auth_global_rows, [36,24])

            # Лист: LDAP серверы (детали)
            for d in user_dirs:
                idp = IDP_XL.get(str(d.get("idp_type","1")),"?")
                safe_name = (d.get("name","") or idp)[:24].replace("/","-")
                ws_ud = self.wb.create_sheet(f"{idp}:{safe_name}"[:31])

                # Основные поля
                ud_rows = []
                if d.get("idp_type","1") == "1":
                    ud_rows = [
                        ("Name",             d.get("name","")),
                        ("Host",             d.get("host","")),
                        ("Port",             d.get("port","")),
                        ("Base DN",          d.get("base_dn","")),
                        ("Bind DN",          d.get("bind_dn","") or "(anonymous)"),
                        ("Bind password",    "***" if d.get("bind_password","") else "(не задан)"),
                        ("StartTLS",         YESNO_XL(d.get("start_tls","0"))),
                        ("Search attribute", d.get("search_attribute","")),
                        ("JIT Provisioning", "Включён" if d.get("provision_status","0")=="1" else "Отключён"),
                        ("Group configuration",
                         {"1":"memberOf","2":"groupOfNames"}.get(str(d.get("group_configuration","1")),"?")),
                        ("Group base DN",        d.get("group_base_dn","") or "—"),
                        ("Group name attr",      d.get("group_name","") or "—"),
                        ("Group member attr",    d.get("group_member","") or "—"),
                        ("User username attr",   d.get("user_username","") or "—"),
                        ("User lastname attr",   d.get("user_lastname","") or "—"),
                        ("User ref attr",        d.get("user_ref_attr","") or "—"),
                        ("Group filter",         d.get("group_filter","") or "—"),
                        ("Description",          d.get("description","") or "—"),
                    ]
                else:
                    ud_rows = [
                        ("Name",               d.get("name","")),
                        ("IDP Entity ID",      d.get("idp_entityid","")),
                        ("SSO URL",            d.get("sso_url","")),
                        ("SLO URL",            d.get("slo_url","") or "—"),
                        ("Username attribute", d.get("username_attribute","")),
                        ("SP Entity ID",       d.get("sp_entityid","")),
                        ("NameID format",      d.get("nameid_format","") or "—"),
                        ("SCIM enabled",       YESNO_XL(d.get("scim_status","0"))),
                        ("JIT Provisioning",   "Включён" if d.get("provision_status","0")=="1" else "Отключён"),
                        ("Sign messages",          YESNO_XL(d.get("sign_messages","0"))),
                        ("Sign assertions",        YESNO_XL(d.get("sign_assertions","0"))),
                        ("Sign authn requests",    YESNO_XL(d.get("sign_authn_requests","0"))),
                        ("Sign logout requests",   YESNO_XL(d.get("sign_logout_requests","0"))),
                        ("Sign logout responses",  YESNO_XL(d.get("sign_logout_responses","0"))),
                        ("Encrypt NameID",         YESNO_XL(d.get("encrypt_nameid","0"))),
                        ("Encrypt assertions",     YESNO_XL(d.get("encrypt_assertions","0"))),
                        ("Group name attr",        d.get("group_name","") or "—"),
                        ("User username attr",     d.get("user_username","") or "—"),
                        ("User lastname attr",     d.get("user_lastname","") or "—"),
                        ("Description",            d.get("description","") or "—"),
                    ]
                self._sheet(ws_ud, ["Параметр","Значение"], ud_rows, [36,40])

                # Лист Group mappings
                pgs = d.get("provision_groups", [])
                if pgs:
                    safe2 = safe_name[:20]
                    ws_pg = self.wb.create_sheet(f"{idp}:{safe2} Groups"[:31])
                    pg_rows = []
                    for pg in pgs:
                        role = pg.get("_role_name","") or pg.get("roleid","?")
                        for ug in pg.get("user_groups", []):
                            grp = ug.get("_grp_name","") or ug.get("usrgrpid","?")
                            pg_rows.append((pg.get("name","*"), role, grp))
                    if pg_rows:
                        self._sheet(ws_pg,
                            ["LDAP group pattern","Zabbix role","Zabbix user group"],
                            pg_rows, [36,24,24])

                # Лист Media mappings
                pms = d.get("provision_media", [])
                if pms:
                    safe3 = safe_name[:20]
                    ws_pm = self.wb.create_sheet(f"{idp}:{safe3} Media"[:31])
                    pm_rows = []
                    for pm in pms:
                        mt = pm.get("_mt_name","") or pm.get("mediatypeid","?")
                        pm_rows.append((
                            pm.get("name",""), mt,
                            pm.get("attribute",""),
                            YESNO_XL(pm.get("active","1")),
                            pm.get("severity",""),
                            pm.get("period",""),
                        ))
                    self._sheet(ws_pm,
                        ["Имя","Media type","Атрибут","Активно","Severity","Период"],
                        pm_rows, [22,22,28,10,12,22])

        # ── Метрики ───────────────────────────────────────────────────────────
        if 'metrics' in sections and metrics and MPL_OK:
            for title, hdata in metrics.items():
                safe = title[:28].replace("/","-").replace("\\","-")
                ws_m = self.wb.create_sheet(f"М:{safe}"[:31])
                ws_m["A1"].value = title
                ws_m["A1"].font  = Font(bold=True,size=12,
                                        color="0F3460",name="Segoe UI")
                r = 3
                for name, pts in hdata.items():
                    for ci, hdr in enumerate(["Метрика","Время","Значение"],1):
                        c = ws_m.cell(r, ci, hdr)
                        c.font = Font(bold=True,color="FFFFFF",
                                      size=9,name="Segoe UI")
                        c.fill = PatternFill("solid",fgColor="0F3460")
                    r += 1; start = r
                    for pt in pts:
                        ws_m.cell(r,1,name[:40])
                        ws_m.cell(r,2,ts2str(pt["clock"]))
                        try: ws_m.cell(r,3,float(pt["value"]))
                        except: ws_m.cell(r,3,pt["value"])
                        r += 1
                    if len(pts) > 1:
                        lc = LineChart(); lc.title = name[:40]
                        lc.y_axis.title = "Значение"
                        lc.width = 18; lc.height = 9
                        vals = Reference(ws_m, min_col=3,
                                          min_row=start-1, max_row=r-1)
                        lc.add_data(vals, titles_from_data=True)
                        ws_m.add_chart(lc, f"E{start-1}")
                    r += 2

        self.wb.save(self.path)


# ══════════════════════════════════════════════════════════════════════════════
#  GUI — Главное окно
# ══════════════════════════════════════════════════════════════════════════════
# ══════════════════════════════════════════════════════════════════════════════
#  Конфиг (config.json рядом со скриптом)
# ══════════════════════════════════════════════════════════════════════════════
_CFG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")

def _cfg_load() -> dict:
    try:
        with open(_CFG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}

def _cfg_save(data: dict) -> None:
    try:
        with open(_CFG_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


# ══════════════════════════════════════════════════════════════════════════════
#  Всплывающее окно логирования
# ══════════════════════════════════════════════════════════════════════════════
class LogWindow(tk.Toplevel):
    TAG_COLORS = {
        "req":  "#89b4fa", "ok":   "#a6e3a1", "err":  "#f38ba8",
        "warn": "#f9e2af", "info": "#cdd6f4",  "ts":   "#6c7086",
    }

    def __init__(self, master):
        super().__init__(master)
        self.title("Zabbix Reporter — Лог / Отладка")
        self.geometry("1020x540")
        self.minsize(500, 240)
        self.configure(bg="#11111b")
        self._master = master
        self._line_count = 0
        self._build()
        self._load_history()

    def _build(self):
        # Тулбар
        tb = tk.Frame(self, bg="#11111b", pady=5)
        tb.pack(fill="x", padx=8)

        def tbtn(text, cmd):
            b = tk.Button(tb, text=text, command=cmd,
                          bg="#1e1e2e", fg="#cdd6f4", relief="flat",
                          font=("Segoe UI", 9), cursor="hand2",
                          padx=10, pady=4,
                          activebackground="#313244", activeforeground="#cdd6f4")
            b.pack(side="left", padx=2)
            return b

        tbtn("🗑  Очистить",    self.clear)
        tbtn("💾  Сохранить",   self._save)
        tbtn("🩺  Диагностика", self._master._run_diagnostics)

        tk.Label(tb, text="  Уровень:", bg="#11111b",
                 fg="#6c7086", font=("Segoe UI", 9)).pack(side="left", padx=(14, 2))
        self._flt = tk.StringVar(value="все")
        cb = ttk.Combobox(tb, textvariable=self._flt, state="readonly",
                          values=["все", "ok", "err", "warn", "req", "info"], width=7)
        cb.pack(side="left")
        cb.bind("<<ComboboxSelected>>", lambda _: self._apply_filter())

        tk.Label(tb, text="  🔍", bg="#11111b",
                 fg="#6c7086", font=("Segoe UI", 9)).pack(side="left", padx=(12, 0))
        self._srch = tk.StringVar()
        self._srch.trace_add("write", lambda *_: self._apply_filter())
        ttk.Entry(tb, textvariable=self._srch, width=24).pack(side="left", padx=4)

        self._cnt_lbl = tk.Label(tb, text="", bg="#11111b",
                                  fg="#6c7086", font=("Segoe UI", 8))
        self._cnt_lbl.pack(side="right", padx=8)

        ttk.Separator(self).pack(fill="x")

        # Текстовое поле
        frm = tk.Frame(self, bg="#11111b")
        frm.pack(fill="both", expand=True, padx=4, pady=4)
        self._txt = tk.Text(
            frm, bg="#11111b", fg="#a6adc8",
            font=("Consolas", 10), relief="flat",
            state="disabled", wrap="none",
            padx=10, pady=6,
            selectbackground="#313244",
        )
        vs = ttk.Scrollbar(frm, orient="vertical",   command=self._txt.yview)
        hs = ttk.Scrollbar(frm, orient="horizontal", command=self._txt.xview)
        self._txt.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        self._txt.grid(row=0, column=0, sticky="nsew")
        vs.grid(row=0, column=1, sticky="ns")
        hs.grid(row=1, column=0, sticky="ew")
        frm.rowconfigure(0, weight=1); frm.columnconfigure(0, weight=1)
        for tag, color in self.TAG_COLORS.items():
            self._txt.tag_configure(
                tag, foreground=color,
                font=("Consolas", 10, "bold") if tag == "err" else ("Consolas", 10))

        # Статусбар
        sb = tk.Frame(self, bg="#0d0d17", height=22)
        sb.pack(fill="x"); sb.pack_propagate(False)
        self._sb_lbl = tk.Label(sb, text="", bg="#0d0d17",
                                 fg="#6c7086", font=("Segoe UI", 8), anchor="w")
        self._sb_lbl.pack(side="left", padx=8, pady=2)

    def _load_history(self):
        buf = self._master._log_buf
        self._txt.configure(state="normal")
        raw = buf.get("1.0", "end")
        self._txt.insert("end", raw)
        for tag in self.TAG_COLORS:
            ranges = buf.tag_ranges(tag)
            for i in range(0, len(ranges), 2):
                try: self._txt.tag_add(tag, str(ranges[i]), str(ranges[i+1]))
                except: pass
        self._line_count = max(0, int(self._txt.index("end-1c").split(".")[0]) - 1)
        self._txt.see("end")
        self._txt.configure(state="disabled")
        self._upd_sb()

    def append(self, text: str, tag: str = "info"):
        self._txt.configure(state="normal")
        self._txt.insert("end", text, tag)
        if text.endswith("\n"):
            self._line_count += 1
        self._txt.see("end")
        self._txt.configure(state="disabled")
        self._upd_sb()

    def clear(self):
        self._master._log_buf.delete("1.0", "end")
        self._txt.configure(state="normal")
        self._txt.delete("1.0", "end")
        self._txt.configure(state="disabled")
        self._line_count = 0
        self._upd_sb()

    def _apply_filter(self):
        lvl = self._flt.get()
        q   = self._srch.get().lower()
        buf = self._master._log_buf
        lines = buf.get("1.0", "end").splitlines()
        self._txt.configure(state="normal")
        self._txt.delete("1.0", "end")
        shown = 0
        for line in lines:
            if not line: continue
            if q and q not in line.lower(): continue
            self._txt.insert("end", line + "\n")
            shown += 1
        for tag in self.TAG_COLORS:
            ranges = buf.tag_ranges(tag)
            for i in range(0, len(ranges), 2):
                try: self._txt.tag_add(tag, str(ranges[i]), str(ranges[i+1]))
                except: pass
        self._txt.see("end")
        self._txt.configure(state="disabled")
        self._cnt_lbl.configure(text=f"{shown} строк")

    def _save(self):
        path = filedialog.asksaveasfilename(
            parent=self, defaultextension=".txt",
            filetypes=[("Text", "*.txt")],
            initialfile=f"zabbix_log_{datetime.date.today()}.txt")
        if path:
            with open(path, "w", encoding="utf-8") as f:
                f.write(self._master._log_buf.get("1.0", "end"))

    def _upd_sb(self):
        ver = self._master._lbl_ver.cget("text")
        self._sb_lbl.configure(
            text=f"Строк: {self._line_count}" +
                 (f"  |  Zabbix {ver}" if ver else ""))
        self._cnt_lbl.configure(text=f"{self._line_count} строк")


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("⚡ Zabbix Reporter v3.0")
        self.geometry("1540x860")
        self.minsize(1200, 700)
        self.configure(bg="#1e1e2e")

        self.zapi          = None
        self._problems     = []
        self._hosts        = []
        self._events       = []
        self._metrics      = {}
        self._items_cache  = []
        self._hosts_detail    = []  # расширенные данные хостов
        self._hosts_latest_ok = {}  # hostid -> True/False (свежие данные)
        self._auth_data    = {}   # authentication.get
        self._user_dirs    = []   # userdirectory.get
        self._log_win      = None   # окно лога (Toplevel)

        self._cfg = _cfg_load()
        self._obj_grp_map = {}   # name -> groupid
        self._obj_tpl_map = {}   # name -> templateid

        self._build_style()
        self._build_ui()
        self._cfg_to_ui()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ── Стиль ──────────────────────────────────────────────────────────────────
    def _build_style(self):
        s = ttk.Style(self); s.theme_use("clam")
        BG="#1e1e2e"; FG="#cdd6f4"; ACC="#89b4fa"; ENT="#313244"; SEL="#45475a"
        s.configure(".", background=BG, foreground=FG,
                     fieldbackground=ENT, font=("Segoe UI",10))
        s.configure("TFrame",       background=BG)
        s.configure("TLabel",       background=BG, foreground=FG)
        s.configure("TButton",      background=ACC, foreground="#1e1e2e",
                     font=("Segoe UI",10,"bold"), relief="flat", padding=6)
        s.map("TButton",
              background=[("active","#74c7ec"),("disabled",SEL)],
              foreground=[("disabled","#6c7086")])
        s.configure("TNotebook",    background=BG, tabmargins=[2,2,2,0])
        s.configure("TNotebook.Tab",background=SEL, foreground=FG,
                     padding=[7,5], font=("Segoe UI",9))
        s.map("TNotebook.Tab",
              background=[("selected",ACC)], foreground=[("selected","#1e1e2e")])
        s.configure("Treeview",     background=ENT, foreground=FG,
                     fieldbackground=ENT, rowheight=26, font=("Segoe UI",9))
        s.configure("Treeview.Heading", background=SEL, foreground=ACC,
                     font=("Segoe UI",9,"bold"))
        s.map("Treeview",
              background=[("selected",ACC)], foreground=[("selected","#1e1e2e")])
        s.configure("TEntry",    fieldbackground=ENT, foreground=FG,
                     insertcolor=FG, selectbackground=ACC, selectforeground="#1e1e2e")
        # Combobox: явно задаём все состояния чтобы текст не сливался с фоном
        s.configure("TCombobox",
                     fieldbackground=ENT, foreground=FG, background=ENT,
                     selectbackground=ENT, selectforeground=FG,
                     insertcolor=FG, arrowcolor=ACC)
        s.map("TCombobox",
              fieldbackground=[("readonly", ENT), ("disabled", BG),
                               ("focus",    ENT), ("!focus",   ENT)],
              foreground=[      ("readonly", FG),  ("disabled", "#6c7086"),
                               ("focus",    FG),  ("!focus",   FG)],
              selectbackground=[("readonly", ENT), ("focus",    ACC),
                               ("!focus",   ENT)],
              selectforeground=[("readonly", FG),  ("focus",    "#1e1e2e"),
                               ("!focus",   FG)],
              background=[      ("readonly", ENT), ("active",   SEL),
                               ("focus",    ENT), ("!focus",   ENT)])
        s.configure("TLabelframe",  background=BG, foreground=ACC,
                     font=("Segoe UI",10,"bold"))
        s.configure("TLabelframe.Label", background=BG, foreground=ACC)
        s.configure("TCheckbutton", background=BG, foreground=FG)
        s.configure("TScrollbar",   background=SEL, troughcolor=BG, arrowcolor=FG)
        s.configure("TProgressbar", troughcolor=ENT, background=ACC)

    # ── Интерфейс ─────────────────────────────────────────────────────────────
    def _build_ui(self):
        hdr = tk.Frame(self, bg="#181825", height=52)
        hdr.pack(fill="x"); hdr.pack_propagate(False)
        tk.Label(hdr, text="⚡ Zabbix Reporter  v3.0",
                 bg="#181825", fg="#89b4fa",
                 font=("Segoe UI",17,"bold")).pack(side="left", padx=18, pady=10)
        self._lbl_ver  = tk.Label(hdr, text="", bg="#181825",
                                   fg="#a6e3a1", font=("Segoe UI",9))
        self._lbl_ver.pack(side="left", padx=8)
        self._lbl_conn = tk.Label(hdr, text="● Не подключено",
                                   bg="#181825", fg="#f38ba8",
                                   font=("Segoe UI",10))
        self._lbl_conn.pack(side="right", padx=18)

        self._prog = ttk.Progressbar(self, mode="indeterminate")
        self._prog.pack(fill="x")

        body = tk.Frame(self, bg="#1e1e2e")
        body.pack(fill="both", expand=True, padx=10, pady=(2,4))

        # Прокручиваемая левая панель
        left_outer = tk.Frame(body, bg="#1e1e2e")
        left_outer.pack(side="left", fill="y", padx=(0,8), pady=4)

        left_label = tk.Label(left_outer, text=" Подключение & Настройки ",
                               bg="#2a2a3e", fg="#89b4fa",
                               font=("Segoe UI",10,"bold"), anchor="w", pady=5)
        left_label.pack(fill="x")

        left_canvas = tk.Canvas(left_outer, bg="#1e1e2e", highlightthickness=0,
                                 width=220)
        left_scroll = ttk.Scrollbar(left_outer, orient="vertical",
                                     command=left_canvas.yview)
        left_canvas.configure(yscrollcommand=left_scroll.set)
        left_canvas.pack(side="left", fill="both", expand=True)
        left_scroll.pack(side="right", fill="y")

        left = tk.Frame(left_canvas, bg="#1e1e2e", padx=8, pady=4)
        left_win = left_canvas.create_window((0, 0), window=left, anchor="nw")

        def _on_left_configure(e):
            left_canvas.configure(scrollregion=left_canvas.bbox("all"))
        def _on_canvas_resize(e):
            left_canvas.itemconfig(left_win, width=e.width)
        left.bind("<Configure>", _on_left_configure)
        left_canvas.bind("<Configure>", _on_canvas_resize)

        # Прокрутка колесом мыши
        def _on_mousewheel(e):
            left_canvas.yview_scroll(int(-1*(e.delta/120)), "units")
        left_canvas.bind_all("<MouseWheel>", _on_mousewheel)

        self._build_left(left)

        right = tk.Frame(body, bg="#1e1e2e")
        right.pack(side="left", fill="both", expand=True, pady=4)
        self._build_tabs(right)

        # ── Строка состояния + кнопка открытия лога ─────────────────────────
        sb = tk.Frame(self, bg="#181825", height=28)
        sb.pack(fill="x"); sb.pack_propagate(False)
        self._lbl_info = tk.Label(sb, text="Готово к работе",
                                   bg="#181825", fg="#a6adc8",
                                   font=("Segoe UI",9), anchor="w")
        self._lbl_info.pack(side="left", padx=12, pady=4)
        self._btn_log_open = tk.Button(
            sb, text="📋 Лог / Отладка", bg="#181825", fg="#89b4fa",
            relief="flat", font=("Segoe UI",9,"bold"),
            cursor="hand2", command=self._open_log_window)
        self._btn_log_open.pack(side="right", padx=12, pady=3)

        # ── Буфер лога (offscreen Text — хранилище всех записей) ─────────────
        self._log_buf = tk.Text(self)   # не pack'ается — только хранит данные
        for tag, color in LogWindow.TAG_COLORS.items():
            self._log_buf.tag_configure(tag, foreground=color)

    # ── Левая панель — коллапсируемые группы ────────────────────────────────
    def _build_left(self, p):
        BG  = "#1e1e2e"
        ACC = "#89b4fa"
        FG  = "#cdd6f4"

        # Используем grid на p — каждая группа занимает одну строку.
        # Это гарантирует, что body разворачивается МЕЖДУ заголовками,
        # а не в конец (проблема pack_forget/pack).
        p.columnconfigure(0, weight=1)
        _row = [0]

        def _next_row():
            r = _row[0]; _row[0] += 1; return r

        def make_group(title, icon, expanded=True):
            """
            Создаёт коллапсируемую секцию в p.
            Заголовок и тело занимают соседние grid-строки.
            При toggle меняется только rowspan/visibility тела — позиция не плывёт.
            """
            hdr_row  = _next_row()
            body_row = _next_row()
            state    = [expanded]

            # ── Шапка секции ─────────────────────────────────────────────────
            hdr = tk.Frame(p, bg="#252535", cursor="hand2")
            hdr.grid(row=hdr_row, column=0, sticky="ew", pady=(2, 0))

            arrow = tk.Label(hdr, text="▼" if expanded else "▶",
                              bg="#252535", fg=ACC, font=("Segoe UI", 8))
            arrow.pack(side="left", padx=(8, 2), pady=3)
            tk.Label(hdr, text=f"{icon}  {title}",
                     bg="#252535", fg=ACC,
                     font=("Segoe UI", 9, "bold")).pack(side="left", pady=3)

            # ── Тело секции ───────────────────────────────────────────────────
            body = tk.Frame(p, bg=BG, padx=10, pady=6)
            if expanded:
                body.grid(row=body_row, column=0, sticky="ew")
            else:
                # занимаем строку нулевым спейсером чтобы grid не схлопывал
                tk.Frame(p, bg=BG, height=0).grid(row=body_row, column=0)

            def _toggle(e=None):
                if state[0]:              # сворачиваем
                    body.grid_remove()    # скрывает, но СОХРАНЯЕТ grid-позицию
                    arrow.configure(text="▶")
                    state[0] = False
                else:                     # разворачиваем
                    body.grid()           # восстанавливает на ту же grid-строку
                    arrow.configure(text="▼")
                    state[0] = True

            hdr.bind("<Button-1>", _toggle)
            for w in hdr.winfo_children():
                w.bind("<Button-1>", _toggle)
            return body

        def lbl(parent, t):
            tk.Label(parent, text=t, bg=BG, fg=FG,
                     font=("Segoe UI", 9)).pack(anchor="w", pady=(4, 1))

        def ent(parent, show=None):
            kw = {"width": 24}
            if show: kw["show"] = show
            e = ttk.Entry(parent, **kw)
            e.pack(fill="x", pady=(0, 2))
            return e

        def btn(parent, text, cmd, state="normal"):
            b = ttk.Button(parent, text=text, command=cmd, state=state)
            b.pack(fill="x", pady=2)
            return b

        # ── Группа: Подключение ───────────────────────────────────────────────
        g1 = make_group("Подключение", "🔌", expanded=True)

        lbl(g1, "URL сервера:")
        self._e_url = ent(g1)
        self._e_url.insert(0, "https://zabbix.example.com")
        lbl(g1, "Пользователь:")
        self._e_user = ent(g1)
        self._e_user.insert(0, "Admin")
        lbl(g1, "Пароль:")
        self._e_pass = ent(g1, show="●")

        self._v_ssl = tk.BooleanVar(value=True)
        tk.Checkbutton(g1, text="Проверять SSL", variable=self._v_ssl,
                       bg=BG, fg=FG, selectcolor="#313244",
                       activebackground=BG, activeforeground=ACC,
                       font=("Segoe UI", 9)).pack(anchor="w", pady=(4, 2))

        self._btn_con = btn(g1, "🔌 Подключиться",     self._connect)
        self._btn_dis = btn(g1, "⏏ Отключиться",       self._disconnect, "disabled")
        btn(g1, "💾 Сохранить настройки", self._cfg_save_ui)

        # ── Группа: Фильтры ───────────────────────────────────────────────────
        g2 = make_group("Фильтры", "⚙", expanded=True)

        lbl(g2, "Мин. серьёзность:")
        self._cmb_sev = ttk.Combobox(g2, width=22, state="readonly",
            values=["0 – Not classified", "1 – Information", "2 – Warning",
                    "3 – Average", "4 – High", "5 – Disaster"])
        self._cmb_sev.current(0)
        self._cmb_sev.pack(fill="x", pady=(0, 2))

        lbl(g2, "Лимит записей:")
        self._v_limit = tk.StringVar(value="500")
        ttk.Entry(g2, textvariable=self._v_limit, width=10).pack(anchor="w")

        now = datetime.datetime.now()
        lbl(g2, "С (YYYY-MM-DD):")
        self._e_from = ttk.Entry(g2, width=14)
        self._e_from.insert(0, (now - datetime.timedelta(days=7)).strftime("%Y-%m-%d"))
        self._e_from.pack(fill="x")
        lbl(g2, "По (YYYY-MM-DD):")
        self._e_till = ttk.Entry(g2, width=14)
        self._e_till.insert(0, now.strftime("%Y-%m-%d"))
        self._e_till.pack(fill="x")

        btn(g2, "🔄 Обновить все данные", self._refresh_all)
        btn(g2, "🩺 Диагностика API",     self._run_diagnostics)

        # ── Группа: Экспорт ───────────────────────────────────────────────────
        g3 = make_group("Экспорт", "💾", expanded=True)

        if PDF_OK:
            btn(g3, "📄 Экспорт PDF",   self._export_pdf)
        else:
            tk.Label(g3, text="PDF: pip install reportlab",
                     bg=BG, fg="#f38ba8", font=("Segoe UI", 8)).pack(anchor="w")
        if EXCEL_OK:
            btn(g3, "📊 Экспорт Excel", self._export_excel)
        else:
            tk.Label(g3, text="Excel: pip install openpyxl",
                     bg=BG, fg="#f38ba8", font=("Segoe UI", 8)).pack(anchor="w")
        btn(g3, "📋 Экспорт CSV", self._export_csv)


    # ── Вкладка «Объекты мониторинга» ────────────────────────────────────────
    def _build_objects_tab(self, parent):
        BG = "#1e1e2e"; BG2 = "#313244"; FG = "#cdd6f4"; ACC = "#89b4fa"

        # ── Панель фильтров (верх) ────────────────────────────────────────────
        flt = tk.Frame(parent, bg=BG2, pady=6)
        flt.pack(fill="x", padx=4, pady=(4,0))

        tk.Label(flt, text="Фильтр:", bg=BG2, fg=ACC,
                 font=("Segoe UI",9,"bold")).grid(row=0,column=0,padx=(10,4),sticky="w")

        # По группе хостов
        tk.Label(flt, text="Группа:", bg=BG2, fg=FG,
                 font=("Segoe UI",9)).grid(row=0,column=1,padx=(0,2),sticky="w")
        self._obj_grp_var = tk.StringVar(value="— все —")
        self._cmb_obj_grp = ttk.Combobox(flt, textvariable=self._obj_grp_var,
                                           state="readonly", width=22)
        self._cmb_obj_grp.grid(row=0,column=2,padx=(0,8))

        # По шаблону
        tk.Label(flt, text="Шаблон:", bg=BG2, fg=FG,
                 font=("Segoe UI",9)).grid(row=0,column=3,padx=(0,2),sticky="w")
        self._obj_tpl_var = tk.StringVar(value="— все —")
        self._cmb_obj_tpl = ttk.Combobox(flt, textvariable=self._obj_tpl_var,
                                           state="readonly", width=28)
        self._cmb_obj_tpl.grid(row=0,column=4,padx=(0,8))

        # Поиск по имени
        tk.Label(flt, text="🔍", bg=BG2, fg=FG).grid(row=0,column=5,padx=(0,2))
        self._obj_search_var = tk.StringVar()
        self._obj_search_var.trace_add("write", lambda *_: self._obj_apply_filter())
        ttk.Entry(flt, textvariable=self._obj_search_var,
                  width=22).grid(row=0,column=6,padx=(0,6))

        ttk.Button(flt, text="🔄 Загрузить",
                   command=self._obj_load).grid(row=0,column=7,padx=(0,10))

        self._cmb_obj_grp.bind("<<ComboboxSelected>>", lambda _: self._obj_apply_filter())
        self._cmb_obj_tpl.bind("<<ComboboxSelected>>", lambda _: self._obj_apply_filter())

        # ── Список хостов (слева) + детали (справа) ─────────────────────────
        paned = tk.PanedWindow(parent, orient="horizontal",
                                bg=BG, sashrelief="flat", sashwidth=5)
        paned.pack(fill="both", expand=True, padx=4, pady=4)

        # Левая часть — таблица хостов
        left_frm = tk.Frame(paned, bg=BG)
        paned.add(left_frm, minsize=300)

        cols_obj = ("Хост", "Видимое имя", "Группы", "Статус", "Доступность", "IP")
        self._tree_obj = ttk.Treeview(left_frm, columns=cols_obj,
                                       show="headings", selectmode="browse")
        for col, w in zip(cols_obj, (140,150,140,70,90,120)):
            self._tree_obj.heading(col, text=col,
                command=lambda c=col: self._sort(self._tree_obj, c))
            self._tree_obj.column(col, width=w, minwidth=40)
        vs_obj = ttk.Scrollbar(left_frm, orient="vertical",
                                command=self._tree_obj.yview)
        hs_obj = ttk.Scrollbar(left_frm, orient="horizontal",
                                command=self._tree_obj.xview)
        self._tree_obj.configure(yscrollcommand=vs_obj.set,
                                  xscrollcommand=hs_obj.set)
        self._tree_obj.grid(row=0,column=0,sticky="nsew")
        vs_obj.grid(row=0,column=1,sticky="ns")
        hs_obj.grid(row=1,column=0,sticky="ew")
        left_frm.rowconfigure(0,weight=1); left_frm.columnconfigure(0,weight=1)
        self._tree_obj._all = []

        # Теги цветов
        self._tree_obj.tag_configure("ok",  foreground="#a6e3a1")
        self._tree_obj.tag_configure("bad", foreground="#f38ba8")
        self._tree_obj.tag_configure("unk", foreground="#f9e2af")
        self._tree_obj.tag_configure("dis", foreground="#6c7086")

        self._tree_obj.bind("<<TreeviewSelect>>", self._obj_on_select)

        # Правая часть — карточка хоста
        right_frm = tk.Frame(paned, bg="#181825")
        paned.add(right_frm, minsize=340)

        tk.Label(right_frm, text="  Карточка хоста",
                 bg="#181825", fg=ACC,
                 font=("Segoe UI",10,"bold")).pack(fill="x", pady=(8,2))
        ttk.Separator(right_frm).pack(fill="x")

        self._obj_detail = tk.Text(
            right_frm, bg="#181825", fg=FG,
            font=("Consolas",9), relief="flat",
            state="disabled", wrap="word",
            padx=12, pady=8,
            selectbackground="#313244",
        )
        obj_sb = ttk.Scrollbar(right_frm, command=self._obj_detail.yview)
        self._obj_detail.configure(yscrollcommand=obj_sb.set)
        self._obj_detail.pack(side="left", fill="both", expand=True)
        obj_sb.pack(side="right", fill="y")

        # Теги для форматирования карточки
        self._obj_detail.tag_configure("hdr",
            foreground=ACC, font=("Segoe UI",10,"bold"))
        self._obj_detail.tag_configure("key",
            foreground="#cba6f7", font=("Consolas",9,"bold"))
        self._obj_detail.tag_configure("val",
            foreground=FG, font=("Consolas",9))
        self._obj_detail.tag_configure("ok",  foreground="#a6e3a1")
        self._obj_detail.tag_configure("bad", foreground="#f38ba8")
        self._obj_detail.tag_configure("dim", foreground="#6c7086")


    # ── Вкладка «Аутентификация» ──────────────────────────────────────────────
    def _build_auth_tab(self, parent):
        BG = "#1e1e2e"; BG2 = "#181825"; FG = "#cdd6f4"; ACC = "#89b4fa"

        # ── Toolbar ───────────────────────────────────────────────────────────
        tb = tk.Frame(parent, bg="#252535", pady=4)
        tb.pack(fill="x", padx=4, pady=(4,0))
        ttk.Button(tb, text="🔄 Загрузить",
                   command=self._auth_load).pack(side="left", padx=6)
        ttk.Button(tb, text="💾 Выгрузить в файл",
                   command=self._auth_export).pack(side="left", padx=(0,6))
        self._auth_status_lbl = tk.Label(
            tb, text="Нажмите «Загрузить» для получения данных",
            bg="#252535", fg="#6c7086", font=("Segoe UI",9))
        self._auth_status_lbl.pack(side="left", padx=8)

        # ── Paned: слева — список LDAP-серверов, справа — детали ─────────────
        paned = tk.PanedWindow(parent, orient="horizontal",
                               bg=BG, sashwidth=5, sashrelief="flat")
        paned.pack(fill="both", expand=True, padx=4, pady=4)

        # Левая часть — дерево LDAP-серверов
        left = tk.Frame(paned, bg=BG)
        paned.add(left, minsize=220)

        tk.Label(left, text="LDAP / SAML серверы",
                 bg=BG, fg=ACC,
                 font=("Segoe UI",9,"bold")).pack(anchor="w", padx=6, pady=(4,2))

        cols_ud = ("Имя", "Тип", "JIT")
        self._tree_ud = ttk.Treeview(left, columns=cols_ud,
                                      show="headings", selectmode="browse")
        for col, w in zip(cols_ud, (140, 60, 40)):
            self._tree_ud.heading(col, text=col)
            self._tree_ud.column(col, width=w, minwidth=30)
        vs_ud = ttk.Scrollbar(left, orient="vertical", command=self._tree_ud.yview)
        self._tree_ud.configure(yscrollcommand=vs_ud.set)
        self._tree_ud.pack(side="left", fill="both", expand=True)
        vs_ud.pack(side="right", fill="y")
        self._tree_ud.bind("<<TreeviewSelect>>", self._auth_on_select)

        # Правая часть — детальный просмотр
        right = tk.Frame(paned, bg=BG2)
        paned.add(right, minsize=420)

        tk.Label(right, text="  Детали конфигурации",
                 bg=BG2, fg=ACC,
                 font=("Segoe UI",10,"bold")).pack(fill="x", pady=(8,2))
        ttk.Separator(right).pack(fill="x")

        self._auth_detail = tk.Text(
            right, bg=BG2, fg=FG, font=("Consolas",9),
            relief="flat", state="disabled", wrap="word",
            padx=12, pady=8, selectbackground="#313244")
        auth_sb = ttk.Scrollbar(right, command=self._auth_detail.yview)
        self._auth_detail.configure(yscrollcommand=auth_sb.set)
        self._auth_detail.pack(side="left", fill="both", expand=True)
        auth_sb.pack(side="right", fill="y")

        # Стили текста карточки
        self._auth_detail.tag_configure("hdr",
            foreground=ACC, font=("Segoe UI",10,"bold"))
        self._auth_detail.tag_configure("subhdr",
            foreground="#cba6f7", font=("Segoe UI",9,"bold"))
        self._auth_detail.tag_configure("key",
            foreground="#cba6f7", font=("Consolas",9,"bold"))
        self._auth_detail.tag_configure("val",
            foreground=FG,  font=("Consolas",9))
        self._auth_detail.tag_configure("ok",  foreground="#a6e3a1")
        self._auth_detail.tag_configure("bad", foreground="#f38ba8")
        self._auth_detail.tag_configure("dim", foreground="#6c7086")
        self._auth_detail.tag_configure("warn",foreground="#f9e2af")

    # ── Загрузка данных аутентификации ────────────────────────────────────────

    def _auth_export(self):
        """Экспортирует настройки аутентификации одновременно в PDF и JSON."""
        if not self._auth_data:
            messagebox.showwarning("Нет данных",
                "Сначала нажмите «Загрузить»"); return

        folder = filedialog.askdirectory(
            title="Папка для сохранения файлов (PDF + JSON)")
        if not folder: return

        ts   = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        base = os.path.join(folder, f"zabbix_auth_{ts}")
        self._busy(True)
        self._auth_status_lbl.configure(text="Экспорт…")

        def _w():
            errors = []
            saved  = []

            # ── JSON ──────────────────────────────────────────────────────
            try:
                payload = {
                    "exported_at":   datetime.datetime.now().isoformat(),
                    "zabbix_version": getattr(self.zapi, "_api_version_str", ""),
                    "authentication": self._auth_data,
                    "user_directories": self._user_dirs,
                }
                json_path = base + ".json"
                with open(json_path, "w", encoding="utf-8") as f:
                    json.dump(payload, f, ensure_ascii=False, indent=2)
                saved.append(os.path.basename(json_path))
            except Exception as ex:
                errors.append(f"JSON: {ex}")

            # ── PDF ───────────────────────────────────────────────────────
            if PDF_OK:
                try:
                    pdf_path = base + ".pdf"
                    _auth_to_pdf(pdf_path, self._auth_data, self._user_dirs)
                    saved.append(os.path.basename(pdf_path))
                except Exception as ex:
                    errors.append(f"PDF: {ex}")
            else:
                errors.append("PDF: библиотека reportlab не установлена")

            def _done():
                self._busy(False)
                if errors:
                    msg = ("Частичный успех\n\nСохранено: " +
                           ", ".join(saved) + "\n\nОшибки:\n" +
                           "\n".join(errors))
                    self._auth_status_lbl.configure(text="Экспорт завершён с ошибками")
                    messagebox.showwarning("Экспорт", msg)
                else:
                    self._auth_status_lbl.configure(
                        text=f"Сохранено: {', '.join(saved)}")
                    messagebox.showinfo("Готово",
                        f"Файлы сохранены в:\n{folder}\n\n" +
                        "\n".join(saved))
            self.after(0, _done)

        threading.Thread(target=_w, daemon=True).start()

    def _auth_load(self):
        if not self.zapi:
            messagebox.showwarning("", "Сначала подключитесь"); return
        self._busy(True)
        self._auth_status_lbl.configure(text="Загрузка...")
        def _w():
            try:
                auth = self.zapi.get_authentication()
                dirs = self.zapi.get_userdirectories()
                self.after(0, lambda: self._auth_populate(auth, dirs))
            except Exception as ex:
                msg = str(ex)
                self.after(0, lambda m=msg: (
                    self._busy(False),
                    self._auth_status_lbl.configure(text=f"Ошибка: {m}"),
                    self.log(f"  [Auth] Ошибка: {m}", "err")))
        threading.Thread(target=_w, daemon=True).start()

    def _auth_populate(self, auth, dirs):
        self._auth_data  = auth
        self._user_dirs  = dirs
        # Заполнить дерево
        t = self._tree_ud
        t.delete(*t.get_children())
        IDP = {"1": "LDAP", "2": "SAML"}
        for d in dirs:
            idp  = IDP.get(str(d.get("idp_type","1")), "?")
            jit  = "Да" if d.get("provision_status","0") == "1" else "—"
            name = d.get("name","") or f"[{idp}]"
            t.insert("","end", values=(name, idp, jit), iid=d["userdirectoryid"])
        # Показать глобальные настройки сразу
        self._auth_show_global(auth, dirs)
        self._busy(False)
        ldap_cnt = sum(1 for d in dirs if d.get("idp_type","1") == "1")
        saml_cnt = sum(1 for d in dirs if d.get("idp_type","1") == "2")
        self._auth_status_lbl.configure(
            text=f"Загружено: {ldap_cnt} LDAP, {saml_cnt} SAML серверов")
        self.log(f"  [Auth] LDAP={ldap_cnt} SAML={saml_cnt}", "ok")

    def _auth_on_select(self, event=None):
        sel = self._tree_ud.selection()
        if not sel: return
        ud_id = sel[0]
        d = next((x for x in self._user_dirs
                   if x["userdirectoryid"] == ud_id), None)
        if d:
            self._auth_show_directory(d)

    def _auth_show_global(self, auth, dirs):
        """Показывает глобальные настройки аутентификации."""
        txt = self._auth_detail
        txt.configure(state="normal"); txt.delete("1.0","end")

        AUTH_TYPE = {"0": "Internal", "1": "LDAP", "2": "HTTP"}
        YESNO = lambda v: ("Да" if str(v)=="1" else "Нет")

        def sec(title): txt.insert("end", f"\n● {title}\n", "hdr")
        def row(key, val, tag="val"):
            txt.insert("end", f"  {key:<32}", "key")
            txt.insert("end", f"{val}\n", tag)

        sec("Глобальные настройки аутентификации")
        at = auth.get("authentication_type","0")
        row("Default authentication",
            AUTH_TYPE.get(at, at),
            "ok" if at=="0" else "warn")
        row("JIT provision interval",    auth.get("jit_provision_interval","—"))
        row("LDAP JIT provisioning",     YESNO(auth.get("ldap_jit_status","0")))
        row("LDAP case sensitive",       YESNO(auth.get("ldap_case_sensitive","1")))

        sec("Парольная политика (Internal)")
        row("Min password length",       auth.get("passwd_min_length","—"))
        rules = int(auth.get("passwd_check_rules","0"))
        rule_labels = [
            (1,  "Содержит заглавные буквы"),
            (2,  "Содержит строчные буквы"),
            (4,  "Содержит цифры"),
            (8,  "Содержит спецсимволы"),
        ]
        for bit, label in rule_labels:
            row(f"  {label}", "Да" if (rules & bit) else "Нет",
                "ok" if (rules & bit) else "dim")

        sec("HTTP аутентификация")
        row("HTTP auth enabled",  YESNO(auth.get("http_auth_enabled","0")))
        row("HTTP case sensitive",YESNO(auth.get("http_case_sensitive","1")))
        row("HTTP strip domains", auth.get("http_strip_domains","") or "—")

        sec("SAML")
        row("SAML auth enabled",  YESNO(auth.get("saml_auth_enabled","0")))
        row("SAML JIT status",    YESNO(auth.get("saml_jit_status","0")))
        row("SAML case sensitive",YESNO(auth.get("saml_case_sensitive","0")))

        sec("MFA")
        row("MFA enabled", YESNO(auth.get("mfa_status","0")))

        if dirs:
            sec("LDAP / SAML серверы")
            txt.insert("end",
                "  Выберите сервер в списке слева для просмотра деталей\n","dim")

        txt.configure(state="disabled")

    def _auth_show_directory(self, d):
        """Полная карточка одного LDAP/SAML user directory."""
        txt = self._auth_detail
        txt.configure(state="normal"); txt.delete("1.0","end")

        IDP  = {"1": "LDAP", "2": "SAML"}
        YESNO = lambda v: ("Да" if str(v) in ("1","true","True") else "Нет")
        idp_type = str(d.get("idp_type","1"))

        def sec(title):
            txt.insert("end", f"\n● {title}\n", "hdr")
        def subsec(title):
            txt.insert("end", f"\n  ▸ {title}\n", "subhdr")
        def row(key, val, tag="val"):
            txt.insert("end", f"  {key:<34}", "key")
            txt.insert("end", f"{val}\n", tag)

        # ── Основное ──────────────────────────────────────────────────────────
        sec(f"{IDP.get(idp_type,'?')} сервер: {d.get('name','') or '(без имени)'}")
        row("ID",              d.get("userdirectoryid",""))
        row("Описание",        d.get("description","") or "—")
        row("JIT Provisioning",
            "Включён" if d.get("provision_status","0")=="1" else "Отключён",
            "ok" if d.get("provision_status","0")=="1" else "dim")

        if idp_type == "1":   # ── LDAP ───────────────────────────────────────
            sec("Подключение к LDAP серверу")
            row("Host",          d.get("host",""))
            row("Port",          d.get("port","389"))
            row("Base DN",       d.get("base_dn",""))
            row("Bind DN",       d.get("bind_dn","") or "(anonymous)")
            row("Bind password", "***" if d.get("bind_password","") else "(не задан)")
            row("StartTLS",      YESNO(d.get("start_tls","0")))

            sec("Поиск пользователей")
            row("Search attribute",  d.get("search_attribute",""))

            if d.get("provision_status","0") == "1":
                sec("JIT Provisioning — настройки")
                row("Group configuration",
                    {"1":"memberOf","2":"groupOfNames"}.get(
                        str(d.get("group_configuration","1")), "?"))
                row("Group base DN",        d.get("group_base_dn","") or "—")
                row("Group name attr",      d.get("group_name","") or "—")
                row("Group member attr",    d.get("group_member","") or "—")
                row("User username attr",   d.get("user_username","") or "—")
                row("User lastname attr",   d.get("user_lastname","") or "—")
                row("User ref attr",        d.get("user_ref_attr","") or "—")
                row("Group filter",         d.get("group_filter","") or "—")

                # ── Group mappings ─────────────────────────────────────────────
                subsec("User Group Mapping (provision_groups)")
                pgs = d.get("provision_groups", [])
                if pgs:
                    for i, pg in enumerate(pgs, 1):
                        txt.insert("end", f"\n    [{i}] LDAP pattern: ", "key")
                        txt.insert("end", f"{pg.get('name','*')}\n", "val")
                        role_name = pg.get("_role_name","") or pg.get("roleid","?")
                        txt.insert("end", f"        Role:          ", "key")
                        txt.insert("end", f"{role_name}\n", "val")
                        ugs = pg.get("user_groups", [])
                        if ugs:
                            txt.insert("end", f"        Zabbix groups: ", "key")
                            grp_names = [ug.get("_grp_name","") or ug.get("usrgrpid","?")
                                         for ug in ugs]
                            txt.insert("end", f"{', '.join(grp_names)}\n", "val")
                else:
                    txt.insert("end", "    (нет маппингов)\n", "dim")

                # ── Media type mappings ────────────────────────────────────────
                subsec("Media Type Mapping (provision_media)")
                pms = d.get("provision_media", [])
                if pms:
                    for i, pm in enumerate(pms, 1):
                        mt_name = pm.get("_mt_name","") or pm.get("mediatypeid","?")
                        txt.insert("end", f"\n    [{i}] {pm.get('name','')}\n","key")
                        row("        Media type", mt_name)
                        row("        Attribute",  pm.get("attribute",""))
                        row("        Active",     YESNO(pm.get("active","1")))
                        row("        Severity",   pm.get("severity","63"))
                        row("        Period",     pm.get("period","1-7,00:00-24:00"))
                else:
                    txt.insert("end", "    (нет маппингов)\n", "dim")

        elif idp_type == "2":  # ── SAML ────────────────────────────────────────
            sec("SAML настройки")
            row("IDP Entity ID",   d.get("idp_entityid",""))
            row("SSO URL",         d.get("sso_url",""))
            row("SLO URL",         d.get("slo_url","") or "—")
            row("Username attr",   d.get("username_attribute",""))
            row("SP Entity ID",    d.get("sp_entityid",""))
            row("NameID format",   d.get("nameid_format","") or "—")
            row("SCIM enabled",    YESNO(d.get("scim_status","0")))

            sec("Подписи и шифрование")
            row("Sign messages",         YESNO(d.get("sign_messages","0")))
            row("Sign assertions",       YESNO(d.get("sign_assertions","0")))
            row("Sign authn requests",   YESNO(d.get("sign_authn_requests","0")))
            row("Sign logout requests",  YESNO(d.get("sign_logout_requests","0")))
            row("Sign logout responses", YESNO(d.get("sign_logout_responses","0")))
            row("Encrypt NameID",        YESNO(d.get("encrypt_nameid","0")))
            row("Encrypt assertions",    YESNO(d.get("encrypt_assertions","0")))

            if d.get("provision_status","0") == "1":
                sec("JIT Provisioning")
                row("Group name attr",  d.get("group_name","") or "—")
                row("User username attr",d.get("user_username","") or "—")
                row("User lastname attr",d.get("user_lastname","") or "—")

                subsec("User Group Mapping")
                pgs = d.get("provision_groups", [])
                if pgs:
                    for i, pg in enumerate(pgs, 1):
                        txt.insert("end",f"\n    [{i}] Pattern: ","key")
                        txt.insert("end",f"{pg.get('name','*')}\n","val")
                        role_name = pg.get("_role_name","") or pg.get("roleid","?")
                        txt.insert("end",f"        Role:     ","key")
                        txt.insert("end",f"{role_name}\n","val")
                        ugs = pg.get("user_groups", [])
                        if ugs:
                            txt.insert("end",f"        Groups:   ","key")
                            grp_names = [ug.get("_grp_name","") or ug.get("usrgrpid","?")
                                         for ug in ugs]
                            txt.insert("end",f"{', '.join(grp_names)}\n","val")
                else:
                    txt.insert("end","    (нет маппингов)\n","dim")

                subsec("Media Type Mapping")
                pms = d.get("provision_media", [])
                if pms:
                    for i, pm in enumerate(pms, 1):
                        mt_name = pm.get("_mt_name","") or pm.get("mediatypeid","?")
                        txt.insert("end",f"\n    [{i}] {pm.get('name','')}\n","key")
                        row("        Media type", mt_name)
                        row("        Attribute",  pm.get("attribute",""))
                        row("        Active",     YESNO(pm.get("active","1")))
                else:
                    txt.insert("end","    (нет маппингов)\n","dim")

        txt.configure(state="disabled")


    # ════════════════════════════════════════════════════════════════════════
    #  Вкладка «Сравнение шаблонов»
    # ════════════════════════════════════════════════════════════════════════
    def _build_compare_tab(self, parent):
        BG = "#1e1e2e"; BG2 = "#181825"; FG = "#cdd6f4"; ACC = "#89b4fa"

        # ── Toolbar: выбор двух шаблонов ─────────────────────────────────────
        tb = tk.Frame(parent, bg="#252535", pady=6)
        tb.pack(fill="x", padx=4, pady=(4, 0))

        tk.Label(tb, text="Шаблон A:", bg="#252535", fg=FG,
                 font=("Segoe UI",9)).grid(row=0, column=0, padx=(10,4))
        self._cmp_a_var = tk.StringVar()
        self._cmb_cmp_a = ttk.Combobox(tb, textvariable=self._cmp_a_var,
                                        state="readonly", width=36)
        self._cmb_cmp_a.grid(row=0, column=1, padx=(0,12))

        tk.Label(tb, text="Шаблон B:", bg="#252535", fg=FG,
                 font=("Segoe UI",9)).grid(row=0, column=2, padx=(0,4))
        self._cmp_b_var = tk.StringVar()
        self._cmb_cmp_b = ttk.Combobox(tb, textvariable=self._cmp_b_var,
                                        state="readonly", width=36)
        self._cmb_cmp_b.grid(row=0, column=3, padx=(0,12))

        ttk.Button(tb, text="🔄 Загрузить шаблоны",
                   command=self._tpl_exp_load).grid(row=0, column=4, padx=(0,8))
        ttk.Button(tb, text="⚖ Сравнить",
                   command=self._cmp_run).grid(row=0, column=5, padx=(0,6))
        ttk.Button(tb, text="📄 PDF",
                   command=self._cmp_export_pdf).grid(row=0, column=6, padx=(0,4))
        ttk.Button(tb, text="📋 CSV",
                   command=self._cmp_export_csv).grid(row=0, column=7, padx=(0,10))

        self._cmp_status = tk.Label(tb, text="Нажмите «Загрузить шаблоны», затем выберите два и нажмите «Сравнить»",
                                     bg="#252535", fg="#6c7086", font=("Segoe UI",8))
        self._cmp_status.grid(row=1, column=0, columnspan=8, padx=10, pady=(2,0), sticky="w")

        # ── Notebook с разделами результатов ─────────────────────────────────
        self._cmp_nb = ttk.Notebook(parent)
        self._cmp_nb.pack(fill="both", expand=True, padx=4, pady=4)

        # Создаём вкладки результатов сравнения
        self._cmp_trees = {}
        sections = [
            ("items",       "📊 Items"),
            ("triggers",    "⚡ Triggers"),
            ("drules",      "🔍 Discovery Rules"),
            ("item_protos", "📊 Item Prototypes"),
            ("trig_protos", "⚡ Trigger Prototypes"),
        ]
        for key, label in sections:
            frm = ttk.Frame(self._cmp_nb)
            self._cmp_nb.add(frm, text=f" {label} ")
            cols = ("Статус", "Имя A", "Имя B", "Key A", "Key B", "Различия")
            t = ttk.Treeview(frm, columns=cols, show="headings",
                             selectmode="browse")
            widths = (90, 200, 200, 200, 200, 220)
            for col, w in zip(cols, widths):
                t.heading(col, text=col)
                t.column(col, width=w, minwidth=40)
            vs = ttk.Scrollbar(frm, orient="vertical",   command=t.yview)
            hs = ttk.Scrollbar(frm, orient="horizontal", command=t.xview)
            t.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
            t.grid(row=0, column=0, sticky="nsew")
            vs.grid(row=0, column=1, sticky="ns")
            hs.grid(row=1, column=0, sticky="ew")
            frm.rowconfigure(0, weight=1); frm.columnconfigure(0, weight=1)
            # Теги цветов
            t.tag_configure("only_a",   foreground="#f38ba8")   # красный — только в A
            t.tag_configure("only_b",   foreground="#a6e3a1")   # зелёный — только в B
            t.tag_configure("differ",   foreground="#f9e2af")   # жёлтый  — различия
            t.tag_configure("similar",  foreground="#fab387")   # оранжевый — похожие
            t.tag_configure("equal",    foreground="#6c7086")   # серый   — совпадают
            self._cmp_trees[key] = t

        self._cmp_result = {}   # последний результат для экспорта

    # ── Логика сравнения ──────────────────────────────────────────────────────
    def _cmp_run(self):
        if not self.zapi:
            messagebox.showwarning("", "Подключитесь к Zabbix"); return
        a_sel = self._cmp_a_var.get()
        b_sel = self._cmp_b_var.get()
        if not a_sel or not b_sel:
            messagebox.showwarning("", "Выберите оба шаблона"); return
        if a_sel == b_sel:
            messagebox.showwarning("", "Выберите разные шаблоны"); return
        a_id = a_sel.split("[")[-1].rstrip("]")
        b_id = b_sel.split("[")[-1].rstrip("]")
        self._busy(True)
        self._cmp_status.configure(text="Загрузка данных шаблонов...")

        def _w():
            try:
                data_a = self.zapi.get_template_full(a_id)
                data_b = self.zapi.get_template_full(b_id)
                result = _compare_templates(data_a, data_b)
                self.after(0, lambda: self._cmp_populate(result, a_sel, b_sel))
            except Exception as ex:
                msg = str(ex)
                self.after(0, lambda m=msg: (
                    self._busy(False),
                    self._cmp_status.configure(text=f"Ошибка: {m}"),
                    self.log(f"  [Compare] Ошибка: {m}", "err")))
        threading.Thread(target=_w, daemon=True).start()

    def _cmp_populate(self, result, a_sel, b_sel):
        self._cmp_result = result
        self._cmp_result["_a_name"] = a_sel.split(" [")[0]
        self._cmp_result["_b_name"] = b_sel.split(" [")[0]

        SECTION_KEYS = {
            "items":       "items",
            "triggers":    "triggers",
            "drules":      "drules",
            "item_protos": "item_protos",
            "trig_protos": "trig_protos",
        }
        totals = {}
        for key, tree_key in SECTION_KEYS.items():
            t = self._cmp_trees[tree_key]
            t.delete(*t.get_children())
            rows = result.get(key, [])
            cnt = {"only_a": 0, "only_b": 0, "differ": 0,
                   "similar": 0, "equal": 0}
            for row in rows:
                status = row[0]
                tag = {
                    "Только A":   "only_a",
                    "Только B":   "only_b",
                    "Различия":   "differ",
                    "Похожие":    "similar",
                    "Совпадает":  "equal",
                }.get(status, "equal")
                t.insert("", "end", values=row, tags=(tag,))
                cnt[tag] = cnt.get(tag, 0) + 1
            totals[key] = cnt

        # Обновить статус
        total_diff = sum(
            v["only_a"] + v["only_b"] + v["differ"] + v["similar"]
            for v in totals.values())
        self._cmp_status.configure(
            text=(f"A: {self._cmp_result['_a_name']}  |  "
                  f"B: {self._cmp_result['_b_name']}  |  "
                  f"Отличий: {total_diff}"))
        self._busy(False)
        self.log(f"  [Compare] Сравнение завершено, отличий: {total_diff}", "ok")

    def _cmp_export_pdf(self):
        if not self._cmp_result:
            messagebox.showwarning("", "Сначала выполните сравнение"); return
        path = filedialog.asksaveasfilename(
            defaultextension=".pdf", filetypes=[("PDF","*.pdf")],
            initialfile=f"compare_{datetime.date.today()}.pdf")
        if not path: return
        self._busy(True)
        def _w():
            try:
                _compare_to_pdf(path, self._cmp_result)
                self.after(0, lambda: (self._busy(False),
                    self._info(f"PDF: {path}"),
                    messagebox.showinfo("Готово", f"PDF сохранён:\n{path}")))
            except Exception as ex:
                msg = str(ex)
                self.after(0, lambda m=msg: (self._busy(False),
                    messagebox.showerror("Ошибка", m)))
        threading.Thread(target=_w, daemon=True).start()

    def _cmp_export_csv(self):
        if not self._cmp_result:
            messagebox.showwarning("", "Сначала выполните сравнение"); return
        folder = filedialog.askdirectory(title="Папка для CSV файлов")
        if not folder: return
        self._busy(True)
        def _w():
            try:
                _compare_to_csv(folder, self._cmp_result)
                self.after(0, lambda: (self._busy(False),
                    self._info(f"CSV: {folder}"),
                    messagebox.showinfo("Готово", f"CSV сохранены в:\n{folder}")))
            except Exception as ex:
                msg = str(ex)
                self.after(0, lambda m=msg: (self._busy(False),
                    messagebox.showerror("Ошибка", m)))
        threading.Thread(target=_w, daemon=True).start()

    def _cmp_update_template_lists(self, templates):
        """Обновить только выпадающие списки сравнения (не трогать дерево экспорта)."""""
        vals = [f"{t.get('name','')} [{t.get('templateid','')}]" for t in templates]
        self._cmb_cmp_a["values"] = vals
        self._cmb_cmp_b["values"] = vals

    # ════════════════════════════════════════════════════════════════════════
    #  Вкладка «Экспорт шаблонов»
    # ════════════════════════════════════════════════════════════════════════
    def _build_tpl_export_tab(self, parent):
        BG = "#1e1e2e"; BG2 = "#181825"; FG = "#cdd6f4"; ACC = "#89b4fa"

        # Toolbar
        tb = tk.Frame(parent, bg="#252535", pady=6)
        tb.pack(fill="x", padx=4, pady=(4, 0))

        ttk.Button(tb, text="🔄 Загрузить список",
                   command=self._tpl_exp_load).pack(side="left", padx=6)
        ttk.Button(tb, text="☑ Выбрать все",
                   command=lambda: self._tpl_exp_select_all(True)).pack(side="left", padx=2)
        ttk.Button(tb, text="☐ Снять все",
                   command=lambda: self._tpl_exp_select_all(False)).pack(side="left", padx=2)

        tk.Label(tb, text="Формат:", bg="#252535", fg=FG,
                 font=("Segoe UI",9)).pack(side="left", padx=(12,4))
        self._tpl_exp_fmt = tk.StringVar(value="yaml")
        for fmt in ("yaml", "xml", "json"):
            ttk.Radiobutton(tb, text=fmt.upper(),
                           variable=self._tpl_exp_fmt, value=fmt).pack(side="left", padx=2)

        ttk.Button(tb, text="📤 Экспортировать выбранные",
                   command=self._tpl_exp_run).pack(side="left", padx=(16,4))

        self._tpl_exp_status = tk.Label(
            tb, text="Нажмите «Загрузить список» или перейдите в «Объекты»",
            bg="#252535", fg="#6c7086", font=("Segoe UI",8))
        self._tpl_exp_status.pack(side="left", padx=8)

        # Список шаблонов с чекбоксами
        list_frm = tk.Frame(parent, bg=BG)
        list_frm.pack(fill="both", expand=True, padx=4, pady=4)

        cols_exp = ("✓", "Имя шаблона", "Группы", "ID")
        self._tree_tpl_exp = ttk.Treeview(list_frm, columns=cols_exp,
                                           show="headings", selectmode="browse")
        for col, w in zip(cols_exp, (30, 340, 260, 80)):
            self._tree_tpl_exp.heading(col, text=col)
            self._tree_tpl_exp.column(col, width=w, minwidth=20)
        vs_te = ttk.Scrollbar(list_frm, orient="vertical",
                               command=self._tree_tpl_exp.yview)
        hs_te = ttk.Scrollbar(list_frm, orient="horizontal",
                               command=self._tree_tpl_exp.xview)
        self._tree_tpl_exp.configure(yscrollcommand=vs_te.set,
                                      xscrollcommand=hs_te.set)
        self._tree_tpl_exp.grid(row=0, column=0, sticky="nsew")
        vs_te.grid(row=0, column=1, sticky="ns")
        hs_te.grid(row=1, column=0, sticky="ew")
        list_frm.rowconfigure(0, weight=1); list_frm.columnconfigure(0, weight=1)

        # Клик по строке — переключить чекбокс
        self._tree_tpl_exp.bind("<ButtonRelease-1>", self._tpl_exp_toggle)
        self._tree_tpl_exp.bind("<space>", self._tpl_exp_toggle)
        self._tpl_exp_checked = set()   # templateid выбранных
        self._tpl_exp_all_templates = []

    def _tpl_exp_load(self):
        if not self.zapi:
            messagebox.showwarning("", "Подключитесь к Zabbix"); return
        self._busy(True)
        self._tpl_exp_status.configure(text="Загрузка...")
        self.log("  [Templates] Начало загрузки списка...", "req")
        def _w():
            try:
                tpls = self._call_or_get_templates()
                self.after(0, lambda: self._tpl_exp_refresh(tpls))
            except Exception as ex:
                msg = str(ex)
                self.after(0, lambda m=msg: (
                    self._busy(False),
                    self._tpl_exp_status.configure(text=f"Ошибка: {m}"),
                    self.log(f"  [Templates] Ошибка загрузки: {m}", "err"),
                    messagebox.showerror("Ошибка загрузки шаблонов", m)))
        threading.Thread(target=_w, daemon=True).start()

    def _call_or_get_templates(self):
        """
        Zabbix 6.2+: selectGroups -> selectTemplateGroups, ответ в templategroups.
        Пробуем три варианта с фоллбэком.
        """
        # Попытка 1: Zabbix 6.2+ selectTemplateGroups
        try:
            tpls = self.zapi._call("template.get", {
                "output":               ["templateid","name","description"],
                "selectTemplateGroups": ["groupid","name"],
                "sortfield":            "name",
            })
            for t in tpls:
                t["groups"] = t.get("templategroups", t.get("groups", []))
            self.log(f"  [Templates] Загружено: {len(tpls)}", "ok")
            return tpls
        except Exception as e1:
            self.log(f"  [Templates] selectTemplateGroups failed: {e1}", "warn")

        # Попытка 2: старый selectGroups (Zabbix < 6.2)
        try:
            tpls = self.zapi._call("template.get", {
                "output":       ["templateid","name","description"],
                "selectGroups": ["groupid","name"],
                "sortfield":    "name",
            })
            for t in tpls:
                if "groups" not in t:
                    t["groups"] = []
            self.log(f"  [Templates] Загружено (fallback): {len(tpls)}", "ok")
            return tpls
        except Exception as e2:
            self.log(f"  [Templates] selectGroups fallback failed: {e2}", "warn")

        # Попытка 3: без групп
        tpls = self.zapi._call("template.get", {
            "output":    ["templateid","name","description"],
            "sortfield": "name",
        })
        for t in tpls:
            t["groups"] = []
        self.log(f"  [Templates] Загружено (без групп): {len(tpls)}", "ok")
        return tpls

    def _tpl_exp_refresh(self, templates):
        self._tpl_exp_all_templates = templates
        self._tpl_exp_checked = set()
        # Заполнить дерево
        t = self._tree_tpl_exp
        t.delete(*t.get_children())
        for tpl in templates:
            grps = ", ".join(g.get("name","") for g in tpl.get("groups", []))
            iid  = str(tpl["templateid"])          # iid всегда строка
            t.insert("", "end",
                     values=("☐", tpl.get("name",""), grps, iid),
                     iid=iid)
        cnt = len(templates)
        status = (f"Загружено шаблонов: {cnt}"
                  if cnt else
                  "Шаблонов не найдено — проверьте права пользователя")
        self._tpl_exp_status.configure(text=status)
        self.log(f"  [Templates] {status}", "ok" if cnt else "warn")
        self._busy(False)
        # Обновить выпадающие списки вкладки сравнения (без рекурсии)
        try:
            vals = [f"{tpl.get('name','')} [{tpl.get('templateid','')}]"
                    for tpl in templates]
            self._cmb_cmp_a["values"] = vals
            self._cmb_cmp_b["values"] = vals
        except Exception:
            pass

    def _tpl_exp_toggle(self, event=None):
        sel = self._tree_tpl_exp.selection()
        if not sel: return
        iid = str(sel[0])   # всегда строка
        if iid in self._tpl_exp_checked:
            self._tpl_exp_checked.discard(iid)
            self._tree_tpl_exp.set(iid, "✓", "☐")
        else:
            self._tpl_exp_checked.add(iid)
            self._tree_tpl_exp.set(iid, "✓", "☑")
        n = len(self._tpl_exp_checked)
        self._tpl_exp_status.configure(text=f"Выбрано: {n}")

    def _tpl_exp_select_all(self, select: bool):
        # Итерируем только по строкам которые реально вставлены в дерево
        existing_iids = set(self._tree_tpl_exp.get_children())
        self._tpl_exp_checked = set()
        for iid in existing_iids:
            if select:
                self._tpl_exp_checked.add(iid)
                self._tree_tpl_exp.set(iid, "✓", "☑")
            else:
                self._tree_tpl_exp.set(iid, "✓", "☐")
        n = len(self._tpl_exp_checked)
        self._tpl_exp_status.configure(text=f"Выбрано: {n}")

    def _tpl_exp_run(self):
        if not self.zapi:
            messagebox.showwarning("", "Подключитесь к Zabbix"); return
        if not self._tpl_exp_checked:
            messagebox.showwarning("", "Выберите хотя бы один шаблон"); return
        folder = filedialog.askdirectory(title="Папка для сохранения YAML/XML/JSON файлов")
        if not folder: return
        fmt = self._tpl_exp_fmt.get()
        ids = list(self._tpl_exp_checked)   # все iid — строки
        # Найти имена выбранных шаблонов (ключи тоже строки)
        tpl_map = {str(t["templateid"]): t["name"] for t in self._tpl_exp_all_templates}
        self._busy(True)
        self._tpl_exp_status.configure(text=f"Экспорт 0/{len(ids)}...")

        def _w():
            ok = 0; errors = []
            for i, tid in enumerate(ids, 1):
                tname = tpl_map.get(tid, tid)
                # Безопасное имя файла
                safe = "".join(c if c.isalnum() or c in " _-.()" else "_"
                               for c in tname).strip()
                ext = {"yaml":".yaml","xml":".xml","json":".json"}.get(fmt,".yaml")
                fpath = os.path.join(folder, f"{safe}{ext}")
                try:
                    content = self.zapi.export_template_yaml(tid) if fmt == "yaml" \
                              else self.zapi._call("configuration.export", {
                                  "format": fmt, "options": {"templates": [tid]}})
                    with open(fpath, "w", encoding="utf-8") as f:
                        f.write(content)
                    ok += 1
                    self.after(0, lambda ii=i, n=len(ids):
                        self._tpl_exp_status.configure(
                            text=f"Экспорт {ii}/{n}..."))
                except Exception as ex:
                    errors.append(f"{tname}: {ex}")
            def _done():
                self._busy(False)
                msg = f"Экспортировано: {ok}/{len(ids)}"
                if errors:
                    msg += f"\nОшибок: {len(errors)}"
                self._tpl_exp_status.configure(text=msg)
                if errors:
                    messagebox.showwarning("Частичный успех",
                        f"{msg}\n\nОшибки:\n" + "\n".join(errors[:5]))
                else:
                    messagebox.showinfo("Готово",
                        f"Экспортировано {ok} шаблонов в:\n{folder}")
            self.after(0, _done)
        threading.Thread(target=_w, daemon=True).start()


    # ════════════════════════════════════════════════════════════════════════
    #  Вкладка «Дубликаты хостов»
    # ════════════════════════════════════════════════════════════════════════
    def _build_dupes_tab(self, parent):
        BG = "#1e1e2e"; BG2 = "#181825"; FG = "#cdd6f4"; ACC = "#89b4fa"

        # ── Toolbar ───────────────────────────────────────────────────────
        tb = tk.Frame(parent, bg="#252535", pady=6)
        tb.pack(fill="x", padx=4, pady=(4, 0))

        ttk.Button(tb, text="🔍 Найти дубликаты",
                   command=self._dupes_run).pack(side="left", padx=6)
        ttk.Button(tb, text="📄 PDF",
                   command=self._dupes_export_pdf).pack(side="left", padx=(0, 4))
        ttk.Button(tb, text="📋 CSV",
                   command=self._dupes_export_csv).pack(side="left", padx=(0, 8))

        # Фильтры типов
        tk.Label(tb, text="Показать:", bg="#252535", fg=FG,
                 font=("Segoe UI", 9)).pack(side="left", padx=(8, 4))
        self._dupes_show_hostname = tk.BooleanVar(value=True)
        self._dupes_show_ip       = tk.BooleanVar(value=True)
        for var, text in [(self._dupes_show_hostname, "По hostname"),
                          (self._dupes_show_ip,       "По IP")]:
            tk.Checkbutton(tb, text=text, variable=var,
                           bg="#252535", fg=FG, selectcolor="#313244",
                           activebackground="#252535", activeforeground=ACC,
                           font=("Segoe UI", 9),
                           command=self._dupes_apply_filter).pack(side="left", padx=2)

        self._dupes_status = tk.Label(
            tb, text="Нажмите «Найти дубликаты» (используются загруженные хосты)",
            bg="#252535", fg="#6c7086", font=("Segoe UI", 8))
        self._dupes_status.pack(side="left", padx=10)

        # ── PanedWindow: список групп (слева) + детали группы (справа) ───
        paned = tk.PanedWindow(parent, orient="horizontal",
                               bg=BG, sashwidth=5, sashrelief="flat")
        paned.pack(fill="both", expand=True, padx=4, pady=4)

        # ── Левая часть — список групп дубликатов ─────────────────────────
        left = tk.Frame(paned, bg=BG)
        paned.add(left, minsize=280)

        tk.Label(left, text="Группы дубликатов",
                 bg=BG, fg=ACC,
                 font=("Segoe UI", 9, "bold")).pack(anchor="w", padx=6, pady=(4, 2))

        cols_g = ("Тип", "Описание", "Хостов")
        self._tree_dupes_groups = ttk.Treeview(
            left, columns=cols_g, show="headings", selectmode="browse")
        for col, w in zip(cols_g, (90, 240, 60)):
            self._tree_dupes_groups.heading(col, text=col)
            self._tree_dupes_groups.column(col, width=w, minwidth=40)
        vs_g = ttk.Scrollbar(left, orient="vertical",
                              command=self._tree_dupes_groups.yview)
        self._tree_dupes_groups.configure(yscrollcommand=vs_g.set)
        self._tree_dupes_groups.pack(side="left", fill="both", expand=True)
        vs_g.pack(side="right", fill="y")

        # Цвета типов
        self._tree_dupes_groups.tag_configure("hostname", foreground="#f9e2af")
        self._tree_dupes_groups.tag_configure("ip",       foreground="#f38ba8")

        self._tree_dupes_groups.bind("<<TreeviewSelect>>",
                                      self._dupes_on_group_select)

        # ── Правая часть — детали выбранной группы ────────────────────────
        right = tk.Frame(paned, bg=BG2)
        paned.add(right, minsize=400)

        tk.Label(right, text="  Детали группы дубликатов",
                 bg=BG2, fg=ACC,
                 font=("Segoe UI", 10, "bold")).pack(fill="x", pady=(8, 2))
        ttk.Separator(right).pack(fill="x")

        # Таблица хостов в группе
        cols_d = ("Host name", "Видимое имя", "Группы", "IP / интерфейсы",
                  "Статус", "Доступн.")
        frm_tree = tk.Frame(right, bg=BG2)
        frm_tree.pack(fill="both", expand=True, padx=4, pady=4)

        self._tree_dupes_detail = ttk.Treeview(
            frm_tree, columns=cols_d, show="headings", selectmode="browse")
        for col, w in zip(cols_d, (160, 160, 160, 200, 70, 80)):
            self._tree_dupes_detail.heading(col, text=col)
            self._tree_dupes_detail.column(col, width=w, minwidth=40)
        vs_d = ttk.Scrollbar(frm_tree, orient="vertical",
                              command=self._tree_dupes_detail.yview)
        hs_d = ttk.Scrollbar(frm_tree, orient="horizontal",
                              command=self._tree_dupes_detail.xview)
        self._tree_dupes_detail.configure(yscrollcommand=vs_d.set,
                                           xscrollcommand=hs_d.set)
        self._tree_dupes_detail.grid(row=0, column=0, sticky="nsew")
        vs_d.grid(row=0, column=1, sticky="ns")
        hs_d.grid(row=1, column=0, sticky="ew")
        frm_tree.rowconfigure(0, weight=1); frm_tree.columnconfigure(0, weight=1)

        self._tree_dupes_detail.tag_configure("enabled",  foreground="#a6e3a1")
        self._tree_dupes_detail.tag_configure("disabled", foreground="#6c7086")

        # Блок с пояснением причины
        self._dupes_reason_lbl = tk.Label(
            right, text="", bg=BG2, fg="#f9e2af",
            font=("Segoe UI", 9), anchor="w", wraplength=600)
        self._dupes_reason_lbl.pack(fill="x", padx=8, pady=(0, 6))

        # Внутреннее состояние
        self._dupes_result      = {}   # {by_name, by_fqdn, by_ip}
        self._dupes_all_groups  = []   # [(type_key, group_dict), ...]
        self._dupes_filtered    = []   # отфильтрованный список

    # ── Поиск дубликатов ──────────────────────────────────────────────────
    def _dupes_run(self):
        if not self._hosts:
            messagebox.showwarning("Нет данных",
                "Сначала загрузите хосты (кнопка «Обновить все данные»)")
            return
        self._busy(True)
        self._dupes_status.configure(text="Анализ…")

        def _w():
            try:
                result = find_duplicates(self._hosts)
                self.after(0, lambda: self._dupes_populate(result))
            except Exception as ex:
                msg = str(ex)
                self.after(0, lambda m=msg: (
                    self._busy(False),
                    self._dupes_status.configure(text=f"Ошибка: {m}"),
                    self.log(f"  [Dupes] Ошибка: {m}", "err")))
        threading.Thread(target=_w, daemon=True).start()

    def _dupes_populate(self, result):
        self._dupes_result = result

        # Собрать плоский список всех групп
        all_groups = []
        for g in result.get("by_hostname", []):
            all_groups.append(("hostname", g))
        for g in result.get("by_ip", []):
            all_groups.append(("ip", g))
        self._dupes_all_groups = all_groups

        total = len(all_groups)
        total_hosts = sum(len(g["hosts"]) for _, g in all_groups)
        self._dupes_status.configure(
            text=f"Найдено групп дубликатов: {total}  (затронуто хостов: {total_hosts})")
        self.log(f"  [Dupes] Групп: {total}, хостов: {total_hosts}", "ok")

        self._dupes_apply_filter()
        self._busy(False)

    def _dupes_apply_filter(self):
        """Применить фильтры и перерисовать список групп."""
        show = {
            "hostname": self._dupes_show_hostname.get(),
            "ip":       self._dupes_show_ip.get(),
        }
        filtered = [(tk, g) for tk, g in self._dupes_all_groups if show.get(tk, True)]
        self._dupes_filtered = filtered

        t = self._tree_dupes_groups
        t.delete(*t.get_children())
        TYPE_LABELS = {"hostname": "По hostname", "ip": "По IP"}
        for i, (tk_, g) in enumerate(filtered):
            label = TYPE_LABELS.get(tk_, tk_)
            reason_short = g["reason"]
            if len(reason_short) > 55:
                reason_short = reason_short[:52] + "…"
            cnt = len(g["hosts"])
            t.insert("", "end",
                     values=(label, reason_short, cnt),
                     tags=(tk_,),
                     iid=str(i))

        # Сбросить детали
        self._tree_dupes_detail.delete(*self._tree_dupes_detail.get_children())
        self._dupes_reason_lbl.configure(text="")

    def _dupes_on_group_select(self, event=None):
        sel = self._tree_dupes_groups.selection()
        if not sel:
            return
        idx = int(sel[0])
        if idx >= len(self._dupes_filtered):
            return
        tk_, group = self._dupes_filtered[idx]
        self._dupes_show_group(group)

    def _dupes_show_group(self, group):
        """Показать хосты выбранной группы дубликатов."""
        t = self._tree_dupes_detail
        t.delete(*t.get_children())
        self._dupes_reason_lbl.configure(text=f"⚠ {group['reason']}")

        for h in group["hosts"]:
            ifaces = h.get("interfaces", [])
            # Собрать все IP интерфейсов
            ip_parts = []
            for iface in ifaces:
                ip = iface.get("ip","").strip()
                if ip and ip not in ("", "0.0.0.0"):
                    itype = IFACE_TYPE.get(str(iface.get("type","1")), "?")
                    ip_parts.append(f"{itype}:{ip}")
            ip_str = "  ".join(ip_parts) if ip_parts else "—"

            grps = ", ".join(g.get("name","") for g in
                             h.get("groups", h.get("hostgroups",[])))
            status = HOST_STATUS.get(str(h.get("status","0")), "?")
            avail_str, _, _ = _iface_avail_str(ifaces)

            tag = "enabled" if str(h.get("status","0")) == "0" else "disabled"
            t.insert("", "end",
                     values=(h.get("host",""), h.get("name",""),
                             grps, ip_str, status, avail_str),
                     tags=(tag,))

    # ── Экспорт ───────────────────────────────────────────────────────────
    def _dupes_export_pdf(self):
        if not self._dupes_all_groups:
            messagebox.showwarning("", "Сначала выполните поиск"); return
        path = filedialog.asksaveasfilename(
            defaultextension=".pdf", filetypes=[("PDF","*.pdf")],
            initialfile=f"duplicates_{datetime.date.today()}.pdf")
        if not path: return
        self._busy(True)
        groups = self._dupes_filtered or self._dupes_all_groups
        def _w():
            try:
                _dupes_to_pdf(path, groups, len(self._hosts))
                self.after(0, lambda: (self._busy(False),
                    self._info(f"PDF: {path}"),
                    messagebox.showinfo("Готово", f"PDF сохранён:\n{path}")))
            except Exception as ex:
                msg = str(ex)
                self.after(0, lambda m=msg: (self._busy(False),
                    messagebox.showerror("Ошибка PDF", m)))
        threading.Thread(target=_w, daemon=True).start()

    def _dupes_export_csv(self):
        if not self._dupes_all_groups:
            messagebox.showwarning("", "Сначала выполните поиск"); return
        path = filedialog.asksaveasfilename(
            defaultextension=".csv", filetypes=[("CSV","*.csv")],
            initialfile=f"duplicates_{datetime.date.today()}.csv")
        if not path: return
        self._busy(True)
        groups = self._dupes_filtered or self._dupes_all_groups
        def _w():
            try:
                _dupes_to_csv(path, groups)
                self.after(0, lambda: (self._busy(False),
                    self._info(f"CSV: {path}"),
                    messagebox.showinfo("Готово", f"CSV сохранён:\n{path}")))
            except Exception as ex:
                msg = str(ex)
                self.after(0, lambda m=msg: (self._busy(False),
                    messagebox.showerror("Ошибка CSV", m)))
        threading.Thread(target=_w, daemon=True).start()

    # ── Вкладки ───────────────────────────────────────────────────────────────
    def _build_tabs(self, parent):
        nb = ttk.Notebook(parent); nb.pack(fill="both", expand=True)

        t1 = ttk.Frame(nb); nb.add(t1, text=" 🔴 Проблемы ")
        self._tree_p = self._tv(t1,
            ("Серьёзность","Проблема","Хост","Начало","Длит.","Подтв."),
            (118,360,160,152,108,68))

        t2 = ttk.Frame(nb); nb.add(t2, text=" 🖥 Хосты ")
        self._build_hosts_tab(t2)

        t3 = ttk.Frame(nb); nb.add(t3, text=" 📋 События ")
        self._build_events_tab(t3)

        t4 = ttk.Frame(nb); nb.add(t4, text=" 📈 Метрики ")
        self._build_metrics_tab(t4)

        t7 = ttk.Frame(nb); nb.add(t7, text=" 🔐 Аутентификация ")
        self._build_auth_tab(t7)

        t8 = ttk.Frame(nb); nb.add(t8, text=" ⚖ Сравнение ")
        self._build_compare_tab(t8)

        t9  = ttk.Frame(nb); nb.add(t9,  text=" 📤 Экспорт шаблонов ")
        self._build_tpl_export_tab(t9)

        t10 = ttk.Frame(nb); nb.add(t10, text=" 🔎 Дубликаты ")
        self._build_dupes_tab(t10)

        t5  = ttk.Frame(nb); nb.add(t5,  text=" 📊 Сводка ")
        self._sum_txt = tk.Text(t5, bg="#181825", fg="#cdd6f4",
                                 font=("Consolas",11), relief="flat",
                                 state="disabled", padx=20, pady=12)
        sb = ttk.Scrollbar(t5, command=self._sum_txt.yview)
        self._sum_txt.configure(yscrollcommand=sb.set)
        self._sum_txt.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")


    # ════════════════════════════════════════════════════════════════════════
    #  Вкладка «Хосты» — объединяет список и детальную карточку
    # ════════════════════════════════════════════════════════════════════════
    def _build_hosts_tab(self, parent):
        BG = "#1e1e2e"; BG2 = "#252535"; FG = "#cdd6f4"; ACC = "#89b4fa"

        # Toolbar с фильтрами
        tb = tk.Frame(parent, bg=BG2, pady=4)
        tb.pack(fill="x", padx=4, pady=(4, 0))

        tk.Label(tb, text="Фильтр:", bg=BG2, fg=ACC,
                 font=("Segoe UI", 9, "bold")).grid(row=0, column=0, padx=(10,4))

        tk.Label(tb, text="Группа:", bg=BG2, fg=FG,
                 font=("Segoe UI", 9)).grid(row=0, column=1, padx=(0,2))
        self._hosts_grp_var = tk.StringVar(value="— все —")
        self._cmb_hosts_grp = ttk.Combobox(tb, textvariable=self._hosts_grp_var,
                                            state="readonly", width=22)
        self._cmb_hosts_grp.grid(row=0, column=2, padx=(0,8))

        tk.Label(tb, text="Шаблон:", bg=BG2, fg=FG,
                 font=("Segoe UI", 9)).grid(row=0, column=3, padx=(0,2))
        self._hosts_tpl_var = tk.StringVar(value="— все —")
        self._cmb_hosts_tpl = ttk.Combobox(tb, textvariable=self._hosts_tpl_var,
                                            state="readonly", width=26)
        self._cmb_hosts_tpl.grid(row=0, column=4, padx=(0,8))

        tk.Label(tb, text="🔍", bg=BG2, fg=FG).grid(row=0, column=5, padx=(0,2))
        self._hosts_search_var = tk.StringVar()
        self._hosts_search_var.trace_add("write",
            lambda *_: self._hosts_apply_filter())
        ttk.Entry(tb, textvariable=self._hosts_search_var,
                  width=22).grid(row=0, column=6, padx=(0,6))

        ttk.Button(tb, text="🔄 Обновить расширенно",
                   command=self._hosts_load_detail).grid(row=0, column=7, padx=(0,4))

        self._hosts_status = tk.Label(
            tb, text="Двойной клик по хосту — открыть карточку",
            bg=BG2, fg="#6c7086", font=("Segoe UI", 8))
        self._hosts_status.grid(row=1, column=0, columnspan=8,
                                 padx=10, pady=(2,0), sticky="w")

        self._cmb_hosts_grp.bind("<<ComboboxSelected>>",
                                  lambda _: self._hosts_apply_filter())
        self._cmb_hosts_tpl.bind("<<ComboboxSelected>>",
                                  lambda _: self._hosts_apply_filter())

        # Таблица хостов
        cols = ("Хост", "Имя", "Группы", "Шаблоны",
                "Статус", "Доступн.", "Интерфейсы",
                "Agent ver.", "Тригг.", "Обслуж.")
        widths = (130, 140, 140, 160, 70, 90, 180, 80, 55, 130)

        tree_frm = tk.Frame(parent, bg=BG)
        tree_frm.pack(fill="both", expand=True, padx=4, pady=4)

        self._tree_h = ttk.Treeview(tree_frm, columns=cols,
                                     show="headings", selectmode="browse")
        for col, w in zip(cols, widths):
            self._tree_h.heading(col, text=col,
                command=lambda c=col, t=self._tree_h: self._sort(t, c))
            self._tree_h.column(col, width=w, minwidth=40)
        vs = ttk.Scrollbar(tree_frm, orient="vertical", command=self._tree_h.yview)
        hs = ttk.Scrollbar(tree_frm, orient="horizontal", command=self._tree_h.xview)
        self._tree_h.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        self._tree_h.grid(row=0, column=0, sticky="nsew")
        vs.grid(row=0, column=1, sticky="ns")
        hs.grid(row=1, column=0, sticky="ew")
        tree_frm.rowconfigure(0, weight=1); tree_frm.columnconfigure(0, weight=1)

        self._tree_h._all = []
        self._tree_h.tag_configure("ok",    foreground="#a6e3a1")
        self._tree_h.tag_configure("bad",   foreground="#f38ba8")
        self._tree_h.tag_configure("unk",   foreground="#f9e2af")
        self._tree_h.tag_configure("maint", foreground="#fab387")

        # Двойной клик по строке → окно карточки
        self._tree_h.bind("<Double-1>", self._hosts_open_card)
        self._tree_h.bind("<Return>",   self._hosts_open_card)

        # Кеш agent-версий и доступности latest data по hostid
        self._hosts_agent_ver = {}   # hostid -> str
        self._hosts_latest_ok = {}   # hostid -> set of iface types seen as working

    # ── Загрузка расширенных данных (group names + templates + agent ver) ──
    def _hosts_load_detail(self):
        if not self.zapi:
            messagebox.showwarning("", "Подключитесь к Zabbix"); return
        self._busy(True)
        self._hosts_status.configure(text="Загрузка расширенных данных…")
        def _w():
            try:
                detail = self.zapi.get_hosts_detail()
                latest_ok = self.zapi.get_hosts_latest_activity(
                    [h["hostid"] for h in (self._hosts or [])],
                    age_seconds=900)  # 15 минут
                self.after(0, lambda: self._hosts_populate_detail(detail, latest_ok))
            except Exception as ex:
                msg = str(ex)
                self.after(0, lambda m=msg: (
                    self._busy(False),
                    self._hosts_status.configure(text=f"Ошибка: {m}"),
                    self.log(f"  [Hosts] Ошибка: {m}", "err")))
        threading.Thread(target=_w, daemon=True).start()

    def _hosts_populate_detail(self, detail, latest_ok):
        """
        detail — результат get_hosts_detail (с templates/macros/tags/agent ver)
        latest_ok — {hostid: bool}  — хост имеет свежие данные
        """
        self._hosts_detail = detail or []
        self._hosts_latest_ok = latest_ok or {}

        # Карта hostid -> детальный объект
        det_by_id = {d["hostid"]: d for d in self._hosts_detail}

        # Обогащаем _hosts данными: templates, agent version
        for h in (self._hosts or []):
            d = det_by_id.get(h["hostid"])
            if d:
                h["_templates"] = d.get("parentTemplates", [])
                h["_agent_version"] = d.get("_agent_version", "")
                h["_tags"]      = d.get("tags", [])
                h["_macros"]    = d.get("macros", [])

        # Обновить выпадающие списки фильтров
        grps = set(); tpls = set()
        for d in self._hosts_detail:
            for g in d.get("groups", d.get("hostgroups", [])):
                grps.add(g.get("name",""))
            for t in d.get("parentTemplates", []):
                tpls.add(t.get("name",""))
        self._cmb_hosts_grp["values"] = ["— все —"] + sorted(filter(None, grps))
        self._cmb_hosts_tpl["values"] = ["— все —"] + sorted(filter(None, tpls))

        # Перерисовать таблицу
        self._fill_h(self._hosts or [])
        self._busy(False)
        self._hosts_status.configure(
            text=f"Расширенные данные загружены: {len(self._hosts_detail)} хостов")
        self.log(f"  [Hosts] Детали: {len(self._hosts_detail)}", "ok")

    def _hosts_apply_filter(self):
        """Применить фильтры к уже заполненной таблице."""
        if not hasattr(self._tree_h, "_all") or not self._tree_h._all:
            return
        grp = self._hosts_grp_var.get()
        tpl = self._hosts_tpl_var.get()
        q   = self._hosts_search_var.get().lower().strip()

        # Перезаполняем
        t = self._tree_h
        t.delete(*t.get_children())
        for iid, row, h in self._tree_h._all:
            # Фильтр по группе
            if grp and grp != "— все —":
                host_grps = [g.get("name","") for g in
                             h.get("groups", h.get("hostgroups",[]))]
                if grp not in host_grps:
                    continue
            # Фильтр по шаблону
            if tpl and tpl != "— все —":
                host_tpls = [tp.get("name","") for tp in h.get("_templates",[])]
                if tpl not in host_tpls:
                    continue
            # Текстовый поиск
            if q and q not in " ".join(str(v) for v in row).lower():
                continue
            tag = ("maint" if str(h.get("maintenance_status","0")) == "1" else
                   "ok"    if row[5] == "Available" else
                   "bad"   if row[5] == "Unavailable" else "unk")
            t.insert("", "end", values=row, tags=(tag,), iid=h["hostid"])

    def _hosts_open_card(self, event=None):
        """Открывает модальное окно с карточкой хоста."""
        sel = self._tree_h.selection()
        if not sel:
            return
        hostid = sel[0]
        # Ищем полный объект — в _hosts_detail если есть, иначе в _hosts
        d = None
        for x in (self._hosts_detail or []):
            if x.get("hostid") == hostid:
                d = x
                break
        if not d:
            for x in (self._hosts or []):
                if x.get("hostid") == hostid:
                    d = x
                    break
        if not d:
            return
        self._open_host_card_window(d)

    def _open_host_card_window(self, h):
        """Модальное окно с полной карточкой хоста."""
        BG = "#181825"; FG = "#cdd6f4"; ACC = "#89b4fa"
        win = tk.Toplevel(self)
        win.title(f"Карточка хоста — {h.get('host','')}")
        win.configure(bg=BG)
        win.geometry("760x640")
        win.minsize(540, 420)
        try: win.transient(self)
        except: pass

        tk.Label(win, text=f"  {h.get('host','')} / {h.get('name','')}",
                 bg=BG, fg=ACC,
                 font=("Segoe UI", 13, "bold")).pack(fill="x", pady=(10, 2))
        ttk.Separator(win).pack(fill="x", padx=8, pady=(0,6))

        txt = tk.Text(win, bg=BG, fg=FG, font=("Consolas", 10),
                      relief="flat", state="normal", wrap="word",
                      padx=16, pady=10, selectbackground="#313244")
        sb = ttk.Scrollbar(win, command=txt.yview)
        txt.configure(yscrollcommand=sb.set)
        txt.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        txt.tag_configure("hdr",  foreground=ACC, font=("Segoe UI",11,"bold"))
        txt.tag_configure("key",  foreground="#cba6f7", font=("Consolas",10,"bold"))
        txt.tag_configure("val",  foreground=FG, font=("Consolas",10))
        txt.tag_configure("ok",   foreground="#a6e3a1")
        txt.tag_configure("bad",  foreground="#f38ba8")
        txt.tag_configure("dim",  foreground="#6c7086")

        def sec(t): txt.insert("end", f"\n● {t}\n", "hdr")
        def row(k, v, tag="val"):
            txt.insert("end", f"  {k:<26}", "key")
            txt.insert("end", f"{v}\n", tag)

        YESNO = lambda v: "Да" if str(v) == "1" else "Нет"

        sec("Основное")
        row("Host name",    h.get("host",""))
        row("Visible name", h.get("name","") or h.get("host",""))
        enabled = str(h.get("status","0")) == "0"
        row("Enabled", "✅ Включён" if enabled else "❌ Отключён",
            "ok" if enabled else "bad")
        row("Monitored by", h.get("_proxy_name","Zabbix Server"))
        if h.get("description"):
            row("Description", h["description"][:120])
        if str(h.get("maintenance_status","0")) == "1":
            row("Обслуживание",
                h.get("_maint_name","") or "(без имени)", "ok")

        sec("Host groups")
        grps = h.get("groups", h.get("hostgroups", []))
        if grps:
            for g in grps:
                txt.insert("end", f"  • {g.get('name','')}\n", "val")
        else:
            txt.insert("end", "  (нет)\n", "dim")

        sec("Templates")
        tpls = h.get("parentTemplates", h.get("_templates", []))
        if tpls:
            for tp in tpls:
                txt.insert("end", f"  • {tp.get('name','?')}\n", "val")
        else:
            txt.insert("end", "  (нет шаблонов)\n", "dim")

        sec("Interfaces")
        ifaces = h.get("interfaces", [])
        if ifaces:
            for iface in ifaces:
                itype = IFACE_TYPE.get(str(iface.get("type","1")), "?")
                addr  = iface.get("ip","") or iface.get("dns","")
                port  = iface.get("port","")
                main  = " [основной]" if iface.get("main") == "1" else ""
                av    = str(iface.get("available","0"))
                av_sym = {"0":"?","1":"✓","2":"✗"}.get(av,"?")
                av_tag = {"0":"dim","1":"ok","2":"bad"}.get(av,"dim")
                txt.insert("end", f"  • {itype:<6} {addr}:{port}{main}  ", "val")
                txt.insert("end", f"[{av_sym}]\n", av_tag)
                if str(iface.get("type","1")) == "1":
                    av_ver = h.get("_agent_version","")
                    txt.insert("end", f"    {'Zabbix agent ver':<24}", "key")
                    txt.insert("end",
                               f"{av_ver if av_ver else '(не получен)'}\n",
                               "val" if av_ver else "dim")
        else:
            txt.insert("end", "  (нет интерфейсов)\n", "dim")

        sec("Tags")
        tags = h.get("tags", h.get("_tags", []))
        if tags:
            for tg in tags:
                v = tg.get("value","")
                txt.insert("end", f"  • {tg.get('tag','')}", "key")
                txt.insert("end", f"{': ' + v if v else ''}\n", "val")
        else:
            txt.insert("end", "  (нет тегов)\n", "dim")

        sec("Host macros")
        macros = h.get("macros", h.get("_macros", []))
        if macros:
            for m in macros:
                mtype = str(m.get("type","0"))
                val = "***" if mtype == "1" else m.get("value","")
                txt.insert("end", f"  {m.get('macro',''):<30}", "key")
                txt.insert("end", f"= {val}", "val")
                if m.get("description"):
                    txt.insert("end", f"   # {m['description'][:40]}", "dim")
                txt.insert("end", "\n")
        else:
            txt.insert("end", "  (нет макросов)\n", "dim")

        txt.configure(state="disabled")

        btn_frm = tk.Frame(win, bg=BG)
        btn_frm.pack(fill="x", side="bottom", pady=6)
        ttk.Button(btn_frm, text="Закрыть",
                   command=win.destroy).pack(side="right", padx=10)

    # ════════════════════════════════════════════════════════════════════════
    #  Вкладка «События» — с фильтрами по серьёзности, хосту, времени
    # ════════════════════════════════════════════════════════════════════════
    def _build_events_tab(self, parent):
        BG = "#1e1e2e"; BG2 = "#252535"; FG = "#cdd6f4"; ACC = "#89b4fa"

        tb = tk.Frame(parent, bg=BG2, pady=4)
        tb.pack(fill="x", padx=4, pady=(4,0))

        # Серьёзность (multi-select через checkboxes)
        tk.Label(tb, text="Серьёзность:", bg=BG2, fg=ACC,
                 font=("Segoe UI", 9, "bold")).grid(row=0, column=0, padx=(10,4))
        self._ev_sev_vars = {}
        for i, (code, name) in enumerate(SEVERITY_NAMES.items()):
            v = tk.BooleanVar(value=True)
            self._ev_sev_vars[code] = v
            cb = tk.Checkbutton(tb, text=name.split()[0] if " " in name else name,
                                variable=v,
                                bg=BG2, fg=FG, selectcolor="#313244",
                                activebackground=BG2, activeforeground=ACC,
                                font=("Segoe UI", 8),
                                command=self._ev_apply_filter)
            cb.grid(row=0, column=1+i, padx=1)

        # Хост
        tk.Label(tb, text="Хост:", bg=BG2, fg=FG,
                 font=("Segoe UI", 9)).grid(row=1, column=0, padx=(10,4), pady=(4,0))
        self._ev_host_var = tk.StringVar(value="— все —")
        self._cmb_ev_host = ttk.Combobox(tb, textvariable=self._ev_host_var,
                                          state="readonly", width=28)
        self._cmb_ev_host.grid(row=1, column=1, columnspan=3, pady=(4,0), sticky="w")
        self._cmb_ev_host.bind("<<ComboboxSelected>>",
                                lambda _: self._ev_apply_filter())

        # Время
        tk.Label(tb, text="С:", bg=BG2, fg=FG,
                 font=("Segoe UI", 9)).grid(row=1, column=4, padx=(8,2), pady=(4,0))
        self._ev_from_var = tk.StringVar()
        e_from = ttk.Entry(tb, textvariable=self._ev_from_var, width=18)
        e_from.grid(row=1, column=5, pady=(4,0))
        tk.Label(tb, text="По:", bg=BG2, fg=FG,
                 font=("Segoe UI", 9)).grid(row=1, column=6, padx=(6,2), pady=(4,0))
        self._ev_till_var = tk.StringVar()
        e_till = ttk.Entry(tb, textvariable=self._ev_till_var, width=18)
        e_till.grid(row=1, column=7, pady=(4,0))

        ttk.Button(tb, text="🔍 Применить",
                   command=self._ev_apply_filter).grid(
                   row=1, column=8, padx=(8,2), pady=(4,0))
        ttk.Button(tb, text="⟲ Сброс",
                   command=self._ev_reset_filter).grid(
                   row=1, column=9, padx=(0,4), pady=(4,0))

        self._ev_status = tk.Label(
            tb, text="Формат времени: YYYY-MM-DD HH:MM (оставьте пустым — без ограничения)",
            bg=BG2, fg="#6c7086", font=("Segoe UI", 8))
        self._ev_status.grid(row=2, column=0, columnspan=10,
                              padx=10, pady=(2,0), sticky="w")

        # Таблица событий
        cols = ("ID", "Серьёзность", "Событие", "Хост", "Время", "Восст.", "Длит.")
        widths = (80, 110, 320, 150, 140, 140, 90)
        tree_frm = tk.Frame(parent, bg=BG)
        tree_frm.pack(fill="both", expand=True, padx=4, pady=4)

        self._tree_e = ttk.Treeview(tree_frm, columns=cols,
                                     show="headings", selectmode="browse")
        for c, w in zip(cols, widths):
            self._tree_e.heading(c, text=c,
                command=lambda _c=c, _t=self._tree_e: self._sort(_t, _c))
            self._tree_e.column(c, width=w, minwidth=40)
        vs = ttk.Scrollbar(tree_frm, orient="vertical", command=self._tree_e.yview)
        hs = ttk.Scrollbar(tree_frm, orient="horizontal", command=self._tree_e.xview)
        self._tree_e.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        self._tree_e.grid(row=0, column=0, sticky="nsew")
        vs.grid(row=0, column=1, sticky="ns")
        hs.grid(row=1, column=0, sticky="ew")
        tree_frm.rowconfigure(0, weight=1); tree_frm.columnconfigure(0, weight=1)
        self._tree_e._all = []

    def _ev_reset_filter(self):
        for v in self._ev_sev_vars.values():
            v.set(True)
        self._ev_host_var.set("— все —")
        self._ev_from_var.set("")
        self._ev_till_var.set("")
        self._ev_apply_filter()

    def _ev_apply_filter(self):
        if not self._tree_e._all:
            return
        # Серьёзности
        sev_show = {code for code, v in self._ev_sev_vars.items() if v.get()}
        host_sel = self._ev_host_var.get()
        from_ts  = self._parse_dt(self._ev_from_var.get())
        till_ts  = self._parse_dt(self._ev_till_var.get())

        t = self._tree_e
        t.delete(*t.get_children())
        shown = 0
        for iid, row, e in t._all:
            sid = str(e.get("severity","0"))
            if sid not in sev_show:
                continue
            if host_sel and host_sel != "— все —":
                eh = e.get("hosts",[{}])
                hname = eh[0].get("host","") if eh else ""
                if hname != host_sel:
                    continue
            clk = int(e.get("clock","0") or 0)
            if from_ts and clk < from_ts:
                continue
            if till_ts and clk > till_ts:
                continue
            t.insert("", "end", values=row, tags=(f"s{sid}",))
            t.tag_configure(f"s{sid}",
                            foreground=SEVERITY_HEX.get(sid,"#cdd6f4"))
            shown += 1
        total = len(t._all)
        self._ev_status.configure(
            text=f"Показано: {shown} из {total}  "
                 "|  Формат: YYYY-MM-DD HH:MM")

    def _parse_dt(self, s: str):
        s = s.strip()
        if not s:
            return None
        for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d"):
            try:
                dt = datetime.datetime.strptime(s, fmt)
                return int(dt.timestamp())
            except Exception:
                continue
        return None

    def _tv(self, parent, cols, widths):
        f = tk.Frame(parent, bg="#1e1e2e")
        f.pack(fill="both", expand=True, padx=4, pady=(4,0))
        t = ttk.Treeview(f, columns=cols, show="headings", selectmode="extended")
        for c, w in zip(cols, widths):
            t.heading(c, text=c,
                      command=lambda _c=c, _t=t: self._sort(_t, _c))
            t.column(c, width=w, minwidth=40)
        vs = ttk.Scrollbar(f, orient="vertical",   command=t.yview)
        hs = ttk.Scrollbar(f, orient="horizontal", command=t.xview)
        t.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        t.grid(row=0, column=0, sticky="nsew"); vs.grid(row=0, column=1, sticky="ns")
        hs.grid(row=1, column=0, sticky="ew")
        f.rowconfigure(0,weight=1); f.columnconfigure(0,weight=1)
        sf = tk.Frame(parent, bg="#313244")
        sf.pack(fill="x", padx=4, pady=(0,4))
        tk.Label(sf, text="🔍", bg="#313244", fg="#cdd6f4").pack(side="left", padx=6)
        sv = tk.StringVar()
        sv.trace_add("write", lambda *_: self._filter(t, sv.get()))
        ttk.Entry(sf, textvariable=sv, width=40).pack(side="left")
        t._sv = sv; t._all = []
        return t

    # ── Вкладка Метрики ───────────────────────────────────────────────────────
    def _build_metrics_tab(self, parent):
        top = tk.Frame(parent, bg="#1e1e2e")
        top.pack(fill="x", padx=8, pady=6)
        ttk.Label(top, text="Хост:").pack(side="left")
        self._cmb_hm = ttk.Combobox(top, width=28, state="readonly")
        self._cmb_hm.pack(side="left", padx=6)
        ttk.Label(top, text="Поиск метрики:").pack(side="left")
        self._e_ms = ttk.Entry(top, width=22)
        self._e_ms.pack(side="left", padx=6)
        ttk.Button(top, text="🔍 Загрузить метрики", command=self._load_items).pack(side="left", padx=4)
        ttk.Button(top, text="📈 Построить график",  command=self._load_history).pack(side="left", padx=2)

        mid = tk.Frame(parent, bg="#1e1e2e")
        mid.pack(fill="both", expand=True, padx=8, pady=4)

        lf = ttk.LabelFrame(mid, text=" Метрики ", padding=4)
        lf.pack(side="left", fill="both", expand=True)
        self._tree_items = ttk.Treeview(lf,
            columns=("Метрика","Ключ","Последнее","Ед."),
            show="headings", selectmode="extended")
        for c, w in [("Метрика",200),("Ключ",180),("Последнее",100),("Ед.",50)]:
            self._tree_items.heading(c,text=c); self._tree_items.column(c,width=w)
        vs2 = ttk.Scrollbar(lf, orient="vertical",
                             command=self._tree_items.yview)
        self._tree_items.configure(yscrollcommand=vs2.set)
        self._tree_items.pack(side="left", fill="both", expand=True)
        vs2.pack(side="right", fill="y")

        self._chart_frame = ttk.LabelFrame(mid, text=" График ", padding=4)
        self._chart_frame.pack(side="left", fill="both", expand=True, padx=(6,0))

    # ── Подключение ──────────────────────────────────────────────────────────
    def _connect(self):
        url  = self._e_url.get().strip()
        user = self._e_user.get().strip()
        pwd  = self._e_pass.get()
        if not url or not user:
            messagebox.showerror("Ошибка","Укажите URL и пользователя"); return
        self._btn_con.configure(state="disabled", text="⏳ Подключение…")
        self._busy(True)
        def _w():
            try:
                api = ZabbixAPI(url, verify_ssl=self._v_ssl.get())
                api.login(user, pwd)
                ver = api.version()
                self.zapi = api
                self.after(0, lambda: self._on_ok(ver))
            except Exception as ex:
                msg = str(ex)
                self.after(0, lambda m=msg: self._on_err(m))
        threading.Thread(target=_w, daemon=True).start()

    def _on_ok(self, ver):
        self._lbl_ver.config(text=f"Zabbix {ver}")
        self._lbl_conn.config(text="● Подключено", fg="#a6e3a1")
        self._btn_con.configure(state="disabled", text="✅ Подключено")
        self._btn_dis.configure(state="normal")
        # Подключить логгер к API-клиенту
        self.zapi._log_cb = lambda msg, lvl="info": self.log(msg, lvl)
        self._busy(False)
        self.log(f"✅ Подключено к Zabbix {ver}", "ok")
        self.log(f"   API version tuple: {self.zapi._api_version}", "info")
        self.log(f"   URL: {self.zapi.url}", "info")
        self._info("Подключено. Загружаю данные…")
        self._refresh_all()

    def _on_err(self, msg):
        self._btn_con.configure(state="normal", text="🔌 Подключиться")
        self._lbl_conn.config(text="● Ошибка", fg="#f38ba8")
        self._busy(False); self._info(f"Ошибка: {msg}")
        self.log(f"✗ Ошибка подключения: {msg}", "err")
        messagebox.showerror("Ошибка подключения", msg)

    def _disconnect(self):
        if self.zapi: self.zapi.logout(); self.zapi = None
        self._btn_con.configure(state="normal", text="🔌 Подключиться")
        self._btn_dis.configure(state="disabled")
        self._lbl_conn.config(text="● Не подключено", fg="#f38ba8")
        self._lbl_ver.config(text=""); self._info("Отключено")
        for t in (self._tree_p, self._tree_h, self._tree_e, self._tree_items):
            t.delete(*t.get_children())

    # ── Загрузка ─────────────────────────────────────────────────────────────
    def _refresh_all(self):
        if not self.zapi:
            messagebox.showwarning("","Сначала подключитесь"); return
        self._busy(True); self._info("Загружаю данные…")
        self.log("━━━ Запрос данных ━━━", "info")
        def _w():
            try:
                sev   = int(self._cmb_sev.get().split()[0])
                limit = int(self._v_limit.get() or 500)
                self.after(0, lambda: self.log(f"  Фильтр: severity>={sev}, limit={limit}", "info"))

                self.after(0, lambda: self.log("  [1/3] problem.get ...", "req"))
                p = self.zapi.get_problems(sev_min=sev, limit=limit)
                self.after(0, lambda n=len(p): self.log(f"  → Проблем получено: {n}", "ok"))

                self.after(0, lambda: self.log("  [2/3] host.get ...", "req"))
                h = self.zapi.get_hosts()
                self.after(0, lambda n=len(h): self.log(f"  → Хостов получено: {n}", "ok"))

                try:
                    df = datetime.datetime.strptime(self._e_from.get().strip(),"%Y-%m-%d")
                    dt = datetime.datetime.strptime(self._e_till.get().strip(),"%Y-%m-%d"
                         ) + datetime.timedelta(days=1)
                except ValueError:
                    df = datetime.datetime.now()-datetime.timedelta(days=7)
                    dt = datetime.datetime.now()
                self.after(0, lambda: self.log(
                    f"  [3/3] event.get  {df.date()} → {dt.date()} ...", "req"))
                ev = self.zapi.get_events(int(df.timestamp()),int(dt.timestamp()),limit=limit)
                self.after(0, lambda n=len(ev): self.log(f"  → Событий получено: {n}", "ok"))

                # Дополнительная диагностика если всё пусто
                if not p and not h and not ev:
                    self.after(0, lambda: self.log(
                        "  ⚠ Все запросы вернули 0 записей. "
                        "Проверьте права пользователя (нужны Zabbix User+) "
                        "и наличие данных в системе.", "warn"))

                latest_ok = {}
                try:
                    latest_ok = self.zapi.get_hosts_latest_activity(
                        [hh["hostid"] for hh in h], age_seconds=900)
                except Exception:
                    pass
                def _done(la=latest_ok, pp=p, hh=h, ee=ev):
                    self._hosts_latest_ok = la
                    self._populate(pp, hh, ee)
                self.after(0, _done)
            except Exception as ex:
                import traceback
                tb = traceback.format_exc()
                msg = f"Ошибка загрузки: {ex}"
                self.after(0, lambda m=msg, t=tb: (
                    self._busy(False),
                    self._info(m),
                    self.log(m, "err"),
                    self.log(t, "err"),
                ))
        threading.Thread(target=_w, daemon=True).start()

    def _populate(self, p, h, ev):
        self._problems, self._hosts, self._events = p, h, ev
        self._fill_p(p); self._fill_h(h); self._fill_e(ev); self._fill_sum()
        # Обновить фильтры объединённой вкладки «Хосты»
        try:
            grps = set()
            for hh in h:
                grp_list = hh.get("groups", hh.get("hostgroups", []))
                for g in grp_list:
                    grps.add(g.get("name",""))
            grp_vals = ["— все —"] + sorted(x for x in grps if x)
            if hasattr(self, "_cmb_hosts_grp"):
                self._cmb_hosts_grp["values"] = grp_vals
        except Exception:
            pass
        self._cmb_hm["values"] = [
            f"{hh.get('host','')}  [{hh.get('hostid','')}]" for hh in h]
        if h: self._cmb_hm.current(0)

        self._busy(False)
        summary = f"Загружено: {len(p)} проблем | {len(h)} хостов | {len(ev)} событий"
        self._info(summary)
        self.log(f"✓ {summary}", "ok")
        # Диагностика первых хостов
        if h:
            sample = h[:3]
            for hh in sample:
                self.log(
                    f"   Хост: {hh.get('host','')}  "
                    f"ifaces={[str(i.get('available','?')) for i in hh.get('interfaces',[])]}  "
                    f"status={hh.get('status','?')}  "
                    f"groups={[g['name'] for g in hh.get('groups',[])]}",
                    "info")
        if p:
            self.log(f"   Первая проблема: {p[0].get('name','')} "
                     f"(sev={p[0].get('severity','?')})", "info")

    def _fill_p(self, data):
        t = self._tree_p; t.delete(*t.get_children()); t._all = []
        for p in data:
            sid = str(p.get("severity","0"))
            hl  = p.get("hosts",[{}]); hn = hl[0].get("host","") if hl else ""
            row = (SEVERITY_NAMES.get(sid,sid), p.get("name",""), hn,
                   ts2str(p.get("clock",0)), dur_str(p.get("clock",0)),
                   len(p.get("acknowledges",[])))
            iid = t.insert("","end",values=row, tags=(f"s{sid}",))
            t.tag_configure(f"s{sid}", foreground=SEVERITY_HEX.get(sid,"#cdd6f4"))
            t._all.append((iid,row))

    def _fill_h(self, data):
        t = self._tree_h
        t.delete(*t.get_children())
        t._all = []
        for h in data:
            ifaces = h.get("interfaces", [])
            # Доступность: используем interface.available + фоллбэк на latest data
            avail_str = self._compute_host_availability(h)

            # Интерфейсы: «Agent 192.168.1.10:10050  SNMP 10.0.0.5:161»
            if ifaces:
                iface_parts = []
                for iface in ifaces:
                    typ = IFACE_TYPE.get(str(iface.get("type","1")), "?")
                    ip  = iface.get("ip","") or iface.get("dns","")
                    port = iface.get("port","")
                    main = " ●" if str(iface.get("main","0")) == "1" else ""
                    iface_parts.append(f"{typ} {ip}:{port}{main}")
                ifaces_str = "  ".join(iface_parts)
            else:
                ifaces_str = "—"

            # Группы и шаблоны
            grps_list = h.get("groups", h.get("hostgroups", []))
            grps_str  = ", ".join(g.get("name","") for g in grps_list)
            tpls_list = h.get("_templates", h.get("parentTemplates", []))
            tpls_str  = ", ".join(tp.get("name","") for tp in tpls_list)

            # Agent version
            agent_ver = h.get("_agent_version","") or ""

            in_maint  = str(h.get("maintenance_status","0")) == "1"
            maint_str = ""
            if in_maint:
                mname     = h.get("_maint_name","")
                maint_str = f"🔧 {mname}" if mname else "🔧 Обслуживание"

            row = (h.get("host",""), h.get("name",""),
                   grps_str, tpls_str,
                   HOST_STATUS.get(str(h.get("status","0")),""),
                   avail_str,
                   ifaces_str,
                   agent_ver,
                   h.get("triggers","0"),
                   maint_str)

            tag = ("maint" if in_maint else
                   "ok"    if avail_str == "Available" else
                   "bad"   if avail_str == "Unavailable" else "unk")
            iid = t.insert("","end", values=row, tags=(tag,), iid=h["hostid"])
            t._all.append((iid, row, h))

    def _compute_host_availability(self, h):
        """
        Вычисляет статус доступности хоста:
        1) Для Agent — использует interface.available
        2) Для SNMP/IPMI/JMX — если interface.available Unknown,
           проверяет «latest data» (кеш self._hosts_latest_ok)
        Возвращает: "Available" / "Unavailable" / "Unknown"
        """
        ifaces = h.get("interfaces", [])
        if not ifaces:
            # Fallback на устаревшее host.available
            return HOST_AVAIL.get(str(h.get("available","0")), "Unknown")

        has_unavail = False
        has_avail   = False
        has_unknown = False
        for iface in ifaces:
            av    = str(iface.get("available","0"))
            itype = str(iface.get("type","1"))
            if av == "1":
                has_avail = True
            elif av == "2":
                has_unavail = True
            else:
                # Unknown — для не-Agent проверяем latest data
                if itype != "1" and self._hosts_latest_ok.get(h["hostid"]):
                    has_avail = True
                else:
                    has_unknown = True

        if has_unavail:
            return "Unavailable"
        if has_avail:
            return "Available"
        return "Unknown"

    def _fill_e(self, data):
        t = self._tree_e
        t.delete(*t.get_children())
        t._all = []
        # Собрать уникальные хосты для combobox
        hosts_set = set()
        for e in data:
            eh = e.get("hosts",[])
            if eh:
                hosts_set.add(eh[0].get("host",""))
        host_list = ["— все —"] + sorted(h for h in hosts_set if h)
        try:
            self._cmb_ev_host["values"] = host_list
            if self._ev_host_var.get() not in host_list:
                self._ev_host_var.set("— все —")
        except Exception:
            pass

        for e in data:
            sid = str(e.get("severity","0"))
            rc  = e.get("r_clock","0") or "0"
            eh  = e.get("hosts",[])
            hname = eh[0].get("host","") if eh else ""
            clk = int(e.get("clock","0") or 0)
            # Длительность
            if rc and rc != "0":
                dur = dur_str_from(clk, int(rc))
            else:
                dur = dur_str(clk)  # от clock до now
            row = (e.get("eventid",""),
                   SEVERITY_NAMES.get(sid, sid),
                   e.get("name",""),
                   hname,
                   ts2str(clk),
                   ts2str(rc) if rc not in ("","0") else "—",
                   dur)
            iid = t.insert("","end", values=row, tags=(f"s{sid}",))
            t.tag_configure(f"s{sid}",
                            foreground=SEVERITY_HEX.get(sid,"#cdd6f4"))
            t._all.append((iid, row, e))
        # Применить фильтры (если они настроены)
        try:
            self._ev_apply_filter()
        except Exception:
            pass

    def _fill_sum(self):
        p, h, ev = self._problems, self._hosts, self._events
        sc: dict = {}
        for x in p: sc[str(x.get("severity","0"))] = sc.get(str(x.get("severity","0")),0)+1
        av = {"Available":0,"Unavailable":0,"Unknown":0}
        for x in h: av[HOST_AVAIL.get(str(x.get("available","0")),"Unknown")] += 1
        lines = ["═"*62,"  СВОДНЫЙ ОТЧЁТ — Zabbix Reporter v3.0",
                 f"  {datetime.datetime.now():%Y-%m-%d %H:%M:%S}","═"*62,"",
                 f"  ПРОБЛЕМЫ (всего: {len(p)})", "  "+"-"*44]
        for k, n in SEVERITY_NAMES.items():
            c = sc.get(k,0)
            lines.append(f"  {n:<20} {c:>4}  {'█'*min(c,30)}")
        lines += ["",f"  ХОСТЫ (всего: {len(h)})", "  "+"-"*44]
        for k, c in av.items(): lines.append(f"  {k:<20} {c:>4}")
        lines += ["",f"  СОБЫТИЯ (всего: {len(ev)})","═"*62]
        self._sum_txt.configure(state="normal")
        self._sum_txt.delete("1.0","end")
        self._sum_txt.insert("1.0","\n".join(lines))
        self._sum_txt.configure(state="disabled")

    # ── Метрики ───────────────────────────────────────────────────────────────
    def _load_items(self):
        if not self.zapi: messagebox.showwarning("","Подключитесь"); return
        sel = self._cmb_hm.get()
        if not sel: return
        hid = sel.split("[")[-1].rstrip("]").strip()
        search = self._e_ms.get().strip()
        self._busy(True)
        def _w():
            try:
                items = self.zapi.get_items([hid], search=search, limit=200)
                self.after(0, lambda: self._fill_items(items))
            except Exception as ex:
                msg = str(ex)
                self.after(0, lambda m=msg: (self._busy(False), self._info(m)))
        threading.Thread(target=_w, daemon=True).start()

    def _fill_items(self, items):
        self._items_cache = items
        t = self._tree_items; t.delete(*t.get_children())
        for it in items:
            t.insert("","end",values=(it.get("name",""),it.get("key_",""),
                                       it.get("lastvalue",""),it.get("units","")))
        self._busy(False); self._info(f"Найдено метрик: {len(items)}")
        self.log(f"  Метрик найдено: {len(items)}", "ok" if items else "warn")

    def _load_history(self):
        if not self.zapi: return
        sel = self._tree_items.selection()
        if not sel: messagebox.showinfo("","Выберите метрики"); return
        try:
            df = datetime.datetime.strptime(self._e_from.get().strip(),"%Y-%m-%d")
            dt = datetime.datetime.strptime(self._e_till.get().strip(),"%Y-%m-%d"
                 ) + datetime.timedelta(days=1)
        except ValueError:
            df = datetime.datetime.now()-datetime.timedelta(days=7)
            dt = datetime.datetime.now()
        chosen = []
        for iid in sel:
            row = self._tree_items.item(iid,"values")
            for it in self._items_cache:
                if it.get("name","") == row[0] and it.get("key_","") == row[1]:
                    chosen.append(it); break
        if not chosen: return
        self._busy(True)
        def _w():
            try:
                res = {}
                for it in chosen:
                    h = self.zapi.get_history([it["itemid"]],
                                               int(df.timestamp()),int(dt.timestamp()),
                                               history=int(it.get("value_type",0)),limit=500)
                    res[it["name"]] = h
                title = (self._cmb_hm.get().split("[")[0].strip()
                         + "  " + datetime.date.today().isoformat())
                self._metrics[title] = res
                self.after(0, lambda: self._draw_chart(res, title))
            except Exception as ex:
                msg = str(ex)
                self.after(0, lambda m=msg: (self._busy(False), self._info(m)))
        threading.Thread(target=_w, daemon=True).start()

    def _draw_chart(self, hdata, title):
        for w in self._chart_frame.winfo_children(): w.destroy()
        if not MPL_TK_OK:
            tk.Label(self._chart_frame, text="Установите matplotlib",
                     bg="#1e1e2e", fg="#f38ba8").pack(expand=True)
            self._busy(False); return
        fig, ax = plt.subplots(figsize=(6,3.5))
        fig.patch.set_facecolor("#1e1e2e"); ax.set_facecolor("#181825")
        clrs = ["#89b4fa","#f38ba8","#a6e3a1","#f9e2af","#cba6f7","#89dceb"]
        for i, (name, pts) in enumerate(hdata.items()):
            if not pts: continue
            xs = [datetime.datetime.fromtimestamp(int(pt["clock"])) for pt in pts]
            ys = []
            for pt in pts:
                try:    ys.append(float(pt["value"]))
                except: ys.append(0.0)
            ax.plot(xs, ys, color=clrs[i % len(clrs)],
                    linewidth=1.8, label=name[:35])
        ax.set_title(title, fontsize=9, color="#89b4fa", pad=5)
        ax.tick_params(colors="#a6adc8", labelsize=7)
        ax.spines[["top","right"]].set_visible(False)
        for sp in ["left","bottom"]: ax.spines[sp].set_color("#45475a")
        ax.grid(axis="y", linestyle="--", alpha=0.3, color="#45475a")
        if len(hdata) > 1:
            ax.legend(fontsize=7, facecolor="#313244",
                      edgecolor="#45475a", labelcolor="#cdd6f4")
        cv = FigureCanvasTkAgg(fig, master=self._chart_frame)
        cv.draw(); cv.get_tk_widget().pack(fill="both", expand=True)
        plt.close(fig)
        self._busy(False)
        self._info(f"График построен  —  точек: {sum(len(v) for v in hdata.values())}")

    # ── Экспорт PDF ────────────────────────────────────────────────────────────
    def _export_pdf(self):
        if not (self._problems or self._hosts or self._events):
            messagebox.showwarning("Нет данных","Загрузите данные"); return
        secs = self._report_settings_dialog("PDF")
        if secs is None: return
        path = filedialog.asksaveasfilename(
            defaultextension=".pdf", filetypes=[("PDF","*.pdf")],
            initialfile=f"zabbix_{datetime.date.today()}.pdf")
        if not path: return
        self._busy(True)
        def _w():
            try:
                PDFReporter(path).build(self._problems, self._hosts,
                                        self._events, self._metrics or None,
                                        sections=secs,
                                        hosts_detail=self._hosts_detail or None,
                                        auth_data=self._auth_data or None,
                                        user_dirs=self._user_dirs or None)
                self.after(0, lambda: (self._busy(False),
                    self._info(f"PDF: {path}"),
                    messagebox.showinfo("Готово", f"PDF сохранён:\n{path}")))
            except Exception as ex:
                msg = str(ex)
                self.after(0, lambda m=msg: (self._busy(False),
                    messagebox.showerror("Ошибка PDF", m)))
        threading.Thread(target=_w, daemon=True).start()

    # ── Экспорт Excel ─────────────────────────────────────────────────────────
    def _export_excel(self):
        if not (self._problems or self._hosts or self._events):
            messagebox.showwarning("Нет данных","Загрузите данные"); return
        secs = self._report_settings_dialog("Excel")
        if secs is None: return
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx", filetypes=[("Excel","*.xlsx")],
            initialfile=f"zabbix_{datetime.date.today()}.xlsx")
        if not path: return
        self._busy(True)
        def _w():
            try:
                ExcelReporter(path).build(self._problems, self._hosts,
                                          self._events, self._metrics or None,
                                          sections=secs,
                                          hosts_detail=self._hosts_detail or None,
                                          auth_data=self._auth_data or None,
                                          user_dirs=self._user_dirs or None)
                self.after(0, lambda: (self._busy(False),
                    self._info(f"Excel: {path}"),
                    messagebox.showinfo("Готово", f"Excel сохранён:\n{path}")))
            except Exception as ex:
                msg = str(ex)
                self.after(0, lambda m=msg: (self._busy(False),
                    messagebox.showerror("Ошибка Excel", m)))
        threading.Thread(target=_w, daemon=True).start()

    # ── Диалог настроек отчёта ───────────────────────────────────────────────
    def _report_settings_dialog(self, fmt: str):
        """Диалог выбора разделов. Возвращает set или None."""
        saved = self._cfg.get("report_sections",
                               {"problems": True, "hosts": True,
                                "events": True, "hosts_detail": True,
                                "auth": True, "metrics": True})
        BG = "#1e1e2e"; ACC = "#89b4fa"; FG = "#cdd6f4"; DIM = "#6c7086"

        dlg = tk.Toplevel(self)
        dlg.title(f"Настройки отчёта — {fmt}")
        dlg.configure(bg=BG)
        dlg.resizable(True, True)
        dlg.minsize(320, 320)
        dlg.grab_set()
        self.update_idletasks()
        x = self.winfo_x() + (self.winfo_width()  - 370) // 2
        y = self.winfo_y() + (self.winfo_height() - 460) // 2
        dlg.geometry(f"370x460+{x}+{y}")

        # ── Шапка (фиксированная) ─────────────────────────────────────────────
        hdr = tk.Frame(dlg, bg=BG)
        hdr.pack(fill="x", padx=20, pady=(14, 0))
        tk.Label(hdr, text=f"Разделы отчёта — {fmt}",
                 bg=BG, fg=ACC,
                 font=("Segoe UI", 11, "bold")).pack(anchor="w")
        tk.Label(hdr, text="Выбор сохраняется автоматически.",
                 bg=BG, fg=DIM,
                 font=("Segoe UI", 8)).pack(anchor="w")
        ttk.Separator(dlg).pack(fill="x", padx=12, pady=(8, 0))

        # ── Прокручиваемая область с чекбоксами ──────────────────────────────
        scroll_area = tk.Frame(dlg, bg=BG)
        scroll_area.pack(fill="both", expand=True, padx=0, pady=0)

        cv = tk.Canvas(scroll_area, bg=BG, highlightthickness=0)
        vsb = ttk.Scrollbar(scroll_area, orient="vertical", command=cv.yview)
        cv.configure(yscrollcommand=vsb.set)
        cv.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        inner = tk.Frame(cv, bg=BG)
        win_id = cv.create_window((0, 0), window=inner, anchor="nw")

        def _on_inner_cfg(e):
            cv.configure(scrollregion=cv.bbox("all"))
        def _on_cv_resize(e):
            cv.itemconfig(win_id, width=e.width)
        inner.bind("<Configure>", _on_inner_cfg)
        cv.bind("<Configure>", _on_cv_resize)
        cv.bind_all("<MouseWheel>",
                    lambda e: cv.yview_scroll(int(-1*(e.delta/120)), "units"))

        SECTIONS = [
            ("problems",     "🔴  Активные проблемы",    "Проблемы из вкладки «Проблемы»"),
            ("hosts",        "🖥  Статус хостов",         "Сводная таблица хостов"),
            ("events",       "📋  История событий",       "События из вкладки «События»"),
            ("hosts_detail", "🗂  Объекты мониторинга",   "Шаблоны, интерфейсы, теги, макросы"),
            ("auth",         "🔐  Аутентификация / LDAP", "Глобальные настройки + все LDAP/SAML серверы с JIT"),
            ("metrics",      "📈  Графики метрик",        "Загруженные графики из вкладки «Метрики»"),
        ]
        vars_ = {}
        for key, label, hint in SECTIONS:
            v = tk.BooleanVar(value=saved.get(key, True))
            vars_[key] = v
            row_f = tk.Frame(inner, bg=BG)
            row_f.pack(fill="x", padx=16, pady=(4, 0))
            tk.Checkbutton(
                row_f, text=label, variable=v,
                bg=BG, fg=FG, selectcolor="#313244",
                activebackground=BG, activeforeground=ACC,
                font=("Segoe UI", 10), anchor="w"
            ).pack(anchor="w")
            tk.Label(row_f, text=f"   {hint}",
                     bg=BG, fg=DIM,
                     font=("Segoe UI", 8)).pack(anchor="w")

        ttk.Separator(dlg).pack(fill="x", padx=12, pady=(6, 0))

        # ── Кнопки (фиксированные внизу) ─────────────────────────────────────
        bf = tk.Frame(dlg, bg=BG)
        bf.pack(fill="x", padx=16, pady=10)

        result = [None]

        def _ok():
            secs = {k for k, v in vars_.items() if v.get()}
            if not secs:
                messagebox.showwarning("Пусто",
                    "Выберите хотя бы один раздел.", parent=dlg)
                return
            self._cfg["report_sections"] = {k: v.get() for k, v in vars_.items()}
            _cfg_save(self._cfg)
            result[0] = secs
            dlg.destroy()

        def _cancel():
            dlg.destroy()

        ttk.Button(bf, text="✅  Создать отчёт",
                   command=_ok, width=18).pack(side="left", padx=(0, 8))
        ttk.Button(bf, text="Отмена",
                   command=_cancel, width=10).pack(side="left")

        dlg.wait_window()
        return result[0]

    # ── Экспорт CSV ───────────────────────────────────────────────────────────
    def _export_csv(self):
        if not (self._problems or self._hosts or self._events):
            messagebox.showwarning("Нет данных","Загрузите данные"); return
        folder = filedialog.askdirectory(title="Папка для CSV")
        if not folder: return
        ts = datetime.date.today().isoformat()
        def wcsv(name, hdrs, rows):
            with open(os.path.join(folder,name),"w",newline="",
                      encoding="utf-8-sig") as f:
                csv.writer(f).writerows([hdrs]+rows)
        wcsv(f"problems_{ts}.csv",
             ["Severity","Problem","Host","Start","Duration","Acks"],
             [(SEVERITY_NAMES.get(str(p.get("severity","0")),""),
               p.get("name",""),
               (p.get("hosts",[{}])[0].get("host","") if p.get("hosts") else ""),
               ts2str(p.get("clock",0)), dur_str(p.get("clock",0)),
               len(p.get("acknowledges",[]))) for p in self._problems])
        wcsv(f"hosts_{ts}.csv",
             ["Host","Name","Group","Status","Availability","IP","Triggers"],
             [(h.get("host",""), h.get("name",""),
               ", ".join(g["name"] for g in h.get("groups",[])),
               HOST_STATUS.get(str(h.get("status","0")),""),
               HOST_AVAIL.get(str(h.get("available","0")),""),
               (h.get("interfaces",[{}])[0].get("ip","") if h.get("interfaces") else ""),
               h.get("triggers","0")) for h in self._hosts])
        wcsv(f"events_{ts}.csv",
             ["EventID","Severity","Name","Host","Time","Resolved"],
             [(e.get("eventid",""),
               SEVERITY_NAMES.get(str(e.get("severity","0")),""),
               e.get("name",""),
               (e.get("hosts",[{}])[0].get("host","") if e.get("hosts") else ""),
               ts2str(e.get("clock",0)),
               (ts2str(e.get("r_clock","0"))
                if e.get("r_clock","0") not in ("","0") else ""))
              for e in self._events])
        for title, hdata in self._metrics.items():
            safe = title[:40].replace("/","-").replace("\\","-")
            wcsv(f"metrics_{safe}_{ts}.csv",
                 ["Metric","Clock","Value"],
                 [(n, pt["clock"], pt["value"])
                  for n, pts in hdata.items() for pt in pts])
        self._info(f"CSV сохранены в {folder}")
        messagebox.showinfo("Готово", f"CSV-файлы сохранены:\n{folder}")

    # ── Утилиты ───────────────────────────────────────────────────────────────
    # ── Лог ───────────────────────────────────────────────────────────────────
    def log(self, msg: str, level: str = "info"):
        """Записать в буфер и, если окно открыто, отобразить там."""
        def _do():
            ts = datetime.datetime.now().strftime("%H:%M:%S")
            self._log_buf.insert("end", f"[{ts}] ", "ts")
            self._log_buf.insert("end", msg + "\n", level)
            if self._log_win and self._log_win.winfo_exists():
                self._log_win.append(f"[{ts}] ", "ts")
                self._log_win.append(msg + "\n", level)
            if level == "err":
                self._btn_log_open.configure(fg="#f38ba8")
        self.after(0, _do)

    def _open_log_window(self):
        """Открыть окно лога или поднять поверх если уже открыто."""
        if self._log_win and self._log_win.winfo_exists():
            self._log_win.lift(); self._log_win.focus_force()
        else:
            self._log_win = LogWindow(self)
        self._btn_log_open.configure(fg="#89b4fa")

    def _clear_log(self):
        self._log_buf.delete("1.0", "end")
        if self._log_win and self._log_win.winfo_exists():
            self._log_win.clear()

    def _save_log(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".txt", filetypes=[("Text","*.txt")],
            initialfile=f"zabbix_log_{datetime.date.today()}.txt")
        if path:
            with open(path, "w", encoding="utf-8") as f:
                f.write(self._log_buf.get("1.0","end"))
            self._info(f"Лог сохранён: {path}")

    # ── Конфиг ────────────────────────────────────────────────────────────────
    def _cfg_to_ui(self):
        """Заполнить поля из config.json."""
        c = self._cfg
        if c.get("url"):
            self._e_url.delete(0, "end"); self._e_url.insert(0, c["url"])
        if c.get("user"):
            self._e_user.delete(0, "end"); self._e_user.insert(0, c["user"])
        if c.get("password"):
            self._e_pass.delete(0, "end"); self._e_pass.insert(0, c["password"])
        self._v_ssl.set(c.get("verify_ssl", True))
        if c.get("severity") is not None:
            try: self._cmb_sev.current(int(c["severity"]))
            except: pass
        if c.get("limit"):
            self._v_limit.set(c["limit"])
        if c.get("date_from"):
            self._e_from.delete(0,"end"); self._e_from.insert(0, c["date_from"])
        if c.get("date_till"):
            self._e_till.delete(0,"end"); self._e_till.insert(0, c["date_till"])

    def _cfg_save_ui(self):
        """Сохранить текущие настройки в config.json."""
        _cfg_save({
            "url":        self._e_url.get().strip(),
            "user":       self._e_user.get().strip(),
            "password":   self._e_pass.get(),
            "verify_ssl": self._v_ssl.get(),
            "severity":   self._cmb_sev.current(),
            "limit":      self._v_limit.get(),
            "date_from":  self._e_from.get().strip(),
            "date_till":  self._e_till.get().strip(),
        })
        self._info("Настройки сохранены в config.json")

    def _on_close(self):
        self._cfg_save_ui()
        if self.zapi:
            try: self.zapi.logout()
            except: pass
        self.destroy()

    def _run_diagnostics(self):
        if not self.zapi:
            messagebox.showwarning("", "Сначала подключитесь"); return
        self.log("━━━ ДИАГНОСТИКА API ━━━", "info")
        self._busy(True)
        def _w():
            try:
                d = self.zapi.diagnose()

                # Пользователь
                u = d.get("current_user", {})
                if u:
                    role = u.get("role", {})
                    grps = u.get("usrgrps", [])
                    self.after(0, lambda: self.log(
                        f"  Пользователь: {u.get('username','?')}  "
                        f"role={role.get('name','?')} (type={role.get('type','?')})",
                        "info"))
                    self.after(0, lambda gs=grps: self.log(
                        f"  Группы польз.: {[g['name'] for g in gs]}", "info"))
                else:
                    err = d.get("current_user_error","?")
                    self.after(0, lambda e=err: self.log(f"  user.get ошибка: {e}", "err"))

                # Группы хостов
                hg = d.get("host_groups", [])
                hge = d.get("host_groups_error", "")
                if hge:
                    self.after(0, lambda e=hge: self.log(f"  hostgroup.get ошибка: {e}", "err"))
                else:
                    self.after(0, lambda n=len(hg), names=[g['name'] for g in hg[:10]]:
                        self.log(f"  Видимые группы хостов ({n}): {names}", "ok" if n else "warn"))

                # Счётчик хостов
                cnt = d.get("hosts_total_count", "?")
                cnt_err = d.get("hosts_count_error", "")
                if cnt_err:
                    self.after(0, lambda e=cnt_err: self.log(f"  host.get count ошибка: {e}", "err"))
                else:
                    self.after(0, lambda c=cnt: self.log(
                        f"  Всего хостов (countOutput): {c}",
                        "ok" if str(c) != "0" else "warn"))

                # Примеры хостов
                hs = d.get("hosts_sample", [])
                hs_err = d.get("hosts_raw_error", "")
                if hs_err:
                    self.after(0, lambda e=hs_err: self.log(f"  host.get sample ошибка: {e}", "err"))
                elif hs:
                    for hh in hs[:5]:
                        self.after(0, lambda h=hh: self.log(
                            f"    Хост: id={h.get('hostid')}  "
                            f"host={h.get('host')}  status={h.get('status')}", "ok"))
                else:
                    self.after(0, lambda: self.log(
                        "  ⚠ host.get вернул 0 хостов даже без фильтров!", "warn"))
                    self.after(0, lambda: self.log(
                        "  Возможные причины:", "warn"))
                    self.after(0, lambda: self.log(
                        "  1. Пользователь не добавлен в группу с доступом к хостам", "warn"))
                    self.after(0, lambda: self.log(
                        "     → Zabbix UI: Administration > User groups > [группа] > Host group permissions", "warn"))
                    self.after(0, lambda: self.log(
                        "  2. Все хосты отключены (status=1)", "warn"))
                    self.after(0, lambda: self.log(
                        "  3. Используйте встроенного Admin вместо reports для теста", "warn"))

                # Проблемы без фильтра
                ps = d.get("problems_sample", [])
                ps_err = d.get("problems_raw_error", "")
                if ps_err:
                    self.after(0, lambda e=ps_err: self.log(f"  problem.get ошибка: {e}", "err"))
                else:
                    self.after(0, lambda n=len(ps): self.log(
                        f"  problem.get (без фильтра, recent=True): {n} записей",
                        "ok" if n else "warn"))

                self.after(0, lambda: self.log("━━━ ДИАГНОСТИКА ЗАВЕРШЕНА ━━━", "info"))
            except Exception as ex:
                msg = str(ex)
                self.after(0, lambda m=msg: self.log(f"  Ошибка диагностики: {m}", "err"))
            finally:
                self.after(0, lambda: self._busy(False))
        threading.Thread(target=_w, daemon=True).start()


    # ── Объекты мониторинга ───────────────────────────────────────────────────
    def _obj_load(self):
        """Загрузить расширенные данные хостов с фильтрами."""
        if not self.zapi:
            messagebox.showwarning("", "Сначала подключитесь"); return
        # Определить фильтры
        grp_sel = self._obj_grp_var.get()
        tpl_sel = self._obj_tpl_var.get()
        groupids   = None
        templateids= None
        if grp_sel and grp_sel != "— все —":
            gid = self._obj_grp_map.get(grp_sel)
            if gid: groupids = [gid]
        if tpl_sel and tpl_sel != "— все —":
            tid = self._obj_tpl_map.get(tpl_sel)
            if tid: templateids = [tid]

        self._busy(True)
        self.log(f"  [Объекты] Загрузка: group={grp_sel!r} tpl={tpl_sel!r}", "req")

        def _w():
            try:
                hosts = self.zapi.get_hosts_detail(
                    groupids=groupids, templateids=templateids)
                self.after(0, lambda: self._obj_populate(hosts))
            except Exception as ex:
                msg = str(ex)
                self.after(0, lambda m=msg: (
                    self._busy(False),
                    self.log(f"  [Объекты] Ошибка: {m}", "err"),
                    self._info(f"Ошибка: {m}")))
        threading.Thread(target=_w, daemon=True).start()

    def _obj_init_filters(self, hosts):
        """Заполнить выпадающие списки групп и шаблонов из загруженных хостов."""
        groups    = {}
        templates = {}
        for h in hosts:
            for g in h.get("groups", []):
                groups[g["name"]] = g["groupid"]
            for t in h.get("parentTemplates", []):
                templates[t["name"]] = t["templateid"]
        self._obj_grp_map = groups
        self._obj_tpl_map = templates
        grp_vals = ["— все —"] + sorted(groups.keys())
        tpl_vals = ["— все —"] + sorted(templates.keys())
        self._cmb_obj_grp["values"] = grp_vals
        self._cmb_obj_tpl["values"] = tpl_vals
        if self._obj_grp_var.get() not in grp_vals:
            self._obj_grp_var.set("— все —")
        if self._obj_tpl_var.get() not in tpl_vals:
            self._obj_tpl_var.set("— все —")

    def _obj_populate(self, hosts):
        self._hosts_detail = hosts
        self._obj_init_filters(hosts)
        self._obj_apply_filter()
        self._busy(False)
        self.log(f"  [Объекты] Загружено: {len(hosts)} хостов", "ok")
        self._info(f"Объекты: {len(hosts)} хостов")

    def _obj_apply_filter(self):
        """Применить текстовый фильтр и показать в дереве."""
        t = self._tree_obj
        t.delete(*t.get_children())
        t._all = []
        q        = self._obj_search_var.get().lower()
        grp_sel  = self._obj_grp_var.get()
        tpl_sel  = self._obj_tpl_var.get()

        for h in self._hosts_detail:
            # Фильтр по группе
            if grp_sel and grp_sel != "— все —":
                if not any(g["name"] == grp_sel for g in h.get("groups",[])):
                    continue
            # Фильтр по шаблону
            if tpl_sel and tpl_sel != "— все —":
                if not any(tp["name"] == tpl_sel
                            for tp in h.get("parentTemplates",[])):
                    continue
            # Текстовый поиск
            host_str = " ".join([
                h.get("host",""), h.get("name",""),
                ", ".join(g["name"] for g in h.get("groups",[])),
            ]).lower()
            if q and q not in host_str:
                continue

            # Определить основной IP
            ifaces = h.get("interfaces",[])
            main_ip = ""
            for iface in ifaces:
                if iface.get("main") == "1":
                    main_ip = iface.get("ip","") or iface.get("dns","")
                    break
            if not main_ip and ifaces:
                main_ip = ifaces[0].get("ip","") or ifaces[0].get("dns","")

            status = "Вкл" if str(h.get("status","0")) == "0" else "Выкл"
            grp_str = ", ".join(g["name"] for g in h.get("groups",[]))
            avail   = HOST_AVAIL.get(str(h.get("available","0")), "?")

            row = (h.get("host",""), h.get("name",""),
                   grp_str, status, avail, main_ip)

            ac  = str(h.get("available","0"))
            dis = str(h.get("status","0")) != "0"
            tag = "dis" if dis else ("ok" if ac=="1" else
                                     "bad" if ac=="2" else "unk")
            iid = t.insert("","end", values=row, tags=(tag,),
                            iid=h["hostid"])
            t._all.append((iid, row))

        # Сбросить карточку
        self._obj_detail.configure(state="normal")
        self._obj_detail.delete("1.0","end")
        self._obj_detail.configure(state="disabled")

    def _obj_on_select(self, event=None):
        """Показать карточку выбранного хоста."""
        sel = self._tree_obj.selection()
        if not sel: return
        hostid = sel[0]
        h = next((x for x in self._hosts_detail
                   if x["hostid"] == hostid), None)
        if not h: return
        self._obj_show_card(h)

    def _obj_show_card(self, h):
        txt = self._obj_detail
        txt.configure(state="normal")
        txt.delete("1.0","end")

        def row(key, val, val_tag="val"):
            txt.insert("end", f"  {key:<22}", "key")
            txt.insert("end", f"{val}\n", val_tag)

        def section(title):
            txt.insert("end", f"\n● {title}\n", "hdr")

        # ── Основная информация ───────────────────────────────────────────────
        section("Основное")
        row("Host name",    h.get("host",""))
        row("Visible name", h.get("name","") or h.get("host",""))
        status_val = "✅ Включён" if str(h.get("status","0"))=="0" else "❌ Отключён"
        status_tag = "ok" if str(h.get("status","0"))=="0" else "bad"
        row("Enabled",      status_val, status_tag)
        row("Monitored by", h.get("_proxy_name","Zabbix Server"))
        if h.get("description"):
            row("Description", h["description"][:120])

        # ── Группы хостов ─────────────────────────────────────────────────────
        section("Host groups")
        for g in h.get("groups",[]):
            txt.insert("end", f"  • {g['name']}\n", "val")

        # ── Шаблоны ───────────────────────────────────────────────────────────
        section("Templates")
        tpls = h.get("parentTemplates",[])
        if tpls:
            for tp in tpls:
                txt.insert("end", f"  • {tp.get('name','?')}\n", "val")
        else:
            txt.insert("end","  (нет шаблонов)\n","dim")

        # ── Интерфейсы ────────────────────────────────────────────────────────
        ITYPE = {"1":"Agent","2":"SNMP","3":"IPMI","4":"JMX"}
        section("Interfaces")
        ifaces = h.get("interfaces",[])
        if ifaces:
            for iface in ifaces:
                itype = ITYPE.get(str(iface.get("type","1")),"?")
                addr  = iface.get("ip","") or iface.get("dns","")
                port  = iface.get("port","")
                main  = " [основной]" if iface.get("main")=="1" else ""
                txt.insert("end",
                    f"  • {itype:<6}  {addr}:{port}{main}\n","val")
                # Agent version
                if str(iface.get("type","1")) == "1":
                    av = h.get("_agent_version","")
                    txt.insert("end",
                        f"  {'Zabbix agent ver':<22}", "key")
                    txt.insert("end",
                        f"{av if av else '(не получен)'}\n",
                        "val" if av else "dim")
        else:
            txt.insert("end","  (нет интерфейсов)\n","dim")

        # ── Теги ──────────────────────────────────────────────────────────────
        section("Tags")
        tags = h.get("tags",[])
        if tags:
            for tg in tags:
                v = tg.get("value","")
                txt.insert("end",
                    f"  • {tg.get('tag','')}",  "key")
                txt.insert("end",
                    f"{': '+v if v else ''}\n", "val")
        else:
            txt.insert("end","  (нет тегов)\n","dim")

        # ── Макросы ───────────────────────────────────────────────────────────
        section("Host macros")
        macros = h.get("macros",[])
        if macros:
            MTYPE = {"0":"Text","1":"Secret","2":"Vault"}
            for m in macros:
                mtype = MTYPE.get(str(m.get("type","0")),"?")
                val   = ("***" if str(m.get("type","0"))=="1"
                         else m.get("value",""))
                desc  = m.get("description","")
                txt.insert("end", f"  {m.get('macro',''):<30}","key")
                txt.insert("end", f"= {val}", "val")
                if desc:
                    txt.insert("end", f"  # {desc[:50]}","dim")
                txt.insert("end","\n")
        else:
            txt.insert("end","  (нет макросов)\n","dim")

        txt.configure(state="disabled")

    def _filter(self, t, q):
        t.delete(*t.get_children()); q = q.lower()
        for iid, row in t._all:
            if not q or any(q in str(v).lower() for v in row):
                t.insert("","end", values=row)

    def _sort(self, t, col):
        rows = [(t.set(k,col),k) for k in t.get_children()]
        for i, (_,k) in enumerate(sorted(rows)): t.move(k,"",i)

    def _busy(self, on):
        if on: self._prog.start(12)
        else:  self._prog.stop()

    def _info(self, text): self._lbl_info.configure(text=text)


# ══════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    App().mainloop()
