"""
Zabbix Reporter v2.0  —  Windows 10/11
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
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    MPL_OK = True
except ImportError:
    MPL_OK = False

try:
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    MPL_TK_OK = True
except ImportError:
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

    def get_events(self, time_from, time_till, limit=1000):
        """
        Zabbix 6.0+: event.get не поддерживает selectHosts напрямую.
        Имена хостов резолвим через trigger.get по objectid.
        """
        events = self._call("event.get", {
            "output": "extend",
            "time_from": time_from, "time_till": time_till,
            "sortfield": ["clock", "eventid"], "sortorder": "DESC",
            "limit": limit, "value": 1,
            "source": 0,   # только триггерные события
        })
        if not events:
            return events

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
        fig, ax = plt.subplots(figsize=(7, 2.8))
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
        plt.close(fig)
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
        canvas.drawString(15*mm, h - 14*mm, "ZABBIX REPORTER v2.0")
        canvas.setFont(_PDF_FONT, 9)
        canvas.drawRightString(w - 15*mm, h - 14*mm,
                               datetime.datetime.now().strftime("%Y-%m-%d %H:%M"))
        canvas.setFillColor(HexColor("#0F3460"))
        canvas.rect(0, 0, w, 10*mm, fill=1, stroke=0)
        canvas.setFillColor(colors.white)
        canvas.setFont(_PDF_FONT, 8)
        canvas.drawCentredString(w/2, 3.5*mm,
                                 f"Страница {doc.page}  |  Zabbix Reporter v2.0")
        canvas.restoreState()

    def build(self, problems, hosts, events, metrics=None, sections=None, hosts_detail=None):
        if sections is None: sections = {'problems','hosts','events','metrics','hosts_detail'}
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

    def build(self, problems, hosts, events, metrics=None, sections=None, hosts_detail=None):
        if sections is None: sections = {'problems','hosts','events','metrics','hosts_detail'}
        if hosts_detail is None: hosts_detail = []
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
        self.title("⚡ Zabbix Reporter v2.0")
        self.geometry("1300x840")
        self.minsize(980, 660)
        self.configure(bg="#1e1e2e")

        self.zapi          = None
        self._problems     = []
        self._hosts        = []
        self._events       = []
        self._metrics      = {}
        self._items_cache  = []
        self._hosts_detail = []   # расширенные данные вкладки «Объекты»
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
                     padding=[14,6], font=("Segoe UI",10))
        s.map("TNotebook.Tab",
              background=[("selected",ACC)], foreground=[("selected","#1e1e2e")])
        s.configure("Treeview",     background=ENT, foreground=FG,
                     fieldbackground=ENT, rowheight=26, font=("Segoe UI",9))
        s.configure("Treeview.Heading", background=SEL, foreground=ACC,
                     font=("Segoe UI",9,"bold"))
        s.map("Treeview",
              background=[("selected",ACC)], foreground=[("selected","#1e1e2e")])
        s.configure("TEntry",       fieldbackground=ENT, foreground=FG)
        s.configure("TCombobox",    fieldbackground=ENT, foreground=FG)
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
        tk.Label(hdr, text="⚡ Zabbix Reporter  v2.0",
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

    # ── Вкладки ───────────────────────────────────────────────────────────────
    def _build_tabs(self, parent):
        nb = ttk.Notebook(parent); nb.pack(fill="both", expand=True)

        t1 = ttk.Frame(nb); nb.add(t1, text=" 🔴 Проблемы ")
        self._tree_p = self._tv(t1,
            ("Серьёзность","Проблема","Хост","Начало","Длит.","Подтв."),
            (118,360,160,152,108,68))

        t2 = ttk.Frame(nb); nb.add(t2, text=" 🖥 Хосты ")
        self._tree_h = self._tv(t2,
            ("Хост","Имя","Группа","Статус","Доступн.","Интерфейсы","Тригг.","Обслуж."),
            (140,160,140,80,100,200,55,150))

        t3 = ttk.Frame(nb); nb.add(t3, text=" 📋 События ")
        self._tree_e = self._tv(t3,
            ("ID","Серьёзность","Событие","Хост","Время","Восст."),
            (80,118,340,155,152,130))

        t4 = ttk.Frame(nb); nb.add(t4, text=" 📈 Метрики ")
        self._build_metrics_tab(t4)

        t6 = ttk.Frame(nb); nb.add(t6, text=" 🗂 Объекты ")
        self._build_objects_tab(t6)

        t5 = ttk.Frame(nb); nb.add(t5, text=" 📊 Сводка ")
        self._sum_txt = tk.Text(t5, bg="#181825", fg="#cdd6f4",
                                 font=("Consolas",11), relief="flat",
                                 state="disabled", padx=20, pady=12)
        sb = ttk.Scrollbar(t5, command=self._sum_txt.yview)
        self._sum_txt.configure(yscrollcommand=sb.set)
        self._sum_txt.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

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
        ttk.Button(top, text="🔍 Найти",        command=self._load_items).pack(side="left", padx=4)
        ttk.Button(top, text="📈 История",      command=self._load_history).pack(side="left", padx=2)

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

                self.after(0, lambda: self._populate(p, h, ev))
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
        # Обновить фильтры вкладки Объекты на основе базовых данных хостов
        # (полная загрузка по кнопке «Загрузить» во вкладке)
        try:
            fake_groups    = {}
            fake_templates = {}
            for hh in h:
                # Support both "groups" and "hostgroups" field names
                grp_list = hh.get("groups", hh.get("hostgroups", []))
                for g in grp_list:
                    fake_groups[g["name"]] = g.get("groupid","")
            self._obj_grp_map = fake_groups
            self._obj_tpl_map = fake_templates
            grp_vals = ["— все —"] + sorted(fake_groups.keys())
            self._cmb_obj_grp["values"] = grp_vals
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
        t = self._tree_h; t.delete(*t.get_children()); t._all = []
        t.tag_configure("ok",    foreground="#a6e3a1")
        t.tag_configure("bad",   foreground="#f38ba8")
        t.tag_configure("unk",   foreground="#f9e2af")
        t.tag_configure("maint", foreground="#fab387")
        for h in data:
            ifaces = h.get("interfaces", [])
            # Доступность вычисляется из интерфейсов (Zabbix 6.0+)
            avail_str, avail_color, ifaces_str = _iface_avail_str(ifaces)
            # Для обратной совместимости: если интерфейсов нет — fallback на host.available
            if not ifaces:
                ac = str(h.get("available", "0"))
                avail_str = HOST_AVAIL.get(ac, "Unknown")
                avail_color = IFACE_AVAIL_COLOR.get(
                    "1" if ac=="1" else "2" if ac=="2" else "0", "#f9e2af")
            grps = h.get("groups", h.get("hostgroups", []))
            # Maintenance
            in_maint  = h.get("maintenance_status", "0") == "1"
            maint_str = ""
            if in_maint:
                mname     = h.get("_maint_name", "")
                maint_str = f"🔧 {mname}" if mname else "🔧 Обслуживание"

            row = (h.get("host",""), h.get("name",""),
                   ", ".join(g["name"] for g in grps),
                   HOST_STATUS.get(str(h.get("status","0")),""),
                   avail_str,
                   ifaces_str,
                   h.get("triggers","0"),
                   maint_str)

            tag = ("maint" if in_maint else
                   "ok"    if avail_str == "Available" else
                   "bad"   if avail_str == "Unavailable" else "unk")
            iid = t.insert("","end", values=row, tags=(tag,))
            t._all.append((iid, row))

    def _fill_e(self, data):
        t = self._tree_e; t.delete(*t.get_children()); t._all = []
        for e in data:
            sid = str(e.get("severity","0")); rc = e.get("r_clock","0")
            eh  = e.get("hosts",[{}])
            row = (e.get("eventid",""), SEVERITY_NAMES.get(sid,sid),
                   e.get("name",""),
                   eh[0].get("host","") if eh else "",
                   ts2str(e.get("clock",0)),
                   ts2str(rc) if rc not in ("","0") else "–")
            iid = t.insert("","end",values=row,tags=(f"s{sid}",))
            t.tag_configure(f"s{sid}", foreground=SEVERITY_HEX.get(sid,"#cdd6f4"))
            t._all.append((iid,row))

    def _fill_sum(self):
        p, h, ev = self._problems, self._hosts, self._events
        sc: dict = {}
        for x in p: sc[str(x.get("severity","0"))] = sc.get(str(x.get("severity","0")),0)+1
        av = {"Available":0,"Unavailable":0,"Unknown":0}
        for x in h: av[HOST_AVAIL.get(str(x.get("available","0")),"Unknown")] += 1
        lines = ["═"*62,"  СВОДНЫЙ ОТЧЁТ — Zabbix Reporter v2.0",
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
                                        hosts_detail=self._hosts_detail or None)
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
                                          hosts_detail=self._hosts_detail or None)
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
                                "metrics": True})
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
