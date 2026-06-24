#!/usr/bin/env python3
"""
generar_dashboards.py — ZAS Talents Dashboard Generator
Genera 6 HTMLs (ARG, CHI, COL, USA, PER, INT) desde Odoo via JSON-RPC.
Diseño: replica exacta del HTML aprobado (finalsahsboardok.html).
"""

import json, os, sys, math, urllib.request, urllib.parse
from datetime import datetime, date, timezone
from collections import defaultdict
import openpyxl

# ─── CONFIG ────────────────────────────────────────────────────────────────────
ODOO_URL  = os.environ.get("ODOO_URL",  "https://zas-talent.odoo.com")
ODOO_DB   = os.environ.get("ODOO_DB",   "zas-talent")
ODOO_USER = os.environ.get("ODOO_USER", "martina.boazzo@zastalents.com")
ODOO_PASS = os.environ.get("ODOO_PASSWORD", "")

BUDGET_FILE = os.environ.get("BUDGET_FILE", "Base_Regional__2_.xlsx")
OUTPUT_DIR  = os.environ.get("OUTPUT_DIR", "docs")

# Tipo de cambio (fallback si no se puede obtener online)
TC = {
    "ARS": 1100,   # BCRA
    "CLP": 940,    # BCCh (presupuesto usa 890)
    "COP": 4250,   # BanRep
    "PEN": 3.72,   # BCRP
    "USD": 1,
}
TC_PRESUP_CLP = 890   # TC especial para presupuesto Chile

# ─── CAMPOS ODOO ───────────────────────────────────────────────────────────────
# ─── CAMPOS CONFIRMADOS (desde Campos__ir_model_fields_.xlsx) ──────────────────
# Campos en project.task
SUBTASK_FIELDS = [
    "id", "name", "stage_id", "date_deadline",
    "company_id", "sale_order_id", "project_id",
    # Pais de campaña (campo relacionado al sale_order en la tarea)
    "x_studio_related_field_8rl_1jhbqu80b",   # Pais de campaña
    "x_studio_related_field_8pi_1jhbqv0jf",   # Pais de campaña (copia)
    # Fecha de publicación = criterio de implementado
    "x_studio_fecha_de_publicacin",
    # Responsable / Implementador
    "x_studio_related_field_7v0_1jidluau0",
    # Fecha estimada (para pipeline)
    "x_studio_fecha_limite_ops",
]

# Campos en sale.order (se descargan por separado)
ORDER_FIELDS = [
    "id", "name", "amount_untaxed", "currency_id",
    "state", "partner_id", "company_id", "date_order",
    # País de campaña en la orden
    "x_studio_campaas",    # Pais de campaña (campo 1)
    "x_studio_campaas_1",  # Pais de campaña (campo 2 — el que diferencia)
    "x_studio_bu_1",       # BU del talento (para internacionales)
    "x_studio_bu",         # BU alternativo
]

# FIELD_MAP vacío (ya no se usa discover_fields)
FIELD_MAP = {}

ORDER_FIELDS = [
    "id", "name", "amount_untaxed", "currency_id",
    "state", "partner_id", "company_id",
    "date_order",
]

# ─── MAPAS DE TRADUCCIÓN DE CAMPOS SELECTION ───────────────────────────────────
# x_studio_campaas (campo 1): valor_interno → label
SEL_CAMPAAS = {
    'Campañas':             'Campañas',
    'Campañas_Chile':       'Campañas Chile',
    'Campañas_Colombia':    'Campañas Colombia',
    'Campañas_US':          'Campañas US',
    'Campañas_Peru':        'Campañas Peru',
    'Campañas_Internacionales': 'Campañas Internacionales',
    # también pueden venir ya con el label completo
    'Campañas Chile':       'Campañas Chile',
    'Campañas Colombia':    'Campañas Colombia',
    'Campañas US':          'Campañas US',
    'Campañas Peru':        'Campañas Peru',
    'Campañas Internacionales': 'Campañas Internacionales',
}

# x_studio_campaas_1 (campo 2): valor_interno → label  (el que realmente diferencia)
SEL_CAMPAAS_1 = {
    'Campañas':             'Campañas Argentinas',
    'Campañas_Chile':       'Campañas Chile',
    'Campañas_Colombia':    'Campañas Colombia',
    'Campañas_US':          'Campañas US',
    'Campañas_Peru':        'Campañas Peru',
    'Campañas_Internacionales': 'Campañas Internacionales',
    # labels directos
    'Campañas Argentinas':  'Campañas Argentinas',
    'Campañas Chile':       'Campañas Chile',
    'Campañas Colombia':    'Campañas Colombia',
    'Campañas US':          'Campañas US',
    'Campañas Peru':        'Campañas Peru',
    'Campañas Internacionales': 'Campañas Internacionales',
}

# BU → país para registros internacionales
BU_TO_PAIS = {
    'ZAS ARGENTINA': 'argentina',
    'ZAS CHILE':     'chile',
    'ZAS COLOMBIA':  'colombia',
    'ZAS USA':       'usa',
    'ZAS PERU':      'peru',
    'ZAS PERU ':     'peru',
}

# Label de campaas_1 → país local
LABEL_TO_PAIS = {
    'Campañas Argentinas': 'argentina',
    'Campañas Chile':      'chile',
    'Campañas Colombia':   'colombia',
    'Campañas US':         'usa',
    'Campañas Peru':       'peru',
}

# company_id → país fallback
COMPANY_TO_PAIS = {1: 'argentina', 2: 'chile', 4: 'peru', 5: 'usa', 6: 'colombia'}

# ─── UTILIDADES ────────────────────────────────────────────────────────────────
def fmt_usd(v):
    """Formatea número a USD legible."""
    if v is None: return "—"
    v = float(v)
    if abs(v) >= 1_000_000: return f"USD {v/1_000_000:.1f}M"
    if abs(v) >= 1_000:     return f"USD {v/1_000:.1f}K"
    return f"USD {v:,.0f}"

def to_usd(amount, currency_code):
    """Convierte monto a USD."""
    if not amount: return 0.0
    code = (currency_code or "USD").upper()
    tc = TC.get(code, 1)
    return float(amount) / tc if tc else float(amount)

def parse_date(s):
    """Parsea fecha string a objeto date."""
    if not s: return None
    try:
        return datetime.strptime(s[:10], "%Y-%m-%d").date()
    except Exception:
        return None

def mes_label(y, m):
    meses = ["", "Jan", "Feb", "Mar", "Apr", "May", "Jun",
             "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    return f"{meses[m]} {str(y)[-2:]}"

def pct_color(pct):
    """Retorna clase CSS según porcentaje de presupuesto."""
    if pct > 1.0:   return "over"
    if pct >= 0.9:  return "ok"
    if pct >= 0.6:  return "warn"
    return "low"

def pct_css_color(pct):
    if pct > 1.0:   return "var(--a5)"
    if pct >= 0.9:  return "var(--a3)"
    if pct >= 0.6:  return "var(--yellow)"
    return "var(--red)"

# ─── ODOO JSON-RPC ─────────────────────────────────────────────────────────────
import http.cookiejar as _cj
_cookie_jar = _cj.CookieJar()
_opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(_cookie_jar))

def odoo_call(url, payload):
    data = json.dumps(payload).encode("utf-8")
    req  = urllib.request.Request(
        url,
        data=data,
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    with _opener.open(req, timeout=120) as r:
        return json.loads(r.read())

def odoo_auth():
    print("  Autenticando en Odoo...")
    # Odoo 19: usar /web/session/authenticate (no /web/dataset/call_kw)
    resp = odoo_call(f"{ODOO_URL}/web/session/authenticate", {
        "jsonrpc": "2.0", "method": "call", "id": 1,
        "params": {
            "db":       ODOO_DB,
            "login":    ODOO_USER,
            "password": ODOO_PASS,
        }
    })
    result = resp.get("result", {})
    uid = result.get("uid")
    if not uid:
        raise Exception(f"Auth fallida: {resp.get('error', result)}")
    print(f"  OK UID={uid} ({result.get('name', '')})")
    return uid

def odoo_search_read(uid, model, domain, fields, limit=5000, offset=0):
    ctx = {"allowed_company_ids": [1, 2, 3, 4, 5, 6]}
    resp = odoo_call(f"{ODOO_URL}/web/dataset/call_kw", {
        "jsonrpc": "2.0", "method": "call", "id": 2,
        "params": {
            "model":  model,
            "method": "search_read",
            "args":   [domain],
            "kwargs": {
                "fields":  fields,
                "limit":   limit,
                "offset":  offset,
                "context": ctx,
            }
        }
    })
    result = resp.get("result")
    if result is None:
        raise Exception(f"search_read error en {model}: {resp.get('error')}")
    return result

def fetch_all(uid, model, domain, fields, batch=2000):
    """Descarga todos los registros en batches."""
    all_recs = []
    offset = 0
    while True:
        batch_recs = odoo_search_read(uid, model, domain, fields, limit=batch, offset=offset)
        all_recs.extend(batch_recs)
        print(f"    {model}: {len(all_recs)} registros...", end="\r")
        if len(batch_recs) < batch:
            break
        offset += batch
    print(f"    {model}: {len(all_recs)} registros totales.       ")
    return all_recs

def _f(t, canonical):
    """Obtiene valor de un campo (alias para t.get, mantenido por compatibilidad)."""
    return t.get(canonical)


# ─── DESCARGA DE DATOS ─────────────────────────────────────────────────────────
def download_data(uid):
    print("  Descargando subtareas...")
    tasks = fetch_all(uid, "project.task",
        domain=[["sale_order_id", "!=", False]],
        fields=SUBTASK_FIELDS,
    )

    print("  Descargando órdenes de venta...")
    orders_raw = fetch_all(uid, "sale.order",
        domain=[["state", "in", ["sale", "done"]]],
        fields=ORDER_FIELDS,
    )

    # Índice de órdenes
    orders = {o["id"]: o for o in orders_raw}
    print(f"  ✓ {len(tasks)} subtareas, {len(orders)} órdenes activas")
    return tasks, orders

# ─── CLASIFICACIÓN ─────────────────────────────────────────────────────────────
def classify_tasks(tasks, orders):
    """
    Clasifica cada tarea en (pais, tab).
    Tab: 'local' o 'intl'
    Retorna: dict {pais: {'local': [...], 'intl': [...]}}
    """
    classified = {
        'argentina': {'local': [], 'intl': []},
        'chile':     {'local': [], 'intl': []},
        'colombia':  {'local': [], 'intl': []},
        'usa':       {'local': [], 'intl': []},
        'peru':      {'local': [], 'intl': []},
        'internacional': {'local': [], 'intl': []},
    }

    today = date.today()

    for t in tasks:
        # Excluir México
        cid = t.get("company_id")
        if isinstance(cid, list):
            cid = cid[0]
        if cid == 3:
            continue

        # Adjuntar orden de venta
        oid = t.get("sale_order_id")
        if isinstance(oid, list):
            oid = oid[0]
        order = orders.get(oid) if oid else None

        # Si la orden está cancelada (por si acaso llegó), skip
        if order and order.get("state") not in ("sale", "done"):
            continue

        t["_order"] = order

        # ── Traducir labels de selección ──
        # x_studio_campaas y x_studio_bu_1 están en sale.order, NO en project.task
        if order:
            raw1 = order.get("x_studio_campaas") or ""
            raw2 = order.get("x_studio_campaas_1") or raw1
        else:
            # fallback: campo relacionado copiado en la tarea
            raw1 = t.get("x_studio_related_field_8rl_1jhbqu80b") or ""
            raw2 = raw1
        label2 = SEL_CAMPAAS_1.get(raw2, raw2)
        if not label2:
            label2 = SEL_CAMPAAS_1.get(raw1, raw1)

        # ── Determinar si es internacional ──
        if label2 == "Campañas Internacionales" or "Internacional" in label2:
            # BU desde la orden de venta
            if order:
                bu = (order.get("x_studio_bu_1") or order.get("x_studio_bu") or "").strip()
            else:
                bu = ""
            pais = BU_TO_PAIS.get(bu)
            if pais is None:
                continue  # BU desconocido → descartar
            t["_pais"] = pais
            t["_tab"]  = "intl"
            t["_label2"] = label2
            classified[pais]["intl"].append(t)
            # También al dashboard Internacional
            classified["internacional"]["intl"].append(t)

        elif label2 in LABEL_TO_PAIS:
            pais = LABEL_TO_PAIS[label2]
            t["_pais"] = pais
            t["_tab"]  = "local"
            t["_label2"] = label2
            classified[pais]["local"].append(t)

        else:
            # Fallback: inferir por company_id
            pais = COMPANY_TO_PAIS.get(cid)
            if pais is None:
                continue
            t["_pais"] = pais
            t["_tab"]  = "local"
            t["_label2"] = label2
            classified[pais]["local"].append(t)

    totals = {p: {tab: len(classified[p][tab]) for tab in ("local","intl")} for p in classified}
    print(f"  Clasificación: {totals}")
    return classified

# ─── CÁLCULO DE KPIs ───────────────────────────────────────────────────────────
def get_amount_usd(t):
    """Retorna el importe en USD de una tarea."""
    order = t.get("_order")
    if not order: return 0.0
    amt = order.get("amount_untaxed") or 0
    cur = order.get("currency_id")
    cur_code = cur[1] if isinstance(cur, list) and len(cur) > 1 else "USD"
    # Extraer código limpio (ej: "ARS" de "ARS" o de "[1, 'ARS']")
    cur_code = cur_code.strip().upper()
    return to_usd(amt, cur_code)

def get_fecha_pub(t):
    return parse_date(t.get("x_studio_fecha_de_publicacin"))

def get_fecha_estimada(t):
    """Fecha estimada para pipeline: x_studio_fecha_limite_ops (Fecha Estimada) o date_deadline."""
    v = t.get("x_studio_fecha_limite_ops") or t.get("date_deadline")
    return parse_date(v)

def is_implementado(t):
    return bool(get_fecha_pub(t))

def compute_kpis(tasks, mes_actual, anio_actual):
    """Calcula KPIs para un conjunto de tareas."""
    today = date.today()

    impl_mes   = 0.0
    impl_ytd   = 0.0
    cnt_mes    = 0
    cnt_ytd    = 0
    campanas_mes = set()
    campanas_ytd = set()
    talentos_mes = set()

    pendientes       = []
    vencidas         = []
    dias_cierre_list = []

    for t in tasks:
        fpub = get_fecha_pub(t)
        order = t.get("_order")
        oid = order["id"] if order else None
        amt = get_amount_usd(t)

        if fpub:
            # Implementado
            if fpub.year == anio_actual and fpub.month == mes_actual:
                impl_mes += amt
                cnt_mes += 1
                if oid: campanas_mes.add(oid)
                talento = _parse_talento(t.get("name", ""))
                talentos_mes.add(talento)

            if fpub.year == anio_actual and fpub.month <= mes_actual:
                impl_ytd += amt
                cnt_ytd += 1
                if oid: campanas_ytd.add(oid)

            # Velocidad de cierre: días desde date_order hasta fecha_pub
            if order and order.get("date_order"):
                dorder = parse_date(order["date_order"])
                if dorder:
                    dias = (fpub - dorder).days
                    if 0 <= dias <= 365:
                        dias_cierre_list.append(dias)
        else:
            # Pendiente
            pendientes.append(t)
            ddl = get_fecha_estimada(t)
            if ddl and ddl < today:
                vencidas.append(t)

    vel_cierre = int(sum(dias_cierre_list) / len(dias_cierre_list)) if dias_cierre_list else 0

    # Tasa de implementación: implementados_mes / (implementados_mes + pendientes_con_fecha_este_mes)
    pendientes_mes = [t for t in pendientes
                      if get_fecha_estimada(t) and
                      get_fecha_estimada(t).year == anio_actual and
                      get_fecha_estimada(t).month == mes_actual]
    total_estimados = cnt_mes + len(pendientes_mes)
    tasa = cnt_mes / total_estimados if total_estimados else 0

    return {
        "implementado_mes":   impl_mes,
        "implementado_ytd":   impl_ytd,
        "cnt_mes":            cnt_mes,
        "cnt_ytd":            cnt_ytd,
        "campanas":           len(campanas_mes),
        "talentos":           len(talentos_mes),
        "tasa_impl":          tasa,
        "total_estimados":    total_estimados,
        "pendientes":         pendientes,
        "vencidas":           vencidas,
        "vel_cierre":         vel_cierre,
        "sin_fecha":          sum(1 for t in pendientes if not get_fecha_estimada(t)),
    }

def _parse_talento(name):
    """Extrae nombre del talento del nombre de la subtarea."""
    if "(" in name:
        return name[:name.index("(")].strip()
    return name.strip()

def _parse_contenido(name):
    """Extrae tipo de contenido del nombre de la subtarea."""
    if "(" in name and ")" in name:
        return name[name.index("(")+1 : name.rindex(")")].strip()
    return ""

def _tipo_contenido(t):
    """Determina el tipo de contenido."""
    c = _parse_contenido(t.get("name", ""))
    raw_tipo = (_f(t, "x_studio_tipo_de_contenido_1") or "").lower()
    nombre_lower = (t.get("name", "") or "").lower()
    # Heurística simple por tipo
    if any(k in nombre_lower or k in c.lower() for k in ["artistico", "artístico", "host", "conducción", "conduccion", "exclusividad"]):
        return "Artistico"
    if any(k in nombre_lower or k in c.lower() for k in ["regional", "yerba", "regional"]):
        return "Regional"
    if any(k in nombre_lower or k in c.lower() for k in ["canje", "barter"]):
        return "Canje"
    return "Comercial"

def _pill_class(tipo):
    mapping = {"Comercial": "pb", "Regional": "pp", "Artistico": "pg", "Canje": "pm2"}
    return mapping.get(tipo, "pb")

# ─── PRESUPUESTO ───────────────────────────────────────────────────────────────
BUDGET_CONFIG = {
    # sheet_name: {tipo: row_idx(0-based), tc_override}
    "argentina": {
        "sheet": "BG ARG",
        "local_total": 7,    # fila 8
        "comercial":   18,   # fila 19
        "artistico":   27,   # fila 28
        "regional":    34,   # fila 35
        "intl":        42,   # fila 43
        "tc_override": None,
    },
    "chile": {
        "sheet": "BG CHI",
        "local_total": None,
        "comercial":   16,   # fila 17
        "artistico":   24,   # fila 25
        "regional":    31,   # fila 32
        "intl":        41,   # fila 42
        "tc_override": TC_PRESUP_CLP,
    },
    "colombia": {
        "sheet": "BG COL",
        "local_total": 7,    # fila 8
        "comercial":   21,   # fila 22
        "artistico":   30,   # fila 31
        "regional":    37,   # fila 38
        "intl":        46,   # fila 47 — Internacional del país
        "tc_override": None,
    },
    "usa": {
        "sheet": "BG USA",
        "local_total": 6,    # fila 7
        "comercial":   17,   # fila 18
        "artistico":   27,   # fila 28
        "regional":    34,   # fila 35
        "intl":        43,   # fila 44
        "tc_override": None,
    },
    "peru": {
        "sheet": "BG ARG",   # si no hay sheet propio usar ARG como fallback
        "local_total": None,
        "comercial":   None,
        "artistico":   None,
        "regional":    None,
        "intl":        None,
        "tc_override": None,
    },
    "internacional": {
        "sheet": "BG INT",
        "local_total": 0,    # fila 1 = Ventas Totales
        "comercial":   None,
        "artistico":   None,
        "regional":    None,
        "intl":        None,
        "tc_override": None,
    },
}

def load_budget(budget_file):
    """Carga el Excel de presupuesto. Retorna dict pais→{tipo→[meses]}."""
    if not os.path.exists(budget_file):
        print(f"  ⚠️  Budget file no encontrado: {budget_file}")
        return {}

    wb = openpyxl.load_workbook(budget_file, read_only=True, data_only=True)
    budgets = {}

    for pais, cfg in BUDGET_CONFIG.items():
        sheet_name = cfg["sheet"]
        if sheet_name not in wb.sheetnames:
            budgets[pais] = {}
            continue

        ws = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))
        tc = cfg.get("tc_override")

        def get_row_months(row_idx):
            """Columnas 3-14 (0-based) = Ene-Dic"""
            if row_idx is None: return [0]*12
            if row_idx >= len(rows): return [0]*12
            row = rows[row_idx]
            vals = []
            for col in range(3, 15):  # columnas D-O
                v = row[col] if col < len(row) else None
                try:
                    v = float(v) if v else 0.0
                except (TypeError, ValueError):
                    v = 0.0
                if tc:
                    v = v / tc
                vals.append(v)
            return vals

        budgets[pais] = {
            "local_total": get_row_months(cfg.get("local_total")),
            "comercial":   get_row_months(cfg.get("comercial")),
            "artistico":   get_row_months(cfg.get("artistico")),
            "regional":    get_row_months(cfg.get("regional")),
            "intl":        get_row_months(cfg.get("intl")),
        }

    wb.close()
    return budgets

# ─── GENERADORES HTML ──────────────────────────────────────────────────────────
CSS = r"""
:root{--bg:#0d0f14;--s1:#161b24;--s2:#1e2535;--bdr:#262f42;
  --a1:#4f8ef7;--a2:#a78bfa;--a3:#34d399;--a4:#fb923c;--a5:#f472b6;
  --tx:#e8eaf0;--tx2:#8b93a8;--tx3:#5a6380;--red:#f87171;--yellow:#fbbf24}
*{box-sizing:border-box;margin:0;padding:0}
body{background:var(--bg);color:var(--tx);font-family:'Inter',system-ui,sans-serif;font-size:13.5px;line-height:1.5}
.hdr{background:var(--s1);border-bottom:1px solid var(--bdr);padding:14px 28px;display:flex;align-items:center;gap:14px;position:sticky;top:0;z-index:200;flex-wrap:wrap}
.hdr-logo{font-size:18px;font-weight:700;color:var(--a1)}
.hdr-sep{color:var(--tx3)}
.hdr-title{font-size:13px;color:var(--tx2);font-weight:500}
.hdr-badge{margin-left:auto;background:rgba(79,142,247,.12);color:var(--a1);border:1px solid rgba(79,142,247,.3);border-radius:20px;padding:3px 12px;font-size:11px;font-weight:600}
.hdr-date{font-size:11px;color:var(--tx3)}
.mtabs{display:flex;gap:0;border-bottom:2px solid var(--bdr);padding:0 28px;background:var(--s1)}
.mtab{padding:14px 24px;cursor:pointer;font-size:14px;font-weight:600;color:var(--tx2);border-bottom:3px solid transparent;margin-bottom:-2px;transition:all .15s;display:flex;align-items:center;gap:8px}
.mtab:hover{color:var(--tx)}
.mtab.local.active{color:var(--a1);border-bottom-color:var(--a1)}
.mtab.intl.active{color:var(--a2);border-bottom-color:var(--a2)}
.mpanel{display:none;padding:24px 28px}
.mpanel.active{display:block}
.sec-label{display:inline-flex;align-items:center;gap:8px;font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:1px;padding:5px 12px;border-radius:20px;margin-bottom:20px}
.sec-label.local{background:rgba(79,142,247,.1);color:var(--a1);border:1px solid rgba(79,142,247,.25)}
.sec-label.intl{background:rgba(167,139,250,.1);color:var(--a2);border:1px solid rgba(167,139,250,.25)}
.stabs{display:flex;gap:2px;border-bottom:1px solid var(--bdr);margin-bottom:24px;overflow-x:auto}
.stab{padding:10px 18px;cursor:pointer;font-size:12.5px;font-weight:500;color:var(--tx2);border-bottom:2px solid transparent;margin-bottom:-1px;white-space:nowrap;transition:all .15s}
.stab:hover{color:var(--tx)}
.stab.active{color:var(--a1);border-bottom-color:var(--a1)}
.spanel{display:none}.spanel.active{display:block}
.badge-r{background:rgba(248,113,113,.2);color:var(--red);border-radius:10px;padding:1px 6px;font-size:10px}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(145px,1fr));gap:12px;margin-bottom:24px}
.kpi{background:var(--s1);border:1px solid var(--bdr);border-radius:10px;padding:16px 18px}
.kpi-l{font-size:10px;text-transform:uppercase;letter-spacing:.8px;color:var(--tx3);margin-bottom:7px}
.kpi-v{font-size:22px;font-weight:700;letter-spacing:-.5px;line-height:1}
.kpi-s{font-size:11px;color:var(--tx3);margin-top:4px}
.kpi-n{font-size:10px;color:var(--tx3);margin-top:3px;font-style:italic;opacity:.8}
.c1{color:var(--a1)}.c2{color:var(--a2)}.c3{color:var(--a3)}.c4{color:var(--a4)}.c5{color:var(--a5)}.cred{color:var(--red)}.cyellow{color:var(--yellow)}
.g2{display:grid;grid-template-columns:1fr 1fr;gap:18px;margin-bottom:18px}
@media(max-width:900px){.g2{grid-template-columns:1fr}}
.card{background:var(--s1);border:1px solid var(--bdr);border-radius:10px;overflow:hidden;margin-bottom:18px}
.card-h{padding:12px 16px;border-bottom:1px solid var(--bdr);font-size:11px;font-weight:600;text-transform:uppercase;letter-spacing:.7px;color:var(--tx2);display:flex;align-items:center;gap:8px;flex-wrap:wrap}
.card-h-note{font-size:10px;color:var(--tx3);font-weight:400;text-transform:none;letter-spacing:0;margin-left:4px}
.dot{width:6px;height:6px;border-radius:50%;flex-shrink:0}
.card-b{padding:14px 16px}
.note{font-size:11.5px;color:var(--tx3);padding:6px 0}
.tw{overflow-x:auto}
table{width:100%;border-collapse:collapse}
th{text-align:left;font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:.6px;color:var(--tx3);padding:6px 9px;border-bottom:1px solid var(--bdr)}
th.r{text-align:right}
td{padding:7px 9px;border-bottom:1px solid #1a2030;font-size:12px;color:var(--tx)}
td.r{text-align:right;font-variant-numeric:tabular-nums}
td.m{color:var(--tx2)}
tr:last-child td{border-bottom:none}
tr:hover td{background:rgba(79,142,247,.04)}
.bars{display:flex;flex-direction:column;gap:9px}
.bi{}.bm{display:flex;justify-content:space-between;margin-bottom:3px}
.bl{font-size:12px;color:var(--tx);overflow:hidden;text-overflow:ellipsis;white-space:nowrap;max-width:62%}
.bv{font-size:11px;color:var(--tx2)}
.bt{height:5px;background:var(--bdr);border-radius:3px;overflow:hidden}
.bf{height:100%;border-radius:3px;background:var(--a1)}
.bf.p{background:var(--a2)}.bf.g{background:var(--a3)}.bf.o{background:var(--a4)}
.hbars{display:flex;flex-direction:column;gap:7px}
.hr{display:flex;align-items:center;gap:10px}
.hl{width:50px;font-size:11px;color:var(--tx2);text-align:right;flex-shrink:0}
.ht{flex:1;height:20px;background:var(--bdr);border-radius:3px;overflow:hidden}
.hf{height:100%;background:linear-gradient(90deg,var(--a1),var(--a2));border-radius:3px;display:flex;align-items:center;padding-left:8px}
.hv{font-size:10px;font-weight:600;color:rgba(255,255,255,.9)}
.pgrid{display:grid;grid-template-columns:repeat(auto-fit,minmax(105px,1fr));gap:10px;margin-bottom:20px}
.pm{background:var(--s2);border:1px solid var(--bdr);border-radius:10px;padding:14px 12px;text-align:center}
.pcur{border-color:var(--a1);background:rgba(79,142,247,.07)}
.pml{font-size:10px;text-transform:uppercase;letter-spacing:.7px;color:var(--tx3);margin-bottom:6px}
.pc{font-size:26px;font-weight:700;color:var(--a2);line-height:1;margin-bottom:3px}
.pi{font-size:10px;color:var(--tx2)}
.bgrow{display:flex;align-items:center;gap:12px;padding:10px 0;border-bottom:1px solid var(--bdr)}
.bgrow:last-child{border-bottom:none}
.bgtipo{font-size:12px;font-weight:600;width:85px;flex-shrink:0}
.bgtrk{flex:1;height:11px;background:var(--bdr);border-radius:6px;overflow:hidden;position:relative}
.bgbg{height:100%;background:rgba(255,255,255,.06);border-radius:6px;position:absolute;inset:0}
.bgreal{height:100%;border-radius:6px;position:absolute;top:0;left:0;transition:width .4s}
.bgreal.ok{background:var(--a3)}.bgreal.warn{background:var(--yellow)}.bgreal.low{background:var(--red)}.bgreal.over{background:var(--a5)}
.bgnums{font-size:11px;color:var(--tx2);min-width:175px;text-align:right}
.bgpct{font-size:11px;font-weight:700;min-width:40px;text-align:right}
.acc{display:flex;flex-direction:column;gap:10px}
.ai{border-radius:9px;padding:13px 15px;display:flex;gap:11px;align-items:flex-start}
.aico{font-size:15px;flex-shrink:0;margin-top:2px}
.albl{font-size:10px;text-transform:uppercase;letter-spacing:.7px;font-weight:600;margin-bottom:4px}
.atxt{font-size:13px;line-height:1.6;color:var(--tx)}
.ai-urgente{background:rgba(248,113,113,.08);border:1px solid rgba(248,113,113,.25)}.ai-urgente .albl{color:var(--red)}
.ai-alerta{background:rgba(251,191,36,.08);border:1px solid rgba(251,191,36,.25)}.ai-alerta .albl{color:var(--yellow)}
.ai-riesgo{background:rgba(251,146,60,.08);border:1px solid rgba(251,146,60,.25)}.ai-riesgo .albl{color:var(--a4)}
.ai-operativo{background:rgba(167,139,250,.08);border:1px solid rgba(167,139,250,.25)}.ai-operativo .albl{color:var(--a2)}
.ai-presupuesto{background:rgba(52,211,153,.08);border:1px solid rgba(52,211,153,.25)}.ai-presupuesto .albl{color:var(--a3)}
.pill{display:inline-block;padding:2px 8px;border-radius:12px;font-size:10px;font-weight:600}
.pb{background:rgba(79,142,247,.15);color:var(--a1)}
.pp{background:rgba(167,139,250,.15);color:var(--a2)}
.pg{background:rgba(52,211,153,.15);color:var(--a3)}
.po{background:rgba(251,146,60,.15);color:var(--a4)}
.pm2{background:rgba(90,99,128,.2);color:var(--tx2)}
.fxnote{background:rgba(167,139,250,.06);border:1px solid rgba(167,139,250,.2);border-radius:8px;padding:9px 13px;font-size:11px;color:var(--tx3);margin-bottom:18px}
.fxnote strong{color:var(--a2)}
.sec-title{font-size:11px;font-weight:600;text-transform:uppercase;letter-spacing:.8px;color:var(--tx3);margin-bottom:12px;padding-bottom:7px;border-bottom:1px solid var(--bdr)}
"""

JS = r"""
function sw(pid,tab){
  document.querySelectorAll('#pg_'+pid+' > .mtabs .mtab').forEach(t=>t.classList.remove('active'));
  document.querySelectorAll('#pg_'+pid+' > .mpanel').forEach(p=>p.classList.remove('active'));
  var el=document.querySelector('#pg_'+pid+' > .mtabs .mtab.'+tab);
  if(el)el.classList.add('active');
  var ep=document.getElementById(pid+'_'+tab);
  if(ep)ep.classList.add('active');
}
function sst(prefix,tab){
  var nav=document.getElementById(prefix+'_nav');
  if(!nav)return;
  nav.querySelectorAll('.stab').forEach(t=>t.classList.remove('active'));
  var pp=nav.parentElement;
  pp.querySelectorAll('.spanel').forEach(p=>p.classList.remove('active'));
  var trig=nav.querySelector('[onclick*="\''+tab+'\'"]');
  if(trig)trig.classList.add('active');
  var sp=document.getElementById(prefix+'_'+tab);
  if(sp)sp.classList.add('active');
}
"""

PAIS_LABEL = {
    "argentina":    "Campañas Argentinas",
    "chile":        "Campañas Chilenas",
    "colombia":     "Campañas Colombianas",
    "usa":          "US Campaigns",
    "peru":         "Campañas Peruanas",
    "internacional":"Campañas Internacionales",
}

def build_implementado_html(tasks, kpis, mes_actual, anio_actual, prefix):
    mes_str = mes_label(anio_actual, mes_actual)

    # Desglose por tipo
    tipo_cnt = defaultdict(int)
    tipo_usd = defaultdict(float)
    for t in tasks:
        fpub = get_fecha_pub(t)
        if fpub and fpub.year == anio_actual and fpub.month == mes_actual:
            tipo = _tipo_contenido(t)
            tipo_cnt[tipo] += 1
            tipo_usd[tipo] += get_amount_usd(t)

    tipo_rows = ""
    for tipo in ["Comercial", "Regional", "Artistico", "Canje"]:
        if tipo_cnt[tipo]:
            tipo_rows += f"""<tr><td><span class="pill {_pill_class(tipo)}">{tipo}</span></td>
              <td class="r">{tipo_cnt[tipo]}</td>
              <td class="r">{fmt_usd(tipo_usd[tipo])}</td></tr>"""

    # Top talentos del mes
    tal_cnt = defaultdict(int)
    tal_usd = defaultdict(float)
    for t in tasks:
        fpub = get_fecha_pub(t)
        if fpub and fpub.year == anio_actual and fpub.month == mes_actual:
            tal = _parse_talento(t.get("name", ""))
            tal_cnt[tal] += 1
            tal_usd[tal] += get_amount_usd(t)
    top_tal = sorted(tal_usd.items(), key=lambda x: -x[1])[:20]
    tal_rows = "".join(
        f'<tr><td>{n}</td><td class="r">{tal_cnt[n]}</td><td class="r">{fmt_usd(u)}</td></tr>'
        for n, u in top_tal
    )

    # Top clientes del mes
    cli_cnt = defaultdict(int)
    cli_usd = defaultdict(float)
    for t in tasks:
        fpub = get_fecha_pub(t)
        if fpub and fpub.year == anio_actual and fpub.month == mes_actual:
            order = t.get("_order")
            cli = ""
            if order and order.get("partner_id"):
                cli = order["partner_id"][1] if isinstance(order["partner_id"], list) else str(order["partner_id"])
            cli_cnt[cli] += 1
            cli_usd[cli] += get_amount_usd(t)
    top_cli = sorted(cli_usd.items(), key=lambda x: -x[1])[:20]
    cli_rows = "".join(
        f'<tr><td>{n}</td><td class="r">{cli_cnt[n]}</td><td class="r">{fmt_usd(u)}</td></tr>'
        for n, u in top_cli
    )

    # Concentración de riesgo
    total_impl_mes = kpis["implementado_mes"] or 1
    conc_rows = "".join(
        f'<tr><td>{n}</td><td class="r">{cli_cnt[n]}</td><td class="r">{fmt_usd(u)}</td>'
        f'<td class="r" style="color:{pct_css_color(u/total_impl_mes)}">{u/total_impl_mes*100:.1f}%</td></tr>'
        for n, u in top_cli[:10]
    )

    # Tabla implementadores
    impl_cnt = defaultdict(int)
    impl_usd = defaultdict(float)
    impl_camp = defaultdict(set)
    for t in tasks:
        fpub = get_fecha_pub(t)
        if fpub and fpub.year == anio_actual and fpub.month == mes_actual:
            imp = _f(t, "x_studio_related_field_7v0_1jidluau0")
            if isinstance(imp, list) and len(imp) > 1:
                imp_name = imp[1]
            elif imp:
                imp_name = str(imp)
            else:
                imp_name = "Sin asignar"
            impl_cnt[imp_name] += 1
            impl_usd[imp_name] += get_amount_usd(t)
            order = t.get("_order")
            if order: impl_camp[imp_name].add(order["id"])
    impl_rows = "".join(
        f'<tr><td>{n}</td><td class="r">{impl_cnt[n]}</td><td class="r">{len(impl_camp[n])}</td><td class="r">{fmt_usd(impl_usd[n])}</td></tr>'
        for n in sorted(impl_usd, key=lambda x: -impl_usd[x])
    )

    vencidas = kpis["vencidas"]
    usd_venc = sum(get_amount_usd(t) for t in vencidas)
    pend = kpis["pendientes"]
    sin_fecha = kpis["sin_fecha"]

    return f"""
<div class="kpis">
  <div class="kpi"><div class="kpi-l">Implementado {mes_str}</div><div class="kpi-v c1">{kpis['cnt_mes']}</div><div class="kpi-s">contenidos publicados</div></div>
  <div class="kpi"><div class="kpi-l">Importe</div><div class="kpi-v c3">{fmt_usd(kpis['implementado_mes'])}</div><div class="kpi-s">en USD</div></div>
  <div class="kpi"><div class="kpi-l">Tasa Impl.</div><div class="kpi-v c3">{kpis['tasa_impl']:.0%}</div><div class="kpi-s">{kpis['cnt_mes']} de {kpis['total_estimados']} est.</div><div class="kpi-n">% de campañas estimadas para el mes efectivamente publicadas</div></div>
  <div class="kpi"><div class="kpi-l">Campañas Vencidas</div><div class="kpi-v cred">{len(vencidas)}</div><div class="kpi-s">{fmt_usd(usd_venc)}</div><div class="kpi-n">Pendientes cuya fecha estimada ya pasó. Requieren seguimiento urgente.</div></div>
  <div class="kpi"><div class="kpi-l">Pendientes</div><div class="kpi-v c4">{len(pend)}</div><div class="kpi-s">{sin_fecha} sin fecha</div></div>
  <div class="kpi"><div class="kpi-l">Vel. Cierre</div><div class="kpi-v c2">{kpis['vel_cierre']}</div><div class="kpi-s">días prom.</div><div class="kpi-n">Días promedio desde pedido hasta publicación efectiva.</div></div>
</div>
<div class="g2">
  <div class="card"><div class="card-h"><div class="dot" style="background:var(--a2)"></div>Por Tipo de Contenido</div>
    <div class="card-b"><div class="tw"><table><thead><tr><th>Tipo</th><th class="r">Cont.</th><th class="r">USD</th></tr></thead><tbody>{tipo_rows}</tbody></table></div></div></div>
  <div class="card"><div class="card-h"><div class="dot" style="background:var(--a3)"></div>Top Talentos del Mes</div>
    <div class="card-b"><div class="tw"><table><thead><tr><th>Talento</th><th class="r">Cont.</th><th class="r">USD</th></tr></thead><tbody>{tal_rows}</tbody></table></div></div></div>
</div>
<div class="card"><div class="card-h"><div class="dot" style="background:var(--a4)"></div>Top Clientes del Mes</div>
  <div class="card-b"><div class="tw"><table><thead><tr><th>Cliente</th><th class="r">Cont.</th><th class="r">USD</th></tr></thead><tbody>{cli_rows}</tbody></table></div></div></div>
<div class="card"><div class="card-h"><div class="dot" style="background:var(--a5)"></div>Concentración de Riesgo por Cliente</div>
  <div class="card-b"><div class="tw"><table><thead><tr><th>Cliente</th><th class="r">Cont.</th><th class="r">USD</th><th class="r">% total</th></tr></thead><tbody>{conc_rows}</tbody></table></div></div></div>
<div class="card"><div class="card-h"><div class="dot" style="background:var(--a1)"></div>Por Implementador</div>
  <div class="card-b"><div class="tw"><table><thead><tr><th>Implementador</th><th class="r">Contenidos</th><th class="r">Campañas</th><th class="r">USD</th></tr></thead><tbody>{impl_rows}</tbody></table></div></div></div>
"""

def build_budget_html(tasks, budget_pais, mes_actual, anio_actual, tab):
    """Genera sección de presupuesto con barras de progreso."""
    if not budget_pais:
        return '<p class="note">Presupuesto no disponible.</p>'

    mes_idx = mes_actual - 1  # 0-based

    # Reales: calcular por tipo
    def get_real(tipo_check):
        total = 0.0
        for t in tasks:
            fpub = get_fecha_pub(t)
            if not fpub: continue
            if fpub.year == anio_actual and fpub.month == mes_actual:
                tipo = _tipo_contenido(t)
                if tipo_check == "total" or tipo.lower() == tipo_check.lower():
                    total += get_amount_usd(t)
        return total

    def get_real_ytd(tipo_check):
        total = 0.0
        for t in tasks:
            fpub = get_fecha_pub(t)
            if not fpub: continue
            if fpub.year == anio_actual and fpub.month <= mes_actual:
                tipo = _tipo_contenido(t)
                if tipo_check == "total" or tipo.lower() == tipo_check.lower():
                    total += get_amount_usd(t)
        return total

    key_tab = "intl" if tab == "intl" else "local_total"

    filas = []
    if tab == "local":
        filas = [
            ("Total Local",  "total",     key_tab,     True),
            ("Comercial",    "comercial",  "comercial", False),
            ("Artístico",    "artistico",  "artistico", False),
            ("Regional",     "regional",   "regional",  False),
        ]
    else:
        filas = [
            ("Internacional","total",      "intl",      True),
        ]

    mes_str = mes_label(anio_actual, mes_actual)
    ytd_str = f"Ene–{mes_label(anio_actual, mes_actual).split()[0]} {anio_actual}"

    rows_html = ""
    for label, tipo_check, bkey, bold in filas:
        bdata = budget_pais.get(bkey)
        if not bdata:
            continue
        presup_mes = bdata[mes_idx] if mes_idx < len(bdata) else 0
        presup_ytd = sum(bdata[:mes_actual])
        real_mes   = get_real(tipo_check)
        real_ytd   = get_real_ytd(tipo_check)

        def brow(lbl, real, presup, is_bold):
            if not presup: return ""
            pct = real / presup if presup else 0
            cls = pct_color(pct)
            bar_w = min(pct * 100, 100)
            style = "font-weight:700" if is_bold else ""
            return f"""<div class="bgrow">
  <span class="bgtipo" style="width:100px;{style}">{lbl}</span>
  <div style="flex:1">
    <div style="margin-bottom:4px">
      <span style="font-size:10px;color:var(--tx3)">Mes ({mes_str})</span>
    </div>
    <div class="bgtrk">
      <div class="bgbg"></div>
      <div class="bgreal {cls}" style="width:{bar_w:.1f}%"></div>
    </div>
  </div>
  <div class="bgnums" style="font-size:11px;color:var(--tx2)">{fmt_usd(real)} / {fmt_usd(presup)}</div>
  <div class="bgpct" style="color:{pct_css_color(pct)}">{pct:.0%}</div>
</div>
<div class="bgrow">
  <span class="bgtipo" style="width:100px;color:var(--tx3);font-size:11px">{lbl} YTD</span>
  <div style="flex:1">
    <div style="margin-bottom:4px">
      <span style="font-size:10px;color:var(--tx3)">YTD (Ene–{mes_label(anio_actual, mes_actual).split()[0]} {anio_actual})</span>
    </div>
    <div class="bgtrk">
      <div class="bgbg"></div>
      <div class="bgreal {pct_color(real_ytd/presup_ytd if presup_ytd else 0)}" style="width:{min((real_ytd/presup_ytd if presup_ytd else 0)*100,100):.1f}%"></div>
    </div>
  </div>
  <div class="bgnums">{fmt_usd(real_ytd)} / {fmt_usd(presup_ytd)}</div>
  <div class="bgpct" style="color:{pct_css_color(real_ytd/presup_ytd if presup_ytd else 0)}">{real_ytd/presup_ytd:.0%}</div>
</div>"""

        rows_html += brow(label, real_mes, presup_mes, bold)

    if not rows_html:
        return '<p class="note">Sin datos de presupuesto para este período.</p>'

    return f"""<div class="card"><div class="card-h"><div class="dot" style="background:var(--a1)"></div>vs Presupuesto — Mes ({mes_str}) y YTD Ene–{mes_label(anio_actual, mes_actual).split()[0]} {anio_actual}</div>
<div class="card-b">
<p class="note" style="margin-bottom:14px">Verde ≥90% · Amarillo ≥60% · Rojo &lt;60% · Rosa = superado.</p>
{rows_html}
</div></div>"""

def build_pipeline_html(tasks, mes_actual, anio_actual):
    today = date.today()
    pendientes = [t for t in tasks if not is_implementado(t)]

    # Agrupar por mes estimado
    by_mes = defaultdict(list)
    sin_fecha_list = []
    for t in pendientes:
        ddl = get_fecha_estimada(t)
        if ddl:
            by_mes[(ddl.year, ddl.month)].append(t)
        else:
            sin_fecha_list.append(t)

    # Grid de meses
    meses_sorted = sorted(by_mes.keys())
    pgrid = ""
    for (y, m) in meses_sorted[-12:]:
        items = by_mes[(y, m)]
        usd = sum(get_amount_usd(t) for t in items)
        cur_class = "pcur" if y == anio_actual and m == mes_actual else ""
        pgrid += f"""<div class="pm {cur_class}">
  <div class="pml">{mes_label(y, m)}</div>
  <div class="pc">{len(items)}</div>
  <div class="pi">{fmt_usd(usd)}</div>
</div>"""

    sin_fecha_note = f'<p class="note">+ {len(sin_fecha_list)} subtareas sin fecha estimada.</p>' if sin_fecha_list else ""

    # Vencidas
    vencidas = sorted(
        [t for t in pendientes if (get_fecha_estimada(t) or date.max) < today],
        key=lambda t: get_fecha_estimada(t) or date.max
    )

    def task_row(t, show_dias=True):
        nombre = t.get("name", "")
        talento = _parse_talento(nombre)
        cont = _parse_contenido(nombre)
        order = t.get("_order")
        cli = ""
        if order and order.get("partner_id"):
            cli = order["partner_id"][1] if isinstance(order["partner_id"], list) else ""
        tipo = _tipo_contenido(t)
        usd = get_amount_usd(t)
        ddl = get_fecha_estimada(t)
        ddl_str = ddl.strftime("%d/%m/%Y") if ddl else "—"
        dias_td = ""
        if show_dias and ddl and ddl < today:
            dias = (today - ddl).days
            dias_td = f'<td class="r" style="color:var(--red)">{dias}d</td>'
        elif show_dias:
            dias_td = '<td class="r">—</td>'
        dias_header = '<th class=r>Días venc.</th>' if show_dias else ''
        return (f'<tr><td class="">{talento}</td><td class="m">{cont}</td>'
                f'<td class="m">{cli}</td>'
                f'<td><span class="pill {_pill_class(tipo)}">{tipo}</span></td>'
                f'<td class="r">{fmt_usd(usd)}</td>'
                f'<td class="m">{ddl_str}</td>{dias_td}</tr>')

    venc_rows = "".join(task_row(t, show_dias=True) for t in vencidas[:50])
    venc_card = ""
    if vencidas:
        venc_card = f"""<div class="card" style="margin-bottom:20px">
  <div class="card-h"><div class="dot" style="background:var(--red)"></div>Campañas Vencidas ({len(vencidas)})<span class="card-h-note">Pendientes con fecha estimada ya pasada — requieren seguimiento inmediato.</span></div>
  <div class="card-b"><div class="tw"><table><thead><tr><th>Talento</th><th>Contenido</th><th>Cliente</th><th>Tipo</th><th class=r>USD</th><th>F.Estimada</th><th class=r>Días venc.</th></tr></thead><tbody>{venc_rows}</tbody></table></div></div>
</div>"""

    # Top pendientes por importe
    top_pend = sorted(pendientes, key=lambda t: -get_amount_usd(t))[:30]
    top_rows = "".join(task_row(t, show_dias=False) for t in top_pend)

    return f"""
<p class="sec-title">Pipeline por Mes Estimado</p>
<div class="pgrid">{pgrid}</div>
{sin_fecha_note}
{venc_card}
<p class="sec-title">Top Pendientes por Importe</p>
<div class="tw"><table><thead><tr><th>Talento</th><th>Contenido</th><th>Cliente</th><th>Tipo</th><th class=r>USD</th><th>F.Estimada</th></tr></thead><tbody>{top_rows}</tbody></table></div>
"""

def build_historial_html(tasks, mes_actual, anio_actual):
    today = date.today()

    # Implementados históricos
    impl_by_mes = defaultdict(lambda: {"cnt": 0, "usd": 0.0})
    for t in tasks:
        fpub = get_fecha_pub(t)
        if not fpub: continue
        key = (fpub.year, fpub.month)
        impl_by_mes[key]["cnt"] += 1
        impl_by_mes[key]["usd"] += get_amount_usd(t)

    meses_sorted = sorted(impl_by_mes.keys())
    max_cnt = max((v["cnt"] for v in impl_by_mes.values()), default=1) or 1

    hbars = ""
    for (y, m) in meses_sorted[-18:]:
        v = impl_by_mes[(y, m)]
        pct = int(v["cnt"] / max_cnt * 100)
        hbars += f"""<div class="hr">
<div class="hl">{mes_label(y, m)}</div>
<div class="ht"><div class="hf" style="width:{pct}%"><span class="hv">{v['cnt']}</span></div></div>
<div style="font-size:10px;color:var(--tx3);white-space:nowrap;min-width:80px;text-align:right">{fmt_usd(v['usd'])}</div>
</div>"""

    # Top talentos históricos
    tal_cnt = defaultdict(int)
    tal_usd = defaultdict(float)
    for t in tasks:
        fpub = get_fecha_pub(t)
        if not fpub: continue
        tal = _parse_talento(t.get("name", ""))
        tal_cnt[tal] += 1
        tal_usd[tal] += get_amount_usd(t)
    top_tal = sorted(tal_usd.items(), key=lambda x: -x[1])[:15]
    tal_rows = "".join(
        f'<tr><td>{n}</td><td class="r">{tal_cnt[n]}</td><td class="r">{fmt_usd(u)}</td></tr>'
        for n, u in top_tal
    )

    # Contenido promedio por influencer mensual
    # Para los últimos 6 meses
    ultimos_6 = sorted(impl_by_mes.keys())[-6:]
    tal_mes = defaultdict(lambda: defaultdict(lambda: {"cnt": 0, "usd": 0.0}))
    for t in tasks:
        fpub = get_fecha_pub(t)
        if not fpub: continue
        key = (fpub.year, fpub.month)
        if key not in ultimos_6: continue
        tal = _parse_talento(t.get("name", ""))
        tal_mes[tal][key]["cnt"] += 1
        tal_mes[tal][key]["usd"] += get_amount_usd(t)

    # Solo talentos con al menos 2 meses de datos
    top_talentos = [n for n, d in tal_mes.items() if len(d) >= 2][:20]
    prom_header = "".join(f'<th class="r">{mes_label(y,m)}</th>' for y, m in ultimos_6)
    prom_rows = ""
    for tal in top_talentos:
        cells = ""
        for key in ultimos_6:
            d = tal_mes[tal].get(key)
            if d and d["cnt"]:
                prom = d["usd"] / d["cnt"]
                usd_str = fmt_usd(d["usd"])
                cnt_str = d["cnt"]
                prom_str = fmt_usd(prom)
                cells += f'<td class="r" title="{usd_str} / {cnt_str} cont.">{prom_str}</td>'
            else:
                cells += '<td class="r" style="color:var(--tx3)">—</td>'
        prom_rows += f"<tr><td>{tal}</td>{cells}</tr>"

    # Top clientes históricos
    cli_cnt = defaultdict(int)
    cli_usd = defaultdict(float)
    for t in tasks:
        fpub = get_fecha_pub(t)
        if not fpub: continue
        order = t.get("_order")
        if order and order.get("partner_id"):
            cli = order["partner_id"][1] if isinstance(order["partner_id"], list) else ""
        else:
            cli = ""
        cli_cnt[cli] += 1
        cli_usd[cli] += get_amount_usd(t)
    top_cli = sorted(cli_usd.items(), key=lambda x: -x[1])[:15]
    cli_rows = "".join(
        f'<tr><td>{n}</td><td class="r">{cli_cnt[n]}</td><td class="r">{fmt_usd(u)}</td></tr>'
        for n, u in top_cli
    )

    return f"""
<div class="g2">
  <div class="card"><div class="card-h"><div class="dot" style="background:var(--a1)"></div>Contenidos por Mes (histórico)</div>
    <div class="card-b"><div class="hbars">{hbars}</div></div></div>
  <div class="card"><div class="card-h"><div class="dot" style="background:var(--a4)"></div>Top Talentos Históricos</div>
    <div class="card-b"><div class="tw"><table><thead><tr><th>Talento</th><th class="r">Cont.</th><th class="r">USD acum.</th></tr></thead><tbody>{tal_rows}</tbody></table></div></div></div>
</div>
<div class="card" style="margin-bottom:18px"><div class="card-h"><div class="dot" style="background:var(--a3)"></div>Contenido Promedio por Influencer (mensual)<span class="card-h-note">USD implementado del mes ÷ cantidad de contenidos del influencer ese mes. Pasá el cursor para ver el detalle.</span></div>
  <div class="card-b"><div class="tw"><table><thead><tr><th>Influencer</th>{prom_header}</tr></thead><tbody>{prom_rows}</tbody></table></div></div></div>
<div class="card"><div class="card-h"><div class="dot" style="background:var(--a5)"></div>Top Clientes Históricos</div>
  <div class="card-b"><div class="tw"><table><thead><tr><th>Cliente</th><th class="r">Cont.</th><th class="r">USD acum.</th></tr></thead><tbody>{cli_rows}</tbody></table></div></div></div>
"""

def build_accionables_html(kpis, pais, tab):
    tag = "[Local]" if tab == "local" else "[Intl]"
    items = []

    vencidas = kpis["vencidas"]
    if vencidas:
        # Buscar la más urgente
        mas_urgente = max(vencidas, key=lambda t: (
            (date.today() - get_fecha_estimada(t)).days
            if get_fecha_estimada(t) else 0
        ))
        ddl_mu = parse_date(mas_urgente.get("date_deadline"))
        dias_mu = (date.today() - ddl_mu).days if ddl_mu else 0
        items.append(("urgente", "🚨", "Urgente",
            f'{tag} {len(vencidas)} campaña(s) vencida(s). Mayor urgencia: {mas_urgente.get("name","")[:50]} — {fmt_usd(get_amount_usd(mas_urgente))} · {dias_mu}d de atraso.'))

    if kpis["sin_fecha"] > 0:
        items.append(("operativo", "🔧", "Operativo",
            f'{tag} {kpis["sin_fecha"]} subtareas sin fecha estimada. Asignar fechas para mejorar la planificación.'))

    if kpis["vel_cierre"] > 30:
        items.append(("riesgo", "⏱", "Riesgo Operativo",
            f'{tag} Velocidad de cierre promedio: {kpis["vel_cierre"]} días. Revisar cuellos de botella en aprobación de contenidos.'))

    if kpis["tasa_impl"] < 0.5 and kpis["total_estimados"] > 5:
        items.append(("alerta", "📉", "Alerta de Implementación",
            f'{tag} Tasa de implementación baja: {kpis["tasa_impl"]:.0%}. Solo {kpis["cnt_mes"]} de {kpis["total_estimados"]} campañas estimadas fueron publicadas.'))

    if not items:
        items.append(("presupuesto", "✅", "Sin alertas",
            f'{tag} No se detectaron alertas críticas para este período.'))

    html = '<div class="acc">'
    for cls, ico, lbl, txt in items:
        html += f'<div class="ai ai-{cls}"><span class="aico">{ico}</span><div><div class="albl">{lbl}</div><div class="atxt">{txt}</div></div></div>'
    html += '</div>'
    return html

def build_tab_panel(pid, prefix, tab_class, tab_label, sec_label, tasks,
                    budget_pais, pais, tab, mes_actual, anio_actual, is_active):
    kpis = compute_kpis(tasks, mes_actual, anio_actual)
    active_cls = "active" if is_active else ""

    impl_html  = build_implementado_html(tasks, kpis, mes_actual, anio_actual, prefix)
    bg_html    = build_budget_html(tasks, budget_pais, mes_actual, anio_actual, tab)
    pipe_html  = build_pipeline_html(tasks, mes_actual, anio_actual)
    hist_html  = build_historial_html(tasks, mes_actual, anio_actual)
    acc_html   = build_accionables_html(kpis, pais, tab)

    venc_count = len(kpis["vencidas"])
    badge_venc = f' <span class="badge-r">{venc_count}</span>' if venc_count else ""

    total_cnt = len(tasks)
    total_usd = sum(get_amount_usd(t) for t in tasks if is_implementado(t))

    fxnote = """<div class="fxnote">💱 <strong>Todos los importes en USD.</strong> TC: BCRA ~1.100 ARS/USD &middot; BCCh 940 CLP/USD &middot; BanRep ~4.250 COP/USD &middot; BCRP ~3.72 PEN/USD. Presupuesto Chile convertido a TC 890.</div>"""

    return f"""
<div class="mpanel {active_cls}" id="{pid}_{tab_class}">
  {fxnote}
  <div class="sec-label {tab_class}">{sec_label}</div>
  <div class="stabs" id="{prefix}_nav">
    <div class="stab active" onclick="sst('{prefix}','impl')">📊 Implementado</div>
    <div class="stab" onclick="sst('{prefix}','bg')">🎯 Presupuesto</div>
    <div class="stab" onclick="sst('{prefix}','pipe')">⏳ Pipeline</div>
    <div class="stab" onclick="sst('{prefix}','hist')">📅 Historial</div>
    <div class="stab" onclick="sst('{prefix}','acc')">⚡ Accionables{badge_venc}</div>
  </div>
  <div class="spanel active" id="{prefix}_impl">{impl_html}</div>
  <div class="spanel" id="{prefix}_bg">{bg_html}</div>
  <div class="spanel" id="{prefix}_pipe">{pipe_html}</div>
  <div class="spanel" id="{prefix}_hist">{hist_html}</div>
  <div class="spanel" id="{prefix}_acc">{acc_html}</div>
</div>"""

def generate_html(pais, local_tasks, intl_tasks, budget, mes_actual, anio_actual):
    """Genera el HTML completo para un país."""
    today_str = datetime.now(timezone.utc).astimezone().strftime("%d %b %Y · %H:%M hs ARG")
    pais_label = PAIS_LABEL[pais]
    pid = f"campanas_{pais}"

    cnt_local = len(local_tasks)
    cnt_intl  = len(intl_tasks)
    usd_local = sum(get_amount_usd(t) for t in local_tasks if is_implementado(t))
    usd_intl  = sum(get_amount_usd(t) for t in intl_tasks  if is_implementado(t))

    budget_pais_local = budget.get(pais, {})
    budget_pais_intl  = budget.get(pais, {})

    sec_local = f"🏠 Ventas Locales — {pais_label}"
    sec_intl  = f"🌐 Ventas Internacionales — {pais_label}"

    tab_local = build_tab_panel(
        pid, f"{pid}_L", "local", "local", sec_local,
        local_tasks, budget_pais_local, pais, "local",
        mes_actual, anio_actual, is_active=True
    )
    tab_intl = build_tab_panel(
        pid, f"{pid}_I", "intl", "intl", sec_intl,
        intl_tasks, budget_pais_intl, pais, "intl",
        mes_actual, anio_actual, is_active=False
    )

    return f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>ZAS | {pais_label} — {datetime.now().strftime('%B %Y')}</title>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<style>{CSS}</style>
</head>
<body>
<div class="hdr">
  <div class="hdr-logo">ZAS</div>
  <div class="hdr-sep">/</div>
  <div class="hdr-title">Dashboard de Implementaciones</div>
  <div class="hdr-badge">{pais_label}</div>
  <div class="hdr-date">{today_str}</div>
</div>
<div id="{pid}">
<div class="mtabs">
  <div class="mtab local active" onclick="sw('{pid}','local')">
    🏠 Local &nbsp;<span style="opacity:.6;font-weight:400;font-size:12px">{cnt_local} cont · {fmt_usd(usd_local)}</span>
  </div>
  <div class="mtab intl" onclick="sw('{pid}','intl')">
    🌐 Internacional &nbsp;<span style="opacity:.6;font-weight:400;font-size:12px">{cnt_intl} cont · {fmt_usd(usd_intl)}</span>
  </div>
</div>
{tab_local}
{tab_intl}
</div>
<script>{JS}</script>
</body>
</html>"""

# ─── MAIN ──────────────────────────────────────────────────────────────────────
def main():
    auto = "--auto" in sys.argv
    now  = datetime.now()
    mes_actual  = now.month
    anio_actual = now.year

    print(f"\n{'='*55}")
    print(f"  ZAS Dashboards — {now.strftime('%d/%m/%Y %H:%M')}")
    print(f"{'='*55}\n")

    # 1. Autenticar
    uid = odoo_auth()

    # 2. Descargar datos
    tasks, orders = download_data(uid)

    # 3. Clasificar
    print("  Clasificando tareas...")
    classified = classify_tasks(tasks, orders)

    # 4. Cargar presupuesto
    print(f"  Cargando presupuesto: {BUDGET_FILE}")
    budget = load_budget(BUDGET_FILE)

    # 5. Generar HTMLs
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    output_map = {
        "argentina":    "argentina.html",
        "chile":        "chile.html",
        "colombia":     "colombia.html",
        "usa":          "usa.html",
        "peru":         "peru.html",
        "internacional":"internacional.html",
    }

    print("\n  Generando HTMLs...")
    for pais, filename in output_map.items():
        local_tasks = classified[pais]["local"]
        intl_tasks  = classified[pais]["intl"]
        print(f"    {pais}: {len(local_tasks)} local / {len(intl_tasks)} intl")

        html = generate_html(pais, local_tasks, intl_tasks, budget, mes_actual, anio_actual)
        path = os.path.join(OUTPUT_DIR, filename)
        with open(path, "w", encoding="utf-8") as f:
            f.write(html)
        print(f"    ✓ {path} ({len(html)//1024}KB)")

    # 6. Index
    _write_index(output_map, now)

    print(f"\n  ✅ {len(output_map)} dashboards generados en '{OUTPUT_DIR}/'")
    print(f"  Links: https://martinaboazzocid.github.io/avances_dashboard/{{pais}}.html\n")

    if not auto:
        input("  Presioná Enter para salir...")

def _write_index(output_map, now):
    items = "".join(
        f'<li><a href="{fname}">{pais.upper()}</a></li>'
        for pais, fname in output_map.items()
    )
    html = f"""<!DOCTYPE html>
<html lang="es">
<head><meta charset="UTF-8"><title>ZAS Dashboards</title>
<style>body{{background:#0d0f14;color:#e8eaf0;font-family:Inter,sans-serif;padding:40px;}}
a{{color:#4f8ef7;text-decoration:none;}}a:hover{{text-decoration:underline;}}
ul{{list-style:none;padding:0;}}li{{margin:10px 0;font-size:18px;}}</style>
</head><body>
<h1 style="color:#4f8ef7">ZAS Dashboards</h1>
<p style="color:#8b93a8">Actualizado: {now.strftime('%d/%m/%Y %H:%M')}</p>
<ul>{items}</ul>
</body></html>"""
    with open(os.path.join(OUTPUT_DIR, "index.html"), "w", encoding="utf-8") as f:
        f.write(html)

if __name__ == "__main__":
    main()
