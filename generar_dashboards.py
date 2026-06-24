#!/usr/bin/env python3
"""
ZAS Talents — Generador de Dashboards de Implementaciones
=========================================================
Descarga datos de Odoo via JSON-RPC, filtra, y genera HTMLs por país.
Salida: docs/argentina.html, docs/chile.html, docs/colombia.html,
        docs/usa.html, docs/peru.html, docs/internacional.html

USO:
    python generar_dashboards.py          # interactivo
    python generar_dashboards.py --auto   # sin prompts (Task Scheduler / GitHub Actions)

CONFIGURACIÓN:
    Variables de entorno (o archivo .env local):
        ODOO_URL       = https://zas-talent.odoo.com
        ODOO_DB        = zas-talent
        ODOO_USER      = martu@zastalents.com
        ODOO_PASSWORD  = tu_password

REQUISITOS:
    pip install requests openpyxl
"""

import sys
import os
import json
import requests
from datetime import datetime, date
from pathlib import Path
import openpyxl

# ─────────────────────────────────────────────────────────────────────────────
# CONFIG — todas las credenciales vienen de variables de entorno
# ─────────────────────────────────────────────────────────────────────────────
ODOO_URL      = os.environ.get("ODOO_URL",      "https://zas-talent.odoo.com")
ODOO_DB       = os.environ.get("ODOO_DB",       "zas-talent")
ODOO_USER     = os.environ.get("ODOO_USER",     "martina.boazzo@zastalents.com")
ODOO_PASSWORD = os.environ.get("ODOO_PASSWORD", "")

AUTO_MODE = "--auto" in sys.argv

# Directorio de salida: docs/ para GitHub Pages
OUTPUT_DIR = Path("docs")
OUTPUT_DIR.mkdir(exist_ok=True)

# Presupuesto — archivo local en el repo
BUDGET_FILE = Path("Base_Regional__2_.xlsx")

# Tipos de cambio (actualizables)
TC = {
    "ARS": 1100,
    "CLP": 940,
    "CLP_BG": 890,   # tipo de cambio usado en el presupuesto de Chile
    "COP": 4250,
    "PEN": 3.72,
    "USD": 1.0,
}

# Compañías Odoo
# 1=ARG, 2=Chile, 3=México (excluir), 4=Perú, 5=USA, 6=Colombia
COMPANY_IDS_ALLOWED = [1, 2, 3, 4, 5, 6]  # se incluye 3 en la descarga, se filtra en Python

# Mapas de traducción de campos selection de Odoo
SEL_CAMPAAS = {
    "Campanas": "Campañas Internacionales",
    "Campañas": "Campañas Internacionales",
}

SEL_CAMPAAS_1 = {
    "Campanas": "Campañas Argentinas",
    "Campañas": "Campañas Argentinas",
    "Chile":    "Campañas Chile",
    "Colombia": "Campañas Colombia",
    "USA":      "US Campaigns",
    "Peru":     "Campañas Peruanas",
    "Perú":     "Campañas Peruanas",
}

# Mapeo BU → país para registros internacionales
BU_PAIS = {
    "ZAS ARGENTINA": "argentina",
    "ZAS CHILE":     "chile",
    "ZAS COLOMBIA":  "colombia",
    "ZAS USA":       "usa",
    "ZAS PERU":      "peru",
    "ZAS PERÚ":      "peru",
}

# Presupuesto: filas por país (1-indexed como en Excel)
BUDGET_ROWS = {
    "argentina": {"sheet": "BG ARG", "comercial": 18, "artistico": 27, "regional": 34, "internacional": 42},
    "chile":     {"sheet": "BG CHI", "comercial": 17, "artistico": 25, "regional": 32, "internacional": 42},
    "colombia":  {"sheet": "BG COL", "comercial": 21, "artistico": 30, "regional": 37, "internacional": 46},
    "usa":       {"sheet": "BG USA", "comercial": 17, "artistico": 27, "regional": 34, "internacional": 43},
}
BUDGET_COL_OFFSET = 3  # col_index = mes + BUDGET_COL_OFFSET - 1


# ─────────────────────────────────────────────────────────────────────────────
# AUTENTICACIÓN ODOO
# ─────────────────────────────────────────────────────────────────────────────

# Sesión compartida para persistir cookies entre llamadas
_session = requests.Session()
_session.headers.update({"Content-Type": "application/json"})


def odoo_authenticate():
    """Autentica en Odoo via /web/session/authenticate y retorna uid."""
    resp = _session.post(
        f"{ODOO_URL}/web/session/authenticate",
        json={
            "jsonrpc": "2.0",
            "method": "call",
            "id": 1,
            "params": {
                "db":       ODOO_DB,
                "login":    ODOO_USER,
                "password": ODOO_PASSWORD,
            },
        },
        timeout=30,
    )
    resp.raise_for_status()
    data = resp.json()
    if data.get("error"):
        raise RuntimeError(f"Autenticación fallida: {data['error']}")
    uid = data.get("result", {}).get("uid")
    if not uid:
        raise RuntimeError(f"Autenticación fallida: uid no recibido. Respuesta: {data}")
    print(f"✓ Autenticado como uid={uid}")
    return uid


def odoo_call(uid, model, method, args=None, kwargs=None):
    """Llamada genérica a Odoo JSON-RPC con contexto multi-company."""
    if args is None:
        args = []
    if kwargs is None:
        kwargs = {}
    kwargs.setdefault("context", {})
    kwargs["context"]["allowed_company_ids"] = COMPANY_IDS_ALLOWED

    resp = _session.post(
        f"{ODOO_URL}/web/dataset/call_kw",
        json={
            "jsonrpc": "2.0",
            "method": "call",
            "id": 1,
            "params": {
                "model":  model,
                "method": method,
                "args":   args,
                "kwargs": kwargs,
            },
        },
        timeout=120,
    )
    resp.raise_for_status()
    data = resp.json()
    if "error" in data:
        raise RuntimeError(f"Error Odoo ({model}.{method}): {data['error']}")
    return data["result"]


# ─────────────────────────────────────────────────────────────────────────────
# DESCARGA DE DATOS
# ─────────────────────────────────────────────────────────────────────────────
SUBTASK_FIELDS = [
    "id", "name", "project_id", "parent_id", "user_ids",
    "date_deadline", "date_last_stage_update",
    "x_studio_related_field_7v0_1jidluau0",   # Responsable de proyecto (implementador) — confirmado
    "sale_order_id", "company_id",
    "stage_id", "state",
]

ORDER_FIELDS = [
    "id", "name", "company_id", "amount_untaxed",
    "currency_id", "state", "date_order",
]

ORDER_LINE_FIELDS = [
    "id", "order_id", "product_id", "price_subtotal",
    "currency_id", "qty_delivered",
]


def discover_fields(uid):
    """Diagnóstico: imprime todos los campos x_ de project.task."""
    print("🔍 Descubriendo campos custom en project.task...")
    all_fields = odoo_call(uid, "project.task", "fields_get", args=[], kwargs={"attributes": ["string", "type"]})
    studio_fields = {k: v for k, v in all_fields.items() if k.startswith("x_")}
    for fname, finfo in sorted(studio_fields.items()):
        print(f"  {fname} ({finfo.get('type','?')}): {finfo.get('string','')}")
    print(f"  → Total campos custom: {len(studio_fields)}")


def download_all_data(uid):
    """Descarga subtareas y órdenes de venta de Odoo."""
    print("↓ Descargando subtareas...")
    subtasks = []
    offset = 0
    batch = 1000
    while True:
        chunk = odoo_call(
            uid, "project.task", "search_read",
            args=[[["parent_id", "!=", False]]],
            kwargs={
                "fields": SUBTASK_FIELDS,
                "limit":  batch,
                "offset": offset,
                "order":  "id asc",
            },
        )
        subtasks.extend(chunk)
        print(f"  subtareas: {len(subtasks)}", end="\r")
        if len(chunk) < batch:
            break
        offset += batch
    print(f"  ✓ {len(subtasks)} subtareas descargadas")

    # ── DIAGNÓSTICO: pedir un registro completo sin filtrar campos ──
    print("DIAG — campos x_ con valor en primera subtarea (sin filtro de fields):")
    sample = odoo_call(
        uid, "project.task", "search_read",
        args=[[["parent_id", "!=", False]]],
        kwargs={"limit": 1, "offset": 0, "order": "id desc"},
    )
    if sample:
        t = sample[0]
        print(f"  Tarea id={t['id']} | {t.get('name','')[:50]!r}")
        for k, v in sorted(t.items()):
            if k.startswith("x_") and v not in (False, None, [], ""):
                print(f"    {k} = {v!r}")
    # ── FIN DIAGNÓSTICO ──

    print("↓ Descargando órdenes de venta...")
    orders = []
    offset = 0
    while True:
        chunk = odoo_call(
            uid, "sale.order", "search_read",
            args=[[["state", "in", ["sale", "done"]]]],
            kwargs={
                "fields": ORDER_FIELDS,
                "limit":  batch,
                "offset": offset,
                "order":  "id asc",
            },
        )
        orders.extend(chunk)
        print(f"  órdenes: {len(orders)}", end="\r")
        if len(chunk) < batch:
            break
        offset += batch
    print(f"  ✓ {len(orders)} órdenes descargadas")

    return subtasks, orders


# ─────────────────────────────────────────────────────────────────────────────
# FILTRADO Y CLASIFICACIÓN
# ─────────────────────────────────────────────────────────────────────────────
def get_company_id(record):
    cid = record.get("company_id")
    if isinstance(cid, (list, tuple)):
        return cid[0]
    return cid


def translate_sel(field_value, translation_map):
    if not field_value:
        return ""
    if isinstance(field_value, (list, tuple)):
        field_value = field_value[0] if field_value else ""
    return translation_map.get(str(field_value), str(field_value))


def classify_subtask(task):
    """
    Retorna (pais, tab) donde:
      pais ∈ {argentina, chile, colombia, usa, peru, internacional, None}
      tab  ∈ {local, internacional}
    None = descartar
    """
    cid = get_company_id(task)
    if cid == 3:
        return None, None

    label_1 = translate_sel(task.get("x_studio_related_field_8rl_1jhbqu80b"),  SEL_CAMPAAS)
    label_2 = translate_sel(task.get("x_studio_related_field_8pi_1jhbqk5st"), SEL_CAMPAAS_1)
    pais_final = (label_1 + label_2).strip()

    LOCAL_MAP = {
        "Campañas Argentinas": "argentina",
        "Campañas Chile":      "chile",
        "Campañas Colombia":   "colombia",
        "US Campaigns":        "usa",
        "Campañas Peruanas":   "peru",
    }

    if pais_final in LOCAL_MAP:
        return LOCAL_MAP[pais_final], "local"

    if "Internacionales" in pais_final or "Internacional" in pais_final:
        # BU (x_studio_bu_1) ya no existe — inferir país desde company_id
        cid_to_pais = {1: "argentina", 2: "chile", 4: "peru", 5: "usa", 6: "colombia"}
        pais = cid_to_pais.get(cid)
        if pais:
            return pais, "internacional"
        return None, None

    return None, None


def filter_and_classify(subtasks, orders):
    """Aplica filtros post-descarga y clasifica registros por país y tab."""
    valid_orders = {}
    for o in orders:
        if get_company_id(o) == 3:
            continue
        valid_orders[o["id"]] = o

    classified = {}
    for pais in ["argentina", "chile", "colombia", "usa", "peru"]:
        classified[pais] = {"local": [], "internacional": []}

    for task in subtasks:
        pais, tab = classify_subtask(task)
        if not pais or not tab:
            continue
        oid = task.get("sale_order_id")
        if isinstance(oid, (list, tuple)):
            oid = oid[0] if oid else None
        task["_order"] = valid_orders.get(oid) if oid else None
        task["_pais"] = pais
        task["_tab"] = tab
        classified[pais][tab].append(task)

    for pais in classified:
        for tab in classified[pais]:
            n = len(classified[pais][tab])
            print(f"  {pais}/{tab}: {n} registros")

    return classified


# ─────────────────────────────────────────────────────────────────────────────
# LECTURA DE PRESUPUESTO
# ─────────────────────────────────────────────────────────────────────────────
def load_budget():
    if not BUDGET_FILE.exists():
        print(f"⚠ Archivo de presupuesto no encontrado: {BUDGET_FILE}")
        return {}

    wb = openpyxl.load_workbook(BUDGET_FILE, data_only=True)
    budget = {}

    for pais, cfg in BUDGET_ROWS.items():
        ws = wb[cfg["sheet"]]
        tc = TC["CLP_BG"] if pais == "chile" else 1.0

        budget[pais] = {}
        for category in ["comercial", "artistico", "regional", "internacional"]:
            row = cfg[category]
            monthly = []
            for mes in range(1, 13):
                col = mes + BUDGET_COL_OFFSET - 1
                val = ws.cell(row=row, column=col).value or 0
                try:
                    val_usd = float(val) / tc
                except (TypeError, ValueError):
                    val_usd = 0.0
                monthly.append(val_usd)
            budget[pais][category] = monthly

    return budget


# ─────────────────────────────────────────────────────────────────────────────
# MÉTRICAS
# ─────────────────────────────────────────────────────────────────────────────
def get_amount_usd(order):
    if not order:
        return 0.0
    amount = order.get("amount_untaxed", 0) or 0
    currency = order.get("currency_id")
    if isinstance(currency, (list, tuple)):
        currency = currency[1] if len(currency) > 1 else ""
    currency = str(currency).upper()

    tc_map = {"ARS": TC["ARS"], "CLP": TC["CLP"], "COP": TC["COP"], "PEN": TC["PEN"], "USD": 1.0}
    tc = tc_map.get(currency, 1.0)
    return float(amount) / tc if tc else 0.0


def compute_kpis(tasks, mes_actual=None):
    if mes_actual is None:
        mes_actual = date.today().month

    implementado_mes = 0.0
    implementado_ytd = 0.0
    pipeline = 0.0
    campanas_set = set()
    talentos_set = set()
    contenidos = []

    today = date.today()

    for t in tasks:
        order = t.get("_order")
        monto = get_amount_usd(order)

        name = t.get("name", "")
        if "(" in name:
            talento = name[:name.index("(")].strip()
            contenido = name[name.index("(") + 1:].replace(")", "").strip()
        else:
            talento = name.strip()
            contenido = ""

        talentos_set.add(talento)

        stage = t.get("stage_id")
        if isinstance(stage, (list, tuple)):
            stage_name = str(stage[1]).lower() if len(stage) > 1 else ""
        else:
            stage_name = ""

        implementado = "implementad" in stage_name or "done" in stage_name or "realiz" in stage_name

        deadline = t.get("date_deadline")
        if deadline:
            try:
                dl = date.fromisoformat(str(deadline)[:10])
                dl_mes = dl.month
                dl_anio = dl.year
            except ValueError:
                dl_mes, dl_anio = mes_actual, today.year
        else:
            dl_mes, dl_anio = mes_actual, today.year

        if implementado:
            if dl_anio == today.year:
                implementado_ytd += monto
            if dl_mes == mes_actual and dl_anio == today.year:
                implementado_mes += monto
                if order:
                    campanas_set.add(order.get("id"))
                contenidos.append({"talento": talento, "contenido": contenido, "monto": monto})
        else:
            pipeline += monto

    return {
        "implementado_mes": implementado_mes,
        "implementado_ytd": implementado_ytd,
        "pipeline": pipeline,
        "campanas": len(campanas_set),
        "talentos": len(talentos_set),
        "contenidos": contenidos,
    }


# ─────────────────────────────────────────────────────────────────────────────
# GENERACIÓN HTML
# ─────────────────────────────────────────────────────────────────────────────
PAIS_LABELS = {
    "argentina":    "🇦🇷 Argentina",
    "chile":        "🇨🇱 Chile",
    "colombia":     "🇨🇴 Colombia",
    "usa":          "🇺🇸 USA",
    "peru":         "🇵🇪 Perú",
    "internacional": "🌐 Internacional",
}

MONTHS_ES = ["Ene", "Feb", "Mar", "Abr", "May", "Jun",
             "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]


def fmt_usd(v):
    return f"${v:,.0f}"


def pct_color(pct):
    if pct >= 100:
        return "#f0fff0", "#22863a"
    if pct >= 90:
        return "#e6f4ea", "#1a7f37"
    if pct >= 60:
        return "#fff8e1", "#b45309"
    return "#fff0f0", "#d1242f"


def build_budget_table(kpis, budget_pais, mes_actual):
    if not budget_pais:
        return "<p>Sin datos de presupuesto.</p>"

    rows_html = ""
    categories = ["comercial", "artistico", "regional"]
    labels = {"comercial": "Comercial", "artistico": "Artístico", "regional": "Regional"}

    for cat in categories:
        monthly = budget_pais.get(cat, [0] * 12)
        presup_mes = monthly[mes_actual - 1] if mes_actual <= 12 else 0
        presup_ytd = sum(monthly[:mes_actual])
        real_mes = kpis["implementado_mes"] / len(categories)
        real_ytd = kpis["implementado_ytd"] / len(categories)

        pct_mes = (real_mes / presup_mes * 100) if presup_mes > 0 else 0
        pct_ytd = (real_ytd / presup_ytd * 100) if presup_ytd > 0 else 0

        bg_m, fg_m = pct_color(pct_mes)
        bg_y, fg_y = pct_color(pct_ytd)

        rows_html += f"""
        <tr>
          <td>{labels[cat]}</td>
          <td>{fmt_usd(presup_mes)}</td>
          <td>{fmt_usd(real_mes)}</td>
          <td style="background:{bg_m};color:{fg_m};font-weight:600">{pct_mes:.0f}%</td>
          <td>{fmt_usd(presup_ytd)}</td>
          <td>{fmt_usd(real_ytd)}</td>
          <td style="background:{bg_y};color:{fg_y};font-weight:600">{pct_ytd:.0f}%</td>
        </tr>"""

    return f"""
    <table class="budget-table">
      <thead>
        <tr>
          <th>Categoría</th>
          <th>Presup. Mes</th><th>Real Mes</th><th>%</th>
          <th>Presup. YTD</th><th>Real YTD</th><th>%</th>
        </tr>
      </thead>
      <tbody>{rows_html}</tbody>
    </table>"""


def build_historial_section(tasks):
    impl_data = {}
    for t in tasks:
        impl = t.get("x_studio_related_field_7v0_1jidluau0") or "Sin asignar"
        if isinstance(impl, (list, tuple)):
            impl = impl[1] if len(impl) > 1 else "Sin asignar"
        order = t.get("_order")
        monto = get_amount_usd(order)
        if impl not in impl_data:
            impl_data[impl] = {"monto": 0.0, "contenidos": 0, "campanas": set()}
        impl_data[impl]["monto"] += monto
        impl_data[impl]["contenidos"] += 1
        if order:
            impl_data[impl]["campanas"].add(order.get("id"))

    if not impl_data:
        return "<p>Sin datos históricos.</p>"

    rows = ""
    for impl, d in sorted(impl_data.items(), key=lambda x: -x[1]["monto"]):
        avg = d["monto"] / d["contenidos"] if d["contenidos"] > 0 else 0
        rows += f"""
        <tr>
          <td>{impl}</td>
          <td>{fmt_usd(d['monto'])}</td>
          <td>{d['contenidos']}</td>
          <td>{len(d['campanas'])}</td>
          <td>{fmt_usd(avg)}</td>
        </tr>"""

    return f"""
    <table class="data-table">
      <thead>
        <tr>
          <th>Implementador</th>
          <th>Total USD</th>
          <th>Contenidos</th>
          <th>Campañas</th>
          <th>Avg/Contenido</th>
        </tr>
      </thead>
      <tbody>{rows}</tbody>
    </table>"""


def build_pipeline_section(tasks):
    pipeline_tasks = []
    for t in tasks:
        stage = t.get("stage_id")
        if isinstance(stage, (list, tuple)):
            stage_name = str(stage[1]).lower() if len(stage) > 1 else ""
        else:
            stage_name = ""
        implementado = "implementad" in stage_name or "done" in stage_name
        if not implementado:
            pipeline_tasks.append(t)

    if not pipeline_tasks:
        return "<p>Sin pipeline activo.</p>"

    rows = ""
    for t in sorted(pipeline_tasks, key=lambda x: x.get("date_deadline") or "9999"):
        name = t.get("name", "")
        deadline = t.get("date_deadline", "—")
        order = t.get("_order")
        monto = get_amount_usd(order)
        stage = t.get("stage_id")
        stage_name = stage[1] if isinstance(stage, (list, tuple)) and len(stage) > 1 else "—"
        rows += f"""
        <tr>
          <td>{name}</td>
          <td>{stage_name}</td>
          <td>{deadline or '—'}</td>
          <td>{fmt_usd(monto)}</td>
        </tr>"""

    return f"""
    <table class="data-table">
      <thead>
        <tr><th>Subtarea</th><th>Etapa</th><th>Deadline</th><th>Monto USD</th></tr>
      </thead>
      <tbody>{rows}</tbody>
    </table>"""


def build_accionables(kpis, pais, tab):
    items = []
    impl = kpis["implementado_mes"]
    pipe = kpis["pipeline"]
    camp = kpis["campanas"]

    if impl == 0:
        items.append("⚠️ No hay implementaciones registradas este mes. Verificar si hay contenidos pendientes de marcar.")
    if pipe > impl * 2:
        items.append(f"🔔 El pipeline ({fmt_usd(pipe)}) supera en más del doble lo implementado este mes. Acelerar cierres.")
    if camp == 0 and impl > 0:
        items.append("🔎 Hay implementaciones sin campañas asociadas. Revisar asignación de órdenes de venta.")
    if kpis["talentos"] > 0 and impl / kpis["talentos"] < 500:
        items.append("📊 Promedio por talento bajo. Evaluar concentración de trabajo en pocos implementadores.")
    if not items:
        items.append("✅ KPIs dentro de parámetros normales.")

    return "<ul>" + "".join(f"<li>{i}</li>" for i in items) + "</ul>"


CSS = """
* { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
       background: #f5f7fa; color: #1a1a2e; }
.header { background: linear-gradient(135deg, #1a1a2e 0%, #16213e 100%);
          color: white; padding: 24px 32px; }
.header h1 { font-size: 1.6rem; font-weight: 700; }
.header .meta { font-size: 0.85rem; opacity: 0.7; margin-top: 4px; }
.tab-bar { display: flex; background: white; border-bottom: 2px solid #e2e8f0;
           padding: 0 24px; gap: 4px; }
.tab-btn { padding: 14px 20px; border: none; background: none; cursor: pointer;
           font-size: 0.95rem; color: #64748b; border-bottom: 3px solid transparent;
           margin-bottom: -2px; transition: all 0.2s; font-weight: 500; }
.tab-btn:hover { color: #1a1a2e; }
.tab-btn.active { color: #1a1a2e; border-bottom-color: #6366f1; font-weight: 700; }
.content { padding: 28px 32px; max-width: 1200px; margin: 0 auto; }
.tab-content { display: none; }
.tab-content.active { display: block; }
.kpi-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 16px; margin-bottom: 28px; }
.kpi-card { background: white; border-radius: 12px; padding: 20px;
            box-shadow: 0 1px 4px rgba(0,0,0,.07); }
.kpi-card .label { font-size: 0.78rem; color: #64748b; text-transform: uppercase;
                   letter-spacing: .05em; margin-bottom: 8px; }
.kpi-card .value { font-size: 1.7rem; font-weight: 800; color: #1a1a2e; }
.kpi-card .sub { font-size: 0.8rem; color: #94a3b8; margin-top: 4px; }
.section-title { font-size: 1rem; font-weight: 700; color: #1e293b;
                 margin: 24px 0 12px; padding-left: 10px;
                 border-left: 4px solid #6366f1; }
.data-table, .budget-table { width: 100%; border-collapse: collapse;
                              background: white; border-radius: 10px;
                              overflow: hidden; box-shadow: 0 1px 4px rgba(0,0,0,.07); }
.data-table th, .budget-table th { background: #1e293b; color: white;
                                    padding: 10px 14px; text-align: left;
                                    font-size: 0.82rem; }
.data-table td, .budget-table td { padding: 10px 14px; border-top: 1px solid #f1f5f9;
                                    font-size: 0.88rem; }
.data-table tr:hover td { background: #f8fafc; }
.sub-tab-bar { display: flex; gap: 4px; margin-bottom: 20px; flex-wrap: wrap; }
.sub-tab-btn { padding: 8px 16px; border-radius: 20px; border: 1px solid #e2e8f0;
               background: white; cursor: pointer; font-size: 0.85rem;
               color: #64748b; transition: all 0.2s; }
.sub-tab-btn.active { background: #6366f1; color: white; border-color: #6366f1; }
.sub-content { display: none; }
.sub-content.active { display: block; }
ul { padding-left: 20px; }
ul li { margin-bottom: 8px; font-size: 0.9rem; line-height: 1.5; }
.badge { display: inline-block; padding: 2px 8px; border-radius: 10px;
         font-size: 0.75rem; font-weight: 600; }
.badge-local { background: #dbeafe; color: #1d4ed8; }
.badge-int   { background: #ede9fe; color: #6d28d9; }
"""

JS = """
function showTab(tabId, btn) {
  document.querySelectorAll('.tab-content').forEach(t => t.classList.remove('active'));
  document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
  document.getElementById(tabId).classList.add('active');
  btn.classList.add('active');
}
function showSubTab(parentId, subId, btn) {
  document.querySelectorAll('#' + parentId + ' .sub-content').forEach(t => t.classList.remove('active'));
  document.querySelectorAll('#' + parentId + ' .sub-tab-btn').forEach(b => b.classList.remove('active'));
  document.getElementById(subId).classList.add('active');
  btn.classList.add('active');
}
"""


def build_tab_content(tab_id, tasks, budget_pais, pais, tab, mes_actual):
    kpis = compute_kpis(tasks, mes_actual)
    badge_class = "badge-local" if tab == "local" else "badge-int"
    badge_label = "Local" if tab == "local" else "Internacional"

    presup_html   = build_budget_table(kpis, budget_pais, mes_actual)
    pipeline_html = build_pipeline_section(tasks)
    historial_html = build_historial_section(tasks)
    accionables_html = build_accionables(kpis, pais, tab)

    sub_tabs = [
        ("Implementado", "impl",   f"""
            <div class='kpi-grid'>
              <div class='kpi-card'><div class='label'>Implementado Mes</div>
                <div class='value'>{fmt_usd(kpis['implementado_mes'])}</div></div>
              <div class='kpi-card'><div class='label'>Implementado YTD</div>
                <div class='value'>{fmt_usd(kpis['implementado_ytd'])}</div></div>
              <div class='kpi-card'><div class='label'>Campañas</div>
                <div class='value'>{kpis['campanas']}</div></div>
              <div class='kpi-card'><div class='label'>Talentos</div>
                <div class='value'>{kpis['talentos']}</div></div>
            </div>"""),
        ("Presupuesto",  "presup", presup_html),
        ("Pipeline",     "pipe",   pipeline_html),
        ("Historial",    "hist",   historial_html),
        ("Accionables",  "acc",    accionables_html),
    ]

    sub_bar = ""
    sub_contents = ""
    for i, (label, key, content) in enumerate(sub_tabs):
        sid = f"{tab_id}_{key}"
        active_class = "active" if i == 0 else ""
        sub_bar += f'<button class="sub-tab-btn {active_class}" onclick="showSubTab(\'{tab_id}\', \'{sid}\', this)">{label}</button>'
        sub_contents += f'<div class="sub-content {active_class}" id="{sid}">{content}</div>'

    return f"""
    <div class="tab-content" id="{tab_id}">
      <span class="badge {badge_class}" style="margin-bottom:16px;display:inline-block">{badge_label}</span>
      <div class="sub-tab-bar">{sub_bar}</div>
      {sub_contents}
    </div>"""


def generate_html(pais, local_tasks, int_tasks, budget, mes_actual):
    label = PAIS_LABELS.get(pais, pais.title())
    fecha = datetime.now().strftime("%d/%m/%Y %H:%M")
    budget_pais = budget.get(pais, {})

    local_html = build_tab_content(
        f"tab_{pais}_local", local_tasks, budget_pais, pais, "local", mes_actual
    )
    int_html = build_tab_content(
        f"tab_{pais}_int", int_tasks, budget_pais, pais, "internacional", mes_actual
    )

    return f"""<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Dashboard {label} — ZAS Talents</title>
  <style>{CSS}</style>
</head>
<body>
  <div class="header">
    <h1>{label} — Dashboard de Implementaciones</h1>
    <div class="meta">ZAS Talents · Actualizado: {fecha} (ARG) · Cifras en USD</div>
  </div>

  <div class="tab-bar">
    <button class="tab-btn active" onclick="showTab('tab_{pais}_local', this)">🏠 Local</button>
    <button class="tab-btn" onclick="showTab('tab_{pais}_int', this)">🌐 Internacional</button>
  </div>

  <div class="content">
    {local_html}
    {int_html}
  </div>

  <script>{JS}</script>
</body>
</html>"""


# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────
def main():
    print("=" * 60)
    print("ZAS Talents — Generador de Dashboards")
    print("=" * 60)

    if not ODOO_PASSWORD:
        print("✗ ODOO_PASSWORD no configurado. Setear variable de entorno.")
        sys.exit(1)

    uid = odoo_authenticate()
    subtasks, orders = download_all_data(uid)

    print("\nClasificando registros...")
    classified = filter_and_classify(subtasks, orders)

    print("\nCargando presupuesto...")
    budget = load_budget()

    mes_actual = date.today().month
    print(f"\nGenerando dashboards (mes {mes_actual})...")

    paises = ["argentina", "chile", "colombia", "usa", "peru"]
    fecha  = datetime.now().strftime("%d/%m/%Y %H:%M")
    for pais in paises:
        local_tasks = classified[pais]["local"]
        int_tasks   = classified[pais]["internacional"]
        html = generate_html(pais, local_tasks, int_tasks, budget, mes_actual)
        outfile = OUTPUT_DIR / f"{pais}.html"
        outfile.write_text(html, encoding="utf-8")
        print(f"  ✓ docs/{pais}.html ({len(local_tasks)} local, {len(int_tasks)} int)")

    build_index(paises, fecha=fecha)
    print(f"\n✅ Dashboards generados en {OUTPUT_DIR}/")

    if not AUTO_MODE:
        input("\nPresioná Enter para salir...")


def build_index(paises, fecha):
    cards = ""
    for pais in paises:
        label = PAIS_LABELS[pais]
        cards += f"""
        <a href="{pais}.html" class="card">
          <div class="card-label">{label}</div>
          <div class="card-sub">Ver dashboard →</div>
        </a>"""

    html = f"""<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>ZAS Talents — Dashboards</title>
  <style>
    body {{ font-family: -apple-system, sans-serif; background: #f5f7fa;
            display: flex; flex-direction: column; align-items: center;
            justify-content: center; min-height: 100vh; margin: 0; }}
    h1 {{ font-size: 1.8rem; color: #1a1a2e; margin-bottom: 6px; }}
    .meta {{ color: #94a3b8; font-size: 0.85rem; margin-bottom: 36px; }}
    .grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
             gap: 16px; max-width: 900px; width: 90%; }}
    .card {{ background: white; border-radius: 14px; padding: 24px;
             text-decoration: none; color: inherit;
             box-shadow: 0 2px 8px rgba(0,0,0,.08); transition: transform .15s;
             text-align: center; }}
    .card:hover {{ transform: translateY(-3px); box-shadow: 0 4px 16px rgba(0,0,0,.12); }}
    .card-label {{ font-size: 1.3rem; font-weight: 700; margin-bottom: 6px; }}
    .card-sub {{ font-size: 0.82rem; color: #6366f1; }}
  </style>
</head>
<body>
  <h1>ZAS Talents · Dashboards</h1>
  <div class="meta">Actualizado: {fecha} (ARG)</div>
  <div class="grid">{cards}</div>
</body>
</html>"""

    (OUTPUT_DIR / "index.html").write_text(html, encoding="utf-8")
    print("  ✓ docs/index.html")


if __name__ == "__main__":
    main()
