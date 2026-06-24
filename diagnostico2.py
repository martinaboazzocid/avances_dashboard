import json, urllib.request, http.cookiejar, os
from collections import Counter, defaultdict

ODOO_URL  = os.environ["ODOO_URL"]
ODOO_DB   = os.environ["ODOO_DB"]
ODOO_USER = os.environ["ODOO_USER"]
ODOO_PASS = os.environ["ODOO_PASSWORD"]

_cj = http.cookiejar.CookieJar()
_opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(_cj))

def call(url, payload):
    data = json.dumps(payload).encode()
    req = urllib.request.Request(url, data=data, headers={"Content-Type":"application/json"}, method="POST")
    with _opener.open(req, timeout=60) as r:
        return json.loads(r.read())

# Auth
r = call(f"{ODOO_URL}/web/session/authenticate", {"jsonrpc":"2.0","method":"call","id":1,
    "params":{"db":ODOO_DB,"login":ODOO_USER,"password":ODOO_PASS}})
uid = r["result"]["uid"]
print(f"UID={uid}\n")

# Traer muestra de 20 subtareas con todos los campos clave
r2 = call(f"{ODOO_URL}/web/dataset/call_kw", {"jsonrpc":"2.0","method":"call","id":2,
    "params":{"model":"project.task","method":"search_read",
        "args":[[["parent_id","!=",False]]],
        "kwargs":{
            "fields":[
                "id","name","company_id","sale_order_id",
                "x_studio_related_field_8rl_1jhbqu80b",
                "x_studio_fecha_de_publicacin",
                "x_studio_fecha_limite_ops",
                "x_studio_related_field_4gg_1jckjh2i6",  # Subtotal copia
                "x_studio_related_field_4ia_1jc785nre",  # Total
                "x_studio_related_field_52s_1j0ff39s4",  # Importe
                "x_studio_related_field_66n_1jc788is6",  # Importe
                "x_studio_related_field_6o9_1jhbqk5st",  # Moneda
                "x_studio_sale_order_id_currency_id",    # Sales Order Currency
                "x_studio_subtotal",
            ],
            "limit":20,
            "context":{"allowed_company_ids":[1,2,3,4,5,6]}
        }
    }
})
tasks = r2.get("result",[])
print(f"=== MUESTRA 20 SUBTAREAS ===")
for t in tasks:
    print(f"\nid={t['id']} | {t['name'][:50]}")
    print(f"  company: {t.get('company_id')}")
    print(f"  sale_order_id: {t.get('sale_order_id')}")
    print(f"  pais_campana (8rl): {t.get('x_studio_related_field_8rl_1jhbqu80b')}")
    print(f"  fecha_pub: {t.get('x_studio_fecha_de_publicacin')}")
    print(f"  fecha_estimada: {t.get('x_studio_fecha_limite_ops')}")
    print(f"  subtotal(4gg): {t.get('x_studio_related_field_4gg_1jckjh2i6')}")
    print(f"  total(4ia): {t.get('x_studio_related_field_4ia_1jc785nre')}")
    print(f"  importe(52s): {t.get('x_studio_related_field_52s_1j0ff39s4')}")
    print(f"  importe(66n): {t.get('x_studio_related_field_66n_1jc788is6')}")
    print(f"  moneda(6o9): {t.get('x_studio_related_field_6o9_1jhbqk5st')}")
    print(f"  order_currency: {t.get('x_studio_sale_order_id_currency_id')}")
    print(f"  subtotal: {t.get('x_studio_subtotal')}")

# Ver una orden de venta con sus campos clave
if tasks and tasks[0].get('sale_order_id'):
    oid = tasks[0]['sale_order_id']
    if isinstance(oid, list): oid = oid[0]
    r3 = call(f"{ODOO_URL}/web/dataset/call_kw", {"jsonrpc":"2.0","method":"call","id":3,
        "params":{"model":"sale.order","method":"search_read",
            "args":[[["id","=",oid]]],
            "kwargs":{
                "fields":["id","name","amount_untaxed","currency_id","state",
                          "x_studio_campaas","x_studio_campaas_1",
                          "x_studio_bu","x_studio_bu_1"],
                "limit":1,
                "context":{"allowed_company_ids":[1,2,3,4,5,6]}
            }
        }
    })
    orders = r3.get("result",[])
    if orders:
        o = orders[0]
        print(f"\n=== ORDEN DE VENTA id={o['id']} ===")
        print(f"  name: {o.get('name')}")
        print(f"  amount_untaxed: {o.get('amount_untaxed')}")
        print(f"  currency_id: {o.get('currency_id')}")
        print(f"  state: {o.get('state')}")
        print(f"  campaas: {o.get('x_studio_campaas')}")
        print(f"  campaas_1: {o.get('x_studio_campaas_1')}")
        print(f"  bu: {o.get('x_studio_bu')}")
        print(f"  bu_1: {o.get('x_studio_bu_1')}")

# Valores únicos del campo pais en tareas
r4 = call(f"{ODOO_URL}/web/dataset/call_kw", {"jsonrpc":"2.0","method":"call","id":4,
    "params":{"model":"project.task","method":"search_read",
        "args":[[["parent_id","!=",False]]],
        "kwargs":{
            "fields":["x_studio_related_field_8rl_1jhbqu80b"],
            "limit":5000,
            "context":{"allowed_company_ids":[1,2,3,4,5,6]}
        }
    }
})
pais_vals = Counter(str(t.get("x_studio_related_field_8rl_1jhbqu80b")) for t in r4.get("result",[]))
print(f"\n=== VALORES ÚNICOS campo pais_campana en tareas (top 20) ===")
for v, n in pais_vals.most_common(20):
    print(f"  {n:5d}x  {v}")
