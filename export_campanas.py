import json, urllib.request, http.cookiejar, os, csv

ODOO_URL  = os.environ["ODOO_URL"]
ODOO_DB   = os.environ["ODOO_DB"]
ODOO_USER = os.environ["ODOO_USER"]
ODOO_PASS = os.environ["ODOO_PASSWORD"]

_cj = http.cookiejar.CookieJar()
_opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(_cj))

def call(url, payload):
    data = json.dumps(payload).encode()
    req = urllib.request.Request(url, data=data, headers={"Content-Type":"application/json"}, method="POST")
    with _opener.open(req, timeout=120) as r:
        return json.loads(r.read())

r = call(f"{ODOO_URL}/web/session/authenticate", {"jsonrpc":"2.0","method":"call","id":1,
    "params":{"db":ODOO_DB,"login":ODOO_USER,"password":ODOO_PASS}})
uid = r["result"]["uid"]
print(f"UID={uid}")

# Órdenes con campaas_1 = 'Campañas'
r2 = call(f"{ODOO_URL}/web/dataset/call_kw", {"jsonrpc":"2.0","method":"call","id":2,
    "params":{"model":"sale.order","method":"search_read",
        "args":[[["x_studio_campaas_1","=","Campañas"],["state","in",["sale","done"]]]],
        "kwargs":{
            "fields":["id","name","company_id","x_studio_campaas","x_studio_campaas_1",
                      "x_studio_bu","x_studio_bu_1","currency_id","amount_untaxed"],
            "limit":2000,
            "context":{"allowed_company_ids":[1,2,3,4,5,6]}
        }
    }
})
orders = r2.get("result", [])
order_map = {o["id"]: o for o in orders}
print(f"Órdenes con campaas_1='Campañas': {len(orders)}")

# Subtareas de esas órdenes
order_ids = [o["id"] for o in orders]
tasks = []
for i in range(0, len(order_ids), 500):
    chunk = order_ids[i:i+500]
    r3 = call(f"{ODOO_URL}/web/dataset/call_kw", {"jsonrpc":"2.0","method":"call","id":3+i,
        "params":{"model":"project.task","method":"search_read",
            "args":[[["parent_id","!=",False],["sale_order_id","in",chunk]]],
            "kwargs":{
                "fields":["id","name","company_id","sale_order_id",
                          "x_studio_fecha_de_publicacin","x_studio_fecha_limite_ops",
                          "x_studio_related_field_52s_1j0ff39s4",
                          "x_studio_related_field_6o9_1jhbqk5st"],
                "limit":5000,
                "context":{"allowed_company_ids":[1,2,3,4,5,6]}
            }
        }
    })
    tasks.extend(r3.get("result", []))
print(f"Subtareas: {len(tasks)}")

# Escribir CSV
with open("campanas_muestra.csv", "w", newline="", encoding="utf-8") as f:
    w = csv.writer(f)
    w.writerow(["task_id","task_name","company","order_name","campaas","campaas_1",
                "bu","bu_1","currency","amount_untaxed","fecha_pub","fecha_estimada","importe_linea"])
    for t in tasks:
        o = order_map.get(t["sale_order_id"][0] if isinstance(t["sale_order_id"],list) else t["sale_order_id"], {})
        def v(x): return x[1] if isinstance(x,list) else (x or "")
        w.writerow([
            t["id"], t["name"], v(t.get("company_id")), v(t.get("sale_order_id")),
            o.get("x_studio_campaas",""), o.get("x_studio_campaas_1",""),
            o.get("x_studio_bu",""), o.get("x_studio_bu_1",""),
            v(o.get("currency_id","")), o.get("amount_untaxed",""),
            t.get("x_studio_fecha_de_publicacin",""), t.get("x_studio_fecha_limite_ops",""),
            t.get("x_studio_related_field_52s_1j0ff39s4",""),
        ])
print(f"Guardado campanas_muestra.csv con {len(tasks)} filas")
