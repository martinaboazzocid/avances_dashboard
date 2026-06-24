import json, urllib.request, http.cookiejar, os

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
            "fields":["id","name","company_id","x_studio_campaas","x_studio_campaas_1","x_studio_bu","x_studio_bu_1","currency_id","amount_untaxed"],
            "limit":2000,
            "context":{"allowed_company_ids":[1,2,3,4,5,6]}
        }
    }
})
orders = r2.get("result", [])
print(f"Órdenes con campaas_1='Campañas': {len(orders)}")

# Guardar como JSON para procesar
with open("/tmp/campanas_orders.json", "w") as f:
    json.dump(orders, f)

# Traer subtareas de esas órdenes
order_ids = [o["id"] for o in orders]
print(f"Buscando subtareas de esas {len(order_ids)} órdenes...")

r3 = call(f"{ODOO_URL}/web/dataset/call_kw", {"jsonrpc":"2.0","method":"call","id":3,
    "params":{"model":"project.task","method":"search_read",
        "args":[[["parent_id","!=",False],["sale_order_id","in",order_ids[:500]]]],
        "kwargs":{
            "fields":["id","name","company_id","sale_order_id","x_studio_fecha_de_publicacin","x_studio_fecha_limite_ops","x_studio_related_field_52s_1j0ff39s4","x_studio_related_field_6o9_1jhbqk5st"],
            "limit":5000,
            "context":{"allowed_company_ids":[1,2,3,4,5,6]}
        }
    }
})
tasks = r3.get("result", [])
print(f"Subtareas encontradas: {len(tasks)}")
with open("/tmp/campanas_tasks.json", "w") as f:
    json.dump({"orders": orders, "tasks": tasks}, f)
print("Guardado en /tmp/campanas_tasks.json")
