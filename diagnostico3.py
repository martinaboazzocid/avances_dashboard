import json, urllib.request, http.cookiejar, os
from collections import Counter

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
print(f"UID={uid}\n")

# Traer TODOS los valores posibles de los campos de selección en sale.order
# Primero: obtener las definiciones del campo (los selection values)
r2 = call(f"{ODOO_URL}/web/dataset/call_kw", {"jsonrpc":"2.0","method":"call","id":2,
    "params":{"model":"sale.order","method":"fields_get",
        "args":[["x_studio_campaas","x_studio_campaas_1","x_studio_bu","x_studio_bu_1"]],
        "kwargs":{"attributes":["string","type","selection"],
                  "context":{"allowed_company_ids":[1,2,3,4,5,6]}}
    }
})
fields_def = r2.get("result", {})
print("=== DEFINICIÓN DE CAMPOS DE SELECCIÓN EN SALE.ORDER ===")
for fname, finfo in fields_def.items():
    print(f"\n{fname} | {finfo.get('string')} | {finfo.get('type')}")
    sel = finfo.get("selection", [])
    for val, label in sel:
        print(f"  valor_interno='{val}' → label='{label}'")

# Ahora ver los valores reales que tienen las órdenes activas
r3 = call(f"{ODOO_URL}/web/dataset/call_kw", {"jsonrpc":"2.0","method":"call","id":3,
    "params":{"model":"sale.order","method":"search_read",
        "args":[[["state","in",["sale","done"]]]],
        "kwargs":{
            "fields":["id","x_studio_campaas","x_studio_campaas_1","x_studio_bu","x_studio_bu_1","currency_id"],
            "limit":5000,
            "context":{"allowed_company_ids":[1,2,3,4,5,6]}
        }
    }
})
orders = r3.get("result", [])
print(f"\n=== VALORES REALES EN {len(orders)} ÓRDENES ===")

camp_counter = Counter(str(o.get("x_studio_campaas")) for o in orders)
camp1_counter = Counter(str(o.get("x_studio_campaas_1")) for o in orders)
bu_counter = Counter(str(o.get("x_studio_bu")) for o in orders)
bu1_counter = Counter(str(o.get("x_studio_bu_1")) for o in orders)

print("\nx_studio_campaas (valores únicos):")
for v, n in camp_counter.most_common():
    print(f"  {n:5d}x  {repr(v)}")

print("\nx_studio_campaas_1 (valores únicos):")
for v, n in camp1_counter.most_common():
    print(f"  {n:5d}x  {repr(v)}")

print("\nx_studio_bu (valores únicos):")
for v, n in bu_counter.most_common():
    print(f"  {n:5d}x  {repr(v)}")

print("\nx_studio_bu_1 (valores únicos):")
for v, n in bu1_counter.most_common():
    print(f"  {n:5d}x  {repr(v)}")

# Ver combinaciones campaas + campaas_1
from collections import defaultdict
combos = Counter((str(o.get("x_studio_campaas")), str(o.get("x_studio_campaas_1"))) for o in orders)
print("\nCombinaciones campaas + campaas_1 (top 30):")
for (c1, c2), n in combos.most_common(30):
    print(f"  {n:5d}x  campaas={repr(c1)}  campaas_1={repr(c2)}")
