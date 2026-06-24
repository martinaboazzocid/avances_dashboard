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
    with _opener.open(req, timeout=60) as r:
        return json.loads(r.read())

r = call(f"{ODOO_URL}/web/session/authenticate", {"jsonrpc":"2.0","method":"call","id":1,
    "params":{"db":ODOO_DB,"login":ODOO_USER,"password":ODOO_PASS}})
print(f"UID={r['result']['uid']}")

# Ver exactamente qué tienen las órdenes US00882, US00881, US00234
r2 = call(f"{ODOO_URL}/web/dataset/call_kw", {"jsonrpc":"2.0","method":"call","id":2,
    "params":{"model":"sale.order","method":"search_read",
        "args":[[["name","in",["US00882","US00881","US00234","US00766","US00808"]]]],
        "kwargs":{"fields":["id","name","state","x_studio_campaas","x_studio_campaas_1","x_studio_bu_1"],
                  "limit":20,"context":{"allowed_company_ids":[1,2,3,4,5,6]}}
    }
})
orders = r2.get("result",[])
print(f"\n=== ÓRDENES INTERNACIONALES DE ARGENTINA ===")
for o in orders:
    print(f"  {o['name']}: state={o['state']} | campaas={repr(o.get('x_studio_campaas'))} | campaas_1={repr(o.get('x_studio_campaas_1'))} | bu_1={repr(o.get('x_studio_bu_1'))}")

# Ver también cuántas órdenes tienen campaas='Internacional' o campaas_1='Internacional'
from collections import Counter
r3 = call(f"{ODOO_URL}/web/dataset/call_kw", {"jsonrpc":"2.0","method":"call","id":3,
    "params":{"model":"sale.order","method":"search_read",
        "args":[[["state","in",["sale","done"]]]],
        "kwargs":{"fields":["x_studio_campaas","x_studio_campaas_1","x_studio_bu_1"],
                  "limit":5000,"context":{"allowed_company_ids":[1,2,3,4,5,6]}}
    }
})
orders_all = r3.get("result",[])
c1 = Counter(str(o.get("x_studio_campaas_1")) for o in orders_all)
c2 = Counter(str(o.get("x_studio_campaas")) for o in orders_all)
print(f"\n=== campaas_1 valores únicos ({len(orders_all)} órdenes activas) ===")
for v,n in c1.most_common(15):
    print(f"  {n:5d}x  {repr(v)}")
print(f"\n=== campaas valores únicos ===")
for v,n in c2.most_common(15):
    print(f"  {n:5d}x  {repr(v)}")
