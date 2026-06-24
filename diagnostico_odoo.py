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

# Auth
r = call(f"{ODOO_URL}/web/session/authenticate", {"jsonrpc":"2.0","method":"call","id":1,
    "params":{"db":ODOO_DB,"login":ODOO_USER,"password":ODOO_PASS}})
uid = r["result"]["uid"]
print(f"UID={uid}")

# Contar todas las tareas sin filtro
r2 = call(f"{ODOO_URL}/web/dataset/call_kw", {"jsonrpc":"2.0","method":"call","id":2,
    "params":{"model":"project.task","method":"search_count","args":[[]],"kwargs":{
        "context":{"allowed_company_ids":[1,2,3,4,5,6]}}}})
print(f"Total tareas (sin filtro): {r2.get('result')}")

# Listar todos los proyectos
r3 = call(f"{ODOO_URL}/web/dataset/call_kw", {"jsonrpc":"2.0","method":"call","id":3,
    "params":{"model":"project.project","method":"search_read","args":[[]],"kwargs":{
        "fields":["id","name"],"limit":100,
        "context":{"allowed_company_ids":[1,2,3,4,5,6]}}}})
proyectos = r3.get("result",[])
print(f"\nProyectos ({len(proyectos)}):")
for p in sorted(proyectos, key=lambda x: x["name"]):
    print(f"  id={p['id']:4d} | {p['name']}")

# Primeras 10 tareas
r4 = call(f"{ODOO_URL}/web/dataset/call_kw", {"jsonrpc":"2.0","method":"call","id":4,
    "params":{"model":"project.task","method":"search_read","args":[[]],"kwargs":{
        "fields":["id","project_id","name","company_id"],"limit":10,
        "context":{"allowed_company_ids":[1,2,3,4,5,6]}}}})
tasks = r4.get("result",[])
print(f"\nPrimeras 10 tareas:")
for t in tasks:
    print(f"  id={t['id']} | proj={t.get('project_id')} | company={t.get('company_id')} | {t['name'][:60]}")

# Subtareas (con parent_id)
r5 = call(f"{ODOO_URL}/web/dataset/call_kw", {"jsonrpc":"2.0","method":"call","id":5,
    "params":{"model":"project.task","method":"search_count",
        "args":[[["parent_id","!=",False]]],"kwargs":{
        "context":{"allowed_company_ids":[1,2,3,4,5,6]}}}})
print(f"\nTareas con parent_id (subtareas): {r5.get('result')}")

# Tareas con sale_order_id
r6 = call(f"{ODOO_URL}/web/dataset/call_kw", {"jsonrpc":"2.0","method":"call","id":6,
    "params":{"model":"project.task","method":"search_count",
        "args":[[["sale_order_id","!=",False]]],"kwargs":{
        "context":{"allowed_company_ids":[1,2,3,4,5,6]}}}})
print(f"Tareas con sale_order_id: {r6.get('result')}")
