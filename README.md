# ZAS Talents — Dashboards de Implementaciones

Dashboards HTML generados automáticamente todos los días a las **9:30am Argentina** desde datos de Odoo.

## Links por país

| País | URL |
|------|-----|
| 🇦🇷 Argentina | `https://TU_ORG.github.io/zas-dashboards/argentina.html` |
| 🇨🇱 Chile | `https://TU_ORG.github.io/zas-dashboards/chile.html` |
| 🇨🇴 Colombia | `https://TU_ORG.github.io/zas-dashboards/colombia.html` |
| 🇺🇸 USA | `https://TU_ORG.github.io/zas-dashboards/usa.html` |
| 🇵🇪 Perú | `https://TU_ORG.github.io/zas-dashboards/peru.html` |
| 🌐 Índice | `https://TU_ORG.github.io/zas-dashboards/` |

> Reemplazá `TU_ORG` por tu usuario u organización de GitHub.

---

## Setup inicial (hacer una sola vez)

### 1. Crear el repositorio en GitHub
```
gh repo create zas-dashboards --public
cd zas-dashboards
git init && git add . && git commit -m "Initial commit"
git push -u origin main
```

### 2. Configurar GitHub Secrets

Ir a: **Repo → Settings → Secrets and variables → Actions → New repository secret**

| Secret | Valor |
|--------|-------|
| `ODOO_URL` | `https://zas-talent.odoo.com` |
| `ODOO_DB` | `zas-talent` |
| `ODOO_USER` | `martu@zastalents.com` |
| `ODOO_PASSWORD` | tu contraseña de Odoo |

### 3. Activar GitHub Pages

Ir a: **Repo → Settings → Pages**
- Source: **Deploy from a branch**
- Branch: `main` / folder: `/docs`
- Guardar

Los links estarán disponibles en ~2 minutos.

### 4. Agregar el archivo de presupuesto

Copiar `Base_Regional__2_.xlsx` a la raíz del repo y hacer commit.

---

## Ejecución manual

Desde GitHub: **Actions → Generar Dashboards ZAS → Run workflow**

Localmente (para testing):
```bash
# Crear .env local (no se commitea)
echo "ODOO_PASSWORD=tu_password" > .env
source .env  # o en Windows: set ODOO_PASSWORD=tu_password

pip install -r requirements.txt
python generar_dashboards.py
```

---

## Estructura del proyecto

```
zas-dashboards/
├── generar_dashboards.py       # Script principal
├── Base_Regional__2_.xlsx      # Presupuesto (actualizar cuando cambie)
├── requirements.txt
├── .github/
│   └── workflows/
│       └── daily.yml           # Automatización diaria 9:30am ARG
└── docs/                       # ← GitHub Pages sirve desde acá
    ├── index.html              # Landing con links a países
    ├── argentina.html
    ├── chile.html
    ├── colombia.html
    ├── usa.html
    └── peru.html
```

---

## Actualizar el presupuesto

1. Reemplazar `Base_Regional__2_.xlsx` con el archivo nuevo
2. Commit y push
3. El workflow del día siguiente usará el nuevo archivo automáticamente

---

*Generado por ZAS Talents Ops · Cualquier duda: martu@zastalents.com*
