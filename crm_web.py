import os
import uuid
from datetime import datetime, timedelta, date

from flask import Flask, request, redirect, Response, render_template_string
from supabase import create_client

# =========================
# CONFIG
# =========================
CRM_TABLE = os.environ.get("CRM_TABLE", "crm_records")
RETOQUE_DAYS = int(os.environ.get("RETOQUE_DAYS", "30"))

SERVICES = [
    "CEJAS",
    "DELINEADO DE OJOS",
    "LABIOS",
    "ACIDO HIALURONICO",
    "BIOESTIMULADORES",
    "RETOQUE",
]

# =========================
# SUPABASE (SIEMPRE DEFINIDO)
# =========================
SUPABASE_URL = os.environ.get("SUPABASE_URL")
SUPABASE_KEY = os.environ.get("SUPABASE_ANON_KEY")

if not SUPABASE_URL or not SUPABASE_KEY:
    raise Exception("Faltan variables SUPABASE_URL o SUPABASE_ANON_KEY en Render")

SUPABASE = create_client(SUPABASE_URL, SUPABASE_KEY)

APP = Flask(_name_)

# =========================
# HELPERS
# =========================
def parse_date_any(s: str) -> date | None:
    """Acepta 'YYYY-MM-DD' o 'DD/MM/YYYY' """
    if not s:
        return None
    s = s.strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None

def iso_date(d: date | None) -> str | None:
    return d.isoformat() if d else None

def calc_fecha_retoque(servicio: str, fecha_base: date | None, fecha_retoque_input: date | None) -> date | None:
    """
    Reglas:
    - Si viene fecha_retoque manual -> se respeta
    - Si servicio == 'RETOQUE' y no viene fecha_retoque -> usa la misma fecha
    - Si NO es 'RETOQUE' -> fecha + RETOQUE_DAYS
    """
    if fecha_retoque_input:
        return fecha_retoque_input

    if not fecha_base:
        return None

    if (servicio or "").upper() == "RETOQUE":
        return fecha_base

    return fecha_base + timedelta(days=RETOQUE_DAYS)

def safe_str(x) -> str:
    return (x or "").strip()

def fetch_rows():
    # Orden por created_at si existe; si no, por fecha desc
    # (si tu tabla no tiene created_at, igual funciona: Supabase ignora order si columna no existe? a veces falla)
    # Mejor: intentamos created_at y si falla, intentamos fecha.
    try:
        res = SUPABASE.table(CRM_TABLE).select("*").order("created_at", desc=True).execute()
        return res.data or []
    except Exception:
        res = SUPABASE.table(CRM_TABLE).select("*").order("fecha", desc=True).execute()
        return res.data or []

# =========================
# UI
# =========================
HTML = """
<!doctype html>
<html>
<head>
  <meta charset="utf-8" />
  <title>CRM Salón</title>
  <style>
    body{font-family:Arial, sans-serif; margin:24px;}
    .row{display:flex; gap:12px; flex-wrap:wrap;}
    input, select, textarea{padding:8px; font-size:14px; width:260px;}
    textarea{width:540px; height:70px;}
    button{padding:10px 14px; font-weight:700; cursor:pointer;}
    table{border-collapse:collapse; width:100%; margin-top:16px;}
    th, td{border:1px solid #ddd; padding:8px; font-size:13px;}
    th{background:#f4f4f4; position:sticky; top:0;}
    .msg{margin:10px 0; color:#b00020; font-weight:700;}
    .ok{color:#0a7b2e;}
    .small{font-size:12px; color:#666;}
  </style>
</head>
<body>
  <h2>CRM Salón</h2>

  {% if msg %}
    <div class="msg">{{ msg }}</div>
  {% endif %}

  <form method="post" action="/save">
    <input type="hidden" name="id" value="{{ editing.get('id','') }}">
    <div class="row">
      <div>
        <div class="small">Nombre</div>
        <input name="nombre" value="{{ editing.get('nombre','') }}" placeholder="Nombre" required>
      </div>
      <div>
        <div class="small">Teléfono (opcional)</div>
        <input name="telefono" value="{{ editing.get('telefono','') }}" placeholder="9xxxxxxxx">
      </div>
      <div>
        <div class="small">Fecha (YYYY-MM-DD o DD/MM/YYYY)</div>
        <input name="fecha" value="{{ editing.get('fecha','') }}" placeholder="2026-02-12" required>
      </div>
      <div>
        <div class="small">Servicio</div>
        <select name="servicio" required>
          <option value="">-- Selecciona --</option>
          {% for s in services %}
            <option value="{{s}}" {% if editing.get('servicio','') == s %}selected{% endif %}>{{s}}</option>
          {% endfor %}
        </select>
      </div>
      <div>
        <div class="small">Fecha retoque (opcional)</div>
        <input name="fecha_retoque" value="{{ editing.get('fecha_retoque','') }}" placeholder="(auto si vacío)">
        <div class="small">Si lo dejas vacío: auto = fecha + {{ retoque_days }} días (o igual si es RETOQUE)</div>
      </div>
    </div>

    <div style="margin-top:10px;">
      <div class="small">Comentario</div>
      <textarea name="comentario" placeholder="Notas...">{{ editing.get('comentario','') }}</textarea>
    </div>

    <div style="margin-top:12px;" class="row">
      <button type="submit">Guardar</button>
      <a href="/" style="padding:10px 14px; text-decoration:none; border:1px solid #ddd;">Limpiar</a>
      <a href="/export.csv" style="padding:10px 14px; text-decoration:none; border:1px solid #ddd;">Exportar CSV</a>
    </div>
  </form>

  <h3 style="margin-top:22px;">Registros ({{ rows|length }})</h3>

  <table>
    <thead>
      <tr>
        <th>Nombre</th>
        <th>Teléfono</th>
        <th>Fecha</th>
        <th>Fecha retoque</th>
        <th>Servicio</th>
        <th>Comentario</th>
        <th>Acciones</th>
      </tr>
    </thead>
    <tbody>
      {% for r in rows %}
        <tr>
          <td>{{ r.get('nombre','') }}</td>
          <td>{{ r.get('telefono','') }}</td>
          <td>{{ r.get('fecha','') }}</td>
          <td>{{ r.get('fecha_retoque','') }}</td>
          <td>{{ r.get('servicio','') }}</td>
          <td style="max-width:380px; white-space:pre-wrap;">{{ r.get('comentario','') }}</td>
          <td>
            <a href="/?edit={{ r.get('id','') }}">Editar</a>
            |
            <a href="/delete?id={{ r.get('id','') }}" onclick="return confirm('¿Eliminar?')">Eliminar</a>
          </td>
        </tr>
      {% endfor %}
    </tbody>
  </table>

</body>
</html>
"""

# =========================
# ROUTES
# =========================
@APP.get("/")
def home():
    msg = request.args.get("msg", "")
    edit_id = request.args.get("edit", "")
    rows = fetch_rows()
    editing = {}
    if edit_id:
        for r in rows:
            if r.get("id") == edit_id:
                editing = dict(r)
                break
    return render_template_string(
        HTML,
        rows=rows,
        editing=editing,
        services=SERVICES,
        retoque_days=RETOQUE_DAYS,
        msg=msg,
    )

@APP.post("/save")
def save():
    try:
        rid = safe_str(request.form.get("id"))
        nombre = safe_str(request.form.get("nombre"))
        telefono = safe_str(request.form.get("telefono"))
        servicio = safe_str(request.form.get("servicio")).upper()
        comentario = safe_str(request.form.get("comentario"))

        fecha = parse_date_any(request.form.get("fecha"))
        fecha_retoque_in = parse_date_any(request.form.get("fecha_retoque"))

        if not nombre:
            return redirect("/?msg=Falta+nombre")
        if not servicio:
            return redirect("/?msg=Falta+servicio")
        if servicio not in SERVICES:
            return redirect("/?msg=Servicio+inv%C3%A1lido")
        if not fecha:
            return redirect("/?msg=Fecha+inv%C3%A1lida+(usa+YYYY-MM-DD+o+DD/MM/YYYY)")

        fecha_retoque = calc_fecha_retoque(servicio, fecha, fecha_retoque_in)

        if not rid:
            rid = str(uuid.uuid4())

        payload = {
            "id": rid,
            "nombre": nombre,
            "telefono": telefono if telefono else None,
            "fecha": iso_date(fecha),
            "servicio": servicio,
            "comentario": comentario if comentario else None,
            "fecha_retoque": iso_date(fecha_retoque),
        }

        # Upsert (crea o actualiza)
        SUPABASE.table(CRM_TABLE).upsert(payload).execute()

        return redirect("/?msg=Guardado+OK")
    except Exception as e:
        # Esto te muestra el error real en pantalla (para no quedarnos a ciegas)
        return redirect("/?msg=ERROR:+{}".format(str(e).replace(" ", "+")))

@APP.get("/delete")
def delete():
    rid = request.args.get("id", "")
    if rid:
        SUPABASE.table(CRM_TABLE).delete().eq("id", rid).execute()
    return redirect("/?msg=Eliminado")

@APP.get("/export.csv")
def export_csv():
    rows = fetch_rows()
    def esc(s):
        s = "" if s is None else str(s)
        s = s.replace('"', '""')
        return f'"{s}"'
    headers = ["nombre", "telefono", "fecha", "fecha_retoque", "servicio", "comentario"]
    out = [",".join(headers)]
    for r in rows:
        out.append(",".join(esc(r.get(h, "")) for h in headers))
    csv_text = "\n".join(out)
    return Response(
        csv_text,
        mimetype="text/csv",
        headers={"Content-Disposition": "attachment; filename=crm_export.csv"},
    )

# Render/Gunicorn usa APP