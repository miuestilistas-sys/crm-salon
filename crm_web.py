import json
import os
import uuid
from copy import deepcopy
from datetime import datetime, timedelta

from flask import Flask, request, redirect, send_file, render_template_string, Response

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

APP = Flask(__name__)
DATA_FILE = "crm_data.json"
UNDO_FILE = "crm_undo.json"
EXPORT_FILE = "crm_export.xlsx"
UNDO_MAX = 30  # historial tipo Ctrl+Z

SERVICES = [
    "CEJAS",
    "DELINEADO DE OJOS",
    "LABIOS",
    "ACIDO HIALURONICO",
    "BIOESTIMULADORES",
    "RETOQUE",
]

COLUMNS_XLSX = ["NOMBRE", "TELEFONO", "FECHA", "FECHA RETOQUE", "SERVICIO", "COMENTARIO"]


# ---------- fechas ----------
def parse_ddmmyyyy(s: str):
    return datetime.strptime(s.strip(), "%d/%m/%Y")


def fmt_ddmmyyyy(dt: datetime):
    return dt.strftime("%d/%m/%Y")


def is_retouch_service(service: str) -> bool:
    return service.strip().upper() == "RETOQUE"


def compute_retouch_date(fecha_str: str, service: str) -> str:
    dt = parse_ddmmyyyy(fecha_str)
    days = 365 if is_retouch_service(service) else 20
    return fmt_ddmmyyyy(dt + timedelta(days=days))


def is_due(retouch_str: str) -> bool:
    try:
        r = parse_ddmmyyyy(retouch_str).date()
        return r <= datetime.now().date()
    except Exception:
        return False


# ---------- persistencia ----------
def load_data():
    if not os.path.exists(DATA_FILE):
        return []
    try:
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        out = []
        for r in data:
            out.append(
                {
                    "id": r.get("id") or str(uuid.uuid4()),
                    "nombre": (r.get("nombre") or "").strip(),
                    "telefono": (r.get("telefono") or "").strip(),
                    "fecha": (r.get("fecha") or "").strip(),  # dd/mm/yyyy
                    "servicio": (r.get("servicio") or "").strip(),
                    "comentario": (r.get("comentario") or "").strip(),
                    "recordatorio": bool(r.get("recordatorio", False)),
                }
            )
        return out
    except Exception:
        return []


def save_data(rows):
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(rows, f, ensure_ascii=False, indent=2)


# ---------- Undo tipo Ctrl+Z (historial) ----------
def load_undo_stack():
    if not os.path.exists(UNDO_FILE):
        return []
    try:
        with open(UNDO_FILE, "r", encoding="utf-8") as f:
            obj = json.load(f)
        stack = obj.get("stack", [])
        return stack if isinstance(stack, list) else []
    except Exception:
        return []


def save_undo_stack(stack):
    stack = stack[-UNDO_MAX:]
    with open(UNDO_FILE, "w", encoding="utf-8") as f:
        json.dump({"stack": stack}, f, ensure_ascii=False, indent=2)


def push_undo_snapshot(before_rows):
    stack = load_undo_stack()
    stack.append(before_rows)
    save_undo_stack(stack)


def pop_undo_snapshot():
    stack = load_undo_stack()
    if not stack:
        return None
    snap = stack.pop()
    save_undo_stack(stack)
    return snap


# ---------- validación ----------
def validate_row(nombre, fecha, servicio):
    if not nombre:
        return "El nombre es obligatorio."
    try:
        parse_ddmmyyyy(fecha)
    except Exception:
        return "Fecha inválida (selecciona una fecha del calendario)."
    if not servicio:
        return "El servicio es obligatorio."
    return None


HTML = r"""
<!doctype html>
<html lang="es">
<head>
  <meta charset="utf-8"/>
  <meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover"/>
  <meta name="theme-color" content="#2563eb"/>
  <link rel="manifest" href="/manifest.webmanifest">
  <title>CRM Salón</title>

  <style>
    body { font-family: Arial, sans-serif; margin: 14px; background:#f6f7fb; }
    .wrap { max-width: 1100px; margin: 0 auto; }
    .card { background: white; border-radius: 14px; padding: 14px; box-shadow: 0 6px 18px rgba(0,0,0,.06); margin-bottom: 12px; }
    h1 { margin: 0 0 10px 0; font-size: 20px; }
    .grid { display:grid; grid-template-columns: 1fr 1fr; gap: 10px; }
    .grid3 { display:grid; grid-template-columns: 1fr 1fr 1fr; gap: 10px; }
    label { font-size: 12px; color: #333; }
    input, select, textarea { width: 100%; padding: 10px; border-radius: 10px; border: 1px solid #d9dbe7; font-size: 14px; background:#fff; }
    textarea { min-height: 70px; resize: vertical; }
    .row { display:flex; gap:10px; flex-wrap: wrap; align-items: center; }
    .btn { padding: 10px 12px; border-radius: 10px; border: 0; cursor:pointer; font-weight:700; text-decoration:none; display:inline-block; }
    .btn-primary { background:#2563eb; color:#fff; }
    .btn-danger { background:#ef4444; color:#fff; }
    .btn-ghost { background:#eef2ff; }
    .btn-ok { background:#16a34a; color:#fff; }
    .btn-warn { background:#f59e0b; color:#111827; }
    .btn-dark { background:#111827; color:#fff; }
    .muted { color:#555; font-size: 13px; }
    .banner { font-weight: 800; font-size: 14px; }
    .error { color:#b91c1c; font-weight:700; }

    table { width:100%; border-collapse: collapse; overflow:hidden; border-radius: 12px; }
    th, td { border: 1px solid #e6e8f2; padding: 10px; text-align: center; vertical-align: middle; font-size: 14px; }
    th { background:#d9ead3; font-size: 13px; position: sticky; top: 0; z-index: 2; }
    td.comment { text-align:left; white-space: pre-wrap; }
    tr.due { background: #fff2cc; }
    tr:hover { outline: 2px solid rgba(37,99,235,.15); }

    /* columnas delgadas */
    th.rem, td.rem { width: 58px; padding-left: 6px; padding-right: 6px; }
    th.act, td.act { width: 58px; padding-left: 6px; padding-right: 6px; }

    /* checkbox verde */
    .chk { width: 22px; height: 22px; accent-color: #16a34a; cursor: pointer; }
    .chkWrap { display:flex; justify-content:center; align-items:center; }

    /* contenedor con scroll para ver TODOS los registros */
    .tableScroll { max-height: 65vh; overflow: auto; border-radius: 12px; }

    @media (max-width: 820px){
      .grid, .grid3 { grid-template-columns: 1fr; }
      th, td { font-size: 13px; padding: 8px; }
      th.rem, td.rem, th.act, td.act { width: 54px; }
      .tableScroll { max-height: 60vh; }
    }
  </style>
</head>

<body>
<div class="wrap">

  <div class="card">
    <h1>CRM Salón</h1>

    <div class="row" style="justify-content:space-between;">
      <div class="banner">
        {{ banner }}
        <br><span class="muted">Tip: toca una fila para cargarla arriba y editarla.</span>
      </div>
      <button class="btn btn-dark" type="button" onclick="goFullscreen()">⛶ Pantalla completa</button>
    </div>

    {% if error %}
      <p class="error">⚠️ {{ error }}</p>
    {% endif %}
  </div>

  <div class="card">
    <form method="post" action="/save" id="mainForm" autocomplete="off">
      <input type="hidden" name="id" id="rid" value="">

      <div class="grid">
        <div>
          <label>NOMBRE</label>
          <input name="nombre" id="nombre" placeholder="Ej: Romina" required autocomplete="off">
        </div>
        <div>
          <label>TELEFONO (opcional)</label>
          <input name="telefono" id="telefono" placeholder="Ej: 999999999 (opcional)" inputmode="tel" autocomplete="off">
        </div>
      </div>

      <div class="grid3" style="margin-top:10px;">
        <div>
          <label>FECHA (calendario)</label>
          <input type="hidden" name="fecha" id="fecha_hidden" value="{{ today_ddmmyyyy }}">
          <input id="fecha_picker" type="date" value="{{ today_iso }}" autocomplete="off">
        </div>

        <div>
          <label>SERVICIO</label>
          <select name="servicio" id="servicio" required autocomplete="off">
            <option value="">-- Selecciona --</option>
            {% for s in services %}
              <option value="{{ s }}">{{ s }}</option>
            {% endfor %}
          </select>
        </div>

        <div>
          <label>FECHA RETOQUE (auto)</label>
          <input id="retoque" readonly>
        </div>
      </div>

      <div style="margin-top:10px;">
        <label>COMENTARIO</label>
        <textarea name="comentario" id="comentario" placeholder="Ej: Se usó black y brown..." autocomplete="off"></textarea>
      </div>

      <div class="row" style="margin-top:12px;">
        <button class="btn btn-primary" type="submit">💾 Guardar (Agregar/Actualizar)</button>
        <button class="btn btn-ghost" type="button" onclick="clearForm()">🧹 Limpiar</button>

        <button class="btn btn-warn" type="submit"
                formaction="/undo" formmethod="post" formnovalidate
                {% if not can_undo %}disabled{% endif %}>
          ↩️ Deshacer
        </button>

        <a class="btn btn-ok" href="/export">📤 Exportar</a>
      </div>
    </form>
  </div>

  <div class="card">
    <div class="row" style="justify-content:space-between;">
      <div style="flex:1; min-width:240px;">
        <label>Filtrar por nombre</label>
        <input id="q" placeholder="Escribe y luego presiona Buscar..." value="{{ q }}" autocomplete="off">
      </div>

      <button class="btn btn-primary" type="button" onclick="applyFilter()">🔎 Buscar</button>
      <button class="btn btn-ghost" type="button" onclick="clearFilter()">✖ Limpiar filtro</button>
    </div>

    <div class="tableScroll" style="margin-top:10px;">
      <table>
        <thead>
          <tr>
            <th>NOMBRE</th>
            <th>TELEFONO</th>
            <th>FECHA</th>
            <th>RETOQUE</th>
            <th>SERVICIO</th>
            <th>COMENTARIO</th>
            <th class="rem">REC.</th>
            <th class="act">ACC.</th>
          </tr>
        </thead>

        <tbody>
          {% for r in rows %}
            <tr class="{{ r.row_class }}"
                onclick='loadRow({{ r|tojson }})'
                style="cursor:pointer;">
              <td><b>{{ r.nombre }}</b></td>
              <td>{{ r.telefono }}</td>
              <td>{{ r.fecha }}</td>
              <td>{{ r.retoque }}</td>
              <td>{{ r.servicio }}</td>
              <td class="comment">{{ r.comentario }}</td>

              <!-- RECORDATORIO con confirm (FIX: guarda bien el check) -->
              <td class="rem" onclick="event.stopPropagation();">
                <form method="post" action="/toggle_reminder" style="margin:0;" onclick="event.stopPropagation();">
                  <input type="hidden" name="id" value="{{ r.id }}">
                  <input type="hidden" name="target" value="">
                  <div class="chkWrap">
                    <input class="chk" type="checkbox"
                           {% if r.recordatorio %}checked{% endif %}
                           onclick="return onReminderClick(event, this);">
                  </div>
                </form>
              </td>

              <td class="act" onclick="event.stopPropagation();">
                <form method="post" action="/delete"
                      onsubmit="return confirm('¿Eliminar este registro?');"
                      style="margin:0;" onclick="event.stopPropagation();">
                  <input type="hidden" name="id" value="{{ r.id }}">
                  <button class="btn btn-danger" type="submit" onclick="event.stopPropagation();">🗑️</button>
                </form>
              </td>
            </tr>
          {% endfor %}

          {% if rows|length == 0 %}
            <tr><td colspan="8" class="muted">No hay resultados.</td></tr>
          {% endif %}
        </tbody>
      </table>
    </div>
  </div>

</div>

<script>
  function goFullscreen(){
    const el = document.documentElement;
    if (el.requestFullscreen) el.requestFullscreen();
    else if (el.webkitRequestFullscreen) el.webkitRequestFullscreen();
  }

  function dateValueToDDMMYYYY(val){
    const parts = (val || "").split("-");
    if(parts.length !== 3) return "";
    const yy = parts[0], mm = parts[1], dd = parts[2];
    if(!yy || !mm || !dd) return "";
    return `${dd}/${mm}/${yy}`;
  }

  function ddmmyyyyToDateValue(ddmmyyyy){
    const parts = (ddmmyyyy || "").split("/");
    if(parts.length !== 3) return "";
    const dd = parts[0], mm = parts[1], yy = parts[2];
    if(dd.length!==2 || mm.length!==2 || yy.length!==4) return "";
    return `${yy}-${mm}-${dd}`;
  }

  function computeRetouchFromHidden(servicio){
    const ddmmyyyy = document.getElementById("fecha_hidden").value.trim();
    try{
      const parts = ddmmyyyy.split("/");
      if(parts.length !== 3) return "";
      const d = parseInt(parts[0],10), m = parseInt(parts[1],10)-1, y = parseInt(parts[2],10);
      const dt = new Date(y, m, d);
      if(isNaN(dt.getTime())) return "";
      const isRet = (servicio || "").trim().toUpperCase() === "RETOQUE";
      const days = isRet ? 365 : 20;
      dt.setDate(dt.getDate() + days);
      const dd = String(dt.getDate()).padStart(2,"0");
      const mm = String(dt.getMonth()+1).padStart(2,"0");
      const yy = dt.getFullYear();
      return `${dd}/${mm}/${yy}`;
    }catch(e){ return ""; }
  }

  function updateRetouch(){
    const s = document.getElementById("servicio").value.trim();
    document.getElementById("retoque").value = computeRetouchFromHidden(s);
  }

  function syncHiddenFromPicker(){
    const picker = document.getElementById("fecha_picker");
    const hidden = document.getElementById("fecha_hidden");
    const ddmmyyyy = dateValueToDDMMYYYY(picker.value);
    if(ddmmyyyy) hidden.value = ddmmyyyy;
    updateRetouch();
  }

  function loadRow(r){
    document.getElementById("rid").value = r.id || "";
    document.getElementById("nombre").value = r.nombre || "";
    document.getElementById("telefono").value = r.telefono || "";
    document.getElementById("servicio").value = r.servicio || "";
    document.getElementById("comentario").value = r.comentario || "";

    const picker = document.getElementById("fecha_picker");
    const hidden = document.getElementById("fecha_hidden");
    hidden.value = r.fecha || hidden.value;
    const iso = ddmmyyyyToDateValue(hidden.value);
    if(iso) picker.value = iso;

    updateRetouch();
    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  function clearForm(){
    document.getElementById("rid").value = "";
    document.getElementById("nombre").value = "";
    document.getElementById("telefono").value = "";
    document.getElementById("servicio").value = "";
    document.getElementById("comentario").value = "";

    document.getElementById("fecha_picker").value = "{{ today_iso }}";
    document.getElementById("fecha_hidden").value = "{{ today_ddmmyyyy }}";

    updateRetouch();
  }

  function applyFilter(){
    const q = document.getElementById("q").value;
    const url = new URL(window.location.href);
    if(q.trim()){
      url.searchParams.set("q", q.trim());
    } else {
      url.searchParams.delete("q");
    }
    window.location.href = url.toString();
  }

  function clearFilter(){
    const url = new URL(window.location.href);
    url.searchParams.delete("q");
    window.location.href = url.toString();
  }

  // Enter en filtro = buscar
  document.addEventListener("keydown", function(e){
    const q = document.getElementById("q");
    if(e.key === "Enter" && document.activeElement === q){
      e.preventDefault();
      applyFilter();
    }
  });

  // FIX checkbox: no deja que cambie solo, pregunta y recién guarda
  function onReminderClick(ev, chk){
    ev.stopPropagation();
    ev.preventDefault();

    const form = chk.closest("form");
    const targetInput = form.querySelector('input[name="target"]');

    const current = chk.checked;
    const desired = !current;

    const msg = desired
      ? "¿Se le envió recordatorio?\n\nOK = Sí / Cancelar = No"
      : "¿Quitar la marca de recordatorio?\n\nOK = Sí / Cancelar = No";

    const ok = confirm(msg);
    if(ok){
      chk.checked = desired;
      targetInput.value = desired ? "1" : "0";
      form.submit();
    }
    return false;
  }

  document.getElementById("fecha_picker").addEventListener("change", syncHiddenFromPicker);
  document.getElementById("servicio").addEventListener("change", updateRetouch);

  syncHiddenFromPicker();

  if ("serviceWorker" in navigator) {
    navigator.serviceWorker.register("/sw.js").catch(()=>{});
  }
</script>
</body>
</html>
"""


@APP.get("/manifest.webmanifest")
def manifest():
    m = {
        "name": "CRM Salon",
        "short_name": "CRM",
        "start_url": "/",
        "display": "fullscreen",
        "background_color": "#f6f7fb",
        "theme_color": "#2563eb",
        "icons": [],
    }
    return Response(json.dumps(m), mimetype="application/manifest+json")


@APP.get("/sw.js")
def sw():
    js = r"""
self.addEventListener("install", (e) => {
  self.skipWaiting();
  e.waitUntil(caches.open("crm-cache-v7").then((cache) => cache.addAll(["/"])));
});
self.addEventListener("activate", (e) => {
  e.waitUntil(self.clients.claim());
});
self.addEventListener("fetch", (e) => {
  e.respondWith(caches.match(e.request).then((r) => r || fetch(e.request)));
});
"""
    return Response(js, mimetype="application/javascript")


@APP.get("/")
def index():
    data = load_data()
    q_raw = request.args.get("q", "")
    q = (q_raw or "").strip().lower()

    rows = []
    nearest = None

    for r in data:
        if q and q not in r["nombre"].lower():
            continue

        retoque = ""
        due = False
        try:
            retoque = compute_retouch_date(r["fecha"], r["servicio"])
            due = is_due(retoque)
        except Exception:
            pass

        if retoque:
            try:
                d = parse_ddmmyyyy(retoque).date()
                if nearest is None or d < nearest[0]:
                    nearest = (d, retoque)
            except Exception:
                pass

        rows.append({**r, "retoque": retoque, "row_class": ("due" if due else "")})

    today_iso = datetime.now().strftime("%Y-%m-%d")
    today_ddmmyyyy = datetime.now().strftime("%d/%m/%Y")

    banner = "Retoque: —"
    if nearest:
        d, txt = nearest
        diff = (d - datetime.now().date()).days
        if diff > 0:
            banner = f"Próximo retoque más cercano: {txt} (faltan {diff} días)"
        else:
            banner = f"⚠️ Hay retoques que ya tocan (ej: {txt})"

    error = request.args.get("error") or ""
    can_undo = len(load_undo_stack()) > 0

    return render_template_string(
        HTML,
        rows=rows,
        services=SERVICES,
        today_iso=today_iso,
        today_ddmmyyyy=today_ddmmyyyy,
        q=q_raw,
        banner=banner,
        error=error,
        can_undo=can_undo,
    )


@APP.post("/save")
def save():
    data = load_data()
    before = deepcopy(data)

    rid = (request.form.get("id") or "").strip()
    nombre = (request.form.get("nombre") or "").strip()
    telefono = (request.form.get("telefono") or "").strip()
    fecha = (request.form.get("fecha") or "").strip()
    servicio = (request.form.get("servicio") or "").strip()
    comentario = (request.form.get("comentario") or "").strip()

    err = validate_row(nombre, fecha, servicio)
    if err:
        return redirect(f"/?error={err}")

    if rid:
        found = False
        for r in data:
            if str(r.get("id")) == str(rid):  # FIX: comparación segura
                rec = bool(r.get("recordatorio", False))  # mantener estado del check
                r.update(
                    {
                        "nombre": nombre,
                        "telefono": telefono,
                        "fecha": fecha,
                        "servicio": servicio,
                        "comentario": comentario,
                        "recordatorio": rec,
                    }
                )
                found = True
                break
        if not found:
            data.append(
                {
                    "id": rid,
                    "nombre": nombre,
                    "telefono": telefono,
                    "fecha": fecha,
                    "servicio": servicio,
                    "comentario": comentario,
                    "recordatorio": False,
                }
            )
    else:
        data.append(
            {
                "id": str(uuid.uuid4()),
                "nombre": nombre,
                "telefono": telefono,
                "fecha": fecha,
                "servicio": servicio,
                "comentario": comentario,
                "recordatorio": False,
            }
        )

    push_undo_snapshot(before)
    save_data(data)
    return redirect("/")


@APP.post("/delete")
def delete():
    data = load_data()
    before = deepcopy(data)

    rid = (request.form.get("id") or "").strip()
    data = [r for r in data if str(r.get("id")) != str(rid)]

    push_undo_snapshot(before)
    save_data(data)
    return redirect("/")


@APP.post("/toggle_reminder")
def toggle_reminder():
    data = load_data()
    before = deepcopy(data)

    rid = (request.form.get("id") or "").strip()
    target = (request.form.get("target") or "").strip()  # "1" marcar, "0" desmarcar
    want = True if target == "1" else False

    for r in data:
        if str(r.get("id")) == str(rid):
            r["recordatorio"] = want
            break

    push_undo_snapshot(before)
    save_data(data)
    return redirect("/")


@APP.post("/undo")
def undo():
    snap = pop_undo_snapshot()
    if snap is None:
        return redirect("/")
    save_data(snap)
    return redirect("/")


# ---------- Export Excel (mismo archivo) ----------
def build_export_excel(path: str, data: list):
    if os.path.exists(path):
        try:
            wb = load_workbook(path)
            if "CRM" in wb.sheetnames:
                ws = wb["CRM"]
                wb.remove(ws)
            ws = wb.create_sheet("CRM", 0)
        except Exception:
            wb = Workbook()
            ws = wb.active
            ws.title = "CRM"
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "CRM"

    ws.append(COLUMNS_XLSX)

    for r in data:
        retoque = ""
        try:
            retoque = compute_retouch_date(r["fecha"], r["servicio"])
        except Exception:
            pass
        ws.append([r["nombre"], r["telefono"], r["fecha"], retoque, r["servicio"], r["comentario"]])

    header_fill = PatternFill("solid", fgColor="D9EAD3")
    header_font = Font(bold=True, size=12)
    cell_font = Font(size=11)

    thin = Side(style="thin", color="999999")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    due_fill = PatternFill("solid", fgColor="FFF2CC")

    col_widths = {"A": 18, "B": 16, "C": 14, "D": 16, "E": 22, "F": 48}
    for col, w in col_widths.items():
        ws.column_dimensions[col].width = w

    ws.row_dimensions[1].height = 26
    for i in range(2, ws.max_row + 1):
        ws.row_dimensions[i].height = 42

    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=len(COLUMNS_XLSX)):
        row_idx = row[0].row
        retoque_cell = ws.cell(row=row_idx, column=4)
        due = (row_idx != 1) and is_due(str(retoque_cell.value or ""))

        for cell in row:
            cell.border = border
            cell.alignment = align
            if row_idx == 1:
                cell.fill = header_fill
                cell.font = header_font
            else:
                cell.font = cell_font
                if due:
                    cell.fill = due_fill

    if "Sheet" in wb.sheetnames and len(wb.sheetnames) > 1:
        sh = wb["Sheet"]
        if sh.max_row == 1 and sh.max_column == 1 and (sh["A1"].value is None):
            wb.remove(sh)

    wb.save(path)


@APP.get("/export")
def export():
    data = load_data()
    build_export_excel(EXPORT_FILE, data)
    return send_file(EXPORT_FILE, as_attachment=True, download_name=EXPORT_FILE)


if __name__ == "__main__":
    APP.run(host="0.0.0.0", port=5000, debug=True)
