
from flask import Flask, request, redirect, session, render_template_string, send_file
import sqlite3
from datetime import date
from openpyxl import Workbook
from io import BytesIO

app = Flask(__name__)
app.secret_key = "sistema_gastos_2026"

def conectar():
    conexion = sqlite3.connect("gastos.db")
    conexion.row_factory = sqlite3.Row
    return conexion

def crear_tabla():
    conexion = conectar()
    conexion.execute("""
        CREATE TABLE IF NOT EXISTS gastos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fecha TEXT NOT NULL,
            descripcion TEXT NOT NULL,
            monto REAL NOT NULL
        )
    """)
    conexion.commit()
    conexion.close()

crear_tabla()

ESTILO = """
<style>
body {font-family:Arial;background:#eef2f7;margin:0;color:#243247}
header {background:#17365d;color:white;padding:20px;text-align:center}
.contenedor {max-width:950px;margin:30px auto;background:white;
padding:25px;border-radius:12px}
input,button {padding:12px;margin:7px 0;border-radius:6px;
border:1px solid #ccc}
input {width:95%}
button {background:#1769aa;color:white;cursor:pointer;border:0}
a {color:#1769aa;text-decoration:none;margin-right:15px}
table {width:100%;border-collapse:collapse;margin-top:20px}
td,th {padding:12px;border-bottom:1px solid #ddd;text-align:left}
th {background:#17365d;color:white}
.total {font-size:24px;color:#1769aa;font-weight:bold}
</style>
"""

LOGIN = """
<!DOCTYPE html>
<html><head><title>Ingreso</title>""" + ESTILO + """
</head><body>
<header><h1>CONTROL DE GASTOS</h1></header>
<div class="contenedor" style="max-width:400px">
<h2>Iniciar sesión</h2>
<form method="POST">
<label>Usuario</label><br>
<input name="usuario" required>
<label>Contraseña</label><br>
<input name="clave" type="password" required>
<button type="submit">Ingresar</button>
</form>
{% if error %}<p style="color:red">{{error}}</p>{% endif %}
</div></body></html>
"""

PAGINA = """
<!DOCTYPE html>
<html><head><title>Mis gastos</title>""" + ESTILO + """
</head><body>
<header><h1>SISTEMA DE CONTROL DE GASTOS</h1></header>
<div class="contenedor">
<h2>Registrar nuevo gasto</h2>
<form method="POST" action="/agregar">
<input type="date" name="fecha" value="{{hoy}}" required>
<input type="text" name="descripcion" placeholder="Descripción del gasto" required>
<input type="number" name="monto" placeholder="Monto en soles" min="0.01" step="0.01" required>
<button type="submit">Guardar gasto</button>
</form>
<hr>
<h2>Reporte de gastos</h2>
<p class="total">Total: S/ {{'%.2f'|format(total)}}</p>
<a href="/excel">Descargar Excel</a>
<a href="/reporte-mensual">Reporte mensual</a>
<a href="/salir">Cerrar sesión</a>
<table>
<tr><th>Fecha</th><th>Descripción</th><th>Monto</th><th>Acción</th></tr>
{% for gasto in gastos %}
<tr>
<td>{{gasto.fecha}}</td>
<td>{{gasto.descripcion}}</td>
<td>S/ {{'%.2f'|format(gasto.monto)}}</td>
<td><a href="/eliminar/{{gasto.id}}" onclick="return confirm('¿Eliminar este gasto?')">Eliminar</a></td>
</tr>
{% endfor %}
</table>
</div></body></html>
"""

@app.route("/", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        if request.form["usuario"] == "admin" and request.form["clave"] == "123":
            session["usuario"] = "admin"
            return redirect("/inicio")
        return render_template_string(LOGIN, error="Usuario o contraseña incorrectos")
    return render_template_string(LOGIN, error=None)

@app.route("/inicio")
def inicio():
    if "usuario" not in session:
        return redirect("/")
    conexion = conectar()
    gastos = conexion.execute("SELECT * FROM gastos ORDER BY fecha DESC, id DESC").fetchall()
    total = conexion.execute("SELECT SUM(monto) FROM gastos").fetchone()[0] or 0
    conexion.close()
    return render_template_string(PAGINA, gastos=gastos, total=total, hoy=date.today().isoformat())
@app.route("/agregar", methods=["POST"])
def agregar():
    if "usuario" not in session:
        return redirect("/")
    fecha = request.form["fecha"]
    descripcion = request.form["descripcion"].strip()
    monto = float(request.form["monto"])
    if not descripcion or monto <= 0:
        return "Datos incorrectos", 400
    conexion = conectar()
    conexion.execute(
        "INSERT INTO gastos (fecha, descripcion, monto) VALUES (?, ?, ?)",
        (fecha, descripcion, monto)
    )
    conexion.commit()
    conexion.close()
    return redirect("/inicio")

@app.route("/eliminar/<int:id>")
def eliminar(id):
    if "usuario" not in session:
        return redirect("/")
    conexion = conectar()
    conexion.execute("DELETE FROM gastos WHERE id = ?", (id,))
    conexion.commit()
    conexion.close()
    return redirect("/inicio")

@app.route("/excel")
def excel():
    if "usuario" not in session:
        return redirect("/")
    conexion = conectar()
    gastos = conexion.execute("SELECT * FROM gastos ORDER BY fecha").fetchall()
    conexion.close()

    libro = Workbook()
    hoja = libro.active
    hoja.title = "Reporte de gastos"
    hoja.append(["ID", "Fecha", "Descripción", "Monto (S/)"])
    for gasto in gastos:
        hoja.append([gasto["id"], gasto["fecha"], gasto["descripcion"], gasto["monto"]])

    archivo = BytesIO()
    libro.save(archivo)
    archivo.seek(0)
    return send_file(archivo, as_attachment=True,
                     download_name="reporte_gastos.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.route("/salir")
def salir():
    session.clear()
    return redirect("/")


@app.route("/reporte-mensual")
def reporte_mensual():
    if "usuario" not in session:
        return redirect("/")

    mes = request.args.get("mes", date.today().strftime("%Y-%m"))

    conexion = conectar()

    gastos = conexion.execute(
        "SELECT * FROM gastos WHERE substr(fecha, 1, 7) = ? ORDER BY fecha DESC",
        (mes,)
    ).fetchall()

    total = conexion.execute(
        "SELECT COALESCE(SUM(monto), 0) FROM gastos WHERE substr(fecha, 1, 7) = ?",
        (mes,)
    ).fetchone()[0]

    conexion.close()

    return render_template_string("""
    <!DOCTYPE html>
    <html>
    <head>
        <title>Reporte mensual</title>
    """ + ESTILO + """
    </head>
    <body>
        <header><h1>REPORTE MENSUAL DE GASTOS</h1></header>
        <div class="contenedor">
            <h2>Seleccionar mes</h2>

            <form method="GET">
                <input type="month" name="mes" value="{{mes}}" required>
                <button type="submit">Consultar</button>
            </form>

            <h2>Total del mes</h2>
            <p class="total">S/ {{'%.2f'|format(total)}}</p>

            <table>
                <tr>
                    <th>Fecha</th>
                    <th>Descripción</th>
                    <th>Monto</th>
                </tr>
                {% for gasto in gastos %}
                <tr>
                    <td>{{gasto.fecha}}</td>
                    <td>{{gasto.descripcion}}</td>
                    <td>S/ {{'%.2f'|format(gasto.monto)}}</td>
                </tr>
                {% endfor %}
            </table>

            <p><a href="/inicio">Volver al inicio</a></p>
        </div>
    </body>
    </html>
    """, mes=mes, gastos=gastos, total=total)



if __name__ == "__main__":
    app.run(debug=True)

