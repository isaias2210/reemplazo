import os
import re
import sys
import json
import requests
from io import StringIO, BytesIO
from datetime import datetime
from flask import (
    Flask, render_template, request, flash, send_file
)
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

# =========================
# CONFIG
# =========================

app = Flask(__name__)
app.secret_key = "super-secret-paseu"

# PLANTILLA en GitHub RAW (si quieres seguir usándola así)
PLANTILLA_URL = "https://raw.githubusercontent.com/isaias2210/reemplazo/main/plantilla.docx"

# Ruta temporal en Render
TEMP_DIR = "/tmp"
PLANTILLA_PATH = os.path.join(TEMP_DIR, "plantilla.docx")

# Google Sheets
DEFAULT_SHEET_ID = "1ahOfi2rjJKh7onjiqcGXRnT-ZysoKaRtuSd0iPNCqUw"
SHEET_NAME = "ifarhu"  # nombre de la pestaña en el Sheet


# =========================
# GOOGLE SHEETS HELPERS
# =========================

def get_sheet():
    """
    Devuelve el worksheet de Google Sheets.
    Usa:
      - GOOGLE_CREDENTIALS  (JSON de service account) en variables de entorno
      - SHEET_ID (opcional) si quieres cambiar el sheet desde Render
    """
    creds_json = os.environ.get("GOOGLE_CREDENTIALS") or os.environ.get("GOOGLE_CREDENTIALS_JSON")
    sheet_id = os.environ.get("SHEET_ID", DEFAULT_SHEET_ID)

    if not creds_json:
        raise RuntimeError("Faltan las credenciales en GOOGLE_CREDENTIALS en Render.")

    creds_dict = json.loads(creds_json)

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]

    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    client = gspread.authorize(creds)

    sh = client.open_by_key(sheet_id)
    try:
        ws = sh.worksheet(SHEET_NAME)
    except gspread.exceptions.WorksheetNotFound:
        # Si no existe la pestaña 'ifarhu', usa la primera hoja
        ws = sh.sheet1

    return ws


# =========================
# DESCARGAR PLANTILLA
# =========================

def cargar_plantilla():
    """Descarga plantilla.docx desde GitHub RAW a /tmp si no existe."""
    if not os.path.exists(PLANTILLA_PATH):
        print("Descargando plantilla.docx...")
        r = requests.get(PLANTILLA_URL)
        r.raise_for_status()
        with open(PLANTILLA_PATH, "wb") as f:
            f.write(r.content)
    return PLANTILLA_PATH


# =========================
# UTILIDADES
# =========================

def limpiar(s: str) -> str:
    return " ".join(s.replace("\t", " ").split())


def extraer_acudiente(texto: str):
    t = texto.replace("\t", " ")
    m = re.search(r"R\.Legal:\s*([A-Za-zÁÉÍÓÚÑ ]+?)\s+Cedula:\s*([0-9-]+)", t, re.IGNORECASE)
    if m:
        return {
            "NOMBRE_ACUDIENTE": limpiar(m.group(1)).upper(),
            "CEDULA_ACUDIENTE": m.group(2),
        }
    return {"NOMBRE_ACUDIENTE": "", "CEDULA_ACUDIENTE": ""}


def parse_tabla_cheques(texto: str):
    lineas = [l for l in texto.splitlines() if l.strip()]
    header_idx = None
    for i, line in enumerate(lineas):
        if ("Regional" in line and "Grado" in line and "Centro" in line
                and "ESTUDIANTE" in line):
            header_idx = i
            break
    if header_idx is None:
        return []

    filas = []
    for line in lineas[header_idx+1:]:
        partes = line.split("\t")
        if len(partes) < 13:
            continue

        estudiante = partes[3].strip()
        tokens = estudiante.split()
        cedula = ""
        nombre = ""
        if tokens and re.match(r"\d{1,2}-\d{3,4}-\d{3,4}", tokens[0]):
            cedula = tokens[0]
            nombre = " ".join(tokens[1:]).upper()
        else:
            nombre = estudiante.upper()

        filas.append({
            "grado": partes[1].strip(),
            "centro": partes[2].strip().upper(),
            "cedula": cedula,
            "nombre": nombre,
            "beca": partes[4].strip(),
            "monto": partes[5].strip(),
            "cheque": partes[6].strip(),
            "planilla": partes[9].strip(),
            "fecha": partes[10].strip(),
            "estado_linea": partes[11].strip(),
            "periodo": partes[12].strip(),
        })

    return filas


def periodo_a_checks(periodo: str):
    periodo = periodo.replace(" ", "")
    d = {"PRIMER_PAGO": "", "SEGUNDO_PAGO": "", "TERCER_PAGO": ""}
    if periodo.startswith("2-"):
        d["PRIMER_PAGO"] = "✓"
    elif periodo.startswith("3-"):
        d["SEGUNDO_PAGO"] = "✓"
    elif periodo.startswith("4-"):
        d["TERCER_PAGO"] = "✓"
    return d


# =========================
# HISTORIAL (Google Sheets)
# =========================

def leer_historial():
    """
    Lee todo el historial desde Google Sheets.
    Supone que la primera fila es encabezado.
    """
    try:
        ws = get_sheet()
        datos = ws.get_all_values()
        if not datos:
            return [], []
        encabezado = datos[0]
        filas = datos[1:]
        return encabezado, filas
    except Exception as e:
        print("Error leyendo historial de Sheets:", e)
        return [], []


def guardar_en_historial(datos, telefono, periodo):
    """
    Agrega una fila nueva al historial en Google Sheets.
    Columnas esperadas:
    FECHA, NOMBRE_ESTUDIANTE, CEDULA_ESTUDIANTE, TELEFONO, ESTADO,
    NOMBRE_ACUDIENTE, CEDULA_ACUDIENTE, PLANILLA, CHEQUE, NIVEL, COLEGIO, PERIODO
    """
    try:
        ws = get_sheet()
        nueva = [
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            datos.get("NOMBRE_ESTUDIANTE", ""),
            datos.get("CEDULA_ESTUDIANTE", ""),
            telefono,
            "PENDIENTE",
            datos.get("NOMBRE_ACUDIENTE", ""),
            datos.get("CEDULA_ACUDIENTE", ""),
            datos.get("PLANILLA", ""),
            datos.get("CHEQUE", ""),
            datos.get("NIVEL", ""),
            datos.get("COLEGIO", ""),
            periodo
        ]
        ws.append_row(nueva, value_input_option="USER_ENTERED")
    except Exception as e:
        print("Error guardando historial en Sheets:", e)


# =========================
# DOCX
# =========================

def formatear_variables(doc, datos):
    campos = [
        "CHEQUE", "PLANILLA", "NIVEL", "COLEGIO",
        "NOMBRE_ESTUDIANTE", "CEDULA_ESTUDIANTE",
        "NOMBRE_ACUDIENTE", "CEDULA_ACUDIENTE",
        "PRIMER_PAGO", "SEGUNDO_PAGO", "TERCER_PAGO",
        "FECHA_ACTUAL"
    ]
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p.runs:
                        txt = r.text
                        if not txt:
                            continue
                        for k in campos:
                            ph = f"{{{{{k}}}}}"
                            if ph in txt:
                                txt = txt.replace(ph, datos.get(k, ""))
                        r.text = txt
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        r.bold = True


def rellenar_docx(fila, acudiente, telefono):
    plantilla = cargar_plantilla()
    doc = Document(plantilla)

    datos = {
        "FECHA_ACTUAL": datetime.now().strftime("%d/%m/%Y"),
        "NOMBRE_ESTUDIANTE": fila["nombre"],
        "CEDULA_ESTUDIANTE": fila["cedula"],
        "NIVEL": fila["grado"],
        "COLEGIO": fila["centro"],
        "PLANILLA": fila["planilla"],
        "CHEQUE": fila["cheque"],
        "NOMBRE_ACUDIENTE": acudiente["NOMBRE_ACUDIENTE"],
        "CEDULA_ACUDIENTE": acudiente["CEDULA_ACUDIENTE"],
    }

    datos.update(periodo_a_checks(fila["periodo"]))

    # Párrafos
    for p in doc.paragraphs:
        for r in p.runs:
            txt = r.text
            if not txt:
                continue
            for k, v in datos.items():
                ph = f"{{{{{k}}}}}"
                if ph in txt:
                    txt = txt.replace(ph, v)
            r.text = txt

    formatear_variables(doc, datos)

    # Archivo temporal
    filename = f"{fila['cedula']}_{fila['periodo']}_{fila['cheque']}.docx"
    filename = filename.replace("/", "-")
    ruta = os.path.join(TEMP_DIR, filename)
    doc.save(ruta)

    guardar_en_historial(datos, telefono, fila["periodo"])

    return ruta


# =========================
# RUTAS
# =========================

@app.route("/", methods=["GET", "POST"])
def index():
    texto = ""
    telefono = ""
    filas = []
    mensaje = ""

    if request.method == "POST":
        texto = request.form.get("texto", "")
        telefono = request.form.get("telefono", "")
        accion = request.form.get("accion")

        filas = parse_tabla_cheques(texto)
        acudiente = extraer_acudiente(texto)

        if not filas:
            mensaje = "No se detectaron filas en el texto."

        if accion == "generar" and filas:
            seleccion = request.form.getlist("fila_idx")
            if not seleccion:
                flash("Debes seleccionar al menos un cheque.", "error")
            else:
                rutas = []
                for idx in seleccion:
                    try:
                        i = int(idx)
                    except:
                        continue
                    if 0 <= i < len(filas):
                        rutas.append(rellenar_docx(filas[i], acudiente, telefono))

                if rutas:
                    archivos = [os.path.basename(r) for r in rutas]
                    return render_template("descargas.html", archivos=archivos)

                flash("No se pudo generar ningún documento.", "error")

    return render_template("index.html", texto=texto, telefono=telefono, filas=filas, mensaje=mensaje)


@app.route("/descargar/<filename>")
def descargar(filename):
    ruta = os.path.join(TEMP_DIR, filename)
    return send_file(ruta, as_attachment=True)


@app.route("/historial")
def historial():
    encabezado, filas = leer_historial()
    return render_template("historial.html", encabezado=encabezado, filas=filas)


@app.route("/buscar", methods=["GET", "POST"])
def buscar():
    resultados = []
    cedula = ""
    encabezado, filas = leer_historial()

    if request.method == "POST":
        cedula = request.form.get("cedula", "").strip()
        if cedula and encabezado:
            idx = {n: i for i, n in enumerate(encabezado)}
            for r in filas:
                if len(r) >= len(encabezado) and r[idx["CEDULA_ESTUDIANTE"]] == cedula:
                    resultados.append(r)

    return render_template("buscar.html", encabezado=encabezado, resultados=resultados, cedula=cedula)


@app.route("/reemplazo", methods=["GET", "POST"])
def reemplazo():
    encabezado, filas = leer_historial()
    pendientes = []
    cedula = ""
    mensaje = ""

    if request.method == "POST" and encabezado:
        cedula = request.form.get("cedula", "").strip()
        accion = request.form.get("accion")
        idx = {n: i for i, n in enumerate(encabezado)}

        if accion == "buscar":
            for r in filas:
                if (len(r) >= len(encabezado)
                    and r[idx["CEDULA_ESTUDIANTE"]] == cedula
                    and r[idx["ESTADO"]] != "REEMPLAZO RECIBIDO"):
                    pendientes.append(r)

    return render_template(
        "reemplazo.html",
        encabezado=encabezado,
        pendientes=pendientes,
        cedula=cedula,
        mensaje=mensaje
    )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
