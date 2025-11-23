import os
import re
import sys
import requests
from io import StringIO, BytesIO
from datetime import datetime
from flask import (
    Flask, render_template, request, flash, send_file
)
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
import pandas as pd

# =========================
# CONFIG
# =========================

app = Flask(__name__)
app.secret_key = "super-secret-paseu"

# URLs permanentes en GitHub RAW
CSV_URL = "https://raw.githubusercontent.com/isaias2210/reemplazo/main/historial_paseu.csv"
PLANTILLA_URL = "https://raw.githubusercontent.com/isaias2210/reemplazo/main/plantilla.docx"

# Ruta temporal en Render
TEMP_DIR = "/tmp"
HISTORIAL_PATH = os.path.join(TEMP_DIR, "historial_paseu.csv")
PLANTILLA_PATH = os.path.join(TEMP_DIR, "plantilla.docx")
SALIDAS_DIR = TEMP_DIR

# =========================
# DESCARGAR ARCHIVOS PERMANENTES DESDE GITHUB
# =========================

def cargar_csv():
    """Descarga historial desde GitHub RAW si no existe localmente."""
    if not os.path.exists(HISTORIAL_PATH):
        print("Descargando historial CSV...")
        r = requests.get(CSV_URL)
        if r.status_code != 200:
            print("No se pudo descargar historial CSV")
            return pd.DataFrame()
        with open(HISTORIAL_PATH, "w", encoding="utf-8") as f:
            f.write(r.text)
    return pd.read_csv(HISTORIAL_PATH, sep=";")

def guardar_csv(df):
    """Guarda CSV en disco local y además sincroniza a GitHub RAW (opcional futuro)."""
    df.to_csv(HISTORIAL_PATH, sep=";", index=False)

def cargar_plantilla():
    """Descarga plantilla.docx desde GitHub RAW."""
    if not os.path.exists(PLANTILLA_PATH):
        print("Descargando plantilla.docx...")
        r = requests.get(PLANTILLA_URL)
        with open(PLANTILLA_PATH, "wb") as f:
            f.write(r.content)
    return PLANTILLA_PATH

# =========================
# UTILIDADES
# =========================

def limpiar(s: str) -> str:
    return " ".join(s.replace("\t"," ").split())

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
    for i,line in enumerate(lineas):
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
    periodo = periodo.replace(" ","")
    d = {"PRIMER_PAGO":"","SEGUNDO_PAGO":"","TERCER_PAGO":""}
    if periodo.startswith("2-"):
        d["PRIMER_PAGO"] = "✓"
    elif periodo.startswith("3-"):
        d["SEGUNDO_PAGO"] = "✓"
    elif periodo.startswith("4-"):
        d["TERCER_PAGO"] = "✓"
    return d

# =========================
# HISTORIAL
# =========================

def leer_historial():
    if not os.path.exists(HISTORIAL_PATH):
        return [], []
    df = cargar_csv()
    encabezado = list(df.columns)
    filas = df.values.tolist()
    return encabezado, filas

def guardar_en_historial(datos, telefono, periodo):
    df = cargar_csv()
    nueva = {
        "FECHA": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "NOMBRE_ESTUDIANTE": datos["NOMBRE_ESTUDIANTE"],
        "CEDULA_ESTUDIANTE": datos["CEDULA_ESTUDIANTE"],
        "TELEFONO": telefono,
        "ESTADO": "PENDIENTE",
        "NOMBRE_ACUDIENTE": datos["NOMBRE_ACUDIENTE"],
        "CEDULA_ACUDIENTE": datos["CEDULA_ACUDIENTE"],
        "PLANILLA": datos["PLANILLA"],
        "CHEQUE": datos["CHEQUE"],
        "NIVEL": datos["NIVEL"],
        "COLEGIO": datos["COLEGIO"],
        "PERIODO": periodo
    }
    df = pd.concat([df, pd.DataFrame([nueva])], ignore_index=True)
    guardar_csv(df)

# =========================
# DOCX
# =========================

def formatear_variables(doc, datos):
    campos = [
        "CHEQUE","PLANILLA","NIVEL","COLEGIO",
        "NOMBRE_ESTUDIANTE","CEDULA_ESTUDIANTE",
        "NOMBRE_ACUDIENTE","CEDULA_ACUDIENTE",
        "PRIMER_PAGO","SEGUNDO_PAGO","TERCER_PAGO",
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
                                txt = txt.replace(ph, datos.get(k,""))
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
            for k,v in datos.items():
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

@app.route("/", methods=["GET","POST"])
def index():
    texto = ""
    telefono = ""
    filas = []
    mensaje = ""

    if request.method == "POST":
        texto = request.form.get("texto","")
        telefono = request.form.get("telefono","")
        accion = request.form.get("accion")

        filas = parse_tabla_cheques(texto)
        acudiente = extraer_acudiente(texto)

        if not filas:
            mensaje = "No se detectaron filas en el texto."

        if accion == "generar" and filas:
            seleccion = request.form.getlist("fila_idx")
            if not seleccion:
                flash("Debes seleccionar al menos un cheque.","error")
            else:
                rutas = []
                for idx in seleccion:
                    try: i = int(idx)
                    except: continue
                    if 0 <= i < len(filas):
                        rutas.append(rellenar_docx(filas[i], acudiente, telefono))

                if rutas:
                    archivos = [os.path.basename(r) for r in rutas]
                    return render_template("descargas.html", archivos=archivos)

                flash("No se pudo generar ningún documento.","error")

    return render_template("index.html", texto=texto, telefono=telefono, filas=filas, mensaje=mensaje)

@app.route("/descargar/<filename>")
def descargar(filename):
    ruta = os.path.join(TEMP_DIR, filename)
    return send_file(ruta, as_attachment=True)

@app.route("/historial")
def historial():
    encabezado, filas = leer_historial()
    return render_template("historial.html", encabezado=encabezado, filas=filas)

@app.route("/buscar", methods=["GET","POST"])
def buscar():
    resultados=[]
    cedula=""
    encabezado, filas = leer_historial()

    if request.method=="POST":
        cedula = request.form.get("cedula","").strip()
        if cedula and encabezado:
            idx = {n:i for i,n in enumerate(encabezado)}
            for r in filas:
                if len(r)>=len(encabezado) and r[idx["CEDULA_ESTUDIANTE"]]==cedula:
                    resultados.append(r)

    return render_template("buscar.html", encabezado=encabezado, resultados=resultados, cedula=cedula)

@app.route("/reemplazo", methods=["GET","POST"])
def reemplazo():
    encabezado, filas = leer_historial()
    pendientes=[]
    cedula=""
    mensaje=""

    if request.method=="POST" and encabezado:
        cedula = request.form.get("cedula","").strip()
        accion = request.form.get("accion")
        idx = {n:i for i,n in enumerate(encabezado)}

        if accion=="buscar":
            for r in filas:
                if (len(r)>=len(encabezado)
                    and r[idx["CEDULA_ESTUDIANTE"]]==cedula
                    and r[idx["ESTADO"]]!="REEMPLAZO RECIBIDO"):
                    pendientes.append(r)

    return render_template("reemplazo.html",
                           encabezado=encabezado,
                           pendientes=pendientes,
                           cedula=cedula,
                           mensaje=mensaje)

if __name__=="__main__":
    app.run(host="0.0.0.0", port=5000)
