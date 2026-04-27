from flask import Flask, render_template, request
from PIL import Image, ImageEnhance, ImageOps
import fitz
import os
import re
import json
import base64
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from google.cloud import vision
from google.oauth2 import service_account

app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

SHEET_ID = "17yVw5YF4MY9Hi5dCYl9zGh0m7k3XsJIn7rUTwbzy6LY"
SHEET_NAME = "Hoja 1"

def get_credentials():
    creds_json = os.environ.get("GOOGLE_CREDENTIALS_JSON")
    if creds_json:
        creds_dict = json.loads(creds_json)
        return creds_dict
    return None

def ocr_from_file(filepath):
    creds_dict = get_credentials()
    credentials = service_account.Credentials.from_service_account_info(
        creds_dict,
        scopes=["https://www.googleapis.com/auth/cloud-platform"]
    )
    client = vision.ImageAnnotatorClient(credentials=credentials)
    
    text = ""
    if filepath.lower().endswith(".pdf"):
        doc = fitz.open(filepath)
        for page in doc:
            mat = fitz.Matrix(3, 3)
            pix = page.get_pixmap(matrix=mat, alpha=False)
            img_bytes = pix.tobytes("png")
            image = vision.Image(content=img_bytes)
            response = client.text_detection(image=image)
            if response.text_annotations:
                text += response.text_annotations[0].description + "\n"
    else:
        with open(filepath, "rb") as f:
            img_bytes = f.read()
        image = vision.Image(content=img_bytes)
        response = client.text_detection(image=image)
        if response.text_annotations:
            text = response.text_annotations[0].description
    return text

def find_patterns(text):
    patterns = {}
    m = re.search(r"\b(\d{22})\b", text)
    if m:
        patterns["Cuenta Destino / CBU"] = m.group(1)
    m = re.search(r"\b(\d{2}-\d{7,8}-\d|\d{11})\b", text)
    if m:
        patterns["CUIT"] = m.group(1)
    m = re.search(r"\b(\d{17})\b", text)
    if m:
        patterns["N° de Referencia"] = m.group(1)
    m = re.search(r"\b(\d{2}/\d{2}/\d{4})\b", text)
    if m:
        patterns["Fecha"] = m.group(1)
    m = re.search(r"\b(\d{2}:\d{2}:\d{2})\b", text)
    if m:
        patterns["Hora"] = m.group(1)
    m = re.search(r"(\d{1,3}(?:\.\d{3})*,\d{2}|\d{1,3}(?:,\d{3})*\.\d{2})", text)
    if m:
        patterns["Importe"] = m.group(1)
    return patterns

def _clean_trailing_fecha(value):
    if not value:
        return value
    return re.sub(r"\bFecha\b\.?\s*$", "", value, flags=re.I).strip()

def parse_bbva_transferencias(text):
    keys = [
        "Cuenta Origen", "N° de Referencia", "Titular Origen", "CUIT Origen",
        "Fecha", "Hora", "Cuenta Destino / CBU", "Titular Destino",
        "CUIT Destino", "Importe", "Motivo", "Concepto"
    ]
    campos = {k: "" for k in keys}
    lines = [l.strip() for l in text.splitlines() if l.strip()]

    for i, line in enumerate(lines):
        if ":" in line and not line.endswith(":"):
            left, right = line.split(":", 1)
            left = left.strip()
            right = right.strip()
            if re.search(r"Cuenta Origen", left, re.I):
                m_cuenta = re.search(r"([A-Z]{0,3}\s*\$?\s*\d{3,}-\d{6,}\/\d?)|(\d{18})", right)
                campos["Cuenta Origen"] = m_cuenta.group(0).strip() if m_cuenta else right
                mref = re.search(r"(\d{17})", line)
                if mref:
                    campos["N° de Referencia"] = mref.group(1)
            elif re.search(r"Cuenta Destino|CBU/CVU Destino|CBU", left, re.I):
                m = re.search(r"(\d{22})", right)
                campos["Cuenta Destino / CBU"] = m.group(1) if m else right
                mh = re.search(r"(\d{2}:\d{2}:\d{2})", right)
                if mh:
                    campos["Hora"] = mh.group(1)
            elif re.search(r"Titularidad", left, re.I):
                if campos["Titular Origen"] == "":
                    campos["Titular Origen"] = _clean_trailing_fecha(right)
                else:
                    campos["Titular Destino"] = right
            elif re.search(r"CUIT|CUIL|CDI", left, re.I):
                m_cuit = re.search(r"(\d{2}-\d{7,8}-\d|\d{11})", right)
                if m_cuit:
                    if campos["CUIT Origen"] == "":
                        campos["CUIT Origen"] = m_cuit.group(1)
                    else:
                        campos["CUIT Destino"] = m_cuit.group(1)
                mfecha = re.search(r"(\d{2}/\d{2}/\d{4})", right)
                if mfecha:
                    campos["Fecha"] = mfecha.group(1)
            elif re.search(r"Importe", left, re.I):
                campos["Importe"] = right
            elif re.search(r"Motivo", left, re.I):
                campos["Motivo"] = right
            elif re.search(r"Concepto", left, re.I):
                campos["Concepto"] = right
            elif re.search(r"Referencia|Nº de Referencia|N° de Referencia", left, re.I):
                mref = re.search(r"(\d{17})", right)
                campos["N° de Referencia"] = mref.group(1) if mref else right

    patterns = find_patterns(text)
    for k, v in patterns.items():
        if k == "Cuenta Destino / CBU" and not campos["Cuenta Destino / CBU"]:
            campos["Cuenta Destino / CBU"] = v
        elif k == "N° de Referencia" and not campos["N° de Referencia"]:
            campos["N° de Referencia"] = v
        elif k == "Fecha" and not campos["Fecha"]:
            campos["Fecha"] = v
        elif k == "Hora" and not campos["Hora"]:
            campos["Hora"] = v
        elif k == "Importe" and not campos["Importe"]:
            campos["Importe"] = v
        elif k == "CUIT" and not campos["CUIT Origen"]:
            campos["CUIT Origen"] = v

    for k in campos:
        if isinstance(campos[k], str):
            campos[k] = campos[k].strip()
    campos["Titular Origen"] = _clean_trailing_fecha(campos["Titular Origen"])

    parsed = [{"Campo": k, "Valor": campos[k]} for k in keys]
    return parsed

def parse_text(text, banco, operacion):
    if banco == "BBVA" and operacion == "Transferencias":
        return parse_bbva_transferencias(text)
    else:
        lines = [line.strip() for line in text.split("\n") if line.strip()]
        return [{"Campo": f"Línea {i+1}", "Valor": line} for i, line in enumerate(lines)]

def send_to_google_sheets(data, filename):
    creds_dict = get_credentials()
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    client = gspread.authorize(creds)
    sheet = client.open_by_key(SHEET_ID).worksheet(SHEET_NAME)

    headers = ["Archivo", "Cuenta Origen", "N° de Referencia", "Titular Origen", "CUIT Origen",
               "Fecha", "Hora", "Cuenta Destino / CBU", "Titular Destino", "CUIT Destino",
               "Importe", "Motivo", "Concepto"]
    try:
        values = sheet.get_all_values()
    except Exception:
        values = []

    if not values or len(values) == 0 or (len(values) == 1 and all(c == "" for c in values[0])):
        sheet.append_row(headers)

    lookup = {item["Campo"]: item["Valor"] for item in data}
    row = [filename] + [lookup.get(h, "") for h in headers[1:]]
    sheet.append_row(row)

@app.route("/", methods=["GET", "POST"])
def index():
    results = []
    banco = None
    operacion = None
    if request.method == "POST":
        banco = request.form.get("banco")
        operacion = request.form.get("operacion")
        files = request.files.getlist("file")
        for file in files:
            if file and file.filename:
                filename = file.filename
                filepath = os.path.join(UPLOAD_FOLDER, filename)
                file.save(filepath)
                text = ocr_from_file(filepath)
                parsed = parse_text(text, banco, operacion)
                results.append({"filename": filename, "parsed": parsed})
                try:
                    send_to_google_sheets(parsed, filename)
                except Exception as e:
                    print("Error al enviar a Google Sheets:", e)
    return render_template("index.html", results=results, banco=banco, operacion=operacion)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port, debug=False)