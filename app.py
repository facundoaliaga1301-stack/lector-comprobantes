# app.py
from flask import Flask, render_template, request
import pytesseract
from PIL import Image, ImageEnhance, ImageOps
import fitz
import os
import re
import gspread
from oauth2client.service_account import ServiceAccountCredentials

app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Configuración Google Sheets
SHEET_ID = "17yVw5YF4MY9Hi5dCYl9zGh0m7k3XsJIn7rUTwbzy6LY"
CREDENTIALS_FILE = "lector-493203-e081dfa428f4.json"
SHEET_NAME = "Hoja 1"

# Opcional: ruta a tesseract en Windows
# pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

def preprocess_image(img):
    img = ImageOps.grayscale(img)
    enhancer = ImageEnhance.Contrast(img)
    img = enhancer.enhance(2.0)
    return img

def pdf_to_images(filepath, zoom=3):
    doc = fitz.open(filepath)
    images = []
    mat = fitz.Matrix(zoom, zoom)
    for page in doc:
        pix = page.get_pixmap(matrix=mat, alpha=False)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        images.append(img)
    return images

def ocr_from_file(filepath):
    text = ""
    if filepath.lower().endswith(".pdf"):
        images = pdf_to_images(filepath, zoom=3)
        for img in images:
            img = preprocess_image(img)
            text += pytesseract.image_to_string(img, lang="spa", config="--psm 6")
            text += "\n"
    else:
        img = Image.open(filepath)
        img = preprocess_image(img)
        text = pytesseract.image_to_string(img, lang="spa", config="--psm 6")
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
        "Cuenta Origen",
        "N° de Referencia",
        "Titular Origen",
        "CUIT Origen",
        "Fecha",
        "Hora",
        "Cuenta Destino / CBU",
        "Titular Destino",
        "CUIT Destino",
        "Importe",
        "Motivo",
        "Concepto"
    ]
    campos = {k: "" for k in keys}
    lines = [l.strip() for l in text.splitlines() if l.strip()]

    # Pase 1: "Campo: valor" en la misma línea
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

    # Pase 2: títulos verticales (campo en una línea y valor en la siguiente)
    i = 0
    while i < len(lines):
        line = lines[i]
        title = line.replace(":", "").strip()
        next_line = lines[i+1] if i+1 < len(lines) else ""
        if re.fullmatch(r"(Fecha|Hora|Motivo|Concepto|Nº de Referencia|N° de Referencia|Cuenta Origen|Cuenta Destino|CBU/CVU Destino|Titularidad|CUIT / CUIL / CDI|CUIT)", title, flags=re.I):
            if re.search(r"Cuenta Origen", title, re.I):
                campos["Cuenta Origen"] = next_line
            elif re.search(r"Cuenta Destino|CBU/CVU Destino|CBU", title, re.I):
                m = re.search(r"(\d{22})", next_line)
                campos["Cuenta Destino / CBU"] = m.group(1) if m else next_line
            elif re.search(r"Titularidad", title, re.I):
                if campos["Titular Origen"] == "":
                    campos["Titular Origen"] = _clean_trailing_fecha(next_line)
                else:
                    campos["Titular Destino"] = next_line
            elif re.search(r"CUIT", title, re.I):
                m = re.search(r"(\d{2}-\d{7,8}-\d|\d{11})", next_line)
                if m:
                    if campos["CUIT Origen"] == "":
                        campos["CUIT Origen"] = m.group(1)
                    else:
                        campos["CUIT Destino"] = m.group(1)
                else:
                    if campos["CUIT Origen"] == "":
                        campos["CUIT Origen"] = next_line
                    else:
                        campos["CUIT Destino"] = next_line
            elif re.search(r"Fecha", title, re.I):
                campos["Fecha"] = next_line
            elif re.search(r"Hora", title, re.I):
                campos["Hora"] = next_line
            elif re.search(r"Nº de Referencia|N° de Referencia|Referencia", title, re.I):
                campos["N° de Referencia"] = next_line
            elif re.search(r"Importe", title, re.I):
                campos["Importe"] = next_line
            elif re.search(r"Motivo", title, re.I):
                campos["Motivo"] = next_line
            elif re.search(r"Concepto", title, re.I):
                campos["Concepto"] = next_line
            i += 2
            continue
        i += 1

    # Pase 3: completar con búsquedas por patrón global
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

    # Limpieza final: recortar espacios y asegurar que Titular Origen no termine con 'Fecha'
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
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, scope)
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
    results = []  # lista de dicts: {"filename":..., "parsed": [...]}
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
