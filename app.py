import sys
import os
os.environ['PYTHONUNBUFFERED'] = '1'

from flask import Flask, render_template, request
import fitz
import re
import json
import base64
import gspread
import requests
from oauth2client.service_account import ServiceAccountCredentials

app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

SHEET_ID = "17yVw5YF4MY9Hi5dCYl9zGh0m7k3XsJIn7rUTwbzy6LY"
SHEET_NAME = "Hoja 1"
MISTRAL_API_KEY = os.environ.get("MISTRAL_API_KEY")

def pdf_to_base64_images(filepath):
    doc = fitz.open(filepath)
    images = []
    mat = fitz.Matrix(2, 2)
    for page in doc:
        pix = page.get_pixmap(matrix=mat, alpha=False)
        img_bytes = pix.tobytes("png")
        b64 = base64.b64encode(img_bytes).decode("utf-8")
        images.append(b64)
    return images

def image_to_base64(filepath):
    with open(filepath, "rb") as f:
        return base64.b64encode(f.read()).decode("utf-8")

def ocr_with_mistral(filepath):
    prompt = """Analizá este comprobante bancario o documento financiero y extraé los datos en formato JSON.
Si no encontrás algún campo dejalo como string vacío.
Devolvé SOLO el JSON sin texto adicional ni markdown.

{
  "tipo_documento": "",
  "banco": "",
  "fecha": "",
  "hora": "",
  "importe": "",
  "moneda": "",
  "cuenta_origen": "",
  "titular_origen": "",
  "cuit_origen": "",
  "cuenta_destino": "",
  "titular_destino": "",
  "cuit_destino": "",
  "nro_referencia": "",
  "motivo": "",
  "concepto": "",
  "estado": ""
}"""

    if filepath.lower().endswith(".pdf"):
        images_b64 = pdf_to_base64_images(filepath)
        if not images_b64:
            sys.stdout.write("ERROR: PDF sin imagenes\n")
            sys.stdout.flush()
            return {}
        b64 = images_b64[0]
        media_type = "image/png"
    else:
        b64 = image_to_base64(filepath)
        ext = filepath.lower().split(".")[-1]
        media_type = "image/jpeg" if ext in ["jpg", "jpeg"] else "image/png"

    sys.stdout.write("Llamando a Mistral...\n")
    sys.stdout.flush()

    response = requests.post(
        "https://api.mistral.ai/v1/chat/completions",
        headers={
            "Authorization": f"Bearer {MISTRAL_API_KEY}",
            "Content-Type": "application/json"
        },
        json={
            "model": "pixtral-12b-latest",
            "messages": [
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "image_url",
                            "image_url": f"data:{media_type};base64,{b64}"
                        },
                        {
                            "type": "text",
                            "text": prompt
                        }
                    ]
                }
            ]
        }
    )

    sys.stdout.write(f"STATUS: {response.status_code}\n")
    sys.stdout.write(f"RESPUESTA: {response.text[:500]}\n")
    sys.stdout.flush()

    try:
        text = response.json()["choices"][0]["message"]["content"].strip()
        text = re.sub(r"```json|```", "", text).strip()
        return json.loads(text)
    except Exception as e:
        sys.stdout.write(f"ERROR PARSEANDO: {e}\n")
        sys.stdout.flush()
        return {}

def get_credentials():
    creds_json = os.environ.get("GOOGLE_CREDENTIALS_JSON")
    if creds_json:
        try:
            return json.loads(creds_json)
        except:
            pass
    if os.path.exists("credentials.json"):
        with open("credentials.json") as f:
            return json.load(f)
    return None

def send_to_google_sheets(data, filename):
    creds_dict = get_credentials()
    if not creds_dict:
        sys.stdout.write("ERROR: Sin credenciales Google\n")
        sys.stdout.flush()
        return
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    client = gspread.authorize(creds)
    sheet = client.open_by_key(SHEET_ID).worksheet(SHEET_NAME)

    headers = ["Archivo", "Tipo", "Banco", "Fecha", "Hora", "Importe", "Moneda",
               "Cuenta Origen", "Titular Origen", "CUIT Origen",
               "Cuenta Destino", "Titular Destino", "CUIT Destino",
               "N° Referencia", "Motivo", "Concepto", "Estado"]

    try:
        values = sheet.get_all_values()
    except:
        values = []

    if not values or (len(values) == 1 and all(c == "" for c in values[0])):
        sheet.append_row(headers)

    row = [
        filename,
        data.get("tipo_documento", ""),
        data.get("banco", ""),
        data.get("fecha", ""),
        data.get("hora", ""),
        data.get("importe", ""),
        data.get("moneda", ""),
        data.get("cuenta_origen", ""),
        data.get("titular_origen", ""),
        data.get("cuit_origen", ""),
        data.get("cuenta_destino", ""),
        data.get("titular_destino", ""),
        data.get("cuit_destino", ""),
        data.get("nro_referencia", ""),
        data.get("motivo", ""),
        data.get("concepto", ""),
        data.get("estado", "")
    ]
    sheet.append_row(row)

@app.route("/", methods=["GET", "POST"])
def index():
    results = []
    if request.method == "POST":
        files = request.files.getlist("file")
        sys.stdout.write(f"ARCHIVOS RECIBIDOS: {len(files)}\n")
        sys.stdout.flush()
        for file in files:
            sys.stdout.write(f"ARCHIVO: {file.filename}\n")
            sys.stdout.flush()
            if file and file.filename:
                filename = file.filename
                filepath = os.path.join(UPLOAD_FOLDER, filename)
                file.save(filepath)
                data = ocr_with_mistral(filepath)
                sys.stdout.write(f"DATA FINAL: {data}\n")
                sys.stdout.flush()
                parsed = [{"Campo": k.replace("_", " ").title(), "Valor": v} for k, v in data.items()]
                results.append({"filename": filename, "parsed": parsed})
                try:
                    send_to_google_sheets(data, filename)
                except Exception as e:
                    sys.stdout.write(f"Error Google Sheets: {e}\n")
                    sys.stdout.flush()
    return render_template("index.html", results=results, banco=None, operacion=None)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port, debug=False)