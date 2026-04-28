from flask import Flask, render_template, request
import fitz
import os
import re
import json
import base64
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import google.generativeai as genai
from PIL import Image
import io

app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

SHEET_ID = "17yVw5YF4MY9Hi5dCYl9zGh0m7k3XsJIn7rUTwbzy6LY"
SHEET_NAME = "Hoja 1"

genai.configure(api_key=os.environ.get("GEMINI_API_KEY"))

def pdf_to_images(filepath):
    doc = fitz.open(filepath)
    images = []
    for page in doc:
        mat = fitz.Matrix(2, 2)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        images.append(img)
    return images

def ocr_with_gemini(filepath):
    model = genai.GenerativeModel("gemini-2.0-flash")
    
    prompt = """Analizá este comprobante bancario o documento financiero y extraé los siguientes datos en formato JSON.
Si no encontrás algún campo, dejalo como string vacío "".
Devolvé SOLO el JSON, sin texto adicional, sin markdown, sin explicaciones.

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
        images = pdf_to_images(filepath)
        if not images:
            return {}
        img = images[0]
    else:
        img = Image.open(filepath)

    response = model.generate_content([prompt, img])
    
    try:
        text = response.text.strip()
        text = re.sub(r"```json|```", "", text).strip()
        return json.loads(text)
    except:
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
        for file in files:
            if file and file.filename:
                filename = file.filename
                filepath = os.path.join(UPLOAD_FOLDER, filename)
                file.save(filepath)
                data = ocr_with_gemini(filepath)
                parsed = [{"Campo": k.replace("_", " ").title(), "Valor": v} for k, v in data.items()]
                results.append({"filename": filename, "parsed": parsed})
                try:
                    send_to_google_sheets(data, filename)
                except Exception as e:
                    print("Error Google Sheets:", e)
    return render_template("index.html", results=results, banco=None, operacion=None)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port, debug=False)