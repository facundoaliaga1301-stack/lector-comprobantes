import pytesseract
from PIL import Image

# Ruta manual al ejecutable de Tesseract
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

img = Image.open("factura.jpg")  # asegurate de que el archivo exista en tu carpeta
texto = pytesseract.image_to_string(img)

print("Texto extraído:")
print(texto)
