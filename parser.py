import fitz
import base64
import re


def leer_pdf(pdf_base64):

    pdf_bytes = base64.b64decode(pdf_base64)

    doc = fitz.open(stream=pdf_bytes, filetype="pdf")

    texto = ""

    for page in doc:
        texto += page.get_text()

    texto = texto.upper()

    resultado = {
        "esBrillaseo": "BRILLASEO" in texto,
        "empresa": "BRILLASEO" if "BRILLASEO" in texto else None,
        "texto": texto
    }

    solicitud = re.search(r"SOLICITUD DE SERVICIO.*?(\d{6,10})", texto, re.S)

    if solicitud:
        resultado["solicitud"] = solicitud.group(1)

    sede = re.search(r"TULUA|CALI|BUGA|PALMIRA|JAMUNDI", texto)

    if sede:
        resultado["sede"] = sede.group(0)

    return resultado