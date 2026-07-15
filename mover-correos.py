import os
import time
import base64
import tempfile

import pythoncom
import win32com.client

from parser import leer_pdf

# ==========================================
# CONFIGURACIÓN
# ==========================================

REMITENTE = "WF-BATCH@comfandi.com.co"
CARPETA_DESTINO = "Brillaseo"

INTERVALO = 30  # segundos


# ==========================================
# BUSCAR CARPETA
# ==========================================

def obtener_carpeta(inbox, nombre):

    for carpeta in inbox.Folders:

        if carpeta.Name.lower() == nombre.lower():
            return carpeta

    return None


# ==========================================
# PROCESAR CORREOS
# ==========================================

def revisar_correos():

    pythoncom.CoInitialize()

    outlook = win32com.client.Dispatch(
        "Outlook.Application"
    ).GetNamespace("MAPI")

    inbox = outlook.GetDefaultFolder(6)

    carpeta_destino = obtener_carpeta(
        inbox,
        CARPETA_DESTINO
    )

    if carpeta_destino is None:
        print(f"No existe la carpeta '{CARPETA_DESTINO}'")
        return

    mensajes = inbox.Items

    mensajes.Sort("[ReceivedTime]", True)

    mensajes = mensajes.Restrict("[Unread]=True")

    print(f"\nBuscando correos de {REMITENTE}...\n")

    for correo in mensajes:

        try:

            remitente = ""

            try:
                remitente = correo.SenderEmailAddress.upper()
            except:
                pass

            if REMITENTE.upper() not in remitente:
                continue

            print("=" * 60)
            print("ASUNTO:", correo.Subject)

            if correo.Attachments.Count == 0:
                print("No tiene adjuntos.")
                continue

            mover = False

            for i in range(1, correo.Attachments.Count + 1):

                adjunto = correo.Attachments.Item(i)

                if not adjunto.FileName.lower().endswith(".pdf"):
                    continue

                print("Leyendo:", adjunto.FileName)

                # Crear archivo temporal
                archivo_temp = tempfile.NamedTemporaryFile(
                    delete=False,
                    suffix=".pdf"
                )

                archivo_temp.close()

                adjunto.SaveAsFile(archivo_temp.name)

                with open(archivo_temp.name, "rb") as f:
                    pdf_base64 = base64.b64encode(
                        f.read()
                    ).decode()

                os.remove(archivo_temp.name)

                resultado = leer_pdf(pdf_base64)

                print("Empresa:", resultado.get("empresa"))
                print("Solicitud:", resultado.get("solicitud"))
                print("Sede:", resultado.get("sede"))

                if resultado.get("esBrillaseo"):

                    print("\n>>> BRILLASEO ENCONTRADO <<<")

                    mover = True
                    break

            if mover:

                correo.UnRead = False
                correo.Save()

                correo.Move(carpeta_destino)

                print("Correo movido correctamente.")

            else:

                print("No pertenece a Brillaseo.")

        except Exception as e:

            print("ERROR:", e)

    pythoncom.CoUninitialize()


# ==========================================
# MAIN
# ==========================================

if __name__ == "__main__":

    print("=" * 60)
    print("SERVICIO INICIADO")
    print("=" * 60)

    while True:

        try:

            revisar_correos()

        except Exception as e:

            print("ERROR GENERAL:", e)

        print(f"\nEsperando {INTERVALO} segundos...\n")

        time.sleep(INTERVALO)