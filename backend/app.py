import sena_db
import requests
from flask import Flask, request, send_file
from flask_cors import CORS
import zipfile
import time
from io import BytesIO
from openpyxl import load_workbook
import urllib3
from openpyxl.drawing.image import Image
import os
import json
import re
import traceback
from tempfile import NamedTemporaryFile
from flask import jsonify
import requests

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

API_URL = "https://five.epicollect.net/api/export/entries/csa-ups-instalacion?form_ref=fff4776480684a35b8765ec74e7c14f8_69c54ba08a99d"

with open("backend/data/tecnicos.json", "r", encoding="utf-8") as f:
    tecnicos = json.load(f)

headers = {
    "User-Agent": "Mozilla/5.0"
}

app = Flask(__name__)

print("VERSION NUEVA ACTIVADA")

CORS(app)

cache_datos = []
ultima_actualizacion = 0

def obtener_datos():

    """
    Consulta la API de EpiCollect y obtiene todos los registros
    disponibles del formulario de instalación de UPS.

    La función implementa paginación automática y cache local
    para reducir el número de consultas realizadas a la API.

    Returns:
        list: Lista de registros obtenidos desde EpiCollect.
    """

    global cache_datos
    global ultima_actualizacion

    # cache 10 minutos
    if time.time() - ultima_actualizacion < 600 and cache_datos:
        print("Usando cache", flush=True)
        return cache_datos

    try:

        todos = []

        # =========================
        # PRIMERA PAGINA
        # =========================

        response = requests.get(
            f"{API_URL}&page=1",
            headers=headers,
            verify=False,
            timeout=30
        )

        print("STATUS PAGINA 1:", response.status_code, flush=True)

        if response.status_code != 200:

            print("ERROR PAGINA 1", flush=True)

            if cache_datos:
                return cache_datos

            return []

        data = response.json()

        total_paginas = data.get("meta", {}).get("last_page", 1)

        print(f"TOTAL PAGINAS: {total_paginas}", flush=True)

        entries = data.get("data", {}).get("entries", [])

        todos.extend(entries)

        # =========================
        # DEMAS PAGINAS
        # =========================

        for page in range(2, total_paginas + 1):

            pagina_cargada = False
            intentos = 0

            while not pagina_cargada and intentos < 10:

                try:

                    print(f"Consultando página {page}", flush=True)

                    response = requests.get(
                        f"{API_URL}&page={page}",
                        headers=headers,
                        verify=False,
                        timeout=30
                    )

                    print(f"STATUS PAGINA {page}: {response.status_code}", flush=True)

                    # RATE LIMIT
                    if response.status_code == 429:

                        intentos += 1

                        print(
                            f"RATE LIMIT PAGINA {page} - intento {intentos}",
                            flush=True
                        )

                        time.sleep(20)

                        continue

                    # ERROR GENERAL
                    if response.status_code != 200:

                        print(f"ERROR PAGINA {page}", flush=True)

                        break

                    data = response.json()

                    entries = data.get("data", {}).get("entries", [])

                    print(
                        f"REGISTROS PAGINA {page}: {len(entries)}",
                        flush=True
                    )

                    todos.extend(entries)

                    print(
                        f"TOTAL ACUMULADO: {len(todos)}",
                        flush=True
                    )

                    pagina_cargada = True

                    time.sleep(5)

                except Exception as e:

                    intentos += 1

                    print(
                        f"ERROR PAGINA {page} - intento {intentos}: {e}",
                        flush=True
                    )

                    time.sleep(10)

        cache_datos = todos
        ultima_actualizacion = time.time()

        print(f"TOTAL REGISTROS FINAL: {len(todos)}", flush=True)

        return todos

    except Exception as e:

        print("ERROR GENERAL:", str(e), flush=True)

        if cache_datos:
            return cache_datos

        return []

@app.route("/generar", methods=["POST"])
def generar():

    """
    Endpoint encargado de generar informes técnicos.

    Recibe una lista de identificadores de registros y
    devuelve un archivo ZIP con los informes generados.

    Returns:
        Response: Archivo ZIP descargable.
    """

    try:
        data = request.json
        seleccionados = data.get("ids", [])

        archivo = sena_db.generar_excel(seleccionados)

        if not archivo:
            return {"error": "No se generaron archivos"}, 500

        # un solo archivo
        if isinstance(archivo, tuple):
            nombre, buffer = archivo

            return send_file(
                BytesIO(buffer),
                as_attachment=True,
                download_name=nombre,
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        # varios (zip)
        return send_file(
            archivo,
            as_attachment=True,
            download_name="informes.zip",
            mimetype="application/zip"
        )
    
    except Exception as e:

        print("=========== ERROR GENERANDO INFORME ===========")
        print(traceback.format_exc())

        return {
            "error": str(e)
        }, 500

@app.route("/datos", methods=["GET"])
def datos():

    """
    Endpoint encargado de consultar registros disponibles.

    Devuelve una versión simplificada de los datos obtenidos
    desde EpiCollect para ser consumidos por el frontend.

    Returns:
    JSON: Lista de registros disponibles.
    """

    datos = obtener_datos()
    

    datos_livianos = []

    for item in datos:

        datos_livianos.append({
            "ec5_uuid": item.get("ec5_uuid"),
            "title": item.get("title"),
            "created_at": item.get("created_at"),
            "7_CIUDAD": item.get("7_CIUDAD"),
            "8_DEPARTAMENTO": item.get("8_DEPARTAMENTO"),
            "66_CAPACIDAD_UPS_KVA": item.get("66_CAPACIDAD_UPS_KVA"),
            "69_NUMERO_DE_SERIE_D": item.get("69_NUMERO_DE_SERIE_D"),
            "11_CODIGO_TECNICO": item.get("11_CODIGO_TECNICO"),
            "5_NOMBRE_SEDE": item.get("5_NOMBRE_SEDE"),
            "6_DIRECCION": item.get("6_DIRECCION"),
            "2_ID_SEDE": item.get("2_ID_SEDE")
        })

    return jsonify(datos_livianos)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)