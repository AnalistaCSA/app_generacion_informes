from . import sena_db
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

    datos = sena_db.obtener_datos()
    

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