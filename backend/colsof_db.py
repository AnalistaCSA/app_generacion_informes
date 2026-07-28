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

API_URL = "https://five.epicollect.net/api/export/entries/ups-colsof?form_ref=055828c4bea94d6fa8a93f8b98adb69f_6a563f807e24e"

with open("backend/data/tecnicos.json", "r", encoding="utf-8") as f:
    tecnicos = json.load(f)

headers = {
    "User-Agent": "Mozilla/5.0"
}

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


def generar_excel(seleccionados=None):

    """
    Genera informes técnicos en formato Excel a partir de los
    registros seleccionados por el usuario.

    También descarga e inserta automáticamente las evidencias
    fotográficas asociadas a cada registro.

    Args:
        seleccionados (list): Lista de identificadores UUID.

    Returns:
        BytesIO: Archivo ZIP con los informes generados.
    """
        
    try:

        print("Consultando API...")

        datos = obtener_datos()
        

        if not datos:
            print("No hay datos, se cancela ejecución")
            return None

        print(f"Registros: {len(datos)}")

        archivos = []

        for item in datos:

            #filtro de seleccion checkbox
            if seleccionados and item.get("ec5_uuid") not in seleccionados:
                continue

            #FUNCION PARA INSERTAR IMAGENES
            def insertar_imagen(ws, item, campo, celda):
                foto1 = item.get(campo)

                if foto1:
                    try:

                        #sacar URL correctamente
                        if isinstance(foto1, list):
                            url = foto1[0].get("file")
                        else:
                            url = foto1

                        #optimizar imagen (más liviana)
                        url = url.replace("entry_original", "entry_small")
                        
                        response = requests.get(foto1, verify=False, timeout=10)

                        if response.status_code == 200:
                            img = Image(BytesIO(response.content))
                            img.width = 358
                            img.height = 340 
                            ws.add_image(img, celda)
                        else:
                            print("No se pudo descargar imagen")
                    except Exception as e:
                        print("Error con imagen:", e)
                else:
                    ws[celda]="N/A"

            print("Procesando: ", item.get("title"))

            wb = load_workbook("formato/formato_informe_ups.xlsx")

            dt_generales = wb["DATOS GENERALES"]
            evi_mantenimiento = wb["EVIDENCIA DEL MANTENIMIENTO"]
            med_entradas_ups = wb["MEDICIONES DE ENTRADA"]
            med_salida_ups = wb["MEDICIONES DE SALIDA"]
            novedades = wb["NOVEDADES"]

            # ***********  escribir datos HOJA DATOS GENERALES" ***************
            dt_generales["B9"] = item.get("3_NOMBRE_OFICINA")
            dt_generales["S9"] = item.get("2_SBAN")
            dt_generales["B10"] = item.get("created_at", "").split("T")[0]
            dt_generales["M10"] = item.get("6_DIRECCION")
            dt_generales["B11"] = item.get("4_REGIONAL")
            dt_generales["M11"] = item.get("7_CIUDAD")
            dt_generales["B12"] = item.get("5_NOMBRE_SEDE")
            dt_generales["M12"] = item.get("8_DEPARTAMENTO")
            dt_generales["B13"] = item.get("9_COORDINADOR__ENCAR")
            dt_generales["M13"] = item.get("10_TELEFONO_DE_ENCAR")

            for tecnico in tecnicos:
        
                if tecnico.get("id") == item.get("11_CODIGO_TECNICO"):
                    dt_generales["C66"] = tecnico.get("nombre")
                    dt_generales["C67"] = tecnico.get("documento")
                    dt_generales["C68"] = tecnico.get("telefono")
                    break


            #Generacion de archivos
            sban = item.get("2_SBAN", "sin_sban")
            sede = item.get("3_NOMBRE_OFICINA", "sin_sede")
            equipo = item.get("34_CAPACIDAD_DE_UPS", "sin_equipo")
            sn = item.get("38_NUMERO_DE_SERIE_U", "sin_sn")

            nombre_archivo = f"informe_mantenimiento_UPS_SBAN_{sban}_{sede}_{equipo}_KVA_SN_{sn}.xlsx"

            # limpiar TODO lo problemático
            nombre_archivo = re.sub(r'[^\w\-.]', '_', nombre_archivo)

            # asegurar extensión correcta
            if not nombre_archivo.lower().endswith(".xlsx"):
                nombre_archivo += ".xlsx"

            with NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                wb.save(tmp.name)

                with open(tmp.name, "rb") as f:
                    contenido = f.read()

            archivos.append((nombre_archivo, contenido))

            print(f"Generado: {nombre_archivo}")

        # varios archivos
        zip_buffer = BytesIO()

        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
            for nombre, archivo in archivos:
                
                if isinstance(archivo, BytesIO):
                    archivo.seek(0)
                    zipf.writestr(nombre, archivo.read())
                else:
                    zipf.writestr(nombre, archivo)

        zip_buffer.seek(0)

        return zip_buffer
    
    except Exception as e:

        print("=========== ERROR EN GENERAR_EXCEL ===========")
        print(traceback.format_exc())

        raise e

def obtener_datos_dashboard():

    datos = obtener_datos()

    datos_livianos = []

    for item in datos:

        datos_livianos.append({
            "ec5_uuid": item.get("ec5_uuid"),
            "title": item.get("title"),
            "created_at": item.get("created_at"),
            "6_CIUDAD": item.get("6_CIUDAD"),
            "7_DEPARTAMENTO": item.get("7_DEPARTAMENTO"),
            "34_CAPACIDAD_DE_UPS": item.get("34_CAPACIDAD_DE_UPS"),
            "38_NUMERO_DE_SERIE_U": item.get("38_NUMERO_DE_SERIE_U"),
            "10_CODIGO_DEL_TECNIC": item.get("10_CODIGO_DEL_TECNIC"),
            "3_NOMBRE_OFICINA": item.get("3_NOMBRE_OFICINA"),
            "8_DIRECCION": item.get("8_DIRECCION"),
            "2_SBAN": item.get("2_SBAN")
        })

    return datos_livianos

# ============================
# RUTAS DEL MODULO COLSOF
# ============================

def registrar_rutas(app):

    @app.route("/colsof/datos", methods=["GET"])
    def datos_colsof():

        return jsonify(obtener_datos_dashboard())


    @app.route("/colsof/generar", methods=["POST"])
    def generar_colsof():

        try:

            data = request.json
            seleccionados = data.get("ids", [])

            archivo = generar_excel(seleccionados)

            if not archivo:
                return {"error": "No se generaron archivos"}, 500

            if isinstance(archivo, tuple):

                nombre, buffer = archivo

                return send_file(
                    BytesIO(buffer),
                    as_attachment=True,
                    download_name=nombre,
                    mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            return send_file(
                archivo,
                as_attachment=True,
                download_name="informes.zip",
                mimetype="application/zip"
            )

        except Exception as e:

            print(traceback.format_exc())

            return {"error": str(e)}, 500