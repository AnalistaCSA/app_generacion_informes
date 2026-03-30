import requests
from openpyxl import load_workbook
import urllib3
from openpyxl.drawing.image import Image

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

API_URL = "https://five.epicollect.net/api/export/entries/csa-ups-instalacion?form_ref=fff4776480684a35b8765ec74e7c14f8_69c54ba08a99d"

headers = {
    "User-Agent": "Mozilla/5.0"
}

def generar_excel():

    print("Consultando API...")

    response = requests.get(API_URL, headers=headers, verify=False)
    datos = response.json()["data"]["entries"]

    print(f"Registros: {len(datos)}")

    for item in datos:

        wb = load_workbook(r"C:\Users\CSA Área TI\Documents\CSA\Epicollect\Generacion_informes\formato\formato_informe_instalacion_ups.xlsx")

        dt_generales = wb["DATOS GENERALES"]
        evi_instalacion = wb["EVIDENCIA DE LA INSTALACIÓN"]
        med_entradas_ups = wb["MEDICIONES ENTRADA DE UPS"]
        med_salida_ups = wb["MEDICIONES SALIDA DE UPS"]
        display_ups = wb["DISPLAY DE LA UPS"]
        baterias = wb["BATERIAS"]
        novedades = wb["NOVEDADES"]

        #escribir datos hoja DATOS GENERALES"
        dt_generales["B9"] = item.get("2_ID_SEDE")
        dt_generales["M9"] = f'{item.get("67_MARCA_UPS")} {item.get("68_MODELO_UPS")}'
        dt_generales["B10"] = item.get("created_at", "").split("T")[0]
        dt_generales["M10"] = item.get("6_DIRECCION")
        dt_generales["B11"] = item.get("4_REGIONAL")
        dt_generales["M11"] = item.get("7_CIUDAD")
        dt_generales["B12"] = item.get("5_NOMBRE_SEDE")
        dt_generales["M12"] = item.get("8_DEPARTAMENTO")
        dt_generales["B13"] = item.get("9_COORDINADOR__ENCAR")
        dt_generales["M13"] = item.get("10_TELEFONO_DE_ENCAR")
        dt_generales["B16"] = item.get("67_MARCA_UPS")
        dt_generales["F16"] = item.get("68_MODELO_UPS")
        dt_generales["J16"] = item.get("66_CAPACIDAD_UPS_KVA")
        dt_generales["M16"] = item.get("33_NUMERO_FASES_ENTR")
        dt_generales["P16"] = item.get("69_NUMERO_DE_SERIE_D")
        dt_generales["S16"] = item.get("13_TIENE_AIRE_ACONDI")
        dt_generales["U16"] = item.get("14_TEMPERATURA_AMBIE")
        dt_generales["B17"] = item.get("41_TIENE_TARJETA_SNM")
        dt_generales["G17"] = item.get("42_MARCA")
        dt_generales["Q17"] = item.get("43_NUMERO_SERIE_SNMP")
        dt_generales["F18"] = item.get("25_NOMBRE_UBICACIN_E")
        if item.get("17_DISTANCIA_AL_LADO")=="SI":
            dt_generales["F21"]="X"
        else: dt_generales["I21"]="X"
        if item.get("18_DISTANCIA_AL_FREN")=="SI":
            dt_generales["F22"]="X"
        else: dt_generales["I22"]="X"
        if item.get("19_DISTANCIA_AL_LADO")=="SI":
            dt_generales["P21"]="X"
        else: dt_generales["T21"]="X"
        if item.get("20_PUERTA_DE_TABLERO")=="SI":
            dt_generales["P22"]="X"
        else: dt_generales["T22"]="X"
        dt_generales["E23"] = item.get("21_OBSERVACIONES_DIS")
        if item.get("24_EQUIPO_EN_CUARTO_")=="SI":
            dt_generales["C26"]="X"
        else: dt_generales["F26"]="X"
        if item.get("16_BUENA_ACCESIBILID")=="SI":
            dt_generales["C27"]="X"
        else: dt_generales["F27"]="X"
        if item.get("23_BUEN_ESTADO_DE_CA")=="SI":
            dt_generales["C29"]="X"
        else: dt_generales["F29"]="X"
        if item.get("26_ALARMAS_ACTIVAS")=="SI":
            dt_generales["C28"]="X"
        else: dt_generales["F28"]="X"
        if item.get("26_ALARMAS_ACTIVAS")=="SI":
            dt_generales["C28"]="X"
            dt_generales["C20"]=item.get("27_INDIQUE_CODIGO_DE")
        else: dt_generales["F28"]="X"
        if item.get("29_CONEXION_ENTRADAS")=="SI":
            dt_generales["M26"]="X"
        else: dt_generales["O26"]="X"
        if item.get("30_CONEXION_SALIDAS_")=="SI":
            dt_generales["M27"]="X"
        else: dt_generales["O27"]="X"
        if item.get("31_CONEXION_DE_BATER")=="SI":
            dt_generales["S26"]="X"
        else: dt_generales["U26"]="X"
        if item.get("32_CONEXION_A_TIERRA")=="SI":
            dt_generales["S27"]="X"
        else: dt_generales["U27"]="X"
        if item.get("44_EQUIPO_EN_GESTION")=="SI":
            dt_generales["S28"]="X"
        else: dt_generales["U28"]="X"
        if item.get("45_PATCH_CORD_CONECT")=="SI":
            dt_generales["M28"]="X"
        else: dt_generales["O28"]="X"
        if item.get("46_CONDUCTORES_EN_BU")=="SI":
            dt_generales["M29"]="X"
        else: dt_generales["O29"]="X"
        dt_generales["S29"] = item.get("248_CANTIDAD_BATERIA")

        titulo = item.get("title", "sin_titulo")
        sede = item.get("5_NOMBRE_SEDE", "sin_sede")

        nombre_archivo = f"informe_{titulo}_{sede}.xlsx"

        # limpiar nombre
        for c in ['/', '\\', ':', '*', '?', '"', '<', '>', '|']:
            nombre_archivo = nombre_archivo.replace(c, "_")

        wb.save(nombre_archivo)

        print(f"Generado: {nombre_archivo}")


if __name__ == "__main__":
    generar_excel()