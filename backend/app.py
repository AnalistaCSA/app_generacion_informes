import requests
import time
from io import BytesIO
from openpyxl import load_workbook
import urllib3
from openpyxl.drawing.image import Image

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

API_URL = "https://five.epicollect.net/api/export/entries/csa-ups-instalacion?form_ref=fff4776480684a35b8765ec74e7c14f8_69c54ba08a99d"

headers = {
    "User-Agent": "Mozilla/5.0"
}

def obtener_datos():
    for intento in range(3):
        try:
            print(f"Intento {intento + 1} de conexión")

            response =requests.get(
                API_URL,
                headers=headers,
                verify=False,
                timeout=10
            )

            response.raise_for_status()

            data = response.json()

            if "data" not in data or "entries" not in data["data"]:
                print("Respuesta invalida de la API")
                return []
            
            return data["data"]["entries"]
        
        except requests.exceptions.RequestException as e:
            print(f"Error en intento {intento + 1} {e}")
            time.sleep(2)

    print("No se pudo encontrar a la API")
    return []


def generar_excel():

    print("Consultando API...")

    datos = obtener_datos()

    if not datos:
        print("No hay datos, se cancela ejecución")
        return

    print(f"Registros: {len(datos)}")

    for item in datos:

        #FUNCION PARA INSERTAR IMAGENES
        def insertar_imagen(ws, item, campo, celda):
            foto1 = item.get(campo)
            if foto1:
                try:
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

        print("Procesando: ", item.get("tittle"))

        wb = load_workbook(r"C:\Users\CSA Área TI\Documents\CSA\Epicollect\Generacion_informes\formato\formato_informe_instalacion_ups.xlsx")

        dt_generales = wb["DATOS GENERALES"]
        evi_instalacion = wb["EVIDENCIA DE LA INSTALACIÓN"]
        med_entradas_ups = wb["MEDICIONES ENTRADA DE UPS"]
        med_salida_ups = wb["MEDICIONES SALIDA DE UPS"]
        display_ups = wb["DISPLAY DE LA UPS"]
        baterias = wb["BATERIAS"]
        novedades = wb["NOVEDADES"]

        # ***********  escribir datos HOJA DATOS GENERALES" ***************
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

        #DATOS GENERALES - DATOS DE LA UPS INSTALADA
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

        #DATOS GENERALES - ÁREA DE UBICACION Y ESTADO DE CONEXIONES
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

        #DATOS GENERALES - INSPECCION DE LAS INSTALACIONES
        dt_generales["S29"] = item.get("249_CANTIDAD_BATERIA")
        dt_generales["I32"] = item.get("33_NUMERO_FASES_ENTR")
        dt_generales["D35"] = item.get("34_CANALIZACION_Y_DI")
        dt_generales["J35"] = item.get("35_LONGITUD_CANALIZA")
        dt_generales["R32"] = item.get("36_NUMERO_FASES_SALI")
        dt_generales["O35"] = item.get("37_CANALIZACION_Y_DI")
        dt_generales["T35"] = item.get("38_LONGITUD_CANALIZA")
        dt_generales["D36"] = item.get("39_OBSERVACIONES_INS")

        #DATOS GENERALES - TABLERO ENTRADA Y SALIDA DE UPS
        #Entrda
        dt_generales["D41"] = item.get("51_BREAKER_ENTRADA_U")
        dt_generales["I41"] = item.get("52_TIPO_ENTRADA_BREA")
        dt_generales["K41"] = item.get("53_MARCA_BREAKER_ENT")
        dt_generales["Q41"] = item.get("54_CAPACIDAD_BREAKER")
        dt_generales["V41"] = item.get("55_CAPACIDAD_DE_CORT")
        dt_generales["D42"] = item.get("48_DPS_ENTRADA")
        dt_generales["I42"] = item.get("49_MARCA_DPS")
        dt_generales["Q42"] = item.get("50_REFERENCIA_DPS")
        #Salida
        dt_generales["D44"] = item.get("60_BREAKER_SALIDA_UP")
        dt_generales["I44"] = item.get("61_TIPO_SALIDA_BREAK")
        dt_generales["K44"] = item.get("62_MARCA_BREAKER_SAL")
        dt_generales["Q44"] = item.get("63_CAPACIDAD_BREAKER")
        dt_generales["V44"] = item.get("64_CAPACIDAD_DE_CORT")
        dt_generales["D45"] = item.get("57_DPS_SALIDA")
        dt_generales["I45"] = item.get("58_MARCA_DPS")
        dt_generales["Q45"] = item.get("59_REFERENCIA_DPS")

        #DATOS SEGUN TIPO DE UPS (MONOFASICA, BIFASICA, TRIFASICA)
        if item.get("70_TIPO_DE_UPS") == "MONOFASICA":
            #Fotos y datos entrada
            dt_generales["E33"] = item.get("72_CALIBRE_ENTRADA_F")
            dt_generales["I33"] = item.get("73_CALIBRE_ENTRADA_N")
            dt_generales["K33"] = item.get("74_CALIBRE_ENTRADA_T")
            dt_generales["E34"] = item.get("75_COLOR_MARQUILLADO")
            
            insertar_imagen(med_entradas_ups, item, "78_FOTO_TENSION_ENTR","A4")
            dt_generales["D49"] = item.get("79_DIGITE_VALOR_MEDI")
            insertar_imagen(med_entradas_ups, item, "80_FOTO_TENSION_ENTR", "A40")
            dt_generales["C50"] = item.get("81_DIGITE_VALOR_MEDI")
            insertar_imagen(med_entradas_ups, item, "82_FOTO_CORRIENTE_EN", "C40")
            dt_generales["L48"] = item.get("83_DIGITE_EL_VALOR_M")
            insertar_imagen(med_entradas_ups, item, "84_FOTO_CORRIENTE_EN", "C58")
            dt_generales["S48"] = item.get("85_DIGITAR_EL_VALOR_")
            insertar_imagen(med_entradas_ups, item, "86_FOTO_CORREINTE_EN", "E58")
            dt_generales["U48"] = item.get("87_DIGITAR_VALOR_ENT")
            med_entradas_ups["A86"] = item.get("88_OBERVACIONES_ENTR")

            #Fotos y datos salida
            dt_generales["R33"] = item.get("89_CALIBRE_SALIDA_FA")
            dt_generales["T33"] = item.get("90_CALIBRE_SALIDA_NE")
            dt_generales["V33"] = item.get("91_CALIBRE_SALIDA_TI")
            dt_generales["R34"] = item.get("92_COLOR_MARQUILLADO")

            insertar_imagen(med_salida_ups, item, "95_FOTO_TENSION_SALI", "A4")
            dt_generales["D53"] = item.get("96_DIGITE_VALOR_MEDI")
            insertar_imagen(med_salida_ups, item, "97_FOTO_TENSION_SALI", "A40")
            dt_generales["M52"] = item.get("98_DIGITE_VALOR_MEDI")
            insertar_imagen(med_salida_ups, item, "99_FOTO_CORRIENTE_SA", "C40")
            dt_generales["L49"] = item.get("100_DIGITE_EL_VALOR_")
            insertar_imagen(med_salida_ups, item, "101_FOTO_CORRIENTE_S", "C58")
            dt_generales["S49"] = item.get("102_DIGITAR_EL_VALOR")
            insertar_imagen(med_salida_ups, item, "103_FOTO_CORREINTE_S", "E58")
            dt_generales["U49"] = item.get("104_DIGITAR_VALOR_SA")
            med_salida_ups["A86"] = item.get("105_OBERVACIONES_SAL")
            
            #valores vacios en ups monofasica
            dt_generales["I34"] = "N/A"
            dt_generales["K34"] = "N/A"
            dt_generales["T34"] = "N/A"
            dt_generales["V34"] = "N/A"
            med_entradas_ups["C4"] = "N/A" 
            med_entradas_ups["E4"] = "N/A"
            med_salida_ups["C4"] = "N/A"
            med_salida_ups["E4"] = "N/A"
            med_entradas_ups["A22"] = "N/A"
            med_entradas_ups["C22"] = "N/A"
            med_entradas_ups["E22"] = "N/A"
            med_salida_ups["A22"] = "N/A"
            med_salida_ups["C22"] = "N/A"
            med_salida_ups["E22"] = "N/A"
            med_entradas_ups["E40"] = "N/A"
            med_salida_ups["E40"] = "N/A"
            med_entradas_ups["A58"] = "N/A"
            med_salida_ups["A58"] = "N/A"
            dt_generales["D48"] = "N/A"           
            dt_generales["F48"] = "N/A"
            dt_generales["H48"] = "N/A"
            dt_generales["F49"] = "N/A"
            dt_generales["H49"] = "N/A"
            dt_generales["N48"] = "N/A"
            dt_generales["Q48"] = "N/A"
            dt_generales["D52"] = "N/A"
            dt_generales["F52"] = "N/A"
            dt_generales["H52"] = "N/A"
            dt_generales["F53"] = "N/A"
            dt_generales["H53"] = "N/A"
            dt_generales["N49"] = "N/A"
            dt_generales["Q49"] = "N/A"

        elif item.get("70_TIPO_DE_UPS") == "BIFASICA":
            #Fotos y datos entrada
            dt_generales["E33"] = item.get("108_CALIBRE_ENTRADA_")
            dt_generales["K33"] = item.get("109_CALIBRE_ENTRADA_")
            dt_generales["E34"] = item.get("110_COLOR_MARQUILLAD")
            dt_generales["I34"] = item.get("111_COLOR_MARQUILLAD")

            insertar_imagen(med_entradas_ups, item, "113_FOTO_TENSION_ENT", "A22")
            dt_generales["D48"] = item.get("114_DIGITE_MEDIDA_TE")
            insertar_imagen(med_entradas_ups, item, "115_FOTO_CORRIENTE_E", "C40")
            dt_generales["L48"] = item.get("116_DIGITE_VALOR_COR")
            insertar_imagen(med_entradas_ups, item, "117_FOTO_CORRIENTE_E", "E40")
            dt_generales["N48"] = item.get("118_DIGITE_VALOR_COR")
            insertar_imagen(med_entradas_ups, item, "119_FOTO_CORREINTE_E", "E58")
            dt_generales["U48"] = item.get("120_DIGITAR_VALOR_EN")
            med_entradas_ups["A86"] = item.get("121_OBERVACIONES_ENT")

            #fotos y datos salida
            dt_generales["R33"] = item.get("122_CALIBRE_SALIDA_F")
            dt_generales["T33"] = item.get("123_CALIBRE_SALIDA_N")
            dt_generales["V33"] = item.get("124_CALIBRE_SALIDA_T")
            dt_generales["R34"] = item.get("125_COLOR_MARQUILLAD")
            dt_generales["T34"] = item.get("126_COLOR_MARQUILLAD")

            insertar_imagen(med_salida_ups, item, "129_FOTO_TENSION_DE_", "A4")
            dt_generales["D53"] = item.get("130_DIGITAR_VALOR_ME")
            insertar_imagen(med_salida_ups, item, "131_FOTO_TENSION_DE_", "C4")
            dt_generales["F53"] = item.get("132_DIGITAR_VALOR_ME")
            insertar_imagen(med_salida_ups, item, "133_FOTO_TENSION_SAL", "A22")
            dt_generales["D52"] = item.get("134_DIGITE_MEDIDA_TE")
            insertar_imagen(med_salida_ups, item, "135_FOTO_TENSION_SAL", "A40")
            dt_generales["M52"] = item.get("136_DIGITE_VALOR_MED")
            insertar_imagen(med_salida_ups, item, "137_FOTO_CORRIENTE_S", "C40")
            dt_generales["L49"] = item.get("138_DIGITE_VALOR_COR")
            insertar_imagen(med_salida_ups, item, "139_FOTO_CORRIENTE_S", "E40")
            dt_generales["N49"] = item.get("140_DIGITE_VALOR_COR")
            insertar_imagen(med_salida_ups, item, "141_FOTO_CORRIENTE_S", "C58")
            dt_generales["S49"] = item.get("142_DIGITAR_EL_VALOR")
            insertar_imagen(med_salida_ups, item, "143_FOTO_CORREINTE_S", "E58")
            dt_generales["U49"] = item.get("144_DIGITAR_VALOR_SA")
            med_salida_ups["A86"] = item.get("145_OBERVACIONES_SAL")

            #valores vacios en ups BIFASICA
            dt_generales["I33"] = "N/A"
            dt_generales["K34"] = "N/A"
            dt_generales["V34"] = "N/A"
            med_entradas_ups["A4"] = "N/A"
            med_entradas_ups["C4"] = "N/A"
            med_entradas_ups["E4"] = "N/A"
            med_entradas_ups["C22"] = "N/A"
            med_entradas_ups["E22"] = "N/A"
            med_entradas_ups["A40"] = "N/A"
            med_entradas_ups["A58"] = "N/A"
            med_entradas_ups["C58"] = "N/A"
            med_salida_ups["E4"] = "N/A"
            dt_generales["H53"] = "N/A"
            dt_generales["F52"] = "N/A"
            dt_generales["H52"] = "N/A"
            med_salida_ups["C22"] = "N/A"
            med_salida_ups["E22"] = "N/A"
            med_salida_ups["A58"] = "N/A"
            dt_generales["Q49"] = "N/A"
            dt_generales["F48"] = "N/A"
            dt_generales["H48"] = "N/A"
            dt_generales["D49"] = "N/A"
            dt_generales["H49"] = "N/A"
            dt_generales["F49"] = "N/A"
            dt_generales["Q48"] = "N/A"
            dt_generales["S48"] = "N/A"
            dt_generales["C50"] = "N/A"

        else:
            #fotos y datos entrada
            dt_generales["E33"] = item.get("148_CALIBRE_ENTRADA_")
            dt_generales["I33"] = item.get("149_CALIBRE_ENTRADA_")
            dt_generales["K33"] = item.get("150_CALIBRE_ENTRADA_")
            dt_generales["E34"] = item.get("151_COLOR_MARQUILLAD")
            dt_generales["I34"] = item.get("152_COLOR_MARQUILLAD")
            dt_generales["K34"] = item.get("153_COLOR_MARQUILLAD")

            insertar_imagen(med_entradas_ups, item, "156_FOTO_TENSION_DE_", "A4")
            dt_generales["D49"] = item.get("157_DIGITAR_VALOR_ME")
            insertar_imagen(med_entradas_ups, item, "158_FOTO_TENSION_DE_", "C4")
            dt_generales["F49"] = item.get("159_DIGITAR_VALOR_ME")
            insertar_imagen(med_entradas_ups, item, "160_FOTO_TENSION_DE_", "E4")
            dt_generales["H49"] = item.get("161_DIGITAR_VALOR_ME")
            insertar_imagen(med_entradas_ups, item, "162_FOTO_TENSION_ENT", "A22")
            dt_generales["D48"] = item.get("163_DIGITE_MEDIDA_TE")
            insertar_imagen(med_entradas_ups, item, "164_FOTO_TENSION_ENT", "C22")
            dt_generales["F48"] = item.get("165_DIGITE_MEDIDA_TE")
            insertar_imagen(med_entradas_ups, item, "166_FOTO_TENSION_ENT", "E22")
            dt_generales["H48"] = item.get("167_DIGITE_MEDIDA_TE")
            insertar_imagen(med_entradas_ups, item, "168_FOTO_TENSION_ENT", "A40")
            dt_generales["C50"] = item.get("169_DIGITE_VALOR_MED")
            insertar_imagen(med_entradas_ups, item, "170_FOTO_CORRIENTE_E", "C40")
            dt_generales["L48"] = item.get("171_DIGITE_VALOR_COR")
            insertar_imagen(med_entradas_ups, item, "172_FOTO_CORRIENTE_E", "E40")
            dt_generales["N48"] = item.get("173_DIGITE_VALOR_COR")
            insertar_imagen(med_entradas_ups, item, "174_FOTO_CORRIENTE_E", "A58")
            dt_generales["Q48"] = item.get("175_DIGITE_VALOR_COR")
            insertar_imagen(med_entradas_ups, item, "176_FOTO_CORRIENTE_E", "C58")
            dt_generales["S48"] = item.get("177_DIGITAR_EL_VALOR")
            insertar_imagen(med_entradas_ups, item, "178_FOTO_CORREINTE_E", "E58")
            dt_generales["U48"] = item.get("179_DIGITAR_VALOR_EN")
            med_entradas_ups["A86"] = item.get("180_OBERVACIONES_ENT")

            #fotos y datos salida
            dt_generales["R33"] = item.get("181_CALIBRE_SALIDA_F")
            dt_generales["T33"] = item.get("182_CALIBRE_SALIDA_N")
            dt_generales["V33"] = item.get("183_CALIBRE_SALIDA_T")
            dt_generales["R34"] = item.get("184_COLOR_MARQUILLAD")
            dt_generales["T34"] = item.get("185_COLOR_MARQUILLAD")
            dt_generales["V34"] = item.get("186_COLOR_MARQUILLAD")
            insertar_imagen(med_salida_ups, item, "189_FOTO_TENSION_DE_", "A4")
            dt_generales["D53"] = item.get("190_DIGITAR_VALOR_ME")
            insertar_imagen(med_salida_ups, item, "191_FOTO_TENSION_DE_", "C4")
            dt_generales["F53"] = item.get("192_DIGITAR_VALOR_ME")
            insertar_imagen(med_salida_ups, item, "193_FOTO_TENSION_DE_", "E4")
            dt_generales["H53"] = item.get("194_DIGITAR_VALOR_ME")
            insertar_imagen(med_salida_ups, item, "195_FOTO_TENSION_SAL", "A22")
            dt_generales["D52"] = item.get("196_DIGITE_MEDIDA_TE")
            insertar_imagen(med_salida_ups, item, "197_FOTO_TENSION_SAL", "C22")
            dt_generales["F52"] = item.get("198_DIGITE_MEDIDA_TE")
            insertar_imagen(med_salida_ups, item, "199_FOTO_TENSION_SAL", "E22")
            dt_generales["H52"] = item.get("200_DIGITE_MEDIDA_TE")
            insertar_imagen(med_salida_ups, item, "201_FOTO_TENSION_SAL", "A40")
            dt_generales["M52"] = item.get("202_DIGITE_VALOR_MED")
            insertar_imagen(med_salida_ups, item, "203_FOTO_CORRIENTE_S", "C40")
            dt_generales["L49"] = item.get("204_DIGITE_VALOR_COR")
            insertar_imagen(med_salida_ups, item, "205_FOTO_CORRIENTE_S", "E40")
            dt_generales["N49"] = item.get("206_DIGITE_VALOR_COR")
            insertar_imagen(med_salida_ups, item, "207_FOTO_CORRIENTE_S", "A58")
            dt_generales["Q49"] = item.get("208_DIGITE_VALOR_COR")
            insertar_imagen(med_salida_ups, item, "209_FOTO_CORRIENTE_S", "C58")
            dt_generales["S49"] = item.get("210_DIGITAR_EL_VALOR")
            insertar_imagen(med_salida_ups, item, "211_FOTO_CORREINTE_S", "E58")
            dt_generales["U49"] = item.get("212_DIGITAR_VALOR_SA")
            med_salida_ups["A86"] = item.get("213_OBERVACIONES_SAL")
        
        if item.get("215_LECTURAS_CORRESP") == "SI":
            dt_generales["P53"] = "X"
        else:
            dt_generales["U53"] ="x"

        if item.get("216_QUE_ALIMENTA_LA_") == "PDU":
            dt_generales["B39"] = "X"
        elif item.get("216_QUE_ALIMENTA_LA_") == "TABLERO REGULADO":
            dt_generales["F39"] = "X"
        elif item.get("216_QUE_ALIMENTA_LA_") == "MULTITOMA":
            dt_generales["K39"] = "X"
        elif item.get("216_QUE_ALIMENTA_LA_") == "RACK DIRECTAMENTE":
            dt_generales["O39"] = "X"
        else:
            dt_generales["R39"] = "X"
            dt_generales["U39"] = item.get("217_CUAL_OTRO")

        insertar_imagen(evi_instalacion, item, "220_PANORAMICA_UBICA", "A3")
        insertar_imagen(evi_instalacion, item, "220_PANORAMICA_UBICA", "B3")
        insertar_imagen(evi_instalacion, item, "220_PANORAMICA_UBICA", "C3")
        insertar_imagen(evi_instalacion, item, "220_PANORAMICA_UBICA", "D3")

        #Generacion de archivos
        titulo = item.get("title", "sin_titulo")
        sede = item.get("5_NOMBRE_SEDE", "sin_sede")

        nombre_archivo = f"informe_instalacion_UPS_{titulo}_{sede}.xlsx"

        # limpiar caracteres especiales en nombre
        for c in ['/', '\\', ':', '*', '?', '"', '<', '>', '|']:
            nombre_archivo = nombre_archivo.replace(c, "_")

        wb.save(nombre_archivo)

        print(f"Generado: {nombre_archivo}")


if __name__ == "__main__":
    generar_excel()