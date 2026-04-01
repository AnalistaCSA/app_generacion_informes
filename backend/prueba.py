
import os
import json


with open(r"C:\Users\CSA Área TI\Documents\CSA\Epicollect\Generacion_informes\backend\data\tecnicos.json", "r", encoding="utf-8") as f:
    tecnicos = json.load(f)

for tecnico in tecnicos:
    
    if tecnico.get("id") == 1:
        nombre_tenico = tecnico.get("nombre")
        print(nombre_tenico)
        break
    else:
        print("El tecnico no existe")