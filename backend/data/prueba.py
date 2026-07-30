from openpyxl import load_workbook as lw

#cargar el archivo

wb = lw(r"C:\Users\CSA Área TI\Documents\CSA\Epicollect\Generacion_informes\backend\data\oficinas_ba.xlsx")
ws = wb.active

sban = int(input("ingrese numero de sban: "))

for dato in ws.iter_rows(min_row=2, values_only=True):
    codigo = dato[0]

    if codigo == sban:
        print(f'Sban: {dato[0]}')
        print(f'Nombre oficina: {dato[1]}')
        print(f'Ciudad: {dato[2]}')
        print(f'Departamente: {dato[3]}')
        print(f'Dirección: {dato[4]}')
        break
    