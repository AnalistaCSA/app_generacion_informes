import sena_db
from flask import Flask, request, send_file
from flask_cors import CORS
from io import BytesIO
import os
import traceback
from flask import jsonify


app = Flask(__name__)

print("VERSION NUEVA ACTIVADA")

CORS(app)

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

    return jsonify(sena_db.obtener_datos_dashboard())

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)