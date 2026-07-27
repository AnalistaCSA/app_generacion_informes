import sena_db
from flask import Flask
from flask_cors import CORS
import os

app = Flask(__name__)

print("VERSION NUEVA ACTIVADA")

CORS(app)

sena_db.registrar_rutas(app)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)