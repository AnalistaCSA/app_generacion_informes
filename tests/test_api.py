import pytest
from unittest.mock import patch

from backend.app import app


@pytest.fixture
def client():
    app.config["TESTING"] = True

    with app.test_client() as client:
        yield client


@patch("backend.app.obtener_datos")
def test_datos_endpoint(mock_obtener_datos, client):

    mock_obtener_datos.return_value = [
        {
            "ec5_uuid": "123",
            "title": "Registro prueba",
            "created_at": "2026-01-01",
            "7_CIUDAD": "Bogotá",
            "8_DEPARTAMENTO": "Cundinamarca",
            "66_CAPACIDAD_UPS_KVA": "10",
            "69_NUMERO_DE_SERIE_D": "SN001",
            "11_CODIGO_TECNICO": "TEC001",
            "5_NOMBRE_SEDE": "Sede prueba",
            "6_DIRECCION": "Dirección prueba",
            "2_ID_SEDE": "001"
        }
    ]

    response = client.get("/datos")

    assert response.status_code == 200

    data = response.get_json()

    assert len(data) == 1

    assert data[0]["title"] == "Registro prueba"


@patch("backend.app.generar_excel")
def test_generar_error(mock_generar_excel, client):

    mock_generar_excel.return_value = None

    response = client.post(
        "/generar",
        json={"ids": ["123"]}
    )

    assert response.status_code == 500


@patch("backend.app.generar_excel")
def test_generar_excepcion(mock_generar_excel, client):

    mock_generar_excel.side_effect = Exception(
        "Error controlado"
    )

    response = client.post(
        "/generar",
        json={"ids": ["123"]}
    )

    assert response.status_code == 500