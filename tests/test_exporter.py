import io
import pytest
import pandas as pd
from exporter import build_excel


def _sample_lines():
    return [
        {
            "proyecto": "TestProject",
            "of_number": "OF001",
            "n_pedido": "2026000001",
            "cliente": "C001526 VOESTALPINE",
            "articulo": "PERFIL C",
            "ref_pedido": "REF001",
            "desarrollo": "309",
            "espesor": "3,00",
            "calidad": "GL1951",
            "nota": "-63000",
            "total_piezas": 3,
            "kilos_teoricos": 100.0,
            "fecha_doc": "2026-05-13",
            "paquete_num": 1,
            "paquete_num_of": 1,
            "kilos_paquete": 100.0,
            "linea": 1,
            "piezas": 3,
            "longitud": 12.713,
            "marca": "11889",
        }
    ]


def test_build_excel_returns_bytes():
    result = build_excel(_sample_lines())
    assert isinstance(result, bytes)
    assert len(result) > 0


def test_build_excel_correct_columns():
    result = build_excel(_sample_lines())
    df = pd.read_excel(io.BytesIO(result))
    # Verify exported columns match EXPORT_COLUMNS specification
    expected_columns = ["N.Pedido", "Paquete_Num", "OF", "Paquete_Num_OF", "Línea", "Marca", "Piezas", "Artículo"]
    for col in expected_columns:
        assert col in df.columns, f"Expected column '{col}' not found in {list(df.columns)}"


def test_build_excel_correct_data():
    result = build_excel(_sample_lines())
    df = pd.read_excel(io.BytesIO(result))
    assert df.iloc[0]["OF"] == "OF001"
    assert df.iloc[0]["Marca"] == 11889


def test_build_excel_empty():
    result = build_excel([])
    df = pd.read_excel(io.BytesIO(result))
    assert len(df) == 0
    assert "OF" in df.columns
