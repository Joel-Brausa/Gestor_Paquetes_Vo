import pytest
from extractor import parse_llm_response


def test_parse_clean_json():
    raw = """{"of_number": "OF001", "paquetes": []}"""
    result = parse_llm_response(raw)
    assert result["of_number"] == "OF001"
    assert result["paquetes"] == []


def test_parse_json_in_code_block():
    raw = '```json\n{"of_number": "OF002", "paquetes": []}\n```'
    result = parse_llm_response(raw)
    assert result["of_number"] == "OF002"


def test_parse_strips_think_tags():
    raw = "<think>internal reasoning</think>\n{\"of_number\": \"OF003\", \"paquetes\": []}"
    result = parse_llm_response(raw)
    assert result["of_number"] == "OF003"


def test_parse_raises_on_invalid_json():
    with pytest.raises(ValueError, match="JSON válido"):
        parse_llm_response("this is not json")


def test_parse_raises_on_missing_of_number():
    with pytest.raises(ValueError, match="of_number"):
        parse_llm_response('{"paquetes": []}')


def test_parse_full_structure():
    raw = """{
        "of_number": "OF26004215",
        "n_pedido": "2026000508 / 1",
        "cliente": "C001526 VOESTALPINE",
        "total_piezas": 28,
        "kilos_teoricos": 2284.0,
        "fecha_doc": "2026-05-13",
        "paquetes": [
            {
                "paquete_num": 1,
                "kilos_paquete": 1285.0,
                "lineas": [
                    {"linea": 1, "piezas": 3, "longitud": 12.713, "marca": "11889"}
                ]
            }
        ]
    }"""
    result = parse_llm_response(raw)
    assert result["paquetes"][0]["lineas"][0]["marca"] == "11889"
    assert result["kilos_teoricos"] == 2284.0
