import os
import pytest
import database
import config


@pytest.fixture(autouse=True)
def use_temp_db(tmp_path, monkeypatch):
    monkeypatch.setattr(config, "DB_PATH", str(tmp_path / "test.db"))
    database.init_db()


def test_create_and_get_project():
    pid = database.create_project("TestProject")
    projects = database.get_projects()
    assert any(p["name"] == "TestProject" for p in projects)
    assert database.get_project_id("TestProject") == pid


def test_duplicate_project_raises():
    database.create_project("Dup")
    with pytest.raises(Exception):
        database.create_project("Dup")


def test_save_packing_list_new():
    pid = database.create_project("P1")
    data = _sample_data("OF001")
    pl_id, replaced = database.save_packing_list(pid, data)
    assert pl_id > 0
    assert replaced is False
    lists = database.get_project_packing_lists(pid)
    assert lists[0]["of_number"] == "OF001"


def test_save_packing_list_replace():
    pid = database.create_project("P2")
    database.save_packing_list(pid, _sample_data("OF002"))
    _, replaced = database.save_packing_list(pid, _sample_data("OF002"))
    assert replaced is True
    lists = database.get_project_packing_lists(pid)
    assert len(lists) == 1  # still only one entry


def test_get_project_lines():
    pid = database.create_project("P3")
    database.save_packing_list(pid, _sample_data("OF003"))
    lines = database.get_project_lines(pid)
    assert len(lines) == 2
    assert lines[0]["marca"] == "11889"


def _sample_data(of_number: str) -> dict:
    return {
        "of_number": of_number,
        "n_pedido": "2026000001",
        "articulo": "P517780300GL1951 PERFIL C",
        "ref_pedido": "4500096097",
        "nota": "-63000",
        "total_piezas": 2,
        "kilos_teoricos": 100.0,
        "paquetes": [
            {
                "paquete_num": 1,
                "kilos_paquete": 100.0,
                "lineas": [
                    {"linea": 1, "piezas": 1, "longitud": 12.713, "marca": "11889"},
                    {"linea": 2, "piezas": 1, "longitud": 9.803,  "marca": "11919"},
                ],
            }
        ],
    }
