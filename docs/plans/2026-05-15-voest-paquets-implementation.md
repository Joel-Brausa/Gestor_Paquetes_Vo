# Voest Paquets — Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Build a Gradio app that processes Brausa packing list PDFs via an OpenRouter LLM, stores extracted data in SQLite, and exports to Excel organized by project.

**Architecture:** Flat module structure at project root — `config.py`, `database.py`, `extractor.py`, `exporter.py`, `app.py`. SQLite DB auto-created at `data/paquets.db`. Business logic modules are independently testable with pytest; Gradio UI is tested manually.

**Tech Stack:** Python 3.10+, Gradio 5.x, openai (OpenRouter client), pdf2image, Pillow, pandas, openpyxl, sqlite3 (stdlib)

**Design doc:** `docs/plans/2026-05-15-voest-paquets-design.md`

---

## Task 1: Project Setup

**Files:**
- Create: `requirements.txt`
- Create: `tests/__init__.py`
- Modify: `.gitignore` (add `data/` and `OPEN_KEY.txt`)

**Step 1: Create `requirements.txt`**

```
gradio>=5.0
openai>=1.0
pdf2image>=1.17
Pillow>=10.0
pandas>=2.0
openpyxl>=3.1
pytest>=8.0
```

**Step 2: Create tests directory**

```bash
mkdir tests
touch tests/__init__.py
```

**Step 3: Add to `.gitignore`**

Append to `.gitignore` (create if it doesn't exist):
```
data/
OPEN_KEY.txt
__pycache__/
*.pyc
.pytest_cache/
```

**Step 4: Install dependencies**

```bash
pip install -r requirements.txt
```

Expected: all packages install without errors.

**Step 5: Commit**

```bash
git add requirements.txt tests/__init__.py .gitignore
git commit -m "chore: project setup and dependencies"
```

---

## Task 2: config.py

**Files:**
- Create: `config.py`

**Step 1: Create `config.py`**

```python
import os

BASE_DIR = os.path.dirname(__file__)

OPEN_KEY_FILE = os.path.join(BASE_DIR, "OPEN_KEY.txt")
DB_PATH = os.path.join(BASE_DIR, "data", "paquets.db")

OPENROUTER_BASE_URL = "https://openrouter.ai/api/v1"
DEFAULT_MODEL = "google/gemini-2.0-flash-001"

def load_api_key() -> str:
    if os.path.exists(OPEN_KEY_FILE):
        with open(OPEN_KEY_FILE, "r", encoding="utf-8") as f:
            return f.read().strip()
    return os.environ.get("OPENROUTER_API_KEY", "")

def save_api_key(key: str) -> None:
    with open(OPEN_KEY_FILE, "w", encoding="utf-8") as f:
        f.write(key.strip())
```

**Step 2: Commit**

```bash
git add config.py
git commit -m "feat: add config module"
```

---

## Task 3: database.py

**Files:**
- Create: `database.py`

**Step 1: Write `database.py`**

```python
import sqlite3
import os
from datetime import datetime, timezone
from typing import Optional

import config


def _connect() -> sqlite3.Connection:
    os.makedirs(os.path.dirname(config.DB_PATH), exist_ok=True)
    conn = sqlite3.connect(config.DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def init_db() -> None:
    with _connect() as conn:
        conn.executescript("""
            CREATE TABLE IF NOT EXISTS projects (
                id         INTEGER PRIMARY KEY AUTOINCREMENT,
                name       TEXT    UNIQUE NOT NULL,
                created_at TEXT    NOT NULL
            );

            CREATE TABLE IF NOT EXISTS packing_lists (
                id             INTEGER PRIMARY KEY AUTOINCREMENT,
                project_id     INTEGER NOT NULL REFERENCES projects(id),
                of_number      TEXT    NOT NULL,
                n_pedido       TEXT,
                cliente        TEXT,
                articulo       TEXT,
                ref_pedido     TEXT,
                desarrollo     TEXT,
                espesor        TEXT,
                calidad        TEXT,
                nota           TEXT,
                total_piezas   INTEGER,
                kilos_teoricos REAL,
                fecha_doc      TEXT,
                imported_at    TEXT    NOT NULL,
                UNIQUE(project_id, of_number)
            );

            CREATE TABLE IF NOT EXISTS lines (
                id               INTEGER PRIMARY KEY AUTOINCREMENT,
                packing_list_id  INTEGER NOT NULL REFERENCES packing_lists(id),
                paquete_num      INTEGER,
                kilos_paquete    REAL,
                linea            INTEGER,
                piezas           INTEGER,
                longitud         REAL,
                marca            TEXT
            );
        """)


# ── Projects ──────────────────────────────────────────────────────────────────

def create_project(name: str) -> int:
    with _connect() as conn:
        cur = conn.execute(
            "INSERT INTO projects (name, created_at) VALUES (?, ?)",
            (name.strip(), datetime.now(timezone.utc).isoformat()),
        )
        return cur.lastrowid


def get_projects() -> list[dict]:
    with _connect() as conn:
        rows = conn.execute("SELECT id, name FROM projects ORDER BY name").fetchall()
        return [dict(r) for r in rows]


def get_project_id(name: str) -> Optional[int]:
    with _connect() as conn:
        row = conn.execute("SELECT id FROM projects WHERE name = ?", (name,)).fetchone()
        return row["id"] if row else None


# ── Packing Lists ─────────────────────────────────────────────────────────────

def save_packing_list(project_id: int, data: dict) -> tuple[int, bool]:
    """
    Insert or replace a packing list.
    Returns (packing_list_id, was_replaced).
    """
    now = datetime.now(timezone.utc).isoformat()
    of_number = data["of_number"]

    with _connect() as conn:
        existing = conn.execute(
            "SELECT id FROM packing_lists WHERE project_id = ? AND of_number = ?",
            (project_id, of_number),
        ).fetchone()

        replaced = False
        if existing:
            conn.execute("DELETE FROM lines WHERE packing_list_id = ?", (existing["id"],))
            conn.execute("DELETE FROM packing_lists WHERE id = ?", (existing["id"],))
            replaced = True

        cur = conn.execute(
            """INSERT INTO packing_lists
               (project_id, of_number, n_pedido, cliente, articulo, ref_pedido,
                desarrollo, espesor, calidad, nota, total_piezas, kilos_teoricos,
                fecha_doc, imported_at)
               VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (
                project_id, of_number,
                data.get("n_pedido"), data.get("cliente"), data.get("articulo"),
                data.get("ref_pedido"), data.get("desarrollo"), data.get("espesor"),
                data.get("calidad"), data.get("nota"),
                data.get("total_piezas"), data.get("kilos_teoricos"),
                data.get("fecha_doc"), now,
            ),
        )
        pl_id = cur.lastrowid

        for paquete in data.get("paquetes", []):
            for line in paquete.get("lineas", []):
                conn.execute(
                    """INSERT INTO lines
                       (packing_list_id, paquete_num, kilos_paquete, linea, piezas, longitud, marca)
                       VALUES (?,?,?,?,?,?,?)""",
                    (
                        pl_id,
                        paquete.get("paquete_num"),
                        paquete.get("kilos_paquete"),
                        line.get("linea"),
                        line.get("piezas"),
                        line.get("longitud"),
                        line.get("marca"),
                    ),
                )

    return pl_id, replaced


def get_project_packing_lists(project_id: int) -> list[dict]:
    with _connect() as conn:
        rows = conn.execute(
            """SELECT of_number, n_pedido, cliente, articulo, fecha_doc, imported_at
               FROM packing_lists WHERE project_id = ?
               ORDER BY imported_at DESC""",
            (project_id,),
        ).fetchall()
        return [dict(r) for r in rows]


def get_project_lines(project_id: int) -> list[dict]:
    """Return all lines for a project, flat, ready for Excel export."""
    with _connect() as conn:
        rows = conn.execute(
            """SELECT
                   p.name          AS proyecto,
                   pl.of_number, pl.n_pedido, pl.cliente, pl.articulo,
                   pl.ref_pedido, pl.desarrollo, pl.espesor, pl.calidad, pl.nota,
                   pl.total_piezas, pl.kilos_teoricos, pl.fecha_doc,
                   l.paquete_num, l.kilos_paquete, l.linea, l.piezas, l.longitud, l.marca
               FROM lines l
               JOIN packing_lists pl ON pl.id = l.packing_list_id
               JOIN projects p       ON p.id  = pl.project_id
               WHERE pl.project_id = ?
               ORDER BY pl.of_number, l.paquete_num, l.linea""",
            (project_id,),
        ).fetchall()
        return [dict(r) for r in rows]
```

**Step 2: Commit**

```bash
git add database.py
git commit -m "feat: add database module with SQLite schema and CRUD"
```

---

## Task 4: Tests for database.py

**Files:**
- Create: `tests/test_database.py`

**Step 1: Write `tests/test_database.py`**

```python
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
        "cliente": "C001526 VOESTALPINE",
        "articulo": "P517780300GL1951 PERFIL C",
        "ref_pedido": "4500096097",
        "desarrollo": "309",
        "espesor": "3,00",
        "calidad": "GL1951",
        "nota": "-63000",
        "total_piezas": 2,
        "kilos_teoricos": 100.0,
        "fecha_doc": "2026-05-13",
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
```

**Step 2: Run tests**

```bash
pytest tests/test_database.py -v
```

Expected: 5 tests PASS.

**Step 3: Commit**

```bash
git add tests/test_database.py
git commit -m "test: add database module tests"
```

---

## Task 5: extractor.py

**Files:**
- Create: `extractor.py`

**Step 1: Write `extractor.py`**

```python
import base64
import io
import json
import re
from typing import Optional

from openai import OpenAI
from PIL import Image

import config

EXTRACTION_PROMPT = """
Analyze this Brausa packing list PDF image and extract ALL data.
Return ONLY a valid JSON object with this exact structure (no markdown, no explanation):

{
  "of_number": "OF26004215",
  "n_pedido": "2026000508 / 1",
  "cliente": "C001526 VOESTALPINE ALEMAN",
  "articulo": "P517780300GL1951 PERFIL C 160x67,5x67,5x19x19 Esp. 3 mm",
  "ref_pedido": "4500096097 CUBE WAREHOUSE T. D11",
  "desarrollo": "309",
  "espesor": "3,00",
  "calidad": "GL1951 A GALVANIZADO S450GD +Z140 MAC",
  "nota": "-63000",
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
}

Rules:
- of_number: the OF number (e.g. "OF26004215")
- fecha_doc: convert date DD/MM/YYYY to YYYY-MM-DD
- kilos_teoricos and kilos_paquete: decimal numbers (use dot as decimal separator, ignore thousand separators)
- longitud: decimal number in meters
- marca: the Marca column value as string
- Include ALL paquetes and ALL lineas
- Return ONLY the JSON, no other text
"""


def pdf_to_base64_images(pdf_bytes: bytes, dpi: int = 150) -> list[str]:
    from pdf2image import convert_from_bytes
    images = convert_from_bytes(pdf_bytes, dpi=dpi, fmt="PNG")
    result = []
    for img in images:
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        result.append(base64.b64encode(buf.getvalue()).decode("utf-8"))
    return result


def extract_data(pdf_bytes: bytes, model: str, api_key: str) -> dict:
    """
    Convert PDF to images, call OpenRouter LLM, return parsed dict.
    Raises ValueError if extraction fails.
    """
    b64_images = pdf_to_base64_images(pdf_bytes)
    if not b64_images:
        raise ValueError("No se pudo convertir el PDF a imágenes.")

    client = OpenAI(base_url=config.OPENROUTER_BASE_URL, api_key=api_key)

    content = [
        {
            "type": "image_url",
            "image_url": {"url": f"data:image/png;base64,{b}"},
        }
        for b in b64_images
    ]
    content.append({"type": "text", "text": EXTRACTION_PROMPT})

    response = client.chat.completions.create(
        model=model,
        messages=[{"role": "user", "content": content}],
    )
    raw = response.choices[0].message.content or ""
    return parse_llm_response(raw)


def parse_llm_response(text: str) -> dict:
    """
    Extract and validate JSON from LLM response text.
    Raises ValueError if JSON is missing or invalid.
    """
    clean = re.sub(r"<think>.*?</think>", "", text, flags=re.DOTALL).strip()

    code_block = re.search(r"```(?:json)?\s*\n(.*?)```", clean, re.DOTALL)
    json_str = code_block.group(1).strip() if code_block else clean

    try:
        data = json.loads(json_str)
    except json.JSONDecodeError as e:
        raise ValueError(f"Respuesta del LLM no es JSON válido: {e}\n\nRespuesta recibida:\n{text[:500]}")

    required = ["of_number", "paquetes"]
    for field in required:
        if field not in data:
            raise ValueError(f"Campo requerido '{field}' no encontrado en la respuesta del LLM.")

    return data
```

**Step 2: Commit**

```bash
git add extractor.py
git commit -m "feat: add extractor module (PDF to images + OpenRouter LLM)"
```

---

## Task 6: Tests for extractor.py

**Files:**
- Create: `tests/test_extractor.py`

**Step 1: Write `tests/test_extractor.py`**

```python
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
```

**Step 2: Run tests**

```bash
pytest tests/test_extractor.py -v
```

Expected: 6 tests PASS.

**Step 3: Commit**

```bash
git add tests/test_extractor.py
git commit -m "test: add extractor module tests"
```

---

## Task 7: exporter.py

**Files:**
- Create: `exporter.py`

**Step 1: Write `exporter.py`**

```python
import io

import pandas as pd


COLUMN_NAMES = {
    "proyecto":        "Proyecto",
    "of_number":       "OF",
    "n_pedido":        "N.Pedido",
    "cliente":         "Cliente",
    "articulo":        "Artículo",
    "ref_pedido":      "Ref.Pedido",
    "desarrollo":      "Desarrollo",
    "espesor":         "Espesor",
    "calidad":         "Calidad",
    "nota":            "Nota",
    "total_piezas":    "Total Piezas",
    "kilos_teoricos":  "Kilos Teóricos",
    "fecha_doc":       "Fecha Doc",
    "paquete_num":     "N.Paquete",
    "kilos_paquete":   "Kilos Paquete",
    "linea":           "Línea",
    "piezas":          "Piezas",
    "longitud":        "Longitud",
    "marca":           "Marca",
}


def build_excel(lines: list[dict]) -> bytes:
    """
    Build an Excel file from a flat list of line dicts.
    Returns bytes ready for download.
    """
    if not lines:
        df = pd.DataFrame(columns=list(COLUMN_NAMES.values()))
    else:
        df = pd.DataFrame(lines)
        df = df.rename(columns=COLUMN_NAMES)
        cols = [COLUMN_NAMES[k] for k in COLUMN_NAMES if COLUMN_NAMES[k] in df.columns]
        df = df[cols]

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Packing Lists")
    return buf.getvalue()
```

**Step 2: Commit**

```bash
git add exporter.py
git commit -m "feat: add exporter module for Excel generation"
```

---

## Task 8: Tests for exporter.py

**Files:**
- Create: `tests/test_exporter.py`

**Step 1: Write `tests/test_exporter.py`**

```python
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
    assert "OF" in df.columns
    assert "Marca" in df.columns
    assert "Proyecto" in df.columns


def test_build_excel_correct_data():
    result = build_excel(_sample_lines())
    df = pd.read_excel(io.BytesIO(result))
    assert df.iloc[0]["OF"] == "OF001"
    assert df.iloc[0]["Marca"] == "11889"


def test_build_excel_empty():
    result = build_excel([])
    df = pd.read_excel(io.BytesIO(result))
    assert len(df) == 0
    assert "OF" in df.columns
```

**Step 2: Run tests**

```bash
pytest tests/test_exporter.py -v
```

Expected: 4 tests PASS.

**Step 3: Run all tests together**

```bash
pytest tests/ -v
```

Expected: all 15 tests PASS.

**Step 4: Commit**

```bash
git add tests/test_exporter.py
git commit -m "test: add exporter module tests"
```

---

## Task 9: app.py — Gradio UI

**Files:**
- Create: `app.py`

This is the main UI file. It wires together all modules. No unit tests — test manually by running the app.

**Step 1: Write `app.py`**

```python
import os
import gradio as gr
import pandas as pd

import config
import database
import extractor
import exporter

database.init_db()

# ── Helpers ───────────────────────────────────────────────────────────────────

def _project_choices() -> list[str]:
    return [p["name"] for p in database.get_projects()]


def _pl_table(project_name: str):
    if not project_name:
        return pd.DataFrame()
    pid = database.get_project_id(project_name)
    if pid is None:
        return pd.DataFrame()
    rows = database.get_project_packing_lists(pid)
    if not rows:
        return pd.DataFrame(columns=["OF", "N.Pedido", "Cliente", "Artículo", "Fecha doc", "Importado"])
    df = pd.DataFrame(rows)
    df = df.rename(columns={
        "of_number": "OF", "n_pedido": "N.Pedido", "cliente": "Cliente",
        "articulo": "Artículo", "fecha_doc": "Fecha doc", "imported_at": "Importado",
    })
    return df[["OF", "N.Pedido", "Cliente", "Artículo", "Fecha doc", "Importado"]]


# ── Gradio app ────────────────────────────────────────────────────────────────

with gr.Blocks(title="Voest Paquets") as demo:

    # ── Sidebar ───────────────────────────────────────────────────────────────
    with gr.Sidebar():
        gr.Markdown("## Configuración")

        api_key_input = gr.Textbox(
            label="API Key OpenRouter",
            value=config.load_api_key(),
            type="password",
            placeholder="sk-or-...",
        )
        save_key_btn = gr.Button("Guardar API Key", size="sm")
        key_status = gr.Markdown("")

        gr.Markdown("---")
        gr.Markdown("**Modelo OpenRouter**")

        model_input = gr.Textbox(
            label="Modelo activo",
            value=config.DEFAULT_MODEL,
            placeholder="proveedor/modelo:tag",
        )

        @save_key_btn.click(inputs=api_key_input, outputs=key_status)
        def save_key(key):
            config.save_api_key(key)
            return "API Key guardada."

    # ── Main content ──────────────────────────────────────────────────────────
    gr.Markdown("# Voest Paquets")
    gr.Markdown("Gestión de packing lists por proyecto")
    gr.Markdown("---")

    # ── 1. Proyecto activo ────────────────────────────────────────────────────
    gr.Markdown("## 1. Proyecto activo")
    with gr.Row():
        project_dd = gr.Dropdown(
            label="Seleccionar proyecto",
            choices=_project_choices(),
            interactive=True,
            scale=3,
        )
        new_project_name = gr.Textbox(
            label="Nuevo proyecto",
            placeholder="Nombre del proyecto...",
            scale=2,
        )
        create_btn = gr.Button("Crear", scale=1)

    project_status = gr.Markdown("")

    @create_btn.click(
        inputs=[new_project_name, project_dd],
        outputs=[project_dd, new_project_name, project_status],
    )
    def create_project(name, current):
        name = name.strip()
        if not name:
            return gr.update(), gr.update(), "Introduce un nombre."
        try:
            database.create_project(name)
            choices = _project_choices()
            return gr.update(choices=choices, value=name), "", f"Proyecto **{name}** creado."
        except Exception as e:
            return gr.update(), gr.update(), f"Error: {e}"

    gr.Markdown("---")

    # ── 2. Importar packing list ──────────────────────────────────────────────
    gr.Markdown("## 2. Importar packing list")

    pdf_upload = gr.File(label="Subir PDF", file_types=[".pdf"])
    process_btn = gr.Button("Procesar con LLM", variant="primary")
    import_status = gr.Markdown("")
    preview_table = gr.Dataframe(label="Líneas extraídas", visible=False)

    @process_btn.click(
        inputs=[pdf_upload, project_dd, model_input, api_key_input],
        outputs=[import_status, preview_table, preview_table],
    )
    def process_pdf(pdf_file, project_name, model, api_key):
        if not pdf_file:
            return "Sube un archivo PDF primero.", gr.update(visible=False), gr.update()
        if not project_name:
            return "Selecciona o crea un proyecto primero.", gr.update(visible=False), gr.update()
        if not api_key:
            return "Introduce la API Key de OpenRouter en el menú lateral.", gr.update(visible=False), gr.update()

        try:
            with open(pdf_file.name, "rb") as f:
                pdf_bytes = f.read()
            data = extractor.extract_data(pdf_bytes, model, api_key)
        except Exception as e:
            return f"Error al procesar: {e}", gr.update(visible=False), gr.update()

        pid = database.get_project_id(project_name)
        _, replaced = database.save_packing_list(pid, data)

        lines = database.get_project_lines(pid)
        of_lines = [l for l in lines if l["of_number"] == data["of_number"]]
        df_preview = pd.DataFrame(of_lines)

        action = "reemplazado" if replaced else "importado"
        msg = f"{data['of_number']} {action} correctamente. {len(of_lines)} líneas guardadas."
        return msg, gr.update(visible=True, value=df_preview), gr.update()

    gr.Markdown("---")

    # ── 3. Packing lists del proyecto ─────────────────────────────────────────
    gr.Markdown("## 3. Packing lists del proyecto")

    pl_table = gr.Dataframe(label="Documentos importados", interactive=False)
    export_btn = gr.Button("Exportar proyecto a Excel", variant="secondary")
    export_file = gr.File(label="Descarga", visible=False)

    @project_dd.change(inputs=project_dd, outputs=pl_table)
    def refresh_table(project_name):
        return _pl_table(project_name)

    @export_btn.click(inputs=project_dd, outputs=[export_file, export_file])
    def export_excel(project_name):
        if not project_name:
            return gr.update(visible=False), gr.update()
        pid = database.get_project_id(project_name)
        lines = database.get_project_lines(pid)
        excel_bytes = exporter.build_excel(lines)

        out_path = os.path.join(config.BASE_DIR, f"{project_name}.xlsx")
        with open(out_path, "wb") as f:
            f.write(excel_bytes)
        return gr.update(visible=True, value=out_path), gr.update()


if __name__ == "__main__":
    demo.launch()
```

**Step 2: Run the app manually**

```bash
python app.py
```

Expected: Gradio launches at `http://127.0.0.1:7860`.

**Manual testing checklist:**
- [ ] Crear un proyecto nuevo y verlo en el dropdown
- [ ] Seleccionar el proyecto recién creado
- [ ] Subir `packing list.pdf` y hacer clic en "Procesar con LLM"
- [ ] Verificar que la tabla preview muestra las líneas extraídas
- [ ] Verificar que la sección 3 muestra el packing list importado
- [ ] Volver a subir el mismo PDF → mensaje debe decir "reemplazado"
- [ ] Clic en "Exportar proyecto a Excel" y abrir el archivo generado

**Step 3: Commit**

```bash
git add app.py
git commit -m "feat: add Gradio UI app"
```

---

## Task 10: Final integration commit

**Step 1: Run full test suite one last time**

```bash
pytest tests/ -v
```

Expected: all tests PASS.

**Step 2: Final commit**

```bash
git add .
git commit -m "feat: complete Voest Paquets app - Gradio UI with SQLite and OpenRouter"
```

---

## Known Limitations / Developer Notes

1. **pdf2image requires Poppler** — on Windows, download Poppler binaries from https://github.com/oschwartz10612/poppler-windows/releases and add to PATH, or pass `poppler_path=` to `convert_from_bytes`.

2. **Gradio Sidebar** — requires Gradio 5.x. If using Gradio 4.x, replace `gr.Sidebar()` with a two-column layout: `with gr.Row(): with gr.Column(scale=1): # sidebar content; with gr.Column(scale=3): # main content`.

3. **LLM model** — the default model `google/gemini-2.0-flash-001` supports vision. Cheaper alternatives: `google/gemini-flash-1.5-8b`, `meta-llama/llama-4-scout:free`. The model must support image input.

4. **Excel output path** — currently saved to project root as `{project_name}.xlsx`. Can be changed in `exporter.py` or `app.py` to save to `data/` folder.
