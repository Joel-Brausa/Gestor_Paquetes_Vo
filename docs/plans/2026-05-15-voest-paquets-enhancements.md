# Voest Paquets — Enhancements Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development to implement this plan task-by-task.

**Goal:** Add model selector with multiple models, fix database structure for global package numbering, simplify data storage, update UI to show package lines, and fix export functionality.

**Architecture:** 
- Database: Add `paquete_num` (global counter per project) and rename existing to `paquete_num_of` (OF-local ID)
- Config: Store available models like app_ejemplo.py does
- App: Add model selector UI (dropdown + add new), display package lines in section 3, fix export callback
- Extractor: Update to handle new data structure

**Tech Stack:** SQLite3, Gradio 5.x, pandas, pytest

---

## Task 1: Update config.py for model management

**Files:**
- Modify: `config.py`

**Step 1: Add models constants**

Add after line 12 (DEFAULT_MODEL):

```python
MODELS_FILE = os.path.join(BASE_DIR, "models", "models.json")

def load_models() -> list[str]:
    if os.path.exists(MODELS_FILE):
        try:
            with open(MODELS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return [DEFAULT_MODEL]
    return [DEFAULT_MODEL]

def save_models(models: list[str]) -> None:
    os.makedirs(os.path.dirname(MODELS_FILE), exist_ok=True)
    with open(MODELS_FILE, "w", encoding="utf-8") as f:
        json.dump(models, f, ensure_ascii=False, indent=2)
```

Add `import json` at the top.

Change `DEFAULT_MODEL` to:
```python
DEFAULT_MODEL = "nvidia/nemotron-3-nano-omni-30b-a3b-reasoning:free"
```

**Step 2: Commit**

```bash
git add config.py
git commit -m "feat: add model management functions to config"
```

---

## Task 2: Update database schema for paquete_num

**Files:**
- Modify: `database.py:159-168` (lines table schema)

**Step 1: Update lines table schema**

Replace the CREATE TABLE lines with:

```python
CREATE TABLE IF NOT EXISTS lines (
    id               INTEGER PRIMARY KEY AUTOINCREMENT,
    packing_list_id  INTEGER NOT NULL REFERENCES packing_lists(id),
    paquete_num      INTEGER NOT NULL,
    paquete_num_of   INTEGER,
    kilos_paquete    REAL,
    linea            INTEGER,
    piezas           INTEGER,
    longitud         REAL,
    marca            TEXT
);
```

**Step 2: Add function to get next paquete_num**

Add after `get_project_lines()` function:

```python
def get_next_paquete_num(project_id: int) -> int:
    with _connect() as conn:
        row = conn.execute(
            "SELECT MAX(paquete_num) FROM lines WHERE packing_list_id IN (SELECT id FROM packing_lists WHERE project_id = ?)",
            (project_id,)
        ).fetchone()
        return (row[0] or 0) + 1
```

**Step 3: Update save_packing_list to use new paquete_num**

In `save_packing_list()`, replace the lines insertion loop (lines 234-249) with:

```python
next_paquete_num = get_next_paquete_num(project_id)
counter = 0
for paquete in data.get("paquetes", []):
    for line in paquete.get("lineas", []):
        conn.execute(
            """INSERT INTO lines
               (packing_list_id, paquete_num, paquete_num_of, kilos_paquete, linea, piezas, longitud, marca)
               VALUES (?,?,?,?,?,?,?,?)""",
            (
                pl_id,
                next_paquete_num + counter,
                paquete.get("paquete_num"),
                paquete.get("kilos_paquete"),
                line.get("linea"),
                line.get("piezas"),
                line.get("longitud"),
                line.get("marca"),
            ),
        )
    counter += 1
```

**Step 4: Update packing_lists table (remove unnecessary fields)**

Replace the CREATE TABLE packing_lists (lines 140-157) to remove: cliente, desarrollo, espesor, calidad, fecha_doc:

```python
CREATE TABLE IF NOT EXISTS packing_lists (
    id             INTEGER PRIMARY KEY AUTOINCREMENT,
    project_id     INTEGER NOT NULL REFERENCES projects(id),
    of_number      TEXT    NOT NULL,
    n_pedido       TEXT,
    articulo       TEXT,
    ref_pedido     TEXT,
    nota           TEXT,
    total_piezas   INTEGER,
    kilos_teoricos REAL,
    imported_at    TEXT    NOT NULL,
    UNIQUE(project_id, of_number)
);
```

**Step 5: Update get_project_lines() query**

Replace query (lines 268-281) to match new schema:

```python
rows = conn.execute(
    """SELECT
           p.name          AS proyecto,
           pl.of_number, pl.n_pedido, pl.articulo,
           pl.ref_pedido, pl.nota,
           pl.total_piezas, pl.kilos_teoricos,
           l.paquete_num, l.paquete_num_of, l.kilos_paquete, l.linea, l.piezas, l.longitud, l.marca
       FROM lines l
       JOIN packing_lists pl ON pl.id = l.packing_list_id
       JOIN projects p       ON p.id  = pl.project_id
       WHERE pl.project_id = ?
       ORDER BY pl.of_number, l.paquete_num, l.linea""",
    (project_id,),
).fetchall()
```

**Step 6: Update tests to match new schema**

In `tests/test_database.py`:
- Update `_sample_data()` helper to remove deleted fields
- Update test assertions for the new schema

**Step 7: Run tests**

```bash
pytest tests/test_database.py -v
```

Expected: All 5 tests PASS

**Step 8: Commit**

```bash
git add database.py tests/test_database.py
git commit -m "feat: add global paquete_num and remove unnecessary fields from db schema"
```

---

## Task 3: Update extractor.py EXTRACTION_PROMPT

**Files:**
- Modify: `extractor.py:417-453` (EXTRACTION_PROMPT)

**Step 1: Update prompt to remove unnecessary fields**

Replace EXTRACTION_PROMPT content to only include:
- of_number, n_pedido, articulo, ref_pedido, nota
- total_piezas, kilos_teoricos
- paquetes with paquete_num, kilos_paquete, lineas with linea, piezas, longitud, marca

Remove: cliente, desarrollo, espesor, calidad, fecha_doc

**Step 2: Commit**

```bash
git add extractor.py
git commit -m "feat: update extraction prompt to remove unnecessary fields"
```

---

## Task 4: Update app.py sidebar — model selector

**Files:**
- Modify: `app.py:40-65` (sidebar)

**Step 1: Add model management imports**

Add at top:
```python
from config import load_models, save_models
```

**Step 2: Replace sidebar model section**

Replace lines 52-59 with:

```python
        gr.Markdown("---")
        gr.Markdown("**Modelo OpenRouter**")

        models_list = gr.State(load_models())
        
        def _get_model_choices():
            return models_list.value

        model_dropdown = gr.Dropdown(
            label="Modelo activo",
            choices=_get_model_choices(),
            value=_get_model_choices()[0] if _get_model_choices() else config.DEFAULT_MODEL,
            interactive=True,
        )
        
        with gr.Row():
            new_model_name = gr.Textbox(
                label="Nuevo modelo",
                placeholder="proveedor/modelo:tag",
                scale=3,
            )
            add_model_btn = gr.Button("➕ Añadir", scale=1)
        
        model_status = gr.Markdown("")

        @add_model_btn.click(inputs=new_model_name, outputs=[model_dropdown, new_model_name, model_status])
        def add_model(new_model):
            new_model = new_model.strip()
            if not new_model:
                return gr.update(), gr.update(), "Introduce un modelo."
            models = _get_model_choices()
            if new_model not in models:
                models.append(new_model)
                save_models(models)
                models_list.value = models
                return gr.update(choices=models, value=new_model), "", "Modelo añadido."
            else:
                return gr.update(), gr.update(), "Modelo ya existe."
```

Update `model_input` reference in `process_pdf` to use `model_dropdown` instead.

**Step 3: Update process_pdf input**

Change line 115 from:
```python
inputs=[pdf_upload, project_dd, model_input, api_key_input],
```

To:
```python
inputs=[pdf_upload, project_dd, model_dropdown, api_key_input],
```

**Step 4: Remove old model_input**

Delete the old `model_input` textbox definition (lines 55-59 in original).

**Step 5: Commit**

```bash
git add app.py
git commit -m "feat: add model dropdown selector with ability to add new models"
```

---

## Task 5: Update app.py section 3 — show package lines

**Files:**
- Modify: `app.py:146-168` (section 3)

**Step 1: Replace section 3 table and logic**

Replace entire section 3 (lines 144-168) with:

```python
    gr.Markdown("---")

    # ── 3. Líneas de paquetes del proyecto ────────────────────────────
    gr.Markdown("## 3. Líneas de paquetes del proyecto")

    lines_table = gr.Dataframe(label="Líneas de paquetes", interactive=False)
    export_btn = gr.Button("Exportar proyecto a Excel", variant="secondary")
    export_file = gr.File(label="Descarga", visible=False)

    def _lines_table(project_name: str):
        if not project_name:
            return pd.DataFrame()
        pid = database.get_project_id(project_name)
        if pid is None:
            return pd.DataFrame()
        rows = database.get_project_lines(pid)
        if not rows:
            return pd.DataFrame(columns=["OF", "N.Pedido", "Artículo", "Paquete_Num", "Paquete_Num_OF", "Kilos", "Línea", "Piezas", "Longitud", "Marca"])
        df = pd.DataFrame(rows)
        df = df.rename(columns={
            "of_number": "OF",
            "n_pedido": "N.Pedido",
            "articulo": "Artículo",
            "paquete_num": "Paquete_Num",
            "paquete_num_of": "Paquete_Num_OF",
            "kilos_paquete": "Kilos",
            "linea": "Línea",
            "piezas": "Piezas",
            "longitud": "Longitud",
            "marca": "Marca",
        })
        cols = ["OF", "N.Pedido", "Artículo", "Paquete_Num", "Paquete_Num_OF", "Kilos", "Línea", "Piezas", "Longitud", "Marca"]
        return df[[c for c in cols if c in df.columns]]

    @project_dd.change(inputs=project_dd, outputs=lines_table)
    def refresh_lines_table(project_name):
        return _lines_table(project_name)

    @export_btn.click(inputs=project_dd, outputs=export_file)
    def export_excel(project_name):
        if not project_name:
            return gr.update(visible=False)
        pid = database.get_project_id(project_name)
        lines = database.get_project_lines(pid)
        if not lines:
            return gr.update(visible=False)
        excel_bytes = exporter.build_excel(lines)
        
        out_path = os.path.join(config.BASE_DIR, f"{project_name}.xlsx")
        with open(out_path, "wb") as f:
            f.write(excel_bytes)
        return gr.update(visible=True, value=out_path)
```

**Step 2: Update process_pdf outputs**

Change line 116 from:
```python
outputs=[import_status, preview_table, pl_table],
```

To:
```python
outputs=[import_status, preview_table, lines_table],
```

And line 142 change:
```python
return msg, gr.update(visible=True, value=df_preview), _pl_table(project_name)
```

To:
```python
return msg, gr.update(visible=True, value=df_preview), _lines_table(project_name)
```

**Step 3: Remove old helpers and tables**

Delete:
- `_pl_table()` function
- Old `pl_table` dataframe definition
- Old `refresh_table()` callback

**Step 4: Update create_project outputs**

Change line 91 from:
```python
outputs=[project_dd, new_project_name, project_status, pl_table],
```

To:
```python
outputs=[project_dd, new_project_name, project_status, lines_table],
```

And line 100:
```python
return gr.update(choices=choices, value=name), "", f"Proyecto **{name}** creado.", _pl_table(name)
```

To:
```python
return gr.update(choices=choices, value=name), "", f"Proyecto **{name}** creado.", _lines_table(name)
```

**Step 5: Commit**

```bash
git add app.py
git commit -m "feat: update section 3 to show package lines instead of packing lists"
```

---

## Task 6: Test everything

**Files:**
- Test: `tests/test_database.py`, `tests/test_exporter.py`

**Step 1: Run all tests**

```bash
pytest tests/ -v
```

Expected: All 15 tests PASS

**Step 2: Test app manually**

Run: `python app.py`
- Create project
- Import packing list
- Verify section 3 shows package lines
- Verify model selector works
- Test export to Excel (should not hang)

**Step 3: Commit if all pass**

```bash
git add .
git commit -m "test: verify all changes work correctly"
```

---

## Known Issues to Watch

1. **Worktree models.json**: Must create `models/` directory on first run
2. **Excel export timeout**: If export still hangs, check if `os.path.join()` is creating correct path
3. **Database migration**: Existing databases won't have `paquete_num` column — users need fresh database
