# Excel Base Feature Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Add Excel base upload, preview, refresh, and download functionality to Voest Paquets app with duplicate validation and formula preservation.

**Architecture:** New `excel_base.py` module handles all Excel operations using Openpyxl. Streamlit UI section added at end of `app.py`. Excel files stored in `Proyectos/` folder per project. Duplicate validation by comparing Paquete_Num in column A before writing.

**Tech Stack:** Openpyxl (Excel manipulation), Streamlit (UI), SQLite (data source), Pandas (preview)

---

### Task 1: Modify config.py to add Proyectos folder path

**Files:**
- Modify: `config.py`

**Step 1: Add PROYECTOS_DIR constant**

Open `config.py` and add at the end (before or after existing constants):

```python
# ── Excel Base Projects ────────────────────────────────────────────────────────

PROYECTOS_DIR = os.path.join(BASE_DIR, "Proyectos")
```

**Step 2: Verify config.py syntax**

Run: `python -c "import config; print(config.PROYECTOS_DIR)"`

Expected: Should print path like `C:\Users\jlopez\Claude-code\Voest_Paquets\Proyectos`

**Step 3: Commit**

```bash
git add config.py
git commit -m "config: add PROYECTOS_DIR for Excel base files"
```

---

### Task 2: Create excel_base.py module with core functions

**Files:**
- Create: `excel_base.py`

**Step 1: Create file with imports and helper functions**

Create new file `excel_base.py` with:

```python
import os
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import config


def get_project_excel_path(project_name: str) -> str:
    """Return the path for a project's Excel base file."""
    os.makedirs(config.PROYECTOS_DIR, exist_ok=True)
    return os.path.join(config.PROYECTOS_DIR, f"{project_name}.xlsx")


def excel_exists(project_name: str) -> bool:
    """Check if Excel base file exists for the project."""
    return os.path.exists(get_project_excel_path(project_name))


def read_excel_preview(project_name: str, rows: int = 10) -> list[dict]:
    """Read first N rows from Excel and return as list of dicts."""
    if not excel_exists(project_name):
        return []
    
    try:
        path = get_project_excel_path(project_name)
        wb = load_workbook(path)
        ws = wb.active
        
        result = []
        for i, row in enumerate(ws.iter_rows(min_row=1, max_row=rows, values_only=True), 1):
            result.append({"row": i, "values": row})
        
        wb.close()
        return result
    except Exception as e:
        raise ValueError(f"Error al leer Excel: {str(e)}")


def get_existing_paquete_nums(project_name: str) -> set[str]:
    """Read column A from Excel and extract all existing Paquete_Num codes."""
    if not excel_exists(project_name):
        return set()
    
    try:
        path = get_project_excel_path(project_name)
        wb = load_workbook(path, data_only=False)
        ws = wb.active
        
        existing = set()
        # Start from row 14 where data begins
        for row in ws.iter_rows(min_row=14, max_row=ws.max_row, min_col=1, max_col=1):
            cell_value = row[0].value
            if cell_value:
                existing.add(str(cell_value))
        
        wb.close()
        return existing
    except Exception as e:
        raise ValueError(f"Error al validar duplicados: {str(e)}")


def find_marca_in_column_am(project_name: str, marca: str) -> str:
    """Search for marca value in column AM and return full matching value."""
    if not excel_exists(project_name):
        return ""
    
    try:
        path = get_project_excel_path(project_name)
        wb = load_workbook(path, data_only=True)
        ws = wb.active
        
        # Column AM is column 39 (A=1, Z=26, AA=27, AM=39)
        col_am = 39
        
        marca_str = str(marca).strip()
        
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=col_am, max_col=col_am):
            cell_value = row[0].value
            if cell_value and marca_str in str(cell_value):
                wb.close()
                return str(cell_value)
        
        wb.close()
        return ""
    except Exception as e:
        raise ValueError(f"Error al buscar marca en AM: {str(e)}")


def write_lines_to_excel(project_name: str, lines: list[dict]) -> dict:
    """
    Write project lines to Excel starting at row 14.
    Validates duplicates and preserves existing data.
    Returns dict with stats: {added: int, duplicates: int, errors: list}
    """
    if not excel_exists(project_name):
        raise ValueError("Sube un Excel base primero")
    
    try:
        path = get_project_excel_path(project_name)
        wb = load_workbook(path)
        ws = wb.active
        
        # Get existing Paquete_Num values to check for duplicates
        existing_paquete_nums = get_existing_paquete_nums(project_name)
        
        added = 0
        duplicates = 0
        errors = []
        
        # Find first empty row starting from row 14
        current_row = 14
        # First, find the last filled row
        for row_num in range(14, ws.max_row + 1):
            if ws[f'A{row_num}'].value is None:
                current_row = row_num
                break
        else:
            # If no empty row found, use max_row + 1
            current_row = ws.max_row + 1
        
        for line in lines:
            try:
                paquete_num = line.get("paquete_num")
                paquete_num_str = f"{paquete_num:04d}" if isinstance(paquete_num, int) else str(paquete_num).zfill(4)
                paquete_code = f"{project_name}/{paquete_num_str}"
                
                # Check for duplicate
                if paquete_code in existing_paquete_nums:
                    duplicates += 1
                    continue
                
                # Column A: Project/Paquete_Num
                ws[f'A{current_row}'] = paquete_code
                
                # Column B: "Bundle"
                ws[f'B{current_row}'] = "Bundle"
                
                # Columns C, D, E: Empty (skip)
                
                # Column F: Formula (adjusts row number)
                ws[f'F{current_row}'] = f"=REDONDEAR.MAS(G{current_row};0)"
                
                # Column H: VLOOKUP formula for Marca search in column AM
                marca = line.get("marca", "")
                marca_value = find_marca_in_column_am(project_name, marca)
                ws[f'H{current_row}'] = marca_value if marca_value else ""
                
                added += 1
                current_row += 1
                existing_paquete_nums.add(paquete_code)
                
            except Exception as e:
                errors.append(f"Fila {current_row}: {str(e)}")
        
        wb.save(path)
        wb.close()
        
        return {
            "added": added,
            "duplicates": duplicates,
            "errors": errors
        }
    
    except ValueError as e:
        raise e
    except Exception as e:
        raise ValueError(f"Error al escribir Excel: {str(e)}")
```

**Step 2: Verify module loads**

Run: `python -c "import excel_base; print('Module loaded successfully')"`

Expected: Should print "Module loaded successfully"

**Step 3: Commit**

```bash
git add excel_base.py
git commit -m "feat: create excel_base module with core functions"
```

---

### Task 3: Modify app.py to add Excel base section UI

**Files:**
- Modify: `app.py:260-304` (after line 260, before existing Section 3)

**Step 1: Add new section 4 at the very end of app.py**

Open `app.py` and add this code at the END of the file (after the current Section 3 code that ends around line 304):

```python
st.markdown("---")

# ── Section 4: Excel Base ───────────────────────────────────────────────────────

st.markdown("## 4. Excel base del proyecto")

if not st.session_state.selected_project:
    st.info("Selecciona un proyecto primero para gestionar su Excel base.")
else:
    import excel_base
    
    col1, col2, col3 = st.columns([2, 1, 1])
    
    # ── File Uploader ──────────────────────────────────────────────────────
    
    with col1:
        uploaded_excel = st.file_uploader(
            "Subir Excel base",
            type="xlsx",
            key="excel_base_upload",
        )
        
        if uploaded_excel:
            try:
                excel_path = excel_base.get_project_excel_path(st.session_state.selected_project)
                with open(excel_path, "wb") as f:
                    f.write(uploaded_excel.getbuffer())
                st.success(f"✅ Excel base cargado: {uploaded_excel.name}")
            except Exception as e:
                st.error(f"❌ Error al cargar Excel: {str(e)}")
    
    # ── Refresh Button ──────────────────────────────────────────────────────
    
    with col2:
        if st.button("🔄 Refrescar", key="excel_refresh_btn"):
            if not excel_base.excel_exists(st.session_state.selected_project):
                st.error("Sube un Excel base primero.")
            else:
                try:
                    pid = database.get_project_id(st.session_state.selected_project)
                    lines = database.get_project_lines(pid)
                    
                    if not lines:
                        st.warning("El proyecto no tiene líneas para exportar.")
                    else:
                        result = excel_base.write_lines_to_excel(
                            st.session_state.selected_project,
                            lines
                        )
                        
                        st.success(
                            f"✅ {result['added']} filas añadidas, "
                            f"{result['duplicates']} duplicados ignorados"
                        )
                        
                        if result['errors']:
                            for error in result['errors']:
                                st.warning(f"⚠️ {error}")
                
                except Exception as e:
                    st.error(f"❌ Error: {str(e)}")
    
    # ── Download Button ────────────────────────────────────────────────────
    
    with col3:
        if st.button("⬇️ Descargar", key="excel_download_btn"):
            if not excel_base.excel_exists(st.session_state.selected_project):
                st.error("No hay Excel base para descargar.")
            else:
                try:
                    excel_path = excel_base.get_project_excel_path(st.session_state.selected_project)
                    with open(excel_path, "rb") as f:
                        st.download_button(
                            label="Descargar Excel",
                            data=f.read(),
                            file_name=f"{st.session_state.selected_project}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="excel_download_file_btn",
                        )
                except Exception as e:
                    st.error(f"❌ Error: {str(e)}")
    
    # ── Preview Table ──────────────────────────────────────────────────────
    
    st.markdown("### Vista previa del Excel base")
    
    if excel_base.excel_exists(st.session_state.selected_project):
        try:
            preview_data = excel_base.read_excel_preview(st.session_state.selected_project, rows=15)
            
            if preview_data:
                # Convert to DataFrame for display
                preview_df = pd.DataFrame([
                    {"Fila": item["row"], **{f"Col{i+1}": val for i, val in enumerate(item["values"])}}
                    for item in preview_data
                ])
                st.dataframe(preview_df, use_container_width=True, hide_index=True)
            else:
                st.info("Excel base vacío.")
        except Exception as e:
            st.error(f"❌ Error al mostrar vista previa: {str(e)}")
    else:
        st.info("Sube un Excel base para ver la vista previa.")
```

**Step 2: Verify app.py syntax**

Run: `python -m py_compile app.py`

Expected: No output (syntax is valid)

**Step 3: Test the section loads**

Run Streamlit and verify the new section appears at the bottom:
```bash
streamlit run app.py --server.port=7860
```

Expected: New "4. Excel base del proyecto" section appears with upload, refresh, download buttons and preview area.

**Step 4: Commit**

```bash
git add app.py
git commit -m "feat: add Excel base section to app.py with upload, refresh, download, and preview"
```

---

### Task 4: Verify folder structure exists

**Files:**
- Check: `Proyectos/` folder

**Step 1: Ensure Proyectos folder exists**

Run: 
```bash
python -c "import config, os; os.makedirs(config.PROYECTOS_DIR, exist_ok=True); print(f'Proyectos folder ready: {config.PROYECTOS_DIR}')"
```

Expected: Should print path and confirm folder is ready

**Step 2: No commit needed** (folder creation is handled in code)

---

### Task 5: Manual integration test

**Files:**
- Test: All files together

**Step 1: Start the app**

```bash
streamlit run app.py --server.port=7860
```

**Step 2: Test workflow**

1. ✅ Create a test project if none exists
2. ✅ In Section 4, upload a test Excel file (must be valid .xlsx)
3. ✅ Verify "Excel base cargado" message appears
4. ✅ Import at least one PDF to populate database
5. ✅ Click "🔄 Refrescar" button
6. ✅ Verify success message with count of rows added
7. ✅ Click "⬇️ Descargar" to download the updated Excel
8. ✅ Open downloaded Excel and verify:
   - Data starts at row 14
   - Column A has format `ProjectName/0001`
   - Column B has "Bundle"
   - Column F has formula `=REDONDEAR.MAS(G14;0)`
   - Column H has Marca lookup results

**Step 3: Verify duplicate handling**

1. ✅ Click "🔄 Refrescar" again
2. ✅ Verify message shows "X duplicados ignorados" (should be the same count as before)
3. ✅ Download Excel again and verify no duplicate rows were added

**Step 4: Test error cases**

1. ✅ Try to refresh without uploading Excel → Should show "Sube un Excel base primero"
2. ✅ Upload a corrupted/invalid xlsx file → Should show error message
3. ✅ Create new project and try to download before uploading Excel → Should show "No hay Excel base"

**Step 5: Commit test results**

```bash
git add .
git commit -m "test: manual integration testing of Excel base feature complete"
```

---

## Success Criteria

✅ New Section 4 appears at end of app.py
✅ File upload accepts .xlsx and saves to `Proyectos/{ProjectName}.xlsx`
✅ Preview shows first 15 rows of uploaded Excel
✅ Refresh button reads project data and writes to Excel
✅ Data written starting at row 14 with correct format
✅ Column A: `{Project}/{Paquete_Num:04d}` format
✅ Column B: "Bundle" literal string
✅ Column F: Formula `=REDONDEAR.MAS(G{row};0)` preserved
✅ Column H: Marca lookup from column AM
✅ Duplicate validation prevents same Paquete_Num being added twice
✅ Download button returns valid Excel file
✅ All error messages appear correctly in Spanish
