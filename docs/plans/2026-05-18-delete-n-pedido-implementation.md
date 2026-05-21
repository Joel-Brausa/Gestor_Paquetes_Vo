# Delete N.Pedido Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Add UI and backend functionality to delete all lines of a N.Pedido from both database and Excel file with two-step confirmation workflow.

**Architecture:** New helper functions in database.py and excel_base.py handle deletion logic. Streamlit UI section added to Section 3 with dropdown selector and confirmation dialog. Deletion cascades to both database and Excel simultaneously.

**Tech Stack:** Streamlit (UI), SQLite (database deletion), Openpyxl (Excel row deletion)

---

### Task 1: Add database helper function to delete N.Pedido

**Files:**
- Modify: `database.py`
- Test: `tests/test_database.py` (if exists, else create)

**Step 1: Review current database.py structure**

Run: `head -50 database.py`

Expected: See imports and existing functions like `get_project_id`, `get_project_lines`, `save_packing_list`, etc.

**Step 2: Write the helper function**

Add this function to the end of `database.py`:

```python
def delete_n_pedido_lines(project_id: int, n_pedido: str) -> int:
    """
    Delete all lines of a N.Pedido from the database.
    Returns count of deleted lines.
    """
    try:
        conn = sqlite3.connect(DATABASE)
        cursor = conn.cursor()
        
        # Get count before deletion
        cursor.execute(
            "SELECT COUNT(*) FROM project_lines WHERE project_id = ? AND n_pedido = ?",
            (project_id, n_pedido)
        )
        count = cursor.fetchone()[0]
        
        # Delete lines
        cursor.execute(
            "DELETE FROM project_lines WHERE project_id = ? AND n_pedido = ?",
            (project_id, n_pedido)
        )
        conn.commit()
        conn.close()
        
        return count
    except Exception as e:
        raise ValueError(f"Error al eliminar N.Pedido: {str(e)}")


def get_unique_n_pedidos(project_id: int) -> list[str]:
    """
    Get all unique N.Pedido values for a project.
    Returns sorted list of N.Pedido strings.
    """
    try:
        conn = sqlite3.connect(DATABASE)
        cursor = conn.cursor()
        cursor.execute(
            "SELECT DISTINCT n_pedido FROM project_lines WHERE project_id = ? ORDER BY n_pedido",
            (project_id,)
        )
        result = [row[0] for row in cursor.fetchall()]
        conn.close()
        return result
    except Exception as e:
        raise ValueError(f"Error al obtener N.Pedidos: {str(e)}")
```

**Step 3: Verify syntax**

Run: `python -m py_compile database.py`

Expected: No output (syntax is valid)

**Step 4: Commit**

```bash
git add database.py
git commit -m "feat: add delete_n_pedido_lines and get_unique_n_pedidos functions"
```

---

### Task 2: Add Excel helper function to delete N.Pedido lines

**Files:**
- Modify: `excel_base.py`

**Step 1: Add function to find rows by N.Pedido**

Add this function to `excel_base.py`:

```python
def get_n_pedido_from_paquete_code(paquete_code: str, project_name: str) -> str:
    """
    Extract N.Pedido from a Paquete_Num by looking up the corresponding line in database.
    Paquete_Code format: "ProjectName/0001"
    This is a helper to map Excel rows back to database records.
    """
    try:
        # Extract just the number part
        parts = paquete_code.split('/')
        if len(parts) < 2:
            return ""
        
        paquete_num_str = parts[1]
        paquete_num = int(paquete_num_str)
        
        # Query database for this paquete_num to get n_pedido
        import database
        conn = sqlite3.connect(database.DATABASE)
        cursor = conn.cursor()
        
        # Get project_id
        cursor.execute("SELECT id FROM projects WHERE name = ?", (project_name,))
        project = cursor.fetchone()
        if not project:
            conn.close()
            return ""
        
        # Get n_pedido for this paquete_num
        cursor.execute(
            "SELECT DISTINCT n_pedido FROM project_lines WHERE project_id = ? AND paquete_num = ? LIMIT 1",
            (project[0], paquete_num)
        )
        result = cursor.fetchone()
        conn.close()
        
        return result[0] if result else ""
    except Exception:
        return ""


def delete_n_pedido_from_excel(project_name: str, n_pedido: str) -> int:
    """
    Delete all rows from Excel that belong to a specific N.Pedido.
    Returns count of deleted rows.
    Rows are deleted completely (shift up), not just cleared.
    """
    if not excel_exists(project_name):
        return 0
    
    try:
        path = get_project_excel_path(project_name)
        wb = load_workbook(path)
        ws = wb.active
        
        # Collect rows to delete (iterate backwards to avoid index issues)
        rows_to_delete = []
        
        for row_num in range(ws.max_row, 13, -1):  # Start from bottom, stop at row 14
            col_a = ws[f'A{row_num}'].value
            
            if col_a:
                # Get n_pedido for this row
                paquete_code = str(col_a)
                row_n_pedido = get_n_pedido_from_paquete_code(paquete_code, project_name)
                
                if row_n_pedido == n_pedido:
                    rows_to_delete.append(row_num)
        
        # Delete rows (already in reverse order from above)
        for row_num in rows_to_delete:
            ws.delete_rows(row_num, 1)
        
        wb.save(path)
        wb.close()
        
        return len(rows_to_delete)
    except Exception as e:
        raise ValueError(f"Error al eliminar N.Pedido del Excel: {str(e)}")
```

**Step 2: Add import at top of excel_base.py**

At the top of the file, after existing imports, add:

```python
import sqlite3
```

**Step 3: Verify syntax**

Run: `python -m py_compile excel_base.py`

Expected: No output (syntax is valid)

**Step 4: Commit**

```bash
git add excel_base.py
git commit -m "feat: add delete_n_pedido_from_excel function"
```

---

### Task 3: Add session state for deletion confirmation dialog

**Files:**
- Modify: `app.py:46-57` (Initialize session state section)

**Step 1: Add deletion dialog state variables**

Find the section "Initialize session state" (around line 46) and add after existing state variables:

```python
if "show_delete_n_pedido_confirm" not in st.session_state:
    st.session_state.show_delete_n_pedido_confirm = False

if "delete_n_pedido_value" not in st.session_state:
    st.session_state.delete_n_pedido_value = None
```

**Step 2: Verify app.py still runs**

Run: `python -m py_compile app.py`

Expected: No output (syntax is valid)

**Step 3: Commit**

```bash
git add app.py
git commit -m "feat: add session state variables for N.Pedido deletion dialog"
```

---

### Task 4: Add N.Pedido deletion UI to Section 3

**Files:**
- Modify: `app.py:270-275` (start of Section 3, after the markdown heading)

**Step 1: Add deletion subsection at start of Section 3**

Find the line `st.markdown("## 3. Líneas de paquetes del proyecto")` and add right after it:

```python
# ── Delete N.Pedido Subsection ────────────────────────────────────────

st.markdown("### Eliminar líneas de un N.Pedido")

if not st.session_state.selected_project:
    st.info("Selecciona un proyecto primero para eliminar N.Pedidos.")
else:
    col1, col2 = st.columns([2, 1])
    
    # Get unique N.Pedidos for this project
    try:
        pid = database.get_project_id(st.session_state.selected_project)
        n_pedidos = database.get_unique_n_pedidos(pid) if pid else []
    except Exception:
        n_pedidos = []
    
    with col1:
        if n_pedidos:
            selected_n_pedido = st.selectbox(
                "Seleccionar N.Pedido a eliminar",
                options=n_pedidos,
                key="delete_n_pedido_select",
            )
        else:
            st.info("Sin N.Pedidos para eliminar.")
            selected_n_pedido = None
    
    with col2:
        if selected_n_pedido and st.button("🗑️ Eliminar", key="delete_n_pedido_btn"):
            st.session_state.show_delete_n_pedido_confirm = True
            st.session_state.delete_n_pedido_value = selected_n_pedido
            st.rerun()

# Show confirmation dialog
if st.session_state.show_delete_n_pedido_confirm and st.session_state.delete_n_pedido_value:
    n_pedido = st.session_state.delete_n_pedido_value
    pid = database.get_project_id(st.session_state.selected_project)
    
    # Count lines to delete
    try:
        lines = database.get_project_lines(pid)
        count_db = len([l for l in lines if l["n_pedido"] == n_pedido])
        
        # Count in Excel
        import excel_base
        count_excel = 0
        if excel_base.excel_exists(st.session_state.selected_project):
            try:
                # Count rows in Excel for this n_pedido
                excel_path = excel_base.get_project_excel_path(st.session_state.selected_project)
                from openpyxl import load_workbook
                wb = load_workbook(excel_path)
                ws = wb.active
                
                for row_num in range(14, ws.max_row + 1):
                    col_a = ws[f'A{row_num}'].value
                    if col_a:
                        paquete_code = str(col_a)
                        row_n_pedido = excel_base.get_n_pedido_from_paquete_code(paquete_code, st.session_state.selected_project)
                        if row_n_pedido == n_pedido:
                            count_excel += 1
                wb.close()
            except Exception:
                count_excel = 0
        
        st.warning(f"⚠️ **Eliminar permanentemente N.Pedido {n_pedido} y todas sus líneas?**\n\nSe eliminarán:\n- {count_db} líneas de la base de datos\n- {count_excel} líneas del Excel" + (" (si existe)" if count_excel == 0 else ""))
        
        col1, col2 = st.columns(2)
        with col1:
            if st.button("Sí, eliminar", key="confirm_delete_n_pedido_btn"):
                try:
                    # Delete from database first
                    deleted_db = database.delete_n_pedido_lines(pid, n_pedido)
                    
                    # Delete from Excel
                    deleted_excel = 0
                    if excel_base.excel_exists(st.session_state.selected_project):
                        try:
                            deleted_excel = excel_base.delete_n_pedido_from_excel(
                                st.session_state.selected_project,
                                n_pedido
                            )
                        except Exception as excel_error:
                            st.warning(f"⚠️ Eliminado de BD pero error en Excel: {str(excel_error)}")
                    
                    st.session_state.show_delete_n_pedido_confirm = False
                    st.success(f"N.Pedido {n_pedido} eliminado: {deleted_db} líneas de BD, {deleted_excel} líneas de Excel")
                    st.rerun()
                except Exception as e:
                    st.error(f"Error al eliminar: {e}")
                    st.session_state.show_delete_n_pedido_confirm = False
        
        with col2:
            if st.button("Cancelar", key="cancel_delete_n_pedido_btn"):
                st.session_state.show_delete_n_pedido_confirm = False
                st.rerun()
    except Exception as e:
        st.error(f"Error al preparar eliminación: {e}")
        st.session_state.show_delete_n_pedido_confirm = False

st.markdown("---")
```

**Step 2: Verify app.py syntax**

Run: `python -m py_compile app.py`

Expected: No output (syntax is valid)

**Step 3: Commit**

```bash
git add app.py
git commit -m "feat: add N.Pedido deletion UI to Section 3 with confirmation dialog"
```

---

### Task 5: Manual testing of N.Pedido deletion

**Files:**
- Test: All modified files together

**Step 1: Start the Streamlit app**

Run: `streamlit run app.py --server.port=7860`

**Step 2: Test workflow**

1. ✅ Create or select an existing project
2. ✅ Import at least one PDF to create lines
3. ✅ Scroll to Section 3
4. ✅ Verify dropdown shows N.Pedidos from imported data
5. ✅ Select one N.Pedido
6. ✅ Click "🗑️ Eliminar"
7. ✅ Verify warning dialog shows correct counts:
   - Count from database
   - Count from Excel (if file exists)
8. ✅ Click "Sí, eliminar"
9. ✅ Verify success message: "N.Pedido XXXXX eliminado: X líneas de BD, Y líneas de Excel"
10. ✅ Verify table below refreshes and no longer shows those lines
11. ✅ Verify dropdown no longer shows that N.Pedido

**Step 3: Test error cases**

1. ✅ Click "Cancelar" in confirmation dialog → Dialog closes, no deletion
2. ✅ Upload Excel base, refresh to populate it
3. ✅ Delete a N.Pedido → Verify both DB and Excel show correct counts
4. ✅ Download Excel → Open in Excel and verify rows were deleted completely (not just cleared)
5. ✅ Delete from a project with no Excel → Verify message says "0 líneas del Excel"

**Step 4: Commit test results**

```bash
git add .
git commit -m "test: manual testing of N.Pedido deletion feature complete"
```

---

## Success Criteria

✅ New subsection "Eliminar líneas de un N.Pedido" appears in Section 3
✅ Dropdown lists all unique N.Pedidos for current project
✅ Delete button disabled if no project selected or no N.Pedidos exist
✅ Warning dialog shows correct line counts for both DB and Excel
✅ Clicking "Sí, eliminar" deletes from both database and Excel file
✅ Deleted rows in Excel are removed completely (rows shift up)
✅ Success message shows accurate counts: "X líneas de BD, Y líneas de Excel"
✅ Table refreshes automatically after deletion
✅ Dropdown refreshes to exclude deleted N.Pedido
✅ Canceling confirmation dialog closes without deleting anything
✅ Error handling works for database errors (stops before Excel deletion)
✅ Error handling works for Excel errors (shows warning but deletion succeeds)
✅ All Spanish error messages appear correctly
