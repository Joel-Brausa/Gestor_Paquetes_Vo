# Delete N.Pedido Feature Design

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:writing-plans to create the implementation plan.

**Goal:** Allow users to delete all lines of a complete N.Pedido (purchase order) from both database and Excel when the order is modified or imported incorrectly.

**Architecture:** New UI subsection in Section 3 with two-step workflow: dropdown selector for N.Pedido, then confirmation button with warning dialog. Backend operations handle both database deletion and Excel row removal with cascade.

**Tech Stack:** Streamlit (UI), SQLite (database), Openpyxl (Excel row deletion)

---

## User Flow

1. User selects a N.Pedido from dropdown in Section 3
2. User clicks "🗑️ Eliminar N.Pedido" button
3. Warning dialog shows affected row count for both DB and Excel
4. User confirms deletion
5. All lines are removed from:
   - SQLite database: `DELETE FROM project_lines WHERE n_pedido = ?`
   - Excel file: Rows deleted completely (shift rows up)
6. Success message shows counts deleted

---

## UI Section: Delete N.Pedido

**Position:** New subsection at the top of Section 3, before the table display

**Components:**

**Step 1: N.Pedido Selector**
- Dropdown listing all unique N.Pedidos in current project
- Label: "Seleccionar N.Pedido a eliminar"
- If no projects or lines exist: Info message "Sin N.Pedidos para eliminar"

**Step 2: Delete Button + Confirmation**
- Button: "🗑️ Eliminar N.Pedido" (disabled if no selection)
- On click: Warning dialog showing:
  ```
  ⚠️ **Eliminar permanentemente N.Pedido XXXXX y todas sus líneas?**
  Se eliminarán:
  - X líneas de la base de datos
  - X líneas del Excel (si existe)
  ```
- Action buttons: "Sí, eliminar" | "Cancelar"

---

## Technical Implementation Details

### Database Operation

```python
DELETE FROM project_lines WHERE n_pedido = ?
```

Must be executed before Excel operation to ensure we know the count.

### Excel Operation

When deleting from Excel:
1. Read all rows starting from row 14
2. Identify rows where column A (Paquete_Num code) corresponds to the deleted N.Pedido
3. Delete those rows completely using openpyxl (this shifts remaining rows up)
4. Save the workbook

**Note:** Column mapping - need to identify which lines belong to deleted N.Pedido. Since Excel stores Paquete_Num in column A but database has n_pedido, need to match via project_lines table before deletion.

### Deletion Order

1. **First:** Get count of lines to delete from database
2. **Second:** Get count from Excel (if exists) by reading Paquete_Nums that belong to N.Pedido
3. **Third:** Delete from database
4. **Fourth:** Delete from Excel (if file exists)
5. **Display:** Combined results

---

## Error Handling

| Scenario | Response |
|----------|----------|
| No N.Pedidos in project | ℹ️ "Sin N.Pedidos para eliminar" |
| Database deletion fails | ❌ "Error al eliminar de BD: {error}" - STOP, don't delete Excel |
| Excel file doesn't exist | ℹ️ "Proyecto sin Excel base - eliminado solo de BD" |
| Excel deletion fails | ⚠️ "Eliminado de BD pero error en Excel: {error}" - allow to continue |
| Deletion succeeds | ✅ "N.Pedido XXXXX eliminado: X líneas de BD, Y líneas de Excel" |
| User cancels | ℹ️ No message, dialog closes |

---

## Implementation Notes

- **Helper function needed:** `delete_n_pedido_from_project()` in database.py
- **Helper function needed:** `delete_n_pedido_from_excel()` in excel_base.py
- Use session state to manage warning dialog (similar to existing delete project confirmation)
- Query unique N.Pedidos dynamically each time Section 3 renders
- Disable selector/button if no project selected
- Refresh table display after successful deletion

