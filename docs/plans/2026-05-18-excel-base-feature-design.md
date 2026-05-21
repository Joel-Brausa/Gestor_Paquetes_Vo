# Excel Base Feature Design

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:writing-plans to create the implementation plan.

**Goal:** Add a new "Excel base" section to allow users to upload a per-project Excel template and automatically populate it with imported PDF data.

**Architecture:** Users upload an Excel base file (once per project) stored in `Proyectos/` folder. When they click "Refrescar/Recargar", the app reads project data from database, validates for duplicates, and writes to the Excel starting at row 14 while preserving formulas and reference data.

**Tech Stack:** Streamlit UI, Openpyxl for Excel manipulation, SQLite for data, Pandas for data handling

---

## User Flow

1. User uploads an Excel base file → Stored in `Proyectos/{ProjectName}.xlsx`
2. User imports PDFs → Data saved to database as usual
3. User clicks "Refrescar/Recargar" → App validates duplicates and updates Excel
4. User clicks "Descargar" → Downloads the updated Excel file

---

## UI Section: "Excel base"

**Position:** New Section 4, at the end (after "Líneas de paquetes del proyecto")

**Components:**
- File uploader: Accepts `.xlsx`, saves to `Proyectos/{ProjectName}.xlsx`
- "Refrescar/Recargar" button: Reads DB, validates duplicates, writes to Excel
- "Descargar" button: Downloads updated Excel
- Preview table: Shows first 10 rows of loaded Excel for verification

---

## Data Mapping to Excel

**All data starts at row 14 and beyond:**

| Column | Content | Notes |
|--------|---------|-------|
| A | `{ProjectName}/{Paquete_Num:04d}` | Format: "466950/0001" |
| B | `"Bundle"` | Literal string |
| C, D, E | Empty | No data |
| F | `=REDONDEAR.MAS(G14;0)` | Formula (row adjusts: G15, G16...) |
| H | VLOOKUP formula | Search Marca in column AM, return full value |

---

## Duplicate Validation Strategy

**Before adding a row to Excel:**
1. Extract `Paquete_Num` from current line
2. Read column A of existing Excel
3. Check if `{ProjectName}/{Paquete_Num:04d}` already exists
4. If duplicate → skip row, increment counter
5. If new → add row

**Output:** Show summary "X filas añadidas, Y duplicados ignorados"

---

## Error Handling

| Scenario | Response |
|----------|----------|
| No Excel base uploaded | ❌ "Sube un Excel base primero" |
| Corrupted/invalid Excel | ❌ "Excel base inválido o corrupto" |
| No project data | ℹ️ "El proyecto no tiene líneas para exportar" |
| Duplicates found | ℹ️ "5 filas ya existen, ignoradas" |
| Upload successful | ✅ "Excel base cargado: {filename}" |
| Refresh successful | ✅ "X filas añadidas, Y duplicados ignorados" |

---

## Implementation Notes

- Use **Openpyxl** to preserve formulas, formats, and reference data in column AM
- Store Excel files in `Proyectos/` folder at application root
- Path format: `{BASE_DIR}/Proyectos/{ProjectName}.xlsx`
- Validate file exists before refresh operation
- Handle file read/write errors gracefully
