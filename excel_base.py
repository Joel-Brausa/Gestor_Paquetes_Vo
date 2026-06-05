import re
from io import BytesIO
from typing import Optional

from openpyxl import load_workbook

import database


# ── Layout definitions ────────────────────────────────────────────────────────
#
# v1 — layout original del template.
# v2 — variante del cliente: todas las columnas desplazadas +1 a la derecha.
#
# Índices de columna en convenio openpyxl (1-based).

_LAYOUTS = {
    "v1": {
        "col_paquete":  1,          # A  — paquete_code  (PREFIX/NNNN)
        "col_bundle":   2,          # B  — "Bundle"
        "col_marca":    8,          # H  — valor del lookup AM
        "col_piezas":   9,          # I  — piezas
        "col_am":       39,         # AM — columna de referencia de marcas
        "col_formula":  7,          # G  — columna con fórmula (no se toca)
        "data_cols":    (1, 2, 8, 9),
    },
    "v2": {
        "col_paquete":  2,          # B
        "col_bundle":   3,          # C
        "col_marca":    9,          # I
        "col_piezas":   10,         # J
        "col_am":       40,         # AN
        "col_formula":  8,          # H
        "data_cols":    (2, 3, 9, 10),
    },
}

_PREFIX_CELL    = (5, 2)   # B5 — igual en ambas versiones
_DATA_START_ROW = 14
_PAQUETE_RE     = re.compile(r'.+/\d{3,}')


# ── Layout detection ──────────────────────────────────────────────────────────

def _detect_layout(ws) -> dict:
    """
    Auto-detecta el layout del Excel inspeccionando la hoja.

    Estrategia (en orden):
    1. Busca el patrón PREFIX/NNNN en col A (→ v1) o col B (→ v2) en filas 14-30.
    2. Si no hay datos, comprueba si la columna de fórmula del template es G (→ v1)
       o H (→ v2) en esas mismas filas.
    3. Fallback: v1.
    """
    v1 = _LAYOUTS["v1"]
    v2 = _LAYOUTS["v2"]

    scan_end = min(_DATA_START_ROW + 17, ws.max_row + 1)

    for row_num in range(_DATA_START_ROW, scan_end):
        val_v1 = ws.cell(row=row_num, column=v1["col_paquete"]).value
        val_v2 = ws.cell(row=row_num, column=v2["col_paquete"]).value
        if val_v1 and _PAQUETE_RE.match(str(val_v1)):
            return v1
        if val_v2 and _PAQUETE_RE.match(str(val_v2)):
            return v2

    # Sin filas de datos — usar columna de fórmula del template como indicador
    for row_num in range(_DATA_START_ROW, scan_end):
        g_val = ws.cell(row=row_num, column=v1["col_formula"]).value
        h_val = ws.cell(row=row_num, column=v2["col_formula"]).value
        if isinstance(g_val, str) and g_val.startswith("="):
            return v1
        if isinstance(h_val, str) and h_val.startswith("="):
            return v2

    return v1  # fallback conservador


# ── Helpers internos ──────────────────────────────────────────────────────────

def _load_excel_bytes(project_name: str) -> Optional[bytes]:
    """Devuelve los bytes del Excel desde la BD, o None si no existe."""
    pid = database.get_project_id(project_name)
    if pid is None:
        return None
    return database.load_project_excel(pid)


def _extract_marca_for_h(marca_raw) -> str:
    """
    Extrae la clave de búsqueda del campo marca de la BD.
    Devuelve el texto después del último '-'.
    Ejemplo: "P62989-07031" → "07031",  "S355-J0" → "J0",  "S355J2" → "S355J2"
    """
    if not marca_raw:
        return ""
    s = str(marca_raw).strip()
    if "-" in s:
        return s.rsplit("-", 1)[1].strip()
    return s


def _get_existing_lines(ws, layout: dict) -> set:
    """
    Devuelve un set de tuplas (paquete_code, marca, piezas) de las filas de datos (14+).
    Lee solo los valores escritos por nosotros, nunca fórmulas.
    """
    existing = set()
    for row_num in range(_DATA_START_ROW, ws.max_row + 1):
        col_p = ws.cell(row=row_num, column=layout["col_paquete"]).value
        if col_p:
            col_m = ws.cell(row=row_num, column=layout["col_marca"]).value
            col_z = ws.cell(row=row_num, column=layout["col_piezas"]).value
            existing.add((
                str(col_p),
                str(col_m) if col_m is not None else "",
                str(col_z) if col_z is not None else "",
            ))
    return existing


def _clear_data_cols(ws, row_num: int, layout: dict) -> None:
    """Borra solo las columnas de datos para la fila indicada. Las fórmulas no se tocan."""
    for col_idx in layout["data_cols"]:
        ws.cell(row=row_num, column=col_idx).value = None


def _build_am_cache(excel_bytes: bytes, layout: dict) -> list[str]:
    """
    Lee todos los valores no vacíos de la columna AM/AN (según layout) con data_only=True
    para obtener los valores cacheados de las fórmulas.
    """
    col_am = layout["col_am"]
    wb = load_workbook(BytesIO(excel_bytes), data_only=True)
    ws = wb.active
    values = []
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=col_am, max_col=col_am):
        v = row[0].value
        if v is not None:
            values.append(str(v).strip())
    wb.close()
    return values


def _lookup_marca_in_am(marca_raw: str, am_values: list[str]) -> str:
    """
    Encuentra el valor coincidente en la columna AM/AN para una marca de la BD.

    Estrategia (en orden):
      1. El valor AM termina con la clave de marca (e.g. endswith "07031").
      2. El valor AM contiene la clave como subcadena (fallback más laxo).
      3. Fallback: devuelve la clave cruda.
    """
    if not marca_raw:
        return ""
    key = str(marca_raw).strip()

    for v in am_values:
        if v.endswith(key):
            return v
    for v in am_values:
        if key in v:
            return v
    return key


# ── Public API ────────────────────────────────────────────────────────────────

def excel_exists(project_name: str) -> bool:
    pid = database.get_project_id(project_name)
    if pid is None:
        return False
    return database.excel_exists_in_db(pid)


def delete_project_excel(project_name: str) -> bool:
    pid = database.get_project_id(project_name)
    if pid is None:
        return False
    return database.delete_project_excel_from_db(pid)


def read_excel_preview(project_name: str) -> list[dict]:
    excel_bytes = _load_excel_bytes(project_name)
    if not excel_bytes:
        return []
    try:
        wb = load_workbook(BytesIO(excel_bytes))
        ws = wb.active
        layout = _detect_layout(ws)
        col_anchor = layout["col_paquete"]  # columna que indica si la fila tiene datos
        # Leemos hasta la columna J (10) para cubrir ambos layouts
        read_cols = 10
        result = []
        for row_num in range(12, ws.max_row + 1):
            row = ws[row_num]
            if row_num <= 13:
                values = tuple(cell.value for cell in row[0:read_cols])
                result.append({"row": row_num, "values": values})
            elif ws.cell(row=row_num, column=col_anchor).value:
                values = tuple(cell.value for cell in row[0:read_cols])
                result.append({"row": row_num, "values": values})
        wb.close()
        return result
    except Exception as e:
        raise ValueError(f"Error al leer Excel: {str(e)}")


def find_marca_in_column_am(project_name: str, marca: str) -> str:
    """Busca el valor de marca en la columna AM/AN y devuelve el valor completo coincidente."""
    excel_bytes = _load_excel_bytes(project_name)
    if not excel_bytes:
        return ""
    try:
        wb_detect = load_workbook(BytesIO(excel_bytes))
        layout = _detect_layout(wb_detect.active)
        wb_detect.close()

        wb = load_workbook(BytesIO(excel_bytes), data_only=True)
        ws = wb.active
        col_am = layout["col_am"]
        marca_str = str(marca).strip()
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=col_am, max_col=col_am):
            cell_value = row[0].value
            if cell_value and marca_str in str(cell_value):
                wb.close()
                return str(cell_value)
        wb.close()
        return ""
    except Exception as e:
        raise ValueError(f"Error al buscar marca en AM/AN: {str(e)}")


def write_lines_to_excel(project_name: str, lines: list[dict]) -> dict:
    """
    Escribe las líneas del proyecto en el Excel a partir de la fila 14.
    Cada línea de la BD → una fila del Excel.

    Columnas escritas por fila (según layout auto-detectado):
      v1:  A — paquete_code  |  B — "Bundle"  |  H — marca (AM lookup)  |  I — piezas
      v2:  B — paquete_code  |  C — "Bundle"  |  I — marca (AN lookup)  |  J — piezas

    La columna de fórmula (G en v1, H en v2) nunca se toca.

    Clave de duplicados: (paquete_code, marca, piezas) — estable entre sincronizaciones.

    Guarda el Excel actualizado en la BD.
    Devuelve dict: {added, duplicates, errors}.
    """
    pid = database.get_project_id(project_name)
    if pid is None:
        raise ValueError("Proyecto no encontrado.")
    excel_bytes = database.load_project_excel(pid)
    if not excel_bytes:
        raise ValueError("Sube un Excel base primero")

    # Detectar layout con el workbook de escritura (preserva fórmulas)
    wb = load_workbook(BytesIO(excel_bytes))
    ws = wb.active
    layout = _detect_layout(ws)

    # Caché del lookup AM/AN (usa data_only=True para leer valores cacheados)
    am_values = _build_am_cache(excel_bytes, layout)

    # Leer prefijo desde B5 (igual en v1 y v2)
    b5_value = ws.cell(row=_PREFIX_CELL[0], column=_PREFIX_CELL[1]).value
    col_a_prefix = str(b5_value).strip() if b5_value else project_name

    existing_lines = _get_existing_lines(ws, layout)

    added = 0
    duplicates = 0
    errors = []

    # Primera fila vacía en la columna de paquete, a partir de la fila 14
    current_row = _DATA_START_ROW
    for row_num in range(_DATA_START_ROW, ws.max_row + 1):
        if ws.cell(row=row_num, column=layout["col_paquete"]).value is None:
            current_row = row_num
            break
    else:
        current_row = ws.max_row + 1

    for line in lines:
        try:
            paquete_num = line.get("paquete_num")
            paquete_num_str = (
                f"{paquete_num:04d}" if isinstance(paquete_num, int)
                else str(paquete_num).zfill(4)
            )
            paquete_code = f"{col_a_prefix}/{paquete_num_str}"

            marca_h   = _lookup_marca_in_am(line.get("marca", ""), am_values)
            piezas    = line.get("piezas", "")
            piezas_str = str(piezas) if piezas else ""

            line_tuple = (paquete_code, marca_h, piezas_str)
            if line_tuple in existing_lines:
                duplicates += 1
                continue

            ws.cell(row=current_row, column=layout["col_paquete"]).value = paquete_code
            ws.cell(row=current_row, column=layout["col_bundle"]).value  = "Bundle"
            ws.cell(row=current_row, column=layout["col_marca"]).value   = marca_h
            ws.cell(row=current_row, column=layout["col_piezas"]).value  = piezas if piezas else ""

            added += 1
            current_row += 1
            existing_lines.add(line_tuple)

        except Exception as e:
            errors.append(f"Fila {current_row}: {str(e)}")

    buf = BytesIO()
    wb.save(buf)
    wb.close()
    database.save_project_excel(pid, buf.getvalue())

    return {"added": added, "duplicates": duplicates, "errors": errors}


def count_n_pedido_rows_in_excel(project_name: str, n_pedido: str) -> int:
    """Cuenta las filas del Excel que pertenecen a un N.Pedido específico."""
    if not excel_exists(project_name):
        return 0
    pid = database.get_project_id(project_name)
    if pid is None:
        return 0

    lines_db = database.get_paquete_nums_for_n_pedido(pid, n_pedido)
    if not lines_db:
        return 0

    excel_bytes = database.load_project_excel(pid)
    if not excel_bytes:
        return 0

    paquete_codes = set()
    for l in lines_db:
        pnum = l["paquete_num"]
        pnum_str = f"{pnum:04d}" if isinstance(pnum, int) else str(pnum).zfill(4)
        paquete_codes.add(f"{project_name}/{pnum_str}")

    try:
        wb = load_workbook(BytesIO(excel_bytes))
        ws = wb.active
        layout = _detect_layout(ws)
        col_p = layout["col_paquete"]
        count = sum(
            1 for row_num in range(_DATA_START_ROW, ws.max_row + 1)
            if ws.cell(row=row_num, column=col_p).value
            and str(ws.cell(row=row_num, column=col_p).value) in paquete_codes
        )
        wb.close()
        return count
    except Exception:
        return 0


def delete_all_data_rows_from_excel(project_name: str) -> int:
    """
    Borra las columnas de datos de todas las filas desde la fila 14.
    Preserva las fórmulas. Guarda el Excel en la BD.
    Devuelve el número de filas borradas.
    """
    pid = database.get_project_id(project_name)
    if pid is None:
        return 0
    excel_bytes = database.load_project_excel(pid)
    if not excel_bytes:
        return 0
    try:
        wb = load_workbook(BytesIO(excel_bytes))
        ws = wb.active
        layout = _detect_layout(ws)
        col_p = layout["col_paquete"]
        cleared = 0
        for row_num in range(_DATA_START_ROW, ws.max_row + 1):
            if ws.cell(row=row_num, column=col_p).value is not None:
                _clear_data_cols(ws, row_num, layout)
                cleared += 1
        buf = BytesIO()
        wb.save(buf)
        wb.close()
        database.save_project_excel(pid, buf.getvalue())
        return cleared
    except Exception as e:
        raise ValueError(f"Error al borrar todas las filas: {str(e)}")


def delete_rows_from_excel(project_name: str, row_numbers: list[int]) -> int:
    """
    Borra las columnas de datos para las filas indicadas (índice 1-based).
    Solo se procesan filas >= 14 (las cabeceras 12-13 están protegidas).
    Preserva las fórmulas. Guarda el Excel en la BD.
    Devuelve el número de filas borradas.
    """
    safe_rows = [r for r in row_numbers if r >= _DATA_START_ROW]
    if not safe_rows:
        return 0
    pid = database.get_project_id(project_name)
    if pid is None:
        return 0
    excel_bytes = database.load_project_excel(pid)
    if not excel_bytes:
        return 0
    try:
        wb = load_workbook(BytesIO(excel_bytes))
        ws = wb.active
        layout = _detect_layout(ws)
        for row_num in safe_rows:
            _clear_data_cols(ws, row_num, layout)
        buf = BytesIO()
        wb.save(buf)
        wb.close()
        database.save_project_excel(pid, buf.getvalue())
        return len(safe_rows)
    except Exception as e:
        raise ValueError(f"Error al borrar filas del Excel: {str(e)}")


def delete_n_pedido_from_excel(project_name: str, n_pedido: str) -> int:
    """
    Borra las columnas de datos de todas las filas que pertenecen a un N.Pedido.
    Preserva las fórmulas. Guarda el Excel en la BD.
    Devuelve el número de filas borradas.
    """
    if not excel_exists(project_name):
        return 0
    pid = database.get_project_id(project_name)
    if pid is None:
        return 0

    lines_db = database.get_paquete_nums_for_n_pedido(pid, n_pedido)
    if not lines_db:
        return 0

    excel_bytes = database.load_project_excel(pid)
    if not excel_bytes:
        return 0

    paquete_codes = set()
    for l in lines_db:
        pnum = l["paquete_num"]
        pnum_str = f"{pnum:04d}" if isinstance(pnum, int) else str(pnum).zfill(4)
        paquete_codes.add(f"{project_name}/{pnum_str}")

    try:
        wb = load_workbook(BytesIO(excel_bytes))
        ws = wb.active
        layout = _detect_layout(ws)
        col_p = layout["col_paquete"]
        cleared = 0
        for row_num in range(_DATA_START_ROW, ws.max_row + 1):
            col_val = ws.cell(row=row_num, column=col_p).value
            if col_val and str(col_val) in paquete_codes:
                _clear_data_cols(ws, row_num, layout)
                cleared += 1
        buf = BytesIO()
        wb.save(buf)
        wb.close()
        database.save_project_excel(pid, buf.getvalue())
        return cleared
    except Exception as e:
        raise ValueError(f"Error al borrar N.Pedido del Excel: {str(e)}")
