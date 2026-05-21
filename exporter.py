import io

import pandas as pd


EXPORT_COLUMNS = [
    ("n_pedido", "N.Pedido"),
    ("paquete_num", "Paquete_Num"),
    ("of_number", "OF"),
    ("paquete_num_of", "Paquete_Num_OF"),
    ("linea", "Línea"),
    ("marca", "Marca"),
    ("piezas", "Piezas"),
    ("articulo", "Artículo"),
]


def build_excel(lines: list[dict]) -> bytes:
    """
    Build an Excel file from a flat list of line dicts.
    Only exports specified columns in specified order.
    Returns bytes ready for download.
    """
    if not lines:
        df = pd.DataFrame(columns=[col[1] for col in EXPORT_COLUMNS])
    else:
        df = pd.DataFrame(lines)

        renamed_df = {}
        for db_col, excel_col in EXPORT_COLUMNS:
            if db_col in df.columns:
                renamed_df[excel_col] = df[db_col]

        df = pd.DataFrame(renamed_df)
        df = df[[col[1] for col in EXPORT_COLUMNS if col[1] in df.columns]]

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Packing Lists")
    return buf.getvalue()
