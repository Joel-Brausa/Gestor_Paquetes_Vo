import os
import psycopg2
import psycopg2.extras
from datetime import datetime, timezone
from typing import Optional


def _get_database_url() -> str:
    # Env var first — set locally via .env (python-dotenv) or by Streamlit Cloud from secrets
    url = os.environ.get("DATABASE_URL", "")
    if url:
        return url
    # Fallback: st.secrets (Streamlit Cloud when env var not present)
    try:
        import streamlit as st
        url = st.secrets.get("DATABASE_URL", "")
        if url:
            return url
    except Exception:
        pass
    raise RuntimeError(
        "DATABASE_URL no configurada. "
        "Añádela en .streamlit/secrets.toml o como variable de entorno."
    )


def _connect() -> psycopg2.extensions.connection:
    return psycopg2.connect(_get_database_url())


def init_db() -> None:
    conn = _connect()
    try:
        with conn.cursor() as cur:
            cur.execute("""
                CREATE TABLE IF NOT EXISTS projects (
                    id         SERIAL PRIMARY KEY,
                    name       TEXT UNIQUE NOT NULL,
                    created_at TEXT NOT NULL
                )
            """)
            cur.execute("""
                CREATE TABLE IF NOT EXISTS packing_lists (
                    id             SERIAL PRIMARY KEY,
                    project_id     INTEGER NOT NULL REFERENCES projects(id),
                    of_number      TEXT NOT NULL,
                    n_pedido       TEXT,
                    articulo       TEXT,
                    ref_pedido     TEXT,
                    nota           TEXT,
                    total_piezas   INTEGER,
                    kilos_teoricos REAL,
                    imported_at    TEXT NOT NULL,
                    UNIQUE(project_id, of_number)
                )
            """)
            cur.execute("""
                CREATE TABLE IF NOT EXISTS lines (
                    id               SERIAL PRIMARY KEY,
                    packing_list_id  INTEGER NOT NULL REFERENCES packing_lists(id),
                    paquete_num      INTEGER NOT NULL,
                    paquete_num_of   INTEGER,
                    kilos_paquete    REAL,
                    linea            INTEGER,
                    piezas           INTEGER,
                    longitud         REAL,
                    marca            TEXT
                )
            """)
            cur.execute("""
                CREATE TABLE IF NOT EXISTS project_excel_files (
                    id          SERIAL PRIMARY KEY,
                    project_id  INTEGER NOT NULL REFERENCES projects(id) ON DELETE CASCADE,
                    excel_data  BYTEA NOT NULL,
                    filename    TEXT,
                    updated_at  TEXT NOT NULL,
                    UNIQUE(project_id)
                )
            """)
        conn.commit()
    finally:
        conn.close()


# ── Projects ──────────────────────────────────────────────────────────────────

def create_project(name: str) -> int:
    conn = _connect()
    try:
        with conn.cursor() as cur:
            cur.execute(
                "INSERT INTO projects (name, created_at) VALUES (%s, %s) RETURNING id",
                (name.strip(), datetime.now(timezone.utc).isoformat()),
            )
            row = cur.fetchone()
        conn.commit()
        return row[0]
    finally:
        conn.close()


def get_projects() -> list[dict]:
    conn = _connect()
    try:
        with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
            cur.execute("SELECT id, name FROM projects ORDER BY name")
            return [dict(r) for r in cur.fetchall()]
    finally:
        conn.close()


def get_project_id(name: str) -> Optional[int]:
    conn = _connect()
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT id FROM projects WHERE name = %s", (name,))
            row = cur.fetchone()
            return row[0] if row else None
    finally:
        conn.close()


def delete_project(project_id: int) -> None:
    conn = _connect()
    try:
        with conn.cursor() as cur:
            cur.execute(
                "DELETE FROM lines WHERE packing_list_id IN "
                "(SELECT id FROM packing_lists WHERE project_id = %s)",
                (project_id,),
            )
            cur.execute("DELETE FROM packing_lists WHERE project_id = %s", (project_id,))
            cur.execute("DELETE FROM projects WHERE id = %s", (project_id,))
        conn.commit()
    finally:
        conn.close()


# ── Packing Lists ─────────────────────────────────────────────────────────────

def save_packing_list(project_id: int, data: dict) -> tuple[int, bool]:
    now = datetime.now(timezone.utc).isoformat()
    of_number = data["of_number"]

    conn = _connect()
    try:
        with conn.cursor() as cur:
            cur.execute(
                "SELECT id FROM packing_lists WHERE project_id = %s AND of_number = %s",
                (project_id, of_number),
            )
            existing = cur.fetchone()

            if existing:
                raise ValueError(
                    f"El packing list '{of_number}' ya existe en este proyecto. "
                    "Elimina primero el N.Pedido correspondiente si deseas reimportarlo."
                )

            cur.execute(
                """INSERT INTO packing_lists
                   (project_id, of_number, n_pedido, articulo, ref_pedido,
                    nota, total_piezas, kilos_teoricos, imported_at)
                   VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s)
                   RETURNING id""",
                (
                    project_id, of_number,
                    data.get("n_pedido"), data.get("articulo"),
                    data.get("ref_pedido"), data.get("nota"),
                    data.get("total_piezas"), data.get("kilos_teoricos"), now,
                ),
            )
            pl_id = cur.fetchone()[0]

            # Get next paquete_num within same transaction to avoid race conditions
            cur.execute(
                "SELECT MAX(paquete_num) FROM lines WHERE packing_list_id IN "
                "(SELECT id FROM packing_lists WHERE project_id = %s)",
                (project_id,),
            )
            row = cur.fetchone()
            next_paquete_num = (row[0] or 0) + 1

            lines_to_insert = []
            counter = 0
            for paquete in data.get("paquetes", []):
                for line in paquete.get("lineas", []):
                    lines_to_insert.append((
                        pl_id,
                        next_paquete_num + counter,
                        paquete.get("paquete_num"),
                        paquete.get("kilos_paquete"),
                        line.get("linea"),
                        line.get("piezas"),
                        line.get("longitud"),
                        line.get("marca"),
                    ))
                counter += 1

            if lines_to_insert:
                psycopg2.extras.execute_values(
                    cur,
                    """INSERT INTO lines
                       (packing_list_id, paquete_num, paquete_num_of, kilos_paquete,
                        linea, piezas, longitud, marca)
                       VALUES %s""",
                    lines_to_insert,
                )
        conn.commit()
        return pl_id
    finally:
        conn.close()


def get_next_paquete_num(project_id: int) -> int:
    conn = _connect()
    try:
        with conn.cursor() as cur:
            cur.execute(
                "SELECT MAX(paquete_num) FROM lines WHERE packing_list_id IN "
                "(SELECT id FROM packing_lists WHERE project_id = %s)",
                (project_id,),
            )
            row = cur.fetchone()
            return (row[0] or 0) + 1
    finally:
        conn.close()


def get_project_lines(project_id: int) -> list[dict]:
    conn = _connect()
    try:
        with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
            cur.execute(
                """SELECT
                       p.name          AS proyecto,
                       pl.of_number, pl.n_pedido, pl.articulo,
                       pl.ref_pedido, pl.nota,
                       pl.total_piezas, pl.kilos_teoricos,
                       l.paquete_num, l.paquete_num_of, l.kilos_paquete,
                       l.linea, l.piezas, l.longitud, l.marca
                   FROM lines l
                   JOIN packing_lists pl ON pl.id = l.packing_list_id
                   JOIN projects p       ON p.id  = pl.project_id
                   WHERE pl.project_id = %s
                   ORDER BY pl.of_number, l.paquete_num, l.linea""",
                (project_id,),
            )
            return [dict(r) for r in cur.fetchall()]
    finally:
        conn.close()


# ── N.Pedido Management ───────────────────────────────────────────────────────

def delete_n_pedido_lines(project_id: int, n_pedido: str) -> int:
    conn = _connect()
    try:
        with conn.cursor() as cur:
            cur.execute(
                "SELECT COUNT(*) FROM lines WHERE packing_list_id IN "
                "(SELECT id FROM packing_lists WHERE project_id = %s AND n_pedido = %s)",
                (project_id, n_pedido),
            )
            count = cur.fetchone()[0]
            cur.execute(
                "DELETE FROM lines WHERE packing_list_id IN "
                "(SELECT id FROM packing_lists WHERE project_id = %s AND n_pedido = %s)",
                (project_id, n_pedido),
            )
            cur.execute(
                "DELETE FROM packing_lists WHERE project_id = %s AND n_pedido = %s",
                (project_id, n_pedido),
            )
        conn.commit()
        return count
    finally:
        conn.close()


def get_unique_n_pedidos(project_id: int) -> list[str]:
    conn = _connect()
    try:
        with conn.cursor() as cur:
            cur.execute(
                "SELECT DISTINCT n_pedido FROM packing_lists "
                "WHERE project_id = %s AND n_pedido IS NOT NULL ORDER BY n_pedido",
                (project_id,),
            )
            return [row[0] for row in cur.fetchall()]
    finally:
        conn.close()


def count_n_pedido_lines(project_id: int, n_pedido: str) -> int:
    conn = _connect()
    try:
        with conn.cursor() as cur:
            cur.execute(
                "SELECT COUNT(*) FROM lines WHERE packing_list_id IN "
                "(SELECT id FROM packing_lists WHERE project_id = %s AND n_pedido = %s)",
                (project_id, n_pedido),
            )
            return cur.fetchone()[0]
    finally:
        conn.close()


def get_paquete_nums_for_n_pedido(project_id: int, n_pedido: str) -> list[dict]:
    """Return paquete_num, marca, piezas for all lines of a given n_pedido."""
    conn = _connect()
    try:
        with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
            cur.execute(
                """SELECT l.paquete_num, l.marca, l.piezas
                   FROM lines l
                   JOIN packing_lists pl ON pl.id = l.packing_list_id
                   WHERE pl.project_id = %s AND pl.n_pedido = %s""",
                (project_id, n_pedido),
            )
            return [dict(r) for r in cur.fetchall()]
    finally:
        conn.close()


def get_n_pedido_for_paquete(project_id: int, paquete_num: int) -> Optional[str]:
    conn = _connect()
    try:
        with conn.cursor() as cur:
            cur.execute(
                """SELECT DISTINCT pl.n_pedido FROM packing_lists pl
                   JOIN lines l ON pl.id = l.packing_list_id
                   WHERE pl.project_id = %s AND l.paquete_num = %s
                   LIMIT 1""",
                (project_id, paquete_num),
            )
            row = cur.fetchone()
            return row[0] if row else None
    finally:
        conn.close()


# ── Manual Line Insertion ─────────────────────────────────────────────────────

def insert_line(
    project_id: int,
    paquete_num: int,
    paquete_num_of: int | None,
    kilos_paquete: float | None,
    linea: int | None,
    piezas: int | None,
    longitud: float | None,
    marca: str | None,
) -> int:
    dummy_of = "_MANUAL_ENTRIES"
    conn = _connect()
    try:
        with conn.cursor() as cur:
            cur.execute(
                "SELECT id FROM packing_lists WHERE project_id = %s AND of_number = %s",
                (project_id, dummy_of),
            )
            pl_row = cur.fetchone()

            if not pl_row:
                now = datetime.now(timezone.utc).isoformat()
                cur.execute(
                    """INSERT INTO packing_lists
                       (project_id, of_number, n_pedido, articulo, ref_pedido,
                        nota, total_piezas, kilos_teoricos, imported_at)
                       VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s)
                       RETURNING id""",
                    (project_id, dummy_of, None, None, None, None, None, None, now),
                )
                pl_id = cur.fetchone()[0]
            else:
                pl_id = pl_row[0]

            cur.execute(
                """INSERT INTO lines
                   (packing_list_id, paquete_num, paquete_num_of, kilos_paquete,
                    linea, piezas, longitud, marca)
                   VALUES (%s,%s,%s,%s,%s,%s,%s,%s)
                   RETURNING id""",
                (pl_id, paquete_num, paquete_num_of, kilos_paquete,
                 linea, piezas, longitud, marca),
            )
            line_id = cur.fetchone()[0]
        conn.commit()
        return line_id
    finally:
        conn.close()


# ── Excel File Storage ────────────────────────────────────────────────────────

def save_project_excel(project_id: int, excel_data: bytes, filename: str = "") -> None:
    """Store or replace the Excel base file for a project in the database."""
    now = datetime.now(timezone.utc).isoformat()
    conn = _connect()
    try:
        with conn.cursor() as cur:
            cur.execute(
                """INSERT INTO project_excel_files (project_id, excel_data, filename, updated_at)
                   VALUES (%s, %s, %s, %s)
                   ON CONFLICT (project_id) DO UPDATE
                   SET excel_data = EXCLUDED.excel_data,
                       filename   = EXCLUDED.filename,
                       updated_at = EXCLUDED.updated_at""",
                (project_id, psycopg2.Binary(excel_data), filename, now),
            )
        conn.commit()
    finally:
        conn.close()


def load_project_excel(project_id: int) -> Optional[bytes]:
    """Load the Excel base file for a project. Returns None if not found."""
    conn = _connect()
    try:
        with conn.cursor() as cur:
            cur.execute(
                "SELECT excel_data FROM project_excel_files WHERE project_id = %s",
                (project_id,),
            )
            row = cur.fetchone()
            return bytes(row[0]) if row else None
    finally:
        conn.close()


def excel_exists_in_db(project_id: int) -> bool:
    conn = _connect()
    try:
        with conn.cursor() as cur:
            cur.execute(
                "SELECT 1 FROM project_excel_files WHERE project_id = %s",
                (project_id,),
            )
            return cur.fetchone() is not None
    finally:
        conn.close()


def delete_project_excel_from_db(project_id: int) -> bool:
    conn = _connect()
    try:
        with conn.cursor() as cur:
            cur.execute(
                "DELETE FROM project_excel_files WHERE project_id = %s RETURNING id",
                (project_id,),
            )
            deleted = cur.fetchone() is not None
        conn.commit()
        return deleted
    finally:
        conn.close()
