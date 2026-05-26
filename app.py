import os
import streamlit as st
import pandas as pd
import time
from dotenv import load_dotenv

load_dotenv(os.path.join(os.path.dirname(__file__), ".env"))

import config
import database
import extractor
import exporter
import excel_base
from config import load_models, save_models

database.init_db()

st.set_page_config(page_title="Gestor de Paquetes", layout="wide")

# ── Custom Styles ────────────────────────────────────────────────────────────

st.markdown("""
<style>
    button {
        background-color: #01547e !important;
        color: white !important;
        box-shadow: 0 4px 8px rgba(1, 84, 126, 0.4) !important;
        border: none !important;
        transition: all 0.3s ease !important;
    }
    button:hover {
        background-color: #01547e !important;
        box-shadow: 0 6px 12px rgba(1, 84, 126, 0.6) !important;
    }
    button:active {
        background-color: #e43739 !important;
        box-shadow: 0 2px 4px rgba(228, 55, 57, 0.4) !important;
    }
</style>
""", unsafe_allow_html=True)

# ── Authentication ────────────────────────────────────────────────────────────

def _check_password():
    if "password_correct" not in st.session_state:
        st.session_state.password_correct = False

    if st.session_state.password_correct:
        return True

    st.markdown("# 🔐 Gestor de Paquetes")
    st.markdown("Introduce la contraseña para continuar")

    with st.form("login_form"):
        password = st.text_input("Contraseña", type="password", key="password_input")
        submitted = st.form_submit_button("Entrar", key="login_btn")

    correct_password = config._get_secret("STREAMLIT_PASSWORD")

    if submitted:
        if password == correct_password and correct_password:
            st.session_state.password_correct = True
            st.rerun()
        else:
            st.error("❌ Contraseña incorrecta")

    st.stop()

_check_password()

# ── Helpers ───────────────────────────────────────────────────────────────────

def _project_choices() -> list[str]:
    return [p["name"] for p in database.get_projects()]

def _lines_table(project_name: str):
    if not project_name:
        return pd.DataFrame()
    pid = database.get_project_id(project_name)
    if pid is None:
        return pd.DataFrame()
    rows = database.get_project_lines(pid)
    if not rows:
        return pd.DataFrame(columns=["OF", "N.Pedido", "Artículo", "Paquete_Num",
                                      "Paquete_Num_OF", "Kilos", "Línea", "Piezas",
                                      "Longitud", "Marca"])
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
    cols = ["OF", "N.Pedido", "Artículo", "Paquete_Num", "Paquete_Num_OF",
            "Kilos", "Línea", "Piezas", "Longitud", "Marca"]
    return df[[c for c in cols if c in df.columns]]

# ── Initialize session state ──────────────────────────────────────────────────

if "selected_project" not in st.session_state:
    projects = _project_choices()
    st.session_state.selected_project = projects[0] if projects else ""

if "show_delete_confirm" not in st.session_state:
    st.session_state.show_delete_confirm = False

if "delete_project_name" not in st.session_state:
    st.session_state.delete_project_name = None

if "show_delete_n_pedido_confirm" not in st.session_state:
    st.session_state.show_delete_n_pedido_confirm = False

if "delete_n_pedido_value" not in st.session_state:
    st.session_state.delete_n_pedido_value = None

if "step2_provider" not in st.session_state:
    st.session_state.step2_provider = "nvidia"

if "excel_preview_update_needed" not in st.session_state:
    st.session_state.excel_preview_update_needed = True

if "logistica_excel_bytes" not in st.session_state:
    st.session_state.logistica_excel_bytes = None

if "logistica_excel_filename" not in st.session_state:
    st.session_state.logistica_excel_filename = None

if "logistica_cache_project" not in st.session_state:
    st.session_state.logistica_cache_project = None

if "pdf_upload_key" not in st.session_state:
    st.session_state.pdf_upload_key = 0

if "excel_upload_key" not in st.session_state:
    st.session_state.excel_upload_key = 0

if "show_excel_delete_confirm" not in st.session_state:
    st.session_state.show_excel_delete_confirm = False

if "excel_rows_to_delete" not in st.session_state:
    st.session_state.excel_rows_to_delete = []

if "show_excel_rows_delete_confirm" not in st.session_state:
    st.session_state.show_excel_rows_delete_confirm = False

if "show_excel_delete_all_confirm" not in st.session_state:
    st.session_state.show_excel_delete_all_confirm = False

if "excel_delete_rows_cache" not in st.session_state:
    st.session_state.excel_delete_rows_cache = None

if "excel_delete_cache_project" not in st.session_state:
    st.session_state.excel_delete_cache_project = None

if "show_delete_line_confirm" not in st.session_state:
    st.session_state.show_delete_line_confirm = False

if "delete_line_id" not in st.session_state:
    st.session_state.delete_line_id = None

if "delete_line_label" not in st.session_state:
    st.session_state.delete_line_label = ""

if "edit_line_last_id" not in st.session_state:
    st.session_state.edit_line_last_id = None

if "edit_of" not in st.session_state:
    st.session_state.edit_of = ""

if "edit_n_pedido" not in st.session_state:
    st.session_state.edit_n_pedido = ""

if "edit_articulo" not in st.session_state:
    st.session_state.edit_articulo = ""

# ── Sidebar Configuration ──────────────────────────────────────────────────────

with st.sidebar:
    st.markdown("## Configuración")

    api_key = config.load_api_key()
    nvidia_key = config.load_nvidia_api_key()

    if nvidia_key:
        st.caption("✅ NVIDIA API Key cargada")
    else:
        st.error("❌ NVIDIA API Key no encontrada\n(NVIDIA_API_KEY / NVIDIA_PAI_KEY)")

    st.markdown("---")
    st.markdown("**Método de extracción**")

    extraction_method = st.radio(
        "Método de extracción",
        options=["nemotron", "pdfplumber"],
        format_func=lambda x: "Nemotron-Parse + Text Model" if x == "nemotron" else "pdfPlumber",
        key="extraction_method",
        index=1,
        horizontal=False,
        label_visibility="collapsed",
    )

    models_list = load_models()
    current_model = models_list[0] if models_list else ""
    nvidia_text_models = config.load_nvidia_text_models()
    current_nvidia_model = nvidia_text_models[0] if nvidia_text_models else ""

    if extraction_method == "nemotron":
        st.markdown("---")
        st.markdown("**Modelo OpenRouter** *(paso 2 — OpenRouter)*")

        current_model = st.selectbox(
            "Modelo OpenRouter activo",
            options=models_list,
            index=0 if models_list else 0,
            key="model_select",
            label_visibility="collapsed",
        )

        col1, col2, col3 = st.columns([3, 1, 1])
        with col1:
            new_model = st.text_input(
                "Nuevo modelo OpenRouter",
                placeholder="proveedor/modelo:tag",
                key="new_model_input",
                label_visibility="collapsed",
            )
        with col2:
            if st.button("➕", key="add_model_btn"):
                new_model = new_model.strip()
                if not new_model:
                    st.error("Introduce un modelo.")
                elif new_model in models_list:
                    st.error("Modelo ya existe.")
                else:
                    models_list.append(new_model)
                    save_models(models_list)
                    st.success("Modelo añadido.")
                    st.rerun()
        with col3:
            if st.button("➖", key="del_model_btn"):
                if len(models_list) <= 1:
                    st.error("Debe quedar al menos un modelo.")
                else:
                    models_list.remove(current_model)
                    save_models(models_list)
                    st.rerun()

        st.markdown("---")
        st.markdown("**Modelo NVIDIA** *(paso 2 — NVIDIA)*")

        current_nvidia_model = st.selectbox(
            "Modelo NVIDIA activo",
            options=nvidia_text_models,
            index=0 if nvidia_text_models else 0,
            key="nvidia_model_select",
            label_visibility="collapsed",
        )

        col1, col2, col3 = st.columns([3, 1, 1])
        with col1:
            new_nvidia_model = st.text_input(
                "Nuevo modelo NVIDIA",
                placeholder="nvidia/modelo",
                key="new_nvidia_model_input",
                label_visibility="collapsed",
            )
        with col2:
            if st.button("➕", key="add_nvidia_model_btn"):
                new_nvidia_model = new_nvidia_model.strip()
                if not new_nvidia_model:
                    st.error("Introduce un modelo.")
                elif new_nvidia_model in nvidia_text_models:
                    st.error("Modelo ya existe.")
                else:
                    nvidia_text_models.append(new_nvidia_model)
                    config.save_nvidia_text_models(nvidia_text_models)
                    st.success("Modelo añadido.")
                    st.rerun()
        with col3:
            if st.button("➖", key="del_nvidia_model_btn"):
                if len(nvidia_text_models) <= 1:
                    st.error("Debe quedar al menos un modelo.")
                else:
                    nvidia_text_models.remove(current_nvidia_model)
                    config.save_nvidia_text_models(nvidia_text_models)
                    st.rerun()

        st.markdown("---")
        st.markdown("**Proveedor paso 2**")
        st.caption("Paso 1: siempre NVIDIA nemotron-parse (imagen→markdown)")

        step2_provider = st.radio(
            "Proveedor paso 2",
            options=["nvidia", "openrouter"],
            format_func=lambda x: "NVIDIA" if x == "nvidia" else "OpenRouter",
            key="step2_provider",
            horizontal=True,
            label_visibility="collapsed",
        )
        if step2_provider == "nvidia":
            st.caption(f"Modelo: {current_nvidia_model}")
        else:
            if api_key:
                st.caption(f"Modelo: {current_model}")
            else:
                st.error("❌ OpenRouter API Key no encontrada en .env")
    else:
        st.markdown("---")
        st.markdown("**Método**")
        st.caption("pdfPlumber")

# ── Main Title ──────────────────────────────────────────────────────────────

st.markdown("# Gestor de Paquetes")
st.markdown("Gestión de packing lists por proyecto")
st.markdown("---")

# ── Section 1: Project Selection ────────────────────────────────────────────────

st.markdown("## 1. Proyecto activo")

col1, col2, col3 = st.columns([3, 2, 1])

with col1:
    projects = _project_choices()
    if projects:
        selected = st.selectbox(
            "Seleccionar proyecto",
            options=projects,
            index=projects.index(st.session_state.selected_project)
                  if st.session_state.selected_project in projects else 0,
            key="project_select",
        )
        st.session_state.selected_project = selected
    else:
        st.session_state.selected_project = ""
        st.info("No hay proyectos creados.")

with col2:
    new_project_name = st.text_input(
        "Nuevo proyecto",
        placeholder="Nombre del proyecto...",
        key="new_project_input",
    )

with col3:
    if st.button("Crear", key="create_btn"):
        new_project_name = new_project_name.strip()
        if not new_project_name:
            st.error("Introduce un nombre.")
        else:
            try:
                database.create_project(new_project_name)
                st.session_state.selected_project = new_project_name
                st.success(f"Proyecto **{new_project_name}** creado.")
                st.rerun()
            except Exception as e:
                st.error(f"Error: {e}")

col1, col2 = st.columns(2)

with col1:
    if st.button("🔄 Actualizar", key="refresh_btn"):
        st.rerun()

with col2:
    if st.button("🗑️ Eliminar", key="delete_btn"):
        if not st.session_state.selected_project:
            st.error("Selecciona un proyecto primero.")
        else:
            st.session_state.show_delete_confirm = True
            st.session_state.delete_project_name = st.session_state.selected_project
            st.rerun()

if st.session_state.show_delete_confirm:
    st.warning(f"⚠️ **Eliminar permanentemente {st.session_state.delete_project_name} y todo su contenido?**")
    col1, col2 = st.columns(2)
    with col1:
        if st.button("Sí, eliminar", key="confirm_delete_btn"):
            try:
                pid = database.get_project_id(st.session_state.delete_project_name)
                if pid:
                    database.delete_project(pid)
                projects = _project_choices()
                st.session_state.selected_project = projects[0] if projects else ""
                st.session_state.show_delete_confirm = False
                st.success(f"Proyecto **{st.session_state.delete_project_name}** eliminado.")
                st.rerun()
            except Exception as e:
                st.error(f"Error al eliminar: {e}")
    with col2:
        if st.button("Cancelar", key="cancel_delete_btn"):
            st.session_state.show_delete_confirm = False
            st.rerun()

st.markdown("---")

# ── Section 2: Import Packing Lists ─────────────────────────────────────────────

st.markdown("## 2. Importar packing list")

uploaded_files = st.file_uploader(
    "Subir PDFs", type="pdf", accept_multiple_files=True,
    key=f"pdf_upload_{st.session_state.pdf_upload_key}",
)

if st.button("🔎 Procesar Documento", key="process_btn"):
    _step2 = st.session_state.get("step2_provider", "nvidia")
    if not uploaded_files:
        st.error("Sube uno o más archivos PDF primero.")
    elif not st.session_state.selected_project:
        st.error("Selecciona o crea un proyecto primero.")
    elif not nvidia_key:
        st.error("❌ NVIDIA API Key no encontrada (NVIDIA_API_KEY / NVIDIA_PAI_KEY)")
    elif _step2 == "openrouter" and not api_key:
        st.error("❌ OpenRouter API Key no encontrada para el paso 2.")
    else:
        progress_container = st.container()
        messages_list = []
        all_dataframes = []
        total = len(uploaded_files)

        with progress_container:
            progress_bar = st.progress(0)
            status_text = st.empty()

        for i, uploaded_file in enumerate(uploaded_files, 1):
            filename = uploaded_file.name
            try:
                t0 = time.time()
                status_text.text(f"[{i}/{total}] {filename} — Leyendo PDF...")
                pdf_bytes = uploaded_file.read()

                def _on_status(msg: str, _i=i, _total=total, _name=filename):
                    status_text.text(f"[{_i}/{_total}] {_name} — {msg}")

                t1 = time.time()
                data = extractor.extract_data(
                    pdf_bytes,
                    extraction_method=st.session_state.get("extraction_method", "nemotron"),
                    openrouter_model=current_model,
                    openrouter_api_key=api_key,
                    nvidia_api_key=nvidia_key,
                    step2_provider=st.session_state.get("step2_provider", "nvidia"),
                    step2_nvidia_model=current_nvidia_model,
                    on_status=_on_status,
                )
                t2 = time.time()

                method_used = data.get("extraction_method", "unknown")
                if "fallback" in method_used:
                    st.warning(f"⚠️ Método utilizado: {method_used}")
                else:
                    st.success(f"✓ Método utilizado: {method_used}")

                pid = database.get_project_id(st.session_state.selected_project)
                pl_id = database.save_packing_list(pid, data)
                t3 = time.time()

                lines = database.get_project_lines(pid)
                of_lines = [l for l in lines if l["of_number"] == data["of_number"]]
                if of_lines:
                    all_dataframes.append(pd.DataFrame(of_lines))

                tiempo_api = f"{t2-t1:.1f}s"
                tiempo_bd = f"{t3-t2:.2f}s"
                msg = (f"✓ {filename}: {data['of_number']} importado "
                       f"({len(of_lines)} líneas) [API:{tiempo_api} BD:{tiempo_bd}]")
                messages_list.append(msg)

            except ValueError as e:
                # OF duplicado — aviso específico, no error genérico
                messages_list.append(f"⚠️ {filename}: {str(e)}")
            except Exception as e:
                messages_list.append(f"✗ {filename}: {str(e)}")

            progress_bar.progress(i / total)

        status_text.empty()
        progress_bar.empty()

        for msg in messages_list:
            if msg.startswith("✓"):
                st.success(msg)
            elif msg.startswith("⚠️"):
                st.warning(msg)
            else:
                st.error(msg)

        if all_dataframes:
            combined_df = pd.concat(all_dataframes, ignore_index=True)
            st.markdown("### Líneas extraídas")
            st.dataframe(combined_df, use_container_width=True)

        # Limpiar el file uploader
        st.session_state.pdf_upload_key += 1
        st.rerun()

st.markdown("---")

# ── Section 3: Project Lines ────────────────────────────────────────────────

if st.session_state.selected_project:
    pid = database.get_project_id(st.session_state.selected_project)
    lines = database.get_project_lines(pid)
    max_paquete = max((l["paquete_num"] for l in lines), default=0)
else:
    lines = []
    max_paquete = 0

st.markdown(f"## 3. Líneas de paquetes del proyecto [{max_paquete} paquetes]")

col1, col2 = st.columns(2)

with col1:
    if st.button("🔄 Refrescar tabla", key="refresh_lines_btn"):
        st.session_state.logistica_cache_project = None  # forzar regeneración
        st.rerun()

# Generar Excel de logística una sola vez por proyecto (cache en session_state)
if lines and st.session_state.selected_project:
    if st.session_state.logistica_cache_project != st.session_state.selected_project:
        try:
            st.session_state.logistica_excel_bytes = exporter.build_excel(lines)
            st.session_state.logistica_excel_filename = (
                f"{st.session_state.selected_project}-Logística.xlsx"
            )
            st.session_state.logistica_cache_project = st.session_state.selected_project
        except Exception as e:
            st.error(f"Error generando Logística: {str(e)}")

with col2:
    if st.session_state.logistica_excel_bytes and lines:
        st.download_button(
            label="⬇️ Exportar Logística",
            data=st.session_state.logistica_excel_bytes,
            file_name=st.session_state.logistica_excel_filename or "Logistica.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="logistica_download_btn",
        )
    elif not lines:
        st.button("🚚 Exportar Logística", key="export_btn_disabled", disabled=True)

lines_df = _lines_table(st.session_state.selected_project)
if not lines_df.empty:
    st.dataframe(lines_df, use_container_width=True, hide_index=True)
else:
    st.info("No hay líneas para mostrar. Importa un packing list primero.")

st.markdown("### Añadir nueva línea")

# Pending success message from a previous add (shown after rerun)
_add_msg = st.session_state.pop("_add_success_msg", None)
if _add_msg:
    st.success(_add_msg)

if not st.session_state.selected_project:
    st.info("Selecciona un proyecto primero para añadir líneas.")
else:
    # Compute next paquete num BEFORE the form so it reflects the latest DB state
    _next_paquete_num = int(database.get_next_paquete_num(
        database.get_project_id(st.session_state.selected_project)
    ))

    with st.form("add_line_form", clear_on_submit=True):
        # Row 0: packing-list level fields
        col_of, col_npedido, col_articulo = st.columns(3)
        # Row 1: line-level numeric fields
        col1, col2, col3, col4 = st.columns(4)
        # Row 2: remaining line fields + button
        col5, col6, col7, col8 = st.columns(4)

        with col_of:
            new_of = st.text_input("OF", value="",
                                   help="Número de OF. Si ya existe se reutiliza su packing list.")
        with col_npedido:
            new_n_pedido = st.text_input("N.Pedido", value="",
                                         help="Solo se aplica si la OF es nueva.")
        with col_articulo:
            new_articulo = st.text_input("Artículo", value="",
                                         help="Solo se aplica si la OF es nueva.")
        with col1:
            paquete_num = st.number_input("Paquete Num", min_value=1, value=_next_paquete_num)
        with col2:
            paquete_num_of = st.number_input("Paquete OF", min_value=0, value=1)
        with col3:
            kilos = st.number_input("Kilos", min_value=0.0, value=0.0, step=0.1)
        with col4:
            linea = st.number_input("Línea", min_value=1, value=1)
        with col5:
            piezas = st.number_input("Piezas", min_value=0, value=1)
        with col6:
            longitud = st.number_input("Longitud (m)", min_value=0.0, value=0.0, step=0.001)
        with col7:
            marca = st.text_input("Marca", value="")
        with col8:
            st.write("")
            submitted_add = st.form_submit_button("➕ Añadir")

    if submitted_add:
        pid = database.get_project_id(st.session_state.selected_project)
        try:
            database.insert_line(
                project_id=pid,
                paquete_num=int(paquete_num),
                paquete_num_of=int(paquete_num_of) if paquete_num_of else None,
                kilos_paquete=float(kilos) if kilos else None,
                linea=int(linea) if linea else None,
                piezas=int(piezas) if piezas else None,
                longitud=float(longitud) if longitud else None,
                marca=marca.strip() if marca.strip() else None,
                of_number=new_of.strip() if new_of.strip() else None,
                n_pedido=new_n_pedido.strip() if new_n_pedido.strip() else None,
                articulo=new_articulo.strip() if new_articulo.strip() else None,
            )
            st.session_state.logistica_cache_project = None
            st.session_state._add_success_msg = (
                f"✓ Línea añadida (Paquete {paquete_num}, Línea {linea})"
            )
            st.rerun()
        except Exception as e:
            st.error(f"Error al añadir línea: {str(e)}")

st.markdown("---")

# ── Fetch lines with IDs — reused by both Eliminar and Editar subsections ────

_lines_with_ids: list[dict] = []
if st.session_state.selected_project:
    _pid_crud = database.get_project_id(st.session_state.selected_project)
    if _pid_crud:
        _lines_with_ids = database.get_project_lines_with_ids(_pid_crud)


def _line_label(l: dict) -> str:
    of_display = "Manual" if l["of_number"] == "_MANUAL_ENTRIES" else (l["of_number"] or "-")
    return (
        f"[{l['line_id']}] Paq {l['paquete_num']} | OF {of_display} "
        f"| L{l['linea']} | {l['marca'] or '-'}"
    )


# ── Eliminar línea individual ──────────────────────────────────────────────────

st.markdown("### Eliminar línea")

if not st.session_state.selected_project:
    st.info("Selecciona un proyecto primero para eliminar líneas.")
elif not _lines_with_ids:
    st.info("No hay líneas para eliminar.")
else:
    _del_options = {l["line_id"]: _line_label(l) for l in _lines_with_ids}

    col_del1, col_del2 = st.columns([4, 1])
    with col_del1:
        _del_selected_id = st.selectbox(
            "Seleccionar línea a eliminar",
            options=list(_del_options.keys()),
            format_func=lambda x: _del_options[x],
            key="delete_line_select",
        )
    with col_del2:
        st.write("")  # vertical alignment
        if st.button("🗑️ Eliminar", key="delete_single_line_btn"):
            st.session_state.delete_line_id = _del_selected_id
            st.session_state.delete_line_label = _del_options.get(_del_selected_id, "")
            st.session_state.show_delete_line_confirm = True
            st.rerun()

if st.session_state.show_delete_line_confirm and st.session_state.delete_line_id is not None:
    st.warning(
        f"⚠️ **¿Eliminar permanentemente esta línea?**\n\n"
        f"`{st.session_state.delete_line_label}`\n\n"
        "Esta acción no se puede deshacer."
    )
    col_yes, col_no = st.columns(2)
    with col_yes:
        if st.button("Sí, eliminar", key="confirm_delete_single_line_btn"):
            try:
                deleted = database.delete_line_by_id(st.session_state.delete_line_id)
                st.session_state.show_delete_line_confirm = False
                st.session_state.delete_line_id = None
                st.session_state.delete_line_label = ""
                st.session_state.logistica_cache_project = None
                if deleted:
                    st.success("✅ Línea eliminada correctamente.")
                else:
                    st.warning("La línea no se encontró (puede que ya hubiera sido eliminada).")
                st.rerun()
            except Exception as e:
                st.error(f"Error al eliminar: {str(e)}")
    with col_no:
        if st.button("Cancelar", key="cancel_delete_single_line_btn"):
            st.session_state.show_delete_line_confirm = False
            st.session_state.delete_line_id = None
            st.rerun()

st.markdown("---")

# ── Editar línea existente ────────────────────────────────────────────────────

st.markdown("### Editar línea existente")

# Pending success message from a previous update (shown after rerun)
_edit_msg = st.session_state.pop("_edit_success_msg", None)
if _edit_msg:
    st.success(_edit_msg)

if not st.session_state.selected_project:
    st.info("Selecciona un proyecto primero para editar líneas.")
elif not _lines_with_ids:
    st.info("No hay líneas para editar.")
else:
    _edit_options = {l["line_id"]: _line_label(l) for l in _lines_with_ids}

    # Selectbox is OUTSIDE the form so changing it immediately reloads the line data
    _edit_selected_id = st.selectbox(
        "Seleccionar línea a editar",
        options=list(_edit_options.keys()),
        format_func=lambda x: _edit_options[x],
        key="edit_line_select",
    )

    # Locate the full dict for the selected line
    _edit_current = next((l for l in _lines_with_ids if l["line_id"] == _edit_selected_id), None)

    if _edit_current:
        # When the selected line changes (or after a successful update), reload
        # session_state from the DB values so the form shows fresh data.
        if _edit_selected_id != st.session_state.edit_line_last_id:
            _of_raw = _edit_current["of_number"] or ""
            st.session_state.edit_of           = "" if _of_raw == "_MANUAL_ENTRIES" else _of_raw
            st.session_state.edit_n_pedido     = str(_edit_current["n_pedido"] or "")
            st.session_state.edit_articulo     = str(_edit_current["articulo"] or "")
            st.session_state.edit_paquete_num    = int(_edit_current["paquete_num"] or 1)
            st.session_state.edit_paquete_num_of = int(_edit_current["paquete_num_of"] or 0)
            st.session_state.edit_kilos          = float(_edit_current["kilos_paquete"] or 0.0)
            st.session_state.edit_linea          = int(_edit_current["linea"] or 1)
            st.session_state.edit_piezas         = int(_edit_current["piezas"] or 0)
            st.session_state.edit_longitud       = float(_edit_current["longitud"] or 0.0)
            st.session_state.edit_marca          = str(_edit_current["marca"] or "")
            st.session_state.edit_line_last_id   = _edit_selected_id

        # All edit fields inside a form — no rerun on each keystroke
        with st.form("edit_line_form"):
            # Row 0: packing-list level fields (editable)
            col_eof, col_enpedido, col_earticulo = st.columns(3)
            # Row 1 & 2: line-level fields
            col_e1, col_e2, col_e3, col_e4 = st.columns(4)
            col_e5, col_e6, col_e7, col_e8 = st.columns(4)

            with col_eof:
                e_of = st.text_input("OF", key="edit_of")
            with col_enpedido:
                e_n_pedido = st.text_input("N.Pedido", key="edit_n_pedido")
            with col_earticulo:
                e_articulo = st.text_input("Artículo", key="edit_articulo")
            with col_e1:
                e_paquete_num = st.number_input("Paquete Num", min_value=1, key="edit_paquete_num")
            with col_e2:
                e_paquete_num_of = st.number_input("Paquete OF", min_value=0, key="edit_paquete_num_of")
            with col_e3:
                e_kilos = st.number_input("Kilos", min_value=0.0, step=0.1, key="edit_kilos")
            with col_e4:
                e_linea = st.number_input("Línea", min_value=1, key="edit_linea")
            with col_e5:
                e_piezas = st.number_input("Piezas", min_value=0, key="edit_piezas")
            with col_e6:
                e_longitud = st.number_input("Longitud (m)", min_value=0.0, step=0.001, key="edit_longitud")
            with col_e7:
                e_marca = st.text_input("Marca", key="edit_marca")
            with col_e8:
                st.write("")
            submit_edit = st.form_submit_button("✏️ Actualizar")

        if submit_edit:
            _pid_edit = database.get_project_id(st.session_state.selected_project)
            try:
                updated = database.update_line_full(
                    project_id=_pid_edit,
                    line_id=_edit_selected_id,
                    of_number=e_of.strip() if e_of.strip() else None,
                    n_pedido=e_n_pedido.strip() if e_n_pedido.strip() else None,
                    articulo=e_articulo.strip() if e_articulo.strip() else None,
                    paquete_num=int(e_paquete_num),
                    paquete_num_of=int(e_paquete_num_of) if e_paquete_num_of else None,
                    kilos_paquete=float(e_kilos) if e_kilos else None,
                    linea=int(e_linea) if e_linea else None,
                    piezas=int(e_piezas) if e_piezas else None,
                    longitud=float(e_longitud) if e_longitud else None,
                    marca=e_marca.strip() if e_marca.strip() else None,
                )
                st.session_state.logistica_cache_project = None
                # Force reload of form values from the freshly updated DB row
                st.session_state.edit_line_last_id = None
                if updated:
                    st.session_state._edit_success_msg = "✅ Línea actualizada correctamente."
                else:
                    st.session_state._edit_success_msg = "⚠️ La línea no se encontró."
                st.rerun()
            except Exception as e:
                st.error(f"Error al actualizar: {str(e)}")

st.markdown("---")

# ── Delete N.Pedido Subsection ────────────────────────────────────────

st.markdown("### Eliminar líneas de una OF completa")

if not st.session_state.selected_project:
    st.info("Selecciona un proyecto primero para eliminar N.Pedidos.")
else:
    col1, col2 = st.columns([2, 1])

    try:
        pid = database.get_project_id(st.session_state.selected_project)
        n_pedidos = database.get_unique_n_pedidos(pid) if pid else []
    except Exception as e:
        st.error(f"Error al cargar N.Pedidos: {str(e)}")
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

if st.session_state.show_delete_n_pedido_confirm and st.session_state.delete_n_pedido_value:
    n_pedido = st.session_state.delete_n_pedido_value
    pid = database.get_project_id(st.session_state.selected_project)

    try:
        count_db = database.count_n_pedido_lines(pid, n_pedido)
        count_excel = excel_base.count_n_pedido_rows_in_excel(
            st.session_state.selected_project, n_pedido
        )

        excel_note = (
            f"\n- {count_excel} filas del Excel base"
            if count_excel > 0
            else "\n- 0 filas del Excel base (o no hay Excel cargado)"
        )
        st.warning(
            f"⚠️ **Eliminar permanentemente N.Pedido {n_pedido}?**\n\n"
            f"Se eliminarán:\n- {count_db} líneas de la base de datos{excel_note}"
        )

        col1, col2 = st.columns(2)
        with col1:
            if st.button("Sí, eliminar", key="confirm_delete_n_pedido_btn"):
                try:
                    deleted_db = database.delete_n_pedido_lines(pid, n_pedido)
                    deleted_excel = 0
                    if excel_base.excel_exists(st.session_state.selected_project):
                        try:
                            deleted_excel = excel_base.delete_n_pedido_from_excel(
                                st.session_state.selected_project, n_pedido
                            )
                        except Exception as excel_error:
                            st.warning(f"⚠️ Eliminado de BD pero error en Excel: {str(excel_error)}")
                    st.session_state.show_delete_n_pedido_confirm = False
                    st.success(
                        f"N.Pedido {n_pedido} eliminado: "
                        f"{deleted_db} líneas de BD, {deleted_excel} filas de Excel"
                    )
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

# ── Section 4: Excel Base ───────────────────────────────────────────────────────

st.markdown("## 4. Excel base del proyecto")

if not st.session_state.selected_project:
    st.info("Selecciona un proyecto primero para gestionar su Excel base.")
else:
    pid = database.get_project_id(st.session_state.selected_project)

    has_excel = excel_base.excel_exists(st.session_state.selected_project)

    # ── Excel Status + Delete Button ──────────────────────────────────────────
    if has_excel:
        col_status, col_del = st.columns([5, 1])
        with col_status:
            st.success("✅ **Excel base cargado** — guardado en base de datos")
        with col_del:
            if st.button("🗑️ Eliminar", key="excel_delete_btn"):
                st.session_state.show_excel_delete_confirm = True
                st.rerun()
    else:
        st.warning("⚠️ **Excel base no cargado** — sube un archivo .xlsx para habilitar la sincronización")

    # ── Confirm Excel Delete ───────────────────────────────────────────────────
    if st.session_state.get("show_excel_delete_confirm"):
        st.warning("⚠️ **¿Eliminar el Excel base de este proyecto?** Esta acción no se puede deshacer.")
        col_yes, col_no = st.columns(2)
        with col_yes:
            if st.button("Sí, eliminar", key="confirm_excel_delete_btn"):
                try:
                    excel_base.delete_project_excel(st.session_state.selected_project)
                    st.session_state.show_excel_delete_confirm = False
                    st.session_state.excel_preview_update_needed = True
                    st.session_state.excel_upload_key += 1
                    st.session_state.excel_delete_rows_cache = None
                    st.success("Excel base eliminado.")
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ Error: {str(e)}")
        with col_no:
            if st.button("Cancelar", key="cancel_excel_delete_btn"):
                st.session_state.show_excel_delete_confirm = False
                st.rerun()

    st.markdown("---")

    # ── File Uploader — solo visible si no hay Excel cargado ──────────────────
    if not has_excel:
        uploaded_excel = st.file_uploader(
            "Subir Excel base",
            type="xlsx",
            key=f"excel_base_upload_{st.session_state.excel_upload_key}",
            help="El archivo se guardará en la base de datos y persistirá entre sesiones.",
        )
        if uploaded_excel:
            try:
                database.save_project_excel(
                    pid,
                    uploaded_excel.getvalue(),
                    uploaded_excel.name,
                )
                st.success(f"✅ Excel base guardado: {uploaded_excel.name}")
                st.session_state.excel_preview_update_needed = True
                st.session_state.excel_upload_key += 1
                st.session_state.excel_delete_rows_cache = None
                st.rerun()
            except Exception as e:
                st.error(f"❌ Error al guardar Excel: {str(e)}")

    # ── Action Buttons — solo visibles si hay Excel cargado ───────────────────
    if has_excel:
        col1, col2, col3 = st.columns(3)

        with col1:
            if st.button("📊 Actualizar Tabla", key="excel_update_table_btn"):
                st.session_state.excel_preview_update_needed = True
                st.rerun()

        with col2:
            if st.button("🔄 Sincronizar", key="excel_sync_btn"):
                try:
                    project_lines = database.get_project_lines(pid)
                    if not project_lines:
                        st.warning("El proyecto no tiene líneas para exportar.")
                    else:
                        result = excel_base.write_lines_to_excel(
                            st.session_state.selected_project,
                            project_lines,
                        )
                        message = f"✅ {result['added']} filas añadidas"
                        if result["duplicates"] > 0:
                            message += f", {result['duplicates']} duplicados ignorados"
                        st.success(message)
                        for error in result.get("errors", []):
                            st.warning(f"⚠️ {error}")
                        st.session_state.excel_preview_update_needed = True
                        st.session_state.excel_delete_rows_cache = None
                        st.rerun()
                except Exception as e:
                    st.error(f"❌ Error: {str(e)}")

        with col3:
            try:
                excel_bytes = database.load_project_excel(pid)
                if excel_bytes:
                    st.download_button(
                        label="⬇️ Descargar",
                        data=excel_bytes,
                        file_name=f"{st.session_state.selected_project}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="excel_download_file_btn",
                    )
            except Exception as e:
                st.error(f"❌ Error: {str(e)}")

    # ── Delete Rows Expander ───────────────────────────────────────────────────
    if has_excel:
        # Load data from DB only when cache is stale or project has changed.
        # This avoids a DB round-trip on every checkbox rerun.
        if (
            st.session_state.excel_delete_rows_cache is None
            or st.session_state.excel_delete_cache_project != st.session_state.selected_project
        ):
            try:
                preview_data = excel_base.read_excel_preview(st.session_state.selected_project)
                st.session_state.excel_delete_rows_cache = [
                    r for r in preview_data if r["row"] >= 14
                ]
                st.session_state.excel_delete_cache_project = st.session_state.selected_project
            except Exception as e:
                st.warning(f"⚠️ Error cargando filas para eliminación: {e}")
                st.session_state.excel_delete_rows_cache = []

        with st.expander("🗑️ Eliminar filas del Excel base"):
            data_rows = st.session_state.excel_delete_rows_cache

            if not data_rows:
                st.info("No hay filas de datos para eliminar.")
            else:
                col_names = [
                    "A (Paquete)", "B (Tipo)", "C", "D", "E", "F", "G",
                    "H (Marca)", "I (Piezas)",
                ]
                rows_df = pd.DataFrame([
                    {
                        "Eliminar": False,
                        "Fila": r["row"],
                        **{
                            col_names[i]: r["values"][i]
                            for i in range(min(len(r["values"]), len(col_names)))
                        },
                    }
                    for r in data_rows
                ])

                edited_df = st.data_editor(
                    rows_df,
                    column_config={
                        "Eliminar": st.column_config.CheckboxColumn(
                            "Eliminar", default=False, width="small"
                        )
                    },
                    disabled=["Fila"] + col_names,
                    hide_index=True,
                    use_container_width=True,
                    key="excel_rows_editor",
                )

                selected_rows = edited_df[edited_df["Eliminar"]]["Fila"].tolist()
                n_selected = len(selected_rows)

                btn_col1, btn_col2 = st.columns([3, 1])
                with btn_col1:
                    label = (
                        f"🗑️ Eliminar {n_selected} fila{'s' if n_selected != 1 else ''} "
                        f"seleccionada{'s' if n_selected != 1 else ''}"
                    )
                    if st.button(label, key="delete_excel_rows_btn", disabled=(n_selected == 0)):
                        st.session_state.excel_rows_to_delete = selected_rows
                        st.session_state.show_excel_rows_delete_confirm = True
                        st.rerun()
                with btn_col2:
                    if st.button("🗑️ Borrar todo", key="delete_all_excel_rows_btn"):
                        st.session_state.show_excel_delete_all_confirm = True
                        st.rerun()

        # ── Confirm Delete Selected Rows ───────────────────────────────────────
        if st.session_state.show_excel_rows_delete_confirm and st.session_state.excel_rows_to_delete:
            n = len(st.session_state.excel_rows_to_delete)
            st.warning(
                f"⚠️ **¿Eliminar permanentemente {n} fila{'s' if n != 1 else ''} "
                f"del Excel base?** Esta acción no se puede deshacer."
            )
            col_yes, col_no = st.columns(2)
            with col_yes:
                if st.button("Sí, eliminar", key="confirm_delete_excel_rows_btn"):
                    try:
                        deleted = excel_base.delete_rows_from_excel(
                            st.session_state.selected_project,
                            st.session_state.excel_rows_to_delete,
                        )
                        st.session_state.excel_rows_to_delete = []
                        st.session_state.show_excel_rows_delete_confirm = False
                        st.session_state.excel_preview_update_needed = True
                        st.session_state.excel_delete_rows_cache = None
                        st.success(
                            f"✅ {deleted} fila{'s' if deleted != 1 else ''} eliminada"
                            f"{'s' if deleted != 1 else ''} del Excel base."
                        )
                        st.rerun()
                    except Exception as e:
                        st.error(f"❌ Error: {str(e)}")
            with col_no:
                if st.button("Cancelar", key="cancel_delete_excel_rows_btn"):
                    st.session_state.excel_rows_to_delete = []
                    st.session_state.show_excel_rows_delete_confirm = False
                    st.rerun()

        # ── Confirm Delete All Rows ────────────────────────────────────────────
        if st.session_state.show_excel_delete_all_confirm:
            st.warning(
                "⚠️ **¿Eliminar TODAS las filas de datos del Excel base (fila 14 en adelante)?** "
                "Las cabeceras se conservan. Esta acción no se puede deshacer."
            )
            col_yes, col_no = st.columns(2)
            with col_yes:
                if st.button("Sí, borrar todo", key="confirm_delete_all_rows_btn"):
                    try:
                        deleted = excel_base.delete_all_data_rows_from_excel(
                            st.session_state.selected_project
                        )
                        st.session_state.show_excel_delete_all_confirm = False
                        st.session_state.excel_preview_update_needed = True
                        st.session_state.excel_delete_rows_cache = None
                        st.success(f"✅ {deleted} filas eliminadas del Excel base.")
                        st.rerun()
                    except Exception as e:
                        st.error(f"❌ Error: {str(e)}")
            with col_no:
                if st.button("Cancelar", key="cancel_delete_all_rows_btn"):
                    st.session_state.show_excel_delete_all_confirm = False
                    st.rerun()

    # ── Preview Table ──────────────────────────────────────────────────────────
    st.markdown("### Vista previa del Excel base")

    if has_excel:
        if st.session_state.excel_preview_update_needed:
            try:
                preview_data = excel_base.read_excel_preview(st.session_state.selected_project)
                if preview_data:
                    preview_df = pd.DataFrame([
                        {"Fila": item["row"],
                         **{f"Col{i+1}": val for i, val in enumerate(item["values"])}}
                        for item in preview_data
                    ])

                    def highlight_header_rows(row):
                        if row.name in [0, 1]:
                            return ["background-color: #f0f0f0"] * len(row)
                        return [""] * len(row)

                    styled_df = preview_df.style.apply(highlight_header_rows, axis=1)
                    st.dataframe(styled_df, use_container_width=True, hide_index=True, height=875)
                else:
                    st.info("Excel base vacío.")
            except Exception as e:
                st.error(f"❌ Error al mostrar vista previa: {str(e)}")
            finally:
                st.session_state.excel_preview_update_needed = False
        else:
            st.info("⏯️ Presiona 'Actualizar Tabla' o 'Sincronizar' para actualizar la vista previa.")
    else:
        st.info("Sube un Excel base para ver la vista previa.")
