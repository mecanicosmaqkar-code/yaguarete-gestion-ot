import os
import re
import subprocess
from datetime import datetime
import pandas as pd
import openpyxl
import requests
import streamlit as st
import plotly.express as px
import google.generativeai as genai

from docx import Document
from docx.shared import Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docxtpl import DocxTemplate, InlineImage

# ==========================================
# CONFIGURACIÓN DE CLOUDINARY
# ==========================================
CLOUD_NAME = "hihbvdgg"
UPLOAD_PRESET = "yaguarete_preset"

def respaldar_trabajo_en_cloudinary(num_ot, ruta_archivo, fotos_subidas=None):
    try:
        url_doc = None
        urls_fotos = []

        if os.path.exists(ruta_archivo):
            with open(ruta_archivo, "rb") as file_to_upload:
                response = requests.post(
                    f"https://api.cloudinary.com/v1_1/{CLOUD_NAME}/raw/upload",
                    data={
                        "upload_preset": UPLOAD_PRESET,
                        "folder": f"Ordenes_de_Trabajo/OT_{num_ot}",
                        "public_id": f"OT_{num_ot}_Documento"
                    },
                    files={"file": file_to_upload}
                )
                if response.status_code == 200:
                    url_doc = response.json().get("secure_url")

        if fotos_subidas:
            for i, foto in enumerate(fotos_subidas, start=1):
                foto.seek(0)
                response_foto = requests.post(
                    f"https://api.cloudinary.com/v1_1/{CLOUD_NAME}/image/upload",
                    data={
                        "upload_preset": UPLOAD_PRESET,
                        "folder": f"Ordenes_de_Trabajo/OT_{num_ot}",
                        "public_id": f"Foto_{i}_{num_ot}"
                    },
                    files={"file": foto.getvalue()}
                )
                if response_foto.status_code == 200:
                    urls_fotos.append(response_foto.json().get("secure_url"))

        return url_doc, urls_fotos
    except Exception as e:
        st.error(f"Error al subir respaldo a Cloudinary: {e}")
        return None, []

# ==========================================
# 1. ESTILO VISUAL ABSOLUTO
# ==========================================
st.set_page_config(page_title="Yaguarete Papeles - Gestión OT", layout="wide", page_icon="📋")

st.markdown("""
    <style>
        /* Forzar fondo blanco general */
        html, body, .stApp, [data-testid="stAppViewContainer"], [data-testid="stHeader"] { 
            background-color: #FFFFFF !important; 
            color: #1E232A !important; 
        }

        label, p, span, div, h1, h2, h3, .stMarkdown { 
            color: #1E232A !important; 
            font-weight: 600 !important; 
        }

        /* Inputs, Textareas y Selectboxes */
        input, select, textarea, div[role="combobox"], [data-baseweb="select"] { 
            background-color: #FFFFFF !important; 
            color: #1E232A !important; 
            border: 1px solid #D5D8DC !important; 
            border-radius: 6px !important; 
        }

        /* POPUP / DESPLEGABLE DE BASEWEB */
        div[data-baseweb="popover"], div[data-baseweb="menu"], ul[role="listbox"], [data-baseweb="popover"] * {
            background-color: #FFFFFF !important;
            color: #1E232A !important;
        }

        /* Opciones individuales dentro del desplegable */
        li[role="option"], [data-baseweb="option"], div[role="option"] {
            background-color: #FFFFFF !important;
            color: #1E232A !important;
        }

        /* Hover sobre opciones */
        li[role="option"]:hover, [data-baseweb="option"]:hover, div[role="option"]:hover {
            background-color: #F2F4F4 !important;
            color: #A61C1C !important;
        }

        /* Botones Oscuros y Descarga */
        div.stButton > button, div.stDownloadButton > button {
            background-color: #A61C1C !important;
            color: #FFFFFF !important;
            border-radius: 6px !important;
            border: none !important;
            padding: 10px 20px !important;
            font-weight: 700 !important;
        }

        div.stDownloadButton > button * {
            color: #FFFFFF !important;
        }

        /* File Uploader */
        [data-testid="stFileUploader"] section {
            background-color: #F8F9F9 !important;
            border: 2px dashed #A61C1C !important;
        }

        [data-testid="stSidebar"] { 
            background-color: #F4F6F6 !important; 
            border-right: 2px solid #E5E8E8 !important; 
        }
    </style>
""", unsafe_allow_html=True)

EXCEL_FILE = "registro_ordenes_servicio.xlsx"
PLANTILLA_FILE = "plantilla_ot.docx"

AREAS = ["Papelote", "Caldera", "Expedición", "Químicos", "Mecánicos", "Km4"]
TECNICOS_OPCIONES = ["Ivan Sosa", "Néstor Medina", "Gerardo Maidana", "Cristian Alvarenga", "Otro (Especificar)"]
CAUSAS_OPCIONES = ["Desgaste natural", "Falta de lubricación", "Error operacional / manipulación", "Sobrecalentamiento", "Fuga hidráulica/neumática", "Falla eléctrica/cortocircuito", "Atascamiento / Muestra atascada", "Falta de mantenimiento preventivo", "Pieza defectuosa", "Llanta"]
MOTIVOS_PENDIENTE_OPCIONES = ["Cambio de turno", "Falta de repuestos / insumos", "Falta de herramientas especializadas", "Priorización de otra urgencia", "Espera de enfriamiento / parada de máquina", "Otro motivo (especificar)"]

MAQUINAS_DICT = {
    "Cat 5": "Cat 5", "Cat 7 (topadora)": "Cat 7", "Cat 8": "Cat 8", "Cat 9": "Cat 9", "Cat 10": "Cat 10", "Cat 11": "Cat 11",
    "Linde 3": "Linde 3", "Linde 7": "Linde 7", "Linde 8": "Linde 8", "Linde 9": "Linde 9", "Linde 10": "Linde 10", "Linde 11": "Linde 11", "Linde 12": "Linde 12",
    "Liugong 3": "Liugong 3", "Liugong 4": "Liugong 4", "Liugong 6": "Liugong 6", "Liugong 7": "Liugong 7", "Liugong 8": "Liugong 8",
    "Clark 2": "Clark 2", "Clark 3": "Clark 3", "Clark 5": "Clark 5", "Clark 6": "Clark 6", "Hyundai": "Hyundai", "Alquilada (especificar)": "ALQUILADA"
}

GEMINI_API_KEY = st.secrets.get("GEMINI_API_KEY", "")
if GEMINI_API_KEY:
    genai.configure(api_key=GEMINI_API_KEY)

# ==========================================
# 2. FUNCIONES Y MANEJO DE ARCHIVOS
# ==========================================
columnas_excel = [
    "Num_OT", "Fecha_Registro", "Estado", "Area", "Codigo_Maq", "Maquina", "Horometro",
    "Tecnico_Inicial", "Tecnico_Final", "Descripcion", "Tipo_Mantenimiento", "Horas_Mantenimiento", "Prioridad", 
    "Causa_Falla", "Categoria_Falla_AI", "Motivo_Pendiente", "Materiales", "Insumo_Cantidad",
    "Fecha_Inicial", "Hora_Final", "Fecha_Entrega", "Observaciones"
]

if not os.path.exists(EXCEL_FILE):
    df_init = pd.DataFrame(columns=columnas_excel)
    df_init.to_excel(EXCEL_FILE, index=False)

def obtener_siguiente_ot():
    """Calcula estrictamente el siguiente correlativo buscando el número máximo guardado en la columna Num_OT."""
    if os.path.exists(EXCEL_FILE):
        try:
            df_ot = pd.read_excel(EXCEL_FILE, usecols=[0])
            if not df_ot.empty:
                numeros = df_ot.iloc[:, 0].astype(str).str.extract(r'(\d+)')[0].dropna().astype(int)
                if not numeros.empty:
                    siguiente = numeros.max() + 1
                    return f"OT-{siguiente:05d}"
        except Exception:
            pass
    return "OT-00001"

def convertir_docx_a_pdf(ruta_docx, ruta_pdf):
    """Genera PDF desde DOCX."""
    try:
        from docx2pdf import convert
        convert(ruta_docx, ruta_pdf)
        return True
    except Exception:
        try:
            subprocess.run(["soffice", "--headless", "--convert-to", "pdf", ruta_docx], check=True)
            return True
        except Exception:
            return False

def rellenar_plantilla(datos_dict, fotos_subidas, ruta_salida_docx):
    """Rellena la plantilla Word manteniendo diseño original."""
    if os.path.exists(PLANTILLA_FILE):
        doc = DocxTemplate(PLANTILLA_FILE)
        imagenes_inline = []
        if fotos_subidas:
            for i, foto in enumerate(fotos_subidas):
                foto.seek(0)
                temp_img_path = f"temp_inline_img_{i}.png"
                with open(temp_img_path, "wb") as f_temp:
                    f_temp.write(foto.read())
                imagenes_inline.append(InlineImage(doc, temp_img_path, width=Cm(12)))

        contexto = {
            'area': datos_dict.get("area", ""),
            'códigomaq': datos_dict.get("códigomaq", ""),
            'Maquina': datos_dict.get("Maquina", ""),
            'horometro': datos_dict.get("horometro", ""),
            'tecnico': datos_dict.get("tecnico", ""),
            'numOT': datos_dict.get("numOT", ""),
            'descripcion_del_servicio': datos_dict.get("descripcion_del_servicio", ""),
            'tipo_mantenimiento': datos_dict.get("tipo_mantenimiento", ""),
            'prioridad': datos_dict.get("prioridad", ""),
            'causa_falla': datos_dict.get("causa_falla", ""),
            'Materiales': datos_dict.get("Materiales", ""),
            'fecha_inicial': datos_dict.get("fecha_inicial", ""),
            'hora_final': datos_dict.get("hora_final", ""),
            'fecha_de_entrega': datos_dict.get("fecha_de_entrega", ""),
            'observaciones': datos_dict.get("observaciones", ""),
            'fotos': imagenes_inline if imagenes_inline else []
        }
        
        doc.render(contexto)
        doc.save(ruta_salida_docx)

        if fotos_subidas:
            for i in range(len(fotos_subidas)):
                temp_img_path = f"temp_inline_img_{i}.png"
                if os.path.exists(temp_img_path):
                    os.remove(temp_img_path)

# ==========================================
# 3. NAVEGACIÓN PRINCIPAL
# ==========================================
st.sidebar.markdown("<h2 style='color: #A61C1C; text-align: center;'>YAGUARETE PAPELES</h2>", unsafe_allow_html=True)
st.sidebar.markdown("---")

opcion = st.sidebar.radio(
    "Navegación principal:",
    [
        "📋 Cargar Orden de Servicio", 
        "⏳ Trabajos Pendientes", 
        "📊 Panel de Estadísticas",
        "📂 Historial PDF"
    ]
)

if opcion == "📋 Cargar Orden de Servicio":
    st.markdown("<h1 style='color: #A61C1C;'>📋 Registro de Orden de Servicio</h1>", unsafe_allow_html=True)
    st.markdown("---")

    # Mantenimiento de Estado de OT en Session
    if "current_ot" not in st.session_state:
        st.session_state["current_ot"] = obtener_siguiente_ot()

    with st.form("form_ot"):
        col1, col2 = st.columns(2)
        with col1:
            num_ot = st.text_input("Número de OT", value=st.session_state["current_ot"], disabled=True)
            estado_ot = st.selectbox("Estado del Trabajo *", ["FINALIZADO", "PENDIENTE / A CONTINUAR"])
            area = st.selectbox("Área *", options=["-- Seleccionar --"] + AREAS)
            maquina_seleccionada = st.selectbox("Equipo o Máquina *", options=["-- Seleccionar --"] + list(MAQUINAS_DICT.keys()))
            
            if maquina_seleccionada == "Alquilada (especificar)":
                maquina_alquilada_detalle = st.text_input("Especifique Marca / Modelo:")
                maquina_final = f"Alquilada ({maquina_alquilada_detalle})" if maquina_alquilada_detalle else "Alquilada"
                codigo_maq = "ALQUILADA"
            elif maquina_seleccionada != "-- Seleccionar --":
                maquina_final = maquina_seleccionada
                codigo_maq = MAQUINAS_DICT[maquina_seleccionada]
            else:
                maquina_final = ""
                codigo_maq = ""

            st.text_input("Código Maq/Eq", value=codigo_maq, disabled=True)
            horometro = st.number_input("Horómetro", min_value=0.0, step=0.1)

            tecnico_seleccionado = st.selectbox("Nombre del Técnico *", options=["-- Seleccionar --"] + TECNICOS_OPCIONES)
            tecnico_final = st.text_input("Especifique el nombre del Técnico:") if tecnico_seleccionado == "Otro (Especificar)" else (tecnico_seleccionado if tecnico_seleccionado != "-- Seleccionar --" else "")

        with col2:
            tipo_mantenimiento = st.selectbox("Tipo Mantenimiento", ["CORRECTIVO", "PREVENTIVO", "PREDICTIVO"])
            horas_mantenimiento = st.selectbox("⚙️ Mantenimiento Periódico:", ["No Aplica", "250 hs", "500 hs", "1000 hs"]) if tipo_mantenimiento in ["PREVENTIVO", "PREDICTIVO"] else "N/A"
            prioridad = st.selectbox("Prioridad *", options=["-- Seleccionar --", "ALTA", "MEDIA", "BAJA"])
            fecha_inicial = st.date_input("Fecha Inicial")
            hora_final = st.time_input("Hora Final")
            fecha_entrega = st.date_input("Fecha de Entrega")

        st.subheader("Detalles del Servicio")
        descripcion_del_servicio = st.text_area("Descripción del Servicio Realizado / Diagnóstico")
        
        motivo_pendiente = ""
        if estado_ot == "PENDIENTE / A CONTINUAR":
            motivo_opcion = st.selectbox("Categoría del Motivo *", ["-- Seleccionar --"] + MOTIVOS_PENDIENTE_OPCIONES)
            motivo_detalle = st.text_area("Explicación Libre del Pendiente *")
            motivo_pendiente = f"[{motivo_opcion}] {motivo_detalle}".strip()

        causas_seleccionadas = st.multiselect("Causas estándar:", options=CAUSAS_OPCIONES)
        causa_detalle_extra = st.text_area("Detalles adicionales de causa:")

        st.subheader("Insumos y Repuestos")
        materiales = st.text_area("Materiales / Repuestos Utilizados")
        insumos_unidades = st.number_input("Cantidad Insumos Usados", min_value=0, value=1)
        observaciones = st.text_area("Observaciones Generales")

        st.subheader("📷 Fotos Adjuntas")
        fotos_subidas = st.file_uploader("Adjuntar evidencia fotográfica:", type=["png", "jpg", "jpeg"], accept_multiple_files=True)

        submitted = st.form_submit_button("💾 Guardar y Registrar Orden de Servicio")

    if submitted:
        if area == "-- Seleccionar --" or maquina_seleccionada == "-- Seleccionar --" or not tecnico_final or prioridad == "-- Seleccionar --":
            st.error("⚠️ Complete todos los campos obligatorios.")
            st.stop()

        partes_causa = []
        if horas_mantenimiento not in ["No Aplica", "N/A"]: partes_causa.append(f"Mantenimiento {horas_mantenimiento}")
        if causas_seleccionadas: partes_causa.append(", ".join(causas_seleccionadas))
        if causa_detalle_extra.strip(): partes_causa.append(causa_detalle_extra.strip())
        causa_falla_final = " - ".join(partes_causa) if partes_causa else "N/A"

        fecha_actual_str = datetime.now().strftime("%Y-%m-%d")
        nombre_base_trabajo = f"{tecnico_final}_{fecha_actual_str}_{maquina_final}_{codigo_maq}_{num_ot}"

        datos_docx = {
            "area": area,
            "códigomaq": codigo_maq,
            "Maquina": maquina_final,
            "horometro": horometro,
            "tecnico": tecnico_final,
            "numOT": num_ot,
            "descripcion_del_servicio": f"[{estado_ot}] {descripcion_del_servicio}",
            "tipo_mantenimiento": f"{tipo_mantenimiento} ({horas_mantenimiento})" if horas_mantenimiento != "N/A" else tipo_mantenimiento,
            "prioridad": prioridad,
            "causa_falla": causa_falla_final,
            "Materiales": materiales,
            "fecha_inicial": str(fecha_inicial),
            "hora_final": str(hora_final),
            "fecha_de_entrega": str(fecha_entrega),
            "observaciones": f"{observaciones} | Pendiente: {motivo_pendiente}" if motivo_pendiente else observaciones
        }

        ruta_salida_docx = f"{nombre_base_trabajo}.docx"
        ruta_salida_pdf = f"{nombre_base_trabajo}.pdf"

        # Generar Documento en Plantilla Word
        rellenar_plantilla(datos_docx, fotos_subidas, ruta_salida_docx)

        # Convertir a PDF
        se_convertio_pdf = convertir_docx_a_pdf(ruta_salida_docx, ruta_salida_pdf)

        # Guardar en Excel
        df_existente = pd.read_excel(EXCEL_FILE)
        nueva_fila = {
            "Num_OT": num_ot, "Fecha_Registro": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "Estado": "PENDIENTE" if estado_ot == "PENDIENTE / A CONTINUAR" else "FINALIZADO",
            "Area": area, "Codigo_Maq": codigo_maq, "Maquina": maquina_final, "Horometro": horometro,
            "Tecnico_Inicial": tecnico_final, "Tecnico_Final": tecnico_final if estado_ot != "PENDIENTE / A CONTINUAR" else "",
            "Descripcion": descripcion_del_servicio, "Tipo_Mantenimiento": tipo_mantenimiento, "Horas_Mantenimiento": horas_mantenimiento,
            "Prioridad": prioridad, "Causa_Falla": causa_falla_final, "Categoria_Falla_AI": "Mecánica", 
            "Motivo_Pendiente": motivo_pendiente, "Materiales": materiales, "Insumo_Cantidad": insumos_unidades, 
            "Fecha_Inicial": str(fecha_inicial), "Hora_Final": str(hora_final), "Fecha_Entrega": str(fecha_entrega), "Observaciones": observaciones
        }
        
        df_actualizado = pd.concat([df_existente, pd.DataFrame([nueva_fila])], ignore_index=True)
        df_actualizado.to_excel(EXCEL_FILE, index=False)

        # Subir Respaldo
        archivo_para_respaldo = ruta_salida_pdf if se_convertio_pdf else ruta_salida_docx
        respaldar_trabajo_en_cloudinary(num_ot, archivo_para_respaldo, fotos_subidas)

        # Actualizar número para la siguiente OT
        st.session_state["current_ot"] = obtener_siguiente_ot()

        st.success(f"✅ Orden {num_ot} registrada exitosamente.")

        # Botones de descarga directos
        col_btn1, col_btn2 = st.columns(2)
        with col_btn1:
            if se_convertio_pdf and os.path.exists(ruta_salida_pdf):
                with open(ruta_salida_pdf, "rb") as file_pdf:
                    st.download_button("📥 Descargar Orden Oficial (.pdf)", data=file_pdf, file_name=ruta_salida_pdf, mime="application/pdf")
            elif os.path.exists(ruta_salida_docx):
                with open(ruta_salida_docx, "rb") as file_docx:
                    st.download_button("📥 Descargar Orden Oficial (.docx)", data=file_docx, file_name=ruta_salida_docx)

        with col_btn2:
            if os.path.exists(ruta_salida_docx):
                with open(ruta_salida_docx, "rb") as file_docx:
                    st.download_button("📄 Descargar Formato Editable (.docx)", data=file_docx, file_name=ruta_salida_docx)

elif opcion == "⏳ Trabajos Pendientes":
    st.markdown("<h1 style='color: #A61C1C;'>⏳ Gestor de Trabajos Pendientes y Finalizados</h1>", unsafe_allow_html=True)
    st.markdown("---")

    df = pd.read_excel(EXCEL_FILE) if os.path.exists(EXCEL_FILE) else pd.DataFrame(columns=columnas_excel)

    tab1, tab2, tab3 = st.tabs(["✍️ Completar Trabajo", "📋 Pendientes Activos", "✅ Historial Completo"])

    with tab1:
        df_pend = df[df["Estado"] == "PENDIENTE"] if "Estado" in df.columns and not df.empty else pd.DataFrame()
        if not df_pend.empty:
            ot_seleccionada = st.selectbox("Seleccione la Orden a Finalizar:", df_pend["Num_OT"].dropna().tolist())
            row_ot = df_pend[df_pend["Num_OT"] == ot_seleccionada].iloc[0]
            
            st.info(f"📌 OT: {ot_seleccionada} | Equipo: {row_ot.get('Maquina')} | Técnico Creador: {row_ot.get('Tecnico_Inicial')}")
            
            with st.form("form_completar_pendiente"):
                tec_final_opc = st.selectbox("Técnico que Finaliza *", ["-- Seleccionar --"] + TECNICOS_OPCIONES)
                horometro_f = st.number_input("Horómetro Final", min_value=0.0, value=float(row_ot.get("Horometro", 0.0)), step=0.1)
                desc_final = st.text_area("Descripción del Trabajo Concluido *")
                materiales_f = st.text_area("Materiales Utilizados", value=str(row_ot.get("Materiales", "")))
                fotos_subidas_f = st.file_uploader("📷 Fotos Evidencia Final:", type=["png", "jpg", "jpeg"], accept_multiple_files=True)

                btn_finalizar = st.form_submit_button("🏁 Finalizar Trabajo y Generar Documento")

            if btn_finalizar:
                if tec_final_opc == "-- Seleccionar --" or not desc_final.strip():
                    st.error("⚠️ Complete los campos obligatorios.")
                    st.stop()

                idx_excel = df[df["Num_OT"] == ot_seleccionada].index[0]
                df.loc[idx_excel, "Estado"] = "FINALIZADO"
                df.loc[idx_excel, "Tecnico_Final"] = tec_final_opc
                df.loc[idx_excel, "Horometro"] = horometro_f
                df.loc[idx_excel, "Descripcion"] = f"{str(row_ot.get('Descripcion', ''))} | [FINALIZADO]: {desc_final}"
                df.loc[idx_excel, "Materiales"] = materiales_f
                df.to_excel(EXCEL_FILE, index=False)

                datos_docx = {
                    "area": row_ot.get("Area", ""),
                    "códigomaq": row_ot.get("Codigo_Maq", ""),
                    "Maquina": row_ot.get("Maquina", ""),
                    "horometro": horometro_f,
                    "tecnico": f"{row_ot.get('Tecnico_Inicial', '')} / {tec_final_opc}",
                    "numOT": ot_seleccionada,
                    "descripcion_del_servicio": f"[FINALIZADO] {desc_final}",
                    "tipo_mantenimiento": row_ot.get("Tipo_Mantenimiento", "CORRECTIVO"),
                    "prioridad": row_ot.get("Prioridad", "MEDIA"),
                    "causa_falla": row_ot.get("Causa_Falla", "N/A"),
                    "Materiales": materiales_f,
                    "fecha_inicial": str(row_ot.get("Fecha_Inicial", "")),
                    "hora_final": str(datetime.now().strftime("%H:%M:%S")),
                    "fecha_de_entrega": str(datetime.now().strftime("%Y-%m-%d")),
                    "observaciones": str(row_ot.get("Observaciones", ""))
                }

                nombre_base = f"{tec_final_opc}_{datetime.now().strftime('%Y-%m-%d')}_{row_ot.get('Maquina','')}_{ot_seleccionada}"
                ruta_salida_docx = f"{nombre_base}.docx"
                ruta_salida_pdf = f"{nombre_base}.pdf"

                rellenar_plantilla(datos_docx, fotos_subidas_f, ruta_salida_docx)
                se_convertio_pdf = convertir_docx_a_pdf(ruta_salida_docx, ruta_salida_pdf)

                archivo_respaldo = ruta_salida_pdf if se_convertio_pdf else ruta_salida_docx
                respaldar_trabajo_en_cloudinary(ot_seleccionada, archivo_respaldo, fotos_subidas_f)

                st.success(f"🎉 Orden {ot_seleccionada} completada exitosamente.")
                
                if se_convertio_pdf and os.path.exists(ruta_salida_pdf):
                    with open(ruta_salida_pdf, "rb") as file_pdf:
                        st.download_button("📥 Descargar Documento Final (.pdf)", data=file_pdf, file_name=ruta_salida_pdf, mime="application/pdf")
                elif os.path.exists(ruta_salida_docx):
                    with open(ruta_salida_docx, "rb") as file_docx:
                        st.download_button("💾 Descargar Documento Final (.docx)", data=file_docx, file_name=ruta_salida_docx)
        else:
            st.info("ℹ️ No hay trabajos pendientes.")

    with tab2:
        df_pendientes = df[df["Estado"] == "PENDIENTE"] if "Estado" in df.columns and not df.empty else pd.DataFrame()
        st.dataframe(df_pendientes, use_container_width=True)

    with tab3:
        df_completados = df[df["Estado"] == "FINALIZADO"] if "Estado" in df.columns and not df.empty else pd.DataFrame()
        st.dataframe(df_completados, use_container_width=True)

elif opcion == "📊 Panel de Estadísticas":
    st.markdown("<h1 style='color: #A61C1C;'>📊 Panel de Estadísticas</h1>", unsafe_allow_html=True)
    if os.path.exists(EXCEL_FILE):
        df = pd.read_excel(EXCEL_FILE)
        if not df.empty:
            c1, c2, c3 = st.columns(3)
            c1.metric("Total Órdenes", len(df))
            c2.metric("Pendientes", len(df[df["Estado"] == "PENDIENTE"]))
            c3.metric("Finalizadas", len(df[df["Estado"] == "FINALIZADO"]))
            st.dataframe(df, use_container_width=True)

elif opcion == "📂 Historial PDF":
    st.markdown("<h1 style='color: #A61C1C;'>📂 Historial PDF y Documentos Generados</h1>", unsafe_allow_html=True)
    st.markdown("---")

    # 1. Buscar primero archivos en la raíz del servidor local
    archivos_locales = [f for f in os.listdir(".") if (f.endswith(".pdf") or f.endswith(".docx")) and f != PLANTILLA_FILE]

    # 2. Si no hay locales o se prefiere consultar el registro, leer desde el Excel
    if os.path.exists(EXCEL_FILE):
        df_historial = pd.read_excel(EXCEL_FILE)
    else:
        df_historial = pd.DataFrame()

    st.subheader("Archivos Disponibles")

    # Si se encontraron archivos físicos en la carpeta local
    if archivos_locales:
        busqueda = st.text_input("🔍 Buscar por Número de OT, Equipo o Técnico:", "")
        if busqueda:
            archivos_locales = [doc for doc in archivos_locales if busqueda.lower() in doc.lower()]

        for archivo in sorted(archivos_locales, reverse=True):
            col_nombre, col_btn = st.columns([3, 1])
            with col_nombre:
                st.write(f"📄 **{archivo}**")
            with col_btn:
                with open(archivo, "rb") as f:
                    mime_type = "application/pdf" if archivo.endswith(".pdf") else "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    st.download_button(
                        label="📥 Descargar",
                        data=f,
                        file_name=archivo,
                        mime=mime_type,
                        key=f"dl_hist_{archivo}"
                    )
            st.markdown("---")

    # Si la app está en la nube (Streamlit Cloud) y los archivos locales no persisten, mostramos el enlace/registro del Excel
    elif not df_historial.empty:
        st.info("ℹ️ Mostrando registro de órdenes desde la base de datos (Excel).")
        busqueda = st.text_input("🔍 Buscar OT en el historial:", "")
        
        if busqueda:
            df_historial = df_historial[df_historial.astype(str).apply(lambda x: x.str.contains(busqueda, case=False)).any(axis=1)]

        for _, fila in df_historial.iterrows():
            num_ot_rec = fila.get("Num_OT", "N/A")
            tec_rec = fila.get("Tecnico_Inicial", "N/A")
            maq_rec = fila.get("Maquina", "N/A")
            fecha_rec = fila.get("Fecha_Registro", "N/A")

            # Intentar re-armar el nombre del archivo si existe localmente
            nombre_esperado = f"{tec_rec}_{str(fecha_rec)[:10]}_{maq_rec}_{num_ot_rec}.pdf"
            
            col1, col2 = st.columns([3, 1])
            with col1:
                st.write(f"📋 **OT:** {num_ot_rec} | **Equipo:** {maq_rec} | **Técnico:** {tec_rec} | **Fecha:** {fecha_rec}")
            with col2:
                if os.path.exists(nombre_esperado):
                    with open(nombre_esperado, "rb") as f:
                        st.download_button("📥 Descargar PDF", data=f, file_name=nombre_esperado, mime="application/pdf", key=f"btn_ex_{num_ot_rec}")
                else:
                    # Enlace directo de respaldo a Cloudinary si está configurado
                    url_cloudinary = f"https://res.cloudinary.com/{CLOUD_NAME}/raw/upload/Ordenes_de_Trabajo/OT_{num_ot_rec}/OT_{num_ot_rec}_Documento"
                    st.markdown(f"[🔗 Ver/Descargar Respaldo]({url_cloudinary})")
            st.markdown("---")
    else:
        st.warning("⚠️ No se encontraron documentos locales ni registros en la base de datos. Genere una nueva OT para visualizarla aquí.")
