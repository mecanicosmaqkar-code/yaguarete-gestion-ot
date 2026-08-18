import os
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
# 1. CONFIGURACIÓN Y ESTILO VISUAL
# ==========================================
st.set_page_config(page_title="Yaguarete Papeles - Gestión OT", layout="wide", page_icon="📋")

st.markdown("""
    <style>
        .stApp, [data-testid="stAppViewContainer"] { background-color: #FFFFFF !important; color: #1E232A !important; }
        label, p, span, div, .stMarkdown { color: #1E232A !important; font-weight: 600 !important; }
        input, select, textarea, div[role="combobox"] { background-color: #F8F9F9 !important; color: #1E232A !important; border: 1px solid #D5D8DC !important; border-radius: 6px !important; }
        [data-testid="stSidebar"] { background-color: #F4F6F6 !important; border-right: 2px solid #E5E8E8 !important; }
        h1, h2, h3, .stHeader { color: #A61C1C !important; font-weight: 700 !important; }
        div.stButton > button:first-child { background-color: #A61C1C !important; color: #FFFFFF !important; border-radius: 6px !important; border: none !important; padding: 10px 20px !important; font-weight: 700 !important; }
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
# 2. FUNCIONES Y BASE DE DATOS
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
    """Calcula dinámicamente el número consecutivo de OT."""
    if os.path.exists(EXCEL_FILE):
        try:
            df = pd.read_excel(EXCEL_FILE)
            if not df.empty and "Num_OT" in df.columns:
                numeros = []
                for ot in df["Num_OT"].dropna().astype(str):
                    digits = ''.join(filter(str.isdigit, ot))
                    if digits:
                        numeros.append(int(digits))
                if numeros:
                    return f"OT-{max(numeros) + 1:05d}"
        except Exception:
            pass
    return "OT-00001"

def analizar_causa_con_gemini(causa_texto):
    if not GEMINI_API_KEY:
        return "Mecánica"
    try:
        model = genai.GenerativeModel('gemini-1.5-flash')
        prompt = f"Clasifica esta falla técnica en UNA sola categoría (Mecánica, Eléctrica, Hidráulica, Neumática, Error Operacional, Desgaste Natural, Llantas): '{causa_texto}'"
        response = model.generate_content(prompt)
        return response.text.strip()
    except Exception:
        return "Mecánica"

def rellenar_plantilla(datos_dict, fotos_subidas, ruta_salida_docx):
    """
    Rellena directamente los campos Jinja2 de tu plantilla 'plantilla_ot.docx'
    e incrusta las imágenes dentro del tag {{ fotos }} o al final de la plantilla.
    """
    if os.path.exists(PLANTILLA_FILE):
        doc = DocxTemplate(PLANTILLA_FILE)
        
        # Procesar fotos para plantilla
        imagenes_inline = []
        if fotos_subidas:
            for i, foto in enumerate(fotos_subidas):
                foto.seek(0)
                temp_img_path = f"temp_inline_img_{i}.png"
                with open(temp_img_path, "wb") as f_temp:
                    f_temp.write(foto.read())
                # Redimensionar la imagen a 12cm de ancho dentro de la plantilla
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
            'fotos': imagenes_inline if imagenes_inline else ""
        }
        
        doc.render(contexto)
        doc.save(ruta_salida_docx)

        # Si la plantilla no tenía la etiqueta {{ fotos }}, se agregan centradas al final
        if fotos_subidas and not imagenes_inline:
            doc_fotos = Document(ruta_salida_docx)
            doc_fotos.add_heading("Anexos Fotográficos", level=2)
            for foto in fotos_subidas:
                foto.seek(0)
                p = doc_fotos.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = p.add_run()
                run.add_picture(foto, width=Cm(12))
            doc_fotos.save(ruta_salida_docx)

        # Limpiar imágenes temporales
        if fotos_subidas:
            for i in range(len(fotos_subidas)):
                temp_img_path = f"temp_inline_img_{i}.png"
                if os.path.exists(temp_img_path):
                    os.remove(temp_img_path)
    else:
        doc = Document()
        doc.add_heading('YAGUARETE PAPELES - ORDEN DE SERVICIO', 0)
        for k, v in datos_dict.items():
            doc.add_paragraph(f"{k}: {v}")
        doc.save(ruta_salida_docx)

# ==========================================
# 3. NAVEGACIÓN Y MENÚ PRINCIPAL
# ==========================================
st.sidebar.markdown("<h2 style='color: #A61C1C; text-align: center;'>YAGUARETE PAPELES</h2>", unsafe_allow_html=True)
st.sidebar.markdown("---")

opcion = st.sidebar.radio(
    "Navegación principal:",
    ["📋 Cargar Orden de Servicio", "⏳ Trabajos Pendientes", "📊 Panel de Estadísticas"]
)

# ==========================================
# SECCIÓN 1: CARGAR ORDEN DE SERVICIO
# ==========================================
if opcion == "📋 Cargar Orden de Servicio":
    st.markdown("<h1 style='color: #A61C1C;'>📋 Registro de Orden de Servicio</h1>", unsafe_allow_html=True)
    st.markdown("---")

    ot_sugerida = obtener_siguiente_ot()

    with st.form("form_ot"):
        col1, col2 = st.columns(2)
        with col1:
            num_ot = st.text_input("Número de OT", value=ot_sugerida, disabled=True)
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

        categoria_ai = analizar_causa_con_gemini(causa_falla_final)
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

        # Generación basada estrictamente en tu plantilla .docx
        rellenar_plantilla(datos_docx, fotos_subidas, ruta_salida_docx)

        df_existente = pd.read_excel(EXCEL_FILE)
        nueva_fila = {
            "Num_OT": num_ot, "Fecha_Registro": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "Estado": "PENDIENTE" if estado_ot == "PENDIENTE / A CONTINUAR" else "FINALIZADO",
            "Area": area, "Codigo_Maq": codigo_maq, "Maquina": maquina_final, "Horometro": horometro,
            "Tecnico_Inicial": tecnico_final, "Tecnico_Final": tecnico_final if estado_ot != "PENDIENTE / A CONTINUAR" else "",
            "Descripcion": descripcion_del_servicio, "Tipo_Mantenimiento": tipo_mantenimiento, "Horas_Mantenimiento": horas_mantenimiento,
            "Prioridad": prioridad, "Causa_Falla": causa_falla_final, "Categoria_Falla_AI": categoria_ai, 
            "Motivo_Pendiente": motivo_pendiente, "Materiales": materiales, "Insumo_Cantidad": insumos_unidades, 
            "Fecha_Inicial": str(fecha_inicial), "Hora_Final": str(hora_final), "Fecha_Entrega": str(fecha_entrega), "Observaciones": observaciones
        }
        
        df_actualizado = pd.concat([df_existente, pd.DataFrame([nueva_fila])], ignore_index=True)
        df_actualizado.to_excel(EXCEL_FILE, index=False)

        respaldar_trabajo_en_cloudinary(num_ot, ruta_salida_docx, fotos_subidas)

        st.success(f"✅ Orden {num_ot} registrada y generada exitosamente con el formato de plantilla.")
        
        if os.path.exists(ruta_salida_docx):
            with open(ruta_salida_docx, "rb") as file_docx:
                st.download_button("📥 Descargar Orden Oficial (.docx)", data=file_docx, file_name=ruta_salida_docx)

# ==========================================
# SECCIÓN 2: TRABAJOS PENDIENTES
# ==========================================
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

                rellenar_plantilla(datos_docx, fotos_subidas_f, ruta_salida_docx)
                respaldar_trabajo_en_cloudinary(ot_seleccionada, ruta_salida_docx, fotos_subidas_f)

                st.success(f"🎉 Orden {ot_seleccionada} completada exitosamente.")
                if os.path.exists(ruta_salida_docx):
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

# ==========================================
# SECCIÓN 3: PANEL DE ESTADÍSTICAS
# ==========================================
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