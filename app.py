import os
from datetime import datetime
import pandas as pd
import openpyxl
import requests
import streamlit as st
import plotly.express as px
import google.generativeai as genai

from docx import Document
from docx.shared import Inches
from docxtpl import DocxTemplate

# ReportLab para generación directa de PDF en servidores Linux / Cloud
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch

# ==========================================
# CONFIGURACIÓN DE CLOUDINARY (UNSIGNED)
# ==========================================
CLOUD_NAME = "hihbvdgg"
UPLOAD_PRESET = "yaguarete_preset"

def respaldar_trabajo_en_cloudinary(num_ot, ruta_archivo, fotos_subidas=None):
    """
    Sube el documento generado (PDF/DOCX) y las fotos adjuntas a Cloudinary utilizando la API REST (Unsigned).
    """
    try:
        url_doc = None
        urls_fotos = []

        # 1. Subir el documento principal
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
                else:
                    st.warning(f"No se pudo subir el documento: {response.text}")

        # 2. Subir imágenes adjuntas si existen
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
        .stApp, [data-testid="stAppViewContainer"] {
            background-color: #FFFFFF !important;
            color: #1E232A !important;
        }

        label, p, span, div, .stMarkdown, .stSelectbox label, .stTextInput label, .stTextArea label, .stNumberInput label, .stMultiSelect label {
            color: #1E232A !important;
            font-weight: 600 !important;
        }

        input, select, textarea, div[role="combobox"], div[data-baseweb="select"] {
            background-color: #F8F9F9 !important;
            color: #1E232A !important;
            border: 1px solid #D5D8DC !important;
            border-radius: 6px !important;
        }

        div[data-baseweb="popover"], div[role="listbox"], ul[role="listbox"] li {
            background-color: #FFFFFF !important;
            color: #1E232A !important;
        }
        
        [data-testid="stSidebar"] {
            background-color: #F4F6F6 !important;
            border-right: 2px solid #E5E8E8 !important;
        }
        
        h1, h2, h3, .stHeader {
            color: #A61C1C !important;
            font-weight: 700 !important;
        }
        
        div.stButton > button:first-child {
            background-color: #A61C1C !important;
            color: #FFFFFF !important;
            border-radius: 6px !important;
            border: none !important;
            padding: 10px 20px !important;
            font-weight: 700 !important;
            font-size: 16px !important;
            transition: all 0.3s ease !important;
        }
        div.stButton > button:first-child:hover {
            background-color: #801414 !important;
            color: #FFFFFF !important;
            box-shadow: 0px 4px 10px rgba(166, 28, 28, 0.4) !important;
        }

        [data-testid="stMetricValue"] {
            color: #A61C1C !important;
            font-weight: bold !important;
        }

        .card-pendiente {
            background-color: #FDEDEC !important;
            border-left: 5px solid #A61C1C !important;
            padding: 15px !important;
            border-radius: 6px !important;
            margin-bottom: 15px !important;
            color: #1E232A !important;
        }
        .card-pendiente h3, .card-pendiente p, .card-pendiente b {
            color: #1E232A !important;
        }
        .card-pendiente h3 {
            color: #A61C1C !important;
        }

        .card-completado {
            background-color: #E8F8F5 !important;
            border-left: 5px solid #27AE60 !important;
            padding: 15px !important;
            border-radius: 6px !important;
            margin-bottom: 15px !important;
            color: #1E232A !important;
        }
        .card-completado h3, .card-completado p, .card-completado b {
            color: #1E232A !important;
        }
        .card-completado h3 {
            color: #27AE60 !important;
        }
    </style>
""", unsafe_allow_html=True)

EXCEL_FILE = "registro_ordenes_servicio.xlsx"
PLANTILLA_FILE = "plantilla_ot.docx"

AREAS = ["Papelote", "Caldera", "Expedición", "Químicos", "Mecánicos", "Km4"]

TECNICOS_OPCIONES = [
    "Ivan Sosa",
    "Néstor Medina",
    "Gerardo Maidana",
    "Cristian Alvarenga",
    "Otro (Especificar)"
]

CAUSAS_OPCIONES = [
    "Desgaste natural",
    "Falta de lubricación",
    "Error operacional / manipulación",
    "Sobrecalentamiento",
    "Fuga hidráulica/neumática",
    "Falla eléctrica/cortocircuito",
    "Atascamiento / Muestra atascada",
    "Falta de mantenimiento preventivo",
    "Pieza defectuosa",
    "Llanta"
]

MOTIVOS_PENDIENTE_OPCIONES = [
    "Cambio de turno",
    "Falta de repuestos / insumos",
    "Falta de herramientas especializadas",
    "Priorización de otra urgencia",
    "Espera de enfriamiento / parada de máquina",
    "Otro motivo (especificar)"
]

MAQUINAS_DICT = {
    "Cat 5": "Cat 5", "Cat 7 (topadora)": "Cat 7", "Cat 8": "Cat 8", "Cat 9": "Cat 9", "Cat 10": "Cat 10", "Cat 11": "Cat 11",
    "Linde 3": "Linde 3", "Linde 7": "Linde 7", "Linde 8": "Linde 8", "Linde 9": "Linde 9", "Linde 10": "Linde 10", "Linde 11": "Linde 11", "Linde 12": "Linde 12",
    "Liugong 3": "Liugong 3", "Liugong 4": "Liugong 4", "Liugong 6": "Liugong 6", "Liugong 7": "Liugong 7", "Liugong 8": "Liugong 8",
    "Clark 2": "Clark 2", "Clark 3": "Clark 3", "Clark 5": "Clark 5", "Clark 6": "Clark 6",
    "Hyundai": "Hyundai",
    "Alquilada (especificar)": "ALQUILADA"
}

GEMINI_API_KEY = st.secrets.get("GEMINI_API_KEY", "")

if GEMINI_API_KEY:
    genai.configure(api_key=GEMINI_API_KEY)

# ==========================================
# 2. BASE DE DATOS Y FUNCIONES
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
    if os.path.exists(EXCEL_FILE):
        try:
            df = pd.read_excel(EXCEL_FILE)
            if not df.empty and "Num_OT" in df.columns:
                ultimas_ots = df["Num_OT"].dropna().astype(str).tolist()
                numeros = []
                for ot in ultimas_ots:
                    digits = ''.join(filter(str.isdigit, ot))
                    if digits:
                        numeros.append(int(digits))
                if numeros:
                    siguiente_num = max(numeros) + 1
                    return f"OT-{siguiente_num:05d}"
        except Exception:
            pass
    return "OT-00001"

def analizar_causa_con_gemini(causa_texto):
    if not GEMINI_API_KEY:
        return "Mecánica"
    try:
        model = genai.GenerativeModel('gemini-1.5-flash')
        prompt = f"""
        Analiza la siguiente Causa de Falla escrita por un técnico en una Orden de Trabajo de Yaguarete Papeles.
        Clasifícala strictly en UNA de estas categorías:
        - Mecánica
        - Eléctrica
        - Hidráulica
        - Neumática
        - Error Operacional / Humano
        - Desgaste Natural
        - Calibración / Instrumentación
        - Llantas / Neumáticos

        Causa: '{causa_texto}'
        Responde únicamente con el nombre de la categoría elegida.
        """
        response = model.generate_content(prompt)
        return response.text.strip()
    except Exception:
        return "Mecánica"

def rellenar_plantilla(datos_dict, fotos_subidas, ruta_salida_docx):
    """
    Rellena la plantilla Word conservando el diseño original.
    """
    if os.path.exists(PLANTILLA_FILE):
        doc = DocxTemplate(PLANTILLA_FILE)
        contexto = {
            'area': datos_dict.get("<<area>>", ""),
            'codigo_maq': datos_dict.get("<<códigomaq>>", ""),
            'maquina': datos_dict.get("<<Maquina>>", ""),
            'horometro': datos_dict.get("<<horometro>>", ""),
            'tecnico': datos_dict.get("<<tecnico>>", ""),
            'num_ot': datos_dict.get("<<numOT>>", ""),
            'descripcion': datos_dict.get("<<descripcion_del_servicio>>", ""),
            'tipo_mantenimiento': datos_dict.get("<<tipo_mantenimiento>>", ""),
            'prioridad': datos_dict.get("<<prioridad>>", ""),
            'causa_falla': datos_dict.get("<<causa_falla>>", ""),
            'materiales': datos_dict.get("<<Materiales>>", ""),
            'fecha_inicial': datos_dict.get("<<fecha_inicial>>", ""),
            'hora_final': datos_dict.get("<<hora_final>>", ""),
            'fecha_entrega': datos_dict.get("<<fecha_de_entrega>>", ""),
            'observaciones': datos_dict.get("<<observaciones>>", "")
        }
        doc.render(contexto)
        doc.save(ruta_salida_docx)

        if fotos_subidas:
            doc_fotos = Document(ruta_salida_docx)
            doc_fotos.add_heading("Fotos Adjuntas", level=2)
            for foto in fotos_subidas:
                foto.seek(0)
                doc_fotos.add_picture(foto, width=Inches(4.5))
            doc_fotos.save(ruta_salida_docx)
    else:
        doc = Document()
        doc.add_heading('YAGUARETE PAPELES - ORDEN DE SERVICIO', 0)
        for k, v in datos_dict.items():
            doc.add_paragraph(f"{k}: {v}")
        if fotos_subidas:
            doc.add_heading("Fotos Adjuntas", level=2)
            for foto in fotos_subidas:
                foto.seek(0)
                doc.add_picture(foto, width=Inches(4.5))
        doc.save(ruta_salida_docx)

def generar_pdf_reportlab(datos_dict, fotos_subidas, ruta_pdf):
    """
    Genera un PDF profesional utilizando ReportLab, totalmente formateado e incluyendo imágenes.
    """
    try:
        doc = SimpleDocTemplate(
            ruta_pdf,
            pagesize=letter,
            rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36
        )
        styles = getSampleStyleSheet()
        
        title_style = ParagraphStyle(
            'TitleStyle',
            parent=styles['Heading1'],
            fontSize=16,
            leading=20,
            textColor=colors.HexColor('#A61C1C'),
            alignment=1,
            spaceAfter=15
        )
        
        label_style = ParagraphStyle(
            'LabelStyle', parent=styles['Normal'], fontSize=9, leading=12, fontName='Helvetica-Bold', textColor=colors.HexColor('#1E232A')
        )
        value_style = ParagraphStyle(
            'ValueStyle', parent=styles['Normal'], fontSize=9, leading=12, fontName='Helvetica', textColor=colors.HexColor('#1E232A')
        )
        sec_style = ParagraphStyle(
            'SecStyle', parent=styles['Heading2'], fontSize=11, leading=14, textColor=colors.HexColor('#A61C1C'), spaceBefore=10, spaceAfter=4
        )

        elements = []
        
        # 1. Título
        elements.append(Paragraph("YAGUARETE PAPELES - ORDEN DE SERVICIO", title_style))

        # 2. Tabla Encabezado
        tabla_datos = [
            [Paragraph("N° OT:", label_style), Paragraph(str(datos_dict.get("<<numOT>>", "")), value_style), Paragraph("Área:", label_style), Paragraph(str(datos_dict.get("<<area>>", "")), value_style)],
            [Paragraph("Equipo / Máquina:", label_style), Paragraph(str(datos_dict.get("<<Maquina>>", "")), value_style), Paragraph("Código Máq:", label_style), Paragraph(str(datos_dict.get("<<códigomaq>>", "")), value_style)],
            [Paragraph("Horómetro:", label_style), Paragraph(str(datos_dict.get("<<horometro>>", "")), value_style), Paragraph("Técnico:", label_style), Paragraph(str(datos_dict.get("<<tecnico>>", "")), value_style)],
            [Paragraph("Tipo Mant.:", label_style), Paragraph(str(datos_dict.get("<<tipo_mantenimiento>>", "")), value_style), Paragraph("Prioridad:", label_style), Paragraph(str(datos_dict.get("<<prioridad>>", "")), value_style)],
            [Paragraph("Fecha Inicial:", label_style), Paragraph(str(datos_dict.get("<<fecha_inicial>>", "")), value_style), Paragraph("Fecha Entrega:", label_style), Paragraph(str(datos_dict.get("<<fecha_de_entrega>>", "")), value_style)],
            [Paragraph("Hora Final:", label_style), Paragraph(str(datos_dict.get("<<hora_final>>", "")), value_style), Paragraph("", label_style), Paragraph("", value_style)],
        ]

        t = Table(tabla_datos, colWidths=[110, 160, 110, 160])
        t.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#F8F9FA')),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#D1D5DB')),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('TOPPADDING', (0, 0), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ]))
        elements.append(t)
        elements.append(Spacer(1, 10))

        # 3. Bloques de Información
        secciones = [
            ("Descripción del Servicio", datos_dict.get("<<descripcion_del_servicio>>", "")),
            ("Causa de Falla / Trabajos Realizados", datos_dict.get("<<causa_falla>>", "")),
            ("Materiales / Repuestos Utilizados", datos_dict.get("<<Materiales>>", "")),
            ("Observaciones Generales", datos_dict.get("<<observaciones>>", ""))
        ]

        for titulo, contenido in secciones:
            elements.append(Paragraph(titulo, sec_style))
            elements.append(Paragraph(str(contenido) if contenido else "N/A", value_style))
            elements.append(Spacer(1, 6))

        # 4. Fotos Adjuntas al Final del PDF
        if fotos_subidas:
            elements.append(Spacer(1, 10))
            elements.append(Paragraph("Evidencia Fotográfica", sec_style))
            elements.append(Spacer(1, 4))
            
            for i, foto in enumerate(fotos_subidas):
                try:
                    foto.seek(0)
                    temp_img_path = f"temp_pdf_img_{i}.png"
                    with open(temp_img_path, "wb") as f_temp:
                        f_temp.write(foto.read())
                    
                    img = RLImage(temp_img_path)
                    max_width = 6.5 * inch
                    max_height = 3.5 * inch

                    aspect = img.imageWidth / float(img.imageHeight)
                    if aspect > 1:
                        img.drawWidth = min(max_width, img.imageWidth)
                        img.drawHeight = img.drawWidth / aspect
                    else:
                        img.drawHeight = min(max_height, img.imageHeight)
                        img.drawWidth = img.drawHeight * aspect

                    elements.append(img)
                    elements.append(Spacer(1, 8))
                except Exception as e_img:
                    elements.append(Paragraph(f"Error adjuntando foto: {e_img}", value_style))

        doc.build(elements)

        # Limpieza de archivos temporales de imagen creados para el PDF
        if fotos_subidas:
            for i in range(len(fotos_subidas)):
                temp_img_path = f"temp_pdf_img_{i}.png"
                if os.path.exists(temp_img_path):
                    os.remove(temp_img_path)

        return ruta_pdf
    except Exception as e:
        st.warning(f"Error generando PDF dinámico: {e}")
        return None

# ==========================================
# 3. NAVEGACIÓN Y MENÚ PRINCIPAL
# ==========================================
st.sidebar.markdown("<h2 style='color: #A61C1C; text-align: center; margin-bottom: 0px;'>YAGUARETE</h2>", unsafe_allow_html=True)
st.sidebar.markdown("<p style='text-align: center; font-weight: bold; color: #1E232A; margin-top: 0px;'>PAPELES</p>", unsafe_allow_html=True)
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
    st.write("Diligencie el formulario correspondiente al mantenimiento o reparación efectuada.")
    st.markdown("---")

    ot_sugerida = obtener_siguiente_ot()

    with st.form("form_ot"):
        col1, col2 = st.columns(2)
        
        with col1:
            num_ot = st.text_input("Número de OT", value=ot_sugerida, disabled=True)
            
            estado_ot = st.selectbox("Estado del Trabajo *", ["FINALIZADO", "PENDIENTE / A CONTINUAR"])
            
            area_opciones = ["-- Seleccionar --"] + AREAS
            area = st.selectbox("Área *", options=area_opciones)
            
            maquinas_opciones = ["-- Seleccionar --"] + list(MAQUINAS_DICT.keys())
            maquina_seleccionada = st.selectbox("Equipo o Máquina *", options=maquinas_opciones)
            
            if maquina_seleccionada == "Alquilada (especificar)":
                maquina_alquilada_detalle = st.text_input("Especifique Marca / Modelo (Máquina Alquilada):")
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

            tec_opciones = ["-- Seleccionar --"] + TECNICOS_OPCIONES
            tecnico_seleccionado = st.selectbox("Nombre del Técnico *", options=tec_opciones)
            
            if tecnico_seleccionado == "Otro (Especificar)":
                tecnico_otro = st.text_input("Especifique el nombre del Técnico:")
                tecnico_final = tecnico_otro if tecnico_otro else "Otro"
            elif tecnico_seleccionado != "-- Seleccionar --":
                tecnico_final = tecnico_seleccionado
            else:
                tecnico_final = ""

        with col2:
            tipo_mantenimiento = st.selectbox("Tipo Mantenimiento", ["CORRECTIVO", "PREVENTIVO", "PREDICTIVO"])
            
            horas_mantenimiento = "N/A"
            if tipo_mantenimiento in ["PREVENTIVO", "PREDICTIVO"]:
                horas_mantenimiento = st.selectbox(
                    "⚙️ Mantenimiento Periódico por Horas:", 
                    ["No Aplica", "250 hs", "500 hs", "1000 hs"]
                )
            
            prio_opciones = ["-- Seleccionar --", "ALTA", "MEDIA", "BAJA"]
            prioridad = st.selectbox("Prioridad *", options=prio_opciones)
            
            fecha_inicial = st.date_input("Fecha Inicial")
            hora_final = st.time_input("Hora Final")
            fecha_entrega = st.date_input("Fecha de Entrega")

        st.subheader("Detalles del Servicio")
        descripcion_del_servicio = st.text_area("Descripción del Servicio Realizado / Diagnóstico")
        
        motivo_pendiente = ""
        if estado_ot == "PENDIENTE / A CONTINUAR":
            st.warning("⚠️ Especifique la razón por la cual no se completó la tarea:")
            motivo_opcion = st.selectbox("Categoría del Motivo *", ["-- Seleccionar --"] + MOTIVOS_PENDIENTE_OPCIONES)
            motivo_detalle = st.text_area("Explicación Libre / Observación sobre el Pendiente *", placeholder="Escriba aquí los detalles del pendiente...")
            
            motivo_cat_str = motivo_opcion if motivo_opcion != "-- Seleccionar --" else ""
            motivo_pendiente = f"[{motivo_cat_str}] {motivo_detalle}".strip() if motivo_cat_str else motivo_detalle

        st.markdown("**Causa de Falla / Trabajo Realizado ***")
        causas_seleccionadas = st.multiselect("1. Selecciona una o más opciones estándar (puedes combinar Llanta, Mantenimientos, etc.):", options=CAUSAS_OPCIONES)
        causa_detalle_extra = st.text_area("2. Escribe o especifica aquí cualquier otra causa libremente:", placeholder="Detalles de la causa o trabajos adicionales...")

        st.subheader("Insumos y Repuestos")
        col_m1, col_m2 = st.columns([3, 1])
        with col_m1:
            materiales = st.text_area("Materiales y/o Repuestos Utilizados")
        with col_m2:
            insumos_unidades = st.number_input("Cantidad Insumos Usados", min_value=0, value=1)

        observaciones = st.text_area("Observaciones Generales")

        st.subheader("📷 Fotos Adjuntas")
        fotos_subidas = st.file_uploader("Adjuntar imágenes como evidencia:", type=["png", "jpg", "jpeg"], accept_multiple_files=True)

        submitted = st.form_submit_button("💾 Guardar y Registrar Orden de Servicio")

    if submitted:
        errores = []
        if area == "-- Seleccionar --": errores.append("Área")
        if maquina_seleccionada == "-- Seleccionar --": errores.append("Equipo o Máquina")
        if tecnico_seleccionado == "-- Seleccionar --" or (tecnico_seleccionado == "Otro (Especificar)" and not tecnico_final.strip()): errores.append("Nombre del Técnico")
        if prioridad == "-- Seleccionar --": errores.append("Prioridad")
        if not causas_seleccionadas and not causa_detalle_extra.strip() and horas_mantenimiento == "No Aplica": errores.append("Causa de Falla o Mantenimiento")
        if estado_ot == "PENDIENTE / A CONTINUAR" and not motivo_pendiente.strip(): errores.append("Detalle de Tareas Pendientes")

        if errores:
            st.error(f"⚠️ **Por favor completa los siguientes campos obligatorios:** {', '.join(errores)}")
            st.stop()

        partes_causa = []
        if horas_mantenimiento != "No Aplica" and horas_mantenimiento != "N/A":
            partes_causa.append(f"Mantenimiento {horas_mantenimiento}")
        if causas_seleccionadas: 
            partes_causa.append(", ".join(causas_seleccionadas))
        if causa_detalle_extra.strip(): 
            partes_causa.append(causa_detalle_extra.strip())
        
        causa_falla_final = " - ".join(partes_causa) if partes_causa else "N/A"
        
        categoria_ai = analizar_causa_con_gemini(causa_falla_final) if causa_falla_final != "N/A" else "N/A"
        fecha_actual_str = datetime.now().strftime("%Y-%m-%d")
        nombre_base_trabajo = f"{tecnico_final}_{fecha_actual_str}_{maquina_final}_{codigo_maq}_{num_ot}"

        datos_docx = {
            "<<area>>": area,
            "<<códigomaq>>": codigo_maq,
            "<<Maquina>>": maquina_final,
            "<<horometro>>": horometro,
            "<<tecnico>>": tecnico_final,
            "<<numOT>>": num_ot,
            "<<descripcion_del_servicio>>": f"[{estado_ot}] {descripcion_del_servicio}",
            "<<tipo_mantenimiento>>": f"{tipo_mantenimiento} ({horas_mantenimiento})" if horas_mantenimiento != "N/A" else tipo_mantenimiento,
            "<<prioridad>>": prioridad,
            "<<causa_falla>>": causa_falla_final,
            "<<Materiales>>": materiales,
            "<<fecha_inicial>>": str(fecha_inicial),
            "<<hora_final>>": str(hora_final),
            "<<fecha_de_entrega>>": str(fecha_entrega),
            "<<observaciones>>": f"{observaciones} | Pendiente: {motivo_pendiente}" if motivo_pendiente else observaciones
        }

        ruta_salida_docx = f"{nombre_base_trabajo}.docx"
        ruta_salida_pdf = f"{nombre_base_trabajo}.pdf"

        # Generar DOCX y PDF
        rellenar_plantilla(datos_docx, fotos_subidas, ruta_salida_docx)
        generar_pdf_reportlab(datos_docx, fotos_subidas, ruta_salida_pdf)

        archivo_a_respaldar = ruta_salida_pdf if os.path.exists(ruta_salida_pdf) else ruta_salida_docx

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

        # SUBIDA A CLOUDINARY
        respaldar_trabajo_en_cloudinary(num_ot, archivo_a_respaldar, fotos_subidas)

        st.success(f"✅ Orden {num_ot} guardada y respaldada en la nube con éxito en estado [{estado_ot}].")
        
        col_down1, col_down2 = st.columns(2)
        with col_down1:
            if os.path.exists(ruta_salida_pdf):
                with open(ruta_salida_pdf, "rb") as file_pdf:
                    st.download_button("📥 Descargar Orden (.pdf)", data=file_pdf, file_name=ruta_salida_pdf, mime="application/pdf")
        with col_down2:
            if os.path.exists(ruta_salida_docx):
                with open(ruta_salida_docx, "rb") as file_docx:
                    st.download_button("📥 Descargar Orden (.docx)", data=file_docx, file_name=ruta_salida_docx)

# ==========================================
# SECCIÓN 2: TRABAJOS PENDIENTES Y COMPLETADOS
# ==========================================
elif opcion == "⏳ Trabajos Pendientes":
    st.markdown("<h1 style='color: #A61C1C;'>⏳ Gestor de Trabajos Pendientes y Finalizados</h1>", unsafe_allow_html=True)
    st.write("Administra las Órdenes de Trabajo pendientes, complétalas o consulta el historial de trabajos finalizados por otros técnicos.")
    st.markdown("---")

    if os.path.exists(EXCEL_FILE):
        df = pd.read_excel(EXCEL_FILE)
    else:
        df = pd.DataFrame(columns=columnas_excel)

    tab1, tab2, tab3 = st.tabs([
        "✍️ Completar / Asumir Trabajo", 
        "📋 Pendientes Activos & Novedades", 
        "✅ Historial de Completados"
    ])

    with tab1:
        df_pend = df[df["Estado"] == "PENDIENTE"] if "Estado" in df.columns and not df.empty else pd.DataFrame()
        
        if not df_pend.empty:
            st.subheader("Transformar Trabajo Pendiente en Orden de Servicio Finalizada")
            
            ot_list = df_pend["Num_OT"].dropna().tolist()
            ot_seleccionada = st.selectbox("Seleccione el Número de Orden a Finalizar:", ot_list)
            
            row_ot = df_pend[df_pend["Num_OT"] == ot_seleccionada].iloc[0]
            
            st.info(f"📌 **Datos Registrados Previamente:** Área: **{row_ot.get('Area', 'N/A')}** | Equipo: **{row_ot.get('Maquina', 'N/A')}** | Técnico Inicial: **{row_ot.get('Tecnico_Inicial', 'N/A')}**")
            st.markdown(f"**Motivo / Tareas Pendientes Registradas:** {row_ot.get('Motivo_Pendiente', 'Sin detalle')}")
            
            with st.form("form_completar_pendiente"):
                col1, col2 = st.columns(2)
                with col1:
                    tec_final_opc = st.selectbox("Nombre del Técnico que Finaliza el Trabajo *", ["-- Seleccionar --"] + TECNICOS_OPCIONES)
                    tec_final_otro = st.text_input("Especifique Técnico (si seleccionó 'Otro'):") if tec_final_opc == "Otro (Especificar)" else ""
                    
                    horometro_val = float(row_ot.get("Horometro", 0.0)) if pd.notna(row_ot.get("Horometro")) else 0.0
                    horometro_f = st.number_input("Horómetro Actualizado", min_value=0.0, value=horometro_val, step=0.1)
                    
                with col2:
                    fecha_entrega_f = st.date_input("Fecha de Entrega Definitiva", value=datetime.now().date())
                    hora_final_f = st.time_input("Hora de Finalización", value=datetime.now().time())

                st.subheader("Detalles de la Reparación Efectuada")
                desc_final = st.text_area("Descripción del Trabajo Realizado para Finalizar la Orden *", placeholder="Explique qué acciones correctivas applied...")
                
                st.markdown("**Causa de Falla / Trabajos Detectados ***")
                causas_comp = st.multiselect("Seleccione causas estándar:", options=CAUSAS_OPCIONES)
                causa_extra_comp = st.text_area("Otras observaciones sobre la causa de falla:", placeholder="Detalles adicionales...")

                st.subheader("Insumos y Repuestos Utilizados")
                col_m1, col_m2 = st.columns([3, 1])
                with col_m1:
                    materiales_f = st.text_area("Materiales / Repuestos Adicionales Utilizados", value=str(row_ot.get("Materiales", "")) if pd.notna(row_ot.get("Materiales")) else "")
                with col_m2:
                    insumos_cant_f = st.number_input("Cantidad de Insumos Usados", min_value=0, value=int(row_ot.get("Insumo_Cantidad", 1)) if pd.notna(row_ot.get("Insumo_Cantidad")) else 1)

                observaciones_f = st.text_area("Observaciones Finales", value=str(row_ot.get("Observaciones", "")) if pd.notna(row_ot.get("Observaciones")) else "")
                
                st.subheader("📷 Fotos Adjuntas del Trabajo Concluido")
                fotos_subidas_f = st.file_uploader("Adjuntar fotos evidencia de finalización:", type=["png", "jpg", "jpeg"], accept_multiple_files=True, key="fotos_completar")

                btn_finalizar = st.form_submit_button("🏁 Finalizar Trabajo y Generar Documento OT")

            if btn_finalizar:
                tecnico_que_termina = tec_final_otro if tec_final_opc == "Otro (Especificar)" else tec_final_opc
                
                if tec_final_opc == "-- Seleccionar --" or not tecnico_que_termina.strip():
                    st.error("⚠️ Debe seleccionar el nombre del técnico que finaliza el trabajo.")
                    st.stop()
                if not desc_final.strip():
                    st.error("⚠️ Ingrese una descripción de los trabajos realizados.")
                    st.stop()

                partes_causa = []
                if causas_comp: partes_causa.append(", ".join(causas_comp))
                if causa_extra_comp.strip(): partes_causa.append(causa_extra_comp.strip())
                causa_falla_final = " - ".join(partes_causa) if partes_causa else str(row_ot.get("Causa_Falla", "N/A"))

                categoria_ai = analizar_causa_con_gemini(causa_falla_final)
                fecha_actual_str = datetime.now().strftime("%Y-%m-%d")

                idx_excel = df[df["Num_OT"] == ot_seleccionada].index[0]
                df.loc[idx_excel, "Estado"] = "FINALIZADO"
                df.loc[idx_excel, "Tecnico_Final"] = tecnico_que_termina
                df.loc[idx_excel, "Horometro"] = horometro_f
                df.loc[idx_excel, "Descripcion"] = f"{str(row_ot.get('Descripcion', ''))} | [COMPLETADO por {tecnico_que_termina}]: {desc_final}"
                df.loc[idx_excel, "Causa_Falla"] = causa_falla_final
                df.loc[idx_excel, "Categoria_Falla_AI"] = categoria_ai
                df.loc[idx_excel, "Materiales"] = materiales_f
                df.loc[idx_excel, "Insumo_Cantidad"] = insumos_cant_f
                df.loc[idx_excel, "Hora_Final"] = str(hora_final_f)
                df.loc[idx_excel, "Fecha_Entrega"] = str(fecha_entrega_f)
                df.loc[idx_excel, "Observaciones"] = observaciones_f

                df.to_excel(EXCEL_FILE, index=False)

                datos_docx = {
                    "<<area>>": row_ot.get("Area", ""),
                    "<<códigomaq>>": row_ot.get("Codigo_Maq", ""),
                    "<<Maquina>>": row_ot.get("Maquina", ""),
                    "<<horometro>>": horometro_f,
                    "<<tecnico>>": f"{row_ot.get('Tecnico_Inicial', '')} / {tecnico_que_termina}",
                    "<<numOT>>": ot_seleccionada,
                    "<<descripcion_del_servicio>>": f"[FINALIZADO] {desc_final}",
                    "<<tipo_mantenimiento>>": row_ot.get("Tipo_Mantenimiento", "CORRECTIVO"),
                    "<<prioridad>>": row_ot.get("Prioridad", "MEDIA"),
                    "<<causa_falla>>": causa_falla_final,
                    "<<Materiales>>": materiales_f,
                    "<<fecha_inicial>>": str(row_ot.get("Fecha_Inicial", "")),
                    "<<hora_final>>": str(hora_final_f),
                    "<<fecha_de_entrega>>": str(fecha_entrega_f),
                    "<<observaciones>>": observaciones_f
                }

                nombre_base = f"{tecnico_que_termina}_{fecha_actual_str}_{row_ot.get('Maquina','')}_{row_ot.get('Codigo_Maq','')}_{ot_seleccionada}"
                ruta_salida_docx = f"{nombre_base}.docx"
                ruta_salida_pdf = f"{nombre_base}.pdf"

                # Generar DOCX y PDF
                rellenar_plantilla(datos_docx, fotos_subidas_f, ruta_salida_docx)
                generar_pdf_reportlab(datos_docx, fotos_subidas_f, ruta_salida_pdf)

                archivo_a_respaldar = ruta_salida_pdf if os.path.exists(ruta_salida_pdf) else ruta_salida_docx

                # SUBIDA A CLOUDINARY
                respaldar_trabajo_en_cloudinary(ot_seleccionada, archivo_a_respaldar, fotos_subidas_f)

                st.success(f"🎉 ¡La Orden de Servicio {ot_seleccionada} fue completada y respaldada exitosamente!")
                
                col_down1, col_down2 = st.columns(2)
                with col_down1:
                    if os.path.exists(ruta_salida_pdf):
                        with open(ruta_salida_pdf, "rb") as file_pdf:
                            st.download_button("💾 Descargar Orden Final (.pdf)", data=file_pdf, file_name=ruta_salida_pdf, mime="application/pdf")
                with col_down2:
                    if os.path.exists(ruta_salida_docx):
                        with open(ruta_salida_docx, "rb") as file_docx:
                            st.download_button("💾 Descargar Orden Final (.docx)", data=file_docx, file_name=ruta_salida_docx)

        else:
            st.info("ℹ️ **No hay trabajos pendientes registrados actualmente.**")

    with tab2:
        df_pendientes = df[df["Estado"] == "PENDIENTE"] if "Estado" in df.columns and not df.empty else pd.DataFrame()
        if not df_pendientes.empty:
            st.info(f"Actualmente existen **{len(df_pendientes)}** trabajo(s) en curso o pendiente(s).")
            for idx, row in df_pendientes.iterrows():
                with st.container():
                    st.markdown(f"""
                    <div class="card-pendiente">
                        <h3 style="margin: 0;">{row['Num_OT']} - {row['Maquina']} ({row['Codigo_Maq']})</h3>
                        <p><b>Área:</b> {row['Area']} | <b>Prioridad:</b> {row['Prioridad']} | <b>Técnico Creador:</b> {row['Tecnico_Inicial']}</p>
                        <p><b>Motivo / Estado del Pendiente:</b><br>{row['Motivo_Pendiente']}</p>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    with st.expander(f"⚙️ Actualizar Novedad / Motivo del Pendiente ({row['Num_OT']})"):
                        with st.form(key=f"form_nov_{idx}"):
                            motivo_opcion_act = st.selectbox("Razón de Continuación de Pendiente:", MOTIVOS_PENDIENTE_OPCIONES, key=f"m_opc_{idx}")
                            motivo_texto_libre = st.text_area("Explicación Libre / Avances Realizados:", placeholder="Escriba aquí la razón detallada...", key=f"m_txt_{idx}")
                            
                            btn_actualizar_nov = st.form_submit_button("💾 Actualizar Novedad")
                            if btn_actualizar_nov:
                                if motivo_texto_libre.strip():
                                    nuevo_motivo_str = f"[{motivo_opcion_act}] {motivo_texto_libre.strip()}"
                                    df.loc[df["Num_OT"] == row['Num_OT'], "Motivo_Pendiente"] = f"{row['Motivo_Pendiente']} || {nuevo_motivo_str}"
                                    df.to_excel(EXCEL_FILE, index=False)
                                    st.success("✅ Novedad registrada correctamente.")
                                    st.rerun()
                                else:
                                    st.error("Ingrese una breve explicación antes de guardar.")
                    st.markdown("---")
        else:
            st.info("ℹ️ No existen trabajos pendientes activos registrados.")

    with tab3:
        df_completados = df[df["Estado"] == "FINALIZADO"] if "Estado" in df.columns and not df.empty else pd.DataFrame()
        if not df_completados.empty:
            st.write(f"Se han finalizado **{len(df_completados)}** trabajo(s) en total.")
            for idx, row in df_completados.iterrows():
                st.markdown(f"""
                <div class="card-completado">
                    <h3 style="margin: 0;">{row['Num_OT']} - {row['Maquina']} ({row['Codigo_Maq']})</h3>
                    <p><b>Área:</b> {row['Area']} | <b>Tipo:</b> {row['Tipo_Mantenimiento']} | <b>Técnico Finalizó:</b> {row['Tecnico_Final']}</p>
                    <p><b>Descripción del Trabajo:</b><br>{row['Descripcion']}</p>
                    <p><b>Causa Falla / Trabajos:</b> {row['Causa_Falla']} | <b>Materiales:</b> {row['Materiales']}</p>
                </div>
                """, unsafe_allow_html=True)
        else:
            st.info("ℹ️ Aún no existen trabajos completados en el historial.")

# ==========================================
# SECCIÓN 3: PANEL DE ESTADÍSTICAS REESTRUCTURADO
# ==========================================
elif opcion == "📊 Panel de Estadísticas":
    st.markdown("<h1 style='color: #A61C1C;'>📊 Panel de Estadísticas e Indicadores</h1>", unsafe_allow_html=True)
    st.markdown("---")

    if os.path.exists(EXCEL_FILE):
        df = pd.read_excel(EXCEL_FILE)
        if not df.empty:
            # Resumen general rápido
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Total Órdenes", len(df))
            c2.metric("Pendientes", len(df[df["Estado"] == "PENDIENTE"]) if "Estado" in df.columns else 0)
            c3.metric("Finalizadas", len(df[df["Estado"] == "FINALIZADO"]) if "Estado" in df.columns else 0)
            c4.metric("Insumos Usados", int(df["Insumo_Cantidad"].sum()) if "Insumo_Cantidad" in df.columns else 0)

            st.markdown("---")

            tab_est1, tab_est2, tab_est3 = st.tabs([
                "🚜 1. Problemas por Cada Máquina", 
                "⚠️ 2. Problemas Frecuentes (Global)", 
                "🛠️ 3. Mantenimientos Realizados"
            ])

            # -------------------------------------------------------------
            # SUBTAB 1: PROBLEMAS POR CADA MÁQUINA Y SU FRECUENCIA
            # -------------------------------------------------------------
            with tab_est1:
                st.subheader("Análisis Individual por Máquina / Equipo")
                maquinas_list = ["Todas"] + sorted(list(df["Maquina"].dropna().unique()))
                maquina_sel = st.selectbox("Seleccione una máquina para examinar:", maquinas_list)

                df_m = df if maquina_sel == "Todas" else df[df["Maquina"] == maquina_sel]

                if not df_m.empty:
                    col_m1, col_m2 = st.columns([2, 1])

                    with col_m1:
                        st.markdown(f"**Frecuencia de Causa de Falla / Trabajo en {maquina_sel}:**")
                        causas_series = df_m["Causa_Falla"].dropna().str.split(" - |, ").explode().str.strip()
                        causas_count = causas_series.value_counts().reset_index()
                        causas_count.columns = ["Problema / Causa", "Frecuencia"]

                        fig_maq = px.bar(
                            causas_count.head(10), x="Frecuencia", y="Problema / Causa", 
                            orientation='h', color_discrete_sequence=['#A61C1C']
                        )
                        fig_maq.update_layout(yaxis=dict(autorange="reversed"), paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
                        st.plotly_chart(fig_maq, use_container_width=True)

                    with col_m2:
                        st.markdown("**Resumen de Registros:**")
                        st.dataframe(causas_count, use_container_width=True, hide_index=True)
                else:
                    st.info("No hay registros para la máquina seleccionada.")

            # -------------------------------------------------------------
            # SUBTAB 2: PROBLEMAS FRECUENTES EN DISTINTAS MÁQUINAS
            # -------------------------------------------------------------
            with tab_est2:
                st.subheader("Ranking Global de Problemas Frecuentes en la Flota")
                
                col_g1, col_g2 = st.columns(2)

                with col_g1:
                    st.markdown("**Top 10 Problemas / Causalidades Más Frecuentes:**")
                    causas_globales = df["Causa_Falla"].dropna().str.split(" - |, ").explode().str.strip()
                    top_causas = causas_globales.value_counts().reset_index()
                    top_causas.columns = ["Falla / Causa", "Total Casos"]

                    fig_top = px.pie(top_causas.head(8), names="Falla / Causa", values="Total Casos", color_discrete_sequence=px.colors.sequential.RdBu)
                    st.plotly_chart(fig_top, use_container_width=True)

                with col_g2:
                    st.markdown("**Distribución de Categorías Técnicas (Análisis IA):**")
                    if "Categoria_Falla_AI" in df.columns:
                        cat_count = df["Categoria_Falla_AI"].value_counts().reset_index()
                        cat_count.columns = ["Categoría", "Frecuencia"]
                        
                        fig_cat = px.bar(cat_count, x="Categoría", y="Frecuencia", color_discrete_sequence=['#801414'])
                        fig_cat.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
                        st.plotly_chart(fig_cat, use_container_width=True)

            # -------------------------------------------------------------
            # SUBTAB 3: MANTENIMIENTOS HECHOS Y SU FRECUENCIA
            # -------------------------------------------------------------
            with tab_est3:
                st.subheader("Registro de Mantenimientos e Intervenciones")

                col_mant1, col_mant2 = st.columns(2)

                with col_mant1:
                    st.markdown("**Frecuencia por Tipo de Mantenimiento:**")
                    tipo_counts = df["Tipo_Mantenimiento"].value_counts().reset_index()
                    tipo_counts.columns = ["Tipo", "Cantidad"]

                    fig_tipo = px.bar(tipo_counts, x="Tipo", y="Cantidad", color="Tipo", color_discrete_sequence=['#27AE60', '#A61C1C', '#F39C12'])
                    fig_tipo.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
                    st.plotly_chart(fig_tipo, use_container_width=True)

                with col_mant2:
                    st.markdown("**Mantenimientos Preventivos por Horas (250h / 500h / 1000h):**")
                    if "Horas_Mantenimiento" in df.columns:
                        df_hrs = df[df["Horas_Mantenimiento"].isin(["250 hs", "500 hs", "1000 hs"])]
                        if not df_hrs.empty:
                            hrs_counts = df_hrs["Horas_Mantenimiento"].value_counts().reset_index()
                            hrs_counts.columns = ["Frecuencia Horas", "Cantidad Ejecutada"]

                            fig_hrs = px.pie(hrs_counts, names="Frecuencia Horas", values="Cantidad Ejecutada", hole=0.4, color_discrete_sequence=px.colors.qualitative.Dark24)
                            st.plotly_chart(fig_hrs, use_container_width=True)
                        else:
                            st.info("Aún no hay registrados mantenimientos de 250 hs, 500 hs o 1000 hs.")
                    else:
                        st.info("Registre nuevas órdenes para comenzar a visualizar esta estadística.")

            st.markdown("---")
            st.subheader("📄 Base de Datos Completa de Órdenes")
            st.dataframe(df, use_container_width=True)
        else:
            st.info("No hay datos en el registro para generar estadísticas.")