import os
import subprocess
from datetime import datetime
import pandas as pd
import requests
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Cm
from nicegui import app, ui

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
            for i, foto_path in enumerate(fotos_subidas, start=1):
                if os.path.exists(foto_path):
                    with open(foto_path, "rb") as foto_file:
                        response_foto = requests.post(
                            f"https://api.cloudinary.com/v1_1/{CLOUD_NAME}/image/upload",
                            data={
                                "upload_preset": UPLOAD_PRESET,
                                "folder": f"Ordenes_de_Trabajo/OT_{num_ot}",
                                "public_id": f"Foto_{i}_{num_ot}"
                            },
                            files={"file": foto_file}
                        )
                        if response_foto.status_code == 200:
                            urls_fotos.append(response_foto.json().get("secure_url"))

        return url_doc, urls_fotos
    except Exception as e:
        print(f"Error al subir respaldo a Cloudinary: {e}")
        return None, []

# ==========================================
# CONSTANTES Y ARCHIVOS
# ==========================================
EXCEL_FILE = "registro_ordenes_servicio.xlsx"
PLANTILLA_FILE = "plantilla_ot.docx"
TEMP_IMG_DIR = "temp_images"

os.makedirs(TEMP_IMG_DIR, exist_ok=True)

AREAS = ["Papelote", "Caldera", "Expedición", "Químicos", "Mecánicos", "Km4"]
TECNICOS_OPCIONES = ["Ivan Sosa", "Néstor Medina", "Gerardo Maidana", "Cristian Alvarenga"]
CAUSAS_OPCIONES = [
    "Desgaste natural", "Falta de lubricación", "Error operacional / manipulación", 
    "Sobrecalentamiento", "Fuga hidráulica/neumática", "Falla eléctrica/cortocircuito", 
    "Atascamiento / Muestra atascada", "Falta de mantenimiento preventivo", "Pieza defectuosa", "Llanta"
]

MAQUINAS_DICT = {
    "Cat 5": "Cat 5", "Cat 7 (topadora)": "Cat 7", "Cat 8": "Cat 8", "Cat 9": "Cat 9", "Cat 10": "Cat 10", "Cat 11": "Cat 11",
    "Linde 3": "Linde 3", "Linde 7": "Linde 7", "Linde 8": "Linde 8", "Linde 9": "Linde 9", "Linde 10": "Linde 10", 
    "Linde 11": "Linde 11", "Linde 12": "Linde 12", "Liugong 3": "Liugong 3", "Liugong 4": "Liugong 4", "Liugong 6": "Liugong 6", 
    "Liugong 7": "Liugong 7", "Liugong 8": "Liugong 8", "Clark 2": "Clark 2", "Clark 3": "Clark 3", "Clark 5": "Clark 5", 
    "Clark 6": "Clark 6", "Hyundai": "Hyundai"
}

columnas_excel = [
    "Num_OT", "Fecha_Registro", "Estado", "Area", "Codigo_Maq", "Maquina", "Horometro",
    "Tecnico_Inicial", "Tecnico_Final", "Descripcion", "Tipo_Mantenimiento", "Horas_Mantenimiento", "Prioridad", 
    "Causa_Falla", "Categoria_Falla_AI", "Motivo_Pendiente", "Materiales", "Insumo_Cantidad",
    "Fecha_Inicial", "Hora_Final", "Fecha_Entrega", "Observaciones"
]

if not os.path.exists(EXCEL_FILE):
    pd.DataFrame(columns=columnas_excel).to_excel(EXCEL_FILE, index=False)

def obtener_siguiente_ot():
    if os.path.exists(EXCEL_FILE):
        try:
            df_ot = pd.read_excel(EXCEL_FILE, usecols=[0])
            if not df_ot.empty:
                numeros = df_ot.iloc[:, 0].astype(str).str.extract(r'(\d+)')[0].dropna().astype(int)
                if not numeros.empty:
                    return f"OT-{(numeros.max() + 1):05d}"
        except Exception:
            pass
    return "OT-00001"

def convertir_docx_a_pdf(ruta_docx, ruta_pdf):
    try:
        subprocess.run(["soffice", "--headless", "--convert-to", "pdf", ruta_docx], check=True)
        return True
    except Exception:
        return False

def rellenar_plantilla(datos_dict, fotos_paths, ruta_salida_docx):
    if os.path.exists(PLANTILLA_FILE):
        doc = DocxTemplate(PLANTILLA_FILE)
        imagenes_inline = []
        if fotos_paths:
            for path in fotos_paths:
                if os.path.exists(path):
                    imagenes_inline.append(InlineImage(doc, path, width=Cm(12)))

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

# ==========================================
# INTERFAZ PRINCIPAL
# ==========================================

@ui.page('/')
def main_page():
    ui.colors(primary='#A61C1C')
    fotos_cargadas_temp = []

    # 1. Menú Lateral Drawer
    with ui.left_drawer(value=False).classes('bg-gray-100 p-2') as left_drawer:
        ui.label('Navegación').classes('font-bold text-gray-700 m-2')
        with ui.tabs().props('vertical').classes('w-full') as tabs:
            tab_cargar = ui.tab('📋 Cargar OT')
            tab_pendientes = ui.tab('⏳ Trabajos Pendientes')
            tab_historial = ui.tab('📂 Historial PDF')
            tab_estadisticas = ui.tab('📊 Estadísticas')

    # 2. Encabezado con Botón de Tres Rayas
    with ui.header().classes('bg-red-800 text-white flex items-center p-4 gap-4'):
        ui.button(icon='menu', on_click=lambda: left_drawer.toggle()).props('flat color=white')
        ui.label('YAGUARETE PAPELES - Gestión OT').classes('text-xl font-bold')

    with ui.tab_panels(tabs, value=tab_cargar).classes('w-full p-4'):
        
        # TAB 1: CARGAR ORDEN
        with ui.tab_panel(tab_cargar):
            ui.label('📋 Registro de Orden de Servicio').classes('text-2xl font-bold text-red-800 mb-4')
            
            with ui.card().classes('w-full p-4'):
                with ui.grid(columns=2).classes('w-full gap-4'):
                    in_num_ot = ui.input('Número de OT', value=obtener_siguiente_ot()).props('readonly')
                    in_estado = ui.select(['FINALIZADO', 'PENDIENTE / A CONTINUAR'], value='FINALIZADO', label='Estado *')
                    in_area = ui.select(AREAS, label='Área *')
                    in_maquina = ui.select(list(MAQUINAS_DICT.keys()), label='Equipo / Máquina *')
                    in_horometro = ui.number('Horómetro', value=0.0, format='%.1f')
                    in_tecnico = ui.select(TECNICOS_OPCIONES, label='Técnico *')
                    in_tipo_mant = ui.select(['CORRECTIVO', 'PREVENTIVO', 'PREDICTIVO'], value='CORRECTIVO', label='Tipo Mantenimiento')
                    in_prioridad = ui.select(['ALTA', 'MEDIA', 'BAJA'], value='MEDIA', label='Prioridad *')
                    in_fecha_ini = ui.input('Fecha Inicial', value=datetime.now().strftime('%Y-%m-%d'))
                    in_fecha_ent = ui.input('Fecha Entrega', value=datetime.now().strftime('%Y-%m-%d'))

                in_descripcion = ui.textarea('Descripción del Servicio / Diagnóstico').classes('w-full mt-2')
                in_causas = ui.select(CAUSAS_OPCIONES, multiple=True, label='Causas Estándar').classes('w-full mt-2')
                in_materiales = ui.textarea('Materiales / Repuestos Utilizados').classes('w-full mt-2')
                in_observaciones = ui.textarea('Observaciones Generales').classes('w-full mt-2')

                ui.label('📷 Adjuntar Fotografías del Servicio').classes('font-bold text-gray-700 mt-4')
                
                def manejar_subida_imagen(e):
                    path_destino = os.path.join(TEMP_IMG_DIR, e.name)
                    with open(path_destino, 'wb') as f:
                        f.write(e.content.read())
                    fotos_cargadas_temp.append(path_destino)
                    ui.notify(f'📷 Imagen subida: {e.name}', type='positive')

                ui.upload(
                    label='Seleccionar o capturar fotos',
                    multiple=True,
                    auto_upload=True,
                    on_upload=manejar_subida_imagen
                ).props('accept="image/*" capture="environment"').classes('w-full mt-2')

                row_descarga = ui.row().classes('w-full my-2')
                with row_descarga:
                    btn_download_doc = ui.button('📥 Descargar Documento').classes('bg-red-800 text-white')
                    btn_download_doc.set_visibility(False)

                def procesar_guardado():
                    if not in_area.value or not in_maquina.value or not in_tecnico.value:
                        ui.notify('⚠️ Complete los campos obligatorios (*)', type='warning')
                        return

                    num_ot_curr = in_num_ot.value
                    codigo_m = MAQUINAS_DICT.get(in_maquina.value, "")
                    nombre_base = f"{in_tecnico.value}_{datetime.now().strftime('%Y-%m-%d')}_{in_maquina.value}_{num_ot_curr}"
                    ruta_docx = f"{nombre_base}.docx"
                    ruta_pdf = f"{nombre_base}.pdf"

                    causa_str = ", ".join(in_causas.value) if in_causas.value else "N/A"

                    datos_docx = {
                        "area": in_area.value, "códigomaq": codigo_m, "Maquina": in_maquina.value,
                        "horometro": in_horometro.value, "tecnico": in_tecnico.value, "numOT": num_ot_curr,
                        "descripcion_del_servicio": f"[{in_estado.value}] {in_descripcion.value}",
                        "tipo_mantenimiento": in_tipo_mant.value, "prioridad": in_prioridad.value,
                        "causa_falla": causa_str, "Materiales": in_materiales.value,
                        "fecha_inicial": in_fecha_ini.value, "hora_final": datetime.now().strftime('%H:%M:%S'),
                        "fecha_de_entrega": in_fecha_ent.value, "observaciones": in_observaciones.value
                    }

                    rellenar_plantilla(datos_docx, fotos_cargadas_temp, ruta_docx)
                    se_convertio = convertir_docx_a_pdf(ruta_docx, ruta_pdf)
                    archivo_final = ruta_pdf if se_convertio and os.path.exists(ruta_pdf) else ruta_docx

                    df_ex = pd.read_excel(EXCEL_FILE)
                    nueva_fila = {
                        "Num_OT": num_ot_curr, "Fecha_Registro": datetime.now().strftime("%Y-%m-%d %H:%M"),
                        "Estado": "PENDIENTE" if in_estado.value == "PENDIENTE / A CONTINUAR" else "FINALIZADO",
                        "Area": in_area.value, "Codigo_Maq": codigo_m, "Maquina": in_maquina.value,
                        "Horometro": in_horometro.value, "Tecnico_Inicial": in_tecnico.value,
                        "Descripcion": in_descripcion.value, "Tipo_Mantenimiento": in_tipo_mant.value,
                        "Prioridad": in_prioridad.value, "Causa_Falla": causa_str, "Materiales": in_materiales.value,
                        "Fecha_Inicial": in_fecha_ini.value, "Fecha_Entrega": in_fecha_ent.value,
                        "Observaciones": in_observaciones.value
                    }
                    pd.concat([df_ex, pd.DataFrame([nueva_fila])], ignore_index=True).to_excel(EXCEL_FILE, index=False)

                    respaldar_trabajo_en_cloudinary(num_ot_curr, archivo_final, fotos_cargadas_temp)

                    btn_download_doc.on_click(lambda: ui.download(archivo_final))
                    btn_download_doc.set_visibility(True)

                    ui.notify(f'✅ Orden {num_ot_curr} guardada correctamente', type='positive')

                    in_descripcion.value = ''
                    in_materiales.value = ''
                    in_observaciones.value = ''
                    fotos_cargadas_temp.clear()
                    in_num_ot.value = obtener_siguiente_ot()

                ui.button('💾 Guardar y Registrar Orden', on_click=procesar_guardado).classes('w-full bg-red-800 text-white font-bold my-4')

        # TAB 2: PENDIENTES
        with ui.tab_panel(tab_pendientes):
            ui.label('⏳ Gestor de Trabajos Pendientes').classes('text-2xl font-bold text-red-800 mb-4')
            
            def refrescar_pendientes():
                container_pendientes.clear()
                df = pd.read_excel(EXCEL_FILE) if os.path.exists(EXCEL_FILE) else pd.DataFrame()
                df_pend = df[df["Estado"] == "PENDIENTE"] if not df.empty and "Estado" in df.columns else pd.DataFrame()
                
                with container_pendientes:
                    if not df_pend.empty:
                        for _, row in df_pend.iterrows():
                            with ui.card().classes('w-full mb-2 p-3'):
                                ui.label(f"📌 OT: {row['Num_OT']} | Equipo: {row['Maquina']} | Técnico: {row['Tecnico_Inicial']}").classes('font-bold')
                                ui.label(f"Descripción: {row['Descripcion']}")
                    else:
                        ui.label('ℹ️ No hay trabajos pendientes registrados.').classes('text-gray-500')

            container_pendientes = ui.column().classes('w-full')
            ui.button('🔄 Refrescar Pendientes', on_click=refrescar_pendientes).classes('mb-2')
            refrescar_pendientes()

        # TAB 3: HISTORIAL
        with ui.tab_panel(tab_historial):
            ui.label('📂 Historial de Documentos Generados').classes('text-2xl font-bold text-red-800 mb-4')
            
            def refrescar_historial():
                container_historial.clear()
                with container_historial:
                    archivos = [f for f in os.listdir('.') if (f.endswith('.pdf') or f.endswith('.docx')) and f != PLANTILLA_FILE]
                    if archivos:
                        for arch in sorted(archivos, reverse=True):
                            with ui.row().classes('w-full items-center justify-between p-2 border-b'):
                                ui.label(f"📄 {arch}")
                                ui.button('Descargar', on_click=lambda a=arch: ui.download(a)).props('flat').classes('text-red-800')
                    else:
                        ui.label('No hay documentos generados aún localmente.').classes('text-gray-500')

            container_historial = ui.column().classes('w-full')
            ui.button('🔄 Actualizar Lista', on_click=refrescar_historial).classes('mb-2')
            refrescar_historial()

        # TAB 4: ESTADÍSTICAS
        with ui.tab_panel(tab_estadisticas):
            ui.label('📊 Indicadores de Gestión OT').classes('text-2xl font-bold text-red-800 mb-4')

            def cargar_estadisticas():
                container_stats.clear()
                with container_stats:
                    if not os.path.exists(EXCEL_FILE):
                        ui.label('No hay datos registrados aún.').classes('text-gray-500')
                        return

                    df = pd.read_excel(EXCEL_FILE)
                    if df.empty:
                        ui.label('El registro de órdenes está vacío.').classes('text-gray-500')
                        return

                    # KPIs
                    total_ot = len(df)
                    finalizadas = len(df[df['Estado'] == 'FINALIZADO']) if 'Estado' in df.columns else 0
                    pendientes = len(df[df['Estado'] == 'PENDIENTE']) if 'Estado' in df.columns else 0

                    with ui.grid(columns=3).classes('w-full gap-4 mb-6'):
                        with ui.card().classes('p-4 text-center bg-gray-50'):
                            ui.label('Total OT').classes('text-sm text-gray-600')
                            ui.label(str(total_ot)).classes('text-3xl font-bold text-red-800')
                        with ui.card().classes('p-4 text-center bg-gray-50'):
                            ui.label('Finalizadas').classes('text-sm text-gray-600')
                            ui.label(str(finalizadas)).classes('text-3xl font-bold text-green-700')
                        with ui.card().classes('p-4 text-center bg-gray-50'):
                            ui.label('Pendientes').classes('text-sm text-gray-600')
                            ui.label(str(pendientes)).classes('text-3xl font-bold text-yellow-700')

                    # Gráficos
                    with ui.grid(columns=2).classes('w-full gap-4'):
                        # OT por Área
                        if 'Area' in df.columns:
                            area_counts = df['Area'].value_counts()
                            ui.chart({
                                'title': {'text': 'Órdenes por Área'},
                                'chart': {'type': 'pie'},
                                'series': [{'name': 'Órdenes', 'data': [{'name': k, 'y': int(v)} for k, v in area_counts.items()]}]
                            }).classes('w-full h-64')

                        # OT por Técnico
                        if 'Tecnico_Inicial' in df.columns:
                            tec_counts = df['Tecnico_Inicial'].value_counts()
                            ui.chart({
                                'title': {'text': 'Órdenes por Técnico'},
                                'chart': {'type': 'column'},
                                'xAxis': {'categories': list(tec_counts.index)},
                                'series': [{'name': 'Cantidad', 'data': [int(v) for v in tec_counts.values]}]
                            }).classes('w-full h-64')

            container_stats = ui.column().classes('w-full')
            ui.button('🔄 Actualizar Estadísticas', on_click=cargar_estadisticas).classes('mb-4')
            cargar_estadisticas()

ui.run(port=8080, title="Yaguarete OT")