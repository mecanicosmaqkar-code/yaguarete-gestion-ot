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
    url_doc = None
    urls_fotos = []

    try:
        # Subir Documento (PDF o DOCX)
        if os.path.exists(ruta_archivo):
            endpoint_url = f"https://api.cloudinary.com/v1_1/{CLOUD_NAME}/raw/upload"
            
            with open(ruta_archivo, "rb") as file_to_upload:
                payload = {
                    "upload_preset": UPLOAD_PRESET,
                    "folder": f"Ordenes_de_Trabajo/OT_{num_ot}"
                }
                files = {"file": (os.path.basename(ruta_archivo), file_to_upload)}
                
                response = requests.post(endpoint_url, data=payload, files=files)
                
                if response.status_code == 200:
                    url_doc = response.json().get("secure_url")
                    print(f"✅ Documento respaldado exitosamente en Cloudinary: {url_doc}")
                else:
                    print(f"❌ Error Cloudinary Documento [{response.status_code}]: {response.text}")

        # Subir Fotografías
        if fotos_subidas:
            endpoint_foto_url = f"https://api.cloudinary.com/v1_1/{CLOUD_NAME}/image/upload"
            for foto_path in fotos_subidas:
                if os.path.exists(foto_path):
                    with open(foto_path, "rb") as foto_file:
                        payload_foto = {
                            "upload_preset": UPLOAD_PRESET,
                            "folder": f"Ordenes_de_Trabajo/OT_{num_ot}"
                        }
                        files_foto = {"file": (os.path.basename(foto_path), foto_file)}
                        
                        resp_foto = requests.post(endpoint_foto_url, data=payload_foto, files=files_foto)
                        if resp_foto.status_code == 200:
                            urls_fotos.append(resp_foto.json().get("secure_url"))
                        else:
                            print(f"❌ Error Cloudinary Imagen [{resp_foto.status_code}]: {resp_foto.text}")

        return url_doc, urls_fotos

    except Exception as e:
        print(f"💥 Excepción al subir a Cloudinary: {e}")
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
    "Fecha_Inicial", "Hora_Final", "Fecha_Entrega", "Observaciones", "URL_Cloudinary"
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

def buscar_archivo_ot(num_ot):
    if not num_ot or pd.isna(num_ot):
        return None
    num_ot_clean = str(num_ot).strip()
    archivos = os.listdir('.')
    for f in archivos:
        if num_ot_clean in f and f.endswith('.pdf'):
            return f
    for f in archivos:
        if num_ot_clean in f and f.endswith('.docx') and f != PLANTILLA_FILE:
            return f
    return None

# ==========================================
# INTERFAZ PRINCIPAL
# ==========================================

@ui.page('/')
def main_page():
    ui.colors(primary='#A61C1C')
    fotos_cargadas_temp = []

    with ui.left_drawer(value=False).classes('bg-gray-100 p-2') as left_drawer:
        ui.label('Navegación').classes('font-bold text-gray-700 m-2')
        with ui.tabs().props('vertical').classes('w-full') as tabs:
            tab_cargar = ui.tab('📋 Cargar OT')
            tab_pendientes = ui.tab('⏳ Trabajos Pendientes')
            tab_historial = ui.tab('📂 Historial PDF')
            tab_estadisticas = ui.tab('📊 Estadísticas')

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
                    try:
                        nombre_archivo = getattr(e, 'name', None) or getattr(e, 'filename', f"img_{len(fotos_cargadas_temp)}.jpg")
                        contenido = getattr(e, 'content', None)
                        
                        if contenido:
                            path_destino = os.path.join(TEMP_IMG_DIR, nombre_archivo)
                            data_bytes = contenido.read() if hasattr(contenido, 'read') else contenido
                            with open(path_destino, 'wb') as f:
                                f.write(data_bytes)
                            fotos_cargadas_temp.append(path_destino)
                            ui.notify(f'📷 Imagen subida: {nombre_archivo}', type='positive')
                    except Exception as err:
                        print(f"Error procesando imagen: {err}")

                ui.upload(
                    label='Seleccionar o capturar fotos',
                    multiple=True,
                    auto_upload=True,
                    on_upload=manejar_subida_imagen
                ).props('accept="image/*" capture="environment"').classes('w-full mt-2')

                def procesar_guardado():
                    if not in_area.value or not in_maquina.value or not in_tecnico.value:
                        ui.notify('⚠️ Complete los campos obligatorios (*)', type='warning')
                        return

                    num_ot_curr = in_num_ot.value
                    codigo_m = MAQUINAS_DICT.get(in_maquina.value, "")
                    nombre_base = f"{in_tecnico.value}_{datetime.now().strftime('%Y-%m-%d')}_{in_maquina.value}_{num_ot_curr}".replace(" ", "_")
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

                    # Subida a Cloudinary
                    url_doc_cloud, _ = respaldar_trabajo_en_cloudinary(num_ot_curr, archivo_final, fotos_cargadas_temp)

                    df_ex = pd.read_excel(EXCEL_FILE)
                    nueva_fila = {
                        "Num_OT": num_ot_curr, 
                        "Fecha_Registro": datetime.now().strftime("%Y-%m-%d %H:%M"),
                        "Estado": "PENDIENTE" if in_estado.value == "PENDIENTE / A CONTINUAR" else "FINALIZADO",
                        "Area": in_area.value, 
                        "Codigo_Maq": codigo_m, 
                        "Maquina": in_maquina.value,
                        "Horometro": in_horometro.value, 
                        "Tecnico_Inicial": in_tecnico.value,
                        "Descripcion": in_descripcion.value, 
                        "Tipo_Mantenimiento": in_tipo_mant.value,
                        "Prioridad": in_prioridad.value, 
                        "Causa_Falla": causa_str, 
                        "Materiales": in_materiales.value,
                        "Fecha_Inicial": in_fecha_ini.value, 
                        "Fecha_Entrega": in_fecha_ent.value,
                        "Observaciones": in_observaciones.value,
                        "URL_Cloudinary": url_doc_cloud if url_doc_cloud else ""
                    }
                    pd.concat([df_ex, pd.DataFrame([nueva_fila])], ignore_index=True).to_excel(EXCEL_FILE, index=False)

                    # Navegación hacia el archivo generado
                    if url_doc_cloud:
                        ui.navigate.to(url_doc_cloud, new_tab=True)
                    elif os.path.exists(archivo_final):
                        ui.download(archivo_final)

                    ui.notify(f'✅ Orden {num_ot_curr} guardada correctamente', type='positive')

                    # Limpieza del formulario y temporales
                    in_descripcion.value = ''
                    in_materiales.value = ''
                    in_observaciones.value = ''
                    
                    for p in fotos_cargadas_temp:
                        if os.path.exists(p):
                            try:
                                os.remove(p)
                            except Exception:
                                pass
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
            ui.label('📊 Panel de Estadísticas y Análisis de Mantenimiento').classes('text-2xl font-bold text-red-800 mb-4')

            def renderizar_estadisticas():
                container_stats.clear()
                with container_stats:
                    if not os.path.exists(EXCEL_FILE):
                        ui.label('No hay datos registrados aún.').classes('text-gray-500')
                        return

                    df = pd.read_excel(EXCEL_FILE)
                    if df.empty:
                        ui.label('El registro de órdenes está vacío.').classes('text-gray-500')
                        return

                    with ui.row().classes('w-full items-start gap-4 flex-col md:flex-row'):
                        
                        # BARRA LATERAL
                        with ui.card().classes('w-full md:w-1/4 p-4 bg-gray-50'):
                            ui.label('🔍 Seleccionar Equipo').classes('font-bold text-lg text-gray-800 mb-2')
                            
                            opciones_maquinas = ['TODAS LAS MÁQUINAS'] + sorted([str(m) for m in df['Maquina'].dropna().unique()])
                            sel_maquina = ui.select(
                                opciones_maquinas, 
                                value='TODAS LAS MÁQUINAS', 
                                label='Filtrar Máquina'
                            ).classes('w-full')

                            ui.separator().classes('my-4')
                            ui.label('💡 Indicación:').classes('text-xs font-bold text-gray-500')
                            ui.label('Selecciona una máquina para ver sus métricas y descargar los PDF asociados.').classes('text-xs text-gray-500')

                        # ÁREA CENTRAL
                        contenido_central = ui.column().classes('w-full md:w-3/4')

                        def actualizar_contenido_central(maq_seleccionada):
                            contenido_central.clear()
                            
                            if maq_seleccionada == 'TODAS LAS MÁQUINAS':
                                df_filtrado = df.copy()
                                titulo_seccion = "Resumen General de la Flota"
                            else:
                                df_filtrado = df[df['Maquina'] == maq_seleccionada]
                                titulo_seccion = f"Análisis y Diagnóstico: {maq_seleccionada}"

                            with contenido_central:
                                ui.label(titulo_seccion).classes('text-xl font-bold text-red-800 mb-2')

                                if df_filtrado.empty:
                                    ui.label('No hay registros disponibles para la selección.').classes('text-gray-500 my-4')
                                    return

                                total_ot = len(df_filtrado)
                                finalizadas = len(df_filtrado[df_filtrado['Estado'] == 'FINALIZADO']) if 'Estado' in df_filtrado.columns else 0
                                pendientes = len(df_filtrado[df_filtrado['Estado'] == 'PENDIENTE']) if 'Estado' in df_filtrado.columns else 0

                                with ui.grid(columns=3).classes('w-full gap-2 mb-4'):
                                    with ui.card().classes('p-3 text-center bg-gray-50'):
                                        ui.label('Total OT').classes('text-xs text-gray-600')
                                        ui.label(str(total_ot)).classes('text-2xl font-bold text-red-800')
                                    with ui.card().classes('p-3 text-center bg-gray-50'):
                                        ui.label('Finalizadas').classes('text-xs text-gray-600')
                                        ui.label(str(finalizadas)).classes('text-2xl font-bold text-green-700')
                                    with ui.card().classes('p-3 text-center bg-gray-50'):
                                        ui.label('Pendientes').classes('text-xs text-gray-600')
                                        ui.label(str(pendientes)).classes('text-2xl font-bold text-yellow-700')

                                # 1. ANÁLISIS DE FALLAS
                                if 'Causa_Falla' in df_filtrado.columns and not df_filtrado['Causa_Falla'].dropna().empty:
                                    causas_list = []
                                    for c in df_filtrado['Causa_Falla'].dropna():
                                        causas_list.extend([x.strip() for x in str(c).split(',') if x.strip() != 'N/A'])
                                    
                                    if causas_list:
                                        causas_series = pd.Series(causas_list).value_counts().head(5)
                                        ui.label('🚨 Fallas y Causas Más Comunes').classes('font-bold text-gray-700 mt-2')
                                        ui.echart({
                                            'tooltip': {'trigger': 'axis'},
                                            'xAxis': {'type': 'value'},
                                            'yAxis': {'type': 'category', 'data': [str(k) for k in causas_series.index[::-1]]},
                                            'series': [{
                                                'data': [int(v) for v in causas_series.values[::-1]],
                                                'type': 'bar',
                                                'itemStyle': {'color': '#A61C1C'}
                                            }]
                                        }).classes('w-full h-48')

                                # 2. HISTORIAL Y DESCARGAS DESDE CLOUDINARY
                                ui.label('📜 Historial de Trabajos e Intervenciones').classes('font-bold text-gray-700 mt-6 mb-2')
                                
                                with ui.card().classes('w-full p-2 max-h-96 overflow-y-auto'):
                                    for _, row in df_filtrado.sort_values(by='Fecha_Registro', ascending=False).iterrows():
                                        num_ot_val = str(row.get('Num_OT', ''))
                                        archivo_encontrado = buscar_archivo_ot(num_ot_val)
                                        url_cloudinary = row.get('URL_Cloudinary', '') if pd.notna(row.get('URL_Cloudinary')) else None

                                        estado_color = 'text-green-700' if row.get('Estado') == 'FINALIZADO' else 'text-yellow-700'
                                        with ui.column().classes('w-full p-2 border-b text-sm'):
                                            with ui.row().classes('w-full justify-between items-center font-bold'):
                                                ui.label(f"OT: {num_ot_val} | Fecha: {str(row.get('Fecha_Registro', 'N/A'))[:10]}")
                                                
                                                with ui.row().classes('items-center gap-2'):
                                                    ui.label(f"[{row.get('Estado', 'N/A')}]").classes(estado_color)
                                                    
                                                    # Prioridad 1: Descargar archivo local si aún existe
                                                    if archivo_encontrado:
                                                        ui.button(
                                                            '📥 PDF Local', 
                                                            on_click=lambda a=archivo_encontrado: ui.download(a)
                                                        ).props('dense size=sm color=red').classes('text-xs')
                                                    # Prioridad 2: Abrir archivo subido a Cloudinary
                                                    elif url_cloudinary and str(url_cloudinary).startswith('http'):
                                                        ui.button(
                                                            '☁️ Abrir PDF', 
                                                            on_click=lambda u=url_cloudinary: ui.navigate.to(u, new_tab=True)
                                                        ).props('dense size=sm color=blue').classes('text-xs')
                                                    else:
                                                        ui.label('Sin PDF').classes('text-xs text-gray-400')

                                            ui.label(f"🔧 Máquina: {row.get('Maquina', 'N/A')} | Técnico: {row.get('Tecnico_Inicial', 'N/A')} | Horómetro: {row.get('Horometro', 0)}")
                                            ui.label(f"📝 Servicio: {row.get('Descripcion', 'Sin descripción')}")

                        sel_maquina.on('update:model-value', lambda e: actualizar_contenido_central(e.args))
                        actualizar_contenido_central('TODAS LAS MÁQUINAS')

            container_stats = ui.column().classes('w-full')
            ui.button('🔄 Refrescar Datos', on_click=renderizar_estadisticas).classes('mb-4 bg-red-800 text-white')
            renderizar_estadisticas()

ui.run(port=8080, title="Yaguarete OT")
