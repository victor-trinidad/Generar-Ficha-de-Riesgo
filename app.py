import pandas as pd
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import streamlit as st
import io # Para manejar la descarga del archivo Word

# --- CONFIGURACIÓN ESPECÍFICA DEL ARCHIVO ---
ARCHIVO_EXCEL = 'LMM_ORG_04 Rev. 00 - Matriz Institucional de Gestión de Riesgos.xlsx'
NOMBRE_HOJA = 'LMM_ORG_04' 
FILA_ENCABEZADOS = 16 

# Renombrar columnas para facilitar el acceso
COLUMNAS_MAP = [
    'Num_Riesgo', 'Entorno_Control', 'Origen_Area', 'Proceso_Documento', 
    'Riesgo_Identificado', 'Impacto_Potencial', 'Efecto', 
    'Gravedad', 'Probabilidad', 'PxG', 'Escala_Riesgo', 
    'Control_Existente', 'Tipo_Control', 'Responsable_Seguimiento', 
    'Eficacia_Seguimiento', 'Version', 'Estado_Control', 
    'Acciones', 'Fecha_Identificacion', 'Ultima_Revision'
]

# --- FUNCIONES AUXILIARES DE GENERACIÓN ---

def agregar_seccion_tabla(document, titulo, datos_dict):
    """Agrega una sección formal usando una tabla de una columna."""
    document.add_heading(titulo, level=2)
    tabla = document.add_table(rows=len(datos_dict), cols=2)
    tabla.style = 'Table Grid'
    tabla.columns[0].width = Inches(2)
    
    i = 0
    for key, value in datos_dict.items():
        row_cells = tabla.rows[i].cells
        row_cells[0].paragraphs[0].add_run(f'{key}').bold = True
        row_cells[1].paragraphs[0].add_run(str(value))
        i += 1
    document.add_paragraph()

def generar_ficha_docx(datos_riesgo):
    """
    Genera la ficha A4 y devuelve el documento como un objeto BytesIO 
    para poder ser descargado en Streamlit.
    """
    document = Document()
    
    # --- Configuración Estilística ---
    section = document.sections[0]
    section.top_margin, section.bottom_margin = Inches(0.5), Inches(0.5)
    section.left_margin, section.right_margin = Inches(0.75), Inches(0.75)
    
    # --- TÍTULO PRINCIPAL DE LA FICHA ---
    document.add_heading(f'FICHA DE GESTIÓN DE RIESGO N° {datos_riesgo["Num_Riesgo"]}', level=0)
    document.add_paragraph(f'Versión: {datos_riesgo["Version"]} | Última Revisión: {datos_riesgo["Ultima_Revision"]}')
    document.add_paragraph('---') 

    # 1) - IDENTIFICACIÓN DEL RIESGO
    identificacion_data = {
        'Riesgo Identificado': datos_riesgo['Riesgo_Identificado'],
        'Entorno de Control': datos_riesgo['Entorno_Control'],
        'Origen / Área Responsable': datos_riesgo['Origen_Area'],
        'Proceso o Documento': datos_riesgo['Proceso_Documento']
    }
    agregar_seccion_tabla(document, '1) IDENTIFICACIÓN DEL RIESGO', identificacion_data)

    # 2) - ANÁLISIS DEL RIESGO
    analisis_data = {
        'Impacto Potencial': datos_riesgo['Impacto_Potencial'],
        'Efecto (Consecuencias)': datos_riesgo['Efecto']
    }
    agregar_seccion_tabla(document, '2) ANÁLISIS DEL RIESGO', analisis_data)

    # 3) - EVALUACIÓN DEL RIESGO
    document.add_heading('3) EVALUACIÓN DEL RIESGO', level=2)
    tabla_evaluacion = document.add_table(rows=2, cols=4)
    tabla_evaluacion.style = 'Table Grid'
    
    # Encabezados
    hdr_cells = tabla_evaluacion.rows[0].cells
    for i, text in enumerate(['Gravedad (G)', 'Probabilidad (P)', 'Resultado (P x G)', 'ESCALA DE RIESGO']):
        hdr_cells[i].paragraphs[0].add_run(text).bold = True
        hdr_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Valores
    val_cells = tabla_evaluacion.rows[1].cells
    val_cells[0].text = str(datos_riesgo['Gravedad'])
    val_cells[1].text = str(datos_riesgo['Probabilidad'])
    val_cells[2].text = str(datos_riesgo['PxG'])
    escala_run = val_cells[3].paragraphs[0].add_run(str(datos_riesgo['Escala_Riesgo']).upper())
    escala_run.bold = True
    val_cells[3].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    document.add_paragraph()

    # 4) - SEGUIMIENTO DEL RIESGO
    seguimiento_data = {
        'Responsable del Seguimiento': datos_riesgo['Responsable_Seguimiento'],
        'Tipo de Control': datos_riesgo['Tipo_Control'],
        'Eficacia del Seguimiento': datos_riesgo['Eficacia_Seguimiento']
    }
    agregar_seccion_tabla(document, '4) SEGUIMIENTO DEL RIESGO', seguimiento_data)
    
    document.add_heading('Descripción del Control Existente', level=3)
    document.add_paragraph(datos_riesgo['Control_Existente'])

    # 5) - SEGUIMIENTO DE VERSIONES Y ACCIONES
    document.add_heading('5) SEGUIMIENTO DE VERSIONES Y ACCIONES', level=2)
    
    tabla_versiones = document.add_table(rows=1, cols=3)
    tabla_versiones.style = 'Table Grid'
    vers_cells = tabla_versiones.rows[0].cells
    vers_cells[0].paragraphs[0].add_run(f'Versión: {datos_riesgo["Version"]}').bold = True
    vers_cells[1].paragraphs[0].add_run(f'Estado: {datos_riesgo["Estado_Control"]}').bold = True
    vers_cells[2].paragraphs[0].add_run(f'Fecha Ident.: {datos_riesgo["Fecha_Identificacion"]}').bold = True
    document.add_paragraph()
    
    document.add_heading('Acciones Pendientes / Recomendadas', level=3)
    document.add_paragraph(datos_riesgo['Acciones'])
    
    # Guardar en memoria (BytesIO) y devolver
    buffer = io.BytesIO()
    document.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


# --- INTERFAZ STREAMLIT PRINCIPAL ---
st.set_page_config(page_title="Generador de Fichas de Riesgo", layout="wide")

st.title("Generador de Fichas de Riesgo Individuales")
st.markdown("Selecciona un riesgo de la lista para generar y descargar su ficha A4.")

@st.cache_data
def cargar_datos(archivo, hoja, encabezados, columnas):
    """Carga y procesa los datos del Excel (Cache para rendimiento)."""
    try:
        df = pd.read_excel(
            archivo, 
            sheet_name=hoja, 
            header=encabezados,
            usecols="B:U"
        )
        df = df.fillna("")
        df.columns = columnas
        # Filtra filas que no tienen Número de Riesgo
        df = df[df['Num_Riesgo'] != ""].reset_index(drop=True)
        return df
    except FileNotFoundError:
        st.error(f"Error: No se encontró el archivo de matriz '{archivo}'. Asegúrate de que está en la misma carpeta.")
        return pd.DataFrame()

df_riesgos = cargar_datos(ARCHIVO_EXCEL, NOMBRE_HOJA, FILA_ENCABEZADOS, COLUMNAS_MAP)

if not df_riesgos.empty:
    
    # Crear una lista de opciones para el selector
    # Se utiliza el Num_Riesgo y la Descripción para fácil identificación
    opciones_riesgos = (df_riesgos['Num_Riesgo'] + ' - ' + df_riesgos['Riesgo_Identificado']).tolist()
    
    # Selector de riesgo en la barra lateral
    riesgo_seleccionado = st.sidebar.selectbox(
        "Seleccionar Riesgo:",
        opciones_riesgos
    )
    
    if riesgo_seleccionado:
        # Extraer el Num_Riesgo de la selección para encontrar la fila
        num_riesgo_buscado = riesgo_seleccionado.split(' - ')[0]
        
        # Obtener la fila (registro) de ese riesgo
        registro_riesgo = df_riesgos[df_riesgos['Num_Riesgo'] == num_riesgo_buscado].iloc[0]
        
        st.header(f"Ficha Seleccionada: {registro_riesgo['Riesgo_Identificado']}")
        
        # Botón de Generación y Descarga
        with st.spinner("Generando ficha..."):
            ficha_docx = generar_ficha_docx(registro_riesgo)
            
            st.download_button(
                label="📥 Descargar Ficha de Riesgo (DOCX)",
                data=ficha_docx,
                file_name=f"Ficha_Riesgo_{registro_riesgo['Num_Riesgo']}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            
            st.info("Presiona el botón de descarga para obtener el documento A4 generado.")
            
else:
    st.warning("No se encontraron datos de riesgos válidos en la matriz. Verifica el archivo.")
