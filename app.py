import pandas as pd
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import streamlit as st
import io 

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
    """Agrega una sección formal usando una tabla de dos columnas (Título del Campo | Valor del Campo)."""
    document.add_heading(titulo, level=2)
    
    # Crear una tabla de 2 columnas
    tabla = document.add_table(rows=len(datos_dict), cols=2)
    tabla.style = 'Table Grid'
    
    # Establecer ancho para el título (Columna 0)
    tabla.columns[0].width = Inches(2)
    
    i = 0
    for key, value in datos_dict.items():
        row_cells = tabla.rows[i].cells
        
        # Título del Campo (negrita)
        row_cells[0].paragraphs[0].add_run(f'{key}').bold = True
        
        # Valor del Campo (texto normal)
        row_cells[1].paragraphs[0].add_run(str(value))
        i += 1
    document.add_paragraph()

def generar_ficha_docx(datos_riesgo):
    """
    Genera la ficha A4 y devuelve el documento como un objeto BytesIO 
    para poder ser descargado en Streamlit.
    """
    document = Document()
    
    # --- Configuración Estilística A4 ---
    section = document.sections[0]
    section.top_margin, section.bottom_margin = Inches(0.5), Inches(0.5)
    section.left_margin, section.right_margin = Inches(0.75), Inches(0.75)
    
    # --- TÍTULO PRINCIPAL DE LA FICHA ---
    document.add_heading(f'FICHA DE GESTIÓN DE RIESGO N° {datos_riesgo["Num_Riesgo"]}', level=0)
    document.add_paragraph(f'Versión: {datos_riesgo["Version"]} | Última Revisión: {datos_riesgo["Ultima_Revision"]}')
    document.add_paragraph('---') 

    # ------------------------------------------------------------------
    # 1) - IDENTIFICACIÓN DEL RIESGO
    # ------------------------------------------------------------------
    identificacion_data = {
        'Riesgo Identificado': datos_riesgo['Riesgo_Identificado'],
        'Entorno de Control': datos_riesgo['Entorno_Control'],
        'Origen / Área Responsable': datos_riesgo['Origen_Area'],
        'Proceso o Documento': datos_riesgo['Proceso_Documento']
    }
    agregar_se
