import streamlit as st
import pandas as pd
from fpdf import FPDF
import base64
from datetime import datetime
import os
import zipfile
import tempfile
import io
import logging
import re
from PIL import Image
import numpy as np
import atexit
import sys

# CONFIGURACIÓN COMPATIBLE CON STREAMLIT CLOUD
st.set_page_config(
    page_title="Sistema de Vales de Resguardo DGA",
    layout="wide",
    page_icon="🏛️",
    initial_sidebar_state="expanded"
)

# Evitar errores de caché
@st.cache_resource(show_spinner=False)
def load_data():
    return None

# Configurar logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Configuración de la página
st.set_page_config(
    page_title="Sistema de Vales de Resguardo - Área de Activo Fijo DGA", 
    layout="wide",
    page_icon="🏛️"
)

# Inicializar session_state para manejo de estado
if 'selected_employee' not in st.session_state:
    st.session_state.selected_employee = None
if 'df_processed' not in st.session_state:
    st.session_state.df_processed = None
if 'empleados_list' not in st.session_state:
    st.session_state.empleados_list = []
if 'file_uploaded' not in st.session_state:
    st.session_state.file_uploaded = False

class PDF(FPDF):
    def __init__(self):
        super().__init__()
    
    def header(self):
        # Logo horizontal en TODAS las páginas
        try:
            self.image("LOGOS_VALE.png", x=10, y=8, w=190)
        except:
            # Si no encuentra el logo, poner título
            self.set_font("Arial", 'B', 14)
            self.cell(0, 5, "VALES DE RESGUARDO INTERNO 2025", 0, 1, 'C')
        self.ln(12)
    
    def footer(self):
        # Pie de página con IMAGEN en TODAS las páginas
        try:
            # Usar la imagen Pie_vale.png en lugar de texto
            self.image("Pie_vale.png", x=10, y=self.h - 30, w=190)
        except:
            # Fallback a texto si la imagen no existe
            self.set_y(-25)
            self.set_font("Arial", 'I', 8)
            self.set_text_color(100, 100, 100)
            self.cell(0, 5, "Agustin Delgado No. 58, Col. Transito, CP. 06820, Alcaldia Cuauhtemoc, CDMX.", 0, 1, 'C')
            self.cell(0, 5, "Tel: (55) 3601 7100", 0, 1, 'C')
            self.cell(0, 5, f"Página {self.page_no()}", 0, 0, 'C')

def procesar_codigo_qr(codigo_qr):
    """
    Procesa el código QR para extraer:
    - No. SEP
    - Número de inventario
    - Descripción
    - Valor
    """
    # Inicializar valores por defecto
    no_sep = ""
    numero_inventario = ""
    descripcion = ""
    valor = 0
    
    try:
        # MEJORA: Validación de código QR vacío
        if not codigo_qr or codigo_qr.strip() == "":
            return no_sep, numero_inventario, descripcion, valor
            
        # MEJORA: Limpieza del código QR
        codigo_qr = codigo_qr.strip()
        
        # Intentar diferentes patrones de códigos QR
        # Patrón 1: Formato estándar con separadores |
        if "|" in codigo_qr:
            partes = codigo_qr.split("|")
            if len(partes) >= 4:
                no_sep = partes[0].strip()
                numero_inventario = partes[1].strip()
                descripcion = partes[2].strip()
                # Extraer valor numérico (eliminar caracteres no numéricos)
                valor_str = re.sub(r'[^\d.]', '', partes[3])
                valor = float(valor_str) if valor_str else 0
        
        # Patrón 2: Formato con otros separadores
        elif ";" in codigo_qr:
            partes = codigo_qr.split(";")
            if len(partes) >= 4:
                no_sep = partes[0].strip()
                numero_inventario = partes[1].strip()
                descripcion = partes[2].strip()
                valor_str = re.sub(r'[^\d.]', '', partes[3])
                valor = float(valor_str) if valor_str else 0
                
        # MEJORA: Patrón 3 - Formato con comas
        elif "," in codigo_qr:
            partes = codigo_qr.split(",")
            if len(partes) >= 4:
                no_sep = partes[0].strip()
                numero_inventario = partes[1].strip()
                descripcion = partes[2].strip()
                valor_str = re.sub(r'[^\d.]', '', partes[3])
                valor = float(valor_str) if valor_str else 0
                
        # Si no coincide con ningún patrón conocido, usar el código completo como descripción
        else:
            descripcion = codigo_qr[:50]  # Limitar a 50 caracteres
            
    except Exception as e:
        logger.error(f"Error al procesar código QR '{codigo_qr}': {str(e)}")
    
    return no_sep, numero_inventario, descripcion, valor

def procesar_dataframe_con_qr(df):
    """
    Procesa el DataFrame que contiene la columna QR y completa las columnas correspondientes
    """
    # MEJORA: Validación de DataFrame vacío
    if df is None or df.empty:
        return df
        
    # Verificar si existe la columna QR
    if 'QR' not in df.columns:
        return df
    
    # Crear copia del DataFrame para no modificar el original
    df_procesado = df.copy()
    
    # Asegurar que existan las columnas necesarias
    columnas_necesarias = ['No. SEP', 'NUMERO DE INVVENTARIO', 'DESCRIPCION', 'VALOR']
    for col in columnas_necesarias:
        if col not in df_procesado.columns:
            df_procesado[col] = ""
            if col == 'VALOR':
                df_procesado[col] = 0.0
    
    # Procesar cada fila con código QR
    for idx, row in df_procesado.iterrows():
        qr_value = str(row['QR']) if pd.notna(row['QR']) else ""
        
        # MEJORA: Validación más robusta de valores QR
        if qr_value and qr_value.strip() not in ["", "nan", "None"]:
            # Verificar si las columnas destino ya tienen datos
            no_sep_existente = str(row['No. SEP']) if pd.notna(row['No. SEP']) else ""
            inventario_existente = str(row['NUMERO DE INVVENTARIO']) if pd.notna(row['NUMERO DE INVVENTARIO']) else ""
            
            # Solo procesar QR si las columnas destino están vacías
            if not no_sep_existente.strip() and not inventario_existente.strip():
                no_sep, numero_inv, descripcion, valor = procesar_codigo_qr(qr_value)
                
                # Actualizar las columnas correspondientes
                if no_sep:
                    df_procesado.at[idx, 'No. SEP'] = no_sep
                if numero_inv:
                    df_procesado.at[idx, 'NUMERO DE INVVENTARIO'] = numero_inv
                if descripcion and (not row['DESCRIPCION'] or pd.isna(row['DESCRIPCION']) or str(row['DESCRIPCION']).strip() == ""):
                    df_procesado.at[idx, 'DESCRIPCION'] = descripcion
                if valor > 0 and (not row['VALOR'] or row['VALOR'] == 0):
                    df_procesado.at[idx, 'VALOR'] = valor
    
    return df_procesado

def generar_vale_pdf(empleado, datos_empleado, inventario_empleado):
    """Genera el contenido PDF y lo retorna como bytes"""
    try:
        # MEJORA: Validación de datos de entrada
        if inventario_empleado is None or inventario_empleado.empty:
            raise Exception("No hay datos de inventario para generar el vale")
            
        if datos_empleado is None or datos_empleado.empty:
            raise Exception("No hay datos del empleado")
        
        # Crear PDF con formato oficial
        pdf = PDF()
        pdf.add_page()
        pdf.set_auto_page_break(auto=True, margin=30)
        pdf.set_margins(left=10, top=25, right=10)
        
        # Título principal
        pdf.set_font("Arial", 'B', 16)
        pdf.cell(0, 8, "VALE DE RESGUARDO 2025", 0, 1, 'C')
        pdf.ln(3)
        
        # Sección de información del responsable - FORMATO MEJORADO CON 2 COLUMNAS PEGADAS
        y_start = pdf.get_y()
        pdf.rect(10, y_start, 190, 42)
        
        # Título dentro del marco
        pdf.set_font("Arial", 'B', 11)
        pdf.set_xy(10, y_start + 3)
        pdf.cell(190, 6, "INFORMACIÓN DEL RESPONSABLE", 0, 1, 'C')
        pdf.line(10, y_start + 9, 200, y_start + 9)  # Línea bajo el título

        # Línea vertical divisoria más a la derecha (120 en lugar de 130)
        pdf.line(120, y_start + 9, 120, y_start + 42)

        # Columna izquierda - Información más pegada
        pdf.set_font("Arial", 'B', 8)
        pdf.set_xy(12, y_start + 12)
        pdf.cell(35, 5, "NOMBRE COMPLETO:", 0, 0)
        pdf.set_font("Arial", '', 8)
        nombre = f"{datos_empleado.get('NOMBRE', '')}"
        # Ajustar nombre largo
        if len(nombre) > 35:
            nombre = nombre[:32] + "..."
        pdf.cell(85, 5, nombre, 0, 1)
        
        pdf.set_font("Arial", 'B', 8)
        pdf.set_xy(12, y_start + 17)
        pdf.cell(15, 5, "CURP:", 0, 0)
        pdf.set_font("Arial", '', 8)
        curp = f"{datos_empleado.get('CURP', '')}"
        pdf.cell(93, 5, curp, 0, 1)
        
        pdf.set_font("Arial", 'B', 8)
        pdf.set_xy(12, y_start + 22)
        pdf.cell(12, 5, "RFC:", 0, 0)
        pdf.set_font("Arial", '', 8)
        rfc = f"{datos_empleado.get('RFC', '')}"
        pdf.cell(96, 5, rfc, 0, 1)
        
        # Área de adscripción - Más compacta y pegada
        pdf.set_font("Arial", 'B', 8)
        pdf.set_xy(12, y_start + 27)
        pdf.cell(42, 5, "AREA DE ADSCRIPCION:", 0, 0)
        pdf.set_font("Arial", '', 8)
        area = f"{datos_empleado.get('AREA O DEPARTAMENTO', '')}"
        
        # Verificar si el área es demasiado larga para una línea
        if len(area) > 20:
            # Dividir en two líneas
            mitad = len(area) // 2
            espacio_idx = area.rfind(' ', 0, mitad)
            if espacio_idx == -1:
                espacio_idx = mitad
                
            linea1 = area[:espacio_idx].strip()
            linea2 = area[espacio_idx:].strip()
            
            pdf.cell(66, 5, linea1, 0, 1)
            pdf.set_x(54)
            pdf.cell(66, 5, linea2, 0, 1)
        else:
            pdf.cell(66, 5, area, 0, 1)
        
        # Campo EDIFICIO
        pdf.set_font("Arial", 'B', 8)
        pdf.set_xy(12, y_start + 37)
        pdf.cell(25, 5, "EDIFICIO:", 0, 0)
        pdf.set_font("Arial", '', 8)
        edificio = f"{datos_empleado.get('EDIFICIO', '')}"
        pdf.cell(83, 5, edificio, 0, 1)
        
        # Columna derecha - Información más pegada y ajustada
        pdf.set_font("Arial", 'B', 8)
        pdf.set_xy(122, y_start + 12)
        pdf.cell(38, 5, "CENTRO DE TRABAJO:", 0, 0)
        pdf.set_font("Arial", '', 8)
        ct = f"{datos_empleado.get('CT', 'COMISIONADO')}"
        pdf.cell(40, 5, ct, 0, 1)
        
        pdf.set_font("Arial", 'B', 8)
        pdf.set_xy(122, y_start + 17)
        pdf.cell(15, 5, "PISO:", 0, 0)
        pdf.set_font("Arial", '', 8)
        piso = f"{datos_empleado.get('PISO', '')}"
        pdf.cell(63, 5, piso, 0, 1)
        
        pdf.set_font("Arial", 'B', 8)
        pdf.set_xy(122, y_start + 22)
        pdf.cell(48, 5, "FECHA LEVANTAMIENTO:", 0, 0)
        pdf.set_font("Arial", '', 8)
        pdf.cell(30, 5, f"{datetime.now().strftime('%d/%m/%Y')}", 0, 1)
        
        pdf.set_font("Arial", 'B', 8)
        pdf.set_xy(122, y_start + 27)
        pdf.cell(33, 5, "TOTAL MUEBLES:", 0, 0)
        pdf.set_font("Arial", '', 8)
        pdf.cell(45, 5, f"{len(inventario_empleado)}", 0, 1)
        
        pdf.set_y(y_start + 43)
        
        # Tabla de bienes - ANCHOS AJUSTADOS PARA OBSERVACIONES
        pdf.set_font("Arial", 'B', 12)
        pdf.cell(0, 8, "INVENTARIO OFICIAL DE BIENES MUEBLES", 0, 1, 'C')
        pdf.ln(2)
        
        # Encabezados de tabla con anchos OPTIMIZADOS para observaciones
        col_widths = [8, 15, 40, 65, 20, 42]
        headers = ["No.", "No. SEP", "NO. INVENTARIO", "DESCRIPCION DEL BIEN", "VALOR", "OBSERV."]
        
        def draw_headers():
            pdf.set_draw_color(0, 51, 102)
            pdf.set_fill_color(230, 240, 250)
            pdf.set_font("Arial", 'B', 7)
            for i, header in enumerate(headers):
                pdf.cell(col_widths[i], 7, header, 1, 0, 'C', fill=True)
            pdf.ln()
            pdf.set_font("Arial", '', 7)
        
        draw_headers()
        
        total_valor = 0
        for idx, (_, item) in enumerate(inventario_empleado.iterrows(), 1):
            if pdf.get_y() > 220:
                pdf.add_page()
                draw_headers()
            
            fill_color = (245, 245, 245) if idx % 2 == 0 else (255, 255, 255)
            pdf.set_fill_color(*fill_color)
            
            pdf.cell(col_widths[0], 6, str(idx), 1, 0, 'C', fill=True)
            
            no_sep = str(item.get('No. SEP', ''))
            if no_sep.lower() == 'nan' or no_sep == '':
                no_sep = ''
            pdf.cell(col_widths[1], 6, no_sep[:8], 1, 0, 'C', fill=True)
            
            no_inv = str(item.get('NUMERO DE INVVENTARIO', ''))
            if no_inv.lower() == 'nan' or no_inv == '':
                no_inv = ''
            
            if len(no_inv) > 20:
                pdf.set_font("Arial", '', 6)
                pdf.cell(col_widths[2], 6, no_inv[:25], 1, 0, 'C', fill=True)
                pdf.set_font("Arial", '', 7)
            else:
                pdf.cell(col_widths[2], 6, no_inv, 1, 0, 'C', fill=True)
            
            descripcion = str(item.get('DESCRIPCION', ''))
            if len(descripcion) > 35:
                descripcion = descripcion[:32] + "..."
            pdf.cell(col_widths[3], 6, descripcion, 1, 0, 'L', fill=True)
            
            valor = item.get('VALOR', 0)
            pdf.cell(col_widths[4], 6, f"${valor:.2f}", 1, 0, 'R', fill=True)
            
            observ = str(item.get('OBSERVACIONES', ''))
            if no_sep == '' and no_inv == '' and valor == 0:
                observ = "EN PROCESO DE ALTA"
            elif observ.lower() == 'nan':
                observ = ''
            
            if len(observ) > 25:
                pdf.set_font("Arial", '', 6)
                pdf.cell(col_widths[5], 6, observ[:35], 1, 0, 'C', fill=True)
                pdf.set_font("Arial", '', 7)
            else:
                pdf.cell(col_widths[5], 6, observ, 1, 0, 'C', fill=True)
            
            pdf.ln()
            total_valor += valor
        
        # Total
        pdf.set_font("Arial", 'B', 8)
        pdf.set_fill_color(220, 230, 240)
        pdf.cell(sum(col_widths[:4]), 7, "VALOR TOTAL DEL INVENTARIO:", 1, 0, 'R', fill=True)
        pdf.cell(col_widths[4], 7, f"${total_valor:.2f}", 1, 0, 'R', fill=True)
        pdf.cell(col_widths[5], 7, "", 1, 1, fill=True)
        
        pdf.ln(8)
        
        # Condiciones de resguardo
        pdf.set_font("Arial", 'B', 9)
        pdf.cell(0, 6, "CONDICIONES DE RESGUARDO", 0, 1, 'C')
        
        pdf.set_font("Arial", '', 7)
        condiciones = [
            "- Los bienes muebles se entregan bajo custodia del resguardante.",
            "- El resguardante es responsable conforme to the Ley General de Bienes Nacionales.",
            "- Los bienes son propiedad institucional, no del personal.",
            "- Debe notificar cualquier movimiento, pérdida or daño inmediatamente.",
            "- En caso de cambio de adscripción or renuncia, debe devolver los bienes."
        ]
        
        for condicion in condiciones:
            pdf.cell(0, 3.5, condicion, 0, 1)
            
        pdf.ln(5)
        
        # Sección de firmas
        y_firmas = pdf.get_y()
        pdf.rect(10, y_firmas, 190, 40)
        
        # Título de la sección
        pdf.set_font("Arial", 'B', 9)
        pdf.set_xy(10, y_firmas + 3)
        pdf.cell(190, 6, "FIRMAS", 0, 1, 'C')
        pdf.line(10, y_firmas + 9, 200, y_firmas + 9)
        
        # Línea vertical divisoria en el centro
        pdf.line(105, y_firmas + 9, 105, y_firmas + 40)
        
        # RESGUARDANTE (LADO IZQUIERDO)
        pdf.set_font("Arial", 'B', 9)
        pdf.set_xy(10, y_firmas + 12)
        pdf.cell(95, 5, "RESGUARDANTE", 0, 0, 'C')
        
        # Espacio para firma
        pdf.set_font("Arial", '', 8)
        pdf.set_xy(25, y_firmas + 22)
        pdf.cell(65, 10, "_________________________", 0, 0, 'C')
        
        nombre_resguardante = f"{datos_empleado.get('NOMBRE', '')}"
        if len(nombre_resguardante) > 35:
            nombre_resguardante = nombre_resguardante[:32] + "..."
        
        pdf.set_xy(10, y_firmas + 32)
        pdf.cell(95, 5, nombre_resguardante, 0, 0, 'C')
        
        # AUTORIZA (LADO DERECHO)
        pdf.set_font("Arial", 'B', 9)
        pdf.set_xy(105, y_firmas + 12)
        pdf.cell(95, 5, "AUTORIZA", 0, 0, 'C')
        
        # Espacio para firma
        pdf.set_font("Arial", '', 8)
        pdf.set_xy(120, y_firmas + 22)
        pdf.cell(65, 10, "_________________________", 0, 0, 'C')
        
        pdf.set_xy(105, y_firmas + 32)
        pdf.cell(95, 5, "EDNA SANCHEZ MARTINEZ", 0, 0, 'C')
        
        pdf.set_font("Arial", 'I', 7)
        pdf.set_xy(105, y_firmas + 37)
        pdf.cell(95, 4, "Coordinadora Administrativa", 0, 0, 'C')
        
        # Retornar el PDF como bytes
        return pdf.output(dest='S').encode('latin1')
        
    except Exception as e:
        logger.error(f"Error al generar el PDF para {empleado}: {str(e)}")
        raise Exception(f"Error al generar el PDF para {empleado}: {str(e)}")

def generar_vale_individual(empleado, df):
    """Genera y descarga un vale de resguardo individual"""
    try:
        # MEJORA: Validación de empleado existente
        if empleado not in df['NOMBRE'].values:
            st.error(f"Empleado '{empleado}' no encontrado en los datos")
            return
            
        # Filtrar datos del empleado
        datos_empleado = df[df['NOMBRE'] == empleado].iloc[0]
        inventario_empleado = df[df['NOMBRE'] == empleado]
        
        # MEJORA: Validación de inventario vacío
        if inventario_empleado.empty:
            st.warning(f"El empleado {empleado} no tiene artículos en el inventario")
            return
        
        # Generar PDF
        pdf_bytes = generar_vale_pdf(empleado, datos_empleado, inventario_empleado)
        
        # Descargar
        filename = f"Vale_Resguardo_{empleado.replace(' ', '_')}.pdf"
        
        st.download_button(
            label="📥 Descargar Vale Oficial",
            data=pdf_bytes,
            file_name=filename,
            mime="application/pdf",
            type="primary",
            key=f"download_{empleado}"
        )
        
        st.success(f"✅ Vale oficial generado para {empleado}")
        
    except Exception as e:
        st.error(f"Error al generar el vale: {str(e)}")

def generar_todos_los_vales(df):
    """Genera todos los vales y retorna un archivo ZIP"""
    try:
        # MEJORA: Validación de DataFrame vacío
        if df is None or df.empty:
            st.error("No hay datos para generar vales")
            return None
            
        empleados = df['NOMBRE'].unique()
        zip_buffer = io.BytesIO()
        
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for empleado in empleados:
                try:
                    datos_empleado = df[df['NOMBRE'] == empleado].iloc[0]
                    inventario_empleado = df[df['NOMBRE'] == empleado]
                    
                    # MEJORA: Saltar empleados sin inventario
                    if inventario_empleado.empty:
                        continue
                        
                    pdf_bytes = generar_vale_pdf(empleado, datos_empleado, inventario_empleado)
                    
                    filename = f"Vale_Resguardo_{empleado.replace(' ', '_')}.pdf"
                    zipf.writestr(filename, pdf_bytes)
                    
                except Exception as e:
                    st.warning(f"⚠️ Error con {empleado}: {str(e)}")
                    continue
        
        zip_buffer.seek(0)
        return zip_buffer.getvalue()
        
    except Exception as e:
        st.error(f"Error al generar el archivo ZIP: {str(e)}")
        return None

def procesar_archivo_excel(uploaded_file):
    """Procesa el archivo Excel y devuelve un DataFrame limpio"""
    try:
        # MEJORA: Validación de archivo vacío
        if uploaded_file is None:
            return None
            
        # Leer el archivo Excel
        df = pd.read_excel(uploaded_file)
        
        # MEJORA: Validación de DataFrame vacío
        if df is None or df.empty:
            st.error("El archivo Excel está vacío")
            return None
        
        # Validar columnas mínimas requeridas
        columnas_requeridas = ['NOMBRE', 'DESCRIPCION', 'VALOR']
        for col in columnas_requeridas:
            if col not in df.columns:
                st.error(f"El archivo debe contener la columna: {col}")
                return None
        
        # Limpiar datos
        df = df.dropna(subset=['NOMBRE'])
        df['VALOR'] = pd.to_numeric(df['VALOR'], errors='coerce').fillna(0)
        
        # Limpiar nombres de columnas (eliminar espacios extra)
        df.columns = df.columns.str.strip()
        
        # Procesar códigos QR si existe la columna
        if 'QR' in df.columns:
            df = procesar_dataframe_con_qr(df)
        
        return df
        
    except Exception as e:
        st.error(f"Error al procesar el archivo: {str(e)}")
        return None

def mostrar_estadisticas(df):
    """Muestra estadísticas del inventario"""
    # MEJORA: Validación de DataFrame
    if df is None or df.empty:
        st.warning("No hay datos para mostrar estadísticas")
        return
        
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Total de empleados", len(df['NOMBRE'].unique()))
    with col2:
        st.metric("Total de artículos", len(df))
    with col3:
        st.metric("Valor total inventario", f"${df['VALOR'].sum():,.2f}")
    with col4:
        valor_promedio = df['VALOR'].mean() if len(df) > 0 else 0
        st.metric("Valor promedio por artículo", f"${valor_promedio:,.2f}")

def mostrar_encabezado_web():
    """Muestra el encabezado de la página web"""
    try:
        # Intenta cargar la imagen del encabezado
        st.image("ENCABEZADO_WEB.png", use_container_width=True)
    except:
        # Si no encuentra la imagen, muestra un encabezado alternativo
        st.markdown("""
        <div style="background-color: #0c4e94; padding: 20px; border-radius: 10px; text-align: center; margin-bottom: 20px;">
            <h1 style="color: white; margin: 0;">Sistema de Vales de Resguardo</h1>
            <p style="color: white; margin: 5px 0 0 0;">Área de Activo Fijo - Dirección General de Administración</p>
        </div>
        """, unsafe_allow_html=True)

def mostrar_pie_web():
    """Muestra el pie de página de la web"""
    try:
        # MEJORA: Usar la imagen Pie_vale.png para el pie de página web también
        st.image("Pie_vale.png", use_container_width=True)
    except:
        # Si no encuentra la imagen, muestra un pie de página alternativo
        st.markdown("""
        <div style="background-color: #f0f2f6; padding: 15px; text-align: center; margin-top: 30px; border-top: 2px solid #0c4e94;">
            <p style="color: #555555; margin: 0;">Sistema desarrollado por el Área de Activo Fijo - DGA 2025</p>
        </div>
        """, unsafe_allow_html=True)

def main():
    """Función principal de la aplicación"""
    # Mostrar encabezado
    mostrar_encabezado_web()
    
    st.title("🏛️ Sistema de Generación de Vales de Resguardo")
    
    # Cargar archivo Excel
    uploaded_file = st.file_uploader("Cargar archivo de inventario Excel", type=["xlsx", "xls"])
    
    if uploaded_file:
        try:
            # Procesar el archivo
            if st.session_state.df_processed is None:
                df = procesar_archivo_excel(uploaded_file)
                if df is None:
                    return
                st.session_state.df_processed = df
                st.session_state.empleados_list = sorted(df['NOMBRE'].unique())
                # Establecer el primer empleado como selección predeterminada
                if st.session_state.selected_employee is None and st.session_state.empleados_list:
                    st.session_state.selected_employee = st.session_state.empleados_list[0]
                st.session_state.file_uploaded = True
            else:
                df = st.session_state.df_processed
                
            # MEJORA: Validación de DataFrame procesado
            if df is None or df.empty:
                st.error("No se pudieron procesar los datos del archivo")
                return
                
            # Mostrar estadísticas generales
            mostrar_estadisticas(df)
            
            # Seleccionar empleado
            empleados = st.session_state.empleados_list
            
            # MEJORA: Validación de lista de empleados
            if not empleados:
                st.warning("No se encontraron empleados en el archivo")
                return
                
            # Obtener el índice actual para mantener la selección
            current_index = 0
            if st.session_state.selected_employee in empleados:
                current_index = empleados.index(st.session_state.selected_employee)
            
            # Widget para seleccionar empleado
            selected_employee = st.selectbox(
                "Seleccionar empleado", 
                options=empleados,
                index=current_index,
                key='employee_selector'
            )
            
            # Actualizar session_state con la selección actual
            st.session_state.selected_employee = selected_employee
            
            # Filtrar datos del empleado seleccionado
            datos_empleado = df[df['NOMBRE'] == selected_employee]
            
            # MEJORA: Validación de datos del empleado
            if datos_empleado.empty:
                st.warning(f"No se encontraron datos para el empleado: {selected_employee}")
                return
            
            # Botones de acción
            col1, col2 = st.columns(2)
            with col1:
                if st.button("📄 Generar Vale Individual", type="primary", use_container_width=True):
                    generar_vale_individual(selected_employee, df)
            
            with col2:
                if st.button("📚 Generar Todos los Vales", use_container_width=True):
                    with st.spinner("🔄 Generando todos los vales, por favor espere..."):
                        zip_data = generar_todos_los_vales(df)
                        
                        if zip_data:
                            st.download_button(
                                label="📦 Descargar Todos los Vales (ZIP)",
                                data=zip_data,
                                file_name="Todos_Los_Vales_de_Resguardo.zip",
                                mime="application/zip",
                                type="primary",
                                key="download_all"
                            )
                            st.success("✅ Todos los vales han sido generados exitosamente")
            
            # Vista previa de datos
            columnas_a_mostrar = ['DESCRIPCION', 'NUMERO DE INVVENTARIO', 'No. SEP', 'VALOR', 'OBSERVACIONES']
            if 'QR' in datos_empleado.columns:
                columnas_a_mostrar = ['QR'] + columnas_a_mostrar
                
            columnas_disponibles = [col for col in columnas_a_mostrar if col in datos_empleado.columns]
            
            st.dataframe(
                datos_empleado[columnas_disponibles],
                height=300,
                use_container_width=True
            )
                    
        except Exception as e:
            st.error(f"Error al procesar el archivo: {str(e)}")
    else:
        # Si no hay archivo cargado, resetear el estado
        if st.session_state.file_uploaded:
            st.session_state.df_processed = None
            st.session_state.empleados_list = []
            st.session_state.selected_employee = None
            st.session_state.file_uploaded = False
            
        st.info("📁 Por favor, carga un archivo Excel para comenzar")
        
        # Mostrar información de ejemplo cuando no hay archivo cargado
        with st.expander("💡 Ver ejemplo de estructura del archivo"):
            st.markdown("""
            **Estructura recomendada del archivo Excel:**
            
            | NOMBRE | CURP | RFC | AREA O DEPARTAMENTO | EDIFICIO | QR | No. SEP | NUMERO DE INVVENTARIO | DESCRIPCION | VALOR | OBSERVACIONES | CT | PISO |
            |--------|------|-----|---------------------|----------|----|---------|----------------------|-------------|-------|--------------|----|------|
            | JUAN PEREZ LOPEZ | PELJ800101HDFRPN01 | PELJ800101ABC | RECURSOS HUMANOS | EDIFICIO A | 12345\|67890\|ESCRITORIO OFICINA\|1500.00 | 12345 | 67890 | ESCRITORIO OFICINA | 1500.00 | BUEN ESTADO | OFICINAS CENTRALES | 2 |
            | MARIA GARCIA HERNANDEZ | GAHM750512MDFRRR02 | GAHM750512DEF | CONTABILIDAD | EDIFICIO B | 54321\|09876\|SILLA EJECUTIVA\|2500.50 | 54321 | 09876 | SILLA EJECUTIVA | 2500.50 | NUEVO | OFICINAS CENTRALES | 3 |
            """)
    
    # Información adicional en sidebar
    st.sidebar.markdown("### ℹ️ Información del Sistema")
    st.sidebar.info("""
    **Sistema de Vales de Resguardo**
    
    🔸 Genera vales en formato PDF oficial  
    🔸 Procesa automáticamente códigos QR  
    🔸 Calcula totales automáticamente  
    
    **Instrucciones:**
    1. Carga tu archivo Excel de inventario
    2. Selecciona un empleado
    3. Genera y descarga el vale
    """)

    # Créditos
    st.sidebar.markdown("---")
    st.sidebar.markdown("**Desarrollado por:**")
    st.sidebar.markdown("Área de Activo Fijo")
    st.sidebar.markdown("DGA 2025")
    st.sidebar.markdown("Pedro Álvaro Pérez Rodríguez")
    
    # Mostrar pie de página
    mostrar_pie_web()

if __name__ == "__main__":

    main()
