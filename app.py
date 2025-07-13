import pandas as pd
import matplotlib.pyplot as plt
from fpdf import FPDF
import streamlit as st
import io

def procesar_archivos(archivo1, archivo2):
    # Leer archivos Excel
    df1 = pd.read_excel(archivo1)
    df2 = pd.read_excel(archivo2)
    
    # Procesamiento de ejemplo: unir los dataframes
    df_final = pd.concat([df1, df2])
    
    # Crear un gráfico simple
    plt.figure()
    if 'valor' in df_final.columns:  # Asegurarse que la columna existe
        df_final['valor'].plot(kind='bar')
        plt.title('Resumen de Valores')
        img_buffer = io.BytesIO()
        plt.savefig(img_buffer, format='png')
        plt.close()
        img_buffer.seek(0)
    else:
        img_buffer = None
    
    return df_final, img_buffer

def crear_pdf(df, img_buffer=None):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    
    # Agregar título
    pdf.cell(200, 10, txt="Reporte Generado", ln=1, align='C')
    
    # Agregar imagen si existe
    if img_buffer:
        pdf.image(img_buffer, x=10, y=30, w=180)
    
    # Agregar tabla de datos
    pdf.set_y(100)
    pdf.cell(200, 10, txt="Datos Resumidos:", ln=1)
    
    # Crear tabla simple (para tablas más complejas necesitarías más lógica)
    for index, row in df.head(10).iterrows():
        pdf.cell(200, 10, txt=str(row.to_dict()), ln=1)
    
    return pdf

def main():
    st.title("Procesador de Archivos Excel")
    st.write("Sube dos archivos Excel para procesar y generar un PDF")
    
    # Subida de archivos
    archivo1 = st.file_uploader("Sube el primer archivo Excel", type=['xlsx', 'xls'])
    archivo2 = st.file_uploader("Sube el segundo archivo Excel", type=['xlsx', 'xls'])
    
    if archivo1 and archivo2:
        # Procesar archivos
        df, img_buffer = procesar_archivos(archivo1, archivo2)
        
        # Mostrar vista previa
        st.write("Vista previa de los datos procesados:")
        st.dataframe(df.head())
        
        if img_buffer:
            st.image(img_buffer, caption='Gráfico generado')
        
        # Crear PDF
        pdf = crear_pdf(df, img_buffer)
        
        # Descargar PDF
        pdf_output = io.BytesIO()
        pdf.output(pdf_output)
        pdf_output.seek(0)
        
        st.download_button(
            label="Descargar PDF",
            data=pdf_output,
            file_name="reporte_generado.pdf",
            mime="application/pdf"
        )

if __name__ == "__main__":
    main()