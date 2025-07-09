"""
PDF Converter - Convertidor de documentos
Autor: Carlos Ramz
Descripción: Script para convertir entre PDF, Word y Excel
"""

import PyPDF2
import os
from docx import Document
import pdfplumber

def leer_pdf_mejorado(ruta_archivo):
    """
    Lee un archivo PDF usando pdfplumber para mejor extracción de texto
    
    Args:
        ruta_archivo (str): Ruta al archivo PDF
        
    Returns:
        str: Texto extraído del PDF
    """
    try:
        import pdfplumber
        
        texto_completo = ""
        
        with pdfplumber.open(ruta_archivo) as pdf:
            print(f"El PDF tiene {len(pdf.pages)} páginas")
            
            for num_pagina, pagina in enumerate(pdf.pages, 1):
                texto_pagina = pagina.extract_text()
                if texto_pagina:
                    texto_completo += texto_pagina + "\n"
                print(f"Página {num_pagina} procesada")
        
        return texto_completo
        
    except ImportError:
        print("❌ Error: pdfplumber no está instalado")
        return None
    except FileNotFoundError:
        print(f"Error: No se encontró el archivo {ruta_archivo}")
        return None
    except Exception as e:
        print(f"Error al leer el PDF: {str(e)}")
        return None

def pdf_a_word(ruta_pdf, directorio_salida="./convertidos"):
    """
    Convierte un PDF a documento Word
    
    Args:
        ruta_pdf (str): Ruta al archivo PDF
        directorio_salida (str): Directorio donde guardar el archivo Word
        
    Returns:
        str: Ruta del archivo Word creado, o None si hay error
    """
    try:
        from docx import Document
        
        # Extraer texto del PDF
        texto_pdf = leer_pdf_mejorado(ruta_pdf)
        if not texto_pdf:
            print("❌ No se pudo extraer texto del PDF")
            return None
        
        # Crear directorio si no existe
        if not os.path.exists(directorio_salida):
            os.makedirs(directorio_salida)
            print(f"📁 Carpeta creada: {directorio_salida}")
        
        # Generar nombre del archivo Word
        nombre_archivo = os.path.basename(ruta_pdf)
        nombre_sin_extension = os.path.splitext(nombre_archivo)[0]
        nombre_word = f"{nombre_sin_extension}.docx"
        ruta_word = os.path.join(directorio_salida, nombre_word)
        
        # Crear documento Word
        documento = Document()
        
        # Agregar el texto manteniendo formato original
        # Dividir por líneas y mantener estructura exacta
        lineas = texto_pdf.split('\n')
        
        for linea in lineas:
            # Mantener líneas vacías para conservar espaciado
            if not linea.strip():
                documento.add_paragraph("")  # Línea vacía
            else:
                documento.add_paragraph(linea)  # Línea con contenido tal como está
        
        # Guardar el documento
        documento.save(ruta_word)
        
        print(f"✅ Archivo Word creado: {ruta_word}")
        return ruta_word
        
    except ImportError:
        print("❌ Error: python-docx no está instalado")
        return None
    except Exception as e:
        print(f"❌ Error al crear archivo Word: {str(e)}")
        return None

def main():
    """Función principal del programa"""
    print("=== PDF Converter ===")
    print("Versión: 1.0.1 - Lector de PDF agregado")
    
    # Probar la función con un PDF
    nombre_pdf = "CODIGO DE ETICA Y CONDUCTA.pdf"
    
    print(f"Buscando archivo: {nombre_pdf}")
    print(f"Directorio actual: {os.getcwd()}")
    
    if os.path.exists(nombre_pdf):
        print(f"✅ Archivo encontrado: {nombre_pdf}")
        print(f"Tamaño del archivo: {os.path.getsize(nombre_pdf)} bytes")
        
        print(f"\nProbando leer el archivo...")
        texto = leer_pdf_mejorado(nombre_pdf)
        
        if texto:
           print("\n--- TEXTO EXTRAÍDO ---")
           print(texto[:200] + "..." if len(texto) > 200 else texto)
    
    # Convertir a Word
           print("\n🔄 Convirtiendo a Word...")
           archivo_word = pdf_a_word(nombre_pdf)
           if archivo_word:
             print(f"✅ Conversión exitosa: {archivo_word}")
        else:
           print("❌ No se pudo extraer texto del PDF")
    else:
        print(f"❌ No se encontró el archivo {nombre_pdf}")
        print("Archivos PDF disponibles:")
        for archivo in os.listdir('.'):
            if archivo.endswith('.pdf'):
                print(f"  - {archivo}")

if __name__ == "__main__":
    main()