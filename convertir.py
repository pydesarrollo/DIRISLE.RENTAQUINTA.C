import os
from fpdf import FPDF
from fpdf.enums import XPos, YPos
from PIL import Image

CARPETA_PRTS = "archivos_prt"
CARPETA_PDFS = "pdfs"
TEXTO_DIVISOR = "MINISTERIO DE SALUD"
LOGO_PATH = "img/Logominsa.jpg"
LINEAS_CENTRADAS = 6
TEXTO_BUSCAR = "NUMERO DE PLAZA     : "
PREFIJO_BUSCAR = "POR IMPUESTO"

# Configuración profesional para formato A4
MARGEN_IZQ = 10
MARGEN_DER = 10
ANCHO_UTIL = 210 - MARGEN_IZQ - MARGEN_DER  # 190mm
FUENTE_SIZE = 6  # Tamaño de fuente óptimo
ALTURA_LINEA = 2.5  # Interlineado mejorado

def obtener_numero_plaza(seccion):
    for linea in seccion:
        if linea and TEXTO_BUSCAR in linea:
            inicio = linea.find(TEXTO_BUSCAR) + len(TEXTO_BUSCAR)
            return linea[inicio:inicio+6].strip()
    return None

def obtener_sufijo_impuesto(seccion):
    for elemento in seccion:
        if elemento and isinstance(elemento, str) and elemento.startswith(PREFIJO_BUSCAR):
            linea_limpia = elemento.strip()
            return linea_limpia[-4:] if len(linea_limpia) >= 4 else linea_limpia.ljust(4, '_')
    return ""

def procesar_archivo(ruta_prt, nombre_base):
    try:
        with open(ruta_prt, 'rb') as f:
            contenido = None
            for encoding in ['cp1252', 'latin-1', 'utf-8']:
                try:
                    contenido = f.read().decode(encoding)
                    break
                except UnicodeDecodeError:
                    continue
            if not contenido:
                raise ValueError("Error de decodificación")

        lineas = contenido.splitlines()
        sufijo_global = "0000"  # Valor por defecto
        if len(lineas) >= 4:
            cuarta_linea = lineas[4].strip()  # Quinta línea (índice 3)
            
            if len(cuarta_linea) >= 4:
                sufijo_global = cuarta_linea[-4:]  # Últimos 4 caracteres
        secciones = []
        seccion_actual = []
        patron_divisor = TEXTO_DIVISOR.lower()

        for linea in lineas:
            if patron_divisor in linea.lower():
                if seccion_actual:
                    secciones.append(seccion_actual)
                seccion_actual = [linea]
                seccion_actual.extend([None] * (LINEAS_CENTRADAS + 1))
            else:
                seccion_actual.append(linea)
        if seccion_actual:
            secciones.append(seccion_actual)

        if not secciones:
            print(f"¡Archivo {nombre_base} sin contenido válido!")
            return

        # Generar PDFs
        for i, seccion in enumerate(secciones):
            try:
                pdf = FPDF(orientation='P', unit='mm', format='A4')
                pdf.add_page()
                pdf.set_font("Courier", size=FUENTE_SIZE)
                pdf.set_margins(left=MARGEN_IZQ, top=20, right=MARGEN_DER)
                pdf.set_auto_page_break(True, margin=15)

                # Insertar logo
                if os.path.exists(LOGO_PATH):
                    try:
                        with Image.open(LOGO_PATH) as img:
                            img_w, img_h = img.size
                            aspect = img_h / img_w
                            logo_width = min(50, ANCHO_UTIL)
                            logo_height = logo_width * aspect
                            x = (210 - logo_width) / 2
                            y_position = pdf.get_y()
                            pdf.image(LOGO_PATH, x=x, y=y_position, w=logo_width, h=logo_height)
                            y_position += logo_height + 4
                            pdf.set_y(y_position)
                    except Exception as e:
                        print(f"Error logo: {str(e)}")

                lineas_centradas_restantes = LINEAS_CENTRADAS
                y_position = pdf.get_y()

                # Obtener datos
                numero_plaza = obtener_numero_plaza(seccion)
                sufijo = obtener_sufijo_impuesto(seccion)

                # Procesar contenido
                for elemento in seccion:
                    if elemento is None:
                        continue
                    if lineas_centradas_restantes > 0:
                        pdf.set_xy(MARGEN_IZQ, y_position)
                        pdf.cell(
                            ANCHO_UTIL,
                            ALTURA_LINEA,
                            elemento,
                            new_x=XPos.LMARGIN,
                            new_y=YPos.NEXT,
                            align='C'
                        )
                        lineas_centradas_restantes -= 1
                        y_position += ALTURA_LINEA
                    else:
                        pdf.set_xy(MARGEN_IZQ, y_position)
                        pdf.multi_cell(
                            w=ANCHO_UTIL,
                            h=ALTURA_LINEA,
                            txt=elemento,
                            new_x=XPos.LMARGIN,
                            new_y=YPos.NEXT
                        )
                        y_position = pdf.get_y()

                # Generar nombre del archivo
                nombre_base_pdf = numero_plaza if numero_plaza else f"{nombre_base}_{i+1}"
                #sufijo = sufijo if sufijo else "0000"  # Valor por defecto
                #nombre_pdf = f"{sufijo}_{nombre_base_pdf}.pdf"
                nombre_pdf = f"{sufijo_global}_{nombre_base_pdf}.pdf"  # Usar sufijo_global
                
                # Manejar duplicados
                ruta_pdf = os.path.join(CARPETA_PDFS, nombre_pdf)
                contador = 1
                while os.path.exists(ruta_pdf):
                    nuevo_nombre = f"{sufijo}_{nombre_base_pdf}_{contador}.pdf"
                    ruta_pdf = os.path.join(CARPETA_PDFS, nuevo_nombre)
                    contador += 1

                pdf.output(ruta_pdf)
                print(f"PDF generado: {os.path.basename(ruta_pdf)}")

            except Exception as e:
                print(f"Error en sección {i+1} de {nombre_base}: {str(e)}")
                continue

    except Exception as e:
        print(f"Error profesional en {nombre_base}: {str(e)}")

def convertir_prt_a_pdf():
    if not os.path.exists(CARPETA_PRTS):
        print(f"No existe la carpeta: {CARPETA_PRTS}")
        return
    
    os.makedirs(CARPETA_PDFS, exist_ok=True)
    
    for archivo in os.listdir(CARPETA_PRTS):
        if not archivo.lower().endswith('.prt'):
            continue
            
        ruta_completa = os.path.join(CARPETA_PRTS, archivo)
        nombre_base = os.path.splitext(archivo)[0]
        procesar_archivo(ruta_completa, nombre_base)

if __name__ == "__main__":
    convertir_prt_a_pdf()
    print("Proceso finalizado. Verifique la carpeta PDFs.")