import sys
import pandas as pd
from jinja2 import FileSystemLoader, Environment
import os
import pdfkit

# Configurar la ruta de wkhtmltopdf
path_wkhtmltopdf = r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe'  # Actualiza esta ruta según tu instalación
config = pdfkit.configuration(wkhtmltopdf=path_wkhtmltopdf)

options = {
    'enable-local-file-access': None,  # Necesario para que wkhtmltopdf pueda acceder a archivos locales
    'page-height': '300mm',
    'page-width': '226mm',
    'orientation': 'Landscape',
    'encoding': 'UTF-8',
    'margin-top': '0in',
    'margin-right': '0in',
    'margin-bottom': '0in',
    'margin-left': '0in',
    'no-outline': None,
    'footer-line': False  # Eliminar la línea de pie de página
}

# Pone en mayuscula la primera letra de cada palabra
def capitalizar_palabras(cadena):
    return cadena.title()

if not os.path.exists('vista/temp'):
    os.makedirs('vista/temp')
    print('Carpeta temporal')

def carpetaSalida():
    if not os.path.exists('salida'):
        os.makedirs('salida')

def leerExcel(nombreArch):
    try:
        print("leyendo Excel")
        df = pd.read_excel(nombreArch)
        return df
    except FileNotFoundError:
        print(f"El archivo {nombreArch} no se encuentra")
    except Exception as e:
        print("Ocurrió un error al leer el archivo", str(e))

def generarPDF(df):
    # Cargar plantilla html con Jinja2
    env = Environment(loader=FileSystemLoader('.'))
    template = env.get_template('vista/index.html')
    
    personas = df["Nombre y apellidos completos"]
    tipoDocumento = df["Tipo de documento de identidad"]
    documento = df["Número de documento de identidad"]

    for i in range(len(personas)):
        if tipoDocumento[i] == "Cédula de ciudadanía":
            tipo_documento_abreviado = "C.C."
        else:
            tipo_documento_abreviado = tipoDocumento[i]

        # Convertir documento[i] a cadena antes de pasar a agregar_puntos()
        numero_documento_str = str(documento[i])

        # Llenar el template html con los datos de las personas
        html_content = template.render(
            persona = capitalizar_palabras(personas[i]),
            documento = tipo_documento_abreviado,
            numeroDocumento = agregar_puntos(numero_documento_str),
            participacion = "Jornadas de capacitación en Provincias Eclesiásticas",
            iglesias = "Iglesias Particulares Seguras y Protectoras",
            dias ="19, 20 y 21",
            mes = "Junio",
            ahno ="2024",
            horas = "25 horas",
            ciudad = "Barranquilla",
            departamento= "Atlántico",
        )

        # Generar el html temporal
        with open(f'vista/temp/temp_{personas[i]}.html', 'w', encoding='utf-8') as f:
            f.write(html_content)
            print(f'creado vista/temp/temp_{personas[i]}.html')

        ruta_leer = f'vista/temp/temp_{personas[i]}.html'
        ruta_salida = f'salida/{personas[i]}.pdf'

        # Generar el archivo pdf del html
        pdfkit.from_file(ruta_leer, output_path=ruta_salida, configuration=config, options=options)
        print(f'pdf generado {personas[i]}')

def agregar_puntos(cadena):
    n = len(cadena)
    grupos = []
    while n > 0:
        grupos.append(cadena[max(0, n - 3):n])
        n -= 3

    grupos.reverse()
    return '.'.join(grupos)

if __name__ == "__main__":
    if len(sys.argv)<2:
        print("recuerde: python certificate.py nombreArchvio.xlsx")
    else:
        carpetaSalida()
        generarPDF(leerExcel(sys.argv[1]))
        