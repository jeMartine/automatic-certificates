import sys
import pandas as pd
from jinja2 import FileSystemLoader, Environment
import os
import pdfkit
import time
import shutil


# Configurar la ruta de wkhtmltopdf
path_wkhtmltopdf = r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe'  # Actualiza esta ruta según tu instalación
config = pdfkit.configuration(wkhtmltopdf=path_wkhtmltopdf)
configuraciones = {}

options = {
    'enable-local-file-access': None,  # Necesario para que wkhtmltopdf pueda acceder a archivos locales
    'page-size': 'A4',
    'orientation': 'Landscape',
    'encoding': 'UTF-8',
    'margin-top': '0in',
    'margin-right': '0in',
    'margin-bottom': '0in',
    'margin-left': '0in',
    'no-outline': None,
}

# Pone en mayuscula la primera letra de cada palabra
def capitalizar_palabras(cadena):
    if isinstance(cadena, str):
        return cadena.title()
    else:
        return ""

def borrar_carpeta_temp():
    ruta = 'vista/temp'
    if os.path.exists(ruta):
        shutil.rmtree(ruta)
        print(f'Carpeta "{ruta}" eliminada.')
    else:
        print(f'La carpeta "{ruta}" no existe.')

if not os.path.exists('vista/temp'):
    os.makedirs('vista/temp')
    print('Carpeta temporal')

if not os.path.exists('formulario'):
    os.makedirs('formulario')
    print('Carpeta formulario')

def carpetaSalida():
    if not os.path.exists('salida'):
        os.makedirs('salida')

def leerConfiguraciones(nombreConfiguraciones):
    global configuraciones
    with open(nombreConfiguraciones, 'r', encoding = "utf-8") as file:
        content = file.readlines()


    for line in content:
        if line.strip() == "":
            break
        key, value = line.strip().split(":", 1)
        configuraciones[key.strip()] = value.strip()

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
    numero_documento = 0
    ubicacion = ""
    personas = df[configuraciones["personas"]]
    tipoDocumento = df[configuraciones["tipoDocumento"]]
    documento = df[configuraciones["documento"]]
    horas = df[configuraciones["horas"]]
    jurisdiccion = df[configuraciones["jurisdiccion"]]

    for i in range(len(personas)):
        # if tipoDocumento[i] == "Cédula de ciudadanía":
        #     tipo_documento_abreviado = "C.C."
        # else: 
        #     if tipoDocumento[i] == "Pasaporte":
        #         tipo_documento_abreviado = "PA"
        #     else: 
        #         tipo_documento_abreviado = tipoDocumento[i]

        #Tipo de documento abreviado
        tipo_doc = ""
        if str(tipoDocumento[i])  != "nan":
            tipo_doc = "Identificado(a) con " + str(tipoDocumento[i])


        #Numero de documento
        if str(documento[i]) == "nan":
            numero_documento = " "
        else:
            numero_documento = agregar_puntos(str(int(documento[i])))

        #Ubicacion
        if configuraciones["departamento"] == "no":
            ubicacion = configuraciones["ciudad"]
        else:
            ubicacion = f"{configuraciones["ciudad"]}, {configuraciones["departamento"]}"


        # Llenar el template html con los datos de las personas
        html_content = template.render(
            persona = capitalizar_palabras(personas[i]),
            documento = tipo_doc,
            numeroDocumento = numero_documento,
            participacion = configuraciones["participacion"],
            iglesias = configuraciones["iglesias"],
            dias = configuraciones["dias"],
            mes = configuraciones["mes"],
            ahno = configuraciones["ahno"],
            horas = int(horas[i]),
            ubicacion_completa = ubicacion,
            jurisdiccion = jurisdiccion[i],
        )

        # Generar el html temporal
        with open(f'vista/temp/temp_{personas[i]}.html', 'w', encoding='utf-8') as f:
            f.write(html_content)
            print(f'creado vista/temp/temp_{personas[i]}.html')

        carpetaLugar(jurisdiccion[i])
        ruta_leer = f'vista/temp/temp_{personas[i]}.html'
        ruta_salida = f'salida/{jurisdiccion[i]}/{personas[i]}.pdf'
        
        # Generar el archivo pdf del html
        pdfkit.from_file(ruta_leer, output_path=ruta_salida, configuration=config, options=options)
        print(f'pdf generado {personas[i]}')

def carpetaLugar(lugar):
    if not os.path.exists(f'salida/{lugar}'):
        os.makedirs(f'salida/{lugar}')

def agregar_puntos(cadena):
    n = len(cadena)
    grupos = []
    while n > 0:
        grupos.append(cadena[max(0, n - 3):n])
        n -= 3

    grupos.reverse()
    return '.'.join(grupos)

if __name__ == "__main__":
    delete = True
    if len(sys.argv)<2:
        print("recuerde: python certificate.py nombreArchvio.xlsx [--test]")
    else:
        if len(sys.argv) > 3:
            if sys.argv[2] == "--test":
                delete = False
        rutaExcel = f'formulario/'+ sys.argv[1]
        print(rutaExcel)
        leerConfiguraciones('entorno.txt')
        carpetaSalida()
        generarPDF(leerExcel(rutaExcel))

        if delete:
            borrar_carpeta_temp()
            print("Carpeta temporal borrada")
        