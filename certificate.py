import sys
import pandas as pd
from jinja2 import FileSystemLoader, Environment
import os

# Pone en mayuscula la primera letra de cada palabra
def capitalizar_palabras(cadena):
    return cadena.title()

if not os.path.exists('vista/temp'):
    os.makedirs('vista/temp')
    print('Carpeta temporal')

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
    #cargar plantilla html con Jinja2
    env = Environment(loader=FileSystemLoader('.'))
    template = env.get_template('vista/index.html')
    
    personas = df["Nombre y apellidos completos"]
    tipoDocumento = df["Tipo de documento de identidad"]
    documento = df ["Número de documento de identidad"]
    horasCertificadas = df["Horas certificadas"]

    for i in range(len(personas)):
        #llenar el template html con los datos de las personas
        html_content = template.render(
            persona = capitalizar_palabras(personas[i]),
            documento = tipoDocumento[i],
            numeroDocumento = documento[i],
            participacion = "texto de prueba",
            iglesias = "texto de prueba",
            dias ="texto de prueba",
            mes = "texto de prueba",
            ahno ="texto de prueba",
            horas = horasCertificadas[i],
            ciudad = "texto de prueba",
            departamento= "texto de prueba",
        )

        with open (f'vista/temp/temp_{personas[i]}.html', 'w', encoding='utf-8') as f:
            f.write(html_content)
            print(f'creado vista/temp/temp_{personas[i]}.html')


if __name__ == "__main__":
    if len(sys.argv)<2:
        print("recuerde: python certificate.py nombreArchvio.xlsx")
    else:
        generarPDF(leerExcel(sys.argv[1]))
        