import pandas as pd

with open('entorno.txt', 'r', encoding = "utf-8") as file:
    content = file.readlines()

configuraciones = {}

for i, line in enumerate(content):
    if line.strip() == "":
        break
    key, value = line.strip().split(":", 1)
    configuraciones[key.strip()] = value.strip()

print("Configuraciones:")
for key, value in configuraciones.items():
    print(f"{key}: {value}")

print( configuraciones["iglesias"])