import os

path_folder = "data/output/resultado_comparacao.xlsx"

if os.path.exists(path_folder):
    os.remove(path_folder)
    print(f"Arquivo apagado con sucesso: {path_folder}")
else:
    print("O arquivo não foi encontrado.")