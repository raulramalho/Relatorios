import os
import pandas as pd

# 1) Caminho base onde estão as pastas de códigos de empresa
base_path = r"Z:\APURAÇAO"
# 2) Nome da pasta de apuração que queremos checar
mes_apuracao = "04-2025"

results = []

# 3) Percorre cada pasta de empresa
for empresa in os.scandir(base_path):
    if not empresa.is_dir():
        continue

    codigo = empresa.name
    status = "NÃO"

    # 4) Para cada subpasta dentro da pasta da empresa
    for sub in os.scandir(empresa.path):
        if not sub.is_dir():
            continue

        # 5) Verifica se existe "04-2025" dentro dessa subpasta
        caminho_apur = os.path.join(sub.path, mes_apuracao)
        if os.path.isdir(caminho_apur):
            status = "OK"
            break

    results.append({
        "Codigo da Empresa": codigo,
        "Apuração": status
    })

# 6) Gera o Excel com os resultados
df = pd.DataFrame(results)
output_file = r"apuracao_04-2025.xlsx"
df.to_excel(output_file, index=False)

print(f"✔ Planilha gerada: {output_file}")
