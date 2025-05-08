import os
import pandas as pd

def contar_arquivos(pasta):
    """Retorna a quantidade de arquivos em uma pasta (não inclui subdiretórios)."""
    if not os.path.exists(pasta):
        return 0
    return sum(
        1 for f in os.listdir(pasta)
        if os.path.isfile(os.path.join(pasta, f))
    )

def gerar_relatorio(base_path, output_file):
    """
    Gera um relatório em planilha Excel com duas colunas:
    'Codigo da empresa' e 'Quantidade de arquivos importados'.
    """
    resultados = []
    # Itera por cada pasta de código de empresa dentro de BASE_PATH
    for empresa in os.listdir(base_path):
        empresa_path = os.path.join(base_path, empresa)
        if not os.path.isdir(empresa_path):
            continue

        # Caminhos para as subpastas de interesse
        pasta_entrada_saida = os.path.join(empresa_path, 'entrada_saida', 'Adicionadas automaticamente')
        pasta_servico = os.path.join(empresa_path, 'servico', 'Adicionadas automaticamente')

        # Contagem de arquivos em cada pasta
        qtd_entrada_saida = contar_arquivos(pasta_entrada_saida)
        qtd_servico = contar_arquivos(pasta_servico)
        total = qtd_entrada_saida + qtd_servico

        resultados.append({
            'Codigo da empresa': empresa,
            'Quantidade de arquivos importados': total
        })

    # Cria DataFrame e salva em arquivo Excel
    df = pd.DataFrame(resultados)
    df.to_excel(output_file, index=False)
    print(f'Relatório salvo em: {output_file}')

if __name__ == '__main__':
    BASE_PATH = r'Z:\MODELO ALTERDATA'
    OUTPUT_FILE = 'relatorio_arquivos_importados.xlsx'
    gerar_relatorio(BASE_PATH, OUTPUT_FILE)
