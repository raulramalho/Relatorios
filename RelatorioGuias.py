import os
import pandas as pd
from datetime import date, timedelta

def mes_passado_ano_mes():
    """
    Retorna uma string no formato MMYYYY representando
    o mês e ano do mês anterior à data de hoje.
    """
    hoje = date.today()
    # Encontrar o primeiro dia do mês atual
    primeiro_deste_mes = hoje.replace(day=1)
    # Subtrair 1 dia para ir ao último dia do mês anterior
    ultimo_dia_mes_passado = primeiro_deste_mes - timedelta(days=1)
    # Formatar como MMYYYY
    return ultimo_dia_mes_passado.strftime("%m%Y")

def process_gisson(base_dir, periodo):
    """
    Para cada pasta de empresa em base_dir (GISSONLINE), entra na subpasta de período
    especificada, conta os PDFs em PRESTADO (prefixo guia_prestada_) e TOMADO (prefixo guia_tomada_),
    e retorna um DataFrame com as colunas: codigo, prestado, tomado.
    """
    registros = []
    for entrada in os.listdir(base_dir):
        caminho_empresa = os.path.join(base_dir, entrada)
        if not os.path.isdir(caminho_empresa):
            continue

        codigo = entrada.split('-', 1)[0].strip()
        prestado_count = 0
        tomado_count   = 0
        caminho_periodo = os.path.join(caminho_empresa, periodo)
        """verifica se existe uma pasta com o nome do periodo mes e ano"""
        if os.path.isdir(caminho_periodo):
            # Conta em PRESTADO
            prestado_dir = os.path.join(caminho_periodo, 'PRESTADO')
            if os.path.isdir(prestado_dir):
                for fn in os.listdir(prestado_dir):
                    nome = fn.lower()
                    if nome.startswith('guia_prestada_') and nome.endswith('.pdf'):
                        prestado_count += 1
            # Conta em TOMADO
            tomado_dir = os.path.join(caminho_periodo, 'TOMADO')
            if os.path.isdir(tomado_dir):
                for fn in os.listdir(tomado_dir):
                    nome = fn.lower()
                    if nome.startswith('guia_tomada_') and nome.endswith('.pdf'):
                        tomado_count += 1
        else:
            print(f"[Aviso GISSONLINE] Empresa {entrada}: pasta '{periodo}' não encontrada.")

        registros.append({
            'codigo': codigo,
            'prestado': prestado_count,
            'tomado': tomado_count
        })

    return pd.DataFrame(registros)


def process_prefeitura(base_dir, periodo):
    """
    Para cada pasta de empresa em base_dir (PREFEITURA SÃO PAULO), entra na subpasta de período
    especificada e verifica se existe a pasta 'GUIA DE PAGAMENTO'.
    Retorna um DataFrame com as colunas: codigo, guia_pagamento_existe (True/False).
    """
    registros = []
    for entrada in os.listdir(base_dir):
        caminho_empresa = os.path.join(base_dir, entrada)
        if not os.path.isdir(caminho_empresa):
            continue

        codigo = entrada.split('-', 1)[0].strip()
        caminho_periodo = os.path.join(caminho_empresa, periodo)
        guia_path = os.path.join(caminho_periodo, 'GUIA DE PAGAMENTO')

        existe = os.path.isdir(guia_path)
        if not os.path.isdir(caminho_periodo):
            print(f"[Aviso PREFEITURA] Empresa {entrada}: pasta '{periodo}' não encontrada.")
            existe = False

        registros.append({
            'codigo': codigo,
            'guia_pagamento_existe': existe
        })

    return pd.DataFrame(registros)


if __name__ == "__main__":
    # caminhos conforme sua máquina
    gis_dir  = r"C:\Robo CONTHABIL\GISSONLINE"
    pref_dir = r"C:\Users\NTBK_03\Desktop\Meu Drive\PREFEITURA SÃO PAULO"
    periodo= mes_passado_ano_mes()

    # processamento
    df_gis  = process_gisson(gis_dir, periodo)
    df_pref = process_prefeitura(pref_dir, periodo)

    # salva relatórios em arquivos separados
    gis_output  = os.path.join(r"Z:\Relatorio Guias",  'relatorio_'+periodo+'_giss.xlsx')
    pref_output = os.path.join(r"Z:\Relatorio Guias", 'relatorio_'+periodo+'_sao_paulo.xlsx')


    df_gis.to_excel(gis_output, index=False)
    df_pref.to_excel(pref_output, index=False)

    print(f"Relatório GISSONLINE salvo em: {gis_output}")
    print(f"Relatório PREFEITURA SÃO PAULO salvo em: {pref_output}")
