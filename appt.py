import pandas as pd
import numpy as np
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import time

# ==============================================================================
# BLOCO 1: FUNÇÕES DE REGRAS DE NEGÓCIO (SEQUENCIAMENTO)
# ==============================================================================

def classificar_prioridade(prioridade_val):
    """ Regra 2: Classificação baseada no prefixo e no 5º dígito """
    p_str = str(prioridade_val).strip()
    
    if p_str.startswith('9'):
        return 'BUFF/RES'
    
    if len(p_str) >= 5:
        ref = p_str[4] 
        mapa_descricoes = {
            '1': 'URGENTE',
            '2': 'EXPORTAÇÃO',
            '3': 'VENDA DIRETA',
            '4': 'TRANSFERÊNCIA CLIENTE NA PONTA',
            '5': 'CLIENTE NA PONTA',
            '6': 'TRANSFERÊNCIA',
            '7': 'OUTROS'
        }
        return mapa_descricoes.get(ref, 'OUTROS')
    
    return 'OUTROS'

def mesma_semana(data1, data2):
    """ Verifica se duas datas estão na mesma semana e ano """
    if pd.isna(data1) or pd.isna(data2): return False
    return data1.isocalendar()[:2] == data2.isocalendar()[:2]

def mesmo_mes(data1, data2):
    """ Verifica se duas datas estão no mesmo mês e ano """
    if pd.isna(data1) or pd.isna(data2): return False
    return (data1.year == data2.year) and (data1.month == data2.month)

def processar_sequenciamento(df_base):
    """ Motor central que cruza prioridades e datas """
    # Certifique-se de que os nomes abaixo correspondem exatamente às colunas do seu arquivo/dados extraídos
    col_data = 'Data Offline'
    col_prioridade = 'Prioridade'
    col_fila = 'Fila_ID'
    col_codigo = 'Codigo_Produto'

    # Converte para datetime para garantir as comparações
    df_base[col_data] = pd.to_datetime(df_base[col_data], errors='coerce')
    
    # Aplica Regra 2
    df_base['Descrição da Prioridade'] = df_base[col_prioridade].apply(classificar_prioridade)
    
    relatorio_ok = []
    
    # Regra 4: Sets de controle para "entra e sai" sem repetir filas
    filas_movidas = set()       
    filas_substituidas = set()  

    for codigo, grupo in df_base.groupby(col_codigo):
        
        # Mapeamento de slots disponíveis no grupo
        slots_disponiveis = grupo[[col_data, col_fila]].sort_values(col_data).to_dict('records')
        
        # Regra 1: Reordenamento mantendo a sequência original em caso de empate
        filas_priorizadas = grupo.sort_values(by=[col_prioridade, col_data], ascending=[True, True])
        
        for idx, row in enumerate(filas_priorizadas.itertuples()):
            fila_atual = getattr(row, col_fila)
            data_original = getattr(row, col_data.replace(' ', '_')) # Pandas substitui espaço por _ no itertuples
            prioridade_desc = getattr(row, 'Descrição_da_Prioridade')
            prioridade_str = str(getattr(row, col_prioridade)).strip()
            
            slot_ideal = slots_disponiveis[idx]
            nova_data = slot_ideal[col_data]
            fila_ocupante = slot_ideal[col_fila]
            
            if fila_atual == fila_ocupante:
                continue
                
            # Regra 3: Filtros de tolerância
            ignorar_troca = False
            if prioridade_desc != 'BUFF/RES' and len(prioridade_str) >= 5:
                quinto_digito = prioridade_str[4]
                
                if quinto_digito in ['1', '2', '4', '6']:
                    if mesma_semana(data_original, nova_data):
                        ignorar_troca = True
                else:
                    if mesmo_mes(data_original, nova_data):
                        ignorar_troca = True
                        
            if ignorar_troca:
                continue
                
            # Regra 4: Validação de duplicidade nas colunas A e E
            if (fila_atual in filas_movidas) or (fila_ocupante in filas_substituidas):
                continue
                
            filas_movidas.add(fila_atual)
            filas_substituidas.add(fila_ocupante)
            
            # Montagem da linha do relatório de ação
            relatorio_ok.append({
                'Ação (Fila Col A)': fila_atual,
                'Nova Data': nova_data,
                'Data Original': data_original,
                'Bloqueador (Fila Col E)': fila_ocupante,
                'Descrição da Prioridade': prioridade_desc,
                'Prioridade Numérica': getattr(row, col_prioridade),
                'Código Produto': getattr(row, col_codigo)
                # Adicione as demais colunas de suporte aqui (F, Q, Z, C, etc.) usando getattr(row, 'Nome_Coluna')
            })

    return pd.DataFrame(relatorio_ok)


# ==============================================================================
# BLOCO 2: CONEXÃO E EXTRAÇÃO VIA SELENIUM
# ==============================================================================

chrome_options = Options()
chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")

servico = Service(ChromeDriverManager().install())
navegador = webdriver.Chrome(service=servico, options=chrome_options)

print(f"Conectado à página: {navegador.title}")

dados_finais = []

def extrair_dados_aprovacao():
    try:
        botao_hist = WebDriverWait(navegador, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Histórico de Aprovações')] | //a[contains(., 'Histórico de Aprovações')]"))
        )
        navegador.execute_script("arguments[0].click();", botao_hist)
        time.sleep(3) 
        
        container = WebDriverWait(navegador, 10).until(
            EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'forceRelatedListCardDesktop')] | //table"))
        )
        return container.text
    except Exception as e:
        return f"Não foi possível extrair: {str(e)}"

total_itens = 71 

for i in range(1, total_itens + 1):
    try:
        print(f"Processando item {i} de {total_itens}...")
        xpath_link = f"(//th[@scope='row']//a)[{i}]"
        link_plano = WebDriverWait(navegador, 10).until(
            EC.presence_of_element_located((By.XPATH, xpath_link))
        )
        nome_plano = link_plano.text
        navegador.execute_script("arguments[0].click();", link_plano)
        
        resultado = extrair_dados_aprovacao()
        
        # IMPORTANTE: Aqui os dados extraídos precisam ser estruturados nas colunas 
        # que o motor de regras utiliza (Prioridade, Data Offline, Fila_ID, Codigo_Produto).
        # Estou salvando no formato base para não quebrar sua estrutura atual.
        dados_finais.append({
            "Plano Nome": nome_plano,
            "Dados de Aprovação": resultado
            # Mapear aqui as colunas estruturadas, ex: "Fila_ID": valor_extraido
        })
        
        print(f"Sucesso: {nome_plano}")
        navegador.back()
        WebDriverWait(navegador, 15).until(EC.presence_of_element_located((By.TAG_NAME, "table")))
        time.sleep(2) 
        
    except Exception as e:
        print(f"Erro no item {i}: {e}")
        navegador.execute_script("window.history.go(-1)") 
        time.sleep(3)


# ==============================================================================
# BLOCO 3: PROCESSAMENTO DOS DADOS E GERAÇÃO DO EXCEL
# ==============================================================================

print("Iniciando o processamento das regras de sequenciamento...")
df_base = pd.DataFrame(dados_finais)

# Executa o motor apenas se as colunas necessárias existirem no DataFrame
colunas_necessarias = ['Data Offline', 'Prioridade', 'Fila_ID', 'Codigo_Produto']
if all(coluna in df_base.columns for coluna in colunas_necessarias):
    df_relatorio_ok = processar_sequenciamento(df_base)
else:
    print("Aviso: As colunas necessárias para o sequenciamento não foram encontradas. Gerando aba OK vazia.")
    df_relatorio_ok = pd.DataFrame(columns=['Ação (Fila Col A)', 'Nova Data', 'Data Original', 'Bloqueador (Fila Col E)'])

print("Gerando arquivo Excel com múltiplas abas...")
with pd.ExcelWriter("Relatorio_Aprovacoes_Salesforce.xlsx", engine='xlsxwriter') as writer:
    # Aba 1: Dados Brutos (como estava antes)
    df_base.to_excel(writer, sheet_name="Base Original", index=False)
    
    # Aba 2: Relatório de Ação (Aba OK)
    df_relatorio_ok.to_excel(writer, sheet_name="OK", index=False)

print("--- PROCESSO CONCLUÍDO COM SUCESSO! ---")
