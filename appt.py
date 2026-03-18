import streamlit as st
import pandas as pd
import io

# ==========================================
# 1. FUNÇÕES DE REGRAS DE NEGÓCIO
# ==========================================

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
    """ Regra 3: Verifica se estão na mesma semana e ano """
    if pd.isna(data1) or pd.isna(data2): return False
    return data1.isocalendar()[:2] == data2.isocalendar()[:2]

def mesmo_mes(data1, data2):
    """ Regra 3: Verifica se estão no mesmo mês e ano """
    if pd.isna(data1) or pd.isna(data2): return False
    return (data1.year == data2.year) and (data1.month == data2.month)

# ==========================================
# 2. MOTOR DE SEQUENCIAMENTO
# ==========================================

def processar_sequenciamento(df_base):
    # ATENÇÃO: Ajuste os nomes abaixo para os cabeçalhos reais da sua planilha
    col_data = 'Data Offline'
    col_prioridade = 'Prioridade'
    col_fila = 'Fila_ID'
    col_codigo = 'Codigo_Produto'

    # Converte para datetime e cria a descrição da prioridade
    df_base[col_data] = pd.to_datetime(df_base[col_data], errors='coerce')
    df_base['Descrição da Prioridade'] = df_base[col_prioridade].apply(classificar_prioridade)
    
    relatorio_ok = []
    
    # Regra 4: Controle de filas únicas nas colunas A e E
    filas_movidas = set()       
    filas_substituidas = set()  

    # Agrupa pelo Código de Produto (AG)
    for codigo, grupo in df_base.groupby(col_codigo):
        
        # Mapeamento de slots (Datas disponíveis ordenadas cronologicamente)
        slots_disponiveis = grupo[[col_data, col_fila]].sort_values(col_data).to_dict('records')
        
        # Regra 1: Reordenamento mantendo a sequência original em caso de empate na prioridade
        filas_priorizadas = grupo.sort_values(by=[col_prioridade, col_data], ascending=[True, True])
        
        for idx, row in enumerate(filas_priorizadas.itertuples()):
            fila_atual = getattr(row, col_fila)
            data_original = getattr(row, col_data.replace(' ', '_')) 
            prioridade_desc = getattr(row, 'Descrição_da_Prioridade')
            prioridade_str = str(getattr(row, col_prioridade)).strip()
            
            # Compara a fila com a melhor data disponível (slot ideal)
            slot_ideal = slots_disponiveis[idx]
            nova_data = slot_ideal[col_data]
            fila_ocupante = slot_ideal[col_fila]
            
            # Se já está no lugar certo, ignora
            if fila_atual == fila_ocupante:
                continue
                
            # Regra 3: Tolerâncias de data (Semana ou Mês)
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
                
            # Regra 4: Validação de duplicidade (só entra no relatório se não repete)
            if (fila_atual in filas_movidas) or (fila_ocupante in filas_substituidas):
                continue
                
            filas_movidas.add(fila_atual)
            filas_substituidas.add(fila_ocupante)
            
            # Adiciona ao relatório
            relatorio_ok.append({
                'Ação (Fila Col A)': fila_atual,
                'Nova Data': nova_data,
                'Data Original': data_original,
                'Bloqueador (Fila Col E)': fila_ocupante,
                'Descrição da Prioridade': prioridade_desc,
                'Prioridade Numérica': getattr(row, col_prioridade),
                'Código Produto': getattr(row, col_codigo)
            })

    return pd.DataFrame(relatorio_ok)

# ==========================================
# 3. INTERFACE STREAMLIT
# ==========================================

st.set_page_config(page_title="Sequenciamento de Produção", layout="wide")
st.title("⚙️ Sequenciamento de Produção")
st.markdown("Faça o upload da planilha base para gerar o relatório de ação.")

arquivo_upload = st.file_uploader("Selecione o arquivo Excel", type=["xlsx", "xls"])

if arquivo_upload is not None:
    try:
        # Lê o arquivo
        df_original = pd.read_excel(arquivo_upload)
        st.success("Arquivo carregado com sucesso! Processando as regras...")
        
        # Roda o motor de regras
        df_ok = processar_sequenciamento(df_original)
        
        # Mostra um preview na tela
        st.subheader("Pré-visualização da Aba OK")
        st.dataframe(df_ok)
        
        # Gera o Excel em memória para download
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df_original.to_excel(writer, sheet_name='Base Original', index=False)
            df_ok.to_excel(writer, sheet_name='OK', index=False)
        
        # Botão de Download
        st.download_button(
            label="📥 Baixar Relatório Consolidado (Excel)",
            data=buffer.getvalue(),
            file_name="Relatorio_Sequenciamento.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"Ocorreu um erro ao processar o arquivo: {e}")
