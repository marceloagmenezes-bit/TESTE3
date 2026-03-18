import streamlit as st
import pandas as pd
import io
import re

# ==========================================
# 1. FUNÇÕES DE REGRAS DE NEGÓCIO
# ==========================================

def classificar_prioridade(prioridade_val):
    """ Regra 2: Classificação baseada no prefixo e no 5º dígito """
    # Remove casas decimais que o Pandas pode adicionar (ex: .0) e limpa tudo que não for número puro
    p_str = str(prioridade_val).split('.')[0]
    p_str = re.sub(r'\D', '', p_str)
    
    if p_str.startswith('9'):
        return 'BUFF/RES'
    
    if len(p_str) >= 5:
        ref = p_str[4]  # Pega exatamente o 5º número
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
    if pd.isna(data1) or pd.isna(data2): return False
    return data1.isocalendar()[:2] == data2.isocalendar()[:2]

def mesmo_mes(data1, data2):
    if pd.isna(data1) or pd.isna(data2): return False
    return (data1.year == data2.year) and (data1.month == data2.month)

def excel_col_to_name(df, col_str):
    """ Converte letra de coluna do Excel (ex: 'AX') para o nome do cabeçalho no Pandas """
    col_str = col_str.upper()
    idx = 0
    for char in col_str:
        idx = idx * 26 + (ord(char) - ord('A')) + 1
    idx -= 1 
    
    if idx < len(df.columns):
        return df.columns[idx]
    return None

# ==========================================
# 2. MOTOR DE SEQUENCIAMENTO
# ==========================================

def processar_sequenciamento(df_base, col_data, col_prioridade, col_fila, col_codigo):
    df_processo = df_base.copy()
    
    df_processo[col_data] = pd.to_datetime(df_processo[col_data], errors='coerce')
    df_processo['Descrição da Prioridade'] = df_processo[col_prioridade].apply(classificar_prioridade)
    
    # Identifica os nomes das colunas baseadas nas letras vitais de suporte
    letras_suporte = ['F', 'Q', 'Z', 'C', 'AX', 'CQ', 'I']
    colunas_suporte = []
    for letra in letras_suporte:
        nome_col = excel_col_to_name(df_processo, letra)
        if nome_col:
            colunas_suporte.append((letra, nome_col))
    
    relatorio_ok = []
    filas_movidas = set()       
    filas_substituidas = set()  

    for codigo, grupo in df_processo.groupby(col_codigo):
        slots_disponiveis = grupo[[col_data, col_fila]].sort_values(col_data).to_dict('records')
        filas_priorizadas = grupo.sort_values(by=[col_prioridade, col_data], ascending=[True, True]).to_dict('records')
        
        for idx, row in enumerate(filas_priorizadas):
            fila_atual = row[col_fila]
            data_original = row[col_data]
            prioridade_desc = row['Descrição da Prioridade']
            
            # Limpeza do número real para aplicar regras de tolerância de data
            prioridade_str = str(row[col_prioridade]).split('.')[0]
            prioridade_str = re.sub(r'\D', '', prioridade_str)
            
            slot_ideal = slots_disponiveis[idx]
            nova_data = slot_ideal[col_data]
            fila_ocupante = slot_ideal[col_fila]
            
            if fila_atual == fila_ocupante:
                continue
                
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
                
            if (fila_atual in filas_movidas) or (fila_ocupante in filas_substituidas):
                continue
                
            filas_movidas.add(fila_atual)
            filas_substituidas.add(fila_ocupante)
            
            # Montagem do Output EXATAMENTE como pedido na primeira versão
            linha_acao = {
                'Identificação da fila (Col A)': fila_atual,
                'Data correta sugerida': nova_data,
                'Data Original': data_original,
                'Fila bloqueadora (Col E)': fila_ocupante,
                'Descrição da Prioridade': prioridade_desc,
                'Prioridade Real': row[col_prioridade],
                'Código do Produto': row[col_codigo]
            }
            
            # Anexa APENAS as colunas de suporte requisitadas
            for letra, nome_col in colunas_suporte:
                linha_acao[f"Col {letra} - {nome_col}"] = row[nome_col]
                    
            relatorio_ok.append(linha_acao)

    return pd.DataFrame(relatorio_ok)

# ==========================================
# 3. INTERFACE STREAMLIT
# ==========================================

st.set_page_config(page_title="Sequenciamento de Produção", layout="wide")
st.title("⚙️ Sequenciamento de Produção")
st.markdown("Faça o upload da planilha base para gerar o relatório de ação com análise de prioridades.")

arquivo_upload = st.file_uploader("Selecione o arquivo Excel", type=["xlsx", "xls"])

if arquivo_upload is not None:
    try:
        df_original = pd.read_excel(arquivo_upload)
        colunas_disponiveis = df_original.columns.tolist()
        
        st.success("Arquivo carregado com sucesso!")
        st.markdown("### 🔍 Mapeamento de Colunas Principais")
        
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            idx_data = colunas_disponiveis.index('Data Offline') if 'Data Offline' in colunas_disponiveis else 0
            col_data = st.selectbox("Coluna de Data", colunas_disponiveis, index=idx_data)
        with c2:
            idx_prio = colunas_disponiveis.index('Prioridade') if 'Prioridade' in colunas_disponiveis else 0
            col_prioridade = st.selectbox("Coluna de Prioridade", colunas_disponiveis, index=idx_prio)
        with c3:
            idx_fila = colunas_disponiveis.index('Fila_ID') if 'Fila_ID' in colunas_disponiveis else 0
            col_fila = st.selectbox("Coluna de Fila (ID)", colunas_disponiveis, index=idx_fila)
        with c4:
            idx_cod = colunas_disponiveis.index('Codigo_Produto') if 'Codigo_Produto' in colunas_disponiveis else 0
            col_codigo = st.selectbox("Coluna de Código de Produto", colunas_disponiveis, index=idx_cod)
            
        st.divider()

        if st.button("🚀 Processar Sequenciamento", type="primary"):
            with st.spinner("Processando..."):
                df_ok = processar_sequenciamento(df_original, col_data, col_prioridade, col_fila, col_codigo)
                
                st.subheader("Pré-visualização da Aba OK")
                if df_ok.empty:
                    st.warning("Nenhuma alteração necessária encontrada.")
                else:
                    if 'Data correta sugerida' in df_ok.columns:
                        df_ok['Data correta sugerida'] = df_ok['Data correta sugerida'].dt.strftime('%d/%m/%Y')
                    if 'Data Original' in df_ok.columns:
                        df_ok['Data Original'] = df_ok['Data Original'].dt.strftime('%d/%m/%Y')
                        
                    st.dataframe(df_ok)
                
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    df_original.to_excel(writer, sheet_name='Base Original', index=False)
                    df_ok.to_excel(writer, sheet_name='OK', index=False)
                
                st.download_button(
                    label="📥 Baixar Relatório Consolidado (Excel)",
                    data=buffer.getvalue(),
                    file_name="Relatorio_Sequenciamento_Atualizado.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        
    except Exception as e:
        st.error(f"Ocorreu um erro ao processar o arquivo: {e}")

