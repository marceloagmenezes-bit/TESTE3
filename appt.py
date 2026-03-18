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
# 2. MOTOR DE SEQUENCIAMENTO E CONFRONTO
# ==========================================

def processar_sequenciamento(df_base, col_data, col_prioridade, col_fila, col_codigo):
    # Cria uma cópia para preservar os dados originais
    df_processo = df_base.copy()
    
    # Padroniza as datas para o cálculo de tolerância
    df_processo[col_data] = pd.to_datetime(df_processo[col_data], errors='coerce')
    
    # Aplica a Regra 2 (Descrição da Prioridade)
    df_processo['Descrição da Prioridade'] = df_processo[col_prioridade].apply(classificar_prioridade)
    
    relatorio_ok = []
    
    # Regra 4: Controle de filas únicas para garantir o "Entra e Sai" sem repetição
    filas_movidas = set()       
    filas_substituidas = set()  

    # Agrupamento (AG) - Passo 1 da Arquitetura Original
    for codigo, grupo in df_processo.groupby(col_codigo):
        
        # Mapeamento de Datas - Passo 2 da Arquitetura Original (Slots cronológicos)
        slots_disponiveis = grupo[[col_data, col_fila]].sort_values(col_data).to_dict('records')
        
        # Reordenamento por Prioridade - Passo 3 da Arquitetura Original
        # Regra 1 incluída: 'col_data' atua como desempate se a prioridade for igual
        filas_priorizadas = grupo.sort_values(by=[col_prioridade, col_data], ascending=[True, True]).to_dict('records')
        
        # Confronto Logístico - Passo 4 da Arquitetura Original
        for idx, row in enumerate(filas_priorizadas):
            fila_atual = row[col_fila]
            data_original = row[col_data]
            prioridade_desc = row['Descrição da Prioridade']
            prioridade_str = str(row[col_prioridade]).strip()
            
            # Puxa o cenário ideal (Casamento de Data x Prioridade)
            slot_ideal = slots_disponiveis[idx]
            nova_data = slot_ideal[col_data]
            fila_ocupante = slot_ideal[col_fila]
            
            # Se não houver divergência (fila já está no lugar certo), ignora
            if fila_atual == fila_ocupante:
                continue
                
            # --- Aplicação da Regra 3 (Tolerâncias) ---
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
                
            # --- Aplicação da Regra 4 (Unicidade Coluna A e E) ---
            if (fila_atual in filas_movidas) or (fila_ocupante in filas_substituidas):
                continue
                
            # Registra as filas para não repeti-las
            filas_movidas.add(fila_atual)
            filas_substituidas.add(fila_ocupante)
            
            # --- Montagem do Output (Relatório de Ação) ---
            linha_acao = {
                'Identificação da Fila (Col A)': fila_atual,
                'Data Correta Sugerida': nova_data,
                'Data Original': data_original,
                'Fila Bloqueadora (Col E)': fila_ocupante,
                'Descrição da Prioridade': prioridade_desc,
                'Prioridade Real': row[col_prioridade],
                'Código do Produto': row[col_codigo]
            }
            
            # Recupera as "Colunas de Suporte" (F, Q, Z, C, AX, CQ, I, etc.) da linha original
            # O loop abaixo adiciona todas as colunas que vieram no arquivo e não estão no escopo principal acima
            colunas_principais = [col_data, col_prioridade, col_fila, col_codigo, 'Descrição da Prioridade']
            for coluna_extra in df_processo.columns:
                if coluna_extra not in colunas_principais:
                    linha_acao[coluna_extra] = row[coluna_extra]
                    
            relatorio_ok.append(linha_acao)

    # Converte tudo para DataFrame final
    return pd.DataFrame(relatorio_ok)

# ==========================================
# 3. INTERFACE STREAMLIT
# ==========================================

st.set_page_config(page_title="Sequenciamento de Produção", layout="wide")
st.title("⚙️ Sequenciamento de Produção")
st.markdown("Faça o upload da planilha base para gerar o relatório de ação com análise de prioridades e tolerâncias.")

arquivo_upload = st.file_uploader("Selecione o arquivo Excel", type=["xlsx", "xls"])

if arquivo_upload is not None:
    try:
        # Lê o arquivo base
        df_original = pd.read_excel(arquivo_upload)
        colunas_disponiveis = df_original.columns.tolist()
        
        st.success("Arquivo carregado com sucesso!")
        st.markdown("### 🔍 Mapeamento de Colunas")
        st.write("Confirme os nomes das colunas correspondentes no seu Excel:")
        
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
            with st.spinner("Confrontando logística e aplicando regras de negócio..."):
                
                # Executa o algoritmo completo
                df_ok = processar_sequenciamento(df_original, col_data, col_prioridade, col_fila, col_codigo)
                
                st.subheader("Pré-visualização da Aba OK")
                if df_ok.empty:
                    st.warning("Nenhuma alteração necessária encontrada após a aplicação das regras de tolerância e prioridade.")
                else:
                    # Ajusta as datas de volta para um formato legível na tela antes de exportar
                    if 'Data Correta Sugerida' in df_ok.columns:
                        df_ok['Data Correta Sugerida'] = df_ok['Data Correta Sugerida'].dt.strftime('%d/%m/%Y')
                    if 'Data Original' in df_ok.columns:
                        df_ok['Data Original'] = df_ok['Data Original'].dt.strftime('%d/%m/%Y')
                        
                    st.dataframe(df_ok)
                
                # Gera o Excel em memória
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
