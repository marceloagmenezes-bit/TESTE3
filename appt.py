import streamlit as st
import pandas as pd
import io
import re

# ==========================================
# 1. FUNÇÕES DE TRATAMENTO E REGRAS DE NEGÓCIO
# ==========================================

def get_prioridade_str(val):
    """ Garante que a prioridade seja tratada puramente como número para extração """
    if pd.isna(val): return ""
    try:
        if isinstance(val, float):
            val = int(val)
        s = str(val).strip()
        return re.sub(r'\D', '', s)
    except:
        return ""

def classificar_prioridade(prioridade_val):
    """ Regra 2: Classificação baseada no prefixo e no 5º dígito """
    p_str = get_prioridade_str(prioridade_val)
    
    if not p_str: return 'OUTROS'
    if p_str.startswith('9'): return 'BUFF/RES'
    
    if len(p_str) >= 5:
        ref = p_str[4]  # Pega o 5º caractere (índice 4)
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
    col_str = col_str.upper()
    idx = 0
    for char in col_str:
        idx = idx * 26 + (ord(char) - ord('A')) + 1
    idx -= 1 
    
    if idx < len(df.columns):
        return df.columns[idx]
    return None

# ==========================================
# 2. MOTOR DE SEQUENCIAMENTO (BASEADO EM GRAFOS/CICLOS)
# ==========================================

def processar_sequenciamento(df_base, col_data, col_prioridade, col_fila, col_codigo, col_dealer):
    df_processo = df_base.copy()
    
    df_processo[col_data] = pd.to_datetime(df_processo[col_data], errors='coerce')
    df_processo['Descrição da Prioridade'] = df_processo[col_prioridade].apply(classificar_prioridade)
    
    # Coluna oculta e estritamente numérica para garantir a ordenação matemática da Prioridade
    df_processo['Prioridade_Num'] = pd.to_numeric(
        df_processo[col_prioridade].apply(get_prioridade_str), errors='coerce'
    ).fillna(999999999) # Se não tiver prioridade, joga pro fim da fila
    
    # Identifica dinamicamente os nomes das colunas baseadas nas letras vitais de suporte
    letras_suporte = ['F', 'Q', 'Z', 'C', 'AX', 'CQ', 'I']
    colunas_suporte = []
    for letra in letras_suporte:
        nome_col = excel_col_to_name(df_processo, letra)
        if nome_col:
            colunas_suporte.append((letra, nome_col))
    
    relatorio_ok = []

    # Agrupa pelo Código de Produto
    for codigo, grupo in df_processo.groupby(col_codigo):
        # Mapeia as datas (slots) disponíveis
        slots_disponiveis = grupo[[col_data, col_fila]].sort_values(col_data).to_dict('records')
        # Reordena o cenário ideal matematicamente usando a nova prioridade limpa
        filas_priorizadas = grupo.sort_values(by=['Prioridade_Num', col_data], ascending=[True, True]).to_dict('records')
        
        # 2.1 Construir a matriz de troca (Quem vai para o lugar de quem)
        f_to_move = {}
        for idx, row in enumerate(filas_priorizadas):
            f_atual = row[col_fila]
            f_to_move[f_atual] = {
                'ocupante': slots_disponiveis[idx][col_fila],
                'nova_data': slots_disponiveis[idx][col_data],
                'data_orig': row[col_data],
                'row': row
            }
            
        # 2.2 Detectar Correntes de Troca (Ciclos Matemáticos de "Entra e Sai")
        visited = set()
        cycles = []
        for f in f_to_move:
            if f not in visited:
                cycle = []
                curr = f
                while curr not in visited:
                    visited.add(curr)
                    cycle.append(curr)
                    curr = f_to_move[curr]['ocupante']
                cycles.append(cycle)
                
        # 2.3 Avaliar e Filtrar as Correntes
        for cycle in cycles:
            if len(cycle) == 1: 
                continue # Fila não mudou de lugar
                
            manter_ciclo = False
            
            # Se QUALQUER etapa do ciclo quebrar as regras de ignore, todo o ciclo deve acontecer
            # para não deixar fila sem destino no relatório.
            for f in cycle:
                move = f_to_move[f]
                ignorar = False
                
                p_desc = move['row']['Descrição da Prioridade']
                p_str = get_prioridade_str(move['row'][col_prioridade])
                
                # Regra 3: Tolerâncias de Data
                if p_desc != 'BUFF/RES' and len(p_str) >= 5:
                    quinto_digito = p_str[4]
                    if quinto_digito in ['1', '2', '4', '6']:
                        if mesma_semana(move['data_orig'], move['nova_data']): ignorar = True
                    else:
                        if mesmo_mes(move['data_orig'], move['nova_data']): ignorar = True
                
                # Regra 5: Mesmo Dealer / Cliente
                dealer_atual = move['row'][col_dealer]
                dealer_ocupante = f_to_move[move['ocupante']]['row'][col_dealer]
                if pd.notna(dealer_atual) and str(dealer_atual).strip() != "":
                    if dealer_atual == dealer_ocupante:
                        ignorar = True # Ignora se vai trocar com o mesmo cliente
                        
                # Se acharmos um movimento que NÃO pode ser ignorado, o ciclo inteiro roda
                if not ignorar:
                    manter_ciclo = True
                    break 
                    
            # 2.4 Preencher o Relatório OK
            if manter_ciclo:
                for f in cycle:
                    move = f_to_move[f]
                    row = move['row']
                    
                    linha_acao = {
                        'Identificação da fila (Col A)': f,
                        'Data correta sugerida': move['nova_data'],
                        'Data Original': move['data_orig'],
                        'Fila bloqueadora (Col E)': move['ocupante'],
                        'Descrição da Prioridade': row['Descrição da Prioridade'],
                        'Prioridade Real': row[col_prioridade],
                        'Código do Produto': row[col_codigo],
                        'Dealer Sugerido': row[col_dealer],
                        'Dealer Bloqueador': f_to_move[move['ocupante']]['row'][col_dealer]
                    }
                    
                    # Anexar as antigas letras de suporte (F, Q, Z, C...)
                    for letra, nome_col in colunas_suporte:
                        linha_acao[f"Col {letra} - {nome_col}"] = row[nome_col]
                            
                    relatorio_ok.append(linha_acao)

    df_final = pd.DataFrame(relatorio_ok)
    
    # ORDENAÇÃO EXIGIDA NO RELATÓRIO: 1) Código do produto, 2) Data correta sugerida
    if not df_final.empty:
        df_final = df_final.sort_values(by=['Código do Produto', 'Data correta sugerida'])
        
    return df_final

# ==========================================
# 3. INTERFACE STREAMLIT
# ==========================================

st.set_page_config(page_title="Sequenciamento de Produção", layout="wide")
st.title("⚙️ Sequenciamento de Produção")
st.markdown("Faça o upload da planilha base para gerar o relatório de ação.")

arquivo_upload = st.file_uploader("Selecione o arquivo Excel", type=["xlsx", "xls", "csv"])

if arquivo_upload is not None:
    try:
        # Tenta ler como Excel, faz fallback para CSV caso você esteja subindo em .csv para testar
        if arquivo_upload.name.endswith('csv'):
            df_original = pd.read_csv(arquivo_upload)
        else:
            df_original = pd.read_excel(arquivo_upload)
            
        colunas_disponiveis = df_original.columns.tolist()
        
        st.success("Arquivo carregado com sucesso!")
        st.markdown("### 🔍 Mapeamento de Colunas Principais")
        
        # Mapeamento expandido para caber a nova coluna do Dealer
        c1, c2, c3 = st.columns(3)
        c4, c5 = st.columns(2)
        
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
        with c5:
            # Tenta preencher automaticamente alguma coluna que tenha DEALER ou CLI no nome
            possiveis_dealer = [c for c in colunas_disponiveis if 'DEALER' in str(c).upper() or 'CLI' in str(c).upper()]
            idx_dealer = colunas_disponiveis.index(possiveis_dealer[0]) if possiveis_dealer else 0
            col_dealer = st.selectbox("Coluna de Dealer / Cliente", colunas_disponiveis, index=idx_dealer)
            
        st.divider()

        if st.button("🚀 Processar Sequenciamento", type="primary"):
            with st.spinner("Calculando logísticas e balanceando matrizes..."):
                df_ok = processar_sequenciamento(df_original, col_data, col_prioridade, col_fila, col_codigo, col_dealer)
                
                st.subheader("Pré-visualização da Aba OK")
                if df_ok.empty:
                    st.warning("Nenhuma alteração necessária encontrada.")
                else:
                    # Formata as datas visualmente
                    if 'Data correta sugerida' in df_ok.columns:
                        df_ok['Data correta sugerida'] = df_ok['Data correta sugerida'].dt.strftime('%d/%m/%Y')
                    if 'Data Original' in df_ok.columns:
                        df_ok['Data Original'] = df_ok['Data Original'].dt.strftime('%d/%m/%Y')
                        
                    st.dataframe(df_ok)
                
                # Geração do arquivo Excel com formatação
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    df_original.to_excel(writer, sheet_name='Base Original', index=False)
                    df_ok.to_excel(writer, sheet_name='OK', index=False)
                    
                    # Aplica as cores no cabeçalho
                    workbook  = writer.book
                    worksheet = writer.sheets['OK']
                    
                    formato_cabecalho = workbook.add_format({
                        'bg_color': '#1F4E78', 
                        'font_color': 'white',
                        'bold': True,
                        'border': 1
                    })
                    
                    for col_num, value in enumerate(df_ok.columns.values):
                        worksheet.write(0, col_num, value, formato_cabecalho)
                        worksheet.set_column(col_num, col_num, 18)
                
                st.download_button(
                    label="📥 Baixar Relatório Consolidado (Excel)",
                    data=buffer.getvalue(),
                    file_name="Relatorio_Sequenciamento_V4.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        
    except Exception as e:
        st.error(f"Ocorreu um erro ao processar o arquivo: {e}")
