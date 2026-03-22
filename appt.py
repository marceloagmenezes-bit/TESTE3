import streamlit as st
import pandas as pd
import io
import re

# ==========================================
# 1. FUNÇÕES DE TRATAMENTO E REGRAS DE NEGÓCIO
# ==========================================

def get_prioridade_str(val):
    if pd.isna(val): return ""
    try:
        if isinstance(val, float):
            val = int(val)
        s = str(val).strip()
        return re.sub(r'\D', '', s)
    except:
        return ""

def classificar_prioridade(prioridade_val):
    p_str = get_prioridade_str(prioridade_val)
    if not p_str: return 'OUTROS'
    if p_str.startswith('9'): return 'BUFF/RES'
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

def encontrar_indice_coluna(colunas_disponiveis, palavras_chave):
    """ Radar inteligente que procura palavras-chave nos nomes das colunas e retorna o índice """
    for i, col in enumerate(colunas_disponiveis):
        col_upper = str(col).upper()
        for palavra in palavras_chave:
            if palavra.upper() in col_upper:
                return i
    return 0 

# ==========================================
# 2. MOTOR DE SEQUENCIAMENTO 
# ==========================================

def processar_sequenciamento(df_base, col_data, col_prioridade, col_fila, col_codigo, col_dealer):
    df_processo = df_base.copy()
    
    df_processo[col_data] = pd.to_datetime(df_processo[col_data], errors='coerce')
    df_processo['Descrição da Prioridade'] = df_processo[col_prioridade].apply(classificar_prioridade)
    
    df_processo['Prioridade_Num'] = pd.to_numeric(
        df_processo[col_prioridade].apply(get_prioridade_str), errors='coerce'
    ).fillna(999999999) 
    
    col_C = excel_col_to_name(df_processo, 'C') # TIPO_DEMANDA
    col_I = excel_col_to_name(df_processo, 'I') # OBS_FILA
    col_Z = excel_col_to_name(df_processo, 'Z') # FILIAL
    col_AX = excel_col_to_name(df_processo, 'AX') # DS_MERCADO
    col_CQ = excel_col_to_name(df_processo, 'CQ') # MARCA
    
    relatorio_ok = []

    for codigo, grupo in df_processo.groupby(col_codigo):
        estado_atual = grupo.to_dict('records')
        estado_atual.sort(key=lambda x: x[col_data])
        
        filas_usadas = set()
        
        for i in range(len(estado_atual)):
            slot_data = estado_atual[i][col_data]
            ocupante_atual = estado_atual[i]
            
            if ocupante_atual[col_fila] in filas_usadas:
                continue
            
            candidatos = [c for c in estado_atual[i:] if c[col_fila] not in filas_usadas]
            
            if not candidatos:
                continue
                
            candidatos = sorted(candidatos, key=lambda x: (x['Prioridade_Num'], x[col_data]))
            ideal = candidatos[0]
            
            if ideal[col_fila] == ocupante_atual[col_fila]:
                continue
                
            data_orig_ideal = ideal[col_data]
            
            ignorar = False
            p_desc = ideal['Descrição da Prioridade']
            p_str = get_prioridade_str(ideal[col_prioridade])
            
            if p_desc != 'BUFF/RES' and len(p_str) >= 5:
                quinto_digito = p_str[4]
                if quinto_digito in ['1', '2', '4', '6']:
                    if mesma_semana(data_orig_ideal, slot_data): ignorar = True
                else:
                    if mesmo_mes(data_orig_ideal, slot_data): ignorar = True
            
            dealer_ideal = ideal.get(col_dealer)
            dealer_ocupante = ocupante_atual.get(col_dealer)
            if pd.notna(dealer_ideal) and str(dealer_ideal).strip() != "":
                if dealer_ideal == dealer_ocupante:
                    ignorar = True
            
            if not ignorar:
                linha_swap = {
                    'Fila': ideal[col_fila],
                    'Data Original': data_orig_ideal,
                    'Data sugerida': slot_data,
                    
                    'Código do Produto': ideal[col_codigo],
                    'Dealer Original': ideal.get(col_dealer),
                    'Observação Fila': ideal.get(col_I),
                    'Demanda Fila': ideal.get(col_C),
                    'Desc. Prioridade Fila': ideal['Descrição da Prioridade'],
                    
                    'Fila inversão': ocupante_atual[col_fila],
                    'Dealer Inversão': ocupante_atual.get(col_dealer),
                    'Observação Inversão': ocupante_atual.get(col_I),
                    'Demanda Inversão': ocupante_atual.get(col_C),
                    'Desc. Prioridade Inversão': ocupante_atual['Descrição da Prioridade'],
                    
                    'Marca (Comum)': ideal.get(col_CQ),
                    'DS Mercado (Comum)': ideal.get(col_AX),
                    'Filial (Comum)': ideal.get(col_Z)
                }
                
                relatorio_ok.append(linha_swap)
                
                filas_usadas.add(ideal[col_fila])
                filas_usadas.add(ocupante_atual[col_fila])
                
                idx_ideal = estado_atual.index(ideal)
                ideal_copia = ideal.copy()
                ocupante_copia = ocupante_atual.copy()
                ideal_copia[col_data] = slot_data
                ocupante_copia[col_data] = data_orig_ideal
                
                estado_atual[i] = ideal_copia
                estado_atual[idx_ideal] = ocupante_copia

    df_final = pd.DataFrame(relatorio_ok)
    
    if not df_final.empty:
        df_final = df_final.sort_values(by=['Código do Produto', 'Data sugerida'])
        df_final['Data Original'] = df_final['Data Original'].dt.strftime('%d/%m/%Y')
        df_final['Data sugerida'] = df_final['Data sugerida'].dt.strftime('%d/%m/%Y')
        
    return df_final

# ==========================================
# 3. INTERFACE STREAMLIT E EXPORTAÇÃO EXCEL
# ==========================================

st.set_page_config(page_title="Smart Order Management", layout="wide")
st.title("⚙️ Smart Order Management")
st.markdown("Faça o upload da base de produção para calcular o balanceamento comercial.")

arquivo_upload = st.file_uploader("Selecione o arquivo Excel", type=["xlsx", "xls", "csv"])

if arquivo_upload is not None:
    try:
        if arquivo_upload.name.endswith('csv'):
            df_original = pd.read_csv(arquivo_upload)
        else:
            df_original = pd.read_excel(arquivo_upload)
            
        colunas_disponiveis = df_original.columns.tolist()
        
        st.success("Arquivo carregado com sucesso!")
        st.markdown("### 🔍 Mapeamento de Colunas Principais")
        
        c1, c2, c3 = st.columns(3)
        c4, c5 = st.columns(2)
        
        # Uso do Radar Inteligente atualizado com as nomenclaturas oficiais
        with c1:
            idx_data = encontrar_indice_coluna(colunas_disponiveis, ['DT_OFFLINE', 'DATA OFFLINE', 'OFFLINE'])
            col_data = st.selectbox("Coluna de Data", colunas_disponiveis, index=idx_data)
        with c2:
            idx_prio = encontrar_indice_coluna(colunas_disponiveis, ['PRIORIDADE OFICIAL', 'PRIORIDADE'])
            col_prioridade = st.selectbox("Coluna de Prioridade", colunas_disponiveis, index=idx_prio)
        with c3:
            idx_fila = encontrar_indice_coluna(colunas_disponiveis, ['NR_FILA_BTO', 'FILA_ID', 'FILA BTO'])
            col_fila = st.selectbox("Coluna de Fila (ID)", colunas_disponiveis, index=idx_fila)
        with c4:
            idx_cod = encontrar_indice_coluna(colunas_disponiveis, ['CD_PRODUTO', 'CODIGO_PRODUTO'])
            col_codigo = st.selectbox("Coluna de Código de Produto", colunas_disponiveis, index=idx_cod)
        with c5:
            idx_dealer = encontrar_indice_coluna(colunas_disponiveis, ['DL01', 'DEALER', 'CLIENTE', 'CLI'])
            col_dealer = st.selectbox("Coluna de Dealer / Cliente", colunas_disponiveis, index=idx_dealer)
            
        st.divider()

        if st.button("🚀 Processar Sequenciamento", type="primary"):
            with st.spinner("Gerando matriz e construindo visual gerencial..."):
                df_ok = processar_sequenciamento(df_original, col_data, col_prioridade, col_fila, col_codigo, col_dealer)
                
                st.subheader("Pré-visualização da Aba OK")
                if df_ok.empty:
                    st.warning("Nenhuma alteração de alto benefício encontrada.")
                else:
                    st.dataframe(df_ok)
                
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    df_original.to_excel(writer, sheet_name='Base Original', index=False)
                    df_ok.to_excel(writer, sheet_name='OK', index=False)
                    
                    workbook  = writer.book
                    worksheet = writer.sheets['OK']
                    
                    worksheet.freeze_panes(1, 1)
                    
                    if not df_ok.empty:
                        (max_row, max_col) = df_ok.shape
                        worksheet.autofilter(0, 0, max_row, max_col - 1)
                    
                    formato_cabecalho = workbook.add_format({
                        'bg_color': '#1F4E78',
                        'font_color': 'white',
                        'bold': True,
                        'border': 1,
                        'valign': 'vcenter'
                    })
                    
                    formato_zebra_1 = workbook.add_format({'bg_color': '#FFFFFF', 'border': 1})
                    formato_zebra_2 = workbook.add_format({'bg_color': '#F2F2F2', 'border': 1})
                    
                    larguras = [12, 14, 14, 18, 25, 20, 15, 20, 14, 25, 20, 15, 20, 15, 18, 15]
                    for col_num, value in enumerate(df_ok.columns.values):
                        worksheet.write(0, col_num, value, formato_cabecalho)
                        w = larguras[col_num] if col_num < len(larguras) else 15
                        worksheet.set_column(col_num, col_num, w)

                    if not df_ok.empty:
                        codigo_anterior = None
                        cor_atual = formato_zebra_1
                        
                        for row_num in range(len(df_ok)):
                            codigo_atual = df_ok.iloc[row_num]['Código do Produto']
                            
                            if codigo_atual != codigo_anterior:
                                cor_atual = formato_zebra_2 if cor_atual == formato_zebra_1 else formato_zebra_1
                                codigo_anterior = codigo_atual
                            
                            for col_num in range(len(df_ok.columns)):
                                val = df_ok.iloc[row_num, col_num]
                                if pd.isna(val): val = ""
                                worksheet.write(row_num + 1, col_num, val, cor_atual)
                
                st.download_button(
                    label="📥 Baixar Relatório Consolidado (Excel)",
                    data=buffer.getvalue(),
                    file_name="Relatorio_Smart_Order_Management.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        
    except Exception as e:
        st.error(f"Ocorreu um erro ao processar o arquivo: {e}")
