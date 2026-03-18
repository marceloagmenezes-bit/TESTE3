import streamlit as st
import pandas as pd
import io
import re

# ==========================================
# 1. FUNÇÕES DE TRATAMENTO E REGRAS DE NEGÓCIO
# ==========================================

def get_prioridade_str(val):
    """ Garante que a prioridade seja tratada puramente como número """
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

# ==========================================
# 2. MOTOR DE SEQUENCIAMENTO (INVERSÃO DIRETA 1-TO-1)
# ==========================================

def processar_sequenciamento(df_base, col_data, col_prioridade, col_fila, col_codigo, col_dealer):
    df_processo = df_base.copy()
    
    df_processo[col_data] = pd.to_datetime(df_processo[col_data], errors='coerce')
    df_processo['Descrição da Prioridade'] = df_processo[col_prioridade].apply(classificar_prioridade)
    
    # Coluna oculta para garantir a ordenação matemática da Prioridade
    df_processo['Prioridade_Num'] = pd.to_numeric(
        df_processo[col_prioridade].apply(get_prioridade_str), errors='coerce'
    ).fillna(999999999) 
    
    # Identifica as colunas baseadas nas letras vitais de suporte
    letras_suporte = ['F', 'Q', 'Z', 'C', 'AX', 'CQ', 'I']
    colunas_suporte = []
    for letra in letras_suporte:
        nome_col = excel_col_to_name(df_processo, letra)
        if nome_col:
            colunas_suporte.append((letra, nome_col))
    
    relatorio_ok = []

    # Agrupa pelo Código de Produto
    for codigo, grupo in df_processo.groupby(col_codigo):
        
        # Converte para lista de dicionários para manipularmos o estado atual
        estado_atual = grupo.to_dict('records')
        
        # Ordena a linha do tempo (slots cronológicos)
        estado_atual.sort(key=lambda x: x[col_data])
        
        # Varre os slots do mais próximo ao mais distante
        for i in range(len(estado_atual)):
            slot_data = estado_atual[i][col_data]
            ocupante_atual = estado_atual[i]
            
            # Identifica quem deveria estar neste slot (O mais prioritário dentre os que sobraram)
            # Ordena os itens restantes (do índice 'i' em diante) por Prioridade -> Data
            candidatos = sorted(estado_atual[i:], key=lambda x: (x['Prioridade_Num'], x[col_data]))
            ideal = candidatos[0]
            
            # Se a fila ideal já for o ocupante atual, não faz nada
            if ideal[col_fila] == ocupante_atual[col_fila]:
                continue
                
            data_orig_ideal = ideal[col_data]
            
            # --- VERIFICAÇÃO DE BENEFÍCIO REAL (REGRAS 3 E 5) ---
            ignorar = False
            p_desc = ideal['Descrição da Prioridade']
            p_str = get_prioridade_str(ideal[col_prioridade])
            
            # Regra 3: Tolerâncias
            if p_desc != 'BUFF/RES' and len(p_str) >= 5:
                quinto_digito = p_str[4]
                if quinto_digito in ['1', '2', '4', '6']:
                    if mesma_semana(data_orig_ideal, slot_data): ignorar = True
                else:
                    if mesmo_mes(data_orig_ideal, slot_data): ignorar = True
            
            # Regra 5: Mesmo Dealer / Cliente
            dealer_ideal = ideal.get(col_dealer)
            dealer_ocupante = ocupante_atual.get(col_dealer)
            if pd.notna(dealer_ideal) and str(dealer_ideal).strip() != "":
                if dealer_ideal == dealer_ocupante:
                    ignorar = True
            
            # --- EXECUÇÃO DA INVERSÃO DIRETA (SWAP 1-TO-1) ---
            if not ignorar:
                # 1. Registra a Fila Ideal tomando o lugar do Ocupante Atual
                linha_1 = {
                    'Identificação da fila (Col A)': ideal[col_fila],
                    'Data correta sugerida': slot_data,
                    'Data Original': data_orig_ideal,
                    'Fila bloqueadora (Col E)': ocupante_atual[col_fila],
                    'Descrição da Prioridade': ideal['Descrição da Prioridade'],
                    'Prioridade Real': ideal[col_prioridade],
                    'Código do Produto': ideal[col_codigo],
                    'Dealer Sugerido': ideal.get(col_dealer),
                    'Dealer Bloqueador': ocupante_atual.get(col_dealer)
                }
                for letra, nome_col in colunas_suporte:
                    linha_1[f"Col {letra} - {nome_col}"] = ideal.get(nome_col)
                relatorio_ok.append(linha_1)
                
                # 2. Registra o Ocupante Atual sendo empurrado para o antigo lugar da Fila Ideal
                linha_2 = {
                    'Identificação da fila (Col A)': ocupante_atual[col_fila],
                    'Data correta sugerida': data_orig_ideal,
                    'Data Original': slot_data,
                    'Fila bloqueadora (Col E)': ideal[col_fila],
                    'Descrição da Prioridade': ocupante_atual['Descrição da Prioridade'],
                    'Prioridade Real': ocupante_atual[col_prioridade],
                    'Código do Produto': ocupante_atual[col_codigo],
                    'Dealer Sugerido': ocupante_atual.get(col_dealer),
                    'Dealer Bloqueador': ideal.get(col_dealer)
                }
                for letra, nome_col in colunas_suporte:
                    linha_2[f"Col {letra} - {nome_col}"] = ocupante_atual.get(nome_col)
                relatorio_ok.append(linha_2)
                
                # 3. Atualiza o Estado (Simula a troca para as próximas datas)
                idx_ideal = estado_atual.index(ideal)
                
                ideal_copia = ideal.copy()
                ocupante_copia = ocupante_atual.copy()
                
                ideal_copia[col_data] = slot_data
                ocupante_copia[col_data] = data_orig_ideal
                
                estado_atual[i] = ideal_copia
                estado_atual[idx_ideal] = ocupante_copia

    df_final = pd.DataFrame(relatorio_ok)
    
    # ORDENAÇÃO: 1) Código do produto, 2) Data correta sugerida
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
        if arquivo_upload.name.endswith('csv'):
            df_original = pd.read_csv(arquivo_upload)
        else:
            df_original = pd.read_excel(arquivo_upload)
            
        colunas_disponiveis = df_original.columns.tolist()
        
        st.success("Arquivo carregado com sucesso!")
        st.markdown("### 🔍 Mapeamento de Colunas Principais")
        
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
            possiveis_dealer = [c for c in colunas_disponiveis if 'DEALER' in str(c).upper() or 'CLI' in str(c).upper()]
            idx_dealer = colunas_disponiveis.index(possiveis_dealer[0]) if possiveis_dealer else 0
            col_dealer = st.selectbox("Coluna de Dealer / Cliente", colunas_disponiveis, index=idx_dealer)
            
        st.divider()

        if st.button("🚀 Processar Sequenciamento", type="primary"):
            with st.spinner("Analisando swaps diretos e cortando cadeias longas..."):
                df_ok = processar_sequenciamento(df_original, col_data, col_prioridade, col_fila, col_codigo, col_dealer)
                
                st.subheader("Pré-visualização da Aba OK")
                if df_ok.empty:
                    st.warning("Nenhuma alteração de alto benefício encontrada.")
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
                    file_name="Relatorio_Sequenciamento_V5.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        
    except Exception as e:
        st.error(f"Ocorreu um erro ao processar o arquivo: {e}")
