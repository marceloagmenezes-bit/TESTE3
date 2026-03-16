import streamlit as st
import pandas as pd
import io
import os

# Função para converter letra de Excel em número de índice
def excel_col_to_num(col_str):
    exp = 0
    col_num = 0
    for char in reversed(col_str.upper()):
        col_num += (ord(char) - ord('A') + 1) * (26**exp)
        exp += 1
    return col_num - 1

st.set_page_config(page_title="AGCO - Sequenciador v3", page_icon="🚜", layout="wide")

# Logo
if os.path.exists("imagem"):
    arquivos = [f for f in os.listdir("imagem") if f.lower().endswith(".jpg")]
    if arquivos: st.image(os.path.join("imagem", arquivos[0]), width=150)

st.title("🚜 Validador de Prioridade Real vs Cronológica")
st.markdown("Abordagem: Mapeamento Direto (Menor Prioridade obrigatoriamente na Data mais antiga)")

uploaded_file = st.file_uploader("Carregue o Excel padrão", type=["xlsx"])

if uploaded_file:
    try:
        # Carrega os dados ignorando a primeira linha (cabeçalho)
        df_raw = pd.read_excel(uploaded_file, header=None)
        
        # Índices das colunas
        c_ct = excel_col_to_num("CT")
        c_j  = excel_col_to_num("J")
        c_d  = excel_col_to_num("D")
        c_ag = excel_col_to_num("AG")
        
        # Colunas extras
        extras = ["F", "Q", "Z", "C", "AX", "CQ", "I"]
        idx_extras = {l: excel_col_to_num(l) for l in extras}

        if st.button("🚀 Executar Mapeamento de Prioridades"):
            # Limpeza de dados para garantir que a matemática funcione
            df = df_raw.iloc[1:].copy()
            df[c_j] = pd.to_datetime(df[c_j], errors='coerce')
            df[c_d] = pd.to_numeric(df[c_d], errors='coerce')
            
            # Remove linhas onde AG, CT ou Data estão vazios para não quebrar a lógica
            df = df.dropna(subset=[c_ag, c_ct, c_j])

            resultado_final = []

            # Agrupar por Código do Produto (AG)
            for produto, grupo in df.groupby(c_ag):
                # 1. Pegar as datas disponíveis para este produto (da menor para a maior)
                datas_disponiveis = sorted(grupo[c_j].tolist())
                
                # 2. Pegar as filas (CT) ordenadas pela Prioridade Real D (do menor para o maior)
                # Se a prioridade for igual, usamos o CT como desempate
                grupo_ideal = grupo.sort_values(by=[c_d, c_ct], ascending=[True, True])
                filas_ordenadas = grupo_ideal[c_ct].tolist()
                prioridades_ordenadas = grupo_ideal[c_d].tolist()

                # 3. Criar o mapa "Onde cada fila DEVERIA estar"
                # Fila de menor D -> Data mais antiga
                mapa_correto = {}
                for i in range(len(filas_ordenadas)):
                    mapa_correto[filas_ordenadas[i]] = datas_disponiveis[i]

                # 4. Verificar quem está no lugar errado hoje
                for _, linha in grupo.iterrows():
                    fila_id = linha[c_ct]
                    data_atual = linha[c_j]
                    data_que_deveria_ser = mapa_correto[fila_id]

                    # Se a data atual for diferente da data que o mapa diz ser a correta
                    if data_atual != data_que_deveria_ser:
                        # Descobrir quem está ocupando a data que deveria ser minha
                        quem_esta_na_minha_data = grupo[grupo[c_j] == data_que_deveria_ser][c_ct].values[0]
                        prioridade_de_quem_esta_la = grupo[grupo[c_j] == data_que_deveria_ser][c_d].values[0]

                        resultado_final.append({
                            "Número da fila (CT)": fila_id,
                            "Prioridade atual (D)": linha[c_d],
                            "Data Offline atual (J)": data_atual.strftime('%d/%m/%Y'),
                            "Data correta sugerida": data_que_deveria_ser.strftime('%d/%m/%Y'),
                            "Fila que ocupa essa data": quem_esta_na_minha_data,
                            "Prioridade da fila que ocupa": prioridade_de_quem_esta_la,
                            "Produto (AG)": produto,
                            "F": linha[idx_extras["F"]],
                            "Q": linha[idx_extras["Q"]],
                            "Z": linha[idx_extras["Z"]],
                            "C": linha[idx_extras["C"]],
                            "AX": linha[idx_extras["AX"]],
                            "CQ": linha[idx_extras["CQ"]],
                            "I": linha[idx_extras["I"]]
                        })

            if resultado_final:
                df_res = pd.DataFrame(resultado_final)
                
                # Remove duplicatas (cada erro de inversão aparece para as duas filas envolvidas, 
                # vamos focar na necessidade de mudança)
                df_res = df_res.drop_duplicates(subset=["Número da fila (CT)"])

                st.warning(f"Foram identificadas {len(df_res)} filas fora da ordem de prioridade real.")
                
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_raw.to_excel(writer, sheet_name='Base Original', index=False, header=False)
                    df_res.to_excel(writer, sheet_name='OK', index=False)
                
                st.success("✅ Processamento concluído!")
                st.download_button("📥 Baixar Planilha Consistente", output.getvalue(), "sequenciamento_agco.xlsx")
                st.dataframe(df_res)
            else:
                st.success("✅ Nenhuma inversão encontrada. O plano de produção está 100% alinhado com as prioridades.")

    except Exception as e:
        st.error(f"Erro no processamento: {e}")