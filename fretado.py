import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import os
import warnings
import io

# Silencia avisos de formatação do Excel
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# --- CAMINHOS FIXOS CONFIGURADOS ---
BASE_SETORES_CSV = "facilities/fretado/Fretado Fab BA 2026(SETORES E EMPRESAS).csv"
BASE_GERAL_EXCEL = "facilities/fretado/Fretado Fab BA 2026.xlsx"

# --- FUNÇÕES DE APOIO ---

def calcular_duracao(inicio, fim):
    """Calcula a diferença entre dois horários HH:MM, tratando virada de dia."""
    try:
        fmt = "%H:%M"
        t1_str = str(inicio).strip()[:5]
        t2_str = str(fim).strip()[:5]
        
        t1 = datetime.strptime(t1_str, fmt)
        t2 = datetime.strptime(t2_str, fmt)
        
        if t2 < t1:
            t2 += timedelta(days=1)
            
        diff = t2 - t1
        horas, segundos = divmod(diff.seconds, 3600)
        minutos = (diff.seconds // 60) % 60
        return f"{horas:02d}:{minutos:02d}"
    except:
        return "00:00"

def e_numero(valor):
    """Verifica se o valor na Coluna D (índice 3) é um número (quantidade)."""
    if pd.isna(valor): return False
    try:
        float(str(valor).replace(',', '.'))
        return True
    except:
        return False
    
def filtrar_dataframe(df, chave_aba):
    """Cria filtros dinâmicos usando um checkbox estável para cada coluna do DataFrame."""
    df_filtrado = df.copy()
    
    # CORREÇÃO: Substituído st.expander por st.checkbox com key.
    # Inicia como False (fechado) e não fecha sozinho após a primeira filtragem.
    if st.checkbox("🔍 Filtrar Colunas", value=False, key=f"ativar_filtros_{chave_aba}"):
        st.markdown("---") # Pequena linha divisória estética
        # Cria campos de seleção para cada coluna de forma vertical e limpa
        for col in df.columns:
            valores_unicos = sorted(df[col].dropna().unique().tolist())
            selecionados = st.multiselect(
                f"Filtrar por {col}", 
                options=valores_unicos, 
                key=f"filter_{chave_aba}_{col}"
            )
            if selecionados:
                df_filtrado = df_filtrado[df_filtrado[col].isin(selecionados)]
        st.markdown("---")
        
    return df_filtrado

def processar_logistica_pepsico(arquivo_excel, arquivo_setores_csv, arquivo_base_geral):
    # 1. CARREGAR REFERÊNCIA DE SETORES (UTF-8)
    if not os.path.exists(arquivo_setores_csv):
        st.error(f"Arquivo base de setores não encontrado no servidor virtual: {arquivo_setores_csv}")
        return None, None, None
        
    try:
        # CORREÇÃO: Adicionado quoting=3 (csv.QUOTE_NONE) para fazer o Pandas ignorar o controle 
        # de aspas do parser e tratar tudo como texto puro. Isso evita o erro do Git/GitHub.
        df_setores_ref = pd.read_csv(
            arquivo_setores_csv, 
            sep=None, 
            engine='python', 
            encoding='utf-8-sig', 
            header=None,
            names=list(range(10)),
            quoting=3
        )
        # CORREÇÃO: Como ignoramos as aspas no parser, limpamos qualquer aspa residual textualmente
        # com o .str.replace('"', '') antes de registrar na lista de setores.
        lista_setores = df_setores_ref[2].dropna().astype(str).str.replace('"', '').str.strip().str.upper().unique().tolist()
    except Exception as e:
        st.error(f"Erro ao ler a base de setores fixa: {e}")
        return None, None, None

    # 2. CARREGAR BASE GERAL
    base_geral_nomes = []
    if arquivo_base_geral and os.path.exists(arquivo_base_geral):
        try:
            df_g = pd.read_excel(arquivo_base_geral)
            base_geral_nomes = df_g.iloc[:, 0].astype(str).str.strip().str.upper().unique().tolist()
        except Exception as e:
            st.warning(f"Não foi possível cruzar com a Base Geral: {e}")

    # 3. ANALISAR EXCEL DE ROTAS
    try:
        abas = pd.read_excel(arquivo_excel, sheet_name=None, header=None)
    except Exception as e:
        st.error(f"Erro ao processar o arquivo de Fretado enviado: {e}")
        return None, None, None

    relatorio_detalhado = []
    relatorio_resumido = []

    for nome_aba, df in abas.items():
        rota_atual = "N/A"
        temp_passageiros = []
        dentro_do_bloco = False

        for idx, row in df.iterrows():
            conteudo_linha = " ".join([str(item).upper() for item in row.values])
            
            if "PONTOS DE REFERÊNCIA" in conteudo_linha or "PONTOS DE REFERENCIA" in conteudo_linha:
                dentro_do_bloco = True
                if idx >= 2:
                    valor_rota = df.iloc[idx - 2, 5] 
                    rota_atual = str(valor_rota).strip() if pd.notna(valor_rota) else "N/A"
                continue

            if "CHEGADA NA FÁBRICA PEPSICO" in conteudo_linha or "CHEGADA NA FABRICA" in conteudo_linha:
                dentro_do_bloco = False
                continue

            if dentro_do_bloco:
                setor_lido = str(row[1]).strip().upper() if pd.notna(row[1]) else ""
                identificado_por_setor = setor_lido in lista_setores
                identificado_por_numero = e_numero(row[3]) 

                if identificado_por_setor or identificado_por_numero:
                    nome_func = str(row[2]).strip().upper() if pd.notna(row[2]) else ""
                    if nome_func and "HORÁRIO" not in nome_func and "HORARIO" not in nome_func:
                        temp_passageiros.append({
                            "aba": nome_aba,
                            "setor": setor_lido,
                            "nome": nome_func,
                            "rota": rota_atual,
                            "linha_idx": idx
                        })

        # Processamento dos dados da aba para os relatórios
        if temp_passageiros:
            df_aba_temp = pd.DataFrame(temp_passageiros)
            for r in df_aba_temp['rota'].unique():
                grupo = df_aba_temp[df_aba_temp['rota'] == r]
                idx_ini, idx_fim = grupo['linha_idx'].min(), grupo['linha_idx'].max()

                # Horários (Coluna E - índice 4)
                h_ini = df.iloc[idx_ini - 1, 4] if idx_ini > 0 else "00:00"
                h_fim = df.iloc[idx_fim + 1, 4] if idx_fim < len(df)-1 else "00:00"
                
                h_ini = str(h_ini) if pd.notna(h_ini) else "00:00"
                h_fim = str(h_fim) if pd.notna(h_fim) else "00:00"
                duracao = calcular_duracao(h_ini, h_fim)

                for _, p in grupo.iterrows():
                    na_base = "SIM" if p['nome'] in base_geral_nomes else "NÃO"
                    relatorio_detalhado.append({
                        "ABA ORIGEM": p['aba'],
                        "SETOR/EMPRESA": p['setor'],
                        "NOME FUNCIONÁRIO": p['nome'],
                        "ROTA": p['rota'],
                        "HORA PARTIDA": h_ini,
                        "HORA CHEGADA FÁBRICA": h_fim,
                        "TEMPO TOTAL ROTA": duracao,
                        "NA BASE GERAL?": na_base if base_geral_nomes else "NÃO TESTADO"
                    })
                
                relatorio_resumido.append({
                    "ABA ORIGEM": nome_aba,
                    "ROTA": r,
                    "QTD FUNCIONÁRIOS": len(grupo),
                    "HORA PARTIDA": h_ini,
                    "HORA CHEGADA FÁBRICA": h_fim,
                    "TEMPO TOTAL ROTA": duracao
                })

    # ALTERAÇÃO: Agora retornamos os DataFrames criados E o arquivo compilado em bytes
    if relatorio_detalhado:
        df_detalhe = pd.DataFrame(relatorio_detalhado)
        df_resumo = pd.DataFrame(relatorio_resumido)
        
        buffer_memoria = io.BytesIO()
        with pd.ExcelWriter(buffer_memoria, engine='openpyxl') as writer:
            df_detalhe.to_excel(writer, sheet_name="DETALHADO", index=False)
            df_resumo.to_excel(writer, sheet_name="RESUMO_ROTAS", index=False)
        return df_detalhe, df_resumo, buffer_memoria.getvalue()
        
    return None, None, None

# --- DESIGN DA INTERFACE WEB ---
st.set_page_config(page_title="Painel de Fretados - PepsiCo", page_icon="🚀", layout="centered")
st.title("Automação de Relatórios de Fretados")
st.markdown("Insira o arquivo diário enviado para consolidar as informações das rotas automaticamente.")

arquivo_fretado = st.file_uploader("Selecione a planilha de Fretado (.xlsx)", type=["xlsx"])

# Controle para limpar a memória se o usuário trocar ou remover o arquivo carregado
if "ultimo_arquivo" not in st.session_state:
    st.session_state.ultimo_arquivo = None

if arquivo_fretado != st.session_state.ultimo_arquivo:
    st.session_state.ultimo_arquivo = arquivo_fretado
    if "df_detalhe" in st.session_state: del st.session_state.df_detalhe
    if "df_resumo" in st.session_state: del st.session_state.df_resumo
    if "resultado_bytes" in st.session_state: del st.session_state.resultado_bytes

if arquivo_fretado is not None:
    st.success("Planilha carregada com sucesso!")
    
    # O botão agora serve EXCLUSIVAMENTE para processar e salvar o resultado no estado da sessão
    if st.button("🚀 PROCESSAR PLANILHA", use_container_width=True):
        with st.spinner("Varrendo abas e aplicando regras de negócio..."):
            df_detalhe, df_resumo, resultado_bytes = processar_logistica_pepsico(arquivo_fretado, BASE_SETORES_CSV, BASE_GERAL_EXCEL)
            
            if resultado_bytes:
                st.session_state.df_detalhe = df_detalhe
                st.session_state.df_resumo = df_resumo
                st.session_state.resultado_bytes = resultado_bytes
                st.balloons()
                st.success("Análise finalizada!")
            else:
                st.warning("Nenhum dado compatível foi processado. Verifique os critérios da planilha.")

    # Se os dados já foram processados, exibe as tabelas e permite filtragem sem resetar
    if "resultado_bytes" in st.session_state:
        # 1. Botão de Download permanece fixo no topo dos resultados
        nome_saida = f"RELATORIO_ROTAS_PEPSICO_{datetime.now().strftime('%d_%m_%H%M')}.xlsx"
        st.download_button(
            label="📥 DOWNLOAD DO RELATÓRIO COMPLETO (.XLSX)",
            data=st.session_state.resultado_bytes,
            file_name=nome_saida,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        
        # 2. Exibição das abas visuais na tela puxando os dados da sessão
        st.markdown("### 📊 Visualização das Tabelas")
        aba_detalhado, aba_resumo = st.tabs(["📋 Relatório Detalhado", "📈 Resumo das Rotas"])
        
        with aba_detalhado:
            # Filtra e exibe o DataFrame detalhado salvo na sessão
            df_detalhe_filtrado = filtrar_dataframe(st.session_state.df_detalhe, "detalhe")
            st.dataframe(df_detalhe_filtrado, use_container_width=True)
            
        with aba_resumo:
            # Filtra e exibe o DataFrame resumo salvo na sessão
            df_resumo_filtrado = filtrar_dataframe(st.session_state.df_resumo, "resumo")
            st.dataframe(df_resumo_filtrado, use_container_width=True)