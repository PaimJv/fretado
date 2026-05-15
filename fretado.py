import pandas as pd
from datetime import datetime, timedelta
import os
import warnings
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk

# Silencia avisos de formatação do Excel para manter o terminal limpo
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

def processar_logistica_pepsico(arquivo_excel, arquivo_setores_csv, arquivo_base_geral, caminho_salvamento):
    # 1. CARREGAR REFERÊNCIA DE SETORES (UTF-8)
    try:
        df_setores_ref = pd.read_csv(arquivo_setores_csv, sep=None, engine='python', encoding='utf-8-sig', header=None)
        lista_setores = df_setores_ref[2].dropna().astype(str).str.strip().str.upper().unique().tolist()
        print(f"✅ {len(lista_setores)} setores carregados da base de referência.")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao ler arquivo de setores fixo:\n{arquivo_setores_csv}\n\nDetalhe: {e}")
        return

    # 2. CARREGAR BASE GERAL
    base_geral_nomes = []
    if arquivo_base_geral and os.path.exists(arquivo_base_geral):
        try:
            df_g = pd.read_excel(arquivo_base_geral)
            base_geral_nomes = df_g.iloc[:, 0].astype(str).str.strip().str.upper().unique().tolist()
            print(f"✅ Base geral carregada com sucesso.")
        except Exception as e:
            print(f"⚠️ Base geral não carregada: {e}")

    # 3. ANALISAR EXCEL DE ROTAS
    print(f"🔍 Analisando todas as abas de: {arquivo_excel}...")
    try:
        abas = pd.read_excel(arquivo_excel, sheet_name=None, header=None)
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao abrir o Excel de Rotas:\n{e}")
        return

    relatorio_detalhado = []
    relatorio_resumido = []

    for nome_aba, df in abas.items():
        rota_atual = "N/A"
        temp_passageiros = []
        dentro_do_bloco = False

        for idx, row in df.iterrows():
            conteudo_linha = " ".join([str(item).upper() for item in row.values])
            
            # REGRA: Localizar Rota (2 linhas acima de PONTOS DE REFERÊNCIA na Coluna F)
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

                # 1. Adicionar ao Detalhado
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
                
                # 2. Adicionar ao Resumido (Uma linha por rota da aba)
                relatorio_resumido.append({
                    "ABA ORIGEM": nome_aba,
                    "ROTA": r,
                    "QTD FUNCIONÁRIOS": len(grupo),
                    "HORA PARTIDA": h_ini,
                    "HORA CHEGADA FÁBRICA": h_fim,
                    "TEMPO TOTAL ROTA": duracao
                })

    # 4. EXPORTAÇÃO COM DUAS ABAS
    if relatorio_detalhado:
        df_detalhe = pd.DataFrame(relatorio_detalhado)
        df_resumo = pd.DataFrame(relatorio_resumido)
        
        try:
            with pd.ExcelWriter(caminho_salvamento, engine='openpyxl') as writer:
                df_detalhe.to_excel(writer, sheet_name="DETALHADO", index=False)
                df_resumo.to_excel(writer, sheet_name="RESUMO_ROTAS", index=False)
            messagebox.showinfo("Sucesso", f"Relatório gerado com sucesso em:\n{caminho_salvamento}")
        except Exception as e:
            messagebox.showerror("Erro ao Salvar", f"Não foi possível salvar o arquivo. Verifique se ele está fechado.\nErro: {e}")
    else:
        messagebox.showwarning("Aviso", "Nenhum dado correspondente foi encontrado para gerar o relatório.")

# --- INTERFACE GRÁFICA (MENU) ---

class AplicacaoLogistica:
    def __init__(self, root):
        self.root = root
        self.root.title("Painel de Logística - PepsiCo")
        self.root.geometry("620x260")
        self.root.resizable(False, False)
        
        style = ttk.Style()
        style.theme_use('clam')
        
        frame = ttk.Frame(root, padding="20")
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Título Principal
        lbl_titulo = ttk.Label(frame, text="Automação de Relatórios de Fretados", font=("Helvetica", 13, "bold"))
        lbl_titulo.grid(row=0, column=0, columnspan=3, pady=(0, 20), sticky="w")
        
        # 1. Campo Arquivo de Fretado (Rotas)
        ttk.Label(frame, text="Arquivo Fretado (.xlsx):").grid(row=1, column=0, sticky="w", pady=5)
        self.txt_rotas = ttk.Entry(frame, width=42)
        self.txt_rotas.grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(frame, text="Selecionar", command=self.selecionar_rotas).grid(row=1, column=2, pady=5)
        
        # 2. Campo Destino de Salvamento
        ttk.Label(frame, text="Salvar Novo Arquivo em:").grid(row=2, column=0, sticky="w", pady=5)
        self.txt_salvamento = ttk.Entry(frame, width=42)
        self.txt_salvamento.grid(row=2, column=1, padx=5, pady=5)
        ttk.Button(frame, text="Alterar Local", command=self.definir_salvamento).grid(row=2, column=2, pady=5)
        
        # Divisor Visual Corrido (Sem bugs de Statics)
        ttk.Separator(frame, orient='horizontal').grid(row=3, column=0, columnspan=3, sticky="ew", pady=15)
        
        # Botão Processar
        self.btn_processar = ttk.Button(frame, text="🚀 CONFIGURAR E PROCESSAR FRETADO", command=self.disparar_processamento)
        self.btn_processar.grid(row=4, column=0, columnspan=3, ipady=6, sticky="ew")

    def selecionar_rotas(self):
        caminho = filedialog.askopenfilename(title="Selecionar Planilha de Fretado", filetypes=[("Excel files", "*.xlsx")])
        if caminho:
            self.txt_rotas.delete(0, tk.END)
            self.txt_rotas.insert(0, caminho)
            
            # Autopreenche uma sugestão de salvamento baseada no local do arquivo de origem
            if not self.txt_salvamento.get():
                sugestao_nome = os.path.join(os.path.dirname(caminho), f"RELATORIO_ROTAS_PEPSICO_{datetime.now().strftime('%d_%m_%H%M')}.xlsx")
                self.txt_salvamento.insert(0, sugestao_nome)

    def definir_salvamento(self):
        caminho = filedialog.asksaveasfilename(title="Definir Local de Destino", defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if caminho:
            self.txt_salvamento.delete(0, tk.END)
            self.txt_salvamento.insert(0, caminho)

    def disparar_processamento(self):
        rotas = self.txt_rotas.get()
        salvamento = self.txt_salvamento.get()
        
        if not rotas:
            messagebox.showwarning("Campo Vazio", "Por favor, selecione a planilha de fretado de entrada.")
            return
        if not salvamento:
            messagebox.showwarning("Campo Vazio", "Por favor, defina um nome e local para salvar o relatório consolidado.")
            return
            
        # Executa o processamento injetando os caminhos internos fixos
        processar_logistica_pepsico(rotas, BASE_SETORES_CSV, BASE_GERAL_EXCEL, salvamento)

if __name__ == "__main__":
    root = tk.Tk()
    app = AplicacaoLogistica(root)
    root.mainloop()