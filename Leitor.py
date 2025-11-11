"""
Extrator de Guias em PDF
Versão: 1.0
Autor: Natanael Silva
Ano: 2025
Licença: MIT

Descrição:
Este software permite selecionar múltiplos arquivos PDF contendo guias
e extrair automaticamente informações como:
- Número da Guia
- Vencimento
- Total a Pagar
- Processo/Protocolo
- Código de Barras

O resultado é exportado para um arquivo Excel (.xlsx).
"""

# extrator_guias.py
# --- IMPORTAÇÕES ---
import fitz                # pip install pymupdf
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import re
import pandas as pd       # pip install pandas
import os
import sys

# =========================
# FUNÇÃO: OCULTAR PROGRESSO
# =========================
def ocultar_progresso(mensagem=""):
    """
    Atualiza a mensagem final, reseta a barra e esconde widgets de progresso.
    """
    progress_label_var.set(mensagem)
    percent_label_var.set("")
    progress_bar['value'] = 0

    # Se os widgets estiverem visíveis, esconda-os
    try:
        lbl_progress.pack_forget()
        progress_bar.pack_forget()
        lbl_percent.pack_forget()
    except Exception:
        pass

    janela.update_idletasks()

# =========================
# FUNÇÃO: EXTRAÇÃO PRINCIPAL
# =========================
def extrair_dados_massa():
    """
    Abre diálogo para selecionar múltiplos PDFs e extrai:
    - Numero Guia (10 dígitos)
    - Vencimento (DD/MM/AAAA)
    - Total a Pagar (último valor com vírgula XX,XX)
    - Processo / Protocolo (variações)
    - Código de Barras (padrões iniciando com 8)
    Salva tudo em um Excel.
    """

    caminhos_pdf = filedialog.askopenfilenames(
        title="Selecione os arquivos PDF",
        filetypes=[("Arquivos PDF", "*.pdf")]
    )
    if not caminhos_pdf:
        return

    total_arquivos = len(caminhos_pdf)
    if total_arquivos == 0:
        return

    # Exibe widgets de progresso
    lbl_progress.pack(pady=(15,5), fill='x')
    progress_bar.pack(fill='x', padx=10)
    lbl_percent.pack(pady=(5,5), fill='x')

    progress_bar['maximum'] = total_arquivos
    progress_bar['value'] = 0
    progress_label_var.set(f"Iniciando... 0 de {total_arquivos} arquivos")
    percent_label_var.set("0%")
    janela.update_idletasks()

    resultados = []
    arquivos_falhados = []

    # Helpers de regex
    def buscar(padrao, texto_busca):
        r = re.search(padrao, texto_busca, re.IGNORECASE | re.DOTALL)
        return r.group(1).strip() if r else ""

    def buscar_barcode(texto_busca):
        # Regra 1: 4 blocos (12 dígitos cada) começando com 8 no primeiro bloco
        barcode_pattern = r"(8\d{11}\s+\d{12}\s+\d{12}\s+\d{12})"
        barcode_match = re.search(barcode_pattern, texto_busca)
        if barcode_match:
            return re.sub(r'\s+', '', barcode_match.group(1))

        # Regra 2: linha única grande começando com 8
        barcode_match_alt = re.search(r"^(8[\d\s]{40,})$", texto_busca, re.MULTILINE)
        if barcode_match_alt:
            return re.sub(r'\s+', '', barcode_match_alt.group(1))

        # Regra 3: qualquer sequência contínua de 44 dígitos que comece com 8
        barcode_match_alt2 = re.search(r"(8\d{43})", re.sub(r'\s+', '', texto_busca))
        if barcode_match_alt2:
            return barcode_match_alt2.group(1)

        return ""

    # Loop principal
    for i, caminho_pdf in enumerate(caminhos_pdf):
        progresso_atual = i + 1
        nome_arquivo = os.path.basename(caminho_pdf)

        percentual = int((progresso_atual / total_arquivos) * 100)
        progress_bar['value'] = progresso_atual
        progress_label_var.set(f"Processando: {progresso_atual} de {total_arquivos} — {nome_arquivo}")
        percent_label_var.set(f"{percentual}%")
        janela.update_idletasks()

        texto = ""
        try:
            # Abre PDF com PyMuPDF
            with fitz.open(caminho_pdf) as doc:
                for pagina in doc:
                    # get_text("text") preserva a ordem textual
                    texto_pagina = pagina.get_text("text")
                    if texto_pagina:
                        texto += texto_pagina + "\n"

            if not texto.strip():
                arquivos_falhados.append(f"{nome_arquivo} (arquivo sem texto - possivelmente imagem)")
                continue

            # TENTATIVA 1: padrões por rótulos (mais confiáveis)
            guia_val = buscar(r"N\.? ?MERO\s*(?:DA\s*)?GUIA.*?([\d]{10})", texto)
            if not guia_val:
                guia_val = buscar(r"\b([\d]{10})\b", texto)  # fallback: primeiro conjunto de 10 dígitos

            # Vencimento (procura por rótulos ou pela primeira data)
            venc_val = buscar(r"02\s*-\s*VENCIMENTO\s*([\d]{2}\/[\d]{2}\/[\d]{4})", texto)
            if not venc_val:
                venc_val = buscar(r"DATA DE VALIDADE\s*([\d]{2}\/[\d]{2}\/[\d]{4})", texto)
            if not venc_val:
                venc_val = buscar(r"([\d]{2}\/[\d]{2}\/[\d]{4})", texto)

            # Total a pagar (procura rótulo ou pega o último valor monetário do documento)
            total_val = buscar(r"26\s*-\s*TOTAL\s*A\s*PAGAR\s*([\d\.\,]+)", texto)
            if not total_val:
                valores_monetarios = re.findall(r"(\d{1,3}(?:\.\d{3})*,\d{2})", texto)
                if valores_monetarios:
                    total_val = valores_monetarios[-1]  # último valor é o total geralmente
                else:
                    # fallback simples: qualquer X,XX
                    tmp = re.findall(r"(\d+,\d{2})", texto)
                    total_val = tmp[-1] if tmp else ""

            # Processo / Protocolo: diferentes padrões
            processo_val = ""
            p = re.search(r"PROCESSO(?:\s*SEI)?\s*[:\-]?\s*([\d\.\/\-]+)", texto, re.IGNORECASE)
            if p:
                processo_val = p.group(1).strip()
            else:
                p2 = re.search(r"PROTOCOLO\s*[:\-]?\s*([\d\.\/\-]+)", texto, re.IGNORECASE)
                if p2:
                    processo_val = p2.group(1).strip()

            # Código de barras
            barcode_val = buscar_barcode(texto)

            # Se tudo estiver vazio (falha), grava debug para análise
            if not (guia_val or venc_val or total_val):
                inicio_texto = texto.strip().replace("\n", " ")[:600]
                arquivos_falhados.append(f"ARQUIVO: {nome_arquivo} - Falha: dados não encontrados. Texto (início): {inicio_texto}...")
                continue

            resultados.append({
                "Arquivo Origem": nome_arquivo,
                "Numero Guia": guia_val,
                "Vencimento": venc_val,
                "Total a Pagar": total_val,
                "Processo/Protocolo": processo_val,
                "Codigo de Barras": barcode_val
            })

        except Exception as e:
            arquivos_falhados.append(f"{nome_arquivo} (Erro: {e})")

    # Após processar todos
    if not resultados:
        msg_falha = "Não foi possível extrair dados de nenhum dos arquivos."
        if arquivos_falhados:
            msg_falha += "\n\n--- LOG DE ERROS ---\n" + "\n".join(arquivos_falhados)
        messagebox.showerror("Falha Total", msg_falha)
        ocultar_progresso()
        return

    # Pergunta onde salvar
    salvar = filedialog.asksaveasfilename(
        title="Salvar Compilado Excel",
        defaultextension=".xlsx",
        filetypes=[("Arquivos Excel", "*.xlsx")],
        initialfile="compilado_guias"
    )
    if not salvar:
        ocultar_progresso("Cancelado pelo usuário.")
        return

    # Salva DataFrame
    df = pd.DataFrame(resultados)
    # Reordena colunas se existirem
    cols = ["Arquivo Origem", "Numero Guia", "Vencimento", "Total a Pagar", "Processo/Protocolo", "Codigo de Barras"]
    df = df[[c for c in cols if c in df.columns]]

    try:
        df.to_excel(salvar, index=False)
        msg_sucesso = f"{len(resultados)} de {len(caminhos_pdf)} arquivos processados e salvos em:\n{salvar}"
        if arquivos_falhados:
            msg_falha = "\n\n--- LOG DE ARQUIVOS QUE FALHARAM ---\n" + "\n".join(arquivos_falhados)
            messagebox.showwarning("Sucesso Parcial", msg_sucesso + msg_falha)
        else:
            messagebox.showinfo("Sucesso", msg_sucesso)
    except Exception as e:
        messagebox.showerror("Erro ao Salvar", f"Não foi possível salvar o arquivo Excel.\nErro: {e}")

    ocultar_progresso(f"Concluído! {len(resultados)} arquivos salvos.")

# =========================
# INTERFACE TKINTER
# =========================

# Paleta genérica (sem referências a empresas)
PRIMARY_BLUE = "#2D6CDF"
PRIMARY_BLUE_HOVER = "#1E4CA8"
BG_LIGHT = "#F4F4F4"
TEXT_COLOR = "#333333"
FOOTER_COLOR = "#888888"
WHITE = "#FFFFFF"

janela = tk.Tk()
janela.title("Extrator de Guias")
janela.geometry("480x360")
janela.configure(bg=BG_LIGHT)
janela.resizable(False, False)

style = ttk.Style(janela)
# Usa tema mais neutro; dependendo do OS pode mudar aparência
try:
    style.theme_use('clam')
except Exception:
    pass

style.configure('TFrame', background=BG_LIGHT)
style.configure('TLabel', background=BG_LIGHT, foreground=TEXT_COLOR, font=('Arial', 10))
style.configure('Title.TLabel', font=('Arial', 14, 'bold'))
style.configure('Percent.TLabel', font=('Arial', 11, 'bold'), foreground=PRIMARY_BLUE)
style.configure('Footer.TLabel', font=('Arial', 8), foreground=FOOTER_COLOR)

style.configure(
    'Main.TButton',
    background=PRIMARY_BLUE,
    foreground=WHITE,
    font=('Arial', 12, 'bold'),
    borderwidth=0,
    padding=(15, 10),
    relief='flat'
)
style.map('Main.TButton', background=[('active', PRIMARY_BLUE_HOVER)])

style.configure('Custom.Horizontal.TProgressbar', troughcolor=BG_LIGHT, bordercolor=PRIMARY_BLUE, background=PRIMARY_BLUE)

frame = ttk.Frame(janela, style='TFrame', padding=(20, 10))
frame.pack(expand=True, fill='both')

lbl_titulo = ttk.Label(frame, text="Extrator de Guias", style='Title.TLabel', anchor='center')
lbl_titulo.pack(pady=(10,5), fill='x')

lbl_desc1 = ttk.Label(frame, text="Selecione múltiplos PDFs para extrair.", anchor='center')
lbl_desc1.pack(fill='x')

lbl_desc2 = ttk.Label(frame, text="Os dados de todos serão salvos em um único Excel.", anchor='center')
lbl_desc2.pack(pady=(0, 20), fill='x')

btn_extrair = ttk.Button(frame, text="Selecionar PDFs e Extrair", command=extrair_dados_massa, style='Main.TButton')
btn_extrair.pack()

# Widgets de progresso (inicialmente ocultos)
progress_label_var = tk.StringVar(value="")
lbl_progress = ttk.Label(frame, textvariable=progress_label_var, style='TLabel', anchor='center')

progress_bar = ttk.Progressbar(frame, orient='horizontal', length=420, mode='determinate', style='Custom.Horizontal.TProgressbar')

percent_label_var = tk.StringVar(value="")
lbl_percent = ttk.Label(frame, textvariable=percent_label_var, style='Percent.TLabel', anchor='center')

# Rodapé pequeno (opcional) - se preferir remova esta linha
lbl_footer = ttk.Label(frame, text="Desenvolvido por: (remova ou altere conforme desejar)", style='Footer.TLabel', anchor='center')
lbl_footer.pack(side='bottom', pady=(10,0))

# Centraliza janela
try:
    janela.eval('tk::PlaceWindow . center')
except Exception:
    pass

# Inicia app
if __name__ == "__main__":
    janela.mainloop()
