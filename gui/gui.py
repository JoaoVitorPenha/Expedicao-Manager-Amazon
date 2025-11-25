import tkinter as tk
from tkinter import filedialog, messagebox
import shutil
import os
import re
from PyPDF2 import PdfReader
from datetime import datetime

# Caminhos
PDFS_DIR = "pdfs/"
MINUTA_DIR = os.path.join(PDFS_DIR, "minuta/")
LISTA_PEDIDOS_DIR = os.path.join(PDFS_DIR, "lista_pedidos/")
DESKTOP = os.path.join(os.path.expanduser("~"), "Desktop")
LOG_FILE = os.path.join(PDFS_DIR, "log.txt")

# Garantir pastas
for pasta in [PDFS_DIR, MINUTA_DIR, LISTA_PEDIDOS_DIR]:
    os.makedirs(pasta, exist_ok=True)

# Limpeza diária de PDFs e TXTs
def limpar_pdfs_se_dia_virou():
    controle_data = os.path.join(PDFS_DIR, "last_run_date.txt")
    hoje = datetime.now().strftime("%Y-%m-%d")

    ultima_data = None
    if os.path.exists(controle_data):
        with open(controle_data, "r", encoding="utf-8") as f:
            ultima_data = f.read().strip()

    if ultima_data != hoje:
        for pasta in [MINUTA_DIR, LISTA_PEDIDOS_DIR]:
            for arquivo in os.listdir(pasta):
                if arquivo.lower().endswith((".pdf", ".txt")):
                    os.remove(os.path.join(pasta, arquivo))
        with open(controle_data, "w", encoding="utf-8") as f:
            f.write(hoje)

limpar_pdfs_se_dia_virou()

def registrar_log(mensagem):
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {mensagem}\n")

def selecionar_arquivos():
    arquivos = filedialog.askopenfilenames(filetypes=[("Arquivos PDF", "*.pdf")])
    if arquivos:
        for arquivo in arquivos:
            if not arquivo.lower().endswith(".pdf"):
                continue
            nome = os.path.basename(arquivo)
            timestamp = datetime.now().strftime("%H%M%S")
            novo_nome = f"{timestamp}_{nome}"
            try:
                shutil.copy(arquivo, os.path.join(MINUTA_DIR, novo_nome))
                shutil.copy(arquivo, os.path.join(LISTA_PEDIDOS_DIR, novo_nome))
            except Exception as e:
                registrar_log(f"Erro ao copiar {arquivo}: {e}")
        messagebox.showinfo("Sucesso", "Arquivos carregados e salvos nas pastas.")

# Parser
STARTWORD = "QUANTIDADE"
STOPWORDS = ["SKU:", "ASIN:", "CONDIÇÃO", "ID DO ITEM"]

def extrair_produtos(pdf_path):
    produtos = {}
    try:
        reader = PdfReader(pdf_path)
    except Exception as e:
        registrar_log(f"Erro ao abrir {pdf_path}: {e}")
        return produtos, 0

    for page in reader.pages:
        try:
            text = page.extract_text()
        except Exception as e:
            registrar_log(f"Erro ao extrair texto de {pdf_path}: {e}")
            continue

        if not text:
            continue

        linhas = text.split("\n")
        collecting = False
        quantidade = 0
        produto_nome = ""

        for linha in linhas:
            linha_clean = linha.strip()

            if STARTWORD in linha_clean.upper():
                collecting = True
                continue

            if collecting:
                if any(sw in linha_clean.upper() for sw in STOPWORDS):
                    if produto_nome and len(produto_nome.split()) > 1:
                        produtos[produto_nome] = produtos.get(produto_nome, 0) + quantidade
                    quantidade = 0
                    produto_nome = ""
                    collecting = False
                    continue

                match = re.match(r"^(\d+)\s+(.*)", linha_clean)
                if match:
                    quantidade = int(match.group(1))
                    produto_nome = match.group(2).strip()
                else:
                    if produto_nome:
                        produto_nome += " " + linha_clean

        if produto_nome and len(produto_nome.split()) > 1:
            produtos[produto_nome] = produtos.get(produto_nome, 0) + quantidade

    return produtos, len(reader.pages)

def gerar_lista_produtos():
    produtos = {}
    total_pedidos = 0

    pdfs = [os.path.join(LISTA_PEDIDOS_DIR, f) for f in os.listdir(LISTA_PEDIDOS_DIR) if f.lower().endswith(".pdf")]
    if not pdfs:
        messagebox.showwarning("Aviso", "Nenhum PDF encontrado em lista_pedidos/.")
        return

    for pdf in pdfs:
        produtos_pdf, pedidos_pdf = extrair_produtos(pdf)
        total_pedidos += pedidos_pdf
        for nome, qtd in produtos_pdf.items():
            produtos[nome] = produtos.get(nome, 0) + qtd

    produtos_ordenados = dict(sorted(produtos.items(), key=lambda x: x[0].lower()))

    data_atual = datetime.now().strftime("_%d%m%y")
    nome_arquivo = f"Lista_de_Produtos{data_atual}.txt"

    linhas_saida = [f"{qtd} - {nome}" for nome, qtd in produtos_ordenados.items()]

    try:
        with open(os.path.join(LISTA_PEDIDOS_DIR, nome_arquivo), "w", encoding="utf-8") as f:
            f.write("\n".join(linhas_saida))
        with open(os.path.join(DESKTOP, nome_arquivo), "w", encoding="utf-8") as f:
            f.write("\n".join(linhas_saida))
    except Exception as e:
        registrar_log(f"Erro ao salvar lista: {e}")
        messagebox.showerror("Erro", f"Não foi possível salvar a lista: {e}")
        return

    registrar_log(f"Lista gerada: {len(produtos)} produtos, {total_pedidos} pedidos")
    messagebox.showinfo("Lista de Produtos", f"Lista consolidada gerada com {len(produtos)} produtos.\n"
                                             f"Total de pedidos: {total_pedidos}\n"
                                             f"Arquivo salvo no Desktop.")

def gerar_minuta():
    pdfs = [os.path.join(MINUTA_DIR, f) for f in os.listdir(MINUTA_DIR) if f.lower().endswith(".pdf")]
    data_atual = datetime.now().strftime("_%d%m%y")
    nome_arquivo = f"Minuta{data_atual}.txt"

    linhas = []
    for pdf in pdfs:
        try:
            reader = PdfReader(pdf)
            for page in reader.pages:
                text = page.extract_text()
                if text:
                    linhas.append(text)
        except Exception as e:
            registrar_log(f"Erro ao gerar minuta de {pdf}: {e}")

    try:
        with open(os.path.join(MINUTA_DIR, nome_arquivo), "w", encoding="utf-8") as f:
            f.write("\n".join(linhas))
        with open(os.path.join(DESKTOP, nome_arquivo), "w", encoding="utf-8") as f:
            f.write("\n".join(linhas))
    except Exception as e:
        registrar_log(f"Erro ao salvar minuta: {e}")
        messagebox.showerror("Erro", f"Não foi possível salvar a minuta: {e}")
        return

    registrar_log(f"Minuta gerada com {len(linhas)} páginas de texto")
    messagebox.showinfo("Minuta", "Minuta gerada e salva no Desktop.")

# GUI
root = tk.Tk()
root.configure(bg="darkblue")
root.geometry("350x150")
root.iconbitmap("./icon/expedicao.ico")
root.title("Expedição Manager Amazon")

btn_selecionar = tk.Button(root, bg="lightblue", text="Selecionar Arquivos PDF", font=("Arial", 12, "bold"), command=selecionar_arquivos)
btn_selecionar.pack(pady=10)

btn_lista = tk.Button(root, bg="lightblue", text="Gerar Lista de Produtos", font=("Arial", 12, "bold"),  command=gerar_lista_produtos)
btn_lista.pack(pady=10)

btn_minuta = tk.Button(root, bg="lightblue", text="Gerar Minuta", font=("Arial", 12, "bold"),  command=gerar_minuta)
btn_minuta.pack(pady=10)

root.mainloop()
