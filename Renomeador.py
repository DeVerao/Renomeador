import os
import re
import fitz  # PyMuPDF
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox
import hashlib
 
PALAVRAS_HOLERITE = ["holerite", "demonstrativo de pagamento", "contracheque", "recibo de pagamento", "folha mensal"]
PALAVRAS_COMPROVANTE = ["comprovante", "transferência", "depósito", "pagamento efetuado", "recibo de transferência"]
 
def renomear_com_seguro(caminho_pdf, novo_caminho, log_func):
    if caminho_pdf == novo_caminho:
        log_func(f"-> Arquivo '{os.path.basename(caminho_pdf)}' já está com o nome correto. Pulando.")
        return
    base, ext = os.path.splitext(novo_caminho)
    contador = 1
    while os.path.exists(novo_caminho):
        novo_caminho = f"{base} ({contador}){ext}"
        contador += 1
    os.rename(caminho_pdf, novo_caminho)
    log_func(f"✔ Renomeado: '{os.path.basename(caminho_pdf)}' → '{os.path.basename(novo_caminho)}'")
 
def identificar_tipo(texto):
    texto_lower = texto.lower()
    if any(p in texto_lower for p in PALAVRAS_HOLERITE):
        return "H"
    if any(p in texto_lower for p in PALAVRAS_COMPROVANTE):
        return "C"
    return "C"
 
def extrair_dados_holerite(texto):
    data_formatada = None
    nome = None
    meses = {
        "janeiro": "01", "fevereiro": "02", "março": "03","marco":"03", "abril": "04",
        "maio": "05", "junho": "06", "julho": "07", "agosto": "08",
        "setembro": "09", "outubro": "10", "novembro": "11", "dezembro": "12"
    }

    match_data = re.search(r"(?:Ref\.|Folha Mensal|Horista|Mensalista)[\s|.:]\s*(\w+)\s*de\s*(\d{4})", texto,
                           re.IGNORECASE)
    if match_data:
        mes_extenso, ano = match_data.groups()
        mes_numerico = meses.get(mes_extenso.lower())
        if mes_numerico:
            data_formatada = f"{ano[-2:]}.{mes_numerico}"
    padroes_nome = [
        r"^\s*([A-ZÂÁÉÍÓÚÃÕÇ ]{5,})\s*\n.*nome do funcionário",
        r"nome do funcionário.*\n\s*([A-ZÂÁÉÍÓÚÃÕÇ ]{5,})\s*$",
        r"^\d{4}[\s-]([A-ZÂÁÉÍÓÚÃÕÇ ]+)",
        r"^\d{4}\s*\n([A-ZÂÁÉÍÓÚÃÕÇ ]+)",
    ]
    for padrao in padroes_nome:
        match_nome = re.search(padrao, texto, re.MULTILINE | re.IGNORECASE)
        if match_nome:
            nome = match_nome.group(1).strip().upper()
            break
    return data_formatada, nome, None
 
def extrair_dados_comprovante(texto):
    data_formatada = None
    nome = None
    identificador = None
    bloco_favorecido_match = re.search(
        r"(?:TRANSFERIDO PARA:|quem recebeu:?|dados do recebedor)([\s\S]+?)(?=NR\. DOCUMENTO|Conta|NR\. AUTENTICACAO|Transação efetuada)",
        texto, re.IGNORECASE)
    texto_alvo = texto
    if bloco_favorecido_match:
        texto_alvo = bloco_favorecido_match.group(1)
    padroes_nome = [
        r"nome do recebedor:\s*([A-Za-zÂÁÉÍÓÚÃÕÇ ]+)",
        r"BENEFICI[AÁ]RIO(?: final)?:?\s+([A-ZÂÁÉÍÓÚÃÕÇ ]+)",
        r"(?:CLIENTE|Nome):?\s*([A-Za-zÂÁÉÍÓÚÃÕÇ ]+)",
        r"favorecido:?\s*([A-Za-zÂÁÉÍÓÚÃÕÇ ]+)"
    ]
    for padrao in padroes_nome:
        match_nome = re.search(padrao, texto_alvo, re.IGNORECASE)
        if match_nome:
            nome = match_nome.group(1).strip().upper()
            break
    match_desc = re.search(r"Descrição:\s*(.*)", texto, re.IGNORECASE)
    if match_desc:
        desc_texto = match_desc.group(1).strip()
        if desc_texto:
            match_data_desc = re.search(r"(\d{2})[\s-](\d{2})$", desc_texto)
            if match_data_desc:
                mes, ano_curto = match_data_desc.groups()
                data_formatada = f"{ano_curto}.{mes}"
            identificador = desc_texto.split()[0].upper()
    if not data_formatada:
        padrao_data_fallback = r"(?:Data de pagamento|data da transferência|DATA DA TRANSFERENCIA|DATA DO PAGAMENTO|Data da operação|Transferido em|efetuada em)?:?\s*(\d{2})[./](\d{2})[./](\d{4})"
        match_data = re.search(padrao_data_fallback, texto)
        if match_data:
            dia, mes, ano = match_data.groups()
            if int(dia) <= 26:
                data_pagamento = datetime(int(ano), int(mes), int(dia))
                data_referencia = data_pagamento - timedelta(days=27)
                data_formatada = data_referencia.strftime("%y.%m")
            else:
                data_formatada = f"{ano[-2:]}.{mes}"
    return data_formatada, nome, identificador
 
def calcular_hash_arquivo(caminho_arquivo):
    sha256_hash = hashlib.sha256()
    try:
        with open(caminho_arquivo, "rb") as f:
            for byte_block in iter(lambda: f.read(4096), b""):
                sha256_hash.update(byte_block)
        return sha256_hash.hexdigest()
    except IOError:
        return None
 
def excluir_duplicados(pasta, log_func):
    hashes = {}
    arquivos_para_verificar = [f for f in os.listdir(pasta) if os.path.isfile(os.path.join(pasta, f))]
    log_func("--- Iniciando verificação de duplicados ---")
    if not arquivos_para_verificar:
        log_func("Nenhum arquivo encontrado para verificar.")
        return
    arquivos_excluidos = 0
    for arquivo in arquivos_para_verificar:
        caminho_completo = os.path.join(pasta, arquivo)
        hash_arquivo = calcular_hash_arquivo(caminho_completo)
        if hash_arquivo is None:
            log_func(f"⚠ Não foi possível ler o arquivo: '{arquivo}'. Pulando.")
            continue
        if hash_arquivo in hashes:
            original = os.path.basename(hashes[hash_arquivo])
            log_func(f"🗑️ Duplicado: '{arquivo}' é idêntico a '{original}'. Excluindo...")
            try:
                os.remove(caminho_completo)
                log_func(f"✔ Arquivo '{arquivo}' excluído.")
                arquivos_excluidos += 1
            except OSError as e:
                log_func(f"❌ Erro ao excluir '{arquivo}': {e}")
        else:
            hashes[hash_arquivo] = caminho_completo
    log_func(f"--- Verificação concluída. Total de {arquivos_excluidos} arquivos duplicados excluídos. ---")
 
class PdfRenamerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Organizador de PDFs")
        self.root.geometry("800x600")
        self.root.minsize(600, 400)
        self.folder_path = tk.StringVar()
        main_frame = tk.Frame(root, padx=10, pady=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        folder_frame = tk.Frame(main_frame)
        folder_frame.pack(fill=tk.X, pady=(0, 10))
        select_button = tk.Button(folder_frame, text="Selecionar Pasta", command=self.select_folder)
        select_button.pack(side=tk.LEFT, padx=(0, 10))
        folder_label = tk.Label(folder_frame, textvariable=self.folder_path, relief=tk.SUNKEN, bg="white", anchor="w")
        folder_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.folder_path.set("Nenhuma pasta selecionada...")
        log_frame = tk.Frame(main_frame)
        log_frame.pack(fill=tk.BOTH, expand=True)
        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, state='disabled', font=("Arial", 10))
        self.log_text.pack(fill=tk.BOTH, expand=True)
        action_frame = tk.Frame(main_frame)
        action_frame.pack(fill=tk.X, pady=(10, 0))
        self.start_button = tk.Button(action_frame, text="Renomear por Conteúdo", command=self.start_processing,
                                      bg="#4CAF50", fg="white", font=("Arial", 12, "bold"))
        self.start_button.pack(fill=tk.X)
        self.delete_button = tk.Button(action_frame, text="Excluir Duplicados", command=self.start_deleting_duplicates,
                                       bg="#F44336", fg="white", font=("Arial", 12, "bold"))
        self.delete_button.pack(fill=tk.X, pady=(5, 0))
 
    def select_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.folder_path.set(path)
            self.log_message(f"Pasta selecionada: {path}\n")
 
    def log_message(self, message):
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state='disabled')
        self.root.update_idletasks()
 
    def start_processing(self):
        path = self.folder_path.get()
        if not os.path.isdir(path):
            messagebox.showerror("Erro", "Por favor, selecione uma pasta válida antes de iniciar.")
            return
        self.log_text.config(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state='disabled')

        self.log_message(f"Iniciando processo de renomeação na pasta: {path}\n" + "=" * 50)
        self.start_button.config(state='disabled', text="Processando...")
        self.delete_button.config(state='disabled')
        try:
            self.processar_pdfs_gui(path)
        except Exception as e:
            self.log_message(f"\nERRO INESPERADO: {e}")
            messagebox.showerror("Erro Crítico", f"Ocorreu um erro inesperado durante o processamento:\n{e}")
        finally:
            self.start_button.config(state='normal', text="Renomear por Conteúdo")
            self.delete_button.config(state='normal')
            self.log_message("\n" + "=" * 50 + "\nProcesso de renomeação concluído!")
 
    def start_deleting_duplicates(self):
        path = self.folder_path.get()
        if not os.path.isdir(path):
            messagebox.showerror("Erro", "Por favor, selecione uma pasta válida antes de iniciar.")
            return
        if not messagebox.askyesno("Confirmação",
                                   "Tem certeza que deseja excluir PERMANENTEMENTE os arquivos duplicados nesta pasta?\n\nEsta ação não pode ser desfeita."):
            self.log_message("Operação de exclusão cancelada pelo usuário.")
            return
        self.log_message("\n" + "=" * 50)
        self.start_button.config(state='disabled')
        self.delete_button.config(state='disabled', text="Excluindo...")
        try:
            excluir_duplicados(path, self.log_message)
        except Exception as e:
            self.log_message(f"\nERRO INESPERADO ao excluir duplicados: {e}")
            messagebox.showerror("Erro Crítico", f"Ocorreu um erro inesperado durante a exclusão:\n{e}")
        finally:
            self.start_button.config(state='normal')
            self.delete_button.config(state='normal', text="Excluir Duplicados")
            self.log_message("\nProcesso de exclusão de duplicados concluído!")
 
    def processar_pdfs_gui(self, pasta):
        arquivos_pdf = [f for f in os.listdir(pasta) if f.lower().endswith(".pdf")]
        if not arquivos_pdf:
            self.log_message("Nenhum arquivo PDF encontrado na pasta.")
            return
        for i, arquivo in enumerate(arquivos_pdf):
            self.log_message(f"\n--- Processando '{arquivo}' ({i + 1}/{len(arquivos_pdf)}) ---")
            caminho_pdf = os.path.join(pasta, arquivo)
            try:
                with fitz.open(caminho_pdf) as doc:
                    texto_completo = ""
                    for page in doc:
                        texto_pagina = page.get_text()
                        if texto_pagina:
                            texto_completo += texto_pagina + "\n"
                if not texto_completo.strip():
                    self.log_message(f"⚠ Aviso: Arquivo '{arquivo}' está vazio ou contém apenas imagens.")
                    continue
                tipo = identificar_tipo(texto_completo)
                if tipo == 'H':
                    data, nome, identificador = extrair_dados_holerite(texto_completo)
                else:
                    data, nome, identificador = extrair_dados_comprovante(texto_completo)
                if data and nome:
                    final_id = identificador if identificador else tipo
                    novo_nome = f"{data} - {nome} - {final_id}.pdf"
                    novo_caminho = os.path.join(pasta, novo_nome)
                    renomear_com_seguro(caminho_pdf, novo_caminho, self.log_message)
                else:
                    erros = []
                    if not data: erros.append("data não encontrada")
                    if not nome: erros.append("nome não encontrado")
                    self.log_message(f"❌ Falha ao extrair dados de '{arquivo}': {', '.join(erros)}.")
            except Exception as e:
                self.log_message(f"❌ Erro Crítico ao processar '{arquivo}': {e}")
 
if __name__ == "__main__":
    root = tk.Tk()
    app = PdfRenamerApp(root)
    root.mainloop()
