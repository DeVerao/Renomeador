import os
import re
import pdfplumber
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox

PALAVRAS_HOLERITE = ["holerite", "demonstrativo de pagamento", "contracheque", "recibo de pagamento", "folha mensal"]
PALAVRAS_COMPROVANTE = ["comprovante", "transferência", "depósito", "pagamento efetuado", "recibo de transferência"]


def renomear_com_seguro(caminho_pdf, novo_caminho, log_func):
    """
    Renomeia um arquivo de forma segura, adicionando sufixo (1), (2), etc. se o nome já existir.
    Recebe a função de log como terceiro argumento para reportar o status na GUI.
    """
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

def identificar_tipo(texto):
    """
    Identifica se é comprovante (C) ou holerite (H).
    A verificação de holerite ocorre primeiro, pois termos como "recibo de pagamento"
    são mais específicos para esse contexto do que para um comprovante genérico.
    """
    texto_lower = texto.lower()
    if any(p in texto_lower for p in PALAVRAS_HOLERITE):
        return "H"
    if any(p in texto_lower for p in PALAVRAS_COMPROVANTE):
        return "C"

    # Se nenhuma palavra-chave for encontrada, assume-se "C" como padrão.
    return "C"


def extrair_dados_holerite(texto):
    """
    Extrai dados especificamente de arquivos de holerite/recibo de pagamento.
    """
    data_formatada = None
    nome = None

    # Dicionário para converter mês por extenso (Português) para número.
    meses = {
        "janeiro": "01", "fevereiro": "02", "março": "03", "abril": "04",
        "maio": "05", "junho": "06", "julho": "07", "agosto": "08",
        "setembro": "09", "outubro": "10", "novembro": "11", "dezembro": "12"
    }

    # 1. Extrair a data de referência (ex: "Ref.: Novembro de 2024")
    match_data = re.search(r"(?:Ref\.|Folha Mensal|Horista|Mensalista)[\s|.:]\s*(\w+)\s*de\s*(\d{4})", texto, re.IGNORECASE)
    if match_data:
        mes_extenso, ano = match_data.groups()
        mes_numerico = meses.get(mes_extenso.lower())
        if mes_numerico:
            data_formatada = f"{ano[-2:]}.{mes_numerico}"

    # 2. Extrair o nome do funcionário.
    #    MELHORIA: Adicionado padrões mais específicos e robustos.
    padroes_nome = [
        r"^\d{4}[\s-]([A-ZÂÁÉÍÓÚÃÕÇ ]+)",
    ]
    for padrao in padroes_nome:
        # Usamos re.MULTILINE para que a busca considere as quebras de linha
        match_nome = re.search(padrao, texto, re.MULTILINE)
        if match_nome:
            nome = match_nome.group(1).strip().upper()
            break

    return data_formatada, nome, None


def extrair_dados_comprovante(texto):
    """
    Extrai dados de comprovantes de transferência.
    """
    data_formatada = None
    nome = None
    identificador = None

    # MELHORIA: Isolar a busca do nome do favorecido para a seção correta.
    # Primeiro, tentamos encontrar o bloco de texto "TRANSFERIDO PARA".
    bloco_favorecido_match = re.search(
        r"TRANSFERIDO PARA:([\s\S]+?)(?=NR\. DOCUMENTO|NR\. AUTENTICACAO|Transação efetuada)", texto, re.IGNORECASE)

    texto_alvo = texto  # Por padrão, busca no texto inteiro
    if bloco_favorecido_match:
        # Se encontrou o bloco, a busca pelo nome será feita apenas dentro dele.
        texto_alvo = bloco_favorecido_match.group(1)

    # Padrões para extrair o nome do favorecido/beneficiário.
    padroes_nome = [
        r"(?:CLIENTE|Nome):\s*([A-Za-zÂÁÉÍÓÚÃÕÇ ]+)",
        r"nome do recebedor:\s*([A-Za-zÂÁÉÍÓÚÃÕÇ ]+)",
        r"BENEFICIARIO:\s*([A-Za-zÂÁÉÍÓÚÃÕÇ ]+)",
    ]

    for padrao in padroes_nome:
        match_nome = re.search(padrao, texto_alvo, re.IGNORECASE)
        if match_nome:
            nome = match_nome.group(1).strip().upper()
            break

    # Tenta extrair data e identificador do campo "Descrição"
    match_desc = re.search(r"Descrição:\s*(.*)", texto, re.IGNORECASE)
    if match_desc:
        desc_texto = match_desc.group(1).strip()
        if desc_texto:
            match_data_desc = re.search(r"(\d{2})[\s-](\d{2})$", desc_texto)
            if match_data_desc:
                mes, ano_curto = match_data_desc.groups()
                data_formatada = f"{ano_curto}.{mes}"
            identificador = desc_texto.split()[0].upper()

    # Se a data não foi encontrada na descrição, usa um padrão de fallback
    if not data_formatada:
        padrao_data_fallback = r"(?:Data de pagamento|data da transferência|DATA DA TRANSFERENCIA|DATA DO PAGAMENTO|Data da operação|Transferido em|efetuada em):?\s*(\d{2})[./](\d{2})[./](\d{4})"
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


class PdfRenamerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Renomeador de PDFs")
        self.root.geometry("800x600")
        self.root.minsize(600, 400)

        self.folder_path = tk.StringVar()

        # --- Frame principal ---
        main_frame = tk.Frame(root, padx=10, pady=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # --- Frame para seleção de pasta ---
        folder_frame = tk.Frame(main_frame)
        folder_frame.pack(fill=tk.X, pady=(0, 10))

        select_button = tk.Button(folder_frame, text="Selecionar Pasta", command=self.select_folder)
        select_button.pack(side=tk.LEFT, padx=(0, 10))

        folder_label = tk.Label(folder_frame, textvariable=self.folder_path, relief=tk.SUNKEN, bg="white", anchor="w")
        folder_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.folder_path.set("Nenhuma pasta selecionada...")

        # --- Frame para o log de atividades ---
        log_frame = tk.Frame(main_frame)
        log_frame.pack(fill=tk.BOTH, expand=True)

        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, state='disabled', font=("Arial", 10))
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # --- Frame para o botão de iniciar ---
        action_frame = tk.Frame(main_frame)
        action_frame.pack(fill=tk.X, pady=(10, 0))

        self.start_button = tk.Button(action_frame, text="Iniciar Processamento", command=self.start_processing,
                                      bg="#4CAF50", fg="white", font=("Arial", 12, "bold"))
        self.start_button.pack(fill=tk.X)

    def select_folder(self):
        """Abre a caixa de diálogo para selecionar uma pasta."""
        path = filedialog.askdirectory()
        if path:
            self.folder_path.set(path)
            self.log_message(f"Pasta selecionada: {path}\n")

    def log_message(self, message):
        """Adiciona uma mensagem à caixa de log."""
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)  # Auto-scroll
        self.log_text.config(state='disabled')
        self.root.update_idletasks()

    def start_processing(self):
        """Inicia o processo de renomeação dos PDFs na pasta selecionada."""
        path = self.folder_path.get()
        if not os.path.isdir(path):
            messagebox.showerror("Erro", "Por favor, selecione uma pasta válida antes de iniciar.")
            return

        self.log_text.config(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state='disabled')

        self.log_message(f"Iniciando processamento na pasta: {path}\n" + "=" * 50)
        self.start_button.config(state='disabled', text="Processando...")

        try:
            self.processar_pdfs_gui(path)
        except Exception as e:
            self.log_message(f"\nERRO INESPERADO: {e}")
            messagebox.showerror("Erro Crítico", f"Ocorreu um erro inesperado durante o processamento:\n{e}")
        finally:
            self.start_button.config(state='normal', text="Iniciar Processamento")
            self.log_message("\n" + "=" * 50 + "\nProcessamento concluído!")

    def processar_pdfs_gui(self, pasta):
        """Versão da função principal adaptada para a GUI."""
        arquivos_pdf = [f for f in os.listdir(pasta) if f.lower().endswith(".pdf")]
        if not arquivos_pdf:
            self.log_message("Nenhum arquivo PDF encontrado na pasta.")
            return

        for i, arquivo in enumerate(arquivos_pdf):
            self.log_message(f"\n--- Processando '{arquivo}' ({i + 1}/{len(arquivos_pdf)}) ---")
            caminho_pdf = os.path.join(pasta, arquivo)

            try:
                with pdfplumber.open(caminho_pdf) as pdf:
                    texto_completo = ""
                    for page in pdf.pages:
                        texto_pagina = page.extract_text(x_tolerance=2, y_tolerance=2)
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
