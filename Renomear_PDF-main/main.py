import os
import re
import shutil
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pdfplumber


def limpar_nome_arquivo(nome: str) -> str:
    """
    Remove caracteres inválidos para nome de arquivo no Windows/Linux.
    """
    nome = re.sub(r'[\\/*?:"<>|]', "", nome)
    nome = re.sub(r"\s+", " ", nome).strip()
    return nome


def extrair_cpf_e_nome(pdf_path: str):
    """
    Extrai CPF e Nome Completo do PDF.
    Retorna (cpf, nome) ou (None, None) se não encontrar.
    """
    texto_completo = []

    try:
        with pdfplumber.open(pdf_path) as pdf:
            for pagina in pdf.pages:
                texto = pagina.extract_text()
                if texto:
                    texto_completo.append(texto)
    except Exception as e:
        raise RuntimeError(f"Erro ao abrir o PDF: {e}")

    texto = "\n".join(texto_completo)

    # Procura a linha onde aparecem CPF + Nome após o cabeçalho
    match = re.search(
        r"CPF\s+Nome Completo\s*\n\s*(\d{3}\.\d{3}\.\d{3}-\d{2})\s+([A-ZÁÀÂÃÉÈÊÍÌÎÓÒÔÕÚÙÛÇ\s]+)",
        texto,
        re.IGNORECASE
    )

    if match:
        cpf = match.group(1).strip()
        nome = match.group(2).strip()

        # corta possíveis trechos extras capturados depois do nome
        nome = re.split(
            r"\n|Natureza do Rendimento|3\. Rendimentos|Rendimentos Tributáveis",
            nome,
            maxsplit=1,
            flags=re.IGNORECASE
        )[0].strip()

        nome = re.sub(r"\s+", " ", nome)
        return cpf, nome

    return None, None


def montar_nome_arquivo(cpf: str, nome: str) -> str:
    nome_final = f"{cpf} - {nome}.pdf"
    return limpar_nome_arquivo(nome_final)


class AppRenomeadorPDF:
    def __init__(self, root):
        self.root = root
        self.root.title("Renomeador de PDFs por CPF e Nome")
        self.root.geometry("760x500")
        self.root.minsize(700, 450)

        self.input_folder = tk.StringVar()
        self.output_folder = tk.StringVar()
        self.prefixo = tk.StringVar(value="")

        self.criar_interface()

    def criar_interface(self):
        frame = ttk.Frame(self.root, padding=15)
        frame.pack(fill="both", expand=True)

        titulo = ttk.Label(
            frame,
            text="Renomear PDFs com CPF e Nome",
            font=("Segoe UI", 14, "bold")
        )
        titulo.pack(anchor="w", pady=(0, 10))

        subtitulo = ttk.Label(
            frame,
            text="Selecione a pasta de origem e a pasta de destino. "
                 "O sistema irá ler cada PDF, extrair CPF e Nome Completo e salvar com o novo nome.",
            wraplength=700,
            justify="left"
        )
        subtitulo.pack(anchor="w", pady=(0, 15))

        # Origem
        origem_frame = ttk.LabelFrame(frame, text="Pasta de origem", padding=10)
        origem_frame.pack(fill="x", pady=5)

        ttk.Entry(origem_frame, textvariable=self.input_folder).pack(
            side="left", fill="x", expand=True, padx=(0, 10)
        )
        ttk.Button(origem_frame, text="Selecionar", command=self.selecionar_origem).pack(side="left")

        # Destino
        destino_frame = ttk.LabelFrame(frame, text="Pasta de destino", padding=10)
        destino_frame.pack(fill="x", pady=5)

        ttk.Entry(destino_frame, textvariable=self.output_folder).pack(
            side="left", fill="x", expand=True, padx=(0, 10)
        )
        ttk.Button(destino_frame, text="Selecionar", command=self.selecionar_destino).pack(side="left")

        # Prefixo opcional
        prefixo_frame = ttk.LabelFrame(frame, text="Opções", padding=10)
        prefixo_frame.pack(fill="x", pady=5)

        ttk.Label(prefixo_frame, text="Prefixo opcional no nome do arquivo:").pack(side="left", padx=(0, 10))
        ttk.Entry(prefixo_frame, textvariable=self.prefixo, width=30).pack(side="left")

        # Botões
        botoes = ttk.Frame(frame)
        botoes.pack(fill="x", pady=15)

        self.btn_processar = ttk.Button(
            botoes,
            text="Processar PDFs",
            command=self.iniciar_processamento
        )
        self.btn_processar.pack(side="left")

        ttk.Button(
            botoes,
            text="Limpar log",
            command=self.limpar_log
        ).pack(side="left", padx=10)

        # Barra de progresso
        self.progress = ttk.Progressbar(frame, mode="determinate")
        self.progress.pack(fill="x", pady=(0, 10))

        # Log
        log_frame = ttk.LabelFrame(frame, text="Log de processamento", padding=10)
        log_frame.pack(fill="both", expand=True)

        self.log_text = tk.Text(log_frame, height=15, wrap="word")
        self.log_text.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        scrollbar.pack(side="right", fill="y")
        self.log_text.configure(yscrollcommand=scrollbar.set)

    def selecionar_origem(self):
        pasta = filedialog.askdirectory(title="Selecione a pasta com os PDFs")
        if pasta:
            self.input_folder.set(pasta)

    def selecionar_destino(self):
        pasta = filedialog.askdirectory(title="Selecione a pasta de destino")
        if pasta:
            self.output_folder.set(pasta)

    def log(self, mensagem: str):
        self.log_text.insert("end", mensagem + "\n")
        self.log_text.see("end")
        self.root.update_idletasks()

    def limpar_log(self):
        self.log_text.delete("1.0", "end")

    def iniciar_processamento(self):
        origem = self.input_folder.get().strip()
        destino = self.output_folder.get().strip()

        if not origem:
            messagebox.showwarning("Atenção", "Selecione a pasta de origem.")
            return

        if not destino:
            messagebox.showwarning("Atenção", "Selecione a pasta de destino.")
            return

        if not os.path.isdir(origem):
            messagebox.showerror("Erro", "A pasta de origem é inválida.")
            return

        if not os.path.isdir(destino):
            messagebox.showerror("Erro", "A pasta de destino é inválida.")
            return

        self.btn_processar.config(state="disabled")
        self.progress["value"] = 0
        self.limpar_log()

        thread = threading.Thread(target=self.processar_pdfs, daemon=True)
        thread.start()

    def processar_pdfs(self):
        origem = self.input_folder.get().strip()
        destino = self.output_folder.get().strip()
        prefixo = self.prefixo.get().strip()

        arquivos_pdf = [f for f in os.listdir(origem) if f.lower().endswith(".pdf")]

        if not arquivos_pdf:
            self.log("Nenhum PDF encontrado na pasta de origem.")
            self.btn_processar.config(state="normal")
            return

        total = len(arquivos_pdf)
        self.progress["maximum"] = total

        processados = 0
        erros = 0

        for i, arquivo in enumerate(arquivos_pdf, start=1):
            caminho_pdf = os.path.join(origem, arquivo)

            try:
                cpf, nome = extrair_cpf_e_nome(caminho_pdf)

                if not cpf or not nome:
                    self.log(f"[ERRO] Não foi possível extrair CPF/Nome: {arquivo}")
                    erros += 1
                    self.progress["value"] = i
                    continue

                novo_nome = montar_nome_arquivo(cpf, nome)

                if prefixo:
                    novo_nome = limpar_nome_arquivo(f"{prefixo} {novo_nome}")

                destino_final = os.path.join(destino, novo_nome)

                # Evita sobrescrever caso já exista
                contador = 1
                base_nome, ext = os.path.splitext(destino_final)
                while os.path.exists(destino_final):
                    destino_final = f"{base_nome} ({contador}){ext}"
                    contador += 1

                shutil.copy2(caminho_pdf, destino_final)
                self.log(f"[OK] {arquivo}  ->  {os.path.basename(destino_final)}")
                processados += 1

            except Exception as e:
                self.log(f"[ERRO] {arquivo}: {e}")
                erros += 1

            self.progress["value"] = i

        self.btn_processar.config(state="normal")
        messagebox.showinfo(
            "Concluído",
            f"Processamento finalizado.\n\n"
            f"Sucesso: {processados}\n"
            f"Erros: {erros}\n"
            f"Total: {total}"
        )


if __name__ == "__main__":
    root = tk.Tk()
    app = AppRenomeadorPDF(root)
    root.mainloop()