
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from openpyxl import load_workbook
from datetime import datetime, date
import os


class ExcelProcv:

    def __init__(self, root):
        self.root = root
        self.root.title("Excel PROCV Automático")
        self.root.geometry("760x750")
        self.root.minsize(680, 680)

        self.arquivo = None
        self.workbook = None

        self.cor_fundo = "#F3F5F7"
        self.cor_azul = "#2457A6"

        self.root.configure(bg=self.cor_fundo)

        self.criar_interface()

    # ==========================================
    # INTERFACE
    # ==========================================

    def criar_interface(self):

        titulo = tk.Label(
            self.root,
            text="EXCEL PROCV AUTOMÁTICO",
            font=("Arial", 20, "bold"),
            bg=self.cor_fundo,
            fg="#1F2937"
        )
        titulo.pack(pady=(22, 5))

        subtitulo = tk.Label(
            self.root,
            text="Automatize suas buscas e preencha sua planilha.",
            font=("Arial", 10),
            bg=self.cor_fundo,
            fg="#64748B"
        )
        subtitulo.pack(pady=(0, 20))

        # ARQUIVO
        self.criar_secao("1. SELECIONE SUA PLANILHA")

        frame_arquivo = tk.Frame(
            self.root, bg=self.cor_fundo
        )
        frame_arquivo.pack(fill="x", padx=30, pady=5)

        self.label_arquivo = tk.Label(
            frame_arquivo,
            text="Nenhuma planilha selecionada",
            bg="white",
            fg="#64748B",
            anchor="w",
            relief="solid",
            borderwidth=1,
            padx=10,
            pady=9
        )
        self.label_arquivo.pack(
            side="left", fill="x", expand=True
        )

        botao_arquivo = tk.Button(
            frame_arquivo,
            text="Procurar Excel",
            command=self.selecionar_arquivo,
            bg=self.cor_azul,
            fg="white",
            font=("Arial", 9, "bold"),
            relief="flat",
            padx=15,
            pady=9,
            cursor="hand2"
        )
        botao_arquivo.pack(side="right", padx=(8, 0))

        # ABA PRINCIPAL
        self.criar_secao("2. CONFIGURE A ABA PRINCIPAL")

        self.combo_principal = self.criar_combo(
            "Aba que receberá o resultado:"
        )

        self.combo_coluna_principal = self.criar_combo(
            "Coluna que contém o código de pesquisa:"
        )

        # ABA DE REFERÊNCIA
        self.criar_secao("3. CONFIGURE A BASE DE REFERÊNCIA")

        self.combo_referencia = self.criar_combo(
            "Aba onde está a tabela de referência:"
        )

        self.combo_coluna_busca = self.criar_combo(
            "Coluna onde o código será procurado:"
        )

        self.combo_coluna_retorno = self.criar_combo(
            "Coluna que contém o dado desejado:"
        )

        # RESULTADO
        self.criar_secao("4. CONFIGURE O RESULTADO")

        frame_resultado = tk.Frame(
            self.root, bg=self.cor_fundo
        )
        frame_resultado.pack(
            fill="x", padx=30, pady=(5, 10)
        )

        tk.Label(
            frame_resultado,
            text="Nome da coluna que receberá o resultado:",
            bg=self.cor_fundo,
            fg="#334155",
            font=("Arial", 9, "bold")
        ).pack(anchor="w", pady=(0, 5))

        self.entrada_resultado = tk.Entry(
            frame_resultado,
            font=("Arial", 10),
            relief="solid",
            borderwidth=1
        )
        self.entrada_resultado.pack(
            fill="x", ipady=7
        )

        # BOTÃO EXECUTAR
        self.botao_executar = tk.Button(
            self.root,
            text="EXECUTAR PROCV E EXPORTAR EXCEL",
            command=self.executar_procv,
            bg="#16834A",
            fg="white",
            font=("Arial", 11, "bold"),
            relief="flat",
            padx=15,
            pady=14,
            cursor="hand2"
        )
        self.botao_executar.pack(
            fill="x", padx=30, pady=(12, 12)
        )

        # STATUS
        self.status = tk.Label(
            self.root,
            text="Aguardando seleção da planilha...",
            bg=self.cor_fundo,
            fg="#64748B",
            font=("Arial", 9),
            wraplength=680
        )
        self.status.pack(pady=(0, 15))

        # EVENTOS DOS COMBOS
        self.combo_principal.bind(
            "<<ComboboxSelected>>",
            self.atualizar_colunas_principal
        )

        self.combo_referencia.bind(
            "<<ComboboxSelected>>",
            self.atualizar_colunas_referencia
        )

    # ==========================================
    # ELEMENTOS DA INTERFACE
    # ==========================================

    def criar_secao(self, texto):

        tk.Label(
            self.root,
            text=texto,
            bg=self.cor_fundo,
            fg=self.cor_azul,
            font=("Arial", 10, "bold")
        ).pack(anchor="w", padx=30, pady=(12, 4))

    def criar_combo(self, texto):

        frame = tk.Frame(self.root, bg=self.cor_fundo)
        frame.pack(fill="x", padx=30, pady=4)

        tk.Label(
            frame,
            text=texto,
            bg=self.cor_fundo,
            fg="#334155",
            font=("Arial", 9)
        ).pack(anchor="w", pady=(0, 4))

        combo = ttk.Combobox(
            frame,
            state="readonly",
            font=("Arial", 10)
        )

        combo.pack(fill="x", ipady=4)

        return combo

    # ==========================================
    # SELECIONAR ARQUIVO
    # ==========================================

    def selecionar_arquivo(self):

        caminho = filedialog.askopenfilename(
            title="Selecione sua planilha Excel",
            filetypes=[
                ("Planilhas Excel", "*.xlsx")
            ]
        )

        if not caminho:
            return

        try:

            wb = load_workbook(
                caminho,
                data_only=False
            )

            self.arquivo = caminho
            self.workbook = wb

            self.label_arquivo.config(
                text=os.path.basename(caminho),
                fg="#166534"
            )

            abas = wb.sheetnames

            self.combo_principal["values"] = abas
            self.combo_referencia["values"] = abas

            self.combo_principal.current(0)
            self.combo_referencia.current(0)

            self.atualizar_colunas_principal()
            self.atualizar_colunas_referencia()

            self.status.config(
                text="Planilha carregada! Configure as colunas."
                ,
                fg="#16834A"
            )

        except Exception as erro:

            messagebox.showerror(
                "Erro ao abrir Excel",
                str(erro)
            )

    # ==========================================
    # IDENTIFICAR COLUNAS DO EXCEL
    # ==========================================

    def obter_colunas(self, nome_aba):

        if not self.workbook or not nome_aba:
            return []

        ws = self.workbook[nome_aba]

        colunas = []

        for numero in range(1, ws.max_column + 1):

            letra = ws.cell(
                row=1,
                column=numero
            ).column_letter

            cabecalho = ws.cell(
                row=1,
                column=numero
            ).value

            if cabecalho is None:
                cabecalho = "(Sem título)"

            descricao = f"{letra} - {cabecalho}"

            colunas.append(descricao)

        return colunas

    def atualizar_colunas_principal(self, event=None):

        aba = self.combo_principal.get()

        colunas = self.obter_colunas(aba)

        self.combo_coluna_principal["values"] = colunas

        if colunas:
            self.combo_coluna_principal.current(0)

    def atualizar_colunas_referencia(self, event=None):

        aba = self.combo_referencia.get()

        colunas = self.obter_colunas(aba)

        self.combo_coluna_busca["values"] = colunas
        self.combo_coluna_retorno["values"] = colunas

        if colunas:
            self.combo_coluna_busca.current(0)
            self.combo_coluna_retorno.current(
                min(1, len(colunas) - 1)
            )

    # ==========================================
    # CONVERTER SELEÇÃO EM NÚMERO DA COLUNA
    # ==========================================

    def numero_coluna(self, selecao):

        if not selecao:
            raise ValueError("Selecione todas as colunas.")

        letra = selecao.split(" - ")[0]

        numero = 0

        for caractere in letra:
            numero = numero * 26 + ord(caractere) - 64

        return numero

    # ==========================================
    # NORMALIZAR CÓDIGOS PARA COMPARAÇÃO
    # ==========================================

    def normalizar(self, valor):

        if valor is None:
            return None

        if isinstance(valor, bool):
            return str(valor).casefold()

        if isinstance(valor, (int, float)):
            if valor == int(valor):
                return str(int(valor))

        if isinstance(valor, (datetime, date)):
            return valor.isoformat()

        return str(valor).strip().casefold()

    # ==========================================
    # CONVERTER RESULTADO PARA TEXTO
    # ==========================================

    def converter_texto(self, valor):

        if valor is None:
            return ""

        if isinstance(valor, bool):
            return "VERDADEIRO" if valor else "FALSO"

        if isinstance(valor, (datetime, date)):
            return valor.isoformat()

        return str(valor)

    # ==========================================
    # EXECUTAR PROCV
    # ==========================================

    def executar_procv(self):

        if not self.workbook or not self.arquivo:

            messagebox.showwarning(
                "Atenção",
                "Primeiro selecione uma planilha Excel."
            )
            return

        try:

            # IDENTIFICAR ABAS

            nome_aba_principal = self.combo_principal.get()
            nome_aba_referencia = self.combo_referencia.get()

            # IDENTIFICAR COLUNAS

            coluna_principal = self.numero_coluna(
                self.combo_coluna_principal.get()
            )

            coluna_busca = self.numero_coluna(
                self.combo_coluna_busca.get()
            )

            coluna_retorno = self.numero_coluna(
                self.combo_coluna_retorno.get()
            )

            nome_resultado = self.entrada_resultado.get().strip()

            if not nome_resultado:

                messagebox.showwarning(
                    "Atenção",
                    "Informe o nome da coluna de resultado."
                )
                return

            ws_principal = self.workbook[
                nome_aba_principal
            ]

            ws_referencia = self.workbook[
                nome_aba_referencia
            ]

            # MONTAR BANCO DE CONSULTA

            banco = {}

            for linha in range(
                2,
                ws_referencia.max_row + 1
            ):

                chave = ws_referencia.cell(
                    row=linha,
                    column=coluna_busca
                ).value

                resultado = ws_referencia.cell(
                    row=linha,
                    column=coluna_retorno
                ).value

                chave_normalizada = self.normalizar(chave)

                if chave_normalizada is None:
                    continue

                # Mantém a primeira ocorrência,
                # como o PROCV tradicional.

                if chave_normalizada not in banco:
                    banco[chave_normalizada] = resultado

            # LOCALIZAR OU CRIAR COLUNA DE RESULTADO

            coluna_resultado = None

            for numero in range(
                1,
                ws_principal.max_column + 1
            ):

                cabecalho = ws_principal.cell(
                    row=1,
                    column=numero
                ).value

                if (
                    cabecalho is not None
                    and str(cabecalho).strip() == nome_resultado
                ):

                    coluna_resultado = numero
                    break

            if coluna_resultado is None:

                coluna_resultado = (
                    ws_principal.max_column + 1
                )

            # ESCREVER CABEÇALHO

            celula_cabecalho = ws_principal.cell(
                row=1,
                column=coluna_resultado
            )

            celula_cabecalho.value = nome_resultado
            celula_cabecalho.number_format = "@"

            # PERCORRER TODAS AS LINHAS

            encontrados = 0
            nao_encontrados = 0

            for linha in range(
                2,
                ws_principal.max_row + 1
            ):

                codigo = ws_principal.cell(
                    row=linha,
                    column=coluna_principal
                ).value

                chave = self.normalizar(codigo)

                if chave is not None and chave in banco:

                    valor = self.converter_texto(
                        banco[chave]
                    )

                    encontrados += 1

                else:

                    valor = ""
                    nao_encontrados += 1

                celula = ws_principal.cell(
                    row=linha,
                    column=coluna_resultado
                )

                # GRAVA O VALOR, NUNCA UMA FÓRMULA

                celula.value = valor

                # FORMATO DE TEXTO NO EXCEL

                celula.number_format = "@"

            # SALVAR NOVO ARQUIVO

            nome_base = os.path.splitext(
                os.path.basename(self.arquivo)
            )[0]

            destino = filedialog.asksaveasfilename(
                title="Salvar planilha com resultado",
                initialfile=f"{nome_base}_resultado.xlsx",
                defaultextension=".xlsx",
                filetypes=[
                    ("Planilha Excel", "*.xlsx")
                ]
            )

            if not destino:
                self.status.config(
                    text="Exportação cancelada. A planilha original não foi alterada."
                )
                return

            self.workbook.save(destino)

            self.status.config(
                text=(
                    f"Concluído! Encontrados: {encontrados} | "
                    f"Não encontrados: {nao_encontrados}"
                ),
                fg="#16834A"
            )

            messagebox.showinfo(
                "PROCV concluído!",
                f"Processamento finalizado!\n\n"
                f"Registros encontrados: {encontrados}\n"
                f"Não encontrados: {nao_encontrados}\n\n"
                f"Arquivo salvo em:\n{destino}\n\n"
                f"Os resultados foram gravados como texto, "
                f"sem fórmulas."
            )

        except Exception as erro:

            messagebox.showerror(
                "Erro no processamento",
                f"Ocorreu um erro:\n\n{erro}"
            )


# ==========================================
# INICIAR PROGRAMA
# ==========================================

if __name__ == "__main__":

    janela = tk.Tk()

    app = ExcelProcv(janela)

    janela.mainloop()
