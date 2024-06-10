import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Optional, List
import matplotlib.pyplot as plt

class PlanilhaApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Busca na Planilha")
        self.df: Optional[pd.DataFrame] = None

        self.setup_ui()

    def setup_ui(self) -> None:
        # Estilo do ttk
        style = ttk.Style(self.root)
        style.theme_use('clam')

        # Frame principal
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Frame para seleção de arquivo
        frame_selecionar_arquivo = ttk.LabelFrame(main_frame, text="Selecionar Arquivo", padding="10")
        frame_selecionar_arquivo.grid(row=0, column=0, padx=10, pady=10, sticky=(tk.W, tk.E))

        label_selecionar_arquivo = ttk.Label(frame_selecionar_arquivo, text="Selecione a planilha:")
        label_selecionar_arquivo.grid(row=0, column=0, padx=5, pady=5)
        
        button_procurar = ttk.Button(frame_selecionar_arquivo, text="Procurar", command=self.carregar_planilha)
        button_procurar.grid(row=0, column=1, padx=5, pady=5)

        # Frame para busca
        frame_buscar = ttk.LabelFrame(main_frame, text="Busca", padding="10")
        frame_buscar.grid(row=1, column=0, padx=10, pady=10, sticky=(tk.W, tk.E))

        label_busca = ttk.Label(frame_buscar, text="Buscar:")
        label_busca.grid(row=0, column=0, padx=5, pady=5)
        
        self.entry_busca = ttk.Entry(frame_buscar, width=50)
        self.entry_busca.grid(row=0, column=1, padx=5, pady=5)
        
        button_buscar = ttk.Button(frame_buscar, text="Buscar", command=self.buscar)
        button_buscar.grid(row=0, column=2, padx=5, pady=5)

        # Frame para filtros avançados
        frame_filtros = ttk.LabelFrame(main_frame, text="Filtros Avançados", padding="10")
        frame_filtros.grid(row=2, column=0, padx=10, pady=10, sticky=(tk.W, tk.E))

        label_colunas = ttk.Label(frame_filtros, text="Colunas:")
        label_colunas.grid(row=0, column=0, padx=5, pady=5)
        
        self.colunas_var = tk.StringVar(value=[])
        self.listbox_colunas = tk.Listbox(frame_filtros, listvariable=self.colunas_var, selectmode='multiple', height=6)
        self.listbox_colunas.grid(row=0, column=1, padx=5, pady=5)

        label_criterio = ttk.Label(frame_filtros, text="Critério:")
        label_criterio.grid(row=1, column=0, padx=5, pady=5)

        self.criterio_var = tk.StringVar(value="Contém")
        self.combobox_criterio = ttk.Combobox(frame_filtros, textvariable=self.criterio_var)
        self.combobox_criterio['values'] = ["Contém", "Igual", "Começa com", "Termina com"]
        self.combobox_criterio.grid(row=1, column=1, padx=5, pady=5)

        # Frame para processamento de dados
        frame_processamento = ttk.LabelFrame(main_frame, text="Processamento de Dados", padding="10")
        frame_processamento.grid(row=3, column=0, padx=10, pady=10, sticky=(tk.W, tk.E))

        label_processo = ttk.Label(frame_processamento, text="Processamento:")
        label_processo.grid(row=0, column=0, padx=5, pady=5)

        self.processamento_var = tk.StringVar(value="Selecione")
        self.combobox_processo = ttk.Combobox(frame_processamento, textvariable=self.processamento_var)
        self.combobox_processo['values'] = ["Filtrar", "Soma", "Média", "Contagem", "Normalizar", "Criar Coluna"]
        self.combobox_processo.grid(row=0, column=1, padx=5, pady=5)

        button_aplicar = ttk.Button(frame_processamento, text="Aplicar", command=self.aplicar_processamento)
        button_aplicar.grid(row=0, column=2, padx=5, pady=5)

        # Frame para visualização de dados
        frame_visualizacao = ttk.LabelFrame(main_frame, text="Visualização de Dados", padding="10")
        frame_visualizacao.grid(row=4, column=0, padx=10, pady=10, sticky=(tk.W, tk.E))

        label_grafico = ttk.Label(frame_visualizacao, text="Gráfico:")
        label_grafico.grid(row=0, column=0, padx=5, pady=5)

        self.grafico_var = tk.StringVar(value="Selecione")
        self.combobox_grafico = ttk.Combobox(frame_visualizacao, textvariable=self.grafico_var)
        self.combobox_grafico['values'] = ["Histograma", "Dispersão", "Pizza", "Barras"]
        self.combobox_grafico.grid(row=0, column=1, padx=5, pady=5)

        button_grafico = ttk.Button(frame_visualizacao, text="Visualizar", command=self.visualizar_grafico)
        button_grafico.grid(row=0, column=2, padx=5, pady=5)

        button_salvar_grafico = ttk.Button(frame_visualizacao, text="Salvar Gráfico", command=self.salvar_grafico)
        button_salvar_grafico.grid(row=0, column=3, padx=5, pady=5)

        # Barra de progresso
        self.progresso = ttk.Progressbar(main_frame, orient='horizontal', mode='determinate')
        self.progresso.grid(row=5, column=0, padx=10, pady=10, sticky=(tk.W, tk.E))

    def carregar_planilha(self) -> None:
        caminho_arquivo: str = filedialog.askopenfilename(filetypes=[("Todos os Arquivos", "*.*"),
                                                                    ("Planilha Excel", "*.xlsx"),
                                                                    ("Arquivo CSV", "*.csv"),
                                                                    ("Arquivo XML", "*.xml")])
        if caminho_arquivo:
            try:
                if caminho_arquivo.endswith('.xlsx'):
                    self.planilhas = pd.read_excel(caminho_arquivo, sheet_name=None)
                    sheet_names = list(self.planilhas.keys())
                    self.selecionar_aba(sheet_names)
                elif caminho_arquivo.endswith('.csv'):
                    self.df = pd.read_csv(caminho_arquivo)
                    self.colunas_var.set(self.df.columns.tolist())
                    messagebox.showinfo("Sucesso", "Arquivo carregado com sucesso!")
                elif caminho_arquivo.endswith('.xml'):
                    self.df = pd.read_xml(caminho_arquivo)
                    self.colunas_var.set(self.df.columns.tolist())
                    messagebox.showinfo("Sucesso", "Arquivo carregado com sucesso!")
                else:
                    raise ValueError("Tipo de arquivo não suportado.")
            except Exception as e:
                messagebox.showerror("Erro", f"Ocorreu um erro ao carregar o arquivo: {e}")

    def selecionar_aba(self, sheet_names: List[str]) -> None:
        def carregar_aba() -> None:
            aba_selecionada = self.aba_var.get()
            self.df = self.planilhas[aba_selecionada]
            self.colunas_var.set(self.df.columns.tolist())
            messagebox.showinfo("Sucesso", f"Aba '{aba_selecionada}' carregada com sucesso!")
            aba_window.destroy()

        aba_window = tk.Toplevel(self.root)
        aba_window.title("Selecionar Aba")
        
        label_abas = ttk.Label(aba_window, text="Selecione a aba:")
        label_abas.pack(padx=10, pady=10)

        self.aba_var = tk.StringVar(value=sheet_names[0])
        combobox_abas = ttk.Combobox(aba_window, textvariable=self.aba_var, values=sheet_names, state='readonly')
        combobox_abas.pack(padx=10, pady=10)

        button_carregar_aba = ttk.Button(aba_window, text="Carregar", command=carregar_aba)
        button_carregar_aba.pack(padx=10, pady=10)

    def buscar(self) -> None:
        if self.df is None:
            messagebox.showwarning("Aviso", "Por favor, carregue o arquivo primeiro.")
            return

        palavra: str = self.entry_busca.get().lower()
        colunas_selecionadas = [self.df.columns[i] for i in self.listbox_colunas.curselection()]
        criterio = self.criterio_var.get()
        linhas_resultados: List[pd.Series] = []

        if not colunas_selecionadas:
            colunas_selecionadas = self.df.columns.tolist()

        self.progresso['maximum'] = len(self.df)
        self.progresso['value'] = 0

        for index, row in self.df.iterrows():
            self.progresso['value'] = index + 1
            self.root.update_idletasks()

            for coluna in colunas_selecionadas:
                valor_celula = str(row[coluna]).lower()
                if pd.notna(row[coluna]) and self.match_criteria(valor_celula, palavra, criterio):
                    linhas_resultados.append(row)
                    break

        self.progresso['value'] = 0

        if not linhas_resultados:
            messagebox.showinfo("Informação", "Nenhuma correspondência encontrada.")
        else:
            self.mostrar_resultados(linhas_resultados)

    def match_criteria(self, valor: str, palavra: str, criterio: str) -> bool:
        if criterio == "Contém":
            return palavra in valor
        elif criterio == "Igual":
            return palavra == valor
        elif criterio == "Começa com":
            return valor.startswith(palavra)
        elif criterio == "Termina com":
            return valor.endswith(palavra)
        return False

    def mostrar_resultados(self, resultados: List[pd.Series]) -> None:
        resultado_window = tk.Toplevel(self.root)
        resultado_window.title("Resultados da Busca")

        tree = ttk.Treeview(resultado_window, columns=self.df.columns.tolist(), show='headings')
        tree.pack(fill=tk.BOTH, expand=True)

        # Configurar cabeçalhos
        for col in self.df.columns:
            tree.heading(col, text=col)
            tree.column(col, minwidth=0, width=100)

        # Inserir resultados
        for linha in resultados:
            tree.insert("", tk.END, values=linha.tolist())

        # Adicionar scrollbar
        scrollbar = ttk.Scrollbar(resultado_window, orient="vertical", command=tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        tree.configure(yscrollcommand=scrollbar.set)

        # Botão para exportar resultados
        button_exportar = ttk.Button(resultado_window, text="Exportar Resultados", command=lambda: self.exportar_resultados(resultados))
        button_exportar.pack(pady=10)

    def exportar_resultados(self, resultados: List[pd.Series]) -> None:
        if not resultados:
            messagebox.showwarning("Aviso", "Nenhum resultado para exportar.")
            return

        caminho_arquivo = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("Arquivo CSV", "*.csv"), ("Planilha Excel", "*.xlsx")])
        if caminho_arquivo:
            try:
                resultados_df = pd.DataFrame(resultados)
                if caminho_arquivo.endswith('.csv'):
                    resultados_df.to_csv(caminho_arquivo, index=False)
                elif caminho_arquivo.endswith('.xlsx'):
                    resultados_df.to_excel(caminho_arquivo, index=False)
                messagebox.showinfo("Sucesso", "Resultados exportados com sucesso!")
            except Exception as e:
                messagebox.showerror("Erro", f"Ocorreu um erro ao exportar os resultados: {e}")

    def aplicar_processamento(self) -> None:
        if self.df is None:
            messagebox.showwarning("Aviso", "Por favor, carregue o arquivo primeiro.")
            return

        processo = self.processamento_var.get()
        if processo == "Filtrar":
            self.aplicar_filtro()
        elif processo == "Soma":
            self.aplicar_agregacao("sum")
        elif processo == "Média":
            self.aplicar_agregacao("mean")
        elif processo == "Contagem":
            self.aplicar_agregacao("count")
        elif processo == "Normalizar":
            self.normalizar_dados()
        elif processo == "Criar Coluna":
            self.criar_coluna()
        else:
            messagebox.showwarning("Aviso", "Processamento não selecionado ou inválido.")

    def aplicar_filtro(self) -> None:
        colunas_selecionadas = [self.df.columns[i] for i in self.listbox_colunas.curselection()]
        if not colunas_selecionadas:
            messagebox.showwarning("Aviso", "Selecione pelo menos uma coluna para filtrar.")
            return

        palavra = self.entry_busca.get().lower()
        criterio = self.criterio_var.get()
        df_filtrado = self.df

        for coluna in colunas_selecionadas:
            if criterio == "Contém":
                df_filtrado = df_filtrado[df_filtrado[coluna].str.contains(palavra, case=False, na=False)]
            elif criterio == "Igual":
                df_filtrado = df_filtrado[df_filtrado[coluna].str.lower() == palavra]
            elif criterio == "Começa com":
                df_filtrado = df_filtrado[df_filtrado[coluna].str.lower().str.startswith(palavra)]
            elif criterio == "Termina com":
                df_filtrado = df_filtrado[df_filtrado[coluna].str.lower().str.endswith(palavra)]

        self.mostrar_resultados(df_filtrado.to_dict(orient='records'))

    def aplicar_agregacao(self, metodo: str) -> None:
        colunas_selecionadas = [self.df.columns[i] for i in self.listbox_colunas.curselection()]
        if not colunas_selecionadas:
            messagebox.showwarning("Aviso", "Selecione pelo menos uma coluna para agregar.")
            return

        try:
            if metodo == "sum":
                resultado = self.df[colunas_selecionadas].sum()
            elif metodo == "mean":
                resultado = self.df[colunas_selecionadas].mean()
            elif metodo == "count":
                resultado = self.df[colunas_selecionadas].count()

            resultado_df = resultado.reset_index()
            resultado_df.columns = ['Coluna', metodo.capitalize()]

            self.mostrar_resultados(resultado_df.to_dict(orient='records'))
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro ao aplicar a agregação: {e}")

    def normalizar_dados(self) -> None:
        colunas_selecionadas = [self.df.columns[i] for i in self.listbox_colunas.curselection()]
        if not colunas_selecionadas:
            messagebox.showwarning("Aviso", "Selecione pelo menos uma coluna para normalizar.")
            return

        try:
            self.df[colunas_selecionadas] = (self.df[colunas_selecionadas] - self.df[colunas_selecionadas].min()) / (self.df[colunas_selecionadas].max() - self.df[colunas_selecionadas].min())

            messagebox.showinfo("Sucesso", "Dados normalizados com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro ao normalizar os dados: {e}")

    def criar_coluna(self) -> None:
        nova_coluna = tk.simpledialog.askstring("Input", "Nome da nova coluna:")
        if not nova_coluna:
            messagebox.showwarning("Aviso", "Nome da nova coluna não fornecido.")
            return

        expressao = tk.simpledialog.askstring("Input", "Expressão para calcular a nova coluna (ex: coluna1 + coluna2):")
        if not expressao:
            messagebox.showwarning("Aviso", "Expressão não fornecida.")
            return

        try:
            self.df[nova_coluna] = self.df.eval(expressao)
            self.colunas_var.set(self.df.columns.tolist())
            messagebox.showinfo("Sucesso", "Nova coluna criada com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro ao criar a nova coluna: {e}")

    def visualizar_grafico(self) -> None:
        if self.df is None:
            messagebox.showwarning("Aviso", "Por favor, carregue o arquivo primeiro.")
            return

        grafico = self.grafico_var.get()
        colunas_selecionadas = [self.df.columns[i] for i in self.listbox_colunas.curselection()]

        if not colunas_selecionadas:
            messagebox.showwarning("Aviso", "Selecione pelo menos uma coluna para visualização.")
            return

        try:
            if grafico == "Histograma":
                self.df[colunas_selecionadas].hist()
            elif grafico == "Dispersão" and len(colunas_selecionadas) >= 2:
                self.df.plot.scatter(x=colunas_selecionadas[0], y=colunas_selecionadas[1])
            elif grafico == "Pizza" and len(colunas_selecionadas) == 1:
                self.df[colunas_selecionadas[0]].value_counts().plot.pie(autopct='%1.1f%%')
            elif grafico == "Barras" and len(colunas_selecionadas) == 1:
                self.df[colunas_selecionadas[0]].value_counts().plot.bar()
            else:
                messagebox.showwarning("Aviso", "Gráfico não selecionado ou colunas insuficientes/erradas para o tipo de gráfico.")
                return

            plt.show()
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro ao visualizar o gráfico: {e}")

    def salvar_grafico(self) -> None:
        caminho_arquivo = filedialog.asksaveasfilename(defaultextension=".png", filetypes=[("Imagem PNG", "*.png"), ("Imagem JPEG", "*.jpg"), ("PDF", "*.pdf")])
        if caminho_arquivo:
            try:
                plt.savefig(caminho_arquivo)
                messagebox.showinfo("Sucesso", "Gráfico salvo com sucesso!")
            except Exception as e:
                messagebox.showerror("Erro", f"Ocorreu um erro ao salvar o gráfico: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = PlanilhaApp(root)
    root.mainloop()
