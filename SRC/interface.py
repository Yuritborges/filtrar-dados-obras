import os
import pandas as pd
import customtkinter as ctk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageTk
import ctypes

# Deixa o visual do programa no modo claro
ctk.set_appearance_mode("light")

class DashboardBrasul(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Truque pro Windows não confundir o programa com o ícone do Python
        try:
            myappid = 'brasul.tecnologia.gestaoinsumos.v12'
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
        except:
            pass

        # Configura o título da janela e o tamanho da tela
        self.title("BRASUL - GESTÃO DE INSUMOS v3.0")
        self.geometry("1580x850")
        self.configure(fg_color="#F5F5F5")

        # Acha onde as pastas e arquivos estão
        diretorio = os.path.dirname(os.path.abspath(__file__))
        self.caminho_db = os.path.join(os.path.dirname(diretorio), "DATA", "output", "Banco_Mestre_Brasul.xlsx")

        # Coloca o ícone da Brasul no topo da janela
        caminho_icone = os.path.join(diretorio, "ICONE_BRASUL.png")
        if os.path.exists(caminho_icone):
            try:
                self.iconbitmap(caminho_icone)
                img_icon = Image.open(caminho_icone)
                self.icon_photo = ImageTk.PhotoImage(img_icon)
                self.wm_iconphoto(False, self.icon_photo)
            except Exception as e:
                print(f"Erro no ícone: {e}")

        # Arruma o grid pra dividir a tela em duas partes
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # Barra lateral branca (Sidebar)
        self.sidebar = ctk.CTkFrame(self, width=280, fg_color="#FFFFFF", corner_radius=0)
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        self.sidebar.grid_propagate(False)

        # Coloca o logo da Brasul na barra lateral
        caminho_logo = os.path.join(diretorio, "LOGOTIPOBRASUL.png")
        if os.path.exists(caminho_logo):
            logo_raw = Image.open(caminho_logo)
            self.logo_img = ctk.CTkImage(logo_raw, size=(220, 90))
            ctk.CTkLabel(self.sidebar, image=self.logo_img, text="").pack(pady=40)

        # Botão verde pra salvar os dados no Excel
        self.btn_export = ctk.CTkButton(self.sidebar, text="📊 EXPORTAR EXCEL",
                                        fg_color="#27ae60", hover_color="#1e8449",
                                        height=45, font=ctk.CTkFont(weight="bold"),
                                        command=self.exportar_excel)
        self.btn_export.pack(pady=10, padx=20, fill="x")

        # Área da direita onde os dados aparecem
        self.main_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.main_frame.grid(row=0, column=1, padx=20, pady=20, sticky="nsew")

        # Caixinha branca onde fica a barra de busca
        self.search_container = ctk.CTkFrame(self.main_frame, fg_color="#FFFFFF", corner_radius=12)
        self.search_container.pack(fill="x", pady=(0, 20))

        # Onde o usuário digita o que quer procurar
        self.entry_busca = ctk.CTkEntry(self.search_container,
                                        placeholder_text="O que você procura? (Ex: Madeira, Aço, %, M2...)",
                                        height=55, border_width=0, fg_color="transparent",
                                        font=ctk.CTkFont(size=15))
        self.entry_busca.pack(side="left", padx=20, fill="x", expand=True)
        self.entry_busca.bind("<Return>", lambda e: self.pesquisar())

        # Botão laranja pra clicar e buscar
        self.btn_search = ctk.CTkButton(self.search_container, text="BUSCAR",
                                        fg_color="#d95947", hover_color="#b04132",
                                        width=140, height=45, font=ctk.CTkFont(weight="bold"),
                                        command=self.pesquisar)
        self.btn_search.pack(side="right", padx=15)

        # Estilo visual da tabela de dados
        style = ttk.Style()
        style.theme_use("default")
        style.configure("Treeview", background="#FFFFFF", fieldbackground="#FFFFFF",
                        rowheight=35, font=('Segoe UI', 10))
        style.configure("Treeview.Heading", font=('Segoe UI', 10, 'bold'), background="#EEEEEE")
        style.map("Treeview", background=[('selected', '#d95947')])

        # Montagem das 9 colunas que o robô extraiu
        self.cols = ("Obra", "Cod", "Desc", "UN", "Q_Orc", "Q_Acum", "Q_Per", "V_Acum", "V_Per")
        self.tabela = ttk.Treeview(self.main_frame, columns=self.cols, show='headings')

        # Nomes que ficam no topo de cada coluna
        headers = ["OBRA/ARQUIVO", "CÓDIGO", "DESCRIÇÃO DO SERVIÇO", "UN", "QTD ORÇ.", "QTD ACUM.", "QTD PER.",
                   "VALOR ACUM.", "VALOR PER."]
        larguras = [160, 95, 420, 50, 100, 100, 100, 120, 120]

        for col, head, w in zip(self.cols, headers, larguras):
            self.tabela.heading(col, text=head)
            self.tabela.column(col, width=w, anchor="center")

        self.tabela.column("Desc", anchor="w")

        # Barra lateral pra descer a tabela se tiver muita coisa
        scrollbar = ttk.Scrollbar(self.main_frame, orient="vertical", command=self.tabela.yview)
        self.tabela.configure(yscrollcommand=scrollbar.set)

        self.tabela.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Botão cinza lá embaixo pra limpar a busca
        self.btn_reset = ctk.CTkButton(self, text="LIMPAR BUSCA", fg_color="#555555",
                                       command=self.resetar, width=120)
        self.btn_reset.place(relx=0.98, rely=0.98, anchor="se")

        self.df_data = pd.DataFrame()
        self.entry_busca.focus()

    # Função que procura as palavras no arquivo Excel
    def pesquisar(self):
        termo = self.entry_busca.get().strip().upper()
        if not termo: return

        if os.path.exists(self.caminho_db):
            try:
                # Carrega o Excel e tira os erros de espaço vazio
                df = pd.read_excel(self.caminho_db).fillna('')

                # Varredura inteligente em 3 frentes: Descrição, Código ou Unidade
                self.df_data = df[df['Descricao'].str.contains(termo, na=False) |
                                  df['Codigo'].astype(str).str.contains(termo, na=False) |
                                  df['UN'].str.contains(termo, na=False)]

                # Limpa o que tava na tela antes de mostrar o novo
                for i in self.tabela.get_children(): self.tabela.delete(i)

                # Joga os resultados na tabela
                for _, r in self.df_data.iterrows():
                    self.tabela.insert("", "end", values=list(r))

                if self.df_data.empty:
                    messagebox.showinfo("Busca", f"Não achei nada para: {termo}")
            except Exception as e:
                messagebox.showerror("Erro", f"Zicou ao ler o Excel: {e}")
        else:
            messagebox.showwarning("Aviso", "O Banco de Dados sumiu. Rode o main primeiro!")

    # Função pra salvar o resultado da tela em um novo arquivo
    def exportar_excel(self):
        if self.df_data.empty:
            messagebox.showwarning("Exportar", "Não tem nada na tela pra exportar.")
            return

        path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                            filetypes=[("Excel files", "*.xlsx")])
        if path:
            try:
                self.df_data.to_excel(path, index=False)
                messagebox.showinfo("Sucesso", "Planilha salva com sucesso!")
            except Exception as e:
                messagebox.showerror("Erro", f"Não consegui salvar: {e}")

    # Limpa a caixa de busca e a tabela
    def resetar(self):
        self.entry_busca.delete(0, 'end')
        for i in self.tabela.get_children(): self.tabela.delete(i)
        self.df_data = pd.DataFrame()
        self.entry_busca.focus()

if __name__ == "__main__":
    app = DashboardBrasul()
    app.mainloop()