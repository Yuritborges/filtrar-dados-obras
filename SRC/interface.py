import os
import pandas as pd
import customtkinter as ctk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageTk
import threading
import shutil
import ctypes
import unicodedata
import sys # Necessário para detectar o executável

# --- FUNÇÃO PARA PEGAR O CAMINHO DOS RECURSOS (ÍCONES/IMAGENS) DENTRO DO EXE ---
def resource_path(relative_path):
    """ Retorna o caminho absoluto para o recurso, seja no PyCharm ou no .exe """
    try:
        # O PyInstaller cria uma pasta temporária e armazena o caminho em _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# COMANDO PARA TENTAR IMPORTAR O MAIN.PY
try:
    from main import extrair_total_brasul
except ImportError:
    pass

ctk.set_appearance_mode("light")

class DashboardBrasul(ctk.CTk):
    def __init__(self):
        super().__init__()

        # --- 1. LÓGICA DE CAMINHOS ---
        if getattr(sys, 'frozen', False):
            # Se for o .exe (na máquina do chefe)
            self.diretorio_base = os.path.dirname(sys.executable)
        else:
            # Se for no PyCharm (na sua máquina)
            diretorio_src = os.path.dirname(os.path.abspath(__file__))
            self.diretorio_base = os.path.dirname(diretorio_src)

        # Pastas de dados (Ficam FORA do .exe para o usuário mexer)
        self.pasta_input = os.path.join(self.diretorio_base, "DATA", "input")
        self.caminho_db = os.path.join(self.diretorio_base, "DATA", "output", "Banco_Mestre_Brasul.xlsx")

        # Imagens e Ícones (Usando resource_path para buscar DENTRO do .exe)
        caminho_icone = resource_path("iconebrasul2.ico")
        caminho_logo = resource_path("LOGOTIPOBRASUL.png")

        # --- 2. CONFIGURAÇÕES DA JANELA ---
        self.title("GESTÃO DE INSUMOS BRASUL v1.0")
        self.geometry("1580x850")
        self.configure(fg_color="#F5F5F5")

        # CONFIGURAÇÃO DOS ÍCONES (Janela e Barra de Tarefas)
        if os.path.exists(caminho_icone):
            try:
                # Força o ícone na Barra de Tarefas do Windows
                myappid = 'brasul.gestao.v12'
                ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)

                # Ícone no Topo da Janela
                img_icon = Image.open(caminho_icone)
                self.icon_photo = ImageTk.PhotoImage(img_icon)
                self.after(200, lambda: self.wm_iconphoto(False, self.icon_photo))
            except:
                pass

        self.df_completo = pd.DataFrame()
        self.df_data = pd.DataFrame()
        self.carregar_banco_inicial()

        # --- 3. LAYOUT PRINCIPAL ---
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # BARRA LATERAL (SIDEBAR)
        self.sidebar = ctk.CTkFrame(self, width=300, fg_color="#FFFFFF", corner_radius=0)
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        self.sidebar.grid_propagate(False)

        # LOGOTIPO GRANDE
        if os.path.exists(caminho_logo):
            img_raw = Image.open(caminho_logo)
            self.logo_img = ctk.CTkImage(img_raw, size=(240, 90))
            ctk.CTkLabel(self.sidebar, image=self.logo_img, text="").pack(pady=(50, 60))

        font_btns = ctk.CTkFont(family="Arial", size=13, weight="bold")

        # BOTÕES LATERAIS
        self.btn_import = ctk.CTkButton(self.sidebar, text="IMPORTAR PDF",
                                        fg_color="#FFCC00", text_color="black", height=50,
                                        font=font_btns, command=self.importar_arquivo_thread)
        self.btn_import.pack(pady=10, padx=25, fill="x")

        self.btn_export = ctk.CTkButton(self.sidebar, text="EXPORTAR BUSCA",
                                        fg_color="#2ecc71", text_color="black", height=50,
                                        font=font_btns, command=self.exportar_excel)
        self.btn_export.pack(pady=10, padx=25, fill="x")

        self.progresso = ctk.CTkProgressBar(self.sidebar, mode="indeterminate", progress_color="#FFCC00")
        self.lbl_status = ctk.CTkLabel(self.sidebar, text="PROCESSANDO...", font=("Arial", 11, "italic"),
                                       text_color="gray")

        # --- 4. ÁREA DE BUSCA ---
        self.main_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.main_frame.grid(row=0, column=1, padx=25, pady=25, sticky="nsew")

        self.search_container = ctk.CTkFrame(self.main_frame, fg_color="#FFFFFF", corner_radius=15,
                                             border_width=1, border_color="#CCCCCC")
        self.search_container.pack(fill="x", pady=(0, 20))

        self.entry_busca = ctk.CTkEntry(self.search_container,
                                        placeholder_text="BUSQUE POR MATERIAL OU CÓDIGO (Ex: AÇO, 02.03.001)...",
                                        height=60, border_width=0, fg_color="transparent", font=ctk.CTkFont(size=16))
        self.entry_busca.pack(side="left", padx=20, fill="x", expand=True)
        self.entry_busca.bind("<Return>", lambda e: self.pesquisar("ACUMULADO"))

        # BOTÕES DE BUSCA
        self.btn_med = ctk.CTkButton(self.search_container, text="ACUMULO DE MEDIÇÃO",
                                     fg_color="#3498db", text_color="black", width=220, height=50,
                                     font=font_btns, command=lambda: self.pesquisar("ACUMULADO"))
        self.btn_med.pack(side="right", padx=15)

        self.btn_orc = ctk.CTkButton(self.search_container, text="PLANILHA QUANTITATIVA",
                                     fg_color="#e67e22", text_color="black", width=220, height=50,
                                     font=font_btns, command=lambda: self.pesquisar("QUANTITATIVA"))
        self.btn_orc.pack(side="right", padx=5)

        # --- 5. TABELA ---
        style = ttk.Style()
        style.theme_use("default")
        style.configure("Treeview", background="#FFFFFF", fieldbackground="#FFFFFF", rowheight=38,
                        font=('Segoe UI', 10), borderwidth=0)
        style.configure("Treeview.Heading", font=('Segoe UI', 10, 'bold'), background="#F2F2F2",
                        foreground="black", relief="groove", borderwidth=1)
        style.map("Treeview", background=[('selected', '#FFCC00')], foreground=[('selected', 'black')])

        self.cols = ("Obra", "Cod", "Desc", "UN", "Q_Orc", "Q_Acum")
        self.tabela = ttk.Treeview(self.main_frame, columns=self.cols, show='headings')

        heads = ["ESCOLA / OBRA", "CÓDIGO", "DESCRIÇÃO DO SERVIÇO", "UN", "QT ORÇADA", "QT ACUMULADA"]
        for c, h in zip(self.cols, heads):
            self.tabela.heading(c, text=h)
            self.tabela.column(c, width=120, anchor="center")
        self.tabela.column("Desc", width=450, anchor="w")

        scrollbar = ttk.Scrollbar(self.main_frame, orient="vertical", command=self.tabela.yview)
        self.tabela.configure(yscrollcommand=scrollbar.set)
        self.tabela.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # --- 6. RODAPÉ ---
        self.lbl_contador = ctk.CTkLabel(self, text="", font=("Arial", 12, "bold"))
        self.lbl_contador.place(relx=0.21, rely=0.97, anchor="sw")

        self.btn_reset = ctk.CTkButton(self, text="🧹 LIMPAR PESQUISA", fg_color="#D1D1D1",
                                       text_color="black", width=180, height=40,
                                       font=font_btns, command=self.limpar)
        self.btn_reset.place(relx=0.98, rely=0.97, anchor="se")
        self.entry_busca.focus()

    def carregar_banco_inicial(self):
        if os.path.exists(self.caminho_db):
            try:
                self.df_completo = pd.read_excel(self.caminho_db).fillna('')
            except: pass

    def remover_acentos(self, txt):
        return "".join(c for c in unicodedata.normalize('NFD', str(txt)) if unicodedata.category(c) != 'Mn')

    def pesquisar(self, tipo):
        termo = self.entry_busca.get().strip().upper()
        if not termo or self.df_completo.empty: return
        termo_n = self.remover_acentos(termo)

        def match(row):
            desc_n = self.remover_acentos(str(row['Desc']).upper())
            cod_n = str(row['Cod']).upper()
            tipo_row = str(row['Tipo']).upper()
            return (termo_n in desc_n or termo_n in cod_n) and (tipo in tipo_row)

        self.df_data = self.df_completo[self.df_completo.apply(match, axis=1)]
        [self.tabela.delete(i) for i in self.tabela.get_children()]

        for _, r in self.df_data.head(1000).iterrows():
            q_ac = r['Q_Acum'] if tipo == "ACUMULADO" else "---"
            self.tabela.insert("", "end", values=[r['Obra'], r['Cod'], r['Desc'], r['UN'], r['Q_Orc'], q_ac])

        self.lbl_contador.configure(text=f"ITENS ENCONTRADOS EM {tipo}: {len(self.df_data)}")

    def limpar(self):
        self.entry_busca.delete(0, 'end')
        [self.tabela.delete(i) for i in self.tabela.get_children()]
        self.lbl_contador.configure(text="")
        self.entry_busca.focus()

    def importar_arquivo_thread(self):
        c = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if c:
            if not os.path.exists(self.pasta_input): os.makedirs(self.pasta_input)
            shutil.copy(c, os.path.join(self.pasta_input, os.path.basename(c)))
            self.btn_import.configure(state="disabled", text="🕒 LENDO...")
            self.progresso.pack(pady=10, padx=25, fill="x")
            self.progresso.start()
            self.lbl_status.pack()
            threading.Thread(target=self.run_ocr, daemon=True).start()

    def run_ocr(self):
        try:
            extrair_total_brasul()
            self.after(0, self.done)
        except Exception as e:
            self.after(0, lambda: messagebox.showerror("Erro", str(e)))

    def done(self):
        self.progresso.stop()
        self.progresso.pack_forget()
        self.lbl_status.pack_forget()
        self.btn_import.configure(state="normal", text="IMPORTAR PDF")
        self.carregar_banco_inicial()
        messagebox.showinfo("SUCESSO", "DADOS INTEGRADOS!")
        self.limpar()

    def exportar_excel(self):
        if not self.df_data.empty:
            p = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                             filetypes=[("Excel files", "*.xlsx")])
            if p:
                try:
                    self.df_data.to_excel(p, index=False)
                    messagebox.showinfo("SUCESSO", "PLANILHA EXPORTADA!")
                except Exception as e:
                    messagebox.showerror("ERRO", str(e))
        else:
            messagebox.showwarning("AVISO", "NÃO HÁ RESULTADOS PARA EXPORTAR.")

if __name__ == "__main__":
    app = DashboardBrasul()
    app.mainloop()