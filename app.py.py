import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog, simpledialog
import re
import csv
import json
import os
import copy
import time
import threading
import webbrowser
import subprocess
import urllib.request
import sys
from datetime import datetime

# --- VERSÃO ATUAL DO PROGRAMA PARA AUTO-UPDATE ---
__version__ = "1.0.7"
URL_VERSAO_REMOTE = "https://raw.githubusercontent.com/getyourwish10-byte/ORGANIZADOR_LIDER/main/version.json"

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    OPENPYXL_DISPONIVEL = True
except ImportError:
    OPENPYXL_DISPONIVEL = False

try:
    from PIL import Image, ImageGrab, ImageOps, ImageEnhance
    PIL_DISPONIVEL = True
except ImportError:
    PIL_DISPONIVEL = False

try:
    import pytesseract
    PYTESSERACT_DISPONIVEL = True
    caminhos_tesseract = [
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        r"C:\Users\{}\AppData\Local\Tesseract-OCR\tesseract.exe".format(os.getenv("USERNAME", ""))
    ]
    for caminho_tes in caminhos_tesseract:
        if os.path.exists(caminho_tes):
            pytesseract.pytesseract.tesseract_cmd = caminho_tes
            break
except ImportError:
    PYTESSERACT_DISPONIVEL = False

try:
    import ollama
    OLLAMA_DISPONIVEL = True
except ImportError:
    OLLAMA_DISPONIVEL = False


CONFIG_PATH = os.path.join(os.path.expanduser("~"), ".organizador_pedidos_config.json")
HISTORICO_MAXIMO = 30


class AplicativoPedidosMagico:
    def __init__(self, root):
        self.root = root
        self.root.title("Organizador Inteligente de Pedidos - by Douglas Oliveira")
        self.root.geometry("1280x880")
        self.root.minsize(980, 680)
        self.root.configure(bg="#F4F6F9")

        # --- REGRAS DO SISTEMA ---
        self.ordem_tamanhos_padrao = [
            "G4", "G3", "G2", "GG", "G", "M", "P",
            "BLG3", "BLG2", "BLGG", "BLG", "BLM", "BLP", "BLPP",
            "13_14_PP", "11_12_GG_INF", "8_10_G_INF", "6_7_M_INF",
            "4_5_P_INF", "2_3_PP_INF", "1_BB", "0_RN"
        ]

        self.conversoes_padrao = {
            "EXG": "G2", "EXGG": "G3", "G1": "GG",
            "BLG1": "BLGG", "BABYLOOK G1": "BLGG",
            "PP": "13_14_PP", "13 A 14 ANOS": "13_14_PP", "13 A 14": "13_14_PP", "13-14 ANOS": "13_14_PP",
            "GG INFANTIL": "11_12_GG_INF", "INFANTIL GG": "11_12_GG_INF", "11 A 12 ANOS": "11_12_GG_INF", "11 A 12": "11_12_GG_INF", "11-12 ANOS": "11_12_GG_INF",
            "G INFANTIL": "8_10_G_INF", "INFANTIL G": "8_10_G_INF", "8 A 10 ANOS": "8_10_G_INF", "8 A 10": "8_10_G_INF", "8-10 ANOS": "8_10_G_INF",
            "M INFANTIL": "6_7_M_INF", "INFANTIL M": "6_7_M_INF", "6 A 7 ANOS": "6_7_M_INF", "6 A 7": "6_7_M_INF", "6-7 ANOS": "6_7_M_INF",
            "P INFANTIL": "4_5_P_INF", "INFANTIL P": "4_5_P_INF", "4 A 5 ANOS": "4_5_P_INF", "4 A 5": "4_5_P_INF", "4-5 ANOS": "4_5_P_INF",
            "PP INFANTIL": "2_3_PP_INF", "INFANTIL PP": "2_3_PP_INF", "2 A 3 ANOS": "2_3_PP_INF", "2 A 3": "2_3_PP_INF", "2-3 ANOS": "2_3_PP_INF"
        }

        self.ordem_tamanhos = list(self.ordem_tamanhos_padrao)
        self.conversoes = dict(self.conversoes_padrao)
        self.preencher_padrao = tk.BooleanVar(value=True)
        self.tema_escuro = tk.BooleanVar(value=False)

        self.carregar_config()
        self._atualizar_termos_busca()

        self.pedidos_atuais = []          
        self.linhas_nao_reconhecidas = [] 
        self.historico = []               
        self.futuro = []                  
        self.arquivo_atual = None         

        self.paletas = {
            "claro": dict(bg="#F4F6F9", header="#1E293B", header_fg="white", panel="#FFFFFF", border="#E2E8F0", text="#0F172A"),
            "escuro": dict(bg="#111827", header="#0B1220", header_fg="#E5E7EB", panel="#1F2937", border="#374151", text="#E5E7EB"),
        }

        self.criar_interface()
        self._configurar_atalhos()
        self._atualizar_titulo()

        # Rotinas automáticas ao iniciar
        self.root.after(1000, self.verificar_e_instalar_ollama_automatico)
        self.root.after(2500, lambda: self.verificar_atualizacoes(silencioso=True))

    def carregar_config(self):
        if os.path.exists(CONFIG_PATH):
            try:
                with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                    dados = json.load(f)
                self.ordem_tamanhos = dados.get("ordem_tamanhos", self.ordem_tamanhos_padrao)
                self.conversoes = dados.get("conversoes", self.conversoes_padrao)
                self.preencher_padrao.set(dados.get("preencher_padrao", True))
            except Exception:
                pass 

    def salvar_config(self):
        try:
            with open(CONFIG_PATH, "w", encoding="utf-8") as f:
                json.dump({
                    "ordem_tamanhos": self.ordem_tamanhos,
                    "conversoes": self.conversoes,
                    "preencher_padrao": self.preencher_padrao.get(),
                }, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def _atualizar_termos_busca(self):
        todos_termos = list(self.ordem_tamanhos) + list(self.conversoes.keys())
        self.termos_busca = sorted(list(set(todos_termos)), key=len, reverse=True)

    def set_status(self, msg, progresso=None):
        timestamp = datetime.now().strftime('%H:%M:%S')
        self.status_var.set(f"{timestamp}  |  {msg}")
        if progresso is not None:
            try:
                self.progress_bar['value'] = progresso
            except Exception:
                pass

    # --- VERIFICAÇÃO E ATUALIZAÇÃO VIA GITHUB ---
    def verificar_atualizacoes(self, silencioso=True):
        def _checar():
            try:
                req = urllib.request.Request(URL_VERSAO_REMOTE, headers={'User-Agent': 'Mozilla/5.0'})
                with urllib.request.urlopen(req, timeout=5) as response:
                    dados = json.loads(response.read().decode('utf-8'))
                
                versao_remota = dados.get("version")
                url_download = dados.get("url")
                changelog = dados.get("changelog", "Melhorias e correções gerais.")

                if versao_remota and versao_remota != __version__:
                    msg = f"Uma nova versão ({versao_remota}) está disponível no GitHub!\n\nNovidades:\n{changelog}\n\nDeseja atualizar agora?"
                    if messagebox.askyesno("Atualização Disponível", msg, parent=self.root):
                        self.root.after(0, lambda: self._executar_atualizacao(url_download))
                elif not silencioso:
                    messagebox.showinfo("Atualizações", "Seu programa já está na versão mais recente!", parent=self.root)
            except Exception as e:
                if not silencioso:
                    messagebox.showerror("Erro", f"Não foi possível verificar atualizações:\n{e}", parent=self.root)

        threading.Thread(target=_checar, daemon=True).start()

    def _executar_atualizacao(self, url_download):
        try:
            self.set_status("A baixar atualização do GitHub...", 10)
            caminho_atual = os.path.abspath(sys.argv[0])
            caminho_temp = caminho_atual + ".tmp"

            def reporthook(blocknum, blocksize, totalsize):
                if totalsize > 0:
                    p = int(blocknum * blocksize * 80 / totalsize) + 10
                    if p > 90: p = 90
                    self.root.after(0, lambda: self.set_status(f"A baixar atualização... {p}%", p))

            urllib.request.urlretrieve(url_download, caminho_temp, reporthook=reporthook)
            self.set_status("Atualização baixada com sucesso!", 100)

            if caminho_atual.endswith(".py"):
                os.replace(caminho_temp, caminho_atual)
                messagebox.showinfo("Sucesso", "Programa atualizado com sucesso! O aplicativo será reiniciado.")
                os.execl(sys.executable, sys.executable, *sys.argv)
            else:
                caminho_bat = os.path.join(os.path.dirname(caminho_atual), "atualizar.bat")
                with open(caminho_bat, "w", encoding="utf-8") as f:
                    f.write(f'@echo off\n')
                    f.write(f'timeout /t 2 /nobreak > nul\n')
                    f.write(f'move /y "{caminho_temp}" "{caminho_atual}"\n')
                    f.write(f'start "" "{caminho_atual}"\n')
                    f.write(f'del "%~f0"\n')
                
                messagebox.showinfo("Atualizando", "O aplicativo será fechado para aplicar a atualização.")
                subprocess.Popen([caminho_bat], shell=True)
                sys.exit(0)
        except Exception as e:
            self.set_status("Falha na atualização.", 0)
            messagebox.showerror("Erro", f"Falha ao atualizar o programa:\n{e}")

    # --- VERIFICAÇÃO E INSTALAÇÃO AUTOMÁTICA DO OLLAMA ---
    def verificar_e_instalar_ollama_automatico(self):
        caminho_ollama_exe = os.path.expanduser(r"~\AppData\Local\Programs\Ollama\ollama.exe")
        instalado = os.path.exists(caminho_ollama_exe)
        if not instalado:
            try:
                res = subprocess.run(["ollama", "--version"], capture_output=True, text=True)
                if res.returncode == 0:
                    instalado = True
            except Exception:
                pass

        if not instalado:
            resposta = messagebox.askyesno(
                "Ollama não detectado",
                "O motor de Inteligência Artificial Local (Ollama) não foi encontrado.\n\n"
                "Deseja que o programa faça o download e a instalação automática agora?",
                parent=self.root
            )
            if resposta:
                self.set_status("A descarregar o instalador do Ollama...", 5)
                threading.Thread(target=self._executar_download_instalacao_ollama, daemon=True).start()
        else:
            self.set_status("Ollama detectado e pronto para uso local.", 100)

    def _executar_download_instalacao_ollama(self):
        try:
            url_installer = "https://ollama.com/download/OllamaSetup.exe"
            caminho_temp = os.path.join(os.environ.get("TEMP", "C:\\Temp"), "OllamaSetup.exe")
            caminho_ollama_exe = os.path.expanduser(r"~\AppData\Local\Programs\Ollama\ollama.exe")
            
            def reporthook_ollama(blocknum, blocksize, totalsize):
                if totalsize > 0:
                    p = int(blocknum * blocksize * 40 / totalsize) + 5
                    self.root.after(0, lambda: self.set_status(f"A baixar o Ollama... {p}%", p))

            self.root.after(0, lambda: self.set_status("A baixar o Ollama...", 10))
            urllib.request.urlretrieve(url_installer, caminho_temp, reporthook=reporthook_ollama)
            
            self.root.after(0, lambda: self.set_status("A executar o instalador do Ollama...", 50))
            subprocess.run([caminho_temp, "/VERYSILENT", "/SUPPRESSMSGBOXES", "/NORESTART"], check=True)
            
            self.root.after(0, lambda: self.set_status("A baixar modelo 'llama3'...", 75))
            subprocess.run([caminho_ollama_exe, "pull", "llama3"], check=True)
            try:
                self.root.after(0, lambda: self.set_status("A baixar modelo de visão 'llava'...", 90))
                subprocess.run([caminho_ollama_exe, "pull", "llava"], check=True)
            except Exception:
                pass
            
            self.root.after(0, lambda: messagebox.showinfo("Sucesso", "Ollama e os modelos de IA foram instalados com sucesso!"))
            self.root.after(0, lambda: self.set_status("IA Local pronta para uso.", 100))
        except Exception as e:
            err = str(e)
            self.root.after(0, lambda: messagebox.showerror("Erro na instalação", f"Não foi possível concluir a instalação:\n{err}"))
            self.root.after(0, lambda: self.set_status("Falha na instalação automática do Ollama.", 0))

    def criar_interface(self):
        estilo = ttk.Style()
        estilo.theme_use("clam")
        self._aplicar_estilo(estilo)

        menubar = tk.Menu(self.root)
        
        menu_arquivo = tk.Menu(menubar, tearoff=0)
        menu_arquivo.add_command(label="Abrir lista de um .txt...", command=self.abrir_txt)
        menu_arquivo.add_command(label="Colar da área de transferência", command=self.colar_da_area_transferencia, accelerator="Ctrl+V")
        menu_arquivo.add_separator()
        menu_arquivo.add_command(label="🔄 Verificar atualizações (GitHub)", command=lambda: self.verificar_atualizacoes(silencioso=False))
        menu_arquivo.add_separator()
        menu_arquivo.add_command(label="🖼️ Ler Imagem com IA Local (Ollama)...", command=self.ler_imagem_com_ollama)
        menu_arquivo.add_command(label="💻 Ler Imagem Offline (Tesseract)", command=self.ler_imagem_ocr_offline)
        menu_arquivo.add_separator()
        menu_arquivo.add_command(label="Exportar Excel (.xlsx)", command=self.salvar_em_xlsx)
        menu_arquivo.add_command(label="Exportar CSV (.csv)", command=self.salvar_em_csv)
        menu_arquivo.add_separator()
        menu_arquivo.add_command(label="Sair", command=self.root.quit)
        menubar.add_cascade(label="Ficheiro", menu=menu_arquivo)

        menu_editar = tk.Menu(menubar, tearoff=0)
        menu_editar.add_command(label="Desfazer", command=self.desfazer, accelerator="Ctrl+Z")
        menu_editar.add_command(label="Refazer", command=self.refazer, accelerator="Ctrl+Y")
        menu_editar.add_separator()
        menu_editar.add_command(label="Adicionar linha manual", command=self.adicionar_linha_manual)
        menu_editar.add_command(label="Excluir linha(s) selecionada(s)", command=self.excluir_linhas_selecionadas, accelerator="Delete")
        menubar.add_cascade(label="Editar", menu=menu_editar)

        menu_config = tk.Menu(menubar, tearoff=0)
        menu_config.add_command(label="Gerenciar tamanhos e conversões...", command=self.abrir_editor_regras)
        menu_config.add_checkbutton(label="Preencher campos vazios (SEM TAMANHO/NÚMERO)", variable=self.preencher_padrao, command=self.salvar_config)
        menu_config.add_checkbutton(label="Tema escuro", variable=self.tema_escuro, command=self.alternar_tema)
        menu_config.add_separator()
        menu_config.add_command(label="Restaurar padrões de fábrica", command=self.restaurar_padroes)
        menubar.add_cascade(label="Configurações", menu=menu_config)

        self.root.config(menu=menubar)

        self.header_frame = tk.Frame(self.root, bg="#1E293B", pady=15)
        self.header_frame.pack(fill="x")
        self.lbl_titulo = tk.Label(self.header_frame, text="ORGANIZADOR INTELIGENTE DE UNIFORMES — TURBO", font=("Segoe UI", 16, "bold"), bg="#1E293B", fg="white")
        self.lbl_titulo.pack()
        self.lbl_subtitulo = tk.Label(self.header_frame, text=f"Versão {__version__} | 100% Local com Ollama (Texto + Visão) e Auto-Update", font=("Segoe UI", 9), bg="#1E293B", fg="#94A3B8")
        self.lbl_subtitulo.pack()

        self.frame_entrada = tk.LabelFrame(self.root, text=" Cole sua lista desorganizada abaixo (uma por linha ou em blocos): ", font=("Segoe UI", 10, "bold"), bg="#F4F6F9", fg="#334155", padx=15, pady=10)
        self.frame_entrada.pack(fill="x", padx=20, pady=(15, 5))

        frame_ocr_rapido = tk.Frame(self.frame_entrada, bg="#F4F6F9")
        frame_ocr_rapido.pack(fill="x", pady=(0, 6))
        
        tk.Button(frame_ocr_rapido, text="🖼️ Ler Imagem (IA Local)", command=self.ler_imagem_com_ollama, bg="#7C3AED", fg="white", font=("Segoe UI", 9, "bold"), bd=0, cursor="hand2", padx=8, pady=4).pack(side="left", padx=(0, 4))
        tk.Button(frame_ocr_rapido, text="📋🖼️ Colar Imagem (IA Local)", command=self.colar_imagem_com_ollama, bg="#7C3AED", fg="white", font=("Segoe UI", 9, "bold"), bd=0, cursor="hand2", padx=8, pady=4).pack(side="left", padx=(0, 4))
        tk.Button(frame_ocr_rapido, text="💻 Ler Offline (Tesseract)", command=self.ler_imagem_ocr_offline, bg="#475569", fg="white", font=("Segoe UI", 9, "bold"), bd=0, cursor="hand2", padx=8, pady=4).pack(side="left")

        self.text_area = scrolledtext.ScrolledText(self.frame_entrada, height=7, font=("Segoe UI", 10), bd=1, relief="solid", undo=True)
        self.text_area.pack(fill="both", expand=True)

        frame_botoes_topo = tk.Frame(self.root, bg="#F4F6F9")
        frame_botoes_topo.pack(fill="x", padx=20, pady=5)

        self.btn_gerar = tk.Button(frame_botoes_topo, text="✨ SEPARAR E ORGANIZAR (Modo Normal)", command=self.processar_texto, bg="#2563EB", fg="white", font=("Segoe UI", 10, "bold"), activebackground="#1D4ED8", activeforeground="white", bd=0, cursor="hand2", padx=8, pady=8)
        self.btn_gerar.pack(side="left", padx=(0, 5), ipadx=5, fill="x", expand=True)

        self.btn_gerar_ia_local = tk.Button(frame_botoes_topo, text="🏠 ORGANIZAR COM IA LOCAL (Ollama)", command=self.processar_texto_com_ia_local_thread, bg="#059669", fg="white", font=("Segoe UI", 10, "bold"), activebackground="#047857", activeforeground="white", bd=0, cursor="hand2", padx=8, pady=8)
        self.btn_gerar_ia_local.pack(side="left", padx=(0, 5), ipadx=5, fill="x", expand=True)

        btn_limpar_tudo = tk.Button(frame_botoes_topo, text="🗑️ Limpar tudo", command=self.limpar_tudo, bg="#64748B", fg="white", font=("Segoe UI", 10, "bold"), bd=0, cursor="hand2", padx=8, pady=8)
        btn_limpar_tudo.pack(side="left", ipadx=5)

        frame_busca = tk.Frame(self.root, bg="#F4F6F9")
        frame_busca.pack(fill="x", padx=20, pady=(0, 5))

        tk.Label(frame_busca, text="🔎 Filtrar tabela:", bg="#F4F6F9", font=("Segoe UI", 9)).pack(side="left")
        self.var_filtro = tk.StringVar()
        self.var_filtro.trace_add("write", lambda *a: self.aplicar_filtro())
        entrada_filtro = tk.Entry(frame_busca, textvariable=self.var_filtro, font=("Segoe UI", 9), width=30)
        entrada_filtro.pack(side="left", padx=5)

        self.lbl_resumo = tk.Label(frame_busca, text="Nenhum pedido gerado ainda.", bg="#F4F6F9", font=("Segoe UI", 9, "bold"), fg="#334155")
        self.lbl_resumo.pack(side="right")

        self.frame_tabela = tk.Frame(self.root, bg="#E2E8F0", bd=1, relief="solid")
        self.frame_tabela.pack(fill="both", expand=True, padx=20, pady=5)

        scroll_y = ttk.Scrollbar(self.frame_tabela)
        scroll_y.pack(side="right", fill="y")
        scroll_x = ttk.Scrollbar(self.frame_tabela, orient="horizontal")
        scroll_x.pack(side="bottom", fill="x")

        colunas = ("NOME", "TAMANHO DE CAMISA", "NÚMERO", "TAMANHO DE CALÇÃO")
        self.tree = ttk.Treeview(self.frame_tabela, columns=colunas, show="headings", yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set, selectmode="extended")

        for col in colunas:
            self.tree.heading(col, text=col, command=lambda c=col: self.ordenar_por_coluna(c))
            self.tree.column(col, width=250, anchor="center")

        self.tree.pack(fill="both", expand=True)
        scroll_y.config(command=self.tree.yview)
        scroll_x.config(command=self.tree.xview)

        self.tree.tag_configure("duplicado", background="#FEE2E2")

        self.tree.bind("<Double-1>", self.editar_celula)
        self.tree.bind("<Delete>", lambda e: self.excluir_linhas_selecionadas())

        self._ordenacao_reversa = {}

        self.frame_avisos = tk.LabelFrame(self.root, text=" ⚠️ Linhas que não foram reconhecidas (revise manualmente): ", font=("Segoe UI", 9, "bold"), bg="#FFF7ED", fg="#9A3412", padx=10, pady=5)
        self.txt_avisos = tk.Text(self.frame_avisos, height=3, font=("Segoe UI", 9), bd=0, bg="#FFF7ED", fg="#7C2D12", wrap="word")
        self.txt_avisos.pack(fill="both", expand=True)
        self.txt_avisos.config(state="disabled")

        frame_acoes = tk.Frame(self.root, bg="#F4F6F9")
        frame_acoes.pack(fill="x", padx=20, pady=5)

        btn_copiar = tk.Button(frame_acoes, text="📋 COPIAR PARA COLAR NO EXCEL", command=self.copiar_para_excel, bg="#16A34A", fg="white", font=("Segoe UI", 10, "bold"), bd=0, cursor="hand2", padx=15, pady=8)
        btn_copiar.pack(side="left", padx=5, expand=True, fill="x")

        btn_salvar_xlsx = tk.Button(frame_acoes, text="📊 SALVAR COMO EXCEL (.XLSX)", command=self.salvar_em_xlsx, bg="#0EA5E9", fg="white", font=("Segoe UI", 10, "bold"), bd=0, cursor="hand2", padx=15, pady=8)
        btn_salvar_xlsx.pack(side="left", padx=5, expand=True, fill="x")

        btn_salvar = tk.Button(frame_acoes, text="💾 SALVAR FICHEIRO (.CSV)", command=self.salvar_em_csv, bg="#D97706", fg="white", font=("Segoe UI", 10, "bold"), bd=0, cursor="hand2", padx=15, pady=8)
        btn_salvar.pack(side="right", padx=5, expand=True, fill="x")

        # --- RODAPÉ PROFISSIONAL COM BARRA DE PROGRESSO DETERMINADA ---
        self.footer_container = tk.Frame(self.root, bg="#1E293B", pady=10, padx=15)
        self.footer_container.pack(fill="x", side="bottom")

        frame_status_bar = tk.Frame(self.footer_container, bg="#1E293B")
        frame_status_bar.pack(fill="x", pady=(0, 6))

        self.status_var = tk.StringVar(value="🟢 Sistema pronto para processamento 100% Local.")
        self.lbl_status_pro = tk.Label(frame_status_bar, textvariable=self.status_var, font=("Segoe UI", 10, "bold"), bg="#1E293B", fg="#F8FAFC", anchor="w")
        self.lbl_status_pro.pack(side="left", fill="x", expand=True)

        self.progress_bar = ttk.Progressbar(frame_status_bar, orient="horizontal", length=220, mode="determinate", maximum=100)
        self.progress_bar.pack(side="right", padx=(10, 0))

        lbl_assinatura = tk.Label(self.footer_container, text="Desenvolvido por Douglas Oliveira  |  Contato: getyourwish10@gmail.com", font=("Segoe UI", 9), bg="#1E293B", fg="#94A3B8")
        lbl_assinatura.pack(anchor="center")

    def _aplicar_estilo(self, estilo):
        estilo.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"), background="#1E293B", foreground="white")
        estilo.configure("Treeview", font=("Segoe UI", 10), rowheight=25)

    def _configurar_atalhos(self):
        self.root.bind("<Control-Return>", lambda e: self.processar_texto())
        self.root.bind("<Control-z>", lambda e: self.desfazer())
        self.root.bind("<Control-y>", lambda e: self.refazer())
        self.root.bind("<Control-v>", lambda e: None)

    def _atualizar_titulo(self):
        extra = f" — {os.path.basename(self.arquivo_atual)}" if self.arquivo_atual else ""
        self.root.title(f"Organizador Inteligente de Pedidos - by Douglas Oliveira{extra}")

    def alternar_tema(self):
        p = self.paletas["escuro"] if self.tema_escuro.get() else self.paletas["claro"]
        self.root.configure(bg=p["bg"])
        self.header_frame.configure(bg=p["header"])
        self.lbl_titulo.configure(bg=p["header"], fg=p["header_fg"])
        self.lbl_subtitulo.configure(bg=p["header"])
        self.frame_entrada.configure(bg=p["bg"], fg=p["text"])
        self.text_area.configure(bg=p["panel"], fg=p["text"], insertbackground=p["text"])
        self.set_status("Tema atualizado.")

    def _salvar_estado_para_undo(self):
        self.historico.append(copy.deepcopy(self.pedidos_atuais))
        if len(self.historico) > HISTORICO_MAXIMO:
            self.historico.pop(0)
        self.futuro.clear()

    def desfazer(self):
        if not self.historico:
            self.set_status("Nada para desfazer.")
            return
        self.futuro.append(copy.deepcopy(self.pedidos_atuais))
        self.pedidos_atuais = self.historico.pop()
        self._redesenhar_tabela()
        self.set_status("Ação desfeita.")

    def refazer(self):
        if not self.futuro:
            self.set_status("Nada para refazer.")
            return
        self.historico.append(copy.deepcopy(self.pedidos_atuais))
        self.pedidos_atuais = self.futuro.pop()
        self._redesenhar_tabela()
        self.set_status("Ação refeita.")

    def obter_peso_ordenacao(self, item):
        tam_camisa = item['TAMANHO DE CAMISA']
        tam_calcao = item['TAMANHO DE CALÇÃO']

        if tam_camisa in self.ordem_tamanhos:
            return (self.ordem_tamanhos.index(tam_camisa), 0)
        if tam_calcao in self.ordem_tamanhos:
            return (self.ordem_tamanhos.index(tam_calcao), 1)
        return (999, 99)

    def ordenar_por_coluna(self, coluna):
        chave_map = {
            "NOME": lambda p: p["NOME"],
            "TAMANHO DE CAMISA": lambda p: (self.ordem_tamanhos.index(p["TAMANHO DE CAMISA"]) if p["TAMANHO DE CAMISA"] in self.ordem_tamanhos else 999),
            "NÚMERO": lambda p: int(p["NÚMERO"]) if str(p["NÚMERO"]).isdigit() else (0 if str(p["NÚMERO"]).upper() in ["PI", "π"] else 9999),
            "TAMANHO DE CALÇÃO": lambda p: (self.ordem_tamanhos.index(p["TAMANHO DE CALÇÃO"]) if p["TAMANHO DE CALÇÃO"] in self.ordem_tamanhos else 999),
        }
        reverso = self._ordenacao_reversa.get(coluna, False)
        self.pedidos_atuais.sort(key=chave_map[coluna], reverse=reverso)
        self._ordenacao_reversa[coluna] = not reverso
        self._redesenhar_tabela(agrupar=False)
        self.set_status(f"Tabela ordenada por {coluna}.")

    # --- PROCESSAMENTO PRINCIPAL ---
    def processar_texto(self):
        texto_bruto = self.text_area.get("1.0", tk.END).strip()
        if not texto_bruto:
            messagebox.showwarning("Aviso", "A caixa de texto está vazia!")
            return

        self._salvar_estado_para_undo()
        self.set_status("A processar dados (Modo Normal)...", 30)

        texto_bruto = re.sub(r'Número Tabela camisa da equipe:?', '', texto_bruto, flags=re.IGNORECASE)

        padroes_ficha = [r'tamanho[\/\s]*[:]', r'nome[\/\s]*[:]', r'número[\/\s]*[:]', r'numero[\/\s]*[:]']
        eh_ficha = any(re.search(p, texto_bruto, re.IGNORECASE) for p in padroes_ficha) or ('___' in texto_bruto)

        pedidos_processados = []
        linhas_nao_reconhecidas = []

        if eh_ficha:
            if re.search(r'[_]{3,}|[-]{3,}', texto_bruto):
                blocos = re.split(r'[_]{3,}|[-]{3,}', texto_bruto)
                blocos = [b.strip() for b in blocos if b.strip()]
            else:
                linhas = [l.strip() for l in texto_bruto.split('\n') if l.strip()]
                blocos = []
                bloco_atual = []
                for linha in linhas:
                    tem_tam = bool(re.search(r'^\s*(Tamanho|Tam)\b', linha, re.IGNORECASE))
                    if tem_tam and any(re.search(r'(Nome|Número|Numero|\d+)', l, re.IGNORECASE) for l in bloco_atual):
                        blocos.append('\n'.join(bloco_atual))
                        bloco_atual = [linha]
                    else:
                        bloco_atual.append(linha)
                if bloco_atual:
                    blocos.append('\n'.join(bloco_atual))
                blocos = [b.strip() for b in blocos if b.strip()]

            for bloco in blocos:
                tamanho_camisa_linha = ""
                tamanho_calcao_linha = ""
                nome = ""
                numero = ""
                is_conjunto = "CONJUNTO" in bloco.upper()

                linhas_bloco = bloco.split('\n')
                for linha in linhas_bloco:
                    linha_orig = linha.strip()
                    if not linha_orig:
                        continue
                    linha_up = linha_orig.upper()

                    m_tam = re.search(r'(?:TAMANHO|TAM)[\/\s]*[:]?\s*([A-Za-z0-9_]+)', linha_up)
                    if m_tam:
                        tamanho_camisa_linha = m_tam.group(1).upper()
                        continue

                    m_num = re.search(r'N[ÚU]MERO(?:[^\d]*)(?:\s+DA\s+CAMISA)?[\/\s]*[:]?\s*([0-9πPI]+)', linha_up, re.IGNORECASE)
                    if m_num:
                        numero = m_num.group(1).upper()
                        continue

                    if "NOME" in linha_up:
                        partes = re.split(r'[:]', linha_orig, 1)
                        if len(partes) > 1:
                            val = partes[1].strip()
                            if val:
                                nome = val
                        continue

                    if not tamanho_camisa_linha:
                        for term in self.termos_busca:
                            if linha_up == term or linha_up.startswith(term):
                                tamanho_camisa_linha = term
                                break

                    if not nome and not re.search(r'^\d+$', linha_orig) and not re.search(r'^(PI|π)$', linha_up, re.IGNORECASE):
                        if not any(k in linha_up for k in ["TAMANHO", "NUMERO", "NÚMERO", "CONJUNTO"]):
                            nome = linha_orig

                nome = re.sub(r'^(nome(\s+na\s+camisa|\s+da\s+camisa)?)[\/\s]*[:]?\s*', '', nome, flags=re.IGNORECASE).strip()
                nome = re.sub(r'[\xa0\t]+', ' ', nome).strip()

                tamanho_camisa_final = self.conversoes.get(tamanho_camisa_linha, tamanho_camisa_linha)
                tamanho_calcao_final = tamanho_camisa_final if is_conjunto else ""

                if nome or numero or tamanho_camisa_final:
                    pedidos_processados.append({
                        'NOME': nome.upper(),
                        'TAMANHO DE CAMISA': tamanho_camisa_final,
                        'NÚMERO': numero,
                        'TAMANHO DE CALÇÃO': tamanho_calcao_final
                    })
                else:
                    if bloco.strip():
                        linhas_nao_reconhecidas.append(bloco)
        else:
            linhas = texto_bruto.split('\n')
            contexto_camisa = ""
            contexto_calcao = ""

            for linha_original in linhas:
                linha_original = linha_original.strip()
                if not linha_original:
                    continue

                linha_temp = linha_original.upper()
                linha_temp = re.sub(r'^\s*\d+\s*[\-\.\)]\s*', '', linha_temp)
                linha_temp = re.sub(r'\bS\s*/\s*N\b', ' ', linha_temp)

                if re.search(r'(PATROCÍNIO|NUMERAÇÃO|JOGO|MATERIAL|PENHAROL)', linha_temp):
                    continue
                if re.match(r'^\*?\s*SUB\s*\d+', linha_temp):
                    continue

                quantidade = 1
                is_kit_line = "KIT" in linha_temp
                is_conjunto = "CONJUNTO" in linha_temp

                if not is_kit_line:
                    match_qtd = re.search(r'^\s*\*?\s*(\d+)\s*(?:X|UN|PCT)\b', linha_temp)
                    if match_qtd:
                        quantidade = int(match_qtd.group(1))
                        linha_temp = re.sub(r'^\s*\*?\s*\d+\s*(?:X|UN|PCT)\b', ' ', linha_temp, count=1)
                else:
                    linha_temp = re.sub(r'^\s*\*?\s*\d+\s', ' ', linha_temp, count=1)

                if "AVULSA" in linha_temp or "AVULSO" in linha_temp:
                    contexto_calcao = ""

                tamanho_camisa_linha = ""
                tamanho_calcao_linha = ""

                for tam in self.termos_busca:
                    prefixo_calcao = r'(TAMANHO\s+(?:D[AEO]\s+)?(?:CAL[CÇ][AÃ]O|CALS[AÃ]O|CAUCAO|SHORT[ES]?|BERMUDA)|(?:CAL[CÇ][AÃ]O|CALS[AÃ]O|CAUCAO|SHORT[ES]?|BERMUDA)\s+TAMANHO|CAL[CÇ][AÃ]O|CALS[AÃ]O|CAUCAO|SHORT[ES]?|BERMUDA)'
                    padrao_calcao = prefixo_calcao + r'[\s\-\:\/\;\.\,]*' + r'(?<![A-Z0-9À-Ÿ])' + re.escape(tam) + r'(?![A-Z0-9À-Ÿ])'
                    if re.search(padrao_calcao, linha_temp):
                        tamanho_calcao_linha = tam
                        linha_temp = re.sub(padrao_calcao, ' ', linha_temp, count=1)
                        break

                for tam in self.termos_busca:
                    prefixo_camisa = r'(TAMANHO\s+(?:D[AEO]\s+)?(?:CAMISA|CAMISETA|CAMIZA|BLUSA)|(?:CAMISA|CAMISETA|CAMIZA|BLUSA)\s+TAMANHO|TAMANHO|TAM\.|TAM|T|CAMISA|CAMISETA|CAMIZA|BLUSA|BABY\s*LOOK)'
                    padrao_tam = prefixo_camisa + r'[\s\-\:\/\;\.\,]*' + r'(?<![A-Z0-9À-Ÿ])' + re.escape(tam) + r'(?![A-Z0-9À-Ÿ])'
                    if re.search(padrao_tam, linha_temp):
                        tamanho_camisa_linha = tam
                        linha_temp = re.sub(padrao_tam, ' ', linha_temp, count=1)
                        break

                tamanhos_soltos = []
                for tam in self.termos_busca:
                    padrao = r'(?<![A-Z0-9À-Ÿ])' + re.escape(tam) + r'(?![A-Z0-9À-Ÿ])'
                    while re.search(padrao, linha_temp):
                        tamanhos_soltos.append(tam)
                        linha_temp = re.sub(padrao, ' ', linha_temp, count=1)

                if not tamanho_camisa_linha and tamanhos_soltos:
                    tamanho_camisa_linha = tamanhos_soltos.pop(0)
                if not tamanho_calcao_linha and tamanhos_soltos:
                    tamanho_calcao_linha = tamanhos_soltos.pop(0)

                if tamanho_camisa_linha or tamanho_calcao_linha:
                    if tamanho_camisa_linha:
                        contexto_camisa = tamanho_camisa_linha
                    if tamanho_calcao_linha:
                        contexto_calcao = tamanho_calcao_linha

                tamanho_camisa_final = tamanho_camisa_linha if tamanho_camisa_linha else contexto_camisa
                tamanho_calcao_final = tamanho_calcao_linha if tamanho_calcao_linha else contexto_calcao

                if is_conjunto and not tamanho_calcao_final and tamanho_camisa_final:
                    tamanho_calcao_final = tamanho_camisa_final

                numero = ""
                match_num_exp = re.search(r'(N[UÚ]MERO|NUM\.|NUM|N[º°O\.]|#|N)[\s\-\:\;]*(\d+|PI|π)', linha_temp, re.IGNORECASE)
                if match_num_exp:
                    numero = match_num_exp.group(2).upper()
                    linha_temp = linha_temp.replace(match_num_exp.group(0), ' ')
                else:
                    match_barra_traco = re.search(r'[\/\-]\s*(\d+|PI|π)', linha_temp, re.IGNORECASE)
                    if match_barra_traco:
                        numero = match_barra_traco.group(1).upper()
                        linha_temp = linha_temp.replace(match_barra_traco.group(0), ' ')
                    else:
                        match_solto = re.search(r'\b(\d{1,3}|PI|π)\b', linha_temp, re.IGNORECASE)
                        if match_solto:
                            numero = match_solto.group(1).upper()
                            linha_temp = re.sub(r'\b' + re.escape(match_solto.group(1)) + r'\b', ' ', linha_temp, count=1, flags=re.IGNORECASE)

                palavras_chave_remover = [
                    r'\bESCREVER\b', r'\bCOM\b', r'\bSEM\b', r'\bNOME\b', r'\bAVULSA\b', r'\bAVULSO\b', r'\bAVULSAS\b',
                    r'\bCAMISETAS?\b', r'\bCAMISAS?\b', r'\bCAMIZAS?\b', r'\bBLUSAS?\b', r'\bBABY\s*LOOK\b',
                    r'\bADULTO\b', r'\bINFANTIL\b', r'\bKITS?\b', r'\bGOLEIRO\b', r'\bDE\b', r'\bE\b',
                    r'\bCONJUNTO\b', r'\bCONJUNTOS\b',
                    r'\bCAL[CÇ][AÃ]O\b', r'\bCALS[AÃ]O\b', r'\bCAUCAO\b', r'\bSHORT[ES]?\b', r'\bBERMUDA\b',
                    r'\bTAMANHO\b', r'\bTAM\.\b', r'\bTAM\b', r'\bT\b', r'\bN[UÚ]MERO\b', r'\bNUM\.\b', r'\bNUM\b', r'\bN[º°O\.]\b', r'#', r'\bN\b'
                ]
                for padrao_remocao in palavras_chave_remover:
                    linha_temp = re.sub(padrao_remocao, ' ', linha_temp)

                linha_temp = re.sub(r'^[\-\/\:\,\.\_\–\—\*\s]+', '', linha_temp)
                linha_temp = re.sub(r'[\-\/\:\,\.\_\–\—\*\s]+$', '', linha_temp)
                linha_temp = re.sub(r"[^A-Z0-9À-Ÿ'\-\s]", ' ', linha_temp)

                nome = re.sub(r'\s+', ' ', linha_temp).strip()
                nome = re.sub(r"(^|\s)[\-']+(\s|$)", r'\1\2', nome).strip()

                if re.match(r'^\d+$', nome):
                    nome = ""

                is_order = False
                if is_kit_line:
                    if numero or nome:
                        is_order = True
                else:
                    if (numero != "" or nome != "") or ("CAMISA" in linha_original.upper() or "BABY" in linha_original.upper()):
                        if "AVULSAS" in linha_original.upper() and not tamanho_camisa_linha:
                            is_order = False
                        else:
                            is_order = True

                if is_order:
                    if not numero or "SEM" in numero:
                        numero = ""
                    if not tamanho_camisa_final or "SEM" in tamanho_camisa_final:
                        tamanho_camisa_final_convertido = ""
                    else:
                        tamanho_camisa_final_convertido = self.conversoes.get(tamanho_camisa_final, tamanho_camisa_final)
                    if not tamanho_calcao_final or "SEM" in tamanho_calcao_final:
                        tamanho_calcao_final_convertido = ""
                    else:
                        tamanho_calcao_final_convertido = self.conversoes.get(tamanho_calcao_final, tamanho_calcao_final)

                    for _ in range(quantidade):
                        pedidos_processados.append({
                            'NOME': nome,
                            'TAMANHO DE CAMISA': tamanho_camisa_final_convertido,
                            'NÚMERO': numero,
                            'TAMANHO DE CALÇÃO': tamanho_calcao_final_convertido
                        })
                else:
                    if linha_original.strip():
                        linhas_nao_reconhecidas.append(linha_original)

        if self.preencher_padrao.get():
            for p in pedidos_processados:
                if not p['TAMANHO DE CAMISA']:
                    p['TAMANHO DE CAMISA'] = "SEM TAMANHO"
                if not p['NÚMERO']:
                    p['NÚMERO'] = "SEM NÚMERO"
                if not p['TAMANHO DE CALÇÃO']:
                    p['TAMANHO DE CALÇÃO'] = "SEM CALÇÃO"

        pedidos_processados.sort(key=self.obter_peso_ordenacao)

        self.pedidos_atuais = pedidos_processados
        self.linhas_nao_reconhecidas = linhas_nao_reconhecidas
        self._redesenhar_tabela()
        self._atualizar_avisos()
        self.set_status(f"Sucesso! {len(pedidos_processados)} pedido(s) organizado(s). {len(linhas_nao_reconhecidas)} linha(s) não reconhecida(s).", 100)

    # --- PROCESSAMENTO COM IA LOCAL (OLLAMA TEXTO) ---
    def processar_texto_com_ia_local_thread(self):
        texto_bruto = self.text_area.get("1.0", tk.END).strip()
        if not texto_bruto:
            messagebox.showwarning("Aviso", "A caixa de texto está vazia!")
            return
        if not OLLAMA_DISPONIVEL:
            messagebox.showerror("Erro", "A biblioteca 'ollama' não está instalada.\n\nInstale executando no terminal:\npip install ollama")
            return

        self.set_status("A IA Local (Ollama) está a processar os dados...", 40)
        threading.Thread(target=self._processar_texto_com_ia_local_exec, args=(texto_bruto,), daemon=True).start()

    def _processar_texto_com_ia_local_exec(self, texto_bruto):
        try:
            prompt = f"""
            És um assistente especialista em organizar listas de uniformes para gráficas.
            Analisa o texto desorganizado abaixo e converte-o numa lista estruturada.
            Cada linha deve conter exatamente 4 colunas separadas por ponto e vírgula (;):
            NOME; TAMANHO DE CAMISA; NÚMERO; TAMANHO DE CALÇÃO

            Regras estritas:
            - Se a linha contiver a palavra "CONJUNTO", o tamanho do calção deve ser igual ao tamanho da camisa.
            - Se o número for "π" ou "PI", mantém na coluna de número.
            - Se faltar o nome, número ou tamanho, preenche com "SEM NOME", "SEM TAMANHO", "SEM NÚMERO", "SEM CALÇÃO".
            - Converte tamanhos informais (ex: EXG para G2, G1 para GG).
            - Retorna APENAS as linhas no formato de texto estruturado com ponto e vírgula, sem explicações extras e sem blocos markdown.

            Texto bruto:
            {texto_bruto}
            """

            resposta = ollama.chat(model='llama3', messages=[{'role': 'user', 'content': prompt}])
            resposta_texto = resposta['message']['content'].replace("```csv", "").replace("```", "").strip()
            linhas_resposta = resposta_texto.splitlines()
            pedidos_processados = []

            for linha in linhas_resposta:
                linha = linha.strip()
                if not linha:
                    continue
                partes = [p.strip() for p in linha.split(';')]
                if len(partes) >= 4:
                    pedidos_processados.append({
                        'NOME': partes[0],
                        'TAMANHO DE CAMISA': partes[1],
                        'NÚMERO': partes[2],
                        'TAMANHO DE CALÇÃO': partes[3]
                    })

            if not pedidos_processados:
                self.root.after(0, lambda: messagebox.showwarning("Aviso", "A IA Local não retornou dados válidos."))
                self.root.after(0, lambda: self.set_status("Aviso: IA Local não retornou dados.", 0))
                return

            pedidos_processados.sort(key=self.obter_peso_ordenacao)

            def atualizar_interface_sucesso():
                self._salvar_estado_para_undo()
                self.pedidos_atuais = pedidos_processados
                self.linhas_nao_reconhecidas = []
                self._redesenhar_tabela()
                self._atualizar_avisos()
                self.set_status(f"Sucesso! {len(pedidos_processados)} pedido(s) organizados por IA Local.", 100)

            self.root.after(0, atualizar_interface_sucesso)

        except Exception as e:
            erro_msg = str(e)
            def mostrar_erro():
                messagebox.showerror("Erro na IA Local (Ollama)", f"Certifique-se de que o aplicativo Ollama está aberto no seu PC.\n\nDetalhes:\n{erro_msg}")
                self.set_status("Erro na IA Local.", 0)
            self.root.after(0, mostrar_erro)

    # --- OCR / LEITURA DE IMAGEM COM OLLAMA (LLAVA COM AUTO-DOWNLOAD) ---
    def ler_imagem_com_ollama(self):
        if not OLLAMA_DISPONIVEL:
            messagebox.showerror("Erro", "A biblioteca 'ollama' não está instalada.")
            return
        caminho = filedialog.askopenfilename(title="Selecione a imagem", filetypes=[("Imagens", "*.png *.jpg *.jpeg *.bmp *.webp")])
        if not caminho:
            return
        self._rodar_ocr_ollama_thread(caminho, os.path.basename(caminho))

    def colar_imagem_com_ollama(self):
        if not OLLAMA_DISPONIVEL:
            messagebox.showerror("Erro", "A biblioteca 'ollama' não está instalada.")
            return
        try:
            imagem = ImageGrab.grabclipboard()
        except Exception:
            imagem = None
        if imagem is None:
            messagebox.showinfo("Aviso", "Nenhuma imagem copiada na área de transferência.")
            return
        
        caminho_temp = os.path.join(os.environ.get("TEMP", "C:\\Temp"), "temp_clipboard_img.png")
        imagem.save(caminho_temp)
        self._rodar_ocr_ollama_thread(caminho_temp, "área de transferência")

    def _rodar_ocr_ollama_thread(self, caminho_imagem, origem):
        self.set_status(f"A ler imagem com IA Local ({origem})...", 20)
        threading.Thread(target=self._rodar_ocr_ollama_exec, args=(caminho_imagem,), daemon=True).start()

    def _rodar_ocr_ollama_exec(self, caminho_imagem):
        try:
            prompt = "Transcreve absolutamente todo o texto contido nesta imagem, mantendo a estrutura exata das linhas e informações dos pedidos."
            texto = ""
            caminho_ollama_exe = os.path.expanduser(r"~\AppData\Local\Programs\Ollama\ollama.exe")
            
            try:
                self.root.after(0, lambda: self.set_status("A processar imagem com IA Local (llava)...", 50))
                response = ollama.chat(
                    model='llava',
                    messages=[{'role': 'user', 'content': prompt, 'images': [caminho_imagem]}]
                )
                texto = response['message']['content']
            except Exception:
                self.root.after(0, lambda: self.set_status("A descarregar o modelo de visão 'llava'...", 30))
                if os.path.exists(caminho_ollama_exe):
                    subprocess.run([caminho_ollama_exe, "pull", "llava"], check=True)
                else:
                    subprocess.run(["ollama", "pull", "llava"], check=True)
                
                self.root.after(0, lambda: self.set_status("A processar imagem com IA Local (llava)...", 70))
                response = ollama.chat(
                    model='llava',
                    messages=[{'role': 'user', 'content': prompt, 'images': [caminho_imagem]}]
                )
                texto = response['message']['content']

            if texto:
                self.root.after(0, lambda: self.set_status("Leitura de imagem por IA concluída.", 100))
                self.root.after(0, lambda: self._mostrar_janela_revisao_ocr(texto))
            else:
                self.root.after(0, lambda: self.set_status("Aviso: Nenhum texto retornado pela IA.", 0))
        except Exception as e:
            err = str(e)
            self.root.after(0, lambda: self.set_status("Erro na leitura de imagem por IA.", 0))
            self.root.after(0, lambda: messagebox.showerror("Erro de IA Local (Visão)", f"Certifique-se de que o aplicativo Ollama está aberto.\n\nDetalhes:\n{err}"))

    # --- OCR OFFLINE (TESSERACT) ---
    def _mostrar_aviso_instalacao_tesseract(self):
        if messagebox.askyesno("Tesseract-OCR", "O Leitor Offline precisa do Tesseract-OCR instalado.\n\nDeseja abrir a página de download?"):
            webbrowser.open("https://github.com/UB-Mannheim/tesseract/wiki")

    def ler_imagem_ocr_offline(self):
        if not PYTESSERACT_DISPONIVEL:
            self._mostrar_aviso_instalacao_tesseract()
            return
        caminho = filedialog.askopenfilename(filetypes=[("Imagens", "*.png *.jpg *.jpeg")])
        if not caminho:
            return
        try:
            self.set_status("A executar OCR offline...", 50)
            texto = pytesseract.image_to_string(Image.open(caminho), lang='por')
            self.set_status("OCR offline concluído.", 100)
            self._mostrar_janela_revisao_ocr(texto)
        except Exception:
            self.set_status("Erro no OCR offline.", 0)
            self._mostrar_aviso_instalacao_tesseract()

    def _mostrar_janela_revisao_ocr(self, texto):
        janela = tk.Toplevel(self.root)
        janela.title("Revisar texto lido")
        janela.geometry("620x480")
        janela.transient(self.root)
        janela.grab_set()

        tk.Label(janela, text="Confira o texto lido:", font=("Segoe UI", 9, "bold")).pack(anchor="w", padx=15, pady=(15, 5))
        txt = scrolledtext.ScrolledText(janela, font=("Segoe UI", 10), undo=True)
        txt.pack(fill="both", expand=True, padx=15)
        txt.insert(tk.END, texto)

        def usar(subst=True):
            val = txt.get("1.0", tk.END).strip()
            if subst:
                self.text_area.delete("1.0", tk.END)
                self.text_area.insert(tk.END, val)
            else:
                self.text_area.insert(tk.END, "\n" + val)
            janela.destroy()

        f = tk.Frame(janela)
        f.pack(fill="x", padx=15, pady=15)
        tk.Button(f, text="✅ Usar (Substituir)", command=lambda: usar(True), bg="#2563EB", fg="white", font=("Segoe UI", 9, "bold")).pack(side="right", padx=5)
        tk.Button(f, text="➕ Adicionar", command=lambda: usar(False), bg="#16A34A", fg="white", font=("Segoe UI", 9, "bold")).pack(side="right", padx=5)

    # --- TABELA E INTERFACE AUXILIARES ---
    def _chave_grupo(self, p):
        if p['TAMANHO DE CAMISA'] in self.ordem_tamanhos:
            return p['TAMANHO DE CAMISA']
        return p['TAMANHO DE CALÇÃO']

    def _redesenhar_tabela(self, agrupar=True):
        for linha in self.tree.get_children():
            self.tree.delete(linha)

        filtro = self.var_filtro.get().strip().upper() if hasattr(self, "var_filtro") else ""
        contagem_nomes = {}
        for p in self.pedidos_atuais:
            nome = p['NOME'].strip().upper()
            if nome and nome != "SEM NOME":
                contagem_nomes[nome] = contagem_nomes.get(nome, 0) + 1

        grupo_atual = None
        for idx, p in enumerate(self.pedidos_atuais):
            if filtro:
                junto = " ".join(str(v) for v in p.values()).upper()
                if filtro not in junto:
                    continue

            if agrupar:
                chave_g = self._chave_grupo(p)
                if grupo_atual is not None and chave_g != grupo_atual:
                    self.tree.insert("", tk.END, values=("", "", "", ""))
                grupo_atual = chave_g

            tag = "duplicado" if contagem_nomes.get(p['NOME'].strip().upper(), 0) > 1 else ""
            self.tree.insert("", tk.END, iid=str(idx), values=(p['NOME'], p['TAMANHO DE CAMISA'], p['NÚMERO'], p['TAMANHO DE CALÇÃO']), tags=(tag,))

        self._atualizar_resumo()

    def aplicar_filtro(self):
        self._redesenhar_tabela(agrupar=not bool(self.var_filtro.get().strip()))

    def _atualizar_resumo(self):
        total = len(self.pedidos_atuais)
        if total == 0:
            self.lbl_resumo.config(text="Nenhum pedido gerado ainda.")
            return
        contagem = {}
        for p in self.pedidos_atuais:
            tam = p['TAMANHO DE CAMISA'] or p['TAMANHO DE CALÇÃO'] or "?"
            contagem[tam] = contagem.get(tam, 0) + 1
        resumo = "  |  ".join(f"{t}: {q}" for t, q in sorted(contagem.items()))
        self.lbl_resumo.config(text=f"Total: {total}  |  {resumo}")

    def _atualizar_avisos(self):
        if self.linhas_nao_reconhecidas:
            self.frame_avisos.pack(fill="x", padx=20, pady=5)
            self.txt_avisos.config(state="normal")
            self.txt_avisos.delete("1.0", tk.END)
            self.txt_avisos.insert(tk.END, "\n".join(f"• {l}" for l in self.linhas_nao_reconhecidas))
            self.txt_avisos.config(state="disabled")
        else:
            self.frame_avisos.pack_forget()

    def editar_celula(self, event):
        item_id = self.tree.identify_row(event.y)
        coluna_id = self.tree.identify_column(event.x)
        if not item_id or not coluna_id:
            return
        valores_atuais = self.tree.item(item_id, "values")
        if not any(valores_atuais):
            return

        col_index = int(coluna_id.replace("#", "")) - 1
        colunas = ("NOME", "TAMANHO DE CAMISA", "NÚMERO", "TAMANHO DE CALÇÃO")
        nome_coluna = colunas[col_index]

        x, y, w, h = self.tree.bbox(item_id, coluna_id)
        entry = tk.Entry(self.tree, font=("Segoe UI", 10))
        entry.insert(0, valores_atuais[col_index])
        entry.select_range(0, tk.END)
        entry.focus()
        entry.place(x=x, y=y, width=w, height=h)

        def salvar(e=None):
            val = entry.get().strip()
            entry.destroy()
            try:
                idx = int(item_id)
                self._salvar_estado_para_undo()
                self.pedidos_atuais[idx][nome_coluna] = val
                self._redesenhar_tabela()
                self.set_status(f"Célula atualizada: {nome_coluna} = {val}")
            except Exception:
                pass

        entry.bind("<Return>", salvar)
        entry.bind("<FocusOut>", salvar)
        entry.bind("<Escape>", lambda e: entry.destroy())

    def adicionar_linha_manual(self):
        self._salvar_estado_para_undo()
        self.pedidos_atuais.append({'NOME': '', 'TAMANHO DE CAMISA': '', 'NÚMERO': '', 'TAMANHO DE CALÇÃO': ''})
        self._redesenhar_tabela(agrupar=False)
        self.set_status("Linha em branco adicionada.")

    def excluir_linhas_selecionadas(self):
        selecionados = self.tree.selection()
        if not selecionados:
            return
        indices = sorted([int(i) for i in selecionados if i.isdigit()], reverse=True)
        self._salvar_estado_para_undo()
        for idx in indices:
            if 0 <= idx < len(self.pedidos_atuais):
                del self.pedidos_atuais[idx]
        self._redesenhar_tabela()
        self.set_status("Linha(s) excluída(s).")

    def limpar_tudo(self):
        if not messagebox.askyesno("Confirmar", "Limpar tudo?"):
            return
        self._salvar_estado_para_undo()
        self.text_area.delete("1.0", tk.END)
        self.pedidos_atuais = []
        self.linhas_nao_reconhecidas = []
        self._redesenhar_tabela()
        self._atualizar_avisos()
        self.set_status("Tudo limpo.", 0)

    def abrir_txt(self):
        caminho = filedialog.askopenfilename(filetypes=[("Texto", "*.txt"), ("Todos", "*.*")])
        if caminho:
            try:
                with open(caminho, "r", encoding="utf-8") as f:
                    self.text_area.delete("1.0", tk.END)
                    self.text_area.insert(tk.END, f.read())
                self.set_status(f"Ficheiro carregado: {os.path.basename(caminho)}")
            except Exception as e:
                messagebox.showerror("Erro", str(e))

    def colar_da_area_transferencia(self):
        try:
            self.text_area.insert(tk.INSERT, self.root.clipboard_get())
            self.set_status("Conteúdo colado.")
        except Exception:
            pass

    def _linhas_para_exportar(self):
        return [self.tree.item(l)['values'] for l in self.tree.get_children()]

    def copiar_para_excel(self):
        linhas = self._linhas_para_exportar()
        if not linhas:
            return
        txt = "\n".join(["\t".join(str(c) for c in v) for v in linhas])
        self.root.clipboard_clear()
        self.root.clipboard_append(txt)
        messagebox.showinfo("Sucesso", "Copiado! Cole no Excel (Ctrl+V).")
        self.set_status("Dados copiados para a área de transferência.")

    def salvar_em_csv(self):
        linhas = self._linhas_para_exportar()
        if not linhas:
            return
        caminho = filedialog.saveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")])
        if caminho:
            try:
                with open(caminho, mode='w', newline='', encoding='utf-8-sig') as f:
                    w = csv.writer(f, delimiter=';')
                    w.writerow(["NOME", "TAMANHO DE CAMISA", "NÚMERO", "TAMANHO DE CALÇÃO"])
                    for l in linhas:
                        w.writerow(l)
                messagebox.showinfo("Sucesso", "CSV salvo com sucesso!")
                self.set_status(f"Ficheiro CSV salvo: {os.path.basename(caminho)}")
            except Exception as e:
                messagebox.showerror("Erro", str(e))

    def salvar_em_xlsx(self):
        if not OPENPYXL_DISPONIVEL:
            messagebox.showerror("Erro", "Instale o openpyxl: pip install openpyxl")
            return
        linhas = self._linhas_para_exportar()
        if not linhas:
            return
        caminho = filedialog.saveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if caminho:
            try:
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.append(["NOME", "TAMANHO DE CAMISA", "NÚMERO", "TAMANHO DE CALÇÃO"])
                for l in linhas:
                    ws.append(list(l))
                wb.save(caminho)
                messagebox.showinfo("Sucesso", "Excel salvo com sucesso!")
                self.set_status(f"Planilha salva: {os.path.basename(caminho)}")
            except Exception as e:
                messagebox.showerror("Erro", str(e))

    def abrir_editor_regras(self):
        janela = tk.Toplevel(self.root)
        janela.title("Gerenciar Regras")
        janela.geometry("500+400")
        tk.Label(janela, text="Em desenvolvimento / Configurações avançadas.").pack(padx=20, pady=20)

    def restaurar_padroes(self):
        self.ordem_tamanhos = list(self.ordem_tamanhos_padrao)
        self.conversoes = dict(self.conversoes_padrao)
        self.salvar_config()
        messagebox.showinfo("Restaurado", "Padrões restaurados.")
        self.set_status("Padrões de fábrica restaurados.", 0)


if __name__ == "__main__":
    root = tk.Tk()
    app = AplicativoPedidosMagico(root)
    root.mainloop()