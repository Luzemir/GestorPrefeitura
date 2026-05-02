import sys
import os
import threading
import subprocess
import customtkinter as ctk
from tkinter import filedialog, messagebox
import pandas as pd
from PIL import Image

def get_app_dir():
    """ Retorna o diretório persistente onde o .exe está, ou a pasta root se em Python """
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

def resource_path(relative_path):
    """ Retorna o caminho absoluto do arquivo empacotado (assets, imgs) pelo PyInstaller. """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)

# Adiciona o diretorio atual ao path para poder importar o bot
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import nfse_bot
import sefaz_bot

# Configuraçoes visuais - Temas Contili
ctk.set_appearance_mode("Dark")  # Forçamos o Dark Mode conforme o design premium
ctk.set_default_color_theme("blue")

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("NFSe Bot - Gestor Fiscal Contili")
        self.geometry("1000x750")
        self.resizable(True, True)
        
        # Grid layout 1x2 (Sidebar e Content)
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)

        self.selected_file_path = None
        self.selected_dir_path = None
        
        # --- SIDEBAR ---
        self.sidebar_frame = ctk.CTkFrame(self, width=200, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(5, weight=1) # Espaço flexível no fim

        # Logo na Sidebar
        try:
            logo_path = resource_path("contili_logo.png")
            if os.path.exists(logo_path):
                logo_img = ctk.CTkImage(light_image=Image.open(logo_path), dark_image=Image.open(logo_path), size=(160, 50))
                self.lbl_logo_sidebar = ctk.CTkLabel(self.sidebar_frame, image=logo_img, text="")
                self.lbl_logo_sidebar.grid(row=0, column=0, padx=20, pady=20)
        except:
            self.lbl_logo_sidebar = ctk.CTkLabel(self.sidebar_frame, text="CONTILI", font=ctk.CTkFont(size=20, weight="bold"))
            self.lbl_logo_sidebar.grid(row=0, column=0, padx=20, pady=20)

        # Botões de Navegação
        self.btn_nav_notas = ctk.CTkButton(self.sidebar_frame, corner_radius=10, height=40, border_spacing=10, text="📄 Extrato de Notas",
                                           fg_color="transparent", text_color=("gray10", "gray90"), hover_color=("gray70", "gray30"),
                                           anchor="w", command=self.btn_nav_notas_event)
        self.btn_nav_notas.grid(row=1, column=0, sticky="ew", padx=20, pady=5)

        self.btn_nav_dominio = ctk.CTkButton(self.sidebar_frame, corner_radius=10, height=40, border_spacing=10, text="📊 Relatórios Domínio",
                                           fg_color="transparent", text_color=("gray10", "gray90"), hover_color=("gray70", "gray30"),
                                           anchor="w", command=self.btn_nav_dominio_event)
        self.btn_nav_dominio.grid(row=2, column=0, sticky="ew", padx=20, pady=5)

        self.btn_nav_comparativo = ctk.CTkButton(self.sidebar_frame, corner_radius=10, height=40, border_spacing=10, text="⚖️ Comparativos",
                                           fg_color="transparent", text_color=("gray10", "gray90"), hover_color=("gray70", "gray30"),
                                           anchor="w", command=self.btn_nav_comparativo_event)
        self.btn_nav_comparativo.grid(row=3, column=0, sticky="ew", padx=20, pady=5)

        self.btn_nav_sefaz = ctk.CTkButton(self.sidebar_frame, corner_radius=10, height=40, border_spacing=10, text="🏛️ Gestor Sefaz",
                                           fg_color="transparent", text_color=("gray10", "gray90"), hover_color=("gray70", "gray30"),
                                           anchor="w", command=self.btn_nav_sefaz_event)
        self.btn_nav_sefaz.grid(row=4, column=0, sticky="ew", padx=20, pady=5)

        # --- FRAMES DE CONTEÚDO ---
        
        # 1. Extrato de Notas (Home)
        self.notas_frame = ctk.CTkScrollableFrame(self, corner_radius=0, fg_color="transparent")
        self.setup_notas_frame()

        # 2. Relatórios Domínio
        self.dominio_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.setup_placeholder_frame(self.dominio_frame, "📊 Relatórios Domínio", "Módulo em desenvolvimento para extração de dados do sistema Domínio.")

        # 3. Comparativos
        self.comparativo_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.setup_placeholder_frame(self.comparativo_frame, "⚖️ Comparativos", "Módulo em desenvolvimento para cruzamento de NFSe x Domínio.")

        # 4. Gestor Sefaz
        self.sefaz_frame = ctk.CTkScrollableFrame(self, corner_radius=0, fg_color="transparent")
        self.setup_sefaz_frame()

        # Selecionar porta padrão
        self.select_frame_by_name("notas")

    def setup_notas_frame(self):
        # Titulo Principal
        self.title_label = ctk.CTkLabel(self.notas_frame, text="Extrator de Notas Fiscais (NFSe)", font=ctk.CTkFont(size=24, weight="bold"))
        self.title_label.pack(pady=(20, 20))
        
        # Frame Passo 1
        self.frame_p1 = ctk.CTkFrame(self.notas_frame, corner_radius=20)
        self.frame_p1.pack(pady=10, padx=40, fill="x")
        
        self.lbl_p1 = ctk.CTkLabel(self.frame_p1, text="Passo 1: Preparar Ambiente", font=ctk.CTkFont(weight="bold", size=15))
        self.lbl_p1.pack(pady=(15, 5))
        self.desc_p1 = ctk.CTkLabel(self.frame_p1, text="É obrigatório fechar todos os navegadores Chrome antes de iniciar.", text_color="gray")
        self.desc_p1.pack(pady=(0, 5))
        self.btn_kill = ctk.CTkButton(self.frame_p1, text="⛔ Encerrar Chrome Aberto", fg_color="#c0392b", hover_color="#e74c3c", corner_radius=20, command=self.kill_chrome)
        self.btn_kill.pack(pady=15)
        
        # Frame Passo 2
        self.frame_p2 = ctk.CTkFrame(self.notas_frame, corner_radius=20)
        self.frame_p2.pack(pady=10, padx=40, fill="x")
        
        self.lbl_p2 = ctk.CTkLabel(self.frame_p2, text="Passo 2: Iniciar Navegador Integrado", font=ctk.CTkFont(weight="bold", size=15))
        self.lbl_p2.pack(pady=(15, 5))
        self.desc_p2 = ctk.CTkLabel(self.frame_p2, text="Clique abaixo para abrir o portal e selecione o certificado.", text_color="gray")
        self.desc_p2.pack(pady=(0, 5))
        self.btn_start_chrome = ctk.CTkButton(self.frame_p2, text="🌐 Abrir Chrome do Robô", fg_color="#0052CC", hover_color="#003D99", corner_radius=20, command=self.start_chrome)
        self.btn_start_chrome.pack(pady=15)
        
        # Frame Passo 3 (Extraçao)
        self.frame_p3 = ctk.CTkFrame(self.notas_frame, corner_radius=20)
        self.frame_p3.pack(pady=10, padx=40, fill="x")
        
        self.lbl_p3 = ctk.CTkLabel(self.frame_p3, text="Passo 3: Dados da Extração", font=ctk.CTkFont(weight="bold", size=15))
        self.lbl_p3.pack(pady=(15, 5))
        
        # Grid para inputs
        self.grid_frame = ctk.CTkFrame(self.frame_p3, fg_color="transparent")
        self.grid_frame.pack(pady=10, padx=20, fill="x")
        
        # Coluna Esq
        self.lbl_comp = ctk.CTkLabel(self.grid_frame, text="Competência (MM/AAAA):")
        self.lbl_comp.grid(row=0, column=0, padx=10, pady=10, sticky="e")
        
        self.entry_comp = ctk.CTkEntry(self.grid_frame, placeholder_text="Ex: 02/2026", width=150, corner_radius=10)
        self.entry_comp.grid(row=0, column=1, padx=10, pady=10, sticky="w")
        self.entry_comp.bind("<KeyRelease>", self.format_date_mask)
        
        # Diretório Destino
        self.btn_dir = ctk.CTkButton(self.grid_frame, text="📂 Pasta Destino (Opcional)", fg_color="gray40", height=35, corner_radius=10, command=self.select_dir)
        self.btn_dir.grid(row=1, column=0, padx=10, pady=10, sticky="e")
        
        self.lbl_dirname = ctk.CTkLabel(self.grid_frame, text="Padrão (pasta 'livros')", text_color="gray60", font=ctk.CTkFont(size=12))
        self.lbl_dirname.grid(row=1, column=1, padx=10, pady=10, sticky="w")
        
        # Arquivo de Relação
        self.btn_file = ctk.CTkButton(self.grid_frame, text="📁 Relação Empresas (.xlsx)", fg_color="gray40", height=35, corner_radius=10, command=self.select_file)
        self.btn_file.grid(row=2, column=0, padx=10, pady=10, sticky="e")
        
        self.lbl_filename = ctk.CTkLabel(self.grid_frame, text="Nenhum arquivo...", text_color="#F1C40F", font=ctk.CTkFont(size=13, weight="bold"))
        self.lbl_filename.grid(row=2, column=1, padx=10, pady=10, sticky="w")
        
        self.btn_model = ctk.CTkButton(self.frame_p3, text="📄 Gerar Modelo Vazio", fg_color="darkblue", text_color="white", corner_radius=20, command=self.generate_model)
        self.btn_model.pack(pady=(0, 15))
        
        # Botão Iniciar Processo
        self.btn_run = ctk.CTkButton(self.notas_frame, text="🚀 INICIAR EXTRAÇÃO", font=ctk.CTkFont(size=22, weight="bold"), height=80, fg_color="#27ae60", hover_color="#2ecc71", corner_radius=20, command=self.start_robot_thread)
        self.btn_run.pack(pady=(20, 10), padx=40, fill="x")
        
        # Barra de status
        self.status_var = ctk.StringVar(value="Aguardando...")
        self.lbl_status = ctk.CTkLabel(self.notas_frame, textvariable=self.status_var, font=ctk.CTkFont(size=12, weight="bold"))
        self.lbl_status.pack(pady=5)

        # Log em Tempo Real
        self.lbl_log_title = ctk.CTkLabel(self.notas_frame, text="Log de Acompanhamento:", font=ctk.CTkFont(weight="bold"))
        self.lbl_log_title.pack(pady=(20, 0), padx=40, anchor="w")
        
        self.log_textbox = ctk.CTkTextbox(self.notas_frame, height=250, corner_radius=15, font=ctk.CTkFont(family="Consolas", size=11))
        self.log_textbox.pack(pady=(5, 40), padx=40, fill="x")
        self.log_textbox.configure(state="disabled")

    def setup_placeholder_frame(self, frame, title, description):
        title_lbl = ctk.CTkLabel(frame, text=title, font=ctk.CTkFont(size=24, weight="bold"))
        title_lbl.pack(pady=(100, 20))
        
        desc_lbl = ctk.CTkLabel(frame, text=description, text_color="gray", font=ctk.CTkFont(size=16))
        desc_lbl.pack(pady=10)
        
        soon_lbl = ctk.CTkLabel(frame, text="PROXIMAMENTE", fg_color="#F1C40F", text_color="black", corner_radius=10, padx=10, pady=5)
        soon_lbl.pack(pady=30)

    # --- Lógica de Navegação ---
    def select_frame_by_name(self, name):
        # Reset buttons colors
        self.btn_nav_notas.configure(fg_color=("gray75", "gray25") if name == "notas" else "transparent")
        self.btn_nav_dominio.configure(fg_color=("gray75", "gray25") if name == "dominio" else "transparent")
        self.btn_nav_comparativo.configure(fg_color=("gray75", "gray25") if name == "comparativo" else "transparent")
        self.btn_nav_sefaz.configure(fg_color=("gray75", "gray25") if name == "sefaz" else "transparent")

        # Show selected frame
        if name == "notas":
            self.notas_frame.grid(row=0, column=1, sticky="nsew")
        else:
            self.notas_frame.grid_forget()
            
        if name == "dominio":
            self.dominio_frame.grid(row=0, column=1, sticky="nsew")
        else:
            self.dominio_frame.grid_forget()
            
        if name == "comparativo":
            self.comparativo_frame.grid(row=0, column=1, sticky="nsew")
        else:
            self.comparativo_frame.grid_forget()

        if name == "sefaz":
            self.sefaz_frame.grid(row=0, column=1, sticky="nsew")
        else:
            self.sefaz_frame.grid_forget()

    def btn_nav_notas_event(self):
        self.select_frame_by_name("notas")

    def btn_nav_dominio_event(self):
        self.select_frame_by_name("dominio")

    def btn_nav_comparativo_event(self):
        self.select_frame_by_name("comparativo")

    def btn_nav_sefaz_event(self):
        self.select_frame_by_name("sefaz")

    # --- Lógica de Funçoes (Originais mantidas) ---
    def kill_chrome(self):
        answer = messagebox.askyesno("Confirmar", "Tem certeza que deseja forçar o fechamento de TODAS as janelas do Google Chrome?\n(Você perderá o que não estiver salvo nas abas)")
        if answer:
            try:
                os.system("taskkill /im chrome.exe /f")
                self.status_var.set("Processos do Chrome encerrados.")
                messagebox.showinfo("Sucesso", "Todas as instâncias do Google Chrome foram devidamente encerradas!")
            except Exception as e:
                messagebox.showerror("Erro", f"Não foi possível encerrar: {e}")

    def start_chrome(self):
        user_data_path = os.path.join(get_app_dir(), "chrome_profile")
        import shutil
        executable = shutil.which("chrome") or shutil.which("chrome.exe")
        if not executable:
            chrome_paths = [
                r"C:\Program Files\Google\Chrome\Application\chrome.exe",
                r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
                os.path.join(os.environ.get("LOCALAPPDATA", ""), r"Google\Chrome\Application\chrome.exe")
            ]
            for path in chrome_paths:
                if os.path.exists(path):
                    executable = path
                    break
        if not executable:
            messagebox.showerror("Erro", "Não foi possível encontrar o executável do Google Chrome neste PC.")
            return
        url_login = "https://nfse.campogrande.ms.gov.br/notafiscal/paginas/portal/index.html#/login"
        chrome_args = [executable, "--remote-debugging-port=9222", f"--user-data-dir={user_data_path}", "--no-first-run", "--no-default-browser-check", url_login]
        try:
            self.status_var.set("Iniciando navegador...")
            subprocess.Popen(chrome_args, stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            self.status_var.set("Navegador aberto.")
            messagebox.showinfo("Sucesso", "Chrome aberto! Siga com o certificado no site.")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao abrir Chrome: {e}")

    def select_file(self):
        filepath = filedialog.askopenfilename(title="Selecione a Relação de Empresas", filetypes=[("Excel files", "*.xlsx *.xls")])
        if filepath:
            self.selected_file_path = filepath
            filename = os.path.basename(filepath)
            self.lbl_filename.configure(text=filename, text_color="white")
            
    def select_dir(self):
        dirpath = filedialog.askdirectory(title="Selecione a Pasta de Destino")
        if dirpath:
            self.selected_dir_path = dirpath
            dirname = os.path.basename(dirpath)
            self.lbl_dirname.configure(text=f".../{dirname}", text_color="white")
            
    def format_date_mask(self, event):
        if event.keysym == "BackSpace": return
        text = self.entry_comp.get().replace("/", "")
        digits = "".join([c for c in text if c.isdigit()])[:6]
        formatted = ""
        if len(digits) > 2: formatted = digits[:2] + "/" + digits[2:]
        else: formatted = digits
        if self.entry_comp.get() != formatted:
            self.entry_comp.delete(0, "end")
            self.entry_comp.insert(0, formatted)
            
    def generate_model(self):
        try:
            model_path = os.path.join(get_app_dir(), "Modelo_Empresas.xlsx")
            df = pd.DataFrame(columns=["Razão Social", "CNPJ"])
            df.to_excel(model_path, index=False)
            messagebox.showinfo("Modelo Criado", f"O modelo foi criado em:\n{model_path}")
            os.startfile(model_path)
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível gerar o modelo: {e}")

    def update_log(self, message):
        def _append():
            self.log_textbox.configure(state="normal")
            agora = pd.Timestamp.now().strftime("%H:%M:%S")
            linha_log = f"[{agora}] {message}"
            self.log_textbox.insert("end", linha_log + "\n")
            self.log_textbox.see("end")
            self.log_textbox.configure(state="disabled")
            
            # Salvar no arquivo de log geral
            log_path = os.path.join(get_app_dir(), "Log_GestorNFSe.txt")
            try:
                with open(log_path, "a", encoding="utf-8") as f:
                    data_agora = pd.Timestamp.now().strftime("%d/%m/%Y %H:%M:%S")
                    f.write(f"[{data_agora}] {message}\n")
            except:
                pass
                
        self.after(0, _append)

    def is_file_locked(self, filepath):
        if not os.path.exists(filepath): return False
        try:
            with open(filepath, 'a'): pass
            return False
        except IOError: return True

    def start_robot_thread(self):
        comp = self.entry_comp.get().strip()
        if not comp or len(comp) < 7:
            messagebox.showwarning("Aviso", "Preencha a competência corretamente (Ex: 02/2026).")
            return
        if not self.selected_file_path:
            messagebox.showwarning("Aviso", "Selecione o arquivo da Relação de Empresas.")
            return
        if self.is_file_locked(self.selected_file_path):
            messagebox.showerror("Arquivo Aberto", "O arquivo de relação está aberto. Feche-o antes de continuar.")
            return
        self.btn_run.configure(state="disabled")
        self.btn_start_chrome.configure(state="disabled")
        self.status_var.set("Robô em Execução...")
        thread = threading.Thread(target=self.run_bot_background, args=(comp, self.selected_file_path, self.selected_dir_path))
        thread.daemon = True
        thread.start()
        
    def run_bot_background(self, comp, filepath, dirpath):
        import time
        try:
            start_time = time.time()
            df_empresas = pd.read_excel(filepath)
            qtd_empresas = len(df_empresas.dropna(how="all"))

            self.update_log(f"▶️ Iniciando competência {comp}...")
            nfse_bot.run_bot(comp, filepath, wait_for_input=False, output_dir=dirpath, log_callback=self.update_log)
            
            end_time = time.time()
            tempo_robo_segundos = int(end_time - start_time)
            tempo_robo_min = tempo_robo_segundos // 60
            tempo_robo_seg = tempo_robo_segundos % 60
            
            # 10 minutos (600 segundos) por empresa manualmente
            tempo_humano_minutos = qtd_empresas * 10
            horas_h = tempo_humano_minutos // 60
            minutos_h = tempo_humano_minutos % 60
            
            economia_segundos = (tempo_humano_minutos * 60) - tempo_robo_segundos
            economia_h = economia_segundos // 3600
            economia_m = (economia_segundos % 3600) // 60
            
            msg_log = (
                "\n==================================================\n"
                "📊 RELATÓRIO DE EFICIÊNCIA DO GESTOR NFSe 📊\n"
                "==================================================\n"
                f"🏢 Empresas Processadas: {qtd_empresas}\n"
                f"⏱️ Tempo Estimado (Manual): {horas_h}h {minutos_h:02d}m ({tempo_humano_minutos} minutos)\n"
                f"⚡ Tempo Gasto pelo Robô: {tempo_robo_min}m {tempo_robo_seg:02d}s\n"
                f"🚀 GANHO EFETIVO DE TEMPO: {economia_h}h {economia_m:02d}m economizadas!\n"
                "=================================================="
            )
            self.update_log(msg_log)
            messagebox.showinfo("Extração Finalizada!", f"Processo concluído com sucesso.\n\nEconomia de tempo gerada: {economia_h}h e {economia_m}m")
        except Exception as e:
            self.update_log(f"❌ ERRO CRÍTICO: {str(e)}")
            messagebox.showerror("Erro na Extração", f"O processo falhou: {e}")
        finally:
            self.btn_run.configure(state="normal")
            self.btn_start_chrome.configure(state="normal")
            self.status_var.set("Pronto.")

    # --- Métodos para o Gestor Sefaz ---
    def setup_sefaz_frame(self):
        # Titulo Principal
        self.title_label_sefaz = ctk.CTkLabel(self.sefaz_frame, text="Gestor Sefaz (NF-e / CT-e)", font=ctk.CTkFont(size=24, weight="bold"))
        self.title_label_sefaz.pack(pady=(20, 20))
        
        # Frame Passo 1
        self.frame_s1 = ctk.CTkFrame(self.sefaz_frame, corner_radius=20)
        self.frame_s1.pack(pady=10, padx=40, fill="x")
        
        self.lbl_s1 = ctk.CTkLabel(self.frame_s1, text="Passo 1: Preparar Ambiente", font=ctk.CTkFont(weight="bold", size=15))
        self.lbl_s1.pack(pady=(15, 5))
        self.btn_kill_sefaz = ctk.CTkButton(self.frame_s1, text="⛔ Encerrar Chrome Aberto", fg_color="#c0392b", hover_color="#e74c3c", corner_radius=20, command=self.kill_chrome)
        self.btn_kill_sefaz.pack(pady=15)
        
        # Frame Passo 2
        self.frame_s2 = ctk.CTkFrame(self.sefaz_frame, corner_radius=20)
        self.frame_s2.pack(pady=10, padx=40, fill="x")
        
        self.lbl_s2 = ctk.CTkLabel(self.frame_s2, text="Passo 2: Iniciar Navegador Sefaz", font=ctk.CTkFont(weight="bold", size=15))
        self.lbl_s2.pack(pady=(15, 5))
        self.desc_s2 = ctk.CTkLabel(self.frame_s2, text="Abra o Chrome para validar o Certificado Digital no e-Serviços.", text_color="gray")
        self.desc_s2.pack(pady=(0, 5))
        self.btn_start_chrome_sefaz = ctk.CTkButton(self.frame_s2, text="🌐 Abrir Chrome da Sefaz", fg_color="#0052CC", hover_color="#003D99", corner_radius=20, command=self.start_chrome_sefaz)
        self.btn_start_chrome_sefaz.pack(pady=15)
        
        # Frame Passo 3 (Extração)
        self.frame_s3 = ctk.CTkFrame(self.sefaz_frame, corner_radius=20)
        self.frame_s3.pack(pady=10, padx=40, fill="x")
        
        self.lbl_s3 = ctk.CTkLabel(self.frame_s3, text="Passo 3: Dados da Extração", font=ctk.CTkFont(weight="bold", size=15))
        self.lbl_s3.pack(pady=(15, 5))
        
        self.grid_frame_s = ctk.CTkFrame(self.frame_s3, fg_color="transparent")
        self.grid_frame_s.pack(pady=10, padx=20, fill="x")
        
        self.lbl_comp_s = ctk.CTkLabel(self.grid_frame_s, text="Competência (MM/AAAA):")
        self.lbl_comp_s.grid(row=0, column=0, padx=10, pady=10, sticky="e")
        
        self.entry_comp_sefaz = ctk.CTkEntry(self.grid_frame_s, placeholder_text="Ex: 02/2026", width=150, corner_radius=10)
        self.entry_comp_sefaz.grid(row=0, column=1, padx=10, pady=10, sticky="w")
        self.entry_comp_sefaz.bind("<KeyRelease>", self.format_date_mask_sefaz)
        
        self.btn_file_sefaz = ctk.CTkButton(self.grid_frame_s, text="📁 Relação Empresas (.xlsx)", fg_color="gray40", height=35, corner_radius=10, command=self.select_file_sefaz)
        self.btn_file_sefaz.grid(row=1, column=0, padx=10, pady=10, sticky="e")
        
        self.lbl_filename_sefaz = ctk.CTkLabel(self.grid_frame_s, text="Nenhum arquivo...", text_color="#F1C40F", font=ctk.CTkFont(size=13, weight="bold"))
        self.lbl_filename_sefaz.grid(row=1, column=1, padx=10, pady=10, sticky="w")

        self.btn_dir_sefaz = ctk.CTkButton(self.grid_frame_s, text="📂 Pasta Destino (Opcional)", fg_color="gray40", height=35, corner_radius=10, command=self.select_dir_sefaz)
        self.btn_dir_sefaz.grid(row=2, column=0, padx=10, pady=10, sticky="e")
        
        self.lbl_dirname_sefaz = ctk.CTkLabel(self.grid_frame_s, text="Padrão (pasta local)", text_color="gray60", font=ctk.CTkFont(size=12))
        self.lbl_dirname_sefaz.grid(row=2, column=1, padx=10, pady=10, sticky="w")
        
        self.btn_model_sefaz = ctk.CTkButton(self.frame_s3, text="📄 Gerar Modelo Vazio Sefaz", fg_color="darkblue", text_color="white", corner_radius=20, command=self.generate_model_sefaz)
        self.btn_model_sefaz.pack(pady=(0, 15))
        
        # Botão Iniciar Processo
        self.btn_run_sefaz = ctk.CTkButton(self.sefaz_frame, text="🚀 INICIAR EXTRAÇÃO SEFAZ", font=ctk.CTkFont(size=22, weight="bold"), height=80, fg_color="#8e44ad", hover_color="#9b59b6", corner_radius=20, command=self.start_sefaz_thread)
        self.btn_run_sefaz.pack(pady=(20, 10), padx=40, fill="x")
        
        self.status_var_sefaz = ctk.StringVar(value="Aguardando...")
        self.lbl_status_sefaz = ctk.CTkLabel(self.sefaz_frame, textvariable=self.status_var_sefaz, font=ctk.CTkFont(size=12, weight="bold"))
        self.lbl_status_sefaz.pack(pady=5)

        self.lbl_log_title_sefaz = ctk.CTkLabel(self.sefaz_frame, text="Log de Acompanhamento Sefaz:", font=ctk.CTkFont(weight="bold"))
        self.lbl_log_title_sefaz.pack(pady=(20, 0), padx=40, anchor="w")
        
        self.log_textbox_sefaz = ctk.CTkTextbox(self.sefaz_frame, height=250, corner_radius=15, font=ctk.CTkFont(family="Consolas", size=11))
        self.log_textbox_sefaz.pack(pady=(5, 40), padx=40, fill="x")
        self.log_textbox_sefaz.configure(state="disabled")

    def select_file_sefaz(self):
        filepath = filedialog.askopenfilename(title="Selecione a Relação de Empresas Sefaz", filetypes=[("Excel files", "*.xlsx *.xls")])
        if filepath:
            self.selected_file_path_sefaz = filepath
            filename = os.path.basename(filepath)
            self.lbl_filename_sefaz.configure(text=filename, text_color="white")
            
    def select_dir_sefaz(self):
        dirpath = filedialog.askdirectory(title="Selecione a Pasta de Destino Sefaz")
        if dirpath:
            self.selected_dir_path_sefaz = dirpath
            dirname = os.path.basename(dirpath)
            self.lbl_dirname_sefaz.configure(text=f".../{dirname}", text_color="white")

    def format_date_mask_sefaz(self, event):
        if event.keysym == "BackSpace": return
        text = self.entry_comp_sefaz.get().replace("/", "")
        digits = "".join([c for c in text if c.isdigit()])[:6]
        formatted = ""
        if len(digits) > 2: formatted = digits[:2] + "/" + digits[2:]
        else: formatted = digits
        if self.entry_comp_sefaz.get() != formatted:
            self.entry_comp_sefaz.delete(0, "end")
            self.entry_comp_sefaz.insert(0, formatted)

    def generate_model_sefaz(self):
        try:
            model_path = os.path.join(get_app_dir(), "Modelo_Empresas_Sefaz.xlsx")
            df = pd.DataFrame(columns=["Código Domínio", "Nome Empresa", "CNPJ", "Inscrição Estadual"])
            df.to_excel(model_path, index=False)
            messagebox.showinfo("Modelo Criado", f"O modelo foi criado em:\n{model_path}")
            os.startfile(model_path)
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível gerar o modelo: {e}")

    def update_log_sefaz(self, message):
        def _append():
            self.log_textbox_sefaz.configure(state="normal")
            agora = pd.Timestamp.now().strftime("%H:%M:%S")
            linha_log = f"[{agora}] {message}"
            self.log_textbox_sefaz.insert("end", linha_log + "\n")
            self.log_textbox_sefaz.see("end")
            self.log_textbox_sefaz.configure(state="disabled")
            
            # Salvar no arquivo de log geral
            log_path = os.path.join(get_app_dir(), "Log_GestorSefaz.txt")
            try:
                with open(log_path, "a", encoding="utf-8") as f:
                    # Adiciona a data no arquivo físico para histórico de longo prazo
                    data_agora = pd.Timestamp.now().strftime("%d/%m/%Y %H:%M:%S")
                    f.write(f"[{data_agora}] {message}\n")
            except:
                pass
                
        self.after(0, _append)

    def start_chrome_sefaz(self):
        user_data_path = os.path.join(get_app_dir(), "chrome_profile_sefaz")
        import shutil
        executable = shutil.which("chrome") or shutil.which("chrome.exe")
        if not executable:
            chrome_paths = [
                r"C:\Program Files\Google\Chrome\Application\chrome.exe",
                r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
                os.path.join(os.environ.get("LOCALAPPDATA", ""), r"Google\Chrome\Application\chrome.exe")
            ]
            for path in chrome_paths:
                if os.path.exists(path):
                    executable = path
                    break
        if not executable:
            messagebox.showerror("Erro", "Não foi possível encontrar o executável do Google Chrome neste PC.")
            return
        url_login = "https://eservicos.sefaz.ms.gov.br/"
        # porta 9223 para o sefaz
        chrome_args = [executable, "--remote-debugging-port=9223", f"--user-data-dir={user_data_path}", "--no-first-run", "--no-default-browser-check", url_login]
        try:
            self.status_var_sefaz.set("Iniciando navegador Sefaz...")
            subprocess.Popen(chrome_args, stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            self.status_var_sefaz.set("Navegador aberto.")
            messagebox.showinfo("Sucesso", "Chrome aberto! Valide o certificado na Sefaz.")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao abrir Chrome: {e}")

    def start_sefaz_thread(self):
        comp = self.entry_comp_sefaz.get().strip()
        if not comp or len(comp) < 7:
            messagebox.showwarning("Aviso", "Preencha a competência corretamente (Ex: 02/2026).")
            return
        if not hasattr(self, 'selected_file_path_sefaz') or not self.selected_file_path_sefaz:
            messagebox.showwarning("Aviso", "Selecione o arquivo da Relação de Empresas da Sefaz.")
            return
        if self.is_file_locked(self.selected_file_path_sefaz):
            messagebox.showerror("Arquivo Aberto", "O arquivo de relação está aberto. Feche-o antes de continuar.")
            return
        self.btn_run_sefaz.configure(state="disabled")
        self.btn_start_chrome_sefaz.configure(state="disabled")
        self.status_var_sefaz.set("Robô Sefaz em Execução...")
        
        dirpath = getattr(self, 'selected_dir_path_sefaz', None)
        thread = threading.Thread(target=self.run_sefaz_background, args=(comp, self.selected_file_path_sefaz, dirpath))
        thread.daemon = True
        thread.start()
        
    def run_sefaz_background(self, comp, filepath, dirpath):
        import time
        try:
            start_time = time.time()
            df_empresas = pd.read_excel(filepath)
            qtd_empresas = len(df_empresas.dropna(how="all"))
            
            self.update_log_sefaz(f"▶️ Iniciando extração Sefaz - {comp}...")
            sefaz_bot.run_sefaz_bot(comp, filepath, output_dir=dirpath, log_callback=self.update_log_sefaz)
            
            end_time = time.time()
            tempo_robo_segundos = int(end_time - start_time)
            tempo_robo_min = tempo_robo_segundos // 60
            tempo_robo_seg = tempo_robo_segundos % 60
            
            tempo_humano_minutos = qtd_empresas * 15
            horas_h = tempo_humano_minutos // 60
            minutos_h = tempo_humano_minutos % 60
            
            economia_segundos = (tempo_humano_minutos * 60) - tempo_robo_segundos
            economia_h = economia_segundos // 3600
            economia_m = (economia_segundos % 3600) // 60
            
            msg_log = (
                "\n==================================================\n"
                "📊 RELATÓRIO DE EFICIÊNCIA DO GESTOR SEFAZ 📊\n"
                "==================================================\n"
                f"🏢 Empresas Processadas: {qtd_empresas}\n"
                f"⏱️ Tempo Estimado (Manual): {horas_h}h {minutos_h:02d}m ({tempo_humano_minutos} minutos)\n"
                f"⚡ Tempo Gasto pelo Robô: {tempo_robo_min}m {tempo_robo_seg:02d}s\n"
                f"🚀 GANHO EFETIVO DE TEMPO: {economia_h}h {economia_m:02d}m economizadas!\n"
                "=================================================="
            )
            self.update_log_sefaz(msg_log)
            messagebox.showinfo("Extração Finalizada!", f"Processo concluído com sucesso.\n\nEconomia de tempo gerada: {economia_h}h e {economia_m}m")
        except Exception as e:
            self.update_log_sefaz(f"❌ ERRO CRÍTICO: {str(e)}")
            messagebox.showerror("Erro na Extração", f"O processo falhou: {e}")
        finally:
            self.btn_run_sefaz.configure(state="normal")
            self.btn_start_chrome_sefaz.configure(state="normal")
            self.status_var_sefaz.set("Pronto.")

if __name__ == "__main__":
    app = App()
    app.mainloop()
