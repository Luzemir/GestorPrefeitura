import sys
import os
import threading
import subprocess
import customtkinter as ctk
from tkinter import filedialog, messagebox
import pandas as pd
from PIL import Image

def resource_path(relative_path):
    """ Retorna o caminho absoluto do arquivo, seja localmente ou empacotado pelo PyInstaller. """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)

# Adiciona o diretorio atual ao path para poder importar o bot
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import nfse_bot

# Configuraçoes visuais
ctk.set_appearance_mode("System")  # Modes: "System" (standard), "Dark", "Light"
ctk.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("NFSe Bot - Gestor Fiscal")
        self.geometry("600x700")
        self.resizable(False, False)
        
        self.selected_file_path = None
        
        self.create_widgets()

    def create_widgets(self):
        # Logo da Contili no topo
        try:
            logo_path = resource_path("contili_logo.png")
            if os.path.exists(logo_path):
                # Aumentando um pouco a logo
                logo_img = ctk.CTkImage(light_image=Image.open(logo_path), dark_image=Image.open(logo_path), size=(200, 70))
                self.lbl_logo = ctk.CTkLabel(self, image=logo_img, text="")
                self.lbl_logo.pack(pady=(15, 0))
        except Exception as e:
            pass # Ignora se falhar ao carregar logo
            
        # Titulo Principal
        self.title_label = ctk.CTkLabel(self, text="Extrator de Notas Fiscais (Campo Grande/MS)", font=ctk.CTkFont(size=20, weight="bold"))
        self.title_label.pack(pady=(10, 10))
        
        # Frame Passo 1
        self.frame_p1 = ctk.CTkFrame(self)
        self.frame_p1.pack(pady=10, padx=20, fill="x")
        
        self.lbl_p1 = ctk.CTkLabel(self.frame_p1, text="Passo 1: Preparar Ambiente", font=ctk.CTkFont(weight="bold"))
        self.lbl_p1.pack(pady=(10, 5))
        self.desc_p1 = ctk.CTkLabel(self.frame_p1, text="É obrigatório fechar todos os navegadores Chrome antes de iniciar.", text_color="gray")
        self.desc_p1.pack(pady=(0, 5))
        self.btn_kill = ctk.CTkButton(self.frame_p1, text="⛔ Encerrar Chrome Aberto", fg_color="#c0392b", hover_color="#e74c3c", command=self.kill_chrome)
        self.btn_kill.pack(pady=10)
        
        # Frame Passo 2
        self.frame_p2 = ctk.CTkFrame(self)
        self.frame_p2.pack(pady=10, padx=20, fill="x")
        
        self.lbl_p2 = ctk.CTkLabel(self.frame_p2, text="Passo 2: Iniciar Navegador Integrado", font=ctk.CTkFont(weight="bold"))
        self.lbl_p2.pack(pady=(10, 5))
        self.desc_p2 = ctk.CTkLabel(self.frame_p2, text="Clique abaixo, insira o certificado no portal e aguarde a tela principal.", text_color="gray")
        self.desc_p2.pack(pady=(0, 5))
        self.btn_start_chrome = ctk.CTkButton(self.frame_p2, text="🌐 Abrir Chrome do Robô", command=self.start_chrome)
        self.btn_start_chrome.pack(pady=10)
        
        # Frame Passo 3 e 4 (Extraçao)
        self.frame_p3 = ctk.CTkFrame(self)
        self.frame_p3.pack(pady=10, padx=20, fill="x")
        
        self.lbl_p3 = ctk.CTkLabel(self.frame_p3, text="Passo 3: Dados da Extração", font=ctk.CTkFont(weight="bold"))
        self.lbl_p3.pack(pady=(10, 5))
        
        # Grid para inputs
        self.grid_frame = ctk.CTkFrame(self.frame_p3, fg_color="transparent")
        self.grid_frame.pack(pady=10, padx=10, fill="x")
        
        # Coluna Esq
        self.lbl_comp = ctk.CTkLabel(self.grid_frame, text="Competência (MM/AAAA):")
        self.lbl_comp.grid(row=0, column=0, padx=10, pady=5, sticky="e")
        
        self.entry_comp = ctk.CTkEntry(self.grid_frame, placeholder_text="Ex: 02/2026", width=120)
        self.entry_comp.grid(row=0, column=1, padx=10, pady=5, sticky="w")
        
        # Coluna Dir
        self.btn_file = ctk.CTkButton(self.grid_frame, text="📁 Selecionar Relação (.xlsx)", fg_color="gray", command=self.select_file)
        self.btn_file.grid(row=1, column=0, padx=10, pady=10, sticky="e")
        
        self.lbl_filename = ctk.CTkLabel(self.grid_frame, text="Nenhum arquivo...", text_color=("#c0392b", "#F1C40F"), font=ctk.CTkFont(size=14, weight="bold"))
        self.lbl_filename.grid(row=1, column=1, padx=10, pady=10, sticky="w")
        
        self.btn_model = ctk.CTkButton(self.grid_frame, text="📄 Gerar Modelo Vazio", fg_color="#2980b9", text_color="white", hover_color="#3498db", border_width=0, command=self.generate_model)
        self.btn_model.grid(row=2, column=0, columnspan=2, pady=(5, 10))
        
        # Botão Iniciar Processo (Expandido)
        self.btn_run = ctk.CTkButton(self, text="🚀 INICIAR EXTRAÇÃO", font=ctk.CTkFont(size=24, weight="bold"), height=70, fg_color="#27ae60", hover_color="#2ecc71", command=self.start_robot_thread)
        self.btn_run.pack(pady=(15, 10))
        
        # Barra de status (Texto dinâmico para modos Claro/Escuro)
        self.status_var = ctk.StringVar(value="Aguardando...")
        self.lbl_status = ctk.CTkLabel(self, textvariable=self.status_var, text_color=("#2c3e50", "#bdc3c7"), font=ctk.CTkFont(size=12, weight="bold"))
        self.lbl_status.pack(pady=5)
        
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
        user_data_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "chrome_profile")
        cmd = f'Start-Process "chrome.exe" -ArgumentList "--remote-debugging-port=9222 --user-data-dir={user_data_path}"'
        try:
            subprocess.Popen(["powershell", "-Command", cmd], creationflags=subprocess.CREATE_NO_WINDOW)
            self.status_var.set("Navegador aberto na porta de depuração.")
            messagebox.showinfo("Sucesso", "Chrome aberto! Siga com o certificado e a tela inicial no site.")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao abrir Chrome: {e}")

    def select_file(self):
        filepath = filedialog.askopenfilename(
            title="Selecione a Relação de Empresas",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if filepath:
            self.selected_file_path = filepath
            filename = os.path.basename(filepath)
            self.lbl_filename.configure(text=filename, text_color="white")
            
    def generate_model(self):
        try:
            model_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "Modelo_Empresas.xlsx")
            df = pd.DataFrame(columns=["Razão Social", "CNPJ"])
            df.to_excel(model_path, index=False)
            messagebox.showinfo("Modelo Criado", f"O modelo foi criado em:\n{model_path}\nEle será aberto agora.")
            os.startfile(model_path)
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível gerar o modelo:\n{e}")

    def start_robot_thread(self):
        comp = self.entry_comp.get().strip()
        if not comp or len(comp) < 7:
            messagebox.showwarning("Aviso", "Preencha a competência corretamente (Ex: 02/2026).")
            return
            
        if not self.selected_file_path:
            messagebox.showwarning("Aviso", "Selecione o arquivo da Relação de Empresas (.xlsx).")
            return
            
        # Desativa os botoes pra nao clonar execução
        self.btn_run.configure(state="disabled")
        self.btn_start_chrome.configure(state="disabled")
        self.status_var.set("Robô em Execução. Não mexa no navegador! Acompanhe o log pelo prompt se estiver aberto.")
        
        # Roda em thread paralela pro UI nao congelar
        thread = threading.Thread(target=self.run_bot_background, args=(comp, self.selected_file_path))
        thread.daemon = True
        thread.start()
        
    def run_bot_background(self, comp, filepath):
        try:
            # Chama a funcao principal que modificamos no nfse_bot.py
            nfse_bot.run_bot(comp, filepath, wait_for_input=False)
            messagebox.showinfo("Finalizado", "A extração foi concluída com sucesso! Verifique a pasta 'livros'.")
        except Exception as e:
            messagebox.showerror("Erro na Extração", f"O processo falhou inesperadamente:\n{e}")
        finally:
            self.btn_run.configure(state="normal")
            self.btn_start_chrome.configure(state="normal")
            self.status_var.set("Aguardando novo comando.")

if __name__ == "__main__":
    app = App()
    app.mainloop()
