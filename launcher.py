import psutil
import sys
import os
import subprocess
import requests
import zipfile
import io
import tkinter as tk
from tkinter import messagebox, ttk
from PIL import Image, ImageTk

# ----------------- CONFIGURAÇÃO -----------------
APP_EXE = "listagem_encomendas_app.exe"
URL_VERSAO = "https://raw.githubusercontent.com/RodrigoBTX/Encomendas_updates/main/version.txt"
URL_RELEASE = "https://github.com/RodrigoBTX/Encomendas_updates/releases/latest/download/app.zip"
VERSAO_LOCAL_FILE = "version.txt"
LOGO_PATH = "logo.ico"
# -------------------------------------------------


def resource_path(relative_path):
    
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

def criar_splash():
    splash = tk.Toplevel()
    splash.overrideredirect(True)
    splash.attributes("-topmost", True)
    
    
    largura, altura = 320, 240 
    screen_width = splash.winfo_screenwidth()
    screen_height = splash.winfo_screenheight()
    x = (screen_width // 2) - (largura // 2)
    y = (screen_height // 2) - (altura // 2)
    splash.geometry(f"{largura}x{altura}+{x}+{y}")

    
    cor_fundo = "#121212" 
    cor_destaque = "#3498db" 
    cor_texto_principal = "#ffffff"
    cor_texto_secundario = "#aaaaaa"
    
    splash.configure(bg=cor_fundo)

    # Container para margens internas
    main_frame = tk.Frame(splash, bg=cor_fundo, padx=20, pady=20)
    main_frame.pack(expand=True, fill="both")

    try:
        img_path = resource_path(LOGO_PATH)
        if os.path.exists(img_path):
            img = Image.open(img_path)           
            img = img.convert("RGBA") 
            img.thumbnail((80, 80), Image.LANCZOS) 
            photo = ImageTk.PhotoImage(img)
            
            label_img = tk.Label(main_frame, image=photo, bg=cor_fundo)
            label_img.image = photo 
            label_img.pack(pady=(10, 15))
    except Exception:        
        pass

    # Texto Principal
    tk.Label(main_frame, text="Listagem de Encomendas", 
             font=("Segoe UI", 14, "bold"), 
             bg=cor_fundo, fg=cor_texto_principal).pack()

    # Texto de Status 
    status_label = tk.Label(main_frame, text="A verificar atualizações...", 
                            font=("Segoe UI", 9), 
                            bg=cor_fundo, fg=cor_texto_secundario)
    status_label.pack(pady=(5, 20))
    
    style = ttk.Style()
    style.theme_use('clam') 
    style.configure("Modern.Horizontal.TProgressbar", 
                    troughcolor=cor_fundo, 
                    bordercolor=cor_fundo, 
                    background=cor_destaque, 
                    lightcolor=cor_destaque, 
                    darkcolor=cor_destaque,
                    thickness=4) 

    progress = ttk.Progressbar(main_frame, orient="horizontal", mode="determinate", 
                                length=180, style="Modern.Horizontal.TProgressbar")
    progress.pack()
    
    # Criamos uma referência global ou passamos o objeto para atualizar
    splash.progress = progress
    splash.status = status_label

    splash.update() 
    return splash



def already_open():
    for p in psutil.process_iter(['name', 'exe']):
        try:
            if p.info['name'] == APP_EXE or (p.info['exe'] and os.path.basename(p.info['exe']) == APP_EXE):
                return True
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
    return False

def ler_versao_local():
    if os.path.exists(VERSAO_LOCAL_FILE):
        with open(VERSAO_LOCAL_FILE, "r") as f:
            return f.read().strip()
    return "0.0.0"

def guardar_versao_local(versao):
    with open(VERSAO_LOCAL_FILE, "w") as f:
        f.write(versao)

def obter_versao_remota():
    try:
        r = requests.get(URL_VERSAO, timeout=10)
        r.raise_for_status()
        return r.text.strip()
    except Exception as e:
        messagebox.showerror("Erro", f"Não foi possível verificar atualizações:\n{e}")
        return None

def download_e_extrair_zip_com_progresso():
    """Download do zip e mostra uma janela com barra de progresso"""
    try:
        # Cria a janela de progresso
        win = tk.Toplevel()
        win.title("A Atualizar...")
        win.geometry("400x100")
        win.resizable(False, False)
        tk.Label(win, text="Download dos arquivos, aguarde...").pack(pady=10)
        progress = ttk.Progressbar(win, orient="horizontal", length=350, mode="determinate")
        progress.pack(pady=10)
        win.update()

        # Faz o download em stream
        r = requests.get(URL_RELEASE, stream=True, timeout=30)
        r.raise_for_status()
        total_length = r.headers.get('content-length')

        if total_length is None:
            # Sem tamanho conhecido
            data = r.content
            progress['value'] = 100
        else:
            total_length = int(total_length)
            data = b""
            chunk_size = 1024*1024  # 1 MB por vez
            downloaded = 0
            for chunk in r.iter_content(chunk_size=chunk_size):
                if chunk:
                    data += chunk
                    downloaded += len(chunk)
                    progress['value'] = (downloaded / total_length) * 100
                    win.update()

        # Extrai zip
        z = zipfile.ZipFile(io.BytesIO(data))
        z.extractall(".")
        win.destroy()
        messagebox.showinfo("Atualização", "Download dos arquivos com sucesso!")

    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao fazer download dos arquivos:\n{e}")
        sys.exit(1)

def iniciar_app():
    if getattr(sys, 'frozen', False):
        exe_path = os.path.join(os.path.dirname(sys.executable), APP_EXE)
    else:
        exe_path = APP_EXE

    if not os.path.exists(exe_path):
        messagebox.showerror("Erro", f"Arquivo {APP_EXE} não encontrado!")
        sys.exit(1)

    try:
        subprocess.Popen([exe_path])
    except Exception as e:
        messagebox.showerror("Erro", f"Não foi possível iniciar a app:\n{e}")
    finally:
        sys.exit(0)

# Função auxiliar para atualizar a barra suavemente
def update_splash(splash, valor, texto):
    splash.status.config(text=texto)
    splash.progress['value'] = valor
    splash.update()        

# ----------------- MAIN -----------------
root = tk.Tk()
root.withdraw()  # Oculta a janela principal do tkinter

splash = criar_splash()

update_splash(splash, 20, "A verificar processos ativos...")
if already_open():
    splash.destroy()
    messagebox.showinfo("Info", f"O {APP_EXE} já se encontra aberto.")
    sys.exit(0)

# Se o exe não existir, download e extrai o zip
update_splash(splash, 40, "A validar ficheiros locais...")
if not os.path.exists(APP_EXE):
    splash.withdraw()
    download_e_extrair_zip_com_progresso()
    splash.deiconify()

# Verifica atualizações se o exe existir
update_splash(splash, 70, "A procurar atualizações no servidor...")
versao_local = ler_versao_local()
versao_remota = obter_versao_remota()

if versao_remota and versao_remota != versao_local:
    update_splash(splash, 85, "Nova versão detetada! A preparar download...")
    splash.withdraw()
    download_e_extrair_zip_com_progresso()
    guardar_versao_local(versao_remota)
    splash.deiconify()


update_splash(splash, 100, "Tudo pronto! A iniciar aplicação...")
# Inicia a aplicação principal
splash.destroy()
iniciar_app()
