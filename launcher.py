import psutil
import sys
import os
import subprocess
import requests
import zipfile
import io
import tkinter as tk
from tkinter import messagebox, ttk

# ----------------- CONFIGURAÇÃO -----------------
APP_EXE = "listagem_encomendas_app.exe"
URL_VERSAO = "https://raw.githubusercontent.com/RodrigoBTX/Encomendas_updates/main/version.txt"
URL_RELEASE = "https://github.com/RodrigoBTX/Encomendas_updates/releases/latest/download/app.zip"
VERSAO_LOCAL_FILE = "version.txt"
# -------------------------------------------------

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
        win.title("Atualizando...")
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

# ----------------- MAIN -----------------
root = tk.Tk()
root.withdraw()  # Oculta a janela principal do tkinter

if already_open():
    messagebox.showinfo("Info", f"O {APP_EXE} já se encontra aberto.")
    sys.exit(0)

# Se o exe não existir, download e extrai o zip
if not os.path.exists(APP_EXE):
    download_e_extrair_zip_com_progresso()

# Verifica atualizações se o exe existir
versao_local = ler_versao_local()
versao_remota = obter_versao_remota()

if versao_remota and versao_remota != versao_local:
    download_e_extrair_zip_com_progresso()
    guardar_versao_local(versao_remota)

# Inicia a aplicação principal
iniciar_app()
