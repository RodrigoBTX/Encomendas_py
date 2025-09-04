import requests, zipfile, io, os, subprocess, sys
import tkinter as tk
from tkinter import ttk, messagebox

# --- Configurações do repositório ---
URL_VERSAO = "https://raw.githubusercontent.com/RodrigoBTX/Encomendas_updates/main/version.txt"
URL_RELEASE = "https://github.com/RodrigoBTX/Encomendas_updates/releases/latest/download/app.zip"

# --- Funções ---
def ler_versao_local():
    if os.path.exists("version.txt"):
        with open("version.txt") as f:
            return f.read().strip()
    return "0.0.0"

def guardar_versao_local(versao):
    with open("version.txt", "w") as f:
        f.write(versao)

def obter_versao_remota():
    try:
        return requests.get(URL_VERSAO, timeout=5).text.strip()
    except Exception as e:
        messagebox.showerror("Erro", f"Não foi possível verificar atualizações:\n{e}")
        return None

def atualizar_app(prog_win):
    try:
        prog_win.label_var.set("Nova versão encontrada. A atualizar...")
        prog_win.update_idletasks()

        r = requests.get(URL_RELEASE, timeout=30)
        z = zipfile.ZipFile(io.BytesIO(r.content))
        z.extractall(".")  # extrai na pasta atual

        prog_win.label_var.set("Atualização concluída!")
        prog_win.update_idletasks()
        prog_win.after(1000, prog_win.destroy)  # fecha popup após 1s
    except Exception as e:
        prog_win.destroy()
        messagebox.showerror("Erro", f"Falha ao atualizar a app:\n{e}")

def iniciar_app():
    if getattr(sys, 'frozen', False):
        exe_path = os.path.join(os.path.dirname(sys.executable), "app.exe")
    else:
        exe_path = "app.exe"

    if not os.path.exists(exe_path):
        messagebox.showerror("Erro", f"O executável principal não foi encontrado:\n{exe_path}")
        return
    
    subprocess.Popen([exe_path])
    sys.exit()  # fecha o updater

# --- Popup de progresso ---
class ProgressWindow(tk.Toplevel):
    def __init__(self, root):
        super().__init__(root)
        self.title("Atualização")
        self.geometry("300x100")
        self.resizable(False, False)
        self.label_var = tk.StringVar()
        self.label_var.set("A verificar atualizações...")
        ttk.Label(self, textvariable=self.label_var).pack(pady=20)
        self.update_idletasks()

# --- Main ---
if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()  # esconde a janela principal

    versao_local = ler_versao_local()
    versao_remota = obter_versao_remota()

    if versao_remota and versao_remota != versao_local:
        prog_win = ProgressWindow(root)
        root.update()  # força a janela aparecer
        atualizar_app(prog_win)
        guardar_versao_local(versao_remota)

    iniciar_app()
