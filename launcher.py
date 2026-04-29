import time
import sys
import os
import subprocess
import zipfile
import io
import threading
import tkinter as tk
from tkinter import messagebox, ttk
import requests
import psutil
from PIL import Image, ImageTk

# ----------------- CONFIGURAÇÃO -----------------
APP_EXE = "listagem_encomendas_app.exe"
URL_VERSAO = (
    "https://raw.githubusercontent.com/RodrigoBTX/Encomendas_updates/main/version.txt"
)
URL_RELEASE = (
    "https://github.com/RodrigoBTX/Encomendas_updates/releases/latest/download/app.zip"
)
VERSAO_LOCAL_FILE = "version.txt"
LOGO_PATH = "logo.ico"
# -------------------------------------------------


def resource_path(relative_path):
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)


def criar_splash(root):
    splash = tk.Toplevel(root)
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

    tk.Label(
        main_frame,
        text="Listagem de Encomendas",
        font=("Segoe UI", 14, "bold"),
        bg=cor_fundo,
        fg=cor_texto_principal,
    ).pack()

    status_label = tk.Label(
        main_frame,
        text="A verificar atualizações...",
        font=("Segoe UI", 9),
        bg=cor_fundo,
        fg=cor_texto_secundario,
    )
    status_label.pack(pady=(5, 20))

    style = ttk.Style()
    style.theme_use("clam")
    style.configure(
        "Modern.Horizontal.TProgressbar",
        troughcolor=cor_fundo,
        bordercolor=cor_fundo,
        background=cor_destaque,
        lightcolor=cor_destaque,
        darkcolor=cor_destaque,
        thickness=4,
    )

    progress = ttk.Progressbar(
        main_frame,
        orient="horizontal",
        mode="determinate",
        length=180,
        style="Modern.Horizontal.TProgressbar",
    )
    progress.pack()

    splash.progress = progress
    splash.status = status_label
    splash.update()
    return splash


def already_open():
    for p in psutil.process_iter(["name", "exe"]):
        try:
            if p.info["name"] == APP_EXE or (
                p.info["exe"] and os.path.basename(p.info["exe"]) == APP_EXE
            ):
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
    except Exception:
        return None


def update_splash(splash, valor, texto):
    splash.status.config(text=texto)
    splash.progress["value"] = valor
    splash.update()


def download_e_extrair(splash, callback_sucesso, callback_erro):
    """Corre numa thread separada. Atualiza o splash via root.after()"""

    def run():
        try:
            r = requests.get(URL_RELEASE, stream=True, timeout=60)
            r.raise_for_status()
            total_length = r.headers.get("content-length")

            data = b""
            if total_length is None:
                splash.after(0, lambda: update_splash(splash, 60, "A descarregar..."))
                data = r.content
            else:
                total_length = int(total_length)
                downloaded = 0
                chunk_size = 1024 * 1024  # 1 MB por vez

                for chunk in r.iter_content(chunk_size=chunk_size):
                    if chunk:
                        data += chunk
                        downloaded += len(chunk)
                        # Mapeia 0%→100% do download para 50%→90% da barra
                        barra_pct = 50 + int((downloaded / total_length) * 40)
                        download_pct = int((downloaded / total_length) * 100)
                        splash.after(
                            0,
                            lambda b=barra_pct, d=download_pct: update_splash(
                                splash, b, f"A descarregar... {d}%"
                            ),
                        )

            splash.after(0, lambda: update_splash(splash, 92, "A extrair ficheiros..."))
            z = zipfile.ZipFile(io.BytesIO(data))
            z.extractall(".")
            splash.after(0, callback_sucesso)

        except Exception as e:
            splash.after(0, lambda err=e: callback_erro(err))

    t = threading.Thread(target=run, daemon=True)
    t.start()


# ----------------- MAIN -----------------
def main():
    root = tk.Tk()
    root.withdraw()

    splash = criar_splash(root)

    def continuar_apos_download(versao_para_guardar=None):
        if versao_para_guardar:
            guardar_versao_local(versao_para_guardar)
        update_splash(splash, 100, "Tudo pronto! A iniciar aplicação...")
        root.after(500, arrancar)

    def arrancar():
        iniciar_app()

    def erro_download(e):
        splash.destroy()
        messagebox.showerror("Erro", f"Falha ao descarregar ficheiros:\n{e}")
        root.destroy()
        sys.exit(1)

    def passo1_verificar_processo():
        update_splash(splash, 20, "A verificar processos ativos...")
        root.after(50, passo2_verificar_exe)

    def passo2_verificar_exe():
        update_splash(splash, 40, "A validar ficheiros locais...")
        root.after(50, passo3_verificar_ficheiro)

    def passo3_verificar_ficheiro():
        if not os.path.exists(APP_EXE):
            update_splash(splash, 50, "A descarregar aplicação pela primeira vez...")

            def apos_primeiro_download():
                versao_remota = obter_versao_remota()
                if versao_remota:
                    guardar_versao_local(versao_remota)
                continuar_apos_download()

            download_e_extrair(
                splash,
                callback_sucesso=apos_primeiro_download,
                callback_erro=erro_download,
            )
        else:
            passo4_verificar_versao()

    def passo4_verificar_versao():
        update_splash(splash, 70, "A procurar atualizações no servidor...")
        root.after(50, passo5_comparar_versoes)

    def passo5_comparar_versoes():
        versao_local = ler_versao_local()
        versao_remota = obter_versao_remota()

        if versao_remota is None:
            messagebox.showwarning(
                "Aviso",
                "Não foi possível verificar atualizações. A iniciar assim mesmo...",
            )
            continuar_apos_download()
            return

        if versao_remota != versao_local:
            splash.attributes("-topmost", False)
            pergunta = messagebox.askyesno(
                "Atualização Disponível",
                f"Nova versão disponível: {versao_remota}\nVersão atual: {versao_local}\n\nDeseja atualizar agora?",
            )
            splash.attributes("-topmost", True)

            if pergunta:
                update_splash(splash, 80, "A descarregar nova versão...")
                download_e_extrair(
                    splash,
                    callback_sucesso=lambda: continuar_apos_download(versao_remota),
                    callback_erro=erro_download,
                )
                return

        continuar_apos_download()

    def iniciar_app():
        if getattr(sys, "frozen", False):
            exe_path = os.path.join(os.path.dirname(sys.executable), APP_EXE)
        else:
            exe_path = APP_EXE

        if not os.path.exists(exe_path):
            messagebox.showerror("Erro", f"Arquivo {APP_EXE} não encontrado!")
            sys.exit(1)

        try:
            update_splash(splash, 100, "A aguardar arranque da aplicação...")
            splash.update()
            processo = subprocess.Popen([exe_path])

            # Espera até 15 segundos que o processo arranque
            for _ in range(15):
                if processo.poll() is not None:
                    messagebox.showerror("Erro", f"{APP_EXE} terminou inesperadamente.")
                    break
                time.sleep(1)
                splash.update()

        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível iniciar a app:\n{e}")
        finally:
            splash.destroy()
            root.destroy()
            sys.exit(0)

    root.after(100, passo1_verificar_processo)
    root.mainloop()


main()
