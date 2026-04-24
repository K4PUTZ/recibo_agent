"""
Recibo Agent - Interface Gráfica (Tkinter)
Inspirado no DRIVE NAVIGATOR
"""

import os
import sys
import platform
import subprocess
import threading

# --- Checagem e instalação de dependências ---
def ensure_dependencies():
    import importlib
    import subprocess
    import sys
    required = [
        ("tkinter", "tk"),
        ("requests", "requests"),
        ("PIL", "pillow"),
        ("colorama", "colorama"),
    ]
    for module, pip_name in required:
        try:
            importlib.import_module(module)
        except ImportError:
            print(f"Instalando dependência: {pip_name}")
            subprocess.check_call([sys.executable, "-m", "pip", "install", pip_name])

ensure_dependencies()

import tkinter as tk
from tkinter import messagebox, scrolledtext

VERSION = "1.0"

# --- Sons de notificação ---
def play_sound(sound_type):
    sounds = {
        "success": {
            "Darwin": "/System/Library/Sounds/Glass.aiff",
            "Windows": "SystemAsterisk",
            "Linux": os.path.expanduser("~/sounds/success.wav")
        },
        "cancel": {
            "Darwin": "/System/Library/Sounds/Funk.aiff",
            "Windows": "SystemExclamation",
            "Linux": os.path.expanduser("~/sounds/cancel.wav")
        }
    }
    sound_path = sounds[sound_type].get(platform.system())
    if sound_path:
        if platform.system() == "Darwin":
            subprocess.run(["afplay", sound_path])
        elif platform.system() == "Windows":
            import winsound
            winsound.PlaySound(sound_path, winsound.SND_ALIAS)
        elif platform.system() == "Linux":
            os.system(f'paplay {sound_path}')
        else:
            os.system('echo -e "\a"')
    else:
        os.system('echo -e "\a"')

def play_success():
    play_sound("success")

def play_cancel():
    play_sound("cancel")

# --- Notificações ---
def show_notification(message, title="Recibo Agent"):
    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo(title, message)
    root.destroy()

# --- Função para processar recibos ---
def processar_recibos(log_callback):
    try:
        # Chama o Recibo Agent em modo --once
        # Ajuste o caminho conforme necessário
        cmd = [sys.executable, os.path.join(os.path.dirname(__file__), "run.py"), "--once"]
        proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
        for line in proc.stdout:
            log_callback(line)
        proc.wait()
        if proc.returncode == 0:
            play_success()
            log_callback("\nProcessamento concluído com sucesso!\n")
        else:
            play_cancel()
            log_callback(f"\nErro no processamento (código {proc.returncode})\n")
    except Exception as e:
        play_cancel()
        log_callback(f"\nErro: {e}\n")

# --- Interface Tkinter ---

class ReciboAgentGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"Recibo Agent - v{VERSION}")
        self.geometry("700x500")
        self.resizable(False, False)
        self.center_window()
        self.create_widgets()
        self.after(100, self.force_focus)
        self.after(200, self.focus_log)
        self.after(300, self.checkup_sistema)

    def force_focus(self):
        self.lift()
        self.attributes('-topmost', True)
        self.after(100, lambda: self.attributes('-topmost', False))
        self.focus_force()

    def center_window(self):
        self.update_idletasks()
        w = self.winfo_width()
        h = self.winfo_height()
        ws = self.winfo_screenwidth()
        hs = self.winfo_screenheight()
        x = (ws // 2) - (w // 2)
        y = (hs // 3) - (h // 2)
        self.geometry(f"+{x}+{y}")

    def create_widgets(self):
        frame = tk.Frame(self, padx=20, pady=20, bg="#222")
        frame.pack(fill=tk.BOTH, expand=True)

        # Título (alinhado à esquerda)
        title = tk.Label(frame, text=f"MATT MAGIC RECIBO_AGENT™ 1.0", font=("Arial", 16, "bold"), bg="#222", fg="#7CFC00", anchor="w", justify=tk.LEFT)
        title.pack(pady=(0, 0), anchor="w", padx=(36,0))
        subsubtitle = tk.Label(frame, text="=================================================", font=("Arial", 10), bg="#222", fg="#CCCCCC", anchor="w", justify=tk.LEFT)
        subsubtitle.pack(pady=(0, 8), anchor="w", padx=(36,0))
        subtitle = tk.Label(frame, text="by Mateus Ribeiro   |   emaildomat@gmail.com", font=("Arial", 10), bg="#222", fg="#CCCCCC", anchor="w", justify=tk.LEFT)
        subtitle.pack(pady=(0, 8), anchor="w", padx=(36,0))

        # Instruções (alinhadas à esquerda)
        instr = (
            "Antes de rodar o programa, verifique se a pasta RECIBOS IN existe no OneDrive.\n"
            "Adicione -c /CAMINHO/RECIBOS IN na linha de comando para configurar a pasta padrão de armazenamento.\n"
            "Adicione -r na linha de comando para fazer login no drive remoto.\n"
            f"Pasta atual: Meus Arquivos/RECIBOS/RECIBOS_IN\n"
        )
        instr_label = tk.Label(frame, text=instr, font=("Consolas", 10), bg="#222", fg="#00BCD4", justify=tk.LEFT, anchor="w")
        instr_label.pack(pady=(0, 10), anchor="w", padx=(36,0))

        self.log_area = scrolledtext.ScrolledText(frame, wrap=tk.WORD, height=15, width=80, font=("Consolas", 11), bg="black", fg="white", insertbackground="white")
        self.log_area.pack(pady=10)
        self.log_area.config(state=tk.DISABLED)

        btn = tk.Button(frame, text="Processar Recibos", font=("Arial", 14), command=self.on_processar)
        btn.pack(pady=10)

    def focus_log(self):
        self.log_area.focus_set()

    def checkup_sistema(self):
        self.log("[CHECKUP] Verificando ambiente...\n")
        # Verifica pasta IN
        from config import WATCH_FOLDER
        if not os.path.exists(WATCH_FOLDER):
            self.log(f"[ERRO] Pasta de entrada não encontrada: {WATCH_FOLDER}\n")
        else:
            self.log(f"[OK] Pasta de entrada encontrada: {WATCH_FOLDER}\n")
        # Verifica conexão API (tentativa simples de autenticação)
        try:
            from auth import get_token
            token = get_token()
            if token:
                self.log("[OK] Conexão com a API Microsoft Graph estabelecida.\n")
            else:
                self.log("[ERRO] Não foi possível obter token de autenticação.\n")
        except Exception as e:
            self.log(f"[ERRO] Falha na conexão com a API: {e}\n")

    def log(self, text):
        self.log_area.config(state=tk.NORMAL)
        self.log_area.insert(tk.END, text)
        self.log_area.see(tk.END)
        self.log_area.config(state=tk.DISABLED)
        self.update_idletasks()

    def on_processar(self):
        self.log("\nIniciando processamento...\n")
        threading.Thread(target=processar_recibos, args=(self.log,), daemon=True).start()

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Recibo Agent - Interface Gráfica ou Processamento Direto")
    parser.add_argument("--cli", action="store_true", help="Executa o processamento direto (run.py --once) sem interface gráfica")
    args = parser.parse_args()

    if args.cli:
        # No modo CLI, pode imprimir o cabeçalho colorido no terminal se quiser
        import colorama
        from colorama import Fore, Back, Style
        colorama.init(autoreset=True)
        CURRENT_FOLDER_PATH = os.path.join("Meus Arquivos", "RECIBOS", "RECIBOS_IN")
        print("\n" * 3)
        os.system("clear" if os.name != "nt" else "cls")
        print("   " + Back.YELLOW + "  " + Back.CYAN + "  " + Back.GREEN + "  " + Back.MAGENTA + "  " + Back.RED + "  " + Back.BLUE + "  " + Back.RESET + " ")
        print(Back.RESET + Fore.GREEN + "   " + Back.RESET + "-----------------------")
        print("   MATT MAGIC RECIBO_AGENT™ 1.0")
        print("   by Mateus Ribeiro")
        print("   emaildomat@gmail.com")
        print("\n")
        print("   Antes de rodar o programa, verifique se a pasta " + Fore.LIGHTCYAN_EX + "RECIBOS IN " + Fore.GREEN + "existe no OneDrive.")
        print("   Adicione " + Fore.LIGHTCYAN_EX + "-c /CAMINHO/RECIBOS IN" + Fore.GREEN + " na linha de comando para configurar a pasta padrão de armazenamento.")
        print("   Adicione " + Fore.LIGHTCYAN_EX + "-r " + Fore.GREEN + "na linha de comando para fazer login no drive remoto.")
        print("   Pasta atual: " + Fore.LIGHTCYAN_EX + CURRENT_FOLDER_PATH)
        print("   ")
        print("\n")
        print(Style.RESET_ALL, end='')
        def log_print(text):
            print(text, end="")
        processar_recibos(log_print)
    else:
        app = ReciboAgentGUI()
        app.mainloop()
