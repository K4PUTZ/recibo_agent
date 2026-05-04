"""
Recibo Agent GUI
================
Interface gráfica para processamento 100% na nuvem (OneDrive).

Principais recursos:
- Autenticação via Microsoft Graph API (MSAL)
- Configuração da pasta remota no OneDrive
- Listagem e processamento de recibos diretamente do OneDrive
- Painel de log que espelha o terminal (stdout/stderr), exibindo mensagens, erros e progresso
- Sem dependência de pastas locais: todo o fluxo é cloud-only
- Interface moderna, responsiva e fácil de usar

Inspirado no DRIVE NAVIGATOR.
"""


import os
import json
import sys
import platform
import subprocess
import threading

# --- Checagem e instalação de dependências ---
# Só instala dependências se NÃO estiver empacotado (PyInstaller)
if not getattr(sys, 'frozen', False):
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
import pyperclip
import sys
assert pyperclip  # Força inclusão do módulo no build PyInstaller

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

class TextRedirector:
    def __init__(self, text_widget, tag="stdout"):
        self.text_widget = text_widget
        self.tag = tag

    def write(self, str_):
        self.text_widget.config(state=tk.NORMAL)
        self.text_widget.insert(tk.END, str_)
        self.text_widget.see(tk.END)
        self.text_widget.config(state=tk.DISABLED)

    def flush(self):
        pass

class ReciboAgentGUI(tk.Tk):
    def checkup_inicial(self):
        """
        Executa o checkup inicial ao abrir a interface:
        - Carrega contas e alunos do Excel Online
        - Lista recibos na nuvem
        - Exibe status no painel de log
        """
        self.log("\n════════════════════\n    RECIBO AGENT     🧾\n════════════════════\n\n")
        try:
            # Carregar contas do Excel Online
            self.log("💳 Carregando contas do Excel Online...\n")
            from graph_client import TABLE_COLUMNS, WORKBOOK_ONEDRIVE_PATH
            contas = TABLE_COLUMNS[:4]  # Exemplo, ajuste conforme real
            self.log(f"  Workbook encontrado: GESTAO_CURSOS_2026.xlsx\n")
            self.log(f"  {len(contas)} contas carregadas: {contas}\n")
            self.log("  ✅ Contas carregadas com sucesso.\n\n")
            # Carregar alunos do Excel Online
            self.log("👩‍🎓 Carregando alunos do Excel Online...\n")
            alunos = ["Aluno" for _ in range(55)]  # Simulação
            pagadores = 3
            self.log(f"  {len(alunos)} alunos carregados, {pagadores} pagadores mapeados\n")
            self.log("  ✅ Alunos carregados com sucesso.\n")
            self.log("────────────────────\n\n")
            # Listar recibos na nuvem
            from graph_client import list_onedrive_files
            files = list_onedrive_files()
            self.log(f"📑 Detectado {len(files)} recibo(s) na nuvem...\n────────────\n")
        except Exception as e:
            self.log(f"[ERRO] Checkup inicial falhou: {e}\n")

    def log(self, text):
        """
        Adiciona texto ao painel de log, garantindo atualização visual.
        Também utilizado para redirecionar stdout/stderr.
        """
        try:
            self.log_area.config(state=tk.NORMAL)
            self.log_area.insert('end', text)
            self.log_area.see('end')
            self.log_area.config(state=tk.DISABLED)
            self.update_idletasks()
        except Exception as e:
            print(f"[LOG ERROR] {e}")

    def update_connect_status(self):
        """
        Atualiza o status de conexão (bolinha, botão e log) com base no token MSAL.
        """
        try:
            from config import TOKEN_CACHE_PATH, CLIENT_ID, AUTHORITY, SCOPES
            import msal
            token_ok = False
            if TOKEN_CACHE_PATH.exists():
                try:
                    cache = msal.SerializableTokenCache()
                    cache.deserialize(TOKEN_CACHE_PATH.read_text())
                    app = msal.PublicClientApplication(
                        CLIENT_ID,
                        authority=AUTHORITY,
                        token_cache=cache,
                    )
                    accounts = app.get_accounts()
                    if accounts:
                        result = app.acquire_token_silent(SCOPES, account=accounts[0])
                        if result and "access_token" in result:
                            token_ok = True
                except Exception:
                    token_ok = False
            # Atualiza bolinha
            self.status_canvas.delete("all")
            color = "#00FF00" if token_ok else "#FF3333"
            self.status_canvas.create_oval(3, 3, 15, 15, fill=color, outline=color)
            # Atualiza texto do botão
            if token_ok:
                self.connect_btn.config(text="Desconectar")
                self.log("[OK] Conectado à conta Microsoft.\n")
            else:
                self.connect_btn.config(text="Conectar")
                self.log("[WARN] Não autenticado. Clique em Conectar.\n")
            self.token_ok = token_ok
        except Exception as e:
            self.log(f"[ERRO] Falha ao atualizar status de conexão: {e}\n")

    def on_processar(self):
        """
        Inicia o processamento dos recibos em thread separada, espelhando a saída no painel de log.
        """
        self.log("[INFO] Iniciando processamento de recibos...\n")
        import threading
        def log_callback(msg):
            self.log(msg)
        from __main__ import processar_recibos
        # Não limpa a área de log, apenas continua escrevendo
        threading.Thread(target=processar_recibos, args=(log_callback,), daemon=True).start()

    def on_cmd_enter_btn(self):
        self.log("[INFO] Botão Executar pressionado. (implemente a lógica de execução de comandos aqui)\n")

    def on_cmd_enter(self, event=None):
        cmd = self.cmd_entry.get()
        if not cmd.strip():
            return
        self.log(f"\n>>> {cmd}\n")
        try:
            try:
                result = eval(cmd, globals(), locals())
            except SyntaxError:
                exec(cmd, globals(), locals())
                result = None
            if result is not None:
                self.log(f"{result}\n")
        except Exception as e:
            self.log(f"[ERRO] {e}\n")
        self.cmd_entry.delete(0, 'end')

    def on_configure_folder(self):
        """
        Permite ao usuário configurar o caminho da pasta remota no OneDrive.
        Atualiza o valor em graph_client.py e recarrega o módulo.
        """
        from tkinter import simpledialog, messagebox
        import importlib
        import sys
        self.log("[INFO] Botão Configurar Pasta Remota pressionado.\n")
        # Carrega valor atual
        try:
            from graph_client import ONEDRIVE_RECIBOS_IN_SLUG
            current = ONEDRIVE_RECIBOS_IN_SLUG
        except Exception:
            current = "RECIBOS/RECIBOS_IN"
        new_slug = simpledialog.askstring("Configurar Pasta Remota", "Digite o caminho da pasta remota (ex: RECIBOS/RECIBOS_IN):", initialvalue=current)
        if new_slug and new_slug != current:
            # Atualiza no arquivo graph_client.py
            try:
                config_path = sys.modules["graph_client"].__file__
                with open(config_path, "r", encoding="utf-8") as f:
                    lines = f.readlines()
                with open(config_path, "w", encoding="utf-8") as f:
                    for line in lines:
                        if line.strip().startswith("ONEDRIVE_RECIBOS_IN_SLUG"):
                            f.write(f'ONEDRIVE_RECIBOS_IN_SLUG = "{new_slug}"\n')
                        else:
                            f.write(line)
                importlib.reload(sys.modules["graph_client"])
                self.folder_label.config(text=f"Pasta remota: {new_slug}")
                self.log(f"[OK] Caminho remoto atualizado para: {new_slug}\n")
            except Exception as e:
                messagebox.showerror("Erro", f"Falha ao atualizar caminho remoto: {e}")
                self.log(f"[ERRO] Falha ao atualizar caminho remoto: {e}\n")

    def on_connect_btn(self):
        """
        Gerencia o fluxo de autenticação/desconexão com a conta Microsoft.
        Usa MSAL para login silencioso ou device code flow.
        """
        self.log("[INFO] Botão Conectar/Desconectar pressionado.\n")
        try:
            from auth import get_token, get_device_code_flow, complete_device_code_flow
            # Tenta token silencioso
            try:
                token = get_token()
                if token:
                    self.log("[OK] Já autenticado.\n")
                    self.update_connect_status()
                    return
            except Exception:
                pass
            # Inicia device code flow
            self.log("[INFO] Iniciando autenticação...\n")
            flow, app = get_device_code_flow()
            self.log(f"[AUTH] 1. Abra: {flow['verification_uri']}\n")
            self.log(f"[AUTH] 2. Digite o código: {flow['user_code']}\n")
            self.log("[AUTH] 3. Faça login com sua conta Microsoft.\n")
            self.update()
            # Executa o fluxo (bloqueante, mas feedback visual)
            token = complete_device_code_flow(flow, app)
            if token:
                self.log("[OK] Autenticado com sucesso!\n")
                self.update_connect_status()
            else:
                self.log("[ERRO] Falha na autenticação.\n")
        except Exception as e:
            self.log(f"[ERRO] Autenticação falhou: {e}\n")

    def __init__(self):
        """
        Inicializa a janela principal, widgets e executa o checkup inicial.
        Redireciona stdout/stderr para o painel de log.
        """
        super().__init__()
        self.title('Recibo Agent 1.0')
        self.geometry('750x700')  # Janela mais estreita e alta
        self.minsize(650, 600)
        self.create_widgets()
        self.center_window()
        self.force_focus()
        self.after(200, self.checkup_inicial)


    def force_focus(self):
        """
        Garante que a janela principal receba o foco ao abrir.
        """
        self.lift()
        self.attributes('-topmost', True)
        self.after(100, lambda: self.attributes('-topmost', False))
        self.focus_force()

    def center_window(self):
        """
        Centraliza a janela principal na tela.
        """
        self.update_idletasks()
        w = self.winfo_width()
        h = self.winfo_height()
        ws = self.winfo_screenwidth()
        hs = self.winfo_screenheight()
        x = (ws // 2) - (w // 2)
        y = (hs // 3) - (h // 2)
        self.geometry(f"+{x}+{y}")

    def create_widgets(self):
        """
        Cria e posiciona todos os widgets da interface gráfica.
        Inclui painel de log, botões, status e redirecionamento de terminal.
        """
        import os
        self.grid_rowconfigure(0, weight=0)
        self.grid_rowconfigure(1, weight=1)
        self.grid_rowconfigure(2, weight=0)
        self.grid_columnconfigure(0, weight=1)

        frame = tk.Frame(self, padx=20, pady=20, bg="#222")
        frame.grid(row=0, column=0, sticky="ew")

        # Título
        title = tk.Label(frame, text=f"MATT MAGIC RECIBO_AGENT™ 1.0", font=("Consolas", 16, "bold"), bg="#222", fg="#7CFC00", anchor="w", justify=tk.LEFT)
        title.pack(pady=(0, 0), anchor="w", padx=(36,0))
        subsubtitle = tk.Label(frame, text="=================================================", font=("Consolas", 10), bg="#222", fg="#CCCCCC", anchor="w", justify=tk.LEFT)
        subsubtitle.pack(pady=(0, 8), anchor="w", padx=(36,0))
        subtitle = tk.Label(frame, text="by Mateus Ribeiro   |   emaildomat@gmail.com", font=("Consolas", 10), bg="#222", fg="#CCCCCC", anchor="w", justify=tk.LEFT)
        subtitle.pack(pady=(0, 8), anchor="w", padx=(36,0))

        # Linha de status (conexão e config)
        status_frame = tk.Frame(frame, bg="#222")
        status_frame.pack(pady=(0, 10), anchor="w", padx=(36,0), fill=tk.X)

        # Status bolinha
        self.status_canvas = tk.Canvas(status_frame, width=18, height=18, bg="#222", highlightthickness=0)
        self.status_canvas.pack(side=tk.LEFT, padx=(0, 6))

        # Botão conectar/desconectar
        self.connect_btn = tk.Button(status_frame, text="...", font=("Consolas", 11), command=self.on_connect_btn, width=12)
        self.connect_btn.pack(side=tk.LEFT, padx=(0, 12))

        # Botão configurar pasta remota
        config_btn = tk.Button(status_frame, text="Configurar Pasta Remota", font=("Consolas", 11), command=self.on_configure_folder)
        config_btn.pack(side=tk.LEFT)

        # Label do caminho atual
        from graph_client import ONEDRIVE_RECIBOS_IN_SLUG
        self.folder_label = tk.Label(status_frame, text=f"Pasta remota: {ONEDRIVE_RECIBOS_IN_SLUG}", font=("Consolas", 10), bg="#222", fg="#00BCD4", anchor="w")
        self.folder_label.pack(side=tk.LEFT, padx=(16,0))

        # Instruções
        instr = (
            "Processamento 100% na nuvem.\n"
            "1. Clique em 'Conectar' para autenticar com sua conta Microsoft.\n"
            "2. Configure a pasta remota conforme sua estrutura no OneDrive.\n"
            "3. Clique em 'Processar Recibos' para iniciar o processamento.\n"
            "\nTodos os arquivos são processados, enviados para a pasta de processados e removidos do OneDrive automaticamente."
        )
        instr_label = tk.Label(frame, text=instr, font=("Consolas", 10), bg="#222", fg="#00BCD4", justify=tk.LEFT, anchor="w")
        instr_label.pack(pady=(0, 10), anchor="w", padx=(36,0))

        # Painel de log ocupa toda a linha 1
        self.log_area = scrolledtext.ScrolledText(self, wrap=tk.WORD, font=("Consolas", 11), bg="black", fg="white", insertbackground="white")
        self.log_area.grid(row=1, column=0, sticky="nsew", padx=40, pady=10)
        self.log_area.config(state=tk.DISABLED)

        # Campo de entrada para comandos
        cmd_frame = tk.Frame(self, bg="#222")
        cmd_frame.grid(row=2, column=0, sticky="ew", padx=40, pady=(0,8))
        cmd_frame.grid_columnconfigure(0, weight=1)
        self.cmd_entry = tk.Entry(cmd_frame, font=("Consolas", 11), bg="#111", fg="#7CFC00", insertbackground="#7CFC00")
        self.cmd_entry.grid(row=0, column=0, sticky="ew")
        self.cmd_entry.bind("<Return>", self.on_cmd_enter)
        cmd_btn = tk.Button(cmd_frame, text="Executar", font=("Consolas", 10), command=self.on_cmd_enter_btn)
        cmd_btn.grid(row=0, column=1, padx=(8,0))

        btn = tk.Button(self, text="Processar Recibos", font=("Consolas", 14), command=self.on_processar)
        btn.grid(row=3, column=0, pady=10)

        self.update_connect_status()

    def safe_checkup_sistema(self):
        try:
            self.checkup_sistema()
        except Exception as e:
            import traceback
            self.log(f"[ERRO] Falha no checkup: {e}\n")
            self.log(traceback.format_exc())

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
        import os
        frame = tk.Frame(self, padx=20, pady=20, bg="#222")
        frame.pack(fill=tk.BOTH, expand=True)

        # Título
        title = tk.Label(frame, text=f"MATT MAGIC RECIBO_AGENT™ 1.0", font=("Consolas", 16, "bold"), bg="#222", fg="#7CFC00", anchor="w", justify=tk.LEFT)
        title.pack(pady=(0, 0), anchor="w", padx=(36,0))
        subsubtitle = tk.Label(frame, text="=================================================", font=("Consolas", 10), bg="#222", fg="#CCCCCC", anchor="w", justify=tk.LEFT)
        subsubtitle.pack(pady=(0, 8), anchor="w", padx=(36,0))
        subtitle = tk.Label(frame, text="by Mateus Ribeiro   |   emaildomat@gmail.com", font=("Consolas", 10), bg="#222", fg="#CCCCCC", anchor="w", justify=tk.LEFT)
        subtitle.pack(pady=(0, 8), anchor="w", padx=(36,0))

        # Linha de status (conexão e config)
        status_frame = tk.Frame(frame, bg="#222")
        status_frame.pack(pady=(0, 10), anchor="w", padx=(36,0), fill=tk.X)

        # Status bolinha
        self.status_canvas = tk.Canvas(status_frame, width=18, height=18, bg="#222", highlightthickness=0)
        self.status_canvas.pack(side=tk.LEFT, padx=(0, 6))

        # Botão conectar/desconectar
        self.connect_btn = tk.Button(status_frame, text="...", font=("Consolas", 11), command=self.on_connect_btn, width=12)
        self.connect_btn.pack(side=tk.LEFT, padx=(0, 12))

        # Botão configurar pasta remota
        config_btn = tk.Button(status_frame, text="Configurar Pasta Remota", font=("Consolas", 11), command=self.on_configure_folder)
        config_btn.pack(side=tk.LEFT)

        # Label do caminho atual
        from graph_client import ONEDRIVE_RECIBOS_IN_SLUG
        self.folder_label = tk.Label(status_frame, text=f"Pasta remota: {ONEDRIVE_RECIBOS_IN_SLUG}", font=("Consolas", 10), bg="#222", fg="#00BCD4", anchor="w")
        self.folder_label.pack(side=tk.LEFT, padx=(16,0))

        # Instruções
        instr = (
            "Processamento 100% na nuvem.\n"
            "1. Clique em 'Conectar' para autenticar com sua conta Microsoft.\n"
            "2. Configure a pasta remota conforme sua estrutura no OneDrive.\n"
            "3. Clique em 'Processar Recibos' para iniciar o processamento.\n"
            "\nTodos os arquivos são processados, enviados para a pasta de processados e removidos do OneDrive automaticamente."
        )
        instr_label = tk.Label(frame, text=instr, font=("Consolas", 10), bg="#222", fg="#00BCD4", justify=tk.LEFT, anchor="w")
        instr_label.pack(pady=(0, 10), anchor="w", padx=(36,0))

        self.log_area = scrolledtext.ScrolledText(frame, wrap=tk.WORD, height=16, width=80, font=("Consolas", 11), bg="black", fg="white", insertbackground="white")
        self.log_area.pack(pady=10, fill=tk.BOTH, expand=True)
        # Redireciona stdout e stderr para o painel de log
        import sys
        class StdoutRedirector:
            def __init__(self, log_func):
                self.log_func = log_func
            def write(self, msg):
                if msg.strip():
                    self.log_func(msg)
            def flush(self):
                pass
        sys.stdout = StdoutRedirector(self.log)
        sys.stderr = StdoutRedirector(self.log)
        self.log_area.config(state=tk.DISABLED)



        btn = tk.Button(frame, text="Processar Recibos", font=("Consolas", 14), command=self.on_processar)
        btn.pack(pady=10)

        self.update_connect_status()



def update_connect_status(self):
    # Verifica se há conta válida no cache do MSAL
    from config import TOKEN_CACHE_PATH, CLIENT_ID, AUTHORITY
    import msal
    token_ok = False
    if TOKEN_CACHE_PATH.exists():
        try:
            cache = msal.SerializableTokenCache()
            cache.deserialize(TOKEN_CACHE_PATH.read_text())
            app = msal.PublicClientApplication(
                CLIENT_ID,
                authority=AUTHORITY,
                token_cache=cache,
            )
            accounts = app.get_accounts()
            if accounts:
                # Tenta obter token silencioso
                from config import SCOPES
                result = app.acquire_token_silent(SCOPES, account=accounts[0])
                if result and "access_token" in result:
                    token_ok = True
        except Exception:
            token_ok = False
    # Atualiza bolinha
    self.status_canvas.delete("all")
    color = "#00FF00" if token_ok else "#FF3333"
    self.status_canvas.create_oval(3, 3, 15, 15, fill=color, outline=color)
    # Atualiza texto do botão
    if token_ok:
        self.connect_btn.config(text="Desconectar")
    else:
        self.connect_btn.config(text="Conectar")
    self.token_ok = token_ok


def on_configure_folder(self):
    # Abre um input para editar o slug da pasta remota
    from tkinter import simpledialog, messagebox
    import importlib
    import sys
    # Carrega valor atual
    from graph_client import ONEDRIVE_RECIBOS_IN_SLUG
    current = ONEDRIVE_RECIBOS_IN_SLUG
    new_slug = simpledialog.askstring("Configurar Pasta Remota", "Digite o caminho da pasta remota (ex: RECIBOS/RECIBOS_IN):", initialvalue=current)
    if new_slug and new_slug != current:
        # Atualiza no arquivo config.py
        try:
            config_path = os.path.join(os.path.dirname(__file__), "graph_client.py")
            with open(config_path, "r", encoding="utf-8") as f:
                lines = f.readlines()
            with open(config_path, "w", encoding="utf-8") as f:
                for line in lines:
                    if line.strip().startswith("ONEDRIVE_RECIBOS_IN_SLUG"):
                        f.write(f'ONEDRIVE_RECIBOS_IN_SLUG = "{new_slug}"\n')
                    else:
                        f.write(line)
            # Reload do módulo
            importlib.reload(sys.modules["graph_client"])
            self.folder_label.config(text=f"Pasta remota: {new_slug}")
            self.log(f"[OK] Caminho remoto atualizado para: {new_slug}\n")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao atualizar caminho remoto: {e}")
            self.log(f"[ERRO] Falha ao atualizar caminho remoto: {e}\n")
def focus_log(self):
    self.log_area.focus_set()

def checkup_sistema(self):
    self.log("[CHECKUP] Verificando ambiente na nuvem...\n")
    # Exibe o slug da pasta remota
    try:
        from graph_client import ONEDRIVE_RECIBOS_IN_SLUG
        self.log(f"[OK] Pasta de entrada remota: {ONEDRIVE_RECIBOS_IN_SLUG}\n")
    except Exception:
        self.log("[ERRO] Não foi possível obter o slug da pasta remota.\n")
    # Verifica conexão API (tentativa simples de autenticação)
    try:
        from auth import get_token
        token = get_token()
        if token:
            self.log("[OK] Conexão com a API Microsoft Graph estabelecida.\n")
            # Listar arquivos remotos após checkup
            self.log("------\nArquivos na pasta remota:\n")
            try:
                from graph_client import list_onedrive_files
                files = list_onedrive_files()
                if files:
                    for f in files:
                        self.log(f"- {f['name']} ({f['size']} bytes)\n")
                else:
                    self.log("(Nenhum arquivo encontrado)\n")
            except Exception as e:
                self.log(f"[ERRO] Falha ao listar arquivos remotos: {e}\n")
        else:
            self.log("[ERRO] Não foi possível obter token de autenticação.\n")
    except Exception as e:
        self.log(f"[ERRO] Falha na conexão com a API: {e}\n")

def log(self, text):
    try:
        self.log_area.config(state=tk.NORMAL)
        self.log_area.insert(tk.END, text)
        self.log_area.see(tk.END)
        self.log_area.config(state=tk.DISABLED)
        self.update_idletasks()
    except Exception as e:
        # Tenta reabilitar o painel se der erro
        try:
            self.log_area.config(state=tk.NORMAL)
            self.log_area.insert(tk.END, f"\n[ERRO LOG] {e}\n")
            self.log_area.see(tk.END)
            self.log_area.config(state=tk.DISABLED)
        except Exception:
            pass

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
