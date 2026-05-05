"""
Recibo Agent GUI
================
Interface gráfica para processamento 100% na nuvem (OneDrive).

Principais recursos:
- Autenticação via Microsoft Graph API (MSAL) com device code flow
- Checkbox "Conectar automaticamente" com preferência persistida
- Painel de log que espelha o terminal
- Barra de progresso + contador regressivo durante autenticação
- Sem dependência de pastas locais: todo o fluxo é cloud-only
"""

import json
import os
import platform
import subprocess
import sys
import threading
import tkinter as tk
import tkinter.ttk as ttk
from tkinter import messagebox, scrolledtext

import pyperclip
assert pyperclip  # força inclusão no build PyInstaller

VERSION = "1.0"
AUTH_TIMEOUT = 900  # segundos (15 min — padrão Microsoft device flow)


# ── Sons de notificação ──

def play_sound(sound_type):
    sounds = {
        "success": {"Darwin": "/System/Library/Sounds/Glass.aiff", "Windows": "SystemAsterisk"},
        "cancel":  {"Darwin": "/System/Library/Sounds/Funk.aiff",  "Windows": "SystemExclamation"},
    }
    sound_path = sounds.get(sound_type, {}).get(platform.system())
    if sound_path and platform.system() == "Darwin":
        subprocess.run(["afplay", sound_path], check=False)
    elif sound_path and platform.system() == "Windows":
        import winsound
        winsound.PlaySound(sound_path, winsound.SND_ALIAS)
    else:
        os.system('echo -e "\\a"')

def play_success(): play_sound("success")
def play_cancel():  play_sound("cancel")


# ── Processamento ──

def processar_recibos(log_callback):
    try:
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


# ── Preferências GUI ──

def _load_prefs() -> dict:
    from config import GUI_PREFS_PATH
    try:
        if GUI_PREFS_PATH.exists():
            return json.loads(GUI_PREFS_PATH.read_text())
    except Exception:
        pass
    return {"auto_connect": False}

def _save_prefs(prefs: dict):
    from config import GUI_PREFS_PATH
    try:
        GUI_PREFS_PATH.write_text(json.dumps(prefs, indent=2))
    except Exception:
        pass


# ── Interface Tkinter ──

class ReciboAgentGUI(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("Recibo Agent 1.0")
        self.geometry("750x700")
        self.minsize(650, 600)
        self.configure(bg="#1a1a1a")

        self.grid_rowconfigure(0, weight=0)
        self.grid_rowconfigure(1, weight=1)
        self.grid_rowconfigure(2, weight=0)
        self.grid_columnconfigure(0, weight=1)

        # Estado interno
        self.auto_connect_var = tk.BooleanVar(value=False)
        self._auth_thread: threading.Thread | None = None
        self._countdown_id = None   # after() handle do contador
        self._countdown_secs = 0
        self._auth_continue_event: threading.Event | None = None
        self._auth_cancelled = False
        self._auth_hint_photo = None  # mantém referência para evitar GC da imagem

        # Carregar preferências e definir checkbox antes de criar widgets
        prefs = _load_prefs()
        self.auto_connect_var.set(prefs.get("auto_connect", False))

        self.create_widgets()
        self.center_window()
        self.force_focus()

        # Startup: checagem rápida (main thread, nunca bloqueia)
        self.after(150, self.checkup_inicial)

    # ── Janela ──

    def force_focus(self):
        self.lift()
        self.attributes("-topmost", True)
        self.after(100, lambda: self.attributes("-topmost", False))
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

    # ── Widgets ──

    def create_widgets(self):
        # ── Cabeçalho ──
        header = tk.Frame(self, padx=20, pady=16, bg="#222")
        header.grid(row=0, column=0, sticky="ew")

        tk.Label(header, text="MATT MAGIC RECIBO_AGENT™ 1.0",
                 font=("Consolas", 16, "bold"), bg="#222", fg="#7CFC00",
                 anchor="w").pack(anchor="w", padx=(36, 0))
        tk.Label(header, text="=" * 49,
                 font=("Consolas", 10), bg="#222", fg="#555",
                 anchor="w").pack(anchor="w", padx=(36, 0))
        tk.Label(header, text="by Mateus Ribeiro   |   emaildomat@gmail.com",
                 font=("Consolas", 10), bg="#222", fg="#888",
                 anchor="w").pack(pady=(0, 10), anchor="w", padx=(36, 0))

        # ── Linha de status ──
        status_row = tk.Frame(header, bg="#222")
        status_row.pack(anchor="w", padx=(36, 0), fill=tk.X)

        # Indicador de conexão (bolinha)
        self.status_canvas = tk.Canvas(status_row, width=18, height=18,
                                       bg="#222", highlightthickness=0)
        self.status_canvas.pack(side=tk.LEFT, padx=(0, 6))

        # Botão Conectar / Desconectar
        self.connect_btn = tk.Button(status_row, text="...", font=("Consolas", 11),
                                     command=self.on_connect_btn, width=14)
        self.connect_btn.pack(side=tk.LEFT, padx=(0, 10))

        # Checkbox auto-connect
        self.auto_connect_chk = tk.Checkbutton(
            status_row,
            text="Conectar automaticamente",
            variable=self.auto_connect_var,
            font=("Consolas", 10),
            bg="#222", fg="#aaa",
            activebackground="#222", activeforeground="#7CFC00",
            selectcolor="#333",
            command=self.on_auto_connect_toggle,
        )
        self.auto_connect_chk.pack(side=tk.LEFT, padx=(0, 16))

        # ── Linha da pasta remota ──
        folder_row = tk.Frame(header, bg="#222")
        folder_row.pack(anchor="w", padx=(36, 0), fill=tk.X, pady=(4, 0))

        self.folder_btn = tk.Button(folder_row, text="Pasta Remota...", font=("Consolas", 10),
                  command=self.on_configure_folder)
        self.folder_btn.pack(side=tk.LEFT)

        # Label pasta atual
        from graph_client import ONEDRIVE_RECIBOS_IN_SLUG
        self.folder_label = tk.Label(folder_row,
                                     text=f"  {ONEDRIVE_RECIBOS_IN_SLUG}",
                                     font=("Consolas", 10), bg="#222", fg="#00BCD4", anchor="w")
        self.folder_label.pack(side=tk.LEFT)

        # ── Barra de progresso de auth (oculta inicialmente) ──
        self.auth_progress_frame = tk.Frame(header, bg="#222")
        self.auth_progress_frame.pack(anchor="w", padx=(36, 36), fill=tk.X, pady=(8, 0))
        self.auth_progress_frame.pack_forget()  # esconde até precisar

        self.auth_progressbar = ttk.Progressbar(
            self.auth_progress_frame, mode="indeterminate", length=400
        )
        self.auth_progressbar.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.auth_countdown_label = tk.Label(
            self.auth_progress_frame, text="", font=("Consolas", 10),
            bg="#222", fg="#FFA500", anchor="w", width=22
        )
        self.auth_countdown_label.pack(side=tk.LEFT, padx=(10, 0))

        # ── Painel de log ──
        self.log_area = scrolledtext.ScrolledText(
            self, wrap=tk.WORD, font=("Consolas", 11),
            bg="#0d0d0d", fg="#e0e0e0", insertbackground="white",
            borderwidth=0, highlightthickness=1, highlightbackground="#333",
        )
        self.log_area.grid(row=1, column=0, sticky="nsew", padx=30, pady=(8, 4))
        self.log_area.config(state=tk.DISABLED)

        # ── Botão processar ──
        self.processar_btn = tk.Button(
            self, text="Processar Recibos", font=("Consolas", 14, "bold"),
            command=self.on_processar, bg="#2a2a2a", fg="#000000",
            activebackground="#3a3a3a", activeforeground="#7CFC00",
            relief=tk.FLAT, pady=8,
        )
        self.processar_btn.grid(row=2, column=0, pady=(4, 14), ipadx=20)

        # Inicializa estado de conexão (sem fazer chamadas de rede)
        self._draw_status_dot(False)
        self.connect_btn.config(text="Conectar")

    # ── Log ──

    def log(self, text, tag=None):
        """Adiciona texto ao painel de log. Thread-safe via after()."""
        def _insert():
            try:
                self.log_area.config(state=tk.NORMAL)
                self.log_area.insert("end", text)
                self.log_area.see("end")
                self.log_area.config(state=tk.DISABLED)
            except Exception:
                pass
        # Se chamado de outro thread, agenda no main thread
        try:
            if threading.current_thread() is threading.main_thread():
                _insert()
            else:
                self.after(0, _insert)
        except Exception:
            pass

    # ── Indicador de status ──

    def _draw_status_dot(self, connected: bool):
        self.status_canvas.delete("all")
        color = "#00e676" if connected else "#ef5350"
        self.status_canvas.create_oval(3, 3, 15, 15, fill=color, outline=color)

    def _set_connected(self, connected: bool):
        self._draw_status_dot(connected)
        if connected:
            self.connect_btn.config(text="Desconectar")
            # Habilita checkbox
            self.auto_connect_chk.config(state=tk.NORMAL)
        else:
            self.connect_btn.config(text="Conectar")
            # Sem token → checkbox não faz sentido estar marcado
            self.auto_connect_var.set(False)
            _save_prefs({"auto_connect": False})
            self.auto_connect_chk.config(state=tk.DISABLED)

    # ── Barra de progresso de auth ──

    def _auth_progress_show(self, expires_in: int = AUTH_TIMEOUT):
        self.auth_progress_frame.pack(anchor="w", padx=(36, 36), fill=tk.X, pady=(8, 0))
        self.auth_progressbar.start(12)
        self._countdown_secs = expires_in
        self._tick_countdown()

    def _auth_progress_hide(self):
        if self._countdown_id:
            self.after_cancel(self._countdown_id)
            self._countdown_id = None
        self.auth_progressbar.stop()
        self.auth_progress_frame.pack_forget()
        self.auth_countdown_label.config(text="")

    def _tick_countdown(self):
        if self._countdown_secs <= 0:
            self.auth_countdown_label.config(text="⏰ Expirado")
            return
        mins, secs = divmod(self._countdown_secs, 60)
        self.auth_countdown_label.config(text=f"⏳ {mins}:{secs:02d} restantes")
        self._countdown_secs -= 1
        self._countdown_id = self.after(1000, self._tick_countdown)

    # ── Startup seguro (nunca bloqueia) ──

    def checkup_inicial(self):
        """
        Checagem de startup: rápida, nunca faz chamadas de rede no main thread.
        Usa check_token_silent() — sem device flow.
        """
        self.log("════════════════════════════════════\n")
        self.log("   RECIBO AGENT 🧾  v" + VERSION + "\n")
        self.log("════════════════════════════════════\n\n")

        from auth import check_token_silent
        token = check_token_silent()

        if token:
            self._set_connected(True)
            self.log("[OK] Sessão ativa. Token válido.\n")
            # Respeitar preferência de auto_connect
            if self.auto_connect_var.get():
                self.log("[INFO] Auto-connect ativo. Carregando dados da nuvem...\n")
                threading.Thread(target=self._checkup_cloud, daemon=True).start()
            else:
                self.log("[INFO] Pronto. Clique em 'Processar Recibos' para iniciar.\n")
        else:
            self._set_connected(False)
            prefs = _load_prefs()
            if prefs.get("auto_connect"):
                # Token expirou mas a pref estava marcada — avisar e desmarcar
                self.log("[WARN] Sessão expirada. Auto-connect desativado.\n")
                self.log("[WARN] Clique em 'Conectar' para fazer login novamente.\n")
                _save_prefs({"auto_connect": False})
            else:
                self.log("[WARN] Não autenticado. Clique em 'Conectar' para iniciar.\n")

    # ── Checkup na nuvem (roda em background thread) ──

    def _checkup_cloud(self):
        """Carrega contas, alunos e lista recibos via Graph API. Sempre em thread."""
        self.log("\n💳 Carregando contas...\n")
        try:
            from graph_client import load_contas
            load_contas()
            self.log("  ✅ Contas carregadas.\n")
        except Exception as e:
            self.log(f"  ⚠️  Falha ao carregar contas: {e}\n")

        self.log("👩‍🎓 Carregando alunos...\n")
        try:
            from graph_client import load_alunos
            load_alunos()
            self.log("  ✅ Alunos carregados.\n")
        except Exception as e:
            self.log(f"  ⚠️  Falha ao carregar alunos: {e}\n")

        self.log("📑 Verificando recibos na fila...\n")
        try:
            from graph_client import list_onedrive_files
            files = list_onedrive_files()
            if files:
                self.log(f"  {len(files)} recibo(s) aguardando processamento:\n")
                for f in files:
                    kb = f.get("size", 0) // 1024
                    self.log(f"    • {f['name']}  ({kb} KB)\n")
            else:
                self.log("  ✅ Nenhum recibo pendente na fila.\n")
        except Exception as e:
            self.log(f"  ⚠️  Falha ao listar recibos: {e}\n")

        self.log("\n────────────────────────────────────\n")
        self.log("[OK] Pronto.\n")

    # ── Autenticação ──

    def on_connect_btn(self):
        from auth import check_token_silent, disconnect

        # Modo desconectar
        if self.connect_btn.cget("text") == "Desconectar":
            disconnect()
            self._set_connected(False)
            self.log("[INFO] Desconectado. Token removido.\n")
            return

        # Já há thread de auth rodando?
        if self._auth_thread and self._auth_thread.is_alive():
            self.log("[WARN] Autenticação já em andamento...\n")
            return

        self.connect_btn.config(state=tk.DISABLED)
        self._auth_thread = threading.Thread(target=self._auth_flow, daemon=True)
        self._auth_thread.start()

    def _auth_flow(self):
        """Fluxo completo de autenticação em background thread."""
        import socket
        import time

        def re_enable_btn():
            self.connect_btn.config(state=tk.NORMAL)

        try:
            # 1. Teste de conectividade
            self.log("[INFO] Verificando conexão com Microsoft...\n")
            try:
                sock = socket.create_connection(("login.microsoftonline.com", 443), timeout=10)
                sock.close()
            except Exception:
                self.log("[ERRO] Sem acesso a login.microsoftonline.com\n")
                self.log("[DICA] Verifique sua internet ou firewall/proxy/VPN.\n")
                self.after(0, re_enable_btn)
                return

            # 2. Tentar token silencioso primeiro
            from auth import check_token_silent, get_device_code_flow, complete_device_code_flow
            token = check_token_silent()
            if token:
                self.log("[OK] Já autenticado (token renovado silenciosamente).\n")
                self.after(0, lambda: self._set_connected(True))
                self.after(0, re_enable_btn)
                threading.Thread(target=self._checkup_cloud, daemon=True).start()
                return

            # 3. Iniciar device code flow
            self.log("[INFO] Iniciando autenticação Microsoft...\n")
            try:
                flow, app = get_device_code_flow()
            except Exception as e:
                self.log(f"[ERRO] Falha ao iniciar autenticação: {e}\n")
                self.after(0, re_enable_btn)
                return

            url  = flow["verification_uri"]
            code = flow["user_code"]
            expires_in = flow.get("expires_in", AUTH_TIMEOUT)

            # 4. Copiar código pro clipboard
            try:
                pyperclip.copy(code)
            except Exception:
                pass

            # 5. Exibir instruções e aguardar Enter ou Esc do usuário
            self._auth_continue_event = threading.Event()
            self._auth_cancelled = False
            self.after(0, lambda c=code, u=url, e=expires_in: self._show_auth_prompt(c, u, e))

            # Bloqueia a thread de auth até o usuário decidir
            self._auth_continue_event.wait()

            if self._auth_cancelled:
                self.log("\n[INFO] Autenticação cancelada pelo usuário.\n")
                self.after(0, re_enable_btn)
                return

            # 6. Usuário confirmou — abrir browser
            import webbrowser
            try:
                webbrowser.open(url)
                self.log("[OK] Navegador aberto. Cole o código quando solicitado.\n")
            except Exception:
                self.log(f"[WARN] Abra manualmente: {url}\n")

            # 7. Mostrar progressbar + countdown
            self.after(0, lambda: self._auth_progress_show(expires_in))
            self.log("[INFO] Aguardando login no navegador...\n\n")

            # 8. Bloquear aqui (em background thread) até o usuário completar
            #    complete_device_code_flow chama acquire_token_by_device_flow que
            #    é uma única chamada bloqueante com polling interno.
            try:
                token = complete_device_code_flow(flow, app)
            except Exception as e:
                err = str(e)
                self.after(0, self._auth_progress_hide)
                self.after(0, re_enable_btn)
                if "expired" in err.lower() or "timed_out" in err.lower():
                    self.log("[ERRO] Tempo esgotado. Clique em Conectar para tentar novamente.\n")
                else:
                    self.log(f"[ERRO] Autenticação falhou: {err}\n")
                    self.log("[DICA] Tente novamente ou verifique internet/proxy.\n")
                return

            # 9. Sucesso
            self.after(0, self._auth_progress_hide)
            self.after(0, lambda: self._set_connected(True))
            self.after(0, re_enable_btn)
            self.log("[OK] ✅ Autenticado com sucesso!\n\n")
            play_success()
            threading.Thread(target=self._checkup_cloud, daemon=True).start()

        except Exception as e:
            self.log(f"[ERRO] Falha inesperada na autenticação: {e}\n")
            self.after(0, self._auth_progress_hide)
            self.after(0, re_enable_btn)

    # ── Prompt de autenticação (Enter / Esc) ──

    def _show_auth_prompt(self, code: str, url: str, expires_in: int):
        """Exibe instruções detalhadas + imagem + prompt Enter/Esc. Roda no main thread."""
        sep = "━" * 48
        self.log("\n" + sep + "\n")
        self.log("  🔐  AUTENTICAÇÃO MICROSOFT\n")
        self.log(sep + "\n\n")
        self.log(f"  Código de acesso :  {code}\n")
        self.log(f"  Conta            :  kantelira@outlook.com\n")
        self.log(f"  URL              :  {url}\n\n")
        self.log("  Como autenticar:\n")
        self.log("  ① O navegador abrirá a página microsoft.com/link\n")
        self.log(f"  ② Digite o código  {code}  no campo 'Enter code'\n")
        self.log("  ③ Faça login com  kantelira@outlook.com\n")
        self.log("  ④ Clique em 'Allow access' / 'Permitir acesso'\n\n")
        self.log(f"  💾  O código já foi copiado para a memória.\n")
        self.log(f"      Use Cmd+V (ou Ctrl+V) para colar.\n\n")

        # Imagem de referência
        self._insert_auth_image()

        self.log("\n" + sep + "\n")
        self.log("  ▶  Enter  →  abrir o navegador e continuar\n")
        self.log("  ✕  Esc    →  cancelar autenticação\n")
        self.log(sep + "\n\n")

        # Bind temporário de teclas
        self.bind("<Return>", self._on_auth_enter)
        self.bind("<Escape>", self._on_auth_esc)
        self.focus_force()

    def _insert_auth_image(self):
        """Insere imagem de referência no painel de log, se disponível."""
        try:
            from PIL import Image, ImageTk
            import os, sys
            base = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
            img_path = os.path.join(base, "auth_hint.png")
            if not os.path.exists(img_path):
                # Fallback dev: ao lado do script
                img_path = os.path.join(
                    os.path.dirname(os.path.abspath(__file__)), "auth_hint.png"
                )
            if not os.path.exists(img_path):
                return
            img = Image.open(img_path)
            img.thumbnail((420, 300), Image.LANCZOS)
            photo = ImageTk.PhotoImage(img)
            self._auth_hint_photo = photo  # evita GC
            self.log_area.config(state=tk.NORMAL)
            self.log_area.insert("end", "  ")
            self.log_area.image_create("end", image=photo)
            self.log_area.insert("end", "\n")
            self.log_area.see("end")
            self.log_area.config(state=tk.DISABLED)
        except Exception:
            pass  # sem imagem — apenas texto

    def _on_auth_enter(self, event=None):
        self.unbind("<Return>")
        self.unbind("<Escape>")
        self._auth_cancelled = False
        self.log("[OK] Abrindo navegador...\n")
        if self._auth_continue_event:
            self._auth_continue_event.set()

    def _on_auth_esc(self, event=None):
        self.unbind("<Return>")
        self.unbind("<Escape>")
        self._auth_cancelled = True
        if self._auth_continue_event:
            self._auth_continue_event.set()

    # ── Auto-connect toggle ──

    def on_auto_connect_toggle(self):
        val = self.auto_connect_var.get()
        _save_prefs({"auto_connect": val})
        state = "ativado" if val else "desativado"
        self.log(f"[INFO] Conectar automaticamente {state}.\n")

    # ── Configurar pasta remota ──

    def on_configure_folder(self):
        from tkinter import simpledialog
        import importlib
        try:
            from graph_client import ONEDRIVE_RECIBOS_IN_SLUG
            current = ONEDRIVE_RECIBOS_IN_SLUG
        except Exception:
            current = "RECIBOS/RECIBOS_IN"

        new_slug = simpledialog.askstring(
            "Configurar Pasta Remota",
            "Caminho da pasta no OneDrive (ex: RECIBOS/RECIBOS_IN):",
            initialvalue=current,
        )
        if not new_slug or new_slug == current:
            return
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
            self.folder_label.config(text=f"  {new_slug}")
            self.log(f"[OK] Pasta remota atualizada para: {new_slug}\n")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao atualizar pasta: {e}")
            self.log(f"[ERRO] Falha ao atualizar pasta: {e}\n")

    # ── Processar recibos ──

    def on_processar(self):
        from auth import check_token_silent
        if not check_token_silent():
            self.log("[WARN] Não autenticado. Clique em 'Conectar' primeiro.\n")
            return
        self.log("[INFO] Iniciando processamento...\n")
        threading.Thread(target=processar_recibos, args=(self.log,), daemon=True).start()


# ── Entry point ──

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Recibo Agent")
    parser.add_argument("--cli", action="store_true", help="Executa sem interface gráfica")
    args, _ = parser.parse_known_args()

    if args.cli:
        import colorama
        from colorama import Fore, Back, Style
        colorama.init(autoreset=True)
        os.system("clear" if os.name != "nt" else "cls")
        print("   MATT MAGIC RECIBO_AGENT™ " + VERSION)
        print("   by Mateus Ribeiro  |  emaildomat@gmail.com\n")
        processar_recibos(lambda text: print(text, end=""))
    else:
        app = ReciboAgentGUI()
        app.mainloop()
