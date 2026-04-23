#!/usr/bin/env python3
"""
Executor para inicialização automática do Recibo Agent com integração tray.
Este script deve ser configurado para iniciar junto com o sistema operacional.
"""
import subprocess
import sys
from pathlib import Path

# Caminho absoluto para o run.py
AGENT_PATH = Path(__file__).parent / "run.py"

if __name__ == "__main__":
    # Executa o run.py normalmente (com integração tray)
    subprocess.Popen([sys.executable, str(AGENT_PATH)])
    print("Recibo Agent iniciado em background com integração tray.")
