#!/bin/bash
# Recibo Agent - Instalador macOS
# Requer: macOS 12+, Python 3.10+
set -e

echo ""
echo "╔══════════════════════════════════════════╗"
echo "║     RECIBO AGENT - INSTALADOR macOS      ║"
echo "║     Processador automático de recibos    ║"
echo "╚══════════════════════════════════════════╝"
echo ""

AGENT_DIR="$(cd "$(dirname "$0")" && pwd)"
VENV_DIR="/Volumes/Expansion/----- PESSOAL -----/PYTHON/.venv"
PYTHON="$VENV_DIR/bin/python"
LABEL="com.mami.recibo_agent"
PLIST_PATH="$HOME/Library/LaunchAgents/$LABEL.plist"
WRAPPER="/usr/local/bin/recibo_agent_start.sh"
LOG="/tmp/recibo_agent.log"

# ── 1. Verificar Python ──
echo "[1/4] Verificando Python 3..."
if ! command -v python3 &>/dev/null; then
    echo "  ❌ Python 3 não encontrado. Instale via https://python.org ou 'brew install python@3.13'"
    exit 1
fi
echo "  ✅ Python $(python3 --version | awk '{print $2}')"

# ── 2. Criar venv e instalar dependências ──
echo "[2/4] Configurando ambiente Python em $VENV_DIR ..."
if [ ! -d "$VENV_DIR" ]; then
    python3 -m venv "$VENV_DIR"
fi
"$VENV_DIR/bin/pip" install -q -r "$AGENT_DIR/requirements.txt"
echo "  ✅ Dependências instaladas"

# ── 3. Autenticação Microsoft ──
echo "[3/4] Autenticação Microsoft..."
echo "  (Será pedido para abrir microsoft.com/devicelogin e inserir um código)"
echo ""
"$PYTHON" -c "
import sys; sys.path.insert(0, '$AGENT_DIR')
from auth import get_token
get_token()
print('Autenticação concluída!')
" || { echo "  ⚠ Autenticação falhou — tente rodar 'python run.py --once --dry-run' manualmente depois."; }

# ── 4. LaunchAgent (autostart) ──
echo ""
echo "[4/4] Configurando autostart..."

# Criar wrapper script
sudo tee "$WRAPPER" > /dev/null << WRAPPER_SCRIPT
#!/bin/bash
export PYTHONUNBUFFERED=1
cd "$AGENT_DIR"
exec "$PYTHON" -u run.py
WRAPPER_SCRIPT
sudo chmod +x "$WRAPPER"
echo "  ✅ Wrapper criado: $WRAPPER"

# Criar plist
mkdir -p "$HOME/Library/LaunchAgents"
cat > "$PLIST_PATH" << PLIST
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN"
  "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>Label</key>
    <string>$LABEL</string>
    <key>ProgramArguments</key>
    <array>
        <string>$WRAPPER</string>
    </array>
    <key>WorkingDirectory</key>
    <string>$AGENT_DIR</string>
    <key>RunAtLoad</key>
    <true/>
    <key>KeepAlive</key>
    <true/>
    <key>ThrottleInterval</key>
    <integer>10</integer>
    <key>ProcessType</key>
    <string>Background</string>
    <key>Nice</key>
    <integer>10</integer>
    <key>StandardOutPath</key>
    <string>$LOG</string>
    <key>StandardErrorPath</key>
    <string>$LOG</string>
</dict>
</plist>
PLIST

launchctl unload "$PLIST_PATH" 2>/dev/null || true
launchctl load "$PLIST_PATH"
echo "  ✅ LaunchAgent carregado: $LABEL"
echo "  Log em: $LOG"

echo ""
echo "╔══════════════════════════════════════════╗"
echo "║    INSTALAÇÃO CONCLUÍDA!                 ║"
echo "╠══════════════════════════════════════════╣"
echo "║  Coloque recibos em:                     ║"
echo "║  MAMI/RECIBOS IN/                        ║"
echo "║                                          ║"
echo "║  Ver log:                                ║"
echo "║  tail -f /tmp/recibo_agent.log           ║"
echo "║                                          ║"
echo "║  Parar serviço:                          ║"
echo "║  launchctl stop com.mami.recibo_agent    ║"
echo "╚══════════════════════════════════════════╝"
echo ""
