"""
Configuração do Recibo Agent.
Funciona em macOS e Windows. Edite os caminhos se necessário.
"""
from pathlib import Path
import platform
import json

# ── Detecção automática de SO ──
IS_MAC = platform.system() == "Darwin"
IS_WIN = platform.system() == "Windows"

# ── Diretório do app (onde ficam token cache, logs, etc.) ──
APP_DIR = Path(__file__).parent

# ── Caminhos locais ──
if IS_MAC:
    WATCH_FOLDER = Path("/Volumes/Expansion/----- MAMI -----/RECIBOS IN")
    PROCESSED_FOLDER = Path("/Volumes/Expansion/----- MAMI -----/RECIBOS PROCESSADOS")
else:
    # Windows — pasta OneDrive da mãe
    WATCH_FOLDER = Path.home() / "OneDrive" / "MAMI" / "RECIBOS IN"
    PROCESSED_FOLDER = Path.home() / "OneDrive" / "MAMI" / "RECIBOS PROCESSADOS"

# ── Pasta no OneDrive onde os recibos processados são salvos ──
# "RECIBOS" = Meus Arquivos/RECIBOS/ na raiz do OneDrive
ONEDRIVE_PROCESSED_PATH = "RECIBOS"

# ── Microsoft Graph API ──
# Preencha CLIENT_ID após criar o App Registration no Azure Portal.
# Instruções: https://portal.azure.com → Azure Active Directory → App registrations
CLIENT_ID = "c2618b42-a033-4db3-b193-31109c2fcb1b"
AUTHORITY = "https://login.microsoftonline.com/consumers"  # Conta pessoal Microsoft
SCOPES = ["Files.ReadWrite.All"]
TOKEN_CACHE_PATH = APP_DIR / ".token_cache.json"

# ── Excel Online ──
WORKBOOK_NAME = "GESTAO_CURSOS_2026.xlsx"
WORKBOOK_ONEDRIVE_PATH = "2026/CURSOS 2026/TURMAS/GESTAO_CURSOS_2026.xlsx"
TABLE_NAME = "tblPagamentos"

# Colunas da tblPagamentos na ordem exata do Excel (11 colunas)
TABLE_COLUMNS = [
    "PaymentID",               # A
    "AlunoBeneficiario",       # B
    "DataPagamento",           # C
    "VALOR",                   # D
    "Competencia",             # E
    "Recibo",                  # F
    "ContaRecebimento",        # G
    "NF",                      # H
    "StatusPagamento",         # I
    "PagadorNome(Opcional)",   # J
    "ObservacoesPagamento",    # K
]

# ── Constantes ──
MESES = ["JAN", "FEV", "MAR", "ABR", "MAI", "JUN",
         "JUL", "AGO", "SET", "OUT", "NOV", "DEZ"]

CONTAS: list[str] = []  # preenchido via load_contas() da tblContas
CONTA_PISTAS: dict[str, str] = {}  # preenchido via load_contas() da tblContas

# Preenchidas ao ler a aba ALUNOS via Graph API
ALUNOS_CONHECIDOS: list[str] = []
# Mapa pagador_nome → aluno_nome (quando coluna PAGADOR está preenchida)
PAGADORES_MAP: dict[str, str] = {}
