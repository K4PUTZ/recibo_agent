import warnings
warnings.filterwarnings("ignore")
#!/usr/bin/env python3
"""
Recibo Agent — Processamento 100% na nuvem (OneDrive).

Este script processa recibos enviados para a pasta remota (OneDrive), extrai dados via OCR e preenche a tabela de pagamentos no Excel Online.
Todos os arquivos são baixados, processados, enviados para a pasta de processados e o original é removido da nuvem.

Fluxo principal:
- Lista arquivos na pasta remota (OneDrive).
- Baixa cada arquivo temporariamente.
- Processa via OCR e extrai dados.
- Insere pagamento na planilha online.
- Faz upload do recibo processado.
- Remove o original do OneDrive.

Não há dependência de pastas locais. Todo o processamento é feito na nuvem.

Uso:
    python run.py                  # Processa todos os arquivos pendentes na nuvem
    python run.py --once           # Processa arquivos pendentes e sai
    python run.py --dry-run        # Mostra o que faria sem escrever no Excel
    python run.py --sync-onedrive  # Limpa OneDrive/RECIBOS: remove órfãos e duplicatas
"""
import argparse
import shutil
import ssl
import sys
import time
from pathlib import Path

# Fix SSL certificate on macOS/venv setups
try:
    import certifi
    ssl._create_default_https_context = lambda: ssl.create_default_context(cafile=certifi.where())
except ImportError:
    pass

# Adiciona o diretório do script ao path
sys.path.insert(0, str(Path(__file__).parent))


from processor import process_receipt
from graph_client import (
    load_alunos, load_contas, insert_payment, upload_receipt, DuplicateReceiptError, sync_onedrive_recibos,
    list_onedrive_files, download_onedrive_file, move_onedrive_file, ONEDRIVE_RECIBOS_IN_SLUG, ONEDRIVE_RECIBOS_PROCESSED_SLUG
)
# Integração com tray

import threading
from tqdm import tqdm

# Integração com tray removida. Interface agora é terminal.
TRAY_AVAILABLE = False


# Elementos visuais para terminal
def linha(tam=20, char="─"):
    print(char * tam)

def titulo(txt, tam=20, char="═"):
    print()
    print(char * tam)
    print(f"{txt:^{tam}} 🧾")
    print(char * tam)

def paragrafo(txt=""):
    print()
    if txt:
        print(txt)

def main():
    parser = argparse.ArgumentParser(description="Recibo Agent")
    parser.add_argument("--once", action="store_true", help="Processa pendentes e sai")
    parser.add_argument("--file", type=str, help="Processa um arquivo específico")
    parser.add_argument("--dry-run", action="store_true", help="Preview sem escrever no Excel")
    parser.add_argument("--sync-onedrive", action="store_true", help="Remove órfãos e duplicatas de OneDrive/RECIBOS")
    args = parser.parse_args()

    titulo("RECIBO AGENT")

    # Carrega contas de recebimento da tblContas
    paragrafo("💳 Carregando contas do Excel Online...")
    try:
        load_contas()
        print("  ✅ Contas carregadas com sucesso.")
    except Exception as e:
        print(f"  ⚠️ Não foi possível carregar contas: {e}")

    # Carrega alunos conhecidos para matching (via Graph API)
    paragrafo("👩‍🎓 Carregando alunos do Excel Online...")
    try:
        load_alunos()
        print("  ✅ Alunos carregados com sucesso.")
    except Exception as e:
        print(f"  ⚠️ Não foi possível carregar alunos: {e}")
        print("  Continuando sem lista de alunos (matching desativado)")

    linha()

    if args.sync_onedrive:
        titulo("SINCRONIZAÇÃO ONEDRIVE/RECIBOS")
        sync_onedrive_recibos(dry_run=args.dry_run)
        linha()
        return

    SUPPORTED = {'.jpg', '.jpeg', '.png', '.webp', '.pdf', '.docx'}

    import tempfile, os
    def process_onedrive_file(file_info, dry_run=False):
        """
        Baixa, processa, insere e remove um arquivo do OneDrive.
        Args:
            file_info (dict): Dict com 'name', 'id', 'size'.
            dry_run (bool): Se True, apenas simula.
        Returns:
            bool: True se processado com sucesso.
        """
        paragrafo()
        print("🔹🔹🔹")
        print(f"⏳ Processando arquivo: {file_info['name']}")
        print("────────────")
        ext = os.path.splitext(file_info['name'])[1].lower()
        if ext not in SUPPORTED:
            print(f"⏭️ Formato não suportado: {file_info['name']}")
            print("────────────")
            return False
        from pathlib import Path
        with tempfile.TemporaryDirectory() as tmpdir:
            local_path = os.path.join(tmpdir, file_info['name'])
            download_onedrive_file(file_info['id'], local_path)
            try:
                data = process_receipt(Path(local_path))
                if not data:
                    print(f"❌ Não foi possível extrair dados de {file_info['name']}")
                    print("────────────")
                    return False
                print(f"📝 Dados extraídos:\n  {data}")
                if dry_run:
                    print("  [DRY RUN] Nenhuma alteração feita.")
                    print("────────────")
                    return True
                pay_id = insert_payment(data)
                print(f"  ✅ Inserido como {pay_id}")
                upload_receipt(local_path, pay_id)
                print(f"  ☁️  Upload concluído: {pay_id}{ext}")
                from graph_client import delete_onedrive_file
                delete_onedrive_file(file_info['id'])
                print(f"  🗑️ Original removido: {file_info['name']}")
                print("────────────")
                return True
            except DuplicateReceiptError:
                print(f"⚠️ Duplicado: {file_info['name']}")
                from graph_client import delete_onedrive_file
                delete_onedrive_file(file_info['id'])
                print(f"  🗑️ Original removido: {file_info['name']}")
                print("────────────")
                return True
            except Exception as e:
                print(f"❌ Erro em {file_info['name']}: {e}")
                print("────────────")
                return False

    if args.file:
        print("[ERRO] Processamento de arquivo individual não suportado no modo OneDrive remoto.")
        print("Use apenas o modo padrão para processar todos os arquivos da nuvem.")
        sys.exit(1)
    elif args.once:
        """
        Processa todos os recibos pendentes na nuvem uma única vez.
        """
        files = list_onedrive_files()
        if not files:
            paragrafo("📭 Nenhum recibo pendente na nuvem.")
            print("────────────")
            return
        paragrafo(f"📑 Processando {len(files)} recibo(s) na nuvem...")
        print("────────────")
        with tqdm(total=len(files), desc="Processando recibos", unit="recibo") as pbar:
            for file_info in files:
                process_onedrive_file(file_info, args.dry_run)
                pbar.update(1)
        print("🏁 Processamento concluído!")
        print("────────────")
    else:
        """
        Processa todos os recibos pendentes na nuvem em modo contínuo (padrão).
        """
        files = list_onedrive_files()
        if not files:
            paragrafo("Nenhum recibo pendente na nuvem.")
            linha()
        else:
            paragrafo(f"Processando {len(files)} recibo(s) na nuvem...")
            linha()
            with tqdm(total=len(files), desc="Processando recibos", unit="recibo") as pbar:
                for file_info in files:
                    process_onedrive_file(file_info, args.dry_run)
                    pbar.update(1)
            linha()
            print("Processamento concluído.")
            linha()


if __name__ == "__main__":
    main()
