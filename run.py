#!/usr/bin/env python3
"""
Recibo Agent — Monitora pasta de recibos e preenche tblPagamentos automaticamente.

Uso:
  python run.py                  # Modo watch (monitora pasta continuamente)
  python run.py --once           # Processa arquivos pendentes e sai
  python run.py --file X.jpg     # Processa um arquivo específico
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


from config import WATCH_FOLDER, PROCESSED_FOLDER
from processor import process_receipt
from graph_client import load_alunos, load_contas, insert_payment, upload_receipt, DuplicateReceiptError, sync_onedrive_recibos
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

    def process_file(f, dry_run=False):
        paragrafo()
        print("🔹🔹🔹")
        print(f"⏳ Processando arquivo: {f.name}")
        print("────────────")
        try:
            data = process_receipt(f)
            if not data:
                print(f"❌ Não foi possível extrair dados de {f.name}")
                print("────────────")
                return False
            print(f"📝 Dados extraídos:\n  {data}")
            if dry_run:
                print("  [DRY RUN] Nenhuma alteração feita.")
                print("────────────")
                return True
            pay_id = insert_payment(data)
            print(f"  ✅ Inserido como {pay_id}")
            upload_receipt(f, pay_id)
            print(f"  ☁️  Upload concluído: {pay_id}{f.suffix.lower()}")
            dest_name = f"{pay_id}{f.suffix.lower()}"
            shutil.move(str(f), str(PROCESSED_FOLDER / dest_name))
            print(f"  📦 Movido para ARQUIVOS PROCESSADOS: {dest_name}")
            print("────────────")
            return True
        except DuplicateReceiptError:
            print(f"⚠️ Duplicado: {f.name}")
            dest_name = f"DUP-{f.name}"
            shutil.move(str(f), str(PROCESSED_FOLDER / dest_name))
            print(f"  📦 Movido para ARQUIVOS PROCESSADOS: {dest_name}")
            print("────────────")
            return True
        except Exception as e:
            print(f"❌ Erro em {f.name}: {e}")
            print("────────────")
            return False

    if args.file:
        f = Path(args.file)
        if not f.exists():
            linha()
            print(f"❌ Arquivo não encontrado: {f}")
            linha()
            sys.exit(1)
        process_file(f, args.dry_run)
    elif args.once:
        files = [f for f in WATCH_FOLDER.iterdir() if f.is_file() and f.suffix.lower() in SUPPORTED]
        if not files:
            paragrafo("📭 Nenhum recibo pendente.")
            print("────────────")
            return
        paragrafo(f"📑 Processando {len(files)} recibo(s)...")
        print("────────────")
        with tqdm(total=len(files), desc="Processando recibos", unit="recibo") as pbar:
            for f in files:
                process_file(f, args.dry_run)
                pbar.update(1)
        print("🏁 Processamento concluído!")
        print("────────────")
    else:
        files = [f for f in WATCH_FOLDER.iterdir() if f.is_file() and f.suffix.lower() in SUPPORTED]
        if not files:
            paragrafo("Nenhum recibo pendente.")
            linha()
        else:
            paragrafo(f"Processando {len(files)} recibo(s)...")
            linha()
            with tqdm(total=len(files), desc="Processando recibos", unit="recibo") as pbar:
                for f in files:
                    process_file(f, args.dry_run)
                    pbar.update(1)
            linha()
            print("Processamento concluído.")
            linha()


if __name__ == "__main__":
    main()
