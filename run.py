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

SUPPORTED = {".jpg", ".jpeg", ".png", ".webp", ".pdf"}


def preview_payment(data: dict) -> str:
    """Retorna string formatada para preview."""
    lines = ["  Dados extraídos:"]
    for key, val in data.items():
        if val:
            lines.append(f"    {key}: {val}")
    return "\n".join(lines)


def process_file(file_path: Path, dry_run: bool = False) -> bool:
    """Processa um único arquivo. Retorna True se sucesso."""
    print(f"\n📄 Processando: {file_path.name}")

    data = process_receipt(file_path)
    if not data:
        return False

    print(preview_payment(data))

    if dry_run:
        print("  [DRY RUN] Nenhuma alteração feita.")
        return True

    try:
        pay_id = insert_payment(data)
        print(f"  ✅ Inserido como {pay_id}")
    except DuplicateReceiptError as e:
        print(f"  ⚠ Duplicata ignorada: {e}")
        _move_to_processed(file_path, f"DUP-{file_path.name}")
        return True
    except Exception as e:
        print(f"  ❌ Erro ao inserir no Excel Online: {e}")
        return False

    # Upload para OneDrive renomeado como {pay_id}.ext
    uploaded = False
    try:
        upload_receipt(file_path, pay_id)
        uploaded = True
    except Exception as e:
        print(f"  ⚠ Upload falhou (arquivo local mantido): {e}")

    # Mover arquivo local para ARQUIVOS PROCESSADOS, renomeado como {pay_id}.ext
    dest_name = f"{pay_id}{file_path.suffix.lower()}"
    _move_to_processed(file_path, dest_name)

    return True


def _move_to_processed(file_path: Path, dest_name: str):
    """Move arquivo para ARQUIVOS PROCESSADOS com o nome indicado."""
    PROCESSED_FOLDER.mkdir(parents=True, exist_ok=True)
    dest = PROCESSED_FOLDER / dest_name
    shutil.move(str(file_path), str(dest))
    print(f"  📁 Movido para ARQUIVOS PROCESSADOS: {dest.name}")


def process_pending(dry_run: bool = False):
    """Processa todos os arquivos pendentes na pasta."""
    if not WATCH_FOLDER.exists():
        print(f"❌ Pasta não encontrada: {WATCH_FOLDER}")
        return

    files = sorted(
        f for f in WATCH_FOLDER.iterdir()
        if f.is_file() and f.suffix.lower() in SUPPORTED
    )

    if not files:
        print("Nenhum recibo pendente.")
        return

    print(f"Encontrados {len(files)} recibo(s) pendente(s)")
    ok = 0
    for f in files:
        if process_file(f, dry_run):
            ok += 1

    print(f"\n{'=' * 40}")
    print(f"Processados: {ok}/{len(files)}")


def watch_mode():
    """Monitora a pasta continuamente com watchdog."""
    try:
        from watchdog.observers import Observer
        from watchdog.events import FileSystemEventHandler
    except ImportError:
        print("❌ watchdog não instalado. Use: pip install watchdog")
        print("   Ou use --once para processar arquivos existentes.")
        sys.exit(1)

    class ReciboHandler(FileSystemEventHandler):
        def on_created(self, event):
            if event.is_directory:
                return
            path = Path(event.src_path)
            if path.suffix.lower() not in SUPPORTED:
                return
            # Espera o arquivo terminar de ser escrito (sync OneDrive)
            time.sleep(3)
            if path.exists():
                process_file(path)

    observer = Observer()
    observer.schedule(ReciboHandler(), str(WATCH_FOLDER), recursive=False)
    observer.start()

    print(f"👀 Monitorando: {WATCH_FOLDER}")
    print("   Ctrl+C para parar\n")

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()


def main():
    parser = argparse.ArgumentParser(description="Recibo Agent")
    parser.add_argument("--once", action="store_true",
                        help="Processa pendentes e sai")
    parser.add_argument("--file", type=str,
                        help="Processa um arquivo específico")
    parser.add_argument("--dry-run", action="store_true",
                        help="Preview sem escrever no Excel")
    parser.add_argument("--sync-onedrive", action="store_true",
                        help="Remove órfãos e duplicatas de OneDrive/RECIBOS")
    args = parser.parse_args()

    # Carrega contas de recebimento da tblContas
    print("Carregando contas do Excel Online...")
    try:
        load_contas()
    except Exception as e:
        print(f"  ⚠ Não foi possível carregar contas: {e}")

    # Carrega alunos conhecidos para matching (via Graph API)
    print("Carregando alunos do Excel Online...")
    try:
        load_alunos()
    except Exception as e:
        print(f"  ⚠ Não foi possível carregar alunos: {e}")
        print("  Continuando sem lista de alunos (matching desativado)")

    if args.sync_onedrive:
        sync_onedrive_recibos(dry_run=args.dry_run)
        return

    if args.file:
        f = Path(args.file)
        if not f.exists():
            print(f"❌ Arquivo não encontrado: {f}")
            sys.exit(1)
        process_file(f, args.dry_run)
    elif args.once:
        process_pending(args.dry_run)
    else:
        # Processa pendentes primeiro, depois entra em watch
        process_pending(args.dry_run)
        if not args.dry_run:
            watch_mode()


if __name__ == "__main__":
    main()
