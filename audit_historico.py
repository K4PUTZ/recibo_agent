#!/usr/bin/env python3
"""
Auditoria e upload dos recibos históricos (arquivos nomeados 1.jpg, 2.pdf, etc.)

O agente foi iniciado após 155 pagamentos já terem sido lançados manualmente
na planilha. Esses recibos estão em RECIBOS PROCESSADOS com nomes numéricos.
Este script mapeia N.ext → PAY-000000{N} e faz upload para OneDrive/RECIBOS.

Uso:
  python audit_historico.py             # Dry-run (só mostra o que faria)
  python audit_historico.py --upload    # Executa uploads reais
  python audit_historico.py --rename    # Renomeia arquivos locais para PAY-ID.ext
  python audit_historico.py --upload --rename  # Upload + renomeia local
"""
import argparse
import re
import sys
import ssl
import urllib.parse
from pathlib import Path

# Fix SSL certificate on macOS/venv
try:
    import certifi
    ssl._create_default_https_context = lambda: ssl.create_default_context(cafile=certifi.where())
except ImportError:
    pass

# Adiciona o diretório do script ao path
sys.path.insert(0, str(Path(__file__).parent))

import requests
from auth import get_token
from config import (
    PROCESSED_FOLDER, ONEDRIVE_PROCESSED_PATH,
    WORKBOOK_ONEDRIVE_PATH, TABLE_NAME
)

GRAPH = "https://graph.microsoft.com/v1.0"
SUPPORTED_EXT = {".jpg", ".jpeg", ".png", ".webp", ".pdf"}
SKIP_EXT = {".docx", ".doc"}

CONTENT_TYPES = {
    ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
    ".png": "image/png", ".webp": "image/webp",
    ".pdf": "application/pdf",
}


def headers() -> dict:
    return {
        "Authorization": f"Bearer {get_token()}",
        "Content-Type": "application/json",
    }


def get_workbook_id() -> str:
    encoded = urllib.parse.quote(WORKBOOK_ONEDRIVE_PATH)
    r = requests.get(f"{GRAPH}/me/drive/root:/{encoded}", headers=headers())
    r.raise_for_status()
    item = r.json()
    print(f"  Workbook: {item['name']}")
    return item["id"]


def get_pay_ids_in_excel(wid: str) -> set[str]:
    """Retorna todos os PaymentIDs presentes na tblPagamentos."""
    url = f"{GRAPH}/me/drive/items/{wid}/workbook/tables/{TABLE_NAME}/rows"
    r = requests.get(url, headers=headers())
    r.raise_for_status()
    ids = set()
    for row in r.json().get("value", []):
        vals = row.get("values", [[]])[0]
        if vals and isinstance(vals[0], str) and vals[0].startswith("PAY-"):
            ids.add(vals[0])
    return ids


def get_onedrive_recibos() -> dict[str, str]:
    """Retorna {stem: item_id} dos arquivos em OneDrive/RECIBOS."""
    folder_encoded = urllib.parse.quote(ONEDRIVE_PROCESSED_PATH)
    r = requests.get(
        f"{GRAPH}/me/drive/root:/{folder_encoded}:/children?$select=name,id",
        headers=headers()
    )
    if r.status_code == 404:
        return {}
    r.raise_for_status()
    result = {}
    for f in r.json().get("value", []):
        stem = f["name"].rsplit(".", 1)[0]
        result[stem] = f["id"]
    return result


def upload_file(local_path: Path, pay_id: str) -> bool:
    """Faz upload do arquivo local para OneDrive/RECIBOS/{pay_id}.ext."""
    ext = local_path.suffix.lower()
    dest = f"{ONEDRIVE_PROCESSED_PATH}/{pay_id}{ext}"
    encoded = urllib.parse.quote(dest)
    url = f"{GRAPH}/me/drive/root:/{encoded}:/content"

    hdrs = headers()
    hdrs["Content-Type"] = CONTENT_TYPES.get(ext, "application/octet-stream")

    with open(local_path, "rb") as f:
        data = f.read()

    r = requests.put(url, headers=hdrs, data=data)
    if r.ok:
        kb = len(data) // 1024
        print(f"  ✅ Upload: {local_path.name} → {pay_id}{ext} ({kb} KB)")
        return True
    else:
        print(f"  ❌ Falha upload {local_path.name}: {r.status_code} {r.text[:200]}")
        return False


def rename_local(local_path: Path, pay_id: str):
    """Renomeia o arquivo local de N.ext para PAY-ID.ext (no mesmo diretório)."""
    new_name = pay_id + local_path.suffix.lower()
    dest = local_path.parent / new_name
    if dest.exists():
        print(f"  ⚠ Renomear: destino já existe ({new_name}), pulando")
        return
    local_path.rename(dest)
    print(f"  📁 Renomeado: {local_path.name} → {new_name}")


def collect_numeric_files() -> list[tuple[int, Path]]:
    """Retorna [(N, path)] para todos os arquivos com nome numérico em PROCESSADOS."""
    results = []
    for f in PROCESSED_FOLDER.iterdir():
        stem = f.stem
        if re.fullmatch(r"\d+", stem):
            results.append((int(stem), f))
    return sorted(results, key=lambda x: x[0])


def main():
    parser = argparse.ArgumentParser(description="Auditoria de recibos históricos")
    parser.add_argument("--upload", action="store_true",
                        help="Executa uploads reais para OneDrive")
    parser.add_argument("--rename", action="store_true",
                        help="Renomeia arquivos locais para PAY-ID.ext após upload")
    args = parser.parse_args()

    dry_run = not args.upload
    if dry_run:
        print("=" * 60)
        print("  DRY-RUN: nenhuma alteração será feita")
        print("  Use --upload para executar de verdade")
        print("=" * 60)

    print("\n🔍 Coletando dados...")
    wid = get_workbook_id()
    pay_ids_excel = get_pay_ids_in_excel(wid)
    onedrive_stems = get_onedrive_recibos()

    print(f"  {len(pay_ids_excel)} PAY IDs na planilha")
    print(f"  {len(onedrive_stems)} arquivos em OneDrive/RECIBOS")

    numeric_files = collect_numeric_files()
    print(f"  {len(numeric_files)} arquivos numéricos em RECIBOS PROCESSADOS\n")

    # Verifica quais numbers estão faltando (sem arquivo local)
    present_nums = {n for n, _ in numeric_files}

    stats = {
        "total": len(numeric_files),
        "pulados_ext": [],      # .docx etc
        "sem_pay_id": [],       # número sem PAY-ID na planilha
        "ja_cloud": [],         # já no OneDrive
        "upload_ok": [],
        "upload_fail": [],
        "renomeados": [],
    }

    for n, fpath in numeric_files:
        ext = fpath.suffix.lower()
        pay_id = f"PAY-{n:09d}"

        # 1. Extensão não suportada
        if ext in SKIP_EXT:
            stats["pulados_ext"].append(fpath.name)
            print(f"  ⏭  {fpath.name} → ignorado (formato {ext} não suportado)")
            continue

        if ext not in SUPPORTED_EXT:
            stats["pulados_ext"].append(fpath.name)
            print(f"  ⏭  {fpath.name} → ignorado (extensão desconhecida {ext})")
            continue

        # 2. PAY-ID existe na planilha?
        if pay_id not in pay_ids_excel:
            stats["sem_pay_id"].append(fpath.name)
            print(f"  ⚠  {fpath.name} → {pay_id} não encontrado na planilha, pulando")
            continue

        # 3. Já está no OneDrive?
        if pay_id in onedrive_stems:
            stats["ja_cloud"].append(pay_id)
            print(f"  ☁  {fpath.name} → já no OneDrive como {pay_id}.*")
            continue

        # 4. Upload
        print(f"  📤 {fpath.name} → {pay_id}{ext}", end="")
        if dry_run:
            print(" [DRY-RUN]")
            stats["upload_ok"].append(pay_id)
        else:
            print()
            ok = upload_file(fpath, pay_id)
            if ok:
                stats["upload_ok"].append(pay_id)
                if args.rename:
                    rename_local(fpath, pay_id)
                    stats["renomeados"].append(pay_id)
            else:
                stats["upload_fail"].append(pay_id)

    # Resumo
    print("\n" + "=" * 60)
    print("RESUMO")
    print("=" * 60)
    print(f"  Total de arquivos numéricos: {stats['total']}")
    print(f"  Já no OneDrive (pulados):    {len(stats['ja_cloud'])}")
    print(f"  Uploads {'simulados' if dry_run else 'realizados'}:         {len(stats['upload_ok'])}")
    if stats["upload_fail"]:
        print(f"  Falhas de upload:            {len(stats['upload_fail'])}")
        for pid in stats["upload_fail"]:
            print(f"    - {pid}")
    if stats["pulados_ext"]:
        print(f"  Formato não suportado:       {len(stats['pulados_ext'])}")
        for f in stats["pulados_ext"]:
            print(f"    - {f}")
    if stats["sem_pay_id"]:
        print(f"  Sem PAY-ID na planilha:      {len(stats['sem_pay_id'])}")
        for f in stats["sem_pay_id"]:
            print(f"    - {f}")
    if args.rename and stats["renomeados"]:
        print(f"  Arquivos locais renomeados:  {len(stats['renomeados'])}")

    # Verifica PAY IDs que existem na planilha mas não têm arquivo local
    print()
    all_nums_from_excel = set()
    for pid in pay_ids_excel:
        try:
            all_nums_from_excel.add(int(pid.replace("PAY-", "")))
        except ValueError:
            pass

    # Histórico: PAY IDs 1-155 sem arquivo local
    historico_range = range(1, max(present_nums) + 1) if present_nums else range(0)
    missing_files = [
        f"PAY-{n:09d}" for n in historico_range
        if n not in present_nums and f"PAY-{n:09d}" in pay_ids_excel
    ]
    if missing_files:
        print(f"  PAY IDs sem arquivo local ({len(missing_files)}):")
        for pid in missing_files:
            print(f"    - {pid}")


if __name__ == "__main__":
    main()
