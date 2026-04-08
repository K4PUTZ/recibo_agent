"""
Cliente Microsoft Graph API para operações no Excel Online.
Escreve direto na tblPagamentos sem tocar no arquivo local —
funciona mesmo com o Excel aberto no browser.
"""
import re
import urllib.parse
import requests

from auth import get_token
from config import (WORKBOOK_ONEDRIVE_PATH, TABLE_NAME, TABLE_COLUMNS,
                    ALUNOS_CONHECIDOS, PAGADORES_MAP, ONEDRIVE_PROCESSED_PATH,
                    CONTAS, CONTA_PISTAS)

GRAPH = "https://graph.microsoft.com/v1.0"

_workbook_id: str | None = None
_known_tx_ids: set[str] = set()


class DuplicateReceiptError(Exception):
    pass


def _headers() -> dict:
    return {
        "Authorization": f"Bearer {get_token()}",
        "Content-Type": "application/json",
    }


def _find_workbook_id() -> str:
    """Obtém o item ID do workbook pelo caminho no OneDrive."""
    global _workbook_id
    if _workbook_id:
        return _workbook_id

    encoded = urllib.parse.quote(WORKBOOK_ONEDRIVE_PATH)
    url = f"{GRAPH}/me/drive/root:/{encoded}"
    r = requests.get(url, headers=_headers())
    r.raise_for_status()

    item = r.json()
    if "id" not in item:
        raise RuntimeError(
            f"'{WORKBOOK_ONEDRIVE_PATH}' não encontrado no OneDrive."
        )

    _workbook_id = item["id"]
    print(f"  Workbook encontrado: {item['name']}")
    return _workbook_id


def _table_url() -> str:
    wid = _find_workbook_id()
    return f"{GRAPH}/me/drive/items/{wid}/workbook/tables/{TABLE_NAME}"


# ── Leitura ──

def load_alunos():
    """Carrega nomes de alunos da aba ALUNOS (col NomeCompleto) e mapa Pagador→Aluno."""
    wid = _find_workbook_id()
    url = f"{GRAPH}/me/drive/items/{wid}/workbook/worksheets/ALUNOS/usedRange"
    r = requests.get(url, headers=_headers())
    r.raise_for_status()

    values = r.json().get("values", [])
    if not values:
        return

    # Descobre índices das colunas pelo cabeçalho
    header = [str(c).strip() for c in values[0]]
    try:
        idx_nome = header.index("NomeCompleto")
    except ValueError:
        idx_nome = 1  # fallback: coluna B
    try:
        idx_pagador = header.index("PAGADOR")
    except ValueError:
        idx_pagador = None

    names = set()
    PAGADORES_MAP.clear()

    for row in values[1:]:
        nome = row[idx_nome].strip() if len(row) > idx_nome and row[idx_nome] else ""
        if not nome or not isinstance(nome, str):
            continue
        names.add(nome)

        if idx_pagador is not None and len(row) > idx_pagador:
            pag = row[idx_pagador]
            if pag and isinstance(pag, str) and pag.strip():
                PAGADORES_MAP[pag.strip().upper()] = nome

    ALUNOS_CONHECIDOS.clear()
    ALUNOS_CONHECIDOS.extend(sorted(names))
    print(f"  {len(ALUNOS_CONHECIDOS)} alunos carregados", end="")
    if PAGADORES_MAP:
        print(f", {len(PAGADORES_MAP)} pagadores mapeados", end="")
    print()


def load_contas():
    """Carrega tblContas da aba CURSOS e atualiza CONTAS e CONTA_PISTAS."""
    wid = _find_workbook_id()
    url = f"{GRAPH}/me/drive/items/{wid}/workbook/tables/tblContas/rows"
    r = requests.get(url, headers=_headers())
    r.raise_for_status()

    CONTAS.clear()
    CONTA_PISTAS.clear()

    for row in r.json().get("value", []):
        vals = row.get("values", [[]])[0]
        if len(vals) < 5:
            continue
        nome, titular, chaves, _, ativo = vals[0], vals[1], vals[2], vals[3], vals[4]
        if str(ativo).strip().upper() != "SIM":
            continue
        nome = str(nome).strip()
        if not nome:
            continue
        CONTAS.append(nome)
        # Cada fragmento de ChavesPix vira uma entrada no mapa
        for pista in str(chaves).split(","):
            pista = pista.strip().lower()
            if pista:
                CONTA_PISTAS[pista] = nome
        # Titular também como pista (primeiras 2 palavras em lower)
        if titular:
            palavras = str(titular).lower().split()
            if len(palavras) >= 2:
                CONTA_PISTAS[f"{palavras[0]} {palavras[1]}"] = nome

    print(f"  {len(CONTAS)} contas carregadas: {CONTAS}")


def get_next_payment_id() -> str:
    """Lê a tabela, retorna o próximo PAY-XXXXXXXXX e atualiza _known_tx_ids."""
    global _known_tx_ids
    url = f"{_table_url()}/rows"
    r = requests.get(url, headers=_headers())
    r.raise_for_status()

    try:
        obs_idx = TABLE_COLUMNS.index("ObservacoesPagamento")
    except ValueError:
        obs_idx = -1

    max_id = 0
    _known_tx_ids = set()

    for row in r.json().get("value", []):
        vals = row.get("values", [[]])[0]
        if vals and isinstance(vals[0], str) and vals[0].startswith("PAY-"):
            try:
                num = int(vals[0].replace("PAY-", ""))
                max_id = max(max_id, num)
            except ValueError:
                pass
        if obs_idx >= 0 and len(vals) > obs_idx:
            m = re.search(r"TX:([A-Z0-9]{20,})", str(vals[obs_idx]))
            if m:
                _known_tx_ids.add(m.group(1))

    return f"PAY-{max_id + 1:09d}"


# ── Escrita ──

def insert_payment(data: dict) -> str:
    """
    Insere uma linha na tblPagamentos via Graph API.
    Retorna o PaymentID gerado.
    Lança DuplicateReceiptError se o recibo já foi cadastrado.
    """
    pay_id = get_next_payment_id()  # também atualiza _known_tx_ids

    # Checa duplicata pelo ID de transação
    tx_id = data.get("_tx_id")
    if tx_id and tx_id in _known_tx_ids:
        raise DuplicateReceiptError(f"TX:{tx_id} já existe na planilha")

    # Monta array de valores na ordem exata das colunas (ignora chaves com _)
    row_values = []
    for col in TABLE_COLUMNS:
        if col == "PaymentID":
            row_values.append(pay_id)
        elif col in data and data[col] is not None:
            row_values.append(data[col])
        else:
            row_values.append("")

    url = f"{_table_url()}/rows/add"
    body = {"values": [row_values]}

    r = requests.post(url, headers=_headers(), json=body)
    r.raise_for_status()

    return pay_id


# ── Upload de recibo ──

def upload_receipt(local_path, pay_id: str) -> str:
    """
    Sobe o recibo para o OneDrive, renomeado como {pay_id}.{ext}.
    Retorna a URL do arquivo no OneDrive.
    """
    from pathlib import Path
    local_path = Path(local_path)
    ext = local_path.suffix.lower()
    dest_name = f"{pay_id}{ext}"
    dest_onedrive = f"{ONEDRIVE_PROCESSED_PATH}/{dest_name}"
    encoded = urllib.parse.quote(dest_onedrive)

    url = f"{GRAPH}/me/drive/root:/{encoded}:/content"

    content_types = {
        ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
        ".png": "image/png", ".webp": "image/webp",
        ".pdf": "application/pdf",
    }
    ct = content_types.get(ext, "application/octet-stream")

    headers = _headers()
    headers["Content-Type"] = ct

    with open(local_path, "rb") as f:
        data = f.read()

    r = requests.put(url, headers=headers, data=data)
    r.raise_for_status()

    item = r.json()
    web_url = item.get("webUrl", "")
    print(f"  ☁ Upload: {dest_name} → OneDrive ({len(data) // 1024} KB)")
    return web_url


# ── Sincronização OneDrive ──

# Preferência de extensão: manter originais (jpeg/jpg/pdf) sobre conversões (png/webp)
_EXT_PREF = {".jpeg": 0, ".jpg": 1, ".pdf": 2, ".png": 3, ".webp": 4}


def sync_onedrive_recibos(dry_run: bool = False) -> None:
    """
    Sincroniza a pasta OneDrive/RECIBOS com a tblPagamentos:
    - Remove arquivos órfãos (PAY IDs que não existem mais na tabela)
    - Remove duplicatas de extensão por PAY ID (mantém o formato mais original)

    Use depois de deletar linhas de teste da planilha, ou quando a pasta
    ficar fora de sincronia. Em uso normal não é necessário.
    """
    from collections import defaultdict

    print("\n🔄 Sincronizando OneDrive/RECIBOS com tblPagamentos...")

    # 1. PAY IDs que existem na planilha
    url = f"{_table_url()}/rows"
    r = requests.get(url, headers=_headers())
    r.raise_for_status()
    pay_ids_excel: set[str] = set()
    for row in r.json().get("value", []):
        vals = row.get("values", [[]])[0]
        if vals and isinstance(vals[0], str) and vals[0].startswith("PAY-"):
            pay_ids_excel.add(vals[0])

    print(f"  {len(pay_ids_excel)} PAY IDs na planilha")

    # 2. Arquivos em OneDrive/RECIBOS
    folder_encoded = urllib.parse.quote(ONEDRIVE_PROCESSED_PATH)
    r2 = requests.get(
        f"{GRAPH}/me/drive/root:/{folder_encoded}:/children?$select=name,id,size",
        headers=_headers()
    )
    r2.raise_for_status()
    files = r2.json().get("value", [])
    print(f"  {len(files)} arquivos em OneDrive/{ONEDRIVE_PROCESSED_PATH}")

    # 3. Agrupa por stem (PAY ID sem extensão)
    by_stem: dict[str, list] = defaultdict(list)
    for f in files:
        parts = f["name"].rsplit(".", 1)
        stem = parts[0]
        ext = "." + parts[1].lower() if len(parts) > 1 else ""
        by_stem[stem].append({**f, "_stem": stem, "_ext": ext})

    orphans = 0
    dupes = 0
    kept = 0

    for stem, items in sorted(by_stem.items()):
        is_orphan = stem not in pay_ids_excel

        if is_orphan:
            for item in items:
                orphans += 1
                if dry_run:
                    print(f"  [DRY-RUN] DELETE {item['name']} (órfão)")
                else:
                    rd = requests.delete(
                        f"{GRAPH}/me/drive/items/{item['id']}",
                        headers=_headers()
                    )
                    status = "✅" if rd.status_code == 204 else f"❌ {rd.status_code}"
                    print(f"  DELETE {item['name']} (órfão) {status}")
        elif len(items) > 1:
            # Duplicata: manter o de menor _EXT_PREF (mais original)
            items_sorted = sorted(items, key=lambda x: _EXT_PREF.get(x["_ext"], 99))
            keep_item = items_sorted[0]
            kept += 1
            for item in items_sorted[1:]:
                dupes += 1
                if dry_run:
                    print(f"  [DRY-RUN] DELETE {item['name']} (duplicata, mantendo {keep_item['name']})")
                else:
                    rd = requests.delete(
                        f"{GRAPH}/me/drive/items/{item['id']}",
                        headers=_headers()
                    )
                    status = "✅" if rd.status_code == 204 else f"❌ {rd.status_code}"
                    print(f"  DELETE {item['name']} (duplicata, mantendo {keep_item['name']}) {status}")

    prefix = "[DRY-RUN] " if dry_run else ""
    print(f"\n{prefix}Concluído: {orphans} órfãos removidos, {dupes} duplicatas removidas, {kept} com duplicata resolvida")
