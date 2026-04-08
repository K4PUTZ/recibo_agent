"""
Processa um recibo (imagem ou PDF) usando EasyOCR + regex parser.
Sem dependência de LLM/Ollama — rápido e leve.
"""
import re
import subprocess
import tempfile
import unicodedata
from pathlib import Path
from datetime import datetime


from config import MESES, CONTAS, CONTA_PISTAS, ALUNOS_CONHECIDOS, PAGADORES_MAP

# Suporte a leitura de DOCX
def _read_docx_text(docx_path: Path) -> str:
    try:
        from docx import Document
    except ImportError:
        print("  ⚠ python-docx não instalado. Use: pip install python-docx")
        return ""
    try:
        doc = Document(str(docx_path))
        return "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
    except Exception as e:
        print(f"  ⚠ Erro ao ler DOCX: {e}")
        return ""

_ocr_reader = None


def _get_reader():
    global _ocr_reader
    if _ocr_reader is None:
        import easyocr
        print("  Carregando modelo OCR (primeira vez ~10s)...")
        _ocr_reader = easyocr.Reader(["pt", "en"], gpu=False, verbose=False)
    return _ocr_reader


def _pdf_first_page_to_image(pdf_path: Path) -> Path:
    tmp = Path(tempfile.mkdtemp()) / "page.png"
    try:
        import fitz
        doc = fitz.open(str(pdf_path))
        pix = doc[0].get_pixmap(dpi=200)
        pix.save(str(tmp))
        doc.close()
        return tmp
    except ImportError:
        pass
    try:
        out = tmp.with_name("page")
        subprocess.run(["pdftoppm", "-png", "-f", "1", "-l", "1", "-r", "200",
                        str(pdf_path), str(out)], check=True, capture_output=True)
        generated = out.with_name("page-1.png")
        if generated.exists():
            return generated
    except (FileNotFoundError, subprocess.CalledProcessError):
        pass
    raise RuntimeError("Instale PyMuPDF: pip install pymupdf")


def process_receipt(file_path: Path) -> dict | None:
    suffix = file_path.suffix.lower()
    temp_image = None
    try:
        if suffix in (".jpg", ".jpeg", ".png", ".webp"):
            # Copia para path temporário sem espaços — OpenCV não lida bem
            tmp = Path(tempfile.mkdtemp()) / f"recibo{suffix}"
            import shutil as _shutil
            _shutil.copy2(str(file_path), str(tmp))
            image_path = tmp
            temp_image = tmp
            reader = _get_reader()
            results = reader.readtext(str(image_path), detail=0)
            text = "\n".join(results)
            if not text.strip():
                print("  ⚠ OCR não extraiu texto")
                return None
            print(f"  OCR extraiu {len(results)} linhas")
            return _parse(text)
        elif suffix == ".pdf":
            image_path = _pdf_first_page_to_image(file_path)
            temp_image = image_path
            reader = _get_reader()
            results = reader.readtext(str(image_path), detail=0)
            text = "\n".join(results)
            if not text.strip():
                print("  ⚠ OCR não extraiu texto")
                return None
            print(f"  OCR extraiu {len(results)} linhas")
            return _parse(text)
        elif suffix == ".docx":
            text = _read_docx_text(file_path)
            if not text.strip():
                print("  ⚠ DOCX vazio ou ilegível")
                return None
            print(f"  DOCX lido com {len(text.splitlines())} linhas")
            # Preencher campos básicos, marcar como OK (DOCX)
            # Tenta extrair valor e data, mas não força
            valor = _extract_valor(text)
            data = _extract_data(text)
            obs = f"Recibo DOCX manual: {text.splitlines()[0][:60]}" if text else "Recibo DOCX manual"
            return {
                "AlunoBeneficiario": "",
                "DataPagamento": data.strftime("%Y-%m-%d") if data else None,
                "VALOR": valor,
                "Competencia": MESES[data.month - 1] if data else None,
                "Recibo": "OK (DOCX)",
                "ContaRecebimento": _extract_conta(text),
                "StatusPagamento": "Confirmado",
                "PagadorNome(Opcional)": "",
                "ObservacoesPagamento": obs,
                "_tx_id": _extract_tx_id(text),
            }
        else:
            print(f"  ⚠ Formato não suportado: {suffix}")
            return None
    except Exception as e:
        print(f"  ⚠ Erro: {e}")
        return None
    finally:
        if temp_image and temp_image.exists():
            temp_image.unlink()


def _parse(text: str) -> dict:
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    valor = _extract_valor(text)
    data = _extract_data(text)
    pagador = _extract_pagador(lines)
    destinatario = _extract_destinatario(lines)
    conta = _extract_conta(text, destinatario)
    metodo = _extract_metodo(text)
    tx_id = _extract_tx_id(text)
    competencia = MESES[data.month - 1] if data else None
    aluno = _match_aluno(pagador) if pagador else None

    obs_parts = []
    if metodo:
        obs_parts.append(metodo)
    if tx_id:
        obs_parts.append(f"TX:{tx_id}")
    if pagador and not aluno:
        obs_parts.append(f"Pagador não encontrado nos alunos: {pagador}")

    campos_faltando = []
    if not valor:
        campos_faltando.append("VALOR")
    if not data:
        campos_faltando.append("DATA")
    if campos_faltando:
        obs_parts.append(f"Não detectado: {', '.join(campos_faltando)}")

    # Recibo fica Pendente se não identificou o aluno ou faltou campo essencial
    recibo = "OK" if (aluno and valor and data) else "Pendente"

    return {
        "AlunoBeneficiario": aluno or "",
        "DataPagamento": data.strftime("%Y-%m-%d") if data else None,
        "VALOR": valor,
        "Competencia": competencia,
        "Recibo": recibo,
        "ContaRecebimento": conta,
        "StatusPagamento": "Confirmado",
        "PagadorNome(Opcional)": pagador or "",
        "ObservacoesPagamento": ", ".join(obs_parts),
        "_tx_id": tx_id,  # não é coluna da tabela — usado para deduplicação
    }


def _extract_valor(text: str) -> float | None:
    m = re.search(r"R\$\s*([\d.,]+)", text, re.IGNORECASE)
    if not m:
        m = re.search(r"\b(\d{1,3}(?:\.\d{3})*,\d{2})\b", text)
    if m:
        raw = m.group(1).strip()
        raw = re.sub(r"\.(?=\d{3})", "", raw)
        raw = raw.replace(",", ".")
        try:
            return float(raw)
        except ValueError:
            pass
    return None


def _extract_data(text: str) -> datetime | None:
    for pat, fmt in [
        (r"(\d{2})/(\d{2})/(\d{4})", "dmy"),
        (r"(\d{4})-(\d{2})-(\d{2})", "ymd"),
    ]:
        m = re.search(pat, text)
        if m:
            try:
                if fmt == "dmy":
                    return datetime(int(m.group(3)), int(m.group(2)), int(m.group(1)))
                else:
                    return datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)))
            except ValueError:
                continue
    return None


# Labels que precedem o destinatário nos comprovantes
_DESTINATARIO_LABELS = re.compile(
    r"^(para|destinat[aá]rio|recebedor)$",
    re.IGNORECASE,
)


def _extract_destinatario(lines: list[str]) -> str | None:
    """Extrai o nome/info do destinatário (seção Para)."""
    for i, line in enumerate(lines):
        if _DESTINATARIO_LABELS.match(line.strip()):
            for j in range(i + 1, min(i + 4, len(lines))):
                c = lines[j].strip()
                if len(c) > 3 and not re.match(r"^[\d\s.\-*/]+$", c):
                    return c
    return None


# Labels que precedem o nome do pagador nos comprovantes
_PAGADOR_LABELS = re.compile(
    r"^(de|pagador|remetente|origem|quem (?:enviou|pagou)|pago por)$",
    re.IGNORECASE,
)
_SKIP_LINE = re.compile(
    r"cpf|cnpj|ag[eê]ncia|conta|banco|institui[cç][aã]o|chave|^\*|realizado|para",
    re.IGNORECASE,
)


def _extract_pagador(lines: list[str]) -> str | None:
    for i, line in enumerate(lines):
        if _PAGADOR_LABELS.match(line.strip()):
            for j in range(i + 1, min(i + 4, len(lines))):
                c = lines[j].strip()
                if (len(c) > 4
                        and not re.match(r"^[\d\s.\-*/]+$", c)
                        and not _SKIP_LINE.search(c)):
                    return c
    return None


def _extract_conta(text: str, destinatario: str | None = None) -> str:
    t = text.lower()
    dest = (destinatario or "").lower()

    if "paypal" in t:
        return "Paypal"
    if re.search(r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}", text):
        return "CNPJ"

    # Verifica pistas (chave Pix, fragmentos de nome do destinatário)
    busca = t + " " + dest
    for pista, conta in CONTA_PISTAS.items():
        if pista in busca:
            return conta

    for conta in CONTAS:
        if conta.lower() in t:
            return conta
    return ""


def _extract_metodo(text: str) -> str:
    t = text.lower()
    if "pix" in t:
        return "PIX"
    if "paypal" in t:
        return "Paypal"
    if "ted" in t or "transferência" in t or "transferencia" in t:
        return "Transferência"
    if "boleto" in t:
        return "Boleto"
    return ""


def _extract_moeda(text: str) -> str:
    if "€" in text or "eur" in text.lower():
        return "€"
    if "us$" in text.lower() or ("$" in text and "r$" not in text.lower()):
        return "US$"
    return "R$"


def _normalizar(nome: str) -> list[str]:
    """Tokeniza nome em lowercase sem acentos, removendo sufixos de ano."""
    sem_acento = unicodedata.normalize("NFD", nome)
    sem_acento = "".join(c for c in sem_acento if unicodedata.category(c) != "Mn")
    return [p for p in sem_acento.lower().split() if not re.match(r"^2?0?\d{2}$", p)]


def _extract_tx_id(text: str) -> str | None:
    """Extrai o ID de transação PIX do comprovante (previne duplicatas)."""
    # Tenta via label 'ID da transação:'
    m = re.search(
        r"ID\s+da\s+transa[cç][aã]o\s*:?\s*[\n\s]*([A-Z0-9]{20,})",
        text, re.IGNORECASE,
    )
    if m:
        return m.group(1).upper()
    # Fallback: formato E2E ID do PIX — E + 8 dígitos ISPB + data + sufixo
    m = re.search(r"\b(E\d{15,}[A-Z0-9]{5,})\b", text)
    if m:
        return m.group(1)
    return None


def _match_aluno(nome: str) -> str | None:
    if not nome or not ALUNOS_CONHECIDOS:
        return None

    # 1. Lookup direto no mapa de pagadores cadastrados
    chave = nome.strip().upper()
    if chave in PAGADORES_MAP:
        return PAGADORES_MAP[chave]

    partes_pagador = _normalizar(nome)
    if not partes_pagador:
        return None

    melhor: str | None = None
    melhor_score = 0

    for aluno in ALUNOS_CONHECIDOS:
        pa = _normalizar(aluno)
        if not pa:
            continue

        # Exact match normalizado
        if partes_pagador == pa:
            return aluno

        # Score: quantos tokens do pagador batem com tokens do aluno
        comuns = sum(1 for t in partes_pagador if t in pa)
        # Bônus se primeiro e último nome coincidem
        bonus = 0
        if partes_pagador[0] == pa[0]:
            bonus += 1
        if len(partes_pagador) > 1 and len(pa) > 1 and partes_pagador[-1] == pa[-1]:
            bonus += 1

        score = comuns + bonus

        # Exige pelo menos 2 tokens comuns E primeiro nome igual para aceitar
        if score >= 2 and partes_pagador[0] == pa[0] and score > melhor_score:
            melhor_score = score
            melhor = aluno

    return melhor
