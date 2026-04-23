"""
Script de diagnóstico para tabelas do GESTAO_CURSOS_2026.xlsx via Graph API.
- Lista problemas comuns que causam erro nos relatórios (linhas em branco, texto em campos numéricos, competências/status fora do padrão, etc).
- Requer configuração correta do Recibo Agent (token, config.py, etc).
"""
from graph_client import _find_workbook_id, _headers, GRAPH
import requests
import re

# Tabelas a verificar
tables = [
    ("tblPagamentos", ["VALOR", "Competencia", "StatusPagamento", "AlunoBeneficiario"]),
    ("tblMatriculas", ["StudentName", "CourseName", "AnoLetivo"]),
    ("tblRateio", ["AlunoDepositante", "AlunoBeneficiario", "Competencia", "CompetenciaDestino", "ALOCAR_VALOR"]),
    ("tblContas", ["NomeConta", "TitularNome", "ChavesPix", "Tipo", "Ativo"]),
]

workbook_id = _find_workbook_id()

problems = []

def check_table(table, columns):
    url = f"{GRAPH}/me/drive/items/{workbook_id}/workbook/tables/{table}/rows"
    r = requests.get(url, headers=_headers())
    r.raise_for_status()
    rows = r.json().get("value", [])
    print(f"\n--- {table} ({len(rows)} linhas) ---")
    for i, row in enumerate(rows):
        vals = row.get("values", [[]])[0]
        if not any(str(v).strip() for v in vals):
            print(f"Linha {i+1}: linha em branco.")
            problems.append((table, i+1, "Linha em branco"))
        for cidx, col in enumerate(columns):
            if cidx >= len(vals):
                print(f"Linha {i+1}: coluna {col} ausente.")
                problems.append((table, i+1, f"Coluna {col} ausente"))
                continue
            v = vals[cidx]
            if col in ("VALOR", "ALOCAR_VALOR"):
                if isinstance(v, str) and (re.search(r"[A-Za-z]", v) or v.strip() == ""):
                    print(f"Linha {i+1}: {col} não numérico: '{v}'")
                    problems.append((table, i+1, f"{col} não numérico: '{v}'"))
            if col in ("Competencia", "CompetenciaDestino"):
                if v not in ("JAN", "FEV", "MAR", "ABR", "MAI", "JUN", "JUL", "AGO", "SET", "OUT", "NOV", "DEZ"):
                    print(f"Linha {i+1}: {col} fora do padrão: '{v}'")
                    problems.append((table, i+1, f"{col} fora do padrão: '{v}'"))
            if col == "StatusPagamento":
                if v not in ("OK", "Pendente", "Estornado", "Confirmado", "OK (DOCX)"):
                    print(f"Linha {i+1}: StatusPagamento inesperado: '{v}'")
                    problems.append((table, i+1, f"StatusPagamento inesperado: '{v}'"))
            if col == "Ativo":
                if v not in ("SIM", "NÃO", "NAO"):
                    print(f"Linha {i+1}: Ativo fora do padrão: '{v}'")
                    problems.append((table, i+1, f"Ativo fora do padrão: '{v}'"))
    print(f"--- Fim {table} ---")

if __name__ == "__main__":
    for table, columns in tables:
        check_table(table, columns)
    print("\nResumo de problemas encontrados:")
    for t, l, p in problems:
        print(f"Tabela {t}, linha {l}: {p}")
    if not problems:
        print("Nenhum problema encontrado!")
