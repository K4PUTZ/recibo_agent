# Recibo Agent — Documentação Completa

> **Para LLMs continuando este trabalho:** leia a seção [Estado atual e pendências](#estado-atual-e-pendências) primeiro.

## Estado atual e pendências

### O que está funcionando (abril 2026)
✅ OCR via EasyOCR (imagens JPEG/PNG/WEBP, PDFs via PyMuPDF, e recibos .docx via python-docx)
- ✅ Match de aluno: PAGADORES_MAP + fuzzy por tokens com normalização de acentos
- ✅ Detecção de conta recebedora: `tblContas` no Excel (aba CURSOS, J1:N6) com `load_contas()`
- ✅ Inserção na `tblPagamentos` via Microsoft Graph API
- ✅ Upload do recibo para OneDrive `/RECIBOS/{PAY-ID}.ext`
- ✅ Arquivo movido para `RECIBOS PROCESSADOS` após processamento (duplicatas com prefixo `DUP-`)
- ✅ Deduplicação por ID de transação PIX (campo `TX:E6xxx...` em ObservacoesPagamento)
- ✅ LaunchAgent macOS (autostart no login, Nice 10, KeepAlive)
- ✅ Aba RELATORIOS com 5 seções de fórmulas dinâmicas (ver abaixo)

### Pendências conhecidas
- ❌ `instalar_windows.bat` ainda referencia Ollama (engine antiga) — não foi atualizado para EasyOCR
- ❌ Aba RELATORIOS S2 (Realizado por Curso) pode ainda mostrar `#VALUE!` em alguns ambientes Excel Online — a fórmula usa `LET + FILTER + ISNUMBER(MATCH)` que requer Excel 365
- ⚠ 7 PAY IDs sem comprovante (marcados `Pendente`): PAY-000000022, 029, 056, 126, 130, 131, 142

### Últimas alterações relevantes (abril 2026)
- `audit_historico.py` adicionado: faz upload e renomeia os 147 recibos históricos (nomes numéricos → PAY-ID)
  - 145 arquivos enviados ao OneDrive e renomeados localmente
  - 2 pulados (.docx não suportado: 25.docx, 41.docx)
  - 8 PAY IDs sem arquivo: 5 corretamente Pendente; 1 OK PayPal (PAY-18, Maria Caputi)
  - PAY-126, PAY-130 corrigidos de OK → Pendente (estavam marcados incorretamente)
  - Typo `Paypall` → `Paypal` corrigido no PAY-000000018
- `sync_onedrive_recibos(dry_run=False)` adicionado a `graph_client.py`
- Flag `--sync-onedrive` adicionada a `run.py`
- `tblContas` criada em CURSOS!J1:N6 com colunas: NomeConta / TitularNome / ChavesPix / Tipo / Ativo
- `load_contas()` adicionado a `graph_client.py`, chamado no startup de `run.py`
- `CONTAS` e `CONTA_PISTAS` em `config.py` agora são listas vazias preenchidas em runtime
- Aba RELATORIOS reescrita com fórmulas dinâmicas em 5 seções

Suporte a recibos .docx adicionado em abril/2026:
      - Recibos .docx agora são lidos automaticamente (python-docx).
      - O texto é extraído e campos principais (valor, data, conta) são detectados se possível.
      - O pagamento é registrado como "OK (DOCX)" e o texto extraído vai para Observações.
      - Se não for possível extrair valor/data, o campo fica em branco para conferência manual.
      - Dependência: python-docx (adicionada ao requirements.txt).

---

## O que é

O **Recibo Agent** é um sistema Python que automatiza o processamento de comprovantes de pagamento (recibos) recebidos no WhatsApp ou por qualquer outro meio. Quando um arquivo é colocado em uma pasta específica, o sistema:

1. Lê o recibo com OCR (reconhecimento de texto)
2. Extrai os dados relevantes (nome, valor, data, conta recebedora, ID de transação)
3. Identifica a qual aluna da tabela o pagamento pertence
4. Checa duplicata pelo ID de transação PIX — aborta se já existir
5. Insere uma linha nova na planilha `GESTAO_CURSOS_2026.xlsx` no OneDrive
6. Faz upload do arquivo original para OneDrive/RECIBOS/, renomeado como `PAY-XXXXXXXXX.ext`
7. Move o arquivo local para `RECIBOS PROCESSADOS/`

Tudo isso acontece automaticamente em segundos, sem nenhuma interação manual.

---

## Estrutura de arquivos

```
recibo_agent/
├── run.py               # Ponto de entrada — modos de execução e watcher
├── processor.py         # OCR + extração de dados com regex
├── graph_client.py      # Comunicação com o Excel Online via Microsoft Graph API
├── auth.py              # Autenticação com a conta Microsoft (MSAL)
├── config.py            # Configurações centrais (caminhos, credenciais, colunas)
├── audit_historico.py   # Script avulso: faz upload/renomeia recibos históricos (N.ext → PAY-ID)
├── requirements.txt     # Dependências Python
├── .token_cache.json    # Token de autenticação (gerado automaticamente, não commitar)
├── instalar_mac.sh      # Instalador macOS
├── instalar_windows.bat # Instalador Windows (desatualizado — ainda referencia Ollama)
└── DOCUMENTACAO.md      # Este arquivo
```

---

## Fluxo completo passo a passo

```
Recibo cai na pasta "RECIBOS IN"
         │
         ▼
   watchdog detecta
         │
         ▼
      processor.py
      ┌─────────────────────────────┐
      │ 1. Converte PDF → imagem    │  (se necessário, via PyMuPDF)
      │ 2. EasyOCR lê a imagem      │  (~3 segundos, CPU)
      │ 3. Lê .docx (se aplicável)  │  (python-docx)
      │ 4. Regex extrai:            │
      │    - Valor (R$ XXX,XX)      │
      │    - Data (DD/MM/AAAA)      │
      │    - Nome do pagador        │
      │    - Método (PIX/TED/etc.)  │
      │    - Conta recebedora       │
      │ 5. Matching com tabela ALUNOS│
      └─────────────────────────────┘
         │
         ▼
   graph_client.py
   ┌─────────────────────────────┐
   │ 1. Gera próximo PaymentID   │  (PAY-000000XXX, lendo a tabela)
   │ 2. Insere linha na planilha │  (via Graph API, funciona com Excel aberto)
   │ 3. Faz upload do arquivo    │  (OneDrive/RECIBOS/PAY-000000XXX.jpeg)
   └─────────────────────────────┘
         │
         ▼
   Arquivo local deletado
```

---

## Módulos em detalhe

### `config.py` — Configurações

Arquivo central que define todos os parâmetros do sistema. É o único arquivo que precisa ser editado para adaptações.

| Variável | Descrição |
|---|---|
| `WATCH_FOLDER` | Pasta monitorada. Mac: `/Volumes/Expansion/----- MAMI -----/RECIBOS IN` |
| `PROCESSED_FOLDER` | Pasta de backup local (usado só se o upload falhar) |
| `ONEDRIVE_PROCESSED_PATH` | Pasta no OneDrive onde os recibos ficam: `RECIBOS` |
| `CLIENT_ID` | ID do App Registration no Azure (não alterar) |
| `WORKBOOK_ONEDRIVE_PATH` | Caminho da planilha no OneDrive |
| `TABLE_NAME` | Nome da tabela Excel: `tblPagamentos` |
| `TABLE_COLUMNS` | Colunas na ordem exata da tabela (11 colunas) |
| `MESES` | Lista de meses abreviados para campo Competência |
| `CONTAS` | Lista vazia — preenchida em runtime por `load_contas()` lendo `tblContas` |
| `CONTA_PISTAS` | Dict vazio — preenchido em runtime por `load_contas()`; chave→NomeConta |
| `ALUNOS_CONHECIDOS` | Lista preenchida em runtime com nomes da aba ALUNOS |
| `PAGADORES_MAP` | Dicionário preenchido em runtime: pagador → aluno (coluna PAGADOR da aba ALUNOS) |

---

### `processor.py` — OCR e extração de dados

**Leitura do arquivo:**
- Imagens (`.jpg`, `.jpeg`, `.png`, `.webp`): passadas direto para o OCR
- PDFs (`.pdf`): primeira página convertida para PNG via PyMuPDF (200 DPI) antes do OCR
- DOCX (`.docx`): texto extraído via python-docx; campos principais detectados por regex.

**OCR com EasyOCR:**
- Modelo rodando localmente, 100% offline, sem custo
- Treinado para português + inglês (`["pt", "en"]`)
- CPU-only — não precisa de placa de vídeo
- Primeira execução carrega o modelo (~10s); nas seguintes fica em memória

**Leitura de DOCX:**
- Utiliza python-docx para extrair texto de recibos .docx.
- Tenta identificar valor, data e conta automaticamente.
- Se não for possível identificar, o campo fica em branco e o texto vai para Observações.

**Extração de campos via regex:**

| Campo | Como detecta |
|---|---|
| **Valor** | Procura `R$ XXX,XX` ou padrão `000.000,00` |
| **Data** | Formatos `DD/MM/AAAA` e `AAAA-MM-DD` |
| **Pagador** | Linha logo após as labels `De`, `Pagador`, `Remetente`, `Origem`, `Pago por` |
| **Método** | Palavras-chave no texto: PIX, TED, Transferência, Boleto, Paypal |
| **Conta** | CNPJ (detectado pelo formato `XX.XXX.XXX/XXXX-XX`), ou palavras Flávia/Silvia/Paypal |
| **Competência** | Derivada do mês da data de pagamento (`ABR` para abril, etc.) |

**Matching de aluno:**

O nome extraído do recibo raramente é idêntico ao cadastro. O sistema faz:

1. **Lookup direto no mapa de pagadores** — se a coluna `PAGADOR` da aba ALUNOS tiver o nome do pagador cadastrado (ex: marido que paga no lugar da esposa), usa diretamente
2. **Normalização** — remove sufixos de ano como `26`, `2026` dos nomes cadastrados (ex: `Katia Puchaski 26` → `katia puchaski`)
3. **Score fuzzy** — conta quantos tokens (palavras) do pagador batem com tokens do aluno; exige mínimo de 2 tokens comuns E primeiro nome igual para aceitar
4. **Retorna o melhor match** com maior score


**Regra de status do Recibo:**
- `OK` — aluno identificado + valor + data encontrados
- `Pendente` — qualquer campo obrigatório ausente ou aluno não identificado

Quando fica `Pendente`, o campo `ObservacoesPagamento` explica o motivo (ex: `"Pagador não encontrado nos alunos: JOAO SILVA"`, `"Não detectado: VALOR"`), para facilitar a conferência manual.

---

### `graph_client.py` — Planilha Excel Online

Toda comunicação com o Excel Online usa a **Microsoft Graph API** — ou seja, a planilha pode estar aberta no browser enquanto o sistema insere dados, sem conflito.

**Funções principais:**

`load_contas()` ← chamada primeiro no startup
- Lê `tblContas` (aba CURSOS, J1:N6) via `/workbook/tables/tblContas/rows`
- Para cada conta ativa (coluna Ativo = "SIM"), popula:
  - `CONTAS`: lista de NomeConta
  - `CONTA_PISTAS`: dict {pista_lower → NomeConta} — cada item de ChavesPix (separado por vírgula) + primeiras 2 palavras do TitularNome

`load_alunos()` ← chamada depois de load_contas()
- Lê a aba `ALUNOS` (coluna `NomeCompleto`, índice 1)
- Preenche `ALUNOS_CONHECIDOS` para matching
- Lê coluna `PAGADOR` e monta `PAGADORES_MAP` (pagadores terceiros → nome da aluna)

`get_next_payment_id()`
- Lê todas as linhas da `tblPagamentos`
- Encontra o maior `PAY-XXXXXXXXX` e retorna o próximo
- **Colateral importante:** also popula `_known_tx_ids` (set global) com todos os IDs de transação `TX:E6xxx...` encontrados em ObservacoesPagamento — usado para deduplicação

`insert_payment(data)`
- Chama `get_next_payment_id()` (atualiza `_known_tx_ids`)
- Checa `data["_tx_id"]` contra `_known_tx_ids` — lança `DuplicateReceiptError` se já existir
- Monta array de 11 valores na ordem exata de `TABLE_COLUMNS`
- Faz `POST /workbook/tables/tblPagamentos/rows/add`

`upload_receipt(local_path, pay_id)`
- Faz `PUT` do arquivo para `OneDrive/RECIBOS/{pay_id}.ext`
- Suporta `.jpg`, `.jpeg`, `.png`, `.webp`, `.pdf`
- Retorna a webUrl do arquivo no OneDrive

`DuplicateReceiptError` — exceção lançada por `insert_payment()` quando TX ID já existe

`sync_onedrive_recibos(dry_run=False)`
- Compara arquivos em `OneDrive/RECIBOS/` com PAY IDs ativos na `tblPagamentos`
- **Remove órfãos**: arquivos cujo PAY ID não existe mais na planilha
- **Remove duplicatas de extensão**: se há `PAY-X.jpeg` e `PAY-X.png`, mantém o formato mais original (ordem de preferência: `.jpeg` > `.jpg` > `.pdf` > `.png` > `.webp`)
- Chamado por `run.py --sync-onedrive`; útil após deletar linhas de teste ou após o processo histórico

---

### `auth.py` — Autenticação Microsoft

Usa **MSAL (Microsoft Authentication Library)** com fluxo **Device Code**.

**Parâmetros críticos:**

| Parâmetro | Valor |
|---|---|
| `CLIENT_ID` | `c2618b42-a033-4db3-b193-31109c2fcb1b` |
| `AUTHORITY` | `https://login.microsoftonline.com/consumers` ← **obrigatório** para conta pessoal Microsoft — `common` ou `organizations` causam erro AADSTS9002346 |
| `SCOPES` | `["Files.ReadWrite.All"]` |
| `TOKEN_CACHE_PATH` | `recibo_agent/.token_cache.json` |
| Conta | `kantelira@outlook.com` |

**Fluxo:**
1. Primeira execução: exibe URL + código → usuário abre `microsoft.com/devicelogin` e autentica
2. Token salvo em `.token_cache.json` (não commitar)
3. Execuções seguintes: `acquire_token_silent()` renova via refresh token, sem interação
4. O token expira se ficar muitos meses sem uso — nesse caso o device flow é acionado novamente

**Se token expirar em produção:** o serviço LaunchAgent vai colocar o device code no log `/tmp/recibo_agent.log`. Ler com `tail -f /tmp/recibo_agent.log` e autenticar pelo browser.

---

### `run.py` — Modos de execução

```bash
# Modo normal: processa pendentes e entra em watch contínuo
python run.py

# Processa o que está na pasta e sai (sem ficar monitorando)
python run.py --once

# Processa um arquivo específico
python run.py --file /caminho/para/recibo.jpeg

# Simulação: mostra o que seria feito sem alterar nada
python run.py --dry-run

# Limpa OneDrive/RECIBOS: remove órfãos e duplicatas de extensão
python run.py --sync-onedrive
python run.py --sync-onedrive --dry-run  # apenas mostra o que seria feito

# Combinações
python run.py --file recibo.pdf --dry-run
python run.py --once --dry-run
```

**Modo watch** (padrão):
- Primeiro processa todos os arquivos que já estão na pasta
- Depois ativa o `watchdog` para monitorar novas entradas em tempo real
- Aguarda 3 segundos após detectar um arquivo novo (para garantir que o upload da OneDrive terminou antes de ler)
- Formatos aceitos: `.jpg`, `.jpeg`, `.png`, `.webp`, `.pdf`

---

### `audit_historico.py` — Migração de recibos históricos

Script avulso para situações em que recibos já processados ficaram com nomes numéricos simples (ex: `1.jpg`, `42.pdf`) e precisam ser vinculados aos PAY IDs da planilha e enviados ao OneDrive.

**Lógica:** arquivo `N.ext` → corresponde a `PAY-{N:09d}`. O PAY ID é verificado na tblPagamentos antes de qualquer ação.

```bash
# Dry-run: mostra o plano sem fazer nada
python audit_historico.py

# Upload para OneDrive + renomeia arquivos locais
python audit_historico.py --upload --rename

# Só upload, sem renomear localmente
python audit_historico.py --upload
```

**O script reporta:**
- Arquivos já presentes no OneDrive (pulados)
- Arquivos com formato não suportado (.docx etc.)
- PAY IDs que existem na planilha mas não têm arquivo local correspondente

**Uso único** — após a migração histórica de abril 2026 já foi executado. Os 147 arquivos foram processados (145 uploads, 2 .docx ignorados). 8 PAY IDs sem arquivo: 1 OK (PayPal - PAY-18), 7 Pendente.

---

## Tabela `tblPagamentos` — Campos preenchidos

| Coluna | Preenchido por | Observação |
|---|---|---|
| PaymentID | Sistema | Gerado automaticamente: PAY-000000XXX |
| AlunoBeneficiario | OCR + matching | Nome exato da aba ALUNOS; vazio se não identificado |
| DataPagamento | OCR | Formato `AAAA-MM-DD` |
| VALOR | OCR | Número decimal (ex: `321.0`) |
| Competencia | Derivado da data | Mês em 3 letras: JAN, FEV, MAR... |
| Recibo | Sistema | `OK` ou `Pendente` |
| ContaRecebimento | OCR + tblContas | NomeConta da conta detectada via CONTA_PISTAS, ou vazio |
| NF | — | Deixado em branco (preenchimento manual) |
| StatusPagamento | Sistema | Sempre `Confirmado` (recibo é evidência) |
| PagadorNome(Opcional) | OCR | Nome bruto como aparece no comprovante |
| ObservacoesPagamento | Sistema | `Arquivo: X | ID: TX:{e2e_id} | Banco: Y | Pasta: Z` — ou aviso de campo ausente |

---

## Pastas

| Pasta | Finalidade |
|---|---|
| `/Volumes/Expansion/----- MAMI -----/RECIBOS IN/` | Entrada — coloque os recibos aqui |
| `/Volumes/Expansion/----- MAMI -----/RECIBOS PROCESSADOS/` | Arquivo local após processamento (renomeado para PAY-ID, ou DUP-{nome} se duplicata) |
| `OneDrive/RECIBOS/` | Destino final dos recibos: `PAY-XXXXXXXXX.ext` |

---

## Autostart no macOS

O sistema está configurado para iniciar automaticamente quando você fizer login no Mac, via **LaunchAgent** do macOS.

**Localização do serviço:** `~/Library/LaunchAgents/com.mami.recibo_agent.plist`

**Comandos de gerenciamento:**
```bash
# Ver status (PID e exit code — 0 = rodando)
launchctl list | grep recibo_agent

# Ver log em tempo real
tail -f /tmp/recibo_agent.log

# Parar
launchctl stop com.mami.recibo_agent

# Iniciar
launchctl start com.mami.recibo_agent

# Desativar autostart permanentemente
launchctl unload ~/Library/LaunchAgents/com.mami.recibo_agent.plist

# Reativar autostart
launchctl load ~/Library/LaunchAgents/com.mami.recibo_agent.plist
```

**Impacto em performance:**
- Em idle (aguardando recibos): **0% CPU**, ~50 MB RAM
- Processando um recibo: ~3-5 segundos de CPU moderada, depois volta ao zero
- Configurado com `Nice 10` e `ProcessType Background` — o macOS automaticamente prioriza games, apps em primeiro plano e qualquer outra coisa sobre ele

---

## Casos especiais e limitações

**Pagador diferente da aluna** (ex: marido que paga)
- Solução: preencher a coluna `PAGADOR` na aba ALUNOS com o nome do pagador (em maiúsculas)
- O sistema fará o vínculo automaticamente no próximo processamento

**Recibo com layout diferente**
- O parser reconhece as labels `De`, `Pagador`, `Remetente`, `Origem`, `Quem enviou`, `Pago por`
- Se o recibo usar outro rótulo, o campo `AlunoBeneficiario` ficará em branco e o Recibo marcará `Pendente`
- O nome bruto sempre fica em `PagadorNome(Opcional)` para conferência manual

**PDF com múltiplas páginas**
- Apenas a primeira página é processada (que é onde fica o comprovante na maioria dos casos)

**Recibos .docx**
- O texto é extraído e os campos principais são detectados se possível.
- O pagamento é registrado como "OK (DOCX)" e o texto extraído vai para Observações.
- Se não for possível extrair valor/data, o campo fica em branco para conferência manual.
- Recibos .docx com layout não convencional podem exigir ajuste manual na planilha.

**Arquivo ilegível ou corrompido**
- O sistema registra o erro no log, não faz nada com a planilha, e o arquivo permanece na pasta de entrada para reprocessamento manual

**Competência vs. data de pagamento**
- A competência é derivada automaticamente do mês da data no recibo
- Se o pagamento for de um mês anterior (ex: pagar março em abril), é necessário corrigir manualmente no Excel após o processamento

---

## Dependências

```
easyocr      # OCR de imagens — modelo local, offline
pymupdf      # Conversão de PDF para imagem
msal         # Autenticação Microsoft
requests     # Chamadas à Graph API
watchdog     # Monitoramento de pasta em tempo real
certifi      # Certificados SSL para macOS/venv
python-docx  # Leitura de recibos .docx
```

Instalar: `pip install -r requirements.txt`

**Nota:** O suporte a .docx requer python-docx instalado (já incluso no requirements.txt a partir de abril/2026).

---

## Estrutura do Excel — GESTAO_CURSOS_2026.xlsx

**Workbook OneDrive path:** `2026/CURSOS 2026/TURMAS/GESTAO_CURSOS_2026.xlsx`  
**Workbook item ID:** `70B8F666C86445B8!s8658e0fef13e42c180df917ce44466ed`

| Aba | Índice | Uso |
|---|---|---|
| ALUNOS | 0 | NomeCompleto (col B), PAGADOR (col O) |
| CURSOS | 1 | tblCursos (A1:I6) + **tblContas (J1:N6)** |
| MATRICULAS | 2 | tblMatriculas — vínculo aluno→curso |
| PAGAMENTOS | 3 | **tblPagamentos** — destino das inserções |
| RATEIO | 4 | tblRateio — redistribuição de pagamentos entre alunos |
| ANIVERSARIOS | 5 | — |
| CONTROLE | 6 | Relatório operacional por aluno (fórmulas dinâmicas) |
| RELATORIOS | 7 | Relatório financeiro (5 seções — ver abaixo) |

### tblContas (CURSOS!J1:N6)

Gerenciada direto no Excel. O sistema lê em cada startup via `load_contas()`.

| NomeConta | TitularNome | ChavesPix | Tipo | Ativo |
|---|---|---|---|---|
| Mateus | Mateus de Oliveira e Souza Ribeiro | `30460639897, mateus de oliveira` | CPF | SIM |
| Silvia | Silvia Maciel Resende | `64800695000179, silvia maciel, 64.800.695` | CNPJ | SIM |
| Flávia | Flávia Betti de Oliveira e Souza | `59164158934, 591641589, flavia betti, flavia betti de o` | CPF | SIM |
| CNPJ | Flávia Betti (CNPJ) | `45236939000198, 45.236.939/0001-98, flavia betti cnpj` | CNPJ | SIM |
| Paypal | Flávia Betti de Oliveira e Souza | `paypal, flavia betti paypal, FLÁVIA BETTI` | Paypal | SIM |

**Como funciona a detecção de conta:** o texto completo do recibo + nome do destinatário são varridos contra todas as chaves de `CONTA_PISTAS`. A primeira correspondência define `ContaRecebimento`.

**Para adicionar/modificar contas:** editar diretamente as células J1:N6 no Excel. Na próxima execução o sistema carrega automaticamente.

### tblPagamentos (PAGAMENTOS) — 11 colunas

`PaymentID | AlunoBeneficiario | DataPagamento | VALOR | Competencia | Recibo | ContaRecebimento | NF | StatusPagamento | PagadorNome(Opcional) | ObservacoesPagamento`

### Aba RELATORIOS — 5 seções (fórmulas dinâmicas)

| Seção | Linhas | Conteúdo |
|---|---|---|
| S1 — Receita por Conta | 1–8 | SUMIFS por ContaRecebimento × mês (excluindo Estornados) |
| S2 — Realizado por Curso | 10–16 | `LET(stu, FILTER(...), SUMPRODUCT(ISNUMBER(MATCH(...))))` — inclui rateio; alunos multi-curso somam em cada curso |
| S3 — Esperado por Curso | 18–24 | Matrícula (mês início) + mensalidade (meses seguintes) por curso |
| S4 — Saldo a Receber | 26–32 | S3 − S2 por curso |
| S5 — MEI Flávia | 34–39 | CNPJ + Paypal + NF via CPF Flávia (sem duplicata) |

---

## Reinstalação / Novo computador

1. Instale Python 3.10+
2. Copie a pasta `recibo_agent/` (incluindo `.token_cache.json` se possível)
3. Crie um venv: `python3 -m venv .venv && source .venv/bin/activate`
4. Instale dependências: `pip install -r requirements.txt`
5. Execute uma vez no terminal para autenticar: `python run.py --once --dry-run`
6. Se não houver `.token_cache.json`: siga as instruções de login (código → `microsoft.com/devicelogin`, conta `kantelira@outlook.com`)
7. A partir daí o token fica salvo e não precisa mais logar
8. Para autostart no macOS: copie o plist para `~/Library/LaunchAgents/` e rode `launchctl load ~/Library/LaunchAgents/com.mami.recibo_agent.plist`

**LaunchAgent (macOS):**
- Plist: `~/Library/LaunchAgents/com.mami.recibo_agent.plist`
- Wrapper: `/usr/local/bin/recibo_agent_start.sh`
- Log: `/tmp/recibo_agent.log`
- `ProcessType: Background`, `Nice: 10`, `KeepAlive: true`, `ThrottleInterval: 10`
- Python: `/Volumes/Expansion/----- PESSOAL -----/PYTHON/.venv/bin/python`
