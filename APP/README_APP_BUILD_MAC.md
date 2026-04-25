# Build do Recibo Agent — Versão Auto Contida (macOS)

Este diretório contém os arquivos e instruções para empacotar o Recibo Agent como um aplicativo auto contido para macOS.

## Passos para empacotar (PyInstaller)

1. **Certifique-se de estar usando o Python da python.org (não Homebrew!)**
   - O Python precisa ter suporte ao Tkinter.

2. **Instale o PyInstaller:**
   ```sh
   pip install pyinstaller
   ```

3. **No terminal, execute:**
   ```sh
   cd /Volumes/Expansion/----- MAMI -----/recibo_agent
   pyinstaller --noconfirm --windowed --onefile --name "Recibo Agent" recibo_agent_gui.py
   ```
   - `--windowed`: não abre terminal junto (GUI pura)
   - `--onefile`: gera um único executável
   - O app será gerado em `dist/Recibo Agent`

4. **Inclua dependências extras:**
   - Se necessário, use o parâmetro `--add-data` para incluir arquivos de configuração, modelos OCR, etc.
   - Exemplo:
     ```sh
     pyinstaller --noconfirm --windowed --onefile --name "Recibo Agent" --add-data "token.json:." recibo_agent_gui.py
     ```

5. **Teste o app:**
   - Rode o executável gerado em `dist/Recibo Agent`.
   - Certifique-se de que o OCR (EasyOCR), PyMuPDF, python-docx e integração com o OneDrive funcionam.

6. **Distribuição:**
   - O arquivo pode ser movido para qualquer Mac com o mesmo chip (Intel/ARM).
   - Para máxima portabilidade, gere o app no mesmo tipo de máquina do usuário final.

---

## Observações
- O app não depende de Ollama/LLM.
- Todos os modelos e dependências Python são empacotados.
- Se precisar de modelos OCR grandes, pode ser necessário incluí-los manualmente (ver docs do EasyOCR).
- Para suporte a atualização automática, considere ferramentas como `pyupdater`.

---

Dúvidas? Consulte a DOCUMENTACAO.md ou peça ajuda ao desenvolvedor.
