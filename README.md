# Painel de Execução Financeira — IF Baiano CSI

Painel estático publicado pelo GitHub Pages com dados consolidados da planilha oficial do Campus Santa Inês.

## Atualização

O arquivo Excel permanece privado. O script lê as abas de 2025 e 2026, concilia o resultado de 2026 com a linha `Total Geral` e só então gera o `index.html` público.

```bash
python -m pip install -r requirements.txt
python scripts/generate_dashboard.py planilha.xlsx \
  --source-name "Planilha Execução Financeira atualizada.xlsx" \
  --source-modified "2026-09-14T21:26:05.492Z"
```

Se a conciliação falhar, nenhum painel válido é gerado e o processo termina com erro. O arquivo `dashboard-source.json` registra a versão da fonte e os totais usados na última geração.

Site: https://bigispectrus.github.io/planilhaIF/
