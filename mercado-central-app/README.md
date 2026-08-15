# Mercado Central - Gestão de contratos e NFs

Este projeto é um MVP para acompanhar contratos, empresas e notas fiscais com uma tela web simples, banco de dados SQLite e fluxo de conferência visual.

## Objetivo

- cadastrar empresas e contratos;
- registrar notas fiscais;
- manter histórico;
- disponibilizar relatório para conferência e impressão;
- preparar a base para OCR e leitura de arquivos PDF em futuras etapas.

## Estrutura

- `app/` - aplicação web
- `templates/` - páginas HTML
- `static/` - arquivos CSS
- `data/` - banco SQLite
- `uploads/` - arquivos enviados

## Instalação

```bash
cd mercado-central-app
python -m venv .venv
.venv\Scripts\activate
python -m pip install --upgrade pip
pip install -r requirements.txt
```

## Execução

```bash
uvicorn app.main:app --reload --host 0.0.0.0 --port 8000
```

Acesse:

- http://localhost:8000

## Fluxo inicial

1. Cadastrar empresa
2. Cadastrar contrato
3. Cadastrar NF
4. Validar relatório e status
5. Preparar para OCR em etapas futuras

## Observação

Este projeto é um MVP funcional, desenhado para demonstrar a base do sistema e permitir testes do banco, das telas e do fluxo.
