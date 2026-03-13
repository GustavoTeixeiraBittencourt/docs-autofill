# 📄 Sistema de Automação de Documentos- Doc-autofill

Automação de preenchimento em lote de documentos Word (`.docx`) a partir de planilhas de dados (`.xlsx` / `.csv`), com suporte a conversão automática para PDF. Desenvolvido para rodar no **Google Colab** com integração ao **Google Drive**.

---

## Funcionalidades

- **Preenchimento em lote** — gera um documento por linha da planilha, substituindo os placeholders `{CHAVE}` pelos valores correspondentes
- **Detecção inteligente de placeholders** — localiza marcadores mesmo em textos com formatação complexa, fragmentados entre múltiplos _runs_, e em cabeçalhos/rodapés
- **Mapeamento fuzzy** — relaciona automaticamente colunas da planilha aos placeholders do documento, com suporte a aliases customizáveis
- **Conversão DOCX → PDF** — célula auxiliar para converter os arquivos gerados em PDF usando LibreOffice
- **Modo Dry-Run** — simula o processamento sem gerar arquivos, útil para validação prévia
- **Logs detalhados** — registra mapeamentos, campos não preenchidos e erros por linha processada
- **Interface interativa** — menu no próprio Colab com opções de execução, debug, configuração e testes

---

## Estrutura do Projeto

```
.
├── docs-autofill.ipynb   # Notebook principal
└── README.md
```

### Células do notebook

| Célula       | Descrição                                                                                                                             |
| ------------ | ------------------------------------------------------------------------------------------------------------------------------------- |
| **Célula 1** | Sistema completo de automação — instalação de dependências, parsing de placeholders, renderização de documentos e execução interativa |
| **Célula 2** | Módulo auxiliar de conversão DOCX → PDF via LibreOffice                                                                               |

---

## 🚀 Como usar

### 1. Pré-requisitos

- Conta Google com acesso ao [Google Colab](https://colab.research.google.com)
- Google Drive para armazenar os arquivos gerados
- Um template `.docx` com placeholders no formato `{NOME_DO_CAMPO}`
- Uma planilha `.xlsx` ou `.csv` com os dados a preencher

### 2. Preparar o template

No seu arquivo `.docx`, insira os campos variáveis entre chaves:

```
Nome: {NOME_DO_INSTRUTOR}
CPF: {CPF_DO_INSTRUTOR}
Valor total: {VALOR_TOTAL}
Data: {DATA}
```

### 3. Preparar a planilha

As colunas da planilha devem corresponder aos placeholders. O sistema faz a correspondência automaticamente (com normalização de acentos, maiúsculas/minúsculas e underscores).

| NOME_DO_INSTRUTOR | CPF_DO_INSTRUTOR | VALOR_TOTAL | DATA       |
| ----------------- | ---------------- | ----------- | ---------- |
| João Silva        | 000.000.000-00   | R$ 1.500,00 | 01/01/2025 |
| Maria Souza       | 111.111.111-11   | R$ 2.000,00 | 05/01/2025 |

### 4. Executar no Colab

1. Abra o notebook no Google Colab
2. Execute a **Célula 1**
3. Escolha a opção `1 — Execução Normal` no menu
4. Faça upload do `.docx` modelo e da planilha quando solicitado
5. Os arquivos gerados serão salvos automaticamente no Google Drive em `/Sistema de Automacao/`

---

## Opções do Menu

| Opção | Descrição                                 |
| ----- | ----------------------------------------- |
| `1`   | Execução normal com upload de arquivos    |
| `2`   | Análise de placeholders (debug detalhado) |
| `3`   | Configuração avançada                     |
| `4`   | Ajuda e documentação                      |
| `5`   | Executar testes                           |
| `6`   | Sair                                      |

---

## Configuração Avançada

É possível configurar aliases para mapear nomes de colunas diferentes dos placeholders:

```python
aliases = {
    "NOME_DO_INSTRUTOR": "INSTRUTOR",   # placeholder → coluna na planilha
    "VALOR_TOTAL": "VALOR",
}
```

Outras configurações disponíveis via `ProcessingConfig`:

| Parâmetro          | Padrão            | Descrição                                                 |
| ------------------ | ----------------- | --------------------------------------------------------- |
| `dry_run`          | `False`           | Simula sem gerar arquivos                                 |
| `missing_policy`   | `"blank"`         | Comportamento para campos ausentes: `"blank"` ou `"keep"` |
| `timezone`         | `"America/Belem"` | Fuso horário para campos de data automáticos              |
| `date_format`      | `"%d/%m/%Y"`      | Formato de data                                           |
| `unique_filenames` | `True`            | Adiciona sufixo único para evitar sobrescrita             |

---

## Dependências

Instaladas automaticamente ao executar o notebook:

- `pandas`
- `openpyxl`
- `python-docx`
- `pytz`
- `rich`
- `unidecode`
- `libreoffice` _(apenas para conversão PDF, instalado via apt)_

---

## Saída

Os arquivos gerados são salvos no Google Drive no caminho:

```
/content/drive/MyDrive/Sistema de Automacao/
```

Com nomes no formato:

```
AAAAMMDD_DOC-<numero>_<instrutor>.docx
```

---

## Exemplo de Fluxo

```
template.docx  +  dados.xlsx
        │
        ▼
  [ Automação ]
        │
        ├─ 20250101_DOC-001_joao_silva.docx
        ├─ 20250101_DOC-002_maria_souza.docx
        └─ 20250101_DOC-003_carlos_lima.docx
```

---

## Licença

Este projeto está sob a licença [MIT](LICENSE).
