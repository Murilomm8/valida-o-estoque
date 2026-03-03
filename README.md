# Sistema de Validação de Localização de Estoque

Aplicação web local (offline) para conferência guiada de estoque por localização.

## Como usar

### Opção recomendada (suporta CSV + XLSX)

1. Rode `python3 server.py`.
2. Acesse `http://localhost:4173` no navegador.
3. Informe o nome do operador.
4. Importe a planilha CSV ou XLSX.
5. Faça a conferência localização por localização usando **Enter** ou o botão **Confirmar e Próximo**.
6. Ao finalizar, exporte o relatório em CSV ou XLSX.

### Abrindo apenas `index.html`

- Funciona para CSV.
- XLSX depende da disponibilidade de uma biblioteca XLSX no navegador.

## Regras implementadas

- Limpeza de localização (`*`, espaços extras, padronização em maiúsculas).
- Parsing no formato `LONGARINA ALTURA.POSIÇÃO`.
- Ordenação lógica por `longarina`, `altura`, `posição`.
- Agrupamento por localização com múltiplos SKUs.
- Sessão com operador, horário de início, progresso e timestamps por localização.
- Registro de divergências (diferente / não encontrado).
- Persistência no `localStorage`.

## Colunas aceitas na importação

- Localização: `Localização` / `Localizacao`
- SKU: `SKU`, `Codigo`
- Produto: `Produto`, `Descrição`
- Quantidade esperada: `Quantidade`, `Qtd`

## Observações

- O parser CSV atual espera vírgula como separador e não trata CSV com aspas complexas.
- O parser XLSX local lê a primeira aba da planilha.
