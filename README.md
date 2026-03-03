# Sistema de Validação de Localização de Estoque

Aplicação web local (offline) para conferência guiada de estoque por localização.

## Como usar

1. Abra `index.html` no navegador (ou rode um servidor local).
2. Informe o nome do operador.
3. Importe uma planilha CSV (ou XLSX quando a biblioteca XLSX estiver disponível).
4. Faça a conferência localização por localização usando **Enter** ou o botão **Confirmar e Próximo**.
5. Ao finalizar, exporte o relatório em CSV ou XLSX.

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

## Observação

- O parser CSV atual espera vírgula como separador e não trata CSV com aspas complexas.
- Para importação/exportação XLSX é necessário disponibilizar a biblioteca `XLSX` no navegador.
