# Sistema de Validação de Localização de Estoque

Aplicação web local (offline) para conferência guiada de estoque por localização.

## Como usar (simples para tablet)

1. Abra `index.html` no navegador.
2. Informe o nome do operador.
3. Importe a planilha CSV ou XLSX.
4. Faça a conferência localização por localização usando **Enter** ou o botão **Confirmar e Próximo**.
5. Ao finalizar, exporte o relatório em CSV ou XLSX.
6. Use a seção de **Histórico** para exportar sessões anteriores em CSV/XLSX.

## Regras implementadas

- Limpeza de localização (`*`, espaços extras, barra inicial e padronização em maiúsculas).
- Parsing no formato `LONGARINA ALTURA.POSIÇÃO`.
- Ordenação lógica por `longarina`, `altura`, `posição`.
- Agrupamento por localização com múltiplos SKUs.
- Conferência de: produto, unidade por caixa, volume, validade, lote e quantidade.
- Sessão com operador, horário de início, progresso e timestamps por localização.
- Persistência no `localStorage` e histórico local de sessões.
- Linhas sem localização válida (ex.: títulos como "Expedição") são ignoradas automaticamente.

## Colunas aceitas na importação

- Obrigatórias: `Localização`, `SKU`, `Produto`, `Quantidade`
- Opcionais: `Unidade por caixa` (ou `Unidade F`), `Volume`, `Validade`, `Lote`

## Observações

- O parser CSV atual detecta `,` ou `;` como separador.
- A importação XLSX funciona no navegador sem servidor adicional (se navegador suportar APIs modernas).
- A exportação XLSX funciona offline (inclusive no histórico), sem depender de bibliotecas externas.
