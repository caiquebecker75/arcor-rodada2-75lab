# Arcor × 75 LAB · Rodada 2 de negociação

Apresentação enxuta (10 telas, palco 1600×900) para defender a redução da proposta de Criação e Ajuste de Artes da Arcor de R$ 2,35 milhões para R$ 0,49 milhão por ano. Continua o deck da rodada 1 (`caiquebecker75/arcor-rfq-75lab`), com a mesma identidade.

A tese: a rodada 1 entregava o máximo da 75 LAB em cada peça (defesa estratégica, mockup de contexto, AR, modelagem de loja, QR de rastreio). A rodada 2 fica só com a entrega central, aplica desconto comercial e passa os complementos a orçamento à parte, projeto a projeto.

## Números
- Rodada 1: R$ 2.348.864 para 2.174 peças (planilha enviada em 29/10).
- Retirada dos complementos: −70% → entrega central R$ 704.659.
- Desconto comercial: −30% sobre a entrega central → R$ 493.261 (−79%).
- Formato 01 · tabela por peça: 4 refações inclusas, da 5ª em diante 50% do valor do item.
- Formato 02 · fee mensal: 65 h × R$ 600 = R$ 39.000 por mês (R$ 468 mil por ano), hora extra R$ 700. Refação consome a franquia; acabadas as horas, a Arcor espera a virada do mês ou paga hora extra.

Os percentuais ficam no topo de `fonte/precos_v2.py`. Mudou algum? Rode o script de novo e ajuste os números das telas.

## Arquivos
- `fonte/telas.html` · conteúdo das telas; `fonte/v2.css` · telas novas; `fonte/base.css` e `fonte/arcor.css` · identidade da rodada 1; `fonte/deck.js` · navegação.
- `python3 -I fonte/build.py` monta `index.html` e `assets/deck.css|js`.
- `python -I fonte/precos_v2.py "<planilha V1.xlsx>"` gera `arcor-75lab-rodada2-tabela-precos.xlsx` (as 101 linhas e os 3 níveis) e `fonte/precos_v2.json` (precisa de openpyxl).
- `node fonte/gerar-pdf.mjs http://localhost:4733/` gera `arcor-75lab-rodada2.pdf` (servir a pasta com `python3 -m http.server 4733`).

Navegação: setas do teclado, clique nas metades da tela, botões do rodapé ou o Sumário (tecla M).
