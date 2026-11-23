"""Rodada 2 · gera a tabela de preço V2 a partir da planilha V1 enviada à Arcor.

Lógica da V2 (a mesma do deck):
  1. Retirada dos complementos que a V1 colocava em cada peça
     (defesa estratégica, mockup de contexto, AR, modelagem de loja, QR de rastreio): -RETIRADA
  2. Desconto comercial sobre a entrega central: -DESCONTO
  3. Refação adicional, a partir da 5ª: REFACAO_EXTRA do valor do item

Rodar (precisa de openpyxl):
  python -I fonte/precos_v2.py "<planilha V1.xlsx>"
Saídas: arcor-75lab-rodada2-tabela-precos.xlsx e fonte/precos_v2.json
"""
import json
import pathlib
import sys

import openpyxl

RETIRADA = 0.70      # parte do preço V1 que era complemento
DESCONTO = 0.30      # desconto comercial sobre a entrega central
REFACAO_EXTRA = 0.50  # por rodada, a partir da 5ª
FATOR = round((1 - RETIRADA) * (1 - DESCONTO), 4)

FEE_HORAS, FEE_VALOR_H, FEE_EXTRA_H = 65, 600, 700

RAIZ = pathlib.Path(__file__).resolve().parent.parent
v1 = pathlib.Path(sys.argv[1])

wb = openpyxl.load_workbook(v1)
ws = wb.active

linhas = []
tot_v1 = tot_v2 = 0.0
pecas = 0
for r in range(4, 104):
    nome = ws.cell(r, 1).value
    g = ws.cell(r, 7).value
    if isinstance(g, str) and '%' in g:  # linhas de alteração extra
        for col in (8, 9, 10):
            if isinstance(ws.cell(r, col).value, (int, float)):
                ws.cell(r, col).value = REFACAO_EXTRA
        continue
    item = {'linha': r, 'nome': (nome or '').strip(), 'unid': ws.cell(r, 2).value, 'niveis': []}
    for qcol, ucol in ((3, 7), (4, 9), (5, 11)):
        q = ws.cell(r, qcol).value or 0
        u1 = ws.cell(r, ucol).value
        if not isinstance(u1, (int, float)):
            continue
        u2 = round(u1 * FATOR, 2)
        ws.cell(r, ucol).value = u2
        tot_v1 += q * u1
        tot_v2 += q * u2
        pecas += q
        item['niveis'].append({'qtd': q, 'v1': u1, 'v2': u2})
    linhas.append(item)

# valores soltos na linha de totais (herança da planilha original)
for ucol in (7, 9, 11):
    u = ws.cell(104, ucol).value
    if isinstance(u, (int, float)):
        ws.cell(104, ucol).value = round(u * FATOR, 2)

pct = lambda x: f'{x * 100:.0f}%'
ws['A110'] = (f'Os valores acima, devem ser considerados até 04 alterações (a partir da 5ª alteração a mesma deverá ser '
              f'cobrada, conforme valor de alterações extras de artes: {pct(REFACAO_EXTRA)} do valor unitário do item).')
ws['A124'] = 'OBSERVAÇÕES DO FORNECEDOR — 75 LAB ESTRATÉGIA, DESIGN E PRODUÇÃO LTDA · RODADA 2 DE NEGOCIAÇÃO'
ws['A126'] = ('2. Os valores consideram a entrega central de cada peça mais até 04 refações inclusas. Saem do preço os '
              'complementos que a rodada 1 colocava em cada peça: defesa estratégica, mockup de contexto, realidade '
              'aumentada, modelagem de loja e QR code de rastreio. Eles passam a ser orçados à parte, projeto a projeto, '
              'fora do escopo deste contrato. Sobre a entrega central há desconto comercial aplicado.')
ws['A127'] = (f'3. Alterações extras de Artes e Alterações extras de Artes Trade: {pct(REFACAO_EXTRA)} do valor unitário '
              'da peça, por rodada, a partir da 5ª.')
ws['A128'] = (f'3.1. Proposta Alternativa · fee mensal: franquia de {FEE_HORAS} horas por mês a R$ {FEE_VALOR_H} por hora '
              f'(R$ {FEE_HORAS * FEE_VALOR_H:,.0f} por mês), hora extra a R$ {FEE_EXTRA_H}. No fee não há cobrança por '
              'refação: a refação consome horas da franquia. Encerradas as horas do mês, a Arcor escolhe aguardar a '
              'virada do mês, quando a franquia reinicia, ou contratar horas extras.').replace(',', '.')

saida = RAIZ / 'arcor-75lab-rodada2-tabela-precos.xlsx'
wb.save(saida)

resumo = {
    'retirada': RETIRADA, 'desconto': DESCONTO, 'fator': FATOR, 'refacao_extra': REFACAO_EXTRA,
    'pecas': pecas, 'total_v1': round(tot_v1, 2), 'total_v2': round(tot_v2, 2),
    'nucleo': round(tot_v1 * (1 - RETIRADA), 2),
    'fee': {'horas': FEE_HORAS, 'valor_h': FEE_VALOR_H, 'extra_h': FEE_EXTRA_H,
            'mes': FEE_HORAS * FEE_VALOR_H, 'ano': FEE_HORAS * FEE_VALOR_H * 12},
    'linhas': linhas,
}
(RAIZ / 'fonte' / 'precos_v2.json').write_text(json.dumps(resumo, ensure_ascii=False, indent=1))
print(f'peças {pecas} · V1 R$ {tot_v1:,.2f} · núcleo R$ {resumo["nucleo"]:,.2f} · V2 R$ {tot_v2:,.2f} · fator {FATOR}')
print(f'fee R$ {resumo["fee"]["mes"]:,.0f}/mês · R$ {resumo["fee"]["ano"]:,.0f}/ano')
print('->', saida)
