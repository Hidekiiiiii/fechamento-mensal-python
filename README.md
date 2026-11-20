# Fechamento mensal automatizado em Python: do export do ERP ao pacote pronto, com as conferências feitas

Um comando lê as 6 exportações brutas do mês (balancete, razão, extrato bancário, contas a receber, contas a pagar e folha). Ele roda 8 conferências e propõe 4 ajustes de fechamento. No fim, entrega um pacote em Excel com veredito: **apresentável**, **com ressalvas** ou **não apresentável**.

O controller deixa de procurar erro e passa a revisar exceções.

> **Aviso:** o Grupo Atlas e todos os dados são fictícios, gerados por script com semente fixa. Nenhum dado real de empresa foi usado.

**Fechamento de out/25 (reextração):**
- **Resultado:** pacote com ressalvas. O EBITDA do mês vai de R$ 1,42 mi para R$ 0,96 mi depois dos ajustes propostos. Se as pendências já medidas se confirmarem, vai a R$ 1,07 mi.
- **Pendências de conciliação:** R$ 520,6 mil, divididas em banco, clientes, folha e saldos inesperados.
- **Fora do corte:** 7 lançamentos, que somam R$ 327,5 mil.
- **Como foi verificado:** os 12 problemas plantados na base aparecem com o valor exato, e um teste automático confere cada um.

## Em um comando

```
python -m fechamento
```

```
Fechamento de out/25 (extracao_2): COM RESSALVAS
  C1 Integridade do balancete   OK                    R$ 0,00
  C2 Razão × balancete          OK                    R$ 0,00
  C3 Conciliação bancária       Pendência       R$ 140.700,00
  C4 Subledgers × razão         Pendência       R$ 156.000,00
  C5 Folha × razão              Pendência        R$ 22.400,00
  C6 Corte e competência        Pendência       R$ 327.500,00
  C7 Saldo inesperado           Pendência       R$ 201.450,00
  C8 Variação atípica           Alerta          R$ 330.000,00
```

O mesmo comando gera o resumo de 5 linhas para o e-mail ao diretor, por regras e sem IA:

> 1. Fechamento de out/25 com ressalvas: C1 e C2 OK; 5 pendências (C3, C4, C5, C6 e C7) e 1 alerta (C8).
> 2. EBITDA do mês: R$ 1,42 mi antes dos ajustes e R$ 0,96 mi após os ajustes propostos (PDD, provisão de férias e 13º e despesas sem nota: −R$ 460,3 mil), contra R$ 1,74 mi em set/25. Se as pendências de C4, C5 e C6 se confirmarem, vai a R$ 1,07 mi.
> 3. Pendências de conciliação, em valor absoluto: R$ 520,6 mil (banco R$ 140,7 mil, clientes R$ 156 mil, folha R$ 22,4 mil e saldos inesperados R$ 201,5 mil). Corte: R$ 286 mil de competência futura a estornar e relançar e R$ 41,5 mil lançados depois do corte, a validar.
> 4. Caixa: 2 itens com a tesouraria: um cheque de R$ 48,7 mil em aberto há 9 dias úteis e um crédito não identificado de R$ 92 mil.
> 5. Fretes sobre vendas da Atlas Distribuição +48% contra set/25 (+R$ 330 mil), com justificativa pendente do gestor. Prazo: 12/11/2025.

## A dor que ele resolve

Num fechamento manual em Excel, a equipe gasta de 2 a 3 dias úteis só em conferências: somar o balancete, bater o razão com os subledgers, conciliar os bancos e montar as provisões em planilhas copiadas de um mês para o outro.

Os erros que escapam raramente são de conta. São de processo:
- fórmula com intervalo errado;
- saldo inicial que não bate com o mês anterior;
- lançamento de novembro dentro de outubro;
- provisão calculada com o salário antigo.

Aqui, cada uma dessas verificações é uma regra explícita, com tolerância e responsável.

## As 8 conferências

| # | Conferência | Regra | Se falhar |
|---|---|---|---|
| C1 | Integridade do balancete | Débitos = créditos; por conta, final = inicial + D − C; inicial = final do pacote anterior | **Bloqueia** |
| C2 | Razão × balancete | Débitos e créditos do razão por conta × unidade = movimento do balancete (tolerância de R$ 1) | **Bloqueia** |
| C3 | Conciliação bancária | Extrato ± itens em aberto = saldo contábil, por banco. Cada item é classificado: em trânsito, cheque não compensado, tarifa ou rendimento não contabilizado, não identificado | Pendência |
| C4 | Subledgers × razão | Aging de clientes e de fornecedores = contas do razão, por unidade | Pendência |
| C5 | Folha × razão | Rubricas da folha por centro de custo = o que a integração lançou nas contas de pessoal | Pendência |
| C6 | Corte e competência | Competência fora do mês ou lançamento depois da data de corte (D+3 úteis) | Pendência |
| C7 | Saldo inesperado | Natureza invertida, transitória com saldo, despesa recorrente sem lançamento | Pendência |
| C8 | Variação atípica | Conta de resultado com Δ contra o mês anterior acima de 30% **e** de R$ 200 mil | Alerta |

C1 e C2 bloqueiam porque, se o balancete não fecha ou não bate com o razão, nenhum número abaixo é confiável. As demais saem como ressalva, com o valor em R$.

## Os 4 ajustes propostos

Nada é lançado no ERP. Cada ajuste sai em partida dobrada, com memória de cálculo e uma coluna "Aprovado (S/N)" para o controller.

| # | Ajuste | Regra | Out/25 |
|---|---|---|---|
| A1 | Complemento da PDD | Aging por faixa de atraso × percentual da política, menos o saldo contábil | R$ 155 mil |
| A2 | Férias e 13º | Salário vigente: férias = salário/30 × dias × 4/3 × 1,34; 13º = salário × avos/12 × 1,34 | R$ 58,3 mil |
| A3 | Despesa sem nota | Despesa recorrente que não foi lançada no mês: provisiona o valor de referência | R$ 247 mil |
| A4 | Itens do banco | Tarifas e IOF vão para despesa financeira; rendimentos, para receita financeira | −R$ 4,3 mil e +R$ 18,6 mil |

## O que sai

Arquivo [`saida/2025-10/extracao_2/fechamento_2025-10.xlsx`](saida/2025-10/extracao_2/fechamento_2025-10.xlsx), com 9 abas:

1. **Status:** veredito, EBITDA antes e depois dos ajustes, efeito potencial das pendências ainda abertas, as 8 conferências com status, valor, responsável e prazo, arquivos lidos com hash e campos de assinatura.
2. **Conferências:** regra, resultado em uma frase, diferença, tolerância, status, responsável e prazo.
3. **Ajustes:** resumo com efeito no EBITDA, os lançamentos em partida dobrada e a memória de cálculo da PDD e das provisões.
4. **Balancete ajustado:** conta × unidade, com o saldo do ERP, os ajustes, o saldo ajustado e subtotais por classe.
5. **DRE:** o mês antes e depois dos ajustes, a variação contra o mês anterior e o acumulado do ano.
6. **Conciliação bancária:** uma por banco, no layout clássico, chegando a diferença zero.
7. **Subledgers:** clientes, fornecedores e folha × razão, além do aging por faixa.
8. **Exceções:** cada lançamento ou saldo marcado, com colunas em branco para justificativa e ação.
9. **Log:** parâmetros, tolerâncias, regras e versão exata (SHA-256) de cada entrada.

Saem também:
- o resumo do e-mail, em `.txt`;
- o balancete ajustado em CSV, no mesmo layout do pacote do mês anterior. Ele vira o saldo inicial do mês seguinte, então o próximo fechamento confere a abertura contra ele.

## Duas extrações: o bloqueio funciona

O gerador cria duas exportações do mesmo mês, como acontece na prática.

**Extração 1:** tem erros de extração e o pacote sai **não apresentável**.
- Uma conta da Logística em que o saldo final não fecha (R$ 12.480,00).
- Um lançamento retroativo em set/25, feito depois do fechamento, que mudou o saldo inicial de duas contas (R$ 37.215,40).
- Três lançamentos liberados depois da extração: estão no balancete, mas não no razão (R$ 64.300,00).

```
python -m fechamento --extracao 1
Fechamento de out/25 (extracao_1): NÃO APRESENTÁVEL
```

**Extração 2:** é a reextração corrigida e sai **com ressalvas**. As duas versões ficam em [`saida/2025-10/`](saida/2025-10/).

## Como sei que funciona: gabarito e armadilhas

A base traz 12 problemas plantados, com o valor esperado de cada um guardado em [`gerador/gabarito.json`](gerador/gabarito.json):

| # | Problema plantado | Encontrado por | Valor esperado |
|---|---|---|---|
| P1 | Saldo final de uma conta ≠ inicial + D − C | C1 (extração 1) | R$ 12.480,00 |
| P2 | Lançamento retroativo em set/25 depois do fechamento | C1 (extração 1) | R$ 37.215,40 |
| P3 | 3 lançamentos no balancete e fora do razão | C2 (extração 1) | R$ 64.300,00 |
| P4 | Cheque não compensado há 9 dias úteis; crédito não identificado | C3 | R$ 48.700 e R$ 92.000 |
| P5 | 7 tarifas e IOF e 1 rendimento no extrato, sem lançamento | C3 e A4 | −R$ 4.286,90 e +R$ 18.640,22 |
| P6 | 2 títulos cancelados no aging sem estorno contábil | C4 | R$ 156.000 |
| P7 | Horas extras do CD Sudeste fora do razão | C5 | R$ 22.400 |
| P8 | 3 lançamentos de nov/25 em out/25 e 4 depois do corte | C6 | R$ 286.000 + R$ 41.500 |
| P9 | Adiantamento na conta de fornecedores; transitória com saldo | C7 | R$ 73.000 e R$ 128.450 |
| P10 | Aluguel do CD e energia do armazém sem lançamento | C7 e A3 | R$ 247.000 |
| P11 | PDD contábil abaixo da política | A1 | R$ 155.000 |
| P12 | Dissídio de 6% da Logística fora das provisões | A2 | R$ 58.300 |

Há também 4 armadilhas, que **não podem** ser marcadas:
- 2 depósitos em trânsito de até 2 dias úteis (R$ 215 mil): são normais;
- 6 pares de estorno e relançamento dentro do mês;
- uma despesa que subiu 35%, mas só R$ 80 mil;
- uma nota de set/25 já provisionada, cuja provisão foi revertida no mês.

Os 26 testes em [`tests/test_fechamento.py`](tests/test_fechamento.py) cobrem:
- as 12 ocorrências, com valores exatos;
- as armadilhas, que não podem aparecer;
- os ajustes em partida dobrada;
- o balancete ajustado fechando em zero;
- o texto do e-mail;
- a integridade do Excel.

## Como rodar

Requer Python 3.11+.

```
pip install -r requirements.txt
python gerador/gerar_dados.py          # opcional: regenera a base (sai idêntica, semente fixa)
python -m fechamento                   # extração mais recente
python -m fechamento --extracao 1      # extração com erros
python -m unittest discover tests      # 26 testes
```

Para usar outra base, aponte os arquivos:

```
python -m fechamento --mes 2025-10 --entrada <pasta> --anterior <csv> --parametros <xlsx> --saida <pasta>
```

As exportações seguem o padrão de ERP brasileiro: CSV com `;`, decimal com vírgula e datas dd/mm/aaaa. As colunas obrigatórias são validadas na leitura.

## Estrutura

```
entradas/
  parametros.xlsx                 mantido pelo controller: datas, tolerâncias, contas, faixas de PDD, recorrentes, bancos
  2025-09/                        pacote do mês anterior (saldo inicial e base de variação)
  2025-10/extracao_1, extracao_2/ as 6 exportações do ERP
fechamento/
  leitura.py                      leitura e validação das entradas
  conferencias.py                 C1 a C8
  ajustes.py                      A1 a A4
  demonstrativos.py               balancete ajustado e DRE
  pacote.py                       o Excel formatado
  pendencias.py                   efeito potencial das pendências no EBITDA
  resumo.py                       o e-mail de 5 linhas
gerador/                          base fictícia e gabarito
saida/2025-10/                    pacotes gerados (as duas extrações)
tests/                            26 testes contra o gabarito
```

## Decisões de método

- **Regras explícitas, não "detecção de anomalia".** Cada regra pode ser lida e auditada, e tem tolerância e responsável. Os parâmetros ficam numa planilha que o controller mantém, não no código.
- **A automação propõe; o humano aprova.** Nenhum ajuste é gravado no ERP.
- **"Após ajustes" não é número final.** O pacote mostra à parte o efeito das pendências já medidas: competência futura, folha e títulos cancelados. Assim o diretor vê a faixa provável, e não um EBITDA que parece definitivo.
- **Convenção de sinal:** nos saldos, positivo é devedor e negativo é credor. Na DRE, receita é positiva e despesa é negativa.
- **Arredondamento meio para cima** nos textos, como no Excel (R$ 520.550 vira "R$ 520,6 mil"), e não o arredondamento bancário do Python.
- **Rastreabilidade:** o log grava o hash SHA-256 de cada arquivo lido. Assim dá para provar com qual versão das entradas o pacote foi gerado.

## Limitações

- Base fictícia: 1 mês, 4 unidades, cerca de 2.200 lançamentos. Não há consolidação societária, fiscal nem moeda estrangeira.
- A conciliação bancária casa os itens por documento e valor. Um ERP real pode exigir casamento aproximado, por data e valor.
- A provisão de férias usa os dias adquiridos informados pela folha. Ela não recalcula períodos aquisitivos.

## Tecnologias

Python 3.13, pandas, NumPy, openpyxl e XlsxWriter. Testes com `unittest`, sem dependências extras.

## Contexto

Este é o 3º projeto do portfólio. O [P2](https://github.com/Hidekiiiiii/painel-controladoria-streamlit) é o painel que analisa o resultado do Grupo Atlas depois do fechamento. Este projeto é a etapa anterior: fechar e conciliar o mês.
