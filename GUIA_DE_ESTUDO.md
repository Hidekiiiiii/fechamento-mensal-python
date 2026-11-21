# Guia de estudo: o fechamento automatizado por dentro

Objetivo: explicar o projeto numa entrevista e fazer pequenas mudanças com segurança. Este é o projeto mais próximo do que você já faz no dia a dia, então a comparação com o Excel e com a rotina de fechamento vale em quase todo passo.

---

## 1. A ideia em uma frase

"O controller deixa de **procurar** erro e passa a **revisar exceções**."

As conferências que tomam 2 a 3 dias num fechamento manual viram **regras explícitas**. Cada uma tem tolerância, responsável e prazo, e roda em segundos. O resultado é um pacote em Excel com um dos três vereditos: **apresentável**, **com ressalvas** ou **não apresentável**.

## 2. Mapa dos arquivos

| Arquivo | Papel | Analogia no fechamento manual |
|---|---|---|
| `entradas/parametros.xlsx` | Datas, tolerâncias, responsáveis, faixas de PDD, despesas recorrentes, bancos | A planilha de premissas que o controller mantém |
| `entradas/2025-10/extracao_*/` | As 6 exportações do ERP | Os relatórios tirados do sistema |
| `entradas/2025-09/balancete_ajustado_2025-09.csv` | O pacote do mês anterior | O fechamento de setembro, base do saldo inicial |
| `fechamento/leitura.py` | Lê e valida as entradas e grava o hash de cada arquivo | Conferir se pegou o relatório certo |
| `fechamento/conferencias.py` | C1 a C8 | O checklist de fechamento |
| `fechamento/ajustes.py` | A1 a A4, em partida dobrada | As planilhas de provisão |
| `fechamento/demonstrativos.py` | Balancete ajustado e DRE | O balancete pós-ajustes |
| `fechamento/pendencias.py` | Efeito potencial no EBITDA do que ainda está aberto | A nota "se X se confirmar, o resultado vai a…" |
| `fechamento/pacote.py`, `resumo.py` | O Excel de 9 abas e o e-mail de 5 linhas | O pacote para a diretoria |
| `gerador/gerar_dados.py` + `gabarito.json` | Base fictícia com 12 problemas plantados e o valor esperado de cada um | Um mês de teste com resposta conhecida |

**O caminho de um comando** (`python -m fechamento`):
1. `__main__.py` acha a extração mais recente.
2. `processar()` (`fechamento/__init__.py`) faz, em ordem: ler → conferir (`C.rodar`) → propor ajustes (`A.propor`) → balancete ajustado → DRE → veredito → e-mail.
3. `gravar()` grava o Excel, o `.txt` do e-mail e o balancete ajustado em CSV.

**A regra que organiza o projeto:**
- **Os parâmetros ficam no Excel, não no código.** Mudar uma tolerância ou um percentual de PDD não exige programar.
- **Os ajustes são propostos, nunca lançados.** O controller aprova cada um na coluna "Aprovado (S/N)".

## 3. Leitura: o que pode dar errado antes de começar

`leitura.py` trata três problemas clássicos de exportação de ERP:
- **Formato brasileiro:** `sep=";"`, `decimal=","` e datas `dd/mm/aaaa`. Sem isso, "1.234,56" vira texto.
- **Códigos como texto:** conta, unidade e centro de custo são lidos como texto (`TEXTO`). Assim, a conta "1.1.01" não vira o número 1,101, e um código com zero à esquerda não perde o zero.
- **Colunas obrigatórias:** se a exportação veio sem uma coluna, o programa para com uma mensagem clara (`OBRIGATORIAS`), em vez de dar um erro obscuro lá na frente.

**Rastreabilidade:** o SHA-256 é uma "impressão digital" do arquivo. Se um único caractere mudar, o hash muda. O log do pacote grava o hash de cada entrada, então dá para provar com qual versão dos relatórios o fechamento foi feito. Na prática, isso responde a pergunta "você usou o razão de antes ou de depois da correção?".

## 4. As conferências, uma a uma

Cada conferência devolve um objeto `Conferencia` com status, diferença em R$, frase de resultado, responsável, prazo e uma tabela de detalhe. O **veredito** (`veredito()`) segue três regras:
- se alguma conferência diz **Bloqueia**, o pacote é **não apresentável**;
- se alguma tem **Pendência** ou **Alerta**, ele sai **com ressalvas**;
- senão, é **apresentável**.

| # | O que confere | Como, em pandas | Equivalente no Excel |
|---|---|---|---|
| **C1** | Débitos = créditos; final = inicial + D − C; inicial = final do pacote anterior | Contas linha a linha e um `merge` com o mês anterior | Coluna de checagem `=F2-(C2+D2-E2)` e `PROCV` no balancete de setembro |
| **C2** | Razão = balancete, por conta × unidade | `groupby` do razão e `merge` com o balancete | Tabela dinâmica do razão ao lado do balancete |
| **C3** | Extrato ± itens em aberto = saldo contábil | `merge(..., indicator=True)`, que mostra o que está só no livro e o que está só no banco | O `PROCV` cruzado entre extrato e razão da conciliação |
| **C4** | Aging = conta de clientes e de fornecedores, por unidade | Dois `groupby` lado a lado | Total do aging contra o saldo do balancete |
| **C5** | Rubricas da folha = contas de pessoal da integração, por centro de custo | `groupby` da folha e do razão (só `origem = INTEGRACAO_FOLHA`) | Resumo da folha contra o razão de pessoal |
| **C6** | Competência fora do mês ou lançamento depois do corte (05/11) | Compara o período de competência e a data de lançamento | Filtro por data no razão |
| **C7** | Natureza invertida, transitória com saldo, despesa recorrente sem lançamento | Regras sobre o balancete e a lista de recorrentes | Olhar o balancete procurando sinal estranho |
| **C8** | Conta de resultado com Δ > 30% **e** > R$ 200 mil contra o mês anterior | Duas condições com `&` | Formatação condicional com `E(...)` |

**Por que C1 e C2 bloqueiam e o resto não?** Se o balancete não fecha ou não bate com o razão, nenhum número abaixo é confiável. As demais são problemas **conhecidos e medidos**: o pacote pode ir para o diretor com a ressalva e o valor em R$.

**O `EPS = 0.005` (meio centavo).** O computador guarda decimais em binário, e 0,1 + 0,2 vira 0,30000000000000004. Comparar "diferença > 0" acusaria erro em contas que fecham. O código arredonda para 2 casas e ignora o que for menor que meio centavo.

### C3 em detalhe: a conciliação do Banco Alfa (out/25)

| | R$ |
|---|---|
| Saldo do extrato | 2.304.655,35 |
| (+) Depósitos em trânsito (no livro, ainda não no banco) | 128.000,00 |
| (−) Cheque não compensado | 48.700,00 |
| **= Saldo do banco ajustado** | **2.383.955,35** |
| Saldo contábil | 2.386.412,55 |
| (−) Tarifas e IOF do extrato sem lançamento | 2.457,20 |
| **= Saldo contábil ajustado** | **2.383.955,35** |
| **Diferença** | **0,00** |

É o layout clássico de conciliação, que você já conhece. A diferença da automação é que **cada item é classificado por regra**:
- **Item no livro e não no banco:** depósito em trânsito ou cheque não compensado. Vira **pendência** se estiver aberto há mais de 5 dias úteis. O cheque de R$ 48,7 mil está há 9.
- **Item no banco e não no livro:** tarifa e IOF, ou rendimento, viram o ajuste A4. O que não dá para identificar, como o crédito de R$ 92 mil no Banco Beta, vira pendência com a tesouraria.

### C6 em detalhe: a armadilha da provisão revertida

Uma nota de **setembro** lançada em outubro parece fora da competência. Mas ela já estava **provisionada em setembro**, e a provisão foi revertida em outubro: o efeito no mês é zero, e não é erro.

O código reconhece isso (linha 240): se existe, no mesmo mês, um débito na conta de provisões a pagar com a mesma unidade e o mesmo valor, o lançamento não é marcado. Sem essa regra, a conferência acusaria um falso positivo todo mês.

## 5. Os ajustes, com a conta feita

Todo ajuste sai em **partida dobrada** pela função `_par`: um débito e um crédito do mesmo valor. Se o valor for negativo, ela inverte os lados, e o complemento vira reversão.

**A1, complemento da PDD (unidade DIG):**

| Faixa | Saldo em aberto | % da política | PDD |
|---|---|---|---|
| A vencer | 2.542.502,00 | 0,5% | 12.712,51 |
| 1 a 30 dias | 315.046,00 | 2% | 6.300,92 |
| 31 a 90 dias | 280.965,00 | 10% | 28.096,50 |
| 91 a 180 dias | 180.306,00 | 50% | 90.153,00 |
| **PDD calculada** | | | **137.262,93** |
| (−) PDD contábil | | | 116.505,93 |
| **Ajuste** | | | **20.757,00** |

Lançamento: D despesa com PDD / C PDD. A faixa de cada título vem de `pd.cut`, que equivale a um `PROCV` com correspondência aproximada numa tabela de faixas. Somando DIG, DIST e SERV, o A1 dá **R$ 155 mil**.

**A2, férias e 13º (um funcionário de SERV):** salário de R$ 3.376,25, 11 dias de férias adquiridos, admitido antes de 2025 (10 avos de 13º até outubro), encargos de 34%.
- Férias = 3.376,25 ÷ 30 × 11 × 4/3 × 1,34 = **R$ 2.211,82**. O 4/3 é o terço constitucional.
- 13º = 3.376,25 × 10/12 × 1,34 = **R$ 3.770,15**.

O problema plantado (P12) é o dissídio de 6% da Logística, que não entrou nas provisões. O cálculo usa o **salário vigente** da folha, então a diferença aparece sozinha: **R$ 58,3 mil**, só na unidade SERV. Nas outras unidades, a diferença é de centavos, e o código não lança diferença menor que R$ 1,00.

**A3, despesa sem nota:** o aluguel do CD e a energia do armazém não foram lançados no mês (o C7 achou). A regra provisiona o valor de referência (**R$ 247 mil**) contra provisões a pagar, com estorno automático no mês seguinte.

**A4, itens do banco:** as tarifas vão para despesa financeira (−R$ 4,3 mil) e o rendimento para receita financeira (+R$ 18,6 mil). Esse ajuste não afeta o EBITDA, porque fica abaixo dele.

## 6. Do balancete ajustado ao EBITDA

`balancete_ajustado()` soma os ajustes ao balancete do ERP por conta × unidade. `dre()` monta a DRE pela `ESTRUTURA_DRE` (`plano.py`), com receita positiva e despesa negativa.

| | Out/25 |
|---|---|
| EBITDA do ERP | R$ 1,42 mi |
| Ajustes A1 + A2 + A3 | −R$ 460,3 mil |
| **EBITDA após ajustes propostos** | **R$ 0,96 mi** |
| Se as pendências medidas se confirmarem: C6 (+286 mil), C5 (−22,4 mil), C4 (−156 mil) | +R$ 107,6 mil |
| **Faixa provável** | **R$ 1,07 mi** |

Ponto que vale ouro na entrevista: **"após ajustes" não é o número final**. `pendencias.py` mostra à parte o efeito do que ainda está aberto. Assim o diretor vê a faixa provável, e não um número com falsa precisão.

**O encadeamento.** O balancete ajustado sai em CSV no mesmo layout do pacote de setembro. Em novembro, ele é o "mês anterior", e o C1 confere se o saldo inicial de novembro bate com ele. Um lançamento retroativo em outubro, feito depois do fechamento, é pego no mês seguinte. Foi exatamente o que aconteceu com setembro na extração 1: o retroativo mudou o saldo inicial de duas contas.

## 7. Duas extrações: o bloqueio na prática

`python -m fechamento --extracao 1` roda a primeira exportação, que veio com erros de extração:
- **C1 bloqueia** com R$ 86.910,80: a conta que não fecha (R$ 12.480,00) mais o retroativo de setembro, que aparece nas duas contas afetadas (2 × R$ 37.215,40).
- **C2 bloqueia** com R$ 64.300,00: 3 lançamentos estão no balancete, mas não no razão.
- O veredito é **NÃO APRESENTÁVEL**, e o programa termina com **código de saída 1**. Num agendamento automático, esse código interrompe o envio do pacote.

A extração 2, que é a reextração corrigida, sai **com ressalvas**.

## 8. Como sei que está certo: gabarito, armadilhas e testes

`python -m unittest discover tests`: **26 testes**.
- **Os 12 problemas plantados** são encontrados com o **valor exato** do `gabarito.json`. Não basta "achar algo": tem que ser R$ 156.000,00, e não R$ 155.999.
- **As 4 armadilhas** não podem ser marcadas:
  - depósitos em trânsito de 1 a 2 dias;
  - pares de estorno e relançamento;
  - uma despesa +35%, mas de só R$ 80 mil;
  - a nota de setembro com provisão revertida.

  Isso mede **falso positivo**, que no dia a dia é o que faz a equipe parar de confiar na ferramenta.
- **Identidades:** todo ajuste tem débito = crédito; o balancete ajustado soma zero; a memória do A2 soma exatamente o lançado.
- **O Excel:** as abas existem, o aging por faixa está preenchido, e o balancete do mês seguinte encadeia.

**Os testes pegam erro de verdade.** Testei duas mutações numa cópia:
- tirar a regra da provisão revertida do C6 derruba 3 testes;
- trocar o "**e**" do C8 por "**ou**" derruba o teste que exige que só o frete seja marcado.

## 9. Cinco exercícios

Faça numa **cópia** de `entradas/parametros.xlsx` e aponte para ela com `--parametros`, ou desfaça depois:

```
git checkout -- entradas saida fechamento
```

Os resultados citados foram conferidos rodando o código.

1. **Afrouxar o critério de variação (fácil).** Na aba `tolerancias`, troque a do C8 de `200000` para `50000`. O C8 passa de **1 conta** (frete, +48%) para **6**, e a armadilha "Viagens CORP +35%, R$ 80 mil" aparece. Discuta: por que o critério exige percentual **e** valor ao mesmo tempo?
2. **Encurtar o prazo do banco (fácil).** Na aba `regras`, troque `dias_uteis_em_aberto` de `5` para `1`. O depósito em trânsito de R$ 128 mil do Banco Alfa (2 dias úteis) vira pendência, e a pendência do C3 vai de R$ 140,7 mil para **R$ 268,7 mil**. O de R$ 87 mil do Banco Gama (1 dia) continua no prazo.
3. **Mudar a política de PDD (médio).** Na aba `pdd`, troque o percentual de "91 a 180 dias" de `0.5` para `0.3`. O efeito do A1 no EBITDA cai de −R$ 155 mil para **−R$ 33,4 mil**, porque em alguma unidade a PDD calculada fica abaixo da contábil e o ajuste vira **reversão**. Ache qual unidade na aba Ajustes.
4. **Mudar os encargos (médio).** Na aba `geral`, troque `taxa_encargos` de `0.34` para `0.30`. O A2 passa a **reverter** provisão em CORP, DIG e DIST, e só SERV continua com complemento (R$ 14,6 mil). Explique por que uma premissa errada de encargos gera um ajuste que parece legítimo.
5. **Quebrar uma regra de propósito (médio).** Em `fechamento/conferencias.py`, linha 324, troque o `&` depois de `limite_rs)` por `|`. Rode os testes e leia `test_c8_so_o_frete`: o que a mudança fez com o significado da regra? Desfaça depois.

## 10. Perguntas de entrevista

1. **O que a automação faz que a planilha não faz?** Ela roda as 8 conferências do mesmo jeito todo mês, em segundos, com tolerância e responsável. Cada exceção sai com o valor em R$, e o veredito impede mandar um pacote que não fecha.
2. **Por que regras explícitas e não detecção de anomalia com IA?** Porque o auditor e o controller precisam ler, entender e aprovar a regra. Cada uma tem tolerância documentada e pode ser ajustada no Excel de parâmetros.
3. **Como você sabe que as conferências estão certas?** Por uma base com 12 problemas plantados, valores exatos num gabarito e 4 armadilhas que não podem ser marcadas. Os 26 testes conferem tudo isso, e mutações no código fazem os testes falhar.
4. **Por que C1 e C2 bloqueiam e o resto não?** Se o balancete não fecha ou não bate com o razão, nenhum número é confiável. O resto é problema medido, que pode ir como ressalva com o valor.
5. **O que é o "efeito potencial das pendências"?** O EBITDA após ajustes não é o final. Mostro à parte quanto as pendências já medidas mexem no resultado se se confirmarem: de R$ 0,96 mi para R$ 1,07 mi.
6. **Como evita falso positivo?** Testando as armadilhas: o trânsito de até 5 dias úteis é normal, os estornos dentro do mês se anulam, a variação precisa ser grande em % e em R$, e a provisão revertida é reconhecida.
7. **Por que a automação não lança no ERP?** Porque o ajuste é uma proposta, com memória de cálculo. A responsabilidade contábil continua do controller, que aprova S/N.
8. **Como garante que usou os arquivos certos?** Pelo hash SHA-256 de cada entrada no log, e pelo balancete ajustado do mês anterior, que vira a base do saldo inicial do mês seguinte.
9. **O que mudaria numa empresa real?** Casamento aproximado na conciliação (por data e valor), leitura direto do banco de dados do ERP, e agendamento logo depois da extração, com o código de saída barrando o envio.
10. **Como você usou IA no projeto?** Escreva sua resposta: as regras vieram da sua experiência de fechamento; o que você definiu (conferências, tolerâncias, armadilhas), o que conferiu, e como o gabarito e os testes validaram o código.
