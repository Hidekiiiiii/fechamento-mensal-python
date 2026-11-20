"""Gera a base fictícia do fechamento de out/25 do Grupo Atlas e o gabarito dos problemas plantados.

Saídas:
- entradas/parametros.xlsx                             parâmetros mantidos pelo controller
- entradas/2025-09/balancete_ajustado_2025-09.csv      pacote do mês anterior (saldo inicial e base de variação)
- entradas/2025-10/extracao_1/*.csv                    primeira extração do ERP, com erros de extração
- entradas/2025-10/extracao_2/*.csv                    reextração corrigida
- gerador/gabarito.json                                valores que a automação precisa encontrar

Todos os dados são fictícios. A semente é fixa: a base sai idêntica a cada execução.
Convenção de sinal dos saldos: positivo = devedor, negativo = credor.
"""
from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd

RAIZ = Path(__file__).resolve().parents[1]
ENTRADAS = RAIZ / "entradas"
rng = np.random.default_rng(20251031)

MES = "2025-10"
COMP = "10/2025"
DATA_BASE = pd.Timestamp("2025-10-31")
DATA_CORTE = pd.Timestamp("2025-11-05")      # D+3 dias úteis
ENC = 0.34
DIAS = list(pd.bdate_range("2025-10-01", "2025-10-30"))
EBITDA_ALVO = 1_420_000.00

# --------------------------------------------------------------------------- cadastros
UNIDADES = [  # unidade, nome, cc da receita, cc de pessoal (ajuste A2), conta de pessoal
    ("DIST", "Atlas Distribuição", "DIST-COM", "DIST-ADM", "5.2.01"),
    ("DIG", "Atlas Digital", "DIG-OPS", "DIG-TEC", "5.2.01"),
    ("SERV", "Atlas Logística", "SERV-COM", "SERV-CDSE", "4.1.02"),
    ("CORP", "Corporativo", "CORP-FIN", "CORP-RH", "5.2.01"),
]
CCS = {  # centro de custo: (pessoas, salário médio, operação logística?)
    "DIST-COM": (45, 5200, False), "DIST-ARM": (90, 2900, False), "DIST-ADM": (15, 6800, False),
    "DIG-MKT": (20, 7200, False), "DIG-OPS": (35, 4300, False), "DIG-TEC": (30, 9800, False),
    "SERV-CDSE": (80, None, True), "SERV-CDNE": (70, 2800, True), "SERV-COM": (15, 6500, False),
    "CORP-DIR": (6, 32000, False), "CORP-FIN": (30, 7800, False), "CORP-RH": (18, 6200, False),
    "CORP-TI": (22, 8900, False), "CORP-FAC": (44, 2700, False),
}
BANCOS = {  # conta contábil: (banco, agência / conta, saldo de abertura, saldo final desejado no livro)
    "1.1.01.01": ("Banco Alfa", "0001 / 41.230-7", 2_850_000.00, 2_386_412.55),
    "1.1.01.02": ("Banco Beta", "1204 / 08.771-2", 1_620_000.00, 1_472_905.18),
    "1.1.01.03": ("Banco Gama", "0450 / 19.054-0", 940_000.00, 763_218.40),
}
CONTAS = [  # conta, descrição, grupo, natureza, tipo
    ("1.1.01.01", "Banco Alfa conta corrente", "Caixa e bancos", "D", "patrimonial"),
    ("1.1.01.02", "Banco Beta conta corrente", "Caixa e bancos", "D", "patrimonial"),
    ("1.1.01.03", "Banco Gama conta corrente", "Caixa e bancos", "D", "patrimonial"),
    ("1.1.02", "Aplicações financeiras", "Caixa e bancos", "D", "patrimonial"),
    ("1.1.03", "Clientes", "Contas a receber", "D", "patrimonial"),
    ("1.1.04", "(-) Perdas estimadas com clientes (PDD)", "Contas a receber", "C", "patrimonial"),
    ("1.1.05", "Estoques", "Estoques", "D", "patrimonial"),
    ("1.1.06", "Impostos a recuperar", "Outros ativos", "D", "patrimonial"),
    ("1.1.07", "Adiantamentos a fornecedores", "Outros ativos", "D", "patrimonial"),
    ("1.1.09", "Conta transitória", "Transitórias", "D", "transitoria"),
    ("1.2.01", "Imobilizado", "Imobilizado", "D", "patrimonial"),
    ("1.2.02", "(-) Depreciação acumulada", "Imobilizado", "C", "patrimonial"),
    ("2.1.01", "Fornecedores", "Fornecedores", "C", "patrimonial"),
    ("2.1.02", "Salários a pagar", "Obrigações trabalhistas", "C", "patrimonial"),
    ("2.1.03", "Encargos sociais a recolher", "Obrigações trabalhistas", "C", "patrimonial"),
    ("2.1.04", "Impostos a recolher", "Obrigações fiscais", "C", "patrimonial"),
    ("2.1.05", "Provisão de férias e encargos", "Obrigações trabalhistas", "C", "patrimonial"),
    ("2.1.06", "Provisão de 13º salário e encargos", "Obrigações trabalhistas", "C", "patrimonial"),
    ("2.1.08", "Provisões a pagar", "Provisões", "C", "patrimonial"),
    ("2.2.01", "Empréstimos e financiamentos", "Empréstimos", "C", "patrimonial"),
    ("2.3.01", "Capital social", "Patrimônio líquido", "C", "patrimonial"),
    ("2.3.02", "Lucros acumulados", "Patrimônio líquido", "C", "patrimonial"),
    ("3.1.01", "Venda de mercadorias", "Receita Bruta", "C", "resultado"),
    ("3.1.02", "Prestação de serviços logísticos", "Receita Bruta", "C", "resultado"),
    ("3.2.01", "Impostos sobre vendas", "Deduções", "D", "resultado"),
    ("3.2.02", "Devoluções e cancelamentos", "Deduções", "D", "resultado"),
    ("4.1.01", "Custo das mercadorias vendidas", "Custo", "D", "resultado"),
    ("4.1.02", "Custo dos serviços: pessoal operacional", "Custo", "D", "resultado"),
    ("4.1.03", "Custo dos serviços: ocupação de CD", "Custo", "D", "resultado"),
    ("5.1.01", "Marketing e mídia", "Despesas Comerciais", "D", "resultado"),
    ("5.1.02", "Comissões de vendas", "Despesas Comerciais", "D", "resultado"),
    ("5.1.03", "Comissão de marketplace", "Despesas Comerciais", "D", "resultado"),
    ("5.1.04", "Fretes sobre vendas", "Despesas Comerciais", "D", "resultado"),
    ("5.1.05", "Perdas estimadas com clientes (PDD)", "Despesas Comerciais", "D", "resultado"),
    ("5.2.01", "Salários", "Despesas com Pessoal", "D", "resultado"),
    ("5.2.02", "Encargos sociais", "Despesas com Pessoal", "D", "resultado"),
    ("5.2.03", "Benefícios", "Despesas com Pessoal", "D", "resultado"),
    ("5.3.01", "Aluguel", "Despesas Administrativas", "D", "resultado"),
    ("5.3.02", "Energia e utilidades", "Despesas Administrativas", "D", "resultado"),
    ("5.3.03", "Software e licenças", "Despesas Administrativas", "D", "resultado"),
    ("5.3.04", "Serviços de terceiros e consultoria", "Despesas Administrativas", "D", "resultado"),
    ("5.3.05", "Viagens", "Despesas Administrativas", "D", "resultado"),
    ("5.3.06", "Materiais e manutenção", "Despesas Administrativas", "D", "resultado"),
    ("6.1.01", "Depreciação e amortização", "D&A", "D", "resultado"),
    ("7.1.01", "Receitas financeiras", "Resultado Financeiro", "C", "resultado"),
    ("7.1.02", "Despesas financeiras", "Resultado Financeiro", "D", "resultado"),
    ("7.1.03", "Juros de antecipação de recebíveis", "Resultado Financeiro", "D", "resultado"),
]
TIPO = {c[0]: c[4] for c in CONTAS}
GRUPO = {c[0]: c[2] for c in CONTAS}
GRUPOS_EBITDA = {"Receita Bruta", "Deduções", "Custo", "Despesas Comerciais", "Despesas com Pessoal",
                 "Despesas Administrativas"}
RECORRENTES = [  # conta, centro de custo, descrição, valor de referência
    ("4.1.03", "SERV-CDSE", "Aluguel do CD Sudeste", 185_000.00),
    ("4.1.03", "SERV-CDNE", "Aluguel do CD Nordeste", 120_000.00),
    ("5.3.01", "CORP-FAC", "Aluguel da sede", 140_000.00),
    ("5.3.01", "DIST-ARM", "Aluguel do armazém da Distribuição", 160_000.00),
    ("5.3.02", "DIST-ARM", "Energia do armazém da Distribuição", 62_000.00),
    ("5.3.02", "SERV-CDSE", "Energia do CD Sudeste", 48_000.00),
    ("5.3.03", "DIG-TEC", "Licenças da plataforma de e-commerce", 95_000.00),
]
FALTANTES = {("4.1.03", "SERV-CDSE"), ("5.3.02", "DIST-ARM")}   # P10: recorrentes sem lançamento em out/25
PDD_FAIXAS = [("A vencer", -99999, 0, 0.005), ("1 a 30 dias", 1, 30, 0.02), ("31 a 90 dias", 31, 90, 0.10),
              ("91 a 180 dias", 91, 180, 0.50), ("Acima de 180 dias", 181, 99999, 1.00)]


def un(cc: str) -> str:
    return cc.split("-")[0]


def repartir(total: float, n: int, sigma: float = 0.5, passo: float = 0.01) -> list[float]:
    """Divide um total em n valores positivos, em múltiplos de `passo`, que somam exatamente o total."""
    w = rng.lognormal(0, sigma, n)
    v = np.round(np.floor(w / w.sum() * total / passo) * passo, 2)
    v[-1] = round(total - v[:-1].sum(), 2)
    assert v.min() > 0, (total, n)
    return [float(x) for x in v]


def dia(ate: int = len(DIAS)) -> pd.Timestamp:
    return DIAS[int(rng.integers(ate))]


# --------------------------------------------------------------------------- razão (lançamentos)
LINHAS: list[dict] = []
_seq = {"id": 250000}
_doc: dict[str, int] = {}


def doc(prefixo: str) -> str:
    _doc[prefixo] = _doc.get(prefixo, 0) + 1
    return f"{prefixo}-{_doc[prefixo]:06d}"


def D(conta: str, onde: str, v: float) -> tuple:
    return (conta, un(onde), onde if "-" in onde else "", round(v, 2), 0.0)


def C(conta: str, onde: str, v: float) -> tuple:
    return (conta, un(onde), onde if "-" in onde else "", 0.0, round(v, 2))


def lancar(data, linhas, historico, *, comp=COMP, origem="MANUAL", usuario="contab.ana", documento="", marca=""):
    _seq["id"] += 1
    assert abs(round(sum(l[3] for l in linhas) - sum(l[4] for l in linhas), 2)) < 0.005, historico
    for conta, unidade, cc, deb, cred in linhas:
        LINHAS.append(dict(id_lanc=_seq["id"], data_lanc=pd.Timestamp(data), data_competencia=comp, conta=conta,
                           unidade=unidade, centro_custo=cc, historico=historico, documento=documento,
                           debito=deb, credito=cred, usuario=usuario, origem=origem, marca=marca))
    return _seq["id"]


def banco() -> str:
    return str(rng.choice(list(BANCOS), p=[0.5, 0.3, 0.2]))


# ---- vendas, impostos, devoluções e recebimentos
VENDAS = {"DIST": (13_100_000, 140, "3.1.01", "DIST-COM"), "DIG": (4_300_000, 80, "3.1.01", "DIG-OPS"),
          "SERV": (2_200_000, 40, "3.1.02", "SERV-COM")}
IMPOSTOS = {"DIST": 0.125, "DIG": 0.125, "SERV": 0.0925}
for u, (total, n, conta, cc) in VENDAS.items():
    for v in repartir(total, n):
        nf = doc("NF")
        lancar(dia(), [D("1.1.03", u, v), C(conta, cc, v)], f"Venda {nf}", origem="INTEGRACAO_FATURAMENTO",
               usuario="integracao", documento=nf)
    imp = round(total * IMPOSTOS[u], 2)
    lancar("2025-10-31", [D("3.2.01", cc, imp), C("2.1.04", u, imp)], "Apuração de impostos sobre vendas out/25",
           origem="INTEGRACAO_FISCAL", usuario="integracao", documento=doc("APU"))
for u, cc, total, n in [("DIST", "DIST-COM", 180_000, 8), ("DIG", "DIG-OPS", 220_000, 12)]:
    for v in repartir(total, n):
        dv = doc("DEV")
        lancar(dia(), [D("3.2.02", cc, v), C("1.1.03", u, v)], f"Devolução de venda {dv}",
               origem="INTEGRACAO_FATURAMENTO", usuario="integracao", documento=dv)
for u, total, n in [("DIST", 12_700_000, 130), ("DIG", 4_150_000, 60), ("SERV", 2_150_000, 40)]:
    for v in repartir(total, n):
        b, cb = banco(), doc("COB")
        lancar(dia(), [D(b, "CORP", v), C("1.1.03", u, v)], f"Recebimento de cliente {cb}",
               origem="INTEGRACAO_COBRANCA", usuario="integracao", documento=cb)
# armadilha: 2 depósitos em trânsito de até 2 dias úteis (R$ 215 mil)
for data, u, b, v in [("2025-10-29", "DIST", "1.1.01.01", 128_000.00), ("2025-10-30", "SERV", "1.1.01.03", 87_000.00)]:
    cb = doc("COB")
    lancar(data, [D(b, "CORP", v), C("1.1.03", u, v)], f"Recebimento de cliente {cb}", origem="INTEGRACAO_COBRANCA",
           usuario="integracao", documento=cb, marca="transito")

# ---- compras e CMV
for u, total, n in [("DIST", 6_300_000, 60), ("DIG", 1_700_000, 25)]:
    for v in repartir(total, n):
        nf = doc("NFE")
        lancar(dia(), [D("1.1.05", u, v), C("2.1.01", u, v)], f"Compra de mercadorias {nf}",
               origem="INTEGRACAO_COMPRAS", usuario="integracao", documento=nf)
for u, cc, total in [("DIST", "DIST-ARM", 4_640_000), ("DIG", "DIG-OPS", 1_650_000)]:
    dias_cmv = DIAS + [pd.Timestamp("2025-10-31")]
    for d_, v in zip(dias_cmv, repartir(total, len(dias_cmv), sigma=0.2)):
        marca = "cmv_ultimo" if (u == "DIST" and d_ == dias_cmv[-1]) else ""
        lancar(d_, [D("4.1.01", cc, v), C("1.1.05", u, v)], f"Baixa de estoque: CMV de {d_:%d/%m}",
               origem="INTEGRACAO_ESTOQUE", usuario="integracao", marca=marca)

# ---- despesas (conta, cc, total do mês, nº de notas, descrição)
DESPESAS = [
    ("5.1.02", "DIST-COM", 260_000, 6, "Comissões de representantes"),
    ("5.1.04", "DIST-ARM", 1_020_000, 30, "Frete de entrega"),
    ("5.1.01", "DIST-COM", 90_000, 4, "Material de ponto de venda"),
    ("5.3.06", "DIST-ARM", 85_000, 8, "Manutenção de equipamentos"),
    ("5.3.04", "DIST-ADM", 70_000, 3, "Assessoria tributária"),
    ("5.3.05", "DIST-COM", 45_000, 6, "Viagens comerciais"),
    ("5.3.01", "DIST-ARM", 160_000, 1, "Aluguel do armazém da Distribuição"),
    ("5.1.01", "DIG-MKT", 520_000, 15, "Mídia digital"),
    ("5.1.03", "DIG-OPS", 310_000, 8, "Comissão de marketplace"),
    ("5.1.04", "DIG-OPS", 240_000, 10, "Frete de e-commerce"),
    ("5.3.03", "DIG-TEC", 95_000, 1, "Licenças da plataforma de e-commerce"),
    ("5.3.04", "DIG-TEC", 60_000, 3, "Desenvolvimento terceirizado"),
    ("4.1.03", "SERV-CDNE", 120_000, 1, "Aluguel do CD Nordeste"),
    ("5.3.02", "SERV-CDSE", 48_000, 1, "Energia do CD Sudeste"),
    ("5.3.06", "SERV-CDSE", 70_000, 6, "Manutenção de empilhadeiras"),
    ("5.3.06", "SERV-CDNE", 50_000, 5, "Materiais de embalagem"),
    ("5.1.02", "SERV-COM", 60_000, 3, "Comissões comerciais"),
    ("5.3.01", "CORP-FAC", 140_000, 1, "Aluguel da sede"),
    ("5.3.02", "CORP-FAC", 35_000, 1, "Energia da sede"),
    ("5.3.03", "CORP-TI", 110_000, 5, "Licenças de software"),
    ("5.3.04", "CORP-FIN", 90_000, 4, "Consultoria contábil"),
    ("5.3.05", "CORP-DIR", 310_000, 12, "Viagens da diretoria"),
    ("5.3.06", "CORP-FAC", 40_000, 4, "Materiais de escritório"),
]
ids_despesa = []
for conta, cc, total, n, desc in DESPESAS:
    for v in repartir(total, n, sigma=0.4):
        nf = doc("NFS")
        i = lancar(dia(19), [D(conta, cc, v), C("2.1.01", cc, v)], f"{desc} {nf}", origem="INTEGRACAO_COMPRAS",
                   usuario="integracao", documento=nf)
        if n > 1:
            ids_despesa.append(i)
# armadilha: 6 pares de estorno e relançamento dentro do mês (efeito líquido zero)
for i in rng.choice(ids_despesa, 6, replace=False):
    orig = [l for l in LINHAS if l["id_lanc"] == i]
    d_ = orig[0]["data_lanc"] + pd.offsets.BDay(1)
    inv = [(l["conta"], l["unidade"], l["centro_custo"], l["credito"], l["debito"]) for l in orig]
    mesmo = [(l["conta"], l["unidade"], l["centro_custo"], l["debito"], l["credito"]) for l in orig]
    lancar(d_, inv, f"Estorno: {orig[0]['historico']} (histórico incorreto)", documento=orig[0]["documento"])
    lancar(d_, mesmo, f"Relançamento: {orig[0]['historico']}", documento=orig[0]["documento"])

# ---- folha de pagamento (out/25)
pessoas = []
mat = 10000
for cc, (n, media, op) in CCS.items():
    for k in range(n):
        mat += 1
        adm = pd.Timestamp("2015-01-05") + pd.Timedelta(days=int(rng.integers(0, 3650)))
        if rng.random() < 0.06:
            adm = pd.Timestamp(2025, int(rng.integers(2, 10)), int(rng.integers(1, 28)))
        pessoas.append(dict(matricula=mat, centro_custo=cc, salario=0.0, he=0.0,
                            beneficios=round(float(rng.uniform(650, 1100)), 2),
                            data_admissao=adm, data_desligamento=pd.NaT,
                            dias_ferias_adquiridos=int(rng.integers(0, 56))))
        if media:
            pessoas[-1]["salario"] = round(float(media * rng.lognormal(0, 0.18)), 2)
        if cc in ("SERV-CDNE", "DIST-ARM") and rng.random() < 0.3:
            pessoas[-1]["he"] = round(float(rng.uniform(150, 900)), 2)
for p in rng.choice([p for p in pessoas if p["centro_custo"] not in ("SERV-CDSE",)], 4, replace=False):
    p["data_desligamento"] = DATA_BASE
# SERV-CDSE: folha total de R$ 412.300,00, com R$ 22.400,00 de horas extras (P7)
cdse = [p for p in pessoas if p["centro_custo"] == "SERV-CDSE"]
for p, v in zip(cdse[:18], repartir(22_400.00, 18, sigma=0.3)):
    p["he"] = v
s_alvo = (412_300.00 - 22_400.00 * (1 + ENC) - sum(p["beneficios"] for p in cdse)) / (1 + ENC)
for p, v in zip(cdse, repartir(round(s_alvo, 2), len(cdse), sigma=0.15)):
    p["salario"] = v
for p in pessoas:
    p["encargos"] = round((p["salario"] + p["he"]) * ENC, 2)
resto = round(412_300.00 - sum(p["salario"] + p["he"] + p["encargos"] + p["beneficios"] for p in cdse), 2)
cdse[0]["beneficios"] = round(cdse[0]["beneficios"] + resto, 2)
assert round(sum(p["salario"] + p["he"] + p["encargos"] + p["beneficios"] for p in cdse), 2) == 412_300.00

folha_por_cc = {}
for cc, (n, media, op) in CCS.items():
    g = [p for p in pessoas if p["centro_custo"] == cc]
    sal = round(sum(p["salario"] for p in g), 2)
    he = round(sum(p["he"] for p in g), 2)
    enc = round(sum(p["encargos"] for p in g), 2)
    ben = round(sum(p["beneficios"] for p in g), 2)
    folha_por_cc[cc] = round(sal + he + enc + ben, 2)
    he_lanc = 0.0 if cc == "SERV-CDSE" else he          # P7: integração não levou as horas extras do CD Sudeste
    u = un(cc)
    if op:
        deb = [D("4.1.02", cc, sal + he_lanc), D("4.1.02", cc, enc), D("4.1.02", cc, ben)]
    else:
        deb = [D("5.2.01", cc, sal + he_lanc), D("5.2.02", cc, enc), D("5.2.03", cc, ben)]
    lancar("2025-10-31", deb + [C("2.1.02", u, sal + he_lanc), C("2.1.03", u, enc), C("2.1.01", u, ben)],
           f"Folha de pagamento out/25: {cc}", origem="INTEGRACAO_FOLHA", usuario="integracao", documento=doc("FOL"))
    ativos = round(sum(p["salario"] for p in g if pd.isna(p["data_desligamento"])), 2)
    fer, dec = round(ativos / 12 * 4 / 3 * (1 + ENC), 2), round(ativos / 12 * (1 + ENC), 2)
    lancar("2025-10-31", [D("4.1.02" if op else "5.2.01", cc, fer + dec), C("2.1.05", u, fer), C("2.1.06", u, dec)],
           f"Provisão mensal de férias e 13º: {cc}")
razao_tmp = pd.DataFrame(LINHAS)
for u in ["DIST", "DIG", "SERV", "CORP"]:
    m = razao_tmp["unidade"] == u
    sal_out = razao_tmp.loc[m & (razao_tmp["conta"] == "2.1.02"), "credito"].sum()
    enc_out = razao_tmp.loc[m & (razao_tmp["conta"] == "2.1.03"), "credito"].sum()
    v = round(float(sal_out * rng.uniform(0.97, 0.995)), 2)
    lancar("2025-10-07", [D("2.1.02", u, v), C("1.1.01.01", "CORP", v)], f"Pagamento da folha de set/25: {u}",
           documento=doc("FOL"))
    v = round(float(enc_out * rng.uniform(0.97, 0.995)), 2)
    lancar("2025-10-20", [D("2.1.03", u, v), C("1.1.01.02", "CORP", v)], f"Recolhimento de INSS e FGTS de set/25: {u}",
           documento=doc("GPS"))
for u in VENDAS:
    v = round(VENDAS[u][0] * IMPOSTOS[u] * 0.98, 2)
    lancar("2025-10-24", [D("2.1.04", u, v), C("1.1.01.01", "CORP", v)], f"Recolhimento de tributos de set/25: {u}",
           documento=doc("DARF"))

# ---- pagamentos a fornecedores
for u, total, n in [("DIST", 7_350_000, 90), ("DIG", 2_550_000, 45), ("SERV", 640_000, 25), ("CORP", 980_000, 30)]:
    for v in repartir(total, n):
        b, pg = banco(), doc("PAG")
        lancar(dia(), [D("2.1.01", u, v), C(b, "CORP", v)], f"Pagamento a fornecedor {pg}", documento=pg,
               usuario="tesouraria.rui")
# P4: cheque emitido em 20/10 e não compensado (9 dias úteis em 31/10)
lancar("2025-10-20", [D("2.1.01", "SERV", 48_700.00), C("1.1.01.01", "CORP", 48_700.00)],
       "Pagamento a fornecedor: cheque 004512", documento="CHQ-004512", usuario="tesouraria.rui", marca="cheque")

# ---- depreciação, resultado financeiro e transitória
for cc, v in [("DIST-ADM", 182_400.00), ("DIG-TEC", 96_150.00), ("SERV-CDSE", 141_820.00), ("CORP-FAC", 61_300.00)]:
    lancar("2025-10-31", [D("6.1.01", cc, v), C("1.2.02", un(cc), v)], "Depreciação do mês")
for v, d_ in [(21_300.55, "2025-10-10"), (18_990.10, "2025-10-24")]:
    j = doc("ANT")
    lancar(d_, [D("7.1.03", "DIG-OPS", v), C("1.1.01.02", "CORP", v)], f"Juros de antecipação de recebíveis {j}",
           documento=j, usuario="tesouraria.rui")
lancar("2025-10-15", [D("7.1.02", "CORP-FIN", 182_000.00), C("1.1.01.01", "CORP", 182_000.00)],
       "Juros de empréstimo: parcela 14/36", documento=doc("EMP"), usuario="tesouraria.rui")
for b, v in [("1.1.01.01", 1_890.00), ("1.1.01.02", 1_240.00), ("1.1.01.03", 760.00)]:
    lancar("2025-10-01", [D("7.1.02", "CORP-FIN", v), C(b, "CORP", v)], "Tarifa de pacote de serviços bancários",
           documento=doc("TAR"), usuario="tesouraria.rui")
lancar("2025-10-31", [D("1.1.01.01", "CORP", 35_420.18), C("7.1.01", "CORP-FIN", 35_420.18)],
       "Rendimento de aplicação financeira", documento=doc("REND"), usuario="tesouraria.rui")
lancar("2025-10-14", [D("1.1.09", "CORP", 15_300.00), C("1.1.01.03", "CORP", 15_300.00)],
       "Pagamento a classificar: TED 7702", documento="TED-007702", usuario="tesouraria.rui")
lancar("2025-10-16", [D("5.3.04", "CORP-FIN", 15_300.00), C("1.1.09", "CORP", 15_300.00)],
       "Reclassificação da TED 7702: honorários advocatícios")
# P9: pagamento a classificar que ficou na transitória
lancar("2025-10-22", [D("1.1.09", "CORP", 128_450.00), C("1.1.01.02", "CORP", 128_450.00)],
       "Pagamento a classificar: TED 7781", documento="TED-007781", usuario="tesouraria.rui")

# ---- problemas plantados no razão
# P3: 3 lançamentos liberados no ERP depois da extração do razão (estão no balancete da extração 1)
for conta, cc, v, h in [("5.3.04", "DIST-ADM", 30_000.00, "Assessoria jurídica NF 88123"),
                        ("5.3.03", "CORP-TI", 21_800.00, "Licença de antivírus NF 4471"),
                        ("5.3.06", "SERV-CDNE", 12_500.00, "Reparo de doca NF 2290")]:
    lancar("2025-11-04", [D(conta, cc, v), C("2.1.01", un(cc), v)], h, origem="INTEGRACAO_COMPRAS",
           usuario="integracao", documento=doc("NFS"), marca="P3")
# P8: 3 lançamentos com competência nov/25 dentro de out/25
for conta, cc, v, h in [("5.1.01", "DIG-MKT", 120_000.00, "Campanha Black Friday (nov/25)"),
                        ("5.3.04", "CORP-FIN", 96_000.00, "Auditoria externa: honorários de nov/25"),
                        ("5.3.06", "DIST-ARM", 70_000.00, "Manutenção de empilhadeiras: contrato de nov/25")]:
    lancar(dia(), [D(conta, cc, v), C("2.1.01", un(cc), v)], h, comp="11/2025", usuario="contab.bia",
           documento=doc("NFS"), marca="P8")
# P8: 4 lançamentos de out/25 feitos depois da data de corte (D+3 úteis = 05/11)
for d_, conta, cc, v, h in [("2025-11-06", "5.3.05", "DIST-COM", 8_200.00, "Viagens comerciais: reembolso"),
                            ("2025-11-06", "5.3.06", "SERV-CDSE", 12_300.00, "Peças de reposição"),
                            ("2025-11-07", "5.3.04", "CORP-FIN", 9_500.00, "Consultoria de RH"),
                            ("2025-11-07", "5.1.04", "DIG-OPS", 11_500.00, "Frete de e-commerce: complemento")]:
    lancar(d_, [D(conta, cc, v), C("2.1.01", un(cc), v)], h, usuario="contab.bia", documento=doc("NFS"), marca="P8")
# armadilha: nota de set/25 que estava provisionada; a provisão é revertida no mesmo mês
lancar("2025-10-08", [D("5.3.04", "DIST-ADM", 45_000.00), C("2.1.01", "DIST", 45_000.00)],
       "Consultoria de processos: NF de set/25", comp="09/2025", documento=doc("NFS"))
lancar("2025-10-08", [D("2.1.08", "DIST", 45_000.00), C("5.3.04", "DIST-ADM", 45_000.00)],
       "Estorno da provisão de set/25: consultoria de processos")
# P2 (reextração): crédito de ICMS sobre imobilizado que tinha sido lançado em set/25 depois do fechamento
lancar("2025-11-03", [D("1.1.06", "CORP", 37_215.40), C("1.2.01", "CORP", 37_215.40)],
       "Reclassificação de ICMS sobre ativo imobilizado (CIAP) de set/25", usuario="contab.bia", marca="retro")

# ---- bancos: aplicação ou resgate para fechar o saldo desejado de cada banco
for b, (_, _, abertura, alvo) in BANCOS.items():
    mov = sum(l["debito"] - l["credito"] for l in LINHAS if l["conta"] == b)
    falta = round(alvo - abertura - mov, 2)
    if falta > 0:
        lancar("2025-10-30", [D(b, "CORP", falta), C("1.1.02", "CORP", falta)], "Resgate de aplicação financeira",
               documento=doc("APL"), usuario="tesouraria.rui")
    else:
        lancar("2025-10-30", [D("1.1.02", "CORP", -falta), C(b, "CORP", -falta)], "Aplicação financeira",
               documento=doc("APL"), usuario="tesouraria.rui")

# ---- calibração: EBITDA do mês, antes dos ajustes, de exatamente R$ 1,42 mi (via CMV do último dia da DIST)
razao = pd.DataFrame(LINHAS)
razao["mov"] = razao["debito"] - razao["credito"]
ebitda = -razao.loc[razao["conta"].map(GRUPO).isin(GRUPOS_EBITDA), "mov"].sum()
plug = round(float(ebitda - EBITDA_ALVO), 2)
idx = razao.index[razao["marca"] == "cmv_ultimo"]
assert plug > -0.5 * razao.loc[idx[0], "debito"], plug
for i in idx:
    for lado in ("debito", "credito"):
        if LINHAS[i][lado]:
            LINHAS[i][lado] = round(LINHAS[i][lado] + plug, 2)
razao = pd.DataFrame(LINHAS)
razao["mov"] = razao["debito"] - razao["credito"]
ebitda = round(float(-razao.loc[razao["conta"].map(GRUPO).isin(GRUPOS_EBITDA), "mov"].sum()), 2)
assert ebitda == EBITDA_ALVO, ebitda

# --------------------------------------------------------------------------- subledgers
# aging de contas a receber: faixas exatas (P11) e 2 títulos cancelados sem estorno contábil (P6)
FAIXAS_CR = [(0, 19_800_000, 700, (-60, 0), 2), (1, 3_100_000, 260, (1, 30), 1), (2, 1_480_000, 140, (31, 90), 1),
             (3, 608_000, 60, (91, 180), 1), (4, 412_000, 40, (181, 420), 1)]
cr = []
n_tit = 0
for faixa, total, n, (d0, d1), passo in FAIXAS_CR:
    for v in repartir(total, n, sigma=0.9, passo=passo):
        n_tit += 1
        u = str(rng.choice(["DIST", "DIG", "SERV"], p=[0.62, 0.13, 0.25]))
        atraso = int(rng.integers(d0, d1 + 1))
        venc = DATA_BASE - pd.Timedelta(days=atraso)
        emis = min(venc - pd.Timedelta(days=int(rng.choice([28, 30, 45, 60]))),
                   DATA_BASE - pd.Timedelta(days=int(rng.integers(0, 25))))
        orig = v if rng.random() < 0.85 else round(v * float(rng.uniform(1.1, 2.0)), 2)
        cr.append(dict(titulo=f"CR-{n_tit:06d}", cliente=f"Cliente {int(rng.integers(1, 420)):04d}", unidade=u,
                       emissao=emis, vencimento=venc, valor_original=orig, valor_aberto=v, situacao="aberto",
                       faixa=faixa))
for v in (98_000.00, 58_000.00):
    n_tit += 1
    cr.append(dict(titulo=f"CR-{n_tit:06d}", cliente=f"Cliente {int(rng.integers(1, 420)):04d}", unidade="DIST",
                   emissao=pd.Timestamp("2025-09-12"), vencimento=pd.Timestamp("2025-10-27"), valor_original=v,
                   valor_aberto=0.0, situacao="cancelado", faixa=-1))
cr = pd.DataFrame(cr)
aging_cr_un = cr.groupby("unidade")["valor_aberto"].sum().round(2)
assert round(aging_cr_un.sum(), 2) == 25_400_000.00
taxas = [f[3] for f in PDD_FAIXAS]
cr_ab = cr[cr["faixa"] >= 0]
pdd_calc = (cr_ab["valor_aberto"] * cr_ab["faixa"].map(dict(enumerate(taxas)))).groupby(cr_ab["unidade"]).sum().round(2)
assert round(pdd_calc.sum(), 2) == 1_025_000.00, pdd_calc.sum()
ajuste_pdd = (pdd_calc / pdd_calc.sum() * 155_000).round(0)
ajuste_pdd["DIST"] += 155_000 - ajuste_pdd.sum()
pdd_contabil = (pdd_calc - ajuste_pdd).round(2)
assert round(pdd_contabil.sum(), 2) == 870_000.00

# aging de contas a pagar: DIG com adiantamento lançado em fornecedores (P9: saldo devedor de R$ 73 mil)
cp = []
for u, total, n in [("DIST", 6_840_000, 330), ("DIG", 412_000, 60), ("SERV", 1_126_000, 150), ("CORP", 1_574_000, 159)]:
    for v in repartir(total, n, sigma=0.8):
        venc = DATA_BASE + pd.Timedelta(days=int(rng.integers(-25, 61)))
        cp.append(dict(titulo=f"CP-{len(cp) + 1:06d}", fornecedor=f"Fornecedor {int(rng.integers(1, 260)):03d}",
                       unidade=u, emissao=min(venc - pd.Timedelta(days=int(rng.choice([30, 45, 60]))),
                                              DATA_BASE - pd.Timedelta(days=int(rng.integers(0, 25)))),
                       vencimento=venc,
                       valor_aberto=v, conta_contabil="2.1.01"))
cp.append(dict(titulo=f"CP-{len(cp) + 1:06d}", fornecedor="Fornecedor 118", unidade="DIG",
               emissao=pd.Timestamp("2025-10-09"), vencimento=pd.Timestamp("2025-10-09"), valor_aberto=-485_000.00,
               conta_contabil="2.1.01"))
cp = pd.DataFrame(cp)
aging_cp_un = cp.groupby("unidade")["valor_aberto"].sum().round(2)
assert aging_cp_un["DIG"] == -73_000.00

# provisões de férias e 13º: calculado (regra do A2) − dissídio de 6% da SERV não refletido (P12)
def meses_13(adm: pd.Timestamp) -> int:
    if adm < pd.Timestamp("2025-01-01"):
        return 10
    return sum(1 for m in range(adm.month, 11)
               if (pd.Timestamp(2025, m, 1) + pd.offsets.MonthEnd(0) - max(adm, pd.Timestamp(2025, m, 1))).days + 1 >= 15)


prov = {}
for p in pessoas:
    if pd.notna(p["data_desligamento"]):
        continue
    u = un(p["centro_custo"])
    fer = round(p["salario"] / 30 * p["dias_ferias_adquiridos"] * 4 / 3 * (1 + ENC), 2)
    dec = round(p["salario"] * meses_13(p["data_admissao"]) / 12 * (1 + ENC), 2)
    a = prov.setdefault(u, [0.0, 0.0])
    a[0] += fer
    a[1] += dec
prov = {u: (round(f, 2), round(d, 2)) for u, (f, d) in prov.items()}
AJ_FER, AJ_DEC = {"SERV": 34_100.00}, {"SERV": 24_200.00}

# --------------------------------------------------------------------------- saldos
mov = razao.groupby(["conta", "unidade"])[["debito", "credito"]].sum().round(2)
mov["mov"] = (mov["debito"] - mov["credito"]).round(2)
pares = set(mov.index)


def m(conta, u):
    return float(mov.loc[(conta, u), "mov"]) if (conta, u) in pares else 0.0


alvo = {  # saldo final desejado (positivo = devedor) das contas patrimoniais
    ("1.1.02", "CORP"): 8_210_000.00, ("1.1.05", "DIST"): 14_300_000.00, ("1.1.05", "DIG"): 4_100_000.00,
    ("1.1.06", "CORP"): 3_152_640.75, ("1.1.07", "DIST"): 340_000.00, ("1.1.07", "CORP"): 262_500.00,
    ("1.1.09", "CORP"): 128_450.00,
    ("1.2.01", "DIST"): 18_400_000.00, ("1.2.01", "DIG"): 6_100_000.00, ("1.2.01", "SERV"): 13_800_000.00,
    ("1.2.01", "CORP"): 4_200_000.00, ("1.2.02", "DIST"): -7_050_000.00, ("1.2.02", "DIG"): -2_140_000.00,
    ("1.2.02", "SERV"): -5_260_000.00, ("1.2.02", "CORP"): -1_310_000.00,
    ("2.1.08", "DIST"): 0.0, ("2.1.08", "SERV"): -36_500.00, ("2.1.08", "CORP"): -214_000.00,
    ("2.2.01", "CORP"): -15_000_000.00,
}
for u in ["DIST", "DIG", "SERV"]:
    alvo[("1.1.03", u)] = float(aging_cr_un[u]) + (156_000.00 if u == "DIST" else 0.0)    # P6
    alvo[("1.1.04", u)] = -float(pdd_contabil[u])
    alvo[("2.1.04", u)] = -float(mov.loc[("2.1.04", u), "credito"])
for u in ["DIST", "DIG", "SERV", "CORP"]:
    alvo[("2.1.01", u)] = -float(aging_cp_un[u])
    alvo[("2.1.02", u)] = -float(mov.loc[("2.1.02", u), "credito"])
    alvo[("2.1.03", u)] = -float(mov.loc[("2.1.03", u), "credito"])
    alvo[("2.1.05", u)] = -round(prov[u][0] - AJ_FER.get(u, 0.0), 2)
    alvo[("2.1.06", u)] = -round(prov[u][1] - AJ_DEC.get(u, 0.0), 2)

# movimento de set/25 (contas de resultado): out/25 "limpo" × ruído, com as exceções plantadas
limpo = razao[~razao["marca"].isin(["P3", "P8", "retro"])]
limpo = limpo[limpo["conta"].map(TIPO) == "resultado"].groupby(["conta", "unidade"])["mov"].sum()
setembro = {k: round(float(v * rng.uniform(0.93, 1.07)), 2) for k, v in limpo.items()}
for (conta, cc) in FALTANTES:
    ref = next(r[3] for r in RECORRENTES if (r[0], r[1]) == (conta, cc))
    setembro[(conta, un(cc))] = round((float(limpo.get((conta, un(cc)), 0.0)) + ref) * float(rng.uniform(0.97, 1.03)), 2)
setembro[("5.1.04", "DIST")] = 690_000.00           # C8: frete da Distribuição sobe 48%
setembro[("5.3.05", "CORP")] = 230_000.00           # armadilha: +35%, mas só R$ 80 mil
setembro.update({("5.1.05", "DIST"): 38_000.00, ("5.1.05", "DIG"): 11_500.00, ("5.1.05", "SERV"): 7_800.00})
jan_ago = {k: round(v * float(rng.uniform(7.7, 8.3)), 2) for k, v in setembro.items()}

abertura = {}
for (conta, u) in sorted(pares | set(alvo) | set(setembro)):
    tipo = TIPO[conta]
    if tipo == "resultado":
        abertura[(conta, u)] = round(jan_ago.get((conta, u), 0.0) + setembro.get((conta, u), 0.0), 2)
    elif conta in BANCOS:
        abertura[(conta, u)] = BANCOS[conta][2]
    else:
        assert (conta, u) in alvo, (conta, u)
        abertura[(conta, u)] = round(alvo[(conta, u)] - m(conta, u), 2)
abertura[("2.3.01", "CORP")] = -30_000_000.00
abertura[("2.3.02", "CORP")] = round(-sum(abertura.values()), 2)
assert abertura[("2.3.02", "CORP")] < 0, abertura[("2.3.02", "CORP")]

# balancete da reextração (extração 2)
linhas_bal = []
for (conta, u), ab in sorted(abertura.items()):
    d_, c_ = (float(mov.loc[(conta, u), "debito"]), float(mov.loc[(conta, u), "credito"])) if (conta, u) in pares else (0.0, 0.0)
    linhas_bal.append(dict(conta=conta, unidade=u, saldo_inicial=ab, debitos=round(d_, 2), creditos=round(c_, 2),
                           saldo_final=round(ab + d_ - c_, 2)))
bal2 = pd.DataFrame(linhas_bal)
desc = {c[0]: c[1] for c in CONTAS}
bal2.insert(1, "descricao", bal2["conta"].map(desc))
assert abs(round(bal2["saldo_final"].sum(), 2)) < 0.01 and abs(round(bal2["saldo_inicial"].sum(), 2)) < 0.01

# conferência interna do C8 (só o frete da Distribuição pode disparar)
for (conta, u), s in setembro.items():
    o = m(conta, u)
    delta = o - s
    dispara = abs(delta) > 200_000 and (s == 0 or abs(delta) / abs(s) > 0.30)
    assert dispara == ((conta, u) == ("5.1.04", "DIST")), (conta, u, s, o)
# conferência interna do C7: nenhuma natureza invertida além das plantadas
nat = {c[0]: c[3] for c in CONTAS}
for r in bal2.itertuples():
    if TIPO[r.conta] != "patrimonial":
        continue
    inv = (nat[r.conta] == "D" and r.saldo_final < -1000) or (nat[r.conta] == "C" and r.saldo_final > 1000)
    assert inv == ((r.conta, r.unidade) == ("2.1.01", "DIG")), (r.conta, r.unidade, r.saldo_final)

# extração 1: balancete com 2 erros (P1 e P2) e razão sem os 3 lançamentos liberados depois (P3)
bal1 = bal2.copy()
i6 = bal1.index[(bal1["conta"] == "1.1.06") & (bal1["unidade"] == "CORP")][0]
i12 = bal1.index[(bal1["conta"] == "1.2.01") & (bal1["unidade"] == "CORP")][0]
bal1.loc[i6, ["saldo_inicial", "debitos"]] += [37_215.40, -37_215.40]
bal1.loc[i12, ["saldo_inicial", "creditos"]] += [-37_215.40, -37_215.40]
i3 = bal1.index[(bal1["conta"] == "2.1.03") & (bal1["unidade"] == "SERV")][0]
bal1.loc[i3, "saldo_final"] += 12_480.00
bal1[["saldo_inicial", "debitos", "creditos", "saldo_final"]] = bal1[["saldo_inicial", "debitos", "creditos", "saldo_final"]].round(2)

COLS_RAZAO = ["id_lanc", "data_lanc", "data_competencia", "conta", "unidade", "centro_custo", "historico", "documento",
              "debito", "credito", "usuario", "origem"]
razao2 = razao.sort_values(["data_lanc", "id_lanc"], kind="stable")
razao1 = razao2[~razao2["marca"].isin(["P3", "retro"])]

# --------------------------------------------------------------------------- extrato bancário
DESCR = {"COB": "CRED COBRANCA BOLETO", "PAG": "PAGTO FORNECEDOR TED", "FOL": "PAGTO FOLHA SALARIOS",
         "GPS": "PAGTO GPS FGTS", "DARF": "PAGTO DARF TRIBUTOS", "ANT": "JUROS ANTECIPACAO RECEBIVEIS",
         "EMP": "DEB PARCELA EMPRESTIMO", "TAR": "TARIFA PACOTE SERVICOS", "REND": "RENDIMENTO APLIC FINANCEIRA",
         "TED": "TED ENVIADA", "APL": "APLIC/RESG AUTOMATICA"}
mov_banco = razao2[razao2["conta"].isin(BANCOS) & ~razao2["marca"].isin(["transito", "cheque"])]
ext = pd.DataFrame(dict(conta=mov_banco["conta"], data=mov_banco["data_lanc"], documento=mov_banco["documento"],
                        descricao=mov_banco["documento"].str.split("-").str[0].map(DESCR),
                        valor=(mov_banco["debito"] - mov_banco["credito"]).round(2)))
nao_lanc = [  # P4 e P5: itens do banco sem lançamento contábil
    ("1.1.01.01", "2025-10-03", "TAR-B-1101", "TARIFA TED/DOC", -1_245.30),
    ("1.1.01.01", "2025-10-15", "IOF-B-1102", "IOF SOBRE SALDO DEVEDOR", -890.00),
    ("1.1.01.02", "2025-10-08", "TAR-B-1103", "TARIFA EMISSAO BOLETOS", -612.45),
    ("1.1.01.02", "2025-10-21", "TAR-B-1104", "TARIFA TED/DOC", -430.15),
    ("1.1.01.03", "2025-10-06", "TAR-B-1105", "TARIFA MANUTENCAO CONTA", -385.00),
    ("1.1.01.03", "2025-10-17", "IOF-B-1106", "IOF OPERACAO DE CAMBIO", -402.10),
    ("1.1.01.01", "2025-10-28", "TAR-B-1107", "TARIFA CUSTODIA DE CHEQUES", -321.90),
    ("1.1.01.02", "2025-10-31", "REN-B-1108", "RENDIMENTO APLIC AUTOMATICA", 18_640.22),
    ("1.1.01.02", "2025-10-22", "TED-B-1109", "TED RECEBIDA 341 0457 REMETENTE NAO INFORMADO", 92_000.00),
]
ext = pd.concat([ext, pd.DataFrame(nao_lanc, columns=["conta", "data", "documento", "descricao", "valor"])
                 .assign(data=lambda x: pd.to_datetime(x["data"]))], ignore_index=True)
ext = ext.sort_values(["conta", "data", "documento"], kind="stable")
ext["banco"] = ext["conta"].map(lambda c: BANCOS[c][0])
ext["agencia_conta"] = ext["conta"].map(lambda c: BANCOS[c][1])
ext["saldo"] = (ext.groupby("conta")["valor"].cumsum() + ext["conta"].map(lambda c: BANCOS[c][2])).round(2)
extrato = ext[["banco", "agencia_conta", "data", "documento", "descricao", "valor", "saldo"]]

# --------------------------------------------------------------------------- folha (arquivo)
folha = []
for p in pessoas:
    base = dict(matricula=p["matricula"], centro_custo=p["centro_custo"], data_admissao=p["data_admissao"],
                data_desligamento=p["data_desligamento"], dias_ferias_adquiridos=p["dias_ferias_adquiridos"])
    for rub, chave in [("SALARIO", "salario"), ("HORA_EXTRA", "he"), ("ENCARGOS", "encargos"), ("BENEFICIOS", "beneficios")]:
        if p[chave]:
            folha.append(base | dict(rubrica=rub, valor=p[chave]))
folha = pd.DataFrame(folha)[["matricula", "centro_custo", "rubrica", "valor", "data_admissao", "data_desligamento",
                             "dias_ferias_adquiridos"]]

# --------------------------------------------------------------------------- pacote de set/25
pac = []
for (conta, u), final in sorted(abertura.items()):
    if TIPO[conta] == "resultado":
        s = setembro.get((conta, u), 0.0)
        d_, c_ = (s, 0.0) if s >= 0 else (0.0, -s)
    else:
        d_ = round(abs(final) * float(rng.uniform(0.05, 0.4)), 2)
        c_ = round(d_ * float(rng.uniform(0.8, 1.2)), 2)
    pac.append(dict(conta=conta, descricao=desc[conta], unidade=u, saldo_inicial=round(final - d_ + c_, 2),
                    debitos=d_, creditos=c_, saldo_final=final))
pacote_set = pd.DataFrame(pac)

# --------------------------------------------------------------------------- gravação
def gravar(df: pd.DataFrame, caminho: Path) -> None:
    caminho.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(caminho, sep=";", decimal=",", float_format="%.2f", date_format="%d/%m/%Y", index=False,
              encoding="utf-8-sig")


gravar(pacote_set, ENTRADAS / "2025-09" / "balancete_ajustado_2025-09.csv")
for n_ext, bal, raz in [(1, bal1, razao1), (2, bal2, razao2)]:
    pasta = ENTRADAS / MES / f"extracao_{n_ext}"
    gravar(bal, pasta / f"balancete_{MES}.csv")
    gravar(raz[COLS_RAZAO], pasta / f"razao_{MES}.csv")
    gravar(extrato, pasta / f"extrato_bancario_{MES}.csv")
    gravar(cr.drop(columns="faixa"), pasta / f"aging_receber_{MES}.csv")
    gravar(cp, pasta / f"aging_pagar_{MES}.csv")
    gravar(folha, pasta / f"folha_{MES}.csv")

with pd.ExcelWriter(ENTRADAS / "parametros.xlsx", engine="xlsxwriter") as xw:
    xw.book.set_properties({"created": datetime(2025, 11, 5)})     # data fixa: o arquivo sai idêntico a cada execução
    pd.DataFrame([("mes", MES), ("data_base", "31/10/2025"), ("data_corte", "05/11/2025"),
                  ("taxa_encargos", ENC), ("unidade_tesouraria", "CORP"), ("preparado_por", "Controladoria")],
                 columns=["parametro", "valor"]).to_excel(xw, sheet_name="geral", index=False)
    pd.DataFrame(UNIDADES, columns=["unidade", "nome", "cc_receita", "cc_pessoal", "conta_pessoal"]).to_excel(
        xw, sheet_name="unidades", index=False)
    pd.DataFrame(CONTAS, columns=["conta", "descricao", "grupo", "natureza", "tipo"]).to_excel(
        xw, sheet_name="contas", index=False)
    pd.DataFrame(PDD_FAIXAS, columns=["faixa", "dias_de", "dias_ate", "percentual"]).to_excel(
        xw, sheet_name="pdd", index=False)
    pd.DataFrame(RECORRENTES, columns=["conta", "centro_custo", "descricao", "valor_referencia"]).to_excel(
        xw, sheet_name="recorrentes", index=False)
    pd.DataFrame([(b, a, c) for c, (b, a, _, _) in BANCOS.items()], columns=["banco", "agencia_conta", "conta_contabil"]
                 ).to_excel(xw, sheet_name="bancos", index=False)
    pd.DataFrame([("C1", 0.00, "Contabilidade", 1), ("C2", 1.00, "Contabilidade", 1), ("C3", 0.00, "Tesouraria", 2),
                  ("C4", 100.00, "Contas a receber e a pagar", 3), ("C5", 500.00, "Departamento pessoal", 3),
                  ("C6", 50_000.00, "Contabilidade", 2), ("C7", 1_000.00, "Contabilidade", 3),
                  ("C8", 200_000.00, "FP&A com o gestor do centro de custo", 5)],
                 columns=["conferencia", "tolerancia", "responsavel", "prazo_dias_uteis"]).to_excel(
        xw, sheet_name="tolerancias", index=False)
    pd.DataFrame([("C3", "dias_uteis_em_aberto", 5), ("C5", "contas_pessoal", "4.1.02;5.2.01;5.2.02;5.2.03"),
                  ("C5", "origem_folha", "INTEGRACAO_FOLHA"), ("C7", "percentual_minimo_recorrente", 0.5),
                  ("C8", "variacao_percentual", 0.30)],
                 columns=["conferencia", "parametro", "valor"]).to_excel(xw, sheet_name="regras", index=False)

gabarito = {
    "mes": MES,
    "extracao_1": {
        "veredito": "não apresentável",
        "C1": {"status": "Bloqueia",
               "erro_aritmetico": [{"conta": "2.1.03", "unidade": "SERV", "diferenca": 12_480.00}],
               "saldo_inicial": [{"conta": "1.1.06", "unidade": "CORP", "diferenca": 37_215.40},
                                 {"conta": "1.2.01", "unidade": "CORP", "diferenca": -37_215.40}]},
        "C2": {"status": "Bloqueia", "linhas": 6, "debitos_faltantes": 64_300.00, "creditos_faltantes": 64_300.00},
    },
    "extracao_2": {
        "veredito": "com ressalvas",
        "ebitda_antes": EBITDA_ALVO, "ebitda_depois": round(EBITDA_ALVO - 155_000 - 58_300 - 247_000, 2),
        "C1": {"status": "OK"}, "C2": {"status": "OK"},
        "C3": {"status": "Pendência", "pendencia": 140_700.00, "cheque_antigo": 48_700.00, "nao_identificado": 92_000.00,
               "transito_no_prazo": 215_000.00, "tarifas": -4_286.90, "rendimentos": 18_640.22,
               "diferenca_nao_explicada": 0.0},
        "C4": {"status": "Pendência", "diferenca": 156_000.00, "unidades": ["DIST"],
               "razao_cr": 25_556_000.00, "aging_cr": 25_400_000.00},
        "C5": {"status": "Pendência", "diferenca": 22_400.00, "centros": ["SERV-CDSE"], "folha": 412_300.00,
               "razao": 389_900.00},
        "C6": {"status": "Pendência", "lancamentos": 7, "valor": 327_500.00, "competencia": 286_000.00,
               "corte": 41_500.00},
        "C7": {"status": "Pendência", "valor": 201_450.00,
               "itens": [{"conta": "2.1.01", "unidade": "DIG", "saldo": 73_000.00},
                         {"conta": "1.1.09", "unidade": "CORP", "saldo": 128_450.00}],
               "recorrentes_faltantes": 247_000.00},
        "C8": {"status": "Alerta", "itens": [{"conta": "5.1.04", "unidade": "DIST", "setembro": 690_000.00,
                                              "outubro": 1_020_000.00}]},
        "A1": 155_000.00, "A2": 58_300.00, "A3": 247_000.00, "A4": {"tarifas": 4_286.90, "rendimentos": 18_640.22},
        "pdd_calculada": 1_025_000.00, "pdd_contabil": 870_000.00,
        "pendencias_total": 520_550.00,
    },
}
(RAIZ / "gerador" / "gabarito.json").write_text(json.dumps(gabarito, ensure_ascii=False, indent=2), encoding="utf-8")

rl = -sum(m(c, u) for c, u in pares if GRUPO[c] in ("Receita Bruta", "Deduções"))
cmv = sum(m("4.1.01", u) for u in ("DIST", "DIG"))
print(f"Razão: {len(razao2)} linhas ({razao2['id_lanc'].nunique()} lançamentos) | extrato: {len(extrato)} | "
      f"aging CR: {len(cr)} | aging CP: {len(cp)} | folha: {len(folha)} linhas, {len(pessoas)} pessoas")
print(f"Receita líquida out/25: {rl:,.2f} | CMV: {cmv:,.2f} | EBITDA: {ebitda:,.2f} | ajuste no CMV: {plug:,.2f}")
print(f"Lucros acumulados (set/25): {abertura[('2.3.02', 'CORP')]:,.2f}")
