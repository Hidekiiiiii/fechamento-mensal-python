"""Ajustes de fechamento propostos pela automação.

Cada ajuste sai como lançamento em partida dupla, com memória de cálculo e status "proposto".
Nada é gravado no ERP: o controller aprova (S/N) na aba Ajustes do pacote.
"""
from __future__ import annotations

from dataclasses import dataclass, field

import pandas as pd

from . import plano
from .conferencias import Conferencia
from .formatacao import brl
from .leitura import Entradas

COLUNAS = ["ajuste", "conta", "unidade", "centro_custo", "historico", "debito", "credito"]


@dataclass
class Ajuste:
    codigo: str
    nome: str
    regra: str
    linhas: pd.DataFrame
    memoria: pd.DataFrame = field(default_factory=pd.DataFrame)

    @property
    def valor(self) -> float:
        return round(float(self.linhas["debito"].sum()), 2)


def _linhas(codigo: str, itens: list[tuple]) -> pd.DataFrame:
    return pd.DataFrame([(codigo, *i) for i in itens], columns=COLUNAS)


def _par(conta_d, un_d, cc_d, conta_c, un_c, cc_c, historico, valor):
    """Dois lados de um lançamento; valor negativo inverte débito e crédito."""
    if valor < 0:
        conta_d, un_d, cc_d, conta_c, un_c, cc_c, valor = conta_c, un_c, cc_c, conta_d, un_d, cc_d, -valor
    v = round(valor, 2)
    return [(conta_d, un_d, cc_d, historico, v, 0.0), (conta_c, un_c, cc_c, historico, 0.0, v)]


def _saldo(e: Entradas, conta: str) -> pd.Series:
    b = e.balancete
    return b[b["conta"] == conta].groupby("unidade")["saldo_final"].sum()


# --------------------------------------------------------------------------- A1
def a1_pdd(e: Entradas) -> Ajuste:
    p = e.p
    cr = e.aging_cr[e.aging_cr["valor_aberto"] != 0].copy()
    cr["dias"] = (p.data_base - cr["vencimento"]).dt.days
    faixas = p.pdd.sort_values("dias_de")
    cr["faixa"] = pd.cut(cr["dias"], bins=list(faixas["dias_de"] - 0.5) + [faixas["dias_ate"].iloc[-1] + 0.5],
                         labels=list(faixas["faixa"]))
    cr["percentual"] = cr["faixa"].map(faixas.set_index("faixa")["percentual"]).astype(float)
    cr["pdd"] = cr["valor_aberto"] * cr["percentual"]
    mem = cr.groupby(["unidade", "faixa"], observed=False).agg(saldo=("valor_aberto", "sum"), pdd=("pdd", "sum"))
    mem = mem.reset_index()
    mem["percentual"] = mem["faixa"].map(faixas.set_index("faixa")["percentual"]).astype(float)
    calc = cr.groupby("unidade")["pdd"].sum().round(2)
    contabil = (-_saldo(e, plano.PDD)).reindex(calc.index).fillna(0).round(2)
    resumo = pd.DataFrame({"pdd_calculada": calc, "pdd_contabil": contabil})
    resumo["ajuste"] = (resumo["pdd_calculada"] - resumo["pdd_contabil"]).round(2)
    itens = []
    for u, r in resumo.iterrows():
        if abs(r.ajuste) >= 0.01:
            cc = p.unidades.loc[u, "cc_receita"]
            itens += _par(plano.DESP_PDD, u, cc, plano.PDD, u, "", f"Complemento da PDD {u}: calculada "
                          f"{brl(r.pdd_calculada, 2)} × contábil {brl(r.pdd_contabil, 2)}", r.ajuste)
    aj = Ajuste("A1", "Complemento da PDD", "Aging por faixa de atraso × percentual da política; ajuste = PDD "
                "calculada − saldo contábil da PDD", _linhas("A1", itens), mem)
    aj.resumo = resumo.rename_axis("unidade").reset_index()
    return aj


# --------------------------------------------------------------------------- A2
def meses_13(admissao: pd.Timestamp, data_base: pd.Timestamp) -> int:
    """Avos de 13º no ano até a data-base: conta o mês com 15 dias ou mais trabalhados."""
    inicio_ano = pd.Timestamp(data_base.year, 1, 1)
    if admissao < inicio_ano:
        return data_base.month
    n = 0
    for m in range(admissao.month, data_base.month + 1):
        ini = pd.Timestamp(data_base.year, m, 1)
        fim = ini + pd.offsets.MonthEnd(0)
        if (fim - max(admissao, ini)).days + 1 >= 15:
            n += 1
    return n


def a2_ferias_13(e: Entradas) -> Ajuste:
    p = e.p
    enc = p.encargos
    f = e.folha
    pessoas = f.groupby("matricula").agg(centro_custo=("centro_custo", "first"), admissao=("data_admissao", "first"),
                                         desligamento=("data_desligamento", "first"),
                                         dias=("dias_ferias_adquiridos", "first"))
    pessoas["salario"] = f[f["rubrica"] == "SALARIO"].groupby("matricula")["valor"].sum()
    ativos = pessoas[pessoas["desligamento"].isna() | (pessoas["desligamento"] > p.data_base)].copy()
    ativos["unidade"] = ativos["centro_custo"].str.split("-").str[0]
    ativos["ferias"] = (ativos["salario"] / 30 * ativos["dias"] * 4 / 3 * (1 + enc)).round(2)
    ativos["avos_13"] = [meses_13(a, p.data_base) for a in ativos["admissao"]]
    ativos["decimo"] = (ativos["salario"] * ativos["avos_13"] / 12 * (1 + enc)).round(2)
    calc = ativos.groupby("unidade")[["ferias", "decimo"]].sum().round(2)
    calc["ferias_contabil"] = (-_saldo(e, plano.PROV_FERIAS)).reindex(calc.index).fillna(0).round(2)
    calc["decimo_contabil"] = (-_saldo(e, plano.PROV_13)).reindex(calc.index).fillna(0).round(2)
    calc["ajuste_ferias"] = (calc["ferias"] - calc["ferias_contabil"]).round(2)
    calc["ajuste_decimo"] = (calc["decimo"] - calc["decimo_contabil"]).round(2)
    calc["ajuste"] = (calc["ajuste_ferias"] + calc["ajuste_decimo"]).round(2)
    calc["ajuste_lancado"] = calc["ajuste"].where(calc["ajuste"].abs() >= 1.00, 0.0)
    calc["observacao"] = calc["ajuste"].map(lambda v: "" if abs(v) >= 1.00 else
                                            ("Sem diferença" if v == 0 else "Abaixo de R$ 1,00 (arredondamento): "
                                             "sem lançamento"))
    itens = []
    for u, r in calc.iterrows():
        if r.ajuste_lancado == 0:      # diferença de arredondamento não vira lançamento
            continue
        conta, cc = p.unidades.loc[u, "conta_pessoal"], p.unidades.loc[u, "cc_pessoal"]
        h = f"Complemento das provisões de férias e 13º {u} (salário vigente em {p.data_base:%m/%Y})"
        itens += _par(conta, u, cc, plano.PROV_FERIAS, u, "", h, r.ajuste_ferias)
        itens += _par(conta, u, cc, plano.PROV_13, u, "", h, r.ajuste_decimo)
    aj = Ajuste("A2", "Complemento de férias e 13º", f"Férias = salário/30 × dias adquiridos × 4/3 × (1 + {enc:.0%}); "
                f"13º = salário × avos/12 × (1 + {enc:.0%}); ajuste = calculado − saldo contábil",
                _linhas("A2", itens), calc.rename_axis("unidade").reset_index())
    return aj


# --------------------------------------------------------------------------- A3
def a3_sem_nota(e: Entradas, c7: Conferencia) -> Ajuste:
    itens = []
    falt = c7.extra.get("faltantes", pd.DataFrame())
    for r in falt.itertuples():
        u = r.centro_custo.split("-")[0]
        itens += _par(r.conta, u, r.centro_custo, plano.PROV_PAGAR, u, "", f"Provisão: {r.descricao} "
                      f"(sem nota no mês; estorno automático no mês seguinte)", r.valor_referencia)
    return Ajuste("A3", "Despesa incorrida sem nota", "Despesa recorrente sem lançamento no mês: provisiona o valor de "
                  "referência contra provisões a pagar, com estorno no mês seguinte", _linhas("A3", itens),
                  falt.drop(columns="Index", errors="ignore"))


# --------------------------------------------------------------------------- A4
def a4_banco(e: Entradas, c3: Conferencia) -> Ajuste:
    p = e.p
    conta_banco = p.bancos.set_index("banco")["conta_contabil"]
    un = str(p.geral["unidade_tesouraria"])
    cc = p.unidades.loc[un, "cc_receita"]
    it = c3.detalhe
    it = it[it["situacao"] == "Ajuste A4 proposto"]
    itens = []
    for r in it.itertuples():
        h = f"{r.descricao} ({r.data:%d/%m}, {r.banco})"
        if r.valor < 0:
            itens += _par(plano.DESP_FIN, un, cc, conta_banco[r.banco], un, "", h, -r.valor)
        else:
            itens += _par(conta_banco[r.banco], un, "", plano.REC_FIN, un, cc, h, r.valor)
    return Ajuste("A4", "Itens do extrato não contabilizados", "Tarifas e IOF para despesa financeira; rendimentos para "
                  "receita financeira; contra a conta do banco, na data do extrato", _linhas("A4", itens),
                  it.drop(columns=["origem", "dias_uteis"], errors="ignore"))


def propor(e: Entradas, conf: dict[str, Conferencia]) -> dict[str, Ajuste]:
    return {a.codigo: a for a in (a1_pdd(e), a2_ferias_13(e), a3_sem_nota(e, conf["C7"]), a4_banco(e, conf["C3"]))}


def efeitos(e: Entradas, ajustes: dict[str, Ajuste]) -> pd.DataFrame:
    """Efeito de cada ajuste no EBITDA e abaixo dele (positivo = melhora o resultado)."""
    grupo = e.p.grupo()
    linhas = []
    for a in ajustes.values():
        l = a.linhas.assign(grupo=a.linhas["conta"].map(grupo))
        res = l[l["conta"].map(e.p.tipo()) == "resultado"]
        mov = -(res["debito"] - res["credito"])
        no_ebitda = mov[res["grupo"].isin(plano.GRUPOS_EBITDA)].sum()
        linhas.append(dict(ajuste=a.codigo, nome=a.nome, regra=a.regra, valor=a.valor,
                           efeito_ebitda=round(float(no_ebitda), 2),
                           efeito_abaixo_ebitda=round(float(mov.sum() - no_ebitda), 2), aprovado="S"))
    return pd.DataFrame(linhas)
