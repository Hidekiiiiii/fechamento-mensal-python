"""Efeito potencial no EBITDA das pendências já medidas que não viraram ajuste proposto.

O EBITDA "após ajustes propostos" não é o número final: algumas pendências mexem no resultado assim que forem
resolvidas. Aqui elas são quantificadas por regra, com a ação esperada, para o diretor ver a faixa provável.
"""
from __future__ import annotations

import pandas as pd

from . import plano


def efeito_potencial(conf: dict, grupo: pd.Series) -> pd.DataFrame:
    linhas = []
    c6 = conf["C6"].detalhe
    if not c6.empty:
        fut = c6[c6["futura"] & c6["conta"].map(grupo).isin(plano.GRUPOS_EBITDA)]
        if len(fut):
            linhas.append(dict(conferencia="C6", pendencia=f"{len(fut)} despesas de competência futura lançadas no mês",
                               efeito=round(float(fut["valor"].sum()), 2),
                               acao="Estornar no mês e relançar na competência"))
    c5 = conf["C5"].detalhe
    if not c5.empty:
        fora = c5[c5["status"] == "Pendência"]
        if len(fora):
            linhas.append(dict(conferencia="C5", pendencia="Folha maior que o razão em " + ", ".join(fora["centro_custo"]),
                               efeito=round(float(-fora["diferenca"].sum()), 2),
                               acao="Lançar o complemento da folha"))
    c4 = conf["C4"].detalhe
    if not c4.empty:
        cr = c4[(c4["status"] == "Pendência") & (c4["conta"] == plano.CLIENTES) & (c4["diferenca"] > 0)]
        if len(cr):
            linhas.append(dict(conferencia="C4", pendencia="Clientes no razão acima do aging em " + ", ".join(cr["unidade"]),
                               efeito=round(float(-cr["diferenca"].sum()), 2),
                               acao="Se o cancelamento das vendas se confirmar (contrapartida a confirmar)"))
    return pd.DataFrame(linhas, columns=["conferencia", "pendencia", "efeito", "acao"]).sort_values(
        "conferencia", ignore_index=True)
