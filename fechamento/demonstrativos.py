"""Balancete ajustado e DRE (mês e acumulado, antes e depois dos ajustes)."""
from __future__ import annotations

import pandas as pd

from .ajustes import Ajuste
from .leitura import Entradas
from .plano import ESTRUTURA_DRE

CHAVE = ["conta", "unidade"]


def balancete_ajustado(e: Entradas, ajustes: dict[str, Ajuste]) -> pd.DataFrame:
    linhas = pd.concat([a.linhas for a in ajustes.values()], ignore_index=True)
    aj = linhas.groupby(CHAVE, as_index=False)[["debito", "credito"]].sum().rename(
        columns={"debito": "debitos_ajuste", "credito": "creditos_ajuste"})
    b = e.balancete.drop(columns="descricao", errors="ignore").merge(aj, on=CHAVE, how="outer").fillna(0)
    b.insert(1, "descricao", b["conta"].map(e.p.contas.set_index("conta")["descricao"]))
    b["ajustes"] = (b["debitos_ajuste"] - b["creditos_ajuste"]).round(2)
    b["saldo_ajustado"] = (b["saldo_final"] + b["ajustes"]).round(2)
    return b.sort_values(CHAVE).reset_index(drop=True)


def para_proximo_mes(bal_aj: pd.DataFrame) -> pd.DataFrame:
    """Balancete ajustado no mesmo layout do pacote anterior: vira o saldo inicial do mês seguinte."""
    return pd.DataFrame({"conta": bal_aj["conta"], "descricao": bal_aj["descricao"], "unidade": bal_aj["unidade"],
                         "saldo_inicial": bal_aj["saldo_inicial"],
                         "debitos": (bal_aj["debitos"] + bal_aj["debitos_ajuste"]).round(2),
                         "creditos": (bal_aj["creditos"] + bal_aj["creditos_ajuste"]).round(2),
                         "saldo_final": bal_aj["saldo_ajustado"]})


def dre(e: Entradas, bal_aj: pd.DataFrame) -> pd.DataFrame:
    grupo = e.p.grupo()

    def por_grupo(df: pd.DataFrame, valor: pd.Series) -> pd.Series:
        g = df["conta"].map(grupo)
        return (-valor).groupby(g).sum()     # receita positiva, despesa negativa

    ant = e.anterior
    colunas = {
        "mes_original": por_grupo(bal_aj, bal_aj["debitos"] - bal_aj["creditos"]),
        "ajustes": por_grupo(bal_aj, bal_aj["ajustes"]),
        "mes_anterior": por_grupo(ant, ant["debitos"] - ant["creditos"]),
        "acumulado_original": por_grupo(bal_aj, bal_aj["saldo_final"]),
        "acumulado_ajustado": por_grupo(bal_aj, bal_aj["saldo_ajustado"]),
    }
    linhas = []
    for nome, grupos, tipo in ESTRUTURA_DRE:
        v = {k: round(float(s.reindex(list(grupos)).fillna(0).sum()), 2) for k, s in colunas.items()}
        v["mes_ajustado"] = round(v["mes_original"] + v["ajustes"], 2)
        v["delta_mes_anterior"] = round(v["mes_ajustado"] - v["mes_anterior"], 2)
        linhas.append(dict(linha=nome, tipo=tipo, **v))
    cols = ["linha", "tipo", "mes_original", "ajustes", "mes_ajustado", "mes_anterior", "delta_mes_anterior",
            "acumulado_original", "acumulado_ajustado"]
    return pd.DataFrame(linhas)[cols]
