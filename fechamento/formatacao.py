"""Formatação de valores em português do Brasil (para textos; o Excel usa formatos numéricos próprios)."""
from __future__ import annotations

from decimal import ROUND_HALF_UP, Decimal

import pandas as pd

MESES = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"]


def num(v: float, casas: int = 0) -> str:
    """Arredonda meio para cima (como o Excel), não pelo arredondamento bancário do Python."""
    q = Decimal(repr(abs(float(v)))).quantize(Decimal(1).scaleb(-casas), rounding=ROUND_HALF_UP)
    s = f"{q:,.{casas}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def lista(itens: list[str]) -> str:
    itens = list(itens)
    return itens[0] if len(itens) == 1 else ", ".join(itens[:-1]) + " e " + itens[-1]


def plural(n: int, singular: str, plural_: str) -> str:
    return f"{n} {singular if n == 1 else plural_}"


def brl(v: float, casas: int = 0) -> str:
    return ("−" if v < 0 else "") + "R$ " + num(v, casas)


def mi(v: float) -> str:
    return ("−" if v < 0 else "") + f"R$ {num(abs(v) / 1e6, 2)} mi"


def mil(v: float) -> str:
    s = num(abs(v) / 1e3, 1)
    if s.endswith(",0"):
        s = s[:-2]
    return ("−" if v < 0 else "") + f"R$ {s} mil"


def compacto(v: float) -> str:
    """R$ 1,42 mi acima de 1 milhão; R$ 460,3 mil abaixo."""
    return mi(v) if abs(v) >= 1e6 else mil(v)


def data(ts: pd.Timestamp) -> str:
    return pd.Timestamp(ts).strftime("%d/%m/%Y")


def mes_curto(mes: str) -> str:
    ano, m = mes.split("-")
    return f"{MESES[int(m) - 1]}/{ano[2:]}"


def mes_anterior(mes: str) -> str:
    return str(pd.Period(mes, "M") - 1)
