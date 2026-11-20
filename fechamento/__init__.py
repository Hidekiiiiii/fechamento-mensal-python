"""Automação do fechamento mensal: lê as exportações do ERP, confere, propõe ajustes e gera o pacote."""
from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

import pandas as pd

from . import ajustes as A
from . import conferencias as C
from . import demonstrativos as D
from .leitura import Entradas, ler
from .pendencias import efeito_potencial
from .resumo import linhas_email

PENDENCIAS = ("C3", "C4", "C5", "C7")


@dataclass
class Resultado:
    mes: str
    extracao: str
    gerado_em: datetime
    entradas: Entradas
    conferencias: dict
    ajustes: dict
    efeitos: pd.DataFrame
    balancete_ajustado: pd.DataFrame
    dre: pd.DataFrame
    veredito: str
    pendencias_total: float
    unidades: dict
    potencial: pd.DataFrame | None = None     # efeito no EBITDA das pendências medidas que não viraram ajuste
    email: list[str] | None = None


def veredito(conferencias: dict) -> str:
    status = {c.status for c in conferencias.values()}
    if "Bloqueia" in status:
        return "não apresentável"
    return "com ressalvas" if status & {"Pendência", "Alerta"} else "apresentável"


def processar(pasta: Path, anterior: Path, parametros: Path) -> Resultado:
    e = ler(Path(pasta), Path(anterior), Path(parametros))
    conf = C.rodar(e)
    aj = A.propor(e, conf)
    bal = D.balancete_ajustado(e, aj)
    r = Resultado(mes=e.p.mes, extracao=Path(pasta).name, gerado_em=datetime.now(), entradas=e, conferencias=conf,
                  ajustes=aj, efeitos=A.efeitos(e, aj), balancete_ajustado=bal, dre=D.dre(e, bal),
                  veredito=veredito(conf),
                  pendencias_total=round(sum(conf[c].diferenca for c in PENDENCIAS if conf[c].status == "Pendência"), 2),
                  unidades=e.p.unidades["nome"].to_dict(), potencial=efeito_potencial(conf, e.p.grupo()))
    r.email = linhas_email(r)
    return r


def gravar(r: Resultado, saida: Path) -> dict[str, Path]:
    from .pacote import gravar_excel

    saida = Path(saida)
    saida.mkdir(parents=True, exist_ok=True)
    arquivos = {"pacote": saida / f"fechamento_{r.mes}.xlsx", "email": saida / f"resumo_email_{r.mes}.txt",
                "balancete": saida / f"balancete_ajustado_{r.mes}.csv"}
    gravar_excel(r, arquivos["pacote"])
    arquivos["email"].write_text("\n".join(f"{i}. {l}" for i, l in enumerate(r.email, 1)) + "\n", encoding="utf-8")
    D.para_proximo_mes(r.balancete_ajustado).to_csv(arquivos["balancete"], sep=";", decimal=",", float_format="%.2f",
                                                    index=False, encoding="utf-8-sig")
    return arquivos
