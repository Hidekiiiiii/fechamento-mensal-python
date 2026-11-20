"""Linha de comando.

    python -m fechamento                  # out/25, extração mais recente
    python -m fechamento --extracao 1     # primeira extração (com erros de extração)
    python -m fechamento --mes 2025-10 --entrada <pasta> --anterior <csv> --parametros <xlsx> --saida <pasta>
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from . import gravar, processar
from .formatacao import brl, mes_anterior, mes_curto

RAIZ = Path(__file__).resolve().parents[1]


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(prog="python -m fechamento", description="Fechamento mensal automatizado")
    ap.add_argument("--mes", default="2025-10", help="mês do fechamento (AAAA-MM)")
    ap.add_argument("--extracao", type=int, help="número da extração (padrão: a mais recente)")
    ap.add_argument("--entrada", type=Path, help="pasta com as 6 exportações do ERP")
    ap.add_argument("--anterior", type=Path, help="balancete ajustado do mês anterior (CSV)")
    ap.add_argument("--parametros", type=Path, default=RAIZ / "entradas" / "parametros.xlsx")
    ap.add_argument("--saida", type=Path, help="pasta de saída do pacote")
    a = ap.parse_args(argv)

    pasta_mes = RAIZ / "entradas" / a.mes
    if a.entrada is None:
        extracoes = sorted(pasta_mes.glob("extracao_*"), key=lambda p: int(p.name.split("_")[1]))
        if not extracoes:
            print(f"Nenhuma extração encontrada em {pasta_mes}", file=sys.stderr)
            return 2
        a.entrada = pasta_mes / f"extracao_{a.extracao}" if a.extracao else extracoes[-1]
    ant = mes_anterior(a.mes)
    a.anterior = a.anterior or RAIZ / "entradas" / ant / f"balancete_ajustado_{ant}.csv"
    a.saida = a.saida or RAIZ / "saida" / a.mes / a.entrada.name

    r = processar(a.entrada, a.anterior, a.parametros)
    arquivos = gravar(r, a.saida)

    print(f"Fechamento de {mes_curto(r.mes)} ({a.entrada.name}): {r.veredito.upper()}")
    for c in r.conferencias.values():
        print(f"  {c.codigo} {c.nome:<26} {c.status:<10} {brl(c.diferenca, 2):>18}")
    print("\nResumo para o e-mail:")
    for i, linha in enumerate(r.email, 1):
        print(f"  {i}. {linha}")
    print("\nArquivos gerados:")
    for p in arquivos.values():
        print(f"  {p.relative_to(RAIZ) if p.is_relative_to(RAIZ) else p}")
    return 1 if r.veredito == "não apresentável" else 0


if __name__ == "__main__":
    raise SystemExit(main())
