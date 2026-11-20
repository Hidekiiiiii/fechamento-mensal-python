"""Resumo de 5 linhas para o e-mail ao diretor, gerado por regras a partir dos resultados (sem IA)."""
from __future__ import annotations

import pandas as pd

from .formatacao import compacto, data, lista, mes_curto, mil, mi, plural


def linhas_email(r) -> list[str]:
    conf, ef, d = r.conferencias, r.efeitos, r.dre.set_index("linha")
    mes = mes_curto(r.mes)
    pend = [c for c in conf.values() if c.status == "Pendência"]
    alertas = [c for c in conf.values() if c.status == "Alerta"]
    bloq = [c for c in conf.values() if c.status == "Bloqueia"]

    if bloq:
        l1 = (f"Fechamento de {mes} não apresentável: " + " e ".join(f"{c.codigo} ({c.nome.lower()})" for c in bloq) +
              " bloqueiam o pacote. Os números abaixo são preliminares; é preciso reextrair o ERP.")
    else:
        ok = [c.codigo for c in conf.values() if c.status == "OK"]
        l1 = (f"Fechamento de {mes} {r.veredito}: " + (lista(ok) + " OK" if ok else "") +
              (f"; {plural(len(pend), 'pendência', 'pendências')} ({lista([c.codigo for c in pend])})" if pend else "")
              + (f" e {plural(len(alertas), 'alerta', 'alertas')} ({lista([c.codigo for c in alertas])})"
                 if alertas else "") + ".")

    afeta = ef[ef["efeito_ebitda"].abs() >= 0.01]
    nomes = {"A1": "PDD", "A2": "provisão de férias e 13º", "A3": "despesas sem nota", "A4": "itens do banco"}
    ant = mes_curto(str(pd.Period(r.mes, "M") - 1))
    l2 = (f"EBITDA do mês: {mi(d.loc['EBITDA', 'mes_original'])} antes dos ajustes e "
          f"{mi(d.loc['EBITDA', 'mes_ajustado'])} após os ajustes propostos ("
          + lista([nomes[a] for a in afeta["ajuste"]]) + f": {compacto(afeta['efeito_ebitda'].sum())}), contra "
          f"{mi(d.loc['EBITDA', 'mes_anterior'])} em {ant}.")
    pot = r.potencial
    if pot is not None and not pot.empty:
        l2 += (f" Se as pendências de {lista(list(pot['conferencia']))} se confirmarem, vai a "
               f"{mi(d.loc['EBITDA', 'mes_ajustado'] + pot['efeito'].sum())}.")

    partes = []
    c4 = conf["C4"].detalhe
    lados = [n for n, k in (("clientes", "receber"), ("fornecedores", "pagar"))
             if (c4["status"].eq("Pendência") & c4["subledger_nome"].str.contains(k)).any()] if not c4.empty else []
    rotulos = {"C3": "banco", "C4": " e ".join(lados) or "subledgers", "C5": "folha", "C7": "saldos inesperados"}
    for cod, rot in rotulos.items():
        if conf[cod].status == "Pendência" and conf[cod].diferenca > 0:
            partes.append(f"{rot} {mil(conf[cod].diferenca)}")
    total = sum(conf[c].diferenca for c in rotulos if conf[c].status == "Pendência")
    c6 = conf["C6"]
    l3 = (f"Pendências de conciliação, em valor absoluto: {mil(total)} ({lista(partes)})." if partes else
          "Sem pendências de conciliação.")
    if c6.status != "OK":
        det6 = c6.detalhe
        fut, tard = det6[det6["futura"]], det6[det6["motivo"].str.startswith("Lançado")]
        trechos = []
        if len(fut):
            trechos.append(f"{mil(fut['valor'].sum())} de competência futura a estornar e relançar")
        if len(tard):
            trechos.append(f"{mil(tard['valor'].sum())} lançados depois do corte, a validar")
        l3 += f" Corte: " + lista(trechos) + "."

    itens = conf["C3"].detalhe
    p3 = itens[itens["situacao"] == "Pendência"] if not itens.empty else itens
    if p3.empty:
        l4 = "Caixa: bancos conciliados, sem itens pendentes com a tesouraria."
    else:
        desc = []
        for x in p3.itertuples():
            if "identificado" in x.tipo:
                desc.append(f"um crédito não identificado de {mil(abs(x.valor))}")
            else:
                desc.append(f"um cheque de {mil(abs(x.valor))} em aberto há {int(x.dias_uteis)} dias úteis")
        l4 = f"Caixa: {plural(len(p3), 'item', 'itens')} com a tesouraria: " + lista(desc) + "."

    c8 = conf["C8"]
    if c8.status == "Alerta":
        x = c8.detalhe.sort_values("delta", key=abs, ascending=False).iloc[0]
        un = r.unidades.get(x.unidade, x.unidade)
        l5 = (f"{x.descricao} da {un} {'+' if x.delta > 0 else '−'}{abs(x.variacao):.0%} contra "
              f"{mes_curto(str(pd.Period(r.mes, 'M') - 1))} ({'+' if x.delta > 0 else '−'}{mil(abs(x.delta))}), com "
              f"justificativa pendente do gestor. Prazo: {data(c8.prazo)}.")
        if len(c8.detalhe) > 1:
            l5 += f" Outras {len(c8.detalhe) - 1} variações na aba Exceções."
    else:
        l5 = "Nenhuma variação atípica contra o mês anterior."
    return [l1, l2, l3, l4, l5]
