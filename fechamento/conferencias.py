"""As 8 conferências automáticas do fechamento.

Status possíveis:
- OK
- Bloqueia: o pacote sai como "não apresentável" (C1 e C2: o número em si não é confiável)
- Pendência: o pacote sai "com ressalvas", com o valor em R$ a resolver
- Alerta: não é erro, mas exige justificativa
"""
from __future__ import annotations

from dataclasses import dataclass, field

import numpy as np
import pandas as pd

from . import plano
from .formatacao import brl, data, num
from .leitura import Entradas

CHAVE = ["conta", "unidade"]
EPS = 0.005   # meio centavo: abaixo disso é ruído de ponto flutuante


@dataclass
class Conferencia:
    codigo: str
    nome: str
    regra: str
    tolerancia: str
    status: str
    diferenca: float
    resultado: str
    responsavel: str
    prazo: pd.Timestamp
    detalhe: pd.DataFrame = field(default_factory=pd.DataFrame)
    extra: dict = field(default_factory=dict)


def _conf(e: Entradas, codigo, nome, regra, tolerancia, status, diferenca, resultado, detalhe=None, **extra):
    return Conferencia(codigo, nome, regra, tolerancia, status, round(float(diferenca), 2), resultado,
                       e.p.responsavel(codigo), e.p.prazo(codigo),
                       detalhe if detalhe is not None else pd.DataFrame(), extra)


def _competencia(s: pd.Series) -> pd.Series:
    """'10/2025' -> Period('2025-10')."""
    return pd.PeriodIndex(pd.to_datetime("01/" + s, format="%d/%m/%Y"), freq="M").to_series(index=s.index)


# --------------------------------------------------------------------------- C1
def c1_integridade(e: Entradas) -> Conferencia:
    b = e.balancete.copy()
    tol = e.p.tol("C1")
    itens = []
    dif_total = round(b["debitos"].sum() - b["creditos"].sum(), 2)
    if abs(dif_total) > tol + EPS:
        itens.append(dict(teste="Total de débitos = total de créditos", conta="", unidade="", diferenca=dif_total))
    b["dif"] = (b["saldo_final"] - (b["saldo_inicial"] + b["debitos"] - b["creditos"])).round(2)
    for r in b[b["dif"].abs() > tol + EPS].itertuples():
        itens.append(dict(teste="Saldo final = inicial + débitos − créditos", conta=r.conta, unidade=r.unidade,
                          diferenca=r.dif))
    ant = e.anterior[CHAVE + ["saldo_final"]].rename(columns={"saldo_final": "final_anterior"})
    m = b[CHAVE + ["saldo_inicial"]].merge(ant, on=CHAVE, how="outer").fillna({"saldo_inicial": 0, "final_anterior": 0})
    m["dif"] = (m["saldo_inicial"] - m["final_anterior"]).round(2)
    for r in m[m["dif"].abs() > tol + EPS].itertuples():
        itens.append(dict(teste="Saldo inicial = saldo final do pacote anterior", conta=r.conta, unidade=r.unidade,
                          diferenca=r.dif))
    det = pd.DataFrame(itens, columns=["teste", "conta", "unidade", "diferenca"])
    ok = det.empty
    resultado = ("Balancete íntegro: débitos = créditos, saldos fecham e abrem no final do mês anterior." if ok else
                 f"{len(det)} falha(s): " + "; ".join(f"{t} ({n})" for t, n in det["teste"].value_counts().items()))
    return _conf(e, "C1", "Integridade do balancete",
                 "Σ débitos = Σ créditos; por conta, final = inicial + D − C; inicial = final do pacote anterior",
                 "R$ 0,00", "OK" if ok else "Bloqueia", det["diferenca"].abs().sum(), resultado, det,
                 erro_aritmetico=det[det["teste"].str.startswith("Saldo final")],
                 saldo_inicial=det[det["teste"].str.startswith("Saldo inicial")])


# --------------------------------------------------------------------------- C2
def c2_razao_balancete(e: Entradas) -> Conferencia:
    tol = e.p.tol("C2")
    r = e.razao.groupby(CHAVE, as_index=False)[["debito", "credito"]].sum()
    m = e.balancete[CHAVE + ["debitos", "creditos"]].merge(r, on=CHAVE, how="outer").fillna(0)
    m["dif_debitos"] = (m["debitos"] - m["debito"]).round(2)
    m["dif_creditos"] = (m["creditos"] - m["credito"]).round(2)
    det = m[(m["dif_debitos"].abs() > tol) | (m["dif_creditos"].abs() > tol)].rename(
        columns={"debitos": "debitos_balancete", "creditos": "creditos_balancete", "debito": "debitos_razao",
                 "credito": "creditos_razao"})
    ok = det.empty
    deb, cred = det["dif_debitos"].sum(), det["dif_creditos"].sum()
    resultado = ("Razão e balancete batem em todas as contas." if ok else
                 f"{len(det)} conta(s) × unidade fora da tolerância: o balancete tem {brl(deb, 2)} de débitos e "
                 f"{brl(cred, 2)} de créditos que não estão no razão.")
    return _conf(e, "C2", "Razão × balancete", "Σ débitos e Σ créditos do razão por conta × unidade = movimento do "
                 "balancete", f"R$ {num(tol, 2)} por conta", "OK" if ok else "Bloqueia",
                 det["dif_debitos"].abs().sum(), resultado, det.reset_index(drop=True),
                 debitos_faltantes=round(deb, 2), creditos_faltantes=round(cred, 2))


# --------------------------------------------------------------------------- C3
def classificar_banco(descricao: str) -> str:
    d = descricao.upper()
    if "TARIFA" in d or "IOF" in d:
        return "Tarifa/IOF não contabilizado"
    if "RENDIMENTO" in d:
        return "Rendimento não contabilizado"
    return "Crédito não identificado" if not d.startswith("DEB") else "Débito não identificado"


def c3_bancos(e: Entradas) -> Conferencia:
    p = e.p
    limite_dias = int(p.regra("C3", "dias_uteis_em_aberto"))
    base = np.datetime64(p.data_base.date()) + np.timedelta64(1, "D")      # conta o próprio dia-base
    bancos = p.bancos.set_index("banco")
    livro = e.razao[e.razao["conta"].isin(bancos["conta_contabil"])].assign(
        valor=lambda x: (x["debito"] - x["credito"]).round(2))[["conta", "data_lanc", "documento", "historico", "valor"]]
    ext = e.extrato.assign(conta=e.extrato["banco"].map(bancos["conta_contabil"]), valor=e.extrato["valor"].round(2))
    m = livro.merge(ext[["conta", "documento", "valor", "data", "descricao"]], on=["conta", "documento", "valor"],
                    how="outer", indicator=True)
    itens, blocos = [], []
    for banco, conta in bancos["conta_contabil"].items():
        so_livro = m[(m["conta"] == conta) & (m["_merge"] == "left_only")]
        so_banco = m[(m["conta"] == conta) & (m["_merge"] == "right_only")]
        for r in so_livro.itertuples():
            dias = int(np.busday_count(np.datetime64(r.data_lanc.date()), base)) - 1
            tipo = "Depósito em trânsito" if r.valor > 0 else "Cheque/pagamento não compensado"
            itens.append(dict(banco=banco, origem="Contabilidade", tipo=tipo, data=r.data_lanc, documento=r.documento,
                              descricao=r.historico, valor=r.valor, dias_uteis=dias,
                              situacao="Pendência" if dias > limite_dias else "No prazo"))
        for r in so_banco.itertuples():
            tipo = classificar_banco(r.descricao)
            itens.append(dict(banco=banco, origem="Banco", tipo=tipo, data=r.data, documento=r.documento,
                              descricao=r.descricao, valor=r.valor, dias_uteis=None,
                              situacao="Pendência" if "identificado" in tipo else "Ajuste A4 proposto"))
        it = pd.DataFrame(itens, columns=["banco", "origem", "tipo", "valor"])
        it = it[it["banco"] == banco]
        saldo_ext = float(e.extrato.loc[e.extrato["banco"] == banco, "saldo"].iloc[-1])
        saldo_cont = float(e.balancete.loc[e.balancete["conta"] == conta, "saldo_final"].sum())
        transito = it.loc[it["tipo"] == "Depósito em trânsito", "valor"].sum()
        cheques = -it.loc[it["tipo"] == "Cheque/pagamento não compensado", "valor"].sum()
        do_banco = it[it["origem"] == "Banco"]
        creditos = do_banco.loc[do_banco["valor"] > 0, "valor"].sum()
        debitos = -do_banco.loc[do_banco["valor"] < 0, "valor"].sum()
        banco_aj = saldo_ext + transito - cheques
        cont_aj = saldo_cont + creditos - debitos
        sem_ajuste = do_banco.loc[do_banco["tipo"].str.contains("identificado"), "valor"].sum()
        blocos.append(dict(banco=banco, conta=conta, saldo_extrato=saldo_ext, depositos_transito=transito,
                           cheques_nao_compensados=cheques, saldo_banco_ajustado=round(banco_aj, 2),
                           saldo_contabil=saldo_cont, creditos_nao_contabilizados=creditos,
                           debitos_nao_contabilizados=debitos, saldo_contabil_ajustado=round(cont_aj, 2),
                           diferenca=round(cont_aj - banco_aj, 2), sem_lancamento_proposto=round(sem_ajuste, 2),
                           saldo_balancete_ajustado=round(cont_aj - sem_ajuste, 2)))
    itens = pd.DataFrame(itens)
    blocos = pd.DataFrame(blocos)
    pend = itens[itens["situacao"] == "Pendência"]
    nao_explicada = float(blocos["diferenca"].abs().sum())
    pendencia = float(pend["valor"].abs().sum()) + nao_explicada
    status = "Pendência" if pendencia > p.tol("C3") + EPS else "OK"
    partes = [f"{len(pend)} item(ns) pendente(s) somando {brl(pend['valor'].abs().sum(), 2)}"]
    partes.append("diferença não explicada de " + brl(nao_explicada, 2) if nao_explicada > EPS else
                  "sem diferença não explicada")
    eh = lambda t: itens["tipo"].eq(t)
    return _conf(e, "C3", "Conciliação bancária",
                 "Saldo do extrato ± itens em aberto = saldo contábil, por banco, com cada item classificado",
                 f"Diferença não explicada R$ 0,00; item em aberto até {limite_dias} dias úteis", status, pendencia,
                 "; ".join(partes) + ".", itens, blocos=blocos, pendencia=round(pendencia, 2),
                 diferenca_nao_explicada=round(nao_explicada, 2),
                 cheque_antigo=round(-pend.loc[pend["tipo"].str.startswith("Cheque"), "valor"].sum(), 2),
                 nao_identificado=round(pend.loc[pend["tipo"].str.contains("identificado"), "valor"].abs().sum(), 2),
                 transito_no_prazo=round(itens.loc[eh("Depósito em trânsito") & itens["situacao"].eq("No prazo"),
                                                   "valor"].sum(), 2),
                 tarifas=round(itens.loc[eh("Tarifa/IOF não contabilizado"), "valor"].sum(), 2),
                 rendimentos=round(itens.loc[eh("Rendimento não contabilizado"), "valor"].sum(), 2))


# --------------------------------------------------------------------------- C4
def c4_subledgers(e: Entradas) -> Conferencia:
    tol = e.p.tol("C4")
    b = e.balancete
    cr = pd.DataFrame({"subledger": e.aging_cr.groupby("unidade")["valor_aberto"].sum(),
                       "razao": b[b["conta"] == plano.CLIENTES].groupby("unidade")["saldo_final"].sum()}).fillna(0)
    cr.insert(0, "conta", plano.CLIENTES)
    cr.insert(0, "subledger_nome", "Contas a receber (aging)")
    cp_conta = e.aging_cp["conta_contabil"].iloc[0] if len(e.aging_cp) else plano.FORNECEDORES
    cp = pd.DataFrame({"subledger": e.aging_cp.groupby("unidade")["valor_aberto"].sum(),
                       "razao": -b[b["conta"] == cp_conta].groupby("unidade")["saldo_final"].sum()}).fillna(0)
    cp.insert(0, "conta", cp_conta)
    cp.insert(0, "subledger_nome", "Contas a pagar (aging)")
    det = pd.concat([cr, cp]).rename_axis("unidade").reset_index()
    det["diferenca"] = (det["razao"] - det["subledger"]).round(2)
    det["status"] = np.where(det["diferenca"].abs() > tol, "Pendência", "OK")
    fora = det[det["status"] == "Pendência"]
    ok = fora.empty
    resultado = ("Aging de clientes e de fornecedores batem com o razão em todas as unidades." if ok else
                 "; ".join(f"{r.subledger_nome.split(' (')[0]} {r.unidade}: razão {brl(r.razao, 2)} × aging "
                           f"{brl(r.subledger, 2)} (Δ {brl(r.diferenca, 2)})" for r in fora.itertuples()) + ".")
    return _conf(e, "C4", "Subledgers × razão", "Aging de clientes = conta de clientes; aging de fornecedores = "
                 "conta de fornecedores, por unidade", f"R$ {num(tol)} por unidade", "OK" if ok else "Pendência",
                 fora["diferenca"].abs().sum(), resultado, det, unidades=list(fora["unidade"]),
                 razao_cr=round(cr["razao"].sum(), 2), aging_cr=round(cr["subledger"].sum(), 2))


# --------------------------------------------------------------------------- C5
def c5_folha(e: Entradas) -> Conferencia:
    tol = e.p.tol("C5")
    contas = str(e.p.regra("C5", "contas_pessoal")).split(";")
    origem = str(e.p.regra("C5", "origem_folha"))
    r = e.razao[e.razao["conta"].isin(contas) & (e.razao["origem"] == origem)]
    det = pd.DataFrame({"folha": e.folha.groupby("centro_custo")["valor"].sum(),
                        "razao": (r["debito"] - r["credito"]).groupby(r["centro_custo"]).sum()}).fillna(0)
    det = det.rename_axis("centro_custo").reset_index()
    det["diferenca"] = (det["folha"] - det["razao"]).round(2)
    det["status"] = np.where(det["diferenca"].abs() > tol, "Pendência", "OK")
    fora = det[det["status"] == "Pendência"]
    ok = fora.empty
    resultado = ("Folha e razão batem em todos os centros de custo." if ok else
                 "; ".join(f"{r.centro_custo}: folha {brl(r.folha, 2)} × razão {brl(r.razao, 2)} "
                           f"(Δ {brl(r.diferenca, 2)})" for r in fora.itertuples()) + ".")
    return _conf(e, "C5", "Folha × razão", "Σ das rubricas da folha por centro de custo = contas de pessoal "
                 "lançadas pela integração da folha", f"R$ {num(tol)} por centro de custo",
                 "OK" if ok else "Pendência", fora["diferenca"].abs().sum(), resultado, det,
                 centros=list(fora["centro_custo"]))


# --------------------------------------------------------------------------- C6
def c6_corte(e: Entradas) -> Conferencia:
    p = e.p
    r = e.razao.copy()
    mes = pd.Period(p.mes, "M")
    r["comp"] = _competencia(r["data_competencia"])
    lanc = r.groupby("id_lanc").agg(data_lanc=("data_lanc", "first"), comp=("comp", "first"),
                                    historico=("historico", "first"), valor=("debito", "sum"))
    deb = r[r["debito"] > 0].groupby("id_lanc").first()[["conta", "unidade", "centro_custo"]]
    lanc = lanc.join(deb)
    # provisão revertida no mês: nota de mês anterior cujo valor teve a provisão estornada (débito em provisões)
    rev = r[(r["conta"] == plano.PROV_PAGAR) & (r["debito"] > 0) & (r["comp"] == mes)]
    revertidas = set(zip(rev["unidade"], rev["debito"].round(2)))
    provisionada = pd.Series([(u, round(v, 2)) in revertidas for u, v in zip(lanc["unidade"], lanc["valor"])],
                             index=lanc.index)
    fora_comp = (lanc["comp"] != mes) & ~((lanc["comp"] < mes) & provisionada)
    apos = lanc["data_lanc"] > p.data_corte
    det = lanc[fora_comp | apos].copy()
    det["motivo"] = np.where(fora_comp[det.index],
                             "Competência " + det["comp"].astype(str).str[5:] + "/" + det["comp"].astype(str).str[:4]
                             + " lançada em " + mes.strftime("%m/%Y"),
                             "Lançado em " + det["data_lanc"].dt.strftime("%d/%m/%Y") + ", depois do corte de "
                             + data(p.data_corte))
    det["material"] = np.where(det["valor"] > p.tol("C6"), "Sim", "Não")
    det["acao"] = np.select(
        [fora_comp[det.index] & (det["comp"] > mes), fora_comp[det.index]],
        [f"Estornar em {mes.strftime('%m/%Y')} e relançar no mês de competência",
         "Reclassificar para a competência correta"],
        "Competência correta: validar documento e aprovação do lançamento tardio")
    det = det.reset_index()[["id_lanc", "data_lanc", "comp", "conta", "unidade", "centro_custo", "historico", "valor",
                             "motivo", "material", "acao"]]
    det["futura"] = fora_comp[det["id_lanc"]].to_numpy() & (lanc.loc[det["id_lanc"], "comp"] > mes).to_numpy()
    det["comp"] = det["comp"].astype(str)
    ok = det.empty
    n_comp, n_corte = int(fora_comp.sum()), int((apos & ~fora_comp).sum())
    v_comp = float(lanc.loc[fora_comp, "valor"].sum())
    v_corte = float(lanc.loc[apos & ~fora_comp, "valor"].sum())
    resultado = ("Nenhum lançamento fora da competência ou depois do corte." if ok else
                 f"{len(det)} lançamento(s): {n_comp} de outra competência ({brl(v_comp, 2)}) e {n_corte} depois "
                 f"do corte ({brl(v_corte, 2)}); {int((det['material'] == 'Sim').sum())} material(is).")
    return _conf(e, "C6", "Corte e competência", "Competência diferente do mês, ou data de lançamento depois da data "
                 f"de corte ({data(p.data_corte)})", f"0 lançamentos; material acima de R$ {num(p.tol('C6'))}",
                 "OK" if ok else "Pendência", det["valor"].sum(), resultado, det, lancamentos=len(det),
                 competencia=round(v_comp, 2), corte=round(v_corte, 2))


# --------------------------------------------------------------------------- C7
def c7_saldos(e: Entradas) -> Conferencia:
    p = e.p
    tol = p.tol("C7")
    b = e.balancete.assign(natureza=e.balancete["conta"].map(p.natureza()), tipo=e.balancete["conta"].map(p.tipo()))
    inv = b[(b["tipo"] == "patrimonial") & (((b["natureza"] == "D") & (b["saldo_final"] < -tol)) |
                                           ((b["natureza"] == "C") & (b["saldo_final"] > tol)))]
    trans = b[(b["tipo"] == "transitoria") & (b["saldo_final"].abs() > tol)]
    itens = [dict(tipo="Natureza invertida", conta=r.conta, unidade=r.unidade, centro_custo="",
                  descricao=f"{r.descricao}: saldo {'credor' if r.saldo_final < 0 else 'devedor'} em conta de "
                            f"natureza {'devedora' if r.natureza == 'D' else 'credora'}",
                  valor=abs(r.saldo_final), tratamento="Pendência") for r in inv.itertuples()]
    itens += [dict(tipo="Transitória com saldo", conta=r.conta, unidade=r.unidade, centro_custo="",
                   descricao=f"{r.descricao}: saldo de {brl(r.saldo_final, 2)} sem classificação",
                   valor=abs(r.saldo_final), tratamento="Pendência") for r in trans.itertuples()]
    pct = float(p.regra("C7", "percentual_minimo_recorrente"))
    raz = e.razao[e.razao["data_competencia"] == pd.Period(p.mes, "M").strftime("%m/%Y")]
    faltantes = []
    for r in p.recorrentes.itertuples():
        x = raz[(raz["conta"] == r.conta) & (raz["centro_custo"] == r.centro_custo)]
        lancado = float((x["debito"] - x["credito"]).sum())
        if lancado < pct * r.valor_referencia:
            faltantes.append(r)
            itens.append(dict(tipo="Despesa recorrente sem lançamento", conta=r.conta,
                              unidade=r.centro_custo.split("-")[0], centro_custo=r.centro_custo,
                              descricao=f"{r.descricao}: lançado {brl(lancado, 2)}, referência "
                                        f"{brl(r.valor_referencia, 2)}",
                              valor=r.valor_referencia, tratamento="Ajuste A3 proposto"))
    det = pd.DataFrame(itens, columns=["tipo", "conta", "unidade", "centro_custo", "descricao", "valor", "tratamento"])
    pend = det[det["tratamento"] == "Pendência"]
    ok = det.empty
    resultado = ("Nenhum saldo inesperado." if ok else
                 f"{len(pend)} saldo(s) inesperado(s) somando {brl(pend['valor'].sum(), 2)}; "
                 f"{len(faltantes)} despesa(s) recorrente(s) sem lançamento, cobertas pelo ajuste A3.")
    status = "Pendência" if not pend.empty else ("Pendência" if faltantes else "OK")
    return _conf(e, "C7", "Saldo inesperado", "Natureza invertida, transitória com saldo ou despesa recorrente sem "
                 f"lançamento (abaixo de {pct:.0%} da referência)", f"R$ {num(tol)}", status, pend["valor"].sum(),
                 resultado, det, faltantes=pd.DataFrame(faltantes))


# --------------------------------------------------------------------------- C8
def c8_variacao(e: Entradas) -> Conferencia:
    p = e.p
    limite_rs, limite_pct = p.tol("C8"), float(p.regra("C8", "variacao_percentual"))
    tipo = p.tipo()

    def movimento(df):
        df = df[df["conta"].map(tipo) == "resultado"]
        return (df["debitos"] - df["creditos"]).groupby([df["conta"], df["unidade"]]).sum()

    m = pd.DataFrame({"anterior": movimento(e.anterior), "atual": movimento(e.balancete)}).fillna(0).round(2)
    m["delta"] = (m["atual"] - m["anterior"]).round(2)
    m["variacao"] = np.where(m["anterior"] != 0, m["delta"] / m["anterior"].abs(), np.nan)
    flag = (m["delta"].abs() > limite_rs) & ((m["anterior"] == 0) | (m["variacao"].abs() > limite_pct))
    det = m[flag].reset_index()
    det.insert(1, "descricao", det["conta"].map(p.contas.set_index("conta")["descricao"]))
    det["justificativa"] = ""
    resultado = ("Nenhuma variação atípica contra o mês anterior." if det.empty else
                 "; ".join(f"{r.descricao} {r.unidade}: {'+' if r.delta > 0 else ''}{r.variacao:.0%} "
                           f"({'+' if r.delta > 0 else '−'}{brl(abs(r.delta))})" for r in det.itertuples()) +
                 ". Exige justificativa do gestor.")
    return _conf(e, "C8", "Variação atípica", f"Conta de resultado com |Δ| contra o mês anterior acima de "
                 f"{limite_pct:.0%} e de R$ {num(limite_rs)}", "Os dois critérios juntos",
                 "OK" if det.empty else "Alerta", det["delta"].abs().sum(), resultado, det)


TODAS = [c1_integridade, c2_razao_balancete, c3_bancos, c4_subledgers, c5_folha, c6_corte, c7_saldos, c8_variacao]


def rodar(e: Entradas) -> dict[str, Conferencia]:
    return {c.codigo: c for c in (f(e) for f in TODAS)}
