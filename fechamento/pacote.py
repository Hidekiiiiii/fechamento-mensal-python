"""Gravação do pacote de fechamento em Excel, formatado como um controller entregaria.

- Números: #.##0 com negativos entre parênteses e em vermelho; centavos só onde há conciliação.
- Cabeçalho congelado e autofiltro na tabela principal de cada aba; totais em negrito com borda dupla.
- Cor só em Bloqueia e Pendência; datas dd/mm/aaaa; abas numeradas.
"""
from __future__ import annotations

import math
from pathlib import Path

import pandas as pd
import xlsxwriter

from .formatacao import data, mes_curto

INT = '#,##0;[Red](#,##0);"–"'
DEC = '#,##0.00;[Red](#,##0.00);"–"'
PCT = '0.0%;[Red]-0.0%;"–"'


class Livro:
    def __init__(self, caminho: Path):
        self.wb = xlsxwriter.Workbook(str(caminho), {"nan_inf_to_errors": True})
        base = {"font_name": "Calibri", "font_size": 10, "valign": "vcenter"}
        f = lambda **k: self.wb.add_format(base | k)
        self.f = {
            "titulo": f(bold=True, font_size=14), "sub": f(italic=True, font_color="#595959"),
            "secao": f(bold=True, font_size=11, font_color="#1F3864"),
            "cab": f(bold=True, bg_color="#D9E1F2", bottom=1, text_wrap=True, valign="top"),
            "texto": f(), "quebra": f(text_wrap=True, valign="top"), "negrito": f(bold=True),
            "int": f(num_format=INT), "dec": f(num_format=DEC), "pct": f(num_format=PCT),
            "data": f(num_format="dd/mm/yyyy", align="left"),
            "t_texto": f(bold=True, top=6), "t_int": f(bold=True, top=6, num_format=INT),
            "t_dec": f(bold=True, top=6, num_format=DEC), "t_pct": f(bold=True, top=6, num_format=PCT),
            "s_texto": f(bold=True), "s_int": f(bold=True, num_format=INT), "s_dec": f(bold=True, num_format=DEC),
            "Bloqueia": f(bold=True, bg_color="#F4CCCC", font_color="#9C0006"),
            "Pendência": f(bold=True, bg_color="#FCE4D6", font_color="#833C0C"),
            "veredito_não apresentável": f(bold=True, font_size=16, bg_color="#F4CCCC", font_color="#9C0006",
                                           border=1),
            "veredito_com ressalvas": f(bold=True, font_size=16, bg_color="#FCE4D6", font_color="#833C0C", border=1),
            "veredito_apresentável": f(bold=True, font_size=16, border=1),
            "campo": f(bottom=1),
        }

    def aba(self, nome: str):
        ws = self.wb.add_worksheet(nome)
        ws.hide_gridlines(2)
        ws.set_zoom(90)
        return ws

    def escrever(self, ws, lin: int, col: int, valor, fmt: str = "texto"):
        if valor is None or (isinstance(valor, float) and math.isnan(valor)) or valor is pd.NaT:
            ws.write_blank(lin, col, None, self.f[fmt])
        elif isinstance(valor, pd.Timestamp):
            ws.write_datetime(lin, col, valor.to_pydatetime(), self.f["data"])
        elif isinstance(valor, (int, float)) and not isinstance(valor, bool):
            ws.write_number(lin, col, float(valor), self.f[fmt])
        else:
            ws.write_string(lin, col, str(valor), self.f[fmt])

    def tabela(self, ws, lin: int, df: pd.DataFrame, colunas: list[tuple], *, total: dict | None = None,
               filtro: bool = False, congelar: bool = False, estilo_linha=None, col0: int = 0) -> int:
        """colunas: (campo, título, formato, largura). Devolve a próxima linha livre."""
        for j, (_, titulo, _, larg) in enumerate(colunas):
            ws.write_string(lin, col0 + j, titulo, self.f["cab"])
            if larg:
                ws.set_column(col0 + j, col0 + j, larg)
        ini = lin + 1
        for i, r in enumerate(df.itertuples(index=False)):
            reg = r._asdict()
            prefixo = estilo_linha(reg) if estilo_linha else ""
            for j, (campo, _, fmt, _) in enumerate(colunas):
                v = reg.get(campo)
                f_ = fmt
                if fmt == "status":
                    f_ = v if v in ("Bloqueia", "Pendência") else "texto"
                elif prefixo:
                    f_ = prefixo + ("texto" if fmt in ("texto", "quebra", "data") else fmt)
                self.escrever(ws, ini + i, col0 + j, v, f_)
        fim = ini + len(df)
        if filtro and len(df):
            ws.autofilter(lin, col0, fim - 1, col0 + len(colunas) - 1)
        if congelar:
            ws.freeze_panes(lin + 1, 0)
        if total is not None:
            for j, (campo, _, fmt, _) in enumerate(colunas):
                if j == 0:
                    ws.write_string(fim, col0, total.get("_rotulo", "Total"), self.f["t_texto"])
                elif campo in total:
                    self.escrever(ws, fim, col0 + j, total[campo], "t_" + ("int" if fmt == "int" else "dec"))
                else:
                    ws.write_blank(fim, col0 + j, None, self.f["t_texto"])
            fim += 1
        return fim + 1

    def fechar(self):
        self.wb.close()


def _cabecalho(lv: Livro, ws, titulo: str, r) -> int:
    ws.write_string(0, 0, titulo, lv.f["titulo"])
    ws.write_string(1, 0, f"Grupo Atlas · fechamento de {mes_curto(r.mes)} · extração: {r.extracao} · "
                          f"gerado em {r.gerado_em:%d/%m/%Y %H:%M}", lv.f["sub"])
    return 3


def gravar_excel(r, caminho: Path) -> None:
    lv = Livro(caminho)
    conf = r.conferencias
    mes = mes_curto(r.mes)
    ant = mes_curto(str(pd.Period(r.mes, "M") - 1))

    # ------------------------------------------------------------------ 1 Status
    ws = lv.aba("1 Status")
    ws.set_landscape()
    ws.set_paper(9)
    ws.fit_to_pages(1, 1)
    lin = _cabecalho(lv, ws, f"Fechamento {mes}: status do pacote", r)
    ws.merge_range(lin, 0, lin, 3, f"Veredito: {r.veredito.upper()}", lv.f[f"veredito_{r.veredito}"])
    ws.set_row(lin, 28)
    lin += 2
    d = r.dre.set_index("linha")
    pot = r.potencial if r.potencial is not None else pd.DataFrame(columns=["efeito"])
    ebitda_aj = d.loc["EBITDA", "mes_ajustado"]
    resumo = [("EBITDA do mês antes dos ajustes", d.loc["EBITDA", "mes_original"], ""),
              ("Ajustes propostos (A1–A4) no EBITDA", d.loc["EBITDA", "ajustes"], ""),
              ("EBITDA após ajustes propostos (A1–A4)", ebitda_aj, "destaque")]
    resumo += [(f"   Pendência {x.conferencia}: {x.pendencia}", x.efeito, "") for x in pot.itertuples()]
    if len(pot):
        resumo.append(("EBITDA se as pendências medidas se confirmarem", ebitda_aj + pot["efeito"].sum(), "destaque"))
    resumo += [("Pendências de conciliação, em valor absoluto (C3, C4, C5, C7)", r.pendencias_total, ""),
               ("Lançamentos fora do corte ou da competência (C6)", conf["C6"].diferenca, "")]
    for i, (rot, v, est) in enumerate(resumo):
        lv.escrever(ws, lin + i, 0, rot, "negrito" if est else "texto")
        lv.escrever(ws, lin + i, 1, v, "s_int" if est else "int")
    lin += len(resumo) + 1
    ws.write_string(lin, 0, "Conferências", lv.f["secao"])
    tab = pd.DataFrame([dict(codigo=c.codigo, nome=c.nome, status=c.status, diferenca=c.diferenca,
                             responsavel=c.responsavel, prazo=c.prazo if c.status != "OK" else None)
                        for c in conf.values()])
    lin = lv.tabela(ws, lin + 1, tab, [("codigo", "Cód.", "texto", 7), ("nome", "Conferência", "texto", 34),
                                       ("status", "Status", "status", 12), ("diferenca", "Diferença (R$)", "int", 16),
                                       ("responsavel", "Responsável", "texto", 34), ("prazo", "Prazo", "data", 12)])
    ws.write_string(lin, 0, "Arquivos lidos", lv.f["secao"])
    arq = r.entradas.arquivos.assign(hash=r.entradas.arquivos["sha256"].str[:12])
    lin = lv.tabela(ws, lin + 1, arq, [("arquivo", "Arquivo", "texto", None), ("linhas", "Linhas", "int", None),
                                       ("hash", "SHA-256 (início)", "texto", None)])
    for i, rot in enumerate(["Preparado por:", "Revisado por:", "Data da revisão:"]):
        ws.write_string(lin + i * 2, 0, rot, lv.f["negrito"])
        ws.merge_range(lin + i * 2, 1, lin + i * 2, 2, "", lv.f["campo"])
    ws.set_column(0, 0, 40)

    # ------------------------------------------------------------------ 2 Conferências
    ws = lv.aba("2 Conferências")
    lin = _cabecalho(lv, ws, "Conferências automáticas", r)
    tab = pd.DataFrame([dict(codigo=c.codigo, nome=c.nome, regra=c.regra, resultado=c.resultado, diferenca=c.diferenca,
                             tolerancia=c.tolerancia, status=c.status, responsavel=c.responsavel,
                             prazo=c.prazo if c.status != "OK" else None) for c in conf.values()])
    lv.tabela(ws, lin, tab, [("codigo", "Cód.", "texto", 6), ("nome", "Conferência", "quebra", 22),
                             ("regra", "Regra", "quebra", 46), ("resultado", "Resultado", "quebra", 60),
                             ("diferenca", "Diferença (R$)", "dec", 15), ("tolerancia", "Tolerância", "quebra", 22),
                             ("status", "Status", "status", 11), ("responsavel", "Responsável", "quebra", 22),
                             ("prazo", "Prazo", "data", 11)], filtro=True, congelar=True)

    # ------------------------------------------------------------------ 3 Ajustes
    ws = lv.aba("3 Ajustes")
    lin = _cabecalho(lv, ws, "Ajustes propostos (nada foi lançado no ERP)", r)
    ef = r.efeitos
    lin = lv.tabela(ws, lin, ef, [("ajuste", "Ajuste", "texto", None), ("nome", "Descrição", "quebra", None),
                                  ("regra", "Regra de cálculo", "quebra", None), ("valor", "Valor (R$)", "dec", None),
                                  ("efeito_ebitda", "Efeito no EBITDA", "dec", None),
                                  ("efeito_abaixo_ebitda", "Efeito abaixo do EBITDA", "dec", None),
                                  ("aprovado", "Aprovado (S/N)", "texto", None)],
                    total={"valor": ef["valor"].sum(), "efeito_ebitda": ef["efeito_ebitda"].sum(),
                           "efeito_abaixo_ebitda": ef["efeito_abaixo_ebitda"].sum()})
    ws.write_string(lin, 0, "Lançamentos propostos (partida dobrada)", lv.f["secao"])
    linhas = pd.concat([a.linhas for a in r.ajustes.values()], ignore_index=True)
    linhas["descricao"] = linhas["conta"].map(r.entradas.p.contas.set_index("conta")["descricao"])
    tot = {"debito": linhas["debito"].sum(), "credito": linhas["credito"].sum(), "_rotulo": "Total (D = C)"}
    lin = lv.tabela(ws, lin + 1, linhas, [("ajuste", "Ajuste", "texto", None), ("conta", "Conta", "texto", None),
                                          ("descricao", "Descrição da conta", "texto", None),
                                          ("unidade", "Unidade", "texto", None),
                                          ("centro_custo", "C. custo", "texto", None), ("debito", "Débito", "dec", None),
                                          ("credito", "Crédito", "dec", None), ("historico", "Histórico", "texto", None)],
                    total=tot, filtro=True)
    for col, larg in enumerate([10, 34, 52, 16, 16, 16, 16, 90]):
        ws.set_column(col, col, larg)
    ws.write_string(lin, 0, "Memória A1: aging de clientes por faixa", lv.f["secao"])
    mem = r.ajustes["A1"].memoria
    lin = lv.tabela(ws, lin + 1, mem, [("unidade", "Unidade", "texto", None), ("faixa", "Faixa", "texto", None),
                                       ("saldo", "Saldo em aberto", "dec", None), ("percentual", "% PDD", "pct", None),
                                       ("pdd", "PDD calculada", "dec", None)],
                    total={"saldo": mem["saldo"].sum(), "pdd": mem["pdd"].sum()})
    res = r.ajustes["A1"].resumo
    lin = lv.tabela(ws, lin, res, [("unidade", "Unidade", "texto", None), ("pdd_calculada", "PDD calculada", "dec", None),
                                   ("pdd_contabil", "PDD contábil", "dec", None), ("ajuste", "Ajuste A1", "dec", None)],
                    total={c: res[c].sum() for c in ["pdd_calculada", "pdd_contabil", "ajuste"]})
    ws.write_string(lin, 0, "Memória A2: provisões de férias e 13º por unidade", lv.f["secao"])
    m2 = r.ajustes["A2"].memoria
    cols2 = [("unidade", "Unidade", "texto", None), ("ferias", "Férias calculadas", "dec", None),
             ("ferias_contabil", "Férias contábil", "dec", None), ("decimo", "13º calculado", "dec", None),
             ("decimo_contabil", "13º contábil", "dec", None), ("ajuste", "Diferença calculada", "dec", None),
             ("ajuste_lancado", "Ajuste A2 lançado", "dec", None), ("observacao", "Observação", "texto", None)]
    lv.tabela(ws, lin + 1, m2, cols2, total={c: m2[c].sum() for c, _, f_, _ in cols2[1:] if f_ == "dec"})

    # ------------------------------------------------------------------ 4 Balancete ajustado
    ws = lv.aba("4 Balancete ajustado")
    lin = _cabecalho(lv, ws, "Balancete ajustado por conta × unidade (positivo = devedor)", r)
    b = r.balancete_ajustado.copy()
    classes = [("1", "Ativo"), ("2", "Passivo e patrimônio líquido"), ("3", "Resultado")]
    partes = []
    cols_v = ["saldo_inicial", "debitos", "creditos", "saldo_final", "ajustes", "saldo_ajustado"]
    for cod, nome in classes:
        sel = b[b["conta"].str[0].isin([cod] if cod != "3" else list("34567"))]
        partes.append(sel)
        partes.append(pd.DataFrame([dict(conta=f"Subtotal {nome}", **{c: round(sel[c].sum(), 2) + 0.0
                                                                       for c in cols_v})]))
    tab = pd.concat(partes, ignore_index=True)
    lv.tabela(ws, lin, tab, [("conta", "Conta", "texto", 26), ("descricao", "Descrição", "texto", 40),
                             ("unidade", "Unidade", "texto", 9), ("saldo_inicial", "Saldo inicial", "dec", 17),
                             ("debitos", "Débitos", "dec", 17), ("creditos", "Créditos", "dec", 17),
                             ("saldo_final", "Saldo final ERP", "dec", 17), ("ajustes", "Ajustes", "dec", 15),
                             ("saldo_ajustado", "Saldo ajustado", "dec", 17)],
              total={"_rotulo": "Total (deve ser zero)", **{c: round(b[c].sum(), 2) + 0.0 for c in cols_v}},
              filtro=True,
              congelar=True, estilo_linha=lambda reg: "s_" if str(reg["conta"]).startswith("Subtotal") else "")

    # ------------------------------------------------------------------ 5 DRE
    ws = lv.aba("5 DRE")
    lin = _cabecalho(lv, ws, f"DRE consolidada: {mes} e acumulado do ano (R$)", r)
    tab = r.dre.copy()
    lv.tabela(ws, lin, tab, [("linha", "Linha", "texto", 30), ("mes_original", f"{mes} original", "int", 15),
                             ("ajustes", "Ajustes A1–A4", "int", 13),
                             ("mes_ajustado", f"{mes} após ajustes", "int", 15),
                             ("mes_anterior", ant, "int", 15), ("delta_mes_anterior", f"Δ vs {ant}", "int", 14),
                             ("acumulado_original", "Acumulado original", "int", 17),
                             ("acumulado_ajustado", "Acumulado ajustado", "int", 17)], congelar=True,
              estilo_linha=lambda reg: {"subtotal": "s_", "total": "t_"}.get(reg["tipo"], ""))
    ws.write_string(lin + len(tab) + 2, 0, "Receitas positivas, custos e despesas negativos; Δ positivo = favorável. "
                                           "Após ajustes = só os ajustes propostos A1–A4; as pendências ainda abertas "
                                           "estão na aba 1. O orçado fica no painel de análise (P2).", lv.f["sub"])

    # ------------------------------------------------------------------ 6 Conciliação bancária
    ws = lv.aba("6 Conciliação bancária")
    lin = _cabecalho(lv, ws, "Conciliação bancária por banco (R$, com centavos)", r)
    blocos = conf["C3"].extra["blocos"]
    for blk in blocos.itertuples():
        ws.write_string(lin, 0, f"{blk.banco} · conta contábil {blk.conta}", lv.f["secao"])
        itens = [("Saldo do extrato em " + data(r.entradas.p.data_base), blk.saldo_extrato, "dec"),
                 ("(+) Depósitos em trânsito", blk.depositos_transito, "dec"),
                 ("(−) Cheques e pagamentos não compensados", -blk.cheques_nao_compensados, "dec"),
                 ("= Saldo bancário ajustado", blk.saldo_banco_ajustado, "t_dec"),
                 ("Saldo contábil", blk.saldo_contabil, "dec"),
                 ("(+) Créditos do banco não contabilizados", blk.creditos_nao_contabilizados, "dec"),
                 ("(−) Débitos do banco não contabilizados", -blk.debitos_nao_contabilizados, "dec"),
                 ("= Saldo contábil + itens do banco", blk.saldo_contabil_ajustado, "t_dec"),
                 ("Diferença não explicada", blk.diferenca, "s_dec")]
        if abs(blk.sem_lancamento_proposto) >= 0.01:
            itens += [("(−) Itens não identificados, sem lançamento proposto", -blk.sem_lancamento_proposto, "dec"),
                      ("= Saldo no balancete ajustado (aba 4)", blk.saldo_balancete_ajustado, "t_dec")]
        for i, (rot, v, fmt) in enumerate(itens):
            lv.escrever(ws, lin + 1 + i, 0, rot, "t_texto" if fmt.startswith("t_") else
                        ("negrito" if fmt.startswith("s_") else "texto"))
            lv.escrever(ws, lin + 1 + i, 1, v, fmt)
        lin += len(itens) + 2
    ws.set_column(0, 0, 44)
    ws.set_column(1, 1, 18)
    ws.write_string(lin, 0, "Itens em aberto e não contabilizados", lv.f["secao"])
    it = conf["C3"].detalhe
    lv.tabela(ws, lin + 1, it, [("banco", "Banco", "texto", None), ("origem", "Lado", "texto", 14),
                                ("tipo", "Tipo", "texto", 32), ("data", "Data", "data", 11),
                                ("documento", "Documento", "texto", 14), ("descricao", "Descrição", "texto", 46),
                                ("valor", "Valor", "dec", 15), ("dias_uteis", "Dias úteis em aberto", "int", 11),
                                ("situacao", "Situação", "status", 20)], filtro=True)

    # ------------------------------------------------------------------ 7 Subledgers
    ws = lv.aba("7 Subledgers")
    lin = _cabecalho(lv, ws, "Subledgers × razão (R$, com centavos)", r)
    det4 = conf["C4"].detalhe
    lin = lv.tabela(ws, lin, det4, [("subledger_nome", "Subledger", "texto", 26), ("unidade", "Unidade", "texto", 10),
                                    ("conta", "Conta", "texto", 10), ("subledger", "Subledger", "dec", 17),
                                    ("razao", "Razão", "dec", 17), ("diferenca", "Diferença", "dec", 15),
                                    ("status", "Status", "status", 11)], filtro=True, congelar=True)
    ws.write_string(lin, 0, "Folha × razão por centro de custo", lv.f["secao"])
    det5 = conf["C5"].detalhe
    lin = lv.tabela(ws, lin + 1, det5, [("centro_custo", "Centro de custo", "texto", None),
                                        ("folha", "Folha", "dec", None), ("razao", "Razão", "dec", None),
                                        ("diferenca", "Diferença", "dec", None), ("status", "Status", "status", None)],
                    total={c: det5[c].sum() for c in ["folha", "razao", "diferenca"]})
    ws.write_string(lin, 0, "Aging de clientes por faixa (base da PDD)", lv.f["secao"])
    piv = r.ajustes["A1"].memoria.pivot_table(index="unidade", columns="faixa", values="saldo", aggfunc="sum",
                                              observed=False)
    faixas = list(r.entradas.p.pdd.sort_values("dias_de")["faixa"])
    piv = piv[faixas]
    piv.columns = [f"f{i}" for i in range(len(faixas))]     # nomes seguros para o itertuples
    piv["total"] = piv.sum(axis=1)
    piv = piv.reset_index()
    cols = [("unidade", "Unidade", "texto", None)] + [(f"f{i}", f_, "dec", None) for i, f_ in enumerate(faixas)] + \
           [("total", "Total", "dec", None)]
    lin = lv.tabela(ws, lin + 1, piv, cols, total={c: piv[c].sum() for c, *_ in cols[1:]})
    ws.write_string(lin, 0, "Fornecedores: o saldo credor do razão aparece positivo para comparar com o aging.",
                    lv.f["sub"])

    # ------------------------------------------------------------------ 8 Exceções
    ws = lv.aba("8 Exceções")
    lin = _cabecalho(lv, ws, "Exceções linha a linha (C6, C7 e C8): preencher justificativa e ação", r)
    ex = []
    for x in conf["C6"].detalhe.itertuples():
        ex.append(dict(conferencia="C6", referencia=f"Lanç. {x.id_lanc}", data=x.data_lanc, conta=x.conta,
                       unidade=x.unidade, centro_custo=x.centro_custo, descricao=x.historico, valor=x.valor,
                       motivo=x.motivo, acao_sugerida=x.acao))
    acao7 = {"Natureza invertida": "Reclassificar o saldo para a conta de natureza correta (ex.: adiantamentos)",
             "Transitória com saldo": "Identificar o pagamento e classificar na conta definitiva",
             "Despesa recorrente sem lançamento": "Provisionado pelo A3; cobrar a nota do fornecedor"}
    for x in conf["C7"].detalhe.itertuples():
        ex.append(dict(conferencia="C7", referencia=x.tipo, data=None, conta=x.conta, unidade=x.unidade,
                       centro_custo=x.centro_custo, descricao=x.descricao, valor=x.valor, motivo=x.tratamento,
                       acao_sugerida=acao7.get(x.tipo, "")))
    for x in conf["C8"].detalhe.itertuples():
        ex.append(dict(conferencia="C8", referencia="Variação atípica", data=None, conta=x.conta, unidade=x.unidade,
                       centro_custo="", descricao=f"{x.descricao}: {ant} {x.anterior:,.0f} → {mes} {x.atual:,.0f}"
                       .replace(",", "."), valor=x.delta, motivo=f"{x.variacao:+.0%} contra {ant}",
                       acao_sugerida="Justificar com o gestor do centro de custo"))
    ex = pd.DataFrame(ex, columns=["conferencia", "referencia", "data", "conta", "unidade", "centro_custo", "descricao",
                                   "valor", "motivo", "acao_sugerida"]).assign(justificativa="", acao="")
    lv.tabela(ws, lin, ex, [("conferencia", "Conf.", "texto", 6), ("referencia", "Referência", "texto", 30),
                            ("data", "Data", "data", 11), ("conta", "Conta", "texto", 9),
                            ("unidade", "Unidade", "texto", 8), ("centro_custo", "C. custo", "texto", 11),
                            ("descricao", "Descrição", "texto", 58), ("valor", "Valor", "dec", 14),
                            ("motivo", "Motivo", "texto", 44), ("acao_sugerida", "Ação sugerida", "texto", 56),
                            ("justificativa", "Justificativa", "texto", 30), ("acao", "Ação tomada", "texto", 24)],
              filtro=True, congelar=True)

    # ------------------------------------------------------------------ 9 Log
    ws = lv.aba("9 Log")
    lin = _cabecalho(lv, ws, "Log: parâmetros, tolerâncias, versão das entradas e regras aplicadas", r)
    p = r.entradas.p
    ger = pd.DataFrame([(k, str(v)) for k, v in p.geral.items()], columns=["parametro", "valor"])
    lin = lv.tabela(ws, lin, ger, [("parametro", "Parâmetro", "texto", 30), ("valor", "Valor", "texto", 70)])
    tol = p.tolerancias.reset_index()
    lin = lv.tabela(ws, lin, tol, [("conferencia", "Conferência", "texto", None),
                                   ("tolerancia", "Tolerância (R$)", "dec", None),
                                   ("responsavel", "Responsável", "texto", None),
                                   ("prazo_dias_uteis", "Prazo (dias úteis após o corte)", "int", None)])
    reg = pd.DataFrame([(c, n, str(v)) for (c, n), v in p.regras.items()], columns=["conferencia", "parametro", "valor"])
    lin = lv.tabela(ws, lin, reg, [("conferencia", "Conferência", "texto", None), ("parametro", "Parâmetro", "texto", None),
                                   ("valor", "Valor", "texto", None)])
    lin = lv.tabela(ws, lin, r.entradas.arquivos, [("arquivo", "Arquivo", "texto", None),
                                                   ("linhas", "Linhas", "int", None),
                                                   ("sha256", "SHA-256", "texto", None)])
    regras = pd.DataFrame([dict(codigo=c.codigo, regra=c.regra) for c in conf.values()] +
                          [dict(codigo=a.codigo, regra=a.regra) for a in r.ajustes.values()])
    lv.tabela(ws, lin, regras, [("codigo", "Código", "texto", None), ("regra", "Regra aplicada", "texto", None)])
    lv.fechar()
