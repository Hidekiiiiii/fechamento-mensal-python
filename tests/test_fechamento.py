"""Testes contra o gabarito: os 12 problemas plantados com valores exatos, nenhuma armadilha marcada e o pacote íntegro.

Rodar a partir da pasta do projeto:  python -m unittest discover tests
"""
from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path

import openpyxl
import pandas as pd

from fechamento import gravar, processar
from fechamento.demonstrativos import para_proximo_mes
from fechamento.leitura import ler_csv

RAIZ = Path(__file__).resolve().parents[1]
GAB = json.loads((RAIZ / "gerador" / "gabarito.json").read_text(encoding="utf-8"))
ENT = RAIZ / "entradas"


def rodar(n: int):
    return processar(ENT / "2025-10" / f"extracao_{n}", ENT / "2025-09" / "balancete_ajustado_2025-09.csv",
                     ENT / "parametros.xlsx")


class Extracao1(unittest.TestCase):
    """Primeira extração: erros de extração precisam bloquear o pacote."""

    @classmethod
    def setUpClass(cls):
        cls.r = rodar(1)
        cls.g = GAB["extracao_1"]

    def test_veredito(self):
        self.assertEqual(self.r.veredito, self.g["veredito"])

    def test_c1_erro_aritmetico_e_saldo_inicial(self):
        c1 = self.r.conferencias["C1"]
        self.assertEqual(c1.status, "Bloqueia")
        ar = c1.extra["erro_aritmetico"]
        self.assertEqual([(x.conta, x.unidade, round(x.diferenca, 2)) for x in ar.itertuples()],
                         [(d["conta"], d["unidade"], d["diferenca"]) for d in self.g["C1"]["erro_aritmetico"]])
        si = c1.extra["saldo_inicial"].sort_values("conta")
        self.assertEqual([(x.conta, x.unidade, round(x.diferenca, 2)) for x in si.itertuples()],
                         [(d["conta"], d["unidade"], d["diferenca"]) for d in self.g["C1"]["saldo_inicial"]])

    def test_c2_lancamentos_fora_do_razao(self):
        c2 = self.r.conferencias["C2"]
        self.assertEqual(c2.status, "Bloqueia")
        self.assertEqual(len(c2.detalhe), self.g["C2"]["linhas"])
        self.assertAlmostEqual(c2.extra["debitos_faltantes"], self.g["C2"]["debitos_faltantes"], places=2)
        self.assertAlmostEqual(c2.extra["creditos_faltantes"], self.g["C2"]["creditos_faltantes"], places=2)

    def test_email_avisa_bloqueio(self):
        self.assertIn("não apresentável", self.r.email[0])
        self.assertIn("reextrair", self.r.email[0])


class Extracao2(unittest.TestCase):
    """Reextração: C1 e C2 passam; as 12 ocorrências aparecem com valores exatos."""

    @classmethod
    def setUpClass(cls):
        cls.r = rodar(2)
        cls.g = GAB["extracao_2"]
        cls.c = cls.r.conferencias

    def test_veredito_e_status(self):
        self.assertEqual(self.r.veredito, self.g["veredito"])
        for cod in ["C1", "C2", "C3", "C4", "C5", "C6", "C7", "C8"]:
            self.assertEqual(self.c[cod].status, self.g[cod]["status"], cod)

    def test_c3_bancos(self):
        x, g = self.c["C3"].extra, self.g["C3"]
        for k in ["pendencia", "cheque_antigo", "nao_identificado", "transito_no_prazo", "tarifas", "rendimentos",
                  "diferenca_nao_explicada"]:
            self.assertAlmostEqual(x[k], g[k], places=2, msg=k)

    def test_c3_armadilha_depositos_em_transito_nao_sao_pendencia(self):
        it = self.c["C3"].detalhe
        transito = it[it["tipo"] == "Depósito em trânsito"]
        self.assertEqual(len(transito), 2)
        self.assertTrue((transito["situacao"] == "No prazo").all())
        self.assertTrue((transito["dias_uteis"] <= 2).all())

    def test_c4_clientes(self):
        c4, g = self.c["C4"], self.g["C4"]
        self.assertAlmostEqual(c4.diferenca, g["diferenca"], places=2)
        self.assertEqual(c4.extra["unidades"], g["unidades"])
        self.assertAlmostEqual(c4.extra["razao_cr"], g["razao_cr"], places=2)
        self.assertAlmostEqual(c4.extra["aging_cr"], g["aging_cr"], places=2)

    def test_c5_folha(self):
        c5, g = self.c["C5"], self.g["C5"]
        self.assertEqual(c5.extra["centros"], g["centros"])
        linha = c5.detalhe.set_index("centro_custo").loc["SERV-CDSE"]
        self.assertAlmostEqual(linha["folha"], g["folha"], places=2)
        self.assertAlmostEqual(linha["razao"], g["razao"], places=2)
        self.assertAlmostEqual(c5.diferenca, g["diferenca"], places=2)

    def test_c6_corte(self):
        c6, g = self.c["C6"], self.g["C6"]
        self.assertEqual(c6.extra["lancamentos"], g["lancamentos"])
        self.assertAlmostEqual(c6.diferenca, g["valor"], places=2)
        self.assertAlmostEqual(c6.extra["competencia"], g["competencia"], places=2)
        self.assertAlmostEqual(c6.extra["corte"], g["corte"], places=2)

    def test_c6_armadilha_provisao_revertida_nao_e_marcada(self):
        hist = " ".join(self.c["C6"].detalhe["historico"])
        self.assertNotIn("Consultoria de processos", hist)

    def test_c6_armadilha_estornos_nao_sao_marcados(self):
        self.assertFalse(self.c["C6"].detalhe["historico"].str.startswith(("Estorno", "Relançamento")).any())

    def test_c7_saldos(self):
        c7, g = self.c["C7"], self.g["C7"]
        self.assertAlmostEqual(c7.diferenca, g["valor"], places=2)
        pend = c7.detalhe[c7.detalhe["tratamento"] == "Pendência"]
        self.assertEqual(sorted((x.conta, x.unidade, round(x.valor, 2)) for x in pend.itertuples()),
                         sorted((d["conta"], d["unidade"], d["saldo"]) for d in g["itens"]))
        rec = c7.detalhe[c7.detalhe["tipo"] == "Despesa recorrente sem lançamento"]
        self.assertAlmostEqual(rec["valor"].sum(), g["recorrentes_faltantes"], places=2)

    def test_c8_so_o_frete(self):
        d = self.c["C8"].detalhe
        g = self.g["C8"]["itens"]
        self.assertEqual([(x.conta, x.unidade, x.anterior, x.atual) for x in d.itertuples()],
                         [(i["conta"], i["unidade"], i["setembro"], i["outubro"]) for i in g])

    def test_ajustes(self):
        ef = self.r.efeitos.set_index("ajuste")
        for cod in ["A1", "A2", "A3"]:
            self.assertAlmostEqual(-ef.loc[cod, "efeito_ebitda"], self.g[cod], places=2, msg=cod)
        a4 = self.r.ajustes["A4"].linhas
        self.assertAlmostEqual(a4.loc[a4["conta"] == "7.1.02", "debito"].sum(), self.g["A4"]["tarifas"], places=2)
        self.assertAlmostEqual(a4.loc[a4["conta"] == "7.1.01", "credito"].sum(), self.g["A4"]["rendimentos"], places=2)
        res = self.r.ajustes["A1"].resumo
        self.assertAlmostEqual(res["pdd_calculada"].sum(), self.g["pdd_calculada"], places=2)
        self.assertAlmostEqual(res["pdd_contabil"].sum(), self.g["pdd_contabil"], places=2)

    def test_ajustes_em_partida_dobrada(self):
        for a in self.r.ajustes.values():
            self.assertAlmostEqual(a.linhas["debito"].sum(), a.linhas["credito"].sum(), places=2, msg=a.codigo)

    def test_ebitda(self):
        d = self.r.dre.set_index("linha")
        self.assertAlmostEqual(d.loc["EBITDA", "mes_original"], self.g["ebitda_antes"], places=2)
        self.assertAlmostEqual(d.loc["EBITDA", "mes_ajustado"], self.g["ebitda_depois"], places=2)

    def test_pendencias_total(self):
        self.assertAlmostEqual(self.r.pendencias_total, self.g["pendencias_total"], places=2)

    def test_balancete_ajustado_fecha(self):
        b = self.r.balancete_ajustado
        self.assertAlmostEqual(b["saldo_final"].sum(), 0, places=2)
        self.assertAlmostEqual(b["saldo_ajustado"].sum(), 0, places=2)

    def test_email(self):
        texto = "\n".join(self.r.email)
        self.assertEqual(len(self.r.email), 5)
        for trecho in ["com ressalvas", "R$ 1,42 mi", "R$ 0,96 mi", "−R$ 460,3 mil", "R$ 1,74 mi em set/25",
                       "C4, C5 e C6 se confirmarem, vai a R$ 1,07 mi", "R$ 520,6 mil",
                       "R$ 286 mil de competência futura a estornar e relançar", "R$ 41,5 mil lançados depois do corte",
                       "R$ 92 mil", "9 dias úteis", "+48%"]:
            self.assertIn(trecho, texto)

    def test_efeito_potencial_das_pendencias(self):
        pot = self.r.potencial.set_index("conferencia")["efeito"]
        self.assertEqual(pot.to_dict(), {"C4": -156_000.00, "C5": -22_400.00, "C6": 286_000.00})

    def test_c6_acao_por_motivo(self):
        d = self.c["C6"].detalhe
        self.assertEqual(int(d["futura"].sum()), 3)
        self.assertTrue(d.loc[d["futura"], "acao"].str.startswith("Estornar").all())
        self.assertTrue(d.loc[~d["futura"], "acao"].str.startswith("Competência correta").all())

    def test_banco_beta_bate_com_o_balancete_ajustado(self):
        blocos = self.c["C3"].extra["blocos"].set_index("conta")
        b = self.r.balancete_ajustado.groupby("conta")["saldo_ajustado"].sum()
        for conta, blk in blocos.iterrows():
            self.assertAlmostEqual(blk["saldo_balancete_ajustado"], b[conta], places=2, msg=conta)

    def test_memoria_a2_soma_o_lancado(self):
        m = self.r.ajustes["A2"].memoria
        self.assertAlmostEqual(m["ajuste_lancado"].sum(), self.g["A2"], places=2)
        self.assertAlmostEqual(self.r.ajustes["A2"].valor, self.g["A2"], places=2)

    def test_aging_por_faixa_preenchido_no_excel(self):
        with tempfile.TemporaryDirectory() as tmp:
            ws = openpyxl.load_workbook(gravar(self.r, Path(tmp))["pacote"])["7 Subledgers"]
            linhas = [r for r in ws.iter_rows(values_only=True) if r and r[0] in ("DIST", "DIG", "SERV")
                      and isinstance(r[1], (int, float)) and len([v for v in r if v is not None]) >= 6]
            self.assertEqual(len(linhas), 3)
            for r in linhas:
                self.assertAlmostEqual(sum(r[1:6]), r[6], places=2)

    def test_pacote_excel_e_encadeamento(self):
        with tempfile.TemporaryDirectory() as tmp:
            arq = gravar(self.r, Path(tmp))
            wb = openpyxl.load_workbook(arq["pacote"])
            self.assertEqual(wb.sheetnames, ["1 Status", "2 Conferências", "3 Ajustes", "4 Balancete ajustado", "5 DRE",
                                             "6 Conciliação bancária", "7 Subledgers", "8 Exceções", "9 Log"])
            erros = [(ws.title, c.coordinate) for ws in wb for row in ws.iter_rows() for c in row
                     if isinstance(c.value, str) and c.value.startswith("#")]
            self.assertEqual(erros, [])
            self.assertEqual(wb["1 Status"]["A4"].value, "Veredito: COM RESSALVAS")
            # o balancete ajustado sai no layout do pacote anterior e vira o saldo inicial do mês seguinte
            novo = ler_csv(arq["balancete"])
            anterior = ler_csv(ENT / "2025-09" / "balancete_ajustado_2025-09.csv")
            self.assertEqual(list(novo.columns), list(anterior.columns))
            self.assertAlmostEqual(novo["saldo_final"].sum(), 0, places=2)
            pd.testing.assert_series_equal(novo["saldo_final"],
                                           para_proximo_mes(self.r.balancete_ajustado)["saldo_final"],
                                           check_names=False)


if __name__ == "__main__":
    unittest.main()
