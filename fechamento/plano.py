"""Contas-chave do plano de contas do Grupo Atlas e estrutura da DRE."""

CLIENTES = "1.1.03"
PDD = "1.1.04"
FORNECEDORES = "2.1.01"
PROV_FERIAS = "2.1.05"
PROV_13 = "2.1.06"
PROV_PAGAR = "2.1.08"
DESP_PDD = "5.1.05"
REC_FIN = "7.1.01"
DESP_FIN = "7.1.02"

GRUPOS_EBITDA = ("Receita Bruta", "Deduções", "Custo", "Despesas Comerciais", "Despesas com Pessoal",
                 "Despesas Administrativas")

# linha da DRE, grupos de contas somados, tipo de linha
ESTRUTURA_DRE = [
    ("Receita bruta", ("Receita Bruta",), "linha"),
    ("Deduções", ("Deduções",), "linha"),
    ("Receita líquida", ("Receita Bruta", "Deduções"), "subtotal"),
    ("Custo", ("Custo",), "linha"),
    ("Lucro bruto", ("Receita Bruta", "Deduções", "Custo"), "subtotal"),
    ("Despesas comerciais", ("Despesas Comerciais",), "linha"),
    ("Despesas com pessoal", ("Despesas com Pessoal",), "linha"),
    ("Despesas administrativas", ("Despesas Administrativas",), "linha"),
    ("EBITDA", GRUPOS_EBITDA, "total"),
    ("Depreciação e amortização", ("D&A",), "linha"),
    ("EBIT", GRUPOS_EBITDA + ("D&A",), "subtotal"),
    ("Resultado financeiro", ("Resultado Financeiro",), "linha"),
    ("Resultado antes do IR (LAIR)", GRUPOS_EBITDA + ("D&A", "Resultado Financeiro"), "total"),
]
