"""Leitura das exportações do ERP, do pacote do mês anterior e dos parâmetros do controller."""
from __future__ import annotations

import hashlib
from dataclasses import dataclass
from pathlib import Path

import pandas as pd

# arquivo esperado na pasta da extração: chave interna -> prefixo do nome
ARQUIVOS = {"balancete": "balancete", "razao": "razao", "extrato": "extrato_bancario", "aging_cr": "aging_receber",
            "aging_cp": "aging_pagar", "folha": "folha"}
DATAS = {"razao": ["data_lanc"], "extrato": ["data"], "aging_cr": ["emissao", "vencimento"],
         "aging_cp": ["emissao", "vencimento"], "folha": ["data_admissao", "data_desligamento"]}
TEXTO = ["conta", "unidade", "centro_custo", "documento", "data_competencia", "titulo", "conta_contabil", "historico",
         "descricao", "rubrica", "situacao", "banco", "agencia_conta", "usuario", "origem", "cliente", "fornecedor"]
OBRIGATORIAS = {
    "balancete": ["conta", "unidade", "saldo_inicial", "debitos", "creditos", "saldo_final"],
    "razao": ["id_lanc", "data_lanc", "data_competencia", "conta", "unidade", "centro_custo", "documento", "debito",
              "credito", "origem"],
    "extrato": ["banco", "data", "documento", "descricao", "valor", "saldo"],
    "aging_cr": ["titulo", "unidade", "vencimento", "valor_aberto", "situacao"],
    "aging_cp": ["titulo", "unidade", "vencimento", "valor_aberto", "conta_contabil"],
    "folha": ["matricula", "centro_custo", "rubrica", "valor", "data_admissao", "data_desligamento",
              "dias_ferias_adquiridos"],
}


def ler_csv(caminho: Path, datas: list[str] = ()) -> pd.DataFrame:
    """CSV no padrão de exportação do ERP: separador ';', decimal ',', datas dd/mm/aaaa."""
    df = pd.read_csv(caminho, sep=";", decimal=",", encoding="utf-8-sig", dtype={c: str for c in TEXTO})
    for c in df.columns.intersection(TEXTO):
        df[c] = df[c].fillna("").str.strip()
    for c in datas:
        df[c] = pd.to_datetime(df[c], format="%d/%m/%Y", errors="coerce")
    return df


def sha256(caminho: Path) -> str:
    return hashlib.sha256(caminho.read_bytes()).hexdigest()


@dataclass
class Parametros:
    geral: dict
    unidades: pd.DataFrame
    contas: pd.DataFrame
    pdd: pd.DataFrame
    recorrentes: pd.DataFrame
    bancos: pd.DataFrame
    tolerancias: pd.DataFrame
    regras: dict

    @property
    def mes(self) -> str:
        return str(self.geral["mes"])

    @property
    def data_base(self) -> pd.Timestamp:
        return pd.to_datetime(self.geral["data_base"], dayfirst=True)

    @property
    def data_corte(self) -> pd.Timestamp:
        return pd.to_datetime(self.geral["data_corte"], dayfirst=True)

    @property
    def encargos(self) -> float:
        return float(self.geral["taxa_encargos"])

    def tol(self, conf: str) -> float:
        return float(self.tolerancias.loc[conf, "tolerancia"])

    def responsavel(self, conf: str) -> str:
        return str(self.tolerancias.loc[conf, "responsavel"])

    def prazo(self, conf: str) -> pd.Timestamp:
        return self.data_corte + pd.offsets.BDay(int(self.tolerancias.loc[conf, "prazo_dias_uteis"]))

    def regra(self, conf: str, nome: str):
        return self.regras[(conf, nome)]

    def natureza(self) -> pd.Series:
        return self.contas.set_index("conta")["natureza"]

    def tipo(self) -> pd.Series:
        return self.contas.set_index("conta")["tipo"]

    def grupo(self) -> pd.Series:
        return self.contas.set_index("conta")["grupo"]


def ler_parametros(caminho: Path) -> Parametros:
    x = pd.read_excel(caminho, sheet_name=None)
    for aba in x.values():
        for c in aba.columns.intersection(["conta", "conta_contabil", "centro_custo", "unidade"]):
            aba[c] = aba[c].astype(str).str.strip()
    geral = dict(zip(x["geral"]["parametro"], x["geral"]["valor"]))
    regras = {(r.conferencia, r.parametro): r.valor for r in x["regras"].itertuples()}
    return Parametros(geral=geral, unidades=x["unidades"].set_index("unidade"), contas=x["contas"], pdd=x["pdd"],
                      recorrentes=x["recorrentes"], bancos=x["bancos"],
                      tolerancias=x["tolerancias"].set_index("conferencia"), regras=regras)


@dataclass
class Entradas:
    pasta: Path
    balancete: pd.DataFrame
    razao: pd.DataFrame
    extrato: pd.DataFrame
    aging_cr: pd.DataFrame
    aging_cp: pd.DataFrame
    folha: pd.DataFrame
    anterior: pd.DataFrame
    p: Parametros
    arquivos: pd.DataFrame   # arquivo, linhas, sha256 (versão exata das entradas usadas)


def ler(pasta: Path, anterior: Path, parametros: Path) -> Entradas:
    p = ler_parametros(parametros)
    dfs, info = {}, []
    for chave, prefixo in ARQUIVOS.items():
        caminho = pasta / f"{prefixo}_{p.mes}.csv"
        if not caminho.exists():
            raise FileNotFoundError(f"Arquivo esperado não encontrado: {caminho}")
        df = ler_csv(caminho, DATAS.get(chave, []))
        faltam = set(OBRIGATORIAS[chave]) - set(df.columns)
        if faltam:
            raise ValueError(f"{caminho.name}: faltam as colunas {sorted(faltam)}")
        dfs[chave] = df
        info.append((caminho.name, len(df), sha256(caminho)))
    ant = ler_csv(anterior)
    info += [(anterior.name, len(ant), sha256(anterior)), (parametros.name, None, sha256(parametros))]
    arquivos = pd.DataFrame(info, columns=["arquivo", "linhas", "sha256"])
    return Entradas(pasta=pasta, anterior=ant, p=p, arquivos=arquivos, **dfs)
