"""
modelador.py  -  executar depois do extrator-ftp-SIH-SUS.py e do extrator_ibge.py

Entrada (na mesma pasta):
  marcapasso_bruto.csv   (extrator-ftp.py: AIHs do procedimento 0406010650 no SIH/SUS)
  ibge_populacao.csv     (extrator_ibge.py: população por ano, região e grupo, projeção oficial do IBGE)

Saída, na pasta resultados/:
  resultados.xlsx          uma pasta de trabalho contendo todas as planilhas modeladas.
  gretl/                   CSVs e script .inp para repetir o Prais-Winsten no gretl
  graficos/                graficos gerados em PNGs
  prais_winsten.csv        a mesma tabela das abas PW

As duas bases usam o MESMO conjunto de AIHs e o MESMO ano (data de internação), então o total do Brasil é idêntico.
Uso:  py -3.12 modelador.py
Precisa de:  pip install pandas numpy scipy matplotlib openpyxl
"""
import json
import os
import sys
import urllib.request

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XlImage
from openpyxl.styles import Alignment, Font
from scipy import stats

# ------------------------------ CONFIGURAÇÃO ------------------------------
BRUTO = "marcapasso_bruto.csv"
POPULACAO = "ibge_populacao.csv"
PASTA_SAIDA = "resultados"
NOME_PASTA = "resultados.xlsx"
ANO_INI, ANO_FIM = 2016, 2025 # filtro das extrações que ocorrem de 2015-2026
CAMPO_ANO = "DT_INTER"        # ano pela data de internação (critério validado contra o TABNET)
IDADE_CORTE = 60
BASE_PRECOS = "media2025"     # "media2025" = R$ de preços médios de 2025 | "dez2025" = R$ de dez/2025
INCLUIR_EXTRAS = False        # False = modelo principal + sensibilidade do plano; True = todos os 7 cenários
ALFA = 0.05
USAR_BCB = True               # tenta baixar o IPCA mensal do Banco Central; se falhar, usa a série embutida
FONTE_POP = ("IBGE, Projeções da População do Brasil e Unidades da Federação, Revisão 2024 (base no Censo 2022), "
             "população por sexo e idade simples em 1º de julho; UFs somadas nas Grandes Regiões")
# ---------------------------------------------------------------------------

BASE = os.path.dirname(os.path.abspath(__file__))
SAIDA = os.path.join(BASE, PASTA_SAIDA)
REG5 = ["Região Norte", "Região Nordeste", "Região Sudeste", "Região Sul", "Região Centro-Oeste"]
REGIOES = REG5 + ["Brasil (Total)"]
ORDEM_REG = {r: i for i, r in enumerate(REGIOES)}
SLUG_REG = {"Região Norte": "norte", "Região Nordeste": "nordeste", "Região Sudeste": "sudeste",
            "Região Sul": "sul", "Região Centro-Oeste": "centrooeste", "Brasil (Total)": "brasil"}
SLUG_GRP = {"Geral": "geral", "60+": "g60"}
GRUPOS = ["Geral", "60+"]
UF2REG = {}
for _faixa, _nome in [((11, 18), REG5[0]), ((21, 30), REG5[1]), ((31, 36), REG5[2]),
                      ((41, 44), REG5[3]), ((50, 54), REG5[4])]:
    for _uf in range(*_faixa):
        UF2REG[_uf] = _nome

# IPCA oficial (variação dezembro/dezembro, %). Fonte: IBGE. Usado só para CONFERIR a série mensal.
IPCA_ANUAL_OFICIAL = {2016: 6.29, 2017: 2.95, 2018: 3.75, 2019: 4.31, 2020: 4.52,
                      2021: 10.06, 2022: 5.79, 2023: 4.62, 2024: 4.83, 2025: 4.26}

# IPCA mensal (% a.m.), 2016-2025. Reserva: o script prefere baixar a série do Banco Central.
IPCA_MENSAL_EMBUTIDO = {
    2016: [1.27, 0.90, 0.43, 0.61, 0.78, 0.35, 0.52, 0.44, 0.08, 0.26, 0.18, 0.30],
    2017: [0.38, 0.33, 0.25, 0.14, 0.31, -0.23, 0.24, 0.19, 0.16, 0.42, 0.28, 0.44],
    2018: [0.29, 0.32, 0.09, 0.22, 0.40, 1.26, 0.33, -0.09, 0.48, 0.45, -0.21, 0.15],
    2019: [0.32, 0.43, 0.75, 0.57, 0.13, 0.01, 0.19, 0.11, -0.04, 0.10, 0.51, 1.15],
    2020: [0.21, 0.25, 0.07, -0.31, -0.38, 0.26, 0.36, 0.24, 0.64, 0.86, 0.89, 1.35],
    2021: [0.25, 0.86, 0.93, 0.31, 0.83, 0.53, 0.96, 0.87, 1.16, 1.25, 0.95, 0.73],
    2022: [0.54, 1.01, 1.62, 1.06, 0.47, 0.67, -0.68, -0.36, -0.29, 0.59, 0.41, 0.62],
    2023: [0.53, 0.84, 0.71, 0.61, 0.23, -0.08, 0.12, 0.23, 0.26, 0.24, 0.28, 0.56],
    2024: [0.42, 0.83, 0.16, 0.38, 0.46, 0.21, 0.38, -0.02, 0.44, 0.56, 0.39, 0.52],
    2025: [0.16, 1.31, 0.56, 0.43, 0.26, 0.24, 0.26, -0.11, 0.48, 0.09, 0.18, 0.33],
}


def caminho(nome):
    return os.path.join(BASE, nome)


# ======================= 1. SIH/SUS: AIHs -> tabelas por região =======================
def ler_bruto():
    p = caminho(BRUTO)
    if not os.path.exists(p):
        sys.exit(f"Não achei {p}. Rode o extrator-ftp.py antes.")
    with open(p, encoding="utf-8-sig") as f:
        sep = ";" if ";" in f.readline() else ","
    df = pd.read_csv(p, sep=sep, dtype=str, encoding="utf-8-sig")
    df["VAL_TOT"] = pd.to_numeric(df["VAL_TOT"], errors="coerce").fillna(0.0)
    idade = pd.to_numeric(df["IDADE"], errors="coerce").fillna(0).astype(int)
    cod = pd.to_numeric(df["COD_IDADE"], errors="coerce").fillna(0).astype(int)
    # COD_IDADE: 4 = anos; 5 = 100 anos ou mais (valor guardado + 100); 1/2/3 = horas/dias/meses (bebê)
    df["idade_anos"] = np.where(cod == 4, idade, np.where(cod == 5, idade + 100, 0))
    df["reg_res"] = pd.to_numeric(df["MUNIC_RES"].str[:2], errors="coerce").map(UF2REG)
    df["reg_real"] = pd.to_numeric(df["MUNIC_MOV"].str[:2], errors="coerce").map(UF2REG)
    df["ano"] = pd.to_numeric(df[CAMPO_ANO].str[:4], errors="coerce")
    n0 = len(df)
    print(f"Bruto: {n0} AIHs; N_AIH duplicados: {int(df['N_AIH'].duplicated().sum())}.")
    df = df[df["ano"].between(ANO_INI, ANO_FIM)]
    print(f"   {len(df)} dentro de {ANO_INI}-{ANO_FIM} (ano pela data de internação); {n0 - len(df)} fora do período.")
    sem = df["reg_res"].isna() | df["reg_real"].isna()
    if sem.any():
        print(f"   ATENÇÃO: {int(sem.sum())} AIHs sem região de residência ou de realização reconhecida; "
              f"excluídas das DUAS tabelas, para que os totais continuem iguais.")
        df = df[~sem]
    return df


def agregar(df, col_regiao, com_valor):
    """Linhas Ano x Região (+ Brasil) x Grupo, a partir de col_regiao (reg_res ou reg_real)."""
    linhas = []
    for grupo in GRUPOS:
        d = df if grupo == "Geral" else df[df["idade_anos"] >= IDADE_CORTE]
        g = d.groupby(["ano", col_regiao]).agg(n=("VAL_TOT", "size"), valor=("VAL_TOT", "sum"))
        for ano in range(ANO_INI, ANO_FIM + 1):
            n_br, v_br = 0, 0.0
            for reg in REG5:
                n, v = (int(g.loc[(ano, reg), "n"]), float(g.loc[(ano, reg), "valor"])) if (ano, reg) in g.index else (0, 0.0)
                linhas.append((ano, reg, grupo, n, round(v, 2)))
                n_br += n
                v_br += v
            linhas.append((ano, "Brasil (Total)", grupo, n_br, round(v_br, 2)))
    t = pd.DataFrame(linhas, columns=["Ano", "Região", "Grupo", "Internações", "Valor"])
    if (t["Internações"] <= 0).any():
        ruins = t[t["Internações"] <= 0].head(5).to_string(index=False)
        sys.exit(f"Há combinações ano/região/grupo sem internações (o Ln não existe):\n{ruins}\n"
                 f"Confira se o marcapasso_bruto.csv cobre {ANO_INI}-{ANO_FIM}.")
    return t if com_valor else t.drop(columns="Valor")


def ler_populacao():
    p = caminho(POPULACAO)
    if not os.path.exists(p):
        sys.exit(f"Não achei {p}. Rode o extrator_ibge.py antes.")
    return pd.read_csv(p, sep=";", encoding="utf-8-sig")


# ================================== 2. IPCA ===================================
def ipca_do_bcb():
    url = ("https://api.bcb.gov.br/dados/serie/bcdata.sgs.433/dados"
           f"?formato=json&dataInicial=01/01/{ANO_INI}&dataFinal=31/12/{ANO_FIM}")
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    dados = json.loads(urllib.request.urlopen(req, timeout=30).read().decode("utf-8"))
    por_ano = {}
    for r in dados:
        dia, mes, ano = r["data"].split("/")
        por_ano.setdefault(int(ano), {})[int(mes)] = float(r["valor"])
    out = {}
    for a in range(ANO_INI, ANO_FIM + 1):
        if sorted(por_ano.get(a, {})) != list(range(1, 13)):
            raise ValueError(f"série do BCB incompleta em {a}")
        out[a] = [por_ano[a][m] for m in range(1, 13)]
    return out


def montar_deflator():
    fonte = "série embutida no script"
    mensal = IPCA_MENSAL_EMBUTIDO
    if USAR_BCB:
        try:
            mensal = ipca_do_bcb()
            fonte = "Banco Central (SGS 433, IPCA mensal)"
        except Exception as e:                                   # noqa: BLE001
            print(f"   (não consegui baixar do Banco Central: {e}\n    usando a série embutida)")
    idx, nivel = {}, 100.0                                       # dez/2015 = 100
    for a in range(ANO_INI, ANO_FIM + 1):
        niveis = []
        for v in mensal[a]:
            nivel *= 1 + v / 100
            niveis.append(nivel)
        idx[a] = niveis
    linhas, dez_ant = [], 100.0
    for a in range(ANO_INI, ANO_FIM + 1):
        medio, dez = float(np.mean(idx[a])), idx[a][-1]
        linhas.append({"Ano": a, "IPCA dez/dez (%)": (dez / dez_ant - 1) * 100,
                       "Índice médio do ano": medio, "Índice de dezembro": dez})
        dez_ant = dez
    t = pd.DataFrame(linhas)
    base = t.loc[t.Ano == ANO_FIM, "Índice médio do ano" if BASE_PRECOS == "media2025" else "Índice de dezembro"].iloc[0]
    t["Fator de correção"] = base / t["Índice médio do ano"]
    print(f"IPCA: {fonte}. Conferência com o IPCA anual oficial (dez/dez):")
    for _, r in t.iterrows():
        of = IPCA_ANUAL_OFICIAL.get(int(r.Ano))
        if of is not None:
            marca = "OK" if abs(r["IPCA dez/dez (%)"] - of) <= 0.02 else "DIFERENTE!"
            print(f"   {int(r.Ano)}: calculado {r['IPCA dez/dez (%)']:.2f}%  oficial {of:.2f}%  {marca}")
    return t, fonte, mensal


# ============================ 3. PRAIS-WINSTEN ================================
def prais_winsten(y, X, tol=1e-10, maxit=500):
    """Prais-Winsten iterado (mesmo esquema do 'prais' do Stata / ar1 --pwe do gretl).
    rho é re-estimado a cada passo com os resíduos do modelo original (y - X b)."""
    y, X = np.asarray(y, float), np.asarray(X, float)
    n, k = X.shape
    beta = np.linalg.lstsq(X, y, rcond=None)[0]
    e0 = y - X @ beta
    dw_orig = float(np.sum(np.diff(e0) ** 2) / np.sum(e0 ** 2))
    rho = 0.0
    for _ in range(maxit):
        e = y - X @ beta
        rho_novo = float(np.clip((e[1:] @ e[:-1]) / (e[:-1] @ e[:-1]), -0.9999, 0.9999))
        w = np.sqrt(1 - rho_novo ** 2)
        ys = np.concatenate([[w * y[0]], y[1:] - rho_novo * y[:-1]])
        Xs = np.vstack([w * X[0], X[1:] - rho_novo * X[:-1]])
        beta_novo = np.linalg.lstsq(Xs, ys, rcond=None)[0]
        fim = abs(rho_novo - rho) < tol and np.max(np.abs(beta_novo - beta)) < tol
        beta, rho = beta_novo, rho_novo
        if fim:
            break
    r = ys - Xs @ beta
    gl = n - k
    s2 = float(r @ r / gl)
    cov = s2 * np.linalg.inv(Xs.T @ Xs)
    se = np.sqrt(np.diag(cov))
    t = beta / se
    p = 2 * stats.t.sf(np.abs(t), gl)
    dw_transf = float(np.sum(np.diff(r) ** 2) / np.sum(r ** 2))
    return {"beta": beta, "se": se, "t": t, "p": p, "rho": rho, "gl": gl, "n": n,
            "dw_orig": dw_orig, "dw_transf": dw_transf}


def classificar(b, p):
    if p < ALFA:
        return "crescente" if b > 0 else "decrescente"
    return "estável"


# Cada cenário: anos EXCLUÍDOS e anos marcados com UMA dummy (1 nos anos listados, 0 nos demais).
CENARIOS = {
    "todos":        {"nome": "Todos os anos (2016-2025)", "excluir": [], "dummy": []},
    "sem2020":      {"nome": "Sem 2020", "excluir": [2020], "dummy": []},
    "dummy2020":    {"nome": "Todos os anos + dummy 2020", "excluir": [], "dummy": [2020]},
    "sem2021":      {"nome": "Sem 2021", "excluir": [2021], "dummy": []},
    "dummy2021":    {"nome": "Todos os anos + dummy 2021", "excluir": [], "dummy": [2021]},
    "sem2020_21":   {"nome": "Sem 2020 e 2021", "excluir": [2020, 2021], "dummy": []},
    "dummy2020_21": {"nome": "Todos os anos + dummy 2020-2021 (uma dummy)", "excluir": [], "dummy": [2020, 2021]},
}
CENARIOS_PLANO = ["todos", "dummy2020_21"]       # modelo principal + sensibilidade prevista no plano
if not INCLUIR_EXTRAS:
    CENARIOS = {k: v for k, v in CENARIOS.items() if k in CENARIOS_PLANO}


def rodar_cenario(anos, y, cenario):
    spec = CENARIOS[cenario]
    anos, y = np.asarray(anos, float), np.asarray(y, float)
    m = ~np.isin(anos, spec["excluir"])
    anos, y = anos[m], y[m]
    cols = [np.ones(len(anos)), anos - anos.min()]
    if spec["dummy"]:
        cols.append(np.isin(anos, spec["dummy"]).astype(float))
    return prais_winsten(y, np.column_stack(cols))


INDICADORES = {"taxa": "Taxa por 100 mil [residência] (ln)",
               "vt_real": "Valor total aprovado corrigido [realização] (ln)",
               "vm_nominal": "Valor médio nominal [realização] (ln)",
               "vm_real": "Valor médio corrigido [realização] (ln)"}


def regressoes(d, dr):
    saida = []
    for grupo in GRUPOS:
        for reg in REGIOES:
            s = d[(d.Grupo == grupo) & (d["Região"] == reg)].sort_values("Ano")
            sr = dr[(dr.Grupo == grupo) & (dr["Região"] == reg)].sort_values("Ano")
            series = {"taxa": (s["Ano"].to_numpy(), s["Taxa"].to_numpy()),
                      "vt_real": (sr["Ano"].to_numpy(), sr["Valor_corr"].to_numpy()),
                      "vm_nominal": (sr["Ano"].to_numpy(), sr["VM"].to_numpy()),
                      "vm_real": (sr["Ano"].to_numpy(), sr["VM_corr"].to_numpy())}
            for ind, (anos, y) in series.items():
                for cen, spec in CENARIOS.items():
                    res = rodar_cenario(anos, np.log(y), cen)
                    b, se, p, gl = res["beta"][1], res["se"][1], res["p"][1], res["gl"]
                    tc = stats.t.ppf(1 - ALFA / 2, gl)
                    r = {"Grupo": grupo, "Região": reg, "Indicador": INDICADORES[ind],
                         "Cenário": spec["nome"], "n": res["n"], "beta1 (ln/ano)": b,
                         "Erro-padrão": se, "t": res["t"][1], "p-valor": p, "rho": res["rho"],
                         "DW original": res["dw_orig"], "DW transformado": res["dw_transf"],
                         "APC (% ao ano)": (np.exp(b) - 1) * 100,
                         "IC95% inf": (np.exp(b - tc * se) - 1) * 100,
                         "IC95% sup": (np.exp(b + tc * se) - 1) * 100,
                         "Tendência": classificar(b, p)}
                    if spec["dummy"]:
                        r["beta dummy"] = res["beta"][2]
                        r["p dummy"] = res["p"][2]
                    saida.append(r)
    return pd.DataFrame(saida)


def variacoes(d, dr, fatores):
    """Variação 2016 -> 2025: valor médio (realização, nominal e real) e taxa (residência)."""
    ini, fim = ANO_INI, ANO_FIM
    ipca_acum = (fatores.loc[fatores.Ano == ini, "Fator de correção"].iloc[0] - 1) * 100
    v_val, v_tax = [], []
    for grupo in GRUPOS:
        for reg in REGIOES:
            s = dr[(dr.Grupo == grupo) & (dr["Região"] == reg)].set_index("Ano")
            v_val.append({"Grupo": grupo, "Região": reg,
                          f"Valor médio {ini} (nominal)": s.loc[ini, "VM"],
                          f"Valor médio {fim} (nominal)": s.loc[fim, "VM"],
                          "Variação nominal (%)": (s.loc[fim, "VM"] / s.loc[ini, "VM"] - 1) * 100,
                          f"Inflação IPCA {ini}-{fim} (%) (médias anuais)": ipca_acum,
                          f"Valor médio {ini} (corrigido)": s.loc[ini, "VM_corr"],
                          f"Valor médio {fim} (corrigido)": s.loc[fim, "VM_corr"],
                          "Variação REAL (%)": (s.loc[fim, "VM_corr"] / s.loc[ini, "VM_corr"] - 1) * 100})
            t = d[(d.Grupo == grupo) & (d["Região"] == reg)].set_index("Ano")
            v_tax.append({"Grupo": grupo, "Região": reg,
                          f"Taxa {ini} (por 100 mil)": t.loc[ini, "Taxa"],
                          f"Taxa {fim} (por 100 mil)": t.loc[fim, "Taxa"],
                          "Variação da taxa (%)": (t.loc[fim, "Taxa"] / t.loc[ini, "Taxa"] - 1) * 100})
    return pd.DataFrame(v_val), pd.DataFrame(v_tax)


# ================================ 4. GRETL ====================================
def exportar_gretl(d, dr):
    pasta = os.path.join(SAIDA, "gretl")
    os.makedirs(pasta, exist_ok=True)
    anos = sorted(d.Ano.unique())
    wide = pd.DataFrame({"ano": anos})
    wide["d2020"] = (wide["ano"] == 2020).astype(int)
    wide["d2021"] = (wide["ano"] == 2021).astype(int)
    wide["dpand"] = wide["ano"].isin([2020, 2021]).astype(int)
    series = []
    for g, gs in SLUG_GRP.items():
        for r, rs in SLUG_REG.items():
            s = d[(d.Grupo == g) & (d["Região"] == r)].sort_values("Ano")
            sr = dr[(dr.Grupo == g) & (dr["Região"] == r)].sort_values("Ano")
            for pref, tab, col in [("taxa", s, "Taxa"), ("vtr", sr, "Valor_corr"),
                                   ("vmn", sr, "VM"), ("vmr", sr, "VM_corr")]:
                wide[f"{pref}_{gs}_{rs}"] = tab[col].round(6).to_numpy()
                wide[f"ln{pref}_{gs}_{rs}"] = np.log(tab[col]).round(8).to_numpy()
            series.append((gs, rs))
    wide.to_csv(os.path.join(pasta, "gretl_dados.csv"), index=False)
    arquivos = {(): "gretl_dados.csv"}
    for nome, excl in [("sem2020", [2020]), ("sem2021", [2021]), ("sem2020_21", [2020, 2021])]:
        if any(spec["excluir"] == excl for spec in CENARIOS.values()):
            wide[~wide.ano.isin(excl)].to_csv(os.path.join(pasta, f"gretl_dados_{nome}.csv"), index=False)
            arquivos[tuple(excl)] = f"gretl_dados_{nome}.csv"
    nome_dummy = {(2020,): "d2020", (2021,): "d2021", (2020, 2021): "dpand"}
    linhas = ["# Gerado por modelador.py. Prais-Winsten iterado (ar1 --pwe) para cada série.",
              "# taxa = região de residência; vtr/vmn/vmr = região de realização (vtr: valor total corrigido).",
              "# Rode em: Arquivo > Arquivos de script > Novo script > cole/abra este arquivo > Executar.",
              "# Se algum comando der erro, me mande a mensagem.",
              "outfile resultados_gretl.txt --write", ""]
    for i, (cen, spec) in enumerate(CENARIOS.items(), 1):
        excl = tuple(spec["excluir"])
        regs = ["const", "ano"] + ([nome_dummy[tuple(spec["dummy"])]] if spec["dummy"] else [])
        linhas += [f"# ===== {i}) {spec['nome']} =====", f"open {arquivos[excl]}",
                   "setobs 1 1 --time-series" if excl else f"setobs 1 {ANO_INI} --time-series"]
        for gs, rs in series:
            for pref in ["taxa", "vtr", "vmn", "vmr"]:
                linhas.append(f"ar1 ln{pref}_{gs}_{rs} {' '.join(regs)} --pwe")
        linhas.append("")
    linhas.append("outfile --end")
    with open(os.path.join(pasta, "gretl_prais.inp"), "w", encoding="utf-8") as f:
        f.write("\n".join(linhas) + "\n")
    return pasta


# ================================ 5. GRÁFICOS =================================
CORES = {"Região Norte": "#1b9e77", "Região Nordeste": "#d95f02", "Região Sudeste": "#7570b3",
         "Região Sul": "#e7298a", "Região Centro-Oeste": "#66a61e", "Brasil (Total)": "black"}


def _marca_pandemia(ax):
    ax.axvspan(2020 - 0.5, 2020 + 0.5, color="0.86", zorder=0)
    ax.axvspan(2021 - 0.5, 2021 + 0.5, color="0.94", zorder=0)
    for a in (2020, 2021):
        ax.text(a, ax.get_ylim()[1], str(a), ha="center", va="top", fontsize=8, color="0.4")


def _tendencia_brasil(ax, s, col):
    anos = s["Ano"].to_numpy(float)
    res = rodar_cenario(anos, np.log(s[col].to_numpy()), "todos")
    X = np.column_stack([np.ones(len(anos)), anos - anos.min()])
    ax.plot(anos, np.exp(X @ res["beta"]), "--", color="black", lw=1, label="Tendência Brasil (Prais-Winsten)")


def grafico_regioes(d, grupo, col, titulo, ylabel, arquivo, tendencia=True):
    fig, ax = plt.subplots(figsize=(9, 5.5))
    for reg in REGIOES:
        s = d[(d.Grupo == grupo) & (d["Região"] == reg)].sort_values("Ano")
        ax.plot(s.Ano, s[col], marker="o", ms=5, lw=2.8 if reg == "Brasil (Total)" else 1.6,
                color=CORES[reg], label=reg.replace("Região ", ""))
    if tendencia:
        _tendencia_brasil(ax, d[(d.Grupo == grupo) & (d["Região"] == "Brasil (Total)")].sort_values("Ano"), col)
    ax.set_xticks(range(ANO_INI, ANO_FIM + 1))
    ax.set_xlabel("Ano de internação")
    ax.set_ylabel(ylabel)
    ax.set_title(titulo)
    ax.grid(alpha=0.3)
    ax.set_ylim(bottom=0)
    _marca_pandemia(ax)
    ax.legend(ncol=2, fontsize=8)
    fig.tight_layout()
    fig.savefig(arquivo, dpi=170)
    plt.close(fig)


def grafico_nominal_vs_real(d, arquivo):
    fig, axs = plt.subplots(1, 2, figsize=(11, 4.8), sharey=True)
    for ax, grupo in zip(axs, GRUPOS):
        s = d[(d.Grupo == grupo) & (d["Região"] == "Brasil (Total)")].sort_values("Ano")
        ax.plot(s.Ano, s.VM, marker="o", color="tab:blue", label="Nominal (R$ da época)")
        ax.plot(s.Ano, s.VM_corr, marker="s", color="tab:red", label="Corrigido pelo IPCA (R$ de 2025)")
        ax.set_title(f"Brasil - {grupo}")
        ax.set_xticks(range(ANO_INI, ANO_FIM + 1))
        ax.tick_params(axis="x", labelrotation=45)
        ax.grid(alpha=0.3)
        ax.set_ylim(bottom=0)
        ax.set_xlabel("Ano de internação")
        _marca_pandemia(ax)
    axs[0].set_ylabel("Valor médio por internação (R$)")
    axs[0].legend(fontsize=8, loc="lower left")
    fig.suptitle("Valor médio por internação: nominal x corrigido pela inflação")
    fig.tight_layout()
    fig.savefig(arquivo, dpi=170)
    plt.close(fig)


def grafico_geral_vs_60(d, arquivo):
    fig, ax = plt.subplots(figsize=(9, 5))
    for grupo, cor, mk in [("Geral", "tab:blue", "o"), ("60+", "tab:red", "s")]:
        s = d[(d.Grupo == grupo) & (d["Região"] == "Brasil (Total)")].sort_values("Ano")
        ax.plot(s.Ano, s.Taxa, marker=mk, color=cor, lw=2, label=grupo)
    ax.set_yscale("log")
    ax.set_xticks(range(ANO_INI, ANO_FIM + 1))
    ax.set_xlabel("Ano de internação")
    ax.set_ylabel("Taxa por 100 mil habitantes (escala log)")
    ax.set_title("Brasil: taxa de internação, população geral x 60+")
    ax.grid(alpha=0.3, which="both")
    _marca_pandemia(ax)
    ax.legend()
    fig.tight_layout()
    fig.savefig(arquivo, dpi=170)
    plt.close(fig)


def graficos(d, dr):
    pasta = os.path.join(SAIDA, "graficos")
    os.makedirs(pasta, exist_ok=True)
    f = lambda n: os.path.join(pasta, n)           # noqa: E731
    grafico_regioes(d, "Geral", "Taxa", "Taxa de internação por região de residência - população geral",
                    "Internações por 100 mil habitantes", f("fig1_taxa_geral_por_regiao.png"))
    grafico_regioes(d, "60+", "Taxa", "Taxa de internação por região de residência - 60 anos ou mais",
                    "Internações por 100 mil habitantes (60+)", f("fig2_taxa_60mais_por_regiao.png"))
    grafico_regioes(dr, "Geral", "VM_corr", "Valor médio por internação corrigido (IPCA), por região de realização - geral",
                    "R$ de 2025", f("fig3_valor_medio_corrigido_geral.png"))
    grafico_regioes(dr, "60+", "VM_corr", "Valor médio por internação corrigido (IPCA), por região de realização - 60+",
                    "R$ de 2025", f("fig4_valor_medio_corrigido_60mais.png"))
    grafico_nominal_vs_real(dr, f("fig5_brasil_nominal_vs_corrigido.png"))
    grafico_geral_vs_60(d, f("fig6_brasil_taxa_geral_vs_60mais.png"))
    return pasta


# ============================ 6. PASTA DE TRABALHO ============================
def _cabecalho(ws, nomes, larguras):
    ws.append(nomes)
    for c in ws[1]:
        c.font = Font(bold=True)
        c.alignment = Alignment(wrap_text=True, vertical="top")
    for i, w in enumerate(larguras, start=1):
        ws.column_dimensions[ws.cell(1, i).column_letter].width = w
    ws.freeze_panes = "A2"


def _df_na_aba(ws, df, larguras=None, inicio_linha=None):
    """Escreve um DataFrame (cabeçalho em negrito) a partir de inicio_linha (ou do início da aba)."""
    r0 = inicio_linha or 1
    for j, c in enumerate(df.columns, start=1):
        cel = ws.cell(r0, j, c)
        cel.font = Font(bold=True)
        cel.alignment = Alignment(wrap_text=True, vertical="top")
    for i, linha in enumerate(df.itertuples(index=False), start=r0 + 1):
        for j, v in enumerate(linha, start=1):
            if isinstance(v, (np.integer,)):
                v = int(v)
            elif isinstance(v, (np.floating, float)):
                v = None if pd.isna(v) else float(v)
            ws.cell(i, j, v)
    for j, w in enumerate(larguras or [], start=1):
        ws.column_dimensions[ws.cell(r0, j).column_letter].width = w


def gravar_pasta(res, rea, pop, fatores, mensal, fonte_ipca, reg, var_val, var_tax, pasta_graficos):
    wb = Workbook()
    n = len(res)                     # = len(rea) = 120
    rotulo = "preços médios de 2025" if BASE_PRECOS == "media2025" else "preços de dez/2025"
    ref = f"R$ {ANO_FIM}"

    # --- Leia-me ---
    ws = wb.active
    ws.title = "Leia-me"
    linhas = [
        f"Implante de marcapasso de câmara dupla transvenoso (SIGTAP 0406010650), SIH/SUS, {ANO_INI}-{ANO_FIM}",
        "Esta pasta de trabalho contem:",
        "Taxa de internação: internações por ano, por REGIÃO DE RESIDÊNCIA do paciente, população residente e taxa por 100 mil habitantes.",
        "Custo de realização: internações por ano, valor total e valor médio por REGIÃO DE REALIZAÇÃO, corrigido por IPCA a preços médios; razão e diferença para a o Sudeste (real e nominal)",
        f"IPCA: variação anual e mensal usada na correção monetária.",
        "Populacao: população do IBGE usada.",
        "Variacao_2016_2025: variação do valor médio por internação (nominal e real) e da taxa de implante.",
        "PW_Taxa: Prais-Winsten da taxa de implante (pergunta: a taxa está crescendo?).",
        "PW_Valor: Prais-Winsten do valor por internação (pergunta: o valor está crescendo?), nominal e corrigido.",
        "Graficos: figuras geradas (também em resultados/graficos). Pasta resultados/gretl: dados e script para repetir no gretl.",
        "",
        "Critérios: as duas bases usam as mesmas AIHs e o mesmo ano (data de internação); por isso o total do Brasil é igual.",
        "Regressão: ln(y) = b0 + b1*ano (+ dummy 2020-2021 no cenário de sensibilidade), Prais-Winsten iterado; "
        "APC = (exp(b1) - 1) x 100.",
        f"População: {FONTE_POP}.",
        f"IPCA: {fonte_ipca}; índice com dez/2015 = 100; valor corrigido = valor x (índice médio de {ANO_FIM} / índice médio do ano).",
        "Reprodução: executar os scripts nesta ordem: extrator-ftp-SIH-SUS.py > extrator_ibge.py > modelador.py",
    ]
    for l in linhas:
        ws.append([l])
    ws["A1"].font = Font(bold=True)
    ws.column_dimensions["A"].width = 135

    # --- Residencia_Taxa ---
    w1 = wb.create_sheet("Taxa de internação")
    _cabecalho(w1, ["Ano", "Região de residência", "Grupo", "Internações", "População",
                    "Taxa por 100 mil", "Ln(taxa)"], [7, 24, 8, 13, 14, 16, 10])
    for i, r in enumerate(res.itertuples(index=False), start=2):
        w1.append([int(r.Ano), r[1], r.Grupo, int(r.Internações), int(r.População),
                   f"=D{i}/E{i}*100000", f"=LN(F{i})"])
        w1[f"D{i}"].number_format = "#,##0"
        w1[f"E{i}"].number_format = "#,##0"
        w1[f"F{i}"].number_format = "0.00"
        w1[f"G{i}"].number_format = "0.0000"

    # --- Realizacao_Custo ---
    w2 = wb.create_sheet("Custo de realização")
    _cabecalho(w2, ["Ano", "Região de realização", "Grupo", "Internações", "Valor total aprovado (R$)",
                    "Valor médio por internação (R$)", "Razão do valor médio / Sudeste",
                    "Diferença p/ Sudeste (R$ nominais)", "Fator de correção (IPCA)",
                    f"Valor total corrigido ({ref})", f"Valor médio corrigido ({ref})",
                    f"Diferença real p/ Sudeste ({ref})"],
              [7, 24, 8, 13, 20, 20, 16, 20, 14, 22, 20, 24])
    ult = n + 1
    sud = lambda col, i: (f'SUMIFS(${col}$2:${col}${ult},$A$2:$A${ult},A{i},'                   # noqa: E731
                          f'$B$2:$B${ult},"Região Sudeste",$C$2:$C${ult},C{i})')
    for i, r in enumerate(rea.itertuples(index=False), start=2):
        w2.append([int(r.Ano), r[1], r.Grupo, int(r.Internações), float(r.Valor),
                   f"=E{i}/D{i}", f"=F{i}/{sud('F', i)}", f"=F{i}-{sud('F', i)}",
                   f"=VLOOKUP(A{i},IPCA!$A$2:$E${1 + len(fatores)},5,FALSE)",
                   f"=E{i}*I{i}", f"=J{i}/D{i}", f"=K{i}-{sud('K', i)}"])
        for col, fmt in zip("DEFGHIJKL", ["#,##0", "#,##0.00", "#,##0.00", "0.0000", "#,##0.00",
                                          "0.0000", "#,##0.00", "#,##0.00", "#,##0.00"]):
            w2[f"{col}{i}"].number_format = fmt

    # --- IPCA ---
    w3 = wb.create_sheet("IPCA")
    _cabecalho(w3, ["Ano", "IPCA dez/dez (%)", "Índice médio do ano", "Índice de dezembro", "Fator de correção"],
               [8, 18, 20, 20, 18])
    for _, r in fatores.iterrows():
        w3.append([int(r.Ano), round(float(r["IPCA dez/dez (%)"]), 2), float(r["Índice médio do ano"]),
                   float(r["Índice de dezembro"]), float(r["Fator de correção"])])
    w3.append([])
    w3.append([f"Fonte: IPCA mensal, {fonte_ipca}. Índice com dez/2015 = 100. Valores corrigidos a {rotulo}: "
               f"valor x (índice médio de {ANO_FIM} / índice médio do ano da internação)."])
    w3.append([])
    mens = pd.DataFrame([[a] + mensal[a] for a in range(ANO_INI, ANO_FIM + 1)],
                        columns=["IPCA mensal (% a.m.)"] + ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun",
                                                           "Jul", "Ago", "Set", "Out", "Nov", "Dez"])
    _df_na_aba(w3, mens, inicio_linha=w3.max_row + 1)

    # --- Populacao ---
    w4 = wb.create_sheet("Populacao")
    _cabecalho(w4, ["Ano", "Região", "Grupo", "População", "Fonte"], [7, 22, 8, 14, 100])
    for r in pop.itertuples(index=False):
        w4.append([int(r.Ano), r[1], r.Grupo, int(r[3]), FONTE_POP])
    for i in range(2, w4.max_row + 1):
        w4[f"D{i}"].number_format = "#,##0"

    # --- Variacao_2016_2025 ---
    w5 = wb.create_sheet("Variacao_2016_2025")
    w5.append([f"Valor médio por internação ({ANO_INI} -> {ANO_FIM}), por região de realização"])
    w5["A1"].font = Font(bold=True)
    _df_na_aba(w5, var_val.round(2), [8, 22, 20, 20, 18, 26, 20, 20, 18], inicio_linha=2)
    r0 = w5.max_row + 3
    w5.cell(r0 - 1, 1, f"Taxa de implante ({ANO_INI} -> {ANO_FIM}), por região de residência").font = Font(bold=True)
    _df_na_aba(w5, var_tax.round(2), inicio_linha=r0)

    # --- PW ---
    pw_taxa = reg[reg["Indicador"].str.contains("residência")].reset_index(drop=True)
    pw_valor = reg[reg["Indicador"].str.contains("realização")].reset_index(drop=True)
    larg = [8, 22, 48, 40] + [12] * (len(reg.columns) - 4)
    _df_na_aba(wb.create_sheet("PW_Taxa"), pw_taxa.round(5), larg)
    _df_na_aba(wb.create_sheet("PW_Valor"), pw_valor.round(5), larg)
    wb["PW_Taxa"].freeze_panes = "A2"
    wb["PW_Valor"].freeze_panes = "A2"

    # --- Graficos ---
    wg = wb.create_sheet("Graficos")
    linha = 1
    for nome in sorted(os.listdir(pasta_graficos)):
        if not nome.endswith(".png"):
            continue
        wg.cell(linha, 1, nome).font = Font(bold=True)
        img = XlImage(os.path.join(pasta_graficos, nome))
        escala = 760 / img.width
        img.width, img.height = 760, int(img.height * escala)
        wg.add_image(img, f"A{linha + 1}")
        linha += int(img.height / 20) + 4

    wb.calculation.fullCalcOnLoad = True
    destino = os.path.join(SAIDA, NOME_PASTA)
    wb.save(destino)
    return destino


# ================================== MAIN ======================================
def main():
    os.makedirs(SAIDA, exist_ok=True)
    df = ler_bruto()
    pop = ler_populacao()

    res = agregar(df, "reg_res", com_valor=False)                 # residência: só internações
    rea = agregar(df, "reg_real", com_valor=True)                 # realização: internações e valores
    for g in GRUPOS:
        a = res[(res.Grupo == g) & (res["Região"] == "Brasil (Total)")]["Internações"].sum()
        b = rea[(rea.Grupo == g) & (rea["Região"] == "Brasil (Total)")]["Internações"].sum()
        print(f"   Conferência Brasil/{g}: residência {a:,} x realização {b:,} -> {'OK, fecham' if a == b else 'ATENÇÃO!'}"
              .replace(",", "."))
        if a != b:
            sys.exit("Os totais por residência e por realização deveriam ser iguais.")

    res = res.merge(pop.rename(columns={"População": "População"}), on=["Ano", "Região", "Grupo"], how="left")
    if res["População"].isna().any():
        sys.exit("Há linhas sem população. Confira o ANO_INI/ANO_FIM do extrator_ibge.py e deste script.")
    res["Taxa"] = res["Internações"] / res["População"] * 1e5

    print()
    fatores, fonte_ipca, mensal = montar_deflator()
    rea = rea.merge(fatores[["Ano", "Fator de correção"]], on="Ano")
    rea["VM"] = rea["Valor"] / rea["Internações"]
    rea["Valor_corr"] = rea["Valor"] * rea["Fator de correção"]
    rea["VM_corr"] = rea["Valor_corr"] / rea["Internações"]

    ordem = lambda t: t.assign(_g=t["Grupo"].map({"Geral": 0, "60+": 1}),             # noqa: E731
                               _r=t["Região"].map(ORDEM_REG)).sort_values(["_g", "Ano", "_r"]) \
        .drop(columns=["_g", "_r"]).reset_index(drop=True)
    res, rea = ordem(res), ordem(rea)
    print(f"\nTaxas: região de RESIDÊNCIA | Valores: região de REALIZAÇÃO")
    print(f"Cenários: {[c['nome'] for c in CENARIOS.values()]}")

    print(f"Rodando Prais-Winsten (4 indicadores x 12 séries x {len(CENARIOS)} cenários)...")
    reg = regressoes(res, rea)
    var_val, var_tax = variacoes(res, rea, fatores)
    reg.round(5).to_csv(os.path.join(SAIDA, "prais_winsten.csv"), sep=";", decimal=",",
                        index=False, encoding="utf-8-sig")
    print("Exportando para o gretl...")
    print("   ", exportar_gretl(res, rea))
    print("Gerando gráficos...")
    pasta_graficos = graficos(res, rea)
    print("   ", pasta_graficos)
    print("Montando a pasta de trabalho...")
    destino = gravar_pasta(res, rea, pop, fatores, mensal, fonte_ipca, reg, var_val, var_tax, pasta_graficos)

    pd.set_option("display.width", 220)
    print("\n=== Brasil: variação do valor médio por internação ===")
    print(var_val[var_val["Região"] == "Brasil (Total)"].round(1).to_string(index=False))
    print("\n=== Brasil - Prais-Winsten (todos os anos) ===")
    cols = ["Grupo", "Indicador", "beta1 (ln/ano)", "p-valor", "APC (% ao ano)", "IC95% inf", "IC95% sup", "Tendência"]
    print(reg[(reg["Região"] == "Brasil (Total)") & (reg["Cenário"] == CENARIOS["todos"]["nome"])][cols]
          .round(4).to_string(index=False))
    nomes = [c["nome"] for c in CENARIOS.values()]
    if len(nomes) > 1:
        pt = reg.pivot_table(index=["Grupo", "Região", "Indicador"], columns="Cenário", values="Tendência",
                             aggfunc="first")[nomes]
        base = pt[CENARIOS["todos"]["nome"]]
        print(f"\n=== Sensibilidade: em quantas das {len(pt)} séries a classificação muda em relação ao modelo principal? ===")
        for nome in nomes[1:]:
            mudou = pt[pt[nome] != base]
            print(f"   {nome:45s}: {len(mudou)} de {len(pt)}")
            for idx, row in mudou.iterrows():
                print(f"        {idx[0]:5s} {idx[1]:20s} {idx[2]:48s} {base[idx]} -> {row[nome]}")
    print(f"\nPronto. Pasta de trabalho: {destino}\nGretl: {os.path.join(SAIDA, 'gretl')} | "
          f"Gráficos: {pasta_graficos}")


if __name__ == "__main__":
    main()
