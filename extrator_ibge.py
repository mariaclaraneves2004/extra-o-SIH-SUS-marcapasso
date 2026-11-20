"""
extrator_ibge.py  -  população ANO A ANO (2016-2025) por Grande Região, a partir da PROJEÇÃO OFICIAL do IBGE
(Projeções da População do Brasil e Unidades da Federação, Revisão 2024, base no Censo 2022; data de referência
1º de julho de cada ano). Gera o total ("Geral") e o grupo 60 anos ou mais ("60+").

Fonte (FTP do IBGE):
  https://ftp.ibge.gov.br/Projecao_da_Populacao/Projecao_da_Populacao_2024/projecoes_2024_tab1_idade_simples.xlsx
  (população por sexo e idade simples, 27 UFs + Brasil, 2000-2070). As UFs são somadas nas 5 regiões.

O script baixa o arquivo para a pasta cache_ibge (e reaproveita se já estiver lá), encontra sozinho o cabeçalho
da planilha, as colunas de idade/sexo/UF e os anos, e confere se a soma das UFs bate com a linha do Brasil.
Se o download falhar, baixe o .xlsx no navegador e coloque-o na pasta cache_ibge (ou ao lado deste script).

Saída:  ibge_populacao.csv   Ano ; Região ; Grupo ; População   (separador ";")
Uso:    py -3.12 extrator_ibge.py
Precisa de: pip install pandas openpyxl
"""
import datetime as dt
import gzip
import os
import re
import sys
import time
import urllib.request

import pandas as pd

# ------------------------------ CONFIGURAÇÃO ------------------------------
PASTA_URL = "https://ftp.ibge.gov.br/Projecao_da_Populacao/Projecao_da_Populacao_2024/"
ARQUIVO = "projecoes_2024_tab1_idade_simples.xlsx"
ANO_INI, ANO_FIM = 2016, 2025
IDADE_CORTE = 60
# ---------------------------------------------------------------------------

BASE = os.path.dirname(os.path.abspath(__file__))
CACHE = os.path.join(BASE, "cache_ibge")
SAIDA = os.path.join(BASE, "ibge_populacao.csv")

REGIOES = ["Região Norte", "Região Nordeste", "Região Sudeste", "Região Sul", "Região Centro-Oeste"]
UF2REG = {}
for _reg, _ufs in zip(REGIOES, ["RO AC AM RR PA AP TO", "MA PI CE RN PB PE AL SE BA", "MG ES RJ SP",
                                "PR SC RS", "MS MT GO DF"]):
    for _uf in _ufs.split():
        UF2REG[_uf] = _reg


# ------------------------------ download ------------------------------
def _get(url):
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 (pesquisa academica)",
                                               "Accept-Encoding": "identity"})
    with urllib.request.urlopen(req, timeout=180) as r:
        dados = r.read()
    return gzip.decompress(dados) if dados[:2] == b"\x1f\x8b" else dados


def obter_arquivo():
    os.makedirs(CACHE, exist_ok=True)
    for pasta in (CACHE, BASE):
        loc = os.path.join(pasta, ARQUIVO)
        if os.path.exists(loc):
            print(f"Usando o arquivo local: {loc}")
            return loc
    nome = ARQUIVO
    try:                                    # procura no índice da pasta do FTP, caso o nome tenha mudado
        html = _get(PASTA_URL).decode("utf-8", "ignore")
        achados = [h for h in re.findall(r'href="([^"]+\.xlsx)"', html, re.I)
                   if "tab1" in h.lower() and "idade" in h.lower()]
        if achados:
            nome = achados[0]
    except Exception as e:                  # noqa: BLE001
        print(f"   (não consegui listar a pasta do FTP: {e}; tento o nome padrão)")
    url = PASTA_URL + nome
    destino = os.path.join(CACHE, ARQUIVO)
    for tentativa in range(1, 4):
        try:
            print(f"Baixando {url} ...", flush=True)
            with open(destino + ".tmp", "wb") as f:
                f.write(_get(url))
            os.replace(destino + ".tmp", destino)
            print(f"   salvo em {destino}")
            return destino
        except Exception as e:              # noqa: BLE001
            print(f"   ! falha ({tentativa}/3): {e}", flush=True)
            time.sleep(3)
    sys.exit(f"Não consegui baixar. Baixe manualmente {url} e coloque o arquivo em {CACHE}")


# ------------------------------ leitura da planilha ------------------------------
def _ano(v):
    try:
        f = float(str(v).strip().replace(",", "."))
    except ValueError:
        return None
    if f != f or f in (float("inf"), float("-inf")):      # NaN / infinito
        return None
    return int(f) if f == int(f) and 2000 <= f <= 2070 else None


def _num(v):
    try:
        f = float(str(v).strip().replace(" ", ""))
    except ValueError:
        return None
    return None if f != f else f


def interpretar(raw):
    """raw = planilha inteira sem cabeçalho (DataFrame de objetos). Devolve (por_uf, brasil, diagnostico)."""
    h = None
    for i in range(min(len(raw), 60)):
        anos = [j for j, v in enumerate(raw.iloc[i]) if _ano(v)]
        if len(anos) >= 10:
            h = i
            break
    if h is None:
        raise ValueError("não achei a linha de cabeçalho com os anos 2000...2070")
    cols_ano = {j: _ano(v) for j, v in enumerate(raw.iloc[h]) if _ano(v)}
    primeira = min(cols_ano)
    nomes = []
    for j in range(primeira):
        a = raw.iat[h, j]
        b = raw.iat[h - 1, j] if h > 0 else ""
        nomes.append((f"{'' if pd.isna(a) else a} {'' if pd.isna(b) else b}").strip().lower())
    c_idade = next((j for j, n in enumerate(nomes) if "idade" in n), None)
    c_sexo = next((j for j, n in enumerate(nomes) if "sexo" in n), None)
    if c_idade is None or c_sexo is None:
        raise ValueError(f"não achei as colunas de idade/sexo (colunas à esquerda dos anos: {nomes})")
    outras = [j for j in range(primeira) if j not in (c_idade, c_sexo)]

    por_uf, brasil, sexos = {}, {}, set()
    n_linhas = 0
    for i in range(h + 1, len(raw)):
        linha = raw.iloc[i]
        sexo = str(linha.iloc[c_sexo]).strip().lower()
        sexos.add(sexo)
        if not sexo.startswith("ambos"):
            continue
        rot = str(linha.iloc[c_idade]).strip()
        total = rot.lower().startswith("total")
        m = re.match(r"(\d+)", rot)
        if not total and not m:
            continue
        idade = None if total else int(m[1])
        textos = [str(linha.iloc[j]).strip() for j in outras if not pd.isna(linha.iloc[j])]
        uf = next((t.upper() for t in textos if t.upper() in UF2REG), None)
        eh_brasil = any(t.lower() == "brasil" or t.upper() == "BR" for t in textos)
        if not uf and not eh_brasil:
            continue
        n_linhas += 1
        for j, ano in cols_ano.items():
            if not ANO_INI <= ano <= ANO_FIM:
                continue
            v = _num(linha.iloc[j])
            if v is None:
                continue
            alvo = brasil if eh_brasil else por_uf.setdefault(uf, {})
            d = alvo.setdefault(ano, {"soma": 0.0, "idosos": 0.0, "total_linha": None})
            if total:
                d["total_linha"] = v
            else:
                d["soma"] += v
                if idade >= IDADE_CORTE:
                    d["idosos"] += v
    if not por_uf:
        raise ValueError(f"nenhuma linha de UF com sexo 'Ambos' (sexos vistos: {sorted(sexos)[:8]})")
    return por_uf, brasil, f"cabeçalho na linha {h + 1}; colunas: idade='{nomes[c_idade]}', sexo='{nomes[c_sexo]}'; " \
                           f"{n_linhas} linhas lidas"


def ler_planilha(caminho):
    xl = pd.ExcelFile(caminho)
    erros = []
    for aba in xl.sheet_names:
        raw = xl.parse(aba, header=None, dtype=object)
        try:
            return interpretar(raw) + (aba,)
        except ValueError as e:
            erros.append(f"aba '{aba}': {e}")
    raise SystemExit("Não consegui interpretar a planilha:\n  " + "\n  ".join(erros)
                     + "\nMe mande esta mensagem e as 10 primeiras linhas da planilha.")


def main():
    caminho = obter_arquivo()
    por_uf, brasil, diag, aba = ler_planilha(caminho)
    print(f"Aba '{aba}': {diag}.")
    if len(por_uf) != 27:
        sys.exit(f"Esperava 27 UFs e achei {len(por_uf)}: {sorted(por_uf)}")

    linhas = []
    for ano in range(ANO_INI, ANO_FIM + 1):
        for grupo, chave in (("Geral", "soma"), ("60+", "idosos")):
            tot_reg = 0
            for reg in REGIOES:
                v = int(round(sum(por_uf[uf][ano][chave] for uf in por_uf if UF2REG[uf] == reg)))
                linhas.append((ano, reg, grupo, v))
                tot_reg += v
            linhas.append((ano, "Brasil (Total)", grupo, tot_reg))
    df = pd.DataFrame(linhas, columns=["Ano", "Região", "Grupo", "População"])
    df.to_csv(SAIDA, sep=";", index=False, encoding="utf-8-sig")
    print(f"\nSalvo: {SAIDA}  ({len(df)} linhas, {ANO_INI}-{ANO_FIM})\n")

    print("Conferências:")
    problemas = 0
    for ano in range(ANO_INI, ANO_FIM + 1):
        soma_ufs = sum(por_uf[uf][ano]["soma"] for uf in por_uf)
        b = brasil.get(ano)
        if b:
            ref = b["total_linha"] if b["total_linha"] is not None else b["soma"]
            if abs(soma_ufs - ref) > 0.0001 * ref:
                problemas += 1
                print(f"   ATENÇÃO {ano}: soma das UFs {soma_ufs:,.0f} x linha do Brasil {ref:,.0f}")
        for uf in por_uf:
            t = por_uf[uf][ano]["total_linha"]
            if t is not None and abs(t - por_uf[uf][ano]["soma"]) > 0.0001 * t:
                problemas += 1
                print(f"   ATENÇÃO {ano} {uf}: soma das idades {por_uf[uf][ano]['soma']:,.0f} x linha Total {t:,.0f}")
    t22 = int(df[(df.Ano == 2022) & (df.Grupo == "Geral") & (df.Região == "Brasil (Total)")]["População"].iloc[0])
    i22 = int(df[(df.Ano == 2022) & (df.Grupo == "60+") & (df.Região == "Brasil (Total)")]["População"].iloc[0])
    pct = i22 / t22 * 100
    ok_t, ok_p = 202_500_000 <= t22 <= 203_700_000, 15.4 <= pct <= 16.2
    print(f"   2022: total {t22:,} ({'OK, perto dos 203,08 milhões do Censo' if ok_t else 'ATENÇÃO: fora do esperado'}) | "
          f"60+ {i22:,} = {pct:.2f}% ({'OK, Censo 2022: 15,81%' if ok_p else 'ATENÇÃO: fora do esperado'})".replace(",", "."))
    if problemas == 0 and ok_t and ok_p:
        print("   Tudo certo.")
    print("\nTabela Brasil (soma das regiões):")
    print(df[df["Região"] == "Brasil (Total)"].pivot(index="Ano", columns="Grupo", values="População").to_string())
    print(f"\nExtraído em {dt.date.today():%d/%m/%Y}. Próximo passo: py -3.12 montar_dados_projeto.py")


if __name__ == "__main__":
    main()
