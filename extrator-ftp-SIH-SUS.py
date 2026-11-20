import ftplib, os, re, sys, time, json, threading
from concurrent.futures import ThreadPoolExecutor, as_completed
import numpy as np
import pandas as pd
from datasus_dbc import decompress

PROC = "0406010650"
ANO_INI, ANO_FIM = 2015, 2026     # 2026 entra para captar AIHs de 2025 processadas depois e 2015 por garantia.
WORKERS = 4                        # downloads simultâneos (se o FTP recusar conexões, baixe para 2)
LIMITE = None                      # coloque 5 para testar só com 5 arquivos

HOST = "ftp.datasus.gov.br"
PASTA = "/dissemin/publicos/SIHSUS/200801_/Dados"
CAMPOS = ["N_AIH", "PROC_REA", "IDADE", "COD_IDADE", "VAL_TOT", "MUNIC_RES",
          "MUNIC_MOV", "DT_INTER", "DT_SAIDA", "ANO_CMPT", "MES_CMPT", "IDENT"]
PADRAO = re.compile(r"^RD([A-Z]{2})(\d{2})(\d{2})\.DBC$")

BASE = os.path.dirname(os.path.abspath(__file__))
CACHE = os.path.join(BASE, "cache")
TMP = os.path.join(BASE, "tmp")

_local = threading.local()
ativos = {}
parar = threading.Event()


def conectar():
    ftp = ftplib.FTP(HOST, timeout=60)
    ftp.login()
    ftp.cwd(PASTA)
    ftp.voidcmd("TYPE I")
    return ftp


def get_ftp():
    if getattr(_local, "ftp", None) is None:
        _local.ftp = conectar()
    return _local.ftp


def reset_ftp():
    try:
        _local.ftp.close()
    except Exception:
        pass
    _local.ftp = None


def listar():
    """Retorna {NOME_MAIÚSCULO: (nome_original, tamanho_bytes)}"""
    ftp = conectar()
    arquivos = {}
    try:
        for nome, fatos in ftp.mlsd(facts=["size"]):
            arquivos[nome.upper()] = (nome, int(fatos.get("size", 0)))
    except Exception:
        for nome in ftp.nlst():
            nome = os.path.basename(nome)
            arquivos[nome.upper()] = (nome, 0)
    try:
        ftp.quit()
    except Exception:
        pass
    return arquivos


def baixar(nome, nome_remoto, destino, tam):
    for tentativa in range(1, 4):
        try:
            ftp = get_ftp()
            baixado = 0
            with open(destino, "wb") as f:
                def cb(bloco):
                    nonlocal baixado
                    f.write(bloco)
                    baixado += len(bloco)
                    ativos[nome] = ("baixando", baixado, tam)
                ftp.retrbinary(f"RETR {nome_remoto}", cb, blocksize=1 << 16)
            return
        except Exception as e:
            reset_ftp()
            print(f"   ! {nome}: falha no download (tentativa {tentativa}/3): {e}", flush=True)
            time.sleep(3)
    raise RuntimeError("download falhou 3 vezes")


def filtrar_dbf(caminho):
    """Lê o DBF em bloco e devolve só as linhas do procedimento (sem loop por linha)."""
    with open(caminho, "rb") as f:
        cab = f.read(32)
        tam_cab = int.from_bytes(cab[8:10], "little")
        tam_reg = int.from_bytes(cab[10:12], "little")
        campos, off = {}, 1
        while True:
            d = f.read(32)
            if not d or d[0] == 0x0D:
                break
            nome = d[:11].split(b"\x00")[0].decode("ascii")
            campos[nome] = (off, d[16])
            off += d[16]
        f.seek(tam_cab)
        bruto = np.fromfile(f, dtype=np.uint8)
    n = len(bruto) // tam_reg
    if n == 0:
        return []
    dados = bruto[: n * tam_reg].reshape(n, tam_reg)
    o, t = campos["PROC_REA"]
    alvo = np.frombuffer(PROC.encode().ljust(t), dtype=np.uint8)
    mask = (dados[:, o:o + t] == alvo).all(axis=1) & (dados[:, 0] != 0x2A)
    saida = []
    for linha in dados[mask]:
        b = linha.tobytes()
        saida.append({c: b[campos[c][0]:campos[c][0] + campos[c][1]].decode("latin-1").strip()
                      for c in CAMPOS if c in campos})
    return saida


def processar(nome, nome_remoto, tam):
    cache_f = os.path.join(CACHE, nome + ".json")
    if os.path.exists(cache_f):
        with open(cache_f, encoding="utf-8") as f:
            return json.load(f), True
    dbc = os.path.join(TMP, nome + ".dbc")
    dbf = os.path.join(TMP, nome + ".dbf")
    try:
        baixar(nome, nome_remoto, dbc, tam)
        ativos[nome] = ("descompactando", 0, 0)
        decompress(dbc, dbf)
        ativos[nome] = ("filtrando", 0, 0)
        regs = filtrar_dbf(dbf)
        for r in regs:
            r["ARQUIVO"] = nome
        with open(cache_f + ".tmp", "w", encoding="utf-8") as f:
            json.dump(regs, f)
        os.replace(cache_f + ".tmp", cache_f)
        return regs, False
    finally:
        ativos.pop(nome, None)
        for p in (dbc, dbf):
            try:
                os.remove(p)
            except OSError:
                pass


def heartbeat():
    while not parar.wait(10):
        itens = []
        for n, (est, b, t) in list(ativos.items()):
            if est == "baixando":
                itens.append(f"{n} baixando " + (f"{b / t * 100:.0f}%" if t else f"{b / 1e6:.1f}MB"))
            else:
                itens.append(f"{n} {est}")
        if itens:
            print("   ... em andamento: " + " | ".join(itens), flush=True)


def fmt(seg):
    seg = int(seg)
    return f"{seg // 3600}h{(seg % 3600) // 60:02d}min"


def main():
    os.makedirs(CACHE, exist_ok=True)
    os.makedirs(TMP, exist_ok=True)
    print("Conectando ao FTP do DATASUS e listando arquivos...", flush=True)
    remoto = listar()

    alvo = []
    for up, (orig, tam) in remoto.items():
        m = PADRAO.match(up)
        if not m:
            continue
        uf, aa, mm = m.groups()
        ano, mes = 2000 + int(aa), int(mm)
        if ANO_INI <= ano <= ANO_FIM and 1 <= mes <= 12:
            alvo.append((ano, mes, uf, orig, tam))
    alvo.sort()
    if LIMITE:
        alvo = alvo[:LIMITE]

    pendentes = [a for a in alvo if not os.path.exists(os.path.join(CACHE, a[3][:-4] + ".json"))]
    gb = sum(a[4] for a in pendentes) / 1e9
    print(f"{len(alvo)} arquivos no período; {len(alvo) - len(pendentes)} já em cache; "
          f"{len(pendentes)} a baixar (~{gb:.1f} GB).", flush=True)

    threading.Thread(target=heartbeat, daemon=True).start()
    ex = ThreadPoolExecutor(WORKERS)
    futs = {ex.submit(processar, a[3][:-4], a[3], a[4]): a[3][:-4] for a in alvo}
    total, feitos, reais, achados, erros = len(alvo), 0, 0, 0, []
    t0 = time.time()
    try:
        for fut in as_completed(futs):
            nome = futs[fut]
            feitos += 1
            try:
                regs, do_cache = fut.result()
            except Exception as e:
                erros.append(nome)
                print(f"[{feitos}/{total}] ERRO em {nome}: {e}", flush=True)
                continue
            achados += len(regs)
            if not do_cache:
                reais += 1
            restantes = total - feitos
            eta = (time.time() - t0) / reais * restantes if reais else 0
            print(f"[{feitos}/{total} | {feitos / total * 100:5.1f}%] {nome} "
                  f"{'(cache)' if do_cache else 'ok'} | {len(regs)} registros "
                  f"| acumulado: {achados} | decorrido {fmt(time.time() - t0)} "
                  f"| faltam ~{fmt(eta)}", flush=True)
    except KeyboardInterrupt:
        print("\nInterrompido. O progresso está salvo; rode de novo para continuar.", flush=True)
        os._exit(1)
    parar.set()
    ex.shutdown()

    todos = []
    for a in alvo:
        p = os.path.join(CACHE, a[3][:-4] + ".json")
        if os.path.exists(p):
            with open(p, encoding="utf-8") as f:
                todos += json.load(f)
    pd.DataFrame(todos).to_csv(os.path.join(BASE, "marcapasso_bruto.csv"), index=False)
    print(f"\nPronto: {len(todos)} registros salvos em marcapasso_bruto.csv", flush=True)
    if erros:
        print(f"ATENÇÃO: {len(erros)} arquivos falharam: {erros}\nRode o script de novo para tentar só esses.")


if __name__ == "__main__":
    main()