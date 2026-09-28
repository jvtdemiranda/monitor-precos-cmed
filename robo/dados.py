"""
O que o robô guarda no repositório (pasta dados/):

- listas/AAAA-MM-DD.csv.gz  -> lista limpa completa; só as 2 últimas ficam
                               (a atual e a anterior, para comparar). As
                               antigas continuam no histórico do git.
- precos/AAAA-MM-DD.csv.gz  -> só código e PMC de cada publicação, para o
                               histórico de meses (pequeno, fica para sempre)
- fontes.json               -> de onde veio cada lista: endereço, data,
                               sha256 do arquivo baixado e o que foi
                               corrigido nela

Os .gz são gravados sem data interna, para o mesmo conteúdo gerar sempre o
mesmo arquivo (e o git não ver mudança onde não houve).
"""
import csv
import gzip
import io
import json
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent / "dados"
TEXTOS = ["ggrem", "ean1", "ean2", "ean3", "produto", "apresentacao", "substancia", "laboratorio",
          "cnpj", "registro", "classe", "tipo", "regime", "tarja", "recursal"]
FLAGS = ["hospitalar", "icms0"]
LISTAS_GUARDADAS = 2


def _gravar_gz(caminho, texto):
    caminho.parent.mkdir(parents=True, exist_ok=True)
    buf = io.BytesIO()
    with gzip.GzipFile(fileobj=buf, mode="wb", mtime=0, filename="") as gz:
        gz.write(texto.encode("utf-8"))
    caminho.write_bytes(buf.getvalue())


def _ler_gz(caminho):
    with gzip.open(caminho, "rt", encoding="utf-8", newline="") as f:
        return list(csv.reader(f))


def _centavos(v):
    return int(v) if v != "" else None


def gravar_lista(data, aliquotas, produtos, raiz=RAIZ):
    cols = TEXTOS + FLAGS + [f"pmc_{a}" for a in aliquotas]
    out = io.StringIO()
    w = csv.writer(out, lineterminator="\n")
    w.writerow(cols)
    for p in produtos:
        w.writerow([p.get(c, "") for c in TEXTOS] + [int(p[f]) for f in FLAGS]
                   + ["" if p["pmc"].get(a) is None else p["pmc"][a] for a in aliquotas])
    _gravar_gz(raiz / "listas" / f"{data}.csv.gz", out.getvalue())

    out = io.StringIO()
    w = csv.writer(out, lineterminator="\n")
    w.writerow(["ggrem", "pmc_0", "pmc_19"])
    for p in produtos:
        w.writerow([p["ggrem"]] + ["" if p["pmc"].get(a) is None else p["pmc"][a] for a in ("0", "19")])
    _gravar_gz(raiz / "precos" / f"{data}.csv.gz", out.getvalue())

    # só as últimas listas completas ficam na pasta
    for velho in sorted((raiz / "listas").glob("*.csv.gz"))[:-LISTAS_GUARDADAS]:
        velho.unlink()


def ler_lista_limpa(data, raiz=RAIZ):
    linhas = _ler_gz(raiz / "listas" / f"{data}.csv.gz")
    cols = linhas[0]
    aliquotas = [c[4:] for c in cols if c.startswith("pmc_")]
    produtos = []
    for linha in linhas[1:]:
        d = dict(zip(cols, linha))
        p = {c: d[c] for c in TEXTOS}
        for f in FLAGS:
            p[f] = d[f] == "1"
        p["pmc"] = {a: _centavos(d[f"pmc_{a}"]) for a in aliquotas}
        produtos.append(p)
    return aliquotas, produtos


def ler_precos(data, raiz=RAIZ):
    """{ggrem: (pmc_0, pmc_19)} de uma publicação."""
    linhas = _ler_gz(raiz / "precos" / f"{data}.csv.gz")
    return {g: (_centavos(a), _centavos(b)) for g, a, b in linhas[1:]}


def publicacoes(raiz=RAIZ):
    """Datas com histórico de preços, da mais antiga pra mais nova."""
    return sorted(p.name[:10] for p in (raiz / "precos").glob("*.csv.gz"))


def listas_completas(raiz=RAIZ):
    return sorted(p.name[:10] for p in (raiz / "listas").glob("*.csv.gz"))


def ler_fontes(raiz=RAIZ):
    arq = raiz / "fontes.json"
    return json.loads(arq.read_text(encoding="utf-8")) if arq.exists() else {}


def gravar_fontes(fontes, raiz=RAIZ):
    raiz.mkdir(parents=True, exist_ok=True)
    (raiz / "fontes.json").write_text(
        json.dumps(dict(sorted(fontes.items())), ensure_ascii=False, indent=1) + "\n", encoding="utf-8")
