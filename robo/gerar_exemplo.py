"""
Gera exemplos/tabela-drogaria-exemplo.csv: a tabela de preços de uma
drogaria fictícia, montada com produtos reais da lista de 09/09/2026, para
experimentar o conferidor. Tem de propósito os casos que aparecem na vida
real: preço acima do PMC, código de barras estragado pelo Excel, produto de
uso hospitalar, produto isento de ICMS, código que não está na lista,
preço em formato "R$ 1.234,56" e uma coluna de código interno antes do EAN.

Rode uma vez (a partir da lista completa em dados/); o arquivo gerado é
commitado e não muda sozinho.
"""
import csv
import random
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import dados  # noqa: E402
import lista  # noqa: E402

DATA = "2026-09-09"
SAIDA = Path(__file__).resolve().parent.parent / "exemplos" / "tabela-drogaria-exemplo.csv"
COMUNS = ["DIPIRONA", "LOSARTANA POTÁSSICA", "AMOXICILINA", "NIMESULIDA", "OMEPRAZOL", "SINVASTATINA",
          "CLORIDRATO DE METFORMINA", "IBUPROFENO", "PARACETAMOL", "AZITROMICINA", "HIDROCLOROTIAZIDA",
          "ATENOLOL", "LORATADINA", "DICLOFENACO SÓDICO", "CAPTOPRIL", "PREDNISONA", "FLUCONAZOL", "CETOPROFENO"]


def reais(centavos):
    return f"{centavos / 100:.2f}".replace(".", ",")


def main():
    rnd = random.Random(8)
    _, produtos = dados.ler_lista_limpa(DATA)
    contagem = {}
    for p in produtos:
        for e in (p["ean1"], p["ean2"], p["ean3"]):
            if e:
                contagem[e] = contagem.get(e, 0) + 1

    def bom(p):
        return (p["ean1"] and lista.gtin_valido(p["ean1"]) and contagem[p["ean1"]] == 1
                and p["pmc"]["19"] and not p["hospitalar"] and not p["icms0"] and p["tipo"] in ("Genérico", "Similar", "Novo"))

    escolhidos = []
    for sub in COMUNS:
        opcoes = sorted((p for p in produtos if p["substancia"] == sub and bom(p)), key=lambda p: p["ggrem"])
        if opcoes:
            escolhidos.append(opcoes[len(opcoes) // 2])

    linhas = []
    for i, p in enumerate(escolhidos):
        fator = rnd.uniform(0.72, 0.97)
        if i in (2, 7, 12):
            fator = rnd.uniform(1.04, 1.18)  # acima do PMC
        linhas.append([f"{1001 + i}", p["ean1"], f"{p['produto']} {p['apresentacao']}"[:70], reais(round(p["pmc"]["19"] * fator))])
    # preço grande no formato com "R$" e ponto de milhar
    caro = next(p for p in produtos if bom(p) and p["pmc"]["19"] > 150000)
    linhas.append(["1101", caro["ean1"], f"{caro['produto']} {caro['apresentacao']}"[:70], "R$ " + f"{caro['pmc']['19'] * 0.9 / 100:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")])
    # isento de ICMS com preço entre o PMC 0% e o PMC 19%
    isento = next(p for p in sorted(produtos, key=lambda p: p["ggrem"])
                  if p["icms0"] and p["pmc"]["19"] and p["pmc"]["0"] and p["pmc"]["19"] - p["pmc"]["0"] > 500
                  and p["ean1"] and contagem[p["ean1"]] == 1)
    linhas.append(["1102", isento["ean1"], f"{isento['produto']} {isento['apresentacao']}"[:70], reais((isento["pmc"]["0"] + isento["pmc"]["19"]) // 2)])
    # uso hospitalar (sem PMC)
    hosp = next(p for p in sorted(produtos, key=lambda p: p["ggrem"]) if p["hospitalar"] and p["ean1"] and contagem[p["ean1"]] == 1)
    linhas.append(["1103", hosp["ean1"], f"{hosp['produto']} {hosp['apresentacao']}"[:70], "89,90"])
    # código de barras que a planilha do Excel transformou em notação científica
    estragado = escolhidos[0]["ean1"]
    linhas.append(["1104", f"{estragado[0]},{estragado[1:6]}E+12", "PRODUTO COM CÓDIGO ESTRAGADO PELO EXCEL", "12,50"])
    # código de barras que não é de medicamento da lista (cosmético)
    linhas.append(["1105", "7891000100103", "PROTETOR SOLAR FPS 50 200ML", "59,90"])
    # preço em branco
    linhas.append(["1106", escolhidos[1]["ean1"], "PRODUTO SEM PREÇO CADASTRADO", ""])

    SAIDA.parent.mkdir(exist_ok=True)
    with SAIDA.open("w", encoding="utf-8", newline="") as f:
        w = csv.writer(f, delimiter=";", lineterminator="\r\n")
        w.writerow(["Código", "EAN", "Descrição", "Preço de venda"])
        w.writerows(linhas)
    print(f"{SAIDA} ({len(linhas)} produtos)")


if __name__ == "__main__":
    main()
