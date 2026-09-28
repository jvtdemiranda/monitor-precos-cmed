"""
O robô.

    python robo/rodar.py                 # confere a página da Anvisa; se saiu lista nova, baixa, limpa e publica
    python robo/rodar.py --historico 12  # idem, e busca até 12 listas anteriores (primeira carga)
    python robo/rodar.py --arquivo X.xlsx --data AAAA-MM-DD   # processa um arquivo local
    python robo/rodar.py --so-publicar   # só regenera public/ a partir de dados/ (sem internet)

Nada é publicado se algo parecer errado: arquivo que não é planilha,
cabeçalho não encontrado, colunas faltando, preços ilegíveis ou uma lista
bem menor que a anterior. Nesses casos o robô para com erro (o GitHub avisa
por e-mail) e o site continua com a última lista boa.
"""
import argparse
import shutil
import dataclasses
import sys
import tempfile
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import coleta  # noqa: E402
import comparar  # noqa: E402
import dados  # noqa: E402
import lista  # noqa: E402
import publicar  # noqa: E402

# lista nova com menos de 80% dos produtos da anterior: provavelmente arquivo
# cortado ou com problema — não publica
QUEDA_MAXIMA = 0.8


def processar(conteudo, data_link, url, raiz=dados.RAIZ):
    sha = coleta.conferir_planilha(conteudo)
    with tempfile.NamedTemporaryFile(suffix=".xlsx") as tmp:
        tmp.write(conteudo)
        tmp.flush()
        info, produtos, q = lista.ler_lista(tmp.name)
    data = info["publicada_em"] or data_link
    if data_link and info["publicada_em"] and info["publicada_em"] != data_link:
        print(f"  aviso: link diz {data_link}, arquivo diz publicada em {info['publicada_em']} — vale a do arquivo")
    anteriores = [d for d in dados.publicacoes(raiz) if d < data]
    if anteriores:
        n_ant = len(dados.ler_precos(anteriores[-1], raiz))
        if len(produtos) < QUEDA_MAXIMA * n_ant:
            raise lista.ListaInvalida(
                f"lista de {data} tem {len(produtos)} produtos contra {n_ant} da anterior — não publico")
    dados.gravar_lista(data, info["aliquotas"], produtos, raiz)
    fontes = dados.ler_fontes(raiz)
    fontes[data] = {"url": url, "sha256": sha, "bytes": len(conteudo), "produtos": len(produtos),
                    "coletada_em": date.today().isoformat(), "qualidade": dataclasses.asdict(q)}
    dados.gravar_fontes(fontes, raiz)
    print(f"  {data}: {len(produtos)} produtos, {q.sem_pmc_hospitalar} de uso hospitalar, "
          f"{q.ean_invalido} códigos de barras inválidos")
    return data


def publicar_tudo(raiz=dados.RAIZ, destino=publicar.PUBLIC):
    datas = dados.listas_completas(raiz)
    if not datas:
        raise SystemExit("nenhuma lista em dados/ — rode o robô primeiro")
    atual = datas[-1]
    anterior = datas[-2] if len(datas) > 1 else None
    aliquotas, produtos = dados.ler_lista_limpa(atual, raiz)
    comp = None
    if anterior:
        _, prods_ant = dados.ler_lista_limpa(anterior, raiz)
        comp = comparar.comparar(prods_ant, produtos)
    fontes = dados.ler_fontes(raiz)
    historico = comparar.resumo_historico(dados.publicacoes(raiz), lambda d: dados.ler_precos(d, raiz))
    fonte = fontes.get(atual, {})
    publicar.gerar_json(destino, atual, fonte.get("url", ""), aliquotas, produtos, comp, anterior,
                        historico, fonte.get("qualidade", {}))
    publicar.gerar_planilha(destino / "lista-cmed.xlsx", atual, anterior, fonte.get("url", ""),
                            aliquotas, produtos, comp)
    exemplo = Path(__file__).resolve().parent.parent / "exemplos" / "tabela-drogaria-exemplo.csv"
    if exemplo.exists():
        shutil.copyfile(exemplo, destino / "exemplo.csv")
    n = len(comp["mudancas"]) if comp else 0
    print(f"publicado: lista de {atual} ({len(produtos)} produtos, {n} mudanças de preço)")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--historico", type=int, default=0, help="buscar até N listas anteriores")
    ap.add_argument("--arquivo", help="processar um xlsx local em vez de baixar")
    ap.add_argument("--data", help="data da lista local (AAAA-MM-DD), se o arquivo não disser")
    ap.add_argument("--so-publicar", action="store_true", help="só regenerar public/ a partir de dados/")
    args = ap.parse_args()

    if args.so_publicar:
        publicar_tudo()
        return
    if args.arquivo:
        processar(Path(args.arquivo).read_bytes(), args.data, f"arquivo local: {Path(args.arquivo).name}")
        publicar_tudo()
        return

    # "já temos" pelo endereço do arquivo, não pela data: se a Anvisa
    # republicar uma lista corrigida no mesmo dia, o nome do arquivo muda
    ja_temos = {f["url"] for f in dados.ler_fontes().values()}
    print(f"Conferindo {coleta.PAGINA}")
    html = coleta.baixar(coleta.PAGINA, prazo=60).decode("utf-8", "replace")
    pendentes = [coleta.lista_mais_recente(html)]
    if args.historico:
        print(f"Buscando até {args.historico} listas anteriores em {coleta.ANTERIORES}")
        html_ant = coleta.baixar(coleta.ANTERIORES, prazo=120).decode("utf-8", "replace")
        pendentes += coleta.listas_na_pagina(html_ant, coleta.ANTERIORES)[: args.historico]
    novas = sorted({(d, u) for d, u in pendentes if u not in ja_temos})
    if not novas:
        print(f"Nenhuma lista nova (a mais recente, de {pendentes[0][0]}, já está em dados/).")
        return
    for data_link, url in novas:  # da mais antiga para a mais nova
        print(f"Baixando lista de {data_link}")
        processar(coleta.baixar(url), data_link, url)
    publicar_tudo()


if __name__ == "__main__":
    main()
