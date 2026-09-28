"""
O que mudou de uma publicação da lista para a seguinte.

A comparação usa o código GGREM (único por apresentação e estável entre
meses) e o PMC na alíquota de 19% — a do Pará. Nas outras alíquotas o
preço muda na mesma proporção.

"Fora do normal": o preço multiplicou ou dividiu por 3 ou mais. Fora do
reajuste anual de abril, o PMC de um medicamento regulado quase não muda;
um salto desses costuma ser correção de um erro da lista anterior ou um
erro novo (nos dados reais apareceu um preço exatamente 10 vezes maior de
um mês para o outro). O robô não decide qual é o certo: marca para alguém
conferir.
"""
FORA_DO_NORMAL = 3.0


def _variacao(antes, depois):
    return round(depois / antes - 1, 4)


def comparar(anteriores, atuais, aliquota="19"):
    """anteriores/atuais: listas de produtos (dicts de lista.ler_lista)."""
    ant = {p["ggrem"]: p for p in anteriores}
    atu = {p["ggrem"]: p for p in atuais}
    novos = [atu[g] for g in atu if g not in ant]
    sairam = [ant[g] for g in ant if g not in atu]
    mudancas = []
    for g, p in atu.items():
        if g not in ant:
            continue
        antes, depois = ant[g]["pmc"].get(aliquota), p["pmc"].get(aliquota)
        if antes == depois:
            continue
        m = {"produto": p, "antes": antes, "depois": depois}
        if antes is None:
            m["tipo"] = "passou a ter PMC"
        elif depois is None:
            m["tipo"] = "deixou de ter PMC"
        else:
            m["tipo"] = "subiu" if depois > antes else "desceu"
            m["variacao"] = _variacao(antes, depois)
            razao = max(depois / antes, antes / depois) if min(antes, depois) > 0 else float("inf")
            m["fora_do_normal"] = razao >= FORA_DO_NORMAL
        mudancas.append(m)
    ordem = {"subiu": 0, "desceu": 1, "passou a ter PMC": 2, "deixou de ter PMC": 3}
    mudancas.sort(key=lambda m: (not m.get("fora_do_normal", False), ordem[m["tipo"]],
                                 -abs(m.get("variacao", 0)), m["produto"]["produto"], m["produto"]["ggrem"]))
    chave = lambda p: (p["produto"], p["apresentacao"], p["ggrem"])  # noqa: E731
    return {"novos": sorted(novos, key=chave), "sairam": sorted(sairam, key=chave), "mudancas": mudancas}


def resumo_historico(publicacoes, ler_precos):
    """
    Uma linha por publicação com quantos preços mudaram em relação à
    anterior — mostra o reajuste anual de abril contra os meses calmos.
    """
    linhas, anterior = [], None
    for data in publicacoes:
        atual = ler_precos(data)
        item = {"data": data, "produtos": len(atual)}
        if anterior is not None:
            subiu = desceu = fora = 0
            for g, (_, pmc) in atual.items():
                if g not in anterior:
                    continue
                antes = anterior[g][1]
                if antes is None or pmc is None or antes == pmc:
                    continue
                if pmc > antes:
                    subiu += 1
                else:
                    desceu += 1
                if min(pmc, antes) == 0 or max(pmc / antes, antes / pmc) >= FORA_DO_NORMAL:
                    fora += 1
            item.update(subiu=subiu, desceu=desceu, fora_do_normal=fora,
                        novos=sum(1 for g in atual if g not in anterior),
                        sairam=sum(1 for g in anterior if g not in atual))
        linhas.append(item)
        anterior = atual
    return linhas
