"""
Gera o que a Vercel publica (pasta public/) a partir de dados/:

- dados/resumo.json      -> números do mês, o que mudou, histórico, qualidade
- dados/catalogo.json    -> produtos para a busca e o conferidor (compacto)
- dados/pmc/<aliq>.json  -> PMC de cada produto numa alíquota (a página só
                            baixa a alíquota escolhida)
(a planilha para baixar, lista-cmed.xlsx, é feita em planilha.py)

Tudo sai idêntico a cada execução com os mesmos dados (o CI confere isso).
"""
import json
from pathlib import Path

PUBLIC = Path(__file__).resolve().parent.parent / "public"


def rotulo_aliquota(chave):
    base = chave.replace("_alc", "")
    nome = "sem impostos" if base == "sem" else base.replace(".", ",") + "%"
    return nome + (" · Área de Livre Comércio" if chave.endswith("_alc") else "")


def _json(caminho, obj):
    caminho.parent.mkdir(parents=True, exist_ok=True)
    caminho.write_text(json.dumps(obj, ensure_ascii=False, separators=(",", ":")) + "\n", encoding="utf-8")


def _flags(p):
    return (1 if p["hospitalar"] else 0) | (2 if p["icms0"] else 0) | (4 if p.get("recursal") else 0)


def _item_mudanca(m):
    p = m["produto"]
    return {"g": p["ggrem"], "produto": p["produto"], "apresentacao": p["apresentacao"],
            "laboratorio": p["laboratorio"], "regime": p.get("regime", ""), "antes": m["antes"],
            "depois": m["depois"], "tipo": m["tipo"], "variacao": m.get("variacao"),
            "fora": m.get("fora_do_normal", False)}


def _item_simples(p):
    return {"g": p["ggrem"], "produto": p["produto"], "apresentacao": p["apresentacao"],
            "laboratorio": p["laboratorio"], "pmc": p["pmc"].get("19"), "hospitalar": p["hospitalar"]}


def gerar_json(destino, data, fonte, aliquotas, produtos, comparacao, anterior, historico, qualidade):
    labs = sorted({p["laboratorio"] for p in produtos})
    subs = sorted({p["substancia"] for p in produtos})
    tipos = sorted({p.get("tipo", "") for p in produtos})
    li, si, ti = ({v: i for i, v in enumerate(x)} for x in (labs, subs, tipos))
    itens = []
    for p in produtos:
        eans = [e for e in (p["ean1"], p.get("ean2", ""), p.get("ean3", "")) if e]
        itens.append([p["ggrem"], eans, p["produto"], p["apresentacao"], si[p["substancia"]],
                      li[p["laboratorio"]], ti[p.get("tipo", "")], _flags(p)])
    _json(destino / "dados" / "catalogo.json", {
        "publicada_em": data,
        "campos": ["ggrem", "eans", "produto", "apresentacao", "substancia", "laboratorio", "tipo", "flags"],
        "flags": {"1": "uso restrito a hospitais (sem PMC)", "2": "isento de ICMS", "4": "preço em análise de recurso"},
        "laboratorios": labs, "substancias": subs, "tipos": tipos, "itens": itens,
    })
    for a in aliquotas:
        _json(destino / "dados" / "pmc" / f"{a}.json", [p["pmc"].get(a) for p in produtos])

    com_pmc = sum(1 for p in produtos if p["pmc"].get("19") is not None)
    _json(destino / "dados" / "resumo.json", {
        "publicada_em": data,
        "anterior": anterior,
        "fonte": fonte,
        "aliquotas": [{"chave": a, "rotulo": rotulo_aliquota(a)} for a in aliquotas if a != "sem"],
        "produtos": len(produtos),
        "com_pmc": com_pmc,
        "hospitalar": sum(1 for p in produtos if p["hospitalar"]),
        "isentos": sum(1 for p in produtos if p["icms0"]),
        "qualidade": qualidade,
        "historico": historico,
        "mudancas": [_item_mudanca(m) for m in comparacao["mudancas"]] if comparacao else [],
        "novos": [_item_simples(p) for p in comparacao["novos"]] if comparacao else [],
        "sairam": [_item_simples(p) for p in comparacao["sairam"]] if comparacao else [],
    })
