"""
Gera o que a Vercel publica (pasta public/) a partir de dados/:

- dados/resumo.json      -> números do mês, o que mudou, histórico, qualidade
- dados/catalogo.json    -> produtos para a busca e o conferidor (compacto)
- dados/pmc/<aliq>.json  -> PMC de cada produto numa alíquota (a página só
                            baixa a alíquota escolhida)
- lista-cmed.xlsx        -> planilha para baixar

Tudo sai idêntico a cada execução com os mesmos dados (o CI confere isso).
"""
import io
import json
import re
import zipfile
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

PUBLIC = Path(__file__).resolve().parent.parent / "public"
DATA_FIXA = datetime(2026, 1, 1)
MARCA = "1F4E5A"
FORA = "FBE8E6"

# ordem de exibição das alíquotas nos filtros e na planilha
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


# ------------------------------------------------------------------ planilha

def _cel(ws, valor, fmt=None, negrito=False, fundo=None, texto=False):
    c = WriteOnlyCell(ws, value=valor)
    if texto and isinstance(valor, str):
        c.data_type = "s"  # nunca vira fórmula, mesmo começando com "="
    if fmt:
        c.number_format = fmt
    if negrito:
        c.font = Font(bold=True, color="FFFFFF" if fundo == MARCA else None)
    if fundo:
        c.fill = PatternFill("solid", fgColor=fundo)
    return c


def _reais(centavos):
    return None if centavos is None else centavos / 100


def _impressao(ws, caber_na_largura=True):
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = "9"  # A4
    if caber_na_largura:
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0


def _aba_tabela(wb, titulo, descricao, cabecalho, linhas, larguras, formatos, destaque=None, caber_na_largura=True):
    ws = wb.create_sheet(titulo)
    _impressao(ws, caber_na_largura)
    ws.print_title_rows = "3:3"
    for i, w in enumerate(larguras, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A4"
    ws.append([_cel(ws, titulo, negrito=True, texto=True)])
    ws.append([_cel(ws, descricao)])
    ws.append([_cel(ws, h, negrito=True, fundo=MARCA) for h in cabecalho])
    for linha in linhas:
        fundo = FORA if destaque and destaque(linha) else None
        ws.append([_cel(ws, v, fmt=f, fundo=fundo, texto=True) for v, f in zip(linha, formatos)])
    ws.auto_filter.ref = f"A3:{get_column_letter(len(cabecalho))}{max(3, len(linhas) + 3)}"
    return ws


def gerar_planilha(caminho, data, anterior, fonte, aliquotas, produtos, comparacao):
    wb = Workbook(write_only=True)
    br = lambda d: f"{d[8:]}/{d[5:7]}/{d[:4]}" if d else "—"  # noqa: E731
    R = '"R$" #,##0.00'

    ws = wb.create_sheet("Leia-me")
    _impressao(ws)
    ws.column_dimensions["A"].width = 120
    for texto, negrito in [
        (f"Lista de preços de medicamentos CMED — publicada em {br(data)}", True),
        (f"Fonte: {fonte}", False),
        (f"Comparada com a lista de {br(anterior)}." if anterior else "Primeira lista coletada.", False),
        ("", False),
        ("Como ler", True),
        ("PMC = preço máximo ao consumidor: o maior preço que farmácia e drogaria podem cobrar.", False),
        ("O PMC depende da alíquota de ICMS do estado. Pará: 19%. Confira a do seu estado — a própria CMED avisa que isso é do comerciante.", False),
        ("Colunas 'ALC' valem para Áreas de Livre Comércio (cidades no AP, AM, RO, RR e AC).", False),
        ("Isento de ICMS = Sim: onde a isenção vale, o teto é o PMC 0%.", False),
        ("Uso hospitalar = Sim: não tem PMC — não pode ser vendido em farmácia pelo PMC.", False),
        ("'Fora do normal' nas mudanças: preço multiplicou ou dividiu por 3 ou mais. Pode ser correção ou erro da lista — confira na fonte.", False),
        ("", False),
        ("Planilha gerada automaticamente pelo robô monitor-precos-cmed a partir do arquivo oficial da Anvisa.", False),
    ]:
        ws.append([_cel(ws, texto, texto=True, negrito=negrito)])

    if comparacao:
        linhas = []
        for m in comparacao["mudancas"]:
            p = m["produto"]
            situacao = "Fora do normal — confira" if m.get("fora_do_normal") else m["tipo"].capitalize()
            linhas.append([p["produto"], p["apresentacao"], p["laboratorio"], p.get("regime", ""),
                           _reais(m["antes"]), _reais(m["depois"]), m.get("variacao"), situacao, p["ggrem"], p["ean1"]])
        _aba_tabela(wb, "Mudanças de preço",
                    f"PMC a 19% (Pará) na lista de {br(anterior)} e na de {br(data)}. Linhas em vermelho: variação fora do normal.",
                    ["Produto", "Apresentação", "Laboratório", "Regime", "PMC antes", "PMC agora", "Variação",
                     "Situação", "Código GGREM", "Código de barras"],
                    linhas, [22, 48, 34, 11, 13, 13, 11, 24, 18, 16],
                    [None, None, None, None, R, R, "+0.00%;-0.00%", None, "@", "@"],
                    destaque=lambda l: l[7].startswith("Fora"))
        for nome, grupo, desc in [
            ("Novos na lista", comparacao["novos"], f"Apresentações que entraram na lista de {br(data)}."),
            ("Saíram da lista", comparacao["sairam"], f"Apresentações da lista de {br(anterior)} que não estão na de {br(data)}."),
        ]:
            _aba_tabela(wb, nome, desc,
                        ["Produto", "Apresentação", "Laboratório", "PMC 19%", "Uso hospitalar", "Código GGREM", "Código de barras"],
                        [[p["produto"], p["apresentacao"], p["laboratorio"], _reais(p["pmc"].get("19")),
                          "Sim" if p["hospitalar"] else "Não", p["ggrem"], p["ean1"]] for p in grupo],
                        [22, 50, 34, 13, 14, 18, 16], [None, None, None, R, None, "@", "@"])

    precos = [a for a in aliquotas if a != "sem"]
    _aba_tabela(wb, "Lista completa", f"Todas as apresentações da lista de {br(data)}, com o PMC em cada alíquota de ICMS.",
                ["Produto", "Apresentação", "Substância", "Laboratório", "Tipo", "Regime", "Tarja", "Uso hospitalar",
                 "Isento de ICMS", "Código GGREM", "Código de barras"] + [f"PMC {rotulo_aliquota(a).replace(' · Área de Livre Comércio', ' ALC')}" for a in precos],
                [[p["produto"], p["apresentacao"], p["substancia"], p["laboratorio"], p.get("tipo", ""), p.get("regime", ""),
                  p.get("tarja", ""), "Sim" if p["hospitalar"] else "Não", "Sim" if p["icms0"] else "Não", p["ggrem"], p["ean1"]]
                 + [_reais(p["pmc"].get(a)) for a in precos] for p in produtos],
                [22, 44, 30, 30, 12, 10, 18, 9, 9, 18, 16] + [11] * len(precos),
                [None] * 9 + ["@", "@"] + [R] * len(precos), caber_na_largura=False)

    wb.properties.creator = "monitor-precos-cmed"
    wb.properties.title = f"Lista CMED {br(data)}"
    wb.properties.created = wb.properties.modified = DATA_FIXA
    buf = io.BytesIO()
    wb.save(buf)
    _zip_deterministico(buf, caminho)


def _zip_deterministico(buf, caminho):
    """O openpyxl carimba a hora atual no zip e em docProps; aqui tudo vira data fixa."""
    buf.seek(0)
    caminho.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(buf) as origem, zipfile.ZipFile(caminho, "w", zipfile.ZIP_DEFLATED) as destino:
        for item in origem.infolist():
            conteudo = origem.read(item.filename)
            if item.filename == "docProps/core.xml":
                conteudo = re.sub(rb"(<dcterms:modified[^>]*>)[^<]*", rb"\g<1>2026-01-01T00:00:00Z", conteudo)
            info = zipfile.ZipInfo(item.filename, date_time=(2026, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            destino.writestr(info, conteudo)
