"""
Planilha para baixar (public/lista-cmed.xlsx).

Feita para ser lida, não só filtrada: cada aba tem título, uma linha
dizendo o que é e uma tabela com poucas colunas. As 26 mil apresentações
ficam em duas abas separadas: a consulta do dia a dia (PMC do Pará e PMC
0%) e a matriz com todas as alíquotas, para outros estados.

As linhas das tabelas são separadas por um fio fino em vez de faixas
alternadas: continua arrumado depois que alguém ordena ou filtra, e aparece
igual no Excel, no celular e no Google Planilhas.

Gerada em modo "write-only" do openpyxl (rápido e leve para 26 mil linhas):
a ordem de escrita é a ordem das linhas, e estilos vão célula a célula.
"""
import io
import math
from copy import copy
import re
import zipfile
from datetime import datetime

from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

DATA_FIXA = datetime(2026, 1, 1)
FONTE = "Arial"
MARCA, TINTA, CINZA, CLARO = "1F4E5A", "1B1F23", "5F6B70", "8A9599"
VERMELHO, VERDE, AMBAR = "B3261E", "1E7F4F", "8A5A00"

REAIS = '"R$" #,##0.00'
INTEIRO = "#,##0"
# a cor da variação vem da fonte, não do formato: cor em formato de número
# ("[Color10]") sai diferente em cada programa — no LibreOffice, quase invisível
VARIACAO = '"▲ "0.0%;"▼ "0.0%;0.0%'
VARIACAO_PEQUENA = '"▲ "0.00%;"▼ "0.00%;0.00%'  # +0,04% não pode aparecer como 0,0%

_fio = Side(style="thin", color="DCE3E5")
_fio_forte = Side(style="medium", color=MARCA)


def _estilo(tam=10, negrito=False, italico=False, cor=TINTA, fundo=None, alinhar=None, quebra=False,
            fmt=None, borda=None, recuo=0):
    return {
        "font": Font(name=FONTE, size=tam, bold=negrito, italic=italico, color=cor),
        "fill": PatternFill("solid", start_color=fundo, end_color=fundo) if fundo else PatternFill(),
        "alignment": Alignment(horizontal=alinhar, vertical="center", wrap_text=quebra, indent=recuo),
        "border": borda or Border(),
        "number_format": fmt or "General",
    }


E = {
    "titulo": _estilo(17, negrito=True, cor=MARCA),
    "subtitulo": _estilo(10, italico=True, cor=CINZA),
    "secao": _estilo(12, negrito=True, cor=MARCA, borda=Border(bottom=_fio_forte)),
    "secao_fio": _estilo(borda=Border(bottom=_fio_forte)),
    "cab": _estilo(10, negrito=True, cor="FFFFFF", fundo=MARCA, alinhar="left", quebra=True, recuo=1),
    "cab_num": _estilo(10, negrito=True, cor="FFFFFF", fundo=MARCA, alinhar="right", quebra=True),
    "cab_grupo": _estilo(9, negrito=True, cor=MARCA, fundo="E3ECEE", recuo=1),
    "txt": _estilo(borda=Border(bottom=_fio), recuo=1),
    "txt_forte": _estilo(negrito=True, borda=Border(bottom=_fio), recuo=1),
    "txt_suave": _estilo(9, cor=CINZA, borda=Border(bottom=_fio), recuo=1),
    "cod": _estilo(9, cor=CINZA, borda=Border(bottom=_fio), fmt="@", recuo=1),
    "reais": _estilo(fmt=REAIS, borda=Border(bottom=_fio)),
    "reais_forte": _estilo(negrito=True, fmt=REAIS, borda=Border(bottom=_fio)),
    "reais_suave": _estilo(9, cor=CINZA, fmt=REAIS, borda=Border(bottom=_fio)),
    # subir é ruim para quem compra (vermelho), descer é bom (verde)
    "variacao_sobe": _estilo(negrito=True, cor=VERMELHO, fmt=VARIACAO, borda=Border(bottom=_fio)),
    "variacao_desce": _estilo(negrito=True, cor=VERDE, fmt=VARIACAO, borda=Border(bottom=_fio)),
    "variacao": _estilo(negrito=True, fmt=VARIACAO, borda=Border(bottom=_fio)),
    "variacao_sobe_pequena": _estilo(negrito=True, cor=VERMELHO, fmt=VARIACAO_PEQUENA, borda=Border(bottom=_fio)),
    "variacao_desce_pequena": _estilo(negrito=True, cor=VERDE, fmt=VARIACAO_PEQUENA, borda=Border(bottom=_fio)),
    "inteiro": _estilo(fmt=INTEIRO, borda=Border(bottom=_fio)),
    "inteiro_destaque": _estilo(negrito=True, fmt=INTEIRO, fundo="FDF1D8", borda=Border(bottom=_fio)),
    "txt_destaque": _estilo(negrito=True, fundo="FDF1D8", borda=Border(bottom=_fio), recuo=1),
    "alerta": _estilo(negrito=True, cor=VERMELHO, borda=Border(bottom=_fio), recuo=1),
    "subiu": _estilo(cor=VERMELHO, borda=Border(bottom=_fio), recuo=1),
    "desceu": _estilo(cor=VERDE, borda=Border(bottom=_fio), recuo=1),
    "neutro": _estilo(cor=CINZA, borda=Border(bottom=_fio), recuo=1),
    "atencao": _estilo(cor=AMBAR, borda=Border(bottom=_fio), recuo=1),
    # resumo
    "rotulo": _estilo(10, cor=CINZA, borda=Border(bottom=_fio)),
    "numero": _estilo(13, negrito=True, fmt=INTEIRO, alinhar="right", borda=Border(bottom=_fio)),
    "numero_alerta": _estilo(13, negrito=True, cor=VERMELHO, fmt=INTEIRO, alinhar="right", borda=Border(bottom=_fio)),
    "nota": _estilo(10, cor=CINZA, quebra=True, borda=Border(bottom=_fio), recuo=1),
    "termo": _estilo(10, negrito=True, borda=Border(bottom=_fio)),
    "rodape": _estilo(9, italico=True, cor=CLARO),
}


def _reais(centavos):
    return None if centavos is None else centavos / 100


def _br(d):
    return f"{d[8:]}/{d[5:7]}/{d[:4]}" if d else "—"


class Aba:
    """Uma aba escrita de cima para baixo, com margem à esquerda (coluna A)."""

    def __init__(self, wb, nome, larguras, caber_na_largura=True, congelar=None):
        self.ws = wb.create_sheet(nome)
        self.n = 0
        if congelar:
            # no modo write-only, só vale se definido antes da primeira linha
            self.ws.freeze_panes = congelar
        self.ws.sheet_view.showGridLines = False
        self.ws.sheet_view.zoomScale = 100
        for i, w in enumerate([2] + larguras, start=1):
            self.ws.column_dimensions[get_column_letter(i)].width = w
        self.largura = larguras
        ps = self.ws.page_setup
        ps.orientation, ps.paperSize = "landscape", "9"  # A4
        if caber_na_largura:
            self.ws.sheet_properties.pageSetUpPr.fitToPage = True
            ps.fitToWidth, ps.fitToHeight = 1, 0

    def celula(self, valor, estilo, texto=False):
        c = WriteOnlyCell(self.ws, value=valor)
        if texto and isinstance(valor, str):
            c.data_type = "s"  # nunca vira fórmula, mesmo começando com "="
        # aplicar fonte, fundo, borda etc. célula a célula custa caro em quase
        # 1 milhão de células; o estilo é montado uma vez e reaproveitado
        cache = self.ws.parent._estilos_prontos
        if estilo not in cache:
            e = E[estilo]
            c.font, c.fill, c.alignment, c.border, c.number_format = (
                e["font"], e["fill"], e["alignment"], e["border"], e["number_format"])
            cache[estilo] = copy(c._style)
        else:
            c._style = copy(cache[estilo])
        return c

    def linha(self, celulas=(), altura=None):
        self.n += 1
        if altura:
            self.ws.row_dimensions[self.n].height = altura
        self.ws.append([None] + list(celulas))
        return self.n

    def espaco(self, altura=8):
        self.linha(altura=altura)

    def titulo(self, texto, subtitulo):
        self.linha([self.celula(texto, "titulo")], altura=30)
        self.linha([self.celula(subtitulo, "subtitulo")], altura=18)
        self.espaco(10)

    def secao(self, texto, colunas):
        self.espaco(14)
        self.linha([self.celula(texto, "secao")] + [self.celula(None, "secao_fio") for _ in range(colunas - 1)], altura=22)

    def cabecalho(self, nomes, numericas=()):
        self.linha([self.celula(n, "cab_num" if i in numericas else "cab") for i, n in enumerate(nomes)], altura=30)
        return self.n


def _altura(texto, largura):
    """Altura de uma linha com texto quebrado numa coluna de `largura` caracteres."""
    # a largura da coluna conta dígitos "0"; em Arial cabem ~15% mais letras comuns
    return 15 * max(1, math.ceil(len(texto) / (largura * 1.15))) + 5


# ------------------------------------------------------------------ abas

def _resumo(wb, data, anterior, fonte, produtos, comparacao, historico):
    a = Aba(wb, "Resumo", [34, 14, 3, 84])
    a.titulo("Preços máximos de medicamentos (CMED)",
             f"Lista oficial publicada em {_br(data)} · coletada e conferida pelo robô monitor-precos-cmed")

    com_pmc = sum(1 for p in produtos if p["pmc"].get("19") is not None)
    a.secao("Esta lista", 4)
    for rotulo, n, nota in [
        ("Apresentações", len(produtos), "cada combinação de remédio, dose e embalagem com preço registrado"),
        ("Com PMC", com_pmc, "podem ser vendidas em farmácia — o PMC é o preço máximo"),
        ("Sem PMC", len(produtos) - com_pmc, "uso restrito a hospitais: não têm preço para farmácia"),
        ("Isentas de ICMS", sum(1 for p in produtos if p["icms0"]), "onde a isenção vale, o teto é o PMC 0%"),
    ]:
        a.linha([a.celula(rotulo, "rotulo"), a.celula(n, "numero"), a.celula(None, "rotulo"), a.celula(nota, "nota")], altura=22)

    if comparacao:
        m = comparacao["mudancas"]
        fora = [x for x in m if x.get("fora_do_normal")]
        subiu = sum(1 for x in m if x["tipo"] == "subiu")
        desceu = sum(1 for x in m if x["tipo"] == "desceu")
        a.secao(f"O que mudou desde {_br(anterior)}", 4)
        for rotulo, n, nota, estilo in [
            ("Preços que mudaram", len(m), f"{subiu} subiram e {desceu} desceram (PMC a 19%) — aba \"Mudanças de preço\"", "numero"),
            ("Fora do normal", len(fora), "multiplicaram ou dividiram por 3 ou mais: confira na fonte antes de usar", "numero_alerta" if fora else "numero"),
            ("Entraram na lista", len(comparacao["novos"]), "aba \"Novos na lista\"", "numero"),
            ("Saíram da lista", len(comparacao["sairam"]), "aba \"Saíram da lista\"", "numero"),
        ]:
            a.linha([a.celula(rotulo, "rotulo"), a.celula(n, estilo), a.celula(None, "rotulo"), a.celula(nota, "nota")], altura=22)

        if fora:
            a.secao("Fora do normal: confira antes de usar", 4)
            for x in fora:
                p = x["produto"]
                texto = (f"{p['apresentacao']} · {p['laboratorio']} · de R$ {_moeda(x['antes'])} "
                         f"para R$ {_moeda(x['depois'])}")
                a.linha([a.celula(p["produto"], "txt_forte", texto=True), a.celula(x.get("variacao"), _cor_variacao(x.get("variacao"))),
                         a.celula(None, "rotulo"), a.celula(texto, "nota", texto=True)], altura=_altura(texto, 84))

    a.secao("Abas desta planilha", 4)
    abas = [
        ("Mudanças de preço", "o que subiu e desceu desde a lista anterior, com os casos fora do normal primeiro"),
        ("Novos na lista / Saíram da lista", "apresentações que entraram e que deixaram a lista"),
        ("Mês a mês", "quantos preços mudaram em cada publicação — o reajuste anual de abril em destaque"),
        ("Consulta PMC", "todas as apresentações com o PMC do Pará (19%) e o PMC 0%; use o filtro do cabeçalho para buscar"),
        ("PMC por alíquota", "o PMC em todas as alíquotas de ICMS, para quem vende em outro estado ou em Área de Livre Comércio"),
    ] if comparacao else [
        ("Consulta PMC", "todas as apresentações com o PMC do Pará (19%) e o PMC 0%"),
        ("PMC por alíquota", "o PMC em todas as alíquotas de ICMS"),
    ]
    for nome, desc in abas:
        a.linha([a.celula(nome, "termo"), a.celula(None, "rotulo"), a.celula(None, "rotulo"), a.celula(desc, "nota")],
                altura=_altura(desc, 84))

    a.secao("Como ler", 4)
    for termo, texto in [
        ("PMC", "Preço máximo ao consumidor: o maior preço que farmácia e drogaria podem cobrar. Vender acima é infração."),
        ("Alíquota de ICMS", "O PMC muda conforme o ICMS do estado. No Pará é 19%. Confira a do seu estado: a própria CMED diz que isso é responsabilidade de quem vende."),
        ("Área de Livre Comércio", "Colunas \"ALC\": valem para cidades de Área de Livre Comércio (no AP, AM, RO, RR e AC)."),
        ("Isento de ICMS", "Onde a isenção vale, o teto é o PMC 0%. Na lista oficial, só parte desses produtos vem marcada com asterisco; aqui a marcação vem da coluna \"ICMS 0%\"."),
        ("Fora do normal", "O robô não decide se é correção de um erro anterior ou erro novo da lista oficial — só avisa para conferir."),
    ]:
        a.linha([a.celula(termo, "termo"), a.celula(None, "rotulo"), a.celula(None, "rotulo"), a.celula(texto, "nota")],
                altura=_altura(texto, 84))

    a.espaco(16)
    a.linha([a.celula(f"Fonte: {fonte}", "rodape", texto=True)])
    a.linha([a.celula("Gerada automaticamente a partir do arquivo oficial da Anvisa. Ferramenta de apoio: não substitui a consulta à lista oficial.", "rodape")])


def _cor_variacao(v):
    if not v:
        return "variacao"
    return ("variacao_sobe" if v > 0 else "variacao_desce") + ("_pequena" if abs(v) < 0.001 else "")


def _moeda(centavos):
    if centavos is None:
        return "—"
    inteiro, frac = divmod(centavos, 100)
    return f"{inteiro:,}".replace(",", ".") + f",{frac:02d}"


SITUACAO = {"subiu": ("▲ Subiu", "subiu"), "desceu": ("▼ Desceu", "desceu"),
            "passou a ter PMC": ("Passou a ter PMC", "neutro"), "deixou de ter PMC": ("Deixou de ter PMC", "neutro")}


def _mudancas(wb, data, anterior, comparacao):
    a = Aba(wb, "Mudanças de preço", [27, 30, 46, 32, 13, 13, 12, 16], congelar=_congelar())
    a.titulo("Mudanças de preço", f"PMC a 19% (Pará) na lista de {_br(anterior)} e na de {_br(data)}. "
             "Os casos fora do normal vêm primeiro.")
    topo = a.cabecalho(["Situação", "Produto", "Apresentação", "Laboratório", "PMC antes", "PMC agora", "Variação",
                        "Código de barras"], numericas=(4, 5, 6))
    for m in comparacao["mudancas"]:
        p = m["produto"]
        rotulo, estilo = ("⚠ Fora do normal — confira", "alerta") if m.get("fora_do_normal") else SITUACAO[m["tipo"]]
        a.linha([a.celula(rotulo, estilo), a.celula(p["produto"], "txt_forte", True), a.celula(p["apresentacao"], "txt", True),
                 a.celula(p["laboratorio"], "txt_suave", True), a.celula(_reais(m["antes"]), "reais_suave"),
                 a.celula(_reais(m["depois"]), "reais_forte"), a.celula(m.get("variacao"), _cor_variacao(m.get("variacao"))),
                 a.celula(p["ean1"], "cod")], altura=20)
    _fechar_tabela(a, topo, 8)


def _lista_simples(wb, nome, subtitulo, grupo):
    a = Aba(wb, nome, [30, 50, 34, 15, 24, 16], congelar=_congelar())
    a.titulo(nome, subtitulo)
    topo = a.cabecalho(["Produto", "Apresentação", "Laboratório", "PMC 19% (Pará)", "Observação", "Código de barras"], numericas=(3,))
    for p in grupo:
        obs = "Uso hospitalar (sem PMC)" if p["hospitalar"] else ("Isento de ICMS" if p["icms0"] else "")
        a.linha([a.celula(p["produto"], "txt_forte", True), a.celula(p["apresentacao"], "txt", True),
                 a.celula(p["laboratorio"], "txt_suave", True), a.celula(_reais(p["pmc"].get("19")), "reais"),
                 a.celula(obs, "neutro"), a.celula(p["ean1"], "cod")], altura=18)
    _fechar_tabela(a, topo, 6)


def _historico(wb, historico):
    a = Aba(wb, "Mês a mês", [26, 12, 12, 16, 12, 12, 16], congelar=_congelar())
    a.titulo("Mês a mês", "Quantos preços mudaram em cada publicação da lista, comparada com a anterior. "
             "Em abril sai o reajuste anual; nos outros meses, poucas dezenas mudam.")
    topo = a.cabecalho(["Publicada em", "Subiram", "Desceram", "Fora do normal", "Entraram", "Saíram", "Apresentações"],
                       numericas=(1, 2, 3, 4, 5, 6))
    for h in reversed(historico):
        abril = h["data"][5:7] == "04" and "subiu" in h
        num = "inteiro_destaque" if abril else "inteiro"
        rotulo = _br(h["data"]) + ("  · reajuste anual" if abril else "")
        a.linha([a.celula(rotulo, "txt_destaque" if abril else "txt")]
                + [a.celula(h.get(k), num) for k in ("subiu", "desceu", "fora_do_normal", "novos", "sairam", "produtos")],
                altura=20)
    _fechar_tabela(a, topo, 7)


def _consulta(wb, data, produtos):
    a = Aba(wb, "Consulta PMC", [30, 48, 30, 30, 12, 16, 15, 13, 27], congelar=_congelar(1))
    a.titulo("Consulta PMC", f"Todas as {len(produtos):,} apresentações da lista de {_br(data)}. ".replace(",", ".")
             + "Para buscar, use a setinha de filtro no cabeçalho (ex.: Produto → contém \"losartana\").")
    topo = a.cabecalho(["Produto", "Apresentação", "Substância", "Laboratório", "Tipo", "Código de barras",
                        "PMC 19% (Pará)", "PMC 0%", "Observação"], numericas=(6, 7))
    for p in produtos:
        obs = "Uso hospitalar (sem PMC)" if p["hospitalar"] and p["pmc"].get("19") is None else (
            "Isento de ICMS: teto PMC 0%" if p["icms0"] else "")
        a.linha([a.celula(p["produto"], "txt_forte", True), a.celula(p["apresentacao"], "txt", True),
                 a.celula(p["substancia"], "txt_suave", True), a.celula(p["laboratorio"], "txt_suave", True),
                 a.celula(p.get("tipo", ""), "txt_suave", True), a.celula(p["ean1"], "cod"),
                 a.celula(_reais(p["pmc"].get("19")), "reais_forte"), a.celula(_reais(p["pmc"].get("0")), "reais_suave"),
                 a.celula(obs, "atencao" if obs.startswith("Isento") else "neutro")], altura=18)
    _fechar_tabela(a, topo, 9)


def _aliquotas(wb, data, aliquotas, produtos):
    normais = [x for x in aliquotas if x not in ("sem",) and not x.endswith("_alc")]
    alc = [x for x in aliquotas if x.endswith("_alc")]
    fmt = lambda x: x.replace("_alc", "").replace(".", ",") + "%"  # noqa: E731
    a = Aba(wb, "PMC por alíquota", [30, 40, 16] + [10] * (len(normais) + len(alc)), caber_na_largura=False,
            congelar=_congelar(2, CAB + 1))  # uma linha a mais: a faixa que agrupa as alíquotas
    a.titulo("PMC por alíquota de ICMS", f"Lista de {_br(data)}. Escolha a coluna da alíquota do seu estado "
             "(Pará: 19%). As colunas ALC valem para Área de Livre Comércio.")
    grupo = [a.celula(None, "cab_grupo") for _ in range(3)]
    grupo += [a.celula("Alíquota de ICMS" if i == 0 else None, "cab_grupo") for i in range(len(normais))]
    grupo += [a.celula("Área de Livre Comércio (ALC)" if i == 0 else None, "cab_grupo") for i in range(len(alc))]
    a.linha(grupo, altura=18)
    topo = a.cabecalho(["Produto", "Apresentação", "Código de barras"] + [fmt(x) for x in normais] + [fmt(x) + " ALC" for x in alc],
                       numericas=range(3, 3 + len(normais) + len(alc)))
    for p in produtos:
        a.linha([a.celula(p["produto"], "txt_forte", True), a.celula(p["apresentacao"], "txt", True), a.celula(p["ean1"], "cod")]
                + [a.celula(_reais(p["pmc"].get(x)), "reais_forte" if x == "19" else "reais") for x in normais + alc], altura=18)
    _fechar_tabela(a, topo, 3 + len(normais) + len(alc))


# título, subtítulo e espaço ocupam as linhas 1 a 3: o cabeçalho da tabela é a 4
CAB = 4


def _congelar(colunas=0, cab=CAB):
    """Célula de congelamento: título e cabeçalho fixos (e as `colunas` primeiras, depois da margem)."""
    return f"{get_column_letter(2 + colunas)}{cab + 1}"


def _fechar_tabela(a, topo, colunas):
    """Liga o filtro no cabeçalho e repete o cabeçalho em cada folha impressa."""
    ultima = get_column_letter(1 + colunas)
    assert a.ws.freeze_panes in (None, _congelar(0, topo), _congelar(1, topo), _congelar(2, topo)), "cabeçalho fora do lugar"
    a.ws.auto_filter.ref = f"B{topo}:{ultima}{max(topo, a.n)}"
    a.ws.print_title_rows = f"{topo}:{topo}"


# ------------------------------------------------------------------ arquivo

def gerar_planilha(caminho, data, anterior, fonte, aliquotas, produtos, comparacao, historico=()):
    # write-only: freeze_panes, filtro e larguras precisam ser definidos antes de
    # a aba ser gravada, o que acontece no save — por isso tudo é montado aqui
    wb = Workbook(write_only=True)
    wb._estilos_prontos = {}
    _resumo(wb, data, anterior, fonte, produtos, comparacao, historico)
    if comparacao:
        _mudancas(wb, data, anterior, comparacao)
        _lista_simples(wb, "Novos na lista", f"Apresentações que entraram na lista de {_br(data)}.", comparacao["novos"])
        _lista_simples(wb, "Saíram da lista", f"Apresentações da lista de {_br(anterior)} que não estão na de {_br(data)}.",
                       comparacao["sairam"])
    if historico:
        _historico(wb, historico)
    _consulta(wb, data, produtos)
    _aliquotas(wb, data, aliquotas, produtos)

    wb.properties.creator = "monitor-precos-cmed"
    wb.properties.title = f"Lista CMED {_br(data)}"
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
