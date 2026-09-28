"""
Coleta na página da Anvisa: acha o link da lista de preços mais recente e
baixa o arquivo — de forma educada (identifica o robô, uma requisição por
vez, pausa entre elas, prazo máximo) e desconfiada (confere se o que veio é
mesmo uma planilha, tenta de novo se a conexão cair).

O robots.txt do gov.br permite robôs nessas páginas (conferido em 09/2026).
"""
import hashlib
import re
import time
import urllib.parse
import urllib.request
from html.parser import HTMLParser

PAGINA = "https://www.gov.br/anvisa/pt-br/assuntos/medicamentos/cmed/precos"
ANTERIORES = PAGINA + "/anos-anteriores/anos-anteriores"
USER_AGENT = "monitor-precos-cmed/1.0 (+https://github.com/jvtdemiranda/monitor-precos-cmed)"
# "xls_conformidade_site_20260909_222937320.xlsx": "site" é a lista de PMC
# (farmácias); "gov" é a de compras públicas, que não interessa aqui.
RE_ARQUIVO = re.compile(r"/xls_conformidade_site_(\d{8})_\d+\.xlsx", re.I)
PAUSA = 2.0
HOSTS_PERMITIDOS = {"www.gov.br"}
# o arquivo real tem ~12 MB; bem acima disso é engano (ou abuso) e não vale encher o disco
TAMANHO_MAXIMO = 100 * 1024 * 1024


class LinkNaoEncontrado(Exception):
    """A página mudou e o link da lista não foi achado — o robô para e avisa."""


class _Links(HTMLParser):
    def __init__(self):
        super().__init__()
        self.links, self._href, self._txt = [], None, []

    def handle_starttag(self, tag, attrs):
        if tag == "a":
            self._href, self._txt = dict(attrs).get("href") or "", []

    def handle_data(self, data):
        if self._href is not None:
            self._txt.append(data)

    def handle_endtag(self, tag):
        if tag == "a" and self._href is not None:
            self.links.append((" ".join("".join(self._txt).split()), self._href))
            self._href = None


def listas_na_pagina(html, base=PAGINA):
    """
    Todos os arquivos de lista de PMC citados na página, do mais novo pro
    mais antigo: [(data "AAAA-MM-DD", url do download), ...].
    """
    p = _Links()
    p.feed(html)
    achados = {}
    for _, href in p.links:
        url = urllib.parse.urljoin(base, href)
        m = RE_ARQUIVO.search(url)
        if not m:
            continue
        if urllib.parse.urlparse(url).hostname != "www.gov.br":
            continue  # só baixa do próprio site da Anvisa
        d = m.group(1)
        data = f"{d[:4]}-{d[4:6]}-{d[6:]}"
        # o link aparece como ".../arquivo.xlsx/@@download/file" ou ".../view"
        raiz = url[: m.end()]
        achados[raiz] = data
    return sorted(((data, raiz + "/@@download/file") for raiz, data in achados.items()), reverse=True)


def lista_mais_recente(html):
    listas = listas_na_pagina(html)
    if not listas:
        raise LinkNaoEncontrado(
            "nenhum link 'xls_conformidade_site_AAAAMMDD_*.xlsx' na página de preços — "
            "o layout do site da Anvisa pode ter mudado")
    return listas[0]


class OrigemRecusada(Exception):
    """O endereço (ou o redirecionamento) saiu do site da Anvisa."""


def _conferir_host(url):
    host = urllib.parse.urlparse(url).hostname
    if urllib.parse.urlparse(url).scheme != "https" or host not in HOSTS_PERMITIDOS:
        raise OrigemRecusada(f"endereço fora do site da Anvisa: {url}")


def baixar(url, prazo=300, tentativas=3, espera=(5, 20, 60)):
    """Baixa com prazo total, tentando de novo em falha de rede. Devolve bytes."""
    _conferir_host(url)
    erro = None
    for n in range(tentativas):
        time.sleep(PAUSA if n == 0 else espera[min(n - 1, len(espera) - 1)])
        inicio = time.time()
        try:
            req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
            with urllib.request.urlopen(req, timeout=60) as r:
                _conferir_host(r.geturl())  # urllib segue redirecionamento sozinho
                partes, total = [], 0
                while bloco := r.read(1 << 16):
                    partes.append(bloco)
                    total += len(bloco)
                    if total > TAMANHO_MAXIMO:
                        raise OrigemRecusada(f"arquivo passou de {TAMANHO_MAXIMO // 2**20} MB: {url}")
                    if time.time() - inicio > prazo:
                        raise TimeoutError(f"download passou de {prazo}s")
                return b"".join(partes)
        except OrigemRecusada:
            raise  # não adianta tentar de novo
        except Exception as e:  # noqa: BLE001 — qualquer falha de rede: tenta de novo
            erro = e
            print(f"  tentativa {n + 1}/{tentativas} falhou: {e!r}", flush=True)
    raise ConnectionError(f"não consegui baixar {url}: {erro!r}")


def conferir_planilha(conteudo):
    """Um .xlsx é um zip; se veio uma página de erro em HTML, isso pega."""
    if conteudo[:4] != b"PK\x03\x04":
        inicio = conteudo[:80].decode("utf-8", "replace")
        raise ValueError(f"o arquivo baixado não é uma planilha xlsx (começa com {inicio!r})")
    return hashlib.sha256(conteudo).hexdigest()
