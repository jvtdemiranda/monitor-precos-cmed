/* Página do monitor: lê os dados publicados pelo robô e liga o conferidor e a busca.
 * Todo texto que vem de fora (lista oficial ou tabela da farmácia) entra na página
 * como texto (textContent), nunca como HTML. */
(function () {
  "use strict";
  var C = window.Conferidor;
  var estado = { resumo: null, catalogo: null, pmc: {}, aliquota: "19", indice: null, busca: null, resultado: null, filtro: null };

  // ---------------------------------------------------------------- utilidades

  function el(tag, attrs) {
    var e = document.createElement(tag);
    if (attrs) Object.keys(attrs).forEach(function (k) {
      if (k === "texto") e.textContent = attrs[k];
      else if (k === "classe") e.className = attrs[k];
      else e.setAttribute(k, attrs[k]);
    });
    for (var i = 2; i < arguments.length; i++) {
      var f = arguments[i];
      if (f === null || f === undefined || f === false) continue;
      e.appendChild(typeof f === "string" || typeof f === "number" ? document.createTextNode(String(f)) : f);
    }
    return e;
  }
  function $(id) { return document.getElementById(id); }
  function limpar(no) { while (no.firstChild) no.removeChild(no.firstChild); return no; }
  function dataBR(d) { return d ? d.slice(8) + "/" + d.slice(5, 7) + "/" + d.slice(0, 4) : "—"; }
  function inteiro(n) { return Number(n || 0).toLocaleString("pt-BR"); }
  function pct(v) {
    if (v === null || v === undefined) return "—";
    var casas = Math.abs(v) < 0.01 ? 2 : 1; // +0,04% não pode aparecer como +0,0%
    var t = (Math.abs(v) * 100).toLocaleString("pt-BR", { maximumFractionDigits: casas, minimumFractionDigits: casas }) + "%";
    return (v >= 0 ? "+" : "−") + t;
  }
  function baixarJSON(url) {
    return fetch(url, { cache: "no-cache" }).then(function (r) {
      if (!r.ok) throw new Error(url + ": HTTP " + r.status);
      return r.json();
    });
  }
  var ICONE = {
    alerta: "M12 8v5M12 16.5v.5M10.3 3.9L2.6 17.5A2 2 0 0 0 4.3 20.5h15.4a2 2 0 0 0 1.7-3L13.7 3.9a2 2 0 0 0-3.4 0z",
    ok: "M5 12.5l4.5 4.5L19 7.5",
    info: "M12 11v6M12 7.5v.5M12 21a9 9 0 1 0 0-18 9 9 0 0 0 0 18z"
  };
  function selo(classe, icone, texto) {
    var s = el("span", { classe: "selo " + classe });
    var svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    svg.setAttribute("viewBox", "0 0 24 24"); svg.setAttribute("fill", "none"); svg.setAttribute("stroke", "currentColor");
    svg.setAttribute("stroke-width", "2.4"); svg.setAttribute("stroke-linecap", "round"); svg.setAttribute("stroke-linejoin", "round");
    svg.setAttribute("aria-hidden", "true");
    var p = document.createElementNS("http://www.w3.org/2000/svg", "path"); p.setAttribute("d", ICONE[icone]);
    svg.appendChild(p); s.appendChild(svg); s.appendChild(document.createTextNode(texto));
    return s;
  }
  function td(rotulo, conteudo, classe) {
    var c = el("td", classe ? { classe: classe } : null);
    if (rotulo) c.setAttribute("data-rotulo", rotulo);
    if (conteudo !== null && conteudo !== undefined) c.appendChild(typeof conteudo === "string" ? document.createTextNode(conteudo) : conteudo);
    return c;
  }
  function nomeProduto(produto, apresentacao) {
    return el("div", null, el("div", { classe: "nome", texto: produto }), el("div", { classe: "apres", texto: apresentacao }));
  }

  // ---------------------------------------------------------------- resumo

  function mostrarResumo(r) {
    estado.resumo = r;
    $("subtitulo").textContent = "Lista oficial da CMED/Anvisa publicada em " + dataBR(r.publicada_em) + ", coletada e conferida por um robô.";
    var fonte = $("fonte");
    fonte.appendChild(document.createTextNode("Fonte: "));
    fonte.appendChild(el("a", { href: "https://www.gov.br/anvisa/pt-br/assuntos/medicamentos/cmed/precos", rel: "noopener" }, "Listas de preços de medicamentos — CMED/Anvisa"));
    fonte.appendChild(document.createTextNode(" (lista de " + dataBR(r.publicada_em) + ")."));

    var mudaram = r.mudancas.length, fora = r.mudancas.filter(function (m) { return m.fora; }).length;
    var kpis = [
      ["Apresentações", inteiro(r.produtos), inteiro(r.com_pmc) + " com PMC"],
      ["Preços que mudaram", inteiro(mudaram), r.anterior ? (fora ? fora + " fora do normal" : "desde " + dataBR(r.anterior)) : "primeira lista"],
      ["Entraram na lista", inteiro(r.novos.length), r.anterior ? "desde " + dataBR(r.anterior) : "—"],
      ["Sem PMC", inteiro(r.produtos - r.com_pmc), "uso restrito a hospitais"]
    ];
    var box = limpar($("kpis"));
    kpis.forEach(function (k) {
      box.appendChild(el("div", { classe: "kpi" }, el("div", { classe: "lbl", texto: k[0] }), el("div", { classe: "val", texto: k[1] }), el("div", { classe: "sub", texto: k[2] })));
    });

    var sel = limpar($("aliquota"));
    r.aliquotas.forEach(function (a) {
      var o = el("option", { value: a.chave, texto: a.rotulo + (a.chave === "19" ? " (Pará)" : "") });
      if (a.chave === "19") o.selected = true;
      sel.appendChild(o);
    });

    mostrarMudancas(r);
    mostrarHistorico(r.historico);
    mostrarQualidade(r);
  }

  function mostrarMudancas(r) {
    if (!r.anterior) {
      $("titulo-mudancas").textContent = "O que mudou";
      $("explica-mudancas").textContent = "Esta é a primeira lista coletada; a comparação aparece a partir da próxima.";
      return;
    }
    $("titulo-mudancas").textContent = "O que mudou desde " + dataBR(r.anterior);
    $("explica-mudancas").textContent = "PMC na alíquota de 19% (Pará) na lista anterior e na atual. Nas outras alíquotas, o preço muda na mesma proporção.";
    var fora = r.mudancas.filter(function (m) { return m.fora; });
    var caixa = limpar($("fora-do-normal"));
    if (fora.length) {
      caixa.appendChild(el("div", { classe: "alerta-caixa" },
        el("h3", null, selo("fora", "alerta", "Fora do normal"), " " + fora.length + (fora.length === 1 ? " preço multiplicou ou dividiu" : " preços multiplicaram ou dividiram") + " por 3 ou mais"),
        el("p", { texto: "Pode ser correção de um erro da lista anterior ou um erro novo na lista oficial. O robô não decide qual é o certo — confira na fonte antes de usar esses valores." })));
    }
    var tb = limpar($("tab-mudancas").querySelector("tbody"));
    r.mudancas.forEach(function (m) {
      var variacao = m.variacao === null || m.variacao === undefined ? m.tipo : pct(m.variacao);
      var cel = el("span", { classe: "num", texto: variacao });
      var tr = el("tr", m.fora ? { classe: "destaque" } : null,
        td(null, nomeProduto(m.produto, m.apresentacao)),
        td("Laboratório", m.laboratorio),
        td("Antes", el("span", { classe: "num", texto: C.reais(m.antes) }), "d"),
        td("Agora", el("span", { classe: "num", texto: C.reais(m.depois) }), "d"),
        td("Variação", m.fora ? el("span", null, cel, " ", selo("fora", "alerta", "confira")) : cel, "d"));
      tb.appendChild(tr);
    });
    if (!r.mudancas.length) tb.appendChild(el("tr", null, el("td", { colspan: "5", texto: "Nenhum preço mudou." })));

    [["det-novos", r.novos, "entraram na lista"], ["det-sairam", r.sairam, "saíram da lista"]].forEach(function (x) {
      var det = $(x[0]), lista = x[1], LIMITE = 60;
      det.querySelector("summary").textContent = inteiro(lista.length) + " apresentações " + x[2];
      var ul = limpar(det.querySelector("ul"));
      lista.slice(0, LIMITE).forEach(function (p) {
        ul.appendChild(el("li", null, el("b", { texto: p.produto }), " " + p.apresentacao + " — " + p.laboratorio +
          (p.hospitalar ? " (uso hospitalar)" : " · PMC 19%: " + C.reais(p.pmc))));
      });
      det.querySelector(".mais").textContent = lista.length > LIMITE ? "Mostrando " + LIMITE + " de " + inteiro(lista.length) + ". A lista completa está na planilha." : "";
      if (!lista.length) det.hidden = true;
    });
  }

  var MESES = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"];
  function mostrarHistorico(h) {
    var tb = limpar($("tab-historico").querySelector("tbody"));
    h.slice().reverse().forEach(function (x) {
      var abril = x.data.slice(5, 7) === "04";
      var primeiro = x.subiu === undefined;
      var n = function (v) { return td(null, el("span", { classe: "num", texto: primeiro ? "—" : inteiro(v) }), "d"); };
      tb.appendChild(el("tr", abril && !primeiro ? { classe: "abril" } : null,
        td(null, el("span", null, dataBR(x.data), abril && !primeiro ? el("span", { classe: "apres", texto: " · reajuste anual" }) : null)),
        n(x.subiu), n(x.desceu), n(x.fora_do_normal), n(x.novos), n(x.sairam)));
    });
  }

  function mostrarQualidade(r) {
    var q = r.qualidade || {};
    var itens = [
      [q.linhas_notas, "linhas de notas antes da tabela, puladas (o cabeçalho é procurado, não assumido)"],
      [q.campos_traco, "campos vazios que vêm preenchidos com um traço, tratados como vazios"],
      [q.espacos_sobrando, "textos com espaços sobrando, limpos"],
      [r.isentos, "apresentações isentas de ICMS — só " + inteiro(q.precos_asterisco) + " vêm com o asterisco que a própria lista promete; o robô usa a coluna \"ICMS 0%\""],
      [q.sem_pmc_hospitalar, "apresentações de uso hospitalar, sem PMC em nenhuma alíquota"],
      [q.ean_repetido, "códigos de barras repetidos em mais de uma apresentação (o conferidor avisa e usa o menor PMC)"],
      [q.ean_invalido, "códigos de barras com dígito verificador inválido (mantidos: é o que está na caixa)"],
      [q.ean_zerado, "códigos de barras \"0000000000000\", descartados"]
    ];
    var box = limpar($("qualidade"));
    itens.forEach(function (i) {
      if (i[0] === undefined || i[0] === null) return;
      box.appendChild(el("div", { classe: "q" }, el("b", { texto: inteiro(i[0]) }), el("span", { texto: i[1] })));
    });
  }

  // ---------------------------------------------------------------- lista e PMC

  function carregarPMC(chave) {
    if (estado.pmc[chave]) return Promise.resolve(estado.pmc[chave]);
    return baixarJSON("dados/pmc/" + encodeURIComponent(chave) + ".json").then(function (v) { estado.pmc[chave] = v; return v; });
  }

  function listaPronta(cat) {
    estado.catalogo = cat;
    estado.indice = C.indiceDeCodigos(cat);
    estado.busca = cat.itens.map(function (it) {
      return C.semAcento([it[2], it[3], cat.substancias[it[4]], cat.laboratorios[it[5]]].join(" "));
    });
    $("carregando-lista").hidden = true;
    $("btn-conferir").disabled = false;
    $("btn-exemplo").disabled = false;
    $("busca").disabled = false;
  }

  // ---------------------------------------------------------------- conferidor

  var ORDEM = ["acima", "atencao", "codigo_estragado", "sem_preco", "sem_codigo", "hospitalar", "nao_encontrado", "ok"];
  var ROTULO = {
    acima: ["acima", "alerta", "Acima do PMC"], atencao: ["atencao", "alerta", "Atenção"], ok: ["ok", "ok", "Dentro do PMC"],
    hospitalar: ["neutro", "info", "Uso hospitalar"], nao_encontrado: ["neutro", "info", "Não está na lista"],
    codigo_estragado: ["atencao", "alerta", "Código estragado"], sem_codigo: ["neutro", "info", "Sem código"], sem_preco: ["neutro", "info", "Sem preço"]
  };
  var OBS = {
    hospitalar: "Produto de uso restrito a hospitais: não tem PMC e não pode ser vendido em farmácia pelo PMC.",
    nao_encontrado: "Esse código de barras não está na lista da CMED (pode não ser medicamento, ou o código está errado)."
  };

  function conferirTexto(texto) {
    var avisos = limpar($("avisos-conferencia"));
    var t = C.lerTabela(texto);
    t.avisos.forEach(function (a) { avisos.appendChild(el("p", { classe: "aviso", texto: a })); });
    if (!t.linhas.length) { $("resultado").hidden = true; return; }
    var chave = $("aliquota").value;
    Promise.all([carregarPMC(chave), carregarPMC("0")]).then(function (v) {
      estado.resultado = C.conferir(t.linhas, estado.catalogo, v[0], v[1], estado.indice);
      estado.filtro = null;
      mostrarResultado();
    }).catch(falhou);
  }

  function mostrarResultado() {
    var res = estado.resultado, cat = estado.catalogo;
    $("resultado").hidden = false;
    var cont = {};
    res.forEach(function (r) { cont[r.status] = (cont[r.status] || 0) + 1; });
    var chips = limpar($("contagens"));
    var todos = el("button", { type: "button", classe: "chip", "aria-pressed": String(!estado.filtro) }, "Todos ", el("span", { classe: "n", texto: String(res.length) }));
    todos.addEventListener("click", function () { estado.filtro = null; mostrarResultado(); });
    chips.appendChild(todos);
    ORDEM.forEach(function (s) {
      if (!cont[s]) return;
      var b = el("button", { type: "button", classe: "chip", "aria-pressed": String(estado.filtro === s) }, ROTULO[s][2] + " ", el("span", { classe: "n", texto: String(cont[s]) }));
      b.addEventListener("click", function () { estado.filtro = estado.filtro === s ? null : s; mostrarResultado(); });
      chips.appendChild(b);
    });

    var tb = limpar($("tab-resultado").querySelector("tbody"));
    res.slice().sort(function (a, b) { return ORDEM.indexOf(a.status) - ORDEM.indexOf(b.status) || a.numero - b.numero; })
      .filter(function (r) { return !estado.filtro || r.status === estado.filtro; })
      .forEach(function (r) {
        var nomeLista = r.itens.length ? cat.itens[r.itens[0]] : null;
        var naLista = nomeLista ? nomeLista[2] + " " + nomeLista[3] : "";
        var repete = naLista && C.semAcento(naLista).indexOf(C.semAcento(r.descricao)) === 0;
        var produto = el("div", null,
          el("div", { classe: "nome", texto: r.descricao || naLista || "(sem descrição)" }),
          el("div", { classe: "apres", texto: (r.codigo || r.codigoBruto || "sem código") + (naLista && r.descricao && !repete ? " · na lista: " + naLista : "") }));
        var obs = (OBS[r.status] ? [OBS[r.status]] : []).concat(r.avisos);
        if (r.status === "acima") obs.unshift("Passa " + C.reais(r.excesso) + " (" + pct(r.excesso / r.pmc).replace("+", "") + ") do teto.");
        if (obs.length) produto.appendChild(el("div", { classe: "obs", texto: obs.join(" ") }));
        var rot = ROTULO[r.status];
        tb.appendChild(el("tr", r.status === "acima" ? { classe: "destaque" } : null,
          td("Linha", el("span", { classe: "num", texto: String(r.numero) })),
          td(null, produto),
          td("Seu preço", el("span", { classe: "num", texto: r.preco === null ? (r.precoBruto || "—") : C.reais(r.preco) }), "d"),
          td("PMC", el("span", { classe: "num", texto: C.reais(r.pmc) }), "d"),
          td("Diferença", el("span", { classe: "num", texto: r.excesso === null ? "—" : (r.excesso > 0 ? "+" : r.excesso < 0 ? "−" : "") + C.reais(Math.abs(r.excesso)) }), "d"),
          td("Situação", selo(rot[0], rot[1], rot[2]))));
      });
  }

  function baixarResultado() {
    var csv = C.paraCSV(estado.resultado, estado.catalogo);
    var a = el("a", { href: URL.createObjectURL(new Blob([csv], { type: "text/csv;charset=utf-8" })), download: "conferencia-pmc.csv" });
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(function () { URL.revokeObjectURL(a.href); }, 1000);
  }

  // ---------------------------------------------------------------- busca

  var esperaBusca;
  function buscar() {
    var q = C.semAcento($("busca").value), box = limpar($("resultados-busca"));
    if (q.length < 3) return;
    var cat = estado.catalogo, achados = [];
    if (/^\d{8,14}$/.test(q)) {
      achados = estado.indice.get(q) || [];
      if (!achados.length) achados = estado.indice.get(C.lerCodigo(q).codigo) || [];
    } else {
      var termos = q.split(/\s+/).filter(Boolean);
      for (var i = 0; i < estado.busca.length && achados.length < 200; i++) {
        var s = estado.busca[i];
        if (termos.every(function (t) { return s.indexOf(t) >= 0; })) achados.push(i);
      }
      achados.sort(function (a, b) {
        return (C.semAcento(cat.itens[b][2]).indexOf(termos[0]) === 0) - (C.semAcento(cat.itens[a][2]).indexOf(termos[0]) === 0) || a - b;
      });
    }
    var chave = $("aliquota").value;
    carregarPMC(chave).then(function (pmc) {
      limpar(box);
      if (!achados.length) { box.appendChild(el("p", { classe: "carregando", texto: "Nada encontrado." })); return; }
      achados.slice(0, 30).forEach(function (i) {
        var it = cat.itens[i], hosp = it[7] & 1;
        var info = [cat.laboratorios[it[5]], cat.tipos[it[6]], it[1].length ? "EAN " + it[1].join(", ") : null].filter(Boolean).join(" · ");
        var direita = el("div", { classe: "pmc" },
          hosp ? selo("neutro", "info", "Uso hospitalar") : el("b", { texto: C.reais(pmc[i]) }),
          el("span", { classe: "apres", texto: hosp ? "sem PMC" : "PMC " + $("aliquota").selectedOptions[0].textContent }));
        var esquerda = nomeProduto(it[2], it[3]);
        esquerda.appendChild(el("div", { classe: "obs", texto: info }));
        if (it[7] & 2) esquerda.appendChild(el("div", { classe: "obs", texto: "Isento de ICMS: onde a isenção vale, o teto é o PMC 0%." }));
        box.appendChild(el("div", { classe: "item" }, esquerda, direita));
      });
      if (achados.length > 30) box.appendChild(el("p", { classe: "mais", texto: "Mostrando 30 resultados. Refine a busca." }));
    }).catch(falhou);
  }

  // ---------------------------------------------------------------- ligações

  function falhou(e) {
    var p = $("erro-carga");
    p.hidden = false;
    p.textContent = "Não consegui carregar os dados da lista (" + e.message + "). Tente recarregar a página.";
  }

  $("btn-conferir").addEventListener("click", function () {
    var arq = $("arquivo").files[0], colado = $("colado").value;
    if (colado.trim()) { conferirTexto(colado); return; }
    if (!arq) { limpar($("avisos-conferencia")).appendChild(el("p", { classe: "aviso", texto: "Escolha um arquivo .csv ou cole a tabela no campo acima." })); return; }
    if (arq.size > 5 * 1024 * 1024) { limpar($("avisos-conferencia")).appendChild(el("p", { classe: "aviso", texto: "Arquivo grande demais (mais de 5 MB). Envie só as colunas de código e preço." })); return; }
    arq.arrayBuffer().then(function (buf) {
      var texto;
      try { texto = new TextDecoder("utf-8", { fatal: true }).decode(buf); }
      catch (e) { texto = new TextDecoder("windows-1252").decode(buf); } // CSV salvo pelo Excel no Windows
      conferirTexto(texto);
    });
  });
  $("btn-exemplo").addEventListener("click", function () {
    fetch("exemplo.csv").then(function (r) { return r.text(); }).then(function (t) {
      $("colado").value = t.replace(/;/g, "\t").replace(/\r/g, "");
      conferirTexto(t);
    }).catch(falhou);
  });
  $("arquivo").addEventListener("change", function () { if (this.files[0]) $("colado").value = ""; });
  $("btn-baixar-resultado").addEventListener("click", baixarResultado);
  $("aliquota").addEventListener("change", function () {
    if (estado.resultado) $("btn-conferir").click();
    if ($("busca").value) buscar();
  });
  $("busca").addEventListener("input", function () { clearTimeout(esperaBusca); esperaBusca = setTimeout(buscar, 200); });

  baixarJSON("dados/resumo.json").then(function (r) {
    mostrarResumo(r);
    return Promise.all([baixarJSON("dados/catalogo.json"), carregarPMC("19")]);
  }).then(function (v) { listaPronta(v[0]); }).catch(falhou);
})();
