/*
 * Conferidor de preços: compara a tabela de preços da farmácia com o PMC da
 * lista oficial. Roda inteiro no navegador — o arquivo da farmácia não é
 * enviado para lugar nenhum.
 *
 * Funções puras (sem tocar na página), testadas em testes/conferidor.test.mjs.
 */
(function (raiz) {
  "use strict";

  function semAcento(s) {
    return String(s).normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase().trim();
  }

  // ---------------------------------------------------------------- leitura

  /** Separa o texto em linhas e células, respeitando aspas ("a;b" é uma célula). */
  function lerCSV(texto, sep) {
    var linhas = [], linha = [], cel = "", aspas = false;
    for (var i = 0; i < texto.length; i++) {
      var c = texto[i];
      if (aspas) {
        if (c === '"' && texto[i + 1] === '"') { cel += '"'; i++; }
        else if (c === '"') aspas = false;
        else cel += c;
      } else if (c === '"' && cel === "") {
        aspas = true;
      } else if (c === sep) {
        linha.push(cel); cel = "";
      } else if (c === "\n" || c === "\r") {
        if (c === "\r" && texto[i + 1] === "\n") i++;
        linha.push(cel); cel = "";
        linhas.push(linha); linha = [];
      } else {
        cel += c;
      }
    }
    if (cel !== "" || linha.length) { linha.push(cel); linhas.push(linha); }
    return linhas.filter(function (l) { return l.some(function (x) { return x.trim() !== ""; }); });
  }

  /** Tab (colado do Excel), ponto e vírgula (CSV brasileiro) ou vírgula. */
  function adivinharSeparador(texto) {
    var primeira = texto.split(/\r?\n/, 1)[0];
    if (primeira.indexOf("\t") >= 0) return "\t";
    var pv = (primeira.match(/;/g) || []).length;
    var v = (primeira.replace(/"[^"]*"/g, "").match(/,/g) || []).length;
    return pv >= v && pv > 0 ? ";" : v > 0 ? "," : ";";
  }

  // nomes de coluna aceitos, do mais específico para o mais genérico: se a
  // tabela tiver "Código" (código interno) e "EAN", vale o EAN
  var NOMES = {
    codigo: [/^(ean|ean ?13|gtin|cod(igo)?\.? ?(de )?barras?|cod(igo)?\.? ?ean|barcode|barras)$/, /^cod(igo)?\.?$/],
    preco: [/^(pre[cç]o|pre[cç]o (de )?venda|pre[cç]o final|pre[cç]o r\$|valor|valor (de )?venda|pv|venda|r\$)$/],
    descricao: [/^(produto|descri[cç][aã]o|nome|medicamento|item)$/]
  };

  function pareceCodigo(v) { return /^\s*=?"?\d{8,14}"?\s*$/.test(v) || /^\s*\d[.,]\d+e\+?\d+\s*$/i.test(v); }
  function parecePreco(v) { return !/^\s*\d{6,}\s*$/.test(v) && lerPreco(v) !== null; }

  function acharColuna(cab, padroes) {
    for (var k = 0; k < padroes.length; k++) {
      var i = cab.findIndex(function (c) { return padroes[k].test(c); });
      if (i >= 0) return i;
    }
    return -1;
  }

  /**
   * Texto do arquivo (ou colado) -> {linhas: [{numero, codigoBruto, precoBruto, descricao}], avisos}.
   * Acha as colunas pelo nome do cabeçalho; sem cabeçalho, pelo conteúdo.
   */
  function lerTabela(texto) {
    texto = String(texto || "").replace(/^\uFEFF/, "");
    var sep = adivinharSeparador(texto);
    var linhas = lerCSV(texto, sep);
    var avisos = [];
    if (!linhas.length) return { linhas: [], avisos: ["O arquivo está vazio."] };

    var inicio = 1, col;
    if (linhas[0].some(pareceCodigo)) {
      // sem cabeçalho: a primeira linha já é produto; acha as colunas pelo conteúdo
      inicio = 0;
      var amostra = linhas[0];
      var c = amostra.findIndex(pareceCodigo);
      var p = amostra.findIndex(function (v, i) { return i !== c && parecePreco(v); });
      var d = amostra.findIndex(function (v, i) { return i !== c && i !== p && /[a-z]{3}/i.test(v); });
      col = { codigo: c, preco: p, descricao: d };
    } else {
      var cab = linhas[0].map(function (x) { return semAcento(x).replace(/[_\s]+/g, " ").replace(/[():]/g, "").trim(); });
      col = { codigo: acharColuna(cab, NOMES.codigo), preco: acharColuna(cab, NOMES.preco), descricao: acharColuna(cab, NOMES.descricao) };
    }
    if (col.codigo < 0 || col.preco < 0) {
      return { linhas: [], avisos: ["Não achei as colunas de código de barras e de preço. Use um cabeçalho como 'EAN;Produto;Preço'."] };
    }
    var saida = [];
    for (var i = inicio; i < linhas.length; i++) {
      var l = linhas[i];
      saida.push({
        numero: i + 1,
        codigoBruto: (l[col.codigo] || "").trim(),
        precoBruto: (l[col.preco] || "").trim(),
        descricao: col.descricao >= 0 ? (l[col.descricao] || "").trim() : ""
      });
    }
    var comCara = saida.filter(function (x) { return pareceCodigo(x.codigoBruto); }).length;
    if (saida.length && comCara < saida.length / 2) {
      avisos.push("A coluna usada como código de barras não parece ter códigos de barras (8 a 14 dígitos) na maioria das linhas. Confira se é a coluna certa.");
    }
    return { linhas: saida, avisos: avisos, separador: sep, colunas: col };
  }

  function gtinValido(c) {
    if (!/^\d+$/.test(c) || [8, 12, 13, 14].indexOf(c.length) < 0) return false;
    var soma = 0, digitos = c.slice(0, -1).split("").reverse();
    digitos.forEach(function (d, i) { soma += Number(d) * (i % 2 === 0 ? 3 : 1); });
    return (10 - (soma % 10)) % 10 === Number(c[c.length - 1]);
  }

  /** '7891106000956', '="789..."', ' 789.110.600.0956 ' -> dígitos; '7,89111E+12' -> problema. */
  function lerCodigo(bruto) {
    var t = String(bruto || "").trim().replace(/^="?|"$/g, "");
    if (!t) return { codigo: "", problema: "sem código" };
    if (/^\d[.,]\d+e\+?\d+$/i.test(t)) {
      return { codigo: "", problema: "estragado pelo Excel (" + t + ")" };
    }
    var d = t.replace(/\D/g, "");
    if (!d) return { codigo: "", problema: "sem código" };
    // o Excel também apaga zeros à esquerda: um EAN-13 que começa com 0 vira
    // 12 dígitos. O zero não muda o dígito verificador, então é seguro repor.
    if (d.length === 12) d = "0" + d;
    return { codigo: d, problema: null };
  }

  /** 'R$ 1.234,56' / '1234,56' / '1234.56' / '12' -> centavos (inteiro); ilegível -> null. */
  function lerPreco(bruto) {
    var t = String(bruto || "").replace(/r\$|\s| /gi, "");
    if (!t || !/^\d[\d.,]*$/.test(t)) return null;
    var ponto = t.lastIndexOf("."), virgula = t.lastIndexOf(",");
    var inteiro, frac;
    if (ponto >= 0 && virgula >= 0) {
      var dec = Math.max(ponto, virgula);
      inteiro = t.slice(0, dec).replace(/[.,]/g, ""); frac = t.slice(dec + 1);
    } else if (virgula >= 0) {
      if ((t.match(/,/g) || []).length > 1) return null;
      inteiro = t.slice(0, virgula); frac = t.slice(virgula + 1);
    } else if (ponto >= 0) {
      // "1.234" no Brasil é mil duzentos e trinta e quatro; "12.9" / "12.90" é decimal
      var partes = t.split(".");
      if (partes.length > 2 || partes[1].length === 3) { inteiro = partes.join(""); frac = ""; }
      else { inteiro = partes[0]; frac = partes[1]; }
    } else {
      inteiro = t; frac = "";
    }
    if (!/^\d+$/.test(inteiro || "0") || !/^\d{0,2}$/.test(frac)) return null;
    return Number(inteiro || "0") * 100 + Number((frac + "00").slice(0, 2));
  }

  // ---------------------------------------------------------------- conferência

  function indiceDeCodigos(catalogo) {
    var mapa = new Map();
    catalogo.itens.forEach(function (item, i) {
      item[1].forEach(function (ean) {
        if (!mapa.has(ean)) mapa.set(ean, []);
        if (mapa.get(ean).indexOf(i) < 0) mapa.get(ean).push(i);
      });
    });
    return mapa;
  }

  /**
   * Confere cada linha. pmc: PMC (centavos) da alíquota escolhida, alinhado
   * com catalogo.itens; pmc0: PMC da alíquota 0% (para isentos de ICMS).
   */
  function conferir(linhas, catalogo, pmc, pmc0, indice) {
    indice = indice || indiceDeCodigos(catalogo);
    return linhas.map(function (l) {
      var r = { numero: l.numero, descricao: l.descricao, codigoBruto: l.codigoBruto, precoBruto: l.precoBruto,
                codigo: "", preco: null, status: "", pmc: null, excesso: null, itens: [], avisos: [] };
      var c = lerCodigo(l.codigoBruto);
      r.codigo = c.codigo;
      r.preco = lerPreco(l.precoBruto);
      if (c.problema && c.problema.indexOf("Excel") >= 0) { r.status = "codigo_estragado"; r.avisos.push("Código " + c.problema + ": formate a coluna como texto e exporte de novo."); return r; }
      if (!r.codigo) { r.status = "sem_codigo"; return r; }
      if (r.preco === null) {
        r.status = "sem_preco";
        r.avisos.push(l.precoBruto ? "Não entendi o preço " + JSON.stringify(l.precoBruto) + "." : "Preço em branco.");
        return r;
      }
      var achados = indice.get(r.codigo) || [];
      r.itens = achados;
      if (!achados.length) { r.status = "nao_encontrado"; return r; }
      var comPmc = achados.filter(function (i) { return pmc[i] !== null && pmc[i] !== undefined; });
      if (!comPmc.length) { r.status = "hospitalar"; return r; }
      // mesmo código de barras em mais de uma apresentação: vale o teto menor
      var teto = Math.min.apply(null, comPmc.map(function (i) { return pmc[i]; }));
      if (achados.length > 1) {
        var distintos = comPmc.map(function (i) { return pmc[i]; }).filter(function (v, k, a) { return a.indexOf(v) === k; });
        r.avisos.push("Este código de barras aparece em " + achados.length + " apresentações da lista" +
          (distintos.length > 1 ? ", com PMC diferentes; usei o menor." : "."));
      }
      r.pmc = teto;
      r.excesso = r.preco - teto;
      r.status = r.preco > teto ? "acima" : "ok";
      var isento = comPmc.filter(function (i) { return catalogo.itens[i][7] & 2; });
      if (isento.length && pmc0) {
        var teto0 = Math.min.apply(null, isento.map(function (i) { return pmc0[i]; }));
        if (teto0 < teto) {
          r.pmc0 = teto0;
          r.avisos.push("Isento de ICMS: se a isenção vale no seu estado, o teto é o PMC 0% (" + reais(teto0) + ")" +
            (r.preco > teto0 && r.status === "ok" ? " — e o seu preço passa dele." : "."));
          if (r.preco > teto0 && r.status === "ok") r.status = "atencao";
        }
      }
      if (achados.some(function (i) { return catalogo.itens[i][7] & 4; })) {
        r.avisos.push("O preço deste produto está em análise de recurso na CMED.");
      }
      return r;
    });
  }

  function reais(centavos) {
    if (centavos === null || centavos === undefined) return "—";
    var s = (Math.abs(centavos) / 100).toFixed(2).split(".");
    return (centavos < 0 ? "-" : "") + "R$ " + s[0].replace(/\B(?=(\d{3})+(?!\d))/g, ".") + "," + s[1];
  }

  // ---------------------------------------------------------------- exportação

  /** Protege contra injeção de fórmula ao abrir o CSV no Excel. */
  function celulaSegura(v) {
    var s = String(v === null || v === undefined ? "" : v);
    if (/^[=+\-@\t\r]/.test(s)) s = "'" + s;
    return /[;"\n\r]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  }

  var SITUACAO = {
    acima: "ACIMA do PMC", atencao: "Atenção (isento de ICMS)", ok: "Dentro do PMC",
    hospitalar: "Uso hospitalar (sem PMC)", nao_encontrado: "Código não está na lista",
    codigo_estragado: "Código estragado pelo Excel", sem_codigo: "Sem código", sem_preco: "Sem preço válido"
  };

  function paraCSV(resultados, catalogo) {
    var linhas = [["Linha", "Produto (sua tabela)", "Código de barras", "Seu preço", "PMC", "Diferença", "Situação", "Produto na lista CMED", "Observações"]];
    resultados.forEach(function (r) {
      var nome = r.itens.length ? r.itens.map(function (i) { return catalogo.itens[i][2] + " " + catalogo.itens[i][3]; }).join(" | ") : "";
      linhas.push([r.numero, r.descricao, r.codigo || r.codigoBruto, r.preco === null ? r.precoBruto : reais(r.preco),
        reais(r.pmc), r.excesso === null ? "" : reais(r.excesso), SITUACAO[r.status] || r.status, nome, r.avisos.join(" ")]);
    });
    return "﻿" + linhas.map(function (l) { return l.map(celulaSegura).join(";"); }).join("\r\n") + "\r\n";
  }

  var api = { lerCSV: lerCSV, lerTabela: lerTabela, lerCodigo: lerCodigo, lerPreco: lerPreco, gtinValido: gtinValido,
              indiceDeCodigos: indiceDeCodigos, conferir: conferir, reais: reais, paraCSV: paraCSV,
              celulaSegura: celulaSegura, SITUACAO: SITUACAO, semAcento: semAcento };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else raiz.Conferidor = api;
})(this);
