// Testes do conferidor (node --test testes/)
import { test } from "node:test";
import assert from "node:assert/strict";
import { createRequire } from "node:module";

const C = createRequire(import.meta.url)("../public/conferidor.js");

// catálogo mínimo no formato de public/dados/catalogo.json
// flags: 1 = uso hospitalar, 2 = isento de ICMS, 4 = em análise de recurso
const catalogo = {
  itens: [
    ["1", ["7891106000956"], "BAYCUTEN N", "CREM X 40 G", 0, 0, 0, 0],
    ["2", ["7896016806469"], "ORENCIA", "250 MG", 0, 0, 0, 1],
    ["3", ["7891010249908"], "HIPOGLOS", "45 G", 0, 0, 0, 2],
    ["4", ["7891010249908"], "HIPOGLOS", "135 G", 0, 0, 0, 0],
    ["5", ["7896523206943"], "BABYMED", "45 G", 0, 0, 0, 4],
    ["6", ["0789110600095"], "COM ZERO NA FRENTE", "X", 0, 0, 0, 0],
  ],
};
const pmc19 = [5090, null, 3053, 7233, 1918, 1000];
const pmc0 = [4122, null, 2500, 6000, 1500, 800];

function conferir(texto) {
  const t = C.lerTabela(texto);
  return C.conferir(t.linhas, catalogo, pmc19, pmc0);
}

test("preços nos formatos que aparecem em tabela de farmácia", () => {
  assert.equal(C.lerPreco("49,90"), 4990);
  assert.equal(C.lerPreco("R$ 1.234,56"), 123456);
  assert.equal(C.lerPreco("1234.56"), 123456);
  assert.equal(C.lerPreco("1,234.56"), 123456);
  assert.equal(C.lerPreco("12.9"), 1290);
  assert.equal(C.lerPreco("1.234"), 123400); // no Brasil, ponto com 3 casas é milhar
  assert.equal(C.lerPreco("12"), 1200);
  assert.equal(C.lerPreco(""), null);
  assert.equal(C.lerPreco("abc"), null);
  assert.equal(C.lerPreco("1,2,3"), null);
  assert.equal(C.lerPreco("12,999"), null);
});

test("código de barras estragado pelo Excel é detectado, não adivinhado", () => {
  assert.match(C.lerCodigo("7,89111E+12").problema, /Excel/);
  assert.match(C.lerCodigo("7.89111e+12").problema, /Excel/);
  assert.equal(C.lerCodigo('="7891106000956"').codigo, "7891106000956");
  assert.equal(C.lerCodigo(" 7891106000956 ").codigo, "7891106000956");
  // Excel também come o zero à esquerda de um EAN-13: 12 dígitos voltam a ser 13
  assert.equal(C.lerCodigo("789110600095").codigo, "0789110600095");
});

test("acha as colunas pelo cabeçalho, com ; , ou tab", () => {
  for (const sep of [";", ",", "\t"]) {
    const t = C.lerTabela(["EAN", "Produto", "Preço"].join(sep) + "\n" + ["7891106000956", "Baycuten", sep === "," ? '"49,90"' : "49,90"].join(sep));
    assert.deepEqual(t.linhas.map((l) => [l.codigoBruto, l.precoBruto, l.descricao]), [["7891106000956", "49,90", "Baycuten"]], sep);
  }
});

test("com coluna 'Código' (interno) e 'EAN', usa o EAN", () => {
  const t = C.lerTabela("Código;EAN;Descrição;Preço de venda\n1001;7891106000956;Baycuten;49,90\n");
  assert.equal(t.linhas[0].codigoBruto, "7891106000956");
});

test("sem cabeçalho, acha as colunas pelo conteúdo", () => {
  const t = C.lerTabela("Baycuten;7891106000956;49,90\nOrencia;7896016806469;100,00\n");
  assert.equal(t.linhas.length, 2);
  assert.equal(t.linhas[0].codigoBruto, "7891106000956");
  assert.equal(t.linhas[0].precoBruto, "49,90");
});

test("avisa quando a coluna de código não tem cara de código de barras", () => {
  const t = C.lerTabela("Código;Preço\n1001;10,00\n1002;11,00\n");
  assert.equal(t.avisos.length, 1);
});

test("situações da conferência", () => {
  const r = conferir([
    "EAN;Produto;Preço",
    "7891106000956;dentro;50,90",       // igual ao PMC: ok
    "7891106000956;acima;51,00",
    "7896016806469;hospitalar;10,00",
    "7899999999999;fora da lista;10,00",
    "7,89111E+12;estragado;10,00",
    "7891106000956;sem preço;",
    "7896523206943;em recurso;10,00",
  ].join("\n"));
  assert.deepEqual(r.map((x) => x.status), ["ok", "acima", "hospitalar", "nao_encontrado", "codigo_estragado", "sem_preco", "ok"]);
  assert.equal(r[1].excesso, 10);
  assert.match(r[6].avisos.join(), /recurso/);
});

test("mesmo código de barras em duas apresentações: vale o menor PMC e avisa", () => {
  const [r] = conferir("EAN;Preço\n7891010249908;40,00\n");
  assert.equal(r.pmc, 3053);
  assert.equal(r.status, "acima");
  assert.match(r.avisos.join(), /2 apresentações/);
});

test("isento de ICMS: preço entre o PMC 0% e o da alíquota vira atenção", () => {
  const cat = { itens: [["3", ["7891010249908"], "HIPOGLOS", "45 G", 0, 0, 0, 2]] };
  const t = C.lerTabela("EAN;Preço\n7891010249908;28,00\n7891010249908;24,00\n");
  const r = C.conferir(t.linhas, cat, [3053], [2500]);
  assert.equal(r[0].status, "atencao");
  assert.equal(r[1].status, "ok");
});

test("CSV do resultado: abre no Excel com acento e sem virar fórmula", () => {
  const r = conferir('EAN;Produto;Preço\n7891106000956;=HYPERLINK("http://x");60,00\n');
  const csv = C.paraCSV(r, catalogo);
  assert.ok(csv.startsWith("﻿"));
  assert.ok(csv.includes(`"'=HYPERLINK(""http://x"")"`));
  assert.ok(csv.includes("ACIMA do PMC"));
});

test("reais", () => {
  assert.equal(C.reais(123456), "R$ 1.234,56");
  assert.equal(C.reais(-5), "-R$ 0,05");
  assert.equal(C.reais(null), "—");
});
