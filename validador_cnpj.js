// Contador de CNPJs da aba "Colar Lista", executado no navegador (JavaScript).
// Enquanto a pessoa digita ou cola, confere cada linha (formato e dígitos
// verificadores, módulo 11) e mostra o resumo, sem enviar nada ao servidor.

function limparCnpj(valor) {
  return String(valor).replace(/\D/g, "");
}

function calcularDigito(base) {
  // Pesos de 2 a 9, aplicados da direita para a esquerda, repetindo o ciclo.
  let soma = 0;
  let peso = 2;
  for (let i = base.length - 1; i >= 0; i--) {
    soma += Number(base[i]) * peso;
    peso = peso === 9 ? 2 : peso + 1;
  }
  const resto = soma % 11;
  return resto < 2 ? 0 : 11 - resto;
}

function validarCnpj(valor) {
  const d = limparCnpj(valor);
  if (d.length !== 14) return false;
  if (/^(\d)\1{13}$/.test(d)) return false; // 00000000000000, 11111111111111...
  const dv1 = calcularDigito(d.slice(0, 12));
  const dv2 = calcularDigito(d.slice(0, 12) + dv1);
  return d.endsWith(String(dv1) + String(dv2));
}

// Classifica cada linha não vazia do texto. Os números de linha são os do campo.
function analisarLista(texto) {
  const r = { total: 0, validos: 0, formatoInvalido: [], digitoInvalido: [] };
  String(texto).split(/\r?\n/).forEach(function (linha, i) {
    if (!linha.trim()) return;
    r.total++;
    const d = limparCnpj(linha);
    if (d.length !== 14) r.formatoInvalido.push(i + 1);
    else if (!validarCnpj(d)) r.digitoInvalido.push(i + 1);
    else r.validos++;
  });
  return r;
}

function plural(n, singular, pluralTxt) {
  return n + " " + (n === 1 ? singular : pluralTxt);
}

function listarLinhas(linhas) {
  const mostra = linhas.slice(0, 8).join(", ");
  return (linhas.length === 1 ? "linha " : "linhas ") + (linhas.length > 8 ? mostra + "…" : mostra);
}

function resumoLista(texto) {
  const a = analisarLista(texto);
  if (a.total === 0) return { tipo: "neutro", texto: "Cole os CNPJs para conferir cada linha." };
  const partes = [plural(a.validos, "válido", "válidos")];
  if (a.formatoInvalido.length) {
    partes.push(plural(a.formatoInvalido.length, "com formato errado", "com formato errado") +
      " (" + listarLinhas(a.formatoInvalido) + ")");
  }
  if (a.digitoInvalido.length) {
    partes.push(plural(a.digitoInvalido.length, "com dígito verificador errado", "com dígito verificador errado") +
      " (" + listarLinhas(a.digitoInvalido) + ")");
  }
  const problemas = a.formatoInvalido.length + a.digitoInvalido.length;
  return {
    tipo: problemas ? "erro" : "ok",
    texto: (problemas ? "⚠ " : "✔ ") + plural(a.total, "linha", "linhas") + ": " + partes.join("; ") + ".",
  };
}

function iniciarContador() {
  const status = document.getElementById("cnpj-status");
  if (!status) return;
  let campo = null;
  let timer = null;

  function atualizar() {
    const m = resumoLista(campo.value);
    status.textContent = m.texto;
    status.className = m.tipo;
  }

  function ligar(t) {
    campo = t;
    campo.addEventListener("input", function () {
      clearTimeout(timer);
      timer = setTimeout(atualizar, 400); // evita anunciar a cada tecla
    });
    atualizar();
  }

  // O campo é criado pelo Streamlit depois deste quadro: procura por alguns segundos.
  let tentativas = 0;
  const busca = setInterval(function () {
    tentativas++;
    let t = null;
    try {
      t = window.parent.document.querySelector('textarea[aria-label="Lista de CNPJs para consulta"]');
    } catch (e) {
      /* sem acesso ao documento pai: o contador fica desligado */
    }
    if (t) { clearInterval(busca); ligar(t); }
    else if (tentativas > 40) { clearInterval(busca); status.textContent = ""; }
  }, 250);
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = { limparCnpj, validarCnpj, analisarLista, resumoLista };
} else {
  iniciarContador();
}
