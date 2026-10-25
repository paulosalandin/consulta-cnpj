// Validador de CNPJ executado no navegador (JavaScript).
// Formata o número enquanto a pessoa digita e confere os dígitos verificadores
// (módulo 11), sem enviar nada ao servidor.

function limparCnpj(valor) {
  return String(valor).replace(/\D/g, "").slice(0, 14);
}

function formatarCnpj(digitos) {
  const d = limparCnpj(digitos);
  let r = d.slice(0, 2);
  if (d.length > 2) r += "." + d.slice(2, 5);
  if (d.length > 5) r += "." + d.slice(5, 8);
  if (d.length > 8) r += "/" + d.slice(8, 12);
  if (d.length > 12) r += "-" + d.slice(12, 14);
  return r;
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

function mensagemStatus(valor) {
  const d = limparCnpj(valor);
  if (d.length === 0) return { tipo: "neutro", texto: "Digite um CNPJ para conferir." };
  if (d.length < 14) {
    return { tipo: "neutro", texto: "Faltam " + (14 - d.length) + " dígito(s)." };
  }
  return validarCnpj(d)
    ? { tipo: "ok", texto: "✔ CNPJ válido (dígitos verificadores corretos)." }
    : { tipo: "erro", texto: "✖ CNPJ inválido: os dígitos verificadores não conferem." };
}

function iniciarValidador() {
  const campo = document.getElementById("cnpj-campo");
  const status = document.getElementById("cnpj-status");
  if (!campo || !status) return;
  campo.addEventListener("input", function () {
    campo.value = formatarCnpj(campo.value);
    const m = mensagemStatus(campo.value);
    status.textContent = m.texto;
    status.className = m.tipo;
    campo.setAttribute("aria-invalid", m.tipo === "erro" ? "true" : "false");
  });
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = { limparCnpj, formatarCnpj, validarCnpj, mensagemStatus };
} else {
  iniciarValidador();
}
