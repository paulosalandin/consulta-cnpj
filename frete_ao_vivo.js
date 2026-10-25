// Estimativa de frete ao vivo, executada no navegador (JavaScript).
// Recalcula peso, valor e prazo enquanto a pessoa altera as quantidades, sem
// esperar o botão "Calcular Frete". As tabelas (pesos, tarifas, regiões) vêm do
// frete.py, injetadas pelo servidor em CONFIG_FRETE: aqui só existe a fórmula.
// O botão continua sendo o que grava o cálculo no histórico.

function quantidade(valor) {
  const n = parseInt(valor, 10);
  return Number.isFinite(n) && n > 0 ? n : 0;
}

function regiaoDaUf(config, uf) {
  const chave = String(uf || "").trim().toUpperCase();
  return config.regiaoPorUf[chave] || config.regiaoPadrao;
}

// Mesma fórmula do frete.py: valor = tarifa_ate_1kg + max(0, peso - 1) * tarifa_kg_extra
function calcularEstimativa(config, qtdCamisa, qtdCalca, qtdBlusa, uf) {
  const peso = Number((
    qtdCamisa * config.pesoUnitario.camisa +
    qtdCalca * config.pesoUnitario.calca +
    qtdBlusa * config.pesoUnitario.blusa
  ).toFixed(3));
  const regiao = regiaoDaUf(config, uf);
  const t = config.regioes[regiao];
  const kgExtra = Math.max(0, peso - 1);
  return {
    pecas: qtdCamisa + qtdCalca + qtdBlusa,
    peso: peso,
    valor: Number((t.ate_1kg + kgExtra * t.kg_extra).toFixed(2)),
    prazo: t.prazo_dias,
    regiao: regiao,
  };
}

function formatarNumero(n, casas) {
  return n.toFixed(casas).replace(".", ",");
}

function textoEstimativa(e) {
  if (e.pecas === 0) return "Informe as quantidades para ver a estimativa ao vivo.";
  return "Estimativa ao vivo: " + formatarNumero(e.peso, 2) + " kg · R$ " + formatarNumero(e.valor, 2) +
    " · " + e.prazo + " dias úteis (região " + e.regiao + "). Valor simulado: clique em Calcular Frete para registrar.";
}

function iniciarFreteAoVivo() {
  const status = document.getElementById("frete-status");
  if (!status || typeof CONFIG_FRETE === "undefined") return;
  const rotulos = ["Camisas", "Calças", "Blusas"];
  let ultimo = "";

  function campo(rotulo) {
    try {
      return window.parent.document.querySelector('input[aria-label="' + rotulo + '"]');
    } catch (e) {
      return null; // sem acesso ao documento pai: a estimativa ao vivo fica desligada
    }
  }

  function atualizar() {
    const campos = rotulos.map(campo);
    if (campos.some(function (c) { return !c; })) return;
    const q = campos.map(function (c) { return quantidade(c.value); });
    const texto = textoEstimativa(calcularEstimativa(CONFIG_FRETE, q[0], q[1], q[2], CONFIG_FRETE.uf));
    if (texto !== ultimo) { // só anuncia quando muda
      ultimo = texto;
      status.textContent = texto;
    }
  }

  // Lê os campos com frequência: cobre digitação e os botões + / − do Streamlit.
  setInterval(atualizar, 250);
  atualizar();
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = { quantidade, regiaoDaUf, calcularEstimativa, textoEstimativa };
} else {
  iniciarFreteAoVivo();
}
