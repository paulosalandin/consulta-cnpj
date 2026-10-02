"""Testes automatizados (pytest) das funções puras de cnpj_core.py e frete.py.

Executar a partir da raiz do projeto:
    pytest tests/test_cnpj_app.py -v
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pytest

from cnpj_core import cor_situacao, extrair_cnpjs, formatar_cnpj
from frete import calcular_frete, calcular_peso_total, regiao_da_uf


def test_formatar_cnpj():
    assert formatar_cnpj("00360305000104") == "00.360.305/0001-04"


@pytest.mark.parametrize("situacao, icone", [
    ("ATIVA", "🟢"),
    ("ativa", "🟢"),
    ("BAIXADA", "🔴"),
    ("SUSPENSA", "🟡"),
    ("INAPTA", "🟡"),
    ("Não encontrado", "⚪"),
])
def test_cor_situacao(situacao, icone):
    assert cor_situacao(situacao) == icone


def test_extrair_cnpjs_validos_e_invalidos():
    validos, invalidos = extrair_cnpjs(["00.360.305/0001-04", "12345", "", "00360305000104"])
    assert validos == ["00360305000104", "00360305000104"]
    assert invalidos == ["12345"]


def test_calcular_peso_total():
    assert calcular_peso_total(0, 0, 0) == 0
    assert calcular_peso_total(1, 0, 0) == pytest.approx(0.20)
    assert calcular_peso_total(2, 1, 3) == pytest.approx(0.40 + 0.50 + 0.90)


def test_regiao_da_uf_conhecida_e_desconhecida():
    assert regiao_da_uf("SP") == "sudeste"
    assert regiao_da_uf("pr") == "sul"
    assert regiao_da_uf("DF") == "centro-oeste"
    assert regiao_da_uf("BA") == "nordeste"
    assert regiao_da_uf("AM") == "norte"
    # UF fora da tabela não quebra: cai na região padrão
    assert regiao_da_uf("XX") == "norte"
    assert regiao_da_uf("") == "norte"


def test_calcular_frete_local_barato_que_norte():
    sudeste = calcular_frete("SP", 1.0)
    norte = calcular_frete("AM", 1.0)
    assert sudeste["valor"] == 14.90
    assert norte["valor"] == 34.90
    assert sudeste["valor"] < norte["valor"]
    assert sudeste["prazo_dias"] < norte["prazo_dias"]


def test_calcular_frete_peso_extra_aumenta_valor():
    ate_1kg = calcular_frete("SP", 1.0)
    com_extra = calcular_frete("SP", 3.0)
    # 2 kg extras * R$ 4,50 acima da tarifa base
    assert com_extra["valor"] == pytest.approx(ate_1kg["valor"] + 2 * 4.50)
    assert com_extra["regiao"] == "sudeste"
    assert com_extra["peso_kg"] == 3.0
