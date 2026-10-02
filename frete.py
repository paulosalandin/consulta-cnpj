"""Cálculo de frete simulado por peso e região de destino (origem: São Paulo/SP).

Não consulta a API dos Correios nem serviços de terceiros: os valores e prazos
são estimativas fixas por região, definidas nas tabelas internas abaixo, com
fins didáticos (Documentação do Sistema, seções 4.4 e 9).
"""

# Seção 9.1 - peso unitário por peça (kg)
PESO_UNITARIO_KG = {
    "camisa": 0.20,
    "calca": 0.50,
    "blusa": 0.30,
}

# Seção 9.2 - tarifa por região
TABELA_REGIOES = {
    "sudeste":      {"ufs": ("SP", "RJ", "MG", "ES"),
                     "ate_1kg": 14.90, "kg_extra": 4.50, "prazo_dias": 3},
    "sul":          {"ufs": ("PR", "SC", "RS"),
                     "ate_1kg": 19.90, "kg_extra": 6.00, "prazo_dias": 5},
    "centro-oeste": {"ufs": ("GO", "MT", "MS", "DF"),
                     "ate_1kg": 22.90, "kg_extra": 7.00, "prazo_dias": 6},
    "nordeste":     {"ufs": ("BA", "PE", "CE", "MA", "PB", "RN", "AL", "SE", "PI"),
                     "ate_1kg": 26.90, "kg_extra": 8.50, "prazo_dias": 7},
    "norte":        {"ufs": ("AM", "PA", "AC", "RO", "RR", "AP", "TO"),
                     "ate_1kg": 34.90, "kg_extra": 11.00, "prazo_dias": 10},
}

REGIAO_POR_UF = {
    uf: regiao for regiao, dados in TABELA_REGIOES.items() for uf in dados["ufs"]
}

# UF fora da tabela: usa a tarifa da região mais cara, para nunca subestimar o frete.
REGIAO_PADRAO = "norte"


def calcular_peso_total(qtd_camisa: int, qtd_calca: int, qtd_blusa: int) -> float:
    """Soma o peso estimado (kg) a partir da quantidade de cada tipo de peça."""
    peso = (
        qtd_camisa * PESO_UNITARIO_KG["camisa"]
        + qtd_calca * PESO_UNITARIO_KG["calca"]
        + qtd_blusa * PESO_UNITARIO_KG["blusa"]
    )
    return round(peso, 3)


def regiao_da_uf(uf: str) -> str:
    """Mapeia a UF de destino para uma das cinco regiões; UF desconhecida cai em REGIAO_PADRAO."""
    return REGIAO_POR_UF.get((uf or "").strip().upper(), REGIAO_PADRAO)


def calcular_frete(uf: str, peso_kg: float) -> dict:
    """Valor e prazo pela região de destino e peso total (seção 9.3).

    valor = tarifa_ate_1kg + max(0, peso - 1.0) * tarifa_kg_extra
    """
    regiao = regiao_da_uf(uf)
    tabela = TABELA_REGIOES[regiao]
    kg_extra = max(0.0, peso_kg - 1.0)
    valor = tabela["ate_1kg"] + kg_extra * tabela["kg_extra"]
    return {
        "peso_kg": round(peso_kg, 3),
        "valor": round(valor, 2),
        "prazo_dias": tabela["prazo_dias"],
        "regiao": regiao,
    }
