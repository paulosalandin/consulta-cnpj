"""Funções puras (sem dependência de Streamlit) usadas pelo app e pelos testes."""
import requests


def formatar_cnpj(cnpj: str) -> str:
    return f"{cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:8]}/{cnpj[8:12]}-{cnpj[12:]}"


def cor_situacao(situacao: str) -> str:
    s = situacao.upper()
    if s == "ATIVA":
        return "🟢"
    elif s == "BAIXADA":
        return "🔴"
    elif s in ("SUSPENSA", "INAPTA"):
        return "🟡"
    else:
        return "⚪"


def extrair_cnpjs(raws: list) -> tuple[list, list]:
    import re
    validos, invalidos = [], []
    for raw in raws:
        limpo = re.sub(r'\D', '', str(raw).strip())
        if len(limpo) == 14:
            validos.append(limpo)
        elif limpo:
            invalidos.append(raw)
    return validos, invalidos


def consultar_cnpj(cnpj: str) -> dict:
    """Consulta um único CNPJ na BrasilAPI e devolve dados normalizados."""
    cnpj_fmt = formatar_cnpj(cnpj)
    try:
        url = f"https://brasilapi.com.br/api/cnpj/v1/{cnpj}"
        resp = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=15)

        if resp.status_code == 200:
            data = resp.json()
            return {
                "cnpj_fmt":  cnpj_fmt,
                "nome":      data.get("razao_social", "N/A"),
                "situacao":  data.get("descricao_situacao_cadastral", "Desconhecida"),
                "municipio": data.get("municipio", ""),
                "uf":        data.get("uf", ""),
                "cep":       data.get("cep", ""),
                "atividade": data.get("cnae_fiscal_descricao", ""),
            }
        elif resp.status_code == 404:
            return {"cnpj_fmt": cnpj_fmt, "nome": "Não encontrado", "situacao": "Não encontrado",
                    "municipio": "", "uf": "", "cep": "", "atividade": ""}
        else:
            return {"cnpj_fmt": cnpj_fmt, "nome": "Erro", "situacao": f"HTTP {resp.status_code}",
                    "municipio": "", "uf": "", "cep": "", "atividade": ""}
    except requests.RequestException as e:
        return {"cnpj_fmt": cnpj_fmt, "nome": "Erro", "situacao": f"Falha na consulta: {str(e)[:60]}",
                "municipio": "", "uf": "", "cep": "", "atividade": ""}
