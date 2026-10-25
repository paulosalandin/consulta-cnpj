import streamlit as st
import streamlit.components.v1 as components
import base64
import io
import json
from pathlib import Path
import pandas as pd
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment

import cnpj_db as db
import frete
from cnpj_core import formatar_cnpj, cor_situacao, extrair_cnpjs, consultar_cnpj

# ── Configuração da página ──────────────────────────────────────────────────
st.set_page_config(
    page_title="Consulta CNPJ - Urbana",
    page_icon="🏢",
    layout="centered",
)

db.inicializar_db()


def _html_frete_ao_vivo(uf: str) -> str:
    """Estimativa de frete ao vivo (JavaScript no navegador, arquivo frete_ao_vivo.js).

    As tabelas vêm do frete.py, para existir uma única fonte de pesos e tarifas.
    """
    config = {
        "pesoUnitario": frete.PESO_UNITARIO_KG,
        "regioes": {
            nome: {k: v for k, v in dados.items() if k != "ufs"}
            for nome, dados in frete.TABELA_REGIOES.items()
        },
        "regiaoPorUf": frete.REGIAO_POR_UF,
        "regiaoPadrao": frete.REGIAO_PADRAO,
        "uf": uf,
    }
    js = (Path(__file__).parent / "frete_ao_vivo.js").read_text(encoding="utf-8")
    return f"""
<style>
  html {{ background: transparent; }}
  body {{ margin: 0; padding: 8px 10px; background: #ffffff; border-radius: 8px;
         font-family: "Source Sans Pro", sans-serif; font-size: 14px; color: #262730; }}
</style>
<div id="frete-status" role="status" aria-live="polite">Informe as quantidades para ver a estimativa ao vivo.</div>
<script>const CONFIG_FRETE = {json.dumps(config)};</script>
<script>{js}</script>
"""


st.markdown("""
<style>
    .main { background-color: #f8f9fa; }
    .stProgress > div > div { background-color: #1f77b4; }
</style>
""", unsafe_allow_html=True)

# ── Cabeçalho ───────────────────────────────────────────────────────────────
col_logo, col_titulo = st.columns([1, 4])
with col_logo:
    with open("urbana.png", "rb") as _f:
        _logo_b64 = base64.b64encode(_f.read()).decode()
    st.markdown(
        f'<img src="data:image/png;base64,{_logo_b64}" alt="Logo da Urbana" '
        'style="width:120px;max-width:100%;">',
        unsafe_allow_html=True,
    )
with col_titulo:
    st.title("Consulta CNPJ - Urbana")
    st.caption("Consulta situação cadastral na Receita Federal via BrasilAPI")
st.divider()

# ── Sidebar ─────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("ℹ️ Como usar")
    st.markdown("""
    **Opção 1 — Excel:**
    1. Arquivo **.xlsx** com CNPJs na coluna A (linha 2 em diante)
    2. Faça upload e clique em **Iniciar Consulta**

    **Opção 2 — Lista:**
    1. Cole os CNPJs na caixa de texto (um por linha)
    2. Clique em **Iniciar Consulta**
    """)
    st.divider()
    st.markdown("**API:** BrasilAPI (gratuita, sem limite fixo)")
    st.markdown("**Formato aceito:** com ou sem pontuação")
    st.caption("Ex: 12.345.678/0001-99 ou 12345678000199")
    st.divider()
    st.caption(
        "🔒 Os dados consultados (CNPJ, razão social, situação cadastral) são "
        "guardados localmente apenas para gerar o histórico da aplicação e "
        "excluídos automaticamente após 12 meses. Veja a aba **Privacidade**."
    )

# ── Funções auxiliares ───────────────────────────────────────────────────────
def gerar_excel(dados: list) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Resultado"

    cabecalhos = ["CNPJ", "Nome Empresarial", "Situação", "Município", "UF", "Atividade Principal"]
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)

    for col, titulo in enumerate(cabecalhos, 1):
        cell = ws.cell(row=1, column=col, value=titulo)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

    cores = {
        "ATIVA":    "C6EFCE",
        "BAIXADA":  "FFC7CE",
        "SUSPENSA": "FFEB9C",
        "INAPTA":   "FFEB9C",
    }

    for i, row in enumerate(dados, 2):
        situacao = str(row.get("situacao", "")).upper()
        fill = PatternFill(start_color=cores.get(situacao, "F2F2F2"),
                           end_color=cores.get(situacao, "F2F2F2"),
                           fill_type="solid")
        valores = [
            row.get("cnpj_fmt", ""),
            row.get("nome", ""),
            row.get("situacao", ""),
            row.get("municipio", ""),
            row.get("uf", ""),
            row.get("atividade", ""),
        ]
        for col, val in enumerate(valores, 1):
            ws.cell(row=i, column=col, value=val).fill = fill

    for col, larg in enumerate([22, 45, 12, 20, 6, 40], 1):
        ws.column_dimensions[ws.cell(row=1, column=col).column_letter].width = larg

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

def executar_consulta(cnpjs: list):
    resultados = []
    progress_bar = st.progress(0, text="Iniciando consulta...")
    tabela_box   = st.empty()

    for i, cnpj in enumerate(cnpjs):
        progress_bar.progress(i / len(cnpjs), text=f"Consultando {i+1}/{len(cnpjs)}: {formatar_cnpj(cnpj)}")

        r = consultar_cnpj(cnpj)
        resultados.append(r)
        db.salvar_consulta(r["cnpj_fmt"], r["nome"], r["situacao"], r["municipio"], r["uf"], r.get("cep", ""))

        df_parcial = pd.DataFrame([{
            "CNPJ":      r["cnpj_fmt"],
            "Nome":      r["nome"],
            "Situação":  f"{cor_situacao(r['situacao'])} {r['situacao']}",
            "Município": r.get("municipio", ""),
            "UF":        r.get("uf", ""),
        } for r in resultados])
        tabela_box.dataframe(df_parcial, use_container_width=True, hide_index=True)

    progress_bar.progress(1.0, text="✅ Consulta finalizada!")
    return resultados

def mostrar_resultados(resultados: list):
    st.divider()
    st.subheader("📊 Resumo")

    contagem = {}
    for r in resultados:
        s = r["situacao"].upper()
        contagem[s] = contagem.get(s, 0) + 1

    cols = st.columns(len(contagem))
    for idx, (sit, qtd) in enumerate(sorted(contagem.items())):
        cols[idx].metric(f"{cor_situacao(sit)} {sit.title()}", qtd)

    st.bar_chart(pd.Series(contagem, name="Quantidade"))

    st.divider()
    excel_bytes = gerar_excel(resultados)
    nome_arquivo = f"resultado_cnpj_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    st.download_button(
        label="📥 Baixar resultado em Excel",
        data=excel_bytes,
        file_name=nome_arquivo,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
        help="Baixa o resultado da consulta em formato .xlsx",
    )

# ── Abas principais ───────────────────────────────────────────────────────────
aba_excel, aba_lista, aba_frete, aba_historico, aba_privacidade = st.tabs(
    ["📂 Importar Excel", "📋 Colar Lista", "🚚 Calcular Frete", "🕘 Histórico", "🔒 Privacidade"]
)

# ── ABA 1: Excel ─────────────────────────────────────────────────────────────
with aba_excel:
    arquivo = st.file_uploader(
        "Selecione o arquivo Excel com os CNPJs", type=["xlsx"],
        help="Arquivo .xlsx com um CNPJ por linha, a partir da linha 2 da coluna A",
    )

    if arquivo:
        try:
            df_entrada = pd.read_excel(arquivo, header=0, usecols=[0], dtype=str)
            df_entrada.columns = ["CNPJ"]
            raws = df_entrada["CNPJ"].dropna().tolist()
            cnpjs, invalidos = extrair_cnpjs(raws)

            col1, col2, col3 = st.columns(3)
            col1.metric("📋 Total no arquivo", len(raws))
            col2.metric("✅ Válidos", len(cnpjs))
            col3.metric("❌ Inválidos", len(invalidos))

            if invalidos:
                with st.expander(f"⚠️ {len(invalidos)} CNPJs inválidos"):
                    st.caption("Linhas que não têm 14 dígitos numéricos após a limpeza da formatação.")
                    st.write(invalidos)

            if st.button("🚀 Iniciar Consulta", type="primary", use_container_width=True, key="btn_excel"):
                if cnpjs:
                    resultados = executar_consulta(cnpjs)
                    mostrar_resultados(resultados)
                else:
                    st.warning("Nenhum CNPJ válido encontrado no arquivo.")

        except Exception as e:
            st.error(f"Não foi possível ler o arquivo enviado. Verifique se é um .xlsx válido. Detalhe: {e}")
    else:
        st.info("👆 Faça o upload de um arquivo Excel para começar.")

# ── ABA 2: Lista colada ───────────────────────────────────────────────────────
with aba_lista:
    st.markdown("Cole os CNPJs abaixo, **um por linha**. Aceita com ou sem pontuação.")
    texto = st.text_area(
        "Lista de CNPJs para consulta",
        height=200,
        placeholder="12.345.678/0001-99\n98765432000100\n11.222.333/0001-81",
        label_visibility="collapsed",
        help="Um CNPJ por linha, com ou sem pontuação",
    )

    if texto.strip():
        raws = [linha.strip() for linha in texto.strip().splitlines() if linha.strip()]
        cnpjs, invalidos = extrair_cnpjs(raws)

        col1, col2, col3 = st.columns(3)
        col1.metric("📋 Total inseridos", len(raws))
        col2.metric("✅ Válidos", len(cnpjs))
        col3.metric("❌ Inválidos", len(invalidos))

        minutos, segundos = divmod(len(cnpjs) * 21, 60)
        st.info(f"⏱️ Tempo estimado: ~{minutos}min {segundos}s")

        if invalidos:
            with st.expander(f"⚠️ {len(invalidos)} CNPJs inválidos"):
                st.caption("Linhas que não têm 14 dígitos numéricos após a limpeza da formatação.")
                st.write(invalidos)

        if st.button("🚀 Iniciar Consulta", type="primary", use_container_width=True, key="btn_lista"):
            if cnpjs:
                resultados = executar_consulta(cnpjs)
                mostrar_resultados(resultados)
            else:
                st.warning("Nenhum CNPJ válido encontrado.")
    else:
        st.info("👆 Cole os CNPJs acima para começar.")

# ── ABA 3: Cálculo de frete ────────────────────────────────────────────────────
with aba_frete:
    st.markdown(
        "Informe o CNPJ do cliente para buscar o endereço automaticamente e a "
        "quantidade de peças do pedido. O cálculo é **simulado** (não consulta "
        "os Correios), com base em peso e região de destino a partir de São Paulo/SP."
    )

    cnpj_frete_raw = st.text_input(
        "CNPJ do cliente",
        placeholder="12.345.678/0001-99",
        help="CNPJ usado para localizar o endereço de entrega (UF/CEP)",
        key="cnpj_frete",
    )

    dados_cliente = None
    if cnpj_frete_raw.strip():
        cnpjs_validos, cnpjs_invalidos = extrair_cnpjs([cnpj_frete_raw])
        if cnpjs_invalidos:
            st.error("CNPJ inválido: precisa ter 14 dígitos numéricos.")
        elif st.button("🔎 Buscar endereço do cliente", key="btn_buscar_endereco"):
            with st.spinner("Consultando dados do cliente..."):
                dados_cliente = consultar_cnpj(cnpjs_validos[0])
                st.session_state["dados_cliente_frete"] = dados_cliente

    dados_cliente = st.session_state.get("dados_cliente_frete")
    if dados_cliente:
        if dados_cliente["situacao"] in ("Não encontrado", "Erro") or dados_cliente["situacao"].startswith("HTTP") or dados_cliente["situacao"].startswith("Falha"):
            st.error(f"Não foi possível obter o endereço: {dados_cliente['situacao']}")
        else:
            st.success(
                f"**{dados_cliente['nome']}** — {dados_cliente['municipio']}/{dados_cliente['uf']} "
                f"— CEP {dados_cliente.get('cep', 'N/A')}"
            )

            st.divider()
            st.subheader("📦 Quantidade de peças")
            c1, c2, c3 = st.columns(3)
            qtd_camisa = c1.number_input("Camisas", min_value=0, value=0, step=1, help="Peso unitário: 0,20 kg")
            qtd_calca  = c2.number_input("Calças", min_value=0, value=0, step=1, help="Peso unitário: 0,50 kg")
            qtd_blusa  = c3.number_input("Blusas", min_value=0, value=0, step=1, help="Peso unitário: 0,30 kg")
            components.html(_html_frete_ao_vivo(dados_cliente["uf"]), height=70)

            if st.button("📐 Calcular Frete", type="primary", use_container_width=True):
                if qtd_camisa + qtd_calca + qtd_blusa == 0:
                    st.warning("Informe ao menos uma peça para calcular o frete.")
                else:
                    peso = frete.calcular_peso_total(qtd_camisa, qtd_calca, qtd_blusa)
                    resultado_frete = frete.calcular_frete(dados_cliente["uf"], peso)
                    db.salvar_frete(
                        dados_cliente["cnpj_fmt"], dados_cliente["uf"],
                        resultado_frete["peso_kg"], resultado_frete["valor"], resultado_frete["prazo_dias"],
                    )
                    col_a, col_b, col_c = st.columns(3)
                    col_a.metric("⚖️ Peso total", f"{peso:.2f} kg")
                    col_b.metric("💰 Valor do frete", f"R$ {resultado_frete['valor']:.2f}")
                    col_c.metric("📅 Prazo estimado", f"{resultado_frete['prazo_dias']} dias úteis")
                    st.caption(
                        f"Região de destino: {resultado_frete['regiao'].title()}. "
                        "Valores simulados para fins didáticos, não refletem tarifa real dos Correios."
                    )

# ── ABA 4: Histórico ────────────────────────────────────────────────────────────
with aba_historico:
    st.markdown("Histórico de consultas e fretes calculados, salvos localmente neste ambiente.")

    sub_consultas, sub_fretes = st.tabs(["Consultas de CNPJ", "Fretes calculados"])

    with sub_consultas:
        consultas = db.listar_consultas()
        if consultas:
            st.dataframe(pd.DataFrame(consultas), use_container_width=True, hide_index=True)
        else:
            st.info("Nenhuma consulta registrada ainda.")

    with sub_fretes:
        fretes = db.listar_fretes()
        if fretes:
            st.dataframe(pd.DataFrame(fretes), use_container_width=True, hide_index=True)
        else:
            st.info("Nenhum frete calculado ainda.")

    st.divider()
    st.subheader("🗑️ Excluir dados de um CNPJ")
    st.caption("Direito de exclusão (LGPD, art. 18) — remove todo o histórico do CNPJ informado.")
    cnpj_excluir = st.text_input("CNPJ a excluir", placeholder="12.345.678/0001-99", key="cnpj_excluir")
    if st.button("Excluir histórico deste CNPJ", key="btn_excluir"):
        if cnpj_excluir.strip():
            cnpjs_validos, _ = extrair_cnpjs([cnpj_excluir])
            if cnpjs_validos:
                db.excluir_por_cnpj(formatar_cnpj(cnpjs_validos[0]))
                st.success("Histórico removido. Atualize a página para ver o resultado.")
            else:
                st.error("CNPJ inválido.")
        else:
            st.warning("Informe um CNPJ.")

# ── ABA 5: Privacidade / LGPD ────────────────────────────────────────────────────
with aba_privacidade:
    st.subheader("🔒 Privacidade e proteção de dados")
    st.markdown("""
    **Finalidade da coleta:** os dados informados (CNPJ) são usados exclusivamente
    para consultar a situação cadastral na Receita Federal e, na aba de frete,
    para estimar o custo de envio. Base legal: execução de procedimento
    solicitado pelo próprio usuário (LGPD, art. 7º, V).

    **O que é armazenado:** CNPJ, razão social, situação cadastral, município,
    UF e CEP retornados pela consulta, além dos fretes calculados. Não
    coletamos CPF de sócios nem outros dados pessoais eventualmente presentes
    na resposta da API.

    **Compartilhamento com terceiros:** o CNPJ digitado é enviado à
    [BrasilAPI](https://brasilapi.com.br) (serviço público gratuito) para
    obter os dados cadastrais. Nenhum dado é enviado a outros serviços.

    **Retenção:** o histórico é mantido por até 12 meses e depois excluído
    automaticamente. É possível solicitar a exclusão imediata de um CNPJ
    específico na aba **Histórico**.

    **Nota:** CNPJ e razão social são dados de pessoa jurídica, fora do escopo
    principal da LGPD (que trata de pessoa física). Se a API retornar nome de
    sócio/responsável, esse dado passa a ser pessoal e recebe o mesmo
    tratamento de minimização e retenção descrito acima.
    """)

    if st.button("🧹 Executar limpeza de histórico com mais de 12 meses"):
        removidos = db.limpar_historico_antigo(dias=365)
        st.success(f"{removidos} registro(s) com mais de 12 meses removido(s).")
