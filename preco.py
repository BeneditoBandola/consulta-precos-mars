from io import BytesIO
import os
import unicodedata
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Consulta de Preços - Mars", layout="centered", page_icon="🐱🐶"
)

if "estado_ativo" not in st.session_state:
    st.session_state.estado_ativo = "MG"

if st.session_state.estado_ativo == "MG":
    fundo_app = "#FFFFFF"
    texto_app = "#0F172A"
    sigla_estado = "MG"
    coluna_preco = "RECOMENDADO MG"
    icone_estado = "⛰️"
else:
    fundo_app = "#E0F2FE"
    texto_app = "#0F172A"
    sigla_estado = "SP"
    coluna_preco = "RECOMENDADO SP"
    icone_estado = "🌆"

st.markdown(
    f"""
<style>
.stApp {{ background-color: {fundo_app} !important; color: {texto_app} !important; }}
h1, h2, h3, p, span, label {{ color: {texto_app} !important; }}

div.stButton > button {{
    background-color: {fundo_app} !important;
    color: {texto_app} !important;
    border: 2px solid #CBD5E1 !important;
    font-weight: 600 !important;
    border-radius: 10px !important;
    transition: all 0.2s ease-in-out;
}}

div.stButton > button:hover {{
    border-color: #10B981 !important;
    color: #10B981 !important;
}}

.card-prod {{
    background-color: #FFFFFF;
    border: 1px solid #CBD5E1;
    border-top: 5px solid #10B981;
    border-radius: 16px;
    padding: 24px 20px;
    box-shadow: 0 8px 24px rgba(0,0,0,0.08);
    text-align: center;
    margin-top: 15px;
    margin-bottom: 25px;
}}
input {{ background-color: #FFFFFF !important; color: #0F172A !important; border: 1px solid #CBD5E1 !important; }}
</style>
""",
    unsafe_allow_html=True,
)


def normalizar_texto(texto):
    if pd.isna(texto):
        return ""
    nfkd = unicodedata.normalize("NFKD", str(texto))
    return "".join([c for c in nfkd if not unicodedata.combining(c)]).lower().strip()


def limpar_codigo(val):
    if pd.isna(val):
        return ""
    txt = str(val).strip()
    return txt[:-2] if txt.endswith(".0") else txt


@st.cache_data
def carregar_dados():
    url_github = "https://raw.githubusercontent.com/SEU_USUARIO/SEU_REPO/main/skeletor_com_codigo_minassal_preenchido.xlsx"
    local = "skeletor_com_codigo_minassal_preenchido.xlsx"
    try:
        df = (
            pd.read_excel(local, dtype=str)
            if os.path.exists(local)
            else pd.read_excel(url_github, dtype=str)
        )
    except:
        return None
    df.columns = [str(c).strip().upper() for c in df.columns]
    for c in ["EAN", "CODIGO_MINASSAL", "SKU"]:
        if c in df.columns:
            df[f"{c}_LIMPO"] = df[c].apply(limpar_codigo)
        else:
            df[f"{c}_LIMPO"] = ""
    df["BUSCA_COMPLETA"] = df.apply(
        lambda r: normalizar_texto(
            f"{r.get('PRODUTO', '')} {r.get('SUBBRAND', '')} {r.get('FAMILY PRICE', '')} {r.get('CODIGO_MINASSAL_LIMPO', '')} {r.get('EAN_LIMPO', '')} {r.get('SKU_LIMPO', '')}"
        ),
        axis=1,
    )
    return df


df_produtos = carregar_dados()
PASTA_FOTOS = "mockups_produtos"


def obter_img(cod):
    for ext in [".png", ".jpg", ".jpeg", ".webp"]:
        p = os.path.join(PASTA_FOTOS, f"{str(cod).strip().replace('.0','')}{ext}")
        if os.path.exists(p):
            return p
    return None


def extrair_prc(row, col):
    v = row.get(col, 0.0)
    if pd.isna(v):
        return 0.0
    try:
        if isinstance(v, (int, float)):
            return float(v)
        txt = str(v).replace("R$", "").strip().replace(",", ".")
        return float(txt)
    except:
        return 0.0


def fmt_moeda(val):
    return (
        f"R$ {val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        if val > 0
        else "Não cadastrado"
    )


st.markdown(
    "<h2 style='text-align: center; margin-bottom: 5px;'>🐱🐶 Consulta de"
    " Preços - Mars</h2>",
    unsafe_allow_html=True,
)
st.markdown(
    "<p style='text-align: center; font-size: 13px; margin-bottom:"
    " 15px;'>Selecione o estado abaixo:</p>",
    unsafe_allow_html=True,
)

col1, col2 = st.columns(2)

with col1:
    if st.button("⛰️ Minas Gerais (MG)", use_container_width=True):
        st.session_state.estado_ativo = "MG"
        st.rerun()

with col2:
    if st.button("🌆 São Paulo (SP)", use_container_width=True):
        st.session_state.estado_ativo = "SP"
        st.rerun()

st.markdown(
    f"<p style='text-align: center; font-size: 13px; font-weight: bold;"
    f" margin-top: 10px;'>Praça Ativa: {sigla_estado} ({'Minas Gerais' if sigla_estado == 'MG' else 'São Paulo'})</p>",
    unsafe_allow_html=True,
)
st.markdown("---")

modo = st.radio(
    "Modo de Operação",
    ["🔍 Consulta de Preços", "🎯 Simulador de Campanhas"],
    horizontal=True,
    label_visibility="collapsed",
)

st.markdown("---")

if modo == "🔍 Consulta de Preços":
    busca = st.text_input(
        "🔍 Buscar Produto:",
        placeholder=(
            "Digite os dígitos finais do EAN, código Minassal ou nome do"
            " produto..."
        ),
    )

    if df_produtos is None:
        st.error("⚠️ Planilha não encontrada.")
    elif busca:
        b_limpa = busca.replace(".0", "").strip()
        if b_limpa.isdigit():
            df_m = df_produtos[
                (df_produtos["EAN_LIMPO"].str.endswith(b_limpa))
                | (df_produtos["CODIGO_MINASSAL_LIMPO"].str.endswith(b_limpa))
                | (df_produtos["SKU_LIMPO"].str.endswith(b_limpa))
            ]
        else:
            tokens = [normalizar_texto(t) for t in b_limpa.split() if t.strip()]
            df_m = df_produtos[
                df_produtos["BUSCA_COMPLETA"].apply(
                    lambda x: all(t in x for t in tokens)
                )
            ]

        if not df_m.empty:
            for idx, row in df_m.iterrows():
                prod = row.get("PRODUTO", "")
                ean = row.get("EAN_LIMPO", "")
                cod_min = row.get("CODIGO_MINASSAL_LIMPO", "")
                familia = row.get("SUBBRAND", "Mars")
                preco = extrair_prc(row, coluna_preco)
                img = obter_img(cod_min) or obter_img(ean)

                col_esq, col_centro, col_dir = st.columns([1, 2.5, 1])
                with col_centro:
                    if img:
                        st.image(img, use_container_width=True)

                prc_fmt = fmt_moeda(preco)
                html = f"""
                    <div class="card-prod">
                        <span style="background:#E2E8F0; color:#0F172A; padding:4px 12px; border-radius:10px; font-size:11px; font-weight:bold;">{familia}</span>
                        <h3 style="margin:8px 0; font-size:18px;">{prod}</h3>
                        <p style="font-size:12px; color:#64748B;">Cód Minassal: {cod_min} | EAN: {ean}</p>
                        <div style="background: {fundo_app}; border: 1.5px solid #CBD5E1; border-top: 4px solid #10B981; padding: 18px; border-radius: 12px; margin-top: 18px; text-align: center;">
                            <span style="font-size:11px; font-weight:bold; color:#64748B; text-transform:uppercase;">{icone_estado} RSP Recomendado ({sigla_estado})</span>
                            <div style="font-size:36px; font-weight:900; color:#059669; font-family:monospace;">{prc_fmt}</div>
                        </div>
                    </div>
                    """
                st.markdown(html, unsafe_allow_html=True)
        else:
            st.error("❌ Nenhum produto encontrado.")

else:  # Modo Simulador de Campanhas
    st.markdown("### 🎯 Simulador de Campanhas & Promoções no PDV")
    st.markdown(
        "Configure a mecânica promocional e pesquise o produto para ver os"
        " valores, preço por unidade e sugestões de arredondamento:"
    )

    tipo_campanha = st.selectbox(
        "Mecânica Promocional:",
        [
            "Leve 2, Pague 1",
            "Leve 3, Pague 2",
            "Desconto Percentual Customizado (%)",
        ],
    )

    perc_desc = 0.0
    if tipo_campanha == "Desconto Percentual Customizado (%)":
        col_p1, col_p2 = st.columns([2, 1])
        with col_p1:
            perc_desc = (
                st.slider("Arraste a porcentagem de desconto:", 1, 70, 15, 1)
                / 100.0
            )
        with col_p2:
            perc_desc_manual = (
                st.number_input(
                    "Ou digite %:", min_value=1, max_value=90, value=int(perc_desc * 100)
                )
                / 100.0
            )
            perc_desc = perc_desc_manual

    busca_campanha = st.text_input(
        "🔍 Buscar Produto para Campanha:",
        placeholder="Digite o nome, EAN ou código Minassal...",
    )

    if df_produtos is not None and busca_campanha:
        b_limpa = busca_campanha.replace(".0", "").strip()
        if b_limpa.isdigit():
            df_m = df_produtos[
                (df_produtos["EAN_LIMPO"].str.endswith(b_limpa))
                | (df_produtos["CODIGO_MINASSAL_LIMPO"].str.endswith(b_limpa))
                | (df_produtos["SKU_LIMPO"].str.endswith(b_limpa))
            ]
        else:
            tokens = [normalizar_texto(t) for t in b_limpa.split() if t.strip()]
            df_m = df_produtos[
                df_produtos["BUSCA_COMPLETA"].apply(
                    lambda x: all(t in x for t in tokens)
                )
            ]

        if not df_m.empty:
            for idx, row in df_m.iterrows():
                prod = row.get("PRODUTO", "")
                ean = row.get("EAN_LIMPO", "")
                cod_min = row.get("CODIGO_MINASSAL_LIMPO", "")
                familia = row.get("SUBBRAND", "Mars")
                preco_cheio = extrair_prc(row, coluna_preco)
                img = obter_img(cod_min) or obter_img(ean)

                col_esq, col_centro, col_dir = st.columns([1, 2.5, 1])
                with col_centro:
                    if img:
                        st.image(img, use_container_width=True)

                # Cálculos das Campanhas
                if tipo_campanha == "Leve 2, Pague 1":
                    unidades_leva = 2
                    unidades_paga = 1
                    total_promocao = preco_cheio * unidades_paga
                    preco_unit_efetivo = (
                        total_promocao / unidades_leva
                        if unidades_leva > 0
                        else 0
                    )
                    detalhe_mec = "Leve 2 unidades e pague apenas 1"
                elif tipo_campanha == "Leve 3, Pague 2":
                    unidades_leva = 3
                    unidades_paga = 2
                    total_promocao = preco_cheio * unidades_paga
                    preco_unit_efetivo = (
                        total_promocao / unidades_leva
                        if unidades_leva > 0
                        else 0
                    )
                    detalhe_mec = "Leve 3 unidades e pague apenas 2"
                else:  # Desconto Percentual Customizado
                    total_promocao = preco_cheio * (1 - perc_desc)
                    preco_unit_efetivo = total_promocao
                    detalhe_mec = (
                        f"Desconto de {int(perc_desc * 100)}% aplicado"
                    )

                # Lógica de Arredondamento Psicológico (,99 ou ,90)
                val_int = int(total_promocao)
                sugestao_99 = val_int - 0.01 if val_int > 0 else 0.99
                if sugestao_99 <= 0 or sugestao_99 > total_promocao:
                    sugestao_99 = (
                        float(val_int - 1) + 0.99 if val_int > 1 else 0.99
                    )

                # Renderização limpa usando containers nativos do Streamlit
                with st.container():
                    st.markdown(
                        f"""
                        <div class="card-prod">
                            <span style="background:#E2E8F0; color:#0F172A; padding:4px 12px; border-radius:10px; font-size:11px; font-weight:bold;">{familia} - SIMULAÇÃO PDV</span>
                            <h3 style="margin:8px 0; font-size:18px;">{prod}</h3>
                            <p style="font-size:12px; color:#64748B;">Cód Minassal: {cod_min} | EAN: {ean}</p>
                            <hr style="border:0; border-top:1px solid #CBD5E1; margin: 12px 0;">
                            <p style="font-size:13px; font-weight:bold; color:#059669; margin-bottom:5px;">Mecânica: {tipo_campanha} ({detalhe_mec})</p>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )

                    c_1, c_2 = st.columns(2)
                    with c_1:
                        st.metric(
                            label="PREÇO CHEIO (RSP)",
                            value=fmt_moeda(preco_cheio),
                        )
                    with c_2:
                        st.metric(
                            label="VALOR PROMOCIONAL EXATO",
                            value=fmt_moeda(total_promocao),
                        )

                    c_3, c_4 = st.columns(2)
                    with c_3:
                        st.info(
                            f"**Valor por Unidade Efetiva:**\n\n{fmt_moeda(preco_unit_efetivo)}"
                        )
                    with c_4:
                        st.success(
                            f"**💡 Sugestão Preço Psicológico:**\n\n{fmt_moeda(sugestao_99)}"
                        )
        else:
            st.error("❌ Nenhum produto encontrado para a campanha.")
