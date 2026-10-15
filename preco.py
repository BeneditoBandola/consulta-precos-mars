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

# Configuração de Cores - Tema Clean
if st.session_state.estado_ativo == "MG":
  fundo_app = "#FFFFFF"
  texto_app = "#1E293B"
  fundo_botao = "#F1F5F9"
  texto_botao = "#0F172A"
  borda_botao = "#CBD5E1"
  sigla_estado = "MG"
  icone_estado = "⛰️"
  coluna_preco = "RECOMENDADO MG"
else:
  fundo_app = "#F0F9FF"
  texto_app = "#0C4A6E"
  fundo_botao = "#BAE6FD"
  texto_botao = "#0369A1"
  borda_botao = "#7DD3FC"
  sigla_estado = "SP"
  icone_estado = "🌆"
  coluna_preco = "RECOMENDADO SP"

st.markdown(
    f"""
<style>
.stApp {{ background-color: {fundo_app} !important; color: {texto_app} !important; }}
h1, h2, h3, p, span, label {{ color: {texto_app} !important; }}
.card-prod {{
    background-color: #FFFFFF;
    border: 1px solid #E2E8F0;
    border-top: 4px solid #10B981;
    border-radius: 12px;
    padding: 20px;
    box-shadow: 0 4px 12px rgba(0,0,0,0.05);
    text-align: center;
    margin-top: 15px;
}}
.card-preco {{
    background: #F8FAFC;
    border: 1px solid #CBD5E1;
    border-radius: 8px;
    padding: 15px;
    margin-top: 15px;
    text-align: center;
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


st.markdown(
    "<h2 style='text-align: center; margin-bottom: 5px;'>🐱🐶 Consulta de"
    " Preços - Mars</h2>",
    unsafe_allow_html=True,
)
st.markdown(
    "<p style='text-align: center; font-size: 13px; margin-bottom:"
    " 20px;'>Selecione o estado abaixo:</p>",
    unsafe_allow_html=True,
)

c1, c2 = st.columns(2)
with c1:
  if st.button(
      "⛰️ Minas Gerais (MG)", use_container_width=True, type="secondary"
  ):
    st.session_state.estado_ativo = "MG"
    st.rerun()
with c2:
  if st.button("🌆 São Paulo (SP)", use_container_width=True, type="secondary"):
    st.session_state.estado_ativo = "SP"
    st.rerun()

st.markdown(
    f"<p style='text-align: center; font-size: 14px; font-weight: bold;'>"
    f"Praça Ativa: <b>{sigla_estado}</b></p>",
    unsafe_allow_html=True,
)
st.markdown("---")

busca = st.text_input(
    "🔍 Buscar Produto:",
    placeholder="Digite EAN, código Minassal ou nome...",
    label_visibility="collapsed",
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

      prc_fmt = (
          f"R$ {preco:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
          if preco > 0
          else "Não cadastrado"
      )
      html = f"""
            <div class="card-prod">
                <span style="background:#E2E8F0; color:#0F172A; padding:4px 12px; border-radius:10px; font-size:11px; font-weight:bold;">{familia}</span>
                <h3 style="margin:8px 0; font-size:18px;">{prod}</h3>
                <p style="font-size:12px; color:#64748B;">Cód Minassal: {cod_min} | EAN: {ean}</p>
                <div class="card-preco">
                    <span style="font-size:11px; font-weight:bold; color:#64748B; text-transform:uppercase;">RSP Recomendado ({sigla_estado})</span>
                    <div style="font-size:36px; font-weight:900; color:#059669; font-family:monospace;">{prc_fmt}</div>
                </div>
            </div>
            """
      st.markdown(html, unsafe_allow_html=True)
  else:
     st.error("❌ Nenhum produto encontrado.")
