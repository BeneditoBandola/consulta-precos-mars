from io import BytesIO
import os
import unicodedata
import pandas as pd
import streamlit as st

# --- 1. CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="Consulta de Preços - Mars", layout="centered", page_icon="🐱🐶"
)

# --- 2. GESTÃO DE ESTADO (MEMÓRIA DO SELETOR) ---
if "estado_ativo" not in st.session_state:
  st.session_state.estado_ativo = "MG"

# Definir as cores dinâmicas com base no estado ativo
if st.session_state.estado_ativo == "MG":
  cor_fundo_app = "#FFFFFF"  # Branco para MG
  cor_texto_app = "#0F172A"  # Texto escuro para contraste no fundo branco
  cor_card_bg = "#F8FAFC"
  cor_card_borda = "#E2E8F0"
  coluna_preco_selecionada = "RECOMENDADO MG"
  sigla_estado = "MG"
  icone_estado = "⛰️"
else:
  cor_fundo_app = "#E0F2FE"  # Azul clarinho para SP
  cor_texto_app = "#0F172A"  # Texto escuro para contraste
  cor_card_bg = "#BAE6FD"
  cor_card_borda = "#7DD3FC"
  coluna_preco_selecionada = "RECOMENDADO SP"
  sigla_estado = "SP"
  icone_estado = "🌆"

# --- 3. ESTILO VISUAL DINÂMICO ---
st.markdown(
    f"""
<style>
.stApp {{ 
    background-color: {cor_fundo_app} !important; 
    color: {cor_texto_app} !important; 
}}
h1, h2, h3, p, span, label {{ 
    color: {cor_texto_app} !important; 
}}
.caixa-produto-info {{
    background-color: {cor_card_bg};
    border: 1px solid {cor_card_borda};
    border-top: 5px solid #10B981;
    border-radius: 16px;
    padding: 24px 20px;
    box-shadow: 0 8px 24px rgba(0, 0, 0, 0.08);
    text-align: center;
    margin-top: 15px;
    margin-bottom: 25px;
}}
.caixa-preco-central {{
    background: {cor_fundo_app};
    border: 1.5px solid {cor_card_borda};
    border-top: 4px solid #10B981;
    padding: 18px;
    border-radius: 12px;
    margin-top: 18px;
    text-align: center;
}}
.titulo-preco {{
    color: #64748B;
    font-size: 13px;
    font-weight: 800;
    text-transform: uppercase;
    letter-spacing: 1px;
}}
.valor-preco {{
    color: #059669;
    font-size: 44px;
    font-weight: 900;
    margin-top: 4px;
    line-height: 1.1;
    font-family: monospace, sans-serif;
}
.badge-familia {{
    display: inline-block;
    background-color: #E2E8F0;
    color: #0F172A;
    padding: 5px 14px;
    border-radius: 14px;
    font-size: 12px;
    font-weight: 700;
    margin-bottom: 10px;
}}
div[data-testid="stImage"] {{
    display: flex;
    justify-content: center;
    align-items: center;
}}
div[data-testid="stImage"] img {{
    object-fit: contain !important;
}}
input, select, textarea {{
    background-color: #FFFFFF !important;
    color: #0F172A !important;
    border: 1px solid #CBD5E1 !important;
}}
</style>
""",
    unsafe_allow_html=True,
)


# --- 4. CARREGAR BASE DO GITHUB ---
def normalizar_texto(texto):
  if pd.isna(texto):
    return ""
  nfkd = unicodedata.normalize("NFKD", str(texto))
  return "".join([c for c in nfkd if not unicodedata.combining(c)]).lower().strip()


def limpar_campo_codigo(val):
  if pd.isna(val):
    return ""
  txt = str(val).strip()
  if txt.endswith(".0"):
    txt = txt[:-2]
  return txt


@st.cache_data
def carregar_dados():
  url_github = "https://raw.githubusercontent.com/SEU_USUARIO/SEU_REPO/main/skeletor_com_codigo_minassal_preenchido.xlsx"
  arquivo_local = "skeletor_com_codigo_minassal_preenchido.xlsx"

  try:
    if os.path.exists(arquivo_local):
      df = pd.read_excel(arquivo_local, dtype=str)
    else:
      df = pd.read_excel(url_github, dtype=str)
  except Exception as e:
    return None

  df.columns = [str(c).strip().upper() for c in df.columns]

  for col_alvo in ["EAN", "CODIGO_MINASSAL", "SKU"]:
    if col_alvo in df.columns:
      df[f"{col_alvo}_LIMPO"] = df[col_alvo].apply(limpar_campo_codigo)
    else:
      df[f"{col_alvo}_LIMPO"] = ""

  df["BUSCA_COMPLETA"] = df.apply(
      lambda r: normalizar_texto(
          f"{r.get('PRODUTO', '')} {r.get('SUBBRAND', '')} {r.get('FAMILY PRICE', '')} "
          f"{r.get('CODIGO_MINASSAL_LIMPO', '')} {r.get('EAN_LIMPO', '')} {r.get('SKU_LIMPO', '')}"
      ),
      axis=1,
  )
  return df


df_produtos = carregar_dados()
PASTA_FOTOS = "mockups_produtos"


def obter_caminho_imagem(codigo_identificador):
  extensoes = [".png", ".jpg", ".jpeg", ".webp", ".PNG", ".JPG", ".JPEG"]
  cod_limpo = str(codigo_identificador).strip().replace(".0", "")
  if os.path.exists(PASTA_FOTOS) and cod_limpo:
    for ext in extensoes:
      caminho_completo = os.path.join(PASTA_FOTOS, f"{cod_limpo}{ext}")
      if os.path.exists(caminho_completo):
        return caminho_completo
  return None


def extrair_preco(row, coluna_alvo):
  preco_raw = row.get(coluna_alvo, 0.0)
  if pd.isna(preco_raw):
    return 0.0
  try:
    if isinstance(preco_raw, (int, float)):
      return float(preco_raw)

    txt = str(preco_raw).strip()
    if not txt or txt.lower() == "nan":
      return 0.0

    txt = txt.replace("R$", "").strip()

    if "," in txt and "." in txt:
      if txt.rfind(",") > txt.rfind("."):
        txt = txt.replace(".", "").replace(",", ".")
      else:
        txt = txt.replace(",", "")
    elif "," in txt:
      parts = txt.split(",")
      if len(parts) == 2 and len(parts[1]) <= 2:
        txt = txt.replace(",", ".")
      else:
        txt = txt.replace(",", "")

    return float(txt)
  except:
    return 0.0


# --- 5. INTERFACE PRINCIPAL ---
st.markdown(
    "<h2 style='text-align: center; margin-bottom: 5px;'>🐱🐶 Consulta de"
    " Preços - Mars</h2>",
    unsafe_allow_html=True,
)
st.markdown(
    "<p style='text-align: center; font-size: 13px; margin-bottom:"
    " 20px;'>Selecione o estado e pesquise pelo produto:</p>",
    unsafe_allow_html=True,
)

# Botões de Seleção de Estado Centralizados
col_btn1, col_btn2 = st.columns(2)
with col_btn1:
  if st.button("⛰️ Minas Gerais (MG)", use_container_width=True):
    st.session_state.estado_ativo = "MG"
    st.rerun()
with col_btn2:
  if st.button("🌆 São Paulo (SP)", use_container_width=True):
    st.session_state.estado_ativo = "SP"
    st.rerun()

st.markdown(
    f"<p style='text-align: center; font-size: 13px; font-weight: bold;"
    f" margin-top: 10px;'>Praça Ativa: {sigla_estado} ({'Minas Gerais' if sigla_estado == 'MG' else 'São Paulo'})</p>",
    unsafe_allow_html=True,
)
st.markdown("---")

# Campo de Busca
codigo_busca = st.text_input(
    "🔍 Buscar Produto:",
    placeholder=(
        "Digite os dígitos finais do EAN, código Minassal ou nome do produto..."
    ),
)

if df_produtos is None:
  st.error(
      "⚠️ A planilha unificada não foi encontrada no GitHub ou localmente."
  )
elif codigo_busca:
  busca_raw = str(codigo_busca).strip()
  busca_limpa = busca_raw.replace(".0", "").strip()

  if busca_limpa.isdigit():
    df_match = df_produtos[
        (df_produtos["EAN_LIMPO"].str.endswith(busca_limpa))
        | (df_produtos["CODIGO_MINASSAL_LIMPO"].str.endswith(busca_limpa))
        | (df_produtos["SKU_LIMPO"].str.endswith(busca_limpa))
        | (df_produtos["EAN_LIMPO"] == busca_limpa)
        | (df_produtos["CODIGO_MINASSAL_LIMPO"] == busca_limpa)
        | (df_produtos["SKU_LIMPO"] == busca_limpa)
    ]
  else:
    tokens = [normalizar_texto(t) for t in busca_raw.split() if t.strip()]

    def match_tokens(texto_registro):
      for tok in tokens:
        if tok not in texto_registro:
          return False
      return True

    df_match = df_produtos[df_produtos["BUSCA_COMPLETA"].apply(match_tokens)]

  if not df_match.empty:
    if len(df_match) > 1:
      st.info(f"ℹ️ Encontrados **{len(df_match)} produtos** correspondentes:")

    for index, row in df_match.iterrows():
      nome_produto = row.get("PRODUTO", "Produto Mars")
      ean_val = row.get("EAN_LIMPO", "N/D")
      cod_minassal = row.get("CODIGO_MINASSAL_LIMPO", "N/D")
      familia_val = str(row.get("SUBBRAND", row.get("CATEGORIA", "Mars")))
      preco_recomendado = extrair_preco(row, coluna_preco_selecionada)

      sku_val = row.get("SKU_LIMPO", "")
      caminho_img = (
          obter_caminho_imagem(cod_minassal)
          or obter_caminho_imagem(ean_val)
          or obter_caminho_imagem(sku_val)
      )

      col_esq, col_centro, col_dir = st.columns([1, 2.8, 1])
      with col_centro:
        if caminho_img and os.path.exists(caminho_img):
          st.image(caminho_img, use_container_width=True)
        else:
          st.markdown(
              "<p style='text-align: center; color: #64748B; font-size: 13px;"
              " margin: 20px 0;'>🖼️ Imagem não disponível</p>",
              unsafe_allow_html=True,
          )

      detalhes_str = f"<b>Cód Minassal:</b> {cod_minassal}"
      if ean_val and ean_val != "N/D":
        detalhes_str += f" | <b>EAN:</b> {ean_val}"

      if preco_recomendado > 0:
        preco_formatado = (
            f"R$ {preco_recomendado:,.2f}"
            .replace(",", "X")
            .replace(".", ",")
            .replace("X", ".")
        )
        bloco_preco = f"""
                <div class="caixa-preco-central">
                    <div class="titulo-preco">{icone_estado} RSP Recomendado ({sigla_estado})</div>
                    <div class="valor-preco">{preco_formatado}</div>
                </div>
                """
      else:
        bloco_preco = (
            '<div style="margin-top: 12px;"><span style="color: #D97706;'
            ' font-size: 13px; font-weight: 700;">⚠️ Preço não'
            f" cadastrado para {sigla_estado}</span></div>"
        )

      html_card = (
          f'<div class="caixa-produto-info"><span'
          f' class="badge-familia">{familia_val}</span><h3 style="margin-top:'
          f' 4px; margin-bottom: 6px; font-size:'
          f' 19px;">{nome_produto}</h3><p style="font-size: 12.5px;'
          f' margin-bottom: 0;">{detalhes_str}</p>{bloco_preco}</div>'
      )
      st.markdown(html_card, unsafe_allow_html=True)
  else:
    st.error(f"❌ Nenhum produto encontrado para: **{codigo_busca}**.")
