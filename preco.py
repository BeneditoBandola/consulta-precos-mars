from io import BytesIO
import os
import unicodedata
import pandas as pd
import streamlit as st

# --- 1. CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="Consulta de Preços - Mars", layout="centered", page_icon="🐱🐶"
)

# --- 2. ESTILO VISUAL E CORES ---
st.markdown(
    """
<style>
.stApp { 
    background-color: #0F172A; 
    color: #F8FAFC; 
}
.caixa-produto-info {
    background-color: #1E293B;
    border: 1px solid #334155;
    border-top: 5px solid #10B981;
    border-radius: 16px;
    padding: 24px 20px;
    box-shadow: 0 8px 24px rgba(0, 0, 0, 0.3);
    text-align: center;
    margin-top: 15px;
    margin-bottom: 25px;
}
.caixa-preco-central {
    background: #0F172A;
    border: 1.5px solid #475569;
    border-top: 4px solid #10B981;
    padding: 18px;
    border-radius: 12px;
    margin-top: 18px;
    text-align: center;
}
.titulo-preco {
    color: #94A3B8;
    font-size: 13px;
    font-weight: 800;
    text-transform: uppercase;
    letter-spacing: 1px;
}
.valor-preco {
    color: #34D399;
    font-size: 44px;
    font-weight: 900;
    margin-top: 4px;
    line-height: 1.1;
    font-family: monospace, sans-serif;
}
.badge-familia {
    display: inline-block;
    background-color: #334155;
    color: #F1F5F9;
    padding: 5px 14px;
    border-radius: 14px;
    font-size: 12px;
    font-weight: 700;
    margin-bottom: 10px;
}
div[data-testid="stImage"] {
    display: flex;
    justify-content: center;
    align-items: center;
}
div[data-testid="stImage"] img {
    object-fit: contain !important;
}
input, select, textarea {
    background-color: #1E293B !important;
    color: #FFFFFF !important;
}
</style>
""",
    unsafe_allow_html=True,
)


# --- 3. CARREGAR BASE DO GITHUB ---
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
  # URL RAW da planilha unificada hospedada no seu GitHub
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


# --- 4. BARRA LATERAL: SELETOR DE ESTADO (MG / SP) ---
st.sidebar.title("⚙️ Painel de Controlo")
st.sidebar.markdown("### Selecione o Estado:")
estado_selecionado = st.sidebar.radio(
    "Praça:", ["Minas Gerais (MG)", "São Paulo (SP)"]
)

if "Minas" in estado_selecionado:
  coluna_preco_selecionada = "RECOMENDADO MG"
  sigla_estado = "MG"
  icone_estado = "⛰️"
else:
  coluna_preco_selecionada = "RECOMENDADO SP"
  sigla_estado = "SP"
  icone_estado = "🌆"

# --- 5. INTERFACE DE CONSULTA RÁPIDA DE PREÇOS ---
st.markdown(
    f"<h2 style='text-align: center; color: #F8FAFC; margin-bottom:"
    f" 2px;'>🐱🐶 Consulta Preços - Mars ({sigla_estado})</h2>",
    unsafe_allow_html=True,
)
st.markdown(
    "<p style='text-align: center; font-size: 13px; color: #94A3B8;'>Digite os"
    " <b>dígitos finais do código de barras (EAN)</b>, o <b>código"
    " Minassal</b> ou o <b>nome do produto</b>:</p>",
    unsafe_allow_html=True,
)

codigo_busca = st.text_input(
    "🔍 Buscar Produto:",
    placeholder="Ex: 97283, whiskas, pedigree, lata...",
    label_visibility="collapsed",
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
            '<div style="margin-top: 12px;"><span style="color: #FBBF24;'
            ' font-size: 13px; font-weight: 700;">⚠️ Preço não'
            f" cadastrado para {sigla_estado}</span></div>"
        )

      html_card = (
          f'<div class="caixa-produto-info"><span'
          f' class="badge-familia">{familia_val}</span><h3 style="color:'
          f' #F8FAFC; margin-top: 4px; margin-bottom: 6px; font-size:'
          f' 19px;">{nome_produto}</h3><p style="font-size: 12.5px; color:'
          f' #94A3B8; margin-bottom: 0;">{detalhes_str}</p>{bloco_preco}</div>'
      )
      st.markdown(html_card, unsafe_allow_html=True)
  else:
    st.error(f"❌ Nenhum produto encontrado para: **{codigo_busca}**.")
