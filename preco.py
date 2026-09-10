import streamlit as st
import pandas as pd
import unicodedata
import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import re
from io import BytesIO

# Importações do ReportLab para geração de PDF
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# --- 1. CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="Gestão de Campo - Mars",
    layout="centered",
    page_icon="🐱🐶"
)

# --- 2. ESTILO VISUAL E CORES (TEMA ELEGANTE COM VERDE) ---
st.markdown("""
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
""", unsafe_allow_html=True)

# --- 3. CARREGAR BASES DE DADOS ---
def normalizar_texto(texto):
    if pd.isna(texto):
        return ""
    nfkd = unicodedata.normalize('NFKD', str(texto))
    return "".join([c for c in nfkd if not unicodedata.combining(c)]).lower().strip()

def limpar_campo_codigo(val):
    if pd.isna(val):
        return ""
    txt = str(val).strip()
    if txt.endswith('.0'):
        txt = txt[:-2]
    return txt

@st.cache_data
def carregar_dados():
    arquivo_base = "skeletor_com_codigo_minassal.xlsx"
    if not os.path.exists(arquivo_base):
        return None
    try:
        df = pd.read_excel(arquivo_base)
    except Exception as e:
        return None

    df.columns = [str(c).strip().upper() for c in df.columns]
    
    for col_alvo in ['EAN', 'CODIGO_MINASSAL', 'SKU']:
        if col_alvo in df.columns:
            df[f'{col_alvo}_LIMPO'] = df[col_alvo].apply(limpar_campo_codigo)
        else:
            df[f'{col_alvo}_LIMPO'] = ""

    df['BUSCA_COMPLETA'] = df.apply(
        lambda r: normalizar_texto(
            f"{r.get('PRODUTO', '')} {r.get('SUBBRAND', '')} {r.get('FAMILY PRICE', '')} "
            f"{r.get('CODIGO_MINASSAL_LIMPO', '')} {r.get('EAN_LIMPO', '')} {r.get('SKU_LIMPO', '')}"
        ),
        axis=1
    )
    return df

@st.cache_data
def carregar_clientes_pocos():
    arquivo_clientes = "clientes com coordenadas.xlsx"
    if not os.path.exists(arquivo_clientes):
        return pd.DataFrame()
    try:
        xls = pd.ExcelFile(arquivo_clientes)
        df_cli = pd.read_excel(arquivo_clientes, sheet_name=xls.sheet_names[0])
        df_cli.columns = [str(c).strip().upper() for c in df_cli.columns]
        
        if 'CIDADE' in df_cli.columns:
            pocos = df_cli[df_cli['CIDADE'].str.contains('POCOS|POÇOS', case=False, na=False)].copy()
            pocos['NOME'] = pocos['NOME'].astype(str).str.strip()
            pocos['ENDEREÇO'] = pocos['ENDEREÇO'].fillna('').astype(str).str.strip()
            pocos['BAIRRO'] = pocos['BAIRRO'].fillna('').astype(str).str.strip()
            return pocos
    except Exception:
        pass
    return pd.DataFrame()

df_produtos = carregar_dados()
df_clientes_pocos = carregar_clientes_pocos()

PASTA_FOTOS = "mockups_produtos"

def obter_caminho_imagem(codigo_identificador):
    extensoes = ['.png', '.jpg', '.jpeg', '.webp', '.PNG', '.JPG', '.JPEG']
    cod_limpo = str(codigo_identificador).strip().replace('.0', '')
    if os.path.exists(PASTA_FOTOS) and cod_limpo:
        for ext in extensoes:
            caminho_completo = os.path.join(PASTA_FOTOS, f"{cod_limpo}{ext}")
            if os.path.exists(caminho_completo):
                return caminho_completo
    return None

def extrair_preco_mg(row):
    preco_raw = 0.0
    for col_preco in ['RSP \nRECOMENDADO', 'RSP \nRecomendado', 'RSP RECOMENDADO', 'RSP']:
        if col_preco in row and pd.notna(row[col_preco]):
            preco_raw = row[col_preco]
            break
    try:
        if isinstance(preco_raw, str):
            preco_raw = preco_raw.replace('R$', '').replace('.', '').replace(',', '.').strip()
        return float(preco_raw)
    except:
        return 0.0

def eh_produto_inovacao_ou_smallbag(row):
    nome = normalizar_texto(row.get('PRODUTO', ''))
    ean = str(row.get('EAN_LIMPO', ''))
    
    # 1. Regra de exclusão solicitada: Remover todos os Optimum e todos de 500g
    if 'optimum' in nome:
        return False
    if '500g' in nome:
        return False

    # 2. EANs de Inovação diretos permitidos
    eans_alvo = [
        "7896029047606", "7896029047620", "7896029047736", "7896029047651",
        "7896029047743", "7896029047842", "7896029047866", "7896029047880",
        "7896029047965", "7896029047941", "7896029047996", "7896029048078",
        "7896029048085"
    ]
    if ean in eans_alvo:
        return True
        
    termos_chave = ['filezito', 'sheba creamy', 'banana e maca']
    for termo in termos_chave:
        if termo in nome:
            return True
            
    # Small Bags: embalagens menores que 3kg (< 3000g ou < 3kg)
    if 'kg' in nome or 'g' in nome:
        if 'dry' in nome or 'racao' in nome or 'bag' in nome or 'adulto' in nome or 'filhote' in nome:
            match_g = re.search(r'(\d+)\s*g', nome)
            if match_g and int(match_g.group(1)) < 3000:
                return True
            match_kg = re.search(r'(\d+[\.,]?\d*)\s*kg', nome)
            if match_kg:
                val_kg = float(match_kg.group(1).replace(',', '.'))
                if val_kg < 3.0:
                    return True

    return False

# Função para gerar o PDF em memória (BytesIO)
def gerar_pdf_relatorio(razao_social, endereco_cliente, bairro_cliente, produtos_presentes, oportunidades_faltantes):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    story = []
    
    styles = getSampleStyleSheet()
    
    # Estilos customizados
    titulo_style = ParagraphStyle(
        'TituloRelatorio',
        parent=styles['Heading1'],
        fontSize=18,
        textColor=colors.HexColor('#0F172A'),
        spaceAfter=6,
        alignment=1
    )
    sub_style = ParagraphStyle(
        'SubTitulo',
        parent=styles['Normal'],
        fontSize=10,
        textColor=colors.HexColor('#64748B'),
        spaceAfter=15,
        alignment=1
    )
    secao_style = ParagraphStyle(
        'SecaoTitulo',
        parent=styles['Heading2'],
        fontSize=13,
        textColor=colors.HexColor('#10B981'),
        spaceBefore=12,
        spaceAfter=6
    )
    texto_style = ParagraphStyle(
        'TextoNormal',
        parent=styles['Normal'],
        fontSize=10,
        textColor=colors.HexColor('#334155')
    )
    
    # Cabeçalho
    story.append(Paragraph("RELATÓRIO DE VERIFICAÇÃO DE PDV", titulo_style))
    story.append(Paragraph("Minassal / Mars — Poços de Caldas (MG)", sub_style))
    story.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor('#CBD5E1'), spaceAfter=15))
    
    # Dados da Loja
    info_loja = f"""
    <b>Cliente / Razão Social:</b> {razao_social}<br/>
    <b>Endereço:</b> {endereco_cliente} — Bairro: {bairro_cliente}<br/>
    <b>Promotora Responsável:</b> Pamela<br/>
    <b>Localidade:</b> Poços de Caldas - MG
    """
    story.append(Paragraph(info_loja, texto_style))
    story.append(Spacer(1, 15))
    
    # Tabela 1: Produtos Encontrados e Preços
    story.append(Paragraph("<b>✅ Produtos Encontrados e Crítica de Preços (vs RSP MG)</b>", secao_style))
    
    tabela_dados = [["Produto", "Linha", "Cód", "Rec. MG", "Lido", "Análise"]]
    
    for item in produtos_presentes:
        analise = "No Preço"
        if item['preco_praticado'] > 0:
            diff = item['preco_praticado'] - item['preco_recomendado']
            if diff > 0.50:
                analise = f"Acima (+R$ {diff:.2f})"
            elif diff < -0.50:
                analise = f"Abaixo (-R$ {abs(diff):.2f})"
        
        tabela_dados.append([
            item['produto'][:32],
            str(item['categoria']),
            str(item['codigo']),
            f"R$ {item['preco_recomendado']:.2f}",
            f"R$ {item['preco_praticado']:.2f}",
            analise
        ])
        
    t1 = Table(tabela_dados, colWidths=[160, 80, 50, 70, 70, 120])
    t1.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#F1F5F9')),
        ('TEXTCOLOR', (0,0), (-1,0), colors.HexColor('#0F172A')),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,-1), 8.5),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('ALIGN', (0,1), (0,-1), 'LEFT'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
        ('TOPPADDING', (0,0), (-1,-1), 5),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E2E8F0')),
    ]))
    story.append(t1)
    story.append(Spacer(1, 15))
    
    # Tabela 2: Oportunidades / Faltantes
    story.append(Paragraph("<b>🚨 Oportunidades de Inovações & Small Bags Ausentes</b>", secao_style))
    
    tabela_faltantes = [["Produto Oportunidade", "Linha", "Cód", "Sugestão RSP (MG)"]]
    
    if oportunidades_faltantes:
        for item in oportunidades_faltantes:
            tabela_faltantes.append([
                item['produto'][:40],
                str(item['categoria']),
                str(item['codigo']),
                f"R$ {item['preco_recomendado']:.2f}"
            ])
    else:
        tabela_faltantes.append(["Nenhuma oportunidade em falta! Mix estratégico 100% executado.", "", "", ""])
        
    t2 = Table(tabela_faltantes, colWidths=[240, 100, 70, 140])
    t2.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#FEF2F2')),
        ('TEXTCOLOR', (0,0), (-1,0), colors.HexColor('#991B1B')),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,-1), 8.5),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('ALIGN', (0,1), (0,-1), 'LEFT'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
        ('TOPPADDING', (0,0), (-1,-1), 5),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#FCA5A5')),
    ]))
    story.append(t2)
    
    doc.build(story)
    buffer.seek(0)
    return buffer

# --- 4. MENU DE NAVEGAÇÃO ENTRE ABAS ---
aba_selecionada = st.radio(
    "Escolha o modo:",
    ["🔍 Consulta Rápida de Preços", "🏪 VERIFICAÇÃO CLIENTE"],
    horizontal=True,
    label_visibility="collapsed"
)

st.markdown("<hr style='border: 0.5px solid #334155; margin-top: 5px; margin-bottom: 20px;'>", unsafe_allow_html=True)

# ==========================================
# ABA 1: CONSULTA RÁPIDA DE PREÇOS
# ==========================================
if aba_selecionada == "🔍 Consulta Rápida de Preços":
    st.markdown("<h2 style='text-align: center; color: #F8FAFC; margin-bottom: 2px;'>🐱🐶 Consulta Preços - Mars</h2>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; font-size: 13px; color: #94A3B8;'>Digite os <b>dígitos finais do código de barras (EAN)</b>, o <b>código Minassal</b> ou o <b>nome do produto</b>:</p>", unsafe_allow_html=True)

    codigo_busca = st.text_input("🔍 Buscar Produto:", placeholder="Ex: 97283, whiskas, pedigree, lata...", label_visibility="collapsed")

    if df_produtos is None:
        st.error("⚠️ Planilha `skeletor_com_codigo_minassal.xlsx` não encontrada no repositório do GitHub.")
    elif codigo_busca:
        busca_raw = str(codigo_busca).strip()
        busca_limpa = busca_raw.replace('.0', '').strip()
        
        if busca_limpa.isdigit():
            df_match = df_produtos[
                (df_produtos['EAN_LIMPO'].str.endswith(busca_limpa)) |
                (df_produtos['CODIGO_MINASSAL_LIMPO'].str.endswith(busca_limpa)) |
                (df_produtos['SKU_LIMPO'].str.endswith(busca_limpa)) |
                (df_produtos['EAN_LIMPO'] == busca_limpa) |
                (df_produtos['CODIGO_MINASSAL_LIMPO'] == busca_limpa) |
                (df_produtos['SKU_LIMPO'] == busca_limpa)
            ]
        else:
            tokens = [normalizar_texto(t) for t in busca_raw.split() if t.strip()]
            def match_tokens(texto_registro):
                for tok in tokens:
                    if tok not in texto_registro:
                        return False
                return True
            df_match = df_produtos[df_produtos['BUSCA_COMPLETA'].apply(match_tokens)]

        if not df_match.empty:
            if len(df_match) > 1:
                st.info(f"ℹ️ Encontrados **{len(df_match)} produtos** correspondentes:")
                
            for index, row in df_match.iterrows():
                nome_produto = row.get('PRODUTO', 'Produto Mars')
                ean_val = row.get('EAN_LIMPO', 'N/D')
                cod_minassal = row.get('CODIGO_MINASSAL_LIMPO', 'N/D')
                familia_val = str(row.get('SUBBRAND', row.get('CATEGORIA', 'Mars')))
                preco_recomendado = extrair_preco_mg(row)
                
                sku_val = row.get('SKU_LIMPO', '')
                caminho_img = obter_caminho_imagem(cod_minassal) or obter_caminho_imagem(ean_val) or obter_caminho_imagem(sku_val)

                col_esq, col_centro, col_dir = st.columns([1, 2.8, 1])
                with col_centro:
                    if caminho_img and os.path.exists(caminho_img):
                        st.image(caminho_img, use_container_width=True)
                    else:
                        st.markdown("<p style='text-align: center; color: #64748B; font-size: 13px; margin: 20px 0;'>🖼️ Imagem não disponível</p>", unsafe_allow_html=True)
                
                detalhes_str = f"<b>Cód Minassal:</b> {cod_minassal}"
                if ean_val and ean_val != "N/D":
                    detalhes_str += f" | <b>EAN:</b> {ean_val}"
                    
                if preco_recomendado > 0:
                    preco_formatado = f"R$ {preco_recomendado:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
                    bloco_preco = f'<div class="caixa-preco-central"><div class="titulo-preco">💰 RSP Recomendado (MG)</div><div class="valor-preco">{preco_formatado}</div></div>'
                else:
                    bloco_preco = '<div style="margin-top: 12px;"><span style="color: #FBBF24; font-size: 13px; font-weight: 700;">⚠️ Preço não cadastrado</span></div>'

                html_card = f'<div class="caixa-produto-info"><span class="badge-familia">{familia_val}</span><h3 style="color: #F8FAFC; margin-top: 4px; margin-bottom: 6px; font-size: 19px;">{nome_produto}</h3><p style="font-size: 12.5px; color: #94A3B8; margin-bottom: 0;">{detalhes_str}</p>{bloco_preco}</div>'
                st.markdown(html_card, unsafe_allow_html=True)
        else:
            st.error(f"❌ Nenhum produto encontrado para: **{codigo_busca}**.")

# ==========================================
# ABA 2: VERIFICAÇÃO CLIENTE (POÇOS DE CALDAS)
# ==========================================
elif aba_selecionada == "🏪 VERIFICAÇÃO CLIENTE":
    st.markdown("<h2 style='text-align: center; color: #F8FAFC; margin-bottom: 5px;'>🏪 Verificação de Cliente - Poços de Caldas</h2>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; font-size: 13px; color: #94A3B8;'>Busque o cliente pelo nome, adicione os produtos encontrados e registre os preços praticados.</p>", unsafe_allow_html=True)

    if df_produtos is None:
        st.error("⚠️ Planilha de produtos não encontrada.")
    elif df_clientes_pocos.empty:
        st.error("⚠️ Planilha de clientes de Poços de Caldas não encontrada ou vazia.")
    else:
        if 'itens_verificacao' not in st.session_state:
            st.session_state.itens_verificacao = []

        termo_cliente = st.text_input("🔍 Digite o nome do cliente para buscar:", placeholder="Ex: Da Roça, Itamaraty, PetShop...")
        
        razao_social = None
        endereco_cliente = ""
        bairro_cliente = ""

        if termo_cliente:
            df_cli_match = df_clientes_pocos[df_clientes_pocos['NOME'].str.contains(termo_cliente, case=False, na=False)]
            if not df_cli_match.empty:
                nomes_encontrados = df_cli_match['NOME'].tolist()
                razao_social = st.selectbox("Selecione o Cliente Encontrado:", nomes_encontrados)
                
                row_cli = df_cli_match[df_cli_match['NOME'] == razao_social].iloc[0]
                endereco_cliente = str(row_cli.get('ENDEREÇO', ''))
                bairro_cliente = str(row_cli.get('BAIRRO', ''))
            else:
                st.warning("Nenhum cliente encontrado com esse nome em Poços de Caldas.")

        if razao_social:
            st.markdown(f"""
            <div style="background-color: #1E293B; border-left: 4px solid #34D399; padding: 10px 15px; border-radius: 8px; margin-bottom: 15px;">
                <span style="font-size: 12px; color: #94A3B8; text-transform: uppercase; font-weight: 700;">Endereço da Loja:</span><br>
                <span style="font-size: 14px; color: #F8FAFC; font-weight: 600;">📍 {endereco_cliente} - Bairro: {bairro_cliente}</span>
            </div>
            """, unsafe_allow_html=True)

        promotor_nome = st.text_input("Promotor Responsável:", value="Pamela", disabled=True)

        st.markdown("---")
        st.subheader("🔍 Adicionar Produtos Encontrados na Loja")
        
        termo_adicao = st.text_input("Digite o nome ou código do produto para adicionar:", placeholder="Ex: Whiskas, Pedigree, 97283...")

        if termo_adicao:
            t_clean = termo_adicao.replace('.0', '').strip()
            if t_clean.isdigit():
                df_busca_add = df_produtos[
                    (df_produtos['EAN_LIMPO'].str.endswith(t_clean)) |
                    (df_produtos['CODIGO_MINASSAL_LIMPO'].str.endswith(t_clean)) |
                    (df_produtos['SKU_LIMPO'].str.endswith(t_clean))
                ]
            else:
                toks = [normalizar_texto(t) for t in t_clean.split() if t.strip()]
                df_busca_add = df_produtos[df_produtos['BUSCA_COMPLETA'].apply(lambda txt: all(tk in txt for tk in toks))]

            if not df_busca_add.empty:
                for idx_prod, r_prod in df_busca_add.iterrows():
                    p_nome = r_prod.get('PRODUTO', 'Produto')
                    p_cod = r_prod.get('CODIGO_MINASSAL_LIMPO', '')
                    p_ean = r_prod.get('EAN_LIMPO', '')
                    p_sku = r_prod.get('SKU_LIMPO', '')
                    p_rec = extrair_preco_mg(r_prod)
                    
                    p_rec_str = f"R$ {p_rec:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.') if p_rec > 0 else "Não cadastrado"
                    caminho_img = obter_caminho_imagem(p_cod) or obter_caminho_imagem(p_ean) or obter_caminho_imagem(p_sku)

                    col_img, col_info, col_prc, col_btn = st.columns([0.8, 2.5, 1.5, 1])
                    
                    with col_img:
                        if caminho_img and os.path.exists(caminho_img):
                            st.image(caminho_img, width=60)
                        else:
                            st.markdown("<span style='color: #64748B; font-size: 11px;'>Sem foto</span>", unsafe_allow_html=True)
                            
                    with col_info:
                        st.markdown(f"""
                        <div style="line-height: 1.3; margin-bottom: 4px;">
                            <span style="color: #F8FAFC; font-size: 16px; font-weight: 700;">{p_nome}</span><br>
                            <span style="color: #94A3B8; font-size: 15px; font-weight: 600;">Cód: {p_cod}</span><br>
                            <span style="color: #34D399; font-size: 16px; font-weight: 800;">💰 Rec. MG: {p_rec_str}</span>
                        </div>
                        """, unsafe_allow_html=True)
                        
                    with col_prc:
                        preco_digitado = st.number_input("Preço R$", min_value=0.0, format="%.2f", key=f"add_prc_{idx_prod}", label_visibility="collapsed")
                        
                    with col_btn:
                        st.write("") 
                        if st.button("➕ Adicionar", key=f"btn_add_{idx_prod}"):
                            novo_item = {
                                'index': idx_prod,
                                'row_data': r_prod,
                                'produto': p_nome,
                                'codigo': p_cod,
                                'categoria': r_prod.get('SUBBRAND', 'Mars'),
                                'preco_recomendado': p_rec,
                                'preco_praticado': preco_digitado
                            }
                            if not any(item['index'] == idx_prod for item in st.session_state.itens_verificacao):
                                st.session_state.itens_verificacao.append(novo_item)
                                st.success(f"Adicionado!")
                                st.rerun()
                    st.markdown("<hr style='border: 0.3px solid #1E293B; margin: 8px 0;'>", unsafe_allow_html=True)
            else:
                st.info("Nenhum produto encontrado com esse termo.")

        st.markdown("---")
        st.subheader(f"📋 Produtos Adicionados para esta Loja ({len(st.session_state.itens_verificacao)} itens)")

        if st.session_state.itens_verificacao:
            for i, item_visita in enumerate(st.session_state.itens_verificacao):
                c_inf1, c_inf2, c_inf3 = st.columns([3, 2, 1])
                with c_inf1:
                    st.write(f"• **{item_visita['produto']}** (Cód: {item_visita['codigo']})")
                    st.write(f"<span style='color: #34D399; font-size: 13px; font-weight: 700;'>Rec. MG: R$ {item_visita['preco_recomendado']:.2f}</span>", unsafe_allow_html=True)
                with c_inf2:
                    st.write(f"Preço Lido: **R$ {item_visita['preco_praticado']:.2f}**")
                with c_inf3:
                    if st.button("❌ Remover", key=f"rem_{i}"):
                        st.session_state.itens_verificacao.pop(i)
                        st.rerun()
        else:
            st.info("Nenhum produto adicionado ainda. Use a busca acima para incluir os itens encontrados na gôndola.")

        st.markdown("---")
        st.subheader("📤 Opções de Envio do E-mail")
        
        tipo_envio = st.radio(
            "Enviar relatório para:",
            ["Enviar somente para o Benedito", "Enviar para a Gestão Completa (Todos)"],
            horizontal=True
        )

        if st.button("🚀 Enviar Verificação e Relatório por E-mail"):
            if not razao_social:
                st.warning("⚠️ Por favor, busque e selecione o Cliente antes de enviar.")
            elif not st.session_state.itens_verificacao:
                st.warning("⚠️ Adicione pelo menos um produto antes de enviar a verificação.")
            else:
                produtos_presentes = st.session_state.itens_verificacao
                indices_presentes = [item['index'] for item in produtos_presentes]
                oportunidades_faltantes = []
                
                for idx, row in df_produtos.iterrows():
                    if idx not in indices_presentes:
                        if eh_produto_inovacao_ou_smallbag(row):
                            oportunidades_faltantes.append({
                                'produto': row.get('PRODUTO', 'Produto'),
                                'codigo': row.get('CODIGO_MINASSAL_LIMPO', ''),
                                'categoria': row.get('SUBBRAND', 'Mars'),
                                'preco_recomendado': extrair_preco_mg(row)
                            })

                # Gera o PDF em memória
                pdf_buffer = gerar_pdf_relatorio(razao_social, endereco_cliente, bairro_cliente, produtos_presentes, oportunidades_faltantes)

                corpo_html = f"""
                <html>
                  <body style="font-family: Arial, sans-serif; color: #333;">
                    <h2 style="color: #0F172A;">🏪 Relatório de Verificação de Cliente</h2>
                    <p><b>Cliente / Razão Social:</b> {razao_social}</p>
                    <p><b>Endereço:</b> {endereco_cliente} - Bairro: {bairro_cliente}</p>
                    <p><b>Promotor:</b> Pamela</p>
                    <p><b>Localidade:</b> Poços de Caldas - MG</p>
                    <hr>
                    <p>Segue em anexo o relatório executivo em formato <b>PDF</b> contendo a verificação de preços e as oportunidades estratégicas (Inovações e Small Bags) para esta loja.</p>
                    <p style="font-size: 11px; color: #777; margin-top: 30px;">Relatório gerado automaticamente pelo App de Gestão de Campo - Minassal / Mars (Poços de Caldas - MG).</p>
                  </body>
                </html>
                """

                remetente = "benedito.bandola@gmail.com"
                senha_app = ""

                try:
                    if "email_config" in st.secrets:
                        remetente = st.secrets["email_config"].get("remetente", "benedito.bandola@gmail.com")
                        senha_app = st.secrets["email_config"].get("senha", "")
                except Exception:
                    pass

                if not senha_app:
                    st.error("⚠️ Erro: A senha de aplicativo do e-mail não foi encontrada nos Secrets do Streamlit. Configure a chave `senha` no painel do Streamlit Cloud.")
                else:
                    try:
                        if tipo_envio == "Enviar somente para o Benedito":
                            lista_destinatarios = ["benedito.bandola@minassal.com.br"]
                        else:
                            lista_destinatarios = [
                                "benedito.bandola@minassal.com.br",
                                "poli@minassal.com.br",
                                "caio.poli@minassal.com.br",
                                "daniel.santini@minassal.com.br",
                                "rubens.porfirio@minassal.com.br"
                            ]

                        msg = MIMEMultipart()
                        msg["Subject"] = f"Verificação de Cliente (PDF): {razao_social} - Poços de Caldas"
                        msg["From"] = remetente
                        msg["To"] = ", ".join(lista_destinatarios)
                        
                        # Anexa o corpo em HTML
                        msg.attach(MIMEText(corpo_html, "html"))

                        # Anexa o arquivo PDF gerado
                        parte_pdf = MIMEBase('application', 'octet-stream')
                        parte_pdf.set_payload(pdf_buffer.read())
                        encoders.encode_base64(parte_pdf)
                        nome_arquivo_pdf = f"Relatorio_Auditoria_{razao_social.replace(' ', '_')}.pdf"
                        parte_pdf.add_header('Content-Disposition', f'attachment; filename="{nome_arquivo_pdf}"')
                        msg.attach(parte_pdf)

                        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
                            server.login(remetente, senha_app)
                            server.sendmail(remetente, lista_destinatarios, msg.as_string())
                            
                        st.success(f"🎉 Relatório em PDF enviado com sucesso para: {', '.join(lista_destinatarios)}!")
                        st.session_state.itens_verificacao = []
                    except Exception as mail_err:
                        st.error(f"❌ Erro ao enviar o e-mail via SMTP: {mail_err}")
                        st.info("💡 Dica: Verifique se a senha de 16 dígitos nos Secrets está correta.")
