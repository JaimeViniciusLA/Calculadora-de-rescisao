from datetime import datetime, date
from dateutil.relativedelta import relativedelta
import streamlit as st
import pandas as pd
import requests
from bs4 import BeautifulSoup
import io

# Importações para geração do PDF e tratamento de imagens
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage, KeepTogether
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from PIL import Image as PILImage

# Configuração da página
st.set_page_config(page_title="Simulador Rescisório FACC", page_icon="🧮", layout="wide")

st.title("🧮 Simulador de Provisionamento Rescisório Multi-Funcionários")
st.write("Calcule individualmente, adicione à lista e exporte os demonstrativos em PDF e Excel.")

# --- FUNÇÃO AUXILIAR DE FORMATAÇÃO DE MOEDA (PT-BR) ---
def fmt_moeda(valor):
    if valor is None:
        return "R$ 0,00"
    if isinstance(valor, str):
        valor = valor.replace("R$", "").replace("*", "").strip()
        try:
            valor = float(valor.replace(".", "").replace(",", "."))
        except ValueError:
            return "R$ 0,00"
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

# --- BUSCA AUTOMÁTICA DE TABELAS E DEDUÇÕES VIA WEB SCRAPING ---
@st.cache_data(ttl=86400)
def carregar_tabelas_online():
    salario_min_default = 1621.00
    teto_inss_default = 988.09  # Teto Máximo Oficial INSS
    deducao_dep_default = 189.59 # Dedução Oficial por Dependente
    simplificado_default = 607.20 # Desconto Simplificado Mensal Oficial
    
    df_inss_default = pd.DataFrame([
        {"Limite Inferior (R$)": 0.0, "Limite Superior (R$)": 1621.00, "Alíquota (%)": 7.5},
        {"Limite Inferior (R$)": 1621.00, "Limite Superior (R$)": 2902.84, "Alíquota (%)": 9.0},
        {"Limite Inferior (R$)": 2902.84, "Limite Superior (R$)": 4354.27, "Alíquota (%)": 12.0},
        {"Limite Inferior (R$)": 4354.27, "Limite Superior (R$)": 8475.55, "Alíquota (%)": 14.0},
    ])
    
    df_irrf_default = pd.DataFrame([
        {"Até (R$)": 2428.80, "Alíquota (%)": 0.0, "Dedução (R$)": 0.0},
        {"Até (R$)": 2826.65, "Alíquota (%)": 7.5, "Dedução (R$)": 182.16},
        {"Até (R$)": 3751.05, "Alíquota (%)": 15.0, "Dedução (R$)": 394.16},
        {"Até (R$)": 4664.68, "Alíquota (%)": 22.5, "Dedução (R$)": 675.49},
        {"Até (R$)": 999999.00, "Alíquota (%)": 27.5, "Dedução (R$)": 908.73},
    ])

    status_mensagem = "Offline (Tabela Padrão Carregada)"

    try:
        url_gov = "https://www.gov.br/receitafederal/pt-br/assuntos/meu-imposto-de-renda/tabelas"
        headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
        response = requests.get(url_gov, headers=headers, timeout=3)
        
        if response.status_code == 200:
            soup = BeautifulSoup(response.text, "html.parser")
            texto_pagina = soup.get_text()
            if "189,59" in texto_pagina:
                deducao_dep_default = 189.59
            if "607,20" in texto_pagina or "564,80" in texto_pagina:
                simplificado_default = 607.20
            status_mensagem = "🟢 Atualizado via Receita Federal (Gov.br)"
    except Exception:
        status_mensagem = "⚠️ Conexão falhou (Usando tabela padrão)"

    return salario_min_default, teto_inss_default, deducao_dep_default, simplificado_default, df_inss_default, df_irrf_default, status_mensagem

salario_min_base, teto_inss_base, deducao_dep_base, simplificado_base, df_inss_base, df_irrf_base, status_conexao = carregar_tabelas_online()

# --- INICIALIZAÇÃO DO ESTADO DE SESSÃO ---
if "lista_funcionarios" not in st.session_state:
    st.session_state.lista_funcionarios = []

# --- BARRA LATERAL (CNPJ E ENDEREÇO ZERADOS) ---
st.sidebar.header("🏢 Dados da Empresa (FACC)")
cnpj_empresa = st.sidebar.text_input("CNPJ da Empresa", value="", placeholder="Digite o CNPJ da empresa")
endereco_empresa = st.sidebar.text_input("Endereço/Cidade", value="", placeholder="Digite o endereço/cidade")

st.sidebar.header("🎨 Identidade Visual")
logo_upload = st.sidebar.file_uploader("Upload da Logomarca FACC", type=["png", "jpg", "jpeg", "jfif"])

st.sidebar.divider()
st.sidebar.header("⚙️ Tabelas de Tributos")
st.sidebar.caption(f"Status da busca: **{status_conexao}**")

st.sidebar.subheader("1. Tabela de INSS")
salario_minimo_nacional = st.sidebar.number_input("Salário Mínimo Vigente (R$)", value=salario_min_base, step=10.0)
df_inss_edited = st.sidebar.data_editor(df_inss_base, num_rows="dynamic", key="editor_inss")
teto_desconto_inss = st.sidebar.number_input("Teto Máximo do Desconto INSS (R$)", value=teto_inss_base, step=10.0)

st.sidebar.divider()
st.sidebar.subheader("2. Tabela de IRRF")
deducao_por_dependente_unitaria = st.sidebar.number_input("Dedução Unitária por Dependente IRRF (R$)", value=deducao_dep_base, step=1.0)
desconto_simplificado_mensal = st.sidebar.number_input("Desconto Simplificado Mensal IRRF (R$)", value=simplificado_base, step=10.0)

df_irrf_edited = st.sidebar.data_editor(df_irrf_base, num_rows="dynamic", key="editor_irrf")

# --- FUNÇÕES AUXILIARES DE CÁLCULO TRABALHISTA E TRIBUTÁRIO ---
def calcular_inss_progressivo(salario, df_inss, teto_maximo):
    inss_total = 0.0
    for _, row in df_inss.iterrows():
        lim_inf = row["Limite Inferior (R$)"]
        lim_sup = row["Limite Superior (R$)"]
        aliquota = row["Alíquota (%)"] / 100.0
        
        if salario > lim_inf:
            base = min(salario, lim_sup) - lim_inf
            inss_total += base * aliquota
        if salario <= lim_sup:
            break
            
    return min(inss_total, teto_maximo)

def calcular_irrf_completo_2026(renda_bruta, valor_inss, qtd_dep, deducao_dep_unitaria, valor_simplificado, df_irrf):
    if renda_bruta <= 0:
        return 0.0, 0.0, "Isento"
    
    deducao_legais = valor_inss + (qtd_dep * deducao_dep_unitaria)
    
    if valor_simplificado > deducao_legais:
        base_calculo = max(0.0, renda_bruta - valor_simplificado)
        regra_aplicada = "Simplificado"
    else:
        base_calculo = max(0.0, renda_bruta - deducao_legais)
        regra_aplicada = "Deduções Legais"

    irrf_bruto = 0.0
    for _, row in df_irrf.iterrows():
        ate = row["Até (R$)"]
        aliquota = row["Alíquota (%)"] / 100.0
        deducao = row["Dedução (R$)"]
        
        if base_calculo <= ate:
            irrf_bruto = (base_calculo * aliquota) - deducao
            break

    irrf_bruto = max(0.0, irrf_bruto)

    if base_calculo <= 5000.00:
        redutor_extra = min(irrf_bruto, 312.89)
    elif 5000.00 < base_calculo <= 7350.00:
        redutor_calculado = 978.62 - (0.133145 * base_calculo)
        redutor_extra = max(0.0, min(irrf_bruto, redutor_calculado))
    else:
        redutor_extra = 0.0

    irrf_final = max(0.0, irrf_bruto - redutor_extra)
    
    return irrf_final, base_calculo, regra_aplicada

def calcular_dias_aviso_previo_com_projecao(dt_admissao, dt_demissao):
    dias_aviso = 30
    while True:
        dt_fim_proj = dt_demissao + relativedelta(days=dias_aviso)
        anos_com_projecao = relativedelta(dt_fim_proj, dt_admissao).years
        novos_dias = min(30 + (anos_com_projecao * 3), 90)
        
        if novos_dias == dias_aviso:
            break
        dias_aviso = novos_dias
        
    return dias_aviso, dt_fim_proj

def calcular_avos_periodo(dt_inicio, dt_fim, limite_ano=False):
    if dt_inicio > dt_fim:
        return 0
        
    avos = 0
    curr = date(dt_inicio.year, dt_inicio.month, 1)
    
    while curr <= dt_fim:
        proximo_mes = curr + relativedelta(months=1)
        inicio_periodo = max(dt_inicio, curr)
        fim_periodo = min(dt_fim, proximo_mes - relativedelta(days=1))
        
        dias_no_mes = (fim_periodo - inicio_periodo).days + 1
        if dias_no_mes >= 15:
            avos += 1
            
        curr = proximo_mes
        
    return min(avos, 12) if limite_ano else avos

def calcular_avos_13_indenizado_direto(dias_aviso):
    avos = dias_aviso // 30
    dias_restantes = dias_aviso % 30
    if dias_restantes >= 15:
        avos += 1
    return avos

def calcular_mes_ano_projetado_fgts(str_ultimo_mes, meses_provisionar):
    try:
        if not str_ultimo_mes or '/' not in str_ultimo_mes:
            return ""
        partes = str_ultimo_mes.strip().split('/')
        mes = int(partes[0])
        ano = int(partes[1])
        dt_base = date(ano, mes, 1)
        dt_proj = dt_base + relativedelta(months=int(meses_provisionar))
        return dt_proj.strftime("%m/%Y")
    except Exception:
        return str_ultimo_mes

# --- GERADOR DE PDF ---
def gerar_pdf_funcionario(func, logo_bytes=None, cnpj="", endereco=""):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=A4, 
        rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30
    )
    story = []
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        'TitleStyle', 
        parent=styles['Heading1'], 
        fontName='Helvetica-Bold', 
        fontSize=11, 
        textColor=colors.HexColor('#1F4E79'), 
        leading=16.5
    )
    subtitle_style = ParagraphStyle('SubtitleStyle', parent=styles['Heading2'], fontName='Helvetica-Bold', fontSize=9.5, textColor=colors.HexColor('#1F4E79'))
    body_style = ParagraphStyle('BodyStyle', parent=styles['Normal'], fontName='Helvetica', fontSize=8)
    bold_style = ParagraphStyle('BoldStyle', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=8)

    texto_empresa = Paragraph(
        f"FUNDAÇÃO DE APOIO AO DESENVOLVIMENTO DA COMPUTAÇÃO CIENTÍFICA - FACC<br/>"
        f"CNPJ: {cnpj if cnpj else 'N/A'} | {endereco if endereco else 'N/A'}<br/>"
        f"DEMONSTRATIVO DE PROVISIONAMENTO DE RESCISÃO CONTRATUAL", 
        title_style
    )

    if logo_bytes is not None:
        try:
            pil_img = PILImage.open(io.BytesIO(logo_bytes))
            orig_w, orig_h = pil_img.size
            
            target_w = 140
            target_h = int(target_w * (orig_h / orig_w))
            
            img_converted_buffer = io.BytesIO()
            pil_img.convert("RGB").save(img_converted_buffer, format="PNG")
            img_converted_buffer.seek(0)
            
            img = RLImage(img_converted_buffer, width=target_w, height=target_h)
            header_data = [[img, texto_empresa]]
        except Exception:
            header_data = [[Paragraph("<b>FACC</b>", title_style), texto_empresa]]
    else:
        header_data = [[Paragraph("<b>FACC</b>", title_style), texto_empresa]]

    header_table = Table(header_data, colWidths=[150, 380])
    header_table.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('ALIGN', (0,0), (0,0), 'LEFT'),
        ('ALIGN', (1,0), (1,0), 'LEFT'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 10),
    ]))
    story.append(header_table)
    story.append(Spacer(1, 8))

    # --- SEÇÃO 1: DADOS CADASTRAIS ---
    if func["Tipo de Contrato"] == "Experiência (Rescisão Antecipada Emp.)":
        txt_contrato = f"Indenização 479 CLT ({func['Dias Faltantes Experiência']} dias p/ fim):"
        txt_termo_proj = f"Experiência ({func['Prazo Experiência Selecionado']}): {func['Data Término / Projeção']}"
    else:
        txt_contrato = f"Aviso Prévio ({func['Dias Aviso Prévio']} dias):"
        txt_termo_proj = f"Término Proj: {func['Data Término / Projeção']}"

    dados_cadastrais = [
        [Paragraph("<b>Matrícula:</b>", bold_style), Paragraph(str(func["Matrícula"]), body_style), Paragraph("<b>Data Admissão:</b>", bold_style), Paragraph(str(func["Data Admissão"]), body_style)],
        [Paragraph("<b>Nome:</b>", bold_style), Paragraph(str(func["Nome"]), body_style), Paragraph("<b>Data Demissão Proj.:</b>", bold_style), Paragraph(str(func["Data Demissão"]), body_style)],
        [Paragraph("<b>Salário Base:</b>", bold_style), Paragraph(fmt_moeda(func['Salário Base (R$)']), body_style), Paragraph(f"<b>{txt_contrato}</b>", bold_style), Paragraph(txt_termo_proj, body_style)],
        [Paragraph("<b>Periculosidade:</b>", bold_style), Paragraph(f"{func['Periculosidade (%)']}%", body_style), Paragraph("<b>Avos 13º (Trab / Av. Inden):</b>", bold_style), Paragraph(f"{func['Avos 13º Prop']}/12  |  {func['Avos 13º Aviso Indenizado']}/12", body_style)],
        [Paragraph("<b>Insalubridade:</b>", bold_style), Paragraph(f"{func['Insalubridade (%)']}%", body_style), Paragraph("<b>Avos Férias (Trab / Av. Inden):</b>", bold_style), Paragraph(f"{func['Avos Férias Prop']}/12  |  {func['Avos Férias Aviso Indenizado']}/12", body_style)],
        [Paragraph("<b>Remuneração Total:</b>", bold_style), Paragraph(f"<b>{fmt_moeda(func['Remuneração Total (R$)'])}</b>", bold_style), Paragraph("<b>Dependentes IRRF:</b>", bold_style), Paragraph(f"{func['Dependentes IRRF']} (Dedução: {fmt_moeda(func['Dedução Dependentes Total (R$)'])})", body_style)],
        [Paragraph("<b>Saldo FGTS em ({0}):</b>".format(func['Último Mês FGTS'] if func['Último Mês FGTS'] else "N/A"), bold_style), Paragraph(fmt_moeda(func['Saldo FGTS Atual (R$)']), body_style), Paragraph("<b>FGTS a Provisionar ({0} mes(es)):</b>".format(func['Meses Provisionados FGTS']), bold_style), Paragraph(f"{fmt_moeda(func['FGTS Futuro Provisionado (R$)'])} (Até {func['Mês/Ano Projetado FGTS'] if func['Mês/Ano Projetado FGTS'] else 'N/A'})", body_style)]
    ]

    t_cadastrais = Table(dados_cadastrais, colWidths=[110, 150, 140, 130])
    t_cadastrais.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F5F7FA')),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E0E0E0')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 3),
        ('BOTTOMPADDING', (0,0), (-1,-1), 3),
    ]))
    
    bloco_sec1 = [
        Paragraph("1. DADOS CADASTRAIS E PROJEÇÃO CONTRATUAL", subtitle_style),
        Spacer(1, 4),
        t_cadastrais,
        Spacer(1, 8)
    ]
    story.append(KeepTogether(bloco_sec1))

    # --- SEÇÃO 2: DEMONSTRATIVO DE VERBAS PROVISIONADAS ---
    verbas_data = [
        [Paragraph("<b>Item / Descrição da Verba</b>", bold_style), Paragraph("<b>Base / Ref. Legal</b>", bold_style), Paragraph("<b>Valor Provisionado (R$)</b>", bold_style)],
        [Paragraph(f"Saldo de Salário ({func['Dias Saldo Salário']} dias)", body_style), Paragraph(f"{func['Dias Saldo Salário']} dia(s)", body_style), Paragraph(fmt_moeda(func['Saldo de Salário Base (R$)']), body_style)]
    ]

    if func["Periculosidade (%)"] > 0:
        verbas_data.append([
            Paragraph(f"Adicional de Periculosidade ({func['Dias Saldo Salário']} dias)", body_style),
            Paragraph(f"{func['Periculosidade (%)']}% s/ Salário Base", body_style),
            Paragraph(fmt_moeda(func['Adicional Periculosidade Saldo (R$)']), body_style)
        ])

    if func["Insalubridade (%)"] > 0:
        verbas_data.append([
            Paragraph(f"Adicional de Insalubridade ({func['Dias Saldo Salário']} dias)", body_style),
            Paragraph(f"{func['Insalubridade (%)']}% s/ Salário Mínimo", body_style),
            Paragraph(fmt_moeda(func['Adicional Insalubridade Saldo (R$)']), body_style)
        ])

    if func["Tipo de Contrato"] == "Experiência (Rescisão Antecipada Emp.)":
        verbas_data.append([
            Paragraph(f"Indenização Art. 479 CLT ({func['Dias Faltantes Experiência']} dias a 50%)", body_style),
            Paragraph("50% s/ saldo do contrato exp.", body_style),
            Paragraph(fmt_moeda(func['Indenização 479 CLT (R$)']), body_style)
        ])
    else:
        verbas_data.append([
            Paragraph(f"Aviso Prévio Indenizado ({func['Dias Aviso Prévio']} dias)", body_style),
            Paragraph("Lei 12.506/2011", body_style),
            Paragraph(fmt_moeda(func['Aviso Prévio (R$)']), body_style)
        ])

    verbas_data.extend([
        [Paragraph("Férias Vencidas + 1/3", body_style), Paragraph(f"{func['Períodos Férias Vencidas']} Período(s)", body_style), Paragraph(fmt_moeda(func['Férias Vencidas + 1/3 (R$)']), body_style)],
        [Paragraph("Férias Proporcionais + 1/3 (Trabalhada)", body_style), Paragraph(f"{func['Avos Férias Prop']}/12 avos", body_style), Paragraph(fmt_moeda(func['Férias Proporcionais + 1/3 (R$)']), body_style)],
        [Paragraph("13º Salário Proporcional (Trabalhado)", body_style), Paragraph(f"{func['Avos 13º Prop']}/12 avos", body_style), Paragraph(fmt_moeda(func['13º Proporcional (R$)']), body_style)],
        [Paragraph("Férias Indenizadas + 1/3 (Aviso Indenizado)", body_style), Paragraph(f"{func['Avos Férias Aviso Indenizado']}/12 avos", body_style), Paragraph(fmt_moeda(func['Férias Indenizadas + 1/3 (R$)']), body_style)],
        [Paragraph("13º Salário Indenizado (Aviso Indenizado)", body_style), Paragraph(f"{func['Avos 13º Aviso Indenizado']}/12 avos", body_style), Paragraph(fmt_moeda(func['13º Indenizado (R$)']), body_style)],
        [Paragraph("FGTS s/ Verbas Rescisórias (8%)", body_style), Paragraph("8% s/ (Bases rescisórias tributáveis)", body_style), Paragraph(fmt_moeda(func['FGTS s/ Verbas Rescisórias (R$)']), body_style)],
        [Paragraph("FGTS s/ 13º Salário (8%)", body_style), Paragraph("8% s/ Total 13º", body_style), Paragraph(fmt_moeda(func['FGTS s/ 13º (R$)']), body_style)],
        [Paragraph("Multa Rescisória FGTS (40%)", body_style), Paragraph(f"40% s/ Soma dos FGTS ({fmt_moeda(func['Base Total FGTS p/ Multa (R$)'])})", body_style), Paragraph(fmt_moeda(func['Multa 40% FGTS (R$)']), body_style)],
        [Paragraph("Multa do Trintídio (Art. 9º Lei 7.238/84)", body_style), Paragraph("Indenização Dissídio", body_style), Paragraph(fmt_moeda(func['Multa Trintídio (R$)']), body_style)],
        [Paragraph("<b>CUSTO TOTAL DE PROVISIONAMENTO (FACC)</b>", bold_style), Paragraph("", body_style), Paragraph(f"<b>{fmt_moeda(func['Custo Total (R$)'])}</b>", bold_style)],
    ])

    t_verbas = Table(verbas_data, colWidths=[230, 150, 150])
    t_verbas.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#2E75B6')),
        ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#D3D3D3')),
        ('ALIGN', (2,0), (2,-1), 'RIGHT'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor('#EAECEE')),
        ('TOPPADDING', (0,0), (-1,-1), 3.5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 3.5),
    ]))

    bloco_sec2 = [
        Paragraph("2. DEMONSTRATIVO DE VERBAS PROVISIONADAS", subtitle_style),
        Spacer(1, 4),
        t_verbas,
        Spacer(1, 8)
    ]
    story.append(KeepTogether(bloco_sec2))

    # --- SEÇÃO 3: BASES DE CÁLCULO DE ENCARGOS E TRIBUTOS ---
    total_base_13_inss = func['13º Proporcional (R$)'] + func['13º Indenizado (R$)']
    base_fgts_verbas = func['Aviso Prévio (R$)'] + func['Saldo de Salário Total (R$)']
    
    regra_inss_saldo = f"Teto {fmt_moeda(func['Teto INSS Aplicado (R$)'])}" if func['INSS Saldo (R$)'] >= func['Teto INSS Aplicado (R$)'] else f"Prog. ({(func['INSS Saldo (R$)']/func['Base INSS Saldo (R$)'])*100:.2f}%)" if func['Base INSS Saldo (R$)'] > 0 else "Prog."
    regra_inss_13 = f"Teto {fmt_moeda(func['Teto INSS Aplicado (R$)'])}" if func['INSS 13º (R$)'] >= func['Teto INSS Aplicado (R$)'] else f"Prog. ({(func['INSS 13º (R$)']/total_base_13_inss)*100:.2f}%)" if total_base_13_inss > 0 else "Prog."

    encargos_data = [
        [Paragraph("<b>Encargo / Tributo</b>", bold_style), Paragraph("<b>Base de Cálculo (R$)</b>", bold_style), Paragraph("<b>Alíquota / Regra</b>", bold_style), Paragraph("<b>Valor Calculado (R$)</b>", bold_style)],
        [Paragraph("FGTS - Verbas Rescisórias (8%)", body_style), Paragraph(fmt_moeda(base_fgts_verbas), body_style), Paragraph("8.00%", body_style), Paragraph(fmt_moeda(func['FGTS s/ Verbas Rescisórias (R$)']), body_style)],
        [Paragraph("FGTS - 13º Salário (8%)", body_style), Paragraph(fmt_moeda(total_base_13_inss), body_style), Paragraph("8.00%", body_style), Paragraph(fmt_moeda(func['FGTS s/ 13º (R$)']), body_style)],
        [Paragraph("Multa Rescisória FGTS (40%)", body_style), Paragraph(fmt_moeda(func['Base Total FGTS p/ Multa (R$)']), body_style), Paragraph("40.00%", body_style), Paragraph(fmt_moeda(func['Multa 40% FGTS (R$)']), body_style)],
        [Paragraph("INSS - Saldo de Salário e Adicional", body_style), Paragraph(fmt_moeda(func['Base INSS Saldo (R$)']), body_style), Paragraph(regra_inss_saldo, body_style), Paragraph(fmt_moeda(func['INSS Saldo (R$)']), body_style)],
        [Paragraph("INSS - 13º Salário (Total)", body_style), Paragraph(fmt_moeda(total_base_13_inss), body_style), Paragraph(regra_inss_13, body_style), Paragraph(fmt_moeda(func['INSS 13º (R$)']), body_style)],
        [Paragraph("IRRF - Saldo de Salário", body_style), Paragraph(fmt_moeda(func['Base IRRF Saldo (R$)']), body_style), Paragraph(f"Tabela 2026 ({func.get('Regra IRRF Saldo', 'Deduções Legais')})", body_style), Paragraph(fmt_moeda(func['IRRF Saldo (R$)']), body_style)],
        [Paragraph("IRRF - 13º Salário (Proporcional)", body_style), Paragraph(fmt_moeda(func['Base IRRF 13º (R$)']), body_style), Paragraph(f"Tabela 2026 ({func.get('Regra IRRF 13º', 'Deduções Legais')})", body_style), Paragraph(fmt_moeda(func['IRRF 13º (R$)']), body_style)],
    ]

    t_encargos = Table(encargos_data, colWidths=[180, 130, 110, 110])
    t_encargos.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#2E75B6')),
        ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#D3D3D3')),
        ('ALIGN', (1,0), (-1,-1), 'RIGHT'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 3.5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 3.5),
    ]))

    bloco_sec3 = [
        Paragraph("3. BASES DE CÁLCULO E DEMONSTRATIVO DE ENCARGOS / TRIBUTOS", subtitle_style),
        Spacer(1, 4),
        t_encargos,
        Spacer(1, 8)
    ]
    story.append(KeepTogether(bloco_sec3))

    # --- SEÇÃO 4: RESUMO FINANCEIRO E LÍQUIDO A PAGAR ---
    resumo_data = [
        [Paragraph("<b>Descrição das Verbas e Descontos</b>", bold_style), Paragraph("<b>Valor (R$)</b>", bold_style)],
        [Paragraph("Total de Verbas Pagas ao Trabalhador (TRCT Bruto)", body_style), Paragraph(fmt_moeda(func['TRCT Bruto (R$)']), body_style)],
        [Paragraph("(-) Total Encargos INSS e IRRF", body_style), Paragraph(f"- {fmt_moeda(func['Total Encargos INSS e IRRF (R$)'])}", body_style)],
        [Paragraph("<b>LÍQUIDO A PAGAR DA RESCISÃO (TRCT)</b>", bold_style), Paragraph(f"<b>{fmt_moeda(func['Líquido a Pagar (R$)'])}</b>", bold_style)],
        [Paragraph("<i>(+) Total de Encargos FGTS Rescisórios (Verbas + 13º + Multa 40%)</i>", body_style), Paragraph(f"<i>{fmt_moeda(func['Total Encargos FGTS (R$)'])}</i>", body_style)]
    ]

    t_resumo = Table(resumo_data, colWidths=[380, 150])
    t_resumo.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#2E75B6')),
        ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#D3D3D3')),
        ('ALIGN', (1,0), (1,-1), 'RIGHT'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BACKGROUND', (0,3), (-1,3), colors.HexColor('#D9E1F2')),
        ('TOPPADDING', (0,0), (-1,-1), 3.5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 3.5),
    ]))

    bloco_sec4 = [
        Paragraph("4. RESUMO FINANCEIRO (LÍQUIDO A PAGAR AO FUNCIONÁRIO)", subtitle_style),
        Spacer(1, 4),
        t_resumo
    ]
    story.append(KeepTogether(bloco_sec4))

    doc.build(story)
    buffer.seek(0)
    return buffer

# --- FORMULÁRIO PRINCIPAL (INICIALIZAÇÃO TOTALMENTE ZERADA E LIMPA) ---
st.divider()
st.subheader("📋 Cadastrar / Calcular Funcionário")

if "form_matricula" not in st.session_state:
    st.session_state.form_matricula = ""
if "form_nome" not in st.session_state:
    st.session_state.form_nome = ""
if "form_ultimo_mes_fgts" not in st.session_state:
    st.session_state.form_ultimo_mes_fgts = ""
if "form_data_admissao" not in st.session_state:
    st.session_state.form_data_admissao = None
if "form_data_demissao" not in st.session_state:
    st.session_state.form_data_demissao = None

col1, col2 = st.columns(2)

with col1:
    matricula = st.text_input("Matrícula", key="form_matricula", placeholder="Digite a matrícula")
    nome = st.text_input("Nome do Funcionário", key="form_nome", placeholder="Digite o nome completo")
    
    tipo_contrato = st.selectbox(
        "Tipo de Contrato de Trabalho",
        options=["Sem Determinação de Prazo", "Experiência (Rescisão Antecipada Emp.)"],
        index=0
    )
    
    salario_base = st.number_input("Salário Base (R$)", value=0.0, step=100.0)
    
    st.markdown("**Adicionais da Remuneração:**")
    col_per_chk, col_per_pct = st.columns([1, 1])
    with col_per_chk:
        periculosidade = st.checkbox("Possui Periculosidade", value=False)
    with col_per_pct:
        pct_periculosidade = st.number_input(
            "Percentual Periculosidade (%)", 
            value=0.0, 
            min_value=0.0, 
            max_value=100.0, 
            step=5.0,
            disabled=not periculosidade
        )
    
    col_ins_chk, col_ins_pct = st.columns([1, 1])
    with col_ins_chk:
        insalubridade = st.checkbox("Possui Insalubridade", value=False)
    with col_ins_pct:
        pct_insalubridade = st.number_input(
            "Percentual Insalubridade (%)", 
            value=0.0, 
            min_value=0.0, 
            max_value=100.0, 
            step=10.0,
            disabled=not insalubridade
        )

    st.markdown("**Férias Vencidas:**")
    col_fer_chk, col_fer_qtd = st.columns([1, 1])
    with col_fer_chk:
        ferias_vencidas_chk = st.checkbox("Possui Férias Vencidas?", value=False)
    with col_fer_qtd:
        qtd_ferias_vencidas = st.number_input(
            "Nº de Períodos Vencidos", 
            value=1, 
            min_value=1, 
            max_value=5, 
            step=1, 
            disabled=not ferias_vencidas_chk
        )

with col2:
    data_admissao = st.date_input("Data de Admissão", key="form_data_admissao", value=None, format="DD/MM/YYYY")
    data_demissao = st.date_input("Data de Demissão Provisionada", key="form_data_demissao", value=None, format="DD/MM/YYYY")
    
    if tipo_contrato == "Experiência (Rescisão Antecipada Emp.)":
        opcao_prazo_exp = st.selectbox(
            "Duração Total do Contrato de Experiência",
            options=["45 + 45 (90 dias)", "30 + 60 (90 dias)", "30 + 30 (60 dias)", "45 dias", "30 dias", "Outro (informar dias)"],
            index=0
        )
        
        if opcao_prazo_exp == "45 + 45 (90 dias)" or opcao_prazo_exp == "30 + 60 (90 dias)":
            dias_duracao_exp = 90
        elif opcao_prazo_exp == "30 + 30 (60 dias)":
            dias_duracao_exp = 60
        elif opcao_prazo_exp == "45 dias":
            dias_duracao_exp = 45
        elif opcao_prazo_exp == "30 dias":
            dias_duracao_exp = 30
        else:
            dias_duracao_exp = st.number_input("Informe a quantidade total de dias de experiência", value=90, min_value=1, max_value=90, step=1)
            
        if data_admissao:
            data_fim_experiencia = data_admissao + relativedelta(days=int(dias_duracao_exp))
            data_termino_projecao_str = data_fim_experiencia.strftime("%d/%m/%Y")
            st.info(f"🗓️ **Término Calculado do Contrato de Experiência:** {data_termino_projecao_str}")
        else:
            data_fim_experiencia = None
    else:
        opcao_prazo_exp = "N/A"
        dias_duracao_exp = 0
        if data_admissao and data_demissao:
            data_fim_experiencia = data_demissao
            dias_aviso_temp, data_fim_proj_temp = calcular_dias_aviso_previo_com_projecao(data_admissao, data_demissao)
            data_termino_projecao_str = data_fim_proj_temp.strftime("%d/%m/%Y")
        else:
            data_fim_experiencia = None

    col_sal_dias, col_dep_ir = st.columns([1, 1])
    with col_sal_dias:
        dias_saldo_salario = st.number_input("Dias de Saldo de Salário", value=0, min_value=0, max_value=31, step=1)
    with col_dep_ir:
        qtd_dependentes_irrf = st.number_input("Número de Dependentes IRRF", value=0, min_value=0, max_value=10, step=1)

    mes_dissidio = st.selectbox(
        "Mês da Data-Base do Dissídio/Convenção",
        options=list(range(1, 13)),
        format_func=lambda x: [
            "1 - Janeiro", "2 - Fevereiro", "3 - Março", "4 - Abril", "5 - Maio", "6 - Junho",
            "7 - Julho", "8 - Agosto", "9 - Setembro", "10 - Outubro", "11 - Novembro", "12 - Dezembro"
        ][x-1],
        index=0
    )
    
    saldo_fgts_atual = st.number_input("Saldo Atual do FGTS (R$)", value=0.0, step=500.0)
    ultimo_mes_fgts = st.text_input("Último Mês Depositado no Saldo do FGTS (MM/AAAA)", key="form_ultimo_mes_fgts", placeholder="Ex: 08/2026")
    meses_provisionar = st.number_input("Meses a Provisionar (FGTS futuro)", value=0, min_value=0)

col_calc, col_clear = st.columns([3, 1])

with col_clear:
    if st.button("🧹 Limpar Campos", use_container_width=True):
        st.session_state.form_matricula = ""
        st.session_state.form_nome = ""
        st.session_state.form_ultimo_mes_fgts = ""
        st.session_state.form_data_admissao = None
        st.session_state.form_data_demissao = None
        st.rerun()

with col_calc:
    btn_calcular = st.button("➕ Calcular e Salvar Funcionário", type="primary", use_container_width=True)
