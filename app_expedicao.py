import re
from io import BytesIO
from datetime import datetime

import streamlit as st
import pdfplumber
import pandas as pd

from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet

# --------- máscaras de dados sensíveis ---------
def mascarar(texto: str) -> str:
    t = texto
    t = re.sub(r"\b\d{2,3}\.?\d{3}\.?\d{3}[\/\-]?\d{4}\-?\d{2}\b", "***.***.***-**", t)  # CNPJ/CPF
    t = re.sub(r"\(?\+?\d{2}\)?\s?\d{4,5}\-?\d{4}", "***-****", t)                       # Telefones
    t = re.sub(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}", "[email oculto]", t)   # E-mails
    return t

# --------- cabeçalho ---------
def extrair_header(texto: str):
    pedido = re.search(r"Pedido de Venda N[º°]\s*(\d+)", texto, re.I)
    pedido_num = pedido.group(1) if pedido else ""
    cliente = ""
    linhas = texto.splitlines()
    for i, ln in enumerate(linhas):
        if "Informações do Cliente" in ln:
            for j in range(i+1, min(i+6, len(linhas))):
                if linhas[j].strip():
                    cliente = linhas[j].strip()
                    break
            break
    ciduf = re.search(r"([A-Za-zÀ-ú\s]+)\s*-\s*([A-Z]{2})\s*-\s*CEP", texto)
    cidade = ciduf.group(1).strip() if ciduf else ""
    uf = ciduf.group(2).strip() if ciduf else ""
    inc = re.search(r"inclu[ií]do em:\s*([0-9/]{10})\s*às\s*([0-9:]{8})", texto, re.I)
    data_inclusao = f"{inc.group(1)} {inc.group(2)}" if inc else ""
    prev = re.search(r"Previs[aã]o de Faturamento:\s*([0-9/]{10})", texto, re.I)
    prev_fat = prev.group(1) if prev else ""
    obs = ""
    mobs = re.search(r"OBS[:\s]+(.+)", texto, re.I)
    if mobs: obs = mobs.group(1).strip()
    return dict(Pedido=pedido_num, Cliente=cliente, Cidade=cidade, UF=uf,
                Data_inclusao=data_inclusao, Previsao_faturamento=prev_fat,
                Obs_expedicao=obs)

# --------- itens (ajustado ao PDF enviado) ---------
def extrair_itens(texto: str):
    itens = []
    bloc = re.search(r"Itens do Pedido de Venda(.*?)(Outras Informa[cç][oõ]es|$)", texto, re.S | re.I)
    if not bloc: return itens
    secao = bloc.group(1)
    linhas = [ln.strip() for ln in secao.splitlines() if ln.strip()]

    # Ex.: "200,00 UN CXT007 CAIXA 100 DOCES/SALGADOS ..."
    rx = re.compile(r"^([\d\.\,]+)\s+([A-Za-z]+)\s+([A-Z0-9\-\/]+)\s+(.+)$")
    for ln in linhas:
        m = rx.match(ln)
        if not m: continue
        qtd_raw, unid, cod, desc = m.groups()
        qtd = float(qtd_raw.replace(".", "").replace(",", "."))
        itens.append(dict(Quantidade=qtd, Unid=unid.upper(), Codigo=cod.strip(), Descricao=desc.strip()))
    return itens

# --------- pipeline ---------
def extrair_do_pdf(pdf_bytes: bytes) -> pd.DataFrame:
    dados = []
    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            texto = page.extract_text() or ""
            txt = mascarar(texto)
            h = extrair_header(txt)
            for it in extrair_itens(txt):
                dados.append({**h, **it})
    return pd.DataFrame(dados)

# --------- PDF de saída (fonte maior) ---------
def guia_pdf(df: pd.DataFrame, tamanho_fonte=16, tamanho_desc=10) -> bytes:
    from io import BytesIO
    from datetime import datetime
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

    buf = BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=24, rightMargin=24, topMargin=24, bottomMargin=24
    )
    styles = getSampleStyleSheet()

    # estilo fixo para a Descrição
    desc_style = ParagraphStyle(
        name="Desc", parent=styles["Normal"],
        fontName="Helvetica", fontSize=tamanho_desc,
        leading=int(tamanho_desc * 1.25)  # quebra harmônica
    )

    story = []
    if df.empty:
        story.append(Paragraph("Sem itens.", styles["Normal"]))
        doc.build(story); return buf.getvalue()

    h = df.iloc[0]

    # Cabeçalho resumido com respiro
    story.append(Paragraph(f"Pedido {h.get('Pedido','')}", styles["Title"]))
    story.append(Spacer(1, 8))
    story.append(Paragraph(f"<b>Cliente:</b> {h.get('Cliente','')}", styles["Normal"]))
    story.append(Spacer(1, 6))
    story.append(Paragraph(f"<b>Inclusão:</b> {h.get('Data_inclusao','')}", styles["Normal"]))
    story.append(Paragraph(f"<b>Previsão Faturamento:</b> {h.get('Previsao_faturamento','')}", styles["Normal"]))
    story.append(Spacer(1, 6))
    obs = h.get("Obs_expedicao") or ""
    story.append(Paragraph(f"<b>Obs.:</b> {obs}", styles["Normal"]))
    story.append(Spacer(1, 12))

    # -------- TABELA --------
    header = ["Qtd", "Unid", "QTD", "CHECK", "Código", "Descrição"]
    linhas = [header]

    for _, r in df.iterrows():
        qtd_br = f"{r['Quantidade']:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        desc_par = Paragraph(str(r.get("Descricao","")), desc_style)
        linhas.append([
            qtd_br,                 # 0 Qtd
            r.get("Unid",""),       # 1 Unid
            "",                     # 2 QTD (vazia p/ escrita manual)
            "",                     # 3 CHECK (vazia)
            r.get("Codigo",""),     # 4 Código
            desc_par                # 5 Descrição (fonte fixa)
        ])

    # larguras: ajuste se quiser
    col_widths = [55, 45, 40, 45, 95, 250]
    tb = Table(linhas, colWidths=col_widths, repeatRows=1)

    tb.setStyle(TableStyle([
        # Cabeçalho
        ("FONT", (0,0), (-1,0), "Helvetica-Bold", tamanho_fonte),
        ("BACKGROUND", (0,0), (-1,0), colors.lightgrey),
        ("ALIGN", (0,0), (-1,0), "CENTER"),

        # Corpo
        ("FONT", (0,1), (1,-1), "Helvetica", tamanho_fonte),  # Qtd/Unid seguem slider
        ("FONT", (4,1), (4,-1), "Helvetica", tamanho_fonte),  # Código segue slider
        # coluna 5 (Descrição) usa Paragraph com fonte fixa -> sem FONT aqui

        ("GRID", (0,0), (-1,-1), 0.25, colors.black),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),

        # alinhamentos por coluna
        ("ALIGN", (0,1), (0,-1), "RIGHT"),    # Qtd
        ("ALIGN", (1,1), (1,-1), "RIGHT"),    # Unid
        ("ALIGN", (2,1), (3,-1), "CENTER"),   # QTD (em branco) e CHECK (em branco)
        ("ALIGN", (4,1), (4,-1), "LEFT"),     # Código
        ("ALIGN", (5,1), (5,-1), "LEFT"),     # Descrição

        # espaço maior para escrita manual nas colunas QTD/CHECK
        ("TOPPADDING", (0,1), (-1,-1), 8),
        ("BOTTOMPADDING", (0,1), (-1,-1), 8),
        ("LEFTPADDING", (0,1), (-1,-1), 6),
        ("RIGHTPADDING", (0,1), (-1,-1), 6),
    ]))

    story.append(tb)
    story.append(Spacer(1, 12))
    story.append(Paragraph(f"Gerado em {datetime.now():%d/%m/%Y %H:%M:%S}", styles["Normal"]))
    doc.build(story)
    return buf.getvalue()
# --------- APP ---------
def main():
    st.set_page_config(page_title="Extrator p/ Expedição", layout="centered")
    st.title("Extrator de Dados p/ Expedição (Papieri)")

    st.write("Envie um PDF de Pedido/NF. O app extrai os itens e gera CSV e um PDF de expedição com fonte maior.")
    fonte = st.slider("Tamanho da fonte dos itens (PDF)", 10, 22, 16)

    up = st.file_uploader("Selecione o PDF", type=["pdf"])
    if up:
        pdf_bytes = up.read()
        with st.spinner("Processando..."):
            df = extrair_do_pdf(pdf_bytes)

        if df.empty:
            st.error("Não encontrei itens. Se o layout variar, ajuste o regex em extrair_itens().")
            return

        st.subheader("Pré-visualização")
        st.dataframe(df, use_container_width=True)

        st.download_button("Baixar CSV", df.to_csv(index=False).encode("utf-8-sig"),
                           file_name="expedicao.csv", mime="text/csv")

        pdf_out = guia_pdf(df, tamanho_fonte=fonte)
        st.download_button("Baixar 'Pedido' (PDF)", data=pdf_out,
                           file_name="pedido.pdf", mime="application/pdf")

    st.caption("Dados sensíveis são mascarados automaticamente.")

if __name__ == "__main__":
    main()