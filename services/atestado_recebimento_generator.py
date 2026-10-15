"""
services/atestado_recebimento_generator.py - Gerador de Atestado de Recebimento de Material/Serviço
Gera em PDF conforme o modelo militar oficial do Estado do Rio Grande do Sul / Brigada Militar / 46º BPM.

Suporta perfeitamente as 3 variações documentadas:
1. Somente NFS-e (Serviço)
2. Somente DANFE (Material)
3. Ambas juntas (DANFE + NFS-e)
"""

from __future__ import annotations

import os
from pathlib import Path
from datetime import datetime
from typing import Dict, Any, Optional

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    Image as RLImage,
    KeepTogether
)

from services.atestado_generator import valor_monetario_extenso, MESES_PT


def formatar_data_militar(dt: Optional[datetime | str] = None, cidade: Optional[str] = None) -> str:
    """Formata a data no padrão militar tradicional (ex: Quartel em Bagé, RS, 1º de outubro de 2026)."""
    if dt is None:
        dt = datetime.now()
    elif isinstance(dt, str):
        try:
            dt = datetime.strptime(dt[:10], '%Y-%m-%d')
        except Exception:
            dt = datetime.now()

    dia_str = "1º" if dt.day == 1 else f"{dt.day:02d}"
    mes_str = MESES_PT.get(dt.month, 'janeiro')
    local = f"Quartel em {cidade.strip()}, " if cidade and cidade.strip() else "Quartel, "
    return f"{local}{dia_str} de {mes_str} de {dt.year}."


def gerar_texto_padrao_recebimento(dados: Dict[str, Any]) -> str:
    """
    Gera o texto oficial padrão do Atestado de Recebimento de Material/Serviço
    adaptado especificamente para:
    - Somente NFS-e (Serviço)
    - Somente DANFE (Material)
    - Ambas juntas (DANFE + NFS-e)
    """
    tipo = str(dados.get('tipo_documento_fiscal') or 'nfse').strip().lower()
    
    # Dados da Viatura
    prefixo = str(dados.get('viatura_prefixo') or '').strip()
    marca = str(dados.get('viatura_marca') or 'Toyota').strip()
    modelo = str(dados.get('viatura_modelo') or 'Hilux').strip()
    placa = str(dados.get('viatura_placa') or '').strip().upper()
    cidade = str(dados.get('cidade_carga') or 'Bagé').strip()

    # Dados da Empresa / Prestador
    empresa = str(dados.get('empresa_nome') or dados.get('oficina') or 'EMPRESA PRESTADORA').strip()
    cnpj = str(dados.get('empresa_cnpj') or '00.000.000/0001-00').strip()

    # Dados do Militar Responsável pelo Recebimento
    nome_resp = str(dados.get('responsavel_nome') or 'JULIANO INOCENTE VELEDA').strip().upper()
    grad_resp = str(dados.get('responsavel_grad') or 'Sd PM').strip()
    id_func = str(dados.get('responsavel_id_func') or '3155358').strip()

    # Dados fiscais
    num_danfe = str(dados.get('numero_danfe') or '').strip()
    val_danfe = float(dados.get('valor_danfe') or 0.0)
    val_danfe_str = f"{val_danfe:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    val_danfe_ext = valor_monetario_extenso(val_danfe)

    num_nfse = str(dados.get('numero_nfse') or '').strip()
    val_nfse = float(dados.get('valor_nfse') or 0.0)
    val_nfse_str = f"{val_nfse:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    val_nfse_ext = valor_monetario_extenso(val_nfse)

    # Identificação do Tipo (nfse, danfe ou ambos)
    tem_danfe = bool(num_danfe and val_danfe > 0) or tipo == 'danfe'
    tem_nfse = bool(num_nfse and val_nfse > 0) or tipo == 'nfse'

    if tipo == 'ambos' or (tem_danfe and tem_nfse):
        # CASO 3: AMBAS JUNTAS (DANFE + NFS-e)
        texto = (
            f"Atesto que o material constante na DANFE nº <b>{num_danfe}</b>, no valor de "
            f"<b>R$ {val_danfe_str}</b> ({val_danfe_ext}) foi recebido e está sob a responsabilidade do "
            f"{grad_resp} <b>{nome_resp}</b>, Id Func {id_func}, bem como foram realizados os serviços "
            f"da Nota Fiscal de Serviços Eletrônica – NFS-e nº <b>{num_nfse}</b>, no valor de "
            f"<b>R$ {val_nfse_str}</b> ({val_nfse_ext}), ambas da empresa <b>{empresa}</b>, CNPJ {cnpj}; "
            f"e destinam-se a manutenção e conservação da viatura BM prefixo <b>{prefixo}</b>, marca {marca}, "
            f"modelo {modelo}, placa <b>{placa}</b>, pertencente à carga da cidade de {cidade}.\n\n"
            f"Conforme consulta, a referida DANFE e a NFS-e encontram-se devidamente cadastradas no site da Secretaria da Fazenda."
        )
    elif tipo == 'danfe' or (tem_danfe and not tem_nfse):
        # CASO 2: SOMENTE DANFE (MATERIAL)
        texto = (
            f"Atesto que o material constante na DANFE nº <b>{num_danfe}</b>, no valor de "
            f"<b>R$ {val_danfe_str}</b> ({val_danfe_ext}) foi recebido e está sob a responsabilidade do "
            f"{grad_resp} <b>{nome_resp}</b>, Id Func {id_func}, da empresa <b>{empresa}</b>, CNPJ {cnpj}; "
            f"e destinam-se a manutenção e conservação da viatura BM prefixo <b>{prefixo}</b>, marca {marca}, "
            f"modelo {modelo}, placa <b>{placa}</b>, pertencente à carga da cidade de {cidade}.\n\n"
            f"Conforme consulta, a referida DANFE encontra-se devidamente cadastrada no site da Secretaria da Fazenda."
        )
    else:
        # CASO 1: SOMENTE NFS-e (SERVIÇO)
        texto = (
            f"Atesto que foi realizado o serviço da Nota Fiscal de Serviços Eletrônica – NFS-e nº <b>{num_nfse}</b>, no valor de "
            f"<b>R$ {val_nfse_str}</b> ({val_nfse_ext}) foi recebido e está sob a responsabilidade do "
            f"{grad_resp} <b>{nome_resp}</b>, Id Func {id_func}, da empresa <b>{empresa}</b>, CNPJ {cnpj}; "
            f"e destina-se a manutenção e conservação da viatura BM prefixo <b>{prefixo}</b>, marca {marca}, "
            f"modelo {modelo}, placa <b>{placa}</b>, pertencente à carga da cidade de {cidade}.\n\n"
            f"Conforme consulta, a referida NFS-e encontra-se devidamente cadastrada no site da Secretaria da Fazenda."
        )

    return texto


def gerar_atestado_recebimento_pdf(
    pdf_path: str | Path,
    dados: Dict[str, Any],
    texto_personalizado: Optional[str] = None
) -> None:
    """
    Gera o PDF oficial do ATESTADO DE RECEBIMENTO DE MATERIAL/SERVIÇO
    formatado com rigor militar e fidelidade ao modelo do 46º BPM.
    """
    pdf_path = Path(pdf_path)
    pdf_path.parent.mkdir(parents=True, exist_ok=True)

    # Margens do modelo oficial A4 retrato: Top 12mm, Left 30mm, Right 15mm, Bottom 20mm
    doc = SimpleDocTemplate(
        str(pdf_path),
        pagesize=A4,
        leftMargin=30 * mm,
        rightMargin=15 * mm,
        topMargin=12 * mm,
        bottomMargin=20 * mm
    )

    largura_util = A4[0] - (45 * mm)  # 210mm - 45mm = 165mm

    # Assets gráficos
    base_dir = Path(__file__).resolve().parent.parent
    img_dir = base_dir / "static" / "img"
    img_brasao_rs = img_dir / "brasao_rs_armas.png"
    img_brasao_bm = img_dir / "brasao_bm_atestado.png"
    img_rodape = img_dir / "rodape_redes_sociais.png"

    # Estilos tipográficos
    style_center_bold = ParagraphStyle(
        'HeaderCenterBold',
        fontName='Times-Bold',
        fontSize=12,
        leading=14,
        alignment=1,
        textColor=colors.black
    )

    style_titulo = ParagraphStyle(
        'TituloRecebimento',
        fontName='Times-Bold',
        fontSize=12,
        leading=14,
        alignment=1,
        textColor=colors.black
    )

    style_corpo = ParagraphStyle(
        'CorpoTextoRecebimento',
        fontName='Times-Roman',
        fontSize=12,
        leading=18,
        alignment=4,  # Justificado
        firstLineIndent=20 * mm,
        textColor=colors.black
    )

    style_data = ParagraphStyle(
        'DataGeracaoRecebimento',
        fontName='Times-Roman',
        fontSize=12,
        leading=15,
        alignment=0,  # À esquerda
        textColor=colors.black
    )

    style_sig_nome = ParagraphStyle(
        'SigNomeRecebimento',
        fontName='Helvetica-Bold',
        fontSize=10,
        leading=12,
        alignment=0,
        textColor=colors.black
    )

    style_sig_cargo = ParagraphStyle(
        'SigCargoRecebimento',
        fontName='Helvetica',
        fontSize=8,
        leading=10,
        alignment=0,
        textColor=colors.black
    )

    style_endereco = ParagraphStyle(
        'EnderecoRecebimento',
        fontName='Helvetica',
        fontSize=8,
        leading=10,
        alignment=0,
        textColor=colors.HexColor("#333333")
    )

    elementos = []

    # 1. Brasão das Armas do RS (Topo)
    if img_brasao_rs.exists():
        img_rs = RLImage(str(img_brasao_rs), width=21 * mm, height=26 * mm)
        img_rs.hAlign = 'CENTER'
        elementos.append(img_rs)
        elementos.append(Spacer(1, 2 * mm))

    # 2. Cabeçalho Oficial
    elementos.append(Paragraph("ESTADO DO RIO GRANDE DO SUL", style_center_bold))
    elementos.append(Paragraph("SECRETARIA DA SEGURANÇA PÚBLICA", style_center_bold))
    elementos.append(Paragraph("BRIGADA MILITAR", style_center_bold))
    elementos.append(Paragraph("CRPMFO – 46º BPM", style_center_bold))
    elementos.append(Spacer(1, 10 * mm))

    # 3. Título do Atestado
    elementos.append(Paragraph("ATESTADO DE RECEBIMENTO DE MATERIAL/SERVIÇO", style_titulo))
    elementos.append(Spacer(1, 10 * mm))

    # 4. Texto do Atestado (Personalizado ou Padrão)
    if texto_personalizado and str(texto_personalizado).strip():
        texto_final = str(texto_personalizado).strip()
    else:
        texto_final = gerar_texto_padrao_recebimento(dados)

    import html
    paragrafos = [p.strip() for p in texto_final.replace('\r\n', '\n').split('\n\n') if p.strip()]
    if not paragrafos:
        paragrafos = [texto_final]

    for i, p_texto in enumerate(paragrafos):
        p_html = p_texto.replace('\n', '<br/>')
        try:
            p_elem = Paragraph(p_html, style_corpo)
        except Exception:
            p_elem = Paragraph(html.escape(p_texto).replace('\n', '<br/>'), style_corpo)
        elementos.append(p_elem)
        if i < len(paragrafos) - 1:
            elementos.append(Spacer(1, 4 * mm))

    elementos.append(Spacer(1, 8 * mm))

    # 5. Data Militar — usa a cidade da viatura dinamicamente
    cidade_doc = str(dados.get('cidade_carga') or '').strip()
    data_str = formatar_data_militar(dados.get('data_servico') or dados.get('data_geracao'), cidade=cidade_doc)
    elementos.append(Paragraph(data_str, style_data))

    # 6. Bloco de Assinatura do Militar Responsável (Espaço de 24mm para GOV.BR)
    nome_resp = str(dados.get('responsavel_nome') or 'JULIANO INOCENTE VELEDA').strip().upper()
    grad_resp = str(dados.get('responsavel_grad') or 'Sd PM').strip()
    cargo_resp = str(dados.get('responsavel_cargo') or 'Aux P4/46º BPM').strip()

    bloco_assinatura = [
        Spacer(1, 24 * mm),  # Espaço para chancela digital GOV.BR
        Paragraph(nome_resp, style_sig_nome),
        Spacer(1, 1 * mm),
        Paragraph(f"{grad_resp} – {cargo_resp}", style_sig_cargo)
    ]

    # 7. Bloco de Rodapé Oficial (Endereço, Brasão BM e Banner Social)
    bloco_endereco = [
        Paragraph("Brigada Militar – CRPMFO / 46º BPM", style_endereco),
        Paragraph("Fone: (53) 3240 3400", style_endereco)
    ]


    bloco_brasao_bm = []
    if img_brasao_bm.exists():
        img_bm = RLImage(str(img_brasao_bm), width=23 * mm, height=23 * mm)
        img_bm.hAlign = 'RIGHT'
        bloco_brasao_bm.append(img_bm)

    tab_rodape = Table([
        [bloco_endereco, bloco_brasao_bm]
    ], colWidths=[largura_util - 30 * mm, 30 * mm])
    tab_rodape.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (-1, -1), 0),
        ('RIGHTPADDING', (0, 0), (-1, -1), 0),
        ('TOPPADDING', (0, 0), (-1, -1), 0),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
        ('ALIGN', (1, 0), (1, 0), 'RIGHT'),
    ]))

    bloco_final = [
        *bloco_assinatura,
        Spacer(1, 9),  # Espaço de tamanho de fonte 9
        tab_rodape,
        Spacer(1, 2 * mm)
    ]

    if img_rodape.exists():
        img_redes = RLImage(str(img_rodape), width=largura_util, height=14.5 * mm)
        img_redes.hAlign = 'CENTER'
        bloco_final.append(img_redes)

    elementos.append(KeepTogether(bloco_final))

    # Constrói o PDF em página única A4
    doc.build(elementos)
