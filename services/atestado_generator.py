"""
Módulo para Geração do ATESTADO DE COMBUSTÍVEIS do 46º BPM em PDF.

Reproduz fielmente o modelo oficial do Estado do Rio Grande do Sul / Brigada Militar / 46º BPM:
- Brasão de Armas do RS no topo
- Cabeçalho oficial centralizado (CRPMFO – 46º BPM)
- Título ATESTADO
- Corpo justificando os gastos totais com conversão monetária por extenso em reais e centavos
- Reconhecimento automático da quinzena (1ª ou 2ª) e do mês de referência
- Data de geração atualizada dinamicamente (ex: Quartel em Bagé, RS, DD de MÊS de AAAA)
- Bloco de assinaturas com Comandante (no impedimento), Subcomandante e Gestor de Frota
- Brasão da Brigada Militar e rodapé oficial com redes sociais
"""

from __future__ import annotations

import os
from pathlib import Path
from datetime import datetime
from typing import List, Dict, Any, Optional

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    Image as RLImage,
    KeepTogether
)

from services.report_config import get_assinatura_config


MESES_PT = {
    1: 'janeiro', 2: 'fevereiro', 3: 'março', 4: 'abril',
    5: 'maio', 6: 'junho', 7: 'julho', 8: 'agosto',
    9: 'setembro', 10: 'outubro', 11: 'novembro', 12: 'dezembro'
}

UNIDADES = ['', 'um', 'dois', 'três', 'quatro', 'cinco', 'seis', 'sete', 'oito', 'nove']
DEZ_A_DEZENOVE = [
    'dez', 'onze', 'doze', 'treze', 'quatorze', 'quinze',
    'dezesseis', 'dezessete', 'dezoito', 'dezenove'
]
DEZENAS = [
    '', '', 'vinte', 'trinta', 'quarenta', 'cinquenta',
    'sessenta', 'setenta', 'oitenta', 'noventa'
]
CENTENAS = [
    '', 'cento', 'duzentos', 'trezentos', 'quatrocentos',
    'quinhentos', 'seiscentos', 'setecentos', 'oitocentos', 'novecentos'
]


def _centena_para_extenso(n: int) -> str:
    """Converte um número de 0 a 999 em palavras."""
    if n == 0:
        return ''
    if n == 100:
        return 'cem'
    c = n // 100
    d = (n % 100) // 10
    u = n % 10
    partes = []
    if c > 0:
        partes.append(CENTENAS[c])
    if d == 1:
        partes.append(DEZ_A_DEZENOVE[u])
    else:
        if d > 1:
            partes.append(DEZENAS[d])
        if u > 0:
            partes.append(UNIDADES[u])
    return ' e '.join(partes)


def numero_para_extenso(n: int) -> str:
    """Converte inteiro positivo em extenso com separação correta de milhares e milhões."""
    if n == 0:
        return 'zero'
    milhoes = n // 1000000
    milhares = (n % 1000000) // 1000
    unidades = n % 1000

    partes = []
    if milhoes > 0:
        txt_m = _centena_para_extenso(milhoes)
        partes.append(f"{txt_m} " + ('milhão' if milhoes == 1 else 'milhões'))
    if milhares > 0:
        if milhares == 1 and not milhoes:
            partes.append('mil')
        else:
            txt_mil = _centena_para_extenso(milhares)
            partes.append(f"{txt_mil} mil")
    if unidades > 0:
        txt_u = _centena_para_extenso(unidades)
        partes.append(txt_u)

    if len(partes) == 1:
        return partes[0]
    return ', '.join(partes[:-1]) + (', ' if len(partes) > 1 else '') + partes[-1]


def valor_monetario_extenso(valor: float) -> str:
    """
    Retorna o valor formatado por extenso em Reais e Centavos.
    Exemplo: 40549.21 -> 'quarenta mil, quinhentos e quarenta e nove reais e vinte e um centavos'
    """
    reais = int(round(valor, 2))
    centavos = int(round((round(valor, 2) - reais) * 100))
    if centavos < 0:
        centavos = 0
    if centavos == 100:
        reais += 1
        centavos = 0

    partes = []
    if reais > 0:
        txt_reais = numero_para_extenso(reais)
        moeda = 'real' if reais == 1 else 'reais'
        partes.append(f"{txt_reais} {moeda}")
    elif centavos == 0:
        return "zero reais"

    if centavos > 0:
        txt_cents = _centena_para_extenso(centavos)
        moeda_c = 'centavo' if centavos == 1 else 'centavos'
        partes.append(f"{txt_cents} {moeda_c}")

    return ' e '.join(partes)


def _parse_data(data_val: Any) -> Optional[datetime]:
    """Converte valores de data em objeto datetime."""
    if not data_val:
        return None
    d_str = str(data_val).strip()
    formatos = [
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%Y-%m-%d",
        "%d/%m/%Y %H:%M:%S",
        "%d/%m/%Y %H:%M",
        "%d/%m/%Y"
    ]
    for fmt in formatos:
        try:
            return datetime.strptime(d_str[:19], fmt)
        except ValueError:
            continue
    return None


def obter_dados_periodo_atestado(abastecimentos: List[Dict[str, Any]]) -> Dict[str, Any]:
    """Calcula os dados dinâmicos do período para o Atestado."""
    datas_dt = [_parse_data(a.get('data_abastecimento')) for a in abastecimentos]
    datas_validas = sorted([d for d in datas_dt if d is not None])

    if not datas_validas:
        alvo = datetime.now()
    else:
        alvo = datas_validas[len(datas_validas) // 2]

    # Regra oficial das quinzenas:
    # 2ª quinzena: dias 10 a 24
    if 10 <= alvo.day <= 24:
        quinzena = "2ª"
        mes_ref = MESES_PT[alvo.month]
        ano_ref = alvo.year
        texto_periodo = f"10 a 24 de {mes_ref} de {ano_ref}"
    elif alvo.day >= 25:
        # 1ª quinzena do próximo mês
        quinzena = "1ª"
        m_prox = alvo.month + 1 if alvo.month < 12 else 1
        a_prox = alvo.year if alvo.month < 12 else alvo.year + 1
        mes_ini = MESES_PT[alvo.month]
        mes_fim = MESES_PT[m_prox]
        if alvo.month == m_prox:
            texto_periodo = f"25 a 09 de {mes_fim} de {a_prox}"
        else:
            texto_periodo = f"25 de {mes_ini} a 09 de {mes_fim} de {a_prox}"
        mes_ref = mes_fim
        ano_ref = a_prox
    else:
        # 1ª quinzena do mês atual (iniciou dia 25 do mês anterior)
        quinzena = "1ª"
        m_ant = alvo.month - 1 if alvo.month > 1 else 12
        a_ant = alvo.year if alvo.month > 1 else alvo.year - 1
        mes_ini = MESES_PT[m_ant]
        mes_fim = MESES_PT[alvo.month]
        if m_ant == alvo.month:
            texto_periodo = f"25 a 09 de {mes_fim} de {alvo.year}"
        else:
            texto_periodo = f"25 de {mes_ini} a 09 de {mes_fim} de {alvo.year}"
        mes_ref = mes_fim
        ano_ref = alvo.year

    nome_arquivo = f"Atestado de Combustíveis - {quinzena} quinzena de {mes_ref}.pdf"

    return {
        'quinzena': quinzena,
        'mes_ref': mes_ref,
        'ano_ref': ano_ref,
        'texto_periodo': texto_periodo,
        'nome_arquivo': nome_arquivo
    }


def gerar_texto_padrao_atestado(
    abastecimentos: Optional[List[Dict[str, Any]]] = None,
    valor_total: Optional[float] = None
) -> str:
    """
    Gera o texto padrão oficial do Atestado de Combustíveis com os valores, período
    e quinzena calculados dinamicamente.
    """
    if abastecimentos is None:
        try:
            import database
            abastecimentos = database.obter_todos_abastecimentos(limit=50000)
        except Exception:
            abastecimentos = []

    if valor_total is None or valor_total <= 0:
        if abastecimentos:
            valor_total = sum(float(a.get('valor_total') or a.get('valor_abastecido') or 0.0) for a in abastecimentos)
        else:
            valor_total = 0.0

    info_periodo = obter_dados_periodo_atestado(abastecimentos or [])
    texto_periodo = info_periodo['texto_periodo']
    quinzena_texto = info_periodo['quinzena']

    val_num_str = f"{valor_total:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    val_extenso_str = valor_monetario_extenso(valor_total)

    return (
        f"Atesto, para devidos fins, os gastos de combustíveis, lubrificantes, "
        f"filtros e lavagem da frota do 46º Batalhão de Polícia Militar, perfazendo o valor total de "
        f"<b>R$ {val_num_str}</b> ({val_extenso_str}), referente ao período de "
        f"<b>{texto_periodo}</b>, correspondente a <b>{quinzena_texto} quinzena</b>, "
        f"de acordo com os dados extraídos do Sistema Ticket Log, estando as empresas em "
        f"conformidade com o que preconiza o Termo de Contrato constante no Processo "
        f"Administrativo nº 05365-24.00/15-0, Pregão Eletrônico nº 100/CELIC/2016, "
        f"bem como a NI 4.7/2018."
    )


def gerar_atestado_pdf(
    pdf_path: str | Path,
    abastecimentos: List[Dict[str, Any]],
    valor_total: Optional[float] = None,
    dados_assinatura: Optional[Dict[str, str]] = None,
    texto_personalizado: Optional[str] = None
) -> None:
    """
    Gera o PDF do ATESTADO DE COMBUSTÍVEIS formatado conforme o padrão do 46º BPM.
    Se texto_personalizado for fornecido, utiliza-o no corpo do documento;
    caso contrário, utiliza o texto padrão oficial com cálculos automáticos.
    """
    pdf_path = Path(pdf_path)
    pdf_path.parent.mkdir(parents=True, exist_ok=True)

    # Margens do modelo ODT: Top 10mm, Left 30mm, Right 15mm, Bottom 20mm (~2cm)
    doc = SimpleDocTemplate(
        str(pdf_path),
        pagesize=A4,
        leftMargin=30 * mm,
        rightMargin=15 * mm,
        topMargin=10 * mm,
        bottomMargin=20 * mm
    )

    largura_util = A4[0] - (45 * mm)  # 210mm - 45mm = 165mm

    # Diretório dos assets
    base_dir = Path(__file__).resolve().parent.parent
    img_dir = base_dir / "static" / "img"
    img_brasao_rs = img_dir / "brasao_rs_armas.png"
    img_brasao_bm = img_dir / "brasao_bm_atestado.png"
    img_rodape = img_dir / "rodape_redes_sociais.png"

    # Cálculo dos valores e períodos
    if valor_total is None or valor_total <= 0:
        valor_total = sum(float(a.get('valor_total') or a.get('valor_abastecido') or 0.0) for a in abastecimentos)

    info_periodo = obter_dados_periodo_atestado(abastecimentos)
    texto_periodo = info_periodo['texto_periodo']
    quinzena_texto = info_periodo['quinzena']

    # Formatação monetária
    val_num_str = f"{valor_total:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    val_extenso_str = valor_monetario_extenso(valor_total)

    # Data de geração atualizada dinamicamente
    hoje = datetime.now()
    data_geracao_str = f"Quartel em Bagé, RS, {hoje.day:02d} de {MESES_PT[hoje.month]} de {hoje.year}."

    # Configuração do militar gestor de frota
    cfg = get_assinatura_config()
    if dados_assinatura:
        nome_gestor = (dados_assinatura.get('nome') or cfg['nome']).strip()
        grad_gestor = (dados_assinatura.get('graduacao') or cfg['graduacao']).strip()
    else:
        nome_gestor = cfg['nome']
        grad_gestor = cfg['graduacao']

    # Estilos de texto
    style_center_bold = ParagraphStyle(
        'HeaderCenterBold',
        fontName='Times-Bold',
        fontSize=12,
        leading=14,
        alignment=1, # Centralizado
        textColor=colors.black
    )

    style_titulo = ParagraphStyle(
        'TituloAtestado',
        fontName='Times-Bold',
        fontSize=12,
        leading=14,
        alignment=1, # Centralizado
        textColor=colors.black,
        spaceAfter=0
    )

    style_corpo = ParagraphStyle(
        'CorpoTexto',
        fontName='Times-Roman',
        fontSize=12,
        leading=18, # 1.5 line spacing
        alignment=4, # Justificado
        firstLineIndent=20 * mm,
        textColor=colors.black
    )

    style_data = ParagraphStyle(
        'DataGeracao',
        fontName='Times-Roman',
        fontSize=12,
        leading=15,
        alignment=0, # À esquerda
        textColor=colors.black
    )

    style_sig_cargo = ParagraphStyle(
        'SigCargo',
        fontName='Helvetica',
        fontSize=8,
        leading=10,
        alignment=0,
        textColor=colors.black
    )

    style_sig_nome = ParagraphStyle(
        'SigNome',
        fontName='Helvetica-Bold',
        fontSize=10,
        leading=12,
        alignment=0,
        textColor=colors.black
    )

    style_sig_obs = ParagraphStyle(
        'SigObs',
        fontName='Helvetica-Oblique',
        fontSize=9,
        leading=11,
        alignment=0,
        textColor=colors.HexColor("#333333")
    )

    elementos = []

    # 1. Brasão das Armas do RS (Centralizado no Topo)
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

    # 3. Título ATESTADO
    elementos.append(Paragraph("ATESTADO", style_titulo))
    elementos.append(Spacer(1, 10 * mm))

    # 4. Texto do Atestado (Editável ou Padrão)
    if texto_personalizado and str(texto_personalizado).strip():
        texto_principal = str(texto_personalizado).strip()
    else:
        texto_principal = gerar_texto_padrao_atestado(abastecimentos, valor_total)

    import html
    # Separa por múltiplos parágrafos se existirem
    paragrafos = [p.strip() for p in texto_principal.replace('\r\n', '\n').split('\n\n') if p.strip()]
    if not paragrafos:
        paragrafos = [texto_principal.strip()]

    for i, p_texto in enumerate(paragrafos):
        # Substitui quebras simples por <br/>
        p_html = p_texto.replace('\n', '<br/>')
        try:
            p_elem = Paragraph(p_html, style_corpo)
        except Exception:
            p_elem = Paragraph(html.escape(p_texto).replace('\n', '<br/>'), style_corpo)
        elementos.append(p_elem)
        if i < len(paragrafos) - 1:
            elementos.append(Spacer(1, 4 * mm))

    elementos.append(Spacer(1, 8 * mm))

    # 5. Data de Geração Atualizada
    elementos.append(Paragraph(data_geracao_str, style_data))
    
    # 6. Bloco de Assinaturas Conforme o Modelo Selecionado
    from services.report_config import get_atestado_config
    cfg_atestado = get_atestado_config()
    if dados_assinatura:
        for k, v in dados_assinatura.items():
            if v is not None and str(v).strip():
                cfg_atestado[k] = str(v).strip()

    modo = cfg_atestado.get('modo', 'impedimento').lower()
    titular_nome = cfg_atestado.get('titular_nome', 'KARLA DE MOURA INCERTI')
    titular_cargo = cfg_atestado.get('titular_cargo', 'Ten Cel PM – Cmt 46º BPM')
    respondendo_nome = cfg_atestado.get('respondendo_nome', 'PATRIQUE ROLIM MARQUES')
    respondendo_cargo = cfg_atestado.get('respondendo_cargo', 'Maj PM – Rsp. p/ Cmt 46º BPM')
    subcomandante_nome = cfg_atestado.get('subcomandante_nome', 'PATRIQUE ROLIM MARQUES')
    subcomandante_cargo = cfg_atestado.get('subcomandante_cargo', 'Maj PM – SCmt 46º BPM')
    gestor_nome = cfg_atestado.get('gestor_nome', cfg.get('nome', 'RÓGER SANTANA MANZZINI'))
    gestor_cargo = cfg_atestado.get('gestor_cargo', f"{cfg.get('graduacao', 'SD PM')} – Gestor de frota")

    # Ajuste de espaçamento que empurra o bloco de assinaturas e rodapé para o final da página (~2cm do fim)
    spacer_antes_ass = 42 * mm if modo == 'impedimento' else 72 * mm
    elementos.append(Spacer(1, spacer_antes_ass))

    col_w = largura_util / 2.0  # ~82.5mm cada coluna

    if modo == 'titular':
        # Modelo 1: Comandante Titular + Gestor de Frota
        celula_esq = [
            Paragraph(titular_nome, style_sig_nome),
            Paragraph(titular_cargo, style_sig_cargo)
        ]
        celula_dir = [
            Paragraph(gestor_nome, style_sig_nome),
            Paragraph(gestor_cargo, style_sig_cargo)
        ]
        tab_assinaturas = Table([
            [celula_esq, celula_dir]
        ], colWidths=[col_w, col_w])

    elif modo == 'respondendo':
        # Modelo 2: Respondendo pelo Comando + Gestor de Frota
        celula_esq = [
            Paragraph(respondendo_nome, style_sig_nome),
            Paragraph(respondendo_cargo, style_sig_cargo)
        ]
        celula_dir = [
            Paragraph(gestor_nome, style_sig_nome),
            Paragraph(gestor_cargo, style_sig_cargo)
        ]
        tab_assinaturas = Table([
            [celula_esq, celula_dir]
        ], colWidths=[col_w, col_w])

    else:
        # Modelo 3: No impedimento (uma assinatura abaixo de outra, com espaço para assinatura GOV)
        celula_esq_sup = [
            Paragraph("No impedimento", style_sig_obs),
            Spacer(1, 1.5 * mm),
            Paragraph(titular_nome, style_sig_nome),
            Paragraph(titular_cargo, style_sig_cargo),
            Spacer(1, 24 * mm),  # Espaço para aposição da assinatura eletrônica GOV.BR
            Paragraph(subcomandante_nome, style_sig_nome),
            Paragraph(subcomandante_cargo, style_sig_cargo)
        ]
        celula_dir_sup = [
            Spacer(1, 4 * mm),  # Alinha visualmente com o primeiro signatário
            Paragraph(gestor_nome, style_sig_nome),
            Paragraph(gestor_cargo, style_sig_cargo)
        ]
        tab_assinaturas = Table([
            [celula_esq_sup, celula_dir_sup]
        ], colWidths=[col_w, col_w])

    tab_assinaturas.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (-1, -1), 0),
        ('RIGHTPADDING', (0, 0), (-1, -1), 0),
        ('TOPPADDING', (0, 0), (-1, -1), 0),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
    ]))

    # 7. Rodapé Oficial: Endereço à esquerda e Brasão da Brigada Militar à direita
    style_rodape_end = ParagraphStyle(
        'RodapeEndereco',
        fontName='Helvetica',
        fontSize=8.5,
        leading=11,
        textColor=colors.HexColor("#444444")
    )

    bloco_endereco = [
        Paragraph("Rua Cândido Portinari, 310, Floresta – Bagé/RS", style_rodape_end),
        Paragraph("Fone: (53) 3240 3400", style_rodape_end),
        Paragraph("E-mail: 46bpm-p4@bm.rs.gov.br", style_rodape_end),
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

    # Bloco final contendo assinaturas, espaço de tamanho de fonte 9, rodapé e banner institucional
    bloco_final = [
        tab_assinaturas,
        Spacer(1, 9),  # Espaço de tamanho de fonte 9 do endereço e contatos
        tab_rodape,
        Spacer(1, 2 * mm)
    ]

    if img_rodape.exists():
        img_redes = RLImage(str(img_rodape), width=largura_util, height=14.5 * mm)
        img_redes.hAlign = 'CENTER'
        bloco_final.append(img_redes)

    elementos.append(KeepTogether(bloco_final))

    # Constrói o PDF em página única
    doc.build(elementos)
