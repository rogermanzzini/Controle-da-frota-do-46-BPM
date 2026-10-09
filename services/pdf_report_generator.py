"""
PDF Report Generator for the 46º BPM consumption control.

Gera a Planilha de Controle de Consumo de Combustível e Lubrificantes em PDF
reproduzindo exatamente o padrão oficial utilizado pelo P4 do 46º BPM:
- Sem brasão
- Cada viatura selecionada em sua própria página (PageBreak entre viaturas)
- Cabeçalho superior militar em 2 colunas com fonte Times-Bold 12pt:
  Esquerda (alinhado à esquerda):
    ESTADO DO RIO GRANDE DO SUL
    SECRETARIA DA SEGURANÇA PÚBLICA
    BRIGADA MILITAR
    CRPM/FO - 46º BPM
  Direita (alinhado à direita):
    BAGÉ, RS
    Em [DATA MILITAR]. (ex: Em 24SET26. ou Em 09SET26.)
- Título do documento centralizado (Times-Bold 12pt):
  Planilha de Controle de Consumo de Combustível e Lubrificantes
- Linha de período centralizada (Times-Bold 12pt):
  De: 10/09/2026 a 24/09/2026      Mês: SETEMBRO de 2026 (2ª quinzena: 10 a 24)
  ou
  De: 25/08/2026 a 09/09/2026      Mês: SETEMBRO de 2026 (1ª quinzena: 25 a 09)
- Tabela com 9 colunas e bordas pretas finas (Times-Bold no cabeçalho, Times-Roman nos dados):
  OPM | Vtr | Data do Abastecimento | Posto Combustível | Motorista | Quantidade Litros | Tipo de Combustível | Valor Unitário | Valor Abastecido
- Linha de Total com valor somado
- Bloco de assinatura oficial:
  RÓGER SANTANA MANZZINI – SD PM
  Auxiliar do P4/46º BPM
"""

from __future__ import annotations

import os
from pathlib import Path
from datetime import datetime
from collections import defaultdict
from typing import List, Dict, Any, Optional, Tuple

from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    PageBreak,
    KeepTogether
)

from services.report_config import get_signature, get_assinatura_config


MESES_PT = {
    1: 'JANEIRO', 2: 'FEVEREIRO', 3: 'MARÇO', 4: 'ABRIL',
    5: 'MAIO', 6: 'JUNHO', 7: 'JULHO', 8: 'AGOSTO',
    9: 'SETEMBRO', 10: 'OUTUBRO', 11: 'NOVEMBRO', 12: 'DEZEMBRO'
}

MESES_MIL = {
    1: 'JAN', 2: 'FEV', 3: 'MAR', 4: 'ABR',
    5: 'MAI', 6: 'JUN', 7: 'JUL', 8: 'AGO',
    9: 'SET', 10: 'OUT', 11: 'NOV', 12: 'DEZ'
}


def _formatar_decimal(valor: Any, decimais: int = 2) -> str:
    """Formata valor com vírgula decimal sem símbolo R$, idêntico ao modelo Word."""
    if valor is None or valor == "":
        return ""
    try:
        val_float = float(valor)
        if val_float == 0.0:
            return ""
        fmt = f"{{:,.{decimais}f}}"
        return fmt.format(val_float).replace(",", "X").replace(".", ",").replace("X", ".")
    except (ValueError, TypeError):
        return str(valor)


def _formatar_unitario(valor: Any) -> str:
    """Formata valor unitário (ex: 7 se inteiro, ou 5,75)."""
    if valor is None or valor == "":
        return ""
    try:
        val_float = float(valor)
        if val_float == 0.0:
            return ""
        if val_float.is_integer():
            return str(int(val_float))
        return f"{val_float:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except (ValueError, TypeError):
        return str(valor)


def _formatar_total_soma(valor: Any) -> str:
    """Formata o total somado para a linha Total."""
    if valor is None:
        return "0,00"
    try:
        val_float = float(valor)
        return f"{val_float:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except (ValueError, TypeError):
        return str(valor)


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


def _formatar_data_abastecimento(data_val: Any) -> str:
    """Formata a data para DD/MM/AAAA HH:MM ou DD/MM/AAAA."""
    dt = _parse_data(data_val)
    if dt:
        if dt.hour != 0 or dt.minute != 0:
            return dt.strftime("%d/%m/%Y %H:%M")
        return dt.strftime("%d/%m/%Y")
    d_str = str(data_val or '').strip()
    if len(d_str) >= 10 and d_str[4] == "-" and d_str[7] == "-":
        partes = d_str[:10].split("-")
        return f"{partes[2]}/{partes[1]}/{partes[0]}"
    return d_str[:16] if d_str else "-"


def obter_info_quinzena(abastecimentos: List[Dict[str, Any]]) -> Dict[str, Any]:
    """
    Calcula todas as informações da quinzena oficial militar com base nos abastecimentos:
    - 1ª quinzena: dias 25 (mês anterior) a 09 (mês de referência) -> Em 09[MES][ANO].
    - 2ª quinzena: dias 10 a 24 (mês de referência) -> Em 24[MES][ANO].

    Retorna dicionário contendo:
    - quinzena: '1ª' ou '2ª'
    - mes_nome: nome do mês em minúsculo (ex: 'setembro')
    - periodo_str: texto oficial do período
    - header_dir: data militar à direita
    - nome_arquivo: nome formatado conforme solicitado
      (ex: 'Planilha de Controle de Combustivel e Lubrificante - 2ª quinzena de setembro.pdf')
    """
    datas_dt = [_parse_data(a.get('data_abastecimento')) for a in abastecimentos]
    datas_validas = sorted([d for d in datas_dt if d is not None])

    if not datas_validas:
        alvo = datetime.now()
    else:
        # Usa data representativa do lote
        alvo = datas_validas[len(datas_validas) // 2]

    # Regra oficial das quinzenas:
    # 2ª quinzena: dia 10 a dia 24 do mês
    if 10 <= alvo.day <= 24:
        quinzena = "2ª"
        d_ini = f"10/{alvo.month:02d}/{alvo.year}"
        d_fim = f"24/{alvo.month:02d}/{alvo.year}"
        mes_ref = MESES_PT[alvo.month]
        ano_ref = alvo.year
        dt_mil = f"Em 24{MESES_MIL[alvo.month]}{alvo.year % 100}."
    elif alvo.day >= 25:
        # 1ª quinzena do próximo mês
        quinzena = "1ª"
        m_prox = alvo.month + 1 if alvo.month < 12 else 1
        a_prox = alvo.year if alvo.month < 12 else alvo.year + 1
        d_ini = f"25/{alvo.month:02d}/{alvo.year}"
        d_fim = f"09/{m_prox:02d}/{a_prox}"
        mes_ref = MESES_PT[m_prox]
        ano_ref = a_prox
        dt_mil = f"Em 09{MESES_MIL[m_prox]}{a_prox % 100}."
    else:
        # Dia 1 a 9: 1ª quinzena do mês atual (iniciou dia 25 do mês anterior)
        quinzena = "1ª"
        m_ant = alvo.month - 1 if alvo.month > 1 else 12
        a_ant = alvo.year if alvo.month > 1 else alvo.year - 1
        d_ini = f"25/{m_ant:02d}/{a_ant}"
        d_fim = f"09/{alvo.month:02d}/{alvo.year}"
        mes_ref = MESES_PT[alvo.month]
        ano_ref = alvo.year
        dt_mil = f"Em 09{MESES_MIL[alvo.month]}{alvo.year % 100}."

    periodo_str = f"De: {d_ini} a {d_fim}      Mês: {mes_ref} de {ano_ref}"
    header_dir = f"BAGÉ, RS<br/>{dt_mil}"
    nome_arquivo = f"Planilha de Controle de Combustivel e Lubrificante - {quinzena} quinzena de {mes_ref.lower()}.pdf"

    return {
        'quinzena': quinzena,
        'mes_nome': mes_ref.lower(),
        'periodo_str': periodo_str,
        'header_dir': header_dir,
        'nome_arquivo': nome_arquivo
    }


def calcular_dados_quinzena(abastecimentos: List[Dict[str, Any]]) -> Tuple[str, str]:
    """Retorna: (texto_periodo, texto_data_militar) para o cabeçalho."""
    info = obter_info_quinzena(abastecimentos)
    return info['periodo_str'], info['header_dir']


def obter_nome_arquivo_relatorio(abastecimentos: List[Dict[str, Any]]) -> str:
    """
    Retorna o nome do arquivo conforme exigido:
    'Planilha de Controle de Combustivel e Lubrificante - 2ª quinzena de setembro.pdf' (ou 1ª quinzena)
    """
    info = obter_info_quinzena(abastecimentos)
    return info['nome_arquivo']


def gerar_pdf(
    pdf_path: str | Path,
    abastecimentos: List[Dict[str, Any]],
    veiculos_selecionados: List[str] | None = None,
    dados_assinatura: Dict[str, str] | None = None
) -> None:
    """
    Cria a Planilha de Controle de Consumo de Combustível e Lubrificantes em PDF
    idêntica ao modelo do Word/Excel do P4 / 46º BPM.
    """
    pdf_path = Path(pdf_path)
    pdf_path.parent.mkdir(parents=True, exist_ok=True)

    # Documento em A4 Paisagem (Landscape) com margens iguais ao modelo Word (20mm laterais, 6mm verticais)
    doc = SimpleDocTemplate(
        str(pdf_path),
        pagesize=landscape(A4),
        leftMargin=20 * mm,
        rightMargin=20 * mm,
        topMargin=8 * mm,
        bottomMargin=8 * mm
    )

    # Estilos usando fonte Times-Bold / Times-Roman (tamanho 12pt para o cabeçalho e títulos como no Word)
    style_header_esq = ParagraphStyle(
        'HeaderEsq',
        fontName='Times-Bold',
        fontSize=12,
        leading=14.5,
        alignment=0, # Alinhado à esquerda
        textColor=colors.black
    )

    style_header_dir = ParagraphStyle(
        'HeaderDir',
        fontName='Times-Bold',
        fontSize=12,
        leading=14.5,
        alignment=2, # Alinhado à direita
        textColor=colors.black
    )

    style_titulo = ParagraphStyle(
        'TituloDoc',
        fontName='Times-Bold',
        fontSize=12,
        leading=15,
        alignment=1, # Centralizado
        textColor=colors.black
    )

    style_periodo = ParagraphStyle(
        'PeriodoDoc',
        fontName='Times-Bold',
        fontSize=12,
        leading=15,
        alignment=1, # Centralizado
        textColor=colors.black
    )

    style_th = ParagraphStyle(
        'TableHead',
        fontName='Times-Bold',
        fontSize=8,
        leading=10,
        alignment=1, # Centralizado
        textColor=colors.black
    )

    style_td_center = ParagraphStyle(
        'TableColCenter',
        fontName='Times-Roman',
        fontSize=7.5,
        leading=9.5,
        alignment=1,
        textColor=colors.black
    )

    style_td_left = ParagraphStyle(
        'TableColLeft',
        fontName='Times-Roman',
        fontSize=7.5,
        leading=9.5,
        alignment=0,
        textColor=colors.black
    )

    style_td_right = ParagraphStyle(
        'TableColRight',
        fontName='Times-Roman',
        fontSize=7.5,
        leading=9.5,
        alignment=2,
        textColor=colors.black
    )

    style_total_label = ParagraphStyle(
        'TableTotalLabel',
        fontName='Times-Bold',
        fontSize=8.5,
        leading=10.5,
        alignment=0, # Alinhado à esquerda
        textColor=colors.black
    )

    style_total_val = ParagraphStyle(
        'TableTotalVal',
        fontName='Times-Bold',
        fontSize=8.5,
        leading=10.5,
        alignment=2,
        textColor=colors.black
    )

    style_ass_nome = ParagraphStyle(
        'AssinaturaNome',
        fontName='Times-Bold',
        fontSize=10,
        leading=12,
        alignment=1,
        textColor=colors.black
    )

    style_ass_cargo = ParagraphStyle(
        'AssinaturaCargo',
        fontName='Times-Roman',
        fontSize=9.5,
        leading=11.5,
        alignment=1,
        textColor=colors.black
    )

    # Calcula a quinzena oficial de todo o lote
    periodo_lote_str, header_dir_lote = calcular_dados_quinzena(abastecimentos)

    # Agrupar abastecimentos por viatura
    grupos_por_veiculo = defaultdict(list)
    for a in abastecimentos:
        chave = str(a.get('prefixo') or a.get('placa') or '').strip().upper()
        if chave:
            grupos_por_veiculo[chave].append(a)

    # Identificar a ordem das viaturas a gerar (somente as selecionadas)
    lista_chaves = []
    if veiculos_selecionados:
        for v in veiculos_selecionados:
            v_norm = str(v).strip().upper()
            if not v_norm:
                continue
            if v_norm in grupos_por_veiculo:
                if v_norm not in lista_chaves:
                    lista_chaves.append(v_norm)
            else:
                achou = False
                for k, itens in grupos_por_veiculo.items():
                    for item in itens:
                        p_pref = str(item.get('prefixo') or '').strip().upper()
                        p_plc = str(item.get('placa') or '').strip().upper()
                        if v_norm == p_pref or v_norm == p_plc:
                            if k not in lista_chaves:
                                lista_chaves.append(k)
                            achou = True
                            break
                    if achou:
                        break
                if not achou and v_norm not in lista_chaves:
                    lista_chaves.append(v_norm)
    else:
        lista_chaves = sorted(list(grupos_por_veiculo.keys()))

    if not lista_chaves and grupos_por_veiculo:
        lista_chaves = sorted(list(grupos_por_veiculo.keys()))

    if not lista_chaves:
        lista_chaves = ["46º BPM"]

    elementos = []

    # Larguras das 9 colunas preenchendo exatamente os 257 mm úteis (728.5 pt)
    col_widths = [
        17 * mm,  # 1. OPM
        17 * mm,  # 2. Vtr
        28 * mm,  # 3. Data do Abastecimento
        52 * mm,  # 4. Posto Combustível
        48 * mm,  # 5. Motorista
        22 * mm,  # 6. Quantidade Litros
        27 * mm,  # 7. Tipo de Combustível
        21 * mm,  # 8. Valor Unitário
        25 * mm   # 9. Valor Abastecido
    ]

    total_veiculos = len(lista_chaves)

    for idx_vtr, vtr_chave in enumerate(lista_chaves):
        itens_vtr = grupos_por_veiculo.get(vtr_chave, [])
        itens_vtr.sort(key=lambda x: str(x.get('data_abastecimento') or ''))

        primeiro_item = itens_vtr[0] if itens_vtr else {}
        opm_vtr = str(primeiro_item.get('fracao') or '46º BPM').strip()
        if not opm_vtr or opm_vtr.upper() == 'NONE':
            opm_vtr = '46º BPM'

        pref_item = str(primeiro_item.get('prefixo') or '').strip()
        plc_item = str(primeiro_item.get('placa') or '').strip()
        viatura_nome = pref_item if pref_item else (plc_item if plc_item else vtr_chave)

        # -------------------------------------------------------------
        # 1. CABEÇALHO MILITAR SUPERIOR (TABLE 1 DO WORD - 12PT BOLD)
        # -------------------------------------------------------------
        texto_esq = (
            "ESTADO DO RIO GRANDE DO SUL<br/>"
            "SECRETARIA DA SEGURANÇA PÚBLICA<br/>"
            "BRIGADA MILITAR<br/>"
            "CRPM/FO - 46º BPM"
        )
        tabela_header_topo = Table(
            [[Paragraph(texto_esq, style_header_esq), Paragraph(header_dir_lote, style_header_dir)]],
            colWidths=[130 * mm, 127 * mm]
        )
        tabela_header_topo.setStyle(TableStyle([
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('LEFTPADDING', (0, 0), (-1, -1), 0),
            ('RIGHTPADDING', (0, 0), (-1, -1), 0),
            ('TOPPADDING', (0, 0), (-1, -1), 0),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
        ]))
        elementos.append(tabela_header_topo)
        elementos.append(Spacer(1, 4 * mm))

        # -------------------------------------------------------------
        # 2. TÍTULO DO DOCUMENTO (TIMES-BOLD 12PT)
        # -------------------------------------------------------------
        elementos.append(Paragraph("Planilha de Controle de Consumo de Combustível e Lubrificantes", style_titulo))
        elementos.append(Spacer(1, 2 * mm))

        # -------------------------------------------------------------
        # 3. LINHA DE PERÍODO / MÊS (TIMES-BOLD 12PT)
        # -------------------------------------------------------------
        elementos.append(Paragraph(periodo_lote_str, style_periodo))
        elementos.append(Spacer(1, 4 * mm))

        # -------------------------------------------------------------
        # 4. TABELA DE 9 COLUNAS (TABLE 2 DO WORD)
        # -------------------------------------------------------------
        header_row = [
            Paragraph("OPM", style_th),
            Paragraph("Vtr", style_th),
            Paragraph("Data do Abastecimento", style_th),
            Paragraph("Posto Combustível", style_th),
            Paragraph("Motorista", style_th),
            Paragraph("Quantidade Litros", style_th),
            Paragraph("Tipo de Combustível", style_th),
            Paragraph("Valor Unitário", style_th),
            Paragraph("Valor Abastecido", style_th),
        ]

        tabela_linhas = [header_row]

        soma_valor = 0.0

        if itens_vtr:
            for item in itens_vtr:
                opm = str(item.get('fracao') or opm_vtr or '46º BPM').strip()
                pref = str(item.get('prefixo') or '').strip()
                plc = str(item.get('placa') or '').strip()
                vtr_exib = pref if pref else plc

                data_str = _formatar_data_abastecimento(item.get('data_abastecimento'))
                posto_str = str(item.get('posto') or '-').strip()
                motorista_str = str(item.get('motorista') or '-').strip()

                litros_val = item.get('litros') or item.get('quantidade_litros')
                combustivel_str = str(item.get('tipo_combustivel') or '-').strip()

                vu_val = item.get('valor_unitario')
                vt_val = float(item.get('valor_total') or item.get('valor_abastecido') or 0.0)
                soma_valor += vt_val

                linha = [
                    Paragraph(opm, style_td_center),
                    Paragraph(vtr_exib, style_td_center),
                    Paragraph(data_str, style_td_center),
                    Paragraph(posto_str, style_td_left),
                    Paragraph(motorista_str, style_td_left),
                    Paragraph(_formatar_decimal(litros_val, 2), style_td_right),
                    Paragraph(combustivel_str, style_td_center),
                    Paragraph(_formatar_unitario(vu_val), style_td_right),
                    Paragraph(_formatar_decimal(vt_val, 2), style_td_right),
                ]
                tabela_linhas.append(linha)
        else:
            for _ in range(4):
                tabela_linhas.append([
                    Paragraph(opm_vtr, style_td_center),
                    Paragraph(viatura_nome, style_td_center),
                    "", "", "", "", "", "", ""
                ])

        # Linha de Total idêntica ao modelo do Word
        linha_total = [
            Paragraph("Total", style_total_label),
            "", "", "", "", "", "", "",
            Paragraph(_formatar_total_soma(soma_valor), style_total_val)
        ]
        tabela_linhas.append(linha_total)

        idx_total = len(tabela_linhas) - 1

        tabela_vtr = Table(tabela_linhas, colWidths=col_widths, repeatRows=1)
        tabela_vtr.setStyle(TableStyle([
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('TOPPADDING', (0, 0), (-1, -1), 3),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
            ('LEFTPADDING', (0, 0), (-1, -1), 3),
            ('RIGHTPADDING', (0, 0), (-1, -1), 3),
            ('SPAN', (0, idx_total), (7, idx_total)),
        ]))

        elementos.append(tabela_vtr)
        elementos.append(Spacer(1, 8 * mm))

        # -------------------------------------------------------------
        # 5. BLOCO DE ASSINATURA OFICIAL (EDITÁVEL)
        # -------------------------------------------------------------
        cfg = get_assinatura_config()
        if dados_assinatura:
            nome_sig = (dados_assinatura.get('nome') or cfg['nome']).strip()
            grad_sig = (dados_assinatura.get('graduacao') or cfg['graduacao']).strip()
            cargo_sig = (dados_assinatura.get('cargo') or cfg['cargo']).strip()
        else:
            nome_sig = cfg['nome']
            grad_sig = cfg['graduacao']
            cargo_sig = cfg['cargo']

        if grad_sig:
            linha_nome = f"{nome_sig} – {grad_sig}"
        else:
            linha_nome = nome_sig

        tabela_assinatura = Table([
            [Paragraph(linha_nome, style_ass_nome)],
            [Paragraph(cargo_sig, style_ass_cargo)]
        ], colWidths=[120 * mm])

        tabela_assinatura.setStyle(TableStyle([
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('TOPPADDING', (0, 0), (-1, -1), 1),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 1),
        ]))

        elementos.append(KeepTogether([tabela_assinatura]))

        # Se houver mais viaturas, quebra para uma nova página
        if idx_vtr < total_veiculos - 1:
            elementos.append(PageBreak())

    # Compila o PDF final
    doc.build(elementos)
