"""
services/excel_processor.py - Motor de Processamento de Planilhas com Pandas e openpyxl
Cruza a planilha geral de abastecimentos do RS com a lista de viaturas do 46º BPM (Bagé),
calculando odômetros, consumos (km/l), custos e gerando relatórios formatados em Excel.
"""

import os
import re
import unicodedata
import gc
from datetime import datetime
from typing import Tuple, Dict, Any, List, Optional
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


def remover_acentos(texto: str) -> str:
    """Remove acentuações e caracteres especiais para comparação flexível."""
    if not isinstance(texto, str):
        texto = str(texto) if texto is not None else ""
    nfkd = unicodedata.normalize('NFKD', texto)
    return u"".join([c for c in nfkd if not unicodedata.combining(c)]).lower().strip()


def normalizar_placa(placa_val: Any) -> str:
    """Higieniza placa: maiúsculas, sem hífens, pontos ou espaços (suporta Mercosul)."""
    if pd.isna(placa_val) or placa_val is None:
        return ""
    placa_str = str(placa_val).strip().upper()
    placa_limpa = re.sub(r'[^A-Z0-9]', '', placa_str)
    return placa_limpa


def normalizar_prefixo(pref_val: Any) -> str:
    """Higieniza prefixo de viatura (ex: 'BM 12450', '12450.0' -> '12450')."""
    if pd.isna(pref_val) or pref_val is None:
        return ""
    if isinstance(pref_val, float):
        if pref_val.is_integer():
            return str(int(pref_val))
    p_str = str(pref_val).strip()
    if p_str.endswith('.0'):
        p_str = p_str[:-2]
    # Se contiver 'BM-' ou 'BM ' normaliza
    p_str = re.sub(r'\s+', ' ', p_str)
    return p_str.upper()


def limpar_numero(val: Any) -> float:
    """Converte números com vírgula ou ponto para float com segurança."""
    if pd.isna(val) or val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip().replace('R$', '').replace('r$', '').replace(' ', '')
    if not s:
        return 0.0
    # Caso 1.234,56
    if ',' in s and '.' in s:
        s = s.replace('.', '').replace(',', '.')
    elif ',' in s:
        s = s.replace(',', '.')
    try:
        return float(s)
    except ValueError:
        return 0.0


# Dicionário de sinônimos de colunas para mapeamento inteligente
# Dicionário de sinônimos de colunas para mapeamento inteligente
COLUNAS_MAPEAMENTO = {
    'placa': ['placa', 'placas', 'placa_veiculo', 'placas_veiculos', 'placa_do_veiculo', 'veiculo', 'veiculos', 'veiculo_placa', 'num_placa', 'placa_carro', 'placa_vtr', 'placa_viatura', 'lista_placas', 'viatura', 'vtr_placa', 'placas_alvo'],
    'prefixo': ['prefixo', 'prefixos', 'prefixo_vtr', 'vtr', 'prefixo_viatura', 'codigo', 'cod_veiculo', 'id_veiculo', 'frota', 'nr_frota', 'viatura_prefixo'],
    'data': ['data', 'data_abastecimento', 'dt_abastecimento', 'data_hora', 'data_transacao', 'dt_transacao', 'data/hora', 'data_movimento'],
    'odometro': ['odometro', 'km', 'hodometro', 'km_atual', 'km_final', 'km_abastecimento', 'quilometragem', 'odometro_atual'],
    'litros': ['litros', 'volume', 'quantidade', 'qtd', 'qtd_litros', 'litros_abastecidos', 'volume_abastecido', 'quant'],
    'valor_total': ['valor_total', 'valor', 'total', 'vl_total', 'valor_abastecimento', 'valor_transacao', 'total_rs', 'valor_r$'],
    'valor_unitario': ['valor_unitario', 'preco_litro', 'vl_unitario', 'preco_unitario', 'unitario', 'valor_litro', 'valor/litro'],
    'tipo_combustivel': ['combustivel', 'tipo_combustivel', 'produto', 'descricao_produto', 'tipo_produto', 'desc_combustivel', 'servico'],
    'posto': ['posto', 'estabelecimento', 'nome_estabelecimento', 'nome_fantasia', 'razao_social', 'nome_posto', 'posto_combustivel'],
    'cidade': ['cidade', 'municipio', 'cidade_posto', 'localidade', 'cidade_abastecimento'],
    'motorista': ['motorista', 'condutor', 'nome_motorista', 'policial'],
    'modelo': ['modelo', 'marca_modelo', 'descricao', 'tipo', 'veiculo_modelo', 'descricao_veiculo', 'marca/modelo'],
    'fracao': ['fracao', 'companhia', 'cia', 'pelotao', 'unidade', 'lotacao', 'opm', 'subsetor', 'batalhao', 'responsavel']
}


def detectar_colunas(df: pd.DataFrame) -> Dict[str, str]:
    """Identifica automaticamente as colunas da planilha usando sinônimos com prioridade para match exato."""
    colunas_encontradas = {}
    colunas_originais = df.columns.tolist()

    # Pass 1: Correspondência Exata de sinônimo
    for col_orig in colunas_originais:
        col_norm = remover_acentos(str(col_orig)).replace('-', '_').replace(' ', '_').replace('.', '_')
        for chave, sinonimos in COLUNAS_MAPEAMENTO.items():
            if chave in colunas_encontradas:
                continue
            for sin in sinonimos:
                sin_norm = remover_acentos(sin).replace('-', '_').replace(' ', '_').replace('.', '_')
                if col_norm == sin_norm:
                    colunas_encontradas[chave] = col_orig
                    break

    # Pass 2: Correspondência Parcial (Substring) para as chaves ainda não mapeadas
    for col_orig in colunas_originais:
        col_norm = remover_acentos(str(col_orig)).replace('-', '_').replace(' ', '_').replace('.', '_')
        for chave, sinonimos in COLUNAS_MAPEAMENTO.items():
            if chave in colunas_encontradas:
                continue
            for sin in sinonimos:
                sin_norm = remover_acentos(sin).replace('-', '_').replace(' ', '_').replace('.', '_')
                if sin_norm in col_norm:
                    colunas_encontradas[chave] = col_orig
                    break

    return colunas_encontradas


def _extensao_arquivo(caminho_ou_arquivo) -> str:
    """Retorna a extensão do arquivo em minúsculas, tratando strings e objetos de arquivo."""
    if isinstance(caminho_ou_arquivo, str):
        return os.path.splitext(caminho_ou_arquivo.lower())[1]
    nome = getattr(caminho_ou_arquivo, 'filename', '') or getattr(caminho_ou_arquivo, 'name', '')
    return os.path.splitext(str(nome).lower())[1]


def carregar_planilha_excel(caminho_ou_arquivo) -> pd.DataFrame:
    """
    Lê com extrema robustez planilhas em TODAS as versões Excel e formatos similares:
      .xlsx  — Excel 2007+ (OpenXML)        → openpyxl
      .xlsm  — Excel com macros (OpenXML)   → openpyxl
      .xls   — Excel 97-2003                → xlrd
      .xlsb  — Excel Binário                → pyxlsb (se instalado)
      .ods   — LibreOffice / Google Sheets  → odf (se instalado)
      .csv / .tsv — Texto delimitado
      Arquivos HTML/XML renomeados como .xls (exportação de sistemas estaduais)
    """
    ext = _extensao_arquivo(caminho_ou_arquivo)

    # ── 1. Excel Binário (.xlsb) ────────────────────────────────────────────
    if ext == '.xlsb':
        try:
            return pd.read_excel(caminho_ou_arquivo, engine='pyxlsb')
        except Exception:
            pass  # fallback abaixo

    # ── 2. ODS / LibreOffice (.ods) ─────────────────────────────────────────
    if ext == '.ods':
        try:
            return pd.read_excel(caminho_ou_arquivo, engine='odf')
        except Exception:
            pass  # fallback abaixo

    # ── 3. CSV / TSV ────────────────────────────────────────────────────────
    if ext in ('.csv', '.tsv', '.txt'):
        sep = '\t' if ext == '.tsv' else None  # None = autodetect pelo Python engine
        for enc in ['utf-8-sig', 'latin1', 'cp1252', 'utf-8']:
            try:
                return pd.read_csv(caminho_ou_arquivo, sep=sep, engine='python', encoding=enc)
            except Exception:
                continue

    # ── 4. Excel legado .xls (Excel 97-2003) ────────────────────────────────
    if ext == '.xls':
        try:
            return pd.read_excel(caminho_ou_arquivo, engine='xlrd')
        except Exception:
            pass  # pode ser HTML renomeado — cai nos fallbacks

    # ── 5. Excel moderno .xlsx / .xlsm e qualquer outro ────────────────────
    try:
        return pd.read_excel(caminho_ou_arquivo, engine='openpyxl')
    except Exception:
        pass

    # ── 6. Tentar xlrd como fallback geral ──────────────────────────────────
    try:
        return pd.read_excel(caminho_ou_arquivo, engine='xlrd')
    except Exception:
        pass

    # ── 7. Tentar pyxlsb como fallback (arquivo .xlsb sem extensão correta) ─
    try:
        return pd.read_excel(caminho_ou_arquivo, engine='pyxlsb')
    except Exception:
        pass

    # ── 8. Tabela HTML (sistemas estaduais exportam HTML renomeado para .xls) ─
    try:
        dfs = pd.read_html(caminho_ou_arquivo)
        if dfs:
            maior_df = max(dfs, key=lambda d: d.shape[0] * d.shape[1])
            if isinstance(maior_df.columns, pd.MultiIndex):
                maior_df.columns = ['_'.join(str(c) for c in col).strip() for col in maior_df.columns.values]
            return maior_df
    except Exception:
        pass

    # ── 9. CSV/TSV como último recurso (arquivo sem extensão reconhecível) ──
    for enc in ['utf-8-sig', 'latin1', 'cp1252', 'utf-8']:
        try:
            return pd.read_csv(caminho_ou_arquivo, sep=None, engine='python', encoding=enc)
        except Exception:
            continue

    # ── 10. Tentativa final com pd.read_excel sem engine forçado ────────────
    return pd.read_excel(caminho_ou_arquivo)


def extrair_placas_de_texto(texto: str) -> List[Dict[str, Any]]:
    """Extrai e normaliza uma lista de placas a partir de texto digitado ou colado."""
    if not texto:
        return []
    
    # Quebra por quebras de linha, vírgulas, ponto e vírgula ou espaços múltiplos
    itens_brutos = re.split(r'[\r\n,;]+', str(texto))
    lista_veiculos = []
    placas_vistas = set()

    for item in itens_brutos:
        item_limpo = item.strip()
        if not item_limpo:
            continue
        placa = normalizar_placa(item_limpo)
        if placa and len(placa) >= 5 and placa not in placas_vistas:
            placas_vistas.add(placa)
            lista_veiculos.append({
                'placa': placa,
                'prefixo': '',
                'modelo': '46º BPM',
                'fracao': '46º BPM - Bagé',
                'tipo_combustivel': 'Gasolina Comum'
            })

    return lista_veiculos


def processar_planilha_viaturas_alvo(caminho_ou_arquivo) -> List[Dict[str, Any]]:
    """
    Lê e padroniza a planilha de viaturas alvo fornecida pelo usuário.
    É ultra-tolerante: aceita planilhas completas com modelo/fração ou apenas uma coluna de placas.
    """
    df = carregar_planilha_excel(caminho_ou_arquivo)
    cols = detectar_colunas(df)

    # Se a primeira linha for título ou vazia, procura o cabeçalho real nas primeiras linhas
    if not cols.get('placa') and not cols.get('prefixo'):
        for r_idx in range(min(6, len(df))):
            vals = [str(x).upper() for x in df.iloc[r_idx].dropna()]
            if any('PLACA' in x for x in vals) or any('PREFIXO' in x for x in vals):
                df.columns = df.iloc[r_idx]
                df = df.iloc[r_idx + 1:].reset_index(drop=True)
                cols = detectar_colunas(df)
                break

    col_placa = cols.get('placa')
    col_prefixo = cols.get('prefixo')
    col_modelo = cols.get('modelo')
    col_fracao = cols.get('fracao')
    col_comb = cols.get('tipo_combustivel')

    # Se a coluna de placa não foi encontrada pelos nomes comuns, inspeciona o conteúdo das colunas
    if not col_placa and not col_prefixo:
        # Se só tem 1 coluna, assume que são as placas
        if len(df.columns) == 1:
            col_placa = df.columns[0]
        else:
            # Procura coluna com padrões de placa (3 letras + 4 números ou padrão Mercosul)
            padrao_placa = re.compile(r'^[A-Za-z]{3}[0-9][A-Za-z0-9][0-9]{2}$')
            melhor_col = None
            maior_taxa = 0
            for c in df.columns:
                amostras = df[c].dropna().astype(str).str.strip().str.replace('-', '').str.replace(' ', '')
                if len(amostras) == 0:
                    continue
                matches = sum(1 for v in amostras if padrao_placa.match(v) or (len(v) == 7 and v[:3].isalpha()))
                taxa = matches / len(amostras)
                if taxa > maior_taxa and taxa >= 0.2:
                    maior_taxa = taxa
                    melhor_col = c
            if melhor_col:
                col_placa = melhor_col
            else:
                # Fallback: primeira coluna
                col_placa = df.columns[0]

    # Também verifica se a própria linha de cabeçalho é uma placa (planilha sem cabeçalho)
    cabecalho_eh_placa = normalizar_placa(col_placa)
    linhas_extras = []
    if cabecalho_eh_placa and len(cabecalho_eh_placa) >= 6 and (cabecalho_eh_placa[:3].isalpha() or cabecalho_eh_placa[3:].isdigit()):
        linhas_extras.append({
            'placa': cabecalho_eh_placa,
            'prefixo': '',
            'modelo': '',
            'fracao': '46º BPM',
            'tipo_combustivel': 'Gasolina Comum'
        })

    lista_veiculos = []
    placas_vistas = set()

    for item in linhas_extras:
        lista_veiculos.append(item)
        placas_vistas.add(item['placa'])

    for _, row in df.iterrows():
        placa = normalizar_placa(row[col_placa]) if col_placa and pd.notna(row.get(col_placa)) else ""
        prefixo = normalizar_prefixo(row[col_prefixo]) if col_prefixo and pd.notna(row.get(col_prefixo)) else ""
        modelo = str(row.get(col_modelo, '')).strip() if col_modelo and pd.notna(row.get(col_modelo)) else ""
        fracao = str(row.get(col_fracao, '46º BPM')).strip() if col_fracao and pd.notna(row.get(col_fracao)) else "46º BPM"
        tipo_comb = str(row.get(col_comb, 'Gasolina Comum')).strip() if col_comb and pd.notna(row.get(col_comb)) else "Gasolina Comum"

        if (placa or prefixo) and placa not in placas_vistas:
            if placa:
                placas_vistas.add(placa)
            lista_veiculos.append({
                'placa': placa,
                'prefixo': prefixo,
                'modelo': modelo,
                'fracao': fracao,
                'tipo_combustivel': tipo_comb
            })

    return lista_veiculos


def processar_cruzamento_abastecimentos(
    caminho_arquivo_rs,
    lista_veiculos_alvo: List[Dict[str, Any]]
) -> Tuple[Dict[str, Any], List[Dict[str, Any]], pd.DataFrame]:
    """
    Processa a planilha do RS e cruza estritamente com os veículos do 46º BPM.
    Calcula quilometragem percorrida, consumo (km/l), custo/km e detecta anomalias.
    Retorna: (resumo_meta, lista_abastecimentos_calculados, df_resumo_viaturas)
    """
    df_rs = carregar_planilha_excel(caminho_arquivo_rs)
    total_linhas_rs = len(df_rs)

    cols_rs = detectar_colunas(df_rs)

    # Identificar colunas essenciais
    col_placa = cols_rs.get('placa')
    col_prefixo = cols_rs.get('prefixo')
    col_data = cols_rs.get('data')
    col_odometro = cols_rs.get('odometro')
    col_litros = cols_rs.get('litros')
    col_valor_total = cols_rs.get('valor_total')
    col_valor_unit = cols_rs.get('valor_unitario')
    col_comb = cols_rs.get('tipo_combustivel')
    col_posto = cols_rs.get('posto')
    col_cidade = cols_rs.get('cidade')
    col_motorista = cols_rs.get('motorista')

    # Criar conjuntos de busca dos veículos alvo
    placas_alvo = set()
    prefixos_alvo = set()
    mapa_veiculos_por_placa = {}
    mapa_veiculos_por_prefixo = {}

    for v in lista_veiculos_alvo:
        p = normalizar_placa(v.get('placa'))
        pref = normalizar_prefixo(v.get('prefixo'))
        if p:
            placas_alvo.add(p)
            mapa_veiculos_por_placa[p] = v
        if pref:
            prefixos_alvo.add(pref)
            mapa_veiculos_por_prefixo[pref] = v

    # Adicionar colunas normalizadas no DataFrame do RS para filtro ágil
    if col_placa:
        df_rs['_placa_norm'] = df_rs[col_placa].apply(normalizar_placa)
    else:
        df_rs['_placa_norm'] = ""

    if col_prefixo:
        df_rs['_prefixo_norm'] = df_rs[col_prefixo].apply(normalizar_prefixo)
    else:
        df_rs['_prefixo_norm'] = ""

    # Filtrar apenas registros correspondentes aos veículos alvo
    mascara = pd.Series(False, index=df_rs.index)
    if placas_alvo and col_placa:
        mascara = mascara | df_rs['_placa_norm'].isin(placas_alvo)
    if prefixos_alvo and col_prefixo:
        mascara = mascara | df_rs['_prefixo_norm'].isin(prefixos_alvo)

    df_filtrado = df_rs[mascara].copy()

    if df_filtrado.empty:
        meta_vazia = {
            'total_linhas_rs': total_linhas_rs,
            'total_filtrado': 0,
            'total_veiculos': 0,
            'total_litros': 0.0,
            'total_valor': 0.0,
            'total_km': 0.0,
            'media_km_l': 0.0
        }
        return meta_vazia, [], pd.DataFrame()

    # Formatar dados padronizados
    df_filtrado['data_abastecimento'] = pd.to_datetime(
        df_filtrado[col_data] if col_data else datetime.now(),
        dayfirst=True,
        errors='coerce'
    )
    # Tratar datas inválidas
    df_filtrado['data_abastecimento'] = df_filtrado['data_abastecimento'].fillna(pd.Timestamp.now())

    df_filtrado['odometro_num'] = df_filtrado[col_odometro].apply(limpar_numero) if col_odometro else 0.0
    df_filtrado['litros_num'] = df_filtrado[col_litros].apply(limpar_numero) if col_litros else 0.0
    df_filtrado['valor_total_num'] = df_filtrado[col_valor_total].apply(limpar_numero) if col_valor_total else 0.0
    df_filtrado['valor_unit_num'] = df_filtrado[col_valor_unit].apply(limpar_numero) if col_valor_unit else 0.0

    # Se valor unitário estiver zerado mas houver total e litros, calcula
    df_filtrado['valor_unit_num'] = df_filtrado.apply(
        lambda r: round(r['valor_total_num'] / r['litros_num'], 3) if r['valor_unit_num'] <= 0 and r['litros_num'] > 0 else r['valor_unit_num'],
        axis=1
    )

    # Identificador principal de agrupamento por viatura
    df_filtrado['chave_viatura'] = df_filtrado.apply(
        lambda r: r['_placa_norm'] if r['_placa_norm'] else r['_prefixo_norm'],
        axis=1
    )

    # Ordenar por viatura e cronologicamente (Data e depois Odômetro)
    df_filtrado = df_filtrado.sort_values(by=['chave_viatura', 'data_abastecimento', 'odometro_num'], ascending=[True, True, True])

    # Lista de saída com os cálculos detalhados
    lista_abastecimentos = []

    for chave, grupo in df_filtrado.groupby('chave_viatura'):
        # Buscar dados de cadastro da viatura
        v_meta = mapa_veiculos_por_placa.get(chave) or mapa_veiculos_por_prefixo.get(chave) or {}
        modelo_meta = v_meta.get('modelo', '')
        fracao_meta = v_meta.get('fracao', '46º BPM')
        prefixo_meta = v_meta.get('prefixo', '')

        km_anterior = None

        for idx, row in grupo.iterrows():
            placa = row['_placa_norm']
            prefixo = row['_prefixo_norm'] or prefixo_meta
            odometro_atual = float(row['odometro_num'])
            litros = float(row['litros_num'])
            valor_total = float(row['valor_total_num'])
            valor_unit = float(row['valor_unit_num'])

            tipo_comb = str(row[col_comb]).strip() if col_comb and pd.notna(row.get(col_comb)) else v_meta.get('tipo_combustivel', 'Gasolina Comum')
            posto = str(row[col_posto]).strip() if col_posto and pd.notna(row.get(col_posto)) else 'Não especificado'
            cidade = str(row[col_cidade]).strip() if col_cidade and pd.notna(row.get(col_cidade)) else 'Bagé'
            motorista = str(row[col_motorista]).strip() if col_motorista and pd.notna(row.get(col_motorista)) else ''

            km_rodado = 0.0
            consumo_km_l = 0.0
            custo_km = 0.0
            anomalia = 0
            anomalia_obs = []

            if km_anterior is not None:
                delta_km = odometro_atual - km_anterior
                if delta_km > 0:
                    km_rodado = delta_km
                    if litros > 0:
                        consumo_km_l = round(km_rodado / litros, 2)
                        custo_km = round(valor_total / km_rodado, 2)

                        # Validação de consumo (viaturas policiais geralmente fazem entre 4 e 16 km/l)
                        if consumo_km_l < 2.5:
                            anomalia = 1
                            anomalia_obs.append(f"Consumo muito alto/atípico ({consumo_km_l} km/l)")
                        elif consumo_km_l > 30.0:
                            anomalia = 1
                            anomalia_obs.append(f"Consumo improvável ({consumo_km_l} km/l) - verificar odômetro")
                elif delta_km < 0:
                    anomalia = 1
                    anomalia_obs.append(f"Odômetro regrediu ({odometro_atual:,.0f} < anterior {km_anterior:,.0f})")
                else:
                    anomalia = 1
                    anomalia_obs.append("Odômetro idêntico ao abastecimento anterior")
            else:
                anomalia_obs.append("Primeiro registro do período (sem odômetro anterior para cálculo)")

            # Anomalia de litragem
            if litros <= 0:
                anomalia = 1
                anomalia_obs.append("Volume zerado ou negativo")
            elif litros > 120:
                anomalia = 1
                anomalia_obs.append(f"Volume acima do tanque comum ({litros:.1f} L)")

            dt_formatada = row['data_abastecimento'].strftime('%Y-%m-%d %H:%M')

            item_dict = {
                'data_abastecimento': dt_formatada,
                'placa': placa if placa else v_meta.get('placa', chave),
                'prefixo': prefixo,
                'modelo': modelo_meta,
                'fracao': fracao_meta,
                'odometro': odometro_atual,
                'km_anterior': km_anterior,
                'km_rodado': km_rodado,
                'litros': litros,
                'consumo_km_l': consumo_km_l,
                'valor_unitario': valor_unit,
                'valor_total': valor_total,
                'custo_km': custo_km,
                'tipo_combustivel': tipo_comb,
                'posto': posto,
                'cidade': cidade,
                'motorista': motorista,
                'anomalia': anomalia,
                'anomalia_obs': " | ".join(anomalia_obs) if anomalia_obs else ""
            }

            lista_abastecimentos.append(item_dict)

            # Atualiza o odômetro anterior apenas se o atual for positivo
            if odometro_atual > 0:
                km_anterior = odometro_atual

    # Montar DataFrame de resumo por viatura
    df_abs = pd.DataFrame(lista_abastecimentos)

    total_filtrado = len(df_abs)
    total_veiculos = df_abs['placa'].nunique()
    total_litros = float(df_abs['litros'].sum())
    total_valor = float(df_abs['valor_total'].sum())
    total_km = float(df_abs[df_abs['km_rodado'] > 0]['km_rodado'].sum())

    # Consumo médio ponderado global (apenas trechos válidos)
    litros_trechos_validos = float(df_abs[(df_abs['km_rodado'] > 0) & (df_abs['anomalia'] == 0)]['litros'].sum())
    media_global_km_l = round(total_km / litros_trechos_validos, 2) if litros_trechos_validos > 0 else 0.0

    meta = {
        'total_linhas_rs': total_linhas_rs,
        'total_filtrado': total_filtrado,
        'total_veiculos': total_veiculos,
        'total_litros': round(total_litros, 2),
        'total_valor': round(total_valor, 2),
        'total_km': round(total_km, 1),
        'media_km_l': media_global_km_l
    }

    # Resumo agrupado por viatura
    resumo_veiculos = []
    for placa, grp in df_abs.groupby('placa'):
        pref = grp['prefixo'].iloc[0] if grp['prefixo'].notna().any() else ""
        mod = grp['modelo'].iloc[0] if grp['modelo'].notna().any() else ""
        frac = grp['fracao'].iloc[0] if grp['fracao'].notna().any() else "46º BPM"
        qtd_abs = len(grp)
        km_ini = grp['odometro'].min()
        km_fim = grp['odometro'].max()
        km_tot = grp[grp['km_rodado'] > 0]['km_rodado'].sum()
        lt_tot = grp['litros'].sum()
        val_tot = grp['valor_total'].sum()

        lt_validos = grp[(grp['km_rodado'] > 0) & (grp['anomalia'] == 0)]['litros'].sum()
        c_medio = round(km_tot / lt_validos, 2) if lt_validos > 0 else 0.0
        custo_km = round(val_tot / km_tot, 2) if km_tot > 0 else 0.0
        anomalias = grp[grp['anomalia'] == 1].shape[0]

        resumo_veiculos.append({
            'placa': placa,
            'prefixo': pref,
            'modelo': mod,
            'fracao': frac,
            'qtd_abastecimentos': qtd_abs,
            'km_inicial': km_ini,
            'km_final': km_fim,
            'km_total_rodado': km_tot,
            'total_litros': round(lt_tot, 2),
            'consumo_medio_km_l': c_medio,
            'total_valor': round(val_tot, 2),
            'custo_km': custo_km,
            'qtd_anomalias': anomalias
        })

    df_resumo_viaturas = pd.DataFrame(resumo_veiculos)

    return meta, lista_abastecimentos, df_resumo_viaturas


def gerar_relatorio_excel_openpyxl(
    caminho_saida: str,
    resumo_meta: Dict[str, Any],
    lista_abastecimentos: List[Dict[str, Any]],
    lista_resumo_viaturas: List[Dict[str, Any]]
):
    """
    Gera uma planilha Excel executiva estilizada com o openpyxl:
    - Aba 1: Resumo_Frota_46BPM
    - Aba 2: Abastecimentos_Detalhados
    - Aba 3: Alertas_e_Inconsistencias
    """
    wb = openpyxl.Workbook()
    # Remove sheet padrão
    wb.remove(wb.active)

    # Cores Militares / Táticas
    COR_CABECALHO_FUNDO = "002B49"   # Azul Petróleo / Brigada Militar
    COR_CABECALHO_TEXTO = "FFFFFF"
    COR_SUB_FUNDO = "0A4D68"
    COR_ZEBRA = "F4F6F9"
    COR_ALERTA_FUNDO = "FFF3CD"
    COR_ALERTA_TEXTO = "856404"
    COR_CARD_BG = "EBF3FA"

    fonte_titulo = Font(name="Calibri", size=16, bold=True, color="002B49")
    fonte_subtitulo = Font(name="Calibri", size=11, italic=True, color="555555")
    fonte_header = Font(name="Calibri", size=11, bold=True, color=COR_CABECALHO_TEXTO)
    fonte_card_titulo = Font(name="Calibri", size=10, bold=True, color="333333")
    fonte_card_valor = Font(name="Calibri", size=14, bold=True, color="002B49")
    fonte_dados = Font(name="Calibri", size=11)
    fonte_alerta = Font(name="Calibri", size=11, color="9C0006", bold=True)

    fill_header = PatternFill(start_color=COR_CABECALHO_FUNDO, end_color=COR_CABECALHO_FUNDO, fill_type="solid")
    fill_zebra = PatternFill(start_color=COR_ZEBRA, end_color=COR_ZEBRA, fill_type="solid")
    fill_card = PatternFill(start_color=COR_CARD_BG, end_color=COR_CARD_BG, fill_type="solid")
    fill_alerta = PatternFill(start_color=COR_ALERTA_FUNDO, end_color=COR_ALERTA_FUNDO, fill_type="solid")

    thin_border = Border(
        left=Side(style='thin', color='D0D7DE'),
        right=Side(style='thin', color='D0D7DE'),
        top=Side(style='thin', color='D0D7DE'),
        bottom=Side(style='thin', color='D0D7DE')
    )

    # -------------------------------------------------------------
    # ABA 1: RESUMO POR VIATURA
    # -------------------------------------------------------------
    ws1 = wb.create_sheet(title="Resumo_Frota_46BPM")
    ws1.views.sheetView[0].showGridLines = True

    # Cabeçalho Institucional
    ws1["A1"] = "46º BATALHÃO DE POLÍCIA MILITAR - BAGÉ / RS"
    ws1["A1"].font = fonte_titulo
    ws1["A2"] = f"Relatório de Consumo de Combustível da Frota | Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    ws1["A2"].font = fonte_subtitulo

    # Cards de KPI (Linhas 4 a 6)
    kpis = [
        ("Viaturas Monitoradas", f"{resumo_meta.get('total_veiculos', 0)} vtrs", "B4:C5"),
        ("Km Total Rodado", f"{resumo_meta.get('total_km', 0):,.1f} km", "E4:F5"),
        ("Litros Abastecidos", f"{resumo_meta.get('total_litros', 0):,.2f} L", "H4:I5"),
        ("Consumo Médio Geral", f"{resumo_meta.get('media_km_l', 0):,.2f} km/l", "K4:L5"),
        ("Gasto Total (R$)", f"R$ {resumo_meta.get('total_valor', 0):,.2f}", "N4:O5")
    ]

    for label, valor, pos_range in kpis:
        c_ini = pos_range.split(":")[0]
        ws1[c_ini] = f"{label}\n{valor}"
        ws1[c_ini].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws1[c_ini].font = fonte_card_valor
        ws1[c_ini].fill = fill_card
        ws1.merge_cells(pos_range)

    # Tabela de Veículos
    colunas_resumo = [
        "Prefixo", "Placa", "Modelo", "Fração/Lotação", "Abastecimentos",
        "Km Inicial", "Km Final", "Km Percorrido", "Litros Totais",
        "Consumo Médio (km/l)", "Valor Total (R$)", "Custo (R$/km)", "Alertas"
    ]

    row_idx = 8
    for col_num, col_nome in enumerate(colunas_resumo, start=1):
        c = ws1.cell(row=row_idx, column=col_num, value=col_nome)
        c.font = fonte_header
        c.fill = fill_header
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = thin_border

    ws1.row_dimensions[row_idx].height = 25

    row_idx = 9
    for i, v in enumerate(lista_resumo_viaturas):
        bg = fill_zebra if i % 2 == 1 else PatternFill(fill_type=None)

        valores = [
            v.get('prefixo', ''),
            v.get('placa', ''),
            v.get('modelo', ''),
            v.get('fracao', ''),
            v.get('qtd_abastecimentos', 0),
            v.get('km_inicial', 0),
            v.get('km_final', 0),
            v.get('km_total_rodado', 0),
            v.get('total_litros', 0),
            v.get('consumo_medio_km_l', 0),
            v.get('total_valor', 0),
            v.get('custo_km', 0),
            v.get('qtd_anomalias', 0)
        ]

        for col_num, val in enumerate(valores, start=1):
            cell = ws1.cell(row=row_idx, column=col_num, value=val)
            cell.font = fonte_dados
            if bg.fill_type:
                cell.fill = bg
            cell.border = thin_border

            # Formatações numéricas
            if col_num in [1, 2, 4]:
                cell.alignment = Alignment(horizontal="center")
            elif col_num == 3:
                cell.alignment = Alignment(horizontal="left")
            elif col_num == 5 or col_num == 13:
                cell.alignment = Alignment(horizontal="center")
                cell.number_format = '#,##0'
            elif col_num in [6, 7, 8]:
                cell.alignment = Alignment(horizontal="right")
                cell.number_format = '#,##0.0'
            elif col_num == 9 or col_num == 10:
                cell.alignment = Alignment(horizontal="right")
                cell.number_format = '#,##0.00'
            elif col_num in [11, 12]:
                cell.alignment = Alignment(horizontal="right")
                cell.number_format = 'R$ #,##0.00'

        row_idx += 1

    # Linha de Totais
    ws1.cell(row=row_idx, column=1, value="TOTAL GERAL").font = Font(name="Calibri", size=11, bold=True)
    ws1.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=4)
    ws1.cell(row=row_idx, column=1).alignment = Alignment(horizontal="center")

    totais_fórmulas = [
        (5, f"=SUM(E9:E{row_idx-1})", '#,##0'),
        (8, f"=SUM(H9:H{row_idx-1})", '#,##0.0'),
        (9, f"=SUM(I9:I{row_idx-1})", '#,##0.00'),
        (10, f"=IF(I{row_idx}>0, H{row_idx}/I{row_idx}, 0)", '#,##0.00'),
        (11, f"=SUM(K9:K{row_idx-1})", 'R$ #,##0.00'),
        (12, f"=IF(H{row_idx}>0, K{row_idx}/H{row_idx}, 0)", 'R$ #,##0.00'),
        (13, f"=SUM(M9:M{row_idx-1})", '#,##0')
    ]
    for c_idx, formula, fmt in totais_fórmulas:
        c = ws1.cell(row=row_idx, column=c_idx, value=formula)
        c.font = Font(name="Calibri", size=11, bold=True)
        c.number_format = fmt
        c.border = thin_border
        c.fill = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")

    # Ajuste automático de largura das colunas
    for col in ws1.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = max(len(str(cell.value or '')) for cell in col)
        ws1.column_dimensions[col_letter].width = max(max_len + 3, 12)

    # -------------------------------------------------------------
    # ABA 2: ABASTECIMENTOS DETALHADOS
    # -------------------------------------------------------------
    ws2 = wb.create_sheet(title="Abastecimentos_46BPM")
    ws2.views.sheetView[0].showGridLines = True

    colunas_detalhes = [
        "Data / Hora", "Prefixo", "Placa", "Modelo", "Odômetro (km)", "Km Anterior",
        "Km Rodado", "Litros", "Consumo (km/l)", "Preço/L (R$)", "Total (R$)",
        "Custo (R$/km)", "Combustível", "Posto / Estabelecimento", "Cidade", "Motorista", "Observações"
    ]

    for col_num, col_nome in enumerate(colunas_detalhes, start=1):
        c = ws2.cell(row=1, column=col_num, value=col_nome)
        c.font = fonte_header
        c.fill = fill_header
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = thin_border
    ws2.row_dimensions[1].height = 25

    row_idx2 = 2
    for i, a in enumerate(lista_abastecimentos):
        bg = fill_zebra if i % 2 == 1 else PatternFill(fill_type=None)
        if a.get('anomalia', 0) == 1:
            bg = fill_alerta

        itens = [
            a.get('data_abastecimento', ''),
            a.get('prefixo', ''),
            a.get('placa', ''),
            a.get('modelo', ''),
            a.get('odometro', 0),
            a.get('km_anterior', '') if a.get('km_anterior') is not None else '-',
            a.get('km_rodado', 0),
            a.get('litros', 0),
            a.get('consumo_km_l', 0),
            a.get('valor_unitario', 0),
            a.get('valor_total', 0),
            a.get('custo_km', 0),
            a.get('tipo_combustivel', ''),
            a.get('posto', ''),
            a.get('cidade', ''),
            a.get('motorista', ''),
            a.get('anomalia_obs', '')
        ]

        for col_num, val in enumerate(itens, start=1):
            c = ws2.cell(row=row_idx2, column=col_num, value=val)
            c.font = fonte_dados
            if bg.fill_type:
                c.fill = bg
            c.border = thin_border

            if col_num in [1, 2, 3, 13, 15]:
                c.alignment = Alignment(horizontal="center")
            elif col_num in [5, 6, 7]:
                c.alignment = Alignment(horizontal="right")
                if isinstance(val, (int, float)):
                    c.number_format = '#,##0'
            elif col_num in [8, 9]:
                c.alignment = Alignment(horizontal="right")
                if isinstance(val, (int, float)):
                    c.number_format = '#,##0.00'
            elif col_num in [10, 11, 12]:
                c.alignment = Alignment(horizontal="right")
                if isinstance(val, (int, float)):
                    c.number_format = 'R$ #,##0.00'

        row_idx2 += 1

    for col in ws2.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = max(len(str(cell.value or '')) for cell in col[:100])
        ws2.column_dimensions[col_letter].width = max(min(max_len + 3, 35), 11)

    # -------------------------------------------------------------
    # ABA 3: ALERTAS E INCONSISTÊNCIAS
    # -------------------------------------------------------------
    ws3 = wb.create_sheet(title="Alertas_Inconsistencias")
    ws3.views.sheetView[0].showGridLines = True

    anomalias = [a for a in lista_abastecimentos if a.get('anomalia', 0) == 1]

    colunas_alertas = [
        "Data / Hora", "Prefixo", "Placa", "Odômetro", "Km Anterior", "Km Rodado",
        "Litros", "Consumo (km/l)", "Posto", "Motivo do Alerta / Inconsistência"
    ]

    for col_num, col_nome in enumerate(colunas_alertas, start=1):
        c = ws3.cell(row=1, column=col_num, value=col_nome)
        c.font = fonte_header
        c.fill = PatternFill(start_color="856404", end_color="856404", fill_type="solid")
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = thin_border
    ws3.row_dimensions[1].height = 25

    row_idx3 = 2
    for a in anomalias:
        vals = [
            a.get('data_abastecimento', ''),
            a.get('prefixo', ''),
            a.get('placa', ''),
            a.get('odometro', 0),
            a.get('km_anterior', '-') if a.get('km_anterior') is not None else '-',
            a.get('km_rodado', 0),
            a.get('litros', 0),
            a.get('consumo_km_l', 0),
            a.get('posto', ''),
            a.get('anomalia_obs', '')
        ]

        for col_num, val in enumerate(vals, start=1):
            c = ws3.cell(row=row_idx3, column=col_num, value=val)
            c.font = fonte_alerta if col_num == 10 else fonte_dados
            c.border = thin_border
            c.fill = fill_alerta
            if col_num in [1, 2, 3]:
                c.alignment = Alignment(horizontal="center")
            elif col_num in [4, 5, 6, 7, 8]:
                c.alignment = Alignment(horizontal="right")

        row_idx3 += 1

    for col in ws3.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = max(len(str(cell.value or '')) for cell in col[:100])
        ws3.column_dimensions[col_letter].width = max(min(max_len + 3, 40), 12)

    # Salva o arquivo gerado
    wb.save(caminho_saida)
    return caminho_saida


# ═════════════════════════════════════════════════════════════════════════════
# PLANILHA CONSOLIDADA DO ESTADO + PLAN2 DE CONFERÊNCIA (PADRÃO AO CRPMFO)
# ═════════════════════════════════════════════════════════════════════════════

MESES_PT = {
    1: 'Janeiro', 2: 'Fevereiro', 3: 'Março', 4: 'Abril',
    5: 'Maio', 6: 'Junho', 7: 'Julho', 8: 'Agosto',
    9: 'Setembro', 10: 'Outubro', 11: 'Novembro', 12: 'Dezembro'
}


def normalizar_cidade_resumo(cid: str) -> str:
    """Normaliza o nome da cidade para exibição padronizada na tabela lateral."""
    c = (cid or '').strip().upper()
    if c in ['EM', 'ESTADO MAIOR', 'ESTADO-MAIOR', '46º BPM - EM']:
        return 'EM'
    if 'BAGE' in c or 'BAGÉ' in c:
        return 'Bagé'
    if 'ACEGUA' in c or 'ACEGUÁ' in c:
        return 'Aceguá'
    if 'HULHA' in c:
        return 'Hulha Negra'
    if 'CANDIOTA' in c:
        return 'Candiota'
    if 'LAVRAS' in c:
        return 'Lavras do Sul'
    if 'PEDRAS' in c:
        return 'Pedras Altas'
    if 'PINHEIRO' in c:
        return 'Pinheiro Machado'
    if 'DOM PEDRITO' in c:
        return 'Dom Pedrito'
    return cid.strip().title() if cid else 'Outros'


def detectar_nome_arquivo_crpmfo(datas_encontradas: List[Any], data_fallback: Optional[datetime] = None) -> str:
    """
    Determina o nome oficial do arquivo baseado na quinzena e mês:
    'Relatório da {1|2}ª quinzena de {Mês} - AO CRPMFO.xlsx'
    """
    data_ref = None
    if datas_encontradas:
        # Pega a data máxima válida do lote
        data_ref = max(datas_encontradas)
    if not data_ref or not isinstance(data_ref, (datetime, pd.Timestamp)):
        data_ref = data_fallback or datetime.now()

    quinzena = 1 if data_ref.day <= 15 else 2
    mes_nome = MESES_PT.get(data_ref.month, 'Mês')
    return f"Relatório da {quinzena}ª quinzena de {mes_nome} - AO CRPMFO.xlsx"


def gerar_planilha_consolidada_estado_e_plan2(
    caminho_arquivo_rs: str,
    viaturas_alvo: List[Dict[str, Any]],
    pasta_destino: str
) -> Tuple[str, str]:
    """
    Gera na mesma planilha:
    - Aba 1: Os abastecimentos brutos de todo o estado do RS (todas as linhas e colunas originais).
    - Aba 2: 'Plan2', contendo a conferência do 46º BPM:
      - Placa e valores individuais de cada abastecimento
      - Total da viatura e Fração/Cidade
      - 2 linhas em branco de espaçamento entre cada viatura
      - Linha final com 'total' e a soma geral
      - Tabela lateral com o resumo consolidado por Fração/Cidade e Total Geral
    Retorna (caminho_completo_arquivo, nome_arquivo_formatado).
    """
    ext = os.path.splitext(caminho_arquivo_rs)[1].lower()
    linhas_brutas = []
    nome_aba_original = "RFCVTITULO"
    datas_lote = []

    # 1. Carregar linhas da planilha estadual
    if ext == '.xls':
        try:
            import xlrd
            wb_src = xlrd.open_workbook(caminho_arquivo_rs)
            s_src = wb_src.sheet_by_name(wb_src.sheet_names()[0])
            nome_aba_original = s_src.name
            for r in range(s_src.nrows):
                row_vals = [s_src.cell_value(r, c) for c in range(s_src.ncols)]
                linhas_brutas.append(row_vals)
                if r > 0 and s_src.ncols > 4:
                    d_raw = s_src.cell_value(r, 4)
                    if isinstance(d_raw, (int, float)):
                        try:
                            datas_lote.append(xlrd.xldate_as_datetime(d_raw, wb_src.datemode))
                        except Exception:
                            pass
        except Exception as e:
            print(f"Aviso ao ler com xlrd: {e}")

    if not linhas_brutas:
        # Fallback para openpyxl ou pandas
        try:
            wb_src = openpyxl.load_workbook(caminho_arquivo_rs, data_only=True)
            s_src = wb_src.active
            nome_aba_original = s_src.title
            for row in s_src.iter_rows(values_only=True):
                linhas_brutas.append(list(row))
        except Exception:
            df = carregar_planilha_excel(caminho_arquivo_rs)
            nome_aba_original = "Abastecimentos_RS"
            linhas_brutas = [list(df.columns)] + df.values.tolist()

    if not linhas_brutas:
        raise ValueError("Não foi possível extrair dados da planilha do estado.")

    # 2. Localizar índices das colunas Placa, Valor e Data/Hora no cabeçalho
    cabecalho = [str(c or '').strip().lower() for c in linhas_brutas[0]]
    idx_placa = 1
    idx_valor = 15

    for i, col_txt in enumerate(cabecalho):
        if 'placa' in col_txt:
            idx_placa = i
            break

    for i, col_txt in enumerate(cabecalho):
        if col_txt in ['valor', 'total', 'valor total', 'valor_total', 'vl_total']:
            idx_valor = i
            break

    # Se não coletou datas pelo xlrd, tentar coletar da coluna de data
    if not datas_lote:
        idx_data = -1
        for i, col_txt in enumerate(cabecalho):
            if any(k in col_txt for k in ['data', 'dt_abast', 'data/hora']):
                idx_data = i
                break
        if idx_data >= 0:
            for row in linhas_brutas[1:1000]:
                if len(row) > idx_data and row[idx_data]:
                    val_d = str(row[idx_data])
                    m = re.search(r'(\d{2})/(\d{2})/(20\d{2})', val_d)
                    if m:
                        try:
                            datas_lote.append(datetime(int(m.group(3)), int(m.group(2)), int(m.group(1))))
                        except Exception:
                            pass

    # 3. Mapear abastecimentos por placa na planilha estadual
    abast_por_placa: Dict[str, List[float]] = {}
    for r in linhas_brutas[1:]:
        if len(r) > idx_placa and r[idx_placa]:
            p = normalizar_placa(r[idx_placa])
            if p:
                v_num = 0.0
                if len(r) > idx_valor and r[idx_valor] is not None:
                    v_num = limpar_numero(r[idx_valor])
                abast_por_placa.setdefault(p, []).append(v_num)

    # 4. Criar o novo workbook com as duas abas
    wb_out = openpyxl.Workbook()

    # Aba 1: Abastecimentos de todo o estado
    ws1 = wb_out.active
    ws1.title = nome_aba_original[:31]
    ws1.views.sheetView[0].showGridLines = True
    for row in linhas_brutas:
        ws1.append(row)

    # Libera linhas_brutas da memória antes de processar a Plan2
    del linhas_brutas
    gc.collect()

    # Aba 2: Plan2 (Conferência 46º BPM)
    ws2 = wb_out.create_sheet(title="Plan2")
    ws2.views.sheetView[0].showGridLines = True

    # Cabeçalho da Plan2
    ws2.cell(row=1, column=1, value="Placa")
    ws2.cell(row=1, column=2, value="Valor")

    cur_row = 2
    cidades_totais: Dict[str, float] = {}

    # Agrupar viaturas alvo na ordem oficial
    for v_obj in viaturas_alvo:
        p_raw = v_obj.get('placa', '')
        p_norm = normalizar_placa(p_raw)
        if not p_norm or p_norm not in abast_por_placa:
            continue

        valores = abast_por_placa[p_norm]
        if not valores:
            continue

        cid_raw = v_obj.get('cidade_carga') or v_obj.get('fracao') or 'BAGÉ'
        frac_raw = v_obj.get('fracao') or ''

        if frac_raw == 'EM' or cid_raw.strip().upper() == 'EM':
            fracao_label = 'EM'
            cidade_resumo = 'EM'
        else:
            cidade_resumo = normalizar_cidade_resumo(cid_raw)
            fracao_label = cidade_resumo.upper()

        tot_placa = sum(valores)
        cidades_totais[cidade_resumo] = cidades_totais.get(cidade_resumo, 0.0) + tot_placa

        end_r = cur_row + len(valores) - 1

        for i, val in enumerate(valores):
            row_num = cur_row + i
            ws2.cell(row=row_num, column=1, value=p_norm)
            c_val = ws2.cell(row=row_num, column=2, value=round(val, 2))
            c_val.number_format = '#,##0.00'

            if i == 0:
                c_tot = ws2.cell(row=row_num, column=3, value=round(tot_placa, 2))
                c_tot.number_format = '#,##0.00'
                ws2.cell(row=row_num, column=4, value=fracao_label)

        # 2 linhas em branco de espaçamento após cada viatura
        cur_row = end_r + 3

    # Linha Total Geral das Viaturas
    total_row = cur_row - 2
    grand_total = sum(cidades_totais.values())
    ws2.cell(row=total_row, column=1, value="total")
    c_gtot = ws2.cell(row=total_row, column=3, value=round(grand_total, 2))
    c_gtot.number_format = '#,##0.00'

    # Tabela Lateral Consolidada (Colunas F e G)
    ordem_cidades = [
        'EM', 'Bagé', 'Aceguá', 'Hulha Negra', 'Candiota',
        'Lavras do Sul', 'Pedras Altas', 'Pinheiro Machado', 'Dom Pedrito'
    ]
    for cid in cidades_totais:
        if cid not in ordem_cidades:
            ordem_cidades.append(cid)

    num_linhas_tabela = len(ordem_cidades) + 1
    start_side_row = max(2, total_row - num_linhas_tabela + 3)

    side_r = start_side_row
    for cid in ordem_cidades:
        tot_cid = cidades_totais.get(cid, 0.0)
        ws2.cell(row=side_r, column=6, value=cid)
        c_ctot = ws2.cell(row=side_r, column=7, value=round(tot_cid, 2))
        c_ctot.number_format = '#,##0.00'
        side_r += 1

    # Total geral da tabela lateral
    c_gtot2 = ws2.cell(row=side_r, column=7, value=round(grand_total, 2))
    c_gtot2.number_format = '#,##0.00'

    # Ajuste de larguras de coluna da Plan2
    ws2.column_dimensions['A'].width = 14
    ws2.column_dimensions['B'].width = 12
    ws2.column_dimensions['C'].width = 14
    ws2.column_dimensions['D'].width = 22
    ws2.column_dimensions['E'].width = 5
    ws2.column_dimensions['F'].width = 20
    ws2.column_dimensions['G'].width = 16

    # 5. Determinar nome oficial formatado
    nome_arquivo_formatado = detectar_nome_arquivo_crpmfo(datas_lote)
    caminho_saida = os.path.join(pasta_destino, nome_arquivo_formatado)
    wb_out.save(caminho_saida)
    wb_out.close()
    del wb_out
    gc.collect()

    return caminho_saida, nome_arquivo_formatado
