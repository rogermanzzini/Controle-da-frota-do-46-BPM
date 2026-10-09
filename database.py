"""
database.py - Módulo de Banco de Dados SQLite para o Sistema do 46º BPM (Bagé/RS)
Gerencia tabelas de viaturas alvo, abastecimentos filtrados e lotes de importação.
"""

import sqlite3
import os
from typing import List, Dict, Any, Optional

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'bpm46_combustivel.db')


def get_connection():
    """Retorna uma conexão ativa com o banco SQLite configurada com row_factory."""
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    # WAL mode allows concurrent readers + one writer and avoids "database is locked"
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA busy_timeout=10000")
    return conn


def init_db():
    """Inicializa as tabelas do banco de dados SQLite caso ainda não existam."""
    conn = get_connection()
    cursor = conn.cursor()

    # Tabela de Viaturas Alvo (46º BPM Bagé)
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS veiculos_alvo (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            placa TEXT UNIQUE NOT NULL,
            prefixo TEXT,
            modelo TEXT,
            fracao TEXT,
            tipo_combustivel TEXT DEFAULT 'Gasolina Comum',
            ativo INTEGER DEFAULT 1,
            data_cadastro TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')

    # Tabela de Importações / Lotes
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS importacoes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            data_hora TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            arquivo_rs TEXT,
            arquivo_alvo TEXT,
            total_linhas_rs INTEGER DEFAULT 0,
            total_filtrado INTEGER DEFAULT 0,
            total_veiculos INTEGER DEFAULT 0,
            total_litros REAL DEFAULT 0.0,
            total_valor REAL DEFAULT 0.0,
            total_km REAL DEFAULT 0.0,
            media_km_l REAL DEFAULT 0.0
        )
    ''')

    # Tabela de Abastecimentos Filtrados do 46º BPM
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS abastecimentos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            importacao_id INTEGER,
            data_abastecimento TEXT,
            placa TEXT NOT NULL,
            prefixo TEXT,
            modelo TEXT,
            fracao TEXT,
            odometro REAL,
            km_anterior REAL,
            km_rodado REAL,
            litros REAL,
            consumo_km_l REAL,
            valor_unitario REAL,
            valor_total REAL,
            custo_km REAL,
            tipo_combustivel TEXT,
            posto TEXT,
            cidade TEXT,
            motorista TEXT,
            anomalia INTEGER DEFAULT 0,
            anomalia_obs TEXT,
            FOREIGN KEY (importacao_id) REFERENCES importacoes (id) ON DELETE CASCADE
        )
    ''')

    # Tabela de Manutenções da Frota
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS manutencoes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            data_servico TEXT NOT NULL,
            tipo TEXT DEFAULT 'Preventiva',
            viatura_prefixo TEXT,
            viatura_placa TEXT,
            viatura_modelo TEXT,
            viatura_marca TEXT,
            cidade_carga TEXT,
            odometro REAL DEFAULT 0.0,
            oficina TEXT,
            descricao_servico TEXT NOT NULL,
            pecas TEXT,
            valor_total REAL DEFAULT 0.0,
            status TEXT DEFAULT 'Concluída',
            data_conclusao TEXT,
            responsavel TEXT,
            observacoes TEXT,
            tipo_documento_fiscal TEXT DEFAULT 'nfse',
            numero_danfe TEXT,
            valor_danfe REAL DEFAULT 0.0,
            numero_nfse TEXT,
            valor_nfse REAL DEFAULT 0.0,
            empresa_nome TEXT,
            empresa_cnpj TEXT,
            responsavel_nome TEXT DEFAULT 'JULIANO INOCENTE VELEDA',
            responsavel_grad TEXT DEFAULT 'Sd PM',
            responsavel_id_func TEXT DEFAULT '3155358',
            responsavel_cargo TEXT DEFAULT 'Aux P4/46º BPM',
            arquivo_os TEXT,
            arquivo_danfe TEXT,
            arquivo_nfse TEXT,
            arquivo_atestado_pdf TEXT,
            data_cadastro TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')

    # Tabela de Arquivos e Documentos Anexados por Viatura (OS, DANFE, NFS-e, Atestados)
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS documentos_manutencao (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            viatura_prefixo TEXT NOT NULL,
            manutencao_id INTEGER,
            tipo_documento TEXT NOT NULL,
            nome_original TEXT NOT NULL,
            nome_arquivo TEXT NOT NULL,
            caminho_relativo TEXT NOT NULL,
            tamanho_bytes INTEGER DEFAULT 0,
            descricao TEXT,
            pasta_os TEXT,
            data_upload TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (manutencao_id) REFERENCES manutencoes (id) ON DELETE SET NULL
        )
    ''')

    # Migração para coluna pasta_os em documentos_manutencao
    try:
        cursor.execute("ALTER TABLE documentos_manutencao ADD COLUMN pasta_os TEXT")
    except Exception:
        pass

    # Migração segura para colunas adicionais em manutencoes caso a tabela já existisse
    colunas_manut = [
        ('viatura_marca', 'TEXT'),
        ('cidade_carga', 'TEXT'),
        ('tipo_documento_fiscal', "TEXT DEFAULT 'nfse'"),
        ('numero_danfe', 'TEXT'),
        ('valor_danfe', 'REAL DEFAULT 0.0'),
        ('numero_nfse', 'TEXT'),
        ('valor_nfse', 'REAL DEFAULT 0.0'),
        ('empresa_nome', 'TEXT'),
        ('empresa_cnpj', 'TEXT'),
        ('responsavel_nome', "TEXT DEFAULT 'JULIANO INOCENTE VELEDA'"),
        ('responsavel_grad', "TEXT DEFAULT 'Sd PM'"),
        ('responsavel_id_func', "TEXT DEFAULT '3155358'"),
        ('responsavel_cargo', "TEXT DEFAULT 'Aux P4/46º BPM'"),
        ('arquivo_os', 'TEXT'),
        ('arquivo_danfe', 'TEXT'),
        ('arquivo_nfse', 'TEXT'),
        ('arquivo_atestado_pdf', 'TEXT'),
        ('numero_os', 'TEXT')
    ]
    for col_nome, col_tipo in colunas_manut:
        try:
            cursor.execute(f"ALTER TABLE manutencoes ADD COLUMN {col_nome} {col_tipo}")
        except Exception:
            pass

    # Migração para veiculos_alvo
    for col_nome, col_tipo in [('marca', 'TEXT'), ('cidade_carga', 'TEXT'), ('cartao_abastecimento', 'TEXT')]:
        try:
            cursor.execute(f"ALTER TABLE veiculos_alvo ADD COLUMN {col_nome} {col_tipo}")
        except Exception:
            pass

    # Migração para documentos_manutencao (subpastas de OS)
    try:
        cursor.execute("ALTER TABLE documentos_manutencao ADD COLUMN pasta_os TEXT")
    except Exception:
        pass

    # Tabela de Equipamentos do 46º BPM
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS equipamentos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            patrimonio TEXT UNIQUE NOT NULL,
            nome TEXT NOT NULL,
            tipo TEXT DEFAULT 'Outro',
            modelo TEXT,
            serie TEXT,
            viatura_prefixo TEXT,
            fracao TEXT DEFAULT '46º BPM',
            status TEXT DEFAULT 'Operacional',
            ultima_revisao TEXT,
            observacoes TEXT,
            data_cadastro TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')

    # Tabela de Gestão de Pneus
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS pneus (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            numero_fogo TEXT UNIQUE NOT NULL,
            marca TEXT,
            modelo TEXT,
            medida TEXT,
            dot TEXT,
            viatura_prefixo TEXT,
            posicao TEXT DEFAULT 'Estoque',
            km_instalacao REAL DEFAULT 0.0,
            km_atual REAL DEFAULT 0.0,
            sulco_mm REAL DEFAULT 8.0,
            status TEXT DEFAULT 'Bom',
            data_instalacao TEXT,
            observacoes TEXT,
            data_cadastro TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')

    # Tabela de Histórico / Movimentação de Pneus
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS movimentacoes_pneus (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            pneu_id INTEGER,
            data_movimentacao TEXT NOT NULL,
            tipo TEXT DEFAULT 'Medição de Sulco',
            viatura TEXT,
            posicao TEXT,
            odometro REAL,
            sulco_mm REAL,
            observacao TEXT,
            responsavel TEXT,
            FOREIGN KEY (pneu_id) REFERENCES pneus (id) ON DELETE CASCADE
        )
    ''')

    # Índices para consultas rápidas
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_abast_placa ON abastecimentos(placa)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_abast_prefixo ON abastecimentos(prefixo)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_abast_data ON abastecimentos(data_abastecimento)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_abast_importacao ON abastecimentos(importacao_id)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_manut_prefixo ON manutencoes(viatura_prefixo)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_manut_data ON manutencoes(data_servico)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_pneu_prefixo ON pneus(viatura_prefixo)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_pneu_fogo ON pneus(numero_fogo)')

    conn.commit()
    conn.close()

    # Inicializar dados padrão se necessário
    popular_dados_iniciais_manutencao_pneus()
    popular_viaturas_padrao_46bpm()


def salvar_veiculos_alvo(lista_veiculos: List[Dict[str, Any]]) -> int:
    """
    Insere ou atualiza viaturas na tabela veiculos_alvo.
    lista_veiculos: lista de dicts com 'placa', 'prefixo', 'modelo', 'fracao', 'tipo_combustivel'
    """
    conn = get_connection()
    cursor = conn.cursor()
    inseridos_ou_atualizados = 0

    for v in lista_veiculos:
        placa = str(v.get('placa', '')).strip().upper().replace('-', '').replace(' ', '')
        if not placa:
            continue
        prefixo = str(v.get('prefixo', '')).strip()
        modelo = str(v.get('modelo', '')).strip()
        fracao = str(v.get('fracao', '46º BPM')).strip()
        tipo_comb = str(v.get('tipo_combustivel', 'Gasolina Comum')).strip()

        marca = str(v.get('marca', '')).strip()
        cidade_carga = str(v.get('cidade_carga', 'Bagé')).strip()

        cursor.execute('''
            INSERT INTO veiculos_alvo (placa, prefixo, marca, modelo, cidade_carga, fracao, tipo_combustivel, ativo)
            VALUES (?, ?, ?, ?, ?, ?, ?, 1)
            ON CONFLICT(placa) DO UPDATE SET
                prefixo = excluded.prefixo,
                marca = CASE WHEN excluded.marca != '' THEN excluded.marca ELSE veiculos_alvo.marca END,
                modelo = CASE WHEN excluded.modelo != '' THEN excluded.modelo ELSE veiculos_alvo.modelo END,
                cidade_carga = CASE WHEN excluded.cidade_carga != '' THEN excluded.cidade_carga ELSE veiculos_alvo.cidade_carga END,
                fracao = CASE WHEN excluded.fracao != '' THEN excluded.fracao ELSE veiculos_alvo.fracao END,
                tipo_combustivel = excluded.tipo_combustivel,
                ativo = 1
        ''', (placa, prefixo, marca, modelo, cidade_carga, fracao, tipo_comb))
        inseridos_ou_atualizados += 1

    conn.commit()
    conn.close()
    return inseridos_ou_atualizados


def obter_veiculos_alvo(apenas_ativos: bool = True) -> List[Dict[str, Any]]:
    """Retorna a lista de viaturas alvo cadastradas no SQLite."""
    conn = get_connection()
    cursor = conn.cursor()
    query = 'SELECT * FROM veiculos_alvo'
    if apenas_ativos:
        query += ' WHERE ativo = 1'
    query += ' ORDER BY CASE WHEN prefixo GLOB "[0-9]*" THEN 0 ELSE 1 END, CAST(prefixo AS INTEGER), prefixo, placa'
    cursor.execute(query)
    rows = [dict(r) for r in cursor.fetchall()]
    conn.close()
    return rows


def obter_veiculos_com_abastecimento(importacao_id: Optional[int] = None) -> List[Dict[str, Any]]:
    """
    Retorna apenas as viaturas que possuem ao menos 1 abastecimento registrado no período/banco.
    """
    conn = get_connection()
    cursor = conn.cursor()
    where_clause = ""
    params = []
    if importacao_id:
        where_clause = " WHERE a.importacao_id = ?"
        params.append(importacao_id)

    query = f'''
        SELECT 
            a.placa,
            COALESCE(v.prefixo, a.prefixo, a.placa) as prefixo,
            COALESCE(v.modelo, a.modelo, 'Não informado') as modelo,
            COALESCE(v.fracao, a.fracao, '46º BPM') as fracao,
            COUNT(a.id) as total_abastecimentos
        FROM abastecimentos a
        LEFT JOIN veiculos_alvo v ON a.placa = v.placa
        {where_clause}
        GROUP BY a.placa
        ORDER BY prefixo, a.placa
    '''
    cursor.execute(query, params)
    rows = [dict(r) for r in cursor.fetchall()]
    conn.close()
    return rows


def salvar_lote_importacao(meta: Dict[str, Any], lista_abastecimentos: List[Dict[str, Any]]) -> int:
    """
    Registra um novo lote de importação e todos os abastecimentos filtrados correspondentes.
    Retorna o importacao_id gerado.
    """
    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute('''
        INSERT INTO importacoes (
            arquivo_rs, arquivo_alvo, total_linhas_rs, total_filtrado,
            total_veiculos, total_litros, total_valor, total_km, media_km_l
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (
        meta.get('arquivo_rs', ''),
        meta.get('arquivo_alvo', ''),
        meta.get('total_linhas_rs', 0),
        meta.get('total_filtrado', 0),
        meta.get('total_veiculos', 0),
        meta.get('total_litros', 0.0),
        meta.get('total_valor', 0.0),
        meta.get('total_km', 0.0),
        meta.get('media_km_l', 0.0)
    ))

    importacao_id = cursor.lastrowid

    for item in lista_abastecimentos:
        cursor.execute('''
            INSERT INTO abastecimentos (
                importacao_id, data_abastecimento, placa, prefixo, modelo, fracao,
                odometro, km_anterior, km_rodado, litros, consumo_km_l,
                valor_unitario, valor_total, custo_km, tipo_combustivel,
                posto, cidade, motorista, anomalia, anomalia_obs
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            importacao_id,
            item.get('data_abastecimento'),
            item.get('placa'),
            item.get('prefixo'),
            item.get('modelo'),
            item.get('fracao'),
            item.get('odometro'),
            item.get('km_anterior'),
            item.get('km_rodado'),
            item.get('litros'),
            item.get('consumo_km_l'),
            item.get('valor_unitario'),
            item.get('valor_total'),
            item.get('custo_km'),
            item.get('tipo_combustivel'),
            item.get('posto'),
            item.get('cidade'),
            item.get('motorista'),
            item.get('anomalia', 0),
            item.get('anomalia_obs', '')
        ))

    conn.commit()
    conn.close()
    return importacao_id


def obter_estatisticas_gerais(importacao_id: Optional[int] = None) -> Dict[str, Any]:
    """Retorna os KPIs agregados (totais de viaturas, litros, km, valor e média km/l)."""
    conn = get_connection()
    cursor = conn.cursor()

    where_clause = ""
    params = []
    if importacao_id:
        where_clause = " WHERE importacao_id = ?"
        params.append(importacao_id)

    query = f'''
        SELECT 
            COUNT(DISTINCT placa) as total_veiculos,
            COUNT(*) as total_abastecimentos,
            COALESCE(SUM(litros), 0) as total_litros,
            COALESCE(SUM(valor_total), 0) as total_valor,
            COALESCE(SUM(CASE WHEN km_rodado > 0 THEN km_rodado ELSE 0 END), 0) as total_km,
            COALESCE(SUM(CASE WHEN anomalia = 1 THEN 1 ELSE 0 END), 0) as total_anomalias
        FROM abastecimentos
        {where_clause}
    '''
    cursor.execute(query, params)
    row = dict(cursor.fetchone())

    # Consumo médio ponderado = total_km / total_litros_consumidos
    # Para o cálculo de consumo, consideramos litros de trechos onde houve km_rodado válido
    query_consumo = f'''
        SELECT 
            SUM(km_rodado) as km_valido,
            SUM(litros) as litros_validos
        FROM abastecimentos
        WHERE km_rodado > 0 AND litros > 0 AND anomalia = 0
        {"AND importacao_id = ?" if importacao_id else ""}
    '''
    cursor.execute(query_consumo, params)
    row_c = dict(cursor.fetchone())
    km_v = row_c.get('km_valido') or 0
    lt_v = row_c.get('litros_validos') or 0
    row['media_km_l'] = round(km_v / lt_v, 2) if lt_v > 0 else 0.0

    # Custo médio por km
    row['custo_medio_km'] = round(row['total_valor'] / km_v, 2) if km_v > 0 else 0.0

    conn.close()
    return row


def obter_resumo_por_viatura(importacao_id: Optional[int] = None) -> List[Dict[str, Any]]:
    """
    Retorna o consolidado por viatura:
    placa, prefixo, modelo, fracao, total abastecimentos, km inicial, km final, km total,
    total litros, consumo médio km/l, gasto total R$, custo R$/km, anomalias.
    """
    conn = get_connection()
    cursor = conn.cursor()

    where_clause = ""
    params = []
    if importacao_id:
        where_clause = " WHERE a.importacao_id = ?"
        params.append(importacao_id)

    query = f'''
        SELECT 
            a.placa,
            COALESCE(v.prefixo, a.prefixo) as prefixo,
            COALESCE(v.modelo, a.modelo, 'Não informado') as modelo,
            COALESCE(v.fracao, a.fracao, '46º BPM') as fracao,
            COUNT(a.id) as qtd_abastecimentos,
            MIN(a.odometro) as km_inicial,
            MAX(a.odometro) as km_final,
            COALESCE(SUM(CASE WHEN a.km_rodado > 0 THEN a.km_rodado ELSE 0 END), 0) as km_total_rodado,
            COALESCE(SUM(a.litros), 0) as total_litros,
            COALESCE(SUM(a.valor_total), 0) as total_valor,
            SUM(CASE WHEN a.anomalia = 1 THEN 1 ELSE 0 END) as qtd_anomalias
        FROM abastecimentos a
        LEFT JOIN veiculos_alvo v ON a.placa = v.placa
        {where_clause}
        GROUP BY a.placa
        ORDER BY prefixo, a.placa
    '''
    cursor.execute(query, params)
    rows = [dict(r) for r in cursor.fetchall()]

    for r in rows:
        # Cálculo de consumo médio km/l por viatura
        km = r.get('km_total_rodado', 0)
        litros = r.get('total_litros', 0)
        valor = r.get('total_valor', 0)

        # Buscar litros apenas dos abastecimentos com km_rodado válido para precisão
        cursor.execute(f'''
            SELECT SUM(litros) as litros_uteis
            FROM abastecimentos
            WHERE placa = ? AND km_rodado > 0 AND anomalia = 0
            {"AND importacao_id = ?" if importacao_id else ""}
        ''', [r['placa']] + (params if importacao_id else []))
        row_uteis = cursor.fetchone()
        litros_uteis = (row_uteis[0] if row_uteis and row_uteis[0] else 0)

        if litros_uteis > 0 and km > 0:
            r['consumo_medio_km_l'] = round(km / litros_uteis, 2)
        elif litros > 0 and km > 0:
            r['consumo_medio_km_l'] = round(km / litros, 2)
        else:
            r['consumo_medio_km_l'] = 0.0

        r['custo_km'] = round(valor / km, 2) if km > 0 else 0.0

    conn.close()
    return rows


def obter_abastecimentos_viatura(placa_ou_prefixo: str, importacao_id: Optional[int] = None) -> List[Dict[str, Any]]:
    """Retorna todos os abastecimentos cronológicos de uma viatura específica."""
    conn = get_connection()
    cursor = conn.cursor()

    where_clause = " WHERE (a.placa = ? OR a.prefixo = ?)"
    params = [placa_ou_prefixo, placa_ou_prefixo]
    if importacao_id:
        where_clause += " AND a.importacao_id = ?"
        params.append(importacao_id)

    query = f'''
        SELECT 
            a.*,
            COALESCE(v.modelo, a.modelo) as modelo_veiculo,
            COALESCE(v.fracao, a.fracao) as fracao_veiculo
        FROM abastecimentos a
        LEFT JOIN veiculos_alvo v ON a.placa = v.placa
        {where_clause}
        ORDER BY a.data_abastecimento ASC, a.odometro ASC
    '''
    cursor.execute(query, params)
    rows = [dict(r) for r in cursor.fetchall()]
    conn.close()
    return rows


def obter_todos_abastecimentos(importacao_id: Optional[int] = None, limit: int = 1000) -> List[Dict[str, Any]]:
    """Retorna lista de todos os abastecimentos filtrados."""
    conn = get_connection()
    cursor = conn.cursor()

    where_clause = ""
    params = []
    if importacao_id:
        where_clause = " WHERE importacao_id = ?"
        params.append(importacao_id)

    query = f'''
        SELECT * FROM abastecimentos
        {where_clause}
        ORDER BY data_abastecimento DESC, id DESC
        LIMIT {limit}
    '''
    cursor.execute(query, params)
    rows = [dict(r) for r in cursor.fetchall()]
    conn.close()
    return rows


def obter_anomalias(importacao_id: Optional[int] = None) -> List[Dict[str, Any]]:
    """Retorna abastecimentos marcados com possíveis anomalias de odômetro ou consumo."""
    conn = get_connection()
    cursor = conn.cursor()

    where_clause = " WHERE anomalia = 1"
    params = []
    if importacao_id:
        where_clause += " AND importacao_id = ?"
        params.append(importacao_id)

    query = f'''
        SELECT * FROM abastecimentos
        {where_clause}
        ORDER BY data_abastecimento DESC
    '''
    cursor.execute(query, params)
    rows = [dict(r) for r in cursor.fetchall()]
    conn.close()
    return rows


def obter_historico_importacoes() -> List[Dict[str, Any]]:
    """Retorna todos os lotes de importação registrados."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM importacoes ORDER BY id DESC')
    rows = [dict(r) for r in cursor.fetchall()]
    conn.close()
    return rows


def limpar_todos_os_dados():
    """Limpa todos os dados das tabelas para reiniciar do zero."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute('DELETE FROM abastecimentos')
    cursor.execute('DELETE FROM importacoes')
    cursor.execute('DELETE FROM veiculos_alvo')
    conn.commit()
    conn.close()


def obter_abastecimentos_por_veiculos(veiculos: List[str]) -> List[Dict[str, Any]]:
    """Retorna todos os abastecimentos cujas viaturas têm placa ou prefixo na lista fornecida.

    Args:
        veiculos: lista de identificadores (placas ou prefixos, ex.: ['IUX9E18', 'BM-12450']).
    """
    if not veiculos:
        return []
    veiculos_limpos = [str(v).strip() for v in veiculos if str(v).strip()]
    if not veiculos_limpos:
        return []
    conn = get_connection()
    cursor = conn.cursor()
    placeholders = ','.join('?' for _ in veiculos_limpos)
    query = f'''SELECT * FROM abastecimentos
               WHERE placa IN ({placeholders}) OR prefixo IN ({placeholders})
               ORDER BY COALESCE(prefixo, placa), data_abastecimento ASC, id ASC'''
    cursor.execute(query, veiculos_limpos + veiculos_limpos)
    rows = [dict(r) for r in cursor.fetchall()]
    conn.close()
    return rows


def obter_abastecimentos_por_prefixos(prefixos: List[str]) -> List[Dict[str, Any]]:
    """Compatibilidade retroativa: chama obter_abastecimentos_por_veiculos."""
    return obter_abastecimentos_por_veiculos(prefixos)


# ==============================================================================
# MÓDULO 2: GESTÃO DE MANUTENÇÃO E EQUIPAMENTOS
# ==============================================================================

def obter_manutencoes(filtro_status: Optional[str] = None, filtro_viatura: Optional[str] = None) -> List[Dict[str, Any]]:
    """Retorna lista de manutenções registradas com filtros opcionais."""
    conn = get_connection()
    cursor = conn.cursor()
    query = 'SELECT * FROM manutencoes WHERE 1=1'
    params = []
    if filtro_status and filtro_status != 'Todos':
        query += ' AND status = ?'
        params.append(filtro_status)
    if filtro_viatura and filtro_viatura != 'Todos':
        query += ' AND (viatura_prefixo = ? OR viatura_placa = ?)'
        params.extend([filtro_viatura, filtro_viatura])
    query += ' ORDER BY data_servico DESC, id DESC'
    cursor.execute(query, params)
    rows = [dict(r) for r in cursor.fetchall()]
    conn.close()
    return rows


def salvar_manutencao(dados: Dict[str, Any]) -> int:
    """Insere ou atualiza um registro de manutenção/ordem de serviço."""
    conn = get_connection()
    cursor = conn.cursor()
    m_id = dados.get('id')
    
    # Tratamento de valores fiscais e totais
    val_tot = float(dados.get('valor_total') or 0.0)
    val_danfe = float(dados.get('valor_danfe') or 0.0)
    val_nfse = float(dados.get('valor_nfse') or 0.0)
    if val_tot <= 0.0:
        val_tot = val_danfe + val_nfse

    if m_id:
        cursor.execute('''
            UPDATE manutencoes SET
                data_servico = COALESCE(?, data_servico),
                tipo = COALESCE(?, tipo),
                viatura_prefixo = COALESCE(?, viatura_prefixo),
                viatura_placa = COALESCE(?, viatura_placa),
                viatura_modelo = COALESCE(?, viatura_modelo),
                viatura_marca = COALESCE(?, viatura_marca),
                cidade_carga = COALESCE(?, cidade_carga),
                odometro = CASE WHEN ? IS NOT NULL THEN ? ELSE odometro END,
                oficina = COALESCE(?, oficina),
                descricao_servico = COALESCE(?, descricao_servico),
                pecas = COALESCE(?, pecas),
                valor_total = CASE WHEN ? > 0 THEN ? ELSE valor_total END,
                status = COALESCE(?, status),
                data_conclusao = COALESCE(?, data_conclusao),
                responsavel = COALESCE(?, responsavel),
                observacoes = COALESCE(?, observacoes),
                tipo_documento_fiscal = COALESCE(?, tipo_documento_fiscal),
                numero_danfe = COALESCE(?, numero_danfe),
                valor_danfe = CASE WHEN ? > 0 THEN ? ELSE valor_danfe END,
                numero_nfse = COALESCE(?, numero_nfse),
                valor_nfse = CASE WHEN ? > 0 THEN ? ELSE valor_nfse END,
                empresa_nome = COALESCE(?, empresa_nome),
                empresa_cnpj = COALESCE(?, empresa_cnpj),
                responsavel_nome = COALESCE(?, responsavel_nome),
                responsavel_grad = COALESCE(?, responsavel_grad),
                responsavel_id_func = COALESCE(?, responsavel_id_func),
                responsavel_cargo = COALESCE(?, responsavel_cargo),
                arquivo_os = COALESCE(?, arquivo_os),
                arquivo_danfe = COALESCE(?, arquivo_danfe),
                arquivo_nfse = COALESCE(?, arquivo_nfse),
                arquivo_atestado_pdf = COALESCE(?, arquivo_atestado_pdf),
                numero_os = COALESCE(?, numero_os)
            WHERE id = ?
        ''', (
            dados.get('data_servico'), dados.get('tipo'), dados.get('viatura_prefixo'),
            dados.get('viatura_placa'), dados.get('viatura_modelo'), dados.get('viatura_marca'),
            dados.get('cidade_carga'), dados.get('odometro'), float(dados.get('odometro') or 0.0),
            dados.get('oficina') or dados.get('empresa_nome'), dados.get('descricao_servico'), dados.get('pecas'),
            val_tot, val_tot, dados.get('status'), dados.get('data_conclusao'),
            dados.get('responsavel') or dados.get('responsavel_nome'), dados.get('observacoes'),
            dados.get('tipo_documento_fiscal'), dados.get('numero_danfe'), val_danfe, val_danfe,
            dados.get('numero_nfse'), val_nfse, val_nfse, dados.get('empresa_nome') or dados.get('oficina'),
            dados.get('empresa_cnpj'), dados.get('responsavel_nome'),
            dados.get('responsavel_grad'), dados.get('responsavel_id_func'),
            dados.get('responsavel_cargo'),
            dados.get('arquivo_os'), dados.get('arquivo_danfe'), dados.get('arquivo_nfse'), dados.get('arquivo_atestado_pdf'),
            dados.get('numero_os'),
            m_id
        ))
        res_id = int(m_id)
    else:
        cursor.execute('''
            INSERT INTO manutencoes (
                data_servico, tipo, viatura_prefixo, viatura_placa, viatura_modelo,
                viatura_marca, cidade_carga, odometro, oficina, descricao_servico,
                pecas, valor_total, status, data_conclusao, responsavel, observacoes,
                tipo_documento_fiscal, numero_danfe, valor_danfe, numero_nfse, valor_nfse,
                empresa_nome, empresa_cnpj, responsavel_nome, responsavel_grad,
                responsavel_id_func, responsavel_cargo, arquivo_os, arquivo_danfe,
                arquivo_nfse, arquivo_atestado_pdf, numero_os
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            dados.get('data_servico'), dados.get('tipo', 'Preventiva'), dados.get('viatura_prefixo'),
            dados.get('viatura_placa'), dados.get('viatura_modelo'), dados.get('viatura_marca'),
            dados.get('cidade_carga'), float(dados.get('odometro') or 0.0),
            dados.get('oficina') or dados.get('empresa_nome'), dados.get('descricao_servico'),
            dados.get('pecas'), val_tot, dados.get('status', 'Concluída'),
            dados.get('data_conclusao'), dados.get('responsavel') or dados.get('responsavel_nome'),
            dados.get('observacoes'), dados.get('tipo_documento_fiscal', 'nfse'),
            dados.get('numero_danfe'), val_danfe, dados.get('numero_nfse'), val_nfse,
            dados.get('empresa_nome') or dados.get('oficina'), dados.get('empresa_cnpj'),
            dados.get('responsavel_nome', 'JULIANO INOCENTE VELEDA'),
            dados.get('responsavel_grad', 'Sd PM'), dados.get('responsavel_id_func', '3155358'),
            dados.get('responsavel_cargo', 'Aux P4/46º BPM'),
            dados.get('arquivo_os'), dados.get('arquivo_danfe'), dados.get('arquivo_nfse'),
            dados.get('arquivo_atestado_pdf'), dados.get('numero_os')
        ))
        res_id = cursor.lastrowid
    conn.commit()
    conn.close()
    return res_id


def obter_manutencao_por_id(manutencao_id: int) -> Optional[Dict[str, Any]]:
    """Retorna uma manutenção específica pelo ID."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM manutencoes WHERE id = ?', (manutencao_id,))
    row = cursor.fetchone()
    conn.close()
    return dict(row) if row else None


def atualizar_atestado_manutencao(manutencao_id: int, caminho_pdf: str) -> None:
    """Atualiza o caminho do arquivo PDF do atestado no registro da manutenção."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute('UPDATE manutencoes SET arquivo_atestado_pdf = ? WHERE id = ?', (caminho_pdf, manutencao_id))
    conn.commit()
    conn.close()


# ==============================================================================
# GESTÃO DE PASTAS E DOCUMENTOS POR VIATURA
# ==============================================================================

def obter_todas_pastas_viaturas() -> List[Dict[str, Any]]:
    """
    Retorna a relação de todas as viaturas conhecidas (cadastradas ou com histórico),
    agrupadas com contagem de documentos na pasta, manutenções e cidade de carga.
    """
    conn = get_connection()
    cursor = conn.cursor()
    
    # 1. Busca todas as viaturas de veiculos_alvo
    cursor.execute('''
        SELECT 
            COALESCE(prefixo, placa) as prefixo,
            placa,
            modelo,
            COALESCE(marca, '') as marca,
            COALESCE(cidade_carga, 'Bagé') as cidade_carga
        FROM veiculos_alvo
        WHERE ativo = 1
    ''')
    viaturas_map: Dict[str, Dict[str, Any]] = {}
    for r in cursor.fetchall():
        d = dict(r)
        pref = str(d['prefixo']).strip().upper()
        viaturas_map[pref] = d

    # 2. Adiciona viaturas que constam em manutenções ou documentos
    cursor.execute('''
        SELECT DISTINCT viatura_prefixo, viatura_placa, viatura_modelo, viatura_marca, cidade_carga
        FROM manutencoes
        WHERE viatura_prefixo IS NOT NULL AND viatura_prefixo != ''
    ''')
    for r in cursor.fetchall():
        d = dict(r)
        pref = str(d['viatura_prefixo']).strip().upper()
        if pref not in viaturas_map:
            viaturas_map[pref] = {
                'prefixo': pref,
                'placa': d.get('viatura_placa') or '',
                'modelo': d.get('viatura_modelo') or 'Viatura BM',
                'marca': d.get('viatura_marca') or '',
                'cidade_carga': d.get('cidade_carga') or 'Bagé'
            }

    # 3. Contagem de documentos e manutenções por viatura
    resultado = []
    for pref, v in viaturas_map.items():
        cursor.execute('SELECT COUNT(*) FROM documentos_manutencao WHERE viatura_prefixo = ?', (pref,))
        total_docs = cursor.fetchone()[0]

        cursor.execute('SELECT COUNT(*), SUM(valor_total) FROM manutencoes WHERE viatura_prefixo = ?', (pref,))
        row_manut = cursor.fetchone()
        total_manut = row_manut[0] or 0
        total_gasto = row_manut[1] or 0.0

        v_info = dict(v)
        v_info['total_arquivos'] = total_docs
        v_info['total_manutencoes'] = total_manut
        v_info['total_gasto'] = total_gasto
        resultado.append(v_info)

    conn.close()
    def _chave_prefixo(x):
        pref = str(x.get('prefixo', '')).strip()
        return (0, int(pref)) if pref.isdigit() else (1, pref)
    resultado.sort(key=_chave_prefixo)
    return resultado


def obter_arquivos_viatura(viatura_prefixo: str) -> List[Dict[str, Any]]:
    """Retorna todos os documentos anexados na pasta de uma viatura."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute('''
        SELECT d.*, m.descricao_servico, m.oficina
        FROM documentos_manutencao d
        LEFT JOIN manutencoes m ON d.manutencao_id = m.id
        WHERE d.viatura_prefixo = ?
        ORDER BY d.data_upload DESC, d.id DESC
    ''', (str(viatura_prefixo).strip().upper(),))
    rows = [dict(r) for r in cursor.fetchall()]
    conn.close()
    return rows


def salvar_documento_viatura(
    viatura_prefixo: str,
    tipo_documento: str,
    nome_original: str,
    nome_arquivo: str,
    caminho_relativo: str,
    tamanho_bytes: int = 0,
    descricao: Optional[str] = None,
    manutencao_id: Optional[int] = None,
    pasta_os: Optional[str] = None
) -> int:
    """Registra um documento na pasta da viatura no SQLite."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute('''
        INSERT INTO documentos_manutencao (
            viatura_prefixo, manutencao_id, tipo_documento,
            nome_original, nome_arquivo, caminho_relativo,
            tamanho_bytes, descricao, pasta_os
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (
        str(viatura_prefixo).strip().upper(), manutencao_id, tipo_documento.upper(),
        nome_original, nome_arquivo, caminho_relativo,
        tamanho_bytes, descricao, pasta_os
    ))
    doc_id = cursor.lastrowid
    conn.commit()
    conn.close()
    return doc_id


def obter_documento_por_id(doc_id: int) -> Optional[Dict[str, Any]]:
    """Retorna dados de um documento específico pelo ID."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM documentos_manutencao WHERE id = ?', (doc_id,))
    row = cursor.fetchone()
    conn.close()
    return dict(row) if row else None


def excluir_documento_viatura(doc_id: int) -> Optional[str]:
    """
    Remove o registro do documento no SQLite e retorna o caminho_relativo
    para que o arquivo físico seja removido do disco.
    """
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute('SELECT caminho_relativo FROM documentos_manutencao WHERE id = ?', (doc_id,))
    row = cursor.fetchone()
    caminho = row['caminho_relativo'] if row else None
    if caminho:
        cursor.execute('DELETE FROM documentos_manutencao WHERE id = ?', (doc_id,))
        conn.commit()
    conn.close()
    return caminho


def excluir_manutencao(manutencao_id: int) -> bool:
    """Exclui um registro de manutenção pelo ID."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute('DELETE FROM manutencoes WHERE id = ?', (manutencao_id,))
    afetados = cursor.rowcount
    conn.commit()
    conn.close()
    return afetados > 0


def obter_estatisticas_manutencao() -> Dict[str, Any]:
    """Retorna KPIs gerais do módulo de manutenção."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute('SELECT COUNT(*) as total, SUM(valor_total) as custo_total FROM manutencoes')
    row_geral = dict(cursor.fetchone() or {})
    cursor.execute("SELECT COUNT(*) as baixadas FROM manutencoes WHERE status IN ('Em Andamento', 'Aguardando Peças')")
    row_baixadas = dict(cursor.fetchone() or {})
    cursor.execute("SELECT COUNT(*) as concluidas FROM manutencoes WHERE status = 'Concluída'")
    row_concluidas = dict(cursor.fetchone() or {})
    cursor.execute("SELECT COUNT(*) as agendadas FROM manutencoes WHERE status = 'Agendada'")
    row_agendadas = dict(cursor.fetchone() or {})
    conn.close()
    return {
        'total': row_geral.get('total') or 0,
        'custo_total': row_geral.get('custo_total') or 0.0,
        'em_andamento': row_baixadas.get('baixadas') or 0,
        'concluidas': row_concluidas.get('concluidas') or 0,
        'agendadas': row_agendadas.get('agendadas') or 0
    }


def obter_equipamentos() -> List[Dict[str, Any]]:
    """Retorna todos os equipamentos patrimoniados do 46º BPM."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM equipamentos ORDER BY nome ASC')
    rows = [dict(r) for r in cursor.fetchall()]
    conn.close()
    return rows


def salvar_equipamento(dados: Dict[str, Any]) -> int:
    """Insere ou atualiza um equipamento do 46º BPM."""
    conn = get_connection()
    cursor = conn.cursor()
    eq_id = dados.get('id')
    if eq_id:
        cursor.execute('''
            UPDATE equipamentos SET
                patrimonio = ?, nome = ?, tipo = ?, modelo = ?, serie = ?,
                viatura_prefixo = ?, fracao = ?, status = ?, ultima_revisao = ?, observacoes = ?
            WHERE id = ?
        ''', (
            dados.get('patrimonio'), dados.get('nome'), dados.get('tipo', 'Outro'),
            dados.get('modelo'), dados.get('serie'), dados.get('viatura_prefixo'),
            dados.get('fracao', '46º BPM'), dados.get('status', 'Operacional'),
            dados.get('ultima_revisao'), dados.get('observacoes'), eq_id
        ))
        res_id = int(eq_id)
    else:
        cursor.execute('''
            INSERT INTO equipamentos (
                patrimonio, nome, tipo, modelo, serie, viatura_prefixo, fracao, status, ultima_revisao, observacoes
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            dados.get('patrimonio'), dados.get('nome'), dados.get('tipo', 'Outro'),
            dados.get('modelo'), dados.get('serie'), dados.get('viatura_prefixo'),
            dados.get('fracao', '46º BPM'), dados.get('status', 'Operacional'),
            dados.get('ultima_revisao'), dados.get('observacoes')
        ))
        res_id = cursor.lastrowid
    conn.commit()
    conn.close()
    return res_id


def excluir_equipamento(eq_id: int) -> bool:
    """Exclui um equipamento pelo ID."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute('DELETE FROM equipamentos WHERE id = ?', (eq_id,))
    afetados = cursor.rowcount
    conn.commit()
    conn.close()
    return afetados > 0


# ==============================================================================
# MÓDULO 3: GESTÃO DE PNEUS
# ==============================================================================

def obter_pneus(filtro_status: Optional[str] = None, filtro_viatura: Optional[str] = None) -> List[Dict[str, Any]]:
    """Retorna pneus cadastrados com filtros opcionais."""
    conn = get_connection()
    cursor = conn.cursor()
    query = 'SELECT * FROM pneus WHERE 1=1'
    params = []
    if filtro_status and filtro_status != 'Todos':
        query += ' AND status = ?'
        params.append(filtro_status)
    if filtro_viatura and filtro_viatura != 'Todos':
        query += ' AND viatura_prefixo = ?'
        params.append(filtro_viatura)
    query += ' ORDER BY viatura_prefixo ASC, posicao ASC'
    cursor.execute(query, params)
    rows = [dict(r) for r in cursor.fetchall()]
    conn.close()
    return rows


def salvar_pneu(dados: Dict[str, Any]) -> int:
    """Cadastra ou atualiza dados de um pneu."""
    conn = get_connection()
    cursor = conn.cursor()
    p_id = dados.get('id')
    sulco = float(dados.get('sulco_mm') or 8.0)
    status = dados.get('status')
    if not status:
        if sulco >= 6.0:
            status = 'Excelente'
        elif sulco >= 4.0:
            status = 'Bom'
        elif sulco >= 2.5:
            status = 'Regular'
        else:
            status = 'Crítico / Substituir'

    if p_id:
        cursor.execute('''
            UPDATE pneus SET
                numero_fogo = ?, marca = ?, modelo = ?, medida = ?, dot = ?,
                viatura_prefixo = ?, posicao = ?, km_instalacao = ?, km_atual = ?,
                sulco_mm = ?, status = ?, data_instalacao = ?, observacoes = ?
            WHERE id = ?
        ''', (
            dados.get('numero_fogo'), dados.get('marca'), dados.get('modelo'),
            dados.get('medida'), dados.get('dot'), dados.get('viatura_prefixo'),
            dados.get('posicao', 'Estoque'), float(dados.get('km_instalacao') or 0.0),
            float(dados.get('km_atual') or 0.0), sulco, status,
            dados.get('data_instalacao'), dados.get('observacoes'), p_id
        ))
        res_id = int(p_id)
    else:
        cursor.execute('''
            INSERT INTO pneus (
                numero_fogo, marca, modelo, medida, dot, viatura_prefixo, posicao,
                km_instalacao, km_atual, sulco_mm, status, data_instalacao, observacoes
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            dados.get('numero_fogo'), dados.get('marca'), dados.get('modelo'),
            dados.get('medida'), dados.get('dot'), dados.get('viatura_prefixo'),
            dados.get('posicao', 'Estoque'), float(dados.get('km_instalacao') or 0.0),
            float(dados.get('km_atual') or 0.0), sulco, status,
            dados.get('data_instalacao'), dados.get('observacoes')
        ))
        res_id = cursor.lastrowid
    conn.commit()
    conn.close()
    return res_id


def excluir_pneu(pneu_id: int) -> bool:
    """Exclui um pneu pelo ID."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute('DELETE FROM pneus WHERE id = ?', (pneu_id,))
    afetados = cursor.rowcount
    conn.commit()
    conn.close()
    return afetados > 0


def registrar_movimentacao_pneu(dados: Dict[str, Any]) -> int:
    """Registra uma movimentação, rodízio ou medição de sulco de pneu."""
    conn = get_connection()
    cursor = conn.cursor()
    p_id = dados.get('pneu_id')
    cursor.execute('''
        INSERT INTO movimentacoes_pneus (
            pneu_id, data_movimentacao, tipo, viatura, posicao, odometro, sulco_mm, observacao, responsavel
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (
        p_id, dados.get('data_movimentacao'), dados.get('tipo', 'Medição de Sulco'),
        dados.get('viatura'), dados.get('posicao'), float(dados.get('odometro') or 0.0),
        float(dados.get('sulco_mm') or 0.0), dados.get('observacao'), dados.get('responsavel')
    ))

    # Atualiza o pneu correspondente
    if p_id and dados.get('sulco_mm'):
        sulco = float(dados.get('sulco_mm'))
        if sulco >= 6.0:
            st = 'Excelente'
        elif sulco >= 4.0:
            st = 'Bom'
        elif sulco >= 2.5:
            st = 'Regular'
        else:
            st = 'Crítico / Substituir'
        cursor.execute('UPDATE pneus SET sulco_mm = ?, status = ? WHERE id = ?', (sulco, st, p_id))

    if p_id and dados.get('posicao'):
        cursor.execute('UPDATE pneus SET posicao = ?, viatura_prefixo = ? WHERE id = ?', (dados.get('posicao'), dados.get('viatura'), p_id))

    res_id = cursor.lastrowid
    conn.commit()
    conn.close()
    return res_id


def obter_estatisticas_pneus() -> Dict[str, Any]:
    """Retorna KPIs gerais do módulo de pneus."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute('SELECT COUNT(*) as total, AVG(sulco_mm) as media_sulco FROM pneus')
    row_geral = dict(cursor.fetchone() or {})
    cursor.execute("SELECT COUNT(*) as em_uso FROM pneus WHERE posicao != 'Estoque' AND status != 'Descartado'")
    row_uso = dict(cursor.fetchone() or {})
    cursor.execute("SELECT COUNT(*) as criticos FROM pneus WHERE sulco_mm < 2.5 OR status LIKE '%Crítico%'")
    row_crit = dict(cursor.fetchone() or {})
    cursor.execute("SELECT COUNT(*) as estoque FROM pneus WHERE posicao = 'Estoque'")
    row_est = dict(cursor.fetchone() or {})
    conn.close()
    return {
        'total': row_geral.get('total') or 0,
        'media_sulco': round(row_geral.get('media_sulco') or 0.0, 1),
        'em_uso': row_uso.get('em_uso') or 0,
        'criticos': row_crit.get('criticos') or 0,
        'estoque': row_est.get('estoque') or 0
    }


def popular_dados_iniciais_manutencao_pneus():
    """Inicializa dados realistas de demonstração para Manutenção e Pneus se estiverem vazios."""
    conn = get_connection()
    cursor = conn.cursor()
    
    # 1. Manutenções (inicia limpo para novas ordens de serviço reais da frota)
    pass

    # 2. Equipamentos (inicia limpo)
    pass

    # 3. Pneus
    cursor.execute('SELECT COUNT(*) FROM pneus')
    total_pneus = cursor.fetchone()[0]
    if total_pneus == 0:
        pneus_demo = [
            ('BM-46-101', 'Bridgestone', 'Dueler A/T', '265/65 R17', '1224', '12402', 'Dianteiro Esquerdo', 75000.0, 88500.0, 5.8, 'Bom', '2026-01-15', 'Desgaste uniforme.'),
            ('BM-46-102', 'Bridgestone', 'Dueler A/T', '265/65 R17', '1224', '12402', 'Dianteiro Direito', 75000.0, 88500.0, 5.7, 'Bom', '2026-01-15', 'Desgaste uniforme.'),
            ('BM-46-103', 'Bridgestone', 'Dueler A/T', '265/65 R17', '1224', '12402', 'Traseiro Esquerdo', 75000.0, 88500.0, 6.2, 'Excelente', '2026-01-15', 'Excelente estado.'),
            ('BM-46-104', 'Bridgestone', 'Dueler A/T', '265/65 R17', '1224', '12402', 'Traseiro Direito', 75000.0, 88500.0, 6.1, 'Excelente', '2026-01-15', 'Excelente estado.'),
            ('BM-46-105', 'Bridgestone', 'Dueler A/T', '265/65 R17', '1024', '12402', 'Estepe', 75000.0, 88500.0, 8.0, 'Excelente', '2026-01-15', 'Estepe novo.'),
            
            ('BM-46-201', 'Pirelli', 'Scorpion ATR', '215/65 R16', '3023', '12401', 'Dianteiro Esquerdo', 30000.0, 45200.0, 3.8, 'Regular', '2025-08-10', 'Próximo do rodízio programado.'),
            ('BM-46-202', 'Pirelli', 'Scorpion ATR', '215/65 R16', '3023', '12401', 'Dianteiro Direito', 30000.0, 45200.0, 3.6, 'Regular', '2025-08-10', 'Desgaste ligeiramente maior no ombro externo.'),
            ('BM-46-203', 'Pirelli', 'Scorpion ATR', '215/65 R16', '3023', '12401', 'Traseiro Esquerdo', 30000.0, 45200.0, 4.9, 'Bom', '2025-08-10', 'Sulco em boas condições.'),
            ('BM-46-204', 'Pirelli', 'Scorpion ATR', '215/65 R16', '3023', '12401', 'Traseiro Direito', 30000.0, 45200.0, 4.8, 'Bom', '2025-08-10', 'Sulco em boas condições.'),
            ('BM-46-205', 'Pirelli', 'Scorpion ATR', '215/65 R16', '3023', '12401', 'Estepe', 30000.0, 45200.0, 7.5, 'Excelente', '2025-08-10', 'Estepe.'),

            ('BM-46-301', 'Goodyear', 'EfficientGrip', '205/55 R16', '4522', '12408', 'Dianteiro Esquerdo', 42000.0, 58900.0, 2.1, 'Crítico / Substituir', '2025-04-12', 'Atingindo indicador TWI. Requer substituição urgente.'),
            ('BM-46-302', 'Goodyear', 'EfficientGrip', '205/55 R16', '4522', '12408', 'Dianteiro Direito', 42000.0, 58900.0, 2.0, 'Crítico / Substituir', '2025-04-12', 'Atingindo indicador TWI. Requer substituição urgente.'),

            ('BM-46-401', 'Michelin', 'LTX Force', '265/65 R17', '0525', None, 'Estoque', 0.0, 0.0, 8.5, 'Excelente', '2026-03-01', 'Pneu novo em reserva no almoxarifado do P4.'),
            ('BM-46-402', 'Michelin', 'LTX Force', '265/65 R17', '0525', None, 'Estoque', 0.0, 0.0, 8.5, 'Excelente', '2026-03-01', 'Pneu novo em reserva no almoxarifado do P4.')
        ]
        cursor.executemany('''
            INSERT INTO pneus (
                numero_fogo, marca, modelo, medida, dot, viatura_prefixo, posicao,
                km_instalacao, km_atual, sulco_mm, status, data_instalacao, observacoes
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', pneus_demo)

    conn.commit()
    conn.close()


def popular_viaturas_padrao_46bpm():
    """Popula a frota oficial completa do 46º BPM (atualizada em 31 AGO 26)."""
    conn = get_connection()
    cursor = conn.cursor()
    frota_oficial = [
        ('7879', 'JBM3576', 'AGRALE', 'MA 12.0 MAXIBUS MC', 'BAGÉ', '6035 7404 2195 1978', 'Diesel S10'),
        ('9532', 'IVB4468', 'NISSAN', 'FRONTIER SVATK 4X4', 'BAGÉ', '6035 7404 2197 5738', 'Diesel S10'),
        ('9565', 'IVB4F30', 'NISSAN', 'FRONTIER SVATK 4X4', 'BAGÉ', '6035 7404 4017 7696', 'Diesel S10'),
        ('9583', 'IVA4820', 'MARCOPOLO', 'VOLARE W8 ON', 'BAGÉ', '6035 7404 2198 0068', 'Diesel S10'),
        ('9766', 'IVI1825', 'YAMAHA', 'LANDER XTZ 250', 'BAGÉ', '6035 7404 2198 7295', 'Gasolina Comum'),
        ('10817', 'PAR2840', 'CHEVROLET', 'S10 LT DD4', 'BAGÉ', '6035 7404 2197 5290', 'Diesel S10'),
        ('10898', 'IYJ1287', 'TOYOTA', 'HILUX CDSRVA4GF', 'BAGÉ', '6035 7404 2199 2345', 'Diesel S10'),
        ('11041', 'IYJ6493', 'TOYOTA', 'COROLLA XEI 2.0', 'DOM PEDRITO', '6035 7404 2197 2784', 'Gasolina Comum'),
        ('11042', 'IYJ6447', 'TOYOTA', 'COROLLA XEI 2.0', 'BAGÉ', '6035 7404 2195 6258', 'Gasolina Comum'),
        ('11375', 'IUX9E18', 'CHEVROLET', 'S10 LTZ FD2', 'BAGÉ', '6035 7404 2196 8436', 'Diesel S10'),
        ('11489', 'IYS9069', 'FIAT', 'PÁLIO WEEK ADVENTURE', 'BAGÉ', '6035 7404 2199 1024', 'Gasolina Comum'),
        ('11509', 'IYS9341', 'FIAT', 'PÁLIO WEEK ADVENTURE', 'PINHEIRO MACHADO', '6035 7404 2195 4105', 'Gasolina Comum'),
        ('11560', 'IYT3841', 'FIAT', 'PÁLIO WEEK ADVENTURE', 'BAGÉ', '6035 7404 4325 1514', 'Gasolina Comum'),
        ('11565', 'IYT3896', 'FIAT', 'PÁLIO WEEK ADVENTURE', 'LAVRAS DO SUL', '6035 7404 2196 7222', 'Gasolina Comum'),
        ('11652', 'IYX4512', 'YAMAHA', 'LANDER XTZ 250', 'BAGÉ', '6035 7404 2199 9878', 'Gasolina Comum'),
        ('11772', 'IZG6F14', 'TOYOTA', 'HILUX CDSRVA2GF', 'BAGÉ', '6035 7404 2199 5926', 'Diesel S10'),
        ('11773', 'IZG6F20', 'TOYOTA', 'HILUX CDSRVA2GF', 'BAGÉ', '6035 7404 2198 8418', 'Diesel S10'),
        ('11774', 'IZG5E51', 'TOYOTA', 'HILUX CDSRVA2GF', 'BAGÉ', '6035 7404 2198 6156', 'Diesel S10'),
        ('11829', 'IZF4J45', 'FIAT', 'PÁLIO WEEK ADVENTURE', 'DOM PEDRITO', '6035 7404 2196 4567', 'Gasolina Comum'),
        ('11871', 'IZF8J81', 'FIAT', 'PÁLIO WEEK ADVENTURE', 'BAGÉ', '6035 7404 2198 7535', 'Gasolina Comum'),
        ('11873', 'IZF8E48', 'FIAT', 'PÁLIO WEEK ADVENTURE', 'BAGÉ', '6036 7404 2198 7592', 'Gasolina Comum'),
        ('12241', 'IZP1E09', 'FIAT', 'PÁLIO WEEK ADVENTURE', 'BAGÉ', '6035 7404 2198 8434', 'Gasolina Comum'),
        ('12300', 'IYZ1G45', 'REBOQUE', 'PREMIUM FLASH', 'BAGÉ', '', 'Outro'),
        ('12983', 'JBA3E79', 'RENAULT', 'DUSTER ZEN 16', 'BAGÉ', '6035 7404 4946 3899', 'Gasolina Comum'),
        ('12984', 'JBA3E93', 'RENAULT', 'DUSTER ZEN 16', 'BAGÉ', '6035 7404 2362 3765', 'Gasolina Comum'),
        ('13034', 'JBB8C49', 'TOYOTA', 'HILUX CDLOWM4FD', 'BAGÉ', '6035 7404 2414 3201', 'Diesel S10'),
        ('13436', 'JBM8H05', 'RENAULT', 'DUSTER ZEN 16', 'CANDIOTA', '6035 7404 2954 3702', 'Gasolina Comum'),
        ('13776', 'JBP8C25', 'RENAULT', 'DUSTER ZEN 16', 'ACEGUÁ', '6035 7404 3024 5958', 'Gasolina Comum'),
        ('13821', 'JBQ1D62', 'RENAULT', 'DUSTER ZEN 16', 'LAVRAS DO SUL', '6035 7404 5681 0412', 'Gasolina Comum'),
        ('13872', 'JBQ7D83', 'RENAULT', 'DUSTER ZEN 16', 'PEDRAS ALTAS', '6035 4704 3083 2045', 'Gasolina Comum'),
        ('13874', 'JBQ7E09', 'RENAULT', 'DUSTER ZEN 16', 'PINHEIRO MACHADO', '6035 7404 3083 2060', 'Gasolina Comum'),
        ('14034', 'JBV9J95', 'RENAULT', 'DUSTER ZEN 16', 'DOM PEDRITO', '6035 7404 3284 0426', 'Gasolina Comum'),
        ('14066', 'JBW2B98', 'RENAULT', 'DUSTER ZEN 16', 'PINHEIRO MACHADO', '6035 7404 3284 0756', 'Gasolina Comum'),
        ('14068', 'JBW2C30', 'RENAULT', 'DUSTER ZEN 16', 'BAGÉ', '6035 7404 3284 0327', 'Gasolina Comum'),
        ('14716', 'TQO0D63', 'TOYOTA', 'COROLLA CROSS', 'DOM PEDRITO', '6035 7404 4779 5060', 'Gasolina Comum'),
        ('14869', 'TQW0F80', 'TOYOTA', 'HILUX', 'DOM PEDRITO', '6035 7404 5162 7043', 'Diesel S10'),
        ('14895', 'TQU0F72', 'TOYOTA', 'HILUX', 'BAGÉ', '6035 7404 5131 7157', 'Diesel S10'),
        ('14916', 'TQW1G99', 'TOYOTA', 'HILUX', 'HULHA NEGRA', '6035 7404 5177 7129', 'Diesel S10'),
        ('14944', 'TQW2D96', 'TOYOTA', 'HILUX', 'LAVRAS DO SUL', '6035 7404 5177 3094', 'Diesel S10'),
        ('15025', 'TQW1H19', 'TOYOTA', 'HILUX', 'ACEGUÁ', '6035 7404 5177 3763', 'Diesel S10'),
        ('15413', 'TQX2C47', 'TOYOTA', 'COROLLA CROSS', 'BAGÉ', '6035 7407 5260 1112', 'Gasolina Comum'),
        ('15492', 'TQZ0B89', 'TOYOTA', 'COROLLA CROSS', 'PEDRAS ALTAS', '6035 7404 5272 6836', 'Gasolina Comum'),
        ('15853', 'MJM4J78', 'VOLKSWAGEN', 'SAVEIRO', 'BAGÉ', '6035 7404 5595 8157', 'Gasolina Comum')
    ]
    for pref, placa, marca, mod, cidade, cartao, combustivel in frota_oficial:
        cursor.execute('''
            INSERT INTO veiculos_alvo (prefixo, placa, marca, modelo, cidade_carga, cartao_abastecimento, fracao, tipo_combustivel, ativo)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, 1)
            ON CONFLICT(placa) DO UPDATE SET
                prefixo = excluded.prefixo,
                marca = excluded.marca,
                modelo = excluded.modelo,
                cidade_carga = excluded.cidade_carga,
                cartao_abastecimento = excluded.cartao_abastecimento,
                fracao = excluded.fracao,
                tipo_combustivel = excluded.tipo_combustivel,
                ativo = 1
        ''', (pref, placa, marca, mod, cidade, cartao, f"46º BPM - {cidade}", combustivel))
    conn.commit()
    conn.close()


if __name__ == '__main__':
    init_db()
    popular_viaturas_padrao_46bpm()
    print("Banco de dados SQLite inicializado com sucesso.")
