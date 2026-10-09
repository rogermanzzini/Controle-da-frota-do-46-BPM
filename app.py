"""
app.py - Servidor Web Flask para Relatório de Abastecimentos do 46º BPM (Bagé/RS)
Backend em Python + Flask com rotas para upload, processamento, relatórios e exportação.
"""

import os
import io
from datetime import datetime
from typing import Optional, Dict, Any, List
from flask import Flask, render_template, request, redirect, url_for, flash, jsonify, send_file, session
from flask_login import LoginManager, UserMixin, login_user, logout_user, login_required, current_user
from werkzeug.utils import secure_filename

import database
from services.excel_processor import (
    processar_planilha_viaturas_alvo,
    processar_cruzamento_abastecimentos,
    gerar_relatorio_excel_openpyxl,
    extrair_placas_de_texto
)
from services.sample_generator import gerar_planilhas_exemplo

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'bpm46-bage-segredo-policial-militar-prod-9821')

# Suporte a Proxy Reverso para Produção (Render, Cloudflare, Nginx, Railway)
from werkzeug.middleware.proxy_fix import ProxyFix
app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1, x_prefix=1)

# ─────────────────────── FLASK-LOGIN ────────────────────────────────────────
login_manager = LoginManager(app)
login_manager.login_view = 'login'
login_manager.login_message = 'Faça login para acessar o sistema.'
login_manager.login_message_category = 'warning'


class Usuario(UserMixin):
    """Classe de usuário para o Flask-Login."""
    def __init__(self, data: dict):
        self.id = str(data['id'])
        self.username = data['username']
        self.nome_completo = data['nome_completo']
        self.papel = data['papel']

    @property
    def is_admin(self):
        return self.papel == 'admin'


@login_manager.user_loader
def load_user(user_id):
    data = database.usuario_get_by_id(int(user_id))
    return Usuario(data) if data else None

# ────────────────────────────────────────────────────────────────────────────

# Configuração de pastas de upload e relatórios
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')
REPORTS_FOLDER = os.path.join(BASE_DIR, 'relatorios_gerados')
EXEMPLOS_FOLDER = os.path.join(BASE_DIR, 'exemplos')
PASTAS_VIATURAS_FOLDER = os.path.join(UPLOAD_FOLDER, 'manutencao', 'viaturas')

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(REPORTS_FOLDER, exist_ok=True)
os.makedirs(EXEMPLOS_FOLDER, exist_ok=True)
os.makedirs(PASTAS_VIATURAS_FOLDER, exist_ok=True)

# Inicializar banco SQLite na subida da aplicação
database.init_db()


@app.context_processor
def inject_global_data():
    """Injeta data e informações globais em todos os templates."""
    from services.report_config import (
        get_assinatura_config,
        get_atestado_config,
        get_atestador_recebimento_config,
        get_empresa_config,
        GRADUACOES_DISPONIVEIS,
        CARGOS_DISPONIVEIS
    )
    from services.atestado_generator import gerar_texto_padrao_atestado

    texto_padrao_atestado = ""
    try:
        texto_padrao_atestado = gerar_texto_padrao_atestado()
    except Exception:
        pass

    # Identifica o módulo ativo para estilização da navegação
    endpoint = request.endpoint or ''
    caminho = request.path or ''
    if 'manutencao' in endpoint or caminho.startswith('/manutencao'):
        modulo_ativo = 'manutencao'
    elif 'pneus' in endpoint or caminho.startswith('/pneus'):
        modulo_ativo = 'pneus'
    else:
        modulo_ativo = 'combustiveis'

    graduacoes_militares = [
        "Sd PM",
        "2º Sgt PM",
        "1º Sgt PM",
        "Subten PM",
        "2º Ten PM",
        "1º Ten PM",
        "Cap PM",
        "Maj PM",
        "Ten Cel PM"
    ]

    cargos_militares = [
        "Aux P4/46º BPM",
        "Auxiliar do P4/46º BPM",
        "Chefe do P4/46º BPM",
        "Gestor de Frota",
        "Encarregado de Garagem"
    ]

    return {
        'agora': datetime.now(),
        'batalhao': 'BRIGADA MILITAR - 46º BPM',
        'app_titulo': 'GESTÃO DE FROTA DE VIATURAS E EQUIPAMENTOS DO 46º BPM',
        'modulo_ativo': modulo_ativo,
        'assinatura_padrao': get_assinatura_config(),
        'atestado_config': get_atestado_config(),
        'atestador_padrao': get_atestador_recebimento_config(),
        'empresa_padrao': get_empresa_config(),
        'texto_padrao_atestado': texto_padrao_atestado,
        'graduacoes_disponiveis': GRADUACOES_DISPONIVEIS,
        'cargos_disponiveis': CARGOS_DISPONIVEIS,
        'graduacoes_militares': graduacoes_militares,
        'cargos_militares': cargos_militares
    }


# ═══════════════════════════ AUTENTICAÇÃO ════════════════════════════════════

@app.route('/login', methods=['GET', 'POST'])
def login():
    """Tela de login do sistema."""
    if current_user.is_authenticated:
        return redirect(url_for('dashboard'))
    if request.method == 'POST':
        username = request.form.get('username', '').strip().lower()
        senha = request.form.get('senha', '')
        usuario_data = database.usuario_verificar_senha(username, senha)
        if usuario_data:
            user_obj = Usuario(usuario_data)
            login_user(user_obj, remember=True)
            database.usuario_atualizar_ultimo_acesso(usuario_data['id'])
            next_page = request.args.get('next')
            flash(f'Bem-vindo, {usuario_data["nome_completo"]}!', 'success')
            return redirect(next_page or url_for('dashboard'))
        else:
            flash('Usuário ou senha incorretos.', 'danger')
    return render_template('login.html')


@app.route('/logout')
@login_required
def logout():
    """Encerra a sessão do usuário."""
    logout_user()
    flash('Você saiu do sistema.', 'info')
    return redirect(url_for('login'))


@app.route('/admin/usuarios')
@login_required
def admin_usuarios():
    """Tela de administração de usuários (apenas admins)."""
    if not current_user.is_admin:
        flash('Acesso restrito a administradores.', 'danger')
        return redirect(url_for('dashboard'))
    usuarios = database.usuarios_listar()
    return render_template('admin_usuarios.html', usuarios=usuarios)


@app.route('/admin/usuarios/criar', methods=['POST'])
@login_required
def admin_criar_usuario():
    """Cria um novo usuário."""
    if not current_user.is_admin:
        return jsonify({'erro': 'Acesso negado'}), 403
    username = request.form.get('username', '').strip().lower()
    nome_completo = request.form.get('nome_completo', '').strip()
    senha = request.form.get('senha', '')
    papel = request.form.get('papel', 'operador')
    if not username or not nome_completo or not senha:
        flash('Preencha todos os campos obrigatórios.', 'danger')
        return redirect(url_for('admin_usuarios'))
    if len(senha) < 6:
        flash('A senha deve ter pelo menos 6 caracteres.', 'danger')
        return redirect(url_for('admin_usuarios'))
    ok = database.usuario_criar(username, nome_completo, senha, papel)
    if ok:
        flash(f'Usuário "{username}" criado com sucesso!', 'success')
    else:
        flash(f'Usuário "{username}" já existe.', 'danger')
    return redirect(url_for('admin_usuarios'))


@app.route('/admin/usuarios/editar/<int:user_id>', methods=['POST'])
@login_required
def admin_editar_usuario(user_id):
    """Edita dados de um usuário."""
    if not current_user.is_admin:
        return jsonify({'erro': 'Acesso negado'}), 403
    nome_completo = request.form.get('nome_completo', '').strip()
    papel = request.form.get('papel', 'operador')
    senha = request.form.get('senha', '').strip() or None
    if senha and len(senha) < 6:
        flash('A nova senha deve ter pelo menos 6 caracteres.', 'danger')
        return redirect(url_for('admin_usuarios'))
    database.usuario_atualizar(user_id, nome_completo, papel, senha)
    flash('Usuário atualizado com sucesso!', 'success')
    return redirect(url_for('admin_usuarios'))


@app.route('/admin/usuarios/excluir/<int:user_id>', methods=['POST'])
@login_required
def admin_excluir_usuario(user_id):
    """Desativa um usuário."""
    if not current_user.is_admin:
        return jsonify({'erro': 'Acesso negado'}), 403
    if int(user_id) == int(current_user.id):
        flash('Você não pode excluir seu próprio usuário.', 'danger')
        return redirect(url_for('admin_usuarios'))
    if database.usuarios_contar_admins() <= 1:
        u = database.usuario_get_by_id(user_id)
        if u and u.get('papel') == 'admin':
            flash('Não é possível remover o único administrador do sistema.', 'danger')
            return redirect(url_for('admin_usuarios'))
    database.usuario_excluir(user_id)
    flash('Usuário removido com sucesso.', 'success')
    return redirect(url_for('admin_usuarios'))


@app.route('/minha-conta', methods=['GET', 'POST'])
@login_required
def minha_conta():
    """Página para o usuário alterar sua própria senha."""
    if request.method == 'POST':
        senha_atual = request.form.get('senha_atual', '')
        nova_senha = request.form.get('nova_senha', '')
        confirmar = request.form.get('confirmar_senha', '')
        ok = database.usuario_verificar_senha(current_user.username, senha_atual)
        if not ok:
            flash('Senha atual incorreta.', 'danger')
        elif nova_senha != confirmar:
            flash('A nova senha e a confirmação não coincidem.', 'danger')
        elif len(nova_senha) < 6:
            flash('A nova senha deve ter pelo menos 6 caracteres.', 'danger')
        else:
            database.usuario_atualizar(int(current_user.id), current_user.nome_completo, current_user.papel, nova_senha)
            flash('Senha alterada com sucesso!', 'success')
    return render_template('minha_conta.html')

# ═════════════════════════════════════════════════════════════════════════════


@app.route('/combustiveis')
@login_required
def modulo_combustiveis():
    """Redireciona para o painel principal do módulo de Combustíveis."""
    return redirect(url_for('dashboard'))


@app.route('/')
@login_required
def dashboard():
    """Painel principal com KPIs, gráficos de consumo e resumo da frota."""
    stats = database.obter_estatisticas_gerais()
    resumo_viaturas = database.obter_resumo_por_viatura()
    importacoes = database.obter_historico_importacoes()
    anomalias = database.obter_anomalias()
    viaturas_cadastradas = database.obter_veiculos_alvo()
    viaturas_com_abastecimento = database.obter_veiculos_com_abastecimento()

    return render_template(
        'dashboard.html',
        stats=stats,
        resumo_viaturas=resumo_viaturas,
        importacoes=importacoes,
        total_anomalias=len(anomalias),
        viaturas_cadastradas=viaturas_cadastradas,
        viaturas_com_abastecimento=viaturas_com_abastecimento
    )


@app.route('/upload')
@login_required
def upload_page():
    """Página de upload das planilhas Excel (Geral do RS e Viaturas do 46º BPM)."""
    viaturas_cadastradas = database.obter_veiculos_alvo()
    return render_template('upload.html', viaturas_cadastradas=viaturas_cadastradas)


@app.route('/processar', methods=['POST'])
@login_required
def processar_arquivos():
    """Processa as planilhas enviadas com Pandas e salva resultados no SQLite."""
    arquivo_rs = request.files.get('arquivo_rs')
    arquivo_alvo = request.files.get('arquivo_alvo')
    texto_placas = request.form.get('texto_placas', '').strip()
    usar_viaturas_cadastradas = request.form.get('usar_viaturas_cadastradas') == '1'

    if not arquivo_rs or arquivo_rs.filename == '':
        flash('Por favor, selecione a planilha geral de abastecimentos do RS (.xlsx ou .xls).', 'danger')
        return redirect(url_for('upload_page'))

    nome_rs = secure_filename(arquivo_rs.filename)
    caminho_rs = os.path.join(UPLOAD_FOLDER, f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{nome_rs}")
    arquivo_rs.save(caminho_rs)

    lista_veiculos_alvo = []
    nome_alvo = "Frota salva no SQLite"

    # Opção 1: Usuário enviou arquivo com placas / viaturas
    if arquivo_alvo and arquivo_alvo.filename != '':
        nome_alvo = secure_filename(arquivo_alvo.filename)
        caminho_alvo = os.path.join(UPLOAD_FOLDER, f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{nome_alvo}")
        arquivo_alvo.save(caminho_alvo)

        try:
            veiculos_planilha = processar_planilha_viaturas_alvo(caminho_alvo)
            if veiculos_planilha:
                lista_veiculos_alvo.extend(veiculos_planilha)
                database.salvar_veiculos_alvo(veiculos_planilha)
        except Exception as e:
            flash(f'Erro ao processar arquivo de viaturas/placas: {str(e)}', 'danger')
            return redirect(url_for('upload_page'))

    # Opção 2: Usuário informou placas diretamente no campo de texto
    if texto_placas:
        veiculos_digitados = extrair_placas_de_texto(texto_placas)
        if veiculos_digitados:
            lista_veiculos_alvo.extend(veiculos_digitados)
            database.salvar_veiculos_alvo(veiculos_digitados)
            nome_alvo = f"Lista de {len(veiculos_digitados)} placas informadas"

    # Opção 3: Usar viaturas já salvas no banco caso nenhuma nova seja informada
    if not lista_veiculos_alvo and usar_viaturas_cadastradas:
        lista_veiculos_alvo = database.obter_veiculos_alvo()

    # Se ainda assim não houver viaturas
    if not lista_veiculos_alvo:
        lista_veiculos_alvo = database.obter_veiculos_alvo()
        if not lista_veiculos_alvo:
            flash('Por favor, informe as placas das viaturas desejadas (por arquivo Excel ou digitando no campo de texto).', 'warning')
            return redirect(url_for('upload_page'))

    # Executar o cruzamento e cálculos via Pandas
    try:
        meta, lista_abastecimentos, df_resumo = processar_cruzamento_abastecimentos(
            caminho_rs,
            lista_veiculos_alvo
        )

        if not lista_abastecimentos:
            flash(f'Nenhum registro correspondente às viaturas do 46º BPM foi encontrado entre as {meta.get("total_linhas_rs", 0)} linhas da planilha do RS.', 'warning')
            return redirect(url_for('upload_page'))

        # Salvar lote no banco de dados SQLite
        meta['arquivo_rs'] = nome_rs
        meta['arquivo_alvo'] = nome_alvo
        importacao_id = database.salvar_lote_importacao(meta, lista_abastecimentos)

        # Gerar cópia do relatório Excel formatado com openpyxl
        nome_excel_saida = f"relatorio_46bpm_lote_{importacao_id}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        caminho_excel_saida = os.path.join(REPORTS_FOLDER, nome_excel_saida)
        gerar_relatorio_excel_openpyxl(
            caminho_excel_saida,
            meta,
            lista_abastecimentos,
            df_resumo.to_dict('records')
        )

        flash(f'Processamento concluído com sucesso! {meta["total_filtrado"]} abastecimentos identificados para {meta["total_veiculos"]} viaturas do 46º BPM.', 'success')
        return redirect(url_for('relatorio'))

    except Exception as e:
        flash(f'Erro durante o processamento das planilhas: {str(e)}', 'danger')
        return redirect(url_for('upload_page'))


@app.route('/relatorio')
@login_required
def relatorio():
    """Tabela detalhada de consumo por viatura com filtros e histórico de abastecimentos."""
    resumo_viaturas = database.obter_resumo_por_viatura()
    abastecimentos = database.obter_todos_abastecimentos(limit=500)
    anomalias = database.obter_anomalias()
    stats = database.obter_estatisticas_gerais()

    return render_template(
        'relatorio.html',
        resumo_viaturas=resumo_viaturas,
        abastecimentos=abastecimentos,
        anomalias=anomalias,
        stats=stats
    )


@app.route('/viatura/<identificador>')
@login_required
def viatura_detalhe(identificador):
    """Ficha analítica de uma viatura específica com histórico cronológico e gráficos."""
    abastecimentos = database.obter_abastecimentos_viatura(identificador)
    if not abastecimentos:
        flash(f'Viatura {identificador} não possui registros de abastecimento.', 'warning')
        return redirect(url_for('relatorio'))

    v_info = abastecimentos[0]
    total_km = sum(a['km_rodado'] for a in abastecimentos if a['km_rodado'] and a['km_rodado'] > 0)
    total_litros = sum(a['litros'] for a in abastecimentos if a['litros'])
    total_valor = sum(a['valor_total'] for a in abastecimentos if a['valor_total'])

    litros_uteis = sum(a['litros'] for a in abastecimentos if a['km_rodado'] and a['km_rodado'] > 0 and a['anomalia'] == 0)
    consumo_medio = round(total_km / litros_uteis, 2) if litros_uteis > 0 else 0.0

    return render_template(
        'veiculo_detalhe.html',
        viatura=v_info,
        abastecimentos=abastecimentos,
        total_km=total_km,
        total_litros=round(total_litros, 2),
        total_valor=round(total_valor, 2),
        consumo_medio=consumo_medio
    )


@app.route('/viaturas')
@login_required
def listar_viaturas():
    """Gestão da frota monitorada do 46º BPM cadastrada no SQLite."""
    viaturas = database.obter_veiculos_alvo()
    return render_template('viaturas.html', viaturas=viaturas)


@app.route('/viatura/adicionar', methods=['POST'])
@login_required
def adicionar_viatura():
    """Cadastra manualmente uma viatura na frota oficial do 46º BPM e cria sua pasta digital."""
    placa = request.form.get('placa', '').strip().upper()
    prefixo = request.form.get('prefixo', '').strip().upper()
    marca = request.form.get('marca', 'Toyota').strip()
    modelo = request.form.get('modelo', '').strip()
    cidade_carga = request.form.get('cidade_carga', 'Bagé').strip()
    fracao = request.form.get('fracao', f"46º BPM - {cidade_carga}").strip()
    tipo_comb = request.form.get('tipo_combustivel', 'Diesel S10').strip()

    if not placa and not prefixo:
        flash('Informe ao menos a placa ou o prefixo da viatura.', 'danger')
        return redirect(request.referrer or url_for('modulo_manutencao'))

    if not placa:
        placa = f"BM-{prefixo}"

    database.salvar_veiculos_alvo([{
        'placa': placa,
        'prefixo': prefixo,
        'marca': marca,
        'modelo': modelo,
        'cidade_carga': cidade_carga,
        'fracao': fracao,
        'tipo_combustivel': tipo_comb
    }])

    # Cria a pasta digital da viatura no disco
    pref_pasta = prefixo or placa
    pasta_viatura = os.path.join(PASTAS_VIATURAS_FOLDER, secure_filename(pref_pasta))
    os.makedirs(pasta_viatura, exist_ok=True)

    flash(f'Viatura BM {prefixo or placa} cadastrada com sucesso e pasta digital criada!', 'success')
    return redirect(request.referrer or url_for('modulo_manutencao'))


@app.route('/upload-placas', methods=['POST'])
@login_required
def upload_placas():
    """Importa viaturas/placas a partir de planilha Excel ou de texto colado."""
    arquivo_placas = request.files.get('arquivo_placas')
    texto_placas = request.form.get('texto_placas', '').strip()
    veiculos = []

    if arquivo_placas and arquivo_placas.filename != '':
        nome = secure_filename(arquivo_placas.filename)
        caminho = os.path.join(UPLOAD_FOLDER, f"placas_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{nome}")
        arquivo_placas.save(caminho)
        try:
            veiculos.extend(processar_planilha_viaturas_alvo(caminho))
        except Exception as e:
            flash(f'Erro ao ler planilha de placas: {str(e)}', 'danger')

    if texto_placas:
        veiculos.extend(extrair_placas_de_texto(texto_placas))

    if veiculos:
        qtd = database.salvar_veiculos_alvo(veiculos)
        flash(f'{qtd} viatura(s) / placa(s) cadastradas com sucesso na frota monitorada do 46º BPM!', 'success')
    else:
        flash('Nenhuma placa válida foi identificada no arquivo ou texto fornecido.', 'warning')

    return redirect(url_for('listar_viaturas'))


@app.route('/exportar-excel')
@login_required
def exportar_excel():
    """Gera e faz o download do relatório Excel formatado (.xlsx) em tempo real via openpyxl."""
    stats = database.obter_estatisticas_gerais()
    resumo_viaturas = database.obter_resumo_por_viatura()
    abastecimentos = database.obter_todos_abastecimentos(limit=10000)

    if not abastecimentos:
        flash('Não há dados de abastecimentos salvos para exportar.', 'warning')
        return redirect(url_for('relatorio'))

    caminho_temporario = os.path.join(REPORTS_FOLDER, f"relatorio_consumo_46bpm_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
    gerar_relatorio_excel_openpyxl(
        caminho_temporario,
        stats,
        abastecimentos,
        resumo_viaturas
    )

    return send_file(
        caminho_temporario,
        as_attachment=True,
        download_name=f"Relatorio_Consumo_46BPM_Bage_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@app.route('/carregar-dados-demonstracao', methods=['POST'])
@login_required
def carregar_dados_demonstracao():
    """Gera planilhas de exemplo realistas e carrega no banco em 1 clique."""
    try:
        caminho_rs, caminho_alvo = gerar_planilhas_exemplo(EXEMPLOS_FOLDER)
        viaturas = processar_planilha_viaturas_alvo(caminho_alvo)
        database.salvar_veiculos_alvo(viaturas)

        meta, abastecimentos, df_resumo = processar_cruzamento_abastecimentos(caminho_rs, viaturas)
        meta['arquivo_rs'] = "abastecimentos_rs_geral_exemplo.xlsx"
        meta['arquivo_alvo'] = "viaturas_46bpm_alvo_exemplo.xlsx"

        importacao_id = database.salvar_lote_importacao(meta, abastecimentos)

        # Gerar o excel correspondente
        caminho_excel = os.path.join(REPORTS_FOLDER, f"relatorio_46bpm_demonstracao.xlsx")
        gerar_relatorio_excel_openpyxl(caminho_excel, meta, abastecimentos, df_resumo.to_dict('records'))

        flash(f'Dados de demonstração carregados com sucesso! {meta["total_filtrado"]} abastecimentos de {meta["total_veiculos"]} viaturas do 46º BPM.', 'success')
        return redirect(url_for('dashboard'))
    except Exception as e:
        flash(f'Erro ao carregar dados de demonstração: {str(e)}', 'danger')
        return redirect(url_for('dashboard'))


@app.route('/baixar-exemplo/<tipo>')
@login_required
def baixar_exemplo(tipo):
    """Permite ao usuário baixar os modelos de planilha Excel para preenchimento."""
    caminho_rs = os.path.join(EXEMPLOS_FOLDER, 'abastecimentos_rs_geral_exemplo.xlsx')
    caminho_alvo = os.path.join(EXEMPLOS_FOLDER, 'viaturas_46bpm_alvo_exemplo.xlsx')

    if not os.path.exists(caminho_rs) or not os.path.exists(caminho_alvo):
        gerar_planilhas_exemplo(EXEMPLOS_FOLDER)

    if tipo == 'rs':
        return send_file(
            caminho_rs,
            as_attachment=True,
            download_name="modelo_abastecimentos_rs.xlsx"
        )
    else:
        return send_file(
            caminho_alvo,
            as_attachment=True,
            download_name="modelo_viaturas_46bpm_alvo.xlsx"
        )


@app.route('/limpar-dados', methods=['POST'])
@login_required
def limpar_dados():
    """Limpa todas as tabelas de abastecimentos do banco de dados SQLite."""
    database.limpar_todos_os_dados()
    flash('Banco de dados de abastecimentos limpo com sucesso!', 'success')
    return redirect(url_for('dashboard'))


# Rota para gerar PDF filtrado por viaturas selecionadas
@app.route('/relatorio/pdf', methods=['POST'])
@login_required
def gerar_relatorio_pdf():
    """Gera um PDF contendo apenas os abastecimentos das viaturas selecionadas (por placa ou prefixo)."""
    raw_veiculos = request.form.getlist('veiculos') or request.form.getlist('prefixos') or request.form.getlist('placas')
    veiculos = [v.strip() for v in raw_veiculos if v and v.strip()]
    if not veiculos:
        flash('Selecione ao menos uma viatura para gerar o relatório.', 'warning')
        return redirect(url_for('dashboard'))
    # Busca os abastecimentos referentes às viaturas selecionadas
    abastecimentos = database.obter_abastecimentos_por_veiculos(veiculos)
    if not abastecimentos:
        flash('Nenhum abastecimento encontrado para as viaturas selecionadas.', 'info')
        return redirect(url_for('dashboard'))

    # Dados customizados de assinatura do formulário
    nome_assinante = request.form.get('nome_assinante', '').strip()
    graduacao_assinante = request.form.get('graduacao_assinante', '').strip()
    cargo_assinante = request.form.get('cargo_assinante', '').strip()
    salvar_como_padrao = request.form.get('salvar_como_padrao') == '1'

    dados_assinatura = None
    if nome_assinante:
        dados_assinatura = {
            'nome': nome_assinante,
            'graduacao': graduacao_assinante or 'SD PM',
            'cargo': cargo_assinante or 'Auxiliar do P4/46º BPM'
        }
        if salvar_como_padrao:
            from services.report_config import set_assinatura_config
            set_assinatura_config(nome_assinante, graduacao_assinante or 'SD PM', cargo_assinante or 'Auxiliar do P4/46º BPM')
            flash('Assinatura salva como novo padrão fixo do sistema!', 'info')

    # Cria o PDF
    from services.pdf_report_generator import gerar_pdf, obter_nome_arquivo_relatorio
    nome_arquivo_download = obter_nome_arquivo_relatorio(abastecimentos)
    pdf_path = os.path.join(REPORTS_FOLDER, f"relatorio_consumo_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf")
    gerar_pdf(pdf_path, abastecimentos, veiculos, dados_assinatura=dados_assinatura)
    return send_file(
        pdf_path,
        as_attachment=True,
        download_name=nome_arquivo_download,
        mimetype='application/pdf'
    )


# Rota para gerar o ATESTADO DE COMBUSTÍVEIS oficial em PDF
@app.route('/atestado/pdf', methods=['GET', 'POST'])
@app.route('/relatorio/atestado-pdf', methods=['GET', 'POST'])
@login_required
def gerar_atestado_pdf_route():
    """Gera o documento oficial ATESTADO DE COMBUSTÍVEIS em PDF com valor total por extenso e quinzena."""
    abastecimentos = database.obter_todos_abastecimentos(limit=50000)
    if not abastecimentos:
        flash('Nenhum abastecimento encontrado para gerar o Atestado de Combustíveis.', 'warning')
        return redirect(url_for('dashboard'))

    stats = database.obter_estatisticas_gerais()
    total_valor = stats.get('total_valor', 0.0)

    from services.report_config import get_atestado_config, set_atestado_config
    cfg = get_atestado_config()

    # Modo de assinatura: titular, respondendo, impedimento
    modo = request.form.get('modo') or request.args.get('modo') or cfg.get('modo', 'impedimento')
    salvar_como_padrao = request.form.get('salvar_como_padrao') == '1'

    dados_assinatura = {
        'modo': modo,
        'titular_nome': request.form.get('titular_nome', '').strip() or cfg.get('titular_nome'),
        'titular_cargo': request.form.get('titular_cargo', '').strip() or cfg.get('titular_cargo'),
        'respondendo_nome': request.form.get('respondendo_nome', '').strip() or cfg.get('respondendo_nome'),
        'respondendo_cargo': request.form.get('respondendo_cargo', '').strip() or cfg.get('respondendo_cargo'),
        'subcomandante_nome': request.form.get('subcomandante_nome', '').strip() or cfg.get('subcomandante_nome'),
        'subcomandante_cargo': request.form.get('subcomandante_cargo', '').strip() or cfg.get('subcomandante_cargo'),
        'gestor_nome': request.form.get('gestor_nome', '').strip() or cfg.get('gestor_nome'),
        'gestor_cargo': request.form.get('gestor_cargo', '').strip() or cfg.get('gestor_cargo'),
    }

    if salvar_como_padrao:
        set_atestado_config(dados_assinatura)
        flash('Configurações de assinatura do Atestado salvas como padrão!', 'info')

    from services.atestado_generator import gerar_atestado_pdf as criar_atestado, obter_dados_periodo_atestado
    info_periodo = obter_dados_periodo_atestado(abastecimentos)
    nome_arquivo = info_periodo['nome_arquivo']

    # Captura o texto personalizado caso o usuário tenha editado no formulário
    texto_atestado = request.form.get('texto_atestado', '').strip() or request.args.get('texto_atestado', '').strip() or None

    pdf_path = os.path.join(REPORTS_FOLDER, f"atestado_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf")
    criar_atestado(
        pdf_path,
        abastecimentos,
        valor_total=total_valor,
        dados_assinatura=dados_assinatura,
        texto_personalizado=texto_atestado
    )

    return send_file(
        pdf_path,
        as_attachment=True,
        download_name=nome_arquivo,
        mimetype='application/pdf'
    )


@app.route('/api/atestado/texto-padrao')
@login_required
def api_texto_padrao_atestado():
    """Retorna o texto oficial padrão do Atestado de Combustíveis em formato JSON."""
    from services.atestado_generator import gerar_texto_padrao_atestado
    texto = gerar_texto_padrao_atestado()
    return jsonify({'texto': texto})


# Rota para configurar as assinaturas padrão do Atestado
@app.route('/configuracoes/atestado', methods=['POST'])
@login_required
def salvar_config_atestado():
    """Salva os dados de assinatura padrão do Atestado de Combustíveis."""
    from services.report_config import set_atestado_config
    dados = {
        'modo': request.form.get('modo', 'impedimento').strip(),
        'titular_nome': request.form.get('titular_nome', '').strip(),
        'titular_cargo': request.form.get('titular_cargo', '').strip(),
        'respondendo_nome': request.form.get('respondendo_nome', '').strip(),
        'respondendo_cargo': request.form.get('respondendo_cargo', '').strip(),
        'subcomandante_nome': request.form.get('subcomandante_nome', '').strip(),
        'subcomandante_cargo': request.form.get('subcomandante_cargo', '').strip(),
        'gestor_nome': request.form.get('gestor_nome', '').strip(),
        'gestor_cargo': request.form.get('gestor_cargo', '').strip(),
    }
    set_atestado_config(dados)
    flash('Configuração de assinaturas do Atestado salva com sucesso!', 'success')
    return redirect(request.referrer or url_for('dashboard'))


# Rota para configurar e salvar a assinatura padrão do sistema
@app.route('/configuracoes/assinatura', methods=['POST'])
@login_required
def salvar_config_assinatura():
    """Salva os dados de assinatura padrão no report_config.json."""
    nome = request.form.get('nome', '').strip()
    graduacao = request.form.get('graduacao', '').strip()
    cargo = request.form.get('cargo', '').strip()
    if not nome:
        flash('Por favor, informe o nome do militar responsável.', 'warning')
        return redirect(request.referrer or url_for('dashboard'))

    from services.report_config import set_assinatura_config
    set_assinatura_config(nome, graduacao or 'SD PM', cargo or 'Auxiliar do P4/46º BPM')
    flash('Configuração da assinatura padrão atualizada com sucesso!', 'success')
    return redirect(request.referrer or url_for('dashboard'))


# Nova rota para remover todas as viaturas cadastradas
@app.route('/viaturas/remover', methods=['POST'])
@login_required
def remover_veiculos():
    """Remove todas as viaturas da tabela veiculos_alvo."""
    database.remover_todas_veiculos_alvo()
    flash('Todas as viaturas foram removidas com sucesso.', 'success')
    return redirect(url_for('listar_viaturas'))


# ==============================================================================
# MÓDULO 2: GESTÃO DE MANUTENÇÃO, PASTAS POR VIATURA E ATESTADO DE RECEBIMENTO
# ==============================================================================

def salvar_arquivo_na_pasta_viatura(
    arquivo_file,
    viatura_prefixo: str,
    tipo_doc: str,
    descricao: str = "",
    manutencao_id: Optional[int] = None,
    pasta_os: Optional[str] = None
) -> Optional[Dict[str, Any]]:
    """Salva um arquivo anexado fisicamente dentro da pasta da viatura e subpasta de OS, registrando no SQLite."""
    if not arquivo_file or not hasattr(arquivo_file, 'filename') or arquivo_file.filename == '':
        return None
    
    prefixo_limpo = secure_filename(str(viatura_prefixo).strip().upper())
    if not prefixo_limpo:
        prefixo_limpo = "GERAL"

    nome_orig = secure_filename(arquivo_file.filename)
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    nome_salvo = f"{tipo_doc.upper()}_{timestamp}_{nome_orig}"

    if pasta_os and str(pasta_os).strip():
        pasta_os_limpa = secure_filename(str(pasta_os).strip())
        pasta_dest = os.path.join(PASTAS_VIATURAS_FOLDER, prefixo_limpo, pasta_os_limpa)
        caminho_rel = os.path.join('uploads', 'manutencao', 'viaturas', prefixo_limpo, pasta_os_limpa, nome_salvo).replace('\\', '/')
    else:
        pasta_os_limpa = None
        pasta_dest = os.path.join(PASTAS_VIATURAS_FOLDER, prefixo_limpo)
        caminho_rel = os.path.join('uploads', 'manutencao', 'viaturas', prefixo_limpo, nome_salvo).replace('\\', '/')

    os.makedirs(pasta_dest, exist_ok=True)
    caminho_abs = os.path.join(pasta_dest, nome_salvo)
    arquivo_file.save(caminho_abs)

    tamanho = os.path.getsize(caminho_abs)

    doc_id = database.salvar_documento_viatura(
        viatura_prefixo=prefixo_limpo,
        tipo_documento=tipo_doc,
        nome_original=nome_orig,
        nome_arquivo=nome_salvo,
        caminho_relativo=caminho_rel,
        tamanho_bytes=tamanho,
        descricao=descricao,
        manutencao_id=manutencao_id,
        pasta_os=pasta_os_limpa
    )
    return {
        'id': doc_id,
        'nome_salvo': nome_salvo,
        'caminho_relativo': caminho_rel,
        'pasta_os': pasta_os_limpa
    }


@app.route('/manutencao/config/assinatura', methods=['POST'])
@login_required
def salvar_config_atestador_route():
    """Salva a configuração do militar atestador padrão do módulo de Manutenção."""
    from services.report_config import set_atestador_recebimento_config
    nome = request.form.get('nome', '').strip()
    graduacao = request.form.get('graduacao', '').strip()
    id_func = request.form.get('id_func', '').strip()
    cargo = request.form.get('cargo', '').strip()

    if not nome or not graduacao or not id_func or not cargo:
        flash('Por favor, preencha todos os campos da assinatura do atestador.', 'warning')
    else:
        set_atestador_recebimento_config(nome, graduacao, id_func, cargo)
        flash('Assinatura do militar atestador salva como padrão com sucesso!', 'success')

    return redirect(request.referrer or url_for('modulo_manutencao'))


@app.route('/manutencao')
@login_required
def modulo_manutencao():
    """Painel do Módulo de Gestão de Manutenção de Viaturas, Pastas e Equipamentos."""
    filtro_status = request.args.get('status', 'Todos')
    filtro_viatura = request.args.get('viatura', 'Todos')
    
    manutencoes = database.obter_manutencoes(filtro_status=filtro_status, filtro_viatura=filtro_viatura)
    equipamentos = database.obter_equipamentos()
    stats = database.obter_estatisticas_manutencao()
    viaturas_cadastradas = database.obter_veiculos_alvo()
    pastas_viaturas = database.obter_todas_pastas_viaturas()
    
    return render_template(
        'manutencao.html',
        manutencoes=manutencoes,
        equipamentos=equipamentos,
        stats=stats,
        viaturas_cadastradas=viaturas_cadastradas,
        pastas_viaturas=pastas_viaturas,
        filtro_status=filtro_status,
        filtro_viatura=filtro_viatura
    )


@app.route('/api/manutencao/extrair-dados-nota', methods=['POST'])
@login_required
def extrair_dados_nota_fiscal():
    """
    Extrai razão social, CNPJ, número da nota e valor total do emitente
    de arquivos XML (NFe/NFS-e) ou PDF de notas fiscais.
    """
    arquivo = request.files.get('arquivo')
    if not arquivo or arquivo.filename == '':
        return jsonify({'erro': 'Nenhum arquivo enviado'}), 400

    nome = arquivo.filename.lower()
    conteudo = arquivo.read()

    empresa_nome = ''
    empresa_cnpj = ''
    numero_nota = ''
    valor_total = ''
    numero_os = ''

    # ---- Tentativa 1: XML (NFe padrão ou NFS-e) ----
    if nome.endswith('.xml'):
        try:
            import xml.etree.ElementTree as ET
            texto_xml = conteudo.decode('utf-8', errors='ignore')
            root = ET.fromstring(texto_xml)

            def _strip_ns(tag):
                return tag.split('}')[-1] if '}' in tag else tag

            def _find_text(node, *tags):
                """Busca aninhada ignorando namespaces."""
                nodes = [node]
                for tag in tags:
                    next_nodes = []
                    for n in nodes:
                        for child in n.iter():
                            if _strip_ns(child.tag) == tag:
                                next_nodes.append(child)
                    nodes = next_nodes
                return nodes[0].text.strip() if nodes and nodes[0].text else ''

            # --- Emitente (NFe) ---
            empresa_nome = _find_text(root, 'emit', 'xNome') or _find_text(root, 'xNome')
            empresa_cnpj = _find_text(root, 'emit', 'CNPJ') or _find_text(root, 'CNPJ')

            # Formatar CNPJ se vier sem pontuação (14 dígitos)
            if empresa_cnpj and len(empresa_cnpj) == 14 and empresa_cnpj.isdigit():
                c = empresa_cnpj
                empresa_cnpj = f"{c[:2]}.{c[2:5]}.{c[5:8]}/{c[8:12]}-{c[12:14]}"

            # --- Número da nota ---
            # NFe: ide/nNF  |  NFS-e: Numero, NumeroNfse, nNFS
            numero_nota = (
                _find_text(root, 'ide', 'nNF') or
                _find_text(root, 'nNF') or
                _find_text(root, 'Numero') or
                _find_text(root, 'NumeroNfse') or
                _find_text(root, 'nNFS') or
                _find_text(root, 'NumeroDaNota')
            )

            # --- Valor total ---
            # NFe: total/ICMSTot/vNF  |  NFS-e: ValorServicos, ValorLiquidoNfse, vNFS
            valor_raw = (
                _find_text(root, 'ICMSTot', 'vNF') or
                _find_text(root, 'vNF') or
                _find_text(root, 'ValorLiquidoNfse') or
                _find_text(root, 'ValorServicos') or
                _find_text(root, 'vNFS') or
                _find_text(root, 'ValorTotal')
            )
            if valor_raw:
                # Converte ponto decimal para vírgula (padrão BR)
                try:
                    v = float(valor_raw.replace(',', '.'))
                    valor_total = f"{v:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
                except ValueError:
                    valor_total = valor_raw

        except Exception as e:
            return jsonify({'erro': f'Falha ao parsear XML: {str(e)}'}), 422

    # ---- Tentativa 2: PDF ----
    elif nome.endswith('.pdf'):
        try:
            import pdfplumber
            import re
            import tempfile

            with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp:
                tmp.write(conteudo)
                tmp_path = tmp.name

            texto_pdf = ''
            linhas_pdf = []
            with pdfplumber.open(tmp_path) as pdf:
                for page in pdf.pages:
                    texto_pagina = page.extract_text() or ''
                    texto_pdf += texto_pagina + '\n'
                    linhas_pdf.extend(texto_pagina.split('\n'))
            try:
                os.unlink(tmp_path)
            except Exception:
                pass

            # ── 1. CNPJ DO FORNECEDOR / EMITENTE ─────────────────────────────────
            # Procura CNPJs no formato com máscara XX.XXX.XXX/XXXX-XX
            # Trata caso onde hífen quebrou de linha
            texto_cnpj_normalizado = re.sub(r'(\d{2}\.\d{3}\.\d{3}/\d{4}-)\s*\n?\s*(\d{2})', r'\1\2', texto_pdf)
            cnpjs = re.findall(r'\b\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}\b', texto_cnpj_normalizado)
            if cnpjs:
                # Descarta o CNPJ do tomador/cliente (Brigada Militar: 89.175.541/0001-64)
                cnpjs_emitentes = [c for c in cnpjs if not c.startswith('89.175.541')]
                empresa_cnpj = cnpjs_emitentes[0] if cnpjs_emitentes else cnpjs[0]

            # ── DETECÇÃO DO TIPO DE DOCUMENTO ────────────────────────────────────
            is_danfe = any('DANFE' in l.upper() for l in linhas_pdf[:25]) or 'RECEBEMOS DE' in texto_pdf.upper()
            is_nfse = any(
                'PRESTADOR DE SERVIÇ' in l.upper() or 'PRESTADOR DE SERVIC' in l.upper() or
                'NFS-E' in l.upper() or 'SECRETARIA MUNICIPAL DA FAZENDA' in l.upper() or
                'NOTA FISCAL DE SERVIÇ' in l.upper()
                for l in linhas_pdf[:20]
            )

            # ── 2. EXTRAÇÃO ESPECÍFICA POR TIPO ──────────────────────────────────
            if is_danfe:
                # A) Nome da empresa no DANFE
                # Canhoto padrão nacional: 'RECEBEMOS DE [EMPRESA] OS PRODUTOS...'
                m_canhoto = re.search(r'RECEBEMOS\s+DE\s+(.+?)\s+OS\s+PRODUTOS', texto_pdf, re.IGNORECASE)
                if m_canhoto and len(m_canhoto.group(1).strip()) >= 3:
                    empresa_nome = m_canhoto.group(1).strip()
                else:
                    # Bloco IDENTIFICACAO DO EMITENTE
                    for i, l in enumerate(linhas_pdf):
                        if 'IDENTIFICAÇ' in l.upper() and 'EMITENTE' in l.upper():
                            termos_descarte = [
                                'DANFE', 'DOCUMENTO AUXILIAR', 'NOTA FISCAL', 'FISCAL ELETRÔNICA', 'FISCAL ELETRONICA',
                                'CHAVE DE ACESSO', '0 - ENTRADA', '1 - SAÍDA', '1 - SAIDA', 'SÉRIE', 'SERIE', 'FOLHA',
                                'Nº', 'N.', 'PROTOCOLO', 'CONSULTA', 'WWW.NFE'
                            ]
                            for j in range(i + 1, min(i + 12, len(linhas_pdf))):
                                cand = linhas_pdf[j].strip()
                                if len(cand) >= 3 and not any(t in cand.upper() for t in termos_descarte) and not re.search(r'^\d[\d\.\-\/]+$', cand):
                                    empresa_nome = cand
                                    break
                            break

                # B) Número da DANFE (formato com 9 algarismos: 000.205.100 ou 9 dígitos seguidos)
                m_num = re.search(r'N[ºo°\.]+\s*(\d{3}\.\d{3}\.\d{3})\b', texto_pdf)
                if not m_num:
                    m_num = re.search(r'\b(\d{3}\.\d{3}\.\d{3})\b', texto_pdf)
                if not m_num:
                    m_num = re.search(r'N[ºo°\.]+\s*(\d{9})\b', texto_pdf)
                if m_num:
                    numero_nota = m_num.group(1)

                # C) Valor da DANFE
                m_val = re.search(r'VALOR TOTAL:\s*R\$\s*([\d\.]+,\d{2})', texto_pdf, re.IGNORECASE)
                if not m_val:
                    m_val = re.search(r'V\.\s*TOTAL\s+DA\s+NOTA\s*[:\s]*([\d\.]+,\d{2})', texto_pdf, re.IGNORECASE)
                if not m_val:
                    m_val = re.search(r'V\.\s*TOTAL\s+PRODUTOS\s*[:\s]*([\d\.]+,\d{2})', texto_pdf, re.IGNORECASE)
                if not m_val:
                    m_val = re.search(r'(?:VALOR\s+DA\s+NOTA|TOTAL\s+DA\s+NOTA)\s*[:\s]*R?\$?\s*([\d\.]+,\d{2})', texto_pdf, re.IGNORECASE)
                if m_val:
                    valor_total = m_val.group(1)

            elif is_nfse:
                # A) Nome da empresa na NFS-e
                # Varre linhas entre 'PRESTADOR DE SERVIÇOS' e 'TOMADOR DE SERVIÇOS'
                idx_p = -1
                idx_t = len(linhas_pdf)
                for i, l in enumerate(linhas_pdf):
                    if 'PRESTADOR DE SERVIÇ' in l.upper() or 'PRESTADOR DE SERVIC' in l.upper():
                        idx_p = i
                    elif idx_p != -1 and ('TOMADOR DE SERVIÇ' in l.upper() or 'TOMADOR DE SERVIC' in l.upper()):
                        idx_t = i
                        break

                if idx_p != -1:
                    termos_descarte = [
                        'PRESTADOR DE SERVIÇOS', 'PRESTADOR DE SERVICOS', 'CNPJ', 'CPF', 'INSCRIÇÃO', 'INSCRICAO',
                        'ENDEREÇO', 'ENDERECO', 'AVENIDA', 'AVE ', 'RUA ', 'CEP', 'TELEFONE', 'FONE', 'MUNICÍPIO',
                        'MUNICIPIO', 'ESTADO', 'NOME/RAZÃO SOCIAL:', 'NOME/RAZAO SOCIAL:', 'RAZÃO SOCIAL:', 'RAZAO SOCIAL:'
                    ]
                    for j in range(idx_p + 1, idx_t):
                        cand = linhas_pdf[j].strip()
                        if len(cand) >= 3 and not any(t in cand.upper() for t in termos_descarte) and not re.search(r'^\d[\d\.\-\/]+$', cand):
                            empresa_nome = cand
                            break

                if not empresa_nome:
                    # Fallback linha com Nome/Razão Social
                    for i, l in enumerate(linhas_pdf):
                        if 'NOME/RAZÃO SOCIAL' in l.upper() or 'NOME/RAZAO SOCIAL' in l.upper():
                            if ':' in l and len(l.split(':', 1)[1].strip()) >= 3:
                                empresa_nome = l.split(':', 1)[1].strip()
                            elif i > 0 and len(linhas_pdf[i-1].strip()) >= 3 and not any(t in linhas_pdf[i-1].upper() for t in ['CNPJ', 'PRESTADOR', 'TOMADOR']):
                                empresa_nome = linhas_pdf[i-1].strip()
                            break

                # B) Número da NFS-e (evita RPS)
                m_num = re.search(r'N[ºo°\.]*\s+da\s+Nota\s*[:\s]+(\d+)', texto_pdf, re.IGNORECASE)
                if not m_num:
                    m_num = re.search(r'N[úu]mero\s+da\s+NFS-?e\s*[:\s]+(\d+)', texto_pdf, re.IGNORECASE)
                if not m_num:
                    m_num = re.search(r'NFS-?e\s*N[ºo°\.]*\s*(\d+)', texto_pdf, re.IGNORECASE)
                if m_num:
                    numero_nota = m_num.group(1)

                # C) Valor da NFS-e
                m_val = re.search(r'Valor\s+L[ií]quido\s+da\s+Nota\s+Fiscal\s*[-:\s]*R\$\s*([\d\.]+,\d{2})', texto_pdf, re.IGNORECASE)
                if not m_val:
                    m_val = re.search(r'VALOR\s+TOTAL\s+DA\s+NOTA\s*[:\s]*([\d\.]+,\d{2})', texto_pdf, re.IGNORECASE)
                if not m_val:
                    m_val = re.search(r'Valor\s+Total\s+dos\s+Servi[çc]os\s*[-:\s]*R\$\s*([\d\.]+,\d{2})', texto_pdf, re.IGNORECASE)
                if m_val:
                    valor_total = m_val.group(1)

            else:
                # Caso genérico / Ordem de Serviço Ticket Log
                m_os_nome = re.search(r'Nome\s+([A-Z0-9\s]+?)\s+CNPJ', texto_pdf)
                if m_os_nome:
                    empresa_nome = m_os_nome.group(1).strip()
                elif linhas_pdf and len(linhas_pdf[0].strip()) >= 5:
                    empresa_nome = linhas_pdf[0].strip()

            # ── 3. NÚMERO DA ORDEM DE SERVIÇO TICKET LOG (8 DÍGITOS) ─────────────
            m_os_num = re.search(r'(?:Ordem\s+de\s+Servi[çc]o|N[úu]mero\s+da\s+Ordem\s+de\s+Servi[çc]o|O\.?S\.?)[:\s]+(\d{8})\b', texto_pdf, re.IGNORECASE)
            if not m_os_num:
                if any(t in texto_pdf.upper() for t in ['TICKET LOG', 'SOU LOG', 'TICKETLOG']):
                    m_os_num = re.search(r'\b(\d{8})\b', texto_pdf)
            if m_os_num:
                numero_os = m_os_num.group(1)

            # Limpeza final do nome da empresa
            if empresa_nome:
                empresa_nome = re.split(r'\s{3,}|\t', empresa_nome)[0].strip()

        except ImportError:
            return jsonify({
                'aviso': 'Biblioteca pdfplumber não disponível. Preencha manualmente.',
                'empresa_nome': '', 'empresa_cnpj': '',
                'numero_nota': '', 'valor_total': '', 'numero_os': ''
            }), 200
        except Exception as e:
            return jsonify({'erro': f'Falha ao ler PDF: {str(e)}'}), 422

    else:
        return jsonify({
            'aviso': 'Formato não suportado para extração. Use XML (recomendado) ou PDF.',
            'empresa_nome': '', 'empresa_cnpj': '',
            'numero_nota': '', 'valor_total': '', 'numero_os': ''
        }), 200

    return jsonify({
        'empresa_nome': empresa_nome,
        'empresa_cnpj': empresa_cnpj,
        'numero_nota': numero_nota,
        'valor_total': valor_total,
        'numero_os': numero_os
    })


@app.route('/manutencao/arquivo/upload', methods=['POST'])
@login_required
def upload_arquivo_viatura_route():
    """Faz upload de arquivo avulso (OS, DANFE, NFS-e ou Outro) na pasta da viatura."""
    viatura_prefixo = request.form.get('viatura_prefixo', '').strip().upper()
    tipo_doc = request.form.get('tipo_documento', 'OUTRO').strip().upper()
    descricao = request.form.get('descricao', '').strip()
    pasta_os_raw = request.form.get('pasta_os', '').strip()
    if pasta_os_raw:
        digitos = re.sub(r'\D', '', pasta_os_raw)
        pasta_os = digitos if len(digitos) == 8 else pasta_os_raw
    else:
        pasta_os = None
    arquivo = request.files.get('arquivo')

    if not viatura_prefixo or not arquivo or arquivo.filename == '':
        flash('Selecione a viatura e o arquivo para upload.', 'warning')
        return redirect(url_for('modulo_manutencao'))

    res = salvar_arquivo_na_pasta_viatura(arquivo, viatura_prefixo, tipo_doc, descricao=descricao, pasta_os=pasta_os)
    if res:
        flash(f'Arquivo anexado com sucesso na pasta da viatura {viatura_prefixo}!', 'success')
    else:
        flash('Falha ao processar o upload do arquivo.', 'danger')

    return redirect(url_for('modulo_manutencao'))


@app.route('/manutencao/arquivo/download/<int:doc_id>')
@login_required
def download_arquivo_viatura_route(doc_id):
    """Envia o arquivo da pasta da viatura para visualização ou download."""
    doc = database.obter_documento_por_id(doc_id)
    if not doc:
        flash('Documento não encontrado.', 'warning')
        return redirect(url_for('modulo_manutencao'))

    caminho_abs = os.path.join(BASE_DIR, doc['caminho_relativo'].replace('/', os.sep))
    if not os.path.exists(caminho_abs):
        flash('Arquivo físico não localizado no disco.', 'danger')
        return redirect(url_for('modulo_manutencao'))

    return send_file(
        caminho_abs,
        as_attachment=False,
        download_name=doc['nome_original']
    )


@app.route('/manutencao/arquivo/excluir/<int:doc_id>', methods=['POST'])
@login_required
def excluir_arquivo_viatura_route(doc_id):
    """Exclui um documento anexado da viatura do disco e do banco."""
    caminho_rel = database.excluir_documento_viatura(doc_id)
    if caminho_rel:
        caminho_abs = os.path.join(BASE_DIR, caminho_rel.replace('/', os.sep))
        try:
            if os.path.exists(caminho_abs):
                os.remove(caminho_abs)
        except Exception:
            pass
        flash('Arquivo removido da pasta da viatura com sucesso.', 'info')
    return redirect(url_for('modulo_manutencao'))


@app.route('/manutencao/atestado/gerar', methods=['POST'])
@login_required
def gerar_atestado_recebimento_route():
    """
    Gera o ATESTADO DE RECEBIMENTO DE MATERIAL/SERVIÇO oficial em PDF,
    salva os arquivos fiscais/OS na pasta da viatura e salva uma cópia do atestado.
    """
    viatura_prefixo = request.form.get('viatura_prefixo', '').strip().upper()
    if not viatura_prefixo:
        flash('Por favor, informe a viatura.', 'warning')
        return redirect(url_for('modulo_manutencao'))

    # Coleta os dados do formulário
    tipo_doc_fiscal = request.form.get('tipo_documento_fiscal', 'nfse').strip().lower()
    
    val_danfe_str = request.form.get('valor_danfe', '0').strip().replace('.', '').replace(',', '.')
    val_nfse_str = request.form.get('valor_nfse', '0').strip().replace('.', '').replace(',', '.')
    val_danfe = float(val_danfe_str) if val_danfe_str else 0.0
    val_nfse = float(val_nfse_str) if val_nfse_str else 0.0
    val_total = val_danfe + val_nfse

    dados = {
        'tipo_documento_fiscal': tipo_doc_fiscal,
        'viatura_prefixo': viatura_prefixo,
        'viatura_marca': request.form.get('viatura_marca', '').strip(),
        'viatura_modelo': request.form.get('viatura_modelo', '').strip(),
        'viatura_placa': request.form.get('viatura_placa', '').strip().upper(),
        'cidade_carga': request.form.get('cidade_carga', 'Bagé').strip(),
        'numero_danfe': request.form.get('numero_danfe', '').strip(),
        'valor_danfe': val_danfe,
        'numero_nfse': request.form.get('numero_nfse', '').strip(),
        'valor_nfse': val_nfse,
        'empresa_nome': request.form.get('empresa_nome', '').strip(),
        'empresa_cnpj': request.form.get('empresa_cnpj', '').strip(),
        'data_servico': request.form.get('data_servico', datetime.now().strftime('%Y-%m-%d')).strip(),
        'descricao_servico': request.form.get('descricao_servico', '').strip() or f"Manutenção e conservação da viatura BM {viatura_prefixo}",
        'pecas': request.form.get('pecas', '').strip(),
        'valor_total': val_total,
        'status': 'Concluída',
        'data_conclusao': request.form.get('data_servico', datetime.now().strftime('%Y-%m-%d')).strip(),
        'responsavel_nome': request.form.get('responsavel_nome', 'JULIANO INOCENTE VELEDA').strip().upper(),
        'responsavel_grad': request.form.get('responsavel_grad', 'Sd PM').strip(),
        'responsavel_id_func': request.form.get('responsavel_id_func', '3155358').strip(),
        'responsavel_cargo': request.form.get('responsavel_cargo', 'Aux P4/46º BPM').strip(),
        'observacoes': request.form.get('observacoes', '').strip()
    }

    # 1. Se solicitado pelo usuário, atualiza a assinatura padrão do atestador
    if request.form.get('salvar_como_padrao_atestador') == '1':
        from services.report_config import set_atestador_recebimento_config
        set_atestador_recebimento_config(
            nome=dados['responsavel_nome'],
            graduacao=dados['responsavel_grad'],
            id_func=dados['responsavel_id_func'],
            cargo=dados['responsavel_cargo']
        )

    # 1b. Se solicitado, salva empresa e CNPJ como padrão para futuros atestados
    if request.form.get('salvar_empresa_padrao') == '1':
        from services.report_config import set_empresa_config
        set_empresa_config(
            nome=dados['empresa_nome'],
            cnpj=dados['empresa_cnpj']
        )

    # 2. Subpasta física da Ordem de Serviço Ticket Log (8 algarismos)
    num_os_raw = request.form.get('numero_os', '').strip()
    num_os_digitos = re.sub(r'\D', '', num_os_raw)
    if len(num_os_digitos) == 8:
        nome_subpasta_os = num_os_digitos
    elif num_os_digitos:
        nome_subpasta_os = num_os_digitos
    elif num_os_raw:
        nome_subpasta_os = secure_filename(num_os_raw)
    else:
        num_referencia = dados.get('numero_nfse') or dados.get('numero_danfe') or f"OS{datetime.now().strftime('%Y%m%d%H%M%S')}"
        nome_subpasta_os = "".join(c for c in str(num_referencia) if c.isalnum() or c in ('-', '_'))

    dados['numero_os'] = nome_subpasta_os

    # Salva a Manutenção no SQLite
    manutencao_id = database.salvar_manutencao(dados)

    pasta_viatura_abs = os.path.join(PASTAS_VIATURAS_FOLDER, secure_filename(viatura_prefixo))
    pasta_os_abs = os.path.join(pasta_viatura_abs, nome_subpasta_os)
    os.makedirs(pasta_os_abs, exist_ok=True)

    # 4. Upload e armazenamento dos arquivos físicos na subpasta da OS da viatura
    # Ordem de Serviço
    arq_os = request.files.get('arquivo_os')
    if arq_os and arq_os.filename != '':
        salvar_arquivo_na_pasta_viatura(arq_os, viatura_prefixo, 'OS', descricao='Ordem de Serviço anexada', manutencao_id=manutencao_id, pasta_os=nome_subpasta_os)

    # Nota Fiscal de Material (DANFE)
    arq_danfe = request.files.get('arquivo_danfe')
    if arq_danfe and arq_danfe.filename != '':
        salvar_arquivo_na_pasta_viatura(arq_danfe, viatura_prefixo, 'DANFE', descricao=f"DANFE nº {dados['numero_danfe']}", manutencao_id=manutencao_id, pasta_os=nome_subpasta_os)

    # Nota Fiscal de Serviços (NFS-e)
    arq_nfse = request.files.get('arquivo_nfse')
    if arq_nfse and arq_nfse.filename != '':
        salvar_arquivo_na_pasta_viatura(arq_nfse, viatura_prefixo, 'NFSE', descricao=f"NFS-e nº {dados['numero_nfse']}", manutencao_id=manutencao_id, pasta_os=nome_subpasta_os)

    # 5. Geração do PDF do Atestado de Recebimento dentro da subpasta da OS
    from services.atestado_recebimento_generator import gerar_atestado_recebimento_pdf
    texto_personalizado = request.form.get('texto_atestado', '').strip() or None

    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    nome_pdf_atestado = f"ATESTADO_RECEBIMENTO_{viatura_prefixo}_{timestamp}.pdf"
    caminho_pdf_atestado = os.path.join(pasta_os_abs, nome_pdf_atestado)

    gerar_atestado_recebimento_pdf(caminho_pdf_atestado, dados, texto_personalizado=texto_personalizado)

    # 6. Registra o próprio PDF do Atestado na pasta da viatura e na OS
    caminho_rel_atestado = os.path.join('uploads', 'manutencao', 'viaturas', secure_filename(viatura_prefixo), nome_subpasta_os, nome_pdf_atestado).replace('\\', '/')
    database.salvar_documento_viatura(
        viatura_prefixo=viatura_prefixo,
        tipo_documento='ATESTADO',
        nome_original=nome_pdf_atestado,
        nome_arquivo=nome_pdf_atestado,
        caminho_relativo=caminho_rel_atestado,
        tamanho_bytes=os.path.getsize(caminho_pdf_atestado),
        descricao=f"Atestado de Recebimento gerado para a viatura BM {viatura_prefixo} ({nome_subpasta_os})",
        manutencao_id=manutencao_id,
        pasta_os=nome_subpasta_os
    )

    # Atualiza o registro da manutenção com o caminho do atestado
    database.atualizar_atestado_manutencao(manutencao_id, caminho_rel_atestado)

    flash(f"Atestado de Recebimento gerado com sucesso para a viatura BM {viatura_prefixo}!", "success")
    return send_file(
        caminho_pdf_atestado,
        as_attachment=True,
        download_name=f"Atestado_Recebimento_{viatura_prefixo}_{dados['tipo_documento_fiscal'].upper()}.pdf",
        mimetype='application/pdf'
    )


@app.route('/api/manutencao/atestado/texto-preview', methods=['POST', 'GET'])
@login_required
def api_texto_preview_recebimento():
    """Retorna o texto padrão dinamicamente calculado para o Atestado de Recebimento."""
    from services.atestado_recebimento_generator import gerar_texto_padrao_recebimento
    dados = request.get_json(silent=True) or request.form or {}
    texto = gerar_texto_padrao_recebimento(dados)
    return jsonify({'texto': texto})


@app.route('/api/manutencao/viatura-info/<viatura_prefixo>')
@login_required
def api_viatura_info(viatura_prefixo):
    """Retorna informações de marca, modelo, placa e cidade de carga de uma viatura cadastrada."""
    viaturas = database.obter_veiculos_alvo()
    pref_busca = str(viatura_prefixo).strip().upper()
    for v in viaturas:
        if str(v.get('prefixo') or '').strip().upper() == pref_busca or str(v.get('placa') or '').strip().upper() == pref_busca:
            return jsonify({
                'prefixo': v.get('prefixo') or pref_busca,
                'placa': v.get('placa') or '',
                'marca': v.get('marca') or '',
                'modelo': v.get('modelo') or '',
                'cidade_carga': v.get('cidade_carga') or 'Bagé'
            })
    
    # Fallback procurando em manutenções anteriores
    manuts = database.obter_manutencoes(filtro_viatura=pref_busca)
    if manuts:
        m = manuts[0]
        return jsonify({
            'prefixo': m.get('viatura_prefixo') or pref_busca,
            'placa': m.get('viatura_placa') or '',
            'marca': m.get('viatura_marca') or '',
            'modelo': m.get('viatura_modelo') or '',
            'cidade_carga': m.get('cidade_carga') or 'Bagé'
        })

    return jsonify({
        'prefixo': pref_busca,
        'placa': '',
        'marca': '',
        'modelo': '',
        'cidade_carga': 'Bagé'
    })


@app.route('/api/manutencao/viatura-arquivos/<viatura_prefixo>')
@login_required
def api_viatura_arquivos(viatura_prefixo):
    """Retorna a listagem de arquivos anexados na pasta da viatura."""
    arquivos = database.obter_arquivos_viatura(viatura_prefixo)
    return jsonify({'arquivos': arquivos})


@app.route('/manutencao/salvar', methods=['POST'])
@login_required
def salvar_manutencao_route():
    """Cadastra ou atualiza uma Ordem de Serviço / Manutenção."""
    dados = {
        'id': request.form.get('id') or None,
        'data_servico': request.form.get('data_servico', '').strip(),
        'tipo': request.form.get('tipo', 'Preventiva').strip(),
        'viatura_prefixo': request.form.get('viatura_prefixo', '').strip().upper(),
        'viatura_placa': request.form.get('viatura_placa', '').strip().upper(),
        'viatura_modelo': request.form.get('viatura_modelo', '').strip(),
        'odometro': request.form.get('odometro', '0').strip().replace(',', '.'),
        'oficina': request.form.get('oficina', '').strip(),
        'descricao_servico': request.form.get('descricao_servico', '').strip(),
        'pecas': request.form.get('pecas', '').strip(),
        'valor_total': request.form.get('valor_total', '0').strip().replace('.', '').replace(',', '.'),
        'status': request.form.get('status', 'Concluída').strip(),
        'data_conclusao': request.form.get('data_conclusao', '').strip() or None,
        'responsavel': request.form.get('responsavel', '').strip(),
        'observacoes': request.form.get('observacoes', '').strip()
    }

    if not dados['data_servico'] or not dados['descricao_servico']:
        flash('Data do serviço e descrição são obrigatórios.', 'warning')
        return redirect(url_for('modulo_manutencao'))

    database.salvar_manutencao(dados)
    flash('Ordem de serviço / manutenção registrada com sucesso!', 'success')
    return redirect(url_for('modulo_manutencao'))


@app.route('/manutencao/excluir/<int:manutencao_id>', methods=['POST'])
@login_required
def excluir_manutencao_route(manutencao_id):
    """Exclui um registro de manutenção."""
    database.excluir_manutencao(manutencao_id)
    flash('Registro de manutenção excluído com sucesso!', 'success')
    return redirect(url_for('modulo_manutencao'))


@app.route('/equipamento/salvar', methods=['POST'])
@login_required
def salvar_equipamento_route():
    """Cadastra ou atualiza um equipamento patrimoniado."""
    dados = {
        'id': request.form.get('id') or None,
        'patrimonio': request.form.get('patrimonio', '').strip().upper(),
        'nome': request.form.get('nome', '').strip(),
        'tipo': request.form.get('tipo', 'Outro').strip(),
        'modelo': request.form.get('modelo', '').strip(),
        'serie': request.form.get('serie', '').strip(),
        'viatura_prefixo': request.form.get('viatura_prefixo', '').strip().upper(),
        'fracao': request.form.get('fracao', '46º BPM').strip(),
        'status': request.form.get('status', 'Operacional').strip(),
        'ultima_revisao': request.form.get('ultima_revisao', '').strip(),
        'observacoes': request.form.get('observacoes', '').strip()
    }

    if not dados['patrimonio'] or not dados['nome']:
        flash('Código de patrimônio e nome do equipamento são obrigatórios.', 'warning')
        return redirect(url_for('modulo_manutencao'))

    database.salvar_equipamento(dados)
    flash(f'Equipamento {dados["patrimonio"]} - {dados["nome"]} salvo com sucesso!', 'success')
    return redirect(url_for('modulo_manutencao'))


@app.route('/equipamento/excluir/<int:eq_id>', methods=['POST'])
@login_required
def excluir_equipamento_route(eq_id):
    """Exclui um equipamento pelo ID."""
    database.excluir_equipamento(eq_id)
    flash('Equipamento excluído com sucesso!', 'success')
    return redirect(url_for('modulo_manutencao'))


# ==============================================================================
# MÓDULO 3: GESTÃO DE PNEUS
# ==============================================================================

@app.route('/pneus')
@login_required
def modulo_pneus():
    """Painel do Módulo de Gestão e Controle de Pneus da Frota."""
    filtro_status = request.args.get('status', 'Todos')
    filtro_viatura = request.args.get('viatura', 'Todos')

    pneus = database.obter_pneus(filtro_status=filtro_status, filtro_viatura=filtro_viatura)
    stats = database.obter_estatisticas_pneus()
    viaturas_cadastradas = database.obter_veiculos_alvo()

    return render_template(
        'pneus.html',
        pneus=pneus,
        stats=stats,
        viaturas_cadastradas=viaturas_cadastradas,
        filtro_status=filtro_status,
        filtro_viatura=filtro_viatura
    )


@app.route('/pneus/salvar', methods=['POST'])
@login_required
def salvar_pneu_route():
    """Cadastra ou atualiza um pneu."""
    dados = {
        'id': request.form.get('id') or None,
        'numero_fogo': request.form.get('numero_fogo', '').strip().upper(),
        'marca': request.form.get('marca', '').strip(),
        'modelo': request.form.get('modelo', '').strip(),
        'medida': request.form.get('medida', '').strip(),
        'dot': request.form.get('dot', '').strip(),
        'viatura_prefixo': request.form.get('viatura_prefixo', '').strip().upper() or None,
        'posicao': request.form.get('posicao', 'Estoque').strip(),
        'km_instalacao': request.form.get('km_instalacao', '0').strip().replace(',', '.'),
        'km_atual': request.form.get('km_atual', '0').strip().replace(',', '.'),
        'sulco_mm': request.form.get('sulco_mm', '8.0').strip().replace(',', '.'),
        'status': request.form.get('status', '').strip() or None,
        'data_instalacao': request.form.get('data_instalacao', '').strip(),
        'observacoes': request.form.get('observacoes', '').strip()
    }

    if not dados['numero_fogo']:
        flash('O número de fogo/série do pneu é obrigatório.', 'warning')
        return redirect(url_for('modulo_pneus'))

    database.salvar_pneu(dados)
    flash(f'Pneu {dados["numero_fogo"]} salvo com sucesso!', 'success')
    return redirect(url_for('modulo_pneus'))


@app.route('/pneus/medicao', methods=['POST'])
@login_required
def medicao_pneu_route():
    """Registra uma medição de sulco ou rodízio de pneu."""
    pneu_id = request.form.get('pneu_id')
    dados = {
        'pneu_id': int(pneu_id) if pneu_id else None,
        'data_movimentacao': request.form.get('data_movimentacao', datetime.now().strftime('%Y-%m-%d')).strip(),
        'tipo': request.form.get('tipo', 'Medição de Sulco').strip(),
        'viatura': request.form.get('viatura', '').strip().upper() or None,
        'posicao': request.form.get('posicao', '').strip(),
        'odometro': request.form.get('odometro', '0').strip().replace(',', '.'),
        'sulco_mm': request.form.get('sulco_mm', '0').strip().replace(',', '.'),
        'observacao': request.form.get('observacao', '').strip(),
        'responsavel': request.form.get('responsavel', '').strip()
    }

    if not dados['pneu_id']:
        flash('Selecione o pneu para registrar a medição/movimentação.', 'warning')
        return redirect(url_for('modulo_pneus'))

    database.registrar_movimentacao_pneu(dados)
    flash('Medição/movimentação de pneu registrada com sucesso!', 'success')
    return redirect(url_for('modulo_pneus'))


@app.route('/pneus/excluir/<int:pneu_id>', methods=['POST'])
@login_required
def excluir_pneu_route(pneu_id):
    """Exclui um pneu pelo ID."""
    database.excluir_pneu(pneu_id)
    flash('Pneu excluído com sucesso!', 'success')
    return redirect(url_for('modulo_pneus'))


if __name__ == '__main__':
    porta = int(os.environ.get('PORT', 5000))
    modo_debug = os.environ.get('FLASK_DEBUG', 'false').lower() in ('true', '1')
    print("=" * 65)
    print("  46º BATALHÃO DE POLÍCIA MILITAR - BAGÉ / RS")
    print("  Sistema de Controle e Relatórios de Consumo de Viaturas")
    print(f"  Servidor iniciado em: http://0.0.0.0:{porta}")
    print(f"  Modo de execução: {'Desenvolvimento (Debug)' if modo_debug else 'Produção'}")
    print("=" * 65)
    app.run(host='0.0.0.0', port=porta, debug=modo_debug)
