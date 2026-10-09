"""
redefinir_frota_oficial.py
Redefine a base de viaturas, limpa ordens de serviço, documentos e pastas físicas,
mantendo estritamente a relação de viaturas e equipamentos em carga fornecida em Excel pelo 46º BPM.
"""

import os
import shutil
import sqlite3
import openpyxl

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, 'bpm46_combustivel.db')
PASTAS_VIATURAS = os.path.join(BASE_DIR, 'uploads', 'manutencao', 'viaturas')
EXCEL_ORIGEM = r'C:\Users\rsman\.gemini\antigravity\brain\f8aefbe2-7410-43fc-a90c-1288549ecad6\.user_uploaded\media_1790887940131.xlsx'
EXCEL_DESTINO = os.path.join(BASE_DIR, 'exemplos', 'frota_oficial_46bpm_em_carga.xlsx')

def main():
    print("="*60)
    print(" INICIANDO REDEFINIÇÃO CONFORME PLANILHA EXCEL OFICIAL")
    print("="*60)

    # 1. Copiar planilha oficial para pasta de exemplos do projeto
    if os.path.exists(EXCEL_ORIGEM):
        os.makedirs(os.path.dirname(EXCEL_DESTINO), exist_ok=True)
        shutil.copy2(EXCEL_ORIGEM, EXCEL_DESTINO)
        print(f"[+] Planilha oficial copiada para: {EXCEL_DESTINO}")
    else:
        print(f"[!] Planilha original não encontrada em: {EXCEL_ORIGEM}")
        return

    # 2. Ler viaturas da planilha Excel
    wb = openpyxl.load_workbook(EXCEL_DESTINO)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))

    frota_oficial = []
    for r in rows[2:]:
        ordem, pref, placa, cartao, cidade, marca_mod = r[0], r[1], r[2], r[3], r[4], r[5]
        if not placa:
            continue
        placa_str = str(placa).strip().upper()
        pref_str = str(pref).strip() if pref is not None else placa_str

        # Se for equipamento (EQP), ignora completamente
        if 'EQP' in placa_str or 'EQP' in pref_str or 'GRAMA' in str(marca_mod).upper() or ordem is None:
            continue

        cidade_str = str(cidade).strip().title() if cidade else 'Bagé'
        marca_mod_str = str(marca_mod).strip() if marca_mod else ''
        cartao_str = str(cartao).strip() if cartao else ''

        mm = marca_mod_str.strip()
        if mm.upper().startswith('AGRALE'):
            marca = 'Agrale'
            modelo = mm.split('/', 1)[-1].strip()
        elif mm.upper().startswith('NISSAN'):
            marca = 'Nissan'
            modelo = mm.split('/', 1)[-1].strip()
        elif mm.upper().startswith('MARCOPOLO'):
            marca = 'Marcopolo'
            modelo = mm.split('/', 1)[-1].strip()
        elif mm.upper().startswith('YAMAHA'):
            marca = 'Yamaha'
            modelo = mm.split('/', 1)[-1].strip()
        elif mm.upper().startswith('CHEVROLET'):
            marca = 'Chevrolet'
            modelo = mm.split('/', 1)[-1].strip()
        elif mm.upper().startswith('TOYOTA'):
            marca = 'Toyota'
            partes = [p.strip() for p in mm.split('/') if p.strip().upper() != 'TOYOTA']
            modelo = ' '.join(partes)
        elif mm.upper().startswith('FIAT'):
            marca = 'Fiat'
            modelo = mm.split('/', 1)[-1].strip()
        elif mm.upper().startswith('RENAULT'):
            marca = 'Renault'
            modelo = mm.split('/', 1)[-1].strip()
        elif mm.upper().startswith('VW'):
            marca = 'Volkswagen'
            modelo = mm.split('/', 1)[-1].strip()
        elif mm.upper().startswith('REB'):
            marca = 'Reboque'
            modelo = mm.split('/', 1)[-1].strip()
        elif '/' in mm:
            partes = mm.split('/', 1)
            marca = partes[0].strip().title()
            modelo = partes[1].strip()
        else:
            marca = ''
            modelo = mm

        # Determina combustível
        mm_upper = marca_mod_str.upper()
        if any(k in mm_upper for k in ['HILUX', 'HILLUX', 'FRONTIER', 'MAXIBUS', 'VOLARE', 'S10', 'TRATOR']):
            combustivel = 'Diesel S10'
        elif any(k in mm_upper for k in ['REB', 'PREMIUM']):
            combustivel = 'Outro'
        else:
            combustivel = 'Gasolina Comum'

        frota_oficial.append({
            'ordem': ordem,
            'prefixo': pref_str,
            'placa': placa_str,
            'marca': marca,
            'modelo': modelo,
            'cidade': cidade_str,
            'cartao': cartao_str,
            'fracao': f"46º BPM - {cidade_str}",
            'combustivel': combustivel
        })

    # Ordena frota em ordem crescente de prefixo
    def _ordem_pref(x):
        p = str(x.get('prefixo', '')).strip()
        return (0, int(p)) if p.isdigit() else (1, p)
    frota_oficial.sort(key=_ordem_pref)

    print(f"[+] Total de viaturas oficiais identificadas no Excel: {len(frota_oficial)}")

    # 3. Limpar tabelas no Banco de Dados SQLite
    conn = sqlite3.connect(DB_PATH, timeout=30)
    cursor = conn.cursor()

    # Deletar documentos de manutenção e manutenções/ordens de serviço
    cursor.execute("DELETE FROM documentos_manutencao")
    qtd_docs = cursor.rowcount
    print(f"[+] Documentos de manutenção excluídos do banco: {qtd_docs}")

    cursor.execute("DELETE FROM manutencoes")
    qtd_manut = cursor.rowcount
    print(f"[+] Ordens de serviço / manutenções excluídas do banco: {qtd_manut}")

    cursor.execute("DELETE FROM equipamentos")
    qtd_eq = cursor.rowcount
    print(f"[+] Equipamentos excluídos do banco: {qtd_eq}")

    # Deletar veiculos_alvo e recadastrar exclusivamente os da planilha oficial
    cursor.execute("DELETE FROM veiculos_alvo")
    print("[+] Tabela veiculos_alvo limpa.")

    for v in frota_oficial:
        cursor.execute('''
            INSERT INTO veiculos_alvo (
                prefixo, placa, marca, modelo, cidade_carga,
                cartao_abastecimento, fracao, tipo_combustivel, ativo
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, 1)
        ''', (
            v['prefixo'], v['placa'], v['marca'], v['modelo'],
            v['cidade'], v['cartao'], v['fracao'], v['combustivel']
        ))

    conn.commit()
    conn.close()
    print(f"[+] {len(frota_oficial)} viaturas/equipamentos da planilha cadastrados com sucesso no banco SQLite!")

    # 4. Limpar e recriar pastas físicas no disco
    if os.path.exists(PASTAS_VIATURAS):
        shutil.rmtree(PASTAS_VIATURAS, ignore_errors=True)
        print(f"[+] Pastas antigas em '{PASTAS_VIATURAS}' removidas completamente.")

    os.makedirs(PASTAS_VIATURAS, exist_ok=True)

    # Cria pastas individuais limpas para cada viatura oficial
    for v in frota_oficial:
        nome_pasta = v['prefixo'] if v['prefixo'] else v['placa']
        nome_pasta_limpo = "".join(c for c in nome_pasta if c.isalnum() or c in ('-', '_'))
        caminho_pasta = os.path.join(PASTAS_VIATURAS, nome_pasta_limpo)
        os.makedirs(caminho_pasta, exist_ok=True)

    print(f"[+] {len(frota_oficial)} pastas de viaturas criadas e limpas em '{PASTAS_VIATURAS}'.")
    print("="*60)
    print(" REDEFINIÇÃO CONCLUÍDA COM SUCESSO!")
    print("="*60)

if __name__ == '__main__':
    main()
