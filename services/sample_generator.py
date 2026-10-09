"""
services/sample_generator.py - Gerador de Planilhas de Teste e Demonstração
Cria planilhas Excel realistas:
1. Planilha geral do RS com múltiplos batalhões e cidades.
2. Planilha alvo de viaturas exclusivas do 46º BPM (Bagé/RS).
"""

import os
import random
from datetime import datetime, timedelta
import pandas as pd


def gerar_planilhas_exemplo(pasta_destino: str):
    """Gera os dois arquivos Excel de demonstração na pasta especificada."""
    os.makedirs(pasta_destino, exist_ok=True)

    caminho_viaturas_alvo = os.path.join(pasta_destino, 'viaturas_46bpm_alvo_exemplo.xlsx')
    caminho_abastecimentos_rs = os.path.join(pasta_destino, 'abastecimentos_rs_geral_exemplo.xlsx')

    # 1. Viaturas Alvo do 46º BPM (Bagé)
    viaturas_46bpm = [
        {"prefixo": "BM-12450", "placa": "IXO8A12", "modelo": "Renault Duster 1.6", "fracao": "1ª Companhia - Bagé", "tipo_combustivel": "Gasolina Comum"},
        {"prefixo": "BM-12451", "placa": "IXO8A15", "modelo": "Renault Duster 1.6", "fracao": "1ª Companhia - Bagé", "tipo_combustivel": "Gasolina Comum"},
        {"prefixo": "BM-13100", "placa": "JBG9A87", "modelo": "Toyota Hilux 4x4", "fracao": "Pelotão Força Tática", "tipo_combustivel": "Diesel S10"},
        {"prefixo": "BM-13105", "placa": "JBG9A90", "modelo": "Toyota Hilux 4x4", "fracao": "Patrulha Rural - Bagé", "tipo_combustivel": "Diesel S10"},
        {"prefixo": "BM-14200", "placa": "JCY4B11", "modelo": "Toyota Corolla 2.0", "fracao": "Comando / Estado-Maior", "tipo_combustivel": "Gasolina Aditivada"},
        {"prefixo": "BM-14205", "placa": "JCY4B15", "modelo": "Toyota Corolla 2.0", "fracao": "1ª Cia - Rádio Patrulha", "tipo_combustivel": "Gasolina Comum"},
        {"prefixo": "BM-15010", "placa": "JZZ7C34", "modelo": "Chevrolet Spin 1.8", "fracao": "2ª Companhia - Caçapava", "tipo_combustivel": "Gasolina Comum"},
        {"prefixo": "BM-15012", "placa": "JZZ7C40", "modelo": "Chevrolet Spin 1.8", "fracao": "2ª Companhia - Lavras do Sul", "tipo_combustivel": "Gasolina Comum"}
    ]

    df_viaturas = pd.DataFrame(viaturas_46bpm)
    df_viaturas.rename(columns={
        "prefixo": "Prefixo Vtr",
        "placa": "Placa do Veículo",
        "modelo": "Modelo / Descrição",
        "fracao": "Fração / Unidade",
        "tipo_combustivel": "Tipo de Combustível Padrão"
    }, inplace=True)
    df_viaturas.to_excel(caminho_viaturas_alvo, index=False, engine='openpyxl')

    # 2. Outras viaturas do RS (para compor a base estadual ampla)
    outras_viaturas_rs = [
        {"prefixo": "BM-11001", "placa": "IQM1A10", "unidade": "1º BPM - Porto Alegre", "cidade": "Porto Alegre", "comb": "Gasolina Comum", "km_ini": 85000, "consumo": 8.5},
        {"prefixo": "BM-11002", "placa": "IQM1A12", "unidade": "9º BPM - Porto Alegre", "cidade": "Porto Alegre", "comb": "Gasolina Comum", "km_ini": 112000, "consumo": 8.0},
        {"prefixo": "BM-11005", "placa": "IRK2B45", "unidade": "11º BPM - Porto Alegre", "cidade": "Porto Alegre", "comb": "Diesel S10", "km_ini": 64000, "consumo": 9.2},
        {"prefixo": "BM-12100", "placa": "IXX3C33", "unidade": "4º BPM - Pelotas", "cidade": "Pelotas", "comb": "Gasolina Comum", "km_ini": 45000, "consumo": 8.8},
        {"prefixo": "BM-12105", "placa": "IXX3C38", "unidade": "4º BPM - Pelotas", "cidade": "Pelotas", "comb": "Diesel S10", "km_ini": 78000, "consumo": 9.5},
        {"prefixo": "BM-13400", "placa": "IYY5D55", "unidade": "12º BPM - Caxias do Sul", "cidade": "Caxias do Sul", "comb": "Gasolina Comum", "km_ini": 92000, "consumo": 7.8},
        {"prefixo": "BM-14050", "placa": "IZW6E66", "unidade": "1º RPMon - Santa Maria", "cidade": "Santa Maria", "comb": "Gasolina Comum", "km_ini": 105000, "consumo": 8.2},
        {"prefixo": "BM-15200", "placa": "JAB7F77", "unidade": "3º RPMon - Passo Fundo", "cidade": "Passo Fundo", "comb": "Diesel S10", "km_ini": 53000, "consumo": 9.0}
    ]

    postos_rs = {
        "Bagé": [
            "Posto da Guarnição Militar Bagé",
            "Auto Posto Shell Santa Tecla (Bagé)",
            "Posto Ipiranga Tiaraju (Bagé)",
            "Posto Petrobras Centro Bagé"
        ],
        "Porto Alegre": [
            "Posto Ipiranga Ipiranga Av",
            "Rede Marajó Farrapos POA",
            "Posto Shell Assis Brasil"
        ],
        "Pelotas": [
            "Posto Bento Gonçalves Pelotas",
            "Posto Dom Joaquim Pelotas"
        ],
        "Caxias do Sul": ["Posto Perimetral Norte Caxias"],
        "Santa Maria": ["Posto Medianeira Santa Maria"],
        "Passo Fundo": ["Posto Brasil Passo Fundo"]
    }

    precos_litro = {
        "Gasolina Comum": 6.19,
        "Gasolina Aditivada": 6.39,
        "Diesel S10": 6.09
    }

    # Montar simulação de 30 dias de abastecimentos
    registros_rs = []
    data_base = datetime.now() - timedelta(days=35)

    # Estado de odômetro das viaturas de Bagé
    estado_bage = {
        "BM-12450": {"km": 42150, "consumo": 8.7, "placa": "IXO8A12", "comb": "Gasolina Comum"},
        "BM-12451": {"km": 51200, "consumo": 8.3, "placa": "IXO8A15", "comb": "Gasolina Comum"},
        "BM-13100": {"km": 68400, "consumo": 9.4, "placa": "JBG9A87", "comb": "Diesel S10"},
        "BM-13105": {"km": 79100, "consumo": 9.1, "placa": "JBG9A90", "comb": "Diesel S10"},
        "BM-14200": {"km": 31000, "consumo": 10.5, "placa": "JCY4B11", "comb": "Gasolina Aditivada"},
        "BM-14205": {"km": 34500, "consumo": 10.1, "placa": "JCY4B15", "comb": "Gasolina Comum"},
        "BM-15010": {"km": 88300, "consumo": 8.9, "placa": "JZZ7C34", "comb": "Gasolina Comum"},
        "BM-15012": {"km": 94100, "consumo": 8.8, "placa": "JZZ7C40", "comb": "Gasolina Comum"}
    }

    # Gerar abastecimentos cronológicos para viaturas do 46º BPM (Bagé)
    for dia in range(35):
        dt_atual = data_base + timedelta(days=dia, hours=random.randint(6, 22), minutes=random.randint(0, 59))
        for pref, dados in estado_bage.items():
            # Abastece a cada 2 a 4 dias
            if random.random() < 0.35:
                litros = round(random.uniform(32.0, 52.0), 2)
                km_percorrido = round(litros * dados["consumo"] * random.uniform(0.93, 1.07), 0)
                dados["km"] += int(km_percorrido)

                tipo_comb = dados["comb"]
                preco_unit = precos_litro[tipo_comb] + round(random.uniform(-0.15, 0.15), 2)
                valor_total = round(litros * preco_unit, 2)
                posto = random.choice(postos_rs["Bagé"])

                registros_rs.append({
                    "Data Abastecimento": dt_atual.strftime("%d/%m/%Y %H:%M"),
                    "Placa": dados["placa"],
                    "Prefixo": pref,
                    "Odômetro": dados["km"],
                    "Litros": litros,
                    "Valor Total": valor_total,
                    "Valor Unitário": round(preco_unit, 2),
                    "Combustível": tipo_comb,
                    "Posto": posto,
                    "Cidade": "Bagé",
                    "Unidade": "46º BPM - Bagé",
                    "Motorista": f"Sd PM {random.choice(['Silva', 'Santos', 'Oliveira', 'Souza', 'Rodrigues', 'Ferreira', 'Pereira'])}"
                })

    # Gerar abastecimentos para outras viaturas do RS
    for vtr in outras_viaturas_rs:
        km_atual = vtr["km_ini"]
        for dia in range(35):
            if random.random() < 0.30:
                dt_atual = data_base + timedelta(days=dia, hours=random.randint(6, 21), minutes=random.randint(0, 59))
                litros = round(random.uniform(30.0, 50.0), 2)
                km_percorrido = round(litros * vtr["consumo"] * random.uniform(0.92, 1.08), 0)
                km_atual += int(km_percorrido)

                tipo_comb = vtr["comb"]
                preco_unit = precos_litro[tipo_comb] + round(random.uniform(-0.10, 0.10), 2)
                valor_total = round(litros * preco_unit, 2)
                posto = random.choice(postos_rs.get(vtr["cidade"], ["Posto Rodoviário RS"]))

                registros_rs.append({
                    "Data Abastecimento": dt_atual.strftime("%d/%m/%Y %H:%M"),
                    "Placa": vtr["placa"],
                    "Prefixo": vtr["prefixo"],
                    "Odômetro": km_atual,
                    "Litros": litros,
                    "Valor Total": valor_total,
                    "Valor Unitário": round(preco_unit, 2),
                    "Combustível": tipo_comb,
                    "Posto": posto,
                    "Cidade": vtr["cidade"],
                    "Unidade": vtr["unidade"],
                    "Motorista": f"Cb PM {random.choice(['Martins', 'Almeida', 'Costa', 'Ribeiro'])}"
                })

    # Embaralhar para simular ordem natural de transações de sistema estadual
    random.shuffle(registros_rs)
    df_rs = pd.DataFrame(registros_rs)
    df_rs.to_excel(caminho_abastecimentos_rs, index=False, engine='openpyxl')

    return caminho_abastecimentos_rs, caminho_viaturas_alvo
