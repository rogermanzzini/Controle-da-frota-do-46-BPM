import os
import json
from typing import Dict, Any

# Caminho para o arquivo de configuração persistente
CONFIG_PATH = os.path.join(os.path.dirname(__file__), "report_config.json")

GRADUACOES_DISPONIVEIS = [
    "SD PM",
    "2º SGT PM",
    "1º SGT PM"
]

CARGOS_DISPONIVEIS = [
    "Auxiliar do P4/46º BPM",
    "Chefe do P4/46º BPM"
]

DEFAULT_CONFIG = {
    "nome": "RÓGER SANTANA MANZZINI",
    "graduacao": "SD PM",
    "cargo": "Auxiliar do P4/46º BPM"
}


def get_assinatura_config() -> Dict[str, str]:
    """Retorna os dados completos da assinatura padrão do relatório."""
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
                nome = data.get("nome") or data.get("signature") or DEFAULT_CONFIG["nome"]
                graduacao = data.get("graduacao") or DEFAULT_CONFIG["graduacao"]
                cargo = data.get("cargo") or DEFAULT_CONFIG["cargo"]
                return {
                    "nome": nome.strip(),
                    "graduacao": graduacao.strip(),
                    "cargo": cargo.strip()
                }
        except Exception:
            pass
    return DEFAULT_CONFIG.copy()


def set_assinatura_config(nome: str, graduacao: str, cargo: str) -> None:
    """Salva os novos dados da assinatura padrão no arquivo JSON."""
    data = {
        "nome": (nome or DEFAULT_CONFIG["nome"]).strip(),
        "graduacao": (graduacao or DEFAULT_CONFIG["graduacao"]).strip(),
        "cargo": (cargo or DEFAULT_CONFIG["cargo"]).strip(),
        "signature": f"{(nome or DEFAULT_CONFIG['nome']).strip()} – {(graduacao or DEFAULT_CONFIG['graduacao']).strip()}"
    }
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def get_signature() -> str:
    """Compatibilidade retroativa: retorna a linha principal da assinatura."""
    cfg = get_assinatura_config()
    return f"{cfg['nome']} – {cfg['graduacao']}"


def set_signature(name: str) -> None:
    """Compatibilidade retroativa."""
    set_assinatura_config(name, DEFAULT_CONFIG["graduacao"], DEFAULT_CONFIG["cargo"])


DEFAULT_ATESTADO_CONFIG = {
    "modo": "impedimento",
    "titular_nome": "KARLA DE MOURA INCERTI",
    "titular_cargo": "Ten Cel PM – Cmt 46º BPM",
    "respondendo_nome": "PATRIQUE ROLIM MARQUES",
    "respondendo_cargo": "Maj PM – Rsp. p/ Cmt 46º BPM",
    "subcomandante_nome": "PATRIQUE ROLIM MARQUES",
    "subcomandante_cargo": "Maj PM – SCmt 46º BPM",
    "gestor_nome": "RÓGER SANTANA MANZZINI",
    "gestor_cargo": "SD PM – Gestor de frota"
}


def get_atestado_config() -> Dict[str, str]:
    """Retorna as configurações e nomes das assinaturas do Atestado de Combustíveis."""
    cfg = DEFAULT_ATESTADO_CONFIG.copy()
    # Se o nome/graduação do militar gestor estiver configurado, usa-o
    ass_cfg = get_assinatura_config()
    cfg['gestor_nome'] = ass_cfg.get('nome') or cfg['gestor_nome']
    cfg['gestor_cargo'] = f"{ass_cfg.get('graduacao', 'SD PM')} – Gestor de frota"

    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
                atestado_data = data.get("atestado")
                if isinstance(atestado_data, dict):
                    for k in cfg:
                        if k in atestado_data and atestado_data[k]:
                            cfg[k] = str(atestado_data[k]).strip()
        except Exception:
            pass
    return cfg


def set_atestado_config(dados: Dict[str, str]) -> None:
    """Salva as opções de assinatura do Atestado no arquivo report_config.json."""
    data = {}
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception:
            data = {}

    current = data.get("atestado", {})
    if not isinstance(current, dict):
        current = {}

    for k, v in dados.items():
        if v is not None:
            current[k] = str(v).strip()

    data["atestado"] = current
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


DEFAULT_ATESTADOR_CONFIG = {
    "nome": "JULIANO INOCENTE VELEDA",
    "graduacao": "Sd PM",
    "id_func": "3155358",
    "cargo": "Aux P4/46º BPM"
}

# -------------------------------------------------
# Empresa configuration for atestado de recebimento
# -------------------------------------------------
DEFAULT_EMPRESA_CONFIG = {
    "empresa_nome": "",
    "empresa_cnpj": ""
}


def get_empresa_config() -> Dict[str, str]:
    """Retorna a configuração padrão da empresa (nome e CNPJ) para o atestado de recebimento."""
    cfg = DEFAULT_EMPRESA_CONFIG.copy()
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
                empresa_data = data.get("empresa")
                if isinstance(empresa_data, dict):
                    for k in cfg:
                        if k in empresa_data and empresa_data[k]:
                            cfg[k] = str(empresa_data[k]).strip()
        except Exception:
            pass
    return cfg


def set_empresa_config(nome: str, cnpj: str) -> None:
    """Salva a configuração padrão da empresa no arquivo JSON."""
    data = {}
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception:
            data = {}
    data["empresa"] = {
        "empresa_nome": (nome or "").strip(),
        "empresa_cnpj": (cnpj or "").strip()
    }
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def get_atestador_recebimento_config() -> Dict[str, str]:
    """Retorna a configuração do militar responsável pelo atestado de recebimento (Manutenção)."""
    cfg = DEFAULT_ATESTADOR_CONFIG.copy()
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
                atestador_data = data.get("atestador_recebimento")
                if isinstance(atestador_data, dict):
                    for k in cfg:
                        if k in atestador_data and atestador_data[k]:
                            cfg[k] = str(atestador_data[k]).strip()
        except Exception:
            pass
    return cfg


def set_atestador_recebimento_config(nome: str, graduacao: str, id_func: str, cargo: str) -> None:
    """Salva as configurações do militar atestador de recebimento no arquivo JSON."""
    data = {}
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception:
            data = {}

    data["atestador_recebimento"] = {
        "nome": (nome or DEFAULT_ATESTADOR_CONFIG["nome"]).strip().upper(),
        "graduacao": (graduacao or DEFAULT_ATESTADOR_CONFIG["graduacao"]).strip(),
        "id_func": (id_func or DEFAULT_ATESTADOR_CONFIG["id_func"]).strip(),
        "cargo": (cargo or DEFAULT_ATESTADOR_CONFIG["cargo"]).strip()
    }

    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

