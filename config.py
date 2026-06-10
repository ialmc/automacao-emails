# -*- coding: utf-8 -*-
"""
config.py — Configurações Gerais da Automação de E-mails
======================================================
Define constantes globais e mapeamentos de relatórios.
Tenta carregar o arquivo local config_local.py para sobrescrever com caminhos e assuntos reais.
"""

import re
from datetime import date
from pathlib import Path

# ══════════════════════════════════════════════════════════════════════════════
# PATHS DE DIRETÓRIOS E ARQUIVOS (Padrão Genérico)
# ══════════════════════════════════════════════════════════════════════════════
# Caminho do diretório do script (resolvido dinamicamente de forma portátil)
SCRIPT_DIR = Path(__file__).parent.resolve()

# Diretório base dos downloads de dados
BASE_ALLPE = SCRIPT_DIR / "data"

# Diretório base do histórico consolidado
BASE_HISTORICO = SCRIPT_DIR / "history"

LOG_DIR = SCRIPT_DIR / "logs"
LOG_FILE = LOG_DIR / "email_downloader.log"

ARQ_LINKS_JSON = BASE_ALLPE / "links_baixados.json"
ARQ_INVENTARIO_JSON = BASE_ALLPE / "inventory_dates.json"
ARQ_LOCK = SCRIPT_DIR / ".coletor.lock"

# ══════════════════════════════════════════════════════════════════════════════
# CONTA OUTLOOK E CONFIGURAÇÃO DE RELATÓRIO
# ══════════════════════════════════════════════════════════════════════════════
NOME_CONTA = "your-email@domain.com"
REPORT_DEST_EMAIL = "alerts@domain.com"
REPORT_TITLE = "Relatório de Ingestão de E-mails"
REPORT_SUBJECT = "Relatório Diário de Cobertura"

# ══════════════════════════════════════════════════════════════════════════════
# DATA DE CORTE GLOBAL
# ══════════════════════════════════════════════════════════════════════════════
# Ignora e-mails recebidos antes desta data
DATA_CORTE_DOWNLOAD_DATE: date = date(2026, 1, 1)

# ══════════════════════════════════════════════════════════════════════════════
# CONFIGURAÇÕES DE DOWNLOAD HTTP
# ══════════════════════════════════════════════════════════════════════════════
REQUEST_TIMEOUT = 90  # segundos por tentativa
REQUEST_RETRIES = 3  # número de tentativas
REQUEST_SLEEP_BETWEEN_RETRIES = 2  # segundos entre tentativas
REQUEST_HEADERS = {
    "User-Agent": "EmailRpaCollector/2.0"
}

# ══════════════════════════════════════════════════════════════════════════════
# REGEX S3 — PADRÃO DE URL DE DOWNLOAD (Amazon S3)
# ══════════════════════════════════════════════════════════════════════════════
# Regex genérica para links de buckets S3 da AWS
S3_REGEX = re.compile(
    r"https://[a-zA-Z0-9.-]+\.s3[.-]*(?:dualstack\.)?[a-zA-Z0-9.-]+\.amazonaws\.com/[^\s\"\'<>]+",
    re.IGNORECASE,
)

# ══════════════════════════════════════════════════════════════════════════════
# MAPEAMENTO DE ASSUNTOS → PASTAS OUTLOOK E PASTAS LOCAIS (Exemplos Genéricos)
# ══════════════════════════════════════════════════════════════════════════════
# Mapeia o assunto exato do e-mail para:
#   - outlook_pasta: subpasta dentro do Inbox onde os e-mails chegam
#   - destino_pasta: pasta física local onde os relatórios baixados serão salvos
ASSUNTOS_PASTAS = {
    '[System][Report] Daily Services Report': {
        'outlook_pasta': 'Services',
        'destino_pasta': BASE_ALLPE / "Services",
    },
    '[System][Report] Daily Products Report': {
        'outlook_pasta': 'Products',
        'destino_pasta': BASE_ALLPE / "Products",
    },
}

# ══════════════════════════════════════════════════════════════════════════════
# PREFIXOS FORÇADOS — Nome canônico dos arquivos baixados
# ══════════════════════════════════════════════════════════════════════════════
# Nome final do arquivo salvo localmente: <prefixo>-DD-MM-YYYY.<ext>
PREFIXOS_FORCADOS = {
    '[System][Report] Daily Services Report': 'report_services',
    '[System][Report] Daily Products Report': 'report_products',
}

# Mapa reverso
PREFIXO_TO_ASSUNTO: dict[str, str] = {v: k for k, v in PREFIXOS_FORCADOS.items()}

# ══════════════════════════════════════════════════════════════════════════════
# DISPLAY_MAP — Metadados de exibição para auditoria de SLA
# ══════════════════════════════════════════════════════════════════════════════
DISPLAY_MAP: dict[str, dict] = {
    "Services": {"nome": "Services", "freq": "Diária", "rag": 1.0},
    "Products": {"nome": "Products", "freq": "Diária", "rag": 1.0},
}

# ══════════════════════════════════════════════════════════════════════════════
# MAPA DE PASTAS E HISTÓRICO (data_manager.py)
# ══════════════════════════════════════════════════════════════════════════════
MAPA_PASTAS: dict[str, Path] = {
    "Services": BASE_ALLPE / "Services",
    "Products": BASE_ALLPE / "Products",
}

MAPA_HISTORICO: dict[str, list[str]] = {
    "Services": ["history_services.xlsx"],
    "Products": ["history_products.xlsx"],
}

REPORTER_KEYWORDS: dict[str, list[str]] = {
    "Services": ["report_services"],
    "Products": ["report_products"],
}

# Relatórios diários (usados no cálculo de SLA D-1/D-2)
DAILY_KEYS = [
    "Services",
    "Products",
]

# ══════════════════════════════════════════════════════════════════════════════
# CARREGAMENTO DE SOBREPOSIÇÃO LOCAL (LOCAL OVERRIDES)
# ══════════════════════════════════════════════════════════════════════════════
# Se o arquivo config_local.py existir (ignorado no Git), ele substitui
# todas as variáveis acima pelas configurações privadas locais da automação.
try:
    import config_local

    for _key in dir(config_local):
        if not _key.startswith("__"):
            globals()[_key] = getattr(config_local, _key)
except ImportError:
    pass
