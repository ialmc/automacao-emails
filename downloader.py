# -*- coding: utf-8 -*-
"""
downloader.py — Coletor Outlook RPA
===================================
Camada de Coleta: conecta ao Outlook via COM/MAPI, extrai links de download dos e-mails
e baixa os arquivos para as pastas correspondentes.

Fluxo de execução (chamado pelo Executar_Coletor_AllPe.cmd):
  1. Adquire lockfile (.coletor.lock) — aborta se outra instância estiver rodando
  2. Limpa arquivos .part órfãos de execuções anteriores interrompidas
  3. Conecta ao Outlook e navega pelas pastas configuradas em config.ASSUNTOS_PASTAS
  4. Extrai links S3 dos e-mails (href + regex com html.unescape)
  5. Baixa arquivos pendentes com retry automático (3 tentativas, 90s timeout)
  6. Salva o estado de deduplicação em links_baixados.JSON (escrita atômica)
  7. Aciona data_manager.py → reporter.py em sequência

Requisitos:
  pip install requests beautifulsoup4 pywin32

⛔ PARA IAs E DESENVOLVEDORES — LEIA ANTES DE MODIFICAR:
  - Constantes de configuração → edite config.py, NÃO este arquivo
  - Para adicionar novo relatório → leia docs/ARCHITECTURE.md §5
  - NÃO mescle "Estoque Diario" e "Estoque Mensal" em uma única entrada
  - NÃO simplifique o S3_REGEX para aceitar apenas um formato de URL
  - NÃO remova o html.unescape() em extrair_links() — é essencial para URLs pré-assinadas
"""

# ── Stdlib ─────────────────────────────────────────────────────────────────
import argparse
import json
import logging
import logging.handlers
import os
import re
import subprocess
import sys
import unicodedata
from datetime import datetime, timedelta, date
from html import unescape
from pathlib import Path
from time import sleep
from typing import Any
from urllib.parse import unquote, urlparse

# ── Third-party ─────────────────────────────────────────────────────────────
import requests
import win32com.client as win32

# ── Local ───────────────────────────────────────────────────────────────────
from config import (
    BASE_ALLPE,
    SCRIPT_DIR,
    LOG_DIR,
    LOG_FILE,
    ARQ_LINKS_JSON,
    ARQ_INVENTARIO_JSON,
    ARQ_LOCK,
    NOME_CONTA,
    DATA_CORTE_DOWNLOAD_DATE,
    REQUEST_TIMEOUT,
    REQUEST_RETRIES,
    REQUEST_SLEEP_BETWEEN_RETRIES,
    REQUEST_HEADERS,
    S3_REGEX,
    ASSUNTOS_PASTAS,
    PREFIXOS_FORCADOS,
    PREFIXO_TO_ASSUNTO,
    DISPLAY_MAP,
    DAILY_KEYS,
)
from detector_datas import extrair_data_do_arquivo, extract_date_from_filename


# ══════════════════════════════════════════════════════════════════════════════
# LOGGING — RotatingFileHandler (máx 5 MB, 7 backups)
# ══════════════════════════════════════════════════════════════════════════════
LOG_DIR.mkdir(parents=True, exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    handlers=[
        logging.handlers.RotatingFileHandler(
            LOG_FILE, maxBytes=5 * 1024 * 1024, backupCount=7, encoding="utf-8"
        ),
        logging.StreamHandler(),
    ],
)
logger = logging.getLogger("downloader")


# ══════════════════════════════════════════════════════════════════════════════
# HELPERS — TEXTO E NORMALIZAÇÃO
# ══════════════════════════════════════════════════════════════════════════════
_token_re = re.compile(r"^[a-z0-9]{8,}$", re.I)
_date_suffix_re = re.compile(r"^\d{2}-\d{2}-\d{4}$")


def normalize_text(s: str) -> str:
    """Normaliza texto: minúsculas, sem acentos, espaços colapsados."""
    if not s:
        return ""
    s = s.strip().casefold()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return re.sub(r"\s+", " ", s)


def sanitize_prefix(prefixo: str) -> str:
    prefixo = (prefixo or "").strip().lower().replace(" ", "_")
    prefixo = re.sub(r"[^a-z0-9_]+", "_", prefixo)
    prefixo = re.sub(r"_+", "_", prefixo).strip("_")
    return prefixo or "arquivo"


def prefixo_do_url(url: str) -> str:
    """Extrai o prefixo semântico de uma URL S3, removendo o token aleatório do fim."""
    p = urlparse(url)
    base = os.path.basename(unquote(p.path))
    stem, _ = os.path.splitext(base)
    stem = (stem or "").strip()
    # Padrão: nome-TOKENALEATORIO ou nome-A-TOKENALEATORIO
    m = re.match(r"^(?P<left>.+?)(?:[-_])A?(?:[-_])(?P<last>[A-Za-z0-9]{8,})$", stem)
    if m and _token_re.match(m.group("last")) and not _date_suffix_re.match(m.group("last")):
        stem = m.group("left")
    else:
        m2 = re.match(r"^(?P<left>.+?)(?:[-_])(?P<last>[A-Za-z0-9]{8,})$", stem)
        if m2 and _token_re.match(m2.group("last")) and not _date_suffix_re.match(m2.group("last")):
            stem = m2.group("left")
    return sanitize_prefix(stem)


def inferir_ext(url: str) -> str:
    """Infere extensão do arquivo a partir do path da URL (ignora query params)."""
    _, ext = os.path.splitext(unquote(urlparse(url).path))
    ext = (ext or ".csv").lower()
    return ext if ext in {".csv", ".txt", ".zip", ".xls", ".xlsx", ".json"} else ".csv"


def arquivo(ass_ref: str) -> str:
    """Retorna nome legível do assunto removendo prefixo de assunto e sufixo automatico."""
    s = ass_ref.replace("[Trinks][AllPe]", "").replace("[System][Report]", "").strip()
    return s.replace("- Email automatico", "").strip()


# ══════════════════════════════════════════════════════════════════════════════
# HELPERS — LÓGICA DE DATAS
# ══════════════════════════════════════════════════════════════════════════════
def data_no_corte_ou_depois(dt: datetime) -> bool:
    if not dt:
        return False
    return dt.date() >= DATA_CORTE_DOWNLOAD_DATE


def calcular_data_referencia(assunto_ref: str, email_date: datetime) -> datetime:
    """Calcula a data de referência dos dados com base no tipo de relatório.

    Consolidados/Formas de Pagamento → D-2 (dados chegam com 2 dias de atraso)
    Agendamentos/Despesas/Estoque    → D-1
    Demais (snapshots)               → data do e-mail
    """
    subject_lower = assunto_ref.lower()
    if any(x in subject_lower for x in ["consolidado de servicos", "consolidado de produtos",
                                         "consolidado de pacotes", "formas de pagamento"]):
        return email_date - timedelta(days=2)
    if any(x in subject_lower for x in ["agendamento", "despesa", "estoque diario"]):
        return email_date - timedelta(days=1)
    return email_date


def obter_chave_inventario(subject_ref: str) -> str:
    """Mapeia o assunto do e-mail para a chave correspondente em DISPLAY_MAP."""
    subject_lower = subject_ref.lower()
    if "servicos" in subject_lower:
        return "Lista Servicos" if ("lista" in subject_lower or "cadastro" in subject_lower) else "Consolidado de servicos"
    if "produtos" in subject_lower:
        return "Lista Produtos" if ("lista" in subject_lower or "relatorio" in subject_lower) else "Consolidado de produtos"
    if "pacotes" in subject_lower:
        return "Consolidado de pacotes"
    if "agendamento" in subject_lower:
        return "Agendamentos"
    if "despesa" in subject_lower:
        return "Despesas"
    if "estoque" in subject_lower:
        return "Estoque Mensal" if "mensal" in subject_lower else "Estoque Diario"
    if "pagamentos" in subject_lower or "forma" in subject_lower:
        return "Formas de Pagamento"
    if "cliente" in subject_lower:
        return "Clientes"
    if "profissionais" in subject_lower:
        return "Profissionais"
    return ""


def obter_datas_pendentes(key: str, inventory: dict, hoje: datetime) -> list[datetime]:
    """Calcula as datas diárias pendentes (seg–sex) desde DATA_CORTE até data_max."""
    if not key or key not in DAILY_KEYS:
        return []
    inv = inventory.get(key, {"dates": []})
    dates_covered = set(inv.get("dates", []))
    atraso_permitido = 2 if ("Consolidado" in key or "Pagamento" in key) else 1
    data_max = hoje.date() - timedelta(days=atraso_permitido)
    missing = []
    curr = DATA_CORTE_DOWNLOAD_DATE
    while curr <= data_max:
        if curr.isoformat() not in dates_covered:
            missing.append(datetime(curr.year, curr.month, curr.day))
        curr += timedelta(days=1)
    return missing


# ══════════════════════════════════════════════════════════════════════════════
# HELPERS — NOME FINAL DO ARQUIVO
# ══════════════════════════════════════════════════════════════════════════════
def gerar_nome_final(assunto_ref: str, dt_ref: datetime, url: str) -> str:
    """Gera nome canônico: <prefixo>-DD-MM-YYYY.<ext>"""
    ext = inferir_ext(url)
    prefixo = PREFIXOS_FORCADOS.get(assunto_ref, prefixo_do_url(url))
    return f"{prefixo}-{dt_ref.strftime('%d-%m-%Y')}{ext}"


# ══════════════════════════════════════════════════════════════════════════════
# STATE FILE — Leitura e escrita atômica
# ══════════════════════════════════════════════════════════════════════════════
def carregar_state() -> dict[str, Any]:
    """Carrega o JSON de deduplicação. Retorna {} em caso de arquivo ausente ou corrompido."""
    if ARQ_LINKS_JSON.exists():
        try:
            return json.loads(ARQ_LINKS_JSON.read_text(encoding="utf-8"))
        except Exception as e:
            logger.error("Falha ao ler links_baixados.JSON: %s — iniciando com estado vazio.", e)
    return {}


def salvar_state(state: dict[str, Any]) -> None:
    """Persiste o estado de deduplicação com escrita atômica (via arquivo .tmp).

    CRITICAL: A escrita atômica via arquivo temporário + os.replace() garante que,
    mesmo que o processo seja interrompido (queda de energia, kill), o JSON nunca
    fique em estado inválido ou vazio. NÃO substitua por write_text() direto.
    """
    tmp = ARQ_LINKS_JSON.with_suffix(".tmp")
    try:
        tmp.write_text(json.dumps(state, indent=2, ensure_ascii=False), encoding="utf-8")
        os.replace(tmp, ARQ_LINKS_JSON)
    except Exception as e:
        logger.error("Falha ao salvar state: %s", e)
        if tmp.exists():
            tmp.unlink(missing_ok=True)


# ══════════════════════════════════════════════════════════════════════════════
# LOCKFILE — Prevenção de execuções concorrentes
# ══════════════════════════════════════════════════════════════════════════════
def adquirir_lock() -> bool:
    """Cria o arquivo .lock. Retorna False se já existe (outra instância em execução)."""
    if ARQ_LOCK.exists():
        try:
            conteudo = ARQ_LOCK.read_text(encoding="utf-8")
            logger.warning(
                "⚠️  Lock ativo detectado (%s). Outra instância pode estar rodando. "
                "Se tiver certeza que não há outra instância, delete manualmente: %s",
                conteudo.strip(), ARQ_LOCK
            )
        except Exception:
            pass
        return False
    try:
        ARQ_LOCK.write_text(f"pid={os.getpid()} started={datetime.now().isoformat()}", encoding="utf-8")
        return True
    except Exception as e:
        logger.error("Não foi possível criar lockfile: %s", e)
        return False


def liberar_lock() -> None:
    """Remove o arquivo .lock ao final da execução."""
    try:
        ARQ_LOCK.unlink(missing_ok=True)
    except Exception as e:
        logger.warning("Não foi possível remover lockfile: %s", e)


# ══════════════════════════════════════════════════════════════════════════════
# LIMPEZA DE ARQUIVOS ÓRFÃOS
# ══════════════════════════════════════════════════════════════════════════════
def limpar_arquivos_part() -> None:
    """Remove arquivos .part e .tmp órfãos de execuções anteriores interrompidas."""
    pastas_destino = {cfg["destino_pasta"] for cfg in ASSUNTOS_PASTAS.values()}
    total_removidos = 0
    for pasta in pastas_destino:
        if not pasta.exists():
            continue
        for f in pasta.rglob("*.part"):
            try:
                f.unlink()
                total_removidos += 1
                logger.debug("Removido órfão: %s", f.name)
            except Exception as e:
                logger.warning("Não foi possível remover %s: %s", f, e)
        for f in pasta.rglob(".recovery_temp_*"):
            try:
                f.unlink()
                total_removidos += 1
            except Exception as e:
                logger.warning("Não foi possível remover %s: %s", f, e)
    if total_removidos:
        logger.info("Limpeza: %d arquivo(s) órfão(s) removido(s).", total_removidos)


# ══════════════════════════════════════════════════════════════════════════════
# EXTRAÇÃO DE LINKS
# ══════════════════════════════════════════════════════════════════════════════
def extrair_links(html: str) -> list[str]:
    """Extrai URLs S3 do corpo HTML do e-mail.

    CRITICAL: Usa DUAS estratégias complementares:
      1. BeautifulSoup extrai hrefs de tags <a> — já decodifica &amp; → &
      2. S3_REGEX no HTML cru — captura links fora de <a> tags
         O html.unescape() é OBRIGATÓRIO aqui pois o HTML cru contém &amp;
         nos query params das URLs pré-assinadas. Sem ele, a assinatura AWS
         falha (HTTP 403). NÃO remova o unescape().
    """
    from bs4 import BeautifulSoup
    soup = BeautifulSoup(html or "", "html.parser")
    links = [a["href"] for a in soup.find_all("a", href=True)]
    # html.unescape converte &amp; → & nos query params das URLs pré-assinadas S3
    links.extend(unescape(url) for url in S3_REGEX.findall(html or ""))
    
    links = list(set(url.strip() for url in links if isinstance(url, str) and url.strip()))
    
    # Remove links sem query params quando a versão completa já foi capturada
    valid_links = []
    for u in links:
        if not any(other != u and other.startswith(u) for other in links):
            valid_links.append(u)
    return valid_links


# ══════════════════════════════════════════════════════════════════════════════
# DOWNLOAD
# ══════════════════════════════════════════════════════════════════════════════
def _is_presigned_url_expired(url: str) -> bool:
    """Verifica se uma URL pré-assinada S3 está expirada com base nos parâmetros X-Amz-*."""
    from urllib.parse import parse_qs
    qs = parse_qs(urlparse(url).query)
    amz_date = qs.get("X-Amz-Date", [None])[0]
    amz_expires = qs.get("X-Amz-Expires", [None])[0]
    if not amz_date or not amz_expires:
        return False
    try:
        criacao = datetime.strptime(amz_date, "%Y%m%dT%H%M%SZ")
        expiracao = criacao + timedelta(seconds=int(amz_expires))
        return datetime.utcnow() > expiracao
    except Exception:
        return False


def baixar_arquivo(url: str, destino: Path, nome_arquivo: str) -> tuple[bool, int | None, str | None]:
    """Baixa um arquivo com retry automático e arquivo temporário para atomicidade.

    Retorna: (sucesso, tamanho_bytes, mensagem_erro)
    """
    destino.mkdir(parents=True, exist_ok=True)
    out_path = destino / nome_arquivo
    temp_path = destino / f".{nome_arquivo}.part"
    err = None

    for attempt in range(1, REQUEST_RETRIES + 1):
        try:
            with requests.get(url, timeout=REQUEST_TIMEOUT, headers=REQUEST_HEADERS, stream=True) as r:
                # Detecta URL pré-assinada expirada antes de propagar erro genérico
                if r.status_code in (400, 403) and _is_presigned_url_expired(url):
                    return False, None, "URL pré-assinada expirada (X-Amz-Expires ultrapassado)"
                r.raise_for_status()
                with temp_path.open("wb") as f:
                    size = 0
                    for chunk in r.iter_content(chunk_size=8192):
                        if chunk:
                            f.write(chunk)
                            size += len(chunk)
                if size == 0:
                    raise ValueError("Resposta vazia: 0 bytes recebidos")
            if out_path.exists():
                out_path.unlink()
            temp_path.rename(out_path)
            return True, size, None
        except Exception as e:
            err = str(e)
            if temp_path.exists():
                temp_path.unlink(missing_ok=True)
            if attempt < REQUEST_RETRIES:
                logger.warning("  Download tentativa %d/%d falhou: %s. Aguardando %ds...",
                               attempt, REQUEST_RETRIES, err, REQUEST_SLEEP_BETWEEN_RETRIES)
                sleep(REQUEST_SLEEP_BETWEEN_RETRIES)

    return False, None, err


# ══════════════════════════════════════════════════════════════════════════════
# PROCESSAMENTO DO OUTLOOK
# ══════════════════════════════════════════════════════════════════════════════
def processar_outlook(dt_ini: datetime) -> int:
    """Processo principal: conecta ao Outlook, extrai links e baixa arquivos.

    CRITICAL: A ordem de processamento dos assuntos segue ASSUNTOS_PASTAS (config.py).
    Não reordenar sem entender o impacto na deduplicação cruzada com o histórico.
    """
    logger.info("Conectando ao Outlook para conta: %s", NOME_CONTA)
    try:
        ns = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
        conta = next(
            (f for f in ns.Folders if normalize_text(f.Name) == normalize_text(NOME_CONTA)),
            None
        )
        if not conta:
            raise Exception(f"Conta '{NOME_CONTA}' não encontrada no Outlook.")

        # Busca robusta pela Caixa de Entrada (varia conforme idioma do Outlook)
        inbox = None
        for f in conta.Folders:
            if normalize_text(f.Name) in ["caixa de entrada", "inbox"]:
                inbox = f
                break
        raiz_busca = inbox if inbox else conta
        logger.info("Raiz de busca: '%s'", raiz_busca.Name)

    except Exception as e:
        logger.error("Erro ao acessar Outlook: %s", e)
        return 0

    state = carregar_state()

    inventory: dict[str, Any] = {}
    if ARQ_INVENTARIO_JSON.exists():
        try:
            inventory = json.loads(ARQ_INVENTARIO_JSON.read_text(encoding="utf-8"))
        except Exception as e:
            logger.warning("Falha ao ler inventory_dates.JSON: %s", e)

    hoje = datetime.now()
    primeiro_dia_mes_atual = hoje.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
    dt_limite_rigoroso = (primeiro_dia_mes_atual - timedelta(days=1)).replace(day=1)

    logger.info("Janela de datas: de %s em diante.", dt_limite_rigoroso.strftime("%d/%m/%Y"))

    dt_ini_consolidado = dt_limite_rigoroso
    dt_corte_geral = dt_limite_rigoroso
    baixados_total = 0

    for ass_ref, cfg in ASSUNTOS_PASTAS.items():
        logger.info("Processando assunto: %s", ass_ref)
        try:
            # Navegar até a pasta do Outlook
            folder = raiz_busca
            for sub_nome in cfg["outlook_pasta"].split("\\"):
                target_norm = normalize_text(sub_nome)
                next_folder = None
                for f in folder.Folders:
                    if normalize_text(f.Name) == target_norm:
                        next_folder = f
                        break
                if not next_folder:
                    logger.warning(
                        "Pasta '%s' não encontrada em '%s'. Usando '%s' como fallback.",
                        sub_nome, folder.Name, raiz_busca.Name
                    )
                    folder = raiz_busca
                    break
                folder = next_folder

            is_consolidado = "consolidado" in ass_ref.lower()
            dt_limite = dt_ini_consolidado if is_consolidado else dt_corte_geral
            filtro = f"[ReceivedTime] >= '{dt_limite.strftime('%d/%m/%Y %I:%M %p')}'"
            items = folder.Items.Restrict(filtro)
            items.Sort("[ReceivedTime]", True)

            chave_inv = obter_chave_inventario(ass_ref)
            history_dates = set(inventory.get(chave_inv, {}).get("history_dates", []))

            is_retroativo = "retroativo" in ass_ref.lower()
            retroativo_count = 0

            for m in items:
                try:
                    m_dt = datetime(
                        m.ReceivedTime.year, m.ReceivedTime.month, m.ReceivedTime.day,
                        m.ReceivedTime.hour, m.ReceivedTime.minute
                    )
                except Exception:
                    continue

                if m_dt < dt_limite:
                    continue
                if normalize_text(ass_ref) not in normalize_text(m.Subject):
                    continue

                if is_retroativo and retroativo_count >= 1:
                    logger.info("  Ignorando e-mail de historico mais antigo: %s (%s)", m.Subject, m_dt.strftime("%d/%m/%Y"))
                    continue

                links = extrair_links(m.HTMLBody)
                if links and is_retroativo:
                    retroativo_count += 1

                for url in links:
                    # ── Lógica de deduplicação e auto-recuperação ──────────────
                    if url in state:
                        link_date_str = (
                            state[url].get("link_date")
                            or (state[url].get("at", "")[:10] if state[url].get("at") else None)
                        )
                        if link_date_str and link_date_str[:10] in history_dates:
                            continue

                        nome_final_estimado = state[url].get("file")
                        if nome_final_estimado:
                            dt_estimada = extract_date_from_filename(nome_final_estimado)
                            if dt_estimada and dt_estimada.strftime("%Y-%m-%d") in history_dates:
                                continue
                            caminho_local = cfg["destino_pasta"] / nome_final_estimado
                            if caminho_local.exists():
                                continue
                            logger.info("  Link no JSON mas arquivo AUSENTE. Re-baixando: %s", nome_final_estimado)
                        else:
                            if not is_consolidado:
                                continue

                    logger.info("  Encontrado link: %s...", url[:80])

                    ext = inferir_ext(url)
                    temp_nome = f".temp_{int(datetime.now().timestamp())}{ext}"

                    ok, sz, err = baixar_arquivo(url, cfg["destino_pasta"], temp_nome)
                    if not ok:
                        logger.error("  Falha no download: %s", err)
                        continue

                    temp_path = cfg["destino_pasta"] / temp_nome
                    usar_data_email = any(
                        x in ass_ref.lower()
                        for x in ["profissionais", "cliente", "lista", "cadastro", "posicao", "despesa", "agendamento"]
                    )

                    dt_real_str = None
                    if not usar_data_email:
                        try:
                            dt_real_str = extrair_data_do_arquivo(str(temp_path))
                        except Exception as e:
                            logger.warning("  Erro ao extrair data do arquivo: %s", e)

                    if dt_real_str:
                        dt_obj = datetime.strptime(dt_real_str, "%Y-%m-%d")
                    else:
                        rt = m.ReceivedTime
                        dt_email = datetime(rt.year, rt.month, rt.day, rt.hour, rt.minute, rt.second)
                        dt_obj = calcular_data_referencia(ass_ref, dt_email)

                    if dt_obj.strftime("%Y-%m-%d") in history_dates:
                        logger.info("  Ignorado (data %s já no histórico consolidado)", dt_obj.strftime("%Y-%m-%d"))
                        temp_path.unlink(missing_ok=True)
                        continue

                    if dt_obj < dt_limite:
                        logger.info("  Ignorado (data %s anterior ao limite %s)", dt_obj.date(), dt_limite.date())
                        temp_path.unlink(missing_ok=True)
                        continue

                    nome_final = gerar_nome_final(ass_ref, dt_obj, url)
                    final_path = cfg["destino_pasta"] / nome_final

                    if final_path.exists():
                        temp_path.unlink(missing_ok=True)
                        logger.info("  Arquivo já existe: %s", nome_final)
                    else:
                        temp_path.rename(final_path)
                        logger.info("  [OK] Baixado: %s (%d KB)", nome_final, (sz or 0) // 1024)

                    state[url] = {
                        "status": "downloaded",
                        "at": datetime.now().isoformat(),
                        "file": nome_final,
                        "subject": m.Subject,
                    }
                    baixados_total += 1

        except Exception as e:
            logger.error("Erro no processamento de '%s': %s", ass_ref, e)

    # ── Fase de Recuperação ──────────────────────────────────────────────────
    try:
        logger.info("==== Fase de Recuperação: e-mails com links S3 reenviados ====")
        recovery_links = varrer_inbox_recovery(raiz_busca, dt_ini, state)
        if recovery_links:
            n_rec = processar_recovery(recovery_links, inventory, state, hoje)
            baixados_total += n_rec
            logger.info("[RECOVERY] ✔ %d arquivo(s) recuperado(s).", n_rec)
        else:
            logger.info("[RECOVERY] Nenhum link de recuperação encontrado.")
    except Exception as exc:
        logger.exception("[RECOVERY] Erro na varredura de recuperação: %s", exc)

    salvar_state(state)
    return baixados_total


# ══════════════════════════════════════════════════════════════════════════════
# RECOVERY — Varredura de e-mails de reenvio manual
# ══════════════════════════════════════════════════════════════════════════════
def varrer_inbox_recovery(
    inbox, dt_ini: datetime, state: dict
) -> dict[str, list[tuple[str, datetime]]]:
    """Varre a Caixa de Entrada por e-mails de recuperação com links S3 de relatórios.

    Ignora e-mails cujo assunto já consta em ASSUNTOS_PASTAS (esses são processados
    no fluxo principal). Captura apenas e-mails manuais (ex: reenvios da Trinks).
    """
    recovery: dict[str, list[tuple[str, datetime]]] = {}
    assuntos_norm = [normalize_text(a) for a in ASSUNTOS_PASTAS]
    filtro = f"[ReceivedTime] >= '{dt_ini.strftime('%d/%m/%Y %I:%M %p')}'"

    try:
        items = inbox.Items.Restrict(filtro)
        items.Sort("[ReceivedTime]", True)
        logger.info("[RECOVERY] Varrendo %d e-mail(s)...", items.Count)
    except Exception as exc:
        logger.warning("[RECOVERY] Falha ao acessar inbox: %s", exc)
        return recovery

    for m in items:
        subj_norm = normalize_text(getattr(m, "Subject", "") or "")
        if any(known in subj_norm for known in assuntos_norm):
            continue

        html_body = getattr(m, "HTMLBody", "") or ""
        body_text = getattr(m, "Body", "") or ""
        links = extrair_links(html_body) or [unescape(u) for u in S3_REGEX.findall(body_text)]

        if not links:
            continue

        received_time = getattr(m, "ReceivedTime", None)
        if not received_time:
            continue

        rt_email = datetime(
            received_time.year, received_time.month, received_time.day,
            received_time.hour, received_time.minute, received_time.second
        )

        for url in links:
            url = url.strip()
            if not S3_REGEX.search(url) or url in state:
                continue
            prefixo = prefixo_do_url(url)
            ass_ref = PREFIXO_TO_ASSUNTO.get(prefixo)
            if not ass_ref:
                logger.warning("[RECOVERY] Prefixo '%s' sem mapeamento: %s", prefixo, url[:80])
                continue
            recovery.setdefault(ass_ref, []).append((url, rt_email))
            logger.info("[RECOVERY] Link detectado: %s → %s", prefixo, arquivo(ass_ref))

    return recovery


def processar_recovery(
    recovery: dict[str, list[tuple[str, datetime]]],
    inventory: dict,
    state: dict,
    inicio_run: datetime,
) -> int:
    """Baixa os links de recuperação e salva com nome canônico."""
    total = 0
    for ass_ref, url_tuples in recovery.items():
        cfg = ASSUNTOS_PASTAS.get(ass_ref)
        if not cfg:
            logger.warning("[RECOVERY] Assunto '%s' não encontrado em ASSUNTOS_PASTAS.", ass_ref)
            continue

        destino = cfg["destino_pasta"]
        destino.mkdir(parents=True, exist_ok=True)

        chave_inv = obter_chave_inventario(ass_ref)
        pending_dates = sorted(obter_datas_pendentes(chave_inv, inventory, inicio_run))

        for i, (url, received_time) in enumerate(url_tuples):
            ext_real = inferir_ext(url)
            temp_nome = f".recovery_temp_{i}_{int(received_time.timestamp())}{ext_real}"
            ok_temp, sz, er_temp = baixar_arquivo(url, destino, temp_nome)

            if not ok_temp:
                logger.error("  [RECOVERY ERRO] %s — %s", url[:80], er_temp)
                continue

            temp_path = destino / temp_nome
            dt_ref = None

            usar_data_email = any(
                x in ass_ref.lower()
                for x in ["profissionais", "cliente", "lista", "cadastro", "posicao", "despesa", "agendamento"]
            )
            if not usar_data_email:
                try:
                    data_str = extrair_data_do_arquivo(str(temp_path))
                    if data_str:
                        dt_ref = datetime.strptime(data_str, "%Y-%m-%d")
                except Exception as e:
                    logger.warning("  [RECOVERY] Erro ao extrair data: %s", e)

            if dt_ref is None:
                if i < len(pending_dates):
                    dt_ref = pending_dates[i]
                else:
                    freq = DISPLAY_MAP.get(chave_inv, {}).get("freq", "Diária") if chave_inv else "Diária"
                    days_back = i if freq == "Diária" else 0
                    dt_ref = calcular_data_referencia(ass_ref, received_time) - timedelta(days=days_back)

            if not data_no_corte_ou_depois(dt_ref):
                logger.info("  [RECOVERY SKIP] Data %s anterior ao corte.", dt_ref.date())
                state[url] = {"status": "ignored_before_cutoff", "link_date": dt_ref.isoformat(),
                              "assunto": ass_ref, "updated_at": datetime.now().isoformat()}
                temp_path.unlink(missing_ok=True)
                continue

            nome = gerar_nome_final(ass_ref, dt_ref, url)
            final_path = destino / nome
            try:
                if final_path.exists():
                    final_path.unlink()
                temp_path.rename(final_path)
                total += 1
                state[url] = {
                    "status": "recovered", "file_name": nome, "link_date": dt_ref.isoformat(),
                    "assunto": ass_ref, "file_size": sz, "at": datetime.now().isoformat(),
                }
                logger.info("  [RECOVERY OK] %s", nome)
            except Exception as e:
                logger.error("  [RECOVERY ERRO] Falha ao mover para %s: %s", nome, e)
                temp_path.unlink(missing_ok=True)

    return total


def verificar_se_ha_atrasos(hoje: datetime) -> bool:
    """Lê inventory_dates.JSON e calcula se há atrasos conforme as regras de SLA."""
    if not ARQ_INVENTARIO_JSON.exists():
        return True
    try:
        inventory = json.loads(ARQ_INVENTARIO_JSON.read_text(encoding="utf-8"))
    except Exception:
        return True
    
    hoje_date = hoje.date()
    
    for key, conf in DISPLAY_MAP.items():
        inv = inventory.get(key, {"dates": []})
        dates_covered = set(inv.get("dates", []))
        
        if conf["freq"] == "Diária":
            atraso_permitido = 2 if ("Consolidado" in key or "Pagamento" in key) else 1
            data_max_item = hoje_date - timedelta(days=atraso_permitido)
            
            curr = DATA_CORTE_DOWNLOAD_DATE
            while curr <= data_max_item:
                if curr.isoformat() not in dates_covered:
                    return True
                curr += timedelta(days=1)
        else:
            all_dates = inv.get("dates", [])
            if all_dates:
                ultima_data = max(all_dates)
                try:
                    atraso_dias = (hoje_date - date.fromisoformat(ultima_data)).days
                    if atraso_dias > conf["rag"] * 1.5:
                        return True
                except Exception:
                    pass
            else:
                return True
            
    return False


# ══════════════════════════════════════════════════════════════════════════════
# ENTRY POINT
# ══════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Coletor Outlook RPA")
    parser.add_argument("--days", type=int, default=15, help="Janela de lookback em dias (padrão: 15)")
    args = parser.parse_args()

    # Adquirir lock — aborta se outra instância estiver rodando
    if not adquirir_lock():
        logger.error("Execução abortada: lockfile existente. Veja %s", ARQ_LOCK)
        sys.exit(1)

    try:
        # Limpeza preventiva de arquivos .part órfãos
        limpar_arquivos_part()

        dt_ini = datetime.now() - timedelta(days=args.days)
        baixados = processar_outlook(dt_ini)
        logger.info("Processamento concluído. Total baixado nesta execução: %d", baixados)

        script_dir = Path(__file__).parent
        logger.info("Acionando Data Manager...")
        subprocess.run([sys.executable, str(script_dir / "data_manager.py")], check=False)

        # Se houver atrasos/gaps de dados, aciona o processa_manuais.py para verificar temp_manuais
        if verificar_se_ha_atrasos(datetime.now()):
            logger.info("Atrasos/gaps de dados detectados no inventário. Acionando Processa Manuais...")
            subprocess.run([sys.executable, str(script_dir / "processa_manuais.py")], check=False)
            logger.info("Re-executando Data Manager para consolidar importações manuais...")
            subprocess.run([sys.executable, str(script_dir / "data_manager.py")], check=False)
        else:
            logger.info("Nenhum atraso/gap detectado no inventário. Pulando processamento manual.")

        logger.info("Acionando Reporter...")
        subprocess.run([sys.executable, str(script_dir / "reporter.py")], check=False)

    finally:
        liberar_lock()
