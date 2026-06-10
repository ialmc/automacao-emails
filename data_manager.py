# -*- coding: utf-8 -*-
"""
data_manager.py — Inventário de Datas All Pé
=============================================
Varre as pastas locais do OneDrive e os arquivos de histórico, construindo
o inventário de datas cobertas por cada tipo de relatório.

Saída: inventory_dates.JSON — lido pelo downloader.py e pelo reporter.py.

⛔ PARA IAs E DESENVOLVEDORES — LEIA ANTES DE MODIFICAR:
  - Paths e MAPA_PASTAS vêm de config.py. NÃO redefina aqui.
  - MAPA_HISTORICO lista os XLSXs consolidados. Se um arquivo histórico for
    renomeado, atualize config.MAPA_HISTORICO — não altere aqui.
  - A escrita do inventário é ATÔMICA (via .tmp + os.replace). NÃO substitua
    por write_text() direto — risco de corromper o arquivo em caso de crash.
"""

import json
import logging
import os
from datetime import datetime
from pathlib import Path
from typing import Any

from config import (
    ARQ_INVENTARIO_JSON,
    BASE_HISTORICO,
    MAPA_PASTAS,
    MAPA_HISTORICO,
    DAILY_KEYS,
)
from detector_datas import extrair_todas_as_datas, extract_date_from_filename

logger = logging.getLogger("data_manager")
if not logger.handlers:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
    )


def carregar_inventario() -> dict[str, Any]:
    """Carrega o inventário existente. Retorna {} se ausente ou corrompido."""
    if ARQ_INVENTARIO_JSON.exists():
        try:
            return json.loads(ARQ_INVENTARIO_JSON.read_text(encoding="utf-8"))
        except Exception as e:
            logger.error("Falha ao ler inventory_dates.JSON: %s — iniciando vazio.", e)
    return {}


def salvar_inventario(inv: dict[str, Any]) -> None:
    """Persiste o inventário com escrita atômica (via .tmp + os.replace).

    CRITICAL: A operação os.replace() é atômica no Windows/Linux. Isso garante
    que o arquivo nunca fique em estado inconsistente. NÃO substitua por
    write_text() direto — em caso de crash durante a escrita, o inventário
    seria perdido completamente.
    """
    inv["_updated_at"] = datetime.now().isoformat()
    tmp = ARQ_INVENTARIO_JSON.with_suffix(".tmp")
    try:
        tmp.write_text(json.dumps(inv, indent=2, ensure_ascii=False), encoding="utf-8")
        os.replace(tmp, ARQ_INVENTARIO_JSON)
        logger.info("Inventário salvo em: %s", ARQ_INVENTARIO_JSON)
    except Exception as e:
        logger.error("Falha ao salvar inventário: %s", e)
        tmp.unlink(missing_ok=True)


def atualizar_inventario() -> None:
    """Varre pastas locais e históricos, atualizando o inventário de datas cobertas."""
    logger.info("Iniciando atualização do inventário de datas...")
    inv = carregar_inventario()

    # CRITICAL: todos_assuntos é a união de MAPA_HISTORICO e MAPA_PASTAS.
    # Qualquer chave em DISPLAY_MAP (config.py) deve estar em pelo menos um deles.
    todos_assuntos = set(list(MAPA_HISTORICO.keys()) + list(MAPA_PASTAS.keys()))

    for assunto in todos_assuntos:
        if assunto not in inv:
            inv[assunto] = {"dates": [], "sources": {}}

        logger.info("Processando: %s", assunto)
        datas_encontradas: set[str] = set()
        datas_historico: set[str] = set()
        sources = inv[assunto].get("sources", {})
        novos_sources: dict[str, Any] = {}

        # ── 1. Arquivos de Histórico (XLSXs consolidados) ─────────────────
        for nome_arq in MAPA_HISTORICO.get(assunto, []):
            caminho = BASE_HISTORICO / nome_arq
            if not caminho.exists():
                logger.warning("  Arquivo histórico não encontrado: %s", caminho)
                continue
            try:
                mtime = os.path.getmtime(caminho)
            except OSError as e:
                logger.warning("  Erro ao obter mtime de %s: %s", caminho, e)
                continue

            chave_src = str(caminho)
            if (chave_src in sources
                    and sources[chave_src].get("mtime") == mtime
                    and "dates" in sources[chave_src]):
                # Cache válido — não relê o arquivo
                novas_datas = set(sources[chave_src]["dates"])
                datas_encontradas.update(novas_datas)
                datas_historico.update(novas_datas)
                novos_sources[chave_src] = sources[chave_src]
            else:
                logger.info("  Lendo histórico: %s", nome_arq)
                try:
                    novas_datas = extrair_todas_as_datas(str(caminho))
                except Exception as e:
                    logger.error("  Falha ao extrair datas de %s: %s", nome_arq, e)
                    novas_datas = set()
                datas_encontradas.update(novas_datas)
                datas_historico.update(novas_datas)
                novos_sources[chave_src] = {
                    "mtime": mtime,
                    "count": len(novas_datas),
                    "type": "history",
                    "dates": sorted(novas_datas),
                }

        # ── 2. Pastas Diárias (arquivos baixados pelo downloader) ──────────
        pasta_d = MAPA_PASTAS.get(assunto)
        if pasta_d and pasta_d.exists():
            for f in pasta_d.rglob("*"):
                if not f.is_file() or f.suffix.lower() not in {".csv", ".xlsx"}:
                    continue
                # Ignora arquivos .part órfãos (não deveria haver, mas por segurança)
                if f.name.startswith("."):
                    continue
                try:
                    mtime = os.path.getmtime(f)
                except OSError as e:
                    logger.warning("  Erro ao obter mtime de %s: %s", f, e)
                    continue

                chave_src = str(f)
                if chave_src in sources and sources[chave_src].get("mtime") == mtime:
                    novos_sources[chave_src] = sources[chave_src]
                    if "date" in sources[chave_src]:
                        datas_encontradas.add(sources[chave_src]["date"])
                    elif "dates" in sources[chave_src]:
                        datas_encontradas.update(sources[chave_src]["dates"])
                else:
                    # Prioriza extrair as datas reais do conteúdo do arquivo para maior integridade (apenas para relatórios diários)
                    # Para relatórios semanais/mensais (como Profissionais e Clientes), não lemos o conteúdo para evitar capturar datas internas
                    # como aniversários ou datas de cadastro. Usamos diretamente o fallback do nome do arquivo.
                    novas_datas = set()
                    if assunto in DAILY_KEYS:
                        try:
                            novas_datas = extrair_todas_as_datas(str(f))
                        except Exception as e:
                            logger.warning("  Falha ao extrair datas de %s: %s", f.name, e)

                    if novas_datas:
                        datas_encontradas.update(novas_datas)
                        novos_sources[chave_src] = {
                            "mtime": mtime,
                            "count": len(novas_datas),
                            "type": "daily_content",
                            "dates": sorted(novas_datas),
                        }
                    else:
                        # Fallback: tenta extrair do nome do arquivo se a leitura de conteúdo falhar
                        dt_nome = extract_date_from_filename(f.name)
                        if dt_nome:
                            dt_str = dt_nome.strftime("%Y-%m-%d")
                            datas_encontradas.add(dt_str)
                            novos_sources[chave_src] = {"mtime": mtime, "date": dt_str, "type": "daily"}

        inv[assunto]["dates"] = sorted(datas_encontradas)
        inv[assunto]["history_dates"] = sorted(datas_historico)
        inv[assunto]["sources"] = novos_sources
        logger.info(
            "  Total de datas cobertas: %d (Histórico: %d)",
            len(datas_encontradas), len(datas_historico)
        )

    salvar_inventario(inv)


if __name__ == "__main__":
    atualizar_inventario()
