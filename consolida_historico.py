# -*- coding: utf-8 -*-
"""
consolida_historico.py — Consolidação de Arquivos Manuais no Histórico
=====================================================================
Este módulo lê arquivos CSV na pasta temp_manuais, identifica o tipo de relatório,
agrupa as datas por mês, e consolida os dados de meses fechados no Excel de histórico
correspondente. Linhas antigas de datas sobrepostas são removidas para evitar duplicados.

Arquivos diários da pasta Automação para o mês fechado consolidado são limpos/arquivados.
"""

import json
import logging
import os
import shutil
import tempfile
from datetime import datetime, date
from pathlib import Path
from typing import Any, Optional

import pandas as pd

from config import (
    BASE_ALLPE,
    BASE_HISTORICO,
    MAPA_HISTORICO,
    MAPA_PASTAS,
    PREFIXOS_FORCADOS,
)
from detector_datas import extrair_todas_as_datas, extrair_data_do_valor

logger = logging.getLogger("consolida_historico")
if not logger.handlers:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
    )

# Diretório de históricos de entrada (CSV) - obtido a partir da base configurada
TEMP_MANUAIS_DIR = BASE_ALLPE / "Historico"

# Mapeamento de Assuntos/Chaves para Prefixos
KEY_TO_PREFIX = {
    'Consolidado de servicos': 'all_pe_consolidado_servicos',
    'Consolidado de produtos': 'all_pe_consolidado_produtos',
    'Consolidado de pacotes': 'all_pe_consolidado_pacotes',
    'Formas de Pagamento': 'all_pe_consolidados_forma_pagamento',
    'Agendamentos': 'all_pe_relatorio_agendamento',
    'Despesas': 'all_pe_despesa_diaria',
    'Estoque Diario': 'all_pe_posicao_estoque_diario'
}

# Configuração de Colunas de Data por categoria de histórico
CONFIG_DATAS_HISTORICO = {
    'Consolidado de servicos': {'coluna': 'DataHora do Fechamento de Contas', 'tipo': 'timestamp'},
    'Consolidado de produtos': {'coluna': 'DataHora do Fechamento de Contas', 'tipo': 'timestamp'},
    'Consolidado de pacotes':  {'coluna': 'DataHora do Fechamento de Contas', 'tipo': 'timestamp'},
    'Formas de Pagamento':     {'coluna': 'DataHora do Fechamento de Contas', 'tipo': 'timestamp'},
    'Agendamentos':            {'coluna': 'Data', 'tipo': 'str_date_br'},
    'Despesas':                {'coluna': 'Data de Vencimento', 'tipo': 'str_datetime_br'},
}

def detectar_categoria_por_colunas(df: pd.DataFrame) -> Optional[str]:
    """Identifica a categoria do relatório com base nas colunas do DataFrame."""
    cols = [str(c).lower().strip() for c in df.columns]
    
    if 'id do agendamento' in cols or 'profissional da vez' in cols:
        return 'Agendamentos'
    if 'codigo despesa' in cols or 'data de vencimento' in cols:
        return 'Despesas'
    if 'tipo forma pagamento' in cols:
        return 'Formas de Pagamento'
        
    if 'id do horario' in cols or 'id profissional' in cols or 'id do profissional' in cols or 'id_profissional' in cols:
        return 'Consolidado de servicos'
        
    if 'item do fechamento de contas' in cols:
        # Encontra o nome original (com as maiúsculas/minúsculas corretas)
        orig_col = [c for c in df.columns if str(c).lower().strip() == 'item do fechamento de contas'][0]
        # Verifica valores na coluna para discernir produtos de pacotes
        sample_vals = df[orig_col].dropna().astype(str).str.lower().str.strip().unique()
        if any('produto' in val for val in sample_vals):
            return 'Consolidado de produtos'
        if any('pacote' in val for val in sample_vals):
            return 'Consolidado de pacotes'
        
        # Fallback de busca iterativa
        for val in df[orig_col].dropna():
            val_str = str(val).lower()
            if 'produto' in val_str:
                return 'Consolidado de produtos'
            if 'pacote' in val_str:
                return 'Consolidado de pacotes'
                
    return None

def obter_excel_historico_ativo(key: str) -> Optional[Path]:
    """Retorna o caminho do arquivo Excel histórico ativo (último da lista no config)."""
    arquivos = MAPA_HISTORICO.get(key)
    if not arquivos:
        return None
    return BASE_HISTORICO / arquivos[-1]

def normalizar_coluna_data(df: pd.DataFrame, col_name: str, target_type: str) -> pd.Series:
    """Normaliza a coluna de data para o tipo e formato esperado pelo Excel histórico."""
    # Extrai datas cruas normalizadas para YYYY-MM-DD
    parsed_dates = df[col_name].apply(extrair_data_do_valor)
    dt_series = pd.to_datetime(parsed_dates, errors='coerce')
    
    if target_type == 'timestamp':
        return dt_series
    elif target_type == 'str_date_br':
        return dt_series.dt.strftime('%d/%m/%Y')
    elif target_type == 'str_datetime_br':
        # Tenta preservar a hora original se houver
        dt_full = pd.to_datetime(df[col_name], errors='coerce')
        # Se falhar, usa a data básica e formata com 0:00
        mask_coerce = dt_full.isna() & dt_series.notna()
        dt_full[mask_coerce] = dt_series[mask_coerce]
        return dt_full.dt.strftime('%d/%m/%Y %H:%M')
    return df[col_name]

def consolidar_mes_no_excel(key: str, df_mes: pd.DataFrame, dates_to_import: set[str]) -> bool:
    """
    Carrega o arquivo histórico ativo, remove dados antigos das datas importadas
    para evitar duplicados e anexa os novos registros de forma atômica.
    """
    excel_path = obter_excel_historico_ativo(key)
    if not excel_path:
        logger.error("Nenhum arquivo histórico configurado para %s", key)
        return False
        
    cfg_data = CONFIG_DATAS_HISTORICO.get(key)
    if not cfg_data:
        logger.error("Configuração de data de histórico ausente para %s", key)
        return False
        
    col_data = cfg_data['coluna']
    tipo_data = cfg_data['tipo']
    
    logger.info("Consolidando em %s (Coluna: %s)", excel_path.name, col_data)
    
    # 1. Carregar Excel atual
    if excel_path.exists():
        try:
            df_excel = pd.read_excel(excel_path)
            logger.info("  Arquivo atual carregado com %d linhas.", len(df_excel))
        except Exception as e:
            logger.error("  Falha ao ler Excel %s: %s", excel_path.name, e)
            return False
    else:
        logger.info("  Arquivo histórico %s não existe. Criando novo.", excel_path.name)
        df_excel = pd.DataFrame(columns=df_mes.columns)
        
    # 2. Filtrar dados antigos correspondentes às datas importadas para evitar duplicações
    if not df_excel.empty and col_data in df_excel.columns:
        excel_dates_str = df_excel[col_data].apply(extrair_data_do_valor)
        mask_keep = ~excel_dates_str.isin(dates_to_import)
        linhas_antes = len(df_excel)
        df_excel = df_excel[mask_keep]
        linhas_removidas = linhas_antes - len(df_excel)
        if linhas_removidas > 0:
            logger.info("  Removidas %d linhas antigas sobrepostas no histórico.", linhas_removidas)
            
    # 3. Normalizar colunas do CSV para coincidir com a estrutura do Excel
    df_mes_aligned = df_mes.copy()
    if col_data in df_mes_aligned.columns:
        df_mes_aligned[col_data] = normalizar_coluna_data(df_mes_aligned, col_data, tipo_data)
        
    # Reindexar colunas para garantir alinhamento exato
    if not df_excel.empty:
        df_mes_aligned = df_mes_aligned.reindex(columns=df_excel.columns)
        
    # 4. Concatenar
    df_final = pd.concat([df_excel, df_mes_aligned], ignore_index=True)
    
    # Remoção estrita de duplicados no dataframe final
    df_final = df_final.drop_duplicates()
    
    logger.info("  Total de linhas após mesclagem (com deduplicação): %d", len(df_final))
    
    # 5. Escrita atômica no Excel
    excel_path.parent.mkdir(parents=True, exist_ok=True)
    tmp_path = excel_path.with_suffix(".xlsx.tmp")
    try:
        df_final.to_excel(tmp_path, index=False)
        os.replace(tmp_path, excel_path)
        logger.info("  [SUCESSO] Arquivo de histórico salvo com segurança.")
        return True
    except Exception as e:
        logger.error("  [ERRO] Falha ao salvar histórico %s: %s", excel_path.name, e)
        if tmp_path.exists():
            tmp_path.unlink()
        return False

def apagar_diarios_por_datas(key: str, datas_importadas: set[str]) -> None:
    """
    Apaga os arquivos diários correspondentes às datas importadas na pasta Automação
    para evitar que o Power BI importe dados duplicados.
    """
    pasta_d = MAPA_PASTAS.get(key)
    if not pasta_d or not pasta_d.exists():
        logger.warning("  Diretorio diario nao existe para chave %s: %s", key, pasta_d)
        return
        
    prefixo = KEY_TO_PREFIX.get(key)
    if not prefixo:
        logger.warning("  Prefixo de arquivo nao encontrado para chave %s", key)
        return
        
    logger.info("  Iniciando a remocao de arquivos diarios para as datas consolidadas de %s...", key)
    
    arquivos_deletados_count = 0
    
    for dt_iso in sorted(list(datas_importadas)):
        try:
            dt_obj = datetime.strptime(dt_iso, "%Y-%m-%d")
            dt_br = dt_obj.strftime("%d-%m-%Y")
        except Exception as e:
            logger.error("  Data ISO invalida %s: %s", dt_iso, e)
            continue
            
        # Padrão de busca: *-DD-MM-YYYY.csv
        padrao = f"*-{dt_br}.csv"
        arquivos_encontrados = list(pasta_d.glob(padrao))
        
        for f in arquivos_encontrados:
            try:
                f.unlink()
                logger.info("    [APAGADO] Arquivo diario deletado: %s", f.name)
                arquivos_deletados_count += 1
            except PermissionError:
                logger.warning("    [AVISO] Arquivo bloqueado pelo Windows: %s. Feche o Excel e tente novamente.", f.name)
            except Exception as e:
                logger.error("    [ERRO] Falha ao deletar arquivo %s: %s", f.name, e)
                
    if arquivos_deletados_count > 0:
        logger.info("  Removidos %d arquivos diarios da pasta %s.", arquivos_deletados_count, pasta_d.name)
    else:
        logger.info("  Nenhum arquivo diario correspondente foi encontrado para exclusao em %s.", pasta_d.name)

def registrar_no_estado(nome_arq_original: str, nome_arq_final: str) -> None:
    """Registra o arquivo manual processado no links_baixados.JSON para manter a rastreabilidade."""
    try:
        from downloader import carregar_state, salvar_state
        state = carregar_state()
        fake_url = f"manual://drive/{nome_arq_original}"
        state[fake_url] = {
            "status": "downloaded",
            "at": datetime.now().isoformat(),
            "file": nome_arq_final,
            "subject": "Enc.: Re: Erro Links - Extrações (Google Drive Manual)"
        }
        salvar_state(state)
        logger.info("  Registrado no links_baixados.JSON: %s", fake_url)
    except Exception as e:
        logger.error("  Falha ao registrar no state JSON: %s", e)

def processar_arquivo_consolidacao(caminho_csv: Path) -> bool:
    """Lê um arquivo manual do temp_manuais, detecta tipo, divide por mês e consolida."""
    logger.info("Analisando arquivo manual para histórico: %s", caminho_csv.name)
    
    # 1. Carregar CSV e detectar separador (com fallback de encoding)
    try:
        # Tenta ler primeiras linhas para detectar delimitador
        with open(caminho_csv, 'r', encoding='utf-8', errors='ignore') as f:
            sample = f.read(4096)
        sep = ';' if ';' in sample else ','

        # Tenta encodings em ordem: utf-8 -> cp1252 -> latin1
        df = None
        encoding_usado = None
        for enc in ('utf-8', 'cp1252', 'latin1'):
            try:
                df = pd.read_csv(caminho_csv, sep=sep, encoding=enc, on_bad_lines='skip')
                encoding_usado = enc
                break
            except UnicodeDecodeError:
                continue
        if df is None:
            raise ValueError(f"Nenhum encoding funcionou para {caminho_csv.name}")
        logger.info("  CSV carregado com %d linhas e %d colunas (encoding: %s).", len(df), len(df.columns), encoding_usado)
    except Exception as e:
        logger.error("  Falha ao ler arquivo %s: %s", caminho_csv.name, e)
        return False
        
    # 2. Detectar categoria do relatório
    key = detectar_categoria_por_colunas(df)
    if not key:
        logger.warning("  Não foi possível identificar a categoria do relatório pelas colunas do arquivo %s.", caminho_csv.name)
        return False
        
    logger.info("  Categoria detectada: %s", key)
    
    cfg_data = CONFIG_DATAS_HISTORICO.get(key)
    if not cfg_data:
        logger.warning("  Categoria %s não possui histórico configurado para escrita. Pulando.", key)
        return False
        
    col_data = cfg_data['coluna']
    if col_data not in df.columns:
        logger.error("  Coluna de data '%s' não localizada no arquivo.", col_data)
        return False
        
    # 3. Mapear datas e separar linhas por ano-mês
    # Adiciona coluna temporária de data ISO YYYY-MM-DD
    df_temp = df.copy()
    df_temp['_date_iso'] = df_temp[col_data].apply(extrair_data_do_valor)
    
    # Filtrar linhas sem data válida
    df_temp = df_temp.dropna(subset=['_date_iso'])
    if df_temp.empty:
        logger.error("  Nenhuma data válida encontrada na coluna '%s'.", col_data)
        return False
        
    # Agrupar por ano-mês (YYYY-MM)
    df_temp['_year_month'] = df_temp['_date_iso'].apply(lambda x: x[:7] if isinstance(x, str) else "")
    df_temp = df_temp[df_temp['_year_month'] != ""]
    
    meses_presentes = sorted(list(df_temp['_year_month'].unique()))
    grupos = df_temp.groupby('_year_month')
    
    hoje = date.today()
    current_month_str = hoje.strftime("%Y-%m")
    
    sucesso_total = True
    
    for ano_mes, df_grupo in grupos:
        # Limpar colunas temporárias para a gravação
        df_gravar = df_grupo.drop(columns=['_date_iso', '_year_month'])
        
        # Verificar se é um mês fechado
        # Um mês é fechado se ano_mes < current_month_str (ex: "2026-05" < "2026-06")
        is_closed_month = ano_mes < current_month_str
        
        datas_grupo = set(df_grupo['_date_iso'].unique())
        
        if is_closed_month:
            logger.info("Processando MÊS FECHADO (%s) com %d registros...", ano_mes, len(df_gravar))
            # Realizar a consolidação no arquivo de histórico
            sucesso_cons = consolidar_mes_no_excel(key, df_gravar, datas_grupo)
            if sucesso_cons:
                # Registrar no estado
                excel_path = obter_excel_historico_ativo(key)
                if excel_path:
                    registrar_no_estado(caminho_csv.name, excel_path.name)
            else:
                sucesso_total = False
        else:
            logger.info("Processando MÊS CORRENTE (%s) com %d registros...", ano_mes, len(df_gravar))
            # Salvar como arquivo diário na pasta de Automação correspondente
            pasta_d = MAPA_PASTAS.get(key)
            prefixo = KEY_TO_PREFIX.get(key)
            if pasta_d and prefixo:
                # A data do nome do arquivo diário será a maior data do grupo
                max_date_iso = sorted(list(datas_grupo))[-1]
                dt_obj = datetime.strptime(max_date_iso, "%Y-%m-%d")
                nome_diario = f"{prefixo}-{dt_obj.strftime('%d-%m-%Y')}.csv"
                destino_diario = pasta_d / nome_diario
                
                logger.info("  Salvando arquivo diário na pasta Automação: %s", nome_diario)
                pasta_d.mkdir(parents=True, exist_ok=True)
                try:
                    df_gravar.to_csv(destino_diario, sep=sep, index=False, encoding='utf-8')
                    logger.info("  [SUCESSO] Arquivo diário gravado.")
                    registrar_no_estado(caminho_csv.name, nome_diario)
                except Exception as e:
                    logger.error("  [ERRO] Falha ao gravar arquivo diário %s: %s", nome_diario, e)
                    sucesso_total = False
            else:
                logger.error("  Configurações de pasta local não encontradas para %s", key)
                sucesso_total = False
                
    # Se gravou com sucesso e o histórico tiver 3 ou mais meses (ex: 05, 04, 03),
    # ele vai inserir tudo e apagar do diário os 2 meses anteriores/últimos para trás (04 e 03)
    if sucesso_total:
        if len(meses_presentes) >= 3:
            mes_mais_recente = meses_presentes[-1]
            datas_excluir = set()
            for dt_iso in df_temp['_date_iso'].unique():
                if isinstance(dt_iso, str) and dt_iso[:7] != mes_mais_recente:
                    datas_excluir.add(dt_iso)
            if datas_excluir:
                apagar_diarios_por_datas(key, datas_excluir)
        else:
            logger.info("  Apenas 1 mes de dados detectado (%s) ou menos de 3 meses. Nao apagando nada do diario.", meses_presentes)

    return sucesso_total

def run_consolida_historico() -> None:
    """Varre a pasta de históricos e processa todos os arquivos CSV."""
    if not TEMP_MANUAIS_DIR.exists():
        logger.info("Pasta de input não encontrada.")
        return
        
    # Pega apenas arquivos CSV na raiz do diretório (ignora subpastas como Backup_Processados)
    arquivos = [p for p in TEMP_MANUAIS_DIR.glob("*.csv") if p.is_file()]
    if not arquivos:
        logger.info("Nenhum arquivo CSV encontrado para consolidar em %s", TEMP_MANUAIS_DIR)
        return
        
    logger.info("Iniciando consolidação de histórico para %d arquivo(s)...", len(arquivos))
    
    backup_dir = TEMP_MANUAIS_DIR / "Backup_Processados"
    backup_dir.mkdir(exist_ok=True)
    
    for f in arquivos:
        sucesso = processar_arquivo_consolidacao(f)
        if sucesso:
            dest = backup_dir / f.name
            try:
                if dest.exists():
                    dest.unlink()
                shutil.move(str(f), str(dest))
                logger.info("Arquivo original consolidado movido para Backup_Processados: %s", f.name)
            except Exception as e:
                logger.error("Falha ao mover arquivo original %s para backup: %s", f.name, e)
        else:
            logger.warning("Falha no processamento do arquivo %s. Ele foi mantido na pasta.", f.name)

if __name__ == '__main__':
    run_consolida_historico()
