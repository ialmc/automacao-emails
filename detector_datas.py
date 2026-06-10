"""
Detector de Data em Arquivos CSV/XLSX
Procura por coluna com 'data' no nome e extrai a data dos dados (apenas YYYY-MM-DD)
Ignora a hora/timestamp.
Suporta múltiplos separadores (vírgula, ponto-e-vírgula).
"""

from datetime import datetime
from pathlib import Path
from typing import Optional

import pandas as pd
import re


def extract_date_from_filename(nome_arquivo: str) -> Optional[datetime]:
    """Extrai data no formato dd-mm-yyyy do nome do arquivo."""
    if not nome_arquivo:
        return None
    match = re.search(r"(\d{2}-\d{2}-\d{4})", nome_arquivo)
    if not match:
        return None
    try:
        return datetime.strptime(match.group(1), "%d-%m-%Y")
    except ValueError:
        return None


def encontrar_coluna_data(df) -> Optional[str]:
    """Encontra coluna que contém 'data' no nome (case-insensitive)."""
    for col in df.columns:
        col_lower = str(col).lower().strip()
        # Procura por 'data' em qualquer parte do nome da coluna
        if "data" in col_lower:
            return col
    return None


def extrair_data_do_valor(valor) -> Optional[str]:
    """
    Extrai a data de um valor e NORMALIZA para YYYY-MM-DD.

    Exemplos:
    - "2 jun., 2026, 12:00 AM" → "2026-06-02"
    - "2026-04-04 13:49:03.0000000" → "2026-04-04"
    - "2026-04-04" → "2026-04-04"
    - "04/04/2026 10:30:00" → "2026-04-04"
    - "04-04-2026" → "2026-04-04"
    """
    if pd.isna(valor) or valor is None:
        return None

    valor_str = str(valor).strip()
    if not valor_str or len(valor_str) < 8:
        return None

    # 1. Tentar formato português extenso abreviado no valor completo: "2 jun., 2026, 12:00 AM"
    MESES_PT = {
        "jan": "01", "fev": "02", "mar": "03", "abr": "04",
        "mai": "05", "jun": "06", "jul": "07", "ago": "08",
        "set": "09", "out": "10", "nov": "11", "dez": "12"
    }
    m_pt = re.search(r"(\d{1,2})\s+([a-z]{3})\.?,\s+(\d{4})", valor_str.lower())
    if m_pt:
        dia = m_pt.group(1).zfill(2)
        mes_str = m_pt.group(2)
        ano = m_pt.group(3)
        if mes_str in MESES_PT:
            return f"{ano}-{MESES_PT[mes_str]}-{dia}"

    # 2. Se tem espaço, tenta separar a data da hora
    partes_espaco = valor_str.split(" ")
    if len(partes_espaco) > 1 and ("-" in partes_espaco[0] or "/" in partes_espaco[0]):
        valor_str = partes_espaco[0]

    # 3. Detectar formato e normalizar para YYYY-MM-DD
    if "-" in valor_str:
        partes = valor_str.split("-")
        if len(partes) >= 3:
            if len(partes[0]) == 4:
                return valor_str[:10]
            else:
                try:
                    dia, mes, ano = partes[0], partes[1], partes[2][:4]
                    return f"{ano}-{mes.zfill(2)}-{dia.zfill(2)}"
                except (IndexError, ValueError):
                    return None

    if "/" in valor_str:
        partes = valor_str.split("/")
        if len(partes) == 3:
            if len(partes[0]) == 4:
                return f"{partes[0]}-{partes[1].zfill(2)}-{partes[2][:2].zfill(2)}"
            else:
                try:
                    dia, mes, ano = partes[0], partes[1], partes[2][:4]
                    return f"{ano}-{mes.zfill(2)}-{dia.zfill(2)}"
                except (IndexError, ValueError):
                    return None

    return None


def ler_csv_com_detector(caminho: str):
    """Tenta ler CSV testando diferentes separadores e opções de formatação."""
    separadores = [";", ",", "\t"]

    for sep in separadores:
        try:
            # Tenta com skiprows=[1] para ignorar linhas divisoras
            df = pd.read_csv(
                caminho,
                encoding="utf-8",
                sep=sep,
                nrows=500,
                skiprows=[1],  # Pula linhas com dashes/divisoras
                on_bad_lines='skip'  # Ignora linhas malformadas
            )
            if len(df.columns) > 1:  # Sucesso se temos múltiplas colunas
                return df
        except Exception:
            pass

    # Tenta sem skiprows
    for sep in separadores:
        try:
            df = pd.read_csv(
                caminho,
                encoding="utf-8",
                sep=sep,
                nrows=500,
                on_bad_lines='skip'
            )
            if len(df.columns) > 1:
                return df
        except Exception:
            pass

    # Fallback: tenta latin-1
    for sep in separadores:
        try:
            df = pd.read_csv(
                caminho,
                encoding="latin-1",
                sep=sep,
                nrows=500,
                skiprows=[1],
                on_bad_lines='skip'
            )
            if len(df.columns) > 1:
                return df
        except Exception:
            pass

    return None


def analisar_arquivo(caminho: str) -> Optional[dict]:
    """
    Analisa arquivo e retorna informações sobre a coluna com 'data'.

    Retorna:
    {
        "coluna": "nome_da_coluna",
        "data_mais_recente": "2026-04-10",
        "data_mais_antiga": "2026-04-01",
        "total_linhas": 100,
        "total_datas_validas": 98,
    }
    """
    p = Path(caminho)

    if not p.exists():
        return None

    try:
        # Tentar ler o arquivo
        if p.suffix.lower() in [".csv", ".txt"]:
            df = ler_csv_com_detector(str(caminho))
            if df is None:
                return None
        elif p.suffix.lower() in [".xlsx", ".xls"]:
            # Para histórico, lemos o arquivo inteiro para garantir que pegamos a data mais recente
            # que geralmente está no final ou espalhada.
            df = pd.read_excel(caminho)
        else:
            return None

    except Exception:
        return None

    # Encontrar coluna com 'data' no nome
    coluna_data = encontrar_coluna_data(df)
    if not coluna_data:
        return None

    # Extrair datas
    datas_validas = []
    for valor in df[coluna_data]:
        data_str = extrair_data_do_valor(valor)
        if data_str:
            datas_validas.append(data_str)

    if not datas_validas or len(datas_validas) < 1:
        return None

    # Ordenar datas (remover duplicatas)
    datas_validas_sorted = sorted(set(datas_validas))

    resultado = {
        "arquivo": p.name,
        "caminho": str(p),
        "coluna": coluna_data,
        "data_mais_recente": datas_validas_sorted[-1],
        "data_mais_antiga": datas_validas_sorted[0],
        "total_linhas": len(df),
        "total_datas_validas": len(datas_validas),
        "percentual": round((len(datas_validas) / len(df)) * 100, 1),
    }

    return resultado


def extrair_todas_as_datas(caminho: str) -> set[str]:
    """
    Retorna um conjunto de todas as datas únicas encontradas no arquivo (formato: YYYY-MM-DD).
    Lê apenas a coluna de data para economia de memória.
    Suporta múltiplas abas em Excel.
    """
    p = Path(caminho)
    if not p.exists():
        return set()

    todas_as_datas = set()
    
    try:
        if p.suffix.lower() in [".csv", ".txt"]:
            # Tenta diferentes encodings
            for enc in ["utf-8", "latin-1", "cp1252", "utf-8-sig"]:
                sucesso_enc = False
                # Para CSV, tenta descobrir o cabeçalho com offsets (pular linhas de lixo)
                for skip in range(6):
                    try:
                        # Tenta primeiro detectar o separador com uma amostra pequena
                        with open(p, mode="r", encoding=enc, errors="ignore") as f_obj:
                            df_head = pd.read_csv(f_obj, nrows=10, skiprows=skip, sep=None, engine='python', on_bad_lines='skip')
                        coluna_data = encontrar_coluna_data(df_head)
                        if coluna_data:
                            # Se achou a coluna, lê o arquivo inteiro (apenas a coluna de data)
                            with open(p, mode="r", encoding=enc, errors="ignore") as f_obj:
                                df = pd.read_csv(f_obj, usecols=[coluna_data], skiprows=skip, sep=None, engine='python', on_bad_lines='skip')
                            
                            # Filtra valores que são apenas traços ou vazios antes de extrair
                            valores = df[coluna_data].dropna().astype(str)
                            valores = valores[~valores.str.contains(r'^[- ]+$')]
                            
                            datas_extraidas = valores.apply(extrair_data_do_valor).dropna().unique()
                            todas_as_datas.update(datas_extraidas)
                            sucesso_enc = True
                            break
                    except Exception:
                        continue
                if sucesso_enc:
                    break
        else:
            # Para Excel, verifica todas as abas
            xl = pd.ExcelFile(caminho)
            for sheet_name in xl.sheet_names:
                for skip in range(6):
                    try:
                        df_head = pd.read_excel(xl, sheet_name=sheet_name, nrows=10, skiprows=skip)
                        coluna_data = encontrar_coluna_data(df_head)
                        if coluna_data:
                            df = pd.read_excel(xl, sheet_name=sheet_name, usecols=[coluna_data], skiprows=skip)
                            todas_as_datas.update(df[coluna_data].apply(extrair_data_do_valor).dropna().unique())
                            break
                    except Exception:
                        continue
    except Exception:
        pass

    return todas_as_datas


def extrair_data_do_arquivo(caminho: str) -> Optional[str]:
    """
    Retorna a data mais recente encontrada no arquivo (formato: YYYY-MM-DD).
    Lê o arquivo inteiro para garantir precisão em relatórios históricos.
    """
    datas = extrair_todas_as_datas(caminho)
    if datas:
        return sorted(list(datas))[-1]
    
    # Fallback para o método rápido de 500 linhas se o anterior falhar por algum motivo
    resultado = analisar_arquivo(caminho)
    if resultado:
        return resultado["data_mais_recente"]
    return None


# ==============================
# EXEMPLO DE USO
# ==============================
if __name__ == "__main__":
    from config import BASE_HISTORICO
    # Teste com arquivos de histórico reais se existirem
    historicos = [
        str(BASE_HISTORICO / "history_services.xlsx")
    ]

    print("=" * 80)
    print("TESTE DE EXTRAÇÃO MASSIVA DE DATAS")
    print("=" * 80)

    for h_str in historicos:
        h_path = Path(h_str)
        if not h_path.exists():
            print(f"❌ Arquivo não encontrado: {h_path}")
            continue

        print(f"\n📂 Analisando: {h_path.name}")
        start = datetime.now()
        datas = extrair_todas_as_datas(str(h_path))
        end = datetime.now()
        
        print(f"   ⏱ Tempo: {end - start}")
        print(f"   ✅ Datas encontradas: {len(datas)}")
        if datas:
            lista = sorted(list(datas))
            print(f"   📅 Range: {lista[0]} até {lista[-1]}")
            if len(lista) > 5:
                print(f"   📋 Amostra: {lista[:3]} ... {lista[-3:]}")

    print("\n" + "=" * 80)
