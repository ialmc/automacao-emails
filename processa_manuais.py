import sys
import shutil
import os
from pathlib import Path
from datetime import datetime

# Garante a importação dos módulos locais de forma portátil
SCRIPT_DIR = Path(__file__).parent.resolve()
sys.path.append(str(SCRIPT_DIR))

from config import ASSUNTOS_PASTAS, PREFIXOS_FORCADOS
from detector_datas import extrair_data_do_arquivo
from downloader import gerar_nome_final, carregar_state, salvar_state

temp_dir = SCRIPT_DIR / "temp_manuais"

def get_ass_ref(filepath: Path):
    filename = filepath.name
    fname = filename.lower()
    if 'agendamento' in fname:
        return '[Trinks][AllPe] Extracao Relatorio de Agendamento - Email automatico'
    if 'forma_pagamento' in fname:
        return '[Trinks][AllPe] Formas de pagamentos nas transacoes - Email automatico'
    if 'estoque_diario' in fname:
        return '[Trinks][AllPe] Extracao de estoque diario - Email automatico'
    if 'pacotes' in fname:
        return '[Trinks][AllPe] Consolidado de pacotes diario - Email automatico'
    if 'produtos' in fname:
        return '[Trinks][AllPe] Consolidado de produtos diario - Email automatico'
    if 'servicos' in fname:
        return '[Trinks][AllPe] Consolidado de servicos diario - Email automatico'
        
    # Fallback para arquivos genéricos (como query_result_...): analisa as colunas e conteúdo
    try:
        import csv
        with open(filepath, mode='r', encoding='utf-8', errors='ignore') as file:
            # Tenta descobrir o separador correto
            sample = file.read(4096)
            file.seek(0)
            sep = ';' if ';' in sample else ','
            reader = csv.reader(file, delimiter=sep)
            headers = next(reader, None)
            if not headers:
                return None
                
            headers_norm = [h.strip().lower() for h in headers]
            
            # 1. Agendamento
            if any('agendamento' in h for h in headers_norm):
                return '[Trinks][AllPe] Extracao Relatorio de Agendamento - Email automatico'
                
            # 2. Despesa
            if any('codigo despesa' in h for h in headers_norm) or any('vencimento' in h for h in headers_norm):
                return '[Trinks][AllPe] Despesa diaria - Email automatico'
                
            # 3. Formas de Pagamento
            if any('tipo forma pagamento' in h for h in headers_norm):
                return '[Trinks][AllPe] Formas de pagamentos nas transacoes - Email automatico'
                
            # 4. Consolidados (Serviço, Produto, Pacote)
            if 'item do fechamento de contas' in headers_norm:
                col_idx = headers_norm.index('item do fechamento de contas')
                # Ler algumas linhas para identificar o tipo do item
                for _ in range(100):
                    row = next(reader, None)
                    if not row:
                        break
                    if col_idx < len(row):
                        val = row[col_idx].strip().lower()
                        if 'produto' in val:
                            return '[Trinks][AllPe] Consolidado de produtos diario - Email automatico'
                        elif 'pacote' in val:
                            return '[Trinks][AllPe] Consolidado de pacotes diario - Email automatico'
                        elif 'servi' in val or 'servico' in val:
                            return '[Trinks][AllPe] Consolidado de servicos diario - Email automatico'
                            
            # 5. Estoque Diário
            if any('estoque' in h for h in headers_norm) and any('posicao' in h for h in headers_norm):
                return '[Trinks][AllPe] Extracao de estoque diario - Email automatico'
                
    except Exception as e:
        print(f"[AVISO] Falha ao identificar o tipo do arquivo {filename}: {e}")
        
    return None

def run():
    from consolida_historico import run_consolida_historico
    run_consolida_historico()

if __name__ == '__main__':
    run()
