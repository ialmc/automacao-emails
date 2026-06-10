# -*- coding: utf-8 -*-
"""
reporter.py — Gerador e Transmissor de Relatório HTML
======================================================
⛔ NÃO redefina DISPLAY_MAP, DATA_CORTE ou REPORTER_KEYWORDS neste arquivo.
   Todas essas constantes vêm de config.py. Alterações aqui serão ignoradas.
"""
import json
import os
import smtplib
from datetime import date, datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
from typing import Any

from config import (
    BASE_ALLPE,
    ARQ_INVENTARIO_JSON,
    ARQ_LINKS_JSON,
    DATA_CORTE_DOWNLOAD_DATE,
    DISPLAY_MAP,
    DAILY_KEYS,
    REPORTER_KEYWORDS,
    REPORT_SUBJECT,
    REPORT_TITLE,
    REPORT_DEST_EMAIL,
)

# ==============================
# DESIGN TOKENS
# ==============================
BASE_AUTOMACAO = BASE_ALLPE  # alias para compatibilidade

# Cores do design.md
COLOR_BG     = "#0B1F2A"
COLOR_BG2    = "#112534"
COLOR_CYAN   = "#00FFFF"
COLOR_GREEN  = "#22C55E"
COLOR_RED    = "#F87171"
COLOR_FG     = "#FFFFFF"
COLOR_FG2    = "#94A3B8"
COLOR_BORDER = "#1A3A4A"

# DISPLAY_MAP e demais constantes vêm de config.py (importado acima)

def carregar_inventario() -> dict[str, Any]:
    if ARQ_INVENTARIO_JSON.exists():
        try:
            return json.loads(ARQ_INVENTARIO_JSON.read_text(encoding="utf-8"))
        except Exception: pass
    return {}

def carregar_links() -> dict[str, Any]:
    if ARQ_LINKS_JSON.exists():
        try:
            return json.loads(ARQ_LINKS_JSON.read_text(encoding="utf-8"))
        except Exception: pass
    return {}

def format_dates_br(iso_list: list[str]) -> str:
    """Formata lista de datas ISO para exibição BR (DD/MM).
    Mostra até 3 datas + total real quando há mais, para que o operador
    saiba a dimensão exata do problema sem precisar contar manualmente.
    """
    if not iso_list:
        return "OK"
    br_dates = [d[8:10] + '/' + d[5:7] for d in iso_list]
    if len(br_dates) > 3:
        return ", ".join(br_dates[:3]) + f" ... (+{len(br_dates) - 3} datas)"
    return ", ".join(br_dates)

def enviar_email(html_content, to_email, user, password):
    msg = MIMEMultipart("alternative")
    msg["Subject"] = f"{REPORT_SUBJECT} — {date.today().strftime('%d/%m/%Y')}"
    msg["From"] = user
    msg["To"] = to_email
    msg.attach(MIMEText(html_content, "html"))
    
    with smtplib.SMTP("smtp.zoho.com", 587) as server:
        server.starttls()
        server.login(user, password)
        server.sendmail(user, to_email, msg.as_string())

def build_html_report(inventory: dict[str, Any], links_state: dict[str, Any], hoje: date) -> str:
    # data_max_esperada agora é calculada por item (atraso_permitido)
    data_corte = DATA_CORTE_DOWNLOAD_DATE
    
    # Mapear links recentes para cada assunto com heurística de palavras-chave
    links_recentes = {}
    links_timestamps = {}
    KEYWORDS = REPORTER_KEYWORDS  # Importado de config.py — não redefinir aqui

    for url, info in links_state.items():
        if info:
            search_text = (url + (info.get("file_name") or info.get("file") or "") + (info.get("assunto") or info.get("subject") or "")).lower()
            updated_at = info.get("updated_at") or info.get("at") or info.get("last_seen_at") or "2000-01-01"
            for key, kws in KEYWORDS.items():
                is_daily = DISPLAY_MAP.get(key, {}).get("freq") == "Diária"
                if is_daily and any(x in search_text for x in ["mensal", "retroativo"]):
                    continue
                if key == "Estoque Diario" and "mensal" in search_text:
                    continue
                if key == "Estoque Mensal" and "diario" in search_text:
                    continue
                
                if any(kw in search_text for kw in kws):
                    if key not in links_timestamps or updated_at > links_timestamps[key]:
                        links_recentes[key] = url
                        links_timestamps[key] = updated_at

    total_emails = len(links_state)
    sucessos = 0
    atrasos_count = 0
    table_rows_html = ""
    gaps_section_html = ""
    
    for key, conf in DISPLAY_MAP.items():
        inv = inventory.get(key, {"dates": [], "sources": {}})
        dates_covered = set(inv.get("dates", []))
        missing = []
        status_text = "OK"
        color_status = COLOR_GREEN
        
        if conf["freq"] == "Diária":
            # Regra de delay: Consolidados e Formas de Pagamento são D-2, o restante é D-1
            atraso_permitido = 2 if ("Consolidado" in key or "Pagamento" in key) else 1
            data_max_item = hoje - timedelta(days=atraso_permitido)
            
            curr = data_corte
            while curr <= data_max_item:
                if curr.isoformat() not in dates_covered:
                    missing.append(curr.isoformat())
                curr += timedelta(days=1)
            if missing:
                # Conta o total de arquivos/datas pendentes, não apenas o número de pastas com atraso
                atrasos_count += len(missing)
                status_text = format_dates_br(missing)
                color_status = COLOR_RED
            else:
                sucessos += 1
        else:
            all_dates = inv.get("dates", [])
            if all_dates:
                ultima_data = max(all_dates)
                atraso_dias = (hoje - date.fromisoformat(ultima_data)).days
                if atraso_dias > conf["rag"] * 1.5:
                    status_text = f"{atraso_dias}d"
                    color_status = COLOR_RED
                    atrasos_count += 1
                else:
                    status_text = "OK"
                    sucessos += 1
            else:
                status_text = "Vazio"
                color_status = COLOR_FG2
        
        all_dates = inv.get("dates", [])
        ultimo_recb = max(all_dates) if all_dates else "-"
        if ultimo_recb != "-": ultimo_recb = datetime.fromisoformat(ultimo_recb).strftime("%d/%m/%Y")
        email_hoje = "Sim" if hoje.isoformat() in dates_covered else "Não"
        
        # TABELA DE DETALHAMENTO (SEM LINKS)
        table_rows_html += f"""
        <tr style="border-bottom: 1px solid {COLOR_BORDER}; color: {COLOR_FG if color_status == COLOR_GREEN else COLOR_RED}">
            <td style="padding: 10px; font-size: 13px;">{conf['nome']}</td>
            <td style="padding: 10px; font-size: 13px;" class="hide-mobile">{conf['freq']}</td>
            <td style="padding: 10px; font-size: 13px;" class="hide-mobile">{email_hoje}</td>
            <td style="padding: 10px; font-size: 13px; text-decoration: underline;">{ultimo_recb}</td>
            <td style="padding: 10px; font-size: 13px; font-weight: bold; color: {color_status}">{status_text}</td>
        </tr>
        """
        
        # SEÇÃO DE PENDÊNCIAS (COM LINKS)
        if color_status == COLOR_RED:
            link_url = links_recentes.get(key)
            link_html = f'<a href="{link_url}" style="color: {COLOR_CYAN}; text-decoration: underline; word-break: break-all;">{link_url}</a>' if link_url else '<span style="color: #666;">Link não localizado recentemente</span>'
            
            if conf["freq"] == "Diária":
                pend_info = f"<b>Datas não recebidas:</b> <span style='color: {COLOR_RED}; font-weight: bold'>{', '.join([d[8:10]+'/'+d[5:7] for d in missing])}</span>"
            else:
                pend_info = f"<b>Atraso detectado:</b> <span style='color: {COLOR_RED}; font-weight: bold'>{status_text}</span> (Último em {ultimo_recb})"

            gaps_section_html += f"""
            <div style="margin-bottom: 15px; padding-bottom: 15px; border-bottom: 1px solid {COLOR_BORDER}">
              <div style="font-size: 12px; color: {COLOR_FG}; margin-bottom: 4px">
                <b>Pasta:</b> {conf['nome']} &nbsp;|&nbsp; <b>Frequência:</b> {conf['freq']}
              </div>
              <div style="font-size: 11px; color: {COLOR_FG2}; margin-bottom: 4px">
                {pend_info}
              </div>
              <div style="font-size: 11px; color: {COLOR_FG2}">
                <b>Último link:</b> {link_html}
              </div>
            </div>
            """

    exec_id = datetime.now().strftime("%Y%m%d_%H%M")
    html = f"""
    <!DOCTYPE html>
    <html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <meta name="color-scheme" content="dark light">
        <style>
            @media only screen and (max-width: 600px) {{
                .hide-mobile {{ display: none !important; }}
                .stat-card {{ width: 100% !important; margin: 5px 0 !important; }}
                .container {{ padding: 15px !important; }}
            }}
        </style>
    </head>
    <body style="margin: 0; padding: 0; background-color: #050d12;">
        <div style="background-color: #050d12; padding: 20px 10px;">
            <div class="container" style="max-width: 800px; margin: 0 auto; background-color: {COLOR_BG2}; border: 1px solid {COLOR_BORDER}; border-radius: 12px; padding: 30px; color: {COLOR_FG};">
                <h1 style="margin: 0 0 20px 0; font-size: 20px; color: {COLOR_CYAN}; font-family: 'DM Sans', sans-serif;">
                    {REPORT_TITLE} — {hoje.strftime('%d/%m/%Y')}
                </h1>
                <div style="display: flex; flex-wrap: wrap; margin-bottom: 30px; gap: 10px;">
                    <div class="stat-card" style="flex: 1; min-width: 120px; background: {COLOR_BG}; padding: 15px; border-radius: 8px; border-left: 4px solid {COLOR_CYAN}">
                        <div style="font-size: 10px; color: {COLOR_FG2}; font-weight: bold; margin-bottom: 5px">E-MAILS</div>
                        <div style="font-size: 24px; font-weight: bold; color: {COLOR_FG}">{total_emails}</div>
                    </div>
                    <div class="stat-card" style="flex: 1; min-width: 120px; background: {COLOR_BG}; padding: 15px; border-radius: 8px; border-left: 4px solid {COLOR_GREEN}">
                        <div style="font-size: 10px; color: {COLOR_FG2}; font-weight: bold; margin-bottom: 5px">SUCESSO</div>
                        <div style="font-size: 24px; font-weight: bold; color: {COLOR_FG}">{sucessos}</div>
                    </div>
                    <div class="stat-card" style="flex: 1; min-width: 120px; background: {COLOR_BG}; padding: 15px; border-radius: 8px; border-left: 4px solid {COLOR_RED}">
                        <div style="font-size: 10px; color: {COLOR_FG2}; font-weight: bold; margin-bottom: 5px">ATRASOS</div>
                        <div style="font-size: 24px; font-weight: bold; color: {COLOR_RED}">{atrasos_count}</div>
                    </div>
                </div>
                <h3 style="margin: 0 0 15px 0; font-size: 14px; color: {COLOR_FG}; text-transform: uppercase;">Detalhamento</h3>
                <table width="100%" style="border-collapse: collapse; margin-bottom: 30px;">
                    <thead>
                        <tr style="border-bottom: 2px solid {COLOR_BORDER}; color: {COLOR_FG2}; font-size: 11px; text-transform: uppercase;">
                            <th align="left" style="padding: 10px;">Pasta</th>
                            <th align="left" style="padding: 10px;" class="hide-mobile">Freq.</th>
                            <th align="left" style="padding: 10px;" class="hide-mobile">Hoje?</th>
                            <th align="left" style="padding: 10px;">Último</th>
                            <th align="left" style="padding: 10px;">Atraso</th>
                        </tr>
                    </thead>
                    <tbody>
                        {table_rows_html}
                    </tbody>
                </table>
                <div style="border: 1px dashed {COLOR_RED}; border-radius: 8px; padding: 20px; background: rgba(248, 113, 113, 0.05)">
                    <h3 style="margin: 0 0 15px 0; color: {COLOR_RED}; font-size: 14px; text-transform: uppercase;">Relatórios não enviados</h3>
                    {gaps_section_html}
                </div>
                <div style="margin-top: 30px; padding-top: 20px; border-top: 1px solid {COLOR_BORDER}; text-align: center; font-size: 10px; color: {COLOR_FG2}">
                    Execução ID: {exec_id} | LMC Inteligência & Data
                </div>
            </div>
        </div>
    </body>
    </html>
    """
    return html

def exibir_status_operador(inventory: dict, links_state: dict, hoje: date):
    DAILY_KEYS = [
        "Consolidado de servicos",
        "Consolidado de produtos",
        "Consolidado de pacotes",
        "Agendamentos",
        "Despesas",
        "Estoque Diario",
        "Formas de Pagamento"
    ]
    
    DAILY_KWS = [
        "consolidado_servicos", "all_pe_consolidado_servicos",
        "consolidado_produtos", "all_pe_consolidado_produtos",
        "consolidado_pacotes", "all_pe_consolidado_pacotes",
        "relatorio_agendamento", "agendamento",
        "despesa_diaria", "despesa",
        "posicao_estoque_diario", "estoque_diario",
        "consolidados_forma_pagamento", "forma_pagamento"
    ]
    
    baixados_hoje_count = 0
    for url, info in links_state.items():
        if info and info.get("status") in ["downloaded", "recovered"]:
            at_str = info.get("at") or info.get("updated_at") or info.get("last_seen_at")
            if at_str:
                try:
                    dt = datetime.fromisoformat(at_str[:19]).date()
                except Exception:
                    continue
                if dt == hoje:
                    search_text = (url + (info.get("file_name") or info.get("file") or "") + (info.get("assunto") or info.get("subject") or "")).lower()
                    if any(x in search_text for x in ["mensal", "retroativo"]):
                        continue
                    if any(kw in search_text for kw in DAILY_KWS):
                        baixados_hoje_count += 1
                        
    if baixados_hoje_count >= 8:
        print("\n==================================================================================")
        print(f"extração  {hoje.strftime('%d/%m/%y')} concluida com sucesso, Não há arquivos nada data de hoje para serem baixados.")
        
        data_corte = DATA_CORTE_DOWNLOAD_DATE
        links_recentes = {}
        links_timestamps = {}
        KEYWORDS = REPORTER_KEYWORDS  # de config.py

        for url, info in links_state.items():
            if info:
                search_text = (url + (info.get("file_name") or info.get("file") or "") + (info.get("assunto") or info.get("subject") or "")).lower()
                updated_at = info.get("updated_at") or info.get("at") or info.get("last_seen_at") or "2000-01-01"
                for key, kws in KEYWORDS.items():
                    is_daily = DISPLAY_MAP.get(key, {}).get("freq") == "Diária"
                    if is_daily and any(x in search_text for x in ["mensal", "retroativo"]):
                        continue
                    if key == "Estoque Diario" and "mensal" in search_text:
                        continue
                    
                    if any(kw in search_text for kw in kws):
                        if key not in links_timestamps or updated_at > links_timestamps[key]:
                            links_recentes[key] = url
                            links_timestamps[key] = updated_at

        pendencias = []
        for key in DAILY_KEYS:
            conf = DISPLAY_MAP[key]
            inv = inventory.get(key, {"dates": [], "sources": {}})
            dates_covered = set(inv.get("dates", []))
            missing = []
            
            atraso_permitido = 2 if ("Consolidado" in key or "Pagamento" in key) else 1
            data_max_item = hoje - timedelta(days=atraso_permitido)
            
            curr = data_corte
            while curr <= data_max_item:
                if curr.isoformat() not in dates_covered:
                    missing.append(curr.isoformat())
                curr += timedelta(days=1)
            
            if missing:
                missing_fmt = [datetime.strptime(d, "%Y-%m-%d").strftime("%d/%m") for d in missing]
                pendencias.append({
                    "pasta": conf["nome"],
                    "freq": conf["freq"],
                    "datas": ", ".join(missing_fmt),
                    "link": links_recentes.get(key, "Link não localizado recentemente")
                })
                
        if pendencias:
            print("há somente pendências anteriores: \n")
            print("Relatórios não enviados")
            for p in pendencias:
                print(f"Pasta: {p['pasta']}  |  Frequência: {p['freq']}")
                print(f"Datas não recebidas: {p['datas']}")
                print(f"Último link: {p['link']}")
        print("==================================================================================\n")
    else:
        print("\n==================================================================================")
        print("[X] ALERTA DE FALHA NA EXTRAÇÃO DIÁRIA")
        print(f"Apenas {baixados_hoje_count} de 8 arquivos diários foram baixados hoje ({hoje.strftime('%d/%m/%Y')})!")
        print("==================================================================================")
        
        data_corte = DATA_CORTE_DOWNLOAD_DATE
        links_recentes = {}
        links_timestamps = {}
        KEYWORDS = REPORTER_KEYWORDS  # de config.py

        for url, info in links_state.items():
            if info:
                search_text = (url + (info.get("file_name") or info.get("file") or "") + (info.get("assunto") or info.get("subject") or "")).lower()
                updated_at = info.get("updated_at") or info.get("at") or info.get("last_seen_at") or "2000-01-01"
                for key, kws in KEYWORDS.items():
                    is_daily = DISPLAY_MAP.get(key, {}).get("freq") == "Diária"
                    if is_daily and any(x in search_text for x in ["mensal", "retroativo"]):
                        continue
                    if key == "Estoque Diario" and "mensal" in search_text:
                        continue
                    
                    if any(kw in search_text for kw in kws):
                        if key not in links_timestamps or updated_at > links_timestamps[key]:
                            links_recentes[key] = url
                            links_timestamps[key] = updated_at

        print("Status dos relatórios de hoje (incluindo falhas):")
        for key in DAILY_KEYS:
            conf = DISPLAY_MAP[key]
            baixado_hoje = False
            for url, info in links_state.items():
                if info and info.get("status") in ["downloaded", "recovered"]:
                    at_str = info.get("at") or info.get("updated_at") or info.get("last_seen_at")
                    if at_str:
                        try:
                            dt = datetime.fromisoformat(at_str[:19]).date()
                        except Exception:
                            continue
                        if dt == hoje:
                            search_text = (url + (info.get("file_name") or info.get("file") or "") + (info.get("assunto") or info.get("subject") or "")).lower()
                            if any(x in search_text for x in ["mensal", "retroativo"]):
                                continue
                            if any(kw in search_text for kw in KEYWORDS[key]):
                                baixado_hoje = True
                                break
            
            status_desc = "[OK] BAIXADO COM SUCESSO" if baixado_hoje else "[X] FALHA NO DOWNLOAD / AUSENTE"
            print(f" - Pasta: {conf['nome']} -> {status_desc}")

        pendencias = []
        for key in DAILY_KEYS:
            conf = DISPLAY_MAP[key]
            inv = inventory.get(key, {"dates": [], "sources": {}})
            dates_covered = set(inv.get("dates", []))
            missing = []
            
            atraso_permitido = 2 if ("Consolidado" in key or "Pagamento" in key) else 1
            data_max_item = hoje - timedelta(days=atraso_permitido)
            
            curr = data_corte
            while curr <= data_max_item:
                if curr.isoformat() not in dates_covered:
                    missing.append(curr.isoformat())
                curr += timedelta(days=1)
            
            if missing:
                missing_fmt = [datetime.strptime(d, "%Y-%m-%d").strftime("%d/%m") for d in missing]
                pendencias.append({
                    "pasta": conf["nome"],
                    "freq": conf["freq"],
                    "datas": ", ".join(missing_fmt),
                    "link": links_recentes.get(key, "Link não localizado recentemente")
                })
                
        if pendencias:
            print("\nRelatórios não enviados / Pendentes no repositório:")
            for p in pendencias:
                print(f"Pasta: {p['pasta']}  |  Frequência: {p['freq']}")
                print(f"Datas não recebidas: {p['datas']}")
                print(f"Último link: {p['link']}")
        print("==================================================================================\n")

if __name__ == "__main__":
    print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Iniciando Reporter...")
    inventory = carregar_inventario()
    links_state = carregar_links()
    if inventory:
        print(f"   Inventário carregado com {len(inventory)} itens.")
        hoje_ref = date.today()
        html = build_html_report(inventory, links_state, hoje_ref)
        (BASE_AUTOMACAO / "preview_report.html").write_text(html, encoding="utf-8")
        
        email_user = os.environ.get("SMTP_EMAIL_USER") or os.environ.get("ALLPE_EMAIL_USER")
        email_pass = os.environ.get("SMTP_EMAIL_PASS") or os.environ.get("ALLPE_EMAIL_PASS")
        if email_user and email_pass:
            try:
                print(f"   Enviando e-mail para {REPORT_DEST_EMAIL} via {email_user}...")
                enviar_email(html, REPORT_DEST_EMAIL, email_user, email_pass)
                print("   [SUCESSO] Relatório enviado.")
            except Exception as e:
                print(f"   [ERRO] Falha ao enviar e-mail: {e}")
        else:
            print("   [AVISO] Envio pulado (credenciais SMTP_EMAIL_USER/PASS não encontradas).")
            
        exibir_status_operador(inventory, links_state, hoje_ref)
    else:
        print("   [AVISO] Inventário vazio ou não encontrado. Relatório não gerado.")
