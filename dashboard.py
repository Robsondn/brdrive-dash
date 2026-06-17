"""
BR DRIVE — Dashboard Generator
Gera o HTML completo com todas as 15 páginas a partir de dados Excel/CSV.
"""

import pandas as pd
import numpy as np
import webbrowser
import os
import sys
from datetime import datetime

# Força UTF-8 no terminal Windows
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
if hasattr(sys.stderr, 'reconfigure'):
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')

# ── Lê o Excel ─────────────────────────────────────────────────────────────
# Caminhos por máquina — adicione novas entradas conforme necessário
_CAMINHOS = {
    "robson.noberto":    r"C:\Users\robson.noberto\Desktop\Power BI Gus\BRDrive_BI_Novo.xlsx",
    "Robo Transporte":   r"C:\Users\Robo Transporte\OneDrive - J&T EXPRESS - FILIAL SP\Power BI Gus\BRDrive_BI_Novo.xlsx",
}
import getpass as _gp
CAMINHO_PADRAO = _CAMINHOS.get(_gp.getuser(), next(iter(_CAMINHOS.values())))
ABA_PADRAO = "fPrincipal"

def carregar_dados(caminho=None, aba=ABA_PADRAO):
    if caminho is None:
        caminho = CAMINHO_PADRAO
    try:
        df = pd.read_excel(caminho, sheet_name=aba)
        print(f"✅ Arquivo carregado: {caminho}")
        print(f"   Aba: {aba}")
    except FileNotFoundError:
        print(f"❌ Arquivo '{caminho}' não encontrado.")
        print("📄 Usando dados de exemplo...")
        df = gerar_dados_exemplo()
    df.columns = [c.strip() for c in df.columns]
    if "DATA" in df.columns:
        df["DATA"] = pd.to_datetime(df["DATA"], errors="coerce")
    # Normaliza e filtra tipos de operação válidos
    if "TIPO DE OPERAÇÃO" in df.columns:
        df["TIPO DE OPERAÇÃO"] = df["TIPO DE OPERAÇÃO"].str.strip().str.title()
        df["TIPO DE OPERAÇÃO"] = df["TIPO DE OPERAÇÃO"].replace({"Coleta Pa": "Coleta PA"})
        tipos_validos = ["Coleta Base", "Coleta PA", "Devolução", "Secundaria"]
        df = df[df["TIPO DE OPERAÇÃO"].isin(tipos_validos)].copy()
    # Normaliza pontualidade (capitalização inconsistente)
    for col in ["PONTUALIDADE SAÍDA", "PONTUALIDADE CHEGADA"]:
        if col in df.columns:
            df[col] = df[col].str.strip().str.title()
            df[col] = df[col].replace({"No Prazo": "No prazo", "Fora Do Prazo": "Fora do prazo"})
    return df

# ── Gera dados de exemplo se não tiver Excel ───────────────────────────────
def gerar_dados_exemplo():
    import random
    random.seed(42)
    condutores = ["ADEMIR ALBUQUERQUE","ANTONIO NETO","CARLOS SILVA","DIEGO FERREIRA",
                  "EDILSON RAMOS","FABIO SANTOS","GILMAR TOMAZ","JOSE MARTINS",
                  "LUCAS ROBERTO","MARCOS OLIVEIRA","PAULO CESAR","ROBSON COSTA"]
    transportadoras = ["JET SP","CEOS EXPRE","SAMY","PACHECO","FAP","DEB","EASY"]
    tipos = ["Coleta base","Devolução","Secundaria","Coleta PA"]
    origens = ["SP BRE","SP CAM","SP OSA","SP SJC"]
    rotas = ["BRE-S247-0500-1","BRE-S393-0300-1","CAM-S100-0200-1","OSA-S050-0100-1"]
    rows = []
    for i in range(200):
        data = pd.Timestamp("2026-05-" + str(random.randint(19,25)).zfill(2))
        mes = "MAIO"
        cond = random.choice(condutores)
        trans = random.choice(transportadoras)
        tipo = random.choice(tipos)
        off_s = random.choices(["OK","OFF"], weights=[85,15])[0]
        off_c = random.choices(["OK","OFF"], weights=[82,18])[0]
        pont_s = random.choices(["No prazo","Atrasado"], weights=[87,13])[0]
        pont_c = random.choices(["No prazo","Atrasado"], weights=[79,21])[0]
        lacre_s = random.choices(["OK","OFF"], weights=[70,30])[0] if pont_s == "Atrasado" else "OK"
        deslacre_c = random.choices(["OK","OFF"], weights=[70,30])[0] if pont_c == "Atrasado" else "OK"
        fez_rota = random.choices([True,False], weights=[90,10])[0]
        rows.append({
            "DATA": data, "MÊS": mes, "Ano": 2026,
            "Semana": "S4 mai", "Semana Ordem Nova": 202621,
            "Semana Nome Nova": "W21 - 2026",
            "Transportador": trans, "CONDUTOR": cond,
            "TIPO DE OPERAÇÃO": tipo,
            "Número do ID": f"SRTR2260{2300000+i}",
            "Tempo Saida OFF": off_s, "Tempo chegada OFF": off_c,
            "PONTUALIDADE SAÍDA": pont_s, "PONTUALIDADE CHEGADA": pont_c,
            "PONTUALIDADE SAÍDA CORRIGIDA": pont_s, "PONTUALIDADE CHEGADA CORRIGIDA": pont_c,
            "STATUS LACRE SAIDA": lacre_s, "STATUS DESLACRE CHEGADA": deslacre_c,
            "ORIGEM": random.choice(origens),
            "Subtipo de linha": random.choice(rotas),
            "Horário real de saída": "07:30" if fez_rota else None,
        })
    return pd.DataFrame(rows)

# ══════════════════════════════════════════════════════════════════════════════
# FUNÇÕES DE CÁLCULO (equivalentes às VAR do DAX)
# ══════════════════════════════════════════════════════════════════════════════

def cor_kpi(v, meta=85, aten=70):
    if v >= meta: return "#27ae60"
    if v >= aten: return "#e67e22"
    return "#e74c3c"

def fmt_pct(v): return f"{v:.1f}%"

# ══════════════════════════════════════════════════════════════════════════════
# COMPONENTES HTML REUTILIZÁVEIS
# ══════════════════════════════════════════════════════════════════════════════

def navbar(pagina_ativa):
    paginas = [
        ("home",           "🏠", "Home",               "首页"),
        ("mural",          "📺", "Mural",              "仪表板墙"),
        ("dashboard",      "📊", "Utilização App",      "应用使用情况"),
        ("pont_op",        "⏱",  "Pontualidade Op.",    "运营准时率"),
        ("evol_semanal",   "📈", "App Semanal",         "应用周度绩效"),
        ("app_mensal",     "📅", "App Mensal",          "APP月度绩效"),
        ("infracoes",      "🚨", "Infrações",           "违章管理"),
        ("motoristas",     "👤", "Motoristas APP",      "司机管理"),
        ("pont_semanal",   "📊", "Pont. Semanal",       "准时率周度"),
        ("pont_mensal",    "📆", "Pont. Mensal",        "准时率月度"),
        ("inf_pontualidade","🔴","Inf. Pontualidade",   "准时率违规"),
        ("pont_motoristas", "🚗","Pont. Motoristas",    "司机准时率"),
    ]
    items = ""
    for pid, icon, label_pt, label_zh in paginas:
        active = "active" if pid == pagina_ativa else ""
        items += (
            f'<a href="{pid}.html" class="nav-item {active}" title="{label_pt} / {label_zh}">'
            f'{icon}'
            f'<span style="display:flex;flex-direction:column;gap:1px;">'
            f'<span style="font-size:11px;">{label_pt}</span>'
            f'<span style="font-size:9px;color:rgba(255,255,255,0.55);">{label_zh}</span>'
            f'</span></a>'
        )
    return f"""
    <nav class="sidebar">
      <div class="sidebar-logo">
        <div class="logo-box">J&T</div>
        <span>BR DRIVE</span>
      </div>
      <div class="nav-links">{items}</div>
      <div class="sidebar-footer">v2.0 · Python</div>
    </nav>"""

def kpi_card(titulo, valor, subtitulo, cor_borda, cor_valor="#fff", extra_html=""):
    return f"""
    <div class="kpi-card" style="border-top:3px solid {cor_borda};">
      <div class="kpi-label">{titulo}</div>
      <div class="kpi-value" style="color:{cor_valor};">{valor}</div>
      <div class="kpi-sub">{subtitulo}</div>
      {extra_html}
    </div>"""

def barra_progresso(pct, cor, width="100%"):
    return f"""
    <div class="progress-bar" style="width:{width};">
      <div class="progress-fill" style="width:{min(pct,100):.0f}%;background:{cor};"></div>
    </div>"""

def badge(texto, cor_bg, cor_txt):
    return f"<span class='badge' style='background:{cor_bg};color:{cor_txt};'>{texto}</span>"

CSS_GLOBAL = """
* { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: Arial, sans-serif; background: #0e0808; color: #fff; display: flex; min-height: 100vh; }

/* SIDEBAR */
.sidebar { width: 220px; min-height: 100vh; background: #1a0a0a; border-right: 1px solid rgba(192,57,43,0.3); display: flex; flex-direction: column; position: fixed; top: 0; left: 0; z-index: 100; overflow-y: auto; }
.sidebar-logo { display: flex; align-items: center; gap: 10px; padding: 16px 14px; border-bottom: 1px solid rgba(192,57,43,0.3); }
.logo-box { background: #C0392B; color: #fff; font-weight: 900; font-size: 13px; padding: 6px 8px; border-radius: 6px; }
.sidebar-logo span { color: #fff; font-size: 14px; font-weight: 700; letter-spacing: 1px; }
.nav-links { flex: 1; padding: 8px 0; }
.nav-item { display: flex; align-items: center; gap: 8px; padding: 8px 14px; color: rgba(255,255,255,0.9); text-decoration: none; font-size: 11px; transition: all 0.15s; }
.nav-item:hover { background: rgba(192,57,43,0.15); color: #fff; }
.nav-item.active { background: rgba(192,57,43,0.25); color: #fff; border-left: 3px solid #C0392B; }
.sidebar-footer { padding: 10px 14px; color: rgba(255,255,255,0.3); font-size: 9px; border-top: 1px solid rgba(255,255,255,0.07); }

/* MAIN */
.main { margin-left: 220px; flex: 1; padding: 16px; min-height: 100vh; }

/* HEADER */
.page-header { display: flex; align-items: center; justify-content: space-between; border-bottom: 1px solid rgba(192,57,43,0.4); padding-bottom: 10px; margin-bottom: 14px; }
.header-left { display: flex; align-items: center; gap: 10px; }
.header-icon { width: 36px; height: 36px; border-radius: 6px; display: flex; align-items: center; justify-content: center; font-weight: 700; color: #fff; font-size: 12px; flex-shrink: 0; }
.header-title { font-size: 14px; font-weight: 700; letter-spacing: 1px; }
.header-sub { color: #ddd; font-size: 9px; margin-top: 2px; }
.header-badges { display: flex; gap: 6px; }
.header-badge { background: #1a1010; border: 1px solid rgba(192,57,43,0.3); border-radius: 5px; padding: 4px 10px; text-align: center; }
.header-badge .hb-label { color: #ddd; font-size: 7px; letter-spacing: 1px; }
.header-badge .hb-value { color: #C0392B; font-size: 11px; font-weight: 700; }

/* KPI CARDS */
.kpi-grid { display: grid; gap: 8px; margin-bottom: 14px; }
.kpi-grid-5 { grid-template-columns: repeat(5,1fr); }
.kpi-grid-4 { grid-template-columns: repeat(4,1fr); }
.kpi-card { background: #1a1010; border-radius: 8px; padding: 10px 14px; }
.kpi-label { font-size: 8px; letter-spacing: 1px; color: #ddd; margin-bottom: 3px; }
.kpi-value { font-size: 24px; font-weight: 700; line-height: 1; }
.kpi-sub { font-size: 9px; color: #ddd; margin-top: 3px; }

/* TABLES */
.table-wrap { background: #1a1010; border-radius: 8px; overflow: hidden; }
.table-header { padding: 8px 14px; display: flex; justify-content: space-between; align-items: center; }
.table-header span { color: #fff; font-size: 11px; font-weight: 700; letter-spacing: 1px; }
.table-header small { color: rgba(255,255,255,0.9); font-size: 9px; }
.table-scroll { overflow-y: auto; max-height: calc(100vh - 340px); }
table { width: 100%; border-collapse: collapse; }
thead { position: sticky; top: 0; z-index: 2; }
th { padding: 8px 10px; font-size: 9px; letter-spacing: 1px; font-weight: 700; }
td { padding: 7px 10px; font-size: 11px; border-bottom: 1px solid rgba(255,255,255,0.04); }
tfoot tr { position: sticky; bottom: 0; z-index: 2; }

/* MISC */
.badge { font-size: 10px; padding: 2px 8px; border-radius: 3px; font-weight: 700; }
.trans-badge { font-size: 10px; padding: 2px 7px; border-radius: 3px; background: rgba(26,82,118,0.3); color: #2980b9; font-weight: 600; }
.progress-bar { background: rgba(255,255,255,0.07); border-radius: 3px; height: 6px; width: 100%; }
.progress-fill { height: 6px; border-radius: 3px; }
.legend-bar { display: flex; gap: 10px; padding: 6px 12px; background: #1a1010; border-radius: 6px; margin-top: 10px; align-items: center; flex-wrap: wrap; }
.legend-item { display: flex; align-items: center; gap: 4px; font-size: 9px; }
.legend-dot { width: 9px; height: 9px; border-radius: 2px; }
.grid-2col { display: grid; grid-template-columns: 1.4fr 1fr; gap: 12px; }
.grid-3col { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 10px; }
"""

def pagina_html(navbar_html, conteudo, titulo="BR DRIVE"):
    return f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>{titulo} — BR DRIVE</title>
  <style>{CSS_GLOBAL}</style>
</head>
<body>
{navbar_html}
<div class="main">
{conteudo}
</div>
</body>
</html>"""

# ══════════════════════════════════════════════════════════════════════════════
# PÁGINAS
# ══════════════════════════════════════════════════════════════════════════════

def pg_home():
    paginas = [
        ("mural",           "📺","仪表板墙",      "Mural de Dashboards",      "4 painéis simultâneos · tela cheia · 实时多屏"),
        ("dashboard",       "📊","应用使用情况",  "Utilização do App",       "Visão geral · utilização · transportadoras"),
        ("pont_op",         "⏱","运营准时率",    "Pontualidade Operacional", "Saída e chegada · 准时出发与到达"),
        ("evol_semanal",    "📈","应用周度绩效",  "Desempenho App (Semanal)", "Comparativo semanal · 周对比"),
        ("app_mensal",      "📅","APP月度绩效",   "Performance Mensal APP",   "Evolução mensal · 月度趋势"),
        ("pont_semanal",    "📊","准时率周度绩效","Pontualidade (Semanal)",   "Análise no prazo · 准时分析"),
        ("pont_mensal",     "📆","准时率月度绩效","Pontualidade (Mensal)",    "Evolução mensal · 月度趋势"),
        ("infracoes",       "🚨","违章管理",      "Gestão de Infrações",      "App OFF · condutor · 按行程明细"),
        ("motoristas",      "👤","司机管理",      "Utilização APP Motoristas","KPI saída · chegada · geral"),
        ("pont_motoristas", "🚗","司机准时率",    "Pontualidade Motoristas",  "KPI saída · chegada · geral"),
        ("inf_pontualidade","🔴","准时率违规",    "Inf. de Pontualidade",     "Detalhe por ID · violações"),
    ]
    cards = ""
    for pid, icon, zh, pt, desc in paginas:
        cards += f"""
        <a href="{pid}.html" class="home-card">
          <div class="home-icon">{icon}</div>
          <div class="home-zh">{zh}</div>
          <div class="home-pt">{pt}</div>
          <div class="home-desc">{desc}</div>
          <div class="home-arrow">›</div>
        </a>"""

    css_home = """
    <style>
    body { background: #C0001A !important; }
    .main { background: #C0001A; }
    .home-grid { display: grid; grid-template-columns: repeat(3,1fr); gap: 10px; padding: 0 8px; }
    .home-card { background: rgba(255,255,255,0.95); border-radius: 12px; padding: 14px; display: flex; flex-direction: column; gap: 4px; text-decoration: none; transition: transform 0.15s, box-shadow 0.15s; }
    .home-card:hover { transform: translateY(-2px); box-shadow: 0 8px 24px rgba(0,0,0,0.3); }
    .home-icon { width: 34px; height: 34px; border-radius: 9px; background: #C0001A; display: flex; align-items: center; justify-content: center; font-size: 16px; }
    .home-zh { font-size: 12px; font-weight: 800; color: #1a0000; }
    .home-pt { font-size: 9px; font-weight: 600; color: #C0001A; }
    .home-desc { font-size: 8px; color: #C0001A; }
    .home-arrow { align-self: flex-end; color: #C0001A; font-size: 20px; font-weight: 700; margin-top: auto; }
    .home-header { display: flex; align-items: center; justify-content: space-between; padding: 16px 24px 12px; border-bottom: 1px solid rgba(255,255,255,0.2); margin-bottom: 10px; }
    .home-logo { background: #fff; border-radius: 10px; padding: 6px 10px; }
    .home-logo-txt { font-size: 22px; font-weight: 900; color: #C0001A; letter-spacing: -2px; }
    .home-logo-sub { font-size: 6px; font-weight: 800; color: #C0001A; letter-spacing: 2.5px; }
    .home-title { color: #fff; font-size: 20px; font-weight: 800; letter-spacing: 1px; }
    .home-sub { color: rgba(255,255,255,0.85); font-size: 9px; letter-spacing: 2px; }
    .home-live { display: flex; align-items: center; gap: 6px; }
    .live-dot { width: 8px; height: 8px; border-radius: 50%; background: #fff; }
    .live-txt { color: #fff; font-size: 10px; font-weight: 700; letter-spacing: 1.5px; }
    .select-lbl { color: #fff; font-size: 11px; letter-spacing: 1.5px; font-weight: 600; text-align: center; padding: 6px 0 8px; }
    .home-footer { margin-top: 12px; padding: 10px 24px; background: rgba(0,0,0,0.15); display: flex; justify-content: space-between; }
    .home-footer span { color: #fff; font-size: 9px; letter-spacing: 1px; }
    </style>"""

    conteudo = f"""{css_home}
    <div class="home-header">
      <div style="display:flex;align-items:center;gap:14px;">
        <div class="home-logo">
          <div class="home-logo-txt">J&T</div>
          <div class="home-logo-sub">EXPRESS</div>
        </div>
        <div>
          <div class="home-title">BR DRIVE</div>
          <div class="home-sub">PAINEL OPERACIONAL · 运营控制台</div>
        </div>
      </div>
      <div class="home-live">
        <div class="live-dot"></div>
        <span class="live-txt">AO VIVO / 实时</span>
      </div>
    </div>
    <div class="select-lbl">SELECIONE UMA PÁGINA / 请选择页面</div>
    <div class="home-grid">{cards}</div>
    <div class="home-footer">
      <span>J&T EXPRESS · BR DRIVE · 承运商监控</span>
      <span>2026</span>
    </div>"""
    return pagina_html("", conteudo, "Home")

def pg_utilizacao_app(df):
    # Dados compactos para o JavaScript (apenas colunas necessárias)
    cols = ['DATA', 'Transportador', 'TIPO DE OPERAÇÃO', 'Tempo Saida OFF', 'Tempo chegada OFF']
    df_js = df[cols].copy()
    df_js['DATA'] = df_js['DATA'].dt.strftime('%Y-%m-%d')
    df_js = df_js.fillna('')
    df_js.columns = ['d', 't', 'op', 's', 'c']  # chaves curtas para reduzir tamanho
    data_json = df_js.to_json(orient='records', force_ascii=False)

    transportadoras = sorted(df['Transportador'].dropna().unique().tolist())
    tipos = sorted(df['TIPO DE OPERAÇÃO'].dropna().unique().tolist())
    date_min = df['DATA'].min().strftime('%Y-%m-%d') if not df.empty else ''
    date_max = df['DATA'].max().strftime('%Y-%m-%d') if not df.empty else ''

    opts_trans = '<option value="Todos">Todos</option>' + ''.join(
        f'<option value="{t}">{t}</option>' for t in transportadoras)
    opts_tipo = '<option value="Todos">Todos</option>' + ''.join(
        f'<option value="{t}">{t}</option>' for t in tipos)

    sel_style = 'width:100%;background:#0e0808;border:1px solid rgba(192,57,43,0.3);color:#fff;border-radius:4px;padding:5px 6px;font-size:11px;'
    inp_style = 'width:100%;background:#0e0808;border:1px solid rgba(192,57,43,0.3);color:#fff;border-radius:4px;padding:5px 6px;font-size:11px;color-scheme:dark;'

    filtro_html = f"""
    <div style="width:210px;flex-shrink:0;">
      <div style="background:#1a1010;border-radius:8px;padding:14px;border:1px solid rgba(192,57,43,0.25);position:sticky;top:16px;">
        <div style="color:#C0392B;font-size:11px;font-weight:700;letter-spacing:1px;margin-bottom:12px;padding-bottom:8px;border-bottom:1px solid rgba(192,57,43,0.3);">FILTROS / 筛选</div>
        <div style="margin-bottom:10px;">
          <div style="color:#ddd;font-size:9px;letter-spacing:1px;margin-bottom:4px;">日期 | Data De</div>
          <input type="date" id="f-date-from" value="{date_min}" onchange="setFiltro('dateFrom',this.value)" style="{inp_style}margin-bottom:4px;">
          <div style="color:#ddd;font-size:9px;letter-spacing:1px;margin-bottom:4px;margin-top:6px;">日期 | Data Até</div>
          <input type="date" id="f-date-to" value="{date_max}" onchange="setFiltro('dateTo',this.value)" style="{inp_style}">
        </div>
        <div style="margin-bottom:10px;">
          <div style="color:#ddd;font-size:9px;letter-spacing:1px;margin-bottom:4px;">承运商 | Transportador</div>
          <select id="f-trans" onchange="setFiltro('trans',this.value)" style="{sel_style}">{opts_trans}</select>
        </div>
        <div style="margin-bottom:10px;">
          <div style="color:#ddd;font-size:9px;letter-spacing:1px;margin-bottom:4px;">运营类型 | Tipo Op.</div>
          <select id="f-tipo" onchange="setFiltro('tipo',this.value)" style="{sel_style}">{opts_tipo}</select>
        </div>
        <div style="margin-bottom:14px;">
          <div style="color:#ddd;font-size:9px;letter-spacing:1px;margin-bottom:4px;">状态 | Status APP</div>
          <select id="f-status" onchange="setFiltro('status',this.value)" style="{sel_style}">
            <option value="Todos">Todos</option>
            <option value="OK">Somente OK</option>
            <option value="OFF">Somente OFF</option>
          </select>
        </div>
        <button onclick="resetFiltros()" style="width:100%;background:rgba(192,57,43,0.2);border:1px solid rgba(192,57,43,0.4);color:#C0392B;border-radius:4px;padding:7px;font-size:10px;font-weight:700;cursor:pointer;letter-spacing:1px;">RESET / 重置</button>
      </div>
    </div>"""

    nav = navbar("dashboard")

    conteudo = """
    <div class="page-header">
      <div class="header-left">
        <div class="header-icon" style="background:#C0392B;">BR</div>
        <div>
          <div class="header-title">BR DRIVE — DASHBOARD OPERACIONAL / 运营仪表盘</div>
          <div class="header-sub">MONITORAMENTO DE TRANSPORTADORES · 承运商监控 · <span id="periodo-txt">—</span></div>
        </div>
      </div>
      <div class="header-badges">
        <div style="display:flex;align-items:center;gap:6px;">
          <div style="width:7px;height:7px;border-radius:50%;background:#27ae60;"></div>
          <span style="font-size:11px;font-weight:600;color:#27ae60;">AO VIVO / 实时</span>
        </div>
        <div class="header-badge"><div class="hb-label">MÊS / 月</div><div class="hb-value" id="kpi-mes">—</div></div>
      </div>
    </div>

    <div style="display:flex;gap:12px;align-items:flex-start;">
      <div style="flex:1;min-width:0;">

        <div class="kpi-grid kpi-grid-5" style="margin-bottom:10px;">
          <div class="kpi-card" style="border-top:3px solid #C0392B;">
            <div class="kpi-label">TOTAL DE IDs / 行程总数</div>
            <div class="kpi-value" id="kpi-total">—</div>
            <div class="kpi-sub">entregas no período</div>
          </div>
          <div class="kpi-card" style="border-top:3px solid #27ae60;">
            <div class="kpi-label">UTILIZAÇÃO APP SAÍDA / 出发使用率</div>
            <div class="kpi-value" id="kpi-util-s" style="color:#27ae60;">—</div>
            <div class="kpi-sub">app ativo na saída / 出发时激活</div>
            <div class="progress-bar" style="margin-top:6px;"><div id="kpi-bar-s" class="progress-fill" style="width:0%;"></div></div>
          </div>
          <div class="kpi-card" style="border-top:3px solid #27ae60;">
            <div class="kpi-label">UTILIZAÇÃO APP CHEGADA / 到达使用率</div>
            <div class="kpi-value" id="kpi-util-c" style="color:#27ae60;">—</div>
            <div class="kpi-sub">app ativo na chegada / 到达时激活</div>
            <div class="progress-bar" style="margin-top:6px;"><div id="kpi-bar-c" class="progress-fill" style="width:0%;"></div></div>
          </div>
          <div class="kpi-card" style="border-top:3px solid #e74c3c;">
            <div class="kpi-label">PIOR TRANSP. / 最差承运商</div>
            <div class="kpi-value" id="kpi-pior" style="font-size:16px;color:#e74c3c;">—</div>
            <div class="kpi-sub" id="kpi-pior-sub"></div>
            <div class="progress-bar" style="margin-top:6px;"><div id="kpi-bar-pior" class="progress-fill" style="width:0%;background:#e74c3c;"></div></div>
          </div>
          <div class="kpi-card" style="border-top:3px solid #27ae60;">
            <div class="kpi-label">MELHOR TRANSP. / 最佳承运商</div>
            <div class="kpi-value" id="kpi-melhor" style="font-size:16px;color:#27ae60;">—</div>
            <div class="kpi-sub" id="kpi-melhor-sub"></div>
            <div class="progress-bar" style="margin-top:6px;"><div id="kpi-bar-melhor" class="progress-fill" style="width:0%;background:#27ae60;"></div></div>
          </div>
        </div>

        <div id="alerta-bar" style="display:flex;gap:8px;padding:7px 12px;background:#1a1010;border-radius:6px;margin-bottom:10px;align-items:center;flex-wrap:wrap;min-height:34px;"></div>

        <div class="grid-2col">
          <div class="table-wrap">
            <div class="table-header" style="background:#C0392B;">
              <span>UTILIZAÇÃO DO APP POR TRANSPORTADOR / 按承运商应用使用率</span>
              <small>saída &amp; chegada / 出发与到达</small>
            </div>
            <div class="table-scroll">
              <table>
                <thead><tr style="background:#1a0808;border-bottom:2px solid rgba(192,57,43,0.4);">
                  <th style="text-align:left;color:#C0392B;">TRANSPORTADOR / 承运商</th>
                  <th style="text-align:center;color:#C0392B;">IDs / 行程数</th>
                  <th style="text-align:center;color:#C0392B;">APP OK SAÍDA</th>
                  <th style="text-align:center;color:#C0392B;">APP OK CHEGADA</th>
                </tr></thead>
                <tbody id="tbody-trans"></tbody>
              </table>
            </div>
          </div>
          <div class="table-wrap">
            <div class="table-header" style="background:#922b21;">
              <span>RANKING UTILIZAÇÃO SAÍDA / 出发使用率排名</span>
              <small>ordenado por volume / 按量排序</small>
            </div>
            <div style="padding:14px 16px;overflow-y:auto;max-height:calc(100vh - 340px);" id="barras-trans"></div>
            <div class="legend-bar" style="padding-top:10px;border-top:1px solid rgba(255,255,255,0.06);">
              <div class="legend-item"><div class="legend-dot" style="background:#27ae60;"></div>≥85% OK / 达标</div>
              <div class="legend-item"><div class="legend-dot" style="background:#e67e22;"></div>70–84% 注意</div>
              <div class="legend-item"><div class="legend-dot" style="background:#e74c3c;"></div>&lt;70% 危急</div>
            </div>
          </div>
        </div>

      </div>
      FILTRO_PLACEHOLDER
    </div>"""

    conteudo = conteudo.replace('FILTRO_PLACEHOLDER', filtro_html)

    # JavaScript — string concatenation para evitar conflito com f-string e chaves JS
    js = (
        '<script>\n'
        'const RAW = ' + data_json + ';\n'
        'const _DM="' + date_max + '",_DMi="' + date_min + '";\n'
        'const _lsM=localStorage.getItem("brdrive_data_max"),_lsF2=localStorage.getItem("brdrive_from"),_lsT2=localStorage.getItem("brdrive_to");\n'
        'if(_lsM!==_DM||(_lsF2&&_lsF2>_DM)||(_lsT2&&_lsT2<_DMi)){localStorage.removeItem("brdrive_from");localStorage.removeItem("brdrive_to");localStorage.setItem("brdrive_data_max",_DM);}\n'
        'const _lsF=localStorage.getItem("brdrive_from"),_lsT=localStorage.getItem("brdrive_to");\n'
        'let fDateFrom=_lsF||"' + date_min + '";\n'
        'let fDateTo=_lsT||"' + date_max + '";\n'
        'let fTrans = "Todos"; let fTipo = "Todos"; let fStatus = "Todos";\n'
        '\n'
        'function cor(v){if(v>=85)return"#27ae60";if(v>=70)return"#e67e22";return"#e74c3c";}\n'
        'function bgCor(v){if(v>=85)return"rgba(39,174,96,0.15)";if(v>=70)return"rgba(230,126,34,0.15)";return"rgba(231,76,60,0.15)";}\n'
        'function pct(a,b){return b>0?(a/b*100):0;}\n'
        '\n'
        'function getFiltrado(){\n'
        '  return RAW.filter(r=>{\n'
        '    if(fDateFrom && r.d < fDateFrom) return false;\n'
        '    if(fDateTo && r.d > fDateTo) return false;\n'
        '    if(fTrans!=="Todos" && r.t!==fTrans) return false;\n'
        '    if(fTipo!=="Todos" && r.op!==fTipo) return false;\n'
        '    if(fStatus==="OK" && (r.s!=="OK" || r.c!=="OK")) return false;\n'
        '    if(fStatus==="OFF" && r.s!=="OFF" && r.c!=="OFF") return false;\n'
        '    return true;\n'
        '  });\n'
        '}\n'
        '\n'
        'function render(){\n'
        '  const data = getFiltrado();\n'
        '  const total = data.length;\n'
        '  const okS = data.filter(r=>r.s==="OK").length;\n'
        '  const okC = data.filter(r=>r.c==="OK").length;\n'
        '  const utilS = pct(okS, total);\n'
        '  const utilC = pct(okC, total);\n'
        '\n'
        '  // Agrupa por transportador\n'
        '  const byT = {};\n'
        '  data.forEach(r=>{\n'
        '    if(!byT[r.t]) byT[r.t]={ids:0,okS:0,okC:0};\n'
        '    byT[r.t].ids++;\n'
        '    if(r.s==="OK") byT[r.t].okS++;\n'
        '    if(r.c==="OK") byT[r.t].okC++;\n'
        '  });\n'
        '  const sorted = Object.entries(byT).sort((a,b)=>b[1].ids-a[1].ids);\n'
        '\n'
        '  // Pior e melhor\n'
        '  let pior=null, melhor=null, pPior=999, pMelhor=-1;\n'
        '  sorted.forEach(([t,v])=>{\n'
        '    const p=pct(v.okS,v.ids);\n'
        '    if(p<pPior){pPior=p;pior=[t,v];}\n'
        '    if(p>pMelhor){pMelhor=p;melhor=[t,v];}\n'
        '  });\n'
        '\n'
        '  // KPIs\n'
        '  document.getElementById("kpi-total").textContent = total;\n'
        '  const elS = document.getElementById("kpi-util-s");\n'
        '  elS.textContent = utilS.toFixed(1)+"%"; elS.style.color = cor(utilS);\n'
        '  const bS = document.getElementById("kpi-bar-s");\n'
        '  bS.style.width = Math.min(utilS,100).toFixed(0)+"%"; bS.style.background = cor(utilS);\n'
        '  const elC = document.getElementById("kpi-util-c");\n'
        '  elC.textContent = utilC.toFixed(1)+"%"; elC.style.color = cor(utilC);\n'
        '  const bC = document.getElementById("kpi-bar-c");\n'
        '  bC.style.width = Math.min(utilC,100).toFixed(0)+"%"; bC.style.background = cor(utilC);\n'
        '  if(pior){\n'
        '    const p=pct(pior[1].okS,pior[1].ids);\n'
        '    const el=document.getElementById("kpi-pior"); el.textContent=pior[0]; el.style.color=cor(p);\n'
        '    document.getElementById("kpi-pior-sub").textContent="Saída: "+p.toFixed(1)+"% · Chegada: "+pct(pior[1].okC,pior[1].ids).toFixed(1)+"%";\n'
        '    const bp=document.getElementById("kpi-bar-pior"); bp.style.width=Math.min(p,100).toFixed(0)+"%"; bp.style.background=cor(p);\n'
        '  }\n'
        '  if(melhor){\n'
        '    const p=pct(melhor[1].okS,melhor[1].ids);\n'
        '    const el=document.getElementById("kpi-melhor"); el.textContent=melhor[0]; el.style.color=cor(p);\n'
        '    document.getElementById("kpi-melhor-sub").textContent="Saída: "+p.toFixed(1)+"% · Chegada: "+pct(melhor[1].okC,melhor[1].ids).toFixed(1)+"%";\n'
        '    const bm=document.getElementById("kpi-bar-melhor"); bm.style.width=Math.min(p,100).toFixed(0)+"%"; bm.style.background=cor(p);\n'
        '  }\n'
        '\n'
        '  // Alerta bar\n'
        '  const alertas = sorted.filter(([t,v])=>pct(v.okS,v.ids)<85);\n'
        '  let ah = "";\n'
        '  if(alertas.length>0){\n'
        '    ah = "<span style=\'color:#e67e22;font-size:10px;font-weight:700;\'>⚠ ATENÇÃO / 注意</span>"'
        '       +"<span style=\'color:#ddd;font-size:10px;margin:0 8px;\'>S. · C.</span>";\n'
        '    alertas.slice(0,6).forEach(([t,v])=>{\n'
        '      const ps=pct(v.okS,v.ids); const pc=pct(v.okC,v.ids); const c=cor(ps);\n'
        '      ah+=`<span style="background:rgba(231,76,60,0.15);border:1px solid ${c};color:${c};border-radius:4px;padding:2px 8px;font-size:10px;font-weight:700;">${t} S:${ps.toFixed(1)}% C:${pc.toFixed(1)}%</span>`;\n'
        '    });\n'
        '  } else if(total>0) {\n'
        '    ah = "<span style=\'color:#27ae60;font-size:10px;font-weight:600;\'>✓ Todos os transportadores dentro da meta / 所有承运商达标</span>";\n'
        '  }\n'
        '  document.getElementById("alerta-bar").innerHTML = ah;\n'
        '\n'
        '  // Tabela\n'
        '  let rows = "";\n'
        '  sorted.forEach(([t,v],i)=>{\n'
        '    const ps=pct(v.okS,v.ids); const pc=pct(v.okC,v.ids);\n'
        '    const bgr = i%2===0?"#1f0e0e":"#1a1010";\n'
        '    rows += `<tr style="background:${bgr};">'
        '<td style="color:#5dade2;font-weight:700;">${t}</td>'
        '<td style="color:#f1c40f;font-weight:700;text-align:center;">${v.ids}</td>'
        '<td><span style="background:${bgCor(ps)};color:${cor(ps)};padding:2px 10px;border-radius:20px;font-size:11px;font-weight:700;">${ps.toFixed(1)}%</span></td>'
        '<td><span style="background:${bgCor(pc)};color:${cor(pc)};padding:2px 10px;border-radius:20px;font-size:11px;font-weight:700;">${pc.toFixed(1)}%</span></td>'
        '</tr>`;\n'
        '  });\n'
        '  document.getElementById("tbody-trans").innerHTML = rows || '
        '  "<tr><td colspan=\'4\' style=\'text-align:center;color:#ddd;padding:20px;\'>Nenhum dado para os filtros selecionados</td></tr>";\n'
        '\n'
        '  // Barras ranking\n'
        '  let barras = "";\n'
        '  sorted.forEach(([t,v])=>{\n'
        '    const ps=pct(v.okS,v.ids); const c=cor(ps);\n'
        '    barras += `<div style="display:grid;grid-template-columns:100px 1fr 44px;align-items:center;gap:8px;margin-bottom:7px;">'
        '<div style="font-size:11px;color:#fff;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">${t}</div>'
        '<div style="height:16px;background:rgba(255,255,255,0.05);border-radius:3px;overflow:hidden;">'
        '<div style="height:100%;width:${Math.min(ps,100).toFixed(0)}%;background:${c};border-radius:3px;"></div></div>'
        '<div style="font-size:11px;font-weight:700;color:${c};text-align:right;">${ps.toFixed(0)}%</div>'
        '</div>`;\n'
        '  });\n'
        '  document.getElementById("barras-trans").innerHTML = barras;\n'
        '\n'
        '  // Período e mês\n'
        '  const dates = data.map(r=>r.d).filter(Boolean).sort();\n'
        '  document.getElementById("periodo-txt").textContent = dates.length>0 ? dates[0]+" → "+dates[dates.length-1] : "—";\n'
        '}\n'
        '\n'
        'function setFiltro(key,val){\n'
        '  if(key==="dateFrom"){fDateFrom=val;localStorage.setItem("brdrive_from",val);}\n'
        '  else if(key==="dateTo"){fDateTo=val;localStorage.setItem("brdrive_to",val);}\n'
        '  else if(key==="trans") fTrans=val;\n'
        '  else if(key==="tipo") fTipo=val;\n'
        '  else if(key==="status") fStatus=val;\n'
        '  render();\n'
        '}\n'
        '\n'
        'function resetFiltros(){\n'
        '  fDateFrom="' + date_min + '"; fDateTo="' + date_max + '";\n'
        '  fTrans="Todos"; fTipo="Todos"; fStatus="Todos";\n'
        '  document.getElementById("f-date-from").value=fDateFrom;\n'
        '  document.getElementById("f-date-to").value=fDateTo;\n'
        '  document.getElementById("f-trans").value="Todos";\n'
        '  document.getElementById("f-tipo").value="Todos";\n'
        '  document.getElementById("f-status").value="Todos";\n'
        '  render();\n'
        '}\n'
        '\n'
        'document.addEventListener("DOMContentLoaded",function(){\n'
        '  var fi=document.getElementById("f-date-from"),ti=document.getElementById("f-date-to");\n'
        '  if(fi)fi.value=fDateFrom;\n'
        '  if(ti)ti.value=fDateTo;\n'
        '  render();\n'
        '});\n'
        '</script>'
    )

    return pagina_html(nav, conteudo + js, "Utilização do App")

def pg_pontualidade_op(df):
    cols = ['DATA', 'Transportador', 'TIPO DE OPERAÇÃO', 'PONTUALIDADE SAÍDA', 'PONTUALIDADE CHEGADA']
    df_js = df[cols].copy()
    df_js['DATA'] = df_js['DATA'].dt.strftime('%Y-%m-%d')
    df_js = df_js.fillna('')
    df_js.columns = ['d', 't', 'op', 'ps', 'pc']
    data_json = df_js.to_json(orient='records', force_ascii=False)

    transportadoras = sorted(df['Transportador'].dropna().unique().tolist())
    tipos = ["Coleta Base", "Coleta PA", "Devolução", "Secundaria"]
    date_min = df['DATA'].min().strftime('%Y-%m-%d') if not df.empty else ''
    date_max = df['DATA'].max().strftime('%Y-%m-%d') if not df.empty else ''

    opts_trans = '<option value="Todos">Todos</option>' + ''.join(
        f'<option value="{t}">{t}</option>' for t in transportadoras)
    opts_tipo = '<option value="Todos">Todos</option>' + ''.join(
        f'<option value="{t}">{t}</option>' for t in tipos)

    sel_style = 'width:100%;background:#0e0808;border:1px solid rgba(26,82,118,0.4);color:#fff;border-radius:4px;padding:5px 6px;font-size:11px;'
    inp_style = 'width:100%;background:#0e0808;border:1px solid rgba(26,82,118,0.4);color:#fff;border-radius:4px;padding:5px 6px;font-size:11px;color-scheme:dark;'

    filtro_html = f"""
    <div style="width:210px;flex-shrink:0;">
      <div style="background:#1a1010;border-radius:8px;padding:14px;border:1px solid rgba(26,82,118,0.3);position:sticky;top:16px;">
        <div style="color:#2980b9;font-size:11px;font-weight:700;letter-spacing:1px;margin-bottom:12px;padding-bottom:8px;border-bottom:1px solid rgba(26,82,118,0.3);">FILTROS / 筛选</div>
        <div style="margin-bottom:10px;">
          <div style="color:#ddd;font-size:9px;letter-spacing:1px;margin-bottom:4px;">日期 | Data De</div>
          <input type="date" id="f-date-from" value="{date_min}" onchange="setFiltro('dateFrom',this.value)" style="{inp_style}margin-bottom:4px;">
          <div style="color:#ddd;font-size:9px;letter-spacing:1px;margin-bottom:4px;margin-top:6px;">日期 | Data Até</div>
          <input type="date" id="f-date-to" value="{date_max}" onchange="setFiltro('dateTo',this.value)" style="{inp_style}">
        </div>
        <div style="margin-bottom:10px;">
          <div style="color:#ddd;font-size:9px;letter-spacing:1px;margin-bottom:4px;">承运商 | Transportador</div>
          <select id="f-trans" onchange="setFiltro('trans',this.value)" style="{sel_style}">{opts_trans}</select>
        </div>
        <div style="margin-bottom:14px;">
          <div style="color:#ddd;font-size:9px;letter-spacing:1px;margin-bottom:4px;">运营类型 | Tipo Op.</div>
          <select id="f-tipo" onchange="setFiltro('tipo',this.value)" style="{sel_style}">{opts_tipo}</select>
        </div>
        <button onclick="resetFiltros()" style="width:100%;background:rgba(26,82,118,0.2);border:1px solid rgba(26,82,118,0.4);color:#2980b9;border-radius:4px;padding:7px;font-size:10px;font-weight:700;cursor:pointer;letter-spacing:1px;">RESET / 重置</button>
      </div>
    </div>"""

    nav = navbar("pont_op")
    conteudo = """
    <div class="page-header">
      <div class="header-left">
        <div class="header-icon" style="background:#1A5276;">⏱</div>
        <div>
          <div class="header-title">PONTUALIDADE OPERACIONAL / 运营准时率</div>
          <div class="header-sub">SAÍDA &amp; CHEGADA · 出发与到达准时率 · <span id="periodo-txt">—</span></div>
        </div>
      </div>
      <div class="header-badges">
        <div style="display:flex;align-items:center;gap:6px;">
          <div style="width:7px;height:7px;border-radius:50%;background:#27ae60;"></div>
          <span style="font-size:11px;font-weight:600;color:#27ae60;">AO VIVO / 实时</span>
        </div>
        <div class="header-badge"><div class="hb-label">META / 目标</div><div class="hb-value" style="color:#27ae60;">≥85%</div></div>
      </div>
    </div>

    <div style="display:flex;gap:12px;align-items:flex-start;">
      <div style="flex:1;min-width:0;">

        <div class="kpi-grid kpi-grid-5" style="margin-bottom:10px;">
          <div class="kpi-card" style="border-top:3px solid #1A5276;">
            <div class="kpi-label">TOTAL IDs / 行程总数</div>
            <div class="kpi-value" id="kpi-total">—</div>
            <div class="kpi-sub">com dados de pontualidade</div>
          </div>
          <div class="kpi-card" style="border-top:3px solid #27ae60;">
            <div class="kpi-label">NO PRAZO SAÍDA / 准时出发率</div>
            <div class="kpi-value" id="kpi-pont-s" style="color:#27ae60;">—</div>
            <div class="kpi-sub" id="kpi-pont-s-sub">saídas no prazo</div>
            <div class="progress-bar" style="margin-top:6px;"><div id="kpi-bar-s" class="progress-fill" style="width:0%;"></div></div>
          </div>
          <div class="kpi-card" style="border-top:3px solid #27ae60;">
            <div class="kpi-label">NO PRAZO CHEGADA / 准时到达率</div>
            <div class="kpi-value" id="kpi-pont-c" style="color:#27ae60;">—</div>
            <div class="kpi-sub" id="kpi-pont-c-sub">chegadas no prazo</div>
            <div class="progress-bar" style="margin-top:6px;"><div id="kpi-bar-c" class="progress-fill" style="width:0%;"></div></div>
          </div>
          <div class="kpi-card" style="border-top:3px solid #e74c3c;">
            <div class="kpi-label">PIOR TRANSP. SAÍDA / 最差出发承运商</div>
            <div class="kpi-value" id="kpi-pior" style="font-size:16px;color:#e74c3c;">—</div>
            <div class="kpi-sub" id="kpi-pior-sub"></div>
            <div class="progress-bar" style="margin-top:6px;"><div id="kpi-bar-pior" class="progress-fill" style="width:0%;background:#e74c3c;"></div></div>
          </div>
          <div class="kpi-card" style="border-top:3px solid #27ae60;">
            <div class="kpi-label">MELHOR TRANSP. SAÍDA / 最佳出发承运商</div>
            <div class="kpi-value" id="kpi-melhor" style="font-size:16px;color:#27ae60;">—</div>
            <div class="kpi-sub" id="kpi-melhor-sub"></div>
            <div class="progress-bar" style="margin-top:6px;"><div id="kpi-bar-melhor" class="progress-fill" style="width:0%;background:#27ae60;"></div></div>
          </div>
        </div>

        <div id="alerta-bar" style="display:flex;gap:8px;padding:7px 12px;background:#1a1010;border-radius:6px;margin-bottom:10px;align-items:center;flex-wrap:wrap;min-height:34px;"></div>

        <div class="grid-2col">
          <div class="table-wrap">
            <div class="table-header" style="background:#1A5276;">
              <span>PONTUALIDADE POR TRANSPORTADOR / 按承运商准时率</span>
              <small>saída &amp; chegada / 出发与到达</small>
            </div>
            <div class="table-scroll">
              <table>
                <thead><tr style="background:#0d2137;border-bottom:2px solid rgba(26,82,118,0.6);">
                  <th style="text-align:left;color:#5dade2;">TRANSPORTADOR / 承运商</th>
                  <th style="text-align:center;color:#5dade2;">IDs</th>
                  <th style="text-align:center;color:#5dade2;">NO PRAZO SAÍDA</th>
                  <th style="text-align:center;color:#5dade2;">NO PRAZO CHEGADA</th>
                  <th style="text-align:center;color:#5dade2;">ATRASADOS</th>
                </tr></thead>
                <tbody id="tbody-pont"></tbody>
              </table>
            </div>
          </div>
          <div class="table-wrap">
            <div class="table-header" style="background:#154360;">
              <span>RANKING PONTUALIDADE SAÍDA / 出发准时率排名</span>
              <small>ordenado por volume / 按量排序</small>
            </div>
            <div style="padding:14px 16px;overflow-y:auto;max-height:calc(100vh - 340px);" id="barras-pont"></div>
            <div class="legend-bar" style="padding-top:10px;border-top:1px solid rgba(255,255,255,0.06);">
              <div class="legend-item"><div class="legend-dot" style="background:#27ae60;"></div>≥85% OK / 达标</div>
              <div class="legend-item"><div class="legend-dot" style="background:#e67e22;"></div>70–84% 注意</div>
              <div class="legend-item"><div class="legend-dot" style="background:#e74c3c;"></div>&lt;70% 危急</div>
            </div>
          </div>
        </div>

      </div>
      FILTRO_PLACEHOLDER
    </div>"""

    conteudo = conteudo.replace('FILTRO_PLACEHOLDER', filtro_html)

    js = (
        '<script>\n'
        'const RAW = ' + data_json + ';\n'
        'const _DM="' + date_max + '",_DMi="' + date_min + '";\n'
        'const _lsM=localStorage.getItem("brdrive_data_max"),_lsF2=localStorage.getItem("brdrive_from"),_lsT2=localStorage.getItem("brdrive_to");\n'
        'if(_lsM!==_DM||(_lsF2&&_lsF2>_DM)||(_lsT2&&_lsT2<_DMi)){localStorage.removeItem("brdrive_from");localStorage.removeItem("brdrive_to");localStorage.setItem("brdrive_data_max",_DM);}\n'
        'const _lsF=localStorage.getItem("brdrive_from"),_lsT=localStorage.getItem("brdrive_to");\n'
        'let fDateFrom=_lsF||"' + date_min + '";\n'
        'let fDateTo=_lsT||"' + date_max + '";\n'
        'let fTrans = "Todos"; let fTipo = "Todos";\n'
        '\n'
        'function cor(v){if(v>=85)return"#27ae60";if(v>=70)return"#e67e22";return"#e74c3c";}\n'
        'function bgCor(v){if(v>=85)return"rgba(39,174,96,0.15)";if(v>=70)return"rgba(230,126,34,0.15)";return"rgba(231,76,60,0.15)";}\n'
        'function pct(a,b){return b>0?(a/b*100):0;}\n'
        '\n'
        'function getFiltrado(){\n'
        '  return RAW.filter(r=>{\n'
        '    if(fDateFrom && r.d < fDateFrom) return false;\n'
        '    if(fDateTo && r.d > fDateTo) return false;\n'
        '    if(fTrans!=="Todos" && r.t!==fTrans) return false;\n'
        '    if(fTipo!=="Todos" && r.op!==fTipo) return false;\n'
        '    return true;\n'
        '  });\n'
        '}\n'
        '\n'
        'function render(){\n'
        '  const data = getFiltrado();\n'
        '  // Só conta registros com valor de pontualidade\n'
        '  const comS = data.filter(r=>r.ps!=="");\n'
        '  const comC = data.filter(r=>r.pc!=="");\n'
        '  const okS = comS.filter(r=>r.ps==="No prazo").length;\n'
        '  const okC = comC.filter(r=>r.pc==="No prazo").length;\n'
        '  const pontS = pct(okS, comS.length);\n'
        '  const pontC = pct(okC, comC.length);\n'
        '\n'
        '  // Agrupa por transportador\n'
        '  const byT = {};\n'
        '  data.forEach(r=>{\n'
        '    if(!byT[r.t]) byT[r.t]={ids:0,comS:0,okS:0,comC:0,okC:0,atr:0};\n'
        '    byT[r.t].ids++;\n'
        '    if(r.ps!==""){byT[r.t].comS++;if(r.ps==="No prazo")byT[r.t].okS++;}\n'
        '    if(r.pc!==""){byT[r.t].comC++;if(r.pc==="No prazo")byT[r.t].okC++;}\n'
        '    if(r.ps==="Atrasado"||r.ps==="Fora do prazo"||r.pc==="Atrasado"||r.pc==="Fora do prazo")byT[r.t].atr++;\n'
        '  });\n'
        '  const sorted = Object.entries(byT).sort((a,b)=>b[1].ids-a[1].ids);\n'
        '\n'
        '  // Pior e melhor (por pontualidade saída)\n'
        '  let pior=null,melhor=null,pPior=999,pMelhor=-1;\n'
        '  sorted.forEach(([t,v])=>{\n'
        '    const p=pct(v.okS,v.comS);\n'
        '    if(v.comS>0&&p<pPior){pPior=p;pior=[t,v];}\n'
        '    if(v.comS>0&&p>pMelhor){pMelhor=p;melhor=[t,v];}\n'
        '  });\n'
        '\n'
        '  // KPIs\n'
        '  document.getElementById("kpi-total").textContent = data.length;\n'
        '  const elS=document.getElementById("kpi-pont-s");\n'
        '  elS.textContent=pontS.toFixed(1)+"%"; elS.style.color=cor(pontS);\n'
        '  document.getElementById("kpi-pont-s-sub").textContent=okS+" de "+comS.length+" saídas no prazo";\n'
        '  const bS=document.getElementById("kpi-bar-s");\n'
        '  bS.style.width=Math.min(pontS,100).toFixed(0)+"%"; bS.style.background=cor(pontS);\n'
        '  const elC=document.getElementById("kpi-pont-c");\n'
        '  elC.textContent=pontC.toFixed(1)+"%"; elC.style.color=cor(pontC);\n'
        '  document.getElementById("kpi-pont-c-sub").textContent=okC+" de "+comC.length+" chegadas no prazo";\n'
        '  const bC=document.getElementById("kpi-bar-c");\n'
        '  bC.style.width=Math.min(pontC,100).toFixed(0)+"%"; bC.style.background=cor(pontC);\n'
        '  if(pior){\n'
        '    const p=pct(pior[1].okS,pior[1].comS);\n'
        '    const el=document.getElementById("kpi-pior"); el.textContent=pior[0]; el.style.color=cor(p);\n'
        '    document.getElementById("kpi-pior-sub").textContent="Saída: "+p.toFixed(1)+"%";\n'
        '    const bp=document.getElementById("kpi-bar-pior"); bp.style.width=Math.min(p,100).toFixed(0)+"%"; bp.style.background=cor(p);\n'
        '  }\n'
        '  if(melhor){\n'
        '    const p=pct(melhor[1].okS,melhor[1].comS);\n'
        '    const el=document.getElementById("kpi-melhor"); el.textContent=melhor[0]; el.style.color=cor(p);\n'
        '    document.getElementById("kpi-melhor-sub").textContent="Saída: "+p.toFixed(1)+"%";\n'
        '    const bm=document.getElementById("kpi-bar-melhor"); bm.style.width=Math.min(p,100).toFixed(0)+"%"; bm.style.background=cor(p);\n'
        '  }\n'
        '\n'
        '  // Alerta\n'
        '  const alertas=sorted.filter(([t,v])=>v.comS>0&&pct(v.okS,v.comS)<85);\n'
        '  let ah="";\n'
        '  if(alertas.length>0){\n'
        '    ah="<span style=\'color:#e67e22;font-size:10px;font-weight:700;\'>⚠ ATENÇÃO / 注意</span><span style=\'color:#ddd;font-size:10px;margin:0 8px;\'>Saída abaixo da meta:</span>";\n'
        '    alertas.slice(0,6).forEach(([t,v])=>{\n'
        '      const p=pct(v.okS,v.comS); const c=cor(p);\n'
        '      ah+=`<span style="background:rgba(231,76,60,0.15);border:1px solid ${c};color:${c};border-radius:4px;padding:2px 8px;font-size:10px;font-weight:700;">${t} ${p.toFixed(1)}%</span>`;\n'
        '    });\n'
        '  } else if(data.length>0){\n'
        '    ah="<span style=\'color:#27ae60;font-size:10px;font-weight:600;\'>✓ Todos os transportadores dentro da meta / 所有承运商达标</span>";\n'
        '  }\n'
        '  document.getElementById("alerta-bar").innerHTML=ah;\n'
        '\n'
        '  // Tabela\n'
        '  let rows="";\n'
        '  sorted.forEach(([t,v],i)=>{\n'
        '    const ps=pct(v.okS,v.comS); const pc=pct(v.okC,v.comC);\n'
        '    const bgr=i%2===0?"#1f0e0e":"#1a1010";\n'
        '    rows+=`<tr style="background:${bgr};">'
        '<td style="color:#5dade2;font-weight:700;">${t}</td>'
        '<td style="color:#f1c40f;font-weight:700;text-align:center;">${v.ids}</td>'
        '<td><span style="background:${bgCor(ps)};color:${cor(ps)};padding:2px 10px;border-radius:20px;font-size:11px;font-weight:700;">${v.comS>0?ps.toFixed(1)+"%":"—"}</span></td>'
        '<td><span style="background:${bgCor(pc)};color:${cor(pc)};padding:2px 10px;border-radius:20px;font-size:11px;font-weight:700;">${v.comC>0?pc.toFixed(1)+"%":"—"}</span></td>'
        '<td style="text-align:center;color:#e74c3c;font-weight:700;">${v.atr}</td>'
        '</tr>`;\n'
        '  });\n'
        '  document.getElementById("tbody-pont").innerHTML=rows||'
        '"<tr><td colspan=\'5\' style=\'text-align:center;color:#ddd;padding:20px;\'>Nenhum dado para os filtros selecionados</td></tr>";\n'
        '\n'
        '  // Barras\n'
        '  let barras="";\n'
        '  sorted.forEach(([t,v])=>{\n'
        '    const ps=pct(v.okS,v.comS); const c=cor(ps);\n'
        '    barras+=`<div style="display:grid;grid-template-columns:100px 1fr 44px;align-items:center;gap:8px;margin-bottom:7px;">'
        '<div style="font-size:11px;color:#fff;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">${t}</div>'
        '<div style="height:16px;background:rgba(255,255,255,0.05);border-radius:3px;overflow:hidden;">'
        '<div style="height:100%;width:${v.comS>0?Math.min(ps,100).toFixed(0):0}%;background:${c};border-radius:3px;"></div></div>'
        '<div style="font-size:11px;font-weight:700;color:${c};text-align:right;">${v.comS>0?ps.toFixed(0)+"%":"—"}</div>'
        '</div>`;\n'
        '  });\n'
        '  document.getElementById("barras-pont").innerHTML=barras;\n'
        '\n'
        '  // Período\n'
        '  const dates=data.map(r=>r.d).filter(Boolean).sort();\n'
        '  document.getElementById("periodo-txt").textContent=dates.length>0?dates[0]+" → "+dates[dates.length-1]:"—";\n'
        '}\n'
        '\n'
        'function setFiltro(key,val){\n'
        '  if(key==="dateFrom"){fDateFrom=val;localStorage.setItem("brdrive_from",val);}\n'
        '  else if(key==="dateTo"){fDateTo=val;localStorage.setItem("brdrive_to",val);}\n'
        '  else if(key==="trans") fTrans=val;\n'
        '  else if(key==="tipo") fTipo=val;\n'
        '  render();\n'
        '}\n'
        'function resetFiltros(){\n'
        '  fDateFrom="' + date_min + '"; fDateTo="' + date_max + '";\n'
        '  fTrans="Todos"; fTipo="Todos";\n'
        '  document.getElementById("f-date-from").value=fDateFrom;\n'
        '  document.getElementById("f-date-to").value=fDateTo;\n'
        '  document.getElementById("f-trans").value="Todos";\n'
        '  document.getElementById("f-tipo").value="Todos";\n'
        '  render();\n'
        '}\n'
        'document.addEventListener("DOMContentLoaded",function(){\n'
        '  var fi=document.getElementById("f-date-from"),ti=document.getElementById("f-date-to");\n'
        '  if(fi)fi.value=fDateFrom;\n'
        '  if(ti)ti.value=fDateTo;\n'
        '  render();\n'
        '});\n'
        '</script>'
    )

    return pagina_html(nav, conteudo + js, "Pontualidade Operacional")

def pg_infracoes(df):
    base_cols = ['DATA', 'MÊS', 'Transportador', 'CONDUTOR', 'TIPO DE OPERAÇÃO',
                 'Tempo Saida OFF', 'Tempo chegada OFF']
    opt_cols = ['Semana Nome Nova', 'Número do ID']
    cols_needed = base_cols + [c for c in opt_cols if c in df.columns]

    df_js = df[[c for c in cols_needed]].copy()
    df_js['DATA'] = df_js['DATA'].dt.strftime('%Y-%m-%d')
    df_js = df_js.fillna('')

    rename_map = {
        'DATA': 'd', 'MÊS': 'm', 'Semana Nome Nova': 'sem',
        'Transportador': 't', 'CONDUTOR': 'c', 'Número do ID': 'nid',
        'TIPO DE OPERAÇÃO': 'op', 'Tempo Saida OFF': 's', 'Tempo chegada OFF': 'ch'
    }
    df_js = df_js.rename(columns={k: v for k, v in rename_map.items() if k in df_js.columns})
    data_json = df_js.to_json(orient='records', force_ascii=False)

    transportadoras = sorted(df['Transportador'].dropna().unique().tolist())
    tipos = ["Coleta Base", "Coleta PA", "Devolução", "Secundaria"]
    date_min = df['DATA'].min().strftime('%Y-%m-%d') if not df.empty else ''
    date_max = df['DATA'].max().strftime('%Y-%m-%d') if not df.empty else ''

    opts_trans = '<option value="Todos">Todos</option>' + ''.join(
        f'<option value="{t}">{t}</option>' for t in transportadoras)
    opts_tipo = '<option value="Todos">Todos</option>' + ''.join(
        f'<option value="{t}">{t}</option>' for t in tipos)

    sel_style = 'width:100%;background:#0e0808;border:1px solid rgba(192,57,43,0.4);color:#fff;border-radius:4px;padding:5px 6px;font-size:11px;'
    inp_style = 'width:100%;background:#0e0808;border:1px solid rgba(192,57,43,0.4);color:#fff;border-radius:4px;padding:5px 6px;font-size:11px;color-scheme:dark;'

    filtro_html = f"""
    <div style="width:210px;flex-shrink:0;">
      <div style="background:#1a1010;border-radius:8px;padding:14px;border:1px solid rgba(192,57,43,0.3);position:sticky;top:16px;">
        <div style="color:#C0392B;font-size:11px;font-weight:700;letter-spacing:1px;margin-bottom:12px;padding-bottom:8px;border-bottom:1px solid rgba(192,57,43,0.3);">FILTROS / 筛选</div>
        <div style="margin-bottom:10px;">
          <div style="color:#ddd;font-size:9px;letter-spacing:1px;margin-bottom:4px;">日期从 | Data De</div>
          <input type="date" id="f-date-from" value="{date_max}" onchange="setFiltro('dateFrom',this.value)" style="{inp_style}margin-bottom:4px;">
          <div style="color:#ddd;font-size:9px;letter-spacing:1px;margin-bottom:4px;margin-top:6px;">日期到 | Data Até</div>
          <input type="date" id="f-date-to" value="{date_max}" onchange="setFiltro('dateTo',this.value)" style="{inp_style}">
        </div>
        <div style="margin-bottom:10px;">
          <div style="color:#ddd;font-size:9px;letter-spacing:1px;margin-bottom:4px;">承运商 | Transportadora</div>
          <select id="f-trans" onchange="setFiltro('trans',this.value)" style="{sel_style}">{opts_trans}</select>
        </div>
        <div style="margin-bottom:10px;">
          <div style="color:#ddd;font-size:9px;letter-spacing:1px;margin-bottom:4px;">运营类型 | Tipo Op.</div>
          <select id="f-tipo" onchange="setFiltro('tipo',this.value)" style="{sel_style}">{opts_tipo}</select>
        </div>
        <div style="margin-bottom:14px;">
          <div style="color:#ddd;font-size:9px;letter-spacing:1px;margin-bottom:4px;">等级 | Nível</div>
          <select id="f-nivel" onchange="setFiltro('nivel',this.value)" style="{sel_style}">
            <option value="Todos">Todos</option>
            <option value="CRITICO">CRÍTICO</option>
            <option value="OK">OK</option>
          </select>
        </div>
        <button onclick="resetFiltros()" style="width:100%;background:rgba(192,57,43,0.2);border:1px solid rgba(192,57,43,0.4);color:#C0392B;border-radius:4px;padding:7px;font-size:10px;font-weight:700;cursor:pointer;letter-spacing:1px;">RESET / 重置</button>
      </div>
    </div>"""

    nav = navbar("infracoes")
    conteudo = """
    <div class="page-header">
      <div class="header-left">
        <div class="header-icon" style="background:#C0392B;">🚨</div>
        <div>
          <div class="header-title">GESTÃO DE INFRAÇÕES / 违规管理</div>
          <div class="header-sub">DETALHE POR ID · 行程违规明细 · <span id="periodo-txt">—</span></div>
        </div>
      </div>
      <div class="header-badges">
        <div class="header-badge"><div class="hb-label">MÊS/月</div><div class="hb-value" id="hb-mes">—</div></div>
        <div class="header-badge"><div class="hb-label">SEMANA/周</div><div class="hb-value" id="hb-sem">—</div></div>
        <div class="header-badge"><div class="hb-label">PERÍODO/期间</div><div class="hb-value" id="hb-periodo">—</div></div>
        <div class="header-badge"><div class="hb-label">NÍVEL/层级</div><div class="hb-value" id="hb-nivel">TODOS</div></div>
      </div>
    </div>

    <div style="display:flex;gap:16px;align-items:flex-start;">
      <div style="flex:1;min-width:0;">

        <div class="kpi-grid kpi-grid-5" style="margin-bottom:12px;">
          <div class="kpi-card" style="border-top:3px solid #C0392B;">
            <div class="kpi-label">TOTAL IDs / 总行程数</div>
            <div class="kpi-value" id="kpi-total">—</div>
            <div class="kpi-sub" id="kpi-total-sub">IDs no período</div>
          </div>
          <div class="kpi-card" style="border-top:3px solid #e74c3c;">
            <div class="kpi-label">CRÍTICO / 严重违规</div>
            <div class="kpi-value" id="kpi-critico" style="color:#e74c3c;">—</div>
            <div class="kpi-sub">1 ou mais OFF / 离线违规</div>
          </div>
          <div class="kpi-card" style="border-top:3px solid #e74c3c;">
            <div class="kpi-label">OFF SAÍDA / 出发离线</div>
            <div class="kpi-value" id="kpi-off-s" style="color:#e74c3c;">—</div>
            <div class="kpi-sub">sem app na saída</div>
          </div>
          <div class="kpi-card" style="border-top:3px solid #e67e22;">
            <div class="kpi-label">OFF CHEGADA / 到达离线</div>
            <div class="kpi-value" id="kpi-off-c" style="color:#e67e22;">—</div>
            <div class="kpi-sub">sem app na chegada</div>
          </div>
        </div>

        <div class="table-wrap">
          <div class="table-header" style="background:#C0392B;">
            <span>DETALHE DE INFRAÇÕES POR CONDUTOR / 按司机违规明细</span>
            <small id="table-count">scroll para ver todos · 滚动查看全部</small>
          </div>
          <div class="table-scroll">
            <table>
              <thead><tr style="background:#1a0808;border-bottom:2px solid rgba(192,57,43,0.4);">
                <th style="text-align:left;color:#C0392B;">CONDUTOR/司机</th>
                <th style="text-align:left;color:#C0392B;">TRANSP./承运商</th>
                <th style="text-align:left;color:#C0392B;">NÚMERO DO ID</th>
                <th style="text-align:left;color:#C0392B;">TIPO OP.</th>
                <th style="text-align:center;color:#C0392B;">OFF SAÍDA</th>
                <th style="text-align:center;color:#C0392B;">OFF CHEGADA</th>
                <th style="text-align:center;color:#C0392B;">NÍVEL</th>
              </tr></thead>
              <tbody id="tbody-inf"></tbody>
            </table>
          </div>
        </div>

        <div class="legend-bar">
          <div class="legend-item" style="color:#e74c3c;"><div class="legend-dot" style="background:#e74c3c;"></div>CRÍTICO = qualquer OFF / 任意离线</div>
          <div class="legend-item" style="color:#27ae60;"><div class="legend-dot" style="background:#27ae60;"></div>OK = app ativo / 应用正常</div>
        </div>

      </div>
      FILTRO_PLACEHOLDER
    </div>"""

    conteudo = conteudo.replace('FILTRO_PLACEHOLDER', filtro_html)

    js = (
        '<script>\n'
        'const RAW = ' + data_json + ';\n'
        'const _DM="' + date_max + '",_DMi="' + date_min + '";\n'
        'const _lsM=localStorage.getItem("brdrive_data_max"),_lsF2=localStorage.getItem("brdrive_from"),_lsT2=localStorage.getItem("brdrive_to");\n'
        'if(_lsM!==_DM||(_lsF2&&_lsF2>_DM)||(_lsT2&&_lsT2<_DMi)){localStorage.removeItem("brdrive_from");localStorage.removeItem("brdrive_to");localStorage.setItem("brdrive_data_max",_DM);}\n'
        'const _lsF=localStorage.getItem("brdrive_from"),_lsT=localStorage.getItem("brdrive_to");\n'
        'let fDateFrom=_lsF||"' + date_max + '";\n'
        'let fDateTo=_lsT||"' + date_max + '";\n'
        'let fTrans = "Todos"; let fTipo = "Todos"; let fNivel = "Todos";\n'
        '\n'
        'function nivel(r){\n'
        '  if(r.s==="OFF"||r.ch==="OFF") return "CRITICO";\n'
        '  return "OK";\n'
        '}\n'
        '\n'
        'function getFiltrado(){\n'
        '  return RAW.filter(r=>{\n'
        '    if(fDateFrom && r.d < fDateFrom) return false;\n'
        '    if(fDateTo && r.d > fDateTo) return false;\n'
        '    if(fTrans!=="Todos" && r.t!==fTrans) return false;\n'
        '    if(fTipo!=="Todos" && r.op!==fTipo) return false;\n'
        '    if(fNivel!=="Todos" && nivel(r)!==fNivel) return false;\n'
        '    return true;\n'
        '  });\n'
        '}\n'
        '\n'
        'function render(){\n'
        '  const data = getFiltrado();\n'
        '  const total = data.length;\n'
        '  const critico = data.filter(r=>nivel(r)==="CRITICO").length;\n'
        '  const offS = data.filter(r=>r.s==="OFF").length;\n'
        '  const offC = data.filter(r=>r.ch==="OFF").length;\n'
        '\n'
        '  document.getElementById("kpi-total").textContent = total;\n'
        '  document.getElementById("kpi-total-sub").textContent = total>0?total+" IDs no período":"sem dados";\n'
        '  document.getElementById("kpi-critico").textContent = critico;\n'
        '  document.getElementById("kpi-off-s").textContent = offS;\n'
        '  document.getElementById("kpi-off-c").textContent = offC;\n'
        '\n'
        '  const dates = data.map(r=>r.d).filter(Boolean).sort();\n'
        '  const periodo = dates.length>0 ? dates[0]+" → "+dates[dates.length-1] : "—";\n'
        '  document.getElementById("periodo-txt").textContent = periodo;\n'
        '  document.getElementById("hb-periodo").textContent = periodo;\n'
        '  const meses = [...new Set(data.map(r=>r.m).filter(Boolean))];\n'
        '  document.getElementById("hb-mes").textContent = meses.length>0 ? meses.join("/") : "—";\n'
        '  const sems = [...new Set(data.map(r=>(r.sem||"")).filter(Boolean))];\n'
        '  document.getElementById("hb-sem").textContent = sems.length>0 ? sems[sems.length-1] : "—";\n'
        '  const nivelLabel = fNivel==="Todos"?"TODOS":fNivel==="CRITICO"?"CRÍTICO":"OK";\n'
        '  document.getElementById("hb-nivel").textContent = nivelLabel;\n'
        '\n'
        '  const MAX_ROWS = 500;\n'
        '  const sorted = data.slice().sort((a,b)=>a.c<b.c?-1:a.c>b.c?1:0);\n'
        '  let rows = "";\n'
        '  sorted.slice(0,MAX_ROWS).forEach((r,i)=>{\n'
        '    const nv = nivel(r);\n'
        '    const bgr = i%2===0?"#1f0e0e":"#1a1010";\n'
        '    const nc = nv==="CRITICO"?"#e74c3c":"#27ae60";\n'
        '    const nb = nv==="CRITICO"?"rgba(231,76,60,0.2)":"rgba(39,174,96,0.2)";\n'
        '    const nvLabel = nv==="CRITICO"?"CRÍTICO":"OK";\n'
        '    const cs = r.s==="OFF"?"#e74c3c":"#27ae60";\n'
        '    const bgs = r.s==="OFF"?"rgba(231,76,60,0.2)":"rgba(39,174,96,0.2)";\n'
        '    const cc = r.ch==="OFF"?"#e74c3c":"#27ae60";\n'
        '    const bgc = r.ch==="OFF"?"rgba(231,76,60,0.2)":"rgba(39,174,96,0.2)";\n'
        '    rows += `<tr style="background:${bgr};">'
        '<td style="color:#5dade2;font-weight:600;">${r.c||"—"}</td>'
        '<td><span class="trans-badge">${r.t||"—"}</span></td>'
        '<td style="color:#f1c40f;font-weight:700;font-size:10px;">${r.nid||"—"}</td>'
        '<td style="color:#fff;font-size:10px;">${r.op||"—"}</td>'
        '<td style="text-align:center;"><span class="badge" style="background:${bgs};color:${cs};">${r.s||"—"}</span></td>'
        '<td style="text-align:center;"><span class="badge" style="background:${bgc};color:${cc};">${r.ch||"—"}</span></td>'
        '<td style="text-align:center;"><span class="badge" style="background:${nb};color:${nc};">${nvLabel}</span></td>'
        '</tr>`;\n'
        '  });\n'
        '  if(sorted.length>MAX_ROWS) rows += `<tr><td colspan="7" style="text-align:center;color:#e67e22;padding:8px;font-size:10px;">... mostrando ${MAX_ROWS} de ${sorted.length} linhas / 显示${MAX_ROWS}/${sorted.length}行</td></tr>`;\n'
        '  document.getElementById("tbody-inf").innerHTML = rows || '
        '"<tr><td colspan=\'7\' style=\'text-align:center;color:#ddd;padding:20px;\'>Nenhum dado para os filtros / 暂无数据</td></tr>";\n'
        '  document.getElementById("table-count").textContent = sorted.length+" IDs";\n'
        '}\n'
        '\n'
        'function setFiltro(key,val){\n'
        '  if(key==="dateFrom"){fDateFrom=val;localStorage.setItem("brdrive_from",val);}\n'
        '  else if(key==="dateTo"){fDateTo=val;localStorage.setItem("brdrive_to",val);}\n'
        '  else if(key==="trans") fTrans=val;\n'
        '  else if(key==="tipo") fTipo=val;\n'
        '  else if(key==="nivel") fNivel=val;\n'
        '  render();\n'
        '}\n'
        '\n'
        'function resetFiltros(){\n'
        '  fDateFrom="' + date_max + '"; fDateTo="' + date_max + '";\n'
        '  fTrans="Todos"; fTipo="Todos"; fNivel="Todos";\n'
        '  document.getElementById("f-date-from").value=fDateFrom;\n'
        '  document.getElementById("f-date-to").value=fDateTo;\n'
        '  document.getElementById("f-trans").value="Todos";\n'
        '  document.getElementById("f-tipo").value="Todos";\n'
        '  document.getElementById("f-nivel").value="Todos";\n'
        '  render();\n'
        '}\n'
        '\n'
        'document.addEventListener("DOMContentLoaded",function(){\n'
        '  var fi=document.getElementById("f-date-from"),ti=document.getElementById("f-date-to");\n'
        '  if(fi)fi.value=fDateFrom;\n'
        '  if(ti)ti.value=fDateTo;\n'
        '  render();\n'
        '});\n'
        '</script>'
    )

    return pagina_html(nav, conteudo + js, "Gestão de Infrações")

def pg_rotas(df):
    if "ORIGEM" not in df.columns:
        df["ORIGEM"] = "SP BRE"
    if "Subtipo de linha" not in df.columns:
        df["Subtipo de linha"] = "ROTA-001"
    if "Horário real de saída" not in df.columns:
        df["Horário real de saída"] = None

    total = len(df)
    fez = df["Horário real de saída"].notna().sum()
    nao_fez = total - fez
    pct_exec = round(fez / max(total,1) * 100, 1)
    c_exec = cor_kpi(pct_exec)
    periodo = f'{df["DATA"].min().strftime("%d/%m/%Y")} → {df["DATA"].max().strftime("%d/%m/%Y")}' if not df.empty else "—"
    mes = df["MÊS"].mode()[0] if not df.empty else "—"

    rows = ""
    for origem, grp in df.groupby("ORIGEM"):
        b_ids = len(grp); b_fez = grp["Horário real de saída"].notna().sum()
        b_nao = b_ids - b_fez; b_pct = round(b_fez/max(b_ids,1)*100,1)
        bc = cor_kpi(b_pct); bw = f"{min(b_pct,100):.0f}"
        rows += f"""<tr style="background:#2a0808;">
          <td colspan="2" style="padding:8px 10px;color:#fff;font-weight:700;font-size:12px;border-top:2px solid rgba(192,57,43,0.5);">BASE: {origem}</td>
          <td style="text-align:center;color:#f1c40f;font-weight:700;border-top:2px solid rgba(192,57,43,0.5);">{b_ids}</td>
          <td style="text-align:center;color:#27ae60;font-weight:700;border-top:2px solid rgba(192,57,43,0.5);">{b_fez}</td>
          <td style="text-align:center;color:#e74c3c;font-weight:700;border-top:2px solid rgba(192,57,43,0.5);">{b_nao}</td>
          <td style="border-top:2px solid rgba(192,57,43,0.5);">
            <div style="display:flex;align-items:center;gap:6px;">
              <span style="color:{bc};font-weight:700;min-width:42px;">{b_pct:.1f}%</span>
              <div style="background:rgba(255,255,255,0.07);border-radius:3px;height:7px;width:80px;">
                <div style="height:7px;border-radius:3px;background:{bc};width:{bw}%;"></div>
              </div>
            </div>
          </td>
          <td style="border-top:2px solid rgba(192,57,43,0.5);"></td>
        </tr>"""
        for i, (_, r) in enumerate(grp.groupby(["Subtipo de linha","Transportador"]).first().reset_index().iterrows()):
            rota = r.get("Subtipo de linha","—"); trans = r.get("Transportador","—")
            r_ids = len(grp[(grp["Subtipo de linha"]==rota)&(grp["Transportador"]==trans)])
            r_fez = grp[(grp["Subtipo de linha"]==rota)&(grp["Transportador"]==trans)]["Horário real de saída"].notna().sum()
            r_nao = r_ids - r_fez; r_pct = round(r_fez/max(r_ids,1)*100,1)
            rc = cor_kpi(r_pct); rw = f"{min(r_pct,100):.0f}"
            bgr = "#1f0e0e" if i%2==0 else "#1a1010"
            cond = grp[(grp["Subtipo de linha"]==rota)]["CONDUTOR"].iloc[0] if len(grp[grp["Subtipo de linha"]==rota]) > 0 else "—"
            rows += f"""<tr style="background:{bgr};">
              <td style="padding:6px 10px 6px 22px;color:#5dade2;font-weight:600;font-size:11px;">{rota}</td>
              <td><span class="trans-badge">{trans}</span></td>
              <td style="text-align:center;color:#f1c40f;font-weight:700;">{r_ids}</td>
              <td style="text-align:center;color:#27ae60;font-weight:700;">{r_fez}</td>
              <td style="text-align:center;color:#e74c3c;font-weight:700;">{r_nao}</td>
              <td>
                <div style="display:flex;align-items:center;gap:6px;">
                  <span style="color:{rc};font-weight:700;min-width:42px;">{r_pct:.1f}%</span>
                  <div style="background:rgba(255,255,255,0.07);border-radius:3px;height:7px;width:80px;">
                    <div style="height:7px;border-radius:3px;background:{rc};width:{rw}%;"></div>
                  </div>
                </div>
              </td>
              <td style="color:#fff;font-size:10px;">{cond}</td>
            </tr>"""

    nav = navbar("rotas")
    conteudo = f"""
    <div class="page-header">
      <div class="header-left">
        <div class="header-icon" style="background:#C0392B;">BR</div>
        <div>
          <div class="header-title">EXECUÇÃO DE ROTAS / 路线执行情况</div>
          <div class="header-sub">FEZ vs NÃO FEZ · 出发执行 · {periodo}</div>
        </div>
      </div>
      <div class="header-badges">
        <div class="header-badge"><div class="hb-label">MÊS/月</div><div class="hb-value">{mes}</div></div>
        <div class="header-badge"><div class="hb-label">PERÍODO/期间</div><div class="hb-value">{periodo}</div></div>
      </div>
    </div>

    <div class="kpi-grid kpi-grid-4">
      {kpi_card("TOTAL IDs / 行程总数", total, "rotas programadas", "#C0392B")}
      {kpi_card("FEZ / 已出发", fez, "saíram no período", "#27ae60", "#27ae60")}
      {kpi_card("NÃO FEZ / 未出发", nao_fez, "não saíram", "#e74c3c", "#e74c3c")}
      {kpi_card("% EXECUÇÃO / 执行率", fmt_pct(pct_exec), "do total programado", c_exec, c_exec)}
    </div>

    <div class="table-wrap">
      <div class="table-header" style="background:#C0392B;">
        <span>EXECUÇÃO POR BASE E ROTA / 按基地和路线执行</span>
        <small>fez vs não fez / 已出发 vs 未出发</small>
      </div>
      <div class="table-scroll">
        <table>
          <thead><tr style="background:#1a0808;border-bottom:2px solid rgba(192,57,43,0.4);">
            <th style="text-align:left;color:#C0392B;width:22%;">BASE/ROTA</th>
            <th style="text-align:left;color:#C0392B;width:14%;">TRANSPORTADORA</th>
            <th style="text-align:center;color:#C0392B;width:8%;">IDs</th>
            <th style="text-align:center;color:#C0392B;width:8%;">FEZ</th>
            <th style="text-align:center;color:#C0392B;width:8%;">NÃO FEZ</th>
            <th style="text-align:left;color:#C0392B;width:22%;">% EXECUÇÃO</th>
            <th style="text-align:left;color:#C0392B;width:18%;">CONDUTOR</th>
          </tr></thead>
          <tbody>{rows}</tbody>
          <tfoot><tr style="background:#C0392B;">
            <td colspan="2" style="font-weight:700;font-size:12px;">TOTAL GERAL</td>
            <td style="text-align:center;font-weight:700;font-size:12px;">{total}</td>
            <td style="text-align:center;font-weight:700;font-size:12px;">{fez}</td>
            <td style="text-align:center;font-weight:700;font-size:12px;">{nao_fez}</td>
            <td style="font-weight:700;font-size:12px;">{pct_exec:.1f}% execução</td>
            <td></td>
          </tr></tfoot>
        </table>
      </div>
    </div>

    <div class="legend-bar">
      <div class="legend-item" style="color:#27ae60;"><div class="legend-dot" style="background:#27ae60;"></div>≥85% execução</div>
      <div class="legend-item" style="color:#e67e22;"><div class="legend-dot" style="background:#e67e22;"></div>70–84% atenção</div>
      <div class="legend-item" style="color:#e74c3c;"><div class="legend-dot" style="background:#e74c3c;"></div>&lt;70% crítico</div>
    </div>"""
    return pagina_html(nav, conteudo, "Execução de Rotas")

def pg_evol_semanal(df):
    # ── Prepara dados ──────────────────────────────────────────────────────
    df2 = df.copy()
    iso = df2['DATA'].dt.isocalendar()
    df2['semana_num'] = iso.week.astype(int)
    df2['ano']        = iso.year.astype(int)
    df2['semana']     = 'W' + df2['semana_num'].astype(str).str.zfill(2) + ' - ' + df2['ano'].astype(str)
    df2['semana_ord'] = df2['ano'] * 100 + df2['semana_num']
    df2['dow']        = df2['DATA'].dt.dayofweek   # 0=seg … 6=dom
    df2['data_str']   = df2['DATA'].dt.strftime('%d/%m')

    transportadoras = sorted(df2['Transportador'].dropna().unique().tolist())
    tipos = ["Coleta Base", "Coleta PA", "Devolução", "Secundaria"]

    # Semana mais recente como padrão
    semana_max_ord = int(df2['semana_ord'].max())
    semana_max_lbl = df2.loc[df2['semana_ord'] == semana_max_ord, 'semana'].iloc[0]

    # Opções de semana (ordenadas)
    semanas_ord = df2[['semana','semana_ord']].drop_duplicates().sort_values('semana_ord')
    opts_sem = ''.join(
        f'<option value="{r.semana_ord}" {"selected" if r.semana_ord == semana_max_ord else ""}>{r.semana}</option>'
        for _, r in semanas_ord.iterrows()
    )
    opts_trans = '<option value="Todos">Todos</option>' + ''.join(
        f'<option value="{t}">{t}</option>' for t in transportadoras)
    opts_tipo = '<option value="Todos">Todos</option>' + ''.join(
        f'<option value="{t}">{t}</option>' for t in tipos)

    # JSON compacto: semana_ord, dia_semana, data_str, transportador, tipo, s, c
    cols = ['semana_ord', 'dow', 'data_str', 'Transportador', 'TIPO DE OPERAÇÃO',
            'Tempo Saida OFF', 'Tempo chegada OFF']
    df_js = df2[cols].fillna('')
    df_js.columns = ['wo', 'dow', 'dt', 't', 'op', 's', 'c']
    data_json = df_js.to_json(orient='records', force_ascii=False)

    nav = navbar("evol_semanal")

    js = (
        '<script>\n'
        'const RAW=' + data_json + ';\n'
        'let fSem=' + str(semana_max_ord) + ',fTrans="Todos",fTipo="Todos";\n'
        'const DIAS=[["SEG","周一"],["TER","周二"],["QUA","周三"],["QUI","周四"],["SEX","周五"],["SÁB","周六"],["DOM","周日"]];\n'
        'const DAY_COR=["#e67e22","#5dade2","#5dade2","#5dade2","#5dade2","#aaa","#e74c3c"];\n'
        '\n'
        'function cor(v){return v>=85?"#27ae60":v>=70?"#e67e22":"#e74c3c";}\n'
        'function pct(a,b){return b>0?Math.round(a/b*1000)/10:0;}\n'
        'function semAnt(wo){const yr=Math.floor(wo/100),wk=wo%100;return wk<=1?(yr-1)*100+52:yr*100+(wk-1);}\n'
        '\n'
        'function getDias(woFilter){\n'
        '  const dias={};\n'
        '  RAW.forEach(r=>{\n'
        '    if(r.wo!==woFilter) return;\n'
        '    if(fTrans!=="Todos"&&r.t!==fTrans) return;\n'
        '    if(fTipo!=="Todos"&&r.op!==fTipo) return;\n'
        '    if(!dias[r.dow]) dias[r.dow]={ids:0,okS:0,okC:0,dt:""};\n'
        '    dias[r.dow].ids++;\n'
        '    if(r.s==="OK") dias[r.dow].okS++;\n'
        '    if(r.c==="OK") dias[r.dow].okC++;\n'
        '    if(r.dt) dias[r.dow].dt=r.dt;\n'
        '  });\n'
        '  return dias;\n'
        '}\n'
        '\n'
        'function buildChart(dAt,dAnt,campo){\n'
        '  const H=155;\n'
        '  const metaY=Math.round((1-0.85)*H); // linha de meta a 85%\n'
        '  let html=\'<div style="display:flex;gap:5px;">\';\n'
        '  let cards=\'<div style="display:flex;gap:5px;margin-top:12px;">\';\n'
        '  for(let d=0;d<7;d++){\n'
        '    const vA=dAt[d]||{ids:0,okS:0,okC:0,dt:""};\n'
        '    const vP=dAnt[d]||{ids:0,okS:0,okC:0,dt:""};\n'
        '    const valA=campo==="s"?pct(vA.okS,vA.ids):pct(vA.okC,vA.ids);\n'
        '    const valP=campo==="s"?pct(vP.okS,vP.ids):pct(vP.okC,vP.ids);\n'
        '    const hA=vA.ids>0?Math.max(6,Math.round(valA/100*H)):0;\n'
        '    const hP=vP.ids>0?Math.max(6,Math.round(valP/100*H)):0;\n'
        '    const cA=cor(valA);\n'
        '    const diff=vP.ids>0?valA-valP:null;\n'
        '    const tStr=diff!==null?(diff>0.5?"▲ +":diff<-0.5?"▼ ":"→ ")+Math.abs(diff).toFixed(1)+"%":"";\n'
        '    const tCor=diff===null?"#aaa":diff>0.5?"#27ae60":diff<-0.5?"#e74c3c":"#aaa";\n'
        '    const dc=DAY_COR[d];\n'
        '    const dn=DIAS[d][0];\n'
        '    const dz=DIAS[d][1];\n'
        '    const dt=vA.dt||vP.dt||"";\n'
        '    const lbl=vA.ids>0?valA.toFixed(1)+"%":"—";\n'
        '    // cor de fundo do card baseada na performance\n'
        '    const cardBg=valA>=85?"rgba(39,174,96,0.1)":valA>=70?"rgba(230,126,34,0.1)":"rgba(231,76,60,0.1)";\n'
        '    const cardBd=valA>=85?"rgba(39,174,96,0.35)":valA>=70?"rgba(230,126,34,0.35)":"rgba(231,76,60,0.35)";\n'
        '\n'
        '    // Coluna do gráfico\n'
        '    html+=`<div style="flex:1;display:flex;flex-direction:column;align-items:center;min-width:80px;">`\n'
        '         +`<div style="font-size:10px;color:${tCor};font-weight:700;height:18px;text-align:center;">${tStr}</div>`\n'
        '         +`<div style="font-size:15px;color:${cA};font-weight:800;height:22px;text-align:center;line-height:22px;">${lbl}</div>`\n'
        '         +`<div style="width:100%;height:${H}px;position:relative;display:flex;align-items:flex-end;justify-content:center;gap:4px;padding:0 14px;box-sizing:border-box;">`\n'
        '         +`  <div style="position:absolute;top:${metaY}px;left:0;right:0;border-top:1.5px dashed rgba(255,255,255,0.3);pointer-events:none;z-index:2;"></div>`\n'
        '         +`  <div style="position:absolute;top:${metaY-11}px;right:6px;font-size:8px;color:rgba(255,255,255,0.75);z-index:2;">85%</div>`\n'
        '         +`  <div style="width:38%;height:${hP}px;background:#1a5276;border-radius:4px 4px 0 0;opacity:0.85;flex-shrink:0;"></div>`\n'
        '         +`  <div style="width:38%;height:${hA}px;background:${cA};border-radius:4px 4px 0 0;flex-shrink:0;"></div>`\n'
        '         +"</div>"\n'
        '         +`<div style="font-size:11px;color:${dc};font-weight:700;margin-top:7px;text-align:center;">${dn}</div>`\n'
        '         +`<div style="font-size:9px;color:rgba(255,255,255,0.6);text-align:center;">${dz}</div>`\n'
        '         +`<div style="font-size:9px;color:#ccc;text-align:center;">${dt}</div>`\n'
        '         +"</div>";\n'
        '\n'
        '    // Card inferior com cor de fundo por performance\n'
        '    const antStr=vP.ids>0?`ant: ${valP.toFixed(1)}%`:"";\n'
        '    cards+=`<div style="flex:1;min-width:90px;background:${cardBg};border:1px solid ${cardBd};border-radius:8px;padding:10px 8px;text-align:center;">`\n'
        '          +`<div style="font-size:9px;color:${dc};font-weight:700;margin-bottom:5px;">${dn}/${dz} ${dt}</div>`\n'
        '          +`<div style="font-size:22px;font-weight:900;color:${cA};line-height:1.1;">${lbl}</div>`\n'
        '          +`<div style="font-size:10px;color:${tCor};font-weight:700;margin-top:5px;">${tStr}</div>`\n'
        '          +`<div style="font-size:9px;color:#ddd;margin-top:3px;">${antStr}</div>`\n'
        '          +`<div style="font-size:8px;color:#ccc;margin-top:2px;">${vA.ids>0?vA.ids+" viagens":""}</div>`\n'
        '          +"</div>";\n'
        '  }\n'
        '  html+="</div>";\n'
        '  cards+="</div>";\n'
        '  return html+cards;\n'
        '}\n'
        '\n'
        'function render(){\n'
        '  const ant=semAnt(fSem);\n'
        '  const dAt=getDias(fSem), dAnt=getDias(ant);\n'
        '\n'
        '  // Labels semana\n'
        '  const lbl=r=>{ const yr=Math.floor(r/100),wk=r%100; return "W"+String(wk).padStart(2,"0")+" - "+yr; };\n'
        '  document.getElementById("kpi-sem").textContent=lbl(fSem);\n'
        '  document.getElementById("kpi-ant").textContent=lbl(ant);\n'
        '\n'
        '  // Médias saída\n'
        '  let tS=0,oS=0,tC=0,oC=0;\n'
        '  for(let d=0;d<7;d++){\n'
        '    const v=dAt[d]; if(!v) continue;\n'
        '    tS+=v.ids; oS+=v.okS; tC+=v.ids; oC+=v.okC;\n'
        '  }\n'
        '  const mS=pct(oS,tS), mC=pct(oC,tC);\n'
        '  const elMS=document.getElementById("kpi-media-s");\n'
        '  elMS.textContent=mS.toFixed(1)+"%"; elMS.style.color=cor(mS);\n'
        '  const elMC=document.getElementById("kpi-media-c");\n'
        '  elMC.textContent=mC.toFixed(1)+"%"; elMC.style.color=cor(mC);\n'
        '\n'
        '  document.getElementById("bloco-saida").innerHTML=buildChart(dAt,dAnt,"s");\n'
        '  document.getElementById("bloco-chegada").innerHTML=buildChart(dAt,dAnt,"c");\n'
        '}\n'
        '\n'
        'function setFiltro(k,v){\n'
        '  if(k==="sem") fSem=parseInt(v);\n'
        '  else if(k==="trans") fTrans=v;\n'
        '  else if(k==="tipo") fTipo=v;\n'
        '  render();\n'
        '}\n'
        'function resetFiltros(){\n'
        '  fSem=' + str(semana_max_ord) + '; fTrans="Todos"; fTipo="Todos";\n'
        '  document.getElementById("f-sem").value=fSem;\n'
        '  document.getElementById("f-trans").value="Todos";\n'
        '  document.getElementById("f-tipo").value="Todos";\n'
        '  render();\n'
        '}\n'
        'document.addEventListener("DOMContentLoaded",render);\n'
        '</script>'
    )

    # ── Reconstrói o conteúdo com IDs simplificados ─────────────────────────
    sel_style2 = 'width:100%;background:#0e0808;border:1px solid rgba(192,57,43,0.3);color:#fff;border-radius:4px;padding:5px 6px;font-size:11px;'
    filtro2 = f"""
    <div style="width:210px;flex-shrink:0;">
      <div style="background:#1a1010;border-radius:8px;padding:14px;border:1px solid rgba(192,57,43,0.25);position:sticky;top:16px;">
        <div style="color:#C0392B;font-size:11px;font-weight:700;letter-spacing:1px;margin-bottom:12px;padding-bottom:8px;border-bottom:1px solid rgba(192,57,43,0.3);">FILTROS / 筛选</div>
        <div style="margin-bottom:10px;">
          <div style="color:#ddd;font-size:9px;letter-spacing:1px;margin-bottom:4px;">月份 | Semana</div>
          <select id="f-sem" onchange="setFiltro('sem',this.value)" style="{sel_style2}">{opts_sem}</select>
        </div>
        <div style="margin-bottom:10px;">
          <div style="color:#ddd;font-size:9px;letter-spacing:1px;margin-bottom:4px;">承运商 | Transportador</div>
          <select id="f-trans" onchange="setFiltro('trans',this.value)" style="{sel_style2}">{opts_trans}</select>
        </div>
        <div style="margin-bottom:14px;">
          <div style="color:#ddd;font-size:9px;letter-spacing:1px;margin-bottom:4px;">运营类型 | Tipo Op.</div>
          <select id="f-tipo" onchange="setFiltro('tipo',this.value)" style="{sel_style2}">{opts_tipo}</select>
        </div>
        <button onclick="resetFiltros()" style="width:100%;background:rgba(192,57,43,0.2);border:1px solid rgba(192,57,43,0.4);color:#C0392B;border-radius:4px;padding:7px;font-size:10px;font-weight:700;cursor:pointer;letter-spacing:1px;">RESET / 重置</button>
        <div style="margin-top:14px;padding-top:10px;border-top:1px solid rgba(255,255,255,0.07);">
          <div style="color:#ddd;font-size:9px;margin-bottom:6px;letter-spacing:1px;">LEGENDA / 图例</div>
          <div style="display:flex;align-items:center;gap:6px;margin-bottom:4px;"><div style="width:10px;height:10px;border-radius:2px;background:#27ae60;"></div><span style="font-size:9px;color:#ddd;">≥85% / 达标</span></div>
          <div style="display:flex;align-items:center;gap:6px;margin-bottom:4px;"><div style="width:10px;height:10px;border-radius:2px;background:#e67e22;"></div><span style="font-size:9px;color:#ddd;">70–84% / 注意</span></div>
          <div style="display:flex;align-items:center;gap:6px;margin-bottom:4px;"><div style="width:10px;height:10px;border-radius:2px;background:#e74c3c;"></div><span style="font-size:9px;color:#ddd;">&lt;70% / 危急</span></div>
          <div style="display:flex;align-items:center;gap:6px;"><div style="width:10px;height:10px;border-radius:2px;background:#1a5276;"></div><span style="font-size:9px;color:#ddd;">Semana anterior / 上周</span></div>
        </div>
      </div>
    </div>"""

    conteudo2 = f"""
    <div class="page-header">
      <div class="header-left">
        <div class="header-icon" style="background:#C0392B;">BR</div>
        <div>
          <div class="header-title">EVOLUÇÃO DIA A DIA / 每日演变 — SAÍDA E CHEGADA / 出发与到达</div>
          <div class="header-sub">COLORIDO=ATUAL/彩=本周 · AZUL=ANTERIOR/蓝=上周 · SEG→DOM/周一→周日</div>
        </div>
      </div>
      <div class="header-badges">
        <div class="header-badge"><div class="hb-label">SEMANA/周</div><div class="hb-value" id="kpi-sem" style="color:#C0392B;">{semana_max_lbl}</div></div>
        <div class="header-badge"><div class="hb-label">ANTERIOR/上周</div><div class="hb-value" id="kpi-ant" style="color:#2980b9;">—</div></div>
        <div class="header-badge"><div class="hb-label">MÉDIA SAÍDA/平均出发</div><div class="hb-value" id="kpi-media-s">—</div></div>
        <div class="header-badge"><div class="hb-label">MÉDIA CHEGADA/平均到达</div><div class="hb-value" id="kpi-media-c">—</div></div>
      </div>
    </div>

    <div style="display:flex;gap:12px;align-items:flex-start;">
      <div style="flex:1;min-width:0;">

        <div class="table-wrap" style="margin-bottom:12px;">
          <div style="background:#C0392B;padding:8px 14px;display:flex;justify-content:space-between;align-items:center;">
            <span style="color:#fff;font-size:11px;font-weight:700;">▶ SAÍDA/出发 — app ativo na saída/出发时应用激活</span>
            <small style="color:rgba(255,255,255,0.95);font-size:9px;">azul/蓝=semana anterior · colorido/彩=semana atual</small>
          </div>
          <div style="padding:12px 14px;" id="bloco-saida"></div>
        </div>

        <div class="table-wrap" style="margin-bottom:12px;">
          <div style="background:#1A5276;padding:8px 14px;display:flex;justify-content:space-between;align-items:center;">
            <span style="color:#fff;font-size:11px;font-weight:700;">▶ CHEGADA/到达 — app ativo na chegada/到达时应用激活</span>
            <small style="color:rgba(255,255,255,0.95);font-size:9px;">azul/蓝=semana anterior · colorido/彩=semana atual</small>
          </div>
          <div style="padding:12px 14px;" id="bloco-chegada"></div>
        </div>

        <div class="legend-bar">
          <div class="legend-item"><div class="legend-dot" style="background:#27ae60;"></div>≥85% meta/达标</div>
          <div class="legend-item"><div class="legend-dot" style="background:#e67e22;"></div>70–84% atenção/注意</div>
          <div class="legend-item"><div class="legend-dot" style="background:#e74c3c;"></div>&lt;70% crítico/危急</div>
          <div class="legend-item"><div class="legend-dot" style="background:#1a5276;"></div>Semana anterior/上周</div>
        </div>
      </div>
      {filtro2}
    </div>"""

    return pagina_html(nav, conteudo2 + js, "App Semanal")

def pg_app_mensal(df):
    df2 = df.copy()
    df2['mes_num'] = df2['DATA'].dt.month
    df2['ano']     = df2['DATA'].dt.year
    df2['mes_ord'] = df2['ano'] * 100 + df2['mes_num']

    tipos   = sorted(df2['TIPO DE OPERAÇÃO'].dropna().unique().tolist())
    ano_max = int(df2['ano'].max())
    anos    = sorted(df2['ano'].dropna().unique().tolist(), reverse=True)

    cols  = ['mes_ord', 'mes_num', 'Transportador', 'TIPO DE OPERAÇÃO',
             'Tempo Saida OFF', 'Tempo chegada OFF']
    df_js = df2[cols].fillna('')
    df_js.columns = ['mo', 'mn', 't', 'op', 's', 'c']
    data_json = df_js.to_json(orient='records', force_ascii=False)

    ss = 'background:#0e0808;border:1px solid rgba(192,57,43,0.3);color:#fff;border-radius:4px;padding:5px 8px;font-size:11px;'
    opts_ano  = '<option value="Todos">Todos</option>' + ''.join(
        f'<option value="{a}" {"selected" if a==ano_max else ""}>{a}</option>' for a in anos)
    opts_tipo = '<option value="Todos">Todos</option>' + ''.join(
        f'<option value="{t}">{t}</option>' for t in tipos)

    nav = navbar("app_mensal")

    conteudo = f"""
    <div class="page-header">
      <div class="header-left">
        <div class="header-icon" style="background:#1A5276;">📅</div>
        <div>
          <div class="header-title">UTILIZAÇÃO DO APP — COMPARATIVO MENSAL / APP月度对比</div>
          <div class="header-sub">App Saída e Chegada · Jan-Dez · por Transportadora / 出发与到达 · 1-12月 · 按承运商</div>
        </div>
      </div>
      <div style="display:flex;gap:10px;align-items:flex-end;">
        <div>
          <div style="color:#ddd;font-size:9px;letter-spacing:1px;margin-bottom:4px;">操作类型 | Tipo de Operação</div>
          <select id="f-tipo" onchange="setFiltro('tipo',this.value)" style="{ss}">{opts_tipo}</select>
        </div>
        <div>
          <div style="color:#ddd;font-size:9px;letter-spacing:1px;margin-bottom:4px;">年份 | Ano</div>
          <select id="f-ano" onchange="setFiltro('ano',this.value)" style="{ss}">{opts_ano}</select>
        </div>
      </div>
    </div>

    <!-- KPI cards por mês -->
    <div style="overflow-x:auto;margin-bottom:14px;">
      <div id="kpi-cards" style="display:flex;gap:8px;padding-bottom:4px;min-width:max-content;"></div>
    </div>

    <!-- 2x2 grid -->
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:12px;">
      <div class="table-wrap">
        <div style="background:#1A5276;padding:8px 12px;">
          <span style="color:#fff;font-size:10px;font-weight:700;">▶ APP SAÍDA/出发 — Evolução Mensal/月度趋势</span>
        </div>
        <div style="padding:12px;" id="chart-saida-evol"></div>
      </div>
      <div class="table-wrap">
        <div style="background:#1A5276;padding:8px 12px;display:flex;justify-content:space-between;align-items:center;">
          <span style="color:#fff;font-size:10px;font-weight:700;">▶ APP SAÍDA/出发 — Por Transportadora/按承运商</span>
          <span id="lbl-ts" style="color:rgba(255,255,255,0.9);font-size:9px;"></span>
        </div>
        <div style="padding:12px;" id="chart-saida-trans"></div>
      </div>
      <div class="table-wrap">
        <div style="background:#C0392B;padding:8px 12px;">
          <span style="color:#fff;font-size:10px;font-weight:700;">▶ APP CHEGADA/到达 — Evolução Mensal/月度趋势</span>
        </div>
        <div style="padding:12px;" id="chart-chegada-evol"></div>
      </div>
      <div class="table-wrap">
        <div style="background:#C0392B;padding:8px 12px;display:flex;justify-content:space-between;align-items:center;">
          <span style="color:#fff;font-size:10px;font-weight:700;">▶ APP CHEGADA/到达 — Por Transportadora/按承运商</span>
          <span id="lbl-tc" style="color:rgba(255,255,255,0.9);font-size:9px;"></span>
        </div>
        <div style="padding:12px;" id="chart-chegada-trans"></div>
      </div>
    </div>

    <div class="legend-bar">
      <div class="legend-item"><div class="legend-dot" style="background:#27ae60;"></div>≥85% meta/达标</div>
      <div class="legend-item"><div class="legend-dot" style="background:#e67e22;"></div>70–84% atenção/注意</div>
      <div class="legend-item"><div class="legend-dot" style="background:#e74c3c;"></div>&lt;70% crítico/危急</div>
      <div class="legend-item"><div style="width:18px;height:3px;background:#f39c12;border-radius:2px;margin-right:2px;"></div>Tendência/趋势</div>
    </div>"""

    js = (
        '<script>\n'
        'const RAW=' + data_json + ';\n'
        'let fAno=' + str(ano_max) + ',fTipo="Todos";\n'
        'const MESES=[["JAN","1月"],["FEV","2月"],["MAR","3月"],["ABR","4月"],["MAI","5月"],["JUN","6月"],'
        '["JUL","7月"],["AGO","8月"],["SET","9月"],["OUT","10月"],["NOV","11月"],["DEZ","12月"]];\n'
        'function cor(v){return v>=85?"#27ae60":v>=70?"#e67e22":"#e74c3c";}\n'
        'function pct(a,b){return b>0?Math.round(a/b*1000)/10:0;}\n'
        '\n'
        'function agg(){\n'
        '  const mes={},trans={};\n'
        '  RAW.forEach(r=>{\n'
        '    if(fAno!=="Todos"&&Math.floor(r.mo/100)!==fAno) return;\n'
        '    if(fTipo!=="Todos"&&r.op!==fTipo) return;\n'
        '    const m=r.mo%100;\n'
        '    if(!mes[m]) mes[m]={ids:0,okS:0,okC:0};\n'
        '    mes[m].ids++;if(r.s==="OK")mes[m].okS++;if(r.c==="OK")mes[m].okC++;\n'
        '    if(r.t){\n'
        '      if(!trans[r.t]) trans[r.t]={};\n'
        '      if(!trans[r.t][m]) trans[r.t][m]={ids:0,okS:0,okC:0};\n'
        '      trans[r.t][m].ids++;if(r.s==="OK")trans[r.t][m].okS++;if(r.c==="OK")trans[r.t][m].okC++;\n'
        '    }\n'
        '  });\n'
        '  return {mes,trans};\n'
        '}\n'
        '\n'
        '// ── KPI cards por mês ──────────────────────────────────────────\n'
        'function buildKPI(mes){\n'
        '  const mks=Object.keys(mes).map(Number).sort((a,b)=>a-b);\n'
        '  return mks.map((m,i)=>{\n'
        '    const v=mes[m],vP=i>0?mes[mks[i-1]]:null;\n'
        '    const pS=pct(v.okS,v.ids),pC=pct(v.okC,v.ids);\n'
        '    const dS=vP?pS-pct(vP.okS,vP.ids):null,dC=vP?pC-pct(vP.okC,vP.ids):null;\n'
        '    const fD=d=>d===null?"":((d>=0?"▲ +":"▼ ")+Math.abs(d).toFixed(1)+"pp");\n'
        '    const cD=d=>d===null?"#ddd":d>=0?"#27ae60":"#e74c3c";\n'
        '    const pt=MESES[m-1][0],zh=MESES[m-1][1];\n'
        '    return `<div style="background:#1a1010;border:1px solid rgba(255,255,255,0.08);border-radius:8px;padding:10px 14px;min-width:130px;">`\n'
        '          +`<div style="font-size:9px;color:#5dade2;font-weight:700;margin-bottom:6px;">${pt}/${zh} · <span style="color:#ddd;font-weight:400;">${v.ids.toLocaleString()} viagens</span></div>`\n'
        '          +`<div style="display:flex;gap:14px;">`\n'
        '          +`<div><div style="font-size:8px;color:#ddd;margin-bottom:2px;">Saída/出发</div>`\n'
        '          +`<div style="font-size:19px;font-weight:900;color:${cor(pS)};line-height:1.1;">${pS.toFixed(1)}%</div>`\n'
        '          +`<div style="font-size:9px;color:${cD(dS)};font-weight:700;">${fD(dS)}</div></div>`\n'
        '          +`<div><div style="font-size:8px;color:#ddd;margin-bottom:2px;">Chegada/到达</div>`\n'
        '          +`<div style="font-size:19px;font-weight:900;color:${cor(pC)};line-height:1.1;">${pC.toFixed(1)}%</div>`\n'
        '          +`<div style="font-size:9px;color:${cD(dC)};font-weight:700;">${fD(dC)}</div></div>`\n'
        '          +"</div></div>";\n'
        '  }).join("");\n'
        '}\n'
        '\n'
        '// ── Gráfico evolução mensal com linha de tendência (SVG) ────────\n'
        'function buildEvol(mes,campo){\n'
        '  const H=175,VW=1200,N=12,cw=VW/N;\n'
        '  const m85=H*(1-0.85),m70=H*(1-0.70);\n'
        '  let lbls="",bars="",names="",pts=[],dots="";\n'
        '  for(let m=1;m<=12;m++){\n'
        '    const v=mes[m]||{ids:0,okS:0,okC:0};\n'
        '    const val=campo==="s"?pct(v.okS,v.ids):pct(v.okC,v.ids);\n'
        '    const h=v.ids>0?Math.max(4,Math.round(val/100*H)):0;\n'
        '    const c=cor(val),lbl=v.ids>0?val.toFixed(1)+"%":"—";\n'
        '    const xC=(m-0.5)*cw,yC=H*(1-val/100);\n'
        '    if(v.ids>0){pts.push(`${xC},${yC}`);dots+=`<circle cx="${xC}" cy="${yC}" r="8" fill="#f39c12" stroke="#fff" stroke-width="2.5"/>`;}\n'
        '    lbls+=`<div style="flex:1;text-align:center;font-size:11px;color:${c};font-weight:800;height:18px;line-height:18px;">${lbl}</div>`;\n'
        '    bars+=`<div style="flex:1;display:flex;align-items:flex-end;justify-content:center;height:${H}px;">`\n'
        '         +`<div style="width:60%;height:${h}px;background:${v.ids>0?c:"rgba(255,255,255,0.06)"};border-radius:3px 3px 0 0;"></div>`\n'
        '         +"</div>";\n'
        '    names+=`<div style="flex:1;text-align:center;padding-top:4px;">`\n'
        '           +`<div style="font-size:10px;color:#5dade2;font-weight:700;">${MESES[m-1][0]}</div>`\n'
        '           +`<div style="font-size:8px;color:rgba(255,255,255,0.6);">${MESES[m-1][1]}</div>`\n'
        '           +"</div>";\n'
        '  }\n'
        '  const line=pts.length>1?`<polyline points="${pts.join(" ")}" fill="none" stroke="#f39c12" stroke-width="4" stroke-linejoin="round" stroke-linecap="round"/>`:""; \n'
        '  const svg=`<svg style="position:absolute;top:0;left:0;width:100%;height:${H}px;pointer-events:none;"`\n'
        '            +` viewBox="0 0 ${VW} ${H}" preserveAspectRatio="none">`\n'
        '            +`<line x1="0" y1="${m85}" x2="${VW}" y2="${m85}" stroke="rgba(255,255,255,0.35)" stroke-width="2" stroke-dasharray="8,5"/>`\n'
        '            +`<text x="10" y="${m85-5}" fill="rgba(255,255,255,0.55)" font-size="22" font-family="Arial">85%</text>`\n'
        '            +`<line x1="0" y1="${m70}" x2="${VW}" y2="${m70}" stroke="rgba(255,255,255,0.18)" stroke-width="1.5" stroke-dasharray="6,5"/>`\n'
        '            +`<text x="10" y="${m70-5}" fill="rgba(255,255,255,0.3)" font-size="22" font-family="Arial">70%</text>`\n'
        '            +line+dots+"</svg>";\n'
        '  return `<div><div style="display:flex;">${lbls}</div>`\n'
        '        +`<div style="position:relative;"><div style="display:flex;">${bars}</div>${svg}</div>`\n'
        '        +`<div style="display:flex;">${names}</div>`\n'
        '        +`<div style="margin-top:6px;display:flex;align-items:center;gap:5px;">`\n'
        '        +`<div style="width:18px;height:3px;background:#f39c12;border-radius:2px;"></div>`\n'
        '        +`<span style="font-size:9px;color:#ddd;">Tendência/趋势</span></div></div>`;\n'
        '}\n'
        '\n'
        '// ── Comparativo por transportadora (últimos 2 meses) ───────────\n'
        'function buildTrans(trans,campo,last2){\n'
        '  if(!last2||last2.length===0) return `<div style="color:#ddd;padding:20px;text-align:center;">Sem dados</div>`;\n'
        '  const mB=last2[last2.length-1],mA=last2.length>1?last2[last2.length-2]:null;\n'
        '  const tList=Object.keys(trans).sort().filter(t=>{\n'
        '    const vA=mA&&trans[t]&&trans[t][mA]?trans[t][mA]:{ids:0,okS:0,okC:0};\n'
        '    const vB=trans[t]&&trans[t][mB]?trans[t][mB]:{ids:0,okS:0,okC:0};\n'
        '    const valA=campo==="s"?pct(vA.okS,vA.ids):pct(vA.okC,vA.ids);\n'
        '    const valB=campo==="s"?pct(vB.okS,vB.ids):pct(vB.okC,vB.ids);\n'
        '    return valA>0&&valB>0;\n'
        '  });\n'
        '  const H=130;\n'
        '  let cols="";\n'
        '  tList.forEach(t=>{\n'
        '    const vA=mA&&trans[t]&&trans[t][mA]?trans[t][mA]:{ids:0,okS:0,okC:0};\n'
        '    const vB=trans[t]&&trans[t][mB]?trans[t][mB]:{ids:0,okS:0,okC:0};\n'
        '    const valA=campo==="s"?pct(vA.okS,vA.ids):pct(vA.okC,vA.ids);\n'
        '    const valB=campo==="s"?pct(vB.okS,vB.ids):pct(vB.okC,vB.ids);\n'
        '    const hA=vA.ids>0?Math.max(3,Math.round(valA/100*H)):0;\n'
        '    const hB=vB.ids>0?Math.max(3,Math.round(valB/100*H)):0;\n'
        '    const lA=vA.ids>0?valA.toFixed(0)+"%":"—",lB=vB.ids>0?valB.toFixed(0)+"%":"—";\n'
        '    cols+=`<div style="flex:1;display:flex;flex-direction:column;align-items:center;">`\n'
        '         +`<div style="width:100%;display:flex;gap:2px;align-items:flex-end;height:${H+18}px;">`\n'
        '         +`  <div style="flex:1;display:flex;flex-direction:column;align-items:center;">`\n'
        '         +`    <div style="font-size:9px;color:#e67e22;font-weight:700;">${lA}</div>`\n'
        '         +`    <div style="width:100%;height:${hA}px;background:#e67e22;border-radius:2px 2px 0 0;"></div>`\n'
        '         +`  </div>`\n'
        '         +`  <div style="flex:1;display:flex;flex-direction:column;align-items:center;">`\n'
        '         +`    <div style="font-size:9px;color:#27ae60;font-weight:700;">${lB}</div>`\n'
        '         +`    <div style="width:100%;height:${hB}px;background:#27ae60;border-radius:2px 2px 0 0;"></div>`\n'
        '         +`  </div>`\n'
        '         +"</div>"\n'
        '         +`<div style="font-size:8px;color:#ddd;text-align:center;margin-top:3px;word-break:break-word;line-height:1.2;">${t.length>9?t.slice(0,9)+"…":t}</div>`\n'
        '         +"</div>";\n'
        '  });\n'
        '  let leg="";\n'
        '  if(mA) leg+=`<div style="display:flex;align-items:center;gap:4px;"><div style="width:10px;height:10px;background:#e67e22;border-radius:2px;"></div><span style="font-size:9px;color:#ddd;">${MESES[mA-1][0]}/${MESES[mA-1][1]}</span></div>`;\n'
        '  leg+=`<div style="display:flex;align-items:center;gap:4px;"><div style="width:10px;height:10px;background:#27ae60;border-radius:2px;"></div><span style="font-size:9px;color:#ddd;">${MESES[mB-1][0]}/${MESES[mB-1][1]}</span></div>`;\n'
        '  return `<div><div style="display:flex;gap:4px;">${cols}</div><div style="margin-top:8px;display:flex;gap:12px;">${leg}</div></div>`;\n'
        '}\n'
        '\n'
        'function render(){\n'
        '  const {mes,trans}=agg();\n'
        '  const mks=Object.keys(mes).map(Number).sort((a,b)=>a-b);\n'
        '  const last2=mks.slice(-2);\n'
        '  const mA=last2.length>1?last2[0]:null,mB=last2[last2.length-1]||null;\n'
        '  const lbl=mA&&mB?MESES[mA-1][0]+"/"+MESES[mA-1][1]+" vs "+MESES[mB-1][0]+"/"+MESES[mB-1][1]:"";\n'
        '  document.getElementById("kpi-cards").innerHTML=buildKPI(mes);\n'
        '  document.getElementById("chart-saida-evol").innerHTML=buildEvol(mes,"s");\n'
        '  document.getElementById("chart-chegada-evol").innerHTML=buildEvol(mes,"c");\n'
        '  document.getElementById("chart-saida-trans").innerHTML=buildTrans(trans,"s",last2);\n'
        '  document.getElementById("chart-chegada-trans").innerHTML=buildTrans(trans,"c",last2);\n'
        '  document.getElementById("lbl-ts").textContent=lbl;\n'
        '  document.getElementById("lbl-tc").textContent=lbl;\n'
        '}\n'
        'function setFiltro(k,v){\n'
        '  if(k==="ano") fAno=v==="Todos"?"Todos":parseInt(v);\n'
        '  else if(k==="tipo") fTipo=v;\n'
        '  render();\n'
        '}\n'
        'document.addEventListener("DOMContentLoaded",render);\n'
        '</script>'
    )
    return pagina_html(nav, conteudo + js, "App Mensal")

def pg_em_construcao(titulo, page_id, cor="#C0392B"):
    nav = navbar(page_id)
    conteudo = f"""
    <div style="display:flex;flex-direction:column;align-items:center;justify-content:center;min-height:60vh;gap:20px;">
      <div style="width:80px;height:80px;border-radius:16px;background:{cor};display:flex;align-items:center;justify-content:center;font-size:32px;">🚧</div>
      <div style="text-align:center;">
        <div style="font-size:20px;font-weight:700;color:#fff;margin-bottom:8px;">{titulo}</div>
        <div style="color:#ddd;font-size:13px;">Conecte seu arquivo <strong style="color:#fff;">dados.xlsx</strong> para ver os dados reais.</div>
        <div style="color:#ddd;font-size:12px;margin-top:6px;">Esta página está pronta para receber seus dados.</div>
      </div>
      <a href="home.html" style="background:{cor};color:#fff;padding:10px 24px;border-radius:8px;text-decoration:none;font-weight:700;font-size:13px;">← Voltar ao Menu</a>
    </div>"""
    return pagina_html(nav, conteudo, titulo)

def pg_inf_pontualidade(df):
    base_cols = ['DATA', 'MÊS', 'Transportador', 'CONDUTOR', 'TIPO DE OPERAÇÃO']
    opt_cols  = ['Semana Nome Nova', 'Número do ID',
                 'PONTUALIDADE SAÍDA', 'PONTUALIDADE CHEGADA']
    cols_needed = base_cols + [c for c in opt_cols if c in df.columns]

    df_js = df[[c for c in cols_needed]].copy()
    df_js['DATA'] = df_js['DATA'].dt.strftime('%Y-%m-%d')
    df_js = df_js.fillna('')

    rename_map = {
        'DATA': 'd', 'MÊS': 'm', 'Semana Nome Nova': 'sem',
        'Transportador': 't', 'CONDUTOR': 'c', 'Número do ID': 'nid',
        'TIPO DE OPERAÇÃO': 'op',
        'PONTUALIDADE SAÍDA': 'ps', 'PONTUALIDADE CHEGADA': 'pc'
    }
    df_js = df_js.rename(columns={k: v for k, v in rename_map.items() if k in df_js.columns})
    data_json = df_js.to_json(orient='records', force_ascii=False)

    transportadoras = sorted(df['Transportador'].dropna().unique().tolist())
    tipos = ["Coleta Base", "Coleta PA", "Devolução", "Secundaria"]
    date_min = df['DATA'].min().strftime('%Y-%m-%d') if not df.empty else ''
    date_max = df['DATA'].max().strftime('%Y-%m-%d') if not df.empty else ''

    opts_trans = '<option value="Todos">Todos</option>' + ''.join(
        f'<option value="{t}">{t}</option>' for t in transportadoras)
    opts_tipo = '<option value="Todos">Todos</option>' + ''.join(
        f'<option value="{t}">{t}</option>' for t in tipos)

    sel_style = 'width:100%;background:#0e0808;border:1px solid rgba(142,68,173,0.5);color:#fff;border-radius:4px;padding:5px 6px;font-size:11px;'
    inp_style = 'width:100%;background:#0e0808;border:1px solid rgba(142,68,173,0.5);color:#fff;border-radius:4px;padding:5px 6px;font-size:11px;color-scheme:dark;'

    filtro_html = f"""
    <div style="width:210px;flex-shrink:0;">
      <div style="background:#1a1010;border-radius:8px;padding:14px;border:1px solid rgba(142,68,173,0.3);position:sticky;top:16px;">
        <div style="color:#8e44ad;font-size:11px;font-weight:700;letter-spacing:1px;margin-bottom:12px;padding-bottom:8px;border-bottom:1px solid rgba(142,68,173,0.3);">FILTROS / 筛选</div>
        <div style="margin-bottom:10px;">
          <div style="color:#ddd;font-size:9px;letter-spacing:1px;margin-bottom:4px;">日期从 | Data De</div>
          <input type="date" id="f-date-from" value="{date_max}" onchange="setFiltro('dateFrom',this.value)" style="{inp_style}margin-bottom:4px;">
          <div style="color:#ddd;font-size:9px;letter-spacing:1px;margin-bottom:4px;margin-top:6px;">日期到 | Data Até</div>
          <input type="date" id="f-date-to" value="{date_max}" onchange="setFiltro('dateTo',this.value)" style="{inp_style}">
        </div>
        <div style="margin-bottom:10px;">
          <div style="color:#ddd;font-size:9px;letter-spacing:1px;margin-bottom:4px;">承运商 | Transportadora</div>
          <select id="f-trans" onchange="setFiltro('trans',this.value)" style="{sel_style}">{opts_trans}</select>
        </div>
        <div style="margin-bottom:10px;">
          <div style="color:#ddd;font-size:9px;letter-spacing:1px;margin-bottom:4px;">运营类型 | Tipo Op.</div>
          <select id="f-tipo" onchange="setFiltro('tipo',this.value)" style="{sel_style}">{opts_tipo}</select>
        </div>
        <div style="margin-bottom:14px;">
          <div style="color:#ddd;font-size:9px;letter-spacing:1px;margin-bottom:4px;">等级 | Nível</div>
          <select id="f-nivel" onchange="setFiltro('nivel',this.value)" style="{sel_style}">
            <option value="Todos">Todos</option>
            <option value="CRITICO">CRÍTICO</option>
            <option value="OK">OK</option>
          </select>
        </div>
        <button onclick="resetFiltros()" style="width:100%;background:rgba(142,68,173,0.2);border:1px solid rgba(142,68,173,0.4);color:#8e44ad;border-radius:4px;padding:7px;font-size:10px;font-weight:700;cursor:pointer;letter-spacing:1px;">RESET / 重置</button>
      </div>
    </div>"""

    nav = navbar("inf_pontualidade")
    conteudo = """
    <div class="page-header">
      <div class="header-left">
        <div class="header-icon" style="background:#8e44ad;">⚠️</div>
        <div>
          <div class="header-title">INF. DE PONTUALIDADE / 准时率违规管理</div>
          <div class="header-sub">DETALHE POR ID · 行程准时明细 · <span id="periodo-txt">—</span></div>
        </div>
      </div>
      <div class="header-badges">
        <div class="header-badge"><div class="hb-label">MÊS/月</div><div class="hb-value" id="hb-mes">—</div></div>
        <div class="header-badge"><div class="hb-label">SEMANA/周</div><div class="hb-value" id="hb-sem">—</div></div>
        <div class="header-badge"><div class="hb-label">PERÍODO/期间</div><div class="hb-value" id="hb-periodo">—</div></div>
        <div class="header-badge"><div class="hb-label">NÍVEL/层级</div><div class="hb-value" id="hb-nivel">TODOS</div></div>
      </div>
    </div>

    <div style="display:flex;gap:16px;align-items:flex-start;">
      <div style="flex:1;min-width:0;">

        <div class="kpi-grid kpi-grid-5" style="margin-bottom:12px;">
          <div class="kpi-card" style="border-top:3px solid #8e44ad;">
            <div class="kpi-label">TOTAL IDs / 总行程数</div>
            <div class="kpi-value" id="kpi-total">—</div>
            <div class="kpi-sub" id="kpi-total-sub">IDs no período</div>
          </div>
          <div class="kpi-card" style="border-top:3px solid #e74c3c;">
            <div class="kpi-label">CRÍTICO / 延误违规</div>
            <div class="kpi-value" id="kpi-critico" style="color:#e74c3c;">—</div>
            <div class="kpi-sub">qualquer atraso / 任意延迟</div>
          </div>
          <div class="kpi-card" style="border-top:3px solid #e74c3c;">
            <div class="kpi-label">ATRASADO SAÍDA / 出发延误</div>
            <div class="kpi-value" id="kpi-atr-s" style="color:#e74c3c;">—</div>
            <div class="kpi-sub">fora do prazo na saída</div>
          </div>
          <div class="kpi-card" style="border-top:3px solid #e67e22;">
            <div class="kpi-label">ATRASADO CHEGADA / 到达延误</div>
            <div class="kpi-value" id="kpi-atr-c" style="color:#e67e22;">—</div>
            <div class="kpi-sub">fora do prazo na chegada</div>
          </div>
          <div class="kpi-card" style="border-top:3px solid #27ae60;">
            <div class="kpi-label">% NO PRAZO / 准时率</div>
            <div class="kpi-value" id="kpi-ok-pct" style="color:#27ae60;">—</div>
            <div class="kpi-sub">saída E chegada no prazo</div>
          </div>
        </div>

        <div class="table-wrap">
          <div class="table-header" style="background:#8e44ad;">
            <span>DETALHE DE PONTUALIDADE POR CONDUTOR / 按司机准时明细</span>
            <small id="table-count">scroll para ver todos · 滚动查看全部</small>
          </div>
          <div class="table-scroll">
            <table>
              <thead><tr style="background:#1a0a1f;border-bottom:2px solid rgba(142,68,173,0.4);">
                <th style="text-align:left;color:#a569bd;">CONDUTOR/司机</th>
                <th style="text-align:left;color:#a569bd;">TRANSP./承运商</th>
                <th style="text-align:left;color:#a569bd;">NÚMERO DO ID</th>
                <th style="text-align:left;color:#a569bd;">TIPO OP.</th>
                <th style="text-align:center;color:#a569bd;">PONT. SAÍDA</th>
                <th style="text-align:center;color:#a569bd;">PONT. CHEGADA</th>
                <th style="text-align:center;color:#a569bd;">NÍVEL</th>
              </tr></thead>
              <tbody id="tbody-inf"></tbody>
            </table>
          </div>
        </div>

        <div class="legend-bar">
          <div class="legend-item" style="color:#e74c3c;"><div class="legend-dot" style="background:#e74c3c;"></div>CRÍTICO = qualquer atraso / 任意延迟</div>
          <div class="legend-item" style="color:#27ae60;"><div class="legend-dot" style="background:#27ae60;"></div>OK = saída E chegada no prazo / 均准时</div>
        </div>

      </div>
      FILTRO_PLACEHOLDER
    </div>"""

    conteudo = conteudo.replace('FILTRO_PLACEHOLDER', filtro_html)

    js = (
        '<script>\n'
        'const RAW = ' + data_json + ';\n'
        'const _DM="' + date_max + '",_DMi="' + date_min + '";\n'
        'const _lsM=localStorage.getItem("brdrive_data_max"),_lsF2=localStorage.getItem("brdrive_from"),_lsT2=localStorage.getItem("brdrive_to");\n'
        'if(_lsM!==_DM||(_lsF2&&_lsF2>_DM)||(_lsT2&&_lsT2<_DMi)){localStorage.removeItem("brdrive_from");localStorage.removeItem("brdrive_to");localStorage.setItem("brdrive_data_max",_DM);}\n'
        'const _lsF=localStorage.getItem("brdrive_from"),_lsT=localStorage.getItem("brdrive_to");\n'
        'let fDateFrom=_lsF||"' + date_max + '";\n'
        'let fDateTo=_lsT||"' + date_max + '";\n'
        'let fTrans = "Todos"; let fTipo = "Todos"; let fNivel = "Todos";\n'
        '\n'
        'function nivel(r){\n'
        '  const latS = r.ps && r.ps !== "No prazo";\n'
        '  const latC = r.pc && r.pc !== "No prazo";\n'
        '  if(latS || latC) return "CRITICO";\n'
        '  return "OK";\n'
        '}\n'
        '\n'
        'function getFiltrado(){\n'
        '  return RAW.filter(r=>{\n'
        '    if(fDateFrom && r.d < fDateFrom) return false;\n'
        '    if(fDateTo && r.d > fDateTo) return false;\n'
        '    if(fTrans!=="Todos" && r.t!==fTrans) return false;\n'
        '    if(fTipo!=="Todos" && r.op!==fTipo) return false;\n'
        '    if(fNivel!=="Todos" && nivel(r)!==fNivel) return false;\n'
        '    return true;\n'
        '  });\n'
        '}\n'
        '\n'
        'function render(){\n'
        '  const data = getFiltrado();\n'
        '  const total = data.length;\n'
        '  const critico = data.filter(r=>nivel(r)==="CRITICO").length;\n'
        '  const atrS = data.filter(r=>r.ps && r.ps!=="No prazo").length;\n'
        '  const atrC = data.filter(r=>r.pc && r.pc!=="No prazo").length;\n'
        '  const ambosOk = data.filter(r=>(!r.ps||r.ps==="No prazo")&&(!r.pc||r.pc==="No prazo")).length;\n'
        '  const okPct = total>0?(ambosOk/total*100):0;\n'
        '\n'
        '  document.getElementById("kpi-total").textContent = total;\n'
        '  document.getElementById("kpi-total-sub").textContent = total>0?total+" IDs no período":"sem dados";\n'
        '  document.getElementById("kpi-critico").textContent = critico;\n'
        '  document.getElementById("kpi-atr-s").textContent = atrS;\n'
        '  document.getElementById("kpi-atr-c").textContent = atrC;\n'
        '  const eOk=document.getElementById("kpi-ok-pct"); eOk.textContent=okPct.toFixed(1)+"%";\n'
        '  eOk.style.color=okPct>=85?"#27ae60":okPct>=70?"#e67e22":"#e74c3c";\n'
        '\n'
        '  const dates = data.map(r=>r.d).filter(Boolean).sort();\n'
        '  const periodo = dates.length>0 ? dates[0]+" → "+dates[dates.length-1] : "—";\n'
        '  document.getElementById("periodo-txt").textContent = periodo;\n'
        '  document.getElementById("hb-periodo").textContent = periodo;\n'
        '  const meses = [...new Set(data.map(r=>r.m).filter(Boolean))];\n'
        '  document.getElementById("hb-mes").textContent = meses.length>0 ? meses.join("/") : "—";\n'
        '  const sems = [...new Set(data.map(r=>(r.sem||"")).filter(Boolean))];\n'
        '  document.getElementById("hb-sem").textContent = sems.length>0 ? sems[sems.length-1] : "—";\n'
        '  document.getElementById("hb-nivel").textContent = fNivel==="Todos"?"TODOS":fNivel==="CRITICO"?"CRÍTICO":"OK";\n'
        '\n'
        '  const MAX_ROWS = 500;\n'
        '  const sorted = data.slice().sort((a,b)=>a.c<b.c?-1:a.c>b.c?1:0);\n'
        '  let rows = "";\n'
        '  sorted.slice(0,MAX_ROWS).forEach((r,i)=>{\n'
        '    const nv = nivel(r);\n'
        '    const bgr = i%2===0?"#12091a":"#0e0616";\n'
        '    const nc = nv==="CRITICO"?"#e74c3c":"#27ae60";\n'
        '    const nb = nv==="CRITICO"?"rgba(231,76,60,0.2)":"rgba(39,174,96,0.2)";\n'
        '    const nvLabel = nv==="CRITICO"?"CRÍTICO":"OK";\n'
        '    const latS = r.ps && r.ps!=="No prazo";\n'
        '    const latC = r.pc && r.pc!=="No prazo";\n'
        '    const cs = latS?"#e74c3c":"#27ae60"; const bgs=latS?"rgba(231,76,60,0.2)":"rgba(39,174,96,0.2)";\n'
        '    const cc = latC?"#e74c3c":"#27ae60"; const bgc=latC?"rgba(231,76,60,0.2)":"rgba(39,174,96,0.2)";\n'
        '    const psLbl = r.ps||"—"; const pcLbl = r.pc||"—";\n'
        '    rows += `<tr style="background:${bgr};">'
        '<td style="color:#5dade2;font-weight:600;">${r.c||"—"}</td>'
        '<td><span class="trans-badge">${r.t||"—"}</span></td>'
        '<td style="color:#f1c40f;font-weight:700;font-size:10px;">${r.nid||"—"}</td>'
        '<td style="color:#fff;font-size:10px;">${r.op||"—"}</td>'
        '<td style="text-align:center;"><span class="badge" style="background:${bgs};color:${cs};font-size:9px;">${psLbl}</span></td>'
        '<td style="text-align:center;"><span class="badge" style="background:${bgc};color:${cc};font-size:9px;">${pcLbl}</span></td>'
        '<td style="text-align:center;"><span class="badge" style="background:${nb};color:${nc};">${nvLabel}</span></td>'
        '</tr>`;\n'
        '  });\n'
        '  if(sorted.length>MAX_ROWS) rows += `<tr><td colspan="7" style="text-align:center;color:#e67e22;padding:8px;font-size:10px;">... mostrando ${MAX_ROWS} de ${sorted.length} linhas / 显示${MAX_ROWS}/${sorted.length}行</td></tr>`;\n'
        '  document.getElementById("tbody-inf").innerHTML = rows || '
        '"<tr><td colspan=\'7\' style=\'text-align:center;color:#ddd;padding:20px;\'>Nenhum dado para os filtros / 暂无数据</td></tr>";\n'
        '  document.getElementById("table-count").textContent = sorted.length+" IDs";\n'
        '}\n'
        '\n'
        'function setFiltro(key,val){\n'
        '  if(key==="dateFrom"){fDateFrom=val;localStorage.setItem("brdrive_from",val);}\n'
        '  else if(key==="dateTo"){fDateTo=val;localStorage.setItem("brdrive_to",val);}\n'
        '  else if(key==="trans") fTrans=val;\n'
        '  else if(key==="tipo") fTipo=val;\n'
        '  else if(key==="nivel") fNivel=val;\n'
        '  render();\n'
        '}\n'
        '\n'
        'function resetFiltros(){\n'
        '  fDateFrom="' + date_max + '"; fDateTo="' + date_max + '";\n'
        '  fTrans="Todos"; fTipo="Todos"; fNivel="Todos";\n'
        '  document.getElementById("f-date-from").value=fDateFrom;\n'
        '  document.getElementById("f-date-to").value=fDateTo;\n'
        '  document.getElementById("f-trans").value="Todos";\n'
        '  document.getElementById("f-tipo").value="Todos";\n'
        '  document.getElementById("f-nivel").value="Todos";\n'
        '  render();\n'
        '}\n'
        '\n'
        'document.addEventListener("DOMContentLoaded",function(){\n'
        '  var fi=document.getElementById("f-date-from"),ti=document.getElementById("f-date-to");\n'
        '  if(fi)fi.value=fDateFrom;\n'
        '  if(ti)ti.value=fDateTo;\n'
        '  render();\n'
        '});\n'
        '</script>'
    )

    return pagina_html(nav, conteudo + js, "Inf. Pontualidade")


def pg_pont_motoristas(df):
    base_cols = ['DATA', 'MÊS', 'Transportador', 'CONDUTOR']
    opt_cols  = ['Semana Nome Nova', 'PONTUALIDADE SAÍDA', 'PONTUALIDADE CHEGADA']
    cols_needed = base_cols + [c for c in opt_cols if c in df.columns]

    df_js = df[[c for c in cols_needed]].copy()
    df_js['DATA'] = df_js['DATA'].dt.strftime('%Y-%m-%d')
    df_js = df_js.fillna('')

    rename_map = {
        'DATA': 'd', 'MÊS': 'm', 'Semana Nome Nova': 'sem',
        'Transportador': 't', 'CONDUTOR': 'c',
        'PONTUALIDADE SAÍDA': 'ps', 'PONTUALIDADE CHEGADA': 'pc'
    }
    df_js = df_js.rename(columns={k: v for k, v in rename_map.items() if k in df_js.columns})
    data_json = df_js.to_json(orient='records', force_ascii=False)

    transportadoras = sorted(df['Transportador'].dropna().unique().tolist())
    condutores      = sorted(df['CONDUTOR'].dropna().unique().tolist())
    date_min = df['DATA'].min().strftime('%Y-%m-%d') if not df.empty else ''
    date_max = df['DATA'].max().strftime('%Y-%m-%d') if not df.empty else ''

    opts_trans = '<option value="Todos">Todos</option>' + ''.join(
        f'<option value="{t}">{t}</option>' for t in transportadoras)
    opts_cond  = '<option value="Todos">Todos</option>' + ''.join(
        f'<option value="{c}">{c}</option>' for c in condutores)

    sel_style = 'width:100%;background:#0e0808;border:1px solid rgba(142,68,173,0.4);color:#fff;border-radius:4px;padding:5px 6px;font-size:11px;'
    inp_style = 'width:100%;background:#0e0808;border:1px solid rgba(142,68,173,0.4);color:#fff;border-radius:4px;padding:5px 6px;font-size:11px;color-scheme:dark;'
    sel_hdr   = 'background:#1a1010;border:1px solid rgba(142,68,173,0.5);color:#fff;border-radius:4px;padding:5px 10px;font-size:11px;min-width:220px;'

    filtro_html = f"""
    <div style="width:210px;flex-shrink:0;">
      <div style="background:#1a1010;border-radius:8px;padding:14px;border:1px solid rgba(142,68,173,0.3);position:sticky;top:16px;">
        <div style="color:#8e44ad;font-size:11px;font-weight:700;letter-spacing:1px;margin-bottom:12px;padding-bottom:8px;border-bottom:1px solid rgba(142,68,173,0.3);">FILTROS / 筛选</div>
        <div style="margin-bottom:10px;">
          <div style="color:#ddd;font-size:9px;letter-spacing:1px;margin-bottom:4px;">日期从 | Data De</div>
          <input type="date" id="f-date-from" value="{date_max}" onchange="setFiltro('dateFrom',this.value)" style="{inp_style}margin-bottom:4px;">
          <div style="color:#ddd;font-size:9px;letter-spacing:1px;margin-bottom:4px;margin-top:6px;">日期到 | Data Até</div>
          <input type="date" id="f-date-to" value="{date_max}" onchange="setFiltro('dateTo',this.value)" style="{inp_style}">
        </div>
        <div style="margin-bottom:14px;">
          <div style="color:#ddd;font-size:9px;letter-spacing:1px;margin-bottom:4px;">承运人 | Transportadora</div>
          <select id="f-trans" onchange="setFiltro('trans',this.value)" style="{sel_style}">{opts_trans}</select>
        </div>
        <button onclick="resetFiltros()" style="width:100%;background:rgba(142,68,173,0.2);border:1px solid rgba(142,68,173,0.4);color:#8e44ad;border-radius:4px;padding:7px;font-size:10px;font-weight:700;cursor:pointer;letter-spacing:1px;">RESET / 重置</button>
      </div>
    </div>"""

    nav = navbar("pont_motoristas")
    conteudo = f"""
    <div class="page-header">
      <div class="header-left">
        <div class="header-icon" style="background:#8e44ad;">🚗</div>
        <div>
          <div class="header-title">PONTUALIDADE MOTORISTAS / 司机准时率</div>
          <div class="header-sub">KPI SAÍDA · KPI CHEGADA · KPI GERAL · <span id="periodo-txt">—</span> · USE OS FILTROS PARA NAVEGAR</div>
        </div>
      </div>
      <div style="display:flex;align-items:center;gap:8px;flex:1;justify-content:center;padding:0 16px;">
        <div style="color:#ddd;font-size:10px;font-weight:600;white-space:nowrap;">驾驶员 | Motoristas</div>
        <select id="f-cond" onchange="setFiltro('cond',this.value)" style="{sel_hdr}">{opts_cond}</select>
      </div>
      <div class="header-badges">
        <div class="header-badge"><div class="hb-label">MÊS/月</div><div class="hb-value" id="hb-mes">—</div></div>
        <div class="header-badge"><div class="hb-label">SEMANA/周</div><div class="hb-value" id="hb-sem">—</div></div>
        <div class="header-badge"><div class="hb-label">PERÍODO/期间</div><div class="hb-value" id="hb-periodo">—</div></div>
      </div>
    </div>

    <div style="display:flex;gap:16px;align-items:flex-start;">
      <div style="flex:1;min-width:0;">

        <div class="kpi-grid kpi-grid-5" style="margin-bottom:12px;">
          <div class="kpi-card" style="border-top:3px solid #8e44ad;">
            <div class="kpi-label">TOTAL MOTORISTAS / 司机</div>
            <div class="kpi-value" id="kpi-total">—</div>
            <div class="kpi-sub">condutores no período</div>
          </div>
          <div class="kpi-card" style="border-top:3px solid #27ae60;">
            <div class="kpi-label">KPI PONT. SAÍDA / 出发准时</div>
            <div class="kpi-value" id="kpi-s" style="color:#27ae60;">—</div>
            <div class="kpi-sub" id="kpi-s-sub"></div>
            <div class="progress-bar" style="margin-top:6px;"><div id="kpi-bar-s" class="progress-fill" style="width:0%;background:#27ae60;"></div></div>
          </div>
          <div class="kpi-card" style="border-top:3px solid #27ae60;">
            <div class="kpi-label">KPI PONT. CHEGADA / 到达准时</div>
            <div class="kpi-value" id="kpi-c" style="color:#27ae60;">—</div>
            <div class="kpi-sub" id="kpi-c-sub"></div>
            <div class="progress-bar" style="margin-top:6px;"><div id="kpi-bar-c" class="progress-fill" style="width:0%;background:#27ae60;"></div></div>
          </div>
          <div class="kpi-card" style="border-top:3px solid #27ae60;">
            <div class="kpi-label">KPI GERAL / 综合准时</div>
            <div class="kpi-value" id="kpi-g" style="color:#27ae60;">—</div>
            <div class="kpi-sub">média saída + chegada</div>
            <div class="progress-bar" style="margin-top:6px;"><div id="kpi-bar-g" class="progress-fill" style="width:0%;background:#27ae60;"></div></div>
          </div>
          <div class="kpi-card" style="border-top:3px solid #e74c3c;">
            <div class="kpi-label">ABAIXO META / 未达标</div>
            <div class="kpi-value" id="kpi-abaixo" style="color:#e74c3c;">—</div>
            <div class="kpi-sub">motoristas KPI abaixo 70%</div>
          </div>
        </div>

        <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px;">

          <div class="table-wrap">
            <div class="table-header" style="background:#1A5276;">
              <span>KPI PONT. SAÍDA / 出发准时</span>
              <small>% no prazo · melhor primeiro</small>
            </div>
            <div class="table-scroll" style="max-height:calc(100vh - 360px);">
              <table>
                <thead><tr style="background:#0d2a3a;border-bottom:2px solid rgba(26,82,118,0.5);">
                  <th style="text-align:center;color:#5dade2;width:24px;">#</th>
                  <th style="text-align:left;color:#5dade2;">CONDUTOR/司机</th>
                  <th style="text-align:left;color:#5dade2;">TRANSP.</th>
                  <th style="text-align:center;color:#5dade2;">IDs</th>
                  <th style="text-align:center;color:#5dade2;">ATR.</th>
                  <th style="text-align:center;color:#5dade2;">KPI S.</th>
                </tr></thead>
                <tbody id="tbody-s"></tbody>
                <tfoot><tr id="tfoot-s" style="background:#1A5276;font-size:11px;font-weight:700;"></tr></tfoot>
              </table>
            </div>
          </div>

          <div class="table-wrap">
            <div class="table-header" style="background:#1A5276;">
              <span>KPI PONT. CHEGADA / 到达准时</span>
              <small>% no prazo · melhor primeiro</small>
            </div>
            <div class="table-scroll" style="max-height:calc(100vh - 360px);">
              <table>
                <thead><tr style="background:#0d2a3a;border-bottom:2px solid rgba(26,82,118,0.5);">
                  <th style="text-align:center;color:#5dade2;width:24px;">#</th>
                  <th style="text-align:left;color:#5dade2;">CONDUTOR/司机</th>
                  <th style="text-align:left;color:#5dade2;">TRANSP.</th>
                  <th style="text-align:center;color:#5dade2;">IDs</th>
                  <th style="text-align:center;color:#5dade2;">ATR.</th>
                  <th style="text-align:center;color:#5dade2;">KPI C.</th>
                </tr></thead>
                <tbody id="tbody-c"></tbody>
                <tfoot><tr id="tfoot-c" style="background:#1A5276;font-size:11px;font-weight:700;"></tr></tfoot>
              </table>
            </div>
          </div>

          <div class="table-wrap">
            <div class="table-header" style="background:#6c3483;">
              <span>KPI GERAL / 综合准时</span>
              <small>média saída + chegada · melhor primeiro</small>
            </div>
            <div class="table-scroll" style="max-height:calc(100vh - 360px);">
              <table>
                <thead><tr style="background:#3b1f4a;border-bottom:2px solid rgba(142,68,173,0.5);">
                  <th style="text-align:center;color:#c39bd3;width:24px;">#</th>
                  <th style="text-align:left;color:#c39bd3;">CONDUTOR/司机</th>
                  <th style="text-align:left;color:#c39bd3;">TRANSP.</th>
                  <th style="text-align:center;color:#c39bd3;">IDs</th>
                  <th style="text-align:center;color:#c39bd3;">KPI S.</th>
                  <th style="text-align:center;color:#c39bd3;">KPI C.</th>
                </tr></thead>
                <tbody id="tbody-g"></tbody>
                <tfoot><tr id="tfoot-g" style="background:#6c3483;font-size:11px;font-weight:700;"></tr></tfoot>
              </table>
            </div>
          </div>

        </div>

        <div class="legend-bar" style="margin-top:8px;">
          <div class="legend-item" style="color:#27ae60;"><div class="legend-dot" style="background:#27ae60;"></div>≥85% — META atingida / 达标</div>
          <div class="legend-item" style="color:#e67e22;"><div class="legend-dot" style="background:#e67e22;"></div>70–84% — PARCIAL / 部分</div>
          <div class="legend-item" style="color:#e74c3c;"><div class="legend-dot" style="background:#e74c3c;"></div>&lt;70% — CRÍTICO / 危急</div>
          <div style="margin-left:auto;font-size:9px;color:#888;">KPI = % viagens "No prazo" · meta ≥85% · ATR. = atrasos no período</div>
        </div>

      </div>
      FILTRO_PLACEHOLDER
    </div>"""

    conteudo = conteudo.replace('FILTRO_PLACEHOLDER', filtro_html)

    js = (
        '<script>\n'
        'const RAW = ' + data_json + ';\n'
        'const _DM="' + date_max + '",_DMi="' + date_min + '";\n'
        'const _lsM=localStorage.getItem("brdrive_data_max"),_lsF2=localStorage.getItem("brdrive_from"),_lsT2=localStorage.getItem("brdrive_to");\n'
        'if(_lsM!==_DM||(_lsF2&&_lsF2>_DM)||(_lsT2&&_lsT2<_DMi)){localStorage.removeItem("brdrive_from");localStorage.removeItem("brdrive_to");localStorage.setItem("brdrive_data_max",_DM);}\n'
        'const _lsF=localStorage.getItem("brdrive_from"),_lsT=localStorage.getItem("brdrive_to");\n'
        'let fDateFrom=_lsF||"' + date_max + '";\n'
        'let fDateTo=_lsT||"' + date_max + '";\n'
        'let fTrans = "Todos"; let fCond = "Todos";\n'
        '\n'
        'function cor(v){if(v>=85)return"#27ae60";if(v>=70)return"#e67e22";return"#e74c3c";}\n'
        'function bgCor(v){if(v>=85)return"rgba(39,174,96,0.15)";if(v>=70)return"rgba(230,126,34,0.15)";return"rgba(231,76,60,0.15)";}\n'
        'function pct(a,b){return b>0?(a/b*100):0;}\n'
        '\n'
        'function getFiltrado(){\n'
        '  return RAW.filter(r=>{\n'
        '    if(fDateFrom && r.d < fDateFrom) return false;\n'
        '    if(fDateTo && r.d > fDateTo) return false;\n'
        '    if(fTrans!=="Todos" && r.t!==fTrans) return false;\n'
        '    if(fCond!=="Todos" && r.c!==fCond) return false;\n'
        '    return true;\n'
        '  });\n'
        '}\n'
        '\n'
        'function aggByDriver(data){\n'
        '  const byC = {};\n'
        '  data.forEach(r=>{\n'
        '    if(!byC[r.c]) byC[r.c]={ids:0,lateS:0,lateC:0,okS:0,okC:0,trans:new Set()};\n'
        '    byC[r.c].ids++;\n'
        '    if(r.ps && r.ps!=="No prazo") byC[r.c].lateS++; else byC[r.c].okS++;\n'
        '    if(r.pc && r.pc!=="No prazo") byC[r.c].lateC++; else byC[r.c].okC++;\n'
        '    if(r.t) byC[r.c].trans.add(r.t);\n'
        '  });\n'
        '  return byC;\n'
        '}\n'
        '\n'
        'function transBadges(transSet){\n'
        '  return [...transSet].map(t=>`<span class="trans-badge">${t}</span>`).join("");\n'
        '}\n'
        '\n'
        'function render(){\n'
        '  const data = getFiltrado();\n'
        '  const byC = aggByDriver(data);\n'
        '  const drivers = Object.entries(byC);\n'
        '\n'
        '  const totalIDs  = data.length;\n'
        '  const totalLateS = data.filter(r=>r.ps && r.ps!=="No prazo").length;\n'
        '  const totalLateC = data.filter(r=>r.pc && r.pc!=="No prazo").length;\n'
        '  const totalOkS  = totalIDs - totalLateS;\n'
        '  const totalOkC  = totalIDs - totalLateC;\n'
        '  const gKpiS = pct(totalOkS, totalIDs);\n'
        '  const gKpiC = pct(totalOkC, totalIDs);\n'
        '  const gKpiG = totalIDs>0 ? (gKpiS + gKpiC) / 2 : 0;\n'
        '  const abaixo = drivers.filter(([,v])=>(pct(v.okS,v.ids)+pct(v.okC,v.ids))/2 < 70).length;\n'
        '\n'
        '  document.getElementById("kpi-total").textContent = drivers.length;\n'
        '  const eS=document.getElementById("kpi-s"); eS.textContent=gKpiS.toFixed(1)+"%"; eS.style.color=cor(gKpiS);\n'
        '  document.getElementById("kpi-s-sub").textContent=totalLateS+"x ATR. · "+totalOkS+" OK";\n'
        '  const bS=document.getElementById("kpi-bar-s"); bS.style.width=Math.min(gKpiS,100).toFixed(0)+"%"; bS.style.background=cor(gKpiS);\n'
        '  const eC=document.getElementById("kpi-c"); eC.textContent=gKpiC.toFixed(1)+"%"; eC.style.color=cor(gKpiC);\n'
        '  document.getElementById("kpi-c-sub").textContent=totalLateC+"x ATR. · "+totalOkC+" OK";\n'
        '  const bC=document.getElementById("kpi-bar-c"); bC.style.width=Math.min(gKpiC,100).toFixed(0)+"%"; bC.style.background=cor(gKpiC);\n'
        '  const eG=document.getElementById("kpi-g"); eG.textContent=gKpiG.toFixed(1)+"%"; eG.style.color=cor(gKpiG);\n'
        '  const bG=document.getElementById("kpi-bar-g"); bG.style.width=Math.min(gKpiG,100).toFixed(0)+"%"; bG.style.background=cor(gKpiG);\n'
        '  document.getElementById("kpi-abaixo").textContent = abaixo;\n'
        '\n'
        '  const dates = data.map(r=>r.d).filter(Boolean).sort();\n'
        '  const periodo = dates.length>0 ? dates[0]+" → "+dates[dates.length-1] : "—";\n'
        '  document.getElementById("periodo-txt").textContent = periodo;\n'
        '  document.getElementById("hb-periodo").textContent = periodo;\n'
        '  const meses = [...new Set(data.map(r=>r.m).filter(Boolean))];\n'
        '  document.getElementById("hb-mes").textContent = meses.length>0 ? meses.join("/") : "—";\n'
        '  const sems = [...new Set(data.map(r=>(r.sem||"")).filter(Boolean))];\n'
        '  document.getElementById("hb-sem").textContent = sems.length>0 ? sems[sems.length-1] : "—";\n'
        '\n'
        '  // KPI SAÍDA — melhor primeiro\n'
        '  const sortedS = drivers.slice().sort((a,b)=>pct(b[1].okS,b[1].ids)-pct(a[1].okS,a[1].ids));\n'
        '  let rS="", tfSids=0, tfSlate=0;\n'
        '  sortedS.forEach(([c,v],i)=>{\n'
        '    const ks=pct(v.okS,v.ids); const bgr=i%2===0?"#1f0e0e":"#1a1010";\n'
        '    const atrLbl=v.lateS>0?`<span style="color:#e74c3c;font-weight:700;">${v.lateS}x</span>&nbsp;<span style="color:#e74c3c;font-size:9px;">ATR.</span>`:`<span style="color:#27ae60;font-size:9px;font-weight:700;">OK</span>`;\n'
        '    tfSids+=v.ids; tfSlate+=v.lateS;\n'
        '    rS+=`<tr style="background:${bgr};">'
        '<td style="text-align:center;color:#888;font-size:10px;">${i+1}</td>'
        '<td style="color:#5dade2;font-weight:600;font-size:10px;">${c}</td>'
        '<td style="font-size:10px;">${transBadges(v.trans)}</td>'
        '<td style="text-align:center;color:#f1c40f;font-weight:700;">${v.ids}</td>'
        '<td style="text-align:center;">${atrLbl}</td>'
        '<td style="text-align:center;"><span style="background:${bgCor(ks)};color:${cor(ks)};padding:2px 8px;border-radius:20px;font-size:10px;font-weight:700;">${ks.toFixed(1)}%</span></td>'
        '</tr>`;\n'
        '  });\n'
        '  document.getElementById("tbody-s").innerHTML = rS || "<tr><td colspan=\'6\' style=\'text-align:center;color:#ddd;padding:16px;\'>Sem dados / 暂无数据</td></tr>";\n'
        '  document.getElementById("tfoot-s").innerHTML=`<td colspan="3" style="font-weight:700;">TOTAL</td>'
        '<td style="text-align:center;font-weight:700;">${tfSids}</td>'
        '<td style="text-align:center;font-weight:700;color:#e74c3c;">${tfSlate}x ATR.</td>'
        '<td style="text-align:center;font-weight:700;">${pct(tfSids-tfSlate,tfSids).toFixed(1)}%</td>`;\n'
        '\n'
        '  // KPI CHEGADA — melhor primeiro\n'
        '  const sortedC = drivers.slice().sort((a,b)=>pct(b[1].okC,b[1].ids)-pct(a[1].okC,a[1].ids));\n'
        '  let rC="", tfCids=0, tfClate=0;\n'
        '  sortedC.forEach(([c,v],i)=>{\n'
        '    const kc=pct(v.okC,v.ids); const bgr=i%2===0?"#1f0e0e":"#1a1010";\n'
        '    const atrLbl=v.lateC>0?`<span style="color:#e74c3c;font-weight:700;">${v.lateC}x</span>&nbsp;<span style="color:#e74c3c;font-size:9px;">ATR.</span>`:`<span style="color:#27ae60;font-size:9px;font-weight:700;">OK</span>`;\n'
        '    tfCids+=v.ids; tfClate+=v.lateC;\n'
        '    rC+=`<tr style="background:${bgr};">'
        '<td style="text-align:center;color:#888;font-size:10px;">${i+1}</td>'
        '<td style="color:#5dade2;font-weight:600;font-size:10px;">${c}</td>'
        '<td style="font-size:10px;">${transBadges(v.trans)}</td>'
        '<td style="text-align:center;color:#f1c40f;font-weight:700;">${v.ids}</td>'
        '<td style="text-align:center;">${atrLbl}</td>'
        '<td style="text-align:center;"><span style="background:${bgCor(kc)};color:${cor(kc)};padding:2px 8px;border-radius:20px;font-size:10px;font-weight:700;">${kc.toFixed(1)}%</span></td>'
        '</tr>`;\n'
        '  });\n'
        '  document.getElementById("tbody-c").innerHTML = rC || "<tr><td colspan=\'6\' style=\'text-align:center;color:#ddd;padding:16px;\'>Sem dados / 暂无数据</td></tr>";\n'
        '  document.getElementById("tfoot-c").innerHTML=`<td colspan="3" style="font-weight:700;">TOTAL</td>'
        '<td style="text-align:center;font-weight:700;">${tfCids}</td>'
        '<td style="text-align:center;font-weight:700;color:#e74c3c;">${tfClate}x ATR.</td>'
        '<td style="text-align:center;font-weight:700;">${pct(tfCids-tfClate,tfCids).toFixed(1)}%</td>`;\n'
        '\n'
        '  // KPI GERAL — melhor primeiro\n'
        '  const sortedG = drivers.slice().sort((a,b)=>{\n'
        '    const ga=(pct(a[1].okS,a[1].ids)+pct(a[1].okC,a[1].ids))/2;\n'
        '    const gb=(pct(b[1].okS,b[1].ids)+pct(b[1].okC,b[1].ids))/2;\n'
        '    return gb-ga;\n'
        '  });\n'
        '  let rG="", tfGids=0;\n'
        '  sortedG.forEach(([c,v],i)=>{\n'
        '    const ks=pct(v.okS,v.ids); const kc=pct(v.okC,v.ids);\n'
        '    const bgr=i%2===0?"#12091a":"#0e0616";\n'
        '    tfGids+=v.ids;\n'
        '    rG+=`<tr style="background:${bgr};">'
        '<td style="text-align:center;color:#888;font-size:10px;">${i+1}</td>'
        '<td style="color:#5dade2;font-weight:600;font-size:10px;">${c}</td>'
        '<td style="font-size:10px;">${transBadges(v.trans)}</td>'
        '<td style="text-align:center;color:#f1c40f;font-weight:700;">${v.ids}</td>'
        '<td style="text-align:center;"><span style="background:${bgCor(ks)};color:${cor(ks)};padding:2px 6px;border-radius:20px;font-size:10px;font-weight:700;">${ks.toFixed(1)}%</span></td>'
        '<td style="text-align:center;"><span style="background:${bgCor(kc)};color:${cor(kc)};padding:2px 6px;border-radius:20px;font-size:10px;font-weight:700;">${kc.toFixed(1)}%</span></td>'
        '</tr>`;\n'
        '  });\n'
        '  document.getElementById("tbody-g").innerHTML = rG || "<tr><td colspan=\'6\' style=\'text-align:center;color:#ddd;padding:16px;\'>Sem dados / 暂无数据</td></tr>";\n'
        '  document.getElementById("tfoot-g").innerHTML=`<td colspan="3" style="font-weight:700;">TOTAL</td>'
        '<td style="text-align:center;font-weight:700;">${tfGids}</td>'
        '<td style="text-align:center;font-weight:700;">${gKpiS.toFixed(1)}% S.</td>'
        '<td style="text-align:center;font-weight:700;">${gKpiC.toFixed(1)}% C.</td>`;\n'
        '}\n'
        '\n'
        'function setFiltro(key,val){\n'
        '  if(key==="dateFrom"){fDateFrom=val;localStorage.setItem("brdrive_from",val);}\n'
        '  else if(key==="dateTo"){fDateTo=val;localStorage.setItem("brdrive_to",val);}\n'
        '  else if(key==="trans") fTrans=val;\n'
        '  else if(key==="cond") fCond=val;\n'
        '  render();\n'
        '}\n'
        '\n'
        'function resetFiltros(){\n'
        '  fDateFrom="' + date_max + '"; fDateTo="' + date_max + '";\n'
        '  fTrans="Todos"; fCond="Todos";\n'
        '  document.getElementById("f-date-from").value=fDateFrom;\n'
        '  document.getElementById("f-date-to").value=fDateTo;\n'
        '  document.getElementById("f-trans").value="Todos";\n'
        '  document.getElementById("f-cond").value="Todos";\n'
        '  render();\n'
        '}\n'
        '\n'
        'document.addEventListener("DOMContentLoaded",function(){\n'
        '  var fi=document.getElementById("f-date-from"),ti=document.getElementById("f-date-to");\n'
        '  if(fi)fi.value=fDateFrom;\n'
        '  if(ti)ti.value=fDateTo;\n'
        '  render();\n'
        '});\n'
        '</script>'
    )

    return pagina_html(nav, conteudo + js, "Pontualidade Motoristas")


def pg_pont_semanal(df):
    # ── Prepara dados ──────────────────────────────────────────────────────
    df2 = df.copy()
    iso = df2['DATA'].dt.isocalendar()
    df2['semana_num'] = iso.week.astype(int)
    df2['ano']        = iso.year.astype(int)
    df2['semana']     = 'W' + df2['semana_num'].astype(str).str.zfill(2) + ' - ' + df2['ano'].astype(str)
    df2['semana_ord'] = df2['ano'] * 100 + df2['semana_num']
    df2['dow']        = df2['DATA'].dt.dayofweek
    df2['data_str']   = df2['DATA'].dt.strftime('%d/%m')

    # Normaliza pontualidade (garante "No prazo")
    for col in ['PONTUALIDADE SAÍDA', 'PONTUALIDADE CHEGADA']:
        if col in df2.columns:
            df2[col] = df2[col].fillna('').str.strip()

    col_ps = 'PONTUALIDADE SAÍDA' if 'PONTUALIDADE SAÍDA' in df2.columns else None
    col_pc = 'PONTUALIDADE CHEGADA' if 'PONTUALIDADE CHEGADA' in df2.columns else None

    transportadoras = sorted(df2['Transportador'].dropna().unique().tolist())
    tipos = ["Coleta Base", "Coleta PA", "Devolução", "Secundaria"]

    semana_max_ord = int(df2['semana_ord'].max())
    semana_max_lbl = df2.loc[df2['semana_ord'] == semana_max_ord, 'semana'].iloc[0]

    semanas_ord = df2[['semana','semana_ord']].drop_duplicates().sort_values('semana_ord')
    opts_sem = ''.join(
        f'<option value="{r.semana_ord}" {"selected" if r.semana_ord == semana_max_ord else ""}>{r.semana}</option>'
        for _, r in semanas_ord.iterrows()
    )
    opts_trans = '<option value="Todos">Todos</option>' + ''.join(
        f'<option value="{t}">{t}</option>' for t in transportadoras)
    opts_tipo = '<option value="Todos">Todos</option>' + ''.join(
        f'<option value="{t}">{t}</option>' for t in tipos)

    # JSON compacto: semana_ord, dow, data_str, transportador, tipo, pont_saida, pont_chegada
    cols_js = ['semana_ord', 'dow', 'data_str', 'Transportador', 'TIPO DE OPERAÇÃO']
    if col_ps: cols_js.append(col_ps)
    if col_pc: cols_js.append(col_pc)
    df_js = df2[cols_js].fillna('')
    new_names = ['wo', 'dow', 'dt', 't', 'op'] + (['ps'] if col_ps else []) + (['pc'] if col_pc else [])
    df_js.columns = new_names
    data_json = df_js.to_json(orient='records', force_ascii=False)

    nav = navbar("pont_semanal")

    js = (
        '<script>\n'
        'const RAW=' + data_json + ';\n'
        'let fSem=' + str(semana_max_ord) + ',fTrans="Todos",fTipo="Todos";\n'
        'const DIAS=[["SEG","周一"],["TER","周二"],["QUA","周三"],["QUI","周四"],["SEX","周五"],["SÁB","周六"],["DOM","周日"]];\n'
        'const DAY_COR=["#e67e22","#5dade2","#5dade2","#5dade2","#5dade2","#aaa","#e74c3c"];\n'
        '\n'
        'function cor(v){return v>=85?"#27ae60":v>=70?"#e67e22":"#e74c3c";}\n'
        'function pct(a,b){return b>0?Math.round(a/b*1000)/10:0;}\n'
        'function semAnt(wo){const yr=Math.floor(wo/100),wk=wo%100;return wk<=1?(yr-1)*100+52:yr*100+(wk-1);}\n'
        '\n'
        'function getDias(woFilter){\n'
        '  const dias={};\n'
        '  RAW.forEach(r=>{\n'
        '    if(r.wo!==woFilter) return;\n'
        '    if(fTrans!=="Todos"&&r.t!==fTrans) return;\n'
        '    if(fTipo!=="Todos"&&r.op!==fTipo) return;\n'
        '    if(!dias[r.dow]) dias[r.dow]={ids:0,okS:0,okC:0,dt:""};\n'
        '    dias[r.dow].ids++;\n'
        '    if(r.ps==="No prazo") dias[r.dow].okS++;\n'
        '    if(r.pc==="No prazo") dias[r.dow].okC++;\n'
        '    if(r.dt) dias[r.dow].dt=r.dt;\n'
        '  });\n'
        '  return dias;\n'
        '}\n'
        '\n'
        'function buildChart(dAt,dAnt,campo){\n'
        '  const H=155;\n'
        '  const metaY=Math.round((1-0.85)*H);\n'
        '  let html=\'<div style="display:flex;gap:5px;">\';\n'
        '  let cards=\'<div style="display:flex;gap:5px;margin-top:12px;">\';\n'
        '  for(let d=0;d<7;d++){\n'
        '    const vA=dAt[d]||{ids:0,okS:0,okC:0,dt:""};\n'
        '    const vP=dAnt[d]||{ids:0,okS:0,okC:0,dt:""};\n'
        '    const valA=campo==="s"?pct(vA.okS,vA.ids):pct(vA.okC,vA.ids);\n'
        '    const valP=campo==="s"?pct(vP.okS,vP.ids):pct(vP.okC,vP.ids);\n'
        '    const hA=vA.ids>0?Math.max(6,Math.round(valA/100*H)):0;\n'
        '    const hP=vP.ids>0?Math.max(6,Math.round(valP/100*H)):0;\n'
        '    const cA=cor(valA);\n'
        '    const diff=vP.ids>0?valA-valP:null;\n'
        '    const tStr=diff!==null?(diff>0.5?"▲ +":diff<-0.5?"▼ ":"→ ")+Math.abs(diff).toFixed(1)+"%":"";\n'
        '    const tCor=diff===null?"#aaa":diff>0.5?"#27ae60":diff<-0.5?"#e74c3c":"#aaa";\n'
        '    const dc=DAY_COR[d]; const dn=DIAS[d][0]; const dz=DIAS[d][1];\n'
        '    const dt=vA.dt||vP.dt||"";\n'
        '    const lbl=vA.ids>0?valA.toFixed(1)+"%":"—";\n'
        '    const cardBg=valA>=85?"rgba(39,174,96,0.1)":valA>=70?"rgba(230,126,34,0.1)":"rgba(231,76,60,0.1)";\n'
        '    const cardBd=valA>=85?"rgba(39,174,96,0.35)":valA>=70?"rgba(230,126,34,0.35)":"rgba(231,76,60,0.35)";\n'
        '\n'
        '    html+=`<div style="flex:1;display:flex;flex-direction:column;align-items:center;min-width:80px;">`\n'
        '         +`<div style="font-size:10px;color:${tCor};font-weight:700;height:18px;text-align:center;">${tStr}</div>`\n'
        '         +`<div style="font-size:15px;color:${cA};font-weight:800;height:22px;text-align:center;line-height:22px;">${lbl}</div>`\n'
        '         +`<div style="width:100%;height:${H}px;position:relative;display:flex;align-items:flex-end;justify-content:center;gap:4px;padding:0 14px;box-sizing:border-box;">`\n'
        '         +`  <div style="position:absolute;top:${metaY}px;left:0;right:0;border-top:1.5px dashed rgba(255,255,255,0.3);pointer-events:none;z-index:2;"></div>`\n'
        '         +`  <div style="position:absolute;top:${metaY-11}px;right:6px;font-size:8px;color:rgba(255,255,255,0.75);z-index:2;">85%</div>`\n'
        '         +`  <div style="width:38%;height:${hP}px;background:#1a5276;border-radius:4px 4px 0 0;opacity:0.85;flex-shrink:0;"></div>`\n'
        '         +`  <div style="width:38%;height:${hA}px;background:${cA};border-radius:4px 4px 0 0;flex-shrink:0;"></div>`\n'
        '         +"</div>"\n'
        '         +`<div style="font-size:11px;color:${dc};font-weight:700;margin-top:7px;text-align:center;">${dn}</div>`\n'
        '         +`<div style="font-size:9px;color:rgba(255,255,255,0.6);text-align:center;">${dz}</div>`\n'
        '         +`<div style="font-size:9px;color:#ccc;text-align:center;">${dt}</div>`\n'
        '         +"</div>";\n'
        '\n'
        '    const antStr=vP.ids>0?`ant: ${valP.toFixed(1)}%`:"";\n'
        '    cards+=`<div style="flex:1;min-width:90px;background:${cardBg};border:1px solid ${cardBd};border-radius:8px;padding:10px 8px;text-align:center;">`\n'
        '          +`<div style="font-size:9px;color:${dc};font-weight:700;margin-bottom:5px;">${dn}/${dz} ${dt}</div>`\n'
        '          +`<div style="font-size:22px;font-weight:900;color:${cA};line-height:1.1;">${lbl}</div>`\n'
        '          +`<div style="font-size:10px;color:${tCor};font-weight:700;margin-top:5px;">${tStr}</div>`\n'
        '          +`<div style="font-size:9px;color:#ddd;margin-top:3px;">${antStr}</div>`\n'
        '          +`<div style="font-size:8px;color:#ccc;margin-top:2px;">${vA.ids>0?vA.ids+" viagens":""}</div>`\n'
        '          +"</div>";\n'
        '  }\n'
        '  html+="</div>"; cards+="</div>";\n'
        '  return html+cards;\n'
        '}\n'
        '\n'
        'function render(){\n'
        '  const ant=semAnt(fSem);\n'
        '  const dAt=getDias(fSem), dAnt=getDias(ant);\n'
        '  const lbl=r=>{ const yr=Math.floor(r/100),wk=r%100; return "W"+String(wk).padStart(2,"0")+" - "+yr; };\n'
        '  document.getElementById("kpi-sem").textContent=lbl(fSem);\n'
        '  document.getElementById("kpi-ant").textContent=lbl(ant);\n'
        '  let tS=0,oS=0,tC=0,oC=0;\n'
        '  for(let d=0;d<7;d++){\n'
        '    const v=dAt[d]; if(!v) continue;\n'
        '    tS+=v.ids; oS+=v.okS; tC+=v.ids; oC+=v.okC;\n'
        '  }\n'
        '  const mS=pct(oS,tS), mC=pct(oC,tC);\n'
        '  const elMS=document.getElementById("kpi-media-s"); elMS.textContent=mS.toFixed(1)+"%"; elMS.style.color=cor(mS);\n'
        '  const elMC=document.getElementById("kpi-media-c"); elMC.textContent=mC.toFixed(1)+"%"; elMC.style.color=cor(mC);\n'
        '  document.getElementById("bloco-saida").innerHTML=buildChart(dAt,dAnt,"s");\n'
        '  document.getElementById("bloco-chegada").innerHTML=buildChart(dAt,dAnt,"c");\n'
        '}\n'
        '\n'
        'function setFiltro(k,v){\n'
        '  if(k==="sem") fSem=parseInt(v);\n'
        '  else if(k==="trans") fTrans=v;\n'
        '  else if(k==="tipo") fTipo=v;\n'
        '  render();\n'
        '}\n'
        'function resetFiltros(){\n'
        '  fSem=' + str(semana_max_ord) + '; fTrans="Todos"; fTipo="Todos";\n'
        '  document.getElementById("f-sem").value=fSem;\n'
        '  document.getElementById("f-trans").value="Todos";\n'
        '  document.getElementById("f-tipo").value="Todos";\n'
        '  render();\n'
        '}\n'
        'document.addEventListener("DOMContentLoaded",render);\n'
        '</script>'
    )

    sel_style2 = 'width:100%;background:#0e0808;border:1px solid rgba(26,82,118,0.4);color:#fff;border-radius:4px;padding:5px 6px;font-size:11px;'
    filtro2 = f"""
    <div style="width:210px;flex-shrink:0;">
      <div style="background:#1a1010;border-radius:8px;padding:14px;border:1px solid rgba(26,82,118,0.3);position:sticky;top:16px;">
        <div style="color:#2980b9;font-size:11px;font-weight:700;letter-spacing:1px;margin-bottom:12px;padding-bottom:8px;border-bottom:1px solid rgba(26,82,118,0.3);">FILTROS / 筛选</div>
        <div style="margin-bottom:10px;">
          <div style="color:#ddd;font-size:9px;letter-spacing:1px;margin-bottom:4px;">周 | Semana</div>
          <select id="f-sem" onchange="setFiltro('sem',this.value)" style="{sel_style2}">{opts_sem}</select>
        </div>
        <div style="margin-bottom:10px;">
          <div style="color:#ddd;font-size:9px;letter-spacing:1px;margin-bottom:4px;">承运商 | Transportador</div>
          <select id="f-trans" onchange="setFiltro('trans',this.value)" style="{sel_style2}">{opts_trans}</select>
        </div>
        <div style="margin-bottom:14px;">
          <div style="color:#ddd;font-size:9px;letter-spacing:1px;margin-bottom:4px;">运营类型 | Tipo Op.</div>
          <select id="f-tipo" onchange="setFiltro('tipo',this.value)" style="{sel_style2}">{opts_tipo}</select>
        </div>
        <button onclick="resetFiltros()" style="width:100%;background:rgba(26,82,118,0.2);border:1px solid rgba(26,82,118,0.4);color:#2980b9;border-radius:4px;padding:7px;font-size:10px;font-weight:700;cursor:pointer;letter-spacing:1px;">RESET / 重置</button>
        <div style="margin-top:14px;padding-top:10px;border-top:1px solid rgba(255,255,255,0.07);">
          <div style="color:#ddd;font-size:9px;margin-bottom:6px;letter-spacing:1px;">LEGENDA / 图例</div>
          <div style="display:flex;align-items:center;gap:6px;margin-bottom:4px;"><div style="width:10px;height:10px;border-radius:2px;background:#27ae60;"></div><span style="font-size:9px;color:#ddd;">≥85% / 达标</span></div>
          <div style="display:flex;align-items:center;gap:6px;margin-bottom:4px;"><div style="width:10px;height:10px;border-radius:2px;background:#e67e22;"></div><span style="font-size:9px;color:#ddd;">70–84% / 注意</span></div>
          <div style="display:flex;align-items:center;gap:6px;margin-bottom:4px;"><div style="width:10px;height:10px;border-radius:2px;background:#e74c3c;"></div><span style="font-size:9px;color:#ddd;">&lt;70% / 危急</span></div>
          <div style="display:flex;align-items:center;gap:6px;"><div style="width:10px;height:10px;border-radius:2px;background:#1a5276;"></div><span style="font-size:9px;color:#ddd;">Semana anterior / 上周</span></div>
        </div>
      </div>
    </div>"""

    conteudo2 = f"""
    <div class="page-header">
      <div class="header-left">
        <div class="header-icon" style="background:#1A5276;">⏱</div>
        <div>
          <div class="header-title">PONTUALIDADE SEMANAL / 准时率周度绩效</div>
          <div class="header-sub">SAÍDA E CHEGADA NO PRAZO · 出发与到达准时 · COLORIDO=ATUAL · AZUL=ANTERIOR</div>
        </div>
      </div>
      <div class="header-badges">
        <div class="header-badge"><div class="hb-label">SEMANA/周</div><div class="hb-value" id="kpi-sem" style="color:#2980b9;">{semana_max_lbl}</div></div>
        <div class="header-badge"><div class="hb-label">ANTERIOR/上周</div><div class="hb-value" id="kpi-ant" style="color:#5d6d7e;">—</div></div>
        <div class="header-badge"><div class="hb-label">PONT. SAÍDA/出发准时</div><div class="hb-value" id="kpi-media-s">—</div></div>
        <div class="header-badge"><div class="hb-label">PONT. CHEGADA/到达准时</div><div class="hb-value" id="kpi-media-c">—</div></div>
      </div>
    </div>

    <div style="display:flex;gap:12px;align-items:flex-start;">
      <div style="flex:1;min-width:0;">

        <div class="table-wrap" style="margin-bottom:12px;">
          <div style="background:#1A5276;padding:8px 14px;display:flex;justify-content:space-between;align-items:center;">
            <span style="color:#fff;font-size:11px;font-weight:700;">▶ PONTUALIDADE SAÍDA / 出发准时率 — % No Prazo por dia</span>
            <small style="color:rgba(255,255,255,0.95);font-size:9px;">azul/蓝=semana anterior · colorido/彩=semana atual</small>
          </div>
          <div style="padding:12px 14px;" id="bloco-saida"></div>
        </div>

        <div class="table-wrap" style="margin-bottom:12px;">
          <div style="background:#8e44ad;padding:8px 14px;display:flex;justify-content:space-between;align-items:center;">
            <span style="color:#fff;font-size:11px;font-weight:700;">▶ PONTUALIDADE CHEGADA / 到达准时率 — % No Prazo por dia</span>
            <small style="color:rgba(255,255,255,0.95);font-size:9px;">azul/蓝=semana anterior · colorido/彩=semana atual</small>
          </div>
          <div style="padding:12px 14px;" id="bloco-chegada"></div>
        </div>

        <div class="legend-bar">
          <div class="legend-item"><div class="legend-dot" style="background:#27ae60;"></div>≥85% no prazo/达标</div>
          <div class="legend-item"><div class="legend-dot" style="background:#e67e22;"></div>70–84% atenção/注意</div>
          <div class="legend-item"><div class="legend-dot" style="background:#e74c3c;"></div>&lt;70% crítico/危急</div>
          <div class="legend-item"><div class="legend-dot" style="background:#1a5276;"></div>Semana anterior/上周</div>
          <div style="margin-left:auto;font-size:9px;color:#888;">KPI = % viagens "No prazo" · meta ≥85%</div>
        </div>
      </div>
      {filtro2}
    </div>"""

    return pagina_html(nav, conteudo2 + js, "Pontualidade Semanal")


def pg_pont_mensal(df):
    df2 = df.copy()
    df2['mes_num'] = df2['DATA'].dt.month
    df2['ano']     = df2['DATA'].dt.year
    df2['mes_ord'] = df2['ano'] * 100 + df2['mes_num']

    tipos   = sorted(df2['TIPO DE OPERAÇÃO'].dropna().unique().tolist())
    ano_max = int(df2['ano'].max())
    anos    = sorted(df2['ano'].dropna().unique().tolist(), reverse=True)

    col_ps = 'PONTUALIDADE SAÍDA' if 'PONTUALIDADE SAÍDA' in df2.columns else None
    col_pc = 'PONTUALIDADE CHEGADA' if 'PONTUALIDADE CHEGADA' in df2.columns else None
    for col in [col_ps, col_pc]:
        if col: df2[col] = df2[col].fillna('').str.strip()

    cols_js = ['mes_ord', 'mes_num', 'Transportador', 'TIPO DE OPERAÇÃO']
    if col_ps: cols_js.append(col_ps)
    if col_pc: cols_js.append(col_pc)
    df_js = df2[cols_js].fillna('')
    new_names = ['mo', 'mn', 't', 'op'] + (['ps'] if col_ps else []) + (['pc'] if col_pc else [])
    df_js.columns = new_names
    data_json = df_js.to_json(orient='records', force_ascii=False)

    ss = 'background:#0e0808;border:1px solid rgba(26,82,118,0.3);color:#fff;border-radius:4px;padding:5px 8px;font-size:11px;'
    opts_ano  = '<option value="Todos">Todos</option>' + ''.join(
        f'<option value="{a}" {"selected" if a==ano_max else ""}>{a}</option>' for a in anos)
    opts_tipo = '<option value="Todos">Todos</option>' + ''.join(
        f'<option value="{t}">{t}</option>' for t in tipos)

    nav = navbar("pont_mensal")

    conteudo = f"""
    <div class="page-header">
      <div class="header-left">
        <div class="header-icon" style="background:#1A5276;">📆</div>
        <div>
          <div class="header-title">PONTUALIDADE — COMPARATIVO MENSAL / 准时率月度对比</div>
          <div class="header-sub">Saída e Chegada No Prazo · Jan-Dez · por Transportadora / 出发与到达准时 · 1-12月 · 按承运商</div>
        </div>
      </div>
      <div style="display:flex;gap:10px;align-items:flex-end;">
        <div>
          <div style="color:#ddd;font-size:9px;letter-spacing:1px;margin-bottom:4px;">操作类型 | Tipo de Operação</div>
          <select id="f-tipo" onchange="setFiltro('tipo',this.value)" style="{ss}">{opts_tipo}</select>
        </div>
        <div>
          <div style="color:#ddd;font-size:9px;letter-spacing:1px;margin-bottom:4px;">年份 | Ano</div>
          <select id="f-ano" onchange="setFiltro('ano',this.value)" style="{ss}">{opts_ano}</select>
        </div>
      </div>
    </div>

    <!-- KPI cards por mês -->
    <div style="overflow-x:auto;margin-bottom:14px;">
      <div id="kpi-cards" style="display:flex;gap:8px;padding-bottom:4px;min-width:max-content;"></div>
    </div>

    <!-- 2x2 grid -->
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:12px;">
      <div class="table-wrap">
        <div style="background:#1A5276;padding:8px 12px;">
          <span style="color:#fff;font-size:10px;font-weight:700;">▶ PONT. SAÍDA/出发准时率 — Evolução Mensal/月度趋势</span>
        </div>
        <div style="padding:12px;" id="chart-saida-evol"></div>
      </div>
      <div class="table-wrap">
        <div style="background:#1A5276;padding:8px 12px;display:flex;justify-content:space-between;align-items:center;">
          <span style="color:#fff;font-size:10px;font-weight:700;">▶ PONT. SAÍDA/出发准时率 — Por Transportadora/按承运商</span>
          <span id="lbl-ts" style="color:rgba(255,255,255,0.9);font-size:9px;"></span>
        </div>
        <div style="padding:12px;" id="chart-saida-trans"></div>
      </div>
      <div class="table-wrap">
        <div style="background:#8e44ad;padding:8px 12px;">
          <span style="color:#fff;font-size:10px;font-weight:700;">▶ PONT. CHEGADA/到达准时率 — Evolução Mensal/月度趋势</span>
        </div>
        <div style="padding:12px;" id="chart-chegada-evol"></div>
      </div>
      <div class="table-wrap">
        <div style="background:#8e44ad;padding:8px 12px;display:flex;justify-content:space-between;align-items:center;">
          <span style="color:#fff;font-size:10px;font-weight:700;">▶ PONT. CHEGADA/到达准时率 — Por Transportadora/按承运商</span>
          <span id="lbl-tc" style="color:rgba(255,255,255,0.9);font-size:9px;"></span>
        </div>
        <div style="padding:12px;" id="chart-chegada-trans"></div>
      </div>
    </div>

    <div class="legend-bar">
      <div class="legend-item"><div class="legend-dot" style="background:#27ae60;"></div>≥85% no prazo/达标</div>
      <div class="legend-item"><div class="legend-dot" style="background:#e67e22;"></div>70–84% atenção/注意</div>
      <div class="legend-item"><div class="legend-dot" style="background:#e74c3c;"></div>&lt;70% crítico/危急</div>
      <div class="legend-item"><div style="width:18px;height:3px;background:#f39c12;border-radius:2px;margin-right:2px;"></div>Tendência/趋势</div>
    </div>"""

    js = (
        '<script>\n'
        'const RAW=' + data_json + ';\n'
        'let fAno=' + str(ano_max) + ',fTipo="Todos";\n'
        'const MESES=[["JAN","1月"],["FEV","2月"],["MAR","3月"],["ABR","4月"],["MAI","5月"],["JUN","6月"],'
        '["JUL","7月"],["AGO","8月"],["SET","9月"],["OUT","10月"],["NOV","11月"],["DEZ","12月"]];\n'
        'function cor(v){return v>=85?"#27ae60":v>=70?"#e67e22":"#e74c3c";}\n'
        'function pct(a,b){return b>0?Math.round(a/b*1000)/10:0;}\n'
        '\n'
        'function agg(){\n'
        '  const mes={},trans={};\n'
        '  RAW.forEach(r=>{\n'
        '    if(fAno!=="Todos"&&Math.floor(r.mo/100)!==fAno) return;\n'
        '    if(fTipo!=="Todos"&&r.op!==fTipo) return;\n'
        '    const m=r.mo%100;\n'
        '    if(!mes[m]) mes[m]={ids:0,okS:0,okC:0};\n'
        '    mes[m].ids++;\n'
        '    if(r.ps==="No prazo") mes[m].okS++;\n'
        '    if(r.pc==="No prazo") mes[m].okC++;\n'
        '    if(r.t){\n'
        '      if(!trans[r.t]) trans[r.t]={};\n'
        '      if(!trans[r.t][m]) trans[r.t][m]={ids:0,okS:0,okC:0};\n'
        '      trans[r.t][m].ids++;\n'
        '      if(r.ps==="No prazo") trans[r.t][m].okS++;\n'
        '      if(r.pc==="No prazo") trans[r.t][m].okC++;\n'
        '    }\n'
        '  });\n'
        '  return {mes,trans};\n'
        '}\n'
        '\n'
        'function buildKPI(mes){\n'
        '  const mks=Object.keys(mes).map(Number).sort((a,b)=>a-b);\n'
        '  return mks.map((m,i)=>{\n'
        '    const v=mes[m],vP=i>0?mes[mks[i-1]]:null;\n'
        '    const pS=pct(v.okS,v.ids),pC=pct(v.okC,v.ids);\n'
        '    const dS=vP?pS-pct(vP.okS,vP.ids):null,dC=vP?pC-pct(vP.okC,vP.ids):null;\n'
        '    const fD=d=>d===null?"":((d>=0?"▲ +":"▼ ")+Math.abs(d).toFixed(1)+"pp");\n'
        '    const cD=d=>d===null?"#ddd":d>=0?"#27ae60":"#e74c3c";\n'
        '    const pt=MESES[m-1][0],zh=MESES[m-1][1];\n'
        '    return `<div style="background:#1a1010;border:1px solid rgba(255,255,255,0.08);border-radius:8px;padding:10px 14px;min-width:130px;">`\n'
        '          +`<div style="font-size:9px;color:#5dade2;font-weight:700;margin-bottom:6px;">${pt}/${zh} · <span style="color:#ddd;font-weight:400;">${v.ids.toLocaleString()} viagens</span></div>`\n'
        '          +`<div style="display:flex;gap:14px;">`\n'
        '          +`<div><div style="font-size:8px;color:#ddd;margin-bottom:2px;">Pont.S/出发</div>`\n'
        '          +`<div style="font-size:19px;font-weight:900;color:${cor(pS)};line-height:1.1;">${pS.toFixed(1)}%</div>`\n'
        '          +`<div style="font-size:9px;color:${cD(dS)};font-weight:700;">${fD(dS)}</div></div>`\n'
        '          +`<div><div style="font-size:8px;color:#ddd;margin-bottom:2px;">Pont.C/到达</div>`\n'
        '          +`<div style="font-size:19px;font-weight:900;color:${cor(pC)};line-height:1.1;">${pC.toFixed(1)}%</div>`\n'
        '          +`<div style="font-size:9px;color:${cD(dC)};font-weight:700;">${fD(dC)}</div></div>`\n'
        '          +"</div></div>";\n'
        '  }).join("");\n'
        '}\n'
        '\n'
        'function buildEvol(mes,campo){\n'
        '  const H=175,VW=1200,N=12,cw=VW/N;\n'
        '  const m85=H*(1-0.85),m70=H*(1-0.70);\n'
        '  let lbls="",bars="",names="",pts=[],dots="";\n'
        '  for(let m=1;m<=12;m++){\n'
        '    const v=mes[m]||{ids:0,okS:0,okC:0};\n'
        '    const val=campo==="s"?pct(v.okS,v.ids):pct(v.okC,v.ids);\n'
        '    const h=v.ids>0?Math.max(4,Math.round(val/100*H)):0;\n'
        '    const c=cor(val),lbl=v.ids>0?val.toFixed(1)+"%":"—";\n'
        '    const xC=(m-0.5)*cw,yC=H*(1-val/100);\n'
        '    if(v.ids>0){pts.push(`${xC},${yC}`);dots+=`<circle cx="${xC}" cy="${yC}" r="8" fill="#f39c12" stroke="#fff" stroke-width="2.5"/>`;}\n'
        '    lbls+=`<div style="flex:1;text-align:center;font-size:11px;color:${c};font-weight:800;height:18px;line-height:18px;">${lbl}</div>`;\n'
        '    bars+=`<div style="flex:1;display:flex;align-items:flex-end;justify-content:center;height:${H}px;">`\n'
        '         +`<div style="width:60%;height:${h}px;background:${v.ids>0?c:"rgba(255,255,255,0.06)"};border-radius:3px 3px 0 0;"></div>`\n'
        '         +"</div>";\n'
        '    names+=`<div style="flex:1;text-align:center;padding-top:4px;">`\n'
        '           +`<div style="font-size:10px;color:#5dade2;font-weight:700;">${MESES[m-1][0]}</div>`\n'
        '           +`<div style="font-size:8px;color:rgba(255,255,255,0.6);">${MESES[m-1][1]}</div>`\n'
        '           +"</div>";\n'
        '  }\n'
        '  const line=pts.length>1?`<polyline points="${pts.join(" ")}" fill="none" stroke="#f39c12" stroke-width="4" stroke-linejoin="round" stroke-linecap="round"/>`:""; \n'
        '  const svg=`<svg style="position:absolute;top:0;left:0;width:100%;height:${H}px;pointer-events:none;"`\n'
        '            +` viewBox="0 0 ${VW} ${H}" preserveAspectRatio="none">`\n'
        '            +`<line x1="0" y1="${m85}" x2="${VW}" y2="${m85}" stroke="rgba(255,255,255,0.35)" stroke-width="2" stroke-dasharray="8,5"/>`\n'
        '            +`<text x="10" y="${m85-5}" fill="rgba(255,255,255,0.55)" font-size="22" font-family="Arial">85%</text>`\n'
        '            +`<line x1="0" y1="${m70}" x2="${VW}" y2="${m70}" stroke="rgba(255,255,255,0.18)" stroke-width="1.5" stroke-dasharray="6,5"/>`\n'
        '            +`<text x="10" y="${m70-5}" fill="rgba(255,255,255,0.3)" font-size="22" font-family="Arial">70%</text>`\n'
        '            +line+dots+"</svg>";\n'
        '  return `<div><div style="display:flex;">${lbls}</div>`\n'
        '        +`<div style="position:relative;"><div style="display:flex;">${bars}</div>${svg}</div>`\n'
        '        +`<div style="display:flex;">${names}</div>`\n'
        '        +`<div style="margin-top:6px;display:flex;align-items:center;gap:5px;">`\n'
        '        +`<div style="width:18px;height:3px;background:#f39c12;border-radius:2px;"></div>`\n'
        '        +`<span style="font-size:9px;color:#ddd;">Tendência/趋势</span></div></div>`;\n'
        '}\n'
        '\n'
        'function buildTrans(trans,campo,last2){\n'
        '  if(!last2||last2.length===0) return `<div style="color:#ddd;padding:20px;text-align:center;">Sem dados</div>`;\n'
        '  const mB=last2[last2.length-1],mA=last2.length>1?last2[last2.length-2]:null;\n'
        '  const tList=Object.keys(trans).sort().filter(t=>{\n'
        '    const vA=mA&&trans[t]&&trans[t][mA]?trans[t][mA]:{ids:0,okS:0,okC:0};\n'
        '    const vB=trans[t]&&trans[t][mB]?trans[t][mB]:{ids:0,okS:0,okC:0};\n'
        '    const valA=campo==="s"?pct(vA.okS,vA.ids):pct(vA.okC,vA.ids);\n'
        '    const valB=campo==="s"?pct(vB.okS,vB.ids):pct(vB.okC,vB.ids);\n'
        '    return valA>0&&valB>0;\n'
        '  });\n'
        '  const H=130;\n'
        '  let cols="";\n'
        '  tList.forEach(t=>{\n'
        '    const vA=mA&&trans[t]&&trans[t][mA]?trans[t][mA]:{ids:0,okS:0,okC:0};\n'
        '    const vB=trans[t]&&trans[t][mB]?trans[t][mB]:{ids:0,okS:0,okC:0};\n'
        '    const valA=campo==="s"?pct(vA.okS,vA.ids):pct(vA.okC,vA.ids);\n'
        '    const valB=campo==="s"?pct(vB.okS,vB.ids):pct(vB.okC,vB.ids);\n'
        '    const hA=vA.ids>0?Math.max(3,Math.round(valA/100*H)):0;\n'
        '    const hB=vB.ids>0?Math.max(3,Math.round(valB/100*H)):0;\n'
        '    const lA=vA.ids>0?valA.toFixed(0)+"%":"—",lB=vB.ids>0?valB.toFixed(0)+"%":"—";\n'
        '    cols+=`<div style="flex:1;display:flex;flex-direction:column;align-items:center;">`\n'
        '         +`<div style="width:100%;display:flex;gap:2px;align-items:flex-end;height:${H+18}px;">`\n'
        '         +`  <div style="flex:1;display:flex;flex-direction:column;align-items:center;">`\n'
        '         +`    <div style="font-size:9px;color:#e67e22;font-weight:700;">${lA}</div>`\n'
        '         +`    <div style="width:100%;height:${hA}px;background:#e67e22;border-radius:2px 2px 0 0;"></div>`\n'
        '         +`  </div>`\n'
        '         +`  <div style="flex:1;display:flex;flex-direction:column;align-items:center;">`\n'
        '         +`    <div style="font-size:9px;color:#27ae60;font-weight:700;">${lB}</div>`\n'
        '         +`    <div style="width:100%;height:${hB}px;background:#27ae60;border-radius:2px 2px 0 0;"></div>`\n'
        '         +`  </div>`\n'
        '         +"</div>"\n'
        '         +`<div style="font-size:8px;color:#ddd;text-align:center;margin-top:3px;word-break:break-word;line-height:1.2;">${t.length>9?t.slice(0,9)+"…":t}</div>`\n'
        '         +"</div>";\n'
        '  });\n'
        '  let leg="";\n'
        '  if(mA) leg+=`<div style="display:flex;align-items:center;gap:4px;"><div style="width:10px;height:10px;background:#e67e22;border-radius:2px;"></div><span style="font-size:9px;color:#ddd;">${MESES[mA-1][0]}/${MESES[mA-1][1]}</span></div>`;\n'
        '  leg+=`<div style="display:flex;align-items:center;gap:4px;"><div style="width:10px;height:10px;background:#27ae60;border-radius:2px;"></div><span style="font-size:9px;color:#ddd;">${MESES[mB-1][0]}/${MESES[mB-1][1]}</span></div>`;\n'
        '  return `<div><div style="display:flex;gap:4px;">${cols}</div><div style="margin-top:8px;display:flex;gap:12px;">${leg}</div></div>`;\n'
        '}\n'
        '\n'
        'function render(){\n'
        '  const {mes,trans}=agg();\n'
        '  const mks=Object.keys(mes).map(Number).sort((a,b)=>a-b);\n'
        '  const last2=mks.slice(-2);\n'
        '  const mA=last2.length>1?last2[0]:null,mB=last2[last2.length-1]||null;\n'
        '  const lbl=mA&&mB?MESES[mA-1][0]+"/"+MESES[mA-1][1]+" vs "+MESES[mB-1][0]+"/"+MESES[mB-1][1]:"";\n'
        '  document.getElementById("kpi-cards").innerHTML=buildKPI(mes);\n'
        '  document.getElementById("chart-saida-evol").innerHTML=buildEvol(mes,"s");\n'
        '  document.getElementById("chart-chegada-evol").innerHTML=buildEvol(mes,"c");\n'
        '  document.getElementById("chart-saida-trans").innerHTML=buildTrans(trans,"s",last2);\n'
        '  document.getElementById("chart-chegada-trans").innerHTML=buildTrans(trans,"c",last2);\n'
        '  document.getElementById("lbl-ts").textContent=lbl;\n'
        '  document.getElementById("lbl-tc").textContent=lbl;\n'
        '}\n'
        'function setFiltro(k,v){\n'
        '  if(k==="ano") fAno=v==="Todos"?"Todos":parseInt(v);\n'
        '  else if(k==="tipo") fTipo=v;\n'
        '  render();\n'
        '}\n'
        'document.addEventListener("DOMContentLoaded",render);\n'
        '</script>'
    )
    return pagina_html(nav, conteudo + js, "Pontualidade Mensal")


def pg_motoristas(df):
    base_cols = ['DATA', 'MÊS', 'Transportador', 'CONDUTOR', 'Tempo Saida OFF', 'Tempo chegada OFF']
    opt_cols = ['Semana Nome Nova']
    cols_needed = base_cols + [c for c in opt_cols if c in df.columns]

    df_js = df[[c for c in cols_needed]].copy()
    df_js['DATA'] = df_js['DATA'].dt.strftime('%Y-%m-%d')
    df_js = df_js.fillna('')

    rename_map = {
        'DATA': 'd', 'MÊS': 'm', 'Semana Nome Nova': 'sem',
        'Transportador': 't', 'CONDUTOR': 'c',
        'Tempo Saida OFF': 's', 'Tempo chegada OFF': 'ch'
    }
    df_js = df_js.rename(columns={k: v for k, v in rename_map.items() if k in df_js.columns})
    data_json = df_js.to_json(orient='records', force_ascii=False)

    transportadoras = sorted(df['Transportador'].dropna().unique().tolist())
    condutores = sorted(df['CONDUTOR'].dropna().unique().tolist())
    date_min = df['DATA'].min().strftime('%Y-%m-%d') if not df.empty else ''
    date_max = df['DATA'].max().strftime('%Y-%m-%d') if not df.empty else ''

    opts_trans = '<option value="Todos">Todos</option>' + ''.join(
        f'<option value="{t}">{t}</option>' for t in transportadoras)
    opts_cond = '<option value="Todos">Todos</option>' + ''.join(
        f'<option value="{c}">{c}</option>' for c in condutores)

    sel_style = 'width:100%;background:#0e0808;border:1px solid rgba(192,57,43,0.4);color:#fff;border-radius:4px;padding:5px 6px;font-size:11px;'
    inp_style = 'width:100%;background:#0e0808;border:1px solid rgba(192,57,43,0.4);color:#fff;border-radius:4px;padding:5px 6px;font-size:11px;color-scheme:dark;'
    sel_hdr = 'background:#1a1010;border:1px solid rgba(192,57,43,0.5);color:#fff;border-radius:4px;padding:5px 10px;font-size:11px;min-width:220px;'

    filtro_html = f"""
    <div style="width:210px;flex-shrink:0;">
      <div style="background:#1a1010;border-radius:8px;padding:14px;border:1px solid rgba(192,57,43,0.3);position:sticky;top:16px;">
        <div style="color:#C0392B;font-size:11px;font-weight:700;letter-spacing:1px;margin-bottom:12px;padding-bottom:8px;border-bottom:1px solid rgba(192,57,43,0.3);">FILTROS / 筛选</div>
        <div style="margin-bottom:10px;">
          <div style="color:#ddd;font-size:9px;letter-spacing:1px;margin-bottom:4px;">日期从 | Data De</div>
          <input type="date" id="f-date-from" value="{date_max}" onchange="setFiltro('dateFrom',this.value)" style="{inp_style}margin-bottom:4px;">
          <div style="color:#ddd;font-size:9px;letter-spacing:1px;margin-bottom:4px;margin-top:6px;">日期到 | Data Até</div>
          <input type="date" id="f-date-to" value="{date_max}" onchange="setFiltro('dateTo',this.value)" style="{inp_style}">
        </div>
        <div style="margin-bottom:14px;">
          <div style="color:#ddd;font-size:9px;letter-spacing:1px;margin-bottom:4px;">承运人 | Transportadora</div>
          <select id="f-trans" onchange="setFiltro('trans',this.value)" style="{sel_style}">{opts_trans}</select>
        </div>
        <button onclick="resetFiltros()" style="width:100%;background:rgba(192,57,43,0.2);border:1px solid rgba(192,57,43,0.4);color:#C0392B;border-radius:4px;padding:7px;font-size:10px;font-weight:700;cursor:pointer;letter-spacing:1px;">RESET / 重置</button>
      </div>
    </div>"""

    nav = navbar("motoristas")
    conteudo = f"""
    <div class="page-header">
      <div class="header-left">
        <div class="header-icon" style="background:#C0392B;">👤</div>
        <div>
          <div class="header-title">GESTÃO DE MOTORISTAS / 司机管理</div>
          <div class="header-sub">KPI SAÍDA · KPI CHEGADA · KPI GERAL · <span id="periodo-txt">—</span> · USE OS FILTROS PARA NAVEGAR</div>
        </div>
      </div>
      <div style="display:flex;align-items:center;gap:8px;flex:1;justify-content:center;padding:0 16px;">
        <div style="color:#ddd;font-size:10px;font-weight:600;white-space:nowrap;">驾驶员 | Motoristas</div>
        <select id="f-cond" onchange="setFiltro('cond',this.value)" style="{sel_hdr}">{opts_cond}</select>
      </div>
      <div class="header-badges">
        <div class="header-badge"><div class="hb-label">MÊS/月</div><div class="hb-value" id="hb-mes">—</div></div>
        <div class="header-badge"><div class="hb-label">SEMANA/周</div><div class="hb-value" id="hb-sem">—</div></div>
        <div class="header-badge"><div class="hb-label">PERÍODO/期间</div><div class="hb-value" id="hb-periodo">—</div></div>
      </div>
    </div>

    <div style="display:flex;gap:16px;align-items:flex-start;">
      <div style="flex:1;min-width:0;">

        <div class="kpi-grid kpi-grid-5" style="margin-bottom:12px;">
          <div class="kpi-card" style="border-top:3px solid #C0392B;">
            <div class="kpi-label">TOTAL MOTORISTAS / 司机</div>
            <div class="kpi-value" id="kpi-total">—</div>
            <div class="kpi-sub">condutores no período</div>
          </div>
          <div class="kpi-card" style="border-top:3px solid #27ae60;">
            <div class="kpi-label">KPI SAÍDA / 出发指标</div>
            <div class="kpi-value" id="kpi-s" style="color:#27ae60;">—</div>
            <div class="kpi-sub" id="kpi-s-sub"></div>
            <div class="progress-bar" style="margin-top:6px;"><div id="kpi-bar-s" class="progress-fill" style="width:0%;background:#27ae60;"></div></div>
          </div>
          <div class="kpi-card" style="border-top:3px solid #27ae60;">
            <div class="kpi-label">KPI CHEGADA / 到达指标</div>
            <div class="kpi-value" id="kpi-c" style="color:#27ae60;">—</div>
            <div class="kpi-sub" id="kpi-c-sub"></div>
            <div class="progress-bar" style="margin-top:6px;"><div id="kpi-bar-c" class="progress-fill" style="width:0%;background:#27ae60;"></div></div>
          </div>
          <div class="kpi-card" style="border-top:3px solid #27ae60;">
            <div class="kpi-label">KPI GERAL / 综合指标</div>
            <div class="kpi-value" id="kpi-g" style="color:#27ae60;">—</div>
            <div class="kpi-sub">média KPI saída + chegada</div>
            <div class="progress-bar" style="margin-top:6px;"><div id="kpi-bar-g" class="progress-fill" style="width:0%;background:#27ae60;"></div></div>
          </div>
          <div class="kpi-card" style="border-top:3px solid #e74c3c;">
            <div class="kpi-label">ABAIXO META / 未达标</div>
            <div class="kpi-value" id="kpi-abaixo" style="color:#e74c3c;">—</div>
            <div class="kpi-sub">motoristas KPI abaixo 70%</div>
          </div>
        </div>

        <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px;">

          <div class="table-wrap">
            <div class="table-header" style="background:#1A5276;">
              <span>KPI SAÍDA / 出发指标</span>
              <small>% ON por motorista · melhor primeiro</small>
            </div>
            <div class="table-scroll" style="max-height:calc(100vh - 360px);">
              <table>
                <thead><tr style="background:#0d2a3a;border-bottom:2px solid rgba(26,82,118,0.5);">
                  <th style="text-align:center;color:#5dade2;width:24px;">#</th>
                  <th style="text-align:left;color:#5dade2;">CONDUTOR/司机</th>
                  <th style="text-align:left;color:#5dade2;">TRANSP.</th>
                  <th style="text-align:center;color:#5dade2;">IDs</th>
                  <th style="text-align:center;color:#5dade2;">OFF</th>
                  <th style="text-align:center;color:#5dade2;">KPI S.</th>
                </tr></thead>
                <tbody id="tbody-s"></tbody>
                <tfoot><tr id="tfoot-s" style="background:#1A5276;font-size:11px;font-weight:700;"></tr></tfoot>
              </table>
            </div>
          </div>

          <div class="table-wrap">
            <div class="table-header" style="background:#1A5276;">
              <span>KPI CHEGADA / 到达指标</span>
              <small>% ON por motorista · melhor primeiro</small>
            </div>
            <div class="table-scroll" style="max-height:calc(100vh - 360px);">
              <table>
                <thead><tr style="background:#0d2a3a;border-bottom:2px solid rgba(26,82,118,0.5);">
                  <th style="text-align:center;color:#5dade2;width:24px;">#</th>
                  <th style="text-align:left;color:#5dade2;">CONDUTOR/司机</th>
                  <th style="text-align:left;color:#5dade2;">TRANSP.</th>
                  <th style="text-align:center;color:#5dade2;">IDs</th>
                  <th style="text-align:center;color:#5dade2;">OFF</th>
                  <th style="text-align:center;color:#5dade2;">KPI C.</th>
                </tr></thead>
                <tbody id="tbody-c"></tbody>
                <tfoot><tr id="tfoot-c" style="background:#1A5276;font-size:11px;font-weight:700;"></tr></tfoot>
              </table>
            </div>
          </div>

          <div class="table-wrap">
            <div class="table-header" style="background:#1e8449;">
              <span>KPI GERAL / 综合指标</span>
              <small>média saída + chegada · melhor primeiro</small>
            </div>
            <div class="table-scroll" style="max-height:calc(100vh - 360px);">
              <table>
                <thead><tr style="background:#0d3320;border-bottom:2px solid rgba(30,132,73,0.5);">
                  <th style="text-align:center;color:#2ecc71;width:24px;">#</th>
                  <th style="text-align:left;color:#2ecc71;">CONDUTOR/司机</th>
                  <th style="text-align:left;color:#2ecc71;">TRANSP.</th>
                  <th style="text-align:center;color:#2ecc71;">IDs</th>
                  <th style="text-align:center;color:#2ecc71;">KPI S.</th>
                  <th style="text-align:center;color:#2ecc71;">KPI C.</th>
                </tr></thead>
                <tbody id="tbody-g"></tbody>
                <tfoot><tr id="tfoot-g" style="background:#1e8449;font-size:11px;font-weight:700;"></tr></tfoot>
              </table>
            </div>
          </div>

        </div>

        <div class="legend-bar" style="margin-top:8px;">
          <div class="legend-item" style="color:#27ae60;"><div class="legend-dot" style="background:#27ae60;"></div>≥85% — META atingida / 达标</div>
          <div class="legend-item" style="color:#e67e22;"><div class="legend-dot" style="background:#e67e22;"></div>70–84% — PARCIAL / 部分</div>
          <div class="legend-item" style="color:#e74c3c;"><div class="legend-dot" style="background:#e74c3c;"></div>&lt;70% — CRÍTICO / 危急</div>
          <div style="margin-left:auto;font-size:9px;color:#888;">KPI = % viagens com app ON · meta ≥85% · Slicers: Data → Transportadora → Motorista</div>
        </div>

      </div>
      FILTRO_PLACEHOLDER
    </div>"""

    conteudo = conteudo.replace('FILTRO_PLACEHOLDER', filtro_html)

    js = (
        '<script>\n'
        'const RAW = ' + data_json + ';\n'
        'const _DM="' + date_max + '",_DMi="' + date_min + '";\n'
        'const _lsM=localStorage.getItem("brdrive_data_max"),_lsF2=localStorage.getItem("brdrive_from"),_lsT2=localStorage.getItem("brdrive_to");\n'
        'if(_lsM!==_DM||(_lsF2&&_lsF2>_DM)||(_lsT2&&_lsT2<_DMi)){localStorage.removeItem("brdrive_from");localStorage.removeItem("brdrive_to");localStorage.setItem("brdrive_data_max",_DM);}\n'
        'const _lsF=localStorage.getItem("brdrive_from"),_lsT=localStorage.getItem("brdrive_to");\n'
        'let fDateFrom=_lsF||"' + date_max + '";\n'
        'let fDateTo=_lsT||"' + date_max + '";\n'
        'let fTrans = "Todos"; let fCond = "Todos";\n'
        '\n'
        'function cor(v){if(v>=85)return"#27ae60";if(v>=70)return"#e67e22";return"#e74c3c";}\n'
        'function bgCor(v){if(v>=85)return"rgba(39,174,96,0.15)";if(v>=70)return"rgba(230,126,34,0.15)";return"rgba(231,76,60,0.15)";}\n'
        'function pct(a,b){return b>0?(a/b*100):0;}\n'
        '\n'
        'function getFiltrado(){\n'
        '  return RAW.filter(r=>{\n'
        '    if(fDateFrom && r.d < fDateFrom) return false;\n'
        '    if(fDateTo && r.d > fDateTo) return false;\n'
        '    if(fTrans!=="Todos" && r.t!==fTrans) return false;\n'
        '    if(fCond!=="Todos" && r.c!==fCond) return false;\n'
        '    return true;\n'
        '  });\n'
        '}\n'
        '\n'
        'function aggByDriver(data){\n'
        '  const byC = {};\n'
        '  data.forEach(r=>{\n'
        '    if(!byC[r.c]) byC[r.c]={ids:0,offS:0,offC:0,okS:0,okC:0,trans:new Set()};\n'
        '    byC[r.c].ids++;\n'
        '    if(r.s==="OFF") byC[r.c].offS++; else byC[r.c].okS++;\n'
        '    if(r.ch==="OFF") byC[r.c].offC++; else byC[r.c].okC++;\n'
        '    if(r.t) byC[r.c].trans.add(r.t);\n'
        '  });\n'
        '  return byC;\n'
        '}\n'
        '\n'
        'function transBadges(transSet){\n'
        '  return [...transSet].map(t=>`<span class="trans-badge">${t}</span>`).join("");\n'
        '}\n'
        '\n'
        'function render(){\n'
        '  const data = getFiltrado();\n'
        '  const byC = aggByDriver(data);\n'
        '  const drivers = Object.entries(byC);\n'
        '\n'
        '  const totalIDs = data.length;\n'
        '  const totalOffS = data.filter(r=>r.s==="OFF").length;\n'
        '  const totalOffC = data.filter(r=>r.ch==="OFF").length;\n'
        '  const totalOkS = totalIDs - totalOffS;\n'
        '  const totalOkC = totalIDs - totalOffC;\n'
        '  const gKpiS = pct(totalOkS, totalIDs);\n'
        '  const gKpiC = pct(totalOkC, totalIDs);\n'
        '  const gKpiG = totalIDs>0 ? (gKpiS + gKpiC) / 2 : 0;\n'
        '  const abaixo = drivers.filter(([,v])=>(pct(v.okS,v.ids)+pct(v.okC,v.ids))/2 < 70).length;\n'
        '\n'
        '  document.getElementById("kpi-total").textContent = drivers.length;\n'
        '  const eS=document.getElementById("kpi-s"); eS.textContent=gKpiS.toFixed(1)+"%"; eS.style.color=cor(gKpiS);\n'
        '  document.getElementById("kpi-s-sub").textContent=totalOffS+"x OFF · "+totalOkS+" ON";\n'
        '  const bS=document.getElementById("kpi-bar-s"); bS.style.width=Math.min(gKpiS,100).toFixed(0)+"%"; bS.style.background=cor(gKpiS);\n'
        '  const eC=document.getElementById("kpi-c"); eC.textContent=gKpiC.toFixed(1)+"%"; eC.style.color=cor(gKpiC);\n'
        '  document.getElementById("kpi-c-sub").textContent=totalOffC+"x OFF · "+totalOkC+" ON";\n'
        '  const bC=document.getElementById("kpi-bar-c"); bC.style.width=Math.min(gKpiC,100).toFixed(0)+"%"; bC.style.background=cor(gKpiC);\n'
        '  const eG=document.getElementById("kpi-g"); eG.textContent=gKpiG.toFixed(1)+"%"; eG.style.color=cor(gKpiG);\n'
        '  const bG=document.getElementById("kpi-bar-g"); bG.style.width=Math.min(gKpiG,100).toFixed(0)+"%"; bG.style.background=cor(gKpiG);\n'
        '  document.getElementById("kpi-abaixo").textContent = abaixo;\n'
        '\n'
        '  const dates = data.map(r=>r.d).filter(Boolean).sort();\n'
        '  const periodo = dates.length>0 ? dates[0]+" → "+dates[dates.length-1] : "—";\n'
        '  document.getElementById("periodo-txt").textContent = periodo;\n'
        '  document.getElementById("hb-periodo").textContent = periodo;\n'
        '  const meses = [...new Set(data.map(r=>r.m).filter(Boolean))];\n'
        '  document.getElementById("hb-mes").textContent = meses.length>0 ? meses.join("/") : "—";\n'
        '  const sems = [...new Set(data.map(r=>(r.sem||"")).filter(Boolean))];\n'
        '  document.getElementById("hb-sem").textContent = sems.length>0 ? sems[sems.length-1] : "—";\n'
        '\n'
        '  // KPI SAÍDA table — melhor primeiro\n'
        '  const sortedS = drivers.slice().sort((a,b)=>pct(b[1].okS,b[1].ids)-pct(a[1].okS,a[1].ids));\n'
        '  let rS="", tfSids=0, tfSoff=0;\n'
        '  sortedS.forEach(([c,v],i)=>{\n'
        '    const ks=pct(v.okS,v.ids); const bgr=i%2===0?"#1f0e0e":"#1a1010";\n'
        '    const offLbl=v.offS>0?`<span style="color:#e74c3c;font-weight:700;">${v.offS}x</span>&nbsp;<span style="color:#e74c3c;font-size:9px;">OFF</span>`:`<span style="color:#27ae60;font-size:9px;font-weight:700;">OK</span>`;\n'
        '    tfSids+=v.ids; tfSoff+=v.offS;\n'
        '    rS+=`<tr style="background:${bgr};">'
        '<td style="text-align:center;color:#888;font-size:10px;">${i+1}</td>'
        '<td style="color:#5dade2;font-weight:600;font-size:10px;">${c}</td>'
        '<td style="font-size:10px;">${transBadges(v.trans)}</td>'
        '<td style="text-align:center;color:#f1c40f;font-weight:700;">${v.ids}</td>'
        '<td style="text-align:center;">${offLbl}</td>'
        '<td style="text-align:center;"><span style="background:${bgCor(ks)};color:${cor(ks)};padding:2px 8px;border-radius:20px;font-size:10px;font-weight:700;">${ks.toFixed(1)}%</span></td>'
        '</tr>`;\n'
        '  });\n'
        '  document.getElementById("tbody-s").innerHTML = rS || "<tr><td colspan=\'6\' style=\'text-align:center;color:#ddd;padding:16px;\'>Sem dados / 暂无数据</td></tr>";\n'
        '  document.getElementById("tfoot-s").innerHTML=`<td colspan="3" style="font-weight:700;">TOTAL</td>'
        '<td style="text-align:center;font-weight:700;">${tfSids}</td>'
        '<td style="text-align:center;font-weight:700;color:#e74c3c;">${tfSoff}x OFF</td>'
        '<td style="text-align:center;font-weight:700;">${pct(tfSids-tfSoff,tfSids).toFixed(1)}%</td>`;\n'
        '\n'
        '  // KPI CHEGADA table — melhor primeiro\n'
        '  const sortedC = drivers.slice().sort((a,b)=>pct(b[1].okC,b[1].ids)-pct(a[1].okC,a[1].ids));\n'
        '  let rC="", tfCids=0, tfCoff=0;\n'
        '  sortedC.forEach(([c,v],i)=>{\n'
        '    const kc=pct(v.okC,v.ids); const bgr=i%2===0?"#1f0e0e":"#1a1010";\n'
        '    const offLbl=v.offC>0?`<span style="color:#e74c3c;font-weight:700;">${v.offC}x</span>&nbsp;<span style="color:#e74c3c;font-size:9px;">OFF</span>`:`<span style="color:#27ae60;font-size:9px;font-weight:700;">OK</span>`;\n'
        '    tfCids+=v.ids; tfCoff+=v.offC;\n'
        '    rC+=`<tr style="background:${bgr};">'
        '<td style="text-align:center;color:#888;font-size:10px;">${i+1}</td>'
        '<td style="color:#5dade2;font-weight:600;font-size:10px;">${c}</td>'
        '<td style="font-size:10px;">${transBadges(v.trans)}</td>'
        '<td style="text-align:center;color:#f1c40f;font-weight:700;">${v.ids}</td>'
        '<td style="text-align:center;">${offLbl}</td>'
        '<td style="text-align:center;"><span style="background:${bgCor(kc)};color:${cor(kc)};padding:2px 8px;border-radius:20px;font-size:10px;font-weight:700;">${kc.toFixed(1)}%</span></td>'
        '</tr>`;\n'
        '  });\n'
        '  document.getElementById("tbody-c").innerHTML = rC || "<tr><td colspan=\'6\' style=\'text-align:center;color:#ddd;padding:16px;\'>Sem dados / 暂无数据</td></tr>";\n'
        '  document.getElementById("tfoot-c").innerHTML=`<td colspan="3" style="font-weight:700;">TOTAL</td>'
        '<td style="text-align:center;font-weight:700;">${tfCids}</td>'
        '<td style="text-align:center;font-weight:700;color:#e74c3c;">${tfCoff}x OFF</td>'
        '<td style="text-align:center;font-weight:700;">${pct(tfCids-tfCoff,tfCids).toFixed(1)}%</td>`;\n'
        '\n'
        '  // KPI GERAL table — melhor primeiro\n'
        '  const sortedG = drivers.slice().sort((a,b)=>{\n'
        '    const ga=(pct(a[1].okS,a[1].ids)+pct(a[1].okC,a[1].ids))/2;\n'
        '    const gb=(pct(b[1].okS,b[1].ids)+pct(b[1].okC,b[1].ids))/2;\n'
        '    return gb-ga;\n'
        '  });\n'
        '  let rG="", tfGids=0;\n'
        '  sortedG.forEach(([c,v],i)=>{\n'
        '    const ks=pct(v.okS,v.ids); const kc=pct(v.okC,v.ids);\n'
        '    const bgr=i%2===0?"#0d1f12":"#0a1a0e";\n'
        '    tfGids+=v.ids;\n'
        '    rG+=`<tr style="background:${bgr};">'
        '<td style="text-align:center;color:#888;font-size:10px;">${i+1}</td>'
        '<td style="color:#5dade2;font-weight:600;font-size:10px;">${c}</td>'
        '<td style="font-size:10px;">${transBadges(v.trans)}</td>'
        '<td style="text-align:center;color:#f1c40f;font-weight:700;">${v.ids}</td>'
        '<td style="text-align:center;"><span style="background:${bgCor(ks)};color:${cor(ks)};padding:2px 6px;border-radius:20px;font-size:10px;font-weight:700;">${ks.toFixed(1)}%</span></td>'
        '<td style="text-align:center;"><span style="background:${bgCor(kc)};color:${cor(kc)};padding:2px 6px;border-radius:20px;font-size:10px;font-weight:700;">${kc.toFixed(1)}%</span></td>'
        '</tr>`;\n'
        '  });\n'
        '  document.getElementById("tbody-g").innerHTML = rG || "<tr><td colspan=\'6\' style=\'text-align:center;color:#ddd;padding:16px;\'>Sem dados / 暂无数据</td></tr>";\n'
        '  document.getElementById("tfoot-g").innerHTML=`<td colspan="3" style="font-weight:700;">TOTAL</td>'
        '<td style="text-align:center;font-weight:700;">${tfGids}</td>'
        '<td style="text-align:center;font-weight:700;">${gKpiS.toFixed(1)}% S.</td>'
        '<td style="text-align:center;font-weight:700;">${gKpiC.toFixed(1)}% C.</td>`;\n'
        '}\n'
        '\n'
        'function setFiltro(key,val){\n'
        '  if(key==="dateFrom"){fDateFrom=val;localStorage.setItem("brdrive_from",val);}\n'
        '  else if(key==="dateTo"){fDateTo=val;localStorage.setItem("brdrive_to",val);}\n'
        '  else if(key==="trans") fTrans=val;\n'
        '  else if(key==="cond") fCond=val;\n'
        '  render();\n'
        '}\n'
        '\n'
        'function resetFiltros(){\n'
        '  fDateFrom="' + date_max + '"; fDateTo="' + date_max + '";\n'
        '  fTrans="Todos"; fCond="Todos";\n'
        '  document.getElementById("f-date-from").value=fDateFrom;\n'
        '  document.getElementById("f-date-to").value=fDateTo;\n'
        '  document.getElementById("f-trans").value="Todos";\n'
        '  document.getElementById("f-cond").value="Todos";\n'
        '  render();\n'
        '}\n'
        '\n'
        'document.addEventListener("DOMContentLoaded",function(){\n'
        '  var fi=document.getElementById("f-date-from"),ti=document.getElementById("f-date-to");\n'
        '  if(fi)fi.value=fDateFrom;\n'
        '  if(ti)ti.value=fDateTo;\n'
        '  render();\n'
        '});\n'
        '</script>'
    )

    return pagina_html(nav, conteudo + js, "Gestão de Motoristas")


def pg_mural():
    return """<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Mural de Dashboards — BR DRIVE</title>
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body { font-family: Arial, sans-serif; background: #0e0808; color: #fff; height: 100vh; display: flex; flex-direction: column; overflow: hidden; }

    /* TOP BAR */
    .mural-bar { background: #1a0a0a; border-bottom: 1px solid rgba(192,57,43,0.4); height: 52px; display: flex; align-items: center; gap: 8px; padding: 0 12px; flex-shrink: 0; overflow-x: auto; }
    .mural-logo { background: #C0392B; color: #fff; font-weight: 900; font-size: 12px; padding: 5px 8px; border-radius: 6px; white-space: nowrap; flex-shrink: 0; }
    .mural-title { color: #fff; font-size: 13px; font-weight: 700; letter-spacing: 1px; white-space: nowrap; flex-shrink: 0; }
    .mural-sep { width: 1px; height: 30px; background: rgba(192,57,43,0.35); flex-shrink: 0; }
    .bar-group { display: flex; align-items: center; gap: 5px; flex-shrink: 0; }
    .bar-label { color: rgba(255,255,255,0.45); font-size: 8px; letter-spacing: 1px; white-space: nowrap; }
    .btn-layout { background: rgba(255,255,255,0.07); border: 1px solid rgba(255,255,255,0.15); color: rgba(255,255,255,0.7); border-radius: 5px; padding: 4px 9px; font-size: 10px; cursor: pointer; transition: all 0.15s; white-space: nowrap; }
    .btn-layout:hover { background: rgba(192,57,43,0.25); border-color: rgba(192,57,43,0.5); color: #fff; }
    .btn-layout.active { background: rgba(192,57,43,0.35); border-color: #C0392B; color: #fff; }
    .btn-shortcut { background: rgba(26,82,118,0.3); border: 1px solid rgba(41,128,185,0.35); color: #5dade2; border-radius: 5px; padding: 4px 8px; font-size: 10px; cursor: pointer; transition: all 0.15s; white-space: nowrap; }
    .btn-shortcut:hover { background: rgba(41,128,185,0.3); border-color: #2980b9; color: #fff; }
    .btn-config { background: rgba(142,68,173,0.25); border: 1px solid rgba(142,68,173,0.4); color: #af7ac5; border-radius: 5px; padding: 4px 9px; font-size: 10px; cursor: pointer; transition: all 0.15s; white-space: nowrap; }
    .btn-config:hover { background: rgba(142,68,173,0.4); color: #fff; }
    .mural-clock { color: #fff; font-size: 14px; font-weight: 700; letter-spacing: 2px; font-variant-numeric: tabular-nums; white-space: nowrap; }
    .btn-fullscreen { background: rgba(39,174,96,0.25); border: 1px solid rgba(39,174,96,0.4); color: #58d68d; border-radius: 5px; padding: 4px 9px; font-size: 10px; cursor: pointer; transition: all 0.15s; white-space: nowrap; }
    .btn-fullscreen:hover { background: rgba(39,174,96,0.4); color: #fff; }
    .btn-home { background: rgba(255,255,255,0.07); border: 1px solid rgba(255,255,255,0.15); color: rgba(255,255,255,0.6); border-radius: 5px; padding: 4px 8px; font-size: 10px; text-decoration: none; display: flex; align-items: center; transition: all 0.15s; white-space: nowrap; }
    .btn-home:hover { background: rgba(192,57,43,0.2); color: #fff; }
    .bar-spacer { flex: 1; min-width: 8px; }

    /* GRID */
    .mural-grid { flex: 1; display: grid; gap: 4px; padding: 4px; overflow: hidden; }
    .layout-2x2 { grid-template-columns: 1fr 1fr; grid-template-rows: 1fr 1fr; }
    .layout-1x4 { grid-template-columns: 1fr; grid-template-rows: 1fr 1fr 1fr 1fr; }
    .layout-2x1 { grid-template-columns: 1fr 1fr; grid-template-rows: 1fr; }
    .layout-1x2 { grid-template-columns: 1fr; grid-template-rows: 1fr 1fr; }
    .layout-1x1 { grid-template-columns: 1fr; grid-template-rows: 1fr; }

    /* PANEL — --pz é o fator de zoom do iframe */
    .mural-panel {
      --pz: 0.62;
      background: #1a0a0a;
      border: 1px solid rgba(192,57,43,0.25);
      border-radius: 6px;
      overflow: hidden;
      display: flex;
      flex-direction: column;
    }
    .panel-bar { height: 26px; background: rgba(0,0,0,0.35); border-bottom: 1px solid rgba(192,57,43,0.2); display: flex; align-items: center; padding: 0 8px; gap: 6px; flex-shrink: 0; z-index: 2; }
    .panel-num { width: 16px; height: 16px; background: #C0392B; border-radius: 3px; display: flex; align-items: center; justify-content: center; font-size: 9px; font-weight: 700; color: #fff; flex-shrink: 0; }
    .panel-title { color: rgba(255,255,255,0.8); font-size: 9px; font-weight: 600; flex: 1; }
    .panel-btn { background: none; border: none; color: rgba(255,255,255,0.35); cursor: pointer; font-size: 12px; padding: 2px 3px; border-radius: 3px; transition: all 0.15s; line-height: 1; }
    .panel-btn:hover { background: rgba(192,57,43,0.25); color: #fff; }

    /* Wrapper + iframe com scale para caber no painel */
    .panel-frame-wrap { flex: 1; position: relative; overflow: hidden; }
    .panel-iframe {
      position: absolute;
      top: 0; left: 0;
      width:  calc(100% / var(--pz));
      height: calc(100% / var(--pz));
      border: none;
      background: #0e0808;
      transform: scale(var(--pz));
      transform-origin: top left;
    }

    /* MAXIMIZE OVERLAY */
    .panel-max-overlay { display: none; position: fixed; inset: 0; z-index: 500; background: #0e0808; flex-direction: column; }
    .panel-max-overlay.open { display: flex; }
    .panel-max-bar { height: 38px; background: #1a0a0a; border-bottom: 1px solid rgba(192,57,43,0.3); display: flex; align-items: center; padding: 0 14px; gap: 10px; flex-shrink: 0; }
    .panel-max-title { color: #fff; font-size: 12px; font-weight: 700; flex: 1; }
    .panel-max-close { background: rgba(192,57,43,0.3); border: 1px solid rgba(192,57,43,0.5); color: #fff; border-radius: 5px; padding: 4px 12px; font-size: 11px; cursor: pointer; }
    .panel-max-close:hover { background: #C0392B; }
    .panel-max-iframe { flex: 1; border: none; width: 100%; }

    /* MODAL */
    .modal-overlay { display: none; position: fixed; inset: 0; background: rgba(0,0,0,0.75); z-index: 1000; align-items: center; justify-content: center; }
    .modal-overlay.open { display: flex; }
    .modal-box { background: #1a0a0a; border: 1px solid rgba(192,57,43,0.45); border-radius: 10px; padding: 20px; width: 520px; max-width: 95vw; }
    .modal-title { color: #fff; font-size: 13px; font-weight: 700; letter-spacing: 1px; margin-bottom: 14px; border-bottom: 1px solid rgba(192,57,43,0.3); padding-bottom: 10px; }
    .modal-row { margin-bottom: 10px; }
    .modal-row label { color: #ddd; font-size: 9px; letter-spacing: 1px; display: block; margin-bottom: 4px; }
    .modal-row select { width: 100%; background: #0e0808; border: 1px solid rgba(192,57,43,0.3); color: #fff; border-radius: 4px; padding: 6px 8px; font-size: 11px; }
    .modal-row select:focus { outline: none; border-color: #C0392B; }
    .modal-btns { display: flex; gap: 8px; justify-content: flex-end; margin-top: 16px; }
    .modal-ok { background: #C0392B; color: #fff; border: none; border-radius: 5px; padding: 7px 18px; font-size: 11px; font-weight: 700; cursor: pointer; }
    .modal-ok:hover { background: #a93226; }
    .modal-cancel { background: rgba(255,255,255,0.08); color: #ddd; border: 1px solid rgba(255,255,255,0.15); border-radius: 5px; padding: 7px 18px; font-size: 11px; cursor: pointer; }
    .modal-cancel:hover { background: rgba(255,255,255,0.15); }
  </style>
</head>
<body>

<!-- TOP BAR -->
<div class="mural-bar">
  <div class="mural-logo">J&amp;T</div>
  <span class="mural-title">BR DRIVE &mdash; MURAL DE DASHBOARDS</span>
  <div class="mural-sep"></div>

  <div class="bar-group">
    <span class="bar-label">LAYOUT</span>
    <button class="btn-layout active" onclick="setLayout('2x2')" id="btn-2x2">&#9638; 2&times;2</button>
    <button class="btn-layout" onclick="setLayout('1x4')" id="btn-1x4">&#8801; 1&times;4</button>
    <button class="btn-layout" onclick="setLayout('2x1')" id="btn-2x1">&#9636; 2&times;1</button>
    <button class="btn-layout" onclick="setLayout('1x1')" id="btn-1x1">&#9635; 1&times;1</button>
  </div>

  <div class="mural-sep"></div>

  <div class="bar-group">
    <span class="bar-label">ATALHOS</span>
    <button class="btn-shortcut" onclick="setPanel(0,'dashboard.html','Utiliza&#231;&#227;o APP')">&#128202; Utiliza&#231;&#227;o</button>
    <button class="btn-shortcut" onclick="setPanel(1,'pont_op.html','Pontualidade Op.')">&#8987; Pontualidade</button>
    <button class="btn-shortcut" onclick="setPanel(2,'app_mensal.html','APP Mensal')">&#128197; APP Mensal</button>
    <button class="btn-shortcut" onclick="setPanel(3,'pont_mensal.html','Pont. Mensal')">&#128198; Pont. Mensal</button>
  </div>

  <div class="mural-sep"></div>

  <button class="btn-config" onclick="openConfig()">&#9881; Configurar Pain&#233;is</button>

  <div class="bar-spacer"></div>

  <span class="mural-clock" id="mural-clock">00:00:00</span>

  <div class="mural-sep"></div>

  <button class="btn-fullscreen" onclick="toggleFullscreen()" id="btn-full">&#9974; Tela Cheia</button>
  <a href="home.html" class="btn-home">&#127968; Home</a>
</div>

<!-- MURAL GRID -->
<div class="mural-grid layout-2x2" id="mural-grid">
  <div class="mural-panel" id="panel-0">
    <div class="panel-bar">
      <div class="panel-num">1</div>
      <span class="panel-title" id="ptitle-0">Utiliza&#231;&#227;o APP</span>
      <button class="panel-btn" onclick="maximizePanel(0)" title="Maximizar">&#10548;</button>
      <button class="panel-btn" onclick="reloadPanel(0)" title="Recarregar">&#8635;</button>
    </div>
    <div class="panel-frame-wrap">
      <iframe class="panel-iframe" id="iframe-0" src="dashboard.html"></iframe>
    </div>
  </div>
  <div class="mural-panel" id="panel-1">
    <div class="panel-bar">
      <div class="panel-num">2</div>
      <span class="panel-title" id="ptitle-1">Pontualidade Op.</span>
      <button class="panel-btn" onclick="maximizePanel(1)" title="Maximizar">&#10548;</button>
      <button class="panel-btn" onclick="reloadPanel(1)" title="Recarregar">&#8635;</button>
    </div>
    <div class="panel-frame-wrap">
      <iframe class="panel-iframe" id="iframe-1" src="pont_op.html"></iframe>
    </div>
  </div>
  <div class="mural-panel" id="panel-2">
    <div class="panel-bar">
      <div class="panel-num">3</div>
      <span class="panel-title" id="ptitle-2">APP Mensal</span>
      <button class="panel-btn" onclick="maximizePanel(2)" title="Maximizar">&#10548;</button>
      <button class="panel-btn" onclick="reloadPanel(2)" title="Recarregar">&#8635;</button>
    </div>
    <div class="panel-frame-wrap">
      <iframe class="panel-iframe" id="iframe-2" src="app_mensal.html"></iframe>
    </div>
  </div>
  <div class="mural-panel" id="panel-3">
    <div class="panel-bar">
      <div class="panel-num">4</div>
      <span class="panel-title" id="ptitle-3">Pont. Mensal</span>
      <button class="panel-btn" onclick="maximizePanel(3)" title="Maximizar">&#10548;</button>
      <button class="panel-btn" onclick="reloadPanel(3)" title="Recarregar">&#8635;</button>
    </div>
    <div class="panel-frame-wrap">
      <iframe class="panel-iframe" id="iframe-3" src="pont_mensal.html"></iframe>
    </div>
  </div>
</div>

<!-- MAXIMIZE OVERLAY -->
<div class="panel-max-overlay" id="max-overlay">
  <div class="panel-max-bar">
    <span class="panel-max-title" id="max-title">&mdash;</span>
    <button class="panel-max-close" onclick="closeMaximize()">&#10005; Fechar</button>
  </div>
  <iframe class="panel-max-iframe" id="max-iframe" src="about:blank"></iframe>
</div>

<!-- CONFIG MODAL -->
<div class="modal-overlay" id="config-modal" onclick="if(event.target===this)closeConfig()">
  <div class="modal-box">
    <div class="modal-title">&#9881; Configurar Pain&#233;is / 配置面板</div>
    <div class="modal-row">
      <label>PAINEL 1</label>
      <select id="cfg-0">
        <option value="dashboard.html" selected>&#128202; Utiliza&#231;&#227;o APP</option>
        <option value="pont_op.html">&#8987; Pontualidade Op.</option>
        <option value="evol_semanal.html">&#128200; App Semanal</option>
        <option value="app_mensal.html">&#128197; APP Mensal</option>
        <option value="infracoes.html">&#128680; Infra&#231;&#245;es</option>
        <option value="motoristas.html">&#128100; Motoristas APP</option>
        <option value="pont_semanal.html">&#128202; Pont. Semanal</option>
        <option value="pont_mensal.html">&#128198; Pont. Mensal</option>
        <option value="inf_pontualidade.html">&#128308; Inf. Pontualidade</option>
        <option value="pont_motoristas.html">&#128663; Pont. Motoristas</option>
      </select>
    </div>
    <div class="modal-row">
      <label>PAINEL 2</label>
      <select id="cfg-1">
        <option value="dashboard.html">&#128202; Utiliza&#231;&#227;o APP</option>
        <option value="pont_op.html" selected>&#8987; Pontualidade Op.</option>
        <option value="evol_semanal.html">&#128200; App Semanal</option>
        <option value="app_mensal.html">&#128197; APP Mensal</option>
        <option value="infracoes.html">&#128680; Infra&#231;&#245;es</option>
        <option value="motoristas.html">&#128100; Motoristas APP</option>
        <option value="pont_semanal.html">&#128202; Pont. Semanal</option>
        <option value="pont_mensal.html">&#128198; Pont. Mensal</option>
        <option value="inf_pontualidade.html">&#128308; Inf. Pontualidade</option>
        <option value="pont_motoristas.html">&#128663; Pont. Motoristas</option>
      </select>
    </div>
    <div class="modal-row">
      <label>PAINEL 3</label>
      <select id="cfg-2">
        <option value="dashboard.html">&#128202; Utiliza&#231;&#227;o APP</option>
        <option value="pont_op.html">&#8987; Pontualidade Op.</option>
        <option value="evol_semanal.html">&#128200; App Semanal</option>
        <option value="app_mensal.html" selected>&#128197; APP Mensal</option>
        <option value="infracoes.html">&#128680; Infra&#231;&#245;es</option>
        <option value="motoristas.html">&#128100; Motoristas APP</option>
        <option value="pont_semanal.html">&#128202; Pont. Semanal</option>
        <option value="pont_mensal.html">&#128198; Pont. Mensal</option>
        <option value="inf_pontualidade.html">&#128308; Inf. Pontualidade</option>
        <option value="pont_motoristas.html">&#128663; Pont. Motoristas</option>
      </select>
    </div>
    <div class="modal-row">
      <label>PAINEL 4</label>
      <select id="cfg-3">
        <option value="dashboard.html">&#128202; Utiliza&#231;&#227;o APP</option>
        <option value="pont_op.html">&#8987; Pontualidade Op.</option>
        <option value="evol_semanal.html">&#128200; App Semanal</option>
        <option value="app_mensal.html">&#128197; APP Mensal</option>
        <option value="infracoes.html">&#128680; Infra&#231;&#245;es</option>
        <option value="motoristas.html">&#128100; Motoristas APP</option>
        <option value="pont_semanal.html">&#128202; Pont. Semanal</option>
        <option value="pont_mensal.html" selected>&#128198; Pont. Mensal</option>
        <option value="inf_pontualidade.html">&#128308; Inf. Pontualidade</option>
        <option value="pont_motoristas.html">&#128663; Pont. Motoristas</option>
      </select>
    </div>
    <div class="modal-btns">
      <button class="modal-cancel" onclick="closeConfig()">Cancelar</button>
      <button class="modal-ok" onclick="applyConfig()">&#10003; Aplicar</button>
    </div>
  </div>
</div>

<script>
// Clock
function tick(){
  const d=new Date();
  document.getElementById('mural-clock').textContent=
    String(d.getHours()).padStart(2,'0')+':'+
    String(d.getMinutes()).padStart(2,'0')+':'+
    String(d.getSeconds()).padStart(2,'0');
}
tick(); setInterval(tick,1000);

// Zoom por layout — ajusta a escala do iframe para caber no painel sem barra de rolagem
const ZOOM_MAP={'2x2':0.62,'1x4':0.82,'2x1':0.65,'1x1':0.95};
const ALL_LAYOUTS=['2x2','1x4','2x1','1x1'];
const LAYOUT_CSS={'2x2':'layout-2x2','1x4':'layout-1x4','2x1':'layout-2x1','1x1':'layout-1x1'};

function setPanelZoom(z){
  document.querySelectorAll('.mural-panel').forEach(p=>{
    p.style.setProperty('--pz', z);
  });
}

function setLayout(name){
  const grid=document.getElementById('mural-grid');
  ALL_LAYOUTS.forEach(l=>{
    grid.classList.remove(LAYOUT_CSS[l]);
    const btn=document.getElementById('btn-'+l);
    if(btn) btn.classList.remove('active');
  });
  grid.classList.add(LAYOUT_CSS[name]);
  const active=document.getElementById('btn-'+name);
  if(active) active.classList.add('active');
  setPanelZoom(ZOOM_MAP[name]||0.62);
}

// Injeta CSS no iframe para esconder a sidebar (mesmo origem)
function injectStyles(idx){
  try{
    const iframe=document.getElementById('iframe-'+idx);
    if(!iframe) return;
    const doc=iframe.contentDocument||iframe.contentWindow.document;
    if(!doc||!doc.head) return;
    let st=doc.getElementById('__mi__');
    if(!st){st=doc.createElement('style');st.id='__mi__';doc.head.appendChild(st);}
    st.textContent=
      '.sidebar{display:none!important;}'+
      '.main{margin-left:0!important;width:100%!important;}'+
      'body{overflow:hidden!important;}';
  }catch(e){}
}

function bindIframe(idx){
  const iframe=document.getElementById('iframe-'+idx);
  if(!iframe) return;
  iframe.addEventListener('load',function(){injectStyles(idx);});
}
[0,1,2,3].forEach(bindIframe);

// Panel control
const PAGE_NAMES={
  'dashboard.html':'Utilização APP','pont_op.html':'Pontualidade Op.',
  'evol_semanal.html':'App Semanal','app_mensal.html':'APP Mensal',
  'infracoes.html':'Infrações','motoristas.html':'Motoristas APP',
  'pont_semanal.html':'Pont. Semanal','pont_mensal.html':'Pont. Mensal',
  'inf_pontualidade.html':'Inf. Pontualidade','pont_motoristas.html':'Pont. Motoristas',
  'home.html':'Home'
};
function setPanel(idx,url,title){
  const iframe=document.getElementById('iframe-'+idx);
  const ptitle=document.getElementById('ptitle-'+idx);
  if(iframe) iframe.src=url;
  if(ptitle) ptitle.textContent=title||PAGE_NAMES[url]||url;
}
function reloadPanel(idx){
  const iframe=document.getElementById('iframe-'+idx);
  if(iframe){const s=iframe.src;iframe.src='about:blank';setTimeout(()=>iframe.src=s,50);}
}
function maximizePanel(idx){
  const iframe=document.getElementById('iframe-'+idx);
  const title=document.getElementById('ptitle-'+idx);
  const overlay=document.getElementById('max-overlay');
  if(iframe&&overlay){
    document.getElementById('max-iframe').src=iframe.src;
    document.getElementById('max-title').textContent=(title?title.textContent:'Painel '+(idx+1))+' — Maximizado';
    overlay.classList.add('open');
  }
}
function closeMaximize(){
  document.getElementById('max-overlay').classList.remove('open');
  document.getElementById('max-iframe').src='about:blank';
}

// Fullscreen
function toggleFullscreen(){
  if(!document.fullscreenElement){
    document.documentElement.requestFullscreen().catch(()=>{});
  }else{
    document.exitFullscreen().catch(()=>{});
  }
}
document.addEventListener('fullscreenchange',()=>{
  const btn=document.getElementById('btn-full');
  if(btn) btn.innerHTML=document.fullscreenElement?'&#9974; Sair Tela Cheia':'&#9974; Tela Cheia';
});

// Config Modal
function openConfig(){
  for(let i=0;i<4;i++){
    const iframe=document.getElementById('iframe-'+i);
    const sel=document.getElementById('cfg-'+i);
    if(iframe&&sel){
      try{const fname=new URL(iframe.src).pathname.split('/').pop();sel.value=fname||iframe.src;}catch(e){}
    }
  }
  document.getElementById('config-modal').classList.add('open');
}
function closeConfig(){
  document.getElementById('config-modal').classList.remove('open');
}
function applyConfig(){
  for(let i=0;i<4;i++){
    const sel=document.getElementById('cfg-'+i);
    if(sel&&sel.value) setPanel(i,sel.value,PAGE_NAMES[sel.value]||sel.value);
  }
  closeConfig();
}
</script>
</body>
</html>"""


# ══════════════════════════════════════════════════════════════════════════════
# GERA TODOS OS ARQUIVOS
# ══════════════════════════════════════════════════════════════════════════════

def gerar_dashboard(pasta="brdrive_output", dados=None):
    os.makedirs(pasta, exist_ok=True)
    df = carregar_dados(dados)
    print(f"📊 Dados carregados: {len(df)} linhas")

    paginas = {
        "home": (pg_home, None),
        "mural": (pg_mural, None),
        "dashboard": (pg_utilizacao_app, df),
        "pont_op": (pg_pontualidade_op, df),
        "evol_semanal": (pg_evol_semanal, df),
        "infracoes": (pg_infracoes, df),
        "app_mensal": (pg_app_mensal, df),
        "motoristas": (pg_motoristas, df),
        "pont_semanal": (pg_pont_semanal, df),
        "pont_mensal": (pg_pont_mensal, df),
        "inf_pontualidade": (pg_inf_pontualidade, df),
        "pont_motoristas": (pg_pont_motoristas, df),
    }

    paginas_simples = []

    # Gera páginas completas
    for nome, (fn, dados_pg) in paginas.items():
        if dados_pg is None:
            html = fn()
        else:
            html = fn(dados_pg)
        with open(f"{pasta}/{nome}.html", "w", encoding="utf-8") as f:
            f.write(html)
        print(f"  ✅ {nome}.html")

    # Gera páginas em construção
    for pid, titulo, cor in paginas_simples:
        html = pg_em_construcao(titulo, pid, cor)
        with open(f"{pasta}/{pid}.html", "w", encoding="utf-8") as f:
            f.write(html)
        print(f"  🔨 {pid}.html (aguarda dados)")

    print(f"\n🎉 Dashboard gerado em: {pasta}/")
    print(f"🌐 Abrindo no navegador...")
    webbrowser.open("file://" + os.path.abspath(f"{pasta}/home.html"))

if __name__ == "__main__":
    dados = sys.argv[1] if len(sys.argv) > 1 else None
    gerar_dashboard(dados=dados)
