"""
Gerador de Dashboard HTML — Análise MDTR por Setor (Alcon)
============================================================
Lê o XLSX MDTRS e gera um arquivo HTML standalone com:
  - Filtros interativos (Gerente, Setor, UF, Bandeira, Marca)
  - Slides por setor + slide de produto
  - Botão de exportação para PowerPoint (PptxGenJS via CDN)

USO:
    python mdtrs_gerar_html.py [caminho_excel.xlsx] [caminho_saida.html]

Padrões:
    Entrada: /mnt/user-data/uploads/MDTRS_FV_TOTAL.xlsx
    Saída:   /mnt/user-data/outputs/MDTRS_Dashboard.html

REUTILIZAÇÃO MENSAL:
    Edite as constantes em CONFIG (MES_CORRENTE, MES_ANTERIOR, YTD_COLS) e rode novamente.
    O HTML resultante é único arquivo, pode ser enviado por email/Teams.
"""

import os
import sys
import json
import warnings
from datetime import datetime

import pandas as pd

warnings.filterwarnings("ignore")

# ============================================================
# CONFIG — edite aqui para virada de mês
# ============================================================
DEFAULT_INPUT = "/mnt/user-data/uploads/MDTRS_FV_TOTAL.xlsx"
DEFAULT_OUTPUT = "/mnt/user-data/outputs/MDTRS_Dashboard.html"
DEFAULT_SHEET = "TOTAL"

TIPO_INFO_ANALISE = "SO"           # SO (sell-out) ou SI_NP (sell-in)
EXCLUIR_NAO_VISITADO = True
THRESHOLD_VARIACAO = 40             # destaque de PDV (variação ≥ X unid.)
THRESHOLD_PDV_PAYLOAD = 20          # PDVs incluídos no payload (margem p/ filtros)

MES_CORRENTE = "202605"             # Mai/26 (com S1, S2, S3, Proj)
MES_ANTERIOR = "202604"             # Abr/26

YTD_COLS = ["202601", "202602", "202603", "202604"]
YTD_ANT_COLS = ["202509", "202510", "202511", "202512"]
YTD_LABEL = "YTD 2026 (Jan-Abr)"
YTD_ANT_LABEL = "Set-Dez 2025"

TOP_BANDEIRAS = 30                  # bandeiras mantidas individualmente; resto vira "OUTRAS"


# ============================================================
# CARREGAMENTO
# ============================================================
def load_data(path, sheet=DEFAULT_SHEET):
    """Lê o XLSX usando cache pickle quando disponível."""
    cache_dir = os.path.join(os.path.expanduser("~"), ".mdtrs_cache")
    os.makedirs(cache_dir, exist_ok=True)
    cache_path = os.path.join(cache_dir, os.path.basename(path).replace(os.sep, "_") + ".pkl")

    if os.path.exists(cache_path) and os.path.getmtime(cache_path) >= os.path.getmtime(path):
        print(f"[1/4] Lendo cache {cache_path} ...")
        df = pd.read_pickle(cache_path)
    else:
        print(f"[1/4] Lendo {path} ...")
        try:
            df = pd.read_excel(path, sheet_name=sheet, engine="calamine")
        except Exception:
            df = pd.read_excel(path, sheet_name=sheet, engine="openpyxl")
        try:
            df.to_pickle(cache_path)
        except Exception as e:
            print(f"      (aviso: cache não salvo: {e})")

    period_cols = [c for c in df.columns if str(c).startswith("20")]
    for c in period_cols:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    print(f"      Linhas: {len(df):,}")
    return df


# ============================================================
# PREPARAÇÃO DE DADOS PARA O HTML
# ============================================================
def build_payload(df):
    """Constrói o objeto JSON que será embutido no HTML."""
    print("[2/4] Filtrando e agregando dados ...")
    df = df[df["TIPO_INFORMACAO"] == TIPO_INFO_ANALISE].copy()
    if EXCLUIR_NAO_VISITADO:
        df = df[df["SETOR_NOME"] != "NÃO VISITADO"]
    print(f"      Linhas filtradas: {len(df):,}")

    # Normaliza bandeira: top N e "OUTRAS"
    top_bands = df["Bandeira"].value_counts().head(TOP_BANDEIRAS).index.tolist()
    df["Bandeira_n"] = df["Bandeira"].where(df["Bandeira"].isin(top_bands), "OUTRAS")
    df["MARCA_clean"] = df["MARCA"].str.replace(" (ALC)", "", regex=False)

    period_cols = (YTD_COLS + [f"{MES_CORRENTE}-S1", f"{MES_CORRENTE}-S2",
                                f"{MES_CORRENTE}-S3", f"{MES_CORRENTE}-Proj"])

    # -------------- AGREGADO PRINCIPAL (rows) --------------
    # Por (Gerente x Setor x UF x Bandeira x Marca) com séries
    agg = df.groupby(["GERENTE", "SETOR_NOME", "UF", "Bandeira_n", "MARCA_clean"],
                     as_index=False)[period_cols].sum()
    agg = agg[(agg[period_cols].sum(axis=1) > 0)]

    # Catálogos
    gerentes = sorted(agg["GERENTE"].unique().tolist())
    setores = sorted(agg["SETOR_NOME"].unique().tolist())
    ufs = sorted(agg["UF"].dropna().unique().tolist())
    bandeiras = sorted(agg["Bandeira_n"].unique().tolist())
    marcas = sorted(agg["MARCA_clean"].unique().tolist())

    gid = {g: i for i, g in enumerate(gerentes)}
    sid = {s: i for i, s in enumerate(setores)}
    uid = {u: i for i, u in enumerate(ufs)}
    bid = {b: i for i, b in enumerate(bandeiras)}
    mid = {m: i for i, m in enumerate(marcas)}

    # Mapa setor → gerente (para reconstruir filtros em cascata)
    setor_gerente = {}
    for _, r in agg.groupby(["SETOR_NOME", "GERENTE"]).size().reset_index().iterrows():
        setor_gerente[sid[r["SETOR_NOME"]]] = gid[r["GERENTE"]]

    rows = []
    for _, r in agg.iterrows():
        rows.append([
            gid[r["GERENTE"]], sid[r["SETOR_NOME"]], uid[r["UF"]],
            bid[r["Bandeira_n"]], mid[r["MARCA_clean"]],
            int(round(r["202601"])), int(round(r["202602"])),
            int(round(r["202603"])), int(round(r["202604"])),
            int(round(r[f"{MES_CORRENTE}-S1"])), int(round(r[f"{MES_CORRENTE}-S2"])),
            int(round(r[f"{MES_CORRENTE}-S3"])), int(round(r[f"{MES_CORRENTE}-Proj"])),
        ])

    # -------------- PDVs RELEVANTES (com |variação| ≥ THRESHOLD_PDV_PAYLOAD) --------------
    pdv_agg = df.groupby(
        ["CNPJ", "PDV", "CIDADE", "UF", "Bandeira_n", "GERENTE", "SETOR_NOME"],
        as_index=False
    ).agg({MES_ANTERIOR: "sum", f"{MES_CORRENTE}-Proj": "sum",
           f"{MES_CORRENTE}-S2": "sum", f"{MES_CORRENTE}-S3": "sum"})
    pdv_agg["var"] = pdv_agg[f"{MES_CORRENTE}-Proj"] - pdv_agg[MES_ANTERIOR]
    pdv_agg = pdv_agg[pdv_agg["var"].abs() >= THRESHOLD_PDV_PAYLOAD].copy()

    # Limpeza do nome do PDV
    pdv_agg["PDV_clean"] = pdv_agg["PDV"].astype(str).str.split(" - ").str[0].str[:45]
    pdv_agg["CIDADE_clean"] = pdv_agg["CIDADE"].astype(str).str.split(" - ").str[0].str[:25]

    pdvs = []
    for _, r in pdv_agg.iterrows():
        if r["GERENTE"] not in gid or r["SETOR_NOME"] not in sid:
            continue
        pdvs.append([
            gid[r["GERENTE"]], sid[r["SETOR_NOME"]],
            uid.get(r["UF"], -1),
            bid.get(r["Bandeira_n"], -1),
            r["PDV_clean"], r["CIDADE_clean"], r["UF"],
            int(round(r[MES_ANTERIOR])),
            int(round(r[f"{MES_CORRENTE}-Proj"])),
            int(round(r[f"{MES_CORRENTE}-S2"])),
            int(round(r[f"{MES_CORRENTE}-S3"])),
            int(round(r["var"])),
        ])

    # -------------- PRODUTO (Marca + Apresentação) --------------
    df["_ytd"] = df[YTD_COLS].sum(axis=1)
    df["_ytd_ant"] = df[YTD_ANT_COLS].sum(axis=1)

    # Por marca: com IDs de filtragem
    marca_rows = []
    g_marca = df.groupby(["GERENTE", "SETOR_NOME", "UF", "Bandeira_n", "MARCA_clean"],
                          as_index=False).agg({"_ytd": "sum", "_ytd_ant": "sum"})
    for _, r in g_marca.iterrows():
        if r["GERENTE"] not in gid or r["SETOR_NOME"] not in sid:
            continue
        marca_rows.append([
            gid[r["GERENTE"]], sid[r["SETOR_NOME"]], uid.get(r["UF"], -1),
            bid.get(r["Bandeira_n"], -1), mid[r["MARCA_clean"]],
            int(round(r["_ytd"])), int(round(r["_ytd_ant"])),
        ])

    # Por apresentação
    apresentacoes = sorted(df["APRESENTACAO"].dropna().unique().tolist())
    aid = {a: i for i, a in enumerate(apresentacoes)}
    apres_rows = []
    g_apres = df.groupby(["GERENTE", "SETOR_NOME", "UF", "Bandeira_n", "APRESENTACAO"],
                          as_index=False).agg({"_ytd": "sum", "_ytd_ant": "sum"})
    for _, r in g_apres.iterrows():
        if r["GERENTE"] not in gid or r["SETOR_NOME"] not in sid:
            continue
        apres_rows.append([
            gid[r["GERENTE"]], sid[r["SETOR_NOME"]], uid.get(r["UF"], -1),
            bid.get(r["Bandeira_n"], -1), aid[r["APRESENTACAO"]],
            int(round(r["_ytd"])), int(round(r["_ytd_ant"])),
        ])

    # Total PDVs por setor (sem filtros) — para mostrar "X de Y PDVs ativos"
    pdvs_por_setor = df.groupby("SETOR_NOME")["CNPJ"].nunique().to_dict()
    pdvs_por_setor_arr = [pdvs_por_setor.get(s, 0) for s in setores]

    return {
        "meta": {
            "gerado_em": datetime.now().strftime("%d/%m/%Y %H:%M"),
            "tipo": TIPO_INFO_ANALISE,
            "mes_corrente": MES_CORRENTE,
            "mes_anterior": MES_ANTERIOR,
            "ytd_label": YTD_LABEL,
            "ytd_ant_label": YTD_ANT_LABEL,
            "threshold": THRESHOLD_VARIACAO,
        },
        "gerentes": gerentes,
        "setores": setores,
        "setor_gerente": setor_gerente,
        "ufs": ufs,
        "bandeiras": bandeiras,
        "marcas": marcas,
        "apresentacoes": apresentacoes,
        "pdvs_por_setor": pdvs_por_setor_arr,
        # Schema: [g, s, u, b, m, jan, fev, mar, abr, s1, s2, s3, proj]
        "rows": rows,
        # Schema: [g, s, u, b, pdv_name, cidade, uf_raw, abr, proj, s2, s3, var]
        "pdvs": pdvs,
        # Schema: [g, s, u, b, m, ytd, ytd_ant]
        "marca_rows": marca_rows,
        # Schema: [g, s, u, b, a, ytd, ytd_ant]
        "apres_rows": apres_rows,
    }


# ============================================================
# GERAÇÃO DO HTML
# ============================================================
def _baixar_libs():
    """Baixa Chart.js e PptxGenJS via npm registry para embutir no HTML.
    Necessário para ambiente corporativo que bloqueia CDNs (jsDelivr/unpkg)."""
    import urllib.request
    import tarfile
    import io as bio

    cache_dir = os.path.join(os.path.expanduser("~"), ".mdtrs_cache")
    os.makedirs(cache_dir, exist_ok=True)
    chart_cache = os.path.join(cache_dir, "chart.umd.js")
    pptx_cache = os.path.join(cache_dir, "pptxgen.bundle.js")

    if os.path.exists(chart_cache) and os.path.exists(pptx_cache):
        with open(chart_cache, "r", encoding="utf-8") as f: chart_js = f.read()
        with open(pptx_cache, "r", encoding="utf-8") as f: pptx_js = f.read()
        return chart_js, pptx_js

    print("      Baixando libs JS (uma única vez)...")
    libs = {
        "chart": ("https://registry.npmjs.org/chart.js/-/chart.js-4.4.0.tgz",
                  "package/dist/chart.umd.js", chart_cache),
        "pptx": ("https://registry.npmjs.org/pptxgenjs/-/pptxgenjs-3.12.0.tgz",
                 "package/dist/pptxgen.bundle.js", pptx_cache),
    }
    js_files = {}
    for name, (url, member, cache_path) in libs.items():
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=30) as resp:
            tgz_data = resp.read()
        with tarfile.open(fileobj=bio.BytesIO(tgz_data), mode="r:gz") as tar:
            f = tar.extractfile(member)
            js = f.read().decode("utf-8")
        with open(cache_path, "w", encoding="utf-8") as out:
            out.write(js)
        js_files[name] = js
        print(f"        ✓ {name} ({len(js)/1024:.0f} KB)")
    return js_files["chart"], js_files["pptx"]


def build_html(payload, output_path):
    print("[3/4] Gerando HTML ...")
    chart_js, pptx_js = _baixar_libs()
    payload_json = json.dumps(payload, separators=(",", ":"), ensure_ascii=False)
    html = (HTML_TEMPLATE
            .replace("{{PAYLOAD_JSON}}", payload_json)
            .replace("{{CHART_JS}}", chart_js)
            .replace("{{PPTX_JS}}", pptx_js))
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)
    size_kb = os.path.getsize(output_path) / 1024
    print(f"[4/4] Salvo em {output_path} ({size_kb:.0f} KB)")


# ============================================================
# TEMPLATE HTML (com filtros + Chart.js + PptxGenJS via CDN)
# ============================================================
HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>MDTR · Análise por Setor — Alcon</title>
<script>
/* Chart.js (UMD build, embutido offline) */
{{CHART_JS}}
</script>
<script>
/* PptxGenJS (bundle, embutido offline) */
{{PPTX_JS}}
</script>
<style>
  :root {
    --navy: #1E2761;
    --navy-dark: #161e4d;
    --ice: #CADCFC;
    --gold: #C9A227;
    --green: #2E7D32;
    --red: #C62828;
    --gray: #595959;
    --gray-light: #E8E8E8;
    --bg: #F4F5F8;
    --card: #FFFFFF;
  }
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body {
    font-family: 'Segoe UI', Calibri, Arial, sans-serif;
    background: var(--bg);
    color: #222;
    font-size: 13px;
    line-height: 1.4;
  }
  /* Top bar */
  header {
    background: var(--navy);
    color: white;
    padding: 14px 24px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    position: sticky; top: 0; z-index: 100;
  }
  header h1 {
    font-family: Georgia, serif;
    font-size: 20px;
    font-weight: bold;
  }
  header h1 small {
    font-family: 'Segoe UI', sans-serif;
    font-size: 11px;
    color: var(--ice);
    font-weight: normal;
    margin-left: 12px;
    letter-spacing: 0.5px;
  }
  header .meta { font-size: 11px; color: var(--ice); }

  /* Filter bar */
  .filters {
    background: white;
    padding: 14px 24px;
    border-bottom: 1px solid var(--gray-light);
    display: flex;
    flex-wrap: wrap;
    gap: 12px;
    align-items: end;
  }
  .filter-group { display: flex; flex-direction: column; gap: 4px; min-width: 160px; }
  .filter-group label {
    font-size: 10px;
    font-weight: bold;
    color: var(--gray);
    text-transform: uppercase;
    letter-spacing: 0.5px;
  }
  .filter-group select {
    padding: 7px 10px;
    border: 1px solid #ccc;
    border-radius: 4px;
    background: white;
    font-size: 13px;
    color: var(--navy);
    cursor: pointer;
    min-width: 180px;
  }
  .filter-group select:focus { outline: 2px solid var(--navy); border-color: var(--navy); }
  .filter-actions { display: flex; gap: 8px; margin-left: auto; }
  button {
    padding: 8px 16px;
    border: none;
    border-radius: 4px;
    font-weight: bold;
    cursor: pointer;
    font-size: 12px;
    transition: all 0.15s;
  }
  .btn-primary { background: var(--navy); color: white; }
  .btn-primary:hover { background: var(--navy-dark); }
  .btn-secondary { background: white; color: var(--navy); border: 1px solid var(--navy); }
  .btn-secondary:hover { background: var(--bg); }
  .btn-gold { background: var(--gold); color: white; }
  .btn-gold:hover { background: #a78318; }

  /* Summary bar */
  .summary {
    background: var(--navy);
    color: white;
    padding: 10px 24px;
    display: flex;
    gap: 32px;
    font-size: 12px;
    border-top: 3px solid var(--gold);
  }
  .summary .item strong {
    display: block;
    font-size: 18px;
    font-weight: bold;
    color: var(--ice);
  }
  .summary .item span {
    color: rgba(255,255,255,0.7);
    font-size: 10px;
    text-transform: uppercase;
    letter-spacing: 0.5px;
  }

  /* Main */
  main { padding: 20px 24px; }
  .section-title {
    font-family: Georgia, serif;
    font-size: 16px;
    font-weight: bold;
    color: var(--navy);
    margin: 16px 0 10px 0;
    padding-bottom: 6px;
    border-bottom: 2px solid var(--gold);
  }

  /* Setor cards grid */
  #setor-list {
    display: grid;
    grid-template-columns: repeat(auto-fill, minmax(380px, 1fr));
    gap: 16px;
  }
  .setor-card {
    background: var(--card);
    border-radius: 6px;
    padding: 14px;
    box-shadow: 0 1px 3px rgba(0,0,0,0.08);
    border: 1px solid var(--gray-light);
    cursor: pointer;
    transition: all 0.15s;
  }
  .setor-card:hover { box-shadow: 0 3px 8px rgba(0,0,0,0.15); border-color: var(--navy); }
  .setor-card h3 {
    color: var(--navy);
    font-family: Georgia, serif;
    font-size: 14px;
    margin-bottom: 2px;
  }
  .setor-card .gerente { font-size: 10px; color: var(--gray); margin-bottom: 8px; }
  .setor-card .kpis {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 8px;
    margin-bottom: 8px;
  }
  .setor-card .kpi {
    background: var(--bg);
    padding: 6px 8px;
    border-radius: 4px;
    border-left: 3px solid var(--navy);
  }
  .setor-card .kpi small {
    font-size: 9px;
    color: var(--gray);
    text-transform: uppercase;
    display: block;
  }
  .setor-card .kpi b {
    font-size: 14px;
    color: var(--navy);
  }
  .setor-card .kpi.gold { border-left-color: var(--gold); }
  .setor-card .kpi.gold b { color: var(--gold); }
  .setor-card .kpi.green { border-left-color: var(--green); }
  .setor-card .kpi.green b { color: var(--green); }
  .setor-card .kpi.red { border-left-color: var(--red); }
  .setor-card .kpi.red b { color: var(--red); }
  .setor-card canvas { height: 70px !important; }

  /* Modal */
  .modal-bg {
    display: none;
    position: fixed; top: 0; left: 0; right: 0; bottom: 0;
    background: rgba(0,0,0,0.6);
    z-index: 200;
    overflow-y: auto;
    padding: 20px;
  }
  .modal-bg.open { display: block; }
  .modal {
    background: white;
    border-radius: 8px;
    max-width: 1280px;
    margin: 0 auto;
    overflow: hidden;
  }
  .modal-header {
    background: var(--navy);
    color: white;
    padding: 16px 24px;
    display: flex;
    justify-content: space-between;
    align-items: center;
    border-bottom: 3px solid var(--gold);
  }
  .modal-header h2 { font-family: Georgia, serif; font-size: 22px; }
  .modal-header .gerente { font-size: 12px; color: var(--ice); }
  .modal-close {
    background: transparent;
    color: white;
    font-size: 22px;
    border: 1px solid white;
    width: 32px; height: 32px;
    border-radius: 4px;
    display: flex; align-items: center; justify-content: center;
  }
  .modal-body { padding: 20px; }
  .kpi-row {
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    gap: 12px;
    margin-bottom: 20px;
  }
  .kpi-card-lg {
    padding: 12px 14px;
    border-radius: 6px;
    background: white;
    border: 1px solid var(--gray-light);
    border-left-width: 4px;
  }
  .kpi-card-lg small {
    font-size: 9.5px;
    color: var(--gray);
    text-transform: uppercase;
    font-weight: bold;
    letter-spacing: 0.3px;
  }
  .kpi-card-lg .val {
    font-size: 24px;
    font-weight: bold;
    color: var(--navy);
    margin: 4px 0;
  }
  .kpi-card-lg .sub { font-size: 11px; color: var(--gray); font-weight: bold; }
  .kpi-card-lg.gold { border-left-color: var(--gold); }
  .kpi-card-lg.gold .val { color: var(--gold); }
  .kpi-card-lg.green { border-left-color: var(--green); }
  .kpi-card-lg.green .val { color: var(--green); }
  .kpi-card-lg.red { border-left-color: var(--red); }
  .kpi-card-lg.red .val { color: var(--red); }
  .sub.green { color: var(--green); }
  .sub.red { color: var(--red); }

  .charts-row {
    display: grid;
    grid-template-columns: 2fr 1fr;
    gap: 16px;
    margin-bottom: 20px;
  }
  .chart-box {
    background: white;
    padding: 14px;
    border: 1px solid var(--gray-light);
    border-radius: 6px;
  }
  .chart-box h4 {
    color: var(--navy);
    font-size: 12px;
    margin-bottom: 8px;
    font-weight: bold;
  }
  .chart-box canvas { max-height: 200px; }

  .tops-row {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 16px;
  }
  .top-table {
    background: white;
    padding: 14px;
    border: 1px solid var(--gray-light);
    border-radius: 6px;
  }
  .top-table h4 {
    font-size: 12px;
    margin-bottom: 8px;
    font-weight: bold;
  }
  .top-table h4.up { color: var(--green); }
  .top-table h4.down { color: var(--red); }
  table {
    width: 100%;
    border-collapse: collapse;
    font-size: 11px;
  }
  th {
    background: var(--navy);
    color: white;
    text-align: left;
    padding: 6px 8px;
    font-weight: bold;
    font-size: 10px;
  }
  th.num, td.num { text-align: right; }
  td {
    padding: 5px 8px;
    border-bottom: 1px solid #f0f0f0;
  }
  tr:nth-child(even) td { background: #fafafa; }
  td.green { color: var(--green); font-weight: bold; }
  td.red { color: var(--red); font-weight: bold; }
  .empty {
    text-align: center;
    color: var(--gray);
    padding: 30px 10px;
    font-style: italic;
  }

  /* Produto */
  .produto-grid {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 16px;
    margin-bottom: 20px;
  }
  #produto-section { background: white; padding: 20px; border-radius: 6px; border: 1px solid var(--gray-light); }
  #apres-table { width: 100%; }

  /* Export modal */
  .export-modal {
    display: none;
    position: fixed; top: 50%; left: 50%;
    transform: translate(-50%, -50%);
    background: white;
    border-radius: 8px;
    padding: 24px;
    box-shadow: 0 10px 40px rgba(0,0,0,0.3);
    z-index: 300;
    min-width: 380px;
  }
  .export-modal.open { display: block; }
  .export-modal h3 {
    color: var(--navy);
    font-family: Georgia, serif;
    margin-bottom: 12px;
  }
  .export-modal label {
    display: block;
    padding: 10px;
    border: 1px solid var(--gray-light);
    border-radius: 4px;
    margin-bottom: 8px;
    cursor: pointer;
  }
  .export-modal label:hover { background: var(--bg); }
  .export-modal label input { margin-right: 8px; }
  .export-modal .actions {
    display: flex;
    gap: 8px;
    justify-content: flex-end;
    margin-top: 16px;
  }
  #progress {
    display: none;
    position: fixed; bottom: 24px; right: 24px;
    background: var(--navy);
    color: white;
    padding: 12px 20px;
    border-radius: 6px;
    box-shadow: 0 4px 12px rgba(0,0,0,0.3);
    z-index: 400;
  }
  #progress.show { display: block; }
</style>
</head>
<body>
<header>
  <h1>MDTR · Análise por Setor <small>ALCON · COMMERCIAL INTELLIGENCE</small></h1>
  <div class="meta" id="meta-info">Carregando…</div>
</header>

<div class="filters">
  <div class="filter-group">
    <label>Gerente</label>
    <select id="f-gerente"><option value="">Todos</option></select>
  </div>
  <div class="filter-group">
    <label>Setor</label>
    <select id="f-setor"><option value="">Todos</option></select>
  </div>
  <div class="filter-group">
    <label>UF</label>
    <select id="f-uf"><option value="">Todas</option></select>
  </div>
  <div class="filter-group">
    <label>Bandeira</label>
    <select id="f-bandeira"><option value="">Todas</option></select>
  </div>
  <div class="filter-group">
    <label>Marca</label>
    <select id="f-marca"><option value="">Todas</option></select>
  </div>
  <div class="filter-actions">
    <button class="btn-secondary" id="btn-reset">Limpar filtros</button>
    <button class="btn-gold" id="btn-export">⬇ Exportar PowerPoint</button>
  </div>
</div>

<div class="summary" id="summary">
  <div class="item"><strong id="s-setores">—</strong><span>Setores</span></div>
  <div class="item"><strong id="s-ytd">—</strong><span id="s-ytd-label">YTD 2026 (un.)</span></div>
  <div class="item"><strong id="s-mai">—</strong><span>Mai-Proj (un.)</span></div>
  <div class="item"><strong id="s-var">—</strong><span>Δ Mai vs Abr</span></div>
  <div class="item"><strong id="s-s3s2">—</strong><span>Δ S3 vs S2</span></div>
  <div class="item"><strong id="s-pdvs">—</strong><span>PDVs em destaque (↑↓)</span></div>
</div>

<main>
  <div class="section-title">Análise de Produto</div>
  <div id="produto-section">
    <div class="produto-grid">
      <div class="chart-box">
        <h4 id="prod-titulo-1">Ranking por Marca — YTD</h4>
        <canvas id="chart-marca-ytd"></canvas>
      </div>
      <div class="chart-box">
        <h4 id="prod-titulo-2">Variação por Marca — YTD vs Período Anterior</h4>
        <canvas id="chart-marca-var"></canvas>
      </div>
    </div>
    <h4 style="color:var(--navy);font-size:12px;margin-bottom:8px;">Por Apresentação</h4>
    <table id="apres-table">
      <thead><tr>
        <th>Apresentação</th>
        <th class="num" id="apres-h1">Período Anterior</th>
        <th class="num" id="apres-h2">YTD Atual</th>
        <th class="num">Δ unid</th>
        <th class="num">Δ %</th>
      </tr></thead>
      <tbody id="apres-tbody"></tbody>
    </table>
  </div>

  <div class="section-title">Setores <span id="contador-setores" style="font-size:12px;font-weight:normal;color:var(--gray);"></span></div>
  <div id="setor-list"></div>
</main>

<!-- Modal detalhe do setor -->
<div class="modal-bg" id="modal">
  <div class="modal">
    <div class="modal-header">
      <div>
        <h2 id="m-titulo">—</h2>
        <div class="gerente" id="m-gerente">—</div>
      </div>
      <button class="modal-close" onclick="fecharModal()">×</button>
    </div>
    <div class="modal-body">
      <div class="kpi-row" id="m-kpis"></div>
      <div class="charts-row">
        <div class="chart-box">
          <h4>Evolução Mensal — Sell-Out (unid.)</h4>
          <canvas id="m-chart-evo"></canvas>
        </div>
        <div class="chart-box">
          <h4>Desempenho Semanal — Maio</h4>
          <canvas id="m-chart-sem"></canvas>
        </div>
      </div>
      <div class="tops-row">
        <div class="top-table">
          <h4 class="up">↑ Top 5 PDVs em ALTA (Mai-Proj vs Abr ≥ +<span class="thr"></span> unid)</h4>
          <table>
            <thead><tr><th>PDV — Cidade/UF</th><th class="num">Abr</th><th class="num">Mai-Proj</th><th class="num">Δ unid</th></tr></thead>
            <tbody id="m-tbl-altas"></tbody>
          </table>
        </div>
        <div class="top-table">
          <h4 class="down">↓ Top 5 PDVs em QUEDA (Mai-Proj vs Abr ≤ −<span class="thr"></span> unid)</h4>
          <table>
            <thead><tr><th>PDV — Cidade/UF</th><th class="num">Abr</th><th class="num">Mai-Proj</th><th class="num">Δ unid</th></tr></thead>
            <tbody id="m-tbl-quedas"></tbody>
          </table>
        </div>
      </div>
    </div>
  </div>
</div>

<!-- Export options modal -->
<div class="export-modal" id="export-modal">
  <h3>Exportar para PowerPoint</h3>
  <p style="font-size:12px;color:var(--gray);margin-bottom:14px;">Escolha o escopo do PPT:</p>
  <label><input type="radio" name="export-scope" value="filtered" checked> Apenas setores visíveis (com filtros aplicados) — <span id="exp-count">—</span> setores</label>
  <label><input type="radio" name="export-scope" value="all"> Todos os 76 setores (ignora filtros)</label>
  <div class="actions">
    <button class="btn-secondary" onclick="fecharExport()">Cancelar</button>
    <button class="btn-primary" onclick="exportarPPT()">Gerar PPT</button>
  </div>
</div>

<div id="progress">Gerando PowerPoint… <span id="progress-pct">0%</span></div>

<script>
// ============================================================
// DADOS EMBUTIDOS
// ============================================================
const DATA = {{PAYLOAD_JSON}};

// ============================================================
// CONSTANTES VISUAIS
// ============================================================
const COLORS = {
  navy: '#1E2761', ice: '#CADCFC', gold: '#C9A227',
  green: '#2E7D32', red: '#C62828', gray: '#595959', grayLight: '#E8E8E8'
};

// Schema de rows: [g, s, u, b, m, jan, fev, mar, abr, s1, s2, s3, proj]
const IDX = { G:0, S:1, U:2, B:3, M:4, JAN:5, FEV:6, MAR:7, ABR:8, S1:9, S2:10, S3:11, PROJ:12 };
// pdvs: [g, s, u, b, pdv_name, cidade, uf_raw, abr, proj, s2, s3, var]
const PIDX = { G:0, S:1, U:2, B:3, NAME:4, CITY:5, UF:6, ABR:7, PROJ:8, S2:9, S3:10, VAR:11 };
// marca_rows: [g, s, u, b, m, ytd, ytd_ant]
const MIDX = { G:0, S:1, U:2, B:3, M:4, YTD:5, YTD_ANT:6 };
// apres_rows: [g, s, u, b, a, ytd, ytd_ant]
const AIDX = { G:0, S:1, U:2, B:3, A:4, YTD:5, YTD_ANT:6 };

let filterState = { gerente: '', setor: '', uf: '', bandeira: '', marca: '' };
let chartInstances = {};

// ============================================================
// UTIL
// ============================================================
function fmtNum(v) {
  if (v == null || isNaN(v)) return '—';
  return Math.round(v).toLocaleString('pt-BR');
}
function fmtPct(v, casas=1) {
  if (v == null || isNaN(v) || !isFinite(v)) return '—';
  const s = (v >= 0 ? '+' : '') + v.toFixed(casas).replace('.', ',') + '%';
  return s;
}
function safe(s) { return String(s || '').replace(/</g, '&lt;'); }

// ============================================================
// FILTRAGEM
// ============================================================
function rowMatches(row, schema = IDX) {
  // Mapeia o estado de filtro pra IDs
  const gid = filterState.gerente === '' ? -1 : DATA.gerentes.indexOf(filterState.gerente);
  const sid = filterState.setor === '' ? -1 : DATA.setores.indexOf(filterState.setor);
  const uid = filterState.uf === '' ? -1 : DATA.ufs.indexOf(filterState.uf);
  const bid = filterState.bandeira === '' ? -1 : DATA.bandeiras.indexOf(filterState.bandeira);
  const mid = filterState.marca === '' ? -1 : DATA.marcas.indexOf(filterState.marca);

  if (gid !== -1 && row[schema.G] !== gid) return false;
  if (sid !== -1 && row[schema.S] !== sid) return false;
  if (uid !== -1 && row[schema.U] !== uid) return false;
  if (bid !== -1 && row[schema.B] !== bid) return false;
  if (mid !== -1 && schema.M !== undefined && row[schema.M] !== mid) return false;
  return true;
}

function rowsFiltradas() { return DATA.rows.filter(r => rowMatches(r, IDX)); }
function pdvsFiltrados() {
  // PDVs ignoram filtro de marca (não temos marca no nível de PDV agregado)
  return DATA.pdvs.filter(r => {
    const gid = filterState.gerente === '' ? -1 : DATA.gerentes.indexOf(filterState.gerente);
    const sid = filterState.setor === '' ? -1 : DATA.setores.indexOf(filterState.setor);
    const uid = filterState.uf === '' ? -1 : DATA.ufs.indexOf(filterState.uf);
    const bid = filterState.bandeira === '' ? -1 : DATA.bandeiras.indexOf(filterState.bandeira);
    if (gid !== -1 && r[PIDX.G] !== gid) return false;
    if (sid !== -1 && r[PIDX.S] !== sid) return false;
    if (uid !== -1 && r[PIDX.U] !== uid) return false;
    if (bid !== -1 && r[PIDX.B] !== bid) return false;
    return true;
  });
}
function marcaRowsFiltradas() { return DATA.marca_rows.filter(r => rowMatches(r, MIDX)); }
function apresRowsFiltradas() {
  // Sem filtro de marca aqui (esse dataset usa apresentação)
  return DATA.apres_rows.filter(r => {
    const gid = filterState.gerente === '' ? -1 : DATA.gerentes.indexOf(filterState.gerente);
    const sid = filterState.setor === '' ? -1 : DATA.setores.indexOf(filterState.setor);
    const uid = filterState.uf === '' ? -1 : DATA.ufs.indexOf(filterState.uf);
    const bid = filterState.bandeira === '' ? -1 : DATA.bandeiras.indexOf(filterState.bandeira);
    if (gid !== -1 && r[AIDX.G] !== gid) return false;
    if (sid !== -1 && r[AIDX.S] !== sid) return false;
    if (uid !== -1 && r[AIDX.U] !== uid) return false;
    if (bid !== -1 && r[AIDX.B] !== bid) return false;
    return true;
  });
}

function setoresVisiveis() {
  // Conjunto de IDs de setor com vendas após filtros
  const ids = new Set();
  rowsFiltradas().forEach(r => ids.add(r[IDX.S]));
  return [...ids].sort((a, b) => DATA.setores[a].localeCompare(DATA.setores[b]));
}

// ============================================================
// CÁLCULO DE MÉTRICAS DE SETOR
// ============================================================
function metricasSetor(setorId) {
  const f = rowsFiltradas().filter(r => r[IDX.S] === setorId);
  const m = { jan:0, fev:0, mar:0, abr:0, s1:0, s2:0, s3:0, proj:0 };
  f.forEach(r => {
    m.jan += r[IDX.JAN]; m.fev += r[IDX.FEV]; m.mar += r[IDX.MAR]; m.abr += r[IDX.ABR];
    m.s1 += r[IDX.S1]; m.s2 += r[IDX.S2]; m.s3 += r[IDX.S3]; m.proj += r[IDX.PROJ];
  });
  m.ytd = m.jan + m.fev + m.mar + m.abr;
  m.varMaiAbr = m.proj - m.abr;
  m.varMaiAbrPct = m.abr > 0 ? (m.varMaiAbr / m.abr * 100) : 0;
  m.varS3S2 = m.s3 - m.s2;
  m.varS3S2Pct = m.s2 > 0 ? (m.varS3S2 / m.s2 * 100) : 0;

  // PDVs do setor (com variação ≥ threshold)
  const pdvs = pdvsFiltrados().filter(r => r[PIDX.S] === setorId);
  const thr = DATA.meta.threshold;
  m.altas = pdvs.filter(r => r[PIDX.VAR] >= thr).sort((a,b) => b[PIDX.VAR] - a[PIDX.VAR]).slice(0, 5);
  m.quedas = pdvs.filter(r => r[PIDX.VAR] <= -thr).sort((a,b) => a[PIDX.VAR] - b[PIDX.VAR]).slice(0, 5);
  m.totalAltas = pdvs.filter(r => r[PIDX.VAR] >= thr).length;
  m.totalQuedas = pdvs.filter(r => r[PIDX.VAR] <= -thr).length;
  m.pdvsTotal = DATA.pdvs_por_setor[setorId];
  return m;
}

// ============================================================
// RENDER
// ============================================================
function init() {
  // Filtros populados
  popularSelect('f-gerente', DATA.gerentes);
  popularSelect('f-setor', DATA.setores);
  popularSelect('f-uf', DATA.ufs);
  popularSelect('f-bandeira', DATA.bandeiras);
  popularSelect('f-marca', DATA.marcas);

  ['f-gerente','f-setor','f-uf','f-bandeira','f-marca'].forEach(id => {
    document.getElementById(id).addEventListener('change', e => {
      const key = id.replace('f-','');
      filterState[key] = e.target.value;
      renderTudo();
    });
  });
  document.getElementById('btn-reset').addEventListener('click', () => {
    filterState = { gerente:'', setor:'', uf:'', bandeira:'', marca:'' };
    document.querySelectorAll('.filters select').forEach(s => s.value = '');
    renderTudo();
  });
  document.getElementById('btn-export').addEventListener('click', abrirExport);

  document.getElementById('meta-info').textContent =
    `Gerado em ${DATA.meta.gerado_em} · ${DATA.meta.tipo} · ${DATA.setores.length} setores`;

  document.querySelectorAll('.thr').forEach(e => e.textContent = DATA.meta.threshold);

  renderTudo();
}

function popularSelect(id, items) {
  const sel = document.getElementById(id);
  items.forEach(v => {
    const opt = document.createElement('option');
    opt.value = v; opt.textContent = v;
    sel.appendChild(opt);
  });
}

function renderTudo() {
  renderSummary();
  renderProduto();
  renderSetores();
}

function renderSummary() {
  const setIds = setoresVisiveis();
  let ytd=0, proj=0, abr=0, s2=0, s3=0, alta=0, queda=0;
  setIds.forEach(sid => {
    const m = metricasSetor(sid);
    ytd += m.ytd; proj += m.proj; abr += m.abr;
    s2 += m.s2; s3 += m.s3;
    alta += m.totalAltas; queda += m.totalQuedas;
  });
  document.getElementById('s-setores').textContent = setIds.length;
  document.getElementById('s-ytd').textContent = fmtNum(ytd);
  document.getElementById('s-mai').textContent = fmtNum(proj);
  const dM = proj - abr;
  const dMpct = abr > 0 ? (dM / abr * 100) : 0;
  document.getElementById('s-var').innerHTML = `${fmtNum(dM)} <span style="font-size:12px;opacity:0.8;">${fmtPct(dMpct)}</span>`;
  const dS = s3 - s2;
  const dSpct = s2 > 0 ? (dS / s2 * 100) : 0;
  document.getElementById('s-s3s2').innerHTML = `${fmtNum(dS)} <span style="font-size:12px;opacity:0.8;">${fmtPct(dSpct)}</span>`;
  document.getElementById('s-pdvs').textContent = `↑${alta}  ↓${queda}`;
  document.getElementById('s-ytd-label').textContent = DATA.meta.ytd_label + ' (un.)';
}

function renderSetores() {
  const cont = document.getElementById('setor-list');
  cont.innerHTML = '';
  const setIds = setoresVisiveis();
  document.getElementById('contador-setores').textContent = `(${setIds.length} setor${setIds.length===1?'':'es'})`;

  // Destrói charts antigos
  Object.keys(chartInstances).filter(k => k.startsWith('s-')).forEach(k => {
    chartInstances[k].destroy(); delete chartInstances[k];
  });

  if (setIds.length === 0) {
    cont.innerHTML = '<div style="grid-column:1/-1;padding:40px;text-align:center;color:var(--gray);">Nenhum setor corresponde aos filtros.</div>';
    return;
  }

  setIds.forEach(sid => {
    const m = metricasSetor(sid);
    const setor = DATA.setores[sid];
    const gerente = DATA.gerentes[DATA.setor_gerente[sid]];

    const card = document.createElement('div');
    card.className = 'setor-card';
    card.dataset.setorId = sid;
    card.innerHTML = `
      <h3>${safe(setor)}</h3>
      <div class="gerente">${safe(gerente)}</div>
      <div class="kpis">
        <div class="kpi"><small>YTD</small><b>${fmtNum(m.ytd)}</b></div>
        <div class="kpi gold"><small>Mai-Proj</small><b>${fmtNum(m.proj)}</b></div>
        <div class="kpi ${m.varMaiAbr>=0?'green':'red'}"><small>Δ vs Abr</small><b>${fmtNum(m.varMaiAbr)} <small style="font-size:9px;">${fmtPct(m.varMaiAbrPct)}</small></b></div>
        <div class="kpi ${m.varS3S2>=0?'green':'red'}"><small>Δ S3 vs S2</small><b>${fmtNum(m.varS3S2)} <small style="font-size:9px;">${fmtPct(m.varS3S2Pct)}</small></b></div>
      </div>
      <canvas id="card-chart-${sid}"></canvas>
      <div style="font-size:9px;color:var(--gray);margin-top:6px;text-align:center;">↑ ${m.totalAltas} PDVs em alta · ↓ ${m.totalQuedas} PDVs em queda  ·  clique para ver detalhes</div>
    `;
    card.addEventListener('click', () => abrirModal(sid));
    cont.appendChild(card);

    // Mini-chart de tendência
    const canvas = document.getElementById(`card-chart-${sid}`);
    chartInstances[`s-${sid}`] = new Chart(canvas, {
      type: 'line',
      data: {
        labels: ['Jan','Fev','Mar','Abr','Mai-P'],
        datasets: [{
          data: [m.jan, m.fev, m.mar, m.abr, m.proj],
          borderColor: COLORS.navy,
          backgroundColor: 'rgba(30,39,97,0.1)',
          tension: 0.3,
          pointRadius: 3,
          pointBackgroundColor: ctx => ctx.dataIndex === 4 ? COLORS.gold : COLORS.navy,
          segment: { borderDash: ctx => ctx.p1DataIndex === 4 ? [4,3] : undefined,
                     borderColor: ctx => ctx.p1DataIndex === 4 ? COLORS.gold : COLORS.navy }
        }]
      },
      options: {
        responsive: true, maintainAspectRatio: false,
        plugins: { legend: { display: false }, tooltip: { enabled: true } },
        scales: {
          x: { display: true, ticks: { font: { size: 9 } }, grid: { display: false } },
          y: { display: false }
        }
      }
    });
  });
}

function renderProduto() {
  // Destrói charts antigos
  ['marca-ytd','marca-var'].forEach(k => {
    if (chartInstances[k]) { chartInstances[k].destroy(); delete chartInstances[k]; }
  });

  // Agrega marca
  const marcaMap = {};
  marcaRowsFiltradas().forEach(r => {
    const m = DATA.marcas[r[MIDX.M]];
    if (!marcaMap[m]) marcaMap[m] = { ytd:0, ytd_ant:0 };
    marcaMap[m].ytd += r[MIDX.YTD];
    marcaMap[m].ytd_ant += r[MIDX.YTD_ANT];
  });
  let marcas = Object.entries(marcaMap)
    .map(([nome, v]) => ({ nome, ...v, var: v.ytd - v.ytd_ant,
                           pct: v.ytd_ant > 0 ? (v.ytd - v.ytd_ant)/v.ytd_ant*100 : null }))
    .filter(m => m.ytd > 0 || m.ytd_ant > 0)
    .sort((a, b) => b.ytd - a.ytd);

  document.getElementById('prod-titulo-1').textContent = `Ranking por Marca — ${DATA.meta.ytd_label}`;
  document.getElementById('prod-titulo-2').textContent = `Variação por Marca — ${DATA.meta.ytd_label} vs ${DATA.meta.ytd_ant_label}`;

  if (marcas.length === 0) {
    document.getElementById('apres-tbody').innerHTML = '<tr><td colspan="5" class="empty">Sem dados para os filtros aplicados.</td></tr>';
    return;
  }

  // Chart 1: ranking YTD
  chartInstances['marca-ytd'] = new Chart(document.getElementById('chart-marca-ytd'), {
    type: 'bar',
    data: {
      labels: marcas.slice(0, 10).map(m => m.nome),
      datasets: [{ data: marcas.slice(0, 10).map(m => m.ytd), backgroundColor: COLORS.navy }]
    },
    options: {
      indexAxis: 'y', responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: { callbacks: { label: ctx => fmtNum(ctx.raw) + ' un.' } },
        datalabels: { display: false }
      },
      scales: {
        x: { display: false },
        y: { ticks: { font: { size: 10 } } }
      }
    },
    plugins: [{
      afterDatasetsDraw(chart) {
        const ctx = chart.ctx;
        chart.data.datasets[0].data.forEach((v, i) => {
          const bar = chart.getDatasetMeta(0).data[i];
          ctx.fillStyle = COLORS.navy; ctx.font = 'bold 10px sans-serif';
          ctx.textAlign = 'left'; ctx.textBaseline = 'middle';
          ctx.fillText(fmtNum(v), bar.x + 4, bar.y);
        });
      }
    }]
  });

  // Chart 2: variação
  chartInstances['marca-var'] = new Chart(document.getElementById('chart-marca-var'), {
    type: 'bar',
    data: {
      labels: marcas.slice(0, 10).map(m => m.nome),
      datasets: [{
        data: marcas.slice(0, 10).map(m => m.var),
        backgroundColor: marcas.slice(0, 10).map(m => m.var >= 0 ? COLORS.green : COLORS.red)
      }]
    },
    options: {
      indexAxis: 'y', responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: { callbacks: { label: ctx => fmtNum(ctx.raw) + ' un. (' + fmtPct(marcas[ctx.dataIndex].pct) + ')' } }
      },
      scales: {
        x: { display: false },
        y: { ticks: { font: { size: 10 } } }
      }
    },
    plugins: [{
      afterDatasetsDraw(chart) {
        const ctx = chart.ctx;
        chart.data.datasets[0].data.forEach((v, i) => {
          const bar = chart.getDatasetMeta(0).data[i];
          ctx.fillStyle = v >= 0 ? COLORS.green : COLORS.red;
          ctx.font = 'bold 9px sans-serif'; ctx.textBaseline = 'middle';
          if (v >= 0) {
            ctx.textAlign = 'left';
            ctx.fillText(fmtNum(v) + ' (' + fmtPct(marcas[i].pct) + ')', bar.x + 4, bar.y);
          } else {
            ctx.textAlign = 'right';
            ctx.fillText(fmtNum(v) + ' (' + fmtPct(marcas[i].pct) + ')', bar.x - 4, bar.y);
          }
        });
      }
    }]
  });

  // Tabela apresentação
  const apresMap = {};
  apresRowsFiltradas().forEach(r => {
    const a = DATA.apresentacoes[r[AIDX.A]];
    if (!apresMap[a]) apresMap[a] = { ytd:0, ytd_ant:0 };
    apresMap[a].ytd += r[AIDX.YTD];
    apresMap[a].ytd_ant += r[AIDX.YTD_ANT];
  });
  const apres = Object.entries(apresMap)
    .map(([nome, v]) => ({ nome, ...v, var: v.ytd - v.ytd_ant,
                           pct: v.ytd_ant > 0 ? (v.ytd - v.ytd_ant)/v.ytd_ant*100 : null }))
    .filter(a => a.ytd > 0 || a.ytd_ant > 0)
    .sort((a, b) => b.ytd - a.ytd);

  document.getElementById('apres-h1').textContent = DATA.meta.ytd_ant_label;
  document.getElementById('apres-h2').textContent = DATA.meta.ytd_label;

  const tbody = document.getElementById('apres-tbody');
  tbody.innerHTML = apres.map(a => `
    <tr>
      <td>${safe(a.nome)}</td>
      <td class="num">${fmtNum(a.ytd_ant)}</td>
      <td class="num"><b>${fmtNum(a.ytd)}</b></td>
      <td class="num ${a.var>=0?'green':'red'}">${fmtNum(a.var)}</td>
      <td class="num ${a.var>=0?'green':'red'}">${fmtPct(a.pct)}</td>
    </tr>
  `).join('');
}

// ============================================================
// MODAL
// ============================================================
function abrirModal(sid) {
  const m = metricasSetor(sid);
  const setor = DATA.setores[sid];
  const gerente = DATA.gerentes[DATA.setor_gerente[sid]];
  document.getElementById('m-titulo').textContent = 'Setor: ' + setor;
  document.getElementById('m-gerente').textContent = 'Gerente: ' + gerente;

  const kpiHtml = `
    <div class="kpi-card-lg">
      <small>${DATA.meta.ytd_label} — ${DATA.meta.tipo}</small>
      <div class="val">${fmtNum(m.ytd)} un.</div>
      <div class="sub">PDVs ativos no setor: ${m.pdvsTotal}</div>
    </div>
    <div class="kpi-card-lg gold">
      <small>Mai-Proj (Maio Projetado)</small>
      <div class="val">${fmtNum(m.proj)} un.</div>
      <div class="sub ${m.varMaiAbr>=0?'green':'red'}">${fmtNum(m.varMaiAbr)} un. vs Abr (${fmtPct(m.varMaiAbrPct)})</div>
    </div>
    <div class="kpi-card-lg ${m.varS3S2>=0?'green':'red'}">
      <small>Mai-S3 vs Mai-S2</small>
      <div class="val">${fmtNum(m.varS3S2)} un.</div>
      <div class="sub ${m.varS3S2>=0?'green':'red'}">${fmtPct(m.varS3S2Pct)}</div>
    </div>
    <div class="kpi-card-lg">
      <small>PDVs em DESTAQUE (≥±${DATA.meta.threshold})</small>
      <div class="val">↑ ${m.totalAltas}   ↓ ${m.totalQuedas}</div>
      <div class="sub">Threshold: ±${DATA.meta.threshold} unid.</div>
    </div>
  `;
  document.getElementById('m-kpis').innerHTML = kpiHtml;

  // Destrói modal charts antigos
  ['m-evo','m-sem'].forEach(k => {
    if (chartInstances[k]) { chartInstances[k].destroy(); delete chartInstances[k]; }
  });

  // Chart evolução
  chartInstances['m-evo'] = new Chart(document.getElementById('m-chart-evo'), {
    type: 'line',
    data: {
      labels: ['Jan','Fev','Mar','Abr','Mai-Proj'],
      datasets: [{
        data: [m.jan, m.fev, m.mar, m.abr, m.proj],
        borderColor: COLORS.navy, backgroundColor: 'rgba(30,39,97,0.1)',
        tension: 0.25, pointRadius: 6, pointHoverRadius: 8, borderWidth: 3,
        pointBackgroundColor: ctx => ctx.dataIndex === 4 ? COLORS.gold : COLORS.navy,
        segment: {
          borderDash: ctx => ctx.p1DataIndex === 4 ? [6,4] : undefined,
          borderColor: ctx => ctx.p1DataIndex === 4 ? COLORS.gold : COLORS.navy
        }
      }]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: { callbacks: { label: ctx => fmtNum(ctx.raw) + ' un.' } }
      },
      scales: {
        x: { grid: { display: false } },
        y: { ticks: { callback: v => fmtNum(v) } }
      }
    },
    plugins: [{
      afterDatasetsDraw(chart) {
        const ctx = chart.ctx;
        chart.data.datasets[0].data.forEach((v, i) => {
          const pt = chart.getDatasetMeta(0).data[i];
          ctx.fillStyle = i === 4 ? COLORS.gold : COLORS.navy;
          ctx.font = 'bold 11px sans-serif'; ctx.textAlign = 'center';
          ctx.fillText(fmtNum(v), pt.x, pt.y - 12);
        });
      }
    }]
  });

  // Chart semanal
  chartInstances['m-sem'] = new Chart(document.getElementById('m-chart-sem'), {
    type: 'bar',
    data: {
      labels: ['S1','S2','S3'],
      datasets: [{
        data: [m.s1, m.s2, m.s3],
        backgroundColor: [COLORS.grayLight, COLORS.ice, COLORS.navy],
        borderRadius: 4
      }]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: { callbacks: { label: ctx => fmtNum(ctx.raw) + ' un.' } }
      },
      scales: {
        x: { grid: { display: false } },
        y: { ticks: { callback: v => fmtNum(v) } }
      }
    },
    plugins: [{
      afterDatasetsDraw(chart) {
        const ctx = chart.ctx;
        chart.data.datasets[0].data.forEach((v, i) => {
          const bar = chart.getDatasetMeta(0).data[i];
          ctx.fillStyle = COLORS.navy; ctx.font = 'bold 12px sans-serif';
          ctx.textAlign = 'center';
          ctx.fillText(fmtNum(v), bar.x, bar.y - 6);
        });
      }
    }]
  });

  // Tops
  document.getElementById('m-tbl-altas').innerHTML = m.altas.length === 0
    ? `<tr><td colspan="4" class="empty">Nenhum PDV com variação ≥ +${DATA.meta.threshold} unid.</td></tr>`
    : m.altas.map(p => `
      <tr>
        <td>${safe(p[PIDX.NAME])} · ${safe(p[PIDX.CITY])}/${safe(p[PIDX.UF])}</td>
        <td class="num">${fmtNum(p[PIDX.ABR])}</td>
        <td class="num"><b>${fmtNum(p[PIDX.PROJ])}</b></td>
        <td class="num green">${fmtNum(p[PIDX.VAR])}</td>
      </tr>`).join('');

  document.getElementById('m-tbl-quedas').innerHTML = m.quedas.length === 0
    ? `<tr><td colspan="4" class="empty">Nenhum PDV com variação ≤ −${DATA.meta.threshold} unid.</td></tr>`
    : m.quedas.map(p => `
      <tr>
        <td>${safe(p[PIDX.NAME])} · ${safe(p[PIDX.CITY])}/${safe(p[PIDX.UF])}</td>
        <td class="num">${fmtNum(p[PIDX.ABR])}</td>
        <td class="num"><b>${fmtNum(p[PIDX.PROJ])}</b></td>
        <td class="num red">${fmtNum(p[PIDX.VAR])}</td>
      </tr>`).join('');

  document.getElementById('modal').classList.add('open');
}

function fecharModal() {
  document.getElementById('modal').classList.remove('open');
}
document.getElementById('modal').addEventListener('click', e => {
  if (e.target.id === 'modal') fecharModal();
});

// ============================================================
// EXPORT PowerPoint
// ============================================================
function abrirExport() {
  document.getElementById('exp-count').textContent = setoresVisiveis().length;
  document.getElementById('export-modal').classList.add('open');
}
function fecharExport() {
  document.getElementById('export-modal').classList.remove('open');
}

async function exportarPPT() {
  fecharExport();
  const scope = document.querySelector('input[name="export-scope"]:checked').value;
  const setIds = scope === 'all'
    ? DATA.setores.map((_, i) => i).filter(i => setoresTodos().has(i))
    : setoresVisiveis();

  if (setIds.length === 0) { alert('Nenhum setor selecionado.'); return; }

  const progress = document.getElementById('progress');
  const progressPct = document.getElementById('progress-pct');
  progress.classList.add('show');
  progressPct.textContent = '0%';

  // Backup do filter state e limpa para gerar "all" se necessário
  const stateBackup = { ...filterState };
  if (scope === 'all') filterState = { gerente:'', setor:'', uf:'', bandeira:'', marca:'' };

  await new Promise(r => setTimeout(r, 50));

  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_WIDE';
  pptx.defineLayout({ name: 'CUSTOM', width: 13.333, height: 7.5 });
  pptx.layout = 'CUSTOM';

  // Capa
  slideCapa(pptx, setIds.length);

  // Produto
  slideProduto(pptx);
  slideApresentacao(pptx);

  // Setores
  for (let i = 0; i < setIds.length; i++) {
    slideSetor(pptx, setIds[i]);
    const pct = Math.round(((i+1) / setIds.length) * 100);
    progressPct.textContent = pct + '%';
    if (i % 5 === 0) await new Promise(r => setTimeout(r, 10));
  }

  filterState = stateBackup;

  const stamp = new Date().toISOString().slice(0,10).replace(/-/g,'');
  await pptx.writeFile({ fileName: `MDTRS_Analise_Setor_${stamp}.pptx` });

  progress.classList.remove('show');
}

function setoresTodos() {
  const ids = new Set();
  DATA.rows.forEach(r => ids.add(r[IDX.S]));
  return ids;
}

function slideCapa(pptx, nSetores) {
  const s = pptx.addSlide();
  s.background = { color: '1E2761' };
  s.addShape('rect', { x: 13.15, y: 0, w: 0.18, h: 7.5, fill: { color: 'C9A227' } });
  s.addText('ALCON | COMMERCIAL INTELLIGENCE', {
    x: 0.8, y: 1.0, w: 11.5, h: 0.5, fontSize: 14, bold: true, color: 'CADCFC', fontFace: 'Calibri'
  });
  s.addText('Análise MDTR por Setor', {
    x: 0.8, y: 1.6, w: 11.5, h: 1.5, fontSize: 44, bold: true, color: 'FFFFFF', fontFace: 'Georgia'
  });
  const desc = filtrosDescritivos();
  s.addText(`Período: ${DATA.meta.ytd_label} → ${DATA.meta.mes_corrente}-Proj  ·  ${nSetores} setor${nSetores===1?'':'es'}`, {
    x: 0.8, y: 3.2, w: 11.5, h: 0.4, fontSize: 16, color: 'CADCFC', fontFace: 'Calibri'
  });
  if (desc) {
    s.addText(`Filtros: ${desc}`, {
      x: 0.8, y: 3.6, w: 11.5, h: 0.4, fontSize: 12, color: 'CADCFC', italic: true, fontFace: 'Calibri'
    });
  }
  s.addText('Performance mensal · Top PDVs em alta e queda · Mai-S3 vs Mai-S2 · Visão de Produto', {
    x: 0.8, y: 4.2, w: 11.5, h: 1.5, fontSize: 14, color: 'FFFFFF', fontFace: 'Calibri'
  });
  s.addText(`Gerado em ${new Date().toLocaleString('pt-BR')}  ·  Fonte: MDTRS_FV_TOTAL`, {
    x: 0.8, y: 6.6, w: 11.5, h: 0.4, fontSize: 10, color: 'CADCFC', fontFace: 'Calibri'
  });
}

function filtrosDescritivos() {
  const partes = [];
  if (filterState.gerente) partes.push(`Gerente=${filterState.gerente}`);
  if (filterState.setor) partes.push(`Setor=${filterState.setor}`);
  if (filterState.uf) partes.push(`UF=${filterState.uf}`);
  if (filterState.bandeira) partes.push(`Bandeira=${filterState.bandeira}`);
  if (filterState.marca) partes.push(`Marca=${filterState.marca}`);
  return partes.join(' · ');
}

function _headerPPT(s, titulo, sub) {
  s.background = { color: 'FFFFFF' };
  s.addShape('rect', { x: 0, y: 0, w: 13.333, h: 0.75, fill: { color: '1E2761' } });
  s.addShape('rect', { x: 0, y: 0.75, w: 13.333, h: 0.05, fill: { color: 'C9A227' } });
  s.addText(titulo, { x: 0.4, y: 0.12, w: 9, h: 0.45,
    fontSize: 20, bold: true, color: 'FFFFFF', fontFace: 'Georgia' });
  if (sub) s.addText(sub, { x: 0.4, y: 0.52, w: 9, h: 0.22,
    fontSize: 10, color: 'CADCFC', fontFace: 'Calibri' });
}

function _kpiPPT(s, x, y, w, h, label, value, sub, accent, valColor) {
  s.addShape('roundRect', { x, y, w, h, fill: { color: 'FFFFFF' },
    line: { color: 'E8E8E8', width: 0.75 }, rectRadius: 0.05 });
  s.addShape('rect', { x, y, w: 0.06, h, fill: { color: accent } });
  s.addText(label.toUpperCase(), { x: x + 0.18, y: y + 0.08, w: w - 0.2, h: 0.25,
    fontSize: 8.5, bold: true, color: '595959', fontFace: 'Calibri' });
  // Ajusta tamanho do valor automaticamente baseado no comprimento
  const valSize = String(value).length > 12 ? 16 : (String(value).length > 9 ? 19 : 22);
  s.addText(value, { x: x + 0.18, y: y + 0.32, w: w - 0.2, h: 0.55,
    fontSize: valSize, bold: true, color: valColor || '1E2761', fontFace: 'Calibri',
    valign: 'middle' });
  if (sub) {
    const subColor = sub.startsWith('+') ? '2E7D32' : (sub.startsWith('-') && !sub.startsWith('—') ? 'C62828' : '595959');
    s.addText(sub, { x: x + 0.18, y: y + h - 0.32, w: w - 0.2, h: 0.25,
      fontSize: 9.5, bold: true, color: subColor, fontFace: 'Calibri' });
  }
}

function slideSetor(pptx, sid) {
  const m = metricasSetor(sid);
  const setor = DATA.setores[sid];
  const gerente = DATA.gerentes[DATA.setor_gerente[sid]];
  const s = pptx.addSlide();
  _headerPPT(s, `Setor: ${setor}`, `Gerente: ${gerente}`);
  s.addText('ANÁLISE MENSAL · MAI vs ABR · S3 vs S2', {
    x: 9.5, y: 0.2, w: 3.5, h: 0.4, fontSize: 9, bold: true, color: 'CADCFC', align: 'right', fontFace: 'Calibri'
  });

  // KPIs
  const kpiW = 3.05, kpiH = 1.1, gap = 0.13, left0 = 0.4, kpiY = 1.0;
  _kpiPPT(s, left0, kpiY, kpiW, kpiH,
    `${DATA.meta.ytd_label} — ${DATA.meta.tipo}`,
    fmtNum(m.ytd) + ' un.',
    `PDVs ativos: ${m.pdvsTotal}`, '1E2761');
  _kpiPPT(s, left0 + (kpiW + gap), kpiY, kpiW, kpiH,
    'MAI-PROJ (Maio Projetado)',
    fmtNum(m.proj) + ' un.',
    `${fmtNum(m.varMaiAbr)} un. vs Abr (${fmtPct(m.varMaiAbrPct)})`,
    'C9A227', 'C9A227');
  _kpiPPT(s, left0 + 2*(kpiW + gap), kpiY, kpiW, kpiH,
    'MAI-S3 vs MAI-S2',
    fmtNum(m.varS3S2) + ' un.',
    fmtPct(m.varS3S2Pct),
    m.varS3S2 >= 0 ? '2E7D32' : 'C62828',
    m.varS3S2 >= 0 ? '2E7D32' : 'C62828');
  _kpiPPT(s, left0 + 3*(kpiW + gap), kpiY, kpiW, kpiH,
    `PDVs em DESTAQUE (≥±${DATA.meta.threshold})`,
    `↑ ${m.totalAltas}   ↓ ${m.totalQuedas}`,
    `Threshold: ±${DATA.meta.threshold} unid.`, '1E2761');

  // Chart evolução mensal (PptxGenJS native chart)
  s.addText('Evolução Mensal — Sell-Out (unid.)',
    { x: 0.4, y: 2.3, w: 7.6, h: 0.3, fontSize: 11, bold: true, color: '1E2761', fontFace: 'Calibri' });
  s.addChart(pptx.ChartType.line, [
    { name: 'Real', labels: ['Jan','Fev','Mar','Abr','Mai-Proj'],
      values: [m.jan, m.fev, m.mar, m.abr, m.proj] }
  ], {
    x: 0.4, y: 2.6, w: 7.6, h: 2.5,
    showLegend: false,
    chartColors: ['1E2761'],
    lineSize: 3,
    lineDataSymbol: 'circle',
    lineDataSymbolSize: 8,
    showValue: true,
    dataLabelFontSize: 9,
    dataLabelColor: '1E2761',
    dataLabelPosition: 't',
    catAxisLabelFontSize: 10,
    valAxisLabelFontSize: 9,
    valGridLine: { style: 'dash', color: 'E8E8E8' },
    catGridLine: { style: 'none' },
  });

  // Chart semanal
  s.addText('Desempenho Semanal — Maio',
    { x: 8.2, y: 2.3, w: 4.7, h: 0.3, fontSize: 11, bold: true, color: '1E2761', fontFace: 'Calibri' });
  s.addChart(pptx.ChartType.bar, [
    { name: 'Semana', labels: ['S1','S2','S3'], values: [m.s1, m.s2, m.s3] }
  ], {
    x: 8.2, y: 2.6, w: 4.7, h: 2.5,
    barDir: 'col', barGrouping: 'standard',
    chartColors: ['1E2761'],
    chartColorsOpacity: 80,
    showLegend: false,
    showValue: true,
    dataLabelFontSize: 10,
    dataLabelColor: '1E2761',
    dataLabelPosition: 'outEnd',
    catAxisLabelFontSize: 11,
    valAxisLabelFontSize: 9,
    valGridLine: { style: 'dash', color: 'E8E8E8' },
  });

  // Top tables
  _topTabelaPPT(s, 0.4, 5.25, 6.3, m.altas, '↑ TOP 5 PDVs em ALTA (Mai-Proj vs Abr ≥ +' + DATA.meta.threshold + ' unid)', '2E7D32');
  _topTabelaPPT(s, 6.9, 5.25, 6.0, m.quedas, '↓ TOP 5 PDVs em QUEDA (Mai-Proj vs Abr ≤ −' + DATA.meta.threshold + ' unid)', 'C62828');
}

function _topTabelaPPT(s, x, y, w, items, titulo, accentColor) {
  s.addText(titulo, { x, y, w, h: 0.28,
    fontSize: 10.5, bold: true, color: accentColor, fontFace: 'Calibri' });

  if (!items || items.length === 0) {
    s.addShape('rect', { x, y: y + 0.32, w, h: 1.85, fill: { color: 'F8F8F8' }, line: { color: 'F8F8F8' } });
    s.addText(`Nenhum PDV com variação ≥ ±${DATA.meta.threshold} unid.`, {
      x, y: y + 1.0, w, h: 0.4, fontSize: 10, color: '595959', align: 'center', italic: true, fontFace: 'Calibri'
    });
    return;
  }

  const rows = [[
    { text: 'PDV — Cidade/UF', options: { bold: true, color: 'FFFFFF', fill: { color: '1E2761' }, fontSize: 9, fontFace: 'Calibri' } },
    { text: 'Abr', options: { bold: true, color: 'FFFFFF', fill: { color: '1E2761' }, fontSize: 9, align: 'right', fontFace: 'Calibri' } },
    { text: 'Mai-Proj', options: { bold: true, color: 'FFFFFF', fill: { color: '1E2761' }, fontSize: 9, align: 'right', fontFace: 'Calibri' } },
    { text: 'Δ unid', options: { bold: true, color: 'FFFFFF', fill: { color: '1E2761' }, fontSize: 9, align: 'right', fontFace: 'Calibri' } },
  ]];
  items.forEach((p, i) => {
    const bg = i % 2 ? 'F8F8F8' : 'FFFFFF';
    rows.push([
      { text: `${p[PIDX.NAME]}  ·  ${p[PIDX.CITY]} / ${p[PIDX.UF]}`,
        options: { color: '1E2761', fill: { color: bg }, fontSize: 9, fontFace: 'Calibri' } },
      { text: fmtNum(p[PIDX.ABR]),
        options: { color: '595959', fill: { color: bg }, fontSize: 9, align: 'right', fontFace: 'Calibri' } },
      { text: fmtNum(p[PIDX.PROJ]),
        options: { color: '1E2761', fill: { color: bg }, fontSize: 9, bold: true, align: 'right', fontFace: 'Calibri' } },
      { text: fmtNum(p[PIDX.VAR]),
        options: { color: accentColor, fill: { color: bg }, fontSize: 9, bold: true, align: 'right', fontFace: 'Calibri' } },
    ]);
  });
  s.addTable(rows, {
    x, y: y + 0.32, w, colW: [w * 0.55, w * 0.15, w * 0.15, w * 0.15],
    rowH: 0.3, fontSize: 9, fontFace: 'Calibri', border: { type: 'solid', color: 'E8E8E8', pt: 0.5 }
  });
}

function slideProduto(pptx) {
  const s = pptx.addSlide();
  _headerPPT(s, 'Análise de Produto · Visão Consolidada',
             `${DATA.meta.ytd_label} vs ${DATA.meta.ytd_ant_label}`);

  // Agrega marca (igual ao HTML)
  const marcaMap = {};
  marcaRowsFiltradas().forEach(r => {
    const m = DATA.marcas[r[MIDX.M]];
    if (!marcaMap[m]) marcaMap[m] = { ytd:0, ytd_ant:0 };
    marcaMap[m].ytd += r[MIDX.YTD]; marcaMap[m].ytd_ant += r[MIDX.YTD_ANT];
  });
  const marcas = Object.entries(marcaMap)
    .map(([nome, v]) => ({ nome, ...v, var: v.ytd - v.ytd_ant,
                           pct: v.ytd_ant > 0 ? (v.ytd - v.ytd_ant)/v.ytd_ant*100 : null }))
    .filter(m => m.ytd > 0 || m.ytd_ant > 0)
    .sort((a, b) => b.ytd - a.ytd);

  const totalYtd = marcas.reduce((a, m) => a + m.ytd, 0);
  const totalAnt = marcas.reduce((a, m) => a + m.ytd_ant, 0);
  const varTotal = totalYtd - totalAnt;
  const varTotalPct = totalAnt > 0 ? (varTotal / totalAnt * 100) : 0;
  const melhor = marcas.reduce((a, b) => (b.var > (a?.var ?? -Infinity)) ? b : a, null);
  const pior = marcas.reduce((a, b) => (b.var < (a?.var ?? Infinity)) ? b : a, null);

  const kpiW = 3.05, kpiH = 1.1, gap = 0.13, left0 = 0.4, kpiY = 1.0;
  _kpiPPT(s, left0, kpiY, kpiW, kpiH, DATA.meta.ytd_label, fmtNum(totalYtd) + ' un.',
    `vs ${DATA.meta.ytd_ant_label}: ${fmtNum(totalAnt)} un.`, '1E2761');
  _kpiPPT(s, left0 + (kpiW + gap), kpiY, kpiW, kpiH,
    'Variação Total', fmtNum(varTotal) + ' un.', fmtPct(varTotalPct),
    varTotal >= 0 ? '2E7D32' : 'C62828', varTotal >= 0 ? '2E7D32' : 'C62828');
  if (melhor) _kpiPPT(s, left0 + 2*(kpiW + gap), kpiY, kpiW, kpiH,
    'Marca em MAIOR ALTA', (melhor.nome || '—').substring(0, 16),
    `${fmtNum(melhor.var)} un.`, '2E7D32', '2E7D32');
  if (pior) _kpiPPT(s, left0 + 3*(kpiW + gap), kpiY, kpiW, kpiH,
    'Marca em MAIOR QUEDA', (pior.nome || '—').substring(0, 16),
    `${fmtNum(pior.var)} un.`, 'C62828', 'C62828');

  // Chart marca YTD
  s.addText(`Ranking por MARCA — ${DATA.meta.ytd_label}`,
    { x: 0.4, y: 2.3, w: 6.2, h: 0.3, fontSize: 11, bold: true, color: '1E2761', fontFace: 'Calibri' });
  const top = marcas.slice(0, 8);
  s.addChart(pptx.ChartType.bar, [
    { name: 'YTD', labels: top.map(m => m.nome).reverse(), values: top.map(m => m.ytd).reverse() }
  ], {
    x: 0.4, y: 2.6, w: 6.2, h: 4.2,
    barDir: 'bar', chartColors: ['1E2761'],
    showLegend: false, showValue: true,
    dataLabelFontSize: 9, dataLabelColor: '1E2761', dataLabelPosition: 'outEnd',
    catAxisLabelFontSize: 9, valAxisLabelFontSize: 8,
    valGridLine: { style: 'none' },
  });

  s.addText(`Variação por MARCA — ${DATA.meta.ytd_label} vs ${DATA.meta.ytd_ant_label}`,
    { x: 6.8, y: 2.3, w: 6.2, h: 0.3, fontSize: 11, bold: true, color: '1E2761', fontFace: 'Calibri' });

  // Tabela de variação (mais legível que chart no PPT)
  const varRows = [[
    { text: 'MARCA', options: { bold: true, color: 'FFFFFF', fill: { color: '1E2761' }, fontSize: 10, fontFace: 'Calibri' } },
    { text: DATA.meta.ytd_ant_label, options: { bold: true, color: 'FFFFFF', fill: { color: '1E2761' }, fontSize: 10, align: 'right', fontFace: 'Calibri' } },
    { text: DATA.meta.ytd_label, options: { bold: true, color: 'FFFFFF', fill: { color: '1E2761' }, fontSize: 10, align: 'right', fontFace: 'Calibri' } },
    { text: 'Δ unid', options: { bold: true, color: 'FFFFFF', fill: { color: '1E2761' }, fontSize: 10, align: 'right', fontFace: 'Calibri' } },
    { text: 'Δ %', options: { bold: true, color: 'FFFFFF', fill: { color: '1E2761' }, fontSize: 10, align: 'right', fontFace: 'Calibri' } },
  ]];
  top.forEach((m, i) => {
    const bg = i % 2 ? 'F8F8F8' : 'FFFFFF';
    const color = m.var >= 0 ? '2E7D32' : 'C62828';
    varRows.push([
      { text: m.nome, options: { color: '1E2761', fill: { color: bg }, fontSize: 9.5, bold: true, fontFace: 'Calibri' } },
      { text: fmtNum(m.ytd_ant), options: { color: '595959', fill: { color: bg }, fontSize: 9.5, align: 'right', fontFace: 'Calibri' } },
      { text: fmtNum(m.ytd), options: { color: '1E2761', fill: { color: bg }, fontSize: 9.5, bold: true, align: 'right', fontFace: 'Calibri' } },
      { text: fmtNum(m.var), options: { color, fill: { color: bg }, fontSize: 9.5, bold: true, align: 'right', fontFace: 'Calibri' } },
      { text: fmtPct(m.pct), options: { color, fill: { color: bg }, fontSize: 9.5, bold: true, align: 'right', fontFace: 'Calibri' } },
    ]);
  });
  s.addTable(varRows, {
    x: 6.8, y: 2.6, w: 6.2, colW: [6.2*0.40, 6.2*0.16, 6.2*0.16, 6.2*0.14, 6.2*0.14],
    rowH: 0.42, fontFace: 'Calibri', border: { type: 'solid', color: 'E8E8E8', pt: 0.5 }
  });
}

function slideApresentacao(pptx) {
  const s = pptx.addSlide();
  _headerPPT(s, 'Análise de Produto · Por Apresentação',
             `${DATA.meta.ytd_label} vs ${DATA.meta.ytd_ant_label}`);

  const apresMap = {};
  apresRowsFiltradas().forEach(r => {
    const a = DATA.apresentacoes[r[AIDX.A]];
    if (!apresMap[a]) apresMap[a] = { ytd:0, ytd_ant:0 };
    apresMap[a].ytd += r[AIDX.YTD]; apresMap[a].ytd_ant += r[AIDX.YTD_ANT];
  });
  const apres = Object.entries(apresMap)
    .map(([nome, v]) => ({ nome, ...v, var: v.ytd - v.ytd_ant,
                           pct: v.ytd_ant > 0 ? (v.ytd - v.ytd_ant)/v.ytd_ant*100 : null }))
    .filter(a => a.ytd > 0 || a.ytd_ant > 0)
    .sort((a, b) => b.ytd - a.ytd)
    .slice(0, 15);

  const head = [
    { text: 'APRESENTAÇÃO', options: { bold: true, color: 'FFFFFF', fill: { color: '1E2761' }, fontSize: 10, fontFace: 'Calibri' } },
    { text: DATA.meta.ytd_ant_label, options: { bold: true, color: 'FFFFFF', fill: { color: '1E2761' }, fontSize: 10, align: 'right', fontFace: 'Calibri' } },
    { text: DATA.meta.ytd_label, options: { bold: true, color: 'FFFFFF', fill: { color: '1E2761' }, fontSize: 10, align: 'right', fontFace: 'Calibri' } },
    { text: 'Δ unid', options: { bold: true, color: 'FFFFFF', fill: { color: '1E2761' }, fontSize: 10, align: 'right', fontFace: 'Calibri' } },
    { text: 'Δ %', options: { bold: true, color: 'FFFFFF', fill: { color: '1E2761' }, fontSize: 10, align: 'right', fontFace: 'Calibri' } },
  ];
  const tbl = [head];
  apres.forEach((a, i) => {
    const bg = i % 2 ? 'F8F8F8' : 'FFFFFF';
    const color = a.var >= 0 ? '2E7D32' : 'C62828';
    tbl.push([
      { text: a.nome.substring(0, 55), options: { color: '1E2761', fill: { color: bg }, fontSize: 9.5, fontFace: 'Calibri' } },
      { text: fmtNum(a.ytd_ant), options: { color: '595959', fill: { color: bg }, fontSize: 9.5, align: 'right', fontFace: 'Calibri' } },
      { text: fmtNum(a.ytd), options: { color: '1E2761', fill: { color: bg }, fontSize: 9.5, bold: true, align: 'right', fontFace: 'Calibri' } },
      { text: fmtNum(a.var), options: { color, fill: { color: bg }, fontSize: 9.5, bold: true, align: 'right', fontFace: 'Calibri' } },
      { text: fmtPct(a.pct), options: { color, fill: { color: bg }, fontSize: 9.5, bold: true, align: 'right', fontFace: 'Calibri' } },
    ]);
  });
  const w = 12.5;
  s.addTable(tbl, {
    x: 0.4, y: 1.0, w, colW: [w*0.46, w*0.135, w*0.135, w*0.135, w*0.135],
    rowH: 0.3, fontFace: 'Calibri', border: { type: 'solid', color: 'E8E8E8', pt: 0.5 }
  });
}

// ============================================================
// BOOT
// ============================================================
init();
</script>
</body>
</html>
"""


def main():
    inp = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_INPUT
    out = sys.argv[2] if len(sys.argv) > 2 else DEFAULT_OUTPUT
    df = load_data(inp)
    payload = build_payload(df)
    build_html(payload, out)
    print("✓ Pronto. Abra o HTML em qualquer navegador moderno.")


if __name__ == "__main__":
    main()
