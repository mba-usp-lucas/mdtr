"""
Gerador de Dashboard HTML — Análise MDTR (Alcon)
=================================================
Lê o XLSX MDTRS e gera um arquivo HTML standalone com:
  - Filtros interativos (Gerente, Setor, UF, Bandeira, Marca)
  - Duas visões: por GERENTE (default, ~7 cards) e por SETOR (76 cards)
  - Slide de produto (Marca + Apresentação)
  - Botão de exportação para PowerPoint (PptxGenJS embutido — funciona offline)

USO:
    python mdtrs_gerar_html.py [caminho_excel.xlsx] [caminho_saida.html]

REUTILIZAÇÃO MENSAL:
    Edite as constantes em CONFIG (MES_CORRENTE, MES_ANTERIOR, YTD_COLS, TIPOS_INFO)
    e rode novamente. O HTML resultante é único arquivo standalone.
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

# Tipos de informação a INCLUIR na análise.
# Use uma LISTA para somar/agregar múltiplos tipos.
# Valores comuns no XLSX: "SO" (sell-out), "SI_NP" (sell-in), "N/I" (sem classificação).
# Exemplos:
#   TIPOS_INFO = ["SO"]                  -> apenas sell-out
#   TIPOS_INFO = ["SI_NP"]               -> apenas sell-in
#   TIPOS_INFO = ["SO", "SI_NP", "N/I"]  -> agrega todos (cuidado: SO + SI conta canal duas vezes)
TIPOS_INFO = ["SO", "SI_NP", "N/I"]

EXCLUIR_NAO_VISITADO = True
THRESHOLD_VARIACAO = 20             # destaque de PDV (variação ≥ X unid., Mai-Proj vs Abr)
THRESHOLD_PDV_PAYLOAD = 10          # PDVs incluídos no payload (margem p/ filtros)

TOP_BANDEIRAS = 30                  # bandeiras mantidas individualmente; resto vira "OUTRAS"
TOP_SETORES_POR_GERENTE = 5         # quantos setores destacar em alta/ofensores no slide gerente

# Filtro de família de produtos no Top 5 PDVs (alta/queda).
# Top 5 é calculado SOMENTE para marcas que casam com TOP5_MARCAS_INCLUI
# E NÃO casam com TOP5_MARCAS_EXCLUI. Use [] para incluir todas.
# Exemplo: focar em produtos SYSTANE (drops oftálmicas) sem incluir LID WIPES (toalha de limpeza).
TOP5_MARCAS_INCLUI = ["SYSTANE"]            # prefixos a incluir (case-insensitive)
TOP5_MARCAS_EXCLUI = ["SYSTANE LID WIPES"]  # prefixos a excluir
TOP5_FAMILIA_LABEL = "Família SYSTANE (excl. LID WIPES)"  # mostrado no aviso


# ============================================================
# DETECÇÃO DINÂMICA DE PERÍODO
# ============================================================
# Mapa mês -> nome curto pra rótulos ("Jan/26", "Mai-S3", etc.)
_NOMES_MES = ["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"]
def _label_mes(yyyymm):
    """'202605' -> 'Mai/26'"""
    s = str(yyyymm)
    return f"{_NOMES_MES[int(s[4:6])-1]}/{s[2:4]}"
def _label_mes_curto(yyyymm):
    """'202605' -> 'Mai'"""
    return _NOMES_MES[int(str(yyyymm)[4:6])-1]


def detectar_estrutura_periodo(df):
    """Examina as colunas do XLSX e devolve metadados de período.

    - mes_corrente: mês com coluna `-Proj` (ex: '202605').
    - mes_anterior: último mês fechado (ex: '202604') antes do corrente.
    - semanas_corrente: lista ordenada de semanas existentes para o mês corrente,
      ex: [1,2,3] hoje. Pode crescer para [1,2,3,4] ou [1,2,3,4,5] em outras coletas.
    - ultima_semana / penultima_semana: maior e segunda-maior semana — usados no
      comparativo "S(N) vs S(N-1)". Se só existe uma semana, penultima_semana=None.
    - col_proj: nome exato da coluna projeção (ex: '202605-Proj').
    - col_semanas: lista dos nomes de coluna semanais ordenados ([..-S1, ..-S2, ..]).
    - meses_acum_ano: meses fechados do ano corrente ordenados (Jan→mês_anterior).
    - meses_5ant: 5 meses anteriores ao primeiro do ano corrente (para YTD ANT).
    - col_ytd / col_ytd_ant: listas para somatório YTD (inclui Mai-Proj) e YTD anterior.
    - label_*: rótulos prontos pra UI/PPT.
    """
    import re
    cols = [str(c) for c in df.columns]
    proj_pat = re.compile(r'^(\d{6})-Proj$')
    sem_pat  = re.compile(r'^(\d{6})-S(\d+)$')
    mes_pat  = re.compile(r'^\d{6}$')

    meses_proj = sorted({proj_pat.match(c).group(1) for c in cols if proj_pat.match(c)})
    if not meses_proj:
        raise ValueError("Nenhuma coluna de projeção '-Proj' encontrada no XLSX.")
    mes_corrente = meses_proj[-1]  # mais recente

    # Semanas do mês corrente
    semanas_corrente = sorted({int(sem_pat.match(c).group(2)) for c in cols
                                if sem_pat.match(c) and sem_pat.match(c).group(1) == mes_corrente})
    if not semanas_corrente:
        raise ValueError(f"Mês corrente {mes_corrente} sem colunas semanais (-S*).")

    ultima_semana = max(semanas_corrente)
    penultima_semana = sorted(semanas_corrente)[-2] if len(semanas_corrente) >= 2 else None

    col_proj = f"{mes_corrente}-Proj"
    col_semanas = [f"{mes_corrente}-S{n}" for n in semanas_corrente]
    col_ultima = f"{mes_corrente}-S{ultima_semana}"
    col_penultima = f"{mes_corrente}-S{penultima_semana}" if penultima_semana else None

    # Meses fechados do ano corrente: começam em <ano>01 e vão até o mês anterior ao corrente
    ano_corr = mes_corrente[:4]
    mes_n = int(mes_corrente[4:6])
    meses_acum_ano = [f"{ano_corr}{m:02d}" for m in range(1, mes_n)
                       if f"{ano_corr}{m:02d}" in cols]
    mes_anterior = meses_acum_ano[-1] if meses_acum_ano else None

    # 5 meses anteriores ao ano corrente (para comparação YTD ANT).
    # Pega os 5 meses imediatamente anteriores a Jan/<ano_corr>.
    todos_mensais = sorted([c for c in cols if mes_pat.match(c)])
    primeiro_ano_corr = f"{ano_corr}01"
    meses_5ant = [m for m in todos_mensais if m < primeiro_ano_corr][-5:]

    col_ytd = meses_acum_ano + [col_proj]  # Jan..AbrFechado + Mai-Proj
    col_ytd_ant = meses_5ant  # 5 meses anteriores fechados

    # Rótulos amigáveis
    if meses_acum_ano:
        label_ytd = f"Últimos {len(col_ytd)} meses ({_label_mes_curto(meses_acum_ano[0])}-{_label_mes_curto(mes_corrente)}/{mes_corrente[2:4]})"
    else:
        label_ytd = f"{_label_mes(mes_corrente)} projetado"
    if col_ytd_ant:
        label_ytd_ant = f"{len(col_ytd_ant)} meses anteriores ({_label_mes_curto(col_ytd_ant[0])}-{_label_mes_curto(col_ytd_ant[-1])}/{col_ytd_ant[-1][2:4]})"
    else:
        label_ytd_ant = "Período anterior (n/d)"
    label_corrente_curto = _label_mes_curto(mes_corrente)
    label_anterior_curto = _label_mes_curto(mes_anterior) if mes_anterior else ""

    label_comparativo_semanal = (
        f"{label_corrente_curto}-S{ultima_semana} vs {label_corrente_curto}-S{penultima_semana}"
        if penultima_semana else
        f"{label_corrente_curto}-S{ultima_semana} (única semana disponível)"
    )

    return {
        "mes_corrente": mes_corrente,
        "mes_anterior": mes_anterior,
        "semanas_corrente": semanas_corrente,
        "ultima_semana": ultima_semana,
        "penultima_semana": penultima_semana,
        "col_proj": col_proj,
        "col_semanas": col_semanas,
        "col_ultima": col_ultima,
        "col_penultima": col_penultima,
        "meses_acum_ano": meses_acum_ano,
        "meses_5ant": meses_5ant,
        "col_ytd": col_ytd,
        "col_ytd_ant": col_ytd_ant,
        "label_ytd": label_ytd,
        "label_ytd_ant": label_ytd_ant,
        "label_corrente_curto": label_corrente_curto,
        "label_anterior_curto": label_anterior_curto,
        "label_comparativo_semanal": label_comparativo_semanal,
        "label_acum_ano": f"Acumulado {_label_mes_curto(meses_acum_ano[0])}-{_label_mes_curto(mes_anterior)}/{mes_anterior[2:4]}" if (meses_acum_ano and mes_anterior) else label_ytd,
    }


# Constantes fallback (sobrescritas pela detecção quando o XLSX é lido).
MES_CORRENTE = "202605"
MES_ANTERIOR = "202604"


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
def _validar_colunas(df, cols, label):
    """Garante que as colunas existem no XLSX. Aborta com erro claro se faltar."""
    missing = [c for c in cols if c not in df.columns]
    if missing:
        disponiveis = [c for c in df.columns if str(c).startswith("20")]
        raise ValueError(
            f"\n❌ Coluna(s) {label} não encontrada(s) no XLSX: {missing}\n"
            f"   Colunas de período disponíveis: {disponiveis}\n"
            f"   Revise o bloco CONFIG no topo do script."
        )


def build_payload(df):
    """Constrói o objeto JSON que será embutido no HTML."""
    print("[2/4] Filtrando e agregando dados ...")

    # DETECTAR estrutura de período dinamicamente — funciona com qualquer combinação
    # de semanas disponíveis (S1, S2, S3, S4, S5...) e qualquer mês corrente.
    est = detectar_estrutura_periodo(df)
    print(f"      Mês corrente: {est['mes_corrente']} ({est['label_corrente_curto']})")
    print(f"      Semanas disponíveis: S{', S'.join(str(s) for s in est['semanas_corrente'])}")
    print(f"      Comparativo semanal: {est['label_comparativo_semanal']}")

    # Aliases locais (substituem as constantes hardcoded)
    MES_CORRENTE = est["mes_corrente"]
    MES_ANTERIOR = est["mes_anterior"]
    YTD_COLS = est["col_ytd"]
    YTD_ANT_COLS = est["col_ytd_ant"]
    YTD_LABEL = est["label_ytd"]
    YTD_ANT_LABEL = est["label_ytd_ant"]
    COL_PROJ = est["col_proj"]
    COL_SEMANAS = est["col_semanas"]
    COL_ULTIMA = est["col_ultima"]
    COL_PENULTIMA = est["col_penultima"]  # pode ser None
    MESES_ACUM_ANO = est["meses_acum_ano"]
    LABEL_ACUM_ANO = est["label_acum_ano"]
    LABEL_COMP_SEMANAL = est["label_comparativo_semanal"]

    # Normaliza TIPOS_INFO para lista (aceita string única também, por robustez)
    tipos = list(TIPOS_INFO) if not isinstance(TIPOS_INFO, str) else [TIPOS_INFO]
    if not tipos:
        raise ValueError("TIPOS_INFO está vazio — escolha pelo menos um tipo.")

    # Valida tipos contra o XLSX
    tipos_validos = df["TIPO_INFORMACAO"].dropna().unique().tolist()
    tipos_nao_existentes = [t for t in tipos if t not in tipos_validos]
    if tipos_nao_existentes:
        print(f"      ⚠ TIPOS_INFO ignorados (não existem no XLSX): {tipos_nao_existentes}")
        print(f"        Disponíveis: {tipos_validos}")
    tipos = [t for t in tipos if t in tipos_validos]
    if not tipos:
        raise ValueError(f"Nenhum TIPOS_INFO válido. Disponíveis no XLSX: {tipos_validos}")
    print(f"      Tipos incluídos: {tipos}")

    df = df[df["TIPO_INFORMACAO"].isin(tipos)].copy()
    if EXCLUIR_NAO_VISITADO:
        df = df[df["SETOR_NOME"] != "NÃO VISITADO"]
    print(f"      Linhas filtradas: {len(df):,}")

    # Valida colunas detectadas (sanity check)
    _validar_colunas(df, YTD_COLS, "YTD_COLS (detectado)")
    if YTD_ANT_COLS:
        _validar_colunas(df, YTD_ANT_COLS, "YTD_ANT_COLS (detectado)")
    _validar_colunas(df, [MES_ANTERIOR] + COL_SEMANAS + [COL_PROJ] if MES_ANTERIOR else COL_SEMANAS + [COL_PROJ],
                     "colunas do mês corrente")

    # Normaliza bandeira: top N e "OUTRAS"
    top_bands = df["Bandeira"].value_counts().head(TOP_BANDEIRAS).index.tolist()
    df["Bandeira_n"] = df["Bandeira"].where(df["Bandeira"].isin(top_bands), "OUTRAS")
    df["MARCA_clean"] = df["MARCA"].astype(str).str.replace(" (ALC)", "", regex=False)
    # Distribuidor ajustado (para filtro/extração por distribuidor)
    df["DIST_n"] = df["DISTRIBUIDOR AJUSTADO"].fillna("S/I").astype(str) \
        if "DISTRIBUIDOR AJUSTADO" in df.columns else "S/I"

    # Meses base para o card/gráfico de tendência (acumulado do ano fechado).
    meses_base = list(MESES_ACUM_ANO)  # Jan, Fev, Mar, Abr quando corrente=Mai
    if meses_base:
        _validar_colunas(df, meses_base, "meses_base (meses fechados do ano)")

    # period_cols: todas as colunas mensais/semanais que vamos somar nos agregados.
    period_cols = meses_base + COL_SEMANAS + [COL_PROJ]

    # -------------- AGREGADO PRINCIPAL (rows) --------------
    agg = df.groupby(["GERENTE", "SETOR_NOME", "UF", "Bandeira_n", "MARCA_clean", "DIST_n"],
                     as_index=False)[period_cols].sum()
    agg = agg[(agg[period_cols].sum(axis=1) > 0)]

    gerentes = sorted(agg["GERENTE"].unique().tolist())
    setores = sorted(agg["SETOR_NOME"].unique().tolist())
    ufs = sorted(agg["UF"].dropna().unique().tolist())
    bandeiras = sorted(agg["Bandeira_n"].unique().tolist())
    marcas = sorted(agg["MARCA_clean"].unique().tolist())
    distribuidores = sorted(agg["DIST_n"].unique().tolist())

    gid = {g: i for i, g in enumerate(gerentes)}
    sid = {s: i for i, s in enumerate(setores)}
    uid = {u: i for i, u in enumerate(ufs)}
    bid = {b: i for i, b in enumerate(bandeiras)}
    mid = {m: i for i, m in enumerate(marcas)}
    did = {d: i for i, d in enumerate(distribuidores)}

    setor_gerente = {}
    for _, r in agg.groupby(["SETOR_NOME", "GERENTE"]).size().reset_index().iterrows():
        setor_gerente[sid[r["SETOR_NOME"]]] = gid[r["GERENTE"]]

    rows = []
    for _, r in agg.iterrows():
        # Meses acumulados como array dinâmico (Jan, Fev, ... até o mês anterior).
        meses_vals = [int(round(r[m])) for m in meses_base]
        s_last_val = int(round(r[COL_ULTIMA]))
        s_prev_val = int(round(r[COL_PENULTIMA])) if COL_PENULTIMA else 0
        rows.append([
            gid[r["GERENTE"]], sid[r["SETOR_NOME"]], uid[r["UF"]],
            bid[r["Bandeira_n"]], mid[r["MARCA_clean"]],
            meses_vals,                          # índice 5: array meses acumulados
            s_last_val,                          # índice 6: última semana
            s_prev_val,                          # índice 7: penúltima semana (0 se não houver)
            int(round(r[COL_PROJ])),             # índice 8: Mai-Proj (projeção do mês corrente)
            did[r["DIST_n"]],                    # índice 9: distribuidor
        ])

    # -------------- PDVs RELEVANTES --------------
    # Top 5 PDVs (alta/queda) é calculado para uma FAMÍLIA específica de produtos
    # (configurável em TOP5_MARCAS_INCLUI / TOP5_MARCAS_EXCLUI). Isso permite focar
    # em SYSTANE (drops oftálmicas) sem misturar com LID WIPES (toalha) que tem
    # comportamento de venda completamente diferente.
    if TOP5_MARCAS_INCLUI or TOP5_MARCAS_EXCLUI:
        marca_str = df["MARCA"].astype(str).str.upper()
        incluir = (
            marca_str.apply(lambda m: any(p.upper() in m for p in TOP5_MARCAS_INCLUI))
            if TOP5_MARCAS_INCLUI else pd.Series(True, index=df.index)
        )
        excluir = (
            marca_str.apply(lambda m: any(p.upper() in m for p in TOP5_MARCAS_EXCLUI))
            if TOP5_MARCAS_EXCLUI else pd.Series(False, index=df.index)
        )
        df_top5 = df[incluir & ~excluir].copy()
        print(f"      Top 5 PDVs filtrado por família ({TOP5_FAMILIA_LABEL}): {len(df_top5):,} linhas")
    else:
        df_top5 = df

    # Agregação por CNPJ + distribuidor (cada PDV-distribuidor é uma linha;
    # permite filtrar PDVs por distribuidor na extração).
    _pdv_cols_agg = {MES_ANTERIOR: "sum", COL_PROJ: "sum", COL_ULTIMA: "sum"}
    if COL_PENULTIMA:
        _pdv_cols_agg[COL_PENULTIMA] = "sum"
    pdv_agg = df_top5.groupby(
        ["CNPJ", "PDV", "CIDADE", "UF", "Bandeira_n", "GERENTE", "SETOR_NOME", "DIST_n"],
        as_index=False
    ).agg(_pdv_cols_agg)
    pdv_agg["var"] = pdv_agg[COL_PROJ] - pdv_agg[MES_ANTERIOR]
    pdv_agg = pdv_agg[pdv_agg["var"].abs() >= THRESHOLD_PDV_PAYLOAD].copy()

    # PDV_clean: nome curto + sufixo do CNPJ (últimos 4 dígitos significativos)
    # Isso evita "DROGARIA VENANCIO" repetido visualmente quando são lojas diferentes
    # da mesma rede (CNPJs distintos).
    def _pdv_nome(row):
        nome = str(row["PDV"]).split(" - ")[0][:35]
        # Sufixo: últimos 4 dígitos do CNPJ (após remover decimais flutuantes)
        try:
            cnpj_int = int(row["CNPJ"])
            suf = str(cnpj_int)[-4:]
            return f"{nome} ({suf})"
        except Exception:
            return nome
    pdv_agg["PDV_clean"] = pdv_agg.apply(_pdv_nome, axis=1)
    pdv_agg["CIDADE_clean"] = pdv_agg["CIDADE"].astype(str).str.split(" - ").str[0].str[:25]

    # Para cada PDV relevante, identificar a MARCA que mais contribuiu para sua variação.
    # Esse contexto evita o "PDV cresceu +60 mas por causa de qual produto?".
    # Usa o MESMO filtro de família (df_top5) para coerência com o top 5 ranking.
    cnpjs_relevantes = set(pdv_agg["CNPJ"].unique())
    df_relev = df_top5[df_top5["CNPJ"].isin(cnpjs_relevantes)]
    cnpj_marca = df_relev.groupby(["CNPJ", "MARCA_clean"], as_index=False).agg(
        {MES_ANTERIOR: "sum", COL_PROJ: "sum"})
    cnpj_marca["var"] = cnpj_marca[COL_PROJ] - cnpj_marca[MES_ANTERIOR]

    # Para cada CNPJ, pegar a marca de variação mais relevante NO MESMO SENTIDO do PDV.
    # Se o PDV está em alta, escolhemos a marca que mais cresceu.
    # Se está em queda, a marca que mais caiu.
    pdv_top_marca = {}  # cnpj -> {"pos": [marca, var, ...], "neg": [...]}
    for cnpj, g in cnpj_marca.groupby("CNPJ"):
        pos = g.loc[g["var"].idxmax()]
        neg = g.loc[g["var"].idxmin()]
        pdv_top_marca[cnpj] = {
            "pos": (str(pos["MARCA_clean"]), int(round(pos["var"]))),
            "neg": (str(neg["MARCA_clean"]), int(round(neg["var"]))),
        }

    pdvs = []
    for _, r in pdv_agg.iterrows():
        if r["GERENTE"] not in gid or r["SETOR_NOME"] not in sid:
            continue
        # Escolhe a marca de maior contribuição no MESMO sentido da variação do PDV
        marca_info = pdv_top_marca.get(r["CNPJ"], {"pos": ("—", 0), "neg": ("—", 0)})
        marca_top = marca_info["pos"] if r["var"] >= 0 else marca_info["neg"]
        pdvs.append([
            gid[r["GERENTE"]], sid[r["SETOR_NOME"]],
            uid.get(r["UF"], -1),
            bid.get(r["Bandeira_n"], -1),
            r["PDV_clean"], r["CIDADE_clean"], r["UF"],
            int(round(r[MES_ANTERIOR])),
            int(round(r[COL_PROJ])),
            int(round(r[COL_PENULTIMA])) if COL_PENULTIMA else 0,
            int(round(r[COL_ULTIMA])),
            int(round(r["var"])),
            marca_top[0],            # nome da marca puxadora
            int(marca_top[1]),       # Δ unid da marca
            did.get(r["DIST_n"], -1),  # índice 14: distribuidor
        ])

    # -------------- TOP PDV POR SETOR (para visão por gerente) --------------
    # Para cada setor, identificar o PDV de maior impacto positivo e negativo.
    # Útil para o slide de gerente — quando se mostra um setor, queremos
    # ver o PDV que mais contribuiu (positivo ou negativo) dentro dele.
    # Usa o MESMO df filtrado por família (df_top5) — coerência com top 5 PDVs.
    pdv_setor_full = df_top5.groupby(
        ["CNPJ", "PDV", "CIDADE", "UF", "GERENTE", "SETOR_NOME"],
        as_index=False
    ).agg({MES_ANTERIOR: "sum", COL_PROJ: "sum"})
    pdv_setor_full["var"] = pdv_setor_full[COL_PROJ] - pdv_setor_full[MES_ANTERIOR]
    def _pdv_nome_short(row):
        nome = str(row["PDV"]).split(" - ")[0][:28]
        try:
            suf = str(int(row["CNPJ"]))[-4:]
            return f"{nome} ({suf})"
        except Exception:
            return nome
    pdv_setor_full["PDV_clean"] = pdv_setor_full.apply(_pdv_nome_short, axis=1)
    pdv_setor_full["CIDADE_clean"] = pdv_setor_full["CIDADE"].astype(str).str.split(" - ").str[0].str[:20]

    top_pdv_setor = {}  # setor_id -> {"alta": {...}, "queda": {...}}
    for setor_nome, g in pdv_setor_full.groupby("SETOR_NOME"):
        if setor_nome not in sid:
            continue
        s_idx = sid[setor_nome]
        alta = g.loc[g["var"].idxmax()] if len(g) else None
        queda = g.loc[g["var"].idxmin()] if len(g) else None
        top_pdv_setor[s_idx] = {
            "alta": [alta["PDV_clean"], alta["CIDADE_clean"], alta["UF"],
                     int(round(alta[MES_ANTERIOR])), int(round(alta[COL_PROJ])),
                     int(round(alta["var"]))] if alta is not None else None,
            "queda": [queda["PDV_clean"], queda["CIDADE_clean"], queda["UF"],
                      int(round(queda[MES_ANTERIOR])), int(round(queda[COL_PROJ])),
                      int(round(queda["var"]))] if queda is not None else None,
        }

    # -------------- TOP MARCA POR SETOR (qual produto puxa o setor) --------------
    # Para cada setor, identificar a MARCA de maior impacto (positivo/negativo).
    # Permite responder: "este setor caiu por causa de qual produto?"
    # Usa df_top5 (filtrado por família) para coerência com Top 5 PDVs.
    marca_setor_full = df_top5.groupby(
        ["GERENTE", "SETOR_NOME", "MARCA_clean"],
        as_index=False
    ).agg({MES_ANTERIOR: "sum", COL_PROJ: "sum"})
    marca_setor_full["var"] = marca_setor_full[COL_PROJ] - marca_setor_full[MES_ANTERIOR]

    top_marca_setor = {}  # setor_id -> {"alta": [marca, abr, proj, var], "queda": [...]}
    for setor_nome, g in marca_setor_full.groupby("SETOR_NOME"):
        if setor_nome not in sid:
            continue
        s_idx = sid[setor_nome]
        alta = g.loc[g["var"].idxmax()] if len(g) else None
        queda = g.loc[g["var"].idxmin()] if len(g) else None
        top_marca_setor[s_idx] = {
            "alta": [alta["MARCA_clean"],
                     int(round(alta[MES_ANTERIOR])),
                     int(round(alta[COL_PROJ])),
                     int(round(alta["var"]))] if alta is not None else None,
            "queda": [queda["MARCA_clean"],
                      int(round(queda[MES_ANTERIOR])),
                      int(round(queda[COL_PROJ])),
                      int(round(queda["var"]))] if queda is not None else None,
        }

    # -------------- MÉTRICAS POR SETOR FILTRADAS PELA FAMÍLIA --------------
    # Para destaques/ofensores do gerente, ordenamos setores pela variação
    # APENAS dos produtos da família (SYSTANE excl. LID WIPES). Coerente com top 5 PDVs.
    setor_familia = df_top5.groupby("SETOR_NOME", as_index=False).agg(
        {MES_ANTERIOR: "sum", COL_PROJ: "sum"})
    setor_familia["var"] = setor_familia[COL_PROJ] - setor_familia[MES_ANTERIOR]
    setor_top5_familia = {}  # setor_id -> [abr, proj, var]
    for _, r in setor_familia.iterrows():
        if r["SETOR_NOME"] not in sid:
            continue
        setor_top5_familia[sid[r["SETOR_NOME"]]] = [
            int(round(r[MES_ANTERIOR])),
            int(round(r[COL_PROJ])),
            int(round(r["var"])),
        ]

    # -------------- AGREGADOS POR GERENTE: UF + BANDEIRA + TOTAIS LIMPOS --------------
    # ESTRUTURA POOL do MDTR:
    #   POOL=1: 1 vendedor atende o brick, valor cheio na linha (ex: 1.0)
    #   POOL=2: 2 vendedores atendem o mesmo brick, valor DIVIDIDO entre eles (0.5 + 0.5 = 1.0)
    # Por isso a soma direta de TODAS as linhas é a verdade (não precisa deduplicar).
    # A chave (BRICK, PDV, SETOR_NOME, MARCA, APRESENTACAO) é única no XLSX.
    gerente_agregados = {}
    for gerente, g in df.groupby("GERENTE"):
        if gerente not in gid:
            continue
        g_idx = gid[gerente]

        # Por UF — soma direta de todas as linhas (POOL=2 mantém 0.5+0.5=1.0 real)
        uf_agg = g.groupby("UF", as_index=False).agg(
            {MES_ANTERIOR: "sum", COL_PROJ: "sum"})
        uf_agg["var"] = uf_agg[COL_PROJ] - uf_agg[MES_ANTERIOR]
        uf_agg = uf_agg.sort_values(COL_PROJ, ascending=False)
        uf_list = [[r["UF"],
                    int(round(r[MES_ANTERIOR])),
                    int(round(r[COL_PROJ])),
                    int(round(r["var"]))]
                   for _, r in uf_agg.iterrows() if pd.notna(r["UF"])]

        # Por Bandeira
        band_agg = g.groupby("Bandeira_n", as_index=False).agg(
            {MES_ANTERIOR: "sum", COL_PROJ: "sum"})
        band_agg["var"] = band_agg[COL_PROJ] - band_agg[MES_ANTERIOR]
        band_agg = band_agg.sort_values(COL_PROJ, ascending=False).head(8)
        band_list = [[r["Bandeira_n"],
                      int(round(r[MES_ANTERIOR])),
                      int(round(r[COL_PROJ])),
                      int(round(r["var"]))]
                     for _, r in band_agg.iterrows()]

        # Totais do gerente: array dinâmico de meses acumulados + s_last/s_prev + proj.
        # Mantém compat com JS: chaves jan/fev/mar/abr/s1/s2/s3 preenchidas quando possível
        # (apenas para o caso clássico Mai/26 — em casos novos, JS lê meses_acum_array).
        meses_array = [int(round(g[m].sum())) for m in meses_base]
        s_last_sum = int(round(g[COL_ULTIMA].sum()))
        s_prev_sum = int(round(g[COL_PENULTIMA].sum())) if COL_PENULTIMA else 0
        proj_sum   = int(round(g[COL_PROJ].sum()))
        kpi_clean = {
            "meses_acum": meses_array,
            "s_last": s_last_sum,
            "s_prev": s_prev_sum,
            "proj":   proj_sum,
            "abr":    meses_array[-1] if meses_array else 0,
            "acum_jan_abr": sum(meses_array),
            "n_pdvs": int(g["CNPJ"].nunique()),
            "n_setores": int(g["SETOR_NOME"].nunique()),
        }

        gerente_agregados[g_idx] = {
            "uf": uf_list,
            "bandeira": band_list,
            "totais": kpi_clean,
        }

    # -------------- PRODUTO (Marca + Apresentação) --------------
    df["_ytd"] = df[YTD_COLS].sum(axis=1)
    df["_ytd_ant"] = df[YTD_ANT_COLS].sum(axis=1)

    # -------------- PRODUTOS FOCO (família SYSTANE via GRUPO_MARCA) --------------
    # Usa a coluna oficial GRUPO_MARCA == "SYSTANE FAMÍLIA" (já exclui LID WIPES,
    # que está classificado em OUTROS). Ranking + variação por marca da família,
    # respeitando filtros de gerente/setor/UF/bandeira/distribuidor (via IDs).
    foco_rows = []  # [g, s, u, b, marca_id, ytd, ytd_ant, abr, proj, dist]
    has_grupo_marca = "GRUPO_MARCA" in df.columns
    if has_grupo_marca:
        df_foco = df[df["GRUPO_MARCA"] == "SYSTANE FAMÍLIA"].copy()
        g_foco = df_foco.groupby(
            ["GERENTE", "SETOR_NOME", "UF", "Bandeira_n", "MARCA_clean", "DIST_n"],
            as_index=False
        ).agg({"_ytd": "sum", "_ytd_ant": "sum",
               MES_ANTERIOR: "sum", COL_PROJ: "sum"})
        for _, r in g_foco.iterrows():
            if r["GERENTE"] not in gid or r["SETOR_NOME"] not in sid:
                continue
            if r["MARCA_clean"] not in mid:
                continue
            foco_rows.append([
                gid[r["GERENTE"]], sid[r["SETOR_NOME"]], uid.get(r["UF"], -1),
                bid.get(r["Bandeira_n"], -1), mid[r["MARCA_clean"]],
                int(round(r["_ytd"])), int(round(r["_ytd_ant"])),
                int(round(r[MES_ANTERIOR])), int(round(r[COL_PROJ])),
                did.get(r["DIST_n"], -1),
            ])

    # -------------- DESEMPENHO POR ASSOCIAÇÃO --------------
    # Agrega vendas por ASSOCIACAO (rede/associação de farmácias) com IDs de filtro.
    assoc_rows = []  # [g, s, u, b, assoc_id, ytd, ytd_ant, abr, proj, dist]
    associacoes = []
    if "ASSOCIACAO" in df.columns:
        df["_assoc"] = df["ASSOCIACAO"].fillna("N/I").astype(str)
        associacoes = sorted(df["_assoc"].unique().tolist())
        asid = {a: i for i, a in enumerate(associacoes)}
        g_assoc = df.groupby(
            ["GERENTE", "SETOR_NOME", "UF", "Bandeira_n", "_assoc", "DIST_n"],
            as_index=False
        ).agg({"_ytd": "sum", "_ytd_ant": "sum",
               MES_ANTERIOR: "sum", COL_PROJ: "sum"})
        for _, r in g_assoc.iterrows():
            if r["GERENTE"] not in gid or r["SETOR_NOME"] not in sid:
                continue
            assoc_rows.append([
                gid[r["GERENTE"]], sid[r["SETOR_NOME"]], uid.get(r["UF"], -1),
                bid.get(r["Bandeira_n"], -1), asid[r["_assoc"]],
                int(round(r["_ytd"])), int(round(r["_ytd_ant"])),
                int(round(r[MES_ANTERIOR])), int(round(r[COL_PROJ])),
                did.get(r["DIST_n"], -1),
            ])

    marca_rows = []
    g_marca = df.groupby(["GERENTE", "SETOR_NOME", "UF", "Bandeira_n", "MARCA_clean", "DIST_n"],
                          as_index=False).agg({"_ytd": "sum", "_ytd_ant": "sum"})
    for _, r in g_marca.iterrows():
        if r["GERENTE"] not in gid or r["SETOR_NOME"] not in sid:
            continue
        marca_rows.append([
            gid[r["GERENTE"]], sid[r["SETOR_NOME"]], uid.get(r["UF"], -1),
            bid.get(r["Bandeira_n"], -1), mid[r["MARCA_clean"]],
            int(round(r["_ytd"])), int(round(r["_ytd_ant"])),
            did.get(r["DIST_n"], -1),
        ])

    apresentacoes = sorted(df["APRESENTACAO"].dropna().unique().tolist())
    aid = {a: i for i, a in enumerate(apresentacoes)}
    apres_rows = []
    g_apres = df.groupby(["GERENTE", "SETOR_NOME", "UF", "Bandeira_n", "APRESENTACAO", "DIST_n"],
                          as_index=False).agg({"_ytd": "sum", "_ytd_ant": "sum"})
    for _, r in g_apres.iterrows():
        if r["GERENTE"] not in gid or r["SETOR_NOME"] not in sid:
            continue
        apres_rows.append([
            gid[r["GERENTE"]], sid[r["SETOR_NOME"]], uid.get(r["UF"], -1),
            bid.get(r["Bandeira_n"], -1), aid[r["APRESENTACAO"]],
            int(round(r["_ytd"])), int(round(r["_ytd_ant"])),
            did.get(r["DIST_n"], -1),
        ])

    pdvs_por_setor = df.groupby("SETOR_NOME")["CNPJ"].nunique().to_dict()
    pdvs_por_setor_arr = [pdvs_por_setor.get(s, 0) for s in setores]

    return {
        "meta": {
            "gerado_em": datetime.now().strftime("%d/%m/%Y %H:%M"),
            "tipos_info": tipos,  # lista (não set!)
            "tipo_label": " + ".join(tipos),
            "mes_corrente": MES_CORRENTE,
            "mes_anterior": MES_ANTERIOR,
            "ytd_label": YTD_LABEL,
            "ytd_ant_label": YTD_ANT_LABEL,
            "threshold": THRESHOLD_VARIACAO,
            "top_setores_gerente": TOP_SETORES_POR_GERENTE,
            "top5_familia_label": (TOP5_FAMILIA_LABEL
                                    if (TOP5_MARCAS_INCLUI or TOP5_MARCAS_EXCLUI)
                                    else ""),
            # Estrutura dinâmica de período (semanas variáveis por execução)
            "label_corrente_curto": est["label_corrente_curto"],
            "label_anterior_curto": est["label_anterior_curto"],
            "label_comparativo_semanal": LABEL_COMP_SEMANAL,
            "label_acum_ano": LABEL_ACUM_ANO,
            "ultima_semana": est["ultima_semana"],
            "penultima_semana": est["penultima_semana"],   # null se só uma semana
            "meses_acum_labels": [_label_mes_curto(m) for m in meses_base],  # ex: ["Jan","Fev","Mar","Abr"]
            "meses_acum_codes": meses_base,                                  # ex: ["202601",...,"202604"]
            "n_meses_acum": len(meses_base),
        },
        "gerentes": gerentes,
        "setores": setores,
        "setor_gerente": setor_gerente,
        "ufs": ufs,
        "bandeiras": bandeiras,
        "marcas": marcas,
        "apresentacoes": apresentacoes,
        "associacoes": associacoes,
        "distribuidores": distribuidores,
        "pdvs_por_setor": pdvs_por_setor_arr,
        # Schema rows: [g, s, u, b, m, meses_acum_array, s_last, s_prev, proj, dist]
        "rows": rows,
        # Schema pdvs: [g, s, u, b, pdv_name, cidade, uf_raw, abr, proj, s2, s3, var, marca_top, var_marca, dist]
        "pdvs": pdvs,
        # Schema marca_rows: [g, s, u, b, m, ytd, ytd_ant]
        "marca_rows": marca_rows,
        # Schema apres_rows: [g, s, u, b, a, ytd, ytd_ant]
        "apres_rows": apres_rows,
        # Schema foco_rows (SYSTANE FAMÍLIA): [g, s, u, b, marca_id, ytd, ytd_ant, abr, proj]
        "foco_rows": foco_rows,
        # Schema assoc_rows: [g, s, u, b, assoc_id, ytd, ytd_ant, abr, proj]
        "assoc_rows": assoc_rows,
        # Mapa setor_id -> {alta, queda} com [pdv, cidade, uf, abr, proj, var]
        "top_pdv_setor": top_pdv_setor,
        # Mapa setor_id -> {alta, queda} com [marca, abr, proj, var]
        "top_marca_setor": top_marca_setor,
        # Mapa setor_id -> [abr, proj, var] usando SÓ as marcas da família (SYSTANE)
        "setor_top5_familia": setor_top5_familia,
        # Mapa gerente_id -> {uf: [[uf, abr, proj, var]], bandeira: [[band, abr, proj, var]]}
        "gerente_agregados": gerente_agregados,
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
    # Só cria diretório se o output tiver pasta (evita erro quando é só nome de arquivo)
    out_dir = os.path.dirname(output_path)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)
    size_kb = os.path.getsize(output_path) / 1024
    abs_path = os.path.abspath(output_path)
    print(f"[4/4] Salvo em {abs_path} ({size_kb:.0f} KB)")


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

  /* Multi-select customizado (substitui select nativo) */
  .ms { position: relative; min-width: 180px; }
  .ms-btn {
    padding: 7px 28px 7px 10px;
    border: 1px solid #ccc;
    border-radius: 4px;
    background: white;
    font-size: 13px;
    color: var(--navy);
    cursor: pointer;
    min-width: 180px;
    text-align: left;
    width: 100%;
    position: relative;
    line-height: 1.2;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
  }
  .ms-btn:hover { border-color: var(--navy); }
  .ms.open .ms-btn { border-color: var(--navy); outline: 2px solid var(--navy); outline-offset: -1px; }
  .ms-btn::after {
    content: '▾';
    position: absolute;
    right: 8px;
    top: 50%;
    transform: translateY(-50%);
    color: var(--gray);
    font-size: 10px;
  }
  .ms-btn.has-selection { color: var(--navy); font-weight: bold; }
  .ms-panel {
    display: none;
    position: absolute;
    top: calc(100% + 4px);
    left: 0;
    z-index: 1000;
    background: white;
    border: 1px solid var(--navy);
    border-radius: 4px;
    box-shadow: 0 4px 16px rgba(0,0,0,0.15);
    min-width: 240px;
    max-width: 320px;
    padding: 6px 0;
  }
  .ms.open .ms-panel { display: block; }
  .ms-search {
    width: calc(100% - 16px);
    margin: 0 8px 6px 8px;
    padding: 6px 8px;
    border: 1px solid #ddd;
    border-radius: 3px;
    font-size: 12px;
    box-sizing: border-box;
  }
  .ms-actions {
    padding: 4px 10px;
    border-bottom: 1px solid #eee;
    margin-bottom: 4px;
    display: flex;
    justify-content: space-between;
    gap: 8px;
  }
  .ms-actions a {
    color: var(--navy);
    font-size: 11px;
    cursor: pointer;
    text-decoration: underline;
    background: none;
    border: none;
    padding: 0;
  }
  .ms-actions a:hover { color: var(--gold); }
  .ms-list {
    max-height: 240px;
    overflow-y: auto;
  }
  .ms-opt {
    padding: 5px 10px;
    cursor: pointer;
    font-size: 12px;
    color: var(--navy);
    display: flex;
    align-items: center;
    gap: 8px;
    user-select: none;
  }
  .ms-opt:hover { background: var(--ice); }
  .ms-opt input[type=checkbox] { margin: 0; cursor: pointer; }
  .ms-opt.hidden { display: none; }
  .ms-empty { padding: 10px; color: var(--gray); font-style: italic; font-size: 11px; text-align: center; }
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
  /* Faixa de aviso (filtro de família aplicado ao top 5) */
  .tops-row-header {
    display: none;
    background: #FFF7E6;
    border-left: 4px solid var(--gold);
    padding: 8px 14px;
    margin-bottom: 10px;
    font-size: 12px;
    color: var(--navy);
    border-radius: 4px;
  }
  .tops-row-header.show { display: block; }
  .tops-row-header b { color: var(--gold); }
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
  /* Sub-linha do produto puxador dentro das tabelas de PDV no modal setor */
  tr.pdv-sub-row td {
    background: #fafafa !important;
    border-bottom: 1px solid #f0f0f0;
  }
  tr.pdv-sub-row td.pdv-sub-cell {
    color: var(--gray);
    font-size: 10px;
    font-style: italic;
    padding-left: 16px;
  }

  /* Produto */
  .produto-grid {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 16px;
    margin-bottom: 20px;
  }
  #produto-section { background: white; padding: 20px; border-radius: 6px; border: 1px solid var(--gray-light); }
  .analise-section { background: white; padding: 20px; border-radius: 6px; border: 1px solid var(--gray-light); margin-bottom: 8px; }
  #assoc-table, #foco-tbl-destaques, #foco-tbl-ofensores { width: 100%; }
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

  /* View toggle */
  .view-toggle {
    display: flex;
    gap: 4px;
    margin: 22px 0 8px 0;
    border-bottom: 2px solid var(--gray-light);
  }
  .view-tab {
    background: transparent;
    border: none;
    padding: 10px 18px;
    font-size: 13px;
    font-weight: bold;
    color: var(--gray);
    cursor: pointer;
    border-bottom: 3px solid transparent;
    margin-bottom: -2px;
    transition: all 0.15s;
    border-radius: 0;
  }
  .view-tab:hover { color: var(--navy); }
  .view-tab.active {
    color: var(--navy);
    border-bottom-color: var(--gold);
  }

  /* Cards de Gerente — grid compacto, conteúdo detalhado abre no modal ao clicar */
  #gerente-list {
    display: grid;
    grid-template-columns: repeat(auto-fill, minmax(440px, 1fr));
    gap: 16px;
  }
  .gerente-card {
    background: var(--card);
    border-radius: 6px;
    padding: 16px;
    box-shadow: 0 1px 3px rgba(0,0,0,0.08);
    border: 1px solid var(--gray-light);
    cursor: pointer;
    transition: all 0.15s;
  }
  .gerente-card:hover {
    box-shadow: 0 4px 12px rgba(0,0,0,0.15);
    border-color: var(--navy);
  }
  .gerente-card h3 {
    color: var(--navy);
    font-family: Georgia, serif;
    font-size: 17px;
    margin-bottom: 2px;
  }
  .gerente-card .meta {
    font-size: 10px;
    color: var(--gray);
    margin-bottom: 10px;
    text-transform: uppercase;
    letter-spacing: 0.3px;
  }
  .gerente-card .kpis {
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    gap: 6px;
    margin-bottom: 10px;
  }
  .gerente-card .kpi {
    background: var(--bg);
    padding: 6px 8px;
    border-radius: 4px;
    border-left: 3px solid var(--navy);
  }
  .gerente-card .kpi small {
    font-size: 9px;
    color: var(--gray);
    text-transform: uppercase;
    display: block;
    letter-spacing: 0.3px;
  }
  .gerente-card .kpi b { font-size: 13px; color: var(--navy); display: block; line-height: 1.1; }
  .gerente-card .kpi.gold { border-left-color: var(--gold); }
  .gerente-card .kpi.gold b { color: var(--gold); }
  .gerente-card .kpi.green { border-left-color: var(--green); }
  .gerente-card .kpi.green b { color: var(--green); }
  .gerente-card .kpi.red { border-left-color: var(--red); }
  .gerente-card .kpi.red b { color: var(--red); }
  /* Container do mini-chart com altura FIXA para evitar loop de auto-resize */
  .gerente-card .chart-wrap {
    position: relative;
    height: 70px;
    width: 100%;
  }
  .gerente-card canvas { max-height: 70px !important; }
  .gerente-card .footer {
    font-size: 10px;
    color: var(--gray);
    margin-top: 8px;
    text-align: center;
    font-style: italic;
  }
  /* Modal de gerente — sub-linha PDV recuada */
  #mg-tbl-destaques tr.pdv-sub td,
  #mg-tbl-ofensores tr.pdv-sub td {
    font-size: 10px;
    color: var(--gray);
    font-style: italic;
    background: #fafafa;
    padding: 3px 8px 6px 8px;
  }
  #mg-tbl-destaques tr.pdv-sub td:first-child,
  #mg-tbl-ofensores tr.pdv-sub td:first-child {
    padding-left: 20px;
  }
  #mg-tbl-destaques tr.setor-main td,
  #mg-tbl-ofensores tr.setor-main td {
    font-weight: bold;
    padding-top: 8px;
  }
  #mg-tbl-destaques tr.setor-main td:first-child,
  #mg-tbl-ofensores tr.setor-main td:first-child {
    color: var(--navy);
  }
</style>
</head>
<body>
<header>
  <h1>MDTR · Análise por Gerente <small>ALCON · COMMERCIAL INTELLIGENCE</small></h1>
  <div class="meta" id="meta-info">Carregando…</div>
</header>

<div class="filters">
  <div class="filter-group">
    <label>Gerente</label>
    <div class="ms" id="ms-gerente" data-key="gerente"></div>
  </div>
  <div class="filter-group">
    <label>Setor</label>
    <div class="ms" id="ms-setor" data-key="setor"></div>
  </div>
  <div class="filter-group">
    <label>UF</label>
    <div class="ms" id="ms-uf" data-key="uf"></div>
  </div>
  <div class="filter-group">
    <label>Bandeira</label>
    <div class="ms" id="ms-bandeira" data-key="bandeira"></div>
  </div>
  <div class="filter-group">
    <label>Marca</label>
    <div class="ms" id="ms-marca" data-key="marca"></div>
  </div>
  <div class="filter-group">
    <label>Distribuidor</label>
    <div class="ms" id="ms-distribuidor" data-key="distribuidor"></div>
  </div>
  <div class="filter-actions">
    <button class="btn-secondary" id="btn-reset">Limpar filtros</button>
    <button class="btn-gold" id="btn-resumo-exec">📋 Resumo Executivo</button>
    <button class="btn-gold" id="btn-export">⬇ Exportar PowerPoint</button>
  </div>
</div>

<div class="summary" id="summary">
  <div class="item"><strong id="s-setores">—</strong><span>Setores</span></div>
  <div class="item"><strong id="s-ytd">—</strong><span id="s-ytd-label">Período principal (un.)</span></div>
  <div class="item"><strong id="s-mai">—</strong><span id="s-mai-label">Mês corrente (un.)</span></div>
  <div class="item"><strong id="s-var">—</strong><span id="s-var-label">Δ Mai vs Abr</span></div>
  <div class="item"><strong id="s-s3s2">—</strong><span id="s-s3s2-label">Δ semanal</span></div>
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

  <!-- Análise Produtos Foco — SYSTANE FAMÍLIA -->
  <div class="section-title">Produtos Foco — SYSTANE Família <span style="font-size:11px;font-weight:normal;color:var(--gray);">(exclui LID WIPES)</span></div>
  <div id="foco-section" class="analise-section">
    <div class="produto-grid">
      <div class="chart-box">
        <h4 id="foco-titulo-1">Ranking SYSTANE por Marca</h4>
        <canvas id="chart-foco-ytd"></canvas>
      </div>
      <div class="chart-box">
        <h4 id="foco-titulo-2">Variação SYSTANE por Marca</h4>
        <canvas id="chart-foco-var"></canvas>
      </div>
    </div>
    <div class="tops-row" style="margin-top:14px;">
      <div class="top-table">
        <h4 class="up">↑ SYSTANE em DESTAQUE — Maior alta (${DATA.meta.label_corrente_curto}-Proj vs ${DATA.meta.label_anterior_curto})</h4>
        <table>
          <thead><tr><th>Marca / Setor</th><th class="num">Abr</th><th class="num">${DATA.meta.label_corrente_curto}-Proj</th><th class="num">Δ unid</th></tr></thead>
          <tbody id="foco-tbl-destaques"></tbody>
        </table>
      </div>
      <div class="top-table">
        <h4 class="down">↓ SYSTANE OFENSORES — Maior queda (${DATA.meta.label_corrente_curto}-Proj vs ${DATA.meta.label_anterior_curto})</h4>
        <table>
          <thead><tr><th>Marca / Setor</th><th class="num">Abr</th><th class="num">${DATA.meta.label_corrente_curto}-Proj</th><th class="num">Δ unid</th></tr></thead>
          <tbody id="foco-tbl-ofensores"></tbody>
        </table>
      </div>
    </div>
  </div>

  <!-- Análise Desempenho por Associação -->
  <div class="section-title">Desempenho por Associação</div>
  <div id="assoc-section" class="analise-section">
    <div class="produto-grid">
      <div class="chart-box">
        <h4 id="assoc-titulo-1">Ranking por Associação — ${DATA.meta.label_corrente_curto}-Proj</h4>
        <canvas id="chart-assoc-ytd"></canvas>
      </div>
      <div class="chart-box">
        <h4 id="assoc-titulo-2">Variação por Associação (${DATA.meta.label_corrente_curto}-Proj vs ${DATA.meta.label_anterior_curto})</h4>
        <canvas id="chart-assoc-var"></canvas>
      </div>
    </div>
    <h4 style="color:var(--navy);font-size:12px;margin:14px 0 8px 0;">Detalhamento por Associação</h4>
    <table id="assoc-table">
      <thead><tr>
        <th>Associação</th>
        <th class="num">Abr</th>
        <th class="num">${DATA.meta.label_corrente_curto}-Proj</th>
        <th class="num">Δ unid</th>
        <th class="num">Δ %</th>
      </tr></thead>
      <tbody id="assoc-tbody"></tbody>
    </table>
  </div>

  <!-- Toggle Gerente / Setor -->
  <div class="view-toggle">
    <button class="view-tab active" data-view="gerente">👤 Por Gerente</button>
    <button class="view-tab" data-view="setor">📍 Por Setor</button>
  </div>

  <!-- Visão por Gerente (default) -->
  <div id="view-gerente">
    <div class="section-title">Gerentes <span id="contador-gerentes" style="font-size:12px;font-weight:normal;color:var(--gray);"></span></div>
    <div id="gerente-list"></div>
  </div>

  <!-- Visão por Setor -->
  <div id="view-setor" style="display:none;">
    <div class="section-title">Setores <span id="contador-setores" style="font-size:12px;font-weight:normal;color:var(--gray);"></span></div>
    <div id="setor-list"></div>
  </div>
</main>

<!-- Modal detalhe do GERENTE -->
<div class="modal-bg" id="modal-gerente">
  <div class="modal">
    <div class="modal-header">
      <div>
        <h2 id="mg-titulo">—</h2>
        <div class="gerente" id="mg-meta">—</div>
      </div>
      <button class="modal-close" onclick="fecharModalGerente()">×</button>
    </div>
    <div class="modal-body">
      <div class="kpi-row" id="mg-kpis"></div>

      <!-- Linha 1: Evolução mensal + Semanal -->
      <div style="display:grid;grid-template-columns:2fr 1fr;gap:14px;margin-bottom:14px;">
        <div class="chart-box">
          <h4>Evolução Mensal Agregada (unid.)</h4>
          <canvas id="mg-chart-evo" style="max-height:170px;"></canvas>
        </div>
        <div class="chart-box">
          <h4>Desempenho Semanal — ${_label_mes_completo(DATA.meta.mes_corrente)}/${DATA.meta.mes_corrente.slice(2,4)}</h4>
          <canvas id="mg-chart-sem" style="max-height:170px;"></canvas>
        </div>
      </div>

      <!-- Linha 2: UF + Bandeira -->
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:14px;">
        <div class="chart-box">
          <h4 id="mg-h4-uf">Desempenho por UF</h4>
          <canvas id="mg-chart-uf" style="max-height:200px;"></canvas>
        </div>
        <div class="chart-box">
          <h4 id="mg-h4-band">Desempenho por Bandeira</h4>
          <canvas id="mg-chart-band" style="max-height:200px;"></canvas>
        </div>
      </div>

      <!-- Linha 3: Top setores com PDV + Marca -->
      <div class="tops-row-header" id="top5-aviso-gerente"></div>
      <div class="tops-row">
        <div class="top-table">
          <h4 class="up">↑ TOP <span class="topn"></span> SETORES em DESTAQUE — Maior alta (${DATA.meta.label_corrente_curto}-Proj vs ${DATA.meta.label_anterior_curto})</h4>
          <table>
            <thead><tr>
              <th>Setor · ↳ PDV · ↳ Produto</th>
              <th class="num">Abr</th>
              <th class="num">${DATA.meta.label_corrente_curto}-Proj</th>
              <th class="num">Δ unid</th>
            </tr></thead>
            <tbody id="mg-tbl-destaques"></tbody>
          </table>
        </div>
        <div class="top-table">
          <h4 class="down">↓ TOP <span class="topn"></span> SETORES OFENSORES — Menor performance (${DATA.meta.label_corrente_curto}-Proj vs ${DATA.meta.label_anterior_curto})</h4>
          <table>
            <thead><tr>
              <th>Setor · ↳ PDV · ↳ Produto</th>
              <th class="num">Abr</th>
              <th class="num">${DATA.meta.label_corrente_curto}-Proj</th>
              <th class="num">Δ unid</th>
            </tr></thead>
            <tbody id="mg-tbl-ofensores"></tbody>
          </table>
        </div>
      </div>
    </div>
  </div>
</div>

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
      <div class="tops-row-header" id="top5-aviso-setor"></div>
      <div class="tops-row">
        <div class="top-table">
          <h4 class="up">↑ Top 5 PDVs em ALTA (${DATA.meta.label_corrente_curto}-Proj vs ${DATA.meta.label_anterior_curto} ≥ +<span class="thr"></span> unid)</h4>
          <table>
            <thead><tr><th>PDV — Cidade/UF</th><th class="num">Abr</th><th class="num">${DATA.meta.label_corrente_curto}-Proj</th><th class="num">Δ unid</th></tr></thead>
            <tbody id="m-tbl-altas"></tbody>
          </table>
        </div>
        <div class="top-table">
          <h4 class="down">↓ Top 5 PDVs em QUEDA (${DATA.meta.label_corrente_curto}-Proj vs ${DATA.meta.label_anterior_curto} ≤ −<span class="thr"></span> unid)</h4>
          <table>
            <thead><tr><th>PDV — Cidade/UF</th><th class="num">Abr</th><th class="num">${DATA.meta.label_corrente_curto}-Proj</th><th class="num">Δ unid</th></tr></thead>
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
  <p style="font-size:12px;color:var(--gray);margin-bottom:10px;font-weight:bold;">1. Visão do PPT</p>
  <label><input type="radio" name="export-view" value="gerente" checked> 👤 Por Gerente — 1 slide por gerente com top setores e PDV de maior impacto (recomendado, ~7 slides)</label>
  <label><input type="radio" name="export-view" value="setor"> 📍 Por Setor — 1 slide por setor (até ~76 slides)</label>

  <p style="font-size:12px;color:var(--gray);margin:14px 0 10px 0;font-weight:bold;">2. Escopo</p>
  <label><input type="radio" name="export-scope" value="filtered" checked> Apenas itens visíveis (com filtros aplicados) — <span id="exp-count">—</span> <span id="exp-count-label">setores</span></label>
  <label><input type="radio" name="export-scope" value="all"> Todos (ignora filtros)</label>
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
// Schema rows: [g, s, u, b, m, meses_acum_arr, s_last, s_prev, proj, dist]
// meses_acum_arr é um array dinâmico de N inteiros (1 por mês acumulado fechado).
const IDX = { G:0, S:1, U:2, B:3, M:4, MESES:5, S_LAST:6, S_PREV:7, PROJ:8, D:9 };
// Schema pdvs: [g, s, u, b, pdv_name, cidade, uf_raw, abr, proj, s_prev, s_last, var, marca, var_marca, dist]
const PIDX = { G:0, S:1, U:2, B:3, NAME:4, CITY:5, UF:6, ABR:7, PROJ:8, S_PREV:9, S_LAST:10, VAR:11, MARCA:12, VAR_MARCA:13, D:14 };
// marca_rows: [g, s, u, b, m, ytd, ytd_ant, dist]
const MIDX = { G:0, S:1, U:2, B:3, M:4, YTD:5, YTD_ANT:6, D:7 };
// apres_rows: [g, s, u, b, a, ytd, ytd_ant, dist]
const AIDX = { G:0, S:1, U:2, B:3, A:4, YTD:5, YTD_ANT:6, D:7 };
// foco_rows: [g, s, u, b, marca_id, ytd, ytd_ant, abr, proj, dist]
const FIDX = { G:0, S:1, U:2, B:3, M:4, YTD:5, YTD_ANT:6, ABR:7, PROJ:8, D:9 };
// assoc_rows: [g, s, u, b, assoc_id, ytd, ytd_ant, abr, proj, dist]
const ASIDX = { G:0, S:1, U:2, B:3, A:4, YTD:5, YTD_ANT:6, ABR:7, PROJ:8, D:9 };

let filterState = { gerente: [], setor: [], uf: [], bandeira: [], marca: [], distribuidor: [] };
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
// Helper: nome completo do mês a partir do código YYYYMM (ex: '202605' → 'Maio')
const _NOMES_MES_COMPLETO = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];
function _label_mes_completo(yyyymm) {
  return _NOMES_MES_COMPLETO[parseInt(String(yyyymm).slice(4,6)) - 1] || '';
}

// Labels dinâmicos derivados da estrutura de período detectada no Python.
// Ex: ultima=4, penultima=3, corrente_curto='Mai' → "S4 vs S3" e "Mai-S4 vs Mai-S3".
function labelDeltaSemanal(curto = false) {
  const m = DATA.meta;
  if (!m.penultima_semana) return curto ? `S${m.ultima_semana}` : `${m.label_corrente_curto}-S${m.ultima_semana}`;
  return curto
    ? `S${m.ultima_semana} vs S${m.penultima_semana}`
    : `${m.label_corrente_curto}-S${m.ultima_semana} vs ${m.label_corrente_curto}-S${m.penultima_semana}`;
}
// Labels das barras do chart semanal — usa as duas últimas semanas dinâmicas.
function labelsSemanaisChart() {
  const m = DATA.meta;
  if (!m.penultima_semana) return [`S${m.ultima_semana}`];
  return [`S${m.penultima_semana}`, `S${m.ultima_semana}`];
}

// ============================================================
// FILTRAGEM
// ============================================================
// Converte array de nomes selecionados em Set de IDs (índices no catálogo).
// Retorna null quando vazio (significa "todos passam").
function selectedIds(stateArr, catalog) {
  if (!stateArr || stateArr.length === 0) return null;
  const ids = new Set();
  stateArr.forEach(name => {
    const i = catalog.indexOf(name);
    if (i !== -1) ids.add(i);
  });
  return ids.size === 0 ? null : ids;
}

function rowMatches(row, schema = IDX) {
  const gids = selectedIds(filterState.gerente, DATA.gerentes);
  const sids = selectedIds(filterState.setor, DATA.setores);
  const uids = selectedIds(filterState.uf, DATA.ufs);
  const bids = selectedIds(filterState.bandeira, DATA.bandeiras);
  const mids = selectedIds(filterState.marca, DATA.marcas);
  const dids = selectedIds(filterState.distribuidor, DATA.distribuidores || []);

  if (gids && !gids.has(row[schema.G])) return false;
  if (sids && !sids.has(row[schema.S])) return false;
  if (uids && !uids.has(row[schema.U])) return false;
  if (bids && !bids.has(row[schema.B])) return false;
  if (mids && schema.M !== undefined && !mids.has(row[schema.M])) return false;
  if (dids && schema.D !== undefined && !dids.has(row[schema.D])) return false;
  return true;
}

function rowsFiltradas() { return DATA.rows.filter(r => rowMatches(r, IDX)); }
function pdvsFiltrados() {
  // PDVs ignoram filtro de marca (não temos marca no nível de PDV agregado)
  const gids = selectedIds(filterState.gerente, DATA.gerentes);
  const sids = selectedIds(filterState.setor, DATA.setores);
  const uids = selectedIds(filterState.uf, DATA.ufs);
  const bids = selectedIds(filterState.bandeira, DATA.bandeiras);
  const dids = selectedIds(filterState.distribuidor, DATA.distribuidores || []);
  return DATA.pdvs.filter(r => {
    if (gids && !gids.has(r[PIDX.G])) return false;
    if (sids && !sids.has(r[PIDX.S])) return false;
    if (uids && !uids.has(r[PIDX.U])) return false;
    if (bids && !bids.has(r[PIDX.B])) return false;
    if (dids && !dids.has(r[PIDX.D])) return false;
    return true;
  });
}
function marcaRowsFiltradas() { return DATA.marca_rows.filter(r => rowMatches(r, MIDX)); }
function apresRowsFiltradas() {
  // Filtros de gerente/setor/UF/bandeira/distribuidor aplicados via rowMatches.
  // Marca não se aplica (esse dataset usa apresentação como chave).
  return DATA.apres_rows.filter(r => rowMatches(r, AIDX));
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
  // Meses acumulados (array dinâmico — N meses fechados antes do corrente)
  const nMeses = DATA.meta.n_meses_acum || 0;
  const mesesAcum = new Array(nMeses).fill(0);
  let s_last = 0, s_prev = 0, proj = 0;
  f.forEach(r => {
    const arr = r[IDX.MESES] || [];
    for (let i = 0; i < arr.length; i++) mesesAcum[i] += arr[i];
    s_last += r[IDX.S_LAST]; s_prev += r[IDX.S_PREV]; proj += r[IDX.PROJ];
  });
  const ytd = mesesAcum.reduce((a, b) => a + b, 0);
  const abr = mesesAcum.length ? mesesAcum[mesesAcum.length - 1] : 0;  // último mês fechado
  const m = {
    mesesAcum, ytd, abr, s_last, s_prev, proj,
    // aliases legados para compat com código existente
    jan: mesesAcum[0] || 0, fev: mesesAcum[1] || 0, mar: mesesAcum[2] || 0,
    s1: 0,  // não temos mais S1 separado por padrão (usamos s_prev/s_last)
    s2: s_prev, s3: s_last,
  };
  m.varMaiAbr = m.proj - m.abr;
  m.varMaiAbrPct = m.abr > 0 ? (m.varMaiAbr / m.abr * 100) : 0;
  // Comparativo: última semana vs penúltima — labels dinâmicos via labelDeltaSemanal()
  m.varS3S2 = m.s_last - m.s_prev;
  m.varS3S2Pct = m.s_prev > 0 ? (m.varS3S2 / m.s_prev * 100) : 0;
  m.penultima_semana_disponivel = DATA.meta.penultima_semana != null;

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
  // Filtros multi-select
  setupMultiSelect('ms-gerente', 'gerente', DATA.gerentes);
  setupMultiSelect('ms-setor', 'setor', DATA.setores);
  setupMultiSelect('ms-uf', 'uf', DATA.ufs);
  setupMultiSelect('ms-bandeira', 'bandeira', DATA.bandeiras);
  setupMultiSelect('ms-marca', 'marca', DATA.marcas);
  setupMultiSelect('ms-distribuidor', 'distribuidor', DATA.distribuidores || []);

  // Fechar painéis ao clicar fora
  document.addEventListener('click', (e) => {
    document.querySelectorAll('.ms.open').forEach(ms => {
      if (!ms.contains(e.target)) ms.classList.remove('open');
    });
  });

  document.getElementById('btn-reset').addEventListener('click', () => {
    filterState = { gerente:[], setor:[], uf:[], bandeira:[], marca:[], distribuidor:[] };
    document.querySelectorAll('.ms').forEach(refreshMultiSelect);
    renderTudo();
  });
  document.getElementById('btn-export').addEventListener('click', abrirExport);
  document.getElementById('btn-resumo-exec').addEventListener('click', exportarResumoExecutivo);

  // Toggle entre visão por gerente / setor
  document.querySelectorAll('.view-tab').forEach(tab => {
    tab.addEventListener('click', () => {
      document.querySelectorAll('.view-tab').forEach(t => t.classList.remove('active'));
      tab.classList.add('active');
      const v = tab.dataset.view;
      document.getElementById('view-gerente').style.display = v === 'gerente' ? 'block' : 'none';
      document.getElementById('view-setor').style.display = v === 'setor' ? 'block' : 'none';
    });
  });

  document.getElementById('meta-info').textContent =
    `Gerado em ${DATA.meta.gerado_em} · ${DATA.meta.tipo_label} · ${DATA.setores.length} setores`;

  document.querySelectorAll('.thr').forEach(e => e.textContent = DATA.meta.threshold);
  document.querySelectorAll('.topn').forEach(e => e.textContent = DATA.meta.top_setores_gerente);

  // Resolve placeholders ${DATA.meta.*} restantes no HTML estático.
  // Os labels dinâmicos (Mai-Proj → Jun-Proj quando o mês mudar, S3 → S4 etc.)
  // ficam aplicados em h4/th/etc sem precisar de manipulação manual via JS.
  resolverPlaceholdersEstaticos();

  renderTudo();
}

function resolverPlaceholdersEstaticos() {
  const subs = {
    '${DATA.meta.label_corrente_curto}': DATA.meta.label_corrente_curto,
    '${DATA.meta.label_anterior_curto}': DATA.meta.label_anterior_curto || 'Anterior',
    '${DATA.meta.label_acum_ano}': DATA.meta.label_acum_ano,
    '${_label_mes_completo(DATA.meta.mes_corrente)}': _label_mes_completo(DATA.meta.mes_corrente),
    '${DATA.meta.mes_corrente.slice(2,4)}': DATA.meta.mes_corrente.slice(2,4),
  };
  // Caminha pelo body inteiro (cobre header, summary, main, modais)
  const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
  const nodes = [];
  let n;
  while ((n = walker.nextNode())) nodes.push(n);
  nodes.forEach(node => {
    let txt = node.nodeValue;
    if (txt.indexOf('${') === -1) return;
    Object.entries(subs).forEach(([k, v]) => { txt = txt.split(k).join(v); });
    if (txt !== node.nodeValue) node.nodeValue = txt;
  });
}

// ============================================================
// MULTI-SELECT
// ============================================================
function setupMultiSelect(containerId, stateKey, items) {
  const cont = document.getElementById(containerId);
  cont.dataset.key = stateKey;
  cont._items = items;

  // Botão
  const btn = document.createElement('button');
  btn.type = 'button';
  btn.className = 'ms-btn';
  cont.appendChild(btn);

  // Painel
  const panel = document.createElement('div');
  panel.className = 'ms-panel';
  panel.innerHTML = `
    <input type="text" class="ms-search" placeholder="Buscar..." />
    <div class="ms-actions">
      <a class="ms-all">Selecionar todos</a>
      <a class="ms-none">Limpar</a>
    </div>
    <div class="ms-list"></div>
  `;
  cont.appendChild(panel);

  const list = panel.querySelector('.ms-list');
  if (items.length === 0) {
    list.innerHTML = '<div class="ms-empty">Nenhuma opção disponível</div>';
  } else {
    items.forEach(v => {
      const opt = document.createElement('label');
      opt.className = 'ms-opt';
      opt.dataset.value = v;
      opt.innerHTML = `<input type="checkbox" value="${escapeHtml(v)}"> <span>${escapeHtml(v)}</span>`;
      list.appendChild(opt);
      opt.querySelector('input').addEventListener('change', (e) => {
        const arr = filterState[stateKey];
        if (e.target.checked) {
          if (!arr.includes(v)) arr.push(v);
        } else {
          filterState[stateKey] = arr.filter(x => x !== v);
        }
        refreshMultiSelect(cont);
        renderTudo();
      });
    });
  }

  // Search
  const search = panel.querySelector('.ms-search');
  search.addEventListener('input', () => {
    const q = search.value.toLowerCase();
    list.querySelectorAll('.ms-opt').forEach(opt => {
      opt.classList.toggle('hidden', q && !opt.dataset.value.toLowerCase().includes(q));
    });
  });
  // Click no input não fecha o painel
  search.addEventListener('click', e => e.stopPropagation());

  // Selecionar todos
  panel.querySelector('.ms-all').addEventListener('click', (e) => {
    e.stopPropagation();
    // Considera só os visíveis (respeita busca)
    const visiveis = [...list.querySelectorAll('.ms-opt:not(.hidden)')].map(o => o.dataset.value);
    filterState[stateKey] = [...new Set([...filterState[stateKey], ...visiveis])];
    refreshMultiSelect(cont);
    renderTudo();
  });
  panel.querySelector('.ms-none').addEventListener('click', (e) => {
    e.stopPropagation();
    filterState[stateKey] = [];
    refreshMultiSelect(cont);
    renderTudo();
  });

  // Toggle do painel
  btn.addEventListener('click', (e) => {
    e.stopPropagation();
    // Fecha outros painéis
    document.querySelectorAll('.ms.open').forEach(o => { if (o !== cont) o.classList.remove('open'); });
    cont.classList.toggle('open');
    if (cont.classList.contains('open')) {
      search.value = '';
      list.querySelectorAll('.ms-opt').forEach(o => o.classList.remove('hidden'));
      setTimeout(() => search.focus(), 50);
    }
  });

  refreshMultiSelect(cont);
}

function refreshMultiSelect(cont) {
  const stateKey = cont.dataset.key;
  const sel = filterState[stateKey] || [];
  const btn = cont.querySelector('.ms-btn');
  if (sel.length === 0) {
    btn.textContent = 'Todos';
    btn.classList.remove('has-selection');
  } else if (sel.length === 1) {
    btn.textContent = sel[0];
    btn.classList.add('has-selection');
  } else {
    btn.textContent = `${sel.length} selecionados`;
    btn.classList.add('has-selection');
  }
  // Atualizar checkboxes
  cont.querySelectorAll('.ms-opt input').forEach(cb => {
    cb.checked = sel.includes(cb.value);
  });
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, c =>
    ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}

function renderTudo() {
  renderSummary();
  renderProduto();
  renderFoco();
  renderAssoc();
  renderGerentes();
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
  // Labels dinâmicos do período corrente / comparativo semanal
  document.getElementById('s-mai-label').textContent =
    `${DATA.meta.label_corrente_curto}-Proj (un.)`;
  document.getElementById('s-var-label').textContent =
    `Δ ${DATA.meta.label_corrente_curto} vs ${DATA.meta.label_anterior_curto || 'Anterior'}`;
  document.getElementById('s-s3s2-label').textContent =
    `Δ ${DATA.meta.label_comparativo_semanal}`;
}

// ============================================================
// MÉTRICAS POR GERENTE
// ============================================================
function gerentesVisiveis() {
  const ids = new Set();
  rowsFiltradas().forEach(r => ids.add(r[IDX.G]));
  return [...ids].sort((a, b) => DATA.gerentes[a].localeCompare(DATA.gerentes[b]));
}

function setoresDoGerente(gId) {
  const ids = new Set();
  rowsFiltradas().forEach(r => { if (r[IDX.G] === gId) ids.add(r[IDX.S]); });
  return [...ids];
}

function metricasGerente(gId) {
  const nMeses = DATA.meta.n_meses_acum || 0;
  let mesesAcum = new Array(nMeses).fill(0);
  let s_last = 0, s_prev = 0, proj = 0;

  const _empty = a => !a || a.length === 0;
  const semFiltrosDimensionais =
    _empty(filterState.setor) && _empty(filterState.uf) && _empty(filterState.bandeira) &&
    _empty(filterState.marca) && _empty(filterState.distribuidor);

  const agg = DATA.gerente_agregados[gId] || {};
  if (semFiltrosDimensionais && agg.totais) {
    // Caminho RÁPIDO + EXATO: usar pré-calculado pelo Python
    const t = agg.totais;
    mesesAcum = (t.meses_acum || []).slice();
    s_last = t.s_last || 0;
    s_prev = t.s_prev || 0;
    proj = t.proj || 0;
  } else {
    // Caminho FILTRADO: somar rows
    rowsFiltradas().forEach(r => {
      if (r[IDX.G] !== gId) return;
      const arr = r[IDX.MESES] || [];
      for (let i = 0; i < arr.length; i++) mesesAcum[i] += arr[i];
      s_last += r[IDX.S_LAST]; s_prev += r[IDX.S_PREV]; proj += r[IDX.PROJ];
    });
  }
  const acumJanAbr = mesesAcum.reduce((a, b) => a + b, 0);
  const abr = mesesAcum.length ? mesesAcum[mesesAcum.length - 1] : 0;
  const m = {
    mesesAcum, acumJanAbr, abr, s_last, s_prev, proj,
    ytd: acumJanAbr,
    // aliases legados
    jan: mesesAcum[0] || 0, fev: mesesAcum[1] || 0, mar: mesesAcum[2] || 0,
    s1: 0, s2: s_prev, s3: s_last,
  };
  m.varMaiAbr = m.proj - m.abr;
  m.varMaiAbrPct = m.abr > 0 ? (m.varMaiAbr / m.abr * 100) : 0;
  m.varS3S2 = m.s_last - m.s_prev;
  m.varS3S2Pct = m.s_prev > 0 ? (m.varS3S2 / m.s_prev * 100) : 0;
  m.penultima_semana_disponivel = DATA.meta.penultima_semana != null;

  // Setores ordenados por variação da FAMÍLIA configurada (SYSTANE) — se houver filtro.
  // Caso contrário, ordena pela variação total (portfólio completo).
  // OBS: os valores Abr/${DATA.meta.label_corrente_curto}-Proj exibidos refletem a família (coerente com top 5 PDVs).
  const setIds = setoresDoGerente(gId);
  const usarFamilia = !!DATA.meta.top5_familia_label;
  const setMetricas = setIds.map(sid => {
    if (usarFamilia && DATA.setor_top5_familia && DATA.setor_top5_familia[sid]) {
      const [abr, proj, varia] = DATA.setor_top5_familia[sid];
      const pct = abr > 0 ? (varia / abr * 100) : 0;
      return { sid, setor: DATA.setores[sid], abr, proj, var: varia, varPct: pct };
    }
    const ms = metricasSetor(sid);
    return { sid, setor: DATA.setores[sid], abr: ms.abr, proj: ms.proj,
             var: ms.varMaiAbr, varPct: ms.varMaiAbrPct };
  });
  const topN = DATA.meta.top_setores_gerente;
  m.destaques = [...setMetricas].sort((a,b) => b.var - a.var).slice(0, topN);
  m.ofensores = [...setMetricas].sort((a,b) => a.var - b.var).slice(0, topN);
  m.nSetores = setIds.length;

  // UF e Bandeira: vêm dos agregados pré-calculados (já com dedupe)
  const ufSet = (filterState.uf && filterState.uf.length) ? new Set(filterState.uf) : null;
  const bandSet = (filterState.bandeira && filterState.bandeira.length) ? new Set(filterState.bandeira) : null;
  m.ufList = (agg.uf || []).filter(u => !ufSet || ufSet.has(u[0]));
  m.bandList = (agg.bandeira || []).filter(b => !bandSet || bandSet.has(b[0]));

  return m;
}

// ============================================================
// PRODUTOS FOCO — SYSTANE FAMÍLIA
// ============================================================
function renderFoco() {
  const rows = (DATA.foco_rows || []).filter(r => rowMatches(r, FIDX));
  // Agrupar por marca
  const marcaMap = {};
  rows.forEach(r => {
    const nome = DATA.marcas[r[FIDX.M]];
    if (!marcaMap[nome]) marcaMap[nome] = { ytd:0, ytd_ant:0, abr:0, proj:0 };
    marcaMap[nome].ytd += r[FIDX.YTD];
    marcaMap[nome].ytd_ant += r[FIDX.YTD_ANT];
    marcaMap[nome].abr += r[FIDX.ABR];
    marcaMap[nome].proj += r[FIDX.PROJ];
  });
  let marcas = Object.entries(marcaMap)
    .map(([nome, v]) => ({ nome, ...v, var: v.ytd - v.ytd_ant }))
    .filter(m => m.ytd > 0 || m.ytd_ant > 0)
    .sort((a, b) => b.ytd - a.ytd);

  if (marcas.length === 0) {
    ['foco-tbl-destaques','foco-tbl-ofensores'].forEach(id =>
      document.getElementById(id).innerHTML = '<tr><td colspan="4" class="empty">Sem dados.</td></tr>');
    return;
  }

  // Chart ranking YTD
  _barChartH('chart-foco-ytd', 'foco-ytd',
    marcas.map(m => m.nome), marcas.map(m => m.ytd), null);
  // Chart variação
  _barChartH('chart-foco-var', 'foco-var',
    marcas.map(m => m.nome), marcas.map(m => m.var),
    marcas.map(m => m.var >= 0 ? COLORS.green : COLORS.red));

  // Top destaques/ofensores por (marca × setor) usando Abr→${DATA.meta.label_corrente_curto}-Proj
  const porMarcaSetor = rows.map(r => ({
    marca: DATA.marcas[r[FIDX.M]],
    setor: DATA.setores[r[FIDX.S]],
    abr: r[FIDX.ABR], proj: r[FIDX.PROJ], var: r[FIDX.PROJ] - r[FIDX.ABR]
  })).filter(x => x.abr !== 0 || x.proj !== 0);

  const dest = [...porMarcaSetor].sort((a,b) => b.var - a.var).slice(0, 8);
  const ofen = [...porMarcaSetor].sort((a,b) => a.var - b.var).slice(0, 8);

  const fmtRow = (x, color) => `
    <tr>
      <td><b>${safe(x.marca)}</b><br><small style="color:var(--gray);">${safe(x.setor)}</small></td>
      <td class="num">${fmtNum(x.abr)}</td>
      <td class="num"><b>${fmtNum(x.proj)}</b></td>
      <td class="num ${color}">${fmtNum(x.var)}</td>
    </tr>`;
  document.getElementById('foco-tbl-destaques').innerHTML =
    dest.length ? dest.map(x => fmtRow(x, 'green')).join('') : '<tr><td colspan="4" class="empty">Sem dados.</td></tr>';
  document.getElementById('foco-tbl-ofensores').innerHTML =
    ofen.length ? ofen.map(x => fmtRow(x, 'red')).join('') : '<tr><td colspan="4" class="empty">Sem dados.</td></tr>';
}

// ============================================================
// DESEMPENHO POR ASSOCIAÇÃO
// ============================================================
function renderAssoc() {
  const rows = (DATA.assoc_rows || []).filter(r => rowMatches(r, ASIDX));
  const map = {};
  rows.forEach(r => {
    const nome = DATA.associacoes[r[ASIDX.A]];
    if (!map[nome]) map[nome] = { abr:0, proj:0 };
    map[nome].abr += r[ASIDX.ABR];
    map[nome].proj += r[ASIDX.PROJ];
  });
  let assocs = Object.entries(map)
    .map(([nome, v]) => ({ nome, ...v, var: v.proj - v.abr,
                           pct: v.abr > 0 ? (v.proj - v.abr)/v.abr*100 : null }))
    .filter(a => a.abr > 0 || a.proj > 0)
    .sort((a, b) => b.proj - a.proj);

  const tbody = document.getElementById('assoc-tbody');
  if (assocs.length === 0) {
    tbody.innerHTML = '<tr><td colspan="5" class="empty">Sem dados para os filtros aplicados.</td></tr>';
    return;
  }

  // Top 12 para os gráficos (associações relevantes)
  const top = assocs.slice(0, 12);
  _barChartH('chart-assoc-ytd', 'assoc-ytd',
    top.map(a => a.nome), top.map(a => a.proj), null);
  _barChartH('chart-assoc-var', 'assoc-var',
    top.map(a => a.nome), top.map(a => a.var),
    top.map(a => a.var >= 0 ? COLORS.green : COLORS.red));

  // Tabela completa
  tbody.innerHTML = assocs.map(a => `
    <tr>
      <td>${safe(a.nome)}</td>
      <td class="num">${fmtNum(a.abr)}</td>
      <td class="num"><b>${fmtNum(a.proj)}</b></td>
      <td class="num ${a.var >= 0 ? 'green' : 'red'}">${fmtNum(a.var)}</td>
      <td class="num ${a.var >= 0 ? 'green' : 'red'}">${a.pct === null ? '—' : fmtPct(a.pct)}</td>
    </tr>`).join('');
}

// Helper: gráfico de barras horizontais genérico
function _barChartH(canvasId, instanceKey, labels, values, colors) {
  if (chartInstances[instanceKey]) { chartInstances[instanceKey].destroy(); delete chartInstances[instanceKey]; }
  const maxV = Math.max(...values.map(v => Math.abs(v)), 1);
  chartInstances[instanceKey] = new Chart(document.getElementById(canvasId), {
    type: 'bar',
    data: { labels, datasets: [{ data: values, backgroundColor: colors || COLORS.navy }] },
    options: {
      indexAxis: 'y', responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: { callbacks: { label: ctx => fmtNum(ctx.raw) + ' un.' } },
        datalabels: { display: false }
      },
      scales: {
        x: { display: false, suggestedMax: maxV * 1.18, suggestedMin: colors ? -maxV*1.18 : 0 },
        y: { ticks: { font: { size: 10 } } }
      }
    },
    plugins: [{
      afterDatasetsDraw(chart) {
        const ctx = chart.ctx;
        chart.data.datasets[0].data.forEach((v, i) => {
          const bar = chart.getDatasetMeta(0).data[i];
          ctx.fillStyle = COLORS.navy; ctx.font = 'bold 10px sans-serif';
          ctx.textBaseline = 'middle';
          ctx.textAlign = v >= 0 ? 'left' : 'right';
          const off = v >= 0 ? 4 : -4;
          ctx.fillText(fmtNum(v), bar.x + off, bar.y);
        });
      }
    }]
  });
}

function renderGerentes() {
  const cont = document.getElementById('gerente-list');
  cont.innerHTML = '';
  const gIds = gerentesVisiveis();
  document.getElementById('contador-gerentes').textContent =
    `(${gIds.length} gerente${gIds.length===1?'':'s'})`;

  // Limpa charts antigos de gerente
  Object.keys(chartInstances).filter(k => k.startsWith('g-')).forEach(k => {
    chartInstances[k].destroy(); delete chartInstances[k];
  });

  if (gIds.length === 0) {
    cont.innerHTML = '<div style="grid-column:1/-1;padding:40px;text-align:center;color:var(--gray);">Nenhum gerente corresponde aos filtros.</div>';
    return;
  }

  gIds.forEach(gid => {
    const m = metricasGerente(gid);
    const gerente = DATA.gerentes[gid];
    const card = document.createElement('div');
    card.className = 'gerente-card';
    card.dataset.gid = gid;
    card.innerHTML = `
      <h3>${safe(gerente)}</h3>
      <div class="meta">${m.nSetores} setor${m.nSetores===1?'':'es'} · clique para ver detalhes</div>
      <div class="kpis">
        <div class="kpi"><small>${DATA.meta.label_acum_ano.replace("Acumulado ","")}</small><b>${fmtNum(m.acumJanAbr)}</b></div>
        <div class="kpi gold"><small>${DATA.meta.label_corrente_curto}-Proj</small><b>${fmtNum(m.proj)}</b></div>
        <div class="kpi ${m.varMaiAbr>=0?'green':'red'}"><small>Δ vs Abr</small><b>${fmtNum(m.varMaiAbr)}<br><small style="font-size:9px;">${fmtPct(m.varMaiAbrPct)}</small></b></div>
        <div class="kpi ${m.varS3S2>=0?'green':'red'}"><small>Δ ${labelDeltaSemanal(true)}</small><b>${fmtNum(m.varS3S2)}<br><small style="font-size:9px;">${fmtPct(m.varS3S2Pct)}</small></b></div>
      </div>
      <div class="chart-wrap"><canvas id="g-card-${gid}"></canvas></div>
    `;
    card.addEventListener('click', () => abrirModalGerente(gid));
    cont.appendChild(card);

    chartInstances[`g-${gid}`] = new Chart(document.getElementById(`g-card-${gid}`), {
      type: 'line',
      data: {
        labels: [...DATA.meta.meses_acum_labels, `${DATA.meta.label_corrente_curto}-P`],
        datasets: [{
          data: [...m.mesesAcum, m.proj],
          borderColor: COLORS.navy,
          backgroundColor: 'rgba(30,39,97,0.1)',
          tension: 0.3, pointRadius: 3,
          pointBackgroundColor: ctx => ctx.dataIndex === DATA.meta.n_meses_acum ? COLORS.gold : COLORS.navy,
          segment: { borderDash: ctx => ctx.p1DataIndex === DATA.meta.n_meses_acum ? [4,3] : undefined,
                     borderColor: ctx => ctx.p1DataIndex === DATA.meta.n_meses_acum ? COLORS.gold : COLORS.navy }
        }]
      },
      options: {
        responsive: true, maintainAspectRatio: false,
        plugins: { legend: { display: false }, tooltip: {
          callbacks: { label: ctx => fmtNum(ctx.raw) + ' un.' }
        } },
        scales: {
          x: { ticks: { font: { size: 9 } }, grid: { display: false } },
          y: { display: false }
        }
      }
    });
  });
}

function abrirModalGerente(gid) {
  const m = metricasGerente(gid);
  const gerente = DATA.gerentes[gid];
  document.getElementById('mg-titulo').textContent = 'Gerente: ' + gerente;
  document.getElementById('mg-meta').textContent =
    `${m.nSetores} setores  ·  TOP ${DATA.meta.top_setores_gerente} setores em destaque e ofensores`;

  // Labels dinâmicos UF/Bandeira (mostram mês corrente correto: Mai-Proj, Jun-Proj, etc.)
  const lblPeriodo = `${DATA.meta.label_corrente_curto}-Proj · variação vs ${DATA.meta.label_anterior_curto || 'período anterior'}`;
  document.getElementById('mg-h4-uf').textContent = `Desempenho por UF (${lblPeriodo})`;
  document.getElementById('mg-h4-band').textContent = `Desempenho por Bandeira (${lblPeriodo})`;

  // Aviso de filtro de família no Top setores/PDV/Produto
  const avisoGer = document.getElementById('top5-aviso-gerente');
  if (DATA.meta.top5_familia_label) {
    avisoGer.innerHTML = `⚠ <b>Top setores</b> e <b>PDV/Produto puxador</b> consideram apenas variação em <b>${DATA.meta.top5_familia_label}</b>. KPIs e gráficos acima refletem o portfólio completo.`;
    avisoGer.classList.add('show');
  } else {
    avisoGer.classList.remove('show');
  }

  // KPIs
  document.getElementById('mg-kpis').innerHTML = `
    <div class="kpi-card-lg">
      <small>ACUMULADO ${DATA.meta.label_acum_ano.toUpperCase().replace("ACUMULADO ","")} — ${DATA.meta.tipo_label}</small>
      <div class="val">${fmtNum(m.acumJanAbr)} un.</div>
      <div class="sub">Abr/${DATA.meta.mes_anterior.slice(2)}: ${fmtNum(m.abr)} un.</div>
    </div>
    <div class="kpi-card-lg gold">
      <small>${DATA.meta.label_corrente_curto}-Proj (${_label_mes_completo(DATA.meta.mes_corrente)} Projetado)</small>
      <div class="val">${fmtNum(m.proj)} un.</div>
      <div class="sub ${m.varMaiAbr>=0?'green':'red'}">${fmtNum(m.varMaiAbr)} un. vs Abr (${fmtPct(m.varMaiAbrPct)})</div>
    </div>
    <div class="kpi-card-lg ${m.varS3S2>=0?'green':'red'}">
      <small>${DATA.meta.label_comparativo_semanal}</small>
      <div class="val">${fmtNum(m.varS3S2)} un.</div>
      <div class="sub ${m.varS3S2>=0?'green':'red'}">${fmtPct(m.varS3S2Pct)}</div>
    </div>
    <div class="kpi-card-lg ${m.varMaiAbr>=0?'green':'red'}">
      <small>Δ ${DATA.meta.label_corrente_curto.toUpperCase()}-PROJ vs ABR (%)</small>
      <div class="val">${fmtPct(m.varMaiAbrPct)}</div>
      <div class="sub ${m.varMaiAbr>=0?'green':'red'}">${fmtNum(m.varMaiAbr)} un.</div>
    </div>
  `;

  // Limpar charts antigos do modal
  ['mg-evo','mg-sem','mg-uf','mg-band'].forEach(k => {
    if (chartInstances[k]) { chartInstances[k].destroy(); delete chartInstances[k]; }
  });

  // Chart evolução
  chartInstances['mg-evo'] = new Chart(document.getElementById('mg-chart-evo'), {
    type: 'line',
    data: {
      labels: [...DATA.meta.meses_acum_labels, `${DATA.meta.label_corrente_curto}-Proj`],
      datasets: [{
        data: [...m.mesesAcum, m.proj],
        borderColor: COLORS.navy, backgroundColor: 'rgba(30,39,97,0.1)',
        tension: 0.25, pointRadius: 6, pointHoverRadius: 8, borderWidth: 3,
        pointBackgroundColor: ctx => ctx.dataIndex === DATA.meta.n_meses_acum ? COLORS.gold : COLORS.navy,
        segment: {
          borderDash: ctx => ctx.p1DataIndex === DATA.meta.n_meses_acum ? [6,4] : undefined,
          borderColor: ctx => ctx.p1DataIndex === DATA.meta.n_meses_acum ? COLORS.gold : COLORS.navy
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

  // Chart Semanal — usa as duas últimas semanas detectadas dinamicamente
  chartInstances['mg-sem'] = new Chart(document.getElementById('mg-chart-sem'), {
    type: 'bar',
    data: {
      labels: labelsSemanaisChart(),
      datasets: [{
        data: m.penultima_semana_disponivel
          ? [m.s_prev, m.s_last]
          : [m.s_last],
        backgroundColor: [COLORS.ice, COLORS.navy],
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
          ctx.fillStyle = COLORS.navy; ctx.font = 'bold 11px sans-serif';
          ctx.textAlign = 'center';
          ctx.fillText(fmtNum(v), bar.x, bar.y - 6);
        });
      }
    }]
  });

  // Chart UF (horizontal bars)
  _renderDimChart('mg-chart-uf', 'mg-uf', m.ufList);

  // Chart Bandeira
  _renderDimChart('mg-chart-band', 'mg-band', m.bandList);

  // Tabelas destaque/ofensores
  document.getElementById('mg-tbl-destaques').innerHTML = renderTabelaSetorPdvMarca(m.destaques, 'green');
  document.getElementById('mg-tbl-ofensores').innerHTML = renderTabelaSetorPdvMarca(m.ofensores, 'red');

  document.getElementById('modal-gerente').classList.add('open');
}

function _renderDimChart(canvasId, instanceKey, items) {
  const canvas = document.getElementById(canvasId);
  if (!canvas) return;
  if (!items || items.length === 0) {
    const ctx = canvas.getContext('2d');
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    return;
  }
  const items8 = items.slice(0, 8);
  const labels = items8.map(it => String(it[0]).substring(0, 18));
  const values = items8.map(it => it[2]);
  const variations = items8.map(it => it[3]);
  const maxVal = Math.max(...values);

  chartInstances[instanceKey] = new Chart(canvas, {
    type: 'bar',
    data: {
      labels: labels,
      datasets: [{ data: values, backgroundColor: COLORS.navy }]
    },
    options: {
      indexAxis: 'y', responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: ctx => {
              const v = ctx.raw;
              const va = variations[ctx.dataIndex];
              const sign = va >= 0 ? '+' : '';
              return `${fmtNum(v)} un. (${sign}${fmtNum(va)})`;
            }
          }
        }
      },
      scales: {
        x: { display: false, max: maxVal * 1.4 },
        y: { ticks: { font: { size: 10 } } }
      }
    },
    plugins: [{
      afterDatasetsDraw(chart) {
        const ctx = chart.ctx;
        chart.data.datasets[0].data.forEach((v, i) => {
          const bar = chart.getDatasetMeta(0).data[i];
          const va = variations[i];
          const sign = va >= 0 ? '+' : '';
          const varColor = va >= 0 ? COLORS.green : COLORS.red;
          ctx.font = 'bold 9px sans-serif'; ctx.textBaseline = 'middle';
          ctx.textAlign = 'left';
          // Valor
          ctx.fillStyle = COLORS.navy;
          ctx.fillText(fmtNum(v), bar.x + 4, bar.y);
          // Variação (colorido)
          const valWidth = ctx.measureText(fmtNum(v)).width;
          ctx.fillStyle = varColor;
          ctx.fillText(`  (${sign}${fmtNum(va)})`, bar.x + 4 + valWidth, bar.y);
        });
      }
    }]
  });
}

function renderTabelaSetorPdvMarca(setores, defaultColor) {
  if (!setores || setores.length === 0) {
    return `<tr><td colspan="4" class="empty">Sem dados.</td></tr>`;
  }
  return setores.map(s => {
    const setorVarColor = s.var >= 0 ? 'green' : 'red';
    const direcao = defaultColor === 'green' ? 'alta' : 'queda';

    // PDV
    const pdvData = DATA.top_pdv_setor[s.sid];
    let pdvRow = '';
    if (pdvData && pdvData[direcao]) {
      const pdv = pdvData[direcao];
      const pdvVarColor = pdv[5] >= 0 ? 'green' : 'red';
      pdvRow = `
        <tr class="pdv-sub">
          <td>↳ PDV: ${safe(pdv[0])} · ${safe(pdv[1])}/${safe(pdv[2])}</td>
          <td class="num">${fmtNum(pdv[3])}</td>
          <td class="num">${fmtNum(pdv[4])}</td>
          <td class="num ${pdvVarColor}">${fmtNum(pdv[5])}</td>
        </tr>`;
    }

    // Marca
    const marcaData = DATA.top_marca_setor[s.sid];
    let marcaRow = '';
    if (marcaData && marcaData[direcao]) {
      const mk = marcaData[direcao];
      const mkVarColor = mk[3] >= 0 ? 'green' : 'red';
      marcaRow = `
        <tr class="pdv-sub">
          <td>↳ Produto: ${safe(mk[0])}</td>
          <td class="num">${fmtNum(mk[1])}</td>
          <td class="num">${fmtNum(mk[2])}</td>
          <td class="num ${mkVarColor}">${fmtNum(mk[3])}</td>
        </tr>`;
    }

    return `
      <tr class="setor-main">
        <td>${safe(s.setor)}</td>
        <td class="num">${fmtNum(s.abr)}</td>
        <td class="num">${fmtNum(s.proj)}</td>
        <td class="num ${setorVarColor}">${fmtNum(s.var)}</td>
      </tr>
      ${pdvRow}
      ${marcaRow}
    `;
  }).join('');
}

function fecharModalGerente() {
  document.getElementById('modal-gerente').classList.remove('open');
}
document.addEventListener('DOMContentLoaded', () => {
  const mg = document.getElementById('modal-gerente');
  if (mg) mg.addEventListener('click', e => { if (e.target.id === 'modal-gerente') fecharModalGerente(); });
});

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
        <div class="kpi"><small>${DATA.meta.label_acum_ano.replace("Acumulado ","")}</small><b>${fmtNum(m.ytd)}</b></div>
        <div class="kpi gold"><small>${DATA.meta.label_corrente_curto}-Proj</small><b>${fmtNum(m.proj)}</b></div>
        <div class="kpi ${m.varMaiAbr>=0?'green':'red'}"><small>Δ vs Abr</small><b>${fmtNum(m.varMaiAbr)} <small style="font-size:9px;">${fmtPct(m.varMaiAbrPct)}</small></b></div>
        <div class="kpi ${m.varS3S2>=0?'green':'red'}"><small>Δ ${labelDeltaSemanal(true)}</small><b>${fmtNum(m.varS3S2)} <small style="font-size:9px;">${fmtPct(m.varS3S2Pct)}</small></b></div>
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
        labels: [...DATA.meta.meses_acum_labels, `${DATA.meta.label_corrente_curto}-P`],
        datasets: [{
          data: [...m.mesesAcum, m.proj],
          borderColor: COLORS.navy,
          backgroundColor: 'rgba(30,39,97,0.1)',
          tension: 0.3,
          pointRadius: 3,
          pointBackgroundColor: ctx => ctx.dataIndex === DATA.meta.n_meses_acum ? COLORS.gold : COLORS.navy,
          segment: { borderDash: ctx => ctx.p1DataIndex === DATA.meta.n_meses_acum ? [4,3] : undefined,
                     borderColor: ctx => ctx.p1DataIndex === DATA.meta.n_meses_acum ? COLORS.gold : COLORS.navy }
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
  const maxYtd = Math.max(...marcas.slice(0, 15).map(m => m.ytd));
  chartInstances['marca-ytd'] = new Chart(document.getElementById('chart-marca-ytd'), {
    type: 'bar',
    data: {
      labels: marcas.slice(0, 15).map(m => m.nome),
      datasets: [{ data: marcas.slice(0, 15).map(m => m.ytd), backgroundColor: COLORS.navy }]
    },
    options: {
      indexAxis: 'y', responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: { callbacks: { label: ctx => fmtNum(ctx.raw) + ' un.' } },
        datalabels: { display: false }
      },
      scales: {
        x: { display: false, max: maxYtd * 1.18 },
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
      labels: marcas.slice(0, 15).map(m => m.nome),
      datasets: [{
        data: marcas.slice(0, 15).map(m => m.var),
        backgroundColor: marcas.slice(0, 15).map(m => m.var >= 0 ? COLORS.green : COLORS.red)
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

  // Aviso de filtro de família no Top 5
  const avisoSetor = document.getElementById('top5-aviso-setor');
  if (DATA.meta.top5_familia_label) {
    avisoSetor.innerHTML = `⚠ <b>Top 5 PDVs</b> considera apenas variação em <b>${DATA.meta.top5_familia_label}</b>. As métricas dos KPIs e gráficos acima refletem o portfólio completo.`;
    avisoSetor.classList.add('show');
  } else {
    avisoSetor.classList.remove('show');
  }

  const kpiHtml = `
    <div class="kpi-card-lg">
      <small>ACUMULADO ${DATA.meta.label_acum_ano.toUpperCase().replace("ACUMULADO ","")} — ${DATA.meta.tipo_label}</small>
      <div class="val">${fmtNum(m.ytd)} un.</div>
      <div class="sub">PDVs ativos no setor: ${m.pdvsTotal}</div>
    </div>
    <div class="kpi-card-lg gold">
      <small>${DATA.meta.label_corrente_curto}-Proj (${_label_mes_completo(DATA.meta.mes_corrente)} Projetado)</small>
      <div class="val">${fmtNum(m.proj)} un.</div>
      <div class="sub ${m.varMaiAbr>=0?'green':'red'}">${fmtNum(m.varMaiAbr)} un. vs Abr (${fmtPct(m.varMaiAbrPct)})</div>
    </div>
    <div class="kpi-card-lg ${m.varS3S2>=0?'green':'red'}">
      <small>${DATA.meta.label_comparativo_semanal}</small>
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
      labels: [...DATA.meta.meses_acum_labels, `${DATA.meta.label_corrente_curto}-Proj`],
      datasets: [{
        data: [...m.mesesAcum, m.proj],
        borderColor: COLORS.navy, backgroundColor: 'rgba(30,39,97,0.1)',
        tension: 0.25, pointRadius: 6, pointHoverRadius: 8, borderWidth: 3,
        pointBackgroundColor: ctx => ctx.dataIndex === DATA.meta.n_meses_acum ? COLORS.gold : COLORS.navy,
        segment: {
          borderDash: ctx => ctx.p1DataIndex === DATA.meta.n_meses_acum ? [6,4] : undefined,
          borderColor: ctx => ctx.p1DataIndex === DATA.meta.n_meses_acum ? COLORS.gold : COLORS.navy
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

  // Chart semanal — duas últimas semanas dinâmicas
  chartInstances['m-sem'] = new Chart(document.getElementById('m-chart-sem'), {
    type: 'bar',
    data: {
      labels: labelsSemanaisChart(),
      datasets: [{
        data: m.penultima_semana_disponivel
          ? [m.s_prev, m.s_last]
          : [m.s_last],
        backgroundColor: [COLORS.ice, COLORS.navy],
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

  // Tops — cada PDV vem com sub-linha do produto que mais puxou sua variação
  function _pdvRows(pdvs, color) {
    if (!pdvs.length) {
      const sign = color === 'green' ? '+' : '−';
      return `<tr><td colspan="4" class="empty">Nenhum PDV com variação ${sign}${DATA.meta.threshold} unid.</td></tr>`;
    }
    return pdvs.map(p => {
      const marca = p[PIDX.MARCA];
      const varMarca = p[PIDX.VAR_MARCA];
      const marcaColor = varMarca >= 0 ? 'green' : 'red';
      const subRow = (marca && marca !== '—')
        ? `<tr class="pdv-sub-row">
             <td class="pdv-sub-cell">↳ Produto: ${safe(marca)}</td>
             <td></td><td></td>
             <td class="num ${marcaColor}" style="font-size:10px;font-style:italic;">${fmtNum(varMarca)}</td>
           </tr>`
        : '';
      return `
        <tr>
          <td>${safe(p[PIDX.NAME])} · ${safe(p[PIDX.CITY])}/${safe(p[PIDX.UF])}</td>
          <td class="num">${fmtNum(p[PIDX.ABR])}</td>
          <td class="num"><b>${fmtNum(p[PIDX.PROJ])}</b></td>
          <td class="num ${color}">${fmtNum(p[PIDX.VAR])}</td>
        </tr>
        ${subRow}`;
    }).join('');
  }
  document.getElementById('m-tbl-altas').innerHTML = _pdvRows(m.altas, 'green');
  document.getElementById('m-tbl-quedas').innerHTML = _pdvRows(m.quedas, 'red');

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
  // Atualiza contador conforme visão atualmente exibida
  const viewAtual = document.querySelector('.view-tab.active').dataset.view;
  // Pré-seleciona a view do PPT igual à visão atual
  document.querySelector(`input[name="export-view"][value="${viewAtual}"]`).checked = true;
  atualizarContagemExport();
  document.querySelectorAll('input[name="export-view"]').forEach(r =>
    r.addEventListener('change', atualizarContagemExport)
  );
  document.getElementById('export-modal').classList.add('open');
}
function atualizarContagemExport() {
  const view = document.querySelector('input[name="export-view"]:checked').value;
  if (view === 'gerente') {
    document.getElementById('exp-count').textContent = gerentesVisiveis().length;
    document.getElementById('exp-count-label').textContent = 'gerente(s)';
  } else {
    document.getElementById('exp-count').textContent = setoresVisiveis().length;
    document.getElementById('exp-count-label').textContent = 'setor(es)';
  }
}
function fecharExport() {
  document.getElementById('export-modal').classList.remove('open');
}

async function exportarPPT() {
  fecharExport();
  const view = document.querySelector('input[name="export-view"]:checked').value;
  const scope = document.querySelector('input[name="export-scope"]:checked').value;

  const stateBackup = { ...filterState };
  if (scope === 'all') filterState = { gerente:[], setor:[], uf:[], bandeira:[], marca:[], distribuidor:[] };

  const ids = view === 'gerente' ? gerentesVisiveis() : setoresVisiveis();
  if (ids.length === 0) {
    filterState = stateBackup;
    alert(`Nenhum ${view} corresponde aos filtros.`);
    return;
  }

  const progress = document.getElementById('progress');
  const progressPct = document.getElementById('progress-pct');
  progress.classList.add('show');
  progressPct.textContent = '0%';

  await new Promise(r => setTimeout(r, 50));

  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'CUSTOM', width: 13.333, height: 7.5 });
  pptx.layout = 'CUSTOM';

  // Capa
  slideCapa(pptx, ids.length, view);
  // Produto (sempre)
  slideProduto(pptx);
  slideApresentacao(pptx);

  // Slides por gerente OU por setor
  const slideFn = view === 'gerente' ? slideGerente : slideSetor;
  for (let i = 0; i < ids.length; i++) {
    slideFn(pptx, ids[i]);
    const pct = Math.round(((i+1) / ids.length) * 100);
    progressPct.textContent = pct + '%';
    if (i % 2 === 0) await new Promise(r => setTimeout(r, 10));
  }

  filterState = stateBackup;

  const stamp = new Date().toISOString().slice(0,10).replace(/-/g,'');
  const fname = view === 'gerente'
    ? `MDTRS_Analise_Gerente_${stamp}.pptx`
    : `MDTRS_Analise_Setor_${stamp}.pptx`;
  await pptx.writeFile({ fileName: fname });

  progress.classList.remove('show');
}

function setoresTodos() {
  const ids = new Set();
  DATA.rows.forEach(r => ids.add(r[IDX.S]));
  return ids;
}

async function exportarResumoExecutivo() {
  // Resumo Executivo: 1 slide para o diretor comercial.
  // Calcula totais nacionais e métricas-chave de cada gerente, gera 1 slide e baixa.
  const progress = document.getElementById('progress');
  const progressPct = document.getElementById('progress-pct');
  progress.classList.add('show');
  progressPct.textContent = '0%';
  await new Promise(r => setTimeout(r, 50));

  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'CUSTOM', width: 13.333, height: 7.5 });
  pptx.layout = 'CUSTOM';

  slideResumoExecutivo(pptx);

  const stamp = new Date().toISOString().slice(0,10).replace(/-/g,'');
  await pptx.writeFile({ fileName: `MDTRS_Resumo_Executivo_${stamp}.pptx` });
  progressPct.textContent = '100%';
  progress.classList.remove('show');
}

function slideResumoExecutivo(pptx) {
  const s = pptx.addSlide();
  s.background = { color: 'FFFFFF' };

  // --- Cabeçalho navy ---
  s.addShape('rect', { x: 0, y: 0, w: 13.333, h: 0.75, fill: { color: '1E2761' } });
  s.addText('Resumo Executivo · MDTR', {
    x: 0.4, y: 0.10, w: 9, h: 0.4,
    fontSize: 22, bold: true, color: 'FFFFFF', fontFace: 'Georgia'
  });
  const gerentes = DATA.gerentes;
  s.addText(`${gerentes.length} gerentes · Acumulado ${DATA.meta.label_acum_ano.replace("Acumulado ","")} · ${DATA.meta.label_corrente_curto}-Proj · ${DATA.meta.tipo_label}`, {
    x: 0.4, y: 0.45, w: 9, h: 0.25,
    fontSize: 11, color: 'CADCFC', fontFace: 'Calibri'
  });
  s.addText(`Gerado em ${DATA.meta.gerado_em}`, {
    x: 9.5, y: 0.25, w: 3.5, h: 0.3,
    fontSize: 10, color: 'CADCFC', align: 'right', fontFace: 'Calibri'
  });
  // Faixa dourada
  s.addShape('rect', { x: 0, y: 0.75, w: 13.333, h: 0.04, fill: { color: 'C9A227' } });

  // --- KPIs nacionais (totalizando todos gerentes) ---
  let totJanAbr = 0, totProj = 0, totAbr = 0, totSPrev = 0, totSLast = 0;
  const dadosGerentes = [];
  for (let gid = 0; gid < gerentes.length; gid++) {
    const m = metricasGerente(gid);
    totJanAbr += m.acumJanAbr;
    totProj += m.proj;
    totAbr += m.abr;
    totSPrev += m.s_prev;
    totSLast += m.s_last;
    dadosGerentes.push({ gid, nome: gerentes[gid], m });
  }
  // Ordenar gerentes por variação % decrescente (best → worst)
  dadosGerentes.sort((a, b) => b.m.varMaiAbrPct - a.m.varMaiAbrPct);

  const varTotAbr = totProj - totAbr;
  const varTotAbrPct = totAbr > 0 ? (varTotAbr / totAbr * 100) : 0;
  const varTotSemanal = totSLast - totSPrev;
  const varTotSemanalPct = totSPrev > 0 ? (varTotSemanal / totSPrev * 100) : 0;

  const kpiW = 3.05, kpiH = 0.95, gap = 0.13, left0 = 0.4, kpiY = 1.0;
  _kpiPPT(s, left0, kpiY, kpiW, kpiH,
    `ACUMULADO ${DATA.meta.label_acum_ano.toUpperCase().replace('ACUMULADO ','')}`,
    fmtNum(totJanAbr) + ' un.',
    `Abr: ${fmtNum(totAbr)} un.`, '1E2761');
  _kpiPPT(s, left0 + (kpiW + gap), kpiY, kpiW, kpiH,
    `${DATA.meta.label_corrente_curto.toUpperCase()}-PROJ NACIONAL`,
    fmtNum(totProj) + ' un.',
    `${fmtNum(varTotAbr)} un. vs Abr (${fmtPct(varTotAbrPct)})`,
    'C9A227', 'C9A227');
  _kpiPPT(s, left0 + 2*(kpiW + gap), kpiY, kpiW, kpiH,
    DATA.meta.label_comparativo_semanal.toUpperCase(),
    fmtNum(varTotSemanal) + ' un.',
    fmtPct(varTotSemanalPct),
    varTotSemanal >= 0 ? '2E7D32' : 'C62828',
    varTotSemanal >= 0 ? '2E7D32' : 'C62828');
  _kpiPPT(s, left0 + 3*(kpiW + gap), kpiY, kpiW, kpiH,
    `Δ ${DATA.meta.label_corrente_curto.toUpperCase()}-PROJ vs ABR`,
    fmtPct(varTotAbrPct),
    `${fmtNum(varTotAbr)} un.`,
    varTotAbr >= 0 ? '2E7D32' : 'C62828',
    varTotAbr >= 0 ? '2E7D32' : 'C62828');

  // --- Título da tabela ---
  s.addText(`Desempenho por Gerente (ordenado por Δ% ${DATA.meta.label_corrente_curto}-Proj vs ${DATA.meta.label_anterior_curto})`, {
    x: 0.4, y: 2.10, w: 12.5, h: 0.3,
    fontSize: 12, bold: true, color: '1E2761', fontFace: 'Calibri'
  });

  // --- Tabela de gerentes ---
  // Colunas: Gerente | Setores | Jan-Abr | ${DATA.meta.label_corrente_curto}-Proj | Δ vs Abr | Δ % | Δ S3 vs S2 | Destaque | Ofensor
  const headers = [
    { text: 'Gerente', w: 1.5 },
    { text: 'Setores', w: 0.7, align: 'center' },
    { text: DATA.meta.label_acum_ano.replace('Acumulado ',''), w: 1.05, align: 'right' },
    { text: `${DATA.meta.label_corrente_curto}-Proj`, w: 1.05, align: 'right' },
    { text: 'Δ vs Abr', w: 0.85, align: 'right' },
    { text: 'Δ %', w: 0.85, align: 'right' },
    { text: 'Δ ' + labelDeltaSemanal(true), w: 1.05, align: 'right' },
    { text: '↑ Setor destaque', w: 2.45 },
    { text: '↓ Setor ofensor', w: 2.45 },
  ];
  let xCursor = 0.4;
  const tableY = 2.45;
  const rowH = 0.55;

  // Header row
  headers.forEach(h => {
    s.addShape('rect', { x: xCursor, y: tableY, w: h.w, h: 0.35,
      fill: { color: '1E2761' }, line: { color: '1E2761' } });
    s.addText(h.text, {
      x: xCursor + 0.08, y: tableY + 0.04, w: h.w - 0.16, h: 0.28,
      fontSize: 9.5, bold: true, color: 'FFFFFF', fontFace: 'Calibri',
      align: h.align || 'left', valign: 'middle'
    });
    xCursor += h.w;
  });

  // Data rows
  dadosGerentes.forEach((d, i) => {
    const y = tableY + 0.35 + (i * rowH);
    const isPos = d.m.varMaiAbr >= 0;
    const accentColor = isPos ? '2E7D32' : 'C62828';
    const bgRow = i % 2 === 0 ? 'F8F8F8' : 'FFFFFF';

    xCursor = 0.4;
    // Background da linha inteira
    headers.forEach(h => {
      s.addShape('rect', { x: xCursor, y, w: h.w, h: rowH,
        fill: { color: bgRow }, line: { color: 'E8E8E8', width: 0.5 } });
      xCursor += h.w;
    });

    // Barra colorida lateral indicando performance
    s.addShape('rect', { x: 0.4, y, w: 0.05, h: rowH, fill: { color: accentColor }, line: { color: accentColor } });

    xCursor = 0.4;
    // 1. Nome do gerente (com nº setores como sub)
    s.addText(d.nome, {
      x: xCursor + 0.12, y: y + 0.06, w: headers[0].w - 0.16, h: 0.25,
      fontSize: 11, bold: true, color: '1E2761', fontFace: 'Calibri'
    });
    s.addText(`${d.m.nSetores} setor${d.m.nSetores === 1 ? '' : 'es'}`, {
      x: xCursor + 0.12, y: y + 0.30, w: headers[0].w - 0.16, h: 0.20,
      fontSize: 8.5, color: '595959', italic: true, fontFace: 'Calibri'
    });
    xCursor += headers[0].w;

    // 2. Setores (centro)
    s.addText(String(d.m.nSetores), {
      x: xCursor, y: y + 0.15, w: headers[1].w, h: 0.30,
      fontSize: 11, color: '1E2761', align: 'center', fontFace: 'Calibri', valign: 'middle'
    });
    xCursor += headers[1].w;

    // 3. Jan-Abr
    s.addText(fmtNum(d.m.acumJanAbr), {
      x: xCursor, y: y + 0.15, w: headers[2].w - 0.08, h: 0.30,
      fontSize: 10.5, color: '1E2761', align: 'right', fontFace: 'Calibri', valign: 'middle'
    });
    xCursor += headers[2].w;

    // 4. ${DATA.meta.label_corrente_curto}-Proj
    s.addText(fmtNum(d.m.proj), {
      x: xCursor, y: y + 0.15, w: headers[3].w - 0.08, h: 0.30,
      fontSize: 10.5, bold: true, color: 'C9A227', align: 'right', fontFace: 'Calibri', valign: 'middle'
    });
    xCursor += headers[3].w;

    // 5. Δ vs Abr (un)
    s.addText(fmtNum(d.m.varMaiAbr), {
      x: xCursor, y: y + 0.15, w: headers[4].w - 0.08, h: 0.30,
      fontSize: 10.5, bold: true, color: accentColor,
      align: 'right', fontFace: 'Calibri', valign: 'middle'
    });
    xCursor += headers[4].w;

    // 6. Δ %
    s.addText(fmtPct(d.m.varMaiAbrPct), {
      x: xCursor, y: y + 0.15, w: headers[5].w - 0.08, h: 0.30,
      fontSize: 10.5, bold: true, color: accentColor,
      align: 'right', fontFace: 'Calibri', valign: 'middle'
    });
    xCursor += headers[5].w;

    // 7. Δ S3 vs S2
    const s3s2Color = d.m.varS3S2 >= 0 ? '2E7D32' : 'C62828';
    s.addText(fmtNum(d.m.varS3S2) + ' (' + fmtPct(d.m.varS3S2Pct) + ')', {
      x: xCursor, y: y + 0.15, w: headers[6].w - 0.08, h: 0.30,
      fontSize: 9.5, color: s3s2Color,
      align: 'right', fontFace: 'Calibri', valign: 'middle'
    });
    xCursor += headers[6].w;

    // 8. Setor destaque (top 1) — cor reflete sinal real
    const dest = d.m.destaques && d.m.destaques[0];
    if (dest) {
      s.addText(dest.setor, {
        x: xCursor + 0.08, y: y + 0.06, w: headers[7].w - 0.16, h: 0.22,
        fontSize: 9.5, bold: true, color: '1E2761', fontFace: 'Calibri'
      });
      const destColor = dest.var >= 0 ? '2E7D32' : 'C62828';
      s.addText(`${fmtNum(dest.var)} un. (${fmtPct(dest.varPct)})`, {
        x: xCursor + 0.08, y: y + 0.28, w: headers[7].w - 0.16, h: 0.22,
        fontSize: 8.5, color: destColor, bold: true, fontFace: 'Calibri'
      });
    } else {
      s.addText('—', { x: xCursor + 0.08, y: y + 0.15, w: headers[7].w - 0.16, h: 0.30,
        fontSize: 9, color: '999999', fontFace: 'Calibri', valign: 'middle' });
    }
    xCursor += headers[7].w;

    // 9. Setor ofensor (top 1) — cor reflete sinal real (alguns gerentes têm todos positivos)
    const ofen = d.m.ofensores && d.m.ofensores[0];
    if (ofen) {
      s.addText(ofen.setor, {
        x: xCursor + 0.08, y: y + 0.06, w: headers[8].w - 0.16, h: 0.22,
        fontSize: 9.5, bold: true, color: '1E2761', fontFace: 'Calibri'
      });
      const ofenColor = ofen.var >= 0 ? '2E7D32' : 'C62828';
      s.addText(`${fmtNum(ofen.var)} un. (${fmtPct(ofen.varPct)})`, {
        x: xCursor + 0.08, y: y + 0.28, w: headers[8].w - 0.16, h: 0.22,
        fontSize: 8.5, color: ofenColor, bold: true, fontFace: 'Calibri'
      });
    } else {
      s.addText('—', { x: xCursor + 0.08, y: y + 0.15, w: headers[8].w - 0.16, h: 0.30,
        fontSize: 9, color: '999999', fontFace: 'Calibri', valign: 'middle' });
    }
  });

  // --- Rodapé ---
  const totalRowsH = 0.35 + (dadosGerentes.length * rowH);
  const footerY = tableY + totalRowsH + 0.15;
  if (footerY < 7.0) {
    s.addText(`Fonte: MDTRS_FV_TOTAL  ·  Variação ${DATA.meta.label_corrente_curto}-Proj vs ${DATA.meta.label_anterior_curto}  ·  Ordenado pelo gerente com melhor Δ% no mês corrente`, {
      x: 0.4, y: footerY, w: 12.5, h: 0.3,
      fontSize: 9, italic: true, color: '595959', fontFace: 'Calibri'
    });
  }
}

function slideGerente(pptx, gid) {
  // Gera 2 slides por gerente: A (visão geral) + B (top setores)
  _slideGerenteOverview(pptx, gid);
  _slideGerenteSetores(pptx, gid);
}

function _slideGerenteOverview(pptx, gid) {
  const m = metricasGerente(gid);
  const gerente = DATA.gerentes[gid];
  const s = pptx.addSlide();
  _headerPPT(s, `Gerente: ${gerente}`,
             `${m.nSetores} setores  ·  Visão Geral`);
  s.addText('EVOLUÇÃO · SEMANAL · UF · BANDEIRA', {
    x: 9.5, y: 0.2, w: 3.5, h: 0.4, fontSize: 9, bold: true, color: 'CADCFC', align: 'right', fontFace: 'Calibri'
  });

  // KPIs
  const kpiW = 3.05, kpiH = 0.95, gap = 0.13, left0 = 0.4, kpiY = 1.0;
  _kpiPPT(s, left0, kpiY, kpiW, kpiH,
    `ACUMULADO ${DATA.meta.label_acum_ano.toUpperCase().replace("ACUMULADO ","")} — ${DATA.meta.tipo_label}`,
    fmtNum(m.acumJanAbr) + ' un.',
    `Abr: ${fmtNum(m.abr)} un.`, '1E2761');
  _kpiPPT(s, left0 + (kpiW + gap), kpiY, kpiW, kpiH,
    `${DATA.meta.label_corrente_curto.toUpperCase()}-PROJ (${_label_mes_completo(DATA.meta.mes_corrente)} Projetado)`,
    fmtNum(m.proj) + ' un.',
    `${fmtNum(m.varMaiAbr)} un. vs Abr (${fmtPct(m.varMaiAbrPct)})`,
    'C9A227', 'C9A227');
  _kpiPPT(s, left0 + 2*(kpiW + gap), kpiY, kpiW, kpiH,
    DATA.meta.label_comparativo_semanal.toUpperCase(),
    fmtNum(m.varS3S2) + ' un.',
    fmtPct(m.varS3S2Pct),
    m.varS3S2 >= 0 ? '2E7D32' : 'C62828',
    m.varS3S2 >= 0 ? '2E7D32' : 'C62828');
  _kpiPPT(s, left0 + 3*(kpiW + gap), kpiY, kpiW, kpiH,
    `Δ ${DATA.meta.label_corrente_curto.toUpperCase()}-PROJ vs ABR`,
    fmtPct(m.varMaiAbrPct),
    `${fmtNum(m.varMaiAbr)} un.`,
    m.varMaiAbr >= 0 ? '2E7D32' : 'C62828',
    m.varMaiAbr >= 0 ? '2E7D32' : 'C62828');

  // Linha 1: Evolução mensal (larga) + Semanal (estreito)
  s.addText(`Evolução Mensal — ${DATA.meta.label_corrente_curto}/${DATA.meta.mes_corrente.slice(2,4)}-Proj`,
    { x: 0.4, y: 2.15, w: 8.0, h: 0.3, fontSize: 11, bold: true, color: '1E2761', fontFace: 'Calibri' });
  s.addChart(pptx.ChartType.line, [
    { name: 'Real',
      labels: [...DATA.meta.meses_acum_labels, `${DATA.meta.label_corrente_curto}-Proj`],
      values: [...m.mesesAcum, m.proj] }
  ], {
    x: 0.4, y: 2.45, w: 8.0, h: 2.2,
    showLegend: false, chartColors: ['1E2761'],
    lineSize: 3, lineDataSymbol: 'circle', lineDataSymbolSize: 7,
    showValue: true, dataLabelFontSize: 9, dataLabelColor: '1E2761', dataLabelPosition: 't',
    catAxisLabelFontSize: 10, valAxisLabelFontSize: 9,
    valGridLine: { style: 'dash', color: 'E8E8E8' },
    catGridLine: { style: 'none' },
  });

  s.addText(`Desempenho Semanal — ${_label_mes_completo(DATA.meta.mes_corrente)}/${DATA.meta.mes_corrente.slice(2,4)}`,
    { x: 8.6, y: 2.15, w: 4.3, h: 0.3, fontSize: 11, bold: true, color: '1E2761', fontFace: 'Calibri' });
  s.addChart(pptx.ChartType.bar, [
    { name: 'Semana',
      labels: labelsSemanaisChart(),
      values: m.penultima_semana_disponivel ? [m.s_prev, m.s_last] : [m.s_last] }
  ], {
    x: 8.6, y: 2.45, w: 4.3, h: 2.2,
    barDir: 'col', barGrouping: 'standard',
    chartColors: ['1E2761'],
    showLegend: false, showValue: true,
    dataLabelFontSize: 10, dataLabelColor: '1E2761', dataLabelPosition: 'outEnd',
    catAxisLabelFontSize: 11, valAxisLabelFontSize: 9,
    valGridLine: { style: 'dash', color: 'E8E8E8' },
  });

  // Linha 2: UF + Bandeira (tabelas compactas em vez de chart — mais claro c/ variação)
  _dimensaoTabelaPPT(s, 0.4, 4.85, 6.3, m.ufList, `Desempenho por UF (${DATA.meta.label_corrente_curto}-Proj)`, 'UF');
  _dimensaoTabelaPPT(s, 6.9, 4.85, 6.0, m.bandList, `Desempenho por Bandeira (${DATA.meta.label_corrente_curto}-Proj)`, 'Bandeira');
}

function _dimensaoTabelaPPT(s, x, y, w, items, titulo, labelCol) {
  s.addText(titulo, { x, y, w, h: 0.28,
    fontSize: 11, bold: true, color: '1E2761', fontFace: 'Calibri' });

  if (!items || items.length === 0) {
    s.addText('Sem dados.', { x, y: y + 1.0, w, h: 0.4,
      fontSize: 10, color: '595959', align: 'center', italic: true, fontFace: 'Calibri' });
    return;
  }

  const top = items.slice(0, 6);
  const rows = [[
    { text: labelCol, options: { bold: true, color: 'FFFFFF', fill: { color: '1E2761' }, fontSize: 9.5, fontFace: 'Calibri' } },
    { text: 'Abr', options: { bold: true, color: 'FFFFFF', fill: { color: '1E2761' }, fontSize: 9.5, align: 'right', fontFace: 'Calibri' } },
    { text: `${DATA.meta.label_corrente_curto}-Proj`, options: { bold: true, color: 'FFFFFF', fill: { color: '1E2761' }, fontSize: 9.5, align: 'right', fontFace: 'Calibri' } },
    { text: 'Δ unid', options: { bold: true, color: 'FFFFFF', fill: { color: '1E2761' }, fontSize: 9.5, align: 'right', fontFace: 'Calibri' } },
  ]];
  top.forEach((it, i) => {
    const bg = i % 2 === 0 ? 'F8F8F8' : 'FFFFFF';
    const varColor = it[3] >= 0 ? '2E7D32' : 'C62828';
    rows.push([
      { text: String(it[0]).substring(0, 22),
        options: { color: '1E2761', fill: { color: bg }, fontSize: 9, bold: true, fontFace: 'Calibri' } },
      { text: fmtNum(it[1]),
        options: { color: '595959', fill: { color: bg }, fontSize: 9, align: 'right', fontFace: 'Calibri' } },
      { text: fmtNum(it[2]),
        options: { color: '1E2761', fill: { color: bg }, fontSize: 9, bold: true, align: 'right', fontFace: 'Calibri' } },
      { text: fmtNum(it[3]),
        options: { color: varColor, fill: { color: bg }, fontSize: 9, bold: true, align: 'right', fontFace: 'Calibri' } },
    ]);
  });

  s.addTable(rows, {
    x, y: y + 0.32, w, colW: [w * 0.45, w * 0.18, w * 0.18, w * 0.19],
    rowH: 0.27, fontFace: 'Calibri', border: { type: 'solid', color: 'E8E8E8', pt: 0.5 }
  });
}

function _slideGerenteSetores(pptx, gid) {
  const m = metricasGerente(gid);
  const gerente = DATA.gerentes[gid];
  const s = pptx.addSlide();
  _headerPPT(s, `Gerente: ${gerente}`,
             `Top ${DATA.meta.top_setores_gerente} setores · PDV e Marca de maior impacto`);
  s.addText('TOP SETORES · PDV · MARCA', {
    x: 9.5, y: 0.2, w: 3.5, h: 0.4, fontSize: 9, bold: true, color: 'CADCFC', align: 'right', fontFace: 'Calibri'
  });

  _setorPdvMarcaTabelaPPT(s, 0.4, 1.05, 6.3, m.destaques,
    `↑ TOP ${DATA.meta.top_setores_gerente} SETORES em DESTAQUE — Maior alta (${DATA.meta.label_corrente_curto}-Proj vs ${DATA.meta.label_anterior_curto})`,
    '2E7D32', 'alta');
  _setorPdvMarcaTabelaPPT(s, 6.9, 1.05, 6.0, m.ofensores,
    `↓ TOP ${DATA.meta.top_setores_gerente} SETORES OFENSORES — Menor performance (${DATA.meta.label_corrente_curto}-Proj vs ${DATA.meta.label_anterior_curto})`,
    'C62828', 'queda');
}

function _setorPdvMarcaTabelaPPT(s, x, y, w, setores, titulo, accentColor, direcao) {
  s.addText(titulo, { x, y, w, h: 0.28,
    fontSize: 11, bold: true, color: accentColor, fontFace: 'Calibri' });

  if (!setores || setores.length === 0) {
    s.addShape('rect', { x, y: y + 0.32, w, h: 5.0, fill: { color: 'F8F8F8' }, line: { color: 'F8F8F8' } });
    s.addText('Sem dados.', { x, y: y + 2.0, w, h: 0.4,
      fontSize: 10, color: '595959', align: 'center', italic: true, fontFace: 'Calibri' });
    return;
  }

  const rows = [[
    { text: 'Setor  ·  ↳ PDV  ·  ↳ Marca',
      options: { bold: true, color: 'FFFFFF', fill: { color: '1E2761' }, fontSize: 9, fontFace: 'Calibri' } },
    { text: 'Abr', options: { bold: true, color: 'FFFFFF', fill: { color: '1E2761' }, fontSize: 9, align: 'right', fontFace: 'Calibri' } },
    { text: `${DATA.meta.label_corrente_curto}-Proj`, options: { bold: true, color: 'FFFFFF', fill: { color: '1E2761' }, fontSize: 9, align: 'right', fontFace: 'Calibri' } },
    { text: 'Δ unid', options: { bold: true, color: 'FFFFFF', fill: { color: '1E2761' }, fontSize: 9, align: 'right', fontFace: 'Calibri' } },
  ]];

  setores.forEach((it, i) => {
    const bgMain = i % 2 === 0 ? 'F0F2F5' : 'FFFFFF';
    const bgSub = i % 2 === 0 ? 'FAFAFA' : 'F8F8F8';
    const setorVarColor = it.var >= 0 ? '2E7D32' : 'C62828';

    // Linha SETOR
    rows.push([
      { text: String(it.setor).substring(0, 42),
        options: { color: '1E2761', fill: { color: bgMain }, fontSize: 9.5, bold: true, fontFace: 'Calibri' } },
      { text: fmtNum(it.abr),
        options: { color: '595959', fill: { color: bgMain }, fontSize: 9, align: 'right', fontFace: 'Calibri' } },
      { text: fmtNum(it.proj),
        options: { color: '1E2761', fill: { color: bgMain }, fontSize: 9, bold: true, align: 'right', fontFace: 'Calibri' } },
      { text: fmtNum(it.var),
        options: { color: setorVarColor, fill: { color: bgMain }, fontSize: 9, bold: true, align: 'right', fontFace: 'Calibri' } },
    ]);

    // Linha PDV
    const pdvData = DATA.top_pdv_setor[it.sid];
    const pdv = pdvData ? pdvData[direcao] : null;
    if (pdv) {
      const pdvVarColor = pdv[5] >= 0 ? '2E7D32' : 'C62828';
      rows.push([
        { text: `   ↳ PDV: ${pdv[0]}  ·  ${pdv[1]}/${pdv[2]}`,
          options: { color: '595959', fill: { color: bgSub }, fontSize: 8.5, italic: true, fontFace: 'Calibri' } },
        { text: fmtNum(pdv[3]),
          options: { color: '595959', fill: { color: bgSub }, fontSize: 8.5, italic: true, align: 'right', fontFace: 'Calibri' } },
        { text: fmtNum(pdv[4]),
          options: { color: '595959', fill: { color: bgSub }, fontSize: 8.5, italic: true, align: 'right', fontFace: 'Calibri' } },
        { text: fmtNum(pdv[5]),
          options: { color: pdvVarColor, fill: { color: bgSub }, fontSize: 8.5, italic: true, bold: true, align: 'right', fontFace: 'Calibri' } },
      ]);
    } else {
      rows.push([
        { text: '   ↳ PDV: —',
          options: { color: '595959', fill: { color: bgSub }, fontSize: 8.5, italic: true, fontFace: 'Calibri' } },
        { text: '—', options: { color: '595959', fill: { color: bgSub }, fontSize: 8.5, italic: true, align: 'right', fontFace: 'Calibri' } },
        { text: '—', options: { color: '595959', fill: { color: bgSub }, fontSize: 8.5, italic: true, align: 'right', fontFace: 'Calibri' } },
        { text: '—', options: { color: '595959', fill: { color: bgSub }, fontSize: 8.5, italic: true, align: 'right', fontFace: 'Calibri' } },
      ]);
    }

    // Linha MARCA
    const marcaData = DATA.top_marca_setor[it.sid];
    const marca = marcaData ? marcaData[direcao] : null;
    if (marca) {
      const mkVarColor = marca[3] >= 0 ? '2E7D32' : 'C62828';
      rows.push([
        { text: `   ↳ Produto: ${marca[0]}`,
          options: { color: '595959', fill: { color: bgSub }, fontSize: 8.5, italic: true, fontFace: 'Calibri' } },
        { text: fmtNum(marca[1]),
          options: { color: '595959', fill: { color: bgSub }, fontSize: 8.5, italic: true, align: 'right', fontFace: 'Calibri' } },
        { text: fmtNum(marca[2]),
          options: { color: '595959', fill: { color: bgSub }, fontSize: 8.5, italic: true, align: 'right', fontFace: 'Calibri' } },
        { text: fmtNum(marca[3]),
          options: { color: mkVarColor, fill: { color: bgSub }, fontSize: 8.5, italic: true, bold: true, align: 'right', fontFace: 'Calibri' } },
      ]);
    } else {
      rows.push([
        { text: '   ↳ Produto: —',
          options: { color: '595959', fill: { color: bgSub }, fontSize: 8.5, italic: true, fontFace: 'Calibri' } },
        { text: '—', options: { color: '595959', fill: { color: bgSub }, fontSize: 8.5, italic: true, align: 'right', fontFace: 'Calibri' } },
        { text: '—', options: { color: '595959', fill: { color: bgSub }, fontSize: 8.5, italic: true, align: 'right', fontFace: 'Calibri' } },
        { text: '—', options: { color: '595959', fill: { color: bgSub }, fontSize: 8.5, italic: true, align: 'right', fontFace: 'Calibri' } },
      ]);
    }
  });

  s.addTable(rows, {
    x, y: y + 0.32, w, colW: [w * 0.55, w * 0.15, w * 0.15, w * 0.15],
    rowH: 0.25, fontFace: 'Calibri', border: { type: 'solid', color: 'E8E8E8', pt: 0.5 }
  });
}

// (Mantida para compatibilidade — usada apenas se algo chamar antiga assinatura)
function _setorPdvTabelaPPT(s, x, y, w, setores, titulo, accentColor, direcao) {
  s.addText(titulo, { x, y, w, h: 0.28,
    fontSize: 10.5, bold: true, color: accentColor, fontFace: 'Calibri' });

  if (!setores || setores.length === 0) {
    s.addShape('rect', { x, y: y + 0.32, w, h: 3.0, fill: { color: 'F8F8F8' }, line: { color: 'F8F8F8' } });
    s.addText('Sem dados.', { x, y: y + 1.2, w, h: 0.4,
      fontSize: 10, color: '595959', align: 'center', italic: true, fontFace: 'Calibri' });
    return;
  }

  // Cabeçalho
  const rows = [[
    { text: 'Setor  ·  ↳ PDV de maior impacto',
      options: { bold: true, color: 'FFFFFF', fill: { color: '1E2761' }, fontSize: 9, fontFace: 'Calibri' } },
    { text: 'Abr', options: { bold: true, color: 'FFFFFF', fill: { color: '1E2761' }, fontSize: 9, align: 'right', fontFace: 'Calibri' } },
    { text: `${DATA.meta.label_corrente_curto}-Proj`, options: { bold: true, color: 'FFFFFF', fill: { color: '1E2761' }, fontSize: 9, align: 'right', fontFace: 'Calibri' } },
    { text: 'Δ unid', options: { bold: true, color: 'FFFFFF', fill: { color: '1E2761' }, fontSize: 9, align: 'right', fontFace: 'Calibri' } },
  ]];

  setores.forEach((it, i) => {
    const bgSetor = i % 2 === 0 ? 'F0F2F5' : 'FFFFFF';
    const bgPdv = i % 2 === 0 ? 'FAFAFA' : 'F8F8F8';
    const setorVarColor = it.var >= 0 ? '2E7D32' : 'C62828';

    rows.push([
      { text: String(it.setor).substring(0, 42),
        options: { color: '1E2761', fill: { color: bgSetor }, fontSize: 9.5, bold: true, fontFace: 'Calibri' } },
      { text: fmtNum(it.abr),
        options: { color: '595959', fill: { color: bgSetor }, fontSize: 9, align: 'right', fontFace: 'Calibri' } },
      { text: fmtNum(it.proj),
        options: { color: '1E2761', fill: { color: bgSetor }, fontSize: 9, bold: true, align: 'right', fontFace: 'Calibri' } },
      { text: fmtNum(it.var),
        options: { color: setorVarColor, fill: { color: bgSetor }, fontSize: 9, bold: true, align: 'right', fontFace: 'Calibri' } },
    ]);

    // Linha PDV (recuada, em itálico)
    const pdvData = DATA.top_pdv_setor[it.sid];
    const pdv = pdvData ? pdvData[direcao] : null;
    if (pdv) {
      const pdvVarColor = pdv[5] >= 0 ? '2E7D32' : 'C62828';
      rows.push([
        { text: `   ↳ ${pdv[0]}  ·  ${pdv[1]}/${pdv[2]}`,
          options: { color: '595959', fill: { color: bgPdv }, fontSize: 8.5, italic: true, fontFace: 'Calibri' } },
        { text: fmtNum(pdv[3]),
          options: { color: '595959', fill: { color: bgPdv }, fontSize: 8.5, italic: true, align: 'right', fontFace: 'Calibri' } },
        { text: fmtNum(pdv[4]),
          options: { color: '595959', fill: { color: bgPdv }, fontSize: 8.5, italic: true, align: 'right', fontFace: 'Calibri' } },
        { text: fmtNum(pdv[5]),
          options: { color: pdvVarColor, fill: { color: bgPdv }, fontSize: 8.5, italic: true, bold: true, align: 'right', fontFace: 'Calibri' } },
      ]);
    } else {
      rows.push([
        { text: '   ↳ sem PDV identificado',
          options: { color: '595959', fill: { color: bgPdv }, fontSize: 8.5, italic: true, fontFace: 'Calibri' } },
        { text: '—', options: { color: '595959', fill: { color: bgPdv }, fontSize: 8.5, italic: true, align: 'right', fontFace: 'Calibri' } },
        { text: '—', options: { color: '595959', fill: { color: bgPdv }, fontSize: 8.5, italic: true, align: 'right', fontFace: 'Calibri' } },
        { text: '—', options: { color: '595959', fill: { color: bgPdv }, fontSize: 8.5, italic: true, align: 'right', fontFace: 'Calibri' } },
      ]);
    }
  });

  s.addTable(rows, {
    x, y: y + 0.32, w, colW: [w * 0.55, w * 0.15, w * 0.15, w * 0.15],
    rowH: 0.27, fontFace: 'Calibri', border: { type: 'solid', color: 'E8E8E8', pt: 0.5 }
  });
}

function slideCapa(pptx, nItems, view) {
  view = view || 'setor';
  const titulo = view === 'gerente' ? 'Análise MDTR por Gerente' : 'Análise MDTR por Setor';
  const itemLabel = view === 'gerente'
    ? `${nItems} gerente${nItems===1?'':'s'}`
    : `${nItems} setor${nItems===1?'':'es'}`;
  const linhaPerf = view === 'gerente'
    ? 'Performance agregada por gerente · Top setores em alta/ofensores · Maior PDV impactador em cada setor · Visão de Produto'
    : 'Performance mensal por setor · Top PDVs em alta e queda · ${DATA.meta.label_comparativo_semanal} · Visão de Produto';

  const s = pptx.addSlide();
  s.background = { color: '1E2761' };
  s.addShape('rect', { x: 13.15, y: 0, w: 0.18, h: 7.5, fill: { color: 'C9A227' } });
  s.addText('ALCON | COMMERCIAL INTELLIGENCE', {
    x: 0.8, y: 1.0, w: 11.5, h: 0.5, fontSize: 14, bold: true, color: 'CADCFC', fontFace: 'Calibri'
  });
  s.addText(titulo, {
    x: 0.8, y: 1.6, w: 11.5, h: 1.5, fontSize: 44, bold: true, color: 'FFFFFF', fontFace: 'Georgia'
  });
  const desc = filtrosDescritivos();
  s.addText(`Período: ${DATA.meta.ytd_label} → ${DATA.meta.mes_corrente}-Proj  ·  ${itemLabel}  ·  Tipo: ${DATA.meta.tipo_label}`, {
    x: 0.8, y: 3.2, w: 11.5, h: 0.4, fontSize: 16, color: 'CADCFC', fontFace: 'Calibri'
  });
  if (desc) {
    s.addText(`Filtros: ${desc}`, {
      x: 0.8, y: 3.6, w: 11.5, h: 0.4, fontSize: 12, color: 'CADCFC', italic: true, fontFace: 'Calibri'
    });
  }
  s.addText(linhaPerf, {
    x: 0.8, y: 4.2, w: 11.5, h: 1.5, fontSize: 14, color: 'FFFFFF', fontFace: 'Calibri'
  });
  s.addText(`Gerado em ${new Date().toLocaleString('pt-BR')}  ·  Fonte: MDTRS_FV_TOTAL`, {
    x: 0.8, y: 6.6, w: 11.5, h: 0.4, fontSize: 10, color: 'CADCFC', fontFace: 'Calibri'
  });
}

function filtrosDescritivos() {
  const partes = [];
  const fmt = (label, arr) => {
    if (!arr || arr.length === 0) return;
    if (arr.length === 1) partes.push(`${label}=${arr[0]}`);
    else partes.push(`${label}=${arr.length} selec.`);
  };
  fmt('Gerente', filterState.gerente);
  fmt('Setor', filterState.setor);
  fmt('UF', filterState.uf);
  fmt('Bandeira', filterState.bandeira);
  fmt('Marca', filterState.marca);
  fmt('Distribuidor', filterState.distribuidor);
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
    `ACUMULADO ${DATA.meta.label_acum_ano.toUpperCase().replace("ACUMULADO ","")} — ${DATA.meta.tipo_label}`,
    fmtNum(m.ytd) + ' un.',
    `PDVs ativos: ${m.pdvsTotal}`, '1E2761');
  _kpiPPT(s, left0 + (kpiW + gap), kpiY, kpiW, kpiH,
    `${DATA.meta.label_corrente_curto.toUpperCase()}-PROJ (${_label_mes_completo(DATA.meta.mes_corrente)} Projetado)`,
    fmtNum(m.proj) + ' un.',
    `${fmtNum(m.varMaiAbr)} un. vs Abr (${fmtPct(m.varMaiAbrPct)})`,
    'C9A227', 'C9A227');
  _kpiPPT(s, left0 + 2*(kpiW + gap), kpiY, kpiW, kpiH,
    DATA.meta.label_comparativo_semanal.toUpperCase(),
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
    { name: 'Real',
      labels: [...DATA.meta.meses_acum_labels, `${DATA.meta.label_corrente_curto}-Proj`],
      values: [...m.mesesAcum, m.proj] }
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
  s.addText(`Desempenho Semanal — ${DATA.meta.label_corrente_curto}`,
    { x: 8.2, y: 2.3, w: 4.7, h: 0.3, fontSize: 11, bold: true, color: '1E2761', fontFace: 'Calibri' });
  s.addChart(pptx.ChartType.bar, [
    { name: 'Semana',
      labels: labelsSemanaisChart(),
      values: m.penultima_semana_disponivel ? [m.s_prev, m.s_last] : [m.s_last] }
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
  _topTabelaPPT(s, 0.4, 5.25, 6.3, m.altas, `↑ TOP 5 PDVs em ALTA (${DATA.meta.label_corrente_curto}-Proj vs ${DATA.meta.label_anterior_curto} ≥ +${DATA.meta.threshold} unid)`, '2E7D32');
  _topTabelaPPT(s, 6.9, 5.25, 6.0, m.quedas, `↓ TOP 5 PDVs em QUEDA (${DATA.meta.label_corrente_curto}-Proj vs ${DATA.meta.label_anterior_curto} ≤ −${DATA.meta.threshold} unid)`, 'C62828');
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
    { text: 'PDV · ↳ Produto puxador', options: { bold: true, color: 'FFFFFF', fill: { color: '1E2761' }, fontSize: 9, fontFace: 'Calibri' } },
    { text: 'Abr', options: { bold: true, color: 'FFFFFF', fill: { color: '1E2761' }, fontSize: 9, align: 'right', fontFace: 'Calibri' } },
    { text: `${DATA.meta.label_corrente_curto}-Proj`, options: { bold: true, color: 'FFFFFF', fill: { color: '1E2761' }, fontSize: 9, align: 'right', fontFace: 'Calibri' } },
    { text: 'Δ unid', options: { bold: true, color: 'FFFFFF', fill: { color: '1E2761' }, fontSize: 9, align: 'right', fontFace: 'Calibri' } },
  ]];
  items.forEach((p, i) => {
    const bgMain = i % 2 === 0 ? 'F0F2F5' : 'FFFFFF';
    const bgSub = i % 2 === 0 ? 'FAFAFA' : 'F8F8F8';
    // Linha do PDV
    rows.push([
      { text: `${p[PIDX.NAME]}  ·  ${p[PIDX.CITY]} / ${p[PIDX.UF]}`,
        options: { color: '1E2761', fill: { color: bgMain }, fontSize: 9, bold: true, fontFace: 'Calibri' } },
      { text: fmtNum(p[PIDX.ABR]),
        options: { color: '595959', fill: { color: bgMain }, fontSize: 9, align: 'right', fontFace: 'Calibri' } },
      { text: fmtNum(p[PIDX.PROJ]),
        options: { color: '1E2761', fill: { color: bgMain }, fontSize: 9, bold: true, align: 'right', fontFace: 'Calibri' } },
      { text: fmtNum(p[PIDX.VAR]),
        options: { color: accentColor, fill: { color: bgMain }, fontSize: 9, bold: true, align: 'right', fontFace: 'Calibri' } },
    ]);
    // Sub-linha do produto que mais puxou
    const marca = p[PIDX.MARCA];
    const varMarca = p[PIDX.VAR_MARCA];
    if (marca && marca !== '—') {
      const marcaColor = varMarca >= 0 ? '2E7D32' : 'C62828';
      rows.push([
        { text: `   ↳ Produto: ${marca}`,
          options: { color: '595959', fill: { color: bgSub }, fontSize: 8.5, italic: true, fontFace: 'Calibri' } },
        { text: '', options: { fill: { color: bgSub }, fontFace: 'Calibri' } },
        { text: '', options: { fill: { color: bgSub }, fontFace: 'Calibri' } },
        { text: fmtNum(varMarca),
          options: { color: marcaColor, fill: { color: bgSub }, fontSize: 8.5, italic: true, bold: true, align: 'right', fontFace: 'Calibri' } },
      ]);
    }
  });
  s.addTable(rows, {
    x, y: y + 0.32, w, colW: [w * 0.55, w * 0.15, w * 0.15, w * 0.15],
    rowH: 0.25, fontSize: 9, fontFace: 'Calibri', border: { type: 'solid', color: 'E8E8E8', pt: 0.5 }
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
