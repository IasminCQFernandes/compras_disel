"""
Painel de Compras por Insumo (Diesel e outros) — LCM Construção
Streamlit + HTML/CSS/JS (ECharts)

Rodar:  streamlit run app.py

Fonte dos dados
  • Banco UAU (padrão): consulta vwHistoricoItensDeCompra via pyodbc / ODBC Driver 17.
    Configure .streamlit/secrets.toml (veja secrets.toml.exemplo).
  • Planilha (alternativa): se o banco não estiver configurado, usa ./diesel.xlsx
    ou um .xlsx enviado pela barra lateral (mesmas colunas da consulta).
"""
from __future__ import annotations

import io
import json
import math
import re
from datetime import date, datetime, timedelta
from pathlib import Path

import logging
import sys

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components


# No Windows, quando o navegador fecha/recarrega a aba, o asyncio registra
# "ConnectionResetError [WinError 10054]". É inofensivo — este filtro só esconde esse aviso.
class _IgnorarConexaoFechada(logging.Filter):
    def filter(self, record: logging.LogRecord) -> bool:
        exc = record.exc_info[1] if record.exc_info else None
        return not isinstance(exc, (ConnectionResetError, ConnectionAbortedError))


if sys.platform == "win32":
    _log_asyncio = logging.getLogger("asyncio")
    if not any(isinstance(f, _IgnorarConexaoFechada) for f in _log_asyncio.filters):
        _log_asyncio.addFilter(_IgnorarConexaoFechada())

BASE = Path(__file__).parent
ASSETS = BASE / "assets"
PLANILHA_PADRAO = BASE / "diesel.xlsx"

UFS = {
    "AC", "AL", "AP", "AM", "BA", "CE", "DF", "ES", "GO", "MA", "MT", "MS", "MG", "PA",
    "PB", "PR", "PE", "PI", "RJ", "RN", "RS", "RO", "RR", "SC", "SP", "SE", "TO",
}

st.set_page_config(page_title="Painel de Compras · LCM", page_icon="⛽", layout="wide",
                   initial_sidebar_state="collapsed")

st.markdown(
    """
    <style>
      #MainMenu, footer, header[data-testid="stHeader"] {visibility:hidden; height:0;}
      .block-container {padding:0 !important; max-width:100% !important;}
      iframe {display:block;}
      .st-key-barra, .st-key-barra_status, .st-key-barra_vazio {max-width:1480px; margin:18px auto 0; padding:0 24px; box-sizing:border-box;}
      .st-key-barra_status, .st-key-barra_vazio {margin-top:0 !important;}
      .st-key-barra [data-testid="stVerticalBlockBorderWrapper"]:first-child, .st-key-barra [data-testid="stForm"] {border:1px solid rgba(128,128,128,.25); border-radius:16px;
        padding:14px 16px 10px;}
      .st-key-barra button[kind="primary"], .st-key-barra button[kind="primaryFormSubmit"] {background:#2a78d6; border-color:#2a78d6; color:#fff;
        font-weight:700; height:42px; border-radius:10px;}
      .st-key-barra button[kind="primary"]:hover, .st-key-barra button[kind="primaryFormSubmit"]:hover {background:#1c5cab; border-color:#1c5cab;}
      .st-key-barra span[data-baseweb="tag"],
      .st-key-barra [data-testid="stMultiSelect"] [data-tag] {background:#2a78d6 !important; color:#fff !important;}
      .st-key-barra [data-testid="stMultiSelect"] [data-tag] * {color:#fff !important;}
      .st-key-barra [data-testid="stCaptionContainer"] {padding-left:4px;}
      .st-key-barra [data-testid="stTextInputRootElement"]:focus-within,
      .st-key-barra [data-baseweb="input"]:focus-within {border-color:#2a78d6 !important;}
      /* mesmo fundo do painel, para a barra e o painel parecerem uma página só */
      @media (prefers-color-scheme: light) { .stApp {background:#f2f2ef;} }
      @media (prefers-color-scheme: dark)  { .stApp {background:#111110;} }
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================================
# Banco UAU
# ============================================================================
def banco_configurado() -> bool:
    try:
        return "uau" in st.secrets
    except Exception:  # sem secrets.toml
        return False


def _conectar():
    import pyodbc

    cfg = st.secrets["uau"]
    driver = cfg.get("driver", "ODBC Driver 17 for SQL Server")
    partes = [f"DRIVER={{{driver}}}", f"SERVER={cfg['server']}", f"DATABASE={cfg['database']}"]
    if cfg.get("trusted_connection", False):
        partes.append("Trusted_Connection=yes")
    else:
        partes += [f"UID={cfg['username']}", f"PWD={cfg['password']}"]
    partes.append("TrustServerCertificate=yes")
    return pyodbc.connect(";".join(partes), timeout=20)


def _executar(sql: str, params: list, timeout: int = 0) -> pd.DataFrame:
    cn = _conectar()
    try:
        cn.timeout = timeout  # tempo máximo da consulta em segundos (0 = sem limite)
        cur = cn.cursor()
        cur.execute(sql, params)
        cols = [c[0] for c in cur.description]
        linhas = [tuple(r) for r in cur.fetchall()]
    finally:
        cn.close()
    return pd.DataFrame.from_records(linhas, columns=cols)


LIMITE_BUSCA = 300


def _like(palavra: str) -> str:
    """Escapa curingas do LIKE do SQL Server e envolve em %…%."""
    p = palavra.replace("[", "[[]").replace("%", "[%]").replace("_", "[_]")
    return f"%{p}%"


@st.cache_data(ttl=3600, show_spinner=False)
def buscar_insumos_banco(termo: str) -> pd.DataFrame:
    """Busca no cadastro de insumos (InsumosGeral) pelo código OU pela descrição.
    Várias palavras = todas precisam aparecer (ex.: "oleo diesel s10")."""
    palavras = [w for w in termo.split() if w][:6]
    filtros, params = [], []
    for w in palavras:
        filtros.append("(g.Cod_ins LIKE ? OR g.Descr_ins LIKE ?)")
        params += [_like(w), _like(w)]
    sql = f"""
SELECT DISTINCT TOP {LIMITE_BUSCA + 1}
       g.Cod_ins AS Cod, g.Descr_ins AS Descr, u.CodUn_UnIns AS Unid
FROM   InsumosGeral g
LEFT JOIN UnidInsumoGeral u ON g.Cod_ins = u.CodIns_UnIns AND u.Padrao_UnIns = 1
WHERE  g.AtInat_ins IN (0, 1) AND {" AND ".join(filtros)}
ORDER BY g.Cod_ins
"""
    df = _executar(sql, params, timeout=20)
    df["Cod"] = df["Cod"].astype(str).str.strip()
    df["Descr"] = df["Descr"].fillna("").astype(str).str.strip()
    df["Unid"] = df["Unid"].fillna("").astype(str).str.strip()
    return df


SQL_DIAG = """
SELECT COUNT(*) AS total,
       SUM(CASE WHEN Data_cot >= ? AND Data_cot < ? THEN 1 ELSE 0 END) AS no_periodo,
       SUM(CASE WHEN Data_cot >= ? AND Data_cot < ? AND Data_item IS NOT NULL AND MaxCotacao > 0 THEN 1 ELSE 0 END) AS validas,
       MIN(Data_cot) AS primeira, MAX(Data_cot) AS ultima
FROM vwHistoricoItensDeCompra
WHERE CodInsProc_Item IN ({marcadores})
"""


def diagnostico_banco(codigos: tuple[str, ...], ini: date, fim: date) -> dict:
    """Explica por que a consulta voltou vazia: código inexistente, fora do período ou barrado por permissão."""
    a = datetime.combine(ini, datetime.min.time())
    b = datetime.combine(fim + timedelta(days=1), datetime.min.time())
    d = _executar(SQL_DIAG.format(marcadores=", ".join("?" * len(codigos))),
                  [a, b, a, b, *codigos], timeout=60).iloc[0]
    out = {k: d[k] for k in ["total", "no_periodo", "validas", "primeira", "ultima"]}
    if not out["total"]:
        parecidos = set()
        for c in codigos:
            prefixo = re.split(r"[.\-\s]", c)[0][:4]
            if prefixo:
                r = _executar("SELECT DISTINCT TOP 15 CodInsProc_Item AS Cod FROM vwHistoricoItensDeCompra "
                              "WHERE CodInsProc_Item LIKE ? ORDER BY CodInsProc_Item", [prefixo + "%"], timeout=30)
                parecidos |= {str(x).strip() for x in r["Cod"]}
        out["parecidos"] = sorted(parecidos)[:15]
    return out


def testar_conexao() -> None:
    _executar("SELECT 1 AS ok", [], timeout=10)


def _sql_historico(qtd_insumos: int) -> str:
    marcadores = ", ".join("?" * qtd_insumos)
    return f"""
SELECT *
FROM (
    SELECT vwHistoricoItensDeCompra.*,
           CASE WHEN ISNULL(NumAc_Ocp, 0) = 0 THEN 0 ELSE 1 END AS OcpAcordoCorp
    FROM   vwHistoricoItensDeCompra
    LEFT JOIN ObrUsr
           ON ObrUsr.Emp_uo = vwHistoricoItensDeCompra.Codigo_Emp
          AND ObrUsr.Obr_uo = vwHistoricoItensDeCompra.Obra_Simc
          AND ObrUsr.Usr_uo = ?
    OUTER APPLY (SELECT Valor_cfg FROM ConfigGerais
                 WHERE Configuracao_cfg = 'WPermissaoObrasVisualizarProcHistPrecos') AS ConfigUtilizarPermissaoObra
    WHERE  CodInsProc_Item IN ({marcadores})
      AND  Data_item IS NOT NULL
      AND ((ConfigUtilizarPermissaoObra.Valor_cfg = 'Sim' AND ObrUsr.Obr_uo IS NOT NULL)
           OR ISNULL(ConfigUtilizarPermissaoObra.Valor_cfg, 'Nao') = 'Nao')
) AS TAB
WHERE TAB.MaxCotacao > 0
  AND TAB.Data_cot >= ?
  AND TAB.Data_cot <  ?
ORDER BY TAB.Data_cot DESC
"""


@st.cache_data(ttl=600, show_spinner=False)
def historico_do_banco(codigos: tuple[str, ...], ini: date, fim: date) -> pd.DataFrame:
    usuario = st.secrets["uau"].get("usuario_uau", "IASMIN")
    params = [usuario, *codigos, datetime.combine(ini, datetime.min.time()),
              datetime.combine(fim + timedelta(days=1), datetime.min.time())]
    return _executar(_sql_historico(len(codigos)), params, timeout=300)


# ============================================================================
# Planilha (alternativa sem banco)
# ============================================================================
@st.cache_data(show_spinner=False)
def ler_planilha(conteudo: bytes | None, caminho: str | None) -> pd.DataFrame:
    return pd.read_excel(io.BytesIO(conteudo) if conteudo is not None else caminho)


def filtrar_planilha(df: pd.DataFrame, codigos: tuple[str, ...], ini: date, fim: date) -> pd.DataFrame:
    d = pd.to_datetime(df["Data_cot"], errors="coerce")
    m = (df["CodInsProc_Item"].astype(str).str.strip().isin(codigos)
         & (d >= pd.Timestamp(ini)) & (d < pd.Timestamp(fim + timedelta(days=1))))
    if "MaxCotacao" in df:
        m &= pd.to_numeric(df["MaxCotacao"], errors="coerce").fillna(0) > 0
    return df[m]


# ============================================================================
# Tratamento
# ============================================================================
def _uf_do_texto(*textos) -> str | None:
    for t in textos:
        if isinstance(t, str):
            achados = [m for m in re.findall(r"/\s*([A-Z]{2})\b", t.upper()) if m in UFS]
            if achados:
                return achados[-1]
    return None


def _nome_curto_obra(nome) -> str:
    s = str(nome).strip()
    s = re.sub(r"^(SCP|CONSORCIO)\s+", "", s, flags=re.I)
    s = re.sub(r"\b100%\s+", "", s)
    s = re.sub(r"\s+-?\s*LCM\b.*$", "", s)
    return re.sub(r"\s{2,}", " ", s).strip(" -")


def _nome_curto_forn(nome) -> str:
    s = str(nome).strip()
    for _ in range(2):
        s = re.sub(r"[,.]?\s*\b(LTDA|S/?A|S\.A|EIRELI|ME|EPP)\.?$", "", s, flags=re.I)
    return s.strip(" ,.-")


def _unidade(u) -> str:
    u = str(u or "").strip().upper()
    return "L" if u in {"L", "LT", "LTS", "LITRO", "LITROS"} else (u or "UN")


def _escala(preco: pd.Series, qtd: pd.Series, subtotal: pd.Series) -> int:
    """O export em planilha às vezes perde a vírgula (6,85 vira 6850000). Detecta pela conta preço × qtd = subtotal."""
    m = (preco > 0) & (qtd > 0) & (subtotal > 0)
    if not m.any():
        return 1
    r = float((subtotal[m] / (preco[m] * qtd[m])).median())
    if r <= 0 or r > 0.5:
        return 1
    return int(10 ** round(math.log10(1 / r) / 2))


def _id(v):
    """Código vindo do UAU: número quando dá (ex.: 16929), senão texto (ex.: '007N3')."""
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return ""
    try:
        f = float(v)
        return int(f) if f.is_integer() else str(v).strip()
    except (TypeError, ValueError):
        return str(v).strip()


def _frete(v) -> str:
    try:
        return "FOB" if int(float(v)) == 1 else "CIF"
    except (TypeError, ValueError):
        return "CIF"


def preparar(df: pd.DataFrame, descr: dict[str, str]) -> tuple[list[dict], dict]:
    df = df.copy()
    preco = pd.to_numeric(df["Preco_Simc"], errors="coerce").astype(float)
    qtd = pd.to_numeric(df["QtdeSimu_Simc"], errors="coerce").astype(float)
    sub = pd.to_numeric(df["SubTotal"], errors="coerce").astype(float)
    esc = _escala(preco, qtd, sub)
    df["preco"], df["qtd"] = preco / esc, qtd / esc
    df["valor"] = sub.fillna(df["preco"] * df["qtd"])
    df["unid"] = df["Unid_temp"].map(_unidade)
    df["ins"] = df["CodInsProc_Item"].astype(str).str.strip()

    uf = df["UF_Obr"].where(df["UF_Obr"].isin(UFS))
    inferida = uf.isna()
    uf = uf.fillna(df.apply(lambda r: _uf_do_texto(r["Descr_Obr"], r["Desc_Emp"]), axis=1))
    df["uf_obra"] = uf.fillna(df["UF_Forn"])
    df["dt_cot"] = pd.to_datetime(df["Data_cot"], errors="coerce")
    df["dt_simc"] = pd.to_datetime(df["Data_Simc"], errors="coerce")

    # unidade principal = a que concentra mais valor (ex.: L). Outras (Vb, UN…) só entram no gasto.
    unidade = df.groupby("unid")["valor"].sum().idxmax()

    regs = []
    for r in df.itertuples(index=False):
        regs.append({
            "ins": r.ins,
            "proc": _id(r.Num_Proc), "cot": _id(r.MaxCotacao), "obraCod": str(_id(r.Obra_Simc)),
            "obra": _nome_curto_obra(r.Descr_Obr), "obraFull": str(r.Descr_Obr).strip(),
            "emp": str(r.Desc_Emp).strip(),
            "forn": _nome_curto_forn(r.nome_pes), "fornFull": str(r.nome_pes).strip(),
            "ufObra": r.uf_obra, "ufForn": r.UF_Forn if isinstance(r.UF_Forn, str) else None,
            "unid": r.unid,
            "preco": round(float(r.preco), 4), "qtd": round(float(r.qtd), 3), "valor": round(float(r.valor), 2),
            "dCot": r.dt_cot.strftime("%Y-%m-%d") if pd.notna(r.dt_cot) else None,
            "dSimc": r.dt_simc.strftime("%Y-%m-%d") if pd.notna(r.dt_simc) else None,
            "frete": _frete(r.tipoFrete),
        })

    gasto_ins = df.groupby("ins")["valor"].sum()
    insumos = [{"cod": c, "desc": descr.get(c, "") or "", "valor": float(gasto_ins.get(c, 0))}
               for c in sorted(df["ins"].unique())]
    meta = {
        "linhas": len(df), "ufInferidas": int(inferida.sum()),
        "outrasUnid": int((df["unid"] != unidade).sum()), "unidade": unidade,
        "escala": esc, "insumos": insumos,
    }
    return regs, meta


# ============================================================================
# Excel (.xlsx) com os dados da consulta
# ============================================================================
def _para_excel(v):
    """Decimal → float, datas com fuso → sem fuso (o Excel não aceita fuso)."""
    from decimal import Decimal
    if isinstance(v, Decimal):
        return float(v)
    if isinstance(v, pd.Timestamp) and v.tzinfo is not None:
        return v.tz_localize(None)
    return v


@st.cache_data(show_spinner=False, max_entries=5)
def gerar_xlsx(df_bruto: pd.DataFrame, regs: list[dict], meta: dict, ini: date, fim: date) -> bytes:
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    un = meta.get("unidade", "L")
    desc = {i["cod"]: i.get("desc", "") for i in meta.get("insumos", [])}
    r = pd.DataFrame(regs)
    compras = pd.DataFrame({
        "Insumo": r["ins"], "Descrição do insumo": r["ins"].map(desc).fillna(""),
        "Data cotação": pd.to_datetime(r["dCot"]), "Data simulação": pd.to_datetime(r["dSimc"]),
        "Processo": r["proc"], "Cotação": r["cot"], "Cód. obra": r["obraCod"], "Obra": r["obraFull"],
        "Empresa": r["emp"], "UF obra": r["ufObra"], "Fornecedor": r["fornFull"], "UF fornecedor": r["ufForn"],
        "Frete": r["frete"], "Unidade": r["unid"], "Quantidade": r["qtd"], "Preço unitário": r["preco"],
        "Total": r["valor"],
    }).sort_values("Data cotação", ascending=False)

    def resumo(chave: str, nome: str) -> pd.DataFrame:
        vol = r[r["unid"] == un]
        g = r.groupby(chave).agg(Gasto=("valor", "sum"), Compras=("valor", "size"))
        gv = vol.groupby(chave).agg(Volume=("qtd", "sum"), _v=("valor", "sum"))
        g = g.join(gv, how="left")
        g[f"Preço médio (R$/{un})"] = g["_v"] / g["Volume"]
        g["Participação"] = g["Gasto"] / g["Gasto"].sum()
        g = g.drop(columns="_v").rename(columns={"Volume": f"Volume ({un})"})
        return g.sort_values("Gasto", ascending=False).reset_index().rename(columns={chave: nome})

    abas = {"Compras": compras, "Resumo por UF": resumo("ufObra", "UF da obra")}
    if len(meta.get("insumos", [])) > 1:
        ri = resumo("ins", "Insumo")
        ri.insert(1, "Descrição", ri["Insumo"].map(desc).fillna(""))
        abas["Resumo por insumo"] = ri
    ro = resumo("obraFull", "Obra")
    abas["Resumo por obra"] = ro
    rf = resumo("fornFull", "Fornecedor")
    abas["Resumo por fornecedor"] = rf
    abas["Consulta UAU (bruta)"] = df_bruto.map(_para_excel) if hasattr(df_bruto, "map") else df_bruto.applymap(_para_excel)

    formatos = {"Data cotação": "DD/MM/YYYY", "Data simulação": "DD/MM/YYYY", "Quantidade": "#,##0.000",
                "Preço unitário": "#,##0.0000", "Total": '"R$" #,##0.00', "Gasto": '"R$" #,##0.00',
                f"Volume ({un})": "#,##0.000", f"Preço médio (R$/{un})": '"R$" #,##0.0000',
                "Participação": "0.0%", "Compras": "0"}
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        info = pd.DataFrame({"Campo": ["Relatório", "Insumos", "Período (data da cotação)", "Registros", "Gerado em"],
                             "Valor": ["Painel de Compras por Insumo — LCM Construção",
                                       ", ".join(f"{c} · {d}" if d else c for c, d in desc.items()),
                                       f"{ini:%d/%m/%Y} a {fim:%d/%m/%Y}", len(r), f"{datetime.now():%d/%m/%Y %H:%M}"]})
        info.to_excel(xw, sheet_name="Sobre", index=False)
        for nome, d in abas.items():
            d.to_excel(xw, sheet_name=nome[:31], index=False)
        cab_fill, cab_font = PatternFill("solid", fgColor="0F1B2E"), Font(bold=True, color="FFFFFF")
        for ws in xw.book.worksheets:
            ws.freeze_panes = "A2"
            if ws.max_row > 1:
                ws.auto_filter.ref = ws.dimensions
            for col in range(1, ws.max_column + 1):
                cab = ws.cell(row=1, column=col)
                cab.fill, cab.font = cab_fill, cab_font
                cab.alignment = Alignment(vertical="center")
                fmt = formatos.get(str(cab.value))
                largura = len(str(cab.value or ""))
                for linha in range(2, min(ws.max_row, 400) + 1):
                    c = ws.cell(row=linha, column=col)
                    largura = max(largura, len(str(c.value)) if c.value is not None else 0)
                if fmt:
                    for linha in range(2, ws.max_row + 1):
                        ws.cell(row=linha, column=col).number_format = fmt
                ws.column_dimensions[get_column_letter(col)].width = min(max(10, largura + 2), 60)
    return buf.getvalue()


def montar_html(regs: list[dict], meta: dict, xlsx: bytes | None = None) -> str:
    import base64

    html = (ASSETS / "painel.html").read_text(encoding="utf-8")
    dados = json.dumps({"regs": regs, "meta": meta}, ensure_ascii=False, separators=(",", ":"))
    html = html.replace("/*__DADOS__*/null", dados.replace("</", "<\\/"))
    html = html.replace("/*__XLSX__*/null", json.dumps(base64.b64encode(xlsx).decode()) if xlsx else "null")
    html = html.replace("/*__GEO__*/null", (ASSETS / "brasil_uf.json").read_text(encoding="utf-8"))
    libs_pdf = "\n".join((ASSETS / f).read_text(encoding="utf-8") for f in ("html-to-image.js", "jspdf.umd.min.js"))
    html = html.replace("/*__PDFLIBS__*/", libs_pdf)
    return html.replace("/*__ECHARTS__*/", (ASSETS / "echarts.min.js").read_text(encoding="utf-8"))


# ============================================================================
# Página
# ============================================================================
with st.sidebar:
    st.markdown("### Fonte dos dados")
    arq = st.file_uploader("Usar uma planilha (.xlsx) em vez do banco", type=["xlsx"])
    st.caption("Mesmas colunas da consulta vwHistoricoItensDeCompra.")

usar_banco = banco_configurado() and arq is None
if usar_banco:
    if not st.session_state.get("conexao_ok"):
        with st.spinner("Conectando ao UAU…"):
            try:
                testar_conexao()
                st.session_state.conexao_ok = True
            except Exception as e:  # noqa: BLE001
                st.error(f"**Não consegui conectar ao banco UAU.**\n\n{e}\n\n"
                         "Confira servidor, banco, usuário e senha em `.streamlit/secrets.toml` e se o "
                         "ODBC Driver 17 está instalado. Enquanto isso, dá para usar uma planilha pela barra lateral.")
                st.stop()
    df_planilha = None
else:
    if arq is not None:
        df_planilha = ler_planilha(arq.getvalue(), None)
    elif PLANILHA_PADRAO.exists():
        df_planilha = ler_planilha(None, str(PLANILHA_PADRAO))
    else:
        st.warning("Configure o banco em **.streamlit/secrets.toml** ou coloque **diesel.xlsx** ao lado do app.py.")
        st.stop()


def buscar_insumos(termo: str) -> pd.DataFrame:
    if usar_banco:
        return buscar_insumos_banco(termo)
    cods = pd.Series(sorted(df_planilha["CodInsProc_Item"].astype(str).str.strip().unique()))
    m = pd.Series(True, index=cods.index)
    for w in termo.upper().split():
        m &= cods.str.upper().str.contains(w, regex=False)
    return pd.DataFrame({"Cod": cods[m], "Descr": "", "Unid": ""}).head(LIMITE_BUSCA + 1)


hoje = date.today()
ini_padrao = (hoje.replace(day=1) - pd.DateOffset(months=6)).date()
# Nada é consultado ao abrir: o painel só carrega depois que o usuário escolhe e clica em "Carregar".
st.session_state.setdefault("params", None)
st.session_state.setdefault("ins_sel", {})     # {código: descrição} marcados
st.session_state.setdefault("ins_ver", 0)      # muda para "resetar" a grade de marcação
params = st.session_state.params


def _rotulo(c: str) -> str:
    d = st.session_state.ins_sel.get(c, "")
    return f"{c} · {d}" if d else c


def _marcar(df: pd.DataFrame, marcar: bool) -> None:
    for r in df.itertuples():
        if marcar:
            st.session_state.ins_sel[r.Cod] = r.Descr
        else:
            st.session_state.ins_sel.pop(r.Cod, None)
    st.session_state.ins_ver += 1


def _limpar_sel() -> None:
    st.session_state.ins_sel = {}
    st.session_state.ins_ver += 1


@st.fragment
def barra_de_escolha() -> None:
    """Busca ao digitar + marcação. Roda como fragmento: digitar não recarrega o painel."""
    sel: dict = st.session_state.ins_sel
    c1, c2, c3 = st.columns([4.2, 2.2, 1], vertical_alignment="bottom")
    kw = dict(placeholder="Digite parte do código ou da descrição — ex.: diesel, COMBU, óleo s10",
              key="termo_busca")
    try:
        termo = c1.text_input("Buscar insumo (código ou descrição)", type="search", live="400ms", **kw)
    except TypeError:  # Streamlit sem busca ao digitar: pesquisa ao teclar Enter
        termo = c1.text_input("Buscar insumo (código ou descrição) — tecle Enter", **kw)
    per = c2.date_input("Período (data da cotação)", value=params[1:] if params else (ini_padrao, hoje),
                        format="DD/MM/YYYY", max_value=hoje + timedelta(days=366), key="periodo")
    carregar = c3.button("Carregar", type="primary", use_container_width=True, key="btn_carregar")

    termo = (termo or "").strip()
    if len(termo) >= 2:
        try:
            res = buscar_insumos(termo)
        except Exception as e:  # noqa: BLE001
            st.error(f"Erro ao buscar insumos: {e}")
            res = pd.DataFrame(columns=["Cod", "Descr", "Unid"])
        cortou = len(res) > LIMITE_BUSCA
        res = res.head(LIMITE_BUSCA).reset_index(drop=True)

        if res.empty:
            st.caption(f"Nenhum insumo encontrado para “{termo}”.")
        else:
            marcados = int(res["Cod"].isin(sel).sum())
            h1, h2, h3 = st.columns([4.2, 1.6, 1.6], vertical_alignment="center")
            h1.caption(f"**{len(res)}** insumo(s) encontrado(s) para “{termo}”"
                       + (f" — mostrando os {LIMITE_BUSCA} primeiros, refine a busca" if cortou else "")
                       + f" · {marcados} marcado(s)")
            h2.button(f"☑ Marcar todos ({len(res)})", on_click=_marcar, args=(res, True),
                      use_container_width=True, key="btn_marcar_todos", disabled=marcados == len(res))
            h3.button("☐ Desmarcar estes", on_click=_marcar, args=(res, False),
                      use_container_width=True, key="btn_desmarcar", disabled=marcados == 0)

            grade = res.assign(Marcar=res["Cod"].isin(sel))[["Marcar", "Cod", "Descr", "Unid"]]
            editado = st.data_editor(
                grade, hide_index=True, use_container_width=True,
                height=min(38 + 35 * len(grade), 290),
                key=f"grade_{abs(hash(termo))}_{st.session_state.ins_ver}",
                disabled=["Cod", "Descr", "Unid"],
                column_config={
                    "Marcar": st.column_config.CheckboxColumn("", width="small"),
                    "Cod": st.column_config.TextColumn("Código", width="small"),
                    "Descr": st.column_config.TextColumn("Descrição", width="large"),
                    "Unid": st.column_config.TextColumn("Unid.", width="small"),
                })
            for r in editado.itertuples():
                if r.Marcar:
                    sel[r.Cod] = r.Descr
                else:
                    sel.pop(r.Cod, None)

    # marcados (chips removíveis)
    if sel:
        s1, s2 = st.columns([6.4, 1], vertical_alignment="bottom")
        ficou = s1.multiselect(f"Insumos marcados ({len(sel)})", options=list(sel), default=list(sel),
                               format_func=_rotulo, key=f"chips_{st.session_state.ins_ver}_{len(sel)}")
        s2.button("Limpar", on_click=_limpar_sel, use_container_width=True, key="btn_limpar")
        if len(ficou) < len(sel):
            for c in [c for c in sel if c not in ficou]:
                sel.pop(c, None)
            st.session_state.ins_ver += 1
            st.rerun(scope="fragment")
    elif not termo:
        st.caption("Busque e marque um ou mais insumos, escolha o período e clique em **Carregar**.")

    if carregar:
        if not sel:
            st.warning("Marque pelo menos um insumo.")
        elif not isinstance(per, (tuple, list)) or len(per) != 2:
            st.warning("Escolha a data inicial **e** a final do período.")
        else:
            st.session_state.params = (tuple(sorted(sel)), per[0], per[1])
            st.session_state.ins_desc = dict(sel)
            st.rerun()  # recarrega o app inteiro com o painel


with st.container(key="barra"):
    with st.container(border=True):
        barra_de_escolha()

descricoes = st.session_state.get("ins_desc", {})
fonte = "Banco UAU" if usar_banco else (arq.name if arq is not None else PLANILHA_PADRAO.name)


def mostrar_html(html: str, altura_fixa: int) -> None:
    if hasattr(st, "iframe"):
        st.iframe(html, height="content")  # Streamlit novo: ajusta a altura sozinho
    else:
        components.html(html, height=altura_fixa, scrolling=True)


if params is None:
    tutorial = (ASSETS / "tutorial.html").read_text(encoding="utf-8").replace("{{FONTE}}", fonte)
    mostrar_html(tutorial, 1100)
    st.stop()

codigos, ini, fim = params
with st.container(key="barra_status"):
    try:
        with st.spinner("Consultando o histórico de compras no UAU… (pode levar alguns segundos)" if usar_banco else "Lendo a planilha…"):
            df = historico_do_banco(codigos, ini, fim) if usar_banco else filtrar_planilha(df_planilha, codigos, ini, fim)
    except Exception as e:  # noqa: BLE001
        st.error(f"Erro ao consultar o banco: {e}")
        st.stop()
    st.caption(f"Fonte: **{fonte}** · {len(df)} registros · {', '.join(codigos)} · "
               f"{ini:%d/%m/%Y} a {fim:%d/%m/%Y} · atualizado às {datetime.now():%H:%M}")

def _usar_codigo(cod: str, ini_: date, fim_: date) -> None:
    st.session_state.params = ((cod,), ini_, fim_)
    st.session_state.ins_sel = {cod: ""}
    st.session_state.ins_desc = {cod: ""}
    st.session_state.ins_ver += 1


if df.empty:
    with st.container(key="barra_vazio"):
        st.info("Nenhuma compra encontrada para esses insumos no período.")
        try:
            with st.spinner("Verificando o motivo…"):
                if usar_banco:
                    dg = diagnostico_banco(codigos, ini, fim)
                else:  # planilha
                    base = df_planilha[df_planilha["CodInsProc_Item"].astype(str).str.strip().isin(codigos)]
                    datas = pd.to_datetime(base["Data_cot"], errors="coerce")
                    no_per = ((datas >= pd.Timestamp(ini)) & (datas < pd.Timestamp(fim + timedelta(days=1)))).sum()
                    dg = {"total": len(base), "no_periodo": no_per, "validas": no_per,
                          "primeira": datas.min(), "ultima": datas.max()}
                    if not len(base):
                        pref = tuple(re.split(r"[.\-\s]", c)[0][:4] for c in codigos)
                        todos = df_planilha["CodInsProc_Item"].astype(str).str.strip().unique()
                        dg["parecidos"] = sorted(c for c in todos if c.startswith(pref))[:15]
        except Exception as e:  # noqa: BLE001
            st.caption(f"Não consegui verificar o motivo: {e}")
            st.stop()

        cods = ", ".join(f"**{c}**" for c in codigos)
        fmt = lambda d: pd.Timestamp(d).strftime("%d/%m/%Y") if pd.notna(d) else "—"  # noqa: E731
        if not dg["total"]:
            st.warning(f"O código {cods} não existe no histórico de compras do UAU. Confira a digitação "
                       "(ex.: **COMBU.0002**, com U).")
            if dg.get("parecidos"):
                st.markdown("**Códigos parecidos encontrados** — clique para carregar:")
                cols = st.columns(5)
                for i, c in enumerate(dg["parecidos"]):
                    cols[i % len(cols)].button(c, key=f"usar_{c}", on_click=_usar_codigo, args=(c, ini, fim),
                                               use_container_width=True)
        elif not dg["no_periodo"]:
            st.warning(f"{cods} tem {int(dg['total'])} compras no UAU, mas nenhuma cotação entre "
                       f"{ini:%d/%m/%Y} e {fim:%d/%m/%Y}. Há cotações de **{fmt(dg['primeira'])}** "
                       f"a **{fmt(dg['ultima'])}** — ajuste o período.")
        elif not dg["validas"]:
            st.warning(f"Há {int(dg['no_periodo'])} registros de {cods} no período, mas nenhum com cotação "
                       "finalizada (Data_item preenchida e MaxCotacao > 0), que são os que a consulta considera.")
        else:
            usuario = st.secrets["uau"].get("usuario_uau", "IASMIN") if usar_banco else ""
            st.warning(f"Há {int(dg['validas'])} compras válidas de {cods} no período, mas todas são de obras que o "
                       f"usuário **{usuario}** não tem permissão de ver no UAU (regra WPermissaoObrasVisualizarProcHistPrecos "
                       "/ ObrUsr). Libere as obras para esse usuário no UAU ou troque `usuario_uau` no secrets.toml.")
    st.stop()

regs, meta = preparar(df, descricoes)
meta.update({"fonte": fonte, "periodo": {"ini": ini.isoformat(), "fim": fim.isoformat()}})

# altura fixa só é usada no Streamlit antigo (sem st.iframe)
nome_arq = "compras_" + "_".join(codigos)[:60].replace(".", "") + f"_{ini:%Y%m%d}-{fim:%Y%m%d}"
meta["arquivo"] = re.sub(r"[^\w-]+", "_", nome_arq)
try:
    xlsx = gerar_xlsx(df, regs, meta, ini, fim)
except Exception as e:  # noqa: BLE001  — o painel abre mesmo se o Excel falhar
    xlsx = None
    st.caption(f"⚠ Não foi possível gerar o Excel: {e}")
mostrar_html(montar_html(regs, meta, xlsx), 4400 + (760 if len(meta["insumos"]) > 1 else 0))
