import streamlit as st
import pandas as pd
import io
import chardet
from collections import Counter

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Stack Files",
    page_icon="📐",
    layout="centered",
)

# ── Custom CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Space+Mono:wght@400;700&family=DM+Sans:wght@300;400;500&display=swap');

html, body, [class*="css"] {
    font-family: 'DM Sans', sans-serif;
    background-color: #f7f7f5;
    color: #1a1a1a;
}

.stApp { background-color: #f7f7f5; }

h1, h2, h3 { font-family: 'Space Mono', monospace !important; }

.hero {
    border: 1px solid #d8d8d0;
    border-left: 4px solid #2d7a2d;
    background: #ffffff;
    padding: 2rem 2.5rem;
    margin-bottom: 2rem;
    border-radius: 2px;
}
.hero h1 {
    font-size: 2rem;
    letter-spacing: -0.03em;
    color: #2d7a2d;
    margin: 0 0 0.25rem 0;
}
.hero p { color: #777; font-size: 0.95rem; margin: 0; }

.file-card {
    background: #ffffff;
    border: 1px solid #d8d8d0;
    border-radius: 2px;
    padding: 0.85rem 1.2rem;
    margin: 0.4rem 0;
    display: flex;
    align-items: center;
    gap: 0.75rem;
    font-family: 'Space Mono', monospace;
    font-size: 0.8rem;
    color: #1a1a1a;
}
.file-card.ok    { border-left: 3px solid #2d7a2d; }
.file-card.error { border-left: 3px solid #cc2222; }

.badge {
    font-family: 'Space Mono', monospace;
    font-size: 0.7rem;
    padding: 0.15rem 0.5rem;
    border-radius: 2px;
    font-weight: 700;
    letter-spacing: 0.05em;
}
.badge-ok    { background: #2d7a2d18; color: #2d7a2d; }
.badge-error { background: #cc222218; color: #cc2222; }
.badge-warn  { background: #bb660018; color: #bb6600; }

.stat-row { display: flex; gap: 1.5rem; margin: 1.5rem 0; }
.stat-box {
    flex: 1;
    background: #ffffff;
    border: 1px solid #d8d8d0;
    padding: 1rem 1.5rem;
    text-align: center;
    border-radius: 2px;
}
.stat-box .num {
    font-family: 'Space Mono', monospace;
    font-size: 2rem;
    color: #2d7a2d;
    display: block;
    font-weight: 700;
}
.stat-box .lbl {
    font-size: 0.75rem;
    color: #999;
    letter-spacing: 0.08em;
    text-transform: uppercase;
}

.section-title {
    font-family: 'Space Mono', monospace;
    font-size: 0.75rem;
    letter-spacing: 0.15em;
    text-transform: uppercase;
    color: #999;
    margin: 1.5rem 0 0.75rem 0;
    padding-bottom: 0.4rem;
    border-bottom: 1px solid #e4e4dc;
}

/* Total geral */
.total-grand {
    background: #2d7a2d;
    color: #ffffff;
    border-radius: 2px;
    padding: 1.25rem 2rem;
    margin: 1rem 0 0.5rem 0;
    display: flex;
    justify-content: space-between;
    align-items: center;
}
.total-grand .tg-label {
    font-family: 'Space Mono', monospace;
    font-size: 0.8rem;
    letter-spacing: 0.12em;
    text-transform: uppercase;
    opacity: 0.85;
}
.total-grand .tg-value {
    font-family: 'Space Mono', monospace;
    font-size: 1.6rem;
    font-weight: 700;
}

/* Cards por moeda */
.currency-row { display: flex; flex-wrap: wrap; gap: 0.75rem; margin: 0.5rem 0 1.5rem 0; }
.currency-card {
    flex: 1;
    min-width: 160px;
    background: #ffffff;
    border: 1px solid #d8d8d0;
    border-top: 3px solid #2d7a2d;
    border-radius: 2px;
    padding: 0.9rem 1.2rem;
    text-align: center;
}
.currency-card .cc-label {
    font-family: 'Space Mono', monospace;
    font-size: 0.7rem;
    letter-spacing: 0.12em;
    text-transform: uppercase;
    color: #999;
    margin-bottom: 0.3rem;
}
.currency-card .cc-value {
    font-family: 'Space Mono', monospace;
    font-size: 1.15rem;
    font-weight: 700;
    color: #1a1a1a;
}
.currency-card .cc-rows {
    font-size: 0.72rem;
    color: #aaa;
    margin-top: 0.2rem;
}

div[data-testid="stFileUploader"] {
    border: 2px dashed #d8d8d0 !important;
    border-radius: 2px !important;
    background: #ffffff !important;
    padding: 2rem !important;
    transition: border-color 0.2s;
}
div[data-testid="stFileUploader"]:hover { border-color: #2d7a2d !important; }

.stButton > button {
    font-family: 'Space Mono', monospace !important;
    font-size: 0.85rem !important;
    letter-spacing: 0.05em !important;
    background: #2d7a2d !important;
    color: #ffffff !important;
    border: none !important;
    border-radius: 2px !important;
    padding: 0.65rem 2rem !important;
    font-weight: 700 !important;
    transition: opacity 0.2s !important;
    width: 100% !important;
}
.stButton > button:hover { opacity: 0.85 !important; }

.stDownloadButton > button {
    font-family: 'Space Mono', monospace !important;
    font-size: 0.85rem !important;
    background: transparent !important;
    color: #2d7a2d !important;
    border: 1px solid #2d7a2d !important;
    border-radius: 2px !important;
    padding: 0.65rem 2rem !important;
    width: 100% !important;
    transition: all 0.2s !important;
}
.stDownloadButton > button:hover {
    background: #2d7a2d !important;
    color: #ffffff !important;
}

.col-list {
    font-family: 'Space Mono', monospace;
    font-size: 0.75rem;
    background: #f0f0ec;
    border: 1px solid #d8d8d0;
    padding: 1rem 1.2rem;
    border-radius: 2px;
    line-height: 2;
    color: #666;
}
.col-missing { color: #cc2222; }
.col-extra   { color: #bb6600; }

.stDataFrame { border: 1px solid #d8d8d0 !important; }

#MainMenu, footer, header { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

# ── Helpers ───────────────────────────────────────────────────────────────────

SEPARATORS = [",", ";", "\t", "|", ":", " "]

CURRENCY_SYMBOLS = {
    "BRL — R$":  "R$",
    "USD — $":   "$",
    "EUR — €":   "€",
    "GBP — £":   "£",
    "ARS — $":   "AR$",
    "CLP — $":   "CL$",
    "MXN — $":   "MX$",
    "Nenhuma":   "",
}


def detect_encoding(raw: bytes) -> str:
    result = chardet.detect(raw)
    enc = result.get("encoding") or "utf-8"
    if enc.lower().startswith("utf"):
        return "utf-8-sig"
    return enc


def sniff_separator(raw: bytes, encoding: str) -> str:
    sample = raw[:4096].decode(encoding, errors="replace")
    counts = {sep: sample.count(sep) for sep in SEPARATORS}
    best = max(counts, key=counts.get)
    return best if counts[best] > 0 else ","


def read_file(uploaded_file):
    name = uploaded_file.name.lower()
    raw  = uploaded_file.read()
    uploaded_file.seek(0)
    try:
        if name.endswith(".csv"):
            enc = detect_encoding(raw)
            sep = sniff_separator(raw, enc)
            df  = pd.read_csv(io.BytesIO(raw), sep=sep, encoding=enc,
                              engine="python", on_bad_lines="warn")
            return df, ""
        elif name.endswith(".xlsx"):
            return pd.read_excel(io.BytesIO(raw), engine="openpyxl"), ""
        elif name.endswith(".xls"):
            return pd.read_excel(io.BytesIO(raw), engine="xlrd"), ""
        else:
            return None, "Formato não suportado."
    except Exception as exc:
        return None, str(exc)


def majority_columns(frames_cols):
    col_tuples = [tuple(sorted(cols)) for _, cols in frames_cols]
    most_common, _ = Counter(col_tuples).most_common(1)[0]
    return list(most_common)


def fmt_number(value: float, symbol: str) -> str:
    """Format a number as currency with thousand separators (pt-BR style)."""
    formatted = f"{value:_.2f}".replace("_", ".").replace(".", ",", 1)
    # fix: replace only the first dot as thousands sep → actually use a cleaner approach
    formatted = f"{value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    if symbol:
        return f"{symbol} {formatted}"
    return formatted


def build_currency_summary_html(df: pd.DataFrame, value_col: str,
                                  currency_col: str | None, symbol: str) -> str:
    """Return HTML for per-currency cards + grand total block."""
    grand_total = pd.to_numeric(df[value_col], errors="coerce").sum()

    # ── Grand total ──
    html = f"""
    <div class="total-grand">
        <span class="tg-label">Total Geral — {value_col}</span>
        <span class="tg-value">{fmt_number(grand_total, symbol)}</span>
    </div>"""

    # ── Per-currency breakdown ──
    if currency_col:
        groups = (
            df.assign(_val=pd.to_numeric(df[value_col], errors="coerce"))
              .groupby(currency_col, dropna=False)
              .agg(total=("_val", "sum"), rows=("_val", "count"))
              .reset_index()
              .sort_values("total", ascending=False)
        )

        cards = ""
        for _, row in groups.iterrows():
            cur_label = str(row[currency_col]) if pd.notna(row[currency_col]) else "—"
            cards += f"""
            <div class="currency-card">
                <div class="cc-label">{cur_label}</div>
                <div class="cc-value">{fmt_number(row['total'], symbol)}</div>
                <div class="cc-rows">{int(row['rows'])} linha(s)</div>
            </div>"""

        html += f'<div class="currency-row">{cards}</div>'

    return html


# ── UI ────────────────────────────────────────────────────────────────────────

st.markdown("""
<div class="hero">
    <h1>STACK FILES</h1>
    <p>Empilha arquivos CSV / XLS / XLSX com as mesmas colunas → exporta CSV padrão US</p>
</div>
""", unsafe_allow_html=True)

uploaded = st.file_uploader(
    "Arraste ou selecione os arquivos",
    type=["csv", "xls", "xlsx"],
    accept_multiple_files=True,
    label_visibility="collapsed",
)

if not uploaded:
    st.markdown("""
    <div style="color:#aaa;font-family:'Space Mono',monospace;font-size:0.75rem;
                text-align:center;padding:1rem;letter-spacing:0.08em;">
    ↑ &nbsp; ARRASTE QUALQUER COMBINAÇÃO DE CSV · XLS · XLSX
    </div>""", unsafe_allow_html=True)
    st.stop()

# ── Process files ─────────────────────────────────────────────────────────────

st.markdown('<div class="section-title">Leitura dos arquivos</div>', unsafe_allow_html=True)

frames_cols = []
frames_df   = []
read_errors = {}

for f in uploaded:
    df, err = read_file(f)
    if err:
        read_errors[f.name] = err
    else:
        frames_cols.append((f.name, list(df.columns)))
        frames_df.append((f.name, df))

    cls       = "error" if err else "ok"
    badge_cls = "badge-error" if err else "badge-ok"
    badge_txt = "ERRO" if err else f"{len(df)} linhas"
    icon      = "❌" if err else "✓"

    st.markdown(f"""
    <div class="file-card {cls}">
        <span>{icon}</span>
        <span style="flex:1">{f.name}</span>
        <span class="badge {badge_cls}">{badge_txt}</span>
    </div>""", unsafe_allow_html=True)

if read_errors and not frames_cols:
    st.error("Nenhum arquivo foi lido com sucesso.")
    st.stop()

# ── Column validation ─────────────────────────────────────────────────────────

st.markdown('<div class="section-title">Validação de colunas</div>', unsafe_allow_html=True)

ref_cols  = majority_columns(frames_cols)
bad_files = {}

for fname, cols in frames_cols:
    missing = [c for c in ref_cols if c not in cols]
    extra   = [c for c in cols if c not in ref_cols]
    if missing or extra:
        bad_files[fname] = {"missing": missing, "extra": extra}

if not bad_files:
    st.markdown("""
    <div class="file-card ok" style="justify-content:center">
        ✓ &nbsp; Todas as colunas são consistentes entre os arquivos
    </div>""", unsafe_allow_html=True)
else:
    n_ok  = len(frames_cols) - len(bad_files)
    n_bad = len(bad_files)
    st.markdown(f"""
    <div class="file-card error">
        ⚠ &nbsp; <strong>{n_bad} arquivo(s) com divergência</strong>
        &nbsp;·&nbsp; Referência: maioria ({n_ok} arquivo(s))
    </div>""", unsafe_allow_html=True)

    for fname, diff in bad_files.items():
        parts = []
        if diff["missing"]:
            cols_str = " &nbsp;·&nbsp; ".join(
                f'<span class="col-missing">{c}</span>' for c in diff["missing"])
            parts.append(f"<strong>Faltando:</strong> {cols_str}")
        if diff["extra"]:
            cols_str = " &nbsp;·&nbsp; ".join(
                f'<span class="col-extra">{c}</span>' for c in diff["extra"])
            parts.append(f"<strong>A mais:</strong> {cols_str}")

        st.markdown(f"""
        <div style="margin:0.5rem 0 0.5rem 1.5rem;">
            <div style="font-family:'Space Mono',monospace;font-size:0.78rem;
                        color:#ff4d4d;margin-bottom:0.25rem;">{fname}</div>
            <div class="col-list">{" &nbsp;&nbsp; ".join(parts)}</div>
        </div>""", unsafe_allow_html=True)

# ── Stats ─────────────────────────────────────────────────────────────────────

good_frames = [(n, df) for n, df in frames_df if n not in bad_files and n not in read_errors]
total_rows  = sum(len(df) for _, df in good_frames)

st.markdown(f"""
<div class="stat-row">
    <div class="stat-box"><span class="num">{len(uploaded)}</span><span class="lbl">Arquivos enviados</span></div>
    <div class="stat-box"><span class="num">{len(good_frames)}</span><span class="lbl">Prontos para empilhar</span></div>
    <div class="stat-box"><span class="num">{total_rows:,}</span><span class="lbl">Linhas totais</span></div>
    <div class="stat-box"><span class="num">{len(ref_cols)}</span><span class="lbl">Colunas</span></div>
</div>""", unsafe_allow_html=True)

with st.expander("Ver colunas de referência"):
    st.markdown(
        '<div class="col-list">' +
        " &nbsp;·&nbsp; ".join(f"<span style='color:#2d7a2d'>{c}</span>" for c in ref_cols) +
        "</div>",
        unsafe_allow_html=True,
    )

# ── Exportar ──────────────────────────────────────────────────────────────────

st.markdown('<div class="section-title">Exportar</div>', unsafe_allow_html=True)

can_merge = len(good_frames) >= 1

if not can_merge:
    st.markdown("""
    <div class="file-card error" style="justify-content:center">
        Nenhum arquivo válido disponível para concatenar.
    </div>""", unsafe_allow_html=True)
    st.stop()

if bad_files:
    st.markdown(f"""
    <div class="file-card" style="border-left:3px solid #ffaa22;margin-bottom:1rem">
        <span class="badge badge-warn">AVISO</span>
        &nbsp; {len(bad_files)} arquivo(s) ignorados por divergência.
        Apenas {len(good_frames)} arquivo(s) válidos serão concatenados.
    </div>""", unsafe_allow_html=True)

# ── Seletores de resumo ───────────────────────────────────────────────────────

# Detecta colunas numéricas a partir dos dados já carregados
sample_df = pd.concat([df for _, df in good_frames], ignore_index=True)
numeric_cols = [c for c in ref_cols if pd.to_numeric(sample_df[c], errors="coerce").notna().any()]
text_cols    = [c for c in ref_cols if c not in numeric_cols]

col_a, col_b, col_c = st.columns(3)

with col_a:
    value_col = st.selectbox(
        "Coluna para somar",
        options=["— nenhuma —"] + numeric_cols,
        help="Selecione a coluna numérica cujo total deseja exibir",
    )

with col_b:
    currency_col = st.selectbox(
        "Coluna de moeda",
        options=["— nenhuma —"] + text_cols,
        help="Selecione a coluna que identifica a moeda (ex: BRL, USD…). O resumo será quebrado por moeda automaticamente.",
    )

with col_c:
    symbol_key = st.selectbox(
        "Símbolo de exibição",
        options=list(CURRENCY_SYMBOLS.keys()),
        help="Símbolo usado na formatação dos valores exibidos",
    )

chosen_value_col    = value_col    if value_col    != "— nenhuma —" else None
chosen_currency_col = currency_col if currency_col != "— nenhuma —" else None
chosen_symbol       = CURRENCY_SYMBOLS[symbol_key]

# ── Botão e resultado ─────────────────────────────────────────────────────────

if st.button("▶ CONCATENAR E BAIXAR CSV"):
    result_df = pd.concat([df[ref_cols] for _, df in good_frames], ignore_index=True)

    # ── Resumo de totais ──────────────────────────────────────────────────────
    if chosen_value_col:
        st.markdown('<div class="section-title">Resumo de totais</div>', unsafe_allow_html=True)
        summary_html = build_currency_summary_html(
            result_df, chosen_value_col, chosen_currency_col, chosen_symbol
        )
        st.markdown(summary_html, unsafe_allow_html=True)
    else:
        st.info("Selecione uma coluna para somar para ver o resumo de totais.")

    # ── Preview ───────────────────────────────────────────────────────────────
    st.markdown('<div class="section-title">Preview (primeiras 50 linhas)</div>',
                unsafe_allow_html=True)
    st.dataframe(result_df.head(50), use_container_width=True, height=260)

    # ── Download ──────────────────────────────────────────────────────────────
    csv_bytes = result_df.to_csv(index=False, sep=",", encoding="utf-8-sig").encode("utf-8-sig")
    st.download_button(
        label="⬇ BAIXAR stack_output.csv",
        data=csv_bytes,
        file_name="stack_output.csv",
        mime="text/csv",
    )
