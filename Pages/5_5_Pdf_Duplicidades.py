"""
Conversor do Relatório de Duplicidades ECAD → Tabela (XLSX / CSV)
com análise de divergências entre registros da mesma obra.

Execução:
    streamlit run ecad_duplicidades_parser.py
"""

import re
import io
from collections import defaultdict

import pdfplumber
import pandas as pd
import streamlit as st

# ──────────────────────────────────────────────────────────────────────────────
# Configuração da página
# ──────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="ECAD Duplicidades → Tabela",
    page_icon="🎵",
    layout="wide",
)

# ──────────────────────────────────────────────────────────────────────────────
# Constantes de parsing — limites de coluna X (em pontos PDF)
# ──────────────────────────────────────────────────────────────────────────────
X_COD    = 77    # cod_titular  | nome
X_PSEUDO = 223   # nome         | pseudonimo
X_ASSOC  = 325   # pseudonimo   | associacao
X_CAT    = 385   # associacao   | cat
X_SUBCAT = 405   # cat          | subcat
X_INICIO = 440   # subcat       | inicio / fim
X_PART   = 515   # inicio/fim   | participacao
X_LINK   = 553   # participacao | link

SKIP_LINES = (
    "RELATÓRIO DE DUPLICIDADES",
    "PERÍODO:",
    "OBRA MUSICAL NACIONAL",
    "TITULAR PRINCIPAL:",
    "CONTRATO DE EDIÇÃO",
    "CÓD. TITULAR",
    "TOTA L:",
    "ECAD.Tec",
    "Página ",
)

# ──────────────────────────────────────────────────────────────────────────────
# Extração do PDF via pdfplumber (compatível com Windows/Linux/Mac)
# ──────────────────────────────────────────────────────────────────────────────

def get_page_count(pdf_bytes: bytes) -> int:
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        return len(pdf.pages)


# ──────────────────────────────────────────────────────────────────────────────
# Parsing baseado em posição X das palavras (robusto para qualquer OS)
# ──────────────────────────────────────────────────────────────────────────────

def _bucket_row(row_words: list) -> dict:
    """Distribui as palavras de uma linha pelos campos da tabela usando X."""
    cod_p, nome_p, pseudo_p, assoc_p, cat_p, subcat_p, \
        inicio_p, fim_p, part_p, link_p = ([] for _ in range(10))

    for w in sorted(row_words, key=lambda x: x["x0"]):
        x, t = w["x0"], w["text"]
        if x < X_COD:
            cod_p.append(t)
        elif x < X_PSEUDO:
            nome_p.append(t)
        elif x < X_ASSOC:
            pseudo_p.append(t)
        elif x < X_CAT:
            assoc_p.append(t)
        elif x < X_SUBCAT:
            cat_p.append(t)
        elif x < X_INICIO:
            subcat_p.append(t)
        elif x < X_PART:
            if not inicio_p:
                inicio_p.append(t)
            else:
                fim_p.append(t)
        elif x < X_LINK:
            part_p.append(t)
        else:
            link_p.append(t)

    return {
        "cod_titular":  " ".join(cod_p),
        "nome":         " ".join(nome_p),
        "pseudonimo":   " ".join(pseudo_p),
        "associacao":   " ".join(assoc_p),
        "cat":          " ".join(cat_p),
        "subcat":       " ".join(subcat_p),
        "inicio":       " ".join(inicio_p),
        "fim":          " ".join(fim_p),
        "participacao": " ".join(part_p),
        "link":         " ".join(link_p),
    }


def parse_report(pdf_bytes: bytes, first_page: int = 1, last_page: int | None = None) -> pd.DataFrame:
    """
    Extrai e parseia o relatório de duplicidades do ECAD usando posições X
    das palavras — funciona em qualquer sistema operacional.
    """
    rows: list[dict] = []
    current_obra = current_codigo = current_resp = None

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        end = last_page or len(pdf.pages)
        pages = pdf.pages[first_page - 1: end]

        for page in pages:
            words = page.extract_words(x_tolerance=3, y_tolerance=3)

            by_row: dict[int, list] = defaultdict(list)
            for w in words:
                by_row[round(w["top"])].append(w)

            last_titular_idx: int | None = None
            last_top: float = -999.0

            for top in sorted(by_row.keys()):
                row_words = sorted(by_row[top], key=lambda w: w["x0"])
                line = " ".join(w["text"] for w in row_words)

                if any(s in line for s in SKIP_LINES):
                    continue
                if re.match(r"^\d{2}/\d{2}/\d{4}\s+-\s+\d{2}:\d{2}", line):
                    continue

                obra_m = re.match(r"OBRA MUSICAL:\s*(.+)", line)
                if obra_m:
                    current_obra = re.sub(r"\s+", " ", obra_m.group(1)).strip()
                    last_titular_idx = None
                    continue

                cod_m = re.match(
                    r"CÓDIGO:\s*(\d+)\s+RESP\.\s+CADASTRO ORIGINAL:\s*(\w+)", line
                )
                if cod_m:
                    current_codigo = cod_m.group(1)
                    current_resp   = cod_m.group(2)
                    last_titular_idx = None
                    continue

                first_word = row_words[0]

                if (
                    re.match(r"^\d{4,}$", first_word["text"])
                    and first_word["x0"] < X_COD
                    and current_obra
                    and current_codigo
                ):
                    bucketed = _bucket_row(row_words)
                    rows.append({
                        "obra_musical":           current_obra,
                        "codigo":                 current_codigo,
                        "resp_cadastro_original": current_resp,
                        **bucketed,
                    })
                    last_titular_idx = len(rows) - 1
                    last_top = top
                    continue

                # Linha de continuação do pseudônimo (linha imediatamente abaixo)
                if (
                    last_titular_idx is not None
                    and (top - last_top) < 12
                    and all(X_PSEUDO <= w["x0"] < X_ASSOC for w in row_words)
                ):
                    extra = " ".join(w["text"] for w in row_words)
                    existing = rows[last_titular_idx]["pseudonimo"]
                    rows[last_titular_idx]["pseudonimo"] = (existing + " " + extra).strip()

    return pd.DataFrame(rows)


# ──────────────────────────────────────────────────────────────────────────────
# Análise de divergências
# ──────────────────────────────────────────────────────────────────────────────

def build_divergence_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    obra_grupos: dict[str, dict] = {}

    for obra, g in df.groupby("obra_musical", sort=False):
        codigos_obra = g["codigo"].unique().tolist()
        titular_map: dict[str, dict[str, str]] = {}
        nome_map:    dict[str, str]             = {}
        tits_por_codigo: dict[str, set]         = {}

        for _, row in g.iterrows():
            ct  = row["cod_titular"]
            cod = row["codigo"]
            titular_map.setdefault(ct, {})[cod] = row["participacao"]
            if ct not in nome_map and row["nome"]:
                nome_map[ct] = row["nome"]
            tits_por_codigo.setdefault(cod, set()).add(ct)

        obra_grupos[obra] = {
            "codigos":         codigos_obra,
            "titular_map":     titular_map,
            "nome_map":        nome_map,
            "tits_por_codigo": tits_por_codigo,
            "todos_titulares": set(titular_map.keys()),
        }

    col_data: dict[str, list] = {
        "total_codigos_obra":              [],
        "titular_em_quantos_codigos":      [],
        "diverge_presenca":                [],
        "ausente_nos_codigos":             [],
        "participacao_por_codigo":         [],
        "diverge_participacao":            [],
        "titulares_so_neste_codigo":       [],
        "titulares_ausentes_neste_codigo": [],
        "ha_divergencia":                  [],
    }

    for _, row in df.iterrows():
        obra = row["obra_musical"]
        cod  = row["codigo"]
        ct   = row["cod_titular"]
        og   = obra_grupos[obra]

        codigos_obra   = og["codigos"]
        titular_map    = og["titular_map"]
        nome_map       = og["nome_map"]
        tits_por_cod   = og["tits_por_codigo"]
        todos_tits     = og["todos_titulares"]

        n_codigos_obra = len(codigos_obra)
        part_map       = titular_map.get(ct, {})
        codigos_com_ct = list(part_map.keys())
        ausentes_em    = [c for c in codigos_obra if c not in part_map]
        diverge_pres   = len(ausentes_em) > 0

        if n_codigos_obra > 1:
            partes_fmt = " | ".join(
                f"{c}: {part_map[c]}%"
                for c in codigos_obra
                if c in part_map
            )
        else:
            partes_fmt = ""

        valores_part = list(part_map.values())
        diverge_part = len(set(valores_part)) > 1 if len(valores_part) > 1 else False

        tits_deste = tits_por_cod.get(cod, set())
        exclusivos = [
            nome_map.get(t, t)
            for t in tits_deste
            if set(titular_map.get(t, {}).keys()) == {cod}
        ]
        exclusivos_str = " | ".join(sorted(exclusivos))

        ausentes_aqui = [
            nome_map.get(t, t)
            for t in todos_tits
            if t not in tits_deste
        ]
        ausentes_aqui_str = " | ".join(sorted(ausentes_aqui))

        ha_div = any([
            diverge_pres,
            diverge_part,
            bool(exclusivos_str),
            bool(ausentes_aqui_str),
        ])

        col_data["total_codigos_obra"].append(n_codigos_obra)
        col_data["titular_em_quantos_codigos"].append(len(codigos_com_ct))
        col_data["diverge_presenca"].append(diverge_pres)
        col_data["ausente_nos_codigos"].append(" | ".join(ausentes_em))
        col_data["participacao_por_codigo"].append(partes_fmt)
        col_data["diverge_participacao"].append(diverge_part)
        col_data["titulares_so_neste_codigo"].append(exclusivos_str)
        col_data["titulares_ausentes_neste_codigo"].append(ausentes_aqui_str)
        col_data["ha_divergencia"].append(ha_div)

    for col, vals in col_data.items():
        df[col] = vals

    return df


# ──────────────────────────────────────────────────────────────────────────────
# Exportadores
# ──────────────────────────────────────────────────────────────────────────────

COLUMN_LABELS = {
    "obra_musical":                    "Obra Musical",
    "codigo":                          "Código",
    "resp_cadastro_original":          "Resp. Cadastro Original",
    "cod_titular":                     "Cód. Titular",
    "nome":                            "Nome",
    "pseudonimo":                      "Pseudônimo",
    "associacao":                      "Associação",
    "cat":                             "Cat.",
    "subcat":                          "Subcat.",
    "inicio":                          "Início",
    "fim":                             "Fim",
    "participacao":                    "Participação (%)",
    "link":                            "Link",
    "total_codigos_obra":              "Total de Códigos (Obra)",
    "titular_em_quantos_codigos":      "Titular em Quantos Códigos",
    "diverge_presenca":                "Diverge: Presença?",
    "ausente_nos_codigos":             "Ausente nos Códigos",
    "participacao_por_codigo":         "Participação por Código",
    "diverge_participacao":            "Diverge: Participação?",
    "titulares_so_neste_codigo":       "Titulares Só Neste Código",
    "titulares_ausentes_neste_codigo": "Titulares Ausentes Neste Código",
    "ha_divergencia":                  "Há Divergência?",
}

N_BASE_COLS = 13


def to_xlsx(df: pd.DataFrame) -> bytes:
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    df_exp = df.rename(columns=COLUMN_LABELS)
    buf    = io.BytesIO()

    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df_exp.to_excel(writer, index=False, sheet_name="Duplicidades")
        ws = writer.sheets["Duplicidades"]

        fill_base     = PatternFill("solid", fgColor="1F3864")
        fill_analysis = PatternFill("solid", fgColor="374151")
        font_hdr      = Font(bold=True, color="FFFFFF", size=10)
        align_ctr     = Alignment(horizontal="center", vertical="center", wrap_text=True)

        for idx, cell in enumerate(ws[1], start=1):
            cell.font      = font_hdr
            cell.fill      = fill_analysis if idx > N_BASE_COLS else fill_base
            cell.alignment = align_ctr

        ws.row_dimensions[1].height = 36

        fill_even      = PatternFill("solid", fgColor="EBF0FA")
        fill_even_anal = PatternFill("solid", fgColor="F0F0F0")
        align_top      = Alignment(vertical="top", wrap_text=True)

        for r_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
            is_even = r_idx % 2 == 0
            for c_idx, cell in enumerate(row, start=1):
                cell.alignment = align_top
                if is_even:
                    cell.fill = fill_even_anal if c_idx > N_BASE_COLS else fill_even

        widths = {
            "Obra Musical":                    40,
            "Código":                          13,
            "Resp. Cadastro Original":         10,
            "Cód. Titular":                    13,
            "Nome":                            38,
            "Pseudônimo":                      22,
            "Associação":                      10,
            "Cat.":                             6,
            "Subcat.":                          8,
            "Início":                          11,
            "Fim":                             11,
            "Participação (%)":                13,
            "Link":                             6,
            "Total de Códigos (Obra)":         10,
            "Titular em Quantos Códigos":      10,
            "Diverge: Presença?":              11,
            "Ausente nos Códigos":             34,
            "Participação por Código":         50,
            "Diverge: Participação?":          12,
            "Titulares Só Neste Código":       50,
            "Titulares Ausentes Neste Código": 50,
            "Há Divergência?":                 11,
        }

        for col in ws.columns:
            hdr = col[0].value
            ws.column_dimensions[col[0].column_letter].width = widths.get(hdr, 18)

        thick_left = Border(left=Side(style="medium", color="374151"))
        for row in ws.iter_rows(
            min_row=1, max_row=ws.max_row,
            min_col=N_BASE_COLS + 1, max_col=N_BASE_COLS + 1,
        ):
            for cell in row:
                cell.border = thick_left

        ws.freeze_panes = "A2"

    return buf.getvalue()


def to_csv(df: pd.DataFrame) -> bytes:
    return (
        df.rename(columns=COLUMN_LABELS)
          .to_csv(index=False, encoding="utf-8-sig")
          .encode("utf-8-sig")
    )


# ──────────────────────────────────────────────────────────────────────────────
# Interface Streamlit
# ──────────────────────────────────────────────────────────────────────────────

def main():
    st.markdown(
        """
        <style>
        .main-header {
            background: linear-gradient(135deg,#1F3864,#2E75B6);
            color:white; padding:1.5rem 2rem; border-radius:10px; margin-bottom:1.5rem;
        }
        .main-header h1 { margin:0; font-size:1.8rem; }
        .main-header p  { margin:0.3rem 0 0; opacity:.85; font-size:.95rem; }
        .metric-card {
            background:#f0f4fa; border-radius:8px; padding:1rem 1.5rem;
            text-align:center; border-left:4px solid #2E75B6;
        }
        .metric-card .value { font-size:2rem; font-weight:700; color:#1F3864; }
        .metric-card .label { font-size:.85rem; color:#555; margin-top:.2rem; }
        </style>
        <div class="main-header">
            <h1>🎵 ECAD Duplicidades → Tabela</h1>
            <p>
               Converte relatório de duplicidades PDF do ECAD → dados estruturados com análise de divergências</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    uploaded = st.file_uploader(
        "📄 Selecione o Relatório de Duplicidades (PDF)",
        type=["pdf"],
        help="Exporte o relatório diretamente do ECAD.Tec Distribuição",
    )

    if not uploaded:
        st.info("Faça o upload do PDF para começar. O arquivo não é armazenado.")
        return

    pdf_bytes = uploaded.read()

    with st.spinner("Contando páginas..."):
        total_pages = get_page_count(pdf_bytes)

    st.success(f"PDF carregado: **{uploaded.name}** — {total_pages} páginas detectadas")

    col_a, col_b = st.columns([2, 1])
    with col_a:
        page_range = st.slider(
            "Intervalo de páginas a processar",
            min_value=1,
            max_value=max(total_pages, 1),
            value=(1, total_pages if total_pages > 0 else 1),
        )
    with col_b:
        output_format = st.radio("Formato de saída", ["XLSX (Excel)", "CSV"], index=0)

    if st.button("🚀 Processar Relatório", type="primary", use_container_width=True):

        prog = st.progress(0, text="Extraindo e parseando PDF...")
        df = parse_report(pdf_bytes, page_range[0], page_range[1])

        if df.empty:
            st.error(
                "Nenhum dado encontrado. "
                "Verifique se o PDF é um Relatório de Duplicidades do ECAD."
            )
            return

        prog.progress(60, text="Calculando divergências...")
        df = build_divergence_columns(df)

        prog.progress(85, text="Gerando arquivo...")

        n_obras     = df["obra_musical"].nunique()
        n_codigos   = df["codigo"].nunique()
        n_titulares = df["cod_titular"].nunique()
        n_com_div   = df[df["ha_divergencia"]]["obra_musical"].nunique()

        prog.progress(100, text="Concluído!")
        prog.empty()

        st.markdown("### 📊 Resumo")
        m1, m2, m3, m4 = st.columns(4)
        for col, label, val in [
            (m1, "Obras Musicais",        n_obras),
            (m2, "Registros (Códigos)",   n_codigos),
            (m3, "Titulares Únicos",      n_titulares),
            (m4, "Obras com Divergência", n_com_div),
        ]:
            col.markdown(
                f'<div class="metric-card">'
                f'<div class="value">{val:,}</div>'
                f'<div class="label">{label}</div>'
                f'</div>',
                unsafe_allow_html=True,
            )

        st.markdown("---")

        st.markdown("### 🔍 Preview dos dados")
        tab1, tab2 = st.tabs(["Todas as linhas", "Apenas linhas com divergência"])

        with tab1:
            n_prev = st.slider("Linhas no preview", 10, 100, 30, step=10, key="p1")
            st.dataframe(
                df.rename(columns=COLUMN_LABELS).head(n_prev),
                use_container_width=True, height=380,
            )
        with tab2:
            df_div = df[df["ha_divergencia"]]
            st.caption(f"{len(df_div):,} linhas com divergência detectada")
            st.dataframe(
                df_div.rename(columns=COLUMN_LABELS),
                use_container_width=True, height=420,
            )

        df_exp = df
        with st.expander("🔎 Filtrar dados antes de exportar"):
            obras    = ["(todas)"] + sorted(df["obra_musical"].unique().tolist())
            obra_sel = st.selectbox("Filtrar por Obra Musical", obras)
            assocs   = ["(todas)"] + sorted(df["associacao"].dropna().unique().tolist())
            ass_sel  = st.selectbox("Filtrar por Associação", assocs)
            so_div   = st.checkbox("Exportar apenas linhas com divergência", value=False)

            df_exp = df.copy()
            if obra_sel != "(todas)":
                df_exp = df_exp[df_exp["obra_musical"] == obra_sel]
            if ass_sel != "(todas)":
                df_exp = df_exp[df_exp["associacao"]   == ass_sel]
            if so_div:
                df_exp = df_exp[df_exp["ha_divergencia"]]

            st.info(f"Linhas a exportar: **{len(df_exp):,}**")

        st.markdown("### ⬇️ Download")
        fname_base = f"ecad_duplicidades_p{page_range[0]}-{page_range[1]}"

        if "CSV" in output_format:
            data  = to_csv(df_exp)
            fname = f"{fname_base}.csv"
            mime  = "text/csv"
            lbl   = "📥 Baixar CSV"
        else:
            data  = to_xlsx(df_exp)
            fname = f"{fname_base}.xlsx"
            mime  = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            lbl   = "📥 Baixar XLSX (Excel)"

        st.download_button(
            label=lbl, data=data, file_name=fname, mime=mime,
            use_container_width=True, type="primary",
        )

        with st.expander("📈 Distribuição por Obra Musical"):
            summary = (
                df.groupby("obra_musical")
                  .agg(
                      codigos=("codigo",        "nunique"),
                      linhas =("cod_titular",   "count"),
                      div    =("ha_divergencia","any"),
                  )
                  .reset_index()
                  .sort_values(["div", "codigos"], ascending=[False, False])
            )
            summary.columns = [
                "Obra Musical", "Registros (Códigos)", "Linhas", "Tem Divergência"
            ]
            st.dataframe(summary, use_container_width=True, height=320)

        st.markdown(
            "> **Nota:** os nomes das obras refletem o texto original do PDF do ECAD."
        )


if __name__ == "__main__":
    main()