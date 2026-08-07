import streamlit as st
import pandas as pd
import io
import re
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

st.title("Processamento de Royalties")

periodo = st.text_input("Período (AAAA-MM)", placeholder="2024-03")

uploaded_file = st.file_uploader("Carregar arquivo de operações (.xlsx)", type=["xlsx"])
incomes_file = st.file_uploader("Carregar lista de incomes (.xlsx)", type=["xlsx"])

def strip_prefix(name):
    """Remove prefixo do tipo '2026FEV SYNC - ' ou '2026FEV LICENCIAMENTO - ' para normalizar o nome."""
    return re.sub(r'^\d{4}\w+\s+(SYNC|LICENCIAMENTO)\s+-\s+', '', str(name))


def classify(row):
    rh = str(row["Rights-Holder"])
    t = str(row["Type"])
    if t == "Advance Refund":
        return "Recuperação de Adiantamentos"
    if re.search(r"License", t, re.IGNORECASE):
        return "Direitos Autorais – Licenciamento" if rh == "PHONOLITE" else "Repasses Editora - Provisão"
    if re.search(r"Synchro", t, re.IGNORECASE):
        return "Direitos Autorais – Sincronização" if rh == "PHONOLITE" else "Repasses Editora - Provisão"
    if rh == "PHONOLITE":
        return "Direitos Autorais – Fonomecânicos Digitais"
    return "Repasses Editora - Provisão"


def is_receita_phonolite(categoria):
    """Tudo que começa com 'Direitos Autorais' é Receita Phonolite."""
    return str(categoria).startswith("Direitos Autorais")


if uploaded_file:
    df = pd.read_excel(uploaded_file)

    # Recuperação de Adiantamentos é apenas uma referência para abater o valor
    # que entrou de uma fonte específica (ex.: ABRAMUS DIGITAL). Não é dinheiro
    # novo processado, então fica marcada com valor negativo e sai fora da
    # soma geral (Total Processado / Quadro Resumo).
    mask_adiantamento = df["Type"] == "Advance Refund"
    df.loc[mask_adiantamento, "Amount"] = -df.loc[mask_adiantamento, "Amount"].abs()

    # Monta payer_map usando nome normalizado (sem prefixo SYNC/LICENCIAMENTO)
    payer_map = {}
    if incomes_file:
        df_incomes = pd.read_excel(incomes_file)
        df_incomes["_base"] = df_incomes["Name"].apply(strip_prefix)
        payer_map = df_incomes.drop_duplicates("_base").set_index("_base")["Payer"].to_dict()

    # --- MÉTRICAS GERAIS ---
    # Total Processado não inclui a Recuperação de Adiantamentos: ela é só
    # referência de abatimento, não soma/subtrai do total geral.
    total = df.loc[~mask_adiantamento, "Amount"].sum()
    st.metric("Total Processado", f"R$ {total:,.2f}")

    st.subheader("Total por Tipo")
    total_por_tipo = df.groupby("Type")["Amount"].sum().reset_index()
    total_por_tipo.columns = ["Tipo", "Total"]
    total_por_tipo["Total"] = total_por_tipo["Total"].map(lambda x: f"R$ {x:,.2f}")
    st.dataframe(total_por_tipo, use_container_width=True, hide_index=True)

    # --- DATAFRAME DE CATEGORIZAÇÃO ---
    st.subheader("Resumo por Categoria Financeira")

    df["Categoria"] = df.apply(classify, axis=1)
    df_resumo = (
        df.groupby(["Categoria", "Name", "Type"])["Amount"]
        .sum()
        .reset_index()
        .rename(columns={"Name": "Nome", "Type": "Tipo", "Amount": "Valor"})
    )

    # Mapeia Fonte via nome normalizado
    df_resumo["_base"] = df_resumo["Nome"].apply(strip_prefix)
    df_resumo["Fonte"] = df_resumo["_base"].map(payer_map).fillna("")
    df_resumo.drop(columns=["_base"], inplace=True)

    if periodo:
        df_resumo.insert(0, "Período", periodo)

    df_display = df_resumo.copy()
    df_display["Valor"] = df_display["Valor"].map(lambda x: f"R$ {x:,.2f}")
    st.dataframe(df_display, use_container_width=True, hide_index=True)

    # --- PREPARA AGRUPAMENTO (Fonte / Categoria) ---
    # A Fonte usada é sempre a da própria coluna "Fonte" (mapeada via incomes),
    # sem nenhuma fonte fixa/hardcoded.
    df_resumo_grp = df_resumo.copy()
    df_resumo_grp["Fonte"] = df_resumo_grp["Fonte"].fillna("(vazio)").replace("", "(vazio)")

    resumo_agrupado = (
        df_resumo_grp.groupby(["Fonte", "Categoria"])["Valor"]
        .sum()
        .reset_index()
        .sort_values(["Fonte", "Categoria"])
    )

    # --- RESUMO EM TELA (sem linha de Total Geral) ---
    st.subheader("Resumo por Fonte e Categoria")

    rows_tela = []
    for fonte in resumo_agrupado["Fonte"].unique():
        grp = resumo_agrupado[resumo_agrupado["Fonte"] == fonte]
        rows_tela.append({"Rótulos de Linha": fonte, "Soma de Valor": grp["Valor"].sum()})
        for _, r in grp.iterrows():
            rows_tela.append({"Rótulos de Linha": f"   {r['Categoria']}", "Soma de Valor": r["Valor"]})

    df_resumo_display = pd.DataFrame(rows_tela)
    df_resumo_display["Soma de Valor"] = df_resumo_display["Soma de Valor"].map(lambda x: f"R$ {x:,.2f}")
    st.dataframe(df_resumo_display, use_container_width=True, hide_index=True)

    # --- QUADRO RESUMO: RECEITA PHONOLITE x REPASSES/OUTROS ---
    st.subheader("Quadro Resumo: Receita Phonolite")

    # Recuperação de Adiantamentos fica fora dessa soma: é só referência de
    # abatimento de uma fonte específica, não compõe o total geral.
    df_quadro = df_resumo[df_resumo["Categoria"] != "Recuperação de Adiantamentos"].copy()
    df_quadro["Classificação"] = df_quadro["Categoria"].apply(
        lambda c: "Receita Phonolite" if is_receita_phonolite(c) else "Repasses e Outros"
    )
    quadro_totais = df_quadro.groupby("Classificação")["Valor"].sum().reset_index()
    quadro_display = quadro_totais.copy()
    quadro_display["Valor"] = quadro_display["Valor"].map(lambda x: f"R$ {x:,.2f}")
    st.dataframe(quadro_display, use_container_width=True, hide_index=True)

    # --- EXPORT XLSX ---
    def write_sheet_plain(ws, df_data):
        """Planilha1: cabeçalho azul escuro/branco bold, linhas pares verde claro, autofilter."""
        df_sorted = df_data.sort_values("Nome") if "Nome" in df_data.columns else df_data

        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill("solid", start_color="1F4E79")
        header_align = Alignment(horizontal="center", vertical="center")
        row_fill_alt = PatternFill("solid", start_color="E2EFDA")
        thin = Side(style="thin")
        brd = Border(left=thin, right=thin, top=thin, bottom=thin)

        headers = list(df_sorted.columns)
        for col, h in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=h)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_align
            cell.border = brd

        for row_idx, row in enumerate(df_sorted.itertuples(index=False), 2):
            fill = row_fill_alt if row_idx % 2 == 0 else PatternFill(fill_type=None)
            for col, val in enumerate(row, 1):
                cell = ws.cell(row=row_idx, column=col, value=val)
                cell.fill = fill
                cell.border = brd

        ws.auto_filter.ref = ws.dimensions
        for col in ws.columns:
            max_len = max(len(str(c.value or "")) for c in col) + 4
            ws.column_dimensions[col[0].column_letter].width = min(max_len, 50)

    def write_resumo_sheet(ws, df_data):
        """Resumo agrupado: Fonte → Categoria → subtotais (sem Total Geral),
        usando a Fonte real de cada operação (sem fonte fixa/hardcoded para
        Recuperação de Adiantamentos, que já entra com valor negativo),
        mais um quadro-resumo lateral de Receita Phonolite x Repasses/Outros."""
        bold_white = Font(bold=True, color="FFFFFF")
        bold_black = Font(bold=True)
        normal = Font(bold=False)

        fill_fonte = PatternFill("solid", start_color="1F4E79")
        fill_total = PatternFill("solid", start_color="BDD7EE")
        no_fill = PatternFill(fill_type=None)

        thin = Side(style="thin")
        medium = Side(style="medium")

        def border_row(left=thin, right=thin, top=thin, bottom=thin):
            return Border(left=left, right=right, top=top, bottom=bottom)

        df_r = df_data.copy()
        df_r["Fonte"] = df_r["Fonte"].fillna("(vazio)").replace("", "(vazio)")

        grouped = df_r.groupby(["Fonte", "Categoria"], sort=True)["Valor"].sum()

        ws.column_dimensions["A"].width = 40
        ws.column_dimensions["B"].width = 18
        ws.column_dimensions["C"].width = 4
        ws.column_dimensions["D"].width = 24
        ws.column_dimensions["E"].width = 18

        # --- Tabela principal: Fonte / Categoria ---
        for c, h in enumerate(["Rótulos de Linha", "Soma de Valor"], 1):
            cell = ws.cell(row=1, column=c, value=h)
            cell.font = bold_white
            cell.fill = fill_fonte
            cell.alignment = Alignment(horizontal="center")
            cell.border = border_row(left=medium, right=medium, top=medium, bottom=medium)

        row_idx = 2
        for fonte, grp in grouped.groupby(level=0):
            subtotal = grp.sum()

            c1 = ws.cell(row=row_idx, column=1, value=fonte)
            c1.font = bold_white
            c1.fill = fill_fonte
            c1.border = border_row(left=medium, right=thin, top=medium, bottom=thin)
            c2 = ws.cell(row=row_idx, column=2, value=subtotal)
            c2.font = bold_white
            c2.fill = fill_fonte
            c2.number_format = '#,##0.00'
            c2.border = border_row(left=thin, right=medium, top=medium, bottom=thin)
            row_idx += 1

            for (_, cat), val in grp.items():
                c1 = ws.cell(row=row_idx, column=1, value=f"   {cat}")
                c1.font = normal
                c1.fill = no_fill
                c1.border = border_row(left=medium, right=thin)
                c2 = ws.cell(row=row_idx, column=2, value=val)
                c2.font = normal
                c2.fill = no_fill
                c2.number_format = '#,##0.00'
                c2.border = border_row(left=thin, right=medium)
                row_idx += 1

        # --- Quadro-resumo lateral: Receita Phonolite x Repasses/Outros ---
        # Recuperação de Adiantamentos fica fora dessa soma: é só referência
        # de abatimento de uma fonte específica, não compõe o total geral.
        df_quadro = df_data[df_data["Categoria"] != "Recuperação de Adiantamentos"].copy()
        df_quadro["Classificação"] = df_quadro["Categoria"].apply(
            lambda c: "Receita Phonolite" if is_receita_phonolite(c) else "Repasses e Outros"
        )
        quadro_totais = df_quadro.groupby("Classificação")["Valor"].sum()

        for c, h in enumerate(["Classificação", "Valor"], 4):
            cell = ws.cell(row=1, column=c, value=h)
            cell.font = bold_white
            cell.fill = fill_fonte
            cell.alignment = Alignment(horizontal="center")
            cell.border = border_row(left=medium, right=medium, top=medium, bottom=medium)

        q_row = 2
        for classificacao, valor in quadro_totais.items():
            c1 = ws.cell(row=q_row, column=4, value=classificacao)
            c1.font = normal
            c1.border = border_row(left=medium, right=thin)
            c2 = ws.cell(row=q_row, column=5, value=valor)
            c2.font = normal
            c2.number_format = '#,##0.00'
            c2.border = border_row(left=thin, right=medium)
            q_row += 1

        c1 = ws.cell(row=q_row, column=4, value="Total Geral")
        c1.font = bold_black
        c1.fill = fill_total
        c1.border = border_row(left=medium, right=thin, top=medium, bottom=medium)
        c2 = ws.cell(row=q_row, column=5, value=quadro_totais.sum())
        c2.font = bold_black
        c2.fill = fill_total
        c2.number_format = '#,##0.00'
        c2.border = border_row(left=thin, right=medium, top=medium, bottom=medium)

    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Planilha1"
    write_sheet_plain(ws1, df_resumo)

    ws2 = wb.create_sheet("Resumo")
    write_resumo_sheet(ws2, df_resumo)

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    fname = f"resumo_royalties_{periodo}.xlsx" if periodo else "resumo_royalties.xlsx"
    st.download_button(
        label="📥 Baixar Resumo para o Financeiro (.xlsx)",
        data=buffer,
        file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )