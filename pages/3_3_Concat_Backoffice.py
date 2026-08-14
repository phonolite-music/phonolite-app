"""
Concat Backoffice — concatenação e totalização de relatórios de royalties.

Principais mudanças em relação à versão anterior:
  * Detecção dinâmica da linha de cabeçalho (antes fixo em header=5, o que
    destruía os arquivos cujo cabeçalho está na linha 1).
  * Leitura dos uploads a partir de bytes em memória -> nunca esbarra em
    ponteiro de buffer já consumido.
  * Validação de schema antes de concatenar, com relatório de divergências.
  * Coluna de origem (arquivo / fonte) no consolidado.
  * Resultados persistidos em st.session_state -> não somem a cada rerun.
  * Detecção de linhas duplicadas e de statements repetidos.

Requisitos: streamlit, pandas, openpyxl, xlsxwriter
"""

from __future__ import annotations

import hashlib
import re
import warnings
from dataclasses import dataclass, field
from io import BytesIO

import pandas as pd
import streamlit as st

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# ---------------------------------------------------------------------------
# Configuração
# ---------------------------------------------------------------------------

APP_TITLE = "Concat Backoffice"
APP_CAPTION = (
    "Concatena, valida e totaliza os relatórios Backoffice "
    "para conferência e inclusão no Reprtoir."
)

# Colunas-âncora usadas para localizar a linha de cabeçalho. Basta que
# MIN_ANCHOR_HITS delas apareçam na linha para considerá-la o cabeçalho.
ANCHOR_COLUMNS = {
    "bo_payeesid",
    "bo_songcode",
    "payee_name",
    "payee_statement_#",
    "publishers_songcode",
    "royalties_to_be_paid",
    "royalties_to_be_paid_$",
    "royaties_gross",
    "royalties_gross",
    "song_title",
    "statement_period",
}
MIN_ANCHOR_HITS = 3

# Quantas linhas do topo inspecionar à procura do cabeçalho.
MAX_HEADER_SCAN = 20

# Colunas candidatas a "valor a pagar", em ordem de preferência.
ROYALTY_COLUMN_CANDIDATES = [
    "ROYALTIES_TO_BE_PAID",
    "ROYALTIES_TO_BE_PAID_$",
    "ROYALTIES_TO_BE_PAID_R$",
]

# Colunas auxiliares acrescentadas pelo app.
COL_ARQUIVO = "_Arquivo_Origem"
COL_HEADER_ROW = "_Linha_Cabecalho"

SUPPORTED_EXTENSIONS = (".xlsx", ".xlsm", ".xls", ".csv")


# ---------------------------------------------------------------------------
# Núcleo (sem dependência de Streamlit — testável isoladamente)
# ---------------------------------------------------------------------------


@dataclass
class LeituraArquivo:
    """Resultado da leitura de um único arquivo."""

    nome: str
    df: pd.DataFrame | None = None
    header_row: int | None = None
    coluna_royalty: str | None = None
    total_royalty: float | None = None
    erro: str | None = None
    avisos: list[str] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return self.df is not None and self.erro is None


def _normaliza(valor) -> str:
    """Normaliza um rótulo de coluna para comparação."""
    if valor is None:
        return ""
    return re.sub(r"\s+", "_", str(valor).strip().lower())


def detecta_linha_cabecalho(
    raw: pd.DataFrame,
    anchors: set[str] = ANCHOR_COLUMNS,
    min_hits: int = MIN_ANCHOR_HITS,
) -> int | None:
    """Devolve o índice (0-based) da linha de cabeçalho, ou None se não achar.

    Percorre as primeiras linhas do arquivo lido sem cabeçalho e escolhe a
    primeira que contenha pelo menos ``min_hits`` colunas-âncora conhecidas.
    """
    melhor_idx, melhor_hits = None, 0
    for idx in range(len(raw)):
        celulas = {_normaliza(v) for v in raw.iloc[idx].tolist()}
        hits = len(celulas & anchors)
        if hits > melhor_hits:
            melhor_idx, melhor_hits = idx, hits
    return melhor_idx if melhor_hits >= min_hits else None


def resolve_coluna_royalty(colunas) -> str | None:
    """Encontra a coluna de valor a pagar, tolerante a variações de nome."""
    mapa = {_normaliza(c): c for c in colunas}
    for candidata in ROYALTY_COLUMN_CANDIDATES:
        chave = _normaliza(candidata)
        if chave in mapa:
            return mapa[chave]
    # Fallback: qualquer coluna que contenha "to_be_paid".
    for chave, original in mapa.items():
        if "to_be_paid" in chave:
            return original
    return None


def le_arquivo(nome: str, conteudo: bytes) -> LeituraArquivo:
    """Lê um arquivo detectando o cabeçalho automaticamente.

    ``conteudo`` são os bytes brutos: cada leitura usa um BytesIO novo, então
    nunca há problema de ponteiro consumido por leituras anteriores.
    """
    resultado = LeituraArquivo(nome=nome)
    ext = "." + nome.rsplit(".", 1)[-1].lower() if "." in nome else ""

    if ext not in SUPPORTED_EXTENSIONS:
        resultado.erro = f"Extensão não suportada ({ext or 'sem extensão'})."
        return resultado

    try:
        if ext == ".csv":
            preview = pd.read_csv(
                BytesIO(conteudo), header=None, nrows=MAX_HEADER_SCAN,
                sep=None, engine="python", dtype=str,
            )
        else:
            preview = pd.read_excel(
                BytesIO(conteudo), header=None, nrows=MAX_HEADER_SCAN, dtype=str
            )
    except Exception as exc:  # noqa: BLE001
        resultado.erro = f"Falha ao abrir o arquivo: {exc}"
        return resultado

    header_row = detecta_linha_cabecalho(preview)
    if header_row is None:
        resultado.erro = (
            f"Cabeçalho não localizado nas primeiras {MAX_HEADER_SCAN} linhas. "
            "Nenhuma coluna conhecida (BO_PayeesID, ROYALTIES_TO_BE_PAID...) foi encontrada."
        )
        return resultado

    try:
        if ext == ".csv":
            df = pd.read_csv(
                BytesIO(conteudo), header=header_row, sep=None, engine="python"
            )
        else:
            df = pd.read_excel(BytesIO(conteudo), header=header_row)
    except Exception as exc:  # noqa: BLE001
        resultado.erro = f"Falha ao ler os dados: {exc}"
        return resultado

    # Limpeza: nomes de coluna sem espaços sobrando, colunas e linhas vazias fora.
    df.columns = [str(c).strip() for c in df.columns]
    df = df.loc[:, ~df.columns.str.match(r"^Unnamed: \d+$")]
    df = df.dropna(how="all")

    if df.empty:
        resultado.erro = "Arquivo sem linhas de dados após o cabeçalho."
        return resultado

    coluna_royalty = resolve_coluna_royalty(df.columns)
    if coluna_royalty is None:
        resultado.avisos.append("Nenhuma coluna de royalties identificada.")
    else:
        valores = pd.to_numeric(df[coluna_royalty], errors="coerce")
        nao_numericos = int(valores.isna().sum() - df[coluna_royalty].isna().sum())
        if nao_numericos > 0:
            resultado.avisos.append(
                f"{nao_numericos} valor(es) não numérico(s) em '{coluna_royalty}' ignorado(s) na soma."
            )
        resultado.total_royalty = float(valores.sum())

    df.insert(0, COL_ARQUIVO, nome)
    resultado.df = df
    resultado.header_row = header_row
    resultado.coluna_royalty = coluna_royalty
    return resultado


def compara_schemas(leituras: list[LeituraArquivo]) -> pd.DataFrame:
    """Relatório de divergências de colunas entre os arquivos lidos.

    O schema de referência é o do primeiro arquivo válido.
    """
    validos = [l for l in leituras if l.ok]
    if not validos:
        return pd.DataFrame()

    referencia = [c for c in validos[0].df.columns if c != COL_ARQUIVO]
    ref_set = set(referencia)

    linhas = []
    for leitura in validos:
        colunas = [c for c in leitura.df.columns if c != COL_ARQUIVO]
        faltando = ref_set - set(colunas)
        sobrando = set(colunas) - ref_set
        mesma_ordem = colunas == referencia
        linhas.append(
            {
                "Arquivo": leitura.nome,
                "Linha do cabeçalho": (leitura.header_row or 0) + 1,
                "Colunas": len(colunas),
                "Linhas": len(leitura.df),
                "Faltando": ", ".join(sorted(faltando)) or "—",
                "Extras": ", ".join(sorted(sobrando)) or "—",
                "Idêntico à referência": "Sim" if mesma_ordem else "Não",
            }
        )
    return pd.DataFrame(linhas)


def concatena(leituras: list[LeituraArquivo]) -> pd.DataFrame:
    """Concatena as leituras válidas preservando a ordem de colunas da referência."""
    validos = [l for l in leituras if l.ok]
    if not validos:
        return pd.DataFrame()

    referencia = list(validos[0].df.columns)
    consolidado = pd.concat([l.df for l in validos], ignore_index=True, sort=False)

    # Colunas da referência primeiro, extras depois (preserva a leitura humana).
    extras = [c for c in consolidado.columns if c not in referencia]
    return consolidado[referencia + extras]


def formata_brl(valor: float) -> str:
    """Formata como moeda brasileira: 1234.5 -> 'R$ 1.234,50'."""
    if pd.isna(valor):
        return "—"
    texto = f"{valor:,.2f}"
    return "R$ " + texto.replace(",", "\x00").replace(".", ",").replace("\x00", ".")


def resumo_por_arquivo(leituras: list[LeituraArquivo]) -> pd.DataFrame:
    """Tabela de totais por arquivo, sem a linha de total (calculada na exibição)."""
    linhas = [
        {
            "Arquivo": l.nome,
            "Linhas": len(l.df) if l.ok else 0,
            "Coluna de royalties": l.coluna_royalty or "—",
            "Total": round(l.total_royalty, 2) if l.total_royalty is not None else None,
        }
        for l in leituras
        if l.ok
    ]
    return pd.DataFrame(linhas)


def diagnostica_duplicidades(consolidado: pd.DataFrame) -> dict:
    """Contagens de linhas duplicadas e de statements repetidos entre arquivos."""
    if consolidado.empty:
        return {}

    colunas_dados = [c for c in consolidado.columns if not c.startswith("_")]
    dupes = int(consolidado.duplicated(subset=colunas_dados).sum())

    resultado = {"linhas_duplicadas": dupes}

    if "Payee_Statement_#" in consolidado.columns:
        por_stmt = consolidado.groupby("Payee_Statement_#")[COL_ARQUIVO].nunique()
        repetidos = por_stmt[por_stmt > 1]
        resultado["statements_em_varios_arquivos"] = repetidos.index.tolist()

    return resultado


LARGURA_MIN = 10
LARGURA_MAX = 45
AMOSTRA_LARGURA = 2000  # linhas amostradas para estimar a largura da coluna


def calcula_largura(serie: pd.Series, titulo: str) -> int:
    """Largura da coluna a partir do p95 do comprimento do texto.

    Imune a colunas vazias, 100% nulas e a mudanças de comportamento do
    ``astype(str)`` entre versões do pandas (no pandas 3 ele preserva NaN em
    vez de gerar a string 'nan'). Por isso o comprimento é calculado com
    ``len(str(v))`` elemento a elemento, que sempre devolve inteiro.
    """
    base = max(len(str(titulo)) + 2, LARGURA_MIN)

    amostra = serie.head(AMOSTRA_LARGURA)
    if len(amostra) == 0:
        return min(base, LARGURA_MAX)

    comprimentos = amostra.map(lambda v: 0 if v is None else len(str(v)))
    comprimentos = pd.to_numeric(comprimentos, errors="coerce").dropna()
    if comprimentos.empty:
        return min(base, LARGURA_MAX)

    p95 = comprimentos.quantile(0.95)
    if pd.isna(p95):
        return min(base, LARGURA_MAX)

    return int(min(max(base, int(p95) + 2), LARGURA_MAX))


def gera_excel(consolidado: pd.DataFrame, resumo: pd.DataFrame) -> bytes:
    """Gera o .xlsx final com aba de dados e aba de resumo, já formatado."""
    buffer = BytesIO()
    abas: list[tuple[str, pd.DataFrame]] = []

    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        if not consolidado.empty:
            consolidado.to_excel(writer, index=False, sheet_name="Consolidado")
            abas.append(("Consolidado", consolidado))
        if not resumo.empty:
            resumo.to_excel(writer, index=False, sheet_name="Resumo")
            abas.append(("Resumo", resumo))

        if not abas:  # nada a escrever: xlsxwriter não aceita pasta sem abas
            aviso = pd.DataFrame({"Aviso": ["Nenhum dado para exportar."]})
            aviso.to_excel(writer, index=False, sheet_name="Consolidado")
            abas.append(("Consolidado", aviso))

        livro = writer.book
        fmt_header = livro.add_format(
            {"bold": True, "bg_color": "#1F2937", "font_color": "#FFFFFF",
             "border": 1, "align": "left", "valign": "vcenter"}
        )
        fmt_moeda = livro.add_format({"num_format": 'R$ #,##0.00'})

        for nome_aba, df in abas:
            aba = writer.sheets[nome_aba]
            for idx, coluna in enumerate(df.columns):
                # iloc em vez de df[coluna]: imune a nomes de coluna repetidos,
                # que fariam df[coluna] devolver um DataFrame em vez de Series.
                serie = df.iloc[:, idx]
                aba.write(0, idx, coluna, fmt_header)

                eh_moeda = any(
                    chave in str(coluna).upper()
                    for chave in ("ROYALT", "ADMIN_FEE", "TOTAL", "GROSS")
                ) and pd.api.types.is_numeric_dtype(serie)

                aba.set_column(
                    idx, idx,
                    calcula_largura(serie, coluna),
                    fmt_moeda if eh_moeda else None,
                )
            aba.freeze_panes(1, 0)
            aba.autofilter(0, 0, len(df), max(len(df.columns) - 1, 0))

    return buffer.getvalue()


def hash_uploads(arquivos) -> str:
    """Identidade do lote de uploads — usada para invalidar o cache."""
    digest = hashlib.sha256()
    for arquivo in arquivos:
        digest.update(arquivo.name.encode("utf-8"))
        digest.update(str(arquivo.size).encode("utf-8"))
    return digest.hexdigest()


# ---------------------------------------------------------------------------
# Interface
# ---------------------------------------------------------------------------


def main() -> None:
    st.set_page_config(page_title=APP_TITLE, page_icon="📊", layout="wide")
    st.title(APP_TITLE)
    st.caption(APP_CAPTION)

    uploaded_files = st.file_uploader(
        "Faça o upload dos relatórios (.xlsx, .xlsm, .xls, .csv)",
        type=None,
        accept_multiple_files=True,
        key="concat_files",
    )

    if not uploaded_files:
        st.info("Aguardando upload dos arquivos...")
        st.session_state.pop("leituras", None)
        st.session_state.pop("lote", None)
        return

    lote = hash_uploads(uploaded_files)

    # Processa uma única vez por lote; o resultado sobrevive aos reruns.
    if st.session_state.get("lote") != lote:
        leituras: list[LeituraArquivo] = []
        barra = st.progress(0.0, text="Lendo arquivos...")
        for i, arquivo in enumerate(uploaded_files, start=1):
            leituras.append(le_arquivo(arquivo.name, arquivo.getvalue()))
            barra.progress(i / len(uploaded_files), text=f"Lendo {arquivo.name}")
        barra.empty()
        st.session_state["leituras"] = leituras
        st.session_state["lote"] = lote

    leituras: list[LeituraArquivo] = st.session_state["leituras"]
    validos = [l for l in leituras if l.ok]
    falhas = [l for l in leituras if not l.ok]

    # --- Painel de leitura -------------------------------------------------
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Arquivos lidos", f"{len(validos)}/{len(leituras)}")
    col2.metric("Linhas", f"{sum(len(l.df) for l in validos):,}".replace(",", "."))
    linhas_cabecalho = sorted({(l.header_row or 0) + 1 for l in validos})
    col3.metric("Linha do cabeçalho", ", ".join(map(str, linhas_cabecalho)) or "—")
    total_geral = sum(l.total_royalty or 0.0 for l in validos)
    col4.metric("Total a pagar", formata_brl(total_geral))

    if falhas:
        with st.expander(f"⚠️ {len(falhas)} arquivo(s) não puderam ser lidos", expanded=True):
            for leitura in falhas:
                st.error(f"**{leitura.nome}** — {leitura.erro}")

    avisos = [(l.nome, a) for l in leituras for a in l.avisos]
    if avisos:
        with st.expander(f"ℹ️ {len(avisos)} aviso(s)"):
            for nome, aviso in avisos:
                st.warning(f"**{nome}** — {aviso}")

    if not validos:
        st.stop()

    aba_validacao, aba_totais, aba_concat = st.tabs(
        ["🔍 Validação de schema", "💰 Totais", "📦 Concatenar"]
    )

    # --- Validação ---------------------------------------------------------
    with aba_validacao:
        relatorio = compara_schemas(leituras)
        divergentes = relatorio[relatorio["Idêntico à referência"] == "Não"]
        if divergentes.empty:
            st.success(
                f"Todos os {len(validos)} arquivos têm exatamente as mesmas colunas, "
                "na mesma ordem. Concatenação segura."
            )
        else:
            st.warning(
                f"{len(divergentes)} arquivo(s) divergem do schema de referência "
                f"(**{validos[0].nome}**). A concatenação vai alinhar por nome de coluna "
                "e preencher as ausentes com vazio."
            )
        st.dataframe(relatorio, width="stretch", hide_index=True)

    # --- Totais ------------------------------------------------------------
    with aba_totais:
        resumo = resumo_por_arquivo(leituras)
        total = float(resumo["Total"].fillna(0).sum())

        pct = st.number_input(
            "Desconto R3 (%)", min_value=0.0, max_value=100.0,
            value=2.5, step=0.5, format="%.2f",
        )
        desconto = round(total * pct / 100, 2)

        exibicao = resumo.copy()
        exibicao["Total"] = exibicao["Total"].apply(formata_brl)
        exibicao.loc[len(exibicao)] = ["TOTAL", resumo["Linhas"].sum(), "", formata_brl(total)]
        st.dataframe(exibicao, width="stretch", hide_index=True)

        c1, c2, c3 = st.columns(3)
        c1.metric("Total bruto", formata_brl(total))
        c2.metric(f"Desconto R3 ({pct:.2f}%)", formata_brl(desconto))
        c3.metric("Total com desconto", formata_brl(total - desconto))

        st.download_button(
            "📥 Baixar resumo (CSV)",
            data=resumo.to_csv(index=False, sep=";", decimal=",").encode("utf-8-sig"),
            file_name="resumo_totais.csv",
            mime="text/csv",
            key="dl_resumo",
        )

    # --- Concatenação ------------------------------------------------------
    with aba_concat:
        consolidado = concatena(leituras)
        diagnostico = diagnostica_duplicidades(consolidado)

        st.success(
            f"{len(validos)} arquivo(s) · "
            f"{len(consolidado):,}".replace(",", ".")
            + f" linhas · {len(consolidado.columns)} colunas"
        )

        dupes = diagnostico.get("linhas_duplicadas", 0)
        if dupes:
            st.warning(f"{dupes} linha(s) idêntica(s) encontrada(s) no consolidado.")
        else:
            st.info("Nenhuma linha duplicada.")

        repetidos = diagnostico.get("statements_em_varios_arquivos", [])
        if repetidos:
            st.warning(
                "Statement(s) presente(s) em mais de um arquivo — possível dupla contagem: "
                + ", ".join(map(str, repetidos[:10]))
                + ("..." if len(repetidos) > 10 else "")
            )

        st.dataframe(consolidado.head(100), width="stretch", hide_index=True)
        st.caption("Prévia das primeiras 100 linhas.")

        st.download_button(
            "📥 Baixar arquivo concatenado (XLSX)",
            data=gera_excel(consolidado, resumo_por_arquivo(leituras)),
            file_name="backoffice_concatenado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_concat",
        )


if __name__ == "__main__":
    main()
