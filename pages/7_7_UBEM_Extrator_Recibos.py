"""
Extrator de Recibos UBEM -> Base XLSX
=====================================
Faz upload de recibos .docx da UBEM, extrai a tabela (colada como imagem EMF),
relaciona cada linha com o respectivo relatorio DIST listado em
"EMPRESAS E PERIODOS", empilha tudo e entrega um .xlsx consolidado.

Executar:
    pip install streamlit openpyxl
    streamlit run app_recibos_ubem.py
"""

from __future__ import annotations

import html
import io
import re
import struct
import zipfile
from dataclasses import dataclass, field

import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

# --------------------------------------------------------------------------- #
# Configuracao
# --------------------------------------------------------------------------- #

COLUNAS = [
    "STATEMENT NAME",
    "PERÍODO",
    "SUBMITTER",
    "PAYEE",
    "SOURCE",
    "AMOUNT_TO_BE_PAID",
    "ROYALTY STATEMENT NUMBER",
    "ARQUIVO",
    "DATA RECIBO",
]

LARGURAS = [62, 20, 14, 14, 14, 20, 16, 34, 14]

MESES = {
    "janeiro": 1, "fevereiro": 2, "março": 3, "marco": 3, "abril": 4,
    "maio": 5, "junho": 6, "julho": 7, "agosto": 8, "setembro": 9,
    "outubro": 10, "novembro": 11, "dezembro": 12,
}

EMR_EXTTEXTOUTW = 84


# --------------------------------------------------------------------------- #
# Parsing
# --------------------------------------------------------------------------- #

def emf_extrair_linhas(blob: bytes) -> list[list[str]]:
    """Le os registros EMR_EXTTEXTOUTW de um EMF e reconstroi as linhas da tabela.

    O texto vetorial do metafile guarda a posicao (x, y) de cada fragmento.
    Agrupamos por y (linha) e ordenamos por x (coluna).
    """
    off, itens = 0, []
    n_total = len(blob)
    while off < n_total - 8:
        tipo, tam = struct.unpack_from("<II", blob, off)
        if tam < 8 or off + tam > n_total:
            break
        if tipo == EMR_EXTTEXTOUTW:
            # EMRTEXT comeca em off+36: ptlReference(8) nChars(4) offString(4)
            _x, _y, n_chars, off_str = struct.unpack_from("<iiII", blob, off + 36)
            txt = blob[off + off_str: off + off_str + n_chars * 2]
            txt = txt.decode("utf-16le", "replace").strip()
            if txt:
                itens.append((_y, _x, txt))
        off += tam

    linhas: dict[int, list[tuple[int, str]]] = {}
    for y, x, s in itens:
        linhas.setdefault(y, []).append((x, s))
    return [[s for _, s in sorted(linhas[y])] for y in sorted(linhas)]


def para_numero(txt: str) -> float:
    """'8.859,51' -> 8859.51"""
    return float(txt.replace(".", "").replace(",", "."))


def texto_do_docx(z: zipfile.ZipFile) -> str:
    xml = z.read("word/document.xml").decode("utf-8")
    xml = re.sub(r"</w:p>", "\n", xml)
    return html.unescape(re.sub(r"<[^>]+>", "", xml))


def data_do_recibo(texto: str) -> str:
    m = re.search(r"Rio de Janeiro,\s*(\d{1,2})\s*de\s*(\w+)\s*de\s*(\d{4})", texto, re.I)
    if not m:
        return ""
    mes = MESES.get(m.group(2).lower(), 0)
    return f"{int(m.group(3)):04d}-{mes:02d}-{int(m.group(1)):02d}"


def total_do_recibo(texto: str) -> float | None:
    m = re.search(r"Total Recebido:\s*R\$\s*([\d\.,]+)", texto)
    return para_numero(m.group(1)) if m else None


@dataclass
class Resultado:
    arquivo: str
    linhas: list[list] = field(default_factory=list)
    total_tabela: float = 0.0
    total_recibo: float | None = None
    avisos: list[str] = field(default_factory=list)
    erro: str | None = None

    @property
    def ok(self) -> bool:
        if self.erro:
            return False
        if self.total_recibo is None:
            return False
        return abs(self.total_tabela - self.total_recibo) < 0.005 and not self.avisos


def processar_docx(nome: str, conteudo: bytes) -> Resultado:
    res = Resultado(arquivo=nome)
    try:
        z = zipfile.ZipFile(io.BytesIO(conteudo))
        texto = texto_do_docx(z)
    except Exception as exc:  # noqa: BLE001
        res.erro = f"Nao foi possivel abrir o .docx ({exc})"
        return res

    dist = [l.strip() for l in texto.split("\n") if l.strip().startswith("DIST")]
    data_rec = data_do_recibo(texto)
    res.total_recibo = total_do_recibo(texto)

    emfs = [n for n in z.namelist() if n.lower().endswith(".emf")]
    if not emfs:
        res.erro = "Nenhuma imagem EMF encontrada (a tabela precisa estar colada como objeto do Excel)."
        return res

    linhas_emf: list[list[str]] = []
    for e in emfs:
        linhas_emf.extend(emf_extrair_linhas(z.read(e)))

    # Linha de dados = SUBMITTER/PAYEE preenchidos e ultima celula = nº do statement
    dados = [
        r for r in linhas_emf
        if len(r) >= 6 and r[1].strip() and r[2].strip()
        and re.fullmatch(r"\d{5,7}", r[-1].strip())
    ]

    if not dados:
        res.erro = "Tabela localizada, mas nenhuma linha de dados reconhecida."
        return res

    if len(dados) != len(dist):
        res.erro = (
            f"Divergencia: {len(dist)} relatorios DIST x {len(dados)} linhas na tabela. "
            "O relacionamento 1:1 nao pode ser feito automaticamente."
        )
        return res

    for linha_dist, r in zip(dist, dados):
        periodo, submitter, payee, source = r[0], r[1], r[2], r[3]
        statement = r[-1].strip()
        try:
            valor = para_numero(r[-2])  # coluna imediatamente antes do nº do statement
        except ValueError:
            res.avisos.append(f"Valor ilegivel na linha '{periodo}'")
            continue

        if source.upper() not in linha_dist.upper():
            res.avisos.append(
                f'SOURCE "{source}" nao confere com "{linha_dist}"'
            )

        res.linhas.append([
            linha_dist, periodo, submitter, payee, source,
            valor, int(statement), nome, data_rec,
        ])

    res.total_tabela = sum(l[5] for l in res.linhas)
    return res


# --------------------------------------------------------------------------- #
# Saida XLSX
# --------------------------------------------------------------------------- #

def gerar_xlsx(linhas: list[list]) -> bytes:
    wb = Workbook()
    # wb.active e tipado como Optional[Worksheet]; garantimos a aba para o type checker
    ws = wb.active or wb.create_sheet()
    ws.title = "BASE"
    ws.append(COLUNAS)
    for l in linhas:
        ws.append(l)

    fonte_cab = Font(bold=True, color="FFFFFF")
    fundo_cab = PatternFill("solid", fgColor="1F4E79")
    fino = Side(style="thin", color="BFBFBF")
    borda = Border(left=fino, right=fino, top=fino, bottom=fino)

    for c in ws[1]:
        c.font = fonte_cab
        c.fill = fundo_cab
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = borda

    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.border = borda
        row[5].number_format = "#,##0.00"
        row[6].number_format = "0"

    ws.freeze_panes = "B2"
    if ws.max_row > 1:
        ws.auto_filter.ref = ws.dimensions
    for i, larg in enumerate(LARGURAS):
        ws.column_dimensions[chr(ord("A") + i)].width = larg
    ws.row_dimensions[1].height = 32

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# --------------------------------------------------------------------------- #
# Interface
# --------------------------------------------------------------------------- #

st.set_page_config(page_title="Recibos UBEM -> Base XLSX", page_icon="📄", layout="wide")


def tabela(dados) -> None:
    """st.dataframe em largura total, compativel com versoes novas e antigas."""
    try:
        st.dataframe(dados, width="stretch", hide_index=True)
    except TypeError:
        st.dataframe(dados, use_container_width=True, hide_index=True)

st.title("📄 Recibos UBEM → Base XLSX")
st.caption(
    "Envie os recibos .docx. O app le a lista de relatorios DIST, extrai a tabela "
    "colada como imagem, relaciona linha a linha e empilha tudo numa base unica."
)

arquivos = st.file_uploader(
    "Recibos (.docx)", type=["docx"], accept_multiple_files=True
)

if not arquivos:
    st.info("Aguardando arquivos.")
    st.stop()

resultados = [processar_docx(f.name, f.getvalue()) for f in arquivos]

# --- Painel de validacao ---------------------------------------------------- #
st.subheader("Validacao por arquivo")
painel = []
for r in resultados:
    if r.erro:
        status = "❌ erro"
    elif r.avisos:
        status = "⚠️ conferir"
    elif r.total_recibo is not None and abs(r.total_tabela - r.total_recibo) >= 0.005:
        status = "⚠️ total divergente"
    else:
        status = "✅ ok"
    painel.append({
        "Arquivo": r.arquivo,
        "Linhas": len(r.linhas),
        "Soma da tabela": f"{r.total_tabela:,.2f}",
        "Total Recebido (recibo)": (
            f"{r.total_recibo:,.2f}" if r.total_recibo is not None else "—"
        ),
        "Status": status,
    })
tabela(painel)

for r in resultados:
    if r.erro:
        st.error(f"**{r.arquivo}** — {r.erro}")
    for a in r.avisos:
        st.warning(f"**{r.arquivo}** — {a}")

# --- Resultado -------------------------------------------------------------- #
todas = [l for r in resultados for l in r.linhas]
if not todas:
    st.stop()

st.subheader("Base consolidada")
c1, c2, c3 = st.columns(3)
c1.metric("Arquivos processados", sum(1 for r in resultados if r.linhas))
c2.metric("Linhas", len(todas))
c3.metric("Total R$", f"{sum(l[5] for l in todas):,.2f}")

tabela([dict(zip(COLUNAS, l)) for l in todas])

st.download_button(
    "⬇️ Baixar base consolidada (.xlsx)",
    data=gerar_xlsx(todas),
    file_name="UBEM_Recibos_BASE.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
)