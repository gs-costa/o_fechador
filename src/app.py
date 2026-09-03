"""Streamlit UI to preview NF-e data and export XLSX reports."""

from __future__ import annotations

import io
import sys
from collections import defaultdict
from datetime import datetime
from pathlib import Path

import streamlit as st

sys.path.insert(0, str(Path(__file__).resolve().parent))

from build_report import build_report_from_data
from convert_nfe_xml_to_xlsx import (
    ConversionResult,
    parse_marketplaces,
    write_consolidated_xlsx,
)
from marketplace_fees import FEE_SOURCES
from report_common import deduplicate_products, num

DEFAULT_INPUT_DIR = Path("NFs")
FOLDER_PATH_KEY = "nfe_folder_path"
PREVIEW_COLUMNS = {
    "invoices": [
        "market_place",
        "numero",
        "xped",
        "data_emissao",
        "valor_produtos",
        "valor_frete",
        "valor_desconto",
        "valor_nf",
        "status",
        "arquivo",
    ],
    "items": [
        "market_place",
        "numero_nf",
        "codigo",
        "ean",
        "descricao",
        "quantidade",
        "valor_unitario",
        "valor_total",
    ],
    "products": ["codigo", "ean", "descricao"],
}


def _marketplace_summary(result: ConversionResult) -> list[dict[str, object]]:
    totals: dict[str, dict[str, float | int]] = defaultdict(
        lambda: {"notas": 0, "faturado": 0.0, "frete": 0.0, "desconto": 0.0}
    )
    for invoice in result.invoices:
        mp = str(invoice.get("market_place", ""))
        bucket = totals[mp]
        bucket["notas"] += 1
        bucket["faturado"] += num(invoice.get("valor_produtos"))
        bucket["frete"] += num(invoice.get("valor_frete"))
        bucket["desconto"] += num(invoice.get("valor_desconto"))

    return [
        {
            "Marketplace": mp,
            "Notas": int(values["notas"]),
            "Faturado": values["faturado"],
            "Frete": values["frete"],
            "Desconto": values["desconto"],
        }
        for mp, values in sorted(totals.items())
    ]


def _preview_rows(
    rows: list[dict[str, object]],
    columns: list[str],
) -> list[dict[str, object]]:
    preview: list[dict[str, object]] = []
    for row in rows:
        preview.append(
            {column: row.get(column, "") for column in columns if column in row}
        )
    return preview


def _export_nfe_bytes(result: ConversionResult) -> bytes:
    buffer = io.BytesIO()
    write_consolidated_xlsx(result.invoices, result.items, buffer)
    buffer.seek(0)
    return buffer.getvalue()


def _export_report_bytes(
    result: ConversionResult,
    *,
    bling_path: Path | None,
    marketplace_paths: dict[str, Path],
    regime_especial: bool,
) -> bytes:
    buffer = io.BytesIO()
    build_report_from_data(
        result.invoices,
        result.items,
        buffer,
        bling_path=bling_path,
        marketplace_paths=marketplace_paths,
        regime_especial=regime_especial,
    )
    buffer.seek(0)
    return buffer.getvalue()


def _configure_page() -> None:
    st.set_page_config(
        page_title="O Fechador",
        page_icon="📊",
        layout="wide",
        initial_sidebar_state="expanded",
    )


def _dialog_start_kwargs(initial: str) -> dict[str, str]:
    if not initial:
        return {}
    start = Path(initial).expanduser()
    if start.is_file():
        return {"initialdir": str(start.parent)}
    if start.is_dir():
        return {"initialdir": str(start)}
    if start.parent.is_dir():
        return {"initialdir": str(start.parent)}
    return {}


def _pick_with_tkinter(picker: str, *, initial: str = "") -> str:
    try:
        import tkinter as tk
        from tkinter import filedialog
    except ImportError:
        return ""

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    kwargs = _dialog_start_kwargs(initial)
    if picker == "directory":
        chosen = filedialog.askdirectory(**kwargs)
    else:
        chosen = filedialog.askopenfilename(
            filetypes=[
                ("Planilhas Excel", "*.xlsx *.xlsm *.xls"),
                ("Excel", "*.xlsx"),
                ("Excel 97-2003", "*.xls"),
                ("Todos os arquivos", "*.*"),
            ],
            **kwargs,
        )
    root.destroy()
    return chosen


def _pick_local_directory(*, initial: str = "") -> str:
    """Open the native folder dialog on the machine running Streamlit."""
    return _pick_with_tkinter("directory", initial=initial)


def _pick_local_file(*, initial: str = "") -> str:
    """Open the native file dialog on the machine running Streamlit."""
    return _pick_with_tkinter("file", initial=initial)


def _assign_picked_file(key: str) -> None:
    selected = _pick_local_file(initial=str(st.session_state.get(key, "")))
    if selected:
        st.session_state[key] = selected


def _render_sidebar() -> Path | None:
    st.sidebar.title("O Fechador")
    st.sidebar.caption("Pré-visualize NF-e e exporte o relatório de fechamento.")

    if FOLDER_PATH_KEY not in st.session_state:
        st.session_state[FOLDER_PATH_KEY] = (
            str(DEFAULT_INPUT_DIR.resolve()) if DEFAULT_INPUT_DIR.exists() else ""
        )

    if st.sidebar.button("Selecionar pasta", use_container_width=True):
        selected = _pick_local_directory(initial=st.session_state[FOLDER_PATH_KEY])
        if selected:
            st.session_state[FOLDER_PATH_KEY] = selected
            st.rerun()

    input_path = st.sidebar.text_input(
        "Pasta com XMLs das NF-e",
        key=FOLDER_PATH_KEY,
        help=(
            "Pasta raiz com subpastas por marketplace, ou pasta única com "
            "arquivos .xml. Use o botão acima para escolher no explorador."
        ),
    )

    if st.sidebar.button("Carregar dados", type="primary", use_container_width=True):
        folder = Path(input_path).expanduser()
        if not folder.is_dir():
            st.sidebar.error(f"Pasta não encontrada: {folder}")
            return None
        try:
            st.session_state["result"] = parse_marketplaces(folder)
            st.session_state["source_folder"] = str(folder.resolve())
        except FileNotFoundError as exc:
            st.sidebar.error(str(exc))
        except Exception as exc:  # noqa: BLE001 - show parsing errors in the UI
            st.sidebar.error(f"Erro ao ler XMLs: {exc}")

    if "source_folder" in st.session_state:
        st.sidebar.success("Dados carregados")
        st.sidebar.caption(st.session_state["source_folder"])

    return Path(input_path).expanduser() if input_path else None


def _render_metrics(result: ConversionResult) -> None:
    products = deduplicate_products(result.items)
    marketplaces = sorted({str(inv.get("market_place", "")) for inv in result.invoices})

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Notas fiscais", len(result.invoices))
    c2.metric("Itens", len(result.items))
    c3.metric("Produtos únicos", len(products))
    c4.metric("Marketplaces", len(marketplaces))


def _render_preview(result: ConversionResult) -> None:
    summary = _marketplace_summary(result)
    products = deduplicate_products(result.items)

    tab_resumo, tab_notas, tab_itens, tab_produtos, tab_erros = st.tabs(
        ["Resumo", "Notas Fiscais", "Itens", "Produtos", "Erros"]
    )

    with tab_resumo:
        st.subheader("Resumo por marketplace")
        st.dataframe(summary, use_container_width=True, hide_index=True)
        if result.processed:
            st.caption(" / ".join(result.processed))

    with tab_notas:
        st.subheader("Notas fiscais")
        st.dataframe(
            _preview_rows(result.invoices, PREVIEW_COLUMNS["invoices"]),
            use_container_width=True,
            hide_index=True,
        )

    with tab_itens:
        st.subheader("Itens das notas")
        st.dataframe(
            _preview_rows(result.items, PREVIEW_COLUMNS["items"]),
            use_container_width=True,
            hide_index=True,
        )

    with tab_produtos:
        st.subheader("Produtos para custos")
        st.caption(
            "Lista deduplicada por código + EAN. Os custos são preenchidos no XLSX exportado."
        )
        st.dataframe(products, use_container_width=True, hide_index=True)

    with tab_erros:
        st.subheader("Arquivos ignorados")
        if result.errors:
            for error in result.errors:
                st.warning(error)
        else:
            st.success("Nenhum arquivo com erro.")


def _render_file_path_input(
    label: str,
    *,
    key: str,
    help_text: str,
    placeholder: str,
) -> str:
    path_col, button_col = st.columns([4, 1], vertical_alignment="bottom")
    with button_col:
        st.button(
            "Selecionar",
            key=f"{key}::pick",
            on_click=_assign_picked_file,
            args=(key,),
            use_container_width=True,
        )
    with path_col:
        return st.text_input(
            label,
            key=key,
            placeholder=placeholder,
            help=help_text,
        ).strip()


def _render_marketplace_sources() -> tuple[Path | None, dict[str, Path]]:
    with st.expander("Planilhas para taxas de marketplace", expanded=True):
        st.caption(
            "Informe o relatório do Bling e as planilhas de taxas disponíveis. "
            "Cada pedido é procurado em todas as planilhas informadas. Campos "
            "vazios geram taxa de marketplace igual a R$ 0."
        )
        bling_value = _render_file_path_input(
            "Caminho do relatório do Bling",
            key="bling_report_path",
            placeholder=r"C:\caminho\relatorio_bling.xls",
            help_text=(
                "A planilha deve conter Número e Número do pedido multiloja para "
                "relacionar cada NF-e ao pedido."
            ),
        )

        marketplace_values: dict[str, str] = {}
        for source in FEE_SOURCES:
            marketplace_values[source.key] = _render_file_path_input(
                f"Planilha de taxas — {source.label}",
                key=f"marketplace_report_path::{source.key}",
                placeholder=r"C:\caminho\planilha_marketplace.xlsx",
                help_text=source.help_text,
            )

        if any(marketplace_values.values()) and not bling_value:
            st.warning(
                "Sem o relatório do Bling, somente as NF-e que "
                "contêm o número do pedido poderão ser relacionadas; "
                "as demais taxas permanecerão em R$ 0."
            )

    bling_path = Path(bling_value).expanduser() if bling_value else None
    marketplace_paths = {
        key: Path(value).expanduser()
        for key, value in marketplace_values.items()
        if value
    }
    return bling_path, marketplace_paths


def _render_fiscal_options() -> bool:
    with st.expander("Opções fiscais", expanded=True):
        regime_especial = st.toggle(
            "Regime Especial",
            value=False,
            help=(
                "Quando ativo, vendas destinadas para fora de MG usam 1,3% da base "
                "de ICMS (vBC) no lugar do ICMS destacado na NF-e."
            ),
        )
        if regime_especial:
            st.success(
                "Regime Especial ATIVO — fora de MG: ICMS = 1,3% × base de ICMS."
            )
        else:
            st.info("Regime Especial INATIVO — será usado o ICMS destacado na NF-e.")
    return regime_especial


def _render_export(
    result: ConversionResult,
    *,
    bling_path: Path | None,
    marketplace_paths: dict[str, Path],
    regime_especial: bool,
) -> None:
    st.subheader("Exportar")
    st.caption(
        "Baixe os arquivos somente quando estiver satisfeito com a pré-visualização."
    )

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    col_nfe, col_report = st.columns(2)

    with col_nfe:
        st.markdown("**Base NF-e**")
        st.write("Abas: Invoices, Items")
        st.download_button(
            label="Baixar nfe_invoices.xlsx",
            data=_export_nfe_bytes(result),
            file_name=f"nfe_invoices_{timestamp}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    with col_report:
        st.markdown("**Relatório de fechamento**")
        st.write(
            "Abas: Resumo, Detalhado, Custos Produtos, Items, Parâmetros, DRE, "
            "Conciliação"
        )
        try:
            report_data = _export_report_bytes(
                result,
                bling_path=bling_path,
                marketplace_paths=marketplace_paths,
                regime_especial=regime_especial,
            )
        except (FileNotFoundError, ValueError) as exc:
            st.error(str(exc))
        else:
            st.download_button(
                label="Baixar relatorio_fechamento.xlsx",
                data=report_data,
                file_name=f"relatorio_fechamento_{timestamp}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )


def main() -> None:
    _configure_page()
    st.title("Pré-visualização do Fechamento")
    st.write(
        "Carregue os XMLs das NF-e, revise os dados nas abas abaixo e exporte o XLSX quando estiver tudo certo."
    )

    _render_sidebar()

    result: ConversionResult | None = st.session_state.get("result")
    if result is None:
        st.info(
            "Selecione a pasta com os XMLs na barra lateral e clique em **Carregar dados**."
        )
        with st.expander("Estrutura esperada da pasta"):
            st.code(
                "NFs/\n"
                "  Mercado Livre/\n"
                "    nota1.xml\n"
                "    nota2.xml\n"
                "  Shopee/\n"
                "    nota3.xml",
                language="text",
            )
        return

    _render_metrics(result)
    st.divider()
    bling_path, marketplace_paths = _render_marketplace_sources()
    st.divider()
    regime_especial = _render_fiscal_options()
    st.divider()
    _render_preview(result)
    st.divider()
    _render_export(
        result,
        bling_path=bling_path,
        marketplace_paths=marketplace_paths,
        regime_especial=regime_especial,
    )


if __name__ == "__main__":
    main()
