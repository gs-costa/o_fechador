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
from convert_nfe_xml_to_xlsx import ConversionResult, parse_marketplaces, write_consolidated_xlsx
from report_common import deduplicate_products, num

DEFAULT_INPUT_DIR = Path("NFs")
PREVIEW_COLUMNS = {
    "invoices": [
        "market_place",
        "numero",
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
        preview.append({column: row.get(column, "") for column in columns if column in row})
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
) -> bytes:
    buffer = io.BytesIO()
    build_report_from_data(
        result.invoices,
        result.items,
        buffer,
        bling_path=bling_path,
        marketplace_paths=marketplace_paths,
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


def _render_sidebar() -> Path | None:
    st.sidebar.title("O Fechador")
    st.sidebar.caption("Pré-visualize NF-e e exporte o relatório de fechamento.")

    default_path = str(DEFAULT_INPUT_DIR.resolve()) if DEFAULT_INPUT_DIR.exists() else ""
    input_path = st.sidebar.text_input(
        "Pasta com XMLs das NF-e",
        value=default_path,
        help="Pasta raiz com subpastas por marketplace, ou pasta única com arquivos .xml",
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
        st.caption("Lista deduplicada por código + EAN. Os custos são preenchidos no XLSX exportado.")
        st.dataframe(products, use_container_width=True, hide_index=True)

    with tab_erros:
        st.subheader("Arquivos ignorados")
        if result.errors:
            for error in result.errors:
                st.warning(error)
        else:
            st.success("Nenhum arquivo com erro.")


def _render_marketplace_sources(
    result: ConversionResult,
) -> tuple[Path | None, dict[str, Path]]:
    marketplaces = sorted(
        {
            str(invoice.get("market_place", ""))
            for invoice in result.invoices
            if invoice.get("market_place")
        }
    )

    with st.expander("Planilhas para taxas de marketplace", expanded=True):
        st.caption(
            "Informe o relatório do Bling e uma planilha para cada marketplace. "
            "Cada pedido é procurado em todas as planilhas informadas, mesmo que a "
            "pasta da NF-e misture marketplaces. Campos vazios geram taxa de "
            "marketplace igual a R$ 0."
        )
        bling_value = st.text_input(
            "Caminho do relatório do Bling",
            key="bling_report_path",
            placeholder=r"C:\caminho\relatorio_bling.xls",
            help=(
                "A planilha deve conter Número e Número do pedido multiloja para "
                "relacionar cada NF-e ao pedido."
            ),
        ).strip()

        marketplace_paths: dict[str, Path] = {}
        for marketplace in marketplaces:
            value = st.text_input(
                f"Planilha de taxas — {marketplace}",
                key=f"marketplace_report_path::{marketplace}",
                placeholder=r"C:\caminho\planilha_marketplace.xlsx",
                help=(
                    "O formato é detectado pelas colunas. Shopee (aba Renda): Ver, "
                    "Preço do produto e Quantia total lançada (R$). Mercado Livre: "
                    "Receita por produtos e Total (BRL)."
                ),
            ).strip()
            if value:
                marketplace_paths[marketplace] = Path(value).expanduser()

        if marketplace_paths and not bling_value:
            st.warning(
                "Sem o relatório do Bling não é possível relacionar as NF-e aos "
                "pedidos; as taxas permanecerão em R$ 0."
            )

    bling_path = Path(bling_value).expanduser() if bling_value else None
    return bling_path, marketplace_paths


def _render_export(
    result: ConversionResult,
    *,
    bling_path: Path | None,
    marketplace_paths: dict[str, Path],
) -> None:
    st.subheader("Exportar")
    st.caption("Baixe os arquivos somente quando estiver satisfeito com a pré-visualização.")

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
        st.info("Informe a pasta com os XMLs na barra lateral e clique em **Carregar dados**.")
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
    bling_path, marketplace_paths = _render_marketplace_sources(result)
    st.divider()
    _render_preview(result)
    st.divider()
    _render_export(
        result,
        bling_path=bling_path,
        marketplace_paths=marketplace_paths,
    )


if __name__ == "__main__":
    main()
