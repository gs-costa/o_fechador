"""Build a marketplace closing report (Fechamento) from the NF-e result workbook.

Reads the workbook produced by ``convert_nfe_xml_to_xlsx.py`` (sheets ``Invoices``
and ``Items``) and writes a new workbook with:

  - Resumo:     one clean grid, metrics as rows, marketplaces as columns + Total.
  - Detalhado:  every column from the Invoices sheet (kept verbatim) followed by the
                appended calculated columns (live Excel formulas).
  - Parametros: editable rates per marketplace (ADS / Afiliado / CANDARU) and global
                rates (IRPJ/CSLL, impostos sobre vendas). Change a rate and the
                whole report recalculates.
  - Items:      line items from the NF-e workbook, with custo_unitario / custo_total.
  - Custos Produtos: deduplicated product list (codigo + ean) for manual cost entry.
  - DRE:        Demonstração do Resultado do Exercício (see ``build_dre.py``).

Only values that can be derived from the NF-e are included. KPIs that depend on
external data (Depositado, Custo Produto, Liquido, Margem, Frete Primario) are not
produced. The appended columns are estimates computed from NF-e values:

  ADS       = taxa_ads(marketplace)      * Faturado (valor_produtos)
  Afiliado  = taxa_afiliado(marketplace) * Faturado (valor_produtos)
  IRPJ/CSLL = taxa_irpj                   * Base ICMS (valor_icms_bc)
  CANDARU   = taxa_candaru(marketplace) * valor_base_comissao
              (3% when market_place contains AMAZON B2B or TIKTOK SHOP; 7% otherwise)
"""

from __future__ import annotations

import argparse
import sys
from datetime import datetime
from pathlib import Path
from typing import BinaryIO

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from build_dre import build_dre
from build_items import build_custos_produtos, build_items
from report_common import (
    ADDED_COLUMNS,
    BORDER,
    CANDARU_RATE_DEFAULT,
    CUSTOS_SHEET,
    DEFAULT_ADS_RATES,
    DEFAULT_AFILIADO_RATES,
    DEFAULT_IRPJ_RATE,
    DEFAULT_SALES_TAX_RATE,
    DET_WIDTHS,
    INT_FMT,
    ITEM_HEADERS,
    ITEMS_SHEET,
    LABEL_FONT,
    MONEY_COLUMNS,
    MONEY_FMT,
    NOTE_FONT,
    PCT_FMT,
    SRC_BASE_COMISSAO,
    SRC_BASE_ICMS,
    SRC_COFINS,
    SRC_DIFAL,
    SRC_FATURADO,
    SRC_ICMS,
    SRC_MARKETPLACE,
    SRC_PIS,
    TEXT_COLUMNS,
    TITLE_FONT,
    TOTAL_FILL,
    as_cell,
    candaru_rate_for,
    cfop_by_invoice,
    deduplicate_products,
    load_result,
    num,
    style_header_cell,
)

# The Detalhado sheet keeps every column from the Invoices sheet (same names and
# order) and then appends the calculated columns below.


def build_parametros(
    ws: Worksheet,
    marketplaces: list[str],
    ads_rates: dict[str, float],
    afiliado_rates: dict[str, float],
    candaru_rates: dict[str, float],
) -> dict[str, str]:
    ws["A1"] = "Taxas por Marketplace"
    ws["A1"].font = TITLE_FONT
    for col, name in zip(
        "ABCD", ("Marketplace", "Taxa ADS", "Taxa Afiliado", "Taxa CANDARU")
    ):
        cell = ws[f"{col}2"]
        cell.value = name
        style_header_cell(cell)

    for idx, mp in enumerate(marketplaces, start=3):
        ws[f"A{idx}"] = mp
        ws[f"B{idx}"] = ads_rates.get(mp, 0.0)
        ws[f"C{idx}"] = afiliado_rates.get(mp, 0.0)
        ws[f"D{idx}"] = candaru_rates.get(mp, CANDARU_RATE_DEFAULT)
        ws[f"B{idx}"].number_format = PCT_FMT
        ws[f"C{idx}"].number_format = PCT_FMT
        ws[f"D{idx}"].number_format = PCT_FMT

    ws["F2"] = "Taxa IRPJ/CSLL"
    ws["F2"].font = LABEL_FONT
    ws["G2"] = DEFAULT_IRPJ_RATE
    ws["G2"].number_format = PCT_FMT

    ws["F3"] = "Taxa Impostos sobre Vendas"
    ws["F3"].font = LABEL_FONT
    ws["G3"] = DEFAULT_SALES_TAX_RATE
    ws["G3"].number_format = PCT_FMT

    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 14
    ws.column_dimensions["F"].width = 26
    ws.column_dimensions["G"].width = 12
    ws.freeze_panes = "A3"
    return {"irpj": "$G$2", "sales_tax": "$G$3"}


def build_detalhado(
    ws: Worksheet,
    invoices: list[dict[str, object]],
    invoice_headers: list[str],
    invoice_cfops: dict[str, str],
    global_rate_cells: dict[str, str],
) -> tuple[dict[str, str], int]:
    """Write the Detalhado sheet and return (column_name -> letter, last_row).

    All Invoices columns are kept verbatim; the calculated columns are appended.
    """
    columns = list(invoice_headers) + ADDED_COLUMNS
    det_col = {name: get_column_letter(i) for i, name in enumerate(columns, start=1)}

    for col_idx, name in enumerate(columns, start=1):
        style_header_cell(ws.cell(row=1, column=col_idx, value=name))

    irpj = f"Parametros!{global_rate_cells['irpj']}"

    ordered = sorted(invoices, key=lambda inv: num(inv.get("numero")))
    for offset, inv in enumerate(ordered):
        r = offset + 2
        chave = str(inv.get("chave_nfe", ""))

        for name in invoice_headers:
            ws[f"{det_col[name]}{r}"] = as_cell(inv.get(name, ""))

        ws[f"{det_col['CFOP']}{r}"] = invoice_cfops.get(chave, "")

        fat = f"{det_col[SRC_FATURADO]}{r}"
        base_comissao = f"{det_col[SRC_BASE_COMISSAO]}{r}"
        base = f"{det_col[SRC_BASE_ICMS]}{r}"
        mp = f"{det_col[SRC_MARKETPLACE]}{r}"
        ws[f"{det_col['ADS']}{r}"] = f"=VLOOKUP({mp},Parametros!$A:$C,2,FALSE)*{fat}"
        ws[f"{det_col['Afiliado']}{r}"] = (
            f"=VLOOKUP({mp},Parametros!$A:$C,3,FALSE)*{fat}"
        )
        ws[f"{det_col['IRPJ/CSLL']}{r}"] = f"={irpj}*{base}"
        ws[f"{det_col['CANDARU']}{r}"] = (
            f"=VLOOKUP({mp},Parametros!$A:$D,4,FALSE)*{base_comissao}"
        )

    last_row = len(invoices) + 1
    for name in columns:
        col = det_col[name]
        if name in MONEY_COLUMNS:
            for r in range(2, last_row + 1):
                ws[f"{col}{r}"].number_format = MONEY_FMT
        elif name in TEXT_COLUMNS:
            for r in range(2, last_row + 1):
                ws[f"{col}{r}"].number_format = "@"

    for name in columns:
        ws.column_dimensions[det_col[name]].width = DET_WIDTHS.get(name, 13)
    ws.freeze_panes = "A2"
    return det_col, last_row


def build_resumo(
    ws: Worksheet,
    marketplaces: list[str],
    det_col: dict[str, str],
) -> None:
    ws["A1"] = "Fechamento - Resumo por Marketplace"
    ws["A1"].font = TITLE_FONT

    n_mp = len(marketplaces)
    label_col = 1
    first_mp_col = 2
    last_mp_col = first_mp_col + n_mp - 1
    total_col = last_mp_col + 1
    pct_col = total_col + 1

    header_row = 3
    style_header_cell(ws.cell(row=header_row, column=label_col, value="Metrica"))
    for i, mp in enumerate(marketplaces):
        style_header_cell(ws.cell(row=header_row, column=first_mp_col + i, value=mp))
    style_header_cell(ws.cell(row=header_row, column=total_col, value="Total"))
    style_header_cell(ws.cell(row=header_row, column=pct_col, value="% s/ Faturado"))

    det = "Detalhado"
    fcol = det_col[SRC_MARKETPLACE]

    metrics = [
        ("Notas Fiscais", "count", None, INT_FMT),
        ("Faturado", "sum", det_col[SRC_BASE_COMISSAO], MONEY_FMT),
        ("ICMS", "sum", det_col[SRC_ICMS], MONEY_FMT),
        ("DIFAL", "sum", det_col[SRC_DIFAL], MONEY_FMT),
        ("PIS", "sum", det_col[SRC_PIS], MONEY_FMT),
        ("COFINS", "sum", det_col[SRC_COFINS], MONEY_FMT),
        ("IRPJ/CSLL", "sum", det_col["IRPJ/CSLL"], MONEY_FMT),
        ("ADS", "sum", det_col["ADS"], MONEY_FMT),
        ("Afiliado", "sum", det_col["Afiliado"], MONEY_FMT),
        ("CANDARU", "sum", det_col["CANDARU"], MONEY_FMT),
    ]

    row = header_row + 1
    faturado_row = row + 1

    for label, kind, metric_col, fmt in metrics:
        label_cell = ws.cell(row=row, column=label_col, value=label)
        label_cell.font = LABEL_FONT
        label_cell.border = BORDER

        for i, _mp in enumerate(marketplaces):
            c = first_mp_col + i
            mp_ref = f"{get_column_letter(c)}${header_row}"
            crit = f"{det}!${fcol}:${fcol}"
            if kind == "count":
                formula = f"=COUNTIF({crit},{mp_ref})"
            else:
                rng = f"{det}!${metric_col}:${metric_col}"
                formula = f"=SUMIF({crit},{mp_ref},{rng})"
            cell = ws.cell(row=row, column=c, value=formula)
            cell.number_format = fmt
            cell.border = BORDER

        first = f"{get_column_letter(first_mp_col)}{row}"
        last = f"{get_column_letter(last_mp_col)}{row}"
        total_cell = ws.cell(row=row, column=total_col, value=f"=SUM({first}:{last})")
        total_cell.number_format = fmt
        total_cell.font = LABEL_FONT
        total_cell.fill = TOTAL_FILL
        total_cell.border = BORDER

        pct_formula = None
        if kind == "sum":
            fat_total = f"{get_column_letter(total_col)}{faturado_row}"
            pct_formula = f"=IFERROR({get_column_letter(total_col)}{row}/{fat_total},0)"
        pct_cell = ws.cell(row=row, column=pct_col, value=pct_formula)
        if pct_formula is not None:
            pct_cell.number_format = PCT_FMT
        pct_cell.border = BORDER
        row += 1

    for i in range(n_mp + 3):
        col = get_column_letter(label_col + i)
        ws.column_dimensions[col].width = 20 if i == 0 else 15
    ws.freeze_panes = ws.cell(row=header_row + 1, column=first_mp_col).coordinate

    note_row = row + 2
    ws.cell(
        row=note_row,
        column=label_col,
        value=(
            "ADS, Afiliado, IRPJ/CSLL e CANDARU sao estimativas (taxa x valor da NF-e); "
            "ajuste as taxas na aba Parametros. Os demais valores vem direto das NF-e."
        ),
    )
    ws.cell(row=note_row, column=label_col).font = NOTE_FONT


def build_report_from_data(
    invoices: list[dict[str, object]],
    items: list[dict[str, object]],
    output: Path | BinaryIO,
) -> dict[str, object]:
    if not invoices:
        raise ValueError("No invoices found in the input data.")

    invoice_headers = [h for h in invoices[0].keys() if h]
    item_headers = [h for h in items[0].keys() if h] if items else list(ITEM_HEADERS)
    invoice_cfops = cfop_by_invoice(items)

    marketplaces = sorted(
        {
            str(inv.get("market_place", ""))
            for inv in invoices
            if inv.get("market_place")
        }
    )

    ads_rates = dict(DEFAULT_ADS_RATES)
    afiliado_rates = dict(DEFAULT_AFILIADO_RATES)
    candaru_rates = {mp: candaru_rate_for(mp) for mp in marketplaces}

    wb = Workbook()
    ws_resumo = wb.active
    assert ws_resumo is not None
    ws_resumo.title = "Resumo"
    ws_det = wb.create_sheet("Detalhado")
    ws_custos = wb.create_sheet(CUSTOS_SHEET)
    ws_items = wb.create_sheet(ITEMS_SHEET)
    ws_par = wb.create_sheet("Parametros")
    ws_dre = wb.create_sheet("DRE")

    global_rate_cells = build_parametros(
        ws_par, marketplaces, ads_rates, afiliado_rates, candaru_rates
    )
    det_col, last_row = build_detalhado(
        ws_det, invoices, invoice_headers, invoice_cfops, global_rate_cells
    )
    custos_col, _ = build_custos_produtos(ws_custos, items)
    items_col, items_last_row = build_items(ws_items, items, item_headers, custos_col)
    build_resumo(ws_resumo, marketplaces, det_col)
    build_dre(
        ws_dre,
        det_col,
        last_row,
        sales_tax_rate_cell=f"Parametros!{global_rate_cells['sales_tax']}",
        items_col=items_col,
        items_last_row=items_last_row,
    )

    if isinstance(output, Path):
        output.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output)

    return {
        "invoices": len(invoices),
        "items": len(items),
        "products": len(deduplicate_products(items)) if items else 0,
        "marketplaces": marketplaces,
    }


def build_report(
    input_path: Path,
    output_path: Path,
) -> dict[str, object]:
    invoices, items = load_result(input_path)
    return build_report_from_data(invoices, items, output_path)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--input",
        type=Path,
        required=True,
        help="Result workbook from convert_nfe_xml_to_xlsx.py (Invoices/Items sheets)",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=None,
        help="Output report path (default: exports/relatorio_fechamento_<timestamp>.xlsx)",
    )
    return parser


def main() -> None:
    args = build_parser().parse_args()
    if not args.input.is_file():
        print(f"Error: input workbook not found: {args.input}")
        sys.exit(1)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = (
        args.output or Path("exports") / f"relatorio_fechamento_{timestamp}.xlsx"
    )

    info = build_report(args.input, output_path)

    marketplaces = info["marketplaces"]
    marketplaces_text = (
        ", ".join(marketplaces) if isinstance(marketplaces, list) else ""
    )
    print(f"Report built from {info['invoices']} invoices ({info['items']} items).")
    print(f"Products in cost sheet: {info.get('products', 0)}")
    print(f"Marketplaces: {marketplaces_text}")
    print(f"Report saved to: {output_path.resolve()}")


if __name__ == "__main__":
    main()
