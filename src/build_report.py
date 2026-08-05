"""Build a marketplace closing report (Fechamento) from the NF-e result workbook.

Reads the workbook produced by ``convert_nfe_xml_to_xlsx.py`` (sheets ``Invoices``
and ``Items``) and writes a new workbook with:

  - Resumo:     one clean grid, metrics as rows, marketplaces as columns + Total.
  - Detalhado:  every column from the Invoices sheet (kept verbatim) followed by the
                appended calculated columns (live Excel formulas).
  - Parametros: editable rates per marketplace (ADS / Afiliado / CANDARU) and global
                rate (IRPJ/CSLL). Change a rate and the whole report recalculates.

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

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

# Default assumption rates (editable later in the Parametros sheet).
DEFAULT_ADS_RATES: dict[str, float] = {
    "Mercado Livre": 0.0597444848231328,
    "Mercado Livre Full": 0.0597444848231328,
    "Shopee": 0.0816910844699686,
    "Shopee Full": 0.0816910844699686,
    "TikTok": 0.0808623294375085,
}
DEFAULT_AFILIADO_RATES: dict[str, float] = {
    "Mercado Livre": 0.0102977242994446,
    "Mercado Livre Full": 0.0102977242994446,
    "Shopee": 0.031148329972608,
    "Shopee Full": 0.031148329972608,
    "TikTok": 0.0924554265462002,
}
DEFAULT_IRPJ_RATE = 0.0308
CANDARU_RATE_LOW = 0.03
CANDARU_RATE_DEFAULT = 0.07
CANDARU_LOW_MARKETPLACE_SUBSTRINGS = ("amazon b2b", "tiktok shop")

MONEY_FMT = "R$ #,##0.00"
PCT_FMT = "0.00%"
INT_FMT = "#,##0"

HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(bold=True, color="FFFFFF")
TOTAL_FILL = PatternFill("solid", fgColor="D9E1F2")
LABEL_FONT = Font(bold=True)
TITLE_FONT = Font(bold=True, size=14, color="1F4E78")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# The Detalhado sheet keeps every column from the Invoices sheet (same names and
# order) and then appends the calculated columns below.
ADDED_COLUMNS = [
    "CFOP",
    "ADS",
    "Afiliado",
    "IRPJ/CSLL",
    "CANDARU",
]

# Kept (Invoices) columns referenced by the calculated formulas.
SRC_MARKETPLACE = "market_place"
SRC_FATURADO = "valor_produtos"
SRC_BASE_COMISSAO = "valor_base_comissao"
SRC_BASE_ICMS = "valor_icms_bc"
SRC_ICMS = "valor_icms"
SRC_DIFAL = "valor_icms_difal"
SRC_PIS = "valor_pis"
SRC_COFINS = "valor_cofins"

# Columns formatted as currency (kept + appended).
MONEY_COLUMNS = {
    "valor_produtos",
    "valor_frete",
    "valor_desconto",
    "valor_nf",
    "valor_base_comissao",
    "valor_icms",
    "valor_icms_difal",
    "valor_icms_bc",
    "valor_pis",
    "valor_cofins",
    "ADS",
    "Afiliado",
    "IRPJ/CSLL",
    "CANDARU",
}
# Columns kept as text to preserve long codes / leading zeros.
TEXT_COLUMNS = {"chave_nfe", "destinatario_doc"}

# Preferred column widths on the Detalhado sheet.
DET_WIDTHS = {
    "market_place": 16,
    "chave_nfe": 36,
    "numero": 10,
    "serie": 7,
    "data_emissao": 22,
    "natureza_operacao": 20,
    "destinatario_doc": 16,
    "destinatario_uf": 8,
    "status": 22,
    "arquivo": 30,
}


def candaru_rate_for(marketplace: str) -> float:
    name = marketplace.strip().casefold()
    if any(sub in name for sub in CANDARU_LOW_MARKETPLACE_SUBSTRINGS):
        return CANDARU_RATE_LOW
    return CANDARU_RATE_DEFAULT


def _num(value: object) -> float:
    if value is None or value == "":
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return 0.0


def _as_cell(value: object) -> str | int | float:
    if value is None:
        return ""
    if isinstance(value, (int, float, str)):
        return value
    return str(value)


def _rows_as_dicts(ws: Worksheet) -> list[dict[str, object]]:
    rows = ws.iter_rows(values_only=True)
    try:
        headers = [str(h) if h is not None else "" for h in next(rows)]
    except StopIteration:
        return []
    return [dict(zip(headers, row)) for row in rows]


def load_result(path: Path) -> tuple[list[dict[str, object]], list[dict[str, object]]]:
    wb = load_workbook(path, read_only=True, data_only=True)
    if "Invoices" not in wb.sheetnames:
        raise ValueError(f"'Invoices' sheet not found in {path}")
    invoices = _rows_as_dicts(wb["Invoices"])
    items = _rows_as_dicts(wb["Items"]) if "Items" in wb.sheetnames else []
    wb.close()
    return invoices, items


def cfop_by_invoice(items: list[dict[str, object]]) -> dict[str, str]:
    result: dict[str, str] = {}
    for item in items:
        key = str(item.get("chave_nfe", ""))
        if key not in result and item.get("cfop"):
            result[key] = str(item.get("cfop"))
    return result


def _style_header_cell(cell) -> None:
    cell.fill = HEADER_FILL
    cell.font = HEADER_FONT
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = BORDER


def build_parametros(
    ws: Worksheet,
    marketplaces: list[str],
    ads_rates: dict[str, float],
    afiliado_rates: dict[str, float],
    candaru_rates: dict[str, float],
) -> dict[str, str]:
    ws["A1"] = "Taxas por Marketplace"
    ws["A1"].font = TITLE_FONT
    for col, name in zip("ABCD", ("Marketplace", "Taxa ADS", "Taxa Afiliado", "Taxa CANDARU")):
        cell = ws[f"{col}2"]
        cell.value = name
        _style_header_cell(cell)

    for idx, mp in enumerate(marketplaces, start=3):
        ws[f"A{idx}"] = mp
        ws[f"B{idx}"] = ads_rates.get(mp, 0.0)
        ws[f"C{idx}"] = afiliado_rates.get(mp, 0.0)
        ws[f"D{idx}"] = candaru_rates.get(mp, CANDARU_RATE_DEFAULT)
        ws[f"B{idx}"].number_format = PCT_FMT
        ws[f"C{idx}"].number_format = PCT_FMT
        ws[f"D{idx}"].number_format = PCT_FMT

    # Global rate block (referenced by Detalhado formulas).
    ws["F2"] = "Taxa IRPJ/CSLL"
    ws["F2"].font = LABEL_FONT
    ws["G2"] = DEFAULT_IRPJ_RATE
    ws["G2"].number_format = PCT_FMT

    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 14
    ws.column_dimensions["F"].width = 16
    ws.column_dimensions["G"].width = 12
    ws.freeze_panes = "A3"
    return {"irpj": "$G$2"}


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
        _style_header_cell(ws.cell(row=1, column=col_idx, value=name))

    irpj = f"Parametros!{global_rate_cells['irpj']}"

    ordered = sorted(invoices, key=lambda inv: _num(inv.get("numero")))
    for offset, inv in enumerate(ordered):
        r = offset + 2
        chave = str(inv.get("chave_nfe", ""))

        for name in invoice_headers:
            ws[f"{det_col[name]}{r}"] = _as_cell(inv.get(name, ""))

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
    _style_header_cell(ws.cell(row=header_row, column=label_col, value="Metrica"))
    for i, mp in enumerate(marketplaces):
        _style_header_cell(ws.cell(row=header_row, column=first_mp_col + i, value=mp))
    _style_header_cell(ws.cell(row=header_row, column=total_col, value="Total"))
    _style_header_cell(ws.cell(row=header_row, column=pct_col, value="% s/ Faturado"))

    det = "Detalhado"
    fcol = det_col[SRC_MARKETPLACE]

    # (label, kind, detalhado column, number_format)
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
    faturado_row = row + 1  # Faturado is the 2nd metric

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
    ws.cell(row=note_row, column=label_col).font = Font(italic=True, color="808080")


def build_report(
    input_path: Path,
    output_path: Path,
) -> dict[str, object]:
    invoices, items = load_result(input_path)
    if not invoices:
        raise ValueError("No invoices found in the input workbook.")

    # Keep the Invoices columns (names and order) for the Detalhado sheet.
    invoice_headers = [h for h in invoices[0].keys() if h]
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
    ws_par = wb.create_sheet("Parametros")

    global_rate_cells = build_parametros(
        ws_par, marketplaces, ads_rates, afiliado_rates, candaru_rates
    )
    det_col, _ = build_detalhado(
        ws_det, invoices, invoice_headers, invoice_cfops, global_rate_cells
    )
    build_resumo(ws_resumo, marketplaces, det_col)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)

    return {
        "invoices": len(invoices),
        "items": len(items),
        "marketplaces": marketplaces,
    }


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
    print(f"Marketplaces: {marketplaces_text}")
    print(f"Report saved to: {output_path.resolve()}")


if __name__ == "__main__":
    main()
