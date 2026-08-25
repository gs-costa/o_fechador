"""Join NF-e invoices to marketplace exports and calculate marketplace fees."""

from __future__ import annotations

import re
import unicodedata
from collections.abc import Iterable, Iterator
from dataclasses import dataclass
from pathlib import Path

from openpyxl import load_workbook
from xlrd import open_workbook

ORDER_NUMBER_COLUMN = "numero_pedido"
MARKETPLACE_FEE_COLUMN = "taxa_marketplace"

_BLING_INVOICE_ALIASES = ("numero", "numero da nota", "numero nf", "nota fiscal")
_BLING_ORDER_ALIASES = (
    "numero do pedido multiloja",
    "pedido multiloja",
    "numero do pedido",
)
_SHOPEE_TYPE_ALIASES = ("ver", "view", "tipo")
_SHOPEE_ORDER_ALIASES = (
    "id do pedido",
    "numero do pedido",
    "pedido",
    "order id",
)
_SHOPEE_SUBTOTAL_ALIASES = (
    "preco do produto",
    "subtotal de mercadoria",
    "subtotal de mercadorias",
    "product subtotal",
)
_SHOPEE_INCOME_ALIASES = (
    "quantia total lancada r",
    "quantia total lancada",
    "rendas do pedido",
    "renda do pedido",
    "order income",
    "order earnings",
)
_MERCADO_LIVRE_ORDER_ALIASES = (
    "n de venda",
    "n o de venda",
    "numero de venda",
    "numero da venda",
    "id da venda",
    "order id",
)
_MERCADO_LIVRE_REVENUE_ALIASES = (
    "receita por produtos",
    "receita por produtos brl",
)
_MERCADO_LIVRE_TOTAL_ALIASES = ("total", "total brl")


@dataclass(frozen=True)
class FeeJoinStats:
    """Matching totals from the two-stage invoice enrichment."""

    invoices: int
    orders_found: int
    fees_found: int
    fees_total: float

    @property
    def invoices_without_order(self) -> int:
        return self.invoices - self.orders_found

    @property
    def orders_without_fee(self) -> int:
        return self.orders_found - self.fees_found


def _normalized_text(value: object) -> str:
    text = unicodedata.normalize("NFKD", str(value or "").strip().casefold())
    without_accents = "".join(char for char in text if not unicodedata.combining(char))
    return " ".join(re.sub(r"[^a-z0-9]+", " ", without_accents).split())


def _normalized_id(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return re.sub(r"\s+", "", str(value).strip()).casefold()


def _normalized_invoice_number(value: object) -> str:
    number = _normalized_id(value)
    if number.isdigit():
        return number.lstrip("0") or "0"
    return number


def _money(value: object) -> float:
    if value is None or value == "":
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip()
    negative = text.startswith("(") and text.endswith(")")
    text = re.sub(r"[^\d,.\-]", "", text)
    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        text = text.replace(".", "").replace(",", ".")

    try:
        amount = float(text)
    except ValueError:
        return 0.0
    return -amount if negative else amount


def _workbook_rows(path: Path) -> Iterator[tuple[str, Iterable[tuple[object, ...]]]]:
    suffix = path.suffix.casefold()
    if suffix == ".xls":
        workbook = open_workbook(path, ignore_workbook_corruption=True)
        for sheet in workbook.sheets():
            yield sheet.name, (tuple(sheet.row_values(row)) for row in range(sheet.nrows))
        return
    if suffix not in {".xlsx", ".xlsm"}:
        raise ValueError(f"Formato não suportado: {path.suffix}. Use .xls ou .xlsx.")

    # read_only=False: Shopee exports often set dimension to A1, which hides
    # the real header row (row 3 of the Renda sheet) in read-only mode.
    workbook = load_workbook(path, read_only=False, data_only=True)
    try:
        for worksheet in workbook.worksheets:
            yield worksheet.title, worksheet.iter_rows(values_only=True)
    finally:
        workbook.close()


def _column(row: tuple[object, ...], aliases: tuple[str, ...]) -> int | None:
    normalized_aliases = {_normalized_text(alias) for alias in aliases}
    for index, value in enumerate(row):
        if _normalized_text(value) in normalized_aliases:
            return index
    return None


def _required_columns(
    row: tuple[object, ...],
    definitions: dict[str, tuple[str, ...]],
) -> dict[str, int] | None:
    columns: dict[str, int] = {}
    for name, aliases in definitions.items():
        index = _column(row, aliases)
        if index is None:
            return None
        columns[name] = index
    return columns


def _value(row: tuple[object, ...], index: int) -> object:
    return row[index] if index < len(row) else None


def load_bling_orders(path: Path) -> dict[str, str]:
    """Return normalized NF-e number -> normalized marketplace order number."""
    if not path.is_file():
        raise FileNotFoundError(f"Planilha do Bling não encontrada: {path}")

    definitions = {
        "invoice": _BLING_INVOICE_ALIASES,
        "order": _BLING_ORDER_ALIASES,
    }
    for _sheet_name, rows in _workbook_rows(path):
        iterator = iter(rows)
        for row in iterator:
            columns = _required_columns(row, definitions)
            if columns is None:
                continue

            orders: dict[str, str] = {}
            for data_row in iterator:
                invoice = _normalized_invoice_number(_value(data_row, columns["invoice"]))
                order = _normalized_id(_value(data_row, columns["order"]))
                if invoice and order:
                    orders[invoice] = order
            return orders

    raise ValueError(
        "A planilha do Bling não contém as colunas 'Número' e "
        "'Número do pedido multiloja'."
    )


def _sum_fees(
    rows: Iterable[tuple[object, ...]],
    columns: dict[str, int],
    *,
    left: str,
    right: str,
    filter_column: str | None = None,
) -> dict[str, float]:
    fees: dict[str, float] = {}
    for row in rows:
        if filter_column is not None:
            row_type = _normalized_text(_value(row, columns[filter_column]))
            if row_type not in {"order", "pedido"}:
                continue

        order = _normalized_id(_value(row, columns["order"]))
        if not order:
            continue
        fee = _money(_value(row, columns[left])) - _money(
            _value(row, columns[right])
        )
        fees[order] = round(fees.get(order, 0.0) + fee, 2)
    return fees


def load_marketplace_fees(path: Path) -> dict[str, float]:
    """Detect a Shopee or Mercado Livre export and return fee totals by order."""
    if not path.is_file():
        raise FileNotFoundError(f"Planilha do marketplace não encontrada: {path}")

    shopee_definitions = {
        "type": _SHOPEE_TYPE_ALIASES,
        "order": _SHOPEE_ORDER_ALIASES,
        "subtotal": _SHOPEE_SUBTOTAL_ALIASES,
        "income": _SHOPEE_INCOME_ALIASES,
    }
    mercado_livre_definitions = {
        "order": _MERCADO_LIVRE_ORDER_ALIASES,
        "revenue": _MERCADO_LIVRE_REVENUE_ALIASES,
        "total": _MERCADO_LIVRE_TOTAL_ALIASES,
    }

    for _sheet_name, rows in _workbook_rows(path):
        iterator = iter(rows)
        for row in iterator:
            shopee_columns = _required_columns(row, shopee_definitions)
            if shopee_columns is not None:
                return _sum_fees(
                    iterator,
                    shopee_columns,
                    left="subtotal",
                    right="income",
                    filter_column="type",
                )

            mercado_livre_columns = _required_columns(
                row, mercado_livre_definitions
            )
            if mercado_livre_columns is not None:
                return _sum_fees(
                    iterator,
                    mercado_livre_columns,
                    left="revenue",
                    right="total",
                )

    raise ValueError(
        f"Não foi possível reconhecer a planilha '{path.name}'. Para Shopee (aba "
        "Renda, cabeçalhos na linha 3), são necessárias as colunas Ver, ID do "
        "pedido, Preço do produto e Quantia total lançada (R$). Para Mercado "
        "Livre, número da venda, Receita por produtos e Total (BRL)."
    )


def enrich_invoices_with_marketplace_fees(
    invoices: list[dict[str, object]],
    *,
    bling_path: Path | None,
    marketplace_paths: dict[str, Path],
) -> tuple[list[dict[str, object]], FeeJoinStats]:
    """Add order number and marketplace fee to invoice copies.

    An order is looked up in every supplied spreadsheet, not only in the one
    given for its own NF-e folder: a folder often mixes orders from more than
    one marketplace.
    """
    orders_by_invoice = load_bling_orders(bling_path) if bling_path else {}
    fees_by_order: dict[str, float] = {}
    for path in marketplace_paths.values():
        fees_by_order.update(load_marketplace_fees(path))

    enriched: list[dict[str, object]] = []
    orders_found = 0
    fees_found = 0
    fees_total = 0.0
    for invoice in invoices:
        result = dict(invoice)
        invoice_number = _normalized_invoice_number(invoice.get("numero"))
        order = orders_by_invoice.get(invoice_number, "")
        fee = fees_by_order.get(order, 0.0) if order else 0.0

        if order:
            orders_found += 1
        if order and order in fees_by_order:
            fees_found += 1
            fees_total += fee

        result[ORDER_NUMBER_COLUMN] = order
        result[MARKETPLACE_FEE_COLUMN] = fee
        enriched.append(result)

    return enriched, FeeJoinStats(
        invoices=len(invoices),
        orders_found=orders_found,
        fees_found=fees_found,
        fees_total=round(fees_total, 2),
    )
