from pathlib import Path
from tempfile import TemporaryDirectory
from unittest import TestCase

from src.convert_nfe_xml_to_xlsx import parse_nfe
from src.marketplace_fees import _order_for_invoice

NFE_TEMPLATE = """\
<nfeProc xmlns="http://www.portalfiscal.inf.br/nfe">
  <NFe>
    <infNFe Id="NFe123">
      <ide><nNF>42</nNF></ide>
      <dest />
      <total><ICMSTot /></total>
      <det nItem="1">
        <prod>
          <cProd>SKU</cProd>
          <xPed>{product_xped}</xPed>
        </prod>
      </det>
      {extra}
    </infNFe>
  </NFe>
</nfeProc>
"""


def _parse_invoice(
    market_place: str,
    *,
    product_xped: str = "",
    extra: str = "",
) -> dict[str, str | float]:
    with TemporaryDirectory() as directory:
        path = Path(directory) / "invoice.xml"
        path.write_text(
            NFE_TEMPLATE.format(product_xped=product_xped, extra=extra),
            encoding="utf-8",
        )
        invoice, _items = parse_nfe(path, market_place=market_place)
    return invoice


class InvoiceOrderExtractionTests(TestCase):
    def test_mercado_livre_uses_xped_inside_compra(self) -> None:
        invoice = _parse_invoice(
            "Mercado Livre",
            product_xped="item-order",
            extra="<compra><xPed>2000014445403521</xPed></compra>",
        )

        self.assertEqual(invoice["xped"], "2000014445403521")

    def test_shopee_uses_any_xped_in_invoice(self) -> None:
        invoice = _parse_invoice("Shopee", product_xped="shopee-order")

        self.assertEqual(invoice["xped"], "shopee-order")

    def test_full_keeps_item_xped_extraction(self) -> None:
        invoice = _parse_invoice("Mercado Livre Full", product_xped="full-order")

        self.assertEqual(invoice["xped"], "full-order")


class OrderFallbackTests(TestCase):
    def test_xml_order_has_priority_for_marketplace(self) -> None:
        invoice = {"market_place": "Shopee", "numero": "42", "xped": "xml-order"}

        order = _order_for_invoice(invoice, {"42": "bling-order"})

        self.assertEqual(order, "xml-order")

    def test_bling_is_used_when_xml_order_is_missing(self) -> None:
        invoice = {"market_place": "Mercado Livre Full", "numero": "42", "xped": ""}

        order = _order_for_invoice(invoice, {"42": "bling-order"})

        self.assertEqual(order, "bling-order")
