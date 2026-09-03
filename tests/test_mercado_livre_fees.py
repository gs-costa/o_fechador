from pathlib import Path
from tempfile import TemporaryDirectory
from unittest import TestCase

from openpyxl import Workbook

from src.marketplace_fees import load_mercado_livre_fees


class MercadoLivreFeeTests(TestCase):
    def test_indexes_and_sums_fees_by_operation_or_package(self) -> None:
        with TemporaryDirectory() as directory:
            path = Path(directory) / "mercado_livre.xlsx"
            workbook = Workbook()
            worksheet = workbook.active
            worksheet.append(["Relatório de conciliação por vendas"])
            worksheet.append(
                [
                    "Número da operação",
                    "Número do pacote",
                    "Valor total de tarifas (desconto já aplicado)",
                ]
            )
            worksheet.append(["operation-1", "package-1", 10])
            worksheet.append(["operation-1", "package-1", 2.5])
            worksheet.append(["operation-2", "package-1", 3])
            worksheet.append(["same-id", "same-id", 4])
            workbook.save(path)

            fees = load_mercado_livre_fees(path)

        self.assertEqual(fees["operation-1"], 12.5)
        self.assertEqual(fees["operation-2"], 3)
        self.assertEqual(fees["package-1"], 15.5)
        self.assertEqual(fees["same-id"], 4)
