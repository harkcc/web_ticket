"""Focused regression tests for the Kaiqi US/EU invoice handlers.

These tests intentionally exercise the handler with synthetic packing-list objects and
cached product data.  They do not connect to MongoDB, fetch product images, or mutate
the production templates.
"""

from __future__ import annotations

import importlib
import sys
import tempfile
import types
import unittest
from hashlib import sha256
from pathlib import Path
from zipfile import ZipFile


WEB_TICKET_ROOT = Path(__file__).resolve().parents[1]
TEMPLATE_DIR = WEB_TICKET_ROOT / "表格模版"
GENERATED_TEMPLATE_DIR = WEB_TICKET_ROOT.parent / "outputs" / "kaiqi-templates-20260818"


def _import_runtime():
    """Import production classes, providing only local test fallbacks if needed.

    The application requirements are not installed in every developer checkout.  The
    handler itself only needs openpyxl for these tests, so missing optional runtime
    modules are stubbed without touching production files.
    """

    try:
        generator_module = importlib.import_module("generator")
        data_module = importlib.import_module("get_ticket_data")
        return generator_module, data_module
    except ModuleNotFoundError:
        numpy_stub = types.ModuleType("numpy")
        numpy_stub.add = lambda left, right: left + right
        fake_numeric_type = type("FakeNumpyScalar", (), {})
        for numeric_name in (
            "short",
            "ushort",
            "intc",
            "uintc",
            "int_",
            "uint",
            "longlong",
            "ulonglong",
            "half",
            "float16",
            "single",
            "double",
            "longdouble",
            "int8",
            "int16",
            "int32",
            "int64",
            "uint8",
            "uint16",
            "uint32",
            "uint64",
            "intp",
            "uintp",
            "float32",
            "float64",
            "bool_",
            "floating",
            "integer",
        ):
            setattr(numpy_stub, numeric_name, fake_numeric_type)
        pandas_stub = types.ModuleType("pandas")

        db_stub = types.ModuleType("db_connector")

        class UnusedMongoDBConnector:
            pass

        db_stub.MongoDBConnector = UnusedMongoDBConnector

        for name, module in {
            "numpy": numpy_stub,
            "pandas": pandas_stub,
            "db_connector": db_stub,
        }.items():
            sys.modules.setdefault(name, module)

        # A failed import can leave a partially initialized module behind.
        sys.modules.pop("generator", None)
        sys.modules.pop("get_ticket_data", None)
        generator_module = importlib.import_module("generator")
        data_module = importlib.import_module("get_ticket_data")
        return generator_module, data_module


generator, ticket_data = _import_runtime()
InvoiceGenerator = generator.InvoiceGenerator
PackingListBox = ticket_data.PackingListBox
PackingListItem = ticket_data.PackingListItem
load_workbook = importlib.import_module("openpyxl").load_workbook
Workbook = importlib.import_module("openpyxl").Workbook


class _StubDB:
    """Context-manager DB double; product lookups are served from the cache."""

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        return False


def _product_cache():
    return {
        "MSKU-A": {
            "cn_name": "产品甲",
            "en_name": "Product A",
            "hs_code": "9503000000",
            "price": "4.25",
            "weight": "0.4",
            "brand": "Brand A",
            "model": "A-01",
            "link": "https://www.amazon.com/dp/B0A1234567",
            "asin": "",
            "material_en": "plastic",
            "material_cn": "塑料",
            "usage_en": "decoration",
            "usage_cn": "装饰",
            "electrified": "否",
            "magnetic": "否",
        },
        "MSKU-B": {
            "cn_name": "产品乙",
            "en_name": "Product B",
            "hs_code": "4202920000",
            "price": 3.5,
            "weight": 0.25,
            "brand": "无",
            "model": "B-02",
            "link": "https://www.amazon.co.uk/dp/B0B7654321",
            "asin": "B0B7654321",
            "material_en": "cotton",
            "material_cn": "棉",
            "usage_en": "storage",
            "usage_cn": "收纳",
            "electrified": "是",
            "magnetic": "否",
        },
        "MSKU-C": {
            "cn_name": "产品丙",
            "en_name": "Product C",
            "hs_code": "8507600000",
            "price": 2.0,
            "weight": 0.6,
            "brand": "Brand C",
            "model": "C-03",
            "link": "https://www.amazon.de/dp/B0C1111111",
            "asin": "B0C1111111",
            "material_en": "metal",
            "material_cn": "金属",
            "usage_en": "tool",
            "usage_cn": "工具",
            "electrified": "否",
            "magnetic": "是",
        },
    }


def _synthetic_boxes():
    box_one = PackingListBox(1)
    box_one.set_dimensions(60, 40, 30)
    box_one.set_weight(8.4)
    box_one.add_item(
        PackingListItem(
            sequence_no=1,
            msku="MSKU-A",
            fnsku="X00-FNSKU-A",
            product_name="产品甲",
            sku="SKU-A",
            quantity=2,
            box_quantities={1: 2},
        )
    )
    box_one.add_item(
        PackingListItem(
            sequence_no=2,
            msku="MSKU-B",
            fnsku="X00-FNSKU-B",
            product_name="产品乙",
            sku="SKU-B",
            quantity=1,
            box_quantities={1: 1},
        )
    )

    box_two = PackingListBox(2)
    box_two.set_dimensions(50, 35, 25)
    box_two.set_weight(6.2)
    box_two.add_item(
        PackingListItem(
            sequence_no=3,
            msku="MSKU-C",
            fnsku="X00-FNSKU-C",
            product_name="产品丙",
            sku="SKU-C",
            quantity=3,
            box_quantities={2: 3},
        )
    )
    return {1: box_one, 2: box_two}


def _synthetic_workbook():
    workbook = Workbook()
    workbook.remove(workbook.active)
    for sheet_name in ("美线+加线-发票导入", "欧线+空派-发票导入", "旧模板"):
        sheet = workbook.create_sheet(sheet_name)
        for row in range(1, 31):
            for column in range(1, 25):
                sheet.cell(row=row, column=column).value = f"SAMPLE-{sheet_name}-{row}-{column}"

        # Keep the production sheet shape and make the target header explicit.
        sheet.cell(row=15, column=24).value = "FNSKU辅助列"

    return workbook


def _address_info(country_code="US", country_name="美国", shipment_name="2026.08.18-911专线-3/3"):
    return {
        "seller_info": {
            "country_code": country_code,
            "country_name": country_name,
        },
        "address_info": {
            "amazonReferenceId": "REF-TEST-001",
            "shipmentName": shipment_name,
            "warehouseId": "FC-TEST",
            "postalCode": "10001",
            "name": "Test Recipient",
            "phoneNumber": "10000000000",
            "city": "Test City",
            "stateOrProvinceCode": "NY",
            "addressLine1": "1 Test Street",
            "addressLine2": "Unit 2",
            "countryCode": country_code,
        },
    }


def _build_generator():
    instance = InvoiceGenerator(
        upload_folder=tempfile.gettempdir(),
        output_folder=tempfile.gettempdir(),
        db_connector=_StubDB(),
    )
    instance.product_cache = _product_cache()
    instance.cache_enabled = True
    instance.missing_products = set()
    instance.image_cache = {}
    instance.image_cache_enabled = True
    instance.image_calls = []
    instance.image_folder = tempfile.gettempdir()

    def image_stub(sheet, cell, msku, image_folder):
        instance.image_calls.append((sheet.title, cell, msku, image_folder))

    # The production image insertion path is intentionally replaced by a recording
    # double; the handler still exercises its image-call boundary for every row.
    instance.insert_original_product_image = image_stub
    return instance


class KaiqiHandlerTests(unittest.TestCase):
    def _run_handler(self, template_name, target_sheet, address):
        invoice_generator = _build_generator()
        workbook = _synthetic_workbook()
        template_path = TEMPLATE_DIR / template_name
        handler = invoice_generator._get_template_handler(str(template_path))
        self.assertIsNotNone(handler)
        handler(
            workbook,
            _synthetic_boxes(),
            code="ADDRESS-CODE",
            address_info=address,
            shipment_id="FBA-TEST",
        )
        return workbook, invoice_generator, target_sheet

    def test_kaiqi_us_routes_rows_to_us_sheet_and_writes_fnsku_x(self):
        workbook, invoice_generator, sheet_name = self._run_handler(
            "凯琦美线.xlsx",
            "美线+加线-发票导入",
            _address_info(),
        )
        sheet = workbook[sheet_name]

        self.assertEqual(workbook.active.title, sheet_name)
        self.assertEqual(sheet.sheet_state, "visible")
        self.assertEqual([sheet.cell(row, 24).value for row in range(16, 19)], [
            "X00-FNSKU-A", "X00-FNSKU-B", "X00-FNSKU-C"
        ])
        self.assertEqual([sheet.cell(row, 18).value for row in range(16, 19)], [1, 1, 1])
        self.assertEqual([sheet.cell(row, 19).value for row in range(16, 19)], [0.8, 0.25, 1.8])
        self.assertEqual([sheet.cell(row, 20).value for row in range(16, 19)], [8.4, 8.4, 6.2])
        self.assertEqual([sheet.cell(row, 21).value for row in range(16, 19)], [60, 60, 50])
        self.assertEqual([sheet.cell(row, 22).value for row in range(16, 19)], [40, 40, 35])
        self.assertEqual([sheet.cell(row, 23).value for row in range(16, 19)], [30, 30, 25])
        self.assertEqual(sheet["B3"].value, "FBA-TEST")
        self.assertEqual(sheet["B16"].value, "FBA-TESTU000001")
        self.assertEqual(sheet["B4"].value, "REF-TEST-001")
        self.assertEqual(sheet["B5"].value, 2)
        self.assertEqual(sheet["B6"].value, "美国")
        self.assertEqual(sheet["B7"].value, "911专线")
        self.assertIsNone(sheet["A19"].value)
        self.assertEqual(len(invoice_generator.image_calls), 3)

    def test_kaiqi_eu_routes_rows_to_eu_sheet_and_preserves_route_columns(self):
        workbook, invoice_generator, sheet_name = self._run_handler(
            "凯琦欧线.xlsx",
            "欧线+空派-发票导入",
            _address_info(country_code="DE", country_name="德国", shipment_name="2026.08.18-T07空派-3/3"),
        )
        sheet = workbook[sheet_name]

        self.assertEqual(workbook.active.title, sheet_name)
        self.assertEqual(sheet.sheet_state, "visible")
        self.assertEqual([sheet.cell(row, 24).value for row in range(16, 19)], [
            "X00-FNSKU-A", "X00-FNSKU-B", "X00-FNSKU-C"
        ])
        self.assertEqual([sheet.cell(row, 18).value for row in range(16, 19)], [
            "https://www.amazon.com/dp/B0A1234567",
            "https://www.amazon.co.uk/dp/B0B7654321",
            "https://www.amazon.de/dp/B0C1111111",
        ])
        self.assertEqual([sheet.cell(row, 19).value for row in range(16, 19)], [
            "B0A1234567", "B0B7654321", "B0C1111111"
        ])
        self.assertEqual([sheet.cell(row, 20).value for row in range(16, 19)], [8.4, 8.4, 6.2])
        self.assertEqual([sheet.cell(row, 21).value for row in range(16, 19)], [60, 60, 50])
        self.assertEqual([sheet.cell(row, 22).value for row in range(16, 19)], [40, 40, 35])
        self.assertEqual([sheet.cell(row, 23).value for row in range(16, 19)], [30, 30, 25])
        self.assertEqual(sheet["B6"].value, "德国")
        self.assertEqual(sheet["B7"].value, "T07空派")
        self.assertEqual(len(invoice_generator.image_calls), 3)

    def test_handler_clears_samples_and_does_not_touch_legacy_sheet(self):
        workbook, _, sheet_name = self._run_handler(
            "凯琦美线.xlsx",
            "美线+加线-发票导入",
            _address_info(),
        )

        target = workbook[sheet_name]
        legacy = workbook["旧模板"]
        for row in range(19, 31):
            self.assertTrue(
                all(target.cell(row=row, column=column).value is None for column in range(1, 25)),
                f"sample data leaked at {target.title}!{row}",
            )
        self.assertEqual(legacy["A16"].value, "SAMPLE-旧模板-16-1")
        self.assertEqual(legacy["X30"].value, "SAMPLE-旧模板-30-24")
        self.assertEqual(legacy.sheet_state, "hidden")

    def test_generated_template_files_are_loadable_by_production_openpyxl(self):
        """Catch the exact failure production load_workbook would see on deploy."""
        for template_name in ("凯琦美线.xlsx", "凯琦欧线.xlsx"):
            template_path = GENERATED_TEMPLATE_DIR / template_name
            if not template_path.exists():
                template_path = TEMPLATE_DIR / template_name
            self.assertTrue(template_path.exists(), f"missing template: {template_path}")
            with self.subTest(template=template_name):
                workbook = load_workbook(template_path)
                self.assertIn(
                    "美线+加线-发票导入" if "美线" in template_name else "欧线+空派-发票导入",
                    workbook.sheetnames,
                )

    def test_generated_workbook_can_be_saved_and_reopened(self):
        """Verify a handler result remains a real downloadable xlsx artifact."""
        cases = (
            ("凯琦美线.xlsx", "美线+加线-发票导入", _address_info()),
            (
                "凯琦欧线.xlsx",
                "欧线+空派-发票导入",
                _address_info(country_code="DE", country_name="德国"),
            ),
        )
        for template_name, sheet_name, address in cases:
            with self.subTest(template=template_name), tempfile.TemporaryDirectory() as temp_dir:
                workbook, _, _ = self._run_handler(template_name, sheet_name, address)
                output_path = Path(temp_dir) / template_name
                workbook.save(output_path)
                reopened = load_workbook(output_path, data_only=False)
                self.assertEqual(reopened.active.title, sheet_name)
                self.assertEqual(reopened[sheet_name]["X16"].value, "X00-FNSKU-A")
                self.assertEqual(reopened[sheet_name]["X18"].value, "X00-FNSKU-C")


class YibaAddressLibraryTests(unittest.TestCase):
    template_path = TEMPLATE_DIR / "一八供应链new.xlsx"
    protected_parts = {
        "xl/worksheets/sheet1.xml": "b0a94767c1bddf35b338e1128bfe5b7e72146b6ebe41e7f6c42fc3da697ff6b6",
        "xl/worksheets/sheet2.xml": "732ca51eb4632e3e105d0720c4f07ceb9d6735f62760e3e8452c2ec34d51dda5",
        "xl/worksheets/sheet4.xml": "3c4c2d7beb79b290d6d56bcdbd5cf5432c855a092e7df2859b8ab8c78a05fb8f",
        "xl/worksheets/sheet5.xml": "5d091c2724982d59ad566b8fdd2bd275ef8f31dbe079a40b609777f64139b218",
        "xl/styles.xml": "ee9b3b98f2a6943727f9bdf07ce6ab7ceea368f0b6560178b01a6cb335253f6e",
        "xl/drawings/drawing1.xml": "08a3bc68dee040df0847f1eb492bee85b3fc9f89df0c7aeb8a2b34dda613ee97",
        "xl/cellimages.xml": "422a762190b51fdb92daaa8b4a4f13e5717e8dd3b18d90e1eff59c9b19258933",
        "xl/_rels/cellimages.xml.rels": "e4e2c298d367e82277d66b5d8ed144d1b001c65f3f7bd6c66cea31f474549d81",
        "xl/media/image1.png": "8ae4aea3635d409a8081cdc98b25466cbc0df59cee3c712ba264d72a3aac4799",
        "xl/media/image2.jpeg": "de6e594fa3cea78f4583cd55901809dd1f60c2fa089de472ad99b67d0f2824ff",
    }

    def test_latest_address_library_rows_are_present(self):
        workbook = load_workbook(self.template_path, data_only=False)
        sheet = workbook["地址库编码表"]
        rows = list(sheet.iter_rows(min_row=2, max_col=14, values_only=True))
        rows_by_code = {}
        for row in rows:
            code = str(row[0] or "").strip()
            if code:
                rows_by_code.setdefault(code, []).append(row)

        expected_new_codes = {
            "PSP3-UPS", "HEA2-UPS", "QXY8", "AWD仓-IUSJ", "XPB2",
            "TOL3", "TOL3-UPS", "XLX6", "MCI4", "MCI4-UPS",
            "BJC1", "QZZ7", "XSE6", "XOR4",
        }
        self.assertEqual(sum(len(items) for items in rows_by_code.values()), 2210)
        self.assertEqual(len(rows_by_code), 2205)
        self.assertTrue(expected_new_codes.issubset(rows_by_code))
        self.assertEqual(rows_by_code["沃尔玛ORD1S"][0][1], "沃尔玛地址")
        self.assertEqual(rows_by_code["LBA8"][0][8], "Moor Way")
        self.assertEqual(rows_by_code["LBA8"][0][10], "Leeds")
        self.assertEqual(str(rows_by_code["LBA8"][0][13]), "LS15 0BF")
        self.assertEqual(
            {str(row[13]) for row in rows_by_code["FTW2"]},
            {"75019", "75261"},
        )

    def test_non_address_sheets_and_images_remain_original(self):
        with ZipFile(self.template_path) as workbook_zip:
            for part_name, expected_digest in self.protected_parts.items():
                with self.subTest(part=part_name):
                    self.assertEqual(
                        sha256(workbook_zip.read(part_name)).hexdigest(),
                        expected_digest,
                    )

    def test_address_only_template_still_generates_and_reopens(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            workbook = load_workbook(self.template_path, data_only=False)
            invoice_generator = _build_generator()
            handler = invoice_generator._get_template_handler(str(self.template_path))
            handler(
                workbook,
                _synthetic_boxes(),
                code="ADDRESS-CODE",
                address_info=_address_info(),
                shipment_id="FBA-TEST",
            )
            output_path = Path(temp_dir) / "一八供应链new-output.xlsx"
            workbook.save(output_path)

            reopened = load_workbook(output_path, data_only=False)
            main = reopened["专线箱单 "]
            self.assertEqual(main["B10"].value, "=VLOOKUP($B$9,地址库编码表!1:1048576,5,0)")
            self.assertEqual(main["U20"].value, "X00-FNSKU-A")
            self.assertEqual(main["U22"].value, "X00-FNSKU-C")

    def test_us_and_canada_leave_declaration_unit_price_blank(self):
        cases = (
            ("US", "美国"),
            ("CA", "加拿大"),
        )
        for country_code, country_name in cases:
            with self.subTest(country=country_code), tempfile.TemporaryDirectory() as temp_dir:
                workbook = load_workbook(self.template_path, data_only=False)
                invoice_generator = _build_generator()
                handler = invoice_generator._get_template_handler(str(self.template_path))
                handler(
                    workbook,
                    _synthetic_boxes(),
                    code="ADDRESS-CODE",
                    address_info=_address_info(
                        country_code=country_code,
                        country_name=country_name,
                    ),
                    shipment_id="FBA-TEST",
                )
                output_path = Path(temp_dir) / f"一八供应链-{country_code}.xlsx"
                workbook.save(output_path)

                main = load_workbook(output_path, data_only=False)["专线箱单 "]
                for row in range(20, 23):
                    self.assertIsNone(main.cell(row=row, column=12).value)
                    self.assertEqual(main.cell(row=row, column=20).value, f"=L{row}*S{row}")

    def test_other_countries_keep_declaration_unit_price(self):
        workbook = load_workbook(self.template_path, data_only=False)
        invoice_generator = _build_generator()
        handler = invoice_generator._get_template_handler(str(self.template_path))
        handler(
            workbook,
            _synthetic_boxes(),
            code="ADDRESS-CODE",
            address_info=_address_info(country_code="DE", country_name="德国"),
            shipment_id="FBA-TEST",
        )

        main = workbook["专线箱单 "]
        self.assertEqual([main.cell(row=row, column=12).value for row in range(20, 23)], [
            "4.25", 3.5, 2.0,
        ])


if __name__ == "__main__":
    unittest.main()
