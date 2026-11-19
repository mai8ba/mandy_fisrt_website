import http.client
import io
import json
import sqlite3
import tempfile
import threading
import unittest
from pathlib import Path
from unittest.mock import patch

import server


class SalesReportTests(unittest.TestCase):
    def setUp(self):
        self.db = sqlite3.connect(":memory:", check_same_thread=False)
        self.db.row_factory = sqlite3.Row
        self.db.executescript(
            """
            CREATE TABLE products (
                id TEXT PRIMARY KEY,
                sku TEXT,
                brand TEXT,
                name TEXT,
                price REAL,
                image TEXT,
                published INTEGER,
                archived_at INTEGER
            );
            CREATE TABLE orders (
                order_no TEXT PRIMARY KEY,
                items_json TEXT,
                created_at INTEGER
            );
            INSERT INTO products VALUES
                ('p1', 'SKU-1', 'Brand A', 'Product A', 3.5, '', 1, 0),
                ('p2', 'SKU-2', 'Brand B', 'Product B', 5.0, '', 1, 0),
                ('p3', 'SKU-3', 'Brand C', 'Unsold product', 8.0, '', 1, 0);
            """
        )
        self.db.executemany(
            "INSERT INTO orders VALUES (?, ?, ?)",
            [
                ("O1", json.dumps([{"id": "p1", "qty": 2, "price": 3.5}]), 100),
                (
                    "O2",
                    json.dumps([
                        {"id": "p1", "qty": 3, "price": 4.0},
                        {"id": "p2", "qty": 8, "price": 5.0},
                    ]),
                    200,
                ),
                ("O3", json.dumps([{"id": "p2", "qty": 1, "price": 5.0}]), 300),
            ],
        )
        self.db.execute("ALTER TABLE products ADD COLUMN colors_json TEXT DEFAULT '[]'")
        self.db.execute("ALTER TABLE products ADD COLUMN description TEXT DEFAULT ''")
        self.db.execute("UPDATE products SET description = 'Colors: 007/08 size: 1cm' WHERE id = 'p2'")
        self.shades = [f'{i:02d}' for i in range(1, 29)] + ['warm pink', '1-2-3', '=literal']
        self.db.execute("UPDATE products SET colors_json=? WHERE id='p1'", (json.dumps(self.shades),))

    def tearDown(self):
        self.db.close()

    def test_units_orders_and_revenue_are_aggregated(self):
        rows = server.sales_report_rows(self.db)
        self.assertEqual([row["id"] for row in rows], ["p2", "p1"])
        self.assertEqual(rows[0]["units_sold"], 9)
        self.assertEqual(rows[0]["order_count"], 2)
        self.assertEqual(rows[0]["revenue"], 45.0)
        self.assertEqual(rows[1]["units_sold"], 5)
        self.assertEqual(rows[1]["order_count"], 2)
        self.assertEqual(rows[1]["revenue"], 19.0)
        all_rows = server.sales_report_rows(self.db, None, include_zero=True)
        self.assertEqual([row["id"] for row in all_rows], ["p2", "p1", "p3"])
        self.assertEqual(all_rows[2]["units_sold"], 0)
        brand_rows = server.sales_report_rows(self.db, None, include_zero=True, sort_by_brand=True)
        self.assertEqual([row["id"] for row in brand_rows], ["p1", "p2", "p3"])

    def test_report_workbook_has_expected_sheet_and_font(self):
        rows = server.sales_report_rows(self.db, None, include_zero=True, sort_by_brand=True, published_only=True)
        workbook_data = server.build_sales_report_workbook(rows)
        self.assertGreater(len(workbook_data), 0)
        from openpyxl import load_workbook

        workbook = load_workbook(io.BytesIO(workbook_data), read_only=False)
        sheet = workbook["Published Products"]
        self.assertEqual(sheet["A1"].value, "Published Products Sales Statistics")
        self.assertEqual(sheet["K6"].value, 5)
        self.assertEqual(sheet["I6"].value, 19)
        self.assertEqual(sheet["K8"].value, 0)
        self.assertEqual(sheet["K5"].value, "Units Sold")
        self.assertEqual(sheet["E5"].value, "Shade/Type")
        self.assertEqual(sheet["J5"].value, "Order details")
        self.assertEqual(sheet["E6"].value.splitlines(), self.shades)
        self.assertEqual(sheet["E7"].value, "007\n08")
        self.assertEqual(sheet["E8"].value, "Default")
        self.assertEqual(sheet["E6"].data_type, 's')
        self.assertIsNone(sheet["J6"].value)
        self.assertEqual(sheet["J6"].number_format, '@')
        self.assertEqual(sheet.auto_filter.ref, 'A5:K8')
        self.assertEqual(sheet['K10'].value, '=SUM(K6:K8)')
        self.assertEqual(sheet['I10'].value, '=SUM(I6:I8)')
        self.assertEqual(sheet.max_column, 11)
        self.assertNotIn('SKU', [cell.value for cell in sheet[5]])
        self.assertEqual(sheet["A1"].font.name, "Arial")
        self.assertEqual(sheet["I6"].font.name, "Arial")
        self.assertTrue(all(cell.font.name == 'Arial' for row in sheet for cell in row if cell.value is not None))
        self.assertEqual(workbook._named_styles['Normal'].font.name, 'Arial')
        from zipfile import ZipFile
        with ZipFile(io.BytesIO(workbook_data)) as archive:
            self.assertNotIn(b'<scheme', archive.read('xl/styles.xml'))
        workbook.close()

    def test_sales_download_excludes_unpublished_archived_and_deleted_products(self):
        from openpyxl import load_workbook

        self.db.execute("UPDATE products SET archived_at=1 WHERE id='p1'")
        self.db.execute("UPDATE products SET published=0 WHERE id='p2'")
        self.db.execute("INSERT INTO orders VALUES (?, ?, ?)", (
            "O4", json.dumps([{"id": "deleted", "qty": 99, "price": 5}]), 400,
        ))
        self.db.commit()

        with patch.object(server, "connect", return_value=self.db):
            with server.ThreadingHTTPServer(("127.0.0.1", 0), server.CatalogHandler) as httpd:
                thread = threading.Thread(target=httpd.serve_forever)
                thread.start()
                client = http.client.HTTPConnection("127.0.0.1", httpd.server_port, timeout=5)
                try:
                    client.request("GET", "/api/admin/sales-report/excel")
                    response = client.getresponse()
                    self.assertEqual(response.status, 401)
                    response.read()
                    with patch.object(server.CatalogHandler, "current_user", return_value="admin"):
                        client.request("GET", "/api/admin/sales-report/excel")
                        response = client.getresponse()
                        self.assertEqual(response.status, 200)
                        self.assertIn("sales-report-published-products-", response.getheader("Content-Disposition"))
                        data = response.read()
                finally:
                    client.close()
                    httpd.shutdown()
                    thread.join()

        workbook = load_workbook(io.BytesIO(data))
        try:
            sheet = workbook["Published Products"]
            self.assertEqual(sheet.auto_filter.ref, "A5:K6")
            self.assertEqual(sheet["C6"].value, "Unsold product")
            self.assertEqual(sheet["G6"].value, "Published")
            self.assertEqual(sheet["K6"].value, 0)
        finally:
            workbook.close()

    def test_product_text_is_literal_and_portrait_images_fit(self):
        from openpyxl import load_workbook
        from openpyxl.utils.units import pixels_to_EMU
        from PIL import Image

        self.db.execute("UPDATE products SET brand='=Brand', name='#N/A' WHERE id='p1'")
        rows = server.sales_report_rows(self.db)
        row = next(row for row in rows if row["id"] == "p1")
        with tempfile.TemporaryDirectory() as directory:
            image_path = Path(directory) / "portrait.jpg"
            Image.new("RGB", (24, 96)).save(image_path)
            with patch.object(server, "make_thumbnail", return_value=image_path):
                data = server.build_sales_report_workbook([row])
        with io.BytesIO(data) as stream:
            workbook = load_workbook(stream)
            sheet = workbook["Published Products"]
            self.assertEqual((sheet["B6"].value, sheet["B6"].data_type), ("=Brand", "s"))
            self.assertEqual((sheet["C6"].value, sheet["C6"].data_type), ("#N/A", "s"))
            self.assertEqual(sheet._images[0].anchor.ext.cx, pixels_to_EMU(24))
            self.assertEqual(sheet._images[0].anchor.ext.cy, pixels_to_EMU(96))
            self.assertLessEqual(sheet.row_dimensions[6].height, 409)
            workbook.close()


if __name__ == "__main__":
    unittest.main()
