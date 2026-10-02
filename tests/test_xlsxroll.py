import sys
import tempfile
import unittest
import zipfile
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "examples"))

from make_sample import make  # noqa: E402

from xlsxroll.__main__ import main  # noqa: E402
from xlsxroll.core import add_months, roll  # noqa: E402

PAIRS = [("2026/09/01-2026/09/30", "2026/10/01-2026/10/31"), ("9月", "10月")]


class AddMonthsTest(unittest.TestCase):
    def test_month_end_stays_month_end(self):
        self.assertEqual(add_months(date(2026, 9, 30), 1), date(2026, 10, 31))
        self.assertEqual(add_months(date(2026, 1, 31), 1), date(2026, 2, 28))

    def test_normal_day_and_year_change(self):
        self.assertEqual(add_months(date(2026, 10, 5), 1), date(2026, 11, 5))
        self.assertEqual(add_months(date(2026, 12, 15), 1), date(2027, 1, 15))


class RollTest(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = Path(self._tmp.name)
        self.src = self.tmp / "2026-09.xlsx"
        self.dst = self.tmp / "2026-10.xlsx"
        make(self.src)

    def tearDown(self):
        self._tmp.cleanup()

    def read(self, path, name):
        with zipfile.ZipFile(path) as z:
            return z.read(name).decode("utf-8")

    def test_text_and_dates(self):
        report = roll(self.src, self.dst, PAIRS, shift_months=1)
        got = {f"{c.cell}": (c.before, c.after) for c in report.changes}
        self.assertEqual(got["A1"], ("2026年9月分 請求書発行資料", "2026年10月分 請求書発行資料"))
        self.assertEqual(got["B2"], ("2026/09/01-2026/09/30", "2026/10/01-2026/10/31"))
        self.assertEqual(got["B3"], ("2026-10-05", "2026-11-05"))  # 独自の日付の表示形式
        self.assertEqual(got["B4"], ("2026-10-31", "2026-11-30"))  # 組み込みの日付の表示形式・月末
        self.assertNotIn("B5", got)  # 金額（桁区切り）は日付ではない
        self.assertNotIn("B6", got)  # 式のセルは触らない

    def test_untouched_parts_are_byte_identical(self):
        roll(self.src, self.dst, PAIRS, shift_months=1)
        for name in ["xl/styles.xml", "xl/workbook.xml", "[Content_Types].xml"]:
            self.assertEqual(self.read(self.src, name), self.read(self.dst, name))
        sheet = self.read(self.dst, "xl/worksheets/sheet1.xml")
        self.assertIn('<col min="1" max="1" width="18.5" customWidth="1"/>', sheet)  # 列幅
        self.assertIn('<pageSetup paperSize="9" orientation="landscape"/>', sheet)  # 印刷設定
        self.assertIn("<f>B5*1.1</f>", sheet)

    def test_furigana_is_not_rewritten(self):
        roll(self.src, self.dst, PAIRS)
        shared = self.read(self.dst, "xl/sharedStrings.xml")
        self.assertIn("<t>10月分の業務委託費</t>", shared)
        self.assertIn("<rPh sb=\"0\" eb=\"1\"><t>ガツ</t></rPh>", shared)

    def test_dry_run_writes_nothing_and_check_finds_leftovers(self):
        report = roll(self.src, None, [("2026年9月", "2026年10月")], check_words=["9月"])
        self.assertFalse(self.dst.exists())
        self.assertEqual([c.cell for c in report.leftovers], ["A5"])  # 「9月分の業務委託費」が残っている

    def test_cli_refuses_to_overwrite_source(self):
        with self.assertRaises(SystemExit):
            main([str(self.src), str(self.src), "--replace", "9月=10月"])


if __name__ == "__main__":
    unittest.main()
