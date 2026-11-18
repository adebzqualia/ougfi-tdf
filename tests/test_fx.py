"""Regression checks for dated FX tabs and the earlier rate-file layouts."""
import json
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook
from pb_automation.extract import _fx


class ExchangeRateTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        config = Path(__file__).resolve().parents[1] / 'config/layouts.json'
        self.profiles = json.loads(config.read_text(encoding='utf-8'))['fx_profiles']

    def tearDown(self):
        self.tmp.cleanup()

    def fixture(self, tabs, filename='TAUX.xlsx'):
        book = Workbook()
        book.remove(book.active)
        for name, rate in tabs:
            sheet = book.create_sheet(name)
            sheet['F2'] = 'TC\nTaux de clôture\nClosing Rate'
            sheet['B6'] = 'GBP'
            sheet['E6'] = 9  # Opening rate: must not be selected.
            sheet['F6'] = rate
            sheet['G6'] = 8  # Average rate: must not be selected.
        path = self.root / filename
        book.save(path)
        book.close()
        return path

    def read_rates(self, path):
        return _fx(path, {'allow_period_in_filename': True}, self.profiles,
                   '2026Q3', {'EUR', 'GBP'}, [])

    def test_dated_sheet_supplies_period_and_closing_rate(self):
        rows = self.read_rates(self.fixture([('2026.09', 0.85443381)]))
        gbp = next(row for row in rows if row['currency'] == 'GBP')
        self.assertEqual(gbp['rate'], 0.85443381)
        self.assertEqual(gbp['date'], '2026-09-30')
        self.assertEqual(gbp['date_evidence'], 'dated worksheet name')

    def test_multiple_months_select_requested_quarter(self):
        rows = self.read_rates(self.fixture([('2026.06', 0.9), ('2026.09', 0.85)]))
        self.assertEqual(next(row['rate'] for row in rows if row['currency'] == 'GBP'), 0.85)

    def test_previous_quarter_is_not_silently_reused(self):
        with self.assertRaisesRegex(ValueError, '2026Q3'):
            self.read_rates(self.fixture([('2026.06', 0.9)]))

    def test_ambiguous_current_tables_stop_run(self):
        with self.assertRaisesRegex(ValueError, 'found 2'):
            self.read_rates(self.fixture([('2026.09', 0.85), ('Sheet1', 0.9)]))

    def test_earlier_generic_sheet_and_filename_still_work(self):
        rows = self.read_rates(self.fixture([('Sheet1', 0.85)], 'taux_q32026.xlsx'))
        self.assertEqual(next(row['rate'] for row in rows if row['currency'] == 'GBP'), 0.85)


if __name__ == '__main__':
    unittest.main()
