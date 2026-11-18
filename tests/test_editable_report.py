"""Regression checks for the workbook's edit and reporting boundaries."""
import json
from pathlib import Path
import tempfile
import unittest

from openpyxl import load_workbook
from pb_automation.reporting import generate_report


class EditableReportTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        root=Path(__file__).resolve().parents[1]
        cls.history=json.loads((root/'database/history.json').read_text(encoding='utf-8'))
        cls.before=json.dumps(cls.history,sort_keys=True)
        cls.config=json.loads((root/'config/presentation.json').read_text(encoding='utf-8'))
        cls.temp=tempfile.TemporaryDirectory()
        cls.path=Path(cls.temp.name)/'editable.xlsx'
        generate_report(cls.history,'2026Q3',cls.path,cls.config)
        cls.formulas=load_workbook(cls.path,data_only=False)
        cls.values=load_workbook(cls.path,data_only=True)

    @classmethod
    def tearDownClass(cls):
        cls.formulas.close();cls.values.close();cls.temp.cleanup()

    def test_generation_does_not_revalue_saved_history(self):
        self.assertEqual(self.before,json.dumps(self.history,sort_keys=True))

    def test_reported_totals_and_rounding_gap_are_retained(self):
        f=self.values['FR France']
        self.assertEqual(f['Z12'].value,2883642)
        self.assertEqual(sum(f[f'Z{r}'].value for r in range(7,12)),2883641)
        self.assertEqual(f['X12'].value,16614)

    def test_only_inputs_unlocked_and_no_external_links(self):
        for name in ('Stock','Flux','Repartition stock','Taux','Other inputs','Comments'):
            self.assertEqual(self.formulas[name].sheet_state,'visible')
            self.assertTrue(self.formulas[name].protection.sheet)
        self.assertFalse(self.formulas['Stock']['F116'].protection.locked)
        self.assertTrue(self.formulas['Stock']['H116'].protection.locked)
        self.assertFalse(self.formulas._external_links)
        self.assertEqual(sum(len(s._charts) for s in self.formulas),21)

    def test_combined_efficiency_uses_oldest_month_inputs(self):
        formulas=[c.value for row in self.formulas['GraphData'] for c in row if c.data_type=='f']
        self.assertTrue(any("'Flux'!$R$156" in f and "'Flux'!$R$150" in f for f in formulas))
        self.assertAlmostEqual(self.values['CORE Flows']['W26'].value,(661018+6180)/(775286+6180))


if __name__=='__main__':unittest.main()
