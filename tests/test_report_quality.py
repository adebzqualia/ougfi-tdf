"""Business-rule regressions: six calendar months, local bands and 100% axes."""
import copy
import json
from pathlib import Path
import tempfile
import unittest
import zipfile
import xml.etree.ElementTree as ET

from openpyxl import load_workbook
from pb_automation.reporting.data import _monthly, _metrics, build_presentation_data
from pb_automation.reporting.model import rolling_months
from pb_automation.reporting.report import generate_report, presentation_history
from test_pipeline import sample_snapshot


class MeanTests(unittest.TestCase):
    def metrics(self,period,values):
        e=sample_snapshot(period)['entities'][0]
        e['flows']=sorted(e['flows'],key=lambda r:r['month'])
        for row,value in zip(e['flows'],values):row['amount']=value
        months=sorted(rolling_months(period))
        return _metrics(e,_monthly(e,2,months),2,period)

    def test_all_quarters_use_six_months_including_zero(self):
        for p in ('2026Q1','2026Q2','2026Q3','2026Q4','2027Q1'):
            with self.subTest(period=p):
                m=self.metrics(p,[0,100,200,300,400,500])
                self.assertEqual(m['monthly_average_local'],250)
                self.assertEqual(m['monthly_average'],125)
                self.assertEqual(m['monthly_average_available_months'],6)
                self.assertEqual(m['monthly_average_months'],sorted(rolling_months(p)))

    def test_missing_month_is_not_zero_or_average_of_five(self):
        m=self.metrics('2026Q3',[None,100,200,300,400,500])
        self.assertIsNone(m['monthly_average_local'])
        self.assertIsNone(m['monthly_average'])
        self.assertEqual(m['monthly_average_available_months'],5)

    def test_duplicate_month_rejected_before_dictionary_can_discard_it(self):
        e=sample_snapshot()['entities'][0];e['flows'].append(copy.deepcopy(e['flows'][0]))
        with self.assertRaisesRegex(ValueError,'duplicate cohort'):
            _monthly(e,1,sorted(rolling_months('2026Q3')))

    def test_out_of_window_flow_excluded(self):
        e=sample_snapshot()['entities'][0]
        outside=copy.deepcopy(e['flows'][0]);outside.update(month='2026-10',amount=999999)
        e['flows'].append(outside)
        m=_metrics(e,_monthly(e,1,sorted(rolling_months('2026Q3'))),1,'2026Q3')
        self.assertEqual(m['monthly_average_local'],100)


class WorkbookQualityTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        root=Path(__file__).resolve().parents[1]
        cls.history=json.loads((root/'database/history.json').read_text(encoding='utf-8'))
        cls.config=json.loads((root/'config/presentation.json').read_text(encoding='utf-8'))
        cls.config['additional_amount_band_countries']=['UK','ZA']
        for eid in ('UK','ZA'):
            cls.config['country_options'][eid]['show_stock_bins']=False
        cls.temp=tempfile.TemporaryDirectory();cls.path=Path(cls.temp.name)/'report.xlsx'
        generate_report(cls.history,'2026Q3',cls.path,cls.config)
        cls.values=load_workbook(cls.path,data_only=True)
        cls.formulas=load_workbook(cls.path,data_only=False)

    @classmethod
    def tearDownClass(cls):
        cls.values.close();cls.formulas.close();cls.temp.cleanup()

    def test_all_efficiency_chart_axes_are_zero_to_one(self):
        ns={'c':'http://schemas.openxmlformats.org/drawingml/2006/chart'}
        found=0
        with zipfile.ZipFile(self.path) as archive:
            for name in archive.namelist():
                if not name.startswith('xl/charts/chart') or not name.endswith('.xml'):continue
                root=ET.fromstring(archive.read(name))
                for axis in root.findall('.//c:valAx',ns):
                    fmt=axis.find('c:numFmt',ns)
                    if fmt is not None and '%' in fmt.attrib.get('formatCode',''):
                        self.assertEqual(axis.find('c:scaling/c:min',ns).get('val'),'0')
                        self.assertEqual(axis.find('c:scaling/c:max',ns).get('val'),'1')
                        found+=1
        self.assertGreaterEqual(found,12)

    def test_source_bands_not_relabelled_or_fabricated(self):
        sh=self.values['UK ZA Amount bands']
        values=[cell.value for row in sh for cell in row]
        self.assertIn('Amount GBP',values);self.assertIn('Amount ZAR',values)
        uk=next(e for e in self.history['snapshots']['2026Q3']['entities'] if e['id']=='UK')
        za=next(e for e in self.history['snapshots']['2026Q3']['entities'] if e['id']=='ZA')
        self.assertIn(uk['stock_bins'][0]['amount'],values)
        self.assertIn(za['stock_bins'][0]['amount'],values)
        for row in uk['stock_bins']+za['stock_bins']:self.assertIn(row['band'],values)
        self.assertTrue(any(cell.data_type=='f' and 'Repartition stock' in cell.value
                            for row in self.formulas['UK ZA Amount bands'] for cell in row))

    def test_group_mean_is_sum_of_country_eur_means(self):
        d=build_presentation_data(presentation_history(self.history,'2026Q3'),'2026Q3',self.config)
        core=next(g for g in d['groups'] if g['id']=='core')
        expected=sum(sum(m['amount_eur'] for m in c['monthly'])/6 for c in core['underlying_countries'])
        self.assertAlmostEqual(core['metrics']['monthly_average'],expected)
        fr=next(c for c in d['countries'] if c['id']=='FR')
        self.assertAlmostEqual(fr['metrics']['average_balance_local'],9444535/46681)

    def test_mean_calculation_trail_has_live_links_and_units(self):
        sh=self.values['Flow averages'];fs=self.formulas['Flow averages']
        row=next(row for row in sh.iter_rows(min_row=6) if row[0].value=='2026Q3' and row[1].value=='FR')
        self.assertEqual(row[3].value,'EUR')
        self.assertEqual(row[4].value,'2026-04');self.assertEqual(row[5].value,'2026-09')
        self.assertEqual(row[6].value,6);self.assertEqual(row[7].value,9444535)
        self.assertAlmostEqual(row[8].value,9444535/6)
        self.assertIn('Flux',fs.cell(row[0].row,8).value)

    def test_brazil_in_is_one_editable_source_for_flow_and_average(self):
        rows=[row for row in self.formulas['Flux'].iter_rows(min_row=6)
              if row[0].value=='2026Q3' and row[1].value=='BR']
        self.assertEqual(len(rows),6)
        for row in rows:
            self.assertEqual(row[5].data_type,'f')
            self.assertIn('Other inputs',row[5].value)
            self.assertTrue(row[5].protection.locked)
            self.assertIn('Other inputs',row[6].value)


if __name__=='__main__':unittest.main()
