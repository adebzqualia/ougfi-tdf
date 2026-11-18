"""Synthetic contract-level files: independent of Excel formula caches."""
from copy import deepcopy
from datetime import datetime
from pathlib import Path
import json
import tempfile
import unittest
import openpyxl

from pb_automation.base import select_base
from pb_automation.extract import ExcelReader


class BaseExtractionTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory();self.path=Path(self.temp.name)/'country.xlsx'
        root=Path(__file__).resolve().parents[1]/'config'
        all_countries=json.loads((root/'countries.json').read_text())
        self.cfg=deepcopy(next(x for x in all_countries['countries'] if x['id']=='DK'))
        self.layout=json.loads((root/'layouts.json').read_text())['nordics']
        profile=deepcopy(json.loads((root/'base.json').read_text())['profiles']['nordics_v1'])
        profile.update(decision_before_origin='error',duplicate_key_policy='error')
        profile['amount_bands']={'basis':'local','upper_bounds':[1,10,100,1000]}
        self.settings={'countries':{'DK':{'mode':'base_if_present','profile':'test','compare_kpi':True}},'profiles':{'test':profile}}
        self.profile=profile

    def tearDown(self):self.temp.cleanup()

    def write(self,rows,period='2026Q3',header=None,sheet='Base'):
        book=openpyxl.Workbook();base=book.active;base.title=sheet
        headers=header or ['Contract number','Amount in local currency','Date of positive balance','Date of decision','Decision status','Country']
        base.append([]);base.append([]);base.append([]);base.append(headers)
        for row in rows:base.append(row)
        metadata=book.create_sheet('KPI_1')
        from pb_automation.common import quarter_end
        metadata['G4']=quarter_end(period)
        # Deliberately no cached values for formulas, and no country KPI tabs.
        kpi=book.create_sheet('KPI_2');kpi['B11']='=SUM(Base!B5:B100)'
        params=book.create_sheet('Parameters');params.append(['Date','SEK','DKK','NOK'])
        params.append([quarter_end(period),.1,.5,.1])
        book.save(self.path);book.close()

    def extract(self,period='2026Q3'):
        reader=ExcelReader(self.path);issues=[]
        try:result=select_base(reader,self.cfg,self.layout,period,issues,self.settings)
        finally:reader.close()
        return result,issues

    @staticmethod
    def row(key='a',amount=100,origin='2026-04-01',decision=None,status='',country='Denmark'):
        dt=lambda s:datetime.fromisoformat(s) if s else None
        return [key,amount,dt(origin),dt(decision),status,country]

    def test_base_works_without_kpi_formula_caches(self):
        self.write([self.row()]);e,issues=self.extract()
        self.assertEqual(e['stock_totals']['remaining_source'],100)
        self.assertEqual(e['extraction_method'],'base')
        self.assertIn('BASE_KPI_COMPARISON_UNAVAILABLE',[i['code'] for i in issues])
        self.assertNotIn('Contract number',json.dumps(e))

    def test_cutoffs_and_future_decisions(self):
        self.write([self.row('a',10,decision='2026-04-02',status='Refund_C'),
                    self.row('b',20,'2026-09-30','2026-10-01','Refund_C'),
                    self.row('c',30,'2026-10-01'),self.row('d',40,'2026-03-31')])
        e,issues=self.extract();self.assertEqual(e['stock_totals']['total'],70)
        self.assertEqual(e['stock_totals']['remaining_source'],60)
        self.assertEqual(sum(r['amount'] for r in e['flows']),30)
        self.assertEqual(sum(r['processed_count'] for r in e['flows']),1)
        self.assertEqual(e['base_audit']['future_origins_excluded'],1)
        self.assertEqual(e['base_audit']['future_decisions_pending'],1)

    def test_q1_rolls_across_year_and_q4_has_december(self):
        for period,months in [('2027Q1',['2026-10','2026-11','2026-12','2027-01','2027-02','2027-03']),
                              ('2026Q4',['2026-07','2026-08','2026-09','2026-10','2026-11','2026-12'])]:
            self.write([self.row(str(i),10,m+'-01') for i,m in enumerate(months)],period)
            e,_=self.extract(period);self.assertEqual([r['month'] for r in e['flows']],months)
            self.assertEqual(sum(r['amount'] for r in e['flows'])/6,10)

    def test_bands_partition_exact_boundaries_for_amount_and_count(self):
        self.write([self.row(str(i),n) for i,n in enumerate([0,1,10,100,1000,1001])])
        e,_=self.extract()
        self.assertEqual([r['count'] for r in e['stock_bins']],[2,1,1,1,1])
        self.assertEqual(sum(r['amount'] for r in e['stock_bins']),2112)
        self.assertEqual(e['stock_bins_totals']['count'],6)

    def test_currency_band_conversion_not_amount_conversion(self):
        self.profile['amount_bands']={'basis':'eur_equivalent','upper_bounds':[1,10,100,1000],
            'rate_source':{'sheet':'Parameters','date_column':1,'first_row':2,'last_row':8,'currency_columns':{'DKK':3}}}
        self.write([self.row('a',2),self.row('b',3)])
        e,_=self.extract();self.assertEqual(e['stock_bins'][0]['amount'],2)
        self.assertEqual(e['stock_bins'][0]['count'],1)
        self.assertEqual(e['stock_bins'][0]['band'],'0 - 1 EUR eq.')
        self.assertEqual(e['stock_bins_totals']['amount'],5)

    def test_units_scale_amounts_not_counts(self):
        self.cfg['units']=1000;self.write([self.row(amount=1)])
        e,_=self.extract();self.assertEqual(e['stock_bins_totals']['amount'],1000)
        self.assertEqual(e['stock_bins_totals']['count'],1)

    def test_unknown_status_and_country_fail(self):
        for row in [self.row(status='NEW STATUS'),self.row(country='Denmrk')]:
            self.write([row])
            with self.assertRaises(ValueError):self.extract()

    def test_negative_and_missing_amounts_fail(self):
        for value in (-1,None):
            self.write([self.row(amount=value)])
            with self.assertRaises(ValueError):self.extract()

    def test_base_formula_is_not_treated_as_raw_or_zero(self):
        self.write([self.row(amount='=1+1')])
        with self.assertRaisesRegex(ValueError,'contains a formula'):self.extract()

    def test_duplicate_exact_record_errors_even_when_keys_allowed(self):
        self.profile['duplicate_key_policy']='warn';self.write([self.row(),self.row()])
        with self.assertRaisesRegex(ValueError,'duplicate Base record'):self.extract()

    def test_repeated_keys_use_explicit_row_counting_policy(self):
        self.write([self.row(amount=1),self.row(amount=2)])
        with self.assertRaisesRegex(ValueError,'duplicate .*key'):self.extract()
        self.profile['duplicate_key_policy']='warn';e,issues=self.extract()
        self.assertEqual(e['stock_bins_totals']['count'],2)
        self.assertIn('BASE_REPEATED_KEY',[i['code'] for i in issues])

    def test_missing_or_invalid_decision_date_requires_explicit_policy(self):
        self.write([self.row(status='Refund_C')])
        with self.assertRaisesRegex(ValueError,'no decision date'):self.extract()
        self.profile['missing_decision_date']='snapshot_status';e,issues=self.extract()
        self.assertEqual(e['stock_totals']['processed_source'],100)
        self.assertIn('BASE_MISSING_DECISION_DATE',[i['code'] for i in issues])
        row=self.row(status='Refund_C');row[3]='31.11.26';self.write([row])
        with self.assertRaisesRegex(ValueError,'invalid'):self.extract()
        self.profile['invalid_decision_date']='snapshot_status';_,issues=self.extract()
        self.assertIn('BASE_INVALID_DECISION_DATE',[i['code'] for i in issues])

    def test_decision_before_origin_is_audited(self):
        self.write([self.row(decision='2026-03-31',status='Refund_C')])
        with self.assertRaisesRegex(ValueError,'precedes'):self.extract()
        self.profile['decision_before_origin']='warn';e,issues=self.extract()
        self.assertEqual(e['base_audit']['decisions_before_origin'],1)

    def test_kpi3_scope_exclusion_does_not_rewrite_kpi2(self):
        self.profile['kpi3_exclusions']=[{'field':'contract','values':['closed']}]
        self.write([self.row('closed'),self.row('open',200)])
        e,_=self.extract();self.assertEqual(e['stock_totals']['remaining_source'],300)
        self.assertEqual(e['stock_bins_totals']['amount'],200)

    def test_custom_sheet_name_is_used_consistently(self):
        self.profile['sheet']='Data';self.write([self.row()],sheet='Data')
        e,_=self.extract();self.assertEqual(e['extraction_method'],'base')

    def test_unmapped_present_base_errors_absent_base_falls_back(self):
        self.settings['countries']['DK']['profile']=None;self.write([self.row()])
        with self.assertRaisesRegex(ValueError,'mapping is unverified'):self.extract()
        self.write([self.row()],sheet='Unused');e,issues=self.extract()
        self.assertIsNone(e);self.assertEqual(issues[0]['code'],'BASE_ABSENT_KPI_FALLBACK')

    def test_duplicate_headers_and_invalid_config_fail(self):
        self.write([self.row()],header=['Contract number','Amount in local currency','Date of positive balance','Date of decision','Decision status','Amount in local currency'])
        with self.assertRaisesRegex(ValueError,'matching header'):self.extract()
        self.write([self.row()]);self.profile['amount_bands']['basis']='typo'
        with self.assertRaisesRegex(ValueError,'basis'):self.extract()


if __name__=='__main__':unittest.main()
