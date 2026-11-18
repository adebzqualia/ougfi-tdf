"""Safety checks for quarter boundaries, validation and historical updates."""
import copy
import json
import tempfile
import unittest
from pathlib import Path

from pb_automation.common import parse_period, quarter_end, rolling_months, snapshot_hash
from pb_automation.validate import validate_snapshot
from pb_automation.history import update_history
from pb_automation.extract import ExcelReader, _select, month_key
from pb_automation.reporting.report import presentation_history
from pb_automation.reporting.data import _global_vintages


def country_config(*identifiers):
    return {'countries':[
        {'id':eid,'name':{'FR':'France','RO':'Romania'}.get(eid,eid),'currency':'EUR',
         'profile':'standard' if eid=='FR' else 'runoff','enabled':True,
         'requires_cst':False,'processed_fields':['refund','unreachable','false_positive'] if eid=='FR' else [],
         'identified_addback_fields':['unreachable'] if eid=='FR' else []}
        for eid in identifiers]}


def sample_snapshot(period='2026Q3'):
    def stock(vintage,amount):
        return dict(vintage=vintage,total=amount,processed_source=amount*.8,
                    refund=amount*.6,unreachable=amount*.2,false_positive=0,
                    authority_refund=None,other_without_refund=None,
                    remaining_source=amount*.2,source='synthetic test fixture')
    total=stock('Total',1000)
    entity=dict(id='FR',name='France',currency='EUR',profile='standard',requires_cst=False,
        processed_fields=['refund','unreachable','false_positive'],identified_addback_fields=['unreachable'],
        stock_basis='source_remaining',flow_basis='source_processed',use_source_stock_totals=True,
        stock_count_basis='stock_vintages',prefer_reported_distribution_totals=True,
        stock=[stock('<2025',0),stock('2025',0),stock('2026',1000)],stock_totals=total,
        flows=[dict(month=m,amount=100,count=10,processed_amount=50,processed_count=5,
                    refund=40,unreachable=10,false_positive=0,authority_refund=None,other_without_refund=None,
                    remaining_amount_source=50,remaining_count_source=5,source='synthetic test fixture')
               for m in rolling_months(period)],
        stock_bins=[dict(band='100 - 1 000 EUR',count=4,amount=400,source='synthetic test fixture')],
        stock_bins_totals=dict(count=4,amount=400,source='synthetic test fixture'),
        stock_vintages=[dict(vintage='2026',count=4,amount=400,source='synthetic test fixture')],
        stock_vintages_totals=dict(count=4,amount=400,source='synthetic test fixture'),
        historical_metrics={},historical_vintages=[],source='synthetic test fixture')
    return dict(period=period,entities=[entity],fx=[dict(currency='EUR',rate=1,
        date=quarter_end(period).isoformat(),source='identity')],issues=[],sources=[])


class QuarterTests(unittest.TestCase):
    def test_folder_and_canonical_periods(self):
        self.assertEqual(parse_period('Q4_2026'),'2026Q4')
        self.assertEqual(parse_period('2026Q4'),'2026Q4')

    def test_rolling_semester_crosses_year_boundary(self):
        self.assertEqual(rolling_months('2027Q1'),[
            '2026-10','2026-11','2026-12','2027-01','2027-02','2027-03'])
        self.assertEqual(quarter_end('2027Q1').isoformat(),'2027-03-31')

    def test_invalid_quarter_rejected(self):
        for period in ['Q0_2026','Q5_2026','2026Q9','quarter four']:
            with self.subTest(period=period),self.assertRaises(ValueError):
                parse_period(period)

    def test_month_headers_are_language_independent(self):
        self.assertEqual(month_key('août-26'), '2026-08')
        self.assertEqual(month_key('Apr-26'), '2026-04')

    def test_duplicate_country_files_stop_extraction(self):
        with self.assertRaisesRegex(ValueError, 'Duplicate inputs'):
            _select([Path('FRANCE.xlsx'), Path('FRANCE_Q32026.xlsx')],
                    ['FRANCE.xlsx', 'FRANCE_Q{quarter}{year}.xlsx'], [], '2026Q3', 'France')


class PresentationTests(unittest.TestCase):
    def test_current_fx_applies_to_both_quarters_without_changing_history(self):
        h={'snapshots':{q:sample_snapshot(q) for q in ['2026Q3','2026Q4']}}
        for q,rate in [('2026Q3',1.0),('2026Q4',0.8)]:
            h['snapshots'][q]['entities'][0]['currency']='GBP'
            h['snapshots'][q]['fx']=[dict(currency='GBP',rate=rate)]
        h['snapshots']['2026Q3']['entities'][0]['historical_metrics']={'identified_eur':100}
        before=copy.deepcopy(h)
        converted=presentation_history(h,'2026Q4')
        self.assertEqual(converted['snapshots']['2026Q3']['entities'][0]['historical_metrics']['identified_eur'],125)
        self.assertEqual(h,before)

    def test_brazil_new_extraction_populates_consolidated_vintages(self):
        country={'id':'BR','profile':'brazil','stock':[], 'stock_origin':[
            {'vintage':'2024','amount':100},{'vintage':'2025','amount':200},{'vintage':'2026','amount':300}]}
        rows=_global_vintages(country,5,'2026Q3')
        self.assertEqual([r['identified_eur'] for r in rows],[20,40,60])

    def test_q1_comparison_selects_previous_year_q4(self):
        h={'snapshots':{q:sample_snapshot(q) for q in ['2026Q4','2027Q1']}}
        self.assertEqual(set(presentation_history(h,'2027Q1')['snapshots']),{'2026Q4','2027Q1'})


class ValidationTests(unittest.TestCase):
    def test_valid_complete_fixture(self):
        snapshot=sample_snapshot()
        validated,report=validate_snapshot(snapshot,country_config('FR'))
        self.assertTrue(report['valid'],report)
        self.assertFalse(report['errors'])
        self.assertEqual(report['snapshot_sha256'],snapshot_hash(validated))

    def test_duplicate_country_rejected(self):
        snapshot=sample_snapshot()
        snapshot['entities'].append(copy.deepcopy(snapshot['entities'][0]))
        _,report=validate_snapshot(snapshot,country_config('FR'))
        self.assertFalse(report['valid'])
        self.assertTrue(report['errors'])

    def test_missing_country_is_explicit_not_carried_forward(self):
        snapshot=sample_snapshot()
        _,report=validate_snapshot(snapshot,country_config('FR','RO'))
        self.assertFalse(report['valid'])
        validated,partial=validate_snapshot(snapshot,country_config('FR','RO'),allow_partial=True)
        self.assertTrue(partial['valid'],partial)
        self.assertEqual([e['id'] for e in validated['entities']],['FR'])
        self.assertIn('RO',partial['coverage']['missing'])

    def test_confirmed_zero_stock_does_not_invent_monthly_flows(self):
        override={'period':'2026Q3','countries':{'RO':{'stock_zero':True,'reason':'Confirmed by reporting owner'}}}
        validated,report=validate_snapshot(sample_snapshot(),country_config('FR','RO'),override)
        self.assertTrue(report['valid'],report)
        romania=next(e for e in validated['entities'] if e['id']=='RO')
        self.assertEqual(romania['historical_metrics']['identified'],0)
        self.assertEqual(romania['historical_metrics']['remaining'],0)
        self.assertEqual(romania['flows'],[])

    def test_override_for_different_quarter_rejected(self):
        override={'period':'2026Q2','countries':{'RO':{'stock_zero':True,'reason':'Previous quarter only'}}}
        _,report=validate_snapshot(sample_snapshot(),country_config('FR','RO'),override)
        self.assertFalse(report['valid'])

    def test_required_currency_rate_must_be_positive(self):
        for rates in [[],[{'currency':'GBP','rate':0,'date':'2026-09-30','source':'test'}]]:
            with self.subTest(rates=rates):
                snapshot=sample_snapshot();snapshot['fx']=rates;snapshot['entities'][0]['currency']='GBP'
                config=country_config('FR');config['countries'][0]['currency']='GBP'
                _,report=validate_snapshot(snapshot,config)
                self.assertFalse(report['valid'])


class HistoryTests(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory()
        self.path=Path(self.tmp.name)/'database'/'history.json'

    def tearDown(self):
        self.tmp.cleanup()

    def accepted(self,period='2026Q3'):
        return validate_snapshot(sample_snapshot(period),country_config('FR'))

    def test_same_quarter_is_idempotent(self):
        snapshot,report=self.accepted()
        first=update_history(self.path,snapshot,report)
        original=self.path.read_bytes()
        second=update_history(self.path,snapshot,report)
        self.assertEqual(first['status'],'added')
        self.assertEqual(second['status'],'unchanged')
        self.assertEqual(self.path.read_bytes(),original)
        self.assertEqual(list(json.loads(original)['snapshots']),['2026Q3'])

    def test_validation_failure_cannot_write_history(self):
        snapshot,report=self.accepted()
        report['valid']=False
        with self.assertRaises(ValueError):update_history(self.path,snapshot,report)
        self.assertFalse(self.path.exists())

    def test_changed_data_requires_revalidation(self):
        snapshot,report=self.accepted()
        snapshot['entities'][0]['stock_totals']['remaining_source']=999
        with self.assertRaises(ValueError):update_history(self.path,snapshot,report)
        self.assertFalse(self.path.exists())

    def test_later_quarter_preserves_previous_and_creates_backup(self):
        snapshot,report=self.accepted('2026Q3')
        update_history(self.path,snapshot,report)
        before=self.path.read_bytes()
        next_snapshot,next_report=self.accepted('2026Q4')
        result=update_history(self.path,next_snapshot,next_report)
        saved=json.loads(self.path.read_text(encoding='utf-8'))
        self.assertEqual(set(saved['snapshots']),{'2026Q3','2026Q4'})
        self.assertEqual(saved['snapshots']['2026Q3'],snapshot)
        self.assertEqual(Path(result['backup']).read_bytes(),before)

    def test_older_snapshot_change_needs_explicit_flag(self):
        snapshot,report=self.accepted('2026Q3');update_history(self.path,snapshot,report)
        newer,newer_report=self.accepted('2026Q4');update_history(self.path,newer,newer_report)
        changed=sample_snapshot('2026Q3')
        changed['entities'][0]['flows'][0]['amount']=101
        changed['entities'][0]['flows'][0]['remaining_amount_source']=51
        changed,changed_report=validate_snapshot(changed,country_config('FR'))
        before=self.path.read_bytes()
        with self.assertRaises(ValueError):update_history(self.path,changed,changed_report)
        self.assertEqual(self.path.read_bytes(),before)
        result=update_history(self.path,changed,changed_report,replace_historical=True)
        self.assertEqual(result['status'],'replaced')


if __name__=='__main__':
    unittest.main()
