"""A new quarter appends history; corrections create auditable revisions."""
import json
from pathlib import Path
import tempfile
import unittest

from pb_automation.common import atomic_json, snapshot_hash
from pb_automation.history import update_history


class HistoryVersioningTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory()
        self.path=Path(self.temp.name)/'history.json'

    def tearDown(self):
        self.temp.cleanup()

    def save(self,period,amount=100,**kwargs):
        snapshot={'period':period,'entities':[{'id':'FR','amount':amount}],'fx':[{'currency':'EUR','rate':1}]}
        report={'valid':True,'errors':[],'snapshot_sha256':snapshot_hash(snapshot)}
        return update_history(self.path,snapshot,report,**kwargs)

    def read(self):
        return json.loads(self.path.read_text(encoding='utf-8'))

    def test_revision_chain_and_identical_rerun(self):
        a=self.save('2026Q3'); original=self.read()
        self.assertEqual(a['revision'],1)
        same=self.save('2026Q3')
        self.assertEqual(same['status'],'unchanged')
        self.assertEqual(self.read(),original)
        b=self.save('2026Q3',200);data=self.read()
        self.assertEqual(b['revision'],2)
        self.assertEqual(data['revisions'][1]['previous_snapshot_sha256'],data['revisions'][0]['snapshot_sha256'])
        backup=self.path.parent/data['revisions'][1]['backup']
        self.assertEqual(json.loads(backup.read_text(encoding='utf-8')),original)
        self.assertEqual(data['snapshots']['2026Q3']['entities'][0]['amount'],200)

    def test_all_four_quarters_and_year_rollover(self):
        periods=['2026Q1','2026Q2','2026Q3','2026Q4','2027Q1']
        for i,period in enumerate(periods):self.save(period,i)
        data=self.read()
        self.assertEqual(list(data['snapshots']),periods)
        self.assertEqual([data['snapshots'][p]['entities'][0]['amount'] for p in periods],list(range(5)))
        self.assertEqual([r['revision'] for r in data['revisions']],[1]*5)

    def test_older_revision_keeps_later_snapshot(self):
        self.save('2026Q4',100);self.save('2027Q1',200)
        before=self.read()
        with self.assertRaisesRegex(ValueError,'older'):self.save('2026Q4',110)
        self.assertEqual(self.read(),before)
        result=self.save('2026Q4',110,replace_historical=True)
        self.assertEqual(result['revision'],2)
        self.assertEqual(self.read()['snapshots']['2027Q1'],before['snapshots']['2027Q1'])

    def test_legacy_history_gains_revision_metadata_without_rewriting_snapshot(self):
        self.save('2026Q3');legacy=self.read();legacy.pop('revisions')
        atomic_json(self.path,legacy)
        result=self.save('2026Q3',101)
        self.assertEqual(result['revision'],2)
        self.assertEqual(json.loads(Path(result['backup']).read_text(encoding='utf-8')),legacy)


if __name__=='__main__':unittest.main()
