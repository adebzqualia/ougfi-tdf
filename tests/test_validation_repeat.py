import unittest
from test_pipeline import sample_snapshot,country_config
from pb_automation.validate import validate_snapshot


class ValidationRepeatTests(unittest.TestCase):
    def test_revalidation_does_not_duplicate_reconciliation_findings(self):
        snapshot=sample_snapshot()
        snapshot['entities'][0]['flows'][0]['processed_amount']=50.1
        first,report=validate_snapshot(snapshot,country_config('FR'))
        self.assertTrue(report['warnings'])
        second,again=validate_snapshot(first,country_config('FR'))
        self.assertEqual(first,second)
        self.assertEqual(report['snapshot_sha256'],again['snapshot_sha256'])
        self.assertEqual(report['warnings'],again['warnings'])
