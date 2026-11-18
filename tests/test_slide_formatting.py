import unittest
from pb_automation.reporting.formatting import slide_amount,slide_percent,stock_axis

class SlideFormattingTests(unittest.TestCase):
    def test_commentary_uses_compact_amounts(self):
        self.assertEqual(slide_amount(1574089,'EUR'),'1.6m EUR')
        self.assertEqual(slide_amount(86982,'EUR'),'87k EUR')
        self.assertEqual(slide_amount(2883642,'EUR'),'2.9m EUR')
        self.assertEqual(slide_amount(0,'EUR'),'0 EUR')

    def test_average_balance_rounds_to_nearest_ten(self):
        self.assertEqual(slide_amount(202.32,'EUR',True),'200 EUR')
        self.assertEqual(slide_amount(235,'EUR',True),'240 EUR')
        self.assertEqual(slide_percent(-.131),'-13%')

    def test_shared_axis_covers_both_stacks(self):
        rows=[{'vintages':[{'identified_eur':2000000,'remaining_eur':1000000},
                           {'identified_eur':1483914,'remaining_eur':1000000}]}]
        self.assertEqual(stock_axis(rows),{'min':0,'max':4000000,'major_unit':500000})
        rows[0]['vintages'][0]['remaining_eur']=9000000
        self.assertGreaterEqual(stock_axis(rows)['max'],10000000)

if __name__=='__main__':unittest.main()
