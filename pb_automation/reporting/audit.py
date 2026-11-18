"""Serializable arithmetic audit, independent of workbook layout or rounding."""
from .data import build_presentation_data
from .report import presentation_history
from .model import complete_sum, ratio


def flow_average_audit(history, period, config):
    data=build_presentation_data(presentation_history(history,period),period,config)
    countries=[]
    for c in data['countries']:
        incoming=[r['amount'] for r in c['monthly']]
        counts=[r['count'] for r in c['monthly']]
        total=complete_sum(incoming)
        mean=ratio(total,6)
        countries.append({'country':c['id'],'name':c['name'],'currency':c['currency'],
            'months':[r['month'] for r in c['monthly']], 'incoming_by_month_local':incoming,
            'months_available':sum(v is not None for v in incoming),'six_month_incoming_local':total,
            'average_monthly_local':mean,'report_average_local':c['metrics']['monthly_average_local'],
            'local_units_per_EUR':c['rate'],'average_monthly_EUR':ratio(mean,c['rate']),
            'incoming_contracts':complete_sum(counts),
            'average_balance_per_contract_local':ratio(total,complete_sum(counts)),
            'complete':all(v is not None for v in incoming)})
    return {'period':period,'definition':'Sum of the six calendar monthly incoming amounts / 6; includes false positives. A missing month is unknown, a reported zero is zero.',
            'country_pages':'Local currency unless explicitly configured otherwise.',
            'group_pages':'EUR: convert each country at current-quarter FX before summing.',
            'countries':countries}
