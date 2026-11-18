"""Step 2: validate and apply explicit manual confirmations; do not write history."""
from copy import deepcopy
import math
from .common import parse_period,quarter_end,rolling_months,snapshot_hash

def validate_snapshot(snapshot,countries_config,overrides=None,allow_partial=False):
    data=deepcopy(snapshot);errors=[];warnings=[]
    def issue(code,message,country='',error=False):
        target=errors if error else warnings
        finding={'code':code,'entity':country,'message':message}
        # Revalidating an imported snapshot must not multiply the same finding
        # or create an artificial historical revision solely from duplicates.
        if finding not in target:target.append(finding)
    def numeric(v):return isinstance(v,(int,float)) and not isinstance(v,bool) and math.isfinite(v)
    try:period=parse_period(data.get('period',''))
    except ValueError as exc:
        return data,{'valid':False,'errors':[{'code':'PERIOD','message':str(exc)}],'warnings':[],'snapshot_sha256':'','coverage':{}}
    data['period']=period
    countries=countries_config.get('countries',[])
    configured={c['id']:c for c in countries if c.get('enabled',True)}
    if len(configured)!=sum(c.get('enabled',True) for c in countries):issue('DUPLICATE_CONFIG','Country configuration contains duplicate IDs.',error=True)
    entities=data.setdefault('entities',[])
    ids=[e.get('id') for e in entities]
    if len(ids)!=len(set(ids)):issue('DUPLICATE_COUNTRY','Two extracted entities have the same country ID.',error=True)
    for source_issue in data.get('issues',[]):
        if source_issue.get('code')=='MISSING_COUNTRY':continue
        issue(source_issue.get('code','EXTRACTION'),source_issue.get('message',''),source_issue.get('entity',''),source_issue.get('severity')=='error')
    overrides=overrides or {}
    if overrides.get('period') and parse_period(overrides['period'])!=period:issue('OVERRIDE_PERIOD','Manual confirmations belong to another quarter.',error=True)
    for eid,override in overrides.get('countries',{}).items():
        if eid not in configured:
            issue('UNKNOWN_OVERRIDE','Manual confirmation country is not enabled.',eid,True);continue
        if not str(override.get('reason','')).strip():
            issue('OVERRIDE_REASON','Every manual confirmation needs a reason.',eid,True);continue
        unknown=set(override)-{'stock_zero','reason','cst'}
        if unknown:issue('UNKNOWN_OVERRIDE',f'Unsupported override fields: {sorted(unknown)}',eid,True);continue
        existing=next((e for e in entities if e['id']==eid),None)
        if override.get('stock_zero'):
            if existing is not None:
                issue('OVERRIDE_CONFLICT','Stock-zero confirmation conflicts with a supplied country file. Remove one input.',eid,True);continue
            cfg=configured[eid];source='Manual confirmation: '+override['reason']
            bands=['0 - 1 €','1 - 10 €','10 - 100 €','100 - 1 000 €','> 1 000 €'] if cfg['currency']=='EUR' else []
            years=[f'<{int(period[:4])-1}',str(int(period[:4])-1),period[:4]]
            existing={'id':eid,'name':cfg['name'],'currency':cfg['currency'],'profile':'stock_only',
                'requires_cst':cfg.get('requires_cst',False),'processed_fields':[],'stock':[],'flows':[],
                'stock_bins':[{'band':b,'count':0,'amount':0,'source':source} for b in bands],
                'stock_vintages':[{'vintage':v,'count':0,'amount':0,'source':source} for v in years],
                'stock_bins_totals':{'count':0,'amount':0,'source':source},'stock_vintages_totals':{'count':0,'amount':0,'source':source},
                'historical_metrics':{'identified':0,'remaining':0,'contracts':0},
                'historical_vintages':[{'vintage':v,'identified_eur':0,'remaining_eur':0,'source':source} for v in years],
                'source':source,'reporting_date':quarter_end(period).isoformat(),'confirmed_zero_stock':True,
                'confirmation_reason':override['reason']}
            entities.append(existing)
            issue('CONFIRMED_ZERO_STOCK','Closing stock is confirmed zero; monthly flows and efficiency remain unavailable.',eid)
        if 'cst' in override:
            if existing is None:issue('CST_WITHOUT_COUNTRY','Supply country data before a CST balance.',eid,True);continue
            cst=override['cst']
            if not numeric(cst.get('amount')) or cst['amount']<0 or cst.get('currency',existing['currency'])!=existing['currency']:
                issue('INVALID_CST','CST must be a nonnegative balance in the country currency.',eid,True)
            else:existing['cst']={**cst,'source':'Manual confirmation: '+override['reason']}
    received={e['id'] for e in entities};missing=sorted(set(configured)-received)
    if not entities:issue('EMPTY','No country data were extracted.',error=True)
    if missing:issue('MISSING_COUNTRY','Missing countries: '+', '.join(missing)+'.',error=not allow_partial)
    for eid in sorted(received-set(configured)):issue('UNKNOWN_COUNTRY','Country not in enabled configuration.',eid,True)
    rates={};fx=data.setdefault('fx',[])
    if not any(r.get('currency')=='EUR' for r in fx):fx.append({'currency':'EUR','rate':1,'date':quarter_end(period).isoformat(),'source':'EUR identity'})
    for r in fx:
        currency=r.get('currency')
        if currency in rates:issue('DUPLICATE_FX',f'Duplicate FX for {currency}.',error=True)
        rates[currency]=r.get('rate')
        if not numeric(r.get('rate')) or r['rate']<=0:issue('INVALID_FX',f'{currency}: exchange rate must be finite and positive.',error=True)
        if r.get('date') and str(r['date'])[:7]!=quarter_end(period).isoformat()[:7]:issue('FX_PERIOD',f'{currency}: rate date does not belong to the quarter-end month.',error=True)
    if rates.get('EUR')!=1:issue('EUR_FX','EUR rate must equal 1.',error=True)
    def finite_tree(obj,eid):
        if isinstance(obj,dict):
            for v in obj.values():finite_tree(v,eid)
        elif isinstance(obj,list):
            for v in obj:finite_tree(v,eid)
        elif isinstance(obj,float) and not math.isfinite(obj):issue('NONFINITE','Input contains NaN or infinity.',eid,True)
    def reconcile(a,b,eid,label,tolerance=3.01):
        if numeric(a) and numeric(b) and abs(a-b)>.01:
            issue('ROUNDING' if abs(a-b)<=tolerance else 'RECONCILIATION',f'{label}: detail {a:,.2f}, reported {b:,.2f}. Reported values retained.',eid)
    for e in entities:
        eid=e['id'];finite_tree(e,eid)
        currency=e.get('currency')
        if currency not in rates:issue('MISSING_FX',f'Missing current-quarter closing rate for {currency}.',eid,True)
        if eid in configured and currency!=configured[eid]['currency']:issue('CURRENCY','Extracted currency disagrees with configuration.',eid,True)
        if e.get('reporting_date') and str(e['reporting_date'])[:10]!=quarter_end(period).isoformat():issue('COUNTRY_PERIOD','Country reporting date is not the requested quarter end.',eid,True)
        flows=e.get('flows',[]);months=[r.get('month') for r in flows]
        if flows and (len(months)!=len(set(months)) or sorted(months)!=rolling_months(period)):issue('FLOW_MONTHS','Expected one row per month in the six-month rolling window.',eid,True)
        if not flows and e.get('profile','standard')=='standard':issue('MISSING_FLOWS','Standard country has no monthly KPI_2 data.',eid,True)
        for row in flows:
            for key in ['amount','count']:
                if not numeric(row.get(key)):issue('MISSING_FLOW',f"{row.get('month')}: {key} is missing or nonnumeric.",eid,True)
            if e.get('profile','standard') not in ('brazil','runoff','stock_only'):
                for key in ['processed_amount','processed_count']:
                    if not numeric(row.get(key)):issue('MISSING_PROCESSED',f"{row.get('month')}: {key} is missing.",eid,True)
        for row in e.get('stock',[])+flows:
            processed=row.get('processed_source',row.get('processed_amount'))
            vals=[row.get(k) for k in e.get('processed_fields',[])]
            if vals and all(numeric(v) for v in vals):reconcile(sum(vals),processed,eid,'Processed components')
            total=row.get('total',row.get('amount'));pending=row.get('remaining_source',row.get('remaining_amount_source'))
            if numeric(total) and numeric(processed):reconcile(total-processed,pending,eid,'Gross less processed')
            for key in ('count','processed_count','remaining_count_source'):
                value=row.get(key)
                if value is not None and (not numeric(value) or value<0 or int(value)!=value):issue('INVALID_COUNT',f'{key} must be a nonnegative whole number.',eid,True)
        for kind,key in [('stock_bins','band'),('stock_vintages','vintage')]:
            rows=e.get(kind,[]);labels=[r.get(key) for r in rows]
            if len(labels)!=len(set(labels)):issue('DUPLICATE_BUCKET',f'Duplicate {kind} label.',eid,True)
            totals=e.get(kind+'_totals',{})
            for field in ('count','amount'):
                vals=[r.get(field) for r in rows]
                if vals and all(numeric(v) for v in vals):reconcile(sum(vals),totals.get(field),eid,f'{kind} {field}',0 if field=='count' else max(3.01,len(rows)*.51))
            for row in rows:
                if row.get('count') is not None and (not numeric(row['count']) or row['count']<0 or int(row['count'])!=row['count']):issue('INVALID_COUNT','Stock count must be a nonnegative whole number.',eid,True)
        for field in ('count','amount'):reconcile(e.get('stock_bins_totals',{}).get(field),e.get('stock_vintages_totals',{}).get(field),eid,'Amount-band total versus vintage '+field,0 if field=='count' else 3.01)
        totals=e.get('stock_totals',{})
        if e.get('stock'):
            if not numeric(totals.get('remaining_source')):issue('MISSING_STOCK','Reported pending stock total is required.',eid,True)
            for field in e.get('identified_addback_fields',[]):
                if not numeric(totals.get(field)):issue('MISSING_ADDBACK',f'Identified-stock category {field} is missing.',eid,True)
        elif not numeric(e.get('historical_metrics',{}).get('identified')):issue('MISSING_STOCK','Closing stock is missing.',eid,True)
    # Freeze the same normalized snapshot the history step will verify.
    data['issues']=[{'severity':'error' if i in errors else 'warning',**i} for i in errors+warnings]
    data['coverage']={'expected':sorted(configured),'received':sorted(received),'missing':missing}
    try:digest=snapshot_hash(data)
    except (ValueError,TypeError):digest=''
    report={'valid':not errors,'period':period,'errors':errors,'warnings':warnings,'coverage':data['coverage'],'snapshot_sha256':digest}
    return data,report
