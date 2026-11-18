"""Aggregate contract-level Base data using reviewed, declarative country rules.

No Excel formula evaluation: schema, statuses, dates, scope and band boundaries
are explicit configuration. Contract identifiers are used for duplicate checks
only and never enter the extracted reporting snapshot or diagnostic messages.
"""
from copy import deepcopy
from datetime import date, datetime
from decimal import Decimal

from .common import quarter_end, rolling_months


def _sum(rows):
    return float(sum((r['amount'] for r in rows), Decimal(0)))


def _read_rows(reader, profile):
    from .extract import normal
    sheet = reader.sheet(profile.get('sheet', 'Base'))
    key=(sheet,profile['header_row'])
    cache=getattr(reader,'_base_grids',{})
    if key not in cache:
        start=profile['header_row']
        values=list(reader.values[sheet].iter_rows(min_row=start,values_only=True))
        formulas=list(reader.formulas[sheet].iter_rows(min_row=start,values_only=True))
        cache[key]=(values,formulas);reader._base_grids=cache
    values,formulas=cache[key]
    if not values:raise ValueError(f'{reader.path.name}!{sheet}: empty Base sheet')
    headers=[normal(v) for v in values[0]];columns={}
    for field,aliases in profile['columns'].items():
        matches=[i for i,h in enumerate(headers) if h and h in {normal(a) for a in aliases}]
        if len(matches)!=1:raise ValueError(f'{reader.path.name}!{sheet}: Base field {field!r} requires one matching header, found {len(matches)}. Check config/base.json.')
        columns[field]=matches[0]
    return sheet,values[1:],formulas[1:],columns


def _band_thresholds(reader, profile, cfg, end):
    from .extract import as_date, normal, number
    bands=profile['amount_bands'];factor=Decimal(1)
    if bands.get('basis','local') not in ('local','eur_equivalent'):
        raise ValueError('Base amount_bands.basis must be local or eur_equivalent')
    if bands.get('basis','local')=='eur_equivalent':
        source=bands['rate_source'];sheet=reader.sheet(source['sheet'])
        currency_column=source['currency_columns'][cfg['currency']]
        matches=[]
        for row_number,row in enumerate(reader.values[sheet].iter_rows(min_row=source['first_row'],max_row=source['last_row'],values_only=True),source['first_row']):
            raw=row[source['date_column']-1]
            try:when=as_date(raw)
            except ValueError:continue
            if when==end:
                formula=reader.formulas[sheet].cell(row_number,currency_column).value
                if isinstance(formula,str) and formula.startswith('='):
                    raise ValueError('Base band FX must be a raw Parameters rate, not an unevaluated formula')
                matches.append(row[currency_column-1])
        if len(matches)!=1:raise ValueError(f'{reader.path.name}: Base band FX requires one Parameters rate for {end} / {cfg["currency"]}')
        value=number(matches[0],'Base band FX')
        if value is None or value<=0:raise ValueError('Base band FX must be positive')
        factor=Decimal(1)/Decimal(str(value))
    bounds=[Decimal(str(x))*factor for x in bands['upper_bounds']]
    if not bounds or bounds[0]<=0 or any(b<=a for a,b in zip(bounds,bounds[1:])):
        raise ValueError('Base amount-band upper_bounds must be positive and strictly increasing')
    return bounds


def extract_base(reader,cfg,profile,period,issues):
    from .extract import _empty_entity, issue, normal, as_date, number, CATEGORIES
    end=quarter_end(period);year=end.year
    meta=profile['reporting_date']
    reported=cfg.get('reporting_date_override') or reader.address(reader.sheet(meta['sheet'].format(prefix=cfg.get('sheet_prefix',''))),meta['cell'])
    if as_date(reported,'Base snapshot reporting date')!=end:
        raise ValueError(f'{reader.path.name}: Base snapshot date does not match {period}')
    category_map={}
    for category,statuses in profile['status_categories'].items():
        if category not in CATEGORIES:raise ValueError(f'Unknown Base category {category}')
        for status in statuses:
            key=normal(status)
            if not key or key in category_map:raise ValueError('Base status mappings must be nonempty and disjoint')
            category_map[key]=category
    pending={normal(v) for v in profile.get('pending_statuses',[''])}
    if pending.intersection(category_map):raise ValueError('Base pending and processed statuses overlap')
    if not set(profile.get('identified_addback_fields',cfg.get('identified_addback_fields',[]))).issubset(profile['status_categories']):
        raise ValueError('Base addback fields must be mapped processing categories')
    if not set(profile.get('kpi3_categories',[None])).issubset(set(profile['status_categories'])|{None}):
        raise ValueError('Base KPI3 categories must be pending/null or mapped processing categories')
    if profile.get('duplicate_key_policy','error') not in ('error','warn'):
        raise ValueError('Unknown Base duplicate_key_policy')
    for key,allowed in [('missing_decision_date',('error','snapshot_status')),
                        ('invalid_decision_date',('error','snapshot_status')),
                        ('decision_before_origin',('error','warn'))]:
        if profile.get(key,'error') not in allowed:raise ValueError(f'Unknown Base {key} policy')
    prior_years=profile.get('stock_prior_years',1)
    if not isinstance(prior_years,int) or not 1<=prior_years<=20:
        raise ValueError('Base stock_prior_years must be an integer from 1 through 20')
    sheet,values,formulas,cols=_read_rows(reader,profile)
    expected_countries={normal(c) for c in cfg.get('country_aliases',[cfg['name']])}
    units=Decimal(str(cfg.get('units',1)))
    if units<=0:raise ValueError('Country units must be positive')
    records=[];seen=set();exact_seen=set();future_origins=0;future_decisions=0;missing_decisions=0;selected=0;early_decisions=0;duplicate_keys=0;invalid_decisions=0
    excluded=0
    for offset,(row,formula_row) in enumerate(zip(values,formulas),profile['header_row']+1):
        get=lambda field:row[cols[field]] if cols[field]<len(row) else None
        if all(get(k) in (None,'') for k in ('contract','amount','origin','status')):continue
        if 'country' in cols:
            country=normal(get('country'))
            allowed={normal(v) for v in profile.get('allowed_countries',[])}
            if allowed and country not in allowed:
                raise ValueError(f'{reader.path.name}!{sheet}, row {offset}: unassigned Base country; review allowed_countries and input labels')
            if country not in expected_countries:continue
        selected+=1
        for name,index in cols.items():
            if index<len(formula_row) and isinstance(formula_row[index],str) and formula_row[index].startswith('='):
                raise ValueError(f'{reader.path.name}!{sheet}, row {offset}: Base field {name} contains a formula; supply raw values or an explicitly supported adapter.')
        identity=get('contract')
        if identity in (None,''):raise ValueError(f'{reader.path.name}!{sheet}, row {offset}: missing contract key')
        origin=as_date(get('origin'),f'Base origin date at row {offset}')
        key=(str(identity).strip(),origin)
        exact=(str(identity).strip(),origin,str(get('amount')),str(get('decision_date')),normal(get('status')))
        if exact in exact_seen:raise ValueError(f'{reader.path.name}!{sheet}: duplicate Base record at row {offset}; records were not deduplicated.')
        exact_seen.add(exact)
        if key in seen:
            if profile.get('duplicate_key_policy','error')=='error':
                raise ValueError(f'{reader.path.name}!{sheet}: duplicate (country, contract, origin-date) key at row {offset}; records were not deduplicated.')
            duplicate_keys+=1
        seen.add(key)
        if origin>end:future_origins+=1;continue
        amount=number(get('amount'),f'Base amount at row {offset}')
        if amount is None:raise ValueError(f'{reader.path.name}!{sheet}, row {offset}: missing amount')
        if amount<0:raise ValueError(f'{reader.path.name}!{sheet}, row {offset}: negative positive-balance amount; review input')
        status=normal(get('status'))
        if status not in pending and status not in category_map:
            raise ValueError(f'{reader.path.name}!{sheet}, row {offset}: unmapped decision status; update status_categories in config/base.json.')
        category=category_map.get(status)
        decision=get('decision_date')
        if decision not in (None,''):
            try:
                raw_decision=decision
                try:decision=as_date(raw_decision,f'Base decision date at row {offset}')
                except ValueError:
                    decision=None
                    for fmt in profile.get('date_formats',[]):
                        try:decision=datetime.strptime(str(raw_decision).strip(),fmt).date();break
                        except ValueError:pass
                    if decision is None:raise ValueError(f'Base decision date at row {offset} is invalid')
            except ValueError:
                if profile.get('invalid_decision_date','error')!='snapshot_status':raise
                invalid_decisions+=1;decision=None
        if decision is not None and decision!='':
            if decision<origin:
                if profile.get('decision_before_origin','error')!='warn':
                    raise ValueError(f'{reader.path.name}!{sheet}, row {offset}: decision precedes positive balance')
                early_decisions+=1
            if decision>end:category=None;future_decisions+=1
        elif category is not None and get('decision_date') in (None,''):
            if profile.get('missing_decision_date','error')=='error':
                raise ValueError(f'{reader.path.name}!{sheet}, row {offset}: processed status has no decision date')
            if profile.get('missing_decision_date')!='snapshot_status':raise ValueError('Unknown missing_decision_date policy')
            missing_decisions+=1
        include_stock=True
        for rule in profile.get('kpi3_exclusions',[]):
            field=rule['field']
            if field not in cols:raise ValueError(f'Base stock exclusion field {field} is not mapped')
            if normal(get(field)) in {normal(v) for v in rule['values']}:
                if not rule.get('categories') or category in rule['categories']:include_stock=False
        if not include_stock:excluded+=1
        records.append(dict(amount=Decimal(str(amount))*units,origin=origin,category=category,stock_eligible=include_stock))
    if not selected:raise ValueError(f'{reader.path.name}!{sheet}: no rows for {cfg["name"]}; cannot infer confirmed zero stock')
    e=_empty_entity(cfg,reader.path)
    e['processed_fields']=list(profile['status_categories'])
    e['identified_addback_fields']=profile.get('identified_addback_fields',cfg.get('identified_addback_fields',[]))
    source=f'{reader.path.name}!{sheet} (aggregated Base)'
    def stock(rows):
        result={k:None for k in CATEGORIES}
        for category in e['processed_fields']:result[category]=_sum(r for r in rows if r['category']==category)
        result.update(total=_sum(rows),processed_source=_sum(r for r in rows if r['category'] is not None),
            remaining_source=_sum(r for r in rows if r['category'] is None),source=source)
        return result
    cutoff=year-profile.get('stock_prior_years',1)
    labels=[f'<{cutoff}']+[str(y) for y in range(cutoff,year+1)]
    for label in labels:
        selected_rows=[r for r in records if (r['origin'].year<cutoff if label.startswith('<') else r['origin'].year==int(label))]
        e['stock'].append(dict(stock(selected_rows),vintage=label))
    e['stock_totals']=stock(records)
    for month in rolling_months(period):
        rows=[r for r in records if r['origin'].strftime('%Y-%m')==month]
        aggregate=stock(rows)
        processed=[r for r in rows if r['category'] is not None]
        e['flows'].append(dict(month=month,amount=aggregate['total'],count=len(rows),
            processed_amount=aggregate['processed_source'],processed_count=len(processed),
            remaining_amount_source=aggregate['remaining_source'],remaining_count_source=len(rows)-len(processed),
            source=source,**{k:aggregate[k] for k in CATEGORIES}))
    stock_categories=profile.get('kpi3_categories',[None])
    retained=[r for r in records if r['category'] in stock_categories and r['stock_eligible']]
    bounds=_band_thresholds(reader,profile,cfg,end)
    def label_amount(x):return format(float(x),'.1f').rstrip('0').rstrip('.') if x<10 else f'{float(x):.0f}'
    for i,upper in enumerate(bounds+[None]):
        lower=Decimal(0) if i==0 else bounds[i-1]
        selected_rows=[r for r in retained if (r['amount']>=0 if i==0 else r['amount']>lower) and (upper is None or r['amount']<=upper)]
        if profile['amount_bands'].get('basis')=='eur_equivalent':
            display=profile['amount_bands']['upper_bounds']
            low=0 if i==0 else display[i-1]
            label=(f'> {low}' if upper is None else f'{low} - {display[i]}')+' EUR eq.'
        else:
            label=(f'> {label_amount(lower)}' if upper is None else f'{label_amount(lower)} - {label_amount(upper)}')+' '+cfg['currency']
        e['stock_bins'].append(dict(band=label,count=len(selected_rows),amount=_sum(selected_rows),source=source))
    old=profile.get('vintage_before',2018)
    for label in [f'<{old}']+[str(y) for y in range(old,year+1)]:
        rows=[r for r in retained if (r['origin'].year<old if label.startswith('<') else r['origin'].year==int(label))]
        e['stock_vintages'].append(dict(vintage=label,count=len(rows),amount=_sum(rows),source=source))
    total=dict(label='Total',count=len(retained),amount=_sum(retained),source=source)
    e['stock_bins_totals']=deepcopy(total);e['stock_vintages_totals']=deepcopy(total)
    e['kpi1']={'amounts':[],'counts':[]}
    for kind in ('amounts','counts'):
        measure=_sum if kind=='amounts' else len
        for label in labels+['Total']:
            rows=records if label=='Total' else [r for r in records if (r['origin'].year<cutoff if label.startswith('<') else r['origin'].year==int(label))]
            months={f'{year}-{m:02}':measure([r for r in rows if r['origin'].year==year and r['origin'].month==m]) for m in range(1,13)}
            e['kpi1'][kind].append(dict(vintage=label,total=measure(rows),months=months,source=source))
    e['extraction_method']='base'
    e['stock_band_basis']=profile['amount_bands'].get('basis','local')
    e['base_audit']={'profile':profile['id'],'selected_rows':selected,'included_as_of_rows':len(records),
        'future_origins_excluded':future_origins,'future_decisions_pending':future_decisions,
        'missing_decision_dates':missing_decisions,'invalid_decision_dates':invalid_decisions,'decisions_before_origin':early_decisions,'stock_rows_excluded':excluded,
        'repeated_contract_origin_keys':duplicate_keys,'band_upper_bounds':profile['amount_bands']['upper_bounds'],
        'count_unit':'Base rows (COUNTIFS convention), not unique contracts','band_convention':'[0, first upper], then (lower, upper]'}
    issue(issues,e['id'],'BASE_EXTRACTED',f'Derived KPI data from {len(records)} Base records using {profile["id"]}; no KPI formula caches required.',reader.path.name)
    for value,code,message in [(future_origins,'BASE_FUTURE_ORIGIN','Future-origin rows excluded from the as-of report'),
        (future_decisions,'BASE_FUTURE_DECISION','Future decisions treated as pending at quarter-end'),
        (missing_decisions,'BASE_MISSING_DECISION_DATE','Configured snapshot-status policy used where decision date is absent'),
        (invalid_decisions,'BASE_INVALID_DECISION_DATE','Configured snapshot-status policy used for an invalid decision date; no replacement date inferred'),
        (duplicate_keys,'BASE_REPEATED_KEY','Repeated contract/origin keys retained as separate nonidentical rows per configured COUNTIFS convention'),
        (early_decisions,'BASE_DECISION_BEFORE_ORIGIN','Decision precedes origin; preserved snapshot category under explicit warn policy')]:
        if value:issue(issues,e['id'],code,f'{message}: {value}.',reader.path.name)
    if e['requires_cst']:issue(issues,e['id'],'CST_REQUIRED','Base CSR movements do not supply the separate CST account closing balance.',reader.path.name)
    return e


def reconcile_with_kpi(reader,cfg,layout,period,entity,issues,tolerance=.01):
    """Optional independent cross-check; unusable caches never feed Base totals."""
    from .extract import _standard, issue
    try:kpi=_standard(reader,cfg,layout,period,[])
    except (ValueError,KeyError) as exc:
        issue(issues,cfg['id'],'BASE_KPI_COMPARISON_UNAVAILABLE',f'Base extracted; optional cached KPI comparison unavailable: {exc}',reader.path.name)
        return
    checked=0;gaps=[]
    def check(path,a,b):
        nonlocal checked
        if isinstance(a,(int,float)) and isinstance(b,(int,float)):
            checked+=1
            if abs(a-b)>tolerance:gaps.append({'metric':path,'base':a,'cached_kpi':b,'difference':a-b})
    for field in ('total','processed_source','remaining_source'):
        check('stock_totals.'+field,entity['stock_totals'][field],kpi['stock_totals'][field])
    prior={r['month']:r for r in kpi['flows']}
    for row in entity['flows']:
        for field in ('amount','count','processed_amount','processed_count'):
            check(f'flows.{row["month"]}.{field}',row[field],prior.get(row['month'],{}).get(field))
    for kind in ('stock_bins','stock_vintages'):
        for field in ('amount','count'):check(kind+'.total.'+field,entity[kind+'_totals'][field],kpi[kind+'_totals'][field])
        if kind=='stock_vintages':
            left={r['vintage'].replace(' ',''):r for r in entity[kind]}
            right={r['vintage'].replace(' ',''):r for r in kpi[kind]}
            if left.keys()!=right.keys():issue(issues,cfg['id'],'BASE_KPI_LAYOUT_DIFFERENCE','KPI3 vintage labels differ; only common labels compared.',reader.path.name)
            pairs=[(label,left[label],right[label]) for label in sorted(left.keys() & right.keys())]
        elif len(entity[kind])==len(kpi[kind]) and entity['base_audit']['band_upper_bounds']==[1,10,100,1000]:
            # These are the five reviewed template bands. Nordic labels are
            # rounded local equivalents of the configured EUR thresholds.
            pairs=[(str(i),a,b) for i,(a,b) in enumerate(zip(entity[kind],kpi[kind]))]
        else:
            issue(issues,cfg['id'],'BASE_KPI_LAYOUT_DIFFERENCE','KPI3 amount-band layout changed; compared totals only, not unmatched bands.',reader.path.name)
            pairs=[]
        for i,a,b in pairs:
            for field in ('amount','count'):check(f'{kind}.{i}.{field}',a[field],b[field])
    entity['base_reconciliation']={'compared_metrics':checked,'tolerance':tolerance,'differences':gaps}
    if gaps:issue(issues,cfg['id'],'BASE_KPI_DIFFERENCES',f'{len(gaps)} Base-derived metrics differ from cached KPI results. Base is authoritative; review base_reconciliation in extracted data.',reader.path.name)


def select_base(reader,cfg,layout,period,issues,configuration):
    """Return a Base entity, or None only for an explicitly allowed KPI route."""
    from .extract import normal,issue
    settings=configuration.get('countries',{}).get(cfg['id'])
    if not settings or settings.get('mode')=='kpi':return None
    mode=settings.get('mode','base_if_present')
    if mode not in ('base','base_if_present'):raise ValueError(f'Unsupported Base extraction mode {mode}')
    profile_id=settings.get('profile')
    profile=deepcopy(configuration.get('profiles',{}).get(profile_id,{}))
    profile.update(settings.get('overrides',{}))
    if 'sheet' in settings:profile['sheet']=settings['sheet']
    present=any(normal(s)==normal(profile.get('sheet','Base')) for s in reader.values.sheetnames)
    if not present:
        if mode=='base':raise ValueError(f'{reader.path.name}: configured Base sheet missing')
        issue(issues,cfg['id'],'BASE_ABSENT_KPI_FALLBACK','Base sheet absent; using configured KPI worksheet fallback.',reader.path.name)
        return None
    if not profile_id or profile_id not in configuration.get('profiles',{}):
        raise ValueError(f'{reader.path.name}: Base exists for {cfg["name"]}, but its mapping is unverified. Configure a reviewed profile in config/base.json or explicitly select mode="kpi".')
    profile['id']=profile_id
    entity=extract_base(reader,cfg,profile,period,issues)
    if settings.get('compare_kpi',True):reconcile_with_kpi(reader,cfg,layout,period,entity,issues,configuration.get('reconciliation_tolerance',.01))
    return entity
