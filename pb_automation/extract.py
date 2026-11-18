"""Step 1: read country workbooks into plain, auditable Python dictionaries.

No history writes happen here. Reviewed Base profiles aggregate raw records;
other countries use KPI worksheets with saved formula caches. A present but
unmapped Base sheet fails explicitly instead of silently using stale KPIs.
"""
from __future__ import annotations

import calendar
import fnmatch
import hashlib
import math
import re
import unicodedata
from datetime import date, datetime
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.utils.datetime import from_excel

from .common import parse_period, quarter_end, rolling_months

CATEGORIES = ("refund", "unreachable", "false_positive", "authority_refund", "other_without_refund")


def normal(value):
    """Compare labels without accents, punctuation or accidental whitespace."""
    value = unicodedata.normalize("NFKD", str(value or "")).encode("ascii", "ignore").decode()
    return re.sub(r"[^a-z0-9]+", " ", value.casefold()).strip()


def number(value, where="value"):
    if value is None or (isinstance(value, str) and value.strip() in ("", "-", "–", "—")):
        return None
    if isinstance(value, bool):
        raise ValueError(f"{where}: boolean is not a numeric amount")
    if isinstance(value, (int, float)):
        result = value
    else:
        text = re.sub(r"[\s\u00a0\u202f]", "", str(value))
        negative = text.startswith("(") and text.endswith(")")
        if negative: text = text[1:-1]
        if re.fullmatch(r"[+-]?\d{1,3}(,\d{3})+(\.\d+)?", text):
            text = text.replace(",", "")
        elif "," in text and "." not in text:
            text = text.replace(",", ".")
        try:
            result = float(text) * (-1 if negative else 1)
        except ValueError as exc:
            raise ValueError(f"{where}: expected a number, got {value!r}") from exc
    if not math.isfinite(result):
        raise ValueError(f"{where}: non-finite number")
    return result


def as_date(value, where="date"):
    if isinstance(value, datetime): return value.date()
    if isinstance(value, date): return value
    if isinstance(value, (int, float)):
        try: return from_excel(value).date()
        except (ValueError, OverflowError, AttributeError): pass
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%Y-%m-%dT%H:%M:%S", "%d-%m-%Y"):
        try: return datetime.strptime(str(value).strip(), fmt).date()
        except ValueError: pass
    raise ValueError(f"{where}: invalid or missing date {value!r}")


def month_key(value):
    if isinstance(value, (datetime, date)): return value.strftime("%Y-%m")
    text = normal(value)
    iso = re.fullmatch(r"(20\d{2}) (\d{1,2})(?: \d{1,2})?", text)
    if iso:
        year, month = map(int, iso.groups())
        if 1 <= month <= 12: return f"{year:04}-{month:02}"
    match = re.fullmatch(r"([a-z]+) (\d{2}|20\d{2})", text)
    if not match: raise ValueError(f"Invalid monthly header: {value!r}")
    word, yr = match.groups(); year = int(yr) + (2000 if len(yr) == 2 else 0)
    aliases = (("jan",), ("feb", "fev"), ("mar",), ("apr", "avr"), ("may", "mai"),
               ("jun", "juin"), ("jul", "juil"), ("aug", "aou"), ("sep",), ("oct",), ("nov",), ("dec",))
    for i, names in enumerate(aliases, 1):
        if any(word.startswith(prefix) for prefix in names): return f"{year}-{i:02}"
    raise ValueError(f"Unknown month: {value!r}")


class ExcelReader:
    """Read summary cells lazily; the Base adapter separately streams raw grids."""
    def __init__(self, path):
        self.path = Path(path)
        self.values = openpyxl.load_workbook(path, data_only=True, read_only=True)
        self.formulas = openpyxl.load_workbook(path, data_only=False, read_only=True)
        self._grids = {}

    def close(self):
        self.values.close(); self.formulas.close()

    def sheet(self, requested):
        exact = [s for s in self.values.sheetnames if normal(s).replace(" ", "") == normal(requested).replace(" ", "")]
        if len(exact) != 1: raise ValueError(f"{self.path.name}: expected one worksheet {requested!r}; found {exact}")
        return exact[0]

    def prepare(self, sheet, max_rows=200, max_cols=30):
        if sheet in self._grids: return
        vals = list(self.values[sheet].iter_rows(max_row=max_rows, max_col=max_cols, values_only=True))
        formulas = list(self.formulas[sheet].iter_rows(max_row=max_rows, max_col=max_cols, values_only=True))
        self._grids[sheet] = (vals, formulas)

    def cell(self, sheet, row, col):
        self.prepare(sheet)
        vals, formulas = self._grids[sheet]
        if row > len(vals) or col > len(vals[row-1]): return None
        value, formula = vals[row-1][col-1], formulas[row-1][col-1]
        where = self.ref(sheet, row, col)
        if isinstance(formula, str) and formula.startswith("=") and value is None:
            raise ValueError(f"{where}: missing formula cache. Recalculate and save the INPUT file in Excel.")
        if isinstance(value, str) and value.startswith(("#REF!", "#DIV/0!", "#VALUE!", "#N/A", "#NAME?", "#NUM!")):
            raise ValueError(f"{where}: Excel error {value}")
        return value

    def address(self, sheet, address):
        from openpyxl.utils.cell import coordinate_to_tuple
        return self.cell(sheet, *coordinate_to_tuple(address))

    def ref(self, sheet, row, col): return f"{self.path.name}!{sheet}!{get_column_letter(col)}{row}"

    def numeric(self, sheet, row, col, units=1):
        value = number(self.cell(sheet, row, col), self.ref(sheet, row, col))
        return None if value is None else value * units

    def find(self, sheet, aliases, start=1, end=150, column=1):
        needles = [normal(x) for x in aliases]
        for row in range(start, end+1):
            value = normal(self.cell(sheet, row, column))
            if any(value.startswith(n) for n in needles): return row
        raise ValueError(f"{self.path.name}!{sheet}: missing label {aliases!r}")


def issue(issues, eid, code, message, source="", severity="warning"):
    issues.append(dict(severity=severity, entity=eid, code=code, message=message, source=source))


def _select(files, patterns, aliases, period, label):
    args = dict(period=period, year=period[:4], quarter=period[-1], folder=f"Q{period[-1]}_{period[:4]}")
    matches = {p for p in files for pat in patterns if fnmatch.fnmatchcase(p.name.casefold(), pat.format(**args).casefold())}
    if not matches and aliases:
        for p in files:
            stem = normal(p.stem)
            stem = re.sub(r"\b(?:20\d{2} q[1-4]|q[1-4] ?20\d{2}|20\d{2}q[1-4])\b", "", stem).strip()
            if stem in {normal(a) for a in aliases}: matches.add(p)
    if len(matches) > 1: raise ValueError(f"Duplicate inputs for {label}: {', '.join(sorted(p.name for p in matches))}")
    return next(iter(matches), None)


def _empty_entity(cfg, path):
    return dict(id=cfg['id'], name=cfg['name'], currency=cfg['currency'], profile=cfg['profile'],
        requires_cst=cfg.get('requires_cst', False), processed_fields=cfg.get('processed_fields', []),
        identified_addback_fields=cfg.get('identified_addback_fields', []),
        stock=[], flows=[], stock_bins=[], stock_vintages=[], historical_metrics={}, historical_vintages=[],
        stock_bins_totals={}, stock_vintages_totals={}, source=path.name,
        stock_basis='source_remaining', use_source_stock_totals=True, flow_basis='source_processed',
        stock_count_basis='stock_vintages', prefer_reported_distribution_totals=True)


def _metadata(reader, sheet, cfg, layout, period, issues, kpi1=False):
    date_value = reader.address(sheet, layout['date_cell'])
    effective = cfg.get('reporting_date_override') or date_value
    if as_date(effective, f"{reader.path.name}!{sheet} date") != quarter_end(period):
        raise ValueError(f"{reader.path.name}!{sheet}: reporting date {effective!r} does not match {period}")
    if cfg.get('reporting_date_override'):
        issue(issues,cfg['id'],'DATE_OVERRIDE','Reporting date explicitly overridden in country configuration.',reader.path.name)
    name = reader.address(sheet, layout['country_cell'])
    allowed = {normal(x) for x in cfg.get('country_aliases', [cfg['name']])}
    if normal(name) not in allowed:
        message = f"{sheet} country label is {name!r}; configured country is {cfg['name']}."
        if kpi1 or cfg.get('allow_country_mismatch'): issue(issues,cfg['id'],'COUNTRY_METADATA',message,reader.path.name)
        else: raise ValueError(f"{reader.path.name}: {message}")
    currency = str(reader.address(sheet, layout['currency_cell']) or '').strip().upper()
    allowed_currency = set(cfg.get('currency_aliases', [cfg['currency']]))
    if currency not in allowed_currency:
        message=f"{sheet} currency label {currency!r}; configured amounts are {cfg['currency']}."
        if cfg.get('allow_currency_mismatch') or kpi1: issue(issues,cfg['id'],'CURRENCY_METADATA',message,reader.path.name)
        else: raise ValueError(f"{reader.path.name}: {message}")


def _distribution(reader, sheet, layout, kind, units):
    cfg = layout['distribution']; prefix = cfg['amount_label'] if kind == 'stock_bins' else cfg['vintage_label']
    start = reader.find(sheet,prefix,column=cfg['section_column']); rows=[]; total={}
    for r in range(start, start+100):
        lab=reader.cell(sheet,r,cfg['label_column'])
        if lab is None or not str(lab).strip(): continue
        record = dict(count=reader.numeric(sheet,r,cfg['count_column']),
                      amount=reader.numeric(sheet,r,cfg['amount_column'],units),source=reader.ref(sheet,r,cfg['label_column']))
        if normal(lab).startswith('total'):
            total=dict(record,label='Total');break
        label = str(int(lab)) if isinstance(lab,(int,float)) and float(lab).is_integer() else str(lab).strip()
        record['band' if kind=='stock_bins' else 'vintage'] = label if kind=='stock_bins' else label.replace(' ','')
        rows.append(record)
    if not rows or not total: raise ValueError(f"{reader.path.name}!{sheet}: incomplete {kind} distribution")
    return rows,total


def _kpi1(reader, sheet, layout, period, cfg, issues):
    result={'amounts':[], 'counts':[]}
    headers=[]
    for r in range(1,80):
        if normal(reader.cell(sheet,r,1))=='vintage':headers.append(r)
    if len(headers)<2:
        issue(issues,cfg['id'],'KPI1_LAYOUT','KPI1 vintage sections were not found; KPI2/KPI3 remain available.',reader.path.name)
        return result
    for kind,head in zip(('amounts','counts'),headers[:2]):
        for r in range(head+1,head+12):
            label=reader.cell(sheet,r,1)
            if label is None:continue
            row={'vintage':str(int(label)) if isinstance(label,(int,float)) else str(label).strip(),
                 'total':reader.numeric(sheet,r,2,cfg.get('units',1) if kind=='amounts' else 1),
                 'months':{f"{period[:4]}-{m:02}":reader.numeric(sheet,r,m+2,cfg.get('units',1) if kind=='amounts' else 1) for m in range(1,13)},
                 'source':reader.ref(sheet,r,1)}
            result[kind].append(row)
            if normal(label)=='total':break
    endmonth=int(period[-1])*3
    for kind in result:
        for row in result[kind]:
            if any(v not in (None,0) and int(m[-2:])>endmonth for m,v in row['months'].items()):
                issue(issues,cfg['id'],'KPI1_OUTSIDE_QUARTER',f"KPI1 {kind} contain later months. Retained as supplied; flow reporting uses only KPI2 rolling-semester data.",reader.path.name)
                break
    return result


def _standard(reader, cfg, layout, period, issues):
    e=_empty_entity(cfg,reader.path);prefix=cfg.get('sheet_prefix','');units=cfg.get('units',1)
    sheets={k:reader.sheet(prefix+name) for k,name in layout['sheets'].items()}
    _metadata(reader,sheets['kpi2'],cfg,layout['metadata'],period,issues)
    _metadata(reader,sheets['kpi1'],cfg,layout['metadata'],period,issues,kpi1=True)
    k3date=reader.address(sheets['kpi3'],layout['distribution']['date_cell'])
    if as_date(k3date,f"{reader.path.name}!{sheets['kpi3']} date") != quarter_end(period):
        raise ValueError(f"{reader.path.name}!{sheets['kpi3']}: reporting date does not match {period}")
    s=sheets['kpi2']; labels=layout['labels']
    def find(key,start=1):
        if key in layout.get('row_overrides',{}):return layout['row_overrides'][key]
        return reader.find(s,labels[key],start=start)
    stock=find('stock_total');processed=find('stock_processed',stock);pending=find('stock_pending',stock)
    flow=find('flow_amount');flowcount=find('flow_count',flow);flowproc=find('flow_processed',flow)
    flown=find('flow_processed_count',flowproc);flowpending=find('flow_pending',flow);flowpn=find('flow_pending_count',flowpending)
    stock_categories={k:reader.find(s,layout['category_labels'][k],start=stock,end=pending-1) for k in e['processed_fields']}
    flow_categories={k:reader.find(s,layout['category_labels'][k],start=flowproc,end=flowpending-1) for k in e['processed_fields']}
    def stock_record(c):
        row={k:None for k in CATEGORIES}
        row.update(total=reader.numeric(s,stock,c,units),processed_source=reader.numeric(s,processed,c,units),
                   remaining_source=reader.numeric(s,pending,c,units),source=reader.ref(s,stock,c))
        for k,r in stock_categories.items():row[k]=reader.numeric(s,r,c,units)
        return row
    e['stock_totals']=stock_record(layout['total_column'])
    for c in layout['stock_columns']:
        label=reader.cell(s,stock-1,c)
        if label is None:raise ValueError(f"{reader.ref(s,stock-1,c)}: missing stock-vintage label")
        vintage=str(int(label)) if isinstance(label,(int,float)) else str(label).replace(' ','')
        e['stock'].append(dict(stock_record(c),vintage=vintage))
    for c in layout['month_columns']:
        row={k:None for k in CATEGORIES}
        row.update(month=month_key(reader.cell(s,flow-1,c)),amount=reader.numeric(s,flow,c,units),count=reader.numeric(s,flowcount,c),
            processed_amount=reader.numeric(s,flowproc,c,units),processed_count=reader.numeric(s,flown,c),
            remaining_amount_source=reader.numeric(s,flowpending,c,units),remaining_count_source=reader.numeric(s,flowpn,c),source=reader.ref(s,flow,c))
        for k,r in flow_categories.items():row[k]=reader.numeric(s,r,c,units)
        e['flows'].append(row)
    e['flows'].sort(key=lambda r:r['month'])
    if [r['month'] for r in e['flows']] != rolling_months(period):
        raise ValueError(f"{reader.path.name}!{s}: expected rolling months {rolling_months(period)}")
    for kind in ('stock_bins','stock_vintages'):
        e[kind],e[kind+'_totals']=_distribution(reader,sheets['kpi3'],layout,kind,units)
    e['kpi1']=_kpi1(reader,sheets['kpi1'],layout,period,cfg,issues)
    if e['requires_cst']:
        issue(issues,e['id'],'CST_REQUIRED','Separate CST account balance is not supplied by KPI2. CSR transfers are movements, not account balances.',reader.path.name)
    return e


def _romania(reader,cfg,layout,period,issues):
    sheet=reader.sheet(layout['sheet']);e=_empty_entity(cfg,reader.path);units=cfg.get('units',1)
    if as_date(reader.address(sheet,layout['date_cell'])) != quarter_end(period):
        raise ValueError(f"{reader.path.name}: Romania reporting date does not match {period}")
    for kind in ('stock_bins','stock_vintages'):
        e[kind],e[kind+'_totals']=_distribution(reader,sheet,layout,kind,units)
    e['historical_metrics']={'identified':number(reader.address(sheet,layout['total_amount_cell']))*units,
        'remaining':number(reader.address(sheet,layout['total_amount_cell']))*units,
        'contracts':number(reader.address(sheet,layout['total_count_cell']))}
    e['historical_vintages']=[dict(vintage=r['vintage'],identified_eur=r['amount'],remaining_eur=r['amount'],source=r['source']) for r in e['stock_vintages']]
    if cfg['currency']!='EUR':raise ValueError('Romania stock-only adapter expects the supplied EUR format; configure a new adapter for local-currency data.')
    issue(issues,e['id'],'STOCK_ONLY','Stock distributions supplied; monthly flow and treatment efficiency are unavailable.',reader.path.name)
    return e


def _brazil(path,cfg,layout,period,issues):
    try: from pyxlsb import open_workbook
    except ImportError as exc:raise ValueError('Brazil .xlsb input requires pyxlsb. Install requirements.txt.') from exc
    cells={}
    with open_workbook(str(path)) as wb:
        with wb.get_sheet(layout['sheet']) as sh:
            for row in sh.rows():
                for c in row:
                    if c.v is not None:cells[(c.r+1,c.c+1)]=c.v
    def num(r,c):return number(cells.get((r,c)),f'{path.name}!{layout["sheet"]}!{get_column_letter(c)}{r}')
    months={}
    for (r,c),value in cells.items():
        if r!=layout['month_header_row']:continue
        if isinstance(value,(int,float)) and 190001<=value<=210012 and int(value)==value:
            y,m=divmod(int(value),100)
            if 1<=m<=12:
                key=f'{y}-{m:02}'
                if key in months:raise ValueError(f'{path.name}: duplicate month {key}')
                months[key]=c
    required=rolling_months(period);last=required[-1]
    if any(m not in months for m in required):raise ValueError(f'{path.name}: Brazil report does not contain all six months {required}')
    end=months[last];e=_empty_entity(cfg,path);e['processed_fields']=[]
    e['historical_metrics']={'identified':num(layout['stock_amount_row'],end),'remaining':num(layout['stock_amount_row'],end),'contracts':num(layout['stock_count_row'],end)}
    e['calendar_movements']=[]
    for mo in required:
        col=months[mo];a=num(layout['in_amount_row'],col);n=num(layout['in_count_row'],col)
        out=num(layout['out_amount_row'],col);outn=num(layout['out_count_row'],col)
        e['calendar_movements'].append([mo,a,out,n,outn,num(layout['stock_amount_row'],col)])
        e['flows'].append(dict(month=mo,amount=a,count=n,processed_amount=None,processed_count=None,
            remaining_amount_source=None,remaining_count_source=None,source=f'{path.name}!{layout["sheet"]}!{get_column_letter(col)}',**{k:None for k in CATEGORIES}))
    origin=[]
    for r in range(layout['vintage_first_row'],layout['vintage_last_row']+1):
        yr=cells.get((r,layout['vintage_year_column']))
        if yr is None:continue
        year=number(yr)
        if year!=int(year) or not 1900<=year<=int(period[:4]):raise ValueError(f'{path.name}: invalid stock vintage {yr}')
        origin.append(dict(vintage=str(int(year)),count=num(r+layout['count_row_offset'],end),amount=num(r,end),source=f'{path.name}!{layout["sheet"]}!{get_column_letter(end)}{r}'))
    e['stock_origin']=origin
    cutoff=layout.get('old_vintage_before',2018);older=[r for r in origin if int(r['vintage'])<cutoff]
    def total(key):
        values=[r[key] for r in older]
        return None if any(v is None for v in values) else sum(values)
    e['stock_vintages']=([dict(vintage=f'<{cutoff}',count=total('count'),amount=total('amount'),source=path.name)] if older else [])+[r for r in origin if int(r['vintage'])>=cutoff]
    e['stock_vintages_totals']=dict(amount=num(layout['distribution_total_amount_row'],end),count=num(layout['distribution_total_count_row'],end),source=path.name,label='Total')
    # Keep LCY here; reporting converts stock_origin into consolidated vintages.
    issue(issues,e['id'],'BRAZIL_CALENDAR','IN/OUT are calendar movements, not treatment of the same monthly cohort. Cohort efficiency is unavailable.',path.name)
    return e


def _filename_period(path):
    text=path.stem.upper().replace('_','').replace('-','').replace(' ','')
    matches=re.findall(r'Q([1-4])(20\d{2})|(20\d{2})Q([1-4])',text)
    periods={f'{y or y2}Q{q or q2}' for q,y,y2,q2 in matches}
    return next(iter(periods)) if len(periods)==1 else None


def _fx(path,config,profiles,period,needed,issues):
    reader=ExcelReader(path)
    try:
        matches=[]
        for profile in profiles:
            for sheet in reader.values.sheetnames:
                if not any(fnmatch.fnmatchcase(sheet.strip().casefold(),x.casefold()) for x in profile['sheets']):continue
                sheet_date=None
                if profile.get('date_from_sheet_name'):
                    match=re.fullmatch(r'(20\d{2})\.(0[1-9]|1[0-2])',sheet.strip())
                    if not match:continue
                    year,month=map(int,match.groups())
                    sheet_date=date(year,month,calendar.monthrange(year,month)[1])
                    # A rate workbook may retain several monthly tabs. Select
                    # only the requested closing month, never the first tab.
                    if sheet_date!=quarter_end(period):continue
                reader.prepare(sheet,max_rows=600,max_cols=20)
                header=reader.address(sheet,profile['closing_header_cell'])
                if any(normal(header).startswith(normal(x)) for x in profile['closing_header_aliases']):matches.append((sheet,profile,sheet_date))
        if len(matches)!=1:raise ValueError(f'{path.name}: expected one identifiable TC / closing-rate table for {period}, found {len(matches)}. Available sheets: {reader.values.sheetnames}. Check config/layouts.json.')
        sheet,p,sheet_date=matches[0]
        raw=reader.address(sheet,p['date_cell']) if p.get('date_cell') else None
        override=config.get('date_override');evidence='embedded date'
        if override is not None:
            reported=as_date(override);evidence='explicit configured date_override'
        elif raw is not None:
            reported=as_date(raw,f'{path.name} FX date')
        elif sheet_date is not None:
            reported=sheet_date;evidence='dated worksheet name'
        elif config.get('allow_period_in_filename',False) and _filename_period(path):
            fileperiod=_filename_period(path)
            if fileperiod!=period:raise ValueError(f'{path.name}: FX filename period {fileperiod} does not match {period}')
            reported=quarter_end(period);evidence='period in filename'
            issue(issues,'FX','FX_DATE_FROM_FILENAME','No embedded FX date. Quarter is taken from the explicit period in the filename.',path.name)
        else:raise ValueError(f'{path.name}: no FX date. Set fx.date_override explicitly after checking the rate table, or use a filename containing its quarter.')
        end=quarter_end(period)
        if (reported.year,reported.month)!=(end.year,end.month):raise ValueError(f'{path.name}: FX month {reported:%Y-%m} does not match {period}')
        rates={}
        for r in range(p['first_data_row'],p.get('last_data_row',500)+1):
            currency=str(reader.cell(sheet,r,p['currency_column']) or '').strip().upper()
            if currency not in needed:continue
            if currency in rates:raise ValueError(f'{path.name}: duplicate FX currency {currency}')
            value=reader.numeric(sheet,r,p['closing_rate_column'])
            if value is None or value<=0:raise ValueError(f'{path.name}: missing or non-positive TC rate for {currency}')
            rates[currency]=dict(currency=currency,rate=value,date=end.isoformat(),source=reader.ref(sheet,r,p['closing_rate_column']),date_evidence=evidence,reported_date=reported.isoformat())
        if 'EUR' in rates and abs(rates['EUR']['rate']-1)>1e-9:raise ValueError('EUR closing rate must be 1')
        rates.setdefault('EUR',dict(currency='EUR',rate=1,date=end.isoformat(),source='EUR identity'))
        missing=needed-set(rates)
        if missing:raise ValueError(f'{path.name}: missing closing rates for {sorted(missing)}')
        return [rates[k] for k in sorted(rates)]
    finally:reader.close()


def extract_folder(input_dir: Path, period: str, countries: dict, layouts: dict, base_config: dict | None = None) -> dict:
    """Read one quarterly input folder. Missing countries are left for validation.

    Do not silently use the previous quarter for missing current-country files.
    No files are modified. File paths in returned provenance are portable names.
    """
    period=parse_period(period);input_dir=Path(input_dir)
    if not input_dir.is_dir():raise ValueError(f'Input folder not found: {input_dir}')
    files=sorted(p for p in input_dir.iterdir() if p.is_file() and p.suffix.lower() in ('.xlsx','.xlsm','.xlsb') and not p.name.startswith('~$'))
    issues=[];entities=[];used=set();readers={}
    configured=[c for c in countries['countries'] if c.get('enabled',True)]
    ids=[c['id'] for c in configured]
    if len(set(ids))!=len(ids):raise ValueError('Duplicate country IDs in configuration')
    try:
        for cfg in configured:
            path=_select(files,cfg.get('file_patterns',[]),cfg.get('aliases',[cfg['name']]),period,cfg['name'])
            if path is None:
                issue(issues,cfg['id'],'MISSING_COUNTRY','No current-quarter input file supplied. No historical values were substituted.');continue
            used.add(path);profile=cfg['profile']
            if profile=='brazil':e=_brazil(path,cfg,layouts[profile],period,issues)
            else:
                if path.suffix.lower()=='.xlsb':raise ValueError(f'{path.name}: .xlsb is supported only for the Brazil profile')
                if path not in readers:readers[path]=ExcelReader(path)
                reader=readers[path]
                if profile=='romania':e=_romania(reader,cfg,layouts[profile],period,issues)
                else:
                    from .base import select_base
                    e=select_base(reader,cfg,layouts[profile],period,issues,base_config or {})
                    if e is None:
                        e=_standard(reader,cfg,layouts[profile],period,issues)
                        e['extraction_method']='kpi'
            entities.append(e)
        fxconfig=countries.get('fx',{})
        fxpath=_select(files,fxconfig.get('file_patterns',['Taux.xlsx']),['Taux'],period,'exchange rates')
        if fxpath is None:raise ValueError('Missing exchange-rate workbook (Taux.xlsx)')
        used.add(fxpath)
        rates=_fx(fxpath,fxconfig,layouts['fx_profiles'],period,{e['currency'] for e in entities}|{'EUR'},issues)
        for path in files:
            if path not in used:issue(issues,'INPUT','UNMATCHED_FILE','Workbook not matched to a configured country or FX input; ignored.',path.name)
        return dict(period=period,date=quarter_end(period).isoformat(),entities=entities,fx=rates,issues=issues,
                    sources=[dict(file=p.name,sha256=hashlib.sha256(p.read_bytes()).hexdigest()) for p in sorted(used)])
    finally:
        for reader in readers.values():reader.close()
