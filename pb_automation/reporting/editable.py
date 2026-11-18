"""Editable Excel inputs feeding the existing presentation pages.

The input tables are a snapshot, not a second history database. Keys/row order
are protected; existing numeric cells may be edited. No macros/external links.
"""
from datetime import datetime
import re

from xlsxwriter.utility import xl_col_to_name as colname

from .report import PresentationReport, METRICS
from .pages import DetailedReport
from .data import _country, _stock_values
from .model import ratio, rolling_months, complete_sum
from .formatting import slide_amount, slide_amount_formula, slide_percent

NA = 'n.d.'


def ref(sheet, row, col):
    return f"'{sheet}'!${colname(col)}${row + 1}"


def numeric(value):
    return isinstance(value, (int, float))


def divide(a, b):
    return f'IF(OR(COUNT({a},{b})<>2,{b}=0),"{NA}",{a}/{b})'


def add(refs):
    return (f'IF(COUNT({",".join(refs)})={len(refs)},SUM({",".join(refs)}),"{NA}")'
            if refs else '0')


class EditableReport(PresentationReport):
    def __init__(self, path, history, period, config):
        super().__init__(path, history, period, config)
        self.countries = {(s['period'], e['id']): _country(e, s, s['period'], config)
                          for s in self.snapshots for e in s['entities']}
        self.stock_refs, self.flow_refs, self.dist_refs, self.metric_refs = {}, {}, {}, {}
        self.flow_rows = {}
        self.cst_refs, self.move_refs, self.note_refs = {}, {}, {}
        self.query_rows = {'Stock': [], 'Flux': [], 'Repartition stock': []}
        self.check_items = []
        self.edit_sheets = ['Stock', 'Flux', 'Repartition stock', 'Taux', 'Other inputs', 'Comments']
        for name in ['Edit guide', 'Other inputs', 'Comments', 'Checks']:
            self.sheets[name] = self.book.add_worksheet(name)
        for name, numfmt in [('edit_num', '#,##0.00;(#,##0.00);0'),
                             ('edit_int', '#,##0;(#,##0);0'), ('edit_pct', '0.0%'),
                             ('edit_rate', '0.00000000')]:
            self.formats[name] = self.book.add_format({'font_name': self.font, 'font_size': 10,
                'bg_color': 'FFF2CC', 'font_color': '1848A0', 'locked': False,
                'num_format': numfmt, 'align': 'right'})
        self.formats['edit_text'] = self.book.add_format({'font_name': self.font, 'font_size': 10,
            'bg_color': 'FFF2CC', 'font_color': '1848A0', 'locked': False,
            'text_wrap': True, 'valign': 'top'})
        self.formats['calc'] = self.book.add_format({'font_name': self.font, 'font_size': 10,
            'num_format': '#,##0.00;(#,##0.00);0', 'font_color': '17613A'})
        self.formats['guide_label'] = self.book.add_format({'font_name': self.font, 'font_size': 10,
            'bold': True, 'font_color': '16324F', 'text_wrap': True, 'valign': 'top'})

    def input_table(self, name, title, headers, note):
        self.start(name, title, headers, [16, 11, 24, 12] + [24] * (len(headers)-4), note)
        sh = self.sheets[name]
        sh.set_tab_color('D6AA36')
        sh.freeze_panes(5, 3)
        sh.hide_gridlines(2)
        sh.set_zoom(85)

    def input(self, sheet, row, col, value, fmt='edit_num'):
        # Empty editable cells are genuinely missing, never an implicit zero.
        sh = self.sheets[sheet]
        if value is None:
            sh.write_blank(row, col, None, self.formats[fmt])
        else:
            sh.write(row, col, value, self.formats[fmt])
        if fmt != 'edit_text':
            sh.data_validation(row, col, row, col, {'validate': 'integer' if fmt=='edit_int' else 'decimal', 'criteria': '>=',
                'value': 0, 'ignore_blank': True, 'error_title': 'Enter a number',
                'error_message': 'Use a non-negative number, or leave blank if unavailable.'})
        raw = ref(sheet, row, col)
        return f'IF(ISNUMBER({raw}),{raw},"{NA}")'

    def calc(self, sheet, row, col, expression, value, fmt='calc'):
        self.formula(sheet, row, col, '=' + expression.lstrip('='), value, fmt)
        return ref(sheet, row, col)

    def finish_table(self, sheet, row, cols):
        self.last[sheet] = row
        self.sheets[sheet].autofilter(4, 0, max(5, row-1), cols-1)

    def rates(self):
        headers = ['Reporting quarter', 'Currency', 'Closing date', 'Local units per EUR']
        self.input_table('Taux', 'Current-quarter exchange rates', headers,
            'One rate per currency. Both quarters use these rates: EUR amount = local amount / rate.')
        self.fxrefs, self.fxvalues = {}, {}
        current = self.history['snapshots'][self.period]
        for i, rate in enumerate(sorted(current['fx'], key=lambda r: r['currency']), 5):
            self.sheets['Taux'].write_row(i, 0, [self.period, rate['currency'], rate['date']])
            if rate['currency'] == 'EUR':
                self.write('Taux', i, 3, 1, 'rate')
                r = ref('Taux', i, 3)
            else:
                r = self.input('Taux', i, 3, rate['rate'], 'edit_rate')
                self.sheets['Taux'].data_validation(i, 3, i, 3,
                    {'validate': 'decimal', 'criteria': '>', 'value': 0, 'ignore_blank': False})
            for p in self.periods:
                self.fxrefs[p, rate['currency']] = r
                self.fxvalues[p, rate['currency']] = rate['rate']
        self.finish_table('Taux', 5 + len(current['fx']), 4)

    def stock(self):
        headers = ['Period', 'Code', 'Country', 'Currency', 'Row / vintage',
            'To process local', 'Processed balances held local', 'Total stock local',
            'Gross KPI2 (reference)', 'Processed KPI2 (reference)', 'Held balance rule',
            'To process EUR', 'Total stock EUR']
        self.input_table('Stock', 'KPI_2 stock inputs — totals and vintages', headers,
            'Yellow = edit. TOTAL drives summaries; vintage rows drive charts. Update both when appropriate. Green = formula.')
        row = 5
        for s in self.snapshots:
            p = s['period']
            for e in s['entities']:
                c = self.countries[p, e['id']]; m = c['metrics']; rate = c['rate']
                retained = None if m['identified'] is None or m['remaining'] is None else m['identified']-m['remaining']
                records = [('TOTAL', m['remaining'], retained, m['identified'], m['total'], m['processed'])]
                if e.get('stock'):
                    for r in e['stock']:
                        pending, total = _stock_values(e, r)
                        held = None if total is None or pending is None else total-pending
                        records.append((r['vintage'], pending, held, total, r.get('total'), r.get('processed_source')))
                else:
                    for v in c['global_vintages']:
                        pending = None if v['remaining_eur'] is None else v['remaining_eur']*rate
                        total = None if v['identified_eur'] is None else v['identified_eur']*rate
                        held = None if pending is None or total is None else total-pending
                        records.append((v['label'], pending, held, total, None, None))
                rule = ', '.join(e.get('identified_addback_fields', ['unreachable'] if 'unreachable' in e.get('processed_fields', []) else [])) or 'None / stock-only'
                for label, pending, held, total, gross, processed in records:
                    sh = self.sheets['Stock']; sh.write_row(row, 0, [p, e['id'], e['name'], e['currency'], label])
                    f = self.input('Stock', row, 5, pending)
                    g = self.input('Stock', row, 6, held)
                    h = (self.calc('Stock', row, 7, add([f, g]), total) if pending is not None and held is not None
                         else self.input('Stock', row, 7, total))
                    self.write('Stock',row,8,gross,'number');self.write('Stock',row,9,processed,'number')
                    i = ref('Stock',row,8); j = ref('Stock',row,9)
                    sh.write(row, 10, rule)
                    fx = self.fxrefs[p, e['currency']]
                    l = self.calc('Stock', row, 11, divide(f, fx), ratio(pending, rate))
                    n = self.calc('Stock', row, 12, divide(h, fx), ratio(total, rate))
                    self.stock_refs[p, e['id'], label] = {'remaining': f, 'held': g, 'identified': h,
                        'total': i, 'processed': j, 'remaining_eur': l, 'identified_eur': n}
                    if label != 'TOTAL':
                        self.query_rows['Stock'].append({'A': p, 'B': e['id'], 'E': label,
                                                        'refs': {'M': f, 'N': h, 'R': l, 'S': n}})
                    row += 1
                details = [self.stock_refs[p, e['id'], r[0]]['identified'] for r in records[1:]]
                self.check_items.append((p, e['name'], 'KPI2 vintage sum minus TOTAL', add(details),
                                        self.stock_refs[p, e['id'], 'TOTAL']['identified'],
                                        complete_sum(v[3] for v in records[1:]),m['identified']))
        self.finish_table('Stock', row, len(headers))

    def flows(self):
        headers = ['Period', 'Code', 'Country', 'Currency', 'Cohort month', 'Incoming amount',
            'Incoming contracts', 'Processed amount', 'Processed contracts', 'Client refund',
            'Unreachable', 'False positive', 'Authority refund', 'Other processing',
            'To process local', 'FX', 'Incoming EUR', 'Processed EUR', 'Efficiency amount', 'Efficiency number']
        self.input_table('Flux', 'KPI_2 monthly flow inputs', headers,
            'Amounts are local currency. Processed amount/count are authoritative inputs; keep breakdown categories consistent with them.')
        row = 5
        for (p, eid), c in self.countries.items():
            for m in c['monthly']:
                self.sheets['Flux'].write_row(row, 0, [p, eid, c['name'], c['currency']])
                self.sheets['Flux'].write_datetime(row, 4, datetime.strptime(m['month'], '%Y-%m'), self.formats['month'])
                cols = ['amount', 'count', 'processed_amount', 'processed_count', 'refund',
                        'unreachable', 'false_positive', 'authority_refund', 'other_without_refund']
                r = {k: self.input('Flux', row, 5+i, m[k], 'edit_int' if k in ('count','processed_count') else 'edit_num')
                     for i, k in enumerate(cols)}
                r['remaining'] = self.calc('Flux', row, 14,
                    self.arithmetic([r['amount'], r['processed_amount']], f"{r['amount']}-{r['processed_amount']}"), m['remaining'])
                fx = self.fxrefs[p, c['currency']]
                self.calc('Flux', row, 15, fx, c['rate'])
                for col, k in [(16, 'amount'), (17, 'processed_amount')]:
                    r[k+'_eur'] = self.calc('Flux', row, col, divide(r[k], fx), m[k+'_eur'])
                for col, k, a, b in [(18, 'efficiency_amount', 'processed_amount', 'amount'),
                                      (19, 'efficiency_count', 'processed_count', 'count')]:
                    r[k] = (self.input('Flux', row, col, m[k], 'edit_pct')
                            if m[k] is not None and (m[a] is None or m[b] is None)
                            else self.calc('Flux', row, col, divide(r[a], r[b]), m[k], 'percent'))
                    flag=f'IF(ISNUMBER({r[k]}),IF(OR({r[k]}<-0.00000001,{r[k]}>1.00000001),1,0),"n.d.")'
                    flag_value=None if m[k] is None else int(m[k] < -1e-8 or m[k] > 1+1e-8)
                    self.check_items.append((p,c['name'],m['month']+' '+k+' outside 0–100%',flag,'0',flag_value,0))
                self.flow_refs[p, eid, m['month']] = r
                self.flow_rows[p,eid,m['month']] = row
                mapping = {'F':'amount','G':'count','H':'refund','I':'unreachable','J':'false_positive',
                    'K':'authority_refund','L':'other_without_refund','M':'processed_amount','N':'processed_count',
                    'O':'remaining','R':'amount_eur','S':'processed_amount_eur','T':'efficiency_amount','U':'efficiency_count'}
                self.query_rows['Flux'].append({'A':p, 'B':eid, 'E':m['month'], 'refs':{col:r[k] for col,k in mapping.items()}})
                active = [k for k in cols[4:] if m[k] is not None]
                if active:
                    self.check_items.append((p, c['name'], m['month']+' categories minus processed',
                                            add([r[k] for k in active]), r['processed_amount'],
                                            complete_sum(m[k] for k in active),m['processed_amount']))
                row += 1
        self.finish_table('Flux', row, len(headers))

    def distribution(self):
        headers = ['Period', 'Code', 'Country', 'Currency', 'Distribution', 'Band / vintage',
                   'Contracts', 'Amount local', 'Amount EUR']
        self.input_table('Repartition stock', 'KPI_3 inputs — detail and reported totals', headers,
            'Yellow totals are reported independently: update them too when correcting details. KPI_3 does not overwrite KPI_2 stock.')
        row = 5
        for (p, eid), c in self.countries.items():
            for kind, label in [('stock_bins', 'Amount bands'), ('stock_vintages', 'Vintages')]:
                details = []
                for record in c[kind]:
                    self.sheets['Repartition stock'].write_row(row, 0, [p, eid, c['name'], c['currency'], label, record['label']])
                    count = self.input('Repartition stock', row, 6, record['count'], 'edit_int')
                    amount = self.input('Repartition stock', row, 7, record['amount'])
                    euro = self.calc('Repartition stock', row, 8, divide(amount, self.fxrefs[p,c['currency']]), record['amount_eur'])
                    details.append((count, amount))
                    self.query_rows['Repartition stock'].append({'A':p, 'B':eid, 'E':'Tranche' if kind=='stock_bins' else 'Millésime',
                        'F':record['label'], 'refs':{'G':count,'H':amount,'J':euro}})
                    row += 1
                total = c[kind+'_totals']
                self.sheets['Repartition stock'].write_row(row, 0, [p, eid, c['name'], c['currency'], label, 'TOTAL'])
                refs = {}
                for i, field in enumerate(('count','amount')):
                    refs[field] = (self.calc('Repartition stock', row, 6+i, add([d[i] for d in details]), total[field])
                        if total[field+'_basis']=='detail_sum' and details else
                        self.input('Repartition stock', row, 6+i, total[field], 'edit_int' if field=='count' else 'edit_num'))
                    if details:
                        self.check_items.append((p,c['name'],f'KPI3 {label}: {field} detail minus total',
                                                add([d[i] for d in details]),refs[field],
                                                complete_sum(d[field] for d in c[kind]),total[field]))
                self.calc('Repartition stock', row, 8, divide(refs['amount'], self.fxrefs[p,c['currency']]),total['amount_eur'])
                self.dist_refs[p,eid,kind] = refs
                if p == self.period:
                    for field in ('count','amount'):
                        self.input_refs[eid,kind,field] = refs[field]
                row += 1
        self.finish_table('Repartition stock', row, len(headers))
        self.other_inputs()

    def other_inputs(self):
        headers = ['Period', 'Code', 'Country', 'Currency', 'Item / month',
                   'Value / IN local', 'OUT local', 'IN contracts', 'OUT contracts', 'Closing stock local']
        self.input_table('Other inputs', 'CST balances and Brazil calendar movements', headers,
            'CST is separate from positive-balance stock. Brazil IN / OUT are calendar movements, not treatment efficiency.')
        row = 5
        for s in self.snapshots:
            p = s['period']
            for e in s['entities']:
                c = self.countries[p,e['id']]; m = c['metrics']
                if e.get('requires_cst') or e.get('cst'):
                    amount = m['cst'] if m['cst'] is not None else (None if m['cst_eur'] is None else m['cst_eur']*c['rate'])
                    self.sheets['Other inputs'].write_row(row,0,[p,e['id'],e['name'],e['currency'],'CST'])
                    self.cst_refs[p,e['id']] = self.input('Other inputs',row,5,amount)
                    row += 1
                for movement in sorted(e.get('calendar_movements', [])):
                    self.sheets['Other inputs'].write_row(row,0,[p,e['id'],e['name'],e['currency'],movement[0]])
                    self.move_refs[p,e['id'],movement[0]] = [self.input('Other inputs',row,i+5,v,
                        'edit_int' if i in (2,3) else 'edit_num') for i,v in enumerate(movement[1:])]
                    # Calendar IN is also Brazil's new monthly inflow. Keep a
                    # single editable source instead of independent duplicate
                    # inputs which could make the average and chart disagree.
                    flow_row=self.flow_rows.get((p,e['id'],movement[0]))
                    if flow_row is not None:
                        refs=self.move_refs[p,e['id'],movement[0]]
                        self.calc('Flux',flow_row,5,refs[0],movement[1])
                        self.calc('Flux',flow_row,6,refs[2],movement[3])
                    row += 1
        self.finish_table('Other inputs', row, len(headers))

    def consolidation(self):
        headers = ['Period','Code','Country','Currency','FX'] + METRICS
        self.start('Consolidation','Formula-linked indicators',headers)
        for row, ((p,eid), c) in enumerate(self.countries.items(),5):
            m = c['metrics']; r = self.stock_refs[p,eid,'TOTAL']; fx = self.fxrefs[p,c['currency']]
            flows = [self.flow_refs[p,eid,x['month']] for x in c['monthly']]
            sumflow = lambda key: add([f[key] for f in flows])
            cst = self.cst_refs.get((p,eid),'0')
            expr = {k:r[k] for k in ('total','processed','remaining','identified','remaining_eur','identified_eur')}
            entity = next(e for e in self.history['snapshots'][p]['entities'] if e['id']==eid)
            basis = entity.get('stock_count_basis','stock_bins')
            unreachable=r['held'] if 'unreachable' in entity.get('identified_addback_fields', ['unreachable']) else '"n.d."'
            expr.update(unreachable=unreachable,cst=cst,including_cst=add([r['identified'],cst]),
                including_cst_eur=add([r['identified_eur'],divide(cst,fx)]),
                contracts=self.dist_refs[p,eid,basis]['count'],flow_eur=sumflow('amount_eur'),
                processed_eur=sumflow('processed_amount_eur'),count=sumflow('count'),processed_count=sumflow('processed_count'),
                oldest_amount=flows[0]['efficiency_amount'],oldest_count=flows[0]['efficiency_count'])
            expr.update(efficiency_amount=divide(expr['processed_eur'],expr['flow_eur']),
                efficiency_count=divide(expr['processed_count'],expr['count']),monthly_average=divide(expr['flow_eur'],'6'))
            # Calendar IN remains the monthly-average basis for Brazil, without
            # inventing a cohort efficiency or classifying OUT as processed.
            moves = [v for (q,code,mo),v in self.move_refs.items() if (q,code)==(p,eid)]
            if moves:
                # Always use the six calendar months, never divide by the
                # number of supplied rows (which would hide missing months).
                expected=sorted(rolling_months(p))
                move_inputs=[self.move_refs.get((p,eid,month),['"n.d."'])[0] for month in expected]
                expr['monthly_average'] = divide(divide(add(move_inputs),'6'),fx)
            quarter = set(rolling_months(p,3))
            for key, field in [('quarter_flow_eur','amount_eur'),('quarter_processed_eur','processed_amount_eur')]:
                expr[key] = add([self.flow_refs[p,eid,mo][field] for mo in sorted(quarter)])
            expr['quarter_efficiency'] = divide(expr['quarter_processed_eur'],expr['quarter_flow_eur'])
            self.sheets['Consolidation'].write_row(row,0,[p,eid,c['name'],c['currency']])
            self.metric_refs[p,eid,'E'] = self.calc('Consolidation',row,4,fx,c['rate'])
            for i,key in enumerate(METRICS,5):
                self.metric_refs[p,eid,colname(i)] = self.calc('Consolidation',row,i,expr[key],m.get(key),
                    'percent' if key in ('efficiency_amount','efficiency_count','oldest_amount','oldest_count','quarter_efficiency') else 'calc')
            self.metrics[p,eid] = m
        self.last['Consolidation'] = 5+len(self.countries)

    def metric(self, period, eid, column):
        if (period,eid,column) in self.metric_refs:
            return '='+self.metric_refs[period,eid,column]
        members = next((r['entity_ids'] for r in self.config.get('consolidated_rows',[]) if r['id']==eid),[])
        if not members:
            return '="n.d."'
        if column=='E':
            return '=1'
        if column in ('X','Y'):
            month = min(rolling_months(period))
            keys = ('processed_amount_eur','amount_eur') if column=='X' else ('processed_count','count')
            expressions = [add([self.flow_refs.get((period,e,month),{}).get(k,'"n.d."') for e in members]) for k in keys]
            return '='+divide(*expressions)
        return '='+add([self.metric(period,e,column)[1:] for e in members])

    def query(self, sheet, column, period, eid, extra=None):
        members = next((r['entity_ids'] for r in self.config.get('consolidated_rows',[]) if r['id']==eid),[eid])
        refs = []
        for r in self.query_rows[sheet]:
            if r['A']!=period or r['B'] not in members:
                continue
            match = True
            for col, criterion in (extra or []):
                if criterion.startswith('DATE('):
                    y,m,_ = map(int,re.findall(r'\d+',criterion)); value=f'{y:04d}-{m:02d}'
                else:
                    value=criterion.strip('"').replace('""','"')
                match &= str(r.get(col))==value
            if match:
                refs.append(r['refs'][column])
        return add(refs) if refs else '"n.d."'

    def helper(self, title, headers, rows, formulas=None):
        if title.endswith(' calendar movements'):
            eid=title.split()[0]
            moves=sorted((mo,v) for (p,c,mo),v in self.move_refs.items() if p==self.period and c==eid)
            formulas=[[None]+['='+v for v in refs]+[None] for mo,refs in moves]
        block=DetailedReport.helper(self,title,headers,rows,formulas)
        if formulas:
            for i, row in enumerate(rows):
                for j,value in enumerate(row):
                    if value is None and formulas[i][j]:
                        self.sheets['GraphData'].write_formula(block['first']+i,j,formulas[i][j],self.pf['helper'],'')
        return block

    def stock_matrix(self, group):
        rows=[]; forms=[]
        for r in group['stock_rows']:
            eid=r['id']; members=r.get('metric_ids',[eid])
            prior=self.metric(self.data['previous_period'],eid,'N')[1:]
            now=self.metric(self.period,eid,'N')[1:]
            change=f'IF(OR(COUNT({prior},{now})<>2,{prior}=0),"n.d.",{now}/{prior}-1)'
            vals=[r['name'],r['previous_identified_eur'],r['current_identified_eur'],r['change'],r['current_remaining_eur']]
            f=[None,'='+prior,'='+now,'='+change,self.metric(self.period,eid,'M')]
            for key in ('identified_eur','remaining_eur'):
                for v in r['vintages']:
                    label=v['label']; refs=[]
                    for (p,code,source_label),fields in self.stock_refs.items():
                        if p!=self.period or code not in members or source_label=='TOTAL':continue
                        normalized=str(source_label).replace(' ','')
                        bucket=f'<{int(self.period[:4])-1}' if normalized.startswith('<') or int(normalized)<int(self.period[:4])-1 else normalized
                        if bucket==label: refs.append(fields[key])
                    vals.append(v[key]); f.append('='+add(refs))
            rows.append(vals); forms.append(f)
        headers=['Country','Prior stock EUR','Current stock EUR','Change','To process EUR']
        headers += [kind+v['label'] for kind in ('Total ','To process ') for v in group['global_vintages']]
        return self.helper(group['id']+' stock matrix',headers,rows,forms)

    def presentation_inputs(self):
        self.comment_row=4
        self.input_table('Comments','Manual presentation comments',['Code / group','Country / group','Comment'],
            'Qualitative notes are editable text. Numerical summary sentences are calculated separately and update automatically.')
        self.sheets['Comments'].set_column(0,0,18);self.sheets['Comments'].set_column(1,1,25);self.sheets['Comments'].set_column(2,2,105)
        items=self.data['countries']+self.data['groups']+[self.data['global']]
        for row,item in enumerate(items,5):
            text='\n'.join(item['comments'])
            self.sheets['Comments'].write_row(row,0,[item['id'],item['name']])
            self.input('Comments',row,2,text,'edit_text');self.sheets['Comments'].set_row(row,52)
            raw=ref('Comments',row,2);self.note_refs[item['id']]=(text,f'IF({raw}="","",{raw})')
        self.finish_table('Comments',5+len(items),3)
        self.write_guide_and_checks()

    def comment_box(self,sheet,r,c,r2,c2,lines,manual=()):
        eid=next((eid for eid,name in self.country_sheets.items() if name==sheet),None)
        if eid is None:
            eid='GLOBAL' if sheet=='Global' else 'core' if sheet.startswith('CORE') else 'noncore'
        if eid in self.note_refs:
            text,formula=self.note_refs[eid]
            lines=list(lines)+[(text,'='+formula)]
            manual=()
        return super().comment_box(sheet,r,c,r2,c2,lines,manual)

    def country_insights(self,country,sheet,notes_start=29):
        # Keep the established country layout and rounded numerical captions.
        entity=next(e for e in self.history['snapshots'][self.period]['entities'] if e['id']==country['id'])
        if not entity.get('calendar_movements') and any(r['amount'] is not None for r in country['monthly']):
            DetailedReport.country_insights(self,country,sheet,notes_start)
            if not country['requires_cst'] and country['previous_metrics'] is not None:
                p=self.metric(self.data['previous_period'],country['id'],'N')[1:];n=self.metric(self.period,country['id'],'N')[1:]
                expr=f'IF(OR(COUNT({p},{n})<>2,{p}=0),"unavailable",TEXT({n}/{p}-1,"0%"))'
                self.sheets[sheet].write_formula(notes_start+6,1,
                    f'="Total stock change vs {self.data["previous_period"]}: approximately "&{expr}&"."',self.pf['insight'],
                    f'Total stock change vs {self.data["previous_period"]}: approximately {slide_percent(country["metrics"]["stock_change"])}.')
            return
        currency=country['display_currency'];eid=country['id'];m=country['metrics']
        amt=m['identified_eur'] if currency=='EUR' else m['identified']
        target=self.metric(self.period,eid,'N' if currency=='EUR' else 'J')[1:]
        lines=[(f'Closing stock: {slide_amount(amt,currency)}.',f'="Closing stock: "&{slide_amount_formula(target,currency)}&"."')]
        if entity.get('calendar_movements'):
            avg=self.metric(self.period,eid,'W')[1:];fx=self.metric(self.period,eid,'E')[1:]
            lines.append((f"Average monthly IN: {slide_amount(m['monthly_average_local'],country['currency'])}.",
                f'="Average monthly IN: "&{slide_amount_formula(f"({avg})*({fx})",country["currency"])}&"."'))
            lines.append(('Calendar OUT includes older balances; cohort treatment efficiency is unavailable.',None))
        else:
            lines.append(('Monthly inflows and treatment efficiency have not been provided.',None))
        if m['stock_change'] is not None:
            p=self.metric(self.data['previous_period'],eid,'N')[1:];n=self.metric(self.period,eid,'N')[1:]
            change=f'IF(OR(COUNT({p},{n})<>2,{p}=0),"unavailable",TEXT({n}/{p}-1,"0%"))'
            lines.append((f"Stock change versus {self.data['previous_period']} at constant FX: approximately {slide_percent(m['stock_change'])}.",
                f'="Stock change versus {self.data["previous_period"]} at constant FX: approximately "&{change}&"."'))
        self.comment_box(sheet,notes_start,1,notes_start+10,31,lines,country['comments'])

    def noncore_tables(self,group,sheet,offset,block):
        super().noncore_tables(group,sheet,offset,block)
        for i,c in enumerate(x for x in group['countries'] if x.get('requires_cst')):
            refs=[]
            for p in (self.data['previous_period'],self.period):
                refs.append(divide(self.cst_refs.get((p,c['id']),'"n.d."'),self.fxrefs.get((p,c['currency']),self.fxrefs[self.period,c['currency']])))
            p,n=refs
            expressions=refs+[f'IF(OR(COUNT({p},{n})<>2,{p}=0),"n.d.",{n}/{p}-1)']
            previous=c.get('previous_cst') or {};current=c.get('cst') or {}
            vals=[previous.get('amount_eur',ratio(previous.get('amount'),c['rate'])),current.get('amount_eur',ratio(current.get('amount'),c['rate']))]
            vals.append(None if vals[0] in (None,0) or vals[1] is None else vals[1]/vals[0]-1)
            for col,expr,value in zip((22,26,30),expressions,vals):
                self.sheets[sheet].write_formula(offset+25+i,col,'='+expr,self.pf['pct' if col==30 else 'num_zero'],NA if value is None else value)

    def write_guide_and_checks(self):
        sh=self.sheets['Edit guide'];sh.hide_gridlines(2);sh.set_tab_color('00627A');sh.set_column('A:A',3);sh.set_column('B:B',27);sh.set_column('C:H',15)
        sh.merge_range('B2:H3','Edit a country and see the report update',self.formats['title'])
        lines=[('1. Choose the period and country','Use the filters in row 5 of an input tab. Yellow cells are editable. Formula cells and row keys are protected without a password.'),
            ('Stock → summary and charts','TOTAL is the authoritative KPI_2 summary. Vintage rows drive the stacked stock charts. Total stock = to process + processed balances still held. Update the TOTAL row and the affected vintage when both change.'),
            ('Flux → monthly tables and charts','Edit monthly local amounts and contract counts. Treatment % = processed / incoming. Processed totals and processing categories are independently reported; keep them consistent.'),
            ('Repartition stock → KPI_3 tables','Edit amount-band or vintage detail and its yellow TOTAL row. Reported totals are retained independently where they differ from detail. KPI_3 does not overwrite KPI_2: their scope can differ.'),
            ('Taux → both comparison quarters','Edit the current closing rate once. Both quarters convert local amounts using this rate. EUR stays at 1.'),
            ('Other inputs and Comments','Edit CST separately, Brazil calendar IN/OUT, or manual presentation text. Rounded numerical captions update automatically; manually written notes must be reviewed after edits.'),
            ('Checks → review discrepancies','Checks recalculate detail/total gaps. Some gaps already exist in the supplied data. They are shown rather than silently corrected. Blank inputs are unknown; country output displays 0 where requested.'),
            ('Keep Excel edits','Save this workbook under a new name. Edits update this Excel file only, not the country files or database/history.json. A Python rerun regenerates the report from those sources.'),
            ('Next quarter or new rows','For another period, country, vintage or month, update the country input files and rerun Python. Do not insert, delete or sort source rows here; filters are safe.'),
            ('Recalculation and chart scale','Excel calculation is set to Automatic. If Excel is in Manual mode, use Formulas → Calculate Now (F9). CORE stock axes share a fixed scale: after a large stock increase, adjust both chart maxima together or regenerate the report.')]
        for i,(title,text) in enumerate(lines):
            r=5+i*3;sh.merge_range(r,1,r+1,1,title,self.formats['guide_label']);sh.merge_range(r,2,r+1,7,text,self.pf['note'])
            sh.set_row(r,25);sh.set_row(r+1,25)
        self.input_table('Checks','Reconciliation checks',['Period','Country','Check','Detail / basis','Reported / comparison','Difference','Status'],
            'Non-zero gaps can be supplied rounding or scope differences. Review them; no input is silently changed.')
        for (p,eid),c in self.countries.items():
            self.check_items.append((p,c['name'],'KPI3 vintages minus KPI2 total stock',self.dist_refs[p,eid,'stock_vintages']['amount'],self.stock_refs[p,eid,'TOTAL']['identified'],c['stock_vintages_totals']['amount'],c['metrics']['identified']))
        for row,(p,name,label,a,b,av,bv) in enumerate(self.check_items,5):
            sh=self.sheets['Checks'];sh.write_row(row,0,[p,name,label]);self.calc('Checks',row,3,a,av);self.calc('Checks',row,4,b,bv)
            difference=av-bv if numeric(av) and numeric(bv) else None
            status='Unavailable' if difference is None else 'OK' if abs(difference)<0.005 else 'Review'
            r=row+1;self.calc('Checks',row,5,f'IF(COUNT(D{r},E{r})=2,D{r}-E{r},"n.d.")',difference)
            self.calc('Checks',row,6,f'IF(NOT(ISNUMBER(F{r})),"Unavailable",IF(ABS(F{r})<0.005,"OK","Review"))',status,'text')
        sh=self.sheets['Checks'];sh.set_column('C:C',49);sh.set_column('D:F',21);self.finish_table('Checks',5+len(self.check_items),7)
        sh.conditional_format(5,6,4+len(self.check_items),6,{'type':'text','criteria':'containing','value':'Review','format':self.formats['warning']})

    def flow_average_audit(self):
        name='Flow averages'
        self.sheets[name]=self.book.add_worksheet(name)
        headers=['Period','Code','Country','Currency','First month','Last month',
                 'Months available','Six-month incoming local','Average monthly local',
                 'Current FX','Average monthly EUR','Mean definition']
        self.input_table(name,'Average monthly new flow — calculation trail',headers,
            'Average = all six monthly incoming amounts / 6. Real zero months count; missing months do not become zero. Incoming amounts include false positives.')
        self.sheets[name].set_tab_color('00627A')
        for row,((p,eid),c) in enumerate(self.countries.items(),5):
            months=sorted(rolling_months(p));m=c['metrics']
            source=[self.flow_refs[p,eid,month]['amount'] for month in months]
            move_keys=[(p,eid,month) for month in months]
            if any(key in self.move_refs for key in move_keys):
                source=[self.move_refs.get(key,['"n.d."'])[0] for key in move_keys]
                values={r[0]:r[1] for e in self.history['snapshots'][p]['entities'] if e['id']==eid for r in e.get('calendar_movements',[])}
                count=sum(values.get(mo) is not None for mo in months)
                total=complete_sum(values.get(mo) for mo in months)
                definition='Calendar IN / 6; not OUT or net movement'
            else:
                values=[r['amount'] for r in c['monthly']]
                count=sum(v is not None for v in values);total=complete_sum(values)
                definition='KPI2 incoming flow / 6; not processed or remaining'
            sh=self.sheets[name]
            sh.write_row(row,0,[p,eid,c['name'],c['currency'],months[0],months[-1]])
            self.calc(name,row,6,f'COUNT({",".join(source)})',count,'integer')
            totalref=self.calc(name,row,7,add(source),total)
            self.calc(name,row,8,divide(totalref,'6'),ratio(total,6))
            fx=self.fxrefs[p,c['currency']]
            self.calc(name,row,9,fx,c['rate'],'rate')
            self.calc(name,row,10,self.metric(p,eid,'W'),m['monthly_average'])
            sh.write(row,11,definition)
        self.sheets[name].set_column('G:G',18);self.sheets[name].set_column('L:L',52)
        self.finish_table(name,5+len(self.countries),len(headers))

    def finalize_workbook(self):
        for name in self.edit_sheets:
            self.sheets[name].hidden=0
            self.sheets[name].protect('',{'autofilter':True,'select_locked_cells':True,'select_unlocked_cells':True})
        if self.comment_row>4:
            self.sheets['Presentation inputs'].hidden=0
        for name in ('Global','CORE Stock','CORE Flows','NON CORE Stock'):
            if name in self.sheets:
                self.sheets[name].write_url('A1',"internal:'Edit guide'!B2",self.formats['note'],'Edit inputs →')
