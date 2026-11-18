"""Turn validated quarterly snapshots into the detailed Excel presentation.

The history is never mutated. Amounts in the comparison quarter are translated
at the reporting quarter's closing rates on a temporary presentation copy.
"""
from copy import deepcopy
from datetime import datetime
from pathlib import Path

from .data import _country, _group, _stock_values, previous_period
from .model import parse_period, ratio
from .pages import DetailedReport
from .formatting import slide_amount, slide_percent, stock_axis

METRICS = ["total", "processed", "remaining", "unreachable", "identified", "cst",
           "including_cst", "remaining_eur", "identified_eur", "including_cst_eur",
           "contracts", "flow_eur", "processed_eur", "count", "processed_count",
           "efficiency_amount", "efficiency_count", "monthly_average", "oldest_amount",
           "oldest_count", "quarter_flow_eur", "quarter_processed_eur", "quarter_efficiency"]
SOURCE_COLUMNS = {"Flux": 22, "Stock": 19, "Repartition stock": 10, "Taux": 5}


def presentation_history(history, period):
    """Keep the current/prior snapshots and apply constant-FX comparisons."""
    period = parse_period(period)
    if period not in history.get("snapshots", {}):
        raise ValueError(f"No validated data for {period}")
    prior = previous_period(period)
    copied = {"schema_version": 1, "snapshots": {
        key: deepcopy(history["snapshots"][key])
        for key in (prior, period) if key in history["snapshots"]}}
    current = copied["snapshots"][period]
    rates = {row["currency"]: row["rate"] for row in current["fx"]}
    rates.setdefault("EUR", 1)
    if not any(row["currency"] == "EUR" for row in current["fx"]):
        current["fx"].append({"currency": "EUR", "rate": 1,
                              "date": current.get("date", period[:4] + "-01-01"), "source": "EUR"})
    for entity in current["entities"]:
        if entity["currency"] not in rates or not rates[entity["currency"]]:
            raise ValueError(f"Missing closing rate for {entity['currency']}")
    if prior in copied["snapshots"]:
        old = copied["snapshots"][prior]
        old_rates = {row["currency"]: row["rate"] for row in old["fx"]}
        active = {entity["id"] for entity in current["entities"]}
        old["entities"] = [e for e in old["entities"] if e["id"] in active]
        for entity in old["entities"]:
            currency = entity["currency"]
            factor = old_rates.get(currency, 1 if currency == "EUR" else None)
            if factor is None:
                raise ValueError(f"Historical closing rate missing for {currency} in {prior}")
            factor /= rates[currency]
            # Raw KPI2 data remains in local currency. Only stored EUR observations
            # need restatement; every other conversion uses the copied rate table.
            for key in ("historical_vintages", "historical_flows"):
                for row in entity.get(key, []):
                    for name, value in list(row.items()):
                        if name.endswith("_eur") and isinstance(value, (int, float)):
                            row[name] = value * factor
            for name, value in list(entity.get("historical_metrics", {}).items()):
                if (name.endswith("_eur") or name == "monthly_average") and isinstance(value, (int, float)):
                    entity["historical_metrics"][name] = value * factor
            cst = entity.get("cst", {})
            if isinstance(cst.get("amount_eur"), (int, float)):
                cst["amount_eur"] *= factor
        old["fx"] = deepcopy(current["fx"])
    return copied


class PresentationReport(DetailedReport):
    """Presentation pages plus hidden calculation tables, with no file columns."""

    def __init__(self, path, history, period, config):
        super().__init__(path, history, period, config)
        for group in self.data["groups"] + [self.data["global"]]:
            group["scope_note"] = (
                f"{self.data['previous_period']} and {period} translated at {period} closing FX. "
                f"{len(group['entity_ids'])} countries included.")
        for key in ("num", "total"):
            extras = {"font_name": self.font, "font_size": 10, "valign": "vcenter",
                      "border": 1, "border_color": "8B969B", "align": "right",
                      "num_format": "#,##0;(#,##0);0"}
            if key == "total":
                extras.update(bg_color=self.color("total_fill"), bold=True)
            self.pf[key + "_zero"] = self.book.add_format(extras)

    def write(self, sheet, row, col, value, fmt="text"):
        if SOURCE_COLUMNS.get(sheet) == col:
            return
        return super().write(sheet, row, col, value, fmt)

    def start(self, sheet, title, headers, widths=None, note=None):
        if headers and headers[-1] == "Source":
            headers = headers[:-1]
            widths = widths[:-1] if widths else None
        return super().start(sheet, title, headers, widths, note)

    def table(self, sheet, headers, count, name):
        return super().table(sheet, headers[:-1] if headers[-1] == "Source" else headers, count, name)

    def merged(self, sheet, r, c, r2, c2, value, fmt="body", formula=None):
        if sheet in self.country_sheets.values():
            if value is None:
                value = 0
            if isinstance(value, str):
                value = value.replace("n.d.", "0")
            if formula and fmt in ("num", "pct", "total", "totalpct"):
                expr = formula.lstrip("=")
                formula = f'=IF(({expr})="n.d.",0,{expr})'
            elif formula and fmt == "insight":
                formula = formula.replace(',"n.d.")', ',"0")')
            if fmt in ("num", "total"):
                fmt += "_zero"
        elif value == 0 and fmt in ("num", "total"):
            fmt += "_zero"
        return super().merged(sheet, r, c, r2, c2, value, fmt, formula)

    def consolidation(self):
        headers = ["Trimestre", "Code pays", "Pays", "Devise", "Taux local pour 1 EUR"] + METRICS
        self.start("Consolidation", "Quarterly indicators", headers, [18] * len(headers),
                   "Calculated during the reporting run; rerun after changing country files.")
        row = 5
        for snapshot in self.snapshots:
            for entity in snapshot["entities"]:
                country = _country(entity, snapshot, snapshot["period"], self.config)
                metrics = country["metrics"]
                self.metrics[(snapshot["period"], country["id"])] = metrics
                values = [snapshot["period"], country["id"], country["name"], country["currency"], country["rate"]]
                for col, value in enumerate(values + [metrics.get(key) for key in METRICS]):
                    fmt = "percent" if col in (20, 21, 23, 24, 27) else "number" if col >= 4 else "text"
                    self.write("Consolidation", row, col, value, fmt)
                row += 1
        self.table("Consolidation", headers, row - 5, "Indicateurs")

    def helper(self, title, headers, rows, formulas=None):
        # The Python pipeline owns semantic transformations; page formulas and
        # native charts link to this run's calculated helper cells.
        return super().helper(title, headers, rows, None)

    def metric(self, period, eid, column):
        combined = next((r for r in self.config.get("consolidated_rows", []) if r["id"] == eid), None)
        if combined and column == "W":
            refs = [super(PresentationReport, self).metric(period, member, column).lstrip("=")
                    for member in combined["entity_ids"]]
            return self.arithmetic(refs, f'SUM({",".join(refs)})')
        return super().metric(period, eid, column)

    def stock(self):
        super().stock()
        row = 5
        columns = {"unreachable": "H", "authority_refund": "J", "refund": "G",
                   "false_positive": "I", "other_without_refund": "K"}
        for snapshot in self.snapshots:
            for entity in snapshot["entities"]:
                for record in entity["stock"]:
                    if entity.get("stock_basis") == "source_remaining":
                        r = row + 1
                        remaining, identified = _stock_values(entity, record)
                        rate = self.fxvalues[(snapshot["period"], entity["currency"])]
                        addbacks = entity.get("identified_addback_fields",
                                              ["unreachable"] if "unreachable" in entity["processed_fields"] else [])
                        refs = [f"M{r}"] + [f"{columns[k]}{r}" for k in addbacks]
                        self.formula("Stock", row, 12, f"=P{r}", remaining)
                        self.formula("Stock", row, 13, self.arithmetic(refs, "+".join(refs)), identified)
                        self.formula("Stock", row, 17, self.arithmetic([f"M{r}", f"Q{r}"], f"M{r}/Q{r}"), ratio(remaining, rate))
                        self.formula("Stock", row, 18, self.arithmetic([f"N{r}", f"Q{r}"], f"N{r}/Q{r}"), ratio(identified, rate))
                    row += 1

    def flows(self):
        super().flows()
        row = 5
        for snapshot in self.snapshots:
            for entity in snapshot["entities"]:
                for record in entity["flows"]:
                    if entity.get("flow_basis") == "source_processed":
                        r = row + 1
                        amount, processed = record["amount"], record["processed_amount"]
                        rate = self.fxvalues[(snapshot["period"], entity["currency"])]
                        self.formula("Flux", row, 12, f"=V{r}", processed)
                        remaining = None if amount is None or processed is None else amount - processed
                        self.formula("Flux", row, 14, self.arithmetic([f"F{r}", f"M{r}"], f"F{r}-M{r}"), remaining)
                        self.formula("Flux", row, 18, self.arithmetic([f"M{r}", f"Q{r}"], f"M{r}/Q{r}"), ratio(processed, rate))
                        self.formula("Flux", row, 19, f'=IF(OR(COUNT(F{r},M{r})<>2,F{r}=0),"n.d.",M{r}/F{r})', ratio(processed, amount), "percent")
                    row += 1

    def stock_matrix(self, group):
        rows = [[r["name"], r["previous_identified_eur"], r["current_identified_eur"], r["change"], r["current_remaining_eur"]]
                + [v[k] for k in ("identified_eur", "remaining_eur") for v in r["vintages"]]
                for r in group["stock_rows"]]
        headers = ["Country", "Prior identified EUR", "Current identified EUR", "Change", "Remaining EUR"]
        headers += [kind + v["label"] for kind in ("Identified ", "Remaining ") for v in group["global_vintages"]]
        return self.helper(group["id"] + " stock matrix", headers, rows)

    def stock_page(self, group, sheet, offset=0, noncore=False):
        self._stock_axis = stock_axis(group['stock_rows']) if sheet == 'CORE Stock' else None
        try:
            super().stock_page(group, sheet, offset, noncore)
        finally:
            self._stock_axis = None
        self.sheets[sheet].write_string(offset + 7, 7, self.data["previous_period"] + "\nat current FX", self.pf["header"])
        note = max(offset + 13 + len(group["countries"]), offset + 23)
        text = "Total stock = stock to process + processed balances still held, using the country reporting rules."
        if noncore:
            text = "Nordic stock excludes separate CST account balances. Stock-only countries have no inferred monthly flows."
        self.sheets[sheet].write_string(note, 1, text, self.pf["insight"])
        self.sheets[sheet].write_string(note + 2, 1, "Both quarters use the reporting quarter’s closing exchange rates.", self.pf["insight"])

    def vintage_chart(self, block, title, firstcol, width=595, height=250):
        chart = super().vintage_chart(block, title, firstcol, width, height)
        if getattr(self, '_stock_axis', None):
            chart.set_y_axis({**self._stock_axis, 'num_format':'#,##0',
                              'num_font':{'name':self.font,'size':8},
                              'major_gridlines':{'visible':True,'line':{'color':'DDE3E5','width':.5}},
                              'line':{'none':True}})
        if len(block["rows"]) > 9:
            for i, row in enumerate(block["rows"]):
                short = self.config.get("chart_short_names", {}).get(row[0])
                if short:
                    self.sheets["GraphData"].write_string(block["first"] + i, 0, short)
            chart.set_x_axis({"num_font": {"size": 8, "rotation": -45}, "interval_unit": 1})
        return chart

    def flow_page(self, group, sheet, offset=0):
        members = group.get("underlying_countries", group["countries"])
        included = [c for c in members if any(m["amount"] is not None for m in c["monthly"])]
        if not included:
            self.page(sheet, group["name"] + ": Monthly flows", offset)
            self.merged(sheet, offset + 7, 1, offset + 13, 31, "Monthly flows have not been provided.", "header")
            return
        if len(included) != len(members):
            group = _group(group["id"], group["name"], included, self.data["months"], self.config, self.period)
            missing = ", ".join(c["name"] for c in members if c not in included)
            group["scope_note"] = f"Monthly flows: {len(included)} countries. Not provided: {missing}."
        return super().flow_page(group, sheet, offset)

    def country_chart(self, country, sheet):
        entity = next(e for e in self.history["snapshots"][self.period]["entities"] if e["id"] == country["id"])
        movements = entity.get("calendar_movements", [])
        if movements:
            rows = [[datetime.strptime(r[0], "%Y-%m").strftime("%b-%y")] + list(r[1:]) + [0] for r in sorted(movements)]
            block = self.helper(country["id"] + " calendar movements", ["Month", "IN", "OUT", "IN count", "OUT count", "Closing stock", "Spacing"], rows)
            chart = self.book.add_chart({"type": "column"})
            for name, col, color, secondary in [("IN (left axis)", 1, "D66B28", False), ("", 6, None, False),
                                               ("", 6, None, True), ("OUT (right axis)", 2, "77B900", True)]:
                chart.add_series({"name": name, "categories": self.series_range(block, 0), "values": self.series_range(block, col),
                                  "y2_axis": secondary, "gap": 80, "overlap": 0,
                                  "fill": {"color": color} if color else {"none": True}, "border": {"none": True}})
            currency = country["currency"]
            self.chart_style(chart, f"{country['name']}: monthly IN and OUT in {currency}", 610, 320)
            chart.set_legend({"position": "bottom", "delete_series": [1, 2], "font": {"size": 9}})
            chart.set_y_axis({"name": f"IN — {currency}", "min": 0, "num_format": "#,##0", "name_font": {"color": "D66B28", "size": 9}})
            chart.set_y2_axis({"name": f"OUT — {currency}", "min": 0, "num_format": "#,##0", "name_font": {"color": "548300", "size": 9}, "major_gridlines": {"visible": False}})
            self.insert(sheet, 6, 1, chart)
            self.merged(sheet, 25, 1, 27, 15, "IN: orange, left axis. OUT: green, right axis. Axes use different scales. Calendar OUT is not cohort treatment.", "note")
            return block
        if not any(r["amount"] is not None for r in country["monthly"]):
            self.merged(sheet, 7, 1, 18, 15, "Monthly flow and treatment data have not been provided.", "header")
            return None
        return super().country_chart(country, sheet)

    def stock_distribution(self, country, sheet, kind, start):
        if not country[kind]:
            label = "Amount-band" if kind == "stock_bins" else "Vintage"
            self.merged(sheet, start, 17, start + 2, 31, label + " distribution has not been provided.", "note")
            return start + 4
        return super().stock_distribution(country, sheet, kind, start)

    def country_page(self, country):
        super().country_page(country)
        sheet = self.sheets[self.country_sheets[country["id"]]]
        sheet.write_string(1, 17, "STOCK OF POSITIVE BALANCES\nBreakdown by amount and vintage", self.pf["header"])
        note = self.config.get("country_options", {}).get(country["id"], {}).get("note", "Stock breakdowns retain each country’s reporting scope.")
        sheet.write_string(4, 1, f"{self.period} — {country['display_currency']}. {note}", self.pf["note"])

    def country_insights(self, country, sheet, notes_start=29):
        entity = next(e for e in self.history["snapshots"][self.period]["entities"] if e["id"] == country["id"])
        m = country["metrics"]
        if entity.get("calendar_movements") or not any(r["amount"] is not None for r in country["monthly"]):
            currency = country["display_currency"]
            amount = m["identified_eur"] if currency == "EUR" else m["identified"]
            lines = [(f"Closing stock: {slide_amount(amount,currency)}." if amount is not None else "Closing stock unavailable.", None)]
            if entity.get("calendar_movements"):
                lines.append((f"Average monthly IN: {slide_amount(m['monthly_average_local'],country['currency'])}.", None))
                lines.append(("Calendar OUT includes older balances; cohort treatment efficiency is unavailable.", None))
            else:
                lines.append(("Monthly inflows and treatment efficiency have not been provided.", None))
            if m["stock_change"] is not None:
                lines.append((f"Stock change versus {self.data['previous_period']} at constant FX: approximately {slide_percent(m['stock_change'])}.", None))
            self.comment_box(sheet, notes_start, 1, notes_start + 10, 31, lines, country["comments"])
            return
        super().country_insights(country, sheet, notes_start)

    def noncore_tables(self, group, sheet, offset, block):
        super().noncore_tables(group, sheet, offset, block)
        self.sheets[sheet].write_string(offset + 30, 17, "CST accounts are separate from positive-balance stock. Missing CST balances remain unavailable.\n\nStock-only countries contribute to closing stock, not to inferred monthly flows or efficiency.", self.pf["note"])
        countries = [c for c in group["countries"] if c.get("requires_cst")]
        for i, c in enumerate(countries):
            prior = c.get("previous_cst") or {}
            current = c.get("cst") or {}
            values = [prior.get("amount_eur", ratio(prior.get("amount"), c["rate"])),
                      current.get("amount_eur", ratio(current.get("amount"), c["rate"]))]
            values.append(None if values[0] in (None, 0) or values[1] is None else values[1] / values[0] - 1)
            for col, value in zip((22, 26, 30), values):
                if value is not None:
                    self.sheets[sheet].write_number(offset + 25 + i, col, value, self.pf["pct" if col == 30 else "num_zero"])

    def build(self):
        try:
            self.rates(); self.stock(); self.flows(); self.distribution(); self.consolidation()
            self.presentation_inputs()
            self.flow_average_audit()
            self.stock_page(self.data["global"], "Global")
            self.flow_page(self.data["global"], "Global", 44)
            self.sheets["Global"].set_h_pagebreaks([44])
            self.sheets["Global"].write_string(0, 1, f"{self.period} — {len(self.data['countries'])} countries — comparison at current-quarter FX", self.pf["note"])
            for group in self.data["groups"]:
                if group["id"] not in self.group_sheets:
                    continue
                stock, flows = self.group_sheets[group["id"]]
                self.stock_page(group, stock, noncore=group["id"] == "noncore")
                if flows:
                    self.flow_page(group, flows)
            for country in self.data["countries"]:
                self.country_page(country)
            self.additional_amount_bands()
            for name in ("Consolidation", "Flux", "Stock", "Repartition stock", "GraphData", "Presentation inputs"):
                self.sheets[name].hide()
            self.sheets["Taux"].write_string(2, 0, "Both quarters use the reporting quarter’s closing rates for the comparison.", self.formats["note"])
            for chart in self.book.charts:
                chart.show_hidden_data()
            self.sheets["Global"].activate()
            self.finalize_workbook()
        finally:
            self.book.close()

    def finalize_workbook(self):
        """Optional final pass for report variants, before the workbook is saved."""

    def flow_average_audit(self):
        """Optional input-linked mean audit, supplied by the editable variant."""

    def additional_amount_bands(self):
        """Local-valued source distributions without inventing band thresholds."""
        ids=self.config.get('additional_amount_band_countries',[])
        countries=[c for c in self.data['countries'] if c['id'] in ids]
        if not countries:return
        name='UK ZA Amount bands'
        sh=self.book.add_worksheet(name);self.sheets[name]=sh
        sh.hide_gridlines(2);sh.set_tab_color('00627A');sh.set_column('A:A',3)
        sh.set_column('B:B',28);sh.set_column('C:D',24);sh.set_column('E:F',20)
        sh.freeze_panes(5,2);sh.set_landscape();sh.fit_to_pages(1,1)
        sh.merge_range('B2:F3',f'Amount-band detail — {self.period}',self.formats['title'])
        for key,pattern in [('band_count','#,##0;(#,##0);0'),
                            ('band_money','#,##0.00;(#,##0.00);0')]:
            base={'font_name':self.font,'font_size':10,'align':'right','num_format':pattern}
            self.formats[key]=self.book.add_format(base)
            self.formats[key+'_total']=self.book.add_format({**base,'bold':True,'top':1,'top_color':'9DAAB6'})
        row=5
        for c in countries:
            currency=c['currency'];eid=c['id'];rows=c['stock_bins'];total=c['stock_bins_totals']
            sh.merge_range(row,1,row,5,f"{c['name']} — amounts in {currency}",self.pf['header'])
            row+=1
            note=self.config.get('country_options',{}).get(eid,{}).get('amount_band_note',
                'Source band thresholds are retained. Amounts are in the country reporting currency; no threshold conversion is inferred.')
            sh.merge_range(row,1,row+2,5,note,self.pf['note']);row+=4
            sh.write_row(row,1,['Source band label','Contracts',f'Amount {currency}','% in number','% in amount'],self.pf['header']);sh.set_row(row,32);row+=1
            first=row;last=first+len(rows)
            for record in rows:
                label=record['label'];sh.write(row,1,label,self.pf['body'])
                criteria=[('E','"Tranche"'),('F','"'+label.replace('"','""')+'"')]
                for col,key,source in [(2,'count','G'),(3,'amount','H')]:
                    self.formula(name,row,col,'='+self.query('Repartition stock',source,self.period,eid,criteria),record[key],'band_count' if key=='count' else 'band_money')
                for col,key,letter in [(4,'pct_count','C'),(5,'pct_amount','D')]:
                    expr=f'=IF(OR(COUNT({letter}{row+1},{letter}{last+1})<>2,{letter}{last+1}=0),"n.d.",{letter}{row+1}/{letter}{last+1})'
                    self.formula(name,row,col,expr,record[key],'percent')
                row+=1
            sh.write(row,1,'Reported total',self.pf['total'])
            for col,key in [(2,'count'),(3,'amount')]:
                if total[key+'_basis']=='detail_sum':
                    letter='C' if col==2 else 'D';rng=f'{letter}{first+1}:{letter}{last}'
                    expr=f'=IF(COUNT({rng})={len(rows)},SUM({rng}),"n.d.")'
                else:expr='='+self.input_refs[eid,'stock_bins',key]
                self.formula(name,row,col,expr,total[key],'band_count_total' if key=='count' else 'band_money_total')
            for col,letter in [(4,'C'),(5,'D')]:
                self.formula(name,row,col,f'=IF(OR(NOT(ISNUMBER({letter}{row+1})),{letter}{row+1}=0),"n.d.",1)',
                             1 if total['count' if col==4 else 'amount'] not in (None,0) else None,'percent')
            row+=4
        sh.print_area(0,0,row,5)


def generate_report(history, period, output, config):
    """Write one detailed workbook; keep input history and configuration unchanged."""
    period = parse_period(period)
    config = deepcopy(config)
    config.setdefault("chronological_month_tables", True)
    config.setdefault("shared_efficiency_country_column", True)
    config.setdefault("percentage_axis_headroom", False)
    config.setdefault("only_nordic_cst_rows", True)
    output = Path(output)
    output.parent.mkdir(parents=True, exist_ok=True)
    prepared = presentation_history(history, period)
    for snapshot in prepared["snapshots"].values():
        for rate in snapshot["fx"]:
            rate.setdefault("source", "")
    if config.get("editable_inputs", True):
        from .editable import EditableReport
        report = EditableReport(output, prepared, period, config)
    else:
        report = PresentationReport(output, prepared, period, config)
    report.build()
