"""Generate a self-contained Excel report, with formulas and Python-calculated caches.

Inputs are snapshots, never workbook links. Excel tables provide filters and stable
key-based summaries. A new quarterly run regenerates all table extents.
"""
from __future__ import annotations

from datetime import datetime
from pathlib import Path

import xlsxwriter
from xlsxwriter.utility import xl_col_to_name as colname

from .model import complete_sum, ratio, rolling_months

NA = "n.d."
SHEETS = ("Consolidation", "Flux", "Stock", "Repartition stock", "Taux")
CATEGORIES = ("refund", "unreachable", "false_positive", "authority_refund", "other_without_refund")


def subtract(a, b):
    return None if a is None or b is None else a - b


def divide(a, b):
    return ratio(a, b)


class Report:
    def __init__(self, path, history, period, front_sheets=None):
        self.history, self.period = history, period
        self.snapshots = [history["snapshots"][p] for p in sorted(history["snapshots"])]
        self.periods = [s["period"] for s in self.snapshots]
        self.entities = {}
        for snap in self.snapshots:
            self.entities.update({e["id"]: e["name"] for e in snap["entities"]})
        self.book = xlsxwriter.Workbook(str(path), {"strings_to_formulas": False, "strings_to_urls": False})
        self.book.set_properties({"title": "Positive Balance", "subject": "Consolidation trimestrielle", "author": "Positive Balance"})
        self.book.set_calc_mode("auto")
        sheet_names = tuple(front_sheets or ()) + SHEETS
        self.sheets = {name: self.book.add_worksheet(name) for name in sheet_names}
        base = {"font_name": "Arial", "font_size": 10, "valign": "vcenter"}
        self.formats = {}
        for name, extras in {
            "text": {}, "input": {"font_color": "#1848A0"},
            "number": {"num_format": '#,##0.00;(#,##0.00);"–"', "align": "right"},
            "integer": {"num_format": '#,##0;(#,##0);"–"', "align": "right"},
            "percent": {"num_format": "0.0%", "italic": True, "align": "right"},
            "rate": {"num_format": "0.00000000", "align": "right"},
            "date": {"num_format": "dd/mm/yyyy"},
            "month": {"num_format": "mmm-yy"},
            "title": {"font_size": 16, "bold": True, "font_color": "#16324F"},
            "section": {"bold": True, "font_color": "#16324F", "bottom": 1, "bottom_color": "#9DAAB6"},
            "header": {"bold": True, "bg_color": "#16324F", "font_color": "#FFFFFF", "text_wrap": True, "align": "center"},
            "note": {"font_color": "#566474", "italic": True},
            "selector": {"bg_color": "#FFF0BA", "font_color": "#1848A0", "bold": True},
            "warning": {"bg_color": "#FFF0E3", "font_color": "#993B15"},
            "missing": {"font_color": "#993B15", "align": "right"},
            "total": {"bold": True, "top": 1, "top_color": "#9DAAB6", "num_format": '#,##0.00;(#,##0.00);"–"', "align": "right"},
        }.items():
            self.formats[name] = self.book.add_format({**base, **extras})
        self.last = {}
        self.metrics = {}
        self.records = {}
        for name, sheet in self.sheets.items():
            sheet.hide_gridlines(2)
            sheet.set_default_row(19)
            sheet.set_column("A:AB", 18, self.formats["text"])
            sheet.set_landscape()
            sheet.set_paper(9)
            sheet.fit_to_pages(1, 0)
            sheet.set_margins(.3, .3, .4, .4)
            if name not in {"Synthese", "Mode emploi"}:
                sheet.freeze_panes(5, 2)
                sheet.repeat_rows(0, 4)
        for name in ("Synthese", "Evolution flux"):
            if name in self.sheets:
                self.sheets[name].set_tab_color("#16324F")

    def write(self, sheet, row, col, value, fmt="text"):
        sh = self.sheets[sheet]
        if value is None:
            sh.write_string(row, col, NA, self.formats["missing"])
        elif isinstance(value, datetime):
            sh.write_datetime(row, col, value, self.formats[fmt])
        else:
            sh.write(row, col, value, self.formats[fmt])

    def formula(self, sheet, row, col, formula, value, fmt="number"):
        self.sheets[sheet].write_formula(row, col, formula, self.formats[fmt], NA if value is None else value)

    def start(self, sheet, title, headers, widths=None, note=None):
        sh = self.sheets[sheet]
        sh.write(1, 0, title, self.formats["title"])
        if note:
            sh.write(2, 0, note, self.formats["note"])
        sh.set_row(4, 42)
        for c, header in enumerate(headers):
            self.write(sheet, 4, c, header, "header")
        if widths:
            for c, width in enumerate(widths):
                sh.set_column(c, c, width, self.formats["text"])
        self.records[sheet] = []

    def table(self, sheet, headers, count, name):
        sh = self.sheets[sheet]
        if count:
            sh.add_table(4, 0, 4 + count, len(headers) - 1,
                         {"name": name, "columns": [{"header": h, "header_format": self.formats["header"]} for h in headers],
                          "style": "Table Style Medium 2"})
            sh.conditional_format(5, 0, 4 + count, len(headers) - 1,
                                  {"type": "text", "criteria": "containing", "value": NA, "format": self.formats["missing"]})
        self.last[sheet] = 5 + count

    def rng(self, sheet, col):
        return f"'{sheet}'!${col}$6:${col}${max(6,self.last[sheet])}"

    def sumkey(self, source, column, target_row, extra=None):
        """Keyed sums stay correct after table sorting. Unknowns cannot sum to zero."""
        amount, periods, ids = (self.rng(source, c) for c in (column, "A", "B"))
        terms = f'{periods},$A{target_row},{ids},$B{target_row}'
        if extra:
            for col, criterion in extra:
                terms += f',{self.rng(source, col)},{criterion}'
        return f'IF(OR(COUNTIFS({terms})=0,COUNTIFS({terms},{amount},"{NA}")>0),"{NA}",SUMIFS({amount},{terms}))'

    @staticmethod
    def arithmetic(refs, expression):
        values = ",".join(refs)
        return f'=IF(COUNT({values})={len(refs)},{expression},"{NA}")'

    def rates(self):
        headers = ["Trimestre", "Devise", "Date des taux", "Devise locale pour 1 EUR", "Type de taux", "Source"]
        self.start("Taux", "Taux de conversion", headers, [15, 12, 17, 25, 20, 64],
                   "Montant EUR = montant en devise locale / taux de clôture. Même convention que la synthèse fournie.")
        self.fxrefs, self.fxvalues = {}, {}
        row = 5
        for snap in self.snapshots:
            for rate in snap["fx"]:
                data = [snap["period"], rate["currency"], datetime.fromisoformat(rate["date"]), rate["rate"], "Clôture", rate["source"]]
                for col, val in enumerate(data):
                    self.write("Taux", row, col, val, "date" if col == 2 else "rate" if col == 3 else "text")
                self.fxrefs[(snap["period"], rate["currency"])] = f"'Taux'!D{row+1}"
                self.fxvalues[(snap["period"], rate["currency"])] = rate["rate"]
                row += 1
        self.table("Taux", headers, row - 5, "TauxChange")
        sh = self.sheets["Taux"]
        sh.write(4, 7, "Trimestres disponibles", self.formats["header"])
        sh.set_column(7, 7, 25)
        for index, period in enumerate(self.periods, 5):
            sh.write_string(index, 7, period)
        self.book.define_name("Periodes", f"='Taux'!$H$6:$H${5+len(self.periods)}")

    def rate_formula(self, row):
        terms = f'{self.rng("Taux", "A")},A{row},{self.rng("Taux", "B")},D{row}'
        return f'=IF(COUNTIFS({terms})=1,SUMIFS({self.rng("Taux", "D")},{terms}),"{NA}")'

    def stock(self):
        headers = ["Trimestre", "Code pays", "Pays", "Devise", "Millésime", "Stock initial local",
                   "Remboursé client local", "Injoignables local", "Faux positifs local", "Autorités local", "Autres sans remboursement local",
                   "Traité calculé local", "A traiter local", "Stock identifié local", "Traité déclaré local", "A traiter déclaré local",
                   "Taux local pour 1 EUR", "A traiter EUR", "Stock identifié EUR", "Source"]
        self.start("Stock", "Stock par millésime issu de KPI_2", headers,
                   [15, 12, 22, 12, 16] + [21] * 14 + [58],
                   "Catégories absentes du modèle pays : n.d. Les calculs utilisent uniquement les catégories applicables.")
        row = 5
        fieldcols = {key: 6+i for i, key in enumerate(CATEGORIES)}
        for snap in self.snapshots:
            for entity in snap["entities"]:
                rate = self.fxvalues.get((snap["period"], entity["currency"]))
                for record in entity["stock"]:
                    r = row+1
                    data = [snap["period"], entity["id"], entity["name"], entity["currency"], record["vintage"], record["total"]] + [record[k] for k in CATEGORIES]
                    for col, value in enumerate(data):
                        self.write("Stock", row, col, value, "number" if col >= 5 else "text")
                    refs = [f"{colname(fieldcols[k])}{r}" for k in entity["processed_fields"]]
                    processed = complete_sum(record[k] for k in entity["processed_fields"])
                    remaining = subtract(record["total"], processed)
                    addback = "unreachable" in entity["processed_fields"]
                    identified = complete_sum([remaining, record["unreachable"]]) if addback else remaining
                    self.formula("Stock", row, 11, self.arithmetic(refs, f'SUM({",".join(refs)})'), processed)
                    self.formula("Stock", row, 12, self.arithmetic([f"F{r}", f"L{r}"], f"F{r}-L{r}"), remaining)
                    self.formula("Stock", row, 13, self.arithmetic([f"M{r}", f"H{r}"] if addback else [f"M{r}"], f"M{r}+H{r}" if addback else f"M{r}"), identified)
                    self.write("Stock", row, 14, record["processed_source"], "number")
                    self.write("Stock", row, 15, record["remaining_source"], "number")
                    self.formula("Stock", row, 16, self.rate_formula(r), rate, "rate")
                    self.formula("Stock", row, 17, self.arithmetic([f"M{r}", f"Q{r}"], f"M{r}/Q{r}"), divide(remaining, rate))
                    self.formula("Stock", row, 18, self.arithmetic([f"N{r}", f"Q{r}"], f"N{r}/Q{r}"), divide(identified, rate))
                    self.write("Stock", row, 19, record["source"])
                    row += 1
        self.table("Stock", headers, row-5, "StockDetail")

    def flows(self):
        headers = ["Trimestre", "Code pays", "Pays", "Devise", "Mois de cohorte", "Flux local", "Contrats entrants",
                   "Remboursé client local", "Injoignables local", "Faux positifs local", "Autorités local", "Autres sans remboursement local",
                   "Traité calculé local", "Contrats traités", "Reste local", "Contrats restants", "Taux local pour 1 EUR",
                   "Flux EUR", "Traité EUR", "Efficacité montant", "Efficacité nombre", "Traité déclaré local", "Source"]
        self.start("Flux", "Flux par cohorte mensuelle et date de reporting", headers,
                   [15, 12, 22, 12, 20] + [21]*17 + [58],
                   "Une même cohorte peut apparaître dans plusieurs trimestres de reporting. Filtrer le trimestre avant d'agréger.")
        row = 5
        fieldcols = {key: 7+i for i, key in enumerate(CATEGORIES)}
        for snap in self.snapshots:
            for entity in snap["entities"]:
                rate = self.fxvalues.get((snap["period"], entity["currency"]))
                for record in entity["flows"]:
                    r = row+1
                    data = [snap["period"], entity["id"], entity["name"], entity["currency"], datetime.strptime(record["month"], "%Y-%m"), record["amount"], record["count"]] + [record[k] for k in CATEGORIES]
                    for col, value in enumerate(data):
                        self.write("Flux", row, col, value, "month" if col == 4 else "integer" if col == 6 else "number" if col >= 5 else "text")
                    refs = [f"{colname(fieldcols[k])}{r}" for k in entity["processed_fields"]]
                    processed = complete_sum(record[k] for k in entity["processed_fields"])
                    self.formula("Flux", row, 12, self.arithmetic(refs, f'SUM({",".join(refs)})'), processed)
                    self.write("Flux", row, 13, record["processed_count"], "integer")
                    self.formula("Flux", row, 14, self.arithmetic([f"F{r}", f"M{r}"], f"F{r}-M{r}"), subtract(record["amount"], processed))
                    self.formula("Flux", row, 15, self.arithmetic([f"G{r}", f"N{r}"], f"G{r}-N{r}"), subtract(record["count"], record["processed_count"]), "integer")
                    self.formula("Flux", row, 16, self.rate_formula(r), rate, "rate")
                    self.formula("Flux", row, 17, self.arithmetic([f"F{r}", f"Q{r}"], f"F{r}/Q{r}"), divide(record["amount"], rate))
                    self.formula("Flux", row, 18, self.arithmetic([f"M{r}", f"Q{r}"], f"M{r}/Q{r}"), divide(processed, rate))
                    self.formula("Flux", row, 19, f'=IF(OR(COUNT(F{r},M{r})<>2,F{r}=0),"{NA}",M{r}/F{r})', ratio(processed, record["amount"]), "percent")
                    self.formula("Flux", row, 20, f'=IF(OR(COUNT(G{r},N{r})<>2,G{r}=0),"{NA}",N{r}/G{r})', ratio(record["processed_count"], record["count"]), "percent")
                    self.write("Flux", row, 21, record["processed_amount"], "number")
                    self.write("Flux", row, 22, record["source"])
                    row += 1
        self.table("Flux", headers, row-5, "FluxDetail")

    def distribution(self):
        headers = ["Trimestre", "Code pays", "Pays", "Devise", "Répartition", "Catégorie", "Contrats", "Montant local", "Taux local pour 1 EUR", "Montant EUR", "Source"]
        self.start("Repartition stock", "Répartitions de stock et périmètre de la source", headers,
                   [15, 12, 22, 12, 18, 25, 18, 21, 24, 21, 58],
                   "Tranches et millésimes décrivent le même stock. Ne pas additionner les deux répartitions. Libellés de tranche conservés.")
        row = 5
        for snap in self.snapshots:
            for entity in snap["entities"]:
                rate = self.fxvalues.get((snap["period"], entity["currency"]))
                for field, kind, label in (("stock_bins", "Tranche", "band"), ("stock_vintages", "Millésime", "vintage")):
                    for record in entity[field]:
                        r = row+1
                        data = [snap["period"], entity["id"], entity["name"], entity["currency"], kind, record[label], record["count"], record["amount"]]
                        for col, value in enumerate(data):
                            self.write("Repartition stock", row, col, value, "integer" if col == 6 else "number" if col == 7 else "text")
                        self.formula("Repartition stock", row, 8, self.rate_formula(r), rate, "rate")
                        self.formula("Repartition stock", row, 9, self.arithmetic([f"H{r}", f"I{r}"], f"H{r}/I{r}"), divide(record["amount"], rate))
                        self.write("Repartition stock", row, 10, record["source"])
                        row += 1
        self.table("Repartition stock", headers, row-5, "StockRepartition")

    def lookup(self, metric_col, period_expr, id_expr):
        amount = self.rng("Consolidation", metric_col)
        criteria = f'{self.rng("Consolidation","A")},{period_expr},{self.rng("Consolidation","B")},{id_expr}'
        return f'=IF(OR(COUNTIFS({criteria})<>1,COUNTIFS({criteria},{amount},"{NA}")>0),"{NA}",SUMIFS({amount},{criteria}))'

