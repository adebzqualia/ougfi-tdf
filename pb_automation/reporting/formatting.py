"""Display-only rounding for slide commentary; source values stay unchanged."""
from decimal import Decimal, ROUND_HALF_UP
import math


def rounded(value, places=0):
    return Decimal(str(value)).quantize(Decimal('1').scaleb(-places), rounding=ROUND_HALF_UP)


def slide_amount(value, currency, per_contract=False):
    if value is None:
        return f'0 {currency}'
    if per_contract:
        amount = rounded(value, -1 if abs(value) >= 10 else 0)
        return f'{amount:,.0f} {currency}'
    if abs(value) >= 1_000_000:
        return f'{rounded(value / 1_000_000, 1):.1f}m {currency}'
    if abs(value) >= 1_000:
        return f'{rounded(value / 1_000):.0f}k {currency}'
    return f'{rounded(value):,.0f} {currency}'


def slide_amount_formula(ref, currency, per_contract=False):
    if per_contract:
        expression=f'TEXT(ROUND({ref},IF(ABS({ref})>=10,-1,0)),"#,##0")&" {currency}"'
    else:
        expression=(f'IF(ABS({ref})>=1000000,TEXT({ref}/1000000,"0.0")&"m {currency}",'
                    f'IF(ABS({ref})>=1000,TEXT({ref}/1000,"0")&"k {currency}",'
                    f'TEXT({ref},"#,##0")&" {currency}"))')
    return f'IF(ISNUMBER({ref}),{expression},"0 {currency}")'


def slide_percent(value):
    return '0%' if value is None else f'{rounded(value * 100):.0f}%'


def stock_axis(rows):
    """One readable zero-based scale covering both stacked stock charts."""
    heights = [sum(max(v.get(key) or 0, 0) for v in row['vintages'])
               for row in rows for key in ('identified_eur', 'remaining_eur')]
    peak=max(heights, default=0)
    if peak <= 0:
        return {'min':0, 'max':1, 'major_unit':.2}
    raw=peak/7
    power=10**math.floor(math.log10(raw))
    step=next(n*power for n in (1,2,2.5,5,10) if n*power>=raw)
    return {'min':0, 'max':math.ceil(peak*1.05/step)*step, 'major_unit':step}
