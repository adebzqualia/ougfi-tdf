"""Small shared helpers; no Excel or reporting rules in this module."""
import calendar
import hashlib
import json
import os
import re
import tempfile
from datetime import date
from pathlib import Path

def parse_period(value):
    text=str(value).strip().upper()
    match=re.fullmatch(r'(20\d{2})Q([1-4])',text)
    if match:return text
    match=re.fullmatch(r'Q([1-4])[_ -]?(20\d{2})',text)
    if match:return f'{match[2]}Q{match[1]}'
    raise ValueError(f'Invalid period {value!r}; use Q4_2026 or 2026Q4.')

def quarter_end(period):
    period=parse_period(period);year=int(period[:4]);month=int(period[-1])*3
    return date(year,month,calendar.monthrange(year,month)[1])

def rolling_months(period,count=6):
    end=quarter_end(period);serial=end.year*12+end.month-1
    return [f'{(serial-i)//12:04d}-{(serial-i)%12+1:02d}' for i in reversed(range(count))]

def read_json(path):
    def reject(value):raise ValueError(f'Invalid JSON numeric constant {value} in {path}')
    with Path(path).open(encoding='utf-8-sig') as stream:return json.load(stream,parse_constant=reject)

def atomic_json(path,value):
    """Replace one file atomically on the same filesystem, never truncate it first."""
    path=Path(path);path.parent.mkdir(parents=True,exist_ok=True)
    content=json.dumps(value,ensure_ascii=False,indent=2,allow_nan=False)+'\n'
    fd,tmp=tempfile.mkstemp(prefix=path.name+'.',suffix='.tmp',dir=path.parent)
    try:
        with os.fdopen(fd,'w',encoding='utf-8') as stream:
            stream.write(content);stream.flush();os.fsync(stream.fileno())
        os.replace(tmp,path)
    finally:
        if os.path.exists(tmp):os.unlink(tmp)

def snapshot_hash(snapshot):
    return hashlib.sha256(json.dumps(snapshot,sort_keys=True,ensure_ascii=False,separators=(',',':'),allow_nan=False).encode('utf-8')).hexdigest()
