"""Step 3: upsert a validated quarter, preserving all other original snapshots."""
import os
from datetime import datetime,timezone
from pathlib import Path
from uuid import uuid4
from .common import read_json,atomic_json,snapshot_hash,parse_period

def update_history(database_path,snapshot,validation_report,*,replace_historical=False):
    if not validation_report.get('valid') or validation_report.get('errors'):
        raise ValueError('Validation failed. The history has not been changed.')
    if snapshot_hash(snapshot)!=validation_report.get('snapshot_sha256'):
        raise ValueError('Validated data changed after validation. Run step 2 again.')
    path=Path(database_path);path.parent.mkdir(parents=True,exist_ok=True)
    lock=path.with_suffix(path.suffix+'.lock')
    try:fd=os.open(lock,os.O_CREAT|os.O_EXCL|os.O_WRONLY)
    except FileExistsError:raise ValueError(f'History is locked by another run: {lock}. If no run is active, remove the stale lock.') from None
    try:
        os.write(fd,str(os.getpid()).encode());os.close(fd)
        data=read_json(path) if path.exists() else {'schema_version':1,'snapshots':{}}
        if data.get('schema_version')!=1 or not isinstance(data.get('snapshots'),dict):raise ValueError('Unsupported database schema; expected schema_version 1 and snapshots.')
        period=parse_period(snapshot['period']);old=data['snapshots'].get(period)
        revisions=data.get('revisions',[])
        if not isinstance(revisions,list):raise ValueError('Invalid history revisions; expected a list.')
        previous_revision=max((r.get('revision',0) for r in revisions if r.get('period')==period),default=1 if old is not None else 0)
        if old is not None and snapshot_hash(old)==snapshot_hash(snapshot):
            return {'status':'unchanged','period':period,'revision':previous_revision,'backup':None}
        if data['snapshots'] and period<max(data['snapshots']) and not replace_historical:
            raise ValueError('This changes a quarter older than the latest history. Use --replace-historical explicitly after review.')
        backup=None
        now=datetime.now(timezone.utc)
        if path.exists():
            stamp=now.strftime('%Y%m%dT%H%M%S%fZ')
            backup=path.parent/'backups'/f'{path.stem}_{stamp}_{uuid4().hex[:8]}.json';atomic_json(backup,data)
        status='replaced' if old is not None else 'added'
        revision=previous_revision+1
        # Metadata belongs outside the validated country snapshot, so validation
        # hashes remain stable. Relative backup paths keep the project portable.
        data['revisions']=revisions+[{
            'period':period,'revision':revision,'status':status,
            'saved_at_utc':now.isoformat(),'snapshot_sha256':snapshot_hash(snapshot),
            'previous_snapshot_sha256':snapshot_hash(old) if old is not None else None,
            'backup':backup.relative_to(path.parent).as_posix() if backup else None,
        }]
        data['snapshots'][period]=snapshot
        data['snapshots']=dict(sorted(data['snapshots'].items()))
        atomic_json(path,data)
        return {'status':status,'period':period,'revision':revision,'backup':str(backup) if backup else None}
    finally:
        lock.unlink(missing_ok=True)
