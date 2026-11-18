"""Shared command-line wiring for the four independently runnable steps."""
import argparse
import hashlib
import json
import sys
from pathlib import Path
from uuid import uuid4
from .common import parse_period,read_json,atomic_json
from .extract import extract_folder
from .validate import validate_snapshot
from .history import update_history

ROOT=Path(__file__).resolve().parents[1]

def main(stage='all'):
    parser=argparse.ArgumentParser(description=f'Positive Balance: {stage}')
    parser.add_argument('--input',type=Path,help='Quarterly folder, e.g. inputs/Q4_2026')
    parser.add_argument('--period',help='2026Q4 or Q4_2026; inferred from input folder when omitted')
    parser.add_argument('--config-dir',type=Path,default=ROOT/'config')
    parser.add_argument('--database',type=Path,default=ROOT/'database'/'history.json')
    parser.add_argument('--work-dir',type=Path,default=ROOT/'work')
    parser.add_argument('--output-dir',type=Path,default=ROOT/'reports')
    parser.add_argument('--allow-partial',action='store_true',help='Explicitly accept missing countries; never carry them forward')
    parser.add_argument('--replace-historical',action='store_true',help='Allow replacing a quarter older than the latest database quarter')
    parser.add_argument('--dry-run',action='store_true',help='Extract and validate only; leave history/reports unchanged')
    args=parser.parse_args()
    try:
        if stage in ('all','extract') and not args.input:raise ValueError('--input is required for extraction.')
        if not args.period and not args.input:raise ValueError('--period is required for this step.')
        period=parse_period(args.period or args.input.name)
        work=args.work_dir/period
        if stage in ('all','extract'):
            configurations={name:read_json(args.config_dir/f'{name}.json') for name in ('countries','layouts')}
            base_path=args.config_dir/'base.json'
            configurations['base']=read_json(base_path) if base_path.exists() else {}
            snapshot=extract_folder(args.input,period,configurations['countries'],configurations['layouts'],configurations['base'])
            config_bytes=json.dumps(configurations,sort_keys=True,ensure_ascii=False).encode('utf-8')
            config_hash=hashlib.sha256(config_bytes).hexdigest()
            snapshot['configuration_sha256']=config_hash
            atomic_json(work/'config_snapshots'/f'{config_hash}.json',configurations)
            atomic_json(work/'extracted.json',snapshot)
            print(f'1. Extracted {len(snapshot["entities"])} countries -> {work / "extracted.json"}')
        if stage in ('all','validate'):
            override=args.config_dir/'overrides'/f'{period}.json'
            snapshot,report=validate_snapshot(read_json(work/'extracted.json'),read_json(args.config_dir/'countries.json'),read_json(override) if override.exists() else {},args.allow_partial)
            atomic_json(work/'validation.json',report)
            if not report['valid']:
                # A failed retry must not leave an older validated snapshot available.
                (work/'validated.json').unlink(missing_ok=True)
                lines='; '.join(i['message'] for i in report['errors'][:5])
                raise ValueError(f'Validation stopped: {lines} See {work / "validation.json"}. History unchanged.')
            atomic_json(work/'validated.json',snapshot)
            print(f'2. Validated; {len(report["warnings"])} findings -> {work / "validation.json"}')
        if args.dry_run:
            print('Dry run finished. History and Excel reports were not changed.');return 0
        if stage in ('all','history'):
            result=update_history(args.database,read_json(work/'validated.json'),read_json(work/'validation.json'),replace_historical=args.replace_historical)
            atomic_json(work/'history_update.json',result)
            print(f'3. History {result["status"]}: {period}')
        if stage in ('all','report'):
            from .reporting import generate_report
            args.output_dir.mkdir(parents=True,exist_ok=True)
            output=args.output_dir/f'Synthese_detaillee_{period}.xlsx'
            temporary=args.output_dir/f'.{output.stem}.{uuid4().hex}.xlsx'
            try:
                generate_report(read_json(args.database),period,temporary,read_json(args.config_dir/'presentation.json'))
                temporary.replace(output)
            finally:temporary.unlink(missing_ok=True)
            print(f'4. Report ready: {output}')
        return 0
    except (ValueError,KeyError,FileNotFoundError,PermissionError) as exc:
        print(f'ERROR: {exc}',file=sys.stderr);return 1
