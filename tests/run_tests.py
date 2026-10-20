#!/usr/bin/env python3
"""Run the synthetic matrix; intermediates stay in an ignored work directory."""
from pathlib import Path
import argparse
import json
import os
import subprocess
import sys

HERE = Path(__file__).resolve().parent
REPO = HERE.parent


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--work-dir', type=Path, default=HERE / '_work')
    parser.add_argument('--engine', default=os.environ.get(
        'LO_ENGINE', str(Path.home() / '.cache/codex-runtimes/codex-primary-runtime/dependencies/native/libreoffice-headless/libreoffice/LibreOfficeDev.app/Contents/MacOS/soffice')))
    args = parser.parse_args()
    engine = Path(args.engine).expanduser().resolve()
    if not engine.is_file():
        parser.error('LibreOffice engine not found; pass --engine or set LO_ENGINE.')
    work = args.work_dir.expanduser().resolve()
    work.mkdir(parents=True, exist_ok=True)
    fixtures = work / 'fixtures'
    output = work / 'converted'
    results = work / 'general-test-results.ndjson'
    report = work / 'general-test-audit.json'

    subprocess.run([sys.executable, str(HERE / 'test_batch_inputs.py')], check=True)
    subprocess.run([sys.executable, str(HERE / 'test_layout_preservation.py')], check=True)
    subprocess.run([sys.executable, str(HERE / 'build_fixtures.py'),
                    '--engine', str(engine), '--output-dir', str(fixtures)], check=True)
    subprocess.run([sys.executable, str(HERE / 'audit_roundtrips.py'),
                    '--engine', str(engine), '--fixtures', str(fixtures)], check=True)
    manifest = json.loads((fixtures / 'expectations.json').read_text())
    inputs = [str(fixtures / item['filename']) for item in manifest['fixtures']]
    with results.open('w', encoding='utf-8') as stream:
        # The matrix deliberately includes invalid files, so conversion exits 1.
        converted = subprocess.run([
            sys.executable, str(REPO / 'engine/converter.py'), '--engine', str(engine),
            '--output-dir', str(output), *inputs], stdout=stream)
    if converted.returncode not in (0, 1):
        raise SystemExit(converted.returncode)
    subprocess.run([sys.executable, str(HERE / 'audit_general_results.py'),
                    '--fixtures', str(fixtures), '--results', str(results),
                    '--report', str(report)], check=True)
    subprocess.run([sys.executable, str(HERE / 'test_body_regressions.py'),
                    '--engine', str(engine), '--work-dir', str(work / 'body-regressions')], check=True)
    subprocess.run([sys.executable, str(HERE / 'test_header_regressions.py'),
                    '--engine', str(engine), '--work-dir', str(work / 'header-regressions')], check=True)
    print(f'Full audit: {report}')


if __name__ == '__main__':
    main()

