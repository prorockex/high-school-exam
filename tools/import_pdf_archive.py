#!/usr/bin/env python3
"""Add a locally supplied PDF directory to the app; no network requests.
Use files you downloaded yourself or an archive supplied with bulk-use permission.
"""
import argparse, hashlib, json, re, shutil
from datetime import datetime
from pathlib import Path
from urllib.parse import urlsplit
from zoneinfo import ZoneInfo


def file_hash(path):
    digest = hashlib.sha256()
    with path.open('rb') as handle:
        while chunk := handle.read(65536): digest.update(chunk)
    return digest.hexdigest()


def save_catalog(root, catalog):
    payload = json.dumps(catalog, ensure_ascii=False, indent=2) + '\n'
    path = root / 'pdf-catalog.json'
    temporary = path.with_suffix('.json.tmp'); temporary.write_text(payload); temporary.replace(path)
    index = root / 'index.html'
    if index.exists():
        pattern = r'(<script id="pdf-catalog" type="application/json">).*?(</script>)'
        encoded = json.dumps(catalog, ensure_ascii=False).replace('<', '\\u003c')
        updated, count = re.subn(pattern, lambda m: m[1] + encoded + m[2], index.read_text(), flags=re.S)
        if count != 1: raise ValueError('index.html must have exactly one pdf-catalog script')
        index.write_text(updated)


def import_archive(directory, root, source='', school=''):
    directory = directory.resolve(); root = root.resolve()
    if not directory.is_dir(): raise ValueError('Input must be an existing PDF directory')
    if source:
        parsed = urlsplit(source)
        if parsed.scheme != 'https' or not parsed.hostname or parsed.username or parsed.password:
            raise ValueError('Source must be an HTTPS URL without credentials')
    if len(school) > 100: raise ValueError('School name too long')
    root.mkdir(parents=True, exist_ok=True)
    destination = root / 'pdfs'; destination.mkdir(exist_ok=True)
    catalog_file = root / 'pdf-catalog.json'
    catalog = json.loads(catalog_file.read_text()) if catalog_file.exists() else {'papers': []}
    existing = {p['id']: p for p in catalog.get('papers', [])}
    failed = []; imported = 0; duplicates = 0
    paths = sorted(directory.rglob('*'))
    for path in paths:
        if not path.is_file() or path.suffix.lower() != '.pdf': continue
        if not path.resolve().is_relative_to(directory):
            failed.append({'name': path.name, 'error': 'Symlink escapes input directory'}); continue
        try:
            digest = hashlib.sha256()
            with path.open('rb') as handle:
                if handle.read(5) != b'%PDF-': raise ValueError('Not a PDF')
                handle.seek(0, 2); size = handle.tell(); handle.seek(max(0, size - 2048))
                if b'%%EOF' not in handle.read(): raise ValueError('Missing PDF EOF marker')
                handle.seek(0)
                while chunk := handle.read(65536): digest.update(chunk)
            identifier = digest.hexdigest(); target = destination / (identifier + '.pdf')
            if identifier in existing and target.exists() and file_hash(target) == identifier:
                duplicates += 1; continue
            if shutil.disk_usage(destination).free < size + 64 * 1024 * 1024:
                raise OSError('Insufficient disk space')
            temporary = destination / (identifier + '.part')
            try:
                shutil.copyfile(path, temporary)
                if file_hash(temporary) != identifier:
                    raise ValueError('Copied file checksum mismatch')
                temporary.replace(target)
            finally: temporary.unlink(missing_ok=True)
            existing[identifier] = {'id': identifier, 'name': path.name, 'path': 'pdfs/' + identifier + '.pdf',
                                    'bytes': size, 'source': source, 'school': school}
            imported += 1
        except (ValueError, OSError) as error: failed.append({'name': path.name, 'error': str(error)})
    catalog.update({'status': 'local_import', 'updatedAt': datetime.now(ZoneInfo('Asia/Taipei')).isoformat(),
                    'papers': list(existing.values()), 'importFailures': failed})
    save_catalog(root, catalog)
    result = {'imported': imported, 'duplicates': duplicates, 'failed': len(failed), 'total': len(existing)}
    print(json.dumps(result, ensure_ascii=False))
    return result


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--input', type=Path, required=True, help='Directory containing downloaded PDFs')
    parser.add_argument('--output', type=Path, default=Path(__file__).resolve().parents[1])
    parser.add_argument('--source', default='', help='Optional original-source HTTPS URL')
    parser.add_argument('--school', default='', help='Optional school shared by this batch; blank means unclassified')
    args = parser.parse_args()
    try: result = import_archive(args.input, args.output, args.source, args.school)
    except (ValueError, OSError) as error: parser.error(str(error))
    raise SystemExit(2 if result['failed'] else 0)
