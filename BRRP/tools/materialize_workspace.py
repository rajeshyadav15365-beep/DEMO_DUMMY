"""Restore archived files without overwriting differing existing content."""
import argparse
import hashlib
import json
from pathlib import Path
import shutil


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--root', required=True, type=Path)
    args = parser.parse_args()
    archive = Path(__file__).resolve().parents[1]
    repo = archive.parent
    root = args.root.resolve()
    if root == repo or repo in root.parents:
        parser.error('Choose a workspace outside the repository.')

    jobs = []
    for entry in json.loads((archive / 'source-manifest.json').read_text()):
        source = (repo / entry['repository_path']).resolve()
        if repo not in source.parents or digest(source) != entry['sha256']:
            raise RuntimeError('Invalid source path/checksum in manifest')
        for alias in entry['workspace_aliases']:
            jobs.append((source, root / alias))
    for source_folder, destination_folder in [('project-review', 'project-review'),
                                              ('deliverables', 'downloads')]:
        folder = archive / source_folder
        for source in folder.rglob('*'):
            if source.is_file():
                jobs.append((source, root / destination_folder / source.relative_to(folder)))

    # Validate all destinations before writing anything.
    for source, destination in jobs:
        resolved = destination.resolve()
        if root not in resolved.parents:
            raise RuntimeError('Destination escapes selected workspace')
        if destination.exists() and (not destination.is_file() or
                                     digest(destination) != digest(source)):
            raise RuntimeError(f'Refusing to overwrite differing file: {destination}')

    written = 0
    for source, destination in jobs:
        if destination.exists():
            continue
        destination.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source, destination)
        written += 1
    print(json.dumps({'workspace': str(root), 'files_restored': written,
                      'files_checked': len(jobs)}))


if __name__ == '__main__':
    main()
