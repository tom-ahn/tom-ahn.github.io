"""Rebuild static pages after a saved data or template change; no web server."""

from pathlib import Path
import runpy
import sys
import time

ROOT = Path(__file__).resolve().parents[1]
GENERATOR = ROOT / 'scripts/render-content.py'


def signature():
    paths = [ROOT / 'dist/content.json', GENERATOR, *sorted((ROOT / 'templates').glob('*.html'))]
    return tuple((str(path), path.stat().st_mtime_ns, path.stat().st_size) for path in paths)


def rebuild():
    generator = runpy.run_path(str(GENERATOR))
    try:
        counts = generator['build'](ROOT / 'dist/content.json', ROOT / 'templates', ROOT / 'dist')
        print(f'Rebuilt {sum(counts.values())} entries; section counts updated.', flush=True)
    except (OSError, ValueError) as error:
        print(f'Build failed; previous pages retained: {error}', file=sys.stderr, flush=True)


def main():
    previous = signature()
    rebuild()
    print('Watching content.json and templates. Press Ctrl+C to stop.', flush=True)
    try:
        while True:
            time.sleep(0.5)
            current = signature()
            if current != previous:
                previous = current
                rebuild()
    except KeyboardInterrupt:
        print('\nStopped watching.', flush=True)


if __name__ == '__main__':
    main()
