"""Prepare the standalone academic page for a future approved GitHub migration."""

import argparse
from html.parser import HTMLParser
from pathlib import Path
import re
import shutil
from urllib.parse import unquote, urlsplit

ROOT = Path(__file__).resolve().parents[1]
PUBLIC_FOLDERS = ('papers', 'sheets', 'assets', 'images', 'img', 'css', 'js')
PUBLIC_EXTENSIONS = {'.html', '.htm', '.pdf', '.png', '.jpg', '.jpeg', '.gif', '.svg', '.webp', '.ico', '.csv', '.xlsx'}


def validate_local_links(page, output):
    class Links(HTMLParser):
        def __init__(self):
            super().__init__()
            self.urls = []

        def handle_starttag(self, tag, attrs):
            attrs = dict(attrs)
            if tag in ('a', 'link') and attrs.get('href'):
                self.urls.append(attrs['href'])
            elif tag in ('script', 'img') and attrs.get('src'):
                self.urls.append(attrs['src'])

    links = Links()
    links.feed(page)
    missing = []
    for url in links.urls:
        parsed = urlsplit(url)
        if parsed.scheme not in ('', 'http', 'https') or parsed.netloc not in ('', 'tom-ahn.github.io'):
            continue
        if not parsed.path:
            continue
        path = unquote(parsed.path).lstrip('/') or 'index.html'
        target = (output / path).resolve()
        if not target.is_relative_to(output) or not target.is_file():
            missing.append(path)
    if missing:
        raise ValueError('Publication stopped: existing files are missing from the output: ' + ', '.join(sorted(set(missing))))


def build_public(source, output, repository_root=None, require_local_links=False):
    source, output = Path(source).resolve(), Path(output).resolve()
    if output == source or output in source.parents or source in output.parents:
        raise ValueError('Use a separate build directory, outside dist.')
    page = (source / 'proposed.html').read_text(encoding='utf-8')
    page = page.replace('<title>Tom Ahn · Proposed layout</title>', '<title>Tom Ahn</title>')
    page = page.replace('<meta name="robots" content="noindex,nofollow">', '')
    page = page.replace('<script src="bridge.js"></script>', '')
    page = re.sub(r'<p>Content from <a href="https://tom-ahn.github.io/".*?</a></p>', '', page)
    output.mkdir(parents=True, exist_ok=True)
    # Retain the existing repository's public resources at their original URLs.
    # This includes the renamed old homepage and standalone supplementary pages.
    if repository_root is not None:
        repository_root = Path(repository_root).resolve()
        if output == repository_root or output in repository_root.parents:
            raise ValueError('Use a separate build directory, outside the repository root.')
        for name in PUBLIC_FOLDERS:
            if (repository_root / name).is_dir():
                shutil.copytree(repository_root / name, output / name, dirs_exist_ok=True)
        for file in repository_root.iterdir():
            if file.is_file() and (file.suffix.lower() in PUBLIC_EXTENSIONS or file.name == 'CNAME'):
                shutil.copyfile(file, output / file.name)
    (output / 'index.html').write_text(page, encoding='utf-8')
    for name in ('proposed.css', 'proposed.js'):
        shutil.copyfile(source / name, output / name)
    # Local supporting files are optional; current published-file links are absolute.
    for name in PUBLIC_FOLDERS:
        destination = output / name
        if (source / name).is_dir():
            shutil.copytree(source / name, destination, dirs_exist_ok=True)
    (output / '.nojekyll').touch()
    if require_local_links:
        validate_local_links(page, output)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--source', type=Path, default=ROOT / 'dist')
    parser.add_argument('--output-dir', type=Path, default=ROOT / 'build/github-pages')
    parser.add_argument('--repository-root', type=Path, default=ROOT)
    parser.add_argument('--require-local-links', action='store_true', help='Stop if an existing same-site link would break.')
    args = parser.parse_args()
    build_public(args.source, args.output_dir, args.repository_root, args.require_local_links)
    print(f'Prepared standalone page in {args.output_dir}; nothing has been published.')


if __name__ == '__main__':
    main()
