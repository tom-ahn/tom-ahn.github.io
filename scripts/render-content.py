"""Build static pages from the section records in dist/content.json."""

import argparse
from datetime import date
import html
import json
from pathlib import Path
import re
import sys
from urllib.parse import urlsplit

ROOT = Path(__file__).resolve().parents[1]
SECTION_IDS = ('economics-papers', 'working-papers', 'other-papers', 'technical-reports', 'media')
TOKEN = re.compile(r'\{\{(count|entries):([a-z-]+)\}\}|\{\{total\}\}')


def load_sections(path):
    data = json.loads(Path(path).read_text(encoding='utf-8'))
    if not isinstance(data, list):
        raise ValueError('The content file must contain a list of sections.')
    ids = [section.get('id') for section in data if isinstance(section, dict)]
    if len(ids) != len(data) or len(set(ids)) != len(ids) or set(ids) != set(SECTION_IDS):
        raise ValueError('Include each of the five section IDs exactly once.')
    for section in data:
        sid = section['id']
        if not isinstance(section.get('items'), list):
            raise ValueError(f'{sid}: items must be a list.')
        for number, item in enumerate(section['items'], 1):
            where = f'{sid}, entry {number}'
            if not isinstance(item, dict) or not isinstance(item.get('title'), str) or not item['title'].strip():
                raise ValueError(f'{where}: a nonempty title is required.')
            for key in ('year', 'authors', 'publication', 'prior'):
                if not isinstance(item.get(key, ''), str):
                    raise ValueError(f'{where}: {key} must be text.')
            links = item.get('links', [])
            if not isinstance(links, list):
                raise ValueError(f'{where}: links must be a list.')
            if sid == 'working-papers' and links:
                raise ValueError(f'{where}: working-paper drafts must not have download links.')
            for link in links:
                if not isinstance(link, dict) or any(not isinstance(link.get(k), str) or not link[k].strip() for k in ('label', 'url')):
                    raise ValueError(f'{where}: each link needs a label and URL.')
                if urlsplit(link['url']).scheme not in ('', 'http', 'https'):
                    raise ValueError(f'{where}: use an https URL or a relative file path.')
    return data


def entry_links(links):
    if not links:
        return ''
    anchors = []
    for link in links:
        label = link['label']
        label = 'PDF' if label.lower() == 'pdf' else label[0].upper() + label[1:]
        anchors.append('<a href="' + html.escape(link['url'], quote=True) +
                       '" target="_blank" rel="noopener noreferrer">' +
                       html.escape(label) + '</a>')
    return '<div class="paper-links">' + ''.join(anchors) + '</div>'


def render_entry(item, media=False):
    title = html.escape(item['title'])
    links = entry_links(item.get('links', []))
    if media:
        return '<li class="entry media"><p class="media-item">' + title + '</p>' + links + '</li>'
    body = '<div class="entry-year">' + html.escape(item.get('year', '')) + '</div><div class="entry-body"><h3>' + title + '</h3>'
    if item.get('authors'):
        body += '<p class="authors">' + html.escape(item['authors']) + '</p>'
    if item.get('publication'):
        body += '<p class="publication">' + item['publication'] + '</p>'
    if item.get('prior'):
        body += '<p class="prior">Previously circulated as: ' + html.escape(item['prior']) + '</p>'
    return '<li class="entry paper">' + body + links + '</div></li>'


def render_template(template, sections):
    records = {section['id']: section['items'] for section in sections}

    def replace(match):
        if match.group(0) == '{{total}}':
            return str(sum(len(items) for items in records.values()))
        kind, sid = match.groups()
        if sid not in records:
            raise ValueError(f'Template references an unknown section: {sid}')
        items = records[sid]
        if kind == 'count':
            return str(len(items))
        if sid == 'working-papers':
            # Working papers are actively maintained: a blank year shows the year the page was built.
            current_year = str(date.today().year)
            items = [dict(item, year=item.get('year', '').strip() or current_year) for item in items]
        return ''.join(render_entry(item, sid == 'media') for item in items)

    result = TOKEN.sub(replace, template)
    if '{{' in result or '}}' in result:
        raise ValueError('The template contains an unresolved placeholder.')
    return result


def build(data_path, template_dir, output_dir, check=False):
    sections = load_sections(data_path)
    pages = {name: render_template((Path(template_dir) / name).read_text(encoding='utf-8'), sections)
             for name in ('index.html', 'proposed.html')}
    # Render every page before writing; invalid input leaves the previous pages intact.
    output_dir = Path(output_dir)
    if check:
        for name, content in pages.items():
            destination = output_dir / name
            if not destination.exists() or destination.read_text(encoding='utf-8') != content:
                raise ValueError(f'{destination} needs rebuilding; run npm run build.')
    else:
        output_dir.mkdir(parents=True, exist_ok=True)
        for name, content in pages.items():
            destination = output_dir / name
            if not destination.exists() or destination.read_text(encoding='utf-8') != content:
                temporary = destination.with_suffix('.html.tmp')
                temporary.write_text(content, encoding='utf-8')
                temporary.replace(destination)
    return {section['id']: len(section['items']) for section in sections}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--data', type=Path, default=ROOT / 'dist/content.json')
    parser.add_argument('--templates', type=Path, default=ROOT / 'templates')
    parser.add_argument('--output-dir', type=Path, default=ROOT / 'dist')
    parser.add_argument('--check', action='store_true', help='Check generated pages match the data.')
    args = parser.parse_args()
    try:
        counts = build(args.data, args.templates, args.output_dir, args.check)
    except (OSError, ValueError) as error:
        print(f'Build failed: {error}', file=sys.stderr)
        return 1
    print(('Checked' if args.check else 'Built') + ' ' + str(sum(counts.values())) +
          ' entries: ' + ', '.join(f'{sid}={n}' for sid, n in counts.items()))
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
