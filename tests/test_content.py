import copy
import html
import json
from pathlib import Path
import re
import runpy
import tempfile
import unittest

ROOT = Path(__file__).resolve().parents[1]
GENERATOR = runpy.run_path(str(ROOT / 'scripts/render-content.py'))
PUBLIC = runpy.run_path(str(ROOT / 'scripts/build-github-pages.py'))


class ContentBuildTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.data_path = self.root / 'content.json'
        self.output = self.root / 'dist'
        self.sections = json.loads((ROOT / 'dist/content.json').read_text())

    def build(self, sections=None, check=False):
        data = self.sections if sections is None else sections
        self.data_path.write_text(json.dumps(data))
        return GENERATOR['build'](self.data_path, ROOT / 'templates', self.output, check)

    def counts(self, page):
        navigation = dict(re.findall(r'href="#([a-z-]+)"><span>[^<]+</span><span class="nav-count">(\d+)</span>', page))
        headings = dict(re.findall(r'<section id="([a-z-]+)".*?<span class="section-count">(\d+)</span>', page))
        return navigation, headings

    def test_each_section_accepts_a_new_record_and_updates_both_counts(self):
        for sid in GENERATOR['SECTION_IDS']:
            with self.subTest(section=sid):
                data = copy.deepcopy(self.sections)
                target = next(section for section in data if section['id'] == sid)
                # Minimal records are supported; optional citation fields may be omitted.
                target['items'].append({'title': 'Temporary <verification> & example'})
                counts = self.build(data)
                page = (self.output / 'proposed.html').read_text()
                navigation, headings = self.counts(page)
                expected = str(len(target['items']))
                self.assertEqual(navigation[sid], expected)
                self.assertEqual(headings[sid], expected)
                self.assertIn('Temporary &lt;verification&gt; &amp; example', page)
                self.assertIn(f'All {sum(counts.values())} entries', (self.output / 'index.html').read_text())

    def test_deleting_all_records_from_a_section_sets_both_counts_to_zero(self):
        self.sections[0]['items'] = []
        self.build()
        navigation, headings = self.counts((self.output / 'proposed.html').read_text())
        self.assertEqual(navigation['economics-papers'], '0')
        self.assertEqual(headings['economics-papers'], '0')

    def test_publishing_a_working_paper_moves_its_record_and_preserves_chosen_order(self):
        economics = next(section for section in self.sections if section['id'] == 'economics-papers')['items']
        working = next(section for section in self.sections if section['id'] == 'working-papers')['items']
        first_title = economics[0]['title']
        moved = working.pop(0)
        moved['links'] = [{'label': 'pdf', 'url': 'papers/new-publication.pdf'}]
        economics.insert(1, moved)
        counts = self.build()
        page = (self.output / 'proposed.html').read_text()
        titles = re.findall(r'<h3>(.*?)</h3>', page)
        self.assertEqual(titles[0], first_title)
        self.assertEqual(titles[1], moved['title'])
        self.assertEqual(counts['economics-papers'], len(economics))
        self.assertEqual(counts['working-papers'], len(working))
        self.assertIn('href="papers/new-publication.pdf"', page)
        self.assertFalse(re.search(r'<h3>.*?<a', page.split('</h3>')[0]))

    def test_working_paper_download_links_are_rejected_without_changing_pages(self):
        self.build()
        previous = (self.output / 'proposed.html').read_bytes()
        working = next(section for section in self.sections if section['id'] == 'working-papers')
        working['items'][0]['links'] = [{'label': 'pdf', 'url': 'papers/draft.pdf'}]
        with self.assertRaisesRegex(ValueError, 'must not have download links'):
            self.build()
        self.assertEqual((self.output / 'proposed.html').read_bytes(), previous)

    def test_invalid_json_or_missing_sections_leave_previous_pages_intact(self):
        self.build()
        previous = {name: (self.output / name).read_bytes() for name in ('index.html', 'proposed.html')}
        self.data_path.write_text('{')
        with self.assertRaises(ValueError):
            GENERATOR['build'](self.data_path, ROOT / 'templates', self.output)
        with self.assertRaisesRegex(ValueError, 'five section IDs'):
            self.build(self.sections[:-1])
        for name, content in previous.items():
            self.assertEqual((self.output / name).read_bytes(), content)

    def test_check_mode_detects_stale_output(self):
        self.build()
        self.build(check=True)
        self.sections[0]['items'].append({'title': 'New record'})
        with self.assertRaisesRegex(ValueError, 'needs rebuilding'):
            self.build(check=True)
        self.build()
        self.build(check=True)

    def test_media_titles_stay_plain_with_separate_link_text(self):
        self.build()
        page = (self.output / 'proposed.html').read_text()
        items = re.findall(r'<li class="entry media">(.*?)</li>', page)
        expected = next(section for section in self.sections if section['id'] == 'media')['items']
        self.assertEqual(len(items), len(expected))
        for item in items:
            self.assertRegex(item, r'<p class="media-item">[^<]*</p>')
            self.assertIn('>Link</a>', item)

    def test_public_build_contains_academic_page_without_prototype_controls(self):
        self.build()
        for name in ('proposed.css', 'proposed.js'):
            (self.output / name).write_bytes((ROOT / 'dist' / name).read_bytes())
        public_dir = self.root / 'public'
        PUBLIC['build_public'](self.output, public_dir)
        page = (public_dir / 'index.html').read_text()
        self.assertIn('<title>Tom Ahn</title>', page)
        self.assertNotIn('noindex', page)
        self.assertNotIn('study-toolbar', page)
        self.assertNotIn('bridge.js', page)
        self.assertEqual(len(re.findall('class="entry ', page)), sum(len(s['items']) for s in self.sections))
        self.assertTrue((public_dir / 'proposed.css').exists())
        self.assertTrue((public_dir / '.nojekyll').exists())
        # Merely building the public artifact must not alter the private comparison workspace.
        self.assertIn('Website layout study', (self.output / 'index.html').read_text())

    def test_public_build_preserves_existing_files_and_checks_local_links(self):
        self.build()
        for name in ('proposed.css', 'proposed.js'):
            (self.output / name).write_bytes((ROOT / 'dist' / name).read_bytes())
        repository = self.root / 'repository'
        repository.mkdir()
        (repository / 'index.html').write_text('Existing page to be replaced')
        (repository / 'index_old.html').write_text('Archived old homepage')
        (repository / 'spreadsheets.html').write_text('Supplementary data')
        (repository / 'excel_viewer.html').write_text('Existing workbook viewer')
        (repository / 'root-paper.pdf').write_bytes(b'existing-root-pdf')
        (repository / 'papers').mkdir()
        (repository / 'papers/extra.pdf').write_bytes(b'existing-pdf')
        # Supply placeholders for the actual links; this test verifies packaging, not PDF contents.
        page = (self.output / 'proposed.html').read_text()
        from urllib.parse import urlsplit
        for url in re.findall(r'href="([^"]+)"', page):
            parsed = urlsplit(html.unescape(url))
            if parsed.netloc != 'tom-ahn.github.io' or parsed.path == '/':
                continue
            file = repository / parsed.path.lstrip('/')
            file.parent.mkdir(parents=True, exist_ok=True)
            file.write_bytes(b'existing-resource')
        public_dir = self.root / 'published'
        PUBLIC['build_public'](self.output, public_dir, repository, True)
        self.assertEqual((public_dir / 'index_old.html').read_text(), 'Archived old homepage')
        self.assertTrue((public_dir / 'spreadsheets.html').exists())
        self.assertTrue((public_dir / 'excel_viewer.html').exists())
        self.assertEqual((public_dir / 'root-paper.pdf').read_bytes(), b'existing-root-pdf')
        self.assertEqual((public_dir / 'papers/extra.pdf').read_bytes(), b'existing-pdf')
        self.assertIn('<title>Tom Ahn</title>', (public_dir / 'index.html').read_text())

    def test_publication_stops_when_a_local_supporting_file_is_missing(self):
        self.build()
        for name in ('proposed.css', 'proposed.js'):
            (self.output / name).write_bytes((ROOT / 'dist' / name).read_bytes())
        repository = self.root / 'empty-repository'
        repository.mkdir()
        with self.assertRaisesRegex(ValueError, 'existing files are missing'):
            PUBLIC['build_public'](self.output, self.root / 'incomplete-public', repository, True)


if __name__ == '__main__':
    unittest.main()
