# Updating the publication lists

Edit `dist/content.json`. It contains five named sections, each with an `items` list. Each item is one publication or reference—equivalent to one row in a table. These are data lists, not visible table layouts or a database.

| Page section | Section ID in the data file |
| --- | --- |
| Economics papers | `economics-papers` |
| Working papers | `working-papers` |
| Other papers | `other-papers` |
| Technical reports | `technical-reports` |
| Media & contact | `media` |

## Add a publication

1. Find the section by its `id` and copy an item inside its `items` list.
2. Replace the title, coauthors, year, publication details, and optional supporting-file links.
3. Place the record where it should appear. Record order is display order; dates never re-sort the list. Keep the student-teacher value-added paper first until the author requests otherwise.
4. If an existing working paper becomes published, move its record to the publication section and update the citation instead of leaving a duplicate.
5. For a downloadable published paper, add the PDF under `dist/papers/` and use `papers/filename.pdf`, or use an existing public URL. Omit the link if no download should be offered. No working-paper download links are allowed.
6. Build and check the page, then publish through the appropriate hosting workflow below.

Example record (placeholder text only; not part of the site):

```json
{
  "title": "New publication title",
  "year": "2027",
  "authors": "with Coauthor One and Coauthor Two.",
  "publication": "Journal name, Vol. 1, Issue 2, pp. 3–10. (2027)",
  "links": [
    {"label": "pdf", "url": "papers/new-publication.pdf"}
  ],
  "prior": ""
}
```

Only `title` is required; the other fields may be omitted or left empty. `publication` is the displayed citation line. Existing records use `<span class="journal">Journal name</span>` to italicize the journal; copying that formatting is optional. Keep any old-title note in `prior`. Use `links: []` for working papers. A media record has its citation in `title` and a link labeled `link`.

The unused legacy `meta` copies were removed so editing the displayed fields is sufficient. The `label` on each section identifies it for people editing the file; section headings and navigation wording live in the layout templates.

## Build and preview

With Python 3.10+ installed, run from the project folder:

```sh
python3 scripts/render-content.py
python3 -m unittest discover -s tests
python3 scripts/render-content.py --check
```

Equivalent shortcuts are `npm run build`, `npm test`, and `npm run check`. There are no third-party dependencies to install. During local editing, `npm run watch` rebuilds the HTML after each saved change to the data or templates. Refresh the preview to see the result.

The build generates both `dist/proposed.html` and the comparison workspace's `dist/index.html` from `templates/`. It updates all section counts and the total automatically. It never edits `dist/current.html`. Bad JSON, missing sections, or working-paper download links fail the build before either generated page is changed.

## GitHub publishing setup
Keep existing PDFs, sheets and supplementary pages in place; the public build carries them forward at their current addresses. Superseded CVs and the old homepage are kept in archive/, which the public build does not publish; do not move them back into papers/ or the top level. To update the CV, upload the new version to papers/ as TomAhn-CV.pdf, replacing the existing file, and save a dated copy (for example TomAhn-CV_Jan27.pdf) in archive/.

After setup, edit dist/content.json in GitHub and commit. Upload published PDFs to papers/ (or dist/papers/) and use their paths in the record. The workflow rebuilds the entries and counts automatically, then checks same-site links before publishing. Working-paper download links remain prohibited.
