# releasenotestera

Private plugin. Give it a list of completed task names, a version number and a date. Get
back `Release Notes MM-DD-YYYY.docx` in the Regal Wealth Advisors house template with only
the contents changed.

```
/releasenotestera 1.0.15, 08/22/2026
Complete Admin Dashboard Page.
Update the Login button styling for Dark Mode.
Fix Roth calculator filing status payload.
```

Claude sorts each line into **New Features**, **Improvements** or **Bug Fixes**, writes a
title, one-line summary, 2–4 **Details** bullets and a **Benefit** line for each, drops a
`Screenshot: <title>` placeholder box underneath, and stamps the version, date and Highlights
counts. You click each placeholder in Word and paste the real screenshot over it.

## Install

```
/plugin marketplace add "C:/Users/stane/Downloads/release notes/releasenotestera-plugin"
/plugin install releasenotestera@releasenotestera
```

Requires `python-docx` and `Pillow`:

```
pip install python-docx Pillow
```

## What is fixed vs. what changes

Fixed, every release — comes from `skills/releasenotestera/template.docx`:

- Header table, intro sentence, section headings, thank-you line, Regal Wealth Advisors
  branding and logo boxes.
- Century Gothic body text, teal `#31849B` New Features numbering, blue `#2E75B5` on
  `#F2F9FF` for Improvements and Bug Fixes, cell borders and column widths.
- The `Title / Summary / Details / Benefit / screenshot` shape of every item.

Changes every release:

- Version, release date, Highlights counts.
- The items in the three tables.

Sections with no items are removed entirely — heading, table and Highlights term.

## Direct use without Claude

```bash
python skills/releasenotestera/build_release_notes.py items.json -o "Release Notes 08-22-2026.docx"
python skills/releasenotestera/build_release_notes.py --selfcheck
```

`items.json`:

```json
{
  "version": "1.0.15",
  "date": "08/22/2026",
  "new_features": [
    {
      "title": "Expenses – New Screen",
      "summary": "Introduced a redesigned Expenses screen with an improved layout and user experience.",
      "details": [
        "Designed and implemented a new Expenses screen.",
        "Enhanced page structure and data organization."
      ],
      "benefit": "Provides a modern and intuitive experience for managing client expenses."
    }
  ],
  "improvements": [],
  "bug_fixes": []
}
```

`version` and `date` are required. `release_title` is optional and defaults to
`Regal Wealth Advisors`. See `example.json`.

## Changing the template

Edit `skills/releasenotestera/template.docx` in Word. Keep each of the three item tables at
exactly **one** row, and keep that row's second cell at exactly **12 paragraphs**:

```
title, summary, blank, "Details", bullet, bullet, bullet,
blank, "Benefit", blank, <image>, blank
```

The builder clones that row per item, so whatever you style there is what every item gets.
Run `--selfcheck` afterwards.
