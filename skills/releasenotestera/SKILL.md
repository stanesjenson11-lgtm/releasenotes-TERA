---
name: releasenotestera
description: "Turn a brief list of completed task names into a Regal Wealth Advisors release-notes Word document. Categorises each task into New Features, Improvements or Bug Fixes, expands it into a title, summary, Details bullets and a Benefit line with a screenshot placeholder, and stamps the version and date. Use when the user types /releasenotestera or asks for release notes, a release doc, or a version write-up from a task list."
---

# /releasenotestera

Input: a rough list of task names plus a version number and date.
Output: `Release Notes MM-DD-YYYY.docx` in the house format, ready for screenshots.

The Word formatting lives entirely in `template.docx` (derived from
`Release Notes 07-01-2026.docx`). Never hand-build a `.docx` — always go through
`build_release_notes.py`, or the fonts, colours, cell shading and column widths drift.

## Steps

**1. Collect version and date.** If the user did not give both, ask. Date format is
`MM/DD/YYYY`. If they point at a file (`.docx`, `.txt`, `.md`) instead of pasting, read it.

**2. Categorise every task line** into exactly one bucket:

| Bucket | What belongs there |
|---|---|
| **New Features** | Something that did not exist before — a new screen, page, module, calculator, integration, or capability. Cues: "Complete X Page", "New Screen", "Add <thing that has no predecessor>". |
| **Improvements** | Something that already existed and got better — UI/UX polish, dark-mode styling, PDF enhancements, performance, session handling, validation hardening, or a removal that simplifies a screen. Cues: "Update", "Enhance", "Revamp", "Remove … section", "Apply … across". |
| **Bug Fixes** | A defect corrected. Cues: "Fix", "Resolve", "Incorrect", "not working", retry/error handling, alignment or display errors. |

Group related lines into one item where they are obviously one piece of work (e.g. six
separate "… for Dark Mode" lines become one *Dark Mode* improvement with each line as a
Details bullet). Do not pad — one task can be one item.

When a line is genuinely ambiguous, put it in **Improvements** and say so in your reply so
the user can move it.

**3. Expand each item** into four fields, matching the house voice:

- **`title`** — `<Area> – <Change>`, using an **en dash (–, U+2013)**, not a hyphen or em dash.
  Some titles legitimately have no dash (`Session Management Improvements`). Title Case.
  Examples: `Expenses – New Screen`, `Roth Calculator – Dark Mode Fixes`.
- **`summary`** — one sentence, past tense, starting with a verb like *Introduced / Added /
  Expanded / Enhanced / Improved / Redesigned / Resolved*. Says what changed, from the user's
  side. No trailing detail dump.
- **`details`** — 2 to 4 terse implementation-level bullets, each a full sentence ending in a
  period. This is what was actually built, not why. Up to 7 is fine for a large feature.
- **`benefit`** — one sentence starting *Provides / Ensures / Improves / Delivers*, stating the
  user-facing value. Never restates the summary.

Write from the task line alone; do not invent specifics (screen names, field names, APIs)
that the user did not mention. If a line is too vague to expand honestly, keep the details
generic and flag it in your reply.

**4. Build.** Write the JSON to the scratchpad, then run `build_release_notes.py` — it sits in
**this same directory as this SKILL.md**, so use that directory's path.

```bash
python "<this skill's directory>/build_release_notes.py" items.json \
  -o "Release Notes MM-DD-YYYY.docx"
```

Omitting `-o` writes `Release Notes MM-DD-YYYY.docx` into the current directory. Write it
where the user is working unless they name somewhere else.

**5. Report** the output path and a one-line count per section, plus anything you flagged as
ambiguous or thin.

## JSON shape

```json
{
  "version": "1.0.15",
  "date": "07/15/2026",
  "new_features": [
    {
      "title": "Expenses – New Screen",
      "summary": "Introduced a redesigned Expenses screen with an improved layout and user experience.",
      "details": [
        "Designed and implemented a new Expenses screen.",
        "Enhanced page structure and data organization.",
        "Integrated with existing financial planning workflows."
      ],
      "benefit": "Provides a modern and intuitive experience for managing client expenses."
    }
  ],
  "improvements": [],
  "bug_fixes": []
}
```

`release_title` is optional and defaults to `Regal Wealth Advisors`.
See `example.json` for a filled file.

## What the builder handles for you

- Version, release date and release title in the header table.
- The intro sentence's trailing `… Version <version>:`.
- The **Highlights** count line, computed from the item counts.
- Numbering each item down its section.
- A gray `Screenshot: <title>` placeholder box under every item — the user clicks it in Word
  and pastes the real screenshot over it.
- An empty section: its heading, table and spacing are removed and it drops out of Highlights.

## Notes

- Sections deliberately keep the reference document's per-section styling differences
  (New Features uses teal `#31849B` numbers in Century Gothic; Improvements and Bug Fixes use
  blue `#2E75B5` on a `#F2F9FF` fill). That is intended — do not "fix" it.
- `python build_release_notes.py --selfcheck` verifies the template and builder still agree.
- Requires `python-docx` and `Pillow`, both already installed.
