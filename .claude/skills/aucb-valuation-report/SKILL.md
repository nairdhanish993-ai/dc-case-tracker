---
name: aucb-valuation-report
description: Generate a complete property valuation report in Alleppey Urban Co-operative Bank (AUCB) format from a title deed/property document, a site inspection report, and site photographs, for Dhanish Nair / DC Valuation & Advisory. Use this whenever the user provides these kinds of case files for an AUCB property and asks for a valuation report, a draft report, or to "prepare"/"generate" the report — including terse follow-ups like "here's the next case," "same bank," or "do the same thing" once this pattern has been established in the conversation, even without the word "valuation" appearing again.
---

# AUCB Valuation Report Builder

## What this produces

A `.docx` valuation report matching Alleppey Urban Co-operative Bank's exact
layout (section order, field labels, table shapes) — see
`references/aucb-field-list.md` for the authoritative structure. This is a
legal/financial document a bank will lend against, so accuracy and honest
handling of gaps matter more than a polished-looking guess.

## Inputs to expect

A case typically arrives as some subset of:
- **Title deed / property document** — a PDF scan, often handwritten Malayalam
  cursive, sometimes typed English. Frequently hard to read even at high
  resolution.
- **Site inspection report** — a PDF with a fairly consistent structure
  (case reference, boundaries, area verification, site particulars,
  geolocation, land rate, engineer sign-off), often with a photo annexure
  as its last page(s).
- **Site photographs** — either standalone images or embedded in the site
  inspection PDF's annexure. Usually GPS-Map-Camera-style photos with a
  geolocation/timestamp overlay baked into the image.
- Occasionally a blank bank template (Excel) and/or a previously-completed
  sample report — useful for confirming format, but not required once you
  have `references/aucb-field-list.md`.

## Process

### 1. Read everything before writing anything
Use the `Read` tool with a `pages` range for large PDFs. If `pdftoppm`
(poppler-utils) isn't available yet, or LibreOffice conversion fails with a
generic "source file could not be loaded" error, see **Environment gotchas**
below before assuming the input file is broken — it's usually a missing
package, not a bad PDF.

If site photos are embedded in the site inspection PDF's annexure page,
extract the original image files with `pdfimages -all -p <pdf> outprefix`
rather than screenshotting/cropping the rendered page — you get clean,
undistorted originals to embed in the output.

### 2. Don't force-read the illegible
Handwritten deed scans are often genuinely not reliably legible, even after
upscaling and cropping. Try once at higher resolution (`pdftoppm -r 400`) if
a number or name matters, but if the user says something like "if it can't
be read, leave it — just complete the rest," take that literally: mark the
field "Not legible / not recorded" and move on. Guessing a name, survey
number, or boundary owner in a bank valuation report is worse than an
honest gap — a wrong fact reads as confidently wrong, while a flagged gap
reads as diligence.

### 3. Cross-check photo GPS/timestamps against the case
Every GPS-Map-Camera-style photo carries lat/long and a timestamp baked
into the image. Compare these against the site inspection report's own
recorded geolocation and inspection date. Photos that are off by a
meaningful distance or from a different day almost certainly belong to a
different case that got mixed in. Exclude them from the report and tell
the user which photo you excluded and why, rather than silently including
a mismatched photo or silently dropping it without a word.

### 4. Map extracted data onto the exact AUCB field list
Follow `references/aucb-field-list.md` section-by-section — don't
reorganize or rename fields, the bank expects this specific shape.

### 5. Flag gaps and assumptions inline — don't fabricate, don't leave blank
This bank's own reports (confirmed from a real signed example) already use
a convention of writing the gap or assumption directly into the value cell
in place of a blank, e.g. "Not specified in site notes — recommend
confirmation" or "Rs. 1,600/sqft — ASSUMED, please confirm/override." Follow
that same convention:
- If a field genuinely isn't in any source document, say so explicitly
  rather than inventing a plausible-sounding value.
- If you must assume something to produce a number at all (most commonly:
  building construction rate and depreciation %, when site notes only give
  land rate), pick a defensible figure for the property's stated age/
  condition/quality, show your arithmetic, and flag it clearly — this
  changes the bottom-line valuation, so it cannot be a buried assumption.
- Visually distinguish flagged content (this skill's script uses italic
  orange text) so a human reviewer can scan the document and immediately
  see everything that still needs their sign-off.

### 6. Compute the valuation
- Land value = land area (cents) × adopted rate/cent
- Building value = BUA (sqft) × building rate/sqft × (1 − depreciation%)
- Market value (FMV) = Land value + Building value
- **Realizable value = 85% of FMV, Distress value = 75% of FMV** — confirmed
  from a real signed AUCB report. This ratio is specific to this bank;
  don't assume it carries over if you're ever asked to produce a different
  bank's format.

### 7. Check the declaration for compliance, not just content
The standard declaration states the IBBI registered valuer "personally
inspected" the property. Compare the name in that boilerplate against the
site inspection form's actual sign-off/engineer name. If a different
associate engineer did the physical visit (common — the registered valuer
reviews and certifies, an associate visits site), reword the declaration to
say so accurately and flag it for the valuer's confirmation. Don't let a
template sentence silently misstate who inspected the property — that's a
compliance detail, not a formatting nicety.

### 8. Build the document
Use `scripts/build_report.template.js` as a starting point (docx-js). Copy
it into your working directory, edit the field values for the new case, and
run it — see the comment header in that file. Key layout points already
solved there, worth keeping rather than rediscovering:
- Page is A4 (11907 dxa) with 700 dxa margins each side → usable width is
  10507 dxa. Keep `TOTAL_W` under that (10300 has headroom) or tables spill
  off the page edge.
- The three boundary sub-tables (deed/sketch/site) need 4 equal columns,
  which doesn't match the 3-column (number/label/value) shape of the rest
  of the report — keep them as a separate `Table`, not extra columns
  bolted onto the main one.
- Section headers spanning the full row use `colSpan` matching that
  table's actual column count.

### 9. Render and visually check before delivering
```
node build_report.template.js
soffice --headless --convert-to pdf your_report.docx
pdftoppm -jpeg -r 110 your_report.pdf preview
```
Then `Read` a couple of the preview images — check text isn't clipped at
the right margin, images fit their cells, and the flagged/orange items
render legibly. This document is going to a bank; look at it before you
say it's done.

### 10. Deliver, don't commit
The output is a client-specific deliverable containing a real person's
property, financial, and personal details — send it directly to the user
(e.g. `SendUserFile`) rather than committing it into this git repository,
even though this skill itself lives in the repo.

## Environment gotchas

These wasted real time the first time through — check for them early
instead of assuming the input is broken:

- **`pdftoppm: command not found`** → `poppler-utils` isn't installed.
  `apt-get install -y poppler-utils`.
- **`soffice --convert-to pdf` fails with "Error: source file could not be
  loaded"** on *every* file, including a trivial one-line `.txt` — this is
  not a corrupt-docx problem. It means only `libreoffice-core`/`-common`
  are installed but the actual application modules aren't. Fix:
  `apt-get install -y libreoffice-writer libreoffice-calc`. Confirm by
  checking whether `libswdlo.so` (Writer's core library) exists anywhere
  under `/usr/lib/libreoffice/program/` — if it's missing, that's the
  cause.
- `docx` (npm) may need `npm install docx` in whatever directory you're
  running the build script from — it is not guaranteed to be globally
  available.

## Extending to another bank

If asked to do this for a different bank, don't assume this same field
list, section order, or 85%/75% realizable/distress ratio apply — every
bank format seen so far has been distinct. Ask for (or derive from a
sample report/blank template) that bank's specific structure, and consider
adding a new `references/<bank>-field-list.md` alongside this one rather
than overwriting it.
