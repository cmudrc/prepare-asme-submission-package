---
name: prepare-asme-submission
description: Build and validate final-production submission packages for ASME journals, including the Journal of Mechanical Design (JMD) and Journal of Computing and Information Science in Engineering (JCISE), from a Word manuscript (.docx) or Overleaf/LaTeX source ZIP. Use when preparing accepted-manuscript production files, extracting numbered TIFF/EPS figures and subfigures, creating native text-only source, rendering the complete clean PDF, checking figures/captions/references/bibliography consistency, or repairing a package rejected by the editorial office.
---

# Prepare an ASME journal submission

Turn one authoritative manuscript source into a synchronized upload package:

- one clean complete PDF containing text, tables, and figures;
- one native text-only Word file or LaTeX source set;
- one TIFF/EPS file per numbered figure or subfigure; and
- a manifest and QA report kept outside the upload ZIP.

## Establish the controlling requirements

1. Read the acceptance letter, current portal instructions, and any direct editorial-office message. Follow the most specific and most recent instruction.
2. Read [references/asme-requirements.md](references/asme-requirements.md) for the current ASME baseline and source links.
3. Read [references/editorial-regressions.md](references/editorial-regressions.md) and turn each applicable failure mode into a closed QA check.
4. Record the paper ID, journal, due date, required statements, multipart-figure policy, supplemental files, and portal naming rules.
5. Keep private correspondence, student identities, and manuscript content out of the skill and repository.

Generic guidance may conflict with a paper-specific instruction. For example, ASME guidance prefers a multipart figure as one file when possible, while a JMD editorial message may require separate `Fig. 9a` through `Fig. 9f` uploads. Follow the message for that paper and record the choice in the manifest.

## Freeze one source

Select the latest approved manuscript as the only source for every deliverable. Never assemble the text-only file, PDF, bibliography, and figures from different revisions.

Before packaging:

- accept or explicitly resolve tracked changes;
- remove review comments and markup from production files;
- verify title, author order, complete affiliations, corresponding-author contact, abstract, references, acknowledgments, funding, conflicts, and data/code availability statements;
- verify that every late correction is present before extracting figures; and
- retain a review-marked manuscript only when requested, outside the production upload set.

## Build the package

Run the packager from the skill directory. Use the bundled Codex Python runtime when available.

```bash
scripts/run_packager.sh manuscript.docx \
  --paper-id MD-26-1234 \
  --journal jmd \
  --multipart separate \
  --output build/MD-26-1234
```

For an Overleaf source ZIP, provide the compiled clean PDF when LaTeX is not installed locally:

```bash
scripts/run_packager.sh overleaf-source.zip \
  --paper-id JCISE-26-1234 \
  --journal jcise \
  --main main.tex \
  --full-pdf accepted-manuscript.pdf \
  --output build/JCISE-26-1234
```

The conservative default classifies raster figures as linework. Override only after visual inspection:

```bash
--figure-kind 2=photo --figure-kind 5=composite
```

When the acceptance letter or editorial office specifies one threshold for all graphics, record and enforce it explicitly:

```bash
--required-dpi 600
```

Record the controlling correspondence and handoff constraints in the same manifest:

```bash
--due-date 2026-09-01 \
--instruction "Upload Fig. 9a through Fig. 9f separately" \
--portal-rule "Only the corresponding author can replace files" \
--supplemental-file supplemental-video.mp4
```

For a LaTeX raster whose final width cannot be inferred, provide its printed width in inches:

```bash
--figure-width 3=3.25 --figure-width 4=6.5
```

Never upsample a low-resolution image to manufacture compliance. Recover or regenerate the native source.

The launcher discovers the bundled Codex Python runtime. Word rendering also requires LibreOffice, and visual QA requires Poppler. If the launcher cannot find them, load the bundled workspace dependencies and set `CODEX_BUNDLED_PYTHON` to the returned Python path. Outside Codex, install `scripts/requirements.txt` in an isolated Python environment.

## Enforce the non-obvious checks

- Remove only graphics from text-only source. Preserve every figure callout and caption.
- Require a separate `.bib` upload for LaTeX even when references appear in the `.tex` or generated `.bbl`.
- Reject `\nocite{*}` unless the paper intentionally cites every bibliography record and the editorial office permits it.
- Compare citation keys to bibliography keys and compare rendered citation/reference numbering for gaps or extras.
- Match numbered figure files to in-text callouts and the figure-caption list, including subfigure letters.
- Name files exactly as figures appear in the paper (`Fig_1`, `Fig_9a`, and so on).
- Treat nominal dpi metadata and effective dpi at printed size as separate checks.
- Preserve a complete replacement set and hashes. Portal warnings, permissions, or editor approval may prevent later edits.

## Inspect and close QA

Read `QA_REPORT.md` and `manifest.json`. A nonzero packager exit means the package is blocked and no upload ZIP is created.

Then follow [references/qa-checklist.md](references/qa-checklist.md):

1. Inspect every rendered page of the complete PDF.
2. Inspect every rendered page of the text-only manuscript.
3. Inspect every figure at intended one- or two-column size.
4. Confirm the upload ZIP contains only production files.
5. Rebuild the entire package after any change.

Do not waive failed resolution, missing/unnumbered figures, ambiguous multipart handling, missing bibliography data, uncited references, compilation failure, or a text-only/full-PDF mismatch.

## Handle portal lockout or editorial correction

If the portal does not allow the needed update:

1. Do not alter unrelated submission metadata or create a new submission.
2. Prepare a complete corrected set: full PDF, text-only native source, all figures, manifest, and a concise delta note.
3. Send or hand off the replacement set to the journal assistant only when the user authorizes that external action.
4. Retain the original upload confirmation and the corrected package hashes until production approval is confirmed.

## Output contract

- `upload/`: files intended for the portal;
- `<paper-id>_submission_package.zip`: upload files only, created after hard checks pass;
- `QA_REPORT.md`: blockers and warnings;
- `manifest.json`: source hash, output hashes, figure classifications/dpi, and controlling requirements; and
- `qa/`: rendered pages for visual inspection.

Do not upload `QA_REPORT.md`, `manifest.json`, or `qa/` unless requested.
