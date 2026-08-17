# Final-package QA checklist

## Source freeze

- Confirm source filename, revision, and approval state.
- Confirm no unresolved tracked text, tracked formatting, change bars, revision colors, or comments.
- Confirm late changes appear in complete PDF, text-only source, bibliography, and figures.
- Confirm code/data links are present, reachable as promised, and identical across outputs.

## Complete PDF

- Inspect every page at 100% zoom.
- Confirm figures, tables, equations, symbols, hyperlinks, citations, and references render correctly.
- Confirm no comments, highlights, editing marks, placeholders, or blank figure boxes appear.
- Confirm there are no blank or unexpectedly sparse trailing pages.
- Confirm page count/order and embedded fonts.
- Confirm figure/table caption lists are present when required.

## Text-only source

- Inspect every rendered page.
- Confirm no embedded production graphics remain.
- Confirm all figure/subfigure callouts and captions remain.
- Confirm tables/equations were not mistaken for figures.
- For Word, confirm comments and tracked changes are gone.
- For LaTeX, confirm root/included `.tex`, `.bib`, and required class/style files are present as separate native files.
- Confirm `\nocite{*}` is absent unless explicitly approved.

## References

- Compare every citation to a reference entry and every reference entry to a citation.
- Investigate numbering gaps, duplicates, and a rendered last citation that does not match the last reference.
- Compile from the exact native source intended for upload.

## Figures

- Match uploaded files one-to-one with numbered captions and callouts.
- Follow the controlling combined/separate multipart rule.
- Classify each raster before applying its threshold.
- Check nominal and effective dpi; do not trust a tag alone.
- Check crop, whitespace, font size, line weight, color, transparency, and grayscale legibility.
- Confirm TIFF/EPS, exact numbering, and 15 MB maximum.
- Confirm transparent raster artwork renders correctly against white rather than black.
- For unusable artwork, record where authorized native/high-resolution sources were searched.
- If broader source access is needed, ask once for permission and name the affected figures and proposed locations.
- If recovery fails, give the user the required width, effective dpi, and minimum pixel dimensions for every replacement.

## Upload and recovery

- Open the upload ZIP and compare it to `upload/`.
- Confirm it contains only production deliverables.
- Compare hashes to `manifest.json`.
- Identify who has portal permission to upload/update.
- Save upload and approval receipts.
- Rebuild the full set after any change; never patch one deliverable in isolation.
