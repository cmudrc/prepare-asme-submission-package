# Editorial-office regression cases

These anonymized cases are requirements, not examples to quote externally.

## Figures

- A package with compliant-resolution graphics was returned because filenames were not numbered as the figures appeared in the paper.
- Raster figures at 96 dpi were rejected against a 266-dpi minimum.
- Multiple JMD/JCISE packages were asked to provide 600-dpi graphics even when generic guidance differed.
- A package was returned because multipart figures had to be uploaded as separate lettered files, despite general guidance preferring a combined multipart file when possible.
- A package was returned because a numbered figure was missing from the figure-caption list.
- A text-only source edit incorrectly removed figure callouts. The correction restored callouts and removed only the graphics.
- TIFF metadata claimed a high nominal dpi even though the source pixels delivered much lower effective resolution at the placed width.
- Transparent TIFF subfigures rendered against black in downstream tooling; flattening alpha against white removed the ambiguity.

Required checks:

- compare figure/subfigure callouts, captions, embedded graphics, and uploaded filenames as sets;
- apply the controlling multipart rule explicitly;
- verify both nominal and effective dpi;
- preserve callouts/captions in text-only source; and
- flatten transparent raster exports against an explicit white background.

When the embedded or supplied graphic cannot satisfy these checks, search only source bundles and project locations already in scope. Search connected storage or communications read-only when the user has authorized those sources; otherwise ask once for permission, naming the figures and proposed locations. If no compliant source exists, request regeneration with exact pixel and width requirements rather than manufacturing dpi.

## Word source and PDF rendering

- Accepting tracked text while retaining tracked formatting records produced visible change bars and revision-colored references in LibreOffice.
- Removing deleted runs without removing the emptied deleted paragraphs created trailing blank PDF pages.
- LibreOffice reflowed a Word manuscript to a different page count than native Microsoft Word even after the revision markup was clean.

Required checks:

- remove text, move, and formatting-revision records from production source;
- remove paragraphs emptied by accepted deletions while preserving section breaks and meaningful objects;
- block blank rendered pages; and
- require native-Word comparison when the complete PDF was rendered with LibreOffice.

## References and LaTeX

- A package was returned because uncited bibliography entries appeared in the reference list.
- `\nocite{*}` caused repeated reference-count mismatches and production rejection.
- A package was returned because the rendered paper contained fewer citation numbers than the reference list.
- A LaTeX package was returned because the `.bib` was not uploaded separately; the portal warning about duplicate overwrite was misleading.
- Archived `.tex`, `.cls`, and `.bst` files copied into the native upload set created false duplicate-source and reference-audit signals.

Required checks:

- reject `\nocite{*}` by default;
- compare reachable citation keys against bibliography keys;
- compare rendered citation/reference numbering for gaps and extras;
- require a separate `.bib`;
- preserve `.tex` and `.bib` as distinct upload files when the portal supports multiple native files; and
- copy the reachable source dependency closure rather than every file in the Overleaf archive.

## Portal state

- Student/coauthor access did not always permit file replacement.
- After editor approval, the portal no longer allowed edits; the complete corrected package had to be sent to the journal assistant.

Required checks:

- identify the actual uploader before the deadline;
- retain source/output hashes and upload receipts;
- prepare a complete replacement set rather than an isolated patch; and
- escalate to the editorial assistant only with user authorization.
