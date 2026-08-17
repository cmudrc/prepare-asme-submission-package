# Editorial-office regression cases

These anonymized cases are requirements, not examples to quote externally.

## Figures

- A package with compliant-resolution graphics was returned because filenames were not numbered as the figures appeared in the paper.
- Raster figures at 96 dpi were rejected against a 266-dpi minimum.
- Multiple JMD/JCISE packages were asked to provide 600-dpi graphics even when generic guidance differed.
- A package was returned because multipart figures had to be uploaded as separate lettered files, despite general guidance preferring a combined multipart file when possible.
- A package was returned because a numbered figure was missing from the figure-caption list.
- A text-only source edit incorrectly removed figure callouts. The correction restored callouts and removed only the graphics.

Required checks:

- compare figure/subfigure callouts, captions, embedded graphics, and uploaded filenames as sets;
- apply the controlling multipart rule explicitly;
- verify both nominal and effective dpi; and
- preserve callouts/captions in text-only source.

## References and LaTeX

- A package was returned because uncited bibliography entries appeared in the reference list.
- `\nocite{*}` caused repeated reference-count mismatches and production rejection.
- A package was returned because the rendered paper contained fewer citation numbers than the reference list.
- A LaTeX package was returned because the `.bib` was not uploaded separately; the portal warning about duplicate overwrite was misleading.

Required checks:

- reject `\nocite{*}` by default;
- compare reachable citation keys against bibliography keys;
- compare rendered citation/reference numbering for gaps and extras;
- require a separate `.bib`; and
- preserve `.tex` and `.bib` as distinct upload files when the portal supports multiple native files.

## Portal state

- Student/coauthor access did not always permit file replacement.
- After editor approval, the portal no longer allowed edits; the complete corrected package had to be sent to the journal assistant.

Required checks:

- identify the actual uploader before the deadline;
- retain source/output hashes and upload receipts;
- prepare a complete replacement set rather than an isolated patch; and
- escalate to the editorial assistant only with user authorization.
