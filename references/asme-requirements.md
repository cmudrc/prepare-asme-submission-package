# ASME journal final-production baseline

Use the acceptance letter, live portal, and direct editorial-office message as the source of truth. Apply this baseline only when they do not say otherwise.

## Deliverables

- Submit a clean complete PDF.
- Submit text-only source in native Word or LaTeX format.
- Submit one TIFF/EPS graphic for each numbered figure or subfigure as instructed.
- Keep graphics out of text-only source, while retaining figure callouts and captions.
- Upload a separate `.bib` with LaTeX source, even if references also appear elsewhere.
- Keep notes, comments, corrections, and review markup out of the complete PDF.

## Text-only order

The current ASME guidance lists:

1. title;
2. author information, including affiliations, addresses, and email;
3. abstract;
4. single-column, double-spaced text;
5. numbered references;
6. table-caption list;
7. figure-caption list;
8. tables, each on a separate page; and
9. page-number footers.

Automation can remove graphics and review markup, but it cannot safely infer every structural rearrangement. Audit the result manually.

## Graphics

| Figure type | ASME baseline minimum |
|---|---:|
| Grayscale or color raster | 266 effective dpi |
| Composite linework and halftone | 500 effective dpi |
| Linework, plots, diagrams, and text-heavy graphics | 900 effective dpi |
| True vector EPS | Resolution-independent |

Older or paper-specific JMD/JCISE messages have requested 600 dpi for all graphics. Follow the explicit message when present.

- Name files with `Fig` or `Figure` followed immediately by the figure number (an underscore may separate them); include the subfigure letter when separate files are requested. The packager uses `Fig_1.tiff`, `Fig_2.tiff`, and so on.
- Keep each graphic at or below 15 MB.
- Evaluate effective dpi at intended printed size. Metadata changes and resampling do not restore missing information.
- Preserve color when supplied and also inspect grayscale legibility.

## Current primary sources

- ASME, [Final Submission: Preparing and Submitting Your Final Digital Files](https://www.asme.org/publications-submissions/journals/information-for-authors/to-use-the-submission-tool/preparing-submitting-final-digital-files)
- ASME, [Writing a Research Paper](https://www.asme.org/publications-submissions/journals/information-for-authors/journal-guidelines/writing-a-research-paper)
- ASME, [Author FAQs](https://www.asme.org/publications-submissions/journals/information-for-authors/author-faqs)
- ASME, [Information for Authors](https://www.asme.org/publications-submissions/journals/information-for-authors)

Recheck these sources because ASME submissions are transitioning to Wiley Research Exchange and portal behavior may change.
