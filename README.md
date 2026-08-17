# Prepare ASME Submission Package

A reusable agent skill and deterministic packager for accepted ASME manuscripts. It turns one authoritative Word document or Overleaf ZIP into a synchronized production-file package: clean full PDF, native text-only source, numbered figure files, upload ZIP, manifest, and QA report.

It prepares files locally. It does **not** upload, publish, email, or submit anything.

## Install the skill

### Codex

Install for your user account:

```bash
mkdir -p "$HOME/.agents/skills"
git clone --depth 1 https://github.com/cmudrc/prepare-asme-submission-package.git \
  "$HOME/.agents/skills/prepare-asme-submission-package"
```

Codex discovers user skills under `~/.agents/skills`. If the skill is not immediately available as `$prepare-asme-submission-package`, restart Codex. See the [Codex skills documentation](https://learn.chatgpt.com/docs/build-skills).

### Claude Code

Install for all of your projects:

```bash
mkdir -p "$HOME/.claude/skills"
git clone --depth 1 https://github.com/cmudrc/prepare-asme-submission-package.git \
  "$HOME/.claude/skills/prepare-asme-submission-package"
```

Or install only for the current project:

```bash
mkdir -p .claude/skills
git clone --depth 1 https://github.com/cmudrc/prepare-asme-submission-package.git \
  .claude/skills/prepare-asme-submission-package
```

Invoke it as `/prepare-asme-submission-package`, or describe the packaging task and let Claude select it. A newly created top-level skills directory may require a Claude Code restart. See the [Claude Code skills documentation](https://code.claude.com/docs/en/skills).

## Runtime prerequisites

The launcher uses Codex's bundled Python when available. Outside Codex, install Python 3.10+ and the Python dependencies in an isolated environment:

```bash
python3 -m venv .venv
.venv/bin/pip install -r scripts/requirements.txt
CODEX_BUNDLED_PYTHON="$PWD/.venv/bin/python" scripts/run_packager.sh --help
```

Also install:

- LibreOffice (`soffice`) for rendering Word manuscripts;
- Poppler (`pdftoppm` and `pdffonts`) for PDF QA; and
- a TeX distribution with `latexmk` when the skill must compile LaTeX locally.

## Quick start

Word manuscript:

```bash
scripts/run_packager.sh manuscript.docx \
  --paper-id MD-26-1234 \
  --journal jmd \
  --multipart separate \
  --output build/MD-26-1234
```

Overleaf ZIP with an accepted clean PDF:

```bash
scripts/run_packager.sh overleaf-source.zip \
  --paper-id JCISE-26-1234 \
  --journal jcise \
  --main main.tex \
  --full-pdf accepted-manuscript.pdf \
  --output build/JCISE-26-1234
```

Successful output has this shape:

```text
build/<paper-id>/
├── upload/                         # portal files
├── <paper-id>_submission_package.zip
├── QA_REPORT.md                    # keep outside the upload set
├── manifest.json                   # hashes, requirements, and QA state
└── qa/                             # rendered pages for visual inspection
```

A nonzero exit means the package is blocked and no upload ZIP is created. A passing package can still contain manual warnings; close every warning in `QA_REPORT.md` before upload. If artwork is missing or unworkable, the skill searches only sources already in scope or authorized, then asks before searching additional locations or requests the native figure from the user.

## Documentation

- [Skill workflow](SKILL.md): authoritative agent instructions and safety boundaries
- [ASME requirements](references/asme-requirements.md): baseline production-file rules and source links
- [Editorial regressions](references/editorial-regressions.md): failure modes learned from prior JMD/JCISE iterations
- [QA checklist](references/qa-checklist.md): required visual and structural closeout
- `scripts/run_packager.sh --help`: complete command-line options

## Development

Run the regression suite from the repository root:

```bash
python3 -m unittest discover -s tests -v
```

Do not commit manuscripts, correspondence, generated packages, or other private submission material. `build/` and `work/` are intentionally ignored.
