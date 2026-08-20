# Prepare ASME Submission Package

[![Tests](https://github.com/cmudrc/prepare-asme-submission-package/actions/workflows/tests.yml/badge.svg?branch=main)](https://github.com/cmudrc/prepare-asme-submission-package/actions/workflows/tests.yml?query=branch%3Amain)
[![Python 3.10–3.13](https://img.shields.io/badge/Python-3.10%E2%80%933.13-3776AB?logo=python&logoColor=white)](https://github.com/cmudrc/prepare-asme-submission-package/actions/workflows/tests.yml)

A reusable agent skill and deterministic packager for accepted ASME manuscripts. It turns one authoritative Word document or Overleaf ZIP into a synchronized production-file package: clean full PDF, native text-only source, numbered figure files, upload ZIP, manifest, and QA report.

It prepares files locally. It does **not** upload, publish, email, or submit anything.

## Install the skill

### Codex (recommended)

Ask Codex's built-in skill installer to install the repository's root skill:

```text
$skill-installer Install the root skill from https://github.com/cmudrc/prepare-asme-submission-package
```

Codex detects newly installed skills automatically. If the skill is not immediately available as `$prepare-asme-submission-package`, restart Codex. See the [Codex skills documentation](https://learn.chatgpt.com/docs/build-skills#install-curated-skills-for-local-use).

### Codex and Claude Code

If Node.js is available, install the skill globally for both agents with the open agent-skills CLI:

```bash
npx skills add cmudrc/prepare-asme-submission-package \
  --global --agent codex --agent claude-code --yes
```

Omit either `--agent` argument when installing for only one agent. See the [`skills` CLI documentation](https://github.com/vercel-labs/skills#install-a-skill).

In Codex, invoke the installed skill as `$prepare-asme-submission-package`. In Claude Code, invoke it as `/prepare-asme-submission-package`. Either agent can also select it automatically from a matching packaging request.

### Manual installation

If neither installer is available, clone the skill into the appropriate user-level directory:

```bash
mkdir -p "$HOME/.agents/skills" "$HOME/.claude/skills"
git clone --depth 1 https://github.com/cmudrc/prepare-asme-submission-package.git \
  "$HOME/.agents/skills/prepare-asme-submission-package"
# Or, for Claude Code only:
git clone --depth 1 https://github.com/cmudrc/prepare-asme-submission-package.git \
  "$HOME/.claude/skills/prepare-asme-submission-package"
```

Codex discovers user skills under `~/.agents/skills`; Claude Code uses `~/.claude/skills`. A newly created top-level skills directory may require an agent restart. See the [Claude Code skills documentation](https://code.claude.com/docs/en/skills).

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
python3 -m pip install -r tests/requirements.txt
python3 -m unittest discover -s tests -v
```

Do not commit manuscripts, correspondence, generated packages, or other private submission material. `build/` and `work/` are intentionally ignored.
