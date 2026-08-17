#!/bin/sh
set -eu

script_dir=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)

if [ -n "${CODEX_BUNDLED_PYTHON:-}" ] && [ -x "$CODEX_BUNDLED_PYTHON" ]; then
    task_python=$CODEX_BUNDLED_PYTHON
elif [ -n "${HOME:-}" ] && [ -x "$HOME/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3" ]; then
    task_python=$HOME/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3
elif command -v python3 >/dev/null 2>&1 && python3 -c 'import lxml, PIL, pypdf' >/dev/null 2>&1; then
    task_python=$(command -v python3)
else
    echo "No compatible Python runtime found. Load Codex workspace dependencies or install scripts/requirements.txt in an isolated environment." >&2
    exit 2
fi

exec "$task_python" "$script_dir/prepare_submission.py" "$@"
