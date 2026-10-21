## What and why

<!-- One or two sentences: what changes, and the problem it solves. Link the issue: Closes #123 -->

## How it was verified

<!-- Paste real output, not "tests pass". -->

```text
uv run ruff check . && uv run ruff format --check . && uv run mypy src && uv run pytest -q tests
```

## Checklist

- [ ] PR title follows Conventional Commits (`feat(query): add offset paging`)
- [ ] New or changed behaviour has tests (Arrange-Act-Assert, including error paths)
- [ ] Tool changes: flat arguments, description <= 30 words, tier caps respected (lite <= 8, standard <= 18, full <= 30)
- [ ] Tool changes: `uv run python scripts/build_manifests.py` re-run (server.json / manifest.json)
- [ ] Writes stay non-destructive (working copy, journal, atomic save) and paths go through the sandbox
- [ ] PLAN.md updated if a design decision or roadmap item changed
- [ ] Tested with a small local model if the tool surface changed (tier: ___, model: ___)
