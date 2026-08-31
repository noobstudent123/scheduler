# data/ — real course spreadsheets (local only)

Put real course spreadsheets here (e.g. `Mandakh_Course_Data*.xlsx`). **This folder
is gitignored** (`data/*` in `.gitignore`), so real data never enters version
control — only this README is committed.

Point the validation harness at a file in here, for example:

```bash
cd tools
node identity-probe.mjs "../data/Mandakh_Course_Data.xlsx"
```

**Sample / fixture spreadsheets** that *should* be committed (small, non-sensitive
test data) go under `tools/fixtures/` instead — that path is not ignored.
