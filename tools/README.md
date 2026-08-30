# Validation harness — `tools/`

Local-only validation for the **rules-as-data** engine refactor. Nothing here is
part of the app build or the deployed site — it exists to prove the refactored
engine is behaviorally identical to the pre-refactor production engine.

Two checks:

- **`identity-probe.mjs`** — the engine-vs-engine identity gate. Runs `scheduleOnce`
  at a **fixed seed** on the **same real data** through *both* engines and asserts:
  - **3a** byte-identical canonical placement signature
    (`session -> day/period/room/parity/cohorts/instructor`)
  - **3b** identical `scoreSchedule` `{parts, overall, metrics}` for `half = null / "h1" / "h2"`
    on one fixed placement.

  The two engines are: the refactored one on this branch
  (`../academic-scheduler.jsx`) and the pre-refactor one pulled from `main`
  (`git show main:academic-scheduler.jsx`). A bug in the probe hits *both* engines
  equally, so it cannot manufacture a false pass — **identity between the two
  engines is the proof**, independent of any absolute score value.

- **`smoke.mjs`** — builds a browser bundle from `../academic-scheduler.jsx`
  (React/xlsx/etc. inlined), renders it headless in **jsdom**, and asserts
  **RENDER OK** with **0 console/React errors**.

## Setup

```bash
cd tools
npm install
```

`node_modules/`, build temp (`.probe-tmp/`), and any `*.xlsx` are gitignored —
they are never committed. **Course data stays local**; pass your own file.

## Run the identity probe

```bash
# pass the data file as an argument …
node identity-probe.mjs "C:/path/to/Mandakh_Course_Data.xlsx"

# … or via env var
DATA="/path/to/Mandakh_Course_Data.xlsx" node identity-probe.mjs
```

Exit code `0` = PASS (engines identical), `1` = FAIL (prints the first differing
placement line and any differing score object), `2` = data file not found.

**Fixed seed:** `12345` (override with `SEED=…`). The seed only needs to be
*fixed and shared* across both engines — its specific value is irrelevant to the
proof, since both engines consume the same `mulberry32(SEED)` stream.

Optional overrides:

| Env var | Default | Meaning |
|---|---|---|
| `DATA` / `argv[2]` | `./Mandakh_Course_Data.xlsx` (placeholder) | path to the course xlsx |
| `SEED` | `12345` | RNG seed fed to `mulberry32` for both engines |
| `CUR_ENGINE` | `../academic-scheduler.jsx` | refactored engine under test |
| `BASE_ENGINE` | *(unset → `git show main:academic-scheduler.jsx`)* | pin the pre-refactor engine to a file instead of `main` |

## Run the jsdom smoke test

```bash
node smoke.mjs
```

Exit `0` = RENDER OK + 0 errors. `CUR_ENGINE` and `BUILD_STAMP` (default `smoke`)
can be overridden.

## Build note — the `xlsx` "keep external" flag (don't rediscover this)

Both engines import `xlsx`, which internally does `require("stream")`. If you
esbuild-bundle the engine for Node with everything inlined, that CommonJS
`require` of a Node builtin becomes:

```
Error: Dynamic require of "stream" is not supported
```

**Fix:** the identity probe bundles the engine for Node with **`packages: 'external'`**
(esbuild JS API) — equivalently `--packages=external` on the CLI. That leaves
`xlsx` / `react` / `@supabase/supabase-js` / `lucide-react` as runtime imports,
so Node's own loader resolves them from `tools/node_modules` (where CJS `require`
of a builtin works). The engine's own code stays bundled; only the npm packages
are external.

The **smoke** build is the opposite on purpose: it targets the browser
(`format: 'iife'`, everything **inlined**, `packages` not external) because the
jsdom `<script>` has no module resolver — the whole app must be self-contained.

## Guardrails

This harness lives only on `refactor/rules-as-data` as a review artifact. It does
not touch `main`, the app build, `/beta/`, or Supabase, and nothing here deploys.
