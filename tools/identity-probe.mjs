// Engine-vs-engine identity gate.
//
// Runs scheduleOnce at a FIXED seed on the SAME real data through BOTH engines
// (the refactored engine on this branch and the pre-refactor engine from main),
// then asserts:
//   (3a) byte-identical canonical placement signature
//        session -> day/period/room/parity/cohorts/instructor
//   (3b) identical scoreSchedule {parts, overall, metrics} for half = null / "h1" / "h2"
//        on one fixed placement.
//
// A probe bug hits both engines equally, so it cannot manufacture a false pass for
// a genuine behavioral difference — identity between the two engines is the proof.
//
// Usage:
//   cd tools && npm install
//   node identity-probe.mjs <path-to-data.xlsx>
//   # or:  DATA=/path/to/data.xlsx node identity-probe.mjs
//   # optional: SEED=12345 (default), CUR_ENGINE=../academic-scheduler.jsx,
//   #           BASE_ENGINE=<file>  (default: `git show main:academic-scheduler.jsx`)
//
// Data files are NOT committed (see .gitignore) — pass your own xlsx path.

import { build } from 'esbuild';
import { readFileSync, writeFileSync, mkdirSync, rmSync, existsSync } from 'node:fs';
import { execFileSync } from 'node:child_process';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { dirname, join, resolve } from 'node:path';
import { createRequire } from 'node:module';

const __dirname = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(__dirname, '..');
const tmp = join(__dirname, '.probe-tmp');
const require = createRequire(import.meta.url);

// ---- config ----
const DATA = process.env.DATA || process.argv[2] || './Mandakh_Course_Data.xlsx';
const SEED = Number(process.env.SEED || 12345);
const CUR_ENGINE = resolve(repoRoot, process.env.CUR_ENGINE || 'academic-scheduler.jsx');
const BASE_ENGINE = process.env.BASE_ENGINE ? resolve(process.env.BASE_ENGINE) : null;

if (!existsSync(DATA)) {
  console.error(`\nData file not found: ${DATA}`);
  console.error('Pass the xlsx path as an argument or via the DATA env var, e.g.:');
  console.error('  node identity-probe.mjs "C:/path/Mandakh_Course_Data.xlsx"\n');
  process.exit(2);
}
const XLSX = require('xlsx'); // resolves from tools/node_modules

// ---- helpers ----
function stable(v) {
  if (Array.isArray(v)) return '[' + v.map(stable).join(',') + ']';
  if (v && typeof v === 'object') return '{' + Object.keys(v).sort().map(k => JSON.stringify(k) + ':' + stable(v[k])).join(',') + '}';
  return JSON.stringify(v);
}
const eq = (x, y) => stable(x) === stable(y);
const clone = x => JSON.parse(JSON.stringify(x));

// Bundle an engine .jsx for Node, exposing its pure functions.
// packages:'external' keeps xlsx/react/etc. as runtime imports so Node's own
// loader resolves them from tools/node_modules — this is the fix for xlsx's
// require("stream") blowing up when bundled ("Dynamic require ... not supported").
async function bundleEngine(srcPath, tag) {
  const src = readFileSync(srcPath, 'utf8');
  const withExports = src + '\nexport { scheduleOnce, scoreSchedule, parseCourseWorkbook, mulberry32, DEFAULT_RULES };\n';
  const entry = join(tmp, `${tag}.export.jsx`);
  const out = join(tmp, `${tag}.mjs`);
  writeFileSync(entry, withExports);
  await build({
    entryPoints: [entry], bundle: true, platform: 'node', format: 'esm',
    packages: 'external', outfile: out, loader: { '.jsx': 'jsx' }, jsx: 'transform',
    define: { __BUILD__: '"probe"' }, logLevel: 'silent',
  });
  return import(pathToFileURL(out).href);
}

// ---- canonical placement signature ----
function sessionId(s) {
  return [s.courseIdx, s.type, s.freq, s.phase, (s.cohorts || []).slice().sort().join('+')].join('|');
}
function sigLines(placed) {
  return placed.map(s => {
    const coh = (s.cohorts || []).slice().sort().join('+');
    const room = (s.room ?? '-') + (s.room2 ? ('+' + s.room2) : '');
    const ins = (s.ins ?? '-') + (s.ins2 ? ('+' + s.ins2) : '');
    return `${sessionId(s)} => day:${s.day} period:${s.period} room:${room} parity:${s.parity ?? '-'} cohorts:${coh} ins:${ins}`;
  }).sort();
}

async function main() {
  rmSync(tmp, { recursive: true, force: true });
  mkdirSync(tmp, { recursive: true });

  // Pre-refactor engine: from main by default (works on any clone of this repo).
  let basePath = BASE_ENGINE;
  if (!basePath) {
    basePath = join(tmp, 'base-engine.jsx');
    const baseSrc = execFileSync('git', ['show', 'main:academic-scheduler.jsx'], { cwd: repoRoot, maxBuffer: 64 * 1024 * 1024 });
    writeFileSync(basePath, baseSrc);
  }

  const A = await bundleEngine(basePath, 'base');   // pre-refactor (main)
  const B = await bundleEngine(CUR_ENGINE, 'cur');  // refactored (this branch)

  const wb = XLSX.readFile(DATA);
  const inA = A.parseCourseWorkbook(wb);
  const inB = B.parseCourseWorkbook(wb);
  const parseIdentical = eq(inA, inB);

  const rA = A.scheduleOnce(clone(inA.courses), clone(inA.teachers), [], A.mulberry32(SEED));
  const rB = B.scheduleOnce(clone(inB.courses), clone(inB.teachers), [], B.mulberry32(SEED));

  const linesA = sigLines(rA.placed), linesB = sigLines(rB.placed);
  const placementIdentical = linesA.join('\n') === linesB.join('\n');
  let firstDiff = null;
  if (!placementIdentical) {
    const n = Math.max(linesA.length, linesB.length);
    for (let i = 0; i < n; i++) if (linesA[i] !== linesB[i]) { firstDiff = { i, a: linesA[i] ?? '(none)', b: linesB[i] ?? '(none)' }; break; }
  }

  // score identity on ONE fixed placement (A's) through BOTH engines
  const fixedPlaced = rA.placed, fixedUnplaced = rA.unplaced, teachers = inA.teachers;
  const scoreResults = {};
  let scoreIdentical = true;
  for (const half of [null, 'h1', 'h2']) {
    const sA = A.scoreSchedule(clone(fixedPlaced), clone(fixedUnplaced), clone(teachers), half);
    const sB = B.scoreSchedule(clone(fixedPlaced), clone(fixedUnplaced), clone(teachers), half);
    const identical = eq(sA, sB);
    scoreIdentical = scoreIdentical && identical;
    scoreResults[String(half)] = { identical, A: sA, B: sB };
  }

  console.log('================ IDENTITY GATE ================');
  console.log(`base engine (main): ${basePath === BASE_ENGINE ? BASE_ENGINE : 'git show main:academic-scheduler.jsx'}`);
  console.log(`cur  engine       : ${CUR_ENGINE}`);
  console.log(`data: ${DATA}`);
  console.log(`seed: ${SEED}`);
  console.log(`courses: A=${inA.courses.length} B=${inB.courses.length} | teachers: A=${inA.teachers.length} B=${inB.teachers.length}`);
  console.log(`[bonus] parseCourseWorkbook identical: ${parseIdentical}`);
  console.log('\n--- GATE 3a: placement signature ---');
  console.log(`placed: A=${rA.placed.length} B=${rB.placed.length} | unplaced: A=${rA.unplaced.length} B=${rB.unplaced.length}`);
  console.log(`PLACEMENT BYTE-IDENTICAL: ${placementIdentical}`);
  if (firstDiff) { console.log(`FIRST DIFF @${firstDiff.i}\n  A: ${firstDiff.a}\n  B: ${firstDiff.b}`); }
  console.log('\n--- GATE 3b: scoreSchedule {parts, overall, metrics} ---');
  for (const half of ['null', 'h1', 'h2']) {
    const r = scoreResults[half];
    console.log(`half=${half}: identical=${r.identical} | overall A=${r.A.overall} B=${r.B.overall}`);
    if (!r.identical) { console.log('  A: ' + JSON.stringify(r.A)); console.log('  B: ' + JSON.stringify(r.B)); }
  }
  console.log(`SCORE IDENTICAL (all halves): ${scoreIdentical}`);

  const pass = placementIdentical && scoreIdentical;
  console.log('\n================ RESULT: ' + (pass ? 'PASS (engines identical)' : 'FAIL (divergence found)') + ' ================');
  if (pass) console.log('score parts (half=null): ' + JSON.stringify(scoreResults['null'].A));

  rmSync(tmp, { recursive: true, force: true });
  process.exit(pass ? 0 : 1);
}

main().catch(e => { console.error(e); rmSync(tmp, { recursive: true, force: true }); process.exit(1); });
