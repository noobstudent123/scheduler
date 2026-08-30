// jsdom smoke test for the engine on this branch.
//
// Builds a browser bundle from ../academic-scheduler.jsx (react/react-dom/xlsx/etc.
// inlined), renders it headless in jsdom, and asserts:
//   - RENDER OK  (#root gets populated)
//   - 0 console.error / window errors (0 React errors)
//
// Usage:
//   cd tools && npm install
//   node smoke.mjs
//   # optional: CUR_ENGINE=../academic-scheduler.jsx  BUILD_STAMP=2026-08-30
//
// This builds the bundle itself, so it validates the SOURCE — no committed
// bundle.js is required.

import { build } from 'esbuild';
import { readFileSync, writeFileSync, mkdirSync, rmSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join, resolve } from 'node:path';
import { JSDOM } from 'jsdom';

const __dirname = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(__dirname, '..');
const tmp = join(__dirname, '.probe-tmp');

const CUR_ENGINE = resolve(repoRoot, process.env.CUR_ENGINE || 'academic-scheduler.jsx');
const STAMP = process.env.BUILD_STAMP || 'smoke';

async function main() {
  rmSync(tmp, { recursive: true, force: true });
  mkdirSync(tmp, { recursive: true });

  const entry = join(tmp, 'entry.jsx');
  writeFileSync(entry,
    `import React from "react";\n` +
    `import { createRoot } from "react-dom/client";\n` +
    `import App from ${JSON.stringify(CUR_ENGINE)};\n` +
    `createRoot(document.getElementById("root")).render(<React.StrictMode><App /></React.StrictMode>);\n`
  );
  const bundleFile = join(tmp, 'bundle.js');
  await build({
    entryPoints: [entry], bundle: true, format: 'iife', platform: 'browser',
    outfile: bundleFile, loader: { '.jsx': 'jsx' }, jsx: 'transform', minify: true,
    define: { __BUILD__: JSON.stringify(STAMP), 'process.env.NODE_ENV': '"production"' },
    // engine sources sit at repo root (no node_modules there); resolve their bare
    // imports (react, xlsx, …) from tools/node_modules regardless of importer dir.
    nodePaths: [join(__dirname, 'node_modules')],
    logLevel: 'silent',
  });
  const bundle = readFileSync(bundleFile, 'utf8');

  const errors = [];
  const dom = new JSDOM(
    '<!doctype html><html><head></head><body><div id="root"></div></body></html>',
    { runScripts: 'dangerously', pretendToBeVisual: true, url: 'http://localhost/' }
  );
  const { window } = dom;
  window.matchMedia = window.matchMedia || (() => ({ matches: false, media: '', onchange: null, addEventListener() {}, removeEventListener() {}, addListener() {}, removeListener() {}, dispatchEvent() { return false; } }));
  if (!window.requestAnimationFrame) window.requestAnimationFrame = (cb) => setTimeout(() => cb(Date.now()), 0);
  if (!window.crypto) window.crypto = { getRandomValues: (a) => { for (let i = 0; i < a.length; i++) a[i] = Math.floor(Math.random() * 256); return a; } };
  window.console.error = (...a) => { errors.push(a.map(String).join(' ')); };
  window.console.warn = () => {};
  window.addEventListener('error', (e) => errors.push('window.onerror: ' + (e.error?.stack || e.message)));
  window.addEventListener('unhandledrejection', (e) => errors.push('unhandledrejection: ' + (e.reason?.stack || e.reason)));

  const script = window.document.createElement('script');
  script.textContent = bundle;
  try { window.document.body.appendChild(script); }
  catch (e) { errors.push('bundle threw synchronously: ' + (e.stack || e.message)); }

  await new Promise((r) => setTimeout(r, 400));

  const root = window.document.getElementById('root');
  const html = root ? root.innerHTML : '';
  const rendered = html.trim().length > 0;
  const reactErrors = errors.filter(e => /Minified React error|Warning:.*React|Cannot read|is not a function|is not defined|Maximum update depth|Objects are not valid as a React child/i.test(e));

  console.log('================ JSDOM SMOKE ================');
  console.log('engine: ' + CUR_ENGINE);
  console.log('#root innerHTML length: ' + html.length);
  console.log('RENDER ' + (rendered ? 'OK' : 'FAIL (root empty)'));
  console.log('console/window errors: ' + errors.length + ' | React-class: ' + reactErrors.length);
  if (errors.length) errors.slice(0, 10).forEach((e, i) => console.log(`  [${i}] ` + e.slice(0, 300)));
  if (rendered) console.log('root content head: ' + html.replace(/\s+/g, ' ').slice(0, 160));

  const pass = rendered && reactErrors.length === 0 && errors.length === 0;
  console.log('SMOKE RESULT: ' + (pass ? 'PASS' : 'FAIL'));
  rmSync(tmp, { recursive: true, force: true });
  process.exit(pass ? 0 : 1);
}

main().catch(e => { console.error(e); rmSync(tmp, { recursive: true, force: true }); process.exit(1); });
