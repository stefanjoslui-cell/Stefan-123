// build.mjs: packs the canvas film into ONE self-contained HTML file.
//   node bdo-automotive/build.mjs   (from the repo root; esbuild and ffmpeg must be available)
//
// - film.js and lib/motion.js are bundled in memory (iife, minified). An onLoad plugin swaps the
//   beats.json / film.json fetches for inlined JSON and removes the film's own preview player.
//   The source files on disk are never modified.
// - Both fonts declared in index.html are inlined as base64 woff2.
// - audio/mix.wav (else audio/music.wav) is transcoded to AAC-in-m4a at 160 kb/s and inlined.
//   With no audio the player runs on its performance.now() clock.
// - player.js is inlined as a plain script. Output: dist/BDO-Automotive-Hele-bildet.html

import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import * as esbuild from 'esbuild';

const ROOT = path.dirname(fileURLToPath(import.meta.url));
const abs = (p) => path.join(ROOT, p);
const exists = (p) => fs.existsSync(abs(p));
const readText = (p) => fs.readFileSync(abs(p), 'utf8');
const readB64 = (p) => fs.readFileSync(abs(p)).toString('base64');
const escapeHtml = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

// Text edits applied to lib/motion.js at load time. Each must match exactly once, or the build stops.
function motionEdits(beatsJson, filmJson) {
  return [
    {
      from: "fetch('beats.json').then((r) => r.json())",
      to: `Promise.resolve(${beatsJson})`,
      label: 'beats.json fetch',
    },
    {
      from: "fetch('film.json').then((r) => (r.ok ? r.json() : {})).catch(() => ({}))",
      to: `Promise.resolve(${filmJson})`,
      label: 'film.json fetch',
    },
    {
      from: 'if (!RENDER) preview(canvas, window.DURATION);',
      to: '/* preview player removed in the packed build: player.js takes over */',
      label: 'preview() call',
    },
  ];
}

function inlineDataPlugin(beatsJson, filmJson) {
  const target = abs('lib/motion.js');
  return {
    name: 'inline-film-data',
    setup(build) {
      build.onLoad({ filter: /lib[\\/]motion\.js$/ }, (args) => {
        if (path.resolve(args.path) !== target) return undefined;
        let src = fs.readFileSync(args.path, 'utf8');
        for (const { from, to, label } of motionEdits(beatsJson, filmJson)) {
          const n = src.split(from).length - 1;
          if (n !== 1) {
            throw new Error(`build: ${label}: expected the text exactly once in lib/motion.js, found ${n}. Update build.mjs to match the file.`);
          }
          src = src.replace(from, () => to);
        }
        return { contents: src, loader: 'js', resolveDir: path.dirname(args.path) };
      });
    },
  };
}

async function bundleFilm() {
  // Validate and minify the JSON once; it is embedded as a JS literal.
  const beatsJson = JSON.stringify(JSON.parse(readText('beats.json')));
  const filmJson = JSON.stringify(JSON.parse(readText('film.json')));
  const result = await esbuild.build({
    entryPoints: [abs('film.js')],
    bundle: true,
    format: 'iife',
    minify: true,
    platform: 'browser',
    target: ['es2020'],
    legalComments: 'none',
    write: false,
    outfile: path.join(os.tmpdir(), 'bdo-automotive-film.js'),
    plugins: [inlineDataPlugin(beatsJson, filmJson)],
    logLevel: 'warning',
  });
  const code = result.outputFiles[0].text;
  if (/<\/script/i.test(code) || code.includes('<!--')) {
    throw new Error('build: the film bundle contains "</script" or "<!--", which would break the inline script tag');
  }
  return code;
}

function inlineFonts(indexHtml) {
  // The @font-face rules are read from index.html so the family names and url() targets stay in one place.
  const rules = indexHtml.match(/@font-face\s*\{[^}]*\}/g) || [];
  if (rules.length === 0) throw new Error('build: no @font-face rules found in index.html');
  return rules.map((rule) => {
    const m = rule.match(/url\((['"]?)([^'")]+\.woff2)\1\)/);
    if (!m) throw new Error(`build: no .woff2 url() in @font-face rule: ${rule}`);
    return rule.replace(m[0], `url(data:font/woff2;base64,${readB64(m[2])})`);
  }).join('\n');
}

function inlineAudio() {
  const src = ['audio/mix.wav', 'audio/music.wav'].find(exists);
  if (!src) {
    console.log('audio: none found (audio/mix.wav, audio/music.wav); the player runs on its clock');
    return '';
  }
  const tmp = path.join(os.tmpdir(), `bdo-automotive-${process.pid}.m4a`);
  try {
    execFileSync('ffmpeg', [
      '-y', '-v', 'error', '-i', abs(src), '-vn',
      '-c:a', 'aac', '-b:a', '160k', '-movflags', '+faststart', tmp,
    ], { stdio: ['ignore', 'ignore', 'pipe'] });
    const b64 = fs.readFileSync(tmp).toString('base64');
    console.log(`audio: ${src} -> AAC 160 kb/s m4a`);
    return `<audio id="track" preload="auto" src="data:audio/mp4;base64,${b64}"></audio>`;
  } finally {
    fs.rmSync(tmp, { force: true });
  }
}

async function main() {
  const indexHtml = readText('index.html');
  const titleMatch = indexHtml.match(/<title>([\s\S]*?)<\/title>/);
  if (!titleMatch) throw new Error('build: no <title> in index.html');
  const title = titleMatch[1].trim();

  const bundle = await bundleFilm();
  const playerJs = readText('player.js');
  if (/<\/script/i.test(playerJs) || playerJs.includes('<!--')) {
    throw new Error('build: player.js contains "</script" or "<!--"');
  }
  const fontCss = inlineFonts(indexHtml);
  const audioTag = inlineAudio();

  const html = `<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>${escapeHtml(title)}</title>
<style>
${fontCss}
html, body { margin: 0; height: 100%; background: #17191C; overflow: hidden; }
</style>
</head>
<body>
<div id="stage"><canvas id="film" width="1920" height="1080"></canvas></div>
${audioTag}
${logoTag()}
<script>${bundle}</script>
<script>${playerJs}</script>
</body>
</html>
`;

  const outDir = abs('dist');
  fs.mkdirSync(outDir, { recursive: true });
  const out = path.join(outDir, (JSON.parse(readText('film.json')).output || 'BDO-Automotive-Hele-bildet') + '.html');
  fs.writeFileSync(out, html);
  const bytes = fs.statSync(out).size;
  console.log(`wrote ${out}`);
  console.log(`size: ${(bytes / 1048576).toFixed(2)} MiB (${bytes} bytes)`);
}

// The official BDO logo, if it has been dropped into the shared brand folder (../brand/bdo-logo.svg|png).
function logoTag() {
  for (const [f, mime] of [['../brand/bdo-logo.svg', 'image/svg+xml'], ['../brand/bdo-logo.png', 'image/png']]) {
    const p = path.join(ROOT, f);
    if (fs.existsSync(p)) { console.log(`logo: ${f}`); return `<script>window.BDO_LOGO="data:${mime};base64,${fs.readFileSync(p).toString('base64')}";</script>`; }
  }
  console.log('logo: none in brand/ (typographic fallback)');
  return '';
}

main().catch((err) => {
  console.error(err.message || err);
  process.exit(1);
});
