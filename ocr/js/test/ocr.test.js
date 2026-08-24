'use strict';

/*
 * The end-to-end suite: real recognition, on real page images, through the
 * real command line. Everything else in this folder tests a component in
 * isolation; this is the one that fails if the pipeline is wired up wrong.
 *
 * It needs tesseract.js and the language data installed (npm install). If
 * they are absent the suite says so and skips, rather than reporting a
 * failure that is really a missing setup step.
 */

const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const os = require('os');
const path = require('path');

const emit = require('../lib/emit.js');
const segment = require('../lib/segment.js');
const { OcrEngine, resolveLangPath } = require('../lib/engine.js');
const cli = require('../ocr-cli.js');

const SAMPLES = path.join(__dirname, '..', '..', 'samples');
const CLEAN = path.join(SAMPLES, 'page_clean.png');
const SCAN = path.join(SAMPLES, 'page_scan.png');

function engineAvailable() {
  try {
    require.resolve('tesseract.js');
  } catch (e) {
    return false;
  }
  return resolveLangPath('eng', undefined, path.join(__dirname, '..')) !== null;
}

const AVAILABLE = engineAvailable();
const SKIP = AVAILABLE ? false : 'tesseract.js or its language data is not installed - run: npm install';

/* The first three lines of the sample page, which is what the tests check for. */
const EXPECTED_WORDS = ('This is a lot of 12 point text to test the '
  + 'ocr code and see if it works on all types '
  + 'of file format').split(/\s+/);

/*
 * Recall rather than an exact string match. OCR is allowed to differ in
 * punctuation and still be correct; demanding a character-perfect match would
 * make the test brittle for no gain.
 */
function recall(text) {
  const got = text.split(/\s+/).filter(Boolean);
  const pool = got.slice();
  let hit = 0;
  for (let i = 0; i < EXPECTED_WORDS.length; i += 1) {
    const at = pool.indexOf(EXPECTED_WORDS[i]);
    if (at >= 0) { hit += 1; pool.splice(at, 1); }
  }
  return hit / EXPECTED_WORDS.length;
}

function tempDir(label) {
  return fs.mkdtempSync(path.join(os.tmpdir(), 'ocrtest-' + label + '-'));
}

test('the sample pages are present', () => {
  assert.ok(fs.existsSync(CLEAN), 'missing ' + CLEAN);
  assert.ok(fs.existsSync(SCAN), 'missing ' + SCAN);
});

test('a clean page is recognised accurately end to end', { skip: SKIP }, async () => {
  const out = tempDir('clean');
  const code = await cli.main(['--in', CLEAN, '--out', out, '--quiet', '--layout']);
  assert.strictEqual(code, 0, 'the run should succeed');

  const xmlPath = path.join(out, 'ocr', 'page_clean.ocr.xml');
  assert.ok(fs.existsSync(xmlPath), 'no XML was written');

  const doc = emit.readDocument(fs.readFileSync(xmlPath, 'utf8'));
  assert.strictEqual(doc.pages.length, 1);
  const page = doc.pages[0];

  assert.ok(page.wordCount > 40, 'expected a page of text, got ' + page.wordCount + ' words');
  assert.ok(page.confidence > 80, 'expected high confidence, got ' + page.confidence);
  assert.ok(recall(page.text) > 0.9, 'text recall was only ' + recall(page.text));

  // The layout view must have been written and must contain the page's text.
  const layout = fs.readFileSync(path.join(out, 'ocr', 'page_clean.layout.txt'), 'utf8');
  // The layout view spaces words by where they physically sat, so the gaps
  // between them are wide and variable. That is the point of the view.
  assert.ok(/12\s+point\s+text/.test(layout), 'the layout view lost the text');

  fs.rmSync(out, { recursive: true, force: true });
});

test('preparing a bad scan reads more of it than not preparing it', { skip: SKIP }, async () => {
  /*
   * THE REGRESSION THAT JUSTIFIES THE PREPROCESSING STAGE.
   *
   * The same degraded page is recognised twice - once handed straight to the
   * engine, once upscaled, evened out, thresholded and straightened first.
   * If preparation ever stops helping, this test says so.
   */
  const rawOut = tempDir('raw');
  const prepOut = tempDir('prep');

  await cli.main(['--in', SCAN, '--out', rawOut, '--quiet', '--no-preprocess']);
  await cli.main(['--in', SCAN, '--out', prepOut, '--quiet']);

  const rawDoc = emit.readDocument(fs.readFileSync(path.join(rawOut, 'ocr', 'page_scan.ocr.xml'), 'utf8'));
  const prepDoc = emit.readDocument(fs.readFileSync(path.join(prepOut, 'ocr', 'page_scan.ocr.xml'), 'utf8'));

  const rawRecall = recall(rawDoc.pages[0].text);
  const prepRecall = recall(prepDoc.pages[0].text);

  assert.ok(
    prepRecall >= rawRecall,
    'preparation made the page WORSE: raw ' + (rawRecall * 100).toFixed(1)
      + '% vs prepared ' + (prepRecall * 100).toFixed(1) + '%',
  );
  assert.ok(
    prepRecall > 0.8,
    'the prepared scan should be read well, got ' + (prepRecall * 100).toFixed(1) + '%',
  );

  // The tilt must have been found and recorded.
  assert.ok(
    Math.abs(prepDoc.pages[0].skew) > 1,
    'the 1.8 degree tilt should have been detected, got ' + prepDoc.pages[0].skew,
  );

  fs.rmSync(rawOut, { recursive: true, force: true });
  fs.rmSync(prepOut, { recursive: true, force: true });
});

test('every word carries a position and the page is reconstructable', { skip: SKIP }, async () => {
  const out = tempDir('pos');
  await cli.main(['--in', CLEAN, '--out', out, '--quiet']);
  const doc = emit.readDocument(fs.readFileSync(path.join(out, 'ocr', 'page_clean.ocr.xml'), 'utf8'));
  const stream = emit.toWordStream(doc);

  assert.ok(stream.length > 40);
  for (let i = 0; i < stream.length; i += 1) {
    const w = stream[i];
    assert.ok(typeof w.X === 'number' && isFinite(w.X), 'word ' + i + ' has no X');
    assert.ok(typeof w.Y === 'number' && isFinite(w.Y), 'word ' + i + ' has no Y');
    assert.ok(w.W > 0 && w.H > 0, 'word ' + i + ' has no size');
    assert.ok(w.Text.length > 0);
    // A 640x480 page at 96 dpi is 480x360 points; nothing may fall outside it.
    assert.ok(w.X >= 0 && w.X < 700, 'word ' + i + ' is off the page at X=' + w.X);
  }

  // Reading order: words must run down the page, not jump about.
  const firstLine = doc.pages[0].regions[0].lines[0];
  for (let i = 1; i < firstLine.words.length; i += 1) {
    assert.ok(
      firstLine.words[i].x >= firstLine.words[i - 1].x,
      'words within a line are not left to right',
    );
  }

  fs.rmSync(out, { recursive: true, force: true });
});

test('a folder of pages is processed in one run', { skip: SKIP }, async () => {
  const inDir = tempDir('batchin');
  const out = tempDir('batchout');
  fs.copyFileSync(CLEAN, path.join(inDir, 'a.png'));
  fs.copyFileSync(SCAN, path.join(inDir, 'b.png'));

  const code = await cli.main(['--in', inDir, '--out', out, '--quiet']);
  assert.strictEqual(code, 0);
  assert.ok(fs.existsSync(path.join(out, 'ocr', 'a.ocr.xml')));
  assert.ok(fs.existsSync(path.join(out, 'ocr', 'b.ocr.xml')));

  fs.rmSync(inDir, { recursive: true, force: true });
  fs.rmSync(out, { recursive: true, force: true });
});

test('an existing result is not overwritten without --force', { skip: SKIP }, async () => {
  const out = tempDir('force');
  await cli.main(['--in', CLEAN, '--out', out, '--quiet']);
  const xmlPath = path.join(out, 'ocr', 'page_clean.ocr.xml');
  fs.writeFileSync(xmlPath, 'SENTINEL', 'utf8');

  await cli.main(['--in', CLEAN, '--out', out, '--quiet']);
  assert.strictEqual(fs.readFileSync(xmlPath, 'utf8'), 'SENTINEL', 'the result was overwritten');

  await cli.main(['--in', CLEAN, '--out', out, '--quiet', '--force']);
  assert.notStrictEqual(fs.readFileSync(xmlPath, 'utf8'), 'SENTINEL', '--force should overwrite');

  fs.rmSync(out, { recursive: true, force: true });
});

test('a file that is not an image is reported, not crashed on', { skip: SKIP }, async () => {
  const inDir = tempDir('badin');
  const out = tempDir('badout');
  fs.writeFileSync(path.join(inDir, 'notreally.png'), 'this is plain text, not a PNG');
  fs.copyFileSync(CLEAN, path.join(inDir, 'good.png'));

  const code = await cli.main(['--in', inDir, '--out', out, '--quiet']);
  // The bad file fails, so the run reports failure - but the good one is done.
  assert.strictEqual(code, 1);
  assert.ok(fs.existsSync(path.join(out, 'ocr', 'good.ocr.xml')), 'one bad file stopped the batch');

  fs.rmSync(inDir, { recursive: true, force: true });
  fs.rmSync(out, { recursive: true, force: true });
});

test('an unknown switch is refused instead of ignored', async () => {
  const code = await cli.main(['--in', CLEAN, '--out', tempDir('x'), '--wrokers', '4']);
  assert.strictEqual(code, 2, 'a typo in a switch must not be silently ignored');
});

test('the engine reports where it found its language data', { skip: SKIP }, () => {
  const engine = new OcrEngine({ lang: 'eng', baseDir: path.join(__dirname, '..') });
  const described = engine.describe();
  assert.strictEqual(described.lang, 'eng');
  assert.notStrictEqual(described.langSource, 'cdn',
    'the language data should be found locally, never fetched from the internet');
});

test('the segmentation of a real page finds lines and regions', { skip: SKIP }, async () => {
  const engine = new OcrEngine({ lang: 'eng', workers: 1, baseDir: path.join(__dirname, '..') });
  try {
    const result = await engine.recognize(CLEAN);
    const page = segment.segmentPage(result.words, { pageWidth: 640 });
    assert.ok(page.regions.length >= 1);
    assert.ok(page.lines.length >= 5, 'expected several lines, got ' + page.lines.length);
    assert.ok(page.metrics.textHeight > 0);
  } finally {
    await engine.stop();
  }
});
