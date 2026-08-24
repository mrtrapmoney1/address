'use strict';

/*
 * engine.js - driving Tesseract.js, and doing it once per run rather than
 * once per page.
 *
 * Three things here are worth more than they look:
 *
 *   1. The language data is found LOCALLY. Tesseract.js downloads a 2-4 MB
 *      .traineddata from a CDN on first use, which fails outright on a
 *      machine behind a corporate proxy - exactly the machines this will run
 *      on. Installed from npm and pointed at with `langPath`, it never
 *      touches the network.
 *
 *   2. Workers are REUSED. The Tesseract.js performance guide is blunt about
 *      creating a worker per image: "This is never the correct option." A
 *      worker costs a few hundred milliseconds and tens of megabytes to
 *      start; a scheduler with a small pool recognises a hundred-page batch
 *      in the time a naive loop takes to do a dozen.
 *
 *   3. A page is recognised, then the parts it was UNSURE about are
 *      recognised again on their own. Tesseract does much better on a small
 *      rectangle containing one line than on that line as part of a whole
 *      page, because it can pick segmentation settings that suit it.
 */

const fs = require('fs');
const path = require('path');
const os = require('os');

/* Page segmentation modes, from Tesseract's own publictypes.h. */
const PSM = {
  OSD_ONLY: '0',
  AUTO_OSD: '1',
  AUTO_ONLY: '2',
  AUTO: '3',
  SINGLE_COLUMN: '4',
  SINGLE_BLOCK_VERT_TEXT: '5',
  SINGLE_BLOCK: '6',
  SINGLE_LINE: '7',
  SINGLE_WORD: '8',
  CIRCLE_WORD: '9',
  SINGLE_CHAR: '10',
  SPARSE_TEXT: '11',
  SPARSE_TEXT_OSD: '12',
  RAW_LINE: '13',
};

const OEM = {
  TESSERACT_ONLY: 0,
  LSTM_ONLY: 1,
  TESSERACT_LSTM_COMBINED: 2,
  DEFAULT: 3,
};

/* ------------------------------------------------- finding the language data */

/*
 * Where to look for <lang>.traineddata, in order of preference. A folder the
 * caller named beats one shipped beside the tool, which beats the npm
 * package, which beats going to the internet.
 *
 * Returns { langPath, gzip, source } or null to mean "let Tesseract.js use
 * its CDN default".
 */
function resolveLangPath(lang, explicitPath, baseDir) {
  const root = baseDir || path.join(__dirname, '..');
  const candidates = [];

  if (explicitPath) {
    // A URL is passed through untouched - Tesseract.js knows what to do with it.
    if (/^https?:\/\//i.test(explicitPath)) {
      return { langPath: explicitPath.replace(/\/$/, ''), gzip: true, source: 'url' };
    }
    candidates.push({ dir: explicitPath, source: 'option' });
  }

  candidates.push({ dir: path.join(root, 'tessdata'), source: 'bundled' });
  candidates.push({
    dir: path.join(root, 'node_modules', '@tesseract.js-data', lang, '4.0.0_best_int'),
    source: 'npm(best_int)',
  });
  candidates.push({
    dir: path.join(root, 'node_modules', '@tesseract.js-data', lang, '4.0.0'),
    source: 'npm(4.0.0)',
  });

  for (let i = 0; i < candidates.length; i += 1) {
    const dir = candidates[i].dir;
    // Both spellings are legal; which one is present decides the gzip flag,
    // and getting that wrong produces a confusing "unknown file format".
    const gz = path.join(dir, lang + '.traineddata.gz');
    const plain = path.join(dir, lang + '.traineddata');
    try {
      if (fs.existsSync(gz)) return { langPath: dir, gzip: true, source: candidates[i].source };
      if (fs.existsSync(plain)) return { langPath: dir, gzip: false, source: candidates[i].source };
    } catch (e) {
      // An unreadable candidate is not an error; try the next one.
    }
  }
  return null;
}

/* ---------------------------------------------------------------- the engine */

/*
 * options:
 *   lang           'eng' (or 'eng+fra')
 *   langPath       folder or URL holding the traineddata
 *   workers        how many worker threads (default: sensible for the machine)
 *   psm            page segmentation mode (default AUTO)
 *   dpi            what the images are, so Tesseract stops guessing
 *   whitelist      restrict the character set
 *   logger         function(status, progress)
 */
function OcrEngine(options) {
  const opts = options || {};
  this.lang = opts.lang || 'eng';
  this.oem = opts.oem === undefined ? OEM.LSTM_ONLY : opts.oem;
  this.psm = opts.psm || PSM.AUTO;
  this.dpi = opts.dpi || 300;
  this.whitelist = opts.whitelist || '';
  this.logger = opts.logger || function noop() {};
  this.baseDir = opts.baseDir;

  /*
   * Each worker holds a full copy of the WASM engine and the language model,
   * so the pool is capped by memory rather than by core count. Four is plenty
   * for a batch and still leaves the machine usable.
   */
  const cpus = (os.cpus() || []).length || 2;
  this.workerCount = Math.max(1, Math.min(opts.workers || Math.min(4, cpus), 8));

  this.resolved = resolveLangPath(this.lang.split('+')[0], opts.langPath, this.baseDir);
  this.scheduler = null;
  this.workers = [];
  this.tesseract = null;
}

OcrEngine.prototype.describe = function describe() {
  return {
    lang: this.lang,
    oem: this.oem,
    psm: this.psm,
    dpi: this.dpi,
    workers: this.workerCount,
    langPath: this.resolved ? this.resolved.langPath : '(tesseract.js CDN default)',
    langSource: this.resolved ? this.resolved.source : 'cdn',
  };
};

/*
 * Load tesseract.js lazily and explain clearly if it is missing. "Cannot find
 * module 'tesseract.js'" is a sentence that has cost people an afternoon; the
 * fix is one command and it belongs in the error.
 */
OcrEngine.prototype._require = function _require() {
  if (this.tesseract !== null) return this.tesseract;
  try {
    this.tesseract = require('tesseract.js'); // eslint-disable-line global-require
  } catch (e) {
    throw new Error(
      'tesseract.js is not installed.\n'
      + 'Run this once, in the ocr\\js folder:\n'
      + '    npm install\n'
      + 'That installs the OCR engine and the English language data into\n'
      + 'ocr\\js\\node_modules, after which nothing is ever downloaded again.\n'
      + '(underlying error: ' + e.message + ')',
    );
  }
  return this.tesseract;
};

OcrEngine.prototype.start = async function start() {
  if (this.scheduler !== null) return;
  const Tesseract = this._require();

  const workerOptions = {
    cacheMethod: 'none',    // the data is already local; caching would just
                            // scatter copies of it through the working folder
    logger: (m) => {
      if (m && m.status) this.logger(m.status, m.progress || 0);
    },
    errorHandler: (err) => {
      this.logger('error', 0, String(err));
    },
  };
  if (this.resolved !== null) {
    workerOptions.langPath = this.resolved.langPath;
    workerOptions.gzip = this.resolved.gzip;
  }

  this.scheduler = Tesseract.createScheduler();
  for (let i = 0; i < this.workerCount; i += 1) {
    /* eslint-disable no-await-in-loop */
    const worker = await Tesseract.createWorker(this.lang, this.oem, workerOptions);
    await worker.setParameters(this.parameters());
    /* eslint-enable no-await-in-loop */
    this.scheduler.addWorker(worker);
    this.workers.push(worker);
  }
};

OcrEngine.prototype.parameters = function parameters(overrides) {
  const params = {
    tessedit_pageseg_mode: this.psm,
    /*
     * Keep the runs of spaces between columns. Tesseract collapses them by
     * default, and while this pipeline reads geometry rather than spacing,
     * the text view is what a person looks at when checking a bad page - and
     * collapsed spacing makes a table unreadable.
     */
    preserve_interword_spaces: '1',
    /*
     * Tell it the resolution. Left unset, Tesseract prints "Invalid
     * resolution 0 dpi. Using 70 instead." and then misjudges what is a
     * heading and what is body text.
     */
    user_defined_dpi: String(this.dpi),
  };
  if (this.whitelist) params.tessedit_char_whitelist = this.whitelist;
  return Object.assign(params, overrides || {});
};

/* Flatten Tesseract's block/paragraph/line/word tree into plain words. */
function harvestWords(data, offsetX, offsetY) {
  const dx = offsetX || 0;
  const dy = offsetY || 0;
  const words = [];
  const blocks = data.blocks || [];

  for (let b = 0; b < blocks.length; b += 1) {
    const paragraphs = blocks[b].paragraphs || [];
    for (let p = 0; p < paragraphs.length; p += 1) {
      const lines = paragraphs[p].lines || [];
      for (let l = 0; l < lines.length; l += 1) {
        const ws = lines[l].words || [];
        for (let i = 0; i < ws.length; i += 1) {
          const w = ws[i];
          if (!w.text || String(w.text).trim() === '') continue;
          words.push({
            text: String(w.text),
            x: w.bbox.x0 + dx,
            y: w.bbox.y0 + dy,
            w: w.bbox.x1 - w.bbox.x0,
            h: w.bbox.y1 - w.bbox.y0,
            conf: w.confidence,
            // Tesseract's own idea of which block and line this came from,
            // kept so its opinion can be compared with ours.
            engineBlock: b,
            engineLine: l,
          });
        }
      }
    }
  }
  return words;
}

/*
 * Recognise one image. `image` is a file path or a Buffer.
 * Returns { words, text, confidence, psm }.
 */
OcrEngine.prototype.recognize = async function recognize(image, options) {
  const opts = options || {};
  await this.start();

  const jobOptions = {};
  if (opts.rectangle) jobOptions.rectangle = opts.rectangle;

  const output = { text: true, blocks: true };
  if (opts.hocr) output.hocr = true;
  if (opts.tsv) output.tsv = true;

  const result = await this.scheduler.addJob('recognize', image, jobOptions, output);
  const data = result.data;

  const offsetX = opts.rectangle ? opts.rectangle.left : 0;
  const offsetY = opts.rectangle ? opts.rectangle.top : 0;

  return {
    words: harvestWords(data, offsetX, offsetY),
    text: data.text || '',
    confidence: data.confidence === undefined ? 0 : data.confidence,
    hocr: data.hocr || null,
    tsv: data.tsv || null,
    psm: this.psm,
  };
};

/*
 * The second look.
 *
 * Group the page's words into lines, find the lines the engine was least sure
 * of, and re-run each one on its own with PSM 7 ("this rectangle is a single
 * line of text"). Given that promise, Tesseract stops trying to work out the
 * layout and spends its effort on the characters - and it is common for a
 * line that came back at 55% confidence in the page pass to come back correct
 * and above 80% on its own.
 *
 * A retry is only KEPT if it is better than what it replaces. An unconditional
 * retry would sometimes make a page worse, which is exactly the kind of
 * quiet damage that destroys trust in a tool.
 */
OcrEngine.prototype.refineLowConfidence = async function refineLowConfidence(image, words, options) {
  const opts = options || {};
  const threshold = opts.threshold === undefined ? 70 : opts.threshold;
  const maxLines = opts.maxLines === undefined ? 40 : opts.maxLines;
  const padding = opts.padding === undefined ? 6 : opts.padding;
  const imageWidth = opts.imageWidth || 0;
  const imageHeight = opts.imageHeight || 0;

  // Required lazily so segment.js and engine.js stay independently testable.
  const segment = require('./segment.js'); // eslint-disable-line global-require
  const lines = segment.groupLines(words);

  const suspect = [];
  for (let i = 0; i < lines.length; i += 1) {
    const ws = lines[i].words;
    let total = 0;
    for (let j = 0; j < ws.length; j += 1) total += ws[j].conf;
    const mean = ws.length === 0 ? 100 : total / ws.length;
    if (mean < threshold) suspect.push({ line: lines[i], mean, index: i });
  }

  suspect.sort(function (a, b) { return a.mean - b.mean; });
  const chosen = suspect.slice(0, maxLines);
  if (chosen.length === 0) return { words, retried: 0, improved: 0 };

  const replaced = [];
  let improved = 0;

  for (let i = 0; i < chosen.length; i += 1) {
    const line = chosen[i].line;
    const left = Math.max(0, Math.round(line.x0 - padding));
    const top = Math.max(0, Math.round(line.y0 - padding));
    let width = Math.round((line.x1 - line.x0) + padding * 2);
    let height = Math.round((line.y1 - line.y0) + padding * 2);
    if (imageWidth > 0) width = Math.min(width, imageWidth - left);
    if (imageHeight > 0) height = Math.min(height, imageHeight - top);
    if (width < 4 || height < 4) continue;

    let retry = null;
    try {
      /* eslint-disable no-await-in-loop */
      const worker = this.workers[0];
      await worker.setParameters(this.parameters({ tessedit_pageseg_mode: PSM.SINGLE_LINE }));
      const result = await worker.recognize(image, { rectangle: { left, top, width, height } }, { blocks: true });
      await worker.setParameters(this.parameters());
      /* eslint-enable no-await-in-loop */
      retry = harvestWords(result.data, left, top);
    } catch (e) {
      retry = null;
    }
    if (retry === null || retry.length === 0) continue;

    let total = 0;
    for (let j = 0; j < retry.length; j += 1) total += retry[j].conf;
    const mean = total / retry.length;

    // Only take the retry if it genuinely beat the original.
    if (mean > chosen[i].mean + 1) {
      improved += 1;
      for (let j = 0; j < retry.length; j += 1) retry[j].refined = true;
      replaced.push({ original: line.words, retry });
    }
  }

  if (replaced.length === 0) return { words, retried: chosen.length, improved: 0 };

  const drop = new Set();
  for (let i = 0; i < replaced.length; i += 1) {
    for (let j = 0; j < replaced[i].original.length; j += 1) drop.add(replaced[i].original[j]);
  }
  let out = words.filter(function (w) { return !drop.has(w); });
  for (let i = 0; i < replaced.length; i += 1) out = out.concat(replaced[i].retry);

  return { words: out, retried: chosen.length, improved };
};

OcrEngine.prototype.stop = async function stop() {
  if (this.scheduler === null) return;
  try {
    await this.scheduler.terminate();
  } catch (e) {
    // Terminating an already-dead worker is not worth failing a run over.
  }
  this.scheduler = null;
  this.workers = [];
};

module.exports = {
  OcrEngine,
  PSM,
  OEM,
  resolveLangPath,
  harvestWords,
};
