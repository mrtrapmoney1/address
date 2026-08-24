#!/usr/bin/env node

'use strict';

/*
 * ocr-cli.js - the Node half of the tool.
 *
 * Images in, one OcrDocument XML per document out. PowerShell drives this for
 * a whole folder, but it stands on its own:
 *
 *     node ocr-cli.js --in page.png --out .\_OcrRun
 *     node ocr-cli.js --in .\pages  --out .\_OcrRun --layout --text
 *
 * Nothing is written outside --out, and nothing existing inside it is
 * overwritten without --force. A run that cannot write its results is a run
 * that wasted its time, so that is checked before any recognition starts.
 */

const fs = require('fs');
const path = require('path');

const png = require('./lib/png.js');
const preprocess = require('./lib/preprocess.js');
const segment = require('./lib/segment.js');
const emit = require('./lib/emit.js');
const imagefile = require('./lib/imagefile.js');
const pdfimg = require('./lib/pdfimg.js');
const invoice = require('./lib/invoice.js');
const { OcrEngine, PSM } = require('./lib/engine.js');

const IMAGE_EXTENSIONS = ['.png', '.jpg', '.jpeg', '.bmp', '.gif', '.webp', '.pbm', '.pnm', '.tif', '.tiff'];
const DOCUMENT_EXTENSIONS = ['.pdf'];

/* ------------------------------------------------------------ argument parsing */

/*
 * A deliberately strict parser. An unrecognised switch is an error rather
 * than something quietly ignored - a typo in --workers should not silently
 * run the slow path.
 */
const FLAGS = {
  '--no-preprocess': 'noPreprocess',
  '--no-refine': 'noRefine',
  '--no-deskew': 'noDeskew',
  '--no-text': 'noText',
  '--crop': 'crop',
  '--layout': 'layout',
  '--json': 'json',
  '--keep-images': 'keepImages',
  '--force': 'force',
  '--invoice': 'invoice',
  '--no-invoice': 'noInvoice',
  '--quiet': 'quiet',
  '--summary-json': 'summaryJson',
  '--help': 'help',
  '-h': 'help',
};

const VALUES = {
  '--in': 'input',
  '--out': 'out',
  '--lang': 'lang',
  '--lang-path': 'langPath',
  '--dpi': 'dpi',
  '--target-dpi': 'targetDpi',
  '--psm': 'psm',
  '--workers': 'workers',
  '--method': 'method',
  '--whitelist': 'whitelist',
  '--refine-below': 'refineBelow',
  '--max-dimension': 'maxDimension',
  '--job': 'job',
  '--name': 'name',
  '--summary-file': 'summaryFile',
};

function parseArgs(argv) {
  const opts = {
    lang: 'eng',
    dpi: 0,
    targetDpi: 300,
    psm: PSM.AUTO,
    method: 'sauvola',
    refineBelow: 70,
  };
  const unknown = [];

  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i];
    if (FLAGS[arg]) { opts[FLAGS[arg]] = true; continue; }
    if (VALUES[arg]) {
      i += 1;
      if (i >= argv.length) throw new Error(arg + ' needs a value');
      opts[VALUES[arg]] = argv[i];
      continue;
    }
    if (arg.charAt(0) === '-') { unknown.push(arg); continue; }
    // A bare path is taken as the input, so the simple case needs no switches.
    if (!opts.input) opts.input = arg;
    else unknown.push(arg);
  }

  if (unknown.length > 0) {
    throw new Error('unrecognised argument(s): ' + unknown.join(' ') + '\nRun with --help to see what is accepted.');
  }

  opts.dpi = Number(opts.dpi) || 0;
  opts.targetDpi = Number(opts.targetDpi) || 300;
  opts.refineBelow = Number(opts.refineBelow) || 70;
  if (opts.workers !== undefined) opts.workers = Number(opts.workers) || undefined;
  if (opts.maxDimension !== undefined) opts.maxDimension = Number(opts.maxDimension) || undefined;
  return opts;
}

const HELP = [
  '',
  '  ocr-cli - scanned page images to positioned words (OcrDocument XML)',
  '',
  '  Usage:',
  '    node ocr-cli.js --in <file-or-folder> --out <folder> [options]',
  '',
  '  Required:',
  '    --in   PATH        an image file, or a folder of them',
  '    --out  PATH        the run folder to write results into (created if absent)',
  '',
  '  Recognition:',
  '    --lang CODE        language, default eng (eng+fra for several)',
  '    --lang-path PATH   folder holding <lang>.traineddata[.gz], or a URL',
  '    --psm N            page segmentation mode, default 3 (auto)',
  '                       4 = single column, 6 = single block, 11 = sparse text',
  '    --workers N        worker threads, default min(4, cpus)',
  '    --whitelist CHARS  restrict recognition to these characters',
  '    --refine-below N   re-read lines under N% confidence, default 70',
  '    --no-refine        skip that second pass',
  '',
  '  Image preparation (PNG input only):',
  '    --dpi N            what the source images ARE, default: read from the file',
  '    --target-dpi N     what to scale them to, default 300',
  '    --method NAME      threshold: sauvola (default) or otsu',
  '    --max-dimension N  never scale beyond this many pixels on the long edge',
  '    --crop             trim scanner borders',
  '    --no-deskew        leave the page tilted',
  '    --no-preprocess    hand the image to the engine untouched',
  '',
  '  Invoices:',
  '    --invoice          read the page as an invoice: repair misread digits,',
  '                       label the amount columns, pull out the invoice',
  '                       number, dates and totals, and check that they add up',
  '                       (on by default; --no-invoice turns it off)',
  '',
  '  Output:',
  '    --layout           also write a plain-text picture of the page',
  '    --json             also write the result as JSON',
  '    --no-text          leave the raw engine text out of the XML',
  '    --keep-images      keep the prepared page images',
  '    --force            overwrite results already in the run folder',
  '    --summary-json     print a machine-readable summary on stdout',
  '    --summary-file F   write that summary to a file instead (what the',
  '                       PowerShell side uses - reading a file avoids every',
  '                       stdout-capture problem)',
  '    --quiet            no progress output',
  '',
  '  Batch:',
  '    --job FILE         read a JSON job file instead of switches',
  '',
].join('\n');

/* ----------------------------------------------------------------- utilities */

function ensureDir(dir) {
  fs.mkdirSync(dir, { recursive: true });
}

function listInputs(input) {
  const stat = fs.statSync(input);
  if (stat.isFile()) return [input];
  if (!stat.isDirectory()) throw new Error('not a file or folder: ' + input);

  const wanted = IMAGE_EXTENSIONS.concat(DOCUMENT_EXTENSIONS);
  return fs.readdirSync(input)
    .filter(function (name) {
      return wanted.indexOf(path.extname(name).toLowerCase()) >= 0;
    })
    .sort()
    .map(function (name) { return path.join(input, name); });
}

/*
 * A PNG can state its own resolution in a pHYs chunk. Believing the file
 * beats making the user measure it, and beats assuming 300 when a fax is 204.
 */
function dpiFromPng(buf) {
  let pos = 8;
  while (pos + 8 <= buf.length) {
    const length = buf.readUInt32BE(pos);
    const type = buf.toString('ascii', pos + 4, pos + 8);
    if (type === 'pHYs' && length >= 9) {
      const perUnitX = buf.readUInt32BE(pos + 8);
      const unit = buf[pos + 16];
      if (unit === 1 && perUnitX > 0) return Math.round(perUnitX * 0.0254);
      return 0;
    }
    if (type === 'IDAT' || type === 'IEND') return 0;
    pos += length + 12;
  }
  return 0;
}

function pad(n, width) {
  let s = String(n);
  while (s.length < width) s = '0' + s;
  return s;
}

/* ----------------------------------------------------------- loading pages */

/*
 * Turn one input file into the page images to recognise.
 *
 * An ordinary image is one page and is used as it stands. A PDF is opened and
 * its embedded page images are lifted out - which for a scanned invoice means
 * the scanner's own pixels, at the scanner's own resolution, with nothing
 * resampled on the way. See lib/pdfimg.js for why that is better than
 * rendering the page.
 */
function loadPages(filePath, opts, log) {
  const extension = path.extname(filePath).toLowerCase();
  const name = path.basename(filePath);

  if (DOCUMENT_EXTENSIONS.indexOf(extension) < 0) {
    const buf = fs.readFileSync(filePath);
    const info = imagefile.identify(buf);
    return [{
      number: 1,
      buffer: buf,
      info,
      path: filePath,
      widthPt: 0,
      heightPt: 0,
      dpi: 0,
      ok: info.readable,
      reason: info.reason,
    }];
  }

  let result;
  try {
    result = pdfimg.extractPageImages(fs.readFileSync(filePath));
  } catch (e) {
    return [{ number: 1, ok: false, path: filePath, reason: 'the PDF could not be read: ' + e.message }];
  }

  for (let i = 0; i < result.warnings.length; i += 1) {
    log('  ' + name + ': ' + result.warnings[i]);
  }

  return result.pages.map(function (page) {
    if (!page.ok) {
      return {
        number: page.number, ok: false, path: filePath,
        widthPt: page.widthPt, heightPt: page.heightPt, reason: page.reason,
      };
    }

    // Written into the run's pages folder so the extracted scan can be looked
    // at, and so the engine has a real file path to open.
    const stem = path.basename(filePath, extension);
    const suffix = page.format === 'jpeg' ? '.jpg' : '.png';
    const target = path.join(opts.pagesDir, stem + '-p' + pad(page.number, 4) + suffix);
    fs.writeFileSync(target, page.buffer);

    return {
      number: page.number,
      buffer: page.buffer,
      info: imagefile.identify(page.buffer),
      path: target,
      widthPt: page.widthPt,
      heightPt: page.heightPt,
      dpi: page.dpi,
      ok: true,
      fromPdf: true,
    };
  });
}

/* --------------------------------------------------------------- one document */

async function processPage(engine, page, sourceName, opts, log) {
  const label = sourceName + (page.fromPdf ? ' p' + page.number : '');

  if (!page.ok) {
    log('  ' + label + ': SKIPPED - ' + page.reason);
    return { ok: false, note: page.reason, words: 0, confidence: 0, regions: 0 };
  }

  const info = page.info;
  if (!info.readable) {
    log('  ' + label + ': SKIPPED - ' + info.reason);
    return { ok: false, note: info.reason, words: 0, confidence: 0, regions: 0 };
  }

  /*
   * Preprocessing needs pixels, and pixels are only available for PNG here.
   * A JPEG still gets recognised - it just goes to the engine as it is.
   * Saying so in the note means nobody has to wonder later why one page has a
   * preprocessing history and another does not.
   */
  let recognitionTarget = page.path;
  let transform = null;
  let steps = null;
  let pixelWidth = info.width;
  let pixelHeight = info.height;
  let sourceDpi = opts.dpi || page.dpi || 0;
  let note = '';
  let temporary = null;

  if (!opts.noPreprocess && imagefile.canPreprocess(info)) {
    if (sourceDpi === 0) sourceDpi = dpiFromPng(page.buffer) || 0;
    /*
     * With no declared resolution, guess from the page size rather than
     * assuming. A 2550-pixel-wide page is US Letter at 300 dpi; a
     * 1275-pixel-wide one is the same page at 150 and badly needs the upscale.
     */
    if (sourceDpi === 0) {
      const longEdge = Math.max(info.width, info.height);
      sourceDpi = longEdge >= 2800 ? 300 : longEdge >= 1900 ? 200 : longEdge >= 1400 ? 150 : 96;
      note = 'source dpi not stated; assumed ' + sourceDpi + ' from the page size';
    }

    const decoded = png.decodePng(page.buffer);
    const prepared = preprocess.preparePage(decoded, {
      sourceDpi,
      targetDpi: opts.targetDpi,
      method: opts.method,
      deskew: !opts.noDeskew,
      crop: opts.crop === true,
      maxDimension: opts.maxDimension,
    });

    transform = prepared.transform;
    steps = prepared.steps;
    pixelWidth = prepared.image.width;
    pixelHeight = prepared.image.height;

    const stem = path.basename(page.path, path.extname(page.path));
    temporary = path.join(opts.pagesDir, stem + '.prep.png');
    fs.writeFileSync(temporary, png.encodePng(prepared.image, { dpi: Math.round(transform.effectiveDpi) }));
    recognitionTarget = temporary;
    log('  ' + label + ': ' + prepared.steps.join(' > '));
  } else {
    if (sourceDpi === 0) sourceDpi = opts.targetDpi;
    if (!imagefile.canPreprocess(info)) {
      note = info.format + ' cannot be prepared in-process; recognised as supplied';
    }
    log('  ' + label + ': recognised as supplied (' + info.format + ')');
  }

  const workingDpi = transform ? transform.effectiveDpi : sourceDpi;
  const result = await engine.recognize(recognitionTarget, { tsv: false });
  let words = result.words;

  let refinement = { retried: 0, improved: 0 };
  if (!opts.noRefine && words.length > 0) {
    refinement = await engine.refineLowConfidence(recognitionTarget, words, {
      threshold: opts.refineBelow,
      imageWidth: pixelWidth,
      imageHeight: pixelHeight,
    });
    words = refinement.words;
    if (refinement.improved > 0) {
      log('    second pass: ' + refinement.improved + ' of ' + refinement.retried + ' weak lines improved');
    }
  }

  const segmented = segment.segmentPage(words, { pageWidth: pixelWidth });

  /*
   * The invoice pass. It runs by default because this toolkit is pointed at
   * invoices; --no-invoice turns it off for anything else. It works on the
   * segmented page, so it can only be done here, after the columns are known.
   */
  let invoiceResult = null;
  if (!opts.noInvoice) {
    try {
      invoiceResult = invoice.readInvoice(segmented);
      if (invoiceResult.repairs.length > 0) {
        log('    repaired ' + invoiceResult.repairs.length + ' misread number(s): '
          + invoiceResult.repairs.slice(0, 4).map(function (r) {
            return r.from + ' -> ' + r.to;
          }).join(', ') + (invoiceResult.repairs.length > 4 ? ', ...' : ''));
      }
      const named = Object.keys(invoiceResult.fields);
      if (named.length > 0) {
        log('    fields: ' + named.join(', '));
      }
      if (invoiceResult.arithmetic.checked) {
        log(invoiceResult.arithmetic.ok
          ? '    the amounts reconcile'
          : '    ' + invoiceResult.arithmetic.reason);
      }
    } catch (e) {
      // A failure to read the page AS AN INVOICE must never cost the OCR
      // result, which is complete and correct regardless.
      log('    the invoice pass failed (the OCR result is unaffected): ' + e.message);
      invoiceResult = null;
    }
  }

  /* Page size in points, from the original image and its true resolution. */
  const srcW = transform ? transform.sourceWidth : pixelWidth;
  const srcH = transform ? transform.sourceHeight : pixelHeight;
  const widthPt = page.widthPt > 0 ? page.widthPt : (srcW * 72) / (sourceDpi || 72);
  const heightPt = page.heightPt > 0 ? page.heightPt : (srcH * 72) / (sourceDpi || 72);

  if (!opts.keepImages && temporary !== null) {
    try { fs.unlinkSync(temporary); } catch (e) { /* not worth failing for */ }
  }

  return {
    ok: true,
    format: info.format,
    words: words.length,
    confidence: emit.averageConfidence(words),
    regions: segmented.regions.length,
    refined: refinement.improved,
    segmented,
    pageSpec: {
      number: page.number,
      widthPt,
      heightPt,
      pixelWidth,
      pixelHeight,
      dpi: workingDpi,
      transform,
      preprocess: steps,
      page: segmented,
      invoice: invoiceResult,
      text: opts.noText ? '' : result.text,
      imagePath: opts.keepImages ? path.relative(opts.out, page.path) : undefined,
      note: note || undefined,
    },
  };
}

/*
 * One input file - which may be a single image or a many-page PDF - to one
 * OcrDocument.
 */
async function processDocument(engine, filePath, opts, log) {
  const name = path.basename(filePath);
  const pages = loadPages(filePath, opts, log);
  const specs = [];
  let words = 0;
  let confidenceSum = 0;
  let good = 0;
  let regions = 0;
  let refined = 0;
  const notes = [];

  for (let i = 0; i < pages.length; i += 1) {
    /* eslint-disable no-await-in-loop */
    const outcome = await processPage(engine, pages[i], name, opts, log);
    /* eslint-enable no-await-in-loop */
    if (outcome.ok) {
      specs.push(outcome.pageSpec);
      words += outcome.words;
      confidenceSum += outcome.confidence;
      regions += outcome.regions;
      refined += outcome.refined;
      good += 1;
    } else if (outcome.note) {
      notes.push(outcome.note);
    }
  }

  return {
    file: filePath,
    ok: good > 0,
    format: path.extname(filePath).replace('.', '') || 'image',
    pageCount: pages.length,
    words,
    confidence: good > 0 ? Math.round((confidenceSum / good) * 10) / 10 : 0,
    regions,
    refined,
    specs,
    note: notes.join('; '),
  };
}

/* --------------------------------------------------------------------- main */

async function main(argv) {
  let opts;
  try {
    opts = parseArgs(argv);
  } catch (e) {
    process.stderr.write('\n  ' + e.message + '\n');
    return 2;
  }

  if (opts.help || argv.length === 0) {
    process.stdout.write(HELP + '\n');
    return 0;
  }

  /* A job file supplies the same options as the switches, for PowerShell. */
  if (opts.job) {
    const job = JSON.parse(fs.readFileSync(opts.job, 'utf8'));
    opts = Object.assign(opts, job);
    opts.dpi = Number(opts.dpi) || 0;
    opts.targetDpi = Number(opts.targetDpi) || 300;
  }

  if (!opts.input) {
    process.stderr.write('\n  --in is required (an image file or a folder of them)\n');
    return 2;
  }
  if (!opts.out) {
    process.stderr.write('\n  --out is required (the run folder to write results into)\n');
    return 2;
  }
  if (!fs.existsSync(opts.input)) {
    process.stderr.write('\n  input not found: ' + opts.input + '\n');
    return 2;
  }

  const log = opts.quiet
    ? function noop() {}
    : function write(message) { process.stderr.write(message + '\n'); };

  const outDir = path.resolve(opts.out);
  const xmlDir = path.join(outDir, 'ocr');
  const pagesDir = path.join(outDir, 'pages');
  ensureDir(outDir);
  ensureDir(xmlDir);
  ensureDir(pagesDir);
  opts.pagesDir = pagesDir;
  opts.out = outDir;

  const inputs = listInputs(opts.input);
  if (inputs.length === 0) {
    process.stderr.write('\n  no images found in ' + opts.input + '\n');
    return 1;
  }

  const engine = new OcrEngine({
    lang: opts.lang,
    langPath: opts.langPath,
    psm: String(opts.psm),
    dpi: opts.targetDpi,
    workers: opts.workers,
    whitelist: opts.whitelist,
    baseDir: __dirname,
  });

  const described = engine.describe();
  log('');
  log('  OCR run');
  log('  ' + inputs.length + ' image(s) -> ' + outDir);
  log('  engine   tesseract.js, ' + described.lang + ', psm ' + described.psm
    + ', ' + described.workers + ' worker(s)');
  log('  language ' + described.langSource + ': ' + described.langPath);
  log('');

  const results = [];
  const started = Date.now();

  try {
    await engine.start();

    for (let i = 0; i < inputs.length; i += 1) {
      const filePath = inputs[i];
      log('[' + pad(i + 1, String(inputs.length).length) + '/' + inputs.length + '] '
        + path.basename(filePath));
      let outcome;
      try {
        /* eslint-disable no-await-in-loop */
        outcome = await processDocument(engine, filePath, opts, log);
        /* eslint-enable no-await-in-loop */
      } catch (e) {
        /*
         * One unreadable file must never cost the other three hundred. The
         * failure is recorded against that file and the batch carries on.
         */
        log('  FAILED: ' + e.message);
        outcome = {
          file: filePath, ok: false, note: e.message, format: 'error',
          words: 0, confidence: 0, specs: [], pageCount: 0,
        };
      }
      results.push(outcome);

      if (outcome.ok) {
        const base = path.basename(filePath, path.extname(filePath));
        const doc = emit.buildDocument({
          source: {
            path: filePath,
            kind: path.extname(filePath).toLowerCase() === '.pdf' ? 'pdf' : 'image',
            pages: outcome.specs.length,
            bytes: fs.statSync(filePath).size,
          },
          engine: {
            name: 'tesseract.js',
            lang: described.lang,
            oem: described.oem,
            psm: described.psm,
            dpi: described.dpi,
            workers: described.workers,
            langPath: described.langPath,
            langSource: described.langSource,
          },
          pages: outcome.specs,
          includeText: !opts.noText,
          tool: 'ocr-cli',
        });

        const xmlPath = path.join(xmlDir, base + '.ocr.xml');
        if (fs.existsSync(xmlPath) && !opts.force) {
          log('  refusing to overwrite ' + path.relative(outDir, xmlPath) + ' (use --force)');
        } else {
          fs.writeFileSync(xmlPath, doc, 'utf8');
          const pageNote = outcome.specs.length > 1 ? outcome.specs.length + ' pages, ' : '';
          log('  -> ' + path.relative(outDir, xmlPath)
            + '  (' + pageNote + outcome.words + ' words, ' + outcome.confidence + '% confidence, '
            + outcome.regions + ' region(s))');
        }

        if (opts.layout) {
          const views = outcome.specs.map(function (spec, index) {
            const header = outcome.specs.length > 1 ? ('--- page ' + (index + 1) + ' ---\n') : '';
            return header + segment.renderLayout(spec.page, { columns: 120 });
          });
          fs.writeFileSync(path.join(xmlDir, base + '.layout.txt'), views.join('\n\n'), 'utf8');
        }
        if (opts.json) {
          fs.writeFileSync(
            path.join(xmlDir, base + '.ocr.json'),
            JSON.stringify(emit.readDocument(doc), null, 2),
            'utf8',
          );
        }
      }
    }
  } finally {
    await engine.stop();
  }

  const elapsed = (Date.now() - started) / 1000;
  const good = results.filter(function (r) { return r.ok; });
  let totalWords = 0;
  let confSum = 0;
  for (let i = 0; i < good.length; i += 1) {
    totalWords += good[i].words;
    confSum += good[i].confidence;
  }

  log('');
  log('  done: ' + good.length + '/' + results.length + ' page(s) in ' + elapsed.toFixed(1) + 's, '
    + totalWords + ' words, mean confidence '
    + (good.length ? (confSum / good.length).toFixed(1) : '0') + '%');
  log('');

  const summary = {
    ok: good.length === results.length,
    out: outDir,
    elapsedSeconds: elapsed,
    engine: described,
    pages: results.map(function (r) {
      return {
        file: r.file,
        ok: r.ok,
        format: r.format,
        words: r.words,
        confidence: r.confidence,
        regions: r.regions || 0,
        refined: r.refined || 0,
        note: r.note || '',
      };
    }),
  };

  /*
   * Writing the summary to a FILE rather than stdout is what the PowerShell
   * side reads. Capturing a child process's stdout from PowerShell while also
   * letting its progress reach the console means merging streams, and merged
   * streams turn a stray warning from a dependency into corrupt JSON. A file
   * has none of those failure modes.
   */
  if (opts.summaryFile) {
    fs.writeFileSync(opts.summaryFile, JSON.stringify(summary, null, 2), 'utf8');
  }
  if (opts.summaryJson) {
    process.stdout.write(JSON.stringify(summary, null, 2) + '\n');
  }

  return good.length === results.length ? 0 : 1;
}

if (require.main === module) {
  main(process.argv.slice(2))
    .then(function (code) { process.exitCode = code; })
    .catch(function (e) {
      process.stderr.write('\n  ' + (e && e.stack ? e.stack : String(e)) + '\n');
      process.exitCode = 1;
    });
}

module.exports = { main, parseArgs, dpiFromPng, listInputs };
