#!/usr/bin/env node

'use strict';

/*
 * ocr-segment.js - words in, OcrDocument XML out.
 *
 * There are two OCR engines in this toolkit. Tesseract.js runs here in Node;
 * Windows.Media.Ocr runs in PowerShell and is built into Windows 10, so it
 * needs no install at all. Both produce the same raw thing - words with
 * boxes - and both need exactly the same work done afterwards: group the
 * words into lines, find the columns, decide the reading order, write it out.
 *
 * That work lives in ONE place, here, rather than being written twice. A
 * second implementation in PowerShell would drift from this one within a
 * month, and then a page would segment differently depending on which engine
 * read it, which is the kind of inconsistency that is very hard to debug.
 *
 *     node ocr-segment.js --in words.json --out page.ocr.xml
 *
 * The input is the format described in docs/OCR-FORMAT.md.
 */

const fs = require('fs');
const path = require('path');

const segment = require('./lib/segment.js');
const emit = require('./lib/emit.js');

function parseArgs(argv) {
  const opts = {};
  for (let i = 0; i < argv.length; i += 1) {
    if (argv[i] === '--in') { i += 1; opts.input = argv[i]; }
    else if (argv[i] === '--out') { i += 1; opts.out = argv[i]; }
    else if (argv[i] === '--layout') { opts.layout = true; }
    else if (argv[i] === '--no-text') { opts.noText = true; }
    else if (argv[i] === '--help' || argv[i] === '-h') { opts.help = true; }
    else if (!opts.input) { opts.input = argv[i]; }
    else throw new Error('unrecognised argument: ' + argv[i]);
  }
  return opts;
}

const HELP = [
  '',
  '  ocr-segment - turn a word list into OcrDocument XML',
  '',
  '    node ocr-segment.js --in words.json --out page.ocr.xml [--layout] [--no-text]',
  '',
  '  Used by the PowerShell side when recognition was done by the Windows',
  '  OCR engine, so that both engines produce identical output.',
  '',
].join('\n');

function main(argv) {
  let opts;
  try {
    opts = parseArgs(argv);
  } catch (e) {
    process.stderr.write('\n  ' + e.message + '\n');
    return 2;
  }

  if (opts.help || !opts.input) {
    process.stdout.write(HELP + '\n');
    return opts.help ? 0 : 2;
  }
  if (!opts.out) {
    process.stderr.write('\n  --out is required\n');
    return 2;
  }

  const job = JSON.parse(fs.readFileSync(opts.input, 'utf8'));
  const pages = [];

  for (let i = 0; i < (job.pages || []).length; i += 1) {
    const page = job.pages[i];
    const words = (page.words || []).map(function (w) {
      return {
        text: String(w.text),
        x: Number(w.x),
        y: Number(w.y),
        w: Number(w.w),
        h: Number(w.h),
        conf: w.conf === undefined || w.conf === null ? null : Number(w.conf),
      };
    }).filter(function (w) {
      return w.text.trim() !== '' && isFinite(w.x) && isFinite(w.y) && w.w > 0 && w.h > 0;
    });

    const segmented = segment.segmentPage(words, { pageWidth: page.pixelWidth || 0 });

    pages.push({
      number: page.number || (i + 1),
      widthPt: page.widthPt,
      heightPt: page.heightPt,
      pixelWidth: page.pixelWidth,
      pixelHeight: page.pixelHeight,
      dpi: page.dpi || 300,
      /*
       * The Windows engine reports the tilt it found but does not correct
       * for it, so there is no transform to invert - the boxes are already
       * in the source image's own pixels. skewDegrees is carried through
       * for the record.
       */
      transform: page.skew ? {
        scale: 1,
        skewDegrees: Number(page.skew),
        rotateDegrees: 0,
        cropX: 0,
        cropY: 0,
        preRotate: { width: page.pixelWidth, height: page.pixelHeight },
        postRotate: { width: page.pixelWidth, height: page.pixelHeight },
      } : null,
      preprocess: page.preprocess || null,
      page: segmented,
      text: opts.noText ? '' : (page.text || ''),
      note: page.note,
      segmented,
    });
  }

  const xml = emit.buildDocument({
    source: job.source || {},
    engine: job.engine || {},
    pages,
    includeText: !opts.noText,
    tool: 'ocr-segment',
  });

  fs.mkdirSync(path.dirname(path.resolve(opts.out)), { recursive: true });
  fs.writeFileSync(opts.out, xml, 'utf8');

  if (opts.layout) {
    for (let i = 0; i < pages.length; i += 1) {
      const layoutPath = opts.out.replace(/\.ocr\.xml$/i, '') + '.layout.txt';
      fs.writeFileSync(layoutPath, segment.renderLayout(pages[i].segmented, { columns: 120 }), 'utf8');
    }
  }

  let total = 0;
  for (let i = 0; i < pages.length; i += 1) {
    pages[i].page.regions.forEach(function (r) {
      r.lines.forEach(function (l) { total += l.words.length; });
    });
  }
  process.stderr.write('  segmented ' + pages.length + ' page(s), ' + total + ' words -> '
    + path.basename(opts.out) + '\n');
  return 0;
}

if (require.main === module) {
  try {
    process.exitCode = main(process.argv.slice(2));
  } catch (e) {
    process.stderr.write('\n  ' + (e && e.stack ? e.stack : String(e)) + '\n');
    process.exitCode = 1;
  }
}

module.exports = { main, parseArgs };
