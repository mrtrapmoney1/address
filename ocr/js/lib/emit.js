'use strict';

/*
 * emit.js - writing the result out as OcrDocument XML.
 *
 * The single most important decision in this file is the COORDINATE SYSTEM,
 * and it is not an aesthetic one.
 *
 * The existing PDF parser hands its layout engine words shaped like this:
 *
 *     { Page, Text, X, Y, W, H, Size }    X/Y in POINTS from the TOP-LEFT
 *
 * If the OCR path emits pixels-from-wherever, then every rule, every test and
 * every piece of spatial reasoning downstream has to be written twice - once
 * for text PDFs and once for scans - and the two will drift apart. Emitting
 * the same shape in the same units means a scanned invoice arrives at the
 * parser indistinguishable from a text one, and the whole second half of the
 * job is already built.
 *
 * So every box travels back through the preprocessing transform to the
 * original page, and then from pixels into points:
 *
 *     points = pixels * 72 / dpi
 *
 * because a point is one seventy-second of an inch, whatever the scanner did.
 */

const { XmlWriter, parseXml, findAll } = require('./xml.js');
const { mapBoxToSource } = require('./preprocess.js');

const FORMAT_VERSION = '1';

function round(value, places) {
  const p = places === undefined ? 2 : places;
  const factor = Math.pow(10, p);
  return Math.round(value * factor) / factor;
}

/*
 * Convert one word from processed-image pixels to source-page points.
 * `transform` may be null, which means the image was recognised as it came in.
 */
function wordToPoints(word, transform, dpi) {
  const box = transform
    ? mapBoxToSource(transform, word.x, word.y, word.w, word.h)
    : { x: word.x, y: word.y, w: word.w, h: word.h };

  const k = 72 / (dpi > 0 ? dpi : 72);
  return {
    text: word.text,
    x: round(box.x * k),
    y: round(box.y * k),
    w: round(box.w * k),
    h: round(box.h * k),
    conf: word.conf === undefined ? null : round(word.conf, 1),
    column: word.column,
    refined: word.refined === true,
  };
}

function averageConfidence(words) {
  if (words.length === 0) return 0;
  let total = 0;
  let counted = 0;
  for (let i = 0; i < words.length; i += 1) {
    if (typeof words[i].conf === 'number') { total += words[i].conf; counted += 1; }
  }
  return counted === 0 ? 0 : round(total / counted, 1);
}


/*
 * Write the invoice block: the headline fields, what the arithmetic said, and
 * every numeric repair that was made.
 *
 * The repairs are written out in full, one element each, with the text as the
 * engine originally read it. A number that was quietly changed on its way to
 * a spreadsheet and cannot be traced back is worse than no number at all -
 * this is what makes each one auditable.
 */
function writeInvoice(w, invoice) {
  w.open('Invoice', {
    fields: Object.keys(invoice.fields).length,
    repairs: invoice.repairs.length,
    reconciles: invoice.arithmetic.checked
      ? (invoice.arithmetic.ok ? 'yes' : 'no')
      : 'unchecked',
  });

  const names = Object.keys(invoice.fields);
  for (let i = 0; i < names.length; i += 1) {
    const field = invoice.fields[names[i]];
    let value = field.value;
    let iso = null;
    if (value !== null && typeof value === 'object' && value.iso) {
      iso = value.iso;
      value = value.iso;
    }
    w.leaf('Field', {
      name: names[i],
      value: typeof value === 'number' ? round(value) : value,
      iso,
      label: field.label,
      from: field.where,
      conf: field.confidence,
      ambiguous: (field.value && field.value.ambiguous) ? '1' : undefined,
    }, field.text);
  }

  if (invoice.arithmetic.checked) {
    w.empty('Arithmetic', {
      ok: invoice.arithmetic.ok ? '1' : '0',
      computed: invoice.arithmetic.computed === undefined
        ? undefined : round(invoice.arithmetic.computed),
      difference: invoice.arithmetic.difference === undefined
        ? undefined : round(invoice.arithmetic.difference),
      note: invoice.arithmetic.reason,
    });
  }

  for (let i = 0; i < invoice.tables.length; i += 1) {
    w.empty('LineItems', {
      region: invoice.tables[i].index,
      rows: invoice.tables[i].rows,
      columns: invoice.tables[i].columns.join(','),
    });
  }

  for (let i = 0; i < invoice.repairs.length; i += 1) {
    w.empty('Repair', {
      from: invoice.repairs[i].from,
      to: invoice.repairs[i].to,
      column: invoice.repairs[i].column,
    });
  }

  for (let i = 0; i < invoice.notes.length; i += 1) {
    w.leaf('Note', null, invoice.notes[i]);
  }

  w.close('Invoice');
}

/*
 * Build the document.
 *
 * spec = {
 *   source:  { path, kind, pages },
 *   engine:  { name, version, lang, oem, psm, dpi, langSource },
 *   pages:   [ {
 *      number, widthPt, heightPt, pixelWidth, pixelHeight, dpi,
 *      transform, preprocess, page (the segmentation result), text,
 *      imagePath, note
 *   } ],
 *   includeText: boolean
 * }
 */
function buildDocument(spec) {
  const w = new XmlWriter();
  const engine = spec.engine || {};
  const source = spec.source || {};

  w.comment('OcrDocument - coordinates are POINTS from the top-left of the page.');
  w.comment('72 points = 1 inch. This is the same frame the PDF text extractor uses,');
  w.comment('so a scanned page and a text page reach the parser in identical shape.');

  w.open('OcrDocument', {
    version: FORMAT_VERSION,
    generated: (spec.generated || new Date()).toISOString(),
    tool: spec.tool || 'ocr-cli',
  });

  w.empty('Source', {
    path: source.path,
    kind: source.kind,
    pages: source.pages,
    bytes: source.bytes,
  });

  w.empty('Engine', {
    name: engine.name || 'tesseract.js',
    version: engine.version,
    lang: engine.lang,
    oem: engine.oem,
    psm: engine.psm,
    dpi: engine.dpi,
    workers: engine.workers,
    langPath: engine.langPath,
    langSource: engine.langSource,
  });

  const pages = spec.pages || [];
  for (let i = 0; i < pages.length; i += 1) {
    const p = pages[i];
    const seg = p.page || { regions: [], metrics: {}, lines: [] };
    const transform = p.transform || null;
    const dpi = p.dpi || (engine.dpi || 300);

    const allWords = [];
    for (let r = 0; r < seg.regions.length; r += 1) {
      for (let l = 0; l < seg.regions[r].lines.length; l += 1) {
        const ws = seg.regions[r].lines[l].words;
        for (let k = 0; k < ws.length; k += 1) allWords.push(ws[k]);
      }
    }

    w.open('Page', {
      number: p.number,
      width: p.widthPt === undefined ? undefined : round(p.widthPt),
      height: p.heightPt === undefined ? undefined : round(p.heightPt),
      unit: 'pt',
      pixelWidth: p.pixelWidth,
      pixelHeight: p.pixelHeight,
      dpi,
      skew: transform ? round(transform.skewDegrees, 1) : 0,
      confidence: averageConfidence(allWords),
      words: allWords.length,
      regions: seg.regions.length,
      preprocess: p.preprocess ? p.preprocess.join(' > ') : undefined,
      image: p.imagePath,
      note: p.note,
    });

    for (let r = 0; r < seg.regions.length; r += 1) {
      const region = seg.regions[r];
      const rb = region.bounds;
      const rBox = transform
        ? mapBoxToSource(transform, rb.x0, rb.y0, rb.x1 - rb.x0, rb.y1 - rb.y0)
        : { x: rb.x0, y: rb.y0, w: rb.x1 - rb.x0, h: rb.y1 - rb.y0 };
      const k = 72 / dpi;

      w.open('Region', {
        index: region.index,
        kind: region.kind,
        x: round(rBox.x * k),
        y: round(rBox.y * k),
        w: round(rBox.w * k),
        h: round(rBox.h * k),
        columns: region.columnCount,
      });

      /*
       * The column boundaries are written out too. Without them a consumer
       * can see which column a word is in but not where the column edge sat,
       * so it cannot tell an empty cell from a missing one - and for a
       * spreadsheet that difference is the whole point.
       */
      if (region.boundaries && region.boundaries.length > 0) {
        w.open('Columns', { count: region.columnCount });
        for (let c = 0; c < region.boundaries.length; c += 1) {
          const pt = transform
            ? mapBoxToSource(transform, region.boundaries[c], rb.y0, 0, 0)
            : { x: region.boundaries[c] };
          w.empty('Boundary', { index: c, x: round(pt.x * k) });
        }
        w.close('Columns');
      }

      for (let l = 0; l < region.lines.length; l += 1) {
        const line = region.lines[l];
        const lBox = transform
          ? mapBoxToSource(transform, line.x0, line.y0, line.x1 - line.x0, line.y1 - line.y0)
          : { x: line.x0, y: line.y0, w: line.x1 - line.x0, h: line.y1 - line.y0 };

        w.open('Line', {
          index: l,
          x: round(lBox.x * k),
          y: round(lBox.y * k),
          w: round(lBox.w * k),
          h: round(lBox.h * k),
          confidence: averageConfidence(line.words),
        });

        for (let n = 0; n < line.words.length; n += 1) {
          const wp = wordToPoints(line.words[n], transform, dpi);
          w.leaf('Word', {
            x: wp.x,
            y: wp.y,
            w: wp.w,
            h: wp.h,
            size: wp.h,          // font size in points, near enough to the box height
            conf: wp.conf,
            col: wp.column,
            refined: wp.refined ? '1' : undefined,
          }, wp.text);
        }
        w.close('Line');
      }
      w.close('Region');
    }

    /*
     * The invoice reading, when one was asked for. It sits inside the page
     * rather than beside it because a multi-page PDF can hold several
     * invoices, and a field belongs to the page it was found on.
     */
    if (p.invoice) {
      writeInvoice(w, p.invoice);
    }

    if (spec.includeText !== false && p.text) {
      w.leaf('Text', null, p.text);
    }
    w.close('Page');
  }

  w.close('OcrDocument');
  return w.end();
}

/* ------------------------------------------------------------------ reading */

function num(value, fallback) {
  if (value === undefined || value === null || value === '') return fallback;
  const n = Number(value);
  return isNaN(n) ? fallback : n;
}

/*
 * Read an OcrDocument back into plain objects. Used by the tests to prove the
 * round trip, and available to any consumer that would rather not parse XML.
 */
function readDocument(xml) {
  const root = parseXml(xml);
  if (root.name !== 'OcrDocument') throw new Error('not an OcrDocument (root is <' + root.name + '>)');

  const sourceNode = findAll(root, 'Source')[0];
  const engineNode = findAll(root, 'Engine')[0];

  const pages = [];
  for (let i = 0; i < root.children.length; i += 1) {
    const pageNode = root.children[i];
    if (pageNode.name !== 'Page') continue;

    const regions = [];
    let textNode = null;

    for (let j = 0; j < pageNode.children.length; j += 1) {
      const child = pageNode.children[j];
      if (child.name === 'Text') { textNode = child; continue; }
      if (child.name !== 'Region') continue;

      const lines = [];
      const boundaries = [];
      for (let k = 0; k < child.children.length; k += 1) {
        const sub = child.children[k];
        if (sub.name === 'Columns') {
          for (let b = 0; b < sub.children.length; b += 1) {
            boundaries.push(num(sub.children[b].attrs.x, 0));
          }
          continue;
        }
        if (sub.name !== 'Line') continue;
        const words = [];
        for (let n = 0; n < sub.children.length; n += 1) {
          const wn = sub.children[n];
          if (wn.name !== 'Word') continue;
          words.push({
            text: wn.text,
            x: num(wn.attrs.x, 0),
            y: num(wn.attrs.y, 0),
            w: num(wn.attrs.w, 0),
            h: num(wn.attrs.h, 0),
            size: num(wn.attrs.size, 0),
            conf: num(wn.attrs.conf, null),
            column: num(wn.attrs.col, 0),
            refined: wn.attrs.refined === '1',
          });
        }
        lines.push({
          index: num(sub.attrs.index, lines.length),
          x: num(sub.attrs.x, 0),
          y: num(sub.attrs.y, 0),
          w: num(sub.attrs.w, 0),
          h: num(sub.attrs.h, 0),
          confidence: num(sub.attrs.confidence, 0),
          words,
        });
      }

      regions.push({
        index: num(child.attrs.index, regions.length),
        kind: child.attrs.kind || 'text',
        x: num(child.attrs.x, 0),
        y: num(child.attrs.y, 0),
        w: num(child.attrs.w, 0),
        h: num(child.attrs.h, 0),
        columnCount: num(child.attrs.columns, 1),
        boundaries,
        lines,
      });
    }

    const invoiceNode = pageNode.children.filter(function (c) { return c.name === 'Invoice'; })[0];
    let invoice = null;
    if (invoiceNode) {
      const fields = {};
      const repairs = [];
      const notes = [];
      let arithmetic = null;
      for (let k = 0; k < invoiceNode.children.length; k += 1) {
        const child = invoiceNode.children[k];
        if (child.name === 'Field') {
          fields[child.attrs.name] = {
            value: child.attrs.iso ? child.attrs.iso : child.attrs.value,
            text: child.text,
            label: child.attrs.label,
            where: child.attrs.from,
            confidence: num(child.attrs.conf, null),
            ambiguous: child.attrs.ambiguous === '1',
          };
        } else if (child.name === 'Arithmetic') {
          arithmetic = {
            ok: child.attrs.ok === '1',
            computed: num(child.attrs.computed, null),
            difference: num(child.attrs.difference, null),
            note: child.attrs.note || '',
          };
        } else if (child.name === 'Repair') {
          repairs.push({ from: child.attrs.from, to: child.attrs.to });
        } else if (child.name === 'Note') {
          notes.push(child.text);
        }
      }
      invoice = { fields, repairs, notes, arithmetic };
    }

    pages.push({
      invoice,
      number: num(pageNode.attrs.number, pages.length + 1),
      width: num(pageNode.attrs.width, 0),
      height: num(pageNode.attrs.height, 0),
      dpi: num(pageNode.attrs.dpi, 0),
      skew: num(pageNode.attrs.skew, 0),
      confidence: num(pageNode.attrs.confidence, 0),
      wordCount: num(pageNode.attrs.words, 0),
      preprocess: pageNode.attrs.preprocess || '',
      note: pageNode.attrs.note || '',
      regions,
      text: textNode ? textNode.text : '',
    });
  }

  return {
    version: root.attrs.version,
    generated: root.attrs.generated,
    source: sourceNode ? sourceNode.attrs : {},
    engine: engineNode ? engineNode.attrs : {},
    pages,
  };
}

/*
 * The flat word stream, in the exact shape the PDF parser's own extractor
 * returns. This is the function that makes a scan and a text PDF
 * interchangeable.
 */
function toWordStream(doc) {
  const words = [];
  for (let i = 0; i < doc.pages.length; i += 1) {
    const page = doc.pages[i];
    for (let r = 0; r < page.regions.length; r += 1) {
      for (let l = 0; l < page.regions[r].lines.length; l += 1) {
        const ws = page.regions[r].lines[l].words;
        for (let n = 0; n < ws.length; n += 1) {
          words.push({
            Page: page.number,
            Text: ws[n].text,
            X: ws[n].x,
            Y: ws[n].y,
            W: ws[n].w,
            H: ws[n].h,
            Size: ws[n].size,
            Conf: ws[n].conf,
            Column: ws[n].column,
            Region: page.regions[r].index,
            RegionKind: page.regions[r].kind,
          });
        }
      }
    }
  }
  return words;
}

module.exports = {
  FORMAT_VERSION,
  buildDocument,
  readDocument,
  toWordStream,
  wordToPoints,
  averageConfidence,
};
