'use strict';

const test = require('node:test');
const assert = require('node:assert');

const { XmlWriter, parseXml, findAll, escapeAttr, stripInvalid, unescapeXml } = require('../lib/xml.js');
const emit = require('../lib/emit.js');

test('the five predefined entities are escaped in text and attributes', () => {
  const w = new XmlWriter();
  w.open('Root', { note: 'a & b < c > d "e" \'f\'' });
  w.leaf('Value', null, 'x < y & z > w');
  w.close('Root');
  const xml = w.end();

  const root = parseXml(xml);
  assert.strictEqual(root.attrs.note, 'a & b < c > d "e" \'f\'');
  assert.strictEqual(findAll(root, 'Value')[0].text, 'x < y & z > w');
});

test('newlines and tabs in attributes survive the round trip', () => {
  // An XML parser normalises raw whitespace in attributes; escaped it does not.
  const w = new XmlWriter();
  w.empty('E', { multi: 'line one\nline two\tend' });
  const xml = w.end();
  assert.ok(xml.indexOf('&#10;') > 0, 'newline should be written as a character reference');
  assert.strictEqual(parseXml(xml).attrs.multi, 'line one\nline two\tend');
});

test('control characters that XML forbids are stripped, not written', () => {
  // OCR of a noisy scan can produce these; writing one makes the whole
  // document unparseable, so they must never reach the file.
  const dirty = 'ok\u0000bad\u0008worse\u001Fend';
  assert.strictEqual(stripInvalid(dirty), 'okbadworseend');

  const w = new XmlWriter();
  w.leaf('W', { a: dirty }, dirty);
  const xml = w.end();
  assert.doesNotThrow(() => parseXml(xml));
});

test('an attribute set to null is omitted, not written empty', () => {
  // "absent" and "present but blank" are different facts about a page.
  const w = new XmlWriter();
  w.empty('W', { there: 1, missing: null, alsoMissing: undefined });
  const xml = w.end();
  assert.ok(xml.indexOf('there="1"') > 0);
  assert.strictEqual(xml.indexOf('missing'), -1);
});

test('an unbalanced document throws instead of producing a broken file', () => {
  const w = new XmlWriter();
  w.open('A');
  w.open('B');
  w.close('B');
  assert.throws(() => w.end(), /still open/);

  const w2 = new XmlWriter();
  w2.open('A');
  assert.throws(() => w2.close('WrongName'), /but <A> is open/);
});

test('numeric character references are decoded on the way back in', () => {
  assert.strictEqual(unescapeXml('&#65;&#x42;&amp;'), 'AB&');
});

test('escapeAttr never leaves a raw quote that would close the attribute', () => {
  assert.strictEqual(escapeAttr('say "hi"'), 'say &quot;hi&quot;');
});

/* ------------------------------------------------------- OcrDocument itself */

function fakePage() {
  const words = [];
  for (let row = 0; row < 4; row += 1) {
    words.push({ text: 'Item' + row, x: 100, y: 100 + row * 40, w: 120, h: 30, conf: 95 });
    words.push({ text: (row + 1) + '00.00', x: 900, y: 100 + row * 40, w: 100, h: 30, conf: 88 });
  }
  const segment = require('../lib/segment.js');
  return segment.segmentPage(words, { pageWidth: 1200 });
}

test('an OcrDocument round-trips through the reader', () => {
  const xml = emit.buildDocument({
    generated: new Date('2026-01-01T00:00:00Z'),
    source: { path: 'C:\\scans\\invoice.png', kind: 'image', pages: 1 },
    engine: { name: 'tesseract.js', lang: 'eng', psm: '3', dpi: 300 },
    pages: [{
      number: 1,
      widthPt: 612,
      heightPt: 792,
      pixelWidth: 2550,
      pixelHeight: 3300,
      dpi: 300,
      transform: null,
      preprocess: ['grayscale', 'sauvola'],
      page: fakePage(),
      text: 'Item0 100.00',
    }],
  });

  const doc = emit.readDocument(xml);
  assert.strictEqual(doc.version, '1');
  assert.strictEqual(doc.engine.lang, 'eng');
  assert.strictEqual(doc.pages.length, 1);
  assert.strictEqual(doc.pages[0].width, 612);
  assert.strictEqual(doc.pages[0].dpi, 300);
  assert.ok(doc.pages[0].regions.length >= 1);
  assert.strictEqual(doc.pages[0].preprocess, 'grayscale > sauvola');
});

test('coordinates are emitted in points, not pixels', () => {
  /*
   * The contract that lets a scanned page and a text PDF reach the parser in
   * the same shape. A word at pixel 900 on a 300 dpi page is at 216 points.
   */
  const segment = require('../lib/segment.js');
  const page = segment.segmentPage(
    [{ text: 'TOTAL', x: 900, y: 600, w: 150, h: 30, conf: 90 }], { pageWidth: 2550 },
  );
  const xml = emit.buildDocument({
    source: {}, engine: { dpi: 300 },
    pages: [{ number: 1, dpi: 300, transform: null, page }],
  });
  const doc = emit.readDocument(xml);
  const word = doc.pages[0].regions[0].lines[0].words[0];
  assert.strictEqual(word.text, 'TOTAL');
  assert.strictEqual(word.x, 216);   // 900 * 72 / 300
  assert.strictEqual(word.y, 144);   // 600 * 72 / 300
  assert.strictEqual(word.w, 36);    // 150 * 72 / 300
});

test('toWordStream produces exactly the shape the PDF parser expects', () => {
  const segment = require('../lib/segment.js');
  const page = segment.segmentPage(
    [{ text: 'Acme', x: 300, y: 300, w: 120, h: 30, conf: 91 }], { pageWidth: 2550 },
  );
  const xml = emit.buildDocument({
    source: {}, engine: { dpi: 300 },
    pages: [{ number: 1, dpi: 300, transform: null, page }],
  });
  const stream = emit.toWordStream(emit.readDocument(xml));
  assert.strictEqual(stream.length, 1);
  const w = stream[0];
  for (const key of ['Page', 'Text', 'X', 'Y', 'W', 'H', 'Size']) {
    assert.ok(Object.prototype.hasOwnProperty.call(w, key), 'missing field ' + key);
  }
  assert.strictEqual(w.Page, 1);
  assert.strictEqual(w.Text, 'Acme');
  assert.strictEqual(w.X, 72);
});

test('a page with no words still produces a valid document', () => {
  const segment = require('../lib/segment.js');
  const xml = emit.buildDocument({
    source: {}, engine: { dpi: 300 },
    pages: [{ number: 1, dpi: 300, transform: null, page: segment.segmentPage([], {}), note: 'blank page' }],
  });
  const doc = emit.readDocument(xml);
  assert.strictEqual(doc.pages[0].regions.length, 0);
  assert.strictEqual(doc.pages[0].note, 'blank page');
});

test('column boundaries are written so an empty cell can be told from a missing one', () => {
  const xml = emit.buildDocument({
    source: {}, engine: { dpi: 300 },
    pages: [{ number: 1, dpi: 300, transform: null, page: fakePage() }],
  });
  const doc = emit.readDocument(xml);
  const region = doc.pages[0].regions[0];
  assert.ok(region.columnCount >= 2, 'expected at least two columns');
  assert.strictEqual(region.boundaries.length, region.columnCount - 1);
});

test('reading something that is not an OcrDocument is refused clearly', () => {
  assert.throws(() => emit.readDocument('<Other/>'), /not an OcrDocument/);
});
