'use strict';

/*
 * Tests for lifting page images out of a scanned PDF.
 *
 * The PDFs are built here rather than checked in as binary fixtures, so what
 * each test exercises is visible in the test itself - and a failure points at
 * a structure rather than at an opaque file.
 */

const test = require('node:test');
const assert = require('node:assert');
const zlib = require('zlib');

const pdfimg = require('../lib/pdfimg.js');
const png = require('../lib/png.js');

/*
 * A minimal but genuine PDF holding one image XObject: catalogue, page tree,
 * page, image, content stream, xref and trailer.
 */
function buildScannedPdf(options) {
  const opts = options || {};
  const imageBytes = opts.imageBytes;
  const filter = opts.filter;
  const width = opts.width;
  const height = opts.height;
  const colorSpace = opts.colorSpace || 'DeviceGray';
  const bpc = opts.bpc === undefined ? 8 : opts.bpc;
  const widthPt = opts.widthPt || 612;
  const heightPt = opts.heightPt || 792;

  const objects = {
    1: '<< /Type /Catalog /Pages 2 0 R >>',
    2: '<< /Type /Pages /Count 1 /Kids [3 0 R] >>',
    3: '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ' + widthPt + ' ' + heightPt + ']'
       + ' /Resources << /XObject << /Im0 4 0 R >> >> /Contents 5 0 R >>',
  };

  const content = Buffer.from('q ' + widthPt + ' 0 0 ' + heightPt + ' 0 0 cm /Im0 Do Q', 'latin1');
  const parts = [];
  let position = 0;
  const offsets = {};

  function push(chunk) {
    const buf = Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk, 'latin1');
    parts.push(buf);
    position += buf.length;
  }

  push('%PDF-1.4\n');
  [1, 2, 3].forEach(function (n) {
    offsets[n] = position;
    push(n + ' 0 obj\n' + objects[n] + '\nendobj\n');
  });

  offsets[4] = position;
  push('4 0 obj\n<< /Type /XObject /Subtype /Image /Width ' + width + ' /Height ' + height
    + ' /ColorSpace /' + colorSpace + ' /BitsPerComponent ' + bpc
    + ' /Filter /' + filter + ' /Length ' + imageBytes.length + ' >>\nstream\n');
  push(imageBytes);
  push('\nendstream\nendobj\n');

  offsets[5] = position;
  push('5 0 obj\n<< /Length ' + content.length + ' >>\nstream\n');
  push(content);
  push('\nendstream\nendobj\n');

  const xref = position;
  let table = 'xref\n0 6\n0000000000 65535 f \n';
  for (let n = 1; n <= 5; n += 1) {
    table += String(offsets[n]).padStart(10, '0') + ' 00000 n \n';
  }
  push(table);
  push('trailer\n<< /Size 6 /Root 1 0 R >>\nstartxref\n' + xref + '\n%%EOF\n');

  return Buffer.concat(parts);
}

function greyPage(width, height) {
  const data = new Uint8Array(width * height);
  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      data[y * width + x] = ((y % 20) < 6 && x > 10 && x < width - 10) ? 0 : 255;
    }
  }
  return data;
}

test('a Flate-compressed greyscale scan is extracted as a PNG', () => {
  const width = 200;
  const height = 120;
  const pixels = greyPage(width, height);
  const pdf = buildScannedPdf({
    imageBytes: zlib.deflateSync(Buffer.from(pixels)),
    filter: 'FlateDecode',
    width,
    height,
    widthPt: 612,
    heightPt: 367,
  });

  const result = pdfimg.extractPageImages(pdf);
  assert.strictEqual(result.ok, true);
  assert.strictEqual(result.pages.length, 1);

  const page = result.pages[0];
  assert.strictEqual(page.ok, true);
  assert.strictEqual(page.format, 'png');
  assert.strictEqual(page.width, width);
  assert.strictEqual(page.height, height);

  // The extracted PNG must decode back to exactly the pixels that went in -
  // this path is lossless by design.
  const decoded = png.decodePng(page.buffer);
  assert.strictEqual(decoded.width, width);
  assert.strictEqual(decoded.channels, 1);
  assert.deepStrictEqual(Array.from(decoded.data), Array.from(pixels));
});

test('the scan resolution is worked out from the page size', () => {
  /*
   * A 1224-pixel image on a 612-point page is 144 dpi. Getting this right is
   * what lets the preprocessing stage decide how much to enlarge the page -
   * and getting it wrong means either a blurry upscale or none at all.
   */
  const pdf = buildScannedPdf({
    imageBytes: zlib.deflateSync(Buffer.from(new Uint8Array(1224 * 100).fill(255))),
    filter: 'FlateDecode',
    width: 1224,
    height: 100,
    widthPt: 612,
    heightPt: 50,
  });
  const page = pdfimg.extractPageImages(pdf).pages[0];
  assert.strictEqual(page.ok, true);
  assert.strictEqual(page.dpi, 144);
  assert.strictEqual(page.widthPt, 612);
});

test('a JPEG scan is passed through untouched, byte for byte', () => {
  /*
   * The point of the DCTDecode path: a JPEG inside a PDF is already a
   * complete JPEG file. Decoding and re-encoding it would lose detail the OCR
   * engine wants, so the bytes must come out exactly as they went in.
   */
  const jpeg = Buffer.from([
    0xff, 0xd8, 0xff, 0xe0, 0x00, 0x10, 0x4a, 0x46, 0x49, 0x46, 0x00, 0x01,
    0x01, 0x00, 0x00, 0x01, 0x00, 0x01, 0x00, 0x00,
    0xff, 0xc0, 0x00, 0x11, 0x08, 0x00, 0x64, 0x00, 0xc8, 0x03,
    0x01, 0x11, 0x00, 0x02, 0x11, 0x01, 0x03, 0x11, 0x01,
    0xff, 0xd9,
  ]);

  const pdf = buildScannedPdf({
    imageBytes: jpeg, filter: 'DCTDecode', width: 200, height: 100,
  });
  const page = pdfimg.extractPageImages(pdf).pages[0];

  assert.strictEqual(page.ok, true);
  assert.strictEqual(page.format, 'jpeg');
  assert.ok(page.buffer.equals(jpeg), 'the JPEG bytes must be untouched');
});

test('a 1-bit scan is expanded to full-range greyscale', () => {
  // Black-and-white is what a document feeder produces by default.
  const width = 16;
  const height = 4;
  const rowBytes = width / 8;
  const packed = Buffer.alloc(rowBytes * height);
  // Alternate bits: 10101010...
  for (let i = 0; i < packed.length; i += 1) packed[i] = 0xaa;

  const pdf = buildScannedPdf({
    imageBytes: zlib.deflateSync(packed),
    filter: 'FlateDecode',
    width,
    height,
    bpc: 1,
  });
  const page = pdfimg.extractPageImages(pdf).pages[0];
  assert.strictEqual(page.ok, true);

  const decoded = png.decodePng(page.buffer);
  // A 1 must become 255 and a 0 must become 0 - not 1 and 0, which would be
  // a page of solid black to every later stage.
  assert.strictEqual(decoded.data[0], 255);
  assert.strictEqual(decoded.data[1], 0);
});

test('an RGB scan comes out as three channels', () => {
  const width = 8;
  const height = 4;
  const pixels = Buffer.alloc(width * height * 3);
  for (let i = 0; i < pixels.length; i += 1) pixels[i] = (i * 7) & 0xff;

  const pdf = buildScannedPdf({
    imageBytes: zlib.deflateSync(pixels),
    filter: 'FlateDecode',
    width,
    height,
    colorSpace: 'DeviceRGB',
  });
  const page = pdfimg.extractPageImages(pdf).pages[0];
  assert.strictEqual(page.ok, true);

  const decoded = png.decodePng(page.buffer);
  assert.strictEqual(decoded.channels, 3);
  assert.deepStrictEqual(Array.from(decoded.data), Array.from(pixels));
});

test('a compression this cannot decode is named, not guessed at', () => {
  /*
   * CCITT and JBIG2 are real formats that real scanners produce. Half-decoding
   * one would hand the OCR engine a page of noise; saying which format it is
   * lets the caller fall back to the Windows renderer, which handles them.
   */
  const pdf = buildScannedPdf({
    imageBytes: Buffer.from([1, 2, 3, 4]),
    filter: 'CCITTFaxDecode',
    width: 100,
    height: 100,
  });
  const result = pdfimg.extractPageImages(pdf);

  assert.strictEqual(result.ok, false);
  assert.strictEqual(result.pages[0].ok, false);
  assert.strictEqual(result.pages[0].format, 'CCITTFaxDecode');
  assert.ok(/CCITT/.test(result.pages[0].reason));
  assert.ok(/render/i.test(result.pages[0].reason), 'the message should say what to do instead');
});

test('a PDF with no images says it needs no OCR', () => {
  const pdf = Buffer.from(
    '%PDF-1.4\n'
    + '1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n'
    + '2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj\n'
    + '3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] >>\nendobj\n'
    + 'trailer\n<< /Size 4 /Root 1 0 R >>\n%%EOF\n', 'latin1');

  const result = pdfimg.extractPageImages(pdf);
  assert.strictEqual(result.pages.length, 1);
  assert.strictEqual(result.pages[0].ok, false);
  assert.ok(/text PDF/.test(result.pages[0].reason));
});

test('the biggest image on a page is taken as the scan', () => {
  /*
   * A scanned invoice often carries a small logo alongside the page image.
   * The scan is the big one; taking the logo would produce a page with three
   * words on it and no explanation.
   */
  const big = zlib.deflateSync(Buffer.from(new Uint8Array(400 * 300).fill(200)));
  const small = zlib.deflateSync(Buffer.from(new Uint8Array(20 * 20).fill(50)));

  let body = '%PDF-1.4\n';
  const offsets = {};
  function add(n, text) {
    offsets[n] = Buffer.byteLength(body, 'latin1');
    body += n + ' 0 obj\n' + text + '\nendobj\n';
  }
  add(1, '<< /Type /Catalog /Pages 2 0 R >>');
  add(2, '<< /Type /Pages /Count 1 /Kids [3 0 R] >>');
  add(3, '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources '
    + '<< /XObject << /Logo 4 0 R /Scan 5 0 R >> >> >>');

  const head = Buffer.from(body, 'latin1');
  const parts = [head];
  let position = head.length;

  function addImage(n, data, w, h) {
    offsets[n] = position;
    const header = Buffer.from(n + ' 0 obj\n<< /Type /XObject /Subtype /Image /Width ' + w
      + ' /Height ' + h + ' /ColorSpace /DeviceGray /BitsPerComponent 8 '
      + '/Filter /FlateDecode /Length ' + data.length + ' >>\nstream\n', 'latin1');
    const tail = Buffer.from('\nendstream\nendobj\n', 'latin1');
    parts.push(header, data, tail);
    position += header.length + data.length + tail.length;
  }
  addImage(4, small, 20, 20);
  addImage(5, big, 400, 300);
  parts.push(Buffer.from('trailer\n<< /Size 6 /Root 1 0 R >>\n%%EOF\n', 'latin1'));

  const page = pdfimg.extractPageImages(Buffer.concat(parts)).pages[0];
  assert.strictEqual(page.ok, true);
  assert.strictEqual(page.width, 400, 'the large image should have been chosen');
  assert.strictEqual(page.images, 2, 'and both should have been seen');
});

test('a damaged xref does not stop the pages being found', () => {
  /*
   * A PDF that has been through a mail gateway very often has xref offsets
   * that are wrong by a few bytes. The reader scans for objects instead of
   * trusting the table, so a stale offset cannot hide a page.
   */
  const width = 40;
  const height = 20;
  const pdf = buildScannedPdf({
    imageBytes: zlib.deflateSync(Buffer.from(new Uint8Array(width * height).fill(128))),
    filter: 'FlateDecode',
    width,
    height,
  });

  // Corrupt every offset in the xref table.
  const text = pdf.toString('latin1');
  const start = text.indexOf('xref');
  const broken = Buffer.from(
    text.slice(0, start) + text.slice(start).replace(/\d{10}/g, '0000000009'),
    'latin1',
  );

  const result = pdfimg.extractPageImages(broken);
  assert.strictEqual(result.ok, true, 'the page should still be found');
  assert.strictEqual(result.pages[0].width, width);
});

test('something that is not a PDF is refused without throwing', () => {
  const result = pdfimg.extractPageImages(Buffer.from('this is not a PDF at all'));
  assert.strictEqual(result.ok, false);
  assert.strictEqual(result.pages.length, 0);
});

test('ASCII85 decoding matches the reference encoding', () => {
  // "Hello world" encodes to exactly this (checked against a reference
  // implementation, not written from memory).
  assert.strictEqual(
    pdfimg.decodeAscii85(Buffer.from('87cURD]j7BEbo7', 'latin1')).toString('latin1'),
    'Hello world',
  );
  // The <~ ~> wrapper a PDF puts around it must be tolerated.
  assert.strictEqual(
    pdfimg.decodeAscii85(Buffer.from('<~87cURD]j7BEbo7~>', 'latin1')).toString('latin1'),
    'Hello world',
  );
  // 'z' is the shorthand for four zero bytes.
  assert.deepStrictEqual(
    Array.from(pdfimg.decodeAscii85(Buffer.from('z', 'latin1'))),
    [0, 0, 0, 0],
  );
});

test('the PNG predictor is undone correctly', () => {
  // Two rows of four bytes, both using filter 2 (Up).
  const data = Buffer.from([
    0, 10, 20, 30, 40,      // filter 0: literal
    2, 1, 1, 1, 1,          // filter 2: add the row above
  ]);
  const out = pdfimg.undoPngPredictor(data, 1, 8, 4);
  assert.deepStrictEqual(Array.from(out), [10, 20, 30, 40, 11, 21, 31, 41]);
});
