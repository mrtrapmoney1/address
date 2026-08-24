'use strict';

const test = require('node:test');
const assert = require('node:assert');
const zlib = require('zlib');

const { decodePng, encodePng, probePng, crc32 } = require('../lib/png.js');

/* A tiny helper so the tests can build images without a fixture file. */
function greyImage(width, height, fill) {
  const data = new Uint8Array(width * height);
  if (typeof fill === 'function') {
    for (let y = 0; y < height; y += 1) {
      for (let x = 0; x < width; x += 1) data[y * width + x] = fill(x, y);
    }
  } else {
    data.fill(fill === undefined ? 255 : fill);
  }
  return { width, height, channels: 1, data };
}

test('crc32 matches the known PNG IEND value', () => {
  // The IEND chunk's CRC is a fixed, published number; if crc32 is wrong,
  // every file written is corrupt in a way no decoder here would notice.
  assert.strictEqual(crc32(Buffer.from('IEND', 'ascii')) >>> 0, 0xae426082);
});

test('grey image survives an encode/decode round trip exactly', () => {
  const src = greyImage(64, 40, (x, y) => (x * 3 + y * 7) & 0xff);
  const decoded = decodePng(encodePng(src));
  assert.strictEqual(decoded.width, 64);
  assert.strictEqual(decoded.height, 40);
  assert.strictEqual(decoded.channels, 1);
  assert.deepStrictEqual(Array.from(decoded.data), Array.from(src.data));
});

test('RGB image survives an encode/decode round trip exactly', () => {
  const width = 17;
  const height = 9;
  const data = new Uint8Array(width * height * 3);
  for (let i = 0; i < data.length; i += 1) data[i] = (i * 5) & 0xff;
  const decoded = decodePng(encodePng({ width, height, channels: 3, data }));
  assert.strictEqual(decoded.channels, 3);
  assert.deepStrictEqual(Array.from(decoded.data), Array.from(data));
});

test('probePng reads dimensions without inflating the image', () => {
  const buf = encodePng(greyImage(123, 45, 200));
  const info = probePng(buf);
  assert.strictEqual(info.width, 123);
  assert.strictEqual(info.height, 45);
  assert.strictEqual(info.bitDepth, 8);
  assert.strictEqual(info.colorType, 0);
});

test('pHYs records the DPI so Tesseract does not have to guess', () => {
  const buf = encodePng(greyImage(10, 10, 255), { dpi: 300 });
  const index = buf.indexOf(Buffer.from('pHYs', 'ascii'));
  assert.ok(index > 0, 'pHYs chunk should be present when a dpi is given');
  // 300 dpi = 11811 pixels per metre, rounded.
  assert.strictEqual(buf.readUInt32BE(index + 4), 11811);
  assert.strictEqual(buf[index + 12], 1, 'unit specifier should be metres');
});

test('a non-PNG buffer is refused rather than misread', () => {
  assert.throws(() => decodePng(Buffer.from('not a png at all')), /not a PNG/);
  assert.strictEqual(probePng(Buffer.from('nope')), null);
});

test('truncated image data is reported, not silently padded', () => {
  const good = encodePng(greyImage(20, 20, 128));
  // Rebuild with a deliberately short IDAT.
  const idatIndex = good.indexOf(Buffer.from('IDAT', 'ascii'));
  const broken = Buffer.from(good);
  broken.writeUInt32BE(4, idatIndex - 4);
  assert.throws(() => decodePng(broken));
});

test('all five scanline filters decode back to the same pixels', () => {
  /*
   * The encoder picks a filter per line by a heuristic, so a round trip does
   * not by itself prove every filter path works. This builds each filter
   * explicitly and checks the decoder reconstructs a known image.
   */
  const width = 8;
  const height = 5;
  const expected = new Uint8Array(width * height);
  for (let i = 0; i < expected.length; i += 1) expected[i] = (i * 11) & 0xff;

  for (let filter = 0; filter <= 4; filter += 1) {
    const raw = Buffer.alloc((width + 1) * height);
    let prev = null;
    for (let y = 0; y < height; y += 1) {
      const line = expected.subarray(y * width, (y + 1) * width);
      raw[y * (width + 1)] = filter;
      for (let x = 0; x < width; x += 1) {
        const a = x >= 1 ? line[x - 1] : 0;
        const b = prev !== null ? prev[x] : 0;
        const c = (prev !== null && x >= 1) ? prev[x - 1] : 0;
        let v;
        if (filter === 0) v = line[x];
        else if (filter === 1) v = line[x] - a;
        else if (filter === 2) v = line[x] - b;
        else if (filter === 3) v = line[x] - ((a + b) >> 1);
        else {
          const p = a + b - c;
          const pa = Math.abs(p - a);
          const pb = Math.abs(p - b);
          const pc = Math.abs(p - c);
          const pred = (pa <= pb && pa <= pc) ? a : (pb <= pc ? b : c);
          v = line[x] - pred;
        }
        raw[y * (width + 1) + 1 + x] = v & 0xff;
      }
      prev = line;
    }

    const ihdr = Buffer.alloc(13);
    ihdr.writeUInt32BE(width, 0);
    ihdr.writeUInt32BE(height, 4);
    ihdr[8] = 8;
    ihdr[9] = 0;
    const chunks = [Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])];
    const build = (type, data) => {
      const out = Buffer.alloc(data.length + 12);
      out.writeUInt32BE(data.length, 0);
      out.write(type, 4, 4, 'ascii');
      data.copy(out, 8);
      out.writeUInt32BE(crc32(out.slice(4, 8 + data.length)), 8 + data.length);
      return out;
    };
    chunks.push(build('IHDR', ihdr));
    chunks.push(build('IDAT', zlib.deflateSync(raw)));
    chunks.push(build('IEND', Buffer.alloc(0)));

    const decoded = decodePng(Buffer.concat(chunks));
    assert.deepStrictEqual(
      Array.from(decoded.data), Array.from(expected),
      'filter type ' + filter + ' did not reconstruct correctly',
    );
  }
});
