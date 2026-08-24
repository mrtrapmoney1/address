'use strict';

const test = require('node:test');
const assert = require('node:assert');

const pp = require('../lib/preprocess.js');

function grey(width, height, fill) {
  const data = new Uint8Array(width * height);
  if (typeof fill === 'function') {
    for (let y = 0; y < height; y += 1) {
      for (let x = 0; x < width; x += 1) data[y * width + x] = fill(x, y);
    }
  } else data.fill(fill);
  return { width, height, channels: 1, data };
}

/* A page of horizontal text lines - enough structure for skew to be real. */
function textPage(width, height) {
  return grey(width, height, (x, y) => {
    const inLine = (y % 20) < 6;
    const inMargin = x < 20 || x > width - 20;
    return (inLine && !inMargin) ? 0 : 255;
  });
}

test('greyscale uses luma weights, not a flat average', () => {
  // Pure green is much brighter to the eye than pure blue.
  const rgb = { width: 2, height: 1, channels: 3, data: new Uint8Array([0, 255, 0, 0, 0, 255]) };
  const g = pp.toGrayscale(rgb);
  assert.strictEqual(g.channels, 1);
  assert.strictEqual(g.data[0], 150);   // 0.587 * 255
  assert.strictEqual(g.data[1], 29);    // 0.114 * 255
  assert.ok(g.data[0] > g.data[1], 'green must come out brighter than blue');
});

test('greyscale of an already-grey image is a copy, not the same buffer', () => {
  const src = grey(4, 4, 120);
  const g = pp.toGrayscale(src);
  g.data[0] = 5;
  assert.strictEqual(src.data[0], 120, 'the input must not be mutated');
});

test('Otsu finds the midpoint between two ink levels', () => {
  const img = grey(100, 100, (x, y) => (y < 20 ? 60 : 200));
  const t = pp.otsuThreshold(img);
  assert.ok(t > 100 && t < 160, 'expected a threshold between the two peaks, got ' + t);
});

test('Otsu on an already-binary page does not collapse to zero', () => {
  /*
   * With ink at 0 and paper at 255 every threshold scores identically, and
   * the naive implementation returns 0 - leaving no headroom, so a single
   * grey pixel from a rescale counts as paper.
   */
  const img = grey(50, 50, (x) => (x < 10 ? 0 : 255));
  const t = pp.otsuThreshold(img);
  assert.ok(t > 100, 'expected a mid-range threshold, got ' + t);
});

test('contrast stretch opens a washed-out page to the full range', () => {
  const img = grey(100, 100, (x, y) => (y < 50 ? 100 : 160));
  const out = pp.stretchContrast(img, 0.5);
  let min = 255;
  let max = 0;
  for (let i = 0; i < out.data.length; i += 1) {
    if (out.data[i] < min) min = out.data[i];
    if (out.data[i] > max) max = out.data[i];
  }
  assert.strictEqual(min, 0);
  assert.strictEqual(max, 255);
});

test('Sauvola survives a lighting gradient that defeats one global threshold', () => {
  /*
   * Text of a constant darkness under a shadow that darkens the page from
   * left to right. A single threshold must lose one end; a local one must not.
   */
  const width = 200;
  const height = 100;
  /*
   * The shadow has to be deep enough that PAPER on the dark side is darker
   * than INK on the light side. Anything milder and one global threshold
   * copes fine, and the test would prove nothing.
   *   left edge:  paper 245, ink 90
   *   right edge: paper  45, ink  0
   * No single cut separates ink from paper across that page.
   */
  const img = grey(width, height, (x, y) => {
    const shade = Math.round((x / width) * 200);
    const isInk = (y % 20) < 6 && (x % 12) < 7;
    return Math.max(0, (isInk ? 90 : 245) - shade);
  });

  const otsu = pp.binarize(img, pp.otsuThreshold(img));
  const sauvola = pp.sauvolaThreshold(img);

  // Count ink in the darkest quarter of the page.
  const inkIn = (bin) => {
    let n = 0;
    for (let y = 0; y < height; y += 1) {
      for (let x = Math.floor(width * 0.75); x < width; x += 1) {
        if (bin.data[y * width + x] < 128) n += 1;
      }
    }
    return n;
  };

  const total = width * 0.25 * height;
  const otsuInk = inkIn(otsu) / total;
  const sauvolaInk = inkIn(sauvola) / total;
  // Otsu floods the shadowed quarter with ink; Sauvola keeps text-like density.
  assert.ok(otsuInk > 0.9, 'the global threshold should have flooded, got ' + otsuInk);
  assert.ok(sauvolaInk < 0.6, 'the local threshold should not flood, got ' + sauvolaInk);
});

test('upscaling to a target DPI multiplies the pixels correctly', () => {
  const img = grey(100, 50, 255);
  const out = pp.scaleToDpi(img, 150, 300);
  assert.strictEqual(out.scale, 2);
  assert.strictEqual(out.image.width, 200);
  assert.strictEqual(out.image.height, 100);
  assert.strictEqual(out.effectiveDpi, 300);
});

test('maxDimension stops a huge scan exhausting memory', () => {
  const img = grey(4000, 3000, 255);
  const out = pp.scaleToDpi(img, 150, 600, 6000);
  assert.ok(out.image.width <= 6000, 'width should be capped');
  assert.ok(out.scale < 4, 'the scale should have been reduced from 4x');
});

test('downscaling averages rather than dropping pixels', () => {
  // A one-pixel checkerboard: point sampling gives all-black or all-white,
  // area averaging gives mid-grey.
  const img = grey(64, 64, (x, y) => (((x + y) % 2) === 0 ? 0 : 255));
  const small = pp.resample(img, 16, 16);
  let sum = 0;
  for (let i = 0; i < small.data.length; i += 1) sum += small.data[i];
  const mean = sum / small.data.length;
  assert.ok(mean > 100 && mean < 155, 'expected mid-grey, got ' + mean);
});

test('skew is measured to a tenth of a degree', () => {
  const page = pp.binarize(textPage(400, 300), 128);
  assert.strictEqual(pp.estimateSkew(page), 0);

  for (const angle of [-3, -1.5, 2, 4]) {
    const tilted = pp.binarize(pp.rotate(page, angle, 255), 128);
    const measured = pp.estimateSkew(tilted);
    assert.ok(
      Math.abs(measured - angle) <= 0.4,
      'tilted by ' + angle + ' but measured ' + measured,
    );
  }
});

test('a blank page reports no skew instead of a random angle', () => {
  assert.strictEqual(pp.estimateSkew(grey(200, 200, 255)), 0);
});

test('rotation grows the canvas so no corner is lost', () => {
  const img = grey(100, 100, 0);
  const out = pp.rotate(img, 45, 255);
  assert.ok(out.width > 100 && out.height > 100, 'the canvas should grow to fit');
  // The corners of the new canvas are outside the original square: white.
  assert.strictEqual(out.data[0], 255);
});

test('despeckle removes dust but keeps real glyphs', () => {
  const width = 100;
  const height = 100;
  const data = new Uint8Array(width * height).fill(255);
  // A solid 10x10 block: a glyph.
  for (let y = 10; y < 20; y += 1) for (let x = 10; x < 20; x += 1) data[y * width + x] = 0;
  // Single isolated pixels: dust.
  data[50 * width + 50] = 0;
  data[60 * width + 70] = 0;

  const out = pp.despeckle({ width, height, channels: 1, data }, 3);
  assert.strictEqual(out.data[50 * width + 50], 255, 'dust should be gone');
  assert.strictEqual(out.data[60 * width + 70], 255, 'dust should be gone');
  assert.strictEqual(out.data[15 * width + 15], 0, 'the glyph must survive');
});

test('cropToContent trims a black scanner edge', () => {
  const width = 100;
  const height = 100;
  const data = new Uint8Array(width * height).fill(255);
  for (let y = 0; y < height; y += 1) for (let x = 0; x < 4; x += 1) data[y * width + x] = 0;  // edge
  for (let y = 40; y < 50; y += 1) for (let x = 40; x < 60; x += 1) data[y * width + x] = 0;   // content

  const out = pp.cropToContent({ width, height, channels: 1, data }, 2);
  assert.ok(out.box.x >= 30, 'the black edge should have been trimmed off, box.x=' + out.box.x);
  assert.ok(out.image.width < width);
});

test('the coordinate transform inverts exactly, including deskew', () => {
  const page = textPage(400, 300);
  const tilted = pp.rotate(page, 2.5, 255);
  const prepared = pp.preparePage(tilted, { sourceDpi: 150, targetDpi: 300 });

  // The centre of the processed image must map back to the centre of the input.
  const centre = pp.mapPointToSource(
    prepared.transform, prepared.image.width / 2, prepared.image.height / 2,
  );
  assert.ok(Math.abs(centre.x - tilted.width / 2) < 1.5, 'x off by ' + (centre.x - tilted.width / 2));
  assert.ok(Math.abs(centre.y - tilted.height / 2) < 1.5, 'y off by ' + (centre.y - tilted.height / 2));
});

test('mapBoxToSource returns a box containing all four mapped corners', () => {
  const transform = {
    scale: 2, rotateDegrees: 0, cropX: 10, cropY: 20,
    preRotate: { width: 100, height: 100 }, postRotate: { width: 100, height: 100 },
  };
  const box = pp.mapBoxToSource(transform, 100, 200, 40, 20);
  assert.strictEqual(box.x, 55);   // (100 + 10) / 2
  assert.strictEqual(box.y, 110);  // (200 + 20) / 2
  assert.strictEqual(box.w, 20);
  assert.strictEqual(box.h, 10);
});

test('preparePage reports every step it took', () => {
  const prepared = pp.preparePage(textPage(200, 200), { sourceDpi: 100, targetDpi: 300 });
  assert.ok(prepared.steps.indexOf('grayscale') >= 0);
  assert.ok(prepared.steps.some((s) => s.indexOf('scale-to-300dpi') === 0), prepared.steps.join(','));
  assert.ok(prepared.steps.indexOf('sauvola') >= 0);
  assert.strictEqual(prepared.transform.effectiveDpi, 300);
});
