'use strict';

const test = require('node:test');
const assert = require('node:assert');

const seg = require('../lib/segment.js');

/*
 * Word boxes are built by hand here rather than produced by OCR, so these
 * tests describe the GEOMETRY rules exactly and never depend on how well
 * Tesseract read a particular image.
 */
function word(text, x, y, w, h) {
  return { text, x, y, w: w === undefined ? text.length * 8 : w, h: h === undefined ? 14 : h, conf: 95 };
}

test('pageMetrics measures the page from its own words', () => {
  const words = [word('one', 0, 0, 30, 14), word('two', 40, 0, 30, 14), word('three', 0, 20, 50, 14)];
  const m = seg.pageMetrics(words);
  assert.strictEqual(m.textHeight, 14);
  assert.ok(m.charWidth > 0, 'a character width should be derivable');
});

test('words on one line are grouped by overlap, not by equal y', () => {
  // A capital and a comma on the same line have different tops and heights.
  const words = [
    word('Total', 0, 100, 40, 14),
    word('due', 50, 102, 25, 11),   // sits 2px lower, shorter
    word(',', 80, 108, 4, 6),       // a comma, much shorter and lower
    word('Next', 0, 130, 35, 14),   // genuinely the next line
  ];
  const lines = seg.groupLines(words);
  assert.strictEqual(lines.length, 2);
  assert.strictEqual(lines[0].words.length, 3);
  assert.strictEqual(seg.lineText(lines[0]), 'Total due ,');
  assert.strictEqual(seg.lineText(lines[1]), 'Next');
});

test('words in a line come back left to right whatever order they arrive in', () => {
  const words = [word('third', 200, 10, 40, 14), word('first', 0, 10, 40, 14), word('second', 100, 10, 40, 14)];
  const lines = seg.groupLines(words);
  assert.strictEqual(seg.lineText(lines[0]), 'first second third');
});

test('a wide white channel splits two independent blocks apart', () => {
  /*
   * Two blocks side by side whose rows do NOT line up - the right block
   * starts lower. This is the case XY-cut is for.
   */
  const words = [];
  for (let i = 0; i < 4; i += 1) words.push(word('left' + i, 0, i * 20, 60, 14));
  for (let i = 0; i < 4; i += 1) words.push(word('right' + i, 400, 200 + i * 20, 60, 14));
  const regions = seg.xyCut(words, {});
  assert.ok(regions.length >= 2, 'expected the two blocks to be separated');
});

test('a table gutter is NOT cut, because the rows carry on across it', () => {
  /*
   * The regression that drove this design. A statement has a huge white
   * channel between the description and the amounts. Cutting there severs
   * every row from its own numbers.
   */
  const words = [];
  for (let i = 0; i < 8; i += 1) {
    words.push(word('01Aug2018', 0, i * 20, 76, 14));
    words.push(word('Description', 100, i * 20, 88, 14));
    words.push(word('1234', 500, i * 20, 32, 14));   // far away, across a wide gap
    words.push(word('99.99', 600, i * 20, 40, 14));
  }
  const mid = 300;
  assert.ok(seg.rowSpanFraction(words, mid) > 0.9, 'every row should span the gap');

  const regions = seg.xyCut(words, {});
  assert.strictEqual(regions.length, 1, 'a table must survive as one region');
});

test('column boundaries never cut through a word', () => {
  /*
   * The bug this test pins down: a column that produced an anchor for its
   * left edge AND its right edge got a boundary through its middle, so
   * numbers printed one above the other landed in different columns.
   */
  const words = [];
  for (let i = 0; i < 6; i += 1) {
    words.push(word('Row' + i, 0, i * 20, 40, 14));
    // A right-aligned money column: different widths, same right edge at 300.
    const amount = i % 2 === 0 ? '800.04' : '1,823.34';
    const w = amount.length * 8;
    words.push({ text: amount, x: 300 - w, y: i * 20, w, h: 14, conf: 95 });
  }
  const lines = seg.groupLines(words);
  const metrics = seg.pageMetrics(words);
  const gutters = seg.findGutters(lines, metrics, {});

  for (let i = 0; i < gutters.length; i += 1) {
    for (let j = 0; j < words.length; j += 1) {
      const straddles = words[j].x + 1 < gutters[i] && gutters[i] < words[j].x + words[j].w - 1;
      assert.ok(!straddles, 'boundary ' + gutters[i] + ' cuts through "' + words[j].text + '"');
    }
  }

  const assignment = seg.assignColumns(lines, gutters);
  // Every amount must end up in the same column as every other amount.
  const amountColumns = new Set();
  for (let i = 0; i < lines.length; i += 1) {
    const last = lines[i].words[lines[i].words.length - 1];
    amountColumns.add(last.column);
  }
  assert.strictEqual(amountColumns.size, 1, 'the money column split in two');
  assert.strictEqual(assignment.columnCount, 2);
});

test('column indices are dense: no phantom empty columns', () => {
  const words = [];
  for (let i = 0; i < 5; i += 1) {
    words.push(word('a', 0, i * 20, 20, 14));
    words.push(word('b', 200, i * 20, 20, 14));
    words.push(word('c', 400, i * 20, 20, 14));
  }
  const lines = seg.groupLines(words);
  const metrics = seg.pageMetrics(words);
  const assignment = seg.assignColumns(lines, seg.findGutters(lines, metrics, {}));
  assert.strictEqual(assignment.columnCount, 3);

  const seen = new Set();
  for (let i = 0; i < lines.length; i += 1) {
    for (let j = 0; j < lines[i].words.length; j += 1) seen.add(lines[i].words[j].column);
  }
  assert.deepStrictEqual(Array.from(seen).sort(), [0, 1, 2]);
});

test('a table is recognised as a table and a paragraph is not', () => {
  const table = [];
  for (let i = 0; i < 6; i += 1) {
    table.push(word('date', 0, i * 20, 40, 14));
    table.push(word('desc', 100, i * 20, 40, 14));
    table.push(word('num', 220, i * 20, 30, 14));
    table.push(word('amt', 340, i * 20, 30, 14));
  }
  const tablePage = seg.segmentPage(table, {});
  assert.ok(
    tablePage.regions.some((r) => r.kind === 'table'),
    'a four-column grid should read as a table',
  );

  // Prose: words flow, with only the left margin shared.
  const prose = [];
  let y = 0;
  for (let i = 0; i < 6; i += 1) {
    let x = 0;
    for (let j = 0; j < 7; j += 1) {
      prose.push(word('word', x, y, 30 + ((i * j) % 17), 14));
      x += 40 + ((i + j) % 13);
    }
    y += 20;
  }
  const prosePage = seg.segmentPage(prose, {});
  assert.ok(
    prosePage.regions.every((r) => r.kind === 'text'),
    'flowing prose should not be reported as a table',
  );
});

test('tableRows lines the cells up by column', () => {
  const words = [];
  const amounts = ['10.00', '20.00', '30.00'];
  for (let i = 0; i < 3; i += 1) {
    words.push(word('2018-01-0' + (i + 1), 0, i * 20, 80, 14));
    words.push(word('Item', 120, i * 20, 40, 14));
    words.push(word(amounts[i], 300, i * 20, 40, 14));
  }
  const page = seg.segmentPage(words, {});
  const region = page.regions[0];
  const rows = seg.tableRows(region);
  assert.strictEqual(rows.length, 3);
  for (let i = 0; i < 3; i += 1) {
    assert.strictEqual(rows[i][rows[i].length - 1], amounts[i],
      'the amount should be the last cell of row ' + i);
  }
});

test('an empty page is handled without throwing', () => {
  const page = seg.segmentPage([], {});
  assert.strictEqual(page.regions.length, 0);
  assert.strictEqual(seg.renderLayout(page, {}), '');
});

test('words with no text or no size are ignored', () => {
  const page = seg.segmentPage([
    word('real', 0, 0, 40, 14),
    { text: '   ', x: 50, y: 0, w: 10, h: 14 },
    { text: 'zero', x: 70, y: 0, w: 0, h: 0 },
  ], {});
  let count = 0;
  page.regions.forEach((r) => r.lines.forEach((l) => { count += l.words.length; }));
  assert.strictEqual(count, 1);
});

test('renderLayout keeps neighbouring words apart', () => {
  const words = [word('Clearing', 0, 0, 60, 14), word('Cheque', 62, 0, 50, 14)];
  const page = seg.segmentPage(words, {});
  const view = seg.renderLayout(page, { columns: 20 });
  // The defect being guarded against is "ClearingCheque" - two words fused
  // into one because they wanted the same character cell. How MUCH space
  // ends up between them depends on the scale and does not matter.
  assert.ok(!/ClearingCheque/.test(view), 'words ran together: ' + JSON.stringify(view));
  assert.ok(/Clearing\s+Cheque/.test(view), 'both words should appear, separated: ' + JSON.stringify(view));
});
