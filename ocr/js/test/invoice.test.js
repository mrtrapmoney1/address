'use strict';

const test = require('node:test');
const assert = require('node:assert');

const inv = require('../lib/invoice.js');
const seg = require('../lib/segment.js');

function W(text, x, y, w) {
  return { text, x, y, w: w === undefined ? text.length * 9 : w, h: 16, conf: 92 };
}

/* A whole invoice, laid out the way one actually is. */
function invoicePage() {
  const words = [
    W('ACME', 60, 40), W('SUPPLIES', 120, 40), W('LTD', 230, 40),
    W('12', 60, 64), W('Depot', 85, 64), W('Road', 150, 64),
    W('Omaha', 60, 86), W('NE', 130, 86), W('68102', 165, 86),

    W('Invoice', 600, 40), W('Number', 680, 40), W('INV-2026-0184', 770, 40),
    W('Invoice', 600, 64), W('Date', 680, 64), W('24/08/2026', 770, 64),
    W('Due', 600, 86), W('Date', 650, 86), W('23/09/2026', 770, 86),
    W('PO', 600, 108), W('Number', 650, 108), W('PO-77321', 770, 108),

    W('Description', 60, 200), W('Qty', 400, 200), W('Unit', 500, 200),
    W('Price', 560, 200), W('Amount', 760, 200),
    W('Blue', 60, 230), W('widgets', 110, 230), W('10', 410, 230),
    W('45.00', 520, 230), W('45O.OO', 770, 230),
    W('Steel', 60, 256), W('bolts', 115, 256), W('25', 410, 256),
    W('12.00', 520, 256), W('300.00', 770, 256),
    W('Packing', 60, 282), W('crate', 140, 282), W('1', 415, 282),
    W('75.50', 520, 282), W('75.50', 770, 282),

    W('Subtotal', 600, 360), W('825.50', 770, 360),
    W('Freight', 600, 386), W('35.00', 770, 386),
    W('Tax', 600, 412), W('Rate', 640, 412), W('7.5%', 770, 412),
    W('Tax', 600, 438), W('64.54', 770, 438),
    W('Total', 600, 470), W('Due', 660, 470), W('925.04', 770, 470),
  ];
  return seg.segmentPage(words, { pageWidth: 900 });
}

/* ------------------------------------------------------------ money and dates */

test('money is read in every convention an invoice uses', () => {
  assert.strictEqual(inv.parseMoney('1,234.56'), 1234.56);
  assert.strictEqual(inv.parseMoney('1.234,56'), 1234.56);   // European
  assert.strictEqual(inv.parseMoney('1 234,56'), 1234.56);   // space separator
  assert.strictEqual(inv.parseMoney('$99.99'), 99.99);
  assert.strictEqual(inv.parseMoney('(500.00)'), -500);      // accounting negative
  assert.strictEqual(inv.parseMoney('1,500.00-'), -1500);    // trailing minus
  assert.strictEqual(inv.parseMoney('250.00CR'), -250);      // credit
  assert.strictEqual(inv.parseMoney('0.00'), 0);
});

test('things that are not money are refused', () => {
  assert.strictEqual(inv.parseMoney('Widgets'), null);
  assert.strictEqual(inv.parseMoney(''), null);
  assert.strictEqual(inv.parseMoney('Net 30'), null);
  assert.strictEqual(inv.parseMoney(null), null);
  assert.strictEqual(inv.parseMoney('INV-2026'), null);
});

test('dates are read, and an ambiguous one says so', () => {
  assert.strictEqual(inv.parseDate('2026-08-24').iso, '2026-08-24');
  assert.strictEqual(inv.parseDate('24 Aug 2026').iso, '2026-08-24');
  assert.strictEqual(inv.parseDate('Aug 24, 2026').iso, '2026-08-24');
  assert.strictEqual(inv.parseDate('01Aug2018').iso, '2018-08-01');

  // 24 cannot be a month, so this one is certain.
  assert.strictEqual(inv.parseDate('24/08/2026').ambiguous, false);
  assert.strictEqual(inv.parseDate('08/24/2026').iso, '2026-08-24');

  // Both could be either. It is read day-first AND flagged, rather than
  // guessed silently.
  const both = inv.parseDate('05/06/2026');
  assert.strictEqual(both.ambiguous, true);

  assert.strictEqual(inv.parseDate('Net 30'), null);
  assert.strictEqual(inv.parseDate('1234.56'), null);
});

/* --------------------------------------------------------------- digit repair */

test('misread digits are repaired inside numbers', () => {
  assert.strictEqual(inv.repairNumeric('1,S00.OO').text, '1,500.00');
  assert.strictEqual(inv.repairNumeric('I,234.56').text, '1,234.56');
  assert.strictEqual(inv.repairNumeric('8OO.04').text, '800.04');
  assert.strictEqual(inv.repairNumeric('l25.00').text, '125.00');
});

test('words are never touched, whatever letters they contain', () => {
  /*
   * The dangerous direction. O -> 0 inside "1,S00.OO" is a fix; the same
   * substitution inside "GOODS" destroys the word. Nothing without a digit
   * in it is ever considered.
   */
  for (const word of ['Widgets', 'GOODS', 'Invoice', 'ISBN', 'O', 'Boston', 'Item']) {
    const result = inv.repairNumeric(word);
    assert.strictEqual(result.changed, false, word + ' should not be altered');
    assert.strictEqual(result.text, word);
  }
});

test('a number that already parses is left exactly alone', () => {
  for (const value of ['125.00', '1,234.56', '0.00', '99']) {
    const result = inv.repairNumeric(value);
    assert.strictEqual(result.changed, false);
    assert.strictEqual(result.text, value);
  }
});

test('a repair that does not produce a number is not applied', () => {
  /*
   * The self-check that makes this safe: substitutions are kept only if the
   * result reads as a number. "OO.OO" has no digit to start with, so it is
   * not even a candidate, and nothing that fails to parse is ever written.
   */
  const result = inv.repairNumeric('OO.OO');
  assert.strictEqual(result.changed, false);
  assert.strictEqual(result.text, 'OO.OO');
});

test('a repair records what it changed, so it can be audited', () => {
  const result = inv.repairNumeric('45O.OO');
  assert.strictEqual(result.changed, true);
  assert.strictEqual(result.from, '45O.OO');
  assert.strictEqual(result.text, '450.00');
});

/* ------------------------------------------------------------------- fields */

test('the headline fields are found by their labels', () => {
  const fields = inv.findFields(invoicePage());

  assert.strictEqual(fields.invoiceNumber.value, 'INV-2026-0184');
  assert.strictEqual(fields.invoiceDate.value.iso, '2026-08-24');
  assert.strictEqual(fields.dueDate.value.iso, '2026-09-23');
  assert.strictEqual(fields.poNumber.value, 'PO-77321');
  assert.strictEqual(fields.subtotal.value, 825.5);
  assert.strictEqual(fields.freight.value, 35);
  assert.strictEqual(fields.tax.value, 64.54);
  assert.strictEqual(fields.taxRate.value, 7.5);
  assert.strictEqual(fields.total.value, 925.04);
});

test('the more specific label wins when two could match', () => {
  /*
   * "Total" and "Total Due" both appear in the label list, and an invoice
   * often prints both. Matching the short one first would take the wrong
   * number, so the longest label at a position is the one that counts.
   */
  const fields = inv.findFields(invoicePage());
  assert.strictEqual(fields.total.label, 'Total Due');
  assert.strictEqual(fields.total.value, 925.04);
});

test('a value is not taken just because it sits next to a label', () => {
  // "Invoice Date" followed by something that is not a date must not produce
  // a date - a wrong date is worse than a missing one.
  const page = seg.segmentPage([
    W('Invoice', 60, 40), W('Date', 140, 40), W('Net', 220, 40), W('30', 260, 40),
    W('Filler', 60, 70), W('text', 130, 70), W('here', 190, 70),
    W('More', 60, 100), W('filler', 120, 100), W('text', 190, 100),
  ], { pageWidth: 600 });
  const fields = inv.findFields(page);
  assert.ok(!fields.invoiceDate, 'no date should have been claimed');
});

test('a value below its label is found as well as one beside it', () => {
  const page = seg.segmentPage([
    W('Invoice', 400, 40), W('Number', 480, 40),
    W('INV-9001', 400, 70),
    W('something', 60, 40), W('else', 160, 40),
    W('another', 60, 70), W('line', 140, 70),
  ], { pageWidth: 700 });
  const fields = inv.findFields(page);
  assert.ok(fields.invoiceNumber, 'the number under the label should be found');
  assert.strictEqual(fields.invoiceNumber.value, 'INV-9001');
  assert.strictEqual(fields.invoiceNumber.where, 'below');
});

/* -------------------------------------------------------------- arithmetic */

test('an invoice that adds up is confirmed', () => {
  const result = inv.readInvoice(invoicePage());
  assert.strictEqual(result.arithmetic.checked, true);
  assert.strictEqual(result.arithmetic.ok, true);
  // 825.50 + 35.00 - 0 + 64.54 = 925.04
  assert.strictEqual(result.arithmetic.computed, 925.04);
});

test('an invoice that does not add up is flagged, and nothing is changed', () => {
  const fields = {
    subtotal: { value: 100 },
    tax: { value: 10 },
    total: { value: 999 },
  };
  const check = inv.checkArithmetic(fields);
  assert.strictEqual(check.ok, false);
  assert.strictEqual(check.computed, 110);
  assert.strictEqual(check.difference, 889);
  assert.ok(/do not add up/.test(check.reason));
  // The values themselves must be untouched.
  assert.strictEqual(fields.total.value, 999);
});

test('a penny of rounding is not called an error', () => {
  const check = inv.checkArithmetic({
    subtotal: { value: 100.00 }, tax: { value: 7.50 }, total: { value: 107.51 },
  });
  assert.strictEqual(check.ok, true);
});

test('too few figures means unchecked, not failed', () => {
  const check = inv.checkArithmetic({ total: { value: 100 } });
  assert.strictEqual(check.checked, false);
});

/* ----------------------------------------------------------------- columns */

test('the line-item columns are given their meaning', () => {
  const page = invoicePage();
  inv.readInvoice(page);

  const items = page.regions.filter(function (r) {
    return r.columnKinds && r.columnKinds.indexOf('amount') >= 0 && r.lines.length === 4;
  })[0];

  assert.ok(items, 'the line-item table should have been found');
  assert.deepStrictEqual(items.columnKinds, ['text', 'number', 'amount', 'amount']);
});

test('the line-item table is not severed by the gaps between its columns', () => {
  /*
   * The regression this whole design turns on. The item table has wide white
   * channels between Description, Qty, Price and Amount, and a totals block
   * below it that sits entirely on the right. An earlier version let that
   * totals block dilute the evidence and cut the table into three pieces,
   * putting the descriptions in one region and the amounts in another.
   */
  const page = invoicePage();
  // Run the invoice pass, so the row below is the finished article the
  // spreadsheet would receive - repaired numbers and all.
  inv.readInvoice(page);

  const items = page.regions.filter(function (r) {
    return r.lines.length === 4 && r.lines[0].words.some(function (w) {
      return w.text === 'Description';
    });
  })[0];

  assert.ok(items, 'the item table should be a single region');
  assert.strictEqual(items.columnCount, 4, 'with four columns');

  const rows = seg.tableRows(items);
  assert.strictEqual(rows[1][0], 'Blue widgets');
  assert.strictEqual(rows[1][1], '10');
  assert.strictEqual(rows[1][2], '45.00');
  assert.strictEqual(rows[1][3], '450.00');   // repaired from 45O.OO
});

test('reading an invoice repairs its numbers and reports them', () => {
  const page = invoicePage();
  const result = inv.readInvoice(page);
  assert.strictEqual(result.repairs.length, 1);
  assert.strictEqual(result.repairs[0].from, '45O.OO');
  assert.strictEqual(result.repairs[0].to, '450.00');
  assert.ok(result.notes.some(function (n) { return /reconcile/.test(n); }));
});

test('an empty page does not throw', () => {
  const result = inv.readInvoice(seg.segmentPage([], {}));
  assert.strictEqual(Object.keys(result.fields).length, 0);
  assert.strictEqual(result.repairs.length, 0);
  assert.strictEqual(result.arithmetic.checked, false);
});

test('an invoice number is told apart from a date and an amount', () => {
  assert.ok(inv.looksLikeInvoiceNumber('INV-2026-0184'));
  assert.ok(inv.looksLikeInvoiceNumber('770123'));
  assert.ok(!inv.looksLikeInvoiceNumber('24/08/2026'), 'a date is not an invoice number');
  assert.ok(!inv.looksLikeInvoiceNumber('1,234.56'), 'an amount is not an invoice number');
  assert.ok(!inv.looksLikeInvoiceNumber('Widgets'), 'a word with no digits is not one');
});
