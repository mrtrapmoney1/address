'use strict';

/*
 * invoice.js - the part that knows it is reading an invoice.
 *
 * Everything else in this toolkit is deliberately document-agnostic: it finds
 * words, lines and columns on any page. This file is the opposite. It assumes
 * an invoice, and that assumption buys three things a general OCR pass cannot
 * give you:
 *
 *   1. NUMBERS COME OUT RIGHT. Tesseract confuses O with 0, l and I with 1,
 *      S with 5, B with 8. In prose that is a typo you can read past. In
 *      "1,S00.OO" it is a number that will not parse, and a total that is
 *      wrong by an order of magnitude is worse than a total that is missing.
 *      Inside a token that is unmistakably an amount, those confusions are
 *      one-directional and safe to undo.
 *
 *   2. COLUMNS GET MEANING. A column whose cells nearly all parse as money is
 *      an amount column - so the line-item table can be read as figures
 *      rather than as text that happens to contain digits.
 *
 *   3. THE FIELDS THAT MATTER ARE FOUND BY POSITION. "Invoice Date" and the
 *      date itself are two separate words on the page; what joins them is
 *      that one sits to the right of, or directly below, the other. That is a
 *      geometric question, and geometry is exactly what this pipeline kept.
 *
 * Nothing here ever silently changes a value. A repair is recorded, and an
 * arithmetic check that fails flags the invoice rather than "fixing" it.
 */

/* ------------------------------------------------------------------ labels */

/*
 * What a field can be called. Real invoices use all of these and more; the
 * list is matched loosely (case and punctuation are stripped) so "INV. NO:"
 * and "Invoice Number" reach the same place.
 */
const LABELS = {
  invoiceNumber: [
    'invoice number', 'invoice no', 'invoice num', 'invoice #', 'invoice',
    'inv no', 'inv num', 'inv #', 'inv', 'invoiceno', 'bill number', 'bill no',
    'document number', 'document no', 'reference', 'ref no', 'our reference',
    'tax invoice', 'credit note', 'statement no',
  ],
  invoiceDate: [
    'invoice date', 'date of invoice', 'inv date', 'date issued', 'issue date',
    'issued', 'date', 'dated', 'billing date', 'document date', 'tax point',
  ],
  dueDate: [
    'due date', 'payment due', 'date due', 'pay by', 'due by', 'due',
    'payment due date', 'net due',
  ],
  poNumber: [
    'po number', 'po no', 'po #', 'p o number', 'purchase order',
    'purchase order no', 'purchase order number', 'your order', 'order no',
    'order number', 'customer po', 'job number',
  ],
  total: [
    'total', 'total due', 'amount due', 'balance due', 'invoice total',
    'total amount', 'total amount due', 'grand total', 'total payable',
    'please pay', 'amount payable', 'net payable', 'total inc vat',
    'total incl tax', 'total including tax',
  ],
  subtotal: [
    'subtotal', 'sub total', 'sub-total', 'net', 'net amount', 'net total',
    'goods total', 'total excluding tax', 'total ex vat', 'total excl tax',
    'merchandise total',
  ],
  tax: [
    'tax', 'vat', 'gst', 'hst', 'pst', 'sales tax', 'tax amount', 'vat amount',
    'total tax', 'tax total', 'gst amount', 'use tax',
  ],
  taxRate: [
    'tax rate', 'vat rate', 'gst rate', 'rate', 'vat %', 'tax %',
  ],
  freight: [
    'freight', 'shipping', 'delivery', 'carriage', 'postage', 'shipping and handling',
    'freight charge', 'delivery charge', 's and h',
  ],
  discount: [
    'discount', 'less discount', 'discount amount', 'trade discount', 'rebate',
  ],
  terms: [
    'terms', 'payment terms', 'terms of payment', 'net terms',
  ],
  accountNumber: [
    'account number', 'account no', 'account #', 'customer number',
    'customer no', 'customer id', 'client number', 'account',
  ],
};

/* Words that mark the head of a line-item table. */
const TABLE_HEADINGS = [
  'description', 'item', 'qty', 'quantity', 'unit price', 'price', 'amount',
  'line total', 'ext price', 'extended', 'part', 'part no', 'sku', 'code',
  'product', 'rate', 'units', 'each',
];

/*
 * Normalise a label for comparison: lower case, no punctuation, single
 * spaces. "INV. NO:" and "Invoice Number" both reduce to something matchable.
 */
function labelKey(text) {
  return String(text)
    .toLowerCase()
    .replace(/[^a-z0-9%# ]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

/* --------------------------------------------------------------- numbers */

const CURRENCY = '\\$\\u00a3\\u20ac\\u00a5';

/*
 * Does this token WANT to be a number?
 *
 * A first pass that has to be generous: the whole point is to catch
 * "1,S00.OO", which does NOT parse as a number, so any test based on how much
 * of it is already digits fails on exactly the tokens that need repairing.
 * (An earlier version scored digits as a fraction of length and rejected
 * "1,S00.OO" at 3 digits out of 8 - it turned down its own headline example.)
 *
 * So this asks a weaker question - is every character either a digit, one of
 * the letters that gets confused for a digit, or a separator, and is there at
 * least one real digit? - and the decision about whether a repair is
 * justified is left to repairNumeric, which can check its own work.
 */
function looksNumeric(text) {
  const trimmed = String(text).trim();
  if (trimmed.length === 0) return false;

  const body = trimmed
    .replace(new RegExp('^[' + CURRENCY + '+\\-(]+'), '')
    .replace(/[)\-]+$/, '')
    .replace(/(CR|DR)$/i, '')
    .trim();

  if (body.length === 0) return false;
  // Without a real digit somewhere it is a word, whatever its letters are.
  if (!/[0-9]/.test(body)) return false;
  return /^[0-9OoQDlIiSsBZzGg|.,'\s]+$/.test(body);
}

/*
 * The confusions, and only in this direction.
 *
 * O -> 0 is safe inside a number and catastrophic outside it.
 * The reverse mapping is never applied: a digit in a numeric token is already
 * what it should be.
 */
const CONFUSIONS = {
  O: '0', o: '0', Q: '0', D: '0',
  l: '1', I: '1', i: '1', '|': '1',
  S: '5', s: '5',
  B: '8',
  Z: '2', z: '2',
  G: '6',
};

/*
 * Repair a numeric token, and CHECK THE REPAIR BEFORE ACCEPTING IT.
 *
 * The rule is simple and self-validating: a substitution is kept only if the
 * result parses as a number and the original did not. That makes it
 * impossible for this to mangle a token it merely guessed wrong about - if
 * the "repair" does not produce something readable, nothing changes.
 *
 * Returns { text, changed, from } so the repair can be recorded. A value that
 * was silently altered is a value nobody can audit.
 */
function repairNumeric(text) {
  const original = String(text);
  if (!looksNumeric(original)) return { text: original, changed: false };

  // Already a number: there is nothing to fix.
  if (parseMoney(original) !== null) return { text: original, changed: false };

  let candidate = '';
  for (let i = 0; i < original.length; i += 1) {
    const ch = original.charAt(i);
    candidate += CONFUSIONS[ch] !== undefined ? CONFUSIONS[ch] : ch;
  }
  // A lone apostrophe is a thousands separator in some locales.
  candidate = candidate.replace(/'/g, ',');

  if (candidate === original) return { text: original, changed: false };
  if (parseMoney(candidate) === null) {
    // The substitution did not produce a readable number, so it was not a
    // misread of a number. Leave it exactly as the engine saw it.
    return { text: original, changed: false };
  }

  return { text: candidate, changed: true, from: original };
}

/*
 * Read a money value. Handles the separator conventions that actually turn up
 * on invoices: 1,234.56 and 1.234,56 and 1 234,56, plus a trailing minus, a
 * parenthesised negative, and a CR suffix.
 */
function parseMoney(text) {
  if (text === null || text === undefined) return null;
  let s = String(text).trim();
  if (s === '') return null;

  let negative = false;
  if (/^\(.*\)$/.test(s)) { negative = true; s = s.slice(1, -1); }
  if (/-$/.test(s)) { negative = true; s = s.replace(/-$/, ''); }
  if (/^-/.test(s)) { negative = true; s = s.replace(/^-/, ''); }
  if (/CR$/i.test(s)) { negative = true; s = s.replace(/CR$/i, ''); }
  s = s.replace(/DR$/i, '');

  s = s.replace(new RegExp('[' + CURRENCY + ']', 'g'), '').replace(/\s/g, '').trim();
  if (s === '' || !/[0-9]/.test(s)) return null;
  if (!/^[0-9.,]+$/.test(s)) return null;

  const lastComma = s.lastIndexOf(',');
  const lastDot = s.lastIndexOf('.');

  if (lastComma >= 0 && lastDot >= 0) {
    // Whichever comes last is the decimal separator.
    if (lastComma > lastDot) s = s.replace(/\./g, '').replace(',', '.');
    else s = s.replace(/,/g, '');
  } else if (lastComma >= 0) {
    const after = s.length - lastComma - 1;
    // "1,234" is thousands; "1,23" is a decimal comma.
    if (after === 3 && s.indexOf(',') === lastComma && /^\d{1,3},\d{3}$/.test(s)) s = s.replace(/,/g, '');
    else if (after === 3) s = s.replace(/,/g, '');
    else s = s.replace(',', '.');
  } else if (lastDot >= 0) {
    const after = s.length - lastDot - 1;
    // "1.234" with nothing else is ambiguous; three decimals on an invoice
    // is far more likely to be a thousands separator than a milli-unit.
    if (after === 3 && /^\d{1,3}\.\d{3}$/.test(s)) s = s.replace(/\./g, '');
  }

  const value = parseFloat(s);
  if (isNaN(value)) return null;
  return negative ? -value : value;
}

/* A date, in the formats invoices actually use. */
const MONTHS = {
  jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6,
  jul: 7, aug: 8, sep: 9, sept: 9, oct: 10, nov: 11, dec: 12,
};

function parseDate(text) {
  if (!text) return null;
  const s = String(text).trim();

  // 2026-08-24
  let m = s.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$/);
  if (m) return normaliseDate(+m[1], +m[2], +m[3]);

  // 24/08/2026 or 08/24/2026 - genuinely ambiguous, resolved below
  m = s.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})$/);
  if (m) {
    let year = +m[3];
    if (year < 100) year += year < 70 ? 2000 : 1900;
    const a = +m[1];
    const b = +m[2];
    // If one of them cannot be a month, that settles it. Otherwise assume
    // day-first, and record the ambiguity rather than pretending.
    if (a > 12 && b <= 12) return normaliseDate(year, b, a);
    if (b > 12 && a <= 12) return normaliseDate(year, a, b);
    const date = normaliseDate(year, b, a);
    if (date) date.ambiguous = true;
    return date;
  }

  // 24 Aug 2026 / Aug 24, 2026 / 24Aug2026
  m = s.match(/^(\d{1,2})\s*[-\s]?\s*([A-Za-z]{3,9})\.?\s*[-,\s]?\s*(\d{2,4})$/);
  if (m && MONTHS[m[2].toLowerCase().slice(0, 4)] !== undefined
      || (m && MONTHS[m[2].toLowerCase().slice(0, 3)] !== undefined)) {
    const month = MONTHS[m[2].toLowerCase().slice(0, 4)] || MONTHS[m[2].toLowerCase().slice(0, 3)];
    let year = +m[3];
    if (year < 100) year += year < 70 ? 2000 : 1900;
    return normaliseDate(year, month, +m[1]);
  }

  m = s.match(/^([A-Za-z]{3,9})\.?\s+(\d{1,2}),?\s+(\d{2,4})$/);
  if (m) {
    const month = MONTHS[m[1].toLowerCase().slice(0, 4)] || MONTHS[m[1].toLowerCase().slice(0, 3)];
    if (month) {
      let year = +m[3];
      if (year < 100) year += year < 70 ? 2000 : 1900;
      return normaliseDate(year, month, +m[2]);
    }
  }

  return null;
}

function normaliseDate(year, month, day) {
  if (month < 1 || month > 12 || day < 1 || day > 31) return null;
  if (year < 1990 || year > 2100) return null;
  const iso = year + '-' + String(month).padStart(2, '0') + '-' + String(day).padStart(2, '0');
  return { year, month, day, iso, ambiguous: false };
}

/*
 * An invoice number: mostly digits or a mixed alphanumeric code, not a date,
 * not an amount. The exclusions matter more than the inclusions - the usual
 * failure is reporting the invoice DATE or the TOTAL as the invoice number.
 */
function looksLikeInvoiceNumber(text) {
  const s = String(text).trim();
  if (s.length < 3 || s.length > 30) return false;
  if (parseDate(s)) return false;
  // A decimal point means it is money, not a reference.
  if (/^[0-9,]+\.[0-9]{2}$/.test(s)) return false;
  if (!/[0-9]/.test(s)) return false;
  if (!/^[A-Za-z0-9][A-Za-z0-9\-/_.#]*$/.test(s)) return false;
  return true;
}

/* ------------------------------------------------------ column classification */

/*
 * What kind of column is this? Decided by what most of its cells parse as,
 * ignoring the header row and blanks.
 */
function classifyColumns(region) {
  const counts = [];
  for (let c = 0; c < region.columnCount; c += 1) {
    counts.push({ money: 0, date: 0, integer: 0, text: 0, filled: 0 });
  }

  for (let l = 0; l < region.lines.length; l += 1) {
    const cells = [];
    for (let c = 0; c < region.columnCount; c += 1) cells.push([]);
    const ws = region.lines[l].words;
    for (let w = 0; w < ws.length; w += 1) {
      const col = ws[w].column === undefined ? 0 : Math.min(region.columnCount - 1, ws[w].column);
      cells[col].push(ws[w].text);
    }

    for (let c = 0; c < region.columnCount; c += 1) {
      const text = cells[c].join(' ').trim();
      if (text === '') continue;
      counts[c].filled += 1;

      /*
       * Classify what the cell WOULD be once its digits are repaired, not
       * what it currently reads as. Otherwise the ordering defeats itself: a
       * column of amounts containing one misread "45O.OO" scores below the
       * threshold, is called text, is therefore never repaired, and stays
       * wrong. Asking the repaired form breaks that circle.
       */
      const repaired = repairNumeric(text).text;

      if (parseDate(text)) counts[c].date += 1;
      else if (/^[0-9]{1,6}$/.test(repaired)) counts[c].integer += 1;
      else if (parseMoney(repaired) !== null && /[.,]/.test(repaired)) counts[c].money += 1;
      else counts[c].text += 1;
    }
  }

  return counts.map(function (count, index) {
    let kind = 'text';
    if (count.filled >= 2) {
      // A clear majority decides; a mixed column stays text, which is the
      // honest answer rather than a guess.
      if (count.money / count.filled >= 0.6) kind = 'amount';
      else if (count.date / count.filled >= 0.6) kind = 'date';
      else if (count.integer / count.filled >= 0.6) kind = 'number';
    }
    return { index, kind, filled: count.filled, counts: count };
  });
}

/* ----------------------------------------------------------- field finding */

function wordsOf(page) {
  const out = [];
  for (let r = 0; r < page.regions.length; r += 1) {
    for (let l = 0; l < page.regions[r].lines.length; l += 1) {
      const line = page.regions[r].lines[l];
      for (let w = 0; w < line.words.length; w += 1) {
        out.push({ word: line.words[w], line, region: page.regions[r] });
      }
    }
  }
  return out;
}

/*
 * Find where a label sits, allowing it to be several words long.
 *
 * Labels are matched over runs of up to four consecutive words on one line,
 * longest first, because "Total" and "Total Amount Due" are different fields
 * and matching the short one first would take the wrong value.
 */
function findLabels(page) {
  const found = [];
  const fieldNames = Object.keys(LABELS);

  for (let r = 0; r < page.regions.length; r += 1) {
    const region = page.regions[r];
    for (let l = 0; l < region.lines.length; l += 1) {
      const words = region.lines[l].words;

      for (let start = 0; start < words.length; start += 1) {
        for (let length = Math.min(4, words.length - start); length >= 1; length -= 1) {
          const span = words.slice(start, start + length);
          const key = labelKey(span.map(function (w) { return w.text; }).join(' '));
          if (key === '') continue;

          for (let f = 0; f < fieldNames.length; f += 1) {
            const field = fieldNames[f];
            if (LABELS[field].indexOf(key) < 0) continue;

            found.push({
              field,
              key,
              words: span,
              length,
              line: region.lines[l],
              lineIndex: l,
              region,
              x0: span[0].x,
              x1: span[span.length - 1].x + span[span.length - 1].w,
              y0: Math.min.apply(null, span.map(function (w) { return w.y; })),
              y1: Math.max.apply(null, span.map(function (w) { return w.y + w.h; })),
              after: words.slice(start + length),
            });
            // Longest match at this position wins; stop shortening.
            start += length - 1;
            length = 0;
            break;
          }
          if (length === 0) break;
        }
      }
    }
  }
  return found;
}

/*
 * Take the value belonging to a label.
 *
 * On the right, on the same line, is the common case. Directly below is the
 * other one - a column heading with its value under it. Both are tried, and
 * the candidate is validated against what the field is supposed to look like,
 * so "Invoice Date: Net 30" does not become a date.
 */
function valueForLabel(label, page, validator) {
  const sameLine = label.after;
  for (let i = 0; i < sameLine.length; i += 1) {
    // Join up to three words: "1,234.56", "24 Aug 2026".
    for (let take = Math.min(3, sameLine.length - i); take >= 1; take -= 1) {
      const span = sameLine.slice(i, i + take);
      const text = span.map(function (w) { return w.text; }).join(' ');
      const value = validator(text);
      if (value !== null && value !== false) {
        return {
          text, value, words: span, where: 'right',
          confidence: averageConf(span),
        };
      }
    }
  }

  /*
   * Below: a word whose horizontal span overlaps the label's, on one of the
   * next couple of lines. Overlap rather than "nearest x" is what stops the
   * value from the column alongside being picked up.
   */
  const lines = label.region.lines;
  for (let l = label.lineIndex + 1; l < Math.min(label.lineIndex + 3, lines.length); l += 1) {
    const words = lines[l].words;
    for (let i = 0; i < words.length; i += 1) {
      const overlap = Math.min(words[i].x + words[i].w, label.x1) - Math.max(words[i].x, label.x0);
      const width = Math.min(words[i].w, label.x1 - label.x0);
      if (width <= 0 || overlap / width < 0.4) continue;

      for (let take = Math.min(3, words.length - i); take >= 1; take -= 1) {
        const span = words.slice(i, i + take);
        const text = span.map(function (w) { return w.text; }).join(' ');
        const value = validator(text);
        if (value !== null && value !== false) {
          return { text, value, words: span, where: 'below', confidence: averageConf(span) };
        }
      }
    }
  }

  return null;
}

function averageConf(words) {
  let total = 0;
  let counted = 0;
  for (let i = 0; i < words.length; i += 1) {
    if (typeof words[i].conf === 'number') { total += words[i].conf; counted += 1; }
  }
  return counted === 0 ? null : Math.round((total / counted) * 10) / 10;
}

/* Validators, one per field shape. */
const VALIDATORS = {
  money: function (text) {
    const value = parseMoney(repairNumeric(text).text);
    return value === null ? null : value;
  },
  date: function (text) {
    return parseDate(text);
  },
  reference: function (text) {
    const trimmed = String(text).trim().replace(/^[:#]\s*/, '');
    return looksLikeInvoiceNumber(trimmed) ? trimmed : null;
  },
  percent: function (text) {
    const m = String(text).match(/^([0-9]{1,2}(?:[.,][0-9]{1,3})?)\s*%?$/);
    if (!m) return null;
    const value = parseFloat(m[1].replace(',', '.'));
    return (value >= 0 && value <= 100) ? value : null;
  },
  text: function (text) {
    const trimmed = String(text).trim();
    return trimmed.length > 0 ? trimmed : null;
  },
};

const FIELD_KINDS = {
  invoiceNumber: 'reference',
  invoiceDate: 'date',
  dueDate: 'date',
  poNumber: 'reference',
  total: 'money',
  subtotal: 'money',
  tax: 'money',
  taxRate: 'percent',
  freight: 'money',
  discount: 'money',
  terms: 'text',
  accountNumber: 'reference',
};

/*
 * Pull the invoice's headline fields off the page.
 *
 * When two labels claim the same field - and they do, because "Total" appears
 * at the foot of the line items AND next to the amount due - the one with the
 * more specific label wins, and failing that the one further down the page,
 * because the payable total is the last one printed.
 */
function findFields(page) {
  const labels = findLabels(page);
  const fields = {};

  for (let i = 0; i < labels.length; i += 1) {
    const label = labels[i];
    const kind = FIELD_KINDS[label.field];
    const found = valueForLabel(label, page, VALIDATORS[kind]);
    if (found === null) continue;

    const candidate = {
      field: label.field,
      label: label.words.map(function (w) { return w.text; }).join(' '),
      labelKey: label.key,
      text: found.text,
      value: found.value,
      where: found.where,
      confidence: found.confidence,
      y: label.y0,
      x: label.x0,
      specificity: label.key.length,
    };

    const existing = fields[label.field];
    if (!existing) {
      fields[label.field] = candidate;
    } else if (candidate.specificity > existing.specificity) {
      fields[label.field] = candidate;
    } else if (candidate.specificity === existing.specificity && candidate.y > existing.y) {
      fields[label.field] = candidate;
    }
  }

  return fields;
}

/*
 * Does the invoice add up?
 *
 *     subtotal + freight - discount + tax = total
 *
 * When it does, every one of those numbers has been confirmed by the others,
 * which is far stronger evidence than any per-character confidence score.
 * When it does not, the invoice is FLAGGED and nothing is changed - a
 * silently "corrected" total is exactly the failure that destroys trust in a
 * tool like this.
 */
function checkArithmetic(fields) {
  const value = function (name) {
    return fields[name] && typeof fields[name].value === 'number' ? fields[name].value : null;
  };

  const subtotal = value('subtotal');
  const tax = value('tax');
  const total = value('total');
  const freight = value('freight');
  const discount = value('discount');

  if (subtotal === null || total === null) {
    return { checked: false, reason: 'not enough figures to check' };
  }

  const computed = subtotal + (freight || 0) - (discount || 0) + (tax || 0);
  const difference = Math.round((total - computed) * 100) / 100;

  // A penny either way is rounding, not a misread.
  if (Math.abs(difference) <= 0.02) {
    return { checked: true, ok: true, computed, difference: 0 };
  }

  return {
    checked: true,
    ok: false,
    computed: Math.round(computed * 100) / 100,
    difference,
    reason: 'the amounts do not add up: subtotal ' + subtotal.toFixed(2)
      + ' + freight ' + (freight || 0).toFixed(2)
      + ' - discount ' + (discount || 0).toFixed(2)
      + ' + tax ' + (tax || 0).toFixed(2)
      + ' = ' + computed.toFixed(2)
      + ', but the total reads ' + total.toFixed(2)
      + ' (out by ' + difference.toFixed(2) + ')',
  };
}

/*
 * Repair the numeric words on a page, in place, and report what changed.
 *
 * Only words inside a column classified as an amount, or that already look
 * numeric, are touched. A word in a description column is never altered, no
 * matter what it looks like.
 */
function repairPage(page) {
  const repairs = [];

  for (let r = 0; r < page.regions.length; r += 1) {
    const region = page.regions[r];
    const columns = classifyColumns(region);

    for (let l = 0; l < region.lines.length; l += 1) {
      const words = region.lines[l].words;
      for (let w = 0; w < words.length; w += 1) {
        const word = words[w];
        const column = columns[word.column === undefined ? 0 : word.column];
        const numericColumn = column && (column.kind === 'amount' || column.kind === 'number');

        if (!numericColumn && !looksNumeric(word.text)) continue;
        if (!looksNumeric(word.text)) continue;

        const repaired = repairNumeric(word.text);
        if (repaired.changed) {
          repairs.push({ from: repaired.from, to: repaired.text, column: word.column });
          word.originalText = repaired.from;
          word.text = repaired.text;
          word.repaired = true;
        }
      }
    }

    /*
     * Classify again now the repairs are in. The first pass had to reason
     * about what the cells could become; this one records what they actually
     * are, which is what the caller reads.
     */
    region.columnKinds = classifyColumns(region).map(function (c) { return c.kind; });
  }

  return repairs;
}

/*
 * The whole invoice read: repair the numbers, classify the columns, find the
 * fields, check the arithmetic.
 */
function readInvoice(page) {
  const repairs = repairPage(page);
  const fields = findFields(page);
  const arithmetic = checkArithmetic(fields);

  const notes = [];
  if (repairs.length > 0) {
    notes.push(repairs.length + ' numeric token(s) repaired');
  }
  if (arithmetic.checked && !arithmetic.ok) {
    notes.push(arithmetic.reason);
  }
  if (arithmetic.checked && arithmetic.ok) {
    notes.push('the amounts reconcile');
  }
  if (fields.invoiceDate && fields.invoiceDate.value && fields.invoiceDate.value.ambiguous) {
    notes.push('the invoice date is ambiguous (day/month order); read as '
      + fields.invoiceDate.value.iso);
  }

  const tables = [];
  for (let r = 0; r < page.regions.length; r += 1) {
    if (page.regions[r].kind === 'table') {
      tables.push({
        index: page.regions[r].index,
        columns: page.regions[r].columnKinds || [],
        rows: page.regions[r].lines.length,
      });
    }
  }

  return { fields, repairs, arithmetic, tables, notes };
}

module.exports = {
  LABELS,
  TABLE_HEADINGS,
  labelKey,
  looksNumeric,
  repairNumeric,
  parseMoney,
  parseDate,
  looksLikeInvoiceNumber,
  classifyColumns,
  findLabels,
  findFields,
  checkArithmetic,
  repairPage,
  readInvoice,
  wordsOf,
};
