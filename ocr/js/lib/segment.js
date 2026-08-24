'use strict';

/*
 * segment.js - putting the page back together from a pile of words.
 *
 * THIS IS THE POINT OF THE WHOLE TOOL.
 *
 * Run an OCR engine over a bank statement and ask it for text, and you get
 * back lines like this:
 *
 *     01Aug2018 Clearing Cheque      4987    36.07        99,914.15
 *     01Aug2018 Incoming Interac e-Transfer      1454 101,583.92
 *
 * Read the second line. Is 1454 a cheque number or a debit? On the page it is
 * obvious, because it sits under a column heading. In the text it is
 * unknowable, and no amount of clever parsing downstream can recover it -
 * the information was destroyed the moment the words were flattened into a
 * string. That is the defect this file exists to fix.
 *
 * The engine knows where every word sat. So:
 *
 *   1. RECURSIVE XY-CUT splits the page into regions along the wide white
 *      channels a human eye reads as "these are two separate blocks" - the
 *      vendor block on the left, the bill-to block on the right.
 *   2. LINE GROUPING joins words into lines by how much their boxes overlap
 *      vertically, never by rounding a coordinate.
 *   3. COLUMN DETECTION finds the invisible vertical rules of a table by
 *      noticing that word edges line up across many rows, and gives every
 *      word a column index.
 *
 * Coordinates are whatever unit the caller passes in (pixels or points), with
 * the origin at the TOP-LEFT and y increasing downwards. Every threshold is
 * expressed as a multiple of the page's own median text height, so the same
 * code works on a 150 dpi fax and a 600 dpi archival scan without tuning.
 */

/* ------------------------------------------------------------------ statistics */

function median(values) {
  if (values.length === 0) return 0;
  const sorted = Array.prototype.slice.call(values).sort(function (a, b) { return a - b; });
  const mid = sorted.length >> 1;
  return sorted.length % 2 === 1 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2;
}

/*
 * The page's own scale, measured from its words. Everything downstream is
 * phrased in these units so that no threshold is ever a raw pixel count.
 */
function pageMetrics(words) {
  const heights = [];
  const charWidths = [];
  for (let i = 0; i < words.length; i += 1) {
    const wd = words[i];
    if (wd.h > 0) heights.push(wd.h);
    const len = wd.text ? wd.text.length : 0;
    if (len > 0 && wd.w > 0) charWidths.push(wd.w / len);
  }
  const textHeight = median(heights) || 10;
  const charWidth = median(charWidths) || textHeight * 0.5;
  return { textHeight, charWidth, count: words.length };
}

function boundsOf(words) {
  if (words.length === 0) return { x0: 0, y0: 0, x1: 0, y1: 0 };
  let x0 = Infinity;
  let y0 = Infinity;
  let x1 = -Infinity;
  let y1 = -Infinity;
  for (let i = 0; i < words.length; i += 1) {
    const wd = words[i];
    if (wd.x < x0) x0 = wd.x;
    if (wd.y < y0) y0 = wd.y;
    if (wd.x + wd.w > x1) x1 = wd.x + wd.w;
    if (wd.y + wd.h > y1) y1 = wd.y + wd.h;
  }
  return { x0, y0, x1, y1 };
}

/* --------------------------------------------------------------- projections */

/*
 * A projection profile is the classic tool here: walk a 1-pixel-wide slit
 * across the page and count how much ink it crosses. Word boxes stand in for
 * ink, which is both faster and less noisy than counting actual pixels - a
 * speck of scanner dust adds a pixel to a real profile but no word to this one.
 *
 * `bucket` keeps the arrays small: a 2550-pixel-wide page profiled at
 * one-point buckets is ~600 entries rather than 2550, and gaps are measured
 * in the same units either way.
 */
function projection(words, from, to, bucket, axis) {
  const size = Math.max(1, Math.ceil((to - from) / bucket));
  const prof = new Float64Array(size);
  for (let i = 0; i < words.length; i += 1) {
    const wd = words[i];
    const start = axis === 'x' ? wd.x : wd.y;
    const end = axis === 'x' ? wd.x + wd.w : wd.y + wd.h;
    // Weight by the word's extent on the OTHER axis, so a long line of body
    // text counts for more than a single stray page number.
    const weight = axis === 'x' ? wd.h : wd.w;
    let b0 = Math.floor((start - from) / bucket);
    let b1 = Math.ceil((end - from) / bucket);
    if (b0 < 0) b0 = 0;
    if (b1 > size) b1 = size;
    for (let b = b0; b < b1; b += 1) prof[b] += weight;
  }
  return prof;
}

/*
 * Every run of empty buckets in a profile, returned widest-first. These are
 * the candidate cut lines.
 */
function findGaps(prof, from, bucket) {
  const gaps = [];
  let runStart = -1;
  for (let i = 0; i < prof.length; i += 1) {
    if (prof[i] === 0) {
      if (runStart < 0) runStart = i;
    } else if (runStart >= 0) {
      gaps.push({ start: from + runStart * bucket, end: from + i * bucket, size: (i - runStart) * bucket });
      runStart = -1;
    }
  }
  if (runStart >= 0) {
    gaps.push({
      start: from + runStart * bucket,
      end: from + prof.length * bucket,
      size: (prof.length - runStart) * bucket,
    });
  }
  return gaps;
}

/*
 * Gaps that touch the outer edge of the region are margins, not separators -
 * cutting at one produces an empty child and an infinite recursion.
 */
function interiorGaps(gaps, from, to, tolerance) {
  return gaps.filter(function (g) {
    return g.start > from + tolerance && g.end < to - tolerance;
  });
}

/* ---------------------------------------------------------------- XY-cut */

/*
 * The guard that stops XY-cut from destroying a table.
 *
 * A wide white channel means one of two very different things. Between the
 * vendor block and the bill-to block it is a real boundary. Between the
 * Description column and the Number column of a statement it is just a sparse
 * table, and cutting there severs every row: the date and the amount end up in
 * different regions with nothing left to say they belonged together.
 *
 * The two cases look identical in a projection profile but not in the rows.
 * If the same text lines continue across the channel - words to the left AND
 * to the right at the same height, row after row - it is a table gutter and
 * must be left alone. Independent blocks do not share rows that consistently.
 */
function rowSpanEvidence(words, mid) {
  const lines = groupLines(words);
  let spanning = 0;
  let occupied = 0;
  let run = 0;
  let longestRun = 0;

  for (let i = 0; i < lines.length; i += 1) {
    let left = false;
    let right = false;
    const ws = lines[i].words;
    for (let j = 0; j < ws.length; j += 1) {
      if (ws[j].x + ws[j].w / 2 < mid) left = true;
      else right = true;
    }
    if (left || right) occupied += 1;
    if (left && right) {
      spanning += 1;
      run += 1;
      if (run > longestRun) longestRun = run;
    } else {
      run = 0;
    }
  }

  return {
    fraction: occupied === 0 ? 0 : spanning / occupied,
    longestRun,
    lines: occupied,
  };
}

/*
 * Kept as its own function because the fraction alone is a useful measure and
 * the tests assert on it directly.
 */
function rowSpanFraction(words, mid) {
  return rowSpanEvidence(words, mid).fraction;
}

/*
 * Recursive XY-cut (Nagy & Seth). At each step, look for the widest white
 * channel running all the way across the region - first horizontally, then
 * vertically - and if it is wider than the threshold for its direction, cut
 * there and recurse on both halves. When neither direction has a wide enough
 * channel, the region is a leaf: a single block of related text.
 *
 * The two thresholds are deliberately different, and that asymmetry is the
 * whole trick:
 *
 *   - a VERTICAL channel (splitting left from right) must be wide, because
 *     the gaps between columns of a table are narrow and must NOT be cut -
 *     cutting them would sever every row of the table;
 *   - a HORIZONTAL channel (splitting top from bottom) can be narrower,
 *     because that is just a paragraph break and cutting it loses nothing.
 *
 * Table columns are recovered later, by alignment, in findColumns() - which is
 * the right tool for them precisely because it preserves the rows.
 */
function xyCut(words, options) {
  const opts = options || {};
  const metrics = pageMetrics(words);
  const bucket = Math.max(1, metrics.textHeight / 4);

  // A page-level gutter is wide: three lines' worth of white. A table's
  // inter-column gap is one or two character widths, so it never qualifies.
  const minVerticalGap = opts.minColumnGap !== undefined
    ? opts.minColumnGap
    : Math.max(metrics.textHeight * 2.5, metrics.charWidth * 6);

  // A block separator is a blank line and a bit.
  const minHorizontalGap = opts.minRowGap !== undefined
    ? opts.minRowGap
    : metrics.textHeight * 1.4;

  // How much shared row structure it takes to veto a vertical cut. At 0.5,
  // half the rows reaching across the channel is enough to call it a table.
  const maxRowSpan = opts.maxRowSpan === undefined ? 0.5 : opts.maxRowSpan;

  // A run of this many consecutive rows all reaching across the channel is a
  // table on its own evidence, whatever the rest of the page is doing.
  const minTableRows = opts.minTableRows === undefined ? 3 : opts.minTableRows;
  const maxDepth = opts.maxDepth === undefined ? 12 : opts.maxDepth;
  const minWords = opts.minWords === undefined ? 2 : opts.minWords;
  const regions = [];

  function cut(subset, depth) {
    if (subset.length === 0) return;
    if (subset.length < minWords || depth >= maxDepth) {
      regions.push(subset);
      return;
    }

    const b = boundsOf(subset);
    const width = b.x1 - b.x0;
    const height = b.y1 - b.y0;

    let bestHorizontal = null;
    if (height > minHorizontalGap * 2) {
      const prof = projection(subset, b.y0, b.y1, bucket, 'y');
      const gaps = interiorGaps(findGaps(prof, b.y0, bucket), b.y0, b.y1, bucket);
      for (let i = 0; i < gaps.length; i += 1) {
        if (gaps[i].size >= minHorizontalGap
          && (bestHorizontal === null || gaps[i].size > bestHorizontal.size)) {
          bestHorizontal = gaps[i];
        }
      }
    }

    let bestVertical = null;
    if (width > minVerticalGap * 2) {
      const prof = projection(subset, b.x0, b.x1, bucket, 'x');
      const gaps = interiorGaps(findGaps(prof, b.x0, bucket), b.x0, b.x1, bucket);
      for (let i = 0; i < gaps.length; i += 1) {
        if (gaps[i].size >= minVerticalGap
          && (bestVertical === null || gaps[i].size > bestVertical.size)) {
          bestVertical = gaps[i];
        }
      }
    }

    /*
     * When both directions offer a cut, prefer the vertical one. Splitting
     * side-by-side blocks apart first is what keeps a two-up address header
     * from being read as one wide paragraph; if the horizontal cut ran first,
     * each band would still straddle both blocks.
     */
    if (bestVertical !== null) {
      const mid = (bestVertical.start + bestVertical.end) / 2;
      /*
       * Refuse the cut if the rows carry on across it: that is a table, and
       * severing it costs more than any layout gain gives back.
       *
       * TWO tests, because the fraction on its own is not enough. On a real
       * invoice the line-item table sits above a totals block that lives
       * entirely on the right-hand side. Those totals rows do not span the
       * channel, so they drag the fraction below the threshold - and the cut
       * went through, slicing the item table into three pieces with the
       * descriptions in one region and the amounts in another.
       *
       * The run length fixes it. A band of consecutive rows that ALL reach
       * across the channel is a table wherever else on the page is empty, and
       * what happens above or below it cannot argue otherwise.
       */
      const evidence = rowSpanEvidence(subset, mid);
      if (evidence.fraction >= maxRowSpan || evidence.longestRun >= minTableRows) {
        bestVertical = null;
      }
    }

    if (bestVertical !== null) {
      const mid = (bestVertical.start + bestVertical.end) / 2;
      const left = [];
      const right = [];
      for (let i = 0; i < subset.length; i += 1) {
        // Assign by centre, so a word that overhangs the channel still lands
        // on the side it mostly occupies.
        if (subset[i].x + subset[i].w / 2 < mid) left.push(subset[i]);
        else right.push(subset[i]);
      }
      if (left.length > 0 && right.length > 0) {
        cut(left, depth + 1);
        cut(right, depth + 1);
        return;
      }
    }

    if (bestHorizontal !== null) {
      const mid = (bestHorizontal.start + bestHorizontal.end) / 2;
      const top = [];
      const bottom = [];
      for (let i = 0; i < subset.length; i += 1) {
        if (subset[i].y + subset[i].h / 2 < mid) top.push(subset[i]);
        else bottom.push(subset[i]);
      }
      if (top.length > 0 && bottom.length > 0) {
        cut(top, depth + 1);
        cut(bottom, depth + 1);
        return;
      }
    }

    regions.push(subset);
  }

  cut(words.slice(), 0);
  return regions;
}

/* ------------------------------------------------------------ line grouping */

/*
 * Words belong to the same line when their vertical spans genuinely overlap.
 *
 * The tempting shortcut - round y to the nearest few pixels and group equal
 * values - fails on every real page: a line with a capital letter or a comma
 * has words whose tops differ by several pixels, and any rounding boundary
 * eventually falls between two words of one line. Overlap has no such seam.
 *
 * The threshold is a fraction of the SHORTER word's height, so a full-height
 * word and a lone comma still pair up.
 */
function groupLines(words, options) {
  const opts = options || {};
  const minOverlap = opts.minOverlap === undefined ? 0.4 : opts.minOverlap;
  if (words.length === 0) return [];

  const sorted = words.slice().sort(function (a, b) {
    if (a.y !== b.y) return a.y - b.y;
    return a.x - b.x;
  });

  const lines = [];
  for (let i = 0; i < sorted.length; i += 1) {
    const wd = sorted[i];
    const top = wd.y;
    const bottom = wd.y + wd.h;
    let placed = false;

    // Search from the most recent line backwards: words arrive in y order, so
    // the match is almost always the last line or the one before it.
    for (let j = lines.length - 1; j >= 0 && !placed; j -= 1) {
      const line = lines[j];
      const overlap = Math.min(bottom, line.y1) - Math.max(top, line.y0);
      if (overlap <= 0) {
        // Once a line ends above this word's top by more than a line height,
        // no earlier line can match either.
        if (line.y1 < top - wd.h) break;
        continue;
      }
      const shorter = Math.min(bottom - top, line.y1 - line.y0);
      if (shorter > 0 && overlap / shorter >= minOverlap) {
        line.words.push(wd);
        if (top < line.y0) line.y0 = top;
        if (bottom > line.y1) line.y1 = bottom;
        placed = true;
      }
    }

    if (!placed) lines.push({ y0: top, y1: bottom, words: [wd] });
  }

  for (let i = 0; i < lines.length; i += 1) {
    lines[i].words.sort(function (a, b) { return a.x - b.x; });
    const b = boundsOf(lines[i].words);
    lines[i].x0 = b.x0;
    lines[i].x1 = b.x1;
    lines[i].y0 = b.y0;
    lines[i].y1 = b.y1;
  }
  lines.sort(function (a, b) {
    if (Math.abs(a.y0 - b.y0) > 1) return a.y0 - b.y0;
    return a.x0 - b.x0;
  });
  return lines;
}

/* -------------------------------------------------------- column detection */

/*
 * One-dimensional clustering: group values that sit within `tolerance` of
 * each other, and report each cluster's centre and how many DISTINCT lines
 * contributed to it. The distinct-line count is what separates a real column
 * from a coincidence - three words at the same x on the same line prove
 * nothing, three words at the same x on three different lines prove a column.
 */
function clusterEdges(entries, tolerance) {
  if (entries.length === 0) return [];
  const sorted = entries.slice().sort(function (a, b) { return a.value - b.value; });
  const clusters = [];
  let current = { values: [sorted[0].value], lines: {} };
  current.lines[sorted[0].line] = true;

  for (let i = 1; i < sorted.length; i += 1) {
    if (sorted[i].value - sorted[i - 1].value <= tolerance) {
      current.values.push(sorted[i].value);
      current.lines[sorted[i].line] = true;
    } else {
      clusters.push(current);
      current = { values: [sorted[i].value], lines: {} };
      current.lines[sorted[i].line] = true;
    }
  }
  clusters.push(current);

  return clusters.map(function (c) {
    return {
      centre: median(c.values),
      min: c.values[0],
      max: c.values[c.values.length - 1],
      count: c.values.length,
      lineCount: Object.keys(c.lines).length,
    };
  });
}

/*
 * Find a table's columns without any ruling lines to help.
 *
 * Text columns are left-aligned, so their words share a left edge. Money
 * columns are right-aligned, so their words share a RIGHT edge and their left
 * edges are all over the place - which is exactly why looking only at left
 * edges finds the description column and misses every amount. Both edges are
 * clustered, and each surviving cluster becomes a column anchor carrying the
 * alignment it was found by.
 */
function findColumns(lines, metrics, options) {
  const opts = options || {};
  const tolerance = opts.tolerance === undefined ? metrics.charWidth * 1.2 : opts.tolerance;
  const minLines = opts.minLines === undefined
    ? Math.max(3, Math.ceil(lines.length * 0.25))
    : opts.minLines;

  const lefts = [];
  const rights = [];
  for (let i = 0; i < lines.length; i += 1) {
    const ws = lines[i].words;
    for (let j = 0; j < ws.length; j += 1) {
      lefts.push({ value: ws[j].x, line: i });
      rights.push({ value: ws[j].x + ws[j].w, line: i });
    }
  }

  /*
   * A shared left edge is not enough on its own. In "01Aug2018 Clearing
   * Cheque", repeated on forty rows, the word "Cheque" starts at the same x
   * every time - but it is the second half of a sentence, not a column. What
   * separates the two is the white in front: a column starts after a wide
   * gap, a word starts after a single space.
   *
   * So an anchor must also be preceded, on most of the lines that support it,
   * by more white than an ordinary word space.
   */
  function precededByGap(position) {
    let wide = 0;
    let seen = 0;
    for (let i = 0; i < lines.length; i += 1) {
      const ws = lines[i].words;
      for (let j = 0; j < ws.length; j += 1) {
        if (Math.abs(ws[j].x - position) > tolerance) continue;
        seen += 1;
        // The left margin itself always counts: nothing precedes it.
        if (j === 0) { wide += 1; break; }
        const gap = ws[j].x - (ws[j - 1].x + ws[j - 1].w);
        if (gap >= metrics.charWidth * 2) wide += 1;
        break;
      }
    }
    return seen > 0 && wide / seen >= 0.6;
  }

  const leftClusters = clusterEdges(lefts, tolerance)
    .filter(function (c) { return c.lineCount >= minLines && precededByGap(c.centre); })
    .map(function (c) { return { position: c.centre, align: 'left', support: c.lineCount }; });

  /* The mirror test for right-aligned columns: wide white AFTER the edge. */
  function followedByGap(position) {
    let wide = 0;
    let seen = 0;
    for (let i = 0; i < lines.length; i += 1) {
      const ws = lines[i].words;
      for (let j = 0; j < ws.length; j += 1) {
        if (Math.abs((ws[j].x + ws[j].w) - position) > tolerance) continue;
        seen += 1;
        if (j === ws.length - 1) { wide += 1; break; }
        const gap = ws[j + 1].x - (ws[j].x + ws[j].w);
        if (gap >= metrics.charWidth * 2) wide += 1;
        break;
      }
    }
    return seen > 0 && wide / seen >= 0.6;
  }

  const rightClusters = clusterEdges(rights, tolerance)
    .filter(function (c) { return c.lineCount >= minLines && followedByGap(c.centre); })
    .map(function (c) { return { position: c.centre, align: 'right', support: c.lineCount }; });

  const anchors = leftClusters.concat(rightClusters)
    .sort(function (a, b) { return a.position - b.position; });

  /*
   * A right-aligned money column and the left edge of the next column can be
   * only a space apart and would otherwise be reported as two columns. Merge
   * anchors that are closer together than a couple of characters, keeping
   * whichever had more support.
   */
  const merged = [];
  for (let i = 0; i < anchors.length; i += 1) {
    const last = merged[merged.length - 1];
    if (last && anchors[i].position - last.position < tolerance) {
      if (anchors[i].support > last.support) merged[merged.length - 1] = anchors[i];
    } else {
      merged.push(anchors[i]);
    }
  }
  return merged;
}

/*
 * Where the column edges actually are.
 *
 * The first attempt at this took the midpoint between two neighbouring
 * alignment anchors. On a real bank statement that put a boundary straight
 * through the middle of the Debits column - because the column had produced
 * an anchor for its left edge AND one for its right, and the midpoint of the
 * two sits inside it. "800.04" and "823.34", printed one directly above the
 * other, landed in different columns.
 *
 * The edges are not between the anchors. They are the WHITE CHANNELS: the
 * vertical strips the region's words never occupy. Reading the boundaries
 * straight off the projection profile makes cutting through a word almost
 * impossible by construction, and it finds edges that no anchor marks - the
 * gap between Date and Description, where both columns are left-aligned and
 * there is only one anchor between them.
 *
 * A channel does not have to be perfectly empty. One overlong description
 * that spills into the next column would erase a strictly-empty gap and
 * merge two columns for the whole table, so a channel is a VALLEY - coverage
 * far below the region's typical density - and every candidate is then
 * checked to make sure it does not saw through a run of words.
 */
function findGutters(lines, metrics, options) {
  const opts = options || {};
  const words = [];
  for (let i = 0; i < lines.length; i += 1) {
    for (let j = 0; j < lines[i].words.length; j += 1) words.push(lines[i].words[j]);
  }
  if (words.length === 0) return [];

  const b = boundsOf(words);
  const bucket = Math.max(0.5, metrics.charWidth / 4);
  const size = Math.max(1, Math.ceil((b.x1 - b.x0) / bucket));

  /*
   * Coverage here is counted in ROWS, not in ink.
   *
   * Weighting by ink made the threshold impossible to reason about: whether a
   * strip counted as empty depended on how tall the type was and how many
   * lines happened to be in the region. Counting rows says exactly what is
   * meant - "a gutter is a vertical strip that hardly any row has text in" -
   * and one long description overhanging its column then costs one row out of
   * forty instead of erasing the gutter altogether.
   */
  const coverage = new Uint32Array(size);
  const touched = new Uint8Array(size);
  for (let i = 0; i < lines.length; i += 1) {
    touched.fill(0);
    const ws = lines[i].words;
    for (let j = 0; j < ws.length; j += 1) {
      let b0 = Math.floor((ws[j].x - b.x0) / bucket);
      let b1 = Math.ceil((ws[j].x + ws[j].w - b.x0) / bucket);
      if (b0 < 0) b0 = 0;
      if (b1 > size) b1 = size;
      for (let k = b0; k < b1; k += 1) touched[k] = 1;
    }
    for (let k = 0; k < size; k += 1) coverage[k] += touched[k];
  }

  /*
   * At most this many rows may intrude into a strip for it to still count as
   * a gutter. It has to tolerate more than nothing: a heading set across the
   * top of a table ("FIRST CHEQUING") legitimately runs over every column
   * below it, and demanding a perfectly clear channel would lose every column
   * edge under a heading.
   */
  const maxRows = opts.maxGutterRows === undefined
    ? Math.max(1, Math.floor(lines.length * 0.15))
    : opts.maxGutterRows;

  // A column gutter is wider than the space between two words.
  const minWidth = opts.minGutter === undefined ? metrics.charWidth * 1.2 : opts.minGutter;

  const valleys = [];
  let runStart = -1;
  for (let i = 0; i <= size; i += 1) {
    const low = i < size && coverage[i] <= maxRows;
    if (low) {
      if (runStart < 0) runStart = i;
    } else if (runStart >= 0) {
      const x0 = b.x0 + runStart * bucket;
      const x1 = b.x0 + i * bucket;
      // Ignore the margins: a channel at the edge of the region separates nothing.
      if (x1 - x0 >= minWidth && x0 > b.x0 + bucket && x1 < b.x1 - bucket) {
        valleys.push((x0 + x1) / 2);
      }
      runStart = -1;
    }
  }

  /*
   * Final safety net. A boundary that cuts through words is not a boundary,
   * whatever the profile said - this is the test that caught the Debits
   * column being split in two.
   */
  const slack = opts.straddleSlack === undefined ? 1 : opts.straddleSlack;
  /*
   * Keep this in step with maxRows. Each row allowed to intrude into a gutter
   * contributes about one straddling word, so a stricter straddle budget than
   * the gutter budget would silently throw away every edge the scan above just
   * accepted - the two tests would be arguing with each other.
   */
  const allowed = opts.maxStraddle === undefined
    ? Math.max(maxRows, Math.floor(words.length * 0.02))
    : opts.maxStraddle;

  return valleys.filter(function (x) {
    let straddling = 0;
    for (let i = 0; i < words.length; i += 1) {
      if (words[i].x + slack < x && x < words[i].x + words[i].w - slack) {
        straddling += 1;
        if (straddling > allowed) return false;
      }
    }
    return true;
  });
}

/*
 * Give every word a column index, then drop the columns nothing landed in and
 * renumber, so the index is always a dense 0..n-1 that a spreadsheet can use
 * as a cell reference without further thought.
 */
function assignColumns(lines, boundaries) {
  const words = [];
  for (let i = 0; i < lines.length; i += 1) {
    for (let j = 0; j < lines[i].words.length; j += 1) words.push(lines[i].words[j]);
  }

  function assign(bounds) {
    for (let i = 0; i < words.length; i += 1) {
      const centre = words[i].x + words[i].w / 2;
      let col = 0;
      while (col < bounds.length && centre > bounds[col]) col += 1;
      words[i].column = col;
    }
  }

  if (!boundaries || boundaries.length === 0 || words.length === 0) {
    assign([]);
    return { boundaries: [], columnCount: 1 };
  }

  const sorted = boundaries.slice().sort(function (a, c) { return a - c; });
  assign(sorted);

  const used = {};
  for (let i = 0; i < words.length; i += 1) used[words[i].column] = true;

  const kept = [];
  for (let c = 0; c <= sorted.length; c += 1) if (used[c]) kept.push(c);
  if (kept.length === 0) {
    assign([]);
    return { boundaries: [], columnCount: 1 };
  }

  const remap = {};
  for (let i = 0; i < kept.length; i += 1) remap[kept[i]] = i;
  for (let i = 0; i < words.length; i += 1) words[i].column = remap[words[i].column];

  const finalBoundaries = [];
  for (let i = 0; i < sorted.length; i += 1) {
    if (used[i] && used[i + 1]) finalBoundaries.push(sorted[i]);
  }
  return { boundaries: finalBoundaries, columnCount: kept.length };
}

/*
 * Is this region actually a table? A block of prose produces one or two left
 * edge clusters (the margin and the indent); a table produces several columns
 * each supported by most of its rows. Requiring three or more anchors and a
 * decent number of rows keeps an address block from being reported as a table.
 */
function looksTabular(lines, columnCount) {
  if (lines.length < 3 || columnCount < 3) return false;

  /*
   * Counting words per line is the wrong test - it demands that every row
   * fill every column, and real tables are full of blanks (a row with a debit
   * has no credit). What marks a table is that most rows put words into
   * SEVERAL DIFFERENT columns. A paragraph, however many words it holds,
   * occupies one run of adjacent columns per line.
   */
  let structured = 0;
  for (let i = 0; i < lines.length; i += 1) {
    const seen = {};
    const ws = lines[i].words;
    for (let j = 0; j < ws.length; j += 1) {
      if (ws[j].column !== undefined) seen[ws[j].column] = true;
    }
    if (Object.keys(seen).length >= 3) structured += 1;
  }
  return structured >= Math.max(2, Math.floor(lines.length * 0.5));
}

/* ------------------------------------------------------------ reading order */

/*
 * Sort regions the way a person reads them: down each column, then on to the
 * next column. Sorting purely by y would interleave the two halves of a
 * side-by-side header, one line at a time, which is precisely the failure
 * this whole file exists to prevent.
 */
function orderRegions(regions, pageWidth) {
  const withBounds = regions.map(function (r, i) {
    return { words: r, bounds: boundsOf(r), index: i };
  });

  // Group regions into vertical bands that overlap horizontally - those are
  // the page's columns.
  const bands = [];
  const byX = withBounds.slice().sort(function (a, b) { return a.bounds.x0 - b.bounds.x0; });
  for (let i = 0; i < byX.length; i += 1) {
    const r = byX[i];
    let placed = false;
    for (let j = 0; j < bands.length && !placed; j += 1) {
      const band = bands[j];
      const overlap = Math.min(r.bounds.x1, band.x1) - Math.max(r.bounds.x0, band.x0);
      const narrower = Math.min(r.bounds.x1 - r.bounds.x0, band.x1 - band.x0);
      if (narrower > 0 && overlap / narrower > 0.5) {
        band.regions.push(r);
        band.x0 = Math.min(band.x0, r.bounds.x0);
        band.x1 = Math.max(band.x1, r.bounds.x1);
        placed = true;
      }
    }
    if (!placed) bands.push({ x0: r.bounds.x0, x1: r.bounds.x1, regions: [r] });
  }

  /*
   * A region spanning most of the page width - a title, a footer, a table that
   * runs edge to edge - is not part of any column. Left in a band it would
   * force everything else into one band and defeat the ordering, so full-width
   * regions are ordered by y alone and the narrow ones by band.
   */
  const full = [];
  const banded = [];
  for (let i = 0; i < bands.length; i += 1) {
    const band = bands[i];
    if (pageWidth > 0 && (band.x1 - band.x0) / pageWidth > 0.75) {
      for (let j = 0; j < band.regions.length; j += 1) full.push(band.regions[j]);
    } else {
      banded.push(band);
    }
  }

  banded.sort(function (a, b) { return a.x0 - b.x0; });
  const ordered = [];
  for (let i = 0; i < banded.length; i += 1) {
    banded[i].regions.sort(function (a, b) { return a.bounds.y0 - b.bounds.y0; });
    for (let j = 0; j < banded[i].regions.length; j += 1) ordered.push(banded[i].regions[j]);
  }

  // Merge the full-width regions back in by vertical position.
  const all = ordered.concat(full).sort(function (a, b) {
    const aFull = full.indexOf(a) >= 0;
    const bFull = full.indexOf(b) >= 0;
    if (aFull || bFull) return a.bounds.y0 - b.bounds.y0;
    return 0;
  });

  return (full.length > 0 ? all : ordered).map(function (r) { return r.words; });
}

/* ---------------------------------------------------------------- the whole job */

/*
 * words -> a structured page.
 *
 * Returns { metrics, regions: [ { bounds, kind, columns, lines: [...] } ] }
 * where every line carries its words in reading order and every word carries
 * the column it belongs to.
 */
function segmentPage(words, options) {
  const opts = options || {};
  const clean = words.filter(function (w) {
    return w && w.text !== undefined && String(w.text).trim() !== '' && w.w > 0 && w.h > 0;
  });

  const metrics = pageMetrics(clean);
  if (clean.length === 0) {
    return { metrics, regions: [], lines: [], pageBounds: { x0: 0, y0: 0, x1: 0, y1: 0 } };
  }

  const pageBounds = boundsOf(clean);
  const pageWidth = opts.pageWidth || (pageBounds.x1 - pageBounds.x0);

  const rawRegions = xyCut(clean, opts);
  const ordered = orderRegions(rawRegions, pageWidth);

  const regions = ordered.map(function (regionWords, index) {
    const lines = groupLines(regionWords, opts);

    /*
     * Assign columns first, then judge whether this is a table. looksTabular
     * needs to see where the words landed, and a text region loses nothing by
     * carrying a column index too - a two-up address header is not a table,
     * but knowing which side each word sat on is exactly what stops the two
     * addresses being read as one.
     */
    const gutters = findGutters(lines, metrics, opts);
    const assignment = assignColumns(lines, gutters);
    const anchors = findColumns(lines, metrics, opts);
    const tabular = looksTabular(lines, assignment.columnCount);

    return {
      index,
      bounds: boundsOf(regionWords),
      kind: tabular ? 'table' : 'text',
      columns: anchors,
      boundaries: assignment.boundaries,
      columnCount: assignment.columnCount,
      lines,
    };
  });

  // A flat, reading-ordered line list, for callers that just want the text.
  const flatLines = [];
  for (let i = 0; i < regions.length; i += 1) {
    for (let j = 0; j < regions[i].lines.length; j += 1) {
      flatLines.push(regions[i].lines[j]);
    }
  }

  return { metrics, regions, lines: flatLines, pageBounds };
}

/* --------------------------------------------------------------- text views */

function lineText(line) {
  return line.words.map(function (w) { return w.text; }).join(' ');
}

/*
 * Render the page back onto a character grid, the way it physically sat. This
 * is the debugging view: if a number is not visible here, no downstream rule
 * can possibly find it, and the problem is in scanning or preprocessing rather
 * than in parsing.
 */
function renderLayout(page, options) {
  const opts = options || {};
  const columns = opts.columns || 120;
  const b = page.pageBounds;
  const width = b.x1 - b.x0;
  const height = b.y1 - b.y0;
  if (width <= 0 || height <= 0) return '';

  const scaleX = columns / width;
  // One text row per line of type, so the output stays roughly page-shaped.
  const rowHeight = page.metrics.textHeight > 0 ? page.metrics.textHeight : 10;
  const rows = Math.max(1, Math.ceil(height / rowHeight) + 1);

  const grid = [];
  for (let i = 0; i < rows; i += 1) grid.push(new Array(columns).fill(' '));

  /*
   * Squeezing a wide page into 120 characters means two words can want the
   * same column. Writing only into blank cells would run them together -
   * "Clearing" and "Cheque" arriving as "ClearingCheque" - which is exactly
   * the ambiguity this view exists to expose. So each word is nudged right
   * until it has a space of its own.
   */
  const all = page.lines;
  for (let i = 0; i < all.length; i += 1) {
    const line = all[i];
    const row = Math.min(rows - 1, Math.max(0, Math.round((line.y0 - b.y0) / rowHeight)));
    let cursor = 0;
    for (let j = 0; j < line.words.length; j += 1) {
      const wd = line.words[j];
      const text = String(wd.text);
      let col = Math.round((wd.x - b.x0) * scaleX);
      if (col < cursor) col = cursor;
      if (col + text.length > columns) col = Math.max(0, columns - text.length);
      for (let k = 0; k < text.length; k += 1) {
        const c = col + k;
        if (c >= 0 && c < columns) grid[row][c] = text.charAt(k);
      }
      cursor = col + text.length + 1;
    }
  }
  return grid.map(function (r) { return r.join('').replace(/\s+$/, ''); }).join('\n');
}

/*
 * The table view: rows of cells, cell n of every row being column n. This is
 * the shape that goes into a spreadsheet, and it is only obtainable because
 * the words kept their positions.
 */
function tableRows(region) {
  return region.lines.map(function (line) {
    const cells = [];
    for (let i = 0; i < region.columnCount; i += 1) cells.push([]);
    for (let j = 0; j < line.words.length; j += 1) {
      const wd = line.words[j];
      const col = wd.column === undefined ? 0 : Math.min(region.columnCount - 1, wd.column);
      cells[col].push(wd.text);
    }
    return cells.map(function (c) { return c.join(' '); });
  });
}

module.exports = {
  median,
  rowSpanFraction,
  rowSpanEvidence,
  pageMetrics,
  boundsOf,
  projection,
  findGaps,
  xyCut,
  groupLines,
  clusterEdges,
  findColumns,
  findGutters,
  assignColumns,
  looksTabular,
  orderRegions,
  segmentPage,
  lineText,
  renderLayout,
  tableRows,
};
