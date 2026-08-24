'use strict';

/*
 * preprocess.js - what happens to a page image before Tesseract ever sees it.
 *
 * This file is where most of the accuracy lives. The Tesseract documentation
 * puts it plainly: "the same image will get much better results if you upscale
 * it before calling recognize". A 150 dpi fax of an invoice recognises badly
 * not because the engine is weak but because the letters are eight pixels tall.
 *
 * Everything here works on a plain object:
 *
 *     { width, height, channels, data }   data = Uint8Array, 8 bits per sample
 *
 * and every function returns a NEW image rather than mutating its input, so a
 * caller can keep the original around to fall back to.
 */

/* ------------------------------------------------------------------ greyscale */

/*
 * ITU-R BT.601 luma. Not a plain average: the eye - and the scanner - carry
 * far more detail in green than in blue, and a plain average visibly muddies
 * text printed in colour on a coloured form.
 */
function toGrayscale(image) {
  if (image.channels === 1) {
    return {
      width: image.width, height: image.height, channels: 1,
      data: Uint8Array.from(image.data.subarray(0, image.width * image.height)),
    };
  }
  const n = image.width * image.height;
  const out = new Uint8Array(n);
  const src = image.data;
  const ch = image.channels;
  for (let i = 0; i < n; i += 1) {
    const r = src[i * ch];
    const g = src[i * ch + 1];
    const b = src[i * ch + 2];
    out[i] = (0.299 * r + 0.587 * g + 0.114 * b + 0.5) | 0;
  }
  return { width: image.width, height: image.height, channels: 1, data: out };
}

/* ------------------------------------------------------------------ resampling */

function clampIndex(v, max) {
  if (v < 0) return 0;
  if (v > max) return max;
  return v;
}

/*
 * Catmull-Rom cubic. Chosen over bilinear because upscaling text with bilinear
 * gives soft, grey-edged glyphs that threshold into blobs, while Catmull-Rom
 * keeps the edge sharp. Chosen over Lanczos because Lanczos rings, and ringing
 * next to a letter stroke reads as a second, fainter stroke.
 */
function cubicWeight(t) {
  const x = Math.abs(t);
  if (x <= 1) return 1.5 * x * x * x - 2.5 * x * x + 1;
  if (x < 2) return -0.5 * x * x * x + 2.5 * x * x - 4 * x + 2;
  return 0;
}

function resample(image, targetWidth, targetHeight) {
  const tw = Math.max(1, Math.round(targetWidth));
  const th = Math.max(1, Math.round(targetHeight));
  if (tw === image.width && th === image.height) {
    return {
      width: tw, height: th, channels: image.channels, data: Uint8Array.from(image.data),
    };
  }

  const { width: sw, height: sh, channels: ch, data: src } = image;
  const out = new Uint8Array(tw * th * ch);
  const xRatio = sw / tw;
  const yRatio = sh / th;

  /*
   * Downscaling with a point-sampled cubic kernel drops pixels and turns a
   * thin stroke into a dotted line. When shrinking, box-average instead: it is
   * the correct thing for area reduction and it is what keeps 600 dpi scans
   * legible when brought down to 300.
   */
  const shrinking = xRatio > 1.05 || yRatio > 1.05;

  if (shrinking) {
    for (let y = 0; y < th; y += 1) {
      const sy0 = y * yRatio;
      const sy1 = Math.min(sh, (y + 1) * yRatio);
      const y0 = Math.floor(sy0);
      const y1 = Math.max(y0 + 1, Math.ceil(sy1));
      for (let x = 0; x < tw; x += 1) {
        const sx0 = x * xRatio;
        const sx1 = Math.min(sw, (x + 1) * xRatio);
        const x0 = Math.floor(sx0);
        const x1 = Math.max(x0 + 1, Math.ceil(sx1));
        for (let c = 0; c < ch; c += 1) {
          let sum = 0;
          let count = 0;
          for (let yy = y0; yy < y1 && yy < sh; yy += 1) {
            for (let xx = x0; xx < x1 && xx < sw; xx += 1) {
              sum += src[(yy * sw + xx) * ch + c];
              count += 1;
            }
          }
          out[(y * tw + x) * ch + c] = count === 0 ? 0 : Math.round(sum / count);
        }
      }
    }
    return { width: tw, height: th, channels: ch, data: out };
  }

  for (let y = 0; y < th; y += 1) {
    const sy = (y + 0.5) * yRatio - 0.5;
    const iy = Math.floor(sy);
    const fy = sy - iy;
    const wy = [cubicWeight(fy + 1), cubicWeight(fy), cubicWeight(fy - 1), cubicWeight(fy - 2)];

    for (let x = 0; x < tw; x += 1) {
      const sx = (x + 0.5) * xRatio - 0.5;
      const ix = Math.floor(sx);
      const fx = sx - ix;
      const wx = [cubicWeight(fx + 1), cubicWeight(fx), cubicWeight(fx - 1), cubicWeight(fx - 2)];

      for (let c = 0; c < ch; c += 1) {
        let acc = 0;
        let norm = 0;
        for (let m = 0; m < 4; m += 1) {
          const yy = clampIndex(iy - 1 + m, sh - 1);
          for (let n = 0; n < 4; n += 1) {
            const xx = clampIndex(ix - 1 + n, sw - 1);
            const w = wy[m] * wx[n];
            acc += src[(yy * sw + xx) * ch + c] * w;
            norm += w;
          }
        }
        const v = norm === 0 ? 0 : acc / norm;
        out[(y * tw + x) * ch + c] = v < 0 ? 0 : (v > 255 ? 255 : Math.round(v));
      }
    }
  }
  return { width: tw, height: th, channels: ch, data: out };
}

/*
 * Bring a page to a target working resolution. Tesseract's models were trained
 * on roughly 300 dpi text; below about 200 dpi accuracy falls off a cliff and
 * above about 400 dpi nothing improves but everything slows down.
 *
 * maxDimension guards against a 1200 dpi archival scan turning into a 40
 * megapixel buffer that exhausts memory before a single word is read.
 */
function scaleToDpi(image, sourceDpi, targetDpi, maxDimension) {
  const limit = maxDimension || 6000;
  const from = sourceDpi > 0 ? sourceDpi : 72;
  let scale = targetDpi / from;

  const cap = limit / Math.max(image.width, image.height);
  if (scale > cap) scale = cap;
  if (Math.abs(scale - 1) < 0.02) {
    return {
      image: {
        width: image.width, height: image.height, channels: image.channels,
        data: Uint8Array.from(image.data),
      },
      scale: 1,
      effectiveDpi: from,
    };
  }

  return {
    image: resample(image, image.width * scale, image.height * scale),
    scale,
    effectiveDpi: from * scale,
  };
}

/* ------------------------------------------------------------------ histogram */

function histogram(gray) {
  const h = new Uint32Array(256);
  for (let i = 0; i < gray.data.length; i += 1) h[gray.data[i]] += 1;
  return h;
}

/*
 * Clip the darkest and lightest `pct` of the pixels and stretch what remains
 * across the full range. A scan whose "white" is 200 and whose "black" is 60
 * becomes a scan whose white is 255 and black is 0, which is what every later
 * threshold decision assumes.
 */
function stretchContrast(gray, pct) {
  const p = pct === undefined ? 0.5 : pct;
  const h = histogram(gray);
  const total = gray.data.length;
  const cut = Math.floor((total * p) / 100);

  let lo = 0;
  let acc = 0;
  while (lo < 255 && acc + h[lo] <= cut) { acc += h[lo]; lo += 1; }

  let hi = 255;
  acc = 0;
  while (hi > 0 && acc + h[hi] <= cut) { acc += h[hi]; hi -= 1; }

  if (hi <= lo) {
    return {
      width: gray.width, height: gray.height, channels: 1, data: Uint8Array.from(gray.data),
    };
  }

  const lut = new Uint8Array(256);
  const span = hi - lo;
  for (let v = 0; v < 256; v += 1) {
    const s = Math.round(((v - lo) * 255) / span);
    lut[v] = s < 0 ? 0 : (s > 255 ? 255 : s);
  }

  const out = new Uint8Array(gray.data.length);
  for (let i = 0; i < gray.data.length; i += 1) out[i] = lut[gray.data[i]];
  return { width: gray.width, height: gray.height, channels: 1, data: out };
}

/* ----------------------------------------------------------------- thresholds */

/*
 * Otsu: choose the threshold that best splits the histogram into two groups by
 * maximising between-class variance. One number for the whole page - fast, and
 * right whenever the page is evenly lit.
 */
function otsuThreshold(gray) {
  const h = histogram(gray);
  const total = gray.data.length;

  let sum = 0;
  for (let i = 0; i < 256; i += 1) sum += i * h[i];

  let sumB = 0;
  let wB = 0;
  let bestVar = -1;
  let bestLow = 0;
  let bestHigh = 0;

  for (let t = 0; t < 256; t += 1) {
    wB += h[t];
    if (wB === 0) continue;
    const wF = total - wB;
    if (wF === 0) break;
    sumB += t * h[t];
    const mB = sumB / wB;
    const mF = (sum - sumB) / wF;
    const between = wB * wF * (mB - mF) * (mB - mF);
    if (between > bestVar) { bestVar = between; bestLow = t; bestHigh = t; }
    else if (between === bestVar) { bestHigh = t; }
  }

  /*
   * An already-binary page has ink at 0 and paper at 255 and nothing between,
   * so every threshold from 0 to 254 scores identically. Taking the first
   * winner would return 0, which is technically a valid split but leaves no
   * headroom: one stray grey pixel from a rescale then counts as paper.
   * Taking the middle of the tied plateau puts the cut where a human would.
   */
  return Math.round((bestLow + bestHigh) / 2);
}

/*
 * Integral images: the trick that makes a windowed threshold affordable. Sum
 * and sum-of-squares over any rectangle become four lookups, so a 31x31 local
 * window over a 2500x3300 page costs one pass instead of nine hundred.
 */
function buildIntegrals(gray) {
  const w = gray.width;
  const h = gray.height;
  const sum = new Float64Array((w + 1) * (h + 1));
  const sqSum = new Float64Array((w + 1) * (h + 1));
  for (let y = 1; y <= h; y += 1) {
    let rowSum = 0;
    let rowSq = 0;
    for (let x = 1; x <= w; x += 1) {
      const v = gray.data[(y - 1) * w + (x - 1)];
      rowSum += v;
      rowSq += v * v;
      sum[y * (w + 1) + x] = sum[(y - 1) * (w + 1) + x] + rowSum;
      sqSum[y * (w + 1) + x] = sqSum[(y - 1) * (w + 1) + x] + rowSq;
    }
  }
  return { sum, sqSum, w, h };
}

function rectSum(integral, x0, y0, x1, y1) {
  const stride = integral.w + 1;
  return integral.sum[y1 * stride + x1]
    - integral.sum[y0 * stride + x1]
    - integral.sum[y1 * stride + x0]
    + integral.sum[y0 * stride + x0];
}

function rectSqSum(integral, x0, y0, x1, y1) {
  const stride = integral.w + 1;
  return integral.sqSum[y1 * stride + x1]
    - integral.sqSum[y0 * stride + x1]
    - integral.sqSum[y1 * stride + x0]
    + integral.sqSum[y0 * stride + x0];
}

/*
 * Sauvola's local threshold. Otsu picks one number for the page, so a scan
 * with a shadow down one side loses that side entirely - either the shadow
 * turns solid black or the text in it disappears. Sauvola decides per pixel
 * from its neighbourhood's mean and standard deviation:
 *
 *     T = mean * (1 + k * (stddev / R - 1))
 *
 * with R = 128 (half the dynamic range) and k around 0.2. On a clean, evenly
 * lit page it lands within a few levels of Otsu; on a bad one it rescues it.
 */
function sauvolaThreshold(gray, windowSize, k) {
  const win = windowSize || Math.max(15, (Math.round(Math.min(gray.width, gray.height) / 40) | 1));
  const kk = k === undefined ? 0.2 : k;
  const radius = Math.max(1, (win - 1) >> 1);
  const integral = buildIntegrals(gray);
  const w = gray.width;
  const h = gray.height;
  const out = new Uint8Array(w * h);

  for (let y = 0; y < h; y += 1) {
    const y0 = Math.max(0, y - radius);
    const y1 = Math.min(h, y + radius + 1);
    for (let x = 0; x < w; x += 1) {
      const x0 = Math.max(0, x - radius);
      const x1 = Math.min(w, x + radius + 1);
      const area = (x1 - x0) * (y1 - y0);
      const s = rectSum(integral, x0, y0, x1, y1);
      const sq = rectSqSum(integral, x0, y0, x1, y1);
      const mean = s / area;
      const variance = Math.max(0, sq / area - mean * mean);
      const std = Math.sqrt(variance);
      const t = mean * (1 + kk * (std / 128 - 1));
      out[y * w + x] = gray.data[y * w + x] > t ? 255 : 0;
    }
  }
  return { width: w, height: h, channels: 1, data: out };
}

function binarize(gray, threshold) {
  const t = threshold === undefined ? otsuThreshold(gray) : threshold;
  const out = new Uint8Array(gray.data.length);
  for (let i = 0; i < gray.data.length; i += 1) out[i] = gray.data[i] > t ? 255 : 0;
  return { width: gray.width, height: gray.height, channels: 1, data: out };
}

/* ---------------------------------------------------------------- ink helpers */

/* An "ink" map: 1 where there is a mark, 0 where there is paper. */
function inkMask(binary) {
  const out = new Uint8Array(binary.width * binary.height);
  for (let i = 0; i < out.length; i += 1) out[i] = binary.data[i] < 128 ? 1 : 0;
  return out;
}

/* --------------------------------------------------------------------- deskew */

/*
 * Skew estimation by projection profile. Rotate the ink by a candidate angle,
 * sum it row by row, and score the result by the sum of squared row totals.
 * When the page is straight, every text line piles its ink into a few rows and
 * that score peaks; when it is crooked, the ink smears across many rows.
 *
 * Two things make this affordable: the search runs on a downsampled copy, and
 * it goes coarse-then-fine rather than sweeping every tenth of a degree.
 */
function scoreAngle(ink, w, h, angleRad) {
  const rows = new Float64Array(h);
  const tan = Math.tan(angleRad);
  const cx = w / 2;
  const cy = h / 2;
  for (let y = 0; y < h; y += 1) {
    for (let x = 0; x < w; x += 1) {
      if (ink[y * w + x] === 0) continue;
      // Shear rather than rotate: for angles under ~10 degrees the difference
      // is below one pixel and a shear costs one multiply.
      const yy = Math.round((y - cy) - (x - cx) * tan + cy);
      if (yy >= 0 && yy < h) rows[yy] += 1;
    }
  }
  let score = 0;
  for (let y = 0; y < h; y += 1) score += rows[y] * rows[y];
  return score;
}

function estimateSkew(binary, maxDegrees) {
  const limit = maxDegrees === undefined ? 6 : maxDegrees;

  // Work at about 600px on the long edge; skew is a page-scale property and
  // full resolution buys nothing but time.
  const long = Math.max(binary.width, binary.height);
  const factor = long > 700 ? 700 / long : 1;
  const small = factor < 1
    ? binarize(resample(binary, binary.width * factor, binary.height * factor), 128)
    : binary;

  const ink = inkMask(small);
  const w = small.width;
  const h = small.height;

  let inkCount = 0;
  for (let i = 0; i < ink.length; i += 1) inkCount += ink[i];
  // A blank or nearly blank page has no skew worth measuring, and the score
  // surface would be pure noise.
  if (inkCount < 50) return 0;

  let best = 0;
  let bestScore = -1;
  for (let deg = -limit; deg <= limit; deg += 0.5) {
    const s = scoreAngle(ink, w, h, (deg * Math.PI) / 180);
    if (s > bestScore) { bestScore = s; best = deg; }
  }
  for (let deg = best - 0.5; deg <= best + 0.5; deg += 0.1) {
    const s = scoreAngle(ink, w, h, (deg * Math.PI) / 180);
    if (s > bestScore) { bestScore = s; best = deg; }
  }
  return Math.round(best * 10) / 10;
}

/*
 * Rotate about the centre, sampling bilinearly, and pad with white. The canvas
 * grows so that no corner of the page is cut off - losing the top-right corner
 * of an invoice is losing the invoice number.
 */
function rotate(image, degrees, fill) {
  const rad = (degrees * Math.PI) / 180;
  if (Math.abs(degrees) < 0.01) {
    return {
      width: image.width, height: image.height, channels: image.channels,
      data: Uint8Array.from(image.data),
    };
  }

  const { width: sw, height: sh, channels: ch, data: src } = image;
  const cos = Math.cos(rad);
  const sin = Math.sin(rad);
  const nw = Math.ceil(Math.abs(sw * cos) + Math.abs(sh * sin));
  const nh = Math.ceil(Math.abs(sw * sin) + Math.abs(sh * cos));
  const pad = fill === undefined ? 255 : fill;
  const out = new Uint8Array(nw * nh * ch).fill(pad);

  const scx = sw / 2;
  const scy = sh / 2;
  const dcx = nw / 2;
  const dcy = nh / 2;

  for (let y = 0; y < nh; y += 1) {
    for (let x = 0; x < nw; x += 1) {
      // Inverse map: for each destination pixel find where it came from.
      const dx = x - dcx;
      const dy = y - dcy;
      const sx = dx * cos + dy * sin + scx;
      const sy = -dx * sin + dy * cos + scy;
      if (sx < 0 || sy < 0 || sx > sw - 1 || sy > sh - 1) continue;

      const x0 = Math.floor(sx);
      const y0 = Math.floor(sy);
      const x1 = Math.min(sw - 1, x0 + 1);
      const y1 = Math.min(sh - 1, y0 + 1);
      const fx = sx - x0;
      const fy = sy - y0;

      for (let c = 0; c < ch; c += 1) {
        const p00 = src[(y0 * sw + x0) * ch + c];
        const p10 = src[(y0 * sw + x1) * ch + c];
        const p01 = src[(y1 * sw + x0) * ch + c];
        const p11 = src[(y1 * sw + x1) * ch + c];
        const top = p00 + (p10 - p00) * fx;
        const bottom = p01 + (p11 - p01) * fx;
        out[(y * nw + x) * ch + c] = Math.round(top + (bottom - top) * fy);
      }
    }
  }
  return { width: nw, height: nh, channels: ch, data: out };
}

/* ----------------------------------------------------------------- despeckle */

/*
 * Remove connected ink blobs smaller than minArea. Scanner dust and JPEG
 * mosquito noise otherwise become full-confidence punctuation - a page of
 * stray commas and full stops that then confuse line grouping.
 *
 * Iterative flood fill with an explicit stack; a recursive one blows the call
 * stack on a large connected region such as a signature or a table rule.
 */
function despeckle(binary, minArea) {
  const min = minArea === undefined ? 3 : minArea;
  const w = binary.width;
  const h = binary.height;
  const ink = inkMask(binary);
  const seen = new Uint8Array(w * h);
  const out = Uint8Array.from(binary.data);
  const stack = new Int32Array(w * h);

  for (let start = 0; start < ink.length; start += 1) {
    if (ink[start] === 0 || seen[start] === 1) continue;
    let top = 0;
    stack[top] = start;
    top += 1;
    seen[start] = 1;
    const component = [];

    while (top > 0) {
      top -= 1;
      const idx = stack[top];
      component.push(idx);
      const x = idx % w;
      const y = (idx / w) | 0;
      for (let dy = -1; dy <= 1; dy += 1) {
        for (let dx = -1; dx <= 1; dx += 1) {
          if (dx === 0 && dy === 0) continue;
          const nx = x + dx;
          const ny = y + dy;
          if (nx < 0 || ny < 0 || nx >= w || ny >= h) continue;
          const n = ny * w + nx;
          if (ink[n] === 1 && seen[n] === 0) {
            seen[n] = 1;
            stack[top] = n;
            top += 1;
          }
        }
      }
    }

    if (component.length < min) {
      for (let i = 0; i < component.length; i += 1) out[component[i]] = 255;
    }
  }
  return { width: w, height: h, channels: 1, data: out };
}

/* -------------------------------------------------------------- border crop */

/*
 * Scanners produce a black margin where the lid did not close, or a grey band
 * from the platen edge. Left in, it dominates the binarisation histogram and
 * the skew estimate. This finds the tight box around the ink, then gives back
 * a small white margin so glyphs are not touching the edge - Tesseract reads
 * text at the very edge of an image noticeably worse.
 */
function cropToContent(binary, marginPx) {
  const margin = marginPx === undefined ? 8 : marginPx;
  const w = binary.width;
  const h = binary.height;
  const ink = inkMask(binary);

  // Row/column ink counts, so a solid black scan edge (a full row of ink) can
  // be told apart from a line of text (a partial row).
  const rowCount = new Uint32Array(h);
  const colCount = new Uint32Array(w);
  for (let y = 0; y < h; y += 1) {
    for (let x = 0; x < w; x += 1) {
      if (ink[y * w + x] === 1) { rowCount[y] += 1; colCount[x] += 1; }
    }
  }

  const rowFull = w * 0.95;
  const colFull = h * 0.95;

  let top = 0;
  while (top < h && (rowCount[top] === 0 || rowCount[top] >= rowFull)) top += 1;
  let bottom = h - 1;
  while (bottom > top && (rowCount[bottom] === 0 || rowCount[bottom] >= rowFull)) bottom -= 1;
  let left = 0;
  while (left < w && (colCount[left] === 0 || colCount[left] >= colFull)) left += 1;
  let right = w - 1;
  while (right > left && (colCount[right] === 0 || colCount[right] >= colFull)) right -= 1;

  if (right <= left || bottom <= top) {
    return { image: binary, box: { x: 0, y: 0, width: w, height: h } };
  }

  const x0 = Math.max(0, left - margin);
  const y0 = Math.max(0, top - margin);
  const x1 = Math.min(w, right + 1 + margin);
  const y1 = Math.min(h, bottom + 1 + margin);
  const cw = x1 - x0;
  const chh = y1 - y0;
  const out = new Uint8Array(cw * chh);

  for (let y = 0; y < chh; y += 1) {
    for (let x = 0; x < cw; x += 1) {
      out[y * cw + x] = binary.data[(y + y0) * w + (x + x0)];
    }
  }
  return {
    image: { width: cw, height: chh, channels: 1, data: out },
    box: { x: x0, y: y0, width: cw, height: chh },
  };
}

/* ------------------------------------------------------------ the whole recipe */

/*
 * The default pipeline, in the order that matters:
 *
 *   1. greyscale        - one channel, correctly weighted
 *   2. contrast stretch - so the threshold step sees a full range
 *   3. upscale to DPI   - the single biggest win, and it must happen on the
 *                         GREY image: upscaling after thresholding just makes
 *                         big jagged pixels
 *   4. threshold        - Sauvola by default, Otsu when asked
 *   5. deskew           - measured on the binary, applied to the binary
 *   6. despeckle        - after deskew, because rotation resampling creates
 *                         its own faint speckle at stroke edges
 *   7. crop             - last, once the page is straight and clean
 *
 * Every stage can be switched off. `steps` records what actually ran, and the
 * geometry needed to map a coordinate in the processed image back to the
 * original page travels back in `transform`.
 */
function preparePage(image, options) {
  const opts = options || {};
  const targetDpi = opts.targetDpi || 300;
  const sourceDpi = opts.sourceDpi || 300;
  const steps = [];

  let work = toGrayscale(image);
  steps.push('grayscale');

  if (opts.stretch !== false) {
    work = stretchContrast(work, opts.stretchPercent);
    steps.push('contrast-stretch');
  }

  let scale = 1;
  if (opts.scale !== false) {
    const scaled = scaleToDpi(work, sourceDpi, targetDpi, opts.maxDimension);
    work = scaled.image;
    scale = scaled.scale;
    if (scale !== 1) steps.push('scale-to-' + targetDpi + 'dpi(x' + scale.toFixed(2) + ')');
  }

  let threshold = null;
  if (opts.binarize !== false) {
    if (opts.method === 'otsu') {
      threshold = otsuThreshold(work);
      work = binarize(work, threshold);
      steps.push('otsu(' + threshold + ')');
    } else {
      work = sauvolaThreshold(work, opts.window, opts.k);
      steps.push('sauvola');
    }
  }

  let skew = 0;
  // Geometry of the rotation, kept so a word box found on the straightened
  // page can be mapped back to where it sits on the page the user has.
  let preRotate = { width: work.width, height: work.height };
  let postRotate = { width: work.width, height: work.height };
  let rotateDegrees = 0;

  if (opts.deskew !== false && opts.binarize !== false) {
    skew = estimateSkew(work, opts.maxSkew);
    if (Math.abs(skew) >= 0.2) {
      preRotate = { width: work.width, height: work.height };
      rotateDegrees = -skew;
      work = binarize(rotate(work, rotateDegrees, 255), 128);
      postRotate = { width: work.width, height: work.height };
      steps.push('deskew(' + skew.toFixed(1) + ')');
    }
  }

  if (opts.despeckle !== false && opts.binarize !== false) {
    const area = opts.minSpeckle === undefined
      ? Math.max(2, Math.round((targetDpi / 300) * 3))
      : opts.minSpeckle;
    work = despeckle(work, area);
    steps.push('despeckle(<' + area + 'px)');
  }

  let cropBox = { x: 0, y: 0, width: work.width, height: work.height };
  if (opts.crop === true && opts.binarize !== false) {
    const cropped = cropToContent(work, opts.cropMargin);
    work = cropped.image;
    cropBox = cropped.box;
    steps.push('crop');
  }

  return {
    image: work,
    steps,
    transform: {
      scale,
      skewDegrees: skew,
      rotateDegrees,
      preRotate,
      postRotate,
      cropX: cropBox.x,
      cropY: cropBox.y,
      sourceWidth: image.width,
      sourceHeight: image.height,
      sourceDpi,
      effectiveDpi: sourceDpi * scale,
    },
    threshold,
  };
}

module.exports = {
  toGrayscale,
  resample,
  scaleToDpi,
  histogram,
  stretchContrast,
  otsuThreshold,
  sauvolaThreshold,
  binarize,
  inkMask,
  estimateSkew,
  rotate,
  despeckle,
  cropToContent,
  preparePage,
};

/*
 * Map a point on the PROCESSED image back to where it sits on the image that
 * came in. Recognition happens on a page that has been enlarged, straightened
 * and cropped; the caller wants coordinates on the page they actually have, so
 * every step has to be undone in reverse order.
 *
 * Without this, deskewing a page would silently move every reported word by
 * several millimetres - which is worse than not deskewing at all, because the
 * error is invisible until someone tries to find the value on the original.
 */
function mapPointToSource(transform, x, y) {
  let px = x;
  let py = y;

  // 3. undo the crop
  px += transform.cropX || 0;
  py += transform.cropY || 0;

  // 2. undo the rotation, about the same centres rotate() used
  const deg = transform.rotateDegrees || 0;
  if (deg !== 0) {
    const rad = (deg * Math.PI) / 180;
    const cos = Math.cos(rad);
    const sin = Math.sin(rad);
    const dcx = transform.postRotate.width / 2;
    const dcy = transform.postRotate.height / 2;
    const scx = transform.preRotate.width / 2;
    const scy = transform.preRotate.height / 2;
    const dx = px - dcx;
    const dy = py - dcy;
    px = dx * cos + dy * sin + scx;
    py = -dx * sin + dy * cos + scy;
  }

  // 1. undo the scaling
  const scale = transform.scale || 1;
  return { x: px / scale, y: py / scale };
}

/*
 * The same for a box. A rectangle that was axis-aligned on the straightened
 * page is a tilted rectangle on the original, so all four corners are mapped
 * and the box that contains them is returned. It is very slightly larger than
 * the true glyph extent, and that is the honest answer.
 */
function mapBoxToSource(transform, x, y, w, h) {
  const corners = [
    mapPointToSource(transform, x, y),
    mapPointToSource(transform, x + w, y),
    mapPointToSource(transform, x, y + h),
    mapPointToSource(transform, x + w, y + h),
  ];
  let x0 = Infinity;
  let y0 = Infinity;
  let x1 = -Infinity;
  let y1 = -Infinity;
  for (let i = 0; i < corners.length; i += 1) {
    if (corners[i].x < x0) x0 = corners[i].x;
    if (corners[i].y < y0) y0 = corners[i].y;
    if (corners[i].x > x1) x1 = corners[i].x;
    if (corners[i].y > y1) y1 = corners[i].y;
  }
  return { x: x0, y: y0, w: x1 - x0, h: y1 - y0 };
}

module.exports.mapPointToSource = mapPointToSource;
module.exports.mapBoxToSource = mapBoxToSource;
