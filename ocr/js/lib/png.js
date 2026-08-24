'use strict';

/*
 * png.js - a PNG decoder and encoder built on nothing but Node's own zlib.
 *
 * Why this exists: everything useful you can do to a scan before recognition
 * (upscale it, grey it, straighten it, threshold it) needs the actual pixels.
 * Every JS library that gives you pixels wants a native build step. A scan is
 * a very tame PNG - 8-bit grey or RGB, not interlaced - so decoding it here
 * costs a few hundred lines once and buys a toolchain that installs with
 * `npm install` and nothing else, on any machine, forever.
 *
 * Decoding supports colour types 0/2/3/4/6, bit depths 1/2/4/8/16, and both
 * interlace methods. Encoding writes 8-bit greyscale or 8-bit RGB, which is
 * all the pipeline ever needs to hand to Tesseract.
 */

const zlib = require('zlib');

const PNG_SIGNATURE = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);

/* Bytes per complete pixel, per colour type. */
const CHANNELS = { 0: 1, 2: 3, 3: 1, 4: 2, 6: 4 };

/* --------------------------------------------------------------------- CRC32 */

const CRC_TABLE = (function buildCrcTable() {
  const table = new Int32Array(256);
  for (let n = 0; n < 256; n += 1) {
    let c = n;
    for (let k = 0; k < 8; k += 1) {
      c = (c & 1) ? (0xedb88320 ^ (c >>> 1)) : (c >>> 1);
    }
    table[n] = c;
  }
  return table;
}());

function crc32(buf) {
  let c = 0xffffffff;
  for (let i = 0; i < buf.length; i += 1) {
    c = CRC_TABLE[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
  }
  return (c ^ 0xffffffff) >>> 0;
}

/* ------------------------------------------------------------------ unfilter */

/*
 * Each scanline is prefixed with a filter byte. Undoing the filter needs the
 * reconstructed bytes of this line and the line above, so this runs in place
 * over the whole inflated block, one line at a time, top to bottom.
 *
 * bpp here is the "byte offset to the pixel to the left", which for sub-byte
 * bit depths is 1, not a fraction. That detail is the classic PNG decoder bug.
 */
function unfilter(raw, width, height, bitDepth, channels) {
  const bitsPerPixel = bitDepth * channels;
  const bpp = Math.max(1, Math.ceil(bitsPerPixel / 8));
  const lineBytes = Math.ceil((width * bitsPerPixel) / 8);
  const out = Buffer.alloc(lineBytes * height);

  let pos = 0;
  for (let y = 0; y < height; y += 1) {
    if (pos >= raw.length) throw new Error('png: truncated image data');
    const filter = raw[pos];
    pos += 1;
    const lineStart = y * lineBytes;
    const prevStart = (y - 1) * lineBytes;

    for (let i = 0; i < lineBytes; i += 1) {
      const x = raw[pos + i];
      const a = i >= bpp ? out[lineStart + i - bpp] : 0;     // left
      const b = y > 0 ? out[prevStart + i] : 0;              // up
      const c = (y > 0 && i >= bpp) ? out[prevStart + i - bpp] : 0; // up-left
      let value;
      switch (filter) {
        case 0: value = x; break;
        case 1: value = x + a; break;
        case 2: value = x + b; break;
        case 3: value = x + ((a + b) >> 1); break;
        case 4: {
          const p = a + b - c;
          const pa = Math.abs(p - a);
          const pb = Math.abs(p - b);
          const pc = Math.abs(p - c);
          let pred;
          if (pa <= pb && pa <= pc) pred = a;
          else if (pb <= pc) pred = b;
          else pred = c;
          value = x + pred;
          break;
        }
        default:
          throw new Error('png: unknown filter type ' + filter + ' on row ' + y);
      }
      out[lineStart + i] = value & 0xff;
    }
    pos += lineBytes;
  }
  return out;
}

/* ---------------------------------------------------------- bit unpacking */

/*
 * Turn a packed scanline block into one byte (or one 16-bit word) per sample,
 * scaled up to the full 0..255 range so that everything downstream can assume
 * 8 bits and stop caring what the file happened to use.
 */
function unpackSamples(data, width, height, bitDepth, channels, isPalette) {
  const samplesPerLine = width * channels;
  const out = new Uint8Array(samplesPerLine * height);
  const lineBits = width * channels * bitDepth;
  const lineBytes = Math.ceil(lineBits / 8);

  if (bitDepth === 8) {
    for (let y = 0; y < height; y += 1) {
      data.copy(out, y * samplesPerLine, y * lineBytes, y * lineBytes + samplesPerLine);
    }
    return out;
  }

  if (bitDepth === 16) {
    // Take the high byte: an OCR pipeline gains nothing from 16-bit depth.
    for (let y = 0; y < height; y += 1) {
      for (let i = 0; i < samplesPerLine; i += 1) {
        out[y * samplesPerLine + i] = data[y * lineBytes + i * 2];
      }
    }
    return out;
  }

  // 1, 2 or 4 bits per sample, packed most-significant-bit first.
  const max = (1 << bitDepth) - 1;
  for (let y = 0; y < height; y += 1) {
    let bitPos = 0;
    for (let i = 0; i < samplesPerLine; i += 1) {
      const byteIndex = y * lineBytes + (bitPos >> 3);
      const shift = 8 - bitDepth - (bitPos & 7);
      const v = (data[byteIndex] >> shift) & max;
      // Palette indices must stay indices; real samples get scaled to 0..255.
      out[y * samplesPerLine + i] = isPalette ? v : Math.round((v * 255) / max);
      bitPos += bitDepth;
    }
  }
  return out;
}

/* ------------------------------------------------------------------- Adam7 */

const ADAM7 = [
  // xStart, yStart, xStep, yStep
  [0, 0, 8, 8], [4, 0, 8, 8], [0, 4, 4, 8], [2, 0, 4, 4],
  [0, 2, 2, 4], [1, 0, 2, 2], [0, 1, 1, 2],
];

function deinterlace(raw, width, height, bitDepth, channels) {
  const bitsPerPixel = bitDepth * channels;
  const full = Buffer.alloc(Math.ceil((width * bitsPerPixel) / 8) * height);
  const fullLineBytes = Math.ceil((width * bitsPerPixel) / 8);
  let offset = 0;

  for (let p = 0; p < ADAM7.length; p += 1) {
    const [xs, ys, xstep, ystep] = ADAM7[p];
    const passW = Math.ceil((width - xs) / xstep);
    const passH = Math.ceil((height - ys) / ystep);
    if (passW <= 0 || passH <= 0) continue;

    const passLineBytes = Math.ceil((passW * bitsPerPixel) / 8);
    const passRaw = raw.slice(offset, offset + (passLineBytes + 1) * passH);
    offset += (passLineBytes + 1) * passH;
    const passData = unfilter(passRaw, passW, passH, bitDepth, channels);
    const passSamples = unpackSamples(passData, passW, passH, bitDepth, channels, true);

    // Write the pass's pixels back into the full grid at 8 bits per sample,
    // then repack below. Sub-byte interlaced PNGs are vanishingly rare in
    // scanning, so clarity beats cleverness here.
    for (let y = 0; y < passH; y += 1) {
      for (let x = 0; x < passW; x += 1) {
        const destX = xs + x * xstep;
        const destY = ys + y * ystep;
        for (let ch = 0; ch < channels; ch += 1) {
          const v = passSamples[(y * passW + x) * channels + ch];
          if (bitDepth === 8) {
            full[destY * fullLineBytes + destX * channels + ch] = v;
          } else if (bitDepth === 16) {
            full[destY * fullLineBytes + (destX * channels + ch) * 2] = v;
          } else {
            const bitPos = (destX * channels + ch) * bitDepth;
            const byteIndex = destY * fullLineBytes + (bitPos >> 3);
            const shift = 8 - bitDepth - (bitPos & 7);
            full[byteIndex] |= (v & ((1 << bitDepth) - 1)) << shift;
          }
        }
      }
    }
  }
  return full;
}

/* ------------------------------------------------------------------ decoding */

/*
 * Returns { width, height, channels, data } where data is a Uint8Array of
 * 8-bit samples, `channels` per pixel, row-major, no padding. Alpha is
 * composited onto white, because a scan with a transparent background that
 * gets thresholded as black is a page of solid ink.
 */
function decodePng(buffer) {
  const buf = Buffer.isBuffer(buffer) ? buffer : Buffer.from(buffer);
  if (buf.length < 8 || !buf.slice(0, 8).equals(PNG_SIGNATURE)) {
    throw new Error('png: not a PNG file (bad signature)');
  }

  let pos = 8;
  let ihdr = null;
  let palette = null;
  let transparency = null;
  const idat = [];

  while (pos + 8 <= buf.length) {
    const length = buf.readUInt32BE(pos);
    const type = buf.toString('ascii', pos + 4, pos + 8);
    const dataStart = pos + 8;
    const dataEnd = dataStart + length;
    if (dataEnd + 4 > buf.length) throw new Error('png: truncated chunk ' + type);
    const data = buf.slice(dataStart, dataEnd);

    if (type === 'IHDR') {
      ihdr = {
        width: data.readUInt32BE(0),
        height: data.readUInt32BE(4),
        bitDepth: data[8],
        colorType: data[9],
        compression: data[10],
        filter: data[11],
        interlace: data[12],
      };
      if (ihdr.width === 0 || ihdr.height === 0) throw new Error('png: zero-sized image');
      if (ihdr.compression !== 0) throw new Error('png: unsupported compression method');
      if (CHANNELS[ihdr.colorType] === undefined) {
        throw new Error('png: unsupported colour type ' + ihdr.colorType);
      }
    } else if (type === 'PLTE') {
      palette = data;
    } else if (type === 'tRNS') {
      transparency = data;
    } else if (type === 'IDAT') {
      idat.push(data);
    } else if (type === 'IEND') {
      break;
    }
    pos = dataEnd + 4; // skip the chunk CRC
  }

  if (ihdr === null) throw new Error('png: no IHDR chunk');
  if (idat.length === 0) throw new Error('png: no image data');

  const inflated = zlib.inflateSync(Buffer.concat(idat));
  const channels = CHANNELS[ihdr.colorType];
  const { width, height, bitDepth, colorType, interlace } = ihdr;

  let unfiltered;
  if (interlace === 1) {
    unfiltered = deinterlace(inflated, width, height, bitDepth, channels);
  } else if (interlace === 0) {
    unfiltered = unfilter(inflated, width, height, bitDepth, channels);
  } else {
    throw new Error('png: unknown interlace method ' + interlace);
  }

  const samples = unpackSamples(unfiltered, width, height, bitDepth, channels, colorType === 3);

  /* Expand palette entries and flatten alpha onto white. */
  if (colorType === 3) {
    if (palette === null) throw new Error('png: palette image with no PLTE chunk');
    const out = new Uint8Array(width * height * 3);
    for (let i = 0; i < width * height; i += 1) {
      const idx = samples[i];
      const alpha = (transparency !== null && idx < transparency.length) ? transparency[idx] : 255;
      const a = alpha / 255;
      for (let ch = 0; ch < 3; ch += 1) {
        const v = palette[idx * 3 + ch];
        out[i * 3 + ch] = Math.round(v * a + 255 * (1 - a));
      }
    }
    return { width, height, channels: 3, data: out };
  }

  if (colorType === 4 || colorType === 6) {
    const colourChannels = colorType === 4 ? 1 : 3;
    const out = new Uint8Array(width * height * colourChannels);
    for (let i = 0; i < width * height; i += 1) {
      const a = samples[i * channels + colourChannels] / 255;
      for (let ch = 0; ch < colourChannels; ch += 1) {
        const v = samples[i * channels + ch];
        out[i * colourChannels + ch] = Math.round(v * a + 255 * (1 - a));
      }
    }
    return { width, height, channels: colourChannels, data: out };
  }

  return { width, height, channels, data: samples };
}

/* ------------------------------------------------------------------ encoding */

function chunk(type, data) {
  const out = Buffer.alloc(data.length + 12);
  out.writeUInt32BE(data.length, 0);
  out.write(type, 4, 4, 'ascii');
  data.copy(out, 8);
  out.writeUInt32BE(crc32(out.slice(4, 8 + data.length)), 8 + data.length);
  return out;
}

/*
 * Pick a filter per scanline by the sum-of-absolute-differences heuristic the
 * PNG specification itself recommends. A binarised page compresses several
 * times smaller this way than with no filtering, which matters when a run
 * writes three hundred page images to disk.
 */
function filterLine(line, prev, bpp, out, outOffset) {
  const n = line.length;
  const candidates = [];

  for (let type = 0; type < 5; type += 1) {
    const buf = Buffer.alloc(n);
    let score = 0;
    for (let i = 0; i < n; i += 1) {
      const x = line[i];
      const a = i >= bpp ? line[i - bpp] : 0;
      const b = prev !== null ? prev[i] : 0;
      const c = (prev !== null && i >= bpp) ? prev[i - bpp] : 0;
      let v;
      switch (type) {
        case 0: v = x; break;
        case 1: v = x - a; break;
        case 2: v = x - b; break;
        case 3: v = x - ((a + b) >> 1); break;
        default: {
          const p = a + b - c;
          const pa = Math.abs(p - a);
          const pb = Math.abs(p - b);
          const pc = Math.abs(p - c);
          let pred;
          if (pa <= pb && pa <= pc) pred = a;
          else if (pb <= pc) pred = b;
          else pred = c;
          v = x - pred;
        }
      }
      buf[i] = v & 0xff;
      // Treat the byte as signed when scoring, per the spec's heuristic.
      score += buf[i] < 128 ? buf[i] : 256 - buf[i];
    }
    candidates.push({ type, buf, score });
  }

  let best = candidates[0];
  for (let i = 1; i < candidates.length; i += 1) {
    if (candidates[i].score < best.score) best = candidates[i];
  }
  out[outOffset] = best.type;
  best.buf.copy(out, outOffset + 1);
}

/*
 * image: { width, height, channels, data }. channels must be 1 (grey) or 3 (RGB).
 */
function encodePng(image, options) {
  const opts = options || {};
  const { width, height, channels, data } = image;
  if (channels !== 1 && channels !== 3) {
    throw new Error('png: encodePng supports 1 or 3 channels, got ' + channels);
  }
  if (data.length < width * height * channels) {
    throw new Error('png: pixel buffer is too small for ' + width + 'x' + height);
  }

  const lineBytes = width * channels;
  const rawSize = (lineBytes + 1) * height;
  const raw = Buffer.alloc(rawSize);
  const lineBuf = Buffer.alloc(lineBytes);
  let prev = null;

  for (let y = 0; y < height; y += 1) {
    for (let i = 0; i < lineBytes; i += 1) lineBuf[i] = data[y * lineBytes + i];
    filterLine(lineBuf, prev, channels, raw, y * (lineBytes + 1));
    prev = Buffer.from(lineBuf);
  }

  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(width, 0);
  ihdr.writeUInt32BE(height, 4);
  ihdr[8] = 8;                              // bit depth
  ihdr[9] = channels === 1 ? 0 : 2;         // colour type: grey or truecolour
  ihdr[10] = 0;                             // compression: deflate
  ihdr[11] = 0;                             // filter: adaptive
  ihdr[12] = 0;                             // interlace: none

  const chunks = [PNG_SIGNATURE, chunk('IHDR', ihdr)];

  /*
   * Writing pHYs matters more than it looks. Tesseract warns "Invalid
   * resolution 0 dpi, using 70 instead" and then makes worse decisions about
   * what is a heading and what is body text. Stamping the real DPI into the
   * file removes the guess.
   */
  if (opts.dpi) {
    const perMetre = Math.round(opts.dpi / 0.0254);
    const phys = Buffer.alloc(9);
    phys.writeUInt32BE(perMetre, 0);
    phys.writeUInt32BE(perMetre, 4);
    phys[8] = 1;                            // unit: metres
    chunks.push(chunk('pHYs', phys));
  }

  const level = opts.level === undefined ? 6 : opts.level;
  chunks.push(chunk('IDAT', zlib.deflateSync(raw, { level })));
  chunks.push(chunk('IEND', Buffer.alloc(0)));
  return Buffer.concat(chunks);
}

/* Read width/height without inflating anything. */
function probePng(buffer) {
  const buf = Buffer.isBuffer(buffer) ? buffer : Buffer.from(buffer);
  if (buf.length < 24 || !buf.slice(0, 8).equals(PNG_SIGNATURE)) return null;
  if (buf.toString('ascii', 12, 16) !== 'IHDR') return null;
  return {
    format: 'png',
    width: buf.readUInt32BE(16),
    height: buf.readUInt32BE(20),
    bitDepth: buf[24],
    colorType: buf[25],
  };
}

module.exports = {
  decodePng,
  encodePng,
  probePng,
  crc32,
  PNG_SIGNATURE,
};
