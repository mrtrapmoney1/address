'use strict';

/*
 * imagefile.js - what kind of image is this, and can Tesseract.js read it?
 *
 * Tesseract.js accepts bmp, jpg, png, pbm, webp and non-animated gif, and
 * nothing else - notably NOT tiff, which is the format most document scanners
 * produce by default, and not pdf. Finding that out from a cryptic decode
 * failure halfway through a four-hundred page batch is no way to learn it, so
 * every file is identified up front by its bytes and the unsupported ones are
 * reported by name with the reason.
 *
 * Identification is by magic number, never by extension. Scanners and mail
 * systems rename files constantly; the bytes do not lie.
 */

const { probePng } = require('./png.js');

/* JPEG: the frame header carries the true dimensions, and it is not at a
 * fixed offset - the markers before it vary in number and length. */
function probeJpeg(buf) {
  if (buf.length < 4 || buf[0] !== 0xff || buf[1] !== 0xd8) return null;
  let pos = 2;
  while (pos + 9 < buf.length) {
    if (buf[pos] !== 0xff) { pos += 1; continue; }
    const marker = buf[pos + 1];
    // Standalone markers carry no length field.
    if (marker === 0xd8 || marker === 0xd9 || (marker >= 0xd0 && marker <= 0xd7) || marker === 0x01) {
      pos += 2;
      continue;
    }
    const length = buf.readUInt16BE(pos + 2);
    // SOF0..SOF15, skipping the DHT/JPG/DAC markers that share the range.
    const isSof = (marker >= 0xc0 && marker <= 0xcf)
      && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc;
    if (isSof) {
      return {
        format: 'jpeg',
        height: buf.readUInt16BE(pos + 5),
        width: buf.readUInt16BE(pos + 7),
        components: buf[pos + 9],
        progressive: marker === 0xc2,
      };
    }
    pos += 2 + length;
  }
  return { format: 'jpeg', width: 0, height: 0 };
}

function probeGif(buf) {
  if (buf.length < 10) return null;
  const sig = buf.toString('ascii', 0, 6);
  if (sig !== 'GIF87a' && sig !== 'GIF89a') return null;
  return { format: 'gif', width: buf.readUInt16LE(6), height: buf.readUInt16LE(8) };
}

function probeBmp(buf) {
  if (buf.length < 26 || buf[0] !== 0x42 || buf[1] !== 0x4d) return null;
  return {
    format: 'bmp',
    width: buf.readInt32LE(18),
    height: Math.abs(buf.readInt32LE(22)),
  };
}

function probeWebp(buf) {
  if (buf.length < 16) return null;
  if (buf.toString('ascii', 0, 4) !== 'RIFF' || buf.toString('ascii', 8, 12) !== 'WEBP') return null;
  const chunk = buf.toString('ascii', 12, 16);
  if (chunk === 'VP8X' && buf.length >= 30) {
    return {
      format: 'webp',
      width: 1 + (buf[24] | (buf[25] << 8) | (buf[26] << 16)),
      height: 1 + (buf[27] | (buf[28] << 8) | (buf[29] << 16)),
    };
  }
  return { format: 'webp', width: 0, height: 0 };
}

/* TIFF is not readable by Tesseract.js, but it must still be RECOGNISED so
 * the message can say "convert this" rather than "this file is broken". */
function probeTiff(buf) {
  if (buf.length < 8) return null;
  const le = buf[0] === 0x49 && buf[1] === 0x49 && buf[2] === 0x2a && buf[3] === 0x00;
  const be = buf[0] === 0x4d && buf[1] === 0x4d && buf[2] === 0x00 && buf[3] === 0x2a;
  if (!le && !be) return null;
  return { format: 'tiff', width: 0, height: 0 };
}

function probePdf(buf) {
  if (buf.length < 5) return null;
  return buf.toString('ascii', 0, 5) === '%PDF-' ? { format: 'pdf', width: 0, height: 0 } : null;
}

/* Netpbm: P1..P6, an ASCII header of whitespace-separated tokens. */
function probePnm(buf) {
  if (buf.length < 8 || buf[0] !== 0x50) return null;
  const type = buf[1] - 0x30;
  if (type < 1 || type > 6) return null;
  const head = buf.toString('ascii', 0, Math.min(buf.length, 128))
    .replace(/#[^\n]*/g, ' ')
    .split(/\s+/)
    .filter(Boolean);
  return {
    format: type === 4 || type === 1 ? 'pbm' : 'pnm',
    width: Number(head[1]) || 0,
    height: Number(head[2]) || 0,
  };
}

/* Formats Tesseract.js will decode directly. */
const TESSERACT_READABLE = { png: true, jpeg: true, bmp: true, gif: true, webp: true, pbm: true, pnm: true };

function identify(buffer) {
  const buf = Buffer.isBuffer(buffer) ? buffer : Buffer.from(buffer);
  const probes = [probePng, probeJpeg, probeGif, probeBmp, probeWebp, probeTiff, probePdf, probePnm];
  for (let i = 0; i < probes.length; i += 1) {
    const info = probes[i](buf);
    if (info !== null && info !== undefined) {
      info.bytes = buf.length;
      info.readable = TESSERACT_READABLE[info.format] === true;
      if (!info.readable) {
        info.reason = info.format === 'pdf'
          ? 'PDF is not an image. Extract or render its pages first.'
          : info.format === 'tiff'
            ? 'Tesseract.js cannot read TIFF. Convert the page to PNG first.'
            : 'unsupported image format';
      }
      return info;
    }
  }
  return {
    format: 'unknown', width: 0, height: 0, bytes: buf.length, readable: false,
    reason: 'the file does not begin with any image signature this tool knows',
  };
}

/*
 * Only PNG can be decoded to pixels in-process, so only PNG can be
 * preprocessed. Everything else Tesseract.js accepts is passed through
 * untouched - still recognised, just without the upscale-and-threshold pass.
 */
function canPreprocess(info) {
  return info.format === 'png';
}

module.exports = {
  identify,
  canPreprocess,
  probeJpeg,
  probeGif,
  probeBmp,
  probeWebp,
  probeTiff,
  probePdf,
  probePnm,
  probePng,
  TESSERACT_READABLE,
};
