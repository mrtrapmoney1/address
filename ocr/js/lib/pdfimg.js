'use strict';

/*
 * pdfimg.js - pulling the page images out of a scanned PDF, in pure JS.
 *
 * WHY NOT JUST RENDER THE PDF?
 *
 * A scanned PDF is not a document that happens to have no text. It is a
 * wrapper: one photograph per page, and nothing else. Rendering it means
 * decoding that photograph, painting it onto a canvas at some chosen
 * resolution, and re-encoding - which needs a rendering library, and which
 * resamples the very pixels the OCR engine is about to read. Every resample
 * softens the letter edges a little.
 *
 * Lifting the embedded image out instead gives the scanner's original pixels,
 * at their original resolution, with no resampling at all. For a JPEG scan it
 * is not even a decode: the bytes are already a complete JPEG file and can be
 * written straight to disk, which Tesseract.js reads natively.
 *
 * WHAT THIS HANDLES, AND WHAT IT DOES NOT
 *
 * Handled: DCTDecode (JPEG) passthrough, FlateDecode raw samples re-encoded
 * as PNG, in DeviceGray, DeviceRGB, and Indexed colour, at 1 or 8 bits per
 * component - which between them is the overwhelming majority of scanned
 * PDFs produced by office equipment.
 *
 * Not handled: CCITTFax (fax G3/G4), JBIG2, and JPX (JPEG 2000). These are
 * real formats that real scanners emit, and decoding them properly is a large
 * job each. Rather than half-decode them and hand the OCR engine something
 * subtly wrong, they are reported by name so the caller can fall back to the
 * Windows renderer - which handles all of them.
 *
 * This is a reader for the scanned case, not a general PDF implementation.
 * The PDF parser here is deliberately small and refuses anything it does not
 * fully understand.
 */

const zlib = require('zlib');
const { encodePng } = require('./png.js');

/* ------------------------------------------------------------------- lexing */

const WHITESPACE = new Set([0x00, 0x09, 0x0a, 0x0c, 0x0d, 0x20]);
const DELIMITERS = new Set([0x28, 0x29, 0x3c, 0x3e, 0x5b, 0x5d, 0x7b, 0x7d, 0x2f, 0x25]);

function isWhite(c) { return WHITESPACE.has(c); }
function isDelim(c) { return DELIMITERS.has(c); }
function isRegular(c) { return !isWhite(c) && !isDelim(c); }

/*
 * The object types. Plain JS shapes rather than classes, because they are
 * only ever inspected, never subclassed.
 */
function Ref(num, gen) { this.num = num; this.gen = gen; }
function Name(value) { this.name = value; }
function PdfStream(dict, start, length) {
  this.dict = dict;
  this.start = start;
  this.length = length;
}

function Lexer(buf, pos) {
  this.buf = buf;
  this.pos = pos || 0;
}

Lexer.prototype.skipWhite = function skipWhite() {
  while (this.pos < this.buf.length) {
    const c = this.buf[this.pos];
    if (isWhite(c)) { this.pos += 1; continue; }
    // A comment runs to the end of the line and is whitespace for all purposes.
    if (c === 0x25) {
      while (this.pos < this.buf.length && this.buf[this.pos] !== 0x0a && this.buf[this.pos] !== 0x0d) {
        this.pos += 1;
      }
      continue;
    }
    break;
  }
};

Lexer.prototype.readToken = function readToken() {
  const start = this.pos;
  while (this.pos < this.buf.length && isRegular(this.buf[this.pos])) this.pos += 1;
  return this.buf.toString('latin1', start, this.pos);
};

Lexer.prototype.readName = function readName() {
  this.pos += 1; // the '/'
  let out = '';
  while (this.pos < this.buf.length && isRegular(this.buf[this.pos])) {
    let ch = this.buf[this.pos];
    // #xx is a hex escape inside a name.
    if (ch === 0x23 && this.pos + 2 < this.buf.length) {
      const hex = this.buf.toString('latin1', this.pos + 1, this.pos + 3);
      const code = parseInt(hex, 16);
      if (!isNaN(code)) {
        out += String.fromCharCode(code);
        this.pos += 3;
        continue;
      }
    }
    out += String.fromCharCode(ch);
    this.pos += 1;
  }
  return new Name(out);
};

Lexer.prototype.readLiteralString = function readLiteralString() {
  this.pos += 1; // the '('
  let depth = 1;
  let out = '';
  while (this.pos < this.buf.length) {
    const c = this.buf[this.pos];
    if (c === 0x5c) {                      // backslash
      this.pos += 1;
      const esc = this.buf[this.pos];
      if (esc === undefined) break;
      out += String.fromCharCode(esc);
      this.pos += 1;
      continue;
    }
    if (c === 0x28) depth += 1;
    if (c === 0x29) {
      depth -= 1;
      if (depth === 0) { this.pos += 1; break; }
    }
    out += String.fromCharCode(c);
    this.pos += 1;
  }
  return out;
};

Lexer.prototype.readHexString = function readHexString() {
  this.pos += 1; // the '<'
  let hex = '';
  while (this.pos < this.buf.length && this.buf[this.pos] !== 0x3e) {
    const c = this.buf[this.pos];
    if (!isWhite(c)) hex += String.fromCharCode(c);
    this.pos += 1;
  }
  this.pos += 1;
  if (hex.length % 2 === 1) hex += '0';
  let out = '';
  for (let i = 0; i < hex.length; i += 2) {
    out += String.fromCharCode(parseInt(hex.substr(i, 2), 16));
  }
  return out;
};

/*
 * Parse one object. `resolveLength` is used only for streams, whose /Length
 * is frequently an indirect reference that has to be looked up before the
 * stream data can be skipped.
 */
Lexer.prototype.parseObject = function parseObject(resolveLength) {
  this.skipWhite();
  if (this.pos >= this.buf.length) return null;

  const c = this.buf[this.pos];

  if (c === 0x2f) return this.readName();
  if (c === 0x28) return this.readLiteralString();

  if (c === 0x3c) {
    if (this.buf[this.pos + 1] === 0x3c) return this.parseDict(resolveLength);
    return this.readHexString();
  }

  if (c === 0x5b) {                        // array
    this.pos += 1;
    const out = [];
    for (;;) {
      this.skipWhite();
      if (this.pos >= this.buf.length) break;
      if (this.buf[this.pos] === 0x5d) { this.pos += 1; break; }
      const item = this.parseObject(resolveLength);
      if (item === null) break;
      out.push(item);
    }
    return out;
  }

  if (c === 0x5d || c === 0x3e || c === 0x7d || c === 0x29) {
    // A closing delimiter where an object was expected: let the caller stop.
    return null;
  }

  const save = this.pos;
  const token = this.readToken();
  if (token === '') { this.pos += 1; return null; }
  if (token === 'true') return true;
  if (token === 'false') return false;
  if (token === 'null') return null;

  /*
   * "12 0 R" is a reference and "12 0 obj" starts an object; both begin with
   * two integers. Look ahead, and rewind if it turns out to be neither.
   */
  if (/^[+-]?\d+$/.test(token)) {
    const afterFirst = this.pos;
    this.skipWhite();
    const secondStart = this.pos;
    const second = this.readToken();
    if (/^\d+$/.test(second)) {
      this.skipWhite();
      const thirdStart = this.pos;
      const third = this.readToken();
      if (third === 'R') {
        return new Ref(parseInt(token, 10), parseInt(second, 10));
      }
      this.pos = thirdStart;
    }
    this.pos = secondStart;
    this.pos = afterFirst;
    return parseInt(token, 10);
  }

  if (/^[+-]?(\d*\.\d*|\d+)$/.test(token)) {
    const value = parseFloat(token);
    return isNaN(value) ? 0 : value;
  }

  // An operator or keyword: hand it back as a marker the caller can react to.
  this.pos = save + token.length;
  return { keyword: token };
};

Lexer.prototype.parseDict = function parseDict(resolveLength) {
  this.pos += 2; // '<<'
  const dict = {};

  for (;;) {
    this.skipWhite();
    if (this.pos >= this.buf.length) break;
    if (this.buf[this.pos] === 0x3e && this.buf[this.pos + 1] === 0x3e) {
      this.pos += 2;
      break;
    }
    if (this.buf[this.pos] !== 0x2f) {
      // Not a key: the dictionary is malformed. Step over and keep going
      // rather than abandoning the whole file.
      this.pos += 1;
      continue;
    }
    const key = this.readName().name;
    const value = this.parseObject(resolveLength);
    dict[key] = value;
  }

  // A stream follows its dictionary.
  const save = this.pos;
  this.skipWhite();
  if (this.buf.toString('latin1', this.pos, this.pos + 6) === 'stream') {
    this.pos += 6;
    // The keyword is followed by CRLF or LF - never by CR alone.
    if (this.buf[this.pos] === 0x0d) this.pos += 1;
    if (this.buf[this.pos] === 0x0a) this.pos += 1;

    let length = dict.Length;
    if (length instanceof Ref && typeof resolveLength === 'function') {
      length = resolveLength(length);
    }
    const start = this.pos;

    if (typeof length !== 'number' || length < 0 || start + length > this.buf.length) {
      /*
       * A broken or indirect /Length that could not be resolved. Rather than
       * give up, find the "endstream" keyword - which is what every tolerant
       * PDF reader does, because broken /Length values are common.
       */
      const index = this.buf.indexOf('endstream', start, 'latin1');
      length = index < 0 ? 0 : index - start;
    }

    this.pos = start + length;
    const after = this.buf.indexOf('endstream', this.pos, 'latin1');
    if (after >= 0) this.pos = after + 9;
    return new PdfStream(dict, start, length);
  }

  this.pos = save;
  return dict;
};

/* -------------------------------------------------------------- the document */

function PdfDocument(buffer) {
  this.buf = Buffer.isBuffer(buffer) ? buffer : Buffer.from(buffer);
  this.objects = new Map();      // "num gen" -> offset
  this.cache = new Map();
  this.compressed = new Map();   // num -> { streamNum, index }
  this.trailer = {};
  this.warnings = [];
}

/*
 * Find every "N G obj" in the file by scanning it.
 *
 * A proper reader follows the cross-reference table from startxref. This
 * scans instead, on purpose: a scanned PDF that has been through a mail
 * gateway, a virus scanner and two SharePoint round trips very often has an
 * xref table whose offsets are wrong by a few bytes, and a reader that trusts
 * it fails on exactly the files that most need reading. Scanning cannot be
 * fooled by a stale offset. It costs one pass over the file.
 */
PdfDocument.prototype.scan = function scan() {
  const buf = this.buf;
  const pattern = Buffer.from('obj', 'latin1');
  let index = buf.indexOf(pattern, 0);

  while (index >= 0) {
    // Walk backwards over "  N G " to find the object number.
    let p = index - 1;
    while (p >= 0 && isWhite(buf[p])) p -= 1;
    const genEnd = p + 1;
    while (p >= 0 && buf[p] >= 0x30 && buf[p] <= 0x39) p -= 1;
    const genStart = p + 1;

    if (genEnd > genStart) {
      while (p >= 0 && isWhite(buf[p])) p -= 1;
      const numEnd = p + 1;
      while (p >= 0 && buf[p] >= 0x30 && buf[p] <= 0x39) p -= 1;
      const numStart = p + 1;

      if (numEnd > numStart && (numStart === 0 || !isRegular(buf[numStart - 1]))) {
        const num = parseInt(buf.toString('latin1', numStart, numEnd), 10);
        const gen = parseInt(buf.toString('latin1', genStart, genEnd), 10);
        if (!isNaN(num) && !isNaN(gen)) {
          // A later definition of the same object wins - that is what an
          // incremental update means.
          this.objects.set(num, index + 3);
        }
      }
    }
    index = buf.indexOf(pattern, index + 3);
  }

  // The trailer, for /Root. Take the LAST one: later updates supersede.
  let trailerIndex = buf.lastIndexOf('trailer', buf.length, 'latin1');
  while (trailerIndex >= 0) {
    const lexer = new Lexer(buf, trailerIndex + 7);
    const dict = lexer.parseObject(this.resolveLength.bind(this));
    if (dict && !(dict instanceof PdfStream) && typeof dict === 'object' && dict.Root) {
      this.trailer = dict;
      break;
    }
    trailerIndex = buf.lastIndexOf('trailer', trailerIndex - 1, 'latin1');
  }

  this.scanObjectStreams();
  return this;
};

PdfDocument.prototype.resolveLength = function resolveLength(ref) {
  const value = this.get(ref);
  return typeof value === 'number' ? value : null;
};

/*
 * Modern PDFs pack most objects into compressed object streams, and a page's
 * dictionary is very often in one. Without this, a perfectly ordinary
 * PDF 1.5+ scan looks like it has no pages at all.
 */
PdfDocument.prototype.scanObjectStreams = function scanObjectStreams() {
  const numbers = Array.from(this.objects.keys());
  for (let i = 0; i < numbers.length; i += 1) {
    let obj;
    try {
      obj = this.getByNumber(numbers[i]);
    } catch (e) {
      continue;
    }
    if (!(obj instanceof PdfStream)) continue;
    const type = obj.dict.Type;
    if (!(type instanceof Name) || type.name !== 'ObjStm') continue;

    let data;
    try {
      data = this.decodeStream(obj);
    } catch (e) {
      this.warnings.push('object stream ' + numbers[i] + ' could not be decompressed');
      continue;
    }
    if (data === null) continue;

    const count = this.get(obj.dict.N);
    const first = this.get(obj.dict.First);
    if (typeof count !== 'number' || typeof first !== 'number') continue;

    // The header is `count` pairs of "objectNumber offset".
    const header = new Lexer(data, 0);
    for (let n = 0; n < count; n += 1) {
      header.skipWhite();
      const objNum = parseInt(header.readToken(), 10);
      header.skipWhite();
      const offset = parseInt(header.readToken(), 10);
      if (isNaN(objNum) || isNaN(offset)) break;
      // Only if the object is not already defined directly in the file.
      if (!this.objects.has(objNum)) {
        this.compressed.set(objNum, { data, offset: first + offset });
      }
    }
  }
};

PdfDocument.prototype.getByNumber = function getByNumber(num) {
  if (this.cache.has(num)) return this.cache.get(num);

  let value = null;
  if (this.objects.has(num)) {
    const lexer = new Lexer(this.buf, this.objects.get(num));
    value = lexer.parseObject(this.resolveLength.bind(this));
  } else if (this.compressed.has(num)) {
    const entry = this.compressed.get(num);
    const lexer = new Lexer(entry.data, entry.offset);
    value = lexer.parseObject(null);
  }

  this.cache.set(num, value);
  return value;
};

/* Follow a reference, however many hops. */
PdfDocument.prototype.get = function get(value) {
  let current = value;
  let hops = 0;
  while (current instanceof Ref) {
    hops += 1;
    // A malformed file can point an object at itself.
    if (hops > 32) return null;
    current = this.getByNumber(current.num);
  }
  return current;
};

/*
 * Undo the stream's filters. Only the ones a scanned PDF's structure uses -
 * the IMAGE filters are deliberately left alone, because for those the
 * compressed bytes are the useful thing.
 */
PdfDocument.prototype.decodeStream = function decodeStream(stream) {
  let data = this.buf.slice(stream.start, stream.start + stream.length);
  if (stream.dict === undefined) return data;

  let filters = this.get(stream.dict.Filter);
  if (!filters) return data;
  if (!Array.isArray(filters)) filters = [filters];

  for (let i = 0; i < filters.length; i += 1) {
    const filter = this.get(filters[i]);
    if (!(filter instanceof Name)) continue;

    if (filter.name === 'FlateDecode' || filter.name === 'Fl') {
      try {
        data = zlib.inflateSync(data);
      } catch (e) {
        // Some producers write a stream with a corrupt final block. Raw
        // inflate recovers the part that is intact, which is usually all of it.
        try {
          data = zlib.inflateRawSync(data.slice(2));
        } catch (e2) {
          throw new Error('FlateDecode failed: ' + e.message);
        }
      }
    } else if (filter.name === 'ASCIIHexDecode' || filter.name === 'AHx') {
      const text = data.toString('latin1').replace(/[^0-9a-fA-F>]/g, '');
      const end = text.indexOf('>');
      const hex = end >= 0 ? text.slice(0, end) : text;
      const bytes = Buffer.alloc(Math.ceil(hex.length / 2));
      for (let n = 0; n < bytes.length; n += 1) {
        bytes[n] = parseInt((hex.substr(n * 2, 2) + '0').slice(0, 2), 16);
      }
      data = bytes;
    } else if (filter.name === 'ASCII85Decode' || filter.name === 'A85') {
      data = decodeAscii85(data);
    } else {
      // An image filter. Stop here and let the caller deal with the bytes.
      break;
    }
  }

  // A predictor is applied after decompression, typically on xref streams.
  const parms = this.get(stream.dict.DecodeParms) || this.get(stream.dict.DP);
  const parm = Array.isArray(parms) ? this.get(parms[0]) : parms;
  if (parm && typeof parm === 'object' && !(parm instanceof Name)) {
    const predictor = this.get(parm.Predictor);
    if (typeof predictor === 'number' && predictor >= 10) {
      data = undoPngPredictor(
        data,
        this.get(parm.Colors) || 1,
        this.get(parm.BitsPerComponent) || 8,
        this.get(parm.Columns) || 1,
      );
    }
  }

  return data;
};

function decodeAscii85(input) {
  const text = input.toString('latin1').replace(/\s/g, '').replace(/^<~/, '');
  const end = text.indexOf('~>');
  const body = end >= 0 ? text.slice(0, end) : text;
  const out = [];
  let i = 0;
  while (i < body.length) {
    if (body[i] === 'z') { out.push(0, 0, 0, 0); i += 1; continue; }
    let group = body.slice(i, i + 5);
    const short = 5 - group.length;
    while (group.length < 5) group += 'u';
    let value = 0;
    for (let n = 0; n < 5; n += 1) value = value * 85 + (group.charCodeAt(n) - 33);
    const bytes = [(value >>> 24) & 0xff, (value >>> 16) & 0xff, (value >>> 8) & 0xff, value & 0xff];
    for (let n = 0; n < 4 - short; n += 1) out.push(bytes[n]);
    i += 5;
  }
  return Buffer.from(out);
}

/* The PNG predictor, as used by xref and some image streams. */
function undoPngPredictor(data, colors, bits, columns) {
  const bpp = Math.max(1, Math.ceil((colors * bits) / 8));
  const rowLength = Math.ceil((colors * bits * columns) / 8);
  const rows = Math.floor(data.length / (rowLength + 1));
  const out = Buffer.alloc(rows * rowLength);

  let prev = Buffer.alloc(rowLength);
  for (let r = 0; r < rows; r += 1) {
    const filter = data[r * (rowLength + 1)];
    const row = data.slice(r * (rowLength + 1) + 1, (r + 1) * (rowLength + 1));
    const decoded = Buffer.alloc(rowLength);

    for (let i = 0; i < rowLength; i += 1) {
      const x = row[i] === undefined ? 0 : row[i];
      const a = i >= bpp ? decoded[i - bpp] : 0;
      const b = prev[i];
      const c = i >= bpp ? prev[i - bpp] : 0;
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
          value = x + ((pa <= pb && pa <= pc) ? a : (pb <= pc ? b : c));
          break;
        }
        default: value = x;
      }
      decoded[i] = value & 0xff;
    }
    decoded.copy(out, r * rowLength);
    prev = decoded;
  }
  return out;
}

/* --------------------------------------------------------------- page tree */

PdfDocument.prototype.getPages = function getPages() {
  const pages = [];
  const root = this.get(this.trailer.Root);

  let pagesNode = null;
  if (root && typeof root === 'object') pagesNode = this.get(root.Pages);

  const seen = new Set();
  const self = this;

  function walk(node, inherited, depth) {
    if (!node || typeof node !== 'object' || depth > 64) return;
    const type = self.get(node.Type);

    // Attributes a page inherits from its parent if it does not set them.
    const carry = {
      Resources: node.Resources !== undefined ? node.Resources : inherited.Resources,
      MediaBox: node.MediaBox !== undefined ? node.MediaBox : inherited.MediaBox,
      Rotate: node.Rotate !== undefined ? node.Rotate : inherited.Rotate,
    };

    if (type instanceof Name && type.name === 'Page') {
      pages.push({ dict: node, inherited: carry });
      return;
    }

    const kids = self.get(node.Kids);
    if (Array.isArray(kids)) {
      for (let i = 0; i < kids.length; i += 1) {
        const key = kids[i] instanceof Ref ? kids[i].num : null;
        if (key !== null) {
          if (seen.has(key)) continue;   // a cyclic page tree
          seen.add(key);
        }
        walk(self.get(kids[i]), carry, depth + 1);
      }
    }
  }

  if (pagesNode) walk(pagesNode, {}, 0);

  /*
   * A file whose catalogue is unreadable still usually has perfectly good
   * page objects in it. Falling back to every object of /Type /Page turns
   * "this PDF is broken" into "here are your pages", which is the whole point.
   */
  if (pages.length === 0) {
    const numbers = Array.from(this.objects.keys()).sort(function (a, b) { return a - b; });
    for (let i = 0; i < numbers.length; i += 1) {
      let obj;
      try { obj = this.getByNumber(numbers[i]); } catch (e) { continue; }
      if (!obj || typeof obj !== 'object' || obj instanceof PdfStream) continue;
      const type = this.get(obj.Type);
      if (type instanceof Name && type.name === 'Page') {
        pages.push({ dict: obj, inherited: {} });
      }
    }
    if (pages.length > 0) {
      this.warnings.push('the page tree was unreadable; pages were found by scanning');
    }
  }

  return pages;
};

/* ------------------------------------------------------------ page images */

function nameOf(value) {
  return value instanceof Name ? value.name : null;
}

/*
 * The image XObjects a page draws, in the order they are declared.
 */
PdfDocument.prototype.getPageImages = function getPageImages(page) {
  const resources = this.get(page.dict.Resources !== undefined
    ? page.dict.Resources : page.inherited.Resources);
  if (!resources || typeof resources !== 'object') return [];

  const xobjects = this.get(resources.XObject);
  if (!xobjects || typeof xobjects !== 'object') return [];

  const found = [];
  const keys = Object.keys(xobjects);
  for (let i = 0; i < keys.length; i += 1) {
    const obj = this.get(xobjects[keys[i]]);
    if (!(obj instanceof PdfStream)) continue;
    const subtype = nameOf(this.get(obj.dict.Subtype));
    if (subtype !== 'Image') continue;
    found.push({ name: keys[i], stream: obj });
  }
  return found;
};

/* Which filter is actually on the image data, after any transport filters. */
PdfDocument.prototype.imageFilter = function imageFilter(stream) {
  let filters = this.get(stream.dict.Filter);
  if (!filters) return null;
  if (!Array.isArray(filters)) filters = [filters];
  for (let i = 0; i < filters.length; i += 1) {
    const name = nameOf(this.get(filters[i]));
    if (name && name !== 'FlateDecode' && name !== 'Fl'
      && name !== 'ASCIIHexDecode' && name !== 'AHx'
      && name !== 'ASCII85Decode' && name !== 'A85') {
      return name;
    }
  }
  return null;
};

const UNSUPPORTED = {
  CCITTFaxDecode: 'CCITT Group 3/4 fax compression',
  JBIG2Decode: 'JBIG2 compression',
  JPXDecode: 'JPEG 2000 compression',
};

/*
 * Turn one image XObject into a file that Tesseract.js can read.
 *
 * Returns { ok, format, buffer, width, height, reason }.
 */
PdfDocument.prototype.extractImage = function extractImage(stream) {
  const dict = stream.dict;
  const width = this.get(dict.Width) || this.get(dict.W);
  const height = this.get(dict.Height) || this.get(dict.H);
  const bpc = this.get(dict.BitsPerComponent) || this.get(dict.BPC) || 8;
  const filter = this.imageFilter(stream);

  if (typeof width !== 'number' || typeof height !== 'number' || width < 1 || height < 1) {
    return { ok: false, reason: 'the image has no usable dimensions' };
  }

  /*
   * A JPEG inside a PDF is a complete JPEG file. Writing the bytes out
   * unchanged is not just the fast path - it is the LOSSLESS one. Decoding
   * and re-encoding would throw away detail the OCR engine wants.
   */
  if (filter === 'DCTDecode' || filter === 'DCT') {
    let data = this.buf.slice(stream.start, stream.start + stream.length);
    // Unwrap any transport filter applied on top of the JPEG.
    let filters = this.get(dict.Filter);
    if (!Array.isArray(filters)) filters = [filters];
    if (filters.length > 1) {
      try {
        data = this.decodeStream(stream);
      } catch (e) {
        return { ok: false, reason: 'the JPEG could not be unwrapped: ' + e.message };
      }
    }
    return { ok: true, format: 'jpeg', buffer: data, width, height };
  }

  if (UNSUPPORTED[filter]) {
    return {
      ok: false,
      format: filter,
      width,
      height,
      reason: UNSUPPORTED[filter] + ' is not decoded here. Render this PDF with the '
        + 'Windows renderer (Invoke-Ocr handles it) or export the pages as PNG.',
    };
  }

  // Everything else: raw samples, which have to be turned into a PNG.
  let raw;
  try {
    raw = this.decodeStream(stream);
  } catch (e) {
    return { ok: false, reason: 'the image data could not be decompressed: ' + e.message };
  }

  const space = this.colourSpace(dict);
  if (space === null) {
    return { ok: false, reason: 'unsupported colour space' };
  }

  const decode = this.get(dict.Decode) || this.get(dict.D);
  const inverted = Array.isArray(decode) && this.get(decode[0]) === 1;

  try {
    const image = samplesToImage(raw, width, height, bpc, space, inverted);
    return { ok: true, format: 'png', buffer: encodePng(image), width, height };
  } catch (e) {
    return { ok: false, reason: 'the samples could not be laid out: ' + e.message };
  }
};

/*
 * Work out how many components a pixel has, and the palette if there is one.
 */
PdfDocument.prototype.colourSpace = function colourSpace(dict) {
  let cs = this.get(dict.ColorSpace !== undefined ? dict.ColorSpace : dict.CS);

  // An ImageMask is one bit per pixel, ink or nothing - very common for the
  // black-and-white scans a document feeder produces.
  if (this.get(dict.ImageMask) === true || this.get(dict.IM) === true) {
    return { components: 1, kind: 'gray', mask: true };
  }

  if (cs instanceof Name) {
    const name = cs.name;
    if (name === 'DeviceGray' || name === 'G' || name === 'CalGray') return { components: 1, kind: 'gray' };
    if (name === 'DeviceRGB' || name === 'RGB' || name === 'CalRGB') return { components: 3, kind: 'rgb' };
    if (name === 'DeviceCMYK' || name === 'CMYK') return { components: 4, kind: 'cmyk' };
    return null;
  }

  if (Array.isArray(cs)) {
    const family = nameOf(this.get(cs[0]));

    if (family === 'Indexed' || family === 'I') {
      const base = this.get(cs[1]);
      const baseName = nameOf(base);
      let baseComponents = 3;
      if (baseName === 'DeviceGray' || baseName === 'G') baseComponents = 1;
      else if (baseName === 'DeviceCMYK') baseComponents = 4;

      let lookup = this.get(cs[3]);
      if (lookup instanceof PdfStream) {
        try { lookup = this.decodeStream(lookup); } catch (e) { lookup = null; }
      } else if (typeof lookup === 'string') {
        lookup = Buffer.from(lookup, 'latin1');
      }
      if (!lookup) return null;
      return { components: 1, kind: 'indexed', palette: lookup, baseComponents };
    }

    if (family === 'ICCBased') {
      const streamObj = this.get(cs[1]);
      const n = streamObj && streamObj.dict ? this.get(streamObj.dict.N) : 3;
      if (n === 1) return { components: 1, kind: 'gray' };
      if (n === 4) return { components: 4, kind: 'cmyk' };
      return { components: 3, kind: 'rgb' };
    }

    if (family === 'DeviceN' || family === 'Separation') {
      // One ink. Treated as greyscale, which for a scanned page is right.
      return { components: 1, kind: 'gray', subtractive: true };
    }
  }

  // No colour space at all: assume greyscale, which is what a bare scan is.
  return { components: 1, kind: 'gray' };
};

/*
 * Unpack the sample data into 8-bit grey or RGB.
 */
function samplesToImage(raw, width, height, bpc, space, inverted) {
  const components = space.components;
  const rowBits = width * components * bpc;
  const rowBytes = Math.ceil(rowBits / 8);

  if (raw.length < rowBytes * height) {
    // A truncated image is still worth reading as far as it goes; the missing
    // part comes out white rather than the whole page being lost.
    const padded = Buffer.alloc(rowBytes * height, 0xff);
    raw.copy(padded, 0, 0, Math.min(raw.length, padded.length));
    raw = padded;
  }

  const maxValue = (1 << bpc) - 1;

  function sample(row, index) {
    if (bpc === 8) return raw[row * rowBytes + index];
    if (bpc === 16) return raw[row * rowBytes + index * 2];
    const bitPos = index * bpc;
    const byte = raw[row * rowBytes + (bitPos >> 3)];
    const shift = 8 - bpc - (bitPos & 7);
    return (byte >> shift) & maxValue;
  }

  if (space.kind === 'indexed') {
    const out = new Uint8Array(width * height * 3);
    const base = space.baseComponents;
    for (let y = 0; y < height; y += 1) {
      for (let x = 0; x < width; x += 1) {
        const idx = sample(y, x);
        for (let c = 0; c < 3; c += 1) {
          const at = idx * base + (base === 1 ? 0 : c);
          const v = at < space.palette.length ? space.palette[at] : 0;
          out[(y * width + x) * 3 + c] = v;
        }
      }
    }
    return { width, height, channels: 3, data: out };
  }

  if (space.kind === 'cmyk') {
    const out = new Uint8Array(width * height * 3);
    for (let y = 0; y < height; y += 1) {
      for (let x = 0; x < width; x += 1) {
        const scale = 255 / maxValue;
        const c = sample(y, x * 4) * scale;
        const m = sample(y, x * 4 + 1) * scale;
        const yl = sample(y, x * 4 + 2) * scale;
        const k = sample(y, x * 4 + 3) * scale;
        out[(y * width + x) * 3] = Math.max(0, 255 - Math.min(255, c + k));
        out[(y * width + x) * 3 + 1] = Math.max(0, 255 - Math.min(255, m + k));
        out[(y * width + x) * 3 + 2] = Math.max(0, 255 - Math.min(255, yl + k));
      }
    }
    return { width, height, channels: 3, data: out };
  }

  if (space.kind === 'rgb') {
    const out = new Uint8Array(width * height * 3);
    const scale = 255 / maxValue;
    for (let y = 0; y < height; y += 1) {
      for (let x = 0; x < width; x += 1) {
        for (let c = 0; c < 3; c += 1) {
          out[(y * width + x) * 3 + c] = Math.round(sample(y, x * 3 + c) * scale);
        }
      }
    }
    return { width, height, channels: 3, data: out };
  }

  // Greyscale, including the 1-bit image mask case.
  const out = new Uint8Array(width * height);
  const scale = 255 / maxValue;
  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      let v = Math.round(sample(y, x) * scale);
      /*
       * An ImageMask stores 1 where the ink goes, which is the opposite of a
       * greyscale sample. Getting this backwards produces a page of white
       * text on black - which OCR reads as an empty page.
       */
      if (space.mask || space.subtractive) v = 255 - v;
      if (inverted) v = 255 - v;
      out[y * width + x] = v;
    }
  }
  return { width, height, channels: 1, data: out };
}

/* --------------------------------------------------------------- the API */

/*
 * The whole job: a scanned PDF's bytes in, one image per page out.
 *
 * Returns { ok, pages: [ { number, ok, format, buffer, width, height,
 * widthPt, heightPt, dpi, reason } ], warnings }.
 */
function extractPageImages(buffer, options) {
  const opts = options || {};
  const doc = new PdfDocument(buffer).scan();
  const pages = doc.getPages();
  const out = [];

  for (let i = 0; i < pages.length; i += 1) {
    const page = pages[i];
    const number = i + 1;

    // MediaBox is [x0 y0 x1 y1] in points - the physical page size, which is
    // what turns pixels into a resolution.
    let box = doc.get(page.dict.MediaBox !== undefined ? page.dict.MediaBox : page.inherited.MediaBox);
    let widthPt = 612;
    let heightPt = 792;
    if (Array.isArray(box) && box.length === 4) {
      const x0 = doc.get(box[0]);
      const y0 = doc.get(box[1]);
      const x1 = doc.get(box[2]);
      const y1 = doc.get(box[3]);
      if ([x0, y0, x1, y1].every(function (v) { return typeof v === 'number'; })) {
        widthPt = Math.abs(x1 - x0);
        heightPt = Math.abs(y1 - y0);
      }
    }

    const images = doc.getPageImages(page);

    if (images.length === 0) {
      out.push({
        number, ok: false, widthPt, heightPt,
        reason: 'the page holds no image - it is probably a text PDF, which needs no OCR',
      });
      continue;
    }

    /*
     * A scan is one image covering the page. When a page carries several,
     * the biggest is the scan and the rest are a logo or a signature strip,
     * so the biggest is the one to read.
     */
    let best = null;
    for (let n = 0; n < images.length; n += 1) {
      const w = doc.get(images[n].stream.dict.Width) || 0;
      const h = doc.get(images[n].stream.dict.Height) || 0;
      if (best === null || w * h > best.area) {
        best = { entry: images[n], area: w * h };
      }
    }

    const extracted = doc.extractImage(best.entry.stream);
    if (!extracted.ok) {
      out.push({
        number, ok: false, widthPt, heightPt,
        format: extracted.format, reason: extracted.reason,
      });
      continue;
    }

    // The resolution the page was scanned at, which the preprocessing stage
    // needs in order to decide how much to enlarge it.
    const dpi = widthPt > 0 ? Math.round((extracted.width * 72) / widthPt) : 0;

    out.push({
      number,
      ok: true,
      format: extracted.format,
      buffer: extracted.buffer,
      width: extracted.width,
      height: extracted.height,
      widthPt,
      heightPt,
      dpi,
      images: images.length,
    });
  }

  return {
    ok: out.some(function (p) { return p.ok; }),
    pages: out,
    warnings: doc.warnings,
  };
}

module.exports = {
  extractPageImages,
  PdfDocument,
  Lexer,
  Ref,
  Name,
  PdfStream,
  decodeAscii85,
  undoPngPredictor,
  samplesToImage,
};
