'use strict';

/*
 * xml.js - a small, strict XML writer and reader.
 *
 * The OCR result is written as XML rather than plain text on purpose. Flat text
 * throws away the one thing OCR knows and a human reader takes for granted:
 * WHERE on the page each word sat. Once "4987   36.07   99,914.15" is a single
 * line of text, nothing downstream can tell which column the 36.07 came from.
 * The XML keeps every word's box, so the column is recoverable.
 *
 * No dependencies. Node's built-ins only.
 */

/* ------------------------------------------------------------------ escaping */

/*
 * The five XML predefined entities. Attribute values additionally must not
 * contain a raw newline or tab, because an XML parser is allowed to normalise
 * those to a space - which would silently corrupt a value.
 */
function escapeText(s) {
  if (s === null || s === undefined) return '';
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function escapeAttr(s) {
  if (s === null || s === undefined) return '';
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;')
    .replace(/\r/g, '&#13;')
    .replace(/\n/g, '&#10;')
    .replace(/\t/g, '&#9;');
}

/*
 * XML 1.0 forbids most control characters outright - there is no escape for
 * them. OCR of a noisy scan can produce them, so they are stripped rather than
 * written, which would make the whole document unparseable.
 */
function stripInvalid(s) {
  if (s === null || s === undefined) return '';
  var out = '';
  var str = String(s);
  for (var i = 0; i < str.length; i += 1) {
    var c = str.charCodeAt(i);
    var ok = c === 0x09 || c === 0x0a || c === 0x0d
      || (c >= 0x20 && c <= 0xd7ff)
      || (c >= 0xe000 && c <= 0xfffd)
      || (c >= 0x10000 && c <= 0x10ffff);
    if (ok) out += str.charAt(i);
  }
  return out;
}

/* ------------------------------------------------------------------- writing */

/*
 * A writer rather than a string-concatenating helper, so that indentation and
 * tag balance are the writer's problem and never the caller's. An unclosed tag
 * throws at end() instead of producing a file that only fails much later, in
 * PowerShell, on a user's machine.
 */
function XmlWriter(options) {
  var opts = options || {};
  this.indentUnit = opts.indent === undefined ? '  ' : opts.indent;
  this.parts = [];
  this.stack = [];
  this.ended = false;
  this.parts.push('<?xml version="1.0" encoding="utf-8"?>');
}

XmlWriter.prototype._indent = function _indent() {
  if (this.indentUnit === '') return '';
  var s = '';
  for (var i = 0; i < this.stack.length; i += 1) s += this.indentUnit;
  return s;
};

XmlWriter.prototype._attrs = function _attrs(attrs) {
  if (!attrs) return '';
  var s = '';
  var keys = Object.keys(attrs);
  for (var i = 0; i < keys.length; i += 1) {
    var v = attrs[keys[i]];
    // An attribute set to null/undefined is omitted rather than written empty:
    // "absent" and "present but blank" are different facts about a page.
    if (v === null || v === undefined) continue;
    s += ' ' + keys[i] + '="' + escapeAttr(stripInvalid(v)) + '"';
  }
  return s;
};

XmlWriter.prototype.open = function open(name, attrs) {
  if (this.ended) throw new Error('XmlWriter: document already ended');
  this.parts.push(this._indent() + '<' + name + this._attrs(attrs) + '>');
  this.stack.push(name);
  return this;
};

XmlWriter.prototype.close = function close(name) {
  if (this.stack.length === 0) throw new Error('XmlWriter: close("' + name + '") with nothing open');
  var top = this.stack.pop();
  if (name !== undefined && name !== top) {
    throw new Error('XmlWriter: close("' + name + '") but <' + top + '> is open');
  }
  this.parts.push(this._indent() + '</' + top + '>');
  return this;
};

/* An element with no children: <Word .../> */
XmlWriter.prototype.empty = function empty(name, attrs) {
  if (this.ended) throw new Error('XmlWriter: document already ended');
  this.parts.push(this._indent() + '<' + name + this._attrs(attrs) + ' />');
  return this;
};

/* An element with text content on one line: <Text>hello</Text> */
XmlWriter.prototype.leaf = function leaf(name, attrs, text) {
  if (this.ended) throw new Error('XmlWriter: document already ended');
  var body = escapeText(stripInvalid(text));
  this.parts.push(this._indent() + '<' + name + this._attrs(attrs) + '>' + body + '</' + name + '>');
  return this;
};

XmlWriter.prototype.comment = function comment(text) {
  // "--" cannot appear inside an XML comment at all; there is no escape for it.
  var body = String(stripInvalid(text)).replace(/-{2,}/g, '-');
  this.parts.push(this._indent() + '<!-- ' + body + ' -->');
  return this;
};

XmlWriter.prototype.end = function end() {
  if (this.stack.length !== 0) {
    throw new Error('XmlWriter: document ended with <' + this.stack.join('>, <') + '> still open');
  }
  this.ended = true;
  return this.parts.join('\n') + '\n';
};

/* ------------------------------------------------------------------- reading */

/*
 * A deliberately small reader: enough to read back a document this writer
 * produced (elements, attributes, text) and no more. It is here so the JS test
 * suite can verify a round trip without pulling in a parser dependency. It is
 * NOT a general XML parser - it does not do namespaces, DTDs or entities beyond
 * the five predefined ones plus numeric character references.
 */
function unescapeXml(s) {
  return String(s).replace(/&(#x?[0-9a-fA-F]+|amp|lt|gt|quot|apos);/g, function (m, ent) {
    if (ent === 'amp') return '&';
    if (ent === 'lt') return '<';
    if (ent === 'gt') return '>';
    if (ent === 'quot') return '"';
    if (ent === 'apos') return "'";
    if (ent.charAt(0) === '#') {
      var code = ent.charAt(1) === 'x' || ent.charAt(1) === 'X'
        ? parseInt(ent.slice(2), 16)
        : parseInt(ent.slice(1), 10);
      if (isNaN(code)) return m;
      return String.fromCodePoint(code);
    }
    return m;
  });
}

function parseAttrs(src) {
  var attrs = {};
  var re = /([A-Za-z_][-A-Za-z0-9_.:]*)\s*=\s*("([^"]*)"|'([^']*)')/g;
  var m = re.exec(src);
  while (m !== null) {
    attrs[m[1]] = unescapeXml(m[3] !== undefined ? m[3] : m[4]);
    m = re.exec(src);
  }
  return attrs;
}

/*
 * Returns { name, attrs, children, text }. Text is the concatenation of the
 * element's own character data, trimmed.
 */
function parseXml(xml) {
  var src = String(xml);
  var root = null;
  var stack = [];
  var re = /<!--[\s\S]*?-->|<\?[\s\S]*?\?>|<!\[CDATA\[([\s\S]*?)\]\]>|<\/\s*([A-Za-z_][-A-Za-z0-9_.:]*)\s*>|<\s*([A-Za-z_][-A-Za-z0-9_.:]*)((?:[^>"']|"[^"]*"|'[^']*')*?)(\/?)>/g;
  var last = 0;
  var m = re.exec(src);

  function addText(t) {
    if (stack.length === 0) return;
    var s = unescapeXml(t);
    if (s.trim() !== '' || /[^\s]/.test(s)) stack[stack.length - 1].text += s;
    else stack[stack.length - 1].text += '';
  }

  while (m !== null) {
    if (m.index > last) addText(src.slice(last, m.index));
    last = re.lastIndex;

    if (m[1] !== undefined) {                       // CDATA
      if (stack.length > 0) stack[stack.length - 1].text += m[1];
    } else if (m[2] !== undefined) {                // closing tag
      if (stack.length === 0) throw new Error('parseXml: unexpected </' + m[2] + '>');
      var done = stack.pop();
      if (done.name !== m[2]) throw new Error('parseXml: </' + m[2] + '> closes <' + done.name + '>');
      done.text = done.text.trim();
      if (stack.length === 0) root = done;
    } else if (m[3] !== undefined) {                // opening or self-closing
      var node = { name: m[3], attrs: parseAttrs(m[4] || ''), children: [], text: '' };
      if (stack.length > 0) stack[stack.length - 1].children.push(node);
      if (m[5] === '/') {
        if (stack.length === 0) root = node;
      } else {
        stack.push(node);
      }
    }
    m = re.exec(src);
  }

  if (stack.length !== 0) throw new Error('parseXml: unclosed <' + stack[stack.length - 1].name + '>');
  if (root === null) throw new Error('parseXml: no root element');
  return root;
}

/* Convenience: every descendant with the given tag name, document order. */
function findAll(node, name) {
  var out = [];
  (function walk(n) {
    for (var i = 0; i < n.children.length; i += 1) {
      if (n.children[i].name === name) out.push(n.children[i]);
      walk(n.children[i]);
    }
  }(node));
  return out;
}

module.exports = {
  XmlWriter: XmlWriter,
  escapeText: escapeText,
  escapeAttr: escapeAttr,
  stripInvalid: stripInvalid,
  unescapeXml: unescapeXml,
  parseXml: parseXml,
  findAll: findAll,
};
