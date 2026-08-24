Sample pages used by the test suite.

page_clean.png
    The standard Tesseract OCR test page, converted to 8-bit greyscale and
    re-encoded. Taken from the Tesseract.js test assets
    (https://github.com/naptha/tesseract.js, Apache-2.0), which in turn take
    it from the Tesseract project's own test suite.

page_scan.png
    The same page put through everything a bad scan does to it: shrunk to 45%
    (about 96 dpi), lit unevenly from one side, speckled with noise and tilted
    by 1.8 degrees. Generated deterministically from page_clean.png by a fixed
    pseudo-random sequence, so it is byte-identical on every machine.

    This is the page the preprocessing regression test uses: recognising it
    untouched and then recognising it prepared must show the prepared run
    reading more of the text, or the preparation stage is not earning its
    place in the pipeline.
