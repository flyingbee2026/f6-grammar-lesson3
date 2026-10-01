# -*- coding: utf-8 -*-
"""Dump every body element (paragraphs + tables) with run-level formatting."""
import sys
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.oxml.ns import qn

path = sys.argv[1]
d = Document(path)

def fmt(run):
    c = run.font.color
    col = None
    if c is not None and c.type is not None:
        col = str(c.rgb) if c.type == 1 else 'theme'
    return "[%s%s%s%s%s]" % (
        run.font.name or '',
        '' if run.font.size is None else '/' + str(run.font.size.pt),
        '/B' if run.font.bold else '',
        '/I' if run.font.italic else '',
        ('/' + col) if col else '')

def pshd(par):
    pPr = par._p.pPr
    if pPr is None: return ''
    shd = pPr.find(qn('w:shd'))
    if shd is None: return ''
    return ' shd=%s' % shd.get(qn('w:fill'))

def borders(par):
    pPr = par._p.pPr
    if pPr is None: return ''
    pBdr = pPr.find(qn('w:pBdr'))
    if pBdr is None: return ''
    return ' pBdr=' + ','.join(e.tag.split('}')[1] for e in pBdr)

for child in d.element.body.iterchildren():
    tag = child.tag.split('}')[1]
    if tag == 'p':
        par = Paragraph(child, d)
        runs = ' + '.join('%s%s' % (fmt(r), r.text) for r in par.runs if r.text)
        print('P[%s|%s%s%s] %s' % (par.style.name, par.alignment, pshd(par), borders(par), runs))
    elif tag == 'tbl':
        t = Table(child, d)
        print('TBL %dx%d style=%s' % (len(t.rows), len(t.columns), t.style.name))
        for ri, r in enumerate(t.rows):
            seen = set()
            out = []
            for c in r.cells:
                if id(c) in seen:
                    out.append('<<merged>>'); continue
                seen.add(id(c))
                inner = ' / '.join(
                    ' '.join('%s%s' % (fmt(run), run.text) for run in p.runs if run.text) for p in c.paragraphs)
                w = c.width
                out.append('{%s} %s' % ('' if w is None else round(w.mm, 1), inner))
            print('  r%d | %s' % (ri, ' || '.join(out)))
