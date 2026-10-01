# -*- coding: utf-8 -*-
"""Dump a docx body in document order (paragraphs + tables) for proofreading."""
import sys
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph

path = sys.argv[1]
d = Document(path)
for child in d.element.body.iterchildren():
    tag = child.tag.split('}')[1]
    if tag == 'p':
        txt = Paragraph(child, d).text
        print('P| ' + txt)
    elif tag == 'tbl':
        t = Table(child, d)
        print('T| --- table %dx%d ---' % (len(t.rows), len(t.columns)))
        for r in t.rows:
            cells = []
            for c in r.cells:
                inner = ' / '.join(p.text.strip() for p in c.paragraphs if p.text.strip())
                cells.append(inner)
            seen, uniq = set(), []
            for c in cells:
                if id(c) not in seen:
                    seen.add(id(c)); uniq.append(c)
            print('   | ' + ' || '.join(uniq))
