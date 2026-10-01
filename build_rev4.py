# -*- coding: utf-8 -*-
"""Rev 4 (from Helen's revised docx):
   teacher  -> her file + the Part A answer key filled into her empty key table
   students -> her file minus every teacher-only item, with answer space added
Usage: python3 build_rev4.py teacher|students
"""
import sys
from docx import Document
from docx.shared import Pt, Mm, RGBColor
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

RED = RGBColor(0xC0, 0x00, 0x00)
SRC = '/root/f6-grammar-lesson3/rev4/helen_teacher.docx'
MODE = sys.argv[1] if len(sys.argv) > 1 else 'teacher'
OUT = ('/root/f6-grammar-lesson3/6B Grammar Lesson 3 - Discussing Solutions (Teacher\'s).docx'
       if MODE == 'teacher' else
       '/root/f6-grammar-lesson3/6B Grammar Lesson 3 - Discussing Solutions (Students).docx')


def shade(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear'); shd.set(qn('w:color'), 'auto'); shd.set(qn('w:fill'), fill)
    tcPr.append(shd)


def set_runs(cell_or_par, runs, first=True):
    """runs = list of (text, bold, colour)"""
    if hasattr(cell_or_par, 'paragraphs'):
        par = cell_or_par.paragraphs[0] if first else cell_or_par.add_paragraph()
    else:
        par = cell_or_par
    for text, bold, colour in runs:
        r = par.add_run(text)
        r.font.bold = bold
        if colour is not None:
            r.font.color.rgb = colour
    return par


def drop(el):
    e = getattr(el, '_element', el)
    e.getparent().remove(e)


def is_red_par(par, needle=None):
    txt = par.text
    if needle and needle not in txt:
        return False
    return any(r.font.color is not None and r.font.color.type is not None
               and str(r.font.color.rgb) == 'C00000' for r in par.runs)


def make_table(rows, cols, widths_mm):
    t = doc.add_table(rows=rows, cols=cols)
    t.style = 'Table Grid'
    t.autofit = False
    t.alignment = WD_TABLE_ALIGNMENT.LEFT
    layout = OxmlElement('w:tblLayout'); layout.set(qn('w:type'), 'fixed')
    t._tbl.tblPr.append(layout)
    for row in t.rows:
        for i, c in enumerate(row.cells):
            c.width = Mm(widths_mm[i])
    return t


doc = Document(SRC)

# ----------------------------------------------------------------- TEACHER
KEYS = [
    ("1",
     "In recent years, food delivery apps have become an indispensable part of daily life for busy "
     "office workers.",
     ["(a) Agreement: \u201capps has\u201d \u2192 \u201capps have\u201d (plural subject).",
      "(b) Verb form: \u201chas became\u201d \u2192 \u201chave become\u201d \u2014 after \u201chave\u201d "
      "the past participle is needed. (Lesson 1, Pattern 1)"]),
    ("2",
     "Living in a fast-paced society, commuters are constantly exposed to noise pollution.",
     ["(a) Part of speech: \u201cfast-pace\u201d \u2192 \u201cfast-paced\u201d (a compound adjective "
      "modifies \u201csociety\u201d)",
      "(b) Agreement: \u201ccommuters is\u201d \u2192 \u201ccommuters are\u201d. (Lesson 1, Pattern 2)"]),
    ("3",
     "In recent times, the issue of the ageing population has sparked widespread concern across "
     "society.",
     ["(a) Plural: \u201cIn recent time\u201d \u2192 \u201cIn recent times\u201d (the fixed pattern "
      "uses the plural).",
      "(b) Verb form: \u201chas spark\u201d \u2192 \u201chas sparked\u201d \u2014 the base form after "
      "\u201chas\u201d leaves the clause without a verb.",
      "(c) Also: \u201cageing population\u201d \u2192 \u201cthe ageing population\u201d (article). "
      "(Lesson 1, Pattern 3)"]),
    ("4",
     "It is undeniable that early childhood education plays a pivotal role in shaping children\u2019s "
     "characters.",
     ["(a) Agreement: \u201ceducation play\u201d \u2192 \u201ceducation plays\u201d.",
      "(b) Verb form: \u201cin shape\u201d \u2192 \u201cin shaping\u201d \u2014 a gerund is needed "
      "after the preposition \u201cin\u201d. (Lesson 1, Pattern 4)"]),
    ("5",
     "To begin with, cultural exchange programmes offer students a golden opportunity to broaden "
     "their horizons.",
     ["(a) Agreement: \u201cprogramme offer\u201d \u2192 \u201cprogrammes offer\u201d (also accepted: "
      "\u201ca cultural exchange programme offers\u201d).",
      "(b) Article: \u201cstudents golden opportunity\u201d \u2192 \u201cstudents a golden "
      "opportunity\u201d.",
      "(c) Also: \u201ctheir horizon\u201d \u2192 \u201ctheir horizons\u201d. (Lesson 2, Pattern 1)"]),
    ("6",
     "Furthermore, joining the school choir offers teenagers a valuable chance to make new friends, "
     "which in turn boosts their confidence.",
     ["(a) Agreement: \u201cchoir offer\u201d \u2192 \u201cchoir offers\u201d (the subject is the "
      "gerund phrase \u201cjoining the school choir\u201d).",
      "(b) Word choice: \u201ca valuable change\u201d \u2192 \u201ca valuable chance\u201d.",
      "(c) Also: \u201cwhich in turn boost\u201d \u2192 \u201cboosts\u201d (the clause refers to the "
      "whole idea). (Lesson 2, Patterns 1 and 2)"]),
    ("7",
     "From the perspective of parents, school counselling services allow them to understand their "
     "children better, which in turn strengthens family relationships.",
     ["(a) Pronoun reference: \u201callow it to understand\u201d \u2192 \u201callow them to "
      "understand\u201d (the parents).",
      "(b) Verb form: \u201cwhich in turn strengthen\u201d \u2192 \u201cstrengthens\u201d \u2014 the "
      "clause needs a finite verb. (Lesson 2, Pattern 2)"]),
    ("8",
     "School libraries serve as a powerful means of encouraging students to read, which enables them "
     "to become lifelong learners.",
     ["(a) Word form: \u201ca powerful mean\u201d \u2192 \u201ca powerful means\u201d (the fixed "
      "pattern).",
      "(b) Verb form: \u201cof encourage\u201d \u2192 \u201cof encouraging\u201d \u2014 a gerund is "
      "needed after the preposition \u201cof\u201d. (Lesson 2, Pattern 3)"]),
    ("9",
     "Admittedly, e-books are more convenient; nevertheless, printed books remain irreplaceable, as "
     "reading on paper helps readers to concentrate better.",
     ["(a) Part of speech: \u201cmore convenience\u201d \u2192 \u201cmore convenient\u201d (an "
      "adjective is needed after \u201care\u201d).",
      "(b) Agreement: \u201creading \u2026 help\u201d \u2192 \u201chelps\u201d (the subject is the "
      "singular gerund \u201creading\u201d).",
      "Note: the semicolon before \u201cnevertheless\u201d is correct \u2014 a comma would be a "
      "comma splice. (Lesson 2, Pattern 4)"]),
    ("10",
     "It is undeniable that a balanced diet plays a pivotal role in maintaining students\u2019 health.",
     ["(a) Agreement: \u201cdiet play\u201d \u2192 \u201cdiet plays\u201d.",
      "(b) Possessive / plural: \u201cstudent\u2019s health\u201d \u2192 \u201cstudents\u2019 "
      "health\u201d (all the students). (Lesson 1, Pattern 4)"]),
]

if MODE == 'teacher':
    key_tbl = None
    for t in doc.tables:
        if len(t.columns) >= 3 and 'The TWO errors' in t.rows[0].cells[2].text:
            key_tbl = t
            break
    assert key_tbl is not None, 'key table not found'
    for num, corrected, errs in KEYS:
        row = key_tbl.add_row()
        c0, c1, c2 = row.cells
        c0.width, c1.width, c2.width = Mm(10), Mm(78), Mm(82)
        set_runs(c0, [(num, True, RED)])
        set_runs(c1, [(corrected, False, RED)])
        for i, e in enumerate(errs):
            set_runs(c2, [(e, False, RED)], first=(i == 0))
    doc.save(OUT)
    print('teacher saved ->', OUT)
    d = Document(OUT)
    for t in d.tables:
        if len(t.columns) >= 3 and 'The TWO errors' in t.rows[0].cells[2].text:
            print('key table rows:', len(t.rows))

# ----------------------------------------------------------------- STUDENTS
else:
    key_tbl = None
    for t in doc.tables:
        if len(t.columns) >= 3 and 'The TWO errors' in t.rows[0].cells[2].text:
            key_tbl = t
            break
    assert key_tbl is not None

    # 1. the red "TEACHER'S VERSION" banner
    for p in list(doc.paragraphs):
        if p.text.strip().startswith('TEACHER\u2019S VERSION'):
            drop(p)

    # 2. Part A: heading -> student instruction; key table -> blank answer table
    heading = None
    for p in doc.paragraphs:
        if p.text.strip().startswith('TEACHER\u2019S ANSWERS'):
            heading = p
    assert heading is not None
    for r in list(heading.runs)[1:]:
        drop(r)
    heading.runs[0].text = 'Write the corrected sentences below:'
    answer_tbl = make_table(11, 2, (10, 160))
    set_runs(answer_tbl.rows[0].cells[0], [('Item', True, None)])
    set_runs(answer_tbl.rows[0].cells[1], [('Corrected sentence', True, None)])
    shade(answer_tbl.rows[0].cells[0], 'D9D9D9')
    shade(answer_tbl.rows[0].cells[1], 'D9D9D9')
    for i in range(10):
        row = answer_tbl.rows[i + 1]
        set_runs(row.cells[0], [('%d.' % (i + 1), True, None)])
        row.cells[1].paragraphs[0].add_run('')
        row.cells[1].add_paragraph()
    key_tbl._tbl.addprevious(answer_tbl._tbl)
    drop(key_tbl)

    # 3. Part B: drop the red teacher notes
    for p in list(doc.paragraphs):
        if is_red_par(p, 'Teaching note:') or is_red_par(p, 'Teacher\u2019s note on the design:'):
            drop(p)

    # 4. Part C: blank the answer cell of every set + drop the red key notes
    for t in doc.tables:
        txt = ' '.join(c.text for r in t.rows for c in r.cells)
        if t.style.name == 'Table Grid' and ('Use Pattern' in txt) and len(t.columns) == 2:
            last = t.rows[-1]
            cell = last.cells[1]
            for p in list(cell.paragraphs)[1:]:
                drop(p)
            for r in list(cell.paragraphs[0].runs):
                drop(r)
            cell.add_paragraph()
            cell.add_paragraph()
    for p in list(doc.paragraphs):
        if is_red_par(p, 'Key note:'):
            drop(p)

    # 5. Part D: model answer -> writing lines; drop the red marking sections
    model_label = None
    for p in doc.paragraphs:
        if p.text.strip().startswith('Teacher\u2019s model answer'):
            model_label = p
    assert model_label is not None
    anchor = model_label._p.getprevious()
    drop(model_label)
    for p in list(doc.paragraphs):
        if p._p.getparent() is None:
            continue
        txt = p.text.strip()
        if any(txt.startswith(x) for x in
               ('Although an extra period', 'To keep us physically fit',
                'A single game of badminton', 'It is also high time that the school looked',
                'I hope the Principal will reconsider')):
            drop(p)
    tail = None
    for p in list(doc.paragraphs):
        if p.text.strip().startswith('What to reward when marking'):
            tail = p
        if p.text.strip().startswith('Common errors to watch when marking'):
            tail = p
    # delete from 'What to reward...' to the end of the body
    started = False
    for p in list(doc.paragraphs):
        if p.text.strip().startswith('What to reward when marking'):
            started = True
        if started:
            drop(p)
    # writing lines
    lines = []
    label = doc.add_paragraph()
    run = label.add_run('Write your paragraph below:')
    run.font.bold = True
    lines.append(label)
    for _ in range(10):
        par = doc.add_paragraph()
        r = par.add_run('_' * 80)
        r.font.name = 'Times New Roman'
        par.paragraph_format.space_after = Pt(10)
        lines.append(par)
    ref = anchor
    for par in lines:
        ref.addnext(par._p)
        ref = par._p
    doc.save(OUT)
    print('students saved ->', OUT)
