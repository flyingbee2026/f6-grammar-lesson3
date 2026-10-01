# -*- coding: utf-8 -*-
"""Rev 5, built FROM Helen's own uploaded docx (rev4/helen_teacher.docx) so that font, size,
style and spacing stay exactly hers: every new paragraph is a deep copy of one of HER
paragraphs, and only the text / bold / colour runs are replaced.

teacher  -> her file + Part A answer key + new Part D (HK-wide issue, three model paragraphs)
students -> her file minus every teacher-only item, with answer space and three writing blocks
Usage: python3 build_rev5.py teacher|students
"""
import sys
from copy import deepcopy
from docx import Document
from docx.shared import Pt, Mm, RGBColor
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from docx.oxml import OxmlElement

SRC = '/root/f6-grammar-lesson3/rev4/helen_teacher.docx'
MODE = sys.argv[1] if len(sys.argv) > 1 else 'teacher'
OUT = ('/root/f6-grammar-lesson3/6B Grammar Lesson 3 - Discussing Solutions (Teacher\'s).docx'
       if MODE == 'teacher' else
       '/root/f6-grammar-lesson3/6B Grammar Lesson 3 - Discussing Solutions (Students).docx')
RED = 'C00000'


# --------------------------------------------------------------- helpers
def drop(el):
    e = getattr(el, '_element', el)
    e.getparent().remove(e)


def find_par(doc, start):
    for p in doc.paragraphs:
        if p.text.strip().startswith(start):
            return p
    return None


def clone_par(model_par, runs):
    """New <w:p> copying the model paragraph's pPr exactly; runs = [(text, bold, colour)]."""
    el = deepcopy(model_par._p)
    for child in list(el):
        if child.tag in (qn('w:r'), qn('w:hyperlink')):
            el.remove(child)
    par = Paragraph(el, model_par._parent)
    for text, bold, colour in runs:
        run = par.add_run(text)
        run.font.bold = bold
        if colour:
            run.font.color.rgb = RGBColor.from_string(colour)
    return el


def chain_after(anchor_el, els):
    cur = anchor_el
    for e in els:
        cur.addnext(e)
        cur = e
    return cur


def set_cell_text(cell, runs, first=True):
    par = cell.paragraphs[0] if first else cell.add_paragraph()
    if first:
        for r in list(par.runs):
            drop(r)
    for text, bold, colour in runs:
        run = par.add_run(text)
        run.font.bold = bold
        if colour:
            run.font.color.rgb = RGBColor.from_string(colour)
    return par


def shade(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear'); shd.set(qn('w:color'), 'auto'); shd.set(qn('w:fill'), fill)
    tcPr.append(shd)


def make_table(doc, rows, widths_mm):
    t = doc.add_table(rows=rows, cols=len(widths_mm))
    t.style = 'Table Grid'
    t.autofit = False
    t.alignment = WD_TABLE_ALIGNMENT.LEFT
    layout = OxmlElement('w:tblLayout'); layout.set(qn('w:type'), 'fixed')
    t._tbl.tblPr.append(layout)
    for r in t.rows:
        for i, c in enumerate(r.cells):
            c.width = Mm(widths_mm[i])
    return t


def is_red(par, needle=None):
    if needle and needle not in par.text:
        return False
    return any(r.font.color is not None and r.font.color.type is not None
               and str(r.font.color.rgb) == RED for r in par.runs)


# --------------------------------------------------------------- shared content
PROMPT = ("You are a Form 6 student. Hong Kong\u2019s population is ageing rapidly, and the needs "
          "of elderly citizens have become a pressing social issue. Write a letter to the editor of "
          "the Hong Kong Young Post, suggesting THREE ways in which our city can care for its "
          "growing number of elderly citizens.")

REQUIREMENTS = [
    "Write THREE body paragraphs (about 100 words each) for this letter.",
    "Begin each paragraph with a topic sentence that makes ONE suggestion, using a different pattern "
    "from today\u2019s lesson in each paragraph.",
    "Then explain why the suggestion helps, give an example or some evidence, and round the "
    "paragraph off.",
    "You do NOT need to use the Lesson 1 or Lesson 2 patterns in the explanation, the example or "
    "the round-off.",
    "Follow the chain: [Topic sentence] \u2192 [Explain] \u2192 [Example / evidence] \u2192 "
    "[Round off].",
]

MODELS = [
    ("Suggested paragraph 1 \u2014 community health care (Pattern 1)",
     [("There is an urgent need for the government to invest in community health services for the "
       "elderly.", "  [Topic sentence \u2014 Pattern 1]"),
      ("Many older citizens find a trip to a hospital exhausting, and a minor complaint that nobody "
       "checks in time easily grows into a serious illness that could have been prevented.", "  [Explain]"),
      ("Take an elderly resident with high blood pressure: a fortnightly visit from a community nurse "
       "would keep the condition under control, while an appointment at a public hospital may not "
       "come for months.", "  [Example]"),
      ("Looking after the elderly in their own neighbourhood therefore protects both their health "
       "and the public purse that pays for their treatment.", "  [Round off]")]),
    ("Suggested paragraph 2 \u2014 community centres against loneliness (Pattern 2)",
     [("To ease the loneliness that many elderly people face, district councils should make it a "
       "priority to turn community centres into lively daily meeting places.",
       "  [Topic sentence \u2014 Pattern 2]"),
      ("Isolation damages both the mind and the body, and it is far cheaper to prevent than to treat "
       "with medicine.", "  [Explain]"),
      ("A centre that runs morning tai chi and afternoon board games, for instance, gives a retired "
       "resident a fixed routine, a reason to leave the flat and the friendships that come with it.", "  [Example]"),
      ("Simple, cheap and human, these centres are the best medicine our city can offer its older "
       "residents.", "  [Round off]")]),
    ("Suggested paragraph 3 \u2014 work and purpose for retired citizens (Pattern 3)",
     [("The government, in collaboration with businesses and charities, should create flexible "
       "part-time jobs for retired workers.", "  [Topic sentence \u2014 Pattern 3]"),
      ("Many experienced citizens still have energy and skills to offer, yet they are pushed out of "
       "the workforce the moment they turn sixty-five, which leaves both the economy and their own "
       "sense of purpose the poorer.", "  [Explain]"),
      ("A retired accountant could mentor young clerks two mornings a week, or a former nurse could "
       "advise a care home part-time, with the government covering part of the wage.", "  [Example]"),
      ("Giving older citizens a role keeps their skills alive and lightens the welfare burden on the "
       "younger generation.", "  [Round off]")]),
]

MARKING = [
    ("Content: ", "three DIFFERENT suggestions that are realistic for Hong Kong, each with its own "
                  "paragraph \u2014 not three versions of the same idea."),
    ("Language: ", "the topic sentence uses the target pattern accurately; the explanation, the "
                   "example and the round-off may be in the students\u2019 own plain English."),
    ("Organisation: ", "one suggestion per paragraph, each following [Topic sentence] \u2192 "
                       "[Explain] \u2192 [Example / evidence] \u2192 [Round off]."),
]

KEYS = [
    ("1", "In recent years, food delivery apps have become an indispensable part of daily life for "
          "busy office workers.",
     ["(a) Agreement: \u201capps has\u201d \u2192 \u201capps have\u201d (plural subject).",
      "(b) Verb form: \u201chas became\u201d \u2192 \u201chave become\u201d \u2014 after "
      "\u201chave\u201d the past participle is needed. (Lesson 1, Pattern 1)"]),
    ("2", "Living in a fast-paced society, commuters are constantly exposed to noise pollution.",
     ["(a) Part of speech: \u201cfast-pace\u201d \u2192 \u201cfast-paced\u201d (a compound adjective "
      "modifies \u201csociety\u201d)",
      "(b) Agreement: \u201ccommuters is\u201d \u2192 \u201ccommuters are\u201d. (Lesson 1, Pattern 2)"]),
    ("3", "In recent times, the issue of the ageing population has sparked widespread concern across "
          "society.",
     ["(a) Plural: \u201cIn recent time\u201d \u2192 \u201cIn recent times\u201d (the fixed pattern "
      "uses the plural).",
      "(b) Verb form: \u201chas spark\u201d \u2192 \u201chas sparked\u201d \u2014 the base form after "
      "\u201chas\u201d leaves the clause without a verb.",
      "(c) Also: \u201cageing population\u201d \u2192 \u201cthe ageing population\u201d (article). "
      "(Lesson 1, Pattern 3)"]),
    ("4", "It is undeniable that early childhood education plays a pivotal role in shaping "
          "children\u2019s characters.",
     ["(a) Agreement: \u201ceducation play\u201d \u2192 \u201ceducation plays\u201d.",
      "(b) Verb form: \u201cin shape\u201d \u2192 \u201cin shaping\u201d \u2014 a gerund is needed "
      "after the preposition \u201cin\u201d. (Lesson 1, Pattern 4)"]),
    ("5", "To begin with, cultural exchange programmes offer students a golden opportunity to "
          "broaden their horizons.",
     ["(a) Agreement: \u201cprogramme offer\u201d \u2192 \u201cprogrammes offer\u201d (also accepted: "
      "\u201ca cultural exchange programme offers\u201d).",
      "(b) Article: \u201cstudents golden opportunity\u201d \u2192 \u201cstudents a golden "
      "opportunity\u201d.",
      "(c) Also: \u201ctheir horizon\u201d \u2192 \u201ctheir horizons\u201d. (Lesson 2, Pattern 1)"]),
    ("6", "Furthermore, joining the school choir offers teenagers a valuable chance to make new "
          "friends, which in turn boosts their confidence.",
     ["(a) Agreement: \u201cchoir offer\u201d \u2192 \u201cchoir offers\u201d (the subject is the "
      "gerund phrase \u201cjoining the school choir\u201d).",
      "(b) Word choice: \u201ca valuable change\u201d \u2192 \u201ca valuable chance\u201d.",
      "(c) Also: \u201cwhich in turn boost\u201d \u2192 \u201cboosts\u201d (the clause refers to the "
      "whole idea). (Lesson 2, Patterns 1 and 2)"]),
    ("7", "From the perspective of parents, school counselling services allow them to understand "
          "their children better, which in turn strengthens family relationships.",
     ["(a) Pronoun reference: \u201callow it to understand\u201d \u2192 \u201callow them to "
      "understand\u201d (the parents).",
      "(b) Verb form: \u201cwhich in turn strengthen\u201d \u2192 \u201cstrengthens\u201d \u2014 the "
      "clause needs a finite verb. (Lesson 2, Pattern 2)"]),
    ("8", "School libraries serve as a powerful means of encouraging students to read, which enables "
          "them to become lifelong learners.",
     ["(a) Word form: \u201ca powerful mean\u201d \u2192 \u201ca powerful means\u201d (the fixed "
      "pattern).",
      "(b) Verb form: \u201cof encourage\u201d \u2192 \u201cof encouraging\u201d \u2014 a gerund is "
      "needed after the preposition \u201cof\u201d. (Lesson 2, Pattern 3)"]),
    ("9", "Admittedly, e-books are more convenient; nevertheless, printed books remain "
          "irreplaceable, as reading on paper helps readers to concentrate better.",
     ["(a) Part of speech: \u201cmore convenience\u201d \u2192 \u201cmore convenient\u201d (an "
      "adjective is needed after \u201care\u201d).",
      "(b) Agreement: \u201creading \u2026 help\u201d \u2192 \u201chelps\u201d (the subject is the "
      "singular gerund \u201creading\u201d).",
      "Note: the semicolon before \u201cnevertheless\u201d is correct \u2014 a comma would be a "
      "comma splice. (Lesson 2, Pattern 4)"]),
    ("10", "It is undeniable that a balanced diet plays a pivotal role in maintaining students\u2019 "
           "health.",
     ["(a) Agreement: \u201cdiet play\u201d \u2192 \u201cdiet plays\u201d.",
      "(b) Possessive / plural: \u201cstudent\u2019s health\u201d \u2192 \u201cstudents\u2019 "
      "health\u201d (all the students). (Lesson 1, Pattern 4)"]),
]

doc = Document(SRC)


def set_run_text(par, idx, text):
    par.runs[idx].text = text


def del_runs_from(par, idx):
    for r in list(par.runs)[idx:]:
        drop(r)


def update_pattern1():
    """Replace 'It is high time that [doer] [PAST-TENSE verb phrase]' with a pattern that works
    as a topic sentence: 'There is an urgent need for [doer] to [base verb phrase]'."""
    # Part B formula box
    for t in doc.tables:
        cell = t.rows[0].cells[0]
        if t.rows[0].cells[0].text.startswith('Pattern 1 \u2014 Calling'):
            par = cell.paragraphs[1]
            set_run_text(par, 1, 'There is an urgent need for [doer] to [base verb phrase].')
            del_runs_from(par, 2)
        for row in t.rows:
            if len(row.cells) < 2:
                continue
            if 'free, accessible sports facilities' in row.cells[1].text:
                par = row.cells[1].paragraphs[0]
                set_run_text(par, 0, 'There is an urgent need for')
                set_run_text(par, 2, 'to invest')
    # Part B teaching note
    for par in doc.paragraphs:
        if par.text.startswith('Teaching note:') and 'It is high time' in par.text:
            set_run_text(par, 1, '\u201cThere is an urgent need for\u201d is followed by [doer] + TO + '
                                 'base verb \u2014 never the base verb on its own (\u201cfor the '
                                 'government invest\u201d) or a gerund (\u201cfor the government '
                                 'investing\u201d). Use it when the problem cannot wait; '
                                 '\u201cpressing\u201d and \u201cgrowing\u201d can replace '
                                 '\u201curgent\u201d.')
    # Part C set answers
    for t in doc.tables:
        for row in t.rows:
            if len(row.cells) < 2:
                continue
            if 'It is high time that the government raised' in row.cells[1].text:
                set_run_text(row.cells[1].paragraphs[0], 0,
                             'There is an urgent need for the government to raise teenagers\u2019 '
                             'awareness of the importance of sleep.')
            if 'It is high time that the government invested more' in row.cells[1].text:
                set_run_text(row.cells[1].paragraphs[0], 0,
                             'There is an urgent need for the government to invest more in '
                             'school-based mental health services for young people.')
    # Part C key note for set 1
    for par in doc.paragraphs:
        if par.text.startswith('Key note:') and 'past-tense verb phrase' in par.text:
            set_run_text(par, 1, 'Accept any sensible base verb phrase: \u201claunch a campaign on '
                                 'healthy sleeping habits\u201d, \u201cmake sleep education part of '
                                 'the junior curriculum\u201d.')
    # Common errors bullet
    for par in doc.paragraphs:
        if par.text.startswith('\u201cIt is high time that the government promote'):
            set_run_text(par, 0, '\u201cThere is an urgent need for the government invest \u2026\u201d '
                                 '\u2192 the infinitive is needed (\u201cto invest\u201d).')


update_pattern1()

# formatting templates cloned from HER paragraphs
TPL_BODY = find_par(doc, 'State your view on the proposal')      # plain black body line
TPL_BOLD = find_par(doc, 'Task requirements:')                   # bold black heading line
TPL_RED = find_par(doc, 'Teaching note:')                        # red note line
TPL_BLANK = None
for p in doc.paragraphs:                                         # her space-after-0 blank line
    if not p.text.strip() and p.paragraph_format.space_after == Pt(0):
        TPL_BLANK = p
        break
assert TPL_BODY is not None and TPL_BOLD is not None and TPL_RED is not None and TPL_BLANK is not None

# ------------------------------------------------------------------ Part A table
key_tbl = None
for t in doc.tables:
    if len(t.columns) >= 3 and 'The TWO errors' in t.rows[0].cells[2].text:
        key_tbl = t
        break
assert key_tbl is not None

if MODE == 'teacher':
    for num, corrected, errs in KEYS:
        row = key_tbl.add_row()
        c0, c1, c2 = row.cells
        c0.width, c1.width, c2.width = Mm(10), Mm(78), Mm(82)
        set_cell_text(c0, [(num, True, RED)])
        set_cell_text(c1, [(corrected, False, RED)])
        for i, e in enumerate(errs):
            set_cell_text(c2, [(e, False, RED)], first=(i == 0))
else:
    banner = find_par(doc, 'TEACHER\u2019S VERSION')
    if banner is not None:
        drop(banner)
    heading = find_par(doc, 'TEACHER\u2019S ANSWERS')
    for r in list(heading.runs)[1:]:
        drop(r)
    heading.runs[0].text = 'Write the corrected sentences below:'
    answer_tbl = make_table(doc, 11, (10, 160))
    set_cell_text(answer_tbl.rows[0].cells[0], [('Item', True, None)])
    set_cell_text(answer_tbl.rows[0].cells[1], [('Corrected sentence', True, None)])
    shade(answer_tbl.rows[0].cells[0], 'D9D9D9')
    shade(answer_tbl.rows[0].cells[1], 'D9D9D9')
    for i in range(10):
        row = answer_tbl.rows[i + 1]
        set_cell_text(row.cells[0], [('%d.' % (i + 1), True, None)])
        row.cells[1].add_paragraph()
    key_tbl._tbl.addprevious(answer_tbl._tbl)
    drop(key_tbl)

# ------------------------------------------------------------------ Part D: prompt
prompt_tbl = None
for t in doc.tables:
    if 'HKDSE Adapted Paper 2 Prompt' in t.rows[0].cells[0].text:
        prompt_tbl = t
        break
assert prompt_tbl is not None
pc = prompt_tbl.rows[0].cells[0]
for par in pc.paragraphs:
    if par.text.strip().startswith('You are a Form 6 student'):
        for r in list(par.runs):
            drop(r)
        run = par.add_run(PROMPT)
        run.font.bold = False
        break
else:
    set_cell_text(pc, [(PROMPT, False, None)], first=False)

# Part D: task requirements (both versions)
for old in ['Write ONE body paragraph', 'State your view on the proposal',
            'Use at least TWO of today\u2019s three patterns', 'Follow the Lesson 2 chain']:
    p = find_par(doc, old)
    if p is not None:
        drop(p)
anchor = find_par(doc, 'Task requirements:')._p
anchor = chain_after(anchor, [clone_par(TPL_BODY, [(line, False, None)]) for line in REQUIREMENTS])

# Part D: old model answer paragraphs
for old in ['Teacher\u2019s model answer', 'Although an extra period', 'To keep us physically fit',
            'A single game of badminton', 'It is also high time that the school looked',
            'I hope the Principal will reconsider']:
    p = find_par(doc, old)
    if p is not None:
        drop(p)

if MODE == 'teacher':
    new_els = [clone_par(TPL_BLANK, [('', False, None)]),
               clone_par(TPL_BOLD, [("Teacher\u2019s three suggested paragraphs (the body of the "
                                     "letter \u2014 the opening \u201cDear Editor,\u201d and the "
                                     "closing are the class\u2019s own work)", True, None)])]
    for title, sentences in MODELS:
        new_els.append(clone_par(TPL_BOLD, [(title, True, None)]))
        runs = []
        for text, tag in sentences:
            runs.append((text, False, None))
            runs.append((tag, False, RED))
        new_els.append(clone_par(TPL_BODY, runs))
    new_els.append(clone_par(TPL_RED, [
        ('Teacher\u2019s note: ', True, RED),
        ('check every paragraph against the chain [Topic sentence] \u2192 [Explain] \u2192 '
         '[Example / evidence] \u2192 [Round off]. The topic sentence is the only place where a '
         'target pattern is required.', False, RED)]))
    chain_after(anchor, new_els)

    # Part D: marking bullets rewritten
    for old, (lead_txt, body) in zip(['Content: a clear stance', 'Language: the target patterns',
                                      'Organisation: topic sentence first'], MARKING):
        p = find_par(doc, old)
        assert p is not None, old
        for r in list(p.runs)[1:]:
            drop(r)
        p.runs[0].text = lead_txt
        p.add_run(body)
else:
    # writing blocks
    els = [clone_par(TPL_BLANK, [('', False, None)]),
           clone_par(TPL_BOLD, [('Write your three paragraphs below:', True, None)])]
    for n in range(1, 4):
        els.append(clone_par(TPL_BOLD, [('Paragraph %d:' % n, True, None)]))
        for _ in range(7):
            els.append(clone_par(TPL_BLANK, [('_' * 80, False, None)]))
    chain_after(anchor, els)

    # drop the red teacher notes and the red marking sections
    for p in list(doc.paragraphs):
        if is_red(p, 'Teaching note:') or is_red(p, 'Teacher\u2019s note on the design:') \
                or is_red(p, 'Key note:'):
            drop(p)
    # blank Part C answer cells
    for t in doc.tables:
        txt = ' '.join(c.text for r in t.rows for c in r.cells)
        if len(t.columns) == 2 and 'Use Pattern' in txt:
            cell = t.rows[-1].cells[1]
            for p in list(cell.paragraphs)[1:]:
                drop(p)
            for r in list(cell.paragraphs[0].runs):
                drop(r)
            cell.add_paragraph()
            cell.add_paragraph()
    started = False
    for p in list(doc.paragraphs):
        if p.text.strip().startswith('What to reward when marking'):
            started = True
        if started:
            drop(p)

doc.save(OUT)
print(MODE, 'saved ->', OUT)
