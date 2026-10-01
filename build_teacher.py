# -*- coding: utf-8 -*-
"""6B Grammar & Writing - Lesson 3: Discussing Solutions (Teacher's version), rev 2.
Rev 2 changes (Helen): (1) Part A = TWO errors per sentence (agreement / tense / parts of
speech / plurals / articles); (2) target patterns end at the main clause - no trailing
participle / purpose phrase; (3) vocabulary table = first 5 rows only; (4) Part D = ONE
paragraph, not a full letter.
Format: A4 20mm margins, Times New Roman 11, 1.15 line spacing, space-after 10pt,
Trebuchet MS bold title block, Table Grid boxes, teacher text red #C00000.
"""
from docx import Document
from docx.shared import Pt, Mm, RGBColor
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import re, sys

RED = RGBColor(0xC0, 0x00, 0x00)
TNR = 'Times New Roman'
TT = 'Trebuchet MS'
B = True
PLAIN = False

OUT = sys.argv[1] if len(sys.argv) > 1 else '/root/f6-grammar-lesson3/out.docx'

doc = Document()
sec = doc.sections[0]
sec.page_width, sec.page_height = Mm(210), Mm(297)
sec.top_margin = sec.bottom_margin = sec.left_margin = sec.right_margin = Mm(20)
st = doc.styles['Normal']
st.font.name = TNR
st.font.size = Pt(11)
st.paragraph_format.line_spacing = 1.15
st.paragraph_format.space_after = Pt(10)


# ---------------------------------------------------------------- helpers
def shade(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill)
    tcPr.append(shd)


def p(runs, space_before=None, space_after=None, align=None, indent=None):
    par = doc.add_paragraph()
    if isinstance(runs, str):
        runs = [(runs, PLAIN, None)]
    for text, bold, colour in runs:
        run = par.add_run(text)
        run.font.bold = bold
        if colour is not None:
            run.font.color.rgb = colour
    if space_before is not None:
        par.paragraph_format.space_before = Pt(space_before)
    if space_after is not None:
        par.paragraph_format.space_after = Pt(space_after)
    if align is not None:
        par.alignment = align
    if indent is not None:
        par.paragraph_format.left_indent = Mm(indent)
    return par


def title_lines(lines):
    for text in lines:
        par = doc.add_paragraph()
        run = par.add_run(text)
        run.font.name = TT
        run.font.size = Pt(10)
        run.font.bold = True
        par.paragraph_format.space_after = Pt(0)
        par.paragraph_format.line_spacing = 1.0
    doc.paragraphs[-1].paragraph_format.space_after = Pt(10)


def heading(text, colour=None):
    return p([(text, B, colour)], space_before=8, space_after=6)


def blank(n=1):
    for _ in range(n):
        par = doc.add_paragraph()
        par.paragraph_format.space_after = Pt(0)


def mk_table(rows, widths_mm, style='Table Grid'):
    t = doc.add_table(rows=rows, cols=len(widths_mm))
    t.style = style
    t.autofit = False
    t.alignment = WD_TABLE_ALIGNMENT.LEFT
    layout = OxmlElement('w:tblLayout')
    layout.set(qn('w:type'), 'fixed')
    t._tbl.tblPr.append(layout)
    for row in t.rows:
        for i, c in enumerate(row.cells):
            c.width = Mm(widths_mm[i])
    return t


def put(cell, runs, first=True, space_after=None):
    if isinstance(runs, str):
        runs = [(runs, PLAIN, None)]
    par = cell.paragraphs[0] if first else cell.add_paragraph()
    for text, bold, colour in runs:
        run = par.add_run(text)
        run.font.bold = bold
        if colour is not None:
            run.font.color.rgb = colour
    if space_after is not None:
        par.paragraph_format.space_after = Pt(space_after)
    return par


def box(rows_of_runs, widths=(170,), fill=None):
    t = mk_table(1, widths)
    cell = t.rows[0].cells[0]
    if fill:
        shade(cell, fill)
    for i, runs in enumerate(rows_of_runs):
        put(cell, runs, first=(i == 0))
    return t


def label_value_table(label, rows, widths=(27.3, 142.7)):
    t = mk_table(len(rows), widths)
    for i, (lab, runs) in enumerate(rows):
        put(t.rows[i].cells[0], [(lab, B, None)])
        put(t.rows[i].cells[1], runs)
    return t


# ================================================================ HEADER
title_lines([
    "St. Joseph's Anglo-Chinese School",
    "F.6 English Language",
    "Grammar & Writing \u2013 Lesson 3: Discussing Solutions",
])
p([("TEACHER\u2019S VERSION \u2014 answers, model answers and teaching notes are printed in red.",
    B, RED)], space_after=12)

# ================================================================ PART A
heading("PART A: PROOFREADING (10 SENTENCES)")
p("Each sentence below uses one of the patterns you learnt in Lesson 1 (introductions) or Lesson 2 "
  "(argument topic sentences), but the content is new. Each sentence contains TWO errors. The errors "
  "are the common ones \u2014 agreement, tense, verb forms, parts of speech, plurals and articles "
  "\u2014 and most of them interrupt the structure of the sentence, so read the whole sentence "
  "before you correct it.")

items = [
    "In recent years, food delivery apps have became an indispensably part of daily life for busy "
    "office workers.",
    "Living in a fast-paced society, commuters is constantly exposed to noise pollutions.",
    "In recent times, the issue of the ageing population has spark widespread concern across the "
    "society.",
    "It is undeniable that early childhood education play a pivotal role in shape children\u2019s "
    "characters.",
    "To begin with, cultural exchange programmes offers students a golden opportunity to broaden "
    "their horizon, paving the way for a global career.",
    "To begin with, joining the school choir offers teenagers a golden opportunity to make new "
    "friends, paved the way for a more confidently attitude to school life.",
    "From the perspective of parents, school counselling services allows them to understand their "
    "children better, which in turn strengthen family relationships.",
    "School libraries serves as a powerful means of encourage students to read, which enables them "
    "to become lifelong learners.",
    "Admittedly, e-books are more convenient, nevertheless, printed books remain irreplaceable, as "
    "reading on paper help readers to concentrate better.",
    "It is undeniable that a balanced diet play a pivot role in maintaining students\u2019 health.",
]
t = mk_table(1, (170,))
cell = t.rows[0].cells[0]
for i, s in enumerate(items):
    put(cell, [(f"{i + 1}. ", B, None), (s, PLAIN, None)], first=(i == 0))
blank(1)

heading("TEACHER\u2019S ANSWERS (LESSON 3, PART A)")
keys = [
    ("1",
     "In recent years, food delivery apps have become an indispensable part of daily life for busy "
     "office workers.",
     ["(a) Verb form (structure): \u201chave became\u201d \u2192 \u201chave become\u201d \u2014 "
      "\u201chave\u201d must be followed by the past participle, so as written the sentence has no "
      "correct verb at all.",
      "(b) Part of speech: \u201can indispensably part\u201d \u2192 \u201can indispensable part\u201d "
      "\u2014 the adjective, not the adverb, modifies the noun \u201cpart\u201d. (Lesson 1, Pattern 1)"]),
    ("2",
     "Living in a fast-paced society, commuters are constantly exposed to noise pollution.",
     ["(a) Agreement: \u201ccommuters is\u201d \u2192 \u201ccommuters are\u201d (plural subject).",
      "(b) Plural: \u201cnoise pollutions\u201d \u2192 \u201cnoise pollution\u201d (\u201cpollution\u201d "
      "is uncountable). (Lesson 1, Pattern 2)"]),
    ("3",
     "In recent times, the issue of the ageing population has sparked widespread concern across "
     "society.",
     ["(a) Verb form (structure): \u201chas spark\u201d \u2192 \u201chas sparked\u201d \u2014 the base "
      "form after \u201chas\u201d leaves the clause without a proper verb.",
      "(b) Article: \u201cacross the society\u201d \u2192 \u201cacross society\u201d (no article when "
      "\u201csociety\u201d is used in the general sense). (Lesson 1, Pattern 3)"]),
    ("4",
     "It is undeniable that early childhood education plays a pivotal role in shaping children\u2019s "
     "characters.",
     ["(a) Agreement: \u201ceducation play\u201d \u2192 \u201ceducation plays\u201d.",
      "(b) Verb form (structure): \u201cin shape\u201d \u2192 \u201cin shaping\u201d \u2014 after the "
      "preposition \u201cin\u201d a gerund is needed; \u201cin shape\u201d reads as a noun phrase and "
      "breaks the pattern. (Lesson 1, Pattern 4)"]),
    ("5",
     "To begin with, cultural exchange programmes offer students a golden opportunity to broaden "
     "their horizons, paving the way for a global career.",
     ["(a) Agreement: \u201cprogrammes offers\u201d \u2192 \u201cprogrammes offer\u201d.",
      "(b) Plural: \u201ctheir horizon\u201d \u2192 \u201ctheir horizons\u201d (more than one "
      "student). (Lesson 2, Pattern 1)"]),
    ("6",
     "To begin with, joining the school choir offers teenagers a golden opportunity to make new "
     "friends, paving the way for a more confident attitude to school life.",
     ["(a) Verb form (structure): \u201cpaved\u201d \u2192 \u201cpaving\u201d \u2014 the comment after "
      "the comma is a participle phrase; \u201cpaved\u201d has no subject and reads as a second main "
      "verb, which breaks the sentence.",
      "(b) Part of speech: \u201cconfidently attitude\u201d \u2192 \u201cconfident attitude\u201d "
      "\u2014 an adverb cannot modify a noun. (Lesson 2, Pattern 1)"]),
    ("7",
     "From the perspective of parents, school counselling services allow them to understand their "
     "children better, which in turn strengthens family relationships.",
     ["(a) Agreement: \u201cservices allows\u201d \u2192 \u201cservices allow\u201d.",
      "(b) Verb form (structure): \u201cstrengthen\u201d \u2192 \u201cstrengthens\u201d \u2014 the "
      "\u201cwhich in turn\u201d clause needs a finite verb. (Lesson 2, Pattern 2)"]),
    ("8",
     "School libraries serve as a powerful means of encouraging students to read, which enables them "
     "to become lifelong learners.",
     ["(a) Agreement: \u201clibraries serves\u201d \u2192 \u201clibraries serve\u201d.",
      "(b) Verb form (structure): \u201cof encourage\u201d \u2192 \u201cof encouraging\u201d \u2014 "
      "after the preposition \u201cof\u201d the gerund is required. (Lesson 2, Pattern 3)"]),
    ("9",
     "Admittedly, e-books are more convenient. Nevertheless, printed books remain irreplaceable, as "
     "reading on paper helps readers to concentrate better.",
     ["(a) Punctuation / sentence boundary (structure): the comma before \u201cnevertheless\u201d "
      "becomes a full stop and \u201cNevertheless\u201d takes a capital letter \u2014 two sentences "
      "cannot be joined by a comma alone.",
      "(b) Agreement: \u201creading \u2026 help\u201d \u2192 \u201chelps\u201d (the subject is the "
      "singular gerund \u201creading\u201d). (Lesson 2, Pattern 4)"]),
    ("10",
     "It is undeniable that a balanced diet plays a pivotal role in maintaining students\u2019 health.",
     ["(a) Agreement: \u201cdiet play\u201d \u2192 \u201cdiet plays\u201d.",
      "(b) Part of speech: \u201ca pivot role\u201d \u2192 \u201ca pivotal role\u201d \u2014 the "
      "adjective \u201cpivotal\u201d is fixed in this pattern. (Lesson 1, Pattern 4)"]),
]
t = mk_table(len(keys) + 1, (10, 78, 82))
hdr = t.rows[0].cells
for i, h in enumerate(("Item", "Corrected sentence", "The TWO errors and why")):
    put(hdr[i], [(h, B, None)])
    shade(hdr[i], 'D9D9D9')
for i, (num, corr, errs) in enumerate(keys, start=1):
    put(t.rows[i].cells[0], [(num, B, RED)])
    put(t.rows[i].cells[1], [(corr, PLAIN, RED)])
    for j, e in enumerate(errs):
        put(t.rows[i].cells[2], [(e, PLAIN, RED)], first=(j == 0))
blank(1)

# ================================================================ PART B
heading("PART B: UPGRADING SENTENCES \u2014 MAKING SUGGESTIONS")
p("In the paragraph that suggests solutions, most students write \u201cSchools can \u2026\u201d, "
  "\u201cThe government can \u2026\u201d and \u201cParents can \u2026\u201d. These sentences are correct, "
  "but they are plain: they show the marker what you think, not how well you can write. Part B gives "
  "you THREE patterns that upgrade exactly these sentences. Each one is only ONE clause, easy to "
  "remember, and each one carries one \u201cwow\u201d feature that impresses the marker.")
p([("Topic for this part: ", B, None),
   ("boost the physical fitness and awareness of citizens", PLAIN, None)], space_after=10)

box([[(u"LEVEL 3 \u2014 THE BASIC SENTENCES STUDENTS USUALLY WRITE", B, None)],
     [("The government can build more sports facilities.", PLAIN, None)],
     [("Schools can organise more sports activities and sports days.", PLAIN, None)],
     [("Parents can exercise with their children at weekends.", PLAIN, None)]],
    fill='D9D9D9')
blank(1)

# Pattern 1
box([[(u"Pattern 1 \u2014 Calling for Immediate Action", B, None)],
     [("Target Formula: ", B, None),
      ("It is high time that [doer] [PAST-TENSE verb phrase].", PLAIN, None)]])
blank(1)
label_value_table("x", [
    ("Level 3:", "The government can build more sports facilities."),
    ("Level 5:", [("It is high time that", B, None),
                  (" the government ", PLAIN, None),
                  ("invested", B, None),
                  (" in free, accessible sports facilities in every district.", PLAIN, None)]),
])
p([("Teaching note: ", B, RED),
   ("\u201cIt is high time that\u201d is followed by the PAST tense \u2014 invested / launched / "
    "reviewed \u2014 never the present tense (\u201cinvests\u201d). If students panic, the easy "
    "alternative is \u201cIt is high time for the government TO invest \u2026\u201d.", PLAIN, RED)])
blank(1)

# Pattern 2
box([[(u"Pattern 2 \u2014 Purpose-First Recommendation", B, None)],
     [("Target Formula: ", B, None),
      ("To [goal], [doer] should make it a priority to [base verb phrase].", PLAIN, None)]])
blank(1)
label_value_table("x", [
    ("Level 3:", "Schools can organise more sports activities and sports days."),
    ("Level 5:", [("To boost students\u2019 physical fitness", B, None),
                  (", schools should ", PLAIN, None),
                  ("make it a priority to", B, None),
                  (" add a daily exercise break to the timetable.", PLAIN, None)]),
])
p([("Teaching note: ", B, RED),
   ("\u201cmake it a priority to\u201d takes the BASE verb (\u201cto add\u201d, never \u201cto "
    "adding\u201d). Opening with \u201cTo + goal\u201d replaces the vague \u201cso that\u201d and "
    "shows the marker your purpose in the first three words.", PLAIN, RED)])
blank(1)

# Pattern 3
box([[(u"Pattern 3 \u2014 Shared Responsibility", B, None)],
     [("Target Formula: ", B, None),
      ("[Doer], in collaboration with [second party], should [verb phrase].", PLAIN, None)]])
blank(1)
label_value_table("x", [
    ("Level 3:", "Parents can exercise with their children at weekends."),
    ("Level 5:", [("Parents, in collaboration with schools,", B, None),
                  (" should encourage their children to adopt a more active lifestyle.",
                   PLAIN, None)]),
])
p([("Teaching note: ", B, RED),
   ("\u201cin collaboration with\u201d = working together with. It is far stronger than \u201cand\u201d "
    "or \u201calso\u201d and shows the marker that you understand shared responsibility. The doer is "
    "swappable: \u201cThe government, in collaboration with district councils, should \u2026\u201d.",
    PLAIN, RED)])
p([("Teacher\u2019s note on the design: ", B, RED),
   ("the three patterns are deliberately kept to ONE clause. Endings of the \u201c, so as to "
    "\u2026\u201d / \u201c, thereby + \u2013ing\u201d type were left out this time \u2014 level 3 "
    "students lose control of the main clause as soon as they open a second one. Add the endings "
    "back only when the class can hold the short form.", PLAIN, RED)])
blank(1)

heading("LEVEL 3 \u2192 LEVEL 5 VOCABULARY (physical fitness and awareness)")
vocab = [
    ("make people healthier", "enhance citizens\u2019 physical well-being",
     "Regular exercise enhances citizens\u2019 physical well-being."),
    ("tell people about the importance of exercise",
     "raise public awareness of the importance of regular exercise",
     "A city-wide campaign would raise public awareness of the importance of regular exercise."),
    ("do exercise", "take regular exercise / adopt an active lifestyle",
     "Citizens who take regular exercise are far less likely to fall ill."),
    ("build more sports places", "provide accessible sports facilities",
     "Accessible sports facilities encourage even the busiest workers to exercise."),
    ("stop people from being lazy", "combat sedentary lifestyles",
     "A morning workout scheme can combat sedentary lifestyles among office workers."),
]
t = mk_table(len(vocab) + 1, (8, 60, 102))
for i, h in enumerate(("#", "Level 3 (what students write)", "Level 5 upgrade (key words in bold)")):
    put(t.rows[0].cells[i], [(h, B, None)])
    shade(t.rows[0].cells[i], 'D9D9D9')
for i, (l3, l5, eg) in enumerate(vocab, start=1):
    put(t.rows[i].cells[0], [(str(i), B, None)])
    put(t.rows[i].cells[1], [(l3, PLAIN, None)])
    put(t.rows[i].cells[2], [(l5, B, None)])
    par = t.rows[i].cells[2].add_paragraph()
    run = par.add_run("e.g. " + eg)
    run.font.italic = True
blank(1)

# ================================================================ PART C
heading("PART C: PRACTICE SETS (6 SETS)")
p("Upgrade each simple sentence with the pattern shown, and keep the pattern to ONE clause, as "
  "taught in Part B. The topics are different from the fitness topic in Part B. Sets 4\u20136 are "
  "more challenging: you must join the TWO simple sentences into ONE upgraded sentence.")

sets = [
    ("Set 1. Topic: Teenage Sleep Deprivation", "Use Pattern 1",
     ["The government can tell teenagers about the importance of sleep."],
     "It is high time that the government raised teenagers\u2019 awareness of the importance of sleep.",
     "Accept any past-tense verb phrase: \u201claunched a campaign on healthy sleeping habits\u201d, "
     "\u201cmade sleep education part of the junior curriculum\u201d."),
    ("Set 2. Topic: Food Waste on Campus", "Use Pattern 2",
     ["Schools can ask students not to waste food in the canteen."],
     "To cut food waste on campus, schools should make it a priority to offer smaller portions and "
     "free second helpings.",
     "Accept \u201cTo reduce food waste \u2026\u201d and any sensible verb phrase after \u201cmake it a "
     "priority to\u201d."),
    ("Set 3. Topic: Cyberbullying", "Use Pattern 3",
     ["Parents can talk to their children about cyberbullying."],
     "Parents, in collaboration with schools, should discuss online safety openly with their "
     "children.",
     "Accept any sensible second party (the school, social workers) and any formal verb phrase "
     "(\u201cset clear rules about online behaviour at home\u201d)."),
    ("Set 4. Topic: Youth Mental Health (more challenging)", "Use Pattern 1",
     ["The government can spend more money on mental health services.",
      "Young people need these services at school."],
     "It is high time that the government invested more in school-based mental health services for "
     "young people.",
     "Idea 2 is folded into the modifier \u201cschool-based\u201d, so the pattern stays one clean "
     "clause. Also accept \u201c\u2026 invested more in mental health services in schools.\u201d"),
    ("Set 5. Topic: Declining Reading Habits (more challenging)", "Use Pattern 2",
     ["Schools can keep the library open after school.",
      "Students need a quiet place to read."],
     "To give students a quiet place to read, schools should make it a priority to keep the library "
     "open after school.",
     "Idea 2 becomes the fronted goal; the second plain sentence must disappear into the pattern."),
    ("Set 6. Topic: Care for the Elderly (more challenging)", "Use Pattern 3",
     ["Teenagers can visit elderly neighbours who live alone.",
      "Community centres can arrange these visits."],
     "Teenagers, in collaboration with community centres, should organise regular visits to elderly "
     "neighbours who live alone.",
     "Idea 2 becomes the second party after \u201cin collaboration with\u201d. Watch the plural: "
     "\u201cthe elderly\u201d is already plural \u2014 never \u201cthe elderlies\u201d."),
]
for title, pattern, simples, answer, note in sets:
    p([(title, B, None)], space_before=6, space_after=4)
    rows = [("Simple sentence" if len(simples) == 1 else "Simple sentences", simples[0])]
    for extra in simples[1:]:
        rows.append(("", extra))
    rows.append((pattern, [(answer, True, RED)]))
    label_value_table("x", rows)
    p([("Key note: ", B, RED), (note, PLAIN, RED)], space_after=8)

# ================================================================ PART D
heading("PART D: WRITING PRACTICE")
box([[("HKDSE Adapted Paper 2 Prompt:", PLAIN, None)],
     [("You are a Form 6 student. Your school is considering cutting one of the two weekly PE "
       "lessons so that F.6 students can have an extra period for examination preparation. Write "
       "a letter to the Principal expressing your opinion on this idea.", PLAIN, None)]],
    fill='F2F2F2')
p([("Task requirements:", B, None)], space_before=8, space_after=4)
for line in [
    "Write ONE body paragraph (about 120 words) for this letter \u2014 not the whole letter.",
    "State your view on the proposal, give at least one reason and at least one suggestion.",
    "Use at least TWO of today\u2019s three patterns and underline every target structure you use.",
    "Follow the Lesson 2 chain: [Topic sentence] \u2192 [Explain] \u2192 [Example] \u2192 [Back to "
    "main idea].",
]:
    p(line, space_after=0, indent=4)

p([("Teacher\u2019s model answer (one paragraph, 129 words)", B, None)], space_before=10, space_after=4)
model = [
    ("Although an extra period for examination preparation sounds attractive, cutting one of our two "
     "weekly PE lessons would do F.6 students more harm than good.",
     "[Topic sentence \u2014 states the stance]"),
    ("To keep us physically fit during the examination year, the school should make it a priority to "
     "protect both lessons and to improve what happens in them.",
     "[Suggestion 1 \u2014 today\u2019s Pattern 2]"),
    ("A single game of badminton or basketball clears the mind far more effectively than another "
     "hour of past papers, and the concentration we gain is worth the time we spend.",
     "[Explain + Example]"),
    ("It is also high time that the school looked for the extra period somewhere else, such as the "
     "morning assembly, rather than taking it from the only lesson that keeps us moving.",
     "[Suggestion 2 \u2014 today\u2019s Pattern 1]"),
    ("I hope the Principal will reconsider the proposal before our health pays for our grades.",
     "[Back to main idea \u2014 Link]"),
]
for text, label in model:
    p([(text, PLAIN, None), ("  " + label, False, RED)], space_after=6)

p([("What to reward when marking", B, RED)], space_before=10, space_after=4)
for line in [
    "Content: a clear stance on the cut itself, plus one realistic alternative \u2014 not just "
    "\u201cPE is good for us\u201d.",
    "Language: the target patterns used accurately; formal verbs (protect, invest in, combat) "
    "rather than \u201chelp\u201d; no \u201cI think it is good / bad\u201d.",
    "Organisation: topic sentence first, then explain and exemplify, and a closing sentence that "
    "returns to the proposal.",
]:
    p(line, space_after=0, indent=4)

p([("Common errors to watch when marking", B, RED)], space_before=8, space_after=4)
for line in [
    "\u201cIt is high time that the government promote \u2026\u201d \u2192 the past tense is needed "
    "(\u201cpromoted\u201d).",
    "\u201cmake it a priority to + \u2013ing\u201d \u2192 base verb only (\u201cto add\u201d).",
    "\u201cin collaboration with\u201d used as a verb (\u201cparents collaborate with schools "
    "should\u2026\u201d) \u2192 keep it as a phrase between the subject and the modal.",
    "Sentence boundaries: a participle phrase after a comma with no matching subject "
    "(\u201cLiving in a fast-paced society, academic pressure \u2026\u201d), or a new sentence "
    "after a comma (\u201c\u2026 flexibility, nevertheless, they cannot \u2026\u201d) \u2192 full "
    "stop and a capital letter.",
    "\u201cthe elderly\u201d treated as singular (\u201cthe elderlies\u201d, \u201cthe elderly "
    "feels\u201d).",
]:
    p(line, space_after=0, indent=4)

doc.save(OUT)

# ---------------------------------------------------------------- verify + word count
d = Document(OUT)
print('saved', OUT)
print('tables:', len(d.tables), '| paragraphs:', len(d.paragraphs))
count = len(re.findall(r"[A-Za-z\u2019'-]+", ' '.join(t for t, _ in model)))
print('model paragraph words:', count)
print('placeholder AA present:', 'AA words' in '\n'.join(x.text for x in d.paragraphs))
