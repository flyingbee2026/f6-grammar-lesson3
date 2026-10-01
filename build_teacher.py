# -*- coding: utf-8 -*-
"""6B Grammar & Writing - Lesson 3: Discussing Solutions (Teacher's version).
Follows the 6B series format: A4 20mm margins, Times New Roman 11, line spacing 1.15,
space-after 10pt, Trebuchet MS bold title block, Table Grid boxes, teacher text red #C00000.
"""
from docx import Document
from docx.shared import Pt, Mm, RGBColor
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import re, os, sys

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
    for r in runs:
        text, bold, colour = r
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
    """Trebuchet MS bold 10pt header block, no space between lines."""
    for i, text in enumerate(lines):
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
        for r in par.runs:
            pass


def mk_table(rows, widths_mm, style='Table Grid'):
    t = doc.add_table(rows=rows, cols=len(widths_mm))
    t.style = style
    t.autofit = False
    t.alignment = WD_TABLE_ALIGNMENT.LEFT
    tblPr = t._tbl.tblPr
    layout = OxmlElement('w:tblLayout')
    layout.set(qn('w:type'), 'fixed')
    tblPr.append(layout)
    for row in t.rows:
        for i, c in enumerate(row.cells):
            c.width = Mm(widths_mm[i])
    return t


def put(cell, runs, first=True, space_after=None):
    """Write runs into a cell: reuse the first empty paragraph, then add more."""
    if isinstance(runs, str):
        runs = [(runs, PLAIN, None)]
    if first:
        par = cell.paragraphs[0]
    else:
        par = cell.add_paragraph()
    for text, bold, colour in runs:
        run = par.add_run(text)
        run.font.bold = bold
        if colour is not None:
            run.font.color.rgb = colour
    if space_after is not None:
        par.paragraph_format.space_after = Pt(space_after)
    return par


def box(rows_of_runs, widths=(170,), fill=None):
    """1-cell bordered box; rows_of_runs = list of run-lists (one paragraph each)."""
    t = mk_table(1, widths)
    cell = t.rows[0].cells[0]
    if fill:
        shade(cell, fill)
    for i, runs in enumerate(rows_of_runs):
        put(cell, runs, first=(i == 0))
    return t


def label_value_table(label, rows, widths=(27.3, 142.7), label_bold=True):
    """label | content rows, used for the drill sets and pattern examples."""
    t = mk_table(len(rows), widths)
    for i, (lab, runs) in enumerate(rows):
        put(t.rows[i].cells[0], [(lab, label_bold, None)])
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
p("Each sentence below contains ONE error in the sentence patterns you learnt in Lesson 1 "
  "(introductions) and Lesson 2 (argument topic sentences). Find the error and write the "
  "corrected sentence.")

items = [
    "In recent years, social media has become indispensable part of daily life for young people.",
    "Living in a fast-paced society, academic pressure has become increasingly heavy for students.",
    "In recent times, the issue of plastic pollution have sparked widespread concern across society.",
    "It is undeniable that sleep plays a pivotal role to shape students\u2019 memory and concentration.",
    "To begin with, part-time work offers students to gain real-world experience, paving the way for their future careers.",
    "To begin with, volunteering offers teenagers a golden opportunity to meet new people, paved the way for a wider social circle.",
    "From the perspective of students, a balanced timetable allows them to manage their time wisely, which in turn boosting their academic performance.",
    "Team sports serve as a powerful means of relieve stress, which enables students to recharge both physically and mentally.",
    "Admittedly, online lessons offer students greater flexibility, nevertheless, they cannot replace "
    "face-to-face interaction, as real-time discussion sparks deeper engagement.",
    "Reading serves as a powerful means of broadening students\u2019 horizons, which enable them to "
    "understand the world beyond the classroom.",
]
t = mk_table(1, (170,))
cell = t.rows[0].cells[0]
for i, s in enumerate(items):
    put(cell, [(f"{i + 1}. ", B, None), (s, PLAIN, None)], first=(i == 0))
blank(1)

heading("TEACHER\u2019S ANSWERS (LESSON 3, PART A)")
keys = [
    ("1",
     "In recent years, social media has become an indispensable part of daily life for young people.",
     "Missing article. The pattern is \u201can indispensable part of\u201d: the singular countable noun "
     "\u201cpart\u201d needs \u201can\u201d before it. (Lesson 1, Pattern 1)"),
    ("2",
     "Living in a fast-paced society, students are constantly exposed to increasing academic pressure.",
     "Dangling participle. After \u201cLiving in a fast-paced society\u201d, the subject of the main clause "
     "must be the people who do the living \u2014 \u201cstudents\u201d, not \u201cacademic pressure\u201d. "
     "(Lesson 1, Pattern 2)"),
    ("3",
     "In recent times, the issue of plastic pollution has sparked widespread concern across society.",
     "Subject\u2013verb agreement. The subject is \u201cthe issue\u201d (singular), so the verb is "
     "\u201chas sparked\u201d, not \u201chave sparked\u201d. (Lesson 1, Pattern 3)"),
    ("4",
     "It is undeniable that sleep plays a pivotal role in shaping students\u2019 memory and concentration.",
     "Wrong preposition. The fixed collocation is \u201cplay a pivotal role IN + gerund\u201d, never "
     "\u201crole to + verb\u201d. (Lesson 1, Pattern 4)"),
    ("5",
     "To begin with, part-time work offers students a golden opportunity to gain real-world experience, "
     "paving the way for their future careers.",
     "\u201cOffer\u201d takes two objects: offer somebody something. Write \u201coffer students a golden "
     "opportunity to gain\u201d \u2014 never \u201coffer students to gain\u201d. (Lesson 2, Pattern 1)"),
    ("6",
     "To begin with, volunteering offers teenagers a golden opportunity to meet new people, paving the "
     "way for a wider social circle.",
     "Participle form. The comment after the comma describes the previous clause, so it must be the "
     "\u2013ing participle \u201cpaving\u201d \u2014 the past participle \u201cpaved\u201d has no subject. "
     "(Lesson 2, Pattern 1)"),
    ("7",
     "From the perspective of students, a balanced timetable allows them to manage their time wisely, "
     "which in turn boosts their academic performance.",
     "Verb form after \u201cwhich in turn\u201d. The clause needs a finite present-tense verb "
     "(\u201cboosts\u201d), not a participle (\u201cboosting\u201d). (Lesson 2, Pattern 2)"),
    ("8",
     "Team sports serve as a powerful means of relieving stress, which enables students to recharge "
     "both physically and mentally.",
     "The pattern is \u201ca means OF + gerund\u201d. After the preposition \u201cof\u201d, use the "
     "\u2013ing form: \u201cof relieving\u201d, not \u201cof relieve\u201d. (Lesson 2, Pattern 3)"),
    ("9",
     "Admittedly, online lessons offer students greater flexibility. Nevertheless, they cannot replace "
     "face-to-face interaction, as real-time discussion sparks deeper engagement.",
     "Comma splice. \u201cNevertheless\u201d begins a new sentence, so the comma before it must become "
     "a full stop (or a semicolon) and \u201cNevertheless\u201d takes a capital letter. (Lesson 2, Pattern 4)"),
    ("10",
     "Reading serves as a powerful means of broadening students\u2019 horizons, which enables them to "
     "understand the world beyond the classroom.",
     "Agreement after \u201cwhich\u201d. \u201cWhich\u201d refers to the activity of reading (singular), "
     "so the verb is \u201cenables\u201d, not \u201cenable\u201d. (Lesson 2, Pattern 3)"),
]
t = mk_table(len(keys) + 1, (10, 80, 80))
hdr = t.rows[0].cells
for i, h in enumerate(("Item", "Corrected sentence", "The error and why (pattern)")):
    put(hdr[i], [(h, B, None)])
    shade(hdr[i], 'D9D9D9')
for i, (num, corr, why) in enumerate(keys, start=1):
    put(t.rows[i].cells[0], [(num, B, RED)])
    put(t.rows[i].cells[1], [(corr, PLAIN, RED)])
    put(t.rows[i].cells[2], [(why, PLAIN, RED)])
blank(1)

# ================================================================ PART B
heading("PART B: UPGRADING SENTENCES \u2014 MAKING SUGGESTIONS")
p("In the paragraph that suggests solutions, most students write \u201cSchools can \u2026\u201d, "
  "\u201cThe government can \u2026\u201d and \u201cParents can \u2026\u201d. These sentences are correct, "
  "but they are plain: they show the marker what you think, not how well you can write. Part B gives "
  "you THREE patterns that upgrade exactly these sentences. Each one is easy to remember, and each one "
  "carries one \u201cwow\u201d feature that impresses the marker.")
p([("Topic for this part: ", B, None),
   ("boost the physical fitness and awareness of citizens", PLAIN, None)], space_after=10)

t = box([[(u"LEVEL 3 \u2014 THE BASIC SENTENCES STUDENTS USUALLY WRITE", B, None)],
         [("The government can build more sports facilities.", PLAIN, None)],
         [("Schools can organise more sports activities and sports days.", PLAIN, None)],
         [("Parents can exercise with their children at weekends.", PLAIN, None)]],
        fill='D9D9D9')
blank(1)

# Pattern 1
t = box([[(u"Pattern 1 \u2014 Calling for Immediate Action", B, None)],
         [("Target Formula: ", B, None),
          ("It is high time that [doer] [PAST-TENSE verb phrase], so as to [purpose].", PLAIN, None)]])
blank(1)
label_value_table("Level 3:", [
    ("Level 3:", "The government can build more sports facilities."),
    ("Level 5:", [("It is high time that", B, None),
                  (" the government ", PLAIN, None),
                  ("invested", B, None),
                  (" in accessible sports facilities, ", PLAIN, None),
                  ("so as to", B, None),
                  (" make regular exercise possible for every citizen.", PLAIN, None)]),
])
p([("Teaching note: ", B, RED),
   ("\u201cIt is high time that\u201d is followed by the PAST tense \u2014 invested / launched / "
    "reviewed \u2014 never the present tense (\u201cinvests\u201d). If students panic, the easy "
    "alternative is \u201cIt is high time for the government TO invest \u2026\u201d. Keep the purpose "
    "clause as \u201cso as to + base verb\u201d.", PLAIN, RED)])
blank(1)

# Pattern 2
t = box([[(u"Pattern 2 \u2014 Purpose-First Recommendation", B, None)],
         [("Target Formula: ", B, None),
          ("To [goal], [doer] should make it a priority to [base verb phrase], thereby "
           "[gerund phrase \u2014 the result].", PLAIN, None)]])
blank(1)
label_value_table("Level 3:", [
    ("Level 3:", "Schools can organise more sports activities and sports days."),
    ("Level 5:", [("To boost students\u2019 physical fitness", B, None),
                  (", schools should ", PLAIN, None),
                  ("make it a priority to", B, None),
                  (" add a daily exercise break to the timetable, ", PLAIN, None),
                  ("thereby", B, None),
                  (" nurturing a habit of staying active.", PLAIN, None)]),
])
p([("Teaching note: ", B, RED),
   ("\u201cmake it a priority to\u201d takes the BASE verb (\u201cto add\u201d, never \u201cto "
    "adding\u201d). \u201cthereby + \u2013ing\u201d names the result and must follow logically from the "
    "action. Opening with \u201cTo + goal\u201d replaces the vague \u201cso that\u201d and shows the "
    "marker your purpose in the first three words.", PLAIN, RED)])
blank(1)

# Pattern 3
t = box([[(u"Pattern 3 \u2014 Shared Responsibility", B, None)],
         [("Target Formula: ", B, None),
          ("[Doer], in collaboration with [second party], should [verb phrase], thus "
           "[gerund phrase \u2014 the result].", PLAIN, None)]])
blank(1)
label_value_table("Level 3:", [
    ("Level 3:", "Parents can exercise with their children at weekends."),
    ("Level 5:", [("Parents, in collaboration with schools,", B, None),
                  (" should set a positive example by exercising with their children at weekends, "
                   , PLAIN, None),
                  ("thus", B, None),
                  (" fostering a culture of fitness at home.", PLAIN, None)]),
])
p([("Teaching note: ", B, RED),
   ("\u201cin collaboration with\u201d = working together with. It is far stronger than \u201cand\u201d "
    "or \u201calso\u201d and shows the marker that you understand shared responsibility. The doer is "
    "swappable: \u201cThe government, in collaboration with district councils, should \u2026\u201d. "
    "\u201cset a positive example by + \u2013ing\u201d adds the means; \u201cthus + \u2013ing\u201d "
    "closes with the result.", PLAIN, RED)])
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
    ("make students care about their health", "cultivate healthy habits",
     "Schools can cultivate healthy habits before bad ones take root."),
    ("help people keep fit", "enable citizens to keep fit",
     "Free fitness classes enable citizens to keep fit regardless of income."),
    ("fewer people do exercise now", "physical inactivity has become alarmingly common",
     "Physical inactivity has become alarmingly common among young office workers."),
]
t = mk_table(len(vocab) + 1, (8, 60, 102))
for i, h in enumerate(("#", "Level 3 (what students write)", "Level 5 upgrade (key words in bold)")):
    put(t.rows[0].cells[i], [(h, B, None)])
    shade(t.rows[0].cells[i], 'D9D9D9')
for i, (l3, l5, eg) in enumerate(vocab, start=1):
    put(t.rows[i].cells[0], [(str(i), B, None)])
    put(t.rows[i].cells[1], [(l3, PLAIN, None)])
    # bold the upgrade phrase, example on a second line
    put(t.rows[i].cells[2], [(l5, B, None)])
    par = t.rows[i].cells[2].add_paragraph()
    run = par.add_run("e.g. " + eg)
    run.font.italic = True
blank(1)

# ================================================================ PART C
heading("PART C: PRACTICE SETS (6 SETS)")
p("Upgrade each simple sentence with the pattern shown. The topics are different from the fitness "
  "topic in Part B. Sets 4\u20136 are more challenging: you must join the TWO simple sentences into "
  "ONE upgraded sentence.")

sets = [
    ("Set 1. Topic: Teenage Sleep Deprivation", "Use Pattern 1",
     ["The government can tell teenagers about the importance of sleep."],
     "It is high time that the government raised teenagers\u2019 awareness of the importance of "
     "sleep, so as to reverse the worrying trend of sleep deprivation among the young.",
     "Accept any answer that keeps the formula and the logic \u2014 e.g. \u201claunched a campaign to "
     "promote healthy sleeping habits, so as to \u2026\u201d."),
    ("Set 2. Topic: Food Waste on Campus", "Use Pattern 2",
     ["Schools can ask students not to waste food in the canteen."],
     "To cut food waste on campus, schools should make it a priority to offer smaller portions and "
     "free second helpings, thereby encouraging students to take only what they can finish.",
     "Accept \u201cTo reduce food waste \u2026\u201d and any result clause that follows logically "
     "(\u201cthereby cutting the amount of food sent to landfill\u201d)."),
    ("Set 3. Topic: Cyberbullying", "Use Pattern 3",
     ["Parents can talk to their children about cyberbullying."],
     "Parents, in collaboration with schools, should talk openly with their children about "
     "cyberbullying, thus building the trust that encourages victims to speak up.",
     "Accept any sensible second party (the school, social workers) and any result clause that "
     "follows logically (\u201cthus creating a safe space where problems can be reported\u201d)."),
    ("Set 4. Topic: Youth Mental Health (more challenging)", "Use Pattern 1",
     ["The government can spend more money on youth mental health services.",
      "Schools can also teach students how to manage stress."],
     "It is high time that the government invested more in youth mental health services, so as to "
     "ensure that schools can teach students how to manage stress.",
     "Require BOTH ideas in ONE sentence. Also accept \u201c\u2026 so as to make stress-management "
     "lessons possible in every school.\u201d"),
    ("Set 5. Topic: Declining Reading Habits (more challenging)", "Use Pattern 2",
     ["Schools can keep the library open after school.",
      "Students will then read more."],
     "To revive students\u2019 interest in reading, schools should make it a priority to keep the "
     "library open after school, thereby giving students a quiet place to read every day.",
     "Require the goal at the front and the result after \u201cthereby\u201d. A plain second sentence "
     "(\u201c\u2026. Students will then read more.\u201d) scores nothing."),
    ("Set 6. Topic: Care for the Elderly (more challenging)", "Use Pattern 3",
     ["Teenagers can visit elderly neighbours with community centres.",
      "The elderly will feel less lonely."],
     "Teenagers, in collaboration with community centres, should organise regular visits to elderly "
     "neighbours, thus easing the loneliness that many of them face.",
     "Watch the plural: \u201cthe elderly\u201d is already plural \u2014 never \u201cthe "
     "elderlies\u201d. Accept any result clause that follows logically from the visits."),
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
t = box([[("HKDSE Adapted Paper 2 Prompt:", PLAIN, None)],
         [("You are a Form 6 student. Your school is considering cutting one of the two weekly PE "
           "lessons so that F.6 students can have an extra period for examination preparation. Write "
           "a letter to the Principal expressing your opinion on this idea.", PLAIN, None)]],
        fill='F2F2F2')
p([("Task requirements:", B, None)], space_before=8, space_after=4)
for line in [
    "Write a letter of about 250 words.",
    "Use at least TWO of today\u2019s three patterns, and underline every target structure you use.",
    "Use the letter format: \u201cDear Principal,\u201d at the start; \u201cYours sincerely,\u201d and "
    "your name and class at the end.",
    "Decide your stance first: you may oppose the change, or support it on one condition.",
]:
    p(line, space_after=2, indent=4)

p([("Teacher\u2019s planning guide", B, RED)], space_before=10, space_after=4)
label_value_table("x", [
    ("Introduction", "The issue + your stance. Use a Lesson 1 pattern (e.g. \u201cIn recent times, "
                     "the issue of \u2026 has sparked widespread concern across society.\u201d)."),
    ("Body 1", "Why the PE lesson matters. A Lesson 2 pattern fits here (\u201c\u2026 offers students "
               "a golden opportunity to \u2026, paving the way for \u2026\u201d), followed by an "
               "Admittedly / Nevertheless counter-argument."),
    ("Body 2", "Your alternative solution \u2014 this is where today\u2019s three patterns do the "
               "work. Give the school something it can actually do."),
    ("Conclusion", "One sentence of recommendation, no new ideas."),
])

p([("Teacher\u2019s model answer (opposing the change, 239 words)", B, None)],
  space_before=10, space_after=4)
model_a = [
    ("Dear Principal,",
     "Salutation"),
    ("In recent times, the issue of students\u2019 physical well-being has sparked widespread concern "
     "across society. I am writing to express my opposition to the proposal of cutting one of our two "
     "weekly PE lessons in favour of an extra examination preparation period.",
     "Introduction \u2014 Lesson 1, Pattern 3 + stance"),
    ("Undoubtedly, F.6 students would welcome more time for revision. Nevertheless, PE lessons offer "
     "us a golden opportunity to recharge after long hours of study, paving the way for sharper "
     "concentration in class. Admittedly, an extra period of examination practice sounds attractive. "
     "Nevertheless, replacing physical activity with yet more revision is short-sighted, as a tired "
     "mind absorbs far less than a refreshed one.",
     "Body 1 \u2014 Lesson 2, Pattern 1 + Pattern 4 (concede, then rebut)"),
    ("Instead of sacrificing our PE lesson, may I propose two practical alternatives? To free up time "
     "for examination preparation, the school should make it a priority to shorten the morning "
     "assembly and the daily roll-call, thereby recovering a full period every week. It is also high "
     "time that the school reviewed the examination timetable, so as to spread the pressure of mock "
     "papers more evenly across the year. Parents, in collaboration with the school, should also "
     "encourage their children to spend half an hour on simple exercise at home, thus turning fitness "
     "into a daily habit rather than a single weekly lesson.",
     "Body 2 \u2014 today\u2019s Patterns 2, 1 and 3 in one paragraph"),
    ("Physical fitness and academic performance are not rivals. Please reconsider this proposal, so "
     "that F.6 students can prepare for the public examination without damaging their health.",
     "Conclusion \u2014 recommendation"),
    ("Yours sincerely,\nChris Wong (6B)",
     "Closing"),
]
for text, label in model_a:
    p([(text, PLAIN, None), ("  [" + label + "]", False, RED)], space_after=6)

p([("Teacher\u2019s model answer B (supporting the change, on one condition \u2014 213 words)",
    B, None)], space_before=10, space_after=4)
model_b = [
    ("Dear Principal,",
     "Salutation"),
    ("In recent times, the issue of examination pressure on F.6 students has sparked widespread "
     "concern across society. I am writing to support the proposal, on the condition that the school "
     "protects our health at the same time.",
     "Introduction \u2014 Lesson 1, Pattern 3 + stance"),
    ("Admittedly, cutting a PE lesson reduces the time we spend exercising. Nevertheless, the final "
     "year of study demands every minute of preparation, as the competition for university places has "
     "never been keener. From the perspective of F.6 students, an extra examination preparation period "
     "allows us to practise past papers under a teacher\u2019s supervision, which in turn reduces the "
     "anxiety of revising alone at home.",
     "Body 1 \u2014 Lesson 2, Pattern 4 + Pattern 2"),
    ("To make up for the lost lesson, it is high time that the school offered a supervised morning "
     "workout before classes, so as to keep our bodies active without touching the teaching "
     "timetable. Teachers should also make it a priority to set homework that leaves room for "
     "exercise, thereby preventing revision from swallowing the whole day, and parents, in "
     "collaboration with the school, should monitor our daily routine, thus ensuring that an extra "
     "study period does not become an extra hour in front of a screen.",
     "Body 2 \u2014 today\u2019s Patterns 1, 2 and 3 in one paragraph"),
    ("I hope the school will consider this compromise, which serves the examination and our health at "
     "the same time.",
     "Conclusion \u2014 recommendation"),
    ("Yours sincerely,\nChris Wong (6B)",
     "Closing"),
]
for text, label in model_b:
    p([(text, PLAIN, None), ("  [" + label + "]", False, RED)], space_after=6)

p([("What to reward when marking", B, RED)], space_before=10, space_after=4)
for line in [
    "Content: a clear stance on the cut itself, plus at least one concrete, realistic alternative "
    "(timetable adjustment, morning workout) \u2014 not just \u201cPE is good for us\u201d.",
    "Language: the target structures used accurately; formal verbs instead of \u201chelp\u201d and "
    "\u201cmake\u201d (establish, invest in, cultivate, combat); no \u201cI think it is good / "
    "bad\u201d.",
    "Organisation: letter format complete (Dear Principal / Yours sincerely / name and class), one "
    "main idea per paragraph, and linkers between paragraphs.",
]:
    p(line, space_after=2, indent=4)

p([("Common errors to watch when marking", B, RED)], space_before=8, space_after=4)
for line in [
    "\u201cIt is high time that the government promote \u2026\u201d \u2192 the past tense is needed "
    "(\u201cpromoted\u201d).",
    "\u201cmake it a priority to + \u2013ing\u201d \u2192 base verb only (\u201cto add\u201d).",
    "\u201cthus / thereby + verb\u201d \u2192 the \u2013ing form is needed (\u201cthus fostering\u201d, "
    "\u201cthereby nurturing\u201d).",
    "\u201cso as to + \u2013ing\u201d \u2192 base verb (\u201cso as to make\u201d).",
    "\u201cin collaboration with\u201d used as a verb (\u201cparents collaborate with schools "
    "should\u2026\u201d) \u2192 keep it as a phrase between the subject and \u201cshould\u201d.",
    "\u201cthe elderly\u201d treated as singular (\u201cthe elderlies\u201d, \u201cthe elderly feels\u201d).",
]:
    p(line, space_after=2, indent=4)

doc.save(OUT)
print('saved', OUT)

# ---------------------------------------------------------------- verify
d = Document(OUT)
print('tables:', len(d.tables), 'paragraphs:', len(d.paragraphs))
words = lambda s: len(re.findall(r"[A-Za-z\u2019'-]+", s))
for name, letters in (('model A', model_a), ('model B', model_b)):
    body = ' '.join(t for t, _ in letters)
    print(name, 'words:', words(body))
