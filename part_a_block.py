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

