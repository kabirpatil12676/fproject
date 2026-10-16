"""
ELEVATE Project Synopsis – Word Document Generator
Produces a clean, black-and-white BTech synopsis .docx
"""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

from docx import Document
from docx.shared import Pt, Inches, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

# ── helpers ─────────────────────────────────────────────────────────────────

def set_font(run, name='Times New Roman', size=12, bold=False, italic=False, underline=False):
    run.font.name = name
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.underline = underline
    run.font.color.rgb = RGBColor(0, 0, 0)

def set_para_spacing(para, before=0, after=6, line=None):
    pf = para.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    if line:
        pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        pf.line_spacing = Pt(line)

def add_para(doc, text='', style='Normal', bold=False, italic=False, size=12,
             align=WD_ALIGN_PARAGRAPH.JUSTIFY, before=0, after=6, indent=None):
    p = doc.add_paragraph(style=style)
    p.alignment = align
    set_para_spacing(p, before=before, after=after)
    if indent is not None:
        p.paragraph_format.first_line_indent = Cm(indent)
    if text:
        r = p.add_run(text)
        set_font(r, size=size, bold=bold, italic=italic)
    return p

def add_heading(doc, text, level=1):
    """Black, Times New Roman headings. Section heading = bold, underlined."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    set_para_spacing(p, before=12, after=4)
    r = p.add_run(text)
    if level == 1:
        set_font(r, size=14, bold=True)
    elif level == 2:
        set_font(r, size=12, bold=True)
    else:
        set_font(r, size=12, bold=False, italic=True)
    return p

def add_body(doc, text, first_indent=True, before=0, after=4):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    set_para_spacing(p, before=before, after=after)
    if first_indent:
        p.paragraph_format.first_line_indent = Cm(0.75)
    r = p.add_run(text)
    set_font(r, size=12)
    return p

def add_body_mixed(doc, parts, first_indent=True, before=0, after=4):
    """parts = list of (text, bold, italic)"""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    set_para_spacing(p, before=before, after=after)
    if first_indent:
        p.paragraph_format.first_line_indent = Cm(0.75)
    for text, bold, italic in parts:
        r = p.add_run(text)
        set_font(r, size=12, bold=bold, italic=italic)
    return p

def add_bullet(doc, label, body, label_bold=True):
    p = doc.add_paragraph(style='Normal')
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    set_para_spacing(p, before=0, after=4)
    p.paragraph_format.left_indent = Cm(0.75)
    p.paragraph_format.first_line_indent = Cm(-0.75)
    r1 = p.add_run(u'\u2022  ')
    set_font(r1, size=12)
    if label:
        r2 = p.add_run(label + ': ')
        set_font(r2, size=12, bold=label_bold)
    r3 = p.add_run(body)
    set_font(r3, size=12)
    return p

def add_bullet_simple(doc, text, indent_cm=0.75):
    p = doc.add_paragraph(style='Normal')
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    set_para_spacing(p, before=0, after=3)
    p.paragraph_format.left_indent = Cm(indent_cm)
    p.paragraph_format.first_line_indent = Cm(-indent_cm)
    r = p.add_run(u'\u2022  ' + text)
    set_font(r, size=12)
    return p

def add_number_bullet(doc, number, text_parts, before=0, after=4):
    """Numbered point: (1)  bold_text: rest_text"""
    p = doc.add_paragraph(style='Normal')
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    set_para_spacing(p, before=before, after=after)
    p.paragraph_format.left_indent = Cm(1.0)
    p.paragraph_format.first_line_indent = Cm(-1.0)
    for text, bold, italic in text_parts:
        r = p.add_run(text)
        set_font(r, size=12, bold=bold, italic=italic)
    return p

def set_cell_border(cell, top=None, bottom=None, left=None, right=None):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for side, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        if val:
            el = OxmlElement(f'w:{side}')
            el.set(qn('w:val'), val)
            el.set(qn('w:sz'), '4')
            el.set(qn('w:space'), '0')
            el.set(qn('w:color'), '000000')
            tcBorders.append(el)
    tcPr.append(tcBorders)

def shade_cell(cell, fill='D9D9D9'):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill)
    tcPr.append(shd)

def cell_text(cell, text, bold=False, italic=False, size=10, align=WD_ALIGN_PARAGRAPH.LEFT):
    cell.text = ''
    p = cell.paragraphs[0]
    p.alignment = align
    set_para_spacing(p, before=2, after=2)
    r = p.add_run(text)
    set_font(r, size=size, bold=bold, italic=italic)

def cell_mixed(cell, parts, size=10, align=WD_ALIGN_PARAGRAPH.LEFT):
    cell.text = ''
    p = cell.paragraphs[0]
    p.alignment = align
    set_para_spacing(p, before=2, after=2)
    for text, bold, italic in parts:
        r = p.add_run(text)
        set_font(r, size=size, bold=bold, italic=italic)

def add_page_break(doc):
    doc.add_page_break()

def set_col_width(table, col_idx, width_cm):
    for row in table.rows:
        row.cells[col_idx].width = Cm(width_cm)

# ── document setup ───────────────────────────────────────────────────────────

doc = Document()

# Page margins (A4)
sec = doc.sections[0]
sec.page_width  = Cm(21.0)
sec.page_height = Cm(29.7)
sec.top_margin    = Cm(2.5)
sec.bottom_margin = Cm(2.5)
sec.left_margin   = Cm(3.0)
sec.right_margin  = Cm(2.0)

# Default style
style = doc.styles['Normal']
style.font.name = 'Times New Roman'
style.font.size = Pt(12)
style.font.color.rgb = RGBColor(0, 0, 0)

# ══════════════════════════════════════════════════════════════════════════════
# COVER PAGE
# ══════════════════════════════════════════════════════════════════════════════

def centered(doc, text, size=12, bold=False, italic=False, before=0, after=6):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_para_spacing(p, before=before, after=after)
    r = p.add_run(text)
    set_font(r, size=size, bold=bold, italic=italic)
    return p

centered(doc, 'DEPARTMENT OF INFORMATION TECHNOLOGY', size=13, bold=True, before=0, after=2)
centered(doc, 'WALCHAND INSTITUTE OF TECHNOLOGY, SOLAPUR', size=14, bold=True, after=2)
centered(doc, 'An Autonomous Institute | Affiliated to Solapur University', size=10, italic=True, after=18)

# Horizontal rule – paragraph border
def add_hrule(doc):
    p = doc.add_paragraph()
    set_para_spacing(p, before=0, after=0)
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '8')
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), '000000')
    pBdr.append(bottom)
    pPr.append(pBdr)
    return p

add_hrule(doc)
add_para(doc, before=6)  # spacing

centered(doc, 'A PROJECT SYNOPSIS', size=13, bold=True, before=6, after=2)
centered(doc, 'Submitted in partial fulfillment of the requirements for the degree of', size=11, italic=True, after=2)
centered(doc, 'Bachelor of Technology in Information Technology', size=12, bold=True, after=18)

centered(doc, 'Design and Evaluation of an Educator-Led Experiential', size=13, bold=True, before=0, after=2)
centered(doc, 'Pedagogical Model and Personalized Digital Platform', size=13, bold=True, after=2)
centered(doc, 'for Holistic Student Growth', size=13, bold=True, after=18)

add_hrule(doc)
add_para(doc, before=6)

# Team table
centered(doc, 'TEAM DETAILS   -   GROUP NO: B-07', size=11, bold=True, before=8, after=8)

tbl = doc.add_table(rows=5, cols=3)
tbl.style = 'Table Grid'
tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
# Header
hdr = tbl.rows[0].cells
cell_text(hdr[0], 'Sr.', bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
cell_text(hdr[1], 'Name of Student', bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
cell_text(hdr[2], 'Roll No.', bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
for c in hdr: shade_cell(c, 'D9D9D9')
data = [('1.', 'Kabir Shitalkumar Patil', 'B-26'),
        ('2.', 'Alok Shriniwas Perla',    'B-34'),
        ('3.', 'Vidhi Jatinbhai Patel',   'B-23'),
        ('4.', 'Rudraksh Rahul Vaidya',   'B-61')]
for i, (sr, name, roll) in enumerate(data, 1):
    row = tbl.rows[i].cells
    cell_text(row[0], sr, align=WD_ALIGN_PARAGRAPH.CENTER)
    cell_text(row[1], name)
    cell_text(row[2], roll, align=WD_ALIGN_PARAGRAPH.CENTER)
# Column widths
for i, w in enumerate([1.5, 8.5, 2.5]):
    set_col_width(tbl, i, w)

add_para(doc, before=14)
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
set_para_spacing(p, before=0, after=2)
r = p.add_run('Project Guide: '); set_font(r, size=11, bold=True)
r = p.add_run('Dr. Ms. D. D. Awasekar'); set_font(r, size=11)
r = p.add_run('     |     '); set_font(r, size=11)
r = p.add_run('Academic Year: '); set_font(r, size=11, bold=True)
r = p.add_run('2025–2026'); set_font(r, size=11)

add_page_break(doc)

# ══════════════════════════════════════════════════════════════════════════════
# 1. INTRODUCTION
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, '1.  Introduction & Background')

add_body_mixed(doc, [
    ('The rapid expansion of digital infrastructure has fundamentally transformed the delivery of education and professional development. Modern learners expect ', False, False),
    ('flexible, personalized, and outcome-focused', True, False),
    (' learning experiences that extend well beyond the passive consumption of recorded lectures. However, a persistent gap exists in the learning technology ecosystem: most platforms are either large-scale multi-vendor marketplaces optimized for content discovery, or heavyweight enterprise ', False, False),
    ('Learning Management Systems (LMS)', True, False),
    (' that replicate the passive, content-centric model of traditional classrooms in digital form.', False, False),
])

add_body_mixed(doc, [
    ('This project addresses that gap by proposing both a new ', False, False),
    ('pedagogical model', True, False),
    (' and the ', False, False),
    ('digital platform', True, False),
    (' that implements it. The pedagogical model - ', False, False),
    ('ELEVATE (Educator-Led Experiential Value and Life-skills Advancement Through Engagement)', True, True),
    (' - is a six-stage structured learning framework grounded in ', False, False),
    ('experiential learning', True, False),
    (', ', False, False),
    ('mastery learning', True, False),
    (', and ', False, False),
    ('reflective practice', True, False),
    ('. The platform serves as the implementation vehicle, enabling a single organizational content authority to design, deliver, and evaluate skill-development journeys for its learners.', False, False),
])

add_body_mixed(doc, [
    ('The central departure from conventional LMS design is the ', False, False),
    ('unit of learning', True, False),
    ('. Traditional platforms organize content into courses composed of video lessons and quizzes; the ELEVATE platform organizes content into ', False, False),
    ('Growth Journeys', True, False),
    (' - structured sequences moving the learner from diagnostic self-assessment through educator-led instruction, experiential activity, multi-source feedback, authentic application, and ', False, False),
    ('reflective portfolio construction', True, False),
    ('. The output is not a completion certificate but a ', False, False),
    ('documented, measurable evidence portfolio', True, False),
    (' of skill development. The platform is ', False, False),
    ('domain-agnostic', True, False),
    (': the underlying data model supports academic instruction, professional certification, technical upskilling, and general life-skill development.', False, False),
])

# Figure 1.1 – keep the comparison image (this is informative, no repetitive text beside it)
doc.add_picture('ppt_images/image3.png', width=Cm(13))
p = doc.paragraphs[-1]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
cap = doc.add_paragraph()
cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
set_para_spacing(cap, before=2, after=8)
r = cap.add_run('Figure 1: '); set_font(r, size=10, bold=True)
r = cap.add_run('Comparison between traditional digital learning platforms and the proposed ELEVATE experiential learning platform.'); set_font(r, size=10, italic=True)

# ══════════════════════════════════════════════════════════════════════════════
# 2. PROBLEM STATEMENT
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, '2.  Problem Statement')

add_body_mixed(doc, [
    ('Engineering and professional education has long recognized a structural deficiency: the curriculum delivers ', False, False),
    ('technical knowledge', True, False),
    (' but rarely develops the ', False, False),
    ('life skills, communication competencies', True, False),
    (', and professional capabilities that determine real-world success. Employer surveys consistently indicate that graduates are technically prepared but deficient in communication, leadership, critical thinking, and professional presence. Existing learning platforms fail to address this gap for the following reasons:', False, False),
])

problems = [
    ('Passive Content-Centric Learning', 'Learning is limited to videos and quizzes with minimal practical application. There is no mechanism for learners to practise, receive structured feedback, or demonstrate applied competence.'),
    ('Lack of Structured Experiential Learning', 'Students have limited opportunities to learn through guided activities, reflection, and hands-on experiences. No mainstream platform provides a structured activity-challenge-feedback cycle.'),
    ('Limited Personalized Feedback', 'Most platforms provide generic assessments without continuous mentoring or individualized guidance. Educator insight is absent from the learning journey in most self-paced platforms.'),
    ('Absence of Evidence-Based Skill Assessment', 'Course completion certificates do not demonstrate practical competencies. There is no platform-level mechanism to measure whether a learner has improved, and by how much, across a defined skill dimension.'),
    ('Poor Support for Holistic Student Development', 'Technical knowledge is prioritized while communication, leadership, teamwork, and life skills receive limited attention. The prevalent model is binary - a learner has either completed a course or not - providing no information about the depth, quality, or trajectory of learning.'),
]
for i, (label, body) in enumerate(problems, 1):
    p = doc.add_paragraph(style='Normal')
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    set_para_spacing(p, before=0, after=4)
    p.paragraph_format.left_indent = Cm(1.0)
    p.paragraph_format.first_line_indent = Cm(-1.0)
    r = p.add_run(f'{i}.  '); set_font(r, size=12)
    r = p.add_run(label + ': '); set_font(r, size=12, bold=True)
    r = p.add_run(body); set_font(r, size=12)

add_body_mixed(doc, [
    ('The problem is therefore not merely technological but ', False, False),
    ('pedagogical', True, False),
    (': there is a need for a platform built on an educationally principled model that moves beyond content delivery to ', False, False),
    ('structured, measurable, experiential growth', True, False),
    ('.', False, False),
], first_indent=True, before=4, after=4)

add_heading(doc, 'Project Aim', level=2)
add_body_mixed(doc, [
    ('To design and evaluate an educator-led experiential pedagogical model and personalized digital platform that supports life-skill development, professional readiness, and holistic student growth beyond conventional classroom learning.', False, False),
], first_indent=False, before=0, after=6)

# Figure 2.1 – challenges visual (IMAGE ONLY – text already explains the 5 challenges above)


# ══════════════════════════════════════════════════════════════════════════════
# 3. OBJECTIVES
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, '3.  Objectives')

add_body(doc, 'dummy')  # placeholder – replace with actual
doc.paragraphs[-1]._element.getparent().remove(doc.paragraphs[-1]._element)

add_body(doc, 'dummy')
doc.paragraphs[-1]._element.getparent().remove(doc.paragraphs[-1]._element)

# Objectives image (PPT-approved visual)


add_para(doc, 'The objectives, formally stated, are as follows:', indent=0, before=4, after=4)

objectives = [
    ('To identify students\' learning needs and skill gaps for holistic development.',
     'Identify students\' preferences, challenges, and skill gaps through surveys and analysis to understand their learning needs and support their overall growth.'),
    ('To design an educator-led experiential pedagogical model integrating learning, practice, feedback, reflection, and real-world application.',
     'Design a structured pedagogical model promoting experiential learning through a continuous cycle - formalized as the six-stage ELEVATE Framework (Explore → Learn → Experience → Validate → Apply → Transform).'),
    ('To develop a personalized digital platform for delivering masterclasses, learning activities, assessments, and individualized learning pathways.',
     'Develop a user-friendly platform offering masterclasses, guided activities, assessments, and personalized pathways, including a multi-source feedback pipeline (educator rubric + peer review) and a Growth Portfolio subsystem.'),
    ('To evaluate the effectiveness of the proposed model and platform in terms of student engagement, skill development, perceived value, and overall learning experience.',
     'Evaluate the impact through user feedback, engagement metrics, and skill improvement, applying the Know → Do → Become measurement framework for pre- and post-journey outcome assessment.'),
]
for i, (label, body) in enumerate(objectives, 1):
    p = doc.add_paragraph(style='Normal')
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    set_para_spacing(p, before=0, after=5)
    p.paragraph_format.left_indent = Cm(1.0)
    p.paragraph_format.first_line_indent = Cm(-1.0)
    r = p.add_run(f'{i}.  '); set_font(r, size=12)
    r = p.add_run(label + ' '); set_font(r, size=12, bold=True)
    r = p.add_run(body); set_font(r, size=12)

# ══════════════════════════════════════════════════════════════════════════════
# 4. STUDENT LEARNING SURVEY
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, '4.  Student Learning Survey & Needs Analysis')

add_body_mixed(doc, [
    ('A structured survey was conducted among ', False, False),
    ('39 student respondents', True, False),
    (' to identify learning preferences, skill gaps, and expectations from a personalized, educator-led learning platform. The findings directly informed the design of the ELEVATE model and platform features. Key findings are summarized below:', False, False),
])

# Survey image (charts – can't be replicated in text)


# Key findings as numbered bullets (concise, complementary to chart)
findings = [
    ('Confidence/public speaking (35.9%)', ' and ', 'Communication (28.2%)', ' are the top non-technical skill challenges faced by students.'),
    ('Short practical videos (66.7%)', ' and doubt-solving communities (41%) are the preferred learning formats most likely to lead to course completion.'),
    ('Most students feel unprepared for ', 'Coding/DSA (74.4%)', ' and Technical Interviews (46.2%), followed by Resume Building (41%).'),
    ('Students are most curious about ', 'AI tools for productivity (53.8%)', ' and getting into product-based companies (48.7%) but lack proper guidance.'),
]
# Simplified as plain bullets:
survey_bullets = [
    'Confidence/public speaking (35.9%) and Communication (28.2%) are the top non-technical skill challenges.',
    'Students prefer short practical videos (66.7%) and doubt-solving communities (41%) to complete courses.',
    'Most students feel unprepared for Coding/DSA (74.4%) and Technical Interviews (46.2%), followed by Resume Building (41%).',
    'Students are most curious about AI tools for productivity (53.8%) and getting into product-based companies (48.7%) but lack guidance.',
]
for b in survey_bullets:
    add_bullet_simple(doc, b)

add_body_mixed(doc, [
    ('These findings validate the need for a structured, educator-led experiential platform addressing ', False, False),
    ('life-skill development', True, False),
    (', ', False, False),
    ('technical preparedness', True, False),
    (', and ', False, False),
    ('personalized guidance', True, False),
    (' - precisely the problem space ELEVATE is designed to solve.', False, False),
], first_indent=True, before=4, after=6)

# ══════════════════════════════════════════════════════════════════════════════
# 5. LITERATURE SURVEY
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, '5.  Literature Survey')
add_heading(doc, '5.1  Existing Learning Platform Paradigms', level=2)

add_body(doc, 'dummy'); doc.paragraphs[-1]._element.getparent().remove(doc.paragraphs[-1]._element)

add_body_mixed(doc, [
    ('Learning delivery systems in current use can be classified into two principal categories, neither of which satisfactorily addresses the pedagogical requirements of this project:', False, False),
], first_indent=False, before=0, after=4)

add_bullet(doc, 'Multi-Vendor Marketplaces (Udemy, Coursera, edX)',
    'These platforms aggregate content from many independent instructors, optimized for content discovery. They cannot support an educator-led, activity-centered, feedback-driven pedagogy; content is linear video sequences with optional quizzes. There is no mechanism for submission, rubric-based feedback, or Growth Portfolio construction [7][8].')
add_bullet(doc, 'Single-Owner LMS Platforms (Moodle, TalentLMS, Docebo)',
    'These platforms support a single content authority managing its own learner population but replicate the passive content model. They do not implement a structured experiential pedagogy, do not provide a multi-source educator feedback pipeline, and do not produce growth portfolios as learning outputs [4][5][6].')

add_heading(doc, '5.2  Structured Review of Related Literature', level=2)

add_body_mixed(doc, [
    ('The following table presents a structured review of key foundational works, their strengths, limitations, identified research gaps, and how ELEVATE addresses each gap (IEEE format):', False, False),
], first_indent=False, before=0, after=6)

# Literature Survey Table (7 rows + header)
lit_cols = ['#', 'Research Paper / Author (Year)', 'Strength', 'Limitations', 'Research Gap', 'How ELEVATE Addresses the Gap']
lit_widths = [0.8, 3.3, 2.5, 2.5, 2.4, 3.0]

lit_data = [
    ('[1]', 'Kolb, D. A. (1984)\nExperiential Learning: Experience as the Source of Learning and Development',
     'Introduced the experiential learning cycle (Concrete Experience → Reflective Observation → Abstract Conceptualization → Active Experimentation).',
     'No digital implementation. Does not address tools for tracking learner progress and competency development.',
     'Need for a digital platform that operationalizes the experiential learning cycle with structured activities, reflection, and feedback.',
     'ELEVATE implements Growth Journeys with experiential activities, reflection journals, educator feedback, and competency tracking.'),
    ('[2]', 'Bloom, B. S. (1956)\nTaxonomy of Educational Objectives: Handbook I – Cognitive Domain',
     'Provides a hierarchical framework for classifying learning objectives in the cognitive domain.',
     'Focuses mainly on the cognitive domain. Does not address affective, psychomotor, or holistic skill development.',
     'Need for an approach that addresses holistic development across technical, professional, and life skills.',
     'ELEVATE integrates technical, professional, and life skills in learning pathways targeting multiple skill domains.'),
    ('[3]', 'Schön, D. A. (1983)\nThe Reflective Practitioner: How Professionals Think in Action',
     'Highlights the importance of reflection-in-action and reflection-on-action for professional growth.',
     'Does not provide digital tools or mechanisms to facilitate and monitor reflection systematically.',
     'Need for digital support for guided reflection and feedback within learning platforms.',
     'ELEVATE provides reflection journals, guided prompts, and educator feedback, making reflection a continuous and measurable part of learning (Transform stage).'),
    ('[4]', 'Dabbagh & Kitsantas (2012)\nPersonal Learning Environments, Social Media, and Self-Regulated Learning. The Internet and Higher Education, vol. 15, no. 1, pp. 3–8',
     'Emphasizes self-regulated learning and the role of personal learning environments in fostering learner autonomy.',
     'Limited focus on structured educator involvement. Lacks mechanisms for continuous feedback and personalized guidance.',
     'Need for an educator-led approach combining self-regulation with personalized guidance and continuous feedback.',
     'ELEVATE combines self-regulated learning with active educator mentorship, feedback, and personalized learning pathways.'),
    ('[5]', 'Siemens, G. (2005)\nConnectivism: A Learning Theory for the Digital Age. Int\'l Journal of Instructional Technology & Distance Learning, vol. 2, no. 1',
     'Highlights the importance of networks, connections, and digital resources for learning in the digital age.',
     'Limited emphasis on competency assessment and measurable learning outcomes.',
     'Need for measurable skill development and evidence-based assessment in networked learning environments.',
     'ELEVATE uses competency-based assessments, real-world activities, and growth portfolios to track measurable skill development.'),
    ('[6]', 'Hattie & Timperley (2007)\nThe Power of Feedback. Review of Educational Research, vol. 77, no. 1, pp. 81–112',
     'Establishes that timely, specific feedback is among the most impactful teaching interventions.',
     'Focuses on theoretical models; does not address automated feedback mechanisms.',
     'Need for scalable, real-time, structured feedback mechanisms that deliver educator-quality feedback at scale.',
     "ELEVATE's multi-source feedback pipeline (educator rubric + peer review) operationalizes the power of feedback at platform scale."),
    ('[7]', 'Ifenthaler & Yau (2020)\nUtilizing Learning Analytics to Support Study Success in Higher Education. ETR&D, vol. 68, no. 4, pp. 1961–1973',
     'Demonstrates that learning analytics significantly improve student outcomes when used for personalized support and early intervention.',
     'Focuses on institutional-level analytics; learner-facing dashboards and individual growth tracking remain underexplored.',
     'Need for learner-facing growth analytics dashboards embedded within the learning journey itself.',
     "ELEVATE's Progress Dashboard and Growth Portfolio give learners real-time visibility into their learning trajectory and skill development evidence."),
]

lit_tbl = doc.add_table(rows=len(lit_data)+1, cols=6)
lit_tbl.style = 'Table Grid'
lit_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER

# Header row
for j, (hdr_txt, w) in enumerate(zip(lit_cols, lit_widths)):
    c = lit_tbl.rows[0].cells[j]
    shade_cell(c, 'D9D9D9')
    cell_text(c, hdr_txt, bold=True, size=9, align=WD_ALIGN_PARAGRAPH.CENTER)
    c.width = Cm(w)

# Data rows
for i, row_data in enumerate(lit_data, 1):
    for j, txt in enumerate(row_data):
        c = lit_tbl.rows[i].cells[j]
        # First col (ref number) – plain, centered
        if j == 0:
            cell_text(c, txt, size=9, align=WD_ALIGN_PARAGRAPH.CENTER)
        elif j == 1:
            # Author / title – bold author, italic title
            lines = txt.split('\n', 1)
            c.text = ''
            p = c.paragraphs[0]
            set_para_spacing(p, before=2, after=2)
            r = p.add_run(lines[0]); set_font(r, size=9, bold=True)
            if len(lines) > 1:
                p.add_run('\n')
                r2 = p.add_run(lines[1]); set_font(r2, size=9, italic=True)
        else:
            cell_text(c, txt, size=9)
        c.width = Cm(lit_widths[j])

cap = doc.add_paragraph()
cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
set_para_spacing(cap, before=3, after=8)
r = cap.add_run('Table 1: '); set_font(r, size=10, bold=True)
r = cap.add_run('Structured Literature Review - Research Gap Analysis and ELEVATE\'s Proposed Contributions (IEEE Format).'); set_font(r, size=10, italic=True)

add_heading(doc, '5.3  Identified Research Gap', level=2)

add_body_mixed(doc, [
    ('The academic literature and commercial product landscape together reveal a consistent gap: no existing platform combines (a) an ', False, False),
    ('educator-led, six-stage experiential pedagogical model', True, False),
    (', (b) ', False, False),
    ('multi-source feedback', True, False),
    (' on practice submissions, and (c) a ', False, False),
    ('Growth Portfolio', True, False),
    (' as the primary learning artifact, within a single organizational platform designed for life-skill and professional development. This project directly addresses that gap.', False, False),
])

# ══════════════════════════════════════════════════════════════════════════════
# 6. PROPOSED ARCHITECTURE & METHODOLOGY
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, '6.  Proposed Architecture & Methodology')
add_heading(doc, '6.1  The ELEVATE Pedagogical Model - Six-Stage Framework', level=2)

add_body_mixed(doc, [
    ('ELEVATE', True, True),
    (' structures every Growth Journey into ', False, False),
    ('six sequential pedagogical stages', True, False),
    ('. Unlike a course-and-lesson hierarchy, a Growth Journey is a unit of measurable skill development with a defined entry state (Learner Growth Profile from the diagnostic), a structured progression through six stages, and a defined exit artifact (the Growth Portfolio). The six stages are:', False, False),
])

stages = [
    ('EXPLORE - Diagnostic Assessment',
     'Before beginning a Growth Journey, the learner completes a structured diagnostic assessment that establishes the Learner Growth Profile - capturing current competency level, prior experience, hesitation areas, and personal goals - providing a measurable baseline for pre/post comparison.'),
    ('LEARN - Educator-Led Learning Capsules',
     'Rather than passive hour-long video lectures, the Learn stage delivers focused 8–15 minute educator-led learning capsules: concept demonstrations, experience-based stories, practical frameworks, and case situations, authored and curated by the organizational educator.'),
    ('EXPERIENCE - Activity Challenge',
     'Every learning capsule is followed by an activity challenge requiring the learner to perform the skill, not merely understand it. Submissions may be video recordings, written documents, or structured responses depending on the skill domain.'),
    ('VALIDATE - Multi-Source Feedback',
     'Submitted activities are assessed through a multi-source feedback pipeline: educator rubric scoring, and optional peer feedback. Learners may iterate on their submission before advancing, consistent with the mastery learning principle.'),
    ('APPLY - Authentic Task',
     'The learner transitions from practice activity to an authentic, real-world application. This stage closes the gap between classroom practice and real-world demonstration (e.g., updating a professional profile after a personal branding learning capsule).'),
    ('TRANSFORM - Growth Portfolio',
     'The learner reflects on their journey and curates a Growth Portfolio: before-state diagnostic results, activity submissions, feedback received, revisions made, final artifact, and self-reflection. The portfolio becomes the primary credential artifact, demonstrably stronger than a completion certificate.'),
]
for i, (label, body) in enumerate(stages, 1):
    p = doc.add_paragraph(style='Normal')
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    set_para_spacing(p, before=0, after=4)
    p.paragraph_format.left_indent = Cm(1.0)
    p.paragraph_format.first_line_indent = Cm(-1.0)
    r = p.add_run(f'Stage {i} - '); set_font(r, size=12)
    r = p.add_run(label + ': '); set_font(r, size=12, bold=True)
    r = p.add_run(body); set_font(r, size=12)

add_heading(doc, '6.2  Know → Do → Become Measurement Framework', level=2)

add_body_mixed(doc, [
    ('Every Growth Journey measures learner progress across three dimensions through ', False, False),
    ('pre- and post-journey assessments', True, False),
    (':', False, False),
], first_indent=False, before=0, after=4)

kdb = [
    ('KNOW', 'Conceptual Understanding', 'Measured through pre- and post-diagnostic assessments, quiz scores, and concept checks.'),
    ('DO', 'Demonstrated Performance', 'Measured through activity challenge submissions evaluated against the educator rubric.'),
    ('BECOME', 'Measurable Behavioral Growth', 'Measured through before/after portfolio comparison, final artifact quality, and rubric improvement scores across iterations.'),
]
for k, sub, desc in kdb:
    p = doc.add_paragraph(style='Normal')
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    set_para_spacing(p, before=0, after=4)
    p.paragraph_format.left_indent = Cm(1.0)
    p.paragraph_format.first_line_indent = Cm(-1.0)
    r = p.add_run(u'\u2022  '); set_font(r, size=12)
    r = p.add_run(k + ' - ' + sub + ': '); set_font(r, size=12, bold=True)
    r = p.add_run(desc); set_font(r, size=12)

add_heading(doc, '6.3  System Architecture', level=2)

add_body_mixed(doc, [
    ('The platform follows a ', False, False),
    ('client-server architecture', True, False),
    (' with clear separation between the Learner Module and the Administrative Module, communicating through a ', False, False),
    ('RESTful API layer', True, False),
    ('. The system is designed as a ', False, False),
    ('modular monolith', True, False),
    (' with a componentized React.js SPA frontend, Node.js + Express.js backend, a core services layer (Authentication, Growth Journey, Assessment, Feedback, Portfolio, Analytics), and a persistent data layer (PostgreSQL + Cloud Storage). Figure 2 below presents the approved system architecture.', False, False),
])

# Architecture diagram – IMAGE ONLY (no table repeat)
doc.add_picture('ppt_images/image8.png', width=Cm(15.5))
p = doc.paragraphs[-1]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
cap = doc.add_paragraph()
cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
set_para_spacing(cap, before=2, after=8)
r = cap.add_run('Figure 2: '); set_font(r, size=10, bold=True)
r = cap.add_run('ELEVATE System Architecture - React Frontend, Node.js + Express API, Core Services, and Data Layer.'); set_font(r, size=10, italic=True)

add_heading(doc, '6.4  Core System Modules', level=2)

modules = [
    ('User Authentication and Role Management', 'Learner / Admin / Educator roles with JWT-based stateless authentication and Role-Based Access Control (RBAC).'),
    ('Learner Growth Profile', 'Diagnostic engine and personalization logic based on the learner\'s initial competency state.'),
    ('Journey and Stage Management (Admin)', 'Authoring Growth Journeys with per-stage configuration, content, and rubric setup.'),
    ('Learning Capsule Delivery', 'Educator-led content playback in 8–15 minute focused capsules.'),
    ('Activity Submission System', 'File/video/text submission with status tracking and revision support.'),
    ('ELEVATE Feedback Pipeline', 'Rubric scoring, and peer review orchestration.'),
    ('Growth Portfolio Builder', 'Evidence aggregation, before/after comparison, and skill badge issuance.'),
    ('Progress Dashboard (Learner)', 'Journey stage status, feedback summary, and portfolio view.'),
    ('Administrative Dashboard', 'Journey analytics, submission queues, and learner cohort monitoring.'),
]
for mod, desc in modules:
    p = doc.add_paragraph(style='Normal')
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    set_para_spacing(p, before=0, after=3)
    p.paragraph_format.left_indent = Cm(1.0)
    p.paragraph_format.first_line_indent = Cm(-1.0)
    r = p.add_run(u'\u2022  '); set_font(r, size=12)
    r = p.add_run(mod + ': '); set_font(r, size=12, bold=True)
    r = p.add_run(desc); set_font(r, size=12)

add_heading(doc, '6.5  Development Methodology', level=2)

add_body_mixed(doc, [
    ('Development will proceed under an ', False, False),
    ('iterative, phase-wise methodology', True, False),
    (' aligned to the ELEVATE model\'s architectural layers:', False, False),
], first_indent=False, before=0, after=4)

phases = [
    ('Phase 1 - Requirement Analysis', 'Formalizing functional and non-functional requirements from the ELEVATE pedagogical specification and representative organizational use cases.'),
    ('Phase 2 - System Design', 'Designing the database schema, API contracts, ELEVATE Engine logic, and UI component architecture for both the learner and administrative interfaces.'),
    ('Phase 3 - Implementation', 'Developing in incremental modules: Authentication & Roles → Diagnostic & Profile → Journey & Stage Management → Capsule Delivery → Activity Submission → Feedback Pipeline → Portfolio Builder → Dashboards.'),
    ('Phase 4 - Testing', 'Functional, integration, and usability testing of each module, followed by end-to-end growth journey simulation testing.'),
    ('Phase 5 - Pedagogical Evaluation', 'Pre/post assessment comparison with a learner cohort to evaluate whether ELEVATE produces measurably better skill-development outcomes than conventional content-centric learning, providing the empirical validation component of the project.'),
]
for ph, desc in phases:
    p = doc.add_paragraph(style='Normal')
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    set_para_spacing(p, before=0, after=4)
    p.paragraph_format.left_indent = Cm(1.0)
    p.paragraph_format.first_line_indent = Cm(-1.0)
    r = p.add_run(u'\u2022  '); set_font(r, size=12)
    r = p.add_run(ph + ': '); set_font(r, size=12, bold=True)
    r = p.add_run(desc); set_font(r, size=12)

# ══════════════════════════════════════════════════════════════════════════════
# 7. HARDWARE & SOFTWARE REQUIREMENTS
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, '7.  Hardware & Software Requirements')

add_body_mixed(doc, [
    ('The table below specifies the hardware and software requirements for the ELEVATE platform. Technology stack selections (marked as candidate) will be finalized during detailed design.', False, False),
], first_indent=False, before=0, after=6)

req_data = [
    ('Operating System', 'Windows 10/11, macOS, or Linux (development); Linux-based container image (deployment)'),
    ('Programming Architecture', 'Client-server; RESTful API-based modular monolith; componentized frontend (SPA)'),
    ('Frontend Framework', 'React.js; component-based SPA with role-based routing'),
    ('Backend Framework', 'Node.js with Express.js; alternative: Django REST Framework (Python)'),

    ('Database Management', 'PostgreSQL (relational) or MongoDB (document-oriented), selected in detailed design'),
    ('Auth & Authorization', 'JWT-based stateless authentication; Role-Based Access Control (RBAC) - Learner / Educator / Admin'),
    ('File/Media Storage', 'Cloud object storage for activity submissions (AWS S3 / Firebase Storage)'),
    ('Version Control & CI/CD', 'Git / GitHub; automated build/test pipeline (GitHub Actions)'),
    ('Hosting / Deployment', 'Containerized deployment (Docker) on cloud IaaS/PaaS (AWS / Azure / Render)'),
    ('Development Hardware', 'Intel i5-class or higher; 8 GB RAM minimum (16 GB recommended); 256 GB SSD'),
]

req_tbl = doc.add_table(rows=len(req_data)+1, cols=2)
req_tbl.style = 'Table Grid'
req_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER

# Header
hcells = req_tbl.rows[0].cells
cell_text(hcells[0], 'COMPONENT', bold=True, size=10, align=WD_ALIGN_PARAGRAPH.CENTER)
shade_cell(hcells[0], 'D9D9D9')
hcells[0].width = Cm(4.8)
cell_text(hcells[1], 'TECHNOLOGY / REQUIREMENT', bold=True, size=10, align=WD_ALIGN_PARAGRAPH.CENTER)
shade_cell(hcells[1], 'D9D9D9')
hcells[1].width = Cm(10.7)

for i, (comp, tech) in enumerate(req_data, 1):
    r = req_tbl.rows[i]
    cell_text(r.cells[0], comp, bold=True, size=10)
    r.cells[0].width = Cm(4.8)
    cell_text(r.cells[1], tech, size=10)
    r.cells[1].width = Cm(10.7)

cap = doc.add_paragraph()
cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
set_para_spacing(cap, before=3, after=8)
r = cap.add_run('Table 2: '); set_font(r, size=10, bold=True)
r = cap.add_run('Hardware & Software Requirements for the ELEVATE Platform.'); set_font(r, size=10, italic=True)

# ══════════════════════════════════════════════════════════════════════════════
# 8. EXPECTED DELIVERABLES
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, '8.  Expected Deliverables')

deliverables = [
    ('ELEVATE Pedagogical Model Specification',
     'A formal description of the six-stage framework (Explore → Learn → Experience → Validate → Apply → Transform), its theoretical foundations, and its operationalization within the digital platform.'),
    ('Student Requirement Survey Report',
     'A documented analysis of learning preferences, skill gaps, and expectations gathered from student respondents, validating the platform design decisions.'),
    ('Cross-Platform Web-Based ELEVATE Platform Prototype',
     'A fully functional prototype implementing all six pedagogical stages across at least two Growth Journeys, with role-based access for Student, Educator, and Administrator.'),
    ('Learner Growth Profile Subsystem',
     'Diagnostic engine, personalization logic, and profile dashboard capturing the learner\'s initial competency state.'),
    ('Educator-Led Growth Journey Feature',
     'Structured learning pathways guiding students through activities, reflections, assessments, and competency development with educator oversight.'),
    ('Activity Submission and Multi-Source Feedback Pipeline',
     'Educator rubric scoring, and peer feedback components for structured, actionable feedback on learner submissions.'),
    ('Personalized Growth Portfolio Builder',
     'Evidence aggregation, before/after comparison, and skill badge issuance; a digital portfolio for students to showcase achievements, learning artifacts, reflections, and evidence of skill development over time.'),
    ('Competency-Based Assessment Framework',
     'A framework evaluating learner progress based on predefined competencies and real-world activities, rather than only course completion.'),
    ('Administrative Interface',
     'Journey authoring, stage configuration, learner management, submission queues, and analytics dashboard.'),
    ('Pedagogical Evaluation Report',
     'Pre/post assessment results from a learner cohort comparing ELEVATE-based outcomes against a baseline, providing empirical validation of the pedagogical model.'),
    ('Complete Project Documentation',
     'Architecture diagrams, database schema, API specification, test cases, system design documents, survey analysis, implementation details, and user manuals.'),
]

for i, (label, desc) in enumerate(deliverables, 1):
    p = doc.add_paragraph(style='Normal')
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    set_para_spacing(p, before=0, after=4)
    p.paragraph_format.left_indent = Cm(1.0)
    p.paragraph_format.first_line_indent = Cm(-1.0)
    r = p.add_run(f'{i:02d}.  '); set_font(r, size=12)
    r = p.add_run(label + ': '); set_font(r, size=12, bold=True)
    r = p.add_run(desc); set_font(r, size=12)

# ══════════════════════════════════════════════════════════════════════════════
# 9. REFERENCES
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, '9.  References')

add_body_mixed(doc, [
    ('References are formatted in accordance with the ', False, False),
    ('IEEE Citation Style', True, False),
    ('.', False, False),
], first_indent=False, before=0, after=4)

refs = [
    'D. A. Kolb, Experiential Learning: Experience as the Source of Learning and Development. Englewood Cliffs, NJ, USA: Prentice Hall, 1984. [Online]. Available: https://www.pearson.com/en-us/subject-catalog/p/Kolb-Experiential-Learning/P200000000384',
    'B. S. Bloom, Taxonomy of Educational Objectives: Handbook I - Cognitive Domain. New York, NY, USA: David McKay, 1956. [Online]. Available: https://archive.org/details/taxonomyofeducat00bloo',
    'D. A. Schön, The Reflective Practitioner: How Professionals Think in Action. New York, NY, USA: Basic Books, 1983. [Online]. Available: https://books.google.com/books/about/The_Reflective_Practitioner.html',
    'N. Dabbagh and A. Kitsantas, "Personal Learning Environments, Social Media, and Self-Regulated Learning: A Natural Formula for Connecting Formal and Informal Learning," The Internet and Higher Education, vol. 15, no. 1, pp. 3–8, 2012. doi: 10.1016/j.iheduc.2011.06.002',
    'G. Siemens, "Connectivism: A Learning Theory for the Digital Age," International Journal of Instructional Technology and Distance Learning, vol. 2, no. 1, 2005. [Online]. Available: http://www.itdl.org/Journal/Jan_05/article01.htm',
    'J. Hattie and H. Timperley, "The Power of Feedback," Review of Educational Research, vol. 77, no. 1, pp. 81–112, 2007. doi: 10.3102/003465430298487',
    'D. Ifenthaler and J. Y.-K. Yau, "Utilizing Learning Analytics to Support Study Success in Higher Education," Educational Technology Research and Development, vol. 68, no. 4, pp. 1961–1973, 2020. doi: 10.1007/s11423-020-09788-z',
    'UNESCO, Reimagining Our Futures Together: A New Social Contract for Education. Paris, France: UNESCO Publishing, 2021. [Online]. Available: https://unesdoc.unesco.org/ark:/48223/pf0000379707',
    'A. Warikoo, WebVeda Learning Platform. [Online]. Available: https://webveda.com',
    'React Documentation. [Online]. Available: https://react.dev',
    'Moodle Learning Management System Documentation. [Online]. Available: https://docs.moodle.org',
    'TalentLMS Platform Overview. [Online]. Available: https://www.talentlms.com',
]

for i, ref in enumerate(refs, 1):
    p = doc.add_paragraph(style='Normal')
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    set_para_spacing(p, before=0, after=4)
    p.paragraph_format.left_indent = Cm(0.9)
    p.paragraph_format.first_line_indent = Cm(-0.9)
    r = p.add_run(f'[{i}]  '); set_font(r, size=11, bold=True)
    r = p.add_run(ref); set_font(r, size=11)

# ══════════════════════════════════════════════════════════════════════════════
# SIGNATURE BLOCK
# ══════════════════════════════════════════════════════════════════════════════
add_para(doc, before=24)  # spacing before signatures

sig_tbl = doc.add_table(rows=3, cols=2)
sig_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER

# Remove table borders
for row in sig_tbl.rows:
    for cell in row.cells:
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        tcBorders = OxmlElement('w:tcBorders')
        for side in ['top', 'bottom', 'left', 'right']:
            el = OxmlElement(f'w:{side}')
            el.set(qn('w:val'), 'none')
            tcBorders.append(el)
        tcPr.append(tcBorders)

# Signature lines via underscores
def sig_line(cell, text, sub1='', sub2=''):
    cell.text = ''
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_para_spacing(p, before=0, after=2)
    r = p.add_run('_' * 40); set_font(r, size=11)

    p2 = cell.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_para_spacing(p2, before=2, after=2)
    r2 = p2.add_run(text); set_font(r2, size=11, bold=True)

    if sub1:
        p3 = cell.add_paragraph()
        p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_para_spacing(p3, before=0, after=1)
        r3 = p3.add_run(sub1); set_font(r3, size=10)
    if sub2:
        p4 = cell.add_paragraph()
        p4.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_para_spacing(p4, before=0, after=1)
        r4 = p4.add_run(sub2); set_font(r4, size=10)

sig_line(sig_tbl.rows[0].cells[0], 'Signatures of Students',
         'Kabir Patil (B-26)   ·   Alok Perla (B-34)',
         'Vidhi Patel (B-23)   ·   Rudraksh Vaidya (B-61)')
sig_line(sig_tbl.rows[0].cells[1], 'Signature of Guide',
         'Dr. Ms. D. D. Awasekar',
         'Dept. of Information Technology, WIT Solapur')

for i in range(2): sig_tbl.columns[i].width = Cm(7.5)

# ── SAVE ────────────────────────────────────────────────────────────────────
out_path = r'e:\Fproject\ELEVATE_Synopsis_Final_v2.docx'
doc.save(out_path)
print('Saved:', out_path)
