#!/usr/bin/env python3
"""
Statement of Qualifications — MYR Group / CSI Electrical Contractors
Clean rewrite with precise coordinate math
"""
import os
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT, TA_JUSTIFY
from reportlab.platypus import (
    BaseDocTemplate, PageTemplate, Frame, Paragraph, Spacer,
    Table, TableStyle, HRFlowable, KeepTogether, PageBreak, NextPageTemplate
)
from reportlab.platypus.flowables import Flowable
from reportlab.lib.colors import HexColor

# ── PAGE & COLORS ─────────────────────────────────────────────────────────────
PW, PH = letter          # 612 x 792 pts
PT     = 1               # 1 point
IN     = inch            # 72 pts

NAVY   = HexColor('#0C1F35')
NAVY2  = HexColor('#1B3A5C')
GOLD   = HexColor('#B8852A')
GOLD2  = HexColor('#D4A84B')
GREEN  = HexColor('#1E6B3C')
LTBG   = HexColor('#F4F7FB')
RULE   = HexColor('#CBD5E0')
TXT    = HexColor('#1C2333')
TSUB   = HexColor('#4A5568')
TLIT   = HexColor('#718096')
WHITE  = colors.white

ML = MR = 0.75 * IN      # left/right margin = 54 pts
CW      = PW - ML - MR   # content width = 504 pts
MB      = 0.55 * IN      # bottom margin

# Header strip on interior pages
HDR_H = 0.42 * IN        # 30 pts

# ── STYLES ────────────────────────────────────────────────────────────────────
S = {}
def ps(n, **kw): S[n] = ParagraphStyle(n, **kw)

ps('h2',      fontName='Helvetica-Bold',        fontSize=13,  textColor=NAVY,  leading=16, spaceBefore=12, spaceAfter=5)
ps('h3',      fontName='Helvetica-Bold',        fontSize=10.5,textColor=NAVY2, leading=13, spaceBefore=10, spaceAfter=4)
ps('body',    fontName='Helvetica',             fontSize=9.5, textColor=TXT,   leading=14, spaceAfter=7, alignment=TA_JUSTIFY)
ps('small',   fontName='Helvetica',             fontSize=8.5, textColor=TSUB,  leading=12, spaceAfter=4)
ps('note',    fontName='Helvetica-Oblique',     fontSize=8,   textColor=TLIT,  leading=11, spaceAfter=3)
ps('bullet',  fontName='Helvetica',             fontSize=9.5, textColor=TXT,   leading=14, leftIndent=12, firstLineIndent=-10, spaceAfter=3)
ps('sec_num', fontName='Helvetica',             fontSize=7.5, textColor=GOLD2, leading=9)
ps('sec_sub', fontName='Helvetica',             fontSize=9,   textColor=GOLD2, leading=12, spaceAfter=6)
ps('th',      fontName='Helvetica-Bold',        fontSize=8.5, textColor=WHITE, leading=11)
ps('td',      fontName='Helvetica',             fontSize=9,   textColor=TXT,   leading=12)
ps('tdb',     fontName='Helvetica-Bold',        fontSize=9,   textColor=TXT,   leading=12)
ps('stat_n',  fontName='Helvetica-Bold',        fontSize=22,  textColor=GOLD,  leading=26, alignment=TA_CENTER)
ps('stat_l',  fontName='Helvetica',             fontSize=7.5, textColor=TSUB,  leading=10, alignment=TA_CENTER)
ps('pname',   fontName='Helvetica-Bold',        fontSize=10.5,textColor=NAVY,  leading=13, spaceAfter=1)
ps('ptitl',   fontName='Helvetica-BoldOblique', fontSize=8.5, textColor=GOLD,  leading=11, spaceAfter=4)
ps('pbody',   fontName='Helvetica',             fontSize=8.5, textColor=TSUB,  leading=12, spaceAfter=3, alignment=TA_JUSTIFY)
ps('avatar',  fontName='Helvetica-Bold',        fontSize=12,  textColor=WHITE, leading=15, alignment=TA_CENTER)
ps('cap_h',   fontName='Helvetica-Bold',        fontSize=10.5,textColor=NAVY,  leading=13, spaceAfter=3)
ps('cap_b',   fontName='Helvetica',             fontSize=8.5, textColor=TSUB,  leading=12, alignment=TA_JUSTIFY)
ps('proj_mw', fontName='Helvetica-Bold',        fontSize=16,  textColor=GOLD,  leading=20, alignment=TA_CENTER)
ps('proj_l',  fontName='Helvetica',             fontSize=7,   textColor=TSUB,  leading=9,  alignment=TA_CENTER)
ps('proj_v',  fontName='Helvetica-Bold',        fontSize=12,  textColor=NAVY,  leading=15, alignment=TA_CENTER)
ps('proj_tag',fontName='Helvetica-Bold',        fontSize=7,   textColor=WHITE, leading=9,  alignment=TA_CENTER)
ps('ref_nm',  fontName='Helvetica-Bold',        fontSize=10.5,textColor=NAVY,  leading=13, spaceAfter=1)
ps('ref_co',  fontName='Helvetica-BoldOblique', fontSize=8.5, textColor=GOLD,  leading=11, spaceAfter=5)
ps('ref_b',   fontName='Helvetica',             fontSize=9,   textColor=TSUB,  leading=13, alignment=TA_JUSTIFY)
ps('org_box', fontName='Helvetica-Bold',        fontSize=9,   textColor=WHITE, leading=12, alignment=TA_CENTER)
ps('fin_big', fontName='Helvetica-Bold',        fontSize=18,  textColor=GOLD,  leading=22, alignment=TA_CENTER)
ps('fin_l',   fontName='Helvetica',             fontSize=7.5, textColor=TSUB,  leading=10, alignment=TA_CENTER)

# ── COVER PAGE ────────────────────────────────────────────────────────────────
def draw_cover(c, doc):
    c.setFillColor(NAVY)
    c.rect(0, 0, PW, PH, fill=1, stroke=0)

    c.saveState()
    c.setStrokeColor(HexColor('#122030'))
    c.setLineWidth(0.5)
    for x in range(-200, int(PW) + 300, 22):
        c.line(x, 0, x + PH, PH)
    c.restoreState()

    c.setFillColor(GOLD)
    c.rect(0, PH - 18, PW, 18, fill=1, stroke=0)
    c.setFillColor(GREEN)
    c.rect(0, PH - 25, PW, 7, fill=1, stroke=0)
    c.setFillColor(NAVY)
    c.setFont('Helvetica-Bold', 7.5)
    c.drawString(ML, PH - 13, 'MYR GROUP INC. / CSI ELECTRICAL CONTRACTORS')
    c.setFont('Helvetica', 7.5)
    c.drawRightString(PW - MR, PH - 13, 'NASDAQ: MYRG')

    # Left accent bar from bottom card to top stripe
    c.setFillColor(GOLD)
    c.rect(ML - 9, 160, 4, PH - 30 - 160, fill=1, stroke=0)

    TX = ML + 8

    # Title block — just under vertical center, left justified
    y = PH / 2 + 20  # start just above center, text flows down

    c.setFillColor(GOLD2)
    c.setFont('Helvetica', 8.5)
    c.drawString(TX, y, 'PREQUALIFICATION SUBMITTAL  \u2014  2026\u20132028')

    y -= 42
    c.setFillColor(WHITE)
    c.setFont('Helvetica-Bold', 44)
    c.drawString(TX, y, 'MYR GROUP')

    y -= 10
    c.setFillColor(GOLD)
    c.rect(TX, y, 3.8 * IN, 2.5, fill=1, stroke=0)

    y -= 32
    c.setFillColor(WHITE)
    c.setFont('Helvetica-Bold', 26)
    c.drawString(TX, y, 'Statement of Qualifications')

    y -= 24
    c.setFillColor(GOLD2)
    c.setFont('Helvetica', 13)
    c.drawString(TX, y, 'Renewable Energy Electrical Construction')

    y -= 18
    c.setFillColor(GOLD)
    c.rect(TX, y, 4.1 * IN, 1.5, fill=1, stroke=0)

    y -= 22
    tags = ['Solar', 'Wind', 'BESS', 'Substation', 'HV Interconnection']
    tx2 = TX
    c.setFont('Helvetica', 8)
    for tag in tags:
        tw = c.stringWidth(tag, 'Helvetica', 8)
        box_w = tw + 14
        c.setFillColor(NAVY2)
        c.roundRect(tx2, y - 3, box_w, 16, 3, fill=1, stroke=0)
        c.setFillColor(GOLD2)
        c.drawString(tx2 + 7, y + 2, tag)
        tx2 += box_w + 10

    card_bot = 33
    card_h   = 100
    card_top = card_bot + card_h

    c.setFillColor(NAVY2)
    c.rect(ML, card_bot, CW, card_h, fill=1, stroke=0)
    c.setFillColor(GOLD)
    c.rect(ML, card_top - 3, CW, 3, fill=1, stroke=0)

    col_w = CW / 3
    card_data = [
        ('Submitted To',   'Pacific Northwest Energy Authority (PNEA)'),
        ('Document Type',  'Statement of Qualifications'),
        ('Submittal Date', 'March 17, 2026'),
    ]
    for i, (lbl, val) in enumerate(card_data):
        cx = ML + i * col_w + 13
        c.setFillColor(GOLD2)
        c.setFont('Helvetica-Bold', 6.5)
        c.drawString(cx, card_top - 20, lbl.upper())
        c.setFillColor(WHITE)
        c.setFont('Helvetica-Bold', 9)
        words = val.split()
        line1, line2 = '', ''
        for w in words:
            test = (line1 + ' ' + w).strip()
            if c.stringWidth(test, 'Helvetica-Bold', 9) < col_w - 22:
                line1 = test
            else:
                line2 = (line2 + ' ' + w).strip()
        c.drawString(cx, card_top - 38, line1)
        if line2:
            c.drawString(cx, card_top - 52, line2)
    for i in [1, 2]:
        xd = ML + i * col_w
        c.setFillColor(NAVY)
        c.rect(xd, card_bot + 10, 1, card_h - 20, fill=1, stroke=0)

    c.setFillColor(GREEN)
    c.rect(0, 0, PW, 23, fill=1, stroke=0)
    c.setFillColor(WHITE)
    c.setFont('Helvetica', 7.5)
    c.drawCentredString(PW / 2, 8,
        'Solar  •  Wind  •  Battery Energy Storage  •  Substation  •  High-Voltage Interconnection')


# ── INTERIOR PAGE HEADER & FOOTER ────────────────────────────────────────────
def draw_interior(c, doc):
    c.setFillColor(NAVY)
    c.rect(0, PH - HDR_H, PW, HDR_H, fill=1, stroke=0)
    c.setFillColor(GOLD)
    c.rect(0, PH - HDR_H - 2, PW, 2, fill=1, stroke=0)
    c.setFillColor(GOLD2)
    c.setFont('Helvetica', 7.5)
    c.drawString(ML, PH - HDR_H + 20, getattr(doc, '_sec_num', ''))
    c.setFillColor(WHITE)
    c.setFont('Helvetica-Bold', 14)
    c.drawString(ML, PH - HDR_H + 6, getattr(doc, '_sec_title', ''))
    sub = getattr(doc, '_sec_sub', '')
    if sub:
        c.setFillColor(GOLD2)
        c.setFont('Helvetica', 8.5)
        c.drawRightString(PW - MR, PH - HDR_H + 10, sub)
    c.setFillColor(NAVY)
    c.rect(0, 0, PW, 28, fill=1, stroke=0)
    c.setFillColor(GOLD)
    c.rect(0, 28, PW, 2, fill=1, stroke=0)
    c.setFillColor(GOLD2)
    c.setFont('Helvetica', 7)
    c.drawString(ML, 10, 'MYR Group Inc. / CSI Electrical Contractors')
    c.setFillColor(TLIT)
    c.drawCentredString(PW / 2, 10, 'Statement of Qualifications — Renewable Energy')
    c.setFillColor(GOLD2)
    pg = getattr(doc, '_pg', 0)
    c.drawRightString(PW - MR, 10, f'Page {pg}')


# ── SECTION MARKER ────────────────────────────────────────────────────────────
class SectionMark(Flowable):
    def __init__(self, num, title, sub=''):
        super().__init__()
        self.num = num; self.title = title; self.sub = sub
        self.width = self.height = 0
    def wrap(self, *a): return (0, 0)
    def draw(self): pass
    def beforeDraw(self):
        d = self.canv._doc
        d._sec_num = self.num
        d._sec_title = self.title
        d._sec_sub = self.sub


# ── DOC CLASS ─────────────────────────────────────────────────────────────────
class SOQDoc(BaseDocTemplate):
    def __init__(self, filename):
        super().__init__(filename, pagesize=letter,
                         leftMargin=ML, rightMargin=MR,
                         topMargin=HDR_H + 14, bottomMargin=MB)
        self._pg = 0
        self._sec_num = self._sec_title = self._sec_sub = ''

        cover_frame = Frame(0, 0, PW, PH,
                            leftPadding=0, rightPadding=0,
                            topPadding=0, bottomPadding=0, id='cover')
        body_frame  = Frame(ML, MB,
                            CW, PH - HDR_H - 14 - MB,
                            leftPadding=0, rightPadding=0,
                            topPadding=0, bottomPadding=0, id='body')
        self.addPageTemplates([
            PageTemplate(id='Cover', frames=[cover_frame], onPage=draw_cover),
            PageTemplate(id='Body',  frames=[body_frame],  onPage=self._on_body),
        ])

    def _on_body(self, c, doc):
        self._pg += 1
        draw_interior(c, doc)


# ── HELPERS ───────────────────────────────────────────────────────────────────
def sec_block(num, title, sub=''):
    items = [
        SectionMark(num, title, sub),
        Paragraph(num, S['sec_num']),
        Paragraph(title, S['h2']),
    ]
    if sub:
        items.append(Paragraph(sub, S['sec_sub']))
    items.append(HRFlowable(width=CW, thickness=1, color=GOLD, spaceAfter=10))
    return items

def stat_row(items):
    n = len(items)
    w = CW / n
    data = [[Paragraph(v, S['stat_n']) for v,_ in items],
            [Paragraph(l, S['stat_l']) for _,l in items]]
    t = Table(data, colWidths=[w]*n)
    t.setStyle(TableStyle([
        ('BACKGROUND',    (0,0),(-1,-1), LTBG),
        ('BOX',           (0,0),(-1,-1), 0.5, RULE),
        ('INNERGRID',     (0,0),(-1,-1), 0.5, RULE),
        ('TOPPADDING',    (0,0),(-1,-1), 10),
        ('BOTTOMPADDING', (0,0),(-1,-1), 8),
        ('LEFTPADDING',   (0,0),(-1,-1), 4),
        ('RIGHTPADDING',  (0,0),(-1,-1), 4),
    ]))
    return t

def cap_card(title, body, w=None):
    w = w or (CW - 6) / 2
    inner = [[
        [Paragraph(title, S['cap_h']), Paragraph(body, S['cap_b'])]
    ]]
    t = Table(inner, colWidths=[w - 16])
    t.setStyle(TableStyle([
        ('VALIGN',        (0,0),(-1,-1), 'TOP'),
        ('LEFTPADDING',   (0,0),(-1,-1), 8),
        ('RIGHTPADDING',  (0,0),(-1,-1), 8),
        ('TOPPADDING',    (0,0),(-1,-1), 9),
        ('BOTTOMPADDING', (0,0),(-1,-1), 9),
        ('BACKGROUND',    (0,0),(-1,-1), LTBG),
        ('BOX',           (0,0),(-1,-1), 0.5, RULE),
    ]))
    return t

def two_cap_row(caps_pair, w=None):
    w = w or (CW - 6) / 2
    row = [cap_card(c[0], c[1], w) if c else '' for c in caps_pair]
    t = Table([row], colWidths=[w, w])
    t.setStyle(TableStyle([
        ('LEFTPADDING',   (0,0),(-1,-1), 0),
        ('RIGHTPADDING',  (0,0),(-1,-1), 0),
        ('TOPPADDING',    (0,0),(-1,-1), 0),
        ('BOTTOMPADDING', (0,0),(-1,-1), 4),
    ]))
    return t


# ── STORY ─────────────────────────────────────────────────────────────────────
def build_story():
    st = []

    st.append(NextPageTemplate('Body'))
    st.append(PageBreak())

    # ── 01 COMPANY PROFILE ────────────────────────────────────────────────────
    st += sec_block('01', 'Company Profile', 'A National Leader in Electrical Construction')
    st.append(Paragraph(
        'MYR Group Inc. (NASDAQ: MYRG) is one of the largest electrical construction firms in North America, '
        'providing a full spectrum of electrical construction services across the commercial, industrial, and '
        'utility sectors. Through our subsidiary CSI Electrical Contractors, we deliver comprehensive renewable '
        'energy construction solutions including utility-scale solar, wind collector systems, battery energy '
        'storage, and high-voltage substation and interconnection work.', S['body']))
    st.append(Paragraph(
        'Headquartered in the Western United States, CSI Electrical has established itself as a premier '
        'contractor for some of the most complex and high-profile renewable energy projects in the nation, '
        'with particular emphasis on the Pacific Northwest, Southwest, and Mountain West regions. Our craft '
        'workforce is 100% IBEW-affiliated, delivering the highest caliber of professional electrical '
        'construction in every market we serve.', S['body']))
    st.append(Spacer(1, 8))
    st.append(stat_row([
        ('$3.6B',  '2025 Annual Revenue'),
        ('9,500+', 'Employees Nationwide'),
        ('40+',    'Years in Business'),
        ('14',     'Subsidiaries / Regions'),
        ('50+',    'States Served'),
        ('0.71',   'EMR Rating (2025)'),
    ]))
    st.append(Spacer(1, 12))

    # Org structure
    st.append(Paragraph('Organizational Structure', S['h3']))
    def org_cell(label, sub, color, w):
        inner = [[
            Paragraph(label, S['org_box']),
            Paragraph(sub, ParagraphStyle('os', fontName='Helvetica', fontSize=7,
                                          textColor=GOLD2, leading=9, alignment=TA_CENTER))
        ]]
        t = Table(inner, colWidths=[w])
        t.setStyle(TableStyle([
            ('BACKGROUND',    (0,0),(-1,-1), color),
            ('BOX',           (0,0),(-1,-1), 1.2, GOLD),
            ('TOPPADDING',    (0,0),(-1,-1), 5),
            ('BOTTOMPADDING', (0,0),(-1,-1), 5),
            ('ALIGN',         (0,0),(-1,-1), 'CENTER'),
            ('VALIGN',        (0,0),(-1,-1), 'MIDDLE'),
        ]))
        return t

    top = Table([[org_cell('MYR Group Inc.', 'Parent Company — NASDAQ: MYRG', NAVY, 2.8*IN)]],
                colWidths=[CW])
    top.setStyle(TableStyle([('ALIGN',(0,0),(-1,-1),'CENTER'),
                              ('TOPPADDING',(0,0),(-1,-1),0),('BOTTOMPADDING',(0,0),(-1,-1),5)]))
    st.append(top)
    mid = Table([[org_cell('CSI Electrical Contractors', 'Subsidiary — Renewable Energy Division', NAVY2, 3.2*IN)]],
                colWidths=[CW])
    mid.setStyle(TableStyle([('ALIGN',(0,0),(-1,-1),'CENTER'),
                              ('TOPPADDING',(0,0),(-1,-1),0),('BOTTOMPADDING',(0,0),(-1,-1),5)]))
    st.append(mid)
    depts = [('Project\nManagement','Planning & Cost Control'),
             ('Engineering','Design & Estimation'),
             ('Field\nOperations','Construction & Testing'),
             ('Safety','HSE & Compliance'),
             ('Quality','QA/QC & Inspections')]
    dcw = CW / 5
    dept_cells = []
    for d in depts:
        cell = Table([[Paragraph(d[0], S['org_box']),
                       Paragraph(d[1], ParagraphStyle('ds', fontName='Helvetica', fontSize=7,
                                                       textColor=GOLD2, leading=9, alignment=TA_CENTER))]],
                     colWidths=[dcw - 6])
        cell.setStyle(TableStyle([
            ('BACKGROUND', (0,0),(-1,-1), HexColor('#1B3A5C')),
            ('BOX',        (0,0),(-1,-1), 1, GOLD),
            ('TOPPADDING', (0,0),(-1,-1),5),('BOTTOMPADDING',(0,0),(-1,-1),5),
            ('ALIGN',      (0,0),(-1,-1),'CENTER'),('VALIGN',(0,0),(-1,-1),'MIDDLE'),
        ]))
        dept_cells.append(cell)
    dept_row = Table([dept_cells], colWidths=[dcw]*5)
    dept_row.setStyle(TableStyle([
        ('LEFTPADDING',(0,0),(-1,-1),3),('RIGHTPADDING',(0,0),(-1,-1),3),
        ('TOPPADDING',(0,0),(-1,-1),0),('BOTTOMPADDING',(0,0),(-1,-1),0),
    ]))
    st.append(dept_row)

    # ── 02 CORE CAPABILITIES ──────────────────────────────────────────────────
    st.append(PageBreak())
    st += sec_block('02', 'Core Capabilities', 'Full-Spectrum Renewable Energy Electrical Construction')
    caps = [
        ('Solar Farm Construction',
         'Utility-scale PV installation: tracker systems, fixed-tilt arrays, DC wiring, combiner boxes, '
         'inverter pads, and AC collection for projects from 50 MW to 500+ MW.'),
        ('Wind Farm Collector Systems',
         'Underground and overhead MV collector systems, turbine pad electrical, step-up transformers, '
         'and meteorological tower wiring for facilities up to 300+ MW.'),
        ('Battery Energy Storage (BESS)',
         'Complete electrical BOP for lithium-ion and emerging technologies: DC power systems, inverter '
         'installation, HVAC electrical, fire suppression wiring, and grid interconnection.'),
        ('Substation Construction',
         'Greenfield and brownfield substations from 34.5 kV to 500 kV — foundations, steel erection, '
         'bus work, equipment setting, grounding grids, control buildings, relay protection.'),
        ('High-Voltage Interconnection',
         'Transmission-level interconnection: gen-tie lines, switching stations, and point-of-interconnection '
         'infrastructure integrating renewable assets to the bulk power grid.'),
        ('Medium-Voltage Distribution',
         'Design-build and design-assist MV distribution including pad-mounted switchgear, underground '
         'duct banks, and overhead distribution up to 34.5 kV.'),
        ('SCADA & Controls',
         'SCADA installation, fiber optic networks, RTU programming, DAS integration, meteorological '
         'monitoring, and real-time plant performance systems.'),
        ('Commissioning & Testing',
         'Relay testing, insulation resistance, hi-pot testing, transformer oil analysis, power quality '
         'analysis, and coordinated system energization.'),
    ]
    for i in range(0, len(caps), 2):
        pair = [caps[i], caps[i+1] if i+1 < len(caps) else None]
        st.append(two_cap_row(pair))

    # ── 03 PROJECT PORTFOLIO ──────────────────────────────────────────────────
    st.append(PageBreak())
    st += sec_block('03', 'Project Portfolio', 'Selected Renewable Energy Projects — Western U.S.')

    TAG_COLOR = {'SOLAR': GOLD, 'WIND': GREEN, 'BESS': NAVY2}
    projects = [
        ('SOLAR','Sagebrush Flats Solar Energy Center','Klickitat County, WA','350 MW','2025','IPP',
         'Complete electrical BOP for a utility-scale solar facility with single-axis tracker arrays. '
         'AC/DC collection system, inverter pads, 34.5 kV overhead collector lines, onsite 230 kV substation, '
         'and gen-tie interconnection to the regional grid.'),
        ('WIND','Cascade Ridge Wind Farm','Sherman County, OR','280 MW','2024','Utility',
         'Underground 34.5 kV collector system for 85 wind turbine generators, turbine transformer wiring, '
         'pad-mounted switchgear, fiber optic SCADA network, and a new 230 kV collector substation with '
         'dual transformer banks.'),
        ('BESS','Silver Creek Energy Storage Facility','Ada County, ID','200 MW / 800 MWh','2025','Developer',
         'Electrical construction for a lithium-ion BESS co-located with an existing solar facility. Inverter '
         'installation, MV switchgear, HVAC power, fire detection wiring, and 230 kV interconnection upgrades.'),
        ('SOLAR','Antelope Mesa Solar Project','Wasco County, OR','200 MW','2024','IPP',
         'Full electrical scope for a bifacial solar facility: 1,200+ string inverters, 42 miles of underground '
         'MV cable, SCADA/DAS systems, and a new 115 kV POI substation with OPGW communication links.'),
        ('WIND','Rattlesnake Hills Wind Energy Center','Benton County, WA','150 MW','2023','Utility',
         '55 turbine pad electrical installations, 34.5 kV underground collector system, permanent met towers, '
         'O&M building electrical, and integration with existing 230 kV transmission switchyard.'),
        ('BESS','Columbia Gateway Storage Project','Umatilla County, OR','150 MW / 600 MWh','2026','Developer',
         'Standalone grid-scale BESS: 120 battery enclosures, power conversion systems, 34.5/230 kV step-up '
         'substation, and utility metering infrastructure.'),
    ]

    pw2 = (CW - 6) / 2
    for i in range(0, len(projects), 2):
        row_cells = []
        for proj in projects[i:i+2]:
            tag, name, loc, cap, yr, client, desc = proj
            tc = TAG_COLOR.get(tag, NAVY2)
            tag_bar = Table([[Paragraph(tag, S['proj_tag'])]], colWidths=[pw2])
            tag_bar.setStyle(TableStyle([
                ('BACKGROUND',(0,0),(-1,-1),tc),
                ('TOPPADDING',(0,0),(-1,-1),4),('BOTTOMPADDING',(0,0),(-1,-1),4),
            ]))
            pw3 = pw2 / 3
            mini = Table([[
                [Paragraph(cap,    S['proj_mw']), Paragraph('Capacity', S['proj_l'])],
                [Paragraph(yr,     S['proj_v']),  Paragraph('Completed', S['proj_l'])],
                [Paragraph(client, S['proj_v']),  Paragraph('Client', S['proj_l'])],
            ]], colWidths=[pw3]*3)
            mini.setStyle(TableStyle([
                ('BACKGROUND',    (0,0),(-1,-1), LTBG),
                ('INNERGRID',     (0,0),(-1,-1), 0.4, RULE),
                ('BOX',           (0,0),(-1,-1), 0.4, RULE),
                ('TOPPADDING',    (0,0),(-1,-1), 5),
                ('BOTTOMPADDING', (0,0),(-1,-1), 5),
            ]))
            inner = [tag_bar, Spacer(1,6),
                     Paragraph(name, S['pname']), Paragraph(loc, S['note']),
                     Spacer(1,4), mini, Spacer(1,6), Paragraph(desc, S['cap_b'])]
            card = Table([[inner]], colWidths=[pw2])
            card.setStyle(TableStyle([
                ('BOX',           (0,0),(-1,-1), 0.5, RULE),
                ('TOPPADDING',    (0,0),(-1,-1), 0),
                ('BOTTOMPADDING', (0,0),(-1,-1), 8),
                ('LEFTPADDING',   (0,0),(-1,-1), 8),
                ('RIGHTPADDING',  (0,0),(-1,-1), 8),
            ]))
            row_cells.append(card)
        if len(row_cells) == 1:
            row_cells.append('')
        row = Table([row_cells], colWidths=[pw2, pw2])
        row.setStyle(TableStyle([
            ('LEFTPADDING',(0,0),(-1,-1),0),('RIGHTPADDING',(0,0),(-1,-1),0),
            ('TOPPADDING',(0,0),(-1,-1),0),('BOTTOMPADDING',(0,0),(-1,-1),6),
        ]))
        st.append(row)

    # ── 04 KEY PERSONNEL ──────────────────────────────────────────────────────
    st.append(PageBreak())
    st += sec_block('04', 'Key Personnel', 'Experienced Leadership Driving Project Success')

    personnel = [
        ('RM','Robert Maynard','Senior Project Director','28 Yrs',['PMP','PE (OR,WA)','OSHA 30'],
         'Oversees CSI Electrical\'s renewable energy portfolio across the Pacific Northwest. Has directed '
         'more than $1.2B in solar, wind, and BESS projects.'),
        ('KL','Karen Liu','Project Manager — Renewables','19 Yrs',['PMP','LEED AP','OSHA 30'],
         'Manages day-to-day execution of utility-scale renewable projects from preconstruction through '
         'commissioning. Led the Sagebrush Flats 350 MW project to on-time completion.'),
        ('JT','James Taggart','General Superintendent','24 Yrs',['Master Elec.','OSHA 500','NFPA 70E'],
         'Directs all field operations for renewable energy construction. IBEW journeyman wireman. '
         'Zero lost-time incidents across three consecutive major projects.'),
        ('SP','Sarah Petersen','Director of Safety','16 Yrs',['CSP','CHST','OSHA 510'],
         'Leads CSI Electrical\'s safety program across all renewable energy operations. EMR maintained '
         'well below industry average for five consecutive years.'),
        ('DW','David Walsh','Chief Estimator — Renewables','21 Yrs',['CPE','PE (WA)','OSHA 30'],
         'Responsible for preconstruction, estimating, and value engineering. Has estimated more than '
         '$2B in awarded renewable energy contracts.'),
    ]
    ph2 = (CW - 6) / 2
    for i in range(0, len(personnel), 2):
        row_cells = []
        for p in personnel[i:i+2]:
            init, name, title, exp, creds, bio = p
            av = Table([[Paragraph(init, S['avatar'])]], colWidths=[0.52*IN])
            av.setStyle(TableStyle([
                ('BACKGROUND',    (0,0),(-1,-1), NAVY2),
                ('TOPPADDING',    (0,0),(-1,-1), 7),
                ('BOTTOMPADDING', (0,0),(-1,-1), 7),
                ('ALIGN',         (0,0),(-1,-1), 'CENTER'),
            ]))
            cred_str = '  ·  '.join(f'<font color="#B8852A"><b>{c}</b></font>' for c in creds)
            top_t = Table([[av, [Paragraph(name, S['pname']),
                                 Paragraph(title, S['ptitl']),
                                 Paragraph(f'<font color="#1B3A5C"><b>{exp} Experience</b></font>', S['small'])]]],
                          colWidths=[0.60*IN, ph2 - 0.75*IN])
            top_t.setStyle(TableStyle([
                ('VALIGN',(0,0),(-1,-1),'TOP'),
                ('LEFTPADDING',(0,0),(-1,-1),0),('RIGHTPADDING',(0,0),(-1,-1),0),
                ('TOPPADDING',(0,0),(-1,-1),0),('BOTTOMPADDING',(0,0),(-1,-1),5),
            ]))
            card_inner = [top_t,
                          Paragraph(f'Credentials: {cred_str}', S['small']),
                          Paragraph(bio, S['pbody'])]
            card = Table([[card_inner]], colWidths=[ph2])
            card.setStyle(TableStyle([
                ('BOX',           (0,0),(-1,-1), 0.5, RULE),
                ('BACKGROUND',    (0,0),(-1,-1), LTBG),
                ('TOPPADDING',    (0,0),(-1,-1), 9),
                ('BOTTOMPADDING', (0,0),(-1,-1), 9),
                ('LEFTPADDING',   (0,0),(-1,-1), 9),
                ('RIGHTPADDING',  (0,0),(-1,-1), 9),
            ]))
            row_cells.append(card)
        if len(row_cells) == 1:
            row_cells.append(Spacer(ph2, 1))
        row = Table([row_cells], colWidths=[ph2, ph2])
        row.setStyle(TableStyle([
            ('LEFTPADDING',(0,0),(-1,-1),0),('RIGHTPADDING',(0,0),(-1,-1),0),
            ('TOPPADDING',(0,0),(-1,-1),0),('BOTTOMPADDING',(0,0),(-1,-1),6),
        ]))
        st.append(KeepTogether(row))

    # ── 05 SAFETY & QUALITY ───────────────────────────────────────────────────
    st.append(PageBreak())
    st += sec_block('05', 'Safety & Quality', 'Industry-Leading Performance & Rigorous Programs')

    st.append(Paragraph('EMR History  (Industry Average: 1.0 — Lower is Better)', S['h3']))

    emr = [('2021','0.78'),('2022','0.75'),('2023','0.73'),('2024','0.72'),('2025','0.71')]

    class EMRChart(Flowable):
        def __init__(self, data, w):
            super().__init__()
            self.data = data; self.width = w; self.height = 100
        def draw(self):
            c = self.canv
            max_h = 68.0
            bw = self.width / len(self.data)
            bar_w = bw * 0.5
            base_y = 22
            for i,(yr,val) in enumerate(self.data):
                bh = (float(val)/1.0) * max_h
                bx = i*bw + (bw-bar_w)/2
                c.setFillColor(NAVY2)
                c.rect(bx, base_y, bar_w, bh, fill=1, stroke=0)
                c.setFillColor(GOLD)
                c.rect(bx, base_y+bh-2, bar_w, 2, fill=1, stroke=0)
                c.setFillColor(NAVY)
                c.setFont('Helvetica-Bold', 10)
                c.drawCentredString(bx+bar_w/2, base_y+bh+4, val)
                c.setFillColor(GOLD2)
                c.setFont('Helvetica-Bold', 7.5)
                c.drawCentredString(bx+bar_w/2, 8, yr)
            c.setFillColor(RULE)
            c.rect(0, base_y-1, self.width, 1, fill=1, stroke=0)

    st.append(EMRChart(emr, CW))
    st.append(Spacer(1, 10))

    metrics = [
        ('Total Recordable Incident Rate (TRIR)', '0.42'),
        ('Days Away / Restricted / Transfer (DART)', '0.18'),
        ('Lost Time Incident Rate (LTIR)', '0.05'),
        ('Experience Modification Rate (EMR)', '0.71'),
        ('Hours Worked Without Lost Time', '2.8M+'),
        ('ISNetworld RAVS Score', 'Grade A'),
    ]
    mrows = [[Paragraph('Safety Performance Metric', S['th']), Paragraph('Result', S['th'])]]
    for k,v in metrics:
        mrows.append([Paragraph(k, S['td']), Paragraph(v, S['tdb'])])
    mt = Table(mrows, colWidths=[CW*0.70, CW*0.30])
    mt.setStyle(TableStyle([
        ('BACKGROUND',    (0,0),(-1,0),  NAVY),
        ('ROWBACKGROUNDS',(0,1),(-1,-1), [WHITE, LTBG]),
        ('GRID',          (0,0),(-1,-1), 0.4, RULE),
        ('TOPPADDING',    (0,0),(-1,-1), 5),
        ('BOTTOMPADDING', (0,0),(-1,-1), 5),
        ('LEFTPADDING',   (0,0),(-1,-1), 8),
        ('RIGHTPADDING',  (0,0),(-1,-1), 8),
        ('TEXTCOLOR',     (1,1),(1,-1),  GOLD),
        ('FONTNAME',      (1,1),(1,-1),  'Helvetica-Bold'),
    ]))
    st.append(mt)
    st.append(Spacer(1, 10))

    iso = [['ISO 9001:2015\nQuality Mgmt', 'ISO 45001:2018\nOccupational H&S', 'ISO 14001:2015\nEnvironmental Mgmt']]
    iso_t = Table(iso, colWidths=[CW/3]*3)
    iso_t.setStyle(TableStyle([
        ('BACKGROUND',    (0,0),(-1,-1), NAVY),
        ('TEXTCOLOR',     (0,0),(-1,-1), WHITE),
        ('FONTNAME',      (0,0),(-1,-1), 'Helvetica-Bold'),
        ('FONTSIZE',      (0,0),(-1,-1), 8.5),
        ('ALIGN',         (0,0),(-1,-1), 'CENTER'),
        ('VALIGN',        (0,0),(-1,-1), 'MIDDLE'),
        ('TOPPADDING',    (0,0),(-1,-1), 8),
        ('BOTTOMPADDING', (0,0),(-1,-1), 8),
        ('INNERGRID',     (0,0),(-1,-1), 0.5, GOLD),
        ('BOX',           (0,0),(-1,-1), 1,   GOLD),
    ]))
    st.append(iso_t)

    # ── 06 FINANCIAL STRENGTH ─────────────────────────────────────────────────
    st.append(PageBreak())
    st += sec_block('06', 'Financial Strength', 'Publicly Traded — Proven Fiscal Stability')

    st.append(Paragraph('Revenue Growth 2021–2025', S['h3']))
    rev = [('2021','$2.3B'),('2022','$2.8B'),('2023','$3.1B'),('2024','$3.4B'),('2025','$3.6B')]

    class RevChart(Flowable):
        def __init__(self, data, w):
            super().__init__()
            self.data = data; self.width = w; self.height = 110
        def draw(self):
            c = self.canv
            vals = [float(v.replace('$','').replace(',','').replace('B','')) for _,v in self.data]
            mx = max(vals)*1.08
            max_h = 75.0
            bw = self.width / len(self.data)
            bbar = bw*0.5
            base_y = 22
            for i,(yr,lbl) in enumerate(self.data):
                bh = (vals[i]/mx)*max_h
                bx = i*bw+(bw-bbar)/2
                c.setFillColor(NAVY2)
                c.rect(bx, base_y, bbar, bh, fill=1, stroke=0)
                c.setFillColor(GREEN)
                c.rect(bx, base_y, bbar, bh*0.22, fill=1, stroke=0)
                c.setFillColor(GOLD)
                c.rect(bx, base_y+bh-2, bbar, 2, fill=1, stroke=0)
                c.setFillColor(NAVY)
                c.setFont('Helvetica-Bold', 10)
                c.drawCentredString(bx+bbar/2, base_y+bh+4, lbl)
                c.setFillColor(GOLD2)
                c.setFont('Helvetica-Bold', 7.5)
                c.drawCentredString(bx+bbar/2, 8, yr)
            c.setFillColor(RULE)
            c.rect(0, base_y-1, self.width, 1, fill=1, stroke=0)

    st.append(RevChart(rev, CW))
    st.append(Spacer(1, 12))
    st.append(stat_row([
        ('$500M+', 'Aggregate Bonding'),
        ('$250M',  'Single Project Limit'),
        ('$200M',  'General Liability'),
        ('MYRG',   'NASDAQ Listed'),
        ('4A1',    'D&B Credit Rating'),
        ('Stat.',  'Workers\' Comp'),
    ]))

    # ── 07 EQUIPMENT & RESOURCES ──────────────────────────────────────────────
    st.append(PageBreak())
    st += sec_block('07', 'Equipment & Resources', 'Owned Fleet & Specialized Capabilities')

    fleet_rows = [
        [Paragraph('Equipment Category', S['th']), Paragraph('Qty', S['th']), Paragraph('Notes', S['th'])],
        [Paragraph('Fleet Vehicles & Rolling Stock', S['td']), Paragraph('1,200+', S['tdb']), Paragraph('Light-duty, medium-duty, specialty transport', S['td'])],
        [Paragraph('Cranes (15–300 Ton)', S['td']),           Paragraph('85+',    S['tdb']), Paragraph('Mobile, rough terrain, lattice boom', S['td'])],
        [Paragraph('Cable Pulling Systems', S['td']),          Paragraph('60+',    S['tdb']), Paragraph('High-tension trailers w/ real-time monitoring', S['td'])],
        [Paragraph('Excavators & Trenchers', S['td']),         Paragraph('40+',    S['tdb']), Paragraph('Track excavators, wheel trenchers, mini', S['td'])],
        [Paragraph('Pile Driving Rigs', S['td']),              Paragraph('30+',    S['tdb']), Paragraph('Hydraulic impact and vibratory drivers', S['td'])],
    ]
    ft = Table(fleet_rows, colWidths=[CW*0.44, CW*0.12, CW*0.44])
    ft.setStyle(TableStyle([
        ('BACKGROUND',    (0,0),(-1,0),  NAVY),
        ('ROWBACKGROUNDS',(0,1),(-1,-1), [WHITE, LTBG]),
        ('GRID',          (0,0),(-1,-1), 0.4, RULE),
        ('TOPPADDING',    (0,0),(-1,-1), 5),
        ('BOTTOMPADDING', (0,0),(-1,-1), 5),
        ('LEFTPADDING',   (0,0),(-1,-1), 8),
        ('RIGHTPADDING',  (0,0),(-1,-1), 8),
        ('TEXTCOLOR',     (1,1),(1,-1),  GOLD),
    ]))
    st.append(ft)
    st.append(Spacer(1, 10))
    st.append(Paragraph('Specialized Renewable Energy Equipment', S['h3']))
    spec = [
        'Hydraulic pile drivers for solar tracker and fixed-tilt foundations',
        'Vermeer RTX1250 and Ditch Witch directional boring machines',
        'High-voltage cable pulling trailers with tension monitoring',
        'Transformer handling and rigging equipment (up to 200 tons)',
        'Bucket trucks and aerial lifts (40–125 ft working height)',
        'Underground cable fault location and testing equipment',
        'Megger insulation resistance and hi-pot testing systems',
        'Doble relay test sets for protective relay commissioning',
        'Fiber optic splicing and OTDR testing equipment',
        'GPS-guided grading and trenching systems',
        'Mobile substations for temporary construction power',
        'Thermal imaging cameras for connection verification',
    ]
    mid_idx = len(spec)//2
    sh = (CW-6)/2
    spec_t = Table([[
        [Paragraph(f'• {s}', S['bullet']) for s in spec[:mid_idx]],
        [Paragraph(f'• {s}', S['bullet']) for s in spec[mid_idx:]],
    ]], colWidths=[sh, sh])
    spec_t.setStyle(TableStyle([
        ('VALIGN',(0,0),(-1,-1),'TOP'),
        ('LEFTPADDING',(0,0),(-1,-1),0),('RIGHTPADDING',(0,0),(-1,-1),0),
        ('TOPPADDING',(0,0),(-1,-1),0),('BOTTOMPADDING',(0,0),(-1,-1),0),
    ]))
    st.append(spec_t)

    # ── 08 CLIENT REFERENCES ──────────────────────────────────────────────────
    st.append(PageBreak())
    st += sec_block('08', 'Client References', 'Contacts Available Upon Request')

    refs = [
        ('Michael Andersen','Clearwater Energy Partners — Director of Construction',
         'Provided oversight of CSI Electrical\'s work on the Sagebrush Flats 350 MW Solar Energy Center '
         'in Washington. Scope included full electrical BOP, substation construction, and grid interconnection. '
         'Project completed on schedule and within budget.'),
        ('Jennifer Nakamura','Ridgeline Power Corporation — VP of Engineering',
         'Managed Ridgeline\'s relationship with CSI Electrical across two wind farm collector system projects '
         'totaling 430 MW in Oregon. Recognized CSI for outstanding safety performance, crew quality, and '
         'proactive schedule management throughout both projects.'),
        ('Thomas Whitfield','Basin Range Storage Development — Project Director',
         'Oversaw CSI Electrical\'s execution of the Silver Creek 200 MW / 800 MWh BESS facility in Idaho. '
         'Commended the team for technical expertise in battery storage electrical systems, attention to '
         'quality, and ability to resolve constructability challenges in the field.'),
    ]
    for nm, co, body in refs:
        init = ''.join(w[0] for w in nm.split())
        av = Table([[Paragraph(init, S['avatar'])]], colWidths=[0.55*IN])
        av.setStyle(TableStyle([
            ('BACKGROUND',(0,0),(-1,-1),GREEN),
            ('TOPPADDING',(0,0),(-1,-1),9),('BOTTOMPADDING',(0,0),(-1,-1),9),
            ('ALIGN',(0,0),(-1,-1),'CENTER'),
        ]))
        card = Table([[av, [Paragraph(nm, S['ref_nm']), Paragraph(co, S['ref_co']),
                             Paragraph(body, S['ref_b']),
                             Paragraph('Phone: [Redacted]  |  Email: [Redacted]', S['note'])]]],
                     colWidths=[0.70*IN, CW - 0.70*IN])
        card.setStyle(TableStyle([
            ('VALIGN',(0,0),(-1,-1),'TOP'),
            ('LEFTPADDING',(0,0),(-1,-1),10),('RIGHTPADDING',(0,0),(-1,-1),10),
            ('TOPPADDING',(0,0),(-1,-1),10),('BOTTOMPADDING',(0,0),(-1,-1),10),
            ('BOX',(0,0),(-1,-1),0.5,RULE),
            ('BACKGROUND',(0,0),(-1,-1),LTBG),
        ]))
        st.append(KeepTogether(card))
        st.append(Spacer(1, 8))

    # ── 09 CERTIFICATIONS ─────────────────────────────────────────────────────
    st.append(PageBreak())
    st += sec_block('09', 'Certifications & Memberships', 'Industry Affiliations & Professional Standards')

    certs = [
        ('NECA',       'Natl Electrical\nContractors Assoc.'),
        ('IBEW',       'Intl Brotherhood of\nElectrical Workers'),
        ('OSHA VPP',   'Voluntary Protection\nPrograms Star Status'),
        ('ISNetworld', 'RAVS Verified\nGrade A Compliance'),
        ('ISO 9001',   'Quality Management\nSystem Certified'),
        ('ISO 14001',  'Environmental Mgmt\nSystem Certified'),
        ('ISO 45001',  'Occupational H&S\nSystem Certified'),
        ('NFPA',       'Natl Fire Protection\nAssociation'),
        ('SEIA',       'Solar Energy\nIndustries Assoc.'),
        ('ACP',        'American Clean\nPower Association'),
        ('ABC / AGC',  'Associated Builders\n& General Contractors'),
        ('NASDAQ MYRG','Publicly Traded\nElectrical Contractor'),
    ]
    ncols = 4
    ccw = CW / ncols
    cert_rows = []
    for i in range(0, len(certs), ncols):
        row = []
        for c2 in certs[i:i+ncols]:
            row.append([
                Paragraph(f'<b>{c2[0]}</b>',
                          ParagraphStyle('ca', fontName='Helvetica-Bold', fontSize=11,
                                         textColor=GOLD, leading=14, alignment=TA_CENTER)),
                Paragraph(c2[1],
                          ParagraphStyle('cs', fontName='Helvetica', fontSize=7.5,
                                         textColor=TSUB, leading=10, alignment=TA_CENTER)),
            ])
        while len(row) < ncols:
            row.append([Spacer(1,1)])
        cert_rows.append(row)
    ct = Table(cert_rows, colWidths=[ccw]*ncols)
    ct.setStyle(TableStyle([
        ('BOX',           (0,0),(-1,-1), 0.5, RULE),
        ('INNERGRID',     (0,0),(-1,-1), 0.5, RULE),
        ('ROWBACKGROUNDS',(0,0),(-1,-1), [WHITE, LTBG]),
        ('TOPPADDING',    (0,0),(-1,-1), 10),
        ('BOTTOMPADDING', (0,0),(-1,-1), 10),
        ('ALIGN',         (0,0),(-1,-1), 'CENTER'),
        ('VALIGN',        (0,0),(-1,-1), 'MIDDLE'),
    ]))
    st.append(ct)
    st.append(Spacer(1, 16))

    closing = Table([[Paragraph(
        '<b>MYR Group Inc. / CSI Electrical Contractors</b> stands ready to serve as your preferred '
        'electrical construction partner for renewable energy development throughout the Pacific Northwest '
        'and beyond. We bring the financial strength of a publicly traded parent company, the regional '
        'expertise of a dedicated renewable energy subsidiary, and the craft excellence of a 100% '
        'IBEW-affiliated workforce to every project we undertake.',
        ParagraphStyle('cl', fontName='Helvetica', fontSize=9.5, textColor=WHITE,
                       leading=14, alignment=TA_JUSTIFY)
    )]], colWidths=[CW])
    closing.setStyle(TableStyle([
        ('BACKGROUND',    (0,0),(-1,-1), NAVY),
        ('TOPPADDING',    (0,0),(-1,-1), 14),
        ('BOTTOMPADDING', (0,0),(-1,-1), 14),
        ('LEFTPADDING',   (0,0),(-1,-1), 16),
        ('RIGHTPADDING',  (0,0),(-1,-1), 16),
        ('BOX',           (0,0),(-1,-1), 1.5, GOLD),
    ]))
    st.append(closing)

    return st


# ── MAIN ──────────────────────────────────────────────────────────────────────
if __name__ == '__main__':
    out = '/Users/mxmbp/portfolio-temp/CedarRidge_SOQ_MYRGroup_CSI.pdf'
    os.makedirs(os.path.dirname(out), exist_ok=True)
    doc = SOQDoc(out)
    doc.build(build_story())
    print(f'Done: {out}')
