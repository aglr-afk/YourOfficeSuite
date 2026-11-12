"""Builds the starter templates that ship with YourOfficeSuite (own designs)."""
import json, os, shutil, datetime
import docx
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, Reference
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import DataBarRule, CellIsRule
from openpyxl.utils import get_column_letter
from pptx import Presentation
from pptx.util import Inches, Pt as PPt
from pptx.dml.color import RGBColor as PRGB
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

OUT = os.path.join(os.path.dirname(__file__), '..', 'shell', 'templates')
INK, VIOLET, ROSE, AQUA, AMBER, MUTED, SOFT, LINE = '14121C', '7C5CFC', 'F0568C', '0FB2C9', 'F59E0B', '6B6784', 'F3F0FB', 'E7E3F3'
manifest = []

def add(kind, id_, name, desc):
    manifest.append({'id': id_, 'kind': kind, 'name': name, 'desc': desc, 'file': f'{kind}/{id_}.{kind}', 'thumb': f'thumbs/{id_}.jpg'})
    return os.path.join(OUT, kind, f'{id_}.{kind}')

# ------------------------------------------------------------------ Word
def rgb(h): return RGBColor.from_string(h)

def base_doc(font='Calibri', size=11, margins=2.2):
    d = docx.Document()
    st = d.styles['Normal']; st.font.name = font; st.font.size = Pt(size); st.element.rPr.rFonts.set(qn('w:eastAsia'), font)
    st.paragraph_format.space_after = Pt(6); st.paragraph_format.line_spacing = 1.15
    for s in d.sections:
        s.left_margin = s.right_margin = Cm(margins); s.top_margin = s.bottom_margin = Cm(2)
    return d

def para(d, text='', size=None, bold=False, color=None, align=None, after=None, before=None, italic=False, caps=False):
    p = d.add_paragraph(); r = p.add_run(text.upper() if caps else text)
    if size: r.font.size = Pt(size)
    r.bold = bold; r.italic = italic
    if color: r.font.color.rgb = rgb(color)
    if align: p.alignment = align
    if after is not None: p.paragraph_format.space_after = Pt(after)
    if before is not None: p.paragraph_format.space_before = Pt(before)
    return p

def rule(p, color=VIOLET, sz=12):
    pPr = p._p.get_or_add_pPr(); b = OxmlElement('w:pBdr'); bt = OxmlElement('w:bottom')
    bt.set(qn('w:val'), 'single'); bt.set(qn('w:sz'), str(sz)); bt.set(qn('w:space'), '4'); bt.set(qn('w:color'), color); b.append(bt); pPr.append(b)

def heading(d, text, color=VIOLET):
    p = para(d, text, size=11, bold=True, color=color, before=14, after=6, caps=True); rule(p, LINE, 6); return p

def shade(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr(); s = OxmlElement('w:shd'); s.set(qn('w:val'), 'clear'); s.set(qn('w:color'), 'auto'); s.set(qn('w:fill'), fill); tcPr.append(s)

def table(d, header, rows, widths=None, head_fill=INK):
    t = d.add_table(rows=1, cols=len(header)); t.style = 'Table Grid'; t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(header):
        c = t.rows[0].cells[i]; c.text = ''; r = c.paragraphs[0].add_run(h); r.bold = True; r.font.color.rgb = rgb('FFFFFF'); r.font.size = Pt(10); shade(c, head_fill)
    for ri, row in enumerate(rows):
        cells = t.add_row().cells
        for i, v in enumerate(row):
            cells[i].text = str(v)
            if ri % 2: shade(cells[i], SOFT)
    if widths:
        for row in t.rows:
            for i, w in enumerate(widths): row.cells[i].width = Cm(w)
    d.add_paragraph()
    return t

def bullets(d, items):
    for it in items:
        p = d.add_paragraph(it, style='List Bullet'); p.paragraph_format.space_after = Pt(2)

def word_resume():
    d = base_doc()
    para(d, 'Your Name', size=30, bold=True, color=INK, after=0)
    para(d, 'Job title you are aiming for', size=13, color=VIOLET, after=4)
    p = para(d, 'City, Country   |   you@email.com   |   +00 000 000 000   |   linkedin.com/in/you', size=10, color=MUTED, after=10); rule(p, VIOLET, 14)
    heading(d, 'Profile'); para(d, 'Two or three lines on who you are, what you are good at and the kind of role you want next. Keep it specific and write it for the job you are applying to.')
    heading(d, 'Experience')
    for role, where, when in [('Job title', 'Company name, City', '2022 to present'), ('Earlier job title', 'Company name, City', '2019 to 2022')]:
        p = d.add_paragraph(); r = p.add_run(role); r.bold = True; r.font.size = Pt(12); p.add_run(f'   {where}').font.color.rgb = rgb(MUTED); p.paragraph_format.space_after = Pt(0)
        para(d, when, size=10, color=MUTED, after=3, italic=True)
        bullets(d, ['Start with a strong verb and say what you did.', 'Add a number where you can, such as time saved or revenue gained.', 'Mention the tools or skills you used.'])
    heading(d, 'Education'); p = d.add_paragraph(); p.add_run('Degree or course name').bold = True; p.add_run('   School or university, year').font.color.rgb = rgb(MUTED)
    heading(d, 'Skills'); para(d, 'Skill one   •   Skill two   •   Skill three   •   Skill four   •   Skill five')
    d.save(add('docx', 'resume', 'Resume', 'Clean one-page CV'))

def word_letter():
    d = base_doc(margins=2.6)
    p = para(d, 'Your Name or Company', size=18, bold=True, color=INK, after=0);
    p = para(d, 'Street address, City, Postcode   |   you@email.com   |   +00 000 000 000', size=9.5, color=MUTED, after=22); rule(p, VIOLET, 12)
    para(d, datetime.date.today().strftime('%d %B %Y'), color=MUTED, after=14)
    for l in ['Recipient name', 'Their job title', 'Company name', 'Street address, City, Postcode']: para(d, l, after=0)
    para(d, 'Subject: What this letter is about', bold=True, before=18, after=12)
    para(d, 'Dear Recipient name,', after=10)
    for t in ['Open with the reason you are writing. One or two sentences are enough.',
              'Use the middle of the letter for the details the reader needs: what happened, what you are asking for, and any dates or amounts.',
              'Close by saying what you would like to happen next and how they can reach you.']: para(d, t, after=10)
    para(d, 'Yours sincerely,', before=10, after=34); para(d, 'Your Name', bold=True, after=0); para(d, 'Your job title', color=MUTED)
    d.save(add('docx', 'letter', 'Business letter', 'Formal letter with letterhead'))

def word_report():
    d = base_doc()
    para(d, 'REPORT', size=10, bold=True, color=VIOLET, before=60, after=6)
    para(d, 'Title of your report', size=34, bold=True, color=INK, after=6)
    p = para(d, 'A short subtitle that says what the reader will learn', size=14, color=MUTED, after=30); rule(p, VIOLET, 18)
    para(d, f'Prepared by Your Name   |   {datetime.date.today().strftime("%B %Y")}', color=MUTED, after=0)
    d.add_page_break()
    heading(d, 'Summary'); para(d, 'Sum up the whole report in one short paragraph. A busy reader should be able to stop here and still know the main finding and what you recommend.')
    heading(d, 'Background'); para(d, 'Explain why this report was written and what question it answers. Mention the period and the data it covers.')
    heading(d, 'Findings'); bullets(d, ['First finding, with the number that backs it up.', 'Second finding, with the number that backs it up.', 'Third finding, with the number that backs it up.'])
    d.add_paragraph()
    table(d, ['Measure', 'Last period', 'This period', 'Change'], [['Measure one', '100', '120', '+20%'], ['Measure two', '80', '76', '-5%'], ['Measure three', '45', '60', '+33%']])
    heading(d, 'Recommendations'); bullets(d, ['What should happen next, and who should do it.', 'What it will cost or save.', 'When it should be done by.'])
    d.save(add('docx', 'report', 'Report', 'Cover page, findings and table'))

def word_notes():
    d = base_doc()
    para(d, 'Meeting notes', size=26, bold=True, color=INK, after=2); p = para(d, 'Project or team name', size=13, color=VIOLET, after=12); rule(p, VIOLET, 12)
    t = d.add_table(rows=3, cols=2); t.style = 'Table Grid'
    for i, (k, v) in enumerate([('Date and time', datetime.date.today().strftime('%d %B %Y') + ', 10:00'), ('Attendees', 'Name, Name, Name'), ('Note taker', 'Name')]):
        t.cell(i, 0).text = k; t.cell(i, 1).text = v; shade(t.cell(i, 0), SOFT); t.cell(i, 0).paragraphs[0].runs[0].bold = True; t.cell(i, 0).width = Cm(4); t.cell(i, 1).width = Cm(12.5)
    heading(d, 'Agenda'); bullets(d, ['First topic', 'Second topic', 'Third topic'])
    heading(d, 'Discussion'); para(d, 'Write the key points for each topic. Keep to what was decided and why, not who said what.')
    heading(d, 'Decisions'); bullets(d, ['Decision one', 'Decision two'])
    heading(d, 'Action items'); table(d, ['Action', 'Owner', 'Due'], [['What needs to be done', 'Name', 'Date'], ['What needs to be done', 'Name', 'Date'], ['What needs to be done', 'Name', 'Date']], widths=[10, 3.5, 3])
    d.save(add('docx', 'meeting-notes', 'Meeting notes', 'Agenda, decisions, action items'))

def word_proposal():
    d = base_doc()
    para(d, 'PROPOSAL', size=10, bold=True, color=ROSE, after=4); para(d, 'Project name', size=30, bold=True, color=INK, after=4)
    p = para(d, 'Prepared for Client name by Your company', size=12, color=MUTED, after=16); rule(p, ROSE, 14)
    heading(d, 'The goal', ROSE); para(d, 'Describe in plain words what the client wants to achieve and why it matters to them.')
    heading(d, 'Our approach', ROSE); bullets(d, ['Step one: what you will do first.', 'Step two: what comes next.', 'Step three: how you will finish and hand over.'])
    heading(d, 'Timeline and cost', ROSE)
    table(d, ['Phase', 'What is included', 'Duration', 'Cost'], [['Discovery', 'Workshops and plan', '2 weeks', '0.00'], ['Build', 'Design and delivery', '6 weeks', '0.00'], ['Launch', 'Testing and handover', '2 weeks', '0.00']], widths=[3, 7.5, 3, 3])
    heading(d, 'Next steps', ROSE); para(d, 'Say how the client can accept the proposal, and how long the offer is valid.')
    d.save(add('docx', 'proposal', 'Proposal', 'Goal, approach, timeline and cost'))

# ------------------------------------------------------------------ Excel
thin = Side(style='thin', color=LINE); BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
def xhead(ws, row, headers, fill=INK, col=1):
    for i, h in enumerate(headers):
        c = ws.cell(row=row, column=col + i, value=h); c.font = Font(bold=True, color='FFFFFF', name='Calibri', size=11)
        c.fill = PatternFill('solid', fgColor=fill); c.alignment = Alignment(horizontal='center', vertical='center'); c.border = BORDER
    ws.row_dimensions[row].height = 22
def xtitle(ws, text, sub=None):
    ws['A1'] = text; ws['A1'].font = Font(bold=True, size=20, color=INK, name='Calibri'); ws.row_dimensions[1].height = 32
    if sub: ws['A2'] = sub; ws['A2'].font = Font(size=11, color=MUTED, name='Calibri')
def widths(ws, ws_):
    for i, w in enumerate(ws_): ws.column_dimensions[get_column_letter(i + 1)].width = w
def box(ws, rng, fmt=None, fill=None, bold=False):
    for row in ws[rng]:
        for c in row:
            c.border = BORDER
            if fmt: c.number_format = fmt
            if fill: c.fill = PatternFill('solid', fgColor=fill)
            if bold: c.font = Font(bold=True, name='Calibri')

def xl_budget():
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = 'Budget'; xtitle(ws, 'Monthly budget', 'Type your own numbers in the Planned and Actual columns')
    xhead(ws, 4, ['Income', 'Planned', 'Actual', 'Difference'], VIOLET)
    inc = [('Salary', 3000, 3000), ('Side income', 400, 250), ('Other', 100, 0)]
    for i, (n, p, a) in enumerate(inc):
        r = 5 + i; ws.cell(r, 1, n); ws.cell(r, 2, p); ws.cell(r, 3, a); ws.cell(r, 4, f'=C{r}-B{r}')
    ws['A8'] = 'Total income'; ws['B8'] = '=SUM(B5:B7)'; ws['C8'] = '=SUM(C5:C7)'; ws['D8'] = '=C8-B8'
    xhead(ws, 10, ['Expenses', 'Planned', 'Actual', 'Difference'], ROSE)
    exp = [('Rent', 1100, 1100), ('Food', 450, 510), ('Transport', 150, 120), ('Bills', 220, 240), ('Fun', 200, 260), ('Savings', 500, 400)]
    for i, (n, p, a) in enumerate(exp):
        r = 11 + i; ws.cell(r, 1, n); ws.cell(r, 2, p); ws.cell(r, 3, a); ws.cell(r, 4, f'=B{r}-C{r}')
    ws['A17'] = 'Total expenses'; ws['B17'] = '=SUM(B11:B16)'; ws['C17'] = '=SUM(C11:C16)'; ws['D17'] = '=B17-C17'
    ws['A19'] = 'Left over'; ws['B19'] = '=B8-B17'; ws['C19'] = '=C8-C17'
    box(ws, 'A5:D8', '#,##0.00'); box(ws, 'A11:D17', '#,##0.00'); box(ws, 'A19:C19', '#,##0.00', SOFT, True)
    for r in (8, 17):
        for c in 'ABCD': ws[f'{c}{r}'].font = Font(bold=True, name='Calibri'); ws[f'{c}{r}'].fill = PatternFill('solid', fgColor=SOFT)
    widths(ws, [24, 14, 14, 14]); ws.freeze_panes = 'A4'
    ch = BarChart(); ch.type = 'col'; ch.title = 'Planned and actual spending'; ch.height = 8.5; ch.width = 15
    ch.add_data(Reference(ws, min_col=2, max_col=3, min_row=10, max_row=16), titles_from_data=True); ch.set_categories(Reference(ws, min_col=1, min_row=11, max_row=16)); ws.add_chart(ch, 'F4')
    wb.save(add('xlsx', 'budget', 'Monthly budget', 'Income, spending and a chart'))

def xl_invoice():
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = 'Invoice'
    ws['A1'] = 'INVOICE'; ws['A1'].font = Font(bold=True, size=26, color=VIOLET, name='Calibri'); ws.row_dimensions[1].height = 38
    ws['A2'] = 'Your company name'; ws['A2'].font = Font(bold=True, size=12); ws['A3'] = 'Street address, City, Postcode'; ws['A4'] = 'you@email.com'
    ws['D2'] = 'Invoice no.'; ws['E2'] = 'INV-001'; ws['D3'] = 'Date'; ws['E3'] = datetime.date.today(); ws['E3'].number_format = 'DD MMM YYYY'; ws['D4'] = 'Due'; ws['E4'] = '=E3+14'; ws['E4'].number_format = 'DD MMM YYYY'
    for c in ('D2', 'D3', 'D4'): ws[c].font = Font(color=MUTED); ws[c].alignment = Alignment(horizontal='right')
    ws['A6'] = 'Bill to'; ws['A6'].font = Font(bold=True, color=VIOLET); ws['A7'] = 'Client name'; ws['A8'] = 'Client address'
    xhead(ws, 10, ['Description', 'Qty', 'Unit price', 'Tax %', 'Amount'])
    for i, (d_, q, p) in enumerate([('Item or service one', 2, 150), ('Item or service two', 1, 320), ('Item or service three', 5, 40), ('', None, None), ('', None, None)]):
        r = 11 + i; ws.cell(r, 1, d_); ws.cell(r, 2, q); ws.cell(r, 3, p); ws.cell(r, 4, 0 if q else None); ws.cell(r, 5, f'=IF(B{r}="","",B{r}*C{r})')
    box(ws, 'A11:E15');
    for r in range(11, 16): ws[f'C{r}'].number_format = '#,##0.00'; ws[f'E{r}'].number_format = '#,##0.00'
    ws['D17'] = 'Subtotal'; ws['E17'] = '=SUM(E11:E15)'; ws['D18'] = 'Tax'; ws['E18'] = '=SUMPRODUCT(B11:B15,C11:C15,D11:D15)/100'; ws['D19'] = 'Total due'; ws['E19'] = '=E17+E18'
    box(ws, 'D17:E19', '#,##0.00'); ws['D19'].font = Font(bold=True, color='FFFFFF'); ws['E19'].font = Font(bold=True, color='FFFFFF'); ws['D19'].fill = ws['E19'].fill = PatternFill('solid', fgColor=VIOLET)
    ws['A21'] = 'Payment details'; ws['A21'].font = Font(bold=True, color=VIOLET); ws['A22'] = 'Bank name, account name, account number'; ws['A24'] = 'Thank you for your business.'; ws['A24'].font = Font(italic=True, color=MUTED)
    widths(ws, [38, 9, 14, 10, 16]); wb.save(add('xlsx', 'invoice', 'Invoice', 'Totals and tax work themselves out'))

def xl_tracker():
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = 'Tasks'; xtitle(ws, 'Project tracker', 'Pick a status from the list in column C')
    xhead(ws, 4, ['Task', 'Owner', 'Status', 'Start', 'Due', 'Done %', 'Notes'], AQUA)
    today = datetime.date.today()
    rows = [('Plan the project', 'Name', 'Done', 0, 5, 1), ('Agree the budget', 'Name', 'Done', 2, 8, 1), ('Design', 'Name', 'In progress', 7, 21, .6), ('Build', 'Name', 'In progress', 14, 40, .25), ('Test', 'Name', 'Not started', 38, 50, 0), ('Launch', 'Name', 'Not started', 50, 55, 0)]
    for i, (t, o, s, a, b, p) in enumerate(rows):
        r = 5 + i; ws.cell(r, 1, t); ws.cell(r, 2, o); ws.cell(r, 3, s); ws.cell(r, 4, today + datetime.timedelta(days=a)); ws.cell(r, 5, today + datetime.timedelta(days=b)); ws.cell(r, 6, p)
    box(ws, 'A5:G24');
    for r in range(5, 25): ws[f'D{r}'].number_format = ws[f'E{r}'].number_format = 'DD MMM'; ws[f'F{r}'].number_format = '0%'
    dv = DataValidation(type='list', formula1='"Not started,In progress,Blocked,Done"', allow_blank=True); ws.add_data_validation(dv); dv.add('C5:C24')
    ws.conditional_formatting.add('F5:F24', DataBarRule(start_type='num', start_value=0, end_type='num', end_value=1, color=AQUA))
    ws.conditional_formatting.add('C5:C24', CellIsRule(operator='equal', formula=['"Done"'], fill=PatternFill('solid', bgColor='D9F5E9'), font=Font(color='0E7A52')))
    ws.conditional_formatting.add('C5:C24', CellIsRule(operator='equal', formula=['"Blocked"'], fill=PatternFill('solid', bgColor='FDE2E8'), font=Font(color='B0233F')))
    ws['A26'] = 'Tasks done'; ws['B26'] = '=COUNTIF(C5:C24,"Done")'; ws['A27'] = 'Tasks open'; ws['B27'] = '=COUNTA(A5:A24)-B26'; box(ws, 'A26:B27', None, SOFT, True)
    widths(ws, [30, 14, 15, 12, 12, 12, 34]); ws.freeze_panes = 'A5'; ws.auto_filter.ref = 'A4:G24'
    wb.save(add('xlsx', 'project-tracker', 'Project tracker', 'Tasks, owners, status, progress'))

def xl_timesheet():
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = 'Timesheet'; xtitle(ws, 'Weekly timesheet', 'Name:                                Week starting:')
    xhead(ws, 4, ['Day', 'Start', 'Finish', 'Break (hours)', 'Hours worked', 'Notes'], AMBER)
    for i, dname in enumerate(['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']):
        r = 5 + i; ws.cell(r, 1, dname)
        if i < 5: ws.cell(r, 2, datetime.time(9, 0)); ws.cell(r, 3, datetime.time(17, 30)); ws.cell(r, 4, 0.5)
        ws.cell(r, 5, f'=IF(OR(B{r}="",C{r}=""),0,(C{r}-B{r})*24-D{r})'); ws[f'B{r}'].number_format = ws[f'C{r}'].number_format = 'HH:MM'; ws[f'E{r}'].number_format = '0.00'
    box(ws, 'A5:F11'); ws['D13'] = 'Total hours'; ws['E13'] = '=SUM(E5:E11)'; ws['D14'] = 'Hourly rate'; ws['E14'] = 20; ws['D15'] = 'Pay'; ws['E15'] = '=E13*E14'
    box(ws, 'D13:E15', '#,##0.00', SOFT, True); widths(ws, [16, 11, 11, 15, 15, 36]); wb.save(add('xlsx', 'timesheet', 'Timesheet', 'Hours and pay for the week'))

def xl_todo():
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = 'To do'; xtitle(ws, 'To-do list', 'Type x in the Done column to tick something off')
    xhead(ws, 4, ['Done', 'Task', 'Priority', 'Due', 'Notes'], ROSE)
    for i, (t, p) in enumerate([('First thing to do', 'High'), ('Second thing to do', 'Medium'), ('Third thing to do', 'Low')]): ws.cell(5 + i, 2, t); ws.cell(5 + i, 3, p)
    box(ws, 'A5:E30'); dv = DataValidation(type='list', formula1='"High,Medium,Low"', allow_blank=True); ws.add_data_validation(dv); dv.add('C5:C30')
    for r in range(5, 31): ws[f'A{r}'].alignment = Alignment(horizontal='center'); ws[f'D{r}'].number_format = 'DD MMM'
    ws.conditional_formatting.add('C5:C30', CellIsRule(operator='equal', formula=['"High"'], font=Font(color='B0233F', bold=True)))
    ws['G4'] = 'Done'; ws['H4'] = '=COUNTIF(A5:A30,"x")'; ws['G5'] = 'Still to do'; ws['H5'] = '=COUNTA(B5:B30)-H4'; box(ws, 'G4:H5', None, SOFT, True)
    widths(ws, [8, 40, 12, 12, 36, 3, 14, 8]); ws.freeze_panes = 'A5'; wb.save(add('xlsx', 'todo', 'To-do list', 'Tasks with priority and dates'))

# ------------------------------------------------------------------ PowerPoint
def prs16():
    p = Presentation(); p.slide_width = Inches(13.333); p.slide_height = Inches(7.5); return p
def slide(p, bg=None):
    s = p.slides.add_slide(p.slide_layouts[6])
    if bg: f = s.background.fill; f.solid(); f.fore_color.rgb = PRGB.from_string(bg)
    return s
def rect(s, x, y, w, h, fill, shape=MSO_SHAPE.RECTANGLE):
    r = s.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h)); r.fill.solid(); r.fill.fore_color.rgb = PRGB.from_string(fill); r.line.fill.background(); r.shadow.inherit = False; return r
def text(s, x, y, w, h, t, size=18, bold=False, color=INK, align=PP_ALIGN.LEFT, font='Calibri', anchor=MSO_ANCHOR.TOP):
    tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h)); tf = tb.text_frame; tf.word_wrap = True; tf.vertical_anchor = anchor
    lines = t if isinstance(t, list) else [t]
    for i, line in enumerate(lines):
        par = tf.paragraphs[0] if i == 0 else tf.add_paragraph(); par.alignment = align; r = par.add_run(); r.text = line
        r.font.size = PPt(size); r.font.bold = bold; r.font.color.rgb = PRGB.from_string(color); r.font.name = font; par.space_after = PPt(size * .45)
    return tb

def ppt_pitch():
    p = prs16()
    s = slide(p, INK); rect(s, 9.2, -1.2, 5.6, 5.6, VIOLET, MSO_SHAPE.OVAL); rect(s, 10.8, 3.4, 3.6, 3.6, ROSE, MSO_SHAPE.OVAL)
    text(s, .9, 2.3, 8, 1.6, 'Company name', 54, True, 'FFFFFF'); text(s, .9, 3.9, 8, 1, 'One line that says what you do and for whom', 22, False, 'CFCBE2'); text(s, .9, 6.3, 6, .5, 'Investor presentation  |  Month Year', 13, False, '8F8AA8')
    def content(kicker, title, body, accent=VIOLET):
        s = slide(p, 'FFFFFF'); rect(s, 0, 0, .28, 7.5, accent); text(s, .9, .65, 8, .4, kicker.upper(), 13, True, accent); text(s, .9, 1.05, 11.4, 1.2, title, 36, True, INK); text(s, .9, 2.6, 11.2, 4, body, 20, False, '3A3653'); return s
    content('The problem', 'What is broken today', ['Describe the pain your customer feels.', 'Say how often it happens and what it costs them.', 'Explain why the current options fall short.'], ROSE)
    content('Our solution', 'How we fix it', ['Say what your product does in one sentence.', 'List the two or three things that make it better.', 'Show what the customer gains.'])
    s = content('Market', 'How big the opportunity is', [], AQUA)
    for i, (n, l) in enumerate([('00 bn', 'Total market'), ('00 bn', 'Market we can serve'), ('00 m', 'Market we can win first')]):
        rect(s, .9 + i * 3.9, 2.8, 3.5, 2.6, SOFT, MSO_SHAPE.ROUNDED_RECTANGLE); text(s, .9 + i * 3.9, 3.2, 3.5, 1, n, 44, True, VIOLET, PP_ALIGN.CENTER); text(s, .9 + i * 3.9, 4.4, 3.5, .6, l, 16, False, MUTED, PP_ALIGN.CENTER)
    content('Traction', 'What we have achieved so far', ['Customers, revenue or users, with the growth rate.', 'Partnerships and pilots.', 'What customers say about you.'], AMBER)
    s = content('Team', 'The people behind it', [])
    for i, n in enumerate(['Name', 'Name', 'Name', 'Name']):
        rect(s, 1.2 + i * 2.9, 2.8, 1.7, 1.7, [VIOLET, ROSE, AQUA, AMBER][i], MSO_SHAPE.OVAL); text(s, .75 + i * 2.9, 4.7, 2.6, .5, n, 18, True, INK, PP_ALIGN.CENTER); text(s, .75 + i * 2.9, 5.15, 2.6, .5, 'Role', 14, False, MUTED, PP_ALIGN.CENTER)
    s = slide(p, INK); rect(s, -1.5, 4.2, 5, 5, VIOLET, MSO_SHAPE.OVAL); text(s, .9, 2.2, 11.5, 1.4, 'We are raising 0.0 m', 50, True, 'FFFFFF'); text(s, .9, 3.7, 11, 1.6, ['What the money will be used for.', 'you@email.com'], 22, False, 'CFCBE2')
    p.save(add('pptx', 'pitch-deck', 'Pitch deck', 'Seven slides for investors'))

def ppt_report():
    p = prs16()
    s = slide(p, 'FFFFFF'); rect(s, 0, 0, 13.333, 2.6, SOFT); rect(s, .9, 2.45, 1.6, .12, VIOLET)
    text(s, .9, 3.0, 11, 1.4, 'Quarterly business review', 46, True, INK); text(s, .9, 4.4, 11, .8, 'Team or department  |  Quarter and year', 20, False, MUTED)
    def page(title):
        s = slide(p, 'FFFFFF'); text(s, .9, .55, 11.5, 1, title, 32, True, INK); rect(s, .9, 1.45, 1.2, .08, VIOLET); return s
    s = page('Highlights')
    for i, (n, l, c) in enumerate([('00%', 'Growth on last quarter', VIOLET), ('000', 'New customers', ROSE), ('00', 'Projects delivered', AQUA), ('0.0', 'Customer rating', AMBER)]):
        rect(s, .9 + i * 2.95, 2.3, 2.7, 2.7, 'FFFFFF', MSO_SHAPE.ROUNDED_RECTANGLE).line.color.rgb = PRGB.from_string(LINE); rect(s, .9 + i * 2.95, 2.3, 2.7, .14, c)
        text(s, .9 + i * 2.95, 2.9, 2.7, 1, n, 44, True, INK, PP_ALIGN.CENTER); text(s, .9 + i * 2.95, 4.05, 2.7, .8, l, 14, False, MUTED, PP_ALIGN.CENTER)
    s = page('What went well'); text(s, .9, 2.1, 11.4, 4.5, ['First win, with the result it produced.', 'Second win, with the result it produced.', 'Third win, with the result it produced.'], 22, False, '3A3653')
    s = page('What we will improve'); text(s, .9, 2.1, 11.4, 4.5, ['First issue, and what we are doing about it.', 'Second issue, and what we are doing about it.'], 22, False, '3A3653')
    s = page('Plan for next quarter')
    for i, (h, b) in enumerate([('Goal one', 'How we will measure it'), ('Goal two', 'How we will measure it'), ('Goal three', 'How we will measure it')]):
        rect(s, .9, 2.2 + i * 1.5, .14, 1.15, [VIOLET, ROSE, AQUA][i]); text(s, 1.3, 2.2 + i * 1.5, 10, .55, h, 22, True, INK); text(s, 1.3, 2.75 + i * 1.5, 10, .5, b, 16, False, MUTED)
    s = slide(p, INK); text(s, .9, 2.9, 11.5, 1.2, 'Thank you', 54, True, 'FFFFFF'); text(s, .9, 4.2, 11, .8, 'Questions and discussion', 22, False, 'CFCBE2')
    p.save(add('pptx', 'business-review', 'Business review', 'Highlights, wins, plan'))

def ppt_minimal():
    p = prs16()
    s = slide(p, 'FBFAFF'); rect(s, .9, 3.05, .9, .1, ROSE); text(s, .9, 3.35, 11, 1.3, 'Presentation title', 50, True, INK); text(s, .9, 4.7, 11, .7, 'Your name  |  Date', 20, False, MUTED)
    for t, b in [('First section', ['A short point.', 'Another short point.', 'One more short point.']), ('Second section', ['A short point.', 'Another short point.']), ('Third section', ['A short point.', 'Another short point.', 'One more short point.'])]:
        s = slide(p, 'FBFAFF'); text(s, .9, .7, 11.5, 1, t, 34, True, INK); rect(s, .9, 1.6, .7, .08, ROSE); text(s, .9, 2.3, 11.2, 4.4, b, 24, False, '3A3653')
    s = slide(p, 'FBFAFF'); text(s, .9, 3.0, 11.5, 1.2, 'Thank you', 50, True, INK); rect(s, .9, 4.3, .9, .1, ROSE)
    p.save(add('pptx', 'minimal', 'Minimal', 'Light and simple'))

if __name__ == '__main__':
    for k in ('docx', 'xlsx', 'pptx'):
        shutil.rmtree(os.path.join(OUT, k), ignore_errors=True); os.makedirs(os.path.join(OUT, k))
    os.makedirs(os.path.join(OUT, 'thumbs'), exist_ok=True)
    d = base_doc(); d.save(os.path.join(OUT, 'blank.docx'))
    wb = openpyxl.Workbook(); wb.active.title = 'Sheet1'; wb.save(os.path.join(OUT, 'blank.xlsx'))
    p = prs16(); p.slides.add_slide(p.slide_layouts[0]); p.save(os.path.join(OUT, 'blank.pptx'))
    word_resume(); word_letter(); word_report(); word_notes(); word_proposal()
    xl_budget(); xl_invoice(); xl_tracker(); xl_timesheet(); xl_todo()
    ppt_pitch(); ppt_report(); ppt_minimal()
    json.dump(manifest, open(os.path.join(OUT, 'manifest.json'), 'w'), indent=1)
    print(len(manifest), 'templates')
