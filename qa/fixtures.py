import docx, openpyxl, shutil, os, io
from pptx import Presentation
from pptx.util import Inches
from docx.shared import Inches as DI
from PIL import Image
from openpyxl.chart import BarChart, Reference
shutil.rmtree('od',ignore_errors=True); shutil.rmtree('home',ignore_errors=True); os.makedirs('od/Reports/2026'); os.makedirs('home')
img=io.BytesIO(); Image.new('RGB',(400,200),(40,110,220)).save(img,'PNG')
d=docx.Document(); d.add_heading('Quarterly Report',1); d.add_paragraph('Original paragraph ALPHA.')
t=d.add_table(rows=2,cols=2); t.cell(0,0).text='A'; t.cell(1,1).text='D'; d.add_paragraph('Closing line.'); d.save('od/Report.docx')
wb=openpyxl.Workbook(); ws=wb.active; ws.title='Data'; ws['A1']=10; ws['A2']=32; ws['A3']='=SUM(A1:A2)'; ws['B1']='hello'; wb.create_sheet('Second')['A1']='s2'; wb.save('od/Budget.xlsx')
p=Presentation(); s=p.slides.add_slide(p.slide_layouts[0]); s.shapes.title.text='Deck Title BETA'; s.placeholders[1].text='Subtitle'
p.slides.add_slide(p.slide_layouts[1]).shapes.title.text='Second'; p.save('od/Reports/Deck.pptx')
# larger, realistic files
d=docx.Document(); d.add_heading('Big Document',0)
for i in range(60):
    d.add_heading(f'Section {i}',2); d.add_paragraph('Lorem ipsum dolor sit amet, consectetur adipiscing elit. '*12)
    if i%10==0: img.seek(0); d.add_picture(img,width=DI(3))
    if i%15==0:
        tb=d.add_table(rows=6,cols=4); tb.style='Table Grid'
        for r in range(6):
            for c in range(4): tb.cell(r,c).text=f'r{r}c{c}'
d.save('od/Reports/2026/Big.docx')
wb=openpyxl.Workbook(); ws=wb.active; ws.append(['Item','Qty','Price','Total'])
for i in range(2,5002): ws.append([f'Item {i}',i%17,round(i*1.37,2),f'=B{i}*C{i}'])
ws['F1']='=SUM(D2:D5001)'; ch=BarChart(); ch.add_data(Reference(ws,min_col=2,min_row=1,max_row=30),titles_from_data=True); ws.add_chart(ch,'H3'); wb.save('od/Reports/2026/Big.xlsx')
p=Presentation()
for i in range(25):
    s=p.slides.add_slide(p.slide_layouts[5]); s.shapes.title.text=f'Slide {i}'; img.seek(0); s.shapes.add_picture(img,Inches(1),Inches(2),width=Inches(4))
p.save('od/Reports/2026/Big.pptx')
open('od/Sheet.csv','w').write('name,qty\napple,3\npear,5\n')
open('od/Notes.txt','w').write('plain text note OMEGA\n')
open('od/Broken.docx','wb').write(open('od/Report.docx','rb').read()[:9000])
import zipfile
z=zipfile.ZipFile('od/BadXml.docx','w'); src=zipfile.ZipFile('od/Report.docx')
for n in src.namelist(): z.writestr(n, b'<w:document><broken' if n=='word/document.xml' else src.read(n))
z.close()
open('od/Locked.xlsx','wb').write(bytes.fromhex('d0cf11e0a1b11ae1')+b'\0'*600)
open('od/Empty.xlsx','wb').write(b'')
open('od/ignore.pdf','wb').write(b'%PDF'); open('od/~$Report.docx','wb').write(b'lock')

# ---- PDF Bench fixtures
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas as _cv
import random, pikepdf
def _pdf(path, n, tag):
    c=_cv.Canvas(path,pagesize=A4)
    for i in range(1,n+1):
        c.setFont('Helvetica-Bold',28); c.drawString(72,760,f'{tag} page {i}'); c.setFont('Helvetica',12)
        for k in range(12): c.drawString(72,720-k*18,f'{tag}{i} line {k} lorem ipsum dolor sit amet')
        c.showPage()
    c.save()
_pdf('od/Sample.pdf',5,'PAGE'); _pdf('od/Second.pdf',3,'SECOND')
random.seed(7)
big=Image.effect_noise((2400,1800),60).convert('RGB'); big.save('od/_big.jpg',quality=95)
png=Image.effect_noise((1200,900),40).convert('RGB'); png.save('od/_big.png')
c=_cv.Canvas('od/Photo.pdf',pagesize=A4)
for i in range(3):
    c.setFont('Helvetica-Bold',22); c.drawString(72,790,f'PHOTO page {i+1}')
    c.drawImage('od/_big.jpg',72,430,width=450,height=337); c.drawImage('od/_big.png',72,90,width=420,height=315); c.showPage()
c.save()
p=pikepdf.open('od/Sample.pdf'); p.pages[0].Rotate=90; p.save('od/Rot90.pdf')
p=pikepdf.open('od/Sample.pdf'); p.save('od/LockedPdf.pdf',encryption=pikepdf.Encryption(owner='o',user='secret',R=4))
open('od/BadPdf.pdf','wb').write(b'%PDF-1.4\n'+b'garbage not a pdf'*50)
Image.new('RGB',(800,600),(124,92,252)).save('od/pic1.png'); Image.new('RGB',(600,900),(240,86,140)).save('od/pic2.jpg',quality=90)
os.remove('od/_big.jpg'); os.remove('od/_big.png'); os.makedirs('out',exist_ok=True)
shutil.rmtree('out',ignore_errors=True); os.makedirs('out')

# ---- metadata fixtures
import glob as _g
for f in _g.glob('od/Meta*'): os.remove(f)
d=docx.Document(); d.add_paragraph('BODY_TEXT_KEEP'); cp=d.core_properties; cp.author='Secret Author'; cp.title='Secret Title'; cp.last_modified_by='Secret Editor'; cp.comments='Secret Comment'; cp.keywords='secret, hidden'; d.save('od/Meta.docx')
wb=openpyxl.Workbook(); ws=wb.active; ws['A1']=5; ws['A2']=6; ws['A3']='=A1*A2'; wb.properties.creator='Secret Author'; wb.properties.title='Secret Title'; wb.properties.lastModifiedBy='Secret Editor'; wb.save('od/Meta.xlsx')
p=Presentation(); s=p.slides.add_slide(p.slide_layouts[0]); s.shapes.title.text='SLIDE_KEEP'; p.core_properties.author='Secret Author'; p.core_properties.title='Secret Title'; p.save('od/Meta.pptx')
c=_cv.Canvas('od/_m.pdf',pagesize=A4); c.setAuthor('Secret Author'); c.setTitle('Secret Title'); c.setSubject('Secret Subject'); c.setKeywords('secret hidden')
for i in (1,2): c.drawString(72,760,f'PDF_KEEP page {i}'); c.showPage()
c.save()
with pikepdf.open('od/_m.pdf') as pd:
    with pd.open_metadata() as m: m['dc:creator']=['Secret Author']; m['dc:title']='Secret Title'
    pd.save('od/Meta.pdf')
os.remove('od/_m.pdf')
im=Image.linear_gradient('L').resize((640,480)).convert('RGB'); ex=Image.Exif(); ex[271]='SecretCam Make'; ex[272]='SecretCam X1'; ex[315]='Secret Author'; ex[274]=6; ex[306]='2026:01:02 03:04:05'
gps=ex.get_ifd(0x8825); gps[1]='N'; gps[2]=(12.0,34.0,56.0); gps[3]='E'; gps[4]=(77.0,35.0,10.0)
im.save('od/Meta.jpg',quality=90,exif=ex,comment=b'Secret comment')
from PIL.PngImagePlugin import PngInfo
pi=PngInfo(); pi.add_text('Author','Secret Author'); pi.add_text('Software','SecretCam'); pi.add_itxt('Description','Secret Comment'); Image.linear_gradient('L').resize((320,240)).convert('RGBA').save('od/Meta.png',pnginfo=pi)
