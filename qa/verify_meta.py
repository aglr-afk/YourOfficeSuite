import sys, json, zipfile, warnings; warnings.simplefilter('ignore')
import docx, openpyxl, pypdf
from pptx import Presentation
from PIL import Image, ImageChops
mode=sys.argv[1]; sfx='clean' if mode=='clean' else 'details'; R={}
def same(a,b):
    x=Image.open(a).convert('RGBA'); y=Image.open(b).convert('RGBA'); return x.size==y.size and ImageChops.difference(x,y).getbbox() is None
def raw(f): return open(f,'rb').read()
def zraw(f):
    z=zipfile.ZipFile(f); assert z.testzip() is None; return b''.join(z.read(n) for n in z.namelist() if n.startswith('docProps/'))
D=docx.Document(f'od/Meta-{sfx}.docx'); c=D.core_properties
X=openpyxl.load_workbook(f'od/Meta-{sfx}.xlsx'); P=Presentation(f'od/Meta-{sfx}.pptx'); pc=P.core_properties
F=pypdf.PdfReader(f'od/Meta-{sfx}.pdf'); md=dict(F.metadata or {})
J=Image.open(f'od/Meta-{sfx}.jpg'); ex=J.getexif(); G=Image.open(f'od/Meta-{sfx}.png')
R['docx']=dict(author=c.author,title=c.title,last=c.last_modified_by,comments=c.comments,keywords=c.keywords,body='BODY_TEXT_KEEP' in '|'.join(p.text for p in D.paragraphs),secret=b'Secret' in zraw(f'od/Meta-{sfx}.docx'),company=b'<Company>QA Co</Company>' in zraw(f'od/Meta-{sfx}.docx'))
R['xlsx']=dict(author=X.properties.creator,title=X.properties.title,formula=X.active['A3'].value,secret=b'Secret' in zraw(f'od/Meta-{sfx}.xlsx'))
R['pptx']=dict(author=pc.author,title=pc.title,slide=P.slides[0].shapes.title.text,secret=b'Secret' in zraw(f'od/Meta-{sfx}.pptx'))
R['pdf']=dict(meta={k:str(v) for k,v in md.items()},xmp=F.xmp_metadata is not None,pages=len(F.pages),text='PDF_KEEP page 2' in F.pages[1].extract_text(),secret=b'Secret' in raw(f'od/Meta-{sfx}.pdf'))
R['jpg']=dict(exif={int(k):str(v) for k,v in ex.items() if k!=0x8825},gps=bool(ex.get_ifd(0x8825)),same=same('od/Meta.jpg',f'od/Meta-{sfx}.jpg'),secret=b'Secret' in raw(f'od/Meta-{sfx}.jpg'),xmp=b'QA Comment' in raw(f'od/Meta-{sfx}.jpg'),comment='comment' in J.info)
R['png']=dict(info={k:str(v) for k,v in G.info.items()},same=same('od/Meta.png',f'od/Meta-{sfx}.png'),secret=b'Secret' in raw(f'od/Meta-{sfx}.png'))
print(json.dumps(R))
