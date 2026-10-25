from pathlib import Path
import fitz,csv
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A3,landscape
from reportlab.lib.colors import HexColor,black,white
O=Path('/workspace/project-review/outputs/bridge_engineering');SRC=next(Path('/workspace/attachments').rglob('SOW & Bridge Report Combined.pdf'));D=fitz.open(SRC)
records=[dict(code='101',name='ICW - ICG LOWER DECK',left='ICW',right='ICG',gap='54.5 m',model='56 m',walk='1.2 m',doc='B774-000-83-43-BPR-0004 Rev 0',pages='133-135',iso=(133,(72,507,553,655)),section=(134,(240,390,378,548)),wind='61.11 / 37.7 m/s; 15-second gust (study values)',land='ICW: grid 4 northside; ICG: grid 3 southside',report='Report pp. 5-7 of 27',loc='SOW 6.1; report PDF p.133'),dict(code='102',name='ICP - ICG LOWER DECK',left='ICP',right='ICG',gap='80.5 m',model='84 m',walk='1.2 m',doc='B774-000-83-43-BPR-0003 Rev 0',pages='196-199',iso=(197,(106,211,557,379)),section=(198,(241,142,377,298)),wind='57.19 / 35.68 m/s; 15-second gust (study values)',land='ICP: grid 3 northside; ICG: scramble-net area',report='Report pp. 5-8 of 20',loc='SOW 6.2; report PDF p.196'),dict(code='103',name='BHS - SLQ LOWER DECK',left='BHS',right='SLQ',gap='40 m',model='42 m',walk='2.324 m',doc='B774-000-83-43-BPR-0005 Rev A',pages='97-100',iso=(98,(104,130,515,350)),section=(99,(203,467,415,686)),wind='65.45 / 39.99 m/s; 3-second gust (study values)',land='BHS: near A1 southside; SLQ: near B2 southside',report='Report pp. 5-8 of 36',loc='SOW 6.3; report PDF pp.97-100')]
W,H=landscape(A3);C=canvas.Canvas(str(O/'Replacement_Bridges_Source_Informed_GA_RevP1.pdf'),pagesize=(W,H));C.setTitle('Replacement bridges - study-based preliminary arrangement P1')
def text(x,y,t,size=10,col=black):C.setFillColor(col);C.setFont('Helvetica',size);C.drawString(x,y,t)
def rect(x,y,w,h):C.setStrokeColor(black);C.setLineWidth(.6);C.rect(x,y,w,h)
def image_source(which,code,key):
 n,clip=which;p=O/f'{code}_source_{key}.png';D[n-1].get_pixmap(matrix=fitz.Matrix(2,2),clip=fitz.Rect(clip),alpha=False).save(p);return p
for r in records:
 rect(22,22,W-44,H-44)
 text(38,H-45,'SHIVANSH / ONGC BRRP - SOURCE-INFORMED PRELIMINARY ARRANGEMENT',16)
 text(38,H-71,r['name']+' | SHV-BRRP-'+r['code']+'-GA-P01 | Rev P1',15)
 text(38,H-93,'STUDY REFERENCE ONLY - NOT FOR FABRICATION, PROCUREMENT, LIFTING OR INSTALLATION',11,HexColor('#AF2020'))
 text(38,H-112,'Source model and cross-section reproduced from uploaded report; no new member sizing or structural verification performed.',10)
 # source model and section
 text(40,695,'A. SOURCE REPLACEMENT-STUDY MODEL',12)
 iso=image_source(r['iso'],r['code'],'model');C.drawImage(str(iso),43,432,width=735,height=246,preserveAspectRatio=True,anchor='c',mask='auto')
 text(43,420,'Source: '+r['doc']+' | uploaded combined PDF p.'+str(r['iso'][0]),9)
 text(815,695,'B. SOURCE CROSS-SECTION',12)
 cs=image_source(r['section'],r['code'],'section');C.drawImage(str(cs),835,442,width=270,height=235,preserveAspectRatio=True,anchor='c',mask='auto')
 text(810,420,'Overall width and structural depth: verify',10)
 text(810,402,'Walkway loading width: '+r['walk']+' (report basis)',10)
 # source-informed reference dimensional schematic
 text(40,376,'C. LONGITUDINAL REFERENCE / SUPPORT INTERFACES - SCHEMATIC, NOT TO SCALE',12)
 x1,x2=130,W-150;y=320
 C.setLineWidth(2);C.line(x1,y,x2,y);C.setLineWidth(.6)
 for x in [x1,x2]:rect(x-15,y-28,30,22)
 text(x1-55,y-48,r['left']+' - sliding',11)
 right=r['right']+' - fixed per study'+(' *' if r['code']=='103' else '')
 text(x2-135,y-48,right,11)
 C.line(x1,y+27,x2,y+27)
 for x in [x1,x2]:C.line(x,y+20,x,y+34)
 label='SACS STUDY MODEL LENGTH '+r['model']+' - NOT A VERIFIED FABRICATION SPAN'
 C.setFillColor(white);C.rect(300,y+19,600,17,fill=1,stroke=0);text(310,y+23,label,10)
 text(45,245,'Approximate 3D-model gap: '+r['gap']+' | final bearing centres, skew, elevations, offsets and movement envelope require survey.',10)
 text(45,227,'Landing reference: '+r['land']+'; confirm against current survey and controlling drawings.',10)
 text(45,209,'The reproduced model indicates braced space-truss framing. No panel spacing, section size or joint detail is adopted here.',10)
 # notes
 text(40,180,'D. DESIGN DEVELOPMENT NOTES',12)
 notes=[
 '1. Retain report geometry only as a study reference; resolve actual support positions and dimensions before detailed design.',
 '2. Report steel yield basis is 50 ksi; procurement grade, toughness, weldability and material requirements still require approval.',
 '3. Report walkway loading is 250 kg/m2; validate loading, combinations, movements, serviceability and fatigue criteria.',
 '4. Do not derive welds, bolt sizes, bearing travel, lifting points or plate thicknesses from model pictures or quantity totals.',
 '5. Include permanent, temporary, lifting, transport and installation checks plus existing-platform/jacket capacity.',
 ('6. SLQ fixed support is supported by the study; SOW 6.3.3(f) says sliding. Resolve this contradiction by technical query.' if r['code']=='103' else '6. Incorporate applicable ICG/ICP jacket strengthening and landing modifications from SOW and structural addendum.')]
 for j,t in enumerate(notes):text(45,161-16*j,t,9)
 text(40,44,'SOURCE: '+r['doc']+'; '+r['report']+'; combined PDF pp.'+r['pages']+' | 25 Oct 2026 | No AFC approval implied.',9)
 C.showPage()
C.save()
P=fitz.open(O/'Replacement_Bridges_Source_Informed_GA_RevP1.pdf')
for i,p in enumerate(P):p.get_pixmap(matrix=fitz.Matrix(1.2,1.2),alpha=False).save(O/f'Bridge_{records[i]["code"]}_GA_RevP1.png')
# Compact source basis register, distinguishing study model from measured gap.
with open(O/'Replacement_Bridge_Verified_Study_References.csv','w',newline='') as f:
 w=csv.writer(f);w.writerow(['Bridge','Approximate gap in report','SACS model length','Walkway loading width','Source report','Combined PDF pages','Status'])
 for r in records:w.writerow([r['name'],r['gap'],r['model'],r['walk'],r['doc'],r['pages'],'Source-verified preliminary study; not AFC'])
# Extract actual source pages, preserving source content and page order.
E=fitz.open()
for n in [97,98,99,100,133,134,135,196,197,198,199]:E.insert_pdf(D,from_page=n-1,to_page=n-1)
E.save(O/'Replacement_Bridge_Source_Study_Excerpts.pdf')
# Write technical handover record.
(O/'RevP1_Engineering_Review_Notes.txt').write_text('Correction to P0: uploaded bridge reports contain preliminary replacement-study lengths and model figures. P1 incorporates these source references; dimensioned B774 fabrication GAs remain unlocated.\nICW: gap54.5m/model56m; ICP: gap80.5m/model84m; BHS: gap40m/model42m. These are different quantities.\nWalkway widths are loading assumptions, not overall structural widths.\nBHS report corroborates fixed SLQ intent but does not remove the SOW conflict.\nWind gust durations differ between reports; do not adopt speeds without reconciling the design basis.\nICW report live-load table156.25kN vs SACS summary144.69kN requires reconciliation. ICP report model84m/walkway1.2m/live load288kN also requires area/load reconciliation; no correction assumed.\nDetailed member, connection, bearing, fatigue and temporary-condition checks remain required.\n')
# Checks
assert len(P)==3
for page in P:assert 'NOT FOR FABRICATION' in page.get_text()
print('Created three source-informed GA sheets P1, an 11-page source-excerpt PDF, study reference CSV and review notes.')
