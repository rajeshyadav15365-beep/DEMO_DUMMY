from pathlib import Path
import json,math,csv,zipfile
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A3,landscape
from reportlab.lib.colors import black,white,HexColor
import fitz
O=Path('/workspace/project-review/outputs/BRRP_Design_Development_P4');OLD=Path('/workspace/project-review/outputs/bridge_engineering');data=json.loads((O/'P4_calculation_data.json').read_text());B=data['bridges'];W,H=landscape(A3)
P=O/'BRRP_P4_Source_Based_Development_Drawings.pdf';c=canvas.Canvas(str(P),pagesize=(W,H));c.setTitle('BRRP P4 Preliminary Design Development - 9 development sheets and source appendix')
navy=HexColor('#17365D');red=HexColor('#AC2020')
def tx(x,y,t,s=10,col=black):c.setFont('Helvetica',s);c.setFillColor(col);c.drawString(x,y,str(t))
def box(x,y,w,h):c.setStrokeColor(black);c.setLineWidth(.6);c.rect(x,y,w,h)
def frame(id,title,source=''):
 box(22,22,W-44,H-44);tx(38,H-43,'SHIVANSH / ONGC BRRP | PRELIMINARY DESIGN DEVELOPMENT',16,navy);tx(38,H-69,title,15)
 tx(38,H-91,'P4 - FOR ENGINEERING REVIEW | CONSTRUCTION RELEASE DEFERRED | NOT FOR PROCUREMENT OR EXECUTION',10,red)
 box(35,34,W-70,58);tx(45,74,id+' | Rev P4 | 25 Oct 2026 | Source-based development sheet; new Shivansh ID',10)
 tx(45,57,source,9);tx(45,41,'No member resistance, connection, bearing or lift approval is implied. Source results are not a new independently solved analysis.',9)
def line(x1,y1,x2,y2):c.setStrokeColor(black);c.setLineWidth(1);c.line(x1,y1,x2,y2)
def arrow(x1,y1,x2,y2):
 line(x1,y1,x2,y2);a=math.atan2(y2-y1,x2-x1)
 for delta in [-.5,.5]:line(x2,y2,x2-8*math.cos(a+delta),y2-8*math.sin(a+delta))
def table(x,y,widths,head,rows,h=29,fs=9):
 total=sum(widths);c.setFillColor(navy);c.rect(x,y-h,total,h,fill=1,stroke=0)
 xx=x
 for wi,t in zip(widths,head):tx(xx+5,y-18,t,fs,white);xx+=wi
 yy=y-h
 for row in rows:
  xx=x
  for wi,t in zip(widths,row):
   box(xx,yy-h,wi,h);tx(xx+5,yy-h+10,t,fs);xx+=wi
  yy-=h
 return yy
# Cover / register
frame('SHV-BRRP-P4-INDEX','DRAWING REGISTER AND ENGINEERING STATUS','Basis: DC3.4 PartI Rev14 / BRRP PartII / B774 scope, addendum and tender sources')
tx(45,680,'PACKAGE PURPOSE',12,navy)
for i,t in enumerate(['Develop existing bridge concepts into a traceable preliminary engineering package.','Use confirmed criteria and source drawings without inventing unverified site or fabrication dimensions.','Three arrangement sheets + three bearing/interface sheets + three detailing schedules.','Six original tender landing sheets are appended unchanged after the development sheets.','Open engineering inputs and calculations are recorded in the companion report and workbook.']):tx(45,655-i*23,t,11)
register=[]
for k in ['GA','BI','DD']:
 for b in B:register.append([f'SHV-BRRP-{b["code"]}-{k}-P4',b['name'],{'GA':'Study-source arrangement','BI':'Bearing / landing interface','DD':'Detail development schedule'}[k],'PRELIMINARY'])
table(45,510,[270,180,390,230],['Proposed ID','Bridge','Sheet scope','Status'],register,28,10)
tx(45,205,'COMPUTATION BOUNDARY',12,navy)
for i,t in enumerate(['Computed: loading arithmetic, source deflection ratios, thermal sensitivity, section geometry and ideal wrap quantities.','Not computed: a new global structural model, code member resistances, connection/bearing capacities or marine-operation design.','The detailed drawing register identifies the calculations and measurements required to close each detail.']):tx(45,180-i*22,t,10)
c.showPage()
# Arrangements, using original report images and explicit source geometry
for b in B:
 frame(f'SHV-BRRP-{b["code"]}-GA-P4',b['name']+' - PRELIMINARY ARRANGEMENT',f'Source B774-000-83-43-{b["report"]}; combined SOW/report PDF pp.{b["pages"]}')
 tx(45,690,'SOURCE-STUDY THREE-DIMENSIONAL MODEL',12,navy)
 c.drawImage(str(OLD/f'{b["code"]}_source_model.png'),45,430,width=730,height=240,preserveAspectRatio=True,anchor='c',mask='auto')
 tx(825,690,'SOURCE CROSS-SECTION',12,navy)
 c.drawImage(str(OLD/f'{b["code"]}_source_section.png'),835,450,width=250,height=220,preserveAspectRatio=True,anchor='c',mask='auto')
 tx(800,425,'Structural depth / overall width: not verified',10)
 tx(800,407,f'Loading-width reference: {b["b"]} m',10)
 tx(45,385,'LONGITUDINAL STUDY REFERENCE - NOT TO SCALE',12,navy)
 x1,x2=120,W-130;y=330;line(x1,y,x2,y)
 for x in [x1,x2]:box(x-14,y-26,28,20)
 tx(x1-45,y-46,b['left']+' - sliding',11);tx(x2-150,y-46,b['right']+' - hinged/fixed intent',11)
 tx(330,y+20,f'Study model length {b["L"]} m; approximate reported gap {b["gap"]} m',12)
 tx(45,255,'Final bearing centres, levels, skew, member coordinates, section mapping and camber profile require engineering verification.',10)
 notes=[f'25-year design life. E=210000 MPa; density=7850 kg/m3; alpha=12x10^-6 /C. (Criteria pp66,119)',f'Global walkway loading2.4525kN/m2; local grating loading4.905kN/m2. Actual loading areas must be reconciled.',f'Study source vertical deflection {b["dv"]}mm / criteria limit {b["vlimit"]:.0f}mm; lateral {b["dh"]}mm / {b["hlimit"]:.0f}mm.',f'Wind gust duration: {15 if b["L"]>50 else 3}sec per Annexure5. Speed and elevation correction require documented calculation.','Camber to compensate verified dead-load deflection; source1:200 statement is not adopted.','BHS new bridge and services must be accepted before existing bridge demolition.' if b['code']=='103' else 'Include approved landing modifications and applicable ICG/ICP jacket strengthening.']
 for j,t in enumerate(notes):tx(45,226-19*j,t,10)
 c.showPage()
# Bearing/interface functional sheets (no component dimensions selected)
for b in B:
 frame(f'SHV-BRRP-{b["code"]}-BI-P4',b['name']+' - BEARING AND LANDING INTERFACE DEVELOPMENT','DC3.4 PartI3.4.16.4.17 pp42-43; source arrangement and platform drawings; dimensions pending')
 tx(45,690,'FUNCTIONAL SUPPORT DIAGRAM - COMPONENTS NOT SIZED OR SELECTED',12,navy)
 # two physical stack concepts explicitly qualitative
 for x,end,kind in [(120,b['left'],'SLIDING'),(700,b['right'],'HINGED / PINNED INTENT')]:
  tx(x,650,end+' PLATFORM | '+kind,12)
  box(x,550,350,45);tx(x+20,568,'Bridge end interface - geometry TBC',10)
  box(x+25,524,300,18);tx(x+40,529,'Self-lubricating PTFE bearing interface',9)
  box(x+5,470,340,45);tx(x+20,489,'Landing / stool / existing load path',10)
  if kind=='SLIDING':arrow(x+65,620,x+270,620);arrow(x+270,609,x+65,609);tx(x+38,440,'Longitudinal movement + lateral/vertical guides',10)
  else:tx(x+22,615,'Axial thrust and rotation provision to be designed',10);tx(x+22,440,'Does not imply moment-rigid frame-end restraint',10)
 tx(45,400,'MOVEMENT AND FORCE DESIGN BASIS',12,navy)
 formulae=['Required directional translation =1.5 x (absolute platform-A translation + absolute platform-B translation).','Required hinged-connection axial action =1.5 x expected axial thrust, with appropriate signed action envelope.','Thermal sensitivity deltaL =alpha x L x deltaT; compatible combined movement must avoid double counting.',f'Thermal sensitivity for this study length = {b["thermal_per_C"]:.3f} mm per degree C; actual design temperature range remains to be set.','Criteria location-tolerance wording must be interpreted for existing supports; it is not an adopted bearing-travel value.']
 for j,t in enumerate(formulae):tx(45,376-23*j,t,10)
 tx(45,238,'DETAILS TO CLOSE BEFORE COMPONENT SIZING',12,navy)
 req=['Coordinates and datum; rotations; signed concurrent reactions; uplift and friction; travel and guide gaps.','Bearing pressure/contact; local plates and stiffeners; anchorage/connection resistance; corrosion and inspection access.','Surveyed existing steel / UT condition; landing/deck/jacket capacity; installation tolerances and replacement method.',('SLQ conflict remains a query: report/methodology fixed intent versus contradictory SOW wording.' if b['code']=='103' else 'Use correct platform drawing and verify member/interface labels before detailed modelling.')]
 for j,t in enumerate(req):tx(45,213-23*j,t,10)
 c.showPage()
# detailing/traceability sheets
for b in B:
 frame(f'SHV-BRRP-{b["code"]}-DD-P4',b['name']+' - DETAILING AND CALCULATION INTERFACE SCHEDULE','Source-based development register; proposed mark system only; no fabrication dimensions issued')
 tx(45,692,'PROPOSED DETAIL MARKS - ASSIGN TO FINAL MODEL AND CALCULATIONS BEFORE DRAWING RELEASE',11,navy)
 detailrows=[['CH','Chords','Section/group mapping; axial/bending/buckling/fatigue','NOT SELECTED'],['BR','Braces / crossframes','Node geometry; effective lengths; joint capacity','NOT SELECTED'],['WB','Walkway beams','Actual spans; local loading; grating support','NOT SELECTED'],['EC','End connection','Concurrent forces; rotations; local resistance','NOT SIZED'],['BE','Bearing / guides','Movement; pressure; uplift; friction; anchorage','NOT SIZED'],['LS','Landing / stool','Survey/UT; support reactions; existing capacity','VERIFY SOURCE'],['SS','Services supports','Routing; operational/hydrotest loads; future allowance','COORDINATE'],['AC','Access / fittings','Clear width; handrails; grating; standards/IRT','SOURCE DETAILS'],['TW','Temporary works','Stage loads; support arrangement; stability','NOT DESIGNED'],['LP','Lifting interface','Verified mass/COG; rigging; skew/dynamics','NOT DESIGNED'],['SF','Transport restraints','Vessel motions; support reactions; deck checks','NOT DESIGNED']]
 table(45,665,[65,210,590,205],['Mark','Detail family','Calculation / evidence required','Current status'],detailrows,29,10)
 tx(45,290,'INSPECTION AND DOCUMENT REQUIREMENTS',12,navy)
 for j,t in enumerate(['Material traceability and certificates; approved welding procedures / qualifications; dimensional control.','NDT extent, hold/witness points and acceptance per applicable specification and approved inspection plan.','Coating exposure/system and inspection; source MTO IRT pp15-34; no generic welding or NDT percentage assumed.','Every final dimension/detail must identify controlling source, calculation reference, revision and engineering reviewer.']):tx(45,264-22*j,t,10)
 tx(45,150,'BRIDGE-SPECIFIC INTERFACE',11,navy)
 tx(45,130,('BHS: piping/cable supports, roofing, service transfer, lifeboat/muster interfaces and demolition package.' if b['code']=='103' else 'IC complex: landing drawing references, addendum jacket strengthening, access and adjacent live-service clearance.'),10)
 c.showPage()
c.save()
# Append original landing drawings unchanged to preserve source evidence.
F=fitz.open(P);m=fitz.open(next(Path('/workspace/attachments').rglob('MTO Drawings IRT.pdf')))
for i in range(1,7):F.insert_pdf(m,from_page=i,to_page=i)
T=O/'BRRP_P4_Drawings_With_Tender_Landing_Appendix.pdf';F.save(T)
# SVG pages preserve vector annotations with embedded source images.
S=fitz.open(P);sv=O/'SVG_Sheets';sv.mkdir(exist_ok=True)
for i,p in enumerate(S):
 (sv/f'P4_Sheet_{i+1:02}.svg').write_text(p.get_svg_image())
 if i in [1,2,3,4]:p.get_pixmap(matrix=fitz.Matrix(1.1,1.1),alpha=False).save(O/f'P4_Preview_Sheet_{i+1:02}.png')
assert len(S)==10 and len(F)==16
for p in S:assert 'CONSTRUCTION RELEASE DEFERRED' in p.get_text()
# Human-readable manifest and single archive.
manifest=['BRRP PRELIMINARY DESIGN DEVELOPMENT PACKAGE P4','Date25 October2026','Construction release deferred per user instruction.','', 'Contents:','- Design development report DOCX','-12-sheet calculation workbook XLSX','-10-page PDF: cover/register +9 source-based development sheets','-16-page PDF: development sheets plus6 original tender landing drawings','-10 SVG sheet exports','-Static calculation snapshot CSV and calculation-data JSON','-Preview images','', 'Source results remain attributed to original bridge studies. New calculations are explicitly identified. No new3D structural solve or member/connection/bearing/lifting adequacy is claimed.','Retain original source PDFs and use source page references in report.']
(O/'README_P4.txt').write_text('\n'.join(manifest))
Z=O.parent/'BRRP_Design_Development_P4.zip'
with zipfile.ZipFile(Z,'w',zipfile.ZIP_DEFLATED) as z:
 for p in sorted(O.rglob('*')):
  if p.is_file():z.write(p,arcname='BRRP_Design_Development_P4/'+str(p.relative_to(O)))
with zipfile.ZipFile(Z) as z:assert z.testzip() is None
print('Development PDF pages',len(S),'with source appendix',len(F),'SVG exports',len(list(sv.glob('*.svg'))))
print('ZIP',Z,'bytes',Z.stat().st_size)
