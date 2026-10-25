from pathlib import Path
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A3,landscape
from reportlab.lib.colors import HexColor,black,white
import fitz,csv,json
O=Path('/workspace/project-review/outputs/bridge_engineering');O.mkdir(exist_ok=True)
W,H=landscape(A3)
bridges=[dict(code='101',name='ICW - ICG LOWER DECK',left='ICW',right='ICG',ls='S: SLIDING (SOW)',rs='F: FIXED (SOW)',clause='6.1.1(c)-(d)',page='42',notes=['ICW landing modifications / extension: -0002 / -0003.','ICG landing modifications / extension: -0004 / -0005.','Include applicable ICG jacket strengthening from structural addendum.','SOW 6.1.3(e) calls the ICG-side work ICW strengthening: query location.']),dict(code='102',name='ICP - ICG LOWER DECK',left='ICP',right='ICG',ls='S: SLIDING (SOW)',rs='F: FIXED (SOW)',clause='6.2.1(c)-(d)',page='44',notes=['ICP landing modifications / extension: -0002 / -0003.','ICG landing modifications / extension: -0004 / -0005.','Include applicable ICP and ICG jacket strengthening from addendum.','Existing platform capacity, corrosion and reactions require verification.']),dict(code='103',name='BHS - SLQ LOWER DECK',left='BHS',right='SLQ',ls='S: SLIDING (SOW)',rs='TYPE UNRESOLVED - QUERY',clause='6.3.1(c)-(d); 6.3.3(f)',page='45-48',notes=['SOW 6.3.1(c) says SLQ fixed; 6.3.3(f) says sliding: resolve before analysis.','BHS / SLQ landing modifications and extensions: -0002 through -0005.','Include piping/cables, cladding/roofing and applicable lifeboat interfaces.','New bridge and services accepted before existing bridge demolition.'])]
C=canvas.Canvas(str(O/'Replacement_Bridges_Schematic_Basis_RevP0.pdf'),pagesize=(W,H))
C.setTitle('Three replacement bridges - source-based schematic basis - not detailed engineering')
def txt(x,y,s,size=10,color=black):
 C.setFillColor(color);C.setFont('Helvetica',size);C.drawString(x,y,s)
def box(x,y,w,h,label=None):
 C.setStrokeColor(black);C.rect(x,y,w,h)
 if label:txt(x+7,y+h-16,label,11)
def dim(x1,x2,y,label):
 C.line(x1,y,x2,y)
 for x in [x1,x2]:C.line(x-4,y-4,x+4,y+4)
 C.setFillColor(white);C.rect((x1+x2)/2-150,y-7,300,16,fill=1,stroke=0);txt((x1+x2)/2-140,y-3,label,10)
for b in bridges:
 C.setStrokeColor(black);C.setLineWidth(.7);C.rect(24,24,W-48,H-48)
 txt(40,H-48,'SHIVANSH / ONGC BRRP | REPLACEMENT BRIDGE DRAWING BASIS',15)
 txt(40,H-71,b['name'],18)
 txt(40,H-92,'SOURCE-BASED SCHEMATIC ONLY | NOT TO SCALE | NOT FOR DESIGN, FABRICATION OR INSTALLATION',11,HexColor('#B22222'))
 txt(40,H-112,'No structural sizing, connection design or site dimensions have been inferred.',10)
 # plan
 y=H-310
 txt(50,y+135,'A. FUNCTIONAL PLAN - GEOMETRY AND MEMBER ARRANGEMENT NOT ESTABLISHED',12)
 box(65,y-10,125,95);txt(78,y+40,b['left']+' PLATFORM',11)
 box(W-195,y-10,125,95);txt(W-184,y+40,b['right']+' PLATFORM',11)
 x1,x2=215,W-220
 C.setDash(5,4);C.rect(x1,y+12,x2-x1,50);C.setDash()
 txt(x1+65,y+40,'BRIDGE ENVELOPE: WIDTH / FRAMING / CLEARANCES TBC',11)
 txt(x1+65,y+23,'Service routes and walkway arrangement subject to approved design',10)
 C.line(190,y+37,x1,y+37);C.line(x2,y+37,W-195,y+37)
 dim(x1,x2,y+99,'L = VERIFIED BEARING-CENTRE SPAN REQUIRED')
 txt(x1,y-12,'B = clear width TBC; support coordinates, skew and orientation TBC',10)
 # elevation
 y2=y-160
 txt(50,y2+95,'B. SUPPORT-INTERFACE SCHEMATIC - NOT AN ANALYSIS OR BEARING DETAIL',12)
 C.setLineWidth(2);C.line(x1,y2+45,x2,y2+45);C.setLineWidth(.7)
 box(x1-15,y2+10,30,28);box(x2-15,y2+10,30,28)
 txt(x1-125,y2-9,b['left']+': '+b['ls'],10)
 txt(x2-175,y2-9,b['right']+': '+b['rs'],10,HexColor('#B22222') if b['code']=='103' else black)
 txt(x1+100,y2+60,'Structural depth / sections / connections: NOT DESIGNED',11)
 txt(x1+100,y2+27,'Deck elevation, support levels and movement envelope: TBC',10)
 txt(50,y2-38,'F/S labels identify SOW support intent only; DOFs, guides, uplift restraints and travel are not assigned.',10)
 # notes lower
 ny=y2-70
 txt(50,ny,'C. BRIDGE-SPECIFIC SCOPE / ENGINEERING QUERIES',12)
 for j,n in enumerate(b['notes']):txt(55,ny-20*(j+1),str(j+1)+'. '+n,10)
 txt(50,ny-110,'D. INPUTS NEEDED: controlling tender drawings; dimensional survey; design criteria; loads; existing structure data;',10)
 txt(50,ny-126,'material specifications; corrosion assessment; service loads; checked model; approved marine and temporary-work basis.',10)
 # title block
 box(40,38,W-80,75)
 txt(50,94,'Proposed document ID: SHV-BRRP-'+b['code']+'-SK-001 | Rev P0 | 25 Oct 2026',11)
 txt(50,77,'SOW reference: B774-000-83-43-SOW-0001 Rev C, '+b['clause']+'; uploaded PDF page(s) '+b['page'],10)
 txt(50,60,'Tender register identifies B774-'+b['code']+'-83-43-0001; narrative cites -83-41-0001. Drawing identity requires clarification.',10)
 txt(50,43,'Prepared as a source-review aid. No professional approval, construction release or calculated adequacy is implied.',9)
 C.showPage()
C.save()
d=fitz.open(O/'Replacement_Bridges_Schematic_Basis_RevP0.pdf')
for i,p in enumerate(d):p.get_pixmap(matrix=fitz.Matrix(1.25,1.25),alpha=False).save(O/f'Bridge_{bridges[i]["code"]}_Schematic.png')
# Engineering deliverable register: proposed IDs distinct from tender IDs.
items=[('001','General arrangement: plan, elevations, sections and coordinates','Survey geometry + approved design basis'),('002','Primary framing and member schedule','Checked global analysis and material specification'),('003','Crossframes, lateral bracing and secondary steel','Global/local stability, serviceability and connection design'),('004','Fixed/guided/sliding bearing and movement details','Confirmed support philosophy, movements and reactions'),('005','End connections, welds and bolts','Checked member forces, local checks and approved fabrication requirements'),('006','Landing extension and existing steel interfaces','Platform as-builts, UT survey, local/global capacity'),('007','Jacket/platform strengthening','Approved addendum interpretation and existing structure checks'),('008','Grating, handrails, stairs and access details','Applicable standards, clearances and access design'),('009','Piping, cable, lighting and service supports','Multidiscipline routing, loads and operating conditions'),('010','Module splits, lifting points and temporary works','Verified weights/COG, lift/transport and temporary-condition calculations'),('011','Loadout, transport supports and sea-fastening','Vessel/transport data and engineered reactions'),('012','Material take-off, weld map and inspection interfaces','Approved detailed design and traceability rules')]
with open(O/'Replacement_Bridge_Engineering_Drawing_Register.csv','w',newline='') as f:
 w=csv.writer(f);w.writerow(['Proposed drawing ID','Bridge','Deliverable','Required basis','Status'])
 for b in bridges:
  for n,title,basis in items:w.writerow([f'SHV-BRRP-{b["code"]}-ST-{n}',b['name'],title,basis,'NOT DESIGNED / INPUTS REQUIRED'])
 w.writerow(['SHV-BRRP-103-ST-013','BHS-SLQ','Existing bridge demolition and temporary stability','Accepted new bridge/services; survey and engineered removal plan','NOT DESIGNED / INPUTS REQUIRED'])
queries=[('TQ-01','All','Drawing number conflict','SOW narrative uses 83-41; annexure uses 83-43','Provide correct controlling drawing IDs and revisions'),('TQ-02','BHS-SLQ','SLQ bearing conflict','6.3.1(c): fixed; 6.3.3(f): sliding','Confirm intended support DOFs and movement requirements'),('TQ-03','BHS-SLQ','Landing strengthening cross references','6.3.2(g) and 6.3.3(e) cite B774-102-83-43-0002','Confirm BHS/SLQ drawings; do not apply ICP drawing by inference'),('TQ-04','BHS-SLQ','Demolition reference conflict','6.3.4 cites B774-005 series; annexure identifies B774-103-83-43-0006','Provide correct demolition drawing and verify existing supports'),('TQ-05','ICW-ICG','Landing location wording conflict','6.1.3(e) describes ICW strengthening within ICG-side section','Confirm platform location and scope'),('TQ-06','All','Missing dimensional basis','No replacement GAs located in inspected uploaded PDFs','Provide controlling tender drawing pack and survey'),('TQ-07','All','Design and load basis incomplete','Referenced design criteria and final service loads need verification','Provide approved criteria, existing models and load registers')]
with open(O/'Replacement_Bridge_Technical_Query_Register.csv','w',newline='') as f:
 w=csv.writer(f);w.writerow(['ID','Bridge','Issue','Evidence','Required resolution']);w.writerows(queries)
print('Created 3 schematic sheets; 37 proposed detailed-design deliverables; 7 technical queries. These are NOT detailed engineering drawings.')
print('PDF pages',len(d),'dimensions',d[0].rect)
for p in O.iterdir():print(p.name,p.stat().st_size)
