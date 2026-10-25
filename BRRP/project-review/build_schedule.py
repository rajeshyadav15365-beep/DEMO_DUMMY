from datetime import date,timedelta
from pathlib import Path
import calendar,csv
import openpyxl
from openpyxl.styles import Font,PatternFill,Alignment
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import ColorScaleRule
O=Path('/workspace/project-review/outputs'); START=date(2026,11,1); A=[]
def add(i,phase,pkg,name,dur,preds='',resource='',owner='',gate='',earliest=0,off=False):
 ps=preds.split(',') if preds else []; start=max([earliest]+[next(a['ef'] for a in A if a['id']==p) for p in ps]); end=start+dur
 A.append(dict(id=i,phase=phase,pkg=pkg,name=name,dur=dur,preds=ps,resource=resource,owner=owner,gate=gate,es=start,ef=end,off=off));return i
add('M000','Milestone','COMMON','Assumed NOA / commencement',0,owner='Contracts')
add('G010','Management','COMMON','Kick-off, obligations, guarantees and funding arrangements',7,'M000','Project team','Project Director')
add('G020','Management','COMMON','Controlled document / addendum and quantity reconciliation',14,'G010','Project team','Engineering','Resolve 1 MT and demolition allocation')
add('S010','Survey','COMMON','Survey permits, mobilisation and access planning',7,'G010','Survey team','Survey Manager')
add('S020','Survey','ALL','Bridge and platform verification survey campaign',14,'S010','Survey spread','Survey Manager','Accepted survey / utility register',off=True)
add('E010','Engineering','COMMON','Design basis and interface approval',14,'G010','Engineering','Engineering Manager')
add('E020','Engineering','ALL','Checked material requisitions / early material release',10,'S020,E010,G020','Engineering','Engineering Manager','Approved procurement MTO; no fabrication release')
add('P010','Procurement','ALL','Vendor approval, steel purchase and phased yard delivery',25,'E020','Procurement','Procurement Manager','Certificates, inspection and traceability')
add('E030','Engineering','REPLACEMENTS','Detailed design, AFC drawings and ITP approval',25,'S020,E010','Design team A','Engineering Manager','Approved for construction')
add('E040','Engineering','REFURBISHMENTS','Bridge and jacket strengthening design / AFC work packs',25,'S020,E010','Design team B','Engineering Manager','ICG / ICP / BHS addendum incorporated')
add('E050','Engineering','BHS-SLQ','Utility transfer, temporary stability and demolition design',35,'S020,E010','Design team A','Technical Authority','ONGC operational interface approval')
add('MR10','Marine','COMMON','Vessel market enquiry, technical audit and conditional reservation',35,'G010','Marine engineering','Marine Manager','Certified vessel data and availability')
add('MR20','Marine','REPLACEMENTS','Lift, loadout, transport and sea-fastening engineering',35,'E030,MR10','Marine engineering','Marine Manager','Verified lift weights / COG; approved procedures')
add('F010','Fabrication','ICW-ICG LD','Fabricate, NDT, dimensional inspection and coating',35,'P010,E030','Fabrication line 1','Fabrication Manager','QA release')
add('F020','Fabrication','BHS-SLQ','Fabricate, NDT, dimensional inspection and coating',55,'P010,E030','Fabrication line 2','Fabrication Manager','QA release')
add('F030','Fabrication','ICG-ICP LD','Fabricate, NDT, dimensional inspection and coating',40,'F010','Fabrication line 1','Fabrication Manager','QA release')
add('F040','Fabrication','ALL STRENGTHENING','Phased fabrication and release of strengthening kits',35,'P010,E040','Dedicated strengthening cell','Fabrication Manager','All kits released at finish; earlier release only by QA')
add('O010','Offshore readiness','ALL','Accommodation/support mobilisation, permits, SIMOPS and readiness',10,'F040,MR10','Support spread','Offshore Manager','Minimum 60 special-personnel accommodation per SCC',off=True)
add('J010','Offshore','ICG / BHS','Addendum jacket strengthening and acceptance',12,'O010','Strengthening crew A','Offshore Manager','Before ICW-ICG LD and BHS replacement',off=True)
add('J020','Offshore','ICP','Addendum jacket strengthening and acceptance',10,'O010','Strengthening crew B','Offshore Manager','Before ICG-ICP LD installation',off=True)
# Sequential workfronts for two independent refurbishment crews.
prev='J010'; crewA=[]
for n,(p,d) in enumerate([('ICW-ICG UD',16),('ICG-ICP UD',22),('SLQ-WIS',18),('NQO-NQD',18)],1):
 i=f'RA{n:02}';add(i,'Offshore',p,'Strengthening, inspection and bridge acceptance',d,prev,'Strengthening crew A','Offshore Manager','PTW, stability, NDT and acceptance',off=True);prev=i;crewA.append(i)
prev='J020';crewB=[]
for n,(p,d) in enumerate([('WIN-NC',12),('SCA-SC1',16),('SCA-SCF',24)],1):
 i=f'RB{n:02}';add(i,'Offshore',p,'Strengthening, inspection and bridge acceptance',d,prev,'Strengthening crew B','Offshore Manager','PTW, stability, NDT and acceptance',off=True);prev=i;crewB.append(i)
add('H010','Marine','COMMON','Heavy-lift spread mobilisation and operational readiness',7,'MR20,O010','Heavy-lift spread','Marine Manager','Survey/class, crane curves, permits and weather',off=True)
add('L010','Marine','ICW-ICG LD','Loadout, sea-fastening and transport',5,'F010,MR20','Cargo spread','Marine Manager','Sail-away release',off=True)
add('I010','Offshore','ICW-ICG LD','Install, connect, inspect and accept replacement bridge',16,'L010,H010,J010','Heavy-lift spread','Offshore Manager','Services / safe access accepted',off=True)
add('L020','Marine','BHS-SLQ','Loadout, sea-fastening and transport',5,'F020,MR20,L010','Cargo spread','Marine Manager','Sail-away release',off=True)
add('I020','Offshore','BHS-SLQ','Install new bridge, transfer services and commission',20,'L020,I010,J010,E050','Heavy-lift spread','Offshore Manager','New bridge and transferred services accepted BEFORE demolition',off=True)
add('D020','Demolition','BHS-SLQ','Isolate, dismantle old bridge and remove to shore',7,'I020','Heavy-lift spread','Offshore Manager','Approved demolition sequence / service isolation',off=True)
add('L030','Marine','ICG-ICP LD','Loadout, sea-fastening and transport',5,'F030,MR20,L020','Cargo spread','Marine Manager','Sail-away release',off=True)
add('I030','Offshore','ICG-ICP LD','Install, connect, inspect and accept replacement bridge',18,'L030,D020,J020','Heavy-lift spread','Offshore Manager','Services / safe access accepted',off=True)
add('D030','Demolition','OTHER BRIDGES','Other removal, waste backload and allocated demolition closure',7,'I030,'+crewA[-1]+','+crewB[-1],'Heavy-lift spread','Offshore Manager','Bridge-specific quantities and method to verify',off=True)
add('O090','Offshore','ALL','Final offshore inspections, punch closure and operational acceptance',7,'D030','Support spread','Completions Manager','All work required for completion accepted',off=True)
add('M100','Milestone','ALL','Planned offshore works complete',0,'O090',owner='Project Director')
add('W010','Allowance','ALL','Offshore weather / access allowance (not productive work)',10,'M100','Reserved marine capacity','Marine Manager','10 days is a planning allowance; not a weather study',off=True)
add('O100','Marine','ALL','Offshore demobilisation',3,'W010','Support spread','Marine Manager','Clearance / backload',off=True)
add('C010','Closeout','ALL','As-built dossiers, e-MB records, QA records and handover compilation',30,'M100','Completions team','QA / Completions')
add('C020','Closeout','ALL','Final document review and completion certificate processing',14,'C010,O100','Project team','Project Director','ONGC acceptance / contractual completion criteria')
add('M200','Milestone','ALL','Internal target: all works complete and handover accepted',0,'C020',owner='Project Director')
# CPM against internal completion only; excludes contractual deadline and warranty.
finish=A[-1]['ef']
for a in reversed(A):
 succ=[b for b in A if a['id'] in b['preds']];a['lf']=min([b['ls'] for b in succ],default=finish);a['ls']=a['lf']-a['dur'];a['float']=a['ls']-a['es']
for a in A:
 a['start']=START+timedelta(days=a['es']);a['finish']=START+timedelta(days=a['ef']-1 if a['dur'] else a['ef'])
endseason=date(2027,5,15)
print('Offshore complete',next(a['finish'] for a in A if a['id']=='M100'),'Demob',next(a['finish'] for a in A if a['id']=='O100'),'Completion',A[-1]['finish'])
print('OFFSHORE MONSOON EXCEPTIONS',[(a['id'],str(a['start']),str(a['finish'])) for a in A if a['off'] and a['finish']>endseason])
# Save data for render step
import json
Path('/workspace/project-review/schedule_data.json').write_text(json.dumps(A,default=str))
