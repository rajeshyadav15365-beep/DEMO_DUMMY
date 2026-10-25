from pathlib import Path
from datetime import date,timedelta
import json,heapq,math,csv
R=Path('/workspace/project-review');O=R/'outputs/BRRP_Integrated_Project_P5';START=date(2026,11,1);DEADLINE=date(2027,10,31)
old=json.loads((R/'final_schedule_data.json').read_text());parents=old['recovery']
# Break down the existing proposed scenario without inventing new resource commitments.
splits={
'G010':[(2,'Contract kick-off and obligations'),(2,'Leadership and authority assignments'),(3,'Guarantee insurance and funding arrangements')],
'G020':[(4,'Source/addendum register and scope boundaries'),(5,'Bridge quantity and demolition reconciliation'),(5,'Interface register and baseline assumptions')],
'S010':[(2,'Survey scope and inspection plan'),(3,'Access permits and operational coordination'),(2,'Survey spread readiness')],
'S020':[(3,'IC complex survey'),(3,'BHS complex survey'),(4,'Other bridge survey campaign'),(4,'Survey data validation and accepted report')],
'E010':[(4,'Criteria hierarchy and code register'),(5,'Loads and environmental basis'),(5,'Design-basis/interface review')],
'E020':[(3,'Preliminary checked quantities'),(3,'Material requisitions and specifications'),(4,'Technical procurement release')],
'P010':[(3,'Vendor approval and purchase authorisation'),(15,'Material manufacturing and expediting'),(4,'Inspection and release documentation'),(3,'Phased receipt and traceability')],
'E030':[(5,'Replacement geometry and model preparation'),(7,'Global strength and serviceability analysis'),(5,'Connections bearings and landing checks'),(4,'Marine constructability and services review'),(4,'Engineering review comments and release')],
'E040':[(5,'Existing condition and strengthening models'),(7,'Member joint and platform interface checks'),(5,'Jacket addendum and temporary conditions'),(4,'Workpack drawings and MTO'),(4,'Engineering review and release')],
'E050':[(7,'BHS service register and operational interfaces'),(10,'New service-support and transfer engineering'),(10,'Temporary stability and demolition design'),(8,'Multidiscipline review and ONGC comments')],
'MR10':[(7,'Marine campaign requirements'),(14,'Vessel enquiries and technical audits'),(7,'Availability and charter evaluation'),(7,'Conditional reservations and assurance plan')],
'MR20':[(7,'Module boundaries weight and COG basis'),(10,'Lift and landing structural engineering'),(10,'Loadout transport and sea-fastening engineering'),(8,'Marine assurance and review')],
'F010':[(3,'Traceable material preparation'),(5,'Assembly and fit-up'),(15,'Welding and production'),(4,'NDT and repair closure'),(3,'Dimensional acceptance'),(5,'Coating preservation and yard release')],
'F020':[(5,'Traceable material preparation'),(8,'Assembly and fit-up'),(25,'Welding and production'),(7,'NDT and repair closure'),(5,'Dimensional acceptance'),(5,'Coating preservation and yard release')],
'F030':[(3,'Traceable material preparation'),(7,'Assembly and fit-up'),(18,'Welding and production'),(5,'NDT and repair closure'),(3,'Dimensional acceptance'),(4,'Coating preservation and yard release')],
'O010':[(3,'Accommodation/support spread mobilisation'),(3,'Offshore crew competence and emergency readiness'),(4,'Permits SIMOPS and workfront release')],
'J010':[(2,'ICG/BHS access and approved temporary conditions'),(7,'Addendum jacket strengthening'),(3,'NDT acceptance and installation-interface release')],
'J020':[(2,'ICP access and approved temporary conditions'),(6,'Addendum jacket strengthening'),(2,'NDT acceptance and installation-interface release')],
'H010':[(3,'Heavy-lift spread mobilisation'),(2,'Equipment/certification verification'),(2,'Operational readiness and lift release')],
'I010':[(2,'Landing and interface readiness'),(4,'Approved lift and bridge landing'),(5,'Structural connections and access fittings'),(3,'Inspection coating restoration and services checks'),(2,'Package acceptance')],
'I020':[(2,'BHS landing and operational readiness'),(4,'Approved lift and new-bridge landing'),(4,'Structural connections and inspection'),(6,'Service transfer and commissioning'),(2,'Operational acceptance before demolition'),(2,'Package release')],
'D020':[(1,'Demolition release and isolation confirmation'),(4,'Approved dismantling and controlled removal'),(2,'Backload and clearance')],
'I030':[(2,'Landing and interface readiness'),(5,'Approved lift and bridge landing'),(6,'Structural connections and access fittings'),(3,'Inspection and service checks'),(2,'Package acceptance')],
'D030':[(2,'Other-bridge demolition quantity and release check'),(3,'Approved residual removal'),(2,'Waste/backload and records')],
'O090':[(3,'Final offshore walkdowns'),(3,'Acceptance-critical punch closure'),(1,'Operational acceptance records')],
'C010':[(10,'As-built and red-line consolidation'),(10,'Quality and commissioning dossiers'),(5,'Accepted measurement and e-MB records'),(5,'Handover submission')],
'C020':[(7,'Client review and comments'),(5,'Document and residual acceptance closure'),(2,'Completion certificate processing')],
'O100':[(1,'Demobilisation clearance'),(1,'Personnel and equipment backload'),(1,'Spread departure and closeout records')]
}
for pid in ['L010','L020','L030']:splits[pid]=[(2,'Loadout and sea-fastening release'),(2,'Marine transport'),(1,'Site receipt and installation readiness')]
for p in parents:
 if p['id'].startswith(('RA','RB')):
  d=p['dur'];splits[p['id']]=[(2,'Permit access and material readiness'),(2,'Approved temporary works and preparation'),(d-8,'Strengthening installation and connections'),(2,'NDT dimensional and coating acceptance'),(2,'Punch closure and package acceptance')]
# First-batch strengthening release is a milestone within one production stream, not duplicate production.
raw=[]
def add(i,name,d,preds,p,resource=None):
 raw.append(dict(id=i,parent=p['id'],name=name,duration=d,preds=list(preds),phase=p['phase'],package=p['pkg'],resource=p['resource'] if resource is None else resource,owner=p['owner'],offshore=p['off'],status='PROPOSED / UNVERIFIED',basis='Illustrative duration allocation within prior scenario; validate measured quantities, productivity and approvals'))
for p in parents:
 pid=p['id']
 if pid=='F041':continue
 if pid=='F040':
  subs=[(3,'Strengthening kit material preparation'),(5,'First-batch assembly and fit-up'),(8,'First-batch production'),(4,'First-batch inspection coating and release'),(10,'Remaining kit production'),(5,'Remaining kit inspection coating and release')]
  prev=p['preds']
  for j,(dur,title) in enumerate(subs,1):
   nid=f'{pid}.{j:02}';add(nid,title,dur,prev,p);prev=[nid]
   if j==4:
    pp=dict(p);pp['id']='F041';add('F041','First approved strengthening kits available',0,[nid],pp,resource='Milestone within strengthening-cell production')
  add(pid,'All strengthening kits available',0,prev,p,resource='Milestone');continue
 if p['duration'] if 'duration' in p else False:pass
 if p['dur']==0:add(pid,p['name'],0,p['preds'],p,resource='Milestone');continue
 subs=splits.get(pid,[(p['dur'],p['name'])]);assert sum(x[0] for x in subs)==p['dur'],pid
 prev=p['preds']
 for j,(dur,title) in enumerate(subs,1):
  nid=f'{pid}.{j:02}';add(nid,title,dur,prev,p);prev=[nid]
 add(pid,p['name']+' - phase complete',0,prev,p,resource='Milestone')
by={x['id']:x for x in raw};assert len(by)==len(raw)
# Correct predecessor coverage after crew-C is introduced.
by['D030.01']['preds']=list(dict.fromkeys(by['D030.01']['preds']+['RA03']))

def first_fit(s,d,off):
 if not off or not d:return s
 while True:
  sd=START+timedelta(days=s);ed=sd+timedelta(days=d-1);blocked=None
  for year in range(sd.year,ed.year+1):
   if sd<=date(year,10,15) and ed>=date(year,5,16):blocked=(date(year,10,16)-START).days;break
  if blocked is None:return s
  s=blocked

def last_fit(f,d,off):
 if not off or not d:return f-d
 while True:
  s=f-d;sd=START+timedelta(days=s);ed=START+timedelta(days=f-1);blocked=None
  for year in range(sd.year,ed.year+1):
   if sd<=date(year,10,15) and ed>=date(year,5,16):blocked=(date(year,5,16)-START).days;break
  if blocked is None:return s
  f=blocked

def solve(items):
 b={x['id']:x for x in items};succ={i:[] for i in b};ind={i:0 for i in b};order=[]
 for x in items:
  for p in x['preds']:
   assert p in b,(x['id'],p);succ[p].append(x['id']);ind[x['id']]+=1
 ready=sorted(i for i,n in ind.items() if not n)
 while ready:
  i=ready.pop(0);order.append(i)
  for j in succ[i]:
   ind[j]-=1
   if not ind[j]:ready.append(j);ready.sort()
 assert len(order)==len(items),'Cycle in schedule'
 for i in order:
  x=b[i];r=max([x.get('earliest',0)]+[b[p]['EF'] for p in x['preds']]);x['ES']=first_fit(r,x['duration'],x['offshore']);x['EF']=x['ES']+x['duration'];x['calendar_wait']=x['ES']-r
 finish=b['M200']['EF'];deadline=(DEADLINE+timedelta(days=1)-START).days
 for prefix,target in [('',finish),('contract_',deadline)]:
  for i in reversed(order):
   x=b[i];lf=min([b[j][prefix+'LS'] for j in succ[i]],default=target);ls=last_fit(lf,x['duration'],x['offshore']);x[prefix+'LS']=ls;x[prefix+'LF']=ls+x['duration'];x[prefix+'float']=ls-x['ES']
 for i in order:
  x=b[i];x['start']=str(START+timedelta(days=x['ES']));x['finish']=str(START+timedelta(days=x['EF']-1 if x['duration'] else x['EF']));x['critical']=x['float']==0
  assert x['float']>=0,(i,x['float'])
  assert all(x['ES']>=b[p]['EF'] for p in x['preds'])
  assert first_fit(x['ES'],x['duration'],x['offshore'])==x['ES']
 # Select one driving chain, while separately recording every zero-float activity.
 chain=[];cur='M200'
 while True:
  chain.append(cur);ps=b[cur]['preds']
  if not ps:break
  cand=[p for p in ps if b[p]['critical']]
  cur=max(cand or ps,key=lambda p:b[p]['EF'])
 chain.reverse()
 return [b[i] for i in order],chain
schedule,chain=solve(raw)
# Baseline comparator uses two strengthening crews with otherwise equal assumptions.
import copy
comp=copy.deepcopy(raw)
for x in comp:
 if x['parent']=='RA04':x['resource']='Strengthening crew A' if x['duration'] else 'Milestone'
 if x['id']=='RA04.01':x['preds']=list(dict.fromkeys(x['preds']+['RA03']))
comparison_initial,_=solve(copy.deepcopy(comp))
# Explicit second-season readiness and remobilisation in the comparator.
season_day=(date(2027,10,16)-START).days
resume_ids={x['id'] for x in comparison_initial if x['offshore'] and x['duration'] and x['ES']>=season_day}
if resume_ids:
 template=dict(id='M2OPEN',parent='SECOND_SEASON',name='Second offshore season opens - planning constraint',duration=0,preds=['M000'],phase='Marine',package='ALL',resource='Milestone',owner='Marine Manager',offshore=False,status='PROPOSED / UNVERIFIED',basis='Explicit planning reopening constraint; actual weather and authorisations control',earliest=season_day)
 mob=dict(template);mob.update(id='M2MOB',name='Second-season marine/crew remobilisation and renewed readiness',duration=7,preds=['M2OPEN'],resource='Second-season marine spread',offshore=True,basis='7-day provisional remobilisation; suppliers, vessels and cost unconfirmed')
 comp.extend([template,mob])
 for x in comp:
  if x['id'] in resume_ids:x['preds']=list(dict.fromkeys(x['preds']+['M2MOB']))
comparison,compchain=solve(comp)
# Verify exclusive named-resource reservations; this is sequencing, not full labour levelling.
exclusive=['Fabrication line 1','Fabrication line 2','Dedicated strengthening cell','Strengthening crew A','Strengthening crew B','Strengthening crew C (additional)','Heavy-lift spread','Cargo spread']
for scen in [schedule,comparison]:
 for resource in exclusive:
  series=sorted([x for x in scen if x['resource']==resource and x['duration']],key=lambda x:x['ES'])
  assert all(y['ES']>=x['EF'] for x,y in zip(series,series[1:])),resource
bridge_data=[['ICW-ICG UD','Strengthening',0,19,0,30,25,110,58,273],['ICW-ICG LD','New/replacement',57,0,65,0,1,118,88,0],['ICG-ICP UD','Strengthening',0,86,0,30,25,93,49,419],['ICG-ICP LD','New/replacement',115,0,45,0,1,176,132,0],['BHS-SLQ','Replacement/services',195,0,87,0,210,88,82,332],['SLQ-WIS','Strengthening',0,61,0,30,15,98,51,0],['NQO-NQD','Strengthening',0,63,0,30,25,93,58,250],['WIN-NC','Strengthening',0,6,0,40,10,45,29,0],['SCA-SC1','Strengthening',0,35,0,28,15,88,55,0],['SCA-SCF','Strengthening',0,91,0,28,30,257,116,0]]
actual_totals=[sum(x[j] for x in bridge_data) for j in range(2,10)];print('MTO ROW SUMS',actual_totals);assert actual_totals==[367,361,197,216,357,1166,718,1274]
summary={s['id']:s for s in schedule};csummary={s['id']:s for s in comparison}
result={'start':str(START),'contract_last_workday':str(DEADLINE),'schedule':schedule,'critical_chain':chain,'comparison':comparison,'comparison_chain':compchain,'bridge_mto':bridge_data,'mto_actual_totals':actual_totals,'statistics':{'activities':len(schedule),'nonzero':sum(x['duration']>0 for x in schedule),'milestones':sum(x['duration']==0 for x in schedule),'critical_nonzero':sum(x['critical'] and x['duration']>0 for x in schedule),'finish':summary['M200']['finish'],'two_crew_finish':csummary['M200']['finish']}}
(O/'P5_integrated_data.json').write_text(json.dumps(result,indent=2))
print(result['statistics']);print('Milestones',[(i,summary[i]['finish'],summary[i]['float'],summary[i]['contract_float']) for i in ['E030','P010','F010','F020','F030','I010','I020','D020','I030','M100','O100','M200']]);print('Critical parent groups',list(dict.fromkeys(summary[i]['parent'] for i in chain)))
# Sensitivity: increase one critical material-manufacturing activity, without changing baseline.
sensitivity=[]
for delay in [0,1,3,4,7,14]:
 test=copy.deepcopy(raw)
 for x in test:
  if x['id']=='P010.02':x['duration']+=delay
 calc,_=solve(test);cb={x['id']:x for x in calc}
 sensitivity.append({'extra_material_days':delay,'offshore_milestone':cb['M100']['finish'],'demob_phase_complete':cb['O100']['finish'],'overall_milestone':cb['M200']['finish'],'calendar_overrun_days':max(0,(date.fromisoformat(cb['M200']['finish'])-DEADLINE).days)})
result['sensitivity']=sensitivity;(O/'P5_integrated_data.json').write_text(json.dumps(result,indent=2))
print('Material-delay sensitivity',sensitivity)
