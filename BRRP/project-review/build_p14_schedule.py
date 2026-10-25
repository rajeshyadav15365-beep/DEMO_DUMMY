import json,pathlib,datetime as dt,math,collections,csv,os
STEEL_LEAD=int(os.getenv("P14_STEEL_LEAD","45"));FAB_RATE=float(os.getenv("P14_FAB_RATE","4"));BAYS=int(os.getenv("P14_BAYS","2"));SCENARIO=os.getenv("P14_SCENARIO","BASE")
out=pathlib.Path('/workspace/project-review/outputs/BRRP_P14_Consolidated_Management');out.mkdir(exist_ok=True)
if SCENARIO!='BASE':out=out/SCENARIO;out.mkdir(exist_ok=True)
old=json.loads(pathlib.Path('/workspace/project-review/outputs/BRRP_Integrated_Project_P5/P5_integrated_data.json').read_text());bridges=old['bridge_mto'];start=dt.date(2026,11,1)
def work(d,cal):
 if cal=='OFF':return not (dt.date(d.year,5,16)<=d<=dt.date(d.year,10,15))
 if cal=='OFFICE':return d.weekday()<5
 if cal=='YARD':return d.weekday()<6
 return True
def dates(earliest,dur,cal):
 if dur==0:return earliest,earliest
 s=earliest
 while not work(s,cal):s+=dt.timedelta(days=1)
 # Offshore task is indivisible: all productive days must fit one open season.
 while True:
  d=s;n=0
  while n<dur:
   if work(d,cal):n+=1
   d+=dt.timedelta(days=1)
  if cal!='OFF' or all(work(s+dt.timedelta(days=i),cal)for i in range((d-s).days)):return s,d
  s=dt.date(s.year,10,16) if s<dt.date(s.year,10,16) else dt.date(s.year+1,10,16)
def latest(lf,dur,cal):
 if dur==0:return lf
 d=lf-dt.timedelta(days=1)
 while True:
  if work(d,cal):
   s=d;n=1
   while n<dur:
    s-=dt.timedelta(days=1)
    if work(s,cal):n+=1
   if cal!='OFF' or all(work(s+dt.timedelta(days=i),cal)for i in range((d-s).days+1)):return s
  d-=dt.timedelta(days=1)
tasks=[];last={}
def add(id,name,dur,pre,cal='OFFICE',bridge='COMMON',resource='',owner='Project Manager',basis='Proposed allowance; replace with verified inputs',source='Management planning'):
 ps=list(pre)
 if resource and resource in last:ps.append(last[resource])
 r={'id':id,'name':name,'duration':dur,'preds':list(dict.fromkeys(ps)),'calendar':cal,'bridge':bridge,'resource':resource,'owner':owner,'basis':basis,'source':source,'status':'PROPOSED / NOT APPROVED'};tasks.append(r)
 if resource:last[resource]=id
 return id
add('START','Assumed commencement — scenario only',0,[],cal='ALL',owner='Contracts',basis='1Nov2026 assumed; actual NOA unconfirmed')
add('G01','Contract/amendment matrix and scope reconciliation',5,['START'],owner='Contracts',source='P8–P12 source reviews')
add('G02','Survey permits, access and mobilisation preparation',10,['G01'],owner='Survey / Operations')
add('G03','Pre-engineering survey campaign',15,['G02'],'OFF',resource='Survey spread',owner='Survey',source='Technical BEC; survey requirement')
add('G04','Survey reports, geometry/condition reconciliation',5,['G03'],owner='Engineering')
add('G05','Survey/basis review and comment closure allowance',10,['G04'],owner='Engineering',source='C8 review rules; specific survey categorisation to confirm')
# Common packages are screened scope items, no orders placed.
common=[('BEAR','Bearings, sliding interfaces and guides','Engineering','G05',35),('GRAT','Grating, handrail and GI sheet supply','Procurement','G05',35),('PIPE','Piping tie-ins, fittings and supports','Piping','G05',45),('ELEC','Electrical materials and bonding','Electrical','G05',40),('INST','Instrumentation/cable interface materials','Instrumentation','G05',40),('COAT','Qualified coating system supply','Coating','G05',25),('NDT','NDT, welding qualification and consumables','QA','G05',20),('MAR','Heavy-lift and accommodation spread readiness','Marine','G05',45),('CARGO','Cargo vessel and shore-base readiness','Logistics','G05',25),('SURV','Survey services mobilisation','Survey','G01',10),('DES','Design agency capacity confirmation','Engineering','G01',10),('TPI','TPI/CA services and audit arrangements','QA','G01',10),('SCAF','Scaffolding/access and temporary works supply','Construction','G05',25),('PPE','Certified PPE and training readiness','HSE','G01',15),('TRAY','Cable trays and support steel interfaces','Electrical','G05',35)]
for code,name,owner,pre,dur in common:
 add('P-'+code,name+' — requisition/technical scope',2 if code=='SURV' else 5,[pre],owner=owner)
 add('R-'+code,name+' — bid/PS review allowance',3 if code=='SURV' else 10,['P-'+code],owner=owner,source='C8; PS applicability and review stage to confirm')
 add('D-'+code,name+' — delivery/readiness',5 if code=='SURV' else dur,['R-'+code],cal='ALL',owner=owner,basis='Unverified vendor/service lead-time allowance; no quote or commitment')
next(t for t in tasks if t['id']=='G03')['preds'].append('D-SURV')
# Priority grouping is a proposed construction sequence, not operational approval.
order=[4,1,3,0,2,5,6,7,8,9];ends=[];eng=[];fab=[];off=[]
for n,i in enumerate(order):
 b=bridges[i];code=f'B{i+1:02d}';name=b[0];q=sum(b[2:6]);new=b[1]!='Strengthening';dem=b[6]
 E=lambda x:code+'-'+x
 add(E('BAS'),'Design basis, loads and interfaces',5,['G05','D-DES'],bridge=name,owner='Engineering',source='Design Criteria PartII; P8 drawing matrix')
 add(E('BRV'),'Basis review and comment closure allowance',10,[E('BAS')],bridge=name,owner='Engineering',source='C8; ten working days used provisionally')
 add(E('ANA'),'Bridge/platform and preservice analysis preparation',20 if new else 12,[E('BRV')],bridge=name,owner='Structural',basis='Proposed engineering effort window; native analysis/check evidence not supplied')
 add(E('ARV'),'Analysis review and comment closure allowance',10,[E('ANA')],bridge=name,owner='Engineering',source='C8 bridge report row10workingdays; re-review risk separate')
 add(E('DET'),'Connection, bearing, landing and workpack detailing',15 if new else 10,[E('ARV')],bridge=name,owner='Structural')
 add(E('DRV'),'Drawing review and comment closure allowance',10,[E('DET')],bridge=name,owner='Engineering',source='C8; mother-analysis dependency preserved')
 add(E('AFC'),'AFC design/MTO independent check and release gate',0,[E('DRV')],cal='ALL',bridge=name,owner='Engineering',basis='Gate; actual signed release not established')
 eng.append(E('AFC'))
 add(E('RFQ'),'Bridge steel RFQ, evaluation and PS preparation',7,[E('BRV')],bridge=name,owner='Procurement')
 add(E('PRV'),'Bridge steel PS review allowance',10,[E('RFQ')],bridge=name,owner='Procurement',source='C8; full project MCI/two-stage applicability to confirm')
 add(E('PO'),'Bridge steel compliant PO release gate',0,[E('AFC'),E('PRV')],cal='ALL',bridge=name,owner='Procurement',basis='No at-risk purchase assumed; amended domestic-origin evidence required')
 add(E('MAT'),'Mill production, inspection and steel delivery',STEEL_LEAD,[E('PO')],cal='ALL',bridge=name,owner='Procurement',basis='Unquoted45calendar-day placeholder; origin, mill capacity and transport unresolved')
 bay='Fabrication line '+str(n%BAYS+1)
 add(E('FAB'),'Fit-up, welding and dimensional control',max(8,math.ceil(q/FAB_RATE)),[E('MAT'),'D-NDT','D-TPI'],'YARD',name,bay,'Fabrication',basis='Assumed4MT per productive bay-day; includes aggregate steel, not validated shop productivity')
 add(E('NDT'),'NDT, repair allowance and dimensional release',6,[E('FAB')],'YARD',name,owner='QA')
 add(E('COA'),'Surface preparation, coating and curing allowance',10,[E('NDT'),'D-COAT'],'YARD',name,'Coating cell','Coating',source='FS2005Rev9; actual system/curing to govern')
 add(E('REL'),'Trial fit, verified weight/COG and loadout release',5,[E('COA')],'YARD',name,owner='QA / Marine')
 fab.append(E('REL'))
 add(E('TRN'),'Loadout, transport and offshore receipt',3,[E('REL'),'D-CARGO'],'OFF',name,'Cargo spread','Logistics')
 add(E('PRE'),'PTW, access, SIMOPS and work-front readiness',3,[E('TRN'),'D-SCAF','D-PPE'],'OFF',name,owner='Construction')
 add(E('STR'),'Landing/platform pre-install strengthening',max(3,math.ceil((b[4]+b[5])/4)),[E('PRE')],'OFF',name,'Strength crew '+str(n%3+1),'Construction',basis='Proposed4MT/crew-day; actual weld/access quantities not validated')
 install=E('INS')
 add(install,'New bridge installation' if new else 'Bridge strengthening and replacement items',7 if new else max(5,math.ceil(b[3]/4)),[E('STR'),'D-MAR','D-BEAR','D-GRAT'],'OFF',name,'Heavy-lift spread' if new else 'Strength crew '+str(n%3+1),'Marine / Construction',basis='Aggregate planning allowance; engineered lifts and temporary stability required')
 add(E('SRV'),'Services transfer, test and accepted reinstatement',5,[install,'D-PIPE','D-ELEC','D-INST','D-TRAY'],'OFF',name,'Service transfer team','Commissioning')
 add(E('DEM'),'Engineered demolition/removal and scrap backload',max(2,math.ceil(dem/15)),[E('SRV')],'OFF',name,'Removal team','Construction',basis='Proposed15MT/day allowance; BHS sequence after safe new service transfer, subject to approved method')
 add(E('PUN'),'Inspection, touch-up and offshore punch closure',3,[E('DEM')],'OFF',name,owner='QA / Commissioning');off.append(E('PUN'))
 add(E('DOS'),'As-built, QA and turnover dossier acceptance',10,[E('PUN')],bridge=name,owner='Document Control')
 ends.append(E('DOS'))
add('OFFDONE','All bridge offshore work completed',0,off,cal='ALL')
add('DEMOB','Spread demobilisation allowance',5,['OFFDONE'],'OFF',resource='Cargo spread',owner='Marine')
add('CLOSE','Integrated acceptance, final measurements and account preparation',10,ends+['DEMOB'],owner='Project Manager')
add('FINISH','Proposed contract-completion milestone',0,['CLOSE'],cal='ALL',owner='Project Director')
by={t['id']:t for t in tasks};adj={k:[]for k in by};degree={k:0 for k in by}
for t in tasks:
 for pre in t['preds']:
  assert pre in by,(t['id'],pre);adj[pre].append(t['id']);degree[t['id']]+=1
queue=collections.deque(k for k,v in degree.items()if not v);top=[]
while queue:
 k=queue.popleft();top.append(k)
 for v in adj[k]:
  degree[v]-=1
  if not degree[v]:queue.append(v)
assert len(top)==len(tasks),'Cycle'
for k in top:
 t=by[k];earliest=max([by[x]['EF']for x in t['preds']]+[start]);t['ES'],t['EF']=dates(earliest,t['duration'],t['calendar']);t['driving_predecessor']=max(t['preds'],key=lambda x:by[x]['EF'])if t['preds']else None
assert [k for k in by if not adj[k]]==['FINISH'], 'Unconnected finish'
assert [k for k in by if not by[k]['preds']]==['START'], 'Unconnected start'
finish=by['FINISH']['EF']
for k in reversed(top):
 t=by[k];t['LF']=min([by[x]['LS']for x in adj[k]]or[finish]);t['LS']=latest(t['LF'],t['duration'],t['calendar']);t['total_float_calendar_days']=(t['LS']-t['ES']).days;assert t['total_float_calendar_days']>=0,(k,t)
for t in tasks:
 assert all(t['ES']>=by[x]['EF']for x in t['preds'])
for t in tasks:
 for key in ['ES','EF','LS','LF']:t[key]=t[key].isoformat()
chain=[];k='FINISH'
while k:
 chain.append(k);k=by[k]['driving_predecessor']
chain.reverse()
res={'scenario':SCENARIO,'steel_lead_days':STEEL_LEAD,'fabrication_rate_MT_per_bay_day':FAB_RATE,'fabrication_bays':BAYS,'start_assumption':str(start),'finish':str(finish),'contract_anniversary_assumption':str(dt.date(2027,11,1)),'calendar_days_after_assumed_anniversary':(finish-dt.date(2027,11,1)).days,'tasks':tasks,'driving_chain':chain,'common_packages':common,'bridge_mto':bridges,'calendars':{'OFFICE':'Mon–Fri8h; holidays absent','YARD':'Mon–Sat; unverified shift/capacity','OFF':'7days;16May–15Oct blocked as planning scenario, indivisible work phase must fit open season','ALL':'7calendar days'},'method':'FS0 DAG plus explicit resource sequencing; forward/backward calendar passes; finite-resource scenario only for named capacity-one chains; not full P6 resource levelling or risk analysis'}
(out/'P14_schedule_data.json').write_text(json.dumps(res,indent=2))
with (out/'BRRP_P14_Integrated_Schedule.csv').open('w')as f:
 w=csv.writer(f);w.writerow(['ID','Bridge','Activity','Duration','Calendar','Predecessors','Resource','Owner','Start','Finish exclusive','Total float calendar days','Status','Basis'])
 for t in tasks:w.writerow([t['id'],t['bridge'],t['name'],t['duration'],t['calendar'],';'.join(t['preds']),t['resource'],t['owner'],t['ES'],t['EF'],t['total_float_calendar_days'],t['status'],t['basis']])
print(json.dumps({'activities':len(tasks),'finish':str(finish),'days_after_assumed_anniversary':res['calendar_days_after_assumed_anniversary'],'driving_chain_length':len(chain),'minimum_float':min(t['total_float_calendar_days']for t in tasks)}))
