p='/workspace/project-review/publish_p14.py';s=open(p).read();s=s.replace('import pathlib,json,math,datetime as dt,csv,sys,shutil,collections,html','import pathlib,json,math,datetime as dt,csv,sys,shutil,collections,html,re,ast')
s=s.replace("sheet('Requirements',['ID'", """# Carry forward historical audit actions with later-disposition precedence.
for x in json.loads((p.parent/'BRRP_P8_Input_Audit/P8_findings.json').read_text()):
 req.append(['P8-'+x[0],x[1],x[2]+' PDF page '+x[3],x[4],x[6],x[5],'Historical finding; later P9/P11 disposition governs'])
for label,script in [('P10','publish_p10.py'),('P11','publish_p11.py'),('P13','publish_p13.py')]:
 tree=ast.parse(pathlib.Path('/workspace/project-review',script).read_text())
 node=next(n for n in tree.body if isinstance(n,ast.Assign)and any(isinstance(t,ast.Name)and t.id=='findings'for t in n.targets))
 for n,x in enumerate(ast.literal_eval(node.value),1):
  req.append([f'{label}-{n:02d}',x[0],x[1],x[2],'Workstream accountable owner',x[3],'Carry-forward action; requires actual evidence'])
sheet('Requirements',['ID'""")
# Insert integration map and forecast sheet before formatting all sheets.
marker='for s in w:\n s.freeze_panes'
insert="""sheet('Stage Integration',['Uploaded model','P14 control','Required correction or integration'],[
['Stage27','Cost Commitments / Requirements / Daily Production','Separate quantity/EV, certified billing and cash; maintain unknown inputs'],
['Stage28','Schedule285 / Calendars and Method','Replace illustrative120activity chain with traceable integrated dependencies; actualL4detail still needs constructability'],
['Stage29','Schedule Scenarios / Resource Commitments','Scenario capacity is uncommitted; verify bottlenecks and cost recovery'],
['Stage30','Workpack Readiness / Daily Production','Verified quantities and accepted productivity required'],
['Stage31','Workpack Readiness / QA and Release Gates','Screen alone cannot authorise work; require evidence and signed release'],
['Stage32','Productivity Forecast / Schedule285','Productive crew-day estimate is not calendarCPM; route remaining workdays to validated network'],
['Stage33','Risk Register / KPI and Escalation','Track early-warning evidence and actual mitigation owner'],
['Stage34','CEO Decisions / Resource Commitments','Use one source of date/cost/progress truth; no inconsistent parallel baselines'],
['Stage35','90Day Mobilisation','D45firstAFCis an objective; not demonstrated by baseline assumptions'],
['Volumes1–6','A–Z Execution Plan and linked registers','Framework text remains planning; fill bridge-specific evidence and authority']])
forecast=[]
for i,b in enumerate(bridges,2):
 forecast.append([b[0],sum(b[2:6]),None,None,None,f'=IF(OR(D{i}=\"\",E{i}=\"\",E{i}<=0),\"\",D{i}/E{i})',f'=IF(OR(C{i}=\"\",D{i}=\"\"),\"\",MAX(0,C{i}-D{i}))',None,None,f'=IF(OR(F{i}=\"\",F{i}<=0,G{i}=\"\",H{i}=\"\",H{i}<=0,I{i}=\"\",I{i}<=0,I{i}>1),\"\",ROUNDUP(G{i}/F{i}/H{i}/I{i},0))',None,'No approved quantity/productivity yet'])
sheet('Productivity Forecast',['Bridge','Source indicative MT','Approved baseline MT','Accepted actual MT','Productive crew-days','Observed MT/productive crew-day','Remaining approved MT','Verified assigned crews','Future productive utilisation0–1','Remaining productive workdays','Updated CPMfinish input','Status/evidence'],forecast)
"""
s=s.replace(marker,insert+'\n'+marker)
s=s.replace("def para(txt,style='Normal'):\n", """def clean(txt):
 txt=re.sub(r'(?<=\\d)(?=[A-Za-z])',' ',txt)
 txt=re.sub(r'(?<=[a-z])(?=\\d)',' ',txt)
 txt=re.sub(r'(?<=[a-z])(?=[A-Z]{2})',' ',txt)
 return txt
def para(txt,style='Normal'):
 txt=clean(txt)
""")
s=s.replace("def heading(txt,level=1):\n", "def heading(txt,level=1):\n txt=clean(txt)\n")
s=s.replace("def P(t,sty='BodyText'):return Paragraph(html.escape(t),S[sty])", "def P(t,sty='BodyText'):return Paragraph(html.escape(clean(t)),S[sty])")
s=s.replace('[[None]*11for _ in range(40)]','[[None]*11 for _ in range(40)]')
open(p,'w').write(s)
