import pathlib,xml.etree.ElementTree as E,json,collections
p=max(pathlib.Path('/workspace/attachments').glob('*/SHIVANSH_PreAward_207_Activities_ProjectLibre_Import.xml'),key=lambda x:x.stat().st_mtime);root=E.parse(p).getroot();ns={'m':'http://schemas.microsoft.com/project'}
def val(x,k):return x.findtext('m:'+k,default='',namespaces=ns)
tasks=root.findall('m:Tasks/m:Task',ns);uids=[val(t,'UID')for t in tasks];links=[];missing=[]
for t in tasks:
 for l in t.findall('m:PredecessorLink',ns):
  r={'successor':val(t,'UID'),'predecessor':val(l,'PredecessorUID'),'type':val(l,'Type'),'lag':val(l,'LinkLag')};links.append(r)
  if r['predecessor']not in uids:missing.append(r)
adj={u:[]for u in uids};indeg={u:0 for u in uids}
for l in links:
 if l['predecessor']in adj:adj[l['predecessor']].append(l['successor']);indeg[l['successor']]+=1
q=collections.deque(u for u,n in indeg.items()if n==0);visited=[]
while q:
 u=q.popleft();visited.append(u)
 for v in adj[u]:
  indeg[v]-=1
  if indeg[v]==0:q.append(v)
r={'path':str(p),'title':val(root,'Title'),'start':val(root,'StartDate'),'minutes_per_day':val(root,'MinutesPerDay'),'task_count':len(tasks),'non_summary_task_count':sum(val(t,'Summary')!='1'for t in tasks),'links':len(links),'link_types':dict(collections.Counter(l['type']for l in links)),'missing_predecessors':missing,'duplicate_UIDs':[u for u,n in collections.Counter(uids).items()if n>1],'unresolved_topological_nodes':[u for u in uids if u not in visited],'resources':len(root.findall('m:Resources/m:Resource',ns)),'assignments':len(root.findall('m:Assignments/m:Assignment',ns)),'calendar_exceptions':len(root.findall('.//m:Exception',ns)),'root_tags':[x.tag.split('}')[-1]for x in root],'examples':[{'uid':val(t,'UID'),'name':val(t,'Name'),'duration':val(t,'Duration'),'start':val(t,'Start'),'finish':val(t,'Finish'),'constraint':val(t,'ConstraintType'),'notes':val(t,'Notes')}for t in tasks[:4]],'tasks':[{'uid':val(t,'UID'),'name':val(t,'Name'),'duration':val(t,'Duration'),'start':val(t,'Start'),'finish':val(t,'Finish')}for t in tasks],'link_records':links}
out=pathlib.Path('/workspace/project-review/outputs/BRRP_P13_Master_Plan_Review');(out/'XML_schedule_audit.json').write_text(json.dumps(r,indent=2));print(json.dumps({k:v for k,v in r.items()if k not in ['tasks','link_records']},indent=2))
