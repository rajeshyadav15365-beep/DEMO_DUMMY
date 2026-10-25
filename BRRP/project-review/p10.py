import pathlib,json,hashlib,fitz,openpyxl,csv
p=pathlib.Path('/workspace/project-review/outputs/BRRP_P10_Input_Review');p.mkdir(exist_ok=True); rows=json.loads(pathlib.Path('/workspace/project-review/p10_manifest.json').read_text());curr={r['result']['path'] for r in rows};prior={}
for f in pathlib.Path('/workspace/attachments').rglob('*'):
 if f.is_file() and str(f) not in curr and f.suffix.lower() in ['.pdf','.xlsx','.csv']:prior.setdefault(hashlib.sha256(f.read_bytes()).hexdigest(),[]).append(str(f))
seen={}; extracts=[]
for r in rows:
 f=pathlib.Path(r['result']['path']);h=hashlib.sha256(f.read_bytes()).hexdigest();r.update(sha256=h,prior=prior.get(h,[]),duplicate=seen.get(h));seen.setdefault(h,r['name']);e={'name':r['name']}
 if f.suffix=='.pdf':
  d=fitz.open(f);e['pages']=[{'page':i+1,'text':' '.join(pg.get_text().split())} for i,pg in enumerate(d)];r['pages']=len(d);print(r['name'],'pages',len(d),'prior',len(r['prior']));print('START',e['pages'][0]['text'][:700])
 elif f.suffix=='.xlsx':
  w=openpyxl.load_workbook(f,data_only=False);v=openpyxl.load_workbook(f,data_only=True);e['sheets']=[]
  for s in w:
   data=[{'row':row[0].row,'cells':{c.coordinate:str(c.value) for c in row if c.value is not None}} for row in s if any(c.value is not None for c in row)];forms=[c for row in s for c in row if c.data_type=='f'];e['sheets'].append({'name':s.title,'rows':s.max_row,'cols':s.max_column,'data':data,'formulas':len(forms),'uncached':sum(v[s.title][c.coordinate].value is None for c in forms)})
  print(r['name'],'prior',len(r['prior']),'duplicate',r['duplicate'],[(s['name'],s['rows'],s['formulas'],s['uncached']) for s in e['sheets']])
 else:e['csv']=list(csv.DictReader(f.open(encoding='utf-8-sig')));print(r['name'],'prior',len(r['prior']),'rows',len(e['csv']), 'sample',e['csv'][:2])
 extracts.append(e)
(p/'inventory.json').write_text(json.dumps(rows,indent=2));(p/'extracts.json').write_text(json.dumps(extracts,indent=2))
