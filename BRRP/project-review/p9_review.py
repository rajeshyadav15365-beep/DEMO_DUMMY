import json,hashlib,pathlib,fitz
root=pathlib.Path('/workspace/project-review'); out=root/'outputs/BRRP_P9_Input_Review';out.mkdir(parents=True,exist_ok=True)
rows=json.loads((root/'p9_batch_manifest.json').read_text()); current={x['result']['path'] for x in rows}; prior={}
for p in pathlib.Path('/workspace/attachments').rglob('*.pdf'):
 if str(p) not in current:
  h=hashlib.sha256(p.read_bytes()).hexdigest();prior.setdefault(h,[]).append(str(p))
seen={};idx=[]
for r in rows:
 p=pathlib.Path(r['result']['path']);h=hashlib.sha256(p.read_bytes()).hexdigest();d=fitz.open(p)
 r.update(sha256=h,pages=len(d),prior_matches=prior.get(h,[]),batch_duplicate_of=seen.get(h));seen.setdefault(h,r['name'])
 pages=[{'page':i+1,'text':' '.join(pg.get_text().split())} for i,pg in enumerate(d)]
 idx.append({'name':r['name'],'path':str(p),'pages':pages})
 if not r['prior_matches'] and not r['batch_duplicate_of']:
  (out/(p.stem+'.txt')).write_text('\n\n'.join('PAGE '+str(x['page'])+'\n'+x['text'] for x in pages))
(out/'inventory.json').write_text(json.dumps(rows,indent=2));(out/'page_index.json').write_text(json.dumps(idx))
print(json.dumps([{'name':r['name'],'pages':r['pages'],'prior':len(r['prior_matches']),'duplicate':r['batch_duplicate_of']} for r in rows],indent=2))
