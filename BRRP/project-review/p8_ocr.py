from pathlib import Path
import json,subprocess,concurrent.futures,os,re
import fitz
R=Path('/workspace/project-review');O=R/'p8_extracts';rows=json.loads((R/'outputs/BRRP_P8_Input_Audit/P8_file_inventory.json').read_text());jobs=[]
for r in rows:
 if r['duplicate_in_batch']:continue
 for n in r['low_text_pages']:jobs.append((r['name'],r['path'],n))
def run(job):
 name,path,n=job;slug=re.sub(r'[^A-Za-z0-9]+','_',Path(name).stem);png=O/f'OCR_{slug}_p{n}.png';base=O/f'OCR_{slug}_p{n}'
 d=fitz.open(path);p=d[n-1];p.get_pixmap(matrix=fitz.Matrix(2,2),alpha=False).save(png)
 env=dict(os.environ);env['OMP_THREAD_LIMIT']='1'
 r=subprocess.run(['tesseract',str(png),str(base),'--psm','3'],stdout=subprocess.PIPE,stderr=subprocess.PIPE,env=env,timeout=90)
 text=base.with_suffix('.txt').read_text() if base.with_suffix('.txt').exists() else ''
 return {'name':name,'page':n,'ocr_path':str(base.with_suffix('.txt')),'chars':len(text),'exit':r.returncode,'snippet':re.sub(r'\s+',' ',text)[:650]}
with concurrent.futures.ThreadPoolExecutor(max_workers=4) as ex:out=list(ex.map(run,jobs))
(R/'outputs/BRRP_P8_Input_Audit/P8_OCR_inventory.json').write_text(json.dumps(out,indent=2))
for x in out:print(x)
