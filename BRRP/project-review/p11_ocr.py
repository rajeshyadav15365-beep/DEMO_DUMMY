import json,pathlib,fitz,subprocess,concurrent.futures,os
p=pathlib.Path('/workspace/project-review/outputs/BRRP_P11_Tender_Review');inv=json.loads((p/'inventory.json').read_text());r=next(r for r in inv if r['name']=='Vol_I_PartIV_ Technical BEC.pdf');d=fitz.open(r['path'])
for i,pg in enumerate(d):pg.get_pixmap(matrix=fitz.Matrix(2,2)).save(str(p/f'BEC_{i+1}.png'))
def run(n):
 subprocess.run(['tesseract',str(p/f'BEC_{n}.png'),str(p/f'BEC_{n}'),'--psm','3'],capture_output=True,check=True,env={**os.environ,'OMP_THREAD_LIMIT':'1'});return n,(p/f'BEC_{n}.txt').read_text()
with concurrent.futures.ThreadPoolExecutor(max_workers=4)as ex:r=list(ex.map(run,range(1,len(d)+1)))
(p/'technical_BEC_OCR.json').write_text(json.dumps(r));print('\n'.join('PAGE '+str(n)+'\n'+t for n,t in r))
