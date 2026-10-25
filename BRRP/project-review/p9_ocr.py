import json,pathlib,fitz,subprocess,concurrent.futures,os
out=pathlib.Path('/workspace/project-review/outputs/BRRP_P9_Input_Review');idx=json.loads((out/'page_index.json').read_text());r=idx[-1];d=fitz.open(r['path']);pages=[1,124,125,160,161,162,163,164,165,166,167,168,246,260]
for n in pages:d[n-1].get_pixmap(matrix=fitz.Matrix(2,2)).save(str(out/f'ocr_{n}.png'))
def task(n):
 subprocess.run(['tesseract',str(out/f'ocr_{n}.png'),str(out/f'ocr_{n}'),'--psm','3'],env={**os.environ,'OMP_THREAD_LIMIT':'1'},capture_output=True,check=True)
 return n,(out/f'ocr_{n}.txt').read_text()
with concurrent.futures.ThreadPoolExecutor(max_workers=4) as ex:
 for n,t in ex.map(task,pages):print('PAGE',n,t[:4200])
