from pathlib import Path
import ast,re,json,zipfile,shutil,hashlib
from docx import Document
from docx.shared import Pt
import openpyxl,fitz
R=Path('/workspace/project-review');O=R/'outputs/BRRP_P8_Input_Audit'
m=ast.parse((R/'beautify_p5.py').read_text());nodes=[]
for n in m.body:
 if isinstance(n,ast.Assign) and any(isinstance(t,ast.Name) and t.id in ['pairs','mp'] for t in n.targets):nodes.append(n)
 if isinstance(n,ast.FunctionDef) and n.name=='clean':nodes.append(n)
ns={'re':re};exec(compile(ast.Module(body=nodes,type_ignores=[]),'clean','exec'),ns);baseclean=ns['clean']
phrases={'allthree':'all three','allten':'all ten','glob al':'global','globaldeckinplace':'global deck in-place','localdeckinplace':'local deck in-place','sourceavailability':'source availability','source-availability':'source-availability','pre-sail':'pre-sail','vesselaudit':'vessel audit','vesselaudits':'vessel audits','firstreviews':'first reviews','subsequentreviews':'subsequent reviews','actualrevisedattachments':'actual revised attachments','source drawings':'source drawings','beforefreezingcommercial/vendorbasis':'before freezing the commercial/vendor basis','BidderPM':'bidder project management','retainedbidderPM':'retained bidder project management','qualifyingactivity':'qualifying activity','wholeworksback-to-backsubcontract':'whole-works back-to-back subcontract','source-GA':'source GA','register0/titleA':'register 0 / title A','revisionconflict':'revision conflict','notAFC':'not AFC','fulltechnicalcheckofeverypage':'full technical check of every page','fullindependentdesign':'full independent design','leadtime':'lead time','specification/MTO/vendor/test':'specification / MTO / vendor / test','technicalquery':'technical query','criticalpre-sail':'critical pre-sail','subcontractingconstraints':'subcontracting constraints','actualaudit':'actual audit','notpositively':'not positively','notfound':'not found','notidentified':'not identified','SVLis':'SVL is','scope applicability':'scope applicability','metadataalone':'metadata alone','filenamesalone':'filenames alone','wholeset':'whole set','newuniquecontent':'new unique content','Exactearliercopy':'Exact earlier copy','Samebatchduplicate':'Same-batch duplicate','Samebatchduplicateof':'Same-batch duplicate of','Earlierexactmatches':'Earlier exact matches','printedtotal':'printed total','retainedbidder':'retained bidder','globalBridge':'global bridge','allplatformchecks':'all platform checks','NOQlabel':'NOQ label','DrawingID':'Drawing ID','FirstPDFpage':'First PDF page','LastPDFpage':'Last PDF page','Reviewstatus':'Review status','PDFpages':'PDF pages'}
def clean(t):
 if not isinstance(t,str):return t
 t=baseclean(t)
 for a,b in sorted(phrases.items(),key=lambda x:-len(x[0])):t=t.replace(a,b)
 return t
repl={
'All16files':'All16files were downloaded, hashed, compared with earlier local copies and registered. Within this batch there are14unique contents and754PDF pages; six unique contents were not present among earlier downloaded copies. Text was extracted and indexed;20low-text pages were OCR-screened; relevant clauses and key source drawings were examined. This is a focused input/impact review, not an independent technical validation of every page or a completed audit of all historical uploads.',
'Theactualsheetsremain':'The source sheets are issued for tender and require contractor verification. They provide concrete geometry, member and connection references for design development, but are not signed AFC drawings. Preserve original source IDs and proposed new IDs separately. Reconcile drawing-index revisions with actual title blocks and older references. Omission from an index alone does not delete contractual work.',
'Developandobtainapproval':'Finalise and obtain approval for the project-specific DCI/MCI. The standard lists are not exhaustive. Use the required project/facility/discipline/document-type numbering and map the temporary Shivansh IDs.',
'Keepdocumentcategory':'Keep document category A/I/R (approval/information/review), review status Cat A/B/C and predecessor state S/R/A as three different control fields.',
'Generallyallow':'In general allow10working days for review. Apply longer first-review periods where explicitly listed; subsequent revisions have10working-day review periods. Use an agreed working calendar rather than assuming calendar days.',
'Engineeringcredit:':'The source review-status rules state Cat C: no engineering progress; Cat B:60percent engineering progress per the stated submission rule; Cat A:100percent. Apply approved document weights. These are not overall project percentages; proposed internal milestone weights cannot replace applicable source rules.',
'Purchase-specificationfirst':'The source allocates5percent of an item procurement weight on first-stage purchase-specification approval and5percent on second-stage approval. Do not treat that as10percent of the whole project or duplicate payment milestones.',
'Submitrevisions':'Submit revisions with a comment-resolution sheet and properly clouded/indexed changes, endorsed through the contractor design agency. Out-of-sequence submissions may be returned without review.',
'Apply30day':'Apply30-day advance-submission notes and MWS approval requirements to the identified, applicable rows. Do not extend generic platform requirements to unrelated facilities without checking.',
'Do not reaffirm the15June':'Do not reaffirm the15June2027scenario as validated after receipt of these inputs. Add the required bridge/platform analyses, review and resubmission cycles, purchase-specification stages, vessel audits and applicable assurance milestones; then recalculate logic, calendars and resources. The earlier scenario remains preliminary. Apply the general vendor/agency amendment together with technical, qualification and exclusion controls.',
'Thecoverreferences':'Match the revised technical SCC and A3 announced on the23April cover against the actual revised attachments. Metadata and filenames alone cannot establish precedence.',
'ConfirmSVL':'Confirm SVL Revision12.3 or the applicable current basis. The separately uploaded12.1list must not be described as the latest revision.',
'ReconcileGA':'Reconcile the ICW-ICG GA101 drawing register Rev0, title-block Rev A and revision-history entry0.',
'Mapold':'Map older101/1020005 references to the revised register and actual sheet content; do not assume omitted work is deleted or request obsolete drawing IDs without checking.',
'ConfirmNOQ':'Confirm the NOQ label in the deliverable annexure against project nomenclature NQO.',
'Reconcileexperience':'Reconcile experience forms and qualification thresholds with the controlling BEC and amendments. Do not determine bidder eligibility from a proforma alone.',
'Engineeringcalculation':'Engineering calculations, current survey verification, actual supplier/crew/vessel commitments, AFC approval and physical execution remain separate open activities.',
'Confirmed:every':'Confirmed: every file in this latest16-file batch is registered and version-checked, and relevant requirements/source changes have been identified with page references. Not confirmed: full incorporation of all historical uploads, applicability of every generic standard, resolution of all clauses, or project/design/AFC completion. OCR transcripts are review aids; the original PDFs control.'}
p=O/'BRRP_P8_Latest_Upload_Analysis_and_Required_Updates.docx';d=Document(p)
for para in d.paragraphs:
 t=para.text
 for prefix,new in repl.items():
  if t.startswith(prefix):t=new;break
 para.text=clean(t)
for table in d.tables:
 for row in table.rows:
  for cell in row.cells:
   for para in cell.paragraphs:
    para.text=clean(para.text)
    for run in para.runs:run.font.size=Pt(8)
d.save(p)
p=O/'BRRP_P8_Input_Inventory_and_Impact_Register.xlsx';w=openpyxl.load_workbook(p)
for s in w:
 for row in s:
  for c in row:
   # Do not change exact hashes, paths, file names or drawing IDs.
   if isinstance(c.value,str) and not (len(c.value)==64 and re.fullmatch('[a-f0-9]+',c.value)) and not c.value.startswith('/workspace/') and not c.value.endswith('.pdf'):c.value=clean(c.value)
w.save(p)
# Keep OCR evidence with the report, without changing originals.
e=O/'OCR_Transcripts';e.mkdir(exist_ok=True);ocr=json.loads((O/'P8_OCR_inventory.json').read_text())
for x in ocr:shutil.copyfile(x['ocr_path'],e/Path(x['ocr_path']).name)
# Validate scope and duplicate relationships.
rows=json.loads((O/'P8_file_inventory.json').read_text());assert len(rows)==16 and len({r['sha256'] for r in rows})==14
assert sum(r['pages'] for r in rows if not r['duplicate_in_batch'])==754
assert sum(not r['prior_matches'] for r in rows if not r['duplicate_in_batch'])==6
assert len(ocr)==20 and all(x['exit']==0 for x in ocr)
assert w['File Incorporation'].max_row==17 and w['Findings and Changes'].max_row==17 and len(w.sheetnames)==6
assert len(fitz.open(O/'P8_Source_Structural_Drawings_and_Deliverable_Matrix.pdf'))==22
for p in O.iterdir():
 if p.suffix in ['.docx','.xlsx']:
  with zipfile.ZipFile(p) as z:assert z.testzip() is None
(O/'README_P8.txt').write_text('Latest16-file batch audit only.16uploads,14unique contents,754unique pages within this batch,6new unique contents compared with earlier local copies.20low-textpages OCR-screened. Key relevant clauses and source drawings reviewed; not a complete independenttechnical validation of everypage or allhistoricaluploads. See Wordreport,6-sheetregister and22-page source structural excerpt. ActualDCIreview/progressrules and previouslymissingGAs now located; earlier scenario assumptions require revision. NoAFC/design/qualification/physicalcompletion is implied. OCRaids must be checked against originals.\n')
files=[]
for p in sorted(O.rglob('*')):
 if p.is_file():files.append({'file':str(p.relative_to(O)),'bytes':p.stat().st_size,'sha256':hashlib.sha256(p.read_bytes()).hexdigest()})
(O/'P8_delivery_manifest.json').write_text(json.dumps(files,indent=2))
Z=O.parent/'BRRP_P8_Latest_Input_Analysis.zip'
with zipfile.ZipFile(Z,'w',zipfile.ZIP_DEFLATED) as z:
 for p in sorted(O.rglob('*')):
  if p.is_file():z.write(p,arcname=O.name+'/'+str(p.relative_to(O)))
with zipfile.ZipFile(Z) as z:assert z.testzip() is None
assert Z.stat().st_size<40*1024*1024
for p in [Z,O/'BRRP_P8_Latest_Upload_Analysis_and_Required_Updates.docx',O/'BRRP_P8_Input_Inventory_and_Impact_Register.xlsx',O/'P8_Source_Structural_Drawings_and_Deliverable_Matrix.pdf']:
 q=Path('/workspace/downloads')/p.name;shutil.copyfile(p,q);q.chmod(0o644)
print('P8validated:16uploads/14unique/754pages;6newunique;20OCRpages;16findings;22pagesourceextract. PackageMB',round(Z.stat().st_size/1e6,2),'savedforlaterdownload.')
