import json
from pathlib import Path
from datetime import date
from docx import Document
from docx.shared import Inches,Pt
from docx.enum.section import WD_SECTION_START,WD_ORIENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
base=Path('/workspace/project-review'); data=json.loads((base/'final_schedule_data.json').read_text()); acts=data['recovery']; by={a['id']:a for a in acts}
d=Document();sec=d.sections[0];sec.top_margin=Inches(.65);sec.bottom_margin=Inches(.65)
normal=d.styles['Normal'];normal.font.name='Calibri';normal.font.size=Pt(10)
d.add_heading('OVERALL PROJECT SCHEDULE',0)
d.add_heading('SHIVANSH — ONGC Mumbai High\nBridge Refurbishment & Replacement Project',1)
d.add_paragraph('Revision 0 | Prepared 25 October 2026 | Provisional management planning')
d.add_paragraph('Planning scenario: commencement on 1 November 2026. Actual award / Notification of Award is unconfirmed. This document is a proposed overall schedule, subject to engineering, supplier, resource and ONGC review; it is not an approved contractual baseline.')
d.add_heading('1. Schedule basis',1)
for t in [
'Contract basis: 12 months from Notification of Award under reviewed GCC 6.3.1.1 and Technical SCC 2.0. Use 31 October 2027 as the conservative last workday for this scenario, subject to confirmation of the contractual date convention. The 360-day assumption in earlier planning files has not been adopted.',
'Scope: ten bridge packages — three replacements and seven refurbishments — with surveys, engineering, procurement, fabrication, transport, marine support, offshore installation, jacket strengthening, service transfer, demolition, inspection and handover.',
'Calendar: durations are elapsed calendar days, assuming sufficient shifts and staffing. Offshore tasks conservatively avoid 16 May–15 October inclusive. An uninterrupted offshore activity that does not fit before the monsoon is deferred to 16 October. This planning rule does not establish contractual extension entitlement.',
'Logic: finish-to-start relationships, zero lag. A successor begins on the day after a predecessor’s last workday. Zero-duration milestones record the following date boundary. F041 is the first-batch release within F040 production; the two durations are not additive.',
'Durations: planner estimates requiring work-quantity, productivity, supplier-lead-time and review-turnaround verification. No actual progress or committed resource availability is inferred.',
'Recovery resources: two fabrication lines, one dedicated strengthening cell, three offshore strengthening crews, one sequential heavy-lift spread and one cargo spread. Supporting accommodation and logistics must be sized for simultaneous work; SCC minimum accommodation for 60 special personnel does not establish adequate total capacity.',
'Weather reserve: ten offshore days are reserved after planned offshore work, followed by three demobilisation days. This allowance must be validated against metocean history and vessel operating limits.'
]:d.add_paragraph(t,style='List Bullet')
d.add_heading('2. Executive milestones',1)
t=d.add_table(rows=1,cols=3);t.style='Light Shading Accent 1'
for c,v in zip(t.rows[0].cells,['Milestone / deliverable','Recovery target','Two-crew comparison']):c.text=v
b={a['id']:a for a in data['two_crew']}
for i in ['M000','E030','P010','F030','I010','I020','D020','I030','M100','O100','M200']:
 r=t.add_row().cells
 for c,v in zip(r,[by[i]['name'],by[i]['finish'],b[i]['finish']]):c.text=v

d.add_paragraph('The recovery option targets offshore work completion at the 2 May 2027 milestone, demobilisation by 14 May 2027 and overall handover at the 15 June 2027 milestone. The two-crew case forecasts overall completion on 13 December 2027, after the assumed unextended deadline. Its seasonal delay and associated cost exposure require resolution; no automatic time extension is assumed.')
d.add_heading('3. Mandatory sequence and management gates',1)
for t in ['Complete surveys and approved design before fabrication or modification. Procurement release requires checked material requisitions; early purchase release is not fabrication approval.','Complete specified ICG, ICP and BHS jacket strengthening before the corresponding new bridge installations.','Install the new BHS–SLQ bridge and accept transferred services before isolating and demolishing the existing bridge.','Release loadout and offshore lifts only after verified weight / centre of gravity, rigging, crane capability, transport, sea-fastening, permits and operational readiness approvals.','Resolve the structural quantity difference (1,142 MT summary versus 1,141 MT bridge rows) and bridge-specific allocation of approximately 357 MT demolition before baseline approval.','Secure the additional third strengthening crew and early first-kit QA release. Validate survey access, engineering review capacity, accommodation, simultaneous operations and vessel commitments.']:
 d.add_paragraph(t,style='List Bullet')
sec=d.add_section(WD_SECTION_START.NEW_PAGE);sec.orientation=WD_ORIENT.LANDSCAPE;sec.page_width=Inches(11.7);sec.page_height=Inches(8.3);sec.left_margin=sec.right_margin=Inches(.45)
d.add_heading('4. Detailed overall schedule — proposed recovery option',1)
d.add_paragraph('Dates: YYYY-MM-DD. Duration: elapsed calendar days. Predecessors: finish-to-start, separated by semicolons. All activities are proposed / unverified.')
t=d.add_table(rows=1,cols=7);t.style='Light Shading Accent 1'
for c,v in zip(t.rows[0].cells,['ID','Package / activity','Days','Start','Finish','Predecessors','Resource']):c.text=v
trPr=t.rows[0]._tr.get_or_add_trPr();repeat=OxmlElement('w:tblHeader');trPr.append(repeat)
for a in acts:
 cells=t.add_row().cells
 for c,v in zip(cells,[a['id'],a['pkg']+' — '+a['name'],str(a['dur']),a['start'],a['finish'],'; '.join(a['preds']) or '—',a['resource'] or 'Milestone']):
  c.text=v
  for p in c.paragraphs:
   p.paragraph_format.space_after=Pt(3)
   for run in p.runs:run.font.size=Pt(8)
 for c in cells:
  for p in c.paragraphs:
   for run in p.runs:run.font.size=Pt(8)
d.add_heading('5. Baseline approval and control',1)
for text in ['Before approval: confirm NOA, controlling documents, work quantities, supplier lead times, crew capacities, vessel suitability and availability, ONGC interfaces, weather workability, cost and funding.','Resource checks performed: exclusive fabrication-line, strengthening-crew, cargo-spread and heavy-lift sequencing. Full labour-loaded and resource-levelled P6 analysis remains outstanding.','Control rhythm: daily workfront and marine coordination; weekly six-week constraint look-ahead and integrated schedule update; monthly completion and cost forecast. Record actual dates and accepted quantities rather than subjective progress.','Warranty: twelve months after actual completion and handover, subject to the contract. Rebase warranty and security-release dates to actual acceptance; they are separate from the EPC execution duration.']:
 d.add_paragraph(text,style='List Bullet')
d.add_paragraph('Source basis: uploaded GCC 2.1.1 / 6.3.1.1; Technical SCC 2.0 / 3.0; SOW Section 1 including 1.3; structural scope addendum; Level3, Stage28 and 360-day scheduling workbooks; Stage34/35 control and mobilisation plans. This proposed schedule does not supersede the executed contract or controlling addenda.')
for section in d.sections:
 footer=section.footer.paragraphs[0];footer.text='SHIVANSH / ONGC BRRP | Overall Schedule Rev 0 | Provisional | Page '
 field=OxmlElement('w:fldSimple');field.set(qn('w:instr'),'PAGE');footer._p.append(field)
output=base/'outputs'/'BRRP_Overall_Project_Schedule.docx';d.save(output)
from zipfile import ZipFile
with ZipFile(output) as z:assert z.testzip() is None
print(str(output));print('Size:',output.stat().st_size,'bytes; 44 schedule activities; DOCX integrity verified.')
