from pathlib import Path
import json,csv,zipfile,shutil
from docx import Document
from docx.shared import Pt
import openpyxl
from openpyxl.styles import Font,PatternFill,Alignment
O=Path('/workspace/project-review/outputs/BRRP_P6_Completion_Register')
# Each group has four distinct closure actions. Actual execution/approval status is not inferred.
G=[
('A','Award, dates and contractual obligations','Contracts Manager','GCC/SCC, executed contract and controlling amendments','Before baseline or binding commitment',[
('Confirm actual NOA, contractual commencement, milestones and completion convention.','Controlled award/date record and obligations register.'),
('Confirm warranty, securities, insurance, payment, LD and notification provisions.','Clause-by-clause commercial obligations with owners and due dates.'),
('Resolve document precedence and all applicable addenda; distinguish BRRP from other projects.','Approved contract/source hierarchy and amendment register.'),
('Confirm project funding authority and limits for purchases, subcontracts and charters.','Named delegation of authority and funding authorisation.')]),
('B','Bridge scope and quantity reconciliation','Engineering Manager / Quantity Surveyor','B774 scope, MTO p1, drawings and addenda','Before purchase quantities are frozen',[
('Freeze inclusions/exclusions for all10bridges, landings, services, temporary works and demolition.','Signed bridge scope/interface matrix.'),
('Reconcile structural row total1141MT with printed1142MT; landing-strengthening rows216 versus217MT.','Traceable corrected or confirmed MTO revision; original retained.'),
('Reconcile handrail1166m versus1168m and grating718m2 versus719m2.','Drawing-based dimensional quantity reconciliation.'),
('Check whether ICG/ICP/BHS addendum wraps and other work are already included in tender quantities.','Inclusion map preventing omissions and double counting.')]),
('C','Leadership, capability and resources','Project Director','P5 organisation, work packages and resource scenario','Before delivery commitments',[
('Assign named Project Director, discipline leads, deputies, checkers and package owners.','Organisation/RACI with availability and authority.'),
('Audit actual engineering, fabrication, QA/coating and offshore workforce capacity against competing work.','Capacity audit and committed resource calendar.'),
('Verify candidate vessels through ownership/charter status, particulars, certificates and project suitability.','Asset-verification and commitment records; website claims alone insufficient.'),
('Confirm third strengthening crew, installation team, support/inspection and peak POB/accommodation capacity.','Resourced work-front plan and crew/support mobilisation commitments.')]),
('D','Application of the supplied design criteria','Engineering Manager / Technical Authority','DC3.4 PartI Rev14, BRRP PartII and specifications','Before design freeze',[
('Issue project application register for25yearlife, codes/editions, materials and design conditions.','Approved design-basis application matrix; full criteria are already supplied.'),
('Apply250kg/m2 global walkway and500kg/m2 local grating loads to the correct elements.','Load definition, area maps and local/global distinction.'),
('Reconcile wind gust duration and elevation:15sec for56/84m studies and3sec for42m.','Documented wind derivation from applicable MH environmental parameters.'),
('Confirm contingencies, corrosion zone treatment, fatigue/VIV, piping20% future allowance and material requirements.','Calculation assumptions showing which allowances are included; no double counting.')]),
('E','Survey and existing condition','Survey Lead / Engineering','Source reports, tender drawings and approved survey scope','Before final dimensions or existing-capacity checks',[
('Verify bridge bearing centres, levels, skew, grid references and coordinate/datum system.','Accepted survey coordinates;56/84/42m remain study lengths until reconciled.'),
('Verify existing members, wall thickness, connections, supports and corrosion using specified inspections.','Condition/UT/NDT register with inaccessible areas identified.'),
('Verify live piping/cables, clearances, access, adjacent bridge and lifeboat/muster interfaces.','Coordinated existing-services and installation-clearance survey.'),
('Resolve tender drawing references and update existing-condition markups for all work fronts.','Controlled survey/as-built input pack with engineering acceptance.')]),
('F','Global structural design','Structural Analysis Lead','Accepted geometry, criteria, loads and source studies','Before member schedules or reaction release',[
('Obtain native source models or reconstruct models from verified coordinates, connectivity and section groups.','Documented reproducible model; native files are not the only possible route.'),
('Verify restraints, offsets, releases, stiffness, existing-platform movement and local/global axes.','Model-validation and support-philosophy record.'),
('Reconcile source gravity/live-load totals, load maps, allowances and all governing combinations.','Applied-load/reaction equilibrium checks and model revision log.'),
('Run and independently check strength, stability, deflection, vibration and applicable fatigue results.','Checked analysis report and complete signed result envelopes; source UC tables alone insufficient.')]),
('G','Member, joint and connection engineering','Structural Detail Lead / Checker','Checked global actions and material specifications','Before fabrication drawings',[
('Select and verify chords, braces, crossframes and walkway beams for applicable resistance and stability.','Member/group schedule with effective lengths, forces and code checks.'),
('Design tubular joints, gussets, stiffeners, end connections, bolts and welds for actual actions.','Local connection calculations and calculation-linked details.'),
('Verify fatigue-sensitive details, stress concentrations, inspectability and material toughness/weldability.','Detail-category and inspection provisions tied to design.'),
('Resolve camber from verified dead-load response and setting geometry; reconcile source1:200 instruction.','Dimensioned camber profile, datum, tolerances and checking record.')]),
('H','Bearings, landings, decks and jackets','Structural / Platform Engineer','Concurrent reactions, survey, movements and addendum','Before support procurement or modification',[
('Resolve hinged/pinned/sliding freedoms including contradictory SLQ wording; fixed intent is not moment-rigid.','Approved support philosophy and consistent analysis/drawing terminology.'),
('Calculate movement and axial thrust using specified150% requirements; design guides, uplift and friction.','Bearing travel/action envelopes, compatible rotations and bearing/guide calculations.'),
('Check bearing seats, landing extensions, existing beams, joints, decks and jackets under concurrent load cases.','Integrated existing-structure adequacy and strengthening report.'),
('Engineer ICG/ICP/BHS wrap strengthening, fit-up, local load transfer, welds and temporary conditions.','Checked strengthening drawings/MTO; ideal shell mass is not a capacity check.')]),
('I','Detailed drawings and AFC control','Engineering Manager / Document Controller','Checked calculations and agreed interface reviews','Before construction release',[
('Complete dimensioned GAs, framing, sections, connections, bearings, landings, access and service details.','Controlled drawing set with source survey/calculation references.'),
('Complete missing101/102 main arrangements and0005 references and103-series landing/demolition details.','Resolved drawing register; authorised reconstructed/new drawings where required.'),
('Complete discipline checking, independent/technical review, interface clearance and applicable ONGC review.','Closed review comments and actual approval evidence.'),
('Authorise issue purpose, signatures and controlled transmittals; withdraw superseded sheets.','Actual AFC release record when later authorised; no invented stamp/date/signature.')]),
('J','Final MTO and material control','Engineering / Procurement / QA','Released detailed drawings and measurement rules','Before final orders and throughout execution',[
('Prepare member/plate/component schedules with grade, size, cut length, finish and drawing revision.','Detailed bridge-specific purchasing MTO.'),
('Develop grating/handrail/sheeting/service/bolt/weld/coating quantities from actual details.','Discipline MTOs; coating area and paint volume not inferred from tonnage.'),
('Separate permanent steel, temporary works, rigging/sea-fastening and demolition; set justified allowances.','Reconciled take-off, inclusion map and approved wastage rules.'),
('Implement receipt, heat/mark traceability, preservation, issue, installed quantity, surplus and returns.','Material ledger and certificates traceable to installed work.')]),
('K','Procurement and subcontracts','Procurement / Commercial Manager','P5 procurement packages, approved requisitions and funding','Before supplier commitments',[
('Obtain compliant quotations and technical/commercial evaluations for all25packages.','Approved bid comparisons, technical acceptance and actual lead times.'),
('Verify vendor/yard/subcontract commitments and applicable client acceptance under amended provisions.','Capacity commitments and required subcontractor letters/approvals.'),
('Issue authorised orders/charters with specifications, inspection, documentation, schedule and change terms.','Executed purchase/subcontract documents; no order has been evidenced here.'),
('Expedite manufacture, inspections, dispatch and receipt; track bearings/lifesaving/services long leads.','Supplier progress evidence and updated procurement dates integrated into schedule.')]),
('L','Yard readiness and fabrication','Fabrication Manager','Released shop work packs, accepted material, qualified procedures','Before production and module release',[
('Confirm bays, cranes, loadout access, workforce, coating and NDT capacity with competing jobs.','Audited production-capacity and yard logistics plan.'),
('Prepare and execute cutting, fit-up, assembly, welding and dimensional-control work packs.','Traceable production records and approved procedure compliance.'),
('Complete specified NDT, repair/retest, dimensional acceptance and suitable trial fit.','Accepted inspection records and release-critical NCR closure.'),
('Complete approved coating, preservation, marking and fabrication dossier.','QA module/kit release with accepted actual quantities and dimensions.')]),
('M','Quality assurance and inspection','QA/QC Manager','Specifications, source IRT and project ITP','Before manufacture and each release',[
('Approve item-specific ITP/QAP, hold/witness points, inspection responsibilities and notices.','Authorised inspection matrix; no generic NDT percentage substituted.'),
('Verify personnel/procedures, welding qualifications, consumables, calibration and material certificates.','Competence, procedure and equipment evidence.'),
('Complete inspections, nonconformances, concessions and corrective actions with competent review.','Accepted inspection/NCR/concession records.'),
('Compile material, weld/NDT, dimensional, coating and release dossiers progressively.','Complete traceable quality dossier for each bridge/module.')]),
('N','Marine engineering and spread commitment','Marine Manager / Marine Engineer','Verified modules, route, field interfaces and vessel data','Before charter/loadout/sail-away',[
('Select technically suitable heavy-lift, cargo, support, accommodation, towage and supply resources.','Suitability assessments and availability/charter commitments.'),
('Complete loadout, grillage, sea-fastening, transport stability/motions and barge/deck checks.','Checked transport/loadout engineering and applicable marine assurance.'),
('Define field positioning, moves, transfers, POB, logistics, operational limits and weather decisions.','Integrated marine campaign and emergency arrangements.'),
('Secure required certification, port/statutory/field permissions and assurance releases as applicable.','Valid certificates/permissions and signed readiness records.')]),
('O','Lifting and temporary conditions','Lifting / Temporary-Works Engineer','Verified weight/COG, crane/radius and sequence','Before any lift or temporary load-path change',[
('Verify module boundaries, actual weight/COG and crane capability at operating radius/configuration.','Weight control/weighing and lift-feasibility report.'),
('Analyse rigging geometry, skew/dynamics, adverse COG/weight and lift-point/local effects.','Checked lift model, lifting-point/spreader/rigging design.'),
('Engineer erection stability, temporary supports, guides, scaffolding/handling interfaces and removal stages.','Checked temporary-work drawings and authorised sequences.'),
('Approve operation-specific lift/landing method, weather window, communications and contingency.','Technical/operational go/no-go release; management schedule is not lift approval.')]),
('P','Shore-base and offshore logistics','Logistics / Marine Manager','Cargo schedule, release records and support capacity','Before sail-away and continuously',[
('Establish storage, preservation, handling, manifests and required customs/transport arrangements.','Shore-base procedure and traceable cargo register.'),
('Deliver complete bridge kits, tools, consumables, inspection equipment and critical spares.','Work-pack completeness and offshore receipt checks.'),
('Coordinate personnel transfers, rotations, accommodation, supply and field moves with actual demand.','Integrated POB/transfer/supply calendar.'),
('Execute backload, waste segregation/disposal, equipment returns and surplus reconciliation.','Accepted backload/disposal/return records.')]),
('Q','Replacement-bridge offshore installation','Offshore Construction Manager','Released work pack, landing/jacket readiness and marine approval','Before and during installation',[
('Close access, permits, isolations, SIMOPS, crew/tool/material and emergency prerequisites.','Signed work-front readiness and applicable ONGC authorisation.'),
('Execute approved loadout/transport/lift/landing and maintain engineered temporary stability.','Operation records with inspections and deviations controlled.'),
('Complete permanent connections, bearings/guides, access fittings, services and coating restoration.','Accepted permanent-work inspection/test records.'),
('Verify final dimensions/settings, accepted quantities and package punch closure.','Client/QA acceptance evidence and as-built markups.')]),
('R','Seven-bridge strengthening execution','Offshore Construction / Technical Authority','Accepted condition survey and bridge-specific repair design','Before and during strengthening',[
('Validate repair locations/thicknesses and live-service restrictions before each work front.','Accepted scope markups and actual condition records.'),
('Provide approved access, protection, temporary support and complete strengthening kits.','Work-pack readiness and temporary-work release.'),
('Execute authorised member/joint/landing modifications, fittings replacement and restoration.','Production and inspection records without unapproved load-path changes.'),
('Accept weld/NDT/dimensions/coating, remove temporary arrangements and close package punch.','Bridge-specific acceptance and measured quantities for all seven packages.')]),
('S','BHS service transfer and demolition','Commissioning / Offshore / ONGC Operations','New-bridge engineering plus live-service and demolition plans','Before service transfer or old-bridge removal',[
('Freeze line/cable/loop lists, support loads, isolations, tie-ins and operational contingency.','Approved multidisciplinary transfer/commissioning package.'),
('Install and test new piping/electrical/instrument interfaces and obtain operational acceptance.','Signed test/reinstatement and service-transfer records.'),
('Confirm new bridge/services accepted, redundant services isolated and demolition stability verified.','Formal demolition release and checked removal method.'),
('Execute authorised dismantling, backload/disposal and remaining-platform restoration.','Demolition completion, waste records and post-work acceptance.')]),
('T','Testing and commissioning','Completions Manager / Discipline Leads','Accepted construction and approved test boundaries/procedures','Before operational acceptance',[
('Prepare system/bridge test packs, prerequisites, witnessing and discipline-specific acceptance values.','Approved procedures; no invented test pressure, torque or electrical threshold.'),
('Complete structural/bearing/access checks and applicable electrical/earthing/lighting tests.','Accepted structural and electrical records.'),
('Complete applicable piping pressure/leak/cleaning/reinstatement and instrument loop/function tests.','Line/loop-specific accepted test and reinstatement records.'),
('Complete operational checks, safety/lifesaving interfaces, punch classification and acceptance.','ONGC operational acceptance and authorised residual-item register.')]),
('U','Schedule, critical path and recovery','Project Controls / Project Director','P5 scenario, actual NOA, quantities, commitments and workability','Before baseline approval; update weekly',[
('Rebase the205activity scenario to actual contract dates and verified supplier/approval durations.','Approved schedule basis, logic and milestones.'),
('Load measured quantities, demonstrated productivity, crews, QA/coating, vessels and POB; level conflicts.','Capacity-backed resource-loaded schedule and histograms.'),
('Validate offshore workability, safe seasonal stopping/revalidation, reserve and demobilisation strategy.','Defensible seasonal margin and contingency/recovery plan.'),
('Update actuals, remaining durations, critical/near-critical paths, look-ahead constraints and forecasts.','Weekly accepted-progress update and authorised recovery decisions.')]),
('V','Cost, cash, guarantees and collections','Finance / Cost / Commercial Manager','Actual quotations, contract payment terms and financing','Before financial sanction; update monthly',[
('Price the16cost categories with committed/actual/remaining costs and controlled shared allocations.','Approved cost baseline and forecast at completion; blanks are not zero.'),
('Model supplier/charter payments, standby/remob, taxes, bank margins, financing and risk contingency.','Full-duration cost/cash model and funding requirement.'),
('Map billing to amended milestones, accepted measurements/e-MB, certification and collection dates.','Auditable invoice/certification/receipts register.'),
('Confirm funding and guarantee/insurance capacity; monitor cash runway and overdue collections.','Financial authorisations, valid securities and updated cash forecast.')]),
('W','HSE, permits and emergency readiness','HSE / Construction / Marine / ONGC','Approved work methods, actual spreads and operational interfaces','Before mobilisation and every operation',[
('Complete hazard reviews and task controls for lifting, hot work, height, dropped objects and live services.','Approved risk assessments and resolved critical actions.'),
('Verify training, medical/competence, fatigue/rotation and relevant certification.','Personnel readiness and supervision records.'),
('Agree PTW/SIMOPS/isolations, emergency communications, rescue, medical evacuation and weather actions.','Operational permits, drills and emergency-readiness evidence.'),
('Implement environmental/waste/spill controls, leadership checks, incident reporting and action closure.','HSE performance and corrective-action records throughout execution.')]),
('X','Changes, notices and contractual entitlement','Contracts / Commercial Manager','Controlling contract, instructions and contemporaneous records','From first change/event through final account',[
('Record changed conditions, instructions, quantity changes and potential delay events immediately.','Dated change/event register and evidence.'),
('Issue required notices within actual contractual time limits; obtain authorised instructions.','Compliant correspondence and instruction trail.'),
('Substantiate time/cost effects, preserve delay/resource records and prevent duplicate recovery.','Supported change submissions and entitlement assessment.'),
('Agree approved changes and flow effects into drawings, MTO, schedule, subcontracts and billing.','Controlled contract/baseline revisions and approved settlement records.')]),
('Y','Handover and operational documentation','Completions / QA / Contracts / ONGC','Accepted work, test evidence and agreed turnover requirements','Before contractual completion',[
('Define bridge/system completion boundaries and agreed dossier/punch requirements.','Approved turnover index and acceptance checklist.'),
('Complete as-builts, calculations/models as required, material/weld/NDT/coating/test and marine records.','Complete accepted engineering/manufacturing/construction dossiers.'),
('Close acceptance-critical punch and document agreed residual items, measurements and e-MB evidence.','Signed punch/quantity/certification records.'),
('Obtain actual completion/handover acceptance and record commencement of warranty obligations.','Client certificates and actual package/overall acceptance dates.')]),
('Z','Warranty, final account and closeout','Project Director / Contracts / Warranty Coordinator','Actual handover dates and contractual release conditions','After acceptance through warranty/final closure',[
('Operate defect notification, response, rectification and verification during contractual warranty.','Closed defect register and accepted corrective work.'),
('Maintain required securities/insurance and obtain releases only when contractual conditions are met.','Guarantee/insurance closeout and release records.'),
('Settle final quantities, approved changes, subcontract accounts, surplus and outstanding payments.','Agreed final account and financial reconciliation.'),
('Archive controlled records and capture productivity, cost, marine downtime and lessons learned.','Final project closure report and reusable verified data.')])]
assert len(G)==26 and all(len(x[5])==4 for x in G)
rows=[]
for letter,title,owner,basis,timing,items in G:
 for i,(action,evidence) in enumerate(items,1):rows.append([f'{letter}{i:02}',letter+' - '+title,action,owner,basis,timing,evidence,'OPEN / COMPLETION NOT EVIDENCED IN THIS SESSION','Assign named owner and target date'])
# Traceability of preparation already done, so supplied inputs are not requested again.
completed=[['Supplied design criteria','Received and indexed','DC3.4 PartI Rev14;BRRP-specific PartII;methodology'],['Source drawings and IRT','Received and indexed;partialbridge drawing coverage','MTO Drawings IRT;reference models and landing sheets'],['Scope/work-package preparation','P5 drafted','100work packages across10bridges'],['Engineering review','P4 preliminary calculations and source sheets','Not independently solved global/member/connection/bearing design'],['AFC preparation','Register/forms prepared','No actualAFC approval'],['Quantity audit','P5 completed arithmetic','1141/1142MT;1166/1168m;718/719m2'],['Procurement preparation','25packages registered','Quotes/orders/actual delivery not evidenced'],['Schedule/CPM','205activity conditional scenario computed','ActualNOA/resource/productivity/workabilityvalidation pending'],['Execution/commissioning/handover plans','P5 drafted','Physicalcompletion/clientacceptance not evidenced'],['Cost/cash controls','Unpricedinputregister prepared','No approvedcost/fundingforecast']]
w=openpyxl.Workbook();s=w.active;s.title='A-Z Closure Register';head=['ID','Workstream','Outstanding action','Responsible role','Required basis','Need-by gate','Closure evidence','Current evidence status','Next control'];s.append(head)
for r in rows:s.append(r)
s.freeze_panes='C2';s.auto_filter.ref=s.dimensions
for i,width in enumerate([12,46,100,45,75,60,110,55,50],1):s.column_dimensions[openpyxl.utils.get_column_letter(i)].width=width
s2=w.create_sheet('Already Supplied or Prepared');s2.append(['Subject','Status','Evidence/limit'])
for r in completed:s2.append(r)
for sh in w:
 sh.freeze_panes='A2';sh.auto_filter.ref=sh.dimensions
 for c in sh[1]:c.font=Font(bold=True,color='FFFFFF');c.fill=PatternFill('solid',fgColor='17365D')
 for row in sh:
  for c in row:c.alignment=Alignment(wrap_text=True,vertical='top')
 if sh==s2:
  for col in ['A','B','C']:sh.column_dimensions[col].width=70
w.save(O/'BRRP_P6_A_to_Z_Completion_Register.xlsx')
with open(O/'BRRP_P6_A_to_Z_Completion_Register.csv','w',newline='') as f:
 z=csv.writer(f);z.writerow(head);z.writerows(rows)
d=Document();d.styles['Normal'].font.name='Calibri';d.styles['Normal'].font.size=Pt(10)
d.add_heading('BRRP — A-to-Z Outstanding Completion Requirements',0)
d.add_paragraph('P6 | 25 October2026 |104closure actions across26workstreams')
d.add_paragraph('This register distinguishes supplied/prepared documents from engineering checks, approvals, commercial commitments and physical work whose completion has not been evidenced in this session. OPEN does not assert that the organisation has never performed the work; it means closure evidence has not been established here. Construction-release approval remains a separate future decision.')
d.add_heading('What is already available',1)
for r in completed:d.add_paragraph(r[0]+': '+r[1]+'. '+r[2]+'.')
d.add_heading('Outstanding workstreams and closure evidence',1)
for letter,title,owner,basis,timing,items in G:
 d.add_heading(letter+' — '+title,1);d.add_paragraph('Responsible role: '+owner+'. Required basis: '+basis+'. Need-by: '+timing+'.')
 for i,(action,evidence) in enumerate(items,1):
  d.add_paragraph(f'{letter}{i:02}. {action}')
  d.add_paragraph('Close only with: '+evidence)
d.add_heading('Priority closure sequence',1)
for t in ['1 Confirm contract/NOA, named authorities, scope and the three quantity discrepancies.','2 Verify survey geometry and load maps; reconstruct/check analysis and existing support capacity.','3 Complete detailed design and reviews; decide construction release later with genuine evidence.','4 Validate prices, supplier lead times, capacity and marine/crew commitments; approve a robust baseline and funding plan.','5 Execute controlled procurement, fabrication, marine/offshore work and inspections under approved methods.','6 Complete service transfer/testing, acceptance, dossiers, warranty and financial closeout.']:
 d.add_paragraph(t)
d.add_paragraph('No blanket re-upload of the design criteria/specifications is required. Requests should identify only a specific unlocated drawing, model/survey datum, commercial commitment or acceptance record after checking the existing set. Native models may be supplied or rebuilt from verified complete inputs; source summary pictures are not an independently checked model.')
d.save(O/'BRRP_P6_A_to_Z_Completion_Requirements.docx')
(O/'P6_completion_data.json').write_text(json.dumps({'groups':G,'actions':rows,'already_available':completed},indent=2))
for p in O.iterdir():
 if p.suffix in ['.xlsx','.docx']:
  with zipfile.ZipFile(p) as z:assert z.testzip() is None
assert len(rows)==104 and len({r[0] for r in rows})==104
Z=O.parent/'BRRP_P6_Completion_Register.zip'
with zipfile.ZipFile(Z,'w',zipfile.ZIP_DEFLATED) as z:
 for p in O.iterdir():z.write(p,arcname=O.name+'/'+p.name)
DL=Path('/workspace/downloads')
for p in [Z,O/'BRRP_P6_A_to_Z_Completion_Requirements.docx',O/'BRRP_P6_A_to_Z_Completion_Register.xlsx',O/'BRRP_P6_A_to_Z_Completion_Register.csv']:
 q=DL/p.name;shutil.copyfile(p,q);q.chmod(0o644)
print('Prepared104specificclosureactions across26workstreams;Word,Excel,CSV andZIPsavedin/workspace/downloads. Noapproval/physicalcompletionassumed.')
