from pathlib import Path
import re,json
from docx import Document
import openpyxl
O=Path('/workspace/project-review/outputs/BRRP_Integrated_Project_P5')
pairs='''projectmanagement|project management
project-specificclauses|project-specific clauses
project-specificprovisions|project-specific provisions
project-specific|project-specific
projectlead|Project lead
workpackages|work packages
workpackage|work package
workpacks|work packs
workpack|work pack
workfrontdecisions|work-front decisions
workfronts|work fronts
workfront|work front
six-weeklookahead|six-week look-ahead
lookahead|look-ahead
sourcerecords|source records
sourcestudies|source studies
sourcemodel|source model
source-baseddrawings|source-based drawings
sourceMTO|source MTO
calculationregister|calculation register
capacityhold|capacity hold
constructionrelease|construction release
source-input|source input
actualprogress|actual progress
global3D|global 3D
memberresistance|member resistance
handrailm|Handrail m
gratingm2|Grating m²
GIsheetm2|GI sheeting m²
StructuralrowMT|Structural row MT
DemolitionMT|Demolition MT
firstkitrelease|first kit release
fullquantity-loadedcrew|fully quantity-loaded crew
vendorcommitments|vendor commitments
measuredproductivity|measured productivity
resourcecommitments|resource commitments
resourcecapacity|resource capacity
namedexclusive|named exclusive
internalfinish|internal finish
handovermilestone|handover milestone
milestoneboundary|milestone boundary
phaseboundary|phase boundary
physicalchildtasks|physical child tasks
drivingoffshoretasks|driving offshore tasks
earliestfinish|earliest finish
unextendeddeadline|unextended deadline
seasonbarrier|seasonal barrier
seasonalclosure|seasonal closure
firstseason|first season
secondseason|second season
second-seasonremobilisation|second-season remobilisation
second-seasonremob|second-season remobilisation
secondseasonremob|second-season remobilisation
safe-suspension|safe suspension
safesuspension|safe suspension
loadcase|load case
loadcases|load cases
loadmap|load map
crewc|crew C
actualNOA|actual NOA
NOAandcontractualcalendar|NOA and contractual calendar
contractualcalendar|contractual calendar
fullcalendars|full calendars
calendarstress|calendar stress
delayscenario|delay scenario
sourcebasis|source basis
quotationsandallocation|required quotations and allocation
costatcompletion|cost at completion
costallocation|cost allocation
remainingforecast|remaining forecast
fundingtriggers|funding triggers
budget/margin/fundingforecast|budget, margin and funding forecast
cost/cash/charterstandby|cost, cash and charter standby
charterstandby|charter standby
charterwindows|charter windows
charterquotations|charter quotations
supplierleadtime|supplier lead time
vendorleadtime|vendor lead time
leadtime|lead time
lead-timeconstraints|lead-time constraints
actualvendor|actual vendor
minimumairgap|minimum air gap
materialrelease|material release
purchaserelease|purchase release
yardrelease|yard release
sail-awayrelease|sail-away release
engineeringrelease|engineering release
operationalrelease|operational release
releasegates|release gates
signedrelease|signed release
NOTAFC|NOT AFC
finaldesign|final design
finaldimensions|final dimensions
procurementquantities|procurement quantities
acceptedquantities|accepted quantities
actualhours|actual hours
labour-hours|labour-hours
actualwork|actual work
engineeringreview|engineering review
approvalmatrix|approval matrix
technicalauthority|technical authority
ProjectDirector|Project Director
TechnicalAuthority|Technical Authority
EngineeringManager|Engineering Manager
ProcurementManager|Procurement Manager
MarineManager|Marine Manager
FinanceManager|Finance Manager
CompletionsManager|Completions Manager
surveydata|survey data
pre-engineeringSurvey|pre-engineering survey
survey-dependent|survey-dependent
localdesign|local design
localchecks|local checks
localgrating|local grating
globalwalkway|global walkway
windduration|wind duration
temporaryconditions|temporary conditions
temporaryworks|temporary works
member/joint|member/joint
existingplatform|existing platform
existingstructure|existing structure
existingcorrosion|existing corrosion
landingstrengthening|landing strengthening
landingmodifications|landing modifications
landing/deck/jacket|landing/deck/jacket
newbridge|new bridge
newbridges|new bridges
oldbridge|old bridge
newfacility|new facility
service-transfer|service transfer
service-transfer|service transfer
servicesacceptance|services acceptance
permanentconnections|permanent connections
concurrentforces|concurrent forces
concurrentreactions|concurrent reactions
weighingoperations|weighing operations
weightcontrol|weight control
mass/COG|mass/COG
cranecurves|crane curves
actualradii|actual radii
liftweights|lift weights
liftrelease|lift release
liftstudy|lift study
lift-able|liftable
approvedmethods|approved methods
approvedprocedures|approved procedures
approveddrawings|approved drawings
approveddetails|approved details
approvedrequisition|approved requisition
approvedtest|approved test
approvedvendor|approved vendor
approvedITP|approved ITP
specifiedtests|specified tests
actualcondition|actual condition
existingcondition|existing condition
NDTcoverage|NDT coverage
qualifiedcoverage|qualified coverage
paintingprocedure|painting procedure
weldingparameters|welding parameters
coatingarea|coating area
qualityrecords|quality records
testrecords|test records
quality/testrecords|quality/test records
wasterecords|waste records
as-builts|as-builts
punchstatus|punch status
punchclassification|punch classification
punchclosure|punch closure
finalaccount|final account
guaranteereleases|guarantee releases
warrantycommencement|warranty commencement
fleetavailability|fleet availability
fleetcapabilities|fleet capabilities
crossprojectconflicts|cross-project conflicts
crossproject|cross-project
allcapacity|all capacity
physicalcompletion|physical completion
UNPRICED|UNPRICED
'''
mp={a:b for line in pairs.splitlines() if line for a,b in [line.split('|',1)]}
# Longer keys first; source file names and identifiers preserved where exact fields.
def clean(t):
 if not isinstance(t,str):return t
 if t.startswith('=') or re.fullmatch(r'[A-Z0-9_.:/;-]+',t):return t
 for k,v in sorted(mp.items(),key=lambda kv:-len(kv[0])):t=t.replace(k,v)
 t=re.sub(r'\b(\d+)(years?|months?|days?|minutes?|seconds?|nonzero|milestones?|activity|activities|bridges?|tasks?|crew|crews|sheet|sheets|personnel|special|calendar|daybulk|dayproductionstream)',r'\1 \2',t)
 t=re.sub(r'\b(\d+)(November|October|June|May|January|February|March|April|December|Nov|MT|MPa|GPa|kN|kg)',r'\1 \2',t)
 t=re.sub(r'(\d)(m2|mm|m/s2|kg/m2|kN/m2|MPa)\b',r'\1 \2',t)
 # Readable numeric boundaries while preserving document/activity IDs and engineering units.
 protected=[]
 def keep(m):
  protected.append(m.group(0));return '§'+str(len(protected)-1)+'§'
 t=re.sub(r'\b(?:[A-Z]{1,8}\d+[A-Z]*(?:[.-]\d+)*|kg/m2|kN/m2|kg/m3|mm2|mm3|mm4|m2|m3|m/s2|3D|2D|L/400|L/500)\b',keep,t)
 t=re.sub(r'(?<=[A-Za-z])(?=\d)|(?<=\d)(?=[A-Za-z])',' ',t)
 for i,v in enumerate(protected):t=t.replace('§'+str(i)+'§',v)
 for a,b in [('PartIII','Part III'),('PartII','Part II'),('PartI','Part I'),('nonzerotasks','nonzero tasks'),('nonzeroactivities','nonzero activities'),('seasonalcapacity','seasonal capacity'),('safestoppoints','safe stopping points'),('actualaward','actual award'),('demobstrategy','demobilisation strategy'),('handoverheadroom','handover headroom'),('firstbatch','first batch'),('zerofloat','zero float')]:t=t.replace(a,b)
 return t
p=O/'BRRP_P5_Integrated_Project_Execution_Plan.docx';doc=Document(p)
for pp in doc.paragraphs:
 new=clean(pp.text)
 if new!=pp.text:pp.text=new
for table in doc.tables:
 for row in table.rows:
  for cell in row.cells:
   for pp in cell.paragraphs:
    new=clean(pp.text)
    if new!=pp.text:pp.text=new
from docx.shared import Pt
for table in doc.tables:
 for row in table.rows:
  for cell in row.cells:
   for pp in cell.paragraphs:
    for run in pp.runs:run.font.size=Pt(8)
doc.save(p)
p=O/'BRRP_P5_Integrated_Project_Control.xlsx';w=openpyxl.load_workbook(p)
for s in w:
 for row in s:
  for c in row:
   if c.data_type!='f':c.value=clean(c.value)
w.save(p)
print('Improved report and workbook prose spacing; preserved IDs, formulas and quantities.')
