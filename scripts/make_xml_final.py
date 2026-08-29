# Generates the twelve FY2025 Cost.Submission XMLs (Gen2 schema) from the FINAL workbook,
# validates each against the official (repaired) XSD, and zips them.
import openpyxl, os, datetime
from collections import defaultdict
from xml.sax.saxutils import escape
from lxml import etree

SRC='/home/user/Clinical-Costing/output/Clinical_Costing_2025_MF4159_FINAL.xlsx'
OUT='/home/user/Clinical-Costing/output/xml_final'
XSD='/home/user/Clinical-Costing/docs/shafafiya/CostSubmission_repaired.xsd'
os.makedirs(OUT,exist_ok=True)
wb=openpyxl.load_workbook(SRC,read_only=True)
ws=wb['Sheet1']
it=ws.iter_rows(values_only=True); hdr=list(next(it)); idx={h:i for i,h in enumerate(hdr) if h}
rows=[r for r in it if r[idx['ID']] is not None]
g=lambda r,n: r[idx[n]]

BUCKET_MAP={'CostBucketDirectAllied':'Allied','CostBucketDirectAnaesthesia':'Anaesthesia',
 'CostBucketDirectED':'ED','CostBucketDirectICU':'ICU','CostBucketDirectImaging':'Imaging',
 'CostBucketDirectLaboratory':'Laboratory','CostBucketDirectPhysician':'Physician',
 'CostBucketDirectOP':'OP','CostBucketDirectOR':'OR','CostBucketDirectOther':'Other',
 'CostBucketDirectPharmacy':'Pharmacy','CostBucketDirectProsthesis':'Prosthesis',
 'CostBucketDirectSupplies':'Other','CostBucketDirectSPS':'SPS','CostBucketDirectWard':'Ward'}
TYPE_COLS=['SWNurs','SWDoc','SWAllied','SWNonClin','Laboratory','Imaging','Pharmacy',
 'Prostheses','MS','Hotel','GS','Depreciation','OHF','OHC','NABMs']
BUCKET_ORDER=['Allied','Anaesthesia','Dental','ED','Home','ICU','Imaging','Laboratory',
 'Physician','OP','OR','Other','Pharmacy','Prosthesis','SPS','Ward']
TYPE_ORDER=['SWNurs','SWDoc','SWAllied','SWNonClin','Laboratory','Imaging','Pharmacy',
 'Prostheses','MS','Hotel','GS','Depreciation','OHF','OHDepreciation','OHC','NABMs']

claims=defaultdict(list); order=[]
for r in rows:
    k=g(r,'ID')
    if k not in claims: order.append(k)
    claims[k].append(r)

def fdt(d): return d.strftime('%d/%m/%Y %H:%M')
def amt(v): return f"{round(v,3):.3f}"

def build_claim(cid, rs):
    r0=rs[0]
    direct=defaultdict(float); types=defaultdict(float)
    for r in rs:
        for col,tag in BUCKET_MAP.items(): direct[tag]+=g(r,col) or 0.0
        for t in TYPE_COLS: types[t]+=g(r,f'CostTypeDetails{t}') or 0.0
    encid=str(g(r0,'EncounterID')).strip()
    diag=str(g(r0,'Diagnosis Code') or '').strip()
    a=[]
    a.append('  <Claim>')
    a.append(f'    <ID>{escape(str(cid))}</ID>')
    a.append('    <ProviderID>MF4159</ProviderID>')
    a.append('    <Encounter>')
    a.append('      <FacilityID>MF4159</FacilityID>')
    a.append(f'      <ID>{escape(encid)}</ID>')
    a.append(f'      <ActivityClinicianID>{escape(str(g(r0,"Doctor No")).strip())}</ActivityClinicianID>')
    a.append(f'      <PatientID>{escape(str(g(r0,"FileNo.")).strip())}</PatientID>')
    a.append('      <Type>1</Type>')
    a.append(f'      <Start>{fdt(g(r0,"StartDateTime"))}</Start>')
    a.append(f'      <End>{fdt(g(r0,"EndDateTime"))}</End>')
    a.append('      <StartType>1</StartType>')
    a.append(f'      <Specialty>{escape(str(g(r0,"Speciality")).strip())}</Specialty>')
    a.append('      <SubSpecialty></SubSpecialty>')
    a.append('      <DRGCode></DRGCode>')
    a.append('      <ServiceCode></ServiceCode>')
    a.append('      <EndType>1</EndType>')
    a.append('      <Diagnosis>')
    a.append('        <Type>Principal</Type>')
    a.append(f'        <Code>{escape(diag)}</Code>')
    a.append('      </Diagnosis>')
    a.append('      <Procedure>')
    a.append('        <Type></Type>')
    a.append('        <Code></Code>')
    a.append('        <Duration>0</Duration>')
    a.append('      </Procedure>')
    a.append('      <EDTriageLevel>0</EDTriageLevel>')
    a.append('      <OPAttendanceType></OPAttendanceType>')
    a.append('      <TheatreAttendanceFlag>0</TheatreAttendanceFlag>')
    a.append('      <CriticalCareTime>0</CriticalCareTime>')
    a.append('      <VentilationTime>0</VentilationTime>')
    a.append('      <CostBucketDirect>')
    for t in BUCKET_ORDER: a.append(f'        <{t}>{amt(direct.get(t,0.0))}</{t}>')
    a.append('      </CostBucketDirect>')
    a.append('      <CostBucketOverheads>')
    for t in BUCKET_ORDER: a.append(f'        <{t}>0.000</{t}>')
    a.append('      </CostBucketOverheads>')
    a.append('      <CostTypeDetails>')
    for t in TYPE_ORDER: a.append(f'        <{t}>{amt(types.get(t,0.0))}</{t}>')
    a.append('      </CostTypeDetails>')
    a.append('    </Encounter>')
    a.append('  </Claim>')
    tot=sum(direct.values())+sum(types.values())
    return '\n'.join(a), tot

bym=defaultdict(list)
for k in order:
    bym[claims[k][0][idx['StartDateTime']].month].append(k)

schema=etree.XMLSchema(etree.parse(XSD))
TX=datetime.datetime.now().strftime('%d/%m/%Y %H:%M')
grand=0.0; excluded=[]
for m in sorted(bym):
    blocks=[]; tot=0.0
    for cid in bym[m]:
        b,t=build_claim(cid,claims[cid])
        if t<=0: excluded.append((m,cid)); continue
        blocks.append(b); tot+=t
    xml=('<Cost.Submission>\n  <Header>\n    <SenderID>MF4159</SenderID>\n'
         '    <ReceiverID>HAAD</ReceiverID>\n'
         f'    <TransactionDate>{TX}</TransactionDate>\n'
         f'    <RecordCount>{len(blocks)}</RecordCount>\n'
         '    <DispositionFlag>PTE_SUBMIT</DispositionFlag>\n  </Header>\n'
         +'\n'.join(blocks)+'\n</Cost.Submission>\n')
    fn=f'{OUT}/CostSubmission_MF4159_2025_M{m:02d}.xml'
    open(fn,'w',encoding='utf-8').write(xml)
    ok=schema.validate(etree.parse(fn))
    if not ok:
        print(f'M{m:02d} XSD FAIL:'); print(schema.error_log)
    grand+=tot
    print(f'M{m:02d}: {len(blocks):>4} claims  {os.path.getsize(fn)/1e6:>5.2f} MB  total {tot:,.3f}  XSD {"PASS" if ok else "FAIL"}')
print(f'grand total {grand:,.3f}')
if excluded:
    print('excluded zero-cost claims:',excluded)
    open(f'{OUT}/EXCLUDED_zero_cost_claims.txt','w').write(
     'Claims excluded from the Cost.Submission XML: total cost 0.000 (DOH blocks zero-total encounters).\n\n'
     +'\n'.join(f'M{m:02d}  {c}' for m,c in excluded)+'\n')
else:
    print('no zero-cost claims — nothing excluded')
