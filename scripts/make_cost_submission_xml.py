import openpyxl, os
from collections import defaultdict
from xml.sax.saxutils import escape

P='/home/user/Clinical-Costing/output/Clinical_Costing_2025_MF4159.xlsx'
OUT='/home/user/Clinical-Costing/output/xml'
os.makedirs(OUT,exist_ok=True)

wb=openpyxl.load_workbook(P,read_only=True)
ws=wb['Sheet1']
it=ws.iter_rows(values_only=True); hdr=list(next(it)); idx={h:i for i,h in enumerate(hdr) if h}
rows=[r for r in it if r[0] is not None]
g=lambda r,n: r[idx[n]]

# Excel column -> (block, tag). Gen2: Supplies bucket removed -> parked in Other (flagged in README).
BUCKET_MAP={'CostBucketDirectAllied':'Allied','CostBucketDirectAnaesthesia':'Anaesthesia',
 'CostBucketDirectED':'ED','CostBucketDirectICU':'ICU','CostBucketDirectImaging':'Imaging',
 'CostBucketDirectLaboratory':'Laboratory','CostBucketDirectPhysician':'Physician',
 'CostBucketDirectOP':'OP','CostBucketDirectOR':'OR','CostBucketDirectOther':'Other',
 'CostBucketDirectPharmacy':'Pharmacy','CostBucketDirectProsthesis':'Prosthesis',
 'CostBucketDirectSupplies':'Other',   # <-- Gen2 schema has no Supplies bucket
 'CostBucketDirectSPS':'SPS','CostBucketDirectWard':'Ward'}
TYPE_COLS=['SWNurs','SWDoc','SWAllied','SWNonClin','Laboratory','Imaging','Pharmacy',
 'Prostheses','MS','Hotel','GS','Depreciation','OHF','OHC','NABMs']
BUCKET_ORDER=['Allied','Anaesthesia','Dental','ED','Home','ICU','Imaging','Laboratory',
 'Physician','OP','OR','Other','Pharmacy','Prosthesis','SPS','Ward']
TYPE_ORDER=['SWNurs','SWDoc','SWAllied','SWNonClin','Laboratory','Imaging','Pharmacy',
 'Prostheses','MS','Hotel','GS','Depreciation','OHF','OHDepreciation','OHC','NABMs']

claims=defaultdict(lambda:{'rows':[]})
for r in rows: claims[g(r,'EncounterID')]['rows'].append(r)

def fdt(d): return d.strftime('%d/%m/%Y %H:%M')
def amt(v): return f"{round(v,3):.3f}"

def build_claim(cid, rs):
    r0=rs[0]
    direct=defaultdict(float); types=defaultdict(float)
    for r in rs:
        for col,tag in BUCKET_MAP.items(): direct[tag]+=g(r,col) or 0.0
        for t in TYPE_COLS: types[t]+=g(r,f'CostTypeDetails{t}') or 0.0
    diag=str(g(r0,'Diagnosis Code') or '').strip()
    lines=[]
    a=lines.append
    a('  <Claim>')
    a(f'    <ID>{escape(str(cid))}</ID>')
    a('    <ProviderID>MF4159</ProviderID>')
    a('    <Encounter>')
    a('      <FacilityID>MF4159</FacilityID>')
    a(f'      <ID>{escape(str(cid))}</ID>')
    a(f'      <ActivityClinicianID>{escape(str(g(r0,"Doctor No")).strip())}</ActivityClinicianID>')
    a(f'      <PatientID>{escape(str(g(r0,"FileNo.")).strip())}</PatientID>')
    a('      <Type>1</Type>')
    a(f'      <Start>{fdt(g(r0,"StartDateTime"))}</Start>')
    a(f'      <End>{fdt(g(r0,"EndDateTime"))}</End>')
    a('      <StartType>1</StartType>')
    a(f'      <Specialty>{escape(str(g(r0,"Speciality")).strip())}</Specialty>')
    a('      <SubSpecialty></SubSpecialty>')
    a('      <DRGCode></DRGCode>')
    a('      <ServiceCode></ServiceCode>')
    a('      <EndType>1</EndType>')
    a('      <Diagnosis>')
    a('        <Type>Principal</Type>')
    a(f'        <Code>{escape(diag)}</Code>')
    a('      </Diagnosis>')
    a('      <Procedure>')
    a('        <Type></Type>')
    a('        <Code></Code>')
    a('        <Duration>0</Duration>')
    a('      </Procedure>')
    a('      <EDTriageLevel>0</EDTriageLevel>')
    a('      <OPAttendanceType></OPAttendanceType>')
    a('      <TheatreAttendanceFlag>0</TheatreAttendanceFlag>')
    a('      <CriticalCareTime>0</CriticalCareTime>')
    a('      <VentilationTime>0</VentilationTime>')
    a('      <CostBucketDirect>')
    for t in BUCKET_ORDER: a(f'        <{t}>{amt(direct.get(t,0.0))}</{t}>')
    a('      </CostBucketDirect>')
    a('      <CostBucketOverheads>')
    for t in BUCKET_ORDER: a(f'        <{t}>0.000</{t}>')
    a('      </CostBucketOverheads>')
    a('      <CostTypeDetails>')
    for t in TYPE_ORDER: a(f'        <{t}>{amt(types.get(t,0.0))}</{t}>')
    a('      </CostTypeDetails>')
    a('    </Encounter>')
    a('  </Claim>')
    total=sum(direct.values())+sum(types.values())
    return '\n'.join(lines), total

# group claims by month of encounter start
bym=defaultdict(list)
for cid,d in claims.items():
    m=d['rows'][0][idx['StartDateTime']].month
    bym[m].append(cid)

TXDATE='26/08/2026 09:00'
grand=0.0; files=[]; excluded=[]
for m in sorted(bym):
    blocks=[]; tot=0.0
    for cid in sorted(bym[m]):
        b,t=build_claim(cid,claims[cid]['rows'])
        if t<=0:
            # DOH blocks zero-total encounters; exclude and log (contributes 0.000 to totals)
            excluded.append((m,cid)); continue
        blocks.append(b); tot+=t
    xml=('<Cost.Submission>\n'
         '  <Header>\n'
         '    <SenderID>MF4159</SenderID>\n'
         '    <ReceiverID>HAAD</ReceiverID>\n'
         f'    <TransactionDate>{TXDATE}</TransactionDate>\n'
         f'    <RecordCount>{len(blocks)}</RecordCount>\n'
         '    <DispositionFlag>PTE_SUBMIT</DispositionFlag>\n'
         '  </Header>\n'
         +'\n'.join(blocks)+'\n</Cost.Submission>\n')
    fn=f'{OUT}/CostSubmission_MF4159_2025_M{m:02d}.xml'
    open(fn,'w',encoding='utf-8').write(xml)
    sz=os.path.getsize(fn)
    files.append((fn,len(blocks),sz,tot)); grand+=tot
    print(f'M{m:02d}: {len(blocks):>4} claims  {sz/1e6:>5.2f} MB  total {tot:,.3f}')
print(f'TOTAL claims {sum(f[1] for f in files)}   grand total {grand:,.3f}')
print('EXCLUDED zero-cost claims:', excluded)
open(f'{OUT}/EXCLUDED_zero_cost_claims.txt','w').write(
  'Claims excluded from the Cost.Submission XML because their total cost is 0.000\n'
  '(DOH portal blocks zero-total encounters; each consists only of code 99555,\n'
  ' "+E Additional Expense of Materials", priced at 0.00 in the CPT table).\n\n'
  +'\n'.join(f'M{m:02d}  {c}' for m,c in excluded)+'\n')
