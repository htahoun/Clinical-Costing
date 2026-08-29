import openpyxl, pickle, build2_lib as B
from collections import defaultdict, Counter

U='/root/.claude/uploads/ad1e78a6-69df-510f-b708-f3b8c52824be'
CALF=f'{U}/e3f8baff-Calibrate_Claims_2025.xlsx'
MYSF=f'{U}/99bf9915-Mysmyle_Claims_2025.xlsx'
OUT='/home/user/Clinical-Costing/output/Clinical_Costing_2025_MF4159_v3_BestOfThree.xlsx'
MONTHS=B.MONTHS

def load_full(path):
    wb=openpyxl.load_workbook(path,data_only=True,read_only=True)
    months={}
    hdr0=None
    for mi,m in enumerate(MONTHS,1):
        ws=wb[m]; it=ws.iter_rows(values_only=True); hdr=list(next(it))
        idx={h:i for i,h in enumerate(hdr) if h}
        for a,b in [('Provider ID','ProviderID'),('DiagnosisCode','Diagnosis Code'),('DiagnosisType','Diagnosis Type')]:
            if a in idx and b not in idx: idx[b]=idx[a]
        if hdr0 is None: hdr0=(hdr,idx)
        rows=[]; empty=0
        for r in it:
            if r[idx['ID']] in (None,''):
                empty+=1
                if empty>200: break
                continue
            empty=0; rows.append(r)
        months[mi]=(rows,idx)
    return months,hdr0

cal,calh=load_full(CALF)
mys,mysh=load_full(MYSF)
X=pickle.load(open('audit3/extra.pkl','rb'))
REPLACE={c for c,_,_ in X['disputes']}          # 15 claims: referee says MySmyle's lines are right
print("claims to take from MySmyle (referee-corrected):",len(REPLACE))

# claim -> rows maps
def bymap(months):
    m=defaultdict(list)
    for mi,(rows,idx) in months.items():
        for r in rows: m[str(r[idx['ID']]).strip()].append((mi,r))
    return m
CALM=bymap(cal); MYSM=bymap(mys)
mys_only=sorted(set(MYSM)-set(CALM))
print("MySmyle-only claims added:",len(mys_only))

# canonical column order = Cherie/v2 template header (Calibrate header equals it modulo variants)
TEMPLATE_HDR=[h for h in calh[0]]
def norm(h):
    return {'Provider ID':'ProviderID','DiagnosisCode':'Diagnosis Code','DiagnosisType':'Diagnosis Type'}.get(h,h)
TEMPLATE_HDR=[norm(h) if h else h for h in TEMPLATE_HDR]

def to_template(r, idx):
    o=[None]*len(TEMPLATE_HDR)
    for j,h in enumerate(TEMPLATE_HDR):
        if not h: continue
        src=idx.get(h)
        if src is None:
            for a,b in [('ProviderID','Provider ID'),('Diagnosis Code','DiagnosisCode'),('Diagnosis Type','DiagnosisType')]:
                if h==a: src=idx.get(b)
        if src is not None and src<len(r): o[j]=r[src]
    return o

pos={h:j for j,h in enumerate(TEMPLATE_HDR) if h}
ENC=pos['EncounterID']; SPEC=pos['Speciality']; DOC=pos['Doctor No']; SDT=pos['StartDateTime']; CODE=pos['ProcedureCode']

merged=defaultdict(list)   # mi -> list of (template_row, source_label)
stats=Counter(); stripped=0
for cid,pairs in CALM.items():
    if cid in REPLACE:
        for mi,r in MYSM[cid]:
            merged[mi].append((to_template(r,MYSM_idx:=mysh[1] if False else mys[mi][1]),'MySmyle (referee-corrected)'))
        stats['referee-corrected']+=1
    else:
        for mi,r in pairs:
            merged[mi].append((to_template(r,cal[mi][1]),'Calibrate'))
        stats['calibrate']+=1
for cid in mys_only:
    for mi,r in MYSM[cid]:
        merged[mi].append((to_template(r,mys[mi][1]),'MySmyle'))
    stats['mysmyle-only']+=1

# hygiene: strip trailing letters from encounter serials; speciality modal fallback
spec_cnt=defaultdict(Counter)
for months_,src in [(mys,'m'),(cal,'c')]:
    for mi,(rows,idx) in months_.items():
        for r in rows:
            d=str(r[idx['Doctor No']] or '').strip(); sv=r[idx['Speciality']]
            if sv not in (None,''): spec_cnt[d][str(sv).strip()]+=1
SPEC_FB={d:c.most_common(1)[0][0] for d,c in spec_cnt.items()}
spec_fixed=0
for mi,rows in merged.items():
    for o,_ in rows:
        e=str(o[ENC] or '')
        e2=e.rstrip('B')
        if e2!=e: o[ENC]=e2; stripped+=1
        if o[SPEC] in (None,''):
            o[SPEC]=SPEC_FB.get(str(o[DOC] or '').strip(),'General Dentist'); spec_fixed+=1
print(f"claims: calibrate-base {stats['calibrate']}, referee-corrected {stats['referee-corrected']}, mysmyle-only {stats['mysmyle-only']}")
print(f"serials cleaned of B suffix: {stripped}; speciality blanks filled: {spec_fixed}")
tot_lines=sum(len(v) for v in merged.values())
print("total lines:",tot_lines," total claims:",stats['calibrate']+stats['referee-corrected']+stats['mysmyle-only'])

# ---------- costing engine (identical method to v2) ----------
T,fixes=B.load_alloc()
B.ORPHAN_MAP=dict(B.ORPHAN_MAP); B.ORPHAN_MAP['27113']='27111'
def code_of(o): return str(o[CODE] or '').strip().zfill(5)
unknown=Counter()
def look(c):
    if c in T: return T[c]
    if c in B.ORPHAN_MAP: return T[B.ORPHAN_MAP[c]]
    if c!='00000': unknown[c]+=1
    return (0.0,[0.0]*17)
MODEL={}; FACT={}
for mi in range(1,13):
    s=sum(look(code_of(o))[0] for o,_ in merged[mi]); MODEL[mi]=s; FACT[mi]=B.GL[mi]/s
if unknown: print("UNKNOWN CODES:",dict(unknown))
cells={}
for mi in range(1,13):
    L=[]
    for o,_ in merged[mi]:
        cu,p=look(code_of(o)); sc=cu*FACT[mi]
        L.append([sc*p[j] for j in range(17)])
    cells[mi]=L
for mi in range(1,13):
    L=cells[mi]
    colsum=[sum(v[j] for v in L) for j in range(17)]
    tgt=[round(c,3) for c in colsum]
    d=round(round(B.GL[mi],3)-sum(tgt),3)
    jm=max(range(17),key=lambda j:tgt[j]); tgt[jm]=round(tgt[jm]+d,3)
    for j in range(17):
        for v in L: v[j]=round(v[j],3)
        s=round(sum(v[j] for v in L),3); dd=round(tgt[j]-s,3)
        if abs(dd)>=0.001:
            k=max(range(len(L)),key=lambda i:L[i][j]); L[k][j]=round(L[k][j]+dd,3)

TMAP={'Anaesthesia':'CostBucketDirectAnaesthesia','Imaging':'CostBucketDirectImaging',
 'Physician':'CostBucketDirectPhysician','OP':'CostBucketDirectOP','OtherLab':'CostBucketDirectLaboratory',
 'Prosthesis':'CostBucketDirectProsthesis','Supplies':'CostBucketDirectSupplies','SWNurs':'CostTypeDetailsSWNurs',
 'SWDoc':'CostTypeDetailsSWDoc','SWAllied':'CostTypeDetailsSWAllied','SWNonClin':'CostTypeDetailsSWNonClin',
 'MS':'CostTypeDetailsMS','Hotel':'CostTypeDetailsHotel','GS':'CostTypeDetailsGS',
 'Depreciation':'CostTypeDetailsDepreciation','OHF':'CostTypeDetailsOHF','OHC':'CostTypeDetailsOHC'}
costcols=[h for h in TEMPLATE_HDR if h and (str(h).startswith('CostBucket') or str(h).startswith('CostTypeDetails'))]

wb=openpyxl.Workbook(write_only=True)
ws=wb.create_sheet('Sheet1')
ws.append(TEMPLATE_HDR+['Source'])
for mi in range(1,13):
    for (o,src),v in zip(merged[mi],cells[mi]):
        for h in costcols: o[pos[h]]=0.0
        tot=0.0
        for j,c in enumerate(B.COLS): o[pos[TMAP[c]]]=v[j]; tot+=v[j]
        if 'total cost' in pos: o[pos['total cost']]=round(tot,3)
        ws.append(o+[src])

names=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
rc=wb.create_sheet('Reconciliation')
rc.append(['Vision Dental Clinic L.L.C · MF4159 · FY2025 · v3 BEST-OF-THREE'])
rc.append(['Universe: Calibrate claims + 23 MySmyle-only records; 15 claims line-corrected per ClaimScanner adjudication.'])
rc.append(['Method: scale-only to the expenditure file. Metadata: Calibrate. Serials cleaned. Source column marks provenance.'])
rc.append([])
rc.append(['Month','Expenditure','Costed total','Difference'])
tg=tc2=0.0
for mi in range(1,13):
    c=round(sum(sum(v) for v in cells[mi]),3); d=round(B.GL[mi]-c,3); tg+=B.GL[mi]; tc2+=c
    rc.append([names[mi-1],round(B.GL[mi],3),c,d])
rc.append(['YEAR',round(tg,3),round(tc2,3),round(tg-tc2,3)])
rc.append([])
rc.append(['Cost-centre mix (% of month)'])
rc.append(['Cost centre']+names+['YEAR %'])
for j,cn in enumerate(B.COLS):
    row=[TMAP[cn]]; yr=0.0
    for mi in range(1,13):
        s=sum(v[j] for v in cells[mi]); yr+=s; row.append(round(s/B.GL[mi]*100,2))
    row.append(round(yr/tc2*100,2)); rc.append(row)

fs=wb.create_sheet('Monthly Factors')
fs.append(['Month','Expenditure','Modelled cost','Factor'])
for mi in range(1,13): fs.append([names[mi-1],round(B.GL[mi],3),round(MODEL[mi],3),round(FACT[mi],6)])

lg=wb.create_sheet('Merge Log')
lg.append(['BEST-OF-THREE MERGE DECISIONS (from Claims_Audit_2025_Three_Sources.xlsx)'])
lg.append([])
lg.append(['1. 15 claims replaced with MySmyle lines (ClaimScanner adjudication sided with MySmyle):'])
for c,a,b in X['disputes']: lg.append([c,f'Calibrate net {a:,.2f}',f'adjudicated/MySmyle {b:,.2f}'])
lg.append([])
lg.append(['2. 23 MySmyle-only zero-value records added (Source=MySmyle).'])
lg.append(['3. 7 Calibrate-only zero-value records kept (in base).'])
lg.append(['4. 25 claims missing from MySmyle are present via the Calibrate base — including claim 1296MF4159-aa1544 (paid 350.00, ref 80285174).'])
lg.append([f'5. {stripped} encounter serials cleaned of stray B suffix.'])
lg.append([f'6. {spec_fixed} blank Speciality cells filled with the clinician\'s modal speciality.'])
lg.append(['7. Metadata (payer/diagnosis/times): Calibrate values kept per Finance Director decision; the 7 payer disputes (ClaimScanner shows C002) and 14 diagnosis + 7 time diffs remain OPEN items, listed in the audit workbook.'])
lg.append(['8. Code 27113 (6 zero-value lines, claim 1689MF4159-522959) priced from sibling 27111; one line of claim 727MF4159-82438e has a blank procedure code and carries zero cost.'])
lg.append(['9. Allocation-table adjustments (as v2):'])
for f in fixes: lg.append(['   '+f])
wb.save(OUT)
print("saved",OUT)
