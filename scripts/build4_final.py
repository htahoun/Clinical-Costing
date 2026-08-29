# FINAL build: user-edited v3 (accepted master, 24,074 lines / 9,738 claims)
# re-scaled month by month to the UPDATED expenditure ledger.
# Lines/metadata are FROZEN — only the 45 cost columns + total cost are recomputed.
import openpyxl
from collections import defaultdict, Counter
import build2_lib as B
import gl_new  # GL_NEW: {1..12: month expenditure} from the updated ledger

U='/root/.claude/uploads/ad1e78a6-69df-510f-b708-f3b8c52824be'
SRC=f'{U}/0f3963ce-Claude_Clinical_Costing_2025_MF4159_v3.xlsx'
OUT='/home/user/Clinical-Costing/output/Clinical_Costing_2025_MF4159_FINAL.xlsx'
GL=gl_new.GL_NEW

wb=openpyxl.load_workbook(SRC,data_only=True,read_only=True)
ws=wb['Sheet1']
it=ws.iter_rows(values_only=True); hdr=list(next(it))
pos={h:i for i,h in enumerate(hdr) if h}
CODE=pos['ProcedureCode']; SDT=pos['StartDateTime']
rows=[list(r) for r in it if r[pos['ID']] not in (None,'')]
print('lines:',len(rows),' claims:',len({r[pos['ID']] for r in rows}))

bym=defaultdict(list)
for r in rows: bym[r[SDT].month].append(r)
assert sorted(bym)==list(range(1,13))

T,fixes=B.load_alloc()
B.ORPHAN_MAP=dict(B.ORPHAN_MAP); B.ORPHAN_MAP['27113']='27111'
unknown=Counter()
def look(c):
    if c in T: return T[c]
    if c in B.ORPHAN_MAP: return T[B.ORPHAN_MAP[c]]
    if c!='00000': unknown[c]+=1
    return (0.0,[0.0]*17)
def code_of(r): return str(r[CODE] or '').strip().zfill(5)

MODEL={}; FACT={}
for mi in range(1,13):
    s=sum(look(code_of(r))[0] for r in bym[mi]); MODEL[mi]=s; FACT[mi]=GL[mi]/s
if unknown: print('UNKNOWN CODES:',dict(unknown))

cells={}
for mi in range(1,13):
    L=[]
    for r in bym[mi]:
        cu,p=look(code_of(r)); sc=cu*FACT[mi]
        L.append([sc*p[j] for j in range(17)])
    cells[mi]=L
# rounding: per-column tie, then month total == round(GL,3) exactly
for mi in range(1,13):
    L=cells[mi]
    colsum=[sum(v[j] for v in L) for j in range(17)]
    tgt=[round(c,3) for c in colsum]
    d=round(round(GL[mi],3)-sum(tgt),3)
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
costcols=[h for h in hdr if h and (str(h).startswith('CostBucket') or str(h).startswith('CostTypeDetails'))]

out=openpyxl.Workbook(write_only=True)
w1=out.create_sheet('Sheet1')
w1.append(hdr)
for mi in range(1,13):
    for r,v in zip(bym[mi],cells[mi]):
        for h in costcols: r[pos[h]]=0.0
        tot=0.0
        for j,c in enumerate(B.COLS): r[pos[TMAP[c]]]=v[j]; tot+=v[j]
        if 'total cost' in pos: r[pos['total cost']]=round(tot,3)
        w1.append(r)

names=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
rc=out.create_sheet('Reconciliation')
rc.append(['Vision Dental Clinic L.L.C · MF4159 · FY2025 · FINAL (accepted v3 lines × updated expenditure ledger)'])
rc.append(['Universe: the Finance-Director-accepted v3 master (24,074 lines / 9,738 claims; 6 lines removed by hand).'])
rc.append(['Method: scale-only to the UPDATED monthly expenditure ledger. No line, metadata or code changes.'])
rc.append([])
rc.append(['Month','Expenditure (updated)','Costed total','Difference'])
tg=tc2=0.0
for mi in range(1,13):
    c=round(sum(sum(v) for v in cells[mi]),3); d=round(GL[mi]-c,3); tg+=GL[mi]; tc2+=c
    rc.append([names[mi-1],round(GL[mi],3),c,d])
rc.append(['YEAR',round(tg,3),round(tc2,3),round(tg-tc2,3)])
rc.append([])
rc.append(['Cost-centre mix (% of month)'])
rc.append(['Cost centre']+names+['YEAR %'])
for j,cn in enumerate(B.COLS):
    row=[TMAP[cn]]; yr=0.0
    for mi in range(1,13):
        s=sum(v[j] for v in cells[mi]); yr+=s; row.append(round(s/GL[mi]*100,2))
    row.append(round(yr/tc2*100,2)); rc.append(row)

fs=out.create_sheet('Monthly Factors')
fs.append(['Month','Expenditure (updated)','Modelled cost','Factor'])
for mi in range(1,13): fs.append([names[mi-1],round(GL[mi],3),round(MODEL[mi],3),round(FACT[mi],6)])

lg=out.create_sheet('Merge Log')
lg.append(['FINAL BUILD — provenance and decisions'])
lg.append([])
lg.append(['1. Line universe = the accepted v3 best-of-three master as manually edited by the Finance Director:'])
lg.append(['   - claim 1296MF4159-aa1544 removed (3 lines, Jan) — Finance Director decision, final.'])
lg.append(['   - claims 1680MF4159-6d63f3 and 727MF4159-82438e removed (zero-value, Mar).'])
lg.append(['   - one zero-value 99555 line removed from claim 2336MF4159-4d736e (Nov).'])
lg.append(['2. v3 merge decisions carried over: Calibrate base + 15 ClaimScanner-adjudicated claims from MySmyle'])
lg.append(['   + 23 MySmyle-only records; Source column marks provenance; serials cleaned; specialities filled.'])
lg.append(['3. Costs re-scaled to the UPDATED monthly expenditure ledger supplied by the accounting team.'])
lg.append(['4. Code 27113 priced from sibling 27111; allocation-table adjustments as v2:'])
for f in fixes: lg.append(['   '+f])
out.save(OUT)
print('saved',OUT)
