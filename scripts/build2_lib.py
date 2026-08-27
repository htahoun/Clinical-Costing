import openpyxl
from collections import defaultdict

U='/root/.claude/uploads/ad1e78a6-69df-510f-b708-f3b8c52824be'
ACT=f'{U}/6923d0f9-Clinical_2025_one_sheet_Claude_Updated_Cherie.xlsx'
ALLOC=f'{U}/17526618-CPT_Allocation_Cheiri_updated.xlsx'
MONTHS=['January','February','March','April','May','June','July','August','September','October','November','December']
GL={1:480185.61748333334,2:424195.7997333334,3:614078.4143333333,4:560747.5943333333,
    5:498821.60129696963,6:722691.5700000001,7:540556.6923522727,8:592151.5070477272,
    9:492461.1900568182,10:537587.8430719697,11:540251.8011386364,12:725123.7600483283}
COLS=['Anaesthesia','Imaging','Physician','OP','OtherLab','Prosthesis','Supplies',
      'SWNurs','SWDoc','SWAllied','SWNonClin','MS','Hotel','GS','Depreciation','OHF','OHC']
CI={c:i for i,c in enumerate(COLS)}
# decision 4: the duplicate 14611 row priced 378 is really code 14612
ORPHAN_MAP={'12702':'12701','29302':'29301','42321':'42311'}

def load_alloc():
    ws=openpyxl.load_workbook(ALLOC,data_only=True,read_only=True)['Pecentage Allocation']
    T={}; fixes=[]
    seen14611=0
    for r in ws.iter_rows(values_only=True):
        if r[0] in (None,'CTP Code'): continue
        code=str(int(r[0])).zfill(5) if isinstance(r[0],(int,float)) else str(r[0]).strip().zfill(5)
        cu=float(r[2] or 0); pct=[float(r[3+j] or 0) for j in range(17)]
        if code=='14611':
            seen14611+=1
            if seen14611==2 and abs(cu-378.0)<0.01:
                code='14612'; fixes.append('14611 dup priced 378.00 re-keyed as 14612 (typo per Finance Director)')
        if code in T:
            fixes.append(f'{code} duplicate row identical — first kept'); continue
        s=sum(pct)
        if s>0 and abs(s-1.0)>1e-9:
            pct=[v/s for v in pct]; fixes.append(f'{code} percentages rescaled from {s*100:.1f}% to exactly 100%')
        T[code]=(cu,pct)
    return T,fixes

def load_activity():
    wb=openpyxl.load_workbook(ACT,data_only=True,read_only=True)
    months={}
    hdr0=None
    for mi,m in enumerate(MONTHS,1):
        ws=wb[m]; it=ws.iter_rows(values_only=True); hdr=list(next(it))
        if hdr0 is None: hdr0=hdr
        idx={h:i for i,h in enumerate(hdr) if h}
        rows=[]
        empty=0
        for r in it:
            if r[idx['ID']] in (None,''):
                empty+=1
                if empty>200: break
                continue
            empty=0
            rows.append(r)
        months[mi]=(rows,idx)
    return months,hdr0

def lookup(T,code):
    c=str(code).strip().zfill(5)
    if c in T: return T[c]
    return T[ORPHAN_MAP[c]]
