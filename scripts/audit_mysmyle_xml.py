import sys
from lxml import etree
from collections import defaultdict
import os
MY, MS = sys.argv[1], sys.argv[2]
print("sizes: mine",os.path.getsize(MY)," mysmyle",os.path.getsize(MS))
ms=etree.parse(MS); my=etree.parse(MY)
schema=etree.XMLSchema(etree.parse('shafafiya/CostSubmission_repaired.xsd'))
okm=schema.validate(ms)
print("XSD: mine =",schema.validate(my),"; MySmyle =",okm)
if not okm:
    for e in schema.error_log[:10]: print("  MS:",e.message)
def hdr(t): return {c.tag:(c.text or '') for c in t.getroot().find('Header')}
print("HEADER mine:   ",hdr(my))
print("HEADER mysmyle:",hdr(ms))
META=['FacilityID','ID','ActivityClinicianID','PatientID','Type','Start','End','StartType','Specialty','SubSpecialty','DRGCode','ServiceCode','EndType','EDTriageLevel','OPAttendanceType','TheatreAttendanceFlag','CriticalCareTime','VentilationTime']
def load(t):
    d={}
    for c in t.getroot().findall('Claim'):
        cid=c.findtext('ID'); e=c.find('Encounter')
        meta={x:(e.findtext(x) or '') for x in META}
        meta['DiagType']=e.findtext('Diagnosis/Type') or ''; meta['DiagCode']=e.findtext('Diagnosis/Code') or ''
        meta['ProcType']=e.findtext('Procedure/Type') or ''; meta['ProcCode']=e.findtext('Procedure/Code') or ''
        am={}
        for blk in ['CostBucketDirect','CostBucketOverheads','CostTypeDetails']:
            for x in e.find(blk): am[f'{blk}.{x.tag}']=float(x.text)
        d[cid]=(meta,am)
    return d
A=load(my); B=load(ms)
print("\nclaims: mine",len(A)," mysmyle",len(B))
print("only in mine:",len(set(A)-set(B)),sorted(set(A)-set(B))[:4])
print("only in mysmyle:",len(set(B)-set(A)),sorted(set(B)-set(A))[:4])
both=set(A)&set(B)
mdiff=defaultdict(int); ex={}
for cid in both:
    for k in A[cid][0]:
        if A[cid][0][k]!=B[cid][0][k]:
            mdiff[k]+=1
            if k not in ex: ex[k]=(cid,A[cid][0][k],B[cid][0][k])
print("\nmeta diffs:")
if not mdiff: print("  none — every identity field matches")
for k,v in sorted(mdiff.items(),key=lambda x:-x[1]):
    c,x,y=ex[k]; print(f"  {k:<20} {v:>4}  e.g. {c}: mine={x!r} mysmyle={y!r}")
tolA=tolB=0.0; kd=defaultdict(lambda:[0,0.0]); big=[]; zero=neg=0
for cid in both:
    aa,ab=A[cid][1],B[cid][1]
    tb=sum(ab.values())
    if tb<=0: zero+=1
    if any(v<0 for v in ab.values()): neg+=1
    for k in set(aa)|set(ab):
        d=ab.get(k,0.0)-aa.get(k,0.0)
        if abs(d)>0.0005:
            kd[k][0]+=1; kd[k][1]+=d
            if abs(d)>1: big.append((abs(d),cid,k,aa.get(k,0.0),ab.get(k,0.0)))
    tolA+=sum(aa.values()); tolB+=tb
print(f"\ntotals: mine {tolA:,.3f}  mysmyle {tolB:,.3f}  diff {tolB-tolA:+,.3f}")
print("zero-total claims in MySmyle:",zero," negatives:",neg)
print("\nfield-level diffs:")
for k,(n,s) in sorted(kd.items(),key=lambda x:-x[1][0])[:15]:
    print(f"  {k:<40} {n:>5} cells  net {s:+,.3f}")
big.sort(reverse=True)
print("\nlargest cell diffs:")
for d,cid,k,va,vb in big[:8]:
    print(f"  {cid} {k}: mine {va:.3f} vs mysmyle {vb:.3f} ({vb-va:+.3f})")
