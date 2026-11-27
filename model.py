import json, io
import numpy as np
import openpyxl
books = {}
models = {}

def emit(x):
    return json.dumps(x, ensure_ascii=False, allow_nan=False)

def load_dataset(dataset_id, raw):
    if dataset_id not in {'basic','small','medium','large'}:
        raise ValueError('알 수 없는 분류표')
    wb = openpyxl.load_workbook(io.BytesIO(bytes(raw)), read_only=True, data_only=True)
    try:
        book = {s: list(wb[s].values) for s in wb.sheetnames}
    finally:
        wb.close()
    def industry_rows(rows):
        end = next(i for i,r in enumerate(rows[6:],6) if str(r[0])=='9590')
        return rows[6:end]
    base = book['국산투입계수(Ad)']
    industries = industry_rows(base)
    codes = [str(r[0]) for r in industries]
    names = [str(r[1]) for r in industries]
    if not codes or len(set(codes)) != len(codes):
        raise ValueError('산업 코드가 없거나 중복됩니다.')
    def matrix(s):
        rows = book[s]
        body = industry_rows(rows)
        cols = [str(c) for c in rows[4][2:2+len(codes)]]
        rr = [str(r[0]) for r in body]
        if sorted(cols)!=sorted(codes) or sorted(rr)!=sorted(codes):
            raise ValueError(s+': 행·열 산업 코드 불일치')
        ri = dict(zip(rr,body)); ci = {c:i for i,c in enumerate(cols,2)}
        a = np.array([[ri[c][ci[k]] for k in codes] for c in codes],dtype=float)
        if not np.isfinite(a).all(): raise ValueError(s + ': 비수치 또는 결측 데이터')
        return a
    matrix('총투입계수(A)')
    L = matrix('생산유발계수'); M = matrix('수입유발계수'); V = matrix('부가가치유발계수')
    ad = matrix('국산투입계수(Ad)'); am = matrix('수입투입계수(Am)')
    rows=book['총투입계수(A)']; ci={str(c):i for i,c in enumerate(rows[4])}
    vr=next(r for r in rows if str(r[0])=='9690')
    v=np.array([vr[ci[c]] for c in codes],dtype=float)
    residuals = {'생산':float(np.max(np.abs((np.eye(len(codes))-ad)@L-np.eye(len(codes))))),
                 '수입':float(np.max(np.abs(am@L-M))),
                 '부가가치':float(np.max(np.abs(v[:,None]*L-V)))}
    if not all(np.isfinite(r) and r < 1e-8 for r in residuals.values()): raise ValueError('계수 검증 실패: '+str(residuals))
    models[dataset_id] = (codes,names,L,M,V)
    books[dataset_id] = book
    return emit({'sectors':[{'code':c,'name':n} for c,n in zip(codes,names)],
                 'sheets':list(book),'residuals':residuals})

def explore(dataset_id, sheet, rowcode):
    rows=books[dataset_id][sheet]
    choices=[{'code':str(r[0]),'name':str(r[1])} for r in rows[6:] if len(r)>1 and r[0] is not None and r[1] is not None]
    chosen=next((r for r in rows[6:] if str(r[0])==rowcode),None)
    if chosen is None: chosen=next(r for r in rows[6:] if r[0] is not None and r[1] is not None)
    cells=[]
    for i in range(2,len(rows[4])):
        if rows[4][i] is None: continue
        val=chosen[i] if i<len(chosen) else None
        if isinstance(val,(float,int)): val=float(val)
        elif val is not None: val=str(val)
        cells.append({'code':str(rows[4][i]),'name':str(rows[5][i] or ''),'value':val})
    return emit({'rows':choices,'selected':str(chosen[0]),'name':str(chosen[1]),
                 'unit':str(rows[3][0] or '係数').replace('係数','계수 / 비율: 행·열 명칭 확인'), 'cells':cells})

def simulate(payload):
    request = json.loads(payload)
    dataset = request['dataset']
    if dataset not in models: raise ValueError('분류표를 먼저 불러오세요.')
    codes,names,L,M,V = models[dataset]
    index = {c:i for i,c in enumerate(codes)}
    f = np.zeros(len(codes))
    for s in request['shocks']:
        a = float(s['amount']); code = str(s['code'])
        if code not in index: raise ValueError('해당 분류표에 없는 코드: '+code)
        if not np.isfinite(a) or a<=0: raise ValueError('유한한 양의 증가액을 입력하세요.')
        f[index[code]] += a
    if not np.isfinite(f.sum()) or f.sum()<=0: raise ValueError('총 증가액을 확인하세요.')
    x=L@f; m=M@f; v=V@f
    result=[{'code':c,'name':n,'production':float(a),'imports':float(b),'valueAdded':float(d)}
            for c,n,a,b,d in zip(codes,names,x,m,v)]
    result.sort(key=lambda r:r['production'],reverse=True)
    return emit({'demand':float(f.sum()),'production':float(x.sum()),'imports':float(m.sum()),
                 'valueAdded':float(v.sum()),'multiplier':float(x.sum()/f.sum()),'rows':result})

def initialize(raw):
    return load_dataset('basic', raw)
