# ===================== Inventory sheet: daily stock snapshots + analysis =====================
# ===================== In Transit sheet (from the MIS 'In Transit' sheet, By Week layout) =====================
def build_mis_intransit(mis_path, DS, BW):
    """Final Pending Qty (col AI) by Final RM Code (col AN), in Monday weeks of Final AtA (col AJ)."""
    wb=wb_ro(mis_path)
    rows=list(wb['In Transit'].iter_rows(values_only=True))
    weeks=[pd.Timestamp(w).normalize() for w in BW['weeks']]; NW=len(weeks); w0=weeks[0]
    code2name={b['mat']:b['common'] for b in DS['blocks'] if b.get('common')}
    alias_=DS.get('alias') or {}
    code2uom={b['mat']:b['uom'] for b in DS['blocks']}
    q={}; outside={}; nodate={}; lines={}; desc={}
    for r in rows[1:]:
        if len(r)<40: continue
        code=str(r[39] or '').strip(); qty=r[34]
        if code in _SWP_INV() and isinstance(qty,(int,float)): qty=qty*_SWP_INV()[code][1]   # MIS tape code -> your BOM code's roll (CODE_SWAP)
        code=alias_.get(code,code)     # alternate code -> dashboard material
        if not code or not isinstance(qty,(int,float)) or isinstance(qty,bool) or qty<=0: continue
        if r[4] and code not in desc: desc[code]=' '.join(str(r[4]).split())[:45]
        d=r[35]
        if not isinstance(d,(dt.datetime,dt.date)): nodate[code]=nodate.get(code,0)+qty; continue
        d=pd.Timestamp(d).normalize(); k=((d-pd.Timedelta(days=d.weekday()))-w0).days//7
        sup=' '.join(str(r[1] or '').split()) or '(no supplier)'
        if 0<=k<NW: q.setdefault(code,{}).setdefault(sup,[0.0]*NW)[k]+=qty; lines[code]=lines.get(code,0)+1
        else: outside[code]=outside.get(code,0)+qty
    # In Transit has no customer (its 'Cust' column is the customs broker), so show who uses the material:
    # customers in MIS Total Plan Req for the plant over the same 26 weeks, with their share of the plan
    tp=list(wb['Total Plan Req'].iter_rows(values_only=True)); w_end=w0+pd.Timedelta(days=7*NW); plant=DS['plant'].upper(); cq={}
    for r in tp[1:]:
        if not tp_ok(r,plant) or not isinstance(r[0],(dt.datetime,dt.date)): continue
        d=pd.Timestamp(r[0]).normalize(); qv=r[14]
        if not (w0<=d<w_end) or not isinstance(qv,(int,float)) or qv<=0: continue
        cu=' '.join(str(r[2] or '').split()) or '(no customer)'; cd=str(r[10] or '').strip()
        if cd in _SWP_INV(): cd,f_=_SWP_INV()[cd]; qv=qv*f_
        cq.setdefault(cd,{}); cq[cd][cu]=cq[cd].get(cu,0.0)+qv
    def cust_text(code):
        d=cq.get(code) or {}; t=sum(d.values())
        if not t: return None
        return ' | '.join(f'{c} {v/t:.0%}' for c,v in sorted(d.items(),key=lambda kv:-kv[1]))
    def split(code,base):
        base=dict(base,customer=cust_text(code) if code else None)
        # one row per supplier (largest first); balances / sheet date only on the material's first row so they are not double counted
        sups=sorted((q.get(code) or {}).items(),key=lambda kv:-sum(kv[1])) if code else []
        if not sups: return [dict(base,supplier=None,w=[0.0]*NW)]
        return [dict(base,supplier=sp,w=w,**({} if i==0 else dict(asof=None,bal1=None,bal2=None))) for i,(sp,w) in enumerate(sups)]
    out=[]; seen=set()
    name2code={v:k for k,v in code2name.items()}
    tok_=lambda t: set(re.findall(r'[A-Z0-9.]+',str(t or '').upper()))
    def code_of(n):                                       # By Week name -> material; else the ONE material whose name words are all in it ('EPE 480 GSM' -> EPE 480)
        if n in name2code: return name2code[n]
        hit=[c_ for c_,v_ in code2name.items() if tok_(v_) and tok_(v_)<=tok_(n) and c_ not in name2code.values() or (tok_(v_) and tok_(v_)<=tok_(n) and c_ not in seen)]
        return hit[0] if len(hit)==1 else None
    for n,m in BW['meta'].items():                       # By Week materials first, same order
        code=code_of(n); seen.add(code)
        out+=split(code,dict(name=n,code=code,uom=m.get('uom') or code2uom.get(code,''),asof=m.get('asof'),bal1=m.get('bal1'),bal2=m.get('bal2')))
    for b in DS['blocks']:                               # dashboard materials not on By Week
        if b['mat'] in seen: continue
        seen.add(b['mat']); out+=split(b['mat'],dict(name=b['common'] or b['mat'],code=b['mat'],uom=b['uom'],asof=None,bal1=None,bal2=None))
    other=[x for c in sorted(q) if c not in seen for x in split(c,dict(name=(f'{c} - {desc[c]}' if c in desc and c.upper().startswith('RM') else c),code=c,uom='',asof=None,bal1=None,bal2=None))]
    return dict(rows=out,other=other,weeks=BW['weeks'],hdr1=BW['hdr1'],hdr2=BW['hdr2'],outside=outside,nodate=nodate,src=os.path.basename(mis_path),
                wsa=(BW.get('data') or {}).get('transit') or {})

def make_intransit_xlsx(IT, path, cache=None, compact=False):
    AR='Calibri'; hf=Font(name=AR,bold=True,color='FFFFFF',size=10); hfill=PatternFill('solid',fgColor='1F3864'); bf=Font(name=AR,size=10)
    bold=Font(name=AR,size=10,bold=True); blue=Font(name=AR,size=10,color='0000FF'); ttl=Font(name=AR,bold=True,size=12)
    ctr=Alignment(horizontal='center',wrap_text=True); num='#,##0;(#,##0);-'; dfy='dd-mmm-yy'
    wb=Workbook(); ws=wb.active; ws.title='In Transit'; NW=len(IT['weeks'])
    ws['A1']='In transit quantity by material and arrival week (MIS RM Dashboard, In Transit sheet)'; ws['A1'].font=ttl
    ws['A2']=(f"Source {IT['src']}, read {dt.date.today():%d-%b-%Y}. Quantity = Final Pending Qty; week = Monday of Final AtA "
              "(MIS: revised ETA port + 5 days; ETAs already past are set to today). Sheet stock date and balances are from the WSA planning sheet. "
              "Shipments carry no customer, so Customer lists who uses the material: customers in MIS Total Plan Req over the same 26 weeks, with their share of the plan.")
    ws['A2'].font=Font(name=AR,size=9,italic=True)
    def table(top,title,items):
        ws.cell(top,1,title).font=bold; h=top+1
        for j,lab in enumerate(['Material','Supplier Name','Customer (who uses it, share of plan)','UOM','Sheet stock date','Total',IT['hdr1'],IT['hdr2']]):
            c=ws.cell(h,1+j,lab); c.font=hf; c.fill=hfill; c.alignment=ctr
        for k,w in enumerate(IT['weeks']):
            c=ws.cell(h,9+k,w); c.number_format='dd-mmm'; c.font=hf; c.fill=hfill; c.alignment=ctr
        ws.row_dimensions[h].height=32
        for i,x in enumerate(items):
            r=h+1+i; ws.cell(r,1,x['name']).font=bf; ws.cell(r,2,x['supplier']).font=bf
            c=ws.cell(r,3,x.get('customer')); c.font=Font(name=AR,size=9); c.alignment=Alignment(wrap_text=True,vertical='top')
            ws.cell(r,4,x['uom']).font=bf
            if x['asof'] is not None: c=ws.cell(r,5,x['asof']); c.font=blue; c.number_format=dfy
            c=ws.cell(r,6,f'=SUM(I{r}:{L(8+NW)}{r})'); c.font=bold; c.number_format=num
            if cache is not None: cache[(r,6)]=sum(x['w'])
            for j,key in ((7,'bal1'),(8,'bal2')):
                if x[key] is not None: c=ws.cell(r,j,x[key]); c.font=blue; c.number_format=num
            for k,v in enumerate(x['w']):
                c=ws.cell(r,9+k,round(v,6) if v else 0); c.font=bf; c.number_format=num
        tr=h+1+len(items); ws.cell(tr,1,'Lines with quantity (count)').font=bold
        for k in range(NW):
            col=L(9+k); c=ws.cell(tr,9+k,f'=COUNTIF({col}{h+1}:{col}{tr-1},">0")'); c.font=bold
            if cache is not None: cache[(tr,9+k)]=sum(1 for x in items if x['w'][k]>0)
        return tr
    if compact:          # Supply sheet: the week-by-week table is the same data as MRP Daily 'Scheduled receipt' - left out
        ws['A1']='Supply - in transit MIS vs WSA planning file, and supplier coverage'
        ws['A2']=(f"Source {IT['src']}, read {dt.date.today():%d-%b-%Y}. In-transit quantities by week are on MRP Daily ('Scheduled receipt'); "
                  "shipment-by-shipment differences (BL / PO) are on the Check sheet.")
        tr=0
    else:
        tr=table(3,'In transit (PPC) - quantity in each material\'s own unit',IT['rows'])
        if IT['other']:
            tr=table(tr+3,'Other RM codes in transit (not on the Daily Dashboard)',IT['other'])
    notes=[]
    if IT['outside']: notes.append('Outside the 26 weeks: '+', '.join(f'{c} {v:,.0f}' for c,v in IT['outside'].items()))
    if IT['nodate']: notes.append('No Final AtA date: '+', '.join(f'{c} {v:,.0f}' for c,v in IT['nodate'].items()))
    for i,t in enumerate(notes): ws.cell(max(tr,2)+2+i,1,t).font=bf
    # ---------- comparison: PPC in transit vs WSA planning file (RM plan) in transit, by material and week ----------
    from openpyxl.formatting.rule import Rule
    from openpyxl.styles.differential import DifferentialStyle
    top=(tr+3+len(notes)) if not compact else 4+len(notes); NWc=NW
    ws.cell(top,1,'Comparison - in transit PPC In_Transit vs WSA RM plan (by arrival week)').font=ttl
    ws.cell(top+1,1,('PPC = PPC In_Transit Final Pending Qty by Final AtA week (as above, all suppliers). WSA = deliveries with a vessel on the WSA planning sheets, by plant-date week '
                     '(pooled glass sheets are split equally between front and back). Difference = PPC - WSA; amber = not equal (rounding under 5 ignored). Shipment-level detail is on the Check sheet.')).font=Font(name=AR,size=9,italic=True)
    h=top+2
    for j,lab in enumerate(['Material','Source','Match?','UOM','','Total']):
        c=ws.cell(h,1+j,lab or None); c.font=hf; c.fill=hfill; c.alignment=ctr
    for k,w in enumerate(IT['weeks']):
        c=ws.cell(h,9+k,w); c.number_format='dd-mmm'; c.font=hf; c.fill=hfill; c.alignment=ctr
    mis_by={}; uom_by={}
    for x in IT['rows']:
        a_=mis_by.setdefault(x['name'],[0.0]*NWc)
        for k,v in enumerate(x['w'][:NWc]): a_[k]+=v
        if x.get('uom'): uom_by.setdefault(x['name'],x['uom'])
    names=[n_ for n_ in dict.fromkeys(list(mis_by)+list(IT.get('wsa') or {})) if any(mis_by.get(n_,[0])) or any((IT.get('wsa') or {}).get(n_,[0]))]
    r=h+1; diff_rows=[]; amber=PatternFill('solid',fgColor='FFEB9C')
    for n_ in names:
        mv=(mis_by.get(n_) or [0.0]*NWc)[:NWc]; wv=list((IT['wsa'].get(n_) or [0.0]*NWc))[:NWc]+[0.0]*max(0,NWc-len(IT['wsa'].get(n_) or []))
        rm,rw,rd=r,r+1,r+2
        for rr,src,vals in ((rm,'PPC',mv),(rw,'WSA',wv)):
            ws.cell(rr,1,n_).font=bf; ws.cell(rr,2,src).font=bf; ws.cell(rr,4,uom_by.get(n_)).font=bf
            c=ws.cell(rr,6,f'=SUM(I{rr}:{L(8+NWc)}{rr})'); c.font=bold; c.number_format=num
            if cache is not None: cache[(rr,6)]=sum(vals)
            for k,v in enumerate(vals): c=ws.cell(rr,9+k,round(v,6) if v else 0); c.font=bf; c.number_format=num
        ws.cell(rd,1,n_).font=bold; ws.cell(rd,2,'Difference (PPC - WSA)').font=bold

        tol=max(5.0,0.0005*max(sum(mv),sum(wv)))          # rounding tolerance
        c=ws.cell(rd,3,f'=IF(SUMPRODUCT(ABS(I{rm}:{L(8+NWc)}{rm}-I{rw}:{L(8+NWc)}{rw}))<=MAX(5,0.0005*MAX(F{rm},F{rw})),"Match","Different")'); c.font=bold
        if cache is not None: cache[(rd,3)]='Match' if sum(abs(a-b) for a,b in zip(mv,wv))<=tol else 'Different'
        c=ws.cell(rd,6,f'=F{rm}-F{rw}'); c.font=bold; c.number_format=num
        if cache is not None: cache[(rd,6)]=sum(mv)-sum(wv)
        for k in range(NWc):
            col=L(9+k); c=ws.cell(rd,9+k,f'={col}{rm}-{col}{rw}'); c.font=bold; c.number_format=num
            d_=mv[k]-wv[k]
            if cache is not None: cache[(rd,9+k)]=d_
            if abs(d_)>5: c.fill=amber
        if sum(abs(a-b) for a,b in zip(mv,wv))>tol: ws.cell(rd,3).fill=amber; ws.cell(rd,6).fill=amber
        diff_rows.append(rd); r+=3
    if diff_rows:
        cells=' '.join(f'I{x}:{L(8+NWc)}{x} F{x}' for x in diff_rows)
        ws.conditional_formatting.add(cells,Rule(type='cellIs',operator='notBetween',formula=['-5','5'],dxf=DifferentialStyle(fill=PatternFill(bgColor='FFFFEB9C',fill_type='solid'))))
        ws.conditional_formatting.add(cells,Rule(type='cellIs',operator='between',formula=['-5','5'],dxf=DifferentialStyle(fill=PatternFill(bgColor='FFFFFFFF',fill_type='solid'))))
        ccol=' '.join(f'C{x}' for x in diff_rows)
        ws.conditional_formatting.add(ccol,Rule(type='containsText',operator='containsText',text='Different',formula=[f'NOT(ISERROR(SEARCH("Different",C{diff_rows[0]})))'],dxf=DifferentialStyle(fill=PatternFill(bgColor='FFFFEB9C',fill_type='solid'))))
        ws.conditional_formatting.add(ccol,Rule(type='containsText',operator='containsText',text='Match',formula=[f'NOT(ISERROR(SEARCH("Match",C{diff_rows[0]})))'],dxf=DifferentialStyle(fill=PatternFill(bgColor='FFC6EFCE',fill_type='solid'))))
    for c_,w_ in zip('ABCDEFGH',[30,34,48,8,13,11,16,16]): ws.column_dimensions[c_].width=w_
    for k in range(NW): ws.column_dimensions[L(9+k)].width=11
    ws.freeze_panes='I5'
    set_print(wb.active,'Supply'); wb.save(path)
    return len({x['name'] for x in IT['rows'] if any(x['w'])})

def own_stock(b):
    """Stock of the material without the JW stock at the OEM (user rule: JW at the OEM is in no calculation).
       After the OEM step b['stock'] is already own stock (jw_out); counted packaging (count_stock) never had JW."""
    s_=float(b.get('stock') or 0.0)
    return s_ if (b.get('jw_out') or b.get('count_stock')) else s_-float(b.get('st_jw') or 0.0)

def multi_sup(b):
    """Supplier detail is shown only when two or more suppliers have stock, in transit or plan for the material."""
    n=0
    for sp in b.get('sup') or []:
        if str(sp.get('name') or '').upper()=='UNASSIGNED': continue
        if float(sp.get('stock') or 0)>0 or sum(sp.get('transit') or [])>0 or sum(sp.get('plan') or [])>0: n+=1
    return n>=2
def cover_days(stock, plan, inflow, openpo):
    """Same rules as the Summary formulas: days until projected stock first goes below zero (len(plan) = no run-out).
       1 existing stock only, 2 + inflow (in transit + adhoc), 3 + inflow + open PO."""
    n=len(plan); out=[]
    for use_in,use_po in ((False,False),(True,False),(True,True)):
        cp=ci=co=0.0; k_=n
        for k in range(n):
            cp+=plan[k]; ci+=inflow[k] if use_in else 0.0; co+=openpo[k] if use_po else 0.0
            if (not use_in and cp>stock) or (use_in and stock+ci+co-cp<0): k_=k; break
        out.append(k_)
    return out


# ===================== days until stock-out (used for Status and the short periods) =====================



# ===================== Substitute Advice: which other material could replace one that runs out =====================
ENCAPS={'EVA','EPE','POE'}
BOM_GROUP_MAP={'aluminium frame':'Frame','aluminum frame':'Frame','silicon sealant':'Sealant','perforated adhesive tap':'Tape','cell fixation tap':'Tape','eva':'Encapsulant',
               'epe':'Encapsulant','poe':'Encapsulant','junction box':'Junction Box','glass':'Glass','busbar':'Busbar','ribbon':'Ribbon','potting a':'Potting A',
               'potting b':'Potting B','backsheet':'Backsheet','cells':'Cell','cell':'Cell'}
def read_bom_groups(mis_path):
    """MIS Standard BOM: the group of every material (column 'Understanding Name' + item 'Sr. no.') and the module types that use it.
       Also the role inside a module: for glass and encapsulant the lower item number of a module type = FRONT, the higher = BACK."""
    out={}
    try: rows=list(wb_ro(mis_path)['Standard BOM'].iter_rows(values_only=True))[1:]
    except Exception as _e_: _swallow(2186,_e_); return out
    per_mt={}
    for r in rows:
        if not r or len(r)<9 or not r[3]: continue
        c=str(r[3]).strip(); grp=BOM_GROUP_MAP.get(str(r[8] or '').strip().lower(),str(r[8] or '').strip())
        x=out.setdefault(c,dict(group=grp,sr=set(),mts=set(),pgr=r[7],role=set()))
        x['sr'].add(r[1]); x['mts'].add(str(r[2] or '').strip())
        per_mt.setdefault((str(r[2] or '').strip(),grp),[]).append((r[1] if isinstance(r[1],(int,float)) else 99,c))
    for (mt,grp),lst in per_mt.items():
        if grp in ('Glass','Encapsulant') and len(lst)>=2:
            lst=sorted(lst)
            out[lst[0][1]]['role'].add('front')
            for _,c in lst[1:]: out[c]['role'].add('back')
    return out



def read_sap_bom(mrp_path):
    """Your SAP BOM (MRP.xlsx 'BOM' sheet): customer, module model, component, description, quantity, unit, valid from, group (RM / PM / CS)."""
    out=[]
    if not mrp_path or not os.path.exists(mrp_path): return out
    try:
        wb=load_workbook(mrp_path,read_only=True,data_only=True)
        if 'BOM' not in wb.sheetnames: return out
        for r in list(wb['BOM'].iter_rows(values_only=True))[1:]:
            if r and len(r)>11 and r[0] and r[4]:
                out.append(dict(cust=' '.join(str(r[0]).split()),model=r[1],code=str(r[4]).strip(),desc=r[5],qty=r[6],uom=r[7],valid=r[8],grp=str(r[11] or '').strip().upper()))
    except Exception as ex_: print(f'NOTE: BOM sheet of {os.path.basename(mrp_path)} not read ({type(ex_).__name__})')
    return out

def read_mis_extra(mis_path, plant):
    """Other MIS data that helps substitution:
       - Total Plan Req 'JB Type' per BOM name (which JB the plan says a module uses)
       - RM Status for OEM JW: 'Make' (approved makers per code), 'Replace Material Code' (MIS's own replacement), stock at other sites
       - Booking Approval Check Summary: pending SO / booking approval per Common Name
       - MB52: unrestricted / quality inspection / blocked per code (plant + warehouse)"""
    import collections
    w=wb_ro(mis_path); X=dict(jbtype={},makes={},replace={},sites={},booking={},mbq={})
    def rows_(n):
        try: return list(w[n].iter_rows(values_only=True))
        except Exception as _e_: _swallow(2226,_e_); return []
    for r in rows_('Total Plan Req')[1:]:
        if r and len(r)>5 and r[4] and r[5] and str(r[21] if len(r)>21 else '').strip().upper() in ('',str(plant).upper()):
            X['jbtype'].setdefault(str(r[4]).strip(),set()).add(str(r[5]).strip())
    rs=rows_('RM Status for OEM JW')
    for r in rs[1:]:
        if not r or len(r)<6 or not r[1]: continue
        c=str(r[1]).strip()
        if c.upper().startswith('JW') or c=='Total': continue
        if r[0] and str(r[0]).strip() not in ('-',''): X['makes'].setdefault(c,[]).append(str(r[0]).strip().title())
        if len(r)>16 and r[16]: X['replace'].setdefault(c,set()).add(str(r[16]).strip()); X['replace'].setdefault(str(r[16]).strip(),set()).add(c)
        st={}
        for j,lab in ((9,'USA Arizona'),):
            if len(r)>j and isinstance(r[j],(int,float)) and r[j]>0: st[lab]=st.get(lab,0)+r[j]
        if st:
            for k,v in st.items(): X['sites'].setdefault(c,{})[k]=X['sites'].get(c,{}).get(k,0)+v
    for r in rs[1:]:                     # job-work stock sits on the JW code
        if r and len(r)>11 and r[1] and str(r[1]).startswith('JW') and isinstance(r[11],(int,float)) and r[11]>0:
            X['sites'].setdefault(str(r[1])[2:].strip(),{})['JW (OEM)']=r[11]
    for r in rows_('Booking Approval Check Summary')[1:]:
        if r and len(r)>32 and r[5] and r[4] and isinstance(r[31],(int,float)):
            X['booking'][str(r[5]).strip().lower()]=(r[31] or 0,r[32] if isinstance(r[32],(int,float)) else 0,(r[33] if len(r)>33 else None))
    pl={str(plant).upper(),str(plant).upper()+' WAREHOUSE'}
    for r in rows_('MB52')[1:]:
        if not r or len(r)<15 or str(r[13] or '').strip().upper() not in pl or not r[2]: continue
        q=X['mbq'].setdefault(str(r[2]).strip(),[0.0,0.0,0.0])
        for i_,j_ in ((0,6),(1,12),(2,9)): q[i_]+=r[j_] if isinstance(r[j_],(int,float)) else 0
    return X


BATCH_MASTER=[{},{}]   # [exact batch -> make, batch code (2-5 letters) -> make] from the 'Batch Master' sheet of MRP.xlsx
USER_ALT_MISS=[]       # your mappings whose main code is not a planned material
USER_ALT={}            # your Material Group 'MRP material' / 'Factor': code -> (main code, factor)
MAT_GROUP={}           # material -> (SAP group, your group name) from the 'Material Group' sheet of MRP.xlsx
MAT_GROUP_A={}         # material -> column A 'Material Group' of that sheet (Glass / Packaging / Label ...), also when it has no column E
MIS_CAT={}             # material -> MIS Standard BOM 'Understanding Name' (column I)
PPC_CAT=[None,{}]      # [PPC file read, material -> 'Item Category' of the PPC 'Count' sheet]
def load_ppc_cat():
    """PPC sheet 'Count', column 'Item Category' (found by its header) per material code; read once per PPC file."""
    p=PPC_SRC[0]
    if PPC_CAT[0]==(p or ''): return PPC_CAT[1]
    PPC_CAT[0]=p or ''; PPC_CAT[1]={}
    if not p: return PPC_CAT[1]
    try:
        rows=[list(r) for r in _wb_ro_raw(p)['Count'].iter_rows(values_only=True)]
        hi=next((i for i,r in enumerate(rows[:10]) if r and str(r[0] or '').strip().lower()=='material'),None)
        hl=[' '.join(str(x or '').split()).lower() for x in rows[hi]] if hi is not None else []
        if 'item category' in hl:
            j=hl.index('item category')
            for r in rows[hi+1:]:
                v=' '.join(str(r[j]).split()) if r and r[0] and j<len(r) and r[j] is not None else ''
                if v and not v.startswith('#'): PPC_CAT[1].setdefault(str(r[0]).strip(),v)
    except Exception: pass
    return PPC_CAT[1]
def item_category(m):
    """Category of a material - ONE rule for Material Status and the ALERT (user, Nov 2026: 'Alert sheet, I want item name and category
       from stock management'): your Material Group name (column E) > MIS Standard BOM Understanding Name (also via your BOM alias) >
       your Material Group column A (packaging / labels have no MIS BOM name) > '-'."""
    return ((MAT_GROUP.get(m) or (None,''))[1] or MIS_CAT.get(m) or next((MIS_CAT[c_] for c_,t_ in (BOM_ALIAS[0] or {}).items() if t_==m and c_ in MIS_CAT),'')
            or ' '.join(str(MAT_GROUP_A.get(m) or '').split()) or '-')
def load_mis_cat(mis_path):
    MIS_CAT.clear()
    try:
        for r in list(wb_ro(mis_path)['Standard BOM'].iter_rows(values_only=True))[1:]:
            if r and len(r)>8 and r[3] and r[8]: MIS_CAT.setdefault(str(r[3]).strip(),str(r[8]).strip())
    except Exception as _e_: _swallow(2273,_e_); pass
def load_user_masters(mrp_path):
    """Your Batch Master and Material Group sheets (MRP.xlsx)."""
    BATCH_MASTER[0].clear(); BATCH_MASTER[1].clear(); MAT_GROUP.clear(); MAT_GROUP_A.clear()
    try: w=wb_ro(mrp_path)
    except Exception as _e_: _swallow(2278,_e_); return
    if 'Batch Master' in w.sheetnames:
        for r in list(w['Batch Master'].iter_rows(values_only=True))[1:]:
            if r and r[0] and len(r)>1 and r[1]:
                b=str(r[0]).strip(); m=' '.join(str(r[1]).split())
                (BATCH_MASTER[0] if len(b)>=8 else BATCH_MASTER[1])[b.upper()]=m
    USER_ALT.clear()
    if 'Material Group' in w.sheetnames:
        rows_=list(w['Material Group'].iter_rows(values_only=True))
        hl_=[' '.join(str(x or '').split()).lower() for x in (rows_[0] if rows_ else [])]
        im_=hl_.index('mrp material') if 'mrp material' in hl_ else None; if_=hl_.index('factor') if 'factor' in hl_ else None
        for r in rows_[1:]:
            if r and len(r)>1 and r[1] and r[0]: MAT_GROUP_A.setdefault(str(r[1]).strip(),str(r[0]).strip())
            if r and len(r)>4 and r[1]:
                c_=str(r[1]).strip(); MAT_GROUP[c_]=(r[0],' '.join(str(r[4] or '').split()))
                t_=str(r[im_] or '').strip() if im_ is not None and im_<len(r) else ''
                if t_:
                    f_=r[if_] if if_ is not None and if_<len(r) and isinstance(r[if_],(int,float)) and r[if_] else 1.0
                    USER_ALT[c_]=(t_,float(f_))             # your mapping: this code counts as t_ (x factor); t_==code = main code

def batch_make(batch):
    """Supplier (make) of an SAP batch from your Batch Master: the exact batch, else its code - characters 2-3 (1JU0369004 -> JU),
       or a longer code (TAIC, VINI, ...) at the start or after the first character."""
    b=str(batch or '').strip().upper()
    if not b: return None
    if b in BATCH_MASTER[0]: return BATCH_MASTER[0][b]
    for L_ in (5,4,3,2):
        for st in (1,0):
            k=b[st:st+L_]
            if len(k)==L_ and k in BATCH_MASTER[1]: return BATCH_MASTER[1][k]
    return None

def batch_suppliers(mis_path, plant, makes=None):
    """MB52 stock per material per batch supplier code (2nd-3rd character of the batch), plant + its warehouse only,
       and a guessed supplier name per code from the PPC In_Transit suppliers of the same materials (fills BATCH_SEEN)."""
    import collections
    w=wb_ro(mis_path); by={}
    try: mb=list(w['MB52'].iter_rows(values_only=True))[1:]
    except Exception as _e_: _swallow(2316,_e_); return by
    pl={str(plant).upper(),str(plant).upper()+' WAREHOUSE'}
    for r in mb:
        if len(r)<15 or str(r[13] or '').strip().upper() not in pl: continue
        b=str(r[3] or '').strip(); c=str(r[2] or '').strip()
        k=b[1:3].upper() if len(b)>=3 and b[1:3].isalpha() else '?'
        mk_=batch_make(b)
        if mk_: k='='+mk_                         # named by your Batch Master
        q=float(r[14]) if isinstance(r[14],(int,float)) else 0.0
        by.setdefault(c,collections.Counter())[k]+=q
        if k!='?' and not k.startswith('='): BATCH_SEEN.setdefault(k,dict(mats=set(),guess=None))['mats'].add(c)
    sup=collections.defaultdict(collections.Counter)
    try:
        for r in list(w['In Transit'].iter_rows(values_only=True))[1:]:
            c=str(r[39] or r[31] or '').strip(); n=' '.join(str(r[1] or '').split())
            if c and n: sup[c][n]+=1
    except Exception as _e_: _swallow(2332,_e_); pass
    votes=collections.defaultdict(collections.Counter)
    for c,cnt in by.items():
        codes=[k for k in cnt if k!='?' and not k.startswith('=')]; names=sup.get(c)
        if not codes or not names: continue
        for k in codes:
            for n in names: votes[k][n]+=(3 if len(codes)==1 and len(names)==1 else 1)+(2 if any(w_.startswith(k[0]) and k[1] in w_ for w_ in n.upper().split()) else 0)
    for c,cnt in by.items():           # MIS RM Status 'Make' of the same material: a make whose name fits the code letters (JU -> Juren)
        for k in [k for k in cnt if k!='?' and not k.startswith('=')]:
            for m_ in (makes or {}).get(c,[]):
                if m_.upper().startswith(k[0]) and k[1] in m_.upper(): votes[k][m_]+=6
    for k,v in votes.items(): BATCH_SEEN[k]['guess']=v.most_common(1)[0][0]
    return by

def sup_name(k):
    if str(k).startswith('='): return str(k)[1:]                 # make from your Batch Master
    n=BATCH_SUP.get(k) or ((BATCH_SEEN.get(k) or {}).get('guess') and f"{BATCH_SEEN[k]['guess']} (guess)")
    return f"{k} = {n}" if n else k

SAP_GROUP_MAP={'W001':'Glass','W002':'Encapsulant','W003':'Cell','W005':'Ribbon','W007':'Tape','W008':'Frame','W009':'Junction Box','W010':'Sealant',
               'W011':'Busbar','W012':'Potting','W013':'Label','W021':'Tape'}
def sap_group_family(grp,desc):
    """Family from the SAP material group in MB52 (column K). Mixed groups are split by the description:
       W009 = junction boxes + connector caps + jumper cables, W010 / W012 = sealant + potting (base = A, catalyst = B)."""
    f=SAP_GROUP_MAP.get(str(grp or '').strip().upper())
    d=str(desc or '').upper()
    if f=='Junction Box':
        if 'CAP' in d: return 'Connector Cap'
        if 'JUMPER' in d or ('CABLE' in d and not re.search(r'JB|JUNCTION|ZH011|DSJB|KBT|VMSJ',d)): return 'Cable'
    if f in ('Sealant','Potting'):
        if 'POTTING' in d or 'PART A' in d or 'PART B' in d:
            return 'Potting B' if re.search(r'PART B|CATALYST|\bB\b',d) else 'Potting A'
        return 'Sealant'
    return f
USER_GROUP_MAP={'glass':'Glass','glass-g12r':'Glass','eva':'Encapsulant','epe':'Encapsulant','poe':'Encapsulant','jb':'Junction Box',
                'frame':'Frame','frame 1.2':'Frame','busbar':'Busbar','ribbon':'Ribbon','rtv':'Sealant','rtv ml':'Sealant','potting a':'Potting A',
                'potting b':'Potting B','perforated tape':'Tape','cell fixing tape (g-g)':'Tape','label':'Label','connector cap':'Connector Cap',
                'solar cell':'Cell','backsheet':'Backsheet','packaging':'Packaging','frame corner key':'Frame Corner Key'}
def _fam(code,cn_row,bom=None,sapg=None,desc=None):
    """Material family: MIS Standard BOM group ('Understanding Name'), else MIS Common Name Category 2,
       else the SAP material group in MB52, else the code prefix (PMPKM ...)."""
    if bom and bom.get(code) and bom[code]['group']: return bom[code]['group']
    if sapg and not (cn_row and len(cn_row)>5 and cn_row[5]):
        f_=sap_group_family(sapg,desc)
        if f_: return f_
    f=str((cn_row or [None]*6)[5] or '').strip() if cn_row and len(cn_row)>5 else ''
    if not f: m=re.match(r'^(?:JW)?([A-Z]+)',code or ''); f=m.group(1) if m else ''
    return 'Encapsulant' if f.upper() in ENCAPS else f
def _spec(desc):
    """Specs read from a material description."""
    d=' '.join(str(desc or '').upper().replace('–','-').replace('*','X').split())
    sp=dict(d=d)
    m=re.search(r'(\d{3,4}(?:\.\d+)?)\s*X\s*(\d{3,4}(?:\.\d+)?)(?:\s*X\s*(\d+(?:\.\d+)?))?',d)
    if m: sp['dim']=(float(m.group(1)),float(m.group(2))); sp['t']=float(m.group(3)) if m.group(3) else None
    m=re.search(r'(\d{2})\s*HS',d); sp['hs']=int(m.group(1)) if m else None
    m=re.search(r'(\d{2,3})\s*GSM',d); sp['gsm']=int(m.group(1)) if m else None
    w=[int(x) for x in re.findall(r'(?<![\d.])(1[0-4]\d\d)\s*MM',d)] or [int(x) for x in re.findall(r'WIDTH\s*:?\s*(\d{4})',d)]
    sp['width']=w[0] if w else None
    m=re.search(r'(\d{2})\s*A\b',d); sp['amp']=int(m.group(1)) if m else None
    m=re.search(r'(\d{3,4})\s*V\b',d); sp['volt']=int(m.group(1)) if m else None
    sp['excl']=bool(re.search(r'EXC\w*\.?\s*CON|EX\.?\s*CONN|EXCLCON',d)); sp['staubli']='STAUBLI' in d
    sp['model']=next((k for k in ('ZH011C','DSJB','KBT4','VMSJ') if k in d),'JB with connector (no make given)' if 'WITH CONNECTOR' in d else 'make not given')
    sp['kind']=next((k for k,ws_ in (('Edge / corner protector',('PROTECTOR',)),('Wooden pallet',('PALLET',)),('Inner box',('INNERBOX','INNER BOX','CARTON')),
                                      ('Strapping',('STRAP',)),('Stretch film',('WRAPPING','STRETCH')),('Glue sheet',('GLUE',)),('Label',('LABEL',)),('Tape',('TAPE',))) if any(w_ in d for w_ in ws_)),None)
    d2=re.sub(r'\d{3,4}\s*V\b','',d); d2=re.sub(r'ZH011C-5|DSJB12Y|DSJBY|KBT4|VMSJ30','',d2)
    pm=re.search(r'\+\s*(\d{3,4})',d2); nm=re.search(r'-\s*(\d{3,4})(?!\d)',d2)
    nums=[int(x) for x in re.findall(r'(?<![\d.])(\d{3,4})(?![\d.])',d2)]
    if pm and nm: sp['cable']=(int(pm.group(1)),int(nm.group(1)))
    elif nums: sp['cable']=(max(nums),max(nums))
    else: sp['cable']=None
    m=re.search(r'(\d)\s*X\s*(0\.\d+)',d); sp['bb']=(int(m.group(1)),float(m.group(2))) if m else None
    m=re.search(r'(0\.\d+)\s*MM\s*ROUND',d); sp['wire']=float(m.group(1)) if m else None
    m=re.search(r'(\d{1,2})\s*MM\s*X\s*(\d+)\s*M\b',d) or re.search(r'(\d{1,2})\s*MM\s*[X*]\s*(\d+)\s*M\b',d) or re.search(r'TAPE.*?(\d{1,2})\s*MM',d)
    sp['tapew']=int(m.group(1)) if m else None
    sp['front']='FRONT' in d or 'ARC' in d or 'HTAR' in d; sp['back']='BACK' in d or 'HOLE' in d or 'PRINTED' in d
    sp['words']=set(w_ for w_ in re.findall(r'[A-Z]{4,}',d) if w_ not in ('WITH','FROM','SIZE','MATERIAL','PART'))
    return sp

def compare_sub(fam,t,c):
    """Grade candidate c as a replacement for target t (both _spec dicts). Returns (grade, notes) - grade 'A' same spec,
       'B' possible with engineering check, None = not suitable."""
    n=[]; g='A'
    def need(cond,msg):
        nonlocal g
        if not cond: g='B'; n.append(msg)
    f=fam.upper()
    if f=='JUNCTION BOX':
        if t['amp'] and c['amp'] and c['amp']<t['amp']: return None,[]
        if t['amp'] and c['amp'] and c['amp']>t['amp']: n.append(f"higher rating {c['amp']}A vs {t['amp']}A - OK electrically")
        if t['volt'] and c['volt'] and c['volt']<t['volt']: return None,[]
        tc,cc=t.get('cable'),c.get('cable')
        if tc and not cc: g='B'; n.append('cable length not in its description - check it is long enough')
        if tc and cc:
            if cc[0]<tc[0] or cc[1]<tc[1]: return None,[]          # shorter cable cannot reach the next module
            need(cc==tc,f"cable +{cc[0]}/-{cc[1]} mm vs +{tc[0]}/-{tc[1]} mm (longer: check string layout / cable clips)")
        need(c['model']==t['model'],f"different JB model ({c['model']} vs {t['model']}) - must be on the module certificate (UL/IEC)")
        if c['excl'] and not t['excl']: g='B'; n.append('connector NOT included - MC4 connectors must be supplied and crimped')
        if c['staubli']!=t['staubli']: g='B'; n.append('different connector make (Staubli vs other) - check connector compatibility')
    elif f in ('FRAME','GLASS'):
        if not t.get('dim') or not c.get('dim'): return None,[]
        if abs(t['dim'][0]-c['dim'][0])>3 or abs(t['dim'][1]-c['dim'][1])>3: return None,[]
        if f=='GLASS':
            if t['front']!=c['front'] and t['back']!=c['back']: return None,[]
            need(t.get('t')==c.get('t'),f"thickness {c.get('t')} vs {t.get('t')} mm - laminate / frame groove / certificate check")
            if ('HOLE' in t['d'])!=('HOLE' in c['d']): return None,[]
            if ('HTAR' in t['d'])!=('HTAR' in c['d']): n.append('different AR coating (HTAR vs ARC) - power bin check')
        else:
            need(t['hs']==c['hs'],f"frame height {c['hs']} vs {t['hs']} HS - mounting / packing / certificate check")
            gt=re.search(r'(\d+(?:\.\d+)?)\s*MM\s*GROOVE',t['d']); gc=re.search(r'(\d+(?:\.\d+)?)\s*MM\s*GROOVE',c['d'])
            if (gt.group(1) if gt else None)!=(gc.group(1) if gc else None):
                g='B'; n.append(f"glass groove differs ({gc.group(1)+' mm groove' if gc else 'standard G-G groove'} vs {gt.group(1)+' mm groove' if gt else 'standard G-G groove'}) - groove must fit the glass stack thickness")
            if re.search(r'PITCH|\(\d+,\d+',c['d']+t['d']) and c['d']!=t['d']: g='B'; n.append('mounting hole pitch may differ - check drawing')
    elif f=='ENCAPSULANT':
        if t['width'] and c['width'] and abs(t['width']-c['width'])>6: return None,[]
        tt=next((k for k in ENCAPS if k in t['d']),''); ct=next((k for k in ENCAPS if k in c['d']),'')
        need(tt==ct,f"{ct} instead of {tt} - lamination recipe / reliability check")
        tf,cf=('FRONT' in t['d']),('FRONT' in c['d']); tb,cb=('BACK' in t['d']),('BACK' in c['d'])
        if (tf and cb) or (tb and cf): g='B'; n.append(f"{'back' if cb else 'front'} layer material used as {'front' if tf else 'back'} - check")
        need(t['gsm']==c['gsm'],f"{c['gsm']} GSM vs {t['gsm']} GSM - BOM / lamination check")
        if t['width'] and c['width'] and t['width']!=c['width']: g='B'; n.append(f"width {c['width']} vs {t['width']} mm")
    elif f=='BUSBAR':
        if not t['bb'] or not c['bb'] or t['bb'][0]!=c['bb'][0]: return None,[]
        need(t['bb'][1]==c['bb'][1],f"thickness {c['bb'][1]} vs {t['bb'][1]} mm - stringer / soldering setting")
        if ('REFLECTIVE' in t['d'])!=('REFLECTIVE' in c['d']): g='B'; n.append('reflective vs plain busbar - power check')
        if re.search(r'SN\d\dPB\d\d',t['d']+c['d']) and re.search(r'SN\d\dPB\d\d',t['d']) and re.search(r'SN\d\dPB\d\d',c['d']) and re.search(r'SN\d\dPB\d\d',t['d']).group(0)!=re.search(r'SN\d\dPB\d\d',c['d']).group(0):
            g='B'; n.append('different coating alloy - soldering temperature check')
        if re.search(r'(\d+)KG',t['d']+c['d']) and re.search(r'(\d+)KG',t['d'] or '')!=None and re.search(r'(\d+)KG',c['d'] or '')!=None and re.search(r'(\d+)KG',t['d']).group(0)!=re.search(r'(\d+)KG',c['d']).group(0):
            n.append('different spool size - check stringer spool holder')
    elif f=='RIBBON':
        if not t['wire'] or not c['wire'] or abs(t['wire']-c['wire'])>0.021: return None,[]
        need(t['wire']==c['wire'],f"wire {c['wire']} vs {t['wire']} mm - stringer setting / power check")
    elif f=='TAPE':
        if t['tapew'] and c['tapew'] and t['tapew']!=c['tapew']: return None,[]
        if ('PERFORAT' in t['d'])!=('PERFORAT' in c['d']) or ('FIXATION' in t['d'])!=('FIXATION' in c['d']): return None,[]
        if 'HIGH TEMP' in c['d'] and 'HIGH TEMP' not in t['d']: g='B'; n.append('high-temperature tape - check process')
        n.append('roll length may differ - convert quantity (rolls / pieces)')
    elif f in ('POTTING A','POTTING B','SEALANT','POTTING'):
        g='B'; n.append('different maker / formulation - adhesion & cure check; potting A and B must be the SAME brand pair' if 'POTTING' in f else 'different maker / drum size - adhesion check')
    elif f=='CONNECTOR CAP':
        if ('MALE' in t['d'].replace('FEMALE',''))!=('MALE' in c['d'].replace('FEMALE','')): return None,[]
    else:   # packaging, consumables and anything else: same kind of item (shared key words) and same / bigger size
        if not t.get('kind') or t['kind']!=c.get('kind'): return None,[]    # same kind of item only (pallet for pallet, box for box ...)
        if t.get('dim') and c.get('dim'):
            if abs(t['dim'][0]-c['dim'][0])>15 or abs(t['dim'][1]-c['dim'][1])>15:
                if c['dim'][0]<t['dim'][0] or c['dim'][1]<t['dim'][1]: return None,[]
                g='B'; n.append(f"bigger size {c['dim'][0]:.0f}x{c['dim'][1]:.0f} vs {t['dim'][0]:.0f}x{t['dim'][1]:.0f} - check packing / container load")
        elif t.get('dim') or c.get('dim'): g='B'; n.append('size not on both descriptions - check')
        nums_t=set(re.findall(r'\d+',t['d'])); nums_c=set(re.findall(r'\d+',c['d']))
        if nums_t!=nums_c and g=='A':
            sz=lambda d_: (re.search(r'\(([^)]*\d[^)]*)\)',d_) or re.search(r'(\d+(?:[X&]\d+)+)',d_))
            st_,sc_=sz(t['d']),sz(c['d'])
            g='B'; n.append(f"size / spec differs ({sc_.group(1) if sc_ else 'no size given'} vs {st_.group(1) if st_ else 'no size given'}) - check it fits")
    return g,n

def make_substitute_advice_xlsx(A, DS, d0, path, mis_path=None):
    """Substitute Advice: for every planned material, other materials of the same family whose description says they could
       replace it if it runs out - same spec (A) or possible after an engineering check (B) - with their stock, in transit and
       how many extra days they would give at the module plan. Advice only: approve, then add the pair on Settings > Substitutes."""
    from openpyxl.styles import Border, Side
    AR='Calibri'; NAVY='1F3864'; thin=Side(style='thin',color='BFBFBF'); bd=Border(left=thin,right=thin,top=thin,bottom=thin)
    cn=DS.get('cn_all') or {}; stk=DS.get('stock_raw') or {}; desc=DS.get('desc_all') or {}; uom=DS.get('uom_all') or {}
    tr=(DS.get('by_code') or {}).get('transit') or {}; alias=DS.get('alias') or {}
    inA={x['b']['mat']:x for x in A}
    ADVICE_ISSUES.clear()
    bom=read_bom_groups(mis_path) if mis_path else {}
    MX=read_mis_extra(mis_path,DS.get('plant')) if mis_path else dict(jbtype={},makes={},replace={},sites={},booking={},mbq={})
    sapg={}; sapt={}
    if mis_path:
        try:
            for r_ in list(wb_ro(mis_path)['MB52'].iter_rows(values_only=True))[1:]:
                if r_ and len(r_)>11 and r_[2]:
                    c_=str(r_[2]).strip(); sapg.setdefault(c_,str(r_[10] or '').strip()); sapt.setdefault(c_,str(r_[11] or '').strip())
                    if r_[5]: desc.setdefault(c_,str(r_[5]))
        except Exception as _e_: _swallow(2505,_e_); pass
    bsup=batch_suppliers(mis_path,DS.get('plant'),MX['makes']) if mis_path else {}
    def sup_txt(c_,name=None):
        # batch suppliers + MIS makes + stock status + other sites + booking, for one code
        out=[]; cnt=bsup.get(c_)
        bt_='; '.join(f"{sup_name(k)}: {q:,.0f}" for k,q in (cnt.most_common() if cnt else []) if q)
        if bt_: out.append('Batches: '+bt_)
        q=MX['mbq'].get(c_)
        if q and (q[1] or q[2]): out.append(f"unrestricted {q[0]:,.0f} / quality insp. {q[1]:,.0f} / blocked {q[2]:,.0f}")
        if MX['makes'].get(c_): out.append('Approved makes (MIS): '+', '.join(dict.fromkeys(MX['makes'][c_])))
        if MX['sites'].get(c_): out.append('Other sites: '+', '.join(f"{k} {v:,.0f}" for k,v in MX['sites'][c_].items()))
        bk=MX['booking'].get(str(name or '').strip().lower())
        if bk and (bk[0] or bk[1]): out.append(f"MIS booking: pending SO {bk[0]:,.0f}, approval given {bk[1]:,.0f}"+(f" ({bk[2]})" if bk[2] else ''))
        return ' | '.join(out)[:400] or None
    # which JB does the MIS plan say each module uses (Total Plan Req 'JB Type', e.g. 500/30A) -> JB codes with that Common Name
    jb_by_type={}
    for c_,r_ in cn.items():
        nm_=str((r_[3] if len(r_)>3 else '') or '').strip().upper()
        m_=re.match(r'^JB\s+(.+)$',nm_)
        if m_: jb_by_type.setdefault(m_.group(1).replace(' ',''),set()).add(c_)
    pool={}
    for c_ in set(cn)|set(stk)|set(inA)|set(bom):
        c_=str(c_).strip()
        if not c_ or c_.upper().startswith('JW') or c_ in alias or c_ in CSPM_ALT: continue     # CS / PM '_Alt' = already in its standard code's stock
        row=cn.get(c_); fam=_fam(c_,row,bom,sapg.get(c_),desc.get(c_))
        if not fam or fam.upper()=='CELL': continue
        x=inA.get(c_); b=x['b'] if x else None
        d_=(b.get('desc') if b else None) or desc.get(c_) or (row[1] if row else '') or ''
        st_=(b['stock'] if b else stk.get(c_,0.0)) or 0.0; it_=(x.get('tr_tot') if x else sum(tr.get(c_) or [])) or 0.0
        sp_=_spec(d_); bi=bom.get(c_)
        if bi and bi['role']:              # front / back from the MIS BOM item order beats the words of the description
            sp_['front']='front' in bi['role']; sp_['back']='back' in bi['role'] and 'front' not in bi['role']
        pool[c_]=dict(code=c_,fam=fam,desc=d_,bom=bi,name=(b['common'] if b else (row[3] if row and row[3] else d_)),uom=(b.get('uom') if b else uom.get(c_)),
                      stock=st_,it=it_,x=x,sp=sp_)
    for p_ in pool.values():                 # what is free for another material: stock + in transit - its own target days of demand
        x=p_['x']
        if x:
            own=(x.get('pl_need') or x.get('avg') or 0)*(x['b'].get('td') or tgt_days(x['b'].get('lt_used')))
            p_['free']=max(0.0,p_['stock']+p_['it']-own); p_['use']=f"Used by plan - {x['status']}, lasts {x['pl_days']:.0f} d at the plan" if x.get('pl_days') is not None else f"Used by plan - {x['status']}"
        else: p_['free']=max(0.0,p_['stock']+p_['it']); p_['use']='Not in the plan - free stock'
    # YOUR SAP BOM (MRP.xlsx) vs MIS Standard BOM: same customer (Settings MIS link -> MIS BOM name), same material group
    mis_bn_fam={}
    try:
        for r_ in list(wb_ro(mis_path)['Standard BOM'].iter_rows(values_only=True))[1:]:
            if r_ and len(r_)>3 and r_[2] and r_[3]:
                c2=str(r_[3]).strip(); f2=(pool.get(c2) or {}).get('fam') or _fam(c2,cn.get(c2),bom,sapg.get(c2),desc.get(c2))
                mis_bn_fam.setdefault((str(r_[2]).strip(),f2),set()).add(c2)
    except Exception as _e_: _swallow(2552,_e_); pass
    sap_alt={}            # MIS code -> [(your code, customer, model, valid from)] where your SAP BOM uses another code of the same group
    lk_={_norm(k):v for k,v in LINKS.items()}
    sap_rows=[x_ for x_ in read_sap_bom(SETTINGS_SRC[0]) if x_['grp']=='RM']
    by_cf={}
    for x_ in sap_rows:
        f2=(pool.get(x_['code']) or {}).get('fam') or _fam(x_['code'],cn.get(x_['code']),bom,sapg.get(x_['code']),x_['desc'])
        by_cf.setdefault((x_['cust'],f2),[]).append(x_)
    for (cu_,f2),xs in sorted(by_cf.items()):
        if not f2 or f2=='Cell': continue
        L_=lk_.get(_norm(cu_)) or {}; bn_=L_.get('bom')
        mis_c=mis_bn_fam.get((str(bn_ or '').strip(),f2),set()); sap_c={x_['code'] for x_ in xs}
        al_=DS.get('alias') or {}
        if not bn_ or not mis_c or {al_.get(c3,c3) for c3 in sap_c}=={al_.get(c3,c3) for c3 in mis_c}: continue   # alternate codes / roll lengths of the same material = same
        for mc in mis_c-sap_c:
            for x_ in xs:
                if x_['code'] not in mis_c: sap_alt.setdefault(mc,[]).append((x_['code'],cu_,x_['model'],x_['valid']))
        ADVICE_ISSUES.append(dict(area='Master data',check='Your BOM (MRP.xlsx) and the MIS BOM use a different raw material',sev='High',material=f"{cu_} - {f2}",code=', '.join(sorted(sap_c)),
            a_lab='Your BOM (MRP.xlsx)',a=', '.join(sorted(sap_c)),b_lab=f'MIS BOM {bn_}',b=', '.join(sorted(mis_c)),diff=None,
            detail=f"{cu_} ({xs[0]['model']}, valid from {xs[0]['valid']}): your BOM has {', '.join(sorted(sap_c))}, MIS Standard BOM {bn_} has {', '.join(sorted(mis_c))}. "
                   "Raw materials are planned with the MIS BOM (your BOM is used for CS / PM), so the MRP plans the MIS code - this row is a cross-check.",
            action='Confirm with MIS / engineering which code the line really uses; correct the MIS BOM (or yours), or add the pair on Settings > Substitutes for this line'))
    wb=Workbook(); ws=wb.active; ws.title='Substitute Advice'; ws.sheet_view.showGridLines=False
    ws['A1']='Substitute Advice - what could replace a material if it runs out (read from material descriptions)'; ws['A1'].font=Font(name=AR,size=14,bold=True,color=NAVY)
    ws['A2']=('ADVICE ONLY - nothing here changes the MRP. Candidates are materials of the same family (MIS Common Name category, or code group for packaging) whose description '
              'matches: A = same spec (only the maker / pack size / rating-above differs) - confirm it is on the module certificate and BOM; B = possible, but the differences listed '
              'need engineering / quality approval. Not listed = not suitable (e.g. lower amp JB, shorter cable, other glass / frame size). Free qty = stock + in transit minus what '
              'the plan needs of the candidate itself for its target days. Extra days = free qty / Usage/Day of the material at the module plan (blank when the units differ - convert first). '
              'Groups come from the MIS Standard BOM (Understanding Name + item no.; front / back = item order in each module type), then MIS Common Name, then the SAP material group in MB52 (W001 glass, W002 EVA/EPE, W008 frame, W009 JB ...). '
              'Cross-check: YOUR BOM (MRP.xlsx BOM sheet, used for CS / PM) uses the candidate where the MIS BOM has this material (also listed on Check). RM planning itself stays on the MIS BOM. '
              'Supplier = 2nd-3rd character of the SAP batch (MB52) - name the codes on Settings > Batch supplier codes. Also used from MIS: Total Plan Req \'JB Type\' (the JB the plan says a module uses), '
              'RM Status \'Make\' and \'Replace Material Code\', stock at Arizona / JW, Booking Approval pending SO, MB52 unrestricted / QI / blocked. '
              'When approved, add the pair on Settings > Substitutes (MRP material, substitute, lines) so the MRP uses it.')
    ws['A2'].font=Font(name=AR,size=9,italic=True); ws.merge_cells('A2:S2'); ws['A2'].alignment=Alignment(wrap_text=True,vertical='top'); ws.row_dimensions[2].height=64
    H=['Material that may run out','Code','Status','Days of cover (+ all in transit)','Usage/Day (module plan)','Description',
       'Option','Substitute code','Substitute description','Match','Differences / what to check','Substitute stock','Substitute in transit','Substitute is used by','Free qty for this material','Extra days at line targets',
       'MIS BOM group of substitute (module types using it)','Substitute: suppliers (batch), stock status, MIS makes, other sites, booking','This material: suppliers (batch), stock status, MIS makes, other sites, booking']
    for j,h in enumerate(H,1):
        c=ws.cell(4,j,h); c.font=Font(name=AR,size=9,bold=True,color='FFFFFF'); c.fill=PatternFill('solid',fgColor=NAVY); c.border=bd; c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
    ws.row_dimensions[4].height=42
    r=5; fA=PatternFill('solid',fgColor='C6EFCE'); fB=PatternFill('solid',fgColor='FFEB9C')
    order_=sorted(A,key=lambda x:rrank(x['b']['mat']))                  # highest shortage risk first
    nmat=0; nwith=0
    for x in order_:
        b=x['b']; t=pool.get(b['mat']) or dict(fam=_fam(b['mat'],cn.get(b['mat']),bom,sapg.get(b['mat']),b.get('desc')),sp=_spec(b.get('desc') or desc.get(b['mat'])),bom=bom.get(b['mat']))
        need_=x.get('pl_need') or x.get('avg') or 0
        plan_bn=set(); plan_jbt=set(); plan_jb=set()
        if t['fam']=='Junction Box' and t.get('bom'):
            plan_bn={bn_ for bn_ in t['bom']['mts'] if bn_ in MX['jbtype']}
            plan_jbt={jt for bn_ in plan_bn for jt in MX['jbtype'][bn_]}
            plan_jb={c2 for jt in plan_jbt for c2 in jb_by_type.get(jt.upper().replace(' ',''),set())}
            if plan_jb and b['mat'] not in plan_jb:
                ADVICE_ISSUES.append(dict(area='Master data',check="MIS BOM and MIS plan disagree on the junction box",sev='High',material=b['common'],code=b['mat'],
                    a_lab='MIS Standard BOM',a=b['mat'],b_lab="Total Plan Req 'JB Type'",b='/'.join(sorted(plan_jbt)),diff=None,
                    detail=f"Standard BOM gives {b['mat']} for {', '.join(sorted(plan_bn))}, but Total Plan Req 'JB Type' for the same module says {'/'.join(sorted(plan_jbt))} ({', '.join(sorted(plan_jb))})",
                    action='Ask MIS / engineering which JB is right and correct the BOM (or add it on Settings > Substitutes)'))
        cands=[]
        for c_,p_ in pool.items():
            if c_==b['mat'] or (p_['fam']!=t['fam']): continue
            g,notes=compare_sub(t['fam'],t['sp'],p_['sp'])
            if not g: continue
            if (p_['uom'] or '')!=(b.get('uom') or '') and p_['uom'] and b.get('uom'): notes=notes+[f"unit {p_['uom']} vs {b.get('uom')} - convert quantity"]
            tb,cb=t.get('bom'),p_.get('bom')
            if t['fam'] in ('Glass','Encapsulant') and tb and cb and tb['role'] and cb['role'] and not (tb['role']&cb['role']):
                if t['fam']=='Glass': continue                     # front glass never replaces back glass (and the other way)
                g='B'; notes=notes+[f"MIS BOM uses it as {'/'.join(sorted(cb['role']))} layer, this material is {'/'.join(sorted(tb['role']))}"]
            if c_ in MX['replace'].get(b['mat'],set()): g='A'; notes=[f"MIS RM Status 'Replace Material Code' links these two codes"]+notes
            if c_ in (ALT_BOM.get(b['mat']) or {}): g='A'; notes=[f"YOUR ALTERNATIVE BOM (PPC BOM code not in the MIS Standard BOM) for {', '.join(sorted(ALT_BOM[b['mat']][c_]))}"]+notes
            sa_=[z_ for z_ in sap_alt.get(b['mat'],[]) if z_[0]==c_]
            if sa_: g='A'; notes=[f"YOUR SAP BOM (MRP.xlsx, CS/PM BOM) uses this code for {', '.join(dict.fromkeys(z_[1] for z_ in sa_))} ({sa_[0][2]}, valid from {sa_[0][3]}) where the MIS BOM has {b['mat']}"]+notes
            if plan_jb and c_ in plan_jb: g='A'; notes=[f"MIS Total Plan Req 'JB Type' says the module(s) {', '.join(sorted(plan_bn))} use JB {'/'.join(sorted(plan_jbt))} = this code"]+notes
            ts_={k for k,q in (bsup.get(b['mat']) or {}).items() if q and k!='?'}; cs_={k for k,q in (bsup.get(c_) or {}).items() if q and k!='?'}
            if ts_ and cs_ and ts_&cs_: notes=notes+[f"same supplier as this material ({', '.join(sorted(ts_&cs_))})"]
            cands.append((g,p_,notes))
        cands.sort(key=lambda z:(z[0],not any(n_.startswith('YOUR SAP BOM') for n_ in z[2]),not any(n_.startswith('MIS ') for n_ in z[2]),-(z[1]['free']),str(z[1]['code'])))   # code last: same order on every run
        nmat+=1; withstock=[z for z in cands if z[1]['stock']+z[1]['it']>0]
        if withstock: nwith+=1
        head=[b['common'],b['mat'],x['status'],(round(x['pl_days'],1) if x.get('pl_days') is not None else None),(round(need_,1) if need_ else None),b.get('desc') or desc.get(b['mat'])]
        show=withstock[:6] or [None]
        for k,z in enumerate(show):
            if z is None:
                why=(f"{len(cands)} same-family material(s) in master data could fit, but none has stock or in transit: "+', '.join(f"{p_['code']} ({g})" for g,p_,_ in cands[:5])) if cands else 'No suitable substitute found in the master data (same family + matching description)'
                vals=head+[None,None,None,None,why,None,None,None,None,None,None,None,sup_txt(b['mat'],b['common'])]
            else:
                g,p_,notes=z
                vals=(head if k==0 else [None,None,None,None,None,None])+[k+1,p_['code'],p_['desc'],{'A':'A - same spec','B':'B - check'}[g],'; '.join(notes) or 'same spec',
                      round(p_['stock']),round(p_['it']) or None,p_['use'],round(p_['free']),(round(p_['free']/need_,1) if need_ and not (p_['uom'] and b.get('uom') and p_['uom']!=b.get('uom')) else None),
                      ((f"{p_['bom']['group']} (item {'/'.join(str(int(s_)) for s_ in sorted(p_['bom']['sr']) if isinstance(s_,(int,float)))}) - {len(p_['bom']['mts'])} module type(s): "
                        +', '.join(sorted(p_['bom']['mts'])[:3])+(' ...' if len(p_['bom']['mts'])>3 else '')) if p_.get('bom') else 'not in MIS BOM')
                       +(f" | SAP group {sapg[p_['code']]} ({sapt.get(p_['code']) or '-'})" if sapg.get(p_['code']) else ''),
                      sup_txt(p_['code'],p_['name']),(sup_txt(b['mat'],b['common']) if k==0 else None)]
            for j,v in enumerate(vals,1):
                c=ws.cell(r,j,v); c.font=Font(name=AR,size=9,bold=(j==1)); c.border=bd; c.alignment=Alignment(vertical='top',wrap_text=j in (1,6,9,11,14,17,18,19))
                if j in (12,13,15): c.number_format='#,##0'
                if j in (4,5,16): c.number_format='#,##0.0'
            if z is not None: ws.cell(r,10).fill=fA if z[0]=='A' else fB
            if k==0:
                ws.cell(r,3).fill=PatternFill('solid',fgColor=STATUS_FILL[x['status']])
                dv=ws.cell(r,4).value
                if isinstance(dv,(int,float)): ws.cell(r,4).fill=PatternFill('solid',fgColor='FFC7CE' if dv<tgt_days(b.get('lt_used'))*2/3 else ('FFEB9C' if dv<tgt_days(b.get('lt_used')) else 'C6EFCE'))
                for j in range(1,len(H)+1): ws.cell(r,j).border=Border(left=thin,right=thin,bottom=thin,top=Side(style='medium',color=NAVY))
            r+=1
    ws.auto_filter.ref=f'A4:{L(len(H))}{max(5,r-1)}'
    for j,w in enumerate([24,15,10,11,10,30,6,15,34,11,48,11,11,26,11,10,40,48,48],1): ws.column_dimensions[L(j)].width=w
    set_print(ws,'Substitute Advice'); wb.save(path)
    return nmat,nwith

# ===================== In Transit Status: MIS 'In Transit' STATUS column analysed =====================
IT_GROUPS=['In transit','At origin','At port (POD)','Held (customs / issue)','Partly received - rest pending / in transit','Transloaded at warehouse','Delivered','Other','No status']
def it_status_group(t):
    """Group the free-text MIS STATUS into a few statuses (the exact MIS text stays on the detail rows)."""
    u=' '.join(str(t or '').replace('\xa0',' ').split()).lower()
    if not u: return 'No status'
    if any(w in u for w in ('stuck','stucked','inspection','hold')): return 'Held (customs / issue)'
    if u in ('in-transit','in transit','intransit','in - transit'): return 'In transit'
    if 'origin' in u: return 'At origin'
    if re.search(r'\bpod\b',u) or 'at port' in u: return 'At port (POD)'
    if 'pending' in u or 'in transit' in u or 'in-transit' in u: return 'Partly received - rest pending / in transit'
    if 'transload' in u: return 'Transloaded at warehouse'
    if 'deliver' in u or 'completed' in u: return 'Delivered'
    return 'Other'

def make_intransit_status_xlsx(mis_path, DS, path):
    """In Transit Status sheet: how much is in which MIS status - by status, by material x status, and every shipment line."""
    from openpyxl.styles import Border, Side
    AR='Calibri'; NAVY='1F3864'; thin=Side(style='thin',color='BFBFBF'); bd=Border(left=thin,right=thin,top=thin,bottom=thin)
    hf=Font(name=AR,size=9,bold=True,color='FFFFFF'); hfill=PatternFill('solid',fgColor=NAVY); bf=Font(name=AR,size=9); bold=Font(name=AR,size=9,bold=True)
    alias=DS.get('alias') or {}; fac=DS.get('alias_factor') or {}
    info={b['mat']:b for b in DS['blocks']}
    rows=list(wb_ro(mis_path)['In Transit'].iter_rows(values_only=True))[1:]
    n_=lambda v: float(v) if isinstance(v,(int,float)) and not isinstance(v,bool) else 0.0
    d_=lambda v: v if isinstance(v,(dt.datetime,dt.date)) else None
    recs=[]
    for r in rows:
        if len(r)<40: continue
        code0=str(r[39] or r[31] or '').strip()
        if not code0 and not r[1]: continue
        if not CODE.match(code0) and not re.match(r'^[A-Z]{2,}[A-Z0-9]*\d{3,}$',code0): code0=''      # e.g. 'No Material Found'
        code=alias.get(code0,code0); b=info.get(code)
        recs.append(dict(grp=it_status_group(r[22]),st=' '.join(str(r[22] or '').replace('\xa0',' ').split()),code=code,code0=code0,
                         mat=(b['common'] if b else ' '.join(str((r[36] if code0 else None) or r[4] or code0 or '(no RM code in MIS)').split())[:60]),uom=(b.get('uom') if b else ''),
                         sup=' '.join(str(r[1] or '').split()),inv=r[2],po=r[6],bl=r[23] or r[24],vessel=' '.join(str(r[12] or '').split()),
                         cont=n_(r[21]),etd=d_(r[26]),atd=d_(r[27]),eta=d_(r[29]) or d_(r[28]),ata=d_(r[35]),
                         qty=n_(r[32]),grn=n_(r[33]),pend=n_(r[34])*fac.get(code0,1.0),rate=n_(r[8]),wh=r[30],planned=b is not None))
    wb=Workbook(); ws=wb.active; ws.title='In Transit Status'; ws.sheet_view.showGridLines=False
    ws['A1']='In Transit Status - how much is in which status (PPC In_Transit, STATUS column)'; ws['A1'].font=Font(name=AR,size=14,bold=True,color=NAVY)
    ws['A2']=("Pending qty = PPC In_Transit 'Final Pending Qty.' (not yet received at the plant; alternate codes converted to the MRP unit). The free-text STATUS is grouped into a few statuses; "
              "the exact MIS text is on every detail row (column B). Shipped qty = 'Final Qty.', GRN = 'PPC GRN'. Value = unit rate x pending qty (invoice currency).")
    ws['A2'].font=Font(name=AR,size=9,italic=True); ws.merge_cells('A2:N2'); ws['A2'].alignment=Alignment(wrap_text=True); ws.row_dimensions[2].height=30
    def head(r,cols,widths=None):
        for j,t in enumerate(cols,1): c=ws.cell(r,j,t); c.font=hf; c.fill=hfill; c.border=bd; c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
        ws.row_dimensions[r].height=30
    grps=[g for g in IT_GROUPS if any(x['grp']==g for x in recs)]
    r=4
    # ---- 2) pending qty by material x status ----
    ws.cell(r,1,'1. PENDING QTY BY MATERIAL AND STATUS').font=Font(name=AR,size=12,bold=True,color=NAVY); r+=1
    pg=[g for g in grps if any(x['grp']==g and x['pend']>0 for x in recs)]
    head(r,['Material','Code','UOM']+pg+['Total pending','First arrival at plant (pending)']); r+=1
    mats={}
    for x in recs:
        if x['pend']<=0: continue
        m=mats.setdefault(x['code'] or x['mat'],dict(mat=x['mat'],code=x['code'],uom=x['uom'],q={},first=None))
        m['q'][x['grp']]=m['q'].get(x['grp'],0.0)+x['pend']
        if x['ata'] and (m['first'] is None or x['ata']<m['first']): m['first']=x['ata']
    for k,m in sorted(mats.items(),key=lambda kv:(rrank(kv[1]['code']),str(kv[1]['mat']).lower())):     # highest shortage risk first
        vals=[m['mat'],m['code'],m['uom']]+[m['q'].get(g) or None for g in pg]+[sum(m['q'].values()),m['first']]
        for j,v in enumerate(vals,1):
            c=ws.cell(r,j,v); c.font=bold if j==len(vals)-1 else bf; c.border=bd
            if 4<=j<=len(vals)-1: c.number_format='#,##0'
            if j==len(vals): c.number_format='dd-mmm-yy'
        r+=1
    r+=2
    # ---- 3) detail (filter here) ----
    ws.cell(r,1,'2. EVERY SHIPMENT LINE (filter by Status / Material)').font=Font(name=AR,size=12,bold=True,color=NAVY); r+=1
    H=['Status','MIS status (exact text)','Material','Code','UOM','Supplier','Invoice','PO','BL','Vessel / voyage','Containers','ETD','ATD','ETA port (revised)',
       'Arrival at plant (Final AtA)','Shipped qty','GRN qty','Pending qty','Warehouse','In MRP']
    head(r,H); h0=r; r+=1
    order_={g:i for i,g in enumerate(IT_GROUPS)}
    for x in sorted(recs,key=lambda x:(x['pend']<=0,rrank(x['code']),order_.get(x['grp'],99),x['ata'] or dt.datetime(2100,1,1))):
        vals=[x['grp'],x['st'] or None,x['mat'],x['code0'],x['uom'],x['sup'],x['inv'],x['po'],x['bl'],x['vessel'] or None,x['cont'] or None,x['etd'],x['atd'],x['eta'],x['ata'],
              x['qty'] or None,x['grn'] or None,x['pend'] or None,x['wh'],'Yes' if x['planned'] else 'No']
        for j,v in enumerate(vals,1):
            c=ws.cell(r,j,v); c.font=bf if x['pend']>0 else Font(name=AR,size=9,color='808080'); c.border=bd; c.alignment=Alignment(vertical='top',wrap_text=j in (2,3))
            if j in (12,13,14,15): c.number_format='dd-mmm-yy'
            if j in (11,16,17,18): c.number_format='#,##0'
        if x['pend']>0: ws.cell(r,18).font=bold
        r+=1
    if r-1>h0: ws.auto_filter.ref=f'A{h0}:{L(len(H))}{r-1}'
    for j,w in enumerate([24,34,26,16,6,26,13,12,16,20,9,10,10,11,12,11,11,11,16,7],1): ws.column_dimensions[L(j)].width=w
    ws.freeze_panes='B4'; set_print(ws,'In Transit Status'); wb.save(path)
    return len(recs)

