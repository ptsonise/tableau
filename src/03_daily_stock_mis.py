# ===================== Daily Stock sheet (from the MIS RM Dashboard) =====================
# Plan scope: every MIS Total Plan Req row of the plant (Company = USA). Set PLAN_WSA_LINES_ONLY=True to count only
# production lines USA* with OEM Category WSA (OEM lines Imperial Star, Runergy, Inox and Arizona excluded).
PLAN_OEM='WSA'
def plan_line_ok(line, oem, company, plant='USA'):
    if str(company or '').strip().upper()!=plant.upper(): return False
    if not PLAN_WSA_LINES_ONLY: return True
    ln=str(line or '').strip().upper()
    return str(oem or '').strip().upper()==PLAN_OEM and (ln.startswith('USA') or re.fullmatch(r'L\d+',ln) is not None)
def tp_ok(r, plant='USA'):
    """A row of MIS Total Plan Req counts in the plan (columns: 1 Plant / line, 18 OEM Category, 21 Company)."""
    return len(r)>21 and plan_line_ok(r[1],r[18],r[21],plant)

TAPE_TYPES=('cell fixation','perforated','heating','double side','kapton','insulation')

def tape_info(desc):
    """Tape family and metres per roll read from the description, e.g. 'G2G PERFORATED TAPE 30MM X 500M' -> ('perforated',30,500),
       'CELL FIXATION TAPE-8MM*100M' -> ('cell fixation',8,100), 'CELL FIXATION TAPE 8MM FF-3665 (T50)' -> ('cell fixation',8,50)."""
    t=' '.join(str(desc or '').upper().split())
    typ=next((k for k in TAPE_TYPES if k.upper() in t),None)
    if not typ or 'TAPE' not in t: return None
    w=re.search(r'(\d+(?:\.\d+)?)\s*MM',t)
    m=re.search(r'(?<![\d.])(\d+(?:\.\d+)?)\s*M(?![A-Z])',t) or re.search(r'\bT(\d+)\b',t)
    if not (w and m): return None
    return (typ,float(w.group(1)),float(m.group(1)))


# ===================== MIS without 'Total Plan Req' / 'Daily Dashboard': rebuilt in memory =====================
CODE_SWAP_F={}     # your BOM code -> factor (1 roll of the MIS code = f rolls of yours), for the planner's typed WSA quantities
def _SWP_INV(): return {o_:(n_,CODE_SWAP_F.get(n_,1.0)) for n_,o_ in CODE_SWAP.items()}   # MIS code -> (your BOM code, factor)
def _bom_code_of(code, mis_path=None):
    """Your BOM code of a MIS tape code (CODE_SWAP rule, usable before prepare_mis): the code itself when it is in your BOM or is no tape;
       else the ONE code of your BOM with the same tape type + width (RMTTT00000010 -> RMTTT00000018)."""
    c=str(code or '').strip()
    if c in _SWP_INV(): return _SWP_INV()[c][0]
    try:
        if '_bc' not in _BOMC:
            _BOMC['_bc']=load_user_bom_rm(SETTINGS_SRC[0]) if SETTINGS_SRC and SETTINGS_SRC[0] else {}
            _BOMC['_d']={}
            if mis_path:
                wb_=wb_ro(mis_path)
                if 'Common Name' in wb_.sheetnames:
                    for r_ in wb_['Common Name'].iter_rows(min_row=2,values_only=True):
                        if r_ and r_[0] and len(r_)>1 and r_[1]: _BOMC['_d'].setdefault(str(r_[0]).strip(),str(r_[1]))
        bc=_BOMC['_bc']; ti=tape_info(_BOMC['_d'].get(c))
        if not bc or c in bc or not ti: return c
        hit=[k for k,m in bc.items() if (tape_info(m.get('desc')) or ())[:2]==ti[:2]]
        return hit[0] if len(hit)==1 else c
    except Exception as _e_: _swallow(1083,_e_); return c
_BOMC={}
CODE_SWAP={}       # your BOM code -> MIS dashboard code it replaces as the material (prepare_mis)
def swap_name(b):
    """Item name of a material planned under YOUR BOM code (CODE_SWAP) when the MIS Common Name names the MIS code's roll length
       (user 20-Nov-2026: 'Cell Fixation Tap-T50' next to BOM 'CELL FIXATION TAPE-8MM*100M' - MIS gives 010 (50 m) and 018 (100 m) the
       same Common Name): the MIS roll length in the name -> your BOM roll length ('Cell Fixation Tap-100m'). None = no change."""
    if b.get('mat') not in CODE_SWAP: return None
    ti_=tape_info(b.get('desc')); f_=CODE_SWAP_F.get(b['mat'])
    if not ti_ or not f_: return None
    ln_=float(ti_[2]); lo_=ln_*float(f_); nm_=str(b.get('common') or '')
    g_=lambda v: f"{v:g}"
    if abs(lo_-ln_)<1e-9: return None
    new_=re.sub(r'\bT'+re.escape(g_(lo_))+r'\b|\b'+re.escape(g_(lo_))+r'\s*m\b',g_(ln_)+'m',nm_,flags=re.I)
    return new_ if new_!=nm_ else (nm_+f" {g_(ln_)}m" if nm_ else None)
MIS_REBUILT=[]     # which MIS sheets were rebuilt this run (shown on the console and on Check)
def _prev_links(prev_path):
    """Customer -> (MIS customer, MIS BOM Name) from the MIS LINK table on the Settings sheet of the previous MRP_Daily."""
    out={}
    if not prev_path or not os.path.exists(prev_path): return out
    try:
        on=False
        for r in wb_ro(prev_path)['Settings'].iter_rows(values_only=True):
            a=str((r or [None])[0] or '').strip()
            if a.upper().startswith('MIS LINK'): on=True; continue
            if on and a and a.isupper() and len(a)>8: break
            if on and a and not a.lower().startswith('production plan customer') and len(r)>2:
                bn=(r[4] if len(r)>4 and r[4] not in (None,'') else r[2]); mc=(r[5] if len(r)>5 and r[5] not in (None,'') else r[1])
                if bn: out[_norm(a)]=(str(mc).strip() if mc else a,str(bn).strip())
    except Exception as _e_: _swallow(1015,_e_); pass
    return out

def prepare_mis(mis_path, mrp_path, prev_path, ndays=141):
    """Your MIS file may have no 'Total Plan Req' and / or no 'Daily Dashboard'. Then they are rebuilt in memory (the file is not changed):
       Total Plan Req = your Production Plan (MRP.xlsx) modules x MIS Standard BOM quantity per module (layers listed twice add up);
       Daily Dashboard = one material per code of the Standard BOM of the module types in your plan (cells left out), named from Common Name.
       Customer -> MIS BOM Name: Settings MIS LINK override, else the previous MRP_Daily's MIS LINK table, else the Module Model of a linked customer."""
    MIS_REBUILT.clear()
    w=wb_ro(mis_path); need_tp='Total Plan Req' not in w.sheetnames; need_dd='Daily Dashboard' not in w.sheetnames
    if not (need_tp or need_dd): return
    if not mrp_path: die("The MIS file has no 'Total Plan Req' / 'Daily Dashboard' - then MRP.xlsx (Production Plan) is needed to build the raw-material plan")
    # plant = MB52 Company of the plant rows
    import collections
    co=collections.Counter(str(r[17]).strip() for r in w['MB52'].iter_rows(min_row=2,values_only=True) if r and len(r)>17 and r[17] and 'WAREHOUSE' not in str(r[17]).upper())
    plant=co.most_common(1)[0][0] if co else 'USA'
    # Standard BOM: module type -> [(sr, code, desc, qty, uom, group)]
    sb={}
    for r in w['Standard BOM'].iter_rows(min_row=2,values_only=True):
        if r and len(r)>8 and r[2] and r[3] and isinstance(r[5],(int,float)):
            sb.setdefault(str(r[2]).strip(),[]).append((r[1],str(r[3]).strip(),r[4],float(r[5]),r[6],str(r[8] or '').strip()))
    cn={}
    for r in w['Common Name'].iter_rows(min_row=2,values_only=True):
        if r and r[0]: cn.setdefault(str(r[0]).strip(),r)
    # your Production Plan
    d0=base_d0(); dend=pd.Timestamp(dt.date.today()).normalize()+pd.Timedelta(days=ndays)   # base date earlier: window starts earlier, same end
    pp=[]; models={}
    for r in pp_rows_std(wb_ro(mrp_path)['Production Plan']):
        if not r or not isinstance(r[0],(dt.datetime,dt.date)) or not isinstance(r[4],(int,float)) or not r[4]: continue
        d=pd.Timestamp(r[0]).normalize()
        if d<d0-pd.Timedelta(days=3) or d>=dend: continue
        ln=' '.join(str(r[1] or '').split()); cu=' '.join(str(r[3] or '').split()); mdl=str(r[7] or '').strip() if len(r)>7 else ''
        pp.append((d,ln,cu,float(r[4]),mdl)); models.setdefault(cu,collections.Counter())[mdl]+=r[4]
    prev=_prev_links(prev_path); link={}
    for cu in models:
        ov=link_override(cu)
        if ov and ov[1]: link[cu]=(ov[0] or cu,ov[1]); continue
        if _norm(cu) in prev: link[cu]=prev[_norm(cu)]
    for cu in models:                                # same Module Model as a linked customer
        if cu in link: continue
        mdl=models[cu].most_common(1)[0][0]
        for c2,(mc2,bn2) in list(link.items()):
            if mdl and mdl in models.get(c2,{}): link[cu]=(cu,bn2); break
    missing_=[cu for cu in models if cu not in link or link[cu][1] not in sb]
    if models and len(missing_)==len(models):
        die("MIS has no 'Total Plan Req', so each Production Plan customer needs its MIS BOM Name. None was found (no Settings MIS LINK, no previous MRP_Daily).\n"
            "       Add the MIS LINK table to the Settings sheet of MRP.xlsx (Production Plan customer | MIS customer | MIS BOM Name | How | Override: MIS BOM Name | Override: MIS customer).")
    if need_tp:
        rows=[tuple(['Date','Plant','Customer','Value','BOM Name','JB Type','Watt','Identify Name','Total MW','Standard BOM.Sr. no.','Standard BOM.Material Code',
                     'Standard BOM.Material Description','Standard BOM.BOM','Standard BOM.UOM','Total Requirement','Common Name.Common Name','Month-Year','OA No.',
                     'OEM Category','Match','Dashboard Key','Company','Final Week','Quarter','Match2','Comman Name Cat.1','Comman Name Cat.2'])]
        for d,ln,cu,q,mdl in pp:
            if cu not in link or link[cu][1] not in sb: continue
            mc,bn=link[cu]
            for sr,code,desc,bq,uom,grp in sb[bn]:
                c_=cn.get(code) or [None]*7
                rows.append((d.to_pydatetime(),ln,mc,q,bn,None,None,None,None,sr,code,desc,bq,uom,q*bq,c_[3] if len(c_)>3 else None,None,f'{ln}|{cu}',
                             PLAN_OEM,code,None,plant,None,None,code,c_[4] if len(c_)>4 else None,c_[5] if len(c_)>5 else None))
        w._s['Total Plan Req']=_CachedSheet(rows,'Total Plan Req'); w.sheetnames.append('Total Plan Req')
        MIS_REBUILT.append(f"Total Plan Req ({len(rows)-1} rows = your Production Plan x MIS Standard BOM)")
    if need_dd:
        used=[]; seen_codes=set(); seen_names=set()
        for bn in sorted({link[cu][1] for cu in link if link[cu][1] in sb},key=lambda b_:-sum(q for d,ln,cu,q,m in pp if link.get(cu,(None,None))[1]==b_)):
            for sr,code,desc,bq,uom,grp in sb[bn]:
                if grp.lower() in ('cells','cell') or code in seen_codes: continue
                c_=cn.get(code) or [None]*7; nm=str(c_[3] or desc or code).strip()
                if nm.lower() in seen_names: continue      # same Common Name = alternate code of a material already listed
                seen_codes.add(code); seen_names.add(nm.lower()); used.append((f'{sr}{bn}',sr,grp or (c_[5] if len(c_)>5 else ''),nm,code,desc,bn))
        mts=[]
        for u in used:
            if u[6] not in mts: mts.append(u[6])
        rows=[('Plant',plant),tuple(['Module Type']+mts[:4]),(),tuple(['Key No.2','Sr No.','Short Name','Common Name','Material','Material Description','Headers','Stock Data']+[f'{d0+pd.Timedelta(days=k):%d-%m-%Y}' for k in range(ndays)])]
        for key,sr,short,nm,code,desc,bn in used: rows.append((key,sr,short,nm,code,desc,'Plan'))
        rows.append(())
        w._s['Daily Dashboard']=_CachedSheet(rows,'Daily Dashboard'); w.sheetnames.append('Daily Dashboard')
        MIS_REBUILT.append(f"Daily Dashboard ({len(used)} materials from the MIS Standard BOM of {len(mts)} module types in your plan)")
    print('MIS sheets rebuilt in memory (not in the file): '+'; '.join(MIS_REBUILT))
    if missing_: print('WARNING: no MIS BOM Name for '+', '.join(missing_)+" - their modules get no raw-material demand. Set it on Settings > MIS LINK (Override: MIS BOM Name).")


def read_phys_counts(mrp_path):
    """Stock Count sheet of MRP.xlsx: material -> dict(base=(qty, unit), bom=(qty, unit), date). Counted = 'Physical Count Qty'
       (column C) is a number; a blank C is 'no counting' (MIS is used), whatever the Qty_BOM formula shows."""
    out={}
    if not PPC_SRC[0]: return out                                       # physical counts come ONLY from the PPC workbook ('Count')
    try: rows=[list(r) for r in wb_ro(mrp_path)['Stock Count'].iter_rows(values_only=True)]
    except Exception as _e_: _swallow(1101,_e_); return out
    hi=next((i for i,r in enumerate(rows[:10]) if r and str(r[0] or '').strip().lower()=='material'),None)
    if hi is None: return out
    hl=[' '.join(str(x or '').split()).lower() for x in rows[hi]]
    ix=lambda *ks: next((hl.index(k) for k in ks if k in hl),None)
    ic=ix('physical count qty','today stock qty'); iu=ix('base unit of measure','stock uom'); ib=ix('bom uom')
    iq=ix('physical count qty_bom','today qty_bom'); idt=ix('physical count date','count date'); isc=ix('supplier code','physical count supplier')
    dcols=[]                                                             # date-header columns: one column per count day
    for j_,hv in enumerate(rows[hi]):
        dd_=hv if isinstance(hv,(dt.datetime,dt.date)) else None
        if dd_ is None and isinstance(hv,str) and re.match(r'^\s*\d{1,2}/\d{1,2}/\d{4}\s*$',hv):
            try: dd_=dt.datetime.strptime(hv.strip(),'%m/%d/%Y')
            except ValueError: dd_=None
        if dd_: dcols.append((j_,pd.Timestamp(dd_).normalize()))
    g=lambda r,i: r[i] if i is not None and i<len(r) else None
    for r in rows[hi+1:]:
        if not r or not r[0]: continue
        q=g(r,ic)
        if not isinstance(q,(int,float)) or isinstance(q,bool): continue          # no counting -> MIS
        c=str(r[0]).strip(); o=out.setdefault(c,dict(base=[0.0,str(g(r,iu) or '').strip().upper()],bom=[0.0,str(g(r,ib) or '').strip().upper()],date=None))
        o['base'][0]+=float(q)
        qb=g(r,iq); o['bom'][0]+=float(qb) if isinstance(qb,(int,float)) and not isinstance(qb,bool) else float(q)
        d=g(r,idt)
        if not isinstance(d,(dt.datetime,dt.date)):                       # latest date column with a count in it
            ds_=[dd_ for j_,dd_ in dcols if j_<len(r) and isinstance(r[j_],(int,float)) and not isinstance(r[j_],bool)]
            d=max(ds_) if ds_ else None
        if isinstance(d,(dt.datetime,dt.date)) and (o['date'] is None or pd.Timestamp(d)>pd.Timestamp(o['date'])): o['date']=pd.Timestamp(d).normalize()
        sc=str(g(r,isc) or '').strip().upper()                       # counted supplier (KB / FL ...)
        if sc:
            qb2=float(qb) if isinstance(qb,(int,float)) and not isinstance(qb,bool) else float(q)
            e=o.setdefault('by_sup',{}).setdefault(sc,[0.0,0.0]); e[0]+=float(q); e[1]+=qb2
    return out


PP_COLS=[('date',('date',)),('line',('production line','line')),('wp',('wp','watt')),('cust',('customer',)),('qty',('qty','quantity','modules')),
         ('mw',('mw',)),('plant',('plant',)),('model',('module model','model'))]
def pp_rows_std(ws):
    """Production Plan rows as (Date, Production Line, Wp, Customer, Qty, MW, Plant, Module Model), found by COLUMN HEADER -
       so an added column (e.g. 'Month') does not shift anything. Without a recognisable header: the old fixed layout A..H."""
    it=ws.iter_rows(values_only=True); hdr=next(it,None) or ()
    hl=[' '.join(str(x or '').split()).lower() for x in hdr]
    idx=[]
    for key,names in PP_COLS:
        j=next((hl.index(n) for n in names if n in hl),None)
        idx.append(j)
    if idx[0] is None or idx[3] is None or idx[4] is None: idx=list(range(8))      # old layout
    for r in it:
        if not r: continue
        yield tuple((r[j] if j is not None and j<len(r) else None) for j in idx)

def build_daily_stock(mis_path, start=None, ndays=None):
    """Same logic as the MIS 'Daily Dashboard', computed from the MIS sheets' saved values:
       Plan      = SUMIFS('Total Plan Req'!O, K=material, V=plant, A=date)
       Stock     = MB52 Total Stock for <plant> and <plant> WAREHOUSE + 'RM Status for OEM JW' JW WSA to OEM (column L)
       Intransit = SUMIFS('In Transit'!AI, AN=material, AJ=date)
       Adhoc     = SUMIFS(Adhoc!D, B=material, E=date)"""
    print(f'  [{time.time()-_T0:5.0f} s] reading MIS (stock, plan, in transit) ...',flush=True)
    wb=wb_ro(mis_path)
    def df(name):
        rows=list(wb[name].iter_rows(values_only=True))
        return pd.DataFrame(rows[1:],columns=[str(c).strip() if c is not None else f'_c{i}' for i,c in enumerate(rows[0])])
    def bycol(d,i): return d[d.columns[i]]
    dd=[list(r) for r in wb['Daily Dashboard'].iter_rows(values_only=True)]
    plant=str(dd[0][1] or 'USA').strip()
    modtypes=[x for x in dd[1][1:5] if isinstance(x,str) and x.strip()]
    ncols_dash=sum(1 for x in dd[3][8:] if x not in (None,''))
    blocks=[]
    for r in range(4,len(dd)-1):
        row=dd[r]
        if len(row)>6 and str(row[6] or '').strip()=='Plan' and isinstance(row[4],str) and row[4].strip():
            sr=row[1]; key=str(row[0] or ''); mt=key[len(str(sr)):] if key.startswith(str(sr)) else ''
            blocks.append(dict(key=key,sr=sr,short=row[2],common=row[3],mat=row[4].strip(),desc=row[5],mtype=mt))
    if not blocks: die("No materials found on the MIS 'Daily Dashboard' sheet (rows with Headers = Plan)")
    start=pd.Timestamp(start or base_d0()).normalize(); ndays=(ndays or ncols_dash or 141)+max(0,(pd.Timestamp(dt.date.today()).normalize()-start).days)   # base date earlier: same last day
    dates=[start+pd.Timedelta(days=k) for k in range(ndays)]; idx={d:k for k,d in enumerate(dates)}
    mb=df('MB52'); rs=df('RM Status for OEM JW'); tp=df('Total Plan Req'); it=df('In Transit')
    ad=pd.DataFrame(columns=range(8))       # MIS 'Adhoc' sheet is NOT used (no Adhoc receipts, no Adhoc stock conversions)
    num=lambda s: pd.to_numeric(s,errors='coerce').fillna(0)
    day=lambda s: pd.to_datetime(s,errors='coerce').dt.normalize()
    # MATERIALS MISSING FROM THE MIS DAILY DASHBOARD: every raw-material code that the plant's plan needs (MIS Total Plan Req,
    # today onward) is planned, even when the Daily Dashboard does not list it (e.g. the Sabanci 3.2 frame RMALF00000244 and
    # JB RMJUB00000123 of 500 TC G2G_G12R_3.2). A code whose MIS Common Name equals a dashboard material stays an alternate code of it.
    cn_={}
    if 'Common Name' in wb.sheetnames:
        for r_ in wb['Common Name'].iter_rows(min_row=2,values_only=True):
            if r_ and r_[0]: cn_.setdefault(str(r_[0]).strip(),r_)
    # YOUR BOM CODE IS THE MATERIAL (user, 12-Nov-2026: "RMTTT*0010 is zero on hand ... check my BOM, you don't see RMTTT*0010. I want
    # the actual material code with its stock numbers"): a MIS dashboard tape whose code is NOT in your BOM, while your BOM lists ONE
    # other roll length of the same tape (type + width), is planned under YOUR code (its own stock, count, unit, usage); the MIS code
    # becomes its alternate code (converted by roll length below: 1 roll of 50 m = 0.5 roll of 100 m). Today: RMTTT00000010 -> 018.
    CODE_SWAP.clear()
    try: ubm0_=load_user_bom_rm(SETTINGS_SRC[0]) if SETTINGS_SRC and SETTINGS_SRC[0] else {}
    except Exception as _e_: _swallow(1268,_e_); ubm0_={}
    _bm0={b['mat'] for b in blocks}
    for b in blocks:
        ti_=tape_info(b.get('desc'))
        if not ti_ or not ubm0_ or b['mat'] in ubm0_: continue
        hit_=[c_ for c_,m_ in ubm0_.items() if c_ not in _bm0 and (tape_info(m_.get('desc')) or ())[:2]==ti_[:2]]
        if len(hit_)!=1: continue
        CODE_SWAP[hit_[0]]=b['mat']; b['mat']=hit_[0]; b['desc']=ubm0_[hit_[0]].get('desc') or b['desc']; _bm0.add(hit_[0])
        # item name (b['common']) stays the MIS name: the WSA sheets and MIS alternate codes are linked by it
    if CODE_SWAP: print('Your BOM code used as the material (the MIS code = its alternate, converted by roll length): '
                        +'; '.join(f'{o_} -> {n_}' for n_,o_ in CODE_SWAP.items()))
    known_={b['mat'] for b in blocks}; dash_names_={str(b.get('common') or '').strip().lower() for b in blocks}; extra_={}
    for d_,ln_,oem_,co_,bn_,c_,ds_,q_ in zip(day(bycol(tp,0)),bycol(tp,1),bycol(tp,18),bycol(tp,21),bycol(tp,4),bycol(tp,10),bycol(tp,11),num(bycol(tp,14))):
        c_=str(c_ or '').strip()
        if not CODE.match(c_) or c_.upper().startswith('JW') or c_ in known_ or q_<=0 or pd.isna(d_) or d_<start or not plan_line_ok(ln_,oem_,co_,plant): continue
        nm_=str((cn_.get(c_) or [None]*4)[3] or '').strip()
        if nm_ and nm_.lower() in dash_names_: continue
        x_=extra_.setdefault(c_,dict(desc=ds_,bn={},name=nm_)); k_=str(bn_ or '').strip(); x_['bn'][k_]=x_['bn'].get(k_,0.0)+q_
    # raw materials of YOUR BOM (MRP.xlsx) that the MIS list does not have (e.g. labels): planned too
    ubm_=load_user_bom_rm(SETTINGS_SRC[0]) if SETTINGS_SRC and SETTINGS_SRC[0] else {}
    for c_,mt_ in (ubm_.items() if UBOM_DEMAND else []):     # (only when your BOM drives demand)
        if not CODE.match(c_) or c_ in known_ or c_ in extra_: continue
        nm_=str((cn_.get(c_) or [None]*4)[3] or '').strip()
        if nm_ and nm_.lower() in dash_names_: continue           # alternate code of a listed material
        ti_=tape_info(mt_.get('desc'))
        if ti_ and any((tape_info(b_.get('desc')) or ())[:2]==ti_[:2] for b_ in blocks): continue   # another roll length of a listed tape
        extra_[c_]=dict(desc=mt_.get('desc'),bn={str(mt_.get('mtype') or 'Your BOM'):1.0},name=nm_ or str(mt_.get('desc') or c_))
    ADDED_MATS.clear()
    for c_,x_ in sorted(extra_.items()):
        mt_=max(x_['bn'],key=x_['bn'].get); r_=cn_.get(c_) or [None]*6
        sr_=1+max([int(b['sr']) for b in blocks if b['mtype']==mt_ and isinstance(b['sr'],(int,float))] or [0])
        blocks.append(dict(key=f'{sr_}{mt_}',sr=sr_,short=(r_[5] if len(r_)>5 and r_[5] else 'Raw material'),common=x_['name'] or str(x_['desc'] or c_),
                           mat=c_,desc=x_['desc'] or (r_[1] if len(r_)>1 else ''),mtype=mt_))
        ADDED_MATS.append((c_,blocks[-1]['common'],mt_))
    if ADDED_MATS: print('Added materials the plan needs but the MIS Daily Dashboard does not list: '+'; '.join(f'{n} ({c})' for c,n,_ in ADDED_MATS))
    uom_all_={str(k).strip():(r_[2] if len(r_)>2 else None) for k,r_ in cn_.items()}
    for c_,u_ in zip(bycol(mb,2),bycol(mb,7)):
        if c_ is not None and u_: uom_all_.setdefault(str(c_).strip(),u_)
    # Stock (SUMIFS text matches ignore case)
    pl=bycol(mb,13).astype(str).str.strip().str.upper()
    mbk=mb[pl.isin([plant.upper(),plant.upper()+' WAREHOUSE'])]
    stock=num(bycol(mbk,14)).groupby(bycol(mbk,2).astype(str).str.strip()).sum().to_dict()
    jw=num(bycol(rs,11)).groupby(bycol(rs,1).astype(str).str.strip()).sum().to_dict()
    raw_stock=dict(stock); raw_jw=dict(jw)          # untouched copies, for the 'raw stock (own unit)' text
    # Alternate codes: other material codes with the same MIS Common Name as a dashboard material (e.g. Cell Fixation Tape T50:
    # dashboard RMTTT00000010 = 50 m rolls, stock and shipments on RMTTT00000018 = 100 m rolls). MIS converts their STOCK through
    # the Adhoc sheet (here 2,343.9 x 2 = 4,687.8), which becomes opening stock below. Their IN-TRANSIT is added to the dashboard
    # material, converted with the same ratio (Adhoc qty / alternate-code stock) when it can be worked out.
    dash_by_name={str(b['common']).strip().lower():b['mat'] for b in blocks if b.get('common')}
    alias={}
    def _al(code,name):
        c_=str(code or '').strip(); n_=str(name or '').strip().lower()
        if c_ and n_ in dash_by_name and c_!=dash_by_name[n_] and not c_.upper().startswith('JW'): alias.setdefault(c_,dash_by_name[n_])
    # ONLY RULE: the MIS 'Common Name' sheet (A = Material Code, D = Common Name). A code is linked to a dashboard material when its
    # Common Name there equals the dashboard material's Common Name. To link another code, add / correct its row on that sheet.
    if 'Common Name' in wb.sheetnames:
        for r_ in list(wb['Common Name'].iter_rows(min_row=2,values_only=True)):
            if r_ and len(r_)>3: _al(r_[0],r_[3])
    # TAPES: each roll length has its own code. Codes of the same tape (type + width) are linked from their DESCRIPTION and converted
    # by roll length: e.g. 1 roll of 200 m = 0.4 roll of 500 m; 1 roll of 100 m = 2 rolls of 50 m. Their stock is counted directly
    # (MIS's own Adhoc conversion for these tapes is then dropped, so nothing is counted twice), in transit is converted the same way.
    desc_of={}
    if 'Common Name' in wb.sheetnames:
        for r_ in wb['Common Name'].iter_rows(min_row=2,values_only=True):
            if r_ and r_[0] and r_[1]: desc_of.setdefault(str(r_[0]).strip(),str(r_[1]))
    for c_,d_ in zip(bycol(mb,2),bycol(mb,5)):
        if c_ is not None and d_: desc_of.setdefault(str(c_).strip(),str(d_))
    for c_,d_ in zip(bycol(it,39),bycol(it,4)):
        if c_ is not None and d_: desc_of.setdefault(str(c_).strip(),str(d_))
    tape_factor={}; tape_dash={}; tape_len={}
    for b in blocks:
        ti=tape_info(b.get('desc') or desc_of.get(b['mat']))
        if not ti: continue
        tape_dash[b['mat']]=ti; tape_len[b['mat']]=ti[2]
        for c_,dsc in desc_of.items():
            tc=tape_info(dsc)
            if c_==b['mat'] or c_.upper().startswith('JW') or not tc or tc[:2]!=ti[:2]: continue
            alias[c_]=b['mat']; tape_factor[c_]=tc[2]/ti[2]; tape_len[c_]=tc[2]
    pack_window={}                                    # dated pack conversions: counted only between 'use from' and 'use until'
    for c_,t_ in PACK_FACTORS.items():               # your own pack conversions (Settings sheet): counted like tape rolls
        d_,f_=t_[0],t_[1]; fr=t_[2] if len(t_)>2 else None; to=t_[3] if len(t_)>3 else None
        if d_ not in {b['mat'] for b in blocks} or not f_: continue
        if c_ in {b['mat'] for b in blocks}: continue      # a substitute that is an MRP material itself (EPE 480 for EPE 420) keeps its OWN stock / plan /
                                                           # in transit - never folded into the other code (user 25-Nov-2026: "show separately ... arrival of EPE 420 will confuse")
        if fr is None and to is None: alias[c_]=d_; tape_factor[c_]=float(f_); tape_dash.setdefault(d_,('pack',0,1.0)); tape_len.setdefault(d_,1.0); tape_len[c_]=float(f_)
        else: pack_window.setdefault(d_,[]).append(dict(code=c_,factor=float(f_),fr=fr,to=to))
    # YOUR mapping (Material Group sheet, MRP.xlsx): 'MRP material' + 'Factor' - overrides the automatic matching above
    load_user_masters(SETTINGS_SRC[0] if SETTINGS_SRC and SETTINGS_SRC[0] else None)
    bset_={b['mat'] for b in blocks}; USER_ALT_MISS.clear()
    for c_,(d_,f_) in USER_ALT.items():
        if c_==d_:                                        # you say: this code is its own material - undo an automatic link
            if c_ in alias: alias.pop(c_,None); tape_factor.pop(c_,None)
            continue
        if d_ not in bset_: USER_ALT_MISS.append((c_,d_)); continue
        alias[c_]=d_; tape_factor[c_]=f_; tape_dash.setdefault(d_,('pack',0,1.0)); tape_len.setdefault(d_,1.0); tape_len[c_]=tape_len.get(d_,1.0)*f_
    uom_raw={}
    for c_,u_ in zip(bycol(mb,2),bycol(mb,7)):
        if c_ is not None and isinstance(u_,str) and u_.strip(): uom_raw.setdefault(str(c_).strip(),u_.strip())

    # conversion ratio alternate code -> dashboard code, from MIS's own Adhoc conversion
    adq={}
    for c_,q_ in zip(bycol(ad,1),num(bycol(ad,3))):
        if c_ is not None: adq[str(c_).strip()]=adq.get(str(c_).strip(),0.0)+q_
    alias_factor=dict(tape_factor); conv_qty={}       # alternate code -> ratio to the dashboard unit ; dashboard code -> Adhoc qty that is a conversion
    for n_,o_ in CODE_SWAP.items(): CODE_SWAP_F[n_]=alias_factor.get(o_,1.0)   # 1 MIS-code roll = f rolls of your BOM code (tape 010 -> 018: 0.5)
    nm_of={}
    for c_,n_ in zip(bycol(mb,2),bycol(mb,18)):
        if c_ is not None and n_: nm_of.setdefault(str(c_).strip(),str(n_))
    first=lambda t: (re.split(r'[^a-z0-9]+',str(t or '').lower()) or [''])[0]
    RATIOS=(0.2,0.25,0.4,0.5,1.0,2.0,2.5,4.0,5.0)
    conv_unlinked={}                    # Adhoc conversion of a code that the Common Name sheet does not link -> Check only
    for d_r,a_r,t_r in zip(bycol(ad,1),num(bycol(ad,3)),day(bycol(ad,4))):
        d_=str(d_r or '').strip()
        if d_ not in {b['mat'] for b in blocks} or a_r<=0 or pd.isna(t_r) or t_r>start or d_ in tape_dash: continue
        dname=next((b['common'] for b in blocks if b['mat']==d_),'')
        hit=False
        for c_,s_ in stock.items():                      # 1) codes linked by the Common Name sheet: ratio from MIS's Adhoc conversion
            if alias.get(c_)!=d_ or s_<=0: continue
            r_=a_r/s_
            if any(abs(r_-x)<=1e-4*x for x in RATIOS):
                alias_factor[c_]=round(r_,6); conv_qty[d_]=conv_qty.get(d_,0.0)+a_r; hit=True; break
        if hit: continue
        for c_,s_ in stock.items():                      # 2) looks like a conversion of an unlinked code: report, do not count
            if c_==d_ or c_ in alias or s_<=0 or c_.upper().startswith('JW') or first(nm_of.get(c_,''))!=first(dname): continue
            r_=a_r/s_
            if any(abs(r_-x)<=1e-4*x for x in RATIOS): conv_unlinked.setdefault(d_,[]).append((c_,nm_of.get(c_,''),round(r_,6),a_r)); break
    # stock by location and unit of measure (for the Inventory sheet)
    loc=pd.DataFrame({'m':bycol(mbk,2).astype(str).str.strip(),'w':pl[pl.isin([plant.upper(),plant.upper()+' WAREHOUSE'])].str.endswith('WAREHOUSE'),'q':num(bycol(mbk,14))})
    st_main=loc[~loc['w']].groupby('m')['q'].sum().to_dict(); st_wh=loc[loc['w']].groupby('m')['q'].sum().to_dict()
    uom={}
    for m_,u_ in zip(bycol(mb,2).astype(str).str.strip(),bycol(mb,7)):
        if isinstance(u_,str) and u_.strip(): uom.setdefault(m_,u_.strip())
    # codes with no MB52 row (no stock yet, e.g. RMTTT00000010, RMGLS00000356): unit from the MIS Common Name sheet, else Total Plan Req
    if 'Common Name' in wb.sheetnames:
        for r_ in wb['Common Name'].iter_rows(min_row=2,values_only=True):
            if r_ and len(r_)>2 and r_[0] and isinstance(r_[2],str) and r_[2].strip(): uom.setdefault(str(r_[0]).strip(),r_[2].strip())
    for m_,u_ in zip(bycol(tp,10).astype(str).str.strip(),bycol(tp,13)):
        if isinstance(u_,str) and u_.strip(): uom.setdefault(m_,u_.strip())
    if CODE_SWAP:                     # MIS Total Plan Req of a swapped tape code -> your BOM code, in its roll (x factor)
        inv_=_SWP_INV(); cc_=bycol(tp,10).astype(str).str.strip(); mk_=cc_.isin(list(inv_))
        if mk_.any():
            tp=tp.copy(); tp.loc[mk_,tp.columns[14]]=num(bycol(tp,14))[mk_]*cc_[mk_].map(lambda o_: inv_[o_][1]); tp.loc[mk_,tp.columns[10]]=cc_[mk_].map(lambda o_: inv_[o_][0])
    tpmask=pd.Series([plan_line_ok(a,b,c,plant) for a,b,c in zip(bycol(tp,1),bycol(tp,18),bycol(tp,21))],index=tp.index)   # USA* lines, OEM Category WSA
    tpf=pd.DataFrame({'m':bycol(tp,10).astype(str).str.strip(),'d':day(bycol(tp,0)),'q':num(bycol(tp,14))})[tpmask.values]
    plan_all={}
    for (m_,d_),q_ in tpf.dropna(subset=['d']).groupby(['m','d'])['q'].sum().items(): plan_all.setdefault(m_,{})[d_]=q_
    def daily(frame,qty_i,mat_i,date_i,extra=None,clamp=False,remap=None):
        f=pd.DataFrame({'m':bycol(frame,mat_i).astype(str).str.strip(),'d':day(bycol(frame,date_i)),'q':num(bycol(frame,qty_i))})
        if remap: f['m']=f['m'].map(lambda x: remap.get(x,x))
        if extra is not None: f=f[extra.values]
        if clamp: f.loc[f['d']<start,'d']=start      # arrivals dated before the first day (e.g. MIS saved yesterday) count on day 1
        f=f[f['d'].isin(idx)]
        out={}
        for (m,d),q in f.groupby(['m','d'])['q'].sum().items(): out.setdefault(m,[0.0]*ndays)[idx[d]]+=q
        return out
    plan=daily(tp,14,10,0,extra=tpmask)
    fc=pd.DataFrame({'m':bycol(tp,10).astype(str).str.strip(),'c':bycol(tp,2).map(lambda x:' '.join(str(x).split()) if x is not None and x==x else '(no customer)'),
                     'd':day(bycol(tp,0)),'q':num(bycol(tp,14))})[tpmask.values]
    fc=fc[fc['d'].isin(idx)&(fc['q']!=0)]
    plan_cust={}
    for (m_,c_,d_),q_ in fc.groupby(['m','c','d'])['q'].sum().items(): plan_cust.setdefault(m_,{}).setdefault(c_,[0.0]*ndays)[idx[d_]]+=q_
    if IT_STATUS_ONLY:                         # MIS STATUS 'IN-TRANSIT' only
        it=it[[it_status_group(v_)=='In transit' for v_ in bycol(it,22)]].reset_index(drop=True)
    itq=it.copy(); itq[itq.columns[34]]=[(q_*alias_factor.get(str(c_).strip(),1.0)) if isinstance(q_,(int,float)) else q_ for q_,c_ in zip(bycol(it,34),bycol(it,39))]
    trans=daily(itq,34,39,35,clamp=True,remap=alias); adh=daily(ad,3,1,4,clamp=True)
    # supplier detail: stock by Make (RM Status for OEM JW: USA + USA Warehouse; the JW stock at the OEM is in no calculation) and in transit by Supplier Name
    mk_stock={}
    for mk,cd,a_,b_ in zip(bycol(rs,0),bycol(rs,1),num(bycol(rs,5)),num(bycol(rs,7))):
        if not isinstance(mk,str) or not mk.strip(): continue
        d_=mk_stock.setdefault(str(cd).strip(),{}); d_[mk.strip().upper()]=d_.get(mk.strip().upper(),0.0)+a_+b_
    f=pd.DataFrame({'m':bycol(itq,39).astype(str).str.strip().map(lambda x: alias.get(x,x)),'s':bycol(itq,1).map(lambda x:' '.join(str(x).split()) if x is not None and x==x else '(no supplier)'),'d':day(bycol(itq,35)),'q':num(bycol(itq,34))})
    f.loc[f['d']<start,'d']=start; f=f[f['d'].isin(idx)&(f['q']>0)]
    trans_sup={}
    po_sup={}
    for po_,sn_ in zip(bycol(it,6),bycol(it,1)):
        if po_ is None or not isinstance(sn_,str) or not sn_.strip(): continue
        k_=str(po_).strip(); k_=k_[:-2] if k_.endswith('.0') else k_
        po_sup.setdefault(k_,' '.join(sn_.split()))
    for (m_,s_,d_),q_ in f.groupby(['m','s','d'])['q'].sum().items(): trans_sup.setdefault(m_,{}).setdefault(s_,[0.0]*ndays)[idx[d_]]+=q_
    # Adhoc rows (dated on or before today) for a material whose alternate codes hold stock are MIS stock conversions
    # (e.g. Cell Fixation 100 m rolls -> 50 m rolls, old busbar codes) = stock you have now -> opening Stock. Other Adhoc rows stay as Adhoc.
    adh_open=dict(conv_qty)            # only Adhoc rows proven to be a conversion of another code's stock
    tape_note={}
    for d_ in tape_dash:              # tapes: alternate-roll stock counted directly by length; MIS's Adhoc conversion for them is dropped
        conv=sum(stock.get(c_,0.0)*f_ for c_,f_ in tape_factor.items() if alias.get(c_)==d_)
        convj=sum(jw.get(c_,0.0)*f_ for c_,f_ in tape_factor.items() if alias.get(c_)==d_)
        stock[d_]=stock.get(d_,0.0)+conv; st_main[d_]=st_main.get(d_,0.0)+conv; jw[d_]=jw.get(d_,0.0)+convj
        a0=sum(q_ for c_,q_,t_ in zip(bycol(ad,1),num(bycol(ad,3)),day(bycol(ad,4))) if str(c_ or '').strip()==d_ and pd.notna(t_) and t_<=start)
        dropped=0.0
        if a0>0 and abs(a0-conv)<=max(1.0,0.02*max(a0,conv)) and d_ in adh: adh[d_][0]=max(0.0,adh[d_][0]-a0); dropped=a0
        tape_note[d_]=dict(conv=conv,adhoc=a0,dropped=dropped,lens={c_:tape_len.get(c_) for c_ in tape_factor if alias.get(c_)==d_},len=tape_len.get(d_))
    # alternate codes holding stock that MIS does NOT convert (no matching Adhoc row) - reported on Check, not counted
    alt_uncounted={}
    for c_,d_ in alias.items():
        if stock.get(c_,0)>0 and c_ not in alias_factor and c_ not in tape_factor: alt_uncounted.setdefault(d_,{})[c_]=stock[c_]
    for m_,q_ in adh_open.items():
        stock[m_]=stock.get(m_,0.0)+q_; st_main[m_]=st_main.get(m_,0.0)+q_
        if m_ in adh: adh[m_][0]=max(0.0,adh[m_][0]-q_)
    alt_of={}
    for c_,d_ in alias.items(): alt_of.setdefault(d_,[]).append(c_)
    for b in blocks:
        m=b['mat']; b['stock']=stock.get(m,0.0)+jw.get(m,0.0)
        b['alt_codes']=sorted(alt_of.get(m,[])); b['adhoc_open']=adh_open.get(m,0.0); b['alt_factor']={c_:alias_factor[c_] for c_ in alt_of.get(m,[]) if c_ in alias_factor}
        b['alt_uncounted']=alt_uncounted.get(m,{}); b['adhoc_other']=max(0.0,sum(adh.get(m,[0.0])[:1])) if m in adh else 0.0
        b['conv_unlinked']=conv_unlinked.get(m,[]); b['tape']=tape_note.get(m)
        b['pack_window']=[dict(w_,qty=raw_stock.get(w_['code'],0.0)*w_['factor'])   # own stock only (JW at OEM ignored in calculations)
                        for w_ in pack_window.get(m,[])]
        parts=[]
        if raw_stock.get(m): parts.append(f"{m} {raw_stock[m]:,.1f} {uom_raw.get(m,'')}".strip())
        if raw_jw.get(m): parts.append(f"JW {raw_jw[m]:,.1f}")
        for c_ in sorted(alt_of.get(m,[])):
            f_=alias_factor.get(c_)
            if f_ and raw_stock.get(c_): parts.append(f"{c_} {raw_stock[c_]:,.1f} {uom_raw.get(c_,'')} x{f_:g}".replace('  ',' '))
            if f_ and raw_jw.get(c_) and c_ in tape_factor: parts.append(f"JW {c_} {raw_jw[c_]:,.1f} x{f_:g}")
        b['raw']=' + '.join(parts) or '0'
        b['plan']=plan.get(m,[0.0]*ndays); b['cust_plan']=dict(sorted(plan_cust.get(m,{}).items(),key=lambda kv:-sum(kv[1]))); b['transit']=trans.get(m,[0.0]*ndays); b['adhoc']=adh.get(m,[0.0]*ndays)
    for b in blocks:
        m=b['mat']; b['st_main']=st_main.get(m,0.0); b['st_wh']=st_wh.get(m,0.0); b['st_jw']=jw.get(m,0.0); b['uom']=uom.get(m,'')
        b['plan_all']=plan_all.get(m,{})
    # RAW MATERIALS COUNTED in the PPC 'Count' sheet: the physical count replaces MIS MB52 for that material (JW at OEM stays)
    pc_=read_phys_counts(SETTINGS_SRC[0]) if SETTINGS_SRC and SETTINGS_SRC[0] else {}
    pcb_={}
    for c_,o_ in pc_.items():
        m_=alias.get(c_,c_); f_=tape_factor.get(c_,1.0) if m_!=c_ else 1.0
        mu_=str(uom.get(m_,'') or '').strip().upper()
        q_=o_['bom'][0] if (o_['bom'][1]==mu_ or not mu_) else (o_['base'][0] if o_['base'][1]==mu_ else o_['bom'][0])   # Qty_BOM when its unit is the MRP unit (glass crates -> PC), else the count's own unit (RTV drums)
        d_=o_['date']
        if m_ in {b['mat'] for b in blocks}: e_=pcb_.setdefault(m_,[0.0,[],None,[]]); e_[0]+=q_*f_; e_[1].append((c_,q_,f_)); e_[3].append((d_,q_*f_))
    for e_ in pcb_.values():           # checker 19-Nov-2026: count date of a material counted under 2+ codes = the OLDEST count date of the codes holding stock
        ds_=[d for d,q in e_[3] if d is not None and q>0.5] or [d for d,_ in e_[3] if d is not None]   # (busbar 047 counted 17-Nov + alias 016 / 033 on 19-Nov:
        e_[2]=(min(ds_) if [q for d,q in e_[3] if d is not None and q>0.5] else max(ds_)) if ds_ else None   # roll forward from 17-Nov, an alias counted 0 does not reset it)
    COUNTED_RM.clear()
    for b in blocks:
        if b['mat'] in pcb_:
            q_,parts_,d_,_=pcb_[b['mat']]
            COUNTED_RM[b['mat']]=dict(qty=q_,parts=parts_,date=d_,mis=b['st_main']+b['st_wh'],by_sup=(pc_.get(b['mat']) or {}).get('by_sup') or {})
            b['stock']=q_+b['st_jw']; b['st_main']=q_; b['st_wh']=0.0
    # STOCK DATA: every source line behind each material's stock, so the workbook can show stock as a formula (Stock Data sheet)
    stock_lines=[]; bmats={b['mat'] for b in blocks if b['mat'] not in COUNTED_RM}
    for m_,e_ in COUNTED_RM.items():
        for c_,q_,f_ in e_['parts']:
            stock_lines.append(dict(code=m_,type='Stock',source='PPC Count (physical count)'+(f" {e_['date']:%d-%b}" if e_['date'] else ''),
                                    src=c_,batch=None,sloc=None,loc=None,qty=float(q_),uom=uom.get(c_,''),factor=f_))
    for r_ in mbk.itertuples(index=False):
        c_=str(r_[2] or '').strip(); q_=r_[14] if isinstance(r_[14],(int,float)) else 0.0
        if c_ in bmats: tgt_,f_=c_,1.0
        elif alias.get(c_) in bmats and c_ in tape_factor: tgt_,f_=alias[c_],tape_factor[c_]
        else: continue
        if not q_: continue
        stock_lines.append(dict(code=tgt_,type='Stock',source='MIS MB52 (Total Stock)',src=c_,batch=r_[3],sloc=' '.join(str(x) for x in (r_[4],r_[8]) if x),
                                loc=r_[13],qty=float(q_),uom=r_[7],factor=f_))
    for m_,q_ in adh_open.items():
        if m_ in bmats and q_ and False: stock_lines.append(dict(code=m_,type='Stock',source='MIS Adhoc (stock of an alternate code converted by MIS)',src=m_,batch=None,sloc=None,loc=None,qty=float(q_),uom=uom.get(m_,''),factor=1.0))
    bmats_all={b['mat'] for b in blocks}
    for r_ in rs.itertuples(index=False):
        c_=str(r_[1] or '').strip(); q_=r_[11] if len(r_)>11 and isinstance(r_[11],(int,float)) else 0.0
        if not q_: continue
        if c_ in bmats_all: tgt_,f_=c_,1.0
        elif alias.get(c_) in bmats_all and c_ in tape_factor: tgt_,f_=alias[c_],tape_factor[c_]
        else: continue
        stock_lines.append(dict(code=tgt_,type='to OEM',source="MIS RM Status for OEM JW (JW WSA to OEM)",src=c_,batch=None,sloc=None,loc='JW at OEM',qty=float(q_),uom=uom.get(c_,''),factor=f_))
    tot_={}
    for x_ in stock_lines: tot_[x_['code']]=tot_.get(x_['code'],0.0)+x_['qty']*x_['factor']
    bad_=[b['mat'] for b in blocks if abs(tot_.get(b['mat'],0.0)-b['stock'])>max(0.01,1e-6*abs(b['stock']))]
    if bad_: print('NOTE: Stock Data lines do not add up to the stock for: '+', '.join(bad_))
    return dict(plant=plant,modtypes=modtypes,blocks=blocks,dates=dates,src=os.path.basename(mis_path),start=start,alias=alias,alias_factor=alias_factor,
                by_code=dict(transit=trans,adhoc=adh),mk_stock=mk_stock,trans_sup=trans_sup,po_sup=po_sup,
                stock_raw=raw_stock,desc_all=desc_of,cn_all=cn_,uom_all=uom_all_,stock_lines=stock_lines)

SC_COL='C'      # column of the stock quantity on 'Stock Count' used in formulas (set per workbook: 'Today Qty_BOM' when present)
def count_qty_idx(header):
    """Index of the stock quantity to use on a Stock Count header row: 'Today Qty_BOM' (count converted to the BOM unit) first."""
    h=[str(x).strip().lower() if x is not None else '' for x in (header or [])]
    for name in ('today qty_bom','today qty bom','today qty (bom)'):
        if name in h: return h.index(name)
    for name in ('today stock qty','today qty'):
        if name in h: return h.index(name)
    return 2
def count_header_and_rows(ws_rows):
    rows=[list(r) for r in ws_rows]
    hr=next((i for i,r in enumerate(rows) if r and str(r[0] or '').strip()=='Material'),None)
    return (rows[hr],rows[hr+1:]) if hr is not None else (None,rows)

GRID_SD=11; GRID_C0=12   # MRP grid: J = Supplier... K = Stock Data, L.. = days (after the Production Line column)
def set_print(ws, title):
    """Print layout used by every report sheet: landscape, fit to page width, header with sheet title + run date, page numbers."""
    try:
        ws.page_setup.orientation='landscape'; ws.page_setup.paperSize=ws.PAPERSIZE_A4
        ws.page_setup.fitToWidth=1; ws.page_setup.fitToHeight=0; ws.sheet_properties.pageSetUpPr.fitToPage=True
        ws.print_options.horizontalCentered=True
        ws.page_margins.left=ws.page_margins.right=0.3; ws.page_margins.top=ws.page_margins.bottom=0.5
        ws.oddHeader.left.text=f'MRP Daily - {title}'; ws.oddHeader.right.text='&D'; ws.oddFooter.center.text='Page &P of &N'
    except Exception as _e_: _swallow(1458,_e_); pass

PRINT_MAX_COVER=22  # DMS BOM Weekly prints / shows only 'Days of cover (stock only)' below 22 days (USER 25-Nov-2026)
PRINT_MIN_SCALE=55   # A4 print: smallest zoom (%) used to fit a column group on the page width (DEFAULT, 25-Nov-2026)
PRINT_AL_GROUPS=('Component','Lines (share of plan volume)','Short from → to (with in transit)','Material lead time (days)',
                 'Production Daily Target_without OEM','Usage/Day (module plan)','In transit (total qty, WSA RM Planning)')
def print_a4(ws, hr, groups, tcols, title):
    """A4 landscape print of a wide sheet (user 25-Nov-2026 "print A4 landscape each sheet BOM material dashboard, ALERT and Action List"):
       the columns are cut into pages only BETWEEN column groups (band row), the code / name columns and the header rows repeat on every
       page, and the zoom is the largest one that keeps the fewest pages wide (never below PRINT_MIN_SCALE). Pages print across, then down,
       so the first sheets of paper hold the highest-risk materials complete. Print settings only - nothing on screen changes."""
    from openpyxl.utils import get_column_letter as _L
    from openpyxl.worksheet.pagebreak import Break, ColBreak, RowBreak
    last=max(g[1] for g in groups); lr=ws.max_row
    while lr>hr and all(ws.cell(lr,c).value in (None,'') for c in range(1,last+1)): lr-=1
    wid={c:(0 if ws.column_dimensions[_L(c)].hidden else ((ws.column_dimensions[_L(c)].width or 8.43)*7+5)/96) for c in range(1,last+1)}
    gw=lambda a,b: sum(wid[c] for c in range(a,b+1))
    tw=gw(*tcols) if tcols else 0
    def pack(avail):
        pages=[]; cur=0; first=True
        for a,b in groups:
            w_=gw(a,b); base=0 if first else tw
            if pages and cur+w_<=avail: pages[-1][1]=b; cur+=w_; continue
            if base+w_>avail: return None
            if pages: first=False; base=tw
            pages.append([a,b]); cur=base+w_
        return pages
    avail0=(11.69-0.25-0.25)*0.92   # 8 % safety: Excel / printer drivers draw columns a little wider than the width x 7 + 5 px rule
    best=None
    for s_ in range(100,PRINT_MIN_SCALE-1,-1):
        pg=pack(avail0*100/s_)
        if pg is None: continue
        if best is None or len(pg)<len(best[1]): best=(s_,pg)
    if best is None: best=(PRINT_MIN_SCALE,[[a,b] for a,b in groups])
    s_,pg=best
    ps=ws.page_setup
    ps.orientation='landscape'; ps.paperSize=ws.PAPERSIZE_A4; ps.scale=s_; ps.fitToWidth=None; ps.fitToHeight=None
    ps.pageOrder='overThenDown'; ws.sheet_properties.pageSetUpPr.fitToPage=False
    ws.page_margins.left=ws.page_margins.right=0.25; ws.page_margins.top=ws.page_margins.bottom=0.55
    ws.page_margins.header=ws.page_margins.footer=0.25
    ws.print_options.horizontalCentered=False; ws.print_options.gridLines=False
    ws.print_area=f'A1:{_L(last)}{lr}'
    ws.print_title_rows=f'{hr-1}:{hr}' if hr>1 else f'{hr}:{hr}'
    ws.print_title_cols=f'{_L(tcols[0])}:{_L(tcols[1])}' if tcols and len(pg)>1 else None
    ws.col_breaks=ColBreak(); ws.row_breaks=RowBreak()
    for a,b in pg[:-1]: ws.col_breaks.append(Break(id=b))
    ws.oddHeader.left.text=title; ws.oddHeader.left.size=9; ws.oddHeader.right.text='Printed &D'; ws.oddHeader.right.size=9
    ws.oddFooter.left.text='&F'; ws.oddFooter.left.size=8; ws.oddFooter.right.text='Page &P of &N'; ws.oddFooter.right.size=8
    ws.oddFooter.center.text=None
    return s_,len(pg)
def print_rows_fit(ws, hr, groups, tcols, title):
    """BOM Material Dashboard - Weekly print (user 25-Nov-2026 "pdf fit rows in page, I don't want to fit everything in 1 page"): the zoom is the
       largest one (max 100 %) at which ALL shown rows of the sheet fit the page HEIGHT (Excel 'Fit all rows on one page'), so no page is cut
       between rows; the columns then run across as many pages as needed, cut between column groups first (a group wider than a page is cut
       by column), Item name (column A) repeats on every page. Rows hidden by the < 22 days filter are not counted. Print settings only."""
    from openpyxl.utils import get_column_letter as _L
    from openpyxl.worksheet.pagebreak import Break, ColBreak, RowBreak
    last=max(g[1] for g in groups); lr=ws.max_row
    while lr>hr and (ws.row_dimensions[lr].hidden or all(ws.cell(lr,c).value in (None,'') for c in range(1,last+1))): lr-=1
    top,bot=0.5,0.45
    th=sum(0 if ws.row_dimensions[r].hidden else (ws.row_dimensions[r].height or 15) for r in range(1,lr+1))
    s_=max(10,min(100,int((8.27-top-bot)*72*0.96/th*100))) if th else 100
    avail=(11.69-0.4)*0.92*100/s_                                    # inches of sheet per page at this zoom (8 % width safety)
    wid={c:(0 if ws.column_dimensions[_L(c)].hidden else ((ws.column_dimensions[_L(c)].width or 8.43)*7+5)/96) for c in range(1,last+1)}
    tw=sum(wid[c] for c in range(tcols[0],tcols[1]+1)) if tcols else 0
    pages=[]; cur=None                                               # [first col, last col, width used]
    def new_page(a):
        nonlocal cur
        base=0 if not pages else tw
        cur=[a,a-1,base]; pages.append(cur)
    for ga,gb in groups:
        gw_=sum(wid[c] for c in range(ga,gb+1))
        if cur is not None and cur[2]+gw_<=avail: cur[1]=gb; cur[2]+=gw_; continue
        new_page(ga)
        if cur[2]+gw_<=avail: cur[1]=gb; cur[2]+=gw_; continue
        for c in range(ga,gb+1):                                     # group wider than a page: cut by column
            if cur[2]+wid[c]>avail and cur[1]>=cur[0]: new_page(c)
            cur[1]=c; cur[2]+=wid[c]
    ps=ws.page_setup
    ps.orientation='landscape'; ps.paperSize=ws.PAPERSIZE_A4; ps.scale=s_; ps.fitToWidth=None; ps.fitToHeight=None
    ps.pageOrder='overThenDown'; ws.sheet_properties.pageSetUpPr.fitToPage=False
    ws.page_margins.left=ws.page_margins.right=0.2; ws.page_margins.top=top; ws.page_margins.bottom=bot
    ws.page_margins.header=ws.page_margins.footer=0.2
    ws.print_options.horizontalCentered=False; ws.print_options.gridLines=False
    ws.print_area=f'A1:{_L(last)}{lr}'
    ws.print_title_rows=None
    ws.print_title_cols=f'{_L(tcols[0])}:{_L(tcols[1])}' if tcols and len(pages)>1 else None
    ws.col_breaks=ColBreak(); ws.row_breaks=RowBreak()
    for pg in pages[:-1]: ws.col_breaks.append(Break(id=pg[1]))
    ws.oddHeader.left.text=title; ws.oddHeader.left.size=9; ws.oddHeader.right.text='Printed &D'; ws.oddHeader.right.size=9
    ws.oddFooter.left.text='&F'; ws.oddFooter.left.size=8; ws.oddFooter.right.text='Page &P of &N'; ws.oddFooter.right.size=8
    ws.oddFooter.center.text=None
    return s_,len(pages)
def _band_groups(ws, band, last):
    """Column groups = the merged cells of the band row (one group per band); columns outside a band join the group before them."""
    ms=sorted((m.min_col,m.max_col) for m in ws.merged_cells.ranges if m.min_row==band==m.max_row)
    out=[]
    for c in range(1,last+1):
        g=next((m for m in ms if m[0]==c),None)
        if g: out.append([g[0],min(g[1],last)])
        elif not out: out.append([c,c])
        elif c>out[-1][1]: out[-1][1]=c
    return [tuple(g) for g in out]
def print_a4_all(wb, label):
    """Apply the A4 print layout to the ALERT, Action List and BOM Material Dashboard - Weekly sheets of a workbook before it is saved."""
    hd=lambda ws,r: {str(ws.cell(r,c).value or '').strip():c for c in range(ws.max_column,0,-1)}
    for ws in wb.worksheets:
        try:
            if ws.title=='ALERT':
                hr=next((r for r in range(1,40) if str(ws.cell(r,1).value or '').strip()=='Sr.'),None)
                if not hr: continue
                h=hd(ws,hr); last=max(h.values()); tc=(h.get('Code'),h.get('Item name'))
                print_a4(ws,hr,_band_groups(ws,hr-1,last),tc if all(tc) else None,f'{label} - ALERT')
            elif ws.title=='Action List':
                h=hd(ws,4); last=max(h.values()); st=sorted({h[k] for k in PRINT_AL_GROUPS if k in h}|{1})
                gr=[(a,(st[i+1]-1 if i+1<len(st) else last)) for i,a in enumerate(st)]
                tc=(h.get('Component'),h.get('Component description'))
                print_a4(ws,4,gr,tc if all(tc) else None,f'{label} - Action List')
            elif ws.title==HEAT_NAME and str(ws.cell(2,1).value or '').strip()=='Item name':
                h=hd(ws,2); last=max(h.values())               # user 25-Nov-2026: all ROWS on one page height, columns across pages
                print_rows_fit(ws,2,_band_groups(ws,1,last),(1,1),f'{label} - {HEAT_NAME}')
        except Exception as _e_: _swallow(1459,_e_)

def make_daily_stock_xlsx(DS, path, title='MRP Daily', table='MRPDaily', weekly_of=None, rowmap=None, cache=None, target_days=None):
    """Build the Daily Stock sheet (or its weekly view) in a temporary workbook, formatted like the MIS 'Daily Dashboard'.
       Columns: A Key, B Sr No., C Short Name, D Common Name, E Material, F Material Description, G Customer, H Supplier, I Headers, J Stock Data, K.. days/weeks.
       Plan is split by customer: the material's Plan row (All customers) is followed by one Plan row per customer.
       Materials with supplier detail (b['sup'], glass) get, after their normal block, one 3-row block per supplier:
       Stock (on hand, in Stock Data only), Intransit, Open PO (no vessel yet). Plan stays on the material block only."""
    from openpyxl.styles import Border, Side
    from openpyxl.styles.colors import Color
    from openpyxl.worksheet.table import Table, TableStyleInfo
    from openpyxl.formatting.rule import ColorScaleRule, Rule
    from openpyxl.styles.differential import DifferentialStyle
    wb=Workbook(); ws=wb.active; ws.title=title
    NAVY='FF002060'; V='Verdana'
    acc='_ * #,##0_ ;_ * \\-#,##0_ ;_ * "-"??_ ;_ @_ '
    thin=Side(style='thin'); dot=Side(style='dotted')
    bthin=Border(left=thin,right=thin,top=thin,bottom=thin); bdot=Border(left=dot,right=dot,top=dot,bottom=dot)
    tfill=lambda th,t: PatternFill('solid',fgColor=Color(theme=th,tint=t))
    ctr=Alignment(horizontal='center',vertical='center',wrap_text=True)
    lab=dict(font=Font(name=V,size=10,bold=True,color=Color(theme=1)),fill=tfill(3,0.8999908444471572),border=bthin,alignment=ctr)
    for r_,t_ in ((1,'Plant'),(2,'Module Type')):
        c=ws.cell(r_,1,t_)
        for k,v in lab.items(): setattr(c,k,v)
    c=ws.cell(1,2,DS['plant']); c.font=Font(name=V,size=10,bold=True,color=Color(theme=1)); c.fill=tfill(5,0.5999938962981048); c.border=bthin; c.alignment=ctr
    for j,mt in enumerate(DS['modtypes'][:4]):
        c=ws.cell(2,2+j,mt); c.font=Font(name=V,size=10,bold=True,color='FFFF0000'); c.border=bthin; c.alignment=ctr
    plan_src=("Plan = module plan (PPC Plan + Revised Plan; without a PPC file the Production Plan of MRP.xlsx) x usage per module (PPC BOM, else MIS Standard BOM), OEM plan left out. ")
    if weekly_of: plan_src=(f"Weekly view of '{weekly_of}' (live formulas). Plan / Intransit / Open PO / Adhoc = week totals; Stock and Inv. Days = end of week. "+plan_src)
    c=ws.cell(1,4,f"Updated {dt.date.today():%d-%b-%Y} from {DS['src']}. Line rows (Customer + Production Line): column I = where that line is supplied from, red cells = days that line is not covered. "+plan_src+f"Stock (RM) = MB52 {DS['plant']} + {DS['plant']} Warehouse + MIS code conversions (JW stock at the OEM left out), replaced by the PPC Count (rolled forward to today) when counted; CS / PM = PPC Count. "
              "Stock row = projected stock including in transit and adhoc. Open PO (no vessel yet, WSA) is NOT in the Stock row; "
              f"'Inv. Days (stock + in transit + Open PO)' adds it. Inv. Days = projected stock / Usage/Day (average plan per production day over the material's transit lead time, from that day); colours: 0 red, half the target yellow, target green. "
              "Glass: supplier rows show each supplier's stock (PPC Count by supplier when counted, else MB52 batches / MIS RM Status by Make; JW at the OEM left out), in transit and Open PO.")
    c.font=Font(name=V,size=8,italic=True,color=NAVY)
    ws.row_dimensions[1].height=34.5; ws.row_dimensions[2].height=36
    dates=DS['dates']; SD=GRID_SD; C0=GRID_C0
    wkn=lambda d: ((d.date()-dt.date(d.year,1,1)).days+dt.date(d.year,1,1).weekday())//7+1   # = WEEKNUM(date,2)
    if weekly_of:                            # one column per Monday week: (week start, first daily index, last daily index)
        per=[]
        for k,d in enumerate(dates):
            mon=d-pd.Timedelta(days=d.weekday())
            if per and per[-1][0]==mon: per[-1][2]=k
            else: per.append([mon,k,k])
    else: per=[[d,k,k] for k,d in enumerate(dates)]
    ND=len(per); CL=C0+ND-1; L_=L; SDL=L_(SD)
    for k,(d,_,_) in enumerate(per):
        c=ws.cell(3,C0+k,f'Week - {wkn(d)}'); c.font=Font(name=V,size=10,bold=True,color=NAVY); c.alignment=ctr
    hdr=['Key No.2','Sr No.','Short Name','Common Name','Material','Material Description','Customer','Production Line','Supplier','Headers','Stock Data']+[f'{d:%d-%m-%Y}' for d,_,_ in per]
    hfill=tfill(5,0.7999816888943144)
    for j,h in enumerate(hdr,1):
        c=ws.cell(4,j,h); c.font=Font(name=V,size=9 if j>=C0 else 10,bold=True,color=NAVY); c.fill=hfill; c.border=bdot; c.alignment=ctr
    ws.row_dimensions[4].height=24
    Dq=("'"+weekly_of.replace("'","''")+"'!") if weekly_of else ''
    DL=lambda k: L_(C0+k)
    r=5; stock_rows=[]; inv_rows=[]; short_rows=[]; inv_td={}      # inv_td: days-of-cover row -> that material's target days (colours)
    body=Font(name=V,size=10,color=NAVY); bold=Font(name=V,size=10,bold=True,color=NAVY); ital=Font(name=V,size=10,italic=True,color='FF595959')
    def ident(rr,b,sup,hd,cust=None,line=None):
        k_=b['key']+(sup if sup and sup!='All suppliers' and sup!=b.get('_sup1') else '')+(f'|{cust.strip()}' if cust and cust!='All customers' else '')+(f'|{line}' if line and not str(line).startswith('All lines') else '')
        for j,v in enumerate([k_,b['sr'],_short_lbl(b),b['common'],b['mat'],b['desc'],cust.strip() if cust else None,line,sup,hd],1): ws.cell(rr,j,v)
        ws.row_dimensions[rr].height=20.1
    # Cached results: formulas are saved together with the value Excel would calculate, so phone / e-mail previews (which do not
    # recalculate) show the numbers, and the colours are also painted as plain fills because previews ignore conditional formats.
    if cache is None: cache={}
    def C(rr,cc,v):
        if rr is not None: cache[(rr,cc)]=v
    def lerp(c1,c2,t): return ''.join(f'{round(int(c1[i:i+2],16)+(int(c2[i:i+2],16)-int(c1[i:i+2],16))*t):02X}' for i in (0,2,4))
    TD=target_days or LEAD_TIME_RM
    def scale_fill(v,td):                  # 0 red -> half the material's target yellow -> its target green (same as the live rule)
        if not isinstance(v,(int,float)): return None
        lo_,mid_,hi_=0,td/2,td
        if v<=lo_: col='FF0000'
        elif v>=hi_: col='00B050'
        elif v<=mid_: col=lerp('FF0000','FFFF00',(v-lo_)/(mid_-lo_))
        else: col=lerp('FFFF00','00B050',(v-mid_)/(hi_-mid_))
        return PatternFill('solid',fgColor='FF'+col)
    negfill=PatternFill('solid',fgColor='FFFFC7CE'); negfont=Font(name=V,size=10,color='FF9C0006')
    def paint(rr,cc,v,kind):
        if rr is None: return
        c=ws.cell(rr,cc)
        if kind=='inv':
            fl=scale_fill(v,inv_td.get(rr,TD))
            if fl: c.fill=fl
        elif kind in ('stock','short') and isinstance(v,(int,float)) and (v<0 if kind=='stock' else v>0): c.fill=negfill; c.font=negfont
    _sty={}; _ctr=Alignment(horizontal='center',vertical='center')
    from copy import copy as _cp
    def style(rows_,font_for):
        # same look as before, but each font / number-format combination is built once and re-used (much faster on 25,000+ cells)
        for rr in rows_:
            inv_=rr in inv_rows
            for j in range(1,CL+1):
                f=font_for(rr,j); nf=('0.0' if inv_ else acc) if j>=SD else None
                key=(id(f),nf); c=ws.cell(rr,j); sa=_sty.get(key)
                if sa is None:
                    c.border=bdot; c.alignment=_ctr; c.font=f
                    if nf: c.number_format=nf
                    _sty[key]=_cp(c._style)
                else: c._style=_cp(sa)
    for b in DS['blocks']:
        if b.get('cust_line'): custs=list(b['cust_line'].keys())          # (customer, production line)
        else: custs=[(c_,None) for c_ in (b.get('cust_plan') or {}).keys()]
        NC=len(custs)
        CPd=b.get('cust_plan') or {}
        def CPx(key):           # plan array of one Plan row: (customer, line) -> cust_line, else customer -> cust_plan
            if b.get('cust_line') and key in b['cust_line']: return b['cust_line'][key]
            cu_=key[0] if isinstance(key,tuple) else key
            return CPd.get(cu_) or [0.0]*len(dates)
        TDg=TD; TD=b.get('td') or TDg          # this material's target days (ABC / XYZ class); restored after the block
        opo=b.get('openpo') or [0.0]*len(dates)
        LTb=lt_of(b)                                 # Usage/Day window of 'Inv. Days' = this material's transit lead time
        has_opo=any(v for v in opo)                 # no Open PO: the Firm receipt row and the '+ Open PO' cover row would only repeat zeros / the row above
        rp=r; rc=list(range(r+1,r+1+NC))
        if has_opo: rs_,rt,ro,ri,rj=range(r+1+NC,r+6+NC)
        else: (rs_,rt,ri),ro,rj=range(r+1+NC,r+4+NC),None,None
        ra=None; rPR=(rj or ri)+1
        inv_lab=('Inv. Days (stock + in transit)','Inv. Days (stock + in transit + Open PO)')
        used_=[sp_ for sp_ in (b.get('sup') or []) if str(sp_.get('name') or '').upper()!='UNASSIGNED' and
               (float(sp_.get('stock') or 0)>0 or sum(sp_.get('transit') or [])>0 or sum(sp_.get('plan') or [])>0)]
        b['_sup1']=used_[0]['name'] if len(used_)==1 else None             # one real supplier: show its name, not 'All suppliers'
        supname=(b['_sup1'] or 'All suppliers') if b.get('sup') else None
        NN=lambda xs: [x_ for x_ in xs if x_ is not None]
        # standard MRP record (gross requirement -> receipts -> projected available balance -> net requirement -> planned orders)
        ident(rp,b,supname,'Gross requirement','All customers' if NC else None,'All lines' if any(ln for _,ln in custs) else None)
        for rr,(cu,ln) in zip(rc,custs):
            src_=(b.get('line_src') or {}).get((_norm(cu),_line_key(ln))) if ln else None
            if src_: src_=src_.replace(b['mat'],'own stock')
            ident(rr,b,(f'from: {src_}' if (src_ and src_ not in ('-','own stock')) else supname),'Gross requirement',cu,ln)
        for rr,hd in zip((rs_,rt,ro,ri,rj),('Projected available balance','In transit (PPC) - by arrival date','Firm receipt (Open PO, no vessel)')+inv_lab):
            if rr is not None: ident(rr,b,supname,hd)
        ident(rPR,b,supname,'Planned order receipt')
        stock_rows.append(rs_); inv_rows+=NN([ri,rj]); inv_td[ri]=TD
        if rj: inv_td[rj]=TD
        if weekly_of:
            for rr in NN([rp]+rc+[rs_,rt,ro,ri,rj,rPR]): ws.cell(rr,SD,f'={Dq}{SDL}{rr}')
            for w,(d,a,z) in enumerate(per):
                col=C0+w; A_,Z_=DL(a),DL(z); rng=lambda rr: f'{Dq}{A_}{rr}:{Z_}{rr}'
                for rr in NN([rp]+rc+[rt,ro]): ws.cell(rr,col,f'=SUM({rng(rr)})')
                for rr in NN([rs_,ri,rj]): ws.cell(rr,col,f'={Dq}{Z_}{rr}')     # end-of-week values from the daily MRP
                ws.cell(rPR,col,f'=SUM({rng(rPR)})')
        else:
            for rr in [rp]+rc: ws.cell(rr,SD,f'=SUM({L_(C0)}{rr}:{L_(CL)}{rr})')
            if b.get('count_stock'):   # new MRP: CS / PM stock = Today Qty on its own Stock Count sheet (live)
                ws.cell(rs_,SD,f"=SUMIFS('Stock Count'!${SC_COL}:${SC_COL},'Stock Count'!$A:$A,$E{rs_})")
            else: ws.cell(rs_,SD,round(b['stock'],6) if VALUES_ONLY else f"=SUMIFS('Stock Data'!$L$5:$L$3000,'Stock Data'!$A$5:$A$3000,$E{rs_})")   # RM: MB52 + JW lines on Stock Data
            ws.cell(rt,SD,f'=SUM({L_(C0)}{rt}:{L_(CL)}{rt})')
            if ro: ws.cell(ro,SD,f'=SUM({L_(C0)}{ro}:{L_(CL)}{ro})')      # total in transit, like Open PO ws.cell(ri,SD,0); ws.cell(rj,SD,0)
            ws.cell(rPR,SD,f'=SUM({L_(C0)}{rPR}:{L_(CL)}{rPR})')
            for k in range(ND):
                col=C0+k; cl=L_(col); pv=L_(col-1)
                ws.cell(rp,col,round(b['plan'][k],6))
                for rr,cu in zip(rc,custs): ws.cell(rr,col,round(CPx(cu)[k],6))
                ws.cell(rt,col,round(b['transit'][k],6))
                if ro: ws.cell(ro,col,round(opo[k],6))
                ws.cell(rs_,col,f'={pv}{rs_}-{cl}{rp}+{cl}{rt}')
                # forward-looking: stock / average demand of the next TD days (blank when no demand ahead)
                ce=L_(min(col+LTb-1,CL)); win=f'{cl}{rp}:{ce}{rp}'          # Usage/Day = average per production day over the lead time
                ws.cell(ri,col,f'=IF(COUNTIF({win},">0")=0,"",{cl}{rs_}/AVERAGEIF({win},">0"))')
                if rj: ws.cell(rj,col,f'=IF(COUNTIF({win},">0")=0,"",({cl}{rs_}+SUM(${L_(C0)}{ro}:{cl}{ro}))/AVERAGEIF({win},">0"))')
                ws.cell(rPR,col,round((b.get('por') or [0.0]*ND)[k],3))
        style(NN([rp]+rc+[rs_,rt,ro,ri,rj,rPR]),lambda rr,j: bold if rr in (ri,rj,rPR) and j>=10 else (ital if rr in rc else body))
        # line rows: red on the days this line is NOT covered (Customer Supplier, substitutes and line ranking applied), next to the receipts
        red_=PatternFill('solid',fgColor='FFC7CE')
        for rr,(cu,ln) in zip(rc,custs):
            k_=(b.get('line_stop') or {}).get(_line_key(ln)) if ln else None
            if k_ is None or k_>=len(dates): continue
            arr_=CPx((cu,ln))
            if weekly_of:
                for w,(d,a,z) in enumerate(per):
                    if z>=k_ and any(arr_[max(a,k_):z+1]): ws.cell(rr,C0+w).fill=red_
            else:
                for k in range(k_,ND):
                    if arr_[k]: ws.cell(rr,C0+k).fill=red_
            ws.cell(rr,8).fill=red_
        if not weekly_of:
            for k in range(min(ND,len(b.get('por') or []))):
                if b['por'][k]: ws.cell(rPR,C0+k).fill=PatternFill('solid',fgColor='FFEB9C')
        # ---- values for the cache (same rules as the formulas) ----
        nd=len(dates); P=list(b.get('plan') or [0.0]*nd); T=b['transit']; A=b['adhoc']; O=opo; S0=float(b.get('stock') or 0.0)
        S=[]; cur=S0
        for k in range(nd): cur=cur-P[k]+T[k]+A[k]; S.append(cur)
        cO=[]; acc_=0.0
        for k in range(nd): acc_+=O[k]; cO.append(acc_)
        win_avg=[]; win_sum=[]
        for k in range(nd):
            u_=usage_day(P,LTb,k); win_sum.append(1 if u_ else 0); win_avg.append(u_ or 0.0)
        if weekly_of:
            C(rp,SD,sum(P))
            for rr,cu in zip(rc,custs): C(rr,SD,sum(CPx(cu)))
            C(rs_,SD,S0); C(rt,SD,sum(b.get('transit') or [])); C(ro,SD,sum(O)); C(ri,SD,0); C(rj,SD,0); C(rPR,SD,sum(b.get('por') or []))
            for w,(d,a,z) in enumerate(per):
                col=C0+w; ps=sum(P[a:z+1])
                C(rp,col,ps); C(rt,col,sum(T[a:z+1])); C(ro,col,sum(O[a:z+1]))
                for rr,cu in zip(rc,custs): C(rr,col,sum(CPx(cu)[a:z+1]))
                C(rs_,col,S[z]); paint(rs_,col,S[z],'stock')
                v=(S[z]/win_avg[z]) if win_sum[z] else ''; C(ri,col,v); paint(ri,col,v,'inv')
                v=((S[z]+cO[z])/win_avg[z]) if win_sum[z] else ''; C(rj,col,v); paint(rj,col,v,'inv')
                C(rPR,col,sum((b.get('por') or [0.0]*nd)[a:z+1]))
        else:
            C(rp,SD,sum(P))
            for rr,cu in zip(rc,custs): C(rr,SD,sum(CPx(cu)))
            if b.get('count_stock'): C(rs_,SD,S0)
            C(rt,SD,sum(T)); C(ro,SD,sum(O)); C(rPR,SD,sum(b.get('por') or []))
            for k in range(nd):
                col=C0+k
                C(rs_,col,S[k]); paint(rs_,col,S[k],'stock')
                v=(S[k]/win_avg[k]) if win_sum[k] else ''; C(ri,col,v); paint(ri,col,v,'inv')
                v=((S[k]+cO[k])/win_avg[k]) if win_sum[k] else ''; C(rj,col,v); paint(rj,col,v,'inv')
        if rowmap is not None: rowmap.append(dict(b=b,rp=rp,rs=rs_,rt=rt,ro=ro,ra=ra,ri=ri,rj=rj,rc=list(rc),custs=list(custs),C0=C0,ND=ND))
        r=rPR+1; TD=TDg
        for sp in (b.get('sup') if multi_sup(b) else None) or []:                 # supplier detail (only when 2+ suppliers share the material)
            nd=len(dates); has_plan='plan' in sp; TDs=b.get('td') or TDg      # supplier rows use the material's own target days
            if has_plan:     # customers mapped to this supplier on the 'Customer Supplier' sheet: own plan, projected stock and cover
                sP,s1,s2,s3,sI=r,r+1,r+2,r+3,r+4
                ident(sP,b,sp['name'],sp.get('plan_label','Plan (mapped customers)'),sp.get('cust_text') or None,'All lines' if b.get('cust_line') else None)
                for rr,hd in zip((s1,s2,s3,sI),('Stock','In transit (PPC) - by arrival date','Firm receipt (Open PO, no vessel)','Inv. Days')): ident(rr,b,sp['name'],hd)
                style((sP,s1,s2,s3,sI),lambda rr,j: bold if rr==sI and j>=10 else ital)
                Pq=sp['plan']; Ss=[]; cur=sp['stock']
                for k in range(nd): cur=cur-Pq[k]+sp['transit'][k]; Ss.append(cur)
                wavg=[usage_day(Pq,LTb,k) or 0.0 for k in range(nd)]; wsum=[1 if v_ else 0 for v_ in wavg]
                stock_rows.append(s1); inv_rows.append(sI); inv_td[sI]=TDs
                if weekly_of:
                    for rr in (sP,s1,s2,s3,sI): ws.cell(rr,SD,f'={Dq}{SDL}{rr}')
                    for w,(d,a,z) in enumerate(per):
                        for rr in (sP,s2,s3): ws.cell(rr,C0+w,f'=SUM({Dq}{DL(a)}{rr}:{DL(z)}{rr})')
                        for rr in (s1,sI): ws.cell(rr,C0+w,f'={Dq}{DL(z)}{rr}')
                        C(sP,C0+w,sum(Pq[a:z+1])); C(s2,C0+w,sum(sp['transit'][a:z+1])); C(s3,C0+w,sum(sp['openpo'][a:z+1]))
                        C(s1,C0+w,Ss[z]); paint(s1,C0+w,Ss[z],'stock')
                        v=(Ss[z]/wavg[z]) if wsum[z] else ''; C(sI,C0+w,v); paint(sI,C0+w,v,'inv')
                    C(sP,SD,sum(Pq)); C(s1,SD,sp['stock']); C(s2,SD,sum(sp['transit'])); C(s3,SD,sum(sp['openpo'])); C(sI,SD,0)
                else:
                    ws.cell(sP,SD,f'=SUM({L_(C0)}{sP}:{L_(CL)}{sP})'); ws.cell(s1,SD,round(sp['stock'],6)); ws.cell(sI,SD,0)
                    for rr,key in ((s2,'transit'),(s3,'openpo')): ws.cell(rr,SD,f'=SUM({L_(C0)}{rr}:{L_(CL)}{rr})')
                    for k in range(ND):
                        col=C0+k; cl=L_(col); pv=L_(col-1); ce=L_(min(col+TDs-1,CL)); win=f'{cl}{sP}:{ce}{sP}'
                        ws.cell(sP,col,round(Pq[k],6)); ws.cell(s2,col,round(sp['transit'][k],6)); ws.cell(s3,col,round(sp['openpo'][k],6))
                        ws.cell(s1,col,f'={pv}{s1}-{cl}{sP}+{cl}{s2}')
                        ws.cell(sI,col,f'=IF(SUM({win})=0,"",{cl}{s1}/AVERAGE({win}))')
                        C(s1,col,Ss[k]); paint(s1,col,Ss[k],'stock')
                        v=(Ss[k]/wavg[k]) if wsum[k] else ''; C(sI,col,v); paint(sI,col,v,'inv')
                    C(sP,SD,sum(Pq)); C(s2,SD,sum(sp['transit'])); C(s3,SD,sum(sp['openpo']))
                r+=5; continue
            s1,s2,s3=r,r+1,r+2
            for rr,hd in zip((s1,s2,s3),('Stock','In transit (PPC) - by arrival date','Firm receipt (Open PO, no vessel)')): ident(rr,b,sp['name'],hd)
            if weekly_of:
                for rr in (s1,s2,s3): ws.cell(rr,SD,f'={Dq}{SDL}{rr}')
                for w,(d,a,z) in enumerate(per):
                    for rr in (s2,s3): ws.cell(rr,C0+w,f'=SUM({Dq}{DL(a)}{rr}:{DL(z)}{rr})')
            else:
                ws.cell(s1,SD,round(sp['stock'],6))
                for rr,key in ((s2,'transit'),(s3,'openpo')):
                    ws.cell(rr,SD,f'=SUM({L_(C0)}{rr}:{L_(CL)}{rr})')
                    for k in range(ND): ws.cell(rr,C0+k,round(sp[key][k],6))
            style((s1,s2,s3),lambda rr,j: ital)
            if weekly_of:
                C(s1,SD,sp['stock']); C(s2,SD,sum(sp['transit'])); C(s3,SD,sum(sp['openpo']))
                for w,(d,a,z) in enumerate(per): C(s2,C0+w,sum(sp['transit'][a:z+1])); C(s3,C0+w,sum(sp['openpo'][a:z+1]))
            else: C(s2,SD,sum(sp['transit'])); C(s3,SD,sum(sp['openpo']))
            r+=3
    last=r-1
    ws.add_table(Table(displayName=table,ref=f'A4:{L_(CL)}{last}',
                       tableStyleInfo=TableStyleInfo(name='TableStyleLight10',showRowStripes=True,showColumnStripes=False)))
    red=DifferentialStyle(font=Font(color='FF9C0006'),fill=PatternFill(bgColor='FFFFC7CE',fill_type='solid'))
    ws.conditional_formatting.add(' '.join(f'{L_(C0)}{x}:{L_(CL)}{x}' for x in stock_rows),Rule(type='cellIs',operator='lessThan',formula=['0'],dxf=red))
    for td_ in sorted(set(inv_td.values())):     # Inv. Days: 0 red / half target yellow / target green - one rule per target (Settings)
        rows_td=[x for x in inv_rows if inv_td.get(x,TD)==td_]
        ws.conditional_formatting.add(' '.join(f'{L_(C0)}{x}:{L_(CL)}{x}' for x in rows_td),
            ColorScaleRule(start_type='num',start_value=0,start_color='FFFF0000',mid_type='num',mid_value=td_/2,mid_color='FFFFFF00',
                           end_type='num',end_value=td_,end_color='FF00B050'))
    # The painted colours (for phone previews) must never outlive an edit in Excel: every cell in these rows is always covered by a
    # live rule - stock >= 0 and blank Inv. Days get a plain white rule that overrides any painted colour.
    plain=DifferentialStyle(font=Font(color='FF002060'),fill=PatternFill(bgColor='FFFFFFFF',fill_type='solid'))
    sr_=' '.join(f'{L_(C0)}{x}:{L_(CL)}{x}' for x in stock_rows); ir_=' '.join(f'{L_(C0)}{x}:{L_(CL)}{x}' for x in inv_rows)
    if stock_rows: ws.conditional_formatting.add(sr_,Rule(type='cellIs',operator='greaterThanOrEqual',formula=['0'],dxf=plain))
    if inv_rows: ws.conditional_formatting.add(ir_,Rule(type='expression',formula=[f'NOT(ISNUMBER({L_(C0)}{inv_rows[0]}))'],dxf=plain))
    if short_rows:
        xr_=' '.join(f'{L_(C0)}{x}:{L_(CL)}{x}' for x in short_rows)
        ws.conditional_formatting.add(xr_,Rule(type='cellIs',operator='greaterThan',formula=['0'],dxf=red))
        ws.conditional_formatting.add(xr_,Rule(type='cellIs',operator='lessThanOrEqual',formula=['0'],dxf=plain))
    for col,w in zip('ABCDEFGHIJKL',[25.68,8,27.26,29,19.16,51.58,20,16,16,14.58,17,13]): ws.column_dimensions[col].width=w
    for col in range(C0,CL+1): ws.column_dimensions[L_(col)].width=13
    ws.column_dimensions['A'].hidden=True
    ws.freeze_panes=f'{L_(C0)}5'; ws.sheet_view.zoomScale=80; ws.sheet_view.showGridLines=True; set_print(ws,title)
    wb.save(path)
    return len(DS['blocks'])

def graft_sheet(pkg_path, src_xlsx, sheet_name, cache=None, hidden=False, tab_color=None):
    """Put the first sheet of src_xlsx (with its table, conditional formats and styles) into the workbook pkg_path as sheet_name,
       replacing a sheet of that name if there is one. Every other part of pkg_path is copied byte for byte."""
    import zipfile, posixpath
    print(f'  [{time.time()-_T0:5.0f} s] sheet: {sheet_name}',flush=True)     # progress, so a slow step is visible
    from lxml import etree
    M='http://schemas.openxmlformats.org/spreadsheetml/2006/main'; R='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    PR='http://schemas.openxmlformats.org/package/2006/relationships'; CT='http://schemas.openxmlformats.org/package/2006/content-types'
    q=lambda t:'{%s}%s'%(M,t)
    WS_T='application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'; TB_T='application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml'
    with zipfile.ZipFile(pkg_path) as z: infos=z.infolist(); parts={i.filename:z.read(i.filename) for i in infos}
    with zipfile.ZipFile(src_xlsx) as z: S={n:z.read(n) for n in z.namelist()}
    X=lambda b: etree.fromstring(b)
    def rels_path(p): d,f=posixpath.split(p); return f'{d}/_rels/{f}.rels'
    def resolve(base,target): return target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join(posixpath.dirname(base),target))
    wbx=X(parts['xl/workbook.xml']); wrels=X(parts['xl/_rels/workbook.xml.rels']); ct=X(parts['[Content_Types].xml'])
    def ct_set(part,typ):
        for o in ct.findall('{%s}Override'%CT):
            if o.get('PartName')=='/'+part: ct.remove(o)
        if typ: e=etree.SubElement(ct,'{%s}Override'%CT); e.set('PartName','/'+part); e.set('ContentType',typ)
    # ---- shared strings ----
    if 'xl/sharedStrings.xml' not in parts:
        parts['xl/sharedStrings.xml']=f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="{M}" count="0" uniqueCount="0"/>'.encode()
        e=etree.SubElement(wrels,'{%s}Relationship'%PR); e.set('Id','rIdSst1'); e.set('Type',R+'/sharedStrings'); e.set('Target','sharedStrings.xml')
        ct_set('xl/sharedStrings.xml','application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml')
    tsst=X(parts['xl/sharedStrings.xml']); tsis=tsst.findall(q('si'))
    plain={}
    for i,si in enumerate(tsis):
        if si.find(q('r')) is None: plain.setdefault(''.join(si.itertext()),i)
    ssst=[''.join(si.itertext()) for si in X(S['xl/sharedStrings.xml']).findall(q('si'))] if 'xl/sharedStrings.xml' in S else []
    smap={}
    for i,t in enumerate(ssst):
        if t not in plain:
            si=etree.SubElement(tsst,q('si')); te=etree.SubElement(si,q('t')); te.text=t
            if t!=t.strip(): te.set('{http://www.w3.org/XML/1998/namespace}space','preserve')
            plain[t]=len(tsst.findall(q('si')))-1
        smap[i]=plain[t]
    tsst.set('uniqueCount',str(len(tsst.findall(q('si')))))
    # ---- styles ----
    tst=X(parts['xl/styles.xml']); sst_=X(S['xl/styles.xml'])
    ORDER=['numFmts','fonts','fills','borders','cellStyleXfs','cellXfs','cellStyles','dxfs','tableStyles','colors','extLst']
    def coll(root,name):
        e=root.find(q(name))
        if e is None:
            e=etree.Element(q(name)); after=None
            for nm in ORDER[:ORDER.index(name)]:
                if root.find(q(nm)) is not None: after=root.find(q(nm))
            if after is None: root.insert(0,e)
            else: after.addnext(e)
        return e
    def canon(e):                       # namespace-free, attribute-order-free fingerprint of a style element
        def num(v):
            try: return repr(round(float(v),6))
            except ValueError: return v
        tag=etree.QName(e).localname
        return (tag,tuple(sorted((etree.QName(k).localname,num(v)) for k,v in e.attrib.items())),tuple(canon(c) for c in e if isinstance(c.tag,str)))
    _idx={}                             # fingerprint -> index, built once per style collection (was re-computed for every lookup)
    def merge(name,elem):
        c=coll(tst,name)
        if name not in _idx:
            d_={}
            for i,k in enumerate(list(c)): d_.setdefault(canon(k),i)
            _idx[name]=d_
        key=canon(elem); d_=_idx[name]
        if key in d_: return d_[key]
        c.append(elem); c.set('count',str(len(c))); d_[key]=len(c)-1; return len(c)-1
    import copy
    # number formats
    tnf=coll(tst,'numFmts'); code2id={n.get('formatCode'):int(n.get('numFmtId')) for n in tnf}
    nfmap={}
    for n in (sst_.find(q('numFmts')) if sst_.find(q('numFmts')) is not None else []):
        fid=int(n.get('numFmtId')); code=n.get('formatCode')
        if code not in code2id:
            new=max([163]+list(code2id.values()))+1; e=etree.SubElement(tnf,q('numFmt')); e.set('numFmtId',str(new)); e.set('formatCode',code); code2id[code]=new
        nfmap[fid]=code2id[code]
    tnf.set('count',str(len(tnf)))
    if len(tnf)==0: tst.remove(tnf)
    sub={nm:[merge(nm,copy.deepcopy(e)) for e in sst_.find(q(nm))] for nm in ('fonts','fills','borders')}
    xfmap=[]
    for xf in sst_.find(q('cellXfs')):
        e=copy.deepcopy(xf)
        for attr,nm in (('fontId','fonts'),('fillId','fills'),('borderId','borders')):
            if e.get(attr) is not None: e.set(attr,str(sub[nm][int(e.get(attr))]))
        if e.get('numFmtId') is not None and int(e.get('numFmtId')) in nfmap: e.set('numFmtId',str(nfmap[int(e.get('numFmtId'))]))
        e.set('xfId','0')
        xfmap.append(merge('cellXfs',e))
    sdx=sst_.find(q('dxfs')); dxmap=[merge('dxfs',copy.deepcopy(d)) for d in (sdx if sdx is not None else [])]
    # ---- the sheet itself ----
    sh=X(S['xl/worksheets/sheet1.xml'])
    def _rc(ref):
        m_=re.match(r'([A-Z]+)(\d+)',ref); n_=0
        for ch_ in m_.group(1): n_=n_*26+ord(ch_)-64
        return int(m_.group(2)),n_
    for c in sh.iter(q('c')):
        if cache and c.find(q('f')) is not None:          # store the calculated value next to the formula
            key=_rc(c.get('r'))
            if key in cache:
                v_=cache[key]
                for old in c.findall(q('v')): c.remove(old)
                if VALUES_ONLY:                            # plain value only (faster file)
                    for f_ in c.findall(q('f')): c.remove(f_)
                    if v_ is None or v_=='': 
                        if 't' in c.attrib: del c.attrib['t']
                        continue
                ve=etree.SubElement(c,q('v'))
                if isinstance(v_,str): c.set('t','str'); ve.text=v_
                elif isinstance(v_,bool): c.set('t','b'); ve.text='1' if v_ else '0'
                else:
                    if 't' in c.attrib: del c.attrib['t']
                    ve.text=repr(round(float(v_),9))
        if c.get('s') is not None: c.set('s',str(xfmap[int(c.get('s'))]))
        if c.get('t')=='s': v=c.find(q('v')); v.text=str(smap[int(v.text)])
    for rw in sh.iter(q('row')):
        if rw.get('s') is not None: rw.set('s',str(xfmap[int(rw.get('s'))]))
    for co in sh.iter(q('col')):
        if co.get('style') is not None: co.set('style',str(xfmap[int(co.get('style'))]))
    for cf in sh.iter(q('cfRule')):
        if cf.get('dxfId') is not None: cf.set('dxfId',str(dxmap[int(cf.get('dxfId'))]))
    if tab_color:                                     # coloured sheet tab (sheetPr must be the first child of the worksheet)
        sp=sh.find(q('sheetPr'))
        if sp is None: sp=etree.Element(q('sheetPr')); sh.insert(0,sp)
        for old in sp.findall(q('tabColor')): sp.remove(old)
        tc=etree.Element(q('tabColor')); tc.set('rgb','FF'+tab_color); sp.insert(0,tc)
    for sv in sh.iter(q('sheetView')):
        if 'tabSelected' in sv.attrib: del sv.attrib['tabSelected']
    # ---- where it goes: replace the existing sheet of that name, or add a new one ----
    sheets=wbx.find(q('sheets')); rid={r_.get('Id'):r_ for r_ in wrels}
    old=next((s for s in sheets if s.get('name')==sheet_name),None)
    if old is not None:
        wpath=resolve('xl/workbook.xml',rid[old.get('{%s}id'%R)].get('Target'))
        rp=rels_path(wpath)
        if rp in parts:                                  # drop its old tables
            for r_ in X(parts[rp]):
                if r_.get('Type','').endswith('/table'):
                    tp_=resolve(wpath,r_.get('Target')); parts.pop(tp_,None); ct_set(tp_,None)
            parts.pop(rp,None)
    else:
        n=1
        while f'xl/worksheets/sheet{n}.xml' in parts: n+=1
        wpath=f'xl/worksheets/sheet{n}.xml'
        ids={r_.get('Id') for r_ in wrels}; k=1
        while f'rId{k}' in ids: k+=1
        e=etree.SubElement(wrels,'{%s}Relationship'%PR); e.set('Id',f'rId{k}'); e.set('Type',R+'/worksheet'); e.set('Target',wpath[3:])
        s=etree.SubElement(sheets,q('sheet')); s.set('name',sheet_name)
        s.set('sheetId',str(max(int(x.get('sheetId')) for x in sheets if x.get('sheetId'))+1)); s.set('{%s}id'%R,f'rId{k}')
        ct_set(wpath,WS_T)
    ent=next(s_ for s_ in sheets if s_.get('name')==sheet_name)
    if hidden: ent.set('state','hidden')
    elif 'state' in ent.attrib: del ent.attrib['state']
    # ---- table ----
    srels=X(S['xl/worksheets/_rels/sheet1.xml.rels']) if 'xl/worksheets/_rels/sheet1.xml.rels' in S else None
    for n_,b_ in S.items():
        if n_.startswith('xl/tables/'):
            names=[c.get('name','').strip().lower() for c in X(b_).iter(q('tableColumn'))]
            dup=sorted({x for x in names if names.count(x)>1})
            if dup: die(f"Internal check: table on sheet '{sheet_name}' has repeated column names {dup} - not saved, so Excel will not need to repair the file.")
    if srels is not None:
        tids=[]; tnames=set()
        for p_,b_ in parts.items():
            if p_.startswith('xl/tables/') and p_.endswith('.xml'):
                t_=X(b_); tids.append(int(t_.get('id'))); tnames.add(t_.get('name'))
        nrels=etree.Element('{%s}Relationships'%PR)
        for r_ in srels:
            if not r_.get('Type','').endswith('/table'): continue
            tb=X(S[resolve('xl/worksheets/sheet1.xml',r_.get('Target'))])
            n=1
            while f'xl/tables/table{n}.xml' in parts: n+=1
            tpath=f'xl/tables/table{n}.xml'; tid=max([0]+tids)+1; tids.append(tid)
            nm=tb.get('name'); base=nm; k=2
            while nm in tnames: nm=f'{base}{k}'; k+=1
            tnames.add(nm); tb.set('id',str(tid)); tb.set('name',nm); tb.set('displayName',nm)
            parts[tpath]=etree.tostring(tb,xml_declaration=True,encoding='UTF-8',standalone=True); ct_set(tpath,TB_T)
            e=etree.SubElement(nrels,'{%s}Relationship'%PR); e.set('Id',r_.get('Id')); e.set('Type',r_.get('Type')); e.set('Target',posixpath.relpath(tpath,posixpath.dirname(wpath)))
        if len(nrels): parts[rels_path(wpath)]=etree.tostring(nrels,xml_declaration=True,encoding='UTF-8',standalone=True)
    parts[wpath]=etree.tostring(sh,xml_declaration=True,encoding='UTF-8',standalone=True)
    # calcChain is rebuilt by Excel; removing it avoids 'repair' messages when formula cells change
    if 'xl/calcChain.xml' in parts:
        parts.pop('xl/calcChain.xml'); ct_set('xl/calcChain.xml',None)
        for r_ in list(wrels):
            if r_.get('Type','').endswith('/calcChain'): wrels.remove(r_)
    calc=wbx.find(q('calcPr'))
    if calc is None: calc=etree.SubElement(wbx,q('calcPr'))
    calc.set('fullCalcOnLoad','1')
    ser=lambda e: etree.tostring(e,xml_declaration=True,encoding='UTF-8',standalone=True)
    parts.update({'xl/workbook.xml':ser(wbx),'xl/_rels/workbook.xml.rels':ser(wrels),'[Content_Types].xml':ser(ct),'xl/styles.xml':ser(tst),'xl/sharedStrings.xml':ser(tsst)})
    tmp=pkg_path+'.tmp'; done=set()
    with zipfile.ZipFile(tmp,'w',zipfile.ZIP_DEFLATED) as zout:
        zout.writestr('[Content_Types].xml',parts['[Content_Types].xml']); done.add('[Content_Types].xml')
        for i in infos:
            if i.filename in parts and i.filename not in done: zout.writestr(i,parts[i.filename]); done.add(i.filename)
        for n_,b_ in parts.items():
            if n_ not in done: zout.writestr(n_,b_)
    try: os.replace(tmp,pkg_path)
    except PermissionError: os.remove(tmp); die(f'Could not save {pkg_path} - is it open in Excel? Close it and run again.')

def attach_wsa(DS, BW, extra=()):
    """Adds WSA Open PO by day and the glass supplier detail to the MIS-based blocks. Returns (n_open, n_glass)."""
    d0=DS['dates'][0]; nd=len(DS['dates']); od=BW.get('open_daily',{}); n_open=0
    for b in DS['blocks']:
        v=[0.0]*nd
        for d_,q_ in od.get(b['common'],{}).items():
            k=max(0,(d_-d0).days)              # plant date already past -> shown on the first day (overdue)
            if k<nd: v[k]+=q_
        b['openpo']=v; n_open+=any(v)
    def _sup_key(name,makes):          # map a full supplier name to its Make (e.g. 'FLAT (VIETNAM) CO., LIMITED' -> FLAT)
        u=name.upper()
        for mk_ in sorted(makes,key=len,reverse=True):
            if mk_ and mk_ in u: return mk_
        for mk_ in ('FLAT','XINYI','KIBING'):
            if mk_ in u: return mk_
        return name
    # stock by supplier from the MB52 BATCHES (your Batch Master); MIS RM Status 'Make' only when no batch can be named
    bmk={}
    if MIS_PATH_[0] and (BATCH_MASTER[0] or BATCH_MASTER[1]):
        pl_={str(DS.get('plant') or '').upper(),str(DS.get('plant') or '').upper()+' WAREHOUSE'}
        al_=DS.get('alias') or {}; fa_=DS.get('alias_factor') or {}
        try:
            for r_ in list(wb_ro(MIS_PATH_[0])['MB52'].iter_rows(values_only=True))[1:]:
                if not r_ or len(r_)<15 or str(r_[13] or '').strip().upper() not in pl_ or not isinstance(r_[14],(int,float)): continue
                c_=str(r_[2] or '').strip(); m_=al_.get(c_,c_); f_=fa_.get(c_,1.0) if m_!=c_ else 1.0
                mk_=batch_make(r_[3])
                key_=re.split(r'[-\s(]',mk_.strip().upper())[0] if mk_ else None
                bmk.setdefault(m_,{}); bmk[m_][key_]=bmk[m_].get(key_,0.0)+r_[14]*f_
        except Exception as ex_: print(f'NOTE: batch split of glass stock skipped ({type(ex_).__name__})')
    n_gl=0
    for b in DS['blocks']:
        if 'glass' not in str(b['common']).lower() and b['mat'] not in extra: continue
        bm_=bmk.get(b['mat']) or {}
        if bm_ and sum(v for k,v in bm_.items() if k)>0.5*max(1.0,sum(bm_.values())):    # most stock named by batch -> use the batches
            makes={k:v for k,v in bm_.items() if k}
        else: makes=dict(DS['mk_stock'].get(b['mat'],{}))
        sup={}
        def S(k): return sup.setdefault(k,dict(name=k,stock=0.0,transit=[0.0]*nd,openpo=[0.0]*nd))
        for mk_,qv in makes.items(): S(mk_)['stock']+=qv
        gap=own_stock(b)-sum(makes.values())                 # own stock only: JW at the OEM never goes into a supplier row
        if abs(gap)>max(0.5,0.005*abs(b['stock'])): S('Unassigned')['stock']+=gap   # ignore rounding-size gaps
        for sn,v in DS['trans_sup'].get(b['mat'],{}).items():
            t_=S(_sup_key(sn,makes))['transit']
            for k in range(nd): t_[k]+=v[k]
        for sn,dd_ in BW.get('open_sup',{}).get(b['common'],{}).items():
            po_,_,shp_=sn.partition('|')
            nm_=DS['po_sup'].get(po_) or DS['po_sup'].get(shp_) or shp_
            o_=S(_sup_key(nm_,makes))['openpo']
            for d_,q_ in dd_.items():
                k=max(0,(d_-d0).days)
                if k<nd: o_[k]+=q_
        rows_=[x for x in sup.values() if abs(x['stock'])>0.5 or any(x['transit']) or any(x['openpo'])]
        rows_.sort(key=lambda x:(x['name']=='Unassigned',-(x['stock']+sum(x['transit'])+sum(x['openpo']))))
        if rows_: b['sup']=rows_; n_gl+=1
        b['sup_sap']={k_:float(v_) for k_,v_ in makes.items()}          # SAP stock (MB52) by supplier, before any count (Stock Check supplier rows)
    return n_open,n_gl


