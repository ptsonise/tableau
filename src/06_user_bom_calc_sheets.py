# ===================== YOUR BOM (MRP.xlsx 'BOM' sheet) - new layout =====================
BOM_HDR_MAP={'customer name':'Customer','customer':'Customer','material code':'Component','component':'Component',
             'material description':'Component desciption','component desciption':'Component desciption','component description':'Component desciption',
             'per module consumption':'Quantity','quantity':'Quantity','uom':'UoM','module model':'Module Model','material group':'Material group',
             'module type':'Module type','valid from':'Valid From','valid to':'Valid to','mis bom name':'MIS BOM Name',
             'fg material':'FG Material','material':'Component','description':'Component desciption','qty per module':'Quantity','line #':'Line #',
             'customers':'Customer group','supplier code':'Supplier Code'}
PP_MODELS={}       # Production Plan customer -> Module Model (to link the BOM's FG Material / customer group to your customers)
def _model_prefix(m):
    m=str(m or '').strip().upper(); i=m.find('-MC4')
    return m[:i+4] if i>0 else m[:22]
def load_pp_models(mrp_path):
    PP_MODELS.clear()
    try:
        for r in pp_rows_std(wb_ro(mrp_path)['Production Plan']):
            if r and r[3] and len(r)>7 and isinstance(r[7],str) and r[7].upper().startswith('BIN'): PP_MODELS.setdefault(' '.join(str(r[3]).split()),r[7])
    except Exception as _e_: _swallow(3235,_e_); pass
def bom_rows_std(rows):
    """Rows of your BOM sheet with standard field names, old layout (Customer / Component / Quantity) or new layout
       (Customer Name / Material Code / Per module Consumption). Returns (list of dicts) or None when no header is found."""
    for i,r in enumerate(rows[:10]):
        hl=[' '.join(str(x or '').split()).lower() for x in r]
        if ('component' in hl and 'quantity' in hl) or ('material code' in hl and 'per module consumption' in hl) or ('material' in hl and 'qty per module' in hl):
            idx={}
            for j,h in enumerate(hl):
                k=BOM_HDR_MAP.get(h)
                if k and k not in idx: idx[k]=j
            out=[]
            for r2 in rows[i+1:]:
                d={k:(r2[j] if j<len(r2) else None) for k,j in idx.items()}
                if not d.get('Component'): continue
                d['Component']=str(d['Component']).strip()
                if isinstance(d.get('Customer'),str): d['Customer']=' '.join(d['Customer'].split())
                if d.get('Customer') or not d.get('FG Material'): out.append(d); continue
                # PPC 'Order' layout: one row per FG Material -> every Production Plan customer with that module model
                if not PP_MODELS and SETTINGS_SRC and SETTINGS_SRC[0]: load_pp_models(SETTINGS_SRC[0])
                pre=_model_prefix(d['FG Material']); hit=False
                for cu,mdl in PP_MODELS.items():
                    if _model_prefix(mdl)==pre: e=dict(d); e['Customer']=cu; e['Module Model']=mdl; out.append(e); hit=True
                if not hit: e=dict(d); e['Customer']=d.get('Customer group') or d['FG Material']; e['Module Model']=d['FG Material']; out.append(e)
            return out
    return None

UBOM_DEMAND=True   # False: raw-material demand = MIS Standard BOM; your BOM (MRP.xlsx) = ALTERNATIVE BOM, used when a standard material is short
UBOM_RM={}         # your BOM, raw materials: normalised customer -> {code: [per-module quantities]} (a code listed twice = two layers, added up)
UBOM_ALT={}        # customer -> {alternative code: its standard code} from your BOM 'Line #' = '<standard>_Alt'
def load_user_bom_rm(mrp_path):
    """Raw-material lines of your BOM (codes RM...), for 'your BOM first, MIS BOM if the customer is not in it'."""
    UBOM_RM.clear(); UBOM_ALT.clear(); META={}
    try:
        rows=[list(r) for r in wb_ro(mrp_path)['BOM'].iter_rows(values_only=True)]
    except Exception as _e_: _swallow(3270,_e_); return META
    for d in bom_rows_std(rows) or []:
        c=d['Component']
        if not c.upper().startswith('RM') or not isinstance(d.get('Quantity'),(int,float)) or not d.get('Customer'): continue
        ln_=str(d.get('Line #') or '').strip()
        if ln_.lower().endswith('_alt'):                                 # your BOM marks an alternative: Line # = '<standard code>_Alt'
            UBOM_ALT.setdefault(_norm(d['Customer']),{})[c]=ln_[:-4].strip()
        UBOM_RM.setdefault(_norm(d['Customer']),{}).setdefault(c,[]).append((float(d['Quantity']),str(d.get('UoM') or '').strip().upper()))
        META.setdefault(c,dict(desc=d.get('Component desciption'),uom=d.get('UoM'),group=d.get('Material group'),mtype=d.get('Module type')))
    return META

PACK_KG={}         # material (common name) -> KG per pack from the WSA sheet unit, e.g. 'DRUM(270KG/DRUM)' -> 270
UBOM_UNIT_ISSUES=[]   # (customer, code, your unit, MRP unit) that could not be converted
def ubom_rate(ucust_norm, block_code, alias, factor, buom=None, bname=None):
    """Your BOM usage per module of a planned material for one customer. Codes of your BOM that are alternates of the same
       material (alternate codes / roll lengths, e.g. Busbar 4 x 0.4 and reflective 4 x 0.36) count ONCE (the largest), a code
       listed twice (front + back layer) counts twice. None = customer not in your BOM."""
    rows=UBOM_RM.get(ucust_norm) if UBOM_DEMAND else None
    if rows is None: return None
    best=0.0; bu=str(buom or '').strip().upper()
    alts_={a for std_,al_ in ALT_BOM.items() for a,cs_ in al_.items()     # an alternative carries no demand ONLY when this customer's BOM
           if any(_norm(x)==ucust_norm for x in cs_) and std_ in rows}         # also lists its standard code; listed alone it IS the demand
    alts_|={a for a,std_ in (UBOM_ALT.get(ucust_norm) or {}).items() if std_ in rows}     # your BOM's '_Alt' rows
    for c,qs in rows.items():
        if alias.get(c,c)!=block_code or c in alts_: continue
        tot=0.0
        for q,u in qs:
            if bu and u and u!=bu and not (c!=block_code and c in factor):     # your unit differs from the MRP unit
                if u=='KG' and PACK_KG.get(bname): q=q/PACK_KG[bname]          # e.g. RTV 0.39 KG per module / 270 KG per drum
                elif (ucust_norm,c,u,bu) not in UBOM_UNIT_ISSUES: UBOM_UNIT_ISSUES.append((ucust_norm,c,u,bu))
            tot+=q
        best=max(best,tot*(factor.get(c,1.0) if c!=block_code else 1.0))
    return best


def ubom_checks(mrp_path, mis_path):
    """Your BOM (raw materials) vs MIS master data: a code whose description has numbers (size, GSM, thickness...) that the MIS
       description of that code does not have = probably the wrong code; units your BOM uses that could not be converted."""
    I=[]; cn={}
    try:
        for r in wb_ro(mis_path)['Common Name'].iter_rows(min_row=2,values_only=True):
            if r and r[0]: cn[str(r[0]).strip()]=(r[1],r[3])
        rows=[list(r) for r in wb_ro(mrp_path)['BOM'].iter_rows(values_only=True)]
    except Exception as _e_: _swallow(3313,_e_); return I
    nums=lambda t: set(re.findall(r'\d+(?:\.\d+)?',str(t or '').upper().replace(',','.')))
    seen=set()
    for d in bom_rows_std(rows) or []:
        c=d['Component']
        if not c.upper().startswith('RM') or c not in cn: continue
        mine=nums(d.get('Component desciption')); mis_=nums(cn[c][0])
        extra=sorted(mine-mis_,key=lambda x:float(x))
        if extra and (c,tuple(extra)) not in seen:
            seen.add((c,tuple(extra))); cus=sorted({' '.join(str(x.get('Customer') or '').split()) for x in (bom_rows_std(rows) or []) if x['Component']==c and nums(x.get('Component desciption'))==mine})
            I.append(dict(area='Your BOM',check='Code and description do not match (MIS master)',sev='High',material=cn[c][1] or c,code=c,
                          a_lab='Your BOM description',a=' '.join(str(d.get('Component desciption') or '').split()),b_lab='MIS description of this code',b=' '.join(str(cn[c][0] or '').split()),diff=None,
                          detail=f"{', '.join(cus)}: your BOM has {', '.join(extra)} in the description, the MIS description of {c} does not - the code may be wrong. The MRP uses the CODE.",
                          action='Correct the code (or the description) in the BOM sheet of MRP.xlsx'))
    for cu,c,u,bu in UBOM_UNIT_ISSUES:
        I.append(dict(area='Your BOM',check='Unit not converted',sev='High',material=c,code=c,a_lab='Your BOM unit',a=u,b_lab='MRP unit',b=bu,diff=None,
                      detail=f'{cu}: your BOM gives {c} per module in {u}, the MRP counts {bu} - used without conversion',
                      action='Put the pack size on the WSA sheet unit (e.g. DRUM(270KG/DRUM)) or give the BOM in the MRP unit'))
    return I


ALT_BOM={}         # standard code -> {alternative code: set(customers)} from your alternative BOM
ALT_EXTRA={}       # code only in your BOM (no standard of that group) -> set(customers)
def build_alt_bom(mis_path, alias):
    """Your BOM = alternative BOM. For each customer: codes of your BOM that are not in the MIS Standard BOM of its module are
       alternatives of the standard code of the same material group (your Material Group, else the MIS group); several standard
       codes in one group -> the one the alternate is linked to (alias / same Common Name). No standard code of that group -> extra."""
    ALT_BOM.clear(); ALT_EXTRA.clear()
    try: rows=list(wb_ro(mis_path)['Standard BOM'].iter_rows(values_only=True))[1:]
    except Exception as _e_: _swallow(3342,_e_); return
    sb={}; mgrp={}
    for r in rows:
        if r and len(r)>8 and r[2] and r[3]:
            sb.setdefault(str(r[2]).strip(),{})[str(r[3]).strip()]=str(r[8] or ''); mgrp.setdefault(str(r[3]).strip(),str(r[8] or ''))
    def grp(c,mis_g=''):
        g_=mis_g or mgrp.get(c)                                 # MIS Standard BOM group of the code (any module) first
        if g_: return BOM_GROUP_MAP.get(str(g_).strip().lower(),str(g_).strip())
        ug=(MAT_GROUP.get(c) or (None,None))[1]
        if ug: return USER_GROUP_MAP.get(ug.lower(),ug.title())
        return BOM_GROUP_MAP.get(str(mis_g).strip().lower(),str(mis_g).strip())
    for cu,codes in UBOM_RM.items():
        L_=next((v for k,v in LINKS.items() if _norm(k)==cu),None); bn=(L_ or {}).get('bom')
        std=sb.get(str(bn or '').strip())
        if not std: continue
        sg={}
        for c,g in std.items(): sg.setdefault(grp(c,g),[]).append(c)
        name=next((k for k in LINKS if _norm(k)==cu),cu)
        for c in codes:
            if c in std or grp(c)=='Cell': continue
            if alias.get(c) in std and c in (BOM_ALIAS[1] or {}): continue    # same material, other roll length / pack (converted) - not an alternative
            cand=sg.get(grp(c)) or []
            if alias.get(c) in std: tgt=[alias[c]]
            elif len(cand)==1: tgt=cand
            elif cand: tgt=[x for x in cand if x not in codes] or cand
            else: ALT_EXTRA.setdefault(c,set()).add(name); continue
            for t in tgt: ALT_BOM.setdefault(t,{}).setdefault(c,set()).add(name)


RM_USAGE_FROM_YOUR_BOM=True   # raw-material usage per module from your BOM (MRP.xlsx) for the customers it lists, else MIS Standard BOM
BOM_ALIAS=[{},{}]  # alias / factor of this run (for your-BOM usage)
STD_RATE={}        # (planned material, MIS BOM name) -> usage per module from the MIS Standard BOM (layers added, alternate codes converted)
CS_RATE={}         # (normalised customer, CS/PM code) -> usage per module from your BOM;  ('model', Module Model, code) -> same by model
def load_rates(mis_path, mrp_path, alias, factor):
    """Usage per module straight from the BOMs (no requirement / module ratios)."""
    STD_RATE.clear(); CS_RATE.clear(); BOM_ALIAS[0]=alias; BOM_ALIAS[1]=factor
    try:
        for r in list(wb_ro(mis_path)['Standard BOM'].iter_rows(values_only=True))[1:]:
            if r and len(r)>5 and r[2] and r[3] and isinstance(r[5],(int,float)):
                c=str(r[3]).strip(); m=alias.get(c,c); f=factor.get(c,1.0) if m!=c else 1.0
                k=(m,str(r[2]).strip()); STD_RATE[k]=STD_RATE.get(k,0.0)+float(r[5])*f
    except Exception as _e_: _swallow(3383,_e_); pass
    try:
        for d in bom_rows_std([list(r) for r in wb_ro(mrp_path)['BOM'].iter_rows(values_only=True)]) or []:
            c=d['Component']
            if c[:2].upper() not in ('CS','PM') or not isinstance(d.get('Quantity'),(int,float)): continue
            k=(_norm(d.get('Customer')),c); CS_RATE[k]=CS_RATE.get(k,0.0)+float(d['Quantity'])
            mdl=str(d.get('Module Model') or '').strip()
            if mdl: CS_RATE.setdefault(('model',mdl,c),0.0); CS_RATE[('model',mdl,c)]=max(CS_RATE[('model',mdl,c)],CS_RATE[k])
    except Exception as _e_: _swallow(3391,_e_); pass

def bom_usage_src(mat, cs, pp_row):
    """(usage per module, source) - source: 'your BOM' (MRP.xlsx), 'MIS Standard BOM', or '' when not used."""
    u=bom_usage(mat,cs,pp_row)
    if cs: return u,('your BOM' if u else '')
    if UBOM_RM.get(_norm(pp_row.get('cust'))) is not None and RM_USAGE_FROM_YOUR_BOM: return u,('your BOM' if u else '')
    return u,('MIS Standard BOM' if u else '')

BLOCK_INFO={}      # material -> (MRP unit, common name) for unit conversion of BOM usage
def bom_usage(mat, cs, pp_row):
    """Usage per module of a material for one Production Plan row (customer / module type), straight from the BOM."""
    if cs:
        return CS_RATE.get((_norm(pp_row.get('cust')),mat)) or CS_RATE.get(('model',str(pp_row.get('model') or '').strip(),mat)) or 0.0
    rows=UBOM_RM.get(_norm(pp_row.get('cust')))
    if rows is not None and RM_USAGE_FROM_YOUR_BOM:            # your BOM lists this customer: use it (PPC way)
        cu_n=_norm(pp_row.get('cust'))
        alts={a for std_,alts_ in ALT_BOM.items() for a,cs_ in alts_.items()     # alternatives for THIS customer, only when its BOM also lists
              if any(_norm(x)==cu_n for x in cs_) and std_ in rows}            # the standard code (listed alone, the alternative carries the demand)
        alts|={a for a,std_ in (UBOM_ALT.get(cu_n) or {}).items() if std_ in rows}         # your BOM's '_Alt' rows
        best=0.0; bu,bname=BLOCK_INFO.get(mat,('',None)); bu=str(bu or '').strip().upper()
        for c,qs in rows.items():
            if c in alts or BOM_ALIAS[0].get(c,c)!=mat: continue
            tot=0.0
            for q,u in qs:                                          # same unit conversion as the plan (ubom_rate): e.g. RTV KG per module -> drums
                if bu and u and u!=bu and not (c!=mat and c in (BOM_ALIAS[1] or {})) and u=='KG' and PACK_KG.get(bname): q=q/PACK_KG[bname]
                tot+=q
            best=max(best,tot*(BOM_ALIAS[1].get(c,1.0) if c!=mat else 1.0))
        return best
    return STD_RATE.get((mat,str(pp_row.get('bom') or '').strip()),0.0)

def rm_plan_from_mis(DS, mis_path, modplan, lineplan=None):
    """Raw-material plan. RM USAGE ALWAYS COMES FROM MIS (your BOM is used for CS / PM only):
         plan = MIS usage per module (Total Plan Req: requirement / modules, all dates) x YOUR modules (Production Plan, via modplan).
       Only for MIS customers of a BOM Name that your Production Plan does not have is the MIS requirement (Total Plan Req) used as it is.
       Returns new blocks (copies of DS['blocks'] with plan, cust_plan, cust_line, plan_all and plan_basis set)."""
    import copy
    dates=DS['dates']; nd=len(dates); idx={pd.Timestamp(d).normalize():k for k,d in enumerate(dates)}
    wb=wb_ro(mis_path); req={}; rate_t={}; seen_=set(); reql={}; lineplan=lineplan or {}
    swp_={o_:(n_,(DS.get('alias_factor') or {}).get(o_,1.0)) for n_,o_ in CODE_SWAP.items()}   # MIS code -> your BOM code (x roll-length factor)
    for r in wb['Total Plan Req'].iter_rows(min_row=2,values_only=True):
        if not tp_ok(r,DS['plant']) or not isinstance(r[0],(dt.datetime,dt.date)) or not isinstance(r[14],(int,float)): continue
        c=str(r[10] or '').strip(); bn=str(r[4] or '').strip(); cu=' '.join(str(r[2] or '').split())
        if c in swp_:                                      # MIS requirement of a swapped tape code -> your BOM code, in its unit
            r=list(r); c,f_=swp_[c]; r[14]=r[14]*f_
        if isinstance(r[3],(int,float)) and r[3]:          # MIS usage per module = requirement / modules (all dates)
            t_=rate_t.setdefault((c,bn),[0.0,0.0]); t_[0]+=r[14]
            pl_=(c,bn,r[0],r[1],r[2],r[17] if len(r)>17 else None)
            if pl_ not in seen_: seen_.add(pl_); t_[1]+=r[3]
        k=idx.get(pd.Timestamp(r[0]).normalize())
        if k is None: continue
        req.setdefault(c,{}).setdefault(bn,{}).setdefault(cu,[0.0]*nd)[k]+=r[14]
        ln_=' '.join(str(r[1] or '').split()); reql.setdefault((c,bn,cu),{}).setdefault(ln_,[0.0]*nd)[k]+=r[14]
    alias=DS.get('alias') or {}; fac=DS.get('alias_factor') or {}
    load_rates(mis_path,SETTINGS_SRC[0],alias,fac)
    mrate={k_:(a/b_) for k_,(a,b_) in rate_t.items() if b_}
    for k_,v_ in STD_RATE.items(): mrate[k_]=v_                 # usage per module = MIS Standard BOM quantity (not requirement / modules)
    ucu={}                                       # MIS customer -> your customer (normalised) that is in your BOM
    for cu_k,L_ in LINKS.items():
        if _norm(cu_k) in UBOM_RM: ucu.setdefault(_norm(L_.get('mis') or cu_k),_norm(cu_k)); ucu.setdefault(_norm(cu_k),_norm(cu_k))
    for k_ in UBOM_RM: ucu.setdefault(k_,k_)
    out=[]
    for b0 in DS['blocks']:
        b=copy.copy(b0); c=b['mat']
        plan=[0.0]*nd; cust={}; cl={}; used_rate=[]; used_mis=[]; pls=b['plan_lines']=[]     # plan_lines: ALERT Plan lines (formulas sheet)
        def addl(cu,ln,v,f=1.0):
            a_=cl.setdefault((cu,ln or None),[0.0]*nd)
            for k in range(nd): a_[k]+=f*v[k]
        bns_=set((req.get(c) or {}).keys())|{bn_ for bn_,cs_ in modplan.items() if any(ubom_rate(ucu.get(_norm(cu_)),c,alias,fac,b0.get('uom'),b0.get('common')) for cu_ in cs_)}
        for bn in sorted(bns_):                              # MIS BOM Names of this material + those where your BOM uses it
            q=mrate.get((c,bn)); covered=set()
            if modplan.get(bn) and (q is not None or any(ubom_rate(ucu.get(_norm(cu_)),c,alias,fac,b0.get('uom'),b0.get('common')) for cu_ in modplan[bn])):
                for cu,v in modplan[bn].items():
                    qu=ubom_rate(ucu.get(_norm(cu)),c,alias,fac,b0.get('uom'),b0.get('common'))          # YOUR BOM first (None = customer not in your BOM)
                    q_=qu if qu is not None else q
                    covered.add(_norm(cu))
                    if not q_: continue
                    tag_=(f'your BOM ({cu})' if qu is not None else f'MIS {bn}')
                    if tag_ not in used_rate: used_rate.append(tag_)
                    arr=cust.setdefault(cu,[0.0]*nd)
                    for k in range(nd): plan[k]+=q_*v[k]; arr[k]+=q_*v[k]
                    lines_=(lineplan.get(bn) or {}).get(cu)
                    if lines_:
                        for ln,lv in lines_.items(): addl(cu,ln,lv,q_); pls.append(dict(cu=cu,ln=ln,bn=bn,rate=q_,mods=list(lv),src=tag_))
                    else: addl(cu,None,v,q_); pls.append(dict(cu=cu,ln=None,bn=bn,rate=q_,mods=list(v),src=tag_))
            # MIS customers of this BOM Name that your Production Plan does not have: MIS requirement as it is
            for cu,v in ((req.get(c) or {}).get(bn) or {}).items():
                if _norm(cu) in covered or not any(v): continue
                if bn not in used_mis: used_mis.append(bn)
                arr=cust.setdefault(cu,[0.0]*nd)
                for k in range(nd): plan[k]+=v[k]; arr[k]+=v[k]
                for ln,lv in (reql.get((c,bn,cu)) or {None:v}).items(): addl(cu,ln,lv); pls.append(dict(cu=cu,ln=ln,bn=bn,rate=1.0,mods=list(lv),src=f'MIS requirement {bn} (quantity, not modules)'))
        b['plan']=plan; b['cust_plan']=dict(sorted({k:v for k,v in cust.items() if any(v)}.items(),key=lambda kv:-sum(kv[1])))
        b['cust_line']=dict(sorted({k:v for k,v in cl.items() if any(v)}.items(),key=lambda kv:(-sum(cust.get(kv[0][0],[0])),kv[0][0],str(kv[0][1]))))
        b['plan_all']={pd.Timestamp(dates[k]).normalize():plan[k] for k in range(nd) if plan[k]}
        parts=[('Usage x your plan: '+', '.join(used_rate)) if used_rate else '',('MIS plan (not in your plan): '+', '.join(used_mis)) if used_mis else '']
        b['plan_basis']=' | '.join(p_ for p_ in parts if p_)
        out.append(b)
    return out

def bom_rate_checks(mis_path, bom, BW, DS_rm, TD=None):
    TD=TD or LEAD_TIME_RM
    """Usage per module: MIS standard BOM (Total Plan Req) vs your BOM sheet, for raw materials in both.
       Units are converted with the pack size written on the WSA sheet unit (e.g. 'DRUM(270KG/DRUM)')."""
    I=[]; wb=wb_ro(mis_path)
    mis_rate={}; tot_={}; unit_={}; seen_=set()
    for r in wb['Total Plan Req'].iter_rows(min_row=2,values_only=True):
        if tp_ok(r,DS_rm['plant']) and isinstance(r[12],(int,float)):
            k_=(str(r[10]).strip(),str(r[4]).strip()); q14_=r[14]
            if k_[0] in _SWP_INV() and isinstance(q14_,(int,float)): q14_=q14_*_SWP_INV()[k_[0]][1]; k_=(_SWP_INV()[k_[0]][0],k_[1])   # MIS tape code -> your BOM code
            if isinstance(r[3],(int,float)) and r[3] and isinstance(q14_,(int,float)):
                tot_.setdefault(k_,[0.0,0.0]); tot_[k_][0]+=q14_
                pl_=(k_,r[0],r[1],r[2],r[17] if len(r)>17 else None)     # modules of one plan line counted once per material
                if pl_ not in seen_: seen_.add(pl_); tot_[k_][1]+=r[3]
            unit_.setdefault(k_,str(r[13] or '').strip())
    for k_,(rq,md) in tot_.items(): mis_rate[k_]=(rq/md,unit_.get(k_,''))   # requirement / modules (each plan line counts its modules once per material line)
    pack={}
    for lab,sh in BW.get('sheets',{}).items():
        m=re.search(r'\((\d+(?:\.\d+)?)\s*KG\s*/\s*(\w+)\)',str(sh['uom'] or '').upper())
        if m:
            for nm in sh['members']: pack[nm]=float(m.group(1))
    blocks={b['mat']:b for b in DS_rm['blocks']}
    grp={}; alt_={}
    for d in bom:
        c=str(d['Component'] or '').strip(); bn=str(d['MIS BOM Name'] or '').strip(); f_=1.0; is_alt=False
        if c in (DS_rm.get('alias') or {}): f_=(DS_rm.get('alias_factor') or {}).get(c,1.0); c=DS_rm['alias'][c]; is_alt=True
        if isinstance(d['Quantity'],(int,float)):
            k=(c,bn,str(d['Module Model'] or '').strip(),str(d['UoM'] or '').strip().upper(),str(d.get('Customer') or '').strip())   # per customer: customers sharing one MIS BOM name are NOT added up
            if is_alt: alt_[k]=max(alt_.get(k,0.0),float(d['Quantity'])*f_)   # alternative code (PPC BOM '_Alt' line / other roll length): used INSTEAD of the main code, never on top
            else: grp[k]=grp.get(k,0.0)+float(d['Quantity'])*f_
    for k,q in alt_.items(): grp.setdefault(k,q)            # only when the main code is not listed for that customer / model
    # MIS rate per module for the same material: sum of its standard BOM lines = requirement / modules
    for (c,bn,mdl,u,cu_),q in grp.items():
        if c not in blocks or (c,bn) not in mis_rate: continue
        b=blocks[c]; mr,mu=mis_rate[(c,bn)]
        if u==mu.upper(): conv=q
        elif u=='KG' and b['common'] in pack: conv=q/pack[b['common']]
        else: continue
        if mr and abs(conv/mr-1)>0.05:
            ratio=conv/mr; P=[x*ratio for x in b['plan']]; T=[a+x for a,x in zip(b['transit'],b['adhoc'])]
            cv0=cover_days(own_stock(b),b['plan'],T,b.get('openpo') or [0.0]*len(P))[1]; cv1=cover_days(own_stock(b),P,T,b.get('openpo') or [0.0]*len(P))[1]
            I.append(dict(area='Plan',check='Usage per module: MIS BOM vs your BOM',sev='High' if (cv1<tgt_days(b.get('lt_used'))<=cv0 or (cv1<tgt_days(b.get('lt_used')) and cv1<cv0)) else 'Medium',
                          material=b['common'],code=c,a_lab=f'MIS per module ({mu})',a=round(mr,6),b_lab=f'Your BOM per module ({mu})',b=round(conv,6),diff=round(conv-mr,6),
                          detail=(f"{bn}: your BOM {q:.4g} {u}"+(f" ({cu_})" if cu_ else '')+(f" = {conv:.5f} {mu} at {pack.get(b['common']):.0f} KG/{mu}" if u!=mu.upper() else '')+
                                  f" vs MIS {mr:.5g} {mu} ({ratio-1:+.0%}). The MRP uses YOUR BOM for this customer (MIS only when your BOM does not list it)."),
                          action='Information: check which rate is right - if your BOM is right, ask MIS to correct its Standard BOM; if MIS is right, correct your BOM'))
    seen=set(); out=[]
    for x in I:
        k=(x['code'],x['detail'].split(':')[0],x['b'])     # one row per MIS BOM name and rate (customers with the same rate shown once)
        if k not in seen: seen.add(k); out.append(x)
    return out

BOM_SRC_NAME=[None]

def _line_key(x):
    """Common production-line key for your Production Plan and MIS: 'USA Line 3' / 'L3' -> 'L3'; Imperial*, Runergy*, Arizona* -> name."""
    t=' '.join(str(x or '').split()).lower()
    m=re.search(r'(?:usa\s*line|^l)\s*(\d+)',t)
    if m: return f'L{m.group(1)}'
    for k in ('imperial','runergy','arizona','inox'):
        if k in t: return k.capitalize()
    return ' '.join(str(x or '').split()) or '(blank)'

def mis_modules_by(mis_path, plant, dates):
    """MIS modules per day by (customer, line key), one count per plan line (date + line + customer + BOM + OA)."""
    wb=wb_ro(mis_path); idx={pd.Timestamp(d).normalize():k for k,d in enumerate(dates)}
    seen=set(); out={}
    for r in wb['Total Plan Req'].iter_rows(min_row=2,values_only=True):
        if not tp_ok(r,plant) or not isinstance(r[0],(dt.datetime,dt.date)) or not isinstance(r[3],(int,float)): continue
        k=idx.get(pd.Timestamp(r[0]).normalize())
        if k is None: continue
        key=(r[0],r[1],r[2],r[4],r[17] if len(r)>17 else None)
        if key in seen: continue
        seen.add(key); cu=' '.join(str(r[2] or '').split())
        out.setdefault((cu,_line_key(r[1])),[0.0]*len(dates))[k]+=r[3]
    return out

def make_plan_compare_xlsx(pp_rows, mis_by, dates, path):
    """Modules: your Production Plan vs MIS, by week - by customer and by production line."""

    AR='Calibri'; NAVY='1F3864'; hf=Font(name=AR,size=9,bold=True,color='FFFFFF'); hfill=PatternFill('solid',fgColor=NAVY)
    bf=Font(name=AR,size=9); bold=Font(name=AR,size=9,bold=True); amber=PatternFill('solid',fgColor='FFEB9C'); green=PatternFill('solid',fgColor='C6EFCE')
    nd=len(dates); per=[]
    for k,d in enumerate(dates):
        mon=pd.Timestamp(d)-pd.Timedelta(days=pd.Timestamp(d).weekday())
        if per and per[-1][0]==mon: per[-1][2]=k
        else: per.append([mon,k,k])
    wb=Workbook(); ws=wb.active; ws.title='Plan Compare'
    ws['A1']='Module plan compare - your Production Plan vs MIS Total Plan Req (modules per week)'; ws['A1'].font=Font(name=AR,size=14,bold=True,color=NAVY)
    ws['A2']=(f'From {pd.Timestamp(dates[0]):%d-%b-%Y}, {nd} days. MIS modules counted once per plan line (date + line + customer + BOM + OA). Your customers are matched to MIS customers '
              '(see Module Plan); lines matched as USA Line n = Ln, Imperial*, Runergy*, Arizona*. Difference = yours - MIS; amber = different. The MRP uses YOUR plan.')
    ws['A2'].font=Font(name=AR,size=9,italic=True); ws.merge_cells('A2:Z2'); ws['A2'].alignment=Alignment(wrap_text=True); ws.row_dimensions[2].height=28
    mine_c={}; mine_l={}
    for x in pp_rows or []:
        cu=x.get('mis_cust') or x.get('cust'); lk=_line_key(x.get('line'))
        for dct,key in ((mine_c,cu),(mine_l,lk)):
            a_=dct.setdefault(key,[0.0]*nd)
            for k,v in enumerate(x['v'][:nd]): a_[k]+=v
    mis_c={}; mis_l={}
    for (cu,lk),v in mis_by.items():
        for dct,key in ((mis_c,cu),(mis_l,lk)):
            a_=dct.setdefault(key,[0.0]*nd)
            for k in range(nd): a_[k]+=v[k]
    def table(top,title,keyname,mine,mis):
        ws.cell(top,1,title).font=Font(name=AR,size=12,bold=True,color=NAVY); h=top+1
        for j,t in enumerate([keyname,'Source','Match?','Total']): c=ws.cell(h,1+j,t); c.font=hf; c.fill=hfill; c.alignment=Alignment(horizontal='center')
        for w,(mon,a,z) in enumerate(per): c=ws.cell(h,5+w,mon.to_pydatetime()); c.number_format='dd-mmm'; c.font=hf; c.fill=hfill; c.alignment=Alignment(horizontal='center')
        keys=sorted(set(mine)|set(mis),key=lambda k_:-(sum(mine.get(k_,[0]))+sum(mis.get(k_,[0]))))
        keys=[k_ for k_ in keys if sum(mine.get(k_,[0]))+sum(mis.get(k_,[0]))>0]
        r=h+1; tm=[0.0]*len(per); tx=[0.0]*len(per); nd_diff=0
        for k_ in keys+['TOTAL']:
            if k_=='TOTAL': wm,wx=tm,tx
            else:
                a_=mine.get(k_,[0.0]*nd); b_=mis.get(k_,[0.0]*nd)
                wm=[sum(a_[a:z+1]) for _,a,z in per]; wx=[sum(b_[a:z+1]) for _,a,z in per]
                tm=[p+q for p,q in zip(tm,wm)]; tx=[p+q for p,q in zip(tx,wx)]
            diff=[p-q for p,q in zip(wm,wx)]; same=all(abs(d_)<0.5 for d_ in diff)
            for i_,(src,vals) in enumerate((('Your Production Plan',wm),('MIS',wx),('Difference',diff))):
                rr=r+i_; fnt=bold if (k_=='TOTAL' or src=='Difference') else bf
                ws.cell(rr,1,k_).font=fnt; ws.cell(rr,2,src).font=fnt
                c=ws.cell(rr,4,sum(vals)); c.number_format='#,##0;-#,##0;"-"'; c.font=fnt
                for w,v in enumerate(vals):
                    c=ws.cell(rr,5+w,v); c.number_format='#,##0;-#,##0;"-"'; c.font=fnt
                    if src=='Difference' and abs(v)>=0.5: c.fill=amber
                if src=='Difference':
                    c=ws.cell(rr,3,'Match' if same else 'Different'); c.font=bold; c.fill=green if same else amber
                    if not same: ws.cell(rr,4).fill=amber
            if not same and k_!='TOTAL': nd_diff+=1
            r+=3
        return r+1,nd_diff,len(keys)
    nxt,dc,nc=table(4,'By customer','Customer',mine_c,mis_c)
    _,dl,nl=table(nxt+1,'By production line','Production line',mine_l,mis_l)
    ws.cell(3,1,f'Different: {dc} of {nc} customers, {dl} of {nl} production lines.').font=Font(name=AR,size=10,bold=True,color='9C0006')
    ws.column_dimensions['A'].width=28; ws.column_dimensions['B'].width=20; ws.column_dimensions['C'].width=10; ws.column_dimensions['D'].width=11
    for w in range(len(per)): ws.column_dimensions[L(5+w)].width=9
    ws.freeze_panes='E6'; ws.sheet_view.zoomScale=90; wb.save(path)
    return dc,dl

def make_no_demand_xlsx(DS, mis_path, bom, modplan, BW, path):
    """Materials with stock (or supply coming) but NO demand in the horizon:
       A) MRP materials whose plan is 0;  B) other RM codes with stock in MB52 (plant + warehouse) or JW that are neither in the MIS plan
       (Total Plan Req, horizon) nor in your BOM for a model that is planned, nor an alternate code of a planned material."""
    from openpyxl.styles import Border, Side
    from openpyxl.worksheet.table import Table, TableStyleInfo
    AR='Calibri'; NAVY='1F3864'; thin=Side(style='thin',color='BFBFBF'); bd=Border(left=thin,right=thin,top=thin,bottom=thin)
    dates=DS['dates']; d0=pd.Timestamp(dates[0]).normalize(); dN=pd.Timestamp(dates[-1]).normalize()
    mis=wb_ro(mis_path)
    plant=DS['plant'].upper(); stock={}; desc={}; uom={}; wh={}
    for r in mis['MB52'].iter_rows(min_row=2,values_only=True):
        pn=str(r[13] or '').strip().upper()
        if pn not in (plant,plant+' WAREHOUSE'): continue
        c=str(r[2] or '').strip(); q=r[14] if isinstance(r[14],(int,float)) else 0
        stock[c]=stock.get(c,0.0)+q; desc.setdefault(c,r[5]); uom.setdefault(c,r[7])
        if pn.endswith('WAREHOUSE'): wh[c]=wh.get(c,0.0)+q
    jw={}
    for r in mis['RM Status for OEM JW'].iter_rows(min_row=2,values_only=True):
        if r and r[1] and isinstance(r[11],(int,float)) and r[11]: jw[str(r[1]).strip()]=jw.get(str(r[1]).strip(),0.0)+r[11]
    intr={}
    for r in list(mis['In Transit'].iter_rows(values_only=True))[1:]:
        if len(r)>39 and isinstance(r[34],(int,float)) and r[34]>0 and it_ok(r): intr[str(r[39] or '').strip()]=intr.get(str(r[39] or '').strip(),0.0)+r[34]
    planned=set()
    for r in mis['Total Plan Req'].iter_rows(min_row=2,values_only=True):
        if tp_ok(r,DS['plant']) and isinstance(r[0],(dt.datetime,dt.date)) and d0<=pd.Timestamp(r[0]).normalize()<=dN and isinstance(r[14],(int,float)) and r[14]>0:
            planned.add(str(r[10] or '').strip())
    live_bn={bn for bn,cu in (modplan or {}).items() if any(any(v) for v in cu.values())}
    for d in bom or []:
        if str(d.get('MIS BOM Name') or '').strip() in live_bn and d.get('Component'): planned.add(str(d['Component']).strip())
    alias=DS.get('alias') or {}
    blocks={b['mat']:b for b in DS['blocks']}
    for c,dc in alias.items():
        if dc in blocks and sum(blocks[dc]['plan'])>0: planned.add(c)
    rows=[]
    for b in DS['blocks']:
        if sum(b['plan'])>0: continue
        sup=sum(b['transit'])+sum(b.get('openpo') or [0])
        if b['stock']>0 or sup>0:
            rows.append(('A  MRP material, plan = 0','PM / CS' if b.get('count_stock') else 'RM',b['common'],b['mat'],b.get('uom'),b['stock'],0.0,sum(b['transit']),sum(b.get('openpo') or [0]),
                         'In the MRP but no customer plan uses it in the horizon'+(' (check the BOM / MIS BOM Name)' if b.get('count_stock') else '')))
    for c in sorted(set(stock)|set(jw)):
        if not c.upper().startswith('RM') or c in blocks or c in planned: continue
        q=stock.get(c,0.0); j=jw.get(c,0.0)
        if q<=0 and j<=0: continue
        if c.upper().startswith('RMFCM'):     # cells: planned by bin / efficiency, not by one code - kept apart
            rows.append(('C  Cells (not planned by this MRP)','Cell',str(desc.get(c) or ''),c,uom.get(c),q,j,intr.get(c,0.0),0.0,
                         'Cells are not in the MIS plan; your BOM names one cell code per model, so other bins / efficiencies show here - check with production'))
            continue
        why='Alternate code of a material with no plan' if c in alias else 'Not in the MIS plan and not in your BOM for a planned model'
        rows.append(('B  Other RM with stock','RM',str(desc.get(c) or ''),c,uom.get(c),q,j,intr.get(c,0.0),0.0,why))
    wb=Workbook(); ws=wb.active; ws.title='No Demand'
    ws['A1']='Materials with no demand plan - stock or supply but nothing planned to use it'; ws['A1'].font=Font(name=AR,size=14,bold=True,color=NAVY)
    ws['A2']=(f'Horizon {d0:%d-%b-%Y} to {dN:%d-%b-%Y}. Demand = MIS plan (Total Plan Req) or your BOM for a planned model (your Production Plan). '
              'Stock = MB52 (or the PPC Count when counted); JW = RM Status for OEM JW (shown only, in no calculation); CS / PM stock = PPC Count. Use it to stop new orders, re-use or return slow stock.')
    ws['A2'].font=Font(name=AR,size=9,italic=True); ws.merge_cells('A2:J2'); ws['A2'].alignment=Alignment(wrap_text=True); ws.row_dimensions[2].height=28
    H=['Group','Type','Material','Code','UOM','Stock','JW at OEM','In transit (PPC)','Open PO (WSA)','Why no demand']
    for j,h in enumerate(H,1):
        c=ws.cell(4,j,h); c.font=Font(name=AR,size=9,bold=True,color='FFFFFF'); c.fill=PatternFill('solid',fgColor=NAVY); c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True); c.border=bd
    rows.sort(key=lambda t:(t[0],-(t[5]+t[6]+t[7]+t[8])))
    for i,t in enumerate(rows):
        for j,v in enumerate(t,1):
            c=ws.cell(5+i,j,round(v,3) if isinstance(v,float) else v); c.font=Font(name=AR,size=9); c.border=bd; c.alignment=Alignment(vertical='top',wrap_text=j in (3,10))
            if j in (6,7,8,9): c.number_format='#,##0.##;-#,##0.##;"-"'
        if t[7] or t[8]:
            ws.cell(5+i,8).fill=PatternFill('solid',fgColor='FFEB9C')
            if t[8]: ws.cell(5+i,9).fill=PatternFill('solid',fgColor='FFEB9C')     # (re-assigning a cell's own fill crashes openpyxl)
    if rows: ws.add_table(Table(displayName='NoDemand',ref=f'A4:J{4+len(rows)}',tableStyleInfo=TableStyleInfo(name='TableStyleLight1',showRowStripes=True)))
    else: ws.cell(5,1,'Every material with stock has a demand plan').font=Font(name=AR,size=10,italic=True)
    ws.cell(3,1,f"{sum(1 for t in rows if t[0].startswith('A'))} MRP materials with plan 0; {sum(1 for t in rows if t[0].startswith('B'))} other RM codes with stock and no demand; "
                 f"{sum(1 for t in rows if t[0].startswith('C'))} cell codes (not planned here). "
                 "Amber = more is still coming (in transit / Open PO).").font=Font(name=AR,size=10,bold=True,color='9C0006')
    for j,w in enumerate([22,9,42,17,7,12,11,12,12,52],1): ws.column_dimensions[L(j)].width=w
    ws.freeze_panes='D5'; set_print(wb.active,'No Demand'); wb.save(path); return len(rows)

MOVE_DAYS=14   # days of expected out / in kept per run in the Movement Log
def read_movement_log(prev_path):
    """Movement Log rows kept in 1_MRP: one row per run date and material."""
    out=[]
    if not prev_path or not os.path.exists(prev_path): return out
    try: wb=wb_ro(prev_path)
    except Exception as _e_: _swallow(3705,_e_); return out
    if 'Movement Log' not in wb.sheetnames: return out
    rows=list(wb['Movement Log'].iter_rows(values_only=True))
    if not rows: return out
    h=list(rows[0])
    for r in rows[1:]:
        if not r or not r[0]: continue
        d=dict(zip(h,r))
        try:
            out.append(dict(date=pd.Timestamp(d['Date']).normalize(),code=str(d['Code']),material=d['Material'],uom=d['UOM'],stock=float(d['Stock'] or 0),
                            oem=(float(d['to OEM Stock']) if isinstance(d.get('to OEM Stock'),(int,float)) else None),
                            plan=[float(x) for x in str(d['Plan next days'] or '').split(';') if x!=''],
                            inb=[float(x) for x in str(d['In transit next days'] or '').split(';') if x!=''],
                            opo=[float(x) for x in str(d['Open PO next days'] or '').split(';') if x!=''],raw=d.get('Raw stock') or '',
                            sup=[(int(a),float(b),c) for a,b,c in (t.split('~',2) for t in str(d.get('In transit by supplier') or '').split('|') if t.count('~')==2)]))
        except Exception as _e_: _swallow(3720,_e_); continue
    return out

def movement_entries(DS):
    d0=pd.Timestamp(DS['dates'][0]).normalize(); out=[]
    for b in DS['blocks']:
        P=b['plan'][:MOVE_DAYS]; I=[a+c for a,c in zip(b['transit'],b['adhoc'])][:MOVE_DAYS]; O=(b.get('openpo') or [0.0]*MOVE_DAYS)[:MOVE_DAYS]
        ts=(DS.get('trans_sup') or {}).get(b['mat'],{})
        sup=[(k,q,sn) for sn,arr in ts.items() for k,q in enumerate(arr[:MOVE_DAYS]) if q>0]
        out.append(dict(date=d0,code=b['mat'],material=b['common'] if not b.get('count_stock') else b['desc'],uom=b.get('uom') or '',stock=float(b['stock']),
                        oem=0.0 if b.get('count_stock') else float(b.get('st_jw') or 0.0),plan=P,inb=I,opo=O,
                        cs=bool(b.get('count_stock')),raw=b.get('raw') or '',sup=sup))
    return out




NO_OEM_KEYS={'L1','L2','L3','L4'}
ATA_SHOW=4          # Arrival at plant: show the next 4 arrival dates, the rest as one total
PLAN_AVG_DAYS=globals().get('PLAN_AVG_DAYS',49)     # Action List daily plan = average of the next N days that have production (> 0) - fixed (user 21-Nov-2026), not a Settings row
def calc_layout(A, RM_):
    """Row positions on the two calculation sheets, decided before the Action List is written (its formulas point at them)."""
    tg=[]; r=2
    for x in A:
        for lk,u in sorted((x.get('tgt_use') or {}).items()):
            tg.append(dict(code=x['b']['mat'],lk=lk,use=u,row=r)); r+=1
    pl={}; r=3
    for x in A:
        if x['b']['mat'] in RM_: pl[x['b']['mat']]=r; r+=2
    r+=1; pll={}                       # per material and production line (MRP by Line)
    for x in A:
        m=x['b']['mat']
        if m not in RM_: continue
        for lk in dict.fromkeys(_line_key(ln) for (cu,ln) in RM_[m]['custs'] if ln):
            pll[(m,lk)]=r; r+=1
    PLAN_LINE_ROWS.clear(); PLAN_LINE_ROWS.update(pll)
    return tg,pl

def make_calc_target_xlsx(tg, path):
    """Calc Target: per material and line that runs a product using it today: usage per module x the line's daily target."""
    wb=Workbook(); ws=wb.active; ws.title='Calc Target'
    for j,h in enumerate(['Component','Line','Group','Usage per module (product run within the lead time)','Line daily target','Usage/Day (= D x E)'],1): ws.cell(1,j,h).font=Font(bold=True)
    for t in tg:
        r=t['row']; lab,v=EFF_TARGET.get(t['lk'],(t['lk'],0))
        ws.cell(r,1,t['code']); ws.cell(r,2,lab.split(' Production')[0]); ws.cell(r,3,'Lines 1-4' if t['lk'] in NO_OEM_KEYS else 'OEM line')
        ws.cell(r,4,round(t['use'],8)); ws.cell(r,5,f"={TGT_CELL[t['lk']]}" if t['lk'] in TGT_CELL else v); ws.cell(r,6,f'=D{r}*E{r}')
        if t['lk'] not in TGT_CELL: ws.cell(r,7,'no Settings target: highest planned day')
        ws.cell(r,8,t['lk'])
    for j,w in enumerate([17,30,11,14,12,14,34],1): ws.column_dimensions[L(j)].width=w
    wb.save(path)

def make_calc_plan_xlsx(A, RM_, pl, d0, path):
    """Calc Plan: per material 4 rows - plan with OEM (= MRP Daily Gross requirement), its running total, plan without OEM
       (= MRP Daily line rows of USA Lines 1-4), its running total. Inv. Days at daily plan = days until the running total passes the stock."""
    wb=Workbook(); ws=wb.active; ws.title='Calc Plan'
    ws.cell(1,1,f'Plan per day from MRP Daily (formulas) - Action List daily plan = average of the first {PLAN_AVG_DAYS} days (from today) with production > 0').font=Font(bold=True)
    any_=next(iter(RM_.values()),None)
    if not any_: wb.save(path); return
    C0,ND=any_['C0'],any_['ND']
    ws.cell(2,1,'Component'); ws.cell(2,2,'Row')
    for k in range(ND): ws.cell(2,3+k,(d0+pd.Timedelta(days=k)).to_pydatetime()).number_format='dd-mmm'
    for x in A:
        m=x['b']['mat']
        if m not in pl: continue
        rm=RM_[m]; r=pl[m]
        no_oem=[rr for rr,(cu,ln) in zip(rm['rc'],rm['custs']) if ln and _line_key(ln) in NO_OEM_KEYS]
        for i,lab in enumerate(['Plan with OEM','Plan without OEM (Lines 1-4)']):
            ws.cell(r+i,1,m); ws.cell(r+i,2,lab)
        for k in range(ND):
            src=L(C0+k)
            ws.cell(r,3+k,f"='MRP Daily'!{src}{rm['rp']}")
            ws.cell(r+1,3+k,('='+'+'.join(f"'MRP Daily'!{src}{rr}" for rr in no_oem)) if no_oem else 0)
    for (m,lk),r in PLAN_LINE_ROWS.items():          # one row per material and production line
        rm=RM_[m]; rows_l=[rr for rr,(cu,ln) in zip(rm['rc'],rm['custs']) if ln and _line_key(ln)==lk]
        ws.cell(r,1,m); ws.cell(r,2,f'Plan line {lk}')
        for k in range(ND): ws.cell(r,3+k,('='+'+'.join(f"'MRP Daily'!{L(C0+k)}{rr}" for rr in rows_l)) if rows_l else 0)
    ws.column_dimensions['A'].width=17; ws.column_dimensions['B'].width=28
    wb.save(path)



def al_group(h):
    if 'module plan' in h: return 'mp'
    if 'OEM' in h and ('arget' in h or 'lan' in h):
        return ('tg_' if 'arget' in h else 'pl_')+('wo' if 'without OEM' in h else 'w')
    if 'WSA RM Planning' in h: return 'wsa'
    if h in ('Status','Exception message','Action','Owner','Due','Status update','Comment'): return 'work'
    return 'base'
def al_hdr_light(h): return ('day by day' in h) or ('stock only' in h)
def al_hdr_color(h):
    """Action List header colours: PURPLE = daily target, GOLD = daily plan (light = stock only / day by day);
       the scenario (without / with OEM) is the coloured banner in row 3 and the end of the header text."""
    g=al_group(h)
    if g in ('tg_w','tg_wo','pl_w','pl_wo'):
        tgt=g.startswith('tg_')
        if al_hdr_light(h): return 'D9C3E9' if tgt else 'FFE699'
        return '7030A0' if tgt else 'BF8F00'
    return {'wsa':'833C0B','work':'595959','mp':'BF8F00'}.get(g,'1F3864')



def sup_short(n):
    """Short supplier name: no country in brackets, no Co./Ltd./Glass/Technology..., first 2 words (FLAT (VIETNAM) CO., LTD -> Flat)."""
    t=re.sub(r'\([^)]*\)',' ',str(n or '')).replace(',',' ').replace('.',' ')
    w=[x for x in t.split() if x.upper() not in ('CO','LTD','LIMITED','COMPANY','CORP','CORPORATION','INC','GLASS','TECHNOLOGY','TECH','INDUSTRIAL','INDUSTRY','NEW','MATERIAL','MATERIALS','SOLAR','ENERGY','JSC','PTE','LLC','GROUP')]
    return ' '.join(w[:2]).title() or str(n or '')




SUMMARY_LINES=[]; DEF_SHEETS_SEEN=[set()]
def audit_workbook(path):
    DEF_SHEETS_SEEN[0]={str(r[0]).strip().lower() for r in DEFINITIONS}|{p_.strip().lower() for r in DEFINITIONS for p_ in re.split(r' / |, | \(',str(r[0]))}
    """Self-check after every run (your rules): (1) a material short in the plan is never labelled safe on a visible sheet;
       (2) every visible sheet lists materials in the same shortage-risk order as the ALERT sheet (Action List is checked like the others).
       Prints what breaks the rules AND which visible sheets were NOT checked and why."""
    try:
        wb=load_workbook(path,read_only=True)
        if 'ALERT' not in wb.sheetnames: print('NOTE: self-check skipped (no ALERT sheet)'); return
        order=[]; short=set(); STOCKOUT={ALERT_CLS[0][0],ALERT_CLS[1][0]}        # reference list = the ALERT rows (class, then first stock-zero day)
        for r in wb['ALERT'].iter_rows(min_row=ALERT_HR+1,values_only=True):
            cls_,code_=r[1],r[2]
            if not code_ or not isinstance(code_,str) or not re.match(r'^(JW)?(RM|PM)[A-Z]*\d+$',code_.strip()): continue
            if code_ not in order: order.append(code_)
            if cls_ in STOCKOUT: short.add(code_)                                   # short in the plan (inside or after the transit lead time)
        if not order: print('NOTE: self-check skipped (ALERT sheet has no rows)'); return
        dup_={c for c in order if sum(1 for r in wb['ALERT'].iter_rows(min_row=ALERT_HR+1,min_col=3,max_col=3,values_only=True) if r[0]==c)>1}   # split in 2 ALERT rows (back glass 333 per WSA sheet)
        dup_|={c_ for v_ in SUBG.values() for c_,*_r in v_}           # Settings SUBSTITUTES: the substitute row sits under its material (sub_cover), not in risk order
        rank={c:i for i,c in enumerate([c for c in order if c not in dup_])}; safe={'In stock','In Stock','Excess','Above reorder','OK','no shortage','No action','covered'}
        issues=[]; NOT_CHECKED=[]; CHECKED=[]
        SKIP={'Settings','Stock Count','BOM','Module Plan','MIS Layout','How it works','Definitions','Movement Log','In Transit Status','RM Issues',
              'Check','Stock Data','No Demand','Substitute Advice','ALERT (formulas)','ALERT Calc','ALERT Plan lines','ALERT Arrival lines','ETD alert'}   # these were hidden (never checked) before all sheets became visible; not in shortage order
        FREE={'ALERT',HEAT_NAME}     # Cover Heatmap: grouped RM / PM / CS (user, Nov 2026), risk order inside each group; ALERT: own order (alert class, first stock-zero day; back glass 333 per customer project): only the 'labelled safe' rule is checked
        for ws in wb.worksheets:
            if ws.sheet_state!='visible': continue
            seq=[]; bad=0; SEVL=[None]
            if ws.title in SKIP:                                         # your input copies / logs kept in source or date order
                NOT_CHECKED.append(f'{ws.title} (source / date order or log)'); continue
            for i,r in enumerate(ws.iter_rows(values_only=True),1):
                if not any(v not in (None,'') for v in r):              # blank row: next list starts
                    bad+=sum(1 for a_,b_ in zip(seq,seq[1:]) if rank[b_]<rank[a_]); seq=[]; continue
                sev_=r[0] if r and r[0] in ('High','Medium','Low') else None
                if sev_ and sev_!=SEVL[0]:                                 # Check sheet: each severity group is its own list
                    bad+=sum(1 for a_,b_ in zip(seq,seq[1:]) if rank[b_]<rank[a_]); seq=[]; SEVL[0]=sev_
                codes=[v for v in r[:5] if isinstance(v,str) and v in rank]
                if not codes and any(isinstance(v,str) and v.strip() and not CODE.match(v.strip()) for v in r[:2]) and not any(isinstance(v,(int,float)) for v in r[:6]):
                    bad+=sum(1 for a_,b_ in zip(seq,seq[1:]) if rank[b_]<rank[a_]); seq=[]; continue      # a title / header row: next table starts
                if not codes: continue
                if codes[0] in short and any(isinstance(v,str) and v.strip() in safe for v in r): issues.append(f'{ws.title} row {i}: {codes[0]} short but labelled safe')
                if codes[0] not in seq: seq.append(codes[0])                 # a material split in 2 rows (back glass 333 per WSA sheet) counts once
            bad+=sum(1 for a,b in zip(seq,seq[1:]) if rank[b]<rank[a])
            (CHECKED if ws.title not in FREE else NOT_CHECKED).append(ws.title if ws.title not in FREE else f'{ws.title} (order not checked - only the labelled-safe rule)')
            if bad and ws.title not in FREE: issues.append(f'{ws.title}: {bad} places out of shortage-risk order')
        print(('Self-check (shortage-risk rules): no problems found' if not issues else 'Self-check WARNING:\n  '+'\n  '.join(issues[:10]))
              +f"\n  Order checked on {len(CHECKED)} sheet(s): {', '.join(CHECKED)}\n  NOT checked ({len(NOT_CHECKED)}): {'; '.join(NOT_CHECKED)}")
        SUMMARY_LINES.append(f'Self-check order test NOT applied to {len(NOT_CHECKED)} sheet(s) (listed above in the Self-check lines)')
        miss_=[w_.title for w_ in wb.worksheets if w_.title!='Definitions' and w_.title.lower() not in DEF_SHEETS_SEEN[0]]
        if miss_: print('WARNING: sheets with NO row on the Definitions sheet: '+', '.join(miss_)); SUMMARY_LINES.append('Sheets without a Definitions row: '+', '.join(miss_))
        else: print('Definitions check: every sheet of this workbook has a row on the Definitions sheet')
    except Exception as ex_: print(f'NOTE: self-check skipped ({type(ex_).__name__})')


def run_summary(title, extra=()):
    """What this run did NOT count / skipped / swallowed - printed and appended to 7_Summary_<date>.txt next to the outputs."""
    L_=[f'--- RUN SUMMARY: {title} ---']
    L_+=list(extra)
    L_+=[m_ for m_ in RUN_NOTES if m_.startswith(('WARNING','ERROR'))]
    L_+=[m_ for m_ in RUN_NOTES if m_.startswith('NOTE')]
    L_+=list(SUMMARY_LINES)
    if SWALLOWED:
        L_.append(f'{len(SWALLOWED)} silent error handler(s) caught an error this run (script line: count, error) - mostly file probing, but check any that look like data:')
        L_+=[f'  line {k}: x{v[0]} {v[1]}: {v[2]}' for k,v in sorted(SWALLOWED.items())]
    else: L_.append('No silent error handler caught an error this run.')
    for l_ in L_: _bi.print(l_)
    try:
        fp_=os.path.join(OUTDIR,outfile_name('sum','txt'))
        with open(fp_,'a',encoding='utf-8') as h_: h_.write('\n'.join(L_)+'\n\n')
    except Exception as e_: _bi.print(f'NOTE: 7_Summary file not written ({e_})')

POST={'path':None,'wb':None}
def post_wb(path):
    """The finished workbook is opened ONCE for all finishing steps (dashboards, sheets, style) and saved once (post_flush)."""
    if POST['path']!=path or POST['wb'] is None: POST['wb']=load_workbook(path); POST['path']=path
    return POST['wb']
def post_flush(path):
    if POST['wb'] is not None and POST['path']==path: _safe_save(POST['wb'],path)
    POST['wb']=None; POST['path']=None

def _fix_views(wb):
    """Excel 'Repaired Records: View from /xl/worksheets/sheetN.xml' (user, Nov 2026): a sheet whose frozen pane was removed
       (_tidy_dashboard, no freeze) kept its <selection pane="..."> entries - pointing at panes that no longer exist, and doubled.
       Every sheet: no pane -> one plain selection; a pane -> one selection per pane that exists for that split (last one kept)."""
    from openpyxl.worksheet.views import Selection
    for ws_ in wb.worksheets:
        try:
            sv_=ws_.sheet_view; sel_=list(sv_.selection or [])
            act_=next((x.activeCell for x in reversed(sel_) if x.activeCell),None) or 'A1'
            pn_=sv_.pane
            if pn_ is None or not (pn_.xSplit or pn_.ySplit):
                sv_.pane=None; sv_.selection=[Selection(activeCell=act_,sqref=act_)]; continue
            ok_={'bottomRight','topRight','bottomLeft'} if (pn_.xSplit and pn_.ySplit) else ({'topRight'} if pn_.xSplit else {'bottomLeft'})
            keep_={}
            for x in sel_:
                if x.pane in ok_: keep_[x.pane]=x
            if not keep_: keep_={pn_.activePane or sorted(ok_)[0]:Selection(pane=pn_.activePane or sorted(ok_)[0],activeCell=act_,sqref=act_)}
            sv_.selection=[keep_[k] for k in ('topRight','bottomLeft','bottomRight') if k in keep_]
        except Exception as _e_: _swallow(0,_e_)
def _safe_save(wb, path):
    """Save to a temporary file first and replace the real file only when the save worked - a failing step never damages 1_MRP."""
    _fix_views(wb)
    print_a4_all(wb,'Daily Material Status' if '4_Status' in os.path.basename(path) else 'MRP Daily')
    tmp_=path+'.saving.xlsx'
    try:
        wb.save(tmp_); os.replace(tmp_,path)
    finally:
        if os.path.exists(tmp_):
            try: os.remove(tmp_)
            except OSError: pass



def strip_code(t):
    """Alternative text without material codes: 'RMJUB00000072 JB 500/30A: 115,961 (Pepper)' -> 'JB 500/30A: 115,961 (Pepper)'."""
    return re.sub(r'\b(RM|PM|CS)[A-Z]*\d{5,}\s*','',str(t or '')).strip() or None
def is_rm(code): return str(code or '').upper().startswith('RM')


def risk_info(x, d0):
    """Shortage risk of one material: first short day (projected stock incl. in transit < 0), lead time, class, short date."""
    w_=x.get('wins') or []; k=w_[0][0] if w_ else None; lt=int(x['b'].get('lt_used') or 0); td=x.get('pl_days')     # days at the module plan
    if x.get('no_demand'): return dict(k=None,lt=lt,cls='No demand',date=None)      # no demand = never Critical / At risk / Watch / Short
    cls=('ALERT' if (k is not None and k<lt) else ('Short later' if k is not None else ('Under lead time' if (td is not None and td<tgt_days(lt)) else ('OK' if x.get('pl_need') else 'No need'))))
    return dict(k=k,lt=lt,cls=cls,date=(d0+pd.Timedelta(days=k)) if k is not None else None)
SUBG={}    # Settings SUBSTITUTES between two MRP materials: main code -> [(substitute code, factor, use from, use until)] - NOT pooled (user 25-Nov-2026)
CUST_FIX_=[]     # customer spellings matched to the BOM
NODEM_=[]            # materials with no demand in this run (mat, name, why)
RISK_RANK={}       # material -> rank (0 = highest shortage risk), same order as the Action List, used by every sheet
def set_risk_rank(A, d0):
    """Rank all materials by shortage risk: ALERT (earliest first) > short later (earliest first) > under the lead time (lowest days) > the rest."""
    G={'ALERT':0,'Short later':1,'Under lead time':2,'Under lead time later':2.5,'OK':3,'No need':4,'No demand':5}   # same six groups as the ALERT sheet
    def key(x):
        ri=risk_info(x,d0); td=x.get('pl_days'); cl=ri['cls']; k_=ri['k']
        if cl=='OK':                                                  # ALERT class 'Under lead time': STORE stock alone covers < the lead time (in transit may cover it)
            st_=float(x['b'].get('stock') or 0.0); nd_=x.get('pl_need')
            if nd_ and max(0.0,st_)/nd_<tgt_days(x['b'].get('lt_used')): cl='Under lead time'
        if cl=='OK':                                                  # ALERT class 'Under lead time later': cover falls below the lead time within 16 weeks
            S_=x.get('S') or []; nd_=x.get('pl_need')
            t_=next((i for i in range(min(len(S_),ALERT_HORIZON)) if nd_ and S_[i]/nd_<tgt_days(x['b'].get('lt_used'))),None)
            if t_ is not None: cl='Under lead time later'; k_=t_
        return (G[cl],k_ if k_ is not None else 0,td if td is not None else 1e9,str(x['b'].get('common') or ''))
    RISK_RANK.clear(); RISK_RANK.update({x['b']['mat']:i for i,x in enumerate(sorted(A,key=key))})
def rrank(code): return RISK_RANK.get(code,10**6)
def short_lt_text(x, d0):
    """All periods with stock below zero that START inside the material transit lead time ('05-Nov → 07-Nov (3 d); 20-Nov → beyond 21-Mar');
       else the first stock-zero day after the transit lead time; else 'not short'. Returns (text, number of periods inside the transit lead time)."""
    if x.get('no_demand'): return 'No demand',0
    w_=x.get('wins') or []; lt_=int(x['b'].get('lt_used') or 0)
    if not w_: return 'not short',0
    D=lambda k: d0+pd.Timedelta(days=k)
    ins=[w for w in w_ if w[0]<lt_]
    if ins:
        return ';\n'.join(f"{D(a):%d-%b} → "+(f"beyond {D(z):%d-%b}" if op else f"{D(z):%d-%b} ({z-a+1} d)") for a,z,op in ins),len(ins)
    return f"{D(w_[0][0]):%d-%b} (after transit lead time {lt_} d)",0

