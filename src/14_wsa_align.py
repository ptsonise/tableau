# ---------- front / back glass: one sheet per glass (user rule, Nov 2026) ----------
WSA_DROP_SHEETS=('Total_Plan_Req',)   # left out of the aligned file. 'Stock' is KEPT: it is plain values, and 'Summary Dashboard' (column B) and the
                                      # Frame Topcon 30HS / 35HS sheets read it - dropping it turned those formulas into =SUMIFS(#REF!,#REF!,...)
_GQ_PLAN=('bal','dem','act','ppcit'); _GQ_DLV=('sch','dlv')
_GT=('po','ship','rem','book','ready','etd','bl','vessel','inv','ref','rem2')
def _gcell(rows,r,c): return rows[r-1][c-1] if (r and c and 0<r<=len(rows) and 0<c<=len(rows[r-1])) else None
def _gset(rows,r,c,v):
    while len(rows)<r: rows.append([])
    if len(rows[r-1])<c: rows[r-1].extend([None]*(c-len(rows[r-1])))
    rows[r-1][c-1]=v
def _glass_side(m, AD):
    """'F' (front glass: Front / HTAR / ARC) or 'B' (back glass: Back / Printed) by the MIS Common Name, else None."""
    t=str(((AD.get('blocks') or {}).get(m) or {}).get('common') or '')
    if _kind_of(t)!='GLASS': return None
    s=_desc_keys(t)['side']
    return next(iter(s)) if len(s)==1 else None
def _it_bl_index():
    """PPC In_Transit lines (all of them, material traced like 1_MRP) by BL number: (line, material, Final Qty.)."""
    ix={}; tail={}
    try: rows=list(wb_ro(MIS_PATH_[0])['In Transit'].iter_rows(values_only=True)) if MIS_PATH_[0] else []
    except Exception: rows=[]
    if len(rows)<2: return ix,tail
    I_=IT_CANON.index; iF,iQ,iQ0,iHB,iMB=I_('Final RM Code'),I_('Final Qty.'),I_('Qty'),I_('HBLNo.'),I_('MBL')
    for k,r in enumerate(rows[1:]):
        q=_num_cell(r[iQ]) if iQ<len(r) else None
        if q is None and iQ0<len(r): q=_num_cell(r[iQ0])
        c=str(r[iF] or '').strip().upper()
        if not c or not q: continue
        for b in _bl_set(r[iHB] if iHB<len(r) else None)|_bl_set(r[iMB] if iMB<len(r) else None):
            ix.setdefault(b,[]).append((k,c,q))
            if len(b)>=8: tail.setdefault(b[-8:],[]).append((k,c,q,b))
    return ix,tail
def _mis_bl_dates(mis_path):
    """PPC In_Transit: BL number (HBLNo. / MBL) -> the plant arrival dates (Final AtA) of its pending lines.
       User rule, Nov 2026: 'MIS is always the most updated' - a WSA delivery row with another date is to be moved to the MIS date."""
    ix={}; tail={}
    try: rows=list(wb_ro(mis_path)['In Transit'].iter_rows(values_only=True)) if mis_path else []
    except Exception: rows=[]
    if len(rows)<2: return ix,tail
    I_=IT_CANON.index; iA,iHB,iMB=I_('Final AtA'),I_('HBLNo.'),I_('MBL')
    for r in rows[1:]:
        d=_date_cell(r[iA]) if iA<len(r) else None
        if d is None: continue
        d=pd.Timestamp(d).normalize()
        for b in _bl_set(r[iHB] if iHB<len(r) else None)|_bl_set(r[iMB] if iMB<len(r) else None):
            ix.setdefault(b,[]).append(d)
            if len(b)>=8: tail.setdefault(b[-8:],[]).append((b,d))
    return ix,tail
def _mis_date_of(bl_cell, idx):
    """The PPC arrival date of a WSA row: the earliest Final AtA of the MIS lines with one of its BL numbers (suffix match like the
       PPC In_Transit match: MBL 'CMDUCHN4068035' = 'CHN4068035'); None when MIS has none of them."""
    ix,tail=idx; ds=[]
    for b in _bl_set(bl_cell):
        ds+=ix.get(b,[])
        if len(b)>=8: ds+=[d for b2,d in tail.get(b[-8:],[]) if b2!=b and (b2.endswith(b) or b.endswith(b2))]
    return min(ds) if ds else None
def _glass_shares(bl_cell, sides, ix, tail, A):
    """Front / back share of one planner row from the In_Transit lines of its BL numbers (None: no line of these glasses)."""
    per={}; seen=set()
    for b in _bl_set(bl_cell):
        hits=list(ix.get(b,[]))
        if len(b)>=8: hits+=[(k,c,q) for k,c,q,b2 in tail.get(b[-8:],[]) if b2!=b and (b2.endswith(b) or b.endswith(b2))]
        for k,c,q in hits:
            s=sides.get(A(c))
            if s and k not in seen: seen.add(k); per[s]=per.get(s,0.0)+q
    tot=sum(per.values())
    return {s:per.get(s,0.0)/tot for s in ('F','B')} if tot>0 else None
def _glass_part(B, side, codes, name, shares, note, block_no, dash, oname):
    """One glass of a planning block that counts front + back: every quantity of the planner's rows split - deliveries
       (Scheduled, Delivered) by the In_Transit lines of the row's BL# (shares), else 1:1; Balance / Demand / Actual / PPC
       arrivals of the kept rows 1:1 (one front + one back glass per module). Delivery text stays only on the rows whose
       shipment has this glass."""
    rv=[list(x) for x in B['rv']]; rf=[list(x) for x in B['rf']]; ro=B['role']
    for d,r in B['dates'].items():
        sd=(shares.get(r) or {}).get(side,0.5)
        has_q=any((_num(_gcell(B['rv'],r,ro[k])) or 0)!=0 for k in _GQ_DLV if k in ro)
        for k in _GQ_PLAN+_GQ_DLV:
            if k not in ro: continue
            v=_num(_gcell(B['rv'],r,ro[k]))
            if v is None: continue
            nv=round(v*(sd if k in _GQ_DLV else 0.5),6)
            _gset(rv,r,ro[k],None if (k in _GQ_DLV and v and not nv) else nv); _gset(rf,r,ro[k],None)     # none of this glass: blank
        if has_q and sd<=1e-9:
            for k in _GT:
                if k in ro: _gset(rv,r,ro[k],None); _gset(rf,r,ro[k],None)
    return dict(B,rv=rv,rf=rf,codes=list(codes),label=name,oname=oname,block=block_no,note=note,dash=dash,src_label=B['label'])
def _glass_split(blocks, AD):
    """Front and back glass on their own sheets (user rule, Nov 2026: 'I want front glass and back glass separated in different
       sheets' - one Balance per glass; the BOM uses both codes). A planning block with front + back glass becomes
       '<sheet> Front' + '<sheet> Back'. A block that lists one glass but counts its partner too (BL# lines of the other glass:
       Glass G12R Topcon Subanci = 356 front + 333 back) is split the same way: its own glass keeps the sheet ('<sheet> Front'),
       the partner gets '<sheet> Back' - the customer project keeps its own sheets (user rule, Nov 2026: 'follow the MRP daily
       settings ... which supplier goes with which customer'; a glass on 2+ sheets is split by customer and supplier in
       _wsa_shared). Delivery rows are split by the In_Transit lines of their BL#, else 1 front : 1 back. The new sheets carry
       'glass only' in A1, so the next run (on the aligned file) does not split them again. Returns (blocks, info); info['arch']
       splits the Archive records of the split blocks the same way."""
    al=AD.get('alias') or {}; A=lambda c: al.get(c,c)
    ix,tail=_it_bl_index()
    res=[]; info=dict(renames={},lines=[],src={},n_bl=0,n_11=0)
    word={'F':'front','B':'back'}
    def shares_of(B,sides):
        sh={}
        if 'bl' not in B['role']: return sh
        for d,r in B['dates'].items():
            has_q=any((_num(_gcell(B['rv'],r,B['role'][k])) or 0)!=0 for k in _GQ_DLV if k in B['role'])
            s_=_glass_shares(_gcell(B['rv'],r,B['role']['bl']),sides,ix,tail,A)
            if s_ is not None: sh[r]=s_
            if has_q: info['n_bl' if s_ is not None else 'n_11']+=1
        return sh
    def split(o,B,mats,sides,extra=''):
        sh_=B['sheet']; nm={'F':f'{sh_[:25]} Front','B':f'{sh_[:26]} Back'}
        shares=shares_of(B,sides)
        own=sides[A(mats[0])] if extra else 'F'; oth='B' if own=='F' else 'F'
        P=_glass_part(B,own,[c for c in mats if sides[A(c)]==own],nm[own],shares,f"({word[own]} glass only - {word[oth]} glass: sheet '{nm[oth]}')",0,True,nm[own])
        Q=_glass_part(B,oth,[c for c in sides if sides[c]==oth and (not extra or c not in mats)] if extra else [c for c in mats if sides[A(c)]==oth],
                      nm[oth],shares,f"({word[oth]} glass only - {word[own]} glass: sheet '{nm[own]}')",100,False,nm[oth])
        res.extend([(o,0,P),(o,1,Q)]); info['renames'][sh_]=nm[own]; info['src'][B['label']]=dict(F=nm['F'],B=nm['B'],sides=sides)
        info['lines'].append(f"{sh_} -> {nm[own]} ({', '.join(P['codes'])}) + {nm[oth]} ({', '.join(Q['codes'])}{extra})")
    for o,B in enumerate(blocks):
        mats=[c for c in B['codes'] if c in AD['blocks']]
        sides={A(c):_glass_side(c,AD) for c in mats}
        mk_=next((v for row in B['rv'][:max(B['hr']-1,0)] for v in row if isinstance(v,str) and 'glass only' in v.lower()),None)
        marked=mk_ is not None
        if marked:                                 # split by an earlier run: keep its note (and so the marker) in A1
            i_=mk_.lower().find('glass only'); j_=mk_.rfind('(',0,i_); B=dict(B,note=mk_[j_ if j_>=0 else i_:].strip())
            if re.search(r"its (front|back) glass RM|includes the (front|back) glass of sheet",mk_):
                info['lines'].append(f"WARNING: '{B['sheet']}' comes from an aligned file of an older version (a customer's glass added to "
                                     "another customer's sheet) - run the script on your own WSA file")
        if mats and all(sides.values()) and set(sides.values())=={'F','B'}:
            split(o,B,mats,sides)
        elif mats and not marked and all(sides.values()) and len(set(sides.values()))==1 and 'bl' in B['role']:
            own=next(iter(sides.values())); oth='B' if own=='F' else 'F'; cnt=Counter()
            for d,r in B['dates'].items():
                for b in _bl_set(_gcell(B['rv'],r,B['role']['bl'])):
                    for k,c,q in ix.get(b,[]):
                        if _glass_side(c,AD)==oth: cnt[A(c)]+=q
            if cnt:                                # its BL# lines carry the partner glass: this sheet counts front + back
                partner=cnt.most_common(1)[0][0]; s2=dict(sides); s2[partner]=oth
                split(o,B,mats,s2,extra=': its BL# lines')
            else: res.append((o,0,B))
        else: res.append((o,0,B))
    res.sort(key=lambda t:(t[0],t[1]))
    def arch(old):
        """Archive records of a block split in this run -> one record per glass (BL# lines, else 1:1)."""
        out=[]
        for r in old:
            sp=info['src'].get(str(r[0]))
            if not sp: out.append(r); continue
            s_=_glass_shares(r[10],sp['sides'],ix,tail,A) or {'F':0.5,'B':0.5}
            q_d,q_s=_num(r[2]),_num(r[3])
            for side in ('F','B'):
                f=s_.get(side,0.5)
                if (q_d or q_s) and f<=1e-9: continue
                n=list(r); n[0]=sp[side]
                if q_d: n[2]=round(q_d*f,6)
                if q_s: n[3]=round(q_s*f,6)
                out.append(n)
        return out
    info['arch']=arch
    return [B for _,_,B in res],info
def _ref_fix(text, renames, gone):
    """Formulas / defined names: references to a renamed sheet follow the new name; references to a removed sheet become
       #REF! (what Excel does when a sheet is deleted)."""
    for o,n in renames.items():
        text=text.replace(f"'{o}'!",f"'{n}'!").replace(f"&apos;{o}&apos;!",f"&apos;{n}&apos;!")
    cell=r"(?:\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?|\$?[A-Z]{1,3}:\$?[A-Z]{1,3}|\$?\d+:\$?\d+)"
    for g in gone:
        q=re.escape(g); alt=[f"'{q}'",f"&apos;{q}&apos;"]
        if re.fullmatch(r'[A-Za-z_][A-Za-z0-9_.]*',g): alt.append(r"(?<![A-Za-z0-9_.'\s;&])"+q)
        text=re.sub('(?:'+'|'.join(alt)+')!'+cell,'#REF!',text)
    return text

def _out_sheet_name(B):
    """Name of a planning block's sheet in the aligned WSA file."""
    if B.get('oname'): return B['oname'][:31]
    if B['block']==0: return B['sheet']
    o=WSA_SPLIT_NAMES.get(B['sheet']) if B['block']==1 else None
    return (o or f"{B['sheet'][:27]} {B['block']+1}")[:31]
def _vendor_rules(AD):
    """Settings ALLOWED VENDORS: {(customer, MRP material): [vendors, top row first]} - the customer under its Settings name and
       its MIS name (LINKS). Like 1_MRP, a customer may use all its vendors, top row used first (user 23-Nov-2026); no row = any vendor."""
    al=AD.get('alias') or {}; out={}
    for r_ in SET_CUSTSUP or []:
        cu=r_.get('Customer'); c=str(r_.get('Code') or '').strip(); v=' '.join(str(r_.get('Supplier') or '').split()).upper()
        if not (cu and c and v): continue
        L_=next((x for k,x in LINKS.items() if _norm(k)==_norm(cu)),None) or {}
        for n in {_norm(cu),_norm(L_.get('mis') or '')}-{''}:
            lst=out.setdefault((n,al.get(c,c)),[])
            if v not in lst: lst.append(v)
    return out
def _vend_hit(a, b):
    """Same vendor? Settings 'FLAT' = batch Make 'FLAT' = 'FLAT (HONG KONG) CO., LIMITED'; KB = KIBING."""
    a=str(a or '').upper().strip(); b=str(b or '').upper().strip()
    a=SUPPLIER_ALIAS.get(a,a); b=SUPPLIER_ALIAS.get(b,b)
    return bool(a and b) and 'UNASSIGNED' not in (a,b) and (a in b or b in a)
VP_UNKNOWN=('','?','UNASSIGNED','OTHER / UNASSIGNED'); VP_WSA={}; VP_MRP={}
def vendor_pools(dem, stock, arr, nd, vend):
    """Settings ALLOWED VENDORS on one projection - ONE rule for the WSA daily (planning sheet) and 1_MRP (user rule, Nov 2026:
       "WSA output must follow settings, e.g. allowed vendors by customer ... all output workbooks must follow settings").
       dem   = {customer name or None: [demand per day]} (OEM left out; None = not split by customer = any vendor)
       stock = {supplier name: qty today}; arr = {supplier name: [arrivals per day]}; vend(customer) = its allowed vendors, preferred first (list; a plain name = one vendor) or None.
       Pools: one per vendor a customer with demand is tied to (that vendor's stock + arrivals), FREE = stock / arrivals of unknown
       vendor (Settings: 'always allowed'), ANY = other vendors (only customers without a vendor rule may use them).
       A tied customer may use its vendors (every row of Settings, top row used first: 'Dow run out then use HU', user 23-Nov-2026) + FREE; a customer without a rule may use everything. Stock-zero day = the first day the
       customers tied to a vendor need more (up to that day) than that vendor + FREE supplied (up to that day), or the pooled
       projection is below zero - the 'up to that day' rule of every projection here (no fixed order of use).
       Returns None when no customer with demand has a rule, else dict(lock=[per day], eff=[per day], pools={label: [own balance:
       supply - demand of that vendor's customers, end of day]}, who={label: [customers]}, first=first day the rule makes short or
       None). eff = projection that follows the rule (< 0 = short by that much); lock = pooled projection - eff = stock + arrivals
       no customer short that day may use (0 = the rule changes nothing that day). Sum of the pools = the pooled projection."""
    def vl(c):                                                        # allowed vendors of a customer, preferred first ([] = no rule)
        v=vend(c) if c is not None else None
        return [] if not v else ([v] if isinstance(v,str) else [x for x in v if x])
    tied=sorted({V for c,v in dem.items() if c is not None and any(x>1e-9 for x in v) for V in vl(c)})
    if not tied: return None
    def pool_of(name):
        n=' '.join(str(name or '').split()).upper()
        if n in VP_UNKNOWN: return 'FREE'
        return next((V for V in tied if _vend_hit(V,n)),'ANY')
    sup={p:[0.0]*nd for p in tied+['FREE','ANY']}; need={'ANY':[0.0]*nd}
    for n,q in stock.items(): sup[pool_of(n)][0]+=float(q or 0.0)
    for n,v in arr.items():
        p=pool_of(n)
        for k in range(min(nd,len(v))): sup[p][k]+=float(v[k] or 0.0)
    who={p:[] for p in tied+['ANY']}; gneed={}                        # gneed: customers with the same allowed vendors (same order) = one group
    for c,v in dem.items():
        vs=vl(c)
        if not vs:
            if c is not None and c not in who['ANY']: who['ANY'].append(c)
            for k in range(min(nd,len(v))): need['ANY'][k]+=float(v[k] or 0.0)
            continue
        for V in vs:
            if c not in who[V]: who[V].append(c)
        g=gneed.setdefault(tuple(vs),[0.0]*nd)
        for k in range(min(nd,len(v))): g[k]+=float(v[k] or 0.0)
    def cum(x):
        o=[]; t_=0.0
        for v in x: t_+=v; o.append(t_)
        return o
    G={g:cum(v) for g,v in gneed.items()}; Sc={V:cum(sup[V]) for V in tied}
    Ad=[a_-b_ for a_,b_ in zip(cum(need['ANY']),cum(sup['ANY']))]; Fc=cum(sup['FREE'])
    # Short on day K when some customers tied to vendors need more (up to K) than THEIR vendors + unknown-vendor supply (up to K) - Hall's
    # condition over every set of customer groups (a customer may use several vendors, user 23-Nov-2026 'both allowed, Dow run out then
    # use HU') - or when everything together is short (= the pooled projection). Customers without a rule never add a condition.
    gl=list(G); subs=[]
    if len(gl)<=10:
        for mask in range(1,1<<len(gl)):
            sg=[gl[i] for i in range(len(gl)) if mask>>i&1]; subs.append((sg,sorted({V for g in sg for V in g})))
    else: subs=[([g],list(g)) for g in gl]+[(gl,tied)]
    # Order of use (pools = display): groups with the fewest vendors first, each takes its first vendor until it is used up, then the next.
    order=sorted(gl,key=lambda g:(len(g),g))
    def alloc(K):
        rem={V:Sc[V][K] for V in tied}
        for g in order:
            left=G[g][K]
            for V in g:
                t_=min(max(rem[V],0.0),left); rem[V]-=t_; left-=t_
            if left>1e-9: rem[g[0]]-=left
        return rem
    dead=sum(need['ANY'])<=1e-6
    Dc=cum(sup['ANY']) if dead else [0.0]*nd
    dst=sum(float(q or 0.0) for n,q in stock.items() if pool_of(n)=='ANY') if dead else 0.0
    lock=[0.0]*nd; eff=[0.0]*nd; Sp=[0.0]*nd
    for K in range(nd):
        S_=Sp[K]=Fc[K]+sum(Sc[V][K] for V in tied)-sum(G[g][K] for g in gl)-Ad[K]
        viol=max([0.0]+[sum(G[g][K] for g in sg)-sum(Sc[V][K] for V in U) for sg,U in subs])-Fc[K]
        eff[K]=min(S_-Dc[K],-viol) if viol>1e-6 else S_-Dc[K]
        lock[K]=max(0.0,S_-eff[K])
    al_=[alloc(K) for K in range(nd)]
    pools={**{V:[al_[K][V] for K in range(nd)] for V in tied},'any vendor':[-x for x in Ad],'unknown vendor':Fc}
    first=next((k for k in range(nd) if lock[k]>0.5 and eff[k]<-0.5 and Sp[k]>-0.5),None)   # the rule makes the day short (pooled not short)
    dT=[x for x in sup['ANY']] if dead else [0.0]*nd
    if dead and nd: dT[0]-=dst                                        # arrivals only (day 1 also holds the stock)
    return dict(lock=lock,eff=eff,pools=pools,who={**{V:who[V] for V in tied},'any vendor':who['ANY'],'unknown vendor':[]},first=first,
                sup={lab:sum(x) for lab,x in (('any vendor' if p=='ANY' else 'unknown vendor' if p=='FREE' else p,v) for p,v in sup.items())},
                dead_st=dst,st_by=({n:float(q or 0.0) for n,q in stock.items() if pool_of(n)=='ANY'} if dead else {}),dead_T=dT,dead_cum=list(Dc),dead=bool(dead and any(abs(x)>1e-9 for x in sup['ANY'])),tied=list(tied))
def ata_days(txt, need):
    """'Arrival at plant (Final AtA)' text with every PPC arrival QUANTITY replaced by its days of cover = qty / Usage/Day (user 25-Nov-2026:
       "replace arrival quantity by cover days of that quantity"; DMS BOM Weekly + web page only - the ALERT / email keep quantities).
       'Sun 29-Nov: 93,480' -> 'Sun 29-Nov: 16.7 days'; 'Passed, not received: 1,234 (..)' and '+N more:  218,880  (..)' the same.
       No Usage/Day = text unchanged. PO / PR ETA texts are not touched (they keep quantities)."""
    if not txt or not need or float(need)<=1e-9: return txt
    f=lambda m_: f"{m_.group(1)}{float(m_.group(2).replace(',',''))/float(need):,.1f} days"
    t=re.sub(r'((?:Mon|Tue|Wed|Thu|Fri|Sat|Sun) \d{2}-[A-Za-z]{3}: )([\d,]+(?:\.\d+)?)',f,str(txt))
    t=re.sub(r'(Passed, not received: )([\d,]+(?:\.\d+)?)',f,t)
    return re.sub(r'(\+\d+ more: +)([\d,]+(?:\.\d+)?)',f,t)
def eta_days(txt, need):
    """'ETA - PO' / 'ETA - PR' text (`pr_po_text`) with every quantity replaced by its days of cover = qty / Usage/Day (user 25-Nov-2026:
       "also pr and po qty make it cover day with that qty"; same rule as `ata_days`). '10-Jan 400,000' -> '10-Jan 11.2 days';
       'Date passed: 1,234 (..)', '+N more: 5,000 (..)' and 'No date: 800' the same. No Usage/Day = text unchanged. DMS BOM Weekly only."""
    if not txt or not need or float(need)<=1e-9: return txt
    f=lambda m_: f"{m_.group(1)}{float(m_.group(2).replace(',',''))/float(need):,.1f} days"
    t=str(txt)
    for pat in (r'(Date passed: )([\d,]+(?:\.\d+)?)',r'(\+\d+ more: +)([\d,]+(?:\.\d+)?)',r'(No date: )([\d,]+(?:\.\d+)?)'):
        t=re.sub(pat,f,t)
    return re.sub(r'(\b\d{2}-[A-Za-z]{3} )([\d,]+(?:\.\d+)?)(?![\d.]| days)',f,t)
def ata_show(txt, codes, need):
    """'Arrival at plant (Final AtA)' as SHOWN - ONE rule for EVERY output (user 26-Nov-2026: "Arrival at plant (Final AtA) in 6_email
       should be same as 4_Status and 1_MRP"): 1_MRP ALERT / Action List (-> MRP by Line), 4_Status BOM Weekly (copied from the ALERT),
       5_Web, 6_Email. = the 4_Status rule of 25-Nov-2026: PPC In_Transit arrivals (`_ata_text`) with every quantity as its days of cover
       (`ata_days`); no arrival -> the dated open PO lines, then the PR lines ('PO ETA: dd-Mmm N days  ;  PR ETA: ...', `pr_po_text`,
       `eta_days`); nothing on any = None (blank). Information only - in no calculation."""
    if txt and str(txt).strip(): return ata_days(txt,need)
    codes=[c_ for c_ in (codes or []) if c_]
    t=[k_+' ETA: '+eta_days(x_,need) for k_ in ('PO','PR') for x_ in [pr_po_text(codes,kind=k_,dated_only=True) if codes and PRPO_LINES else '-'] if x_ and x_!='-']
    return '  ;  '.join(t) or None
def _ata_text(shs, d0):
    """'Arrival at plant (Final AtA)' text of a list of shipments, one line each: dates already passed (not yet received) as one total,
       the next ATA_SHOW arrival dates (two per line), the rest as one total. None when nothing is on the way."""
    by_d={}
    for s_ in shs:
        if s_.get('raw') is not None: by_d[s_['raw']]=by_d.get(s_['raw'],0.0)+s_['qty']
    past_=[(d_,q_) for d_,q_ in sorted(by_d.items()) if q_ and d_<d0]; fut_=[(d_,q_) for d_,q_ in sorted(by_d.items()) if q_ and d_>=d0]
    ln_=[]
    if past_: ln_.append(f"Passed, not received: {sum(q for _,q in past_):,.0f}"+(f" ({past_[0][0]:%d-%b})" if len(past_)==1 else f" ({past_[0][0]:%d-%b} → {past_[-1][0]:%d-%b})"))
    nx_=[f"{d_:%a %d-%b}: {q_:,.0f}" for d_,q_ in fut_[:ATA_SHOW]]
    for i_ in range(0,len(nx_),2): ln_.append('   |   '.join(nx_[i_:i_+2]))
    if len(fut_)>ATA_SHOW: rest_=fut_[ATA_SHOW:]; ln_.append(f"+{len(rest_)} more:  {sum(q for _,q in rest_):,.0f}  ({rest_[0][0]:%d-%b} → {rest_[-1][0]:%d-%b})")
    return '\n'.join(ln_) or None
def vp_dead(m):
    """Stock today / arrivals per day of material m that no customer may use (Settings ALLOWED VENDORS, 1_MRP rule = vendor_pools):
       (dead stock, [dead arrivals per day], is_dead(supplier name)). (0, None, None) when the rule leaves everything usable."""
    vp=VP_MRP.get(m) or {}
    if not vp.get('dead'): return 0.0,None,None
    def is_dead(n):
        n=' '.join(str(n or '').split()).upper()
        return n not in VP_UNKNOWN and not any(_vend_hit(V,n) for V in vp.get('tied') or [])
    return float(vp.get('dead_st') or 0.0),(vp.get('dead_T') if any(abs(v)>1e-9 for v in (vp.get('dead_T') or [])) else None),is_dead
def vp_nc(m, ships=None, d0=None):
    """SHOW IT, DON'T COUNT IT (user 28-Nov-2026: "it is a physical fact ... ignore means don't count the stock and the arrival in
       calculation because it is no demand yet"): stock / in transit of material m that no customer with demand may use (`vp_dead`).
       Every shown stock / In transit / SAP / count / overdue / Arrival at plant field includes it; every calculation (cover, stock-zero,
       Inv. Days, projection, class) leaves it out. Returns dict(st=stock, it=in transit total, sups=[supplier names], txt=one line
       'Not counted (no customer demand yet): stock N + in transit N (dates)', also the last line of 'Arrival at plant') or None."""
    dst,dT,is_dead=vp_dead(m)
    if not is_dead: return None
    sh_=[s_ for s_ in ((ships or {}).get(m) or []) if s_.get('raw') is not None and is_dead(s_.get('sup'))]
    vp=VP_MRP.get(m) or {}
    sups=sorted({str(s_.get('sup')) for s_ in sh_ if s_.get('sup')}|{str(n) for n,q in (vp.get('st_by') or {}).items() if is_dead(n) and abs(q)>1e-9})
    sups=[n for n in sups if not any(n!=o_ and n.upper() in o_.upper() for o_ in sups)]   # 'FLAT' + 'FLAT (HONG KONG) CO., LIMITED' = one name
    it=sum(dT or [])
    if dst<=1e-9 and it<=1e-9: return None
    t_=[]
    if dst>1e-9: t_.append(f"stock {dst:,.0f}")
    if it>1e-9:
        ds_=sorted(s_['raw'] for s_ in sh_)
        t_.append(f"in transit {it:,.0f}"+(f" ({ds_[0]:%d-%b})" if ds_ and ds_[0]==ds_[-1] else (f" ({ds_[0]:%d-%b} → {ds_[-1]:%d-%b})" if ds_ else '')))
    h_='Not counted (no customer demand yet'+(': '+', '.join(sups) if sups else '')+'): '
    return dict(st=float(dst),it=float(it),sups=sups,txt=h_+' + '.join(t_))
def _mrp_vendor_pools(b, P, T, stk, rules, ships, d0, nd):
    """vendor_pools on the 1_MRP projection of one material: demand per customer (b['cust_line'], OEM already out), stock by
       supplier (b['sup']: MB52 batches / count by supplier), in transit by supplier (b['sup'] transit, else the PPC In_Transit shipments);
       what cannot be named goes to 'unknown vendor' (stock) / 'not split by customer' (demand), so the pools add up to S."""
    m=b['mat']; dem={}; nv={}
    for k_,v_ in (b.get('cust_line') or {}).items():
        cu=(k_[0] if isinstance(k_,tuple) else k_); nm=' '.join(str(cu or '').split())
        a_=dem.setdefault(nm,[0.0]*nd)
        for k in range(min(nd,len(v_))): a_[k]+=float(v_[k] or 0.0)
        nv[nm]=list(rules.get((_norm(nm),m)) or []) or None
    r_=[P[k]-sum(v[k] for v in dem.values()) for k in range(nd)]
    if any(abs(x)>1e-6 for x in r_): dem[None]=r_
    st={}; av={}
    if b.get('sup'):
        for sp in b['sup']:
            n=str(sp.get('name') or ''); st[n]=st.get(n,0.0)+float(sp.get('stock') or 0.0)
            av.setdefault(n,[0.0]*nd); av[n]=[av[n][k]+float((sp.get('transit') or [0.0]*nd)[k] if k<len(sp.get('transit') or []) else 0.0) for k in range(nd)]
    else:
        for s_ in ships.get(m,[]):
            if s_.get('date') is None: continue
            k=(s_['date']-d0).days
            if 0<=k<nd: av.setdefault(str(s_.get('sup') or ''),[0.0]*nd)[k]+=float(s_['qty'])
    rs=stk-sum(st.values())
    if abs(rs)>1e-6: st['']=st.get('',0.0)+rs
    ra=[T[k]-sum(v[k] for v in av.values()) for k in range(nd)]
    if any(abs(x)>1e-6 for x in ra): av['']=[av.get('',[0.0]*nd)[k]+ra[k] for k in range(nd)]
    return vendor_pools(dem,st,av,nd,lambda c: nv.get(c))
def _wsa_sheet_series(m, d0, nd):
    """Per WSA planning sheet of material m (step 2, aligned file): {sheet: (stock today, end-of-day projection)} = stock + PPC arrivals
       (column T) - Demand, minus the sheet's own vendor lock (column AK) - the build_alert_rows rule for a material on 2+ sheets."""
    dts=[(d0+pd.Timedelta(days=k)).to_pydatetime() for k in range(nd)]; out={}
    for g in WSA_REFS[0] or []:
        if m not in (g.get('codes') or []) or not g.get('dem_by_date'): continue
        bb=g.get('bal_by_date') or {}; st=bb.get(dts[0])
        if st is None: st=bb.get(g['asof'])
        st=float(st or 0.0); cur=st; S=[]
        for d in dts:
            cur=cur-float((g.get('dem_by_date') or {}).get(d) or 0.0)+float((g.get('ppcit_by_date') or {}).get(d) or 0.0)
            S.append(cur-float((g.get('vlock_by_date') or {}).get(d) or 0.0))
        out[g['sheet']]=(st,S)
    return out
def _vp_from_sheets(S, sheets):
    """vendor_pools-like result of a material split over 2+ WSA sheets: short when any sheet is short (its customers may not use the
       other sheet's suppliers) - eff = the short sheets together (never above the pooled projection S); lock = S - eff."""
    nd=len(S); eff=[0.0]*nd; lock=[0.0]*nd
    for k in range(nd):
        v=[sv[1][k] for sv in sheets.values() if k<len(sv[1])]
        neg=sum(x for x in v if x<0)
        eff[k]=min(S[k],neg) if neg<-1e-6 else S[k]; lock[k]=max(0.0,S[k]-eff[k])
    return dict(lock=lock,eff=eff,first=next((k for k in range(nd) if lock[k]>0.5),None),sheets=sheets,pools={sh:sv[1] for sh,sv in sheets.items()},who={},sup={})
def vp_console(VP, d0):
    """One console text for vendor_pools results {name: vp}: the rule and whether it moved a stock-zero day."""
    out=[]
    for nm,vp in VP.items():
        lk=vp['first']
        rt_=('split by WSA sheet as the WSA daily: '+', '.join(f"'{sh}' stock {sv[0]:,.0f}" for sh,sv in vp['sheets'].items())) if vp.get('sheets') else vp_rules_txt(vp)
        out.append(f"{nm} [{rt_}] -> "+(f"the rule makes {d0+pd.Timedelta(days=lk):%d-%b} a stock-zero day (the allowed vendors cannot cover every customer; pooled stock still positive)" if lk is not None else 'no stock-zero day changes'))
    return '; '.join(out)
def vp_rules_txt(vp):
    """Short text of a vendor_pools result: who uses which vendor (Settings) and the supply of each."""
    t=[]
    for lab,cs in vp['who'].items():
        if lab=='unknown vendor' and not vp['sup'].get(lab): continue
        if lab=='any vendor' and not cs and not vp['sup'].get(lab): continue
        t.append(f"{lab}: {', '.join(cs) or ('every customer' if lab=='unknown vendor' else 'no customer')} (stock + arrivals {vp['sup'].get(lab,0):,.0f})")
    return '; '.join(t)
WSA_PUT={}   # PPC In_Transit line -> the 2_WSA sheet it is put on (material on 2+ sheets), for 3_Inbound 'Shown on'
def _it_key(x): return '|'.join(str(x.get(k_) or '') for k_ in ('supplier','invoice','po','description','qty','mat','arrive'))
def wsa_put_of(x):
    """Sheet 2_WSA put this PPC In_Transit line on (step 1 = this process, step 2 = env UPD_WSA_PUT set by step 1), else None."""
    if not WSA_PUT and os.environ.get('UPD_WSA_PUT'):
        import json as _js
        try: WSA_PUT.update(_js.loads(os.environ['UPD_WSA_PUT']))
        except ValueError: pass
    return WSA_PUT.get(_it_key(x))
def _wsa_shared(blocks, AD, label_of, since=None):
    """A material on 2+ planning sheets (back glass 333: 'Glass G12R Topcon SB Back' + 'Glass G12R Topcon Subanci Back').
       User rule, Nov 2026: 'you have the demand of each customer - follow the MRP daily settings and definitions to calculate the
       WSA daily, so you know which supplier goes with which customer':
       - customers: each (customer, line) of the material goes to ONE sheet - the sheet with numbers in its customer column on the
         rows from the 1st of the month (`since`; the planner's sheets keep old customers on 2025 rows), else the sheet whose
         customers have the same first allowed vendor (Settings ALLOWED VENDORS), else the first sheet; demand not split by
         customer stays on the first sheet;
       - stock by supplier, as the 1_MRP supplier rows (MB52 batches / count by supplier): a supplier's stock goes to the sheet
         whose customers have it as FIRST allowed vendor (2+ sheets: by their demand); a supplier no customer is tied to and
         'Unassigned' stock go to the sheets with 'any vendor' customers (by demand), else the first sheet;
       - PPC In Transit arrivals: the sheet where you typed the PO + BL# when its customers may use that supplier (tied to it, or
         any vendor), else by supplier like the stock.
       Returns {material: dict(sheets, first, of={key: block}, stock={block: qty}, by_v={block: {vendor: qty}},
       cust={block: {vendor or '': [customers]}}, arr={block: [qty per day]}, sup={block: {day: {supplier names}}})}."""
    where={}
    for j,B in enumerate(blocks):
        for m in B['codes']:
            if m in AD['blocks'] and j not in where.setdefault(m,[]): where[m].append(j)
    shared={m:js for m,js in where.items() if len(js)>1}
    if not shared: return {}
    rules=_vendor_rules(AD); nd=len(AD['dates']); dix={d:i for i,d in enumerate(AD['dates'])}; PPC=AD.get('ppc') or {}
    def labs_of(B):
        out={}
        for d,r in B['dates'].items():
            for c,lab_ in B['cust']:
                if _num(_gcell(B['rv'],r,c)): out.setdefault(since is None or d>=since,set()).add(_norm(lab_))
        return out.get(True) or out.get(False) or set()          # the current customers; a sheet with none from `since`: all rows
    labs={j:labs_of(blocks[j]) for js in shared.values() for j in js}
    res={}
    for m,js in shared.items():
        a=AD['blocks'][m]; first=js[0]
        vend=lambda key: (rules.get((key[0],m)) or [None])[0]
        dem={key:sum(v) for key,v in a['cl'].items() if key not in a['oem']}
        of={}
        for key in a['cl']:
            mo=AD['mods'].get(key); dn=a['disp'].get(key,key)
            mis_=mo['mis'] if mo else dn[0]; lab=label_of(mis_,mo['line'] if mo else dn[1]) or mis_
            hit=[j for j in js if _norm(lab) in labs[j]]
            if len(hit)==1: of[key]=hit[0]
        for key in a['cl']:
            if key in of: continue
            v=vend(key); hit=sorted({of[k] for k in of if v and vend(k)==v})
            of[key]=hit[0] if len(hit)==1 else first
        claim={}; anyw={}; cust={j:{} for j in js}
        for key,j in of.items():
            v=vend(key); w_=dem.get(key,0.0)+1e-9
            if key not in a['oem']:
                if v: claim.setdefault(v,{})[j]=claim.get(v,{}).get(j,0.0)+w_
                else: anyw[j]=anyw.get(j,0.0)+w_
                nm_=(AD['mods'].get(key) or {}).get('mis') or a['disp'].get(key,key)[0]     # OEM customers are in no calculation: not listed
                lst=cust[j].setdefault(v or '',[])
                if nm_ and nm_ not in lst: lst.append(nm_)
        def tied(name):
            out={}
            for v,wd in claim.items():
                if _vend_hit(v,name):
                    for j,w_ in wd.items(): out[j]=out.get(j,0.0)+w_
            return out
        st={j:0.0 for j in js}; by_v={j:{} for j in js}
        def put(q,name,wts):
            s_=sum(wts.values())
            for j,w_ in wts.items(): st[j]+=q*w_/s_; by_v[j][name]=by_v[j].get(name,0.0)+q*w_/s_
        named=[(n,q) for n,q in (a.get('sup') or []) if n and n.upper()!='UNASSIGNED' and abs(q)>1e-9]
        tn=sum(q for _,q in named); stock=a['stock']
        if tn>stock+0.5 and tn>0: named=[(n,q*stock/tn) for n,q in named]; tn=stock     # JW at the OEM left out / count: the MRP stock is the total
        for n,q in named: put(q,n,tied(n) or anyw or {first:1.0})
        if abs(stock-tn)>1e-6: put(stock-tn,'other / unassigned',anyw or {first:1.0})
        arr={j:[0.0]*nd for j in js}; sup={j:{} for j in js}; arr_v={j:{} for j in js}
        for x in PPC.get('rows') or []:
            if x.get('mat')!=m: continue
            i=dix.get(x.get('arrive'))
            if i is None: continue
            sn=str(x.get('supplier') or ''); t_=tied(sn); ws_=set(x.get('wsa_sheets') or [])
            tr=[j for j in js if blocks[j]['label'] in ws_ or blocks[j].get('src_label') in ws_]
            ok=[j for j in tr if j in t_ or j in anyw]
            j=ok[0] if ok else (max(t_,key=t_.get) if t_ else (max(anyw,key=anyw.get) if anyw else (tr[0] if tr else first)))
            arr[j][i]+=x['qty_mrp']; sup[j].setdefault(i,set()).add(sn or '?'); WSA_PUT[_it_key(x)]=_out_sheet_name(blocks[j])   # 3_Inbound 'Shown on' = this sheet
            arr_v[j].setdefault(sn,[0.0]*nd)[i]+=x['qty_mrp']
        res[m]=dict(sheets=js,first=first,of=of,stock=st,by_v=by_v,cust=cust,arr=arr,sup=sup,arr_v=arr_v)
    return res
def _shared_txt(S_, jb, blocks):
    """Row 2 note of a sheet that shares its material: whose stock this is (Settings ALLOWED VENDORS) and what the other sheets hold."""
    fmt=lambda d: ', '.join(f'{n} {q:,.0f}' for n,q in sorted(d.items(),key=lambda t:-t[1]) if abs(q)>0.5) or '0'
    who='; '.join(f"{', '.join(v)} -> {k or 'any vendor'}" for k,v in sorted(S_['cust'][jb].items(),key=lambda t:(t[0]=='',t[0])))
    oth='; '.join(f"'{_out_sheet_name(blocks[j])}' {fmt(S_['by_v'][j])}" for j in S_['sheets'] if j!=jb)
    return f"stock by supplier (Settings ALLOWED VENDORS, as 1_MRP): here {fmt(S_['by_v'][jb])} for {who or 'no customer'}; other sheet: {oth}"

WSA_OLD_DIR='old WSA files'; WSA_BACKUP=[None]; WSA_MOVED=[]
WSA_DAMAGED=[None]  # (file, safety copy) when an in-place write failed half-way AND could not be undone - repeated at the end of the run
WSA_LOCKED=[None]   # set when your WSA file could not be replaced (open in Excel / OneDrive): the WSA daily went to WSA_LOCKED[0]
def _put_file(tmp, out):
    """Put tmp in place of out when Windows may lock out (OneDrive sync / Excel). Returns None when done, else the last error (tmp kept).
       1) replace, 5 tries over ~15 s (OneDrive holds a file a few seconds); 2) write the content INTO out (a rename-only lock - OneDrive -
       still allows it; Excel's lock refuses the open, so nothing is touched). Before writing in place, out is copied to '<tmp>.keep'; a
       write that fails half-way is undone from it, and if even that fails the console names the copy (it is kept). Removing tmp after a
       good save is cleanup only - its failure never undoes the save (checker findings Nov 2026)."""
    import shutil
    last=None
    for w_ in (0,1,2,4,8):
        if w_: time.sleep(w_)
        try: os.replace(tmp,out); return None
        except PermissionError as e_: last=e_
    if not os.path.exists(out): return last
    try: f_=open(out,'r+b')                                           # refused while Excel has it open -> the file is not touched
    except OSError as e_: return e_
    keep=tmp+'.keep'
    try: shutil.copyfile(out,keep)
    except OSError as e_: f_.close(); return e_                       # no safety copy -> no in-place write
    try:
        with f_, open(tmp,'rb') as s_:
            shutil.copyfileobj(s_,f_); f_.truncate()
    except Exception as e_:                                           # half written: put the file back as it was
        try: shutil.copyfile(keep,out); os.remove(keep)
        except OSError as e2_:
            WSA_DAMAGED[0]=(out,keep)
            print('#'*100+f"\nWARNING: {os.path.basename(out)} WAS HALF WRITTEN AND COULD NOT BE PUT BACK ({e2_}).\n"
                  f"  Your file as it was before this run: {keep} - close Excel / pause OneDrive and copy it over {os.path.basename(out)}.\n"+'#'*100,flush=True)
        return e_
    for f2_ in (tmp,keep):                                            # cleanup only
        try: os.remove(f2_)
        except OSError as e_: _swallow(0,e_)
    return None
def _wsa_save(tmp, out, stem, d0):
    """Save the rebuilt WSA file as today's 2_WSA (your WSA file is never written - user answer '1', 11-Nov-2026). History: (user, Nov 2026: "I don't see the WSA daily anymore" - the file is on OneDrive and Windows said
       'Access is denied' when the rename replaced it, so no WSA daily was made and step 2 ran on the old layout). _put_file (retry, write
       in place); when the file stays locked (open in Excel): save as '<stem> (aligned <date>).xlsx' next to it - today's WSA daily exists,
       the next run with the file closed reads the newest of the two and puts the result back in '<stem>.xlsx' (the aligned one goes to
       deleted by _wsa_tidy). Returns the path written."""
    WSA_LOCKED[0]=None
    last=_put_file(tmp,out)
    if last is None: return out
    alt=os.path.join(os.path.dirname(os.path.abspath(out)),f'{stem} (2).xlsx')     # today's WSA daily is open in Excel: a second name
    os.replace(tmp,alt)                                               # raises when this fails too (then nothing was saved)
    WSA_LOCKED[0]=(alt,f'{type(last).__name__}: {last}'); return alt
def _same_file(a, b):
    import filecmp
    try: return os.path.getsize(a)==os.path.getsize(b) and filecmp.cmp(a,b,shallow=False)
    except OSError: return False
def delete_old_outputs(folder, inbound_today):
    """User, Nov 2026: "Too many output, can old output be deleted after the script produces the new output file?" - once today's files
       are saved, the files of EARLIER dates go (1_MRP / 2_WSA / 6_Email already did): 4_Status_<date>.xlsx,
       7_Summary_<date>.txt and 3_Inbound_<date>.xlsx (all of them - the user does not use the 'Changes' sheet, which compares with
       the previous Inbound file; today's file is read for it before this runs). A kind is only cleaned when today's file of that kind exists. --keep-old keeps everything."""
    gone=[]; kept=[]
    if globals().get('KEEP_OLD') or not os.path.isdir(folder): return gone,kept
    today_=f'{dt.date.today():%Y-%m-%d}'; names=sorted(os.listdir(folder))
    def _old(rx):
        return [(m_.group(1),n_) for n_ in names for m_ in [re.fullmatch(rx,n_)] if m_ and m_.group(1)<today_]
    todo=[]
    if os.path.exists(os.path.join(folder,outfile_name('dms','xlsx'))): todo+=[n_ for d_,n_ in _old(out_rx('dms')+r'_(\d{4}-\d{2}-\d{2})\.xlsx')]
    if os.path.exists(os.path.join(folder,outfile_name('sum','txt'))): todo+=[n_ for d_,n_ in _old(out_rx('sum')+r'_(\d{4}-\d{2}-\d{2})\.txt')]
    if inbound_today and os.path.basename(inbound_today)==outfile_name('inb','xlsx') and os.path.dirname(os.path.abspath(inbound_today))==os.path.abspath(folder) and os.path.exists(inbound_today):
        todo+=[n_ for d_,n_ in _old(out_rx('inb')+r'_(\d{4}-\d{2}-\d{2})\.xlsx')]      # all of them: the user does not use the 'Changes' sheet (12-Nov-2026)
    for n_ in todo:
        try: os.remove(os.path.join(folder,n_)); gone.append(n_)
        except OSError: kept.append(n_)
    return gone,kept
def _wsa_backup(out, ran_on, d0):
    """No backup copies any more (user, Nov 2026: "old WSA files folder? I don't want this, too complicated"): the WSA file lives on
       OneDrive, whose version history keeps the earlier versions; the save itself is safe (atomic replace, or an in-place write with a
       temporary safety copy - _put_file). Kept as a stub so callers stay unchanged."""
    return None
def _wsa_tidy(out, stem, ran_on=None, ran_mtime=None):
    """Your WSA file is never changed (user, Nov 2026, answer '1'); the WSA daily is '2_WSA_<date>.xlsx' next to it. Like the older
       1_MRP files, 2_WSA files of EARLIER dates are deleted (today's file and its '(2)' copy stay). Old '<your file> (aligned
       <date>).xlsx' files of earlier versions are only named in a NOTE (never read when your own file is there). The 'old WSA files'
       folder of earlier versions is removed - only the names the script made there (stem = your file's name)."""
    d_=os.path.dirname(os.path.abspath(out)); gone=[]
    today_=f'{dt.date.today():%Y-%m-%d}'
    for n_ in sorted(os.listdir(d_)):
        p_=os.path.join(d_,n_); m_=re.fullmatch(out_rx('wsa')+r'_(\d{4}-\d{2}-\d{2})( \(2\))?\.xlsx',n_)
        if m_ and m_.group(1)<today_ and p_!=os.path.abspath(out) and not globals().get('KEEP_OLD'):   # --keep-old keeps them, like 1_MRP
            try: os.remove(p_); gone.append(n_)
            except OSError as e_: print(f'NOTE: {n_} could not be deleted ({e_}) - close it in Excel; it can be deleted by hand')
        elif n_.lower().endswith('.xlsx') and _ALIGN_SFX.search(os.path.splitext(n_)[0]) and _ALIGN_SFX.sub('',os.path.splitext(n_)[0])==stem and p_!=(os.path.abspath(ran_on) if ran_on else None):
            print(f"NOTE: {n_} = an older WSA output (not read, not changed) - the WSA daily is now 2_WSA_<date>.xlsx; delete it when you no longer need it")
        elif n_==f'{stem}.xlsx.tmp.keep':                             # safety copy of an earlier version's in-place write
            try: os.remove(p_); gone.append(n_)
            except OSError as e_: _swallow(0,e_)
    od=os.path.join(d_,WSA_OLD_DIR)
    if os.path.isdir(od):
        mine_=re.compile(re.escape(stem)+r'( \(aligned \d{4}-\d{2}-\d{2}\))?( \(before \d{4}-\d{2}-\d{2} \d{6}\))?\.xlsx',re.I)
        for n_ in os.listdir(od):
            if mine_.fullmatch(n_) and ('(before ' in n_ or '(aligned ' in n_):
                try: os.remove(os.path.join(od,n_))
                except OSError as e_: _swallow(0,e_)
        try: os.rmdir(od); gone.append(f"folder '{WSA_OLD_DIR}'")
        except OSError: print(f"NOTE: folder '{WSA_OLD_DIR}' kept - it holds files the script did not make")
    return gone
def align_wsa(src, AD, today=None, name_from=None):
    """Write the rebuilt WSA file as '2_WSA_<date>.xlsx' next to your WSA file (yours is only read). Returns (path, report rows)."""
    ALIGN_STAT['ppc_sheets']=[0,0]
    import zipfile, posixpath
    from collections import defaultdict
    from lxml import etree
    from openpyxl.formula.translate import Translator
    M_='http://schemas.openxmlformats.org/spreadsheetml/2006/main'; M='{%s}'%M_
    RNS='http://schemas.openxmlformats.org/officeDocument/2006/relationships'; PNS='http://schemas.openxmlformats.org/package/2006/relationships'
    CTN='http://schemas.openxmlformats.org/package/2006/content-types'
    from openpyxl.utils import column_index_from_string as _CI
    def CX_order(s__):
        m_=re.match(r'<c r="([A-Z]+)',s__); return _CI(m_.group(1)) if m_ else 0
    _ug=lambda u: next((g for g,ws in (('pc',('PC','PCS','NOS','SET','EA')),('m2',('M2','SQM','SQ')),('roll',('ROL','ROLL'))) if any(str(u).upper().startswith(w) for w in ws)),str(u).upper()[:2])
    d0=pd.Timestamp(today or base_d0()).normalize(); keep_from=d0.replace(day=1)
    dix={d:i for i,d in enumerate(AD['dates'])}; nd=len(AD['dates'])
    serial=lambda d: (pd.Timestamp(d).normalize()-pd.Timestamp('1899-12-30')).days
    nf_=name_from or src; stem=_ALIGN_SFX.sub('',os.path.splitext(os.path.basename(nf_))[0])
    own_stem=stem; stem=outfile_name('wsa')               # user, Nov 2026 (answer '1'): YOUR WSA FILE IS NEVER CHANGED -
    out=os.path.join(os.path.dirname(os.path.abspath(nf_)),f'{stem}.xlsx')     # the WSA daily is a new file each day (was: saved over your file, option 2):
    wv=load_workbook(src,data_only=True,read_only=True); wf=load_workbook(src,read_only=True)
    blocks,GS=_glass_split(_wsa_read(wv,wf),AD)
    look,lookcu=_wsa_lookup(src)
    lab_cache={}
    def label_of(mis,line):
        k=(_norm(mis),line)
        if k in lab_cache: return lab_cache[k]
        lab=next((look[(p,_norm(mis))] for p in _plant_keys(line) if (p,_norm(mis)) in look),None)
        if lab is None and _norm(mis) in lookcu: lab=lookcu[_norm(mis)].most_common(1)[0][0]
        lab_cache[k]=lab; return lab
    T=ALIGN_SHEET; q=lambda s_: str(s_).replace('"','""')
    z=zipfile.ZipFile(src); names_=set(z.namelist())
    wbx=etree.fromstring(z.read('xl/workbook.xml')); rels=etree.fromstring(z.read('xl/_rels/workbook.xml.rels'))
    tg={r.get('Id'):r.get('Target') for r in rels}
    shs=wbx.find(M+'sheets'); sheet_el={s.get('name'):s for s in shs}
    part={s.get('name'):('xl/'+tg[s.get('{%s}id'%RNS)].lstrip('/').replace('xl/','',1)) for s in shs}
    for o_,n_ in GS['renames'].items():          # front + back glass sheet: the front glass keeps the sheet, under its new name
        if o_ in sheet_el and n_ not in sheet_el: sheet_el[o_].set('name',n_); sheet_el[n_]=sheet_el[o_]; part[n_]=part[o_]
    STY=_Styles(z.read('xl/styles.xml')); yellow=STY.yellow_xfs()
    report=[]; skipped=[]; unmatched=Counter(); misonly=Counter(); other=Counter(); unit_notes=[]; dropped=set(); arch_new=[]
    # ---------- 1) the plan table rows of every block (module plan x BOM) and the customer columns ----------
    work=[]; tab=[]; lab_tot=Counter(); lab_oem=set(); hist_lab={}
    SHR=_wsa_shared(blocks,AD,label_of,keep_from)   # a material on 2+ sheets: customers, stock and arrivals by Settings ALLOWED VENDORS
    VRULES=_vendor_rules(AD); VP_WSA.clear()
    for jb,B in enumerate(blocks):
        mats=[c for c in B['codes'] if c in AD['blocks']]
        mine=lambda m,key: m not in SHR or SHR[m]['of'].get(key,SHR[m]['first'])==jb
        if not mats: skipped.append(f"{B['label']} (codes {', '.join(B['codes']) or 'none'} not in the MRP - sheet left as it is)"); continue
        U_=str(AD['blocks'][mats[0]].get('uom') or '').upper()
        if B['uom'] and U_ and _ug(B['uom'])!=_ug(U_) and not ({_ug(B['uom']),_ug(U_)}=={'pc','roll'}):
            unit_notes.append(f"{B['label']}: unit '{B['uom']}' -> '{U_}'")
        bid=B['label']; dropped.update(f"{B['sheet']}: {h}" for h in B['dropped'])
        stock=sum((SHR[m]['stock'][jb] if m in SHR else AD['blocks'][m]['stock'])
                  +(AD['blocks'][m]['jw'] if WSA_OEM_IN_BALANCE and (m not in SHR or SHR[m]['first']==jb) else 0.0) for m in mats)
        def want_m(m,i):                                      # MRP plan of the material; shared: its customers on this sheet (+ the unsplit rest on the first sheet)
            a_=AD['blocks'][m]; tot_=a_['plan'][i]+(a_['plan_oem'][i] if WSA_OEM_IN_BALANCE else 0.0)
            if m not in SHR: return tot_
            ks_=[k_ for k_ in a_['cl'] if WSA_OEM_IN_BALANCE or k_ not in a_['oem']]
            return (sum(a_['cl'][k_][i] for k_ in ks_ if mine(m,k_))
                    +((tot_-sum(a_['cl'][k_][i] for k_ in ks_)) if SHR[m]['first']==jb else 0.0))
        want=[sum(want_m(m,i) for m in mats) for i in range(nd)]
        oemk=set().union(*[AD['blocks'][m]['oem'] for m in mats])
        keys=sorted(set().union(*[{k_ for k_ in AD['blocks'][m]['cl'] if mine(m,k_)} for m in mats]),key=lambda k:(k in oemk,k))
        brow=[]; got=[0.0]*nd
        for key in keys:
            per={m:AD['blocks'][m]['cl'].get(key) for m in mats if mine(m,key) and any(AD['blocks'][m]['cl'].get(key) or [])}
            use=[sum(v[i] for v in per.values()) for i in range(nd)]
            if not any(u>1e-12 for u in use): continue
            is_oem=key in oemk; inbal='No' if (is_oem and not WSA_OEM_IN_BALANCE) else 'Yes'
            mo=AD['mods'].get(key)
            dn=next((AD['blocks'][m]['disp'][key] for m in per if key in AD['blocks'][m]['disp']),key)
            mis_d=mo['mis'] if mo else dn[0]; line_d=mo['line'] if mo else dn[1]
            lab=label_of(mis_d,line_d)
            if lab is None: unmatched[f"{mis_d} / {line_d}"]+=1; lab=mis_d          # not on the Lookup sheet: its own column
            if mo is None and inbal=='Yes': misonly[f"{bid}: {mis_d} ({line_d or 'no line'})"]+=1
            srcs=sorted({('PPC BOM' if f'your BOM ({dn[0]})' in AD['blocks'][m]['basis'] else 'MIS BOM') for m in per})
            for i in range(nd):
                u=use[i]; d=AD['dates'][i]
                if u<=1e-12 or d<d0: continue
                m_=mo['v'][i] if mo else 0.0; ok=m_>1e-12
                brow.append(dict(date=d,serial=serial(d),line=line_d,cust=', '.join(mo['cust']) if mo else '',mis=mis_d,block=bid,col=lab,inbal=inbal,
                                 mods=m_ if ok else None,qty=(u/m_) if ok else None,req=u,mw=((mo.get('mw') or [0.0]*nd)[i] if ok else None),
                                 unit=U_ or B['uom'],codes='+'.join(m for m in per if per[m][i]>1e-12),src=(' + '.join(srcs) if ok else 'MIS plan (no modules)')))
                if inbal=='Yes': got[i]+=u
        for i in range(nd):                                   # anything the customer split does not explain (keeps Demand = MRP plan)
            if abs(want[i]-got[i])>1e-6 and AD['dates'][i]>=d0:
                d=AD['dates'][i]; other[bid]+=1
                brow.append(dict(date=d,serial=serial(d),line='',cust='',mis='',block=bid,col='(not split by customer)',inbal='Yes',mods=None,qty=None,
                                 req=want[i]-got[i],mw=None,unit=U_ or B['uom'],codes='+'.join(mats),src='MRP plan'))
        for x in brow:
            lab_tot[x['col']]+=x['mods'] or 0.0
            if x['inbal']=='No': lab_oem.add(x['col'])
        for d,r in B['dates'].items():                        # customers with numbers on the rows kept before today
            if keep_from<=d<d0:
                for c,lab_ in B['cust']:
                    if _num(cv(B['rv'],r,c)): hist_lab.setdefault(_norm(lab_),lab_)
        ufix=bool(B['uom'] and U_ and _ug(B['uom'])!=_ug(U_) and not ({_ug(B['uom']),_ug(U_)}=={'pc','roll'}))   # your label is another unit: the numbers are in the MRP (SAP) unit
        kd_,km_=_kind_of(B['mname'] or ''),_kind_of(' '.join(str(AD['blocks'][m].get('common') or '') for m in mats))
        mfix=(AD['blocks'][mats[0]].get('desc') or AD['blocks'][mats[0]].get('common') or mats[0]) if (kd_ and km_ and kd_!=km_) else None   # e.g. 'Frame 72 cells' on the EVA sheet
        if not mfix and len(mats)==1 and mats[0] in CODE_SWAP: mfix=AD['blocks'][mats[0]].get('desc') or None   # tape planned under your BOM code: its description (roll length)
        if mfix: unit_notes.append(f"{B['label']}: material '{str(B['mname']).strip()}' -> '{mfix}'")
        vp=None                                               # Settings ALLOWED VENDORS on this sheet (one material per sheet)
        if len(mats)==1 and any(k_[1]==mats[0] for k_ in VRULES):
            m=mats[0]; a_=AD['blocks'][m]; dem_={}; nvend={}
            for k_ in a_['cl']:
                if not mine(m,k_) or (k_ in a_['oem'] and not WSA_OEM_IN_BALANCE): continue
                mo=AD['mods'].get(k_); nm_=(mo['mis'] if mo else a_['disp'].get(k_,k_)[0]) or k_[0]
                v_=dem_.setdefault(nm_,[0.0]*nd)
                for i in range(nd): v_[i]+=a_['cl'][k_][i]
                nvend[nm_]=list(VRULES.get((k_[0],m)) or VRULES.get((_norm(nm_),m)) or []) or None
            rest_=[want[i]-sum(v_[i] for v_ in dem_.values()) for i in range(nd)]
            if any(abs(x)>1e-6 for x in rest_): dem_[None]=rest_
            if m in SHR: st_={n_:q_ for n_,q_ in SHR[m]['by_v'][jb].items()}; av_={n_:list(v_) for n_,v_ in SHR[m]['arr_v'][jb].items()}
            else:
                named=[(n_,q_) for n_,q_ in (a_.get('sup') or []) if n_ and n_.upper()!='UNASSIGNED' and abs(q_)>1e-9]
                tn=sum(q_ for _,q_ in named)
                if tn>stock+0.5 and tn>0: named=[(n_,q_*stock/tn) for n_,q_ in named]; tn=stock
                st_={}
                for n_,q_ in named: st_[n_]=st_.get(n_,0.0)+q_
                if abs(stock-tn)>1e-6: st_['']=st_.get('',0.0)+stock-tn
                av_={}
                for x in (AD.get('ppc') or {}).get('rows') or []:
                    i=dix.get(x.get('arrive'))
                    if x.get('mat')!=m or i is None: continue
                    av_.setdefault(str(x.get('supplier') or ''),[0.0]*nd)[i]+=x['qty_mrp']
                pa_=((AD.get('ppc') or {}).get('arr') or {}).get(m) or [0.0]*nd
                r_=[pa_[i]-sum(v_[i] for v_ in av_.values()) for i in range(nd)]
                if any(abs(x)>1e-6 for x in r_): av_.setdefault('',[0.0]*nd); av_['']=[av_[''][i]+r_[i] for i in range(nd)]
            vp=vendor_pools(dem_,st_,av_,nd,lambda c: nvend.get(c))
            if vp: VP_WSA[bid]=vp
        work.append(dict(B=B,mats=mats,bid=bid,stock=stock,brow=brow,U=U_ or str(B['uom']).upper(),jb=jb,ufix=ufix,mfix=mfix,vp=vp,
                         shr=' | '.join(f"{m} {_shared_txt(SHR[m],jb,blocks)}" for m in mats if m in SHR)))
    seen=set(); slots=[]
    for lab in (sorted([l for l in lab_tot if l not in lab_oem and l!='(not split by customer)'],key=lambda l:(-lab_tot[l],l))
                +sorted([l for l in lab_tot if l in lab_oem],key=lambda l:(-lab_tot[l],l))+list(hist_lab.values())):
        if _norm(lab) not in seen: seen.add(_norm(lab)); slots.append(lab)
    while len(slots)<WSA_SLOTS: slots.append('')
    canon={_norm(s_):s_ for s_ in slots if s_}
    for w in work:
        for x in w['brow']: x['col']=canon.get(_norm(x['col']),x['col'])
        tab+=w['brow']
    cols=WSA_FIX1+[(f'c{i}',lab,11) for i,lab in enumerate(slots)]+WSA_FIX2
    CX={k:j+1 for j,(k,_,_) in enumerate(cols)}; LC=L(len(cols)); XL=lambda k: L(CX[k])
    end_d=max([AD['dates'][-1]]+[max(w['B']['dates']) for w in work if w['B']['dates']])
    days=[keep_from+pd.Timedelta(days=k) for k in range((end_d-keep_from).days+1)]
    R0=WSA_HDR+1; rowof=lambda d: R0+(d-keep_from).days; t=rowof(d0); last=rowof(end_d)
    # ---------- 2) styles ----------
    fb=STY.font(bold=True); f14=STY.font(bold=True,size=14); fw=STY.font(bold=True,rgb='FF1F3864'); bd=STY.border()
    yel=STY.fill('FFFFFF00')
    HDRF={'g':STY.fill('FFD8E4BC'),'b':STY.fill('FFB7DEE8'),'s':STY.fill('FFD9D9D9'),'d':STY.fill('FFEBF1DE'),'p':STY.fill('FFFCE4D6')}
    SX=dict(title=STY.xf(font=f14),label=STY.xf(font=fb),info=STY.xf(),note=STY.xf(font=fw),
            date=STY.xf(nf=STY.numfmt('dd-mmm-yy'),border=bd,h='center'),text=STY.xf(border=bd),day=STY.xf(border=bd,h='center'),
            mod=STY.xf(nf=3,border=bd),days=STY.xf(nf=STY.numfmt('0.0'),border=bd,h='right'),mw=STY.xf(nf=STY.numfmt('0.000'),border=bd),
            tmod=STY.xf(nf=3,font=fb),tmw=STY.xf(nf=STY.numfmt('0.000'),font=fb),
            t_date=STY.xf(nf=STY.numfmt('dd-mmm-yy')),t_mod=STY.xf(nf=3),t_qty=STY.xf(nf=STY.numfmt('#,##0.0#####')),
            t_req=STY.xf(nf=STY.numfmt('#,##0.###')),t_mw=STY.xf(nf=STY.numfmt('0.000')),t_head=STY.xf(font=fb,wrap=True,v='center'))
    for k_,fl in HDRF.items(): SX['h'+k_]=STY.xf(font=fb,fill=fl,border=bd,h='center',v='center',wrap=True)
    def qfmt(u): return '#,##0' if _ug(u) in ('pc','roll') or str(u).upper().startswith(('SET','NOS','PCS')) else '#,##0.0'
    for code_ in ('#,##0','#,##0.0'):
        nfid=3 if code_=='#,##0' else STY.numfmt(code_)
        SX['q'+code_]=STY.xf(nf=nfid,border=bd); SX['bal'+code_]=STY.xf(nf=nfid,border=bd)
        SX['typed'+code_]=STY.xf(nf=nfid,font=fb,fill=yel,border=bd); SX['tot'+code_]=STY.xf(nf=nfid,font=fb)
    DX_NEG=STY.dxf('FF9C0006','FFFFC7CE'); DX_TODAY=STY.dxf('FF000000','FFFFF2CC',bold=True)
    hgroup=lambda k: 'b' if k.startswith('c') and k[1:].isdigit() else ('g' if k in ('date','day','bal','plan','dem') else ('s' if k in ('actmod','act','mw','actmw','pdays','pport') else ('p' if k in ('ppcit','ppcsup','ppcpp') else 'd')))
    PPC=AD.get('ppc') or {}; PARR=PPC.get('arr') or {}
    # ---------- 3) one standard sheet per block ----------
    def keyrow(B,d): return B['dates'].get(d)
    # your typed Delivered / Scheduled are in YOUR BOM code's unit (user, 13-Nov-2026: "check my BOM first ... the WSA plan uses that code")
    def ival(B,role,r): return cv(B['rv'],r,B['role'][role]) if (r and role in B['role']) else None
    def ifx(B,role,r): return cv(B['rf'],r,B['role'][role]) if (r and role in B['role']) else None
    def bal_styles(B):
        """style index of the planner's Balance cells on the rows kept before today (to keep their yellow 'typed' look)"""
        want_={f"{L(B['role']['bal'])}{r}" for d,r in B['dates'].items() if keep_from<=d<d0}
        if not want_: return {}
        res={}
        for _,c in etree.iterparse(z.open(part[B['sheet']]),tag=M+'c'):
            if c.get('r') in want_: res[c.get('r')]=c.get('s')
            c.clear()
        return res
    old_arch=[]
    if ARCHIVE_SHEET in wv.sheetnames:
        for r in wv[ARCHIVE_SHEET].iter_rows(min_row=2,values_only=True):
            r=list(r)+[None]*len(ARCH_COLS)
            if r[0] and isinstance(r[1],(dt.datetime,dt.date)): old_arch.append(r[:len(ARCH_COLS)])
    old_arch=GS['arch'](old_arch)                 # records of a block split into front / back glass: one per glass
    akey=lambda r: (str(r[0]),pd.Timestamp(r[1]).normalize(),_num(r[3]),str(r[4] or ''),str(r[10] or ''))
    arch_have={akey(r) for r in old_arch}; arch_old_by={}
    for r in old_arch: arch_old_by.setdefault(str(r[0]),[]).append(r)
    plan_sheets=[]          # (output sheet name, xml, last row)
    out_name={}
    for w in work:
        B=w['B']; bid=w['bid']; brow=w['brow']; qf=qfmt(w['U']); SQ=SX['q'+qf]; STYP=SX['typed'+qf]; SBAL=SX['bal'+qf]; STOT=SX['tot'+qf]
        oname=_out_sheet_name(B)
        out_name[(B['sheet'],B['block'])]=oname
        S_lab=defaultdict(float); S_mod=defaultdict(float); S_req=defaultdict(float); S_mw=defaultdict(float)
        for x in brow:
            s_=x['serial']
            if x['mods'] is not None: S_lab[(s_,str(x['col']).lower())]+=x['mods']
            if x['inbal']=='Yes':
                if x['mods'] is not None: S_mod[s_]+=x['mods']; S_mw[s_]+=x['mw'] or 0.0
                S_req[s_]+=(x['mods']*x['qty'] if x['mods'] is not None and x['qty'] is not None else x['req'])
        colmap={B['role'][k]:XL(k) for k in B['role'] if k in CX}
        for j,(c,lab_) in enumerate(B['cust']):
            k_=next((f'c{i}' for i,s_ in enumerate(slots) if s_ and _norm(s_)==_norm(lab_)),None)
            if k_: colmap[c]=XL(k_)
        bst=bal_styles(B)
        for d,ri in sorted(B['dates'].items()):                   # deleted rows (before the 1st of the month) with a delivery record -> Archive
            if d>=keep_from: continue
            vals={k_:ival(B,k_,ri) for k_ in ('dlv','sch','po','ship','rem','book','ready','etd','bl','vessel','inv','ref')}
            dl_=_num(vals['dlv']) or 0; sch_=_num(vals['sch']) or 0
            if not (dl_>0 or sch_>0 or vals['po'] or vals['bl'] or vals['vessel']): continue
            rem_=str(vals['rem'] or '')
            state='Cancelled' if 'canc' in rem_.lower() else ('Delivered' if dl_>0 else ('Not delivered' if sch_>0 else ''))
            arch_new.append([bid,d.to_pydatetime()]+[vals[k_] for k_ in ('dlv','sch','po','ship','rem','book','ready','etd','bl','vessel','inv','ref')]+[state])
        carried=sum((_num(r_[3]) or 0) for r_ in (arch_old_by.get(bid,[])+[a_ for a_ in arch_new if a_[0]==bid and akey(a_) not in arch_have]) if r_[14]=='Not delivered')
        # old numbers for the console report (today and the next 7 days of the sheet as it was)
        r_t=keyrow(B,d0); old_bal=_num(ival(B,'bal',r_t)); old_d7=sum(_num(ival(B,'dem',keyrow(B,d0+pd.Timedelta(days=k)))) or 0 for k in range(7))
        old_neg=next((d for d in sorted(B['dates']) if d>=d0 and (_num(ival(B,'bal',B['dates'][d])) or 0)<0),None)
        rowsx=[]
        # rows 1-6: material info
        title=' + '.join(dict.fromkeys(str(AD['blocks'][m].get('common') or m) for m in w['mats']))
        r1=_xcell('A1',f'{bid}  -  {title}'+(f"  {B['note']}" if B.get('note') else ''),SX['title'])
        SCH=XL('sch'); ARC=f'SUMIFS({ARCHIVE_SHEET}!$D:$D,{ARCHIVE_SHEET}!$A:$A,"{q(bid)}",{ARCHIVE_SHEET}!$O:$O,"Not delivered")'
        r2=(_xcell('A2',f'Stock on {d0:%d-%b-%y}',SX['label'])+_xcell('C2',round(w['stock'],3),STOT,f'C{t}')+
            _xcell('E2','MB52 + PPC count'+('' if WSA_OEM_IN_BALANCE else ', JW at OEM excluded')+(f" - {w['shr']}" if w['shr'] else ''),SX['note'])+
            _xcell(f"{XL('actmod')}2",f'Open deliveries dated before {keep_from:%d-%b-%y} (Archive, State = Not delivered):',SX['note'])+
            _xcell(f'{SCH}2',carried,STOT,ARC))
        r2=''.join(sorted(re.findall(r'<c [^>]*?(?:/>|>.*?</c>)',r2),key=CX_order))
        r3=(_xcell('A3','Supplier & Material',SX['label'])+_xcell('C3',w['mfix'] or B['mname'] or title,SX['info'])+
            _xcell(f"{XL('ppcit')}3",(f"PPC In Transit = {PPC['src']} sheet {PPC.get('sheet') or 'In_Transit'}: Final Pending Qty. on Final AtA (late = today); material traced by your WSA PO + BL#, description, supplier, RM Code (Inbound sheet 'PPC In Transit')" if PARR else 'PPC In Transit: EMPTY - '+'; '.join(PPC.get('why') or ['no PPC file / In_Transit sheet'])),SX['note']))
        vpn_=w.get('vp')
        r3+=_xcell(f"{XL('vlock')}3",('Settings ALLOWED VENDORS: '+vp_rules_txt(vpn_)+'. Customers without a vendor rule may use the other vendors\' surplus; a customer may use every vendor listed for it, top row first (+ stock of unknown vendor). '
                                       f"Column V (Plant + Port Days, PPC In Transit) subtracts column {XL('vlock')} = stock + arrivals no customer short that day may use (0 = the rule changes nothing that day)."
                                       if w.get('vp') else 'Settings ALLOWED VENDORS: no vendor rule for the customers of this sheet (any vendor).'),SX['note'])
        u4_=w['U'] if w['ufix'] else (B['uom'] or w['U'])          # user, Nov 2026: keep M2 (SAP) - a wrong label ('Pcs' on EVA / EPE 420) shows the MRP unit
        r4=_xcell('A4','Unit of Measurement',SX['label'])+_xcell('C4',u4_,SX['info'])+(_xcell('E4',f"MRP unit: {w['U']}",SX['info']) if w['U'] and _ug(u4_)!=_ug(w['U']) else (_xcell('E4',B['unote'],SX['info']) if B.get('unote') else ''))
        r5=_xcell('A5','Material code(s)',SX['label'])+''.join(_xcell(f'{L(3+i)}5',c_,SX['info']) for i,c_ in enumerate(B['codes']))
        rowsx+= [f'<row r="1" ht="21" customHeight="1">{r1}</row>',f'<row r="2">{r2}</row>',f'<row r="3">{r3}</row>',f'<row r="4">{r4}</row>',f'<row r="5">{r5}</row>']
        # rows 8..: one per day
        drows=[]; bal=None; prevr=None; cum=0.0; cum_p=0.0; tot=defaultdict(float); new_vals={}
        PA_=lambda m: (SHR[m]['arr'][w['jb']] if m in SHR else PARR.get(m)) or [0.0]*nd      # shared material: the arrivals of this sheet
        parr=[sum(PA_(m)[i] for m in w['mats']) for i in range(nd)]
        ALIGN_STAT['ppc_sheets'][1]+=1; ALIGN_STAT['ppc_sheets'][0]+=int(any(parr))
        PSUP=PPC.get('sup') or {}       # supplier names of the PPC In_Transit lines arriving that day
        PS_=lambda m: (SHR[m]['sup'][w['jb']] if m in SHR else PSUP.get(m)) or {}
        psup=[' / '.join(sorted(set().union(*[PS_(m).get(i,set()) for m in w['mats']]))) for i in range(nd)]
        # SHOW IT, DON'T COUNT IT (user 28-Nov-2026, option A): arrivals of a vendor no customer with demand may use stay in column T / U
        # (physical fact); column AK holds them (= `vendor_pools` lock), so column V does not count them
        for d in days:
            r=rowof(d); ri=keyrow(B,d); cells=[]; s_=serial(d)
            cells.append(_xcell(f'A{r}',d.to_pydatetime(),SX['date'])); cells.append(_xcell(f'B{r}',d.strftime('%a'),SX['day']))
            _s=_num(ival(B,'sch',ri)) or 0.0; _rm=str(ival(B,'rem',ri) or ''); _dl=ival(B,'dlv',ri)
            if 'canc' not in _rm.lower() and (_dl is None or _dl==''): cum+=_s     # = SUMPRODUCT(Scheduled, Remarks without 'canc', Delivered blank) from row 8
            if d<d0:                                                    # rows kept before today: the numbers as they were
                st_=bst.get(f"{L(B['role']['bal'])}{ri}") if ri else None
                cells.append(_xcell(f'C{r}',ival(B,'bal',ri),STYP if st_ in yellow else SBAL))
                for k_,sty_ in (('plan',SX['mod']),('dem',SQ)): cells.append(_xcell(f'{XL(k_)}{r}',ival(B,k_,ri),sty_))
                hv={}
                for c,lab_ in B['cust']:
                    k_=next((f'c{i}' for i,s2 in enumerate(slots) if s2 and _norm(s2)==_norm(lab_)),None)
                    v_=_num(cv(B['rv'],ri,c)) if (k_ and ri) else None
                    if v_ is not None: hv[k_]=(hv.get(k_) or 0)+v_          # blank stays blank
                for i,s2 in enumerate(slots): cells.append(_xcell(f"{XL(f'c{i}')}{r}",hv.get(f'c{i}'),SX['mod']))
                for k_,sty_ in (('actmod',SX['mod']),('act',SQ),('mw',SX['mw']),('actmw',SX['mw']),('pdays',SX['days']),('pport',SX['days']),('ppcit',SQ),('ppcsup',SX['text']),('ppcpp',SX['days']),('vlock',SQ),('vpool',SX['text'])):
                    cells.append(_xcell(f'{XL(k_)}{r}',ival(B,k_,ri),sty_))
            else:                                                       # today and later: live formulas on PPC_Plan_Req
                ID=q(bid); dref=f'$A{r}'
                if d==d0: bal=round(w['stock'],3); cells.append(_xcell(f'C{r}',bal,STYP))
                else:
                    bal=bal-new_vals[(prevr,'act')]+(_num(ival(B,'dlv',keyrow(B,d-pd.Timedelta(days=1)))) or 0.0)
                    cells.append(_xcell(f'C{r}',bal,SBAL,f"C{prevr}-{XL('act')}{prevr}+{XL('dlv')}{prevr}"))
                pv=S_mod.get(s_,0.0); dv=S_req.get(s_,0.0); mv=S_mw.get(s_,0.0)
                cells.append(_xcell(f"{XL('plan')}{r}",pv,SX['mod'],f'SUMIFS({T}!$H:$H,{T}!$E:$E,"{ID}",{T}!$G:$G,"Yes",{T}!$A:$A,{dref})'))
                cells.append(_xcell(f"{XL('dem')}{r}",dv,SQ,f'SUMIFS({T}!$J:$J,{T}!$E:$E,"{ID}",{T}!$G:$G,"Yes",{T}!$A:$A,{dref})'))
                for i,s2 in enumerate(slots):
                    a_=f"{XL(f'c{i}')}"
                    if s2: cells.append(_xcell(f'{a_}{r}',S_lab.get((s_,s2.lower()),0.0),SX['mod'],f'SUMIFS({T}!$H:$H,{T}!$E:$E,"{ID}",{T}!$F:$F,{a_}${WSA_HDR},{T}!$A:$A,{dref})'))
                    else: cells.append(_xcell(f'{a_}{r}',None,SX['mod']))
                    if s2: tot[f'c{i}']+=S_lab.get((s_,s2.lower()),0.0)
                cells.append(_xcell(f"{XL('actmod')}{r}",pv,SX['mod'],f"{XL('plan')}{r}"))
                cells.append(_xcell(f"{XL('act')}{r}",dv,SQ,f"{XL('dem')}{r}"))
                cells.append(_xcell(f"{XL('mw')}{r}",mv,SX['mw'],f'SUMIFS({T}!$K:$K,{T}!$E:$E,"{ID}",{T}!$G:$G,"Yes",{T}!$A:$A,{dref})'))
                cells.append(_xcell(f"{XL('actmw')}{r}",mv,SX['mw'],f"{XL('mw')}{r}"))
                cells.append(_xcell(f"{XL('pdays')}{r}",(bal/dv) if dv else 'na',SX['days'],f'IFERROR(C{r}/{XL("dem")}{r},"na")'))
                U1,X1,T1=XL('sch'),XL('rem'),XL('dlv')
                cells.append(_xcell(f"{XL('pport')}{r}",((bal+carried+cum)/dv) if dv else 'na',SX['days'],
                                    f'IFERROR((C{r}+${U1}$2+SUMPRODUCT({U1}${R0}:{U1}{r},--ISERROR(SEARCH("canc",{X1}${R0}:{X1}{r})),--({T1}${R0}:{T1}{r}="")))/{XL("dem")}{r},"na")'))
                ii=dix.get(d); av=parr[ii] if ii is not None else 0.0; cum_p+=av; P1=XL('ppcit')
                cells.append(_xcell(f'{P1}{r}',av if av else None,SQ))
                cells.append(_xcell(f"{XL('ppcsup')}{r}",(psup[ii] or None) if ii is not None else None,SX['text']))
                VP_=w.get('vp'); lk_=(VP_['lock'][ii] if (VP_ and ii is not None) else 0.0)
                if VP_:                                          # Settings ALLOWED VENDORS: column V leaves out what no short customer may use
                    VL=XL('vlock')
                    cells.append(_xcell(f"{XL('ppcpp')}{r}",((bal+cum_p-lk_)/dv) if dv else 'na',SX['days'],
                                        f'IFERROR((C{r}+SUM({P1}${t}:{P1}{r})-{VL}{r})/{XL("dem")}{r},"na")'))
                    cells.append(_xcell(f'{VL}{r}',round(lk_,3) if ii is not None else None,SQ))
                    if ii is not None:
                        cells.append(_xcell(f"{XL('vpool')}{r}",' | '.join(f"{lab_} {(p_[ii] if abs(p_[ii])>=0.5 else 0.0):,.0f}" for lab_,p_ in VP_['pools'].items()
                                                                        if VP_['who'].get(lab_) or VP_['sup'].get(lab_)),SX['text']))
                else:
                    cells.append(_xcell(f"{XL('ppcpp')}{r}",((bal+cum_p)/dv) if dv else 'na',SX['days'],
                                        f'IFERROR((C{r}+SUM({P1}${t}:{P1}{r}))/{XL("dem")}{r},"na")'))
                new_vals[(r,'act')]=dv; tot['plan']+=pv; tot['dem']+=dv; tot['mw']+=mv; tot['ppcit']+=av
            # delivery columns: as they are (the planner's constants and same-row formulas kept)
            for k_ in ('dlv','sch','po','ship','rem','book','ready','etd','bl','vessel','inv','ref','rem2'):
                v_=ival(B,k_,ri); f_=ifx(B,k_,ri); fx=None
                if isinstance(f_,str) and f_.startswith('='): fx=_carry_formula(f_,ri,colmap,r)
                sty_=SX['date'] if (isinstance(v_,(dt.datetime,dt.date)) or k_ in ('book','ready','etd')) else (SQ if k_ in ('dlv','sch') else SX['text'])
                cells.append(_xcell(f'{XL(k_)}{r}',v_,sty_,fx))
                if k_ in ('dlv','sch') and d>=d0: tot[k_]+=_num(v_) or 0.0
            sch_=_num(ival(B,'sch',ri)) or 0; rem_=str(ival(B,'rem',ri) or ''); dl_=_num(ival(B,'dlv',ri)) or 0; po_=ival(B,'po',ri); ve_=ival(B,'vessel',ri)
            st_v='' if not sch_>0 else ('Cancelled' if 'canc' in rem_.lower() else ('Delivered' if dl_>0 else ('PO Pending' if str(po_ or '').lower()=='new po' else ('Open PO' if ve_ in (None,'') else 'In transit'))))
            U1,X1,T1,V1,AC1=XL('sch'),XL('rem'),XL('dlv'),XL('po'),XL('vessel')
            cells.append(_xcell(f"{XL('status')}{r}",st_v,SX['text'],
                                f'IF({U1}{r}>0,IF(ISNUMBER(SEARCH("canc",{X1}{r})),"Cancelled",IF({T1}{r}>0,"Delivered",IF({V1}{r}="NEW PO","PO Pending",IF({AC1}{r}="","Open PO","In transit")))),"")'))
            cells=sorted((c_ for c_ in cells if c_),key=CX_order)
            drows.append(f'<row r="{r}">'+''.join(cells)+'</row>'); prevr=r
        # row 6 totals (today and later) and row 7 headers
        tc=[_xcell('A6','Total from today',SX['label'])]
        for k_ in ('plan','dem','mw','ppcit','dlv','sch')+tuple(f'c{i}' for i,s2 in enumerate(slots) if s2):
            sty_=SX['tmw'] if k_=='mw' else (SX['tmod'] if k_=='plan' or k_.startswith('c') else STOT)
            tc.append(_xcell(f'{XL(k_)}6',tot.get(k_,0.0),sty_,f'SUM({XL(k_)}{t}:{XL(k_)}{last})'))
        tc=sorted((c_ for c_ in tc if c_),key=CX_order)
        rowsx.append('<row r="6">'+''.join(tc)+'</row>')
        rowsx.append(f'<row r="{WSA_HDR}" ht="45" customHeight="1">'+''.join(_xcell(f'{L(j+1)}{WSA_HDR}',h_ if h_ else None,SX['h'+hgroup(k_)]) for j,(k_,h_,_) in enumerate(cols))+'</row>')
        rowsx+=drows
        # sheet view: same zoom / tab selection as before, frozen at the header and column C, opens at today's row
        old_sv={}
        try:
            for _,e in etree.iterparse(z.open(part[B['sheet']]),tag=M+'sheetView'):
                old_sv={k:e.get(k) for k in ('tabSelected','zoomScale','zoomScaleNormal')}; break
        except Exception as _e_: _swallow(8850,_e_); pass
        if B['block']>0: old_sv.pop('tabSelected',None)
        sva=' '.join(f'{k}="{v}"' for k,v in old_sv.items() if v)
        top=max(R0,t-3)
        TODAY_F=("TODAY()" if d0>=pd.Timestamp(dt.date.today()).normalize() else f"DATE({d0.year},{d0.month},{d0.day})")   # base-date run: the stock row is "today"
        widths=''.join(f'<col min="{j+1}" max="{j+1}" width="{wd}" customWidth="1"/>' for j,(_,_,wd) in enumerate(cols))
        xml=('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
             f'<worksheet xmlns="{M_}" xmlns:r="{RNS}"><dimension ref="A1:{LC}{last}"/>'
             f'<sheetViews><sheetView {sva} workbookViewId="0"><pane xSplit="3" ySplit="{WSA_HDR}" topLeftCell="D{top}" activePane="bottomRight" state="frozen"/>'
             f'<selection pane="topRight"/><selection pane="bottomLeft"/><selection pane="bottomRight" activeCell="D{t}" sqref="D{t}"/></sheetView></sheetViews>'
             f'<sheetFormatPr defaultRowHeight="15"/><cols>{widths}</cols><sheetData>'+''.join(rowsx)+'</sheetData>'
             f'<autoFilter ref="A{WSA_HDR}:{LC}{last}"/>'
             f'<conditionalFormatting sqref="C{R0}:C{last} {XL("pdays")}{R0}:{XL("pport")}{last} {XL("ppcpp")}{R0}:{XL("ppcpp")}{last}"><cfRule type="cellIs" dxfId="{DX_NEG}" priority="1" operator="lessThan"><formula>0</formula></cfRule></conditionalFormatting>'
             f'<conditionalFormatting sqref="A{R0}:B{last}"><cfRule type="expression" dxfId="{DX_TODAY}" priority="2"><formula>$A{R0}={TODAY_F}</formula></cfRule></conditionalFormatting>'
             '<pageMargins left="0.5" right="0.5" top="0.6" bottom="0.6" header="0.3" footer="0.3"/></worksheet>').encode('utf-8')
        plan_sheets.append((oname,xml,last,B))
        new_neg=None; b_=round(w['stock'],3)
        for d in days:
            if d<d0: continue
            if d>d0: b_=b_-S_req.get(serial(d-pd.Timedelta(days=1)),0.0)+(_num(ival(B,'dlv',keyrow(B,d-pd.Timedelta(days=1)))) or 0.0)
            if b_<0: new_neg=d; break
        report.append(dict(label=bid,sheet=oname,codes='+'.join(w['mats']),old_bal=old_bal,new_bal=w['stock'],old_d7=old_d7,
                           new_d7=sum(S_req.get(serial(d0+pd.Timedelta(days=k)),0.0) for k in range(7)),old_neg=old_neg,new_neg=new_neg))
    # ---------- 4) PPC_Plan_Req and Archive ----------
    tab.sort(key=lambda x:(x['serial'],str(x['block']),x['inbal']!='Yes',str(x['col']),str(x['line'])))
    def plan_xml(rows):
        hdr='<row r="1">'+''.join(_xcell(f'{L(j+1)}1',h_,SX['t_head']) for j,h_ in enumerate(ALIGN_COLS))+_xcell('P1',
            f"Built by update_inbound.py on {dt.date.today():%d-%b-%Y}"+(f" (base date {d0:%d-%b-%Y} = your MB52 date)" if d0<pd.Timestamp(dt.date.today()).normalize() else '')+": module plan (PPC Plan + Revised Plan) x quantity per module (PPC BOM, else MIS BOM). "
            "The planning sheets read this table with SUMIFS (customer columns, Planned modules, Demand, MW), so a change here updates them at once. "
            "Counted in Balance = No: OEM plan (made at the OEM from JW stock, not from our warehouse). Rebuilt from today on every run.",SX['info'])+'</row>'
        body=[]
        for i,x in enumerate(rows,2):
            c=[_xcell(f'A{i}',x['date'].to_pydatetime(),SX['t_date']),_xcell(f'B{i}',x['line']),_xcell(f'C{i}',x['cust']),_xcell(f'D{i}',x['mis']),
               _xcell(f'E{i}',x['block']),_xcell(f'F{i}',x['col']),_xcell(f'G{i}',x['inbal']),_xcell(f'H{i}',x['mods'],SX['t_mod']),_xcell(f'I{i}',x['qty'],SX['t_qty'])]
            if x['mods'] is not None and x['qty'] is not None: c.append(_xcell(f'J{i}',x['mods']*x['qty'],SX['t_req'],f'H{i}*I{i}'))
            else: c.append(_xcell(f'J{i}',x['req'],SX['t_req']))
            c+=[_xcell(f'K{i}',x['mw'],SX['t_mw']),_xcell(f'L{i}',x['unit']),_xcell(f'M{i}',x['codes']),_xcell(f'N{i}',x['src'])]
            body.append(f'<row r="{i}">'+''.join(c)+'</row>')
        n=len(rows)+1; widths=[11,16,26,24,26,26,10,10,13,13,9,7,30,16,3,90]
        return (n,('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                f'<worksheet xmlns="{M_}" xmlns:r="{RNS}"><dimension ref="A1:P{n}"/><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
                '<selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/>'
                '<cols>'+''.join(f'<col min="{j}" max="{j}" width="{w_}" customWidth="1"/>' for j,w_ in enumerate(widths,1))+'</cols><sheetData>'
                +hdr+''.join(body)+f'</sheetData><autoFilter ref="A1:N{n}"/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>').encode('utf-8'))
    arch_add=[r for r in arch_new if akey(r) not in arch_have]
    arch=sorted(old_arch+arch_add,key=lambda r:(pd.Timestamp(r[1]),str(r[0])))
    def arch_xml(rows):
        hdr='<row r="1">'+''.join(_xcell(f'{L(j+1)}1',h_,SX['t_head']) for j,h_ in enumerate(ARCH_COLS))+_xcell('Q1',
            'Delivery records of the planning-sheet rows deleted by update_inbound.py (rows before the 1st of the month). Kept here so nothing is lost; not read by the script.',SX['info'])+'</row>'
        body=[]
        for i,r in enumerate(rows,2):
            c=[]
            for j,v in enumerate(r):
                sty_=SX['t_date'] if isinstance(v,(dt.datetime,dt.date)) else (SX['t_mod'] if j in (2,3) else 0)
                c.append(_xcell(f'{L(j+1)}{i}',v,sty_))
            body.append(f'<row r="{i}">'+''.join(c)+'</row>')
        n=len(rows)+1; widths=[26,12,11,12,13,16,22,12,11,11,18,26,13,13,13,3,90]
        return (n,('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                f'<worksheet xmlns="{M_}" xmlns:r="{RNS}"><dimension ref="A1:Q{max(n,1)}"/><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
                '<selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/>'
                '<cols>'+''.join(f'<col min="{j}" max="{j}" width="{w_}" customWidth="1"/>' for j,w_ in enumerate(widths,1))+'</cols><sheetData>'
                +hdr+''.join(body)+f'</sheetData><autoFilter ref="A1:O{n}"/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>').encode('utf-8'))
    # ---------- 5) workbook: sheet list, parts, names ----------
    old_order=[s.get('name') for s in shs]; patched={}; drop=set(['xl/calcChain.xml'])
    ct=etree.fromstring(z.read('[Content_Types].xml'))
    gone=[]
    for nm_ in WSA_DROP_SHEETS:                    # Stock / Total_Plan_Req: not needed in the WSA daily file (user, Nov 2026)
        if nm_ not in sheet_el or nm_ not in part or sheet_el[nm_].getparent() is None: continue
        el_=sheet_el[nm_]; rid_=el_.get('{%s}id'%RNS); shs.remove(el_)
        for r_ in list(rels):
            if r_.get('Id')==rid_: rels.remove(r_)
        pth_=part[nm_]; drop.add(pth_); rp_=posixpath.join(posixpath.dirname(pth_),'_rels',posixpath.basename(pth_)+'.rels')
        if rp_ in names_:
            for rel in etree.fromstring(z.read(rp_)):
                if rel.get('TargetMode')!='External': drop.add(posixpath.normpath(posixpath.join(posixpath.dirname(pth_),rel.get('Target'))))
            drop.add(rp_)
        gone.append(nm_)
    def free_part():
        k_=1
        while f'xl/worksheets/sheet{k_}.xml' in names_ or f'xl/worksheets/sheet{k_}.xml' in patched: k_+=1
        return f'xl/worksheets/sheet{k_}.xml'
    def add_sheet(name, after_el=None):
        pth=free_part(); ids_=set(r.get('Id') for r in rels); n_=1
        while f'rId{n_}' in ids_: n_+=1
        etree.SubElement(rels,'{%s}Relationship'%PNS,Id=f'rId{n_}',Type=RNS+'/worksheet',Target=pth.replace('xl/','',1))
        e_=etree.Element(M+'sheet',name=name,sheetId=str(max(int(s.get('sheetId')) for s in shs)+1)); e_.set('{%s}id'%RNS,f'rId{n_}')
        if after_el is not None: after_el.addnext(e_)
        else: shs.append(e_)
        etree.SubElement(ct,'{%s}Override'%CTN,PartName='/'+pth,ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml')
        part[name]=pth; patched[pth]=b''; return e_
    redesigned=set()
    for oname,xml,last_,B in plan_sheets:
        if oname in part and oname==B['sheet']: patched[part[oname]]=xml; redesigned.add(oname)
        elif oname in part: patched[part[oname]]=xml; redesigned.add(oname)         # an earlier run's split sheet (input already redesigned)
        else:
            prev_el=sheet_el[B['sheet']]
            while prev_el.getnext() is not None and prev_el.getnext().get('name') in [p[0] for p in plan_sheets if p[3]['sheet']==B['sheet']]: prev_el=prev_el.getnext()
            add_sheet(oname,prev_el); patched[part[oname]]=xml; redesigned.add(oname)
    # replaced sheets lose their old comments / drawings / print settings (all on deleted rows)
    for oname in redesigned:
        rp=posixpath.join(posixpath.dirname(part[oname]),'_rels',posixpath.basename(part[oname])+'.rels')
        if rp in names_:
            for rel in etree.fromstring(z.read(rp)):
                if rel.get('TargetMode')=='External': continue
                tp=posixpath.normpath(posixpath.join(posixpath.dirname(part[oname]),rel.get('Target')))
                drop.add(tp)
            drop.add(rp)
    n_plan,pxml=plan_xml(tab)
    if T not in part: add_sheet(T)
    patched[part[T]]=pxml
    n_arch,axml=arch_xml(arch)
    if ARCHIVE_SHEET not in part: add_sheet(ARCHIVE_SHEET)
    patched[part[ARCHIVE_SHEET]]=axml
    for o in list(ct):
        pn=str(o.get('PartName') or '').lstrip('/')
        if pn in drop: ct.remove(o)
    new_order=[s.get('name') for s in shs]; idx_new={n:i for i,n in enumerate(new_order)}
    imap={i:idx_new[n] for i,n in enumerate(old_order) if n in idx_new}
    dns=wbx.find(M+'definedNames')
    if dns is None:
        dns=etree.Element(M+'definedNames'); cp_=wbx.find(M+'calcPr')
        if cp_ is not None: cp_.addprevious(dns)
        else: wbx.append(dns)
    red_old={i for i,n in enumerate(old_order) if n in redesigned or n in gone}
    for e in list(dns):
        ls=e.get('localSheetId')
        if ls is None: continue
        if int(ls) in red_old or e.get('name')=='_xlnm._FilterDatabase' and new_order[imap.get(int(ls),0)] in (T,ARCHIVE_SHEET): dns.remove(e); continue
        if int(ls) in imap: e.set('localSheetId',str(imap[int(ls)]))
    def qn(n): return "'"+n.replace("'","''")+"'"
    for oname,xml,last_,B in plan_sheets:
        etree.SubElement(dns,M+'definedName',name='_xlnm._FilterDatabase',localSheetId=str(idx_new[oname]),hidden='1').text=f"{qn(oname)}!$A${WSA_HDR}:${LC}${last_}"
    etree.SubElement(dns,M+'definedName',name='_xlnm._FilterDatabase',localSheetId=str(idx_new[T]),hidden='1').text=f"{T}!$A$1:$N${n_plan}"
    etree.SubElement(dns,M+'definedName',name='_xlnm._FilterDatabase',localSheetId=str(idx_new[ARCHIVE_SHEET]),hidden='1').text=f"{ARCHIVE_SHEET}!$A$1:$O${n_arch}"
    bv=wbx.find(f'{M}bookViews/{M}workbookView')
    if bv is not None:
        for a_ in ('activeTab','firstSheet'):
            if bv.get(a_) is not None: bv.set(a_,str(imap.get(int(bv.get(a_)),0)))
    cp_=wbx.find(M+'calcPr')
    if cp_ is None: cp_=etree.SubElement(wbx,M+'calcPr')
    cp_.set('fullCalcOnLoad','1')
    for r in list(rels):
        if str(r.get('Target') or '').endswith('calcChain.xml'): rels.remove(r)
    # ---------- 6) Dashboard: one lookup column for every planning sheet ----------
    dash_note=''
    if 'Dashboard' in part and part['Dashboard'] not in patched:
        try:
            droot=etree.fromstring(z.read(part['Dashboard'])); sd=droot.find(M+'sheetData'); masters={}
            for c in sd.iter(M+'c'):
                f=c.find(M+'f')
                if f is not None and f.get('t')=='shared' and f.get('ref') is not None: masters[f.get('si')]=(c.get('r'),f.text or '')
            for c in sd.iter(M+'c'):
                f=c.find(M+'f')
                if f is None or f.get('t')!='shared': continue
                mr,mt=masters.get(f.get('si'),(None,None))
                if mr is None: continue
                f.text=mt if c.get('r')==mr else Translator('='+mt,origin=mr).translate_formula(c.get('r'))[1:]
                for a_ in ('t','ref','si'):
                    if a_ in f.attrib: del f.attrib[a_]
            info_={}
            for w in work:
                B=w['B']
                if B.get('dash') is False: continue        # back glass of a split sheet: no Dashboard row of its own
                info_.setdefault(B['sheet'],[]).append((B['c0'],out_name[(B['sheet'],B['block'])]))
            for oname,xml,last_,B in plan_sheets: info_.setdefault(oname,[(1,oname)])
            sst_=[]; dash_fixed=[]
            if 'xl/sharedStrings.xml' in names_:
                sst_=[''.join(si.itertext()) for si in etree.fromstring(z.read('xl/sharedStrings.xml')).findall(M+'si')]
            pat=re.compile(r"('(?:[^']|'')+'|[A-Za-z0-9_.]+)!\$A:\$[A-Z]{1,3}")
            nfix=0; PP=CX['pport']
            for rowel in sd.iter(M+'row'):
                rr=int(rowel.get('r'))
                if rr<2: continue
                cells=rowel.findall(M+'c'); fx=[c.find(M+'f') for c in cells if c.find(M+'f') is not None and c.find(M+'f').text]
                hit={m.group(1).strip("'").replace("''","'") for f in fx for m in pat.finditer(f.text)}
                hit=[h_ for h_ in hit if h_ in info_]
                if not hit: continue
                acell=next((c for c in cells if c.get('r')==f'A{rr}'),None); av=None
                if acell is not None and acell.find(M+'v') is not None:
                    try: av=float(acell.find(M+'v').text)
                    except (TypeError,ValueError): av=None
                tgt={}
                for h_ in hit:                                   # the block the old column index pointed into (Potting B = column AW ...)
                    blks=sorted(info_[h_]); tgt[h_]=blks[0][1]
                    for c0_,nm_ in blks:
                        if av is not None and av>=c0_: tgt[h_]=nm_
                if len(set(tgt.values()))>1:                     # one row reading two sheets (EPE 420 GSM row read EVA 420 GSM in column D): use the sheet named in column B
                    bcell=next((c for c in cells if c.get('r')==f'B{rr}'),None); blab=''
                    if bcell is not None:
                        if bcell.get('t')=='s' and bcell.find(M+'v') is not None:
                            try: blab=sst_[int(bcell.find(M+'v').text)]
                            except (ValueError,IndexError): blab=''
                        elif bcell.get('t')=='inlineStr': blab=''.join(bcell.itertext())
                        elif bcell.find(M+'v') is not None: blab=bcell.find(M+'v').text or ''
                    own=next((v_ for v_ in tgt.values() if _norm(v_)==_norm(blab)),None)
                    if own:
                        for h_ in list(tgt):
                            if tgt[h_]!=own: dash_fixed.append(f"row {rr} '{blab}' read '{tgt[h_]}' -> '{own}'"); tgt[h_]=own
                def _sub(m):
                    h_=m.group(1).strip("'").replace("''","'")
                    return (qn(tgt[h_])+f'!$A:${LC}') if h_ in tgt else m.group(0)
                for f in fx: f.text=pat.sub(_sub,f.text)
                if acell is not None:
                    for ch in list(acell): acell.remove(ch)
                    if 't' in acell.attrib: del acell.attrib['t']
                    etree.SubElement(acell,M+'v').text=str(PP)
                nfix+=1
            # back-glass rows (user, Nov 2026: 'Dashboard rows for the back-glass sheets?' - 'Yes'): one row per '<x> Back' sheet that no
            # row reads yet = a copy of its front row (same formulas / format, reading the Back sheet), added below the table; a rerun on
            # the aligned file finds the rows and adds nothing
            import copy as _cp
            rows_=sorted(sd.findall(M+'row'),key=lambda e:int(e.get('r'))); read_={}; last_=1
            for rowel in rows_:
                rr=int(rowel.get('r')); used=False
                for c in rowel.findall(M+'c'):
                    f=c.find(M+'f')
                    if f is not None and f.text:
                        used=True
                        for m in pat.finditer(f.text): read_.setdefault(m.group(1).strip("'").replace("''","'"),rowel)
                    elif c.find(M+'v') is not None or c.find(M+'is') is not None: used=True
                if used and rr>1: last_=max(last_,rr)
            old_last=last_; back_add=[]
            for fname,rowel in sorted(read_.items(),key=lambda kv:int(kv[1].get('r'))):
                if not fname.endswith(' Front'): continue
                bk=[n_ for n_ in part if n_.endswith(' Back') and n_[:-5].startswith(fname[:-6])]
                if len(bk)!=1 or bk[0] in read_: continue
                bname=bk[0]; src_r=int(rowel.get('r')); last_+=1; nr=last_
                new=_cp.deepcopy(rowel); new.set('r',str(nr))
                for c in new.findall(M+'c'):
                    col_=re.match(r'[A-Z]+',c.get('r')).group(0); old_ref=c.get('r'); c.set('r',f'{col_}{nr}')
                    f=c.find(M+'f')
                    if f is not None and f.text:
                        for a_ in ('t','ref','si'):
                            if a_ in f.attrib: del f.attrib[a_]
                        t_=Translator('='+f.text,origin=old_ref).translate_formula(c.get('r'))[1:]
                        f.text=re.sub(r"('(?:[^']|'')+'|[A-Za-z0-9_.]+)!",lambda m:(qn(bname)+'!') if m.group(1).strip("'").replace("''","'")==fname else m.group(0),t_)
                        v_=c.find(M+'v')
                        if v_ is not None: c.remove(v_)                 # Excel calculates on open (fullCalcOnLoad)
                        if c.get('t') in ('str','e','n','b'): del c.attrib['t']
                    elif col_=='B':                                     # label = the old label + ' Back'
                        lab=''
                        if c.get('t')=='s' and c.find(M+'v') is not None:
                            try: lab=sst_[int(c.find(M+'v').text)]
                            except (ValueError,IndexError): lab=''
                        elif c.get('t')=='inlineStr': lab=''.join(c.itertext())
                        for ch in list(c): c.remove(ch)
                        c.set('t','inlineStr'); is_=etree.SubElement(c,M+'is'); etree.SubElement(is_,M+'t').text=(lab or fname[:-6])+' Back'
                pos_=next((i for i,e in enumerate(list(sd)) if int(e.get('r'))>=nr),None)
                if pos_ is not None and int(list(sd)[pos_].get('r'))==nr: sd.remove(list(sd)[pos_])   # an empty formatted row there
                pos_=next((i for i,e in enumerate(list(sd)) if int(e.get('r'))>nr),None)
                if pos_ is None: sd.append(new)
                else: sd.insert(pos_,new)
                back_add.append(f"row {nr} '{bname}' (copy of row {src_r})")
            if back_add:                                                # filter / conditional formats / size: cover the new rows
                def _ext(ref):
                    return ' '.join(re.sub(rf'^([A-Z]+\$?)(\d+):([A-Z]+\$?){old_last}$',lambda m:f"{m.group(1)}{m.group(2)}:{m.group(3)}{last_}",x_) for x_ in ref.split())
                for e in droot.iter(M+'autoFilter',M+'conditionalFormatting'):
                    k_='ref' if e.tag==M+'autoFilter' else 'sqref'
                    if e.get(k_): e.set(k_,_ext(e.get(k_)))
                for e in (wbx.find(M+'definedNames') if wbx.find(M+'definedNames') is not None else []):   # the hidden filter name follows the filter
                    if e.get('name')=='_xlnm._FilterDatabase' and re.match(r"'?Dashboard'?!",e.text or ''):
                        e.text=re.sub(rf'(\$[A-Z]+\$){old_last}$',lambda m:f"{m.group(1)}{last_}",e.text)
                dm=droot.find(M+'dimension')
                if dm is not None and dm.get('ref'):
                    m=re.match(r'([A-Z]+\d+):([A-Z]+)(\d+)$',dm.get('ref'))
                    if m and int(m.group(3))<last_: dm.set('ref',f"{m.group(1)}:{m.group(2)}{last_}")
            patched[part['Dashboard']]=etree.tostring(droot,xml_declaration=True,encoding='UTF-8',standalone=True)
            dash_note=f'Dashboard: {nfix} rows now read column {PP} (Plant + Port Days) of the new layout'+(('; fixed: '+'; '.join(dash_fixed)) if dash_fixed else '')+(('; back-glass rows added: '+'; '.join(back_add)) if back_add else '')
        except Exception as ex_: dash_note=f'NOTE: Dashboard not updated ({type(ex_).__name__}: {ex_})'
    if GS['renames'] or gone:
        own_=set(part[o_] for o_ in redesigned)|{part[T],part[ARCHIVE_SHEET]}; nref=0
        for nm_ in z.namelist():
            if not nm_.startswith(('xl/worksheets/sheet','xl/charts/chart')) or nm_ in drop or nm_ in own_: continue
            raw_=patched.get(nm_) or z.read(nm_)
            t_=raw_.decode('utf-8')
            if not any(k_ in t_ for k_ in list(GS['renames'])+gone): continue
            t2_=_ref_fix(t_,GS['renames'],gone)
            if t2_!=t_: patched[nm_]=t2_.encode('utf-8'); nref+=1
        for e in dns:
            if e.text: e.text=_ref_fix(e.text,GS['renames'],gone)
    patched['xl/styles.xml']=STY.bytes()
    patched['xl/workbook.xml']=etree.tostring(wbx,xml_declaration=True,encoding='UTF-8',standalone=True)
    patched['xl/_rels/workbook.xml.rels']=etree.tostring(rels,xml_declaration=True,encoding='UTF-8',standalone=True)
    patched['[Content_Types].xml']=etree.tostring(ct,xml_declaration=True,encoding='UTF-8',standalone=True)
    tmp=out+'.tmp'
    with zipfile.ZipFile(tmp,'w',zipfile.ZIP_DEFLATED) as zo:
        seen_=set()
        for info in z.infolist():
            if info.filename in drop: continue
            data=patched.get(info.filename); seen_.add(info.filename)
            zo.writestr(info,data if data is not None else z.read(info.filename),compress_type=zipfile.ZIP_DEFLATED)
        for nm_,data in patched.items():
            if nm_ not in seen_: zo.writestr(nm_,data,compress_type=zipfile.ZIP_DEFLATED)
    z.close(); wv.close(); wf.close()
    try:                                                              # never replace anything with a broken file: every part must read back
        with zipfile.ZipFile(tmp) as zc:
            bad_=zc.testzip()
            if bad_: raise ValueError(f'damaged part {bad_}')
            for nm_ in zc.namelist():
                if nm_.endswith(('.xml','.rels')): etree.fromstring(zc.read(nm_))
    except Exception as ex_:
        try: os.remove(tmp)
        except OSError: pass
        raise RuntimeError(f'the aligned file did not pass its own check ({type(ex_).__name__}: {ex_}) - nothing was replaced')
    try:
        ran_mt_=os.path.getmtime(nf_) if os.path.exists(nf_) else None   # the file this run read: second files newer than it are kept (_wsa_tidy)
        WSA_BACKUP[0]=_wsa_backup(out,nf_,d0)                        # no backup folder (user, Nov 2026); OneDrive version history keeps old versions
        out=_wsa_save(tmp,out,stem,d0)                                # OneDrive / Excel lock: retry, write in place, else '<stem> (aligned <date>).xlsx'
    except Exception as ex_:                                          # every way failed: your file stays as it is, the run goes on
        try: os.remove(tmp)
        except OSError: pass
        raise RuntimeError(f"could not save {os.path.basename(out)} ({type(ex_).__name__}: {ex_}) - is it open in Excel? Close it and run again; "
                           "your WSA file was NOT changed") from ex_
    WSA_MOVED[:]=_wsa_tidy(out,own_stem,nf_)           # 2_WSA files of earlier dates: deleted
    # ---------- console ----------
    oct_open=[r for r in arch_add if r[-1]=='Not delivered' and pd.Timestamp(r[1])>=keep_from-pd.Timedelta(days=31)]
    print(f"WSA aligned: {os.path.basename(out)} - {len(plan_sheets)} planning sheets in one layout (header row {WSA_HDR}, columns A to {LC}), rows {keep_from:%d-%b-%y} to {end_d:%d-%b-%y}")
    print(f"  Customer columns (same on every sheet): {', '.join(s_ for s_ in slots if s_)}")
    if PPC.get('src'):
        rows_=PPC.get('rows') or []; late_=sum(1 for x in rows_ if x['late']); notm_=sorted({x['mat'] for x in rows_ if not x['in_mrp']})
        print(f"  PPC In Transit ({PPC['src']}, sheet {PPC.get('sheet') or 'In_Transit'}): {len(rows_)} shipment lines with a pending quantity ({late_} late - counted on today)"
              +(f"; {len(PPC['skipped'])} lines not mapped to a BOM material: "+'; '.join(PPC['skipped'][:4]) if PPC.get('skipped') else '')
              +(f"; codes not in the MRP: {', '.join(notm_)}" if notm_ else ''))
        if PPC.get('fixes'):
            print(f"  PPC In Transit: {len(PPC['fixes'])} lines traced to a different BOM code than their Final RM Code (your WSA PO + BL#, description, supplier, RM Code), e.g. "
                  +'; '.join(f"{x['final_code'] or '-'} -> {x['code']} {x['qty']:,.0f}" for x in PPC['fixes'][:3])+" - every line is on the Inbound sheet 'PPC In Transit'")
        chk_=[x for x in (PPC.get('lines') or []) if 'CHECK' in str(x.get('how'))]
        if chk_: print(f"  PPC In Transit: {len(chk_)} lines to CHECK - your WSA sheet has the shipment as another size / HS than the description says: "
                       +'; '.join(f"{x['description'][:40]} {x['qty']:,.0f} on {', '.join(x['wsa_sheets'])}" for x in chk_[:3]))
        for w_ in PPC.get('why') or []: print('  NOTE (PPC In Transit): '+w_)
    else: print("  PPC In Transit: NOT READ - "+'; '.join(PPC.get('why') or ['no PPC file with an In_Transit sheet'])+" - the PPC columns stay empty")
    print(f"  From today: live formulas on the sheet {T} ({len(tab):,} rows = module plan x BOM); Balance = stock typed on today's row (yellow), then Balance - Actual + Delivered")
    print("  Balance on today's row = stock 1_MRP uses (MB52, PPC count rolled forward)"+('' if WSA_OEM_IN_BALANCE else ', JW stock at the OEM left out')+
          "; Demand = module plan x BOM"+('' if WSA_OEM_IN_BALANCE else ' without the OEM plan (OEM columns still show their modules)'))
    print(f"  Rows before {keep_from:%d-%b-%y} deleted; {len(arch_add)} delivery records from them moved to the sheet {ARCHIVE_SHEET}"
          +(f" ({len(oct_open)} of {keep_from-pd.Timedelta(days=1):%b-%y} with no Delivered quantity and not cancelled)" if oct_open else ''))
    splits=[f"{B['sheet']} -> {o}" for o,_,_,B in plan_sheets if o!=B['sheet'] and not B.get('src_label')]
    for l_ in GS['lines']:
        if l_.startswith('WARNING'): print('  '+l_)
    if [l_ for l_ in GS['lines'] if not l_.startswith('WARNING')]:
        print('  Front and back glass on their own sheets (one Balance per glass): '+'; '.join(l_ for l_ in GS['lines'] if not l_.startswith('WARNING')))
        print(f"    delivery rows split by the In_Transit lines of their BL#: {GS['n_bl']}; by 1 front : 1 back (no BL# line found): {GS['n_11']}")
    for m_,S_ in SHR.items():
        print(f"  {m_} on {len(S_['sheets'])} sheets - each customer on its own sheet, stock and PPC arrivals by supplier (Settings ALLOWED VENDORS, as 1_MRP): "
              +' | '.join(f"'{_out_sheet_name(blocks[j])}' stock "+(', '.join(f'{n} {q:,.0f}' for n,q in sorted(S_['by_v'][j].items(),key=lambda t:-t[1]) if abs(q)>0.5) or '0')
                          +f", PPC arrivals {sum(S_['arr'][j]):,.0f} ("+'; '.join(f"{', '.join(v)} -> {k or 'any vendor'}" for k,v in sorted(S_['cust'][j].items(),key=lambda t:(t[0]=='',t[0])))+')'
                          for j in S_['sheets']))
    if gone: print('  Sheets left out of the WSA daily file (not needed): '+', '.join(gone))
    if splits: print('  Two materials on one sheet -> second material on its own sheet: '+'; '.join(splits))
    if dash_note: print('  '+dash_note)
    w_=max(len(x['label']) for x in report) if report else 10
    fmt=lambda v: '-' if v is None else f'{v:,.0f}'
    print(f"  {'Block':<{w_}}  {'Balance today old -> new':>28}  {'Demand next 7 d old -> new':>30}  {'Balance < 0 from old -> new':>28}")
    for x in report:
        print(f"  {x['label']:<{w_}}  {fmt(x['old_bal']):>13} -> {fmt(x['new_bal']):<11}  {fmt(x['old_d7']):>14} -> {fmt(x['new_d7']):<12}  "
              f"{(x['old_neg'].strftime('%d-%b-%y') if x['old_neg'] is not None else 'never'):>12} -> {(x['new_neg'].strftime('%d-%b-%y') if x['new_neg'] is not None else 'never')}")
    if skipped: print('  Not rebuilt: '+'; '.join(skipped))
    if unmatched: print("  NOTE: plan customers not on the WSA 'Lookup' sheet (own column, named after the MIS customer): "+'; '.join(sorted(unmatched)))
    if misonly: print('  NOTE: MIS demand of customers not in the module plan (in Demand, no modules): '+'; '.join(sorted(misonly)))
    if other: print(f"  NOTE: demand not split by customer (row '(not split by customer)' on {T}): "+', '.join(f'{k} ({v} d)' for k,v in other.items()))
    if dropped: print('  NOTE: columns not in the standard layout (left out): '+'; '.join(sorted(dropped)))
    if unit_notes: print('  Labels fixed (your text did not fit the material; the numbers are in the SAP / MRP unit): '+'; '.join(unit_notes))
    if VP_WSA: print(f"  Settings ALLOWED VENDORS on {len(VP_WSA)} planning sheets (columns {XL('vlock')} / {XL('vpool')}; column V subtracts {XL('vlock')}): "+vp_console(VP_WSA,d0))
    return out,report

# ===================== Material Plan sheet: one column per day =====================

