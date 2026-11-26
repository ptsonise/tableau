# ---- PPC In_Transit -> YOUR BOM material: supplier + description + RM Code / Final RM Code (Final RM Code is sometimes wrong) ----
_KIND_KW=[('CELL FIXATION','TAPE'),('FIXATION TAPE','TAPE'),('PERFORATED TAPE','TAPE'),('TAPE','TAPE'),('EPE','EPE'),('EVA','EVA'),('POE','POE'),
          ('BUSBAR','BUSBAR'),('BUS BAR','BUSBAR'),('ROUND WIRE','RIBBON'),('RIBBON HOLDER','OTHER'),('RIBBON','RIBBON'),('GLASS','GLASS'),('FRAME','FRAME'),
          ('JUNCTION BOX','JB'),('J-BOX','JB'),('CABLE','JB'),('SEALANT','SEALANT'),('RTV','SEALANT'),('POTTING','POTTING'),('LABEL','LABEL'),('BARCODE','LABEL'),
          ('PALLET','PALLET'),('INNERBOX','BOX'),('BOX','BOX'),('N-TYPE','CELL'),('SOLAR CELL','CELL')]
_SUP_PLACES={'SHANGHAI','VIETNAM','CHINA','HONG','KONG','ZHEJIANG','JIANGSU','NINGBO','HANGZHOU','TAICANG','NINGHAI','SHANDONG','THAILAND','INDIA',
             'USA','HK','SUZHOU','CHANGZHOU','WUXI','GUANGDONG','SHENZHEN','JIAXING'}
_SUP_STOP={'CO','LTD','LIMITED','INC','LLC','COMPANY','COMPA','CORP','CORPORATION','GROUP','THE','AND','OF','PVT','PRIVATE','TRADE','TRADING','INTERNATIONAL',
           'TECHNOLOGY','MATERIAL','MATERIALS','NEW','INDUSTRY','INDUSTRIAL','INVESTMENT','SCIENCE','APPLIED','ENERGY','SOLAR','PV','POWER','ELECTRIC',
           'ELECTRONIC','ELECTRONICS','MANUFACTURING','IMPORT','EXPORT','HOLDINGS','SYSTEMS','DATA','FILM'}
SUPPLIER_ALIAS={'KB':'KIBING'}      # short supplier names on your BOM sheet -> the name on the shipping documents
def _kind_of(t):
    """Product type of a description = the FIRST product word in it ('EPE FILM - EPE (EVA, ...)' is EPE, not EVA)."""
    u=' '.join(str(t or '').upper().split()); best=None
    for kw,k in _KIND_KW:
        i=u.find(kw)
        if i>=0 and (best is None or i<best[0] or (i==best[0] and len(kw)>best[2])): best=(i,k,len(kw))
    return best[1] if best else None
def _sup_tokens(s):
    """Distinctive words of a supplier name (no place names, no 'CO LTD ...'): 'TAICANG JUREN INTERNATIONAL TRADE CO., LTD' -> {'JUREN'}."""
    return {SUPPLIER_ALIAS.get(w,w) for w in re.split(r'[^A-Z0-9]+',str(s or '').upper()) if len(w)>1 and w not in _SUP_STOP and w not in _SUP_PLACES}
def _code_family(c):
    c=str(c).upper()
    for g in (('RMBUS','RMRIB'),('RMEVA','RMEPE'),('RMTAP','RMTTT'),('RMGLS',),('RMALF',),('RMRTV',),('RMJUB',),('RMLAB',),('PMPK',),('CST',),('RMFC',)):
        if c.startswith(g): return g
    return (c[:5],)
def in_transit_catalog(ppc_path, mis_path, DS=None):
    """The materials an In_Transit line can be mapped to: YOUR PPC BOM (code, material description, supplier) first, then the
       other MRP materials (description from MIS Common Name / MB52)."""
    cat={}
    try:
        for r in wb_ro(ppc_path)['BOM'].iter_rows(min_row=2,values_only=True):
            if r and len(r)>6 and r[4]:
                c=str(r[4]).strip().upper(); e=cat.setdefault(c,dict(desc=str(r[5] or '').strip(),sup=set(),src='PPC BOM'))
                if r[6]: e['sup'].add(' '.join(str(r[6]).split()))
    except Exception as _e_: _swallow(7775,_e_); pass
    if mis_path:
        try:
            w=wb_ro(mis_path)
            for r in w['Common Name'].iter_rows(min_row=2,values_only=True):
                if r and r[0] and r[1] and str(r[0]).strip().upper().startswith(('RM','PM','CS')): cat.setdefault(str(r[0]).strip().upper(),dict(desc=str(r[1]),sup=set(),src='MIS'))
            for r in w['MB52'].iter_rows(min_row=2,values_only=True):
                if r and len(r)>5 and r[2] and r[5] and str(r[2]).strip().upper().startswith(('RM','PM','CS')): cat.setdefault(str(r[2]).strip().upper(),dict(desc=str(r[5]),sup=set(),src='MB52'))
        except Exception as _e_: _swallow(7783,_e_); pass
    for b in (DS or {}).get('blocks') or []:
        cat.setdefault(str(b['mat']).upper(),dict(desc=str(b.get('common') or b['mat']),sup=set(),src='MRP'))
    return cat
def _po_set(v):
    """PO numbers in a cell (SAP numbers of 8+ digits; one cell can hold several)."""
    return set(re.findall(r'\d{8,}',str(v).replace('﻿',''))) if v not in (None,'') else set()
def _bl_set(v):
    """BL numbers in a cell (HBL / MBL / BL#; one cell can hold several): words of 6+ letters / digits with 4+ digits."""
    if v in (None,''): return set()
    if isinstance(v,float) and v.is_integer(): v=int(v)
    return {t for t in re.split(r'[^A-Z0-9]+',str(v).upper().replace('﻿','')) if len(t)>=6 and sum(ch.isdigit() for ch in t)>=4}
def wsa_ref_index(groups):
    """PO and BL# of your WSA planning sheets (every dated row, and the Archive sheet of an aligned file) -> lookup tables."""
    sheets=[]; by_po={}; by_bl={}; by_tail={}
    for g in groups or []:
        k=len(sheets); sheets.append(dict(label=g['label'],sheet=g['sheet'],codes=set(g['codes'])))
        for po,bl in g.get('refs') or []:
            ps=_po_set(po); bs=_bl_set(bl)
            for p in ps: by_po.setdefault(p,set()).add(k)
            for b in bs:
                by_bl.setdefault(b,[]).append((k,ps))
                if len(b)>=8: by_tail.setdefault(b[-8:],[]).append((k,ps,b))
    return dict(sheets=sheets,po=by_po,bl=by_bl,tail=by_tail)
def wsa_find(idx, po, bls):
    """The WSA sheets that have this shipment: rows with the PO AND one of its BL numbers (HBL / MBL; 'CMDUCHN4068035' = 'CHN4068035'),
       else rows with one of its BL numbers, else rows with the PO. Returns (how matched, [sheet numbers])."""
    if not idx: return None,[]
    ps=_po_set(po); pb=set(); b_=set()
    for b in bls:
        hits=list(idx['bl'].get(b,[]))
        if len(b)>=8: hits+=[(k,p_) for k,p_,b2 in idx['tail'].get(b[-8:],[]) if b2!=b and (b2.endswith(b) or b.endswith(b2))]
        for k,p_ in hits:
            b_.add(k)
            if ps&p_: pb.add(k)
    if pb: return 'PO + BL#',sorted(pb)
    if b_: return 'BL#',sorted(b_)
    p_=set().union(*(idx['po'].get(p,set()) for p in ps)) if ps else set()
    return ('PO',sorted(p_)) if p_ else (None,[])

def map_in_transit(lines, cat, wsa=None):
    """Each line -> the best material of the catalog.
       1) YOUR WSA sheets first (user rule, Nov 2026): the sheet(s) where you typed this shipment's PO + BL# (HBLNo. / MBL; else the
          BL#, else the PO) - the material of that sheet that fits the description (same product type, no contradicting size /
          thickness / HS / GSM / front-back) is taken; several fit -> the best description match. A WSA sheet whose material
          contradicts the description is reported ('check') and not used.
       2) Else the description (sizes, thickness, HS, GSM, front / back, product type - a contradicting size counts against),
          the supplier (your BOM supplier, or the codes this supplier ships on the same sheet), RM Code (shipped) > Final RM Code
          (PO) > Match / PO-wise codes. The Final RM Code is replaced only when another code is clearly better (2+ points); a line
          with no code at all needs a strong, specific description match, else it stays unmapped.
       0) Before both (user, 21-Nov-2026 'go', Kibing KB-WSA-260716B-5 / -6): 2+ lines of ONE invoice with the SAME description but
          DIFFERENT unit prices, each price with its own Final RM Code (PPC sets it from 'PO and Unit Price Key') = the description
          was copied, so neither the description nor your WSA sheet can tell them apart -> the Final RM Code (unit price) is used."""
    CODE_=re.compile(r'^(RM|PM|CS)[A-Z0-9\-]{6,}$')
    al=(wsa or {}).get('alias') or {}; A=lambda c: al.get(c,c); idx=(wsa or {}).get('idx')
    grp_={}
    for x in lines:
        if x.get('invoice') not in (None,'') and x.get('price') and x['description']:
            grp_.setdefault((str(x['invoice']).strip().upper(),x['description'].upper()),[]).append(x)
    copied={}
    for (inv_,_d),xs in grp_.items():
        pc={}
        for x in xs: pc.setdefault(round(x['price'],4),set()).add(x['final_code'])
        fcs=[next(iter(v)) for v in pc.values() if len(v)==1]
        if len(pc)>=2 and len(fcs)==len(pc) and len(set(fcs))==len(pc) and all(c in cat for c in fcs):
            txt=' / '.join(f'{p_:g} = {next(iter(v))}' for p_,v in sorted(pc.items()))
            for x in xs: copied[id(x)]=(inv_,txt)
            if inv_ not in IT_COPIED:
                IT_COPIED[inv_]=txt
                print(f"NOTE: PPC In_Transit invoice {inv_}: {len(xs)} lines with the same description ({xs[0]['description'][:45]}) at different"
                      f" unit prices - description looks copied, material by unit price / Final RM Code: {txt}",flush=True)
    sheet_sup={}
    for x in lines:
        for c in (x['rm_code'],x['final_code']):
            if CODE_.match(c) and c in cat: sheet_sup.setdefault(c,set()).update(_sup_tokens(x['supplier']))
    kd_cache={}
    def keys(t):
        if t not in kd_cache: kd_cache[t]=(_desc_keys(t),_kind_of(t))
        return kd_cache[t]
    for x in lines:
        desc=x['description']; st=_sup_tokens(x['supplier']); ka,k1=keys(desc)
        ev={}
        for c,w in ((x['rm_code'],1.5),(x['final_code'],1.0),(x['match'],0.5)):
            if CODE_.match(c) and c in cat: ev[c]=ev.get(c,0)+w
        for c in x['po_codes']:
            if c in cat: ev[c]=ev.get(c,0)+0.5
        sup_ok=lambda c: bool(st) and (any(_sup_tokens(s_)&st for s_ in cat[c]['sup']) or bool(sheet_sup.get(c,set())&st))
        lvl,ks=wsa_find(idx,x.get('po'),x.get('bls') or set())
        wsh=[idx['sheets'][k] for k in ks]; wcodes={A(c) for s in wsh for c in s['codes']}
        x['wsa_how']=lvl; x['wsa_sheets']=[s['label'] for s in wsh]
        fams={_code_family(c) for c in ev}
        cands=set(ev)|{c for c in cat if (fams and _code_family(c) in fams) or (st and sup_ok(c)) or A(c) in wcodes}
        sc={}
        for c in cands:
            kb,k2=keys(cat[c]['desc']); s_=_desc_score(desc,cat[c]['desc'])
            for f in ('dim','th','hs','wd','gsm','len'):
                if ka[f] and kb[f] and (ka[f]-kb[f]) and (kb[f]-ka[f]): s_-=3          # a size / thickness / HS / GSM that contradicts
            if k1 and k2: s_+=4 if k1==k2 else -6
            if sup_ok(c): s_+=1.5
            sc[c]=s_+ev.get(c,0)
        def fits(c):                       # same product type and nothing in the description contradicts it
            kb,k2=keys(cat[c]['desc'])
            return bool(k1 and k2 and k1==k2) and not any(ka[f] and kb[f] and (ka[f]-kb[f]) and (kb[f]-ka[f]) for f in ('dim','th','hs','wd','gsm','len','side'))
        def same_part(c):                  # same product type and side as the line, but another size / thickness / HS / GSM
            kb,k2=keys(cat[c]['desc'])
            if not (k1 and k2 and k1==k2) or (ka['side'] and kb['side'] and ka['side']!=kb['side']): return False
            return any(ka[f] and kb[f] and (ka[f]-kb[f]) and (kb[f]-ka[f]) for f in ('dim','th','hs','wd','gsm','len'))
        x['score']=None; x['final_score']=sc.get(x['final_code']); x['bom_desc']=''; x['bom_sup']=''; x['src']=''
        fc=x['final_code']; where=', '.join(x['wsa_sheets'])
        okw=sorted((c for c in sc if A(c) in wcodes and fits(c)),key=lambda c:-sc[c])
        if id(x) in copied:                                            # 0) copied description: the unit price decides
            best=fc; sc.setdefault(best,0.0)
            how=f"same description on 2+ lines of invoice {copied[id(x)][0]} at different unit prices ({copied[id(x)][1]}): description looks copied - Final RM Code (unit price) used"
        elif okw:                                                      # 1) your WSA sheet decides
            best=okw[0]; near=[c for c in okw if sc[best]-sc[c]<1.0]
            for pref in (x['rm_code'],fc,x['match']):
                if pref in near: best=pref; break
            how=f"WSA {lvl} on sheet {where}: "+('Final RM Code confirmed' if best==fc else f"Final RM Code {fc or '-'} replaced")
        else:                                                          # 2) description + supplier + codes
            if not sc: x['code']=None; x['how']='not a BOM material (no code, no matching supplier / description)'; continue
            order=sorted(sc,key=lambda c:-sc[c]); best=order[0]
            near=[c for c in order if sc[best]-sc[c]<1.0]
            for pref in (x['rm_code'],fc,x['match']):
                if pref in near: best=pref; break
            if fc in cat and best!=fc and sc[best]-sc[fc]<2.0: best=fc; how='Final RM Code kept (nothing clearly better)'
            elif fc in cat and best==fc: how='Final RM Code confirmed by description' + (' + supplier' if sup_ok(fc) else '')
            elif not ev:
                spec=any(ka[f] for f in ('dim','th','hs','wd','gsm','len','side'))
                if not (sc[best]>=6 and k1 and k1==keys(cat[best]['desc'])[1] and spec):
                    x['code']=None; x['how']='not a BOM material (description matches nothing clearly)'; continue
                how='mapped by description + supplier (no code on the line)'
            elif best==x['rm_code']: how=f"RM Code used - its description fits; Final RM Code {fc or '-'} does not"
            else: how=f"mapped by description + supplier (Final RM Code {fc or '-'}, RM Code {x['rm_code'] or '-'})"
            if wsh and A(best) not in wcodes:                         # your WSA sheet has the same part in another size / HS = check
                cl_=sorted(c for c in cat if A(c) in wcodes and same_part(c))
                if cl_: how+=f" - CHECK: your WSA sheet {where} has this {lvl} as {cl_[0]} {cat[cl_[0]]['desc']}"
        x['code']=best; x['how']=how; x['score']=sc[best]
        x['bom_desc']=cat[best]['desc']; x['bom_sup']=' / '.join(sorted(cat[best]['sup'])); x['src']=cat[best]['src']
    return lines

IT_HDR_KEYS=('final pending qty.','final pending qty','final rm code','rm code','final qty.','final ata2','final ata','revised eta port','supplier name','ppc grn')
def _it_sheet(names):
    """Your In_Transit sheet: 'In_Transit', else the same name written another way ('In Transit', 'IN-TRANSIT', 'PPC In Transit')."""
    k_=lambda s: re.sub(r'[^a-z0-9]','',str(s).lower())
    return (next((s for s in names if s=='In_Transit'),None) or next((s for s in names if k_(s)=='intransit'),None)
            or next((s for s in names if 'intransit' in k_(s)),None))
def _num_cell(v):
    """A quantity cell -> float; a number typed as text ('1,200') counts too. None when it is not a number."""
    if v is None or isinstance(v,bool): return None
    if isinstance(v,(int,float)): return float(v)
    s_=str(v).strip().replace(',','')
    return float(s_) if re.fullmatch(r'-?\d+(?:\.\d+)?',s_) else None
def _date_cell(v):
    """A date cell -> Timestamp (day): a real date, an Excel date number, or a date typed as text (2026-11-10, 10-Nov-2026, 10 Nov 26)."""
    if isinstance(v,(dt.datetime,dt.date)): return pd.Timestamp(v).normalize()
    if isinstance(v,(int,float)) and not isinstance(v,bool): return (pd.Timestamp('1899-12-30')+pd.Timedelta(days=int(v))) if 20000<v<80000 else None
    if isinstance(v,str):
        for f_ in ('%Y-%m-%d','%Y-%m-%d %H:%M:%S','%d-%b-%Y','%d-%b-%y','%d %b %Y','%d %b %y','%b %d, %Y','%d-%B-%Y'):
            try: return pd.Timestamp(dt.datetime.strptime(v.strip(),f_))
            except ValueError: pass
    return None
def _it_formulas(path, sn, rows_, cols_):
    """How many of these cells (Excel row numbers x column indexes) hold a formula: such a cell reads as empty when the file
       was last saved outside Excel (no saved result)."""
    try: wb=load_workbook(path,read_only=True)
    except Exception as _e_: _swallow(7926,_e_); return 0
    want=set(rows_); n=0
    for i,r in enumerate(wb[sn].iter_rows(max_row=max(want),values_only=True),1):
        if i in want: n+=sum(1 for j in cols_ if j is not None and j<len(r) and isinstance(r[j],str) and r[j].startswith('='))
    wb.close(); return n
def _ppc_files():
    """Workbooks with 'PPC' in the name in the folders the PPC file is searched in (newest first)."""
    fs=[os.path.join(fo,n) for fo in (globals().get('_PPC_DIRS') or []) if os.path.isdir(fo) for n in os.listdir(fo)
        if n.lower().endswith('.xlsx') and 'ppc' in n.lower() and not n.startswith(('~$','MRP_Daily_','Inbound_'))]
    return sorted(dict.fromkeys(fs),key=os.path.getmtime,reverse=True)

def read_ppc_in_transit(path, cat=None, why=None, info=None, wsa=None):
    """Your PPC 'In_Transit' sheet (the PPC In_Transit columns + your own): one dict per shipment line with a quantity still
       pending - qty = Final Pending Qty. (else Final Qty. - PPC GRN), date = Final ATA2 / Final AtA (the plant arrival = port
       ETA + 5 days; else Revised ETA Port + 5). The material is traced by map_in_transit (supplier + description + RM Code +
       Final RM Code) against YOUR BOM. Returns the lines ([] = none pending), or None when there is no such sheet.
       Also reads: the sheet named another way ('In Transit'), title rows above the header, numbers / dates typed as text.
       why (list) gets the reason when nothing is read, and a NOTE for quantity formulas with no saved result; info (dict) the sheet."""
    why=[] if why is None else why; info={} if info is None else info; bn=os.path.basename(path)
    try: wb=load_workbook(path,read_only=True,data_only=True)
    except Exception as ex_: why.append(f'{bn} could not be opened ({type(ex_).__name__}: {ex_})'); return None
    sn=_it_sheet(wb.sheetnames)
    if not sn: why.append(f"{bn} has no In_Transit sheet (its sheets: {', '.join(wb.sheetnames)})"); wb.close(); return None
    info['sheet']=sn; raw=[]; empty=0
    for i,r in enumerate(wb[sn].iter_rows(values_only=True),1):
        if not any(v not in (None,'') for v in r):
            empty+=1
            if empty>(500 if raw else 2000): break       # the sheet is formatted down to row 1,048,000
            continue
        empty=0; raw.append((i,tuple(r)))
    wb.close()
    if not raw: why.append(f'{bn} sheet {sn} is empty'); return []
    is_hdr=lambda r: sum(k_ in {_hnorm(v) for v in r} for k_ in IT_HDR_KEYS)>=2
    hi=next((k for k,(i,r) in enumerate(raw[:30]) if is_hdr(r)),0); info['header_row']=raw[hi][0]
    hdr=[_hnorm(h) for h in raw[hi][1]]
    ix=lambda *ns: next((hdr.index(n) for n in ns if n in hdr),None)
    ixf=lambda f: next((j for j,h in enumerate(hdr) if h and f(h)),None)
    g=lambda r,i: r[i] if (i is not None and i<len(r)) else None
    iC=ix('final rm code'); iC2=ix('rm code'); iC3=ix('match'); iPW=ix('po wise material desc.'); iQ=ix('final pending qty.','final pending qty')
    if iQ is None: iQ=ixf(lambda h:'pending' in h and 'qty' in h)
    iFQ=ix('final qty.','qty'); iG=ix('ppc grn'); iA2=ix('final ata2'); iA1=ix('final ata'); iE=ix('revised eta port'); iPO=ix('po'); iB=ix('hblno.','mbl'); iHB=ix('hblno.','hbl no.','hbl'); iMB=ix('mbl')
    iV=ix('vessel / voy'); iS=ix('status'); iPD=ix('pod'); iSup=ix('supplier name'); iInv=ix('invoice number'); iDs=ix('description'); iUR=ix('unit rate')
    iETD,iEB=ix('etd'),ix('eta as per booking'); has_=_ppc_it_has(iETD,iE,iEB,iHB,iMB,iA1)
    up=lambda v: str(v or '').strip().upper()
    out=[]; noq=[]
    for xr,r in raw[hi+1:]:
        q=_num_cell(g(r,iQ))
        if q is None:
            fq=_num_cell(g(r,iFQ)); gr=_num_cell(g(r,iG))
            if fq is None and any(g(r,j_) not in (None,'') for j_ in (iC,iC2,iSup)): noq.append(xr)
            q=(fq or 0.0)-(gr or 0.0)
        if not q or q<=0: continue
        mi_=_ppc_it_missing(g(r,iETD),g(r,iE),g(r,iEB),g(r,iHB),g(r,iMB),g(r,iA1),has_)   # = ppc_it_rows (MRP_Daily): ETD + ETA + BL# + Final AtA
        if mi_: PPC_NODATE.append(_ppc_nodate_txt(g(r,iDs),g(r,iInv),q,mi_)); continue
        d=_ppc_it_date(g(r,iA2),g(r,iA1),g(r,iE),iA2 is not None,iA1 is not None)
        if d is None: PPC_NODATE.append(_ppc_nodate_txt(g(r,iDs),g(r,iInv),q)); continue   # no arrival date = not counted (user, 21-Nov-2026: 1c)
        out.append(dict(final_code=up(g(r,iC)),rm_code=up(g(r,iC2)),match=up(g(r,iC3)),po_codes=re.findall(r'\b(?:RM|PM|CS)[A-Z0-9]{6,}\b',up(g(r,iPW))),
                        description=' '.join(str(g(r,iDs) or '').split()),qty=float(q),date=d,
                        po=g(r,iPO),bl=g(r,iB),bls=_bl_set(g(r,iHB))|_bl_set(g(r,iMB)),vessel=' '.join(str(g(r,iV) or '').split()),status=' '.join(str(g(r,iS) or '').split()),pod=' '.join(str(g(r,iPD) or '').split()),
                        supplier=' '.join(str(g(r,iSup) or '').split()),invoice=g(r,iInv),price=_num_cell(g(r,iUR)),raw_code=str(g(r,iC) or '').strip()))
    if noq and _it_formulas(path,sn,noq,[iQ,iFQ]):
        why.append(f"{len(noq)} rows of {bn} sheet {sn} have a quantity formula with no saved result (the file was last saved outside Excel)"
                   f" - open {bn} in Excel and save it once")
    if not out and not why:
        why.append(f"{bn} sheet {sn}: no 'Final Pending Qty.' column (header row {raw[hi][0]})" if iQ is None and iFQ is None
                   else f"{bn} sheet {sn}: {len(raw)-hi-1} rows, none with a pending quantity above 0")
    return map_in_transit(out,cat or {},wsa)

def ppc_arrivals(DS, dates):
    """PPC In_Transit -> arrivals per MRP material and day (MRP units: the same code conversions as PPC In_Transit, e.g.
       RMTTT00000018 = 2 x RMTTT00000010). A late shipment (arrival before today, still pending) arrives today.
       why = why there are no arrivals (shown in T3 of the WSA sheets, the Inbound sheet 'PPC In Transit', the end of the console)."""
    why=[]; info={}
    none_=lambda lines_=None: dict(arr={},rows=[],skipped=[],src=None,fixes=[],lines=lines_ or [],why=why,sheet=info.get('sheet'))
    if not PPC_SRC[0]:
        why.append('no PPC file found (a workbook with Count + BOM sheets next to MRP.xlsx or the script)'); return none_()
    cat=in_transit_catalog(PPC_SRC[0],MIS_PATH_[0] or globals().get('MIS'),DS)
    wsa_=dict(idx=wsa_ref_index(WSA_REFS[0]),alias=DS.get('alias') or {}) if WSA_REFS[0] else None
    rows=read_ppc_in_transit(PPC_SRC[0],cat,why,info,wsa_)
    if rows is None:
        oth=[os.path.basename(p_) for p_ in _ppc_files() if os.path.abspath(p_)!=os.path.abspath(PPC_SRC[0]) and _it_sheet(_sheetnames(p_))]
        if oth: why.append(f"an In_Transit sheet is in {', '.join(oth)}, but the PPC file used is {os.path.basename(PPC_SRC[0])} "
                           "(the newest workbook with Count + BOM sheets) - keep one PPC file in the folder")
        return none_()
    if not rows: return none_()
    fixes=[x for x in rows if x.get('code') and x['code']!=x['final_code']]
    d0=pd.Timestamp(dates[0]).normalize(); nd=len(dates); idx={pd.Timestamp(d).normalize():k for k,d in enumerate(dates)}
    alias=DS.get('alias') or {}; fac=DS.get('alias_factor') or {}; mats={b['mat'] for b in DS['blocks']}
    arr={}; used=[]; skipped=[]; sup={}
    for x in rows:
        c=x['code']
        x['mat']=None; x['qty_mrp']=None; x['arrive']=None; x['late']=False; x['in_mrp']=False
        if not c: skipped.append(f"{x['description'][:28]} - {x['supplier'][:24]} PO {x['po']} qty {x['qty']:,.0f}"); continue
        m=alias.get(c,c); q=x['qty']*fac.get(c,1.0)
        if x['date'] is None: skipped.append(f"{c} PO {x['po']} qty {x['qty']:,.0f} (no arrival date)"); continue
        x.update(mat=m,qty_mrp=q,arrive=max(x['date'],d0),late=x['date']<d0,in_mrp=m in mats)
        k=idx.get(max(x['date'],d0))
        if k is None: continue                                        # after the planning horizon
        arr.setdefault(m,[0.0]*nd)[k]+=q; used.append(x); sup.setdefault(m,{}).setdefault(k,set()).add(x['supplier'] or '?')
    if not arr:
        nd_=sum(1 for x in rows if x.get('code') and x['date'] is None); nm_=sum(1 for x in rows if not x.get('code'))
        why.append(f"{len(rows)} lines read from {os.path.basename(PPC_SRC[0])}, none arrives in the planning period ({nd_} with no arrival date, {nm_} not BOM materials)")
    return dict(arr=arr,rows=used,skipped=skipped,src=os.path.basename(PPC_SRC[0]),fixes=fixes,lines=rows,why=why,sheet=info.get('sheet'),sup=sup)
def _ppc_arrivals_safe(DS, dates):
    """ppc_arrivals; an unexpected cell in the In_Transit sheet must not stop the WSA alignment - it is reported instead."""
    try: return ppc_arrivals(DS,dates)
    except Exception as ex_:
        import traceback
        tb_=traceback.extract_tb(ex_.__traceback__)[-1]
        return dict(arr={},rows=[],skipped=[],src=None,fixes=[],lines=[],sheet=None,
                    why=[f'error reading the PPC In_Transit sheet ({type(ex_).__name__}: {ex_}, line {tb_.lineno})'])

def align_capture(DS, pp_rows, dates):
    """After the OEM step of build_new_mrp: per material its non-OEM plan, OEM plan, own stock and per (MIS customer, line) use;
       the module plan (and MW) per (MIS customer, line)."""
    n_=len(dates); mods={}
    for x in pp_rows or []:
        k=(_norm(x.get('mis_cust') or x.get('cust')),' '.join(str(x.get('line') or '').split()))
        a=mods.setdefault(k,dict(cust=[],mis=' '.join(str(x.get('mis_cust') or x.get('cust') or '').split()),line=k[1],v=[0.0]*n_,mw=[0.0]*n_))
        cu=' '.join(str(x.get('cust') or '').split())
        if cu and cu not in a['cust']: a['cust'].append(cu)
        for i,v in enumerate((x.get('v') or [])[:n_]): a['v'][i]+=float(v or 0)
        for i,v in enumerate((x.get('mw') or [])[:n_]): a['mw'][i]+=float(v or 0)
    bl={}
    for b in DS['blocks']:
        cl={}; oem=set(); disp={}
        for src,is_oem in ((b.get('cust_line') or {},False),(b.get('cust_line_oem') or {},True)):
            for k,v in src.items():
                cu,ln=(k if isinstance(k,tuple) else (k,None))
                key=(_norm(cu),' '.join(str(ln or '').split()))
                disp.setdefault(key,(' '.join(str(cu or '').split()),' '.join(str(ln or '').split())))
                a=cl.setdefault(key,[0.0]*n_)
                for i,q in enumerate(list(v)[:n_]): a[i]+=float(q or 0)
                if is_oem: oem.add(key)
        sup=[]                                      # stock by supplier = the MRP_Daily supplier rows: MB52 batches, or the count by supplier
        for sp in b.get('sup') or []:
            nm=str(sp.get('name') or ''); st=float(sp.get('stock') or 0.0)
            sup.append((nm,st))                     # counted: already the count by supplier, rolled forward (apply_count_by_supplier)
        bl[b['mat']]=dict(plan=list(b.get('plan') or [0.0]*n_),plan_oem=list(b.get('plan_oem') or [0.0]*n_),cl=cl,oem=oem,disp=disp,
                          basis=str(b.get('plan_basis') or ''),stock=float(b.get('stock') or 0),jw=float(b.get('st_jw') or 0) if b.get('jw_out') else 0.0,
                          uom=b.get('uom'),common=b.get('common'),desc=b.get('desc'),sup=sup)
    return dict(dates=[pd.Timestamp(d).normalize() for d in dates],blocks=bl,mods=mods,alias=dict(DS.get('alias') or {}),ppc=_ppc_arrivals_safe(DS,dates))

def _wsa_lookup(src):
    """WSA 'Lookup' sheet: (MIS plant, MIS customer) -> customer column label; customer -> labels (any plant)."""
    by={}; bycu={}
    try:
        for r in wb_ro(src)['Lookup'].iter_rows(min_row=2,values_only=True):
            if not r or len(r)<4 or not r[1] or not r[2] or not r[3]: continue
            pl,cu,lab=_norm(r[1]),_norm(r[2]),' '.join(str(r[3]).split())
            by[(pl,cu)]=lab; bycu.setdefault(cu,Counter())[lab]+=1
    except Exception as ex_: print(f"NOTE: WSA 'Lookup' sheet not read ({type(ex_).__name__}: {ex_}) - customer columns named after the MIS customer")
    return by,bycu

def _plant_keys(line):
    """Plan production line -> the plant names the WSA Lookup uses ('USA Line 3' -> 'l3', 'Runergy' -> 'runergy l1' ...)."""
    n=_norm(line); m=re.match(r'^(?:usa\s*)?line\s*(\d+)$',n) or re.match(r'^l\s*(\d+)$',n)
    if m: return [f'l{m.group(1)}']
    if n.startswith('runergy') or n.startswith('runenergy'): return [n,'runergy l1','runergy l2','runergy']
    if n.startswith('imperial') or n.startswith('ic star'): return [n,'imperial star']
    if n.startswith('arizona'): return [n,'arizona']
    if n.startswith('inox'): return [n,'inox 1','inox 2']
    return [n]

def _hnorm(x): return ' '.join(str(x if x is not None else '').replace('\n',' ').split()).lower()
def _hdr_role(h):
    """Planning-sheet header (normalised) -> field."""
    if not h: return None
    if h.startswith('required plant date'): return 'date'
    if h in ('d','day'): return 'day'
    if h.startswith('balance'): return 'bal'
    if h.startswith('planned modules') or h.startswith('modules per day'): return 'plan'
    if h.startswith('demand'): return 'dem'
    if h.startswith('actual modules'): return 'actmod'
    if h.startswith('actual mw'): return 'actmw'
    if h=='actual' or h.startswith('actual consum'): return 'act'
    if h.startswith('mw per day'): return 'mw'
    if h.startswith('plant days'): return 'pdays'
    if h.startswith('plant + port days (ppc'): return 'ppcpp'
    if h.startswith('ppc in transit (supplier'): return 'ppcsup'
    if h.startswith('ppc in transit'): return 'ppcit'
    if h.startswith('plant + port'): return 'pport'
    if h.startswith('delivered'): return 'dlv'
    if h.startswith('scheduled'): return 'sch'
    if h=='po' or h.startswith('po '): return 'po'
    if h.startswith('shipper'): return 'ship'
    if h.startswith('remarks'): return 'rem'
    if h.startswith('booking'): return 'book'
    if h.startswith('readiness'): return 'ready'
    if h=='etd' or h.startswith('etd '): return 'etd'
    if h.startswith('bl'): return 'bl'
    if h.startswith('vessel'): return 'vessel'
    if h.startswith('invoice'): return 'inv'
    if h.startswith('reference'): return 'ref'
    if h.startswith('status'): return 'status'
    if h.startswith('vendor rule: not usable'): return 'vlock'
    if h.startswith('vendor rule'): return 'vpool'
    return 'other'

def _wsa_read(wv, wf):
    """Planning blocks of a WSA file in any layout: the column of each field by its header, the customer columns (between
       Balance and Delivered), codes, unit, material name and the row of each date."""
    out=[]
    for sh in wv.sheetnames:
        if sh in SKIP or sh in (ALIGN_SHEET,ARCHIVE_SHEET): continue
        ws=wv[sh]
        if ws.sheet_state!='visible': continue
        rv=[list(r) for r in ws.iter_rows(values_only=True)]
        starts=sorted([(r+1,c+1) for r,row in enumerate(rv[:12]) for c,x in enumerate(row[:200]) if str(x).strip().startswith('Required Plant Date')],key=lambda x:x[1])
        if not starts: continue
        rf=[list(r) for r in wf[sh].iter_rows(values_only=True)]
        maxc=max(len(r) for r in rv)
        for bi,(hr,c0) in enumerate(starts):
            c1=(starts[bi+1][1]-1) if bi+1<len(starts) else maxc
            role={'date':c0}; cust=[]; dropped=[]
            for c in range(c0+1,c1+1):
                raw=cv(rv,hr,c); ro=_hdr_role(_hnorm(raw))
                if ro is None: continue
                if ro=='day' and c==c0+1: role['day']=c; continue
                if ro in ('other','day'):
                    if 'bal' in role and 'dlv' not in role: cust.append((c,' '.join(str(raw).split())))
                    elif 'dlv' in role: dropped.append(' '.join(str(raw).split()))
                    continue
                if ro=='rem' and 'rem' in role: ro='rem2'
                if ro=='act': role['act']=c; continue              # 'Actual consumption' twice (EPE 480): the last one is the one Balance uses
                role.setdefault(ro,c)
            if 'bal' not in role or 'sch' not in role or 'dem' not in role: continue
            if 'book' not in role and 'rem' in role and 'ready' in role:          # a Booking column without a header (EPE 420 GSM)
                for c in range(role['rem']+1,role['ready']):
                    if not _hnorm(cv(rv,hr,c)): role['book']=c; break
            codes=[]; uom=''; mname=''
            for r in range(1,hr):
                for c in range(c0,c1+1):
                    x=cv(rv,r,c)
                    if isinstance(x,str):
                        if CODE.match(x.strip()) and x.strip() not in codes: codes.append(x.strip())
                        if x.strip().startswith('Unit of Measurement'): uom=str(cv(rv,r,c+2) or '').strip()
                        if x.strip().startswith('Supplier & Material'): mname=str(cv(rv,r,c+2) or '').strip()
            dates={}
            for r in range(hr+1,len(rv)+1):
                x=cv(rv,r,c0)
                if isinstance(x,(dt.datetime,dt.date)): dates.setdefault(pd.Timestamp(x).normalize(),r)
            out.append(dict(sheet=sh,block=bi,nblocks=len(starts),label=LABELS.get((sh,bi)) or (sh if len(starts)==1 else f'{sh} #{bi+1}'),
                            hr=hr,c0=c0,c1=c1,role=role,cust=cust,dropped=dropped,codes=codes,uom=uom,mname=mname,dates=dates,rv=rv,rf=rf))
    return out

class _Styles:
    """styles.xml of the WSA file + the few cell formats the rebuilt sheets use (appended once, reused on the next run)."""
    def __init__(self, sty):
        import copy
        from lxml import etree
        self.E=etree; self.cp=copy; self.M='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
        self.st=etree.fromstring(sty); M=self.M
        self.nfs=self.st.find(M+'numFmts')
        if self.nfs is None: self.nfs=etree.Element(M+'numFmts',count='0'); self.st.insert(0,self.nfs)
        self.fonts=self.st.find(M+'fonts'); self.fills=self.st.find(M+'fills'); self.borders=self.st.find(M+'borders'); self.xfs=self.st.find(M+'cellXfs')
        self.dxfs=self.st.find(M+'dxfs')
        if self.dxfs is None:
            self.dxfs=etree.Element(M+'dxfs',count='0'); cs=self.st.find(M+'cellStyles')
            (cs if cs is not None else self.xfs).addnext(self.dxfs)
        self.f0=self.fonts.findall(M+'font')[0]
    @staticmethod
    def _sig(e): return (e.tag,tuple(sorted(e.attrib.items())),tuple(_Styles._sig(c) for c in e))
    def _add(self, parent, el, tag):
        M=self.M; kids=parent.findall(M+tag); sg=self._sig(el)
        for i,k in enumerate(kids):
            if self._sig(k)==sg: return i
        parent.append(el); parent.set('count',str(len(kids)+1)); return len(kids)
    def numfmt(self, code):
        M=self.M
        for e in self.nfs.findall(M+'numFmt'):
            if e.get('formatCode')==code: return int(e.get('numFmtId'))
        nid=max([163]+[int(e.get('numFmtId')) for e in self.nfs.findall(M+'numFmt')])+1
        self.E.SubElement(self.nfs,M+'numFmt',numFmtId=str(nid),formatCode=code); self.nfs.set('count',str(len(self.nfs.findall(M+'numFmt'))))
        return nid
    def font(self, bold=False, size=None, rgb=None):
        M=self.M; f=self.cp.deepcopy(self.f0)
        for ch in list(f):
            if ch.tag in (M+'b',M+'sz',M+'color') or (ch.tag==M+'scheme' and rgb): f.remove(ch)
        if bold: f.insert(0,self.E.Element(M+'b'))
        sz=self.E.Element(M+'sz'); sz.set('val',str(size or self._size0())); f.append(sz)
        if rgb: c=self.E.Element(M+'color'); c.set('rgb',rgb); f.append(c)
        else:
            c0=self.f0.find(M+'color')
            if c0 is not None: f.append(self.cp.deepcopy(c0))
        return self._add(self.fonts,f,'font')
    def _size0(self):
        s=self.f0.find(self.M+'sz'); return s.get('val') if s is not None else '11'
    def fill(self, rgb):
        M=self.M; f=self.E.Element(M+'fill'); p=self.E.SubElement(f,M+'patternFill',patternType='solid')
        self.E.SubElement(p,M+'fgColor',rgb=rgb); self.E.SubElement(p,M+'bgColor',indexed='64')
        return self._add(self.fills,f,'fill')
    def border(self):
        M=self.M; b=self.E.Element(M+'border')
        for side in ('left','right','top','bottom'):
            s=self.E.SubElement(b,M+side,style='thin'); self.E.SubElement(s,M+'color',rgb='FFBFBFBF')
        self.E.SubElement(b,M+'diagonal')
        return self._add(self.borders,b,'border')
    def xf(self, nf=0, font=0, fill=0, border=0, h=None, v=None, wrap=False):
        M=self.M; a={'numFmtId':str(nf),'fontId':str(font),'fillId':str(fill),'borderId':str(border),'xfId':'0'}
        if nf: a['applyNumberFormat']='1'
        if font: a['applyFont']='1'
        if fill: a['applyFill']='1'
        if border: a['applyBorder']='1'
        e=self.E.Element(M+'xf',**a)
        if h or v or wrap:
            e.set('applyAlignment','1'); al=self.E.SubElement(e,M+'alignment')
            if h: al.set('horizontal',h)
            if v: al.set('vertical',v)
            if wrap: al.set('wrapText','1')
        return self._add(self.xfs,e,'xf')
    def dxf(self, font_rgb, fill_rgb, bold=False):
        M=self.M; d=self.E.Element(M+'dxf'); f=self.E.SubElement(d,M+'font')
        if bold: self.E.SubElement(f,M+'b')
        self.E.SubElement(f,M+'color',rgb=font_rgb)
        p=self.E.SubElement(self.E.SubElement(d,M+'fill'),M+'patternFill'); self.E.SubElement(p,M+'bgColor',rgb=fill_rgb)
        return self._add(self.dxfs,d,'dxf')
    def yellow_xfs(self):
        """cellXfs indexes with a yellow fill (the planner's typed balances)."""
        M=self.M; fills=self.fills.findall(M+'fill'); out=set()
        for i,x in enumerate(self.xfs.findall(M+'xf')):
            try:
                pf=fills[int(x.get('fillId') or 0)].find(M+'patternFill'); fg=pf.find(M+'fgColor') if pf is not None else None
                if fg is not None and str(fg.get('rgb') or '').upper() in ('FFFFFF00','FFFF00'): out.add(str(i))
            except (IndexError,ValueError): pass
        return out
    def bytes(self): return self.E.tostring(self.st,xml_declaration=True,encoding='UTF-8',standalone=True)

def _xcell(ref, v=None, s=0, f=None):
    """One <c> element as text: formula + cached value, a number, a date (as serial), text (inline) or an empty styled cell."""
    from xml.sax.saxutils import escape
    sa=f' s="{s}"' if s else ''
    if isinstance(v,(dt.datetime,dt.date)): v=(pd.Timestamp(v).normalize()-pd.Timestamp('1899-12-30')).days
    if isinstance(v,float) and v!=v: v=None                        # NaN
    def num(x): return str(int(x)) if float(x)==int(x) and abs(x)<1e15 else repr(float(x))
    if f:
        fx='<f>'+escape(f)+'</f>'
        if isinstance(v,bool): return f'<c r="{ref}"{sa} t="b">{fx}<v>{int(v)}</v></c>'
        if isinstance(v,(int,float)): return f'<c r="{ref}"{sa}>{fx}<v>{num(v)}</v></c>'
        if isinstance(v,str): return f'<c r="{ref}"{sa} t="str">{fx}<v>{escape(v)}</v></c>'
        return f'<c r="{ref}"{sa}>{fx}</c>'
    if isinstance(v,bool): return f'<c r="{ref}"{sa} t="b"><v>{int(v)}</v></c>'
    if isinstance(v,(int,float)): return f'<c r="{ref}"{sa}><v>{num(v)}</v></c>'
    if isinstance(v,str) and v!='':
        sp=' xml:space="preserve"' if v!=v.strip() or '\n' in v else ''
        return f'<c r="{ref}"{sa} t="inlineStr"><is><t{sp}>{escape(v)}</t></is></c>'
    return f'<c r="{ref}"{sa}/>' if s else ''

def _carry_formula(fx, r_old, colmap, r_new):
    """The planner's formula in a delivery column, moved to the new layout: only constants (=10000*4) or references to cells of
       the same row (=A747-40) whose column exists in the new layout; anything else -> None (the value is copied instead)."""
    s=str(fx)[1:] if str(fx).startswith('=') else str(fx)
    if '!' in s or ':' in s or '"' in s or '[' in s: return None
    from openpyxl.utils import column_index_from_string as CI
    out=[]; pos=0
    for m in re.finditer(r'(\$?)([A-Z]{1,3})(\$?)(\d+)',s):
        if m.start()>0 and (s[m.start()-1].isalpha() or s[m.start()-1]=='_'): return None
        if int(m.group(4))!=r_old: return None
        nc=colmap.get(CI(m.group(2)))
        if nc is None: return None
        out.append(s[pos:m.start()]+m.group(1)+nc+m.group(3)+str(r_new)); pos=m.end()
    out.append(s[pos:])
    if re.search(r'[A-Za-z]{2,}\(',s): return None                     # functions: not moved
    return ''.join(out)

