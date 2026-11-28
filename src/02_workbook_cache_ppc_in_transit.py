# ---------- read-only workbook cache: each input workbook / sheet is parsed once per run (big speed-up) ----------
def _ref_formula_counts(path):
    """{sheet name: number of formulas containing #REF!} read straight from the xlsx XML (no workbook is opened)."""
    import zipfile, html, posixpath
    out={}
    try:
        with zipfile.ZipFile(path) as z_:
            wbx=z_.read('xl/workbook.xml').decode('utf-8'); rel=z_.read('xl/_rels/workbook.xml.rels').decode('utf-8')
            tgt={}
            for m_ in re.finditer(r'<Relationship\b[^>]*>',rel):
                t_=m_.group(0); i_=re.search(r'\bId="([^"]+)"',t_); g_=re.search(r'\bTarget="([^"]+)"',t_)
                if i_ and g_: tgt[i_.group(1)]=g_.group(1)
            for m_ in re.finditer(r'<(?:\w+:)?sheet\b[^>]*>',wbx):
                t_=m_.group(0); nm_=re.search(r'\bname="([^"]*)"',t_); rid_=re.search(r'\br:id="([^"]+)"',t_)
                if not nm_ or not rid_ or rid_.group(1) not in tgt: continue
                pth_=tgt[rid_.group(1)].lstrip('/'); pth_=pth_ if pth_.startswith('xl/') else posixpath.join('xl',pth_)
                try: x_=z_.read(pth_).decode('utf-8','ignore')
                except KeyError: continue
                k_=len(re.findall(r'<f\b[^>]*>[^<]*#REF!',x_))
                if k_: out[html.unescape(nm_.group(1))]=k_
    except Exception as _e_: _swallow(0,_e_)
    return out
_WB_CACHE={}
class _CachedSheet:
    def __init__(self,rows,title): self._rows=rows; self.title=title; self.max_row=len(rows); self.max_column=max((len(r) for r in rows),default=0)
    def iter_rows(self,min_row=1,max_row=None,min_col=1,max_col=None,values_only=True):
        for r in self._rows[(min_row or 1)-1:max_row]:
            yield r[(min_col or 1)-1:max_col] if (max_col or (min_col or 1)>1) else r

USER_MB52=[None]   # SAP MB52 stock: the PPC file's 'MB52' sheet (user 20-Nov-2026), else your MB52*.xlsx; MIS MB52 only gives the extra columns (plant name, final code, key, company, categories)
USER_MB52_SHEET=[None]   # sheet of USER_MB52 that holds the MB52 lines (None = first sheet of MB52.xlsx)
def ppc_mb52_sheet(p):
    """'MB52' sheet of the PPC file when its header has Plant / Material / Unrestricted, else None."""
    try:
        w=load_workbook(p,read_only=True,data_only=True); n_=next((n for n in w.sheetnames if n.strip().lower()=='mb52'),None)
        if not n_: return None
        h=next(w[n_].iter_rows(max_row=1,values_only=True)); hl=[str(x or '').strip().lower() for x in h]
        return n_ if {'plant','material','unrestricted'}<=set(hl) else None
    except Exception as _e_: _swallow(0,_e_); return None
def _is_mis_file(p):
    try: return {'MB52','Standard BOM'}<=set(_sheetnames(p))
    except Exception as _e_: _swallow(253,_e_); return False
def find_user_mb52(folders):
    """Newest MB52*.xlsx in the folders whose first sheet has Plant / Material / Unrestricted (not the MIS file)."""
    best=[]
    for fo in folders:
        for f in glob.glob(os.path.join(fo,'*.xlsx')):
            bn=os.path.basename(f)
            if bn.startswith('~$') or 'mb52' not in bn.lower() or 'dashboard' in bn.lower(): continue
            try:
                w=load_workbook(f,read_only=True,data_only=True); h=next(w[w.sheetnames[0]].iter_rows(max_row=1,values_only=True))
                hl=[str(x or '').strip().lower() for x in h]
                if {'plant','material','unrestricted'}<=set(hl): best.append((os.path.getmtime(f),f))
            except Exception as _e_: _swallow(265,_e_); continue
    return max(best)[1] if best else None
def user_mb52_rows(mis_rows, wbc):
    """Your MB52 in the MIS MB52 layout (20 columns). Your file gives Plant ... Quality Inspection (columns found by header);
       Plant Master.Plant Name / Company by plant and Final Material Code / Dashboard Key / Comman Name Cat.1 / Cat.2 by material
       come from MIS MB52 (else MIS Common Name / 'plant name-final code'); Total Stock = Unrestricted + Quality Inspection (MIS rule)."""
    hdr=list(mis_rows[0]); body=mis_rows[1:]
    pl_name={}; pl_comp={}; fin={}; key={}; cat={}
    for r in body:
        if not r or len(r)<20: continue
        pc=str(r[0] or '').strip().split('.')[0]
        if r[13]: pl_name.setdefault(pc,r[13])
        if r[17]: pl_comp.setdefault(pc,r[17])
        m=str(r[2] or '').strip()
        if m:
            if r[15]: fin.setdefault(m,str(r[15]).strip())
            if r[16]: key.setdefault(m,r[16])
            if r[18] or r[19]: cat.setdefault(m,(r[18],r[19]))
    try:
        for r in wbc['Common Name'].iter_rows(min_row=2,values_only=True):
            if r and r[0] and len(r)>4: cat.setdefault(str(r[0]).strip(),(r[3],r[4]))
    except Exception as _e_: _swallow(286,_e_); pass
    try:
        for r in wbc['Plant Master'].iter_rows(min_row=2,values_only=True):
            if r and r[0] and len(r)>1 and r[1]: pl_name.setdefault(str(r[0]).strip().split('.')[0],r[1])
    except Exception as _e_: _swallow(290,_e_); pass
    uw=load_workbook(USER_MB52[0],read_only=True,data_only=True); us=uw[USER_MB52_SHEET[0] or uw.sheetnames[0]]
    urows=list(us.iter_rows(values_only=True)); uh=[str(x or '').strip().lower() for x in urows[0]]
    want=['plant','name 1','material','batch','storage location','material description','unrestricted','base unit of measure',
          'descr. of storage loc.','blocked','material group','material type','quality inspection']
    ix=[uh.index(w) if w in uh else None for w in want]
    num=lambda v: float(v) if isinstance(v,(int,float)) and not isinstance(v,bool) else (float(str(v).replace(',','')) if str(v or '').replace(',','').replace('.','',1).lstrip('-').isdigit() else 0.0)
    out=[tuple(hdr)]
    for r in urows[1:]:
        if not r or all(v in (None,'') for v in r): continue
        v=[(r[j] if j is not None and j<len(r) else None) for j in ix]
        if not v[2]: continue
        for k_ in (6,9,12): v[k_]=num(v[k_])
        pc=str(v[0] or '').strip().split('.')[0]; m=str(v[2]).strip()
        pn=pl_name.get(pc); f_=fin.get(m,m); c1,c2=cat.get(f_,cat.get(m,(None,None)))
        try: v[0]=int(pc)
        except ValueError: pass
        row=v+[pn,v[6]+v[12],f_,key.get(m) or (f"{pn}-{f_}" if pn else None),pl_comp.get(pc,pn),c1,c2]
        out.append(tuple(row+[None]*(len(hdr)-len(row))))
    print(f"Stock (MB52) read from: {os.path.basename(USER_MB52[0])}"+(f" sheet '{USER_MB52_SHEET[0]}'" if USER_MB52_SHEET[0] else '')+f" - {len(out)-1} lines (MIS MB52 used only for plant name, final code, key, company, categories)",flush=True)
    return out

ROWS_CACHE_DIR=os.environ.get('WSA_ROWS_CACHE')      # set only by the step-1 process for its step-2 child; removed when the run ends
def _rows_cache_file(path,name):
    if not ROWS_CACHE_DIR: return None
    import hashlib
    try: st_=os.stat(path)
    except OSError: return None
    k_=hashlib.md5(f'{os.path.abspath(path)}|{st_.st_size}|{st_.st_mtime_ns}|{name}'.encode('utf-8')).hexdigest()
    return os.path.join(ROWS_CACHE_DIR,k_+'.pkl')
def _rows_cache_get(path,name):
    f_=_rows_cache_file(path,name)
    if f_ and os.path.exists(f_):
        try:
            import pickle
            with open(f_,'rb') as h_: return pickle.load(h_)
        except Exception as _e_: _swallow(326,_e_); return None
    return None
def _rows_cache_put(path,name,rows):
    f_=_rows_cache_file(path,name)
    if f_:
        try:
            import pickle
            with open(f_+'.tmp','wb') as h_: pickle.dump(rows,h_,protocol=pickle.HIGHEST_PROTOCOL)
            os.replace(f_+'.tmp',f_)
        except Exception as _e_: _swallow(335,_e_); pass
# ---------- UNIT = YOUR BOM UNIT (user 26-Nov-2026): every quantity is converted to the PPC BOM 'UoM' where it ENTERS ----------
UNIT_SYN={'EA':'PC','NOS':'PC','PCS':'PC','LITER':'L','LITRE':'L','LTR':'L','ROLL':'ROL','RL':'ROL','DRUM':'DR','DRUMS':'DR'}   # same unit, other spelling (answers 3-5)
UNIT_T={}          # loaded once: bom = code -> your BOM unit (as written), fac = (code, unit) -> [(supplier code, factor)] (PPC Count), sap = code -> SAP unit
UNIT_LOG={}        # (source, code, unit, BOM unit, supplier code, factor, how) -> [lines, qty in, qty out]  (Check: every converted line)
UNIT_NONE={}       # (source, code, unit, BOM unit) -> [lines, qty]  (Check 'UNIT ?': no conversion known, LEFT OUT of the numbers)
def _usyn(u):
    u=' '.join(str(u if u is not None else '').split()).upper()
    return '' if u in ('','NONE','-') else UNIT_SYN.get(u,u)
def _unit_tables():
    """Your BOM unit per code (PPC 'BOM' UoM) and the conversions of the PPC 'Count' sheet: a row whose Base Unit differs from its
       BOM UOM gives 'Quantity' = BOM units per base unit for (code, base unit, Supplier Code), e.g. RTV 006 DR -> KG: HU 270, DOW 250;
       laser labels BOX -> PC 1,000. SAP unit per code = MB52 Base Unit (PPC MB52 sheet, else MIS MB52), MIS Common Name, PPC Count."""
    if UNIT_T: return UNIT_T
    UNIT_T.update(bom={},fac={},sap={})
    if not PPC_SRC[0]: return UNIT_T
    try:
        w=_wb_ro_raw(PPC_SRC[0])
        for d in bom_rows_std([list(r) for r in w['BOM'].iter_rows(values_only=True)]) or []:
            u=' '.join(str(d.get('UoM') or '').split()).upper()
            if _usyn(u): UNIT_T['bom'].setdefault(str(d['Component']).strip(),u)
        rows=[list(r) for r in _count_main_table(w).iter_rows(values_only=True)] if 'Count' in w.sheetnames else []
        hi=next((i for i,r in enumerate(rows[:10]) if r and str(r[0] or '').strip()=='Material'),None)
        if hi is not None:
            hl=[_hnorm(x) for x in rows[hi]]; ix=lambda k: hl.index(k) if k in hl else None
            iu,ib,iq,isc=ix('base unit of measure'),ix('bom uom'),ix('quantity'),ix('supplier code')
            g=lambda r,i: r[i] if i is not None and i<len(r) else None
            for r in rows[hi+1:]:
                c=str(g(r,0) or '').strip(); f=g(r,iq)
                if not c: continue
                if _usyn(g(r,iu)): UNIT_T['sap'].setdefault(('count',c),_usyn(g(r,iu)))
                if not isinstance(f,(int,float)) or isinstance(f,bool) or f<=0 or _usyn(g(r,iu))==_usyn(g(r,ib)): continue
                L=UNIT_T['fac'].setdefault((c,_usyn(g(r,iu)),_usyn(g(r,ib))),[]); sc=str(g(r,isc) or '').strip().upper()
                if all(s_!=sc for s_,_ in L): L.append((sc,float(f)))
        for p_,sn_ in ((PPC_SRC[0],'MB52'),(MIS_PATH_[0],'MB52'),(MIS_PATH_[0],'Common Name')):
            if not p_ or sn_ not in _sheetnames(p_): continue
            w_=load_workbook(p_,read_only=True,data_only=True); rr=list(w_[sn_].iter_rows(values_only=True)); w_.close()   # raw (not yet relabelled) units
            if not rr: continue
            hl=[_hnorm(x) for x in rr[0]]
            im,iu=(hl.index('material') if 'material' in hl else (0 if sn_=='Common Name' else None)),(hl.index('base unit of measure') if 'base unit of measure' in hl else (hl.index('uom') if 'uom' in hl else None))
            if im is None or iu is None: continue
            for r in rr[1:]:
                if r and len(r)>max(im,iu) and r[im] and _usyn(r[iu]): UNIT_T['sap'].setdefault(str(r[im]).strip(),_usyn(r[iu]))
    except Exception as _e_: _swallow(0,_e_)
    return UNIT_T
def bom_unit(code):
    """Your BOM unit of a code (PPC BOM 'UoM', as written), '' = not in your BOM (the SAP unit stays, e.g. RMALF00000131 SET)."""
    return _unit_tables()['bom'].get(str(code or '').strip(),'')
def sap_unit(code):
    T=_unit_tables(); c=str(code or '').strip()
    return T['sap'].get(c) or T['sap'].get(('count',c)) or ''
def _sup_pick(sup, opts):
    """Which supplier's conversion: one supplier = it; else the supplier code in the batch ('1HU096320K' -> HU), the supplier
       code / name ('SHANGHAI HUITIAN ...' -> HU, word starts with the code); unknown = the MAIN supplier (first PPC Count row, answer 9)."""
    if len(opts)==1: return opts[0]+('only',)
    s=str(sup or '').strip().upper(); codes=[c for c,_ in opts if c]
    hit=[c for c in codes if c==s]
    if not hit and s and ' ' not in s and re.search(r'\d',s) and len(s)>=6:       # an SAP batch
        mk=None
        try: mk=str(batch_make(s) or '').upper()
        except Exception as _e_: _swallow(0,_e_)
        hit=[c for c in codes if c==mk] or [c for c in codes if s[1:1+len(c)]==c or s[:len(c)]==c]
    if not hit and s:
        ws=[w for w in re.split(r'[^A-Z0-9]+',SUPPLIER_ALIAS.get(s,s)) if w]
        hit=[c for c in codes if len(c)>=2 and any(w.startswith(c) for w in ws)]
    if hit:
        c=max(hit,key=len); return next(o for o in opts if o[0]==c)+('supplier',)
    return opts[0]+('main supplier (supplier not known)',)
def to_bom(code, qty, unit=None, sup=None, src='', log=True):
    """ONE conversion of a quantity into YOUR BOM unit (user 26-Nov-2026, 'UNIT = YOUR BOM UNIT'). unit = the unit the quantity is in
       (None = the SAP unit of the code); sup = supplier code / name / SAP batch. Same unit or a synonym (EA = NOS = PC, LITER = L) = unchanged;
       other unit = x the PPC Count 'Quantity' of (code, unit, supplier); no conversion known = None (line LEFT OUT, Check 'UNIT ?')."""
    if not isinstance(qty,(int,float)) or isinstance(qty,bool): return qty
    c=str(code or '').strip(); b=bom_unit(c)
    u=_usyn(unit) if unit not in (None,'') else sap_unit(c)
    if not b or not u or u==_usyn(b): return qty
    opts=_unit_tables()['fac'].get((c,u,_usyn(b)))
    if not opts:
        if log: e=UNIT_NONE.setdefault((src,c,u,b),[0,0.0]); e[0]+=1; e[1]+=float(qty)
        return None
    sc,f,how=_sup_pick(sup,opts)
    if log: e=UNIT_LOG.setdefault((src,c,u,b,sc,f,how),[0,0.0,0.0]); e[0]+=1; e[1]+=float(qty); e[2]+=float(qty)*f
    return float(qty)*f
# sheet -> (code column, [(qty column, its unit column or None = SAP unit)], supplier / batch column, unit columns to relabel) by header name
UNIT_SHEETS={'MB52':('material',[('total stock','base unit of measure'),('unrestricted','base unit of measure'),('blocked','base unit of measure'),
                                 ('quality inspection','base unit of measure')],'batch',['base unit of measure']),
             'Total Plan Req':('standard bom.material code',[('standard bom.bom','standard bom.uom'),('total requirement','standard bom.uom')],None,['standard bom.uom']),
             'Standard BOM':('material code',[('bom','uom')],None,['uom']),
             'Common Name':('material code',[],None,['uom']),
             'RM Status for OEM JW':('material code',[(k,'uom') for k in ('usa','usa warehouse','usa arizona','jw wsa to oem','usa depot')],'make',['uom']),
             'PR':('material',[('quantity requested','unit of measure'),('quantity ordered','unit of measure')],'name of supplier',['unit of measure']),
             'PO':('material',[('order quantity','order unit'),('still to be delivered (qty)','order unit'),('still to be invoiced (qty)','order unit'),
                               ('quantity in sku','stockkeeping unit')],'supplier/supplying plant',['order unit','stockkeeping unit']),
             'GR':('material',[('qty in unit of entry','unit of entry'),('quantity','base unit of measure'),('qty in order unit','order unit')],'batch',
                   ['unit of entry','base unit of measure','order unit'])}
def _unit_rows(name, rows, src):
    """Rows of an input sheet with every quantity in YOUR BOM unit (UNIT_SHEETS; header = first row) and the unit cells relabelled."""
    sp=UNIT_SHEETS.get(name)
    if not sp or not rows or not _unit_tables()['bom']: return rows
    hl=[_hnorm(x) for x in rows[0]]; ix=lambda k: hl.index(k) if k and k in hl else None
    ic,isp=ix(sp[0]),ix(sp[2]); qs=[(ix(a),ix(b)) for a,b in sp[1]]; qs=[(a,b) for a,b in qs if a is not None]; rl=[ix(k) for k in sp[3] if ix(k) is not None]
    if ic is None: return rows
    out=[rows[0]]
    for r in rows[1:]:
        c=str((r[ic] if r and ic<len(r) else None) or '').strip()
        b=bom_unit(c) if c else ''
        if not b: out.append(r); continue
        r=list(r)
        uu=lambda u: (r[u] if u is not None and u<len(r) else None) or None; sp_=r[isp] if isp is not None and isp<len(r) else None
        isn=lambda v: isinstance(v,(int,float)) and not isinstance(v,bool)
        if any(a<len(r) and isn(r[a]) and r[a] for a,_ in qs):   # Check: each line once, with its MAIN quantity (MB52 = Total Stock = Unrestricted + QI)
            a0,u0=qs[0]; to_bom(c,(r[a0] if a0<len(r) and isn(r[a0]) else 0.0),uu(u0),sp_,f'{src} {name}')
        for a,u in qs:                                   # a quantity with no known conversion becomes None = left out (answer 10)
            if a<len(r) and isn(r[a]) and r[a]: r[a]=to_bom(c,r[a],uu(u),sp_,f'{src} {name}',log=False)
        for j in rl:                                     # every quantity of the row is now in your BOM unit
            if j<len(r): r[j]=b
        out.append(tuple(r))
    return out
class _CachedWB:
    def __init__(self,path):
        self._wb=load_workbook(path,read_only=True,data_only=True); self.sheetnames=list(self._wb.sheetnames); self._s={}; self._path=path
    def __getitem__(self,name):
        if name not in self._s:
            if name not in self.sheetnames:          # a sheet missing from an input file: warn once and carry on with no rows
                print(f"NOTE: sheet '{name}' not found in {os.path.basename(self._path)} - that part is skipped")
                self._s[name]=_CachedSheet([],name); return self._s[name]
            rows=_rows_cache_get(self._path,name)
            if rows is None:
                rows=[tuple(r) for r in self._wb[name].iter_rows(values_only=True)]; _rows_cache_put(self._path,name,rows)
            if name=='In Transit' and rows and _is_mis_file(self._path): rows=ppc_it_rows(rows)   # IN TRANSIT = PPC In_Transit ONLY (user, 21-Nov-2026)
            if name=='In Transit' and rows: rows=normalize_in_transit(rows,self)     # fixed column layout + checked material code
            if name=='MB52' and rows and USER_MB52[0] and _is_mis_file(self._path): rows=user_mb52_rows(rows,self)   # stock from YOUR MB52
            if name in UNIT_SHEETS and rows:                 # quantities in YOUR BOM unit (user 26-Nov-2026)
                if _is_mis_file(self._path): rows=_unit_rows(name,rows,'MIS' if not (name=='MB52' and USER_MB52[0]) else 'SAP')
                elif PPC_SRC[0] and os.path.abspath(self._path)==os.path.abspath(PPC_SRC[0]): rows=_unit_rows(name,rows,'PPC')
            self._s[name]=_CachedSheet(rows,name)
        return self._s[name]
    def __contains__(self,name): return name in self.sheetnames
PPC_IT_NOTE=[False]
PPC_NODATE=[]      # read_ppc_in_transit: pending lines with no arrival date (not counted)
IT_COPIED={}       # map_in_transit: invoice -> text, lines with a copied description told apart by unit price (console NOTE once)
def _ppc_it_date(ata2, ata, eta, has_ata2, has_ata):
    """Plant arrival date of a PPC In_Transit line - ONE rule for WSA column T and 1_MRP: Final ATA2 (PPC's date) only when Final AtA
       holds a date too (PPC fills ATA2 with the file date when AtA is blank - user, 21-Nov-2026, answer 1c), else Revised ETA Port + 5;
       None = no arrival date: the line is NOT counted until a date is entered."""
    a_=_date_cell(ata) if has_ata else None
    if a_ is not None or not has_ata:
        a2_=_date_cell(ata2) if has_ata2 else None
        if a2_ is not None: return a2_
        if a_ is not None: return a_
    e_=_date_cell(eta)
    return (e_+pd.Timedelta(days=5)) if e_ is not None else None
PPC_IT_NEED=('ETD','ETA','BL#','AtA')
def _ppc_it_missing(etd, eta, eta_b, hbl, mbl, ata, has):
    """User, 21-Nov-2026: "a row counts as in transit only with enough information: ETD, ETA, shipping documents, ATA". ONE rule for
       WSA column T and 1_MRP: a pending line needs an ETD date, an ETA date (Revised ETA Port, else ETA as Per Booking), a BL#
       (HBLNo. or MBL: 6+ letters / digits with 4+ digits) and a Final AtA date. Returns the missing fields ([] = counted).
       has = the fields whose column exists on the sheet: a column missing from the sheet is not checked (WARNING once)."""
    m=[]
    if 'ETD' in has and _date_cell(etd) is None: m.append('ETD')
    if 'ETA' in has and _date_cell(eta) is None and _date_cell(eta_b) is None: m.append('ETA')
    if 'BL#' in has and not (_bl_set(hbl) or _bl_set(mbl)): m.append('BL#')
    if 'AtA' in has and _date_cell(ata) is None: m.append('AtA')
    return m
PPC_IT_COLWARN=[False]
def _ppc_it_has(iETD, iE, iEB, iHB, iMB, iA):
    """Which completeness fields can be checked on this sheet (column found); WARNING once for a missing one."""
    has=[k for k,ok in (('ETD',iETD is not None),('ETA',iE is not None or iEB is not None),('BL#',iHB is not None or iMB is not None),('AtA',iA is not None)) if ok]
    if len(has)<len(PPC_IT_NEED) and not PPC_IT_COLWARN[0]:
        PPC_IT_COLWARN[0]=True
        print(f"WARNING: PPC In_Transit has no column for {', '.join(k for k in PPC_IT_NEED if k not in has)} - that field is not checked "
              "(a line needs ETD + ETA + BL# + Final AtA to count as in transit)",flush=True)
    return has
def _ppc_nodate_txt(desc, inv, q, miss=None):
    try: q_=f"{float(q):,.0f}"
    except Exception: q_=str(q)
    return f"{' '.join(str(desc or '').split())[:40]} inv {inv} qty {q_}"+(f" (no {' / '.join(miss)})" if miss else '')
def ppc_it_rows(mis_rows):
    """User, 21-Nov-2026: "Script find in transit from PPC only. Ignore mis in transit". Every reader of the MIS 'In Transit' sheet
       (1_MRP projection, ALERT, dashboards, email, Check, In Transit Status, glass split by BL#) gets the rows of your PPC
       'In_Transit' sheet instead - the same columns as MIS. Same rules as WSA column T (`read_ppc_in_transit`): date 'Final AtA' =
       `_ppc_it_date` (Final ATA2 only when Final AtA has a date, else Revised ETA Port + 5; no date = line left out); 'Final Pending Qty.' blank = Final Qty. - PPC GRN. No PPC In_Transit = NO in transit (WARNING); MIS
       In Transit is never read for quantities or dates (its header row is kept only when the PPC sheet cannot be read)."""
    p_=PPC_SRC[0]; sn=None; raw=[]
    if p_:
        try:
            wb=load_workbook(p_,read_only=True,data_only=True); sn=_it_sheet(wb.sheetnames); empty=0
            if sn:
                for r in wb[sn].iter_rows(values_only=True):
                    if not any(v not in (None,'') for v in r):
                        empty+=1
                        if empty>(500 if raw else 2000): break
                        continue
                    empty=0; raw.append(tuple(r))
            wb.close()
        except Exception as ex_: print(f'WARNING: PPC In_Transit could not be read ({type(ex_).__name__}: {ex_})'); raw=[]
    hi=next((k for k,r in enumerate(raw[:30]) if sum(k_ in {_hnorm(v) for v in r} for k_ in IT_HDR_KEYS)>=2),None)
    if hi is None:
        if not PPC_IT_NOTE[0]:
            PPC_IT_NOTE[0]=True
            print('WARNING: NO IN TRANSIT this run - '+(f"{os.path.basename(p_)} has no readable In_Transit sheet" if p_ else 'no PPC file found')
                  +' (MIS In Transit is not used: in transit = PPC In_Transit only)',flush=True)
        return [tuple(mis_rows[0])] if mis_rows else []
    hdr=[str(h).strip() if h is not None else '' for h in raw[hi]]; hn=[_hnorm(h) for h in hdr]
    ix=lambda *ns: next((hn.index(n) for n in ns if n in hn),None)
    iA2,iA,iQ,iFQ,iG,iE=ix('final ata2'),ix('final ata'),ix('final pending qty.','final pending qty'),ix('final qty.'),ix('ppc grn'),ix('revised eta port')
    out=[tuple(hdr)]
    iDs,iInv=ix('description'),ix('invoice number'); nod_=[]
    iETD,iEB,iHB,iMB=ix('etd'),ix('eta as per booking'),ix('hblno.','hbl no.','hbl'),ix('mbl'); has_=_ppc_it_has(iETD,iE,iEB,iHB,iMB,iA)
    for r in raw[hi+1:]:
        r=list(r)+[None]*max(0,len(hdr)-len(r))
        g_=lambda i_: r[i_] if i_ is not None else None
        d_=_ppc_it_date(g_(iA2),g_(iA),g_(iE),iA2 is not None,iA is not None)    # = read_ppc_in_transit (WSA column T)
        mi_=_ppc_it_missing(g_(iETD),g_(iE),g_(iEB),g_(iHB),g_(iMB),g_(iA),has_)   # raw cells, before Final AtA is overwritten below
        if iA is not None: r[iA]=d_.to_pydatetime() if d_ is not None else None
        if iQ is not None:                                   # quantity: Final Pending Qty., else Final Qty. - PPC GRN (text numbers read; else 0)
            q_=_num_cell(r[iQ])
            if q_ is None:
                fq=_num_cell(r[iFQ]) if iFQ is not None else None; gr=_num_cell(r[iG]) if iG is not None else None
                q_=(fq or 0.0)-(gr or 0.0)
            r[iQ]=q_
        if iQ is not None and (_num_cell(r[iQ]) or 0)>0:                       # pending: needs ETD + ETA + BL# + Final AtA (user, 21-Nov-2026)
            if mi_ or (d_ is None and iA is not None):
                nod_.append(_ppc_nodate_txt(g_(iDs),g_(iInv),r[iQ],mi_ or ['arrival date'])); continue
        out.append(tuple(r))
    if not PPC_IT_NOTE[0]:
        PPC_IT_NOTE[0]=True
        print(f"IN TRANSIT = {os.path.basename(p_)} sheet {sn} ONLY ({len(out)-1} lines; MIS In Transit is not used - user, 21-Nov-2026)",flush=True)
        if nod_: print(f"NOTE: {len(nod_)} PPC In_Transit line(s) with a pending qty are NOT counted as in transit - a line needs ETD + ETA + BL# + Final AtA "
                       "(user, 21-Nov-2026): "+'; '.join(nod_[:10])+(' ...' if len(nod_)>10 else ''),flush=True)
    return out
# ---------- PPC In_Transit: column layout by header name, and the material code checked against the shipment description ----------
IT_CANON=['GID','Supplier Name','Invoice Number','Product Name','Description','HSN','PO','Qty','Unit Rate','Invoice value','BOOKING NO','LINE',
          'VESSEL / VOY','inco','OF','Cust','TPT','POL','POD',"40'HQ","20'GP",'Total Cont','STATUS','HBLNo.','MBL','Containers','ETD','ATD',
          'ETA as Per Booking','Revised ETA Port','warehouse','RM Code','Final Qty.','PPC GRN','Final Pending Qty.','Final AtA',
          'PO Wise Material Desc.','PO and Unit Price Key','True/False','Final RM Code','GR. Date']
IT_FIX=[]          # shipments whose MIS Final RM Code was replaced by the traced material (WSA PO + BL#, description, supplier; listed on Check)
IT_MISSING=[]      # In Transit headers not found in the MIS file this run (their column is read as blank - listed on Check)
_FRONT={'HTAR','ARC','FRONT'}; _BACK={'PRINTED','BACK'}
def _desc_keys(t):
    """Distinguishing features of a description: dimensions, thickness, HS frame height, roll length, front / back glass, GSM."""
    u=' '.join(str(t or '').upper().replace('\n',' ').split())
    u=re.sub(r'(\d+)\s+HS\b',r'\1HS',u)
    k={}
    k['hs']=set(re.findall(r'(\d+)HS',u))
    nums=set(re.findall(r'(?<![\d.])(\d{4})(?![\d.])',u)); k['dim']=nums
    th=set(re.findall(r'X\s*(\d+(?:\.\d+)?)\s*MM\b',u)) | set(re.findall(r'\b(\d+\.\d+)\s*MM\b',u))
    k['th']={str(float(x)) for x in th if float(x)<10}
    k['len']=set(re.findall(r'(\d+)\s*M\b(?!M)',u)) if 'TAPE' in u else set()     # roll length only tells tapes apart
    k['gsm']=set(re.findall(r'(\d+)\s*GSM',u))
    k['wd']=set()                                                       # cross-section of busbar / ribbon: '0.36*4 MM', '4 X 0.36MM', '5 X 0.36'
    for a_,b_ in re.findall(r'(\d+(?:\.\d+)?)\s*[X\*]\s*(\d+(?:\.\d+)?)\s*(?:MM|\b)',u):
        if float(a_)<20 and float(b_)<20: k['wd']|={str(float(a_)),str(float(b_))}
    w=set(re.split(r'[^A-Z0-9.]+',u))
    k['side']={'F'} if w&_FRONT else ({'B'} if w&_BACK else set())
    k['words']={x for x in w if len(x)>2 and not x.isdigit()}
    return k
def _desc_score(ship, cand):
    a,b=_desc_keys(ship),_desc_keys(cand); sc=0.0
    for f,wt in (('hs',3),('dim',2),('th',2),('len',2),('gsm',2),('side',3),('wd',3)):
        if a[f] and b[f]: sc+=wt*(len(a[f]&b[f])) - wt*(1 if not (a[f]&b[f]) else 0)*2
    sc+=0.2*len(a['words']&b['words'])
    return sc
def normalize_in_transit(rows, wbc):
    """1) Puts the In Transit columns in the usual order by their header names (MIS sometimes adds / removes a column).
       2) The material of each shipment is traced exactly like the PPC In Transit of the WSA / Inbound files (user rule, Nov 2026:
          'did you do the same with the MRP output'): your WSA sheet with this PO + BL# (HBLNo. / MBL) first, then description +
          supplier + RM Code against your BOM (map_in_transit). MIS 'Final RM Code' comes from the PO; it is replaced when another
          code clearly fits (e.g. EP304 0.50 mm on your EPE 420 sheet, 2272 glass on a 2376 PO, 35HS frame on a 30HS PO).
          Lines that are not BOM materials keep the MIS code. Every replacement is listed on the Check sheet (IT_FIX)."""
    hdr=[str(h).strip() if h is not None else '' for h in rows[0]]
    pos={h:i for i,h in enumerate(hdr) if h}
    IT_MISSING[:]=[h for h in IT_CANON if h not in pos]
    extra=[i for i,h in enumerate(hdr) if h and h not in IT_CANON]
    out=[tuple(IT_CANON+[hdr[i] for i in extra])]
    for r in rows[1:]:
        out.append(tuple([r[pos[h]] if (h in pos and pos[h]<len(r)) else None for h in IT_CANON]+[r[i] if i<len(r) else None for i in extra]))
    I_=IT_CANON.index; iR,iF,iD,iQ,iS,iPW,iPO,iHB,iMB,iInv,iUR=(I_(h) for h in ('RM Code','Final RM Code','Description','Final Pending Qty.','Supplier Name',
                                                                      'PO Wise Material Desc.','PO','HBLNo.','MBL','Invoice Number','Unit Rate'))
    up=lambda v: str(v or '').strip().upper()
    lines=[]
    for k,r in enumerate(out[1:],1):
        if not (r[iD] or r[iR] or r[iF]): continue
        lines.append(dict(k=k,final_code=up(r[iF]),rm_code=up(r[iR]),match='',po_codes=re.findall(r'\b(?:RM|PM|CS)[A-Z0-9]{6,}\b',up(r[iPW])),
                          description=' '.join(str(r[iD] or '').split()),qty=r[iQ] if isinstance(r[iQ],(int,float)) and not isinstance(r[iQ],bool) else 0,
                          po=r[iPO],bls=_bl_set(r[iHB])|_bl_set(r[iMB]),supplier=' '.join(str(r[iS] or '').split()),
                          invoice=r[iInv],price=_num_cell(r[iUR])))
    cat=in_transit_catalog(PPC_SRC[0],getattr(wbc,'_path',None) or MIS_PATH_[0])
    map_in_transit(lines,cat,dict(idx=wsa_ref_index(WSA_REFS[0]),alias={}) if WSA_REFS[0] else None)
    IT_FIX.clear(); fixed=list(out)
    for x in lines:
        c=x.get('code'); fc=str(out[x['k']][iF] or '').strip()
        if not c or c==x['final_code']: continue
        r=list(out[x['k']]); r[iF]=c; fixed[x['k']]=tuple(r)          # the traced material replaces the MIS code
        IT_FIX.append(dict(desc=x['description'][:80],shipped=c,mis=fc or '(blank)',qty=x['qty'],po=r[iPO],sup=r[iS],
                           cd_s=cat[c]['desc'],cd_m=(cat.get(up(fc)) or {}).get('desc',''),how=x.get('how') or ''))
    iFQ,iG=I_('Final Qty.'),I_('PPC GRN')
    for k in range(1,len(fixed)):                     # quantities (SAP unit of the traced material) -> YOUR BOM unit by the line's supplier
        r=list(fixed[k]); c=str(r[iF] or '').strip()
        if not c or not bom_unit(c): continue
        for j in (iQ,iFQ,iG):
            if isinstance(r[j],(int,float)) and not isinstance(r[j],bool) and r[j]: r[j]=to_bom(c,r[j],None,r[iS],'PPC In_Transit (1_MRP)',log=j==iQ)
        fixed[k]=tuple(r)
    return fixed

PPC_SRC=[None]     # PPC workbook (sheets 'Count' + 'BOM' + 'Plan'): the ONLY source of physical stock counts; BOM and plan (else MRP.xlsx)
PP_HDR=('Date','Production Line','Wp','Customer','Qty','MW','Plant','Module Model')
def _cu_clean(x): return ' '.join(str(x or '').split())
def read_revised_plan(ws):
    """PPC 'Revised Plan': blocks side by side, one per production line (line name above the block, then Wp / Customer / Qty / MW),
       one row per date. A month block further down (new header row) is read the same way.
       Returns [(date, line, wp, customer, qty, mw)] for every cell with a Qty (0 included = line not running that day)."""
    rows=[list(r) for r in ws.iter_rows(values_only=True)]
    out=[]; blocks=None; dcol=None
    for i,r in enumerate(rows):
        hl=[' '.join(str(x or '').split()).lower() for x in r]
        if 'customer' in hl and 'qty' in hl:                      # header row: one block per 'Customer' column
            above=rows[i-1] if i>0 else []; blocks=[]
            for j,h in enumerate(hl):
                if h!='customer': continue
                q=next((k for k in range(j+1,len(hl)) if hl[k]=='qty'),None)
                if q is None: continue
                w=j-1 if j>0 and hl[j-1] in ('wp','watt') else None
                m=q+1 if q+1<len(hl) and hl[q+1]=='mw' else None
                st=w if w is not None else j; ln=None
                for k in range(st,-1,-1):                          # line name: the nearest text above the block, at or left of its start
                    v=above[k] if k<len(above) else None
                    if isinstance(v,str) and v.strip(): ln=_cu_clean(v); break
                    if blocks and k<=blocks[-1]['end']: break             # do not take the previous block's name
                blocks.append(dict(line=ln,wp=w,cu=j,qty=q,mw=m,end=(m if m is not None else q)))
            dcol=None; continue
        if not blocks: continue
        if dcol is None:
            dcol=next((k for k,v in enumerate(r) if isinstance(v,(dt.datetime,dt.date))),None)
        d=r[dcol] if dcol is not None and dcol<len(r) else None
        if not isinstance(d,(dt.datetime,dt.date)): continue
        for b in blocks:
            if not b['line']: continue
            q=r[b['qty']] if b['qty']<len(r) else None
            if not isinstance(q,(int,float)) or isinstance(q,bool): continue
            g=lambda k: (r[k] if k is not None and k<len(r) else None)
            out.append((pd.Timestamp(d).normalize(),b['line'],g(b['wp']),_cu_clean(g(b['cu'])),float(q),g(b['mw'])))
    return out
PLAN_NOTE=[None]
def _has_ppc_plan(names):
    """PPC workbook has a module plan: sheet 'Plan' or 'WSA Plan' (PPC of 20-Nov-2026 has only 'WSA Plan' + 'OEM Plan' + 'Revised Plan')."""
    return any(str(n).strip().lower() in ('plan','wsa plan') for n in (names or []))
def ppc_plan_sheet(ppc):
    """Module plan = PPC 'Plan' with PPC 'Revised Plan' on top: for every date + production line in the Revised Plan, its rows
       replace the 'Plan' rows of that date + line; all other dates and lines stay as on 'Plan'.
       Revised Plan customer names are matched to the 'Plan' names (same line first) so the BOM and MIS links keep working.
       Returned in the Production Plan layout (Date, Production Line, Wp, Customer, Qty, MW, Plant, Module Model)."""
    import difflib
    hit=getattr(ppc,'_pp_merged',None)
    if hit is not None: return hit
    base=[r for r in pp_rows_std(ppc['Plan']) if r and isinstance(r[0],(dt.datetime,dt.date))] if 'Plan' in ppc.sheetnames else []
    wsa_=next((n for n in ppc.sheetnames if n.strip().lower()=='wsa plan'),None)
    if wsa_:                                                       # user, 19-Nov-2026: "WSA plan is correct. Use WSA plan" - lines 1-4 from 'WSA Plan' (one column per line); Wp / model from the 'Plan' row of the same day + line + customer
        info={(pd.Timestamp(r[0]).normalize(),_line_key(r[1]),_cu_clean(r[3])):r for r in base}
        wb_rows=[]; rows_=[list(r) for r in ppc[wsa_].iter_rows(values_only=True)]
        hi=next((i for i,r in enumerate(rows_[:10]) if r and any(str(x or '').strip().lower()=='date' for x in r) and any(str(x or '').strip().lower().startswith('usa line') for x in r)),None)
        if hi is not None:
            hl=[' '.join(str(x or '').split()).lower() for x in rows_[hi]]; jd=hl.index('date'); jc=hl.index('customer') if 'customer' in hl else None
            jw=next((j for j,h in enumerate(hl) if h.startswith('watt')),None); jwp=hl.index('wp') if 'wp' in hl else None
            jm=hl.index('module model') if 'module model' in hl else None     # no 'Plan' sheet (PPC 20-Nov-2026): Wp / model from WSA Plan itself
            lj=[(j,' '.join(str(rows_[hi][j]).split())) for j,h in enumerate(hl) if h.startswith('usa line')]
            for r in rows_[hi+1:]:
                if not r or jd>=len(r) or not isinstance(r[jd],(dt.datetime,dt.date)) or jc is None: continue
                d_=pd.Timestamp(r[jd]).normalize(); cu_=_cu_clean(r[jc])
                for j,ln in lj:
                    q=r[j] if j<len(r) else None
                    if not isinstance(q,(int,float)) or isinstance(q,bool) or not q: continue
                    o=info.get((d_,_line_key(ln),cu_)); wp=o[2] if o else None
                    if wp is None and jwp is not None and jwp<len(r) and isinstance(r[jwp],(int,float)) and not isinstance(r[jwp],bool): wp=r[jwp]
                    if wp is None and jw is not None and jw<len(r):
                        m_=re.search(r'\d+',str(r[jw] or '')); wp=int(m_.group()) if m_ else None
                    wb_rows.append((d_.to_pydatetime(),ln,wp,cu_,float(q),(float(q)*wp/1e6 if isinstance(wp,(int,float)) else None),(o[6] if o else 'USA'),(o[7] if o else (r[jm] if jm is not None and jm<len(r) else None))))
            if wb_rows: base=wb_rows
    # PPC of 19-Nov-2026: 'Plan' = USA lines 1-4 only, the OEM plan is its own sheet 'OEM Plan' (same columns); 'WSA Plan' = a combined
    # print of lines 1-4 (not read). Old PPC: OEM rows were inside 'Plan'. OEM rows without a quantity are left out.
    oem_=next((n for n in ppc.sheetnames if n.strip().lower()=='oem plan'),None)
    OEM_PLAN_CUST.clear()
    if oem_:
        orows_=[r for r in pp_rows_std(ppc[oem_]) if r and isinstance(r[0],(dt.datetime,dt.date)) and isinstance(r[4],(int,float)) and not isinstance(r[4],bool)]
        own_={_cu_clean(r[3]).lower() for r in base if r and r[4]}            # customers on the plant's own plan (WSA Plan / Plan)
        both_=sorted({_cu_clean(r[3]) for r in orows_ if r[4] and _cu_clean(r[3]).lower() in own_})
        OEM_PLAN_CUST.update(c_ for c_ in {_cu_clean(r[3]) for r in orows_ if r[4]} if c_ and c_ not in both_)   # every 'OEM Plan' customer = OEM (user 28-Nov-2026)
        if both_: print(f"NOTE: customer(s) on both PPC 'WSA Plan' and 'OEM Plan' are NOT treated as OEM (their plant lines need the stock): {', '.join(both_)}",flush=True)
        base+=orows_
    rev=read_revised_plan(ppc['Revised Plan']) if 'Revised Plan' in ppc.sheetnames else []
    names={}; models={}; plants={}
    for r in base:
        cu=_cu_clean(r[3])
        if not cu or not isinstance(r[4],(int,float)) or not r[4]: continue
        names.setdefault(cu,set()).add(_line_key(r[1]))
        mdl=str(r[7] or '').strip()
        if mdl.upper().startswith('BIN'): models.setdefault(cu,{}); models[cu][mdl]=models[cu].get(mdl,0)+r[4]
        if r[6]: plants.setdefault(_line_key(r[1]),r[6])
    ren={}
    def canon(cu,lk):
        if not cu: return cu
        n=cu.lower()
        for c in names:
            if c.lower()==n: return c
        best,sc=None,0.0
        for c in names:
            s=difflib.SequenceMatcher(None,n,c.lower()).ratio()
            if c.lower().startswith(n) or n.startswith(c.lower()): s=max(s,0.9)
            if lk in names[c]: s+=0.1
            if s>sc: best,sc=c,s
        if best and sc>=0.75: ren[cu]=best; return best
        return cu
    rkeys={(d,_line_key(ln)) for d,ln,*_ in rev}
    merged=[r for r in base if (pd.Timestamp(r[0]).normalize(),_line_key(r[1])) not in rkeys]
    for d,ln,wp,cu,q,mw in rev:
        c=canon(cu,_line_key(ln)); mdl=max(models[c].items(),key=lambda kv:kv[1])[0] if models.get(c) else None
        merged.append((d.to_pydatetime(),ln,wp,c,q,mw,plants.get(_line_key(ln),'USA'),mdl))
    merged.sort(key=lambda r:(pd.Timestamp(r[0]),str(r[1])))
    if rev:
        ds=sorted({d for d,_ in rkeys}); lns=sorted({ln for _,ln,*_ in rev})
        PLAN_NOTE[0]=(f"Module plan from PPC: "+("'WSA Plan'" if wsa_ else "'Plan'")+(" + 'OEM Plan'" if oem_ else '')+f" + 'Revised Plan' ({ds[0]:%d-%b} to {ds[-1]:%d-%b}, {', '.join(lns)} replaced)"
                      +(' - Revised Plan names matched: '+'; '.join(f"'{a}' -> '{b}'" for a,b in sorted(ren.items())) if ren else ''))
    else: PLAN_NOTE[0]="Module plan from PPC "+("'WSA Plan'" if wsa_ else "'Plan'")+(" + 'OEM Plan'" if oem_ else '')+" (no 'Revised Plan' sheet)"
    print(PLAN_NOTE[0],flush=True)
    sh=_CachedSheet([PP_HDR]+[tuple(r) for r in merged],'Production Plan')
    try: ppc._pp_merged=sh
    except Exception as _e_: _swallow(503,_e_); pass
    return sh

class _RedirWB:
    """Your MRP.xlsx as the script sees it: 'BOM' and 'Stock Count' come from the PPC workbook ('BOM', 'Count'), all else from MRP.xlsx.
       When the PPC workbook has a 'Plan' sheet, 'Production Plan' = PPC 'Plan' + 'Revised Plan' (see ppc_plan_sheet)."""
    MAP={'BOM':'BOM','Stock Count':'Count'}
    def __init__(self,base,ppc): self._b=base; self._p=ppc
    def _has_plan(self): return _has_ppc_plan(self._p.sheetnames)
    @property
    def sheetnames(self): return list(dict.fromkeys(list(self._b.sheetnames)+list(self.MAP)+(['Production Plan'] if self._has_plan() else [])))
    def __contains__(self,n): return n in self.sheetnames
    def __getitem__(self,n):
        if n=='Production Plan' and self._has_plan(): return ppc_plan_sheet(self._p)
        if n=='Stock Count': return _count_main_table(self._p)
        return self._p[self.MAP[n]] if n in self.MAP else self._b[n]
    def __getattr__(self,a): return getattr(self._b,a)
_COUNT_MAIN={}
def _count_main_table(ppc):
    """PPC 'Count' = its MAIN table only (checker finding, 12-Nov-2026): the planner's sheet has a second small list further down
       (row 170: 'Material / Description / Stock', 24 packaging rows in STOCK units - rolls, not the BOM unit). Those rows were read as
       counts with Quantity = 1 and replaced the main rows: PET strapping 016 showed 1,399 KG instead of 1,512 rolls x 21.8 = 32,962 KG.
       Everything from a second header row (column A = 'Material' again) down is left out; a console NOTE says how many rows."""
    k=id(ppc)
    if k not in _COUNT_MAIN:
        sh=ppc['Count']; rows=[tuple(r) for r in sh.iter_rows(values_only=True)]
        isH=lambda r: bool(r) and str(r[0] or '').strip()=='Material'
        h1=next((i for i,r in enumerate(rows) if isH(r)),None)
        h2=next((i for i,r in enumerate(rows) if h1 is not None and i>h1 and isH(r)),None)
        if h2 is not None:
            n_=sum(1 for r in rows[h2+1:] if r and r[0] not in (None,''))
            print(f"NOTE: PPC Count sheet: the second list from row {h2+1} ({' / '.join(str(x) for x in rows[h2] if x not in (None,''))}; {n_} rows) is NOT read as counts - only the main table (rows {h1+1}-{h2}) is. Put a new count in the main table's Physical Count Qty / Counting Date columns.",flush=True)
            rows=rows[:h2]
        if h1 is not None:                                               # user, 19-Nov-2026 (new Count layout, no 'Physical Count Qty' column):
            hl_=[' '.join(str(x or '').split()).lower() for x in rows[h1]]   # "must read Physical Count Qty_BOM and Count date" - the count = Physical
            ix_=lambda k: hl_.index(k) if k in hl_ else None                 # Count Qty_BOM (BOM unit; text 'no counting' = not counted), dated by 'Count date'.
            iqb,iqc=ix_('physical count qty_bom'),next((ix_(k) for k in ('count closing stock (bun)','count closing stock','counting qty') if k in hl_),None)
            if 'physical count qty' not in hl_ and 'today stock qty' not in hl_ and iqb is not None and iqc is not None:
                iqf=ix_('quantity')       # count column = Qty_BOM / Quantity (stock unit): the MRP unit stays as SAP / MIS give it (RTV 006 plans in DR:
                isn=lambda v: isinstance(v,(int,float)) and not isinstance(v,bool)   # 39,460 KG / 270 = 146 DR); CS / PM keep using Qty_BOM itself
                r_=list(rows[h1]); r_[iqc]='Physical Count Qty'; rows[h1]=tuple(r_)
                for i_ in range(h1+1,len(rows)):
                    r_=list(rows[i_])
                    if not r_ or r_[0] in (None,'') or iqb>=len(r_): continue
                    q_=r_[iqb]; f_=r_[iqf] if iqf is not None and iqf<len(r_) and isn(r_[iqf]) and r_[iqf] else 1.0
                    r_[iqc]=float(q_)/float(f_) if isn(q_) else None
                    rows[i_]=tuple(r_)
                ik,idt_=ix_('key'),ix_('count date')                 # checker 19-Nov-2026: RMRIBRW000036 listed twice (rows 16 + 17, same Key 'RMRIBRW000036|',
                if ik is not None:                                  # same SAP / count / date, units M2 + KG) was added up = count doubled (21,923 instead of 10,962).
                    seen_,dup_,keep_=set(),[],rows[:h1+1]           # A row with the same Key + count + count date as an earlier row = a duplicate, counted once.
                    for i_ in range(h1+1,len(rows)):
                        r_=rows[i_]; g_=lambda j: r_[j] if j is not None and j<len(r_) else None
                        k_=(str(g_(ik) or '').strip().upper(),g_(iqb),str(g_(idt_)))
                        if r_ and r_[0] not in (None,'') and k_[0] and isn(k_[1]) and k_ in seen_: dup_.append(f"row {i_+1} {r_[0]} ({k_[1]:,.1f})"); continue
                        seen_.add(k_); keep_.append(r_)
                    if dup_ and not _COUNT_MAIN.get('_dup_said'):
                        _COUNT_MAIN['_dup_said']=True
                        print(f"NOTE: PPC Count sheet: {len(dup_)} duplicate row(s) (same Key, Physical Count Qty_BOM and Count date as an earlier row) counted ONCE: {'; '.join(dup_)}. Delete the duplicate in the PPC file.",flush=True)
                    if dup_: rows=keep_
        _COUNT_MAIN[k]=(ppc,_CachedSheet(rows,'Count'))
    return _COUNT_MAIN[k][1]
def find_ppc(folders):
    """Newest workbook with 'Count' and 'BOM' sheets (a file name with 'PPC' wins), not an 1_MRP / Inbound output."""
    best=[]
    for fo in folders:
        for f in glob.glob(os.path.join(fo,'*.xlsx')):
            bn=os.path.basename(f)
            if bn.startswith('~$') or is_output_file(bn): continue
            try: n=set(_sheetnames_raw(f))
            except Exception as _e_: _swallow(527,_e_); continue
            if {'Count','BOM'}<=n: best.append(('ppc' in bn.lower(),os.path.getmtime(f),f))
    return max(best)[2] if best else None
def wb_ro(path):
    """wb_ro(path), parsed once and cached; re-read if the file changes. Your MRP.xlsx: BOM / Stock Count from the PPC workbook."""
    base=_wb_ro_raw(old_names_copy(path))
    try:
        um_={os.path.abspath(p_) for p_ in (globals().get('USER_MRP'),(SETTINGS_SRC[0] if SETTINGS_SRC else None)) if p_}
        if PPC_SRC[0] and os.path.abspath(path) in um_:
            return _RedirWB(base,_wb_ro_raw(PPC_SRC[0]))
    except NameError: pass
    return base
def _wb_ro_raw(path):
    key=os.path.abspath(path)
    try: mt=os.path.getmtime(key)
    except OSError: mt=None
    hit=_WB_CACHE.get(key)
    if hit is None or hit[0]!=mt: hit=(mt,_CachedWB(path)); _WB_CACHE[key]=hit
    return hit[1]

RUN_NOTES=[]
class _CaptureDone(Exception): pass    # step 1 of the WSA alignment stops build_new_mrp after align_capture()
_QUIET=[False]                          # step 1 of the WSA alignment: only errors on the console (step 2 prints everything)
import builtins as _bi
def print(*a, **k):                     # every NOTE / WARNING / ERROR line is also kept (RUN_NOTES)
    m=' '.join(str(x) for x in a)
    if m.startswith(('NOTE','WARNING','ERROR')): RUN_NOTES.append(m)
    if _QUIET[0] and not m.lstrip().startswith('ERROR'): return
    _bi.print(*a, **k)
def die(m): print('ERROR:',m); sys.exit(1)
ARGS=sys.argv[1:]; PREV=None; MRP=None; MIS=None
NO_MRP='--no-mrp' in ARGS
if NO_MRP: ARGS.remove('--no-mrp')
KEEP_OLD='--keep-old' in ARGS                  # keep the older 1_MRP_<date>.xlsx files (default: they are deleted once the new one is saved)
if KEEP_OLD: ARGS.remove('--keep-old')
NO_ALIGN='--no-align' in ARGS                   # default: the WSA planning columns are rebuilt (aligned copy of the WSA file) on every run
if NO_ALIGN: ARGS.remove('--no-align')
if '--align-wsa' in ARGS: ARGS.remove('--align-wsa')  # older command line: aligning is the default now
ALIGN_WSA=not NO_ALIGN
for flag in ('--prev','--mrp','--mrp-out','--mis'):
    if flag in ARGS:
        i=ARGS.index(flag); val=ARGS[i+1] if i+1<len(ARGS) else die(f'{flag} needs a file'); del ARGS[i:i+2]
        if flag=='--prev': PREV=val
        elif flag=='--mrp': MRP=val
        elif flag=='--mis': MIS=val
        else: MRP_OUT=val
def _release(path):
    """Close the cached read-only handles of a workbook (typed_rows / wb_ro) so the file can be deleted or replaced."""
    k=os.path.abspath(path)
    for p_ in [p_ for p_ in list(_PARTS) if os.path.abspath(p_)==k]:
        try: _PARTS.pop(p_)[0].close()
        except Exception as _e_: _swallow(576,_e_); pass
    hit=_WB_CACHE.pop(k,None)
    if hit is not None:
        try: hit[1]._wb.close()
        except Exception as _e_: _swallow(580,_e_); pass
    import gc; gc.collect()
def _sheetnames(p):
    try: return wb_ro(p).sheetnames
    except Exception as _e_: _swallow(584,_e_); return []
def _sheetnames_raw(p):
    """Sheet names straight from xl/workbook.xml (no workbook is opened) - used when scanning a folder for the right file."""
    import zipfile, html
    try:
        with zipfile.ZipFile(p) as z_: x_=z_.read('xl/workbook.xml').decode('utf-8')
        return [html.unescape(n_) for n_ in re.findall(r'<(?:\w+:)?sheet\b[^>]*?\bname="([^"]*)"',x_)]
    except Exception as _e_: _swallow(591,_e_); return []
_WSA_MARK={'Total_Plan_Req','PPC_Plan_Req'}     # a WSA daily file: the planner's (Total_Plan_Req) or an aligned one (PPC_Plan_Req)
if not ARGS:      # double-click / no file names: newest WSA daily file (Total_Plan_Req or PPC_Plan_Req sheet) in this folder or the script's folder
    _cands=[]
    for _f in dict.fromkeys([os.getcwd(),os.path.dirname(os.path.abspath(sys.argv[0]))]):
        for _n in os.listdir(_f):
            _nl=_n.lower()
            if _nl.endswith('.xlsx') and not _n.startswith('~$') and not _nl.endswith(('.sorting.xlsx','.tmp.keep')) and not _nl.startswith('mrp') and not is_output_file(_n) and 'mis' not in _nl and 'dashboard' not in _nl:
                _cands.append(os.path.join(_f,_n))
    _cands=[p for p in sorted(_cands,key=os.path.getmtime,reverse=True) if _WSA_MARK&set(_sheetnames_raw(p))]
    _own=[p for p in _cands if not re.search(r'\(aligned \d{4}-\d{2}-\d{2}\)',os.path.basename(p))]   # your own WSA file first (the script never writes it)
    if _own: _cands=_own+[p for p in _cands if p not in _own]
    if not _cands: die('No WSA daily file found here. Usage: python update_inbound.py <WSA daily file.xlsx> [MRP.xlsx] [MIS RM Dashboard.xlsx]')
    ARGS=[_cands[0]]; print(f'WSA daily file: {_cands[0]}')
def _is_mrp(p):
    n=_sheetnames(p)            # the MRP workbook has a By Week sheet; an Inbound output also has one, but it has an Inbound sheet too
    return 'By Week' in n and 'Inbound' not in n and not (_WSA_MARK&set(n))
def _is_mis(p):
    n=_sheetnames(p); return 'MB52' in n and ('Daily Dashboard' in n or 'Standard BOM' in n)
for a_ in list(ARGS):                     # an MIS RM Dashboard given in any position
    if MIS is None and a_.lower().endswith('.xlsx') and os.path.exists(a_) and _is_mis(a_): MIS=a_; ARGS.remove(a_)
if not ARGS: die('Usage: python update_inbound.py <daily_file.xlsx> [MRP.xlsx] [MIS_RM_Dashboard.xlsx] [output.xlsx] [--prev ...]')
if len(ARGS)>=2 and os.path.exists(ARGS[0]) and os.path.exists(ARGS[1]):      # accept the daily file and the MRP workbook in either order
    n0,n1=_sheetnames(ARGS[0]),_sheetnames(ARGS[1])
    if not (_WSA_MARK&set(n0)) and _is_mrp(ARGS[0]) and _WSA_MARK&set(n1): ARGS[0],ARGS[1]=ARGS[1],ARGS[0]
SRC=ARGS[0]
if not os.path.exists(SRC): die(f'File not found: {SRC}')
OUT=None
def _is_user_mrp(p):
    n=_sheetnames(p)                     # Nov 2026: the planner's MRP.xlsx may hold only Settings (plan + BOM + counts from the PPC file)
    return ('Settings' in n or ('Production Plan' in n and 'BOM' in n)) and not is_output_file(os.path.basename(p))
USER_MRP=MRP if (MRP and os.path.exists(MRP)) else None     # your MRP workbook (BOM, Stock Count, Production Plan, Customer Supplier): read only
MRP=None                                                      # the old By Week / MRP_<date> update is retired - 1_MRP is the output
if len(ARGS)>1:
    if os.path.exists(ARGS[1]) and USER_MRP is None and _is_user_mrp(ARGS[1]): USER_MRP=ARGS[1]
    else: OUT=ARGS[1]
if len(ARGS)>2 and OUT is None: OUT=ARGS[2]
OUTDIR=os.path.dirname(os.path.abspath(SRC))
def _find_user_mrp():
    """MRP.xlsx next to the daily file (or in the current folder); else the newest MRP*.xlsx with BOM + Production Plan sheets."""
    folders=[OUTDIR]+([os.path.abspath(os.getcwd())] if os.path.abspath(os.getcwd())!=OUTDIR else [])
    for f in folders:
        p=os.path.join(f,'MRP.xlsx')
        if os.path.exists(p) and _is_user_mrp(p): return p
    cands=[os.path.join(f,n) for f in folders for n in os.listdir(f)
           if n.lower().startswith('mrp') and n.lower().endswith('.xlsx') and not n.startswith('~$') and not is_output_file(n)]
    for p in sorted(cands,key=os.path.getmtime,reverse=True):
        if _is_user_mrp(p): return p
    return None
if USER_MRP is None and not NO_MRP: USER_MRP=_find_user_mrp()
if USER_MRP: print(f'Your MRP workbook (read only): {USER_MRP}')
else: print('NOTE: no MRP.xlsx (with BOM and Production Plan sheets) found - 1_MRP uses the MIS plan only.')
_PPC_DIRS=list(dict.fromkeys([os.path.dirname(os.path.abspath(USER_MRP)) if USER_MRP else os.getcwd(),OUTDIR,os.path.abspath(os.getcwd()),os.path.dirname(os.path.abspath(__file__))]))
PPC_SRC[0]=find_ppc(_PPC_DIRS)
if PPC_SRC[0]: print(f'BOM and Stock Count read from: {os.path.basename(PPC_SRC[0])} (sheets BOM, Count'+(", module plan from Plan / WSA Plan + Revised Plan)" if _has_ppc_plan(_sheetnames(PPC_SRC[0])) else ')'))
else: print('WARNING: no PPC workbook (sheets Count + BOM) in the folder - physical stock counts come ONLY from the PPC Count sheet, so this run uses MB52 stock for every material; BOM is read from MRP.xlsx')
USER_MB52_SHEET[0]=ppc_mb52_sheet(PPC_SRC[0]) if PPC_SRC[0] else None      # user 20-Nov-2026: "Find MB52 sheet from PPC from now on"
if USER_MB52_SHEET[0]: USER_MB52[0]=PPC_SRC[0]
else:
    USER_MB52[0]=find_user_mb52(list(dict.fromkeys([os.path.dirname(os.path.abspath(USER_MRP)) if USER_MRP else os.getcwd(),OUTDIR,os.path.abspath(os.getcwd()),os.path.dirname(os.path.abspath(__file__))])))
    if PPC_SRC[0]: print(f"NOTE: no 'MB52' sheet (Plant / Material / Unrestricted) in {os.path.basename(PPC_SRC[0])} - "+('stock from '+os.path.basename(USER_MB52[0]) if USER_MB52[0] else 'stock is read from the MIS MB52 sheet'))
if not USER_MB52[0]: print('NOTE: no MB52 found (PPC sheet MB52 or MB52.xlsx) - stock is read from the MIS MB52 sheet')
_BASE_D0=[None]
def base_d0():
    """The run's day 1 (user rule, Nov 2026: "If you have MB52 Nov 09, then calc everything from Nov 09; the stock count is also
       considered counted on Nov 09"): the date of your MB52.xlsx when it is older than today, else today. Projections, the WSA
       daily's Balance row, the ALERT / dashboards and late shipments all start on this day; a PPC count dated on / after it counts as
       taken on it (no roll-forward), an older count is rolled forward to it. File names keep the run date."""
    if _BASE_D0[0] is None:
        t_=pd.Timestamp(dt.date.today()).normalize(); b_=t_
        try:
            if USER_MB52[0]:
                m_=file_data_date(USER_MB52[0])
                if m_ is not None and m_<t_: b_=pd.Timestamp(m_).normalize()
        except Exception as _e_: _swallow(0,_e_)
        _BASE_D0[0]=b_
        if b_<t_ and not _QUIET[0]: print(f"BASE DATE {b_:%d-%b-%Y}: your MB52 ({os.path.basename(USER_MB52[0])}) is dated {b_:%d-%b} - everything is calculated from {b_:%d-%b} (stock, counts, plan, in transit); file names keep today's date")
    return _BASE_D0[0]
try: import lxml.etree
except ImportError: die("This script needs one more add-on. Run this once, then run the script again:\n       pip install lxml")
if MIS is None:   # newest MIS*Dashboard*.xlsx next to the daily file (or in the current folder)
    cands=[]
    for f in {OUTDIR,os.path.abspath(os.getcwd())}:
        for n in os.listdir(f):
            nl=n.lower()
            if nl.endswith('.xlsx') and not n.startswith('~$') and ('mis' in nl or 'dashboard' in nl): cands.append(os.path.join(f,n))
    for p in sorted(cands,key=os.path.getmtime,reverse=True):
        if _is_mis(p): MIS=p; print(f'MIS RM Dashboard found: {p}'); break
if MIS and not os.path.exists(MIS): die(f'MIS RM Dashboard not found: {MIS}')
if not MIS: print('NOTE: no MIS RM Dashboard found (needs MB52 and Standard BOM sheets) - 1_MRP is not built, only the Inbound file.')
OUT=OUT or os.path.join(OUTDIR,outfile_name('inb','xlsx'))

CODE=re.compile(r'^(JW)?RM[A-Z]{3,5}\d+$')
LABELS={('Tapes',0):'Tapes - Cell Fixation',('Tapes',1):'Tapes - Perforated 500m',('Potting Material',0):'Potting Material - Part A',('Potting Material',1):'Potting Material - Part B'}
SKIP={'Total_Plan_Req','Stock'}
def _num(x): return float(x) if isinstance(x,(int,float)) and not isinstance(x,bool) else None
def cv(rows,r,c):
    if r<1 or c<1 or r>len(rows): return None
    row=rows[r-1]; return row[c-1] if c<=len(row) else None

_PARTS={}
def typed_rows(path, sheet, col):
    """Rows whose cell in column 'col' holds a typed number (not a formula), read straight from the sheet XML - much faster than
       loading the whole workbook a second time just to see which balances are typed."""
    import zipfile
    from lxml import etree
    M='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
    if path not in _PARTS:                 # sheet name -> xml part, read once per file
        z=zipfile.ZipFile(path); wbx=etree.fromstring(z.read('xl/workbook.xml')); rels=etree.fromstring(z.read('xl/_rels/workbook.xml.rels'))
        tg={r.get('Id'):r.get('Target') for r in rels}
        _PARTS[path]=(z,{sh.get('name'):tg.get(sh.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')) for sh in wbx.find(M+'sheets')})
    z,pm=_PARTS[path]; tgt=pm.get(sheet)
    if tgt is None: return set()
    part=tgt.lstrip('/') if tgt.startswith('/') else 'xl/'+tgt
    letter=L(col); out=set()
    for _,c in etree.iterparse(z.open(part),tag=M+'c'):
        ref=c.get('r') or ''
        if ref.rstrip('0123456789')==letter and c.find(M+'f') is None and c.find(M+'v') is not None and c.get('t') in (None,'n'):
            out.add(int(ref[len(letter):]))
        c.clear()
    return out

def extract_fast(path, plan_end):
    print(f'  [{time.time()-_T0:5.0f} s] reading the WSA planning file ...',flush=True)
    wv=load_workbook(path,data_only=True,read_only=True)
    groups=[]; skipped=[]
    for sh in wv.sheetnames:
        if sh in SKIP or sh==ARCHIVE_SHEET: continue
        ws=wv[sh]
        if ws.sheet_state!='visible': continue
        head=[list(r) for r in ws.iter_rows(min_row=1,max_row=12,values_only=True)]
        starts=[(r+1,c+1) for r,row in enumerate(head) for c,x in enumerate(row[:200]) if str(x).strip().startswith('Required Plant Date')]
        if not starts: continue
        rv=[list(r) for r in ws.iter_rows(values_only=True)]
        maxc=max(len(r) for r in rv)
        starts.sort(key=lambda x:x[1])
        for bi,(hr,c0) in enumerate(starts):
            c1=(starts[bi+1][1]-1) if bi+1<len(starts) else maxc
            hdr={}
            for c in range(c0,c1+1):
                x=str(cv(rv,hr,c) or '').replace('\n',' ').strip().lower()
                if x: hdr.setdefault(x,c)
            def col(*ps):
                for p in ps:
                    for x,c in hdr.items():
                        if x.startswith(p): return c
                return None
            FIELDS=[('balance',lambda h:h.startswith('balance')),('planned',lambda h:h.startswith('planned modules per day')),('demand',lambda h:h.startswith('demand')),
                    ('actmod',lambda h:h.startswith('actual modules per day')),('actual',lambda h:h=='actual' or h.startswith('actual consum')),('mw',lambda h:h.startswith('mw per day')),
                    ('actmw',lambda h:h.startswith('actual mw per day')),('plantdays',lambda h:h.startswith('plant days')),('plantport',lambda h:h.startswith('plant + port') and '(ppc' not in h),('delivered',lambda h:h.startswith('delivered')),
                    ('ppcport',lambda h:h.startswith('plant + port days (ppc')),('ppcit',lambda h:h.startswith('ppc in transit') and 'supplier' not in h),
                    ('vlock',lambda h:h.startswith('vendor rule: not usable'))]
            fcols={k:next((c for h,c in hdr.items() if fn(h)),None) for k,fn in FIELDS}
            def info(r): return {k:(_num(cv(rv,r,c)) if c else None) for k,c in fcols.items()}
            cB=col('balance'); cDl=col('delivered'); cS=col('scheduled'); cPO=col('po'); cSh=col('shipper'); cRem=col('remarks'); cBk=col('booking approv'); cRd=col('readiness'); cETD=col('etd'); cBL=col('bl'); cVs=col('vessel')
            if cB is None or cS is None: skipped.append(f'{sh} (no Balance / Scheduled Delivery column)'); continue
            codes=set(); uom=''; mname=''
            for r in range(1,hr):
                for c in range(c0,c1+1):
                    x=cv(rv,r,c)
                    if isinstance(x,str):
                        if CODE.match(x.strip()): codes.add(x.strip())
                        if x.strip().startswith('Unit of Measurement'): uom=str(cv(rv,r,c+2) or '').strip()
                        if x.strip().startswith('Supplier & Material'): mname=str(cv(rv,r,c+2) or '').strip()
            uom,_=wsa_block_to_bom(rv,None,hr,{'bal':cB,'dem':fcols['demand'],'act':fcols['actual'],'dlv':cDl,'sch':cS,'ppcit':fcols['ppcit'],'vlock':fcols['vlock'],'ship':cSh},
                                   sorted(codes),uom,sh,log=False)      # YOUR BOM unit (= 2_WSA `_wsa_read`; listed on Check once, from there)
            rows=[(r,cv(rv,r,c0)) for r in range(hr+1,len(rv)+1) if isinstance(cv(rv,r,c0),dt.datetime)]
            if not rows: skipped.append(f'{sh} (no dated rows)'); continue
            last=None; typed=typed_rows(path,sh,cB)          # balances typed by the planner (not formulas) = real stock dates
            for r,d in rows:
                if r in typed and isinstance(cv(rv,r,cB),(int,float)) and not isinstance(cv(rv,r,cB),bool): last=(r,d)
            if last is None: skipped.append(f'{sh} (no typed balance, so no stock date)'); continue
            asof=last[1]; inbound=[]; last_rec=None
            bal_by_date={d:_num(cv(rv,r,cB)) for r,d in rows if _num(cv(rv,r,cB)) is not None}
            dem_by_date={d:info(r).get('demand') for r,d in rows if isinstance(info(r).get('demand'),(int,float))}
            pdays_by_date={d:(info(r).get('plantdays'),info(r).get('plantport'),info(r).get('ppcport')) for r,d in rows}
            ppcit_by_date={d:info(r).get('ppcit') for r,d in rows if info(r).get('ppcit')}
            vlock_by_date={d:info(r).get('vlock') for r,d in rows if info(r).get('vlock')}     # Settings ALLOWED VENDORS (aligned file, column AK)
            cU_=next((c for h,c in hdr.items() if h.startswith('ppc in transit (supplier')),None)
            ppcsup_names={' '.join(n_.split()) for r,d in rows if cU_ and d>=asof for n_ in str(cv(rv,r,cU_) or '').split(' / ') if n_.strip()}
            for r,d in rows:
                if cDl:
                    qd=_num(cv(rv,r,cDl))
                    if qd and qd>0 and d<=asof: last_rec=(d,qd)
                if d>plan_end: continue
                q=_num(cv(rv,r,cS))
                if q and q>0 and d>=asof-dt.timedelta(days=14):
                    g=lambda c: cv(rv,r,c) if c else None
                    inbound.append(dict(date=d,qty=q,row=r,remarks=str(g(cRem) or '').strip(),po=g(cPO),shipper=g(cSh),booking=g(cBk),readiness=g(cRd),etd=g(cETD),bl=g(cBL),vessel=str(g(cVs) or '').replace('\n',' ').strip(),info=info(r)))
            label=LABELS.get((sh,bi)) or (sh if len(starts)==1 else f'{sh} - {mname[:18]}')
            refs=[(cv(rv,r,cPO) if cPO else None,cv(rv,r,cBL) if cBL else None) for r,d in rows]     # PO + BL# of every row: In_Transit match
            groups.append(dict(label=label,sheet=sh,block=bi,uom=uom,codes=sorted(codes),asof=asof,inbound=inbound,last_receipt=last_rec,asof_info=info(last[0]),bal_by_date=bal_by_date,dem_by_date=dem_by_date,pdays_by_date=pdays_by_date,ppcit_by_date=ppcit_by_date,vlock_by_date=vlock_by_date,ppcsup_names=ppcsup_names,
                               refs=[t for t in refs if any(v not in (None,'') for v in t)]))
    if ARCHIVE_SHEET in wv.sheetnames:          # aligned WSA file: the delivery records of the deleted rows (their PO + BL# count for the match too)
        it_=wv[ARCHIVE_SHEET].iter_rows(values_only=True); h_=[_hnorm(v) for v in next(it_,())]
        iS_,iP_,iB_=(h_.index(k) if k in h_ else None for k in ('planning sheet','po','bl#'))
        byl_={g['label']:g for g in groups}
        v_=lambda r,i: r[i] if (i is not None and i<len(r)) else None
        if iS_ is not None:
            for r in it_:
                g=byl_.get(v_(r,iS_))
                if g and (v_(r,iP_) not in (None,'') or v_(r,iB_) not in (None,'')): g['refs'].append((v_(r,iP_),v_(r,iB_)))
    return groups,skipped

def _sheet_pools(sheets):
    """BW['sheets'] with the sheets that share a material merged (back glass 333 on 'Glass G12R Topcon SB Back' + 'Glass G12R Topcon
       Subanci Back' = one MIS material): balance = total of the sheets (None if one has none), demand = total per date."""
    labs=list(sheets); par={l:l for l in labs}
    def root(l):
        while par[l]!=l: par[l]=par[par[l]]; l=par[l]
        return l
    owner={}
    for l in labs:
        for m in sheets[l].get('members') or []:
            if m in owner: par[root(l)]=root(owner[m])
            else: owner[m]=l
    grp={}
    for l in labs: grp.setdefault(root(l),[]).append(l)
    out={}
    for ls in grp.values():
        if len(ls)==1: out[ls[0]]=sheets[ls[0]]; continue
        shs=[sheets[l] for l in ls]; dem={}
        for sh in shs:
            for d,v in (sh.get('dem') or {}).items(): dem[d]=dem.get(d,0.0)+(v or 0.0)
        out[' + '.join(ls)]=dict(shs[0],members=sorted({m for sh in shs for m in sh.get('members') or []}),dem=dem,n=len(shs),
                                bal=sum(float(sh['bal']) for sh in shs) if all(sh.get('bal') is not None for sh in shs) else None,
                                sheet=' + '.join(sh['sheet'] for sh in shs),asof=max(sh['asof'] for sh in shs))
    return out
def below_zero_periods(g):
    """From the sheet stock date: the periods (from, to) where the sheet's Plant Days < 0 and where its Plant + Port Days < 0.
       A day with no demand ('na') keeps the state of the day before (Plant Days: the sign of the Balance).
       'pi' = Plant + Port Days (PPC In Transit) - the column with the arrivals of your PPC In_Transit sheet.
       Returns {'pd': [...], 'pp': [...], 'pi': [...]} with dict(fr, to, low, open) - open = still below zero on the last date of the sheet."""
    asof=g['asof']; pdd=g.get('pdays_by_date') or {}; bal=g.get('bal_by_date') or {}
    ds=sorted(d for d in (set(pdd)|set(bal)) if d>=asof)
    res={}
    for key,ix in (('pd',0),('pp',1),('pi',2)):
        per=[]; cur=None; prev=False
        for d in ds:
            t_=pdd.get(d) or ()
            v=t_[ix] if ix<len(t_) else None
            neg=((bal.get(d) is not None and bal[d]<0) if ix==0 else prev) if v is None else v<0
            if neg:
                if cur is None: cur=dict(fr=d,to=d,low=v,open=False)
                else:
                    cur['to']=d
                    if v is not None and (cur['low'] is None or v<cur['low']): cur['low']=v
            elif cur is not None: per.append(cur); cur=None
            prev=neg
        if cur is not None: cur['open']=True; per.append(cur)
        res[key]=per
    return res
def periods_text(per, most=4, end='end of sheet'):
    """'08-Nov-26 -> 11-Nov-26; 20-Nov-26 -> 31-Mar-27 (end of sheet)' - or 'none'."""
    if not per: return 'none'
    t_='; '.join(f"{p['fr']:%d-%b-%y} \u2192 {p['to']:%d-%b-%y}"+(f' ({end})' if p['open'] else '') for p in per[:most])
    return t_+(f' ... +{len(per)-most} more (Below zero sheet)' if len(per)>most else '')
def other_materials_calc(AD, covered):
    """The WSA-sheet calculation for every MRP material WITHOUT a WSA planning sheet (nothing is written to the WSA file):
       Balance(day) = stock today - demand of the days before; Plant Days = Balance / Demand;
       Plant + Port Days (PPC In Transit) = (Balance + arrivals of the PPC In_Transit sheet up to that day) / Demand.
       Stock = the stock 1_MRP uses (MB52 / PPC count, JW at the OEM left out), Demand = module plan x BOM (OEM plan left out)."""
    dates=AD['dates']; nd=len(dates); arr=(AD.get('ppc') or {}).get('arr') or {}; out=[]
    for m,b in AD['blocks'].items():
        if m in covered: continue
        plan=b['plan']; st=float(b['stock'] or 0); a=arr.get(m) or [0.0]*nd
        bal=st; cum=0.0; pdd={}; bal_by={}
        for i,d in enumerate(dates):
            if i: bal-=plan[i-1]
            cum+=a[i]; dem=plan[i]
            pdd[d]=((bal/dem) if dem else None,None,((bal+cum)/dem) if dem else None); bal_by[d]=bal
        per=below_zero_periods(dict(asof=dates[0],pdays_by_date=pdd,bal_by_date=bal_by))
        nxt=next(((dates[i],a[i]) for i in range(nd) if a[i]>0),None)
        out.append(dict(mat=m,common=b.get('common') or m,uom=b.get('uom') or '',kind='Raw material' if (str(m).upper().startswith('RM') and not str(m).upper().startswith('RMLAB') and not is_consumable(m)) else 'Packaging / consumable',   # barcode / back label = packaging (user, Nov 2026)
                        stock=st,dem7=sum(plan[:7]),dem30=sum(plan[:30]),pd0=pdd[dates[0]][0],pi0=pdd[dates[0]][2],per=per,it=sum(a),nxt=nxt,
                        ppc_by={dates[i]:a[i] for i in range(nd) if a[i]>0},asof=dates[0]))
    first=lambda x: min([p['fr'] for p in x['per']['pi']]+[p['fr'] for p in x['per']['pd']]+[pd.Timestamp('2100-01-01')])
    return sorted(out,key=lambda x:(first(x),str(x['common'])))

def add_other_materials(path, rows, ppc, wsa_of=None):
    """Inbound: sheet 'Other materials' (one row per material), their periods added to 'Below zero', and sheet 'PPC In Transit'
       (every line of your PPC In_Transit sheet: what it says, the BOM material it was traced to and why, where it is counted)."""
    from openpyxl.styles import Border, Side
    wb=load_workbook(path); AR='Calibri'
    bf=Font(name=AR,size=10); hf=Font(name=AR,bold=True,color='FFFFFF',size=10)
    hfill=PatternFill('solid',fgColor='1F3864'); ctr=Alignment(horizontal='center',vertical='center',wrap_text=True); dfy='dd-mmm-yy'
    if 'Other materials' in wb.sheetnames: del wb['Other materials']
    pos=(wb.sheetnames.index('Below zero')+1) if 'Below zero' in wb.sheetnames else 1
    ws=wb.create_sheet('Other materials',pos); ws.sheet_view.showGridLines=False
    ws['A1']='Other materials - the WSA calculation for the BOM materials that have no WSA planning sheet (calculated by the script, not written to the WSA file)'
    ws['A1'].font=Font(name=AR,size=12,bold=True)
    ws['A2']=('Same rules as the WSA sheets: Balance = stock today (MB52 / PPC count, JW at the OEM left out) minus the demand of the days before; '
              'Demand = module plan x BOM (OEM plan left out); Plant Days = Balance / Demand; Plant + Port Days (PPC In Transit) = (Balance + arrivals of your '
              f"PPC In_Transit sheet up to that day) / Demand{(' - ' + ppc['src']) if ppc and ppc.get('src') else ' - no PPC In_Transit sheet found'}. "
              'Solar cells are not in the MRP, so they are not listed.'); ws['A2'].font=bf
    hd=['Material','Code','Type','UOM','Stock today','Demand next 7 days','Demand next 30 days','Plant Days today','Plant + Port Days (PPC In Transit) today',
        'Plant Days < 0 (from \u2192 to)','Plant + Port Days (PPC In Transit) < 0 (from \u2192 to)','PPC In Transit qty (all)','Next PPC arrival','Next PPC arrival qty']
    for j,h in enumerate(hd,1):
        c=ws.cell(4,j,h); c.font=hf; c.fill=hfill; c.alignment=ctr
    ws.row_dimensions[4].height=44
    for i,x in enumerate(rows):
        r=5+i
        vals=[x['common'],x['mat'],x['kind'],x['uom'],x['stock'],x['dem7'],x['dem30'],x['pd0'] if x['pd0'] is not None else 'no demand today',
              x['pi0'] if x['pi0'] is not None else 'no demand today',periods_text(x['per']['pd'],end='end of plan'),periods_text(x['per']['pi'],end='end of plan'),x['it'] or None,
              x['nxt'][0] if x['nxt'] else None,x['nxt'][1] if x['nxt'] else None]
        for j,v in enumerate(vals,1):
            c=ws.cell(r,j,v); c.font=bf
            if j in (5,6,7,12,14): c.number_format='#,##0'
            elif j in (8,9): c.number_format='0.0'
            elif j==13: c.number_format=dfy
        for j,k_ in ((10,'pd'),(11,'pi')):
            ws.cell(r,j).font=Font(name=AR,size=10,color='C00000' if x['per'][k_] else '008000'); ws.cell(r,j).alignment=Alignment(wrap_text=True,vertical='top')
    for c_,w_ in zip('ABCDEFGHIJKLMN',[40,17,20,7,13,13,13,11,14,40,40,13,12,12]): ws.column_dimensions[c_].width=w_
    ws.freeze_panes='B5'; ws.auto_filter.ref=f'A4:N{4+max(1,len(rows))}'
    if 'Below zero' in wb.sheetnames:                                  # their periods go to the Below zero list too
        bz=wb['Below zero']; old=[]
        for r in bz.iter_rows(min_row=5,max_col=13,values_only=True):
            if r and r[0] and r[4]: old.append(tuple(r))
        new=[]
        for x in rows:
            for key_,meas_ in (('pd','Plant Days < 0'),('pi','Plant + Port Days (PPC In Transit) < 0')):
                for p_ in x['per'][key_]:
                    pa=sorted((d_,q_) for d_,q_ in x['ppc_by'].items() if d_>=p_['fr'])
                    new.append(('(no WSA sheet)',f"{x['common']} ({x['mat']})",x['uom'],x['asof'].to_pydatetime(),meas_,p_['fr'].to_pydatetime(),p_['to'].to_pydatetime(),
                                (p_['to']-p_['fr']).days+1,p_['low'],'Yes' if p_['open'] else '',pa[0][0].to_pydatetime() if pa else None,pa[0][1] if pa else None,'PPC In Transit' if pa else ''))
        allr=sorted(old+new,key=lambda t_:(pd.Timestamp(t_[5]),str(t_[4]),str(t_[0])))
        for row_ in bz.iter_rows(min_row=5,max_row=max(bz.max_row,5)):
            for c in row_: c.value=None
        for i,t_ in enumerate(allr):
            for j,v in enumerate(t_,1):
                c=bz.cell(5+i,j,v); c.font=bf
                if j in (4,6,7,11): c.number_format=dfy
                elif j==9: c.number_format='0.0'
                elif j in (8,12): c.number_format='#,##0'
            if str(t_[4]).startswith('Plant + Port'): bz.cell(5+i,5).font=Font(name=AR,size=10,bold=True,color='C55A11' if 'PPC' in str(t_[4]) else 'C00000')
        bz.auto_filter.ref=f'A4:M{4+max(1,len(allr))}'
    lines=(ppc or {}).get('lines') or []; why_='; '.join((ppc or {}).get('why') or [])
    if lines:                                                          # the trace of every In_Transit line
        if 'PPC In Transit' in wb.sheetnames: del wb['PPC In Transit']
        tr=wb.create_sheet('PPC In Transit',wb.sheetnames.index('Other materials')+1); tr.sheet_view.showGridLines=False
        tr['A1']="PPC In Transit - every line of your PPC In_Transit sheet traced to YOUR BOM material (supplier + description + RM Code + Final RM Code)"
        tr['A1'].font=Font(name=AR,size=12,bold=True)
        tr['A2']=('Your WSA first: the sheet where you typed this PO + BL# (HBLNo. / MBL; else the BL#, else the PO) gives the material when it fits the '
                  'description. Else the Final RM Code is replaced only when the description (size, thickness, HS, GSM, front / back, product type) and the '
                  'supplier clearly fit another code - usually the RM Code. Amber = traced to another code than Final RM Code; red = CHECK (your WSA sheet '
                  'has the shipment as another size / HS) or not a BOM material (not counted). Late lines are counted on today. A supplier no customer '
                  'may use (Settings ALLOWED VENDORS) is not counted (Counted on blank), same as the 1_MRP ALERT.'); tr['A2'].font=bf
        if why_: tr['A3']='NOTE: '+why_; tr['A3'].font=Font(name=AR,size=10,bold=True,color='C00000')
        hd2=['Supplier','Invoice','PO','Description (In_Transit)','RM Code','Final RM Code','Traced to BOM code','BOM material description','BOM supplier',
             'Description from','How it was traced','Pending qty','MRP code (after code conversion)','Qty in MRP unit','Plant arrival (Final AtA)','Counted on','Shown on',
             'Your WSA sheet with this PO / BL#']
        for j,h in enumerate(hd2,1):
            c=tr.cell(4,j,h); c.font=hf; c.fill=hfill; c.alignment=ctr
        tr.row_dimensions[4].height=44
        amber=PatternFill('solid',fgColor='FFE699')
        wsa_of=wsa_of or {}
        def shown(x):
            if not x.get('code'): return '-'
            ls_=wsa_of.get(x.get('mat'))
            if isinstance(ls_,str): ls_=[ls_]
            if ls_ and len(ls_)>1:                                     # a material on 2+ sheets: the sheet with this PO / BL#, else by supplier
                t_=[l_ for l_ in ls_ if l_ in set(x.get('wsa_sheets') or [])]
                return f"WSA sheet: {t_[0]}" if len(t_)==1 else 'WSA sheet: '+' or '.join(ls_)+' (by supplier - Settings ALLOWED VENDORS)'
            if ls_: return f"WSA sheet: {ls_[0]}"
            if x.get('in_mrp'): return 'Other materials'
            return 'not in the MRP (no plan for it)'
        order_=lambda x:(0 if (x.get('code') and x['code']!=x['final_code']) else (1 if 'CHECK' in str(x.get('how')) else (2 if not x.get('code') else 3)),str(x['supplier']),pd.Timestamp(x['date']) if x.get('date') is not None else pd.Timestamp('2100-01-01'))
        def dead_(x):                                                  # Settings ALLOWED VENDORS: a supplier no customer may use = not counted (= 1_MRP ALERT, `vp_dead`)
            dn_=vp_dead(x['mat'])[2] if x.get('mat') and x.get('arrive') is not None else None
            return bool(dn_ and dn_(x.get('supplier')))
        for i,x in enumerate(sorted(lines,key=order_)):
            r=5+i; dd_=dead_(x)
            vals=[x['supplier'],x['invoice'],x['po'],x['description'],x['rm_code'] or None,x['raw_code'] or None,x.get('code'),x.get('bom_desc') or None,x.get('bom_sup') or None,
                  x.get('src') or None,x.get('how'),x['qty'],x.get('mat'),x.get('qty_mrp'),x['date'].to_pydatetime() if x.get('date') is not None else None,
                  (x['arrive'].to_pydatetime() if x.get('arrive') is not None and not dd_ else None),
                  ('not counted: no customer may use this supplier (Settings ALLOWED VENDORS)' if dd_ else shown(x)),
                  (f"{', '.join(x['wsa_sheets'])} ({x['wsa_how']})" if x.get('wsa_sheets') else '-')]
            for j,v in enumerate(vals,1):
                c=tr.cell(r,j,v); c.font=bf
                if j in (12,14): c.number_format='#,##0.##'
                elif j in (15,16): c.number_format=dfy
            if x.get('code') and x['code']!=x['final_code']:
                for j in (7,11): tr.cell(r,j).fill=amber
            if not x.get('code') or 'CHECK' in str(x.get('how')): tr.cell(r,11).font=Font(name=AR,size=10,color='C00000',bold='CHECK' in str(x.get('how')))
            if x.get('late'): tr.cell(r,16).font=Font(name=AR,size=10,color='C55A11')
            if dd_: tr.cell(r,17).font=Font(name=AR,size=10,color='C00000')
        for c_,w_ in zip('ABCDEFGHIJKLMNOPQR',[30,13,12,46,15,17,15,36,26,10,48,11,15,12,12,12,30,34]): tr.column_dimensions[c_].width=w_
        tr.freeze_panes='E5'; tr.auto_filter.ref=f'A4:R{4+len(lines)}'
    else:                                                              # nothing read: say why (also in T3 of every WSA sheet)
        if 'PPC In Transit' in wb.sheetnames: del wb['PPC In Transit']
        tr=wb.create_sheet('PPC In Transit',(wb.sheetnames.index('Other materials')+1) if 'Other materials' in wb.sheetnames else len(wb.sheetnames))
        tr.sheet_view.showGridLines=False
        tr['A1']="PPC In Transit - NOTHING READ: column T 'PPC In Transit (arrivals)' of the WSA sheets is empty"; tr['A1'].font=Font(name=AR,size=12,bold=True,color='C00000')
        tr['A3']='Why: '+(why_ or 'no PPC file with an In_Transit sheet'); tr['A3'].font=bf
    if 'Summary' in wb.sheetnames:
        c=wb['Summary']['D4']; c.value=f"Other BOM materials (no WSA sheet): {len(rows)} on the sheet 'Other materials'"; c.font=Font(name=AR,size=10,bold=True,color='C55A11')
    if 'Read me' in wb.sheetnames:
        rd_=wb['Read me']; r_=rd_.max_row+2
        rd_.cell(r_,1,"Other materials - the WSA calculation (Balance, Plant Days, Plant + Port Days with your PPC In_Transit) for the BOM materials without a WSA planning sheet; their below-zero periods are on the Below zero sheet.").font=bf
        rd_.cell(r_+1,1,"PPC In Transit - every line of your PPC In_Transit sheet, the BOM material it was traced to (your WSA PO + BL#, description, supplier, RM Code, Final RM Code), why, where it is counted, and your WSA sheet with that PO / BL#.").font=bf
    wb.save(path)

def clean_status(rem):
    l=rem.lower()
    if 'canc' in l: return 'Cancelled'
    if 'grn' in l: return 'GRN pending'
    if 'approved' in l: return 'Approved for booking'
    if 'planned' in l: return 'Planned'
    return ''
def port_of(rem):

    if clean_status(rem): return ''
    return rem.title() if rem else ''
def ymd(v): return pd.Timestamp(v).date() if hasattr(v,'year') else None


