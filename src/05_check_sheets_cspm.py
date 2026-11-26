# ===================== Check sheet: discrepancies between MIS, the WSA planning file and the MRP workbook =====================
def build_checks(DS, BW, mis_path, mrp_path=None, LIMIT_DAYS=30):
    """Returns a list of issues (dicts). Sources: MIS RM Dashboard (DS + raw sheets), WSA planning file (BW), MRP input workbook."""
    import re as _re
    I=[]; today=pd.Timestamp(DS['dates'][0]).normalize(); nd=len(DS['dates']); H=min(LIMIT_DAYS,nd)
    def add(area,check,sev,material='',code='',a_lab='',a=None,b_lab='',b=None,detail='',action=''):
        diff=(b-a) if isinstance(a,(int,float)) and isinstance(b,(int,float)) else None
        I.append(dict(area=area,check=check,sev=sev,material=material,code=code,a_lab=a_lab,a=a,b_lab=b_lab,b=b,diff=diff,detail=detail,action=action))
    blocks={b['mat']:b for b in DS['blocks']}; c2n={b['mat']:b['common'] for b in DS['blocks']}; n2c={v:k for k,v in c2n.items()}
    for a_,d_ in (DS.get('alias') or {}).items(): c2n.setdefault(a_,c2n.get(d_))
    avgplan={m:(sum(b['plan'][:H])/H if H else 0) for m,b in blocks.items()}
    mis=wb_ro(mis_path)
    num=lambda v: float(v) if isinstance(v,(int,float)) and not isinstance(v,bool) else (float(v) if isinstance(v,str) and _re.fullmatch(r'-?\d+(\.\d+)?',v.strip()) else None)
    fam=lambda n: {'G12R Front Glass':'G12R Glass','G12R Back Glass':'G12R Glass','M10R Glass Front':'M10R Glass','M10R Glass Back':'M10R Glass'}.get(str(n),str(n))

    # ---------- STOCK ----------
    for m,b in blocks.items():
        fac=b.get('alt_factor') or {}
        tp_=b.get('tape')
        if tp_:
            add('Master data','Tape rolls converted by length (from the description)','Low',b['common'],m,f"{m} roll",tp_.get('len'),'Other rolls in stock (converted)',round(tp_['conv'],1),
                f"1 roll of {m} = {tp_.get('len'):g} m; "+', '.join(f"{c} = {l:g} m (x{l/tp_['len']:g})" for c,l in tp_['lens'].items() if l)+
                (f"; MIS Adhoc conversion {tp_['adhoc']:,.1f} dropped (counted directly instead)" if tp_.get('dropped') else (f"; MIS Adhoc {tp_['adhoc']:,.1f} kept (does not match the converted stock)" if tp_.get('adhoc') else '')),
                'OK - check the roll lengths in the descriptions are right')
            fac={}
        if fac:
            add('Master data','Alternate code converted by MIS (counted)','Low',b['common'],m,'',None,'Opening stock from conversion',b.get('adhoc_open') or None,
                'Stock of '+', '.join(f'{c} (x{f:g})' for c,f in fac.items())+f' is in {m} opening stock through the MIS Adhoc conversion; their in-transit is counted with the same ratio',
                'OK - no action, or ask MIS to move the dashboard to the code in use')
        unc=b.get('alt_uncounted') or {}
        if unc:
            tot=sum(unc.values()); ap=avgplan.get(m,0)
            add('Stock','Same-name stock NOT counted (alternate code)','High' if ap and tot>=ap*3 else 'Medium',b['common'],m,'Stock counted',own_stock(b),'Not counted',tot,
                ', '.join(f'{c} {q:,.1f}' for c,q in unc.items())+f' has the same MIS Common Name (alternate code, other unit) - not counted (MIS Adhoc is not used); if usable for {m}, that is'+(f' {tot/ap:.1f} more days' if ap else ' more stock'),
                'Confirm with production: if usable, add it on Settings > Substitutes (MRP material, substitute, factor) so the MRP counts it')
        for c_,nm_,r_,q_ in (b.get('conv_unlinked') or []):
            add('Master data','Code not linked on the MIS Common Name sheet','High',b['common'],m,'',None,'MIS Adhoc today',q_,
                f"MIS converts {c_} ('{nm_}') into {m} on its Adhoc sheet (x{r_:g}), but the Common Name sheet calls it '{nm_}', not '{b['common']}'. "
                f"Its stock shows on the Adhoc row (day 1), not in Stock, and its in-transit is not counted",
                f"Add / correct the row on the MIS Common Name sheet: {c_} -> Common Name '{b['common']}'")
        ao=(b.get('adhoc_other') or 0)-sum(q_ for _,_,_,q_ in (b.get('conv_unlinked') or []))
        if ao>0.5:
            add('Stock','MIS Adhoc quantity counted today (not a code conversion)','High' if ao>own_stock(b) else 'Medium',b['common'],m,'MB52 stock (JW at OEM out)',own_stock(b),'Adhoc today',ao,
                f'MIS Adhoc adds {ao:,.0f} on day 1 ({ao/max(own_stock(b),1):.1f}x the stock) - counted as in MIS; the WSA sheet balance may not include it',
                'Confirm what this Adhoc quantity is (stock elsewhere? receipt?) - if not real, ask MIS to remove it')

    # 1 MIS stock vs WSA sheet balance on the sheet stock date (pooled sheets compared as a whole; sheets sharing a material together)
    for lab,sh in _sheet_pools(BW.get('sheets',{})).items():
        mems=[m for m in sh['members'] if m in n2c]
        if not mems or sh['bal'] is None: continue
        est_=lambda c: STOCK_EST[c]['est'] if (c in STOCK_EST and STOCK_EST[c].get('count') is not None and STOCK_EST[c].get('est') is not None) else None   # counted = count rolled forward (= ALERT stock)
        mis_st=sum(est_(n2c[m]) if est_(n2c[m]) is not None else own_stock(blocks[n2c[m]]) for m in mems)  # own stock (JW at the OEM in no calculation); pooled sheet (e.g. glass front + back): WSA balance = total of the materials
        wsa=float(sh['bal']); d=mis_st-wsa; ap=sum(avgplan[n2c[m]] for m in mems)
        days=abs(d)/ap if ap else None
        if abs(d)>max(1,0.02*max(abs(wsa),abs(mis_st))):
            sev='High' if (days is not None and days>=2) else 'Medium' if (days is None or days>=0.5) else 'Low'
            add('Stock','MIS stock vs WSA sheet balance',sev,' / '.join(mems),', '.join(n2c[m] for m in mems),
                'WSA balance',wsa,'MIS stock',mis_st,
                f"WSA sheet{'s' if sh.get('n',1)>1 else ''} '{sh['sheet']}' balance on {sh['asof']:%d-%b} vs MIS stock today ({DS['src']})"+(f"; gap = {days:.1f} days of plan" if days is not None else '')
                +("; pooled sheet, MIS = total of its materials" if len(mems)>1 else '')+("; one material on 2+ sheets (by customer and supplier) - WSA = total of the sheets" if sh.get('n',1)>1 else ''),
                'Check which count is current (receipts or issues booked on one side only)')
    # 3 Supplier split gap (RM Status by Make vs MB52) for glass
    for m,b in blocks.items():
        for sp in b.get('sup') or []:
            if sp['name']=='Unassigned':
                add('Stock','Glass stock without supplier (RM Status vs MB52)','Low',b['common'],m,'MB52 total',own_stock(b),'Sum by Make (RM Status)',own_stock(b)-sp['stock'],
                    'Batches in MB52 that MIS RM Status does not assign to a Make','Ask MIS owner to map these batches to a supplier')
    # 5 RM unit: WSA sheet UOM vs MIS MB52 unit
    misu={}
    for r in mis['MB52'].iter_rows(min_row=2,values_only=True):
        c=str(r[2] or '').strip()
        if c and r[7]: misu.setdefault(c,str(r[7]).strip())
    for lab,sh in BW.get('sheets',{}).items():
        for mname in sh['members']:
            c=n2c.get(mname)
            syn=lambda u: {'SQM':'M2','M2':'M2','PCS':'EA','PC':'EA','NOS':'EA','EA':'EA','SET':'EA','SETS':'EA','ROLL':'EA','ROLLS':'EA','ROL':'EA','KG':'KG','KGS':'KG','DR':'DR'}.get(_re.sub(r'[^A-Z0-9]','',str(u).upper()),_re.sub(r'[^A-Z0-9]','',str(u).upper()))
            if c and sh['uom'] and misu.get(c) and syn(sh['uom'])!=syn(misu[c]):
                add('Units','WSA sheet unit vs SAP unit','Low',mname,c,'',None,'',None,f"WSA sheet '{sh['sheet']}' says {sh['uom']}, MB52 uses {misu[c]} (quantities look like {misu[c]})",'Correct the unit label on the planning sheet')

    # ---------- PLAN (next LIMIT_DAYS days) ----------
    # 6 MIS plan vs WSA sheet demand
    for lab,sh in _sheet_pools(BW.get('sheets',{})).items():
        mems=[m for m in sh['members'] if m in n2c]
        if not mems or not sh['dem']: continue
        wsa=sum(v for d,v in sh['dem'].items() if today<=pd.Timestamp(d)<today+pd.Timedelta(days=H))
        mis_p=sum(sum(blocks[n2c[m]]['plan'][:H]) for m in mems)
        if abs(mis_p-wsa)>max(1,0.10*max(mis_p,wsa)):
            add('Plan',f'MIS plan vs WSA sheet demand (next {H} days)','High' if abs(mis_p-wsa)>0.25*max(mis_p,wsa,1) else 'Medium',' / '.join(mems),', '.join(n2c[m] for m in mems),
                'WSA demand',wsa,'MIS plan',mis_p,f"WSA sheet{'s' if sh.get('n',1)>1 else ''} '{sh['sheet']}' Demand column vs MIS Total Plan Req"+(' (pooled sheet, MIS = total of its materials)' if len(mems)>1 else '')
                +(' (one material on 2+ sheets, by customer and supplier: WSA = total of the sheets)' if sh.get('n',1)>1 else '')
                +(' - about 2x: this glass sheet seems to plan front + back glass sets under one material' if mis_p and 1.8<wsa/mis_p<2.2 and 'glass' in sh['sheet'].lower() else ''),
                'Align the production plan used by the planning file and MIS')
    # 9 Customer names
    mis_c=set()
    for r in mis['Total Plan Req'].iter_rows(min_row=2,values_only=True):
        if tp_ok(r,DS['plant']) and r[2]: mis_c.add(' '.join(str(r[2]).split()))
    pp_c=set(); mp_cust=set()
    import difflib
    allc=sorted(mis_c|pp_c|{c for c in mp_cust if c})
    for i,a in enumerate(allc):
        for b in allc[i+1:]:
            if a.lower()!=b.lower() and difflib.SequenceMatcher(None,a.lower(),b.lower()).ratio()>=0.75:
                where=lambda x: '/'.join(n for n,s_ in (('MIS',mis_c),('Production Plan',pp_c),('Material Plan',mp_cust)) if x in s_)
                add('Master data','Customer names that look alike','Low',f'{a}  vs  {b}','','',None,'',None,f"'{a}' in {where(a)}; '{b}' in {where(b)}",'Confirm if same customer and use one spelling')

    # ---------- IN TRANSIT: shipment by shipment (BL number) ----------
    it=list(mis['In Transit'].iter_rows(values_only=True))
    M=[]
    for r in it[1:]:
        if len(r)<40: continue
        q=num(r[34])
        if not q or q<=0: continue
        code=str(r[39] or '').strip()
        if code in _SWP_INV(): q=q*_SWP_INV()[code][1]; code=_SWP_INV()[code][0]   # MIS tape code -> your BOM code (its roll), as the MRP
        M.append(dict(code=code,mat=c2n.get(code),fam=fam(c2n.get(code)) if code in c2n else code,qty=q,po=r[6],sup=r[1],vessel=r[12],
                      date=pd.Timestamp(r[35]).normalize() if isinstance(r[35],(dt.datetime,dt.date)) else None,
                      keys={t for f_ in (r[23],r[24],r[10]) for t in _re.split(r'[^A-Z0-9]+',str(f_ or '').upper()) if len(t)>=8 and any(ch.isdigit() for ch in t)}))
    W=[dict(d,fam=fam(d['name']),keys={t for t in _re.split(r'[^A-Z0-9]+',str(d['bl'] or '').upper()) if len(t)>=8 and any(ch.isdigit() for ch in t)})
       for d in BW.get('deliv',[]) if d['stage']=='In transit']
    idx={}
    for j,m in enumerate(M):
        for t in m['keys']: idx.setdefault(t,set()).add(j)
    parent={}
    def find(x):
        while parent.setdefault(x,x)!=x: x=parent[x]
        return x
    matchedW=set(); matchedM=set()
    for i,w in enumerate(W):
        for t in w['keys']:
            for j in idx.get(t,()):
                parent[find(('W',i))]=find(('M',j)); matchedW.add(i); matchedM.add(j)
    grp={}
    for i in matchedW: grp.setdefault(find(('W',i)),{'W':[],'M':[]})['W'].append(i)
    for j in matchedM: grp.setdefault(find(('M',j)),{'W':[],'M':[]})['M'].append(j)
    for g in grp.values():
        ws_=[W[i] for i in g['W']]; ms_=[M[j] for j in g['M']]
        bl=', '.join(sorted({t for w in ws_ for t in w['keys']})[:2]); ves=str(ws_[0]['vessel'] or '')[:35]
        wq={}; mq={}
        for w in ws_: wq[w['fam']]=wq.get(w['fam'],0)+w['qty']
        for m in ms_: mq[m['fam']]=mq.get(m['fam'],0)+m['qty']
        if set(wq)!=set(mq):
            add('In transit','Same shipment, different material','Low' if abs(sum(wq.values())-sum(mq.values()))<=2 else 'Medium',f"WSA: {', '.join(sorted(wq))} | MIS: {', '.join(sorted(mq))}",'','WSA qty',sum(wq.values()),'PPC qty',sum(mq.values()),
                f'BL {bl} ({ves}) - material coded differently in the two files','Agree one material code for this item')
        else:
            for f_,a in wq.items():
                b=mq.get(f_,0)
                if abs(a-b)>max(2,0.005*max(a,b)):
                    add('In transit','Same shipment, different quantity','High' if abs(a-b)>0.1*max(a,b) else 'Medium',f_,'','WSA qty',a,'PPC qty',b,f'BL {bl} ({ves})','Check the packing list / invoice quantity')
        wd=min(w['date'] for w in ws_); md=min((m['date'] for m in ms_ if m['date'] is not None),default=None)
        if md is not None:
            dd_=(md-pd.Timestamp(wd).normalize()).days
            if abs(dd_)>=3:
                add('In transit','Same shipment, different arrival date','Medium' if abs(dd_)>=7 else 'Low',', '.join(sorted(wq)),'','WSA plant date',None,'PPC Final AtA',None,
                    f"BL {bl} ({ves}): WSA {pd.Timestamp(wd):%d-%b} vs PPC {md:%d-%b} ({dd_:+d} days)",f"PPC In_Transit is the latest: move the WSA row to {md:%d-%b}")
    for i,w in enumerate(W):
        if i not in matchedW:
            add('In transit','In WSA, not in PPC In_Transit','High' if w['bl'] else 'Medium',w['name'],n2c.get(w['name'],''),'WSA qty',w['qty'],'PPC qty',0.0,
                f"Plant date {pd.Timestamp(w['date']):%d-%b}; PO {w['po'] or '-'}; BL {w['bl'] or '-'}; {str(w['vessel'] or '')[:35]}",
                'Ask MIS owner to add the shipment, or confirm it was received' if w['bl'] else 'No BL yet - confirm booking with supplier')
    for j,m in enumerate(M):
        if j not in matchedM:
            ap_=avgplan.get(m['code'],0)
            add('In transit','In PPC In_Transit, not in WSA','High' if (ap_ and m['qty']>=ap_) else ('Medium' if m['qty']>=100 else 'Low'),m['mat'] or m['code'],m['code'],'WSA qty',0.0,'PPC qty',m['qty'],
                f"{str(m['sup'] or '')[:30]}; PO {m['po'] or '-'}; BL {', '.join(sorted(m['keys'])[:2]) or '-'}; AtA {m['date']:%d-%b}" if m['date'] is not None else f"PO {m['po']}",
                'Ask the planner to add it to the WSA sheet, or confirm it was received')
    # codes in PPC In_Transit that MRP cannot place
    oth={}
    for m in M:
        if m['code'] not in c2n: oth[m['code']]=oth.get(m['code'],0)+m['qty']
    for c,q in oth.items():
        add('Master data','PPC In_Transit code not on Daily Dashboard','High' if c.lower().startswith('no material') else 'Medium',c,c,'',None,'Pending qty',q,
            'Quantity cannot be counted in Daily Stock' + (" - MIS lookup failed (RM code missing)" if c.lower().startswith('no material') else ''),
            'Fix the RM code in MIS' if c.lower().startswith('no material') else 'Add the code to the dashboard, or map it to the right material')
    # ---------- OPEN PO ----------
    mis_po={}
    for m in M:
        if m['po']: mis_po.setdefault(str(m['po']).strip().removesuffix('.0'),set()).add(m['fam'])
    for d in BW.get('deliv',[]):
        if d['stage']!='Open PO' or not d['first']: continue
        po=str(d['po'] or '').strip().removesuffix('.0')
        if pd.Timestamp(d['date']).normalize()<today:
            add('Open PO','Open PO past its plant date','High',d['name'],n2c.get(d['name'],''),'',None,'Qty',d['qty'],f"Plant date {pd.Timestamp(d['date']):%d-%b}, still no vessel; PO {po or '-'}",'Get a new date from the supplier')
        elif po and po in mis_po and fam(d['name']) in mis_po[po]:
            add('Open PO','WSA Open PO already shipping in MIS?','Low',d['name'],n2c.get(d['name'],''),'',None,'Qty',d['qty'],
                f"PO {po} has shipments in PPC In_Transit; this WSA line (plant date {pd.Timestamp(d['date']):%d-%b}) has no vessel yet",'Check whether this line is already booked (avoid double counting)')
    return I

# ===================== MIS layout: columns read by POSITION (checked every run, shown on Check) =====================
# In Transit is read by header name (normalize_in_transit). These sheets are read by column position, so a column MIS inserts
# or removes would silently shift the numbers. Every run: (1) the header in each position is compared with the last run
# (stored on the hidden 'MIS Layout' sheet of MRP_Daily), (2) the data in each column is checked for the right kind of value.
# (column index 0 = A, purpose, kind: date / num / code / text / any)
MIS_POSITIONS={
 'Total Plan Req':[(0,'Date','date'),(1,'Plant = production line','text'),(2,'Customer','text'),(3,'Modules per plan line (Value)','num'),
                   (4,'MIS BOM Name','text'),(10,'Material code','code'),(12,'Number used to select rows for the BOM rate check','any'),
                   (13,'Unit of measure','text'),(14,'Requirement quantity','num'),(17,'OA (plan line key)','any'),(18,'OEM Category','text'),(21,'Company (plant)','text')],
 'MB52':[(2,'Material code','code'),(5,'Material description','text'),(7,'Unit of measure','text'),(13,'Plant / storage (USA, USA WAREHOUSE)','text'),
         (14,'Total stock','num'),(18,'Material name','text')],
 'RM Status for OEM JW':[(0,'Make (supplier)','text'),(1,'Material code','code'),(5,'Stock USA','num'),(7,'Stock USA Warehouse','num'),(11,'JW WSA to OEM','num')],
 'Common Name':[(0,'Material code','code'),(1,'Material description','text'),(2,'Unit of measure','text'),(3,'Common Name','text')],
 'Standard BOM':[(2,'BOM Name','text'),(3,'Material code','code'),(5,'Quantity per module','num')]}
MIS_OPTIONAL={'Standard BOM','Common Name'}
IT_USED=[IT_CANON[i] for i in (1,4,6,7,8,12,23,24,31,34,35,39)]     # In Transit columns the script uses
_MCODE=re.compile(r'^[A-Z]{2,}[A-Z0-9]*\d{3,}$')
def _kind_ok(v,kind):
    if kind=='date': return isinstance(v,(dt.datetime,dt.date))
    if kind=='num': return isinstance(v,(int,float)) and not isinstance(v,bool)
    if kind=='code': return bool(_MCODE.match(str(v).strip().upper()))
    if kind=='text': return isinstance(v,str)
    return True
def mis_layout_check(mis_path, prev_path=None, sample=3000):
    """Returns (Check issues, rows for the Check section and the hidden 'MIS Layout' sheet)."""
    wb=wb_ro(mis_path); prev={}
    if prev_path and os.path.exists(prev_path):
        try:
            pw=wb_ro(prev_path)
            if 'MIS Layout' in pw.sheetnames:
                for r in list(pw['MIS Layout'].iter_rows(values_only=True))[1:]:
                    if r and r[0] and r[1] is not None: prev[(str(r[0]),str(r[1]))]=str(r[2] if r[2] is not None else '')
        except Exception: prev={}
    iss=[]; rows=[]
    def add(sev,check,sheet,col,a_lab,a,b_lab,b,detail,action):
        iss.append(dict(area='MIS layout',check=check,sev=sev,material=sheet,code=col,a_lab=a_lab,a=a,b_lab=b_lab,b=b,diff=None,detail=detail,action=action))
    for sheet,cols in MIS_POSITIONS.items():
        if sheet not in wb.sheetnames:
            if sheet not in MIS_OPTIONAL:
                add('High','MIS sheet missing',sheet,'','',None,'',None,f"Sheet '{sheet}' is not in {os.path.basename(mis_path)} - everything read from it is zero / blank",'Get the full MIS RM Dashboard file')
            continue
        data=wb[sheet]._rows; hdr=list(data[0]) if data else []; body=data[1:1+sample]
        for i,purpose,kind in cols:
            col=L(i+1); cur=str(hdr[i]).strip() if i<len(hdr) and hdr[i] is not None else ''
            old=prev.get((sheet,col)); vals=[r[i] for r in body if i<len(r) and r[i] not in (None,'')]
            ok_n=sum(1 for v in vals if _kind_ok(v,kind)); share=(ok_n/len(vals)) if vals else None
            moved=old is not None and old!=cur
            bad=kind!='any' and vals and len(vals)>=20 and share<0.8
            empty=kind!='any' and not vals and len(body)>=20
            res='OK'
            if moved: res='HEADER CHANGED'; add('High','MIS column header changed since the last run (column moved?)',sheet,col,'Header last run',old,'Header now',cur,
                    f"Column {col} is read as '{purpose}' by position. If MIS inserted or removed a column, every number from this sheet is wrong.",
                    'Check the MIS file layout; if the column moved, tell MIS to restore it (or update the script)')
            if bad: res=(res+' / ' if res!='OK' else '')+'WRONG DATA'; add('High','MIS column does not hold the expected data',sheet,col,'Expected',kind,'Matching values',f'{share:.0%}',
                    f"Column {col} (header '{cur}') should be '{purpose}' ({kind}) but only {share:.0%} of {len(vals)} values look right.",'Check the MIS file layout before using the results')
            if empty: res=(res+' / ' if res!='OK' else '')+'EMPTY'; add('Medium','MIS column read by position is empty',sheet,col,'Expected',purpose,'',None,
                    f"Column {col} (header '{cur}') has no values in {len(body)} rows.",'Check the MIS file')
            rows.append(dict(sheet=sheet,col=col,purpose=purpose,kind=kind,cur=cur,old=old,share=share,res=res))
    if IT_MISSING:
        used=[h for h in IT_MISSING if h in IT_USED]; other=[h for h in IT_MISSING if h not in IT_USED]
        if used: add('High','PPC In_Transit: columns used by the script are missing','In Transit',', '.join(used),'',None,'',None,
                     'In Transit is read by header name; these headers were not found, so they are read as blank: '+', '.join(used),'Ask MIS to restore the header names')
        if other: add('Low','PPC In_Transit: other columns not found','In Transit',', '.join(other)[:60],'',None,'',None,'Not used for the numbers: '+', '.join(other),'None needed')
    if not prev: add('Low','MIS layout recorded - compared from the next run','(all)','','',None,'',None,
                     'No previous MRP_Daily with a MIS Layout sheet: the headers read by position are stored now and compared on the next run. Data checks were done.','Check once that the headers below are the right columns')
    return iss,rows

def make_mis_layout_xlsx(rows, path, store=False):
    """store=False: the Check section 'MIS columns read by position'; store=True: the hidden 'MIS Layout' sheet kept for the next run."""
    from openpyxl.styles import Border, Side
    AR='Calibri'; NAVY='1F3864'; thin=Side(style='thin',color='BFBFBF'); bd=Border(left=thin,right=thin,top=thin,bottom=thin)
    wb=Workbook(); ws=wb.active
    if store:
        ws.title='MIS Layout'; ws.append(['Sheet','Column','Header'])
        for x in rows: ws.append([x['sheet'],x['col'],x['cur']])
        wb.save(path); return
    ws.title='MIS columns'
    ws['A1']='MIS columns read by position - header now vs last run, and whether the data looks right'; ws['A1'].font=Font(name=AR,size=13,bold=True,color=NAVY)
    ws['A2']=('These MIS sheets are read by column position (In Transit is read by header name). A header that changed, or a column whose values are not '
              'the expected kind, means MIS changed the layout and the numbers from that sheet cannot be trusted.')
    ws['A2'].font=Font(name=AR,size=9,italic=True); ws.merge_cells('A2:H2'); ws['A2'].alignment=Alignment(wrap_text=True); ws.row_dimensions[2].height=28
    H=['Sheet','Column','Read as','Expected kind','Header now','Header last run','Values of the expected kind','Result']
    for j,h in enumerate(H,1):
        c=ws.cell(4,j,h); c.font=Font(name=AR,size=9,bold=True,color='FFFFFF'); c.fill=PatternFill('solid',fgColor=NAVY); c.border=bd; c.alignment=Alignment(horizontal='center',wrap_text=True)
    for i,x in enumerate(rows):
        vals=[x['sheet'],x['col'],x['purpose'],x['kind'],x['cur'],x['old'] if x['old'] is not None else '(first run)',x['share'],x['res']]
        for j,v in enumerate(vals,1):
            c=ws.cell(5+i,j,v); c.font=Font(name=AR,size=9,bold=j==8); c.border=bd
            if j==7: c.number_format='0%'
        ws.cell(5+i,8).fill=PatternFill('solid',fgColor='C6EFCE' if x['res']=='OK' else 'FFC7CE')
    for j,w in enumerate([20,8,40,10,28,28,12,22],1): ws.column_dimensions[L(j)].width=w
    wb.save(path)

def make_check_xlsx(issues, DS, path):
    from openpyxl.styles import Border, Side
    from openpyxl.worksheet.table import Table, TableStyleInfo
    AR='Calibri'; NAVY='1F3864'
    wb=Workbook(); ws=wb.active; ws.title='Check'
    hf=Font(name=AR,bold=True,color='FFFFFF',size=10); hfill=PatternFill('solid',fgColor=NAVY); bf=Font(name=AR,size=9)
    thin=Side(style='thin',color='BFBFBF'); bd=Border(left=thin,right=thin,top=thin,bottom=thin); ctr=Alignment(horizontal='center',vertical='center',wrap_text=True)
    sevfill={'High':'FFC7CE','Medium':'FFEB9C','Low':'DDEBF7'}
    ws['A1']='Data check - discrepancies between MIS RM Dashboard, WSA planning file and MRP workbook'; ws['A1'].font=Font(name=AR,size=14,bold=True,color=NAVY)
    ws['A2']=(f"Run {dt.date.today():%d-%b-%Y}. Areas: Stock, Plan (next 30 days), In transit (matched shipment by shipment on BL number), Open PO, Units, Master data. "
              "Severity: High = can change a shortage decision; Medium = should be fixed; Low = tidy-up (collapsed - click + on the left). "
              "MIS layout = MIS columns read by position (header moved / wrong data). Below: MIS columns, Segmentation, Unit check, Pack sizes, Plan compare (collapsed - click + to open).")
    ws['A2'].font=Font(name=AR,size=9,italic=True); ws['A2'].alignment=Alignment(wrap_text=True,vertical='top'); ws.merge_cells('A2:K2'); ws.row_dimensions[2].height=30
    areas=['MIS layout','Stock','Plan','In transit','Open PO','Units','Master data']
    for j,t in enumerate(['Area','High','Medium','Low','Total']): c=ws.cell(4,2+j,t); c.font=hf; c.fill=hfill; c.alignment=ctr; c.border=bd
    for i,a in enumerate(areas+['Total']):
        r=5+i; c=ws.cell(r,2,a); c.font=Font(name=AR,size=10,bold=(a=='Total')); c.border=bd
        sel=[x for x in issues if a=='Total' or x['area']==a]
        for j,sv in enumerate(['High','Medium','Low',None]):
            n=sum(1 for x in sel if sv is None or x['sev']==sv); c=ws.cell(r,3+j,n); c.alignment=ctr; c.border=bd; c.font=Font(name=AR,size=10,bold=True)
            if sv and n: c.fill=PatternFill('solid',fgColor=sevfill[sv])
    top=5+len(areas)+3
    # ---- TOP ISSUES: the few things to fix this week (High first, biggest first), one line each ----
    order0={'High':0,'Medium':1,'Low':2}
    tops=sorted([x for x in issues if x['sev'] in ('High','Medium')],key=lambda x:(order0[x['sev']],rrank(str(x.get('code') or '').strip()),-(abs(x['diff']) if isinstance(x['diff'],(int,float)) else 0)))[:15]
    if tops:
        ws.cell(top-1,1,'TOP ISSUES - fix these first (all issues are listed below)').font=Font(name=AR,size=12,bold=True,color='C00000')
        for j,t in enumerate(['Severity','Area','Check','Material','Code']): c=ws.cell(top,1+j,t); c.font=hf; c.fill=hfill; c.alignment=ctr; c.border=bd
        for i,x in enumerate(tops):
            r=top+1+i
            for j,v in enumerate([x['sev'],x['area'],x['check'],x['material'],x['code']]):              # no Action column (user, Nov 2026: 'no help')
                c=ws.cell(r,1+j,v); c.font=bf; c.border=bd; c.alignment=Alignment(vertical='top',wrap_text=j in (2,3))
            ws.cell(r,1).fill=PatternFill('solid',fgColor=sevfill[x['sev']])
        top=top+len(tops)+4
        ws.cell(top-1,1,'ALL ISSUES').font=Font(name=AR,size=12,bold=True,color=NAVY)
    hdr=['Severity','Area','Check','Material','Code','Source A','Value A','Source B','Value B','Difference (B - A)','Detail']
    for j,t in enumerate(hdr): c=ws.cell(top,1+j,t); c.font=hf; c.fill=hfill; c.alignment=ctr; c.border=bd
    order={'High':0,'Medium':1,'Low':2}
    issues=sorted(issues,key=lambda x:(order[x['sev']],rrank(str(x.get('code') or '').strip()),areas.index(x['area']) if x['area'] in areas else 9,-(abs(x['diff']) if isinstance(x['diff'],(int,float)) else 0)))   # severity, then highest shortage risk
    for i,x in enumerate(issues):
        r=top+1+i
        vals=[x['sev'],x['area'],x['check'],x['material'],x['code'],x['a_lab'] or None,x['a'],x['b_lab'] or None,x['b'],x['diff'],x['detail']]
        for j,v in enumerate(vals,1):
            c=ws.cell(r,j,round(v,3) if isinstance(v,float) else v); c.font=bf; c.border=bd; c.alignment=Alignment(vertical='top',wrap_text=j in (3,4,11))
            if j in (7,9,10): c.number_format='#,##0;-#,##0;0'
        ws.cell(r,1).fill=PatternFill('solid',fgColor=sevfill[x['sev']]); ws.cell(r,1).font=Font(name=AR,size=9,bold=True)
    last=top+max(1,len(issues))
    if not issues: ws.cell(top+1,1,'No discrepancies found')
    ws.add_table(Table(displayName='DataCheck',ref=f'A{top}:K{last}',tableStyleInfo=TableStyleInfo(name='TableStyleLight1',showRowStripes=True)))
    for col,w in zip('ABCDEFGHIJK',[9,11,30,30,16,16,12,16,12,13,70]): ws.column_dimensions[col].width=w
    ws.freeze_panes=f'A{top+1}'; ws.sheet_view.zoomScale=90
    set_print(wb.active,'Check'); lows=[r_ for r_ in range(1,ws.max_row+1) if ws.cell(r_,1).value=='Low']
    if lows:                                        # Low = tidy-up: collapsed, click + on the left to open
        ws.row_dimensions.group(lows[0],lows[-1],hidden=True,outline_level=1)
        ws.sheet_properties.outlinePr.summaryBelow=False
    wb.save(path)
    return {s_:sum(1 for x in issues if x['sev']==s_) for s_ in ('High','Medium','Low')}

# ===================== CS / PM in the new MRP: user-maintained BOM + Stock Count =====================
BOM_COLS=['Customer','Module Model','Item','Ict','Component','Component desciption','Quantity','UoM','Valid From','Valid to','MIS BOM Name']
def _guess_mis_bom(model, mis_names):
    m=str(model or '').upper()
    for n in mis_names:
        u=n.upper()
        if 'EPEP' in m and 'EPEP' in u: return n
        if ('BIN-21' in m or '625' in m) and u.endswith('G12R_USA'): return n
    return ''

def fill_count_bom(hdr, rows, path):
    """A workbook last saved by a tool other than Excel keeps formulas but no results: Today Qty_BOM then reads as empty.
       Recalculate it as Today Stock Qty x Quantity so packaging stock is never 0 by mistake."""
    hl=[str(h).strip().lower() if h is not None else '' for h in (hdr or [])]
    if 'today qty_bom' not in hl or 'today stock qty' not in hl: return 0
    io,iq=hl.index('today qty_bom'),hl.index('today stock qty'); i_f=hl.index('quantity') if 'quantity' in hl else None
    n=0
    for r in rows:
        if len(r)>max(io,iq) and not isinstance(r[io],(int,float)) and isinstance(r[iq],(int,float)):
            f=r[i_f] if (i_f is not None and len(r)>i_f and isinstance(r[i_f],(int,float))) else 1.0
            r[io]=r[iq]*f; n+=1
    path=PPC_SRC[0] or path
    if n: print(f'NOTE: {n} Count rows in {os.path.basename(path)} had no saved Today Qty_BOM (file last saved outside Excel) - '
                f'calculated as Today Stock Qty x Quantity. Open {os.path.basename(path)} in Excel and save it once to restore its formula results.')
    return n

def read_user_inputs(prev_daily, old_mrp, mis_names, model2bn=None):
    """BOM and Stock Count from YOUR MRP workbook (MRP.xlsx) - it is the master. The previous MRP_Daily is used only when your
       MRP has no such sheet. MIS BOM Name (link to the MIS plan) is not on your BOM: it is taken from the BOM sheet column
       if you add one, else from your Production Plan (customers using that Module Model), else the previous MRP_Daily, else a guess."""
    bom=[]; count_hdr=None; count=[]; src=None; prev_map={}
    model2bn=model2bn or {}
    def load(path):
        if not path or not os.path.exists(path): return None
        try: return wb_ro(path)
        except Exception as _e_: _swallow(3115,_e_); return None
    wp=load(prev_daily)
    if wp is not None and 'BOM' in wp.sheetnames:          # MIS BOM Name chosen earlier for each Module Model
        rows=[list(r) for r in wp['BOM'].iter_rows(values_only=True)]
        hr=next((i for i,r in enumerate(rows) if 'Component' in r and 'Quantity' in r),None)
        if hr is not None and 'MIS BOM Name' in rows[hr]:
            im,ib=rows[hr].index('Module Model'),rows[hr].index('MIS BOM Name')
            for r in rows[hr+1:]:
                if r[im] and r[ib]: prev_map.setdefault(str(r[im]).strip(),r[ib])
    for path in (old_mrp,prev_daily):
        wb=load(path)
        if wb is None: continue
        if 'BOM' in wb.sheetnames and not bom:
            rows=[list(r) for r in wb['BOM'].iter_rows(values_only=True)]
            std_=bom_rows_std(rows)                                  # old or new column layout
            if std_ is not None:
                for d0_ in std_:
                    d={c:d0_.get(c) for c in BOM_COLS}; d['Material group']=d0_.get('Material group'); d['Line #']=d0_.get('Line #')     # '<standard>_Alt' = CS / PM alternative (cspm_alt_fold)
                    if not d['Component']: continue
                    mdl=str(d['Module Model'] or '').strip()
                    if not d['MIS BOM Name']: d['MIS BOM Name']=model2bn.get(mdl) or prev_map.get(mdl) or _guess_mis_bom(mdl,mis_names)
                    bom.append(d)
                src=src or path
        if 'Stock Count' in wb.sheetnames and count_hdr is None and PPC_SRC[0] and path==old_mrp:   # ONLY the PPC 'Count' sheet
            hdr_,rws_=count_header_and_rows(wb['Stock Count'].iter_rows(values_only=True))
            if hdr_ is not None:
                n_=max(j for j,x in enumerate(hdr_) if x is not None)+1
                count_hdr=hdr_[:n_]; count=[list(r[:n_])+[None]*max(0,n_-len(r)) for r in rws_ if r and r[0]]
                hl_=[str(h_ or '').strip().lower() for h_ in count_hdr]
                # your names: 'Physical Count Qty' (C) = Today Stock Qty, 'Physical Count Qty_BOM' (G) = Today Qty_BOM (same columns, new names)
                ren_={'physical count qty':'Today Stock Qty','physical count qty_bom':'Today Qty_BOM'}
                count_hdr=[ren_.get(h_,count_hdr[i_]) if not (h_=='physical count qty' and 'today stock qty' in hl_) and not (h_=='physical count qty_bom' and 'today qty_bom' in hl_) else count_hdr[i_] for i_,h_ in enumerate(hl_)]
                hl_=[str(h_ or '').strip().lower() for h_ in count_hdr]
                fill_count_bom(count_hdr,count,path)                       # Today Qty_BOM = count x Quantity when Excel's result is missing
                ipc=hl_.index('today qty_bom') if 'today qty_bom' in hl_ else (6 if len(hl_)>6 else 2)      # column G: Physical Count Qty_BOM
                PHYS_COUNT.clear(); PHYS_DATE.clear()
                idt=next((i_ for i_,h_ in enumerate(hl_) if h_ in ('physical count date','count date')),None)    # your 'Physical Count Date' column (any position)
                for r_ in count:
                    c_=str(r_[0] or '').strip(); v_=r_[ipc] if ipc<len(r_) else None
                    if c_ and isinstance(v_,(int,float)) and not isinstance(v_,bool): PHYS_COUNT[c_]=PHYS_COUNT.get(c_,0.0)+float(v_)
                    if c_ and idt is not None and idt<len(r_) and isinstance(r_[idt],(dt.datetime,dt.date)):
                        PHYS_DATE[c_]=max(PHYS_DATE.get(c_,r_[idt]),r_[idt])
                pass
    return bom,count_hdr,count,src

MODPLAN_BY_PLANT={}    # (MIS BOM Name, customer, Plant line) -> modules per day, for the Module Plan sheet
def module_plan(mis_path, plant, dates):
    """Modules per day by MIS BOM Name and customer, from MIS Total Plan Req ('Value', one per line / customer / BOM / day)."""
    wb=wb_ro(mis_path); idx={pd.Timestamp(d).normalize():k for k,d in enumerate(dates)}
    seen=set(); mp={}; MODPLAN_BY_PLANT.clear()
    for r in wb['Total Plan Req'].iter_rows(min_row=2,values_only=True):
        if not tp_ok(r,plant) or not isinstance(r[0],(dt.datetime,dt.date)) or not isinstance(r[3],(int,float)): continue
        k=idx.get(pd.Timestamp(r[0]).normalize())
        if k is None: continue
        key=(r[0],r[1],r[2],r[4],r[17] if len(r)>17 else None)   # one plan line = date + production line + customer + BOM + OA
        if key in seen: continue
        seen.add(key); bn=str(r[4] or '').strip(); cu=' '.join(str(r[2] or '').split())
        mp.setdefault(bn,{}).setdefault(cu,[0.0]*len(dates))[k]+=r[3]
        pl=' '.join(str(r[1] or '').split())                    # MIS 'Plant' = production line (L1..L4, Arizona, OEM lines)
        MODPLAN_BY_PLANT.setdefault((bn,cu,pl),[0.0]*len(dates))[k]+=r[3]
    return mp

CSPM_ALT_QTY={}     # CS / PM alternative code -> its counted stock (added to the standard code; Stock Data / Movement Log show it)
CSPM_ALT={}         # CS / PM alternative code -> its standard code (PPC BOM 'Line #' = '<standard code>_Alt', e.g. PMPKM00000532 -> PMPKM00000004)
def cspm_alt_fold(bom):
    """CS / PM '_Alt' lines of your BOM: the alternative is the SAME material as its standard code (user 20-Nov-2026: 532 = alternate of
       004, shown separately it confuses). Its line is used INSTEAD of the standard line only when that customer / model has no standard
       line (never on top: it doubled the wrapping-roll demand), and its stock is added to the standard code (one row, like the busbar
       substitutes). Returns the BOM rows to plan with; fills CSPM_ALT."""
    CSPM_ALT.clear(); CSPM_ALT_QTY.clear(); std_={}; out=[]
    for d in bom:
        c=str(d.get('Component') or '').strip(); ln_=str(d.get('Line #') or '').strip()
        if c[:2].upper() in ('CS','PM') and ln_.lower().endswith('_alt') and ln_[:-4].strip() and ln_[:-4].strip()!=c: CSPM_ALT[c]=ln_[:-4].strip()
        else:
            if c[:2].upper() in ('CS','PM'): std_.setdefault((d.get('Customer'),str(d.get('Module Model') or '').strip()),set()).add(c)
            out.append(d)
    for d in bom:
        c=str(d.get('Component') or '').strip(); s_=CSPM_ALT.get(c)
        if s_ and s_ not in std_.get((d.get('Customer'),str(d.get('Module Model') or '').strip()),set()): out.append(dict(d,Component=s_))
    return out
def cspm_alt_stock(stock):
    """Stock of each CS / PM alternative code added to its standard code (same unit: your BOM / count unit)."""
    for a_,s_ in CSPM_ALT.items():
        if a_ in stock: CSPM_ALT_QTY[a_]=stock[a_]; stock[s_]=stock.get(s_,0.0)+stock.pop(a_)
    return stock
def cspm_from_bom(bom, count, modplan, dates, sr0, count_hdr=None, lineplan=None):
    """One block per CS / PM component: plan = average BOM qty per module (per MIS BOM Name) x modules that day."""
    nd=len(dates); per={}; bom=cspm_alt_fold(bom)
    for d in bom:
        c=str(d['Component'] or '').strip(); bn=str(d['MIS BOM Name'] or '').strip()
        if c[:2].upper() not in ('CS','PM') or not isinstance(d['Quantity'],(int,float)): continue
        x=per.setdefault(c,dict(desc=d['Component desciption'],uom=d['UoM'],q={}))
        if bn:
            mq=x['q'].setdefault(bn,{}); mdl=str(d['Module Model'] or '').strip()
            mq[mdl]=mq.get(mdl,0.0)+float(d['Quantity'])        # rows of one model add up; models are averaged below
    qi=count_qty_idx(count_hdr)
    stock=cspm_alt_stock({str(r[0]).strip():(float(r[qi]) if len(r)>qi and isinstance(r[qi],(int,float)) else 0.0) for r in count if r and r[0]})
    hl=[str(h).strip().lower() if h is not None else '' for h in (count_hdr or [])]
    ix=lambda n: hl.index(n) if n in hl else None
    i_q,i_u,i_f=ix('today stock qty'),ix('stock uom'),ix('quantity')
    raw={}
    for r in count:
        if not r or not r[0] or i_q is None: continue
        qv=r[i_q] if len(r)>i_q else None; u=r[i_u] if (i_u is not None and len(r)>i_u) else ''; fv=r[i_f] if (i_f is not None and len(r)>i_f) else None
        if isinstance(qv,(int,float)):
            raw[str(r[0]).strip()]=f"{qv:,.0f} {u}"+(f" x {fv:g}" if isinstance(fv,(int,float)) and fv not in (0,1) else '')
    out=[]; sr=sr0
    for c in sorted(per):
        x=per[c]; plan=[0.0]*nd; cust={}; cl={}; pls=[]
        for bn,mq in x['q'].items():
            q=sum(mq.values())/len(mq)
            for cu,v in (modplan.get(bn) or {}).items():
                arr=cust.setdefault(cu,[0.0]*nd)
                for k in range(nd): plan[k]+=q*v[k]; arr[k]+=q*v[k]
                for ln,lv in (((lineplan or {}).get(bn) or {}).get(cu) or {None:v}).items():
                    a_=cl.setdefault((cu,ln),[0.0]*nd)
                    for k in range(nd): a_[k]+=q*lv[k]
                    pls.append(dict(cu=cu,ln=ln,bn=bn,rate=q,mods=list(lv),src=f'your BOM {bn} (average of {len(mq)} model(s))'))
        sr+=1; desc=' '.join(str(x['desc'] or '').split())
        out.append(dict(key=f'{sr}CS/PM',sr=sr,short='CS - Consumable' if c.upper().startswith('CS') else 'PM - Packaging',common=desc,mat=c,desc=desc,mtype='CS/PM',
                        stock=stock.get(c,0.0),uom=x['uom'] or '',plan=plan,cust_plan=dict(sorted({k:v for k,v in cust.items() if any(v)}.items(),key=lambda kv:-sum(kv[1]))),
                        cust_line=dict(sorted({k:v for k,v in cl.items() if any(v)}.items(),key=lambda kv:(-sum(cust.get(kv[0][0],[0])),kv[0][0],str(kv[0][1])))),
                        transit=[0.0]*nd,adhoc=[0.0]*nd,openpo=[0.0]*nd,count_stock=True,raw=raw.get(c,'not on Stock Count')+''.join(f' + {raw[a_]} ({a_})' for a_,s_ in CSPM_ALT.items() if s_==c and a_ in raw),plan_lines=pls,
                        st_main=stock.get(c,0.0),st_wh=0.0,st_jw=0.0,plan_all={pd.Timestamp(dates[k]).normalize():plan[k] for k in range(nd) if plan[k]}))
    return out


