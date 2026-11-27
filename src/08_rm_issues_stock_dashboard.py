# ===================== RM ISSUES log (one issue per raw material at risk; your entries are carried over each run) =====================
RMI_PREV={}
def load_rmi_prev(paths):
    """Your RM Issues entries from the newest existing 4_Status (else 1_MRP of an older run; same-day re-runs included) - read before it is overwritten."""
    RMI_PREV.clear()
    dirs_={os.path.dirname(os.path.abspath(q)) for q in paths if q}          # since 19-Nov-2026 RM Issues lives in 4_Status (older runs: 1_MRP)
    dms_=sorted([os.path.join(d_,n_) for d_ in dirs_ if os.path.isdir(d_) for n_ in os.listdir(d_) if re.fullmatch(out_rx('dms')+r'_\d{4}-\d{2}-\d{2}\.xlsx',n_)],
                key=lambda f_: os.path.basename(f_),reverse=True)
    for p_ in [q for q in dms_+list(paths) if q and os.path.exists(q)]:
        try:
            pw=load_workbook(old_names_copy(p_),read_only=True,data_only=True)
            if 'RM Issues' not in pw.sheetnames: continue
            rows_=list(pw['RM Issues'].iter_rows(values_only=True))
            hi=next((i for i,r in enumerate(rows_[:8]) if r and 'Issue ID' in r),None)
            if hi is None: continue
            h_=list(rows_[hi])
            for r in rows_[hi+1:]:
                if r and r[0]: RMI_PREV[str(r[0])]=dict(zip(h_,r))
            if RMI_PREV: return
        except Exception as _e_: _swallow(4668,_e_); continue
RMI_COLS=['Issue ID','Date Raised','Project','Material','Supplier','PO Reference','Qty Required','Qty Received','Shortfall','Short from (stock = 0)','Required Date (at Factory)',
          'Revised ETA','Shipped qty','Pending qty','Delay (Days)','Production Impact (MW)','Priority','Owner','Status','Date Closed','Aging (Days)']      # no Action / Mitigation (user, Nov 2026)
def add_rm_issues_sheet(path, A, AL_ORDER, d0, nd, ships, prev_path=None):
    """RM Issues: an issue for every raw material that runs short or falls below the lead time in the next 16 weeks.
       Required date = the day cover drops below the lead time (stock 0 if earlier); Qty required = what is missing to keep lead-time days at the worst point;
       Qty received = PPC GRN since the issue was raised; Revised ETA / supplier / PO = the next PPC In_Transit shipment after the required date;
       Production impact (MW) = modules that cannot be built at the worst point x module Wp. Owner, Status, Date Closed and your own
       Action text are kept from the previous 1_MRP; an issue whose material is no longer at risk is closed automatically."""
    from openpyxl.styles import Border, Side
    from openpyxl.worksheet.table import Table, TableStyleInfo
    from openpyxl.worksheet.datavalidation import DataValidation
    wb=post_wb(path)
    if 'RM Issues' in wb.sheetnames: del wb['RM Issues']
    prev=dict(RMI_PREV)                                        # your entries, read at the START of the run (before any file is overwritten)
    HZ=min(nd,ALERT_HORIZON); dH=d0+pd.Timedelta(days=HZ)
    grn={}                                                     # PPC GRN by code and arrival date (Qty received since the issue was raised)
    try:
        for r in list(wb_ro(MIS_PATH_[0])['In Transit'].iter_rows(values_only=True))[1:]:
            if len(r)>39 and r[39] and isinstance(r[33],(int,float)) and r[33]>0 and isinstance(r[35],(dt.datetime,dt.date)):
                grn.setdefault(str(r[39]).strip(),[]).append((pd.Timestamp(r[35]).normalize(),float(r[33])))
    except Exception as _e_: _swallow(4689,_e_); pass
    # Production Plan inside the horizon, per (line, customer): modules, Wp x modules, MW  -> real Wp and an MW cap
    ppl={}
    try:
        for r in pp_rows_std(wb_ro(USER_MRP or SETTINGS_SRC[0])['Production Plan']):
            dd,ln,wp,cu,q,mw_,_,_=r
            if not isinstance(dd,(dt.datetime,dt.date)) or not isinstance(q,(int,float)) or q<=0: continue
            if not (d0<=pd.Timestamp(dd).normalize()<dH): continue
            cu_=' '.join(str(cu or '').split()); cu_=(LINKS.get(cu_) or {}).get('mis') or cu_   # plan name -> MIS customer (the names cust_line uses)
            e=ppl.setdefault((_line_key(str(ln or '')),_norm(cu_)),[0.0,0.0,0.0])
            e[0]+=q; e[1]+=q*(wp if isinstance(wp,(int,float)) and wp>0 else 0.0)
            e[2]+=(mw_ if isinstance(mw_,(int,float)) and mw_>0 else (q*wp/1e6 if isinstance(wp,(int,float)) and wp>0 else 0.0))
    except Exception as _e_: _swallow(4701,_e_); pass
    out=[]; seen=set()
    base=[x for x in AL_ORDER if not x.get('sup_row') and is_rm(x['b']['mat']) and not is_pm_like(x['b']['mat']) and not is_consumable(x['b']['mat']) and not x.get('no_demand')]   # tapes / barcode / back label = PM (user, Nov 2026)
    byc={x_['b']['mat']:x_ for x_ in A}
    for x in base:
        b=x['b']; m=b['mat']; S=x.get('S') or []
        P=b.get('plan') or []; mem=[m]; fx={m:1.0}
        if not S or not any(v>0 for v in P[:HZ]): continue
        cw=cover_watch(x,d0,ships,nd)
        t=cw['t'] if cw else None
        ri=risk_info(x,d0)
        if t is None and ri['k'] is None: continue
        t=min([v for v in (t,ri['k']) if v is not None])
        req_d=d0+pd.Timedelta(days=t)
        TGT=tgt_days(b.get('lt_used')); gap=max([max(0.0,sum(P[k+1:k+1+TGT])-S[k]) for k in range(t,min(HZ,len(S)))] or [0.0])
        if ri['k'] is None and gap<=0.5: continue            # below the lead time but the rest of the plan is covered and stock never hits 0: no issue (checker finding)   # missing to keep lead-time days of (pooled) demand at the worst point
        sh=sorted([s_ for s_ in (ships or {}).get(m,[]) if s_.get('raw') is not None],key=lambda s_:s_['raw'])
        nx=next((s_ for s_ in sh if s_['raw']>=req_d),None) or (sh[-1] if sh else None)
        iid=f"RMI-{m}"; pv=prev.get(iid,{})
        raised=pd.Timestamp(pv['Date Raised']).normalize() if isinstance(pv.get('Date Raised'),(dt.datetime,dt.date)) else d0
        if str(pv.get('Status') or '').strip().lower()=='closed' and isinstance(pv.get('Date Closed'),(dt.datetime,dt.date)): raised=d0   # re-opened: new date
        rec=sum(q*fx.get(c,1.0)/fx.get(m,1.0) for c in mem for dd,q in grn.get(c,[]) if raised<=dd<=d0)
        # production impact: modules that cannot be built at the worst point x Wp of the plan rows that use it, capped at the MW planned for them
        pairs=set()
        for c in mem:
            for (cu,ln),arr in ((byc.get(c) or {}).get('b',{}).get('cust_line') or {}).items():
                if any((arr[k] if k<len(arr) else 0) for k in range(HZ)): pairs.add((_line_key(str(ln or '')),_norm(cu)))
        rows_pp=[v for k_,v in ppl.items() if k_ in pairs] or [v for k_,v in ppl.items() if k_[0] in {l_ for l_,_ in pairs}]
        mods=sum(v[0] for v in rows_pp); wpq=sum(v[1] for v in rows_pp); mw_plan=sum(v[2] for v in rows_pp)
        wp_avg=(wpq/mods) if mods and wpq else 600.0                   # 600 Wp only when the plan has no Wp
        per_mod=(sum(P[:HZ])/mods) if mods else None
        worst=max(0.0,-min(S[:HZ])) if S else 0.0
        mw=(min(worst/per_mod*wp_avg/1e6,mw_plan) if (per_mod and mw_plan) else (worst/per_mod*wp_avg/1e6 if per_mod else None))
        slt_=short_lt_text(x,d0)
        prio='High' if (ri['cls']=='ALERT' or slt_[1]) else ('Medium' if ri['k'] is not None else 'Low')     # stock zero inside the lead time = High
        late=((nx['raw']-req_d).days if nx else None)
        st=pv.get('Status') if str(pv.get('Status') or '').strip() and str(pv.get('Status')).strip().lower()!='closed' else 'Open'
        proj=str(x.get('proj') or '').replace('\n',', ') or ', '.join(sorted({str(cu) for (cu,ln) in (b.get('cust_line') or {})}))[:60]
        same=[s_ for s_ in sh if nx and s_['raw']==nx['raw'] and s_.get('sup')==nx.get('sup') and s_.get('po')==nx.get('po')]
        out.append([iid,raised.to_pydatetime(),proj,f"{str(b.get('common') or '')} ({' + '.join(mem)})",(nx['sup'] if nx else 'no supplier on the way'),
                    (str(nx.get('po')) if nx and nx.get('po') else None),round(gap+rec),round(rec),round(gap),slt_[0],req_d.to_pydatetime(),(nx['raw'].to_pydatetime() if nx else None),
                    (round(sum(s_.get('shq') or 0 for s_ in same)) if nx else None),(round(sum(s_['qty'] for s_ in same)) if nx else None),
                    late,(round(mw,2) if mw is not None else None),prio,pv.get('Owner'),st,None,(d0-raised).days])
        seen.add(iid)
    for iid,pv in prev.items():                                # no longer at risk: closed automatically (kept for the record)
        if iid in seen: continue
        dr=pv.get('Date Raised'); dc=pv.get('Date Closed') if isinstance(pv.get('Date Closed'),(dt.datetime,dt.date)) else d0.to_pydatetime()
        out.append([iid,dr,pv.get('Project'),pv.get('Material'),pv.get('Supplier'),pv.get('PO Reference'),pv.get('Qty Required'),pv.get('Qty Received'),0,pv.get('Short from (stock = 0)'),
                    pv.get('Required Date (at Factory)'),pv.get('Revised ETA'),pv.get('Shipped qty'),pv.get('Pending qty'),pv.get('Delay (Days)'),pv.get('Production Impact (MW)'),pv.get('Priority'),pv.get('Owner'),
                    'Closed',dc,((pd.Timestamp(dc)-pd.Timestamp(dr)).days if isinstance(dr,(dt.datetime,dt.date)) else None)])
    # open issues by Production Impact (MW), highest first; issues without MW after them; closed issues at the bottom
    out.sort(key=lambda r_:(r_[18]=='Closed',r_[15] is None,-(r_[15] if isinstance(r_[15],(int,float)) else 0)))
    # ---- sheet
    ws=wb.create_sheet('RM Issues'); ws.sheet_view.showGridLines=False
    F=lambda **k: Font(name='Calibri',**k)
    ws['A1']='RM Issues — raw materials at risk of shortage'; ws['A1'].font=F(size=20,bold=True,color='14213D')
    ws['A2']=(f"One issue per raw material that runs short or falls below its lead time (days) in the next {HZW()}, sorted by Production Impact (MW) high to low (closed issues at the bottom), rebuilt every run. "
              "Required date = first day under the lead time (Inv. Days; or stock 0); Qty required = what is missing to keep lead-time days at the worst point; Qty received = PPC GRN since the issue was raised; "
              f"Revised ETA / supplier / PO = next PPC In_Transit shipment, Shipped qty / Pending qty = PPC Final Qty. / Final Pending Qty. of that shipment (same supplier, PO, ETA); Production impact = modules that cannot be built at the worst point x the Wp of the plan rows using it (capped at the MW planned for them in {HZW()}); a substitute group = ONE issue (group demand, stock and shipments). "
              "Yellow columns are yours (Owner, Status, Date Closed) - kept from the previous 1_MRP. Issues no longer at risk are closed automatically.")
    ws['A2'].font=F(size=9,color='8A949E'); ws.merge_cells('A2:U2'); ws['A2'].alignment=Alignment(wrap_text=True,vertical='top'); ws.row_dimensions[2].height=48
    H=list(RMI_COLS)
    for j,h in enumerate(H,1):
        c=ws.cell(4,j,h); c.font=F(size=10,bold=True,color='1F2D3D'); c.fill=PatternFill('solid',fgColor='F4F6F8')
        c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True); c.border=Border(bottom=Side(style='medium',color='C9D2DC'))
    ws.row_dimensions[4].height=32
    fm=[None,'dd-mmm-yy',None,None,None,None,'#,##0','#,##0','#,##0',None,'dd-mmm-yy','dd-mmm-yy','#,##0','#,##0','0','0.00',None,None,None,'dd-mmm-yy','0']
    PC={'High':('FCE8E6','C0392B'),'Medium':('FEF3E2','B9770E'),'Low':('FEF9E7','9A7D0A')}
    thin=Side(style='thin',color='E3E8EE')
    for i,row in enumerate(out,5):
        for j,v in enumerate(row,1):
            c=ws.cell(i,j,v); c.font=F(size=10,color='1F2D3D'); c.border=Border(bottom=thin)
            c.alignment=Alignment(vertical='center',wrap_text=(j in (3,10)),horizontal=('left' if isinstance(v,str) else 'right'))
            if fm[j-1]: c.number_format=fm[j-1]
            if j in (18,19,20): c.fill=PatternFill('solid',fgColor='FFFDE7')
        bg,fg=PC.get(row[16],('F2F4F6','5D6D7E')); ws.cell(i,17).fill=PatternFill('solid',fgColor=bg); ws.cell(i,17).font=F(size=10,bold=True,color=fg)
        if isinstance(row[14],(int,float)) and row[14]>0: ws.cell(i,15).font=F(size=10,bold=True,color='C0392B')
        if isinstance(row[9],str) and '→' in row[9]: ws.cell(i,10).font=F(size=10,bold=True,color='C0392B')
        if row[18]=='Closed':
            for j in range(1,22): ws.cell(i,j).font=F(size=10,color='A6ACAF')
        ws.row_dimensions[i].height=30
    for j,w_ in enumerate([22,11,26,30,24,13,11,11,11,30,12,11,11,11,8,11,9,14,10,11,8],1): ws.column_dimensions[L(j)].width=w_
    if out: ws.add_table(Table(displayName='RMIssues',ref=f'A4:U{4+len(out)}',tableStyleInfo=TableStyleInfo(name='TableStyleLight1',showRowStripes=False)))
    dv=DataValidation(type='list',formula1='"Open,In progress,Expedited,Waiting supplier,Closed"',allow_blank=True); ws.add_data_validation(dv)
    if out: dv.add(f'S5:S{4+len(out)}')
    ws.freeze_panes='E5'
    pos=(wb.sheetnames.index('Requirements')+1) if 'Requirements' in wb.sheetnames else 3
    wb.move_sheet('RM Issues',offset=pos-wb.sheetnames.index('RM Issues'))

def add_stock_check_sheet(path, AL_ORDER, d0, mlog=None):
    """Stock Check: SAP stock (MB52) next to the physical count and its date, and the stock the MRP uses today
       (count rolled forward by usage and receipts when the count is older than today). Highest shortage risk first."""
    from openpyxl.styles import Border, Side
    from openpyxl.worksheet.table import Table, TableStyleInfo
    wb=post_wb(path)
    if 'Stock Check' in wb.sheetnames: del wb['Stock Check']
    ws=wb.create_sheet('Stock Check'); ws.sheet_view.showGridLines=False
    F=lambda **k: Font(name='Calibri',**k); INK='1F2D3D'; MUTED='8A949E'; NAVY='14213D'
    rowb=Border(bottom=Side(style='thin',color='E3E8EE')); hb=Border(bottom=Side(style='medium',color='C9D2DC'))
    ws['A1']='Stock Check — SAP stock vs physical count'; ws['A1'].font=F(size=20,bold=True,color=NAVY)
    ws['A2']=('SAP stock = PPC sheet MB52 (else your MB52.xlsx) (Unrestricted + Quality Inspection; MIS MB52 only for plant name / final code / categories) · Physical count = PPC Count sheet, count date = the date column holding the count · '
              'when the count is older than today it is rolled forward: Estimated stock = count − usage per day × days since count + received since count (PPC GRN after the count date); '
              'usage per day = average Production Plan demand of the next 7 days (OEM plan left out); not counted: Estimated stock = SAP stock (MB52) · Used in MRP = the stock behind every shortage date in this workbook (= ALERT Estimated stock today) · '
              'a material with stock from 2+ suppliers = one row per supplier (Item name - supplier: MB52 batches / count by supplier; usage and receipts by the count share; Lines / Project / Usage / Module = the whole material).')
    ws['A2'].font=F(size=10,color=MUTED); ws.merge_cells('A2:T2'); ws['A2'].alignment=Alignment(wrap_text=True,vertical='top'); ws.row_dimensions[2].height=58
    H=['Code','Item name','Category','UOM','SAP stock (MB52)','Physical count','Count date','Days since count','Usage / day (next 7 days, roll-forward only)','Usage since count','Received since count',
       'Estimated stock today','Used in MRP (own stock)','Estimate − SAP','Source','Shortage risk',
       'Lines (share of plan volume)','Project Name (share of plan)','Usage per module by customer (source: your BOM / MIS Standard BOM)',
       'Module (MIS BOM, share of plan)']                   # Q..T moved here from the Action List (user, Nov 2026); Item name + Category = the ALERT (Nov 2026)
    for j,h in enumerate(H,1):
        c=ws.cell(4,j,h); c.font=F(size=10,bold=True,color=INK); c.fill=PatternFill('solid',fgColor='F4F6F8'); c.border=hb
        c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
    ws.row_dimensions[4].height=34
    if not MIS_CAT and MIS_PATH_[0]: load_mis_cat(MIS_PATH_[0])     # Category (item_category = the ALERT rule)
    base=[x for x in AL_ORDER if not x.get('sup_row')]
    supx={(x['b']['mat'],x['sup_row']):x for x in AL_ORDER if x.get('sup_row')}     # Action List supplier rows: customers allowed that vendor
    apos={}                                                        # sorted like its label: the ALERT order first, then the rest (OK / No need / No demand) as before
    for i_,r_ in enumerate(ALERT_ROWS): apos.setdefault(r_['code'],i_)
    base=sorted(base,key=lambda x:apos.get(x['b']['mat'],10**6))   # the ALERT order first
    def sup_label(m,nm,ri):                                        # a supplier row of a material on 2+ WSA sheets (333): the ALERT row of the sheet holding its stock
        rs_=[r_ for r_ in ALERT_ROWS if r_['code']==m and r_.get('sups') and nm in r_['sups']]
        if not rs_: return alert_label(m,ri)
        r_=rs_[0]; w_=r_.get('why') or {}; t=ALERT_CLS[r_['cls']][0]
        if r_.get('first') is not None: t+=f" {r_['first']:%d-%b}"
        elif r_['cls']==3 and w_.get('below') is not None: t+=f" from {w_['below']:%d-%b}"
        return t+(f" ({r_['where']})" if r_.get('where') else ''),r_['cls']
    r=5; nsup=0
    for x in base:
        b=x['b']; m=b['mat']; e=STOCK_EST.get(m) or {}; ri=risk_info(x,d0); ss=stock_src(b)
        own=ss['own']; est=ss['est']; sapv=ss['sap']; cnt=ss['cnt']
        nm0=str(b.get('common') or m)                              # = ALERT Item name (MIS Common Name)
        lines=[dict(nm=nm0,sap=sapv,cnt=cnt,est=est,own=own,sh=None,lab=alert_label(m,ri),x=x,src=e.get('src') or 'SAP (MB52)')]
        sp=sup_split(b)
        if sp:                                                     # do not combine suppliers (user, Nov 2026: glass FLAT and KIBING separated)
            ct_=sum(o['cnt'] or 0.0 for o in sp); et_=sum(o['est'] for o in sp)
            lines=[]
            for o in sp:
                sh_=((o['cnt'] or 0.0)/ct_ if (cnt is not None and ct_>0) else ((o['est']/et_) if et_>0 else 0.0))
                dn_=bool(vp_dead(m)[2] and vp_dead(m)[2](o['name']))                              # Settings ALLOWED VENDORS: no customer with demand may use this supplier
                lines.append(dict(nm=f"{nm0} - {o['name']}",sap=o['sap'],cnt=o['cnt'],est=o['est'],own=(0.0 if dn_ else o['est']),sh=sh_,
                                  rec=((b.get('sup_rec') or {}).get(o['name'],0.0) if b.get('sup_rec') is not None else None),   # a receipt goes to ITS OWN supplier (user 26-Nov-2026 'B yes')
                                  lab=sup_label(m,o['name'],ri),
                                  x=x,src=(e.get('src') or 'SAP (MB52)')+(' · count by supplier' if o['cnt'] is not None else ' · MB52 batches by supplier')
                                      +(f" · customers allowed this supplier: {supx[(m,o['name'])].get('custs')}" if supx.get((m,o['name'])) else '')))
            nsup+=1
        for ln in lines:
            sh_=ln['sh']; k_=(1.0 if sh_ is None else sh_)
            rolled=e.get('est') is not None                         # counted / rolled forward: Estimate − SAP shown (as before)
            vals=[m,ln['nm'],item_category(m),b.get('uom'),
                  (round(ln['sap']) if isinstance(ln['sap'],(int,float)) else None),(round(ln['cnt']) if ln['cnt'] is not None else 'no counting'),
                  (pd.Timestamp(e['date']).to_pydatetime() if e.get('date') is not None else None),e.get('days'),
                  (round(e['use_day']*k_,1) if e.get('use_day') else None),(round(e['use']*k_) if e.get('use') else None),(round(ln['rec'] if ln.get('rec') is not None else e['rec']*k_) if e.get('count') is not None else None),
                  round(ln['est']),round(ln['own']),
                  (round(ln['est']-ln['sap']) if (rolled and isinstance(ln['sap'],(int,float))) else None),ln['src'],
                  ln['lab'][0],
                  x.get('custs'),x.get('proj'),x.get('usage_src'),x.get('mods')]     # Q..T = the whole material (line / project shares)
            vals=vals[:16]+[(', '.join(map(str,v)) if isinstance(v,(list,tuple)) else v) for v in vals[16:]]
            fm=[None,None,None,None,'#,##0','#,##0','dd-mmm-yy','0','#,##0.0','#,##0','#,##0','#,##0','#,##0','+#,##0;-#,##0;0',None,None,None,None,None,None]
            for j,(v,f_) in enumerate(zip(vals,fm),1):
                c=ws.cell(r,j,v); c.font=F(size=10,color=INK); c.border=rowb
                c.alignment=Alignment(vertical='center',wrap_text=(j>=16),horizontal=('left' if j in (2,3,15,16) or isinstance(v,str) else 'right'))
                if f_: c.number_format=f_
            if isinstance(e.get('days'),int) and e['days']>0:
                for j in (7,8): ws.cell(r,j).fill=PatternFill('solid',fgColor='FEF3E2'); ws.cell(r,j).font=F(size=10,bold=True,color='B9770E')
                ws.cell(r,12).font=F(size=10,bold=True,color=INK)
            ci_=ln['lab'][1]; bg,fg=(ALERT_CHIP[ci_] if ci_ is not None else ('E3F4EC','1E8449'))   # = the ALERT chip colours
            ws.cell(r,16).fill=PatternFill('solid',fgColor=bg); ws.cell(r,16).font=F(size=10,bold=True,color=fg)
            d_=ws.cell(r,14).value
            if isinstance(d_,(int,float)) and isinstance(ln['sap'],(int,float)) and ln['sap'] and abs(d_)>0.1*abs(ln['sap']): ws.cell(r,14).font=F(size=10,bold=True,color='C0392B')
            nl_=max([str(v or '').count('\n')+1 for v in vals[15:]]+[1])
            ws.row_dimensions[r].height=max(19,min(15*nl_,150)); r+=1
    for j,w_ in enumerate([15,30,16,6,13,13,11,9,10,12,12,13,13,12,34,16,18,30,46,18],1): ws.column_dimensions[L(j)].width=w_
    if r>5: ws.add_table(Table(displayName='StockCheck',ref=f'A4:{L(len(H))}{r-1}',tableStyleInfo=TableStyleInfo(name='TableStyleLight1',showRowStripes=False)))
    ws.freeze_panes='C5'
    pos=wb.sheetnames.index(HEAT_NAME)+1 if HEAT_NAME in wb.sheetnames else 1
    wb.move_sheet('Stock Check',offset=pos-wb.sheetnames.index('Stock Check'))
    pass                                                           # saved once at the end (post_flush)






# ===================== ONE STYLE FOR THE WHOLE WORKBOOK + DEFINITIONS sheet =====================
DEFINITIONS=[
 ('ALERT','Alert','One simple list, strong alert (first tab of 1_MRP and of the Inbound file - the same rows)','STOCK-OUT IN LEAD TIME (red) = a stock-zero day before a new order can arrive: expedite / use an alternative · STOCK-OUT LATER (orange) = order by the date (first stock-zero day − lead time) · UNDER LEAD TIME (yellow) = under the lead time of stock but in transit covers the plan: watch the arrivals'),
 ('ALERT','Item name / Category','What the material is','Item name = MIS Common Name; a material and its Settings substitute = one row each with its own numbers (2nd line = when / how much the substitute covers, sub_cover), back glass 333 = one row per WSA sheet · Category = your MRP.xlsx Material Group name (column E) > MIS Standard BOM Understanding Name > your Material Group column A (Packaging / Label ...)'),
 ('ALERT','Inv. Days / Days of cover (+ all in transit)','Cover today','Stock today (MB52 / PPC count rolled forward, JW at the OEM left out) ÷ Usage/Day (module plan, average per production day over the transit lead time); incl. in transit = (stock + all in transit) ÷ Usage/Day. Listed when Inv. Days < 21 or a stock-zero day falls inside the lead time'),
 ('ALERT','Stock-zero days inside the lead time / Stock zero from → to','Days without stock','A stock-zero day = that day\'s need is not covered: stock + in transit arrived − plan < 0. Counted from today to today + lead time − 1; every period of the plan is listed, ⚠ = it starts inside the lead time. Back glass 333: one row per WSA sheet (customer project, own supplier stock + PPC arrivals)'),
 ('Substitute Advice','Usage/Day (module plan) / Days of cover (+ all in transit)','Module plan need','Usage/Day = module plan × usage per module, average per production day over the material\'s transit lead time (OEM left out; each substitute code on its own) - the same number as the ALERT; Days of cover (+ all in transit) = (stock + in transit) ÷ Usage/Day.'),
 ('All sheets','Usage / day (plan)','Production Plan demand','Modules in the Production Plan on that day × usage per module, per customer / line, OEM customers left out. Drives the projection: stock, shortage dates, planned orders, Inv. Days, risk lists. Plan usage/day of a count = average of the next 7 days.'),
 ('MRP Daily / Weekly','Inv. Days','How long the projected stock lasts','Projected stock at the end of the day (stock + in transit by arrival date − plan usage) ÷ Usage/Day from that day (average plan per production day over the material\'s transit lead time). Blank when no production in that window.'),
 ('All sheets','OEM','Left out of the calculations','Customers with OEM = Yes (CUSTOMER DAILY TARGET): their plan and the JW stock at the OEM are not part of any projection, shortage date, planned order or risk list. The JW stock stays visible as data (Stock Data "to OEM" lines, Movement Log, No Demand).'),
 ('In Transit','Shipment code check','Which material a shipment is','Description, Supplier Name, RM Code, Final RM Code and the codes on the PO line are compared; the code whose description fits the shipment clearly best (sizes, thickness, busbar / ribbon width, GSM, front / back) is used; blank / unknown codes are matched among the materials the same supplier ships. Every change is listed on Check.'),
 ('MRP Daily / MRP Weekly','Supplier rows','Glass per supplier','Shown only when 2 or more suppliers have stock, in transit or allowed demand for the material - otherwise the row would repeat the material row.'),
 ('RM Issues (4_Status)','Short From (Stock = 0)','Every stock-zero period inside the lead time','All periods where projected stock (counts / MB52 rolled forward + PPC in transit by arrival date − plan) is below zero that START inside the material lead time, each with dates and length ("beyond" = still short at the end of the shortage horizon); none inside: first stock-zero day marked "after lead time"; none at all: not short. Any period inside the lead time = High risk (RM Issues priority High).'),
 ('Inputs','SAP stock (MB52)','Where stock comes from','PPC file sheet MB52 (else your MB52.xlsx, newest MB52*.xlsx): Unrestricted + Quality Inspection = Total Stock; Plant name, Final Material Code, Dashboard Key, Company, Comman Name Cat.1 / Cat.2 are looked up from the MIS MB52 (else MIS Common Name). Stock date = save date of your MB52 file (older than today: rolled forward by usage)'),
 ('RM Issues (4_Status)','Issue log','One issue per raw material at risk',f'Material short or under the lead time in {HZW()}; Required date = first day under the lead time (or stock 0); Qty required = missing to keep lead-time days at the worst point; Qty received = PPC GRN since raised; Short from (stock = 0) = every period stock is below zero that starts inside the lead time (High priority); Revised ETA / Supplier / PO = next PPC In_Transit shipment; Shipped qty = PPC Final Qty. and Pending qty = MIS Final Pending Qty. of that shipment (same supplier, PO and ETA); Delay = ETA − required date; Production impact = modules not buildable at the worst point × Production Plan Wp, capped at the MW planned on those lines / customers in {HZW()}; substitute group = one issue; Owner / Status / Date Closed / Action kept from the previous run; auto-closed when no longer at risk'),
 ('Inputs','BOM and Stock Count','Where they come from','From the PPC workbook (sheets BOM and Count, newest file with both, name with PPC preferred); MRP.xlsx only when no PPC workbook is in the folder'),

 ('ALERT','Estimated stock today / SAP stock (MB52) / Physical count','Stock of each ALERT row','Estimated stock today = the Stock Check rule: physical count rolled forward to today (count − usage since the count + goods received), not counted: SAP stock (MB52); JW stock at the OEM left out · SAP stock (MB52) / Physical count = the Stock Check columns (count as counted; substitute member = its own code only; back glass 333 per WSA sheet = the suppliers whose stock is on that sheet)'),
 ('ALERT (Material Status columns)','In transit / Next arrival / Overdue in transit','Supply of each ALERT row','Stock = Estimated stock today, JW stock at the OEM left out · In transit = all pending shipments (PPC In_Transit; back glass 333 per WSA sheet: its PPC arrivals) · Next arrival = first arrival from today · Overdue = PPC pending qty whose arrival date has passed (333 per WSA sheet: in today\'s arrivals)'),
 ('Material Status','Rows','Which materials','EVERY material: raw materials, packaging and consumables, also the ones with no demand (brought back Nov 2026). Columns: UOM (after Category; 4_Status Materials Status other-material rows = PPC Count Base Unit of Measure; that file also has Item Category = PPC Count Item and Consumable Unit = PPC PM_CS_Usage), Stock, SAP stock (MB52), Physical count (= the Stock Check columns), In transit, Next arrival, Overdue in transit, Usage/Day, Short from, Inv. Days, Shortage risk, Days of cover (+ all in transit) (no INVENTORY / QUALITY columns since Nov 2026)'),
 ('ETD alert','Alert','WSA deliveries not on PPC In_Transit',f'Your WSA delivery rows (not cancelled) with an ETD whose BL# is not on PPC In_Transit (= not shipped as far as PPC In_Transit knows) · ARRIVES AFTER STOCK-OUT = expected arrival after the material\'s first stock-zero day (ALERT) · ETD PASSED · SHIPS WITHIN {ETD_SOON} DAYS · ON PLAN; the same rows as the Inbound file'),
 ('ETD alert','ETD vs today / Expected arrival / Delay','How late, when it can arrive','Late N d = base date − ETD; In N d = days until the ETD · expected arrival = your WSA plant date; ETD passed = base date + planned transit (WSA plant date − ETD), i.e. if it ships today · delay = expected arrival − WSA plant date · these rows are in no projection (the outputs use PPC In_Transit)'),
 ('Stock Check','SAP stock / Physical count / Count date','Both stock sources side by side','SAP = PPC sheet MB52 (else your MB52.xlsx) Unrestricted + Quality Inspection (MIS MB52 only for plant name, final code, key, company, categories) · count = PPC Count sheet; count date = the date column that holds the count'),
 ('ALERT','Below lead time from / Why (cause) / Days late / Arrival at plant',f'Which materials drop below the lead time in the next {HZW()}, and why',f'Below lead time from = first day projected stock covers < lead-time days of need · Why: LATE ARRIVAL = what is on the way would keep lead time or more but arrives after that day, LATE + NOT ENOUGH, NO SUPPLY · Days late = arrival of the first shipment after the below-lead-time day minus that day (late = would have kept lead-time cover if on time) · Arrival at plant (Final AtA) = passed-not-received total + the next arrival dates (PPC Final AtA) · Action second line = expedite / order more. UNDER LEAD TIME LATER (blue) = lead time or more today, below the lead time within {HZW()}, no stock-zero day. Materials with no demand are not listed'),
 ('MRP by Line','Usage/Day / Inv. Days at module plan / Below Lead Time From (module plan) / Days Late (module plan)','MODULE PLAN scenario (gold band)','Usage/Day = module plan × usage per module, average per production day over the transit lead time, OEM left out · Inv. days = (own stock + in transit) ÷ that need · Below Lead Time From / Days Late = same as the target columns but with the module-plan need = the ALERT values'),
 ('All sheets','No demand','Material without demand = no shortage status, no action','Needs at least one of: production requirement (plan x BOM usage) in the next shortage-horizon days (Settings, 60), daily target need, average planned demand. Otherwise: not Critical / At risk / Watch / Short, no Expedite / Order now / Release order, not on the ALERT'),
 ('BOM Material Dashboard - Weekly','Heatmap cells','When it runs short','One column per week = the daily number on that date (days the stock lasts from that date on the planned demand) = the same number as BOM Material Dashboard - Daily and the RM Material Dashboard on that date; red = short (below 1 day), orange / amber = below the lead time, blue = lead time or more, na = no demand that day; same rows as the ALERT sheet, grouped RM / PM / CS, highest risk first (same rule as every dashboard): first day the daily cover goes below 1 (stock-zero day), then below 5 days, then the lowest cover today'),
 ('Stock Check','Data dates','How old the inputs are','MIS = save date of the MIS file, WSA = report date of the planning file, Counts = oldest–newest count date; red when older than 2 days. Old MIS: SAP stock of materials that are not counted is rolled forward by usage per day × days since the MIS date.'),
 ('Stock Check','Current month: opening / GR / used / closing','Stock movement of this month','Opening = the first count of the month (PPC Count date columns; a material with several count rows needs every row counted); not counted: the stock of the first MRP run of the month (Movement Log, fills as you run daily) · GR qty = PPC In_Transit PPC GRN with GR. Date in this month (alternate codes converted) · Used = opening + GR after the opening date up to the closing date − closing (negative = a GR or count is missing) · Closing = the latest count (not counted: SAP stock today) · Projected = MRP projection (stock + in transit − module plan) at month end'),
 ('Stock Check','Shortage risk','The ALERT label','The same class as the ALERT sheet: STOCK-OUT IN LEAD TIME / STOCK-OUT LATER (+ first stock-zero day) · UNDER LEAD TIME · UNDER LEAD TIME LATER (+ the date it falls below the lead time) · OK = not on the ALERT (No demand / No need = no plan demand) · rows sorted in the ALERT order, then OK · back glass 333: one line per WSA sheet; a substitute member = its own ALERT row (EPE 480 / EPE 420 separately)'),
 ('Stock Check','Estimated stock today','Count rolled forward to today','Count − usage per day × days since count + received since count (PPC GRN with GR. Date after the count date); usage per day = average Production Plan demand of the next 7 days (OEM plan left out); not counted = SAP stock (MB52). This is the stock every shortage date uses (= ALERT Estimated stock today).'),
 ('Stock Check','Item name / Category / supplier rows','What the material is, stock per supplier','Item name / Category = the ALERT (MIS Common Name; item_category) · a material with stock from 2+ suppliers (e.g. glass FLAT / KIBING) = one row per supplier, "Item name - supplier": SAP = MB52 batches of that supplier (Batch Master), count = the PPC count by supplier (Supplier Code), estimated = the MRP Daily supplier row; usage / received since count split by the count share; Unassigned = stock MB52 / the count does not name'),
 ('RM Issues (4_Status)','Priority','Shortage priority','High = a stock-zero day inside the lead time · Medium = stock zero later (time to order) · Low = under the lead time at the module plan, no stock-zero day · OK · No demand. The stock lines behind the Stock column (source, source code, factor, storage location) are on the Stock Data sheet'),
 ('Stock Check','Lines (share of plan volume)','Which lines use the material','Share of the material\'s planned use over the plan horizon by production line (Production Plan × BOM usage)'),
 ('Stock Check','Project Name (share of plan)','Which projects use it','Same share by Project Name (Settings MIS LINK)'),
 ('Stock Check','Usage per module by customer','BOM usage behind the need','Usage per module per customer with its source: your PPC BOM sheet first, else MIS Standard BOM'),
 ('Stock Check','Module (MIS BOM, share of plan)','Module types using it','Short module name from the MIS BOM name with its share'),
 ('ALERT / Material Status','Usage/Day','Production view','Module plan (PPC Plan + Revised Plan) × BOM usage, average per production day over the transit lead time, OEM left out (no Reorder level / Status / On hold columns since Nov 2026)'),
 ('ALERT / Material Status','Lead time (days)','Order-to-arrival time','Settings TRANSIT LEAD TIME "Your lead time" (column F) if typed, else the WSA ETD → plant history (median), else the Settings fallback (RM 49, packaging 14)'),
 ('Stock Data / Movement Log','Stock / to OEM Stock','Stock here / at the OEM','MB52 (or physical count) / MIS RM Status "JW WSA to OEM" - the JW stock at the OEM is in no calculation'),
 ('ALERT / MRP by Line','Arrival at plant (Final AtA)','Next arrivals','Passed-not-received total, next 4 arrival dates, rest as one total'),
 ('ALERT / 4_Status / email / web page','Covered by substitute','Which code covers which days, with how much','Settings SUBSTITUTES between two MRP materials (EPE 420 <- EPE 480): every sheet shows each code SEPARATELY with its own stock, plan, in transit and arrivals (never pooled). Day by day (stock never below 0) the material uses its own stock + arrivals first; on a day they do not cover, the substitute\'s stock + arrivals left after its own plan cover the rest (factor; Use from / Use until respected). Text = Days of cover (stock only) per code, own cover to, substitute periods + quantities, stock-zero periods with the substitute. An alternative listed ALONE for a customer in your BOM carries that customer\'s demand.'),
 ('MRP by Line','Line stops because of','Which material stops the line','Code and name of the material(s) that are short on this line on the Line stops on date (the same for every row of the line); blank = the line runs to the end of the plan'),
 ('MRP by Line','Line can run (days) / Line stops on','How long the line runs','Earliest day a material it needs is short after sharing stock between lines (Line ranking, Allowed vendors, Settings Substitutes) - and after using your approved ALTERNATIVE (your BOM) for that line\'s customers: free stock of the alternative = what its own plan does not need within its lead time'),
 ('MRP by Line','First needed / Needed now','When this line needs it','First plan day the line uses it; Yes when within the material lead time'),
 ('MRP by Line','Plan on this line','Who / when on this line','Usage windows and customers in date order; "Same as today" = unchanged'),
 ('MRP by Line','This line: Usage/Day','This line\'s need','This line\'s module plan × usage per module, average per production day over the transit lead time (OEM left out)'),
 ('MRP by Line','In transit usable by this line','Supply this line may use','PPC in transit from suppliers allowed for this line (Settings ALLOWED VENDORS)'),
 ('MRP Daily','Gross requirement / Projected available balance','Day-by-day MRP record','Plan demand per day; balance = previous − requirement + in transit'),
 ('MRP Daily','Firm receipt / Planned order receipt','Open PO / order suggestion','Open PO without vessel; planned receipt keeps stock at target days (lead-time offset)'),
 ('Action List','Purpose','One row per material, highest shortage risk first','Same shortage-risk order as ALERT / RM Issues. Priority: High = stock zero inside the lead time, Medium = short later, Low = under the lead time, OK. Materials with no demand get no shortage status.'),
 ('Action List','Component / Component description / UOM','The material','Material code, MIS Common Name | description, unit of measure (MRP unit)'),
 ('Action List','Lines / Project Name / Usage per module by customer / Module (MIS BOM)','Who uses the material','Share of plan volume by production line and project; usage per module by customer (your PPC BOM first, else MIS Standard BOM); module type (MIS BOM name) and its share of the plan'),
 ('Action List','Short from → to (with in transit) / Below Lead Time From / Days Late','When the material runs short','Same logic as ALERT: projected stock (stock - plan + PPC in transit) below zero; first day cover is under the lead time; days between the shortage and the arrival at plant'),
 ('Action List','Material lead time (days) / Stock / to OEM Stock','Lead time and stock','Lead time = Settings TRANSIT LEAD TIME (your lead time) / WSA history / fallback; Stock = MB52 or PPC count rolled forward (JW at the OEM left out of calculations); to OEM Stock = JW stock at the OEM, shown only'),
 ('Action List','In transit today / Arrival at plant (Final AtA)','PPC In_Transit shipments not yet received','PPC In_Transit Final Pending Qty., next arrival dates (Final AtA; late lines counted today)'),
 ('Action List','Production Daily Target / Inv. Days at line daily target (without OEM, with OEM)','Cover at the customers\' daily targets','Target need per day = customer daily target x usage per module (Settings CUSTOMER DAILY TARGET); Inv. Days = (stock + in transit) ÷ target need per day, red under the lead time; "stock only" excludes transit; without OEM = customers with OEM = No, with OEM adds the OEM customer and JW stock at OEM'),
 ('Action List','Usage/Day / Inv. Days at module plan / Below Lead Time From (module plan) / Days Late (module plan)','Cover at the module plan','Average need per production day (module plan x usage per module, over the transit lead time, OEM left out); Inv. Days = (stock + in transit) ÷ that need; below-21 and days-late as on ALERT'),
 ('Action List','In transit / Open PO (WSA RM Planning) / Alternative BOM','Cross-checks','Quantities on the WSA planning sheets (in transit with a vessel, Open PO without vessel); alternative BOM = codes in your BOM that replace the standard code, with their stock and customers'),
 ('ALERT (formulas)','All columns','Same rows as ALERT, as live Excel formulas','Yellow = input, white = formula. Reads ALERT Calc (projection), ALERT Plan lines and ALERT Arrival lines; used to prove every ALERT row adds up. Column "Check vs ALERT" compares each calculated column with the ALERT sheet (OK / DIFF)'),
 ('ALERT Calc','Daily projection','Projection behind each ALERT row',f'Per ALERT row: Plan = SUMIFS of ALERT Plan lines, Arrivals = SUMIFS of ALERT Arrival lines (counted on max(arrival, today)); projected stock = previous day + arrivals - plan; stock zero = 1 when projected stock < 0; below the lead time = 1 when projected stock ÷ Usage/Day < 21 (first {HZW()})'),
 ('ALERT Plan lines','Demand (formula)','Plan behind ALERT Calc','One row per ALERT row x material x day x customer x production line: Modules x Qty per module x Factor (OEM customers left out; substitute members on their own rows; qty per module from your PPC BOM else MIS Standard BOM)'),
 ('ALERT Arrival lines','Qty (formula)','Shipments behind ALERT Calc','One row per ALERT row x shipment: PPC In_Transit pending qty x Factor, counted on MAX(arrival, today); material traced to your BOM / WSA PO + BL#'),
 ('In Transit Status','Status groups / shipment lines','PPC In_Transit by STATUS','Free-text PPC In_Transit STATUS grouped (In transit, At origin, At port, Held, Partly received, Transloaded, Delivered, Other, No status); pending qty = Final Pending Qty.; every shipment line listed below the summary'),
 ('No Demand','All columns','Stock or supply with nothing planned to use it','A = MRP materials with plan 0; B = other RM codes with MB52 / JW stock not in the MIS plan nor in your BOM for a planned model; C = cell codes (not planned here). Amber = more still coming'),
 ('Settings','All settings','Every planning rule in one place','Target days (A / B / C / packaging, + Z extra), line cover limit, daily targets, MIS LINK, allowed vendors, substitutes, batch supplier codes, line ranking. Read first on every run; copy the sheet into MRP.xlsx to keep your edits'),
 ('Module Plan','All columns','Modules per day by line and customer','From the PPC file: Plan with Revised Plan on top (a date + line in the Revised Plan replaces that date + line in Plan); MIS customer / BOM name = the Settings MIS LINK'),
 ('BOM','All columns','Copy of the BOM in the PPC file','Edit it in the PPC file, not here. MIS BOM Name comes from the module plan (customers using each Module Model). Used for CS / PM and for raw materials of customers it lists'),
 ('Stock Count','All columns','Copy of the PPC Count sheet','Today Stock Qty x Quantity = Today Qty_BOM; a physical count replaces MB52 for that material and is rolled forward to today (receipts - usage since the count date)'),
 ('How it works','Text','Short description of the current logic','Inputs, calculation order, sheet list and daily routine'),
 ('MIS Layout','Sheet / Column / Header','Headers of MIS columns read by position','Stored each run and compared with the next run; a changed header is reported on the Check sheet (MIS layout)'),
 ('MRP Weekly','All columns','Weekly view of MRP Daily','Live formulas on MRP Daily: Plan / In transit / Open PO = week totals; stock and Inv. Days = end of week'),
 ('Check','All columns','Discrepancies between MIS, the WSA planning file, your PPC file and the MRP','Severity High / Medium / Low; areas MIS layout, Stock, Plan (next 30 days), In transit (matched shipment by shipment on BL number), Open PO, Units, Master data, Your BOM. Difference = value B - value A. Low rows are collapsed (click + on the left)'),
]

def _web_hue(rgb):
    """Old colour (hex) -> 5_Web family: 'r' red, 'o' orange, 'g' green, 'b' blue, 'p' purple, 'k' dark ink, 'n' neutral grey."""
    import colorsys
    h_=str(rgb or '')[-6:]
    try: r_,g_,b_=(int(h_[i:i+2],16)/255 for i in (0,2,4))
    except ValueError: return 'n'
    hh,ss,vv=colorsys.rgb_to_hsv(r_,g_,b_); hh*=360
    if vv<0.38: return 'k'
    if ss<0.12: return 'n'
    if hh<20 or hh>340: return 'r'
    if hh<52: return 'o'
    if hh<170: return 'g'
    if hh<255: return 'b'
    return 'p'
def web_style_mrp(path, specs):
    """1_MRP sheets 1-5 (ALERT, Material Status, ETD alert, Action List, ALERT (formulas)) in the 5_Web Summary tab look (user 27-Nov-2026:
       "design sheet 1 to sheet 5 of 1_MRP like 5_Web summary tab style"): same palette / fonts as `_dms_weekly_style` (Segoe UI, white page,
       no grid, light grey row lines, one coloured group band over the table, muted grey header with a group-colour underline, red bold
       numbers, status chips = light red / grey, purple / blue text families). specs = {sheet: (header row, band row or None, title rows)}.
       Presentation only: values, formulas, rows, columns, number formats, merges, filters and print setup are not changed."""
    from openpyxl.styles import Border, Side, GradientFill
    wb=post_wb(path)
    WINK='15191A'; WMUT='6B7370'; WLINE='E4E7E4'; WRED='C9372C'; WPO='6A4FC7'; WSUB='2F6FC6'; FN='Segoe UI'
    FAM={'r':'C9372C','o':'F5A623','g':'1E8A3C','b':'3B9FD8','p':'6A4FC7','k':WINK,'n':'9AA29D'}
    TXT={'r':WRED,'o':'9A6A12','g':'1E8A3C','b':WSUB,'p':WPO,'k':WINK,'n':WMUT}
    KEY=(('STORE','b'),('PRODUCTION','g'),('COVER','g'),('WAREHOUSE','p'),('SUPPLY','p'),('TRANSIT LEAD','r'),('SHORTAGE','r'),('BELOW','o'),('INVENTORY','o'),('QUALITY','g'))
    line=Side(style='thin',color=WLINE); gap=Side(style='medium',color='FFFFFF')
    def rgb_of(cell):
        f=cell.fill
        return str(f.fgColor.rgb or '')[-6:].upper() if (f is not None and f.fill_type=='solid' and isinstance(f.fgColor.rgb,str)) else ''
    def font_of(cell):
        c=cell.font.color
        return str(c.rgb)[-6:].upper() if (c is not None and isinstance(c.rgb,str)) else ''
    for name,(hr,band,titles) in specs.items():
        if name not in wb.sheetnames: continue
        ws=wb[name]; last=ws.max_column; nc=range(1,last+1)
        ws.sheet_view.showGridLines=False; ws.sheet_properties.tabColor='7DC242'
        end=hr
        while end<ws.max_row and any(ws.cell(end+1,c).value is not None for c in (1,2,3)): end+=1      # table = rows until the first empty row
        grp={}                                                    # column -> group family (from the band row, running right)
        cur=('k','')
        for c in nc:
            if band:
                t=str(ws.cell(band,c).value or '').strip()
                if t:
                    fam=next((f_ for k_,f_ in KEY if k_ in t.upper()),None) or _web_hue(rgb_of(ws.cell(band,c)))
                    cur=(fam,t)
                grp[c]=cur
            else: grp[c]=('b','')
        if band:                                                  # band row: one coloured title per group (merged cells kept)
            for c in nc:
                x=ws.cell(band,c); fam,t=grp[c]; dark=(fam=='k')
                x.font=Font(name=FN,size=10,bold=True,color='FFFFFF')
                x.fill=GradientFill(type='linear',degree=0,stop=('3B9FD8','7DC242')) if dark else PatternFill('solid',fgColor=FAM[fam])
                x.alignment=Alignment(horizontal='center',vertical='center',wrap_text=False); x.border=Border(left=gap,right=gap)
            ws.row_dimensions[band].height=24
        for c in nc:                                              # header row: muted grey bold 9 pt on white, group-colour underline
            x=ws.cell(hr,c); fam=grp[c][0]
            x.font=Font(name=FN,size=9,bold=True,color=WMUT); x.fill=PatternFill('solid',fgColor='FFFFFF')
            x.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
            x.border=Border(bottom=Side(style='medium',color=FAM['b'] if fam=='k' else FAM[fam]))
        for r_ in titles:                                         # title / info lines above the header (ETD alert, Action List)
            for c in nc:
                x=ws.cell(r_,c)
                if x.value is None and not x.fill.fill_type: continue
                x.fill=PatternFill('solid',fgColor='FFFFFF')
                x.font=Font(name=FN,size=14 if r_==titles[0] else 9,bold=(r_==titles[0]),color=WINK if r_==titles[0] else WMUT)
        heads={c:str(ws.cell(hr,c).value or '').strip() for c in nc}
        for r_ in range(hr+1,end+1):
            for c in nc:
                x=ws.cell(r_,c); v=x.value; fl=rgb_of(x); fc=font_of(x)
                num=isinstance(v,(int,float)) and not isinstance(v,bool)
                fam_f=_web_hue(fc) if fc else 'k'
                bold=bool(x.font.b); col=TXT.get(fam_f,WINK) if fam_f!='k' else WINK; sz=9; fill='FFFFFF'; chip=False
                if fl and fl not in ('FFFFFF','000000'):
                    rr_,gg_,bb_=(int(fl[i:i+2],16) for i in (0,2,4))
                    if rr_+gg_+bb_<735:                           # a status chip (tints are lighter than this and go white)
                        fm=_web_hue(fl); fill,col=('F6C4C4',WRED) if fm=='r' else ('E9EDE9',TXT.get(fam_f,'26302A') if fam_f not in ('k','n') else '26302A'); bold=True; chip=True
                if heads[c] in ('Item name','Component description'): bold=True
                elif fam_f=='n' and not chip: col=WMUT
                x.font=Font(name=FN,size=sz,bold=bold,color=col); x.fill=PatternFill('solid',fgColor=fill); x.border=Border(bottom=line)
                x.alignment=Alignment(horizontal=x.alignment.horizontal,vertical='center',wrap_text=x.alignment.wrap_text)
        for r_ in range(end+1,ws.max_row+1):                      # boxes / notes under the table (Material Status): same fonts + web colours
            for c in nc:
                x=ws.cell(r_,c)
                if x.value is None and not x.fill.fill_type: continue
                fl=rgb_of(x); fam_f=_web_hue(font_of(x)) if font_of(x) else 'k'
                if fl and fl not in ('FFFFFF','000000'):
                    rr_,gg_,bb_=(int(fl[i:i+2],16) for i in (0,2,4)); fm=_web_hue(fl)
                    x.fill=PatternFill('solid',fgColor=FAM[fm] if rr_+gg_+bb_<600 else 'F4F6F4')
                dark_bg=bool(fl) and sum(int(fl[i:i+2],16) for i in (0,2,4))<600
                x.font=Font(name=FN,size=x.font.sz or 10,bold=x.font.b,color='FFFFFF' if dark_bg else TXT.get(fam_f,WINK))
        for r_ in range(hr+1,end+1):                              # keep the row heights (wrapped lines stay visible); min 19
            h_=ws.row_dimensions[r_].height
            ws.row_dimensions[r_].height=max(19,h_) if h_ else 19
    return wb
def unify_workbook_style(path):
    """Same look on every sheet: Calibri, navy titles and table headers, dark-grey text, no gridlines, navy tabs, no charts."""
    from openpyxl.styles import Border, Side
    from openpyxl.utils import range_boundaries
    wb=post_wb(path)
    NAVY='14213D'; INK='2C3E50'; thin=Side(style='thin',color='DDE3E8')
    # Definitions sheet
    if 'Definitions' in wb.sheetnames: del wb['Definitions']
    d=wb.create_sheet('Definitions')
    d['A1']='Definitions — every header and how it is calculated'; d['A2']='Sources: MIS RM Dashboard (MB52, RM Status; in transit = PPC In_Transit), WSA RM planning file, PPC file (Plan + Revised Plan = module plan, BOM = quantity per module, Count = physical stock, In_Transit), MB52.xlsx (SAP stock), MRP.xlsx (Settings, Batch Master), Customer_Tracker files.'
    for j,h in enumerate(['Sheet','Header','Definition','Calculation / source'],1): d.cell(4,j,h)
    for i,(a,b,c,e) in enumerate(DEFINITIONS,5):
        for j,v in enumerate((a,b,c,e),1):
            x=d.cell(i,j,v); x.alignment=Alignment(wrap_text=True,vertical='top')
    for col,w_ in zip('ABCD',[22,34,46,90]): d.column_dimensions[col].width=w_
    from openpyxl.worksheet.table import Table, TableStyleInfo
    d.add_table(Table(displayName='Definitions',ref=f'A4:D{4+len(DEFINITIONS)}',tableStyleInfo=TableStyleInfo(name='TableStyleLight1',showRowStripes=True)))
    hdr_rows={'Action List':[4],'MRP by Line':[4]}
    cut_=wb.sheetnames.index('In Transit Status') if 'In Transit Status' in wb.sheetnames else len(wb.sheetnames)
    hidden_=set(wb.sheetnames[cut_:])                                   # hidden sheets: no cell-by-cell restyle (saves time)
    HEAD='F4F6F8'; INK2='1F2D3D'; hb=Border(bottom=Side(style='medium',color='C9D2DC'),left=Side(style='thin',color='E3E8EE'),right=Side(style='thin',color='E3E8EE'))
    CARD={'green':('E3F4EC','1E8449'),'blue':('E8F1FB','2471A3'),'amber':('FEF3E2','B9770E'),'grey':('F2F4F6','5D6D7E'),'red':('FCE8E6','C0392B'),'teal':('E3F2F0','138D75')}
    def card_of(rgb):                                       # old banner colour -> card tint of the same family
        h=str(rgb or '')[-6:].upper()
        if h in ('548235','70AD47','375623','229954'): return CARD['green']
        if h in ('2E75B6','2F5597','1F4E79','2E86C1','2471A3'): return CARD['blue']
        if h in ('833C0B','C65911','ED7D31','D68910'): return CARD['amber']
        if h in ('C00000','C0392B'): return CARD['red']
        return CARD['grey']
    for ws in wb.worksheets:
        ws._charts=[]; ws._images=[]                                    # no charts anywhere
        ws.sheet_state='visible'
        ws.sheet_view.showGridLines=False; ws.sheet_properties.tabColor=('C00000' if ws.title=='ALERT' else NAVY)
        if ws.title in hidden_ and ws.title!='Definitions': continue
        big_=ws.title in ('MRP Daily','MRP Weekly')                     # written in Calibri already: no cell-by-cell font pass
        own=ws.title in ('ALERT',ALERT_FX,ALERT_CALC,ALERT_PL,ALERT_AL,'Material Status',HEAT_NAME,'Stock Check','RM Issues','Action List')   # already in house style
        if not own:
            for row in ([] if big_ else ws.iter_rows()):
                for c in row:
                    if c.value is None: continue
                    f=c.font; col=f.color.rgb if (f.color is not None and isinstance(f.color.rgb,str)) else None
                    c.font=Font(name='Calibri',size=(f.sz or 10),bold=f.b,italic=f.i,color=(col if col and col[-6:] not in ('000000',) else INK))
            t=ws.cell(1,1)
            if isinstance(t.value,str): t.font=Font(name='Calibri',size=18,bold=True,color=NAVY)
            rows=set(hdr_rows.get(ws.title,[]))
            for tb in ws.tables.values():
                c1,r1,c2,r2=range_boundaries(tb.ref); rows.add(r1)
                for j in range(c1,c2+1):
                    c=ws.cell(r1,j); c.fill=PatternFill('solid',fgColor=HEAD); c.font=Font(name='Calibri',size=10,bold=True,color=INK2)
                    c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True); c.border=hb
            for r in hdr_rows.get(ws.title,[]):
                for c in ws[r]:
                    if c.value is None: continue
                    c.fill=PatternFill('solid',fgColor=HEAD); c.font=Font(name='Calibri',size=10,bold=True,color=INK2)
                    c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True); c.border=hb
            if ws.title=='Action List':                         # row 3 group banners -> soft cards (same palette as the KPI cards)
                for c in ws[3]:
                    fg=c.fill.fgColor.rgb if c.fill and c.fill.fill_type=='solid' else None
                    if not fg or str(fg)[-6:] in ('FFFFFF','000000'): continue
                    bg,tx=card_of(fg); c.fill=PatternFill('solid',fgColor=bg); c.border=Border(top=Side(style='medium',color=tx))
                    if c.value is not None: c.font=Font(name='Calibri',size=10,bold=True,color=tx)
        code_re=re.compile(r'^(RM|PM|CS)[A-Z]*\d{5,}$')                 # every material-code cell: same format on every sheet
        cfont=Font(name='Calibri',size=10,bold=False,color='2C3E50'); cbord=Border(left=thin,right=thin,top=thin,bottom=thin)
        for row in ([] if ws.title in (ALERT_PL,ALERT_AL,HEAT_NAME) else ws.iter_rows()):      # the line sheets: plain data (speed); the daily dashboard keeps its dark heatmap look
            for c in row:
                if isinstance(c.value,str) and code_re.match(c.value.strip()):
                    c.font=cfont; c.border=cbord; c.alignment=Alignment(horizontal=c.alignment.horizontal,vertical='center',wrap_text=True)
        for tb in ws.tables.values():                           # table safety: header names unique, not blank, same as the table's column list
            c1,r1,c2,r2=range_boundaries(tb.ref); seen_=set()
            for k_,j in enumerate(range(c1,c2+1)):
                c=ws.cell(r1,j); v=str(c.value).strip() if c.value not in (None,'') else f'Column{k_+1}'
                while v in seen_: v=v+' '
                seen_.add(v); c.value=v
                if k_<len(tb.tableColumns): tb.tableColumns[k_].name=v
    d['A1'].font=Font(name='Calibri',size=18,bold=True,color=NAVY); d['A2'].font=Font(name='Calibri',size=10,italic=True,color='8A949E')
    for w_ in wb.worksheets: w_.sheet_state='visible'                # every sheet visible (your choice, Nov 2026)
    wb.active=0
    pass                                                           # saved once at the end (post_flush)







# ===================== RISK CHARTS: treemap (where the shortage is) + heatmap (when it runs short) - drawn with cells =====================

# ===================== STOCK MANAGEMENT dashboard (Store > Warehouse > Inventory > Production > Quality > Aging) =====================
def add_stock_mgmt_dashboard(path, A, AL_ORDER, mlog, ships, d0, nd, mis_path):
    """One row per raw material, grouped like a stock management system:
       STORE opening stock (previous MRP run) · WAREHOUSE stock in (PPC arrivals since then) / stock out (implied) / closing ·
       INVENTORY reorder level · PRODUCTION Usage/Day · QUALITY status + on hold (MB52 blocked + QI) ·
       AGING last arrival / days of supply / movement class / recommended action / stock value. Summary boxes + shortage watch below.
       Brought back Nov 2026 (user: "bring back stock management and include all materials not just raw materials") after it was merged
       into the ALERT: rows = EVERY material (RM + packaging + consumables, also no demand), same columns and rules as before; the ALERT
       keeps its Material Status columns."""
    from openpyxl.styles import Border, Side
    wb=post_wb(path)
    if 'Material Status' in wb.sheetnames: del wb['Material Status']
    ws=wb.create_sheet('Material Status'); ws.sheet_view.showGridLines=False; ws.sheet_properties.tabColor='14213D'
    F=lambda **k: Font(name='Calibri',**k)
    NAVY='14213D'; INK='1F2D3D'; MUTED='8A949E'; LINE='E3E8EE'
    STORE,WH,INV,PROD,QUAL,AGE=('2E86C1','D6EAF8'),('229954','D5F5E3'),('E67E22','FDEBD0'),('6C3483','E8DAEF'),('17A589','D1F2EB'),('1B2631','D6DBDF')
    def fill(c,col): c.fill=PatternFill('solid',fgColor=col)
    thin=Side(style='thin',color=LINE); bd=Border(left=thin,right=thin,top=thin,bottom=thin)
    widths=[5,16,18,14,17,11,11,11,10,10,11,11,11,11,10,11,9,10,12]
    for j,w_ in enumerate(widths,1): ws.column_dimensions[L(j)].width=w_
    # ---------- data ----------
    base=[x for x in AL_ORDER if not x.get('sup_row')]                 # EVERY material: RM, packaging, consumables (user, Nov 2026: 'include all materials')
    price={}; qi={}; blk={}
    mis=wb_ro(mis_path) if mis_path else None
    if mis:
        for r in list(mis['In Transit'].iter_rows(values_only=True))[1:]:
            if r and len(r)>39 and isinstance(r[8],(int,float)) and r[8]>0: price.setdefault(str(r[39] or '').strip(),[]).append(r[8])
        for r in list(mis['MB52'].iter_rows(values_only=True))[1:]:
            if r and len(r)>14 and r[2]:
                c=str(r[2]).strip(); m=(BOM_ALIAS[0] or {}).get(c,c)
                qi[m]=qi.get(m,0)+(r[12] if isinstance(r[12],(int,float)) else 0); blk[m]=blk.get(m,0)+(r[9] if isinstance(r[9],(int,float)) else 0)
    price={k:sum(v)/len(v) for k,v in price.items()}
    load_mis_cat(mis_path)                                       # category: item_category() (the ALERT uses the same)
    prev={}
    for e in sorted(mlog or [],key=lambda e:e['date']):
        if e['date']<d0: prev[e['code']]=e
    rows=[]
    for i,x in enumerate(base,1):
        b=x['b']; m=b['mat']; need=x.get('pl_need') or 0.0; lt=int(b.get('lt_used') or 0)   # module plan Usage/Day
        close=float(b['stock']); pe=prev.get(m); opening=float(pe['stock']) if pe else close
        since=pe['date'] if pe else d0
        sin=sum(s_['qty'] for s_ in ships.get(m,[]) if s_.get('raw') is not None and since<s_['raw']<=d0)
        sout=max(0.0,opening+sin-close)
        rop=need*lt+need*tgt_days(lt)
        status='Out of Stock' if close<=0 else ('Below reorder' if close<rop else 'Above reorder')
        if status=='Above reorder' and (x.get('wins') or []): status=f"Short {(d0+pd.Timedelta(days=x['wins'][0][0])):%d-%b}"     # never labelled safe when short
        hold=qi.get(m,0)+blk.get(m,0)
        last=max((s_['raw'] for s_ in ships.get(m,[]) if s_.get('raw') is not None and s_['raw']<=d0),default=None)
        dos=(max(0.0,close)/need) if need else None
        mv='No demand' if dos is None else ('Fast' if dos<=30 else ('Normal' if dos<=60 else ('Slow' if dos<=90 else 'Excess')))
        ri=risk_info(x,d0); k_=ri['k']; mv=ri['cls']                      # shortage-risk class (not stock ageing)
        val=((max(0.0,close)+(x.get('tr_tot') or 0))/need) if need else None      # days incl. in transit (no unit cost)
        sh_=ships.get(m,[])
        nxt_=min((s_['raw'] for s_ in sh_ if s_.get('raw') is not None and s_['raw']>=d0),default=None)
        ovd_=sum(s_['qty'] for s_ in sh_ if s_.get('raw') is not None and s_['raw']<d0)
        rows.append(dict(i=i,code=m,it=float(x.get('tr_tot') or 0),nxt=nxt_,ovd=ovd_,name=str(b.get('common') or m),cat=item_category(m),uom=b.get('uom'),
                         op=opening,sin=sin,sout=sout,close=close,rop=rop,need=need,status=status,hold=hold,last=last,dos=dos,mv=mv,val=val,k=k_,lt=lt,
                         alt=strip_code(str(AL_VAL.get((m,'Alternative BOM (yours): code - stock - customers')) or '').split('\n')[0]) or ''))
    # ---------- ALIGN WITH THE ALERT SHEET (user, Nov 2026: 'stock management sheet and alert sheet are not aligned') ----------
    # ONE source of truth: when a material is on the ALERT sheet, its numbers, class and short-from date come from the ALERT row, so the two
    # sheets can never disagree (before: Under-21-days used another rule, back glass 333 was one row here and one row per WSA sheet there).
    CLS_={0:'ALERT',1:'Short later',2:'Under lead time',3:'Under lead time later',4:'OK',5:'No demand'}
    def _from_alert(ar, i_, x_=None):
        nm=str(ar['name'])+(f"\n(WSA sheet: {ar['where']})" if ar.get('where') else '')
        k_=ar.get('k0'); need_=float(ar['need'] or 0.0)
        return dict(i=i_,code=ar['code'],it=float(ar.get('it') or 0.0),nxt=(pd.Timestamp(ar['nx'][0]) if ar.get('nx') else None),
                    ovd=(ar.get('ovd') if ar.get('ovd') is not None else (x_['ovd'] if x_ else 0.0)),ovd_txt=("in today's arrivals" if (ar.get('where') and ar.get('ovd') is None) else None),
                    name=nm,cat=item_category(ar['code']),uom=ar.get('uom'),op=float(ar['stock'] or 0.0),sin=0.0,sout=0.0,close=float(ar['stock'] or 0.0),
                    rop=float(ar['rop']),need=need_,status=ar['status'],hold=float(ar.get('hold') or 0.0),last=(x_['last'] if x_ else None),
                    dos=(float(ar['days']) if ar['days'] is not None else None),mv=CLS_[ar['cls']],val=(float(ar['days_it']) if ar['days_it'] is not None else None),k=k_,lt=int(ar['lt'] or 0),alt=(x_['alt'] if x_ else ''),ar=ar)
    AR={}
    for ar in ALERT_ROWS: AR.setdefault(ar['code'],[]).append(ar)
    new_rows=[]
    for x in rows:
        ars=AR.get(x['code']) or []
        if len(ars)>1 or (ars and ars[0].get('where')):               # split per WSA sheet (back glass 333): same rows as the ALERT
            for ar in ars: new_rows.append(_from_alert(ar,0,x))
        elif ars:                                                      # same material, same numbers as the ALERT row
            new_rows.append(_from_alert(ars[0],0,x))
        else:
            if x['mv'] in ('ALERT','Short later','Under lead time'):      # nothing on the ALERT -> it is not at risk on this sheet either
                x.update(mv='OK' if x['need'] else 'No need',k=None)
            new_rows.append(x)
    on_alert=sorted([y for y in new_rows if y.get('ar') or y['mv'] in CLS_.values()],key=lambda y:(list(CLS_.values()).index(y['mv']) if y['mv'] in CLS_.values() else 9,
                    y['k'] if y['k'] is not None else ((y['ar']['why']['t'] or 0) if y.get('ar') else 0),y['val'] if y['val'] is not None else 1e9,str(y['name'])))
    rest=[y for y in new_rows if not (y.get('ar') or y['mv'] in CLS_.values())]
    rows=on_alert+rest
    for i_,y in enumerate(rows,1): y['i']=i_
    # Stock SAP + Stock Count next to Stock (user, Nov 2026): SAP = MB52 stock of the material; Count = PPC physical count (not rolled forward);
    # Stock = what the MRP uses (count rolled forward to today, else SAP). Same rule as the Stock Check sheet. Rows per WSA sheet (333): not SAP / count -> '-'.
    # SAP stock (MB52) / Physical count = the Stock Check columns (user, Nov 2026; stock_src): ALERT row -> its SAP / count (substitute
    # group member shown with its own stock -> its own; back glass 333 per WSA sheet -> the suppliers of that sheet, else '-').
    byb_={x['b']['mat']:x['b'] for x in base}
    for y in rows:
        ar_=y.get('ar'); y['nosrc']=False
        if ar_ is not None:
            y['sap']=ar_.get('sap'); y['cnt']=ar_.get('cnt'); y['nosrc']=bool(ar_.get('where') and ar_.get('sups') is None)
        else:
            ss_=stock_src(byb_[y['code']]) if y['code'] in byb_ else dict(sap=None,cnt=None)
            y['sap']=ss_['sap']; y['cnt']=ss_['cnt']
    # rows 1-7 (title banner + process chevrons) removed (user request Nov 2026): the group band is row 1, the header row 2
    # ---------- table headers ----------
    groups=[('',1,5,NAVY),('STORE · stock',6,8,STORE[0]),('WAREHOUSE · supply',9,11,WH[0]),('INVENTORY',12,12,INV[0]),('PRODUCTION',13,13,PROD[0]),('QUALITY',14,15,QUAL[0]),('AGING ANALYSIS',16,19,NAVY)]
    for t,c1,c2,col in groups:
        if c2>c1: ws.merge_cells(start_row=1,start_column=c1,end_row=1,end_column=c2)
        c=ws.cell(1,c1,t); c.font=F(size=10,bold=True,color='FFFFFF'); c.alignment=Alignment(horizontal='center',vertical='center')
        for j in range(c1,c2+1): fill(ws.cell(1,j),col)
    ws.row_dimensions[1].height=20
    H=['Sr.','Item code','Item name','','Category','Stock','SAP stock (MB52)','Physical count','In transit','Next arrival','Overdue in transit','Reorder level','Usage/Day','Status','On hold','Short from','Days of cover (stock only)','Shortage risk','Days of cover (+ all in transit)']
    hc=[NAVY]*5+[STORE[0]]*3+[WH[0]]*3+[INV[0]]+[PROD[0]]+[QUAL[0]]*2+[NAVY]*4
    for j,(h,col) in enumerate(zip(H,hc),1):
        c=ws.cell(2,j,h); c.font=F(size=9,bold=True,color='FFFFFF'); fill(c,col); c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True); c.border=bd
    ws.merge_cells('C2:D2'); ws.row_dimensions[2].height=30
    CH={'Above reorder':('EAF2F8','1F618D'),'Below reorder':('FCF3CF','9A7D0A'),'Out of Stock':('FADBD8','C0392B'),
        'ALERT':('FADBD8','C0392B'),'Short later':('FDEBD0','B9770E'),'Under lead time':('FCF3CF','9A7D0A'),'Under lead time later':('EAF2F8','1F618D'),'OK':('D5F5E3','1E8449'),'No need':('F2F4F6','8A949E'),'No demand':('F2F4F6','8A949E')}
    r=3
    for x in rows:
        vals=[x['i'],x['code'],x['name'],None,x['cat'],round(x['close']),(round(x['sap']) if x.get('sap') is not None else '-'),(round(x['cnt']) if x.get('cnt') is not None else ('-' if x.get('nosrc') else 'no counting')),round(x['it']),(x['nxt'].to_pydatetime() if x['nxt'] is not None else 'none'),(x.get('ovd_txt') or (round(x['ovd']) if x['ovd'] else '-')),round(x['rop']),round(x['need'],1),
              x['status'],round(x['hold']) if x['hold'] else '-',((d0+pd.Timedelta(days=x['k'])).to_pydatetime() if x['k'] is not None else 'none'),(round(x['dos']) if x['dos'] is not None else '-'),x['mv'],(round(x['val']) if x['val'] is not None else '-')]
        fmts=[None,None,None,None,None,'#,##0','#,##0','#,##0','#,##0','dd-mmm','#,##0','#,##0','#,##0.0',None,'#,##0','dd-mmm-yy','0',None,'0']
        for j,(v,fm) in enumerate(zip(vals,fmts),1):
            c=ws.cell(r,j,v); c.font=F(size=10,color=INK); c.border=Border(bottom=thin)
            c.alignment=Alignment(horizontal=('left' if j in (2,3,5) else 'center'),vertical='center')
            if fm: c.number_format=fm
            if j in (6,7,8):  fill(c,'F4F9FD')
            if j in (9,10,11): fill(c,'F4FBF7')
            if j==12: fill(c,'FEF8F1')
            if j==13: fill(c,'F8F4FB')
            if j in (7,8) and isinstance(v,str): c.font=F(size=10,color=MUTED)
        ws.merge_cells(start_row=r,start_column=3,end_row=r,end_column=4)
        for j,key in ((14,x['status']),(18,x['mv'])):
            bg,fg=CH.get(key,('FDEBD0','B9770E')); fill(ws.cell(r,j),bg); ws.cell(r,j).font=F(size=10,bold=True,color=fg)
        if x['k'] is not None: ws.cell(r,16).font=F(size=10,bold=True,color='C0392B')
        if isinstance(vals[14],(int,float)): ws.cell(r,15).font=F(size=10,bold=True,color='B9770E')
        ws.row_dimensions[r].height=19; r+=1
    # ---------- summary boxes ----------
    r+=1; top=r
    def box(c1,c2,title,col,lines):
        ws.merge_cells(start_row=top,start_column=c1,end_row=top,end_column=c2)
        t=ws.cell(top,c1,title); t.font=F(size=11,bold=True,color='FFFFFF'); t.alignment=Alignment(vertical='center',indent=1)
        for j in range(c1,c2+1): fill(ws.cell(top,j),col)
        for i_,(a,v,fm,vc) in enumerate(lines,1):
            ws.merge_cells(start_row=top+i_,start_column=c1,end_row=top+i_,end_column=c2-1)
            x_=ws.cell(top+i_,c1,'•  '+a); x_.font=F(size=10,color=(vc if vc and fm is None and v is None else INK))
            y_=ws.cell(top+i_,c2,v); y_.font=F(size=11,bold=True,color=vc or NAVY); y_.alignment=Alignment(horizontal='right')
            if fm: y_.number_format=fm
            for j in range(c1,c2+1): fill(ws.cell(top+i_,j),'F7F9FB'); ws.cell(top+i_,j).border=Border(bottom=Side(style='thin',color='ECEFF3'))
    LRr=line_run_days([x for x in A if is_rm(x['b']['mat'])],d0,nd,ALLOC_STOP_)
    stop14=[l for l in LRr if l['t'][0]<LINE_LT.get(l.get('key'),LEAD_TIME_RM)*2/3]; first_stop=min(LRr,key=lambda l:l['t'][0]) if LRr else None   # line CRITICAL = under 2/3 of its longest material lead time
    shorts=[x for x in rows if x['k'] is not None]; f0=min(shorts,key=lambda x:x['k']) if shorts else None
    box(1,4,'⚠️  KEY RISK HIGHLIGHTS',NAVY,[('Short inside transit lead time (ALERT)',sum(1 for x in rows if x['mv']=='ALERT'),'0','C0392B'),('Short later (order in time)',sum(1 for x in rows if x['mv']=='Short later'),'0','B9770E'),('Under lead timeays (+ later)',sum(1 for x in rows if x['mv'] in ('Under lead time','Under lead time later')),'0','9A7D0A'),
        ('Earliest shortage: '+(f"{f0['name'][:16]} {(d0+pd.Timedelta(days=f0['k'])):%d-%b}" if f0 else 'none'),None,None,'C0392B' if f0 else '1E8449'),
        ('Lines CRITICAL (< 2/3 lead time)',len(stop14),'0','C0392B' if stop14 else '1E8449'),
        ('First line stop: '+(f"{first_stop['name']} {(d0+pd.Timedelta(days=first_stop['t'][0])):%d-%b}" if first_stop and first_stop['t'][0]<nd else 'none'),None,None,'C0392B' if (first_stop and first_stop['t'][0]<nd) else '1E8449')])
    b_=[0]*6                                                      # shortage timeline from the SAME rows as the table / ALERT sheet
    for x in rows:
        k=x['k']; b_[5 if k is None else (0 if k<=7 else (1 if k<=14 else (2 if k<=30 else (3 if k<=60 else 4))))]+=1
    tl=list(zip(['≤ 7 DAYS','8 – 14 DAYS','15 – 30 DAYS','31 – 60 DAYS','> 60 DAYS','NO SHORTAGE'],b_))
    box(6,9,'⏱  SHORTAGE TIMELINE (first short day)',NAVY,[(lab.title().replace('Days','days'),v,'0',('C0392B' if i_<2 else ('B9770E' if i_<4 else '9A7D0A'))) for i_,(lab,v) in enumerate(tl[:5])])
    al=[x for x in rows if x['k'] is not None and x['k']<x['lt']]
    box(11,19,'⚠️  SHORTAGE WATCH (short inside transit lead time)','C0392B',
        [(f"{x['name'][:26]} – short {(d0+pd.Timedelta(days=x['k'])):%d-%b}"+(f" · alt {x['alt'].split(':')[0][:24]}" if x['alt'] else ''),None,None,'C0392B') for x in al[:5]] or [('No material short inside its transit lead time',None,None,'1E8449')])
    # five cards
    top2=top+7
    cards=[('🏪 STORE','Short within 7 days',tl[0][1],'0',STORE),('🏭 WAREHOUSE','Lines stopping ≤ 14 days',len(stop14),'0',WH),
           ('⚙️ PRODUCTION','Materials short in the plan',len(shorts),'0',PROD)]      # no INVENTORY / QUALITY card (user, Nov 2026)
    spans=[(1,5),(6,10),(11,19)]                                  # 11..19 -> 11..16 after the 3 columns below are deleted
    for (t,lab,v,fm,(dark,light)),(c1,c2) in zip(cards,spans):
        for rr in range(top2,top2+4):
            ws.merge_cells(start_row=rr,start_column=c1,end_row=rr,end_column=c2)
            for j in range(c1,c2+1): fill(ws.cell(rr,j),dark if rr==top2 else light)
        a=ws.cell(top2,c1,t); a.font=F(size=12,bold=True,color='FFFFFF'); a.alignment=Alignment(horizontal='center',vertical='center')
        b_=ws.cell(top2+1,c1,lab); b_.font=F(size=10,color=dark); b_.alignment=Alignment(horizontal='center')
        n_=ws.cell(top2+2,c1,v); n_.font=F(size=24,bold=True,color=dark); n_.alignment=Alignment(horizontal='center',vertical='center'); n_.number_format=fm
        ws.row_dimensions[top2+2].height=34
    ws.cell(top2+5,1,'Every material (raw materials, packaging, consumables) in the risk order of the MRP · In transit = PPC pending · Overdue = PPC arrival date passed, not received · '
                     'Shortage risk: ALERT = short inside the transit lead time · Short later · Under lead time (no shortage in plan) · Days of cover (+ all in transit) = (stock + PPC in transit) ÷ Usage/Day.').font=F(size=8,italic=True,color=MUTED)
    ws.cell(top2+6,1,'Stock = the stock the MRP uses (physical count rolled forward to today, else SAP) = ALERT Estimated stock today · SAP stock (MB52) / Physical count = the Stock Check columns (count as counted; no counting = not counted; back glass 333 per WSA sheet = the suppliers of that sheet) · Why Stock differs: a counted material = Count − plan usage per day × days since the count + goods received since the count (never below 0); an uncounted one = SAP · count dates, usage and receipts per material: Stock Check sheet.').font=F(size=8,italic=True,color=MUTED)
    _del_cols(ws,[12,14,15])                                      # no INVENTORY (Reorder level) / QUALITY (Status, On hold) (user, Nov 2026: "delete quality and inventory")
    _add_uom_col(ws,[x.get('uom') for x in rows])                # UOM after Category (user, Nov 2026)
    wb.move_sheet('Material Status',offset=-wb.sheetnames.index('Material Status'))
    pass                                                           # saved once at the end (post_flush)

def _copy_sheet_to(src, dst):
    """Copy one finished sheet (values, styles, merges, widths, heights) into a sheet of ANOTHER workbook."""
    from copy import copy
    _sc={}        # SPEED: a cell's style is copied once per distinct style (the copies are never changed afterwards), not once per cell
    for row in src.iter_rows():
        for c in row:
            n=dst.cell(c.row,c.column,c.value)
            if c.has_style:
                k_=tuple(c._style); v_=_sc.get(k_)
                if v_ is None: v_=_sc[k_]=(copy(c.font),copy(c.fill),copy(c.border),copy(c.alignment),c.number_format)
                n.font,n.fill,n.border,n.alignment,n.number_format=v_
    for rng in src.merged_cells.ranges: dst.merge_cells(str(rng))
    for k,d in src.column_dimensions.items():
        if d.width: dst.column_dimensions[k].width=d.width
        if d.hidden: dst.column_dimensions[k].hidden=True
    for k,d in src.row_dimensions.items():
        if d.height: dst.row_dimensions[k].height=d.height
    dst.sheet_view.showGridLines=src.sheet_view.showGridLines; dst.sheet_view.zoomScale=src.sheet_view.zoomScale
    if src.sheet_properties.tabColor is not None: dst.sheet_properties.tabColor=copy(src.sheet_properties.tabColor)

DASH_CODES={}   # sheet title -> {row: the WSA sheet's material code(s)} (RM dashboards: Code / Item name columns)
DASH_MAP={}     # old row -> new row of the last Dashboard sorted by _wsa_dashboard_to (risk order)
def sort_wsa_dashboard(path, d0):
    """The aligned WSA file's own 'Dashboard' sorted highest risk first (user answer Nov 2026: 'Yes' - the same order as the RM Material
       Dashboard of 4_Status). XML level: the rows move with their formulas (each formula's own-row references, e.g. $A5,
       follow the row) and every unquoted 'Dashboard!B5' reference in other sheets (hidden Summary Dashboard) is moved too. Cached values
       stay (Excel recalculates on open). Returns the number of rows moved."""
    import zipfile, shutil, tempfile
    from openpyxl import Workbook as _WB
    _wsa_dashboard_to(path,_WB().active,d0)                     # works out today's values + the risk order (DASH_MAP)
    mp=dict(DASH_MAP)
    z=zipfile.ZipFile(path); wbx=z.read('xl/workbook.xml').decode(); rels=z.read('xl/_rels/workbook.xml.rels').decode()
    rid=re.search(r'<sheet[^>]*name="Dashboard"[^>]*r:id="([^"]+)"',wbx).group(1)
    tg=(re.search(r'Id="%s"[^>]*Target="([^"]+)"'%rid,rels) or re.search(r'Target="([^"]+)"[^>]*Id="%s"'%rid,rels)).group(1)
    dpath='xl/'+tg.replace('/xl/','').lstrip('/')
    ref_rx=re.compile(r'(?<![!\w$:])(\$?[A-Z]{1,3}\$?)(\d+)(?![\d(])')
    def _fix_f(m_):
        segs=re.split(r"('[^']*'|\"[^\"]*\"|&quot;.*?&quot;)",m_.group(2))      # sheet names / text in quotes are never cell references
        f=''.join(sg if k_%2 else ref_rx.sub(lambda q: q.group(1)+str(mp.get(int(q.group(2)),int(q.group(2)))),sg) for k_,sg in enumerate(segs))
        return m_.group(1)+f+m_.group(3)
    out={}
    x=z.read(dpath).decode('utf8')
    sd=re.search(r'<sheetData>(.*)</sheetData>',x,re.S)
    parts=re.findall(r'(<row [^>]*?r="(\d+)"(?:[^>]*/>|.*?</row>))',sd.group(1),re.S)
    newp=[]
    for txt,rn in parts:
        rn=int(rn); nr=mp.get(rn,rn)
        if nr!=rn:
            txt=re.sub(r'(<row [^>]*?r=")%d(")'%rn,r'\g<1>%d\2'%nr,txt,count=1)
            txt=re.sub(r'(<c r="[A-Z]{1,3})%d(")'%rn,r'\g<1>%d\2'%nr,txt)
            txt=re.sub(r'(<f[^>]*>)([^<]*)(</f>)',_fix_f,txt)
        newp.append((nr,txt))
    newp.sort(key=lambda t:t[0])
    out[dpath]=x[:sd.start(1)]+''.join(t for _,t in newp)+x[sd.end(1):]
    def _f1(m_):                                                 # Dashboard F1 = day 1 of the date row (user rule: base date = MB52 date)
        c_=m_.group(0); b_=pd.Timestamp(d0).normalize()
        if b_<pd.Timestamp(dt.date.today()).normalize():
            return re.sub(r'<f[^>]*>(?:TODAY\(\)|DATE\(\d+,\d+,\d+\))</f>(?:<v>[^<]*</v>)?',f'<f>DATE({b_.year},{b_.month},{b_.day})</f><v>{(b_-pd.Timestamp("1899-12-30")).days}</v>',c_)
        return re.sub(r'<f>DATE\(\d+,\d+,\d+\)</f>(?:<v>[^<]*</v>)?','<f ca="1">TODAY()</f>',c_)   # an earlier base-date run: live again
    if not mp:                                                   # rows already in risk order: only F1 may need the base date / TODAY() again
        x1=re.sub(r'<c r="F1"[^>]*>.*?</c>',_f1,x,count=1)
        if x1==x: z.close(); return 0
        out[dpath]=x1
    else: out[dpath]=re.sub(r'<c r="F1"[^>]*>.*?</c>',_f1,out[dpath],count=1)
    ext_rx=re.compile(r"(?<![\w'])(Dashboard!\$?[A-Z]{1,3}\$?)(\d+)(?!\d)")
    for n in z.namelist():                                       # other sheets pointing at Dashboard cells
        if n.startswith('xl/worksheets/') and n.endswith('.xml') and n!=dpath:
            t=z.read(n).decode('utf8')
            if 'Dashboard!' not in t: continue
            t2=ext_rx.sub(lambda q: q.group(1)+str(mp.get(int(q.group(2)),int(q.group(2)))),t)
            if t2!=t: out[n]=t2
    tmp=path[:-5]+'.sorting.xlsx'
    with zipfile.ZipFile(tmp,'w',zipfile.ZIP_DEFLATED) as zo:
        for it in z.infolist(): zo.writestr(it,out[it.filename].encode('utf8') if it.filename in out else z.read(it.filename))
    z.close()
    try:
        zipfile.ZipFile(tmp).testzip(); load_workbook(tmp,read_only=True)['Dashboard']   # read-back check before it replaces the file
        err_=_put_file(tmp,path)                                     # OneDrive / Excel lock: retry, write in place (checker finding Nov 2026)
        if err_ is not None: raise err_
    finally:
        if os.path.exists(tmp):                                      # never leave '<stem>.sorting.xlsx' behind (the next run could read it)
            try: os.remove(tmp)
            except OSError as e_: _swallow(0,e_)
    return len(mp)
CAT_COL=[False]
def _wsa_dashboard_to(wsa_path, dst, d0):
    """'RM Material Dashboard' = the Dashboard sheet of the daily WSA file (user, Nov 2026: "copy dashboard from daily wsa to daily
       material status"). Its formulas VLOOKUP the planning sheets of the WSA file, which another workbook cannot see, and its saved results
       are from the last time Excel calculated it - so every formula is worked out here the same way: date row 1 = TODAY() + 1 per column,
       IFERROR(VLOOKUP(date, '<sheet>'!A:.., column, 0), fallback) on the planning sheet's numbers (the script's own values for today on).
       Same layout: fonts / fills / borders / number formats, column widths, frozen pane, zoom, the 4 colour rules (NA grey, < 5 red,
       < 15 green, > 15 gold). Returns the number of materials, None when the file has no Dashboard."""
    import zipfile
    from copy import copy as _cp
    from openpyxl.formatting.rule import CellIsRule, FormulaRule
    def _styles_light(archive, wb):                              # read-only load WITH fonts / fills / borders (the speed patch drops them)
        from openpyxl.utils.indexed_list import IndexedList
        node=_oxl_fromstring(archive.read('xl/styles.xml'))
        for ch in list(node):
            if ch.tag.rsplit('}',1)[-1] not in ('numFmts','cellXfs','fonts','fills','borders'): node.remove(ch)
        ss=_OxlStylesheet.from_tree(node)
        wb._borders=IndexedList(ss.borders); wb._fonts=IndexedList(ss.fonts); wb._fills=IndexedList(ss.fills)
        wb._cell_styles=ss.cell_styles; wb._number_formats=ss.number_formats; wb._date_formats=ss.date_formats
        wb._timedelta_formats=ss.timedelta_formats; wb._alignments=ss.alignments; wb._protections=ss.protections
        return wb
    _prev=_oxl_reader.apply_stylesheet; _oxl_reader.apply_stylesheet=_styles_light
    try: wf=load_workbook(wsa_path,read_only=True)
    finally: _oxl_reader.apply_stylesheet=_prev
    if 'Dashboard' not in wf.sheetnames: return None
    wv=load_workbook(wsa_path,read_only=True,data_only=True); src=list(wf['Dashboard'].iter_rows())
    hdr={}; tab={}
    rx=re.compile(r"^=IFERROR\(VLOOKUP\(([A-Z]+)\$1,'?(.+?)'?!\$A:\$[A-Z]+,\$A(\d+)(-1)?,0\),\"(.*?)\"\)$")
    def table(nm):
        if nm not in tab:
            t_={}
            if nm in wv.sheetnames:
                for r_ in wv[nm].iter_rows(values_only=True):
                    if r_ and isinstance(r_[0],(dt.datetime,dt.date)): t_.setdefault(pd.Timestamp(r_[0]).normalize(),r_)
            tab[nm]=t_
        return tab[nm]
    nmat=0; buf={}; rsheet={}
    for row in src:
        for c in row:
            if not hasattr(c,'column') or c.column is None: continue
            v=c.value; out=v
            if c.row==1 and isinstance(v,str) and v.startswith('='):
                if v.upper()=='=TODAY()' or re.match(r'^=DATE\(\d+,\d+,\d+\)$',v.upper()): out=pd.Timestamp(d0).to_pydatetime()   # DATE() = a base-date run (F1)
                else:
                    m_=re.match(r'^=([A-Z]+)1\+(\d+)$',v)
                    out=(pd.Timestamp(hdr[m_.group(1)])+pd.Timedelta(days=int(m_.group(2)))).to_pydatetime() if m_ and m_.group(1) in hdr else None
                hdr[c.column_letter]=out
            elif isinstance(v,str) and v.startswith('='):
                m_=rx.match(v); out=None
                if m_:
                    rsheet.setdefault(c.row,m_.group(2).replace('&amp;','&'))
                    a_=src[c.row-1][0].value
                    idx=(int(a_) if isinstance(a_,(int,float)) else None)
                    if idx is not None and m_.group(4): idx-=1
                    fb=m_.group(5) or None; dd_=hdr.get(m_.group(1)); out=fb
                    if idx and dd_ is not None:
                        r_=table(m_.group(2)).get(pd.Timestamp(dd_).normalize())
                        if r_ is not None and idx-1<len(r_):
                            x_=r_[idx-1]
                            out=fb if (isinstance(x_,str) and x_.startswith('#')) else (0 if x_ is None else x_)
            buf.setdefault(c.row,[]).append((c,out))
        if row and getattr(row[0],'row',1)>1 and isinstance(getattr(row[0],'value',None),(int,float)): nmat+=1
    # highest risk on top (user, Nov 2026: "All dashboard, sort from highest risk to lowest"): first day the cover (days) goes below 1 (stock-zero day),
    # then below 5, then the lowest 'Today's Morning stock'; rows with no plan last; row 1 (dates) and rows without a material stay put
    def _risk(rn):
        vs=[o for c,o in buf[rn] if c.column>=6]; nums=[v for v in vs if isinstance(v,(int,float)) and not isinstance(v,bool)]
        dv=next((o for c,o in buf[rn] if c.column==4),None)
        if not nums and not isinstance(dv,(int,float)): return (1,0,0,0)
        neg=next((i for i,v in enumerate(vs) if isinstance(v,(int,float)) and v<1),10**6)     # stock-zero day (below 1), as the BOM dashboards
        lo=next((i for i,v in enumerate(vs) if isinstance(v,(int,float)) and v<RISK_DAYS),10**6)
        return (0,neg,lo,dv if isinstance(dv,(int,float)) else 1e9)
    data=[rn for rn in sorted(buf) if rn>1 and isinstance(next((o for c,o in buf[rn] if c.column==1),None),(int,float))]
    to={rn:rn for rn in buf}
    for new_,old_ in zip(data,sorted(data,key=_risk)): to[old_]=new_
    DASH_MAP.clear(); DASH_MAP.update({o_:n_ for o_,n_ in to.items() if o_!=n_})
    for rn,cells in buf.items():
        for c,out in cells:
            n_=dst.cell(to[rn],c.column,out)
            if c.has_style:
                n_.font=_cp(c.font); n_.fill=_cp(c.fill); n_.border=_cp(c.border); n_.alignment=_cp(c.alignment); n_.number_format=c.number_format
    # Category in column C (user, Nov 2026: "rm material dashboard add category") = item_category of the WSA sheet's material code(s)
    if not MIS_CAT and MIS_PATH_[0]: load_mis_cat(MIS_PATH_[0])
    sh_codes={}
    for g_ in (WSA_REFS[0] or []): sh_codes.setdefault(str(g_.get('sheet')),[]).extend(g_.get('codes') or [])
    h_=dst.cell(1,3,'Category'); h_._style=_cp(dst.cell(1,2)._style)
    mrp_m={str(x_['b']['mat']).strip() for x_ in (AL_ORDER_ or []) if not x_.get('sup_row')}      # the sheet's MRP material(s) only (not old / alias codes; checker finding)
    for rn in data:
        cl_=sh_codes.get(rsheet.get(rn,''),[]); cm_=[c_ for c_ in cl_ if c_ in mrp_m] or cl_
        cs_=list({str(k_).strip().lower():str(k_).strip() for k_ in reversed([item_category(c_) for c_ in cm_]) if k_ not in (None,'','-')}.values())[::-1]   # one per name, any case
        k_=dst.cell(to[rn],3,' / '.join(cs_) or '-'); k_._style=_cp(dst.cell(to[rn],2)._style); k_.alignment=Alignment(horizontal='left',vertical='center')
        DASH_CODES.setdefault(dst.title,{})[to[rn]]=list(dict.fromkeys(str(c_).strip() for c_ in cm_))     # code + item name columns of the dark design
    CAT_COL[0]=True
    try:                                                            # column widths, frozen pane, zoom, colour rules from the sheet's own XML
        z_=zipfile.ZipFile(wsa_path); wbx=z_.read('xl/workbook.xml').decode(); rels=z_.read('xl/_rels/workbook.xml.rels').decode()
        rid=re.search(r'<sheet[^>]*name="Dashboard"[^>]*r:id="([^"]+)"',wbx).group(1)
        tg=(re.search(r'Id="%s"[^>]*Target="([^"]+)"'%rid,rels) or re.search(r'Target="([^"]+)"[^>]*Id="%s"'%rid,rels)).group(1)
        x_=z_.read('xl/'+tg.replace('/xl/','').lstrip('/')).decode()
        for a_,b_,w_ in re.findall(r'<col min="(\d+)" max="(\d+)" width="([\d.]+)"',x_):
            for j_ in range(int(a_),min(int(b_),dst.max_column)+1): dst.column_dimensions[L(j_)].width=float(w_)
        for r_,h_ in re.findall(r'<row r="(\d+)"[^>]*?ht="([\d.]+)"[^>]*customHeight="1"',x_): dst.row_dimensions[int(r_)].height=float(h_)
        dst.column_dimensions['C'].width=18                       # Category column (was a thin spacer)   # row heights (rotated dates in row 1)
        if re.search(r'<sheetView[^>]*showGridLines="0"',x_): dst.sheet_view.showGridLines=False                # gridlines as in the source
        pn=re.search(r'<pane xSplit="(\d+)" ySplit="(\d+)"',x_)
        if pn: dst.freeze_panes=f"{L(int(pn.group(1))+1)}{int(pn.group(2))+1}"
        zm=re.search(r'<sheetView[^>]*zoomScale="(\d+)"',x_)
        if zm: dst.sheet_view.zoomScale=int(zm.group(1))
        sq=re.search(r'<conditionalFormatting sqref="([^"]+)"',x_)
        rng=sq.group(1) if sq else f"D2:{L(dst.max_column)}{dst.max_row}"
    except Exception as ex_:
        print(f'NOTE: Dashboard layout not fully copied ({type(ex_).__name__}: {ex_})'); rng=f"D2:{L(dst.max_column)}{dst.max_row}"
    tl=rng.split()[0].split(':')[0]
    dst.conditional_formatting.add(rng,FormulaRule(formula=[f'NOT(ISERROR(SEARCH("NA",{tl})))'],fill=PatternFill('solid',bgColor='BFBFBF'),font=Font(color='BFBFBF'),stopIfTrue=True))
    dst.conditional_formatting.add(rng,CellIsRule(operator='lessThan',formula=['5'],fill=PatternFill('solid',bgColor='FF0000'),font=Font(color='000000'),stopIfTrue=True))
    dst.conditional_formatting.add(rng,CellIsRule(operator='lessThan',formula=['15'],fill=PatternFill('solid',bgColor='00B050'),font=Font(color='006100'),stopIfTrue=True))
    dst.conditional_formatting.add(rng,CellIsRule(operator='greaterThan',formula=['15'],fill=PatternFill('solid',bgColor='BF8F00'),font=Font(color='FFFFFF')))
    return nmat
def _drop_card_rows(w, cr, n=4):
    """Remove the n-row card block that starts at row cr (+ the blank row after it): its merges and cells go, the rows below move up."""
    for m_ in list(w.merged_cells.ranges):
        if cr<=m_.min_row<=cr+n-1: w.unmerge_cells(str(m_))
    for k_ in [k_ for k_ in w._cells if cr<=k_[0]<=cr+n]: del w._cells[k_]
    mr_=w.max_row
    if mr_>cr+n:
        for m_ in [m_ for m_ in list(w.merged_cells.ranges) if m_.min_row>cr+n]:
            w.unmerge_cells(str(m_)); w.merge_cells(start_row=m_.min_row-n,start_column=m_.min_col,end_row=m_.max_row-n,end_column=m_.max_col)
        hs={r_:w.row_dimensions[r_].height for r_ in range(cr+n+1,mr_+1)}
        w.move_range(f"A{cr+n+1}:{L(w.max_column)}{mr_}",rows=-n)
        for r_ in range(cr,mr_+1): w.row_dimensions[r_].height=None
        for r_,h_ in hs.items(): w.row_dimensions[r_-n].height=h_
    else:
        for r_ in range(cr,mr_+1): w.row_dimensions[r_].height=None
def _remap_rows(w, f):
    """Move / delete whole rows of a finished sheet: f(old row) -> new row, or None = delete. Cells (value + style), merges (all rows of a
       merge must land on consecutive rows, else the merge is dropped), row heights, freeze panes, autofilter and print title rows follow."""
    import re as _re
    cells=[(k_,c_) for k_,c_ in w._cells.items()]
    merges=[(m_.min_row,m_.min_col,m_.max_row,m_.max_col) for m_ in w.merged_cells.ranges]
    hs={r_:d_.height for r_,d_ in w.row_dimensions.items() if d_.height}
    for m_ in list(w.merged_cells.ranges): w.unmerge_cells(str(m_))
    w._cells={}
    for r_ in list(w.row_dimensions.keys()): w.row_dimensions[r_].height=None
    for (r_,c_),cell in cells:
        n_=f(r_)
        if n_ is None: continue
        cell.row=n_; w._cells[(n_,c_)]=cell
    for r_,h_ in hs.items():
        n_=f(r_)
        if n_ is not None: w.row_dimensions[n_].height=h_
    for a,b,c,d in merges:
        na,nc=f(a),f(c)
        if na is not None and nc is not None and nc-na==c-a: w.merge_cells(start_row=na,start_column=b,end_row=nc,end_column=d)
    def _cell(ref):
        m_=_re.match(r'^([A-Z]+)(\d+)$',ref or '')
        if not m_: return ref
        n_=f(int(m_.group(2))); return f"{m_.group(1)}{n_}" if n_ else None
    if w.freeze_panes: w.freeze_panes=_cell(w.freeze_panes)
    if w.auto_filter.ref and ':' in w.auto_filter.ref:
        a_,b_=w.auto_filter.ref.split(':'); a2,b2=_cell(a_),_cell(b_)
        w.auto_filter.ref=f"{a2}:{b2}" if a2 and b2 else None
    pt=w.print_title_rows
    if pt:
        r1,r2=[int(x) for x in _re.findall(r'\d+',pt)][:2]
        keep=[f(x) for x in range(r1,r2+1) if f(x) is not None]
        w.print_title_rows=f"{min(keep)}:{max(keep)}" if keep else None
IT_SUP_HEAD='In transit supplier name (qty)'; IT_CODE=re.compile(r'^(JW)?[A-Z]{1,8}\d{4,}$')     # any material code (RM / PM / CS)
def _sup_short(n):
    """Supplier name without the legal form ('FLAT (HONG KONG) CO., LIMITED' -> 'FLAT (HONG KONG)')."""
    t=' '.join(str(n or '').split())
    t=re.sub(r'[,.]?\s*(CO\.?,?\s*(LTD|LIMITED)\.?|COMPANY LIMITED|LIMITED|LTD\.?|SDN\.?\s*BHD\.?|INC\.?|LLC|CORP(ORATION)?\.?)\s*$','',t,flags=re.I)
    return t.strip(' ,.') or '?'
def it_sup_text(codes, sheet=None, names=False):
    """In transit by supplier of a 4_Status row (user, Nov 2026: "Add in transit supplier name in all sheet of daily material
       status"): 'SUPPLIER qty; ...' (largest first). A material on 2+ WSA sheets (333) with its sheet = the PPC In_Transit lines of the
       suppliers shown in that sheet's column U (= the sheet's 'In transit'); else the PPC In_Transit shipments of the code (substitute
       group row: every member, in this code's unit; any other row: the code only) = the 'In transit' column."""
    by={}
    if sheet:
        g=next((g_ for g_ in (WSA_REFS[0] or []) if sheet in (g_.get('sheet'),g_.get('label'))),None)
        nm=(g or {}).get('ppcsup_names') or set()
        for x in ((CAP[0] or {}).get('ppc') or {}).get('rows') or []:
            sn=' '.join(str(x.get('supplier') or '').split())
            if x.get('mat') in codes and sn in nm and x.get('qty_mrp'): by[sn]=by.get(sn,0.0)+float(x['qty_mrp'])
    else:
        for c in codes:
            for s_ in SHIPS_.get(c,[]):
                if s_.get('qty') and s_['qty']>0: by[s_.get('sup') or '?']=by.get(s_.get('sup') or '?',0.0)+float(s_['qty'])
    agg={}
    for n,q in by.items(): agg[_sup_short(n)]=agg.get(_sup_short(n),0.0)+q
    if names: return [n for n,q in sorted(agg.items(),key=lambda t:-t[1]) if q>0.5]      # supplier names only (daily email)
    return '; '.join(f"{n} {q:,.0f}" for n,q in sorted(agg.items(),key=lambda t:-t[1]) if q>0.5) or 'none'
def _email_suppliers(r, bom_sup):
    """Supplier name(s) of an ALERT row for the daily email (user, Nov 2026: "put supplier names in the critical table"): the suppliers
       of its shipments in transit (333 per WSA sheet = that sheet's suppliers), else the supplier on
       your PPC BOM (the one you order from), else the stock suppliers of the MRP supplier rows (glass / materials with 2+ suppliers),
       else the Shipper typed on your WSA delivery rows; 'no supplier on file' when none is known."""
    codes=[r['code']]
    try: n_=it_sup_text([r['code']],sheet=r['where'],names=True) if r.get('where') else it_sup_text(codes,names=True)
    except Exception: n_=[]
    if not n_:
        al_=[str(a_).strip() for a_,m_ in (BOM_ALIAS[0] or {}).items() if str(m_).strip() in codes]     # other roll length on your BOM (tape 018 -> 010)
        n_=[_sup_short(v_) for c_ in codes+al_ for v_ in sorted(bom_sup.get(str(c_).upper(),()))]
    if not n_:
        for x_ in (AL_ORDER_ or []):
            b_=x_['b']
            if x_.get('sup_row') or str(b_.get('mat')) not in codes: continue
            n_+=[_sup_short(sp_.get('name')) for sp_ in (b_.get('sup') or []) if str(sp_.get('name') or '').upper()!='UNASSIGNED' and float(sp_.get('stock') or 0)>0]
    if not n_: n_=[f"{v_} (WSA shipper)" for c_ in codes for v_ in dict.fromkeys(WSA_SHIPPER.get(str(c_).upper(),[]))]
    return ' / '.join(dict.fromkeys(n_)) or 'no supplier on file'
EMAIL_CFG='Daily_Email_Settings.txt'
def _email_settings(out_dir):
    """To / Cc of the daily email from the text file Daily_Email_Settings.txt (next to the outputs, the inputs or the
       script; lines 'To: a@x; b@y', 'Cc: ...'; Name / Title lines of older files are ignored - no signature). The addresses are the planner's own data, so they are in
       that file and not in the script. No file yet -> an empty one is written once for the planner to fill."""
    cfg={}
    dirs=list(dict.fromkeys([out_dir]+list(globals().get('_PPC_DIRS') or [])+[os.getcwd(),os.path.dirname(os.path.abspath(__file__))]))
    f_=next((os.path.join(d_,EMAIL_CFG) for d_ in dirs if d_ and os.path.isfile(os.path.join(d_,EMAIL_CFG))),None)
    if f_ is None:
        try:
            with open(os.path.join(out_dir,EMAIL_CFG),'w',encoding='utf-8') as h_:
                h_.write('# Daily email settings - fill once; the script reads this file on every run.\n# Several addresses: separate with ;\nTo: \nCc: \n')
        except OSError: pass
        return cfg
    try:
        raw_=open(f_,'rb').read()
        try: txt_=raw_.decode('utf-8-sig')
        except UnicodeDecodeError: txt_=raw_.decode('cp1252',errors='replace')     # saved by Notepad as ANSI
        for ln_ in txt_.splitlines():
            ln_=ln_.strip()
            if not ln_ or ln_.startswith('#') or ':' not in ln_: continue
            k_,v_=ln_.split(':',1); k_=k_.strip().lower(); v_=v_.strip()
            if k_ in ('to','cc'): v_=', '.join(a_.strip() for a_ in re.split(r'[;,]',v_) if a_.strip())
            if k_ in ('to','cc') and v_: cfg[k_]=v_
    except Exception as ex_: print(f'NOTE: {EMAIL_CFG} not read ({type(ex_).__name__}: {ex_})')
    return cfg
