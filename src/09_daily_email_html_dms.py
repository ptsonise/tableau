PER_HEAD='Stock zero from → to (⚠ = starts inside the transit lead time)'
def write_daily_email(rows, d0, out_dir, attach=None):
    """Daily email to top management, written from the ALERT rows (the same rows as the ALERT of 1_MRP / Inbound /
       4_Status) in the layout of the email the user sends (18-Nov-2026 sample): Dear all, intro, Needs attention, lead-in line, '1. Critical' + one line with '1.1 Raw Materials' / '1.2 Packaging' / '1.3 Consumables' tables (stock-out inside
       the transit lead time, with supplier), '2. Upcoming ETD Alert' + one line with '2.x' tables (FUTURE ETDs arriving after the
       stock-out, one line per material), Notes (no signature). Section numbers follow the sections / groups that have rows. No Summary, no Order this week / Overdue /
       passed-ETD tables; consumables are left out. Two files next to 1_MRP: 6_Email_<date>.eml = an Outlook draft
       (4_Status attached; To / Cc from Daily_Email_Settings.txt) and 6_Email_<date>.html = the same
       body. Facts only - no advice text."""
    import html as _h
    from email.message import EmailMessage
    from email.utils import formatdate
    if not rows: return None
    E=lambda v: _h.escape(str(v))
    D=lambda d_: f"{pd.Timestamp(d_):%d-%b}"
    def cover(v):
        if v is None: return '-'
        if v<0.05: return '0 days'
        return (f"{v:.1f} days" if v<1 else (f"{v:,.0f} day" if round(v)==1 else f"{v:,.0f} days"))
    bom_sup={}
    try:
        for x_ in wb_ro(PPC_SRC[0])['BOM'].iter_rows(min_row=2,values_only=True):
            if x_ and len(x_)>6 and x_[4] and x_[6]: bom_sup.setdefault(str(x_[4]).strip().upper(),set()).add(' '.join(str(x_[6]).split()))
    except Exception as _e_: pass                                 # no PPC file: in-transit / stock suppliers only
    def dn(code,name):                                       # user, 19-Nov-2026: the email shows the Component Description from the PPC BOM instead of the material name (no BOM description = the name)
        return BOM_DESC.get(str(code or '').strip().upper()) or str(name)
    def mat(r): return dn(r.get('code'),r['name'])+(f" ({str(r['where']).replace('Glass G12R Topcon ','')})" if r.get('where') else '')
    def note(r):
        # Situation (user, 19-Nov-2026): blank when nothing is in transit; next arrival after the stock-out = how many days late;
        # otherwise (in transit, but not enough / too late) = how many days are not covered (the first stock-zero period)
        if not (r.get('tr_tot') or 0): return ''
        n_=[]; fi_=pd.Timestamp(r['first']).normalize() if r.get('first') is not None else None
        if fi_ is not None:
            nx_=pd.Timestamp(r['nx'][0]).normalize() if r.get('nx') else None
            if nx_ is not None and nx_>fi_:
                d_=(nx_-fi_).days; n_.append(f"Arrives {d_} day{'' if d_==1 else 's'} late (stock-out {D(fi_)}, next arrival {D(nx_)})")
            else:
                w_=floor_wins(r.get('S') or [],r.get('stock'))         # first stock-zero period, stock never below 0 (= 'Stock zero from -> to' on the same row)
                if w_:
                    k_,z_,o_=w_[0]; c_=z_-k_+1
                    n_.append(f"Not covered: {c_}{'+' if o_ else ''} day{'' if (c_==1 and not o_) else 's'} ({D(fi_)} → "+(f"beyond {D(fi_+pd.Timedelta(days=c_-1))}" if o_ else D(fi_+pd.Timedelta(days=c_-1)))+')')
                else: n_.append('In transit is not enough or arrives too late')
        if isinstance(r.get('ovd'),(int,float)) and r['ovd']>0: n_.append(f"{r['ovd']:,.0f} overdue (arrival date passed)")
        if r.get('subnote'): n_.append(r['subnote'])
        return '; '.join(n_)
    def egrp(code):                                          # 0 raw material, 1 packaging, 2 consumable (the dashboards' RM / PM / CS groups)
        c_=str(code or '').upper(); it_=str(count_item_of(code) or '').strip().lower()
        if it_.startswith('consum'): return 2
        if it_.startswith('packag') or c_.startswith('RMLAB'): return 1
        if it_ or c_.startswith('RM') or c_ in ('','-'): return 0
        return 2 if c_.startswith('CS') else 1
    # Consumables (tapes, barcodes, back label) ARE in the email again (user, 12-Nov-2026: "email daily need mention consumables
    # items"; the 11-Nov rule left them out): their own tables 'Critical – Consumables' / 'Consumables' (ETD Alert), after packaging
    crit=[r for r in rows if r['cls']==0]
    GRP_=[(0,'Raw Materials'),(1,'Packaging'),(2,'Consumables')]   # own numbered tables (1.1, 1.2 ...), only when they have a row
    sup_={id(r):_email_suppliers(r,bom_sup) for r in crit}
    subs=[f"{dn(r.get('code'),r['name'])}: {r['subcov']}" for r in crit if r.get('subcov')]   # Settings SUBSTITUTES: which code covers when, how much (sub_cover)
    # Layout = the email the user sends (11-Nov-2026: "Email should be like this attached"): Dear all, intro, Needs attention, one
    # lead-in line, Critical table, ETD Alert table, Notes (no signature). NO Summary section, no counts, no emoji.
    def _fd(r): return pd.Timestamp(r['first']).normalize() if r.get('first') is not None else None
    t0_=pd.Timestamp(d0).normalize(); bk_=[[],[],[],[]]
    for r in sorted([r for r in crit if _fd(r) is not None],key=_fd):
        n_=(_fd(r)-t0_).days; bk_[0 if n_<=0 else (1 if n_<=7 else (2 if n_<=30 else 3))].append(r)
    later_=sorted([r for r in rows if r['cls']==1 and _fd(r) is not None],key=_fd)
    def _nm(l_): return ', '.join(l_[:3])+(' and others' if len(l_)>3 else '')
    nxt_=[r for k_ in (1,2,3) for r in bk_[k_]]
    subj=f"Daily Material Status – {pd.Timestamp(d0):%d-%b-%Y} | "+(f"Out of stock today: {_nm([mat(r) for r in bk_[0]])}" if bk_[0]
          else (f"Next stock-out: {mat(nxt_[0])} {D(_fd(nxt_[0]))}" if nxt_ else (f"Next stock-out: {mat(later_[0])} {D(_fd(later_[0]))}" if later_ else 'no stock-out in the plan')))
    nopo=[mat(r) for r in crit if not (r.get('tr_tot') or 0)]
    late_=[mat(r) for r in crit if (r.get('tr_tot') or 0)]
    att=[]
    if nopo: att.append(f"Nothing in transit: {', '.join(nopo)}")
    if late_: att.append(f"In transit, but the stock runs out first or the quantity is not enough: {', '.join(late_)}")
    cfg_=_email_settings(out_dir)
    FF='font-family:Aptos,Calibri,Helvetica,sans-serif;font-size:9pt'
    P=f'style="margin:0;{FF};color:#000"'
    TH=f'style="background:#1F3864;color:#fff;padding:5px 8px;text-align:left;{FF};border:1px solid #1F3864"'
    TD=f'style="padding:5px 8px;border:1px solid #D0D5DD;{FF};color:#1F2D3D;vertical-align:top"'
    TDR=f'style="padding:5px 8px;border:1px solid #D0D5DD;{FF};vertical-align:top;color:#C0392B;font-weight:bold;white-space:nowrap"'
    TDN=TD[:-1]+';white-space:nowrap"'                                 # Arrival at plant: one arrival line = one line
    def table(hd,body,nw=-1):
        return ('<table style="border-collapse:collapse;margin:6px 0 14px 0">'+'<tr>'+''.join(f'<th {TH}>{E(h_)}</th>' for h_ in hd)+'</tr>'
                +''.join('<tr>'+''.join(f'<td {TDR if j_==hl else (TDN if j_==nw else TD)}>{E(v_).replace(chr(10),"<br>")}</td>' for j_,v_ in enumerate(vals))+'</tr>' for vals,hl in body)+'</table>')   # hl = the column in red
    def ul(items): return f'<ul style="margin-top:0;margin-bottom:0">'+''.join(f'<li {P}>{E(a_)}</li>' for a_ in items)+'</ul>'
    BR=f'<p {P}><br></p>'
    intro=f"Here is today's raw material, packaging and consumables status, based on SAP MB52 + the physical count, the module plan, and PPC In_Transit arrival dates (Final AtA) (base date {pd.Timestamp(d0):%d-%b-%Y})."
    H=['<div>',f'<p {P}>Dear all,</p>',f'<p {P}>{E(intro)}</p>',BR]
    T=[subj,'','Dear all,','',intro]
    ACT_='Please either reschedule the PO or arrange a special shipment to avoid a production shortage.'
    if att:
        H.append(f'<p {P}><b>Needs attention</b></p>'+ul(att)+BR)      # no reschedule sentence here: sections 1 / 2 carry it (user: duplicate)
        T+=['','Needs attention']+[f"- {a_}" for a_ in att]
    # ETD Alert rows first (the lead-in line names only the sections that exist). FUTURE ETDs only, and only deliveries that ARRIVE
    # AFTER THE STOCK-OUT (the user's email has no 'ships within 7 days' lines, no Alert / Runs out columns); one line per material,
    # the same rows / rule as the 'ETD alert' sheet, which keeps every alert
    fut_={}
    for e_ in ETD_LATE:
        if e_['passed']: continue
        code_,name_,cls_,first_,uom_=_etd_alert(e_); fl_=_etd_flag(e_,first_)
        if fl_!=0: continue
        k_=(code_,e_['sheet'] if code_=='RMGLS00000333' else '')
        t_=fut_.setdefault(k_,dict(pk=egrp(code_),name=dn(code_,name_)+(f" ({e_['sheet'].replace('Glass G12R Topcon ','')})" if k_[1] else ''),sup=set(),n=0,qty=0.0,
                                   uom=uom_ or '',etd0=e_['etd'],pl0=e_['plant'],first=first_))
        t_['n']+=1; t_['qty']+=float(e_['qty'] or 0); t_['etd0']=min(t_['etd0'],e_['etd']); t_['pl0']=min(t_['pl0'],e_['plant'])
        ar_=_etd_arow(e_)
        if ar_ is not None: t_['sup'].add(_email_suppliers(ar_,bom_sup))
        elif e_.get('shipper') and not str(e_['shipper']).replace('.0','').isdigit(): t_['sup'].add(str(e_['shipper']).strip())
    futl=sorted(fut_.values(),key=lambda t_:(t_['first'] if t_['first'] is not None else t_['etd0'],t_['etd0'],t_['name']))
    lead_=('Please find the critical materials and the upcoming ETD alert (the PO remains open) below:' if (crit and futl)
           else ('Please find the critical materials below:' if crit else ('Please find the upcoming ETD alert (the PO remains open) below:' if futl else '')))
    if lead_:
        H.append(f'<p {P}>{E(lead_)}</p>'); T+=['',lead_]
    # numbered sections (user's email 18-Nov-2026): '1. Critical' + one line, '1.1 Raw Materials' / '1.2 Packaging' / '1.3 Consumables'
    # (only groups with rows, numbered in order), '2. Upcoming ETD Alert' + one line, '2.1 ...'; numbers follow the sections that exist
    sec_=[0]
    def head_(txt_,line_):
        sec_[0]+=1; H.append(f'<p {P}><b>{sec_[0]}. {E(txt_)}</b></p><p {P}>{E(line_)}</p>'); T.extend(['',f"{sec_[0]}. {txt_}",line_])
    def sub_(k_,gl_):
        H.append(f'<p {P}><br></p><p {P}><b>{sec_[0]}.{k_} {E(gl_)}</b></p>'); T.extend(['',f"{sec_[0]}.{k_} {gl_}"])
    if crit:
        head_('Critical','Materials will run out of stock within the transit lead time. '+ACT_)
        # the two columns = the ALERT's own text (user, 18-Nov-2026): 'Stock zero from → to (⚠ ...)' = r['per'], 'Arrival at plant
        # (Final AtA)' = `ata_show` (ONE rule for every output, user 26-Nov-2026: same as 4_Status and 1_MRP)
        hd=['Material','Supplier','Cover today',PER_HEAD,'Arrival at plant (Final AtA)','Situation']; k_=0
        for pk_,gl_ in GRP_:
            body=[([mat(r),sup_[id(r)],cover(r.get('days')),str(r.get('per') or '-'),ata_show(r.get('ata'),[r['code']],r.get('need')) or '',note(r)],3) for r in crit if egrp(r['code'])==pk_]
            if not body: continue
            k_+=1; sub_(k_,gl_); H.append(table(hd,body,nw=4))
            T+=[f"- {b_[0][0]} | {b_[0][1]} | cover {b_[0][2]} | stock zero {b_[0][3].replace(chr(10),'; ')} | arrival at plant {b_[0][4].replace(chr(10),'; ')} | {b_[0][5]}" for b_ in body]
    # no Summary / Order this week / Overdue / passed-ETD tables (user, Nov 2026 - the ETD alert sheet keeps that detail)
    if futl:
        head_('Upcoming ETD Alert','The ETD is late: the delivery arrives after the stock-out. '+ACT_); k_=0
        for pk_,gl_ in GRP_:
            body=[([t_['name'],' / '.join(sorted(t_['sup'])) or '-',str(t_['n']),f"{t_['qty']:,.0f} {t_['uom']}".strip(),D(t_['etd0']),D(t_['pl0'])],-1) for t_ in futl if t_['pk']==pk_]
            if not body: continue
            k_+=1; sub_(k_,gl_); H.append(table(['Material','Supplier','Deliveries','Quantity','Next ETD','WSA plant date'],body))
            T+=[f"- {b_[0][0]} ({b_[0][1]}): {b_[0][2]} deliver{'y' if b_[0][2]=='1' else 'ies'}, {b_[0][3]}, next ETD {b_[0][4]}, WSA plant date {b_[0][5]}" for b_ in body]
    notes=subs+([f"Full details: the attached {os.path.basename(attach)} (BOM Material Dashboard – Weekly = every material in risk order)."] if attach else [])
    if notes:
        H.append(f'<p {P}><br></p><p {P}><b>Notes</b></p>'+ul(notes))
        T+=['','Notes']+[f"- {n_}" for n_ in notes]
    H.append('</div>')                                       # no signature (user, 18-Nov-2026): Outlook adds the planner's own
    html_='\n'.join(H)
    stem=os.path.join(out_dir,outfile_name("mail"))
    with open(stem+'.html','w',encoding='utf-8') as f_:
        f_.write(f'<!doctype html><html><head><meta charset="utf-8"><title>{E(subj)}</title></head><body>{html_}</body></html>')
    msg=EmailMessage(); msg['Subject']=subj; msg['To']=cfg_.get('to') or ''; msg['Date']=formatdate(localtime=True); msg['X-Unsent']='1'   # X-Unsent: Outlook opens it as a draft
    if cfg_.get('cc'): msg['Cc']=cfg_['cc']
    msg.set_content('\n'.join(T[2:])); msg.add_alternative(f'<html><body>{html_}</body></html>',subtype='html')
    if attach and os.path.exists(attach):
        with open(attach,'rb') as f_: msg.add_attachment(f_.read(),maintype='application',subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet',filename=os.path.basename(attach))
    with open(stem+'.eml','wb') as f_: f_.write(bytes(msg))
    for n_ in ([] if globals().get('KEEP_OLD') else sorted(os.listdir(out_dir))):     # older daily emails go, like 1_MRP
        if re.fullmatch(out_rx('mail')+r'_\d{4}-\d{2}-\d{2}\.(eml|html)',n_) and not n_.startswith(os.path.basename(stem)):
            try: os.remove(os.path.join(out_dir,n_))
            except OSError: pass
    print(f"DAILY EMAIL: {stem}.eml (double-click = Outlook draft, 4_Status attached; "
          +(f"To / Cc from {EMAIL_CFG}" if cfg_.get('to') else f"add the To line - or fill {EMAIL_CFG} once")
          +f"; check and send) + .html - subject '{subj}'; {len(crit)} critical; ETD alert (future ETDs arriving after the stock-out) "
          f"{sum(t_['n'] for t_ in futl)} deliveries on {len(futl)} line(s); consumables included",flush=True)
    return stem
def in_lead_time(r):
    """Web pages list a material when its Days of cover (stock only) <= ITS OWN transit lead time (user, 23-Nov-2026; was <= 21 days), OR its first
       stock-zero day is within lead time + RELEASE_WINDOW days (the ALERT's 'ORDER NOW' window; user 23-Nov-2026 "want to see" pallet 535: 18.7 days vs
       lead time 14, runs out 09-Dec = lead time + 5). No demand = not listed."""
    if r.get('days') is None: return False
    lt=int(r.get('lt') or r.get('_lt') or 0); k0=r.get('k0')
    return float(r['days'])<=lt or (k0 is not None and k0-lt<=RELEASE_WINDOW)
def _cust_vendor(nm, code, rl):
    """First allowed vendor (Settings ALLOWED VENDORS) of a customer name for a material: Settings name / MIS name, else the customer whose name
       words contain its words ('Lucky 7 Solar Farm LLC' = 'Sabanci-Lucky 7 Solar'); None = no rule (any vendor)."""
    v_=rl.get((_norm(nm),code))
    if v_: return v_[0]
    return (_cust_vendors(nm,code,rl) or [None])[0]
def _cust_vendors(nm, code, rl):
    """All allowed vendors (Settings ALLOWED VENDORS rows, top row first) of a customer name for a material; [] = no rule (any vendor)."""
    v_=rl.get((_norm(nm),code))
    if v_: return list(v_)
    tk_=_toks(nm); hit_={tuple(x_) for k_,x_ in rl.items() if k_[1]==code and tk_ and tk_<=_toks(k_[0])}
    return list(next(iter(hit_))) if len(hit_)==1 else []
def usage_vendors(code, where=None):
    """Vendors the customers of a web card may use (the customers of its BOM usage lines, `bom_use_of`): set of Settings vendors, or None when a
       customer has no rule (any vendor allowed) or the card has no customer line. User 22-Nov-2026: 'why do I still see FLAT arrive in glass for SB'."""
    ln_=bom_use_of(code,where)[1]; rl_=_vendor_rules({}); c_=str(code or '').strip(); out=set()
    import html as _h
    for l_ in ln_:
        v_=_cust_vendors(_h.unescape(l_.split(':')[0]),c_,rl_)
        if not v_: return None
        out.update(v_)
    return out or None
def vend_ok(vs, name):
    """Shipper / supplier text allowed for the vendor set vs? Unknown names and no rule = allowed; 'A / B' = allowed when any part is."""
    if not vs: return True
    parts=[x_.strip() for x_ in re.split(r'\s+/\s+|;',str(name or '')) if x_.strip()]
    if not parts or all(p_.upper() in VP_UNKNOWN for p_ in parts): return True
    return any(p_.upper() in VP_UNKNOWN or any(_vend_hit(v_,p_) for v_ in vs) for p_ in parts)
def bom_use_of(code, where=None):
    """BOM usage per module of a material for the web pages (user 21/22-Nov-2026 "show bom usage" - all 3 pages): short text ('1.2 m' or
       'Customer 1.2 · Customer 1.5 m') + the full lines with source (= 'Usage per module by customer' of Stock Check; substitute '+' lines left out)."""
    import html as _h
    e=_h.escape; use={}
    for x_ in (AL_ORDER_ or []):
        if x_.get('sup_row') or not x_.get('usage_src'): continue
        ln_=[l_.strip() for l_ in str(x_['usage_src']).split('\n') if l_.strip() and not l_.strip().startswith('+')]
        if ln_: use.setdefault(str(x_['b']['mat']).strip().upper(),ln_)
    ln_=use.get(str(code or '').strip().upper()) or []
    if where and ln_:                                        # a row of ONE WSA sheet (back glass 333 on 2 sheets): only the customers whose first allowed vendor (Settings ALLOWED VENDORS) is a supplier of THAT sheet (user 22-Nov-2026 "why dont follow rule customers suppliers?")
        g_=next((g for g in (WSA_REFS[0] or []) if g.get('sheet')==where and str(code or '').strip() in (g.get('codes') or [])),None)
        sn_=[n_ for n_ in ((g_ or {}).get('ppcsup_names') or [])]
        rl_=_vendor_rules({}); c_=str(code or '').strip()
        vend_=lambda l_: _cust_vendors(l_.split(':')[0],c_,rl_)
        keep_=[l_ for l_ in ln_ if not vend_(l_) or any(_vend_hit(v2_,n_) for v2_ in vend_(l_) for n_ in sn_)] if sn_ else ln_
        ln_=keep_ or ln_
    pr_=[re.match(r'(.+?):\s*([\d.eE+-]+)\s*(.*?)\s*\((.*)\)$',l_) for l_ in ln_]
    if not ln_ or not all(pr_): return (e(' · '.join(ln_)) if ln_ else ''),[e(l_) for l_ in ln_]
    q_=list(dict.fromkeys(m_.group(2) for m_ in pr_)); u_=pr_[0].group(3)
    sh_=(f"{q_[0]} {u_}".strip() if len(q_)==1 else ' · '.join(f"{' '.join(m_.group(1).split()[:2])} {m_.group(2)}" for m_ in pr_)+(f" {u_}" if u_ else ''))
    return e(sh_),[e(l_) for l_ in ln_]

WEEK_VIEW_WEEKS=16    # 'Weeks' tab = week-by-week calendar, 16 weeks from this week (user 23-Nov-2026)
CAL_DAYS=28    # Calendar tab = the next 4 weeks from TODAY (user 23-Nov-2026)
WSA_DLV={}     # open WSA deliveries per ALERT row (key = 'where' or code), filled in step 2 from the 2_WSA sheets (`wsa_open_deliveries`)
DET_DAYS=112   # chart length of the detail panel (days from the base date) = ALERT_HORIZON
def det_data(r, d0):
    """Detail panel of Material_Shortage / Shortage_Board = the 2_WSA detail (user 22-Nov-2026 "have detail like wsa has"): projected stock chart,
       PPC arrivals table, open WSA deliveries table. From the ALERT row itself (`S` = stock + PPC in transit - plan, `Tp` arrivals); presentation only."""
    S=list(r.get('S') or [])[:DET_DAYS]; T=list(r.get('Tp') or [])[:DET_DAYS]; P=list(r.get('P') or [])[:DET_DAYS]
    n=len(S); cum=0.0; d=[]; F_,Sf_=floor_path(S,r.get('stock'))
    for k in range(n):
        t_=float(T[k]) if k<len(T) and T[k] else 0.0; cum+=t_
        d.append([f"{(pd.Timestamp(d0)+pd.Timedelta(days=k)):%Y-%m-%d}",round(max(0.0,S[k]-cum),1),(round(float(P[k]),1) if k<len(P) else 0),round(t_,1),round(F_[k],1)])
    z=next((k for k in range(n) if Sf_[k]<0),None)
    arr=[[f"{pd.Timestamp(x[0]):%d-%b-%y}",x[3]] for x in d if x[3]]
    key=str(r.get('where') or r.get('code') or '').strip(); dl=WSA_DLV.get(key) or WSA_DLV.get(str(r.get('code') or '').strip().upper()) or ([],0)
    vs_=usage_vendors(r.get('code'),r.get('where')); dlv_=[x_ for x_ in dl[0] if vend_ok(vs_,x_[3])]   # shippers the card's customers may not use are out (Settings ALLOWED VENDORS)
    return dict(d=d,z=z,arr=arr[:12],na=len(arr),dlv=dlv_,nd=(dl[1]-(len(dl[0])-len(dlv_))))
DET_CSS=""".dch{background:var(--card);border:1px solid var(--line);border-radius:6px;padding:5px;margin-top:6px;grid-column:1/-1}.dch svg{width:100%;height:150px;display:block}
.dch .lg{color:var(--mu);font-size:12px;display:flex;gap:10px;flex-wrap:wrap;padding:0 2px 3px}.dch .lg i{display:inline-block;width:14px;height:3px;vertical-align:3px;margin-right:5px;border-radius:2px}
.dtb{background:var(--card);border:1px solid var(--line);border-radius:6px;padding:5px 7px;margin-top:6px;overflow-x:auto;grid-column:1/-1}.dtb h4{margin:0 0 3px;font-size:13px;color:var(--mu);font-weight:600}
.dtb table{width:100%;border-collapse:collapse;font-size:13px}.dtb th,.dtb td{padding:2px 6px;text-align:left;border-bottom:1px solid var(--line);white-space:nowrap;background:none;position:static}
.dtb th{color:var(--mu);font-weight:500;font-size:12px}.dtb tr:last-child td{border-bottom:0}.dtb td.r,.dtb th.r{text-align:right;font-variant-numeric:tabular-nums;padding-right:14px}
"""
DET_JS=r"""
const _df=s=>{const t=new Date(s+"T00:00:00");return t.getDate()+"-"+["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"][t.getMonth()]+"-"+String(t.getFullYear()).slice(2)};
const _fn=v=>v==null?"-":Math.round(v).toLocaleString();
function detChart(r){const d=r.d,W=900,H=200,P=52,n=d.length;if(!n)return"";let lo=0,hi=0;d.forEach(x=>{[x[1],x[4]].forEach(v=>{if(v!=null){lo=Math.min(lo,v);hi=Math.max(hi,v)}})});if(hi===lo)hi=lo+1;
 const X=i=>P+(W-P-6)*i/Math.max(1,n-1),Y=v=>10+(H-30)*(hi-v)/(hi-lo);
 const ln=k=>d.map((x,i)=>x[k]==null?"":`${i&&d[i-1][k]!=null?"L":"M"}${X(i).toFixed(1)},${Y(x[k]).toFixed(1)}`).join("");
 let g=`<line x1="${P}" x2="${W-6}" y1="${Y(0)}" y2="${Y(0)}" stroke="#6b7280" stroke-dasharray="3 3"/>`;
 d.forEach((x,i)=>{if(x[3])g+=`<line x1="${X(i)}" x2="${X(i)}" y1="${H-20}" y2="${H-28}" stroke="#7dc242" stroke-width="3"/>`});
 for(let i=0;i<n;i+=14)g+=`<text x="${X(i)}" y="${H-4}" fill="#8b93a3" font-size="10" text-anchor="middle">${_df(d[i][0]).slice(0,-3)}</text>`;
 g+=`<text x="${P-6}" y="${Y(hi)+4}" fill="#8b93a3" font-size="10" text-anchor="end">${_fn(hi)}</text><text x="${P-6}" y="${Y(0)+4}" fill="#8b93a3" font-size="10" text-anchor="end">0</text>${lo<0?`<text x="${P-6}" y="${Y(lo)}" fill="#8b93a3" font-size="10" text-anchor="end">${_fn(lo)}</text>`:""}`;
 if(r.z!=null)g+=`<line x1="${X(r.z)}" x2="${X(r.z)}" y1="10" y2="${H-20}" stroke="#ef4444" stroke-width="1.5"/>`;
 g+=`<path d="${ln(1)}" fill="none" stroke="#8b93a3" stroke-width="1.5"/><path d="${ln(4)}" fill="none" stroke="#2b7a3e" stroke-width="2"/>`;
 return `<div class="dch"><div class="lg"><span><i style="background:#2b7a3e"></i>Stock + PPC arrivals</span><span><i style="background:#8b93a3"></i>Stock only (no arrivals)</span><span><i style="background:#7dc242"></i>PPC arrival</span><span><i style="background:#ef4444"></i>Stock-zero day</span></div><svg viewBox="0 0 ${W} ${H}" preserveAspectRatio="none">${g}</svg></div>`}
function detTbl(t,h,rows,rc){return rows.length?`<div class="dtb"><h4>${t}</h4><table><tr>${h.map((x,i)=>`<th${rc.includes(i)?' class="r"':""}>${x}</th>`).join("")}</tr>${rows.map(r=>`<tr>${r.map((x,i)=>`<td${rc.includes(i)?' class="r"':""}>${rc.includes(i)?_fn(x):(x||"")}</td>`).join("")}</tr>`).join("")}</table></div>`:""}
const detBlock=r=>(r.ch?detChart(r.ch):"")+(r.ch?detTbl(`PPC In_Transit arrivals (${r.ch.na})`,["Plant date","Qty"],r.ch.arr,[1])+detTbl(`Open WSA deliveries - Scheduled Delivery, not cancelled (${r.ch.nd})`,["Plant date","Qty","PO","Shipper","ETD","BL#","Vessel ETA","Status"],r.ch.dlv,[1]):"");
"""
def det_inject(page):
    return page.replace('</style>',DET_CSS+'</style>',1).replace('<script>','<script>'+DET_JS,1)

BOARD_MAX_DAYS=63       # Shortage_Board: day strip up to the longest transit lead time of the listed rows, at most 63 days (DEFAULT)
def lead_arrivals(r, d0):
    """Arrivals inside the material's transit lead time = [[date text, qty], ...] from the daily in-transit path `Tp` (ONE function: the 5_Web
       Summary 'next:' line and the 4_Status BOM Weekly detail row both read it, user 27-Nov-2026). Nothing is calculated here."""
    def num(v):
        try: return None if v is None else round(float(v),1)
        except Exception: return None
    return [[f"{(pd.Timestamp(d0)+pd.Timedelta(days=k_)):%d-%b}",num(t_)] for k_,t_ in enumerate(list(r.get('Tp') or [])[:int(r.get('lt') or r.get('_lt') or 0)]) if t_ and t_>0]
def next_arrival_text(r, d0):
    """4_Status BOM Weekly detail line (user 27-Nov-2026: '(+5 more) is waste if you do not list all'): EVERY arrival inside the lead time
       from `lead_arrivals`: 'next (SET): 29-Nov 30,000 \u00b7 03-Dec 12,000'; none -> 'next: none within the lead time'."""
    a=lead_arrivals(r,d0)
    if not a: return 'next: none within the lead time'
    u=str(r.get('uom') or '').strip()
    return f"next{f' ({u})' if u else ''}: "+' \u00b7 '.join(f"{d} {q:,.0f}" for d,q in a)
def write_material_daily_update(rows, d0, out_dir):
    """5_Web_<date>.html (user, 22-Nov-2026: "new dashboard inspired by OneFootball match page", then 4 screenshots of the app): FOURTH web
       page, one material = one 'match'. Match list (cards) -> tap = match screen like the app: dark scoreboard header (Days of cover (stock only) vs Inv. Days
       + 21 d arrivals, status = ALERT class, 'goals' = PPC arrivals, 'red cards' = stock-out periods), pill tabs Overview / Ticker / Stats /
       BOM usage. Overview = Recent form (W/D/L per week: cover; arrivals), Details card, Live table (risk rank with neighbours), Top arrivals,
       'Will it last?' bar (stock / arrivals / gap vs need until the lead-time end). Stats = Match Momentum (daily cover up, stock-out down, arrival
       marks, lead-time line) + stat rows. Same rows / order as Material_Shortage. Presentation only - no calculation of its own."""
    import json as _j, html as _h
    if not rows: return None
    exp=_heat_expand(rows,d0); n_all=len(exp); e=_h.escape
    srt=sorted(exp,key=lambda r:(_heat_grp(r),)+bom_risk_key(r))      # user 25-Nov-2026: ALL materials like the BOM Material Dashboard (was only in_lead_time), in its row order
    if DMS_ORDER:
        _ix={k_:i_ for i_,k_ in enumerate(DMS_ORDER)}; _nm=lambda r:str(r.get('name') or r.get('code') or '')+(f" ({str(r['where']).replace('Glass G12R Topcon ','')})" if r.get('where') else '')
        srt=[r for _,r in sorted(enumerate(srt),key=lambda t:(_ix.get((str(t[1].get('code') or '').strip().upper(),_nm(t[1])),10**6),t[0]))]
    def num(v):
        try: return None if v is None else round(float(v),1)
        except Exception: return None
    NW=max(1,min(-(-BOARD_MAX_DAYS//7),-(-(max([int(r.get('lt') or r.get('_lt') or 0) for r in srt] or [0])+1)//7)))    # weeks of the strip = up to the longest lead time of the listed rows (same rule as Shortage_Board)
    data=[]
    for r in srt:
        lt=int(r.get('lt') or r.get('_lt') or 0); N=max(35,min(DET_DAYS,lt+14),7*NW)
        cvd=bom_day_cover(r)[:N]; Tp=list(r.get('Tp') or [])[:N]
        so=[(not isinstance(v,str)) and v is not None and v<1 for v in cvd]
        runs=[]; k=0
        while k<len(so):
            if so[k]:
                j=k
                while j+1<len(so) and so[j+1]: j+=1
                runs.append([k,j]); k=j+1
            else: k+=1
        zd=[k for k,v in enumerate(bom_day_cover(r)[:DET_DAYS+WEEK_VIEW_WEEKS*7]) if (not isinstance(v,str)) and v is not None and v<1]     # stock-zero days over the whole horizon, same rule as `runs` (calendar tab, user 23-Nov-2026)
        arr=[[k,num(t)] for k,t in enumerate(Tp) if t and t>0]
        aw=[[k,num(t)] for k,t in enumerate(list(r.get('Tp') or [])[:DET_DAYS+WEEK_VIEW_WEEKS*7]) if t and t>0]     # arrivals for the Weeks tab (16 weeks from this week)
        ac=[]                                  # cover after each arrival, stock NEVER below 0 (user 23-Nov-2026: an arrival cannot cover days already lost); no demand that day -> next production day, shift = days later
        for k_,_q in arr:
            j_=k_
            while j_<len(cvd) and (cvd[j_] is None or isinstance(cvd[j_],str)): j_+=1
            ac.append([round(float(cvd[j_]),1) if j_<len(cvd) else None,j_-k_])      # cvd = bom_day_cover (stock never below 0) = the week strip
        stock=float(r.get('stock') or 0.0); need=float(r.get('need') or 0.0)
        wk=[next((num(v) for v in cvd[7*w:7*w+7] if v is not None and not isinstance(v,str)),None) for w in range(NW)]
        wa=[round(sum(t for t in Tp[7*w:7*w+7] if t and t>0)) for w in range(NW)]
        W16=WEEK_VIEW_WEEKS; cvf_=bom_day_cover(r); Tf_=list(r.get('Tp') or [])     # Overview + Arrival strips = next 16 weeks (user 25-Nov-2026); list cards keep NW
        wkx=lambda arr: [next((num(v) for v in arr[7*w:7*w+7] if v is not None and not isinstance(v,str)),None) for w in range(W16)] if arr else None
        wk16_=wkx(cvf_); wa16_=[round(sum(t for t in Tf_[7*w:7*w+7] if t and t>0)) for w in range(W16)]
        bu_,bl_=bom_use_of(r.get('code'),r.get('where'))
        HZ=DET_DAYS+WEEK_VIEW_WEEKS*7                                     # Settings SUBSTITUTES on every tab (user 25-Nov-2026): substitute arrivals, days / qty it covers, cover with it
        sx_=[dict(n=e(f"{x_['n']} {x_['c']}"),u=e(str(x_['u'] or '')),ar=[[k_,num(t_)] for k_,t_ in enumerate(x_['T'][:HZ]) if t_>0.5],
                  us=[[k_,round(float(v_),3)] for k_,v_ in enumerate(x_['use'][:HZ]) if v_>1e-6]) for x_ in r.get('sub_x') or []]     # use: 3 decimals so period totals = the subcov text (checker: 94,342 vs 94,341)
        zst_=set(r.get('sub_zero') or []); zc_=[k_ for k_ in zd if k_ not in zst_] if sx_ else []
        cvs_=r.get('cv_sub'); wks_=[next((num(v) for v in cvs_[7*w:7*w+7] if v is not None and not isinstance(v,str)),None) for w in range(NW)] if cvs_ else None
        cds_=[c_.strip() for c_ in re.split(r'\s*(?:/|,|\+|\n)\s*',str(r.get('code') or '')) if c_.strip()]
        pb_,pu_=_prpo_by(cds_) if (cds_ and PRPO_LINES) else ({},0.0)     # open PO / PR lines (user 25-Nov-2026 "I also need PO ETA and PR ETA info"): information, in no cover
        pp_=[[(pd.Timestamp(k_[0])-pd.Timestamp(d0).normalize()).days,k_[1],num(q_)] for k_,q_ in sorted(pb_.items())]
        ppc_=[]; wkp_=None; wkp16_=None                  # user 25-Nov-2026 "how many coverage days with each PR ETA and PO ETA": scenario = PPC arrivals + every DATED open PO / PR line on its ETA
        if pp_ and r.get('S'):              # (date passed = on the base date); daily cover = bom_day_cover (stock never below 0, no later arrivals), so the cover on an
            S0_=list(r['S']); n2_=len(S0_); ex_=[0.0]*n2_   # ETA day counts every PO / PR up to that day. Information only - in no Excel output.
            for k_,_t,q_ in pp_:
                if k_<n2_: ex_[max(0,k_)]+=float(q_ or 0)
            cu_=0.0; S2_=[]
            for k_ in range(n2_): cu_+=ex_[k_]; S2_.append(S0_[k_]+cu_)
            T0_=(list(r.get('Tp') or [])+[0.0]*n2_)[:n2_]
            cv2_=bom_day_cover(dict(r,S=S2_,Tp=[(T0_[k_] or 0.0)+ex_[k_] for k_ in range(n2_)]))
            for k_,_t,q_ in pp_:
                j_=max(0,k_)
                while j_<n2_ and (cv2_[j_] is None or isinstance(cv2_[j_],str)): j_+=1
                ppc_.append([round(float(cv2_[j_]),1),j_-max(0,k_)] if (k_<n2_ and j_<n2_) else [None,0])
            wkp_=[next((num(v) for v in cv2_[7*w:7*w+7] if v is not None and not isinstance(v,str)),None) for w in range(NW)]; wkp16_=wkx(cv2_)
        gv_=r.get('sub_give'); gv_=dict(n=e(f"{gv_['n']} {gv_['c']}"),us=[[k_,round(float(v_),3)] for k_,v_ in enumerate(gv_['use'][:HZ]) if v_>1e-6]) if gv_ else None
        c_=r.get('cls'); c_=4 if c_ is None else int(c_)
        data.append(dict(n=e((r.get('name') or r.get('code') or '')+(f" ({str(r['where']).replace('Glass G12R Topcon ','')})" if r.get('where') else '')),c=e(str(r.get('code') or '')),d=e(BOM_DESC.get(str(r.get('code') or '').strip().upper()) or ''),ch=det_data(r,d0),
            a21=lead_arrivals(r,d0),
            g=_heat_grp(r),k=c_,lt=lt,N=N,st=num(stock),u=e(str(r.get('uom') or '')),nd=num(need),
            dy=num(r.get('days')),di=num(r.get('days_it')),dn=num(r.get('days_nx')),
            dw=num(r.get('days_nx')) if need>0 else None,   # next arrival ONLY (user 23-Nov-2026)
            per=e(str(r.get('per') or '').replace('\n',' · ')),wh=e(str(r.get('where') or '')),
            ata=[e(l_) for l_ in re.split(r'\n|  ;  ',str(ata_show(r.get('ata'),[r.get('code')],r.get('need')) or '')) if l_.strip()],bu=bu_,bl=bl_,sb=e(str(r.get('subcov') or '')),
            runs=runs,zd=zd,ar=arr,aw=aw,ac=ac,arl=sum(t_ for k_,t_ in arr if k_<=lt),
            cv=[(None if (isinstance(v,str) or v is None) else round(min(float(v),60.0),1)) for v in cvd],wk=wk,wa=wa,
            nx=([f"{pd.Timestamp(r['nx'][0]):%d-%b}",num(r['nx'][1])] if r.get('nx') else None),
            pp=pp_,ppc=ppc_,wkp=wkp_,wk16=wk16_,wa16=wa16_,wks16=wkx(r.get('cv_sub')),wkp16=wkp16_,pu=num(pu_) if pu_>1e-9 else None,sx=sx_,zc=zc_,wks=wks_,gv=gv_,sn=e(str(r.get('subnote') or '')),il=bool(in_lead_time(r)),
            bd=[[e(h_),e(v_)] for h_,v_ in DMS_INFO.get((str(r.get('code') or '').strip().upper(),str(r.get('name') or r.get('code') or '')+(f" ({str(r['where']).replace('Glass G12R Topcon ','')})" if r.get('where') else '')),[])]))   # every BOM Weekly column (user 25-Nov-2026)
    nbd=sum(1 for x_ in data if x_['bd'])
    if data: print(f"  5_Web: BOM Material Dashboard columns shown for {nbd} of {len(data)} listed materials"+('' if nbd==len(data) else ' (NOTE: rows without a BOM Weekly row: '+', '.join(x_['c'] for x_ in data if not x_['bd'])+')'),flush=True)
    meta=dict(H=DET_DAYS,CH=CAL_DAYS,WW=WEEK_VIEW_WEEKS,to=max(0,(pd.Timestamp(dt.date.today())-pd.Timestamp(d0).normalize()).days),d0=f"{d0:%Y-%m-%d}",d0t=f"{d0:%d-%b-%Y}",tot=n_all,run=f"{dt.datetime.now():%d-%b-%Y %H:%M}",
              cls=[c[0].title().replace('In Transit Lead Time','In Lead Time').replace('Ok','OK') for c in ALERT_CLS])
    js_data=_j.dumps(dict(meta=meta,rows=data)).replace('</','<\\/')
    page=r"""<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Material Daily Update</title><style>
@import url("https://fonts.googleapis.com/css2?family=Poppins:wght@400;500;600;700;800&display=swap");
:root{--bg:#f7f7f7;--card:#fff;--c2:#eef0ee;--card2:#eef0ee;--tx:#15191a;--mu:#6b7370;--line:#e4e7e4;--red:#e5384d;--acc:#1e8a3c;--acc2:#2b7a3e;--or:#f5a623;color-scheme:light}
*{box-sizing:border-box;-webkit-tap-highlight-color:transparent}body{margin:0;background:var(--bg);color:var(--tx);font:16px/1.35 Poppins,"Segoe UI",system-ui,sans-serif}
.pg{max-width:620px;margin:0 auto;padding:0 14px 110px}
.tt{background:#fff;margin:0 -14px;padding:18px 16px 12px;position:sticky;top:0;z-index:5;box-shadow:0 1px 0 var(--line)}.tt h1{display:inline-block;margin:0 0 6px;padding:5px 18px;border-radius:12px;background:linear-gradient(90deg,#3b9fd8,#7dc242);color:#fff;font-size:21px;font-weight:700;letter-spacing:0}.tt .sub{color:#2b6b3a;font-size:13.5px}
.brand{display:flex;flex-direction:column;line-height:1;font-style:italic}.brand b{font-size:26px;font-weight:800;color:#2b7a3e;letter-spacing:-.6px}.brand span{font-size:10px;font-weight:700;color:#f5a623;letter-spacing:.02em;margin-top:1px}
#q{width:100%;margin:12px 0;font:inherit;font-size:16px;padding:10px 14px;border-radius:999px;border:1px solid var(--line);background:var(--card)}
.mc{background:var(--card);border:1px solid var(--line);border-radius:16px;padding:12px 14px;margin-bottom:10px;cursor:pointer}.mc:active{background:var(--c2)}
.row{display:flex;justify-content:space-between;gap:10px;align-items:center}.nm{font-weight:600;font-size:16px;line-height:1.25}
.tg{font-size:11px;color:var(--mu);border:1px solid var(--line);border-radius:6px;padding:0 5px;margin-left:5px;font-weight:500;vertical-align:1px}
.sub2{color:var(--mu);font-size:12.5px;margin-top:2px}.sc{font-weight:800;font-size:28px;line-height:1;text-align:right;white-space:nowrap}.sc small{display:block;font-size:10.5px;font-weight:500;color:var(--mu);margin-top:3px;white-space:normal;max-width:96px;margin-left:auto;line-height:1.25}.sc.r{color:var(--red)}.rd{color:var(--red);font-weight:600}
.tl{position:relative;height:10px;border-radius:5px;background:#e3e8e3;margin:10px 0 4px}.tl s{position:absolute;top:0;bottom:0;background:var(--red);border-radius:3px}
.tl u{position:absolute;top:-5px;border-left:5px solid transparent;border-right:5px solid transparent;border-top:8px solid var(--acc);transform:translateX(-5px)}
.tl b{position:absolute;top:-3px;bottom:-3px;width:2px;background:var(--tx)}
.fade{position:fixed;left:0;right:0;bottom:0;height:96px;background:linear-gradient(rgba(247,247,247,0),#f7f7f7 62%);pointer-events:none;z-index:9}
.nav{position:fixed;left:50%;transform:translateX(-50%);bottom:14px;display:flex;gap:4px;background:#fff;border:1px solid var(--line);border-radius:999px;padding:6px;box-shadow:0 4px 18px rgba(0,0,0,.12);z-index:10;max-width:calc(100vw - 24px)}
.nav{overflow-x:auto;scrollbar-width:none;padding:6px;gap:0;width:calc(100vw - 24px);max-width:596px;justify-content:space-between}.nav::-webkit-scrollbar{display:none}.nav button{border:0;background:none;font:inherit;flex:1 1 auto;font-size:clamp(12.5px,3.5vw,15px);font-weight:700;padding:11px 6px;border-radius:999px;cursor:pointer;color:var(--tx);white-space:nowrap}.nav button.on{background:linear-gradient(90deg,#3b9fd8,#7dc242);color:#fff}.nav small{color:var(--mu);font-weight:700}.nav button.on small{color:#fff}
.ov{position:fixed;inset:0;background:var(--bg);display:none;z-index:20;overflow-y:auto}.ov.on{display:block}
.shw{max-width:620px;margin:0 auto}
.hero{background:linear-gradient(180deg,#1b5c2d 0,#0b0f0c 62%);color:#fff;padding:14px 14px 14px}
.bt{display:flex;justify-content:space-between}.cb{width:42px;height:42px;border-radius:50%;border:0;background:rgba(255,255,255,.16);color:#fff;font-size:20px;cursor:pointer}
.bd0{display:grid;grid-template-columns:1fr auto 1fr;align-items:start;gap:6px;text-align:center;margin-top:6px}
.tn{font-size:14px;margin-top:6px}.xg{font-size:12px;opacity:.65;margin-top:3px}
.bd0 .n{font-size:40px;font-weight:800;line-height:1.1;margin-top:6px}.bd0 .n.r{color:#ff9aa6}.mid{padding-top:6px}.chip{display:inline-block;font-size:12px;font-weight:600;background:rgba(255,255,255,.18);border-radius:999px;padding:4px 10px}.chip.rc{background:var(--red)}.vs{opacity:.85;font-size:13px;font-weight:600;margin-top:10px;line-height:1.25}
.sg{display:grid;grid-template-columns:max-content 22px max-content;justify-content:center;align-items:center;gap:6px 10px;margin:16px 6px 2px;font-size:13.5px;white-space:nowrap}.sg .l{text-align:right}.sg .m{text-align:center;line-height:1;font-size:12px;color:#7ddc8f}.sg .m.cL{color:#ff6b6b}.sg .m.cD{color:#f6c4c4}.sg .m.cW{color:#c9cfca}.sg .m.cX{color:#8f9792}.sg .m.k{color:#ff6b6b;font-size:14px}.sg .rr{text-align:left;color:#ff9aa6}
.sg b{opacity:.65;font-weight:500}
.pills{display:grid;grid-template-columns:repeat(3,1fr);gap:8px;max-width:420px;margin:0 auto;padding:14px 14px 4px}
.pill{min-width:0;border:2px solid var(--line);background:var(--card);border-radius:999px;padding:9px 4px;font:inherit;font-size:15px;cursor:pointer;color:var(--tx);text-align:center;white-space:nowrap}.pill.on{border:3px solid #1c2224;padding:8px 3px;font-weight:600}
.sec{padding:0 14px}.sec h3{font-size:20px;font-weight:700;margin:22px 0 10px}
.card{background:var(--card);border:1px solid var(--line);border-radius:14px;padding:12px 14px}
.wk{display:flex;gap:4px}.wk>div{flex:1;min-width:0;text-align:center}.wk>div.x{opacity:.4}.wk b{display:block;height:38px;line-height:38px;border-radius:8px;font-size:15px;font-weight:700;color:#fff}.wk .tri{display:flex;justify-content:center;align-items:flex-start;height:16px;padding-top:3px;line-height:0;font-size:0}.wk .tri u{display:block;border-left:6px solid transparent;border-right:6px solid transparent;border-bottom:9px solid #3b7a3e}.wk .tri u{background:none!important}.wk .tri u.L{border-bottom-color:#c9372c}.wk .tri u.D{border-bottom-color:#e49a9a}.wk .tri u.W{border-bottom-color:#9aa29d}.wk .tri u.X{border-bottom-color:#c5cbc7}.wk .dd{display:block;font-weight:700;font-size:14px;line-height:1.1}.wk .mo{display:block;font-size:12px;color:var(--mu)}.fc{font-size:12px;color:var(--mu);margin-bottom:10px}
.W{background:#e9ede9;color:#26302a!important}.L{background:#c9372c}.D{background:#f6c4c4;color:#26302a!important}.X{background:#e9ede9;color:#9aa29d!important}
.dt{display:flex;gap:14px;padding:13px 4px;border-bottom:1px solid var(--line);align-items:center}.dt:last-child{border:0}.dt .ic{width:30px;font-size:22px;text-align:center}.dt .t{font-size:16.5px}.dt .t small{display:block;color:var(--mu);font-size:14px}
table{width:100%;border-collapse:collapse}th{font-weight:400;color:var(--mu);font-size:13.5px;text-align:left;padding:4px 4px 10px}td{padding:11px 4px;font-size:15.5px}tr.me td{font-weight:700}tr.me{background:#f3f8f4}td.r,th.r{text-align:right}
.dot{display:inline-block;width:9px;height:9px;border-radius:50%;border:2px solid #9aa29d;margin:0 6px}.dot.k{border-color:var(--red);background:var(--red)}
.rate{display:inline-block;min-width:54px;text-align:center;padding:4px 8px;border-radius:999px;font-weight:700}.rate.g{background:#1e6b36;color:#fff}
.ev{display:flex;gap:12px;align-items:flex-start;padding:11px 0;border-bottom:1px solid var(--line)}.ev:last-child{border:0}.mi{min-width:56px;font-weight:700;color:var(--mu);font-size:13.5px;text-align:right}
.ico{width:26px;height:26px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:13px;color:#fff;flex:none}.ico.a{background:var(--acc)}.ico.s{background:#c9372c;border-radius:7px}.ico.cL{background:#c9372c}.ico.cD{background:#f6c4c4;color:#c9372c}.ico.cW{background:#e9ede9;color:#26302a}.ico.cX{background:#e9ede9;color:#9aa29d}.ico.trg{background:none;border-radius:0}.ico.trg.cL u{border-bottom-color:#c9372c}.ico.trg.cD u{border-bottom-color:#e49a9a}.ico.trg.cW u{border-bottom-color:#9aa29d}.ico.trg.cX u{border-bottom-color:#c5cbc7}.ico.trg u{display:block;border-left:8px solid transparent;border-right:8px solid transparent;border-bottom:13px solid #3b7a3e}.ico.w{background:var(--tx)}.ico.k{background:#8a948c}
.et{font-weight:600;font-size:15px}.ed{font-size:13px;color:var(--mu)}
.mm{display:flex;justify-content:space-between;font-size:12px;color:var(--mu)}
.kv div{padding:6px 0;border-bottom:1px solid var(--line)}.kv span{color:var(--mu);font-size:13px;display:block}.kv div:last-child{border:0}
.empty{color:var(--mu);text-align:center;padding:30px 0}
.mc .wk{margin:12px 0 4px}.w16s{overflow-x:auto;-webkit-overflow-scrolling:touch;margin:0 -4px;padding:0 4px}.wk.w16>div{flex:0 0 40px}

.pn{background:var(--card);border:1px solid var(--line);border-radius:16px;padding:9px 12px;margin:0 0 4px;overflow-y:auto;overflow-x:hidden;touch-action:pan-y;-webkit-overflow-scrolling:touch}
.pn .ph{display:flex;align-items:baseline;gap:8px;margin-bottom:3px}.pn .ph b{font-size:22px;color:var(--red)}.pn .ph span{color:var(--mu);font-size:12.5px}
.pn table{border-collapse:collapse;width:100%;table-layout:fixed}.pn th{padding:4px 2px;color:var(--mu);font-weight:500;font-size:10.5px;text-align:right;position:sticky;top:-9px;background:var(--card)}.pn th:first-child{text-align:left;width:40%}.pn td{padding:8px 2px 0;vertical-align:top;font-size:13.5px}.pn td.r{text-align:right;font-weight:600;overflow-wrap:normal}.pn td.mt b{font-size:13.5px;line-height:1.2;display:block}.pn td.sub{padding:3px 2px 8px;border-bottom:1px solid var(--line);font-size:11.5px;line-height:1.35}.pn td.sub .z{color:var(--red)}.pn td.sub .nx{color:var(--mu)}.pn td.sub .z b{font-weight:500}.pn tbody tr.tp{cursor:pointer}.pn .sm{color:var(--mu);font-size:12px;font-weight:400}.pn .rd{color:var(--red);font-weight:700}
.cal{margin-bottom:14px}.cal h4{margin:0 0 6px;font-size:15px;display:flex;justify-content:space-between;align-items:baseline}.cal h4 span{font-size:12px;font-weight:400;color:var(--mu)}
.cg{display:grid;grid-template-columns:repeat(7,1fr);gap:3px}.cg i{font-style:normal;text-align:center;font-size:11px;color:var(--mu);padding-bottom:2px}
.cg div{aspect-ratio:1;border-radius:8px;background:#eef0ee;display:flex;flex-direction:column;align-items:center;justify-content:center;font-size:13px;line-height:1.1;font-weight:600;min-width:0}
.cg div small{font-size:10px;font-weight:700;margin-top:1px}.cg .o{background:none;color:#c5cbc7;font-weight:400}.cg .t{outline:2px solid var(--tx)}.cg .z{color:#fff;background:#c9372c}.cg .zb{color:#fff;background:#3b7dd8}.sbx{color:#2f6fc6}.mu{color:var(--mu)}.cg u{display:block;border-left:4px solid transparent;border-right:4px solid transparent;border-bottom:6px solid #F5A623;margin-top:2px}.cg .z u{border-bottom-color:#FFD54F}.mrow .sq{display:inline-block;width:9px;height:9px;border-radius:2px;background:#c9372c;margin-right:7px}.mrow u.ar{display:inline-block;border-left:5px solid transparent;border-right:5px solid transparent;border-bottom:8px solid #F5A623;margin-right:6px}.mrow{padding:3px 0;font-size:14px;cursor:pointer}.dl2{align-items:flex-start}.cg .p{outline:3px solid #1c2224}.cg .zk{cursor:pointer}.wv{margin:10px 0}.wv .wh{display:flex;justify-content:space-between;align-items:center;cursor:pointer;font-weight:700;font-size:14px;margin-bottom:6px}.wv .wh span{font-weight:400;font-size:12px;color:var(--mu)}.wv .wh u{display:inline-block;border-left:5px solid transparent;border-right:5px solid transparent;border-bottom:8px solid #F5A623;margin-left:6px}.wv .wh .sq{display:inline-block;width:9px;height:9px;border-radius:2px;background:#c9372c;margin-left:6px}.cg .f{opacity:.4}.wm{padding:7px 0}.wm>div{margin:1px 0}.wm .sq{margin-right:7px}.wv .wl{margin-top:8px;border-top:1px solid var(--line);padding-top:6px}
.cl{display:flex;gap:10px;flex-wrap:wrap;font-size:12px;color:var(--mu);margin:4px 0 12px}.cl i{display:inline-block;width:12px;height:12px;border-radius:4px;vertical-align:-2px;margin-right:4px}
.dl .ev{cursor:pointer}.ico.po{background:#6a4fc7!important;color:#fff!important;font-size:11px;font-weight:700}.ico.pr{background:#9aa3ae!important;color:#fff!important;font-size:11px;font-weight:700}.ppx{color:#6a4fc7}.ico.sbt{background:#3b7dd8!important;color:#fff!important}.ico.trg.sbt u{border-bottom-color:#fff!important}
.dch .lg,.dtb h4{margin:0}.dtb{margin-top:8px}.dch{margin-top:0}
</style></head><body><div class="pg">
<div class="tt"><h1>Material Daily Update</h1><div class="sub" id="sub"></div></div>
<div id="sumv" style="display:none"><div class="pn" id="w21"></div></div>
<div id="calv" style="display:none"><div class="card" style="margin-top:12px"><div class="cl"><span><i style="background:#eef0ee"></i>no stock-zero</span><span><i style="background:#c9372c"></i>at least one RM material at stock zero</span><span><i style="background:#3b7dd8"></i>RM own stock zero, covered by its substitute</span><span><i style="background:#9aa39d"></i>PM / CS stock zero: listed below, day not coloured</span><span><u style="display:inline-block;border-left:5px solid transparent;border-right:5px solid transparent;border-bottom:8px solid #F5A623"></u> PPC arrival that day</span><span>tap a day to see the materials</span></div><div id="cal"></div></div><div id="cald"></div></div>
<div id="wkv" style="display:none"><div class="card" style="margin-top:12px"><div class="cl"><span><i style="background:#eef0ee"></i>no stock-zero</span><span><i style="background:#c9372c"></i>an RM material at stock zero</span><span><i style="background:#3b7dd8"></i>RM covered by a substitute</span><span><i style="background:#9aa39d"></i>PM / CS stock zero (listed, day not coloured)</span><span><u style="display:inline-block;border-left:5px solid transparent;border-right:5px solid transparent;border-bottom:8px solid #F5A623"></u> PPC arrival</span><span>tap a week to open it</span></div></div><div id="wkd"></div></div>
<input id="q" placeholder="Search material" aria-label="Search"><div id="out"></div></div>
<div class="fade"></div><div class="nav" id="nav"></div>
<div class="ov" id="ov"><div class="shw" id="sh"></div></div>
<script>
const D=__DATA__,M=D.meta,R=D.rows,$=s=>document.querySelector(s);R.forEach(r=>r.g=["RM","PM","CS"][r.g]||"");
const f1=v=>v==null?'-':(Math.abs(v)>=100?Math.round(v).toLocaleString('en-US'):(Math.round(v*10)/10).toString());
const q0=v=>v==null?'-':Math.round(v).toLocaleString('en-US');
const dd=k=>{const d=new Date(M.d0+'T00:00:00');d.setDate(d.getDate()+k);return d.toLocaleDateString('en-GB',{day:'2-digit',month:'short'})};
const STAT=['Stock-out in lead time','Stock-out later','Under 21 days','Under 21 days later','Low cover'];
const stl=r=>r.k<4?STAT[r.k]:(r.il?'Low cover':(r.k==5?'No demand':'OK'));   // all materials listed (25-Nov-2026): ALERT OK / NO DEMAND keep their own label (checker finding)
let F='all',Q='',TAB='tk',CUR=0;
const brk=v=>String(v).replace(/\s+;\s+PR ETA/g,' ;<br>PR ETA');   // 'PO ETA: ...  ;' then 'PR ETA: ...' on the next line (user 25-Nov-2026 'for easy read')
const sxn=r=>(r.sx||[]).map(x=>x.n).join(' / ');
const cvd_=r=>new Set(r.zc||[]);
const zr=r=>{const c=cvd_(r);return r.zd.filter(k=>!c.has(k))};   // stock-zero days still open WITH the substitute
const sxq=(r,a,b)=>(r.sx||[]).map(x=>[x,x.us.filter(y=>y[0]>=a&&y[0]<=b)]).filter(t=>t[1].length);
const sxar=(r,a,b)=>(r.sx||[]).flatMap(x=>x.ar.filter(y=>y[0]>=a&&y[0]<=b).map(y=>[x,y]));
const usr=(us)=>rn(us.map(y=>y[0])).map(g=>[g[0],g[g.length-1],us.filter(y=>y[0]>=g[0]&&y[0]<=g[g.length-1]).reduce((s,y)=>s+y[1],0)]);
const w16=(r,vs)=>`<div class="w16s">${wkb(r,vs,r.wa16,'w16')}</div>`;   // next 16 weeks, scroll sideways (Overview + Arrival, user 25-Nov-2026)
const ppStrip=(r,b)=>(b?r.wkp16:r.wkp)?`<div class="sub2 ppx" style="margin-top:6px">If the open PO + PR lines arrive on their ETA (with PPC arrivals) · information only</div>${b?w16(r,r.wkp16):wkb(r,r.wkp)}`:'';
const subStrip=(r,b)=>(b?r.wks16:r.wks)?`<div class="sub2" style="color:#2f6fc6;margin-top:6px">⇄ With substitute ${sxn(r)} (days of cover incl. its stock)</div>${b?w16(r,r.wks16):wkb(r,r.wks)}`:'';
function wkb(r,vs,wa,cl){const L=Math.floor(r.lt/7),A=wa||r.wa;return `<div class="wk${cl?' '+cl:''}">${(vs||r.wk).map((v,w)=>{const c=wl(v),d=dd(7*w).split(' ');return `<div class="${w>L?'x':''}"><b class="${c}">${v==null?'-':Math.max(0,Math.floor(v))}</b><span class="tri">${A[w]>0?`<u class="${c}"></u>`:''}</span><span class="dd">${d[0]}</span><span class="mo">${d[1]}</span></div>`}).join('')}</div>`}
const tl=wkb;   // every bar = the Shortage_Board week strip (user 22-Nov-2026 "I want all the bar look like this")
const lead=r=>r.runs.length&&r.runs[0][0]<=r.lt;
function draw(){const rs=R.map((r,i)=>[r,i]).filter(([r])=>(F=='all'||(F=='lt'&&r.k==0)||(F=='low'&&r.k!=0&&r.il))&&(!Q||(r.n+' '+r.g+' '+r.wh).toLowerCase().includes(Q)));
 $('#out').innerHTML=rs.length?rs.map(([r,i])=>`<div class="mc" data-i="${i}"><div class="row"><div><span class="nm">${r.n}</span><span class="tg">${r.g}</span>
 <div class="sub2"><span class="${r.k==0?'rd':''}">${stl(r)}</span></div><div class="sub2">Usage/Day ${f1(r.nd)} ${r.u} · stock ${q0(r.st)}</div>${r.bu?`<div class="sub2">BOM usage <b>${r.bu}</b> / module</div>`:''}<div class="sub2">Cover + next arrival: <b>${f1(r.dw)}</b> days</div></div><div class="sc${r.dy<7?' r':''}">${f1(r.dy)}<small>Days of cover (stock only)</small></div></div>${tl(r)}
 <div class="sub2${r.per?' rd':''}">${r.per?'Stock zero '+r.per:'No stock-zero day in '+r.N+' days'}</div>${subStrip(r)}${r.sn?`<div class="sub2" style="color:#2f6fc6">⇄ ${r.sn}</div>`:''}</div>`).join(''):'<div class="empty">No materials match.</div>'}
let CK=null;
function calHtml(rs,one){const H=M.CH,T0=M.to,d0=new Date(M.d0+'T00:00:00'),t0=new Date(d0);t0.setDate(t0.getDate()+T0);const st=new Date(t0);st.setDate(st.getDate()-(t0.getDay()+6)%7);const nw=Math.ceil(((t0-st)/864e5+H)/7),z=new Set(),B=new Set(),A=new Set(),P=new Set();/* red / blue day = an RM material only (user 25-Nov-2026); PM / CS stock zero = P: listed, day not coloured */rs.forEach(([r])=>{const rm=one||r.g=='RM';zr(r).forEach(k=>{if(k>=T0&&k<T0+H)(rm?z:P).add(k)});(r.zc||[]).forEach(k=>{if(k>=T0&&k<T0+H)(rm?B:P).add(k)});r.ar.forEach(x=>{if(x[0]>=T0&&x[0]<T0+H)A.add(x[0])});if(one)sxar(r,T0,T0+H-1).forEach(([x,y])=>A.add(y[0]))});
 let c=['Mo','Tu','We','Th','Fr','Sa','Su'].map(x=>'<i>'+x+'</i>').join('');
 for(let n=0;n<nw*7;n++){const dt=new Date(st);dt.setDate(dt.getDate()+n);const k=Math.round((dt-d0)/864e5),lb=dt.getDate()==1||n==0?`<small>${dt.toLocaleDateString('en-GB',{month:'short'})}</small>`:'';
  if(k<T0||k>=T0+H){c+=`<div class="o">${dt.getDate()}${lb}</div>`;continue}
  const tri=A.has(k)?'<u></u>':'',cl=(z.has(k)?'z ':(B.has(k)?'zb ':''))+((z.has(k)||B.has(k)||A.has(k)||P.has(k))&&!one?'zk ':'')+(k==T0?'t ':'')+(k==CK?'p':'');
  c+=`<div class="${cl}" data-k="${k}">${dt.getDate()}${lb}${tri}</div>`}
 return `<div class="cal"><div class="cg">${c}</div></div>`}
function calDraw(){const H=M.CH,T0=M.to,rs=R.map((r,i)=>[r,i]).filter(([r])=>!Q||(r.n+' '+r.g+' '+r.wh).toLowerCase().includes(Q));
 $('#cal').innerHTML=calHtml(rs,false);
 let h='';for(let k=T0;k<T0+H;k++){if(CK!=null&&CK!=k)continue;const l=rs.filter(([r])=>r.zd.includes(k)),ar=rs.map(([r,i])=>[r,i,r.ar.filter(x=>x[0]==k).reduce((s,x)=>s+x[1],0)]).filter(x=>x[2]>0),gv=rs.filter(([r])=>r.gv&&r.gv.us.some(y=>y[0]==k));if(!l.length&&!ar.length&&!gv.length)continue;
  h+=`<div class="ev dl2"><div class="mi">${dd(k)}</div><div style="flex:1">${l.map(([r,i])=>`<div class="mrow" data-j="${i}">${(r.zc||[]).includes(k)?`<span class="sq" style="background:#3b7dd8"></span><b>${r.n}</b> <span class="tg">${r.g}</span> <span style="color:#2f6fc6">own stock zero, covered by ${sxq(r,k,k).map(([x,u])=>x.n+' '+q0(u[0][1])+' '+x.u).join(' / ')}</span>`:`<span class="sq"${r.g=='RM'?'':' style="background:#9aa39d"'}></span><b>${r.n}</b> <span class="tg">${r.g}</span> <span class="${r.g=='RM'?'rd':'mu'}">stock zero${r.sx&&r.sx.length?' (also with '+sxn(r)+')':''}</span>`}</div>`).join('')}${gv.map(([r,i])=>`<div class="mrow" data-j="${i}"><span class="sq" style="background:#3b7dd8"></span><b>${r.n}</b> <span class="tg">${r.g}</span> <span style="color:#2f6fc6">used for ${r.gv.n}: ${q0(r.gv.us.find(y=>y[0]==k)[1])} ${r.u}</span></div>`).join('')}${ar.map(([r,i,q])=>`<div class="mrow" data-j="${i}"><u class="ar"></u><b>${r.n}</b> <span class="tg">${r.g}</span> arrival <b>+${q0(q)} ${r.u}</b></div>`).join('')}</div></div>`}
 $('#cald').innerHTML=`<h3 style="font-size:18px;margin:16px 0 8px">${CK!=null?dd(CK)+' · tap the day again to see all':'Stock zero and arrivals, by day'}</h3><div class="card">${h||'<div class="empty">No stock-zero day or arrival in the next 4 weeks.</div>'}</div>`}
$('#calv').addEventListener('click',e=>{const j=e.target.closest('[data-j]');if(j){TAB='ov';show(+j.dataset.j);$('#ov').scrollTop=0;return}const c=e.target.closest('.zk');if(c){CK=CK==+c.dataset.k?null:+c.dataset.k;calDraw()}});
let WO=new Set([0]);
const WD=['Mon','Tue','Wed','Thu','Fri','Sat','Sun'];
const rn=z=>z.reduce((o,k)=>{o.length&&o[o.length-1][o[o.length-1].length-1]==k-1?o[o.length-1].push(k):o.push([k]);return o},[]);
function wkHtml(rs,one){const T0=M.to,d0=new Date(M.d0+'T00:00:00'),t0=new Date(d0);t0.setDate(t0.getDate()+T0);const st=new Date(t0);st.setDate(st.getDate()-(t0.getDay()+6)%7);const k0=Math.round((st-d0)/864e5);let h='';
 for(let w=0;w<M.WW;w++){const a=k0+7*w,z=new Set(),B=new Set(),A=new Set(),ln=[];
  rs.forEach(([r,i])=>{const lo=Math.max(a,T0),cs=cvd_(r),zs=r.zd.filter(k=>k>=lo&&k<a+7&&!cs.has(k)),zb=r.zd.filter(k=>k>=lo&&k<a+7&&cs.has(k)),as=r.aw.filter(x=>x[0]>=lo&&x[0]<a+7),sa=one?sxar(r,lo,a+6):[],pq=(r.pp||[]).filter(x=>x[0]>=lo&&x[0]<a+7),su=sxq(r,lo,a+6),gu=r.gv?r.gv.us.filter(y=>y[0]>=lo&&y[0]<a+7):[];
   if(one||r.g=='RM'){zs.forEach(k=>z.add(k));zb.forEach(k=>B.add(k))}as.forEach(x=>A.add(x[0]));sa.forEach(([x,y])=>A.add(y[0]));if(zs.length||zb.length||as.length||sa.length||su.length||gu.length||pq.length)ln.push([r,i,zs,as,zb,sa,su,gu,pq])});
  let c='';for(let n=0;n<7;n++){const k=a+n,dt=new Date(d0);dt.setDate(dt.getDate()+k);const lb=dt.getDate()==1||(w==0&&n==0)?`<small>${dt.toLocaleDateString('en-GB',{month:'short'})}</small>`:'';c+=`<div class="${k<T0?'f ':''}${z.has(k)?'z ':(B.has(k)?'zb ':'')}${k==T0?'t':''}">${dt.getDate()}${lb}${A.has(k)?'<u></u>':''}</div>`}
  const op=one||WO.has(w),body=op?`<div class="wl">${ln.map(([r,i,zs,as,zb,sa,su,gu,pq])=>`<div class="mrow wm"${one?'':` data-j="${i}"`}>${one?'':`<div><b>${r.n}</b> <span class="tg">${r.g}</span></div>`}${zs.length?`<div class="${one||r.g=='RM'?'rd':'mu'}"><i class="sq"${one||r.g=='RM'?'':' style="background:#9aa39d"'}></i>stock zero ${rn(zs).map(g=>g.length==1?dd(g[0]):dd(g[0])+' – '+dd(g[g.length-1])).join(', ')}</div>`:''}${as.length?`<div><u class="ar"></u>arrival ${as.map(x=>'<b>+'+q0(x[1])+' '+r.u+'</b> '+dd(x[0])).join(', ')}</div>`:''}${su.map(([x,u])=>`<div class="sbx"><i class="sq" style="background:#3b7dd8"></i>covered by ${x.n}: ${usr(u).map(g=>(g[0]==g[1]?dd(g[0]):dd(g[0])+' – '+dd(g[1]))+' <b>'+q0(g[2])+' '+x.u+'</b>').join(', ')}</div>`).join('')}${sa.map(([x,y])=>`<div class="sbx"><u class="ar"></u>substitute ${x.n} arrival <b>+${q0(y[1])} ${x.u}</b> ${dd(y[0])}</div>`).join('')}${pq.length?`<div class="ppx">${['PO','PR'].map(t=>pq.filter(x=>x[1]==t)).filter(l=>l.length).map(l=>l[0][1]+' ETA '+l.map(x=>'<b>+'+q0(x[2])+' '+r.u+'</b> '+dd(x[0])).join(', ')).join(' · ')}</div>`:''}${gu.length?`<div class="sbx"><i class="sq" style="background:#3b7dd8"></i>used for ${r.gv.n}: ${usr(gu).map(g=>(g[0]==g[1]?dd(g[0]):dd(g[0])+' – '+dd(g[1]))+' <b>'+q0(g[2])+' '+r.u+'</b>').join(', ')}</div>`:''}</div>`).join('')||'<div class="empty" style="padding:6px 0">No stock-zero day or arrival this week.</div>'}</div>`:'';
  h+=`<div class="card wv"><div class="wh"${one?'':` data-w="${w}"`}><div>Week ${w+1} · ${dd(a)} – ${dd(a+6)}${z.size?'<i class="sq"></i>':''}${B.size?'<i class="sq" style="background:#3b7dd8"></i>':''}${A.size?'<u></u>':''}</div>${one?'':`<span>${op?'hide':'show'}</span>`}</div><div class="cg">${c}</div>${body}</div>`}
 return h}
function wkDraw(){$('#wkd').innerHTML=wkHtml(R.map((r,i)=>[r,i]).filter(([r])=>!Q||(r.n+' '+r.g+' '+r.wh).toLowerCase().includes(Q)),false)}
$('#wkv').addEventListener('click',e=>{const j=e.target.closest('[data-j]');if(j){TAB='ov';show(+j.dataset.j);$('#ov').scrollTop=0;return}const t=e.target.closest('[data-w]');if(t){const w=+t.dataset.w;WO.has(w)?WO.delete(w):WO.add(w);wkDraw()}});
function wl(v){return v==null?'X':v>=21?'W':v>=7?'D':'L'}
function form(r){return `<div class="card"><div class="fc">Next ${M.WW} weeks · days of cover on the first day of each week · ▲ = PPC arrival that week · scroll sideways${r.wks?' · own stock':''}</div>${w16(r,r.wk16)}${subStrip(r,1)}${ppStrip(r,1)}</div>`}
function subCard(r){if(!(r.sx&&r.sx.length)&&!r.gv)return '';let h='';
 (r.sx||[]).forEach(x=>{const g=usr(x.us);h+=`<div class="dt"><div class="ic">⇄</div><div class="t">Substitute ${x.n}<small>${g.length?'covers '+g.map(t=>(t[0]==t[1]?dd(t[0]):dd(t[0])+' → '+dd(t[1]))+' '+q0(t[2])+' '+x.u).join(' · ')+' (total '+q0(x.us.reduce((s,y)=>s+y[1],0))+' '+x.u+')':'not used'}<br>${x.ar.length?'its arrivals: '+x.ar.map(y=>dd(y[0])+' +'+q0(y[1])+' '+x.u).join(' · '):'nothing in transit'}<br>${zr(r).length?'<span class="rd">still stock zero with it: '+rn(zr(r)).map(g=>g.length==1?dd(g[0]):dd(g[0])+' → '+dd(g[g.length-1])).slice(0,4).join(', ')+'</span>':'no stock-zero day left with it'}</small></div></div>`});
 if(r.gv){const g=usr(r.gv.us);h+=`<div class="dt"><div class="ic">⇄</div><div class="t">Substitute for ${r.gv.n}<small>${g.length?'used '+g.map(t=>(t[0]==t[1]?dd(t[0]):dd(t[0])+' → '+dd(t[1]))+' '+q0(t[2])+' '+r.u).join(' · '):'not needed / nothing left to give'}</small></div></div>`}
 return `<div class="card">${h}</div>`}
function show(i){const r=R[i];CUR=i;const L=lead(r),stt=stl(r);
 const ev=[];r.ar.forEach((x,i)=>ev.push([x[0],'a',x[1],wl(r.ac[i][0])]));r.runs.forEach(x=>ev.push([x[0],'s',x[1]-x[0]+1]));(r.sx||[]).forEach(x=>usr(x.us).forEach(g=>ev.push([g[0],'b',g[2],x])));ev.sort((a,b)=>a[0]-b[0]);
 const sg=ev.slice(0,6).map(x=>x[1]=='b'?`<div class="l">${x[3].n.split(' ').slice(0,2).join(' ')} ${q0(x[2])} <b>${dd(x[0])}</b></div><div class="m" style="color:#7fb2ff">⇄</div><div></div>`:x[1]=='a'?`<div class="l">+${q0(x[2])} ${r.u} <b>${dd(x[0])}</b></div><div class="m c${x[3]}">▲</div><div></div>`:`<div></div><div class="m k">■</div><div class="rr"><b>${dd(x[0])}</b> ${r.sx&&r.sx.length?'own ':''}stock-out ${x[2]} d</div>`).join('');
 let h=`<div class="hero"><div class="bt"><button class="cb" onclick="$('#ov').classList.remove('on')">←</button></div>
 <div class="bd0"><div><div class="n${r.dy<7?' r':''}">${f1(r.dy)}</div><div class="tn">Days of cover (stock only)</div><div class="xg">${q0(r.st)} ${r.u} in stock</div></div>
 <div class="mid"><div class="chip${r.k==0?' rc':''}">${stt}</div><div class="vs">${r.n}</div></div>
 <div><div class="n${r.dw!=null&&r.dw<7?' r':''}">${f1(r.dw)}</div><div class="tn">Days of cover (+ next arrival only)</div><div class="xg">${r.ar.length?q0(r.ar[0][1]):0} ${r.u} next arrival</div></div></div>
 <div class="sg">${sg}</div></div>
 <div class="pills">${[['ov','Overview'],['tk','Arrival'],['wk','Weeks']].map(t=>`<button class="pill${t[0]==TAB?' on':''}" onclick="TAB='${t[0]}';show(${i})">${t[1]}</button>`).join('')}</div><div class="sec">`;
 if(TAB=='ov'){
  h+=`<h3>Coverage</h3>${form(r)}${(r.sx&&r.sx.length)||r.gv?'<h3>Substitute</h3>'+subCard(r):''}<h3>Stock-zero calendar</h3><div class="card">${calHtml([[r,i]],true)}</div>`;
  h+=`<h3>Details</h3><div class="card"><div class="dt"><div class="ic">🏭</div><div class="t">Supplier / sheet<small>${r.wh||r.ata[0]||'-'}</small></div></div>
   <div class="dt"><div class="ic">📅</div><div class="t">Base date ${M.d0t}<small>built ${M.run}</small></div></div>
   <div class="dt"><div class="ic">🏷</div><div class="t">${r.g} · ${stt}<small>Matchday = ${r.per?'stock zero '+r.per:'no stock-zero day'}</small></div></div>
   <div class="dt"><div class="ic">⚙</div><div class="t">Usage/Day ${f1(r.nd)} ${r.u}<small>Estimated stock ${q0(r.st)} ${r.u}</small></div></div>
   ${r.bu?`<div class="dt"><div class="ic">🧩</div><div class="t">Per module ${r.bu}<small>${(r.bl||[]).join('<br>')}</small></div></div>`:''}
   ${r.sb?`<div class="dt"><div class="ic">⇄</div><div class="t">Covered by substitute<small>${r.sb}</small></div></div>`:''}
   <div class="dt"><div class="ic">🚚</div><div class="t">${r.nx?'Next arrival '+r.nx[0]+' · '+q0(r.nx[1])+' '+r.u:'No next arrival'}<small>Days of cover (+ next arrival only) ${f1(r.dn)} · incl. in transit ${f1(r.di)}</small></div></div><div class="dt"><div class="ic">#</div><div class="t">${r.c}<small>${r.d||'No BOM description'}</small></div></div></div>`;
  h+=`<h3>BOM Material Dashboard</h3><div class="card">${r.bd&&r.bd.length?`<table>${r.bd.map(x=>`<tr><td style="color:var(--mu);font-size:12.5px;width:42%">${x[0]}</td><td style="font-size:13px">${brk(x[1])}</td></tr>`).join('')}</table>`:'<div class="empty">Not on the BOM Material Dashboard of 4_Status.</div>'}</div>`;
  if(r.ch)h+=`<h3>Projected stock, arrivals and open deliveries</h3>${detBlock(r)}`;
  const lo=Math.max(0,Math.min(i-2,R.length-5)),nb=R.slice(lo,lo+5).map((x,k)=>[x,lo+k]);
  h+=`<h3>Live table</h3><div class="card"><table><tr><th>#</th><th>Material</th><th class="r">Days</th><th class="r">Stock</th></tr>${nb.map(([x,k])=>`<tr class="${k==i?'me':''}"><td>${k+1}<span class="dot${x.k==0?' k':''}"></span></td><td>${x.n}</td><td class="r">${f1(x.dy)}</td><td class="r">${q0(x.st)}</td></tr>`).join('')}</table></div>`;
  const ta=[...r.ar].sort((a,b)=>b[1]-a[1]).slice(0,3);
  h+=`<h3>Top arrivals</h3><div class="card">${ta.length?`<table><tr><th>#</th><th>Arrival</th><th class="r">Qty</th></tr>${ta.map((x,k)=>`<tr><td>${k+1}</td><td>${dd(x[0])}${x[0]>r.lt?' (after lead time)':''}</td><td class="r"><span class="rate${k==0?' g':''}">${q0(x[1])}</span></td></tr>`).join('')}</table>`:'<div class="empty">No PPC arrivals in the horizon.</div>'}</div>`;}
 if(TAB=='wk'){h+=`${(r.sx&&r.sx.length)||r.gv?'<h3>Substitute</h3>'+subCard(r):''}<h3>Next ${M.WW} weeks</h3>${wkHtml([[r,i]],true)}`}
 if(TAB=='tk'){const all=[[0,'k','K','Kick-off',`Stock ${q0(r.st)} ${r.u} · Usage/Day ${f1(r.nd)} ${r.u}`]];
  r.ar.forEach((x,i)=>{const c=r.ac[i][0],sh=r.ac[i][1],cs=c==null?' · <b>no demand after this arrival</b>':' · <b>Days of cover after this arrival '+f1(c)+'</b>'+(sh>0?' (first production day '+dd(x[0]+sh)+')':'');all.push([x[0],'trg c'+wl(c),'<u></u>','Arrival +'+q0(x[1])+' '+r.u,dd(x[0])+(x[0]<=r.lt?' · inside lead time':' · after lead time')+cs])});
  r.runs.forEach(x=>all.push([x[0],'s','',(r.sx&&r.sx.length?'Own stock-out (own stock only)':'Stock-out'),dd(x[0])+(x[1]>x[0]?' to '+dd(x[1]):'')+' · '+(x[1]-x[0]+1)+' day'+(x[1]>x[0]?'s':'')+(x[0]<=r.lt?' · starts inside lead time':'')]));
  (r.sx||[]).forEach(x=>{x.ar.filter(y=>y[0]<r.N).forEach(y=>all.push([y[0],'trg sbt','<u></u>','Substitute '+x.n+' arrival +'+q0(y[1])+' '+x.u,dd(y[0])+' · its own plan first, the rest can cover '+r.n]));
   usr(x.us).filter(g=>g[0]<r.N).forEach(g=>all.push([g[0],'sbt','⇄','Covered by '+x.n,dd(g[0])+(g[1]>g[0]?' to '+dd(g[1]):'')+' · <b>'+q0(g[2])+' '+x.u+'</b> of the substitute used']))});
  if(r.sx&&r.sx.length)rn(zr(r).filter(k=>k<r.N)).forEach(g=>all.push([g[0],'s','','Stock-out even with '+sxn(r),dd(g[0])+(g.length>1?' to '+dd(g[g.length-1]):'')]));
  (r.pp||[]).forEach((x,j)=>{const c=(r.ppc||[])[j]||[null,0],cs=c[0]==null?' · no demand after it':' · <b>Days of cover if it arrives '+f1(c[0])+'</b>'+(c[1]>0?' (first production day '+dd(Math.max(0,x[0])+c[1])+')':'')+' (with PPC arrivals + earlier PO / PR)';
   all.push([Math.max(0,x[0]),(x[1]=='PO'?'po':'pr'),x[1],x[1]+' ETA +'+q0(x[2])+' '+r.u,(x[0]<0?'date passed: '+dd(x[0])+' · ':dd(x[0])+' · ')+(x[1]=='PO'?'open PO, not shipped yet':'requisition, no PO yet')+cs])});
  if(r.pu)all.push([0,'po','PO','PO, no delivery date: '+q0(r.pu)+' '+r.u,'open PO quantity without a date · information, not in the cover']);
  if(r.gv)usr(r.gv.us).filter(g=>g[0]<r.N).forEach(g=>all.push([g[0],'sbt','⇄','Used for '+r.gv.n,dd(g[0])+(g[1]>g[0]?' to '+dd(g[1]):'')+' · <b>'+q0(g[2])+' '+r.u+'</b>']));
  all.push([Math.min(r.lt,r.N-1),'w','⏱','Lead-time end',dd(Math.min(r.lt,r.N-1))+' · orders placed today arrive after this']);all.sort((a,b)=>a[0]-b[0]);
  const ab=(r.bd||[]).filter(x=>/^(Arrival at plant|ETA - PO|ETA - PR|Pending GRN|POD|In transit STATUS)/.test(x[0]));
  h+=`<h3>Arrival</h3><div class="card"><div class="fc">Next ${M.WW} weeks · days of cover on the first day of each week · scroll sideways</div>${w16(r,r.wk16)}${subStrip(r,1)}${ppStrip(r,1)}${ab.length?`<table style="margin:8px 0 4px">${ab.map(x=>`<tr><td style="color:var(--mu);font-size:12.5px;width:38%">${x[0]}</td><td style="font-size:13px">${brk(x[1])}</td></tr>`).join('')}</table>`:''}${all.map(x=>`<div class="ev"><div class="mi">${dd(x[0])}</div><div class="ico ${x[1]}">${x[2]}</div><div><div class="et">${x[3]}</div><div class="ed">${x[4]}</div></div></div>`).join('')}</div>`}
 $('#sh').innerHTML=h+'</div><div style="height:40px"></div>';$('#ov').classList.add('on')}
$('#out').addEventListener('click',e=>{const c=e.target.closest('.mc');if(c){TAB='tk';show(+c.dataset.i);$('#ov').scrollTop=0}});
document.addEventListener('keydown',e=>{if(e.key=='Escape')$('#ov').classList.remove('on')});
const rdc=v=>v!=null&&v<7?' rd':'';
$('#w21').innerHTML=`<div class="ph"><b>${R.filter(r=>r.il).length}</b><span>materials within their transit lead time (days of cover, stock only) · tap a line</span></div><table><thead><tr><th>Material</th><th>Days of cover<br>(stock only)</th><th>Days of cover<br>(+ next arrival only)</th><th>Days of cover<br>(+ all in transit)</th></tr></thead><tbody>${R.map((r,i)=>[r,i]).filter(([r])=>r.il).map(([r,i])=>`<tr class="tp" data-j="${i}"><td class="mt"><b>${r.n}</b><span class="sm">${r.g}${r.wh?' \u00b7 '+r.wh:''}</span></td><td class="r${rdc(r.dy)}">${f1(r.dy)}</td><td class="r${rdc(r.dw)}">${f1(r.dw)}</td><td class="r${rdc(r.di)}">${f1(r.di)}</td></tr><tr class="tp" data-j="${i}"><td class="sub" colspan="4">${r.per?'<div class="z">'+r.per.replace(/ \u00b7 /g,' \u00b7 ')+'</div>':'<div class="nx">no stock-zero day</div>'}<div class="nx">next: ${r.a21.length?`${r.a21[0][0]} \u00b7 ${q0(r.a21[0][1])} ${r.u}${r.a21.length>1?` (+${r.a21.length-1} more, ${q0(r.a21.reduce((s,x)=>s+x[1],0))} within the lead time)`:''}`:'none within the lead time'}</div></td></tr>`).join('')}</tbody></table>`;
window.fitP=()=>{};(function(){const P=$('#w21'),fit=window.fitP=()=>{if(!P.offsetParent)return;P.style.maxHeight='';const rs=[...P.querySelectorAll('tbody tr')],h=P.querySelector('.ph').offsetHeight+P.querySelector('thead').offsetHeight+(P.offsetHeight-P.clientHeight)+14;P.style.maxHeight=Math.max(240,innerHeight-P.getBoundingClientRect().top-scrollY-100)+'px'};fit();addEventListener('resize',fit);
 P.onclick=e=>{const t=e.target.closest('tr[data-j]');if(t){TAB='ov';show(+t.dataset.j);$('#ov').scrollTop=0}}})();
$('#sub').textContent=`All ${R.length} materials (BOM Material Dashboard order) · ${R.filter(r=>r.il).length} need action: stock will not last until a new order could arrive · base date ${M.d0t}`;
const nl=R.filter(r=>r.k==0).length;
$('#nav').innerHTML=[['all','All',R.length],['sum','Summary',''],['cal','Calendar',''],['wk','Weeks',''],['lt','Runs out',nl],['low','Low cover',R.filter(r=>r.il&&r.k!=0).length]].map(c=>`<button class="${c[0]==F?'on':''}" data-f="${c[0]}">${c[1]} <small>${c[2]}</small></button>`).join('');
$('#nav').addEventListener('click',e=>{const b=e.target.closest('button');if(!b)return;F=b.dataset.f;document.querySelectorAll('#nav button').forEach(x=>x.classList.toggle('on',x===b));const sv=F=='sum',cv=F=='cal',wv=F=='wk';$('#sumv').style.display=sv?'block':'none';$('#calv').style.display=cv?'block':'none';$('#wkv').style.display=wv?'block':'none';$('#out').style.display=(sv||cv||wv)?'none':'';$('#q').style.display=sv?'none':'';if(sv)window.fitP();else if(cv)calDraw();else if(wv)wkDraw();else draw()});
$('#q').addEventListener('input',e=>{Q=e.target.value.trim().toLowerCase();F=='cal'?calDraw():F=='wk'?wkDraw():draw()});draw();
</script></body></html>""".replace('__DATA__',js_data)
    page=det_inject(page)
    stem=os.path.join(out_dir,outfile_name("web","html"))
    with open(stem,'w',encoding='utf-8') as f_: f_.write(page)
    for n_ in ([] if globals().get('KEEP_OLD') else sorted(os.listdir(out_dir))):
        if re.fullmatch(r'(Match_Centre|Material_Shortage|Material_Dashboard|Shortage_Board|'+out_rx('web')+'|'+out_rx('wsa')+r')_\d{4}-\d{2}-\d{2}\.html',n_) and n_!=os.path.basename(stem):
            try: os.remove(os.path.join(out_dir,n_))
            except OSError: pass
    print(f"MATERIAL DAILY UPDATE: {stem} (all {len(data)} materials of the BOM Material Dashboard, {sum(1 for x_ in data if x_['il'])} within the transit lead time, {sum(1 for x_ in data if x_['k']==0)} run out inside the lead time"
          +(f"; substitute shown on every tab for {', '.join(x_['c'] for x_ in data if x_['sx'] or x_['gv'])}" if any(x_['sx'] or x_['gv'] for x_ in data) else '')+')',flush=True)
    return stem

def wsa_open_deliveries(path, d0):
    """Open WSA deliveries per planning sheet of a 2_WSA file, for the Details of 5_Web (table 'Open WSA deliveries'; the
       other web pages are deleted, user 23-Nov-2026). [] for a file without aligned sheets. One dict per planning sheet: n = sheet name,
       codes = codes of row 5, dlv = open deliveries (Scheduled Delivery, not cancelled) from the base date, nd = their number. Reads the cached
       values the script wrote into 2_WSA (columns found BY HEADER) - presentation only, no calculation."""
    import html as _h
    d0=pd.Timestamp(d0).normalize(); e=_h.escape
    def num(v):
        try: return None if v is None or isinstance(v,(str,bool)) else float(v)
        except Exception: return None
    with warnings.catch_warnings():
        warnings.simplefilter('ignore')
        wb=load_workbook(path,data_only=True,read_only=True)
    sheets=[]
    try:
        for ws in wb.worksheets:
            if getattr(ws,'sheet_state','visible')!='visible': continue
            top=[list(r) for r in ws.iter_rows(min_row=1,max_row=WSA_HDR,values_only=True)]
            if len(top)<WSA_HDR or not str(top[WSA_HDR-1][0] or '').startswith('Required Plant Date'): continue
            hd=[str(x or '').strip() for x in top[WSA_HDR-1]]
            def col(*keys):
                for k_ in keys:
                    for i_,h_ in enumerate(hd):
                        if h_.lower().startswith(k_.lower()): return i_
                return None
            iC,iE,iV=col('Balance'),col('Demand'),col('Plant + Port Days (PPC')
            iX,iPO,iSh,iETD,iBL,iVE,iSt,iRm=col('Scheduled Delivery'),col('PO'),col('Shipper'),col('ETD'),col('BL#'),col('Vessel ETA'),col('Status'),col('Remarks')
            if None in (iC,iE,iV): continue
            g=lambda r_,i_: (r_[i_] if i_ is not None and i_<len(r_) else None)
            dlv=[]; seen=False
            for r_ in ws.iter_rows(min_row=WSA_HDR+1,values_only=True):
                dd=_any_date(r_[0]) if r_ and r_[0] is not None else None
                if dd is None: continue
                dd=pd.Timestamp(dd).normalize()
                if dd<d0: continue
                seen=True
                x_=num(g(r_,iX)); st_=str(g(r_,iSt) or '')
                if x_ and 'canc' not in (st_+' '+str(g(r_,iRm) or '')).lower():
                    etd_=_any_date(g(r_,iETD)) if g(r_,iETD) is not None else None
                    dlv.append([f"{dd:%d-%b-%y}",x_,e(str(g(r_,iPO) or '')),e(str(g(r_,iSh) or '')),
                                (f"{pd.Timestamp(etd_):%d-%b-%y}" if etd_ is not None else e(str(g(r_,iETD) or ''))),e(str(g(r_,iBL) or '')),e(str(g(r_,iVE) or '')),e(st_)])
            if not seen: continue
            sheets.append(dict(n=e(ws.title),dlv=dlv[:12],nd=len(dlv),codes=re.findall(r'[A-Z]{2,}[A-Z0-9]*\d{5,}',str(top[4][2] or '').upper())))
    finally:
        try: wb.close()
        except Exception as _e_: _swallow(0,_e_)
    return sheets
def wsa_alert_match(sheets, rows):
    """WSA planning sheet -> its ALERT row (the dashboard rows, `_heat_expand`): the row of that sheet (`where`, back glass 333 on 2
       sheets), else the first code of the sheet's row 5 (or its MRP / alias code) with an ALERT row. Sets s['_a'] (None = no ALERT row: no demand / not in the MRP)."""
    by_w={str(r.get('where')):r for r in rows if r.get('where')}
    by_c={}
    for r in rows:
        if not r.get('where'): by_c.setdefault(str(r.get('code') or '').strip().upper(),r)
    for s_ in sheets:
        cand=list(s_['codes'])                               # + the MRP code of a code typed on the sheet (busbar 032 -> 047, PACK_FACTORS / Settings
        for c_ in s_['codes']:                               # SUBSTITUTES) and the other roll length of a tape (BOM alias, either direction: 018 <-> 025)
            pf_=PACK_FACTORS.get(c_)
            if pf_: cand.append(str(pf_[0]).strip().upper())
            for a_,m_ in (BOM_ALIAS[0] or {}).items():
                a_=str(a_).strip().upper(); m_=str(m_).strip().upper()
                if a_==c_: cand.append(m_)
                if m_==c_: cand.append(a_)
        s_['_a']=by_w.get(_h_unesc(s_['n'])) or next((by_c[c_] for c_ in cand if c_ in by_c),None)
    return sheets
def _h_unesc(t):
    import html as _h
    return _h.unescape(str(t or ''))
def _add_it_sup_col(w, hr, at, rows_, code_col, text_col, dark=False, head=IT_SUP_HEAD, fn=None, oneline=False, width=None):
    """Insert the 'In transit supplier name (qty)' column before column `at` (header row hr, data rows rows_); freeze pane and
       autofilter follow. Codes: the code cell (several joined by ' / ' or ', '); the WSA sheet: '(WSA sheet: X)' or a
       '(... Back)' label of the text cell - only for a material on 2+ WSA sheets. head / fn(codes, sheet, own) = another column the
       same way (a number: 'Pending GRN Qty', `pending_grn_qty`)."""
    from copy import copy as _cp
    from openpyxl.utils import column_index_from_string as _CI
    shared={}
    for g_ in WSA_REFS[0] or []:
        for c_ in g_.get('codes') or []: shared.setdefault(c_,[]).append(g_.get('sheet'))
    shared={c_:v_ for c_,v_ in shared.items() if len(v_)>1}
    fz=w.freeze_panes; flt=w.auto_filter.ref
    _ins_cols(w,at,1)
    h=w.cell(hr,at,head); h._style=_cp(w.cell(hr,at-1)._style); h.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
    if hr>1 and w.cell(hr-1,at-1).has_style and not w.cell(hr-1,at).has_style: w.cell(hr-1,at)._style=_cp(w.cell(hr-1,at-1)._style)
    n_=0
    for r_ in rows_:
        codes=[c_ for c_ in re.split(r'\s*(?:/|,|\+|\n)\s*',str(w.cell(r_,code_col).value or '')) if IT_CODE.match(c_.strip())]
        if not codes:                                              # a Count row without a material code
            c=w.cell(r_,at,'-'); c._style=_cp(w.cell(r_,at-1)._style); c.number_format='General'; continue
        tx=str(w.cell(r_,text_col).value or ''); sh=None
        if any(c_ in shared for c_ in codes):
            cand=[s_ for c_ in codes for s_ in shared.get(c_,[])]
            m_=re.search(r'\(WSA sheet:\s*([^)]+)\)',tx)
            if m_: sh=next((s_ for s_ in cand if s_==m_.group(1).strip()),None)
            m_=re.search(r'\(([^()]+)\)\s*$',tx)
            if sh is None and m_: sh=next((s_ for s_ in cand if s_.endswith(' '+m_.group(1).strip())),None)
        v_=(fn or it_sup_text)(codes,sheet=sh)
        c=w.cell(r_,at,v_); c._style=_cp(w.cell(r_,at-1)._style); c.number_format='General'
        if not isinstance(v_,str):                                 # a quantity
            c.number_format='#,##0'; c.alignment=Alignment(horizontal='right',vertical='center'); n_+=1; continue
        c.alignment=Alignment(horizontal='left',vertical='center',wrap_text=not oneline); n_+=1
        ln_=1 if oneline else max(1,-(-len(v_)//38))
        if ln_>1: w.row_dimensions[r_].height=max(w.row_dimensions[r_].height or 15,min(13*ln_+4,80))
    if dark:                                                       # dashboards: no white hole on the black page (title / band / No Plan / legend rows; checker finding)
        for r_ in range(1,w.max_row+1):
            c=w.cell(r_,at); o=w.cell(r_,text_col)                 # the item-name cell: page colour, or the band colour on a band row
            if c.value is None and not c.fill.fill_type and o.fill.fill_type: c.fill=_cp(o.fill)
    w.column_dimensions[L(at)].width=width or (40 if fn is None else 13)
    if fz:
        cf_,rf_=re.match(r'([A-Z]+)(\d+)',str(fz)).groups()
        if _CI(cf_)>=at: w.freeze_panes=f"{L(_CI(cf_)+1)}{rf_}"
    if flt:
        a_,b_=flt.split(':'); cb_,rb_=re.match(r'([A-Z]+)(\d+)',b_).groups()
        if _CI(cb_)>=at-1: w.auto_filter.ref=f"{a_}:{L(_CI(cb_)+1)}{rb_}"
    return n_
PGRN_HEAD='Pending GRN Qty'; PGRN=[None]
POD_HEAD='POD (port of discharge)'; STS_HEAD='In transit STATUS'      # PPC In_Transit columns POD / STATUS (user 25-Nov-2026)
PO_HEAD='ETA - PO (PO delivery date)'; PR_HEAD='ETA - PR (requisition delivery date, no PO yet)'
GRR_HEAD='SAP GR since count (PPC GR, reference)'; GRL_HEAD='Last SAP GR (posting date)'
def read_pending_gr(path):
    """PPC sheet 'Pending GR' (user, 18-Nov-2026): shipments at the plant waiting for the SAP goods receipt. One dict per row with a
       PO.No or an Invoice No: po, invoice, supplier, item, inv_qty, row. Header row = the first row with 'PO.No' + 'Invoice No'.
       [] = no such sheet / no rows."""
    try: wb=load_workbook(path,read_only=True,data_only=True)
    except Exception: return []
    sn=next((n_ for n_ in wb.sheetnames if re.sub(r'[^a-z]','',n_.lower()) in ('pendinggr','pendinggrn')),None)
    if not sn: wb.close(); return []
    rows=list(wb[sn].iter_rows(values_only=True)); wb.close()
    nh=lambda v: re.sub(r'[^a-z]','',str(v or '').lower())
    hi=next((i for i,r in enumerate(rows[:20]) if r and {'pono','invoiceno'}<={nh(v) for v in r}),None)
    if hi is None: return []
    hd=[nh(v) for v in rows[hi]]; ix=lambda *ks: next((hd.index(k) for k in ks if k in hd),None)
    iP,iI,iS,iT,iQ=ix('pono'),ix('invoiceno'),ix('suppliername'),ix('item'),ix('invoiceqty')
    g=lambda r,i: r[i] if (i is not None and i<len(r)) else None
    return [dict(po=g(r,iP),invoice=g(r,iI),supplier=' '.join(str(g(r,iS) or '').split()),item=' '.join(str(g(r,iT) or '').split()),
                 inv_qty=g(r,iQ),row=hi+2+k) for k,r in enumerate(rows[hi+1:]) if r and (g(r,iP) not in (None,'') or g(r,iI) not in (None,''))]
def pending_grn():
    """Pending GRN Qty per material (user, 18-Nov-2026): PPC 'Pending GR' rows matched to PPC In_Transit lines on PO + Invoice No +
       supplier (first word); qty = the line's Final Pending Qty in the MRP unit (`qty_mrp`), material = the traced code (`map_in_transit`).
       One invoice can hold 2+ materials (EPE + EVA): every line counts. Rows with no matching pending line -> console NOTE.
       Returns dict(qty={material: qty}, lines=[(sheet, In_Transit line)], miss=[Pending GR row]). Information only - in no calculation."""
    if PGRN[0] is not None: return PGRN[0]
    k=lambda v: re.sub(r'\.0$','',' '.join(str(v if v is not None else '').split())).upper()
    w1=lambda v: (re.findall(r'[A-Z0-9]+',str(v or '').upper()) or [''])[0]
    lines=((CAP[0] or {}).get('ppc') or {}).get('lines') or []
    qty={}; hit=[]; miss=[]; seen=set()
    for p in (read_pending_gr(PPC_SRC[0]) if PPC_SRC and PPC_SRC[0] else []):
        m_=[x for x in lines if p['po'] not in (None,'') and p['invoice'] not in (None,'') and k(x.get('po'))==k(p['po'])
            and k(x.get('invoice'))==k(p['invoice']) and w1(x.get('supplier'))[:4]==w1(p['supplier'])[:4]]
        if not m_: miss.append(p); continue
        for x in m_:
            if id(x) in seen: continue                              # the same invoice twice on Pending GR = counted once
            seen.add(id(x)); c_=x.get('mat') or x.get('code') or x.get('final_code') or '?'
            q_=float(x['qty_mrp'] if x.get('qty_mrp') is not None else x.get('qty') or 0.0)
            qty[c_]=qty.get(c_,0.0)+q_; hit.append((p,x,c_,q_))
    PGRN[0]=dict(qty=qty,lines=hit,miss=miss)
    return PGRN[0]
def pending_grn_qty(codes, sheet=None, own=False):
    """'Pending GRN Qty' of a 4_Status row: the codes of the row added up; a material on 2+ WSA sheets with its sheet =
       only the suppliers of that sheet (same split as `it_sup_text`). '-' = nothing pending GR."""
    d=pending_grn(); al_=BOM_ALIAS[0] or {}; cs=set(codes)|{al_.get(c_,c_) for c_ in codes}
    nm=None
    if sheet:
        g=next((g_ for g_ in (WSA_REFS[0] or []) if sheet in (g_.get('sheet'),g_.get('label'))),None); nm=(g or {}).get('ppcsup_names')
    t=sum(q_ for p,x,c_,q_ in d['lines'] if c_ in cs and (not nm or ' '.join(str(x.get('supplier') or '').split()) in nm))
    return round(t) if t else '-'
def it_pod_status(codes, sheet=None, own=False, kind='pod'):
    """'POD (port of discharge)' / 'In transit STATUS' of a 4_Status row (user 25-Nov-2026: "bom material dashboard and web dashboard
       has info about POD and STATUS column from in transit"): the PPC In_Transit columns POD / STATUS of the lines that are counted as in
       transit for this material (same lines, same sheet / supplier split as 'Pending GRN Qty'). POD = the port (one port: its name; 2+ ports:
       'PORT: dates' per port); STATUS = 'status text: arrival dates' per distinct status (first 4 dates + '+N more'). '-' = nothing in transit /
       empty. Information only - in no calculation."""
    al_=BOM_ALIAS[0] or {}; cs=set(codes)|{al_.get(c_,c_) for c_ in codes}; nm=None
    if sheet:
        g=next((g_ for g_ in (WSA_REFS[0] or []) if sheet in (g_.get('sheet'),g_.get('label'))),None); nm=(g or {}).get('ppcsup_names')
    lines=((CAP[0] or {}).get('ppc') or {}).get('lines') or []
    xs=sorted([x for x in lines if x.get('mat') in cs and x.get('date') is not None and (not nm or ' '.join(str(x.get('supplier') or '').split()) in nm)],key=lambda x: x['date'])
    key='pod' if kind=='pod' else 'status'
    g_={}                                                        # text -> arrival dates (shortest list first read: 'IN-TRANSIT: 29-Nov, 01-Dec, 05-Dec +4 more')
    for x in xs:
        t_=str(x.get(key) or '').strip().replace(' | ','; ')     # the PPC text itself can hold ' | '
        if t_: g_.setdefault(t_,[]).append(f"{x['date']:%d-%b}")
    if not g_: return '-'
    if kind=='pod' and len(g_)==1: return next(iter(g_))        # one port: its name only
    return ' | '.join(f"{t_}: {', '.join(list(dict.fromkeys(d_))[:4])}"+(f" +{len(set(d_))-4} more" if len(set(d_))>4 else '') for t_,d_ in g_.items())
def pod_text(codes, sheet=None, own=False): return it_pod_status(codes,sheet,own,'pod')
def status_text(codes, sheet=None, own=False): return it_pod_status(codes,sheet,own,'status')
def make_daily_material_status(out_path, src_wb, AL_ORDER, bom, cnt_hdr, cnt_rows, d0, mis_path, ships=None):
    """4_Status_<date>.xlsx (user, Nov 2026) - 2 sheets: 'Materials Status' = (1) + (2) in ONE table (user, Nov 2026: "combine
       direct materials and other material then rename to Materials Status") + 'ALERT' (copy of the 1_MRP ALERT, without the 4 class cards).
       Columns F 'Item Category' (PPC Count column I 'Item') and H 'Consumable Unit (Production Confirmed)' (PPC PM_CS_Usage) added.
       1) Direct Material   = the 'Material Status' sheet of 1_MRP (materials of your BOM), copied as it is
       2) Other Materials (was 'Indirect Material', renamed Nov 2026) = materials on the PPC 'Count' sheet that are NOT in the BOM: SAP stock (MB52),
          physical count, in transit; Usage/Day = PPC Count 'Material Req Qty per Day' (user, Nov 2026; rows of one material added up) ->
          days of supply, days incl. in transit, short from (stock + PPC arrivals - Usage/Day, day by day), shortage risk.
       Returns the file path (None when there is nothing to write)."""
    from openpyxl import Workbook
    from openpyxl.styles import Border, Side
    if 'Material Status' not in src_wb.sheetnames: return None
    F=lambda **k: Font(name='Calibri',**k); NAVY='14213D'; INK='1F2D3D'; MUTED='8A949E'; LINE='E3E8EE'; STORE=('2E86C1','D6EAF8'); QUAL=('17A589','D1F2EB')
    thin=Side(style='thin',color=LINE); bd=Border(left=thin,right=thin,top=thin,bottom=thin)
    out_file=os.path.join(os.path.dirname(os.path.abspath(out_path)),outfile_name('dms','xlsx'))
    # ---- which codes are direct (BOM) ----
    direct={str(d.get('Component') or '').strip() for d in (bom or []) if isinstance(d,dict)}
    direct|={str(x['b']['mat']).strip() for x in AL_ORDER if not x.get('sup_row')}
    al_=BOM_ALIAS[0] or {}; also={c for k,v in al_.items() for c in (str(k).strip(),str(v).strip())}      # alias / substitute codes of a BOM material are inside its direct row
    direct|=also
    BOM_DESC.clear()
    for d_ in (bom or []):
        if isinstance(d_,dict) and d_.get('Component') and d_.get('Component desciption'):
            BOM_DESC.setdefault(str(d_['Component']).strip().upper(),' '.join(str(d_['Component desciption']).split()))
    for k_,v_ in al_.items():                                      # alias code (tape 010 -> BOM code 018): the BOM code's description
        if str(v_).strip().upper() in BOM_DESC: BOM_DESC.setdefault(str(k_).strip().upper(),BOM_DESC[str(v_).strip().upper()])
    # ---- the PPC Count sheet: one entry per material code, in the order of the sheet ----
    hl=[' '.join(str(h or '').split()).lower() for h in (cnt_hdr or [])]
    ix=lambda *ks: next((hl.index(k) for k in ks if k in hl),None)
    i_m,i_d,i_u,i_s,i_i=ix('material'),ix('material description'),ix('base unit of measure','stock uom'),ix('storage location'),ix('item','item category')
    i_q=ix('material req qty per day','material req qty / day','req qty per day')          # Usage/Day of a material that is not in the BOM
    g=lambda r,i: (r[i] if (i is not None and i<len(r)) else None)
    ent={}
    for r in cnt_rows or []:
        if not r or not g(r,i_m): continue
        c=' '.join(str(r[i_m]).split()).strip(); e=ent.setdefault(c,dict(desc=None,uom=None,loc=None,item=None,need=None))
        q_=g(r,i_q)
        if isinstance(q_,(int,float)) and not isinstance(q_,bool): e['need']=(e['need'] or 0.0)+float(q_)   # several count rows: added up
        for k,i in (('desc',i_d),('uom',i_u),('loc',i_s),('item',i_i)):
            v=g(r,i)
            if e[k] is None and v not in (None,'') and not (isinstance(v,str) and v.startswith('#')): e[k]=' '.join(str(v).split())
    pc=read_phys_counts(SETTINGS_SRC[0]) if SETTINGS_SRC and SETTINGS_SRC[0] else {}      # counted materials: qty, unit, count date
    sap={}
    try:
        mis=wb_ro(mis_path)
        for r in list(mis['MB52'].iter_rows(values_only=True))[1:]:
            if r and len(r)>14 and r[2] and isinstance(r[14],(int,float)): sap[str(r[2]).strip()]=sap.get(str(r[2]).strip(),0.0)+r[14]
    except Exception as ex_: print(f'NOTE: MB52 not read for the Other Materials sheet ({type(ex_).__name__}: {ex_})')
    ind=[(c,e) for c,e in ent.items() if c not in direct]
    skipped=[c for c in ent if c in also and c not in {str(d.get('Component') or '').strip() for d in (bom or []) if isinstance(d,dict)} and c not in {str(x['b']['mat']).strip() for x in AL_ORDER}]
    # ---- workbook ----
    wb=Workbook(); ws1=wb.active; ws1.title='Direct Material'
    _copy_sheet_to(src_wb['Material Status'],ws1)
    # no STORE / WAREHOUSE / PRODUCTION cards in the daily file (user, Nov 2026: "No need from daily material status"; 1_MRP keeps them)
    cr=next((c.row for c in ws1['A'] if isinstance(c.value,str) and c.value.startswith('🏪')),None)
    if cr: _drop_card_rows(ws1,cr)
    ws=wb.create_sheet('Other Materials'); ws.sheet_view.showGridLines=False; ws.sheet_properties.tabColor='14213D'
    WH=('229954','D5F5E3'); INV=('E67E22','FDEBD0'); PROD=('6C3483','E8DAEF')
    # SAME 19 columns, widths, bands, header colours and cell styles as the Direct Material sheet (Material Status)
    for j,w_ in enumerate([5,16,18,14,17,11,11,11,10,10,11,11,11,11,10,11,9,10,12],1): ws.column_dimensions[L(j)].width=w_
    for t,c1,c2,col in (('',1,5,NAVY),('STORE · stock',6,8,STORE[0]),('WAREHOUSE · supply',9,11,WH[0]),('INVENTORY',12,12,INV[0]),('PRODUCTION',13,13,PROD[0]),('QUALITY',14,15,QUAL[0]),('AGING ANALYSIS',16,19,NAVY)):
        if c2>c1: ws.merge_cells(start_row=1,start_column=c1,end_row=1,end_column=c2)
        c=ws.cell(1,c1,t); c.font=F(size=10,bold=True,color='FFFFFF'); c.alignment=Alignment(horizontal='center',vertical='center')
        for j in range(c1,c2+1): ws.cell(1,j).fill=PatternFill('solid',fgColor=col)
    ws.row_dimensions[1].height=20
    H=['Sr.','Item code','Item name','','Category','Stock','SAP stock (MB52)','Physical count','In transit','Next arrival','Overdue in transit','Reorder level','Usage/Day','Status','On hold','Short from','Days of cover (stock only)','Shortage risk','Days of cover (+ all in transit)']
    hc=[NAVY]*5+[STORE[0]]*3+[WH[0]]*3+[INV[0]]+[PROD[0]]+[QUAL[0]]*2+[NAVY]*4
    for j,(h,col) in enumerate(zip(H,hc),1):
        c=ws.cell(2,j,h); c.font=F(size=9,bold=True,color='FFFFFF'); c.fill=PatternFill('solid',fgColor=col); c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True); c.border=bd
    ws.merge_cells('C2:D2'); ws.row_dimensions[2].height=30
    CH={'Out of Stock':('FADBD8','C0392B'),'No demand':('F2F4F6','8A949E'),'Short':('FDEBD0','B9770E'),'Under lead time':('FCF3CF','9A7D0A'),'OK':('D5F5E3','1E8449')}
    HZ_=180; nneed=0
    hold=mb52_hold(mis_path); fmts=[None,None,None,None,None,'#,##0','#,##0','#,##0','#,##0','dd-mmm','#,##0','#,##0','#,##0.0',None,'#,##0','dd-mmm-yy','0',None,'0']
    r=3; ncount=0; nsap=0
    for i_,(c,e) in enumerate(ind,1):
        p=pc.get(c); cq=p['base'][0] if p else None; sv=sap.get(c)
        if cq is not None: ncount+=1
        if sv is not None: nsap+=1
        stock=cq if cq is not None else sv                                     # a physical count replaces the SAP stock (no demand: nothing to roll forward)
        sh_=(ships or {}).get(c,[]); it=sum(s_['qty'] for s_ in sh_)
        nxt=min((s_['raw'] for s_ in sh_ if s_.get('raw') is not None and s_['raw']>=d0),default=None)
        ovd=sum(s_['qty'] for s_ in sh_ if s_.get('raw') is not None and s_['raw']<d0)
        hq=float(hold.get(c) or 0)
        status='Out of Stock' if (stock is not None and stock<=0) else 'No demand'
        need=e.get('need') if (e.get('need') or 0)>0 else None              # PPC Count 'Material Req Qty per Day'
        if need and cq is not None and p.get('date') is not None:          # counted + a need: rolled forward like Stock Check (checker finding)
            dc_=max(0,(d0-pd.Timestamp(p['date']).normalize()).days)
            if dc_>0: stock=max(0.0,cq-need*dc_+gr_between(c,pd.Timestamp(p['date']).normalize(),d0))
        kz=None; dos=dit=None; risk='No demand'
        if need:
            nneed+=1; st0=float(stock or 0.0); dos=max(0.0,st0)/need; dit=max(0.0,st0+it)/need
            arr=[0.0]*HZ_
            for s_ in sh_:
                if s_.get('raw') is not None:
                    k_=max(0,(s_['raw']-d0).days)
                    if k_<HZ_: arr[k_]+=s_['qty']
            cur=st0
            for k_ in range(HZ_):                                           # stock + arrivals - Usage/Day: first day below zero
                cur+=arr[k_]-need
                if cur<0: kz=k_; break
            risk=('Short' if kz is not None else ('Under lead time' if dit<tgt_days(LT_BY_CODE.get(str(c).upper())) else 'OK'))
        nm=e['desc'] or ''
        vals=[i_,c,nm,None,((item_category(c) if re.search(r'\d',c) else None) or '-') if (item_category(c) if re.search(r'\d',c) else None) not in (None,'','-') else (e['item'] or '-'),(round(stock) if stock is not None else '-'),(round(sv) if sv is not None else '-'),(round(cq) if cq is not None else 'no counting'),
              round(it),(nxt.to_pydatetime() if nxt is not None else 'none'),(round(ovd) if ovd else '-'),'-',(round(need,1) if need else '-'),status,(round(hq) if hq else '-'),
              ((d0+pd.Timedelta(days=kz)).to_pydatetime() if kz is not None else 'none'),(round(dos) if dos is not None else '-'),
              (f"Short {(d0+pd.Timedelta(days=kz)):%d-%b}" if kz is not None else risk),(round(dit) if dit is not None else '-')]
        for j,(v,f_) in enumerate(zip(vals,fmts),1):
            cc=ws.cell(r,j,v); cc.font=F(size=10,color=INK); cc.border=Border(bottom=thin)
            cc.alignment=Alignment(horizontal=('left' if j in (2,3,5) else 'center'),vertical='center')
            if f_: cc.number_format=f_
            if j in (6,7,8): cc.fill=PatternFill('solid',fgColor='F4F9FD')
            if j in (9,10,11): cc.fill=PatternFill('solid',fgColor='F4FBF7')
            if j==12: cc.fill=PatternFill('solid',fgColor='FEF8F1')
            if j==13: cc.fill=PatternFill('solid',fgColor='F8F4FB')
        ws.merge_cells(start_row=r,start_column=3,end_row=r,end_column=4)
        if kz is not None: ws.cell(r,16).font=F(size=10,bold=True,color='C0392B')
        for j,key in ((14,status),(18,risk)):
            bg,fg=CH[key]; ws.cell(r,j).fill=PatternFill('solid',fgColor=bg); ws.cell(r,j).font=F(size=10,bold=True,color=fg)
        if hq: ws.cell(r,15).font=F(size=10,bold=True,color='B9770E')
        ws.row_dimensions[r].height=19; r+=1
    _del_cols(ws,[12,14,15])                                      # same columns as Direct Material (no INVENTORY / QUALITY, Nov 2026)
    _add_uom_col(ws,[e.get('uom') for _,e in ind])                # UOM after Category = the PPC Count Base Unit of Measure (user, Nov 2026)
    ws.cell(r+1,1,f'Other materials = on the PPC Count sheet but not in the BOM ({len(ind)} materials; {ncount} counted, {nsap} with MB52 stock, {nneed} with a Usage/Day) · same columns as the BOM materials above · Stock = physical count if counted (with a Usage/Day: rolled forward to today = count − need × days since the count + received since the count, like Stock Check), else SAP (MB52) stock · '
                  'SAP stock (MB52) = MB52 stock · Physical count = PPC physical count (no counting = not counted) · In transit / Overdue = MIS pending shipments · '
                  f'Usage/Day = PPC Count column Material Req Qty per Day (rows of one material added up) · Inv. Days = stock ÷ need · Days of cover (+ all in transit) = (stock + in transit) ÷ need · '
                  f'Short from = first day stock + PPC arrivals − Usage/Day < 0 (next {HZ_} days) · Shortage risk: Short dd-mmm / Under lead time (incl. in transit) / OK / No demand (no Usage/Day on the Count sheet).').font=F(size=8,italic=True,color=MUTED)
    # ---- Item Category (PPC Count column I 'Item') after Category + Consumable Unit (PPC PM_CS_Usage) after UOM (user, Nov 2026) ----
    from copy import copy as _cp
    cons={}
    try:
        _pw=wb_ro(PPC_SRC[0]) if PPC_SRC and PPC_SRC[0] else None
        _sn=next((n_ for n_ in ('PM_CS_Usage','CS_Usage') if _pw is not None and n_ in _pw.sheetnames),'PM_CS_Usage')   # PPC renamed the sheet to CS_Usage (Nov 2026)
        _r=list(_pw[_sn].iter_rows(values_only=True)) if _pw is not None else []
        _h=next((i for i,x in enumerate(_r[:10]) if x and 'code' in [' '.join(str(v or '').split()).lower() for v in x]),None)
        if _h is not None:
            _hl=[' '.join(str(v or '').split()).lower() for v in _r[_h]]
            _ic=_hl.index('code'); _iu=next((j for j,h in enumerate(_hl) if h.startswith('consumable unit')),None)
            for x in _r[_h+1:]:
                if not x or _ic>=len(x) or not x[_ic] or _iu is None or _iu>=len(x): continue
                v_=x[_iu]
                if v_ not in (None,'','-',0) and not (isinstance(v_,str) and v_.startswith('#')): cons.setdefault(' '.join(str(x[_ic]).split()),' '.join(str(v_).split()))
    except Exception as ex_: print(f'NOTE: PPC PM_CS_Usage (Consumable Unit) not read ({type(ex_).__name__}: {ex_})')
    item_of={c_:e_.get('item') for c_,e_ in ent.items()}
    _io=item_of.get; item_of=type('D',(dict,),{'get':lambda self,k,d=None:(('Consumables' if is_consumable(k) else None) or _io(k) or count_item_of(k) or d)})(item_of)   # a code not on Count: its alias code's Item (tape 010 <- 018)
    def _data_rows(w):
        out=[]; r_=3
        while isinstance(w.cell(r_,1).value,int) and w.cell(r_,2).value: out.append(r_); r_+=1
        return out
    def _add_col(w, at, head, vals_by_row, band_to):
        _ins_cols(w,at,1)
        for m_ in list(w.merged_cells.ranges):
            if m_.min_row==1 and m_.max_row==1 and m_.min_col==1 and m_.max_col<band_to: w.unmerge_cells(str(m_)); w.merge_cells(start_row=1,start_column=1,end_row=1,end_column=band_to)
        w.cell(1,at)._style=_cp(w.cell(1,1)._style)
        h=w.cell(2,at,head); h._style=_cp(w.cell(2,at-1)._style)
        for r_,v_ in vals_by_row.items():
            c=w.cell(r_,at,v_ if v_ not in (None,'') else '-'); c._style=_cp(w.cell(r_,at-1)._style); c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
        w.column_dimensions[L(at)].width=14
    for w in (ws1,ws):
        drs=_data_rows(w)
        _add_col(w,6,'Item Category',{r_:item_of.get(' '.join(str(w.cell(r_,2).value).split())) for r_ in drs},6)
        _add_col(w,8,'Consumable Unit (Production Confirmed)',{r_:cons.get(' '.join(str(w.cell(r_,2).value).split())) for r_ in drs},8)
    # ---- ONE sheet 'Materials Status' = BOM materials, then the other Count materials (user, Nov 2026) ----
    ms=wb.create_sheet('Materials Status',0); ms.sheet_view.showGridLines=False; ms.sheet_properties.tabColor='14213D'
    def _copy_rows(src,r1,r2,off):
        for rr in range(r1,r2+1):
            for c in src[rr]:
                if c.value is None and not c.has_style: continue
                n_=ms.cell(rr+off,c.column); n_._value=c._value; n_.data_type=c.data_type; n_._style=_cp(c._style)
            if src.row_dimensions[rr].height: ms.row_dimensions[rr+off].height=src.row_dimensions[rr].height
        for m_ in src.merged_cells.ranges:
            if r1<=m_.min_row and m_.max_row<=r2: ms.merge_cells(start_row=m_.min_row+off,start_column=m_.min_col,end_row=m_.max_row+off,end_column=m_.max_col)
    d1=_data_rows(ws1); d2=_data_rows(ws); ncol=ws1.max_column
    lastD=d1[-1] if d1 else 2
    _copy_rows(ws1,1,lastD,0)
    sep=lastD+1                                                    # band row: where the other materials start
    ms.merge_cells(start_row=sep,start_column=1,end_row=sep,end_column=ncol)
    sc=ms.cell(sep,1,f'OTHER MATERIALS · on the PPC Count sheet, not in the BOM ({len(d2)}) · Usage/Day = PPC Count Material Req Qty per Day')
    sc.font=F(size=10,bold=True,color='FFFFFF'); sc.fill=PatternFill('solid',fgColor=NAVY); sc.alignment=Alignment(horizontal='left',vertical='center',indent=1)
    ms.row_dimensions[sep].height=20
    if d2:                                                         # highest risk first (user, Nov 2026: "sort from highest risk to lowest"):
        hd_=[ws.cell(2,j).value for j in range(1,ws.max_column+1)]  # Short dd-mmm (earliest first), Under lead time, OK, No demand
        jr_=hd_.index('Shortage risk')+1 if 'Shortage risk' in hd_ else None; jf_=hd_.index('Short from')+1 if 'Short from' in hd_ else None
        ji_=hd_.index('Days of cover (+ all in transit)')+1 if 'Days of cover (+ all in transit)' in hd_ else None
        def _ok(r_):
            rk=str(ws.cell(r_,jr_).value or '') if jr_ else ''; sf=ws.cell(r_,jf_).value if jf_ else None; di=ws.cell(r_,ji_).value if ji_ else None
            g_=0 if rk.startswith('Short') else (1 if rk.startswith('Under') else (2 if rk=='OK' else 3))
            return (g_,sf if isinstance(sf,dt.datetime) else dt.datetime(2999,1,1),di if isinstance(di,(int,float)) else 1e12,r_)
        for k_,r_ in enumerate(sorted(d2,key=_ok)):
            _copy_rows(ws,r_,r_,sep+1+k_-r_)
        for k_,r_ in enumerate(range(sep+1,sep+1+len(d2)),1): ms.cell(r_,1).value=len(d1)+k_      # Sr. continues
    endO=sep+len(d2)
    _copy_rows(ws1,lastD+1,ws1.max_row,endO-lastD)               # summary boxes + notes of the BOM materials
    _copy_rows(ws,(d2[-1] if d2 else 2)+1,ws.max_row,ms.max_row+1-((d2[-1] if d2 else 2)+1))   # other materials note
    for k_,d_ in ws1.column_dimensions.items():
        if d_.width: ms.column_dimensions[k_].width=d_.width
    ms.freeze_panes='C3'
    for row_ in ms.iter_rows(min_row=endO+1):                     # the summary boxes count the BOM materials only (checker finding)
        for c in row_:
            if isinstance(c.value,str) and c.value.startswith(('⚠️  KEY RISK','⏱  SHORTAGE TIMELINE','⚠️  SHORTAGE WATCH')): c.value=c.value+' · BOM materials'
    del wb['Direct Material']; del wb['Other Materials']
    # ---- ALERT copy (user, Nov 2026: "Copy Alert this workbook but remove STOCK-OUT IN TRANSIT LEAD TIME / STOCK-OUT LATER / UNDER 21 DAYS / UNDER 21 DAYS LATER" = the 4 class cards) ----
    if 'ALERT' in src_wb.sheetnames:
        wa=wb.create_sheet('ALERT',1); sa=src_wb['ALERT']
        _copy_sheet_to(sa,wa); wa.freeze_panes=sa.freeze_panes; wa.sheet_view.zoomScale=sa.sheet_view.zoomScale
        if sa.auto_filter.ref: wa.auto_filter.ref=sa.auto_filter.ref
        wa.print_title_rows=sa.print_title_rows; wa.page_setup.orientation=sa.page_setup.orientation       # print like the 1_MRP ALERT (checker finding)
        wa.page_setup.fitToWidth=sa.page_setup.fitToWidth; wa.page_setup.fitToHeight=sa.page_setup.fitToHeight; wa.sheet_properties.pageSetUpPr.fitToPage=True
        cr=next((c.row for c in wa['A'] if c.row>ALERT_HR and c.value==ALERT_CLS[0][0]),None)
        if cr: _drop_card_rows(wa,cr)
    # ---- 'Usage per module by customer' (= Stock Check / old Action List column) as the last column of both sheets (user, Nov 2026) ----
    UH='Usage per module by customer (source: your BOM / MIS Standard BOM)'
    use_of={}
    for x_ in AL_ORDER:
        if x_.get('sup_row'): continue
        v_=x_.get('usage_src'); v_=', '.join(map(str,v_)) if isinstance(v_,(list,tuple)) else v_
        if v_: use_of.setdefault(str(x_['b']['mat']).strip(),v_)
    def _usage_col(w, hr, rows_, code_col, band_r=None, last_r=None):
        nc=w.max_column if last_r is None else last_r
        j=nc+1
        h=w.cell(hr,j,UH); h._style=_cp(w.cell(hr,nc)._style); h.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
        if band_r:
            b_=w.cell(band_r,j,'BOM · usage'); b_._style=_cp(w.cell(band_r,1)._style); b_.alignment=Alignment(horizontal='center',vertical='center')
        for r_ in rows_:
            cd_=' '.join(str(w.cell(r_,code_col).value or '').split())
            v_=use_of.get(cd_) or '-'
            c=w.cell(r_,j,v_); c._style=_cp(w.cell(r_,nc)._style); c.alignment=Alignment(horizontal='left',vertical='center',wrap_text=True)
            n_=sum(max(1,-(-len(t_)//44)) for t_ in str(v_).split('\n'))      # wrapped lines too (width 46 ~ 44 characters; checker finding)
            w.row_dimensions[r_].height=max(w.row_dimensions[r_].height or 15,min(15*n_,150))
        w.column_dimensions[L(j)].width=46
        return j
    jm=_usage_col(ms,2,list(range(3,lastD+1))+list(range(sep+1,endO+1)),2,band_r=1,last_r=ncol)
    for m_ in list(ms.merged_cells.ranges):                       # band row of the other materials spans the new column too
        if m_.min_row==sep and m_.max_row==sep: ms.unmerge_cells(str(m_)); ms.merge_cells(start_row=sep,start_column=1,end_row=sep,end_column=jm)
    if 'ALERT' in wb.sheetnames:
        wa=wb['ALERT']; ar_=[r_ for r_ in range(ALERT_HR+1,wa.max_row+1) if isinstance(wa.cell(r_,1).value,int) and wa.cell(r_,3).value]
        lastA=max([c.column for c in wa[ALERT_HR] if c.value is not None]+[1])
        ja=_usage_col(wa,ALERT_HR,ar_,3,band_r=ALERT_HR-1,last_r=lastA)
        if wa.auto_filter.ref: wa.auto_filter.ref=f"A{ALERT_HR}:{L(ja)}{ar_[-1] if ar_ else ALERT_HR+1}"
    # ---- Cover Heatmap (1_MRP, grouped RM / PM / CS) (user, Nov 2026) ----
    for n_ in (HEAT_NAME,):
        if n_ in src_wb.sheetnames:
            wc=wb.create_sheet(n_); sc_=src_wb[n_]; _copy_sheet_to(sc_,wc); wc.freeze_panes=sc_.freeze_panes
            for cf_ in sc_.conditional_formatting:                  # the daily sheet's colour rules
                for rl_ in cf_.rules: wc.conditional_formatting.add(str(cf_.sqref),rl_)
    if 'RM Issues' in src_wb.sheetnames:                           # user, 19-Nov-2026: "move sheet RM Issues from daily MRP to daily materials status" (table, Status list, frozen header kept)
        from openpyxl.worksheet.table import Table as _Tb
        wi=wb.create_sheet('RM Issues'); si=src_wb['RM Issues']; _copy_sheet_to(si,wi); wi.freeze_panes=si.freeze_panes
        for cf_ in si.conditional_formatting:
            for rl_ in cf_.rules: wi.conditional_formatting.add(str(cf_.sqref),rl_)
        for t_ in si.tables.values(): wi.add_table(_Tb(displayName=t_.displayName,ref=t_.ref,tableStyleInfo=t_.tableStyleInfo))
        for dv_ in si.data_validations.dataValidation: wi.data_validations.append(dv_)
    # In transit supplier name (qty) on EVERY sheet (user, Nov 2026: "Add in transit supplier name in all sheet of daily material status"):
    # right after 'In transit' (Materials Status, ALERT, BOM Weekly), 
    def _col_of(w,hr,t_): return next((c.column for c in w[hr] if isinstance(c.value,str) and c.value.strip().startswith(t_)),None)
    def _code_rows(w,r0,cc):
        return [r_ for r_ in range(r0,w.max_row+1) if any(IT_CODE.match(x_.strip()) for x_ in re.split(r'\s*(?:/|,|\+|\n)\s*',str(w.cell(r_,cc).value or '')))]
    nsup={}
    try:
        if 'Materials Status' in wb.sheetnames:
            w=wb['Materials Status']; j=_col_of(w,2,'In transit')
            if j: nsup['Materials Status']=_add_it_sup_col(w,2,j+1,list(range(3,lastD+1))+list(range(sep+1,endO+1)),2,3)
        if 'ALERT' in wb.sheetnames:
            w=wb['ALERT']; j=_col_of(w,ALERT_HR,'In transit')
            if j: nsup['ALERT']=_add_it_sup_col(w,ALERT_HR,j+1,[r_ for r_ in range(ALERT_HR+1,w.max_row+1) if isinstance(w.cell(r_,1).value,int) and w.cell(r_,3).value],3,4)
        if HEAT_NAME in wb.sheetnames:
            w=wb[HEAT_NAME]; j=_col_of(w,4,'In transit')
            if j: nsup[HEAT_NAME]=_add_it_sup_col(w,4,j+1,_code_rows(w,5,1),1,2,dark=True)
    except Exception as ex_:
        import traceback; traceback.print_exc(); print(f'NOTE: In transit supplier column not added everywhere ({type(ex_).__name__}: {ex_})')
    # Pending GRN Qty on EVERY sheet (user, 18-Nov-2026: PPC 'Pending GR' PO + Invoice + Supplier -> In_Transit Final Pending Qty):
    # BOM Weekly after 'Next arrival' (below, once its columns are rebuilt). Information only.
    npg={}; npo=0; neta=0
    # user, Nov 2026: "Materials Status, delete row 1, then hide this sheet; ALERT, delete row 10 then move row 1 to 9 to bottom"
    if 'Materials Status' in wb.sheetnames:                        # the group band row goes: header = row 1, data from row 2
        _remap_rows(wb['Materials Status'],lambda r_: None if r_==1 else r_-1)
    if 'ALERT' in wb.sheetnames:                                   # group band (row ALERT_HR-1) goes; summary boxes (rows 1 .. ALERT_HR-2) below the table
        wa=wb['ALERT']; top=ALERT_HR-2; end_=wa.max_row-(top+1)    # rows after the band move up by top + 1
        _remap_rows(wa,lambda r_: (end_+1+r_ if r_<=top else (None if r_==top+1 else r_-(top+1))))
    for n_ in (HEAT_NAME,):          # user, Nov 2026: "Dont freeze or wrap text ... all dashboards look organized"
        if n_ in wb.sheetnames: _tidy_dashboard(wb[n_])
    if HEAT_NAME in wb.sheetnames: _heat_final(wb[HEAT_NAME])             # rows 1-2 out, 9 pt, Material Status columns grouped (user, Nov 2026)
    watch_=None
    try:                                                           # user, 12-Nov-2026: the look of their WSA Material Coverage page (light, 3 colours, Shortage Watch first)
        for n_ in (HEAT_NAME,):
            if n_ in wb.sheetnames: _coverage_light(wb[n_])
        watch_=_shortage_watch(wb)
        org_=_organize_bom_weekly(wb)
        if HEAT_NAME in wb.sheetnames:
            w=wb[HEAT_NAME]; j=_col_of(w,2,'Arrival at plant (Final AtA)')
            if j: npg[HEAT_NAME]=_add_it_sup_col(w,2,j+1,_code_rows(w,3,1),1,2,head=PGRN_HEAD,fn=pending_grn_qty)
            j=_col_of(w,2,'Arrival at plant (Final AtA)')           # user 24-Nov-2026: ETA (PPC PR_PO delivery date) right next to AtA
            if j and PRPO_LINES:                                    # user 24-Nov-2026 'Segregate PR and PO ETA': two columns, PO next to AtA, PR after it
                for h_,k_ in ((PR_HEAD,'PR'),(PO_HEAD,'PO')):       # each is inserted right after AtA, so the last one (PO) ends up next to it
                    npo=_add_it_sup_col(w,2,j+1,_code_rows(w,3,1),1,2,head=h_,fn=lambda c,sheet=None,own=False,k=k_: pr_po_text(c,sheet,own,kind=k),oneline=True,width=34)
            neta=_eta_cols_days(w,2,_code_rows(w,3,1),1)             # user 25-Nov-2026: PO / PR quantity -> days of cover
            j=_col_of(w,2,PGRN_HEAD)                                # user 25-Nov-2026: POD + STATUS of the PPC In_Transit lines, right after Pending GRN Qty
            if j:
                for h_,f_ in ((STS_HEAD,status_text),(POD_HEAD,pod_text)):
                    _add_it_sup_col(w,2,j+1,_code_rows(w,3,1),1,2,head=h_,fn=f_,oneline=True,width=14)
            if GR_REF:                                          # user 20-Nov-2026: PPC GR for reference only (no calculation)
                for h_,f_,a_ in ((GRR_HEAD,gr_ref_since,STS_HEAD),(GRL_HEAD,gr_ref_last,GRR_HEAD)):
                    j=_col_of(w,2,a_)
                    if j: _add_it_sup_col(w,2,j+1,_code_rows(w,3,1),1,2,head=h_,fn=f_)
            jc_=_col_of(w,2,'Code'); ck_={r_:str(w.cell(r_,jc_).value or '').strip() for r_ in _code_rows(w,3,jc_)} if jc_ else {}
            _bom_weekly_keep(wb[HEAT_NAME])
            _dms_info_capture(wb[HEAT_NAME],ck_)                # the web page shows every column of this sheet (user 25-Nov-2026)
            _dms_two_rows(wb[HEAT_NAME],ck_,d0)                 # LAST step: two rows per material (user 27-Nov-2026, 5_Web Summary layout); rows shift, so after the capture above
    except Exception as ex_:
        import traceback; traceback.print_exc(); print(f'NOTE: coverage-page look not applied to every dashboard ({type(ex_).__name__}: {ex_})')
    for n_ in (WATCH_NAME,'Materials Status','ALERT'):               # user, 15-Nov-2026: "delete shortage watch, materials status, alert sheet from output daily material status" (built above only to feed the dashboards)
        if n_ in wb.sheetnames: del wb[n_]
    if 'RM Issues' in wb.sheetnames: wb.move_sheet('RM Issues',offset=(wb.sheetnames.index(HEAT_NAME)+1 if HEAT_NAME in wb.sheetnames else 0)-wb.sheetnames.index('RM Issues'))
    for w_ in wb.worksheets:                                       # user, 18-Nov-2026: "hide all sheets except BOM Material Dashboard - Weekly" (+ RM Issues, user 19-Nov-2026)
        if w_.title not in (HEAT_NAME,'RM Issues'): w_.sheet_state='hidden'
    vis=[i_ for i_,w_ in enumerate(wb.worksheets) if w_.sheet_state=='visible']
    wb.active=vis[0] if vis else 0                                 # the first VISIBLE sheet opens (a hidden active sheet = Excel repair)
    for i_,w_ in enumerate(wb.worksheets): w_.sheet_view.tabSelected=(i_==wb.index(wb.active))
    if wb.views: wb.views[0].firstSheet=wb.index(wb.active)   # first tab shown = a visible sheet (checker note)
    _safe_save(wb,out_file)
    print(f"Daily Material Status saved: {out_file} - {HEAT_NAME} + RM Issues (the only visible sheets, {len(DMS_WEEKLY_KEEP)} info columns + weeks; the other sheets of the file are not built any more) (Shortage Watch / Materials Status / ALERT are not in this file, user 15-Nov-2026; {len(d1)} BOM materials); other materials ({len(ind)} Count materials not in the BOM: {ncount} counted, {nsap} with MB52 stock, {nneed} with a Usage/Day)"
          +(f"; left out of Other Materials because they are alias / substitute codes of a BOM material: {', '.join(skipped)}" if skipped else ''),flush=True)
    if watch_: print(f"  Coverage: {watch_[0]} materials - {watch_[1]} short today, {watch_[2]} more short within their transit lead time",flush=True)
    if nsup: print(f"  '{IT_SUP_HEAD}' added: "+', '.join(f'{k} {v} rows' for k,v in nsup.items()),flush=True)
    pg_=pending_grn()
    if npg: print(f"  '{PGRN_HEAD}' (PPC Pending GR -> In_Transit on PO + Invoice + supplier): {len(pg_['lines'])} In_Transit lines, "
                  +(', '.join(f'{c_} {q_:,.0f}' for c_,q_ in sorted(pg_['qty'].items())) or 'none')+'; added: '+', '.join(f'{k} {v} rows' for k,v in npg.items()),flush=True)
    if npo: print(f"  '{PO_HEAD}' + '{PR_HEAD}' (PPC sheet PR_PO: open PR / PO lines, information only - in no calculation): {len(PRPO_LINES)} materials with open lines, added: {npo} rows",flush=True)
    print(f"  '{PO_HEAD}' / '{PR_HEAD}': {neta} cells show days of cover (qty / Usage/Day) instead of quantity",flush=True)
    if pg_['miss']: print(f"NOTE: {len(pg_['miss'])} PPC Pending GR rows have no pending In_Transit line with the same PO + Invoice No + supplier (not in {PGRN_HEAD}): "
                          +'; '.join(f"row {p['row']} {p['item'] or '?'} {p['supplier']} PO {p['po'] if p['po'] is not None else '-'} inv {p['invoice'] if p['invoice'] is not None else '-'}" for p in pg_['miss']),flush=True)
    return out_file

