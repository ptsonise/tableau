# ===================== Settings sheet (simple config in Excel) =====================
BASIC_ROWS=[('Lead time raw materials (days) - if no ETD history','LEAD_TIME_RM','Fallback transit lead time; every day threshold = the lead time'),
            ('Lead time packaging / consumables (days)','LEAD_TIME_CSPM','Fallback transit lead time (packaging / consumables)'),
            ('Dashboard: red when Inv. Days are under (days)','COV_RED_DAYS','Dashboards: red = cover under this many days (colour step 1 = 1 .. this), amber = under the lead time, green = covered'),
            ('Dashboard: colour step 2 ends at (days)','DASH_STEP2','Dashboards (1_MRP dark colours): cover from the red limit up to this many days'),
            ('Dashboard: first blue = lead time up to lead time + (days)','DASH_BLUE1','Dashboards (1_MRP dark colours)'),
            ('Dashboard: second blue ends at lead time + (days)','DASH_BLUE2','Dashboards (1_MRP dark colours); beyond = light blue'),
            ('Dashboard: a lead time under this counts as (days)','DASH_MIN_LT','Dashboards: minimum lead time used for the colour steps'),
            ('Risk order: second step = first day cover is below (days)','RISK_DAYS','Every dashboard: first day cover is below 1, then below this many days, then lowest cover today'),
            ('ALERT: look-ahead for "below lead time later" (days)','ALERT_HORIZON','112 = 16 weeks: ALERT Below lead time from / Why (cause) / RM Issues'),
            ('ETD alert: "ships within" (days)','ETD_SOON','ETD this close and still not on PPC In_Transit = SHIPS WITHIN N DAYS'),
            ('Arrival at plant: next arrival dates shown','ATA_SHOW','ALERT / Material Status show this many dates, the rest as one total')]
ADV_ROWS=[('ABC: A = top % of value','ABC_A_PCT','80'),('ABC: B = up to % of value','ABC_B_PCT','95 (rest = C)'),
          ('ABC: days of requirement used for the value','ABC_HORIZON','90'),('XYZ: X if weekly variation below','XYZ_X','0.25'),
          ('XYZ: Y if weekly variation below (above = Z)','XYZ_Y','0.5'),('Shortage measured over the next (days)','SHORT_HORIZON','60'),
          ('RELEASE ORDER when an order is due within (days)','RELEASE_WINDOW','7'),('RESCHEDULE OUT when stock is more than x times the target','RESCHED_OUT_X','3'),
          ('Count only WSA lines (Yes / No)','PLAN_WSA_LINES_ONLY','No'),
          ('Inbound: short lead if ETD to plant date is under (days)','SHORT_LEAD','20'),('Inbound: long lead if ETD to plant date is over (days)','LONG_LEAD','75')]
SETTING_ROWS=BASIC_ROWS+ADV_ROWS
SUBS_X=[]          # substitutes of this run, one per substitute code ('A or B' split), resolved to material codes
SUBS=[]            # substitutes from the Settings sheet: dict(mrp, sub, factor, lines(set or None), fr, to, order) - text as typed
def _date(x):
    try: return pd.Timestamp(x).normalize() if x not in (None,'') else None
    except Exception as _e_: _swallow(5868,_e_); return None
def read_settings(paths):
    """Settings sheet: from your MRP.xlsx if it has one, else from the previous 1_MRP. Sets the rules for this run.
       Sections: basic numbers, unlimited materials, allowed vendors (by customer), substitutes, line ranking, advanced numbers.
       Old sections (Customer Supplier, Pack conversions, Vendor Priority) are still read and turned into the new ones."""
    g=globals(); src=None
    for pth in paths:
        if not pth or not os.path.exists(pth): continue
        try: w=wb_ro(pth)
        except Exception as _e_: _swallow(5877,_e_); continue
        sn_=next((n_ for n_ in ('Settings',sheet_new('mrp','Settings')) if n_ in w.sheetnames),None)      # a Settings sheet copied from 1_MRP keeps its number
        if not sn_: continue
        rows=[list(r)+[None]*8 for r in w[sn_].iter_rows(values_only=True)]
        key={a.lower():b for a,b,_ in SETTING_ROWS}
        unl=set(); words=[]; rank={}; rmat={}; lhdr=None; cs_=[]; subs=[]; sec=None; vhdr=None; lo={}; ltg={}; bsp={}; ct_=[]; lto={}
        for r in rows:
            a=str(r[0] or '').strip(); al=a.lower()
            if al.startswith('line ranking'): lhdr=None
            for t,sc in (('mis link','k'),('unlimited materials','u'),('allowed vendors','c'),('customer supplier','c'),('substitutes','s'),('pack conversions','p'),
                         ('line ranking','l'),('vendor priority','v'),('batch supplier codes','b'),('customer daily target','t'),('transit lead time','lt'),('advanced','x'),('basic','x')):
                if al.startswith(t): sec=sc; vhdr=None; break
            else:
                if 'daily target' in al and isinstance(r[1],(int,float)) and not isinstance(r[1],bool):   # 'USA Line 1 Production Plan Daily Target'
                    ltg[_line_key(re.split(r'production plan|daily target',a,flags=re.I)[0])]=(a,float(r[1])); continue
                if al in key and r[1] not in (None,''):
                    nm=key[al]
                    if nm=='PLAN_WSA_LINES_ONLY': g[nm]=str(r[1]).strip().lower() in ('yes','y','true','1')
                    else:
                        try: g[nm]=(float(r[1]) if nm in ('XYZ_X','XYZ_Y') else int(float(r[1])))
                        except (TypeError,ValueError): pass
                    continue
                if not a and not any(r[1:6]): continue
                if sec in ('u','c') and not any(r[1:6]) and len(a)>40 and ' ' in a: continue     # the section's help sentence, not a rule (user 22-Nov-2026)
                if sec=='t':
                    if a and not al.startswith('customer') and isinstance(r[2] if len(r)>2 else None,(int,float)):
                        ct_.append((a,str(r[1] or '').strip(),float(r[2]),'Yes' if str(r[3] if len(r)>3 else '').strip().lower() in ('yes','y','oem','x') else 'No'))
                    continue
                if sec=='lt':                 # TRANSIT LEAD TIME: code | material | type | WSA history | deliveries | YOUR lead time | used | source
                    v_=r[5] if len(r)>5 else None
                    if a and not al.startswith('code') and isinstance(v_,(int,float)) and not isinstance(v_,bool) and v_>0: lto[a.upper()]=int(round(float(v_)))
                    continue
                if sec=='b':
                    if a and not al.startswith('batch code') and r[1] not in (None,''): bsp[a.upper()[:2]]=str(r[1]).strip()
                    continue
                if sec=='u' and a and not al.startswith(('code','material code')):
                    (unl.add(a) if re.match(r'^[A-Z]{2}[A-Z0-9]{6,}$',a.upper()) else words.append(al))
                elif sec=='c' and al!='customer' and a:
                    cs_.append({'Customer':a,'Material':r[1],'Code':r[2],'Supplier':r[3],'Share %':None})
                elif sec=='s' and not al.startswith('mrp material') and (a and r[1]):
                    lines=None if r[3] in (None,'') else {_line_key(x) for x in re.split(r'[,;/]',str(r[3])) if x.strip()}
                    try: f=float(r[2]) if r[2] not in (None,'') else 1.0
                    except (TypeError,ValueError): f=1.0
                    subs.append(dict(mrp=a,sub=str(r[1]),factor=f,lines=lines,fr=_date(r[4]),to=_date(r[5]),order=r[6] if isinstance(r[6],(int,float)) else 1))
                elif sec=='p' and a and not al.startswith('code') and r[1] and r[2] not in (None,''):     # old: pack conversions
                    try: subs.append(dict(mrp=str(r[1]),sub=a,factor=float(r[2]),lines=None,fr=_date(r[3]),to=_date(r[4]),order=1))
                    except (TypeError,ValueError): pass
                elif sec=='k' and al.startswith('production plan customer'):
                    hk_=[' '.join(str(x or '').split()).lower() for x in r]; PROJ_COL[0]=hk_.index('project name') if 'project name' in hk_ else None
                elif sec=='k' and a and PROJ_COL[0] is not None and PROJ_COL[0]<len(r) and r[PROJ_COL[0]] and (PROJECT_OF.__setitem__(_norm(a),' '.join(str(r[PROJ_COL[0]]).split())) or r[4] in (None,'')): pass
                elif sec=='k' and a and not al.startswith('production plan customer') and r[4] not in (None,''):
                    lo[_norm(a)]=(str(r[5]).strip() if r[5] not in (None,'') else None,str(r[4]).strip()); LINK_OVR_NAME[_norm(a)]=a
                elif sec=='l':
                    if al.startswith('material'): lhdr=r; continue
                    if lhdr is not None:                         # matrix: one row per material, one column per line
                        rk={}
                        for j in range(3,len(lhdr)):
                            if lhdr[j] and isinstance(r[j],(int,float)): rk[_line_key(lhdr[j])]=float(r[j])
                        if rk:
                            if al.startswith('all materials'): rank.update(rk)
                            else: rmat[str(r[2] or a or r[1]).strip()]=dict(rk,_name=a,_desc=r[1])
                    elif a and not al.startswith('production line') and r[1] not in (None,''):   # old format: line | rank
                        try: rank[_line_key(a)]=float(r[1])
                        except (TypeError,ValueError): pass
                elif sec=='v':                                                                         # old: vendor priority matrix
                    if al=='material': vhdr=r; continue
                    if vhdr is not None and (a or r[2]) and not r[1]:      # rows without vendor -> substitutes per line
                        for j in range(3,len(vhdr)):
                            if vhdr[j] and isinstance(r[j],(int,float)) and r[j]>0:
                                subs.append(dict(mrp=a,sub=str(r[2] or a),factor=1.0,lines={_line_key(vhdr[j])},fr=None,to=None,order=r[j]))
        if unl or words: g['UNLIMITED_CODES']=unl; g['UNLIMITED_WORDS']=tuple(words)     # empty Settings section = keep the script list (523 / 338, user rule)
        g['LINE_RANK']=rank; g['LINE_RANK_MAT']=rmat
        LINE_TARGET.clear(); LINE_TARGET.update(ltg); BATCH_SUP.clear(); BATCH_SUP.update(bsp); LT_OVERRIDE.clear(); LT_OVERRIDE.update(lto)
        # no CUSTOMER DAILY TARGET section: each customer takes its LINE's daily target from Settings (e.g. USA Line 1 = 4,000)
        CUST_TARGET[:]=ct_ or [(c_,l_,((ltg.get(_line_key(l_)) or (None,None))[1] or t_),o_) for c_,l_,t_,o_ in CUST_TARGET_DEFAULT]
        g['SET_CUSTSUP']=[x for x in cs_ if x.get('Customer')]; g['SUBS']=subs; LINK_OVERRIDE.clear(); LINK_OVERRIDE.update(lo); src=pth; break
    return src

def resolve_substitutes(mis_path):
    """Turns the Substitutes table (code OR description, as you typed it) into material codes, using the MIS descriptions.
       Substitutes for ALL lines become pack conversions (counted in the material's stock, with dates); substitutes for
       some lines only are used when sharing stock out to those lines. Returns Check rows for anything not found."""
    g=globals(); iss=[]; mis=wb_ro(mis_path); desc_of={}; dash={}; names={}
    for r in mis['Daily Dashboard'].iter_rows(min_row=5,values_only=True):
        if len(r)>6 and str(r[6] or '').strip()=='Plan' and r[4]: dash[str(r[4]).strip()]=f"{r[3]} {r[5] or ''}"; names[str(r[4]).strip()]=str(r[3] or '').strip()
    for r in mis['Common Name'].iter_rows(min_row=2,values_only=True):
        if r and r[0] and r[1]: desc_of.setdefault(str(r[0]).strip(),f"{r[1]} {r[3] or ''}")
    for r in mis['MB52'].iter_rows(min_row=2,values_only=True):
        if r[2] and r[5]: desc_of.setdefault(str(r[2]).strip(),str(r[5]))
    tok=lambda t: {w_ for w_ in re.split(r'[^A-Z0-9.]+',str(t or '').upper()) if w_ and w_ not in ('MM','X','FOR','THE','OF','GG','G2G','WITH')}
    def find(text,pool):
        t=str(text or '').strip()
        if t.upper() in {c.upper() for c in pool}: return next(c for c in pool if c.upper()==t.upper()),1.0       # a code
        for c,nm in names.items():                                                                              # an MRP material name
            if c in pool and nm.upper()==t.upper(): return c,1.0
        tt=tok(t)
        def sc_(kv):          # best of: words of the description, or how much of what you typed is found in it
            tk=tok(kv[1]); return max(len(tt&tk)/max(1,len(tt|tk)), len(tt&tk)/max(1,len(tt))*0.9)
        best=max(pool.items(),key=sc_,default=(None,''))
        sc=sc_(best) if best[0] else 0
        return (best[0],sc) if sc>=0.5 else (None,sc)
    for c_,n_,*_r in ADDED_MATS: dash.setdefault(str(c_).strip(),str(n_ or ''))      # materials added because the plan needs them are MRP materials too
    pf={}; lines=[]; SUBS_X[:]=[]
    exp_=[]                                                                      # 'A or B' / 'A; B' = several substitutes in one cell
    for x in SUBS:
        parts_=[p_.strip() for p_ in re.split(r'\s+or\s+|\s*;\s*|\s*\|\s*',str(x['sub'] or ''),flags=re.I) if p_.strip()]
        for p_ in (parts_ or [x['sub']]): exp_.append(dict(x,sub=p_))
    for x in exp_:
        m,sm=find(x['mrp'],dash); c,sc=find(x['sub'],{**desc_of,**dash})
        if not m or not c or c==m:
            iss.append(dict(area='Master data',check='Substitute not understood (Settings)',sev='Medium',material=str(x['mrp']),code=str(x['sub']),a_lab='',a=None,b_lab='',b=None,diff=None,
                            detail=f"MRP material '{x['mrp']}' -> {m or 'not found'}; substitute '{x['sub']}' -> {c or 'not found'}",action='Use the code, or more words from the description')); continue
        x['mrp_code'],x['sub_code']=m,c; x['sub_desc']=desc_of.get(c,dash.get(c,'')); SUBS_X.append(x)
        iss.append(dict(area='Master data',check='Substitute read from Settings (check it is the right material)',sev='Low',material=f"{x['mrp']} -> {m}",code=c,a_lab='',a=None,b_lab='Factor',b=x['factor'],diff=None,
                        detail=f"'{x['sub']}' found as {c} {x['sub_desc'][:50]}; "+('all lines' if x['lines'] is None else 'lines: '+', '.join(sorted(x['lines'])))+
                               (f"; from {x['fr']:%d-%b}" if x['fr'] is not None else '')+(f" until {x['to']:%d-%b}" if x['to'] is not None else ''),action='If wrong, type the code instead'))
        if x['lines'] is None: pf[c]=(m,x['factor'],x['fr'],x['to'])
        else: lines.append(x)
    g['PACK_FACTORS']={**_PACK_FACTORS_SCRIPT,**pf}; g['SUBS_LINES']=lines      # script entries kept (review finding Nov 2026)
    return iss

def make_settings_sheet(path, src=None):
    from openpyxl.styles import Border, Side
    AR='Calibri'; NAVY='1F3864'; thin=Side(style='thin',color='BFBFBF'); bd=Border(left=thin,right=thin,top=thin,bottom=thin)
    wb=Workbook(); ws=wb.active; ws.title='Settings'; ws.sheet_view.showGridLines=False
    ws['A1']='Settings - all planning rules in one place. Type in the yellow cells, save, run the script'; ws['A1'].font=Font(name=AR,size=14,bold=True,color=NAVY)
    ws['A2']=('Tip: copy this sheet into your MRP.xlsx once (right-click the tab > Move or Copy > Create a copy) and keep editing it there. Leave a rule blank when there is no real constraint.'
              +(f'  Read this run from: {os.path.basename(src)}' if src else '  (no Settings sheet found - defaults)'))
    ws['A2'].font=Font(name=AR,size=9,italic=True)
    yel=PatternFill('solid',fgColor='FFF2CC'); hf=Font(name=AR,size=10,bold=True,color='FFFFFF'); hfill=PatternFill('solid',fgColor=NAVY)
    def title(r,t,sub=''):
        ws.cell(r,1,t).font=Font(name=AR,size=12,bold=True,color=NAVY)
        if sub: ws.cell(r+1,1,sub).font=Font(name=AR,size=9,italic=True,color='595959'); return r+2
        return r+1
    def head(r,cols):
        for j,t in enumerate(cols,1): c=ws.cell(r,j,t); c.font=hf; c.fill=hfill; c.border=bd; c.alignment=Alignment(wrap_text=True,vertical='center')
        return r+1
    def numrows(r,rows_):
        r=head(r,['Setting','Value','Note'])
        for lab,nm,expl in rows_:
            v=globals()[nm]; v=('Yes' if v else 'No') if isinstance(v,bool) else v
            ws.cell(r,1,lab).border=bd; c=ws.cell(r,2,v); c.fill=yel; c.border=bd; c.alignment=Alignment(horizontal='center'); ws.cell(r,3,expl).border=bd; r+=1
        return r
    def table(r,cols,data,yellow_cols,n_blank=5,date_cols=()):
        r=head(r,cols)
        for x in list(data)+[[None]*len(cols)]*n_blank:
            for j in range(len(cols)):
                v=x[j] if j<len(x) else None
                if isinstance(v,pd.Timestamp): v=v.to_pydatetime()
                c=ws.cell(r,j+1,v); c.border=bd
                if j in yellow_cols or not any(x): c.fill=yel
                if j in date_cols: c.number_format='dd-mmm-yy'
            r+=1
        return r
    r=4
    r=title(r,'BASIC SETTINGS'); r=numrows(r,BASIC_ROWS)
    tg_=sorted(LINE_TARGET.items()) or [(f'L{i}',(f'USA Line {i} Production Plan Daily Target',None)) for i in range(1,5)]
    for _,(lab,v) in tg_:                  # modules per day each line should make - used for 'days stock + in transit lasts' and the daily report
        ws.cell(r,1,lab).border=bd; c=ws.cell(r,2,v); c.fill=yel; c.border=bd; c.alignment=Alignment(horizontal='center')
        ws.cell(r,3,'Modules per day').border=bd; r+=1
    r+=1
    r=title(r,'TRANSIT LEAD TIME (days, per material)','Every day threshold of a material = its transit lead time (Usage/Day window, ALERT under / below lead time, safety stock, look-ahead). '
            'Type a number in column F "Your lead time" to set it; blank = WSA history (median ETD -> plant days of your WSA deliveries), else the fallback above. Copy this sheet into your MRP.xlsx to keep your numbers.')
    r=table(r,['Code','Material','Type','WSA history (median ETD -> plant days)','Deliveries in history','Your lead time (days)','Used this run','Source'],
            [list(x) for x in sorted(LT_TABLE,key=lambda x:(x[2]!='Raw material',str(x[1] or ''),str(x[0])))]
            +[[k,'(not in the MRP this run)',None,None,None,v,None,'Your lead time - material not in this run'] for k,v in sorted(LT_OVERRIDE.items()) if k not in {str(x[0]).upper() for x in LT_TABLE}],[5],0)+1
    r=title(r,'UNLIMITED MATERIALS','Never short - one material code, or a word from the description, per row')
    r=table(r,['Code or description word'],[[v] for v in sorted(UNLIMITED_CODES)+list(UNLIMITED_WORDS)],[0],3)+1
    r=title(r,'ALLOWED VENDORS (by customer)','A customer may use ONLY the vendors listed for a material (several rows = several vendors, top row preferred). '
            'No row / blank Supplier = any vendor. Stock of unknown vendor is always allowed.')
    r=table(r,['Customer','Material','Code','Supplier'],[[x.get('Customer'),x.get('Material'),x.get('Code'),x.get('Supplier')] for x in CUSTSUP_ROWS],[3])+1
    r=title(r,'SUBSTITUTES','Another material that can replace an MRP material. Type the CODE or words from the DESCRIPTION (the script finds it). '
            'Lines blank = all lines (counted in the material stock); otherwise only for the lines listed (e.g. USA Line 1, USA Line 2). Dates blank = always.')
    r=table(r,['MRP material (code or description)','Substitute (code or description)','Factor (MRP units per 1 unit)','Lines (blank = all)','Use from','Use until','Order (1 = first)'],
            [[x['mrp'],x['sub'],x['factor'],', '.join(sorted(x['lines'])) if x['lines'] else None,x['fr'],x['to'],x['order']] for x in SUBS],[1,2,3,4,5,6],5,(4,5))+1
    r=title(r,'CUSTOMER DAILY TARGET','Daily need of a material = sum over these customers of: daily target x usage per module (MIS Standard BOM / your BOM for CS-PM), '
            'like your PPC Usage sheet. OEM = Yes counts only in the "with OEM" columns.')
    r=table(r,['Customer','Line','Daily target (modules)','OEM (Yes / No)'],[[c_,l_,v_,o_] for c_,l_,v_,o_ in (CUST_TARGET or CUST_TARGET_DEFAULT)],[3],3)+1
    r=title(r,'BATCH SUPPLIER CODES','The 2nd and 3rd characters of an SAP batch tell the supplier (1JU0369004 -> JU). Type / correct the supplier name in column B. '
            '"(guess)" = the script guessed it from PPC In_Transit suppliers of the same materials - overwrite it with the right name')
    bc_=sorted(set(BATCH_SUP)|set(BATCH_SEEN))
    r=table(r,['Batch code (2 letters)','Supplier name','Seen on materials (MB52 batches)'],
            [[k,BATCH_SUP.get(k) or ((BATCH_SEEN.get(k) or {}).get('guess') and f"{BATCH_SEEN[k]['guess']} (guess)"),
              ', '.join(sorted((BATCH_SEEN.get(k) or {}).get('mats') or []))[:120] or None] for k in bc_],[1],2)+1
    r=title(r,'LINE RANKING','Which line gets a scarce material first: 1 = first, 2 = second ... per material row (only materials used by 2 or more lines are listed). '
            'Row "All materials" = default for every material. Blank = closest to stopping first')
    lines_=list(LINES_SEEN) or sorted({_line_key(k) for v in LINE_RANK_MAT.values() for k in v if not k.startswith('_')})
    lk_=[_line_key(n) for n in lines_]
    data=[['All materials (default)','','']+[LINE_RANK.get(k) for k in lk_]]
    seen_=set()
    for nm,ds,cd,lns in SHARED_MATS:
        rk=LINE_RANK_MAT.get(cd) or next((v for k,v in LINE_RANK_MAT.items() if k.upper() in (str(nm).upper(),str(ds).upper())),{})
        data.append([nm,ds,cd]+[(rk.get(k) if k in lns else '-') for k in lk_]); seen_.add(cd)
    for k,v in LINE_RANK_MAT.items():                    # rows you typed for materials not shared this run - kept
        if k not in seen_ and v.get('_name') is not None and k not in {d[2] for d in data}: data.append([v.get('_name'),v.get('_desc'),k]+[v.get(x) for x in lk_])
    r=table(r,['Material (common name)','Component description','Code']+lines_,data,list(range(3,3+len(lines_))),2)+1
    r=title(r,'MIS LINK (Production Plan customer -> MIS plan / BOM)','How each customer of your Production Plan is linked to the MIS plan (for raw-material usage). '
            'Found by the script: same day + same production line in both plans (no names needed). Only if it is wrong: type the MIS BOM Name (and MIS customer) in the yellow cells')
    lk_rows=[[cu,L_.get('mis'),L_.get('bom'),(f"{L_['how']}"+(f" ({L_['conf']:.0%})" if L_.get('conf') is not None else '')),
              (link_override(cu,quiet=True) or (None,None))[1],(link_override(cu,quiet=True) or (None,None))[0]] for cu,L_ in sorted(LINKS.items())]
    in_=({_norm(cu) for cu in LINKS})                       # your overrides of customers without a plan this run are kept (user 22-Nov-2026: Crimson / PGR were dropped)
    lk_rows+=[[LINK_OVR_NAME.get(k_,k_),None,None,'your override (Settings) - customer not in the plan this run',v_[1],v_[0]] for k_,v_ in sorted(LINK_OVERRIDE.items()) if k_ not in in_]
    r=table(r,['Production Plan customer','MIS customer (found)','MIS BOM Name (found)','How it was found','Override: MIS BOM Name','Override: MIS customer'],lk_rows,[4,5],2)+1
    r=title(r,'ADVANCED SETTINGS','Leave as they are unless you have a reason'); r=numrows(r,ADV_ROWS)
    for col,w in zip('ABCDEFG',[46,40,16,22,12,12,12]): ws.column_dimensions[col].width=w
    for j in range(8,16): ws.column_dimensions[L(j)].width=12
    set_print(ws,'Settings'); wb.save(path)

def segment(DS, mis_path, H=None):
    """ABC by value (next H days of gross requirement x unit price from PPC In_Transit 'Unit Rate'), XYZ by demand variability
       (coefficient of variation of weekly requirement over the next 13 weeks). Sets b['abc'], b['xyz'], b['value'], b['price'], b['td']."""
    H=H or ABC_HORIZON
    mis=wb_ro(mis_path); alias=DS.get('alias') or {}; fac=DS.get('alias_factor') or {}
    num={}; den={}
    for r in list(mis['In Transit'].iter_rows(values_only=True))[1:]:
        if len(r)<40 or not isinstance(r[8],(int,float)) or r[8]<=0: continue
        c=str(r[39] or '').strip(); q=r[34] if isinstance(r[34],(int,float)) and r[34]>0 else (r[7] if isinstance(r[7],(int,float)) and r[7]>0 else 1.0)
        f=1.0
        if c in alias: f=fac.get(c,1.0) or 1.0; c=alias[c]
        num[c]=num.get(c,0.0)+(r[8]/f)*q*f; den[c]=den.get(c,0.0)+q*f      # price per MRP unit, weighted by quantity
    rows=[]
    for b in DS['blocks']:
        b['price']=(num[b['mat']]/den[b['mat']]) if den.get(b['mat']) else None
        P=b['plan']; b['value']=(sum(P[:H])*b['price']) if b['price'] else None
        wk=[sum(P[i:i+7]) for i in range(0,min(len(P),91),7)]
        m_=sum(wk)/len(wk) if wk else 0
        cv=((sum((x-m_)**2 for x in wk)/len(wk))**0.5/m_) if m_ else None
        b['cv']=cv; b['xyz']='-' if cv is None else ('X' if cv<XYZ_X else 'Y' if cv<XYZ_Y else 'Z')
    rm=[b for b in DS['blocks'] if not b.get('count_stock') and b['value']]
    tot=sum(b['value'] for b in rm) or 1; acc=0
    for b in sorted(rm,key=lambda b:-b['value']):
        acc+=b['value']; b['abc']='A' if acc-b['value']<ABC_A_PCT/100*tot else ('B' if acc-b['value']<ABC_B_PCT/100*tot else 'C')
    for b in DS['blocks']:
        if b.get('count_stock'): b['abc']='PM'
        else: b.setdefault('abc','C')
    return rows

def make_segment_xlsx(DS, path, H=90):
    from openpyxl.styles import Border, Side
    from openpyxl.worksheet.table import Table, TableStyleInfo
    AR='Calibri'; NAVY='1F3864'; thin=Side(style='thin',color='BFBFBF'); bd=Border(left=thin,right=thin,top=thin,bottom=thin)
    wb=Workbook(); ws=wb.active; ws.title='Segmentation'
    ws['A1']='Segmentation - ABC by value, XYZ by demand variability, and the target days each class gets'; ws['A1'].font=Font(name=AR,size=13,bold=True,color=NAVY)
    ws['A2']=(f'ABC: value = next {H} days of requirement x unit price (PPC In_Transit Unit Rate); A = top 80% of value, B = next 15%, C = rest or no price. '
              'XYZ: weekly requirement variability over 13 weeks (X stable < 0.25, Y 0.25-0.5, Z erratic > 0.5). PM = packaging / consumables. Target days are set on the Settings sheet.')
    ws['A2'].font=Font(name=AR,size=9,italic=True); ws.merge_cells('A2:J2'); ws['A2'].alignment=Alignment(wrap_text=True); ws.row_dimensions[2].height=30
    H_=['Material','Code','UOM','Class (ABC)','Variability (XYZ)','CV weekly','Unit price','Value next %d d'%H,'Share of value','Target days']
    for j,h in enumerate(H_,1):
        c=ws.cell(4,j,h); c.font=Font(name=AR,size=9,bold=True,color='FFFFFF'); c.fill=PatternFill('solid',fgColor=NAVY); c.border=bd; c.alignment=Alignment(horizontal='center',wrap_text=True)
    tot=sum(b['value'] for b in DS['blocks'] if b.get('value')) or 1
    rows=sorted(DS['blocks'],key=lambda b:rrank(b['mat']))          # highest shortage risk first
    for i,b in enumerate(rows):
        vals=[b['desc'] if b.get('count_stock') else b['common'],b['mat'],b.get('uom'),b['abc'],b['xyz'],round(b['cv'],2) if b.get('cv') is not None else None,
              round(b['price'],4) if b.get('price') else None,round(b['value']) if b.get('value') else None,(b['value']/tot) if b.get('value') else None,b['td']]
        for j,v in enumerate(vals,1):
            c=ws.cell(5+i,j,v); c.font=Font(name=AR,size=9,bold=j in (4,10)); c.border=bd
            if j==8: c.number_format='#,##0'
            if j==9: c.number_format='0.0%'
        ws.cell(5+i,4).fill=PatternFill('solid',fgColor={'A':'FFC7CE','B':'FFEB9C','C':'C6EFCE','PM':'DDEBF7'}.get(b['abc'],'FFFFFF'))
    ws.add_table(Table(displayName='Segmentation',ref=f'A4:J{4+len(rows)}',tableStyleInfo=TableStyleInfo(name='TableStyleLight1',showRowStripes=True)))
    for j,w in enumerate([32,18,6,9,9,9,10,14,9,9],1): ws.column_dimensions[L(j)].width=w
    wb.save(path)


LINE_MODS={}       # production line key -> modules per day (your Production Plan) - for 'modules/day affected' in the emails
def material_report(A, d0, fname=''):
    """ONE short daily report, easy to read as a text message: materials whose stock + in transit lasts fewer than
       their lead time at the module plan, shortest first."""
    tg=([(c_,v_) for c_,l_,v_,o_ in CUST_TARGET] if CUSTOMER_TARGETS else [(lab.split(' Production')[0].replace('USA ','').replace(' (plan peak)','*'),v) for _,(lab,v) in sorted(EFF_TARGET.items()) if v])
    low=sorted([x for x in A if x.get('pl_days') is not None and x['pl_days']<tgt_days(x['b'].get('lt_used'))],key=lambda x:((x.get('tr_tot') or 0)>0,x['pl_days']))
    n=len(low)
    subj=f"Material report {d0:%d %b}: {n} under lead time" if n else f"Material report {d0:%d %b}: all OK (lead time or more)"
    body=[]
    if not tg: body.append('No line daily targets on Settings - no report.')
    else:
        body.append(f"Under lead time at the module plan - nothing in transit first")
        body.append('('+' / '.join(f'{a} {v:,.0f}' for a,v in tg)+')')
        if any(a.endswith('*') for a,_ in tg): body.append('* no target on Settings: highest planned day')
        body.append('')
        for i,x in enumerate(low,1):
            b=x['b']; u=b.get('uom') or ''
            body.append(f"{i}) {b['common']}")
            body.append(f"   {x['pl_days']:.1f} days | stock {b['stock']:,.0f} | in transit {x.get('tr_tot') or 0:,.0f} {u}".rstrip())
            ln_=[EFF_TARGET[k_][0].split(' Production')[0].replace(' (plan peak)','').replace('USA ','') for k_ in (x.get('tgt_lines') or []) if k_ in EFF_TARGET]
            if ln_: body.append(f"   Lines: {', '.join(ln_)}")
        if not low: body.append(f"All materials lead time or more.")
    return subj,body

def unfreeze_all(path, keep=()):
    """Remove every frozen pane from the output workbook (edits the sheet XML directly, so formulas, cached values and formats stay),
       except on the sheets named in keep."""
    import zipfile, shutil
    tmp_=path+'.unfreeze'
    with zipfile.ZipFile(path) as z0:
        wbx=z0.read('xl/workbook.xml').decode('utf-8'); rels=z0.read('xl/_rels/workbook.xml.rels').decode('utf-8')
    keep_parts=set()
    for n_ in keep:
        m_=re.search(r'<sheet\b[^>]*name="'+re.escape(n_)+r'"[^>]*r:id="([^"]+)"',wbx)
        if m_:
            t_=re.search(r'<Relationship\b[^>]*Id="'+m_.group(1)+r'"[^>]*Target="([^"]+)"',rels) or re.search(r'<Relationship\b[^>]*Target="([^"]+)"[^>]*Id="'+m_.group(1)+'"',rels)
            if t_: keep_parts.add('xl/'+t_.group(1).lstrip('/').replace('xl/','',1))
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp_,'w',zipfile.ZIP_DEFLATED) as zout:
        for it in zin.infolist():
            data=zin.read(it.filename)
            if it.filename.startswith('xl/worksheets/sheet') and it.filename.endswith('.xml') and it.filename not in keep_parts:
                t=data.decode('utf-8')
                t=re.sub(r'<pane\b[^>]*/>','',t)                                   # frozen / split pane
                t=re.sub(r'(<selection\b[^>]*?)\s+pane="[^"]*"',r'\1',t)            # selections that pointed at a pane
                data=t.encode('utf-8')
            zout.writestr(it,data)
    shutil.move(tmp_,path)

def reorder_sheets(pkg_path, first):
    """Puts the sheets named in 'first' at the front in that order (others keep their order after them) by editing the
       workbook XML only - charts and saved formula results stay. Sheet-scoped names (print titles) follow their sheet."""
    import zipfile, shutil
    from lxml import etree
    M='http://schemas.openxmlformats.org/spreadsheetml/2006/main'; q=lambda t:'{%s}%s'%(M,t)
    z=zipfile.ZipFile(pkg_path); parts={n:z.read(n) for n in z.namelist()}; z.close()
    wb=etree.fromstring(parts['xl/workbook.xml']); sh=wb.find(q('sheets')); old=list(sh)
    new=[e for n in first for e in old if e.get('name')==n]+[e for e in old if e.get('name') not in first]
    pos={id(e):i for i,e in enumerate(old)}; newidx={pos[id(e)]:i for i,e in enumerate(new)}
    for e in old: sh.remove(e)
    for e in new: sh.append(e)
    dn=wb.find(q('definedNames'))
    if dn is not None:
        for d in dn:
            if d.get('localSheetId') is not None: d.set('localSheetId',str(newidx.get(int(d.get('localSheetId')),int(d.get('localSheetId')))))
    bv=wb.find(q('bookViews'))
    if bv is not None:
        for v in bv: v.set('activeTab','0'); v.set('firstSheet','0')
    parts['xl/workbook.xml']=etree.tostring(wb,xml_declaration=True,encoding='UTF-8',standalone=True)
    # only the first sheet is 'selected'
    rels=etree.fromstring(parts['xl/_rels/workbook.xml.rels']); tg={r.get('Id'):r.get('Target') for r in rels}
    RID='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
    for i,e in enumerate(new):
        pth='xl/'+tg[e.get(RID)].lstrip('/').replace('xl/','')
        x=etree.fromstring(parts[pth])
        for sv in x.iter(q('sheetView')):
            if i==0: sv.set('tabSelected','1')
            elif 'tabSelected' in sv.attrib: del sv.attrib['tabSelected']
        parts[pth]=etree.tostring(x,xml_declaration=True,encoding='UTF-8',standalone=True)
    tmpz=pkg_path+'.reorder'
    with zipfile.ZipFile(tmpz,'w',zipfile.ZIP_DEFLATED) as zo:
        for n,b in parts.items(): zo.writestr(n,b)
    shutil.move(tmpz,pkg_path)


ALLOC_STOP_={}
def module_short_names(mis_path, bns):
    """Short module name from the MIS BOM name: drop '500 TC G2G_' and 'USA', add the cell size (G12R / M10R) when the name
       does not say it (read from the BOM's glass / frame), e.g. '500 TC G2G_G12R_3.2' -> 'G12R 3.2', '500 TC G2G_EPEP_USA' -> 'M10R EPEP'."""
    out={}; size={}; hs={}
    try:
        cn={}
        for r in wb_ro(mis_path)['Common Name'].iter_rows(min_row=2,values_only=True):
            if r and r[0]: cn[str(r[0]).strip()]=str(r[3] or '')
        for r in wb_ro(mis_path)['Standard BOM'].iter_rows(min_row=2,values_only=True):
            if r and len(r)>4 and r[2] and r[3]:
                t=(cn.get(str(r[3]).strip(),'')+' '+str(r[4] or '')).upper()
                mh=re.search(r'(\d\d)\s*HS',t)
                if mh and str(r[8] or '').lower().startswith('alum'): hs.setdefault(str(r[2]).strip(),mh.group(1)+'HS')
                for z in ('G12R','M10R','G12','M10'):
                    if z in t: size.setdefault(str(r[2]).strip(),{}); size[str(r[2]).strip()][z]=size[str(r[2]).strip()].get(z,0)+1; break
    except Exception as _e_: _swallow(6251,_e_); pass
    for bn in bns:
        if not bn: continue
        t=re.sub(r'^\d+\s*(TC|MP)\s*G2G_?','',str(bn).strip())
        parts=[p_ for p_ in re.split(r'[_\s]+',t) if p_ and p_.upper()!='USA']
        nm=' '.join(parts) or str(bn)
        if not re.search(r'G12|M10',nm,re.I) and size.get(str(bn).strip()):
            nm=f"{max(size[str(bn).strip()],key=size[str(bn).strip()].get)} {nm}"
        if re.fullmatch(r'G12R?',nm) and hs.get(str(bn).strip()): nm=f"{nm} {hs[str(bn).strip()]}"     # plain G12R: add the frame height
        out[bn]=nm
    return out

def plan_windows(arr, d0, gap=3):
    """Runs of days with demand (gaps of up to 'gap' days without production = same run): [(start_k, end_k)]."""
    out=[]; st=None; last=None
    for k,v in enumerate(arr):
        if v>0:
            if st is None: st=k
            elif k-last>gap+1: out.append((st,last)); st=k
            last=k
    if st is not None: out.append((st,last))
    return out

def future_plan_text(b, lk, d0, nd, k_today):
    """How this line uses the material over the plan, in one short line:
       'Same as today' = used from the line's first production day to the end of the plan by one customer;
       otherwise when it is used (starts later / stops / only a few days) | the customers in date order."""
    per={}
    for (cu,ln),arr in (b.get('cust_line') or {}).items():
        if not ln or _line_key(ln)!=lk or not any(arr): continue
        us=UCMAP.get((_norm(cu),lk)) or {}
        nm=max(us,key=us.get) if us else cu
        a_=per.setdefault(nm,[0.0]*nd)
        for k in range(min(nd,len(arr))): a_[k]+=arr[k]
    if not per: return None,False
    tot=[sum(a_[k] for a_ in per.values()) for k in range(nd)]
    k0=k_today or 0; dd=lambda k: f"{(d0+pd.Timedelta(days=k)):%d-%b}"
    w=plan_windows(tot,d0)
    whole=len(w)==1 and w[0][0]<=k0+3 and w[0][1]>=nd-4
    seq=[]                                            # customer runs in date order (gaps without production ignored)
    for k in range(nd):
        if tot[k]<=0: continue
        nm=max(per,key=lambda n_: per[n_][k])
        if not seq or seq[-1][0]!=nm: seq.append((nm,k))
    if whole and len(seq)==1: return 'Same as today',False
    if whole: use='Used all plan'
    else:
        txt=[]
        for a,z in w: txt.append(f"{dd(a)} → "+("end of plan" if z>=nd-4 else dd(z)))
        use=('Starts later: ' if w[0][0]>k0+3 else 'Used ')+'; '.join(txt)+('' if w[-1][1]>=nd-4 else ' (then not used)')
    cust=' → '.join(f"{nm} ({dd(k)})" for nm,k in seq)
    return f"{use} | {cust}",True

def line_plan_need(b, lk, nd):
    """Module plan Usage/Day of one material on one production line: average per production day (> 0), over the transit lead time."""
    arr=[0.0]*min(nd,lt_of(b))
    for (cu_,ln_),a_ in (b.get('cust_line') or {}).items():
        if ln_ and _line_key(ln_)==lk:
            for k in range(len(arr)): arr[k]+=a_[k] if k<len(a_) else 0.0
    nz=[v for v in arr if v>0]
    return round(sum(nz)/len(nz),1) if nz else None
def make_line_view_xlsx(A, LR, d0, nd, path, TD=None, AL_COLS=None):
    """MRP by Line: one row per material and production line, sorted like the Action List (material, then line).
       Line columns: how long the line can run, this line's need at its daily target and at its daily plan (average of the next
       21 days with production), when this line is short (allocated) and where it is supplied from.
       Then the Action List columns Stock ... Daily plan with OEM for the material (formulas = the Action List row),
       repeated on every line row of the material (no merged cells), so filtering a line still shows them."""
    from openpyxl.styles import Border, Side
    AR='Calibri'; NAVY='1F3864'; TEAL='1F6E6E'; thin=Side(style='thin',color='BFBFBF'); bd=Border(left=thin,right=thin,top=thin,bottom=thin)
    wb=Workbook(); ws=wb.active; ws.title='MRP by Line'; ws.sheet_view.showGridLines=False
    ws['A1']='MRP by Line - each material per production line, with the module-plan numbers of the material'; ws['A1'].font=Font(name=AR,size=14,bold=True,color=NAVY)
    ws['A2']=('Filter column A (line) and / or column E (component). Column D = the material(s) that stop the line on the Line stops on date. Teal columns = this line only. The other columns are the material totals (module plan: Usage/Day, '
              'Inv. Days, below the lead time from, days late - the same numbers as the ALERT); they are repeated on every line row of the material, so a line filter keeps them visible. '
              '')
    ws['A2'].font=Font(name=AR,size=9,italic=True); ws.merge_cells('A2:N2'); ws['A2'].alignment=Alignment(wrap_text=True,vertical='top'); ws.row_dimensions[2].height=40
    LCOL=['Production line','Line can run (days)','Line stops on','Line stops because of','Component','Component description','UOM','First needed (this line)','Needed now (within lead time)',
          'Plan on this line (who / when) - differs from today?',
          'This line: Usage/Day','This line: short from (allocated)','This line: supplied from (in order)']
    AL_COLS=[(h_,c_) for h_,c_ in (AL_COLS or []) if 'oem' not in h_.lower() and 'daily target' not in h_.lower()]   # targets / OEM: Action List only
    H=LCOL+[('In transit usable by this line (PPC In_Transit, allowed suppliers)' if h_==IT_TODAY_HEAD else h_) for h_,_ in AL_COLS]
    for j_,h in enumerate(H,1):
        c=ws.cell(4,j_,h); c.border=bd; c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
        if j_<=len(LCOL): col_,txt_=(TEAL if j_>=8 else NAVY),'FFFFFF'
        else: col_=al_hdr_color(h); txt_='1F1F1F' if al_hdr_light(h) else 'FFFFFF'
        c.fill=PatternFill('solid',fgColor=col_); c.font=Font(name=AR,size=9,bold=True,color=txt_)
    ws.row_dimensions[4].height=48
    info={r_['key']:r_ for r_ in (LR or [])}; lorder={r_['key']:i_ for i_,r_ in enumerate(LR or [])}
    order=[x for x in (AL_ORDER_ or A) if not x.get('sup_row')] or A          # same order as the Action List: highest shortage risk first
    r=5
    for x in order:
        b=x['b']; lines={}
        for (cu,ln),arr in (b.get('cust_line') or {}).items():
            if not ln or not any(arr): continue
            a_=lines.setdefault(_line_key(ln),[0.0]*nd)
            for k in range(min(nd,len(arr))): a_[k]+=arr[k]
        if not lines: continue
        r0=r
        for lk in sorted(lines,key=lambda k_:(lorder.get(k_,99),k_)):
            a_=lines[lk]; L_=info.get(lk); kt,limt,_=L_['t'] if L_ else (nd,[],[])
            f_=next((k for k in range(nd) if a_[k]),None)
            fut_txt,fut_diff=future_plan_text(b,lk,d0,nd,LINE_TODAY.get(lk))
            vals=[(L_['name'] if L_ else lk),(kt if kt<nd else f'{nd}+'),((d0+pd.Timedelta(days=kt)).to_pydatetime() if kt<nd else None),
                  ('; '.join(L_.get('stop_by') or []) if (L_ and kt<nd) else ''),b['mat'],
                  (b['desc'] if b.get('count_stock') else b['common']),b.get('uom'),((d0+pd.Timedelta(days=f_)).to_pydatetime() if f_ is not None else None),
                  (('Yes' if (f_ is not None and f_<tgt_days(b.get('lt_used'))) else 'No') if VALUES_ONLY else f'=IF(AND(ISNUMBER(H{r}),H{r}<DATE({d0.year},{d0.month},{d0.day})+{tgt_days(b.get("lt_used"))}),"Yes","No")'),
                  fut_txt,
                  line_plan_need(b,lk,nd),                                    # this line's module plan x usage, average per production day (9 weeks)
                  ((d0+pd.Timedelta(days=ALLOC_STOP_[(b['mat'],lk)])).to_pydatetime() if ALLOC_STOP_.get((b['mat'],lk),nd)<nd else
                   ('No demand' if x.get('no_demand') else 'covered here – short on other lines' if (x.get('wins') or []) else 'covered')),
                  (x.get('used') or {}).get(lk) or '']
            if VALUES_ONLY:
                vals+=[(round(ALLOC_IT[(b['mat'],lk)]) if (h_==IT_TODAY_HEAD and (b['mat'],lk) in ALLOC_IT) else AL_VAL.get((b['mat'],h_))) for h_,_ in AL_COLS]
            else: vals+=[f"=INDEX('Action List'!${cl}$5:${cl}$1000,MATCH($E{r},'Action List'!$A$5:$A$1000,0))" for _,cl in AL_COLS]
            for j_,v in enumerate(vals,1):
                c=ws.cell(r,j_,v); c.font=Font(name=AR,size=9); c.border=bd
                c.alignment=Alignment(vertical=('top' if (j_<=len(LCOL) or 'Arrival' in H[j_-1]) else 'center'),wrap_text=j_ in (6,10,13) or (j_>len(LCOL) and 'Arrival' in H[j_-1]))
                if j_ in (3,8): c.number_format='dd-mmm-yy'
                if j_==4 and v: c.font=Font(name=AR,size=9,bold=True,color='9C0006')
                if j_==11: c.number_format='#,##0.0'
                if j_==9: c.alignment=Alignment(horizontal='center',vertical='top')
                if j_==12 and isinstance(v,dt.datetime): c.number_format='dd-mmm'; c.font=Font(name='Calibri',size=10,bold=True,color='9C0006')
                if j_>len(LCOL):
                    hh=H[j_-1]; c.number_format=('dd-mmm-yy' if isinstance(v,dt.datetime) else '0.0' if hh.startswith(('Inv. Days','Days of cover')) else ('#,##0.0' if 'Daily' in hh else '#,##0;-#,##0;"-"'))
                    if hh.startswith(('Inv. Days','Days of cover')) and isinstance(v,(int,float)) and v<tgt_days(b.get('lt_used')):     # under the lead time: red highlight
                        c.fill=PatternFill('solid',fgColor='FFC7CE'); c.font=Font(name='Calibri',size=10,bold=True,color='9C0006')
            if L_ and kt<nd: ws.cell(r,2).font=Font(name=AR,size=9,bold=True,color='9C0006')
            if fut_diff: ws.cell(r,10).fill=PatternFill('solid',fgColor='FFEB9C')      # the plan on this line is not the same as today
            r+=1
        n_=str(x.get('ata_txt') or '').count('\n')+1; need_h=12.5*n_; cnt_=r-r0      # room for the arrival lines
        if cnt_ and need_h>15*cnt_:
            for rr_ in range(r0,r): ws.row_dimensions[rr_].height=need_h/cnt_
        for j_ in range(1,len(H)+1): ws.cell(r0,j_).border=Border(left=thin,right=thin,bottom=thin,top=Side(style='medium',color=NAVY))
    if r>5: ws.auto_filter.ref=f'A4:{L(len(H))}{r-1}'
    widths=[16,9,10,34,16,30,6,10,9,40,11,10,30]+[(42 if 'Arrival' in h_ else 11) for h_,_ in AL_COLS]
    for j_,w in enumerate(widths,1): ws.column_dimensions[L(j_)].width=w
    set_print(ws,'MRP by Line'); wb.save(path)

def cspm_from_bom_pp(bom, count, pp_rows, dates, sr0, count_hdr=None):
    """CS / PM plan straight from YOUR BOM and YOUR Production Plan, row by row:
       the Production Plan customer is found in the BOM by customer name; if not, by its Module Model (the BOM of that model).
       Returns (blocks, Check rows for Production Plan customers that have no BOM)."""
    nd=len(dates); iss=[]; bom=cspm_alt_fold(bom)
    bycm={}                                            # (customer, model) -> {component: qty per module} (rows of one model add up)
    meta={}
    for d in bom:
        c=str(d.get('Component') or '').strip()
        if c[:2].upper() not in ('CS','PM') or not isinstance(d.get('Quantity'),(int,float)): continue
        cu=_norm(d.get('Customer')); mdl=str(d.get('Module Model') or '').strip()
        bycm.setdefault((cu,mdl),{}); bycm[(cu,mdl)][c]=bycm[(cu,mdl)].get(c,0.0)+float(d['Quantity'])
        meta.setdefault(c,(d.get('Component desciption'),d.get('UoM')))
    bycust={}; bymodel={}
    for (cu,mdl),comp in bycm.items():
        bycust.setdefault(cu,[]).append(comp); bymodel.setdefault(mdl,[]).append(comp)
    def avg(lst):
        out={}
        for comp in lst:
            for c,q in comp.items(): out[c]=out.get(c,0.0)+q/len(lst)
        return out
    plan={}; cust={}; cl={}; how={}; pls={}
    for x in pp_rows or []:
        key=_norm(x['cust']); mdl=str(x.get('model') or '').strip(); cu=(LINKS.get(x['cust']) or {}).get('mis') or x['cust']   # MIS name shown
        if key in bycust: comp=avg(bycust[key]); how[cu]='customer'
        elif mdl and mdl in bymodel: comp=avg(bymodel[mdl]); how[cu]=f'model {mdl[:14]}'
        else:
            if sum(x['v']) and cu not in how:
                how[cu]=None
                iss.append(dict(area='Plan',check='Production Plan customer has no packaging BOM (CS / PM not planned)',sev='Medium',material=x['cust'],code=x['line'],
                                a_lab='',a=None,b_lab='Modules in horizon',b=round(sum(x['v'])),diff=None,
                                detail=f"{x['cust']} ({mdl or 'no model'}): not in your BOM by customer or by model",action='Add its BOM rows, or the Module Model on your Production Plan / Index'))
            continue
        for c,q in comp.items():
            p_=plan.setdefault(c,[0.0]*nd); a_=cust.setdefault(c,{}).setdefault(cu,[0.0]*nd); l_=cl.setdefault(c,{}).setdefault((cu,x['line']),[0.0]*nd)
            for k in range(nd):
                v=q*x['v'][k]; p_[k]+=v; a_[k]+=v; l_[k]+=v
            pls.setdefault(c,[]).append(dict(cu=cu,ln=x['line'],bn=None,rate=q,mods=list(x['v']),src=f"your BOM (packaging, by {how.get(cu) or 'customer'})"))
    qi=count_qty_idx(count_hdr)
    stock=cspm_alt_stock({str(r[0]).strip():(float(r[qi]) if len(r)>qi and isinstance(r[qi],(int,float)) else 0.0) for r in count if r and r[0]})
    hl=[str(h).strip().lower() if h is not None else '' for h in (count_hdr or [])]
    ix=lambda n: hl.index(n) if n in hl else None
    i_q,i_u,i_f=ix('today stock qty'),ix('stock uom'),ix('quantity')
    raw={}
    for r in count:
        if not r or not r[0] or i_q is None: continue
        qv=r[i_q] if len(r)>i_q else None; u=r[i_u] if (i_u is not None and len(r)>i_u) else ''; fv=r[i_f] if (i_f is not None and len(r)>i_f) else None
        if isinstance(qv,(int,float)): raw[str(r[0]).strip()]=f"{qv:,.0f} {u}"+(f" x {fv:g}" if isinstance(fv,(int,float)) and fv not in (0,1) else '')
    out=[]; sr=sr0
    for c in sorted(set(meta)|set(plan)):
        sr+=1; desc=' '.join(str((meta.get(c) or ('',''))[0] or '').split()); P=plan.get(c,[0.0]*nd)
        cp=dict(sorted({k:v for k,v in (cust.get(c) or {}).items() if any(v)}.items(),key=lambda kv:-sum(kv[1])))
        out.append(dict(key=f'{sr}CS/PM',sr=sr,short='CS - Consumable' if c.upper().startswith('CS') else 'PM - Packaging',common=desc,mat=c,desc=desc,mtype='CS/PM',
                        stock=stock.get(c,0.0),uom=(meta.get(c) or ('',''))[1] or '',plan=P,cust_plan=cp,
                        cust_line=dict(sorted({k:v for k,v in (cl.get(c) or {}).items() if any(v)}.items(),key=lambda kv:(-sum(cp.get(kv[0][0],[0])),kv[0][0],str(kv[0][1])))),
                        transit=[0.0]*nd,adhoc=[0.0]*nd,openpo=[0.0]*nd,count_stock=True,raw=raw.get(c,'not on Stock Count')+''.join(f' + {raw[a_]} ({a_})' for a_,s_ in CSPM_ALT.items() if s_==c and a_ in raw),plan_lines=pls.get(c,[]),
                        st_main=stock.get(c,0.0),st_wh=0.0,st_jw=0.0,plan_all={pd.Timestamp(dates[k]).normalize():P[k] for k in range(nd) if P[k]}))
    return out,iss

def make_bom_sheet(bom, mis_names, path):
    from openpyxl.worksheet.datavalidation import DataValidation
    from openpyxl.worksheet.table import Table, TableStyleInfo
    wb=Workbook(); ws=wb.active; ws.title='BOM'
    src_=os.path.basename(PPC_SRC[0]) if PPC_SRC[0] and 'BOM' in _sheetnames(PPC_SRC[0]) else BOM_SRC_NAME[0]
    ws['A1']=(f'Copy of the BOM in {src_} - edit it THERE. MIS BOM Name (link to the MIS plan) comes from the module plan (customers using each Module Model).'
              if BOM_SRC_NAME and BOM_SRC_NAME[0] and not BOM_SRC_NAME[0].startswith('1_MRP') else
              'BOM - your input. Edit, add or delete rows here; the next run uses this sheet. MIS BOM Name links each model to the MIS production plan.')
    ws['A1'].font=Font(name='Arial',size=10,bold=True,color='C00000')
    for j,h in enumerate(BOM_COLS,1):
        c=ws.cell(3,j,h); c.font=Font(name='Arial',size=10,bold=True,color='FFFFFF'); c.fill=PatternFill('solid',fgColor='1F3864')
    for i,d in enumerate(bom):
        for j,h in enumerate(BOM_COLS,1): ws.cell(4+i,j,d[h])
    last=3+max(1,len(bom))
    ws.add_table(Table(displayName='BOMInput',ref=f'A3:{L(len(BOM_COLS))}{last}',tableStyleInfo=TableStyleInfo(name='TableStyleLight9',showRowStripes=True)))
    if mis_names:
        lst=wb.create_sheet('lists'); [lst.cell(i+1,1,n) for i,n in enumerate(mis_names)]
        ws.cell(2,14,'MIS BOM Names: '+' | '.join(mis_names)).font=Font(name='Arial',size=9,italic=True)
        dv=DataValidation(type='list',formula1='"'+','.join(mis_names)[:250]+'"',allow_blank=True); ws.add_data_validation(dv); dv.add(f'K4:K{last+200}')
        del wb['lists']
    for j,w in enumerate([24,42,7,5,22,40,10,7,11,11,26],1): ws.column_dimensions[L(j)].width=w
    ws.freeze_panes='A4'; wb.save(path)

def make_count_sheet(hdr, rows, path, bom=None, cache=None):
    """Stock Count (your input). With the columns Today Stock Qty / Stock UOM / BOM UOM / Quantity / Today Qty_BOM, BOM UOM and
       Today Qty_BOM stay formulas (Today Qty_BOM = Today Stock Qty x Quantity), and the MRP uses Today Qty_BOM."""
    wb=Workbook(); ws=wb.active; ws.title='Stock Count'
    src_=(os.path.basename(PPC_SRC[0])+' (sheet Count)') if PPC_SRC[0] else BOM_SRC_NAME[0]
    ws['A1']=(f'Copy of the stock count in {src_} - update it THERE (Today Stock Qty x Quantity = Today Qty_BOM, used by the MRP).'
              if BOM_SRC_NAME and BOM_SRC_NAME[0] and not BOM_SRC_NAME[0].startswith('1_MRP') else
              'Stock Count - your input (CS / PM physical count). Enter Today Stock Qty in the stock unit and Quantity = BOM units per stock unit '
              '(e.g. 21.8 KG per roll); the MRP uses Today Qty_BOM.')
    ws['A1'].font=Font(name='Arial',size=10,bold=True,color='C00000')
    hdr=list(hdr or ['Material','Material Description','Today Qty'])
    hl=[str(h).strip().lower() if h is not None else '' for h in hdr]
    ix=lambda n: hl.index(n) if n in hl else None
    iq,ibu,ifac,iout=ix('today stock qty'),ix('bom uom'),ix('quantity'),ix('today qty_bom')
    bomu={str(d['Component']).strip():d['UoM'] for d in (bom or []) if d.get('Component') and d.get('UoM')}
    for j,h in enumerate(hdr,1):
        c=ws.cell(3,j,h); c.font=Font(name='Arial',size=10,bold=True,color='FFFFFF'); c.fill=PatternFill('solid',fgColor='1F3864')
        if isinstance(h,(dt.datetime,dt.date)): c.number_format='dd-mmm'
    for i,r in enumerate(rows):
        rr=4+i
        for j,v in enumerate(r,1):
            if (j-1)==ibu:          # BOM UOM looked up from this workbook's BOM sheet (Component = E, UoM = H)
                ws.cell(rr,j,f'=IFERROR(INDEX(BOM!$H:$H,MATCH($A{rr},BOM!$E:$E,0)),"")')
                if cache is not None: cache[(rr,j)]=str(bomu.get(str(r[0]).strip(),'') or '')
            elif (j-1)==iout and iq is not None and ifac is not None and (iq>=len(r) or isinstance(r[iq],(int,float)) or not isinstance(v,str)):   # 'no counting' text stays (not a formula 0)
                ws.cell(rr,j,f'={L(iq+1)}{rr}*IF({L(ifac+1)}{rr}="",1,{L(ifac+1)}{rr})')
                qv=r[iq] if isinstance(r[iq],(int,float)) else 0; fv=r[ifac] if isinstance(r[ifac],(int,float)) else 1
                if cache is not None: cache[(rr,j)]=float(qv)*float(fv)
                ws.cell(rr,j).font=Font(bold=True)
            else: ws.cell(rr,j,None if isinstance(v,str) and v.strip() in ('#N/A','#REF!','#VALUE!','#DIV/0!','#NAME?') else v)
    for j,w in enumerate([18,42,14,10,10,10,15],1): ws.column_dimensions[L(j)].width=w
    if iout is not None:
        for rr in range(4,4+len(rows)): ws.cell(rr,iout+1).fill=PatternFill('solid',fgColor='E2EFDA')
    ws.freeze_panes='C4'; wb.save(path)

def cspm_checks(bom, count, modplan, mis_names, today, H=30):
    I=[]
    def add(area,check,sev,material='',code='',a_lab='',a=None,b_lab='',b=None,detail='',action=''):
        I.append(dict(area=area,check=check,sev=sev,material=material,code=code,a_lab=a_lab,a=a,b_lab=b_lab,b=b,
                      diff=(b-a) if isinstance(a,(int,float)) and isinstance(b,(int,float)) else None,detail=detail,action=action))
    used={str(d['MIS BOM Name'] or '').strip() for d in bom}
    for bn,cu in modplan.items():
        q=sum(sum(v[:H]) for v in cu.values())
        if q and bn not in used:
            add('Plan','MIS BOM Name has no BOM rows','High',bn,'','',None,f'Modules next {H} days',q,
                'Customers: '+', '.join(sorted(cu))+' - no CS / PM plan for these modules','Add BOM rows (or set MIS BOM Name on existing rows) in the BOM sheet')
    for d in bom:
        bn=str(d['MIS BOM Name'] or '').strip()
        if not bn: add('Master data','BOM row without MIS BOM Name','Medium',str(d['Module Model'] or ''),str(d['Component'] or ''),'',None,'',None,'This row is not used in the plan','Fill MIS BOM Name in the BOM sheet')
        elif mis_names and bn not in mis_names: add('Master data','BOM row with unknown MIS BOM Name','Medium',bn,str(d['Component'] or ''),'',None,'',None,'Name not found in MIS Total Plan Req','Pick a name from the list')
    pm_bom={str(d['Component']).strip() for d in bom if str(d['Component'] or '')[:2].upper() in ('CS','PM')}
    cnt={str(r[0]).strip() for r in count if r and r[0]}
    for c in sorted(pm_bom-cnt): add('Stock','CS / PM component not on Stock Count','Medium','',c,'',None,'',None,'Stock is taken as 0','Add the material to the Stock Count sheet')
    return I

# ===================== Customer -> supplier map (your input) =====================
MAP_COLS=['Customer','Material','Code','Supplier','Share %','Note','Suggested supplier']
def _norm(x): return ' '.join(str(x or '').split()).lower()
_LO_SAID=set()
def link_override(cu, quiet=False):
    """Settings MIS LINK override for a plan customer: exact name first, else the same name spelled slightly differently
       (e.g. 'Clenera Crimson' in the PPC plan vs 'Cleanera Crimson' on Settings, similarity >= 0.9). Prints a note once."""
    n=_norm(cu); ov=LINK_OVERRIDE.get(n)
    if ov or not n or not LINK_OVERRIDE: return ov
    import difflib
    k_=max(LINK_OVERRIDE,key=lambda k2:difflib.SequenceMatcher(None,n,k2).ratio())
    if difflib.SequenceMatcher(None,n,k_).ratio()<0.9: return None
    if not quiet and n not in _LO_SAID:
        _LO_SAID.add(n); print(f"Settings MIS LINK '{k_}' used for plan customer '{cu}' (spelling differs - make them the same to remove this note)",flush=True)
    return LINK_OVERRIDE[k_]

OEM_PLAN_CUST=set()   # customers with modules on the PPC 'OEM Plan' sheet (filled by ppc_plan_sheet): OEM like Settings OEM = Yes (user 28-Nov-2026)
def _oem_hit(n, ns): return bool(n) and any(n==o_ or n.startswith(o_) or o_.startswith(n) for o_ in ns)
def oem_names():
    """THE list of OEM customers (normalised names) - one rule for every calculation (user 28-Nov-2026: OEM plan in NO calculation):
       Settings CUSTOMER DAILY TARGET OEM = Yes + every customer on the PPC 'OEM Plan' sheet, each with the MIS customer it is linked to
       (LINKS, e.g. Desri -> Banita (Encap)). A MIS customer that a non-OEM plan customer also links to stays in (it serves the plant lines)."""
    ns={_norm(cu) for cu,ln,tg,o in CUST_TARGET if str(o).strip().lower()=='yes'}|{_norm(c_) for c_ in OEM_PLAN_CUST}
    ns.discard('')
    keep={_norm(L_.get('mis')) for c_,L_ in LINKS.items() if L_.get('mis') and not _oem_hit(_norm(c_),ns)}
    return ns|{_norm(L_['mis']) for c_,L_ in LINKS.items() if L_.get('mis') and _oem_hit(_norm(c_),ns) and _norm(L_['mis']) not in keep}
def is_oem_cust(cu, ns=None):
    """True = customer cu (plan or MIS name) is OEM: its plan is left out of every calculation (`oem_names`)."""
    return _oem_hit(_norm(cu),oem_names() if ns is None else ns)

def read_supplier_map(prev_path):
    rows=[]
    if not prev_path or not os.path.exists(prev_path): return rows
    try: wb=wb_ro(prev_path)
    except Exception as _e_: _swallow(6548,_e_); return rows
    if 'Customer Supplier' not in wb.sheetnames: return rows
    data=[list(r) for r in wb['Customer Supplier'].iter_rows(values_only=True)]
    hr=next((i for i,r in enumerate(data) if 'Customer' in r and 'Supplier' in r),None)
    if hr is None: return rows
    h=[str(x).strip() if x is not None else '' for x in data[hr]]
    for r in data[hr+1:]:
        d={c:(r[h.index(c)] if c in h and h.index(c)<len(r) else None) for c in MAP_COLS}
        if d['Customer'] and (d['Code'] or d['Material']): rows.append(d)
    return rows

def apply_supplier_map(DS, maprows, add_new=False):
    """Puts each supplier's share of the plan on its supplier block. The sheet is yours: rows are only added when add_new
       (first creation); customers missing from it are reported on Check instead. Returns (rows, issues for Check)."""
    nd=len(DS['dates']); issues=[]
    def add(sev,check,material,code,detail,action):
        issues.append(dict(area='Plan',check=check,sev=sev,material=material,code=code,a_lab='',a=None,b_lab='',b=None,diff=None,detail=detail,action=action))
    rows=list(maprows); have={(_norm(r['Customer']),str(r['Code'] or '').strip()) for r in rows}
    for b in (DS['blocks'] if add_new else []):
        if not b.get('sup') or b.get('count_stock'): continue
        for cu,v in (b.get('cust_plan') or {}).items():
            if sum(v)>0 and (_norm(cu),b['mat']) not in have:
                rows.append({'Customer':cu,'Material':b['common'],'Code':b['mat'],'Supplier':None,'Share %':None})
                have.add((_norm(cu),b['mat']))
    bycode={}
    for r in rows: bycode.setdefault(str(r['Code'] or '').strip(),[]).append(r)
    for b in DS['blocks']:
        mr=bycode.get(b['mat'])
        if not mr or b.get('count_stock'): continue
        cp={_norm(k):(k,v) for k,v in (b.get('cust_plan') or {}).items()}
        sup=b.setdefault('sup',[]); byname={s_['name'].upper():s_ for s_ in sup}
        share={}; mapped=set()
        for r in mr:
            sname=str(r['Supplier'] or '').strip().upper(); cu=_norm(r['Customer'])
            if not sname: continue
            try: pct=float(r.get('Share %')) if r.get('Share %') not in (None,'') else 100.0
            except (TypeError,ValueError): pct=100.0
            share.setdefault(cu,{}); share[cu][sname]=share[cu].get(sname,0)+pct; mapped.add(cu)
        for cu,ss in share.items():
            if abs(sum(ss.values())-100)>0.5:
                add('Medium','Customer Supplier shares do not add up to 100%',b['common'],b['mat'],f"{cp.get(cu,(cu,))[0]}: "+', '.join(f'{k} {v:g}%' for k,v in ss.items()),'Correct Share % on the Customer Supplier sheet')
        for cu,(name,v) in cp.items():
            if sum(v[:30])>0 and cu not in mapped:
                add('Medium','Customer has plan but no supplier assigned',b['common'],b['mat'],f'{name}: {sum(v[:30]):,.0f} in the next 30 days not linked to a supplier',
                    'Add / fill this customer on the Customer Supplier sheet of your MRP workbook')
        for cu,ss in share.items():
            if cu not in cp: continue
            name,v=cp[cu]; tot=sum(ss.values()) or 100
            for sname,pct in ss.items():
                s_=byname.get(sname)
                if s_ is None:
                    s_=dict(name=sname,stock=0.0,transit=[0.0]*nd,openpo=[0.0]*nd); sup.append(s_); byname[sname]=s_
                    add('Low','Supplier on the map has no stock or shipment',b['common'],b['mat'],f'{sname} is assigned to {name} but has no stock, in-transit or open PO for this material','Check the supplier name, or plan a PO')
                s_.setdefault('plan',[0.0]*nd); s_.setdefault('custs',[])
                f=pct/tot
                for k in range(nd): s_['plan'][k]+=v[k]*f
                s_['custs'].append(f'{name}' if pct>=99.5 else f'{name} {pct:g}%')
        for s_ in sup:
            if 'plan' in s_: s_['cust_text']=' | '.join(s_['custs'])
    return rows,issues


def make_supplier_cover_sheet(DS, path, TD=None):
    from openpyxl.styles import Border, Side
    from openpyxl.worksheet.table import Table, TableStyleInfo
    nd=len(DS['dates']); d0=pd.Timestamp(DS['dates'][0]).normalize(); AR='Calibri'; NAVY='1F3864'
    thin=Side(style='thin',color='BFBFBF'); bd=Border(left=thin,right=thin,top=thin,bottom=thin)
    wb=Workbook(); ws=wb.active; ws.title='Supplier Coverage'
    ws['A1']=f'Supplier coverage - each supplier\'s stock and shipments against the plan of the customers mapped to it'; ws['A1'].font=Font(name=AR,size=14,bold=True,color=NAVY)
    ws['A2']=('From Settings ALLOWED VENDORS (which supplier each customer may use). Run-out date = first day that supplier\'s projected stock (stock + in transit) goes below zero. '
              'Status = CRITICAL / AT RISK / WATCH / OK against the material\'s own target days (Settings). Red run-out date = inside that target. Stock by supplier = the PPC Count by supplier when counted, else MB52 batches (Batch Master) / MIS RM Status by Make (JW at the OEM left out); "Unassigned" stock is not counted for any supplier.')
    ws['A2'].font=Font(name=AR,size=9,italic=True); ws.merge_cells('A2:M2'); ws['A2'].alignment=Alignment(wrap_text=True); ws.row_dimensions[2].height=28
    H=['Material','Supplier','Customers mapped','Stock today','Avg plan / day (next target days)','Run-out date (with in transit)','In transit (total)','Open PO (total)','Next arrival','Status']
    for j,h in enumerate(H,1):
        c=ws.cell(4,j,h); c.font=Font(name=AR,size=9,bold=True,color='FFFFFF'); c.fill=PatternFill('solid',fgColor=NAVY); c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True); c.border=bd
    ws.row_dimensions[4].height=36; r=5; n=0
    for b in sorted(DS['blocks'],key=lambda b_:rrank(b_['mat'])):
        for sp in b.get('sup') or []:
            if 'plan' not in sp: continue
            TDm=b.get('td') or TD or LEAD_TIME_RM
            cov=cover_days(sp['stock'],sp['plan'],sp['transit'],sp['openpo']); avg=sum(sp['plan'][:TDm])/TDm
            k=next((i for i in range(nd) if sp['transit'][i]>0),None)
            st=material_status(cov,TDm)
            vals=[b['common'],sp['name'],sp.get('cust_text'),round(sp['stock']),round(avg),
                  (d0+pd.Timedelta(days=cov[1])).to_pydatetime() if cov[1]<nd else None,round(sum(sp['transit'])),round(sum(sp['openpo'])),
                  (f"{(d0+pd.Timedelta(days=k)):%d-%b}: {sp['transit'][k]:,.0f}" if k is not None else None),st]
            for j,v in enumerate(vals,1):
                c=ws.cell(r,j,v); c.font=Font(name=AR,size=9); c.border=bd; c.alignment=Alignment(vertical='top',wrap_text=j==3)
                if j in (4,5,7,8): c.number_format='#,##0'
                if j==6: c.number_format='dd-mmm-yy'
            if cov[1]<TDm: ws.cell(r,6).fill=PatternFill('solid',fgColor='FFC7CE')
            ws.cell(r,10).fill=PatternFill('solid',fgColor=STATUS_FILL[st])
            r+=1; n+=1
    if n:
        ws.add_table(Table(displayName='SupplierCoverage',ref=f'A4:J{r-1}',tableStyleInfo=TableStyleInfo(name='TableStyleLight1',showRowStripes=True)))
    else: ws.cell(5,1,'No supplier mapping yet - fill the Customer Supplier sheet').font=Font(name=AR,size=10,italic=True)
    for j,w in enumerate([22,14,44,12,12,13,12,12,18,11],1): ws.column_dimensions[L(j)].width=w
    ws.freeze_panes='C5'; wb.save(path); return n

# ===================== Module Plan from YOUR Production Plan (columns A-F) =====================
_STOP={'solar','farm','llc','inc','the','encap','ltd','co','project','g12r','m10r','energy'}
def _toks(s): return {t for t in re.split(r'[^a-z0-9]+',str(s or '').lower()) if t and t not in _STOP}
MODEL2BN={}
LINKS={}; LINK_OVERRIDE={}; LINK_OVR_NAME={}   # LINK_OVR_NAME: normalised -> customer name as typed on Settings MIS LINK
def module_plan_from_pp(mrp_path, mis_path, plant, dates):
    """Modules per day from your Production Plan (A Date, B Production Line, C Wp, D Customer, E Qty, F MW).
       Each customer is linked to the MIS customer / MIS BOM Name it matches (same name, or shared words such as 'Lucky 7'),
       so the BOM sheet and the Customer Supplier sheet keep working. Returns (modplan, sheet rows, issues) or None."""
    try: wb=wb_ro(mrp_path)
    except Exception as _e_: _swallow(6657,_e_); return None
    if 'Production Plan' not in wb.sheetnames: return None
    idx={pd.Timestamp(d).normalize():k for k,d in enumerate(dates)}; nd=len(dates)
    # MIS customer -> BOM Name (most modules)
    mw=wb_ro(mis_path); cb={}; mis_day={}; seen_=set()
    for r in mw['Total Plan Req'].iter_rows(min_row=2,values_only=True):
        if len(r)>21 and str(r[21] or '').strip().upper()==plant.upper() and r[2] and r[4] and isinstance(r[3],(int,float)):
            cu=' '.join(str(r[2]).split()); cb.setdefault(cu,{}); cb[cu][str(r[4]).strip()]=cb[cu].get(str(r[4]).strip(),0)+r[3]
            if isinstance(r[0],(dt.datetime,dt.date)):
                kk=(r[0],r[1],r[2],r[4],r[17] if len(r)>17 else None)
                if kk not in seen_:                         # one count per MIS plan line
                    seen_.add(kk); mis_day.setdefault((pd.Timestamp(r[0]).normalize(),_line_key(r[1])),[]).append((cu,str(r[4]).strip(),r[3]))
    mis_bom={cu:max(d,key=d.get) for cu,d in cb.items()}
    def link(cu):
        n=' '.join(str(cu or '').split())
        for m in mis_bom:
            if m.lower()==n.lower(): return m
        t=_toks(n); best=None; sc=0
        for m in mis_bom:
            o=len(t & _toks(m))
            if o>sc: best,sc=m,o
        return best if sc>0 else None
    lines={}; mws={}; modplan={}; issues=[]
    models={}
    for r in pp_rows_std(wb['Production Plan']):
        if not r or not isinstance(r[0],(dt.datetime,dt.date)) or not isinstance(r[4],(int,float)) or not r[4]: continue
        k=idx.get(pd.Timestamp(r[0]).normalize())
        if k is None: continue
        ln=' '.join(str(r[1] or '').split()); cu=' '.join(str(r[3] or '').split())
        mdl=str(r[7] or '').strip() if len(r)>7 else ''
        if mdl and mdl not in ('0','#N/A'): models.setdefault(cu,{}); models[cu][mdl]=models[cu].get(mdl,0)+r[4]
        if PLAN_WSA_LINES_ONLY and not (ln.upper().startswith('USA') or re.fullmatch(r'L\d+',ln.upper())): continue
        key=(ln,cu); lines.setdefault(key,[0.0]*nd)[k]+=r[4]
        mw_=r[5] if len(r)>5 and isinstance(r[5],(int,float)) and not isinstance(r[5],bool) else (r[4]*r[2]/1e6 if isinstance(r[2],(int,float)) else 0.0)
        mws.setdefault(key,[0.0]*nd)[k]+=mw_                  # MW of the plan line (for the WSA 'MW per day' column)
    # LINK (most reliable first):  1) your override on the Settings sheet;  2) the two plans themselves - the MIS plan line
    # on the SAME DAY and the SAME PRODUCTION LINE (no names needed);  3) customer name;  4) Module Model of another customer
    d_of={k:d for d,k in idx.items()}; votes={}
    for (ln,cu),v in lines.items():
        lk=_line_key(ln)
        for k,q in enumerate(v):
            if not q: continue
            for mc_,bn_,mq in mis_day.get((d_of[k],lk),[]):
                w=min(q,mq)*(1.5 if (_toks(cu)&_toks(mc_)) else 1.0)
                votes.setdefault(cu,{}); votes[cu][(mc_,bn_)]=votes[cu].get((mc_,bn_),0)+w
    LINKS.clear()
    for cu in {c for (_,c) in lines}:
        ov=link_override(cu)
        if ov and ov[1]: LINKS[cu]=dict(mis=ov[0] or link(cu),bom=ov[1],how='your override (Settings)',conf=1.0); continue
        vt=votes.get(cu)
        if vt:
            (mc_,bn_),w=max(vt.items(),key=lambda kv:kv[1]); tot=sum(vt.values())
            LINKS[cu]=dict(mis=mc_,bom=bn_,how='same day + same line in the MIS plan',conf=w/tot); continue
        mc_=link(cu)
        if mc_: LINKS[cu]=dict(mis=mc_,bom=mis_bom.get(mc_),how='customer name',conf=None); continue
        LINKS[cu]=dict(mis=None,bom=None,how='not found',conf=None)
    for cu,L_ in LINKS.items():                           # 4) still nothing: a customer with the same Module Model
        if L_['bom'] is None and models.get(cu):
            mdl=max(models[cu].items(),key=lambda kv:kv[1])[0]
            for c2,L2 in LINKS.items():
                if c2!=cu and L2['bom'] and mdl in (models.get(c2) or {}): LINKS[cu]=dict(mis=L2['mis'],bom=L2['bom'],how=f'same Module Model as {c2}',conf=None); break
    rows=[]
    for (ln,cu),v in lines.items():
        mc=LINKS[cu]['mis']; bn=LINKS[cu]['bom']
        mdl_=max((models.get(cu) or {}).items(),key=lambda kv:kv[1])[0] if models.get(cu) else ''
        rows.append(dict(line=ln,cust=cu,mis_cust=mc,bom=bn,v=v,model=mdl_,mw=mws.get((ln,cu))))
        if bn: arr=modplan.setdefault(bn,{}).setdefault(mc,[0.0]*nd); [arr.__setitem__(k,arr[k]+v[k]) for k in range(nd)]
        else:
            issues.append(dict(area='Plan',check='Production Plan customer not linked to a MIS BOM Name',sev='High',material=cu,code=ln,a_lab='',a=None,
                               b_lab='Modules in horizon',b=sum(v),diff=None,detail=f'{ln} / {cu}: not found in the MIS plan (same day / line), by name or by Module Model - these modules get no raw-material plan',
                               action='Set its MIS BOM Name in Settings - MIS LINK'))
    for cu,L_ in LINKS.items():
        if L_['bom'] and str(L_['how']).startswith('same day') and not (_toks(cu)&_toks(L_['mis'] or '')) and not (L_['conf'] is not None and L_['conf']<0.6):
            issues.append(dict(area='Master data',check='Production Plan customer name differs from MIS - MIS name used',sev='Low',material=cu,code=L_['bom'],a_lab='Confidence',a=round(L_['conf'] or 0,2),
                               b_lab='',b=None,diff=None,detail=f"Your Production Plan says '{cu}', MIS says '{L_['mis']}' (same days / lines) - the MRP shows the MIS name",
                               action="Use the MIS name on your Production Plan when convenient"))
        if L_['bom'] and L_['conf'] is not None and L_['conf']<0.6:
            issues.append(dict(area='Plan',check='Customer link to MIS is uncertain',sev='Medium' if L_['conf']<0.4 else 'Low',material=cu,code=L_['bom'],a_lab='Confidence',a=round(L_['conf'],2),b_lab='',b=None,diff=None,
                               detail=f"{cu} -> {L_['mis']} / {L_['bom']} ({L_['how']}, {L_['conf']:.0%} of matching days)",action='Check it, or set it in Settings - MIS LINK'))
    rows.sort(key=lambda x:(str(x['bom']),x['line'],x['cust']))
    m2b={}                                                  # Module Model -> MIS BOM Name (most modules)
    for x in rows:
        for mdl,q in (models.get(x['cust']) or {}).items():
            if x['bom']: m2b.setdefault(mdl,{}); m2b[mdl][x['bom']]=m2b[mdl].get(x['bom'],0)+q
    MODEL2BN.clear(); MODEL2BN.update({m:max(d,key=d.get) for m,d in m2b.items()})
    return modplan,rows,issues

def line_plan(pp_rows):
    """{MIS BOM Name: {MIS customer: {production line: modules per day}}} from the Module Plan rows."""
    lp={}
    for x in pp_rows or []:
        if not x.get('bom') or not x.get('mis_cust'): continue
        d=lp.setdefault(x['bom'],{}).setdefault(x['mis_cust'],{})
        arr=d.setdefault(x['line'],[0.0]*len(x['v']))
        for k,v in enumerate(x['v']): arr[k]+=v
    return lp

def module_plan_sheet(modplan, dates, path, pp_rows=None, src=None):
    wb=Workbook(); ws=wb.active; ws.title='Module Plan'
    if pp_rows is not None:
        ppc_=os.path.basename(PPC_SRC[0]) if PPC_SRC[0] and _has_ppc_plan(_sheetnames(PPC_SRC[0])) else None
        ws['A1']=((f"Modules per day from your PPC file ({ppc_}: sheet 'Plan' or 'WSA Plan', with sheet 'Revised Plan' on top) by Production Line and Customer. " if ppc_
                   else f'Modules per day from your Production Plan ({src}, columns A-F) by Production Line and Customer. ')
                  +'MIS Customer / MIS BOM Name = the MIS customer and module BOM the plan customer is linked to (Settings MIS LINK, else the closest name). '
                  +(f"Quantity per module = your PPC 'BOM' sheet; customers it does not list use the MIS Standard BOM. " if ppc_ else 'Quantity per module = your BOM sheet, else the MIS Standard BOM. ')
                  +'Which supplier a customer may use = Settings ALLOWED VENDORS.')
        hdr=['Production Line','Customer','MIS Customer','MIS BOM Name','Total']
        data=[[x['line'],x['cust'],x['mis_cust'],x['bom'],x['v']] for x in pp_rows]
    else:
        ws['A1']='Modules per day by MIS BOM Name and customer (from MIS Total Plan Req) - used with the BOM sheet for the CS / PM plan.'
        hdr=['MIS BOM Name','Customer','Total']
        data=[[bn,cu,v] for bn in sorted(modplan) for cu,v in sorted(modplan[bn].items(),key=lambda kv:-sum(kv[1]))]
    ws['A1'].font=Font(name='Arial',size=10,bold=True,color='1F3864')
    nfix=len(hdr); hdr=hdr+[pd.Timestamp(d).to_pydatetime() for d in dates]
    for j,h in enumerate(hdr,1):
        c=ws.cell(3,j,h); c.font=Font(name='Arial',size=9,bold=True,color='FFFFFF'); c.fill=PatternFill('solid',fgColor='1F3864')
        if j>nfix: c.number_format='dd-mmm'
    for i,row in enumerate(data):
        v=row[-1]; r=4+i
        for j,x in enumerate(row[:-1],1):
            c=ws.cell(r,j,x)
            if x is None and j>=nfix-1: c.fill=PatternFill('solid',fgColor='FFC7CE')
        ws.cell(r,nfix,sum(v)).number_format='#,##0'
        for k,x in enumerate(v): ws.cell(r,nfix+1+k,x or None)
    for j,w in enumerate([18,26,24,26,11][:nfix],1): ws.column_dimensions[L(j)].width=w
    ws.freeze_panes=ws.cell(4,nfix+1).coordinate; wb.save(path)

