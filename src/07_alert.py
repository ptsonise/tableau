# ===================== ALERT sheet: ONE simple list for both daily outputs =====================
ALERT_ROWS=[]      # built once in build_new_mrp: the first sheet of 1_MRP AND of the Inbound file (WSA daily) - the same rows
ALERT_D0=[None]    # the date of those rows (MRP day 1 = today)
ALERT_CLS=[('STOCK-OUT IN TRANSIT LEAD TIME','C00000','FFFFFF','EXPEDITE / use an alternative'),
           ('STOCK-OUT LATER','E26B0A','FFFFFF','order by the date'),
           ('UNDER LEAD TIME','FFC000','000000','in transit covers - watch the arrivals'),
           (f'UNDER LEAD TIME LATER','2E75B6','FFFFFF',f'falls below its lead time within {HZW()} - expedite / order before the date')]
ALERT_CLS+=[('OK','1E8449','FFFFFF','cover above the lead time - no action'),('NO DEMAND','8A949E','FFFFFF','no demand in the plan')]   # every material is listed (user, Nov 2026)
ALERT_SHORT=['Stock-out in LT','Stock-out later','Under lead time','Under lead time later','OK','No demand']
def alert_why(S, need, d0, lt, ships_=None, T=None, sup=None):
    """The 'why' columns of the ALERT sheet (were the Risk Charts sheet, merged into ALERT Nov 2026, user request): Below lead time
       from = first day projected stock covers < its lead time (days of need) (next 16 weeks); Cause = LATE ARRIVAL (what is on the way
       would keep lead time or more, but arrives after that day) / LATE + NOT ENOUGH / NO SUPPLY; Days late = first arrival after that day minus
       that day; Action detail = the Risk Charts action. ships_ = PPC In_Transit shipments (cover_watch result 'cw' style), else arrivals T per day."""
    TGT=tgt_days(lt); HZ=min(len(S),ALERT_HORIZON)
    if not need or not S: return dict(t=None,below=None,cause=None,late=None,detail='')
    t=next((k for k in range(HZ) if S[k]/need<TGT),None)
    if t is None: return dict(t=None,below=None,cause=None,late=None,detail='')
    dt_=d0+pd.Timedelta(days=t)
    if ships_ is not None: later=[(s_['raw'],s_['qty'],s_.get('sup')) for s_ in ships_]
    else: later=[(d0+pd.Timedelta(days=k),T[k],sup) for k in range(t+1,min(len(T),len(S))) if T[k]>0]
    later=sorted([l_ for l_ in later if l_[0]>dt_],key=lambda l_:l_[0])
    gap=TGT*need-S[t]; on_way=sum(q for _,q,_ in later); gap_all=max(0.0,TGT*need-min(S[t:HZ]))
    cause='NO SUPPLY' if not later else ('LATE ARRIVAL' if on_way>=gap else 'LATE + NOT ENOUGH')
    nx=later[0] if later else None
    if cause=='LATE ARRIVAL': det=f"Expedite {nx[2] or 'shipment'} to {dt_:%d-%b} (due {nx[0]:%d-%b})"+(f" + order {gap_all:,.0f} for later weeks" if gap_all>0 else '')
    elif cause=='LATE + NOT ENOUGH': det=f"Expedite {nx[2] or ''} + order {gap_all:,.0f} more".replace('  ',' ')
    else: det=f"Order {gap_all:,.0f} (nothing on the way after {dt_:%d-%b})"
    return dict(t=t,below=dt_,cause=cause,late=(nx[0]-dt_).days if nx else None,detail=det)
def floor_path(S, st0):
    """Stock NEVER below 0 (user 23-Nov-2026: an arrival cannot cover days already lost): projected stock S (carries the shortage) -> (F, Sf):
       F = stock after every day floored at 0, Sf = floored stock + that day's net flow BEFORE the floor (< 0 = a stock-zero day)."""
    F=[]; Sf=[]; sp=float(st0 or 0.0); prv=max(0.0,sp)
    for v in S:
        x=prv+(v-sp); Sf.append(x); prv=max(0.0,x); F.append(prv); sp=v
    return F,Sf
def floor_wins(S, st0):
    """Stock-zero periods (first day, last day, still open at the end) from the floored path."""
    Sf=floor_path(S,st0)[1]; nd=len(Sf); wins=[]; k=0
    while k<nd:
        if Sf[k]<0:
            k2=k
            while k2+1<nd and Sf[k2+1]<0: k2+=1
            wins.append((k,k2,k2+1>=nd)); k=k2+1
        else: k+=1
    return wins
def wins_txt(wins, d0, lt):
    """'Stock zero from → to' text of floor_wins periods - ONE rule (ALERT rows + Settings ALLOWED VENDORS supplier sub-rows). '' = none."""
    D=lambda k_: d0+pd.Timedelta(days=k_)
    def ptxt(a,z,op):
        t_=f"{D(a):%d-%b} → beyond {D(z):%d-%b}" if op else (f"{D(a):%d-%b} (1 d)" if a==z else f"{D(a):%d-%b} → {D(z):%d-%b} ({z-a+1} d)")
        return ('⚠ ' if a<int(lt or 0) else '')+t_
    return '\n'.join(ptxt(*w) for w in (wins or [])[:6])+(f"\n... +{len(wins)-6} more" if len(wins or [])>6 else '')
def vp_subs(r):
    """Settings ALLOWED VENDORS supplier sub-rows of an ALERT row (user 28-Nov-2026: "ALERT, Material Status, 4_Status, 5_Web: a material with
       Settings ALLOWED VENDORS rules gets one sub-row per supplier under the material row (= Action List supplier rows). Numbers come from
       vendor_pools, no new calculation"). = the Action List supplier rows of the material (AL_ORDER_ 'sup_row', built in build_new_mrp from
       vendor_pools: own stock, own PPC arrivals, own projection). Presentation only: no days of cover (no Usage/Day per supplier).
       [] for a row per WSA sheet (333: that row is already one supplier) or a material with one supplier / no rule."""
    m=str(r.get('code') or '').strip()
    if r.get('where') or m not in VP_MRP: return []
    d0=ALERT_D0[0]; out=[]
    for px in AL_ORDER_ or []:
        if not px.get('sup_row') or px['b']['mat']!=m: continue
        nc=bool(px.get('no_cust'))
        ata=sup_ata(px,SHIPS_,d0)                                        # = the Action List 'Arrival at plant (Final AtA)' of the supplier row
        wins=px.get('wins') or []
        out.append(dict(sup=str(px['sup_row']),custs=('no customer may use this vendor' if nc else str(px.get('custs') or '')),nc=nc,
                        st=float(px.get('own_stock') or 0.0),it=float(px.get('tr_tot') or 0.0),
                        per=('not counted: no customer demand yet' if nc else (wins_txt(wins,d0,r.get('lt')) if d0 is not None else '')),
                        first=(d0+pd.Timedelta(days=wins[0][0]) if (wins and not nc and d0 is not None) else None),ata=ata,
                        nxl='next: '+(' | '.join(x_.strip() for x_ in str(ata).split('\n') if x_.strip()) if ata else 'none')))
    return out
def sup_ata(x, ships, d0):
    """'Arrival at plant (Final AtA)' of an Action List row (material or supplier row) - ONE rule for the Action List and the supplier sub-rows.
       A supplier no customer with demand may use (312 FLAT): its arrivals are NOT days of cover (show it, don't count it, user 28-Nov-2026)
       -> the `vp_nc` line 'Not counted (no customer demand yet: ...): in transit N (dd-Mmm → dd-Mmm)'. Else `ata_show` (days of cover)."""
    if x.get('sup_row') and x.get('no_cust'):
        nc=vp_nc(x['b']['mat'],ships,d0)
        if nc and nc.get('txt_it'): return nc['txt_it']
    return ata_show(x.get('ata_txt'),[x['b']['mat']],x.get('pl_need'))
def vp_sub_name(s):
    """Item name of a supplier sub-row: '↳ <supplier> · <customers allowed>'."""
    return f"   ↳ {s['sup']} · {s['custs']}"
def alert_eval(name, code, uom, stock, S, T, need, lt, tr_tot, d0, where=None, sup=None, why=None, ata=None):
    """ONE rule for the ALERT sheet of both daily outputs (user request Nov 2026: 'list of less than 21 days inventory, how many days
       stock zero within transit lead time and zero from - to; the other sheets are too complicated - one simple sheet, strong alert'):
       days of stock = stock / Usage/Day (module plan, average per production day over the transit lead time);
       stock-zero day = that day's need is not covered: stock + in transit arrived up to that day - plan up to that day < 0
       (= the Action List 'Short from'; on a WSA sheet: Plant + Port Days (PPC In Transit) below 1 - the red cell comes a day later);
       stock-zero days inside the transit lead time = such days from today until today + transit lead time - 1.
       Listed when days of stock < the lead time or a stock-zero day falls inside the transit lead time; None = not listed."""
    if not need:                                                  # no demand: still listed (class 5), no days of stock
        nx=next(((k_,T[k_]) for k_ in range(min(len(S),len(T or []))) if T[k_]>0),None)      # its own next arrival (a substitute without own plan, EPE 480)
        return dict(cls=5,code=code,name=name,where=where,uom=str(uom or ''),stock=stock,need=0.0,days=None,days_it=None,days_nx=None,lt=int(lt or 0),zin=0,
                    first=None,first_in=False,per='none in the plan',nx=(d0+pd.Timedelta(days=nx[0]),nx[1]) if nx else None,sup=sup,act='No demand in the plan',k0=None,
                    why=dict(t=None,below=None,cause=None,late=None,detail=''),ata=ata,S=list(S),tr_tot=tr_tot)
    nd=len(S); lt=int(lt or 0); wins=floor_wins(S,stock)                  # stock never below 0 (user 23-Nov-2026)
    zin=sum(min(z,lt-1)-a+1 for a,z,_ in wins if a<lt)
    days=max(0.0,stock)/need; days_it=max(0.0,stock+(tr_tot or 0.0))/need
    why=why or dict(t=None,below=None,cause=None,late=None,detail='')
    D=lambda k_: d0+pd.Timedelta(days=k_)
    per=wins_txt(wins,d0,lt)
    k0=wins[0][0] if wins else None; cls=0 if zin>0 else (1 if wins else (2 if days<tgt_days(lt) else (3 if why['t'] is not None else 4)))
    if cls==0: act='EXPEDITE / use an alternative'
    elif cls==1: act='ORDER NOW' if k0-lt<=RELEASE_WINDOW else f"ORDER BY {D(k0-lt):%d-%b}"
    elif cls==2: act='Watch the arrivals (in transit covers the plan)'
    elif cls==3: act=f"Keep {tgt_days(lt)} days (lead time): act before {why['below']:%d-%b}"
    else: act='OK - no action'
    if why.get('detail'): act+='\n'+why['detail']
    nx=next(((k_,T[k_]) for k_ in range(min(nd,len(T))) if T[k_]>0),None)
    days_nx=max(0.0,stock+(nx[1] if nx else 0.0))/need            # Inv. Days incl. the NEXT arrival only (web pages + chat pictures, user 21-Nov-2026); presentation only
    return dict(cls=cls,code=code,name=name,where=where,uom=str(uom or ''),stock=stock,need=need,days=days,days_it=days_it,days_nx=days_nx,lt=lt,zin=zin,
                first=D(k0) if k0 is not None else None,first_in=(k0 is not None and k0<lt),per=per or 'none in the plan',
                nx=(D(nx[0]),nx[1]) if nx else None,sup=sup,act=act,k0=k0,why=why,ata=ata,S=list(S),tr_tot=tr_tot)
def build_alert_rows(A, d0, nd):
    """The ALERT list, built ONCE (1_MRP first tab = Inbound first tab). One row per material with demand from the Action List
       projection (stock + PPC in transit by arrival date - module plan; a Settings substitute = its own row, own numbers, right under its material - sub_cover). A material that the
       WSA daily splits by customer project and supplier (on 2+ WSA sheets: back glass 333 = SB Back + Subanci Back) gets one row
       per sheet from that sheet's stock, demand and PPC arrivals - the same rule on the numbers of the sheet."""
    on={}
    for g in WSA_REFS[0] or []:
        for c in g.get('codes') or []: on.setdefault(c,[]).append(g)
    dts=[(d0+pd.Timedelta(days=k)).to_pydatetime() for k in range(nd)]; out=[]
    for x in A:
        b=x['b']; m=b['mat']
        lt=b.get('lt_used')
        shs=[g for g in on.get(m,[]) if g.get('dem_by_date')]
        if len(shs)>1:                                       # one row per customer project (WSA sheet)
            for g in shs:
                bb=g.get('bal_by_date') or {}; st=bb.get(dts[0])
                if st is None: st=bb.get(g['asof'])
                st=float(st or 0.0)
                P=[float((g.get('dem_by_date') or {}).get(d) or 0.0) for d in dts]; T=[float((g.get('ppcit_by_date') or {}).get(d) or 0.0) for d in dts]
                cur=st; S=[]; VL_=[float((g.get('vlock_by_date') or {}).get(d) or 0.0) for d in dts]   # Settings ALLOWED VENDORS (WSA column AK)
                for k in range(nd): cur=cur-P[k]+T[k]; S.append(cur-VL_[k])
                nd_=usage_day(P,lt_of(b))                     # Usage/Day over the transit lead time
                fut=[(dts[k],T[k]) for k in range(nd) if T[k]>0]
                ata='\n'.join('   |   '.join(f"{d_:%a %d-%b}: {q_:,.0f}" for d_,q_ in fut[i_:i_+2]) for i_ in range(0,min(len(fut),ATA_SHOW),2)) or None
                r=alert_eval(b.get('common') or m,m,b.get('uom'),st,S,T,nd_,lt,sum(T),d0,where=g['sheet'],why=alert_why(S,nd_,d0,lt,T=T),ata=ata)
                if r: r.update(P=list(P),Tp=list(T),vlock=(VL_ if any(VL_) else None)); out.append(r)
            continue
        nx=x.get('nx') or {}
        dst_,dT_,dn_=vp_dead(m)                                          # Settings ALLOWED VENDORS: stock / arrivals no customer may use are out
        sh_=[s_ for s_ in SHIPS_.get(m,[]) if s_.get('raw') is not None and not (dn_ and dn_(s_.get('sup')))]       # = cover_watch (Action List module-plan columns)
        T_al=[t_-(dT_[k_] if dT_ and k_<len(dT_) else 0.0) for k_,t_ in enumerate(b['transit'])]
        if dn_:                                                          # next-arrival supplier text: allowed suppliers only
            k1_=next((k_ for k_,t_ in enumerate(T_al) if t_>0.5),None)
            nx=dict(nx,sup=', '.join(dict.fromkeys(str(s_.get('sup')) for s_ in sh_ if s_.get('date') is not None and k1_ is not None
                                                    and (s_['date']-d0).days==k1_ and not dn_(s_.get('sup')))) or nx.get('sup'))
        r=alert_eval(b.get('common') or m,m,b.get('uom'),float(b['stock'] or 0.0)-dst_,x.get('S') or [],
                     T_al,(None if x.get('no_demand') else x.get('pl_need')),lt,float(x.get('tr_tot') or 0.0)-sum(dT_ or []),d0,sup=nx.get('sup'),
                     why=alert_why(x.get('S') or [],x.get('pl_need'),d0,lt,ships_=sh_),ata=x.get('ata_txt'))
        if r: r['nc']=vp_nc(m,SHIPS_,d0)                                  # shown, not counted (user 28-Nov-2026): r['stock'] / r['tr_tot'] stay the counted numbers
        if r:                                                # plan + arrivals of the projection S (ALERT (formulas) / ALERT Calc rebuild S from them)
            vl_=(VP_MRP.get(m) or {}).get('lock')                          # Settings ALLOWED VENDORS (already in x['S'])
            dc_=(VP_MRP.get(m) or {}).get('dead_cum') if (vl_ and dn_) else None     # dead stock / arrivals are already out of stock / Tp: only the rest is a lock
            if dc_: vl_=[max(0.0,v_-(dc_[k_] if k_<len(dc_) else 0.0)) for k_,v_ in enumerate(vl_)]
            r['vlock']=list(vl_)[:nd] if vl_ and any(v_>1e-6 for v_ in vl_) else None
            r.update(P=([0.0]*nd if x.get('no_demand') else list(b.get('plan') or [])[:nd]),
                     Tp=[a_+c_-(dT_[k_] if dT_ and k_<len(dT_) else 0.0) for k_,(a_,c_) in enumerate(zip(b['transit'],b.get('adhoc') or [0.0]*nd))][:nd])
            out.append(r)
    for r in out: r.update(alert_stock_fields(r,d0))
    try: alert_stock_src(out,A)                                       # SAP stock (MB52) / physical count next to Estimated stock today
    except Exception as ex_: import traceback; traceback.print_exc(); print(f'NOTE: ALERT SAP stock / physical count not filled ({type(ex_).__name__}: {ex_})')
    try: alert_lines(out,A,d0,nd)                                     # plan / arrival lines behind P / Tp (ALERT (formulas) SUMIFS)
    except Exception as ex_: import traceback; traceback.print_exc(); print(f'NOTE: ALERT lines not built ({type(ex_).__name__}: {ex_})')
    for c in {r['code'] for r in out if r.get('where')}:              # 333 per WSA sheet: on-hold stock (blocked + QI) sits with the stock -
        rs=[r for r in out if r['code']==c and r.get('where')]           # the sheet that holds the supplier's stock (by stock share; none: 1st)
        tot=sum(max(0.0,float(r['stock'] or 0)) for r in rs); h=rs[0]['hold']
        for i,r in enumerate(rs): r['hold']=(h*max(0.0,float(r['stock'] or 0))/tot) if tot>0 else (h if i==0 else 0.0)
    out.sort(key=lambda r:(r['cls'],r['k0'] if r['k0'] is not None else (r['why']['t'] or 0),r['days_it'] if r['days_it'] is not None else 1e9,str(r['name'])))   # high to low risk
    try: out=sub_cover(out,d0)                                        # Settings SUBSTITUTES: who covers when, how much; substitute row under its material
    except Exception as ex_: import traceback; traceback.print_exc(); print(f'NOTE: substitute cover not built ({type(ex_).__name__}: {ex_})')
    return out
def sub_cover(rows, d0):
    """Settings SUBSTITUTES between two MRP materials (user 25-Nov-2026: "show EPE 480 and EPE 420 SEPARATELY ... When calculate inv day must
       show when and what quantity cover by which material and its substitution code"). Every code keeps its OWN ALERT numbers (own stock,
       plan, PPC arrivals, class, dates). Here only, day by day with stock never below 0 (`floor_path` rule): the main material uses its own
       stock + arrivals first; on a day they do not cover, the substitute's stock + arrivals LEFT AFTER ITS OWN PLAN cover the rest
       (factor = main units per 1 substitute unit; Use from / Use until respected). Sets on the main row and on the substitute row:
       r['subcov'] = full text (days of cover per code, periods + quantities per code, stock zero with the substitute) and r['subnote'] =
       short note under the item name. The substitute row is moved right under its main material. Presentation only."""
    by={r['code']:r for r in rows if not r.get('where')}
    D=lambda k_: f"{(d0+pd.Timedelta(days=int(k_))):%d-%b}"
    q=lambda v_: f"{v_:,.0f}"
    def segs(use):
        out_=[]; k=0
        while k<len(use):
            if use[k]>1e-6:
                k2=k
                while k2+1<len(use) and use[k2+1]>1e-6: k2+=1
                out_.append((k,k2,sum(use[k:k2+1]))); k=k2+1
            else: k+=1
        return out_
    move=[]
    for m,subs in SUBG.items():
        rm=by.get(m); nd=len((rm or {}).get('P') or [])
        if not nd: continue
        Pm=list(rm['P']); Tm=(list(rm.get('Tp') or [])+[0.0]*nd)[:nd]; sm=max(0.0,float(rm['stock'] or 0)); um=rm.get('uom') or ''
        ss=[]
        for c,f,fr,to in subs:
            rs=by.get(c)
            if rs is None or not f: continue
            ss.append(dict(c=c,r=rs,f=float(f),fr=fr,to=to,st=max(0.0,float(rs['stock'] or 0)),P=(list(rs.get('P') or [])+[0.0]*nd)[:nd],
                           T=(list(rs.get('Tp') or [])+[0.0]*nd)[:nd],use=[0.0]*nd))
        if not ss: continue
        own0=None; zero=[]; mor=[]
        for k in range(nd):
            sm+=Tm[k]; mor.append((sm,[z['st']+z['T'][k] for z in ss])); take=min(sm,Pm[k]); sm-=take; rest=Pm[k]-take
            if rest>1e-6 and own0 is None: own0=k
            dk=d0+pd.Timedelta(days=k)
            for z in ss:
                z['st']=max(0.0,z['st']+z['T'][k]-z['P'][k])                     # the substitute's own plan first
                if rest<=1e-6 or (z['fr'] is not None and pd.Timestamp(z['fr'])>dk) or (z['to'] is not None and pd.Timestamp(z['to'])<dk): continue
                u=min(z['st']*z['f'],rest); z['st']-=u/z['f']; rest-=u; z['use'][k]=u/z['f']
            if rest>1e-6: zero.append(k)
        okw=lambda z,k: not ((z['fr'] is not None and pd.Timestamp(z['fr'])>d0+pd.Timedelta(days=k)) or (z['to'] is not None and pd.Timestamp(z['to'])<d0+pd.Timedelta(days=k)))
        last=max([k for k in range(nd) if Pm[k]>1e-9],default=-1); tl=float(rm.get('need') or 0.0); cvs=[]
        for k in range(nd):                     # cover INCL. the substitute (user 25-Nov-2026: "heat map also reflected"): same meaning as bom_day_cover
            if Pm[k]<=1e-9: cvs.append('na'); continue       # (days the stock lasts from that morning, no later arrivals), the substitute's stock after its own plan fills what own stock does not
            o_,st_=mor[k][0],list(mor[k][1]); v_=None
            for j in range(k,last+1):
                tk=min(o_,Pm[j]); o_-=tk; rest=Pm[j]-tk
                for i_,z in enumerate(ss):
                    st_[i_]=max(0.0,st_[i_]-z['P'][j])
                    if rest>1e-6 and okw(z,j): u=min(st_[i_]*z['f'],rest); st_[i_]-=u/z['f']; rest-=u
                if rest>1e-6: v_=(j-k)+(Pm[j]-rest)/Pm[j]; break
            if v_ is None: v_=(last+1-k)+((o_+sum(x_*z['f'] for x_,z in zip(st_,ss)))/tl if tl>1e-9 else 0.0)
            cvs.append(v_)
        rm['cv_sub']=cvs; rm['sub_zero']=list(zero)
        rm['sub_x']=[dict(c=z['c'],n=str(z['r']['name']),f=z['f'],u=z['r'].get('uom') or um,T=list(z['T']),use=list(z['use'])) for z in ss]
        for z in ss: z['r']['sub_give']=dict(c=m,n=str(rm['name']),use=list(z['use']))
        nm_m=f"{rm['name']} {m}"; need=rm.get('need') or 0.0
        dd=[f"own {max(0.0,float(rm['stock'] or 0))/need:,.1f}"]+[f"{z['r']['name']} {max(0.0,float(z['r']['stock'] or 0))*z['f']/need:,.1f}" for z in ss] if need else []
        tot_d=sum(max(0.0,float(x_['stock'] or 0))*f_ for x_,f_ in [(rm,1.0)]+[(z['r'],z['f']) for z in ss])/need if need else None
        main_t=[]; main_n=[]
        if need: main_t.append(f"{tot_d:,.1f} days of cover (stock only): {' + '.join(dd)}")   # first (user 25-Nov-2026 text order), then 'Then <sub> covers ...', then 'With <sub>: stock zero ...'
        for z in ss:
            sg=segs(z['use']); used=sum(z['use']); zn=f"{z['r']['name']} {z['c']}"; uz=z['r'].get('uom') or um
            arr=[(k_,v_) for k_,v_ in enumerate(z['T']) if v_>0.5]
            at=', '.join(f"{D(k_)} {q(v_)}" for k_,v_ in arr[:4])+(f" +{len(arr)-4} more" if len(arr)>4 else '')
            src=f"its stock {q(max(0.0,float(z['r']['stock'] or 0)))} {uz}"+(f" + arrivals {at}" if arr else ' (nothing in transit)')+(f", after its own plan {q(sum(z['P']))}" if sum(z['P'])>0.5 else '')
            if sg:
                pt='; '.join((f"{D(a_)} → {D(b_)}" if b_>a_ else D(a_))+f" {q(v_)} {uz}" for a_,b_,v_ in sg[:4])+(f"; +{len(sg)-4} more" if len(sg)>4 else '')
                main_t.append(f"then {zn} covers {pt} (total {q(used)} {uz}; {src})"); main_n.append(f"substitute {z['r']['name']} {z['c']} covers {D(sg[0][0])} → {D(sg[-1][1])}: {q(used)} {uz}")
                st_=(f"Substitute for {nm_m}: used {pt} (total {q(used)} {uz}; {src})")
                z['r']['subnote']=f"substitute for {rm['name']} {m}: used {D(sg[0][0])} → {D(sg[-1][1])}: {q(used)} {uz}"
            elif own0 is None:
                st_=f"Substitute for {nm_m}: not needed - {rm['name']} covers its plan ({src})"; z['r']['subnote']=f"substitute for {rm['name']}: not needed"
            else:
                st_=f"Substitute for {nm_m}: nothing left to give ({src})"; z['r']['subnote']=f"substitute for {rm['name']}: nothing left to give"
            z['r']['subcov']=st_
        zw=[]                                                            # stock-zero periods WITH the substitute (days without plan between two short days bridge them)
        for k in zero:
            if zw and all(Pm[j]<=1e-6 for j in range(zw[-1][1]+1,k)): zw[-1][1]=k
            else: zw.append([k,k])
        ztx=', '.join((f"{D(a_)} → beyond {D(b_)}" if b_>=nd-1 else (f"{D(a_)} → {D(b_)}" if b_>a_ else D(a_))) for a_,b_ in zw[:4])+(f" +{len(zw)-4} more" if len(zw)>4 else '')
        if own0 is not None and any(sum(z['use'])>0 for z in ss):
            zt=(f"with {' / '.join(z['r']['name'] for z in ss)}: stock zero {ztx}" if zero else f"with {' / '.join(z['r']['name'] for z in ss)}: no stock zero in the plan")
            main_t.append(zt); main_n.append(zt)
        elif own0 is not None: main_n.append(f"{' / '.join(z['r']['name'] for z in ss)} has nothing to give")
        else: main_n.append(f"own stock covers the plan; substitute {' / '.join(z['r']['name'] for z in ss)} not needed")
        rm['subcov']='. '.join(t_[0].upper()+t_[1:] for t_ in main_t)+'.'; rm['subnote']='; '.join(main_n)
        move+=[(m,z['c']) for z in ss]
    for m,c in move:                                                     # substitute row right under its material (user: "material and its substitutions")
        rs=by.get(c); rm=by.get(m)
        if rs is None or rm is None or rs not in rows: continue
        rows.remove(rs); rows.insert(rows.index(rm)+1,rs)
    return rows
def _subset_match(groups, target, tol):
    """Keys of `groups` (key -> list per day) whose sum equals `target` day by day (smallest subset first); None if none / too many."""
    import itertools
    ks=[k for k,v in groups.items() if any(abs(x)>1e-9 for x in v)]
    if len(ks)>16: return None
    tt=sum(target)
    for n_ in range(0,len(ks)+1):
        for sub in itertools.combinations(ks,n_):
            if abs(sum(sum(groups[k]) for k in sub)-tt)>tol*max(1.0,len(target)): continue
            if all(abs(sum(groups[k][d] for k in sub)-target[d])<=tol for d in range(len(target))): return list(sub)
    return None
def alert_lines(rows, A, d0, nd):
    """The lines behind each ALERT row's daily Plan (P) and Arrivals (Tp), for the 'ALERT Plan lines' / 'ALERT Arrival lines' sheets
       (user, Nov 2026: "I want to see formulas in ALERT (formulas): Plan (demand), Arrivals (in transit)"):
       plan line = customer x production line x day: modules x qty per module (your BOM / MIS BOM; MIS requirement = quantity), OEM out;
       arrival line = one PPC In_Transit shipment (traced code, alias factor), counted on max(arrival, today); + MIS Adhoc / substitute
       receipts per day. Back glass 333 per WSA sheet: the customers / POs whose sum
       is the sheet's Demand / column T. Every row is checked: the lines must add up to P / Tp day by day; when they do not, the row
       gets one line per day with the projection value (r['pl_ok'] / r['al_ok'] = False, console NOTE)."""
    byb={x['b']['mat']:x['b'] for x in A}; oem=OEM_IS[0] or (lambda cu: False)
    tol=lambda v: max(0.01,abs(v)*1e-6)
    def plines(c,fac):
        out=[]
        for pl in (byb.get(c) or {}).get('plan_lines') or []:
            if oem(pl['cu']): continue
            for k in range(min(nd,len(pl['mods']))):
                if pl['mods'][k]: out.append(dict(code=c,k=k,cu=pl['cu'],ln=pl['ln'],src=pl['src'],mods=float(pl['mods'][k]),rate=float(pl['rate']),f=fac))
        return out
    def alines(c,fac,mem=()):
        out=[]; sub_=(byb.get(c) or {}).get('adhoc_sub') or {}
        dn_=vp_dead(c)[2]                                              # Settings ALLOWED VENDORS: shipments no customer may use are not arrivals
        for s_ in SHIPS_.get(c,[]):
            if s_.get('raw') is None or (dn_ and dn_(s_.get('sup'))): continue
            out.append(dict(code=c,sup=s_.get('sup') or '',po=str(s_.get('po') or '').replace('.0',''),bl=str(s_.get('bl') or ''),raw=s_['raw'],qty=float(s_['qty']),f=fac,src='PPC In_Transit'))
        for k,v in enumerate(((byb.get(c) or {}).get('adhoc') or [])[:nd]):
            v-=sum((arr_[k] if k<len(arr_) else 0.0) for sc_,arr_ in sub_.items() if sc_ in mem)   # group: a member's own stock is not a receipt (= _adh)
            if abs(v)>1e-9: out.append(dict(code=c,sup='Adhoc / substitute receipt',po='',bl='',raw=d0+pd.Timedelta(days=k),qty=float(v),f=fac,src='MIS Adhoc sheet / Settings substitute (per day)'))
        return out
    def daysum(ls,qf,kf):
        t=[0.0]*nd
        for l_ in ls:
            k=kf(l_)
            if 0<=k<nd: t[k]+=qf(l_)
        return t
    pk=lambda l_: l_['k']; pq=lambda l_: l_['mods']*l_['rate']*l_['f']
    ak=lambda l_: max(0,(l_['raw']-d0).days); aq=lambda l_: l_['qty']*l_['f']
    bad=[]
    for r in rows:
        m=r['code']; P=list(r.get('P') or [0.0]*nd)[:nd]; T=list(r.get('Tp') or [0.0]*nd)[:nd]
        P+= [0.0]*(nd-len(P)); T+=[0.0]*(nd-len(T))
        if r.get('where'):                                               # 333 per WSA sheet: the customers / POs of this sheet
            pl_all=plines(m,1.0); al_all=alines(m,1.0)
            byc={}
            for l_ in pl_all: byc.setdefault(l_['cu'],[]).append(l_)
            sub=_subset_match({c_:daysum(v_,pq,pk) for c_,v_ in byc.items()},P,0.01)
            pls=[l_ for c_ in (sub or []) for l_ in byc[c_]] if sub is not None else None
            byp={}
            for l_ in al_all: byp.setdefault(l_['po'] or l_['src'],[]).append(l_)
            sub=_subset_match({c_:daysum(v_,aq,ak) for c_,v_ in byp.items()},T,0.01)
            als=[l_ for c_ in (sub or []) for l_ in byp[c_]] if sub is not None else None
        else:
            pls=plines(m,1.0); als=alines(m,1.0)
        if r.get('cls')==5: pls=[]                                        # no demand: no plan lines
        if pls is None or any(abs(a-b)>tol(b) for a,b in zip(daysum(pls,pq,pk),P)):
            bad.append(f"{m}{' ('+r['where']+')' if r.get('where') else ''} plan")
            pls=[dict(code=m,k=k,cu='(projection value - lines do not add up)',ln=None,src='ALERT projection',mods=P[k],rate=1.0,f=1.0) for k in range(nd) if P[k]]
        if als is None or any(abs(a-b)>tol(b) for a,b in zip(daysum(als,aq,ak),T)):
            bad.append(f"{m}{' ('+r['where']+')' if r.get('where') else ''} arrivals")
            als=[dict(code=m,sup='(projection value - lines do not add up)',po='',bl='',raw=d0+pd.Timedelta(days=k),qty=T[k],f=1.0,src='ALERT projection') for k in range(nd) if T[k]]
        r['plan_l']=pls; r['arr_l']=als
    n_p=sum(len(r['plan_l']) for r in rows); n_a=sum(len(r['arr_l']) for r in rows)
    print(f"ALERT (formulas): {n_p} plan lines, {n_a} arrival lines behind {len(rows)} ALERT rows - "+("every row adds up to the projection" if not bad else "NOT split (lines do not add up): "+'; '.join(bad)),flush=True)
HOLD_={}
def mb52_hold(mis_path):
    """On hold per MRP code = MIS MB52 Blocked (column J) + Quality Inspection (column M) - the old Material Status rule."""
    if HOLD_.get('_read') or not mis_path: return HOLD_
    HOLD_['_read']=True
    try:
        mis=wb_ro(mis_path)
        for r in list(mis['MB52'].iter_rows(values_only=True))[1:]:
            if r and len(r)>14 and r[2]:
                c=str(r[2]).strip(); m=(BOM_ALIAS[0] or {}).get(c,c)
                HOLD_[m]=HOLD_.get(m,0)+(r[12] if isinstance(r[12],(int,float)) else 0)+(r[9] if isinstance(r[9],(int,float)) else 0)
    except Exception as ex_: print(f'NOTE: MB52 on-hold stock not read ({type(ex_).__name__}: {ex_})')
    return HOLD_
def alert_stock_fields(r, d0):
    """The Material Status columns of an ALERT row (user, Nov 2026: 'combine alert and stock management into 1 sheet'):
       in transit (= what the row's projection uses), overdue in transit (PPC arrival date passed, not received), on hold
       (MB52 blocked + QI), reorder level = Usage/Day x (transit lead time + SAFETY_DAYS), status Out of Stock / Below reorder /
       Above reorder ('Short dd-mmm' when above reorder but a stock-zero day is in the plan - never labelled safe when short).
       333 per WSA sheet: its PPC arrivals (late ones are on today)."""
    m=r['code']; hold=mb52_hold(MIS_PATH_[0])
    hq=float(hold.get(m) or 0)
    ovd_m=sum(s_['qty'] for s_ in SHIPS_.get(m,[]) if s_.get('raw') is not None and s_['raw']<d0)     # material total, shown (physical: no-demand supplier included, user 28-Nov-2026)
    ovd=None if r.get('where') else ovd_m                             # per WSA sheet: overdue lines are in today's PPC arrivals
    it=float(r.get('tr_tot') or 0.0)+float((r.get('nc') or {}).get('it') or 0.0)    # In transit SHOWN = physical; 'Days of cover (+ all in transit)' uses tr_tot (counted)
    rop=(r['need'] or 0.0)*(int(r['lt'] or 0)+tgt_days(r['lt'])); st=float(r['stock'] or 0.0)   # safety stock = lead-time days of need
    status='Out of Stock' if st<=0 else ('Below reorder' if st<rop else 'Above reorder')
    if status=='Above reorder' and r.get('first') is not None: status=f"Short {r['first']:%d-%b}"
    return dict(it=it,ovd=ovd,ovd_m=ovd_m,hold=hq,rop=rop,status=status)
def stock_src(b):
    """SAP stock (MB52), physical count and estimated stock today of a material - ONE rule for Stock Check, Material Status and the
       ALERT (user, Nov 2026: "Estimated stock today in ALERT = same calc logic as Stock Check, then SAP stock (MB52), Physical count").
       sap = MB52 stock (Unrestricted + QI; not counted and no MB52 line: the stock the MRP uses) · cnt = PPC physical count as counted
       (None = not counted) · est = count rolled forward to today (Stock Check rule), else the stock the MRP uses (= SAP) · own = the
       stock every shortage date uses (JW at the OEM out). est = own for every material (same number on every sheet)."""
    m=b['mat']; e=STOCK_EST.get(m) or {}
    own=float(b['stock'])-(0.0 if b.get('jw_out') else float(b.get('st_jw') or 0))
    sap=e.get('sap'); cnt=e.get('count')
    if not isinstance(sap,(int,float)): sap=(own if cnt is None else None)
    est=float(e['est']) if e.get('est') is not None else own
    return dict(sap=(float(sap) if sap is not None else None),cnt=(float(cnt) if cnt is not None else None),est=est,own=own)
def sup_split(b):
    """Stock of a material by supplier (user, Nov 2026: "Stock Check: do not combine suppliers, e.g. glass FLAT and KIBING separated"):
       [dict(name, sap, cnt, est)] when 2+ suppliers hold stock, else None. est = the supplier rows of MRP Daily (b['sup']: MB52 batches,
       or the count by supplier rolled forward); sap = MB52 batches by supplier (b['sup_sap']); cnt = the PPC count by supplier as counted
       (b['sup_cnt']). What MB52 / the count does not name goes to 'Unassigned', so the supplier rows add up to the material."""
    sups=b.get('sup') or []
    if not sups: return None
    ms=stock_src(b); est_by={}
    for sp in sups: est_by[str(sp.get('name'))]=est_by.get(str(sp.get('name')),0.0)+float(sp.get('stock') or 0.0)
    sapd={str(k):float(v) for k,v in (b.get('sup_sap') or {}).items() if k}; cntd={str(k):float(v) for k,v in (b.get('sup_cnt') or {}).items()}
    if sapd and ms['sap'] is not None and abs(ms['sap']-sum(sapd.values()))>0.5: sapd['Unassigned']=sapd.get('Unassigned',0.0)+ms['sap']-sum(sapd.values())
    if ms['cnt'] is not None and cntd and abs(ms['cnt']-sum(cntd.values()))>0.5: cntd['Unassigned']=cntd.get('Unassigned',0.0)+ms['cnt']-sum(cntd.values())
    names=list(dict.fromkeys(list(est_by)+list(sapd)+list(cntd)))
    names=[n for n in names if n!='Unassigned']+[n for n in names if n=='Unassigned']
    out=[dict(name=n,sap=(sapd.get(n,0.0) if sapd else None),cnt=(cntd.get(n,0.0) if ms['cnt'] is not None else None),est=est_by.get(n,0.0)) for n in names]
    out=[o for o in out if any(abs(v)>0.5 for v in (o['sap'] or 0.0,o['cnt'] or 0.0,o['est']))]
    return out if len(out)>=2 else None
def alert_stock_src(rows, A):
    """SAP stock (MB52) / physical count of every ALERT row (r['sap'], r['cnt']): one material = stock_src
       (in the row's unit; count = the counted members, None when none is counted); a material on 2+ WSA sheets (back glass 333) = the
       suppliers whose stock is the sheet's stock today (sup_split; r['sups']) - none found: None (shown '-')."""
    byb={x['b']['mat']:x['b'] for x in A}
    for r in rows:
        m=r['code']; b=byb.get(m); r['sap']=r['cnt']=None; r['sups']=None
        if b is None: continue
        if r.get('where'):
            sp=sup_split(b) or []; st=float(r['stock'] or 0.0)
            sub=_subset_match({o['name']:[o['est']] for o in sp},[st],max(1.0,0.002*abs(st))) if sp else None
            if sub is not None:
                r['sups']=list(sub); hit=[o for o in sp if o['name'] in sub]
                r['sap']=sum(o['sap'] or 0.0 for o in hit) if any(o['sap'] is not None for o in sp) else None
                r['cnt']=sum(o['cnt'] or 0.0 for o in hit) if any(o['cnt'] is not None for o in sp) else None
            continue
        ss_=stock_src(b); r['sap']=ss_['sap']; r['cnt']=ss_['cnt']      # physical: a supplier no customer with demand may use is shown too (user 28-Nov-2026)
ALERT_CHIP=[('FADBD8','C0392B'),('FDEBD0','B9770E'),('FCF3CF','9A7D0A'),('EAF2F8','1F618D'),('E3F4EC','1E8449'),('F2F4F6','8A949E')]   # ALERT status chips (Material Status colours)
def alert_label(code, ri=None):
    """Shortage label of a material = its ALERT class (user rule, Nov 2026: 'ALERT and Stock Check use the same label' - before,
       Stock Check used risk_info (days incl. in transit) and showed 'OK' for 280 / 268 that the ALERT lists as UNDER LEAD TIME).
       Text = ALERT class + first stock-zero day (UNDER LEAD TIME LATER: + below-lead-time date); a material on 2+ WSA sheets (333) =
       one line per sheet. Not on the ALERT: OK / No demand / No need (risk_info).
       Returns (text, worst ALERT class index or None)."""
    rows=[r for r in ALERT_ROWS if r['code']==code]
    if not rows: return ((ri or {}).get('cls') if (ri or {}).get('cls') in ('No demand','No need') else 'OK'),None
    def one(r):
        t=ALERT_CLS[r['cls']][0]; w_=r.get('why') or {}
        if r.get('first') is not None: t+=f" {r['first']:%d-%b}"
        elif r['cls']==3 and w_.get('below') is not None: t+=f" from {w_['below']:%d-%b}"
        return t+(f" ({r['where']})" if r.get('where') and len(rows)>1 else '')
    return '\n'.join(one(r) for r in rows),min(r['cls'] for r in rows)
def write_alert_sheet(wb, rows, d0, where):
    """The ALERT sheet = ALERT + Material Status in ONE sheet (user, Nov 2026: "combine alert and stock management into 1 sheet,
       style design stock management"; first tab of 1_MRP and of the Inbound file, the same rows). Material Status design:
       row 1 = group band (STORE / INVENTORY / QUALITY / PRODUCTION / TRANSIT LEAD TIME / BELOW LEAD TIME / WAREHOUSE), row 2 = header
       coloured by group, one row per ALERT row from row 3 (Sr. + light tinted group columns + soft status chips), summary boxes +
       four counter cards and the definitions line below. Rows = the ALERT rows (every old Material Status row is on it; EPE 420 =
       the EPE 480 + 420 group row). Material Status columns added: On hold, Reorder level, Status, In transit, Overdue in transit
       (`alert_stock_fields`); its duplicates (Short from, Inv. Days, Shortage risk, Days of cover (+ all in transit)) = the ALERT columns.
       No Action column (user, Nov 2026)."""
    from openpyxl.styles import Border, Side
    if 'ALERT' in wb.sheetnames: del wb['ALERT']
    if not MIS_CAT and MIS_PATH_[0]: load_mis_cat(MIS_PATH_[0])     # Category (item_category)
    ws=wb.create_sheet('ALERT',0); ws.sheet_view.showGridLines=False; ws.sheet_properties.tabColor='C00000'
    F=lambda **k: Font(name='Calibri',**k)
    def fill(c,col): c.fill=PatternFill('solid',fgColor=col)
    NAVY='14213D'; INK='1F2D3D'; MUTED='8A949E'; LINE='E3E8EE'
    STORE,INV,QUAL,PROD,LEAD,BELOW,WH=(('2E86C1','F4F9FD'),('E67E22','FEF8F1'),('17A589','F2FBF8'),('6C3483','F8F4FB'),
                                       ('C0392B','FDF3F2'),('1B2631','F4F6F8'),('229954','F4FBF7'))     # Material Status palette
    thin=Side(style='thin',color=LINE); bd=Border(left=thin,right=thin,top=thin,bottom=thin)
    H=['Sr.','Alert','Code','Item name','Category','UOM',
       'Estimated stock today','Days of cover (stock only)',
       'Reorder level',
       'Status','On hold (blocked + QI)',
       'Usage/Day (module plan)','Days of cover (+ all in transit)',
       'Transit lead time (days)','Stock-zero days inside the transit lead time','First stock-zero day','Stock zero from → to (⚠ = starts inside the transit lead time)',
       f'Below lead time from','Why (cause)','Days late',
       'In transit','Next arrival','Overdue in transit','Arrival at plant (Final AtA)']                   # no Action column (user, Nov 2026)
    W=[5,22,16,34,16,7, 12,11, 12, 14,11, 11,10, 9,12,11,38, 11,13,8, 12,24,11,34]
    WR={2,4,10,17,22,23,24}                                                            # wrapped columns
    LEFT={4,5,17,22,24}
    for j,w_ in enumerate(W,1): ws.column_dimensions[L(j)].width=w_
    NCOL=len(H)
    G=[(f'ALERT · STOCK MANAGEMENT · {d0:%d-%b-%Y}',1,6,NAVY),('STORE · stock',7,8,STORE[0]),('INVENTORY',9,9,INV[0]),('QUALITY',10,11,QUAL[0]),
       ('PRODUCTION',12,13,PROD[0]),('TRANSIT LEAD TIME · stock-out',14,17,LEAD[0]),(f'BELOW LEAD TIME · why',18,20,BELOW[0]),('WAREHOUSE · supply',21,24,WH[0])]
    # ---------- group band (row GB) ----------
    HR=ALERT_HR; GB=HR-1
    for t,c1,c2,col in G:
        if c2>c1: ws.merge_cells(start_row=GB,start_column=c1,end_row=GB,end_column=c2)
        c=ws.cell(GB,c1,t); c.font=F(size=10,bold=True,color='FFFFFF'); c.alignment=Alignment(horizontal='center',vertical='center')
        for j in range(c1,c2+1): fill(ws.cell(GB,j),col)
    ws.row_dimensions[GB].height=20
    # ---------- row 2: header ----------
    hc={}; TINT={}
    for t,c1,c2,col in G:
        for j in range(c1,c2+1):
            hc[j]=col
            if col!=NAVY: TINT[j]=dict((STORE,INV,QUAL,PROD,LEAD,BELOW,WH))[col]
    for j,h in enumerate(H,1):
        c=ws.cell(HR,j,h); c.font=F(size=9,bold=True,color='FFFFFF'); fill(c,hc[j]); c.border=bd
        c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
    ws.row_dimensions[HR].height=44
    # ---------- data from row HR+1 ----------
    CHIP=ALERT_CHIP                                                                   # = Material Status status chips
    CC={'LATE ARRIVAL':('FDEBD0','B9770E'),'LATE + NOT ENOUGH':('FADBD8','C0392B'),'NO SUPPLY':('FADBD8','C0392B')}
    SCH={'Above reorder':('EAF2F8','1F618D'),'Below reorder':('FCF3CF','9A7D0A'),'Out of Stock':('FADBD8','C0392B')}   # Material Status Status chips
    fm=[None,None,None,None,None,None,'#,##0','0.0','#,##0',None,'#,##0','#,##0.0','0.0','0','0','dd-mmm-yy',None,'dd-mmm-yy',None,'0','#,##0',None,'#,##0',None]
    ALERT_POS.clear(); rr=HR; SUBR=[]
    for i,r in enumerate(rows):
        rr+=1; ALERT_POS.append(rr); lab=ALERT_CLS[r['cls']][0]; w_=r.get('why') or {}
        nm=r['name']+(f"\n(WSA sheet: {r['where']})" if r.get('where') else '')+(f"\n({r['subnote']})" if r.get('subnote') else '')
        nx=f"{r['nx'][0]:%d-%b}: {r['nx'][1]:,.0f}"+(f" ({str(r['sup'])[:22]})" if r.get('sup') else '') if r.get('nx') else 'none'
        ovd=r.get('ovd'); hold=r.get('hold') or 0
        ata_=ata_show(r.get('ata'),[r['code']],r.get('need'))             # ONE arrival text for every output (user 26-Nov-2026)
        vals=[i+1,lab,r['code'],nm,item_category(r['code']),r['uom'],stock_show(r),r['days'],r.get('rop'),r.get('status') or '-',(hold if hold else '-'),
              r['need'],r['days_it'],r['lt'],r['zin'],r['first'].to_pydatetime() if r['first'] is not None else 'none',r['per'],
              w_['below'].to_pydatetime() if w_.get('below') is not None else 'not below',w_.get('cause') or '-',w_.get('late'),
              r.get('it') or 0,nx,('in today\'s arrivals' if ovd is None else (ovd if ovd else '-')),ata_]
        for j,(v,f_) in enumerate(zip(vals,fm),1):
            c=ws.cell(rr,j,v); c.font=F(size=10,color=INK); c.border=Border(bottom=thin)
            c.alignment=Alignment(horizontal=('left' if j in LEFT else 'center'),vertical='center',wrap_text=j in WR)
            if f_: c.number_format=f_
            if j in TINT: fill(c,TINT[j])
        bg,fg=CHIP[r['cls']]; fill(ws.cell(rr,2),bg); ws.cell(rr,2).font=F(size=10,bold=True,color=fg)
        sb,sf=SCH.get(str(vals[9]),('FDEBD0','B9770E')); fill(ws.cell(rr,10),sb); ws.cell(rr,10).font=F(size=10,bold=True,color=sf)
        if hold: ws.cell(rr,11).font=F(size=10,bold=True,color='B9770E')
        if r['days'] is not None and r['days']<tgt_days(r['lt']): ws.cell(rr,8).font=F(size=10,bold=True,color='C0392B')
        if r['zin']>0: fill(ws.cell(rr,15),'FADBD8'); ws.cell(rr,15).font=F(size=12,bold=True,color='C0392B')
        else: ws.cell(rr,15).font=F(size=10,color=MUTED)
        if r['first_in']: ws.cell(rr,16).font=F(size=10,bold=True,color='C0392B')
        if w_.get('t') is not None and w_['t']<=14: ws.cell(rr,18).font=F(size=10,bold=True,color='C0392B')
        if w_.get('cause'):
            bg2,fg2=CC[w_['cause']]; fill(ws.cell(rr,19),bg2); ws.cell(rr,19).font=F(size=10,bold=True,color=fg2)
        if isinstance(w_.get('late'),int) and w_['late']>0: ws.cell(rr,20).font=F(size=10,bold=True,color='B9770E')
        if isinstance(ovd,(int,float)) and ovd>0: ws.cell(rr,23).font=F(size=10,bold=True,color='C0392B')
        nl_=max(1,str(r['per']).count('\n')+1,nm.count('\n')+1,2 if len(nx)>24 else 1,str(ata_ or '').count('\n')+1,2 if len(lab)>18 else 1)
        ws.row_dimensions[rr].height=max(19,15*nl_)            # every period / arrival line visible in Excel
        for s in vp_subs(r):                                   # Settings ALLOWED VENDORS: one sub-row per supplier (= Action List supplier rows), no code / Sr.
            rr+=1; SUBR.append(rr)
            sv={4:vp_sub_name(s),6:r['uom'],7:s['st'],16:(s['first'].to_pydatetime() if s.get('first') is not None else None),17:s['per'] or 'no stock-zero day',21:s['it'],24:s['ata']}
            for j in range(1,NCOL+1):
                c=ws.cell(rr,j,sv.get(j)); c.font=F(size=9,italic=True,color=('8A949E' if s['nc'] else '4A5868')); c.border=Border(bottom=thin); fill(c,'F7F9FB')
                c.alignment=Alignment(horizontal=('left' if j in LEFT else 'center'),vertical='center',wrap_text=j in WR)
                if fm[j-1]: c.number_format=fm[j-1]
            if s['first'] is not None: ws.cell(rr,17).font=F(size=9,italic=True,bold=True,color='C0392B')
            ws.row_dimensions[rr].height=max(17,13*max(str(s['per']).count('\n')+1,str(s['ata'] or '').count('\n')+1))
    if not rows: ws.cell(HR+1,1,f'No material under lead time, none below lead time within {HZW()} and none out of stock inside its transit lead time.').font=F(size=11,bold=True,color='1E8449')
    last=max(rr,HR+1)
    ws.auto_filter.ref=f"A{HR}:{L(NCOL)}{last}"
    # no summary boxes on top, no class cards / definition note below the table (user 27-Nov-2026: ALERT rows 1-9 and 65-70 not needed)
    ws.freeze_panes=f'E{HR+1}'; ws.sheet_view.zoomScale=90; ws.print_title_rows=f'{GB}:{HR}'
    ws.page_setup.orientation='landscape'; ws.page_setup.fitToWidth=1; ws.page_setup.fitToHeight=0; ws.sheet_properties.pageSetUpPr.fitToPage=True
    # ---------- PRODUCTION (need / day, days incl. in transit) moves next to STORE stock (user, Nov 2026) ----------
    # new column order: STORE G-H | PRODUCTION I-J | INVENTORY K | QUALITY L-M (old order 9..13 = INVENTORY, QUALITY x2, PRODUCTION x2)
    from copy import copy as _cp
    for mr in list(ws.merged_cells.ranges):
        if mr.min_row==GB and mr.max_row==GB and 9<=mr.min_col and mr.max_col<=13: ws.unmerge_cells(str(mr))
    snap={(r_,j_):(ws.cell(r_,j_).value,_cp(ws.cell(r_,j_)._style)) for r_ in range(GB,last+1) for j_ in range(9,14)}
    for r_ in range(GB,last+1):
        for new_,old_ in zip(range(9,14),(12,13,9,10,11)):
            v_,st_=snap[(r_,old_)]; c_=ws.cell(r_,new_); c_.value=v_; c_._style=st_
    for new_,old_ in zip(range(9,14),(12,13,9,10,11)): ws.column_dimensions[L(new_)].width=W[old_-1]
    ws.merge_cells(start_row=GB,start_column=9,end_row=GB,end_column=10); ws.merge_cells(start_row=GB,start_column=12,end_row=GB,end_column=13)
    # ---------- SAP stock (MB52) + Physical count next to Estimated stock today (user, Nov 2026; = Stock Check columns) ----------
    _ins_cols(ws,8,2)
    for j,h in ((8,'SAP stock (MB52)'),(9,'Physical count')):
        c=ws.cell(HR,j,h); c._style=_cp(ws.cell(HR,7)._style)
        ws.column_dimensions[L(j)].width=12
    for rr in SUBR:
        for j in (8,9): ws.cell(rr,j)._style=_cp(ws.cell(rr,7)._style)
    for i,r in enumerate(rows):
        rr=ALERT_POS[i]; one=r.get('where') and r.get('sups') is None          # 333 sheet row whose suppliers were not found: '-'
        for j,v in ((8,(r['sap'] if r.get('sap') is not None else '-')),(9,(r['cnt'] if r.get('cnt') is not None else ('-' if one else 'no counting')))):
            c=ws.cell(rr,j,v); c._style=_cp(ws.cell(rr,7)._style)
            if isinstance(v,str): c.font=F(size=10,color=MUTED)
    # ---------- no INVENTORY / QUALITY on the ALERT (user, Nov 2026: "Alert sheet, delete quality and inventory"): M Reorder level, N Status, O On hold ----------
    _del_cols(ws,[13,14,15])
    ws.auto_filter.ref=f"A{HR}:{L(NCOL+2-3)}{last}"
    nc_note(ws,HR,rows,pos=ALERT_POS)
    return ws
def stock_show(r):
    """'Estimated stock today' as SHOWN = physical stock (user 28-Nov-2026 "show it, don't count it"): the counted stock r['stock'] +
       the stock no customer with demand may use (`vp_nc`). Every cover / stock-zero number keeps r['stock']."""
    return float(r.get('stock') or 0.0)+float((r.get('nc') or {}).get('st') or 0.0)
def nc_note(ws, hr, rows, first=None, pos=None):
    """Excel note on 'Estimated stock today' of every row with stock / in transit not counted (`vp_nc` text): why the days of cover
       are lower than the stock shown. Rows from hr+1 (first = other first data row)."""
    from openpyxl.comments import Comment
    js=next((c_.column for c_ in ws[hr] if str(c_.value or '').strip()=='Estimated stock today'),None)
    if js is None: return
    for i,r in enumerate(rows):
        t_=(r.get('nc') or {}).get('txt')
        if t_:
            c_=ws.cell(pos[i] if pos else (first or hr+1)+i,js); c_.comment=Comment(t_+'. Shown here, left out of every days of cover / stock-zero number.','update_inbound')
            c_.font=Font(name='Calibri',size=10,color='1F2D3D',underline='single')
def _add_uom_col(ws, uoms, top=3):
    """UOM column right after Category (column F) on Material Status / Other Materials (user, Nov 2026: "Add UOM into daily material
       status"): everything from column F moves one right; the navy band A1:E1 grows to A1:F1; header F2 'UOM'; data rows from `top`."""
    from copy import copy as _cp
    _ins_cols(ws,6,1)
    for m_ in list(ws.merged_cells.ranges):
        if m_.min_row==1 and m_.max_row==1 and m_.min_col==1 and m_.max_col==5: ws.unmerge_cells(str(m_)); ws.merge_cells('A1:F1')
    ws.cell(1,6)._style=_cp(ws.cell(1,5)._style)
    h=ws.cell(2,6,'UOM'); h._style=_cp(ws.cell(2,5)._style)
    for i,u in enumerate(uoms):
        c=ws.cell(top+i,6,(str(u).strip() if u not in (None,'') else '-')); c._style=_cp(ws.cell(top+i,5)._style)
        c.alignment=Alignment(horizontal='center',vertical='center')
    ws.column_dimensions['F'].width=7
def _del_cols(ws, cols):
    """Delete whole columns (list of indexes): cells right of them move left; merged ranges shrink (a range left with no column is
       dropped; when its first cell is deleted, its value + style go to its first remaining column); widths move left."""
    dl=sorted(set(cols)); ds=set(dl)
    if not dl: return
    nc=lambda c: c-sum(1 for d in dl if d<c)
    rngs=[(m_.min_row,m_.min_col,m_.max_row,m_.max_col) for m_ in ws.merged_cells.ranges]
    for m_ in list(ws.merged_cells.ranges): ws.unmerge_cells(str(m_))
    for r1,c1,r2,c2 in rngs:                                      # anchor deleted: keep its value on the first remaining column
        keep=[c for c in range(c1,c2+1) if c not in ds]
        if c1 in ds and keep:
            a=ws.cell(r1,c1); b=ws.cell(r1,keep[0]); b._value=a._value; b.data_type=a.data_type; b._style=a._style
    maxc=ws.max_column
    wd={j:ws.column_dimensions[L(j)].width for j in range(1,maxc+1) if L(j) in ws.column_dimensions}
    mv=sorted([(k,c) for k,c in ws._cells.items() if k[1]>dl[0]],key=lambda kc:kc[0][1])
    for (r_,c_),cell in mv: del ws._cells[(r_,c_)]
    for (r_,c_),cell in mv:
        if c_ in ds: continue
        n_=ws.cell(r_,nc(c_)); n_._value=cell._value; n_.data_type=cell.data_type; n_._style=cell._style
    for j in range(dl[0],maxc+1):
        if j in ds: continue
        if wd.get(j): ws.column_dimensions[L(nc(j))].width=wd[j]
    for r1,c1,r2,c2 in rngs:
        keep=[c for c in range(c1,c2+1) if c not in ds]
        if not keep: continue
        a_,b_=nc(keep[0]),nc(keep[-1])
        if r2>r1 or b_>a_: ws.merge_cells(start_row=r1,start_column=a_,end_row=r2,end_column=b_)
def _ins_cols(ws, at, n):
    """Insert n empty columns before column `at`: cells (value + style), merged ranges (a range across `at` grows), column widths
       move right. openpyxl's insert_cols does not move merged ranges or widths."""
    rngs=[(m_.min_row,m_.min_col,m_.max_row,m_.max_col) for m_ in ws.merged_cells.ranges]
    for m_ in list(ws.merged_cells.ranges): ws.unmerge_cells(str(m_))
    mv=sorted([(k,c) for k,c in ws._cells.items() if k[1]>=at],key=lambda kc:-kc[0][1])
    for (r_,c_),cell in mv:
        del ws._cells[(r_,c_)]
        nc=ws.cell(r_,c_+n); nc._value=cell._value; nc.data_type=cell.data_type; nc._style=cell._style
    wd={j:ws.column_dimensions[L(j)].width for j in range(at,ws.max_column+1) if L(j) in ws.column_dimensions}
    for j in sorted(wd,reverse=True):
        if wd[j]: ws.column_dimensions[L(j+n)].width=wd[j]
    for r1,c1,r2,c2 in rngs:
        ws.merge_cells(start_row=r1,start_column=c1+(n if c1>=at else 0),end_row=r2,end_column=c2+(n if c2>=at else 0))
ALERT_POS=[]     # sheet row of each ALERT row (supplier sub-rows sit between them) - ALERT (formulas) compares with that row
ALERT_HR=2       # ALERT sheet: row 1 group band, row 2 header, data from row 3 (user 27-Nov-2026: no summary boxes / cards / note)
ALERT_FX="ALERT (formulas)"; ALERT_CALC="ALERT Calc"; ALERT_PL="ALERT Plan lines"; ALERT_AL="ALERT Arrival lines"; OEM_IS=[None]
def write_alert_formulas(wb, rows, d0):
    """'ALERT (formulas)' + 'ALERT Calc' (user, Nov 2026: "New Alert sheet, I want a new same sheet with excel formulas so that I can
       check your work"): the SAME rows and columns as the ALERT, but every calculated column is an Excel formula. Inputs (yellow) =
       Stock today, On hold, Transit lead time, In transit, Overdue + per day the plan and the arrivals (ALERT Calc). ALERT Calc rebuilds the
       projection: Projected stock = stock + arrivals - plan, day by day; 'Stock zero' = projected stock < 0; 'Below lead time' = projected
       stock / need < 21 in the first 16 weeks. Column Y 'Check vs ALERT' compares each calculated column with the ALERT sheet
       (OK / DIFF + what differs in Z). Both outputs. Excel calculates on open (no cached values)."""
    from openpyxl.styles import Border, Side
    for n_ in (ALERT_FX,ALERT_CALC,ALERT_PL,ALERT_AL):
        if n_ in wb.sheetnames: del wb[n_]
    pos=wb.sheetnames.index('ALERT')+1 if 'ALERT' in wb.sheetnames else 0
    fx=wb.create_sheet(ALERT_FX,pos); cs=wb.create_sheet(ALERT_CALC,pos+1)
    for w_ in (fx,cs): w_.sheet_view.showGridLines=False
    fx.sheet_properties.tabColor='E26B0A'; cs.sheet_properties.tabColor='8A949E'
    F=lambda **k: Font(name='Calibri',**k)
    def fill(c,col): c.fill=PatternFill('solid',fgColor=col)
    NAVY='14213D'; INK='1F2D3D'; MUTED='8A949E'; LINE='E3E8EE'; INP='FFF2CC'
    thin=Side(style='thin',color=LINE); bd=Border(left=thin,right=thin,top=thin,bottom=thin)
    TG='lead time'; HZ=ALERT_HORIZON
    nd=max([len(r.get('P') or []) for r in rows]+[1])
    LC=L(5+nd)                                                     # last date column of ALERT Calc (dates from column F)
    # ---------------- ALERT Plan lines / ALERT Arrival lines (the data behind Plan and Arrivals) ----------------
    pls_=wb.create_sheet(ALERT_PL,pos+2); als_=wb.create_sheet(ALERT_AL,pos+3)
    for w_ in (pls_,als_): w_.sheet_view.showGridLines=False; w_.sheet_properties.tabColor='8A949E'
    def _lines_sheet(w_,title,note,H_,W_,rows_):
        w_['A1']=title; w_['A1'].font=F(size=16,bold=True,color=NAVY)
        w_['A2']=note; w_['A2'].font=F(size=9,italic=True,color=MUTED); w_.merge_cells(start_row=2,start_column=1,end_row=2,end_column=len(H_))
        w_['A2'].alignment=Alignment(wrap_text=True,vertical='top'); w_.row_dimensions[2].height=40
        for j,(h,wd) in enumerate(zip(H_,W_),1):
            c=w_.cell(4,j,h); fill(c,NAVY); c.font=F(size=9,bold=True,color='FFFFFF'); c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
            w_.column_dimensions[L(j)].width=wd
        w_.row_dimensions[4].height=32
        fi_=PatternFill('solid',fgColor=INP)                          # one style object for all cells (16,000 rows: speed)
        for i_,(vals,fms,inp) in enumerate(rows_,5):
            for j,(v_,f_) in enumerate(zip(vals,fms),1):
                c=w_.cell(i_,j,v_)
                if f_: c.number_format=f_
                if j in inp: c.fill=fi_
        w_.freeze_panes='C5'; w_.auto_filter.ref=f"A4:{L(len(H_))}{4+max(1,len(rows_))}"
        return 4+max(1,len(rows_))
    prow=[]; arow=[]
    for i,r in enumerate(rows):
        for l_ in r.get('plan_l') or []:
            n_=5+len(prow)
            prow.append(([i+1,r['code'],l_['code'],(d0+pd.Timedelta(days=l_['k'])).to_pydatetime(),l_['cu'],l_['ln'],l_['src'],l_['mods'],l_['rate'],l_['f'],f'=H{n_}*I{n_}*J{n_}'],
                         [None,None,None,'dd-mmm-yy',None,None,None,'#,##0.##','0.######','0.####','#,##0.##'],{8,9,10}))
        for l_ in r.get('arr_l') or []:
            n_=5+len(arow)
            arow.append(([i+1,r['code'],l_['code'],l_['sup'],(int(l_['po']) if str(l_['po']).isdigit() else (l_['po'] or None)),l_['bl'] or None,l_['raw'].to_pydatetime(),f"=MAX(G{n_},'{ALERT_CALC}'!$F$3)",l_['qty'],l_['f'],f'=I{n_}*J{n_}',l_['src']],
                         [None,None,None,None,None,None,'dd-mmm-yy','dd-mmm-yy','#,##0.##','0.####','#,##0.##',None],{7,9,10}))
    PL_last=_lines_sheet(pls_,"ALERT Plan lines - the plan behind 'Plan (demand)' on ALERT Calc",
        f"One row per ALERT row (Sr.) x material x day x customer x production line, from the module plan (PPC Plan + Revised Plan; OEM customers left out). "
        f"Demand = Modules x Qty per module x Factor (Excel formula). Qty per module = your PPC BOM, else the MIS Standard BOM (column G says which); 'MIS requirement' lines = "
        f"MIS quantity (Modules = quantity, Qty per module = 1). Factor = unit of the ALERT row (substitute group: member -> main material). Yellow = input. "
        f"ALERT Calc Plan = SUMIFS(Demand, Sr., Date).",
        ['Sr. (ALERT row)','ALERT code','Material code','Date','Customer','Production line','Qty per module from','Modules','Qty per module','Factor','Demand (formula)'],
        [8,16,16,11,24,16,30,11,11,8,12],prow)
    AL_last=_lines_sheet(als_,"ALERT Arrival lines - the shipments behind 'Arrivals (in transit)' on ALERT Calc",
        f"One row per ALERT row (Sr.) x shipment: PPC In_Transit (pending qty, Final AtA, material traced to your BOM / WSA PO + BL#; alias codes converted), plus MIS Adhoc / "
        f"substitute receipts per day. Counted on = MAX(arrival, today) - late lines are counted today (Excel formula). Qty = Pending qty x Factor (formula). "
        f"Back glass 333 per WSA sheet: the POs of that sheet. ALERT Calc Arrivals = SUMIFS(Qty, Sr., Counted on); In transit = all lines of the row up to the plan end.",
        ['Sr. (ALERT row)','ALERT code','Material code','Supplier','PO','BL / HBL','Arrival at plant (PPC)','Counted on (formula)','Pending qty (MRP unit)','Factor','Qty (formula)','Source'],
        [8,16,16,30,12,18,12,12,13,8,13,26],arow)
    PLK=f"'{ALERT_PL}'!$A$5:$A${PL_last}"; PLD=f"'{ALERT_PL}'!$D$5:$D${PL_last}"; PLQ=f"'{ALERT_PL}'!$K$5:$K${PL_last}"
    ALK=f"'{ALERT_AL}'!$A$5:$A${AL_last}"; ALD=f"'{ALERT_AL}'!$H$5:$H${AL_last}"; ALQ=f"'{ALERT_AL}'!$K$5:$K${AL_last}"
    # ---------------- ALERT Calc ----------------
    cs['A1']='ALERT Calc - the daily projection behind the ALERT (Excel formulas)'; cs['A1'].font=F(size=16,bold=True,color=NAVY)
    cs['A2']=(f"Per ALERT row: Plan = SUMIFS of '{ALERT_PL}' (modules x qty per module per customer / line / day, OEM left out; substitute members on their own rows; "
              f"back glass 333 = the customers of the WSA sheet) and Arrivals = SUMIFS of '{ALERT_AL}' (PPC In_Transit shipments counted on max(arrival, today); 333 = the POs of the sheet). "
              f"Projected stock = previous day + arrivals - plan, never below 0 (day 1: Estimated stock today on '{ALERT_FX}'). Stock zero = 1 when previous stock + arrivals - plan < 0 (lost demand is not owed). "
              f"Below {TG} = 1 when projected stock / Usage/Day < {TG} (first {HZW()} = day number < {HZ}).")
    cs['A2'].font=F(size=9,italic=True,color=MUTED); cs.merge_cells('A2:N2'); cs['A2'].alignment=Alignment(wrap_text=True,vertical='top'); cs.row_dimensions[2].height=42
    for j,h in enumerate(['Sr.','Code','Item name','Line','Estimated stock today'],1):
        for rr_,v_ in ((3,h),(4,'Day number' if j==4 else None)):
            c=cs.cell(rr_,j,v_); fill(c,NAVY); c.font=F(size=9,bold=True,color='FFFFFF'); c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
    for k in range(nd):
        c=cs.cell(3,6+k,(d0+pd.Timedelta(days=k)).to_pydatetime()); c.number_format='dd-mmm'; fill(c,NAVY); c.font=F(size=9,bold=True,color='FFFFFF'); c.alignment=Alignment(horizontal='center')
        c=cs.cell(4,6+k,k); fill(c,'D6DBDF'); c.font=F(size=8,color=INK); c.alignment=Alignment(horizontal='center')
        cs.column_dimensions[L(6+k)].width=9
    for j,w_ in enumerate([5,16,30,26,11],1): cs.column_dimensions[L(j)].width=w_
    DT=f"'{ALERT_CALC}'!$F$3:${LC}$3"; IDX=f"'{ALERT_CALC}'!$F$4:${LC}$4"
    F8=F(size=8,color=INK); FB8=F(size=8,color=INK,bold=True); FIL_=PatternFill('solid',fgColor='EAF2F8'); BL_=Border(bottom=Side(style='thin',color=LINE)); BL2_=Border(bottom=Side(style='thin',color='8A949E'))
    LINES=['Plan (demand)','Arrivals (in transit)','Projected stock',f'Stock zero (1 = yes)',f'Below {TG} (1 = yes, first {HZW()})']
    for i,r in enumerate(rows):
        b0=5+5*i; fr=3+i
        for q,ln in enumerate(LINES):
            rr=b0+q
            ln_=(ln+' - Settings ALLOWED VENDORS: minus stock / arrivals no short customer may use' if (q==2 and r.get('vlock')) else ln)
            for j,v_ in enumerate([i+1,r['code'],r['name']+(f" ({r['where']})" if r.get('where') else ''),ln_],1):
                c=cs.cell(rr,j,v_ if (q==0 or j==4 or (j==1 and q==1)) else None); c.font=F(size=9,bold=(j==4 and q==2),color=INK); c.alignment=Alignment(vertical='center')
            if q==0:
                c=cs.cell(rr,5,f"='{ALERT_FX}'!G{fr}"); c.number_format='#,##0'; c.font=F(size=9,color=INK)
            for k in range(nd):
                col=L(6+k); prv=L(5+k)
                if q==0: v_=f'=SUMIFS({PLQ},{PLK},$A{b0},{PLD},{col}$3)'
                elif q==1: v_=f'=SUMIFS({ALQ},{ALK},$A{b0},{ALD},{col}$3)'
                elif q in (2,3):                                 # stock never below 0 (user 23-Nov-2026): row 3 = floored stock, row 4 = 1 when floored stock + arrivals - plan < 0
                    e_=(f"$E{b0}+{col}{b0+1}-{col}{b0}" if k==0 else f"{prv}{b0+2}+{col}{b0+1}-{col}{b0}")
                    vl_=r.get('vlock')                           # Settings ALLOWED VENDORS: stock + arrivals no short customer may use (change vs the day before)
                    dl_=((vl_[k] if k<len(vl_) else 0.0)-((vl_[k-1] if k-1<len(vl_) else 0.0) if k else 0.0)) if vl_ else 0.0
                    if abs(dl_)>1e-9: e_+=f"-({round(dl_,6)})"
                    v_=f"=MAX(0,{e_})" if q==2 else f"=IF({e_}<0,1,0)"
                else: v_=f"=IF(AND({col}$4<{HZ},'{ALERT_FX}'!$L{fr}>0),IF({col}{b0+2}/'{ALERT_FX}'!$L{fr}<MAX(1,'{ALERT_FX}'!$N{fr}),1,0),0)"
                c=cs.cell(rr,6+k,v_); c.font=(FB8 if q==2 else F8); c.number_format='#,##0' if q<3 else '0'
                if q<2: c.fill=FIL_
            for j in range(1,6+nd): cs.cell(rr,j).border=(BL2_ if q==4 else BL_)
    cs.freeze_panes='F5'
    # ---------------- ALERT (formulas) ----------------
    H=['Sr.','Alert','Code','Item name','Category','UOM','Estimated stock today','Days of cover (stock only)','Reorder level','Status','On hold (blocked + QI)',
       'Usage/Day (module plan)','Days of cover (+ all in transit)','Transit lead time (days)','Stock-zero days inside the transit lead time','First stock-zero day',
       'Stock-zero days in the plan (count)',f'Below {TG} from','Why (cause)','Days late','In transit','Next arrival','Overdue in transit',
       'In transit arriving after the plan end','Check vs ALERT','What differs',
       f'Helper: day number below {TG}','Helper: arrivals after that day',f'Helper: first arrival day after it',f'Helper: gap to {TG} that day','Helper: next arrival day']
    W=[5,22,16,30,14,7,12,11,12,14,11,11,10,9,12,11,11,11,13,8,12,18,11,12,10,30,10,12,10,12,10]
    for j,w_ in enumerate(W,1): fx.column_dimensions[L(j)].width=w_
    G=[('ALERT (formulas) · same rows as ALERT · yellow = input, white = Excel formula',1,6,NAVY),('STORE · stock',7,8,'2E86C1'),('INVENTORY',9,9,'E67E22'),
       ('QUALITY',10,11,'17A589'),('PRODUCTION',12,13,'6C3483'),('TRANSIT LEAD TIME · stock-out',14,17,'C0392B'),(f'BELOW LEAD TIME · why',18,20,'1B2631'),
       ('WAREHOUSE · supply',21,24,'229954'),('CHECK',25,26,'C00000'),('HELPER (formulas)',27,31,'8A949E')]
    hc={}
    for t,c1,c2,col in G:
        if c2>c1: fx.merge_cells(start_row=1,start_column=c1,end_row=1,end_column=c2)
        c=fx.cell(1,c1,t); c.font=F(size=10,bold=True,color='FFFFFF'); c.alignment=Alignment(horizontal='center',vertical='center')
        for j in range(c1,c2+1): fill(fx.cell(1,j),col); hc[j]=col
    for j,h in enumerate(H,1):
        c=fx.cell(2,j,h); c.font=F(size=9,bold=True,color='FFFFFF'); fill(c,hc[j]); c.border=bd; c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
    fx.row_dimensions[1].height=20; fx.row_dimensions[2].height=44
    INPUTS={7,11,14,23}
    fm={7:'#,##0',8:'0.0',9:'#,##0',11:'#,##0',12:'#,##0.0',13:'0.0',14:'0',15:'0',16:'dd-mmm-yy',17:'0',18:'dd-mmm-yy',20:'0',21:'#,##0',23:'#,##0',24:'#,##0',27:'0',28:'#,##0',29:'0',30:'#,##0',31:'0'}
    AO=ALERT_HR-2; A_="ALERT!"
    for i,r in enumerate(rows):
        n=3+i; b0=5+5*i; ap=ALERT_POS[i] if i<len(ALERT_POS) else n+AO     # the same material on the ALERT sheet (supplier sub-rows in between)
        PL=f"'{ALERT_CALC}'!$F${b0}:${LC}${b0}"; AR=f"'{ALERT_CALC}'!$F${b0+1}:${LC}${b0+1}"; SR=f"'{ALERT_CALC}'!$F${b0+2}:${LC}${b0+2}"
        ZR=f"'{ALERT_CALC}'!$F${b0+3}:${LC}${b0+3}"; BR=f"'{ALERT_CALC}'!$F${b0+4}:${LC}${b0+4}"
        ovd=r.get('ovd')
        f_={1:i+1,3:r['code'],4:r['name']+(f"\n(WSA sheet: {r['where']})" if r.get('where') else ''),5:item_category(r['code']),6:r['uom'],
            7:float(r['stock'] or 0.0),11:float(r.get('hold') or 0.0),14:int(r['lt'] or 0),
            21:f'=SUMIFS({ALQ},{ALK},$A{n},{ALD},"<="&\'{ALERT_CALC}\'!${LC}$3)',
            23:("in today's arrivals" if ovd is None else float(ovd or 0.0)),
            2:f'=IF(L{n}=0,\"{ALERT_CLS[5][0]}\",IF(O{n}>0,\"{ALERT_CLS[0][0]}\",IF(ISNUMBER(P{n}),\"{ALERT_CLS[1][0]}\",IF(H{n}<MAX(1,N{n}),\"{ALERT_CLS[2][0]}\",IF(ISNUMBER(R{n}),\"{ALERT_CLS[3][0]}\",\"{ALERT_CLS[4][0]}\")))))',
            12:f'=IFERROR(AVERAGEIFS({PL},{PL},">0",{IDX},"<"&MAX(1,N{n})),0)',
            8:f'=IF(L{n}>0,MAX(0,G{n})/L{n},"")',
            9:f'=L{n}*(N{n}+MAX(1,N{n}))',
            10:f'=IF(G{n}<=0,"Out of Stock",IF(G{n}<I{n},"Below reorder",IF(ISNUMBER(P{n}),"Short "&TEXT(P{n},"dd-mmm"),"Above reorder")))',
            13:f'=IF(L{n}>0,MAX(0,G{n}+U{n})/L{n},"")',
            15:f'=IF(N{n}>0,SUMIFS({ZR},{IDX},"<"&N{n}),0)',
            16:f'=IFERROR(INDEX({DT},MATCH(1,{ZR},0)),"none")',
            17:f'=SUM({ZR})',
            27:f'=IFERROR(MATCH(1,{BR},0)-1,"")',
            18:f'=IF(AA{n}="","not below",INDEX({DT},AA{n}+1))',
            28:f'=IF(AA{n}="","",SUMIFS({AR},{IDX},">"&AA{n}))',
            29:f'=IF(AA{n}="","",IF(COUNTIFS({AR},">0",{IDX},">"&AA{n})=0,"",_xlfn.MINIFS({IDX},{AR},">0",{IDX},">"&AA{n})))',
            30:f'=IF(AA{n}="","",MAX(1,N{n})*L{n}-INDEX({SR},AA{n}+1))',
            19:f'=IF(AA{n}="","-",IF(AC{n}="","NO SUPPLY",IF(AB{n}>=AD{n},"LATE ARRIVAL","LATE + NOT ENOUGH")))',
            20:f'=IF(OR(AA{n}="",AC{n}=""),"",AC{n}-AA{n})',
            31:f'=IF(COUNTIF({AR},">0")=0,"",_xlfn.MINIFS({IDX},{AR},">0"))',
            22:f'=IF(AE{n}="","none",TEXT(INDEX({DT},AE{n}+1),"dd-mmm")&": "&TEXT(INDEX({AR},AE{n}+1),"#,##0"))',
            24:f'=SUMIFS({ALQ},{ALK},$A{n},{ALD},">"&\'{ALERT_CALC}\'!${LC}$3)'}
        cl=lambda a,b: f'ABS(N({a})-N({b}))<=MAX(0.01,ABS(N({b}))*0.000001)'
        # ALERT columns: H:I = SAP stock (MB52) + Physical count (not here), no Reorder level / Status / On hold (Nov 2026) -> letters differ from this sheet's
        parts=[(f'B{n}=ALERT!B{ap}','Alert'),(cl(f'H{n}',f'ALERT!J{ap}'),'Days of cover (stock only)'),(cl(f'L{n}',f'ALERT!K{ap}'),'Usage/Day'),(cl(f'M{n}',f'ALERT!L{ap}'),'Days of cover (+ all in transit)'),(f'O{n}=ALERT!N{ap}','Stock-zero days in transit lead time'),
               (f'IF(ISNUMBER(P{n}),P{n}=ALERT!O{ap},ALERT!O{ap}="none")','First stock-zero day'),(f'IF(ISNUMBER(R{n}),R{n}=ALERT!Q{ap},ALERT!Q{ap}="not below")',f'Below {TG} from'),
               (f'S{n}=ALERT!R{ap}','Why'),(cl(f'U{n}',f'ALERT!T{ap}'),'In transit'),(f'N(T{n})=N(ALERT!S{ap})','Days late'),(f'LEFT(ALERT!U{ap},LEN(V{n}))=V{n}','Next arrival')]
        f_[26]='=TRIM('+'&'.join(f'IF({c_},"","{t_}; ")' for c_,t_ in parts)+')'
        f_[25]=f'=IF(Z{n}="","OK","DIFF")'
        for j in range(1,32):
            c=fx.cell(n,j,f_.get(j)); c.font=F(size=10,color=(MUTED if j>=27 else INK)); c.border=Border(bottom=thin)
            c.alignment=Alignment(horizontal=('left' if j in (4,5,22,26) else 'center'),vertical='center',wrap_text=j in (2,4,10,23,26))
            if j in fm: c.number_format=fm[j]
            if j in INPUTS: fill(c,INP)
        fx.row_dimensions[n].height=max(19,15*(f_[4].count('\n')+1),30)
    last=2+max(1,len(rows))
    from openpyxl.formatting.rule import CellIsRule
    for k_,(lab_,_,_,_) in enumerate(ALERT_CLS):                   # Alert chips = the ALERT sheet colours
        fx.conditional_formatting.add(f"B3:B{last}",CellIsRule(operator='equal',formula=[f'"{lab_}"'],fill=PatternFill('solid',fgColor=ALERT_CHIP[k_][0]),font=Font(name='Calibri',bold=True,color=ALERT_CHIP[k_][1])))
    fx.conditional_formatting.add(f"Y3:Y{last}",CellIsRule(operator='equal',formula=['"OK"'],fill=PatternFill('solid',fgColor='D5F5E3'),font=Font(name='Calibri',bold=True,color='1E8449')))
    fx.conditional_formatting.add(f"Y3:Y{last}",CellIsRule(operator='equal',formula=['"DIFF"'],fill=PatternFill('solid',fgColor='FADBD8'),font=Font(name='Calibri',bold=True,color='C0392B')))
    tr_=last+2
    fx.cell(tr_,1,'Rows checked').font=F(size=10,bold=True,color=NAVY); fx.cell(tr_,4,f'=COUNTA(Y3:Y{last})').font=F(size=10,bold=True,color=NAVY)
    fx.cell(tr_+1,1,'OK (same as ALERT)').font=F(size=10,bold=True,color='1E8449'); fx.cell(tr_+1,4,f'=COUNTIF(Y3:Y{last},"OK")').font=F(size=10,bold=True,color='1E8449')
    fx.cell(tr_+2,1,'DIFF').font=F(size=10,bold=True,color='C0392B'); fx.cell(tr_+2,4,f'=COUNTIF(Y3:Y{last},"DIFF")').font=F(size=10,bold=True,color='C0392B')
    nt=fx.cell(tr_+4,1,(f"Yellow = input (Estimated stock today = PPC count rolled forward, else SAP (MB52), JW at the OEM left out - see Stock Check / Stock Data; On hold = MIS MB52 blocked + QI; "
                        f"Transit lead time = Settings; Overdue = PPC arrival date passed). In transit = SUMIFS of '{ALERT_AL}' up to the plan end; plan + arrivals per day on '{ALERT_CALC}' = SUMIFS of '{ALERT_PL}' / '{ALERT_AL}'. "
                        f"White = Excel formula. Usage/Day = average plan per production day (plan > 0) in the over the transit lead time. Inv. Days = stock / need. "
                        f"Reorder level = need × (transit lead time + transit lead time). Stock-zero day = projected stock < 0; inside the transit lead time = day number < transit lead time. "
                        f"Below {TG} from = first day projected stock / need < {TG} ({HZW()}). Why: NO SUPPLY = no arrival after that day; LATE ARRIVAL = arrivals after "
                        f"that day ≥ the gap to {TG}; else LATE + NOT ENOUGH. Days late = first arrival after that day − that day. Back glass 333: one row per WSA sheet; "
                        f"a material and its substitute = one row each (own numbers). Check vs ALERT = every calculated column compared with the ALERT sheet (same row; the ALERT has "
                        f"SAP stock (MB52) + Physical count in H:I and no Reorder level / Status / On hold, so its column letters differ from this sheet's)."))
    nt.font=F(size=8,italic=True,color=MUTED); nt.alignment=Alignment(wrap_text=True,vertical='top')
    fx.merge_cells(start_row=tr_+4,start_column=1,end_row=tr_+4,end_column=26); fx.row_dimensions[tr_+4].height=64
    fx.freeze_panes='E3'; fx.sheet_view.zoomScale=90; fx.auto_filter.ref=f"A2:{L(31)}{last}"
    try: wb.calculation.fullCalcOnLoad=True                      # Excel calculates every formula when the file opens
    except Exception as _e_: _swallow(4540,_e_); pass
    return fx
ETD_LATE=[]        # WSA delivery rows not on PPC In_Transit, ETD passed or to come (built with the Inbound file; Inbound + 1_MRP + daily email)
ETD_SHEET='ETD alert'
ETD_SOON=7         # 'Ships within 7 days': ETD this close and still not on PPC In_Transit
WSA_SHIPPER={}     # material code -> shipper names typed on your WSA delivery rows
def _etd_arow(r):
    """The ALERT row of a WSA delivery row's material (333: the row of that WSA sheet); None when not on the ALERT."""
    mrp_m={str(x_['b']['mat']).strip() for x_ in (AL_ORDER_ or []) if not x_.get('sup_row')}
    for c_ in ([c_ for c_ in r['codes'] if c_ in mrp_m] or list(r['codes'])):
        al=[a_ for a_ in ALERT_ROWS if a_['code']==c_]
        if len(al)>1: al=[a_ for a_ in al if a_.get('where')==r['sheet']] or al
        if al: return al[0]
    return None
def _etd_alert(r):
    """ALERT class + first stock-zero day of the planning sheet's material (333: the row of that WSA sheet)."""
    mrp_m={str(x_['b']['mat']).strip() for x_ in (AL_ORDER_ or []) if not x_.get('sup_row')}
    cs=[c_ for c_ in r['codes'] if c_ in mrp_m] or list(r['codes'])
    for c_ in cs:
        al=[a_ for a_ in ALERT_ROWS if a_['code']==c_]
        if len(al)>1: al=[a_ for a_ in al if a_.get('where')==r['sheet']] or al
        if al:
            a_=al[0]
            return c_,str(a_['name']).split('\n')[0],a_['cls'],(a_['first'] if a_.get('first') is not None else None),(a_.get('uom') or r.get('uom'))
    return (cs[0] if cs else '-'),'-',None,None,r.get('uom')
ETD_FLAGS=[('ARRIVES AFTER STOCK-OUT','C00000','FFFFFF'),('ETD PASSED','E67E22','FFFFFF'),(f'SHIPS WITHIN {ETD_SOON} DAYS','F4D03F','1F2D3D'),('ON PLAN','1E8449','FFFFFF')]
def _etd_flag(r, first):
    """0 = arrives after the stock-out (expected arrival after the material's first stock-zero day), 1 = ETD passed, 2 = ETD within
       ETD_SOON days, 3 = on plan."""
    if first is not None and r['eta_now']>pd.Timestamp(first).normalize(): return 0
    if r['passed']: return 1
    return 2 if -r['late']<=ETD_SOON else 3
def write_etd_late_sheet(wb, rows, where):
    """Sheet 'ETD alert' (Inbound and 1_MRP, the same rows): your WSA delivery rows (not cancelled) with an ETD whose BL# is not on
       PPC In_Transit, ETD passed AND to come. Alert: ARRIVES AFTER STOCK-OUT (expected arrival after the material's first stock-zero
       day), ETD PASSED, SHIPS WITHIN ETD_SOON DAYS, ON PLAN. ETD vs today: 'Late N d' / 'In N d'. Expected arrival = WSA plant date; ETD
       passed = base date + planned transit (WSA plant date - ETD) if it ships today; delay = that - WSA plant date. With the material's
       ALERT class and first stock-zero day. Not in any projection (the outputs use PPC In_Transit) - information to chase."""
    from openpyxl.styles import Border, Side
    for n_ in (ETD_SHEET,'ETD late'):
        if n_ in wb.sheetnames: del wb[n_]
    ix_=max([i_+1 for i_,n_ in enumerate(wb.sheetnames) if n_.startswith('ALERT')]+[0])
    ws=wb.create_sheet(ETD_SHEET,ix_); ws.sheet_view.showGridLines=False; ws.sheet_properties.tabColor='C0392B'
    F=lambda **k: Font(name='Calibri',**k); thin=Side(style='thin',color='D0D5DD'); BD=Border(bottom=thin)
    base_=rows[0]['base'] if rows else None
    info=[(r,)+_etd_alert(r) for r in rows]
    info=[t_+(_etd_flag(t_[0],t_[4]),) for t_ in info]
    info.sort(key=lambda t_:(t_[6],-t_[0]['late'] if t_[0]['passed'] else 0,t_[0]['etd'],t_[0]['label']))
    cnt=Counter(t_[6] for t_ in info)
    ws['A1']='ETD alert - WSA deliveries not in transit yet'; ws['A1'].font=F(size=20,bold=True,color='1F3864')
    ws['A2']=(f"Your WSA delivery rows (not cancelled) with an ETD whose BL# is not on PPC In_Transit (base date {base_:%d-%b-%Y}). " if base_ is not None else "Your WSA delivery rows (not cancelled) with an ETD whose BL# is not on PPC In_Transit. ")+\
             (f"Alert: ARRIVES AFTER STOCK-OUT = expected arrival after the material's first stock-zero day; ETD PASSED; SHIPS WITHIN {ETD_SOON} DAYS; ON PLAN. "
              "Expected arrival = your WSA plant date; ETD passed = base date + planned transit (plant date - ETD) if it ships today. "
              "Not in the projections (they use PPC In_Transit). "+where)
    ws['A2'].font=F(size=9,color='5D6D7E')
    ws['A3']='   '.join(f"{lab_}: {cnt.get(k_,0)}" for k_,(lab_,_,_) in enumerate(ETD_FLAGS)); ws['A3'].font=F(size=10,bold=True,color='1F2D3D')
    H=['Sr.','Alert','ETD vs today','Status','Planning sheet','Code','Item name','ALERT','First stock-zero day','PO','Shipper','Quantity','UOM','ETD',
       'WSA plant date','Planned transit (days)','Expected arrival','Delay vs WSA plant date (days)','BL#','Remarks (as written)']
    for j_,h_ in enumerate(H,1):
        c=ws.cell(4,j_,h_); c.font=F(size=10,bold=True,color='1F2D3D'); c.fill=PatternFill('solid',fgColor='F4F6F8')
        c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True); c.border=Border(bottom=Side(style='medium',color='8A949E'))
    ws.row_dimensions[4].height=32
    CH={0:('C00000','FFFFFF'),1:('E67E22','FFFFFF'),2:('F4D03F','1F2D3D'),3:('5DADE2','FFFFFF'),4:('1E8449','FFFFFF'),5:('8A949E','FFFFFF')}
    for i_,(r,code,name,cls,first,uom_,fl) in enumerate(info):
        rr=5+i_
        when=(f"Late {r['late']} d" if r['passed'] else ('Today' if r['late']==0 else f"In {-r['late']} d"))
        v=[i_+1,ETD_FLAGS[fl][0],when,r['status'],r['sheet'],code,name,(ALERT_SHORT[cls] if cls is not None and cls<len(ALERT_SHORT) else '-'),
           (first.to_pydatetime() if first is not None else 'none'),r['po'] or '-',r['shipper'] or '-',r['qty'],uom_ or '',r['etd'].to_pydatetime(),
           r['plant'].to_pydatetime(),r['transit'],r['eta_now'].to_pydatetime(),r['delay'],r['bl'] or '-',r['remarks'] or None]
        for j_,x_ in enumerate(v,1):
            c=ws.cell(rr,j_,x_); c.font=F(size=10); c.border=BD; c.alignment=Alignment(vertical='center',horizontal=('left' if j_ in (4,5,7,11,19,20) else 'center'))
            if hasattr(x_,'year'): c.number_format='dd-mmm-yy'
            if j_==12: c.number_format='#,##0'
        c=ws.cell(rr,2); c.fill=PatternFill('solid',fgColor=ETD_FLAGS[fl][1]); c.font=F(size=9,bold=True,color=ETD_FLAGS[fl][2])
        if r['passed']: ws.cell(rr,3).font=F(size=11,bold=True,color=('C0392B' if r['late']>=14 else 'B9770E'))
        if cls is not None:
            bg,fg=CH.get(cls,('8A949E','FFFFFF')); c=ws.cell(rr,8); c.fill=PatternFill('solid',fgColor=bg); c.font=F(size=9,bold=True,color=fg)
        if fl==0: ws.cell(rr,17).font=F(size=10,bold=True,color='C0392B')     # arrives after the stock-out
        if r['status'].startswith('BL#'): ws.cell(rr,4).font=F(size=10,color='9A7D0A')
        ws.row_dimensions[rr].height=18
    if not rows: ws.cell(5,1,'Every WSA delivery row with an ETD is on PPC In_Transit (or cancelled).').font=F(size=10)
    for j_,w_ in enumerate([5,24,11,30,30,15,26,14,12,13,24,11,6,10,11,9,11,11,18,30],1): ws.column_dimensions[L(j_)].width=w_
    ws.freeze_panes='D5'; ws.auto_filter.ref=f"A4:T{4+max(1,len(rows))}"
    return len(rows)
def add_alert_inbound(path, rows, d0):
    """Inbound file: the ETD alert sheet. NO ALERT sheets in the Inbound file (user 21-Nov-2026: "Daily inbound, delete all sheet alert"); the ALERT is on 1_MRP."""
    wb=load_workbook(path)
    try: write_etd_late_sheet(wb,ETD_LATE,'This is the Inbound copy (the same rows as 1_MRP).')
    except Exception as ex_: import traceback; traceback.print_exc(); print(f'NOTE: ETD late sheet not added to the Inbound file ({type(ex_).__name__}: {ex_})')
    for w_ in wb.worksheets: w_.sheet_view.tabSelected=False
    wb.active=0; wb.worksheets[0].sheet_view.tabSelected=True
    _safe_save(wb,path)

def cover_watch(x, d0, ships, nd, need=None):
    """ONE calculation for the Risk Charts list AND the Action List columns 'Below Lead Time From' / 'Short From (Stock = 0)' / 'Days Late'.
       Below Lead Time From = first day projected stock (rolled forward + PPC in transit by arrival date - plan) covers fewer than lead-time days of need.
       Short From = first day projected stock is below zero. Days Late = arrival of the first shipment that comes AFTER the below-lead-time day, minus that day.
       Returns None when the material has no demand (never shown as a shortage) or no Usage/Day."""
    if x.get('no_demand'): return None
    b=x['b']; TGT=tgt_days(b.get('lt_used')); HZ=min(nd,ALERT_HORIZON); m=b['mat']; S=x.get('S') or []
    need=need or x.get('pl_need')                            # default: module plan Usage/Day; the Action List passes the daily-target need
    if not need or not S: return None
    cov=[v/need for v in S[:HZ]]
    t=next((k for k in range(len(cov)) if cov[k]<TGT),None)
    ri=risk_info(x,d0)
    out=dict(need=need,cov=cov,cov0=cov[0],t=t,ri=ri,short=ri['date'],below=None,late=None,nx=None,cause=None,gap_all=0.0,ov=[],ovq=0.0)
    dn_=vp_dead(m)[2]                                                         # Settings ALLOWED VENDORS: shipments no customer may use are not supply
    sh=[s_ for s_ in (ships or {}).get(m,[]) if s_.get('raw') is not None and not (dn_ and dn_(s_.get('sup')))]
    out['ov']=[s_ for s_ in sh if s_['raw']<d0]; out['ovq']=sum(s_['qty'] for s_ in out['ov'])
    if t is None: return out
    dt_=d0+pd.Timedelta(days=t); out['below']=dt_
    later=sorted([s_ for s_ in sh if s_['raw']>dt_],key=lambda s_:s_['raw'])
    gap=TGT*need-S[t]; on_way=sum(s_['qty'] for s_ in later)
    out['gap_all']=max(0.0,TGT*need-min(S[t:HZ]))                     # deepest gap in the 16 weeks (back to the lead time); arrivals already in S
    out['cause']='NO SUPPLY' if not later else ('LATE ARRIVAL' if on_way>=gap else 'LATE + NOT ENOUGH')
    out['nx']=later[0] if later else None
    out['late']=(later[0]['raw']-dt_).days if later else None
    return out

SHIPS_={}
HEAT_NAME='BOM Material Dashboard - Weekly'     # the weekly cover heatmap sheet (was 'Cover Heatmap', then 'BOM Material Dashboard'; renamed by the user, Nov 2026)
def bom_day_cover(r):
    """Days of cover of an ALERT / heatmap row: one value per day = HOW MANY DAYS THE STOCK LASTS FROM THAT DAY (user 23-Nov-2026: "make the
       week box mean how many days the stock lasts from that day"; the old rule = that morning's stock / that single day's demand, copied from WSA
       column V, gave 357 days for RTV on a 3.6 DR holiday). Stock that morning (incl. that day's arrival) is used up day by day on the PLANNED
       demand of the following days (holidays / ramps count as they are planned), no later arrivals; the last day is a fraction; after the last day of the plan the rest of the stock runs at Usage/Day
       (empty days after the plan end do not count as days that last). 'na' = no demand that day. Below 1 = the day's own demand is not covered
       (stock-zero day, same as before). STOCK NEVER GOES BELOW 0 (user 23-Nov-2026): demand lost on a stock-zero day is not owed."""
    P=list(r.get('P') or []); S=list(r.get('S') or [])
    n=len(S); Pn=[(P[k] if k<len(P) else 0.0) or 0.0 for k in range(n)]
    last=max([k for k,v in enumerate(Pn) if v>1e-9],default=-1)                  # last day with planned demand: empty days AFTER the plan do not 'last'
    cum=[0.0]
    for v in Pn[:last+1]: cum.append(cum[-1]+v)
    pos=[v for v in Pn if v>1e-9]; tail=float(r.get('need') or 0.0)
    if tail<=1e-9: tail=(sum(pos[-28:])/len(pos[-28:])) if pos else 0.0           # beyond the plan: Usage/Day (lead-time average), else the last plan days
    out=[]; prev=None
    for k in range(n):
        pk=Pn[k]
        s_prev=S[k-1] if k>0 else (S[0]+(P[0] if P else 0.0)-((r.get('Tp') or [0.0])[0] or 0.0))
        mor=(prev if prev is not None else max(0.0,s_prev))+(S[k]-s_prev+pk)     # morning stock incl. that day's arrival
        prev=max(0.0,mor-pk)
        if pk<=1e-9: out.append('na'); continue
        tgt=cum[k]+mor
        if tgt>=cum[last+1]: out.append((last+1-k)+((tgt-cum[last+1])/tail if tail>1e-9 else 0.0))     # stock outlasts the plan: rest at Usage/Day
        else: out.append((bisect.bisect_right(cum,tgt)-1-k)+(tgt-cum[bisect.bisect_right(cum,tgt)-1])/Pn[bisect.bisect_right(cum,tgt)-1])
    return out
def bom_risk_key(r):
    """Highest risk first - the SAME rule as the RM Material Dashboards (user, Nov 2026: \"all dashboard must sort from highest to lowest
       risk\"): first day the daily cover goes below 1 (stock-zero day), then below RISK_DAYS (Settings, 5), then the lowest cover today; rows with no demand last."""
    vs=bom_day_cover(r); P=list(r.get('P') or [])
    nums=[v for v in vs if not isinstance(v,str)]
    if not nums: return (1,0,0,0,str(r.get('name')))
    neg=next((i for i,v in enumerate(vs) if not isinstance(v,str) and v<1),10**6)     # below 1 = the day's need is not covered = stock-zero day (checker finding: <0 missed 1-day stock-outs)
    lo=next((i for i,v in enumerate(vs) if not isinstance(v,str) and v<RISK_DAYS),10**6)
    d_=(float(r.get('stock') or 0.0)/P[0]) if (P and P[0]>1e-9) else 1e9
    return (0,neg,lo,d_,str(r.get('name')))
COUNT_ITEM={}      # PPC Count 'Item' per code (Cover Heatmap groups, MRP Daily label) - filled in build_new_mrp
# Consumables named by the user (11-Nov-2026: "These are Consumables: ..."): perforated tapes, cell fixation tapes, barcodes, blank back
# label. They win over the PPC Count 'Item' and over the older answers (tape = PM, barcode / back label = PM).
CONSUMABLE_CODES={'RMTAP00000018','RMTAP00000025','RMTTT00000018','RMTTT00000009','RMLAB00000097','RMLAB00000098','RMLAB00000121'}
def is_consumable(code):
    """True for a code on the user's consumables list, or the other roll length of one (BOM alias in either direction: Cell Fixation
       tape RMTTT00000010 <- RMTTT00000018)."""
    c=str(code or '').strip().upper()
    if not c: return False
    if c in CONSUMABLE_CODES: return True
    for a,m in (BOM_ALIAS[0] or {}).items():
        a_=str(a).strip().upper(); m_=str(m).strip().upper()
        if (m_==c and a_ in CONSUMABLE_CODES) or (a_==c and m_ in CONSUMABLE_CODES): return True
    return False
def count_item_of(code):
    """PPC Count 'Item' of a code; a code not on the Count sheet takes the Item of its alias code (other roll length, e.g. Cell
       Fixation tape RMTTT00000010 <- RMTTT00000018). A code on the user's consumables list = 'Consumables' whatever the Count says."""
    c=str(code or '').strip()
    if is_consumable(c): return 'Consumables'
    if COUNT_ITEM.get(c): return COUNT_ITEM[c]
    return next((COUNT_ITEM[a] for a,m in (BOM_ALIAS[0] or {}).items() if str(m).strip()==c and COUNT_ITEM.get(str(a).strip())),None)
def is_pm_like(code):
    """Packaging per your rule (user answers Nov 2026: tape = PM, barcode / back label = PM): PPC Count Item 'Packaging' (alias too) or RMLAB."""
    if is_consumable(code): return False                     # the user's consumables list (Nov 2026) wins: CS, not PM
    return str(count_item_of(code) or '').lower().startswith('packag') or str(code or '').upper().startswith('RMLAB')
def _short_lbl(b):
    """MRP Daily / Weekly column C: 'PM - Packaging' for a material your PPC Count calls Packaging (Kraft labels, barcodes, back label)
       or for RMLAB codes, when its own label says raw material / consumable (user, Nov 2026); other labels unchanged."""
    sh=b.get('short'); m=str(b.get('mat') or '').strip()
    if is_consumable(m) and (str(sh or '').lower().startswith(('cs','pm','raw material')) or not sh): return 'CS - Consumable'   # user's consumables list (Nov 2026)
    pk=is_pm_like(m)
    if pk and (str(sh or '').lower().startswith(('cs','raw material')) or not sh): return 'PM - Packaging'
    return sh
def count_items(cnt_hdr, cnt_rows):
    """PPC Count column 'Item' (Item Category) per material code: first non-empty, non-#N/A value."""
    hl=[' '.join(str(h or '').split()).lower() for h in (cnt_hdr or [])]
    im=hl.index('material') if 'material' in hl else None; ii=next((hl.index(k_) for k_ in ('item','item category') if k_ in hl),None); out={}   # header 'Item' or 'Item Category' (user, Nov 2026: "read Item Category as Item")
    if im is None or ii is None: return out
    for r in cnt_rows or []:
        if not r or im>=len(r) or ii>=len(r) or not r[im]: continue
        v=r[ii]
        if v in (None,'') or (isinstance(v,str) and v.startswith('#')): continue
        out.setdefault(' '.join(str(r[im]).split()),' '.join(str(v).split()))
    return out
def _heat_expand(rows, d0):
    """ALERT rows for the dashboards (every substitute is its own ALERT row, user 25-Nov-2026)."""
    return list(rows)
def _heat_grp(r):                                             # RM first, then PM, then CS (user, Nov 2026: "group rm material, then pm material then cs material")
    c_=str(r.get('code') or '').upper()                # group = your PPC Count 'Item' (user, Nov 2026: "refer to item category from count ppc"):
    it_=str(count_item_of(r.get('code')) or '').strip().lower()      # Packaging -> PM, Consumable -> CS, any other item -> RM (alias code's Item too)
    if is_consumable(c_): return 2                     # the user's consumables list (Nov 2026: tapes, barcodes, back label = CS)
    if c_.startswith('RMLAB'): return 1                # barcode / back label = PM whatever the Count says (user answer Nov 2026; = is_pm_like, MRP Daily label)
    if it_: return 1 if it_.startswith('packag') else (2 if it_.startswith('consum') else 0)
    return 0 if c_.startswith('RM') else (1 if c_.startswith('PM') else 2)
def add_cover_heatmap(path, rows, d0, nd, items=None):
    """Cover Heatmap: the SAME rows as the ALERT sheet (Risk Charts list merged into ALERT, Nov 2026), weekly projected cover vs 21 days.
       Sorted from high to low risk (user request Nov 2026): alert class (stock-out inside the lead time > stock-out later > under the lead time >
       under the lead time later), then the number of weeks out of stock, then the weeks under the lead time, then the first stock-zero day."""
    from openpyxl.styles import Border, Side
    wb=post_wb(path); HEAT_SUBNOTE.clear(); HEAT_NEED.clear()
    for n_ in ('Risk Charts','Cover Heatmap','BOM Material Dashboard',HEAT_NAME):
        if n_ in wb.sheetnames: del wb[n_]
    ws=wb.create_sheet(HEAT_NAME); ws.sheet_view.showGridLines=False; ws.sheet_properties.tabColor=DASH_TAB
    F=lambda **k: Font(name='Calibri',**k)
    def fill(c,col): c.fill=PatternFill('solid',fgColor=col)
    NW=16; r0=1; DARK=DASH_BG; TXT=DASH_TXT; MUT=DASH_MUT
    cL=2; cK=3; cC=4; cM=5; cS=5+len(HEAT_MS_COLS)      # material name in B, Category in C (= Material Status, user Nov 2026), alert class in D,
                                                            # E..J = Material Status columns (user, Nov 2026; filled by _heat_ms_fill), weeks from K
    ws.column_dimensions['A'].width=15; ws.column_dimensions['B'].width=44;     # A = Code (user, Nov 2026: "code and item name in all dashboard")
    ws.column_dimensions['C'].width=18; ws.column_dimensions['D'].width=15
    if not MIS_CAT and MIS_PATH_[0]: load_mis_cat(MIS_PATH_[0])
    for w in range(NW): ws.column_dimensions[L(cS+w)].width=7
    for j_,(h_,w_) in enumerate(HEAT_MS_COLS): ws.column_dimensions[L(cM+j_)].width=w_
    def weeks(r):                                            # the DAILY number on each week's date (user, Nov 2026: "01-Dec in the BOM dashboard
        v_=r.get('cv_sub') or bom_day_cover(r)               # user 25-Nov-2026: a material covered by a substitute shows the cover WITH it (italic + note)                                  # should be the same as 01-Dec in the RM dashboard") = cover that morning / that day's demand
        # week cell = cover on the week's first day WITH demand (same rule as 5_Web; holiday week start 25-Dec / 01-Jan no longer blank)
        return [next((x_ for x_ in v_[7*w:7*w+7] if not isinstance(x_,str) and x_ is not None),'na') if 7*w<len(v_) else 'na' for w in range(NW)] if any(not isinstance(x_,str) for x_ in v_) else []
    exp=_heat_expand(rows,d0); mgrp=_heat_grp
    GN=['RAW MATERIALS (RM)','PACKAGING MATERIALS (PM)','CONSUMABLES (CS)']
    srt=sorted(exp,key=lambda r:(mgrp(r),)+bom_risk_key(r))      # one risk rule for every dashboard (user, Nov 2026)
    rows_=[]                                                 # band row (str) before each group, then its materials in risk order
    for g_ in range(3):
        mem_=[r for r in srt if mgrp(r)==g_]
        if mem_: rows_.append(f"{GN[g_]} · {len(mem_)}"); rows_.extend(mem_)
    height=4+len(rows_)+5; HEAT_MS_ROWS.clear()
    for rr in range(r0,r0+height):
        for cc in range(1,cS+NW+1): fill(ws.cell(rr,cc),DARK)
    ws.cell(r0,cL,'Weekly cover of the ALERT materials vs their transit lead time').font=F(size=16,bold=True,color=TXT)
    ws.cell(r0+1,cL,f"Each column = that date's cover (days the stock lasts from that date on the planned demand) - the same number as the daily cover on that date: red = short (below 1 day) · orange / amber = below the material's lead time · blue = lead time or more · na = no demand that day. "
                    f'Grouped by your PPC Count Item: raw materials (RM), then packaging (PM), then consumables (CS); in each group highest risk on top (same rule as every dashboard): first day the daily cover goes below 1 (stock-zero day), then below 5 days, then the lowest cover today; no demand last. Same materials as the ALERT sheet; a substitute group (EPE 480 + EPE 420) is shown per material (own stock, plan, in transit).').font=F(size=9,color=MUT)
    c=ws.cell(r0+3,cC,'Alert'); c.font=F(size=8,color=MUT); c.alignment=Alignment(horizontal='center')
    c=ws.cell(r0+3,cK,'Category'); c.font=F(size=8,color=MUT); c.alignment=Alignment(horizontal='left',indent=1)
    c=ws.cell(r0+3,1,'Code'); c.font=F(size=8,color=MUT); c.alignment=Alignment(horizontal='left',indent=1)
    for j_,(h_,w_) in enumerate(HEAT_MS_COLS):
        c=ws.cell(r0+3,cM+j_,h_); c.font=F(size=8,color=MUT); c.alignment=Alignment(horizontal='center',vertical='bottom',wrap_text=True)
    ws.row_dimensions[r0+3].height=30
    c=ws.cell(r0+3,cL,'Item name'); c.font=F(size=8,color=MUT); c.alignment=Alignment(horizontal='right',indent=1)
    for w in range(NW):
        c=ws.cell(r0+3,cS+w,f"{(d0+pd.Timedelta(days=7*w)):%d-%b}"); c.font=F(size=8,color=MUT); c.alignment=Alignment(horizontal='center')
    bins=DAY_BINS()                                          # the daily colour rule (short = below 1 day), as the RM / daily dashboards
    def colr(v):
        for a_,b_,c_ in DAY_BINS(_ROW_LT[0]):
            if a_<=v<b_: return c_
        return DAY_NA
    bdk=Side(style='thin',color=DARK)
    for i,r in enumerate(rows_):
        rr=r0+4+i
        if isinstance(r,str):                                # group band
            c=ws.cell(rr,cL,r); c.font=F(size=9,bold=True,color='FFFFFF'); c.alignment=Alignment(horizontal='left',indent=1)
            for cc in range(1,cS+NW): fill(ws.cell(rr,cc),DASH_BAND)      # band from column A (Code column; checker finding)
            ws.row_dimensions[rr].height=16; continue
        _ROW_LT[0]=r.get('lt')                               # this row's colours: amber below / blue from its lead time
        nm=str(r['name'])+(f" ({str(r['where']).replace('Glass G12R Topcon ','')})" if r.get('where') else '')    # full item name (= ALERT; checker finding)
        c=ws.cell(rr,1,r.get('code')); c.font=F(size=8,color=MUT); c.alignment=Alignment(horizontal='left',indent=1)
        HEAT_MS_ROWS.append((rr,str(r.get('code') or '').strip(),r.get('where')))
        HEAT_NEED[(str(r.get('code') or '').strip().upper(),nm)]=r.get('need'); HEAT_NEED.setdefault((str(r.get('code') or '').strip().upper(),None),r.get('need'))
        c=ws.cell(rr,cL,nm); c.font=F(size=9,color=TXT); c.alignment=Alignment(horizontal='right',indent=1)
        c=ws.cell(rr,cK,item_category(r.get('code'))); c.font=F(size=8,color=MUT); c.alignment=Alignment(horizontal='left',indent=1)
        _,bg,fg,_=ALERT_CLS[r['cls']] if r['cls']<len(ALERT_CLS) else ('','1E8449','FFFFFF','')
        c=ws.cell(rr,cC,ALERT_SHORT[r['cls']] if r['cls']<len(ALERT_SHORT) else 'Not on ALERT'); fill(c,bg); c.font=F(size=8,bold=True,color=fg); c.alignment=Alignment(horizontal='center')
        c.border=Border(left=bdk,right=bdk,top=bdk,bottom=bdk)
        own_=[next((x_ for x_ in bom_day_cover(r)[7*w:7*w+7] if not isinstance(x_,str) and x_ is not None),None) for w in range(NW)] if r.get('cv_sub') else None
        sbn_=' / '.join(f"{x_['n']} {x_['c']}" for x_ in r.get('sub_x') or [])
        for w,v in enumerate(weeks(r)):
            if isinstance(v,str): cl=ws.cell(rr,cS+w,'na'); fill(cl,DAY_NA); cl.font=F(size=7,color=DAY_NA_TXT)
            else:
                cl=ws.cell(rr,cS+w,max(0,math.floor(round(v,1)))); fill(cl,colr(v)); cl.font=F(size=8,color='FFFFFF')      # whole days rounded DOWN = 5_Web
                if own_ and own_[w] is not None and math.floor(round(v,1))>math.floor(round(own_[w],1)):     # the substitute adds days this week: italic bold + a note with its own cover
                    cl.font=F(size=8,color='FFFFFF',italic=True,bold=True)
                    HEAT_SUBNOTE[(str(r.get('code') or '').strip().upper(),nm,w)]=f"Incl. substitute {sbn_}: {max(0,math.floor(round(v,1)))} days. Own stock only: {max(0,math.floor(round(own_[w],1)))} days."   # note added on the DMS sheet (_dms_info_capture; column moves drop comments)
            cl.alignment=Alignment(horizontal='center')
            cl.border=Border(left=bdk,right=bdk,top=bdk,bottom=bdk)
        ws.row_dimensions[rr].height=15
    rl=r0+4+len(rows_)+1
    ws.cell(rl,cS,"Cover, days the stock lasts from that morning:").font=F(size=9,color=TXT)
    labs=DAY_LABS()
    for k,c_ in enumerate([b_[2] for b_ in bins]+[DAY_NA]):
        c=ws.cell(rl+1,cS+2*k,labs[k]); fill(c,c_); c.font=F(size=8,bold=True,color=(DASH_NATXT if c_==DAY_NA else 'FFFFFF')); c.alignment=Alignment(horizontal='center')
    ws.cell(rl+3,cL,"Source: the ALERT projection (stock rolled forward to today + PPC in transit by arrival date − plan demand; back glass 333 per WSA sheet). One column per week = the daily number on that date (BOM Material Dashboard - Daily).").font=F(size=8,color=MUT)
    for rr in range(r0,r0+height): ws.row_dimensions[rr].height=ws.row_dimensions[rr].height or 16
    pos=wb.sheetnames.index('ALERT')+1 if 'ALERT' in wb.sheetnames else 0
    wb.move_sheet(HEAT_NAME,offset=pos-wb.sheetnames.index(HEAT_NAME))
# + Stock, SAP stock (MB52), Physical count = Material Status 'STORE · stock' (user, Nov 2026: "BOM Material Dashboard, add these")
HEAT_MS_COLS=[('Stock',11),('SAP stock (MB52)',11),('Physical count',11),('In transit',11),('Next arrival',9),('Usage/Day',9),('Short from',10),('Days of cover (stock only)',8),('Days of cover (+ all in transit)',9)]
HEAT_NEED={}        # (code, item name) / (code, None) -> Usage/Day of a BOM Weekly row (add_cover_heatmap -> ata_days in the DMS 'Arrival at plant' column)
HEAT_SUBNOTE={}     # (code, item name, week index) -> note of a week cell that includes a substitute's cover (add_cover_heatmap -> DMS comment)
HEAT_MS_ROWS=[]     # (row, code, WSA sheet) of the weekly dashboard - filled from the Material Status sheet by _heat_ms_fill
def _heat_ms_fill(wb):
    """BOM Material Dashboard - Weekly columns E..J = the Material Status values of the row's material (user, Nov 2026: "Add these from
       material status sheet into bom material dashboard: In transit, Next arrival, Usage/Day, Short from, Inv. Days, Days incl. in
       transit"). Copied cell by cell from the finished Material Status sheet (same value, same number format), matched by code; back glass
       333 by its WSA sheet ('(WSA sheet: ...)' in the Material Status name). Returns the number of rows filled."""
    if HEAT_NAME not in wb.sheetnames or 'Material Status' not in wb.sheetnames or not HEAT_MS_ROWS: return 0
    ms=wb['Material Status']; ws=wb[HEAT_NAME]
    hd={str(ms.cell(2,j).value or '').strip():j for j in range(1,ms.max_column+1)}
    jc=hd.get('Item code'); jn=hd.get('Item name')
    if not jc: return 0
    src={}
    for r_ in range(3,ms.max_row+1):
        c_=str(ms.cell(r_,jc).value or '').strip()
        if not re.match(r'^(RM|PM|CS)[A-Z]*\d',c_): continue
        m_=re.search(r'\(WSA sheet: ([^)]+)\)',str(ms.cell(r_,jn).value or '') if jn else '')
        src.setdefault((c_,m_.group(1).strip() if m_ else None),r_); src.setdefault((c_,None),r_)
    F=lambda **k: Font(name='Calibri',**k); TXT=DASH_TXT; n_=0
    for rr,code,where in HEAT_MS_ROWS:
        r_=src.get((code,str(where).strip() if where else None)) or src.get((code,None))
        if r_ is None: continue
        n_+=1
        for j_,(h_,w_) in enumerate(HEAT_MS_COLS):
            sj=hd.get(h_)
            if not sj: continue
            sc=ms.cell(r_,sj); c=ws.cell(rr,5+j_,sc.value); c.number_format=sc.number_format
            isd=isinstance(sc.value,(dt.datetime,dt.date))
            c.font=F(size=8,color=(DASH_RED if (h_=='Short from' and isd) else TXT),bold=(h_=='Short from' and isd))
            c.alignment=Alignment(horizontal='center',vertical='center')
    nt=next((r_ for r_ in range(1,4) if isinstance(ws.cell(r_,2).value,str) and ws.cell(r_,2).value.startswith('Each column')),None)
    if nt: ws.cell(nt,2).value=ws.cell(nt,2).value+" Columns Stock .. Days of cover (+ all in transit) = the Material Status sheet."
    return n_
DAY_NA='474D57'      # 'na' cell (no demand that day); was 1B2335 on the black page
# Dashboard background (user, Nov 2026: "Change all Dashboard background from black to better color"): light theme - soft off-white
# page, navy text, white-ish gaps between the heat cells; the cover colours (DAY_BINS) and the navy group bands stay.
# BLACK AGAIN (user, Nov 2026: "change background back to black as before") - the values before the light theme.
# SOFTER DARK (user, Nov 2026: "Background BOM Material Dashboard too dark"): charcoal 33383F instead of 12131C (light = rejected).
DAY_NA_TXT='C8CED6'; DASH_BAND='2B5A84'   # 'na' text 5.4:1 on DAY_NA; group band steel blue (navy 2C3E50 vanished on charcoal - checker)
DASH_BG='33383F'; DASH_TXT='E5E8EE'; DASH_MUT='AEB6BF'; DASH_TAB='33383F'; DASH_RED='FF8A80'; DASH_NATXT='FFFFFF'
_ROW_LT=[None]     # lead time of the dashboard row being painted (DAY_BINS of that row)
def DAY_BINS(lt=None):
    """Colour rule of the daily dashboards (legend row of BOM Material Dashboard - Daily): cover in days that morning; short = below 1 day.
       Amber ends and blue starts at the material's TRANSIT LEAD TIME (user rule, Nov 2026: every day threshold = lead time; was 21 days):
       short < 1, 1-7, 7-14, 14 - LT (below the lead time), LT - LT+14, LT+14 - LT+39, LT+39 + (LT at least 14)."""
    R=max(2,int(COV_RED_DAYS)); M=max(int(DASH_STEP2),R+1); T=max(int(DASH_MIN_LT),tgt_days(lt if lt is not None else _ROW_LT[0]),M)
    return [(-1e9,1,'D40000'),(1,R,'E8590C'),(R,M,'F08C00'),(M,T,'C9A227'),(T,T+int(DASH_BLUE1),'1A5A8E'),(T+int(DASH_BLUE1),T+int(DASH_BLUE2),'1E7BC4'),(T+int(DASH_BLUE2),1e9,'0EA5F5')]
def DAY_LABS():
    R=max(2,int(COV_RED_DAYS)); M=max(int(DASH_STEP2),R+1)
    return ['short (< 1)',f'1–{R}',f'{R}–{M}',f'{M} – lead time (LT)',f'LT – LT+{int(DASH_BLUE1)}',f'LT+{int(DASH_BLUE1)} – LT+{int(DASH_BLUE2)}',f'LT+{int(DASH_BLUE2)} +','na']

