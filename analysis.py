"""GoodBelly case study: complete analysis, reproducible from the raw dataset.
Run:  python3 analysis_final.py      (needs 'Case Study Dataset.xlsx' in the same folder or one level up)
Steps: 1 load/prepare  2 Model 1  3 assumption checks  4 stepwise with interactions  5 exhaustive 2/3/4-way search
       6 final Model 2 + checks  7 Q3 demo tests  8 Q4 placement  9 Q5 refinements  10 figures + results.json"""
import os, json, itertools, warnings; warnings.filterwarnings('ignore')
import numpy as np, pandas as pd, statsmodels.api as sm, statsmodels.formula.api as smf
import matplotlib; matplotlib.use('Agg'); import matplotlib.pyplot as plt
from scipy import stats
from statsmodels.stats.stattools import durbin_watson
from statsmodels.stats.outliers_influence import variance_inflation_factor as vif
from sklearn.model_selection import GroupKFold
from sklearn.linear_model import LinearRegression
os.makedirs('fig',exist_ok=True)
plt.rcParams.update({'font.size':9,'axes.spines.top':False,'axes.spines.right':False,'figure.dpi':150}); C1,C2,C3='#1f4e79','#c0504d','#7f7f7f'
# ---------- 1. load and prepare
src=next(p for p in ['Case Study Dataset.xlsx','../Case Study Dataset.xlsx'] if os.path.exists(p))
d=pd.read_excel(src,sheet_name='Real Data').rename(columns={'Units Sold':'Units','Average Retail Price':'Price','Sales Rep':'Rep','Demo1-3':'D13','Demo4-5':'D45'})
d=d.sort_values(['Store','Date']).reset_index(drop=True)
V=['Price','Rep','Endcap','Demo','D13','D45','Natural','Fitness']; y=d.Units.values; R={}
R['n']=len(d); R['stores']=int(d.Store.nunique()); R['weeks']=int(d.Date.nunique()); R['n_endcap']=int(d.Endcap.sum()); R['n_endcap_stores']=int(d[d.Endcap==1].Store.nunique())
R['demo_overlap_rows']=int(((d.Demo+d.D13+d.D45)>=2).sum())
def tab(m): ci=m.conf_int(); return {n:dict(b=m.params[n],se=m.bse[n],t=m.tvalues[n],p=m.pvalues[n],lo=ci.loc[n,0],hi=ci.loc[n,1]) for n in m.params.index}
def summ(m): return dict(r2=m.rsquared,ar2=m.rsquared_adj,F=m.fvalue,se=float(np.sqrt(m.mse_resid)),dw=durbin_watson(m.resid),skew=stats.skew(m.resid),kurt=stats.kurtosis(m.resid,fisher=False))
# ---------- 2. Model 1 (all 8 variables, Q1)
m1=smf.ols('Units~'+'+'.join(V),d).fit(); R['model1']=dict(coef=tab(m1),**summ(m1))
R['stdbeta1']={v:m1.params[v]*d[v].std()/d.Units.std() for v in V}
Xm=sm.add_constant(d[V]); R['vif1']={v:vif(Xm.values,i+1) for i,v in enumerate(V)}
inf1=m1.get_influence(); R['model1']['n_sr3']=int((abs(inf1.resid_studentized_internal)>3).sum()); R['model1']['max_cook']=float(inf1.cooks_distance[0].max())
# ---------- 4. stepwise with 16 hand-picked interactions (p-in .05, p-out .10)
pairs=[('Rep','Endcap'),('Rep','Demo'),('Rep','D13'),('Rep','D45'),('Endcap','Demo'),('Endcap','D13'),('Endcap','D45'),('Price','Rep'),('Price','Endcap'),('Price','Demo'),('Price','D13'),('Price','D45'),('Natural','Rep'),('Fitness','Rep'),('Natural','Endcap'),('Fitness','Demo')]
dd=d.copy()
for a,b in pairs: dd[f'{a}_x_{b}']=dd[a]*dd[b]
def stepwise(df,cands,yv='Units',pin=.05,pout=.10):
    sel=[];path=[]
    while True:
        best=None
        for v in [c for c in cands if c not in sel]:
            m=smf.ols(f'{yv}~'+'+'.join(sel+[v]),df).fit()
            if m.pvalues[v]<pin and (best is None or abs(m.tvalues[v])>abs(best[1])): best=(v,m.tvalues[v])
        if best is None: break
        prev=sm.OLS(df[yv],sm.add_constant(df[sel])).fit().rsquared if sel else 0
        sel.append(best[0]); m=smf.ols(f'{yv}~'+'+'.join(sel),df).fit()
        path.append(dict(step=len(path)+1,var=best[0],t=float(best[1]),p=float(m.pvalues[best[0]]),r2=m.rsquared,ar2=m.rsquared_adj,dr2=m.rsquared-prev))
        for v in list(sel):
            if m.pvalues[v]>pout: sel.remove(v); m=smf.ols(f'{yv}~'+'+'.join(sel),df).fit()
    return sel,path
dd['RE']=dd.Rep*dd.Endcap; I16=[f'{a}_x_{b}' for a,b in pairs]; sel,path=stepwise(dd,V+I16); R['stepwise16']=dict(selected=sel,path=path)
sel_main,_=stepwise(dd,V); R['stepwise_main_only']=sel_main
# ---------- 5. exhaustive search: all 2-,3-,4-way products
X={'*'.join(c):np.prod([d[v].values for v in c],axis=0) for k in (2,3,4) for c in itertools.combinations(V,k)}
nz={n:int((x!=0).sum()) for n,x in X.items()}; cand={n:x for n,x in X.items() if nz[n]>=10}
R['search']=dict(n_terms={k:sum(1 for n in X if n.count('*')==k-1) for k in (2,3,4)},usable={k:sum(1 for n in cand if n.count('*')==k-1) for k in (2,3,4)},pools={})
for name,pool in [('main only',V),('main + 2-way',V+[n for n in cand if n.count('*')==1]),('main + 2,3-way',V+[n for n in cand if n.count('*')<=2]),('main + 2,3,4-way',V+list(cand))]:
    df2=pd.concat([d,pd.DataFrame({n.replace('*','_x_'):cand[n] for n in cand})],axis=1)
    names=[c.replace('*','_x_') for c in pool]; s,_=stepwise(df2,names); R['search']['pools'][name]=dict(pool_size=len(pool),selected=s)
# ---------- 6. final Model 2
d['RE']=d.Rep*d.Endcap; F='RE+Rep+D13+Demo+D45+Price'; m2=smf.ols('Units~'+F,d).fit(); R['model2']=dict(coef=tab(m2),**summ(m2))
R['stdbeta2']={v:m2.params[v]*d[v].std()/d.Units.std() for v in F.split('+')}
Xf=sm.add_constant(d[F.split('+')].astype(float)); R['vif2']={v:vif(Xf.values,i+1) for i,v in enumerate(F.split('+'))}
inf=m2.get_influence(); k=len(F.split('+')); sr=inf.resid_studentized_internal; lev=inf.hat_matrix_diag; cook=inf.cooks_distance[0]
R['model2']['outliers']=dict(n_sr3=int((abs(sr)>3).sum()),lev_cut=2*(k+1)/len(d),n_lev=int((lev>2*(k+1)/len(d)).sum()),max_cook=float(cook.max()))
B=d[['Rep','D13','Demo','D45','Price']].copy(); B['RE']=d.RE
cvrmse=lambda Xm,mk=LinearRegression:float(np.sqrt(np.mean([((y[te]-mk().fit(Xm[tr],y[tr]).predict(Xm[te]))**2).mean() for tr,te in GroupKFold(10).split(Xm,y,d.Store.values)])))
R['cv_rmse']=dict(model1=cvrmse(d[V].values),model2=cvrmse(B.values))
# ---------- 7. Q3: demo duration (partial F-tests)
T=lambda s:(float(m2.f_test(s).fvalue),float(m2.f_test(s).pvalue))
R['q3']=dict(D13_eq_D45=T('D13=D45'),Demo_eq_D13=T('Demo=D13'),Demo_eq_D45=T('Demo=D45'),total_units_per_demo=float(m2.params['Demo']+3*m2.params['D13']+2*m2.params['D45']))
# ---------- 8. Q4: placement
g=lambda q:d.query(q).Units; R['q4']=dict(groups={'no_endcap_no_rep':(len(g('Endcap==0 and Rep==0')),float(g('Endcap==0 and Rep==0').mean())),'no_endcap_rep':(len(g('Endcap==0 and Rep==1')),float(g('Endcap==0 and Rep==1').mean())),'endcap_no_rep':(len(g('Endcap==1 and Rep==0')),float(g('Endcap==1 and Rep==0').mean())),'endcap_rep':(len(g('Endcap==1 and Rep==1')),float(g('Endcap==1 and Rep==1').mean()))},
   endcap_stores_without_rep=sorted(d.query('Endcap==1 and Rep==0').Store.unique().tolist()))
fe=smf.ols('Units~'+F+'+Endcap',d).fit(); R['q4']['plain_endcap_added_back']=dict(b=float(fe.params['Endcap']),p=float(fe.pvalues['Endcap']))
d['Reg']=d.Region.fillna('NA')
# ---------- 9. Q5: refinements
def pf(extra): mm=smf.ols('Units~'+F+extra,d).fit(); f=mm.compare_f_test(m2); return dict(F=float(f[0]),p=float(f[1]),adj_r2=mm.rsquared_adj)
R['q5']=dict(natural_fitness=pf('+Natural+Fitness'),region=pf('+C(Reg)'),price_squared=pf('+I(Price**2)'),week_trend=pf('+I(Date.rank(method="dense"))'),store_dummies=pf('+C(Store)'))
mi=smf.ols('Units~'+'+'.join(V)+'+RE',d).fit(); f=mi.compare_f_test(m1); R['q5']['add_interaction_to_model1']=dict(F=float(f[0]),p=float(f[1]),r2=mi.rsquared)
ml=smf.ols('np.log(Units)~'+F,d).fit(); R['q5']['log_version']=dict(r2=ml.rsquared,kurt=float(stats.kurtosis(ml.resid,fisher=False)))
R['q5']['other_interactions_p']={c:float(smf.ols('Units~'+F+'+'+c,dd).fit().pvalues[c]) for c in I16 if c!='Rep_x_Endcap'}
R['elasticity_at_mean']=float(m2.params['Price']*d.Price.mean()/d.Units.mean())
# ---------- 10. figures + save
def fig_diag(m,fn,title):
    e=m.resid; z=(e-e.mean())/e.std(); f=m.fittedvalues; fg,ax=plt.subplots(2,2,figsize=(7.2,5.6))
    ax[0,0].scatter(f,e,s=5,alpha=.4,color=C1); ax[0,0].axhline(0,color=C2,lw=1); ax[0,0].set(title='Residuals vs fitted',xlabel='Fitted',ylabel='Residual')
    ax[0,1].hist(z,bins=40,density=True,color=C1,alpha=.7); x=np.linspace(-6,6,200); ax[0,1].plot(x,stats.norm.pdf(x),color=C2); ax[0,1].set(title='Histogram of std. residuals',xlabel='Std. residual')
    n=len(z); p=(np.arange(1,n+1)-.5)/n; ax[1,0].scatter(stats.norm.cdf(np.sort(z)),p,s=5,color=C1,alpha=.5); ax[1,0].plot([0,1],[0,1],color=C2); ax[1,0].set(title='Normal P-P plot',xlabel='Observed cum. prob',ylabel='Expected cum. prob')
    stats.probplot(z,plot=ax[1,1]); ax[1,1].get_lines()[0].set(markersize=2,color=C1); ax[1,1].get_lines()[1].set_color(C2); ax[1,1].set_title('Normal Q-Q plot')
    fg.suptitle(title,fontsize=10,fontweight='bold'); fg.tight_layout(); fg.savefig(fn); plt.close()
fig_diag(m1,'fig/diag_m1.png','Model 1 (all 8 variables, no interaction): residual diagnostics'); fig_diag(m2,'fig/diag_m2.png','Model 2 (stepwise, with Rep x Endcap): residual diagnostics')
fg,ax=plt.subplots(figsize=(6,3)); G=R['q4']['groups']; ks=['no_endcap_no_rep','no_endcap_rep','endcap_no_rep','endcap_rep']; v=[G[x][1] for x in ks]
ax.bar(['No endcap,\nno rep','No endcap,\nrep','Endcap,\nno rep','Endcap,\nrep'],v,color=[C3,C1,C2,C1]); [ax.text(i,x+12,f'{x:.0f}\n(n={G[k][0]})',ha='center',fontsize=8) for i,(x,k) in enumerate(zip(v,ks))]
ax.set(ylabel='Avg weekly units',title='Endcap pays off only where a regional rep is present',ylim=(0,max(v)*1.25)); fg.tight_layout(); fg.savefig('fig/endcap.png'); plt.close()
ci=m2.conf_int(); items=[('RE','Endcap (in rep stores)'),('Rep','Regional rep'),('Demo','Demo week'),('D13','Demo wk 1-3'),('D45','Demo wk 4-5+')]
fg,ax=plt.subplots(figsize=(5.6,3)); bb=[m2.params[k] for k,_ in items]; ax.errorbar(bb,range(5),xerr=[np.array(bb)-[ci.loc[k,0] for k,_ in items],[ci.loc[k,1] for k,_ in items]-np.array(bb)],fmt='o',color=C1,capsize=3)
ax.set_yticks(range(5)); ax.set_yticklabels([l for _,l in items]); ax.axvline(0,color='grey',lw=.8); ax.set(xlabel='Extra units per store-week (95% CI)',title='Model 2: effect of each lever'); fg.tight_layout(); fg.savefig('fig/coef.png'); plt.close()
kk=['Demo','D13','D45']; fg,ax=plt.subplots(figsize=(5.2,3)); bv=[m2.params[x] for x in kk]
ax.bar(['Demo week','Weeks 1-3\nafter','Weeks 4-5+\nafter'],bv,color=C1,yerr=[[m2.params[x]-ci.loc[x,0] for x in kk],[ci.loc[x,1]-m2.params[x] for x in kk]],capsize=4); [ax.text(i,p+6,f'+{p:.0f}',ha='center') for i,p in enumerate(bv)]
ax.set(ylabel='Extra units per store-week',title='Demo lift (Model 2, 95% CI)'); fg.tight_layout(); fg.savefig('fig/lift.png'); plt.close()
fg,ax=plt.subplots(figsize=(5.2,2.6)); ax.hist(d.Units,bins=40,color=C1,alpha=.8); ax.set(title='Weekly units per store',xlabel='Units',ylabel='Store-weeks'); fg.tight_layout(); fg.savefig('fig/dist.png'); plt.close()
w=d.groupby('Date').Units.mean(); fg,ax=plt.subplots(figsize=(5.6,2.8)); ax.plot(w.index,w.values,marker='o',color=C1); ax.set(title='Average weekly units per store over time',ylabel='Units'); fg.autofmt_xdate(); fg.tight_layout(); fg.savefig('fig/time.png'); plt.close()
json.dump(R,open('results.json','w'),default=float,indent=1)
print('MODEL 2: Units =',' '.join(f'{m2.params[k]:+.1f}*{k}' if k!='Intercept' else f'{m2.params[k]:.1f}' for k in m2.params.index)); print('R2 %.4f adjR2 %.4f DW %.2f skew %.2f kurt %.2f'%(m2.rsquared,m2.rsquared_adj,durbin_watson(m2.resid),stats.skew(m2.resid),stats.kurtosis(m2.resid,fisher=False)))
