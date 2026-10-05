"""GoodBelly case study, JASP-order analysis. Reads GoodBelly_JASP.csv (the file loaded into JASP) and reproduces,
step by step, every table and figure in GoodBelly_JASP_Final_Report.docx.
Run:  python3 jasp_report/analysis_jasp.py        (writes jasp_report/fig/*.png and jasp_report/results_jasp.json)
Order: 1 base model | 2 its assumptions | 3 log-level model | 4 its assumptions | 5 interaction screen (16 terms)
       6 stepwise final model + assumptions | 7 answers to Q3-Q5 (tests used in the report)"""
import os, json, warnings; warnings.filterwarnings('ignore')
import numpy as np, pandas as pd, statsmodels.api as sm, statsmodels.formula.api as smf
import matplotlib; matplotlib.use('Agg'); import matplotlib.pyplot as plt
from scipy import stats
from statsmodels.stats.stattools import durbin_watson
from statsmodels.stats.diagnostic import het_breuschpagan
from statsmodels.stats.outliers_influence import variance_inflation_factor as vif
from sklearn.model_selection import GroupKFold
from sklearn.linear_model import LinearRegression

HERE = os.path.dirname(os.path.abspath(__file__)); ROOT = os.path.dirname(HERE)
FIG = os.path.join(HERE, 'fig'); os.makedirs(FIG, exist_ok=True)
NAVY, RED, GREY = '#1f4e79', '#c0504d', '#7f7f7f'
plt.rcParams.update({'font.size': 9, 'axes.spines.top': False, 'axes.spines.right': False, 'figure.dpi': 170})

d = pd.read_csv(os.path.join(ROOT, 'GoodBelly_JASP.csv'))
V = ['Price', 'Rep', 'Endcap', 'Demo', 'D13', 'D45', 'Natural', 'Fitness']
I16 = [c for c in d.columns if '_x_' in c]
assert len(d) == 1386 and len(I16) == 16
R = {'n': len(d), 'stores': int(d.Store.nunique()), 'weeks': int(d.Date.nunique()), 'dates': [d.Date.min(), d.Date.max()]}

def coef_table(m, X=None):
    ci = m.conf_int(); sd = m.model.data.frame if hasattr(m.model.data, 'frame') else None
    out = {}
    for n in m.params.index:
        out[n] = dict(b=m.params[n], se=m.bse[n], t=m.tvalues[n], p=m.pvalues[n], lo=ci.loc[n, 0], hi=ci.loc[n, 1])
    return out

def vifs(m, names):
    X = sm.add_constant(d[names].astype(float)); return {v: float(vif(X.values, i + 1)) for i, v in enumerate(names)}

def checks(m, name, k):
    e = m.resid; inf = m.get_influence(); sr = inf.resid_studentized_internal; lev = inf.hat_matrix_diag; ck = inf.cooks_distance[0]
    bp = het_breuschpagan(e, m.model.exog); jb = stats.jarque_bera(e)
    out = dict(r2=m.rsquared, ar2=m.rsquared_adj, F=float(m.fvalue), rmse=float(np.sqrt(m.mse_resid)), dw=float(durbin_watson(e)),
               skew=float(stats.skew(e)), kurt=float(stats.kurtosis(e, fisher=False)), min_res=float(e.min()), max_res=float(e.max()),
               bp_p=float(bp[1]), jb_p=float(jb[1]), n_sr3=int((abs(sr) > 3).sum()), max_cook=float(ck.max()),
               lev_cut=2 * (k + 1) / len(d), n_lev=int((lev > 2 * (k + 1) / len(d)).sum()), max_lev=float(lev.max()),
               ss_reg=float(m.ess), ss_res=float(m.ssr), ss_tot=float(m.centered_tss), df_reg=int(m.df_model), df_res=int(m.df_resid))
    g = pd.DataFrame({'e': e, 'E': d.Endcap, 'R': d.Rep}).groupby(['E', 'R']).e.mean()
    out['res_by_group'] = {f'{int(a)}{int(b)}': float(v) for (a, b), v in g.items()}
    out['sr3_in_endcap_norep'] = int(((abs(sr) > 3) & (d.Endcap == 1) & (d.Rep == 0)).sum())
    top = np.argsort(-lev)[:20]; out['top20_lev_rep_x_endcap'] = int(d.Rep_x_Endcap.values[top].sum())
    return out

def fig_diag(m, fn, title):
    e = m.resid; z = (e - e.mean()) / e.std(); f = m.fittedvalues; fg, ax = plt.subplots(2, 2, figsize=(7.0, 5.4))
    ax[0, 0].scatter(f, e, s=5, alpha=.4, color=NAVY); ax[0, 0].axhline(0, color=RED, lw=1); ax[0, 0].set(title='Residuals vs. predicted', xlabel='Predicted value', ylabel='Residual')
    ax[0, 1].hist(z, bins=40, density=True, color=NAVY, alpha=.7); x = np.linspace(-7, 7, 200); ax[0, 1].plot(x, stats.norm.pdf(x), color=RED); ax[0, 1].set(title='Histogram of standardized residuals', xlabel='Standardized residual')
    n = len(z); p = (np.arange(1, n + 1) - .5) / n
    ax[1, 0].scatter(stats.norm.cdf(np.sort(z)), p, s=5, color=NAVY, alpha=.5); ax[1, 0].plot([0, 1], [0, 1], color=RED); ax[1, 0].set(title='Normal P-P plot', xlabel='Observed cumulative probability', ylabel='Expected cumulative probability')
    stats.probplot(z, plot=ax[1, 1]); ax[1, 1].get_lines()[0].set(markersize=2, color=NAVY); ax[1, 1].get_lines()[1].set_color(RED)
    ax[1, 1].set(title='Q-Q plot standardized residuals', xlabel='Theoretical quantiles', ylabel='Standardized residuals')
    fg.suptitle(title, fontsize=10, fontweight='bold'); fg.tight_layout(); fg.savefig(os.path.join(FIG, fn)); plt.close()

# ------------------------------------------------------------------ descriptives
desc = d[['Units', 'Price', 'Natural', 'Fitness']].agg(['mean', 'std', 'min', 'max']).T
R['desc'] = {k: dict(mean=float(v['mean']), sd=float(v['std']), min=float(v['min']), max=float(v['max'])) for k, v in desc.iterrows()}
R['dummy_counts'] = {v: int(d[v].sum()) for v in ['Rep', 'Endcap', 'Demo', 'D13', 'D45']}
R['n_endcap_stores'] = int(d[d.Endcap == 1].Store.nunique()); R['demo_overlap_rows'] = int(((d.Demo + d.D13 + d.D45) >= 2).sum())
R['groups'] = {f'{int(e)}{int(r)}': dict(n=int(len(g)), mean=float(g.Units.mean())) for (e, r), g in d.groupby(['Endcap', 'Rep'])}
R['endcap_norep_stores'] = sorted(d[(d.Endcap == 1) & (d.Rep == 0)].Store.unique().tolist())
R['units_skew'] = float(stats.skew(d.Units))

# ------------------------------------------------------------------ 1-2 base model
m1 = smf.ols('Units~' + '+'.join(V), d).fit()
R['m1'] = dict(coef=coef_table(m1), **checks(m1, 'm1', len(V)))
R['m1']['std_beta'] = {v: float(m1.params[v] * d[v].std() / d.Units.std()) for v in V}
R['m1']['vif'] = vifs(m1, V)
fig_diag(m1, 'diag_base.png', 'Base model (Units on all 8 variables): residual diagnostics')

# ------------------------------------------------------------------ 3-4 log-level model
ml = smf.ols('lnUnits~' + '+'.join(V), d).fit()
R['ml'] = dict(coef=coef_table(ml), **checks(ml, 'ml', len(V)))
R['ml']['pct'] = {v: float((np.exp(ml.params[v]) - 1) * 100) for v in V}
R['ml']['vif'] = vifs(ml, V)
fig_diag(ml, 'diag_log.png', 'Log-level model (ln Units on all 8 variables): residual diagnostics')

# ------------------------------------------------------------------ 5 interaction screen (each of the 16 added to the base model)
rows = []
for c in I16:
    m = smf.ols('Units~' + '+'.join(V) + '+' + c, d).fit(); f = m.compare_f_test(m1)
    rows.append(dict(term=c, nonzero=int((d[c] != 0).sum()), b=float(m.params[c]), t=float(m.tvalues[c]), p=float(m.pvalues[c]), dr2=float(m.rsquared - m1.rsquared), r2=float(m.rsquared)))
R['screen'] = sorted(rows, key=lambda r: -r['dr2'])
fg, ax = plt.subplots(figsize=(6.4, 4.0)); s = R['screen'][::-1]
ax.barh([r['term'].replace('_x_', ' × ') for r in s], [r['dr2'] * 100 for r in s], color=[RED if r['term'] == 'Rep_x_Endcap' else NAVY for r in s])
ax.set(xlabel='Increase in R² (percentage points) when the term is added to the base model', title='Screening the 16 interaction candidates, one at a time'); ax.tick_params(axis='y', labelsize=8)
fg.tight_layout(); fg.savefig(os.path.join(FIG, 'screen.png')); plt.close()
# why Endcap x Rep: group means + base-model residuals by group
lab = ['No endcap,\nno rep', 'No endcap,\nrep', 'Endcap,\nno rep', 'Endcap,\nrep']; keys = ['00', '01', '10', '11']
fg, ax = plt.subplots(1, 2, figsize=(7.0, 2.9))
mu = [R['groups'][k]['mean'] for k in keys]; ax[0].bar(lab, mu, color=[GREY, NAVY, RED, NAVY])
for i, k in enumerate(keys): ax[0].text(i, mu[i] + 12, f"{mu[i]:.0f}\n(n={R['groups'][k]['n']})", ha='center', fontsize=7.5)
ax[0].set(ylabel='Average weekly units', title='Raw averages by endcap and rep', ylim=(0, max(mu) * 1.25)); ax[0].tick_params(axis='x', labelsize=7.5)
rr = [R['m1']['res_by_group'][k] for k in keys]; ax[1].bar(lab, rr, color=[GREY, NAVY, RED, NAVY]); ax[1].axhline(0, color='black', lw=.7)
for i, v in enumerate(rr): ax[1].text(i, v + (10 if v >= 0 else -38), f'{v:+.0f}', ha='center', fontsize=8)
ax[1].set(ylabel='Mean residual of base model (units)', title='What the base model gets wrong', ylim=(-350, 190)); ax[1].tick_params(axis='x', labelsize=7.5)
fg.tight_layout(); fg.savefig(os.path.join(FIG, 'why_endcap_rep.png')); plt.close()

# ------------------------------------------------------------------ 6 stepwise (JASP: Method = Stepwise, p-in 0.05, p-out 0.10), 24 candidates
def stepwise(df, cands, yv, pin=.05, pout=.10):
    sel, path = [], []
    while True:
        best = None
        for v in [c for c in cands if c not in sel]:
            m = smf.ols(f'{yv}~' + '+'.join(sel + [v]), df).fit()
            if m.pvalues[v] < pin and (best is None or abs(m.tvalues[v]) > abs(best[1])): best = (v, m.tvalues[v])
        if best is None: break
        prev = smf.ols(f'{yv}~' + '+'.join(sel), df).fit().rsquared if sel else 0.0
        sel.append(best[0]); m = smf.ols(f'{yv}~' + '+'.join(sel), df).fit()
        path.append(dict(step=len(path) + 1, var=best[0], t=float(best[1]), fchg=float(best[1]) ** 2, p=float(m.pvalues[best[0]]), r2=float(m.rsquared), ar2=float(m.rsquared_adj), dr2=float(m.rsquared - prev), rmse=float(np.sqrt(m.mse_resid))))
        for v in list(sel):
            if m.pvalues[v] > pout: sel.remove(v); m = smf.ols(f'{yv}~' + '+'.join(sel), df).fit(); path.append(dict(removed=v))
    return sel, path
sel, path = stepwise(d, V + I16, 'Units'); R['stepwise'] = dict(selected=sel, path=path, n_candidates=len(V + I16))
assert sel == ['Rep_x_Endcap', 'Rep', 'D13', 'Demo', 'D45', 'Price'], sel
sel_main, _ = stepwise(d, V, 'Units'); R['stepwise_main_only'] = sel_main
sel_ln, path_ln = stepwise(d, V + I16, 'lnUnits'); mln = smf.ols('lnUnits~' + '+'.join(sel_ln), d).fit()
R['stepwise_ln'] = dict(selected=sel_ln, n_terms=len(sel_ln), **{k: v for k, v in checks(mln, 'mln', len(sel_ln)).items() if k in ('r2', 'ar2', 'dw', 'skew', 'kurt', 'bp_p', 'n_sr3')})

F = sel; m2 = smf.ols('Units~' + '+'.join(F), d).fit(); mrob = smf.ols('Units~' + '+'.join(F), d).fit(cov_type='cluster', cov_kwds={'groups': d.Store})
R['m2'] = dict(coef=coef_table(m2), **checks(m2, 'm2', len(F)))
R['m2']['std_beta'] = {v: float(m2.params[v] * d[v].std() / d.Units.std()) for v in F}
R['m2']['vif'] = vifs(m2, F); R['m2']['cluster'] = {v: dict(se=float(mrob.bse[v]), p=float(mrob.pvalues[v])) for v in m2.params.index}
fig_diag(m2, 'diag_final.png', 'Final stepwise model (with Endcap × Rep): residual diagnostics')
# out-of-sample check, whole stores held out
y = d.Units.values; g = d.Store.values
def cv(X): return float(np.sqrt(np.mean([((y[te] - LinearRegression().fit(X[tr], y[tr]).predict(X[te])) ** 2).mean() for tr, te in GroupKFold(10).split(X, y, g)])))
R['cv_rmse'] = dict(base=cv(d[V].values), final=cv(d[F].values))
# exhaustive-search result (run once by analysis.py in the repo root; not repeated here)
R['search_note'] = 'All 2-, 3- and 4-way products of the 8 variables (154 terms, 108 usable) give the same six terms (analysis.py).'

# ------------------------------------------------------------------ 7 tests used in the answers
T = lambda s: dict(F=float(m2.f_test(s).fvalue), p=float(m2.f_test(s).pvalue))
R['q3'] = dict(D13_eq_D45=T('D13=D45'), Demo_eq_D13=T('Demo=D13'), Demo_eq_D45=T('Demo=D45'),
               diff_D45_D13=dict(zip(['est', 'lo', 'hi'], [float(m2.params['D45'] - m2.params['D13'])] + [float(x) for x in m2.t_test('D45-D13=0').conf_int()[0]])),
               total=float(m2.params['Demo'] + 3 * m2.params['D13'] + 2 * m2.params['D45']))
ci = m2.conf_int(); R['q3']['total_note'] = 'Demo + 3*D13 + 2*D45'
R['q3']['base_units_mean_no_promo'] = float(d[(d.Rep == 0) & (d.Endcap == 0) & (d.Demo + d.D13 + d.D45 == 0)].Units.mean())
fe = smf.ols('Units~' + '+'.join(F) + '+Endcap', d).fit(); R['q4'] = dict(plain_endcap=dict(b=float(fe.params['Endcap']), p=float(fe.pvalues['Endcap'])),
    rep_and_endcap=float(m2.params['Rep'] + m2.params['Rep_x_Endcap']), blend=float(m2.params['Rep_x_Endcap'] * R['dummy_counts']['Rep'] * 0 + m2.params['Rep_x_Endcap'] * 36 / 53))
d['Reg'] = d.Region.fillna('NA'); d['WeekNo'] = d.Week
def pf(extra):
    mm = smf.ols('Units~' + '+'.join(F) + extra, d).fit(); f = mm.compare_f_test(m2); return dict(F=float(f[0]), p=float(f[1]), ar2=float(mm.rsquared_adj))
R['q5'] = dict(natural_fitness=pf('+Natural+Fitness'), region=pf('+C(Reg)'), price_sq=pf('+I(Price**2)'), week=pf('+WeekNo'), store=pf('+C(Store)'),
               other15={c: float(smf.ols('Units~' + '+'.join(F) + '+' + c, d).fit().pvalues[c]) for c in I16 if c != 'Rep_x_Endcap'})
R['elasticity'] = float(m2.params['Price'] * d.Price.mean() / d.Units.mean())
R['price_range'] = [float(d.Price.min()), float(d.Price.max())]
R['mean_units'] = float(d.Units.mean())
# base-model endcap coefficient is a blend of "with rep" and "without rep" weeks
R['blend_check'] = dict(b_endcap_m1=float(m1.params['Endcap']), expected=float(m2.params['Rep_x_Endcap'] * 36 / 53))

# ------------------------------------------------------------------ figures for the final model
items = [('Rep_x_Endcap', 'Endcap (in stores with a rep)'), ('Rep', 'Regional sales rep'), ('Demo', 'Demo week'), ('D13', 'Demo 1-3 weeks ago'), ('D45', 'Demo 4-5+ weeks ago')]
fg, ax = plt.subplots(1, 2, figsize=(7.0, 2.9), gridspec_kw={'width_ratios': [1.3, 1]})
bb = np.array([m2.params[k] for k, _ in items]); lo = np.array([ci.loc[k, 0] for k, _ in items]); hi = np.array([ci.loc[k, 1] for k, _ in items])
ax[0].errorbar(bb, range(5)[::-1], xerr=[bb - lo, hi - bb], fmt='o', color=NAVY, capsize=3); ax[0].set_yticks(range(5)[::-1]); ax[0].set_yticklabels([l for _, l in items], fontsize=8)
ax[0].axvline(0, color=GREY, lw=.8); ax[0].set(xlabel='Extra units per store-week (95% CI)', title='Effect of each lever (final model)')
kk = ['Demo', 'D13', 'D45']; bv = [m2.params[x] for x in kk]
ax[1].bar(['Demo\nweek', 'Weeks\n1-3 after', 'Weeks\n4-5+ after'], bv, color=NAVY, yerr=[[m2.params[x] - ci.loc[x, 0] for x in kk], [ci.loc[x, 1] - m2.params[x] for x in kk]], capsize=4)
for i, v in enumerate(bv): ax[1].text(i, v + 14, f'+{v:.0f}', ha='center', fontsize=8)
ax[1].set(ylabel='Extra units per store-week', title='Does the demo lift fade?', ylim=(0, 150)); ax[1].tick_params(axis='x', labelsize=8)
fg.tight_layout(); fg.savefig(os.path.join(FIG, 'levers.png')); plt.close()

json.dump(R, open(os.path.join(HERE, 'results_jasp.json'), 'w'), indent=1, default=float)
print('final:', ' '.join(f'{m2.params[k]:+.2f}*{k}' if k != 'Intercept' else f'{m2.params[k]:.2f}' for k in m2.params.index))
print('M1 R2 %.4f | LOG R2 %.4f | FINAL R2 %.4f DW %.2f' % (m1.rsquared, ml.rsquared, m2.rsquared, durbin_watson(m2.resid)))
