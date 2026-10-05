# GoodBelly: the analysis, built one step at a time (stepwise + interactions version)

Each step says **what we do**, **why**, **how (in Excel/JASP, as in class)** and **what we got**. Numbers come from our data (1,386 store-weeks). The order follows your syllabus, Sessions 1-8.

Variables: Units (Y), Price, Rep, Endcap, Demo, D13 (= Demo1-3), D45 (= Demo4-5), Natural, Fitness.

---

## Step 0: Understand the columns

- Units, Price, Natural and Fitness are numbers. Rep, Endcap, Demo, D13 and D45 are **dummy variables** (1 = yes, 0 = no).
- Demo, D13 and D45 mark **how recently a store ran a demo** (this week / 1-3 weeks ago / 4-5+ weeks ago). "No recent demo" (all three = 0) is the **base**, so a dummy's coefficient means "the difference from no recent demo, holding other variables fixed" (Session 5-6). They can overlap: 34 store-weeks have two or more flags set (repeat demos), and in those rows the effects simply add.

## Step 1: Look at the data

**Excel:** Data Analysis → Descriptive Statistics, plus a histogram.

- Units: mean 254, SD 111, range 48-1,041, right-skewed.
- Only 53 store-weeks (3.8%) have an endcap, in 12 stores. This turns out to matter a lot.
- Raw averages: demo 401 vs 245 (no demo); rep 299 vs 199 (no rep); endcap 584 vs 241 (no endcap).
- These raw gaps overlap (a store can have a demo, an endcap and a rep at once), so we need regression to separate the effects (Session 4: "control").

## Step 2: Model 1, a plain linear regression with all 8 variables

**Model:** Units = β0 + β1·Price + β2·Rep + β3·Endcap + β4·Demo + β5·D13 + β6·D45 + β7·Natural + β8·Fitness + ε

**Why:** the natural first model, and the baseline we will test.

**Excel:** Data → Data Analysis → Regression. Y = Units, X = the 8 columns, tick Residuals, Residual Plots and Normal Probability Plot.

**We got:** R² = 0.673, adjusted R² = 0.671, Se = 63.7 units, F = 353.7 (p < 0.001).

| Variable | Coef | t | p |
|---|---|---|---|
| Price | -28.5 | -7.2 | <0.001 |
| Rep | +77.4 | 20.0 | <0.001 |
| Endcap | +305.1 | 33.7 | <0.001 |
| Demo | +111.1 | 15.0 | <0.001 |
| D13 | +73.5 | 15.0 | <0.001 |
| D45 | +67.6 | 10.3 | <0.001 |
| Natural | -1.6 | -0.9 | 0.37 |
| Fitness | -1.0 | -0.9 | 0.35 |

Read it as: "holding the others fixed, an endcap adds about 305 units a week." The F-test says the model is useful overall; each t-test (H₀: coefficient = 0) says which variables matter. Natural and Fitness don't.

## Step 3: Multicollinearity (VIF)

VIF = 1/(1 − Rⱼ²). Rule of thumb: a problem above 4 (Session 4). Ours are all between 1.02 and 1.26, so **no problem**.

## Step 4: Check the assumptions on Model 1 (Question 2)

| Assumption | Check | Model 1 result |
|---|---|---|
| Linearity, constant variance | Residuals vs fitted | **Fails.** Uneven spread, and a block of big negative residuals (down to about -370) at high fitted values |
| Independence | Durbin-Watson (≈ 2 is fine) | **Fails.** 1.39 (positive autocorrelation) |
| Normality | Histogram, P-P, Q-Q, skew, kurtosis | **Fails.** Skew -0.93, kurtosis 8.5, heavy left tail |

Lesson: a high R² (0.67) does not mean a valid model. The residual plot is telling us something is missing.

## Step 5: Build the model by stepwise regression, with interactions (Session 7 + 6)

**Why stepwise:** instead of choosing variables by eye, let the data pick them using the partial F-test (p-in 0.05, p-out 0.10).

**Why interactions:** an interaction term (product of two variables) lets the effect of one variable **depend on** the other (Session 6: dummy × variable). We added 16 candidates: Rep, Endcap and Price crossed with the demo stages, Rep × Endcap, and so on.

**How (Session 7 algorithm):**
1. Start empty. Add the variable with the largest partial correlation with Units, if its partial F is significant (p < 0.05).
2. Re-check variables already in (remove if p > 0.10).
3. Repeat until nothing enters or leaves.

**We got:**

| Step | Variable entered | t | R² | ΔR² |
|---|---|---|---|---|
| 1 | **Endcap × Rep** | 41.0 | 0.548 | 0.548 |
| 2 | Rep | 22.1 | 0.666 | 0.118 |
| 3 | D13 | 16.6 | 0.722 | 0.056 |
| 4 | Demo | 16.7 | 0.769 | 0.047 |
| 5 | D45 | 14.3 | 0.798 | 0.030 |
| 6 | Price | -7.3 | 0.806 | 0.007 |

Natural, Fitness, plain Endcap and the other 15 interactions never enter.

**The key discovery:** the Endcap × Rep interaction alone explains 55% of the variation, and the plain Endcap variable drops out. In other words, **an endcap only works in stores that also have a regional rep**.

Check it directly (average weekly units):

| Group | Store-weeks | Avg units |
|---|---|---|
| No endcap, no rep | 607 | 198 |
| No endcap, rep | 726 | 276 |
| Endcap, no rep | 17 | **218** |
| Endcap, rep | 36 | **757** |

The 17 "endcap without rep" weeks come from just three stores (Bethesda, Coral Gables, Sarasota). They sold almost nothing extra, and they were exactly the big negative residuals in Step 4.

**Excel:** create a column =Rep*Endcap, add it as a regressor, and use the partial F-test (compare SSE of the models with and without it) to confirm.

**Did we miss a better combination? The exhaustive check.** We repeated stepwise on every 2-, 3- and 4-way product of the 8 variables (28 + 56 + 70 = 154 terms; 108 usable after dropping terms non-zero in fewer than 10 rows). Result: **the same 6-term model in every pool.** Adding each of the 117 remaining terms to Model 2 one at a time, none is significant (smallest p = 0.09; about 6 would pass at 5% by chance). In 10-fold cross-validation (whole stores held out) Model 2 predicts best: RMSE 49.1 vs 68.0 for Model 1, 49.5 for lasso on all terms, 50.4 for gradient boosting, 52.6 for random forest. Putting all 108 terms in one OLS overfits (RMSE 305).

## Step 6: Model 2, the final stepwise model

**Units = 276.6 + 454.4·(Endcap × Rep) + 59.4·Rep + 106.8·Demo + 73.4·D13 + 74.6·D45 − 22.1·Price**

R² = 0.806, adjusted R² = 0.805, Se = 49.0, F = 954. All p < 0.001.

| Term | Meaning |
|---|---|
| Endcap × Rep = 454 | endcap adds about 454 units a week, but only where there is a rep |
| Rep = 59 | rep alone adds about 59 units a week |
| Demo = 107 | demo week adds about 107 units |
| D13 = 73, D45 = 75 | weeks after: about 73-75 units each week |
| Price = -22 | each extra $1 removes about 22 units |

Standardised betas (to compare levers): Endcap × Rep 0.65 > Rep 0.27 > D13 0.24 > Demo 0.23 > D45 0.18 > Price -0.09.

## Step 7: Re-check the assumptions on Model 2

| Check | Model 1 | Model 2 |
|---|---|---|
| Residuals vs fitted | uneven, big negative block | centred on zero, even |
| Durbin-Watson | 1.39 | **2.11** |
| Skew / kurtosis | -0.93 / 8.5 | **-0.01 / 3.05** (normal is 0 / 3) |
| Max VIF | 1.26 | 1.27 |
| |std resid| > 3 | 23 points | 5 points |
| Max Cook's D | 0.096 | 0.024 |

**Conclusion:** all assumptions are reasonably met. The problem with Model 1 was a missing interaction (omitted variable bias, Session 4), not the scale of sales.

**Why we don't use a log model anymore:** the same 6 terms on ln(units) give R² = 0.61 and kurtosis 5.8, which is worse. The earlier log fix was only hiding the missing interaction.

## Step 8: Outlier analysis (Session 8)

- Standardised residual beyond ±3: only 5 points.
- Leverage cut-off 2(k+1)/n = 0.010: the five highest-leverage points are all endcap-with-rep weeks (Pearl, Redmond, Redwood City, Roosevelt Square), because endcaps are rare.
- Cook's distance maximum 0.024, far below 1, so **no influential points**. Keep all data.

## Step 9: Do demos work, and for how long? (Question 3)

All three demo coefficients are positive, p < 0.001, so yes.

**Does the lift fade? Partial F-test:**
- D13 = D45: F = 0.04, **p = 0.85**, so no difference. No fade within five weeks.
- Demo = D13: F = 24.0, **p < 0.001**, so the demo week is higher than the afterglow.

**Excel way:** run the full model (SSE_full), then a reduced model with D13 + D45 combined into one column (SSE_red), and compute F = [(SSE_red − SSE_full)/1] / [SSE_full/(n−k−1)]. For one restriction this equals t².

**Payoff of one demo:** 107 + 3×73 + 2×75 ≈ **476 extra units over six weeks**. Only about 22% falls on demo day.

**Limit:** D45 is "4-5 weeks or more", so we only know the lift lasts at least five weeks.

## Step 10: Does placement matter? (Question 4)

- Endcap adds about 454 units a week **in rep stores** (95% CI 438 to 471), and about zero without a rep (p = 0.96 when plain Endcap is added back).
- Combined with a rep, the gain is about 59 + 454 = 513 units a week.
- Other interactions (Endcap × Demo, Rep × Demo, all Price interactions and so on): none significant (p between 0.15 and 0.95), so demos work the same with or without a rep or endcap.
- Region dummies: partial F p = 0.54, not significant.
- **Caution:** the no-rep result rests on only 3 stores and 17 store-weeks. Treat it as strong but worth confirming.

## Step 11: Refinements (Question 5)

| Refinement | Result |
|---|---|
| Stepwise with interactions | R² 0.67 → 0.81 |
| Add Natural + Fitness back (partial F) | p = 0.96, excluded |
| Add plain Endcap back | p = 0.96, excluded |
| Region dummies (partial F) | p = 0.54, not needed |
| Other 15 interactions | all insignificant |
| All 154 two-, three-, four-way products | same model selected; none of 117 remaining terms significant; best out-of-sample error |
| Log transformation | worse, not used |

Ideas beyond the data: store traffic and size (a likely omitted variable), competitor prices, holidays, week-by-week demo dummies, more endcap tests in non-rep stores, and a randomised pilot.

## Step 12: Recommendations (Question 6)

| Finding | Recommendation |
|---|---|
| Endcap +454 only with a rep | Pair every endcap with a regional rep. Extend rep coverage first. |
| Demo about 476 units, no fade in 5 weeks | Keep the demos and judge them over six weeks. |
| Rep +59 on its own | Add reps to uncovered stores. |
| Price elasticity about -0.36 | Don't use price cuts for volume: revenue falls. |
| Natural, Fitness, Region not significant | Target by store traffic, not neighbourhood. |

**Limitations:** observational data, 11 weeks, one retailer, few endcap stores, non-random assignment.

---

## One-minute version, in your own words

"A model with all eight variables fit well (R² 0.67) but failed the assumptions: autocorrelation and a heavy left tail. We used stepwise regression with interaction terms, and it picked an Endcap × Rep interaction first. That showed an endcap only works in stores that also have a regional rep: about +454 units a week there, nothing without one. With that term the model explains 81%, and every assumption is met. Demos add about 107 units in the demo week and about 75 a week after, with no fade in five weeks, worth about 476 units per demo. Price is inelastic and neighbourhood doesn't matter."

## Where each step sits in your slides

| Step | Session |
|---|---|
| 2 | 2, 3, 5 (regression output, t and F tests) |
| 3 | 4 (VIF) |
| 4, 7 | 3, 8 (assumptions, residual plots, Durbin-Watson) |
| 5, 6 | 6, 7 (interaction variables, stepwise, partial F) |
| 8 | 8 (outliers, leverage, Cook's) |
| 9, 10, 11 | 5, 6 (partial F, dummies) |
