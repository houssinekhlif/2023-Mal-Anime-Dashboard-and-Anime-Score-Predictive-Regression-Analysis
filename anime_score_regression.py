"""
Anime Score Prediction -- Multiple Linear Regression Analysis
================================================================
Dataset: MyAnimeList (MAL) Kaggle dataset, pre-processed with one-hot
encoded categorical features (Type, Source, Rating, Genre, Studio).

This script documents and reproduces the full model-building process
end to end, so every number in the report can be traced back to code
that actually ran, rather than to a claim:

  1. Feature engineering (log transforms, missing-value imputation)
  2. Why Popularity / Favorites / Scored By are excluded as predictors
  3. Baseline regression on all 60 engineered predictors
  4. Backward elimination on statistical significance (p < 0.05)
  5. Stepwise selection using information criteria (AIC, BIC) as a
     robustness check against the significance-based selection
  6. Multicollinearity diagnostics (VIF)
  7. Side-by-side comparison of every candidate model
  8. Final coefficient table

Run with: pip install pandas numpy statsmodels openpyxl
          python anime_score_regression.py
"""

import numpy as np
import pandas as pd
import statsmodels.api as sm
from statsmodels.stats.outliers_influence import variance_inflation_factor

FILE_PATH = "animesscoreprediction.xlsx"   # <-- update to your file path
SHEET_NAME = "Sheet1"


# ---------------------------------------------------------------
# 1. LOAD DATA
# ---------------------------------------------------------------
df = pd.read_excel(FILE_PATH, sheet_name=SHEET_NAME)

# Only rows with a known Score are usable for training
keep = df["Score"].notna()
d = df.loc[keep].copy()
print(f"Rows with a known Score: {len(d)} of {len(df)}")


# ---------------------------------------------------------------
# 2. FEATURE ENGINEERING
# ---------------------------------------------------------------
# Episodes and Members are extremely right-skewed count variables
# (raw skewness ~32.6 and ~7.6 respectively -- a few extreme outliers
# like long-running shows or mega-popular titles would otherwise
# dominate an OLS fit under squared-error loss). Log-transforming
# compresses that scale (skew drops to ~0.7 and ~0.5) and turns the
# coefficient into a "percent change in X -> point change in Score"
# semi-elasticity, which is a more natural read on a variable
# spanning several orders of magnitude.
ep_full = df["Episodes"].fillna(df["Episodes"].median())
dur_full = pd.to_numeric(df["Duration (min)"], errors="coerce")
dur_full = dur_full.fillna(dur_full.median())
yr_full = df["Premier Year"].fillna(df["Premier Year"].median()) - 2000

ep = ep_full.loc[keep].values
dur = dur_full.loc[keep].values
yr = yr_full.loc[keep].values
members = d["Members"].values


# ---------------------------------------------------------------
# 3. WHY Popularity / Favorites / Scored By ARE EXCLUDED
# ---------------------------------------------------------------
# These three, plus Members, are all "engagement" metrics computed
# from the same audience that also produces the Score. Two concrete
# problems with including all four as predictors:
#   (a) Reverse causality / leakage risk: a high Score plausibly
#       *causes* more favorites and a higher popularity ranking, not
#       just the other way around.
#   (b) Severe multicollinearity: Scored By and Members correlate at
#       0.989 on the raw scale (0.988 on the log scale) -- they are
#       effectively duplicate signals.
# Only Members is kept, as the least Score-contingent of the four.
# The two matrices below are the evidence for that decision.
engagement = d[["Popularity", "Favorites", "Scored By", "Members", "Score"]]
print("\n--- Engagement-cluster correlation matrix (raw scale) ---")
print(engagement.corr().round(3))

log_engagement = pd.DataFrame({
    "LN(Popularity)": np.log(d["Popularity"]),
    "LN(Favorites+1)": np.log(d["Favorites"] + 1),
    "LN(Scored By)": np.log(d["Scored By"]),
    "LN(Members)": np.log(d["Members"]),
    "Score": d["Score"],
})
print("\n--- Engagement-cluster correlation matrix (log scale) ---")
print(log_engagement.corr().round(3))


# ---------------------------------------------------------------
# 4. BUILD THE FULL 60-PREDICTOR DESIGN MATRIX
# ---------------------------------------------------------------
type_cols = ["Type_Special", "Type_Movie", "Type_OVA", "Type_ONA", "Type_Music"]
source_TtoAA = [
    "Source_Visual novel", "Source_Novel", "Source_Original",
    "Source_Light novel", "Source_4-koma manga", "Source_Web manga",
    "Source_Game", "Source_Web novel",
]
source_other_radio = d["Source_Other"].values + d["Source_Radio"].values
source_ACtoAG = [
    "Source_Book", "Source_Mixed media", "Source_Picture book",
    "Source_Card game", "Source_Music",
]
rating_cols = [
    "Rating_PG-13 - Teens 13 or older", "Rating_R+ - Mild Nudity",
    "Rating_PG - Children", "Rating_G - All Ages",
    "Rating_Rx - Hentai", "Rating_null",
]
genre_cols = [
    "Genre_Action", "Genre_Adventure", "Genre_Avant Garde",
    "Genre_Award Winning", "Genre_Boys Love", "Genre_Comedy", "Genre_Drama",
    "Genre_Ecchi", "Genre_Erotica", "Genre_Fantasy", "Genre_Girls Love",
    "Genre_Gourmet", "Genre_Hentai", "Genre_Horror", "Genre_Mystery",
    "Genre_Romance", "Genre_Sci-Fi", "Genre_Slice of Life", "Genre_Sports",
    "Genre_Supernatural", "Genre_Suspense",
]
studio_BKtoBP = [
    "Studio_Toei Animation", "Studio_Sunrise", "Studio_J.C.Staff",
    "Studio_Madhouse", "Studio_Production I.G", "Studio_TMS Entertainment",
]
studio_BRtoBT = ["Studio_Studio Deen", "Studio_OLM", "Studio_Pierrot"]

cols_order = (
    ["LN(Episodes)", "Duration (min)", "Year (since 2000)", "LN(Members)"]
    + type_cols
    + source_TtoAA
    + ["Source_Other+Radio"]
    + source_ACtoAG
    + ["Source_null"]
    + rating_cols
    + genre_cols
    + studio_BKtoBP
    + studio_BRtoBT
)

X_parts = [np.log(ep), dur, yr, np.log(members)]
for c in type_cols:
    X_parts.append(d[c].values)
for c in source_TtoAA:
    X_parts.append(d[c].values)
X_parts.append(source_other_radio)
for c in source_ACtoAG:
    X_parts.append(d[c].values)
X_parts.append(d["Source_null"].values)
for c in rating_cols:
    X_parts.append(d[c].values)
for c in genre_cols:
    X_parts.append(d[c].values)
for c in studio_BKtoBP:
    X_parts.append(d[c].values)
for c in studio_BRtoBT:
    X_parts.append(d[c].values)

X = np.column_stack(X_parts).astype(float)
y = d["Score"].values.astype(float)
print(f"\nDesign matrix: {X.shape[0]} rows x {X.shape[1]} predictors")


def fit_ols(idx):
    """Fit OLS on a subset of predictor columns (by index) plus an intercept."""
    Xc = sm.add_constant(X[:, idx]) if len(idx) else np.ones((X.shape[0], 1))
    return sm.OLS(y, Xc).fit()


# ---------------------------------------------------------------
# 5. BASELINE MODEL -- ALL 60 PREDICTORS
# ---------------------------------------------------------------
full_idx = list(range(len(cols_order)))
model_full = fit_ols(full_idx)
print(
    f"\nFull model (60 predictors): R2={model_full.rsquared:.4f}  "
    f"AdjR2={model_full.rsquared_adj:.4f}  F={model_full.fvalue:.2f}  "
    f"n={int(model_full.nobs)}"
)


# ---------------------------------------------------------------
# 6. BACKWARD ELIMINATION ON SIGNIFICANCE (p < 0.05)
# ---------------------------------------------------------------
# True backward elimination is iterative: drop the single worst
# predictor, refit, recheck significance, and repeat until every
# remaining predictor clears the threshold -- not a one-shot drop of
# everything that looked insignificant in the first fit.
def backward_elimination_pvalue(idx, alpha=0.05):
    idx = list(idx)
    while True:
        model = fit_ols(idx)
        pvals = model.pvalues[1:]  # index 0 is the intercept
        worst_i = int(np.argmax(pvals))
        if pvals[worst_i] > alpha:
            idx.pop(worst_i)
        else:
            return idx, model


sig_idx, model_sig = backward_elimination_pvalue(full_idx)
print(
    f"\nSignificance-based model ({len(sig_idx)} predictors): "
    f"R2={model_sig.rsquared:.4f}  AdjR2={model_sig.rsquared_adj:.4f}  "
    f"F={model_sig.fvalue:.2f}"
)
dropped_sig = sorted(
    set(cols_order[i] for i in full_idx) - set(cols_order[i] for i in sig_idx)
)
print("Dropped:", dropped_sig)


# ---------------------------------------------------------------
# 7. STEPWISE SELECTION ON INFORMATION CRITERIA (AIC / BIC)
# ---------------------------------------------------------------
# AIC and BIC penalize model complexity differently from a raw
# p-value cutoff: AIC's implicit bar for keeping a variable is
# roughly p ~ 0.16 (more lenient than 0.05), while BIC penalizes each
# added parameter by log(n) -- for n=15,692 that's a much harsher
# tax. Running both directions (forward and backward) checks that
# the result isn't an artifact of which end you searched from.
def forward_selection(criterion="aic"):
    selected, remaining = [], list(range(len(cols_order)))
    cur_score = getattr(fit_ols(selected), criterion)
    while remaining:
        scored = [(getattr(fit_ols(selected + [c]), criterion), c) for c in remaining]
        best_score, best_c = min(scored)
        if best_score < cur_score - 1e-9:
            selected.append(best_c)
            remaining.remove(best_c)
            cur_score = best_score
        else:
            break
    return selected, cur_score


def backward_elimination_ic(criterion="aic"):
    selected = list(range(len(cols_order)))
    cur_score = getattr(fit_ols(selected), criterion)
    improved = True
    while improved and len(selected) > 1:
        scored = [
            (getattr(fit_ols([c for c in selected if c != cand]), criterion), cand)
            for cand in selected
        ]
        best_score, best_c = min(scored)
        if best_score < cur_score - 1e-9:
            selected.remove(best_c)
            cur_score = best_score
        else:
            improved = False
    return selected, cur_score


fwd_aic_idx, fwd_aic_score = forward_selection("aic")
bwd_aic_idx, bwd_aic_score = backward_elimination_ic("aic")
bwd_bic_idx, bwd_bic_score = backward_elimination_ic("bic")

print(f"\nForward AIC  ({len(fwd_aic_idx)} predictors): AIC={fwd_aic_score:.1f}")
print(f"Backward AIC ({len(bwd_aic_idx)} predictors): AIC={bwd_aic_score:.1f}")
print(f"Backward BIC ({len(bwd_bic_idx)} predictors): BIC={bwd_bic_score:.1f}")
print("Forward and backward AIC agree on the same set:", set(fwd_aic_idx) == set(bwd_aic_idx))


# ---------------------------------------------------------------
# 8. MODEL COMPARISON TABLE
# ---------------------------------------------------------------
comparison = pd.DataFrame(
    {
        "Method": ["Backward, p<0.05", "Forward, AIC", "Backward, AIC", "Backward, BIC"],
        "# Predictors": [len(sig_idx), len(fwd_aic_idx), len(bwd_aic_idx), len(bwd_bic_idx)],
        "Adj R2": [
            fit_ols(sig_idx).rsquared_adj,
            fit_ols(fwd_aic_idx).rsquared_adj,
            fit_ols(bwd_aic_idx).rsquared_adj,
            fit_ols(bwd_bic_idx).rsquared_adj,
        ],
    }
)
print("\n--- Model comparison ---")
print(comparison.round(4).to_string(index=False))


# ---------------------------------------------------------------
# 9. VIF FOR THE FINAL (SIGNIFICANCE-BASED) MODEL
# ---------------------------------------------------------------
Xc_final = sm.add_constant(X[:, sig_idx])
vif_table = pd.DataFrame(
    {
        "Variable": [cols_order[i] for i in sig_idx],
        "VIF": [
            variance_inflation_factor(Xc_final, j)
            for j in range(1, Xc_final.shape[1])
        ],
    }
).sort_values("VIF", ascending=False)
print("\n--- VIF, final model (top 10) ---")
print(vif_table.head(10).to_string(index=False))
print(f"Max VIF: {vif_table['VIF'].max():.2f}   Median VIF: {vif_table['VIF'].median():.2f}")


# ---------------------------------------------------------------
# 10. FINAL COEFFICIENT TABLE
# ---------------------------------------------------------------
coef_table = pd.DataFrame(
    {
        "Variable": [cols_order[i] for i in sig_idx] + ["Intercept"],
        "Coefficient": list(model_sig.params[1:]) + [model_sig.params[0]],
        "Std Error": list(model_sig.bse[1:]) + [model_sig.bse[0]],
        "t-stat": list(model_sig.tvalues[1:]) + [model_sig.tvalues[0]],
        "p-value": list(model_sig.pvalues[1:]) + [model_sig.pvalues[0]],
    }
)
print("\n--- Final coefficient table ---")
print(coef_table.round(4).to_string(index=False))

coef_table.to_csv("final_model_coefficients.csv", index=False)
comparison.to_csv("model_comparison.csv", index=False)
vif_table.to_csv("vif_table.csv", index=False)
print("\nSaved: final_model_coefficients.csv, model_comparison.csv, vif_table.csv")
