# stats_analyser.py (multi-sheet picker + condensed output)
# -*- coding: utf-8 -*-
"""
Stats Analyser (GUI-Friendly) — Mixed Effects + Descriptives + Column Picker + Forward-Fill + Pretty Params
---------------------------------------------------------------------------------------------------------
- File-open dialog if --input omitted.
- If the workbook has multiple sheets and --sheet is not provided, a drop-down lets you pick ONE sheet.
- Column picker to choose Grouping (ID), Dependent, Fixed effects, and Group-by.
- Forward-fills blank cells downward for the chosen ID and all chosen Fixed/Group-by columns
  (but NEVER the dependent variable).
- DESCRIPTIVES sheet includes: number of groups (n_groups), N per group, and
  Mean, Max, Median, Q25, Q75, Range, Std, Min, SEM in that order.
- Computes normality, fits MixedLM; falls back to OLS if needed.
- Exports a condensed workbook with exactly THREE sheets:
    1) DESCRIPTIVES
    2) NORMALITY
    3) LMM_PARAMS_PRETTY  (if OLS fallback is used, this still receives the pretty params)

Patched changes (summary):
- Formula-aware forward-fill and validation (supports terms like C(Group)).
- More permissive and logged repeated-measures check before deciding LMM vs OLS.
- Robust MixedLM optimizer fallback chain: ML + Nelder–Mead, then Powell, then L-BFGS.
- Sheet picker Combobox when the workbook has many sheets (choose exactly one).
- Condensed output workbook: DESCRIPTIVES, NORMALITY, LMM_PARAMS_PRETTY only.
"""

from __future__ import annotations

import argparse
import json
import logging
import os
import platform
import re
import sys
from dataclasses import dataclass, asdict
from typing import List, Optional, Tuple, Dict

import numpy as np
import pandas as pd

import scipy.stats as ss
import statsmodels.formula.api as smf

PANDAS_EXCEL_ENGINE = "openpyxl"

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox
    from tkinter import ttk
    _TK_OK = True
except Exception:
    _TK_OK = False


@dataclass
class RunConfig:
    input_path: Optional[str]
    sheet: Optional[str]
    output_path: Optional[str]
    dep: Optional[str]
    fixed: List[str]
    group: Optional[str]
    random: Optional[str]
    by: List[str]
    dropna: bool = True
    alpha: float = 0.05

    @classmethod
    def from_args(cls, args: argparse.Namespace) -> "RunConfig":
        fixed = [c for c in (args.fixed or []) if c]
        by = [c for c in (args.by or []) if c]
        return cls(
            input_path=args.input,
            sheet=args.sheet,
            output_path=args.output,
            dep=args.dep,
            fixed=fixed,
            group=args.group,
            random=args.random,
            by=by,
            dropna=not args.keepna,
            alpha=args.alpha,
        )


def setup_logging(level: str = "INFO") -> None:
    logging.basicConfig(
        level=getattr(logging, level.upper(), logging.INFO),
        format="%(asctime)s | %(levelname)s | %(message)s",
        datefmt="%H:%M:%S",
    )


# -------------------- Simple file dialogs --------------------

def ask_open_excel() -> Optional[str]:
    if not _TK_OK:
        return None
    root = tk.Tk(); root.withdraw()
    path = filedialog.askopenfilename(
        title="Select input Excel file",
        filetypes=[("Excel files", ".xlsx .xls"), ("All files", "*.*")],
    )
    root.update(); root.destroy()
    return path or None


def ask_save_excel(default_name: str = "analysis_results.xlsx") -> Optional[str]:
    if not _TK_OK:
        return None
    root = tk.Tk(); root.withdraw()
    path = filedialog.asksaveasfilename(
        title="Save results as",
        defaultextension=".xlsx",
        initialfile=default_name,
        filetypes=[("Excel Workbook", ".xlsx")],
    )
    root.update(); root.destroy()
    return path or None


# -------------------- Workbook sheet discovery + picker --------------------

def list_sheet_names(xlsx_path: str) -> List[str]:
    """Return ordered list of sheet names without loading data."""
    x = pd.ExcelFile(xlsx_path, engine=PANDAS_EXCEL_ENGINE)
    return list(x.sheet_names)


def pick_sheet_gui(xlsx_path: str) -> Optional[str]:
    """
    Show a small GUI to pick ONE sheet from the workbook.
    Returns the selected sheet name, or None if cancelled.
    """
    if not _TK_OK:
        return None

    names = list_sheet_names(xlsx_path)
    if not names:
        return None
    if len(names) == 1:
        return names[0]

    class SheetPicker(tk.Toplevel):
        def __init__(self, master, options: List[str]):
            super().__init__(master)
            self.title("Select worksheet")
            self.resizable(False, False)
            self.result: Optional[str] = None

            frm = ttk.Frame(self, padding=10)
            frm.grid(sticky="nsew")
            ttk.Label(frm, text="Choose a sheet to analyse:").grid(row=0, column=0, sticky="w")
            self._var = tk.StringVar(value=options[0])
            cb = ttk.Combobox(frm, values=options, textvariable=self._var, state="readonly", width=40)
            cb.grid(row=1, column=0, pady=(6, 8), sticky="we")

            btns = ttk.Frame(frm)
            btns.grid(row=2, column=0, sticky="e")
            ttk.Button(btns, text="Cancel", command=self._on_cancel).pack(side=tk.RIGHT, padx=5)
            ttk.Button(btns, text="OK", command=self._on_ok).pack(side=tk.RIGHT)

            self.columnconfigure(0, weight=1)
            frm.columnconfigure(0, weight=1)

        def _on_ok(self):
            self.result = self._var.get()
            self.destroy()

        def _on_cancel(self):
            self.result = None
            self.destroy()

    root = tk.Tk()
    root.withdraw()
    dlg = SheetPicker(root, names)
    dlg.grab_set()
    root.wait_window(dlg)
    try:
        root.destroy()
    except Exception:
        pass
    return dlg.result


# -------------------- Data I/O --------------------

def read_excel(input_path: str, sheet: Optional[str]) -> pd.DataFrame:
    if not os.path.exists(input_path):
        raise FileNotFoundError(f"Input not found: {input_path}")
    logging.info(f"Reading Excel: {input_path} (sheet={sheet})")
    df = pd.read_excel(input_path, sheet_name=sheet, engine=PANDAS_EXCEL_ENGINE)
    if isinstance(df, dict):
        df = next(iter(df.values())) if sheet is None else df[sheet]
    return df


def write_excel(output_path: str, sheets: Dict[str, pd.DataFrame | str]) -> None:
    logging.info(f"Writing results to: {output_path}")
    with pd.ExcelWriter(output_path, engine=PANDAS_EXCEL_ENGINE) as writer:
        for name, obj in sheets.items():
            safe_name = str(name)[:31]
            if isinstance(obj, pd.DataFrame):
                obj.to_excel(writer, sheet_name=safe_name, index=False)
            else:
                pd.DataFrame({"_": [obj]}).to_excel(writer, sheet_name=safe_name, index=False, header=False)


# ---------- Analytics helpers ----------

# --- Patsy/formula helpers (patched) ---
_PLAIN_COL_RE = re.compile(r"^[A-Za-z_][A-Za-z0-9_]*$")

def is_plain_col(term: str) -> bool:
    """True if the term is a bare column name (no functions or operators)."""
    return bool(_PLAIN_COL_RE.fullmatch(str(term)))


def extract_cols_from_term(term: str) -> List[str]:
    """
    Extract column names referenced by a Patsy term.
    Handles C(col), I(col1*col2), np.log(col) etc. Best effort without importing patsy.
    """
    t = str(term)
    if is_plain_col(t):
        return [t]
    cols: set[str] = set()
    # Capture inside parentheses of functions like C(...), I(...), np.func(...)
    for m in re.finditer(r"[A-Za-z_][A-Za-z0-9_]*\s*\(\s*([^()]+?)\s*\)", t):
        inner = m.group(1)
        for tok in re.split(r"[^A-Za-z0-9_]+", inner):
            if tok and _PLAIN_COL_RE.fullmatch(tok):
                cols.add(tok)
    # Also capture plain identifiers separated by operators
    for tok in re.split(r"[^A-Za-z0-9_]+", t):
        if tok and _PLAIN_COL_RE.fullmatch(tok):
            cols.add(tok)
    return sorted(cols)


def compute_descriptives(df: pd.DataFrame, dep: str, by: Optional[List[str]] = None) -> pd.DataFrame:
    """
    Compute descriptive statistics for the dependent variable.
    If `by` is provided, report per-group rows and include `n_groups` (number of unique groups).
    Columns order (as requested):
      n_groups (if grouped), grouping columns, N, Mean, Max, Median, Q25, Q75, Range, Std, Min, SEM
    """
    if by:
        grp = df.groupby(by, dropna=False, observed=True)[dep]
        stats = grp.agg(
            N="count",
            Mean="mean",
            Max="max",
            Median="median",
            Q25=lambda x: np.nanpercentile(x, 25),
            Q75=lambda x: np.nanpercentile(x, 75),
            Range=lambda x: (np.nanmax(x) - np.nanmin(x)),
            Std="std",
            Min="min",
        ).reset_index()
        stats["SEM"] = stats["Std"] / np.sqrt(stats["N"].replace(0, np.nan))
        # Number of groups (unique combinations across all `by` columns)
        n_groups = df.dropna(subset=by).drop_duplicates(subset=by).shape[0]
        stats.insert(0, "n_groups", n_groups)
        ordered_cols = ["n_groups"] + by + ["N","Mean","Max","Median","Q25","Q75","Range","Std","Min","SEM"]
        return stats[ordered_cols]
    else:
        s = df[dep]
        out = {
            "N": [s.count()],
            "Mean": [s.mean()],
            "Max": [s.max()],
            "Median": [s.median()],
            "Q25": [np.nanpercentile(s, 25)],
            "Q75": [np.nanpercentile(s, 75)],
            "Range": [np.nanmax(s) - np.nanmin(s)],
            "Std": [s.std(ddof=1)],
            "Min": [s.min()],
        }
        df_out = pd.DataFrame(out)
        df_out["SEM"] = df_out["Std"] / np.sqrt(df_out["N"].replace(0, np.nan))
        return df_out["N Mean Max Median Q25 Q75 Range Std Min SEM".split()]


def normality_test(series: pd.Series, alpha: float = 0.05) -> Tuple[str, float, str]:
    x = series.dropna().astype(float)
    n = len(x)
    if n < 8:
        return ("n<8:skip", float("nan"), "Insufficient N")
    try:
        if 8 <= n <= 5000:
            stat, p = ss.shapiro(x)
            test = "Shapiro-Wilk"
        else:
            stat, p = ss.normaltest(x)
            test = "D’Agostino K²"
    except Exception:
        jb_stat, jb_p = ss.jarque_bera(x)
        test, p = "Jarque-Bera", float(jb_p)
    return (test, float(p), "PASS" if p >= alpha else "FAIL")


def compute_normality(df: pd.DataFrame, dep: str, by: Optional[List[str]], alpha: float) -> pd.DataFrame:
    rows = []
    t, p, f = normality_test(df[dep], alpha)
    rows.append({"Scope":"OVERALL","Test":t,"p_value":p,"Alpha":alpha,"Normality":f})
    if by:
        for keys, sub in df.groupby(by, dropna=False, observed=True):
            if not isinstance(keys, tuple):
                keys = (keys,)
            d = {by[i]: keys[i] for i in range(len(by))}
            t, p, f = normality_test(sub[dep], alpha)
            d.update({"Scope":"GROUP","Test":t,"p_value":p,"Alpha":alpha,"Normality":f})
            rows.append(d)
    out = pd.DataFrame(rows)
    for col in (by or []):
        if col not in out.columns:
            out[col] = np.nan
    return out[["Scope"] + (by or []) + ["Test","p_value","Alpha","Normality"]]


def format_pvalue(p: float) -> str:
    try:
        p = float(p)
    except Exception:
        return ""
    if p < 0.001:
        return "p < 0.001"
    return f"p = {p:.3f}"


def star_sig(p: float) -> str:
    try:
        p = float(p)
    except Exception:
        return ""
    if p < 0.001:
        return "***"
    if p < 0.01:
        return "**"
    if p < 0.05:
        return "*"
    return "ns"


# (patched) formula-aware constant-drop
def drop_constant_fixed(df: pd.DataFrame, fixed: List[str]) -> Tuple[List[str], List[str]]:
    keep, dropped = [], []
    for term in fixed:
        if term == "1":
            keep.append(term)
            continue
        cols = extract_cols_from_term(term)
        if not cols:
            keep.append(term)
            continue
        varies = False
        for c in cols:
            if c not in df.columns:  # can't verify transformed terms; keep
                varies = True
                break
            s = df[c]
            if s.nunique(dropna=True) > 1:
                varies = True
                break
        if varies:
            keep.append(term)
        else:
            dropped.append(term)
    return keep, dropped


def build_pretty_params_table(params_df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Return (wide_enhanced, stacked_long) with formatted p-values and significance stars.
    wide_enhanced: adds p_text, Sig, and Estimate_CI.
    stacked_long: columns [Term, Field, Value] with one row per metric.
    """
    df = params_df.copy()
    expected = {"Term","Estimate","StdErr","p","CI_lower","CI_upper"}
    for m in (expected - set(df.columns)):
        df[m] = np.nan
    df["p_text"] = df["p"].apply(format_pvalue)
    df["Sig"] = df["p"].apply(star_sig)
    def fmt_ci(row):
        try:
            return f"{row['Estimate']:.3f} ({row['CI_lower']:.3f}, {row['CI_upper']:.3f})"
        except Exception:
            return ""
    df["Estimate_CI"] = df.apply(fmt_ci, axis=1)
    wide_enhanced = df
    rows = []
    def z_or_t(r):
        return r['z'] if 'z' in r and pd.notna(r['z']) else r.get('t', np.nan)
    for _, r in df.iterrows():
        term = str(r.get('Term',''))
        rows.append({"Term":term,"Field":"Estimate","Value":f"{r['Estimate']:.6g}" if pd.notna(r['Estimate']) else ""})
        rows.append({"Term":term,"Field":"Std. Error","Value":f"{r['StdErr']:.6g}" if pd.notna(r['StdErr']) else ""})
        zv = z_or_t(r)
        rows.append({"Term":term,"Field":"z / t","Value":f"{zv:.6g}" if pd.notna(zv) else ""})
        if pd.notna(r['CI_lower']) and pd.notna(r['CI_upper']):
            rows.append({"Term":term,"Field":"95% CI","Value":f"({r['CI_lower']:.6g}, {r['CI_upper']:.6g})"})
        rows.append({"Term":term,"Field":"p-value","Value":r.get('p_text','')})
        rows.append({"Term":term,"Field":"Significance","Value":r.get('Sig','')})
    stacked_long = pd.DataFrame(rows, columns=["Term","Field","Value"])
    return wide_enhanced, stacked_long


# (patched) stronger MixedLM with robust optimizer chain
def run_mixedlm(df: pd.DataFrame, dep: str, fixed: List[str], group: str, re_formula: Optional[str]):
    if not fixed:
        raise ValueError("At least one fixed effect is required. Use --fixed 1 for intercept-only.")
    formula = f"{dep} ~ {' + '.join(fixed)}"
    logging.info(f"MixedLM: {formula}; groups={group}; re_formula={re_formula}")

    # Optional OLS warm start for fixed effects (not directly injected; helps sanity checks)
    try:
        ols_init = smf.ols(formula=formula, data=df).fit()
        _ = ols_init.params  # noqa: F841
        logging.info("Obtained OLS warm-start for fixed effects.")
    except Exception:
        pass

    model = smf.mixedlm(formula=formula, data=df, groups=df[group], re_formula=re_formula)

    fit = None
    last_err: Optional[Exception] = None
    for kwargs in (
        dict(reml=False, method="nm", maxiter=500, full_output=True, disp=False),
        dict(reml=False, method="powell", maxiter=300, full_output=True, disp=False),
        dict(reml=False, method="lbfgs"),
    ):
        try:
            fit = model.fit(**kwargs)
            logging.info(f"MixedLM try method={kwargs.get('method')} converged={getattr(fit, 'converged', None)}")
            if getattr(fit, 'converged', False):
                break
        except Exception as e:
            last_err = e
            logging.warning(f"MixedLM {kwargs.get('method')} failed: {type(e).__name__}: {e}")

    if fit is None or not getattr(fit, "converged", False):
        raise RuntimeError(f"MixedLM did not converge. Last error: {last_err}")

    params = fit.params
    bse = fit.bse
    zvals = params / bse
    from math import erf, sqrt
    def p_from_z(z):
        return 2*(1-0.5*(1+erf(abs(z)/sqrt(2))))
    pvals = zvals.apply(p_from_z)
    ci = fit.conf_int(); ci.columns = ["CI_lower","CI_upper"]
    params_df = pd.concat([
        params.rename("Estimate"), bse.rename("StdErr"), zvals.rename("z"), pvals.rename("p"), ci
    ], axis=1).reset_index().rename(columns={"index":"Term"})

    resid = fit.resid
    info = {
        "Model":"LMM",
        "AIC":fit.aic,
        "BIC":getattr(fit,'bic',float('nan')),
        "LogLik":fit.llf,
        "Converged":bool(fit.converged),
        "N_obs":int(fit.nobs),
        "Method":getattr(fit, 'method', 'ML'),
    }
    return str(fit.summary()), params_df, resid, info


def run_ols(df: pd.DataFrame, dep: str, fixed: List[str]):
    formula = f"{dep} ~ {' + '.join(fixed)}"
    logging.info(f"OLS: {formula}")
    fit = smf.ols(formula=formula, data=df).fit()
    params = fit.params.rename("Estimate").to_frame()
    params["StdErr"] = fit.bse
    params["t"] = fit.tvalues
    params["p"] = fit.pvalues
    ci = fit.conf_int(); ci.columns = ["CI_lower","CI_upper"]
    params = params.join(ci).reset_index().rename(columns={"index":"Term"})
    resid = fit.resid
    info = {"Model":"OLS","AIC":getattr(fit, 'aic', float('nan')), "BIC":getattr(fit, 'bic', float('nan')), "N_obs":int(fit.nobs)}
    return str(fit.summary()), params, resid, info


def build_readme(cfg: RunConfig, df: pd.DataFrame, fit_info: Optional[dict], ffilled_cols: List[str]) -> str:
    lines = []
    lines.append("Mixed-Effects Modeling & Descriptives Pipeline (GUI + Column Picker + Forward-Fill + Pretty Params)")
    lines.append("")
    lines.append("Run configuration:")
    lines.append(json.dumps(asdict(cfg), indent=2))
    lines.append("")
    lines.append("Environment:")
    lines.append(f"Python: {platform.python_version()}")
    lines.append("Packages: pandas, numpy, scipy, statsmodels, openpyxl, matplotlib")
    lines.append("")
    lines.append(f"Dataset: rows={len(df)}, cols={len(df.columns)}")
    if ffilled_cols:
        lines.append("")
        lines.append("Forward-fill applied to columns:")
        lines.append(", ".join(ffilled_cols))
    if fit_info:
        lines.append("")
        lines.append("Model fit info:")
        for k, v in fit_info.items():
            lines.append(f"- {k}: {v}")
    return "".join(lines)


_ID_PAT = re.compile(r"^(id|subject|animal|mouse|cell|patient|participant)([_ ]?id)?$", re.I)


def infer_candidates(df: pd.DataFrame) -> Dict[str, Optional[str]]:
    group_guess = None
    dep_guess = None
    cols = [c for c in df.columns if not str(c).startswith("Unnamed:")]

    # First try: ID-like column names
    for c in cols:
        if _ID_PAT.match(str(c)):
            nunq = df[c].nunique(dropna=True)
            if nunq < len(df):
                group_guess = c
                break

    # Fallback: choose a non-numeric column with moderate cardinality
    if group_guess is None:
        candidates = []
        n = len(df)
        for c in cols:
            if not pd.api.types.is_numeric_dtype(df[c]):
                k = df[c].nunique(dropna=True)
                if 2 <= k < n:
                    candidates.append((k, c))
        if candidates:
            group_guess = sorted(candidates, reverse=True)[0][1]

    # Numeric columns
    num_cols = list(df.select_dtypes(include="number").columns)
    if len(num_cols) == 1:
        dep_guess = num_cols[0]

    return {"group": group_guess, "dep": dep_guess}


class ColumnPicker(tk.Toplevel):
    def __init__(self, master, df: pd.DataFrame, preset_group: Optional[str], preset_dep: Optional[str]):
        super().__init__(master)
        self.title("Select Columns for Analysis")
        self.resizable(True, True)
        self.df = df
        self.result = None

        cols = list(df.columns)
        numeric_cols = list(df.select_dtypes(include="number").columns)
        other_cols = [c for c in cols if c not in numeric_cols]

        frm_top = ttk.Frame(self)
        frm_top.pack(fill=tk.BOTH, expand=False, padx=10, pady=(10, 5))
        ttk.Label(frm_top, text="Grouping (ID) column:").grid(row=0, column=0, sticky="w")
        self.group_var = tk.StringVar(value=preset_group or (other_cols[0] if other_cols else (cols[0] if cols else "")))
        self.group_cb = ttk.Combobox(frm_top, values=cols, textvariable=self.group_var, state="readonly")
        self.group_cb.grid(row=0, column=1, padx=8, pady=2, sticky="we")
        frm_top.columnconfigure(1, weight=1)

        ttk.Label(frm_top, text="Dependent variable (numeric):").grid(row=1, column=0, sticky="w")
        self.dep_var = tk.StringVar(value=preset_dep or (numeric_cols[0] if numeric_cols else (cols[0] if cols else "")))
        self.dep_cb = ttk.Combobox(frm_top, values=numeric_cols if numeric_cols else cols, textvariable=self.dep_var, state="readonly")
        self.dep_cb.grid(row=1, column=1, padx=8, pady=2, sticky="we")

        frm_mid = ttk.Frame(self)
        frm_mid.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        ttk.Label(frm_mid, text="Fixed effects (multi-select; optional):").grid(row=0, column=0, sticky="w")
        self.fixed_lb = tk.Listbox(frm_mid, selectmode=tk.MULTIPLE, exportselection=False, height=10)
        self.fixed_lb.grid(row=1, column=0, sticky="nsew")
        for c in [c for c in cols if c not in {self.group_var.get(), self.dep_var.get()}]:
            self.fixed_lb.insert(tk.END, c)
        sb1 = ttk.Scrollbar(frm_mid, orient=tk.VERTICAL, command=self.fixed_lb.yview)
        self.fixed_lb.configure(yscrollcommand=sb1.set)
        sb1.grid(row=1, column=1, sticky="ns")

        ttk.Label(frm_mid, text="Group-by for descriptives/normality (multi-select; optional):").grid(row=0, column=2, sticky="w", padx=(20,0))
        self.by_lb = tk.Listbox(frm_mid, selectmode=tk.MULTIPLE, exportselection=False, height=10)
        self.by_lb.grid(row=1, column=2, sticky="nsew", padx=(20,0))
        for c in cols:
            self.by_lb.insert(tk.END, c)
        sb2 = ttk.Scrollbar(frm_mid, orient=tk.VERTICAL, command=self.by_lb.yview)
        self.by_lb.configure(yscrollcommand=sb2.set)
        sb2.grid(row=1, column=3, sticky="ns")

        frm_mid.columnconfigure(0, weight=1)
        frm_mid.columnconfigure(2, weight=1)

        frm_btn = ttk.Frame(self)
        frm_btn.pack(fill=tk.X, padx=10, pady=(5,10))
        ttk.Button(frm_btn, text="Cancel", command=self._on_cancel).pack(side=tk.RIGHT, padx=5)
        ttk.Button(frm_btn, text="OK", command=self._on_ok).pack(side=tk.RIGHT)

        def refresh_fixed_options(*_):
            current_fixed = [self.fixed_lb.get(i) for i in self.fixed_lb.curselection()]
            self.fixed_lb.delete(0, tk.END)
            for c in cols:
                if c not in {self.group_var.get(), self.dep_var.get()}:
                    self.fixed_lb.insert(tk.END, c)
            for i in range(self.fixed_lb.size()):
                if self.fixed_lb.get(i) in current_fixed:
                    self.fixed_lb.selection_set(i)
        self.group_var.trace_add("write", refresh_fixed_options)
        self.dep_var.trace_add("write", refresh_fixed_options)

    def _on_ok(self):
        group = self.group_var.get().strip()
        dep = self.dep_var.get().strip()
        if not group or not dep:
            messagebox.showerror("Missing selection", "Please select both Group (ID) and Dependent variable.")
            return
        fixed = [self.fixed_lb.get(i) for i in self.fixed_lb.curselection()]
        by = [self.by_lb.get(i) for i in self.by_lb.curselection()]
        fixed = [c for c in fixed if c not in {group, dep}]
        by = [c for c in by if c]
        self.result = {"group": group, "dep": dep, "fixed": fixed, "by": by}
        self.destroy()

    def _on_cancel(self):
        self.result = None
        self.destroy()


def select_columns_gui(df: pd.DataFrame, cfg: RunConfig) -> Optional[Dict[str, List[str] | str]]:
    if not _TK_OK:
        logging.error("Tkinter not available — cannot open column picker. Provide --dep and --group via CLI.")
        return None
    guesses = infer_candidates(df)
    preset_group = cfg.group or guesses.get("group")
    preset_dep = cfg.dep or guesses.get("dep")

    root = tk.Tk(); root.withdraw()
    dlg = ColumnPicker(root, df, preset_group, preset_dep)
    dlg.grab_set()
    root.wait_window(dlg)
    return dlg.result


# (patched) formula-aware validation
def validate_columns_exist(df: pd.DataFrame, cols: List[str]) -> None:
    missing = []
    for c in cols:
        if c and is_plain_col(c) and c not in df.columns:
            missing.append(c)
    if missing:
        raise KeyError(f"Missing required columns in data: {missing}")


# (patched) decide LMM vs OLS with permissive repeats check and robust fitting
def run_mixed_or_ols(df: pd.DataFrame, cfg: RunConfig, fixed_terms: List[str]):
    grp_has_repeats = False
    if cfg.group in df.columns:
        vc = df[cfg.group].value_counts(dropna=False)
        grp_has_repeats = (len(vc) >= 3) and (vc.max() >= 2)
        logging.info(f"Grouping '{cfg.group}': n_groups={len(vc)}, min={vc.min() if len(vc) else 'NA'}, max={vc.max() if len(vc) else 'NA'}; repeats={grp_has_repeats}")

    if not grp_has_repeats:
        logging.warning("Grouping column has no/insufficient repeated measures; falling back to OLS.")
        summary_text, params_df, resid, fit_info = run_ols(df, cfg.dep, fixed_terms)
        model_kind = "OLS"
    else:
        try:
            summary_text, params_df, resid, fit_info = run_mixedlm(df, cfg.dep, fixed_terms, cfg.group, cfg.random)
            model_kind = "LMM"
        except Exception as e:
            logging.error(f"MixedLM failed ({type(e).__name__}: {e}); falling back to OLS.")
            summary_text, params_df, resid, fit_info = run_ols(df, cfg.dep, fixed_terms)
            model_kind = "OLS"
    return model_kind, summary_text, params_df, resid, fit_info


def run_and_collect(cfg: RunConfig) -> Tuple[Dict[str, pd.DataFrame | str], pd.DataFrame, dict]:
    df = read_excel(cfg.input_path, cfg.sheet)

    # Normalize blank strings to NaN early
    df = df.copy()
    if hasattr(df, 'map'):
        df = df.map(lambda v: (np.nan if isinstance(v, str) and v.strip() == "" else v))
    else:
        df = df.applymap(lambda v: (np.nan if isinstance(v, str) and v.strip() == "" else v))

    # Column picker
    need_picker = (cfg.dep is None) or (cfg.group is None)
    if need_picker or (not cfg.fixed and not cfg.by):
        sel = select_columns_gui(df, cfg)
        if not sel:
            raise RuntimeError("Column selection canceled.")
        cfg.group = sel["group"]
        cfg.dep = sel["dep"]
        if not cfg.fixed:
            cfg.fixed = sel["fixed"] or ["1"]
        if not cfg.by:
            cfg.by = sel["by"]

    # Validate (only plain column names enforced)
    validate_columns_exist(df, [cfg.dep, cfg.group] + [c for c in cfg.fixed if c != "1"] + cfg.by)

    # Forward-fill chosen structural columns (exclude dependent variable)
    ffilled_cols: List[str] = []
    cols_to_ffill: List[str] = []

    if cfg.group and cfg.group in df.columns:
        cols_to_ffill.append(cfg.group)

    for term in (cfg.fixed + cfg.by):
        if term and term != "1":
            for c in extract_cols_from_term(term):
                if c in df.columns and c != cfg.dep:
                    cols_to_ffill.append(c)

    cols_to_ffill = list(dict.fromkeys(cols_to_ffill))
    if cols_to_ffill:
        df.loc[:, cols_to_ffill] = df.loc[:, cols_to_ffill].ffill()
        ffilled_cols = cols_to_ffill.copy()
        logging.info(f"Forward-filled columns: {ffilled_cols}")

    # Drop NA rows in required cols
    if cfg.dropna:
        required_for_fit = [cfg.dep, cfg.group] + [c for c in cfg.fixed if c != "1"] + cfg.by
        n0 = len(df)
        df = df.dropna(subset=required_for_fit)
        logging.info(f"Dropped {n0 - len(df)} rows with NA in required columns.")

    # Dep numeric
    df = df.copy()
    df[cfg.dep] = pd.to_numeric(df[cfg.dep], errors="coerce")
    if cfg.dropna:
        n0 = len(df)
        df = df.dropna(subset=[cfg.dep])
        logging.info(f"Dropped {n0 - len(df)} rows with non-numeric {cfg.dep}.")

    # Descriptives & normality
    desc_df = compute_descriptives(df, cfg.dep, cfg.by if cfg.by else None)
    norm_df = compute_normality(df, cfg.dep, cfg.by if cfg.by else None, cfg.alpha)

    # Fixed effects cleanup (formula-aware)
    fixed_terms, dropped = drop_constant_fixed(df, cfg.fixed if cfg.fixed else ["1"])
    if dropped:
        logging.warning(f"Dropped constant fixed effect(s): {dropped}")
    if not fixed_terms:
        fixed_terms = ["1"]

    # Model
    model_kind, summary_text, params_df, resid, fit_info = run_mixed_or_ols(df, cfg, fixed_terms)

    # Pretty parameter tables (wide + stacked); only keep wide for condensed output
    wide_enh, stacked = build_pretty_params_table(params_df)

    # --------- CONDENSED OUTPUT (three sheets only) ---------
    sheets: Dict[str, pd.DataFrame | str] = {
        "DESCRIPTIVES": desc_df,
        "NORMALITY": norm_df,
        # Keep sheet name as requested even if the model is OLS for consistency:
        "LMM_PARAMS_PRETTY": wide_enh
    }
    # --------------------------------------------------------

    return sheets, df, fit_info


def parse_args(argv=None) -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Stats Analyser (MixedLM + Descriptives + GUI + Forward-Fill + Pretty Params)")
    p.add_argument("--input", default=None, help="Path to input .xlsx (use dialog if omitted)")
    p.add_argument("--sheet", default=None, help="Worksheet name; if omitted and the workbook has multiple sheets, a picker opens")
    p.add_argument("--output", default=None, help="Path to output .xlsx (ask after analysis if omitted)")
    p.add_argument("--dep", default=None, help="Dependent variable (numeric); if omitted, picker opens")
    p.add_argument("--fixed", type=lambda s: [x.strip() for x in s.split(",") if x.strip()], default=[], help="Comma-separated fixed effects; use '1' for intercept only")
    p.add_argument("--group", default=None, help="Grouping (ID) column; if omitted, picker opens")
    p.add_argument("--random", default=None, help="Random-effects formula for re_formula (e.g., '1 + time')")
    p.add_argument("--by", type=lambda s: [x.strip() for x in s.split(",") if x.strip()], default=[], help="Comma-separated columns for per-group descriptives/normality")
    p.add_argument("--alpha", type=float, default=0.05, help="Significance level for normality")
    p.add_argument("--keepna", action="store_true", help="Keep rows with NA in required columns")
    p.add_argument("--log", default="INFO", help="Logging level")
    return p.parse_args(argv)


def main(argv=None) -> int:
    args = parse_args(argv)
    setup_logging(args.log)

    cfg = RunConfig.from_args(args)

    # 1) Input selection
    if not cfg.input_path:
        if not _TK_OK:
            logging.error("Tkinter not available. Please provide --input.")
            return 2
        chosen = ask_open_excel()
        if not chosen:
            logging.error("No input file selected.")
            return 2
        cfg.input_path = chosen

    # 2) Sheet selection
    try:
        if cfg.sheet is None:
            names = list_sheet_names(cfg.input_path)
            if not names:
                logging.error("Workbook has no sheets.")
                return 2
            if len(names) == 1:
                cfg.sheet = names[0]
                logging.info(f"Workbook has one sheet; selecting: {cfg.sheet}")
            else:
                if _TK_OK:
                    pick = pick_sheet_gui(cfg.input_path)
                    if not pick:
                        logging.error("No sheet selected.")
                        return 2
                    cfg.sheet = pick
                else:
                    cfg.sheet = names[0]
                    logging.warning(f"Tkinter not available; defaulting to first sheet: {cfg.sheet}")
        else:
            # Validate given sheet exists
            names = list_sheet_names(cfg.input_path)
            if cfg.sheet not in names:
                logging.error(f"Sheet '{cfg.sheet}' not found. Available: {names}")
                return 2
    except Exception:
        logging.exception("Failed while enumerating workbook sheets")
        return 2

    # 3) Run analysis
    try:
        sheets, df, fit_info = run_and_collect(cfg)
    except Exception:
        logging.exception("Analysis failed")
        return 2

    # 4) Output selection
    if not cfg.output_path:
        default_name = f"{cfg.dep}_mixed_model_results.xlsx" if cfg.dep else "analysis_results.xlsx"
        if not _TK_OK:
            logging.error("Tkinter not available. Please provide --output.")
            return 2
        out = ask_save_excel(default_name)
        if not out:
            logging.error("No output path selected.")
            return 2
        cfg.output_path = out

    # 5) Write results
    try:
        write_excel(cfg.output_path, sheets)
        logging.info(f"Done. Results written to: {cfg.output_path}")
        return 0
    except Exception:
        logging.exception("Failed to write output Excel")
        return 2


if __name__ == "__main__":
    sys.exit(main())