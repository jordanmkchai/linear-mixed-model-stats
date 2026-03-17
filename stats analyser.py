
# stats_analyser.py
# -*- coding: utf-8 -*-
"""
Stats Analyser (GUI-Friendly) — Multi-Sheet + Mixed Effects + Descriptives + Column Picker + Forward-Fill + Pretty Params
-----------------------------------------------------------------------------------------------------------------------
- File-open dialog if --input omitted.
- NEW: If the workbook has multiple sheets, a dialog lets you pick ONE sheet or **All sheets**.
- One column-picker: selections (ID/Dep/Fixed/By) are applied to all selected sheets.
- Forward-fills blank cells downward for the chosen ID and all chosen Fixed/Group-by columns (NEVER the dependent).
- DESCRIPTIVES sheet: includes `n_groups`, then N, Mean, Max, Median, Q25, Q75, Range, Std, Min, SEM, grouped by your Group-by.
- MixedLM fit with OLS fallback; PRETTY/STACKED parameter tables with conventional p-values + stars.
- NEW: When multiple sheets are analysed, results are combined into the same output workbook:
    * Most tables are **stacked vertically** with a leading `Sheet` column.
    * `*_PARAMS_STACKED` are combined **horizontally** (columns per sheet) so you can compare at a glance.
    * Text outputs (README, SUMMARIES) are concatenated with clear `===== Sheet: name =====` headers.
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


# -------------------- File dialogs --------------------

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


# -------------------- Workbook utilities --------------------

def list_sheet_names(xlsx_path: str) -> List[str]:
    x = pd.ExcelFile(xlsx_path, engine=PANDAS_EXCEL_ENGINE)
    return list(x.sheet_names)


def choose_sheets_gui(xlsx_path: str) -> List[str]:
    """Ask the user to pick ONE sheet or All sheets. Returns a list of sheet names."""
    names = list_sheet_names(xlsx_path)
    if len(names) <= 1 or not _TK_OK:
        return names

    class SheetPicker(tk.Toplevel):
        def __init__(self, master, options: List[str]):
            super().__init__(master)
            self.title("Select sheet(s)")
            self.resizable(False, False)
            self.result: List[str] = []

            ttk.Label(self, text="Select one sheet or tick 'Analyse all sheets':").grid(row=0, column=0, columnspan=2, sticky="w", padx=10, pady=(10,5))
            self.sheet_var = tk.StringVar(value=options[0])
            self.combo = ttk.Combobox(self, values=options, textvariable=self.sheet_var, state="readonly", width=40)
            self.combo.grid(row=1, column=0, columnspan=2, padx=10, sticky="we")

            self.all_var = tk.BooleanVar(value=False)
            chk = ttk.Checkbutton(self, text="Analyse all sheets", variable=self.all_var, command=self._toggle)
            chk.grid(row=2, column=0, columnspan=2, padx=10, pady=(8,8), sticky="w")

            btnf = ttk.Frame(self)
            btnf.grid(row=3, column=0, columnspan=2, sticky="e", padx=10, pady=(0,10))
            ttk.Button(btnf, text="Cancel", command=self._on_cancel).pack(side=tk.RIGHT, padx=5)
            ttk.Button(btnf, text="OK", command=self._on_ok).pack(side=tk.RIGHT)

            self.columnconfigure(0, weight=1)

        def _toggle(self):
            self.combo.config(state="disabled" if self.all_var.get() else "readonly")

        def _on_ok(self):
            if self.all_var.get():
                self.result = names
            else:
                self.result = [self.sheet_var.get()]
            self.destroy()

        def _on_cancel(self):
            self.result = []
            self.destroy()

    root = tk.Tk(); root.withdraw()
    dlg = SheetPicker(root, names)
    dlg.grab_set(); root.wait_window(dlg)
    return dlg.result or []


# -------------------- Data I/O --------------------

def read_sheet(input_path: str, sheet: Optional[str]) -> pd.DataFrame:
    logging.info(f"Reading Excel: {input_path} (sheet={sheet})")
    df = pd.read_excel(input_path, sheet_name=sheet, engine=PANDAS_EXCEL_ENGINE)
    if isinstance(df, dict):
        df = next(iter(df.values()))
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

def compute_descriptives(df: pd.DataFrame, dep: str, by: Optional[List[str]] = None) -> pd.DataFrame:
    """
    If `by` is provided, report per-group rows and include `n_groups` (unique combinations across all `by`).
    Order: n_groups, (by...), N, Mean, Max, Median, Q25, Q75, Range, Std, Min, SEM
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
        return df_out[["N","Mean","Max","Median","Q25","Q75","Range","Std","Min","SEM"]]


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


def drop_constant_fixed(df: pd.DataFrame, fixed: List[str]) -> Tuple[List[str], List[str]]:
    keep, dropped = [], []
    for term in fixed:
        if term == "1":
            keep.append(term)
            continue
        s = df[term]
        nunq = s.nunique(dropna=True)
        if nunq <= 1 or (np.issubdtype(s.dtype, np.number) and (float(s.std(ddof=0)) == 0.0)):
            dropped.append(term)
        else:
            keep.append(term)
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


def run_mixedlm(df: pd.DataFrame, dep: str, fixed: List[str], group: str, re_formula: Optional[str]):
    if not fixed:
        raise ValueError("At least one fixed effect is required. Use --fixed 1 for intercept-only.")
    formula = f"{dep} ~ {' + '.join(fixed)}"
    logging.info(f"MixedLM: {formula}; groups={group}; re_formula={re_formula}")
    model = smf.mixedlm(formula=formula, data=df, groups=df[group], re_formula=re_formula)
    try:
        fit = model.fit(method="lbfgs")
    except Exception:
        fit = model.fit(method="powell", maxiter=200)
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
    info = {"Model":"LMM","AIC":fit.aic, "BIC":getattr(fit,'bic',float('nan')), "LogLik":fit.llf, "Converged":bool(fit.converged), "N_obs":int(fit.nobs)}
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


def build_readme(cfg: RunConfig, df: pd.DataFrame, fit_info: Optional[dict], ffilled_cols: List[str], sheet_name: Optional[str]=None) -> str:
    lines = []
    header = "Mixed-Effects Modeling & Descriptives Pipeline (Multi-Sheet + GUI + Forward-Fill + Pretty Params)"
    if sheet_name:
        header = f"===== Sheet: {sheet_name} =====" + header
    lines.append(header)
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
    dtypes = df.dtypes
    for c in cols:
        if _ID_PAT.match(str(c)):
            nunq = df[c].nunique(dropna=True)
            if nunq < len(df):
                group_guess = c
                break
    if group_guess is None:
        candidates = []
        n = len(df)
        for c in cols:
            if not np.issubdtype(dtypes[c], np.number):
                k = df[c].nunique(dropna=True)
                if 2 <= k < n:
                    candidates.append((k, c))
        if candidates:
            group_guess = sorted(candidates, reverse=True)[0][1]
    num_cols = [c for c in cols if np.issubdtype(dtypes[c], np.number)]
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
        dtypes = df.dtypes
        numeric_cols = [c for c in cols if np.issubdtype(dtypes[c], np.number)]
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
    dlg.grab_set(); root.wait_window(dlg)
    return dlg.result


# -------------------- One-sheet pipeline --------------------

def process_one_sheet(df: pd.DataFrame, cfg: RunConfig) -> Tuple[Dict[str, pd.DataFrame | str], pd.DataFrame, dict]:
    # Normalize blank strings to NaN early (prefer DataFrame.map if available)
    df = df.copy()
    if hasattr(df, 'map'):
        df = df.map(lambda v: (np.nan if isinstance(v, str) and v.strip() == "" else v))
    else:
        df = df.applymap(lambda v: (np.nan if isinstance(v, str) and v.strip() == "" else v))

    # Column picker if needed
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

    # Validate
    missing = [c for c in [cfg.dep, cfg.group] + [c for c in cfg.fixed if c != "1"] + cfg.by if c not in df.columns]
    if missing:
        raise KeyError(f"Missing required columns in data: {missing}")

    # Forward-fill chosen structural columns (exclude dependent)
    ffilled_cols: List[str] = []
    cols_to_ffill = []
    if cfg.group and cfg.group in df.columns:
        cols_to_ffill.append(cfg.group)
    for c in (cfg.fixed + cfg.by):
        if c != "1" and c in df.columns and c != cfg.dep:
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

    # Fixed effects cleanup
    fixed_terms, dropped = drop_constant_fixed(df, cfg.fixed if cfg.fixed else ["1"])
    if dropped:
        logging.warning(f"Dropped constant fixed effect(s): {dropped}")
    if not fixed_terms:
        fixed_terms = ["1"]

    # Model: LMM if valid groups; else OLS
    model_kind = "LMM"
    fit_info = None
    resid = None

    grp_has_repeats = False
    if cfg.group in df.columns:
        vc = df[cfg.group].value_counts(dropna=False)
        grp_has_repeats = (len(vc) < len(df)) and (vc.min() >= 2)

    if not grp_has_repeats:
        logging.warning("Grouping column has no (or insufficient) repeated measures; falling back to OLS.")
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

    # Residual diagnostics
    resid_df = pd.DataFrame({"Residual": resid})
    rdesc = compute_descriptives(resid_df.rename(columns={"Residual": cfg.dep}), cfg.dep, None)
    rtest, rp, rflag = normality_test(resid, cfg.alpha)
    rdiag = rdesc.copy(); rdiag["Test"]=rtest; rdiag["p_value"]=rp; rdiag["Normality"]=rflag

    # Base sheets dict for this sheet
    sheets: Dict[str, pd.DataFrame | str] = {}

    sheets["DATA_SNAPSHOT"] = df.head(50)
    sheets["DESCRIPTIVES"] = desc_df
    sheets["NORMALITY"] = norm_df

    if model_kind == "LMM":
        sheets["LMM_SUMMARY"] = summary_text
        sheets["LMM_PARAMS"]  = params_df
    else:
        sheets["OLS_SUMMARY"] = summary_text
        sheets["OLS_PARAMS"]  = params_df

    wide_enh, stacked = build_pretty_params_table(params_df)
    if model_kind == "LMM":
        sheets["LMM_PARAMS_PRETTY"] = wide_enh
        sheets["LMM_PARAMS_STACKED"] = stacked
    else:
        sheets["OLS_PARAMS_PRETTY"] = wide_enh
        sheets["OLS_PARAMS_STACKED"] = stacked

    sheets["RESIDUALS_SUMMARY"] = rdiag

    return sheets, df, fit_info


# -------------------- Multi-sheet orchestration --------------------

def combine_multisheet_results(per_sheet_results: List[Tuple[str, Dict[str, pd.DataFrame | str], pd.DataFrame, dict]], cfg: RunConfig) -> Dict[str, pd.DataFrame | str]:
    """Merge multiple per-sheet results into one workbook-like dict."""
    combined: Dict[str, pd.DataFrame | str] = {}

    # Collect data snapshots
    snap_rows = []
    for sheet, sheets, df, fit_info in per_sheet_results:
        tmp = df.head(50).copy()
        tmp.insert(0, "Sheet", sheet)
        snap_rows.append(tmp)
    if snap_rows:
        combined["DATA_SNAPSHOT"] = pd.concat(snap_rows, ignore_index=True)

    # Helper to vertical-stack with Sheet column
    def vstack(name_variants: List[str]) -> None:
        rows = []
        for sheet, sheets, df, fit_info in per_sheet_results:
            for nm in name_variants:
                if nm in sheets and isinstance(sheets[nm], pd.DataFrame):
                    t = sheets[nm].copy()
                    t.insert(0, "Sheet", sheet)
                    rows.append(t)
        if rows:
            combined[name_variants[0]] = pd.concat(rows, ignore_index=True)

    # Vertical stacks
    vstack(["DESCRIPTIVES"])  # includes Sheet
    vstack(["NORMALITY"])    # includes Sheet
    vstack(["RESIDUALS_SUMMARY"])  # includes Sheet

    # Params (vertical)
    vstack(["LMM_PARAMS"])  # only sheets that had LMM
    vstack(["OLS_PARAMS"])  # only sheets that had OLS
    vstack(["LMM_PARAMS_PRETTY"])  # vertical stack
    vstack(["OLS_PARAMS_PRETTY"])  # vertical stack

    # STACKED params — combine horizontally (columns per sheet)
    def hcombine_stacked(key: str, outname: Optional[str] = None):
        pivoted: List[Tuple[str, pd.DataFrame]] = []
        for sheet, sheets, df, fit_info in per_sheet_results:
            if key in sheets and isinstance(sheets[key], pd.DataFrame):
                st = sheets[key]
                # Expect columns: Term, Field, Value
                pv = st.pivot(index=["Term","Field"], columns=None, values="Value")  # start as single series
                # The above returns a Series; convert to DataFrame with this sheet as column
                pv = pv.to_frame(name=sheet)
                pivoted.append((sheet, pv))
        if not pivoted:
            return
        # Outer-join across sheets on (Term, Field)
        out = None
        for sheet, pv in pivoted:
            out = pv if out is None else out.join(pv, how="outer")
        out = out.reset_index()
        combined[outname or key] = out

    hcombine_stacked("LMM_PARAMS_STACKED")
    hcombine_stacked("OLS_PARAMS_STACKED")

    # SUMMARIES & README — concatenate text blocks with headers
    def concat_text_blocks(name_variants: List[str], outname: Optional[str] = None):
        blocks = []
        for sheet, sheets, df, fit_info in per_sheet_results:
            for nm in name_variants:
                if nm in sheets and isinstance(sheets[nm], str):
                    blocks.append(f"===== Sheet: {sheet} =====" + str(sheets[nm]).strip())
        if blocks:
            combined[outname or name_variants[0]] = "".join(blocks)

    concat_text_blocks(["LMM_SUMMARY"])  # combined LMM summaries if any
    concat_text_blocks(["OLS_SUMMARY"])  # combined OLS summaries if any

    # Combined README (per-sheet readme stitched)
    readme_blocks = []
    for sheet, sheets, df, fit_info in per_sheet_results:
        # Build a per-sheet mini README using same structure
        # Note: We don't have ffilled_cols per sheet at this stage; omit or infer from README text if previously built.
        readme_blocks.append(f"===== Sheet: {sheet} =====")
        readme_blocks.append(f"Dependent: {cfg.dep} | Group(ID): {cfg.group} | Fixed: {', '.join(cfg.fixed) if cfg.fixed else '1'} | By: {', '.join(cfg.by) if cfg.by else '-'}")
        readme_blocks.append(f"Rows used: {len(df)} | Cols: {len(df.columns)}")
    combined["README"] = "".join(readme_blocks)

    return combined


# -------------------- CLI & Main --------------------

def parse_args(argv=None) -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Stats Analyser (Multi-Sheet + MixedLM + Descriptives + GUI + Forward-Fill + Pretty Params)")
    p.add_argument("--input", default=None, help="Path to input .xlsx (use dialog if omitted)")
    p.add_argument("--sheet", default=None, help="Worksheet name, comma-separated list, or 'ALL' (dialog if omitted)")
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

    # Input selection
    if not cfg.input_path:
        if not _TK_OK:
            logging.error("Tkinter not available. Please provide --input.")
            return 2
        chosen = ask_open_excel()
        if not chosen:
            logging.error("No input file selected.")
            return 2
        cfg.input_path = chosen

    # Determine sheet set
    all_sheets = list_sheet_names(cfg.input_path)
    if cfg.sheet:
        if cfg.sheet.strip().upper() == "ALL":
            chosen_sheets = all_sheets
        elif "," in cfg.sheet:
            chosen_sheets = [s.strip() for s in cfg.sheet.split(",") if s.strip()]
        else:
            chosen_sheets = [cfg.sheet]
    else:
        chosen_sheets = choose_sheets_gui(cfg.input_path)
        if not chosen_sheets:
            # If dialog canceled or not available, default to first sheet
            chosen_sheets = [all_sheets[0]] if all_sheets else []
    chosen_sheets = [s for s in chosen_sheets if s in all_sheets]
    if not chosen_sheets:
        logging.error("No valid sheet(s) to analyse.")
        return 2

    # Column picker only once: use the first sheet's data to decide
    try:
        df_first = read_sheet(cfg.input_path, chosen_sheets[0])
        # Trigger picker if needed
        need_picker = (cfg.dep is None) or (cfg.group is None) or (not cfg.fixed and not cfg.by)
        if need_picker:
            _ = process_one_sheet(df_first, cfg)  # This will set cfg via picker but we discard result
    except Exception:
        logging.exception("Failed during initial column selection stage")
        return 2

    # Process each chosen sheet and collect
    per_sheet_results: List[Tuple[str, Dict[str, pd.DataFrame | str], pd.DataFrame, dict]] = []
    for sh in chosen_sheets:
        try:
            df = read_sheet(cfg.input_path, sh)
            sheets_dict, df_used, fit_info = process_one_sheet(df, cfg)
            per_sheet_results.append((sh, sheets_dict, df_used, fit_info))
        except Exception as e:
            logging.exception(f"Analysis failed for sheet '{sh}' — skipping.")
            continue

    if not per_sheet_results:
        logging.error("No sheets were successfully analysed.")
        return 2

    # Combine into one workbook
    combined = combine_multisheet_results(per_sheet_results, cfg)

    # Save file
    if not cfg.output_path:
        default_name = f"{cfg.dep}_multi_sheet_results.xlsx" if cfg.dep else "analysis_results.xlsx"
        if not _TK_OK:
            logging.error("Tkinter not available. Please provide --output.")
            return 2
        out = ask_save_excel(default_name)
        if not out:
            logging.error("No output path selected.")
            return 2
        cfg.output_path = out

    try:
        write_excel(cfg.output_path, combined)
        logging.info(f"Done. Results written to: {cfg.output_path}")
        return 0
    except Exception:
        logging.exception("Failed to write output Excel")
        return 2


if __name__ == "__main__":
    sys.exit(main())
