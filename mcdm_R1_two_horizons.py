"""Baseline MCDM analysis for one respondent across two evaluation horizons.
Run: python mcdm_single_respondent.py
Dependencies: numpy, scipy, artifact_tool (for Excel input/output).
"""
from pathlib import Path
import numpy as np
from scipy.stats import rankdata
from artifact_tool import Blob, Workbook, SpreadsheetFile

# Change these settings for a different respondent or evaluation horizon.
INPUT = Path(__file__).with_name("R1_decision_matrices.xlsx")
SHEETS = {"2025-2035": "R1-2025-2035", "after-2035": "R1-after-2035"}
OUTPUT = Path(__file__).with_name("R1_results.xlsx")
PATHWAYS = ["Base Case"] + [f"P{i}" for i in range(1, 21)]
CRITERIA = [f"B{i}" for i in range(1, 19)]
METHODS = ["ARAS", "COPRAS", "MOORA", "TOPSIS", "VIKOR", "WASPAS"]
V_VIKOR = 0.5
LAMBDA_WASPAS = 0.5
TIE_RULE = "min"
ROUND = 10
EPS = 1e-12

# Read and validate the respondent's 21 x 18 decision matrix.
def load_matrix(path, sheet):
    wb = SpreadsheetFile.import_xlsx(Blob.load(str(path)))
    ws = wb.worksheets.get_item(sheet)
    rows = ws.get_range("A1:S22").values
    headers = [str(c).strip() for c in rows[0]]
    if headers != ["Pathway"] + CRITERIA:
        raise ValueError("Expected columns Pathway, B1, ..., B18 in A1:S1.")
    names = [str(row[0]).strip() for row in rows[1:]]
    if len(set(names)) != 21 or set(names) != set(PATHWAYS):
        raise ValueError("Expected exactly Base Case and P1 to P20, without duplicates.")
    if any(v is None or isinstance(v, bool) for row in rows[1:] for v in row[1:]):
        raise ValueError("Missing or invalid rating.")
    X = np.asarray([rows[1 + names.index(p)][1:] for p in PATHWAYS], dtype=float)
    if not np.isfinite(X).all() or np.any((X < 1) | (X > 5)) or np.any(X != np.floor(X)):
        raise ValueError("Ratings must be integers between 1 and 5.")
    return X

# Original six scoring functions. Every internal score is higher-is-better.
def s_aras(X, w, v=V_VIKOR, lam=LAMBDA_WASPAS):
    Xe = np.vstack([X.min(0), X])                      # row 0 = optimal alternative
    b = 1.0 / Xe                                       # cost criteria
    S = (b / b.sum(0) * w).sum(1)
    return S[1:] / S[0]


def s_copras(X, w, v=V_VIKOR, lam=LAMBDA_WASPAS):
    s = (X / X.sum(0) * w).sum(1)                      # S-; S+ = 0 (all cost)
    return s.min() * s.sum() / (s * (s.min() / s).sum())


def s_moora(X, w, v=V_VIKOR, lam=LAMBDA_WASPAS):
    return -(X / np.sqrt((X ** 2).sum(0)) * w).sum(1)


def s_topsis(X, w, v=V_VIKOR, lam=LAMBDA_WASPAS):
    V = X / np.sqrt((X ** 2).sum(0)) * w
    d_pos = np.sqrt(((V - V.min(0)) ** 2).sum(1))
    d_neg = np.sqrt(((V - V.max(0)) ** 2).sum(1))
    return d_neg / (d_pos + d_neg + EPS)


def vikor_srq(X, w, v=V_VIKOR):
    lo, hi = X.min(0), X.max(0)
    rng = hi - lo
    d = np.divide(X - lo, rng, out=np.zeros_like(X), where=rng > 0) * w  # zero-range -> 0
    S, R = d.sum(1), d.max(1)

    def scale(a):
        return (a - a.min()) / (a.max() - a.min()) if a.max() > a.min() else np.zeros_like(a)

    return S, R, v * scale(S) + (1 - v) * scale(R)


def s_vikor(X, w, v=V_VIKOR, lam=LAMBDA_WASPAS):
    return -vikor_srq(X, w, v)[2]                      # lower Q = better


def s_waspas(X, w, v=V_VIKOR, lam=LAMBDA_WASPAS):
    r = X.min(0) / X
    return lam * (r * w).sum(1) + (1 - lam) * np.prod(r ** w, axis=1)


SCORERS = {"ARAS": s_aras, "COPRAS": s_copras, "MOORA": s_moora,
           "TOPSIS": s_topsis, "VIKOR": s_vikor, "WASPAS": s_waspas}


def method_ranks(X, w, v=V_VIKOR, lam=LAMBDA_WASPAS, methods=METHODS, tie=TIE_RULE):
    """21 x len(methods) array of ranks (1 = best)."""
    return np.column_stack([
        rankdata(-np.round(SCORERS[m](X, w, v=v, lam=lam), ROUND), method=tie)
        for m in methods])


# Export the original decision matrix and computed scores/ranks.
def write_sheet(wb, name, headers, rows, last_col, number_format=None):
    ws = wb.worksheets.add(name)
    ws.get_range(f"A1:{last_col}{len(rows)+1}").values = [headers] + rows
    ws.get_range(f"A1:{last_col}1").format = {
        "fill": "#23435B", "font": {"bold": True, "color": "#FFFFFF"},
        "row_height": 24, "horizontal_alignment": "center"}
    ws.get_range(f"A1:A{len(rows)+1}").format.column_width = 18
    ws.get_range(f"B1:{last_col}{len(rows)+1}").format.column_width = 15 if number_format else 9
    if number_format:
        ws.get_range(f"B2:{last_col}{len(rows)+1}").set_number_format(number_format)
    ws.freeze_panes.freeze_rows(1)
    return ws


def main():
    wb = Workbook.create()
    results = {}
    for period, sheet in SHEETS.items():
        X = load_matrix(INPUT, sheet)
        w = np.full(X.shape[1], 1.0 / X.shape[1])
        ranks = method_ranks(X, w)
        scores = np.column_stack([SCORERS[m](X, w) for m in METHODS])
        scores[:, METHODS.index("VIKOR")] *= -1  # conventional positive Q
        ws = write_sheet(wb, f"Ranks {period}", ["Pathway"] + METHODS + ["Mean rank"],
                         [[p] + ranks[i].astype(int).tolist() + [float(ranks[i].mean())]
                          for i, p in enumerate(PATHWAYS)], "H")
        ws.get_range("H2:H22").formulas = [[f"=AVERAGE(B{i}:G{i})"] for i in range(2, 23)]
        ws.get_range("H2:H22").set_number_format("0.00")
        ws.get_range("H1:H22").format.column_width = 15
        write_sheet(wb, f"Scores {period}", ["Pathway"] + METHODS,
                    [[p] + scores[i].tolist() for i, p in enumerate(PATHWAYS)], "G", "0.0000000000")
        results[period] = ranks
        print(period, "\n", ranks)
    a, b = (results[p] for p in SHEETS)
    ws = write_sheet(wb, "Table 4 comparison", ["Pathway"] + [f"{m} (2025-2035)" for m in METHODS] + [f"{m} (after-2035)" for m in METHODS],
                     [[p] + a[i].astype(int).tolist() + b[i].astype(int).tolist() for i,p in enumerate(PATHWAYS)], "M")
    ws.get_range("B1:M22").format.column_width = 23
    write_sheet(wb, "Settings", ["Setting", "Value"], [["Criteria", "All cost-type"], ["Weight per criterion", 1/18], ["VIKOR v", V_VIKOR], ["WASPAS lambda", LAMBDA_WASPAS], ["Tie rule", TIE_RULE], ["Score rounding", ROUND], ["Rank 1", "Least barrier-constrained"], ["VIKOR scores", "Q (lower is better)"]], "B").get_range("A1:B9").format.column_width = 32
    SpreadsheetFile.export_xlsx(wb).save(str(OUTPUT))
    print(f"Saved {OUTPUT}")
    return results


if __name__ == "__main__":
    main()
