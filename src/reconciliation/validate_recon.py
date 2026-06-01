"""K3_Code-agnostic validation: proves the transform logic matches the stale
embedded sheets on every column except the (now-stale) K3_Code source field."""
import warnings; warnings.simplefilter("ignore")
from collections import Counter
import numpy as np, openpyxl, pandas as pd
import recon

WB = recon.DEFAULT_WORKBOOK

def read_table(sheet, ref):
    wb = openpyxl.load_workbook(WB, read_only=True, data_only=True)
    rows = [[c.value for c in r] for r in wb[sheet][ref]]
    wb.close()
    return pd.DataFrame(rows[1:], columns=rows[0]).dropna(how="all")

def n(v):
    if v is None or (isinstance(v, float) and np.isnan(v)) or pd.isna(v): return 0.0
    try: return round(float(v), 2)
    except (ValueError, TypeError): return str(v)

def s(v):
    if v is None or (isinstance(v, float) and np.isnan(v)) or pd.isna(v): return ""
    return str(v).strip()

res = recon.run()

# ---------- schKToK2Part2Diff, ignoring K3_Code ----------
exp = read_table("schKToK2Part2Diff-Flat", "A1:J177")
got = res["schKToK2Part2Diff"]

def sig_rows(df):
    rows = []
    for _, r in df.iterrows():
        rows.append((s(r.get("Source")), s(r.get("SchKLine")), s(r.get("Part 2 Line")),
                     s(r.get("Activity_Num")), s(r.get("Activity_Name")),
                     s(r.get("Activity_Allocation")), s(r.get("Attribute")),
                     n(r.get("Schedule K")), n(r.get("Part 2"))))
    return sorted(rows)

es, gs = sig_rows(exp), sig_rows(got)
print("=== schKToK2Part2Diff (K3-agnostic) ===")
print(f"  expected {len(es)} rows, got {len(gs)} rows")
match = es == gs
print(f"  MULTISET MATCH (all cols except K3_Code): {match}")
if not match:
    ce, cg = Counter(es), Counter(gs)
    only_e = list((ce - cg).elements()); only_g = list((cg - ce).elements())
    print(f"  only in expected: {len(only_e)}, only in got: {len(only_g)}")
    for row in only_e[:8]: print("   E:", row)
    for row in only_g[:8]: print("   G:", row)

# ---------- grossIncome: amount-triple distribution (K3-agnostic) ----------
exp_g = read_table("(K.02) Sch K to K2 Recon", "C12:F60")
got_g = res["grossIncome"]

def amt_multiset(df):
    return sorted([(n(a), n(b), n(c)) for a, b, c in
                   zip(df["Part 2"], df["Part X"], df["Difference"])])

print("\n=== grossIncome (K3-agnostic amount triples) ===")
print(f"  expected {len(exp_g)} rows, got {len(got_g)} rows")
em, gm = amt_multiset(exp_g), amt_multiset(got_g)
print(f"  MULTISET MATCH: {em == gm}")
if em != gm:
    ce, cg = Counter(em), Counter(gm)
    print("  only in expected:", list((ce - cg).elements())[:10])
    print("  only in got:", list((cg - ce).elements())[:10])
