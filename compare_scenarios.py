#!/usr/bin/env python3
"""
Haifa model scenario comparison report generator.

Compares two model scenario folders under raw/ (default: 2000 = before data
update, 2001 = after data update) and writes Excel comparison workbooks:

  comparison/Comparison_Report.xlsx  - summary, file inventory, all report
                                       tables (Trip/Car/Transit/Boarding/
                                       Cordon), matrix totals, input files,
                                       super-zone analysis (grouping column
                                       set by SZ_COLUMN below): motorization,
                                       transit boardings, passengers per car
                                       and mode totals.
  comparison/Matrix_Details.xlsx     - per-zone origin/destination totals for
                                       every demand matrix and every period,
                                       plus super-zone O/D matrices for the
                                       main modes (Driver/Transit/Passenger).

Data formats handled:
  *.rep       '@'-delimited report tables, one block per period (AM/OP/PM/NE).
  mat/*.mat   raw little-endian float32 matrices stacked in the order listed
              in the matching *.in file; dimension = sqrt(size/4/n_matrices).
              808 zones = the centroid list in CentroidGroup.in; the 829-zone
              FinalDemand files add 21 park&ride station zones (their position
              in the zone order is detected by aligning row sums against the
              808-zone Demand5 file of the same period).
  CentroidGroup.in   Emme group definition listing all 808 centroid numbers.
  TAZ_North3.csv     TAZ attributes incl. superZone used for aggregation.
  TransitCapacity.in Emme "matrices by zones" listing, per-zone values.
  PeriodZone.csv     super-zone level demand (present only in scenario 2000).

Usage:  python3 compare_scenarios.py [--base 2000] [--updated 2001]
"""

import argparse
import hashlib
import math
import re
from pathlib import Path

import numpy as np
import pandas as pd
import xlsxwriter

# ---------------------------------------------------------------------------
# configuration
# ---------------------------------------------------------------------------

ROOT = Path(__file__).resolve().parent
RAW = ROOT / "raw"
OUT = ROOT / "comparison"

PERIOD_ORDER = ["AM", "OP", "PM", "NE"]

# TAZ_North3.csv column that assigns each TAZ to a super zone, used by the
# "SZ ..." analysis sheets (motorization, transit boardings, pass per car,
# mode totals). Change here if your TAZ file names the column differently.
SZ_COLUMN = "SZ_new"
MAT_PERIOD_DIRS = {"am": "AM", "op": "OP", "pm": "PM", "eve": "NE"}

REPORT_FILES = ["Trip.rep", "Car.rep", "Transit.rep", "Boarding.rep", "Cordon.rep"]

# fixed, human-readable names for the two header rows of Car.rep
CAR_COLUMNS = [
    "Car*Km V/C<=0.75",
    "Car*Km V/C 0.75-1.0",
    "Car*Km V/C 1.0-1.25",
    "Car*Km V/C>1.25",
    "Car*Km Total",
    "Car*Hour Total",
    "Avg Speed km/h",
]

BLOCK_HEADER_RE = re.compile(r"^(\d+):([A-Z]{2}):\s*(.*)$")


def is_number(s: str) -> bool:
    s = s.strip()
    if not s or s in ("n/a", "N/A"):
        return False
    try:
        float(s)
        return True
    except ValueError:
        return False


def to_number(s: str):
    s = s.strip()
    if not s or s in ("n/a", "N/A"):
        return np.nan
    try:
        return float(s)
    except ValueError:
        return np.nan


# ---------------------------------------------------------------------------
# report (.rep) parsing
# ---------------------------------------------------------------------------

def split_blocks(path: Path):
    """Split a .rep file into (period, title, [lines]) blocks."""
    blocks = []
    period = title = None
    lines = []
    for raw_line in path.read_text().splitlines():
        m = BLOCK_HEADER_RE.match(raw_line)
        if m:
            if period is not None:
                blocks.append((period, title, lines))
            period, title, lines = m.group(2), m.group(3), []
        elif period is not None and raw_line.strip():
            lines.append(raw_line)
    if period is not None:
        blocks.append((period, title, lines))
    return blocks


def parse_trip(path: Path) -> pd.DataFrame:
    rows = []
    for period, _title, lines in split_blocks(path):
        header = None
        for line in lines:
            f = line.split("@")
            if f[0].strip() == "From":
                header = [c.strip() for c in f[2:]]
                continue
            if header is None or len(f) < 3:
                continue
            rows.append(
                {"Period": period, "From": f[0].strip(), "To": f[1].strip(),
                 **{h: to_number(v) for h, v in zip(header, f[2:])}}
            )
    return pd.DataFrame(rows)


def parse_car(path: Path) -> pd.DataFrame:
    rows = []
    for period, _title, lines in split_blocks(path):
        for line in lines:
            f = line.split("@")
            if not is_number(f[1] if len(f) > 1 else ""):
                continue  # the two header lines
            rows.append(
                {"Period": period, "Area": f[0].strip(),
                 **{h: to_number(v) for h, v in zip(CAR_COLUMNS, f[1:])}}
            )
    return pd.DataFrame(rows)


def qualify_rows(lines):
    """Yield (qualified_label, fields) using indentation to build a hierarchy
    so repeated labels (e.g. 'Auto' under two parents) stay unique."""
    stack = []  # list of (indent, label)
    for line in lines:
        indent = len(line) - len(line.lstrip(" "))
        label = line.split("@")[0].strip()
        while stack and stack[-1][0] >= indent:
            stack.pop()
        qualified = " / ".join([s[1] for s in stack] + [label]) if stack else label
        yield qualified, line.split("@"), indent
        stack.append((indent, label))


def parse_transit(path: Path) -> pd.DataFrame:
    rows = []
    for period, _title, lines in split_blocks(path):
        header = None
        body = []
        for line in lines:
            f = line.split("@")
            if f[0].strip() == "" and not is_number(f[1]):
                header = [c.strip() for c in f[1:]]
                continue
            body.append(line)
        for qualified, f, _ind in qualify_rows(body):
            vals = f[1:]
            if len(vals) == 1:  # trailing scalar rows (Demand w/o Inner, ...)
                rows.append({"Period": period, "Row": qualified,
                             "Total value": to_number(vals[0])})
            else:
                rows.append({"Period": period, "Row": qualified,
                             **{h: to_number(v) for h, v in zip(header, vals)}})
    df = pd.DataFrame(rows)
    cols = ["Period", "Row"] + [c for c in df.columns if c not in ("Period", "Row")]
    return df[cols]


def parse_boarding(path: Path) -> pd.DataFrame:
    rows = []
    for period, _title, lines in split_blocks(path):
        modes = None
        user_class = None
        for line in lines:
            f = line.split("@")
            if f[0].strip() == "" and not is_number(f[1]):
                modes = [c.strip() for c in f[1:]]
                continue
            if len(f) == 2:  # e.g. 'Train_users@16811.1054'
                user_class = f[0].strip()
                rows.append({"Period": period, "User class": user_class,
                             "Measure": "Users total",
                             "Class total": to_number(f[1])})
            else:
                rows.append({"Period": period, "User class": user_class,
                             "Measure": f[0].strip(),
                             **{m: to_number(v) for m, v in zip(modes, f[1:])}})
    df = pd.DataFrame(rows)
    cols = ["Period", "User class", "Measure", "Class total"] + [
        c for c in df.columns if c not in ("Period", "User class", "Measure", "Class total")]
    return df[cols]


def parse_cordon(path: Path) -> pd.DataFrame:
    """Cordon.rep is a stack of small heterogeneous tables -> long format."""
    rows = []
    for period, _title, lines in split_blocks(path):
        section, col_names = "", []
        for qualified, f, _ind in qualify_rows(lines):
            non_numeric_tail = [x for x in f[1:] if x.strip() and not is_number(x)]
            if non_numeric_tail:  # section header line
                section = f[0].strip() or "/".join(x.strip() for x in f[1:] if not is_number(x))
                col_names = [x.strip() if not is_number(x) else "" for x in f[1:]]
                # a header may still carry numeric fields (e.g. Pass*Km ratio)
                for i, v in enumerate(f[1:]):
                    if is_number(v):
                        rows.append({"Period": period, "Section": section,
                                     "Row": "(header value)", "Column": f"V{i + 1}",
                                     "Value": to_number(v)})
                continue
            for i, v in enumerate(f[1:]):
                if not v.strip():
                    continue
                col = col_names[i] if i < len(col_names) and col_names[i] else f"V{i + 1}"
                rows.append({"Period": period, "Section": section,
                             "Row": qualified, "Column": col,
                             "Value": to_number(v)})
    return pd.DataFrame(rows)


REPORT_PARSERS = {
    "Trip.rep": (parse_trip, ["Period", "From", "To"]),
    "Car.rep": (parse_car, ["Period", "Area"]),
    "Transit.rep": (parse_transit, ["Period", "Row"]),
    "Boarding.rep": (parse_boarding, ["Period", "User class", "Measure"]),
    "Cordon.rep": (parse_cordon, ["Period", "Section", "Row", "Column"]),
}


def report_title(scen_dir: Path) -> str:
    """Model run title from the first report block header."""
    for rep in REPORT_FILES:
        p = scen_dir / "report" / rep
        if p.exists():
            m = BLOCK_HEADER_RE.match(p.read_text().splitlines()[0])
            if m:
                return m.group(3)
    return "n/a"


# ---------------------------------------------------------------------------
# input files
# ---------------------------------------------------------------------------

def read_centroids(scen_dir: Path):
    nums = []
    for line in (scen_dir / "CentroidGroup.in").read_text().splitlines():
        m = re.match(r"^a\s+gh1:\s+(.*)$", line)
        if m:
            nums += [int(x) for x in m.group(1).split()]
    return nums


def parse_transit_capacity(path: Path) -> pd.DataFrame:
    """Parse the per-zone table of the Emme 'matrices by zones' listing."""
    scen_of_md = {}
    for m in re.finditer(r"(md\d+)\s+:\s+\S+\s+Transit Capacity, scenario (\d+)",
                         path.read_text()):
        scen_of_md[m.group(1)] = m.group(2)
    # column order in the data table (first column repeats the last md)
    header_mds = None
    rows = []
    for line in path.read_text().splitlines():
        if header_mds is None and re.match(r"^\s*origin\s+destin", line):
            header_mds = re.findall(r"md\d+", line)
            continue
        if header_mds and re.match(r"^\s*\d+(\s+[-.\d]+)+\s*$", line):
            parts = line.split()
            zone, vals = int(parts[0]), [to_number(v) for v in parts[1:]]
            row = {"Zone": zone}
            for md, v in zip(header_mds, vals):
                scen = scen_of_md.get(md, md)
                # 12000->AM, 22000->OP, 32000->PM, 42000->NE
                per = {"1": "AM", "2": "OP", "3": "PM", "4": "NE"}.get(scen[:1], md)
                row[f"Capacity {per}"] = v
            rows.append(row)
    df = pd.DataFrame(rows)
    return df.loc[:, ~df.columns.duplicated(keep="last")]


def parse_taz(path: Path) -> pd.DataFrame:
    return pd.read_csv(path)


# ---------------------------------------------------------------------------
# matrices
# ---------------------------------------------------------------------------

def read_in_labels(path: Path):
    labels = []
    for line in path.read_text().splitlines():
        parts = line.split()
        if len(parts) >= 2 and parts[0].isdigit():
            labels.append(parts[1])
    return labels


def read_mat(path: Path, n_matrices: int):
    size = path.stat().st_size
    n2 = size // (4 * n_matrices)
    n = math.isqrt(n2)
    if n * n != n2 or size != 4 * n_matrices * n * n:
        raise ValueError(f"{path}: size {size} does not fit {n_matrices} square float32 matrices")
    return np.fromfile(path, dtype="<f4").reshape(n_matrices, n, n)


def find_extra_positions(small_rowsums, big_rowsums, rel_tol=1e-3):
    """Positions in the big zone system that are absent from the small one,
    found by aligning row sums in order."""
    i = j = 0
    extra = []
    while i < len(small_rowsums) and j < len(big_rowsums):
        if abs(small_rowsums[i] - big_rowsums[j]) <= rel_tol * max(1.0, abs(small_rowsums[i])):
            i += 1
            j += 1
        else:
            extra.append(j)
            j += 1
    extra += list(range(j, len(big_rowsums)))
    if i < len(small_rowsums):
        return None  # alignment failed
    return extra


class MatrixStore:
    """All matrices of one scenario, with zone labels."""

    def __init__(self, scen_dir: Path, centroids):
        self.scen_dir = scen_dir
        self.centroids = centroids
        # {(period, file_stem, matrix_label): 2-D array}
        self.mats = {}
        # {(period, file_stem): [zone labels]}
        self.zones = {}
        self._load()

    def _load(self):
        mat_root = self.scen_dir / "mat"
        for sub in sorted(mat_root.iterdir()):
            period = MAT_PERIOD_DIRS.get(sub.name, sub.name)
            for in_file in sorted(sub.glob("*.in")):
                mat_file = in_file.with_suffix(".mat")
                if not mat_file.exists():
                    continue
                labels = read_in_labels(in_file)
                data = read_mat(mat_file, len(labels))
                stem = re.sub(r"_\d+$", "", in_file.stem)  # Demand5_1 -> Demand5
                for lab, m in zip(labels, data):
                    self.mats[(period, stem, lab)] = m
                self.zones[(period, stem)] = data.shape[1]
        self._label_zones()

    def _label_zones(self):
        """Replace zone counts with zone-number label lists."""
        n_small = len(self.centroids)
        # detect where the extra (P&R station) zones sit, using any period
        # that has both a small and a big matrix of the Driver demand
        extra_positions = None
        for (period, stem, lab), m in self.mats.items():
            if m.shape[0] != n_small or lab != "Driver":
                continue
            big = self.mats.get((period, "FinalDemand", "Driver"))
            if big is None:
                continue
            extra_positions = find_extra_positions(m.sum(axis=1), big.sum(axis=1))
            if extra_positions:
                break
        for key, n in list(self.zones.items()):
            if n == n_small:
                self.zones[key] = [str(z) for z in self.centroids]
            elif extra_positions and n == n_small + len(extra_positions):
                labels = [str(z) for z in self.centroids]
                for k, pos in enumerate(extra_positions):
                    labels.insert(pos, f"PR_STN_{k + 1:02d}")
                self.zones[key] = labels
            else:
                self.zones[key] = [f"idx_{i + 1}" for i in range(n)]


# ---------------------------------------------------------------------------
# Excel writing helpers
# ---------------------------------------------------------------------------

class ReportWriter:
    def __init__(self, path: Path, base_name: str, upd_name: str):
        self.wb = xlsxwriter.Workbook(str(path), {"nan_inf_to_errors": True})
        self.base_name = base_name
        self.upd_name = upd_name
        f = self.wb.add_format
        self.f_title = f({"bold": True, "font_size": 14, "font_color": "#1F4E79"})
        self.f_note = f({"italic": True, "font_color": "#666666"})
        self.f_head = f({"bold": True, "bg_color": "#1F4E79", "font_color": "white",
                         "border": 1, "align": "center", "valign": "vcenter", "text_wrap": True})
        self.f_subhead = f({"bold": True, "bg_color": "#D9E2F3", "border": 1, "align": "center"})
        self.f_key = f({"border": 1})
        self.f_num = f({"border": 1, "num_format": "#,##0.0"})
        self.f_num_base = f({"border": 1, "num_format": "#,##0.0", "bg_color": "#F2F2F2"})
        self.f_diff = f({"border": 1, "num_format": "[Blue]#,##0.0;[Red]-#,##0.0;#,##0.0"})
        # 'precise' variants: up to 3 decimals, for ratio tables
        self.f_num3 = f({"border": 1, "num_format": "#,##0.0##"})
        self.f_num3_base = f({"border": 1, "num_format": "#,##0.0##", "bg_color": "#F2F2F2"})
        self.f_diff3 = f({"border": 1, "num_format": "[Blue]#,##0.0##;[Red]-#,##0.0##;#,##0.0##"})
        self.f_pct = f({"border": 1, "num_format": "[Blue]0.0%;[Red]-0.0%;0.0%"})
        self.f_sect = f({"bold": True, "font_size": 12, "font_color": "#1F4E79"})

    def close(self):
        self.wb.close()

    def sheet(self, name):
        return self.wb.add_worksheet(name[:31])

    def write_comparison(self, ws, df_base, df_upd, key_cols, title=None,
                         start_row=0, pct=True, val_cols=None, key_widths=None,
                         precise=False):
        """Side-by-side comparison table: for every value column write
        base / updated / diff (/ %diff) sub-columns."""
        if val_cols is None:
            val_cols = [c for c in df_base.columns if c not in key_cols]
        b = df_base.copy()
        b["_ord"] = range(len(b))
        merged = b.merge(df_upd, on=key_cols, how="outer",
                         suffixes=("|base", "|upd"), sort=False)
        merged = merged.sort_values("_ord", na_position="last")

        sub = [self.base_name, self.upd_name, "Diff"] + (["%Diff"] if pct else [])
        nsub = len(sub)
        r0 = start_row
        if title:
            ws.write(r0, 0, title, self.f_sect)
            r0 += 1
        # header rows
        for i, k in enumerate(key_cols):
            ws.merge_range(r0, i, r0 + 1, i, k, self.f_head)
        c = len(key_cols)
        for v in val_cols:
            if nsub > 1:
                ws.merge_range(r0, c, r0, c + nsub - 1, v, self.f_head)
            else:
                ws.write(r0, c, v, self.f_head)
            for j, s in enumerate(sub):
                ws.write(r0 + 1, c + j, s, self.f_subhead)
            c += nsub
        f_base = self.f_num3_base if precise else self.f_num_base
        f_num = self.f_num3 if precise else self.f_num
        f_diff = self.f_diff3 if precise else self.f_diff

        # data
        def write_val(rr, cc, v, fmt):
            if v is None or (isinstance(v, float) and not math.isfinite(v)):
                ws.write_blank(rr, cc, None, fmt)
            else:
                ws.write(rr, cc, v, fmt)

        r = r0 + 2
        for _, row in merged.iterrows():
            for i, k in enumerate(key_cols):
                ws.write(r, i, row[k], self.f_key)
            c = len(key_cols)
            for v in val_cols:
                vb = row.get(f"{v}|base", row.get(v, np.nan))
                vu = row.get(f"{v}|upd", np.nan)
                vb = np.nan if pd.isna(vb) else float(vb)
                vu = np.nan if pd.isna(vu) else float(vu)
                diff = vu - vb if not (pd.isna(vb) or pd.isna(vu)) else np.nan
                write_val(r, c, vb, f_base)
                write_val(r, c + 1, vu, f_num)
                write_val(r, c + 2, diff, f_diff)
                if pct:
                    p = diff / vb if diff == diff and vb != 0.0 else np.nan
                    write_val(r, c + 3, p, self.f_pct)
                c += nsub
            r += 1
        # layout
        widths = key_widths or [max(12, min(34, int(merged[k].astype(str).str.len().max() or 10) + 2))
                                for k in key_cols]
        for i, w in enumerate(widths):
            ws.set_column(i, i, w)
        ws.set_column(len(key_cols), len(key_cols) + len(val_cols) * nsub - 1, 12)
        ws.freeze_panes(r0 + 2, len(key_cols))
        return r  # next free row


# ---------------------------------------------------------------------------
# building blocks of the report
# ---------------------------------------------------------------------------

def file_inventory(base_dir: Path, upd_dir: Path) -> pd.DataFrame:
    def digest(p: Path):
        return hashlib.md5(p.read_bytes()).hexdigest()

    rels = sorted(
        {p.relative_to(d).as_posix() for d in (base_dir, upd_dir)
         for p in d.rglob("*") if p.is_file()}
    )
    rows = []
    for rel in rels:
        pb, pu = base_dir / rel, upd_dir / rel
        if pb.exists() and pu.exists():
            status = "identical" if digest(pb) == digest(pu) else "DIFFERS"
        elif pb.exists():
            status = f"only in {base_dir.name}"
        else:
            status = f"only in {upd_dir.name}"
        rows.append({"File": rel,
                     f"Size {base_dir.name}": pb.stat().st_size if pb.exists() else None,
                     f"Size {upd_dir.name}": pu.stat().st_size if pu.exists() else None,
                     "Status": status})
    return pd.DataFrame(rows)


def matrix_totals(store_b: MatrixStore, store_u: MatrixStore) -> pd.DataFrame:
    rows = []
    keys = sorted(set(store_b.mats) | set(store_u.mats),
                  key=lambda k: (PERIOD_ORDER.index(k[0]) if k[0] in PERIOD_ORDER else 9,
                                 k[1], k[2]))
    for key in keys:
        period, stem, lab = key
        mb = store_b.mats.get(key)
        mu = store_u.mats.get(key)
        tb = float(mb.sum()) if mb is not None else np.nan
        tu = float(mu.sum()) if mu is not None else np.nan
        if mb is not None and mu is not None and mb.shape == mu.shape:
            d = mu - mb
            max_cell = float(np.abs(d).max())
            n_changed = int((np.abs(d) > 0.01).sum())
        else:
            max_cell, n_changed = np.nan, np.nan
        rows.append({"Period": period, "File": stem, "Matrix": lab,
                     "Total": (tb, tu),
                     "Max abs cell diff": max_cell,
                     "Cells changed (>0.01)": n_changed})
    # split the tuple into base/upd frames handled by write_comparison
    base = pd.DataFrame([{"Period": r["Period"], "File": r["File"], "Matrix": r["Matrix"],
                          "Total": r["Total"][0]} for r in rows])
    upd = pd.DataFrame([{"Period": r["Period"], "File": r["File"], "Matrix": r["Matrix"],
                         "Total": r["Total"][1]} for r in rows])
    extra = pd.DataFrame([{"Period": r["Period"], "File": r["File"], "Matrix": r["Matrix"],
                           "Max abs cell diff": r["Max abs cell diff"],
                           "Cells changed (>0.01)": r["Cells changed (>0.01)"]} for r in rows])
    return base, upd, extra


def zone_totals_frame(store: MatrixStore, period: str, stems):
    """DataFrame keyed by Zone with '<Matrix> Orig'/'<Matrix> Dest' columns."""
    out = None
    for stem in stems:
        labs = [k[2] for k in store.mats if k[0] == period and k[1] == stem]
        if not labs:
            continue
        zones = store.zones[(period, stem)]
        df = pd.DataFrame({"Zone": zones})
        for lab in labs:
            m = store.mats[(period, stem, lab)]
            df[f"{lab} Orig"] = m.sum(axis=1)
            df[f"{lab} Dest"] = m.sum(axis=0)
        out = df if out is None else out.merge(df, on="Zone", how="outer")
    return out


TOUR_PURPOSES = {"W": "Work", "E": "Education", "O": "Other"}


def tours_summary(store: MatrixStore, taz: pd.DataFrame) -> pd.DataFrame:
    """Tour totals, trips-per-tour and tours-per-capita headline measures.

    Tours: mat/tours/Tour_5 = all daily tours by exit period (period 0-4);
    AutoObjective_3 / TransitObjective_3 split the same tours by mode and
    purpose (their sum equals the Tour_5 total). Auto_5/Transit_5 count each
    tour twice (exit & return leg) and are therefore not used for totals.
    Trips: mat/<period>/Demand5 = trips of the 4 modelled periods only
    (AM/OP/PM/NE) - the model does not output trip matrices for period 0,
    so 'trips per tour' relates modelled-period trips to all daily tours.
    """
    def tot(period, stem, labels=None):
        return sum(float(m.sum()) for (p, s, lab), m in store.mats.items()
                   if p == period and s == stem and (labels is None or lab in labels))

    trips_by_period = {p: tot(p, "Demand5") for p in PERIOD_ORDER}
    trips_total = sum(trips_by_period.values())
    auto_trips = sum(tot(p, "Demand5", {"Driver", "Passenger"}) for p in PERIOD_ORDER)
    transit_trips = sum(tot(p, "Demand5", {"Transit", "P&R", "K&R"}) for p in PERIOD_ORDER)

    tours_total = tot("tours", "Tour")
    auto_tours = tot("tours", "AutoObjective")
    transit_tours = tot("tours", "TransitObjective")
    purpose_tours = {name: tot("tours", "AutoObjective", {f"Auto_{k}"})
                     + tot("tours", "TransitObjective", {f"Trns_{k}"})
                     for k, name in TOUR_PURPOSES.items()}
    population = float(taz["population"].sum())

    def ratio(a, b):
        return a / b if b else np.nan

    rows = [("Trips " + p + " (Demand5, all modes)", trips_by_period[p]) for p in PERIOD_ORDER]
    rows += [
        ("Trips total, modelled periods", trips_total),
        ("  Auto trips (Driver + Passenger)", auto_trips),
        ("  Transit trips (Transit + P&R + K&R)", transit_trips),
        ("Tours total, all day (Tour_5)", tours_total),
        ("  Auto tours (AutoObjective)", auto_tours),
        ("  Transit tours (TransitObjective)", transit_tours),
        ("  Work tours", purpose_tours["Work"]),
        ("  Education tours", purpose_tours["Education"]),
        ("  Other tours", purpose_tours["Other"]),
        ("Trips per tour (modelled-period trips / all-day tours)",
         ratio(trips_total, tours_total)),
        ("  Auto trips per auto tour", ratio(auto_trips, auto_tours)),
        ("  Transit trips per transit tour", ratio(transit_trips, transit_tours)),
        ("Population (TAZ_North3)", population),
        ("Tours per capita", ratio(tours_total, population)),
        ("  Auto tours per capita", ratio(auto_tours, population)),
        ("  Transit tours per capita", ratio(transit_tours, population)),
        ("  Work tours per capita", ratio(purpose_tours["Work"], population)),
        ("  Education tours per capita", ratio(purpose_tours["Education"], population)),
        ("  Other tours per capita", ratio(purpose_tours["Other"], population)),
    ]
    return pd.DataFrame(rows, columns=["Measure", "Value"])


def tours_per_capita_by_superzone(store: MatrixStore, taz: pd.DataFrame,
                                  sz_of_zone: dict) -> pd.DataFrame:
    """Per-superZone tours (by origin = home zone), population and ratios."""
    zones = store.zones[("tours", "Tour")]
    groups = [str(sz_of_zone.get(z, "EXT")) for z in zones]

    def orig_sums(stem, labels=None):
        out = np.zeros(len(zones))
        for (p, s, lab), m in store.mats.items():
            if p == "tours" and s == stem and (labels is None or lab in labels):
                out += m.sum(axis=1)
        return out

    df = pd.DataFrame({"SuperZone": groups,
                       "Tours": orig_sums("Tour"),
                       "Auto tours": orig_sums("AutoObjective"),
                       "Transit tours": orig_sums("TransitObjective"),
                       "Work tours": orig_sums("AutoObjective", {"Auto_W"})
                       + orig_sums("TransitObjective", {"Trns_W"}),
                       "Education tours": orig_sums("AutoObjective", {"Auto_E"})
                       + orig_sums("TransitObjective", {"Trns_E"}),
                       "Other tours": orig_sums("AutoObjective", {"Auto_O"})
                       + orig_sums("TransitObjective", {"Trns_O"})})
    agg = df.groupby("SuperZone", as_index=False).sum()
    pop = (taz.assign(SuperZone=taz["superZone"].astype(int).astype(str))
           .groupby("SuperZone", as_index=False)["population"].sum()
           .rename(columns={"population": "Population"}))
    agg = agg.merge(pop, on="SuperZone", how="left")
    total = agg.drop(columns="SuperZone").sum()
    total["SuperZone"] = "TOTAL"
    agg = pd.concat([agg, total.to_frame().T], ignore_index=True)
    agg = agg.sort_values(
        "SuperZone",
        key=lambda s: s.map(lambda x: int(x) if str(x).isdigit()
                            else (9998 if x == "EXT" else 9999)),
        ignore_index=True)
    for col in ["Tours", "Auto tours", "Transit tours",
                "Work tours", "Education tours", "Other tours"]:
        with np.errstate(divide="ignore", invalid="ignore"):
            agg[f"{col} per capita"] = np.where(
                agg["Population"].astype(float) > 0,
                agg[col].astype(float) / agg["Population"].astype(float), np.nan)
    cols = (["SuperZone", "Population", "Tours", "Tours per capita"]
            + [c for base in ["Auto tours", "Transit tours", "Work tours",
                              "Education tours", "Other tours"]
               for c in (base, f"{base} per capita")])
    return agg[cols]


def sz_sort_key(s: pd.Series):
    """Numeric super zones first, then EXT, then TOTAL."""
    return s.map(lambda x: int(x) if str(x).isdigit() else (9998 if x == "EXT" else 9999))


def sz_new_map(taz: pd.DataFrame) -> dict:
    return {str(int(t)): str(int(s)) for t, s in zip(taz["TAZ"], taz[SZ_COLUMN])}


def add_total_row(df: pd.DataFrame, key_col: str = "SuperZone") -> pd.DataFrame:
    total = df.drop(columns=key_col).sum(numeric_only=True)
    total[key_col] = "TOTAL"
    return pd.concat([df, total.to_frame().T], ignore_index=True)


def sz_motorization(taz: pd.DataFrame) -> pd.DataFrame:
    agg = (taz.assign(SuperZone=taz[SZ_COLUMN].astype(int).astype(str))
           .groupby("SuperZone", as_index=False)[["population", "car"]].sum()
           .rename(columns={"population": "Population", "car": "Cars"}))
    agg = add_total_row(agg)
    with np.errstate(divide="ignore", invalid="ignore"):
        agg["Motorization (cars/1000 pop)"] = np.where(
            agg["Population"].astype(float) > 0,
            agg["Cars"].astype(float) / agg["Population"].astype(float) * 1000.0,
            np.nan)
    return agg.sort_values("SuperZone", key=sz_sort_key, ignore_index=True)


def demand5_origin_by_sz(store: MatrixStore, labels, sz_map: dict) -> pd.DataFrame:
    """Super zone (SZ_COLUMN) x period origin sums of the Demand5 matrices in `labels`."""
    out = {}
    for period in PERIOD_ORDER:
        zones = store.zones.get((period, "Demand5"))
        if zones is None:
            continue
        tot = np.zeros(len(zones))
        for lab in labels:
            m = store.mats.get((period, "Demand5", lab))
            if m is not None:
                tot += m.sum(axis=1)
        out[period] = pd.Series(tot).groupby(
            pd.Series([sz_map.get(z, "EXT") for z in zones])).sum()
    df = pd.DataFrame(out).fillna(0.0)
    df.index.name = "SuperZone"
    return df.reset_index()


def sz_transit_boardings(store: MatrixStore, sz_map: dict) -> pd.DataFrame:
    df = demand5_origin_by_sz(store, ["Transit", "P&R", "K&R"], sz_map)
    periods = [p for p in PERIOD_ORDER if p in df.columns]
    df["Daily total"] = df[periods].sum(axis=1)
    return add_total_row(df).sort_values("SuperZone", key=sz_sort_key, ignore_index=True)


def sz_pass_per_car(store: MatrixStore, sz_map: dict) -> pd.DataFrame:
    drv = add_total_row(demand5_origin_by_sz(store, ["Driver"], sz_map)).set_index("SuperZone")
    pas = add_total_row(demand5_origin_by_sz(store, ["Passenger"], sz_map)).set_index("SuperZone")
    pas = pas.reindex(drv.index).fillna(0.0)
    out = pd.DataFrame(index=drv.index)
    out["Drivers (day)"] = drv.sum(axis=1).astype(float)
    out["Car passengers (day)"] = pas.sum(axis=1).astype(float)
    with np.errstate(divide="ignore", invalid="ignore"):
        for p in [p for p in PERIOD_ORDER if p in drv.columns]:
            d, x = drv[p].astype(float), pas[p].astype(float)
            out[f"Pass per car {p}"] = np.where(d > 0, (d + x) / d, np.nan)
        d, x = out["Drivers (day)"], out["Car passengers (day)"]
        out["Pass per car (day)"] = np.where(d > 0, (d + x) / d, np.nan)
    return out.reset_index().sort_values("SuperZone", key=sz_sort_key, ignore_index=True)


def sz_mode_totals(store: MatrixStore, sz_map: dict) -> pd.DataFrame:
    def day_total(labels, name):
        df = demand5_origin_by_sz(store, labels, sz_map).set_index("SuperZone")
        return df.sum(axis=1).rename(name)

    out = pd.concat([day_total(["Driver"], "Drivers"),
                     day_total(["Passenger"], "Car passengers"),
                     day_total(["Transit", "P&R", "K&R"], "Transit passengers")],
                    axis=1).fillna(0.0).reset_index()
    return add_total_row(out).sort_values("SuperZone", key=sz_sort_key, ignore_index=True)


def superzone_matrix(store: MatrixStore, period: str, stem: str, label: str,
                     sz_of_zone: dict):
    m = store.mats.get((period, stem, label))
    if m is None:
        return None
    zones = store.zones[(period, stem)]
    groups = [sz_of_zone.get(z, "STA" if z.startswith("PR_STN") else "EXT") for z in zones]
    order = sorted(set(groups), key=lambda g: (isinstance(g, str), g))
    idx = {g: i for i, g in enumerate(order)}
    gi = np.array([idx[g] for g in groups])
    agg = np.zeros((len(order), len(order)))
    np.add.at(agg, (gi[:, None], gi[None, :]), m)
    return pd.DataFrame(agg, index=order, columns=order)


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--base", default="2000", help="baseline scenario folder under raw/")
    ap.add_argument("--updated", default="2001", help="updated scenario folder under raw/")
    args = ap.parse_args()

    base_dir, upd_dir = RAW / args.base, RAW / args.updated
    base_lbl, upd_lbl = f"{args.base} (before)", f"{args.updated} (after)"
    OUT.mkdir(exist_ok=True)

    print(f"Comparing raw/{args.base} (baseline) vs raw/{args.updated} (updated)")

    # ---- parse everything -------------------------------------------------
    reports = {}
    for rep, (parser, keys) in REPORT_PARSERS.items():
        pb, pu = base_dir / "report" / rep, upd_dir / "report" / rep
        if pb.exists() and pu.exists():
            reports[rep] = (parser(pb), parser(pu), keys)
            print(f"  parsed {rep}: {len(reports[rep][0])} rows per scenario")

    centroids = read_centroids(base_dir)
    print(f"  zone system: {len(centroids)} centroids")
    store_b = MatrixStore(base_dir, centroids)
    store_u = MatrixStore(upd_dir, centroids)
    print(f"  matrices: {len(store_b.mats)} in baseline, {len(store_u.mats)} in updated")

    taz_b, taz_u = parse_taz(base_dir / "TAZ_North3.csv"), parse_taz(upd_dir / "TAZ_North3.csv")
    cap_b = parse_transit_capacity(base_dir / "TransitCapacity.in")
    cap_u = parse_transit_capacity(upd_dir / "TransitCapacity.in")
    inv = file_inventory(base_dir, upd_dir)

    sz_of_zone = {str(int(r.TAZ)): int(r.superZone) for r in taz_b.itertuples()}

    # ---- validation: matrix totals vs Trip.rep -----------------------------
    trip_b = reports["Trip.rep"][0]
    tot_row = trip_b[(trip_b["From"] == "Metropolin Total") & (trip_b["To"] == "Metropolin Total")
                     & (trip_b["Period"] == "AM")]
    if not tot_row.empty:
        mat_total = store_b.mats.get(("AM", "FinalDemand", "Driver"))
        if mat_total is not None:
            rep_val = float(tot_row["DRIVER"].iloc[0])
            m_val = float(mat_total.sum())
            ok = abs(rep_val - m_val) <= 0.01 * rep_val
            print(f"  validation Driver AM: report={rep_val:,.1f} matrix={m_val:,.1f} "
                  f"{'OK' if ok else 'MISMATCH!'}")

    # ======================================================================
    # workbook 1: Comparison_Report.xlsx
    # ======================================================================
    w = ReportWriter(OUT / "Comparison_Report.xlsx", base_lbl, upd_lbl)

    # ---- Summary -----------------------------------------------------------
    ws = w.sheet("Summary")
    ws.write(0, 0, "Haifa model scenario comparison", w.f_title)
    ws.write(1, 0, f"Baseline:  raw/{args.base}  -  {report_title(base_dir)}", w.f_note)
    ws.write(2, 0, f"Updated:   raw/{args.updated}  -  {report_title(upd_dir)}", w.f_note)
    ws.write(3, 0, "Diff = updated - baseline;  %Diff = Diff / baseline "
                   "(blue = increase, red = decrease)", w.f_note)

    def trip_totals(df):
        t = df[(df["From"] == "Metropolin Total") & (df["To"] == "Metropolin Total")].copy()
        return t.drop(columns=["From", "To"])

    r = w.write_comparison(ws, trip_totals(reports["Trip.rep"][0]),
                           trip_totals(reports["Trip.rep"][1]), ["Period"],
                           title="Total trips (Metropolin Total -> Metropolin Total, from Trip.rep)",
                           start_row=5)

    def car_totals(df):
        t = df[df["Area"] == "Metropolin, Total"].copy()
        return t[["Period", "Car*Km Total", "Car*Hour Total", "Avg Speed km/h"]]

    r = w.write_comparison(ws, car_totals(reports["Car.rep"][0]),
                           car_totals(reports["Car.rep"][1]), ["Period"],
                           title="Network performance (Metropolin Total, from Car.rep)",
                           start_row=r + 2)

    def transit_totals(df):
        t = df[df["Row"] == "All Public Transport Modes"].copy()
        return t[["Period", "Initial Boarding", "Boarding", "Passenger*Km", "Passenger*Hour"]]

    w.write_comparison(ws, transit_totals(reports["Transit.rep"][0]),
                       transit_totals(reports["Transit.rep"][1]), ["Period"],
                       title="Public transport (All PT modes, from Transit.rep)",
                       start_row=r + 2)

    # ---- Inventory ----------------------------------------------------------
    ws = w.sheet("File Inventory")
    ws.write(0, 0, "File inventory (md5 comparison of raw/ contents)", w.f_title)
    identical = [f for f, s in zip(inv["File"], inv["Status"])
                 if s == "identical" and "/" not in f]
    one_sided = [f"{f} ({s})" for f, s in zip(inv["File"], inv["Status"])
                 if s.startswith("only in")]
    notes = []
    if identical:
        notes.append("Identical top-level input files: " + ", ".join(identical) + ".")
    if one_sided:
        notes.append("Files present in only one scenario: " + ", ".join(one_sided) + ".")
    notes.append("Note: TransitCapacity.in header contains run timestamps, so it can show as "
                 "DIFFERS even when all data values match (see its sheet).")
    for i, n in enumerate(notes):
        ws.write(1 + i, 0, n, w.f_note)
    hr = 5
    for i, c in enumerate(inv.columns):
        ws.write(hr, i, c, w.f_head)
    for j, (_, row) in enumerate(inv.iterrows()):
        for i, c in enumerate(inv.columns):
            v = row[c]
            ws.write(hr + 1 + j, i, "" if pd.isna(v) else v,
                     w.f_num if isinstance(v, (int, float)) and not isinstance(v, bool) else w.f_key)
    ws.set_column(0, 0, 42)
    ws.set_column(1, 3, 16)
    ws.freeze_panes(hr + 1, 0)

    # ---- report sheets -------------------------------------------------------
    sheet_names = {"Trip.rep": "Trip", "Car.rep": "Car", "Transit.rep": "Transit",
                   "Boarding.rep": "Boarding", "Cordon.rep": "Cordon"}
    for rep, (dfb, dfu, keys) in reports.items():
        ws = w.sheet(sheet_names[rep])
        ws.write(0, 0, f"{rep} comparison", w.f_title)
        w.write_comparison(ws, dfb, dfu, keys, start_row=2)

    # ---- matrix totals --------------------------------------------------------
    tb, tu, extra = matrix_totals(store_b, store_u)
    ws = w.sheet("Matrix Totals")
    ws.write(0, 0, "Demand matrix totals (from mat/*.mat binary matrices)", w.f_title)
    r = w.write_comparison(ws, tb, tu, ["Period", "File", "Matrix"], start_row=2)
    # append the cell-level stats to the right of the table
    c0 = 3 + 4  # keys + 4 sub-columns of 'Total'
    ws.write(3, c0, "Max abs cell diff", w.f_head)
    ws.write(3, c0 + 1, "Cells changed (>0.01)", w.f_head)
    for j, (_, row) in enumerate(extra.iterrows()):
        for k, col in enumerate(["Max abs cell diff", "Cells changed (>0.01)"]):
            v = row[col]
            if pd.isna(v):
                ws.write_blank(4 + j, c0 + k, None, w.f_num)
            else:
                ws.write(4 + j, c0 + k, v, w.f_num)
    ws.set_column(c0, c0 + 1, 18)

    # ---- tours analysis --------------------------------------------------------
    ws = w.sheet("Tours Analysis")
    ws.write(0, 0, "Tours vs trips - trips per tour and tours per capita", w.f_title)
    tours_notes = [
        "mat/tours matrices are TOURS (whole day): Tour_5 = all tours by exit period "
        "(0-4); AutoObjective_3/TransitObjective_3 split the same tours by mode and "
        "purpose (Work/Education/Other) - their sum equals the Tour_5 total.",
        "Auto_5/Transit_5 count every tour twice (exit leg + return leg) and are "
        "therefore not used for the totals below.",
        "Trip matrices (mat/<period>/Demand5) exist only for the 4 modelled periods "
        "(AM/OP/PM/NE), so 'trips per tour' relates modelled-period trips to all-day "
        "tours - a full-day trip rate would be roughly 2 trips per tour.",
    ]
    for i, n in enumerate(tours_notes):
        ws.write(1 + i, 0, n, w.f_note)
    w.write_comparison(ws, tours_summary(store_b, taz_b), tours_summary(store_u, taz_u),
                       ["Measure"], start_row=5, precise=True,
                       key_widths=[52])

    ws = w.sheet("Tours per Capita by SZ")
    ws.write(0, 0, "Tours per capita by superZone (tour origin = home zone; "
                   "population from TAZ_North3.csv)", w.f_title)
    ws.write(1, 0, "EXT = external zones 51-88 (no population, so no per-capita value).",
             w.f_note)
    w.write_comparison(ws,
                       tours_per_capita_by_superzone(store_b, taz_b, sz_of_zone),
                       tours_per_capita_by_superzone(store_u, taz_u, sz_of_zone),
                       ["SuperZone"], start_row=3, precise=True)

    # ---- super-zone (SZ_COLUMN) analysis ------------------------------------------
    sznew_b, sznew_u = sz_new_map(taz_b), sz_new_map(taz_u)

    ws = w.sheet("SZ Motorization")
    ws.write(0, 0, "Population, cars and motorization by super zone "
                   f"({SZ_COLUMN} from TAZ_North3.csv)", w.f_title)
    ws.write(1, 0, "Motorization = cars / population * 1000.", w.f_note)
    w.write_comparison(ws, sz_motorization(taz_b), sz_motorization(taz_u),
                       ["SuperZone"], start_row=3, precise=True)

    ws = w.sheet("SZ Transit Boardings")
    ws.write(0, 0, f"Transit boardings by super zone ({SZ_COLUMN}) and period", w.f_title)
    for i, n in enumerate([
        "Transit demand (Transit + P&R + K&R, mat/<period>/Demand5) summed by origin "
        f"zone and aggregated to {SZ_COLUMN} super zones - i.e. boardings without transfers.",
        "Boardings by transit mode are only reported network-wide (Boarding.rep - see "
        "the Boarding sheet); no zone-level mode split exists in the model outputs, "
        "so only total transit boardings are shown per super zone.",
        "EXT = external station zones (not listed in TAZ_North3.csv).",
    ]):
        ws.write(1 + i, 0, n, w.f_note)
    w.write_comparison(ws, sz_transit_boardings(store_b, sznew_b),
                       sz_transit_boardings(store_u, sznew_u),
                       ["SuperZone"], start_row=5)

    ws = w.sheet("SZ Pass per Car")
    ws.write(0, 0, f"Passengers per car by super zone ({SZ_COLUMN})", w.f_title)
    ws.write(1, 0, "Passengers per car = (drivers + car passengers) / drivers, from the "
                   "Demand5 Driver and Passenger matrices by origin super zone; "
                   "(day) = AM+OP+PM+NE.", w.f_note)
    w.write_comparison(ws, sz_pass_per_car(store_b, sznew_b),
                       sz_pass_per_car(store_u, sznew_u),
                       ["SuperZone"], start_row=3, precise=True)

    ws = w.sheet("SZ Mode Totals")
    ws.write(0, 0, "Drivers, car passengers and transit passengers by super zone "
                   f"({SZ_COLUMN})", w.f_title)
    ws.write(1, 0, "Daily totals (AM+OP+PM+NE) by origin super zone from Demand5; "
                   "transit = Transit + P&R + K&R.", w.f_note)
    w.write_comparison(ws, sz_mode_totals(store_b, sznew_b),
                       sz_mode_totals(store_u, sznew_u),
                       ["SuperZone"], start_row=3)

    # ---- TAZ ---------------------------------------------------------------------
    num_cols = [c for c in taz_b.columns if c != "TAZ" and pd.api.types.is_numeric_dtype(taz_b[c])]
    summary_b = pd.DataFrame({"Attribute": num_cols,
                              "Sum": [taz_b[c].sum() for c in num_cols]})
    summary_u = pd.DataFrame({"Attribute": num_cols,
                              "Sum": [taz_u[c].sum() for c in num_cols]})
    ws = w.sheet("TAZ Attributes")
    ws.write(0, 0, "TAZ_North3.csv - attribute sums over all TAZs", w.f_title)
    m = taz_b.merge(taz_u, on="TAZ", suffixes=("_b", "_u"))
    changed = m[[any(m[f"{c}_b"].iloc[i] != m[f"{c}_u"].iloc[i] for c in num_cols)
                 for i in range(len(m))]] if len(m) else m
    ws.write(1, 0, f"Zone-level rows with any change: {len(changed)} of {len(m)}", w.f_note)
    w.write_comparison(ws, summary_b, summary_u, ["Attribute"], start_row=3)

    # ---- Transit capacity -----------------------------------------------------------
    ws = w.sheet("Transit Capacity")
    ws.write(0, 0, "TransitCapacity.in - per-zone transit capacity by period", w.f_title)
    w.write_comparison(ws, cap_b, cap_u, ["Zone"], start_row=2)

    # ---- PeriodZone -----------------------------------------------------------------
    pz_b_path, pz_u_path = base_dir / "PeriodZone.csv", upd_dir / "PeriodZone.csv"
    pz_keys = ["perodFrom", "periodTo", "zoneFrom", "zoneTo"]

    def pz_summary(pz):
        return (pz.groupby(["perodFrom", "periodTo"], as_index=False)[["Auto", "Transit"]]
                .sum().rename(columns={"perodFrom": "Period from", "periodTo": "Period to"}))

    if pz_b_path.exists() and pz_u_path.exists():
        pz_b, pz_u = pd.read_csv(pz_b_path), pd.read_csv(pz_u_path)
        ws = w.sheet("PeriodZone Summary")
        ws.write(0, 0, "PeriodZone.csv - super-zone demand, totals by period pair", w.f_title)
        w.write_comparison(ws, pz_summary(pz_b), pz_summary(pz_u),
                           ["Period from", "Period to"], start_row=2)
        ws = w.sheet("PeriodZone Detail")
        ws.write(0, 0, "PeriodZone.csv - full zone-pair comparison "
                       f"({len(pz_b)} rows in {args.base}, {len(pz_u)} rows in {args.updated}; "
                       "blank = row missing in that scenario)", w.f_title)
        w.write_comparison(ws, pz_b, pz_u, pz_keys, start_row=2)
        print(f"  wrote PeriodZone comparison ({len(pz_b)} vs {len(pz_u)} rows)")
    elif pz_b_path.exists() or pz_u_path.exists():
        only_dir, pz_path = ((args.base, pz_b_path) if pz_b_path.exists()
                             else (args.updated, pz_u_path))
        agg = pz_summary(pd.read_csv(pz_path))
        ws = w.sheet(f"PeriodZone ({only_dir} only)")
        ws.write(0, 0, f"PeriodZone.csv exists only in raw/{only_dir} - no comparison possible.",
                 w.f_title)
        ws.write(1, 0, "Totals by period pair (super-zone level demand):", w.f_note)
        for i, c in enumerate(agg.columns):
            ws.write(3, i, c, w.f_head)
        for j, (_, row) in enumerate(agg.iterrows()):
            for i, c in enumerate(agg.columns):
                ws.write(4 + j, i, row[c], w.f_num)
        ws.set_column(0, 3, 14)

    w.close()
    print(f"  wrote {OUT / 'Comparison_Report.xlsx'}")

    # ======================================================================
    # workbook 2: Matrix_Details.xlsx
    # ======================================================================
    w2 = ReportWriter(OUT / "Matrix_Details.xlsx", base_lbl, upd_lbl)

    ws = w2.sheet("ReadMe")
    ws.write(0, 0, "Matrix details", w2.f_title)
    for i, t in enumerate([
        "Zone totals sheets: origin (row) and destination (column) sums per zone for every",
        "matrix in mat/<period>/FinalDemand and mat/tours, compared between scenarios.",
        "SZ sheets: matrices aggregated to superZone (TAZ_North3.csv); EXT = external",
        "stations zones 51-88, STA = park&ride station zones of the 829-zone system.",
        "Diff = updated - baseline.",
    ]):
        ws.write(1 + i, 0, t, w2.f_note)

    for period in PERIOD_ORDER:
        zb = zone_totals_frame(store_b, period, ["FinalDemand"])
        zu = zone_totals_frame(store_u, period, ["FinalDemand"])
        if zb is None or zu is None:
            continue
        ws = w2.sheet(f"{period} Zone Totals")
        ws.write(0, 0, f"{period} - per-zone origin/destination totals (FinalDemand)", w2.f_title)
        w2.write_comparison(ws, zb, zu, ["Zone"], start_row=2, pct=False)
        print(f"  wrote zone totals {period}")

    zb = zone_totals_frame(store_b, "tours", ["Tour", "Auto", "Transit",
                                              "AutoObjective", "TransitObjective"])
    zu = zone_totals_frame(store_u, "tours", ["Tour", "Auto", "Transit",
                                              "AutoObjective", "TransitObjective"])
    if zb is not None and zu is not None:
        ws = w2.sheet("Tours Zone Totals")
        ws.write(0, 0, "Tours - per-zone origin/destination totals", w2.f_title)
        w2.write_comparison(ws, zb, zu, ["Zone"], start_row=2, pct=False)
        print("  wrote zone totals tours")

    for period in PERIOD_ORDER:
        for lab in ["Driver", "Transit", "Passenger"]:
            mb = superzone_matrix(store_b, period, "FinalDemand", lab, sz_of_zone)
            mu = superzone_matrix(store_u, period, "FinalDemand", lab, sz_of_zone)
            if mb is None or mu is None:
                continue
            ws = w2.sheet(f"{period} SZ {lab}")
            r = 0
            for name, mat in [(base_lbl, mb), (upd_lbl, mu), ("Diff (updated - baseline)", mu - mb)]:
                ws.write(r, 0, f"{period} {lab} - {name} (superZone O/D)", w2.f_sect)
                ws.write(r + 1, 0, "O \\ D", w2.f_head)
                for i, ccol in enumerate(mat.columns):
                    ws.write(r + 1, 1 + i, str(ccol), w2.f_head)
                for j, (idx, row) in enumerate(mat.iterrows()):
                    ws.write(r + 2 + j, 0, str(idx), w2.f_head)
                    fmt = w2.f_diff if name.startswith("Diff") else w2.f_num
                    for i, v in enumerate(row):
                        ws.write(r + 2 + j, 1 + i, float(v), fmt)
                r += len(mat) + 4
            ws.set_column(0, len(mb.columns), 10)
        print(f"  wrote superzone O/D {period}")

    w2.close()
    print(f"  wrote {OUT / 'Matrix_Details.xlsx'}")
    print("Done.")


if __name__ == "__main__":
    main()
