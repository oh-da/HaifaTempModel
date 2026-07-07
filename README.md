# HaifaTempModel

Haifa transport model — scenario data and comparison tooling.

## Data (`raw/`)

Two model scenario folders with identical structure, both representing year 2020:

| Folder | Meaning |
|--------|---------|
| `raw/2000` | Scenario **before** the data update (based on old data) |
| `raw/2001` | Scenario **after** updating the data in the last month |

Each scenario contains:

- `report/*.rep` — `@`-delimited model output reports, one block per time
  period (**AM / OP / PM / NE**): `Trip.rep` (O/D trips by mode), `Car.rep`
  (car network performance), `Transit.rep` (PT performance), `Boarding.rep`
  (boardings by mode/user class), `Cordon.rep` (cordon-line summaries).
- `mat/<am|op|pm|eve|tours>/*.mat` — binary demand matrices: raw
  little-endian float32 square matrices stacked in the order listed in the
  matching `*.in` file (dimension = √(file size ⁄ 4 ⁄ n matrices)).
  808-zone matrices follow the centroid list in `CentroidGroup.in`;
  the 829-zone `FinalDemand` matrices add 21 park&ride station zones.
- `CentroidGroup.in` — Emme group definition with all 808 centroid numbers.
- `TAZ_North3.csv` — TAZ attributes (population, employment, superZone, …).
- `TransitCapacity.in` — Emme "matrices by zones" listing of transit capacity.
- `PeriodZone.csv` — super-zone level demand (present only in `raw/2000`).

## Comparison report

```bash
pip install -r requirements.txt
python3 compare_scenarios.py            # defaults: --base 2000 --updated 2001
```

Writes two workbooks into `comparison/`:

- **`Comparison_Report.xlsx`** — `Summary` (headline KPIs per period),
  `File Inventory` (which files differ), one sheet per report
  (`Trip`, `Car`, `Transit`, `Boarding`, `Cordon`), `Matrix Totals`
  (total of every demand matrix + cell-level change stats),
  `TAZ Attributes`, `Transit Capacity`, `PeriodZone (base only)`.
- **`Matrix_Details.xlsx`** — per-zone origin/destination totals for every
  FinalDemand and tours matrix, plus super-zone O/D matrices
  (Driver / Transit / Passenger per period; baseline, updated and diff).

Every value is shown as *baseline / updated / Diff / %Diff*
(Diff = updated − baseline; blue = increase, red = decrease).
The script validates itself by checking that matrix totals reproduce the
`Trip.rep` report totals.
