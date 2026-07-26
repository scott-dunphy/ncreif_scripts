"""Pull NCREIF ODCE vintage (YrBuiltorLastRen) and market value at legal property
share, grouped by property type.

Vintage field: YrBuiltorLastRen = year built or last renovated, verified against the
live API (DataTypeId=3). YrBuilt (original construction year only) also exists --
switch the VINTAGE constant to use it. For ODCE 2025Q1 the renovation-adjusted
average is ~3 years newer (2002 vs 1999).

Metrics returned per PropertyType per quarter (VINTAGE = YrBuiltorLastRen):
  - VW_<VINTAGE>          : vintage weighted by market value at 100%
                            = SUM(VINTAGE * MV) / SUM(MV)
  - VW_<VINTAGE>_atShare  : vintage weighted by market value at legal property share
                            = SUM(VINTAGE * MV * LegalPropertyShare)
                              / SUM(MV * LegalPropertyShare)
  - EW_<VINTAGE>          : equal-weighted (simple average) vintage
  - MV_at_Share           : SUM(MV * LegalPropertyShare)
  - MV_100pct             : SUM(MV)
  - Prop_Count            : property count

Universe: DataTypeId=3 (transitioned NPI database), ODCE funds only (FundType='D').
No NPI_Plus=1 membership flag applied.

Credentials: copy .env.example to .env and fill in NCREIF_EMAIL / NCREIF_PASSWORD.
The .env file is gitignored. Real environment variables and the constants at the
top of this file also work -- see the Credentials block below for precedence.

Run with no arguments and it pulls every quarter of DEFAULT_YEAR, including the
vintage-band breakdown, and writes an .xlsx into the working directory.

Usage:
    export NCREIF_EMAIL=you@firm.com        # optional if set in the file
    export NCREIF_PASSWORD='...'
    python ncreif_odce_vintage.py                   # all quarters of DEFAULT_YEAR -> .xlsx
    python ncreif_odce_vintage.py --year 2024       # all quarters of 2024
    python ncreif_odce_vintage.py --start 20241 --end 20261   # explicit range
    python ncreif_odce_vintage.py --no-buckets      # skip the vintage-band queries
    python ncreif_odce_vintage.py --out custom.xlsx # override the output filename
"""

from __future__ import annotations

import argparse
import datetime as dt
import os
import sys
import time

import pandas as pd
import requests

# ---------------------------------------------------------------------------
# Credentials
#
# Preferred: put them in a .env file next to this script (it is gitignored, so
# it never reaches GitHub). Copy .env.example to .env and fill it in:
#
#     NCREIF_EMAIL=you@firm.com
#     NCREIF_PASSWORD=your_password
#
# These two constants are a last-resort override. Anything typed here is
# tracked by git and can be pushed by accident -- leave them blank.
#
# Resolution order: constants below -> real environment variables -> .env file.
# ---------------------------------------------------------------------------
NCREIF_EMAIL = ""
NCREIF_PASSWORD = ""

# Where to look for the .env file: alongside this script, then the working dir.
ENV_FILE = ".env"

# Year pulled when no --year/--start/--end is given. None = current calendar year.
DEFAULT_YEAR: int | None = None

# How many years to walk back looking for data before giving up. NCREIF freezes a
# quarter well after it ends, so early in a calendar year the current year is often
# still empty.
MAX_YEAR_FALLBACK = 2

BASE_URL = "https://qt-api.ncreif.org"
DATA_TYPE_ID = 3  # transitioned NPI (the official NPI as of 2026Q1)

# ---------------------------------------------------------------------------
# TLS verification
#
# False skips certificate validation, which is what you want behind a corporate
# proxy that re-signs traffic (Zscaler, Netskope, Blue Coat) and otherwise
# throws SSLCertVerificationError.
#
# Note this turns off MITM protection on the login call that carries the
# password. The verifying alternative, if IT can give you the proxy's root CA:
#     export REQUESTS_CA_BUNDLE=/path/to/corporate-root-ca.pem
# then set this back to True.
# ---------------------------------------------------------------------------
VERIFY_SSL = False

if not VERIFY_SSL:
    # urllib3 warns on every unverified request; one notice at startup is enough.
    import urllib3

    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Vintage field. YrBuiltorLastRen = year built or last renovated (verified live);
# YrBuilt = original construction year only. Both exist in DataTypeId=3.
VINTAGE = "YrBuiltorLastRen"

# ODCE funds and clean vintage/value records only. No NPI_Plus membership filter --
# this covers all ODCE property-quarters, including those outside the NPI.
#
# Every field in WHERE is bracketed. The API silently DROPS the entire WHERE clause
# when it meets a field name it does not recognize (returning the full 1978+ history
# unfiltered), and bare YYYYQ breaks the parser outright -- brackets avoid both.
BASE_WHERE = (
    "[FundType]='D' "
    f"AND [{VINTAGE}] Is Not Null AND [{VINTAGE}] > 1800 "
    "AND [MV] > 0"
)

SELECT_MAIN = (
    # Vintage weighted by market value at 100%.
    f"SUM({VINTAGE}*MV)/SUM(MV) AS 'VW_{VINTAGE}', "
    # Vintage weighted by market value at legal property share.
    f"SUM({VINTAGE}*MV*LegalPropertyShare)/SUM(MV*LegalPropertyShare) "
    f"AS 'VW_{VINTAGE}_atShare', "
    f"AVG({VINTAGE}) AS 'EW_{VINTAGE}', "
    "SUM(MV*LegalPropertyShare) AS 'MV_at_Share', "
    "SUM(MV) AS 'MV_100pct', "
    "COUNT(MV) AS 'Prop_Count'"
)

# NCREIF's own loaded templates use the Exp_-prefixed type field; it carries the same
# full-name values ('Residential', 'Industrial', 'Office', ...).
PT_FIELD = "Exp_PropertyType"

GROUPBY_MAIN = f"[Year],[YYYYQ],[{PT_FIELD}]"


class NCREIFClient:
    """Thin client for the NCREIF query-tool API. Handles JWT refresh on 401."""

    def __init__(self, email: str, password: str, pause: float = 8.0):
        self.email = email
        self.password = password
        self.pause = pause  # be polite: ~8 queries/min
        self.token: str | None = None
        self.session = requests.Session()
        # Applies to every request made through this session.
        self.session.verify = VERIFY_SSL
        self._last_call = 0.0
        # Every query's exact-as-returned rows, kept for the Raw_Data output tab:
        # list of (label, DataFrame straight off the wire, no derived columns).
        self.raw_log: list[tuple[str, pd.DataFrame]] = []

    def login(self) -> None:
        resp = self.session.post(
            f"{BASE_URL}/Login/Login",
            json={"Email": self.email, "Password": self.password},
            headers={"Content-Type": "application/json"},
            timeout=60,
        )
        resp.raise_for_status()
        # API returns {"message": "<JWT>"} -- the token is the message field.
        self.token = resp.json()["message"]

    def _throttle(self) -> None:
        elapsed = time.monotonic() - self._last_call
        if self._last_call and elapsed < self.pause:
            time.sleep(self.pause - elapsed)
        self._last_call = time.monotonic()

    def query(
        self,
        select: str,
        where: str,
        groupby: str = "Year, YYYYQ",
        query_data: int = 0,
        label: str = "",
    ) -> pd.DataFrame:
        if self.token is None:
            self.login()

        payload = {
            "p_DataTypeId": DATA_TYPE_ID,
            "p_SelectQuery": select,
            "p_WhereClause": where,
            "p_GroupbyClause": groupby,
            "p_QueryData": query_data,
            "kpi": "",
        }

        for attempt in range(2):
            self._throttle()
            resp = self.session.post(
                f"{BASE_URL}/QT_V2/ExecuteQuery",
                json=payload,
                headers={
                    "Authorization": f"Bearer {self.token}",
                    "Content-Type": "application/json",
                },
                timeout=180,
            )
            if resp.status_code == 401 and attempt == 0:
                self.login()  # token expires after 20 minutes
                continue
            if resp.status_code == 500:
                raise RuntimeError(
                    f"NCREIF 500 -- check SELECT/WHERE/GROUP BY syntax.\n{resp.text[:500]}"
                )
            resp.raise_for_status()
            break

        rows = resp.json().get("rows", [])
        records = [{f["name"]: f["value"] for f in row["fields"]} for row in rows]
        df = pd.DataFrame(records)
        if not df.empty:
            self.raw_log.append((label or where, df.copy()))
        return _coerce_numeric(df)


def _coerce_numeric(df: pd.DataFrame) -> pd.DataFrame:
    for col in df.columns:
        if "PropertyType" in col or col.startswith("Vintage"):
            continue
        converted = pd.to_numeric(df[col], errors="coerce")
        if converted.notna().any():
            df[col] = converted
    return df


def load_dotenv(filename: str = ENV_FILE) -> dict[str, str]:
    """Read a simple KEY=VALUE .env file. No dependency on python-dotenv.

    Looks next to this script first, then in the working directory. Blank lines
    and #-comments are skipped; surrounding quotes on the value are stripped.
    Returns an empty dict when no file is found.
    """
    here = os.path.dirname(os.path.abspath(__file__))
    for path in (os.path.join(here, filename), os.path.abspath(filename)):
        if not os.path.isfile(path):
            continue
        values: dict[str, str] = {}
        with open(path, encoding="utf-8") as fh:
            for line in fh:
                line = line.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                key, _, val = line.partition("=")
                val = val.strip()
                if len(val) >= 2 and val[0] == val[-1] and val[0] in "\"'":
                    val = val[1:-1]
                values[key.strip()] = val
        return values
    return {}


def resolve_credentials() -> tuple[str | None, str | None, str]:
    """Constants win, then real env vars, then the .env file. Also reports source."""
    dotenv = load_dotenv()
    email = NCREIF_EMAIL or os.environ.get("NCREIF_EMAIL") or dotenv.get("NCREIF_EMAIL")
    password = (
        NCREIF_PASSWORD
        or os.environ.get("NCREIF_PASSWORD")
        or dotenv.get("NCREIF_PASSWORD")
    )
    if NCREIF_EMAIL or NCREIF_PASSWORD:
        source = "constants in this file"
    elif os.environ.get("NCREIF_EMAIL"):
        source = "environment variables"
    elif dotenv:
        source = ENV_FILE
    else:
        source = "nowhere"
    return email, password, source


def year_range(year: int) -> tuple[int, int]:
    """All four quarters of a calendar year as YYYYQ bounds: 2025 -> (20251, 20254)."""
    return year * 10 + 1, year * 10 + 4


def report_quarter_coverage(df: pd.DataFrame, year: int) -> None:
    """Say which quarters of the year actually came back, and which did not."""
    got = sorted(int(q) for q in df["YYYYQ"].unique())
    missing = [year * 10 + q for q in (1, 2, 3, 4) if year * 10 + q not in got]
    print(f"Quarters returned: {', '.join(str(q) for q in got)}")
    if missing:
        print(
            "Not available (likely not yet frozen/released): "
            + ", ".join(str(q) for q in missing)
        )


def period_clause(start: int | None, end: int | None) -> str:
    parts = []
    if start is not None:
        parts.append(f"[YYYYQ] >= {start}")
    if end is not None:
        parts.append(f"[YYYYQ] <= {end}")
    return (" AND " + " AND ".join(parts)) if parts else ""


def pull_by_property_type(
    client: NCREIFClient, start: int | None, end: int | None
) -> pd.DataFrame:
    """Vintage + market value by property type, per quarter."""
    df = client.query(
        select=SELECT_MAIN,
        where=BASE_WHERE + period_clause(start, end),
        groupby=GROUPBY_MAIN,
        label="vintage_by_property_type",
    )
    if df.empty:
        return df

    df = df.sort_values(["YYYYQ", PT_FIELD]).reset_index(drop=True)
    for col in (f"VW_{VINTAGE}", f"VW_{VINTAGE}_atShare", f"EW_{VINTAGE}"):
        df[col] = df[col].round(1)
    df["Share_Pct_of_MV"] = df["MV_at_Share"] / df["MV_100pct"]

    # Each quarter's MV-at-share mix across property types.
    df["Pct_of_Quarter_MV_at_Share"] = df.groupby("YYYYQ")["MV_at_Share"].transform(
        lambda s: s / s.sum()
    )
    return df


# Rolling age bands: age measured against the observation [Year], so back-history
# classifies correctly. Mutually exclusive and collectively exhaustive; the <= 9
# band also absorbs any negative age (renovation year ahead of the observation
# year), which would otherwise fall through all three.
AGE_BANDS = [
    ("1_Last 10 yrs", f"[Year]-[{VINTAGE}] <= 9"),
    ("2_Prior 10 yrs", f"[Year]-[{VINTAGE}] BETWEEN 10 AND 19"),
    ("3_Over 20 yrs", f"[Year]-[{VINTAGE}] >= 20"),
]

# Fixed calendar cohorts on the vintage year itself. Same-year cutoffs regardless of
# observation quarter. BASE_WHERE's > 1800 floor keeps junk years out of Pre-2005.
YEAR_COHORTS = [
    ("1_2015 or newer", f"[{VINTAGE}] >= 2015"),
    ("2_2005-2014", f"[{VINTAGE}] BETWEEN 2005 AND 2014"),
    ("3_Pre-2005", f"[{VINTAGE}] < 2005"),
]


def pull_banded(
    client: NCREIFClient,
    start: int | None,
    end: int | None,
    bands: list[tuple[str, str]],
    tag: str,
) -> pd.DataFrame:
    """Market value at legal share split across a set of vintage bands.

    NCREIF has no bucket field, so run one query per band and stack the results.
    Band definitions must be mutually exclusive and collectively exhaustive.
    """
    frames = []
    for label, clause in bands:
        part = client.query(
            select=(
                "SUM(MV*LegalPropertyShare) AS 'MV_at_Share', "
                "SUM(MV) AS 'MV_100pct', "
                "COUNT(MV) AS 'Prop_Count'"
            ),
            where=f"{BASE_WHERE}{period_clause(start, end)} AND {clause}",
            groupby=GROUPBY_MAIN,
            label=f"{tag}_{label}",
        )
        if part.empty:
            continue
        part.insert(0, "Vintage_Band", label)
        frames.append(part)

    if not frames:
        return pd.DataFrame()

    out = pd.concat(frames, ignore_index=True)
    out["Pct_of_Type_MV_at_Share"] = out.groupby(["YYYYQ", PT_FIELD])[
        "MV_at_Share"
    ].transform(lambda s: s / s.sum())
    return out.sort_values(["YYYYQ", PT_FIELD, "Vintage_Band"]).reset_index(drop=True)


def check_band_coverage(main_df: pd.DataFrame, buckets_df: pd.DataFrame) -> pd.DataFrame:
    """Verify the bands partition the universe -- no overlap, nothing dropped.

    Sums each band's property count per quarter/type and compares against the
    unbanded total. A positive gap means properties fell through the bands; a
    negative gap means a property was counted in more than one band.
    """
    keys = ["YYYYQ", PT_FIELD]
    banded = buckets_df.groupby(keys)["Prop_Count"].sum().rename("Banded_Props")
    total = main_df.set_index(keys)["Prop_Count"].rename("Total_Props")
    cmp = pd.concat([total, banded], axis=1).fillna(0)
    cmp["Gap"] = cmp["Total_Props"] - cmp["Banded_Props"]
    return cmp[cmp["Gap"] != 0].reset_index()


def latest_quarter_pivot(df: pd.DataFrame) -> pd.DataFrame:
    q = df["YYYYQ"].max()
    snap = df[df["YYYYQ"] == q].set_index(PT_FIELD)
    return snap[
        [
            f"VW_{VINTAGE}",
            f"VW_{VINTAGE}_atShare",
            f"EW_{VINTAGE}",
            "MV_at_Share",
            "MV_100pct",
            "Prop_Count",
            "Pct_of_Quarter_MV_at_Share",
        ]
    ]


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument(
        "--year",
        type=int,
        default=None,
        help="calendar year to pull, all four quarters (default: DEFAULT_YEAR)",
    )
    ap.add_argument("--start", type=int, default=None, help="start YYYYQ, e.g. 20241")
    ap.add_argument("--end", type=int, default=None, help="end YYYYQ, e.g. 20261")
    ap.add_argument(
        "--no-buckets",
        dest="buckets",
        action="store_false",
        help="skip the vintage-band breakdown (on by default)",
    )
    ap.add_argument(
        "--out",
        default=None,
        help="output path; defaults to ncreif_odce_vintage_<year>.xlsx in the "
             "working directory. Pass 'none' to skip writing a file.",
    )
    args = ap.parse_args()

    email, password, source = resolve_credentials()
    if not email or not password:
        print(
            "No credentials found. Copy .env.example to .env and fill in "
            "NCREIF_EMAIL and NCREIF_PASSWORD (or export them as env vars).",
            file=sys.stderr,
        )
        return 1
    print(f"Credentials from: {source}")
    if not VERIFY_SSL:
        print("TLS certificate verification is OFF (VERIFY_SSL = False).")

    client = NCREIFClient(email, password)

    # An explicit --start/--end wins; otherwise pull a full calendar year, walking
    # back a year at a time if the requested one has not been released yet.
    explicit_range = args.start is not None or args.end is not None
    if explicit_range:
        year = None
        start, end = args.start, args.end
        main_df = pull_by_property_type(client, start, end)
    else:
        year = args.year or DEFAULT_YEAR or dt.date.today().year
        for attempt in range(MAX_YEAR_FALLBACK + 1):
            start, end = year_range(year)
            print(f"Pulling all quarters of {year} ({start}-{end}) ...")
            main_df = pull_by_property_type(client, start, end)
            if not main_df.empty:
                break
            if attempt < MAX_YEAR_FALLBACK:
                year -= 1
                print(f"  no data; falling back to {year}")

    if main_df.empty:
        print(
            "No rows returned -- check the period filter, credentials, or masking "
            "(>=3 properties from >=3 managers).",
            file=sys.stderr,
        )
        return 2

    if year is not None:
        report_quarter_coverage(main_df, year)

    pd.set_option("display.width", 160)
    pd.set_option("display.float_format", lambda v: f"{v:,.4f}")
    print(f"\nODCE vintage + market value by property type -- latest quarter "
          f"{int(main_df['YYYYQ'].max())}\n")
    print(latest_quarter_pivot(main_df).to_string())

    def show_banded(df: pd.DataFrame, title: str) -> None:
        q = df["YYYYQ"].max()
        snap = df[df["YYYYQ"] == q]
        print(f"\nShare of MV at legal share by {title} -- {int(q)}\n")
        print(
            snap.pivot(
                index=PT_FIELD,
                columns="Vintage_Band",
                values="Pct_of_Type_MV_at_Share",
            ).to_string()
        )
        print(f"\nMV at legal share by {title} -- {int(q)}\n")
        print(
            snap.pivot(
                index=PT_FIELD, columns="Vintage_Band", values="MV_at_Share"
            ).to_string()
        )
        gaps = check_band_coverage(main_df, df)
        if not gaps.empty:
            print(
                f"\nWARNING: {title} counts do not reconcile to the unbanded total "
                "(masking may suppress small bands):\n"
            )
            print(gaps.to_string(index=False))

    buckets_df = pd.DataFrame()
    cohorts_df = pd.DataFrame()
    if args.buckets:
        buckets_df = pull_banded(client, start, end, AGE_BANDS, tag="age")
        if not buckets_df.empty:
            show_banded(buckets_df, "age band")

        cohorts_df = pull_banded(client, start, end, YEAR_COHORTS, tag="cohort")
        if not cohorts_df.empty:
            show_banded(cohorts_df, "vintage cohort")

    # Default to an .xlsx in the working directory; 'none' opts out.
    out = args.out
    if out is None:
        tag = str(year) if year is not None else f"{start}_{end}"
        out = f"ncreif_odce_vintage_{tag}.xlsx"
    if out.lower() == "none":
        return 0

    out_path = os.path.abspath(out)
    if out.endswith(".xlsx"):
        # Every API response exactly as returned, stacked, tagged by source query --
        # no rounding, no derived columns.
        raw_df = pd.concat(
            [df.assign(Query=label) for label, df in client.raw_log],
            ignore_index=True,
        ) if client.raw_log else pd.DataFrame()

        with pd.ExcelWriter(out_path, engine="openpyxl") as xw:
            latest_quarter_pivot(main_df).to_excel(xw, sheet_name="Latest_Snapshot")
            main_df.to_excel(xw, sheet_name="By_PropertyType", index=False)
            if not buckets_df.empty:
                buckets_df.to_excel(xw, sheet_name="Age_Bands", index=False)
                buckets_df.pivot_table(
                    index=[PT_FIELD, "YYYYQ"],
                    columns="Vintage_Band",
                    values="Pct_of_Type_MV_at_Share",
                ).to_excel(xw, sheet_name="AgeBand_Pct_by_Quarter")
            if not cohorts_df.empty:
                cohorts_df.to_excel(xw, sheet_name="Vintage_Cohorts", index=False)
                cohorts_df.pivot_table(
                    index=[PT_FIELD, "YYYYQ"],
                    columns="Vintage_Band",
                    values="Pct_of_Type_MV_at_Share",
                ).to_excel(xw, sheet_name="Cohort_Pct_by_Quarter")
            if not raw_df.empty:
                cols = ["Query"] + [c for c in raw_df.columns if c != "Query"]
                raw_df[cols].to_excel(xw, sheet_name="Raw_Data", index=False)
    else:
        main_df.to_csv(out_path, index=False)
    print(f"\nWrote {out_path}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
