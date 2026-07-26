"""Pull NCREIF ODCE vintage (YrBuiltorRenov) and market value at legal property share,
grouped by property type.

Metrics returned per PropertyType per quarter:
  - VW_YrBuiltorRenov         : year built/renovated weighted by market value at 100%
                                = SUM(YrBuiltorRenov * MV) / SUM(MV)
  - VW_YrBuiltorRenov_atShare : year built/renovated weighted by market value at legal
                                property share
                                = SUM(YrBuiltorRenov * MV * LegalPropertyShare)
                                  / SUM(MV * LegalPropertyShare)
  - EW_YrBuiltorRenov         : equal-weighted (simple average) year built or renovated
  - MV_at_Share               : SUM(MV * LegalPropertyShare)
  - MV_100pct                 : SUM(MV)
  - Prop_Count                : property count

Universe: DataTypeId=3 (transitioned NPI database), ODCE funds only (FundType='D').
No NPI_Plus=1 membership flag applied.

Credentials: set NCREIF_EMAIL / NCREIF_PASSWORD at the top of this file, or leave
them blank to pick up the env vars of the same name.

Usage:
    export NCREIF_EMAIL=you@firm.com        # optional if set in the file
    export NCREIF_PASSWORD='...'
    python ncreif_odce_vintage.py                       # latest 4 quarters
    python ncreif_odce_vintage.py --start 20241 --end 20261
    python ncreif_odce_vintage.py --buckets             # add vintage-decade breakdown
    python ncreif_odce_vintage.py --out odce_vintage.xlsx
"""

from __future__ import annotations

import argparse
import os
import sys
import time

import pandas as pd
import requests

# ---------------------------------------------------------------------------
# Credentials
#
# Set these directly if you want, but env vars are safer -- this file is in a
# git repo, and anything typed here can be committed and pushed by accident.
# Leave them blank to fall back to NCREIF_EMAIL / NCREIF_PASSWORD.
# ---------------------------------------------------------------------------
NCREIF_EMAIL = ""
NCREIF_PASSWORD = ""

BASE_URL = "https://qt-api.ncreif.org"
DATA_TYPE_ID = 3  # transitioned NPI (the official NPI as of 2026Q1)

# ODCE funds and clean vintage/value records only. No NPI_Plus membership filter --
# this covers all ODCE property-quarters, including those outside the NPI.
BASE_WHERE = (
    "FundType='D' "
    "AND YrBuiltorRenov Is Not Null AND YrBuiltorRenov > 1800 "
    "AND MV > 0"
)

SELECT_MAIN = (
    # Vintage weighted by market value at 100%.
    "SUM(YrBuiltorRenov*MV)/SUM(MV) AS 'VW_YrBuiltorRenov', "
    # Vintage weighted by market value at legal property share.
    "SUM(YrBuiltorRenov*MV*LegalPropertyShare)/SUM(MV*LegalPropertyShare) "
    "AS 'VW_YrBuiltorRenov_atShare', "
    "AVG(YrBuiltorRenov) AS 'EW_YrBuiltorRenov', "
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
        self._last_call = 0.0

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
        return _coerce_numeric(df)


def _coerce_numeric(df: pd.DataFrame) -> pd.DataFrame:
    for col in df.columns:
        if "PropertyType" in col or col.startswith("Vintage"):
            continue
        converted = pd.to_numeric(df[col], errors="coerce")
        if converted.notna().any():
            df[col] = converted
    return df


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
    )
    if df.empty:
        return df

    df = df.sort_values(["YYYYQ", PT_FIELD]).reset_index(drop=True)
    for col in ("VW_YrBuiltorRenov", "VW_YrBuiltorRenov_atShare", "EW_YrBuiltorRenov"):
        df[col] = df[col].round(1)
    df["Share_Pct_of_MV"] = df["MV_at_Share"] / df["MV_100pct"]

    # Each quarter's MV-at-share mix across property types.
    df["Pct_of_Quarter_MV_at_Share"] = df.groupby("YYYYQ")["MV_at_Share"].transform(
        lambda s: s / s.sum()
    )
    return df


def pull_vintage_buckets(
    client: NCREIFClient, start: int | None, end: int | None
) -> pd.DataFrame:
    """Market value at legal share distributed across vintage-decade buckets.

    NCREIF has no vintage-bucket field, so bucket by [Year]-[YrBuiltorRenov] age bands
    with one query per band and stack the results.
    """
    bands = [
        ("0-4 yrs", "[Year]-[YrBuiltorRenov] <= 4"),
        ("5-9 yrs", "[Year]-[YrBuiltorRenov] BETWEEN 5 AND 9"),
        ("10-19 yrs", "[Year]-[YrBuiltorRenov] BETWEEN 10 AND 19"),
        ("20-29 yrs", "[Year]-[YrBuiltorRenov] BETWEEN 20 AND 29"),
        ("30+ yrs", "[Year]-[YrBuiltorRenov] >= 30"),
    ]
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


def latest_quarter_pivot(df: pd.DataFrame) -> pd.DataFrame:
    q = df["YYYYQ"].max()
    snap = df[df["YYYYQ"] == q].set_index(PT_FIELD)
    return snap[
        [
            "VW_YrBuiltorRenov",
            "VW_YrBuiltorRenov_atShare",
            "EW_YrBuiltorRenov",
            "MV_at_Share",
            "MV_100pct",
            "Prop_Count",
            "Pct_of_Quarter_MV_at_Share",
        ]
    ]


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--start", type=int, default=None, help="start YYYYQ, e.g. 20241")
    ap.add_argument("--end", type=int, default=None, help="end YYYYQ, e.g. 20261")
    ap.add_argument("--buckets", action="store_true", help="add vintage-band breakdown")
    ap.add_argument("--out", default=None, help="write results to .xlsx or .csv")
    args = ap.parse_args()

    email = NCREIF_EMAIL or os.environ.get("NCREIF_EMAIL")
    password = NCREIF_PASSWORD or os.environ.get("NCREIF_PASSWORD")
    if not email or not password:
        print(
            "No credentials. Set NCREIF_EMAIL / NCREIF_PASSWORD at the top of this "
            "file, or export them as env vars.",
            file=sys.stderr,
        )
        return 1

    client = NCREIFClient(email, password)
    main_df = pull_by_property_type(client, args.start, args.end)
    if main_df.empty:
        print("No rows returned -- check the period filter or masking (>=3 props, >=3 managers).")
        return 2

    pd.set_option("display.width", 160)
    pd.set_option("display.float_format", lambda v: f"{v:,.4f}")
    print(f"\nODCE vintage + market value by property type -- latest quarter "
          f"{int(main_df['YYYYQ'].max())}\n")
    print(latest_quarter_pivot(main_df).to_string())

    buckets_df = pd.DataFrame()
    if args.buckets:
        buckets_df = pull_vintage_buckets(client, args.start, args.end)
        if not buckets_df.empty:
            q = buckets_df["YYYYQ"].max()
            print(f"\nMV at legal share by vintage band -- {int(q)}\n")
            print(
                buckets_df[buckets_df["YYYYQ"] == q]
                .pivot(
                    index=PT_FIELD,
                    columns="Vintage_Band",
                    values="Pct_of_Type_MV_at_Share",
                )
                .to_string()
            )

    if args.out:
        if args.out.endswith(".xlsx"):
            with pd.ExcelWriter(args.out, engine="openpyxl") as xw:
                main_df.to_excel(xw, sheet_name="By_PropertyType", index=False)
                if not buckets_df.empty:
                    buckets_df.to_excel(xw, sheet_name="Vintage_Bands", index=False)
        else:
            main_df.to_csv(args.out, index=False)
        print(f"\nWrote {args.out}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
