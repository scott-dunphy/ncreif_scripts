---
name: ncreif-query-tool
description: >
  Query the NCREIF Property Database via the REST API to retrieve commercial real estate
  performance data including returns, cap rates, NOI growth, occupancy, leveraged returns,
  and more. Use this skill whenever the user wants to pull NCREIF data, build NPI queries,
  analyze property index returns, compare property types or geographies using NCREIF,
  look up CBSA codes, or construct advanced WHERE/SELECT/GROUP BY clauses for the
  NCREIF Query Tool or API. Also trigger when the user mentions "NPI", "NCREIF",
  "Expanded NPI", "Classic NPI", "ODCE", "query tool", "property index", or asks about
  institutional real estate benchmarks, property-level performance data, or NCREIF field names.
  Also covers the 2026Q1 NPI transition (Expanded NPI becoming the official NPI, Classic
  query-tool retirement, 2020-Census CBSAs).
---

# NCREIF Query Tool Skill

Build and execute queries against the NCREIF Property Database REST API. This skill
translates natural-language questions about commercial real estate performance into
properly formatted API calls and returns structured results.

## ⚠️ 2026Q1 NPI Transition (read first)

As of the **First Quarter 2026 release**, the Expanded NPI **became the official NPI**.
This changes the default mental model for every query:

- **Classic NPI query-tool access is retired (April 1, 2026).** The Classic Research
  (live) and Classic Frozen databases are no longer queryable through the tool. The only
  surviving Classic product is the **Classic NPI Spreadsheet Report** (classic frozen
  dataset, **national and property-type returns only** — no Classic subtype or metro
  returns going forward).
- **There is now effectively one NPI.** Query it with `NPI_Plus = 1` and **full-name
  property types** (`'Industrial'`, `'Office'`, etc.). The single-letter Classic codes
  (`A`, `I`, `O`, `R`, `H`, `X`) are no longer the working values in the query tool.
- **Q1 2026 data was formally frozen.** Subsequent updates introduce: (a) **CBSA
  definitions realigned to the 2020 Census** (the codes in `references/geography.md`
  were 2010-Census — several metro division codes change; see that file), (b) data
  cleanup/corrections to fields with history back to the 1990s, and (c) **Industrial
  Outdoor Storage (IOS)** as a new `Design` option under the `Industrial: Specialized`
  subtype.
- **Membership flag:** the transitioned data dictionary exposes `NPI_Plus` (and
  `NPI_Plus_Lag1`) as the NPI membership flag. A standalone `NPI` flag / `NPILag*`
  fields are not in the new dictionary.

**Net effect on this skill:** keep using the API mechanics below with
`NPI_Plus = 1` and full-name property types as the default. Treat anything that relies
on Classic codes, the `NPI=1` flag, or DataTypeId 1/2 as legacy. (The NCREIF API User
Guide documents the `p_DataTypeId` mapping as `1` = Classic Research, `2` = Classic Frozen,
`3` = Expanded NPI — so `3` is the value for the NPI. Note the guide predates the
April 2026 Classic retirement, so it still lists the Classic datasets as queryable; in
practice DataTypeId 1/2 access is retired from the query tool.)

## Quick Reference

**API Base URL**: `https://qt-api.ncreif.org`

**Endpoints** (per the official NCREIF API User Guide). All accept a JSON request body and
use the `POST` method. The `/QT/*` (V1) endpoints return **XML**; the `/QT_V2/*` endpoints
return **JSON** (a `rows`/`fields` structure with datatypes) — except `GetExcel`, which
returns an Excel file.

| Endpoint | In | Out |
|----------|----|----|
| `/Login/Login` | `{"email","password"}` | `{"message":"<JWT>"}` |
| `/QT/ExecuteQuery` | JSON params | **XML** (field names as node names) |
| `/QT/ExecuteQueryFromXMLString` | JSON-wrapped XML string (`xmlQuery`) | **XML** |
| `/QT_V2/ExecuteQuery` | JSON params | **JSON** (`rows`/`fields` + datatypes) |
| `/QT_V2/ExecuteQueryGetExcel` | JSON params | **Excel file** |
| `/QT_V2/ExecuteQueryFromXMLString` | JSON-wrapped XML string (`xmlQuery`) | **JSON** |
| `/QT_V2/ExecuteQueryFromXMLFile` | `multipart/form-data` saved query-tool XML file | **JSON** |

**Pick the endpoint by the output you want:** `/QT_V2/ExecuteQuery` for JSON, `/QT/ExecuteQuery`
for XML (this is the endpoint the guide's Python example uses), `/QT_V2/ExecuteQueryGetExcel`
for an Excel download. The `FromXMLString`/`FromXMLFile` variants let you replay queries
saved/exported from the NCREIF query tool — element names match the exported files.

> The request body schema is the same for the param-based endpoints:
> `p_DataTypeId`, `p_SelectQuery`, `p_WhereClause`, `p_GroupbyClause`, `kpi`, `p_QueryData`.
> A live OpenAPI/SwaggerUI spec is also published at
> `https://qt-api.ncreif.org/swagger/index.html`.

**Rate Limits**: not specified in the API User Guide. In practice budget ~8s between
queries (≈8/min) during business hours to avoid throttling.

**Authentication**: Requires NCREIF username and password. Authenticate first via
`POST /Login/Login` with `{"email":"...","password":"..."}` (the guide's working code
examples capitalize the keys — `{"Email":"...","Password":"..."}` — and note that property
names and passwords are **case-sensitive**). The response is **JSON shaped like
`{"message":"<JWT>"}`** — the bearer token is the `message` field, NOT the raw response
body, and the response also carries token metadata, so extract the `message` value
specifically. Include it in subsequent requests as `Authorization: Bearer <token>`.
**Tokens expire after 20 minutes** — re-authenticate when a request returns HTTP 401.

### HTTP Status Codes
| Code | Meaning |
|------|---------|
| 200 | Query executed; results (if any) are in the response body |
| 400 | Incorrect endpoint |
| 401 | Token missing/expired, or account lacks query-tool access — re-login and retry |
| 500 | Query could not be executed (check SELECT/WHERE/GROUP BY syntax) |

## Query Parameters (JSON Payload)

Every query requires these parameters:

| Parameter | Required | Description |
|-----------|----------|-------------|
| `p_DataTypeId` | Yes | Database selector. `3` = transitioned NPI (the official NPI as of 2026Q1) — **use this**. `1`/`2` were Classic Research/Frozen and are retired from the query tool (April 1, 2026). |
| `p_SelectQuery` | Yes* | SQL-style SELECT aggregations (e.g., `SUM(NOI) / SUM(Denom) AS 'Income_Return'`) |
| `p_WhereClause` | Yes | Filter criteria (e.g., `NPI_Plus = 1 AND PropertyType = 'Residential'`) |
| `p_GroupbyClause` | No | Grouping fields, default `YYYYQ`. Comma-separated. Do NOT repeat in SELECT. |
| `p_QueryData` | No | Manager filter. `0` = all managers (default), `1` = my manager's properties, `2` = all properties except my manager's |
| `kpi` | No | Shortcut that auto-assigns the appropriate SELECT/WHERE for the named KPI. Official values per the API guide: `NPIClassic`, `NPIExpanded`, `CapRates`, `PercentLeased` (more KPIs to be added). Use `NPIExpanded` for the NPI; `NPIClassic` is legacy (Classic query-tool access is retired). |

*Not required if using `kpi` parameter.

### Syntax Rules
- Use single quotes for string values: `PropertyType = 'Office'` (never double quotes)
- Use single quotes for aliases: `AS 'Income_Return'` (never double quotes)
- Avoid spaces in aliases: use `Income_Return` not `Income Return` (spaces become `%20`)
- Field names in WHERE can use brackets: `[PropertyType] = 'Office'`
- Fields in GROUP BY should use brackets: `[YYYYQ],[PropertyType]`
- Do NOT include GROUP BY fields in SELECT — they appear automatically in results

## Database Selection Guide

Read `references/database-selection.md` for the full decision matrix. Post-transition
summary:

| DataTypeId | WHERE flag | Database | PropertyType values |
|------------|------------|----------|---------------------|
| 3 | `NPI_Plus=1` | **Transitioned NPI (the official NPI)** | Full names: Residential, Industrial, Office, Retail, Hotel, Self-Storage, Seniors Housing, Other, Land |
| 1 / 2 | `NPI=1` | Classic Research / Frozen — **RETIRED from query tool (Apr 1 2026)** | (legacy) A, I, O, R, H, X |

**Default:** `p_DataTypeId=3` with `NPI_Plus=1` and full-name property types. This is now
the official NPI, not just the "expanded" alternative.

## Common Query Templates

### NPI Returns
```json
{
  "p_DataTypeId": 3,
  "p_SelectQuery": "SUM(NOI) AS NOI, SUM(CapEx) AS CapEx, SUM(MV) AS MV, SUM(MVLag1) AS MVLag1, SUM(PSales) AS PSales, SUM(Denom) AS Denom, SUM(NOI)/SUM(Denom) AS 'Income_Return', (SUM(MV)-SUM(MVLag1)-SUM(CapEx)+SUM(PSales))/SUM(Denom) AS 'Capital_Return', (SUM(NOI)+SUM(MV)-SUM(MVLag1)-SUM(CapEx)+SUM(PSales))/SUM(Denom) AS 'Total_Return', COUNT(MV) AS 'Prop_Count'",
  "p_WhereClause": "NPI_Plus=1",
  "p_GroupbyClause": "Year, YYYYQ",
  "p_QueryData": 0,
  "kpi": ""
}
```

### NPI Returns by Property Type
Same SELECT as above, add to GROUP BY: `Year, YYYYQ, [PropertyType]`
WHERE: `NPI_Plus=1`

### ODCE Properties (100% / unweighted)
WHERE: `NPI_Plus=1 AND FundType = 'D'`

### ODCE Returns "at Share" (official NCREIF template — weight by LegalPropertyShare)
NCREIF's official "ODCE Returns at Share" template weights **every component** by
`LegalPropertyShare` (not a flat 100%). This matches the ownership-share convention used
for fund questionnaires.
```
Income Return  = SUM(NOI*LegalPropertyShare) / SUM(Denom*LegalPropertyShare)
Capital Return = (SUM(MV*LegalPropertyShare) - SUM(MVLag1*LegalPropertyShare)
                 - SUM(CapEx*LegalPropertyShare) + SUM(PSales*LegalPropertyShare))
                 / SUM(Denom*LegalPropertyShare)
Total Return   = Income Return + Capital Return
WHERE: NPI_Plus=1 AND FundType='D'
```
(`EffectivePropertyShare` is also available if you need effective rather than legal share.)

### Appraisal Cap Rates (official NCREIF template)
NCREIF does **not** simply average the pre-calc `AppCapRate` field. The official template
annualizes quarterly NOI over value (×4), adds partial sales to value, applies an outlier
qualifier, and excludes hotels:
```
EqWtd Cap Rate  = ROUND(AVG(NOI/(MV+PSales+1))*4, 4)
ValWtd Cap Rate = ROUND(SUM(NOI)/(SUM(MV)+SUM(PSales))*4, 4)
WHERE: NPI_Plus=1 AND PropertyType <> 'Hotel'
       AND Abs((NOI/(MV+PSales+1))*4) < .30      -- tighten below national level
```
(The `+1` avoids divide-by-zero; tighten the .30 outlier band for sub-national cuts. The
pre-calc `AppCapRate`/`AppCapRate4QNOI` fields remain available as a simpler alternative.)

### Transaction Cap Rates (official NCREIF template — sold properties only)
Uses `NOILag1` because only a partial quarter of NOI is available in the sale quarter
post-freeze; tighter outlier band than appraisal cap rates:
```
EqWtd Cap Rate  = ROUND(AVG(NOILag1/(MV+PSales+1))*4, 4)
ValWtd Cap Rate = ROUND(SUM(NOILag1)/(SUM(MV)+SUM(PSales))*4, 4)
WHERE: NPI_Plus=1 AND PropertyType <> 'Hotel' AND SaleQtr = 1
       AND (NOILag1/(MV+PSales+1))*4 >= .02
       AND (NOILag1/(MV+PSales+1))*4 <= .14
```

### Percent Leased / Occupancy
```
p_SelectQuery: Avg(PercentLeased) AS Occupancy, Count(PercentLeased) AS Props
p_WhereClause: PercentLeased is not Null and NPI_Plus = 1
```

### NOI Growth (official NCREIF template)
Excludes post-freeze partial-sale quarters (partial NOI understates growth) and requires a
prior-quarter NOI:
```
p_SelectQuery: Sum(NOI) AS NOI, Sum(NOILag1) AS NOILag1, (Sum(NOI)/Sum(NOILag1))-1 AS 'NOI_Growth', Count(NOILag1) AS 'Prop_Count'
p_WhereClause: NPI_Plus=1 AND PartialSaleQtr = 0 AND NOILag1 Is Not Null
```

### Cash Flow Returns (official NCREIF template)
CapEx is subtracted from NOI (not from the ending value), so the price-change term carries
**no** CapEx deduction. Total return equals the standard NPI total return.
```
CashFlow Return = (SUM(NOI)-SUM(CapEx)) / SUM(Denom)
Price Change    = (SUM(MV)-SUM(MVLag1)+SUM(PSales)) / SUM(Denom)
Total Return    = (SUM(NOI)-SUM(CapEx)+SUM(MV)-SUM(MVLag1)+SUM(PSales)) / SUM(Denom)
WHERE: NPI_Plus=1
```

### Market Value Indicators — MVI / FCFY / CXR (official NCREIF template)
These exclude property-quarters with substantial capex via the `MVIFLAG` filter:
```
MVI  = (SUM(MV)-SUM(MVLag1)+SUM(PSales)) / SUM(MVLag1)
FCFY = (SUM(NOI)-SUM(CapEx)) / SUM(MVLag1)
CXR  = SUM(CapEx) / SUM(MVLag1)
WHERE: NPI_Plus=1 AND MVIFLAG = 1
```
(Equal-weighted variant: wrap each in `AVG(...)` and add the supplied `*_stdev` columns.)

### Equal-Weighted NPI Returns (official NCREIF template)
Averages each property's return rather than value-weighting; includes cross-sectional
standard deviation:
```
p_SelectQuery: Avg(IncRet) AS AveIncReturn, Avg(AppRet) AS AveAppReturn, Avg(TotRet) AS AveTotReturn, StDev(TotRet) AS StdDevofTotRet, Count(TotRet) AS 'Prop_Count'
p_WhereClause: NPI_Plus=1
```

### Sold Properties (by Year or Quarter)
```
p_SelectQuery: Count(Prop) AS 'Prop_Count', Sum(SalePrice) AS TotalSale, Avg(SalePrice) AS AvgSale
p_WhereClause: NPI_Plus=1 AND SaleQtr = 1
p_GroupbyClause: Year      (or YYYYQ for quarterly; note 20024 returns no row — freeze quarter)
```

### Leveraged Returns
```
p_SelectQuery: SUM(NOI-Interest-Principal+LoanProceeds) AS LevNOI, SUM(LevDenom) AS LevDenom, (SUM(NOI-Interest-Principal+LoanProceeds))/SUM(LevDenom) AS 'Lev_Income_Return'
p_WhereClause: NPI_Plus=1 AND INCL_LV=1
```

## Return Formulas (NPI)

These are the standard NCREIF return calculations:

- **Income Return** = `SUM(NOI) / SUM(Denom)`
- **Capital Return** = `(SUM(MV) - SUM(MVLag1) - SUM(CapEx) + SUM(PSales)) / SUM(Denom)`
- **Total Return** = `(SUM(NOI) + SUM(MV) - SUM(MVLag1) - SUM(CapEx) + SUM(PSales)) / SUM(Denom)`
  - PSales (partial sales proceeds) must be included in the total return numerator
- **Denom** = `MVLag1 + 0.5*CapEx - PSales` (pre-calculated in the database)

For **Cash Flow Returns**, replace NOI with `NOI - CapEx`:
- **Cash Flow Income Return** = `SUM(NOI - CapEx) / SUM(Denom)`

## Trailing Period Returns and Aggregation

**CRITICAL**: The API returns single-quarter returns. When the user asks for trailing
period returns (1-year, 3-year, 5-year, 10-year, etc.), you MUST:

1. **Pull the full quarterly time series** — e.g., 40 quarters for a 10-year return
2. **Geometrically link** (time-weighted return) — multiply wealth factors `(1 + r)`, NOT arithmetic average
3. **Annualize** for periods > 1 year — raise to the power of `(1/years)`

**Formulas:**
- Cumulative: $R = \prod_{t=1}^{N}(1+r_t) - 1$
- Annualized: $R_{ann} = (1+R_{cum})^{4/N} - 1$ where N = number of quarters

**Quarter counts**: 1-year = 4, 3-year = 12, 5-year = 20, 10-year = 40

**Convention**: QTD and YTD are reported as cumulative (not annualized). 1-year is
cumulative (which equals annualized). 3-year+ are annualized.

Read `references/return-aggregation.md` for complete methodology, Python/Excel
implementation code, query strategy, and edge cases.

## Property Types and Subtypes

Read `references/property-types.md` for the complete mapping.

### Transitioned NPI (DataTypeId=3, NPI_Plus=1) — the official NPI
PropertyType values (full names): `Residential`, `Industrial`, `Office`, `Retail`, `Hotel`, `Self-Storage`, `Seniors Housing`, `Other`, `Land`

PropertySubType examples:
- Residential: `Residential: Apartment`, `Residential: Student Housing`, `Residential: Single Family Rental`, `Residential: Manufactured Housing`
- Industrial: `Industrial: Warehouse`, `Industrial: Flex`, `Industrial: Life Science`, `Industrial: Manufacturing`, `Industrial: Specialized`
- Office: `Office: Central Business District`, `Office: Suburban`, `Office: Urban`, `Office: Medical Office`, `Office: Secondary Business District`, `Office: Life Science`
- Retail: `Retail: Mall`, `Retail: Strip`, `Retail: Street`

New `Design` option coming post-freeze: **Industrial Outdoor Storage (IOS)** under the
`Industrial: Specialized` subtype.

> NCREIF's own loaded templates sometimes reference `Exp_PropertyType` /
> `Exp_PropertySubType` for type/subtype filters (e.g. `Exp_PropertySubType =
> 'Residential: Apartment'`). The plain `PropertyType` / `PropertySubType` columns now
> carry the same full-name values. If a filter silently returns unfiltered results, try
> the `Exp_`-prefixed name (and verify spelling — a misspelled field is ignored, not
> errored). Spreadsheet reports spell it `Self Storage`; the query tool uses `Self-Storage`.

### Classic NPI codes (legacy — query-tool access retired Apr 1 2026)
Single-letter PropertyType codes `A` (Apartment), `I` (Industrial), `O` (Office),
`R` (Retail), `H` (Hotel), `X` (Other) and the 2–3 letter subtype codes survive only in
the **Classic NPI Spreadsheet Report** (national + property-type returns). They are no
longer the working values for query-tool calls.

## Geographic Filtering

Read `references/geography.md` for full CBSA code tables.

### Key Fields
- `Region`: E (East), M (Midwest), S (South), W (West)
- `Division`: EN, ME, NE, SE, SW, WM, WN, WP
- `State`: Two-letter code, e.g., `'CA'`, `'NY'`, `'TX'`
- `CBSAorDiv`: Preferred CBSA field — uses division code where one exists, CBSA code otherwise
- `CBSA`: Always the full CBSA (ignores divisions)
- `CBSADiv`: Division code only (NULL if no division)
- `City`, `Zip`, `County`

### Common CBSA Codes (CBSAorDiv)
| Market | CBSAorDiv |
|--------|-----------|
| New York (NYC proper) | 35614 |
| Nassau-Suffolk (Long Island) | 35004 |
| Los Angeles | 31084 |
| Orange County (Anaheim) | 11244 |
| Chicago | 16974 → **16984** (2020) |
| San Francisco | 41884 |
| Oakland | 36084 |
| San Jose / Silicon Valley | 41940 |
| Washington DC | 47894 → **47764** (2020) |
| Boston | 14454 |
| Dallas | 19124 |
| Houston | 26420 |
| Atlanta | 12060 → **12054** (2020) |
| Phoenix | 38060 |
| Denver | 19740 |
| Seattle | 42644 |
| Miami | 33124 |
| Fort Lauderdale | 22744 |
| West Palm Beach | 48424 |
| San Diego | 41740 |
| Minneapolis | 33460 |
| Detroit | 19804 |
| Philadelphia | 37964 |
| Tampa | 45300 → **45294** (2020) |
| Orlando | 36740 |
| Portland OR | 38900 |
| Austin | 12420 |
| Nashville | 34980 |
| Charlotte | 16740 |
| Raleigh | 39580 |
| Indianapolis | 26900 |
| Sacramento | 40900 |
| Salt Lake City | 41620 |
| Las Vegas | 29820 |
| Silver Spring / MD suburbs | 43524 → **23224** (2020) |

For the full CBSA table, read `references/geography.md`.

> **CBSA vintage (2026Q1 transition):** the codes above are **2010-Census** vintage, which
> matches the query tool today. The transitioned NPI is realigning to **2020-Census** CBSA
> definitions post-freeze. A handful of major-market division codes change — e.g.
> **Chicago 16974→16984, Washington DC 47894→47764, Atlanta 12060→12054, Tampa
> 45300→45294, and the old MD-suburbs division 43524 is replaced** (Frederick-Gaithersburg-
> Bethesda = 23224). Most other metros keep the same code (only the name label changes).
> See `references/geography.md` for the old→new mapping and verify which vintage your
> quarter's data uses with a live query.

## Advanced WHERE Clause Patterns

### Filter by Size
```
[SqFt] > 1000 AND [SqFt] < 100000
```

### Filter by Value per SF
```
MV / [SqFt] > 300 AND [SqFt] > 1000
```

### Occupancy Range (Core Properties)
```
[PercentLeased] > .90
```

### Vintage Year Analysis
```
[YYYYQ] > 20051 AND [StartDate] >= 20051 AND [StartDate] <=20063 AND [AcqDate] >=200501 AND [AcqDate] <=200609 AND [SaleCode] <> 'S' AND NPI_Plus=1
```

### Same Store / Constant Properties
```
[YYYYQ] >= 20191 AND [YYYYQ] <= 20244 AND [StartDate] < 20191 AND [EndDate] >= 20244
```

### Exclude Sale Quarters
```
SaleQtr = 0
```

### Leveraged vs Unleveraged
- Only unleveraged: `[INCL_LV] IS NULL`
- Only leveraged: `[INCL_LV] = 1`
- High leverage: `[BalanceLag1] / [MVLag1] > .90`

### Exclude Outlier Returns
```
Abs(TotRet) < 0.05
```
Or asymmetric: `[TotRet] > -0.05 AND [TotRet] < 0.075`

### Recently Built Properties (< 5 years old)
```
[Year]-[YrBuilt]<=5
```

### Fund Types
- ODCE: `FundType = 'D'`
- All Open End: `(FundType = 'O' OR FundType = 'D')`
- Separate Accounts: `FundType = 'S'`
- Closed End: `FundType = 'C'`

### Combining AND/OR with Parentheses
```
((PropertyType='Office' AND PropertySubtype = 'Office: Suburban') OR PropertyType = 'Residential') AND State = 'CA' AND NPI_Plus = 1
```

### Time Period Filtering
YYYYQ format is a 5-digit number: year + quarter (e.g., 20241 = Q1 2024)
```
[YYYYQ] >= 20201 AND [YYYYQ] <= 20244
```
For "last 5 years" from current quarter, calculate the appropriate YYYYQ value.

### Expanded NPI Specific Fields
- `PropertyType` — Use full names: `'Residential'`, `'Industrial'`, `'Office'`, `'Retail'`, `'Hotel'`, `'Self-Storage'`, `'Seniors Housing'`, `'Other'`
- `PropertySubType` — Full names: `'Residential: Apartment'`, `'Industrial: Warehouse'`, etc.
- `Usage` — Retail usage: `'Retail: High-End with Grocer'`, `'Retail: Not High-End without Grocer'`, etc.
- `Clusters` — Apartment subtypes: `'Garden'`, `'Low-Rise'`, `'Mid-Rise'`, `'High-Rise'`
- `Design` — Detailed design types

### Property type fields (transitioned NPI)
- `PropertyType` — full names: `'Residential'`, `'Industrial'`, `'Office'`, `'Retail'`, `'Hotel'`, `'Self-Storage'`, `'Seniors Housing'`, `'Other'`, `'Land'`
- `PropertySubType` — full names: `'Residential: Apartment'`, `'Industrial: Warehouse'`, etc.
- `Exp_PropertyType` / `Exp_PropertySubType` — alternate names NCREIF's own templates use; same full-name values
- Classic single-letter codes (`'A'`, `'I'`, `'O'`, `'R'`, …) are legacy (query-tool access retired)

## GROUP BY Patterns

Default: `Year, YYYYQ` or `[Period],[YYYYQ],[Year],[Quarter]`

Common additions:
- By property type: `Year, YYYYQ, [PropertyType]`
- By geography: `Year, YYYYQ, [CBSAorDiv]`
- By subtype: `Year, YYYYQ, [PropertyType], [PropertySubType]`
- By usage (retail): `[Period],[YYYYQ],[Year],[Quarter],[Usage]`
- By clusters (apartments): `[Period],[YYYYQ],[Year],[Quarter],[PropertySubType],[Clusters]`
- Annual only: `Year`

## Building Queries: Step-by-Step

1. **Database**: Use the transitioned NPI (DataTypeId=3, `NPI_Plus=1`). Classic (1/2) is retired.
2. **Determine the metric**: Returns, cap rates, occupancy, NOI growth, etc.
3. **Build the SELECT**: Use the appropriate official template above.
4. **Build the WHERE**: Start with `NPI_Plus=1`, then add property type (full names), geography, time period, and other filters.
5. **Build the GROUP BY**: Start with time period (Year, YYYYQ), add any dimensions the user wants to see.
6. **Check for common pitfalls**:
   - Don't use AND to combine multiple property types — use OR or IN
   - Include `[SqFt] > 0` when filtering by square footage (some properties don't report it)
   - Include `PercentLeased > 0` when filtering by occupancy (0 may mean unreported)
   - Remember AcqDate is YYYYMM format, not YYYYQ
   - Masking criteria: results need ≥3 properties from ≥3 managers per group

## Masking Criteria

NCREIF applies masking to protect confidentiality. Results are only returned when a group
contains at least **3 properties from at least 3 different managers**. If a query is too
specific (narrow geography + property type + time), some quarters may return no data.

## Implementation Notes

### Python Example
```python
import requests

BASE_URL = "https://qt-api.ncreif.org"

# Step 1: Authenticate
login_resp = requests.post(
    f"{BASE_URL}/Login/Login",
    json={"Email": "your_email", "Password": "your_password"},
    headers={"Content-Type": "application/json"}
)
token = login_resp.json()["message"]  # API returns {"message": "<JWT>"}

# Step 2: Execute Query
headers = {
    "Authorization": f"Bearer {token}",
    "Content-Type": "application/json"
}
payload = {
    "p_DataTypeId": 3,
    "p_SelectQuery": "SUM(NOI)/SUM(Denom) AS 'Income_Return', (SUM(MV)-SUM(MVLag1)-SUM(CapEx)+SUM(PSales))/SUM(Denom) AS 'Capital_Return', (SUM(NOI)+SUM(MV)-SUM(MVLag1)-SUM(CapEx)+SUM(PSales))/SUM(Denom) AS 'Total_Return', COUNT(MV) AS 'Prop_Count'",
    "p_WhereClause": "NPI_Plus=1",
    "p_GroupbyClause": "Year, YYYYQ",
    "p_QueryData": 0,
    "kpi": ""
}
resp = requests.post(
    f"{BASE_URL}/QT_V2/ExecuteQuery",
    json=payload,
    headers=headers
)
# /QT_V2/ExecuteQuery returns JSON as {"rows":[{"fields":[{"name","value","datatype"},...]}]}.
# Flatten it into records:
rows = resp.json().get("rows", [])
records = [{f["name"]: f["value"] for f in row["fields"]} for row in rows]
# import pandas as pd; df = pd.DataFrame(records)

# Alternatively use /QT/ExecuteQuery (V1), which returns XML with field names as node
# names — this is the endpoint the API User Guide's Python sample uses
# (xml.etree.ElementTree to parse). Or /QT_V2/ExecuteQueryGetExcel for an Excel file.
```

### Excel Power Query
Read `references/power-query.md` for the full M code template that reads query
parameters from a table and authenticates automatically.

## Reference Files

For detailed field definitions, property type codes, CBSA lookups, and implementation
examples, read the appropriate reference file:

- `references/field-definitions.md` — All data fields with types and descriptions
- `references/property-types.md` — Complete property type/subtype/usage/design/cluster values
- `references/geography.md` — Full CBSA code table, regions, divisions, states
- `references/database-selection.md` — DataTypeId decision matrix and NPI vs NPI_Plus logic
- `references/return-aggregation.md` — **Time-weighted return methodology**: geometric linking, annualization, trailing period calculations, Python/Excel code
- `references/power-query.md` — Excel Power Query M code template for batch queries
- `references/faq.md` — Common questions and troubleshooting
