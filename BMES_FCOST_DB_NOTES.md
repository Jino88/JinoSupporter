# BMES F-COST DB Notes

Last updated: 2026-06-17

이 문서는 `BMES > Test` 화면에서 BOM, 자재명, 단가를 읽기 위해 확인한 BMES DB 탐색 이력과 테이블 역할을 정리한 것이다. BMES 서버에는 절대 쓰기 작업을 하지 않는다. 앱 쿼리는 `SELECT`만 사용하고, SQL Server 연결 문자열에는 `ApplicationIntent=ReadOnly`를 넣는다.

## Connection

- BMES DB host checked during exploration: `10.6.0.21`
- Database: `BMES_LIV`
- Local app DB stores the BMES connection setting under key `BmesFcostTest.DbConnection`.
- Do not record the BMES password, session cookie, or SSO token in this file.

## BMES Web Endpoint Clues

- F-COST screen: `/MES050027`
- Main data endpoint: `GET /MES050027/SearchList`
- Exchange rate endpoint: `POST /MES050027/SearchListExchangeRate`
- `SearchList` returns summary rows in `data.contents`.
- `SearchList` also returns detailed grid data as a JSON string in `data.verify`. After parsing, rows include fields such as `Product Code`, `Product Name`, `Material Code`, `Material Name`, `Work Date`, `Standard Price(VND)`, `Actual Input Price(VND)`, and `F-COST(VND)`.
- `SearchListExchangeRate` returns rows with `GDATU`, `FCURR`, and `ZEXCH`. Example currencies observed: `CNY`, `EUR`, `GBP`, `JPY`, `USD`, `VND`.

## Table Roles

### `dbo.MATE`

- Preferred material/product name source.
- Important columns:
  - `MATNR`: material or product code.
  - `MAKTX`: material or product name.
- Verified examples:
  - `P-S-151100800` -> `TIU-C11-20-R-ZZ`
  - `C-S-000005300` -> `ASSY FRAME(C11-20-R)`
  - `M-P-027044900` -> `SHTW 0.036 KTP18`
- Current app name-source preference: `MATE.MAKTX`, then `MAKT`, then `BGPD`, then `ANH_TEXT`.

### `dbo.BGPD`

- Also has names such as `MAKTX`, but it was incomplete for several material codes.
- Useful as fallback only.

### `dbo.MAST`

- BOM header/link table.
- Important columns:
  - `MATNR`: product or assembly material code.
  - `WERKS`: plant.
  - `STLAN`: BOM usage/category.
  - `STLNR`: BOM number.
  - `STLAL`: BOM alternative.
- Current app starts from `MAST`:
  - If Product Code is typed, `MAST.MATNR = @ProductCode`.
  - If Product Code is blank, model names from local `modelGroup` are matched to `MATE.MAKTX`, then those matched `MATNR` values are used as roots.

### `dbo.BOMC`

- BOM component line table.
- Join key from `MAST`:
  - `BOMC.STLNR = MAST.STLNR`
  - `BOMC.STLAL = MAST.STLAL`
  - `BOMC.STLAN = MAST.STLAN`
- Important columns:
  - `POSNR`: BOM line number.
  - `CMATE`: child material code.
  - `MENGE`: usage quantity.
  - `MEINS`: usage unit.
  - `SDATE`, `EDATE`: BOM validity dates.
  - `USEYN`: active flag.
- Current date filter:
  - `SDATE <= @WorkDate`
  - `EDATE >= @WorkDate`
  - `ISNULL(USEYN, 'Y') = 'Y'`
- If a component code starts with `C-`, the app treats it as a possible assembly and recursively reads its child BOM from `MAST` + `BOMC`.

### `dbo.INFR`

- Current app price source.
- Important columns:
  - `MATNR`: material code.
  - `EKORG`: purchase organization. Often plant-like values such as `3210`, `1110`.
  - `LIFNR`: vendor.
  - `KBETR`: price amount.
  - `KPEIN`: price unit denominator.
  - `KONWA`: currency.
  - `DATAB`, `DATBI`: price validity dates, stored as `yyyyMMdd` style values.
  - `KNUMH`, `INFNR`: condition/info record references.
- Current app price lookup:
  - `INFR.MATNR = BOM material code`
  - `DATAB <= @WorkDateKey`
  - `DATBI >= @WorkDateKey`
  - Prefer same `EKORG` as the BOM plant, then same first two digits, then newest valid price.
- UI meaning:
  - `Standard Price` displays `KBETR / KPEIN KONWA`.
  - `1 EA Price` displays `KBETR / KPEIN`.
- Example for `M-P-027044900` around `20260616`:
  - `EKORG=3210`, `KBETR=41.17`, `KPEIN=100`, `KONWA=USD`, so one-unit price is `0.4117 USD`.
  - `EKORG=1110`, `KBETR=382`, `KPEIN=1`, `KONWA=KRW`, so one-unit price is `382 KRW`.

### `dbo.A018`, `dbo.A018_LOG`, `dbo.A018_LOG_2`

- Condition header/history tables found during exploration.
- Important columns:
  - `MATNR`
  - `KNUMH`
  - `DATAB`, `DATBI`
  - `LIFNR`, `EKORG`
- These can map a material to a condition number, but the current app does not depend on them because `INFR` already exposes `MATNR`, `KBETR`, `KPEIN`, and `KONWA` directly.

### `dbo.COST`

- Cost center master, not material price.
- Columns observed: `KOKRS`, `KHINR`, `KOSTL`, `KTEXT`, `KOSAR`, `FUNC_AREA`.
- Do not use for BOM material standard price.

### `dbo.PRCR`

- Purchase requisition/status style table, not material price.
- Columns observed: `PURNO`, `PURSQ`, `USEYN`, `BANFN`, `BNFPO`, `IFSTA`, `IFMSG`, `IFDAT`.
- Do not use for BOM material standard price.

## Read-Only Discovery Queries

Find tables containing a column:

```sql
USE BMES_LIV;

SELECT
    s.name AS schema_name,
    t.name AS table_name,
    c.name AS column_name
FROM sys.tables t
JOIN sys.schemas s ON s.schema_id = t.schema_id
JOIN sys.columns c ON c.object_id = t.object_id
WHERE c.name = N'KNUMH'
ORDER BY t.name;
```

Check current price rows for one material:

```sql
USE BMES_LIV;

SELECT TOP (20)
    EKORG,
    MATNR,
    KMEIN,
    LIFNR,
    KBETR,
    KPEIN,
    KONWA,
    DATAB,
    DATBI,
    KNUMH,
    INFNR
FROM dbo.INFR
WHERE MATNR = N'M-P-027044900'
  AND DATAB <= N'20260616'
  AND DATBI >= N'20260616'
ORDER BY DATAB DESC;
```

Read one product BOM:

```sql
USE BMES_LIV;

SELECT TOP (100)
    m.MATNR AS ProductCode,
    m.WERKS AS Plant,
    m.STLNR AS BomNo,
    m.STLAL AS BomAlt,
    b.POSNR AS BomLine,
    b.CMATE AS MaterialCode,
    b.MENGE AS UsageQty,
    b.MEINS AS UsageUnit,
    b.SDATE,
    b.EDATE
FROM dbo.MAST m
JOIN dbo.BOMC b
    ON b.STLNR = m.STLNR
   AND b.STLAL = m.STLAL
   AND b.STLAN = m.STLAN
WHERE m.MATNR = N'P-S-151100800'
  AND b.SDATE <= CONVERT(date, N'20260616', 112)
  AND b.EDATE >= CONVERT(date, N'20260616', 112)
  AND ISNULL(b.USEYN, N'Y') = N'Y'
ORDER BY b.POSNR;
```

## Current App Flow

Related files:

- `JinoSupporter.Web/Components/Pages/BmesFcostTestPage.razor`
- `JinoSupporter.Web/Services/BmesFcostMaterialPriceService.cs`

Flow:

1. User enters BMES DB connection and Work Date.
2. If Product Code is entered, the service reads that product BOM.
3. If Product Code is blank, the page reads local modelGroup middle-group Material names and searches BMES products whose `MATE.MAKTX` matches those names.
4. The service reads root BOM lines from `MAST` + `BOMC`.
5. For every component starting with `C-`, the service recursively reads child BOM lines.
6. Names are read from `MATE.MAKTX` if available.
7. Prices are read from `INFR` by material and work date.
8. The page displays BOM columns separately: Plant, BOM No, Alt, Line, Usage Qty, Unit.

Blank Product Code matching:

- Source names are local `ModelGroups > MidGroups.Material`, which corresponds to NG Rate `MATERIALNAME`.
- The Test page also includes a minimum required seed list for models that must appear even if the local group data is incomplete:
  `TIU-C11-20-R-ZZ`, `MSM-X626B-TOP`, `MSM-X526B-TOP`, `SI-SM-G736B`, `MSM-S936U`, `NSM-NX14-L`, `NSM-NX14-R`, `MSM-S931B`, `SI-SM-S908U`, `ASSY REAR...338...`, `TIU-C11-20-L-ZZ`, `TIU-L5S3-01-L-ZZ`, `TIU-L5S3-01-R-ZZ`, `BRS-161016S08ZZ`, `MSU-L20S15-07`.
- BMES name comparison is separator-tolerant: spaces, `-`, `_`, `.`, and `/` are removed before comparing against `MATE.MAKTX`.
- This is still a deterministic match, not a broad fuzzy search.

Not implemented yet:

- Full MES050027 F-COST total reconciliation.
- Currency exchange conversion to VND for every price row.
- Actual input price comparison.
