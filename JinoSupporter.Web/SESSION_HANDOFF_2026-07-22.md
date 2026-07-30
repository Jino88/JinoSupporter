# JinoSupporter.Web BMES Report 정적 HTML 세션 handoff — 2026-07-22

이 문서는 다음 세션이 BMES Report 탭 HTML화 작업을 재조사 없이 이어가기 위한 재개
지점이다. `InferenceDataAIService/SESSION_HANDOFF_2026-07-21.md`와는 별개의 작업이다.

## 1. 목적

`/report/bmes` 페이지의 7개 탭을 Blazor 라이브 컴포넌트가 아니라 **자체 완결형 정적
HTML 한 파일**로 생성해 iframe에 띄운다.

이유는 Chrome 메모리 문제다. 기존에는 7개 리포트가 동시에 살아있는 Blazor 컴포넌트로
유지돼 브라우저 메모리가 터졌다. 현재 구조는 탭 본문을 전부 비활성 `<template>`에
넣어두고, 탭 클릭 시 그 하나만 DOM에 clone해 붙이고 이전 것은 버린다. 따라서 어느
시점에도 리포트 하나만 레이아웃/페인트된다.

## 2. 구조

- 생성기: `Services/BmesReportHtmlExportService.cs`
- 페이지: `Components/Pages/BmesReportPage.razor`
  - Report Setup / Progress 패널만 Blazor로 남기고, 결과는
    `/report/bmes/view/{token}` iframe 하나로 표시한다.
- 출력 위치: `AppStoragePaths.Combine("_temp", "bmes-report")/{token}/report.html`
  - token은 32자리 hex. 생성 시 이전 token 폴더를 전부 지운다.
  - `ResolveReportFile`에 token 검증과 path-escape 가드가 있다.
- 렌더링 방식: Blazor 독립형 `HtmlRenderer`로 각 탭 컴포넌트를 export 모드로 렌더한다.

## 3. 탭과 컴포넌트 매핑

메뉴 순서는 서비스의 `Tabs` 배열이 정한다.

| 탭 키 | 라벨 | 컴포넌트 | 데이터 |
|---|---|---|---|
| `daily` | Daily | `NgRateForDailyReportPage` | BMES 실제 조회 (유일) |
| `weekly` | Weekly | `NgRateForWeeklyReportPage` | Daily의 hierarchy 재사용 |
| `cause-monthly` | 원인 비중 | `BmesCauseMonthlyReportPage` | Daily의 hierarchy 재사용 |
| `fcost` | F-COST | `BmesFCostPage` | 이 탭만 F-COST 리포트 계산 |
| `fcost-all` | F-COST(전체) | `BmesFCostPage` | `fcost` 스냅샷 재사용 |
| `fcost-weekly` | 목표 불량률 | `BmesFCostPage` | `fcost` 스냅샷 재사용 |
| `fcost-weekly-all` | 목표 불량률(전체) | `BmesFCostPage` | `fcost` 스냅샷 재사용 |

렌더 순서는 `daily → cause-monthly → weekly → fcost ×4`이며 progress stage 순서와
맞춘다. 표시 순서는 `Tabs` 배열 순서와 별개다.

F-COST 4개 탭의 플래그 조합은 서비스의 `FCostVariants`에 있다.

| 탭 키 | `ShowFCostWeeklyReport` | `ShowAllPeriodColumns` | `ShowRegularFCostDetails` |
|---|---|---|---|
| `fcost` | false | false | true |
| `fcost-all` | false | true | true |
| `fcost-weekly` | true | false | false |
| `fcost-weekly-all` | true | true | false |

## 4. 확정된 설계 원칙

### BMES 조회는 리포트 전체에서 1회

Daily만 실제로 스크레이프하고, `OnExportComputed`로 두 가지를 밖으로 내보낸다.

- `HierReports` — 원인 비중, Weekly가 재사용
- `NgRateReportService.NgRateReport` — F-COST가 `ExportSharedNgTrend`로 재사용

탭마다 다시 조회하지 않는다.

### F-COST 리포트 계산은 4개 탭에서 1회

`BmesFCostPage`를 그냥 4번 렌더하면 `BackfillRawAsync` + `GenerateRawRangeReportAsync`
+ `BuildRawMaterialBreakdownAsync`가 4번 돈다. 이를 막기 위해 `FCostExportSnapshot`
레코드를 도입했다.

- 첫 변형(`fcost`)만 계산하고 `OnExportComputed`로 스냅샷을 넘긴다.
- 나머지 3개는 `ExportSharedReport`를 받아 데이터 접근 없이 렌더만 한다.
- 스냅샷에 담기는 것: `Report`, `NgRateTrendByMid`, `ModelGroups`,
  `RawMaterialBreakdown`, `RawStatus`, `StartDate`, `EndDate`

### 정적 스냅샷에서는 접기/펼치기를 전부 펼친다

2026-07-21 사용자 판단으로 확정했다. 정적 HTML에는 누를 토글이 없으므로 접힌 행은
영영 도달 불가다. 따라서 export 모드에서만 모든 계층을 펼친다. 표가 길어지는 것은
감수한다.

| 컴포넌트 | 펼치는 대상 | 진입점 |
|---|---|---|
| `BmesFCostPage` | 계층 행 walker, Trend 행, Raw breakdown 3단계 | `IsHierarchyRowExpanded`, `IsFCostRowVisible` |
| `NgRateForWeeklyReportPage` | 행 walker, `_expandedGroups`, `_expandedGroupsModel`, `_expandedMids` | `IsSectionExpanded`, `IsWrRowVisible` |
| `NgRateForDailyReportPage` | Reason 섹션 | `IsReasonSectionCollapsed` |

기간 컬럼도 같은 이유로 export 시 전부 펼친다.

- `BmesFCostPage.IsPeriodExpanded` → `ExportMode || ...`
- `NgRateForWeeklyReportPage`의 `VisibleWeekCols` / `VisibleMonthCols` → `ExportMode || ...`
- `Top10*Cols`는 의도된 고정 cap이므로 건드리지 않았다.

### 렌더 모드를 선언한 컴포넌트는 정적 렌더 불가

독립형 `HtmlRenderer`는 렌더 모드를 지원하지 않는다. `@rendermode`를 선언한 컴포넌트를
넘기면 다음 오류로 실패한다.

```
Cannot supply a component of type '...BmesFCostPage' because the current platform
does not support the render mode 'InteractiveServerRenderMode'.
```

`BmesFCostPage`가 유일하게 `@page` + `@rendermode`를 갖고 있어서 이 문제가 났다.

- `Components/Pages/BmesFCost.razor`를 새로 만들어 `@page "/bmes/f-cost"`와
  `@rendermode InteractiveServer`를 옮겼다. 이 wrapper가 `<BmesFCostPage />`를 감싼다.
- `BmesFCostPage`는 나머지 3개 탭 컴포넌트처럼 라우트 없는 순수 컴포넌트가 됐다.
- 메뉴로 들어가는 `/bmes/f-cost`는 이전과 동일하게 인터랙티브하다.

**앞으로 export 대상 컴포넌트에 `@rendermode`를 다시 붙이면 안 된다.**

### 차트는 JS interop이 아니라 정적 payload로 넘긴다

`HtmlRenderer`는 `OnAfterRenderAsync`를 호출하지 않는다. Weekly는 캔버스 2개
(`ngrateByGroupCanvas`, `ngrateByGroupModelCanvas`)를 `OnAfterRenderAsync`의 JS
interop으로 그리고 있어서 정적 export에서는 빈 캔버스가 나갔다.

생성기 wrapper에는 원래부터 `initCharts`가 있었지만, 그것이 읽을
`<script type="application/json" data-chart>` payload를 내보내는 쪽이 어디에도 없어서
한 번도 동작한 적이 없는 코드였다.

이번에 다음을 추가했다.

- `BuildGroupTrendChartData()` — labels/series 생성부를 분리. 인터랙티브 경로와
  export 경로가 공유한다.
- `ExportChartPayload(canvasId)` — export 모드에서만 `MarkupString`으로 payload를 emit.
  - `<script>` 본문은 raw text라 HTML 인코딩된 따옴표가 디코드되지 않아 `JSON.parse`가
    깨진다. 그래서 `MarkupString`으로 내보낸다.
  - System.Text.Json 기본 encoder가 `<`, `>`를 이스케이프하므로 그룹명이 script 태그를
    조기 종료시킬 수 없다.

**차트를 가진 탭을 추가할 때는 같은 방식을 따라야 한다.**

## 5. Export 모드 파라미터 규약

각 탭 컴포넌트는 `ExportMode`가 켜지면 `OnParametersSetAsync`에서 스스로 1회
생성한다(`_exportGenerated` 가드).

- 공통: `ExportMode`, `ExportStart`, `ExportEnd`, `ExportGroups`
- `NgRateForDailyReportPage`: `ExportProgress`, `OnExportComputed(HierReports?, NgRateReport?)`
- `NgRateForWeeklyReportPage`: `ExportSharedHierarchy`, `ExportProgress`
- `BmesCauseMonthlyReportPage`: `ExportSharedHierarchy`
- `BmesFCostPage`: `ExportSharedNgTrend`, `ExportSharedReport`, `ExportProgress`,
  `OnExportComputed(FCostExportSnapshot)`

## 6. 진행률 표시

stage 이름은 `BmesReportHtmlExportService`가 소유한다. 페이지가 상수를 따로 갖고
있으면 어긋나므로 옮겼다.

- `StageNames` = `Daily`, `원인 비중`, `Weekly`, `F-COST`
- `BmesReportPage`는 `new ReportProgressTracker(BmesReportHtmlExportService.StageNames, ...)`로
  tracker를 만든다.
- `GenerateAllTabsAsync(start, end, groups, threshold, tracker, log)` 시그니처로
  tracker를 직접 받는다. 예전의 단일 `IProgress<string>` 인자는 없앴다.
- 이전에는 stage별 progress가 연결되지 않아 `SettleUnfinished`로 4개 바가 한꺼번에
  초록이 됐다. 지금은 단계별로 차오른다.

## 7. 현재 검증 상태

- 컴파일: `dotnet build JinoSupporter.Web/JinoSupporter.Web.csproj` 오류 0, 경고 0
- 서버: 포트 5050에서 정상 기동 확인
- 라우트: `/report/bmes`, `/bmes/f-cost` 모두 살아있음(로그인 리다이렉트 302)
- placeholder: 7개 탭 전부 실제 렌더링 경로로 교체됨. `PlaceholderBody`는
  `BuildCombinedHtml`의 안전망으로만 남아있다.

**실제 리포트 출력은 아직 검증하지 않았다.** BMES 실조회가 필요해서 `Get Report`를
실행하지 않았다.

## 8. 다음 세션의 첫 작업

1. `/report/bmes`에서 짧은 기간으로 `Get Report`를 1회 실행한다.
2. 7개 탭에 placeholder 문구가 남아있지 않은지 확인한다.
3. **Weekly 차트 2개가 실제로 그려지는지 확인한다.** 이번 변경 중 가장 불확실한
   부분이다. 안 그려지면 생성된 `report.html`에서 `data-chart` script가 실제로
   들어갔는지, `initCharts`의 `JSON.parse`가 성공하는지부터 본다.
4. F-COST 4개 탭의 숫자가 서로 정합한지 확인한다. 스냅샷 공유가 잘못되면 여기서
   틀어진다.
5. 전부 펼친 표의 길이가 실사용에 견딜만한지 본다. 부담되면 §9의 대안으로 바꾼다.
6. Chrome 메모리가 실제로 개선됐는지 이전 구조와 비교한다. 이 작업의 원래 목적이다.

## 9. 보류한 대안

행 확장 방식으로 다음 안이 있었으나 "전부 펼침"을 선택했다.

- 모든 행을 `data-parent` 속성과 함께 내보내고 하위 행은 `hidden`으로 두되, 토글
  JS를 wrapper에 추가해 정적 HTML 안에서 접기/펼치기를 재현하는 방식.
- DOM 크기는 전부 펼침과 동일하고 기존 화면과 동작이 같지만 작업량이 크다.
- 표가 너무 길어 불편하다는 판단이 서면 이 안으로 전환한다.

## 10. 주의사항

- export 대상 컴포넌트에 `@rendermode`를 붙이지 않는다.
- `IsPeriodExpanded`가 `ExportMode`를 무조건 우선하므로 `ShowAllPeriodColumns`는
  `fcost` 계열 본표에서 무력화된다. §11 참조.
- `OnAfterRenderAsync`에 의존하는 렌더링을 export 경로에서 기대하지 않는다.
- 새 탭을 추가할 때 데이터를 다시 조회하지 말고 Daily/F-COST의 공유 결과를 받는다.
- `_temp/bmes-report`는 생성 시마다 이전 token을 전부 지운다. 결과를 보존해야 하면
  따로 복사한다.
- 이 세션에서 JinoSupporter 저장소의 미커밋 변경을 commit, revert, clean하지 않았다.

## 11. 2026-07-22 07:04 - 실제 Get Report 출력 검증

§8의 1~4번을 실행했다. 산출물은
`D:\000. MyWorks\002. DB\_temp\bmes-report\53e963ea3c2a4805bc34634b74c0a01f\report.html`
(937KB).

### 통과한 항목

- **placeholder 0건.** `'{key}' 탭 HTML 생성은 준비 중입니다` 문자열이 없다.
  7개 탭 `<template>`이 모두 실제 렌더 결과로 채워졌다.
- **Weekly 차트 payload 정상.** `<script type="application/json" data-chart>` 2개가
  실데이터로 emit됐고 `canvasId`가 `ngrateByGroupCanvas`,
  `ngrateByGroupModelCanvas`로 캔버스 id와 일치한다.
  labels = `["W30","W29","W28","W27","","M7"]` (빈 문자열은 주/월 블록 구분자).
  가장 불확실하다고 봤던 부분인데 문제 없었다.
- **F-COST 스냅샷 공유 정합.** `fcost-weekly`의 모델 5종이 전부 `fcost`의 모델 41종에
  포함된다. 고아 모델이 없으므로 스냅샷 재사용이 어긋나지 않았다.

### 탭별 분량

| 탭 | chars | tables | rows |
|---|---|---|---|
| daily | 202,104 | 15 | 161 |
| weekly | 167,444 | 10 | 264 |
| cause-monthly | 69,312 | 2 | 145 |
| fcost | 231,501 | 2 | 263 |
| fcost-all | 231,501 | 2 | 263 |
| fcost-weekly | 20,643 | 1 | 9 |
| fcost-weekly-all | 30,574 | 1 | 9 |

### 발견한 문제: `fcost`와 `fcost-all`이 바이트 단위로 동일

`cmp`로 확인했다. 두 탭이 완전히 같은 231KB이므로 `F-COST(전체)` 탭은 현재 의미가 없고
파일의 절반(462KB)이 중복이다. 메모리 절감이 목적인 작업에서 역효과다.

원인은 두 경로가 겹친 것이다.

- 본표 컬럼 가시성은 `VisibleFCostColumnIndexes` → `ShouldShowPeriodColumn` →
  `IsPeriodExpanded`를 탄다. `IsPeriodExpanded`가 `ExportMode || ...`라서 export에서는
  항상 전개된다. 즉 `ShowAllPeriodColumns`가 본표에 개입할 여지가 없다.
- `ShowAllPeriodColumns`를 실제로 읽는 곳은 `WeeklyReportColumnIndexes` 분기(1881행)와
  Trend 카드 제목(170행)뿐이다. Trend 카드는 `IsTrendReportRoute && trendRows.Count > 0`
  조건이라 export 경로에서 렌더되지 않아 제목 차이조차 안 나온다.
- `fcost-weekly` / `fcost-weekly-all`은 1881행 분기를 타므로 20,643 vs 30,574로 정상
  구분된다. 이 쌍은 문제 없다.

§4의 "전부 펼침" 원칙은 "접힌 것은 정적 HTML에서 영영 도달 불가"가 근거였는데, 기간
컬럼은 `(전체)` 탭이라는 도달 경로가 이미 있으므로 이 근거가 적용되지 않는다.

### 미확인

- §8의 5번(표 길이 실사용 판단), 6번(Chrome 메모리 실측)은 브라우저에서 봐야 한다.
- 인터랙티브 경로의 `OnAfterRenderAsync`(1947~1948행)가 두 캔버스에 같은
  `BuildGroupTrendChartData()` 결과를 넘긴다. export는 이 동작을 그대로 재현했을 뿐이라
  이번 작업의 회귀는 아니지만, "그룹별"과 "그룹모델별" 차트가 같은 그림이라는 뜻이라
  별도 확인이 필요하다.

## 12. 2026-07-22 - TREND 탭 복구 + 뷰어 툴바 구현

§11의 문제와 사용자 요청을 반영해 코드를 수정했다. **컴파일만 확인했고 실제 출력은
아직 검증하지 않았다.**

### 12.1 사라진 "주요 모델 불량률 TREND" 표 복구

사용자가 이 표가 export에 없다고 지적했다. 조사 결과 **사라진 게 아니라 한 번도
포함된 적이 없었다.**

`BmesFCostPage.razor:166`의 `@if (IsTrendReportRoute && trendRows.Count > 0)` 카드인데,
`IsTrendReportRoute`는 `TrendReportOnly` 파라미터를 그대로 반환한다. 그런데
`TrendReportOnly = true`로 세팅하는 코드가 **웹 프로젝트 전체에 없었다.**

- 서비스의 파라미터 딕셔너리에 항목 자체가 없어 4개 변형 모두 기본값 false
- `/bmes/f-cost`(`BmesFCost.razor:12`)도 `<BmesFCostPage />`를 파라미터 없이 렌더
- `AppMenus.cs`에 trend 메뉴 없음

`BuildMajorTrendRows()`, 카드 마크업, `.fcost-trend-title` CSS는 전부 살아있었고
진입점만 없었다. 표준 F-COST 표와 컬럼 소스가 달라(`WeeklyReportColumnIndexes` =
4주 + 3개월) 스크린샷의 W28~W25 / Jul~May 구성과 일치한다.

조치: `FCostVariants`에 `TrendOnly` 플래그를 추가하고 `("trend", false, false, true)`
변형을 넣었다. `Tabs`에 `("trend", "주요 모델 불량률 TREND")` 추가. **탭이 8개가 됐다.**

이것이 §11에서 `fcost`와 `fcost-all`이 바이트 동일이었던 이유도 설명한다.
`ShowAllPeriodColumns`가 non-weekly 변형에서 차이를 내는 지점(`:170` 제목, `:273`)이
둘 다 `IsTrendReportRoute` 안에 있어 렌더 자체가 안 됐다.

### 12.2 §11의 IsPeriodExpanded 수정은 철회

사용자가 한 번 승인했으나 적용하지 않았다. 뷰어 툴바가 컬럼 개수를 **늘릴** 수 있으려면
HTML에 전체 컬럼이 들어있어야 하는데, 그 수정은 `fcost` 탭에서 컬럼을 잘라낸다.
`IsPeriodExpanded`는 `ExportMode || ...` 그대로 두었다.

`fcost` vs `fcost-all` 중복은 여전히 남아있다. 툴바로 컬럼 수를 조절할 수 있게 된 만큼
`fcost-all` 탭은 이제 불필요하다고 볼 여지가 있으나 판단 보류했다.

### 12.3 뷰어 툴바

`BuildCombinedHtml`에 sticky 툴바를 추가했다. 전부 show/hide + 재합산이며 데이터는
건드리지 않는다. 값은 `localStorage`(`bmesReportViewerSettings`)에 저장되고 탭 전환 시
`show()`가 `applyAll()`을 다시 부른다. `@media print`에서 툴바와 탭 메뉴는 숨긴다.

| 컨트롤 | 구현 |
|---|---|
| 폰트 | `--tb-font` → `#rpt-host`, 셀은 `inherit` |
| 크기 | `--tb-zoom` (50~250%). `zoom`을 쓴 이유는 리포트 마크업의 9px/10px 하드코딩을 비율로 같이 줄이기 위해서다. `font-size` 오버라이드는 그 의도된 차이를 뭉갠다 |
| 굵기 | `--tb-weight`. `.fw-bold`는 `!important`라 총계 행 굵기는 유지된다 |
| Date / Week / Month | 표기 개수. 빈 값 = 전체 |
| Min PPM | 기본값은 Report Setup에서 넘어온 값(사용자 지정 500) |

### 12.4 기간 컬럼 제어 방식 — 75개 루프를 고치지 않은 이유

`VisibleDateCols` 등을 도는 루프가 Daily 26개, Weekly 27개, F-COST 22개로 총 75곳이다.
전부 `data-pk`/`data-pi`를 붙이는 대신, 생성 HTML의 구조적 규칙성을 이용했다.

- 기간 컬럼은 **항상 각 행의 오른쪽 끝**에 `date → sep → week → sep → month` 순
- 따라서 블록 크기 3개만 알면 셀을 오른쪽부터 세어 찾을 수 있다
- 각 탭이 `<script type="application/json" data-periods>{"date":n,"week":n,"month":n}</script>`
  를 1회 emit (`ExportPeriodMeta()`, Daily·Weekly 각 1곳)

오른쪽부터 세는 방식이라 선행 라벨 컬럼 개수가 표마다 달라도 안전하다. `sep-td`/`sep-th`는
세지 않고 건너뛴다. `blk-date`/`blk-week`/`blk-month` colspan 헤더 행은 별도 분기로
colspan만 조정한다.

**한계: `data-periods`가 없는 탭은 컬럼 컨트롤이 동작하지 않는다.** 확인 결과
`blk-*` 마커는 daily 1개 표, weekly 10개 표에만 있고 `fcost`/`cause-monthly`/
`fcost-weekly`에는 없다. F-COST 계열은 `_report.Columns`의 `FCostPeriodKind` 기반이라
마크업 구조가 달라 `ExportPeriodMeta()`를 아직 넣지 않았다. **미완 항목이다.**

### 12.5 Minimum PPM 동적화

- export 시 `DailyReasonPpmThreshold`를 **0으로 강제**해 모든 reason 행을 내보낸다.
  정적 파일에서 잘라낸 행은 복구가 불가능하기 때문이다.
- 각 상세 행에 `data-ppm`(기준일 PPM)을 붙인다 — `ExportReasonPpm()`. export 모드가
  아니면 null을 반환해 Blazor가 속성을 생략하므로 인터랙티브 화면 마크업은 그대로다.
- 섹션 제목은 `data-ppm-label` 템플릿(`{v}` 치환)을 달아 JS가 다시 쓴다.
- **Total 행 재계산이 필요하다.** 원본은 필터를 통과한 상세만 합산하므로(`:1426`),
  임계값이 바뀌면 합계도 바뀌어야 한다. JS가 보이는 상세 행의 셀을 오른쪽부터
  `periodCount`개만 파싱해 다시 합산한다. 남은 상세가 없으면 Total 행도 숨긴다.
- 서비스 시그니처의 `dailyReasonPpmThreshold`는 **필터에서 툴바 초기값으로 의미가 바뀌었다.**

### 12.6 변경 파일

- `Services/BmesReportHtmlExportService.cs` — Tabs, FCostVariants, TrendReportOnly,
  BuildCombinedHtml(툴바 CSS/HTML/JS), daily threshold 0
- `Components/Pages/NgRateForDailyReportPage.razor` — `ExportPeriodMeta()`,
  `ExportReasonPpm()`, `ExportReasonLabelTemplate()`, rank-row `data-ppm`, 섹션 제목
- `Components/Pages/NgRateForWeeklyReportPage.razor` — `ExportPeriodMeta()`

### 12.7 검증 상태

- `dotnet msbuild -t:Compile` 오류 0. (`-t:Build`는 실행 중인 서버가 exe를 잠가
  MSB3027 복사 실패가 나지만 컴파일과 무관하다.)
- **런타임 미검증.** 서버가 구버전으로 떠 있어 재시작 후 `Get Report`를 다시 돌려야 한다.

### 12.8 다음 세션의 첫 작업

1. 서버 재시작(`restart-web.cmd`) 후 `/report/bmes`에서 `Get Report` 재실행.
2. **주요 모델 불량률 TREND 탭이 실제로 렌더되는지** 확인. 스냅샷 재사용 변형이라
   `trendRows.Count > 0`이 안 되면 빈 탭이 나온다.
3. 툴바 7개 컨트롤 동작 확인. 특히:
   - Date/Week/Month를 줄였을 때 헤더 colspan과 본문 셀이 어긋나지 않는지
   - Min PPM 변경 시 Total 행 숫자가 맞는지 (이번 구현에서 가장 불확실한 부분)
4. threshold 0 export로 daily 탭이 얼마나 커졌는지 측정. 202KB 대비 증가폭이
   과하면 하한을 두는 안을 재검토한다.
5. F-COST 계열 탭에 `ExportPeriodMeta()` 대응 추가 (§12.4 미완 항목).

## 13. 2026-07-22 - Report Setup의 minimum PPM 제거 + 툴바 기본값 확정

### 13.1 Report Setup에서 "Daily Reason Table minimum PPM" 입력 제거

§12.5로 export가 threshold 0을 쓰게 되면서 이 입력은 데이터를 거르지 않게 됐고,
툴바 값이 `localStorage`에 저장되므로 두 번째 조회부터는 덮어써졌다. 첫 조회에만
영향을 주는 입력이라 혼란만 남아 제거했다.

- `BmesReportPage.razor` — `<AdditionalSettings>` 블록, `_dailyReasonPpmThreshold`
  필드, 호출 인자 삭제
- `BmesReportHtmlExportService.GenerateAllTabsAsync`에서 `dailyReasonPpmThreshold`
  파라미터 삭제. 대신 `DefaultReasonPpmThreshold = 500d` 상수를 툴바 초기값으로 쓴다.

주의: `NgRateForDailyReportPage`의 `DailyReasonPpmThreshold` **파라미터 자체는 남아있다.**
인터랙티브 Daily 화면(`ShowSetup`)이 자체 입력으로 계속 쓰기 때문이다. 제거한 것은
`/report/bmes` 생성 화면의 입력뿐이다.

### 13.2 툴바 기본값 확정 (사용자 지정)

| 컨트롤 | 기본값 |
|---|---|
| 폰트 | 맑은 고딕 |
| 크기 | 90% |
| 굵기 | 보통(400) |
| Date | 7 일 |
| Week | 4 주 |
| Month | 3 개월 |
| Min PPM | 500 |

마크업의 `value`/`selected` 속성과 JS `DEFAULTS` 객체 양쪽에 두었다. 전자는 JS 실행
전 첫 페인트용, 후자는 `초기화` 버튼용이다. **둘을 따로 고치면 어긋나므로 같이 바꿔야
한다.**

### 13.3 설정 범위는 탭별이 아니라 전역

`localStorage` 항목 하나(`bmesReportViewerSettings`)를 모든 탭이 공유하고, 탭 전환 시
`show()`가 새 본문에 `applyAll()`을 다시 적용한다.

탭마다 기간 블록 구성이 달라도 안전하게 무시된다. 예: Weekly는 date 블록이 0이라 Date
값이 무시되고, Daily는 month 컬럼이 1개뿐이라 Month 3을 넣어도 있는 만큼만 보인다.
`applyPeriods`가 `Math.min(limit, blockSize)`를 쓰기 때문이다.

### 13.4 검증 상태

`dotnet msbuild -t:Compile` 오류 0. **런타임은 여전히 미검증** — §12.8의 항목이 그대로
남아있다.

## 14. 2026-07-22 - TREND 카드를 F-COST 탭 안으로 이동 + 스크린샷 표 불일치 확인

### 14.1 별도 trend 탭을 취소하고 F-COST 탭 상단으로

§12.1에서 `trend` 탭을 새로 만들었으나, 사용자가 "원래 F-COST에 현재 나오는 것 위에
있었다"고 지적했다. 재생성본에서 `trend` 탭 자체는 231KB로 정상 렌더됐음을 확인했고
(`fcost-trend-title` 존재, 실제 모델/불량률/F-COST/비중 데이터 포함) 위치만 틀렸다.

조치:

- `BmesFCostPage`에 `[Parameter] ShowMajorTrendCard` 추가.
  카드 조건을 `IsTrendReportRoute` → `(IsTrendReportRoute || ShowMajorTrendCard)`로 변경.
  카드는 마크업상 이미 일반 F-COST 표(`:563`/`:591`/`:803`)보다 위에 있어 위치 조정 불필요.
- `Tabs`에서 `trend` 항목 제거 → **다시 7개 탭**.
- `FCostVariants`의 4번째 요소를 `TrendOnly` → `TrendCard`로 바꾸고
  `fcost`/`fcost-all` 두 변형에 true. 파라미터도 `TrendReportOnly` → `ShowMajorTrendCard`.

`TrendReportOnly`는 여전히 어디서도 true가 되지 않는다(§12.1). 트렌드 전용 라우트를
살릴 계획이 없다면 정리 대상이다.

부수 효과: 트렌드 카드 제목이 `ShowAllPeriodColumns`에 따라 "FCOST Trend" /
"FCOST Trend (전체)"로 갈리므로 §11의 `fcost` == `fcost-all` 바이트 동일 문제가
부분적으로 해소된다. 본표 컬럼은 여전히 동일하다.

### 14.2 사용자 스크린샷의 표는 이 코드베이스 산출물이 아니다

사용자가 제시한 "■ 주요 모델 불량률 TREND" 표와 현재 트렌드 카드는 **다른 표다.**

| 항목 | 스크린샷 | 현재 코드 |
|---|---|---|
| 1열 헤더 | 제품명 | 모델 |
| 2열 헤더 | 항목 | 지표 |
| 3열 | **26년 목표 불량률** | 없음 |
| 지표명 | F-COST 비중 | 비중 |
| 모델명 | BRS-161016S08ZZ | BRS-161016S08ZZ_E2 |

확인한 것:

- `JinoSupporter.Web`과 `BmesNgRateStandalone`의 트렌드 카드 마크업은 동일하며 둘 다
  목표 불량률 컬럼이 없다(`:175-176`).
- `목표 불량률` 컬럼은 Weekly Report 표(`:446` `weekly-target-th`)에만 있고, 그 표는
  `구분 | 모델 | 목표 불량률 | 달성률 | 기간…`으로 모델당 1행이라 구조가 다르다.
- git 이력 `-S "제품명"`, `-S "trend-target"`, `-S "26년"` 검색 결과 해당 형태 없음.
  `제품명`은 초기 커밋 869b841의 F-Cost 본표 헤더로만 등장한다.

**결론: 스크린샷은 Excel/PPT 등 외부 산출물이거나 사용자가 원하는 목표 디자인이다.**
현재 코드로 그 표를 그대로 재현하려면 트렌드 카드에 목표 불량률 컬럼 추가와 헤더 라벨
변경이 별도로 필요하다. 사용자 확인 대기 중.

### 14.3 재생성본 관측치 (token 7bdf523e…)

threshold 0 export의 실제 비용이 측정됐다.

| 탭 | 이전(§11) | 이번 | 변화 |
|---|---|---|---|
| daily | 202,104 | 470,655 | **+133%** |
| 전체 파일 | 937KB | 1.47MB | +57% |

daily가 2.3배가 됐다. Chrome 메모리가 이 작업의 원래 목적이므로 실측 후 하한 도입을
재검토할 필요가 있다.

### 14.4 검증 상태

`dotnet msbuild -t:Compile` 오류 0. 런타임 미검증 — 서버 재시작 후 재생성 필요.

## 15. 2026-07-22 - 동시 생성 레이스 버그 수정 + 툴바 동작 확인

### 15.1 툴바 컬럼 컨트롤은 정상 동작 확인 (런타임 첫 검증)

사용자 화면에서 확인했다. Summary 헤더의 자체 카운터는 `D 94/94 - W 17/17 - M 4/4`
(= DOM에 전체 컬럼이 들어있음)인데 실제 표시는 Date 7개(07-22~07-15), Week 4개
(W30~W27)였다. 툴바 값 Date 7 / Week 4 / Month 3과 일치한다.

**§12.4의 "오른쪽부터 세기" 방식이 실제로 동작한다.** 헤더 colspan도 어긋나지 않았다.

### 15.2 "Could not find a part of the path ...\report.html" 원인과 수정

사용자 화면에 이 에러가 떴다. 서빙 경로 문제가 아니다 —
`Program.cs:265`의 `/report/bmes/view/{token}`은 `ResolveReportFile`이 null이면
`Results.NotFound()`를 깨끗이 반환한다. **생성 중 쓰기 실패였다.**

원인: `GenerateAllTabsAsync`가 **시작 시점에** `CleanupOldTokens()`로 모든 토큰 폴더를
삭제했다. BMES 스크레이프 때문에 생성이 오래 걸리므로, 진행 중에 `Get Report`를 다시
누르면 두 번째 실행의 cleanup이 첫 번째 실행의 디렉터리를 지우고, 첫 번째가
`File.WriteAllTextAsync`를 할 때 경로가 없어 예외가 난다. 에러의 토큰
`5c7554e7…`는 디스크에 없고 `661074a9…`만 남아있던 정황과 일치한다.

수정 (`BmesReportHtmlExportService.cs`):

- `SemaphoreSlim GenerationLock = new(1, 1)` 추가. `GenerateAllTabsAsync`는 락을 잡는
  래퍼가 되고 본체는 `GenerateAllTabsCoreAsync`로 분리했다.
- `CleanupOldTokens()`를 **쓰기 이후**로 이동하고 `CleanupOldTokens(string keepToken)`
  으로 바꿔 방금 만든 토큰은 건너뛴다.

이제 늦게 누른 요청은 앞 요청이 끝날 때까지 대기하며, 정리는 새 파일이 디스크에
존재한 뒤에만 일어난다.

### 15.3 알아둘 것: localStorage가 기본값보다 우선한다

사용자 화면의 크기가 90이 아니라 100이었다. §13.2의 기본값은 저장된 설정이 없을 때만
적용되고, 한 번이라도 툴바를 만졌으면 `localStorage`의 값이 이긴다. 기본값으로
되돌리려면 `초기화` 버튼을 눌러야 한다. 의도된 동작이다.

### 15.4 검증 상태

`dotnet msbuild -t:Compile` 오류 0. 레이스 수정은 서버 재시작 후 `Get Report` 연타로
확인해야 한다. §14.1의 TREND 카드 F-COST 탭 상단 배치도 아직 미확인이다.

## 16. 2026-07-22 - Daily 탭을 Summary + 모델별 하위 탭으로 분리

### 16.1 목적

Daily 탭 하나가 470KB(§14.3)로 가장 무겁다. Summary 카드와 모델별 블록이 한 페이지에
연속으로 쌓여 있어 모델이 많을수록 DOM이 커진다. 최상위 탭과 같은 방식으로 하위 탭을
만들어 **선택된 섹션 하나만 DOM에 올린다.**

### 16.2 섹션 마커 (컴포넌트)

`NgRateForDailyReportPage`의 최상위 블록 두 종류에 속성을 붙였다.

| 블록 | 위치 | `data-daily-section` | `data-daily-label` |
|---|---|---|---|
| Summary 카드 | `:73` | `summary` | `Summary` |
| 모델별 블록 | `:228` `daily-model-copy-target` | `model.Key` | `model.Material` |

`ExportSectionKey(key)`가 export 모드가 아니면 null을 반환하므로 Blazor가 속성을
생략한다. **인터랙티브 Daily 화면은 기존처럼 연속 스크롤 그대로다.**

모델 블록은 원래부터 `daily-model-copy-target` div로 감싸져 있어 래퍼를 새로 만들 필요가
없었다. Summary는 `@if (!ReasonOnly)` 안의 카드 div(`:73`~`:189`)가 그대로 경계였다.

### 16.3 하위 탭 렌더링 (wrapper JS)

`initSubTabs(scope)`:

- `[data-daily-section]`이 2개 미만이면 아무것도 하지 않는다. **마커가 없는 탭은 기존
  동작 그대로다.**
- 각 섹션을 `template.content.appendChild(s)`로 **라이브 DOM에서 분리**한다.
- 알약 모양 버튼 바(`.rpt-subtabs`)와 `subHost`를 붙이고 첫 섹션을 보여준다.
- 전환 시 `subHost.innerHTML = ''`로 이전 섹션을 버린다 — 최상위 탭과 같은 방식.

`showSub()`는 매번 `applyPeriods(host)` / `applyPpm(host)` / `initCharts(subHost)` /
`sizeTables()`를 다시 부른다. 새로 mount된 섹션에도 툴바 설정이 적용돼야 하기 때문이다.

`data-periods` 메타 스크립트는 섹션 **바깥**(본문 최상단)에 있어 분리 대상이 아니고
`host`에 남는다. 그래서 `applyPeriods(host)`가 계속 찾을 수 있다. **섹션 안으로 옮기면
안 된다.**

`sizeTables()`는 기존 `bmesReportTableSizer` 호출을 함수로 뽑은 것이다.

### 16.4 알아둘 것

- 하위 탭 선택은 저장되지 않는다. 최상위 탭을 바꿨다 돌아오면 항상 Summary로 초기화된다.
- 섹션 바깥에 남는 것은 `<style>`(NgRateReportStyles)과 `data-periods` 스크립트뿐이라
  보이지 않는다. 그래서 하위 탭 바가 사실상 본문 최상단에 온다.
- Weekly 탭에는 마커가 없어 차트 2개가 기존대로 `initCharts(host)` 경로를 탄다.

### 16.5 검증 상태

`dotnet msbuild -t:Compile` 오류 0. **런타임 미검증.** 재생성 후 확인할 것:

1. Daily 탭에 Summary + 모델 수만큼 하위 탭이 뜨는지
2. 하위 탭 전환 시 툴바의 Date/Week/Month와 Min PPM이 계속 먹는지
3. 하위 탭 전환 후 Chrome 메모리가 실제로 내려가는지 (이 작업의 목적)

## 17. 2026-07-22 - BMES LPA 메뉴 신규 추가

### 17.1 범위

좌측 메뉴 BMES 그룹의 Test 4 아래에 `LPA`를 추가하고, BMES `MES073260/SearchList`
응답을 표로 표시한다. 기존 리포트 작업(§11~§16)과는 별개 기능이다.

### 17.2 요청 스펙 (사용자 제공)

- `POST https://bmes.bujeon.com/MES073260/SearchList`
- `Content-Type: application/json`, `X-Requested-With: XMLHttpRequest`
- Payload:
  `{"Condition":{"FACCO":"GN","AUDAT_FR":"19000101","AUDAT_TO":"20501231","LQRNO":"","LQBNO":"","AULOC":"","IMPLV":"","DICNO":"","LASEQ":"","CHKER":"","ZSTAT":"","USEYN":"Y"}}`
- 응답 약 290KB JSON

**사용자가 브라우저 세션 쿠키(`ASP.NET_SessionId`, `__RequestVerificationToken`,
`SingleSignOn`)를 함께 붙여넣었으나 코드/파일에 저장하지 않았다.** 다른 스크레이퍼와
동일하게 `NgRateSettingsService`의 자격증명으로 매번 로그인해 쿠키를 새로 받는다.

### 17.3 추가/변경 파일

| 파일 | 내용 |
|---|---|
| `Services/BmesLpaScrapeService.cs` | 신규. 토큰 획득 → 로그인 → JSON POST → 파싱 |
| `Components/Pages/BmesLpaPage.razor` | 신규. `@page "/bmes/lpa"`, 조회 폼 + 표 |
| `Services/AppMenus.cs` | `BmesLpa = "bmes-lpa"` 상수, `All` 항목, 역할 기본값 4곳 |
| `Components/Layout/NavMenu.razor` | Test 4 바로 아래 NavLink |
| `Program.cs` | `AddScoped<BmesLpaScrapeService>()` |

### 17.4 설계 판단

**응답 스키마를 고정하지 않았다.** `BmesRoutingScrapeService`는 컬럼 배열을 하드코딩하지만
LPA는 응답 컬럼 구성을 확인할 방법이 없었다(실제 호출을 못 해봄). 그래서 반환된 객체의
모든 프로퍼티를 최초 등장 순서대로 컬럼으로 만든다. 서버 스키마가 바뀌면 컬럼이 하나
늘어날 뿐 데이터가 조용히 누락되지 않는다.

`ParseRows`는 `data.contents` → `data`(배열) → `contents` → 루트 배열 순으로 찾는다.
`data.contents`가 다른 BMES 엔드포인트의 형태라 우선순위를 가장 높게 뒀지만 **실제 LPA
응답이 이 중 무엇인지 미확인이다.**

DB 저장은 하지 않는다. 요청이 "받아서 표로 표시"였다.

BaseUrl은 `https`를 썼다. 사용자 제공 요청이 https였고 `NgRateService`/`FCostService`도
https다. (`BmesRoutingScrapeService` 등 일부는 http를 쓴다.)

### 17.5 검증 상태

`dotnet msbuild -t:Compile` 오류 0. **런타임 미검증 — 실제 BMES 호출을 못 했다.**

재시작 후 확인할 것:

1. 좌측 메뉴에 LPA가 보이는지 (권한 기본값에 넣었지만 **기존 사용자의 DB 권한 행에는
   없으므로 Admin 외 역할은 Setting에서 권한을 다시 저장해야 할 수 있다**)
2. `/bmes/lpa`에서 Search가 200을 받는지
3. **표에 컬럼이 뜨는지.** 안 뜨면 `ParseRows`의 배열 위치 추정이 틀린 것이므로 응답
   JSON의 최상위 구조를 확인해야 한다.

## 18. 2026-07-22 - LPA 권한 Admin 한정 + 응답 파서 구조 비의존화

### 18.1 권한을 Admin으로 한정

§17.3에서 Manager/ManagerAi/Leader 등 4개 역할 기본값에 `BmesLpa`를 넣었던 것을 되돌렸다.
이제 `AppMenus.All`에만 있고, Admin은 `All.Select(m => m.Id)`로 전체를 받으므로 **Admin
전용**이 된다. 나중에 다른 역할에 열려면 §17.3의 역할 배열에 `BmesLpa`를 다시 넣으면 된다.

### 18.2 ParseRows를 응답 구조에 의존하지 않게 변경

기존에는 `data.contents` → `data` → `contents` → 루트 배열 순으로 **경로를 추측**했다.
실제 LPA 응답 구조를 확인할 방법이 없어 추측이 틀리면 빈 표가 나오는 구조였다.

이제 문서 전체를 순회해 **객체 배열 중 원소가 가장 많은 것**을 행 집합으로 택한다.
결과 payload에서 행 배열은 다른 메타데이터 배열보다 압도적으로 크므로 이 방식이면
래퍼 모양(`data.contents`든 루트 배열이든)과 무관하게 동작한다.

- `FindLargestObjectArray(element, depth, ref best, ref bestCount)` 재귀
- `MaxScanDepth = 8`로 깊이 제한
- 배열 안에 또 배열이 중첩된 경우(그룹형 결과)도 한 단계 더 내려가 본다

컬럼은 여전히 등장 순서대로 동적 생성이라 스키마를 하드코딩하지 않는다.

### 18.3 검증 상태

`dotnet msbuild -t:Compile` 오류 0. **런타임 미검증** — BMES 실호출을 못 했다.
표가 비면 응답이 객체 배열을 포함하지 않는다는 뜻이므로 그때는 응답 JSON의 실제 형태를
봐야 한다.

## 19. 2026-07-22 - LPA 상세 화면(MES073261) 추가 + 날짜 기본값

### 19.1 응답 구조가 실물로 확인됨

사용자가 MES073261 실응답을 제공했다. `{"result":true,"message":"SUCCESS","data":{"contents":[...]}}`
형태이며 §18.2의 "가장 큰 객체 배열" 파서가 그대로 처리한다. 목록(MES073260)도 같은
래퍼일 것으로 보인다.

상세 레코드 필드: `LQRNO, LORSQ, RESUT, DESCR, LQBNO, LOBVE, LOBSQ, CATGY, SBCAT,
DETAL, TYPRC, IMPLV, LCITM, LCPUR, LCMET, LCEVA, ZIMAG, ZIMAG_TX, ZSORT, USEYN,
ERNAM, ERDAT`

`LCITM`/`LCPUR`/`LCMET`/`LCEVA`는 **베트남어 + 한국어가 `\n`으로 이어진 장문**이다.
그래서 상세는 표가 아니라 라벨/값 그리드로 렌더하고 `white-space: pre-wrap`으로 줄바꿈을
보존한다.

### 19.2 서비스 리팩터링

`FetchAsync`의 본문을 `PostSearchAsync(endpoint, condition, progress)`로 분리했다.
토큰 → 로그인 → JSON POST → 파싱 흐름이 두 엔드포인트에서 동일하고 경로와 Condition
객체만 다르기 때문이다.

- `ListEndpoint   = "/MES073260/SearchList"` — 기존 목록
- `DetailEndpoint = "/MES073261/SearchList"` — 신규 상세
- `LpaDetailQuery` 레코드: `Lqrno`(필수), `Facco`, `Dicno`, `Auloc`, `Implv`, `Lcitm`

### 19.3 페이지 동작

- 목록 행 클릭 → 해당 `LQRNO`의 상세를 조회해 목록 아래 카드로 표시. 선택 행은 하이라이트.
- 상세 조건의 `DICNO`/`AULOC`/`IMPLV`는 **선택한 목록 행에서 읽는다.** 목록 응답에 해당
  컬럼이 없으면 빈 문자열이 가고, BMES는 빈 값을 "전체"로 취급하므로 조회 범위만 넓어진다.
- 상세 항목마다 `LORSQ` 뱃지 + `TYPRC` + `IMPLV`/`RESUT` 뱃지를 헤더로 얹고, 나머지 필드는
  값이 있는 것만 라벨/값 그리드로 나열한다.
- **라벨은 원본 필드명 그대로 쓴다.** 스키마 문서가 없어 한국어 라벨을 붙이면 추측이 된다.
- 새로 Search하면 열려있던 상세는 닫는다(이전 결과의 것이므로).

### 19.4 날짜 기본값

`Audit Date From = 2026-06-08` 고정(`LpaFirstDate` 상수), `Audit Date To = DateTime.Today`.
`To`는 필드 이니셜라이저라 컴포넌트가 만들어질 때마다 그날 날짜가 된다.

### 19.5 검증 상태

`dotnet msbuild -t:Compile` 오류 0. **런타임 미검증.** 확인할 것:

1. 목록 Search 결과가 표에 뜨는지
2. 행 클릭 시 상세가 뜨는지. 안 뜨면 목록 응답에 `DICNO`/`AULOC` 컬럼이 없어 조건이
   과하게 넓거나 좁을 수 있으므로 목록 컬럼명을 먼저 확인한다.
3. 장문 필드 줄바꿈이 살아있는지
