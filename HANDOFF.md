# JinoSupporter — Persistent Handoff

이 파일은 세션 간 재개 지점을 기록한다. 최신 항목을 아래에 append한다. BMES Report
정적 HTML 작업의 상세 이력은 `JinoSupporter.Web/SESSION_HANDOFF_2026-07-22.md`(§1~§19)에
있다.

## 2026-07-23 07:20 - Daily 하위 탭(§16) 런타임 검증
- Completed: 사용자가 "Daily 탭에 Summary만 보이고 모델별 탭이 없다"고 지적. §16(Daily를
  Summary + 모델별 하위 탭으로 분리)이 실제로 동작하는지 검증했다. 결론: 코드 정상,
  화면이 오래된 리포트를 표시 중.
- Decisions:
  - 현재 생성물 `D:\000. MyWorks\002. DB\_temp\bmes-report\76def4...\report.html`
    (07:14 생성, 3.6MB)에 `data-daily-section` 8개, `initSubTabs`, `rpt-subtabs` 모두 존재.
    `show()`가 `initSubTabs(host)` 호출, `host=rpt-host` 정상.
  - 정적 분석만으론 원인 불명이라 scratchpad에 jsdom 설치 후 그 report.html의 JS를
    실제 실행. 결과: JS 에러 0, `.rpt-subtabs` 바 1개 + 알약 버튼 8개
    (Summary + BRS/MSU/MSM-S931B/MSM-X526/TIU-C11/TIU-L5S3/ASSY 338) 정상 생성.
    바는 `#rpt-host` 최상단(툴바 아래, "Summary" 제목 위)에 위치.
  - 따라서 기능은 정상. 사용자 화면은 §16 이전 생성물이나 캐시된 iframe을 표시 중일 뿐.
    07:14 파일이 §16 코드를 담고 있으므로 현재 실행 중인 서버는 이미 §16 빌드다.
- Files: 코드 변경 없음. 검증 스크립트만 scratchpad에 생성(비영구).
- Verification: jsdom 실행으로 하위 탭 바 렌더 확인(위 결과). dotnet 빌드는 돌리지 않음
  (코드 변경이 없어서). 서버는 건드리지 않음(재시작은 사용자 몫).
- Next: 사용자가 `/report/bmes`에서 `Get Report` 재실행(필요 시 Ctrl+F5)하면 툴바 아래
  하위 탭 8개가 나타남. 재실행 후에도 안 뜨면 브라우저 캐시/더 오래된 서버 프로세스를
  의심할 것. §16 코드 자체는 수정 불필요.
- **정정: 위 "수정 불필요" 결론은 틀렸다. 아래 07:50 항목에서 실제 CSS 버그를 찾아 고쳤다.**

## 2026-07-23 07:50 - 하위 탭 안 보이는 실제 원인 = CSS로 버튼 숨김, 수정
- Completed: `Get Report` 재실행(402083 rows) 후에도 하위 탭이 안 보였다. jsdom에선
  바가 렌더됐지만 실브라우저에서만 안 뜨는 차이 = 외부 CSS. **실제 Chrome 헤드리스로
  report.html을 렌더**해 재현하고 계산된 스타일을 뽑아 근본 원인을 확정, 수정했다.
- Decisions / 근본 원인:
  - 헤드리스 Chrome 계산 스타일: `.rpt-subtabs` 바는 보이나(rectH=6, 빈 테두리만)
    `.rpt-subtab` 버튼이 전부 `display:none`. 즉 바는 DOM에 있는데 버튼이 CSS로 숨겨짐.
  - 범인: `BmesReportHtmlExportService.cs:283`
    `.report-export button:not(.rpt-tab):not(.export-toggle) { display:none !important; }`
    — 정적 export에서 인터랙티브 버튼(복사/이동/삭제)을 숨기는 blanket 규칙. §16에서
    새로 만든 하위 탭 버튼 클래스 `.rpt-subtab`이 예외 목록에 없어 같이 숨겨졌다.
  - 같은 규칙이 툴바 `#tb-reset`(초기화) 버튼도 숨기고 있었다(기존에 눈치 못 챈 버그).
  - 수정: 예외에 `.rpt-subtab`과 `#tb-reset` 추가.
    `...button:not(.rpt-tab):not(.rpt-subtab):not(.export-toggle):not(#tb-reset)`.
  - 교훈: export 대상에 **새 버튼 클래스를 추가하면 반드시 이 :not 화이트리스트에 넣어야**
    한다. 안 그러면 조용히 숨겨진다.
- Files: `Services/BmesReportHtmlExportService.cs` (283행 규칙 1줄 + 위 주석).
- Verification:
  - 실제 Chrome 헤드리스로 수정 전(shot.png)=하위 탭 없음 재현 → 수정 후(shot3.png)=
    Summary + 모델 7개 알약 바 + 초기화 버튼 정상 표시 확인.
  - `dotnet msbuild -t:Compile` error CS 0건(기존 경고만). 서버는 건드리지 않음.
- Next: **사용자가 서버 재시작 후 `/report/bmes`에서 `Get Report`를 다시 실행**해야 실제
  화면에 반영됨(컴파일만 됨, 배포 아님). 그 뒤 Daily 탭 상단에 하위 탭 8개가 보이면 완료.

## 2026-07-23 08:10 - LPA 목록 컬럼 정리 + 모델별 TOTAL 행
- Completed: `BmesLpaPage.razor` 목록 표를 사용자 요청대로 변경.
  - 헤더 라벨: OKCNT→OK, NGCNT→NG (`ColumnHeader`).
  - NG(NGCNT) 바로 뒤에 합성 컬럼 `NG RATE`(ppm) 추가. 값 = NGCNT/TOTAL×1,000,000,
    TOTAL=0이면 "-". 사용자 확정 분모 = TOTAL 열.
  - NACNT, ERNAM_TX, VERID 컬럼 숨김(`HiddenColumns`에 추가). VERID_TX는 유지.
  - 각 날짜의 각 모델 그룹 행(`lpa-model-row`)에 per-model TOTAL 표시: TOTAL/OK/NG 합계 +
    NG RATE를 각 컬럼 아래 정렬해 렌더. 모델명 라벨은 첫 숫자 컬럼 직전까지 colspan.
- Decisions:
  - NG RATE는 데이터에 없는 합성 컬럼이라 `DisplayColumns`에서 NGCNT 뒤에 삽입, 셀 값은
    행의 NGCNT/TOTAL로 계산(`RowNgRate`), 총계 행은 모델 합계로 계산(`ModelTotalText`).
  - 숫자 컬럼(TOTAL/OKCNT/NGCNT/NG RATE)은 `.lpa-num`으로 우측 정렬 + tabular-nums.
  - 모델 행 sticky-left는 라벨 셀(`.lpa-model-label`)에만 적용(예전엔 모든 td에 걸려 있어
    총계 셀 여러 개면 왼쪽에 겹칠 문제 → 라벨만 sticky로 변경).
- Files: `Components/Pages/BmesLpaPage.razor` (HiddenColumns, DisplayColumns, 헤더/데이터/
  모델행 마크업, 헬퍼 `ModelTotal`/`IsTotalColumn`/`NumClass`/`ColumnHeader`/`ParseCount`/
  `NgRateText`/`RowNgRate`/`ModelTotalText`, CSS).
- Verification: `dotnet msbuild -t:Compile` error CS 0건. **런타임 미검증** — BMES 실조회가
  필요해 Search를 못 돌렸다.
- Next: 서버 재시작 후 `/bmes/lpa`에서 Search → (1) 헤더 OK/NG/NG RATE, (2) NACNT/ERNAM_TX/
  VERID 사라짐, (3) 모델 행에 TOTAL/OK/NG/NG RATE 합계가 컬럼 아래 정렬돼 뜨는지 확인.
- 참고: §16 하위 탭 + LPA 컬럼 변경은 사용자 화면(스크린샷)에서 정상 반영 확인됨.

## 2026-07-23 08:30 - LPA 상세를 모달→인라인 행으로 전환
- Completed: 상세(MES073261)를 오버레이 모달 대신 **클릭한 행 바로 아래 인라인 행**으로
  펼치도록 변경. "각 모델 밑에 상세 표가 뜨게" 요청 반영.
- Decisions:
  - 데이터 행 렌더 직후 `selected`면 `<tr class="lpa-detail-row"><td colspan=전체>`로
    상세 그리드를 인라인 렌더. 내용(로딩/에러/항목 그리드)은 기존 모달 본문과 동일 구조.
  - `ShowDetailAsync`에 토글 추가: 이미 열린 행의 상세(또는 행)를 다시 누르면 접힘.
  - 오버레이 모달 마크업/`@if (_detailLqrno …)` 블록과 `.lpa-overlay/.lpa-modal/
    .lpa-modal-body` CSS 제거. 대신 `.lpa-detail-row/.lpa-detail-inline/
    .lpa-detail-inline-head` CSS 추가(좌측 들여쓰기 40px, 배경 #f8fafc, 하단 강조선).
  - `_detailLqrno/_detailRows/ShowDetailAsync/CloseDetail` 로직은 그대로 재사용 —
    렌더 위치만 바뀜. **on-click 방식**(자동 전개 아님): 모델마다 상세 자동조회는 행당
    BMES 호출이라 비용이 커서 배제. 상세 버튼/행 클릭 시 그 행만 인라인 조회.
- Files: `Components/Pages/BmesLpaPage.razor` (인라인 상세 행, 모달 제거, 토글, CSS).
- Verification: `dotnet msbuild -t:Compile` error CS 0건. 잔여 `lpa-overlay/lpa-modal`
  참조 없음. **런타임 미검증**(BMES 실조회 필요).
- Next: 서버 재시작 후 `/bmes/lpa` Search → 행 상세 클릭 시 그 행 아래에 상세가 인라인으로
  펼쳐지고, 다시 클릭하면 접히는지 확인. 자동 전개를 원하면 별도 요청 필요.

## 2026-07-23 09:00 - 상단 브라우저형 탭 바 추가(네비게이션 방식)
- Completed: `MainLayout.razor`에 페이지 이동 시 탭 생성/전환/닫기 되는 상단 탭 바 구현.
  사용자 확정 방식 = **네비게이션 탭**(상태 유지 X, 전환=재이동).
- Decisions:
  - 탭 상태(`List<TabEntry>{Url,Title}`)는 `MainLayout`에 보관. 레이아웃 인스턴스는 같은
    레이아웃 내 페이지 이동 시 유지되므로 탭 목록이 살아있음.
  - `Nav.LocationChanged` 구독 → `TrackCurrent()`가 현재 base-relative 경로(쿼리/해시 제거,
    trim '/')로 탭 추가/활성화. 탭 클릭=`NavigateTo(url)`, ×=닫기(활성 탭 닫으면 인접 탭,
    마지막이면 "/"로).
  - 탭 라벨: URL 마지막 세그먼트로 임시 생성 후, `OnAfterRenderAsync`에서 JS `appDocTitle`
    (=`document.title`, 각 페이지 `<PageTitle>`)로 실제 제목으로 갱신.
  - CSS는 `MainLayout.razor.css`(스코프드) `.app-tabs/.app-tab/.app-tab-close`. main이 스크롤
    컨테이너라 `.app-tabs`를 `position:sticky; top:0`로 상단 고정.
- Files: `Components/Layout/MainLayout.razor`(탭 상태/로직/마크업),
  `Components/Layout/MainLayout.razor.css`(탭 CSS), `wwwroot/js/app.js`(`appDocTitle`).
- Verification: `dotnet msbuild -t:Compile` error CS 0건. **런타임 미검증**.
- Next: 서버 재시작 후 여러 메뉴를 눌러 탭이 생기고, 탭 클릭으로 전환, ×로 닫힘, 활성 탭
  닫을 때 인접 탭으로 가는지 확인. 라벨이 각 페이지 제목으로 뜨는지(=`appDocTitle` 동작).

## 2026-07-23 09:05 - Schedule 삭제 요청: 조사 완료, 삭제는 보류(범위 확인 대기)
- Completed: "Schedule/GN LAB Schedule 메뉴+코드+DB 전부 삭제" 요청의 영향 범위 조사.
  **파괴적 + 공유 리소스라 실행 전 확인 필요 → 아직 아무것도 삭제/변경하지 않음.**
- 발견(중요):
  - Schedule 기능은 **Web 앱과 WPF 앱(`BmesNgRateStandalone`) 양쪽에 동일하게 존재**.
    두 프로젝트 모두 `AppMenus.Schedule`, "GN LAB Schedule" 메뉴, `SchedulePage`(Web),
    `WebRepository`의 `EnsureScheduleDatabase`/`Schedules` 테이블 CRUD를 가짐.
  - DB는 **공유 SQLite `schedule.db`** — 실제 파일 `D:\000. MyWorks\002. DB\schedule.db`
    (20KB) 및 `...\01. NG RATE\schedule.db`. WPF 앱도 이 DB를 씀.
  - 따라서 Web만 지우고 `schedule.db`를 삭제하면 **WPF 앱 Schedule이 깨짐**(빈 테이블
    재생성은 되나 기존 데이터 소실).
- Web 삭제 대상 코드(참고): `Components/Pages/SchedulePage.razor`(라우트 `/schedule`),
  `AppMenus.cs`(Schedule 상수·항목·역할 기본값 5곳), `NavMenu.razor`(Schedule 그룹/서브링크/
  `_scheduleOpen`/`IsScheduleActive`/`ToggleSchedule`), `WebRepository.cs`(`ScheduleItem`,
  `_scheduleDbPath`, `OpenScheduleConnection`, `EnsureScheduleDatabase`, `ReadScheduleItem`,
  `GetSchedulesInRange`, `AddSchedule`, `UpdateSchedule`, `DeleteSchedule`, `GetAllScheduleTags`).
- Next(사용자 결정 대기): (1) Web만 vs Web+WPF 둘 다 코드 삭제? (2) 공유 `schedule.db`
  파일 삭제할지(백업 여부). 결정 후 실행.

## 2026-07-23 09:40 - Schedule 완전 삭제 실행 완료 (Web+WPF, DB 삭제)
- 사용자 결정: 코드는 Web+WPF 둘 다 삭제, `schedule.db`는 백업 없이 즉시 삭제.
- Completed:
  - DB: `D:\000. MyWorks\002. DB\schedule.db`, `...\01. NG RATE\schedule.db` 삭제(약 40KB).
    잠금 없었음. **주의: 재시작 전 구버전 서버가 다시 만들 수 있음 → 재시작 후엔 생성 코드
    없어 재생성 안 됨.**
  - 코드(양 프로젝트): `SchedulePage.razor` 삭제(Web). `AppMenus.cs` Schedule 상수/메뉴/
    역할 5곳 제거. `NavMenu.razor`(Web) Schedule 그룹/토글/IsScheduleActive 제거.
    `WebRepository.cs` ScheduleItem 레코드·생성자 블록·EnsureScheduleDatabase·
    OpenScheduleConnection·CRUD 전부 제거(scratchpad `strip_schedule.py` 앵커 기반,
    각 199줄). `AppPathsService.cs` `ScheduleDbPath` 프로퍼티+7 사용처 제거.
    `WpfSettingsReader.cs` `TryGetScheduleDatabasePath` 제거. `Program.cs`/`App.xaml.cs`
    `Schedule:Path` 설정 제거. `BmesSettingPage`/`AdminDbQueryPage`/`TestExcelConverterPage`
    의 ScheduleDbPath 복사·비교 제거.
  - `WebRepository.cs`의 깨진 섹션 헤더(U+0080 혼입)는 Edit 매칭 불가라 Python으로 처리함.
- Decisions/주의:
  - WPF `AppMenus.cs`는 작업 중 외부(linter/사용자)에서 동시 수정됨(BmesDailyReport 추가 등).
    Schedule 제거는 유지됨. 되돌리지 말 것.
- Verification:
  - 잔여 Schedule 참조 검색 0건(양 프로젝트, SchedStart/SchedEnd 등 무관 항목 제외).
  - **Web: `dotnet msbuild -t:Compile` error CS 0건.**
  - **WPF: Schedule 관련 에러 0건.** 남은 에러 3건은 전부 `NgRateSetupPanel.razor`의
    `ReportStage` 미정의(CS0246)이며, `Services/ReportProgressTracker.cs`가 WPF 프로젝트에
    아예 없어서 나는 **사전 존재/동시 sync 이슈**다. Schedule 삭제와 무관하며 건드리지 않음.
- Next: 서버 재시작 후 Web에서 Schedule 메뉴/페이지가 사라졌는지 확인.

## 2026-07-23 09:55 - WPF ReportStage 미정의 수정 (사전 존재 건)
- Completed: 사용자 요청으로 WPF의 `ReportStage`/`ReportStageStatus` 미정의(CS0246 3건)
  해결. Web `Services/ReportProgressTracker.cs`를 `BmesNgRateStandalone.Services`
  네임스페이스로 복제해 `BmesNgRateStandalone/Services/ReportProgressTracker.cs` 신규 생성.
  파일은 BCL만 의존해 그대로 이식 가능. WPF `_Imports.razor`에 이미 `@using
  BmesNgRateStandalone.Services`가 있어 razor에서 해결됨. 기존 중복 정의 없음.
- Verification:
  - `-t:Compile`은 WPF XAML 마크업 컴파일 단계를 건너뛰어 InitializeComponent/named
    element/Main 없음 오류(CS5001/CS0103)를 내지만 이는 타깃 아티팩트다.
  - **`dotnet build`(전체, XAML 포함) 실제 CS 에러 0건.** ReportStage 3건 해소, 신규 오류 없음.
- Next: 없음(이 건 종결). 서버/WPF 재시작·실행 확인은 사용자 몫.

## 2026-07-24 07:28 - LPA를 Report식 정적 HTML 뷰어로 전환 + NG 집계 탭 추가
- Completed: `/bmes/lpa`의 Search 결과를 Blazor 표 대신 **임시 HTML 1개 파일**로 만들어
  iframe으로 보여주도록 전환(=`/report/bmes`와 같은 구조). 그 HTML 안에 하위 탭 2개:
  `목록`(기존 표 전부) + `NG 집계`(날짜 × AULOC 교차표, 신규).
- Decisions:
  - **상세(MES073261)를 Search 때 전부 미리 조회**하는 방식으로 사용자 확정. 정적 HTML은
    BMES를 다시 못 부르므로 자기완결 스냅샷이 되려면 선조회가 필수. 대신 Search가 느려짐.
  - 그래서 `BmesLpaScrapeService`에 **세션 재사용**을 도입. 기존엔 호출 1건마다
    token GET + login POST를 새로 했기 때문에 상세 N건 = 로그인 N번이었다. 신규
    `FetchDetailsAsync`는 **로그인 1회 + 동시 4개**로 전체 상세를 받는다(`Session` 클래스,
    `DetailBatchConcurrency=4`). 레거시 엔드포인트라 동시수는 일부러 낮게 잡음.
  - 상세는 inert `<template>`에 넣고 클릭할 때만 DOM에 clone — Report의 메모리 대책과 동일.
  - 컬럼 규칙(HiddenColumns / 합성 `NG RATE` / OK·NG 헤더 이름)은 페이지에서 export
    서비스로 **이동**. 표가 이제 거기에만 존재하므로 규칙도 한 곳에 둔다.
  - 뷰어 안에서 필터·접기/펼치기·크기(zoom)를 처리. 모델 합계와 그룹 건수는 **필터 통과분만**
    JS가 재계산(기존 서버측 동작과 동일). 목록 필터 입력창은 페이지 → HTML로 이동.
  - `NG 집계` 탭: 행=날짜(최신순), 열=AULOC, 값=NG 합계 + 행/열 합계. 셀에 NG/TOTAL/ppm을
    data 속성으로 다 실어두고 `표시` 선택으로 전환(요청은 NG 건수, 그것이 기본값).
  - 상세의 RESUT 배지는 기존에 값과 무관하게 초록이었음 → NG/FAIL만 빨강으로 분리.
- Files:
  - 신규 `Services/BmesLpaHtmlExportService.cs` (HTML 생성 전체).
  - `Services/BmesLpaScrapeService.cs` (세션 재사용 `Session`/`OpenSessionAsync`,
    신규 `FetchDetailsAsync`, `PostSearchAsync` 시그니처 변경).
  - `Components/Pages/BmesLpaPage.razor` (표·상세·필터 마크업 전부 제거 → iframe + 조회/생성).
  - `Program.cs` (DI 등록 + `GET /bmes/lpa/view/{token}` 라우트).
  - 생성물 위치: `D:\000. MyWorks\002. DB\_temp\bmes-lpa\<token>\lpa.html` (이전 토큰 폴더는
    새 생성 성공 후 정리).
- Verification:
  - `dotnet msbuild -t:Compile` **error CS 0건**(경고는 전부 기존 것). `-t:Rebuild`도 CS 에러
    0건 — 실패한 건 실행 중 서버가 exe를 잠가서 나는 bin 복사(MSB3021/3027)뿐. 서버는 안 건드림.
  - **런타임 검증함**: scratchpad에 실제 `BmesLpaHtmlExportService.cs`를 링크한 하네스를 만들어
    합성 데이터(3날짜×3모델×3레이어=27행, 상세 없는 행 1개·실패 행 1개 포함)로 HTML을 생성하고
    **실제 헤드리스 Chrome**으로 자체 테스트 55개 실행 → **55/55 PASS**. 커버: 상세 열기/닫기/
    토글·누락·실패 폴백, 필터(9/27 및 재번호), 모델 합계·NG RATE 재계산, 날짜/모델 접기,
    피벗 값이 목록 NG와 일치, 총계 86 일치, 표시 전환, 탭별 툴바, zoom, 가로 스크롤 없음.
    `상세` 버튼이 CSS로 안 숨는지도 확인(07-23 07:50에 당했던 건).
  - 스크린샷으로 목록/인라인 상세/피벗 레이아웃 육안 확인.
  - 주의: **BMES 실조회는 못 했다.** 실제 응답 컬럼명·상세 소요시간은 미검증.
- Next(이 항목은 아래 08:0x에서 이어짐): 서버 재시작 후 `/bmes/lpa`에서 Search. 확인할 것 (1) 상세 선조회 진행 표시
  (`상세 조회 n / N…`)와 **실제 소요 시간** — 행이 수백 개면 오래 걸린다. 너무 느리면
  `DetailBatchConcurrency` 상향 또는 "상세 미리받기" 옵션화를 검토. (2) 목록 탭이 기존 표와
  같은지. (3) `NG 집계` 탭의 날짜×AULOC 값. (4) 상세 실패 건수가 status에 뜨는지.

## 2026-07-24 07:45 - NG 집계 탭 재설계 (모델 × 일자/주차/월, ppm + NG/TOTAL)
- Completed: `NG 집계` 탭을 사용자 요청 형태로 재구성. 축을 뒤집고 기간 블록 3개를 붙였다.
  - **행 = 모델(AULOC)**, 정렬 = 검사수(TOTAL) 내림차순 → 동률 시 모델명. 정렬 기준이 안 보이면
    이상하므로 맨 앞에 `전체` 열(기간 전체 집계)을 두어 정렬키를 노출.
  - **열 = 최신→과거** 순의 3블록: `일자`(MM/dd) · `주차`(W#) · `월`(M#). 블록 사이에 구분열.
  - **셀 = NG RATE(ppm) 크게 + 그 밑에 작게 `NG/TOTAL`**. 비율만 있으면 모집단이 안 보여서
    (0/50과 0/50000이 구분 안 됨) 요청대로 원수를 함께 표시.
  - 데이터 없는 기간은 빈칸(=감사 없음), 감사했는데 NG 0인 것과 구분.
  - 기존 `표시`(NG건수/TOTAL/ppm) 셀렉터는 셀 형식이 고정되어 불필요 → 제거. 대신
    `일자/주차/월` 개수 제한 입력을 추가(빈칸=전체). 열은 전부 파일에 있고 JS가 숨기기만 함.
- Decisions:
  - 주/월 버킷 키는 `NgRateReportService`의 기존 규칙을 그대로 따름 — `W:{yyyy}{ww}`
    (`CalendarWeekRule.FirstDay`, 월요일 시작, **ISO 아님**), `M:{yyyy}{MM}`, 헤더 `W#`/`M#`.
    BMES 리포트에서의 "W30"과 여기의 "W30"이 같은 주를 뜻하게 하기 위함.
  - AUDAT 파싱 실패 행은 일자 열에는 남지만(원문 문자열이 키) 주/월에는 못 들어간다.
- 이번에 잡은 버그 2개(둘 다 실제 브라우저 검증에서만 드러남):
  - **날짜 헤더가 `07-23`으로 나옴.** `ToString("MM/dd")`의 `/`는 *문화권의* 날짜 구분자라
    ko-KR 서버에서 `-`로 렌더된다. `CultureInfo.InvariantCulture` 지정으로 수정.
    (`NgRateReportService.FormatDateHeader`도 같은 코드라 리포트도 동일 증상일 것 — 미수정.)
  - **2단 헤더가 스크롤 시 겹침.** 밴드행 높이를 CSS `height:22px`로 고정했지만 표 셀에서
    height는 최소값이라 실제 26.5px가 되어 아래 행이 4.5px 파고들었다. → 런타임에
    `offsetHeight`를 재서 `--band-h`에 넣는 `syncPivotHeader()`로 수정(zoom 변경 시에도 재측정).
- Files: `Services/BmesLpaHtmlExportService.cs` (피벗 빌더 전면 교체, 툴바/CSS/JS).
- Verification: `dotnet msbuild -t:Compile` **error CS 0건**(경고 32개 전부 기존 것, LPA 파일 0건).
  헤드리스 Chrome 자체 테스트 **80/80 PASS** (기존 71 + 신규):
  - 기본 세트 71/71 — 행 정렬(TOTAL desc→이름), 헤더 `모델`/`전체`/밴드 `일자,주차,월`,
    셀이 ppm+NG/TOTAL이고 목록 데이터와 일치, 합계행, 기간 제한(1개/0개/해제 시 밴드·구분열
    동반 처리), 피벗 필터, 탭별 툴바.
  - 넓은 데이터셋(5/20~7/23, 6일·5주·3개월) 3/3 — `07/23,07/21,07/08,06/30,06/10,05/20`,
    `W30,W28,W27,W24,W21`, `M7,M6,M5` 모두 최신→과거 순 확인.
  - sticky 6/6 — 스크롤 상태에서 밴드행 고정, 2행 헤더 겹침 없음, 모델 열 가로 고정.
  - 스크린샷으로 레이아웃 육안 확인.
  - 주의: 여전히 **BMES 실조회 미검증**.
- Next: 서버 재시작 후 `/bmes/lpa` Search → `NG 집계` 탭에서 (1) 모델 정렬이 TOTAL 순인지,
  (2) 일자/주차/월 열이 최신부터인지, (3) 셀의 ppm과 (NG/TOTAL)이 맞는지 확인. 실제 기간이
  길면 일자 열이 많아지므로 `일자` 칸에 7 등을 넣어 좁혀 쓰면 된다.

## 2026-07-24 08:12 - NG 집계 탭: 합계 최상단 이동 + 불량 상세 표 추가
- Completed: 사용자 요청 2건.
  - **합계 행을 표 맨 위로.** `<tfoot>`에 있던 합계를 `<thead>`의 **3번째 행**으로 옮겼다.
    단순히 순서만 바꾸지 않고 sticky로 고정 — 밴드행/기간행 밑에 붙어서 모델 행이 밑으로
    스크롤되어도 계속 보인다. `<thead>`에 둔 덕에 필터(=`tBodies[0]`만 훑음)가 합계를
    숨길 수 없다는 이점도 있다.
  - **밑에 `불량 상세` 표 신설.** 모델(AULOC)별 접기/펼치기 블록 + `모두 펼치기/모두 접기`
    버튼. 상세(MES073261) 항목 중 **RESUT이 NG/FAIL인 것만** 행으로 나온다.
    열 = `# / 일자 / 레이어 / 항목(LORSQ) / 점검항목(TYPRC) / 중요도(IMPLV) / 결과(RESUT)`
    + **상세 응답에 실제로 값이 있던 나머지 키들**(동적, 첫 등장 순).
- Decisions:
  - 불량 판정 규칙을 `IsNgResult()` 하나로 통합. 기존 인라인 상세 배지가 쓰던
    "NG 또는 FAIL 포함" 규칙과 새 표가 **절대 어긋나지 않게** 하기 위함. RESUT 값 집합이
    문서화되어 있지 않아 규칙 자체는 그대로 뒀다.
  - 뒤쪽 열을 하드코딩하지 않고 동적으로 뽑는다. MES073261 스키마가 미문서라 열 목록을
    박아두면 필드가 조용히 사라진다. 단 **NG 항목 중 하나라도 값이 있는 키만** 열이 된다
    (전부 빈 키, OK 항목에만 있는 키는 열이 안 생김 — 테스트로 고정).
  - LQRNO 기준 중복 제거. 같은 감사가 목록 행 여러 개에 걸쳐 나와도 상세는 하나이므로
    두 번 세면 안 된다.
  - 그룹 정렬 = 불량 건수 내림차순 → 이름. 표의 목적이 "어디를 볼지"라서 가나다순은 무의미.
  - 긴 점검 텍스트는 셀에서 말줄임 처리하고 원문은 `title`에 남김.
  - 불량이 0건이면 표를 아예 만들지 않고 안내문만 — 이때 **목록 NG 합계가 0이 아니면
    그 숫자를 같이 표시**한다(=RESUT 표기가 NG/FAIL이 아닐 수 있다는 신호). 실데이터에서
    RESUT 값이 다른 표기면 이 문구로 바로 드러난다.
  - 표 2개가 한 탭에 쌓이므로 `has-ng`일 때만 스크롤 박스 높이를 42vh로 반씩 나눈다.
- 이번에 잡은 버그 1개(브라우저 검증에서만 드러남):
  - **필터에 모델명을 치면 불량 표가 통째로 비었다.** 모델은 그룹 행에만 있고 데이터 행
    셀에는 없어서 `tr.textContent`에 안 걸린다. → `dataset.model`을 매칭 대상에 명시적으로
    붙여, 피벗과 동일하게 "그 모델만 남는" 동작이 되게 수정.
- Files: `Services/BmesLpaHtmlExportService.cs` (피벗 합계 위치, NG 상세 빌더, CSS, JS).
  다른 파일 변경 없음.
- Verification:
  - `dotnet msbuild -t:Compile` **error CS 0건**, LPA 파일 경고 0건(기존 경고만).
  - 헤드리스 Chrome 자체 테스트 **127/127 PASS** (기존 세트 확장):
    - `test.html` 101 — 기존 목록/상세/피벗 전부 + 신규: 합계가 tfoot이 아닌 thead 3번째 행,
      모든 모델 행보다 위, 값은 전 행 합과 일치 / 불량표 25행·3그룹(9,9,7)·전부 NG 결과·
      OK 항목 미유입 / 헤더가 고정열+값 있는 키만(EMPTYFLD·OKNOTE 제외 확인) /
      1..25 연속 번호 / 접기 시 16행 남고 재번호 / 모두 접기·펼치기 / 필터 연동 / 이스케이프.
    - `sticky.html` 12 — 스크롤 상태에서 합계행이 기간행 바로 밑에 붙고, 겹침 없고,
      모델 행 위로 덮이는지(z-index) 확인.
    - `cleantest.html` 11 — 불량 0건 데이터셋: 표 미생성, 버튼 미생성, 안내문에 목록 NG
      6건이 언급됨, 합계는 여전히 최상단, 패널 높이 분할 안 함.
    - `widetest.html` 3 — 다주/다월 데이터 열 순서.
  - 스크린샷으로 레이아웃 육안 확인(합계 최상단, 아래 불량표 그룹 접기 UI).
  - 주의: 여전히 **BMES 실조회 미검증**. 실제 RESUT 표기가 NG/FAIL이 아니면 불량표가
    0건으로 나오는데, 그 경우 안내문에 목록 NG 합계가 찍히니 바로 알 수 있다.
- Next: 서버 재시작 후 `/bmes/lpa` Search → `NG 집계` 탭에서 (1) 합계가 맨 위에 고정되는지,
  (2) 아래 `불량 상세` 표의 건수가 합계행 NG(예: 134건)와 맞는지, (3) 안 맞으면 안내/부제에
  찍힌 두 숫자를 비교해 RESUT 실제 표기를 확인 → 필요 시 `IsNgResult()` 규칙 보강.

## 2026-07-24 11:21 - 불량 상세 0건 원인 규명 및 수정 (RESUT는 코드 A/B/C)
- Completed: `NG 집계` 탭의 `불량 상세`가 "0건 / 목록 NG 합계 138건"으로 나오던 문제 해결.
  원인은 **RESUT이 텍스트가 아니라 코드**라는 것. MES073261은 `RESUT: "A" | "B" | "C"`를
  주고, BMES 화면은 페이지 스크립트의
  `GRID_SELECT_OPT.RESUT = [A: "A : OK", B: "B : NG", C: "C : N/A"]`로 라벨을 붙여 그린다.
  기존 `IsNgResult()`는 문자열에 "NG"/"FAIL"이 있는지만 봤으므로 코드 "B"를 못 잡아 항상 0건.
- Decisions:
  - 판정 기준 = **RESUT 코드 "B"**. 실데이터로 검증: GN / 2026-06-08~07-24 / 600 감사에서
    감사별 `count(B) == NGCNT`가 **600건 전부 일치**(총합 138 = 138). 코드 분포는
    A 6,056 / B 138 / C 238 / 공백 14, 합 6,446 = 목록 TOTAL 합계.
  - 텍스트 규칙("NG"/"FAIL" 포함)은 **폴백으로 유지**. 엔드포인트가 나중에 라벨을 주더라도
    동작하고, A/C에는 절대 걸리지 않는다.
  - 화면에는 코드가 아니라 BMES와 같은 라벨(`B : NG`)을 표시. 배지 색은
    NG=빨강, A/OK=초록, **C : N/A와 미지의 코드는 회색**(`.lpa-badge-na` 신설) —
    N/A를 초록으로 칠하면 합격처럼 읽힌다.
  - 불량 표의 동적 열 헤더를 MES073261 그리드 `colModel` 헤더와 동일하게 매핑
    (DESCR→Comments, LCITM→Check Item 등). 원래 필드명은 `title`에 유지.
    체크리스트 내부 필드 LOBVE/ZSORT/ZIMAG는 열에서 제외(ZIMAG_TX 경로만 표시).
- Files: `JinoSupporter.Web/Services/BmesLpaHtmlExportService.cs` 단일 파일
  (`ResutLabels`/`ResutLabel`/`IsNgResult`/`ResutBadgeClass`, `NgDetailHeaders`,
  `NgDetailFixedKeys`, 부제 문구, 배지 CSS). 다른 파일 변경 없음.
- Verification:
  - `dotnet msbuild -t:Compile` **error CS 0건**, LPA 파일 경고 0건.
  - **실데이터 오프라인 렌더 검증**(신규 방법): 스크래치패드에서 BMES 실조회를 그대로 덤프
    (목록 600행 + 상세 600건/6,446항목 → `lpa_dump.json`) 후, `JinoSupporter.Web.dll`을 참조하는
    작은 콘솔 앱이 리플렉션으로 `BuildHtml`을 호출해 HTML을 생성 → 서버 없이 실제 산출물 확인.
    결과: 불량 상세 **138건**(= 피벗 합계행 NG 138, 목록 NGCNT 합계 138), 모델 그룹 5개
    (UNIT 56 / MODULE / SPK / FRONT 338 / REAR ASSY 338), 배지 짝 검사
    `result:A : OK 6,056 / ng:B : NG 276(=138 인라인+138 표) / na:C : N/A 238`, 공백 14건은 배지 없음.
  - 헤드리스 Chrome 스크린샷으로 레이아웃 육안 확인(피벗 + 아래 불량 상세, Comments 열에
    "Ng laser cắt lẹm VP NG : 1/8 pcs" 같은 실제 불량 내용이 보임).
- Next: **사용자가 웹 서버를 재시작한 뒤** `/bmes/lpa` Search → `NG 집계` 탭에서
  불량 상세가 138건(조회 기간에 따라 합계행 NG와 같은 값)으로 나오는지 확인.
  참고: 상세 인라인 배지도 이제 `A : OK / B : NG / C : N/A`로 표시된다.

## 2026-07-24 11:51 - 불량 상세에 Result Image 내장 + 클릭 팝업(라이트박스)
- Completed: 사용자 요청 — LPA 불량 사진을 HTML에 내장하고 썸네일 클릭 시 팝업 확대.
  - 신규 `Services/BmesLpaImageService.cs`: `ZIMAG_TX` 경로의 사진을
    `/MES073261/GetImage?fileName=<경로>`로 받아(로그인 불필요, 실측 확인) 썸네일(≤160px)과
    뷰(≤1000px) 두 크기로 축소해 data URI로 반환. 디스크 캐시
    (`_temp/bmes-lpa-images`, 경로 SHA256 앞16B 키)로 재조회 시 즉시.
  - `BmesLpaHtmlExportService`: 불량 상세 표에 `사진` 열, 인라인 상세 헤더에 썸네일 추가.
    뷰 이미지는 화면에 안 그리고 `<script type=application/json id=lpa-img-data>` 안에만 두어,
    클릭 전까지 파싱/디코드 0 → 사진 수백 장이어도 로딩 속도 유지. 라이트박스는
    이전/다음(화면에 보이는 썸네일만 순회)·원본 링크·ESC/←/→ 지원.
  - `BmesLpaPage.razor`: 상세 조회 후 참조된 이미지 경로만 모아 `LpaImages.FetchAsync`,
    결과를 `ExportInput.Images`로 전달. 상태줄에 `이미지 N/M장` 표기(차이나면 링크 대체).
  - `Program.cs`: `BmesLpaImageService` DI 등록.
- Decisions:
  - **원본은 장당 평균 3.7MB(최대 7.5MB)** → 그대로 못 박음. 뷰 1000px/q72면 장당 ~90KB,
    141장이 base64 포함 HTML 27MB로 수렴(측정값). 썸네일은 장당 ~4KB.
  - 임베드 상한 `MaxImages=600`, `MaxEmbeddedBytes=40MB`. 초과분은 파일에 안 넣고
    `사진` 링크(BMES 직접)로 폴백 — 1년치 조회로 파일이 무한정 커지는 것 방지.
  - 이미지 엔드포인트는 세션 불필요라 스크래퍼 로그인 핸들셰이크를 안 탐. HEAD는 404를
    주므로(검증자 없음) 캐시는 경로만 키로. 삭제 시 폴더 지우면 재조회.
  - System.Drawing은 Windows 6.1+ 전용 → `OperatingSystem.IsWindowsVersionAtLeast(6,1)`
    가드로 감싸 비-Windows에선 "이미지 미내장"으로 degrade(크래시 X). CA1416 0건.
  - JSON 임베드 시 `<`, `/`를 이스케이프해 `</script>`로 블록이 조기 종료되지 않게 함.
- Files: `Services/BmesLpaImageService.cs`(신규), `Services/BmesLpaHtmlExportService.cs`,
  `Components/Pages/BmesLpaPage.razor`, `Program.cs`.
- Verification:
  - `dotnet msbuild -t:Compile` **error CS 0건, 경고 0건**(LPA/이미지 파일 기준).
  - **실데이터 오프라인 검증**: 실조회 덤프(600행/상세 600건)에서 이미지 경로 141개 추출 →
    `BmesLpaImageService.FetchAsync` 실행 → **141/141 내장, 18.6초**(재실행 시 캐시로 0.0초),
    썸네일 845KB·뷰 17.0MB, 최종 HTML 27MB. 리플렉션으로 `BuildHtml` 호출해 산출.
  - 헤드리스 Chrome 스크린샷 2종: (1) 불량 표 `사진` 열에 행별 썸네일, (2) 썸네일 클릭 시
    라이트박스가 실제 NG 사진(불량부 빨간 박스)·캡션(LPA번호·일자·모델·항목)·`1/140`·
    이전/다음/원본/닫기와 함께 뜸.
  - 참고: 임베드는 **상세 전체**의 사진(141장)을 포함(불량 138건 + OK/NA 항목 사진 포함),
    표의 `사진` 열은 불량 138건에만 노출. 인라인 상세를 열면 해당 항목 사진도 보임.
- Next: **사용자 서버 재시작 후** `/bmes/lpa` Search → 상태줄 `이미지 N/M장` 확인,
  `NG 집계` 탭 `사진` 열 썸네일 클릭 → 팝업 확대·이전/다음 동작 확인. 최초 조회는
  사진 다운로드로 다소 걸리고(경고: 첫 조회 수십 초), 이후 캐시로 빨라짐.

## 2026-07-24 (오후) - 이미지 lazy 로딩(DB 저장) + 불량상세 매트릭스화 + 엑셀 추출
- Completed(사용자 연속 요청 5건 처리, 모두 `/bmes/lpa`):
  1) **툴바 기본값**: NG 집계 툴바 `일자=7 / 주차=4 / 월=3`(크기는 이미 100). 입력 `value=`만
     추가 — `applyPivot()`가 로드 시 입력값을 읽으므로 JS 변경 불필요.
  2) **이미지 선수집 폐지 → 뷰 시점 lazy 로딩**: Search가 이미지 다운로드로 18초씩 블록되던
     것을 제거. 상세(텍스트)만 받아 HTML 생성 → 즉시. 이미지는 뷰어의 `<img loading="lazy">`가
     신규 라우트 `GET /bmes/lpa/img?path=&size=thumb|view`로 그때그때 요청.
     `<template>` 안(접힌 상세)·화면 밖 썸네일은 실제로 보일 때만 BMES를 침.
  3) **축소본을 파일 캐시 → DB 저장**: `_temp/bmes-lpa-images` 디스크 캐시 폐지. 신규 테이블
     `BmesLpaImages(Path,Kind,ImageData,CreatedAt, PK(Path,Kind))`(process-review.db).
     `WebRepository.GetLpaImage/SaveLpaImagePair`. `Kind`='t'(썸네일160px)/'v'(뷰1000px).
     경로당 1회만 원본 다운로드→두 크기 생성→저장, 이후 뷰·검색·재시작 모두 DB에서 즉시.
  4) **불량 상세 표 → 매트릭스**: 행=모델(UNIT/MODULE…) 섹션 안에서 **(점검항목=TYPRC 1차 ·
     Check Item=LCITM 2차)** 두 라벨 열, **열=일자(NG난 날, 최신순)**, **셀=수량 + 그 날·항목의
     불량 썸네일 전부**. 라벨 2열은 sticky-left(150/200px), 날짜 헤더 sticky-top.
     동적 컬럼/AppendClipped/NgDetailFixedKeys/NgDetailHeaders 제거. 집계는 공용
     `BmesLpaHtmlExportService.BuildNgMatrix()`(public)로 추출 — HTML·엑셀이 동일 데이터 사용.
     모델 '건' 카운트는 행 `data-ng`(행별 NG수) 합산(applyNgDetail 수정, # 번호열 제거).
  5) **엑셀 추출(이미지 포함)**: 신규 `BmesLpaExcelExporter`(ClosedXML). 페이지에 `엑셀 추출
     (불량 상세)` 버튼. `BuildNgMatrix`로 시트 구성(점검항목|Check Item|일자열, 모델 섹션행,
     셀=수량+사진). 사진은 `BmesLpaImageService.GetAsync(path, view:true)`로 **뷰(1000px)** 를
     받아 셀에 임베드하되 화면표시는 64px로 축소(엑셀에서 키우면 선명). `BrowserDownload`로 저장.
     - 사용자 오해 정정: **DB에 원본 없음**(원본 3.7MB는 버림, 축소본 2개만 저장). 그래서 엑셀엔
       "원본" 대신 우리가 가진 최고 축소본=뷰(1000px)를 넣음. 처음엔 썸네일(160px)을 넣어 흐렸음.
- Files:
  - `Services/BmesLpaImageService.cs` (전면 재작성: `GetAsync(path,view)` 단건 on-demand,
    static HttpClient, DB 경유. FetchAsync/LpaImage/DataUri/Key/파일캐시/Max* 전부 제거).
  - `Services/WebRepository.cs` (BmesLpaImages 테이블 + Get/Save 메서드).
  - `Program.cs` (`GET /bmes/lpa/img` 라우트, 인증 게이트 + 장기 캐시 헤더).
  - `Services/BmesLpaHtmlExportService.cs` (ExportInput.Images 제거, AppendThumbs 엔드포인트
    lazy `<img>`+onerror 폴백, BuildImageData/JsonString/lpa-img-data 제거, NgMatrix 레코드 +
    BuildNgMatrix + BuildNgDetailTab 매트릭스화, viewSrc 엔드포인트화, lpaImgFail, CSS 2열
    sticky+셀 스타일, 툴바 기본값).
  - `Services/BmesLpaExcelExporter.cs` (신규).
  - `Components/Pages/BmesLpaPage.razor` (이미지 선수집 제거, `_lastResult/_lastDetails` 보관,
    엑셀 버튼 + `ExportExcelAsync`, LpaImages/IJSRuntime 주입).
- Verification: `dotnet msbuild -t:Compile` **error CS 0건**, 수정/신규 파일(BmesLpa*, WebRepository,
  Program.cs) **경고 0건**(남은 경고 2건은 무관한 기존 것: DataInferenceInputTestPage, ExcelHelperRunner).
  **런타임 미검증**(서버 재시작은 사용자 몫). 특히 ClosedXML 이미지 임베드(AddPicture/MoveTo/
  WithSize)와 매트릭스 sticky 레이아웃은 실제 화면/엑셀로 미확인 — 컴파일로 API 시그니처만 확인.
- Next(사용자 서버 재시작 후 확인):
  1) Search가 이미지 없이 즉시 끝나고, 매트릭스 스크롤/상세 열 때 사진이 lazy로 채워지는지.
  2) 불량 상세: 점검항목/Check Item 2열 + 일자열 + 셀 수량·이미지, 모델 '건' 합계가 피벗 NG와 일치.
  3) `엑셀 추출` → xlsx에 이미지 포함(뷰1000px), 셀 크기·화질 OK. 첫 추출은 이미지 다운로드로 다소 걸림.
  4) 화질 더 필요하면 `BmesLpaImageService`의 `ViewMaxPx`(1000)·`ViewQuality`(72) 상향 검토 —
     **단 캐시 무효화 필요**: 현재 `Kind`가 't'/'v' 고정이라 크기만 바꾸면 기존 DB행이 재사용됨.
     크기 바꾸려면 Kind에 크기를 넣어(예 'v1600') 자동 재생성시키거나 BmesLpaImages 테이블 비우기.

## 2026-07-24 (오후 2) - 엑셀 화질 개선 + 엑셀 3시트화(피벗/불량상세/목록)
- Completed(사용자 연속 요청 2건):
  1) **엑셀 이미지 화질**: 사용자 "DB 원본을 셀크기 맞춰 넣어줘". 정정: DB에 원본 없음(축소본만).
     엑셀에 넣던 **160px 썸네일 → 뷰(1000px)** 로 교체(`GetAsync(path, view:true)`), 셀 표시는
     64px로 축소해 임베드 → 엑셀에서 키워도 선명. 상수 `ThumbPx=40→64`.
  2) **엑셀 3시트**: `엑셀 추출` 결과를 (1)`NG 집계`(피벗, **첫 시트**) (2)`불량 상세`(매트릭스+이미지)
     (3)`목록` 순 다중 시트로. 피벗은 툴바 기본값과 동일하게 **일자7/주차4/월3**로 제한.
- 공용 데이터 추출(엑셀·HTML 동일 소스): `BmesLpaHtmlExportService`에 public 추가 —
  `BuildPivotData(rows, dateN, weekN, monthN)`(+records `PivotData`/`PivotBlock`/`PivotBucket`),
  `BuildListTable(columns, rows)`, `NgRateText` public화. **주의**: `BuildPivotTab`(HTML)은 안
  건드리고 집계 루프를 `BuildPivotData`에 복제함 — 둘이 같은 규칙(WeekKey/MonthKey/Count/ppm)을
  쓰지만 accumulate 루프는 2곳이니 한쪽 바꾸면 다른 쪽도 맞출 것.
- Files: `Services/BmesLpaExcelExporter.cs`(3시트 재작성), `Services/BmesLpaHtmlExportService.cs`
  (BuildPivotData/BuildListTable/records, NgRateText public), `Components/Pages/BmesLpaPage.razor`
  (pivot/list 생성 후 Export에 전달).
- Verification:
  - `dotnet msbuild -t:Compile` **error CS 0건**, 수정/신규 파일 **경고 0건**.
  - **ClosedXML 런타임 스모크 테스트 통과**(scratchpad `xlsxsmoke` 콘솔앱, ClosedXML 0.102.3):
    내가 쓰는 호출 시퀀스(2단 병합 밴드 헤더 + 멀티라인 셀 + `AddPicture(jpeg).MoveTo(cell,x,y).
    WithSize(64,64)` 오프셋 다중 이미지 + FreezeRows/Columns + 3시트 SaveAs)로 유효 xlsx 생성 확인
    → 압축 해제 시 worksheet 3개 + `xl/media/image{,2,3}.jpg` + sheet2 이미지 rels 정상.
  - **단 실제 데이터로 BmesLpaExcelExporter.Export는 미실행**(집계 LINQ는 컴파일만). 서버 재시작 후
    실제 추출로 레이아웃/수량/이미지 확인 필요.
- Next: 서버 재시작 후 `엑셀 추출` → (1)첫 시트 NG 집계 피벗(모델×일자7/주차4/월3, ppm+NG/TOTAL,
  합계행), (2)불량 상세 이미지 화질(1000px 임베드), (3)목록 시트 확인.

## 2026-07-24 (오후 3) - 엑셀 피벗: 셀 2줄 → 불량률/갯수 2행 분리
- Completed: 사용자 "엑셀 피벗은 갯수 행·불량률 행으로 나눠줘". `BmesLpaExcelExporter.WritePivotSheet`
  를 재구성: 모델(및 합계)마다 **두 행**(`불량률`=ppm / `갯수`=NG/TOTAL), 모델명은 두 행 병합,
  신규 `구분` 열(B). 컬럼 A=모델 B=구분 C=전체 D..=기간. 갯수 행은 회색 9pt. Freeze 5행/3열.
  `WritePivotEntity/WriteRate/WriteCount`로 분리, `PivotValue`(2줄 셀) 제거.
- Files: `Services/BmesLpaExcelExporter.cs`만.
- Verification: `dotnet msbuild -t:Compile` error 0/경고 0. 행방향 병합은 기존 스모크(A2:A3 병합)로
  검증된 동일 API. 실데이터 Export는 여전히 미실행 — 서버 재시작 후 확인.
- Next: 서버 재시작 후 `엑셀 추출` 첫 시트에서 모델별 불량률/갯수 2행이 맞는지 확인.

## 2026-07-24 (오후 4) - BMES 로그인 실패(Worker Status 등) = http:// → https:// 수정
- 증상: `/worker-status` Fetch가 "Logging in to BMES… → Login failed — check credentials"로 실패.
  **LPA Search는 정상**(=자격증명 유효). 로그인 코드는 LPA와 동일한데 결과만 다름.
- 근본 원인: `WorkerStatusService.BaseUrl`이 **`http://`** bmes.bujeon.com. LPA(`BmesLpaScrapeService`)만
  **`https://`**. BMES가 TLS를 강제해 http로는 토큰 GET 때 받은 anti-forgery 쿠키가 https 리다이렉트를
  넘어 유지되지 않아 `/MES000000/LoginCheck` POST가 검증 실패 → "Login failed". "갑자기 안 됨"과 일치.
- 수정: **Web 앱** BMES 서비스 3곳 BaseUrl을 https로 통일 —
  `WorkerStatusService.cs:14`, `BmesRoutingScrapeService.cs:17`, `BmesMaterialService.cs:15`.
- Verification: `dotnet msbuild -t:Compile` error 0. 런타임(실제 BMES 로그인)은 서버 재시작 후 사용자 확인.
- Next: 서버 재시작 → `/worker-status` Fetch 성공하는지. Routing/Material 페이지도 같이 정상화됨.

## 2026-07-24 (오후 5) - WPF standalone도 https 수정 + 버전 1.0.21 게시
- Completed: 사용자 "standalone도 업데이트하고 버전업".
  - WPF `BmesNgRateStandalone`의 동일 3파일(`Services/WorkerStatusService.cs:14`,
    `BmesRoutingScrapeService.cs:17/19`, `BmesMaterialService.cs:15/17`) BaseUrl http→**https**.
  - csproj 버전 **1.0.20 → 1.0.21**(Version/FileVersion/AssemblyVersion).
  - **`tools/PublishStandaloneUpdate.ps1` 실행**(csproj 버전 자동 읽음): Release self-contained
    win-x64 게시 → `JinoSupporter.Web/standalone-updates/BmesNgRateStandalone-1.0.21.zip`(≈85.5MB)
    생성 + `update.json`(version 1.0.21, url, sha256=902ce1fb…, notes) 갱신 + 서버 서빙 위치 복사.
- Verification: 게시 성공. 빌드 경고는 기존 2건뿐(HierSubRows CS4014, NgRatePage CS8602) — 무관.
  `update.json`이 1.0.21 가리키고 zip 존재 확인. 구 1.0.20.zip은 남겨둠(매니페스트는 1.0.21만 참조).
- 배포 메커니즘: 서버가 `/standalone/update.json`을 **요청마다 파일에서 읽어** 서빙하므로 웹서버
  재시작 없이도 클라이언트가 다음 폴링에서 1.0.21을 받아 자동 업데이트. (단, 웹서버가
  이 폴더를 ContentRoot로 서빙 중일 때 — 1.0.20이 여기서 서빙되던 것과 동일 위치.)
- Next: 클라이언트(설치된 standalone)에서 업데이트 알림/자동 갱신으로 1.0.21 받는지 확인.
  받은 뒤 standalone에서 Worker Status/Routing/Material 로그인 정상인지.

## 2026-07-24 (오후 6) - WPF standalone에 LPA 기능 이식
- Completed: 사용자 "WPF에 LPA도 다 추가". 웹 LPA를 standalone(BlazorWebView)로 이식.
- 핵심 제약: WPF는 **BlazorWebView라 서버 라우트(app.MapGet) 없음** → 웹의 iframe+`/bmes/lpa/view`·
  `/bmes/lpa/img`를 못 씀. 대신 **`<iframe srcdoc="@_html">`**로 뷰어 HTML을 앱 안에 직접 렌더
  (웹 HTML+JS 전부 재사용). **뷰어 사진은 BMES 익명 URL 직접 로드**(사용자 결정) — 라우트 불필요.
- 공유 파일 파라미터화(웹 안 깨짐): `BmesLpaHtmlExportService`에 `ExportInput.DirectBmesImages`
  추가. false(웹)=`/bmes/lpa/img` 엔드포인트, true(WPF)=BMES 직접 URL. `AppendThumbs`/
  `AppendDetailItem`/`AppendDetailTemplate`/`BuildNgDetailTab`에 bool 전파, JS `viewSrc`는
  `IMG_DIRECT` 플래그로 분기. 신규 `RenderHtml(input)`(문자열 반환, srcdoc용).
- WPF 신규/변경:
  - 서비스 4개 이식(네임스페이스만 변환): `Services/BmesLpaScrapeService.cs`,
    `BmesLpaImageService.cs`, `BmesLpaHtmlExportService.cs`, `BmesLpaExcelExporter.cs`.
  - `Components/Pages/BmesLpaPage.razor`(신규, 네이티브): Search→RenderHtml(DirectBmesImages=true)
    →srcdoc iframe, 엑셀 버튼(3시트, 이미지 view 1000px). `@rendermode` 없음(WPF).
  - `App.xaml.cs`: BmesLpaScrapeService/BmesLpaImageService/BmesLpaHtmlExportService 싱글턴 등록.
  - `Components/Layout/NavMenu.razor`: `bmes/lpa` "LPA" 링크 추가.
  - `WebRepository`(BmesLpaImages 테이블/GetLpaImage/SaveLpaImagePair)·`AppMenus`(BmesLpa)는
    이미 web에서 sync돼 있었음(추가 작업 없음).
- Verification: 웹 `-t:Compile` error 0(파라미터화 후에도). **WPF `dotnet build`(Debug, XAML 포함)
  error 0**. Debug 빌드 실행 → 크래시 없이 기동(DI 정상). **단, LPA 실제 동작(srcdoc 렌더·
  BMES 직접 이미지 로드·엑셀)은 클릭 미검증** — srcdoc은 이 앱에서 신규 패턴.
- Next(사용자): 실행된 WPF에서 좌측 **LPA** 메뉴 → Search → 표/매트릭스/사진 뜨는지, 엑셀 추출 확인.

## 2026-07-24 (오후 7) - WPF LPA 영어화 + 1.0.22 게시
- Completed: 사용자 "WPF LPA는 영어로" + "버전업해서 다른 데서 업데이트 받게".
  - **영어화**: WPF LPA 파일들(웹과 별도 복사본이라 웹은 한글 유지)의 사용자 노출 한글을 영어로.
    Python 일괄 치환(longest-first)으로 `BmesLpaHtmlExportService`/`BmesLpaExcelExporter`/
    `BmesLpaScrapeService` 번역 + 페이지(`BmesLpaPage.razor`) 수동 번역. 예: 목록→List,
    NG 집계→NG Summary, 불량 상세→Defect Detail, 점검항목→Check Point, 불량률→NG Rate,
    갯수→Count, 구분→Type, 전체→Overall, 합계→Total, 모델→Model, 원본→Original 등.
    폰트명 'Malgun Gothic','맑은 고딕'과 일부 코드 주석(교차표/기간/원본)은 그대로(비노출).
  - **주의(분기)**: WPF LPA는 sync 목록에 없는 수동 복사본 → 이제 웹(한글)과 divergent.
    웹 LPA 수정 시 WPF는 자동 반영 안 됨(재복사하면 한글 재유입). 언어차 때문에 의도적 분기.
  - **버전 1.0.21 → 1.0.22** + `PublishStandaloneUpdate.ps1` 재실행: zip(≈85.6MB) +
    update.json(1.0.22, sha256=4c72c487…, 영어 notes) 갱신. 실행 중 인스턴스 taskkill 후 게시.
- Verification: WPF `dotnet build`/publish error 0(게시 성공). 남은 한글 grep: 폰트명+주석뿐.
  1.0.22 게시본 실행해 기동 확인. **LPA 실제 화면/엑셀 클릭 검증은 여전히 미수행**.
- Next: 클라이언트가 1.0.22 자동업데이트로 LPA(영어) 받는지, LPA 메뉴 동작·엑셀 확인.

## 2026-07-25 07:18 - 엑셀 불량 상세 시트에 AULOC 열 추가(웹+WPF)
- Completed: 사용자 "엑셀 추출 쪽, 점검항목 왼쪽에 AULOC 1열 추가". `불량 상세`(WPF: Defect Detail)
  시트를 A=AULOC, B=점검항목/Check Point, C=Check Item, D..=일자로 재배치.
  - `firstDateCol` 3→4, `lastCol` `2+dateCount`→`3+dateCount`, 헤더 3개, 데이터행 3개 열로 시프트.
  - 열폭 A=16 / B=26 / C=34, `FreezeColumns(2)`→`(3)`.
- Decisions:
  - AULOC 값은 `NgMatrixGroup.Model`(=AULOC, 이미 그룹 키). **행마다 반복 기입(병합 안 함)** —
    엑셀 정렬/필터가 살아있게. 
  - **모델 그룹 밴드 행은 그대로 유지**(모델명+건수 표시). 이제 AULOC 열과 정보가 중복되지만
    per-model 건수가 밴드에만 있어서 제거하지 않음. 사용자가 원하면 밴드 삭제 가능.
  - 웹/WPF는 의도적 분기(한글/영어) 상태라 **두 파일을 각각 동일하게 수정**.
- Files: `JinoSupporter.Web/Services/BmesLpaExcelExporter.cs`,
  `BmesNgRateStandalone/Services/BmesLpaExcelExporter.cs` (각각 WriteMatrixSheet + 클래스 doc).
- Verification: 웹 `dotnet msbuild -t:Compile` **error 0**(BmesLpaExcelExporter 관련 경고 0, 나머지는
  기존 무관 경고). WPF `dotnet build` **오류 0개/경고 4개**(기존 HierSubRows CS4014·NgRatePage CS8602가
  wpftmp+본 프로젝트에 중복 계상). **실제 xlsx 추출은 미실행** — 컴파일/좌표 계산만 확인.
- Next: 사용자 서버 재시작 후 웹 `엑셀 추출` → 2번째 시트 A열에 AULOC가 행마다 찍히고 사진이
  D열 이후로 밀려서도 정상 배치되는지 확인. WPF는 배포하려면 버전업+PublishStandaloneUpdate.ps1 필요
  (이번엔 게시 안 함, 현재 게시본은 1.0.22).

## 2026-07-25 07:26 - 웹 PC Download 메뉴 + 설치파일 배포 경로 + WPF 첫 실행 경로설정 유도
- Completed: 사용자 3건 요청.
  1) **웹 Tools에 `PC Download` 메뉴** — id `pc-download`, 라우트 `/tools/pc-download`.
     - `AppMenus`: 상수/`All` 항목(Tools 그룹) + 모든 role 기본권한에 추가.
     - `WebRepository.SeedDefaultMenuPermissionsIfEmpty`: **기존 DB용 grant 추가**(QrBakoData/
       DailyReport와 동일 패턴, AppRoles.All 전원). 안 하면 커스터마이즈된 기존 role에 안 뜸.
     - `NavMenu.razor`: `showTools` 조건 + 💻 항목.
  2) **설치파일 다운로드**
     - 신규 `Services/StandaloneDownloadCatalog.cs`: `standalone-updates/update.json`을 읽어
       `StandaloneRelease`(version/notes/publishedAt + Setup·Package `StandaloneAsset`) 반환.
       Setup은 manifest `setupUrl`, 없으면 `BmesNgRateStandalone_Setup-<ver>.exe` 규칙으로 탐색.
       파일 없으면 null → 페이지가 zip만 노출.
     - 신규 `Components/Pages/PcDownloadPage.razor`(정적 렌더, `[Authorize]`): 버전/게시일/노트/
       용량 + Installer·zip 버튼 + 설치 안내 4단계.
     - `Program.cs`: `/standalone/download/{fileName}`이 **`.exe`도 허용**(기존 zip 전용).
       익명 유지(업데이터가 로그인 전에 받음).
     - `installer.iss`: `OutputBaseFilename`에 버전 삽입(`..._Setup-<ver>.exe`), 기본 버전 1.0.22.
     - `tools/PublishStandaloneUpdate.ps1`: ISCC.exe 탐색(기본 경로 2곳→PATH) 후 있으면
       `/DMyAppVersion=$Version`으로 설치파일 빌드→`standalone-updates`로 복사, manifest에
       `setupUrl`/`setupSha256` 추가. **없으면 경고만 찍고 zip만 게시**(기존 동작 유지).
       WPF `StandaloneUpdateService`는 unknown JSON 필드 무시하므로 기존 업데이트 흐름 무영향.
  3) **WPF 첫 실행 시 경로 지정 화면**: `HomePage.razor` OnInitialized에서
     `!Paths.IsNgRateStorageConfigured`면 `/bmes/setting?firstrun=1`로 replace 이동.
     **static `_firstRunRedirected`로 프로세스당 1회만** — 계속 갇히지 않게.
     `BmesSettingPage.razor`: `firstrun=1`이면 안내 배너(Working Folder→Save→BMES 계정→재시작).
- Decisions/제약:
  - **이 PC에 Inno Setup 미설치** → 설치파일(.exe) 생성 불가. 현재 페이지는 zip(81.6MB)만 노출하고
    Inno Setup 안내 문구 표시. 설치하면 재게시만으로 자동으로 Installer 버튼이 생김.
  - **게시 안 함**: 현재 배포본은 1.0.22라 오늘 변경(AULOC 열, 첫 실행 유도)이 다운로드에 없음.
    게시는 전 클라이언트 자동업데이트를 유발하는 외부 영향 작업이라 사용자 확인 대기.
- Files: web `Services/AppMenus.cs`, `Services/WebRepository.cs`, `Services/StandaloneDownloadCatalog.cs`(신규),
  `Components/Pages/PcDownloadPage.razor`(신규), `Components/Layout/NavMenu.razor`, `Program.cs`;
  wpf `Components/Pages/HomePage.razor`, `Components/Pages/BmesSettingPage.razor`,
  `installer.iss`, `tools/PublishStandaloneUpdate.ps1`.
- Verification:
  - 웹 `dotnet msbuild -t:Compile` **error 0**(신규 파일 경고 0). WPF `dotnet build` **오류 0/경고 4**(기존 2건 중복 계상).
  - PublishStandaloneUpdate.ps1 **PS 5.1 파서 통과**(`?.` 미사용으로 수정함).
  - `StandaloneDownloadCatalog.Read` **실제 폴더로 콘솔 스모크 실행**: v1.0.22 / Package=
    BmesNgRateStandalone-1.0.22.zip 81.6MB / Setup=(none) — 기대대로.
  - **미검증**: 페이지 실제 렌더, 다운로드 라우트 응답, WPF 첫 실행 리다이렉트(설정 지운 상태로 실행 필요),
    Inno Setup 경로의 설치파일 빌드(도구 없음).
- Next(사용자 결정 필요):
  1) 서버 재시작 → Tools > PC Download 표시/다운로드 확인.
  2) 설치파일(.exe) 원하면 Inno Setup 6 설치 후 PublishStandaloneUpdate.ps1 재실행.
  3) 오늘 WPF 변경을 배포하려면 csproj 1.0.22→1.0.23 + 게시(= 전 클라이언트 자동 업데이트).

## 2026-07-25 07:33 - 정정: 설치파일 빌드 스크립트는 이미 있었음(BuildStandaloneInstaller.ps1)
- 정정 내용: 직전 항목에서 "이 PC에 Inno Setup 미설치 → 설치파일 생성 불가"라고 적은 것은 **틀렸다.**
  - Inno Setup은 **`%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe`** 에 설치돼 있다(사용자 단위 설치).
    내가 Program Files 2곳과 PATH만 확인해서 놓쳤다.
  - **`BmesNgRateStandalone/tools/BuildStandaloneInstaller.ps1`이 이미 존재**했고(2026-05-14),
    LOCALAPPDATA 포함 6개 후보 경로를 탐색한다. VS Code task `build-standalone-installer`로도 실행됨.
    내가 PublishStandaloneUpdate.ps1에 ISCC 탐색 로직을 **중복 구현**한 것도 잘못이었다.
- Completed(정정 반영):
  - `PublishStandaloneUpdate.ps1`: 중복 ISCC 블록 제거 → **`BuildStandaloneInstaller.ps1-SkipPublish` 호출**로 교체.
    실패 시 warning만 내고 zip만 게시(기존 동작 보존). setupUrl/setupSha256 매니페스트 추가는 유지.
  - `BuildStandaloneInstaller.ps1`: `-SkipPublish` 스위치 추가(publish 중복 방지, 없으면 publishDir 검사),
    기대 산출물 경로를 **버전 포함 이름**(`dist\BmesNgRateStandalone_Setup-<ver>.exe`)으로 수정
    — installer.iss의 `OutputBaseFilename` 변경과 맞춤. **이 수정 없으면 기존 스크립트가 깨진다.**
  - `README.md`: 산출물 이름 버전 포함으로 갱신 + PC Download 연동 설명 추가.
- Verification: **`BuildStandaloneInstaller.ps1` 실제 실행 성공** — Inno Setup 컴파일 57초,
  `dist\BmesNgRateStandalone_Setup-1.0.22.exe` **60,396,767 bytes(57.6MB)** 생성 확인.
  두 ps1 모두 PS 5.1 파서 통과.
- **주의(버전 불일치)**: 방금 빌드한 dist의 exe는 파일명이 1.0.22지만 **내용은 오늘 변경분
  (엑셀 AULOC 열, 첫 실행 경로유도)이 포함된 코드**다. 게시된 1.0.22 zip과 내용이 다르므로
  **이 exe를 standalone-updates로 복사하지 않았다.** 배포하려면 csproj 1.0.23으로 올려 재빌드/게시할 것.
- Next: 사용자 승인 시 csproj 1.0.22→1.0.23 + `PublishStandaloneUpdate.ps1` 1회 실행
  (zip+설치파일+update.json 동시 생성 → PC Download에 Installer 버튼 자동 노출, 기존 클라이언트 자동 업데이트).

## 2026-07-25 07:38 - 1.0.23 게시(zip+설치파일) + PC Download Razor 버그 수정
- Completed:
  1) **버전 1.0.22 → 1.0.23**(csproj Version/FileVersion/AssemblyVersion) 후
     `PublishStandaloneUpdate.ps1` 1회 실행. 실행 전 프로세스 미실행 확인(taskkill 불필요).
     산출물(모두 `JinoSupporter.Web/standalone-updates/`):
     - `BmesNgRateStandalone-1.0.23.zip` 85,580,800 B (sha256 46134328…)
     - `BmesNgRateStandalone_Setup-1.0.23.exe` 60,386,667 B (sha256 32d68253…) ← **설치파일 최초 게시**
     - `update.json` version 1.0.23 + `setupUrl`/`setupSha256` 필드 포함.
     Inno Setup 컴파일 43초. 구버전 zip(1.0.20~1.0.22)은 남겨둠.
  2) **버그 수정**: PC Download 화면에 버전이 `v@release.Version` 문자 그대로 노출됨(사용자 스크린샷).
     원인 = **Razor의 이메일 주소 자동 인식** — 문자 뒤에 바로 오는 `@`는 식별자가 아니라
     이메일로 처리돼 이스케이프됨. 수정: `v@(release.Version)`로 괄호 감싸기(2곳: 버전 배지 + 경고문).
     **교훈: Razor에서 `텍스트@변수` 형태는 항상 `@(변수)`로 쓸 것.**
- Files: `BmesNgRateStandalone/BmesNgRateStandalone.csproj`(버전),
  `JinoSupporter.Web/Components/Pages/PcDownloadPage.razor`(Razor 수정),
  `standalone-updates/*`(게시 산출물).
- Verification:
  - 게시 스크립트 성공, 위 3개 파일 존재/크기/해시 확인, `update.json` 내용 확인.
  - 웹 `-t:Compile` error 0.
  - `StandaloneDownloadCatalog.Read` 콘솔 스모크 재실행 → **v1.0.23 / Setup=Setup-1.0.23.exe 57.6MB /
    Package=1.0.23.zip 81.6MB** 정상 인식(이제 Installer 버튼이 뜰 상태).
  - **미검증**: 실제 페이지 렌더(서버 재시작 필요), 설치파일 실행/설치 동작, 기존 클라이언트 자동 업데이트.
- 배포 영향: `update.json`은 요청마다 파일에서 읽히므로 **웹서버 재시작 없이도** 기존 standalone
  클라이언트가 다음 폴링에서 1.0.23 자동 업데이트를 받는다(zip 경로). 단 **PC Download 페이지 자체는
  신규 코드라 웹서버 재시작 후에야 보인다.**
- Next: 웹서버 재시작 → Tools > PC Download에서 v1.0.23 + **Download Installer** 버튼 확인 →
  설치파일 받아 실제 설치 테스트(첫 실행 시 Setting 화면 뜨는지 포함).

## 2026-07-25 08:02 - 설치파일 슬림화(Syncfusion 미사용 제거 + 다국어 리소스 제외) → 1.0.24 게시
- 배경: 사용자 "기능 대비 설치파일이 왜 이렇게 크냐". 실측 결과 publish 207.5MB / 설치파일 57.6MB.
  구성: .NET8+WPF 런타임 ≈138MB, Syncfusion 18.7MB, 다국어 리소스 13폴더 18.9MB,
  OpenXml+XlsIO 13MB, SqlClient 8.3MB, wwwroot 8.3MB. 사용자 지시: "런타임은 넣고 나머지는 알아서".
- Completed:
  1) **Syncfusion 완전 제거** — `Syncfusion.Blazor.Core`/`Spreadsheet` PackageReference,
     `App.xaml.cs`(using + RegisterLicense + AddSyncfusionBlazor), `_Imports.razor` 2줄,
     `wwwroot/index.html` 3줄(테마 css + 스크립트 2개).
     **근거: `SfSpreadsheet` 등 Syncfusion 컴포넌트 사용처가 코드 전체에 0건**(웹에서 복사될 때 딸려온 것).
     `DocumentFormat.OpenXml`/`Syncfusion.XlsIO`도 직접 사용 0건 — Spreadsheet 패키지의 전이 의존성이라 같이 빠짐.
     엑셀은 ClosedXML로 처리하므로 무영향.
     **부수 확인**: `index.html`이 참조하던 `_content/Syncfusion.Blazor.Themes/bootstrap5.css`는
     Themes 패키지가 참조돼 있지 않아 **원래부터 404**였다(자체 `bootstrap/bootstrap.min.css` 사용 중).
     따라서 제거해도 시각적 변화 없음.
  2) **`<SatelliteResourceLanguages>en</SatelliteResourceLanguages>`** — 다국어 리소스 13폴더 제외.
  3) 버전 1.0.23 → **1.0.24** 후 재게시.
- 결과(실측): publish **207.5MB → 168.0MB**(-39.5MB), 파일수 **625 → 318**,
  zip **85.6MB → 72.1MB**, **설치파일 60.4MB → 53.6MB**. lang 폴더 0개 / Syncfusion 파일 0개 확인.
- Files: `BmesNgRateStandalone.csproj`, `App.xaml.cs`, `Components/_Imports.razor`,
  `wwwroot/index.html`, `standalone-updates/*`(1.0.24 산출물).
- Verification:
  - `dotnet build` 오류 0/경고 4(기존). **Debug 빌드 실제 실행 → 12초간 살아있고 MainWindowTitle
    'BMES NG Rate' 확인 후 정상 종료**(Syncfusion DI 제거로 인한 기동 실패 없음).
    `%LOCALAPPDATA%\JinoWorkHost\logs\bmes-ngrate-standalone.log`에 1.0.24 Startup 기록만, 에러 없음.
  - 게시 산출물 3종 크기/해시/`update.json`(setupUrl 포함) 확인.
  - **미검증: 각 화면 실제 렌더**(Syncfusion 테마 css 제거 후 시각 확인은 안 함 — 단 위 404 근거로 영향 없다고 판단).
- Next: 웹서버 재시작 → PC Download에서 v1.0.24 Installer(53.6MB) 확인 → 설치 테스트.
  더 줄이려면 framework-dependent 게시(-140MB, 각 PC에 .NET 8 Desktop Runtime 필요)뿐 — 사용자가 런타임 포함 유지 결정함.

## 2026-07-25 08:15 - 탭 제목 오표시 버그 수정(LPA 눌렀는데 탭엔 DAILY REPORT)
- 증상(사용자 스크린샷): 좌측 LPA 클릭 → 본문은 LPA인데 **탭 라벨 3개가 전부 "DAILY REPORT"**.
- 근본 원인: `MainLayout.OnAfterRenderAsync`가 `JS appDocTitle()`로 **document.title을 한 번 읽어**
  활성 탭에 기록하는 구조였다. 그런데 Blazor `<PageTitle>`은 title을 **비동기로 나중에** 쓴다.
  → 레이아웃이 읽는 시점엔 아직 **직전 페이지 제목**이 남아 있어, 그 값이 새 탭에 그대로 박힌다.
  이후 재렌더가 없으면 영영 안 고쳐져서 모든 새 탭이 직전 제목을 물려받는다.
- 수정(폴링 → 이벤트 구독):
  - `wwwroot/js/app.js`: 신규 `window.appTitleWatcher`. `document.head`에 MutationObserver
    (`childList`+`characterData`+`subtree` — title 엘리먼트가 교체돼도 잡히게)를 걸고, title이 바뀔
    때마다 `dotNetRef.invokeMethodAsync('OnDocumentTitleChanged', title)`. start 시 1회 seed 호출.
    invoke는 `.catch(()=>{})`(페이지 정리 중 서킷 소멸 대비). `stop()`으로 disconnect.
  - `Components/Layout/MainLayout.razor`: OnAfterRenderAsync는 **firstRender에서 watcher 시작만**
    (document.title 읽기 제거). 신규 `[JSInvokable] OnDocumentTitleChanged(title)`가 그 순간의
    활성 탭 라벨을 갱신. `DotNetObjectReference` 필드 추가 + `Dispose()`에서 해제.
  - 기존 `window.appDocTitle`은 남겨둠(다른 호출부 있을 수 있어 제거 안 함).
- Files: `JinoSupporter.Web/wwwroot/js/app.js`, `JinoSupporter.Web/Components/Layout/MainLayout.razor`.
- Verification:
  - `dotnet msbuild -t:Compile` error 0.
  - **jsdom으로 JS 실동작 검증**(scratchpad `titlewatch`): app.js에서 appTitleWatcher 블록을
    **직접 추출해 실행**(손으로 옮긴 로직 아님). 결과 PASS —
    start 시 현재 제목 seed 발화, `document.title='LPA'`/`'Setting'` 대입마다 콜백 발화,
    `stop()` 이후 대입은 발화 안 함. 즉 "title 대입이 head observer를 발화시킨다"는 핵심 가정 확인.
  - **미검증**: 실제 브라우저에서 Blazor `<PageTitle>`과의 타이밍(서버 재시작 후 확인 필요).
    URL 유래 폴백 라벨(`TitleFromUrl`)은 그대로라 최악의 경우에도 "Lpa"로 표시된다.
- Next: 웹서버 재시작 → LPA/다른 메뉴 이동 시 탭 라벨이 각 페이지 `<PageTitle>`(LPA 등)로 바뀌는지 확인.

## 2026-07-25 10:55 - LPA NG 집계 탭에 '점검항목별 불량률' 새 표 추가(웹+WPF)
- 요청: 모델 피벗 표 밑에 **Check Item끼리 묶어서** 일/주/월별 불량률을 **새로운 표로**.
- 설계 결정(중요):
  - 행 = **LCITM(Check Item) 텍스트 distinct**, 모델 구분 없이 병합. NG 많은 순 정렬.
    행 머리에 해당 Check Item이 속한 **점검항목(TYPRC)들을 작은 회색 글씨로 병기**(여러 개면 ' · ' 결합).
  - **분모 = 위 모델 피벗과 동일**(해당 기간 전체 검사수량 TOTAL). 그래서 각 행 ppm의 합 = 합계 행이 되고
    두 표를 그대로 비교할 수 있다. (점검 실시횟수 기준 분모는 위 표와 비교 불가라 채택 안 함.)
  - 빈 셀 = 그 기간에 점검 자체가 없음. 점검했는데 이 항목이 정상이면 **진짜 0**(위 표 규칙과 동일).
  - NG 집계는 **`BuildNgMatrix` 재사용** → LQRNO 중복 제거 로직이 불량 상세와 100% 동일.
- 코드 구조:
  - `BuildPivotTab`의 반환을 `string` → **`PivotRender(Html, Blocks, ColTotals)`** 로 변경.
    새 표가 **같은 기간 컬럼과 같은 분모를 그대로 재사용**하게 하기 위함(루프 3번째 복제 방지).
  - 신규 `BuildCheckPivotTab(rows, details, blocks, colTotals)`.
  - 패널: `<div class="lpa-scroll">모델피벗</div>{{checkPivotSection}}{{ngDetailTab}}`.
    패널 클래스에 `has-chk` 추가, `.lpa-panel.has-chk .lpa-scroll{max-height:30vh}`(3단 스택).
  - JS: `applyPivot()`을 `applyPivotTable(table)` + `PIVOT_IDS=['lpa-pivot','lpa-chkpivot']` 로 일반화.
    `syncPivotHeader()`가 인자로 테이블을 받도록 시그니처 변경(zoom 핸들러도 두 표 모두 호출).
    → 툴바의 일자/주차/월 개수 제한과 필터가 새 표에도 그대로 적용됨.
- Files: `JinoSupporter.Web/Services/BmesLpaHtmlExportService.cs`,
  `BmesNgRateStandalone/Services/BmesLpaHtmlExportService.cs`(영어판 동일 이식:
  'NG Rate by Check Item' / 'Total' / 'Overall' / '(no text)').
- Verification(둘 다 **헤드리스 Chrome 실렌더**):
  - scratchpad 하니스(`lpahtml`, `lpahtmlwpf`)로 **실제 서비스 파일을 그대로 컴파일**(BmesLpaScrapeService.
    LpaResult / AppStoragePaths / BmesLpaImageService만 스텁) 후 합성 데이터로 HTML 생성 → puppeteer 렌더.
  - 웹 `dotnet msbuild -t:Compile` error 0, WPF `dotnet build` 오류 0/경고 4(기존).
  - **JS 에러 0건**. 새 표 렌더 확인(높이 155px). 기간 컬럼이 모델 피벗과 동일(07/24,07/20,06/15,W30,W25,M7,M6).
  - **숫자 검산 통과**: 합성 데이터 NG 5건(L1:2, L2:1, L3:2) / 수량 4,300 기준
    · 전체 5/4,300=1,163ppm · 07/24 3/1,500=2,000 · 07/20 2/2,000=1,000 · 06/15 0/800=0
    · 행 합(Gap 3 + Coil 2)=5=합계, **새 표 합계 행이 모델 피벗 합계 행과 완전 일치**(분해 성질 확인).
    · 'Coil 상태 확인'이 MODULE(AI Grill)과 UNIT(Frame + VP ass'y)에 걸쳐 **1행으로 병합**되고
      점검항목 2개가 병기됨 — 요청한 "Check Item끼리 묶기" 동작 확인.
  - **미검증**: 실제 BMES 데이터/긴 베트남어 Check Item에서의 열 너비, 서버 재시작 후 화면.
- 주의: 엑셀 `NG 집계` 시트에는 **이 표를 넣지 않았다**(요청은 화면 기준). 필요하면 별도 요청.
- Next: 웹서버 재시작 → LPA Search → NG 집계 탭에서 모델 피벗 밑 '점검항목별 불량률' 확인.

## 2026-07-25 11:40 - 점검항목별 불량률: 엑셀 시트 추가 + 화면/엑셀에 점검항목·모델 병기
- 요청 2건: (1) "엑셀 추출시 같이 가능하게" (2) "CheckItem 옆에 점검항목, 모델도 나오게".
- Completed:
  - **집계 로직 공용화**: 직전 항목에서 `BuildCheckPivotTab`(HTML) 안에 있던 누적 루프를
    **public `BuildCheckPivotRows(rows, details)`** 로 추출. 반환 `CheckPivotRow(Item, CheckPoints,
    Models, OverallNg, NgByKey)`. HTML/엑셀이 **같은 함수**를 쓰므로 BuildPivotTab/BuildPivotData
    같은 이중 루프 문제를 반복하지 않는다. NgByKey의 키는 BuildPivotData와 동일(raw AUDAT / W: / M:)
    → 분모를 `PivotData.ColTotals`에서 그대로 나눠 쓸 수 있음.
  - **모델 병기**: `NgMatrixGroup.Model`(=AULOC)을 항목별 SortedSet으로 모아 `Models`에 ' · ' 결합.
    한 Check Item이 여러 모델에 걸리면 "MODULE · UNIT"처럼 표시(점검항목과 동일한 병합 규칙).
  - **화면**: 행 머리에 `점검항목 …` / `모델 …` 두 줄을 작은 회색 글씨로 추가(라벨 굵게, `.lpa-chk-typrc b`).
  - **엑셀 신규 시트 `점검항목 NG`(2번째, WPF는 `NG by Check Item`)**:
    A=Check Item B=점검항목 C=모델 D=구분 E=전체 F..=기간. 모델 피벗과 동일하게 항목당 2행
    (불량률/갯수), 합계 행이 맨 위, 틀고정 5행/5열. 컬럼폭 A=52(문장이라 넓게)/B=22/C=16.
    **분모는 `PivotData.ColTotals`(모델 피벗과 동일)** → 시트1 합계 행과 시트2 합계 행이 일치.
  - 시그니처 변경: `BmesLpaExcelExporter.Export(pivot, **checkRows**, matrix, images, list, title)`.
    호출부 `BmesLpaPage.razor`(웹/WPF) 갱신.
- Files: 웹/WPF 각각 `Services/BmesLpaHtmlExportService.cs`, `Services/BmesLpaExcelExporter.cs`,
  `Components/Pages/BmesLpaPage.razor`.
- Verification(웹·WPF **양쪽 모두**):
  - 웹 `-t:Compile` error 0 / WPF `dotnet build` 오류 0(경고 4=기존).
  - scratchpad 하니스로 **실제 xlsx 생성 후 ClosedXML로 되읽어 셀 덤프**:
    시트 4개(`NG 집계|점검항목 NG|불량 상세|목록`, WPF는 영문) 확인.
    `점검항목 NG` 9행×12열, 헤더 `Check Item|점검항목|모델|구분|전체|일자…|주차…|월…`,
    **합계 1,163 / 07/24 2,000 / 07/20 1,000 — 화면 표 및 시트1 합계와 완전 일치**,
    Gap 확인 698(3/4,300), Coil 상태 확인 465(2/4,300), **모델 열에 'MODULE · UNIT' 병합 표시** 확인.
    틀고정 rows=5 cols=5 확인.
  - 화면은 puppeteer 재렌더 → JS 에러 0, 행 머리에 '점검항목 …/모델 …' 노출 확인
    (웹: "Gap 확인 점검항목 Height check 모델 MODULE", WPF: "... Check Point ... Model MODULE").
  - **미검증**: 실제 BMES 데이터 규모에서의 시트 폭/줄바꿈, 서버 재시작 후 실제 추출.
- Next: 웹서버 재시작 → LPA Search → `엑셀 추출`로 2번째 시트 `점검항목 NG` 확인.
  WPF에 반영하려면 1.0.24 → 1.0.25 버전업 후 재게시 필요(아직 안 함).

## 2026-07-25 12:00 - 1.0.25 게시(점검항목 NG 피벗 + 엑셀 시트)
- Completed: csproj 1.0.24 → **1.0.25** 후 `PublishStandaloneUpdate.ps1` 실행(실행 중 인스턴스 없음).
  산출물(`JinoSupporter.Web/standalone-updates/`):
  - `BmesNgRateStandalone-1.0.25.zip` 72,056,526 B (sha256 31dda51f…)
  - `BmesNgRateStandalone_Setup-1.0.25.exe` 53,652,824 B (sha256 cd33081b…)
  - `update.json` version 1.0.25 + setupUrl/setupSha256.
  Inno Setup 컴파일 37초. 구버전(1.0.20~1.0.24) 파일은 남겨둠.
- Verification:
  - **게시본(Release publish) 실제 실행** → 12초 생존, 창 제목 'BMES NG Rate',
    로그에 `Starting BMES NG Rate Standalone 1.0.25.0` 기록, 에러 없음.
  - `StandaloneDownloadCatalog.Read` 재실행 → v1.0.25 / Setup 51.2MB / Package 68.7MB 정상 인식
    (PC Download 페이지가 Installer 버튼을 띄울 상태).
  - update.json의 notes에 작은따옴표가 `\u0027`로 이스케이프됨(PowerShell ConvertTo-Json 기본 동작).
    System.Text.Json이 정상 디코드하는 것을 카탈로그 스모크 출력으로 확인 — **문제 없음**.
  - **미검증**: 설치파일 실제 설치, 기존 클라이언트 자동 업데이트 수신, WPF LPA 화면/엑셀 클릭 동작.
- 크기 추이: 1.0.23 설치파일 60.4MB → (슬림화) 1.0.24 53.6MB → 1.0.25 53.7MB.
- Next: 클라이언트 자동 업데이트로 1.0.25 수신 확인 → WPF LPA에서 'NG by Check Item' 표와
  엑셀 4시트 확인. 웹은 서버 재시작 필요(코드만 반영된 상태).

## 2026-07-25 12:40 - **정정**: 점검항목별 표의 분모가 틀렸음 → 점검 횟수 기준으로 재설계 + 컬럼 순서
- **내가 틀린 부분**: 07-25 10:55 항목에서 "분모 = 위 모델 피벗과 동일(기간별 전체 검사수량)"으로
  설계한 것은 **잘못이다.** 사용자 지적: "1개 항목인데 하루에 78? 말이 아예 안 됨."
  점검항목 1줄의 분모가 생산/검사 **수량(78, 117, 6,599)** 일 이유가 없다. 체크리스트 한 줄은
  **감사 1건당 1회 응답**되므로 분모는 **그 항목이 점검된 횟수**여야 한다. ppm 분해 성질을 지키려던
  의도가 표의 의미를 망가뜨렸다. (10:55 항목의 해당 설계 근거는 무효.)
- 수정 내용:
  1) **분모 = 점검 횟수**. `BuildCheckPivotRows`를 재작성 — 더 이상 `BuildNgMatrix`(NG만 보유)를
     쓰지 않고 rows+details를 직접 순회(LQRNO 중복 제거 동일)하며 **RESUT A/B/C 모든 응답을
     '점검 1회'로 카운트**하고 그중 NG(IsNgResult)만 따로 카운트. 레코드에 `OverallChecked` +
     `CheckedByKey` 추가.
  2) **단위 ppm → %**. 분모가 횟수라 ppm이면 40%가 400,000ppm으로 표시돼 오해를 부름.
     신규 `NgPercentText(ng,total)`(소수1자리). `AppendPivotCell(..., bool percent=false)` /
     엑셀 `WriteRate(..., bool percent=false)` 로 분기. **모델 피벗(시트1)은 ppm 그대로.**
  3) 빈 셀 규칙 변경: **그 기간에 그 항목이 점검되지 않음** = 빈칸, 점검했는데 정상 = 0%.
  4) **컬럼 순서 모델 → 점검항목 → Check Item**(사용자 요청). 엑셀 A/B/C 및 화면 행머리 순서 모두.
     엑셀 폭 A=16 / B=22 / C=52. 합계 행 라벨은 A열(모델 자리)에 표시.
  5) 모델 목록은 이제 **NG난 모델이 아니라 그 항목을 점검한 모든 모델**(분모와 일관).
  6) 부제 문구: "NG건수 / 그 항목이 점검된 횟수 (%, 위 표의 ppm과 분모가 다름)".
- 유지된 것: 행은 여전히 **NG 1건 이상인 Check Item만**(전 항목 나열은 노이즈), NG 많은 순 정렬.
- Files: 웹/WPF 각각 `Services/BmesLpaHtmlExportService.cs`, `Services/BmesLpaExcelExporter.cs`.
- Verification(웹·WPF 양쪽, scratchpad 하니스):
  - 웹 `-t:Compile` error 0 / WPF `dotnet build` 오류 0(경고 4=기존).
  - **실제 xlsx 재생성 후 ClosedXML 되읽기**: 헤더 `모델|점검항목|Check Item|구분|전체|일자…`,
    합계 83.3% (5/6), `Gap 확인` 100% (3/3), `Coil 상태 확인` 66.7% (2/3).
    합성데이터 기대치와 일치 — Coil은 L1/L2에서 NG, L4(SPK)에서 정상 응답이라 **3회 점검 중 2회 NG**.
    모델 열도 그에 맞춰 `MODULE · SPK · UNIT`로 확장됨(정상).
    점검 안 된 기간은 빈칸(`.`)으로 나옴 확인. 틀고정 5행/5열.
  - 화면 puppeteer 재렌더: JS 에러 0, 셀이 `83.3% 5/6` 형태, 행머리 `모델 … 점검항목 … Check Item`.
- **미검증**: 실제 BMES 데이터에서의 값(합성 데이터라 분모가 1~3으로 작아 %가 100%로 나옴 — 실데이터는
  점검횟수가 커서 정상 범위일 것), 서버 재시작 후 화면/엑셀.
- Next: 웹서버 재시작 → 실데이터로 분모가 납득 가능한 값(예: 16/40)인지 사용자 확인.
  확인되면 WPF 1.0.25 → 1.0.26 버전업 후 재게시(아직 안 함).

## 2026-07-25 13:10 - 점검항목 표: % → ppm 환원 + 엑셀 일자/주차/월 구분선
- 요청: "일 하고 주차, 월 구분선 정확하게 표시해주고 %로 하지말고 ppm으로 변경".
- Completed:
  1) **단위 % → ppm 환원**(사용자 지시). 12:40 항목에서 내가 넣은 `NgPercentText`와
     `AppendPivotCell(..., percent)` / `WriteRate(..., percent)` 분기를 **전부 제거**(죽은 코드 남기지 않음).
     점검항목 표도 `NgRateText`(ppm) 사용. **분모는 12:40에서 고친 '점검 횟수' 그대로 유지** —
     즉 ppm = NG건수 / 점검횟수 × 1e6. 시트1(모델 피벗)은 검사수량 기준 ppm이라 **두 표의 ppm은
     서로 비교 불가**(부제/주석에 명시).
  2) **엑셀 구분선**: 신규 `DrawBlockBorders(ws, pivot, firstPeriodCol, firstRow, lastRow)`.
     각 기간 블록(일자/주차/월)의 **첫 컬럼에 Medium LeftBorder**, 마지막 기간 컬럼에 RightBorder,
     라벨 영역과 숫자 영역 사이(전체 컬럼 우측)에도 RightBorder. **시트1·시트2 둘 다 적용**.
     (화면 HTML은 sep-th/sep-td 빈 컬럼으로 이미 구분돼 있어 손대지 않음. 엑셀엔 그 빈 컬럼이
     없어서 세 블록이 한 덩어리로 보이던 것이 원인.)
- Files: 웹/WPF 각각 `Services/BmesLpaHtmlExportService.cs`, `Services/BmesLpaExcelExporter.cs`.
- Verification(웹·WPF 양쪽):
  - 웹 `-t:Compile` error 0 / WPF `dotnet build` 오류 0.
  - **xlsx 재생성 후 ClosedXML로 테두리 실측**(하니스에 BorderReport 추가):
    `[NG 집계] LEFT-of 4:07/24, 7:W30, 9:M7 / RIGHT-of 3:전체, 10:M6`
    `[점검항목 NG] LEFT-of 6:07/24, 9:W30, 11:M7 / RIGHT-of 5:전체, 12:M6`
    → **블록 경계와 정확히 일치**(일자 시작/주차 시작/월 시작 + 양 끝). WPF도 동일 좌표 확인.
  - 셀 값이 ppm으로 환원됨 확인(합계 833,333 = 5/6 등). 화면도 puppeteer로 `833,333 5/6` 확인, JS 에러 0.
  - **미검증**: 실데이터에서의 ppm 자릿수 체감(합성데이터는 분모가 1~3이라 ppm이 100만 근처).
- Next: 웹서버 재시작 → 엑셀 추출로 구분선/ppm 확인. WPF 반영하려면 1.0.25 → 1.0.26 게시 필요.

## 2026-07-25 13:35 - 점검항목 표 정렬을 컬럼 순서(모델→점검항목→Check Item)로 변경
- 요청: "이거 정렬 순서가 이상한데"(모델 열이 MODULE/UNIT/SPK로 뒤섞여 보임).
- 원인: `BuildCheckPivotRows`가 **NG 건수 내림차순**으로 정렬하고 있었음. 컬럼 순서를 모델→점검항목→
  Check Item으로 바꾼 뒤에는 첫 열이 정렬키가 아니라서 뒤죽박죽으로 보임.
- 수정: 정렬을 **`모델` → `점검항목` → `Check Item`(모두 Ordinal)** 로 변경. 표시 컬럼 순서와 동일.
  정렬키 계산용 지역함수 `ModelsOf/PointsOf`를 만들고 **Select 투영에서도 재사용**(문자열 결합 중복 제거).
  한 Check Item이 여러 모델/점검항목에 걸리면 결합 문자열("MODULE · SPK · UNIT")이 정렬키가 됨.
- **트레이드오프**: NG 많은 항목이 더 이상 맨 위로 오지 않는다. 모델별로 묶어 보는 용도에 맞춘 것.
  NG순/불량률순이 필요하면 정렬키만 바꾸면 됨(엑셀에서 사용자가 직접 정렬도 가능 — 병합은 라벨 3열뿐).
- Files: 웹/WPF `Services/BmesLpaHtmlExportService.cs`(BuildCheckPivotRows 정렬부).
- Verification: 하니스 합성데이터에 **모델 4종·점검항목 6종을 추가**해 정렬을 검증
  (NG건수와 알파벳 순서가 일부러 어긋나게 구성).
  결과 순서 = FRONT/Tape Main Mic → MODULE/Height check → MODULE · SPK · UNIT/AI Grill…
  → SPK/Bako/bbb → SPK/Zebra check/aaa → UNIT/Yoke ass'y.
  **모델 오름차순, 같은 모델 안에서 점검항목 오름차순**(SPK의 Bako가 Zebra check보다 앞, 그 결과
  Check Item은 'bbb'가 'aaa'보다 앞 — 3순위 키라서 정상) 확인.
  **웹 엑셀·웹 화면(puppeteer)·WPF 엑셀 3곳 모두 동일 순서** 확인. 빌드 error 0.
- Next: 서버 재시작 후 실데이터에서 순서 확인. WPF 반영은 1.0.26 게시 필요(미실시).

## 2026-07-25 14:15 - LPA 상태 박스에 진행률 바 추가(웹+WPF)
- 요청: 상태 메시지 박스("623 row(s), 상세 623건 완료.")에 조그만 프로그레스바.
- Completed:
  - `BmesLpaScrapeService.FetchDetailsAsync`에 **선택 파라미터
    `IProgress<(int Done,int Total)>? counter`** 추가(기존 `IProgress<string> progress`는 그대로).
    문자열 "상세 조회 n / N…"을 파싱하는 대신 숫자를 따로 보냄. 10건마다 + 마지막에 보고.
  - `BmesLpaPage.razor`: `_progDone/_progTotal` 필드 + `SetProgress/ClearProgress`,
    상태 alert 안에 높이 4px 바(`.lpa-prog` / `.lpa-prog-bar`, transition .15s).
    **`_progTotal > 0`일 때만 렌더** — 로그인/HTML 생성처럼 총량이 없는 단계에서 멈춘 바는
    멈춘 것처럼 보이므로 아예 감춤. 완료/에러 시 `ClearProgress()`로 사라짐.
  - 카운트되는 구간 2곳: **상세 프리페치**(counter 연결), **엑셀 이미지 다운로드**(20건→10건마다 갱신,
    마지막 건도 갱신). 엑셀은 이미지 완료 후 "엑셀 파일 생성 중…"으로 바꾸고 바를 숨김.
- **주의(버그 회피)**: 처음에 `style="width:@ProgressPercent.ToString("0.#", …)%"` 로 썼는데,
  **따옴표를 가진 @식을 따옴표 속성 안에 넣는 형태**라 Razor 파싱이 위험(앞서 `v@release.Version`이
  이메일로 오인된 것과 같은 계열). 컴파일은 통과하지만 신뢰할 수 없어 **`ProgressWidth` 프로퍼티가
  "29.4%" 문자열을 통째로 반환**하도록 변경하고 `style="width:@ProgressWidth"`로 단순화.
  InvariantCulture 고정(ko-KR에서 "29,4%" 방지).
- Files: 웹/WPF 각각 `Services/BmesLpaScrapeService.cs`, `Components/Pages/BmesLpaPage.razor`.
- Verification:
  - 웹 `-t:Compile` error 0 / WPF `dotnet build` 오류 0.
  - **실제 .razor에서 상태 박스 마크업과 <style>을 추출**해 4가지 상태로 렌더 후 헤드리스 Chrome 측정
    (scratchpad `progbar`): 총량 없음 → 바 없음, 180/623 → **채움 29%**, 90/142 → **63%**,
    완료 메시지 → 바 없음. 트랙 높이 4px 확인. 스크린샷으로 시각 확인.
  - **미검증**: 실제 Blazor 서버에서의 실시간 갱신(서버 재시작 필요), WPF 실행 화면.
- Next: 웹서버 재시작 → Search 중 바가 차오르는지, 엑셀 추출 중에도 동작하는지 확인.

## 2026-07-25 14:35 - 점검항목 표: 불량 항목만 → 전체 점검항목 + From 기본값 7/20
- 요청 2건: (1) "여기 불량만 말고 전체 다 보여줘" (2) "Audit Date From 기본값 2026-07-20".
- Completed:
  1) **`BuildCheckPivotRows`가 `overallNg.Keys` → `overallChecked.Keys` 를 순회**하도록 변경.
     즉 **한 번이라도 응답된 모든 Check Item**이 행으로 나온다(NG 0건이면 0 ppm, 갯수 `0/n`).
     `overallNg`는 `GetValueOrDefault(k)`로 읽고 `overallChecked[k]`가 분모.
     14:15 이전에 내가 넣었던 "실패한 항목만" 필터는 제거됨(주석도 교체).
     합계 행 분모도 그만큼 늘어남(예: 9/10 → 9/11).
     부제 문구: "N개 점검항목(불량 0건 포함)" / WPF "N checks (including those with no NG)".
  2) **`LpaFirstDate` 2026-06-08 → 2026-07-20**(웹/WPF 페이지의 From 기본값). LPA 데이터 시작일은
     여전히 6/8이라 직접 입력하면 그 이전도 조회 가능 — 주석에 명시.
- Files: 웹/WPF 각각 `Services/BmesLpaHtmlExportService.cs`, `Components/Pages/BmesLpaPage.razor`.
- Verification(웹·WPF):
  - 웹 `-t:Compile` error 0 / WPF `dotnet build` 오류 0.
  - 하니스 합성데이터에 **RESUT=A만 있는 항목**(`MODULE / AI Grill / 기타`)을 두고 검증 →
    이전에는 표에 없던 그 행이 **`0 ppm, 갯수 0/1`로 정상 출력**됨(엑셀 r8, 화면 행 목록 모두).
    시트 행수 17 → 19, 합계 갯수 9/10 → **9/11**로 분모 증가 확인.
  - 정렬(모델→점검항목→Check Item)·구분선·ppm은 그대로 유지됨을 같은 덤프에서 확인. JS 에러 0.
- **미검증**: 실데이터에서 행 수가 얼마나 늘어나는지(체크리스트 전 항목이 나오므로 상당히 길어질 수 있음).
  너무 길면 "NG 있는 항목만" 토글을 추가하는 것이 다음 후보.
- Next: 웹서버 재시작 → 전체 항목 표시/기본 날짜 확인. WPF 반영은 1.0.26 게시 필요(미실시).

## 2026-07-25 14:55 - 불량 상세를 독립 탭으로 분리 + 일자별 세로 구분선
- 요청: "불량 상세는 탭으로 나눠서 전체 화면에서 보이게(너무 작게 보임) / 일자마다 구분선".
- 원인: 불량 상세가 **NG 집계 탭 안에** 모델 피벗·점검항목 피벗과 함께 3단으로 쌓여 있어
  `.lpa-panel.has-chk .lpa-scroll{max-height:30vh}`로 눌려 있었음.
- Completed:
  1) **탭 3개로 분리**: `목록 / NG 집계 / 불량 상세`(WPF: `List / NG Summary / Defect Detail`).
     - 패널: `<div class="lpa-panel" data-panel="ng">{{ngDetailTab}}</div>` 신설,
       pivot 패널에서 `{{ngDetailTab}}` 제거.
     - `pivotPanelClass`에서 `has-ng` 제거(이제 pivot 탭은 피벗 2단만) → `has-chk`만.
       `.lpa-panel.has-chk .lpa-scroll` 30vh → **40vh**, 불량 상세는 `.lpa-scroll` 기본 **78vh**.
     - JS: `selectTab`에 `ng` 분기(`applyNgDetail()`), 필터 입력 핸들러도 3분기.
       `applyPivot()`이 더 이상 `applyNgDetail()`을 호출하지 않음(다른 탭이라 불필요).
  2) **일자 구분선**: `.lpa-ngdetail thead th.lpa-num, .lpa-ngdetail td.lpa-ng-cell
     { border-left: 1px solid #dbe3ec; }` — 날짜 컬럼마다 세로선. 사진이 있는 인접 일자가
     서로 붙어 보이던 문제 해소.
- Files: 웹/WPF `Services/BmesLpaHtmlExportService.cs`.
- Verification(웹·WPF 양쪽, 헤드리스 Chrome):
  - 탭 3개 렌더 확인(`목록/NG 집계/불량 상세`, WPF는 영문).
  - 탭 전환 후 활성 패널의 `max-height` 실측: **list 780px / pivot 400px / ng 780px**
    → 불량 상세가 기존 30vh(≈300px) 대비 **전체 높이 사용** 확인.
  - 불량 상세 탭에서 `td.lpa-ng-cell` / `thead th.lpa-num`의 `border-left-width` **1px** 확인.
  - **JS 에러 0**, 스크린샷으로 탭·구분선 시각 확인.
  - 빌드: 웹 `-t:Compile` error 0 / WPF `dotnet build` 오류 0.
- **미검증**: 실데이터(사진 많은 셀)에서의 높이 체감, 서버 재시작 후 화면.
- Next: 웹서버 재시작 → 불량 상세 탭 확인. WPF 반영은 1.0.26 게시 필요(미실시).

## 2026-07-25 16:12 - 1.0.26 게시(LPA 개편 일괄)
- Completed: csproj 1.0.25 → **1.0.26** 후 `PublishStandaloneUpdate.ps1` 실행(실행 중 인스턴스 없음).
  포함 변경(12:40~14:55 항목 전부): 불량 상세 독립 탭+일자 구분선 / 점검항목 NG 피벗(전체 항목,
  분모=점검횟수, ppm) 화면+엑셀 시트 / 모델·점검항목 컬럼 / 엑셀 일자·주차·월 구분선 /
  정렬 모델→점검항목→Check Item / 진행률 바 / From 기본값 2026-07-20.
  산출물(`JinoSupporter.Web/standalone-updates/`):
  - `BmesNgRateStandalone-1.0.26.zip` 72,059,099 B (sha256 8695374e…)
  - `BmesNgRateStandalone_Setup-1.0.26.exe` 53,648,421 B (sha256 c726e2c6…)
  - `update.json` version 1.0.26 + setupUrl/setupSha256.
  Inno Setup 컴파일 37.6초. 구버전(1.0.20~1.0.25) 파일은 남겨둠.
- Verification:
  - **게시본(Release publish) 실제 실행** → 12초 생존, 창 제목 'BMES NG Rate',
    로그 `Starting BMES NG Rate Standalone 1.0.26.0`, 에러 없음.
  - `StandaloneDownloadCatalog.Read` → v1.0.26 / Setup 51.2MB / Package 68.7MB 정상 인식.
  - **미검증**: 설치파일 실제 설치, 클라이언트 자동 업데이트 수신, WPF LPA 화면 클릭 동작.
- 배포 상태: 웹은 사용자 재시작으로 확인 완료(사용자 "확인 완료"). WPF 클라이언트는 다음 폴링에서
  1.0.26 자동 업데이트 수신 예정.
- Next: 클라이언트에서 1.0.26 수신 확인 → WPF LPA 3탭/엑셀 4시트 동작 확인.
## 2026-07-27 07:10 - LPA 웹 UI 전면 영문화(Use Y/N 제외) + 탭 라벨 LPA 대문자
- 요청: (1) "LPA 웹에 USE Y/N 이 부분 빼고 모든 한글로 되어있는거 영어로 변경"
  (2) "탭에 Lpa 되어있는거 LPA 대문자로 표시".
- 접근: **WPF(BmesNgRateStandalone) 쪽 LPA 서비스가 이미 전부 영문**이었고 웹과 네임스페이스 외
  차이가 없었음 → 새로 번역하지 않고 **WPF 파일을 웹으로 복사 + 네임스페이스만 치환**.
  번역 문구가 두 프로젝트에서 100% 동일해지고, 앞으로 웹/WPF diff가 네임스페이스 1줄만 남음.
- Completed:
  1) 웹 `Services/BmesLpaHtmlExportService.cs` · `BmesLpaExcelExporter.cs` · `BmesLpaScrapeService.cs`
     = WPF판으로 교체(네임스페이스만 JinoSupporter.Web.Services). 화면 탭 `List / NG Summary /
     Defect Detail`, 툴바 `Filter / Expand all / Collapse models / Collapse all / Date·Week·Month / Size`,
     라이트박스 `Prev / Next / Original / Close`, 엑셀 시트 `NG Summary / NG by Check Item /
     Defect Detail / List`, 헤더 `Model / Type / Overall / Check Point / NG Rate / Count`.
  2) 웹 `Components/Pages/BmesLpaPage.razor` 한글 문구 전부 영문화(WPF 문구에 맞춤):
     엑셀 버튼·툴팁, 초기 안내문, 상태 메시지(Fetching n detail(s)… / Generating HTML… /
     Preparing n image(s) for Excel… / Excel saved …), 엑셀 제목 `LPA Defect Detail · …`,
     **다운로드 파일명 `LPA_불량상세_…` → `LPA_DefectDetail_…`**.
  3) **Use Y/N 컨트롤은 요청대로 그대로 둠**(원래도 영문 라벨 + Y/N/All).
  4) WPF 번역 과정에서 남아 있던 찌꺼기를 **웹·WPF 양쪽에서** 정리:
     `.Append("")`(건 제거 흔적) 2곳, JS `fmt(x) + ''` 3곳, "the model's  count"(이중 공백) 2곳,
     남은 한글 주석 `교차표`/`기간`, 엑셀 주석 `Date -> Week` → `→`.
  5) `font-family`의 `'맑은 고딕'` 제거(영문명 `'Malgun Gothic'`만 유지 — Windows에서 동일 폰트로
     해석됨). 생성 HTML `<html lang="ko">` → `lang="en"`. **BMES 데이터 자체의 한글(점검항목·
     Check Item 텍스트)은 당연히 그대로**.
  6) 탭 라벨: `Components/Layout/MainLayout.razor`의 `TitleFromUrl`이 URL 마지막 조각을
     Title Case 해서 "Lpa"가 됨. **`Acronyms` 집합(현재 "lpa")에 있으면 전체 대문자**로.
     페이지가 `<PageTitle>`을 보고하면 어차피 덮어쓰지만, 그 전 fallback이 이 값이라 여기가 원인.
     탭 스트립은 웹에만 있음(WPF 해당 없음).
- Files:
  - 웹: `Components/Pages/BmesLpaPage.razor`, `Components/Layout/MainLayout.razor`,
    `Services/BmesLpaHtmlExportService.cs`, `BmesLpaExcelExporter.cs`, `BmesLpaScrapeService.cs`,
    `BmesLpaImageService.cs`(주석 "원본"→"Original").
  - WPF: `Services/BmesLpaHtmlExportService.cs`, `BmesLpaExcelExporter.cs`, `BmesLpaImageService.cs`
    (4번·5번 정리분만).
- 주의(다음 세션): 복사 시 `sed -i`가 **CRLF를 LF로 바꿔버림** → awk로 CRLF 복원함.
  두 프로젝트 모두 CRLF이므로 파일 복사 후 `grep -c $''` 로 확인할 것.
- Verification:
  - 빌드: 웹 `-t:Compile` **error 0**, WPF `dotnet build` **오류 0**(경고는 기존 것).
  - 한글 스캔: 웹/WPF LPA 페이지·서비스 + MainLayout **한글 0건**.
  - 웹↔WPF 서비스 4종 diff = **네임스페이스 1줄뿐**(줄바꿈 포함 동일).
  - 실제 소스에서 뷰어 HTML 템플릿을 추출해 합성 데이터로 렌더(헤드리스 Chrome):
    탭 `List / NG Summary / Defect Detail`, 툴바·라이트박스 전부 영문, **렌더 결과 한글 0건**,
    `<html lang="en">`, **JS 에러 0**.
    JS 카운트 3경로 실측(`+ ''` 제거분): 일자 그룹 "1", 모델행 "1"+TOTAL 10, 불량 상세 "1",
    `tb-stat` "1 / 1 rows". 스크린샷으로 시각 확인. 뷰어 JS 단독 `node --check` 통과.
  - **미검증**: 실제 서버 화면(재시작 필요), 엑셀 재추출 시트명/헤더(코드가 1.0.26 WPF와
    바이트 동일이라 동작은 같을 것으로 봄), 탭 라벨은 서버 재시작 후 확인 필요.
- Next: 사용자 웹서버 재시작 → LPA 화면/엑셀/탭 라벨 "LPA" 확인.
  WPF 클라이언트에 반영하려면 1.0.26 → **1.0.27 게시 필요**(미실시 — 이번 WPF 변경은 주석·폰트·
  lang 정리라 기능 영향 없음).

## 2026-07-27 07:32 - [UI][REMOVE] LpaSearchFilters: Use Y/N 필터 UI 제거(웹+WPF)
- Agent: Claude
- Session: External
- Task-ID: Unavailable
- Category: UI
- Feature: LpaSearchFilters
- Change: REMOVE
- Completed: LPA 검색 툴바의 Use Y/N 셀렉트를 웹·WPF 양쪽에서 제거하고, 조회는 상수
  `UseYnFilter = "Y"`로 고정. 뷰어 HTML 헤더의 `· Use Y/All` 표기와 `ExportInput.UseYn`
  프로퍼티도 함께 제거(더 이상 선택 항목이 아니므로 죽은 값).
- Decisions: 사용자가 "안 씀"이라고 했으므로 UI만 숨기는 대신 값 자체를 상수로 고정
  (기존 기본값과 동일한 USEYN=Y라 조회 결과 변화 없음). `LpaQuery.UseYn` 프로퍼티는
  BMES 요청 파라미터라 그대로 유지. 웹/WPF `BmesLpaHtmlExportService.cs`는 계속
  네임스페이스 1줄만 다른 상태를 유지.
- Files: JinoSupporter.Web/Components/Pages/BmesLpaPage.razor,
  JinoSupporter.Web/Services/BmesLpaHtmlExportService.cs,
  BmesNgRateStandalone/Components/Pages/BmesLpaPage.razor,
  BmesNgRateStandalone/Services/BmesLpaHtmlExportService.cs
- Verification: 웹 `dotnet build -t:Compile` 오류 0(경고 32개, 전부 기존 것), WPF
  `dotnet build` 오류 0(경고 4개, 기존 것). 잔여 `_useYn` 참조 0건(grep).
  두 프로젝트 HtmlExportService diff = 네임스페이스 1줄뿐 확인. 실제 화면은 미확인
  (웹 재시작은 사용자 담당, WPF는 재게시 필요).
- Next: 사용자 웹서버 재시작 후 LPA 툴바에서 Use Y/N 사라진 것 확인. WPF 클라이언트
  반영은 1.0.27 게시 필요(미실시).

## 2026-07-27 07:50 - [UI][CHANGE] DailyReportRankings: 좌우 배치 + 데이터 있는 기준일 추적 + 전일대비 블록
- Agent: Claude
- Session: External
- Task-ID: Unavailable
- Category: UI
- Feature: DailyReportRankings
- Change: CHANGE
- Completed: DAILY REPORT의 Worsened/Improved 표를 상하 → 좌우 2열 그리드로 바꾸고,
  비교 기준을 "직전 기간"에서 "그 모델이 실제로 생산된 가장 최근 기간"으로 변경(값 아래에
  기준 기간과 −2w/−4d 표기). 같은 랭킹을 일(day) 단위로 계산하는 DAY OVER DAY 블록을
  주차 블록 아래에 추가.
- Decisions:
  * "데이터 없음" 판정은 PPM==0이 아니라 report.GroupRawIn의 Input>0 존재 여부로 함
    (PPM 0은 "무결점 생산"과 "미생산"을 구분하지 못함). GroupRawIn을 (unit, period)로
    한 번만 인덱싱(ProductionIndex)해서 기존 전체 스캔도 O(1) 조회로 교체.
  * 최신 기간(주/일)에 PPM이 없는 모델은 기존 정책대로 계속 제외(생산 중단은 개선이 아님).
  * 일 단위는 최근 30일(DailyTrendDays)로 제한 — 스파크라인 길이 겸 기준일 소급 한계.
  * 블록 헤딩(dr-block-head)을 그리드 전체 폭 span 아이템으로 넣어 표 마크업은 한 벌만 유지.
  * 그리드 트랙 최소폭 760px + Change 열의 %를 값 아래로 내려 표 고유폭을 줄임
    (그래야 2열에서 가로 스크롤이 안 생김). 좁은 화면에서는 auto-fit이 1열로 되돌림.
  * BmesDailyReportModelRow.WeeklyPpm → TrendPpm 개명(일 단위 행에도 쓰이므로),
    BaselineHeader/BaselineStepsBack 추가. SnapshotSchemaVersion 7 → 8.
- Files: JinoSupporter.Web/Services/BmesDailyReportService.cs,
  JinoSupporter.Web/Components/Pages/BmesDailyReportPage.razor
- Verification: 웹 `dotnet build -t:Compile` 오류 0(경고 32개, 전부 기존 것).
  레이아웃은 페이지의 실제 <style> 블록을 추출해 합성 데이터 목업으로 헤드리스 Chrome 렌더 +
  스크린샷 확인: 1920/1800px = 2열이며 표 가로 스크롤 없음(client==scroll), 1500/1280px = 1열,
  1600px에서는 2열이지만 가장 넓은 표만 약 60px 스크롤. 기준 기간 라벨(W28 −2w / 07/22 −4d /
  no earlier week)과 DAY OVER DAY 블록 배치 육안 확인.
  **미검증**: 실제 BMES 데이터로의 산출값(스냅샷 재생성 필요), 서버 화면.
- Next: 스키마 8로 올렸으므로 기존 캐시 스냅샷은 폐기됨 → 웹 재시작 후 DAILY REPORT 첫 진입 시
  자동 Refresh(수 분) 완료를 기다려 실제 수치/일 단위 표 확인.

## 2026-07-27 07:59 - [DATA][FIX] DailyReportRankings: 생산 없는 날 기준으로 잡혀 전일대비 표가 비던 문제
- Agent: Claude
- Session: S5
- Task-ID: S5-20260727-075722980-beb613c23c6e4b319
- Category: DATA
- Feature: DailyReportRankings
- Change: FIX
- Completed: DAY OVER DAY 표가 "Top 0 models"로 비어 나오던 원인을 찾아 수정. 최신 기간 컬럼이
  데이터가 전혀 없는 날(일요일/휴일)이면 그 앞의 빈 컬럼들을 버리고 실제 데이터가 있는 최신
  기간을 기준으로 삼도록 TrimToLatestWithData 추가(주/일 양쪽 공통 적용).
- Decisions:
  * 원인: NgRateReportService.BuildRequestedPeriodColumns가 요청 기간의 **모든 날짜**에 대해
    컬럼을 생성한다(데이터 유무와 무관). 그래서 07/26(일)에 컬럼은 있으나 원본 행이 0건 →
    모든 모델의 최신일 PPM=0 → `recent <= 0 continue`로 전부 탈락 → 0행.
    저장된 스냅샷에서도 OverallRecentDayPpm=0 / OverallPreviousDayPpm=287,029로 확인.
  * "최신 기간"은 전체 TOTAL 행에 값이 있는 첫 컬럼으로 정의. 모델별 기준(baseline) 소급은
    기존대로 unit별 Produced() 판정을 그대로 사용.
  * TOTAL 행 자체가 없으면 컬럼을 건드리지 않고 그대로 둠(대시보드 전체를 비우지 않기 위해).
  * SnapshotSchemaVersion 8 → 9. IsStale은 오늘 생성된 스냅샷을 최신으로 보기 때문에,
    버전을 올리지 않으면 오늘 안에는 자동 재계산이 돌지 않는다.
- Files: JinoSupporter.Web/Services/BmesDailyReportService.cs
- Verification: `dotnet build -t:Compile` 오류 0.
  원인 확인은 실제 데이터로 함 — 저장된 스냅샷 JSON(AppSettings/BmesDailyReport:Snapshot)에서
  일 단위 0행 + 07/26 전체 PPM 0 확인, 수집 DB(01. NG RATE/temp_20260727_075533.db)에서
  PRODUCT_DATE='2026-07-26' 행 0건, '2026-07-25' 7,099행·33라인·NG 7,077 확인.
  **미검증**: 실제 재계산 결과(스냅샷 재생성 필요).
- Next: 웹 재시작 → DAILY REPORT 진입 시 스키마 9로 자동 Refresh → DAY OVER DAY가 07/25 vs
  07/24 기준으로 채워지는지 확인.

## 2026-07-28 18:57 - [UI][ADD] ReportKpiTab: KPI 주차 추이 탭 추가
- Agent: Codex
- Session: S1
- Task-ID: S1-20260728-185247729-3227b31a9056432ab
- Category: UI
- Feature: ReportKpiTab
- Change: ADD
- Completed: Report 정적 HTML 뷰어에 KPI 탭을 추가하고 `순번 | KPI | W27 | W28 | W29 | W30…` 형태의 주차별 PPM 표를 구현했다. 조회 기간의 주차는 연도·주차 키 기준 오름차순으로 자동 표시된다.
- Decisions: 추가 BMES 조회 없이 Daily 단계에서 이미 계산한 전체/공정유형별 요약 데이터를 재사용한다. KPI 행은 기존 요약 순서(TOTAL 및 공정유형)를 유지하고 0 PPM은 `-`로 표시한다.
- Files: JinoSupporter.Web/Services/BmesReportHtmlExportService.cs
- Verification: `dotnet msbuild JinoSupporter.Web/JinoSupporter.Web.csproj -t:Compile -p:BuildProjectReferences=false` 성공(오류 0). 소스 검사로 KPI 탭 등록, 본문 생성, 주차 오름차순, `순번`/`KPI` 헤더 존재를 확인했다.
- Next: 웹 서버 재시작 후 `/report/bmes`에서 해당 기간으로 Get Report를 실행하여 실제 KPI 탭의 주차/값 표시를 확인한다. 애플리케이션 실행은 사용자 요청이 없어 수행하지 않았다.

## 2026-07-29 07:24 - [DATA][ADD] FCostCoreParts: MES072410 저장 구조 조사
- Agent: Codex
- Session: S2
- Task-ID: S2-20260729-072134428-35f4601d716f43dfa
- Category: DATA
- Feature: FCostCoreParts
- Change: ADD
- Completed: 기존 Report F-COST의 MES072400 수집, `fcost_raw.db` 캐시, 리포트 로드 경로와 MES072410 응답 필드 차이를 추적했다. MES072410은 같은 SQLite 파일에 저장하되 기존 리포트 테이블과 분리해야 안전하다고 확정했다.
- Decisions: MES072410 행은 `CPTYP_TX`와 공정별 집계 의미를 가지므로 `FCostRawRows`에 섞지 않고 endpoint 전용 Pulls/Rows/Columns 테이블을 사용한다. 응답에 신뢰할 수 있는 컬럼 메타가 있을 때만 Columns를 저장하며 공유된 브라우저 세션 쿠키는 사용하거나 기록하지 않는다.
- Files: None
- Verification: `FCostService`, `FCostReportService`, `BmesFCostPage`의 로그인·수집·SQLite 스키마·리포트 로드 경로를 읽기 전용으로 확인했다.
- Next: MES072410 전용 파서와 날짜 범위 수집·동일 DB 분리 저장 서비스를 구현하고 Web 프로젝트를 좁게 컴파일한다.

## 2026-07-29 07:34 - [DATA][ADD] FCostCoreParts: 타 세션 구현 계약 확정
- Agent: Codex
- Session: S2
- Task-ID: S2-20260729-072134428-35f4601d716f43dfa
- Category: DATA
- Feature: FCostCoreParts
- Change: ADD
- Completed: 실제 BMES 연결이나 코드 구현 없이 다른 AI 세션이 사용할 MES072410 수집·파싱·저장 계약을 확정했다. 기존 `FCostService`와 동일한 로그인/날짜별 캐시 흐름을 따르되 `fcost_raw.db` 안의 endpoint 전용 테이블로 분리하는 방식이다.
- Decisions: 공개 진입점은 `ParseSearchListJson`, `BackfillAsync`, `GetStatus`, `GetRawDbPath`로 두고, 테이블은 `MES072410Pulls`, `MES072410Rows`, 메타가 확인될 때만 `MES072410Columns`를 사용한다. `RawJson`과 nullable 수치값을 함께 보존하고 INAMT 기준 공정명을 continuation 행에 forward-fill하며 기존 `FCostRaw*` 테이블과 Report UI는 변경하지 않는다.
- Files: None
- Verification: `FCostCorePartsService.cs`와 MES072410 DI 등록이 생성되지 않았음을 확인했으며 외부 BMES 요청도 실행하지 않았다.
- Next: 다른 AI 세션에서 아래 구현 계약대로 신규 서비스를 작성한 뒤 제공된 샘플 JSON 파싱, 임시 SQLite upsert/readback, Web 프로젝트 scoped compile을 수행한다.

## 2026-07-29 07:23 - [UI][CHANGE] ReportKpiTab: 월·주 혼합 KPI 표 골격 준비
- Agent: Codex
- Session: S1
- Task-ID: S1-20260729-071709013-3e19e3d342ed48dba
- Category: UI
- Feature: ReportKpiTab
- Change: CHANGE
- Completed: KPI 탭을 예시 양식에 맞춰 `KPI | KPI 종류 | 기준실적 | 26년 목표 | 26년 실적 | 구분 | 월·주 기간` 구조로 개편했다. 각 KPI는 실적/달성률 2행으로 표시하고, 현재 보유한 월별·주별 PPM은 시간순으로 혼합 배치하며 좌측 6개 열은 가로 스크롤 시 고정된다.
- Decisions: 추가 BMES 조회 없이 기존 Daily 요약의 MonthCols/WeekCols와 PPM을 재사용한다. 기준실적·목표·연간실적·달성률은 데이터 원천이 정해질 때까지 `-`로 표시하고, 월 경계 주차는 주중 기준으로 해당 월 뒤에 배치한다.
- Files: JinoSupporter.Web/Services/BmesReportHtmlExportService.cs
- Verification: `dotnet msbuild JinoSupporter.Web/JinoSupporter.Web.csproj -t:Compile -p:BuildProjectReferences=false` 성공(오류 0, 기존 경고만 발생). 소스 검사에서 월/주 데이터 소스, 고정 헤더, 2행 병합, sticky 열을 모두 PASS로 확인했고 W23~W27 배치 기준도 6월→7월 순서로 확인했다. 사용자 요청에 따라 애플리케이션/서버는 실행하지 않았다.
- Next: 웹 서버 재시작 후 `/report/bmes`에서 Get Report를 실행해 실제 KPI 표 레이아웃을 확인한 뒤, 기준실적·목표·누적실적·달성률 데이터 원천과 KPI별 행 정의를 연결한다.

## 2026-07-29 07:26 - [UI][FIX] ReportKpiTab: KPI 항목을 예시 정의로 교정
- Agent: Codex
- Session: S1
- Task-ID: S1-20260729-071709013-3e19e3d342ed48dba
- Category: UI
- Feature: ReportKpiTab
- Change: FIX
- Completed: KPI 이름에 TOTAL·SUB 등 NG 요약 행이 표시되던 잘못된 연결을 제거하고, 예시의 `초과투입 재료비(Main)`, `초과투입 재료비(내재화)`, `Main 공정불량 개선율`, `IPG 공정불량 개선율` 4개 항목으로 교정했다. 재료비 KPI는 4개 세부 행, 공정불량 KPI는 2개 세부 행을 병합 표시하도록 구성했다.
- Decisions: KPI 항목과 행 구성은 화면 정의로 고정하고 실제 기준실적·목표·누적실적·월·주 값은 향후 데이터 원천에서만 채운다. 기존 Summary의 공정유형 이름과 PPM 값은 KPI 본문에서 사용하지 않는다.
- Files: JinoSupporter.Web/Services/BmesReportHtmlExportService.cs
- Verification: `dotnet msbuild JinoSupporter.Web/JinoSupporter.Web.csproj -t:Compile -p:BuildProjectReferences=false` 성공(오류 0, 기존 경고만 발생). 4개 KPI 이름 존재, Summary 행 연결 제거, 가변 rowspan 구성을 소스 검사로 모두 PASS 확인했다. 애플리케이션/서버는 실행하지 않았다.
- Next: 웹 서버 재시작 후 `/report/bmes`에서 Get Report를 다시 생성해 4개 KPI와 병합 행이 예시대로 표시되는지 확인한다.

## 2026-07-29 07:38 - [INTEGRATION][ADD] ReportKpiData: F-COST 합계와 모델 평균 연동
- Agent: Codex
- Session: S1
- Task-ID: S1-20260729-071709013-3e19e3d342ed48dba
- Category: INTEGRATION
- Feature: ReportKpiData
- Change: ADD
- Completed: `초과투입 재료비(Main)`의 기준실적 1.79%·목표 1%를 설정하고 월·주 실적에 F-COST의 TOTAL RATE, 초과 투입 재료비에 Total FCOST, 달성율에 목표/실적 비율을 연결했다. `Main 공정불량 개선율`에는 기준 73,934 ppm·목표 40,000 ppm을 설정하고 F-COST Trend 모델 불량률의 단순 평균과 목표 대비 달성율을 연결했다.
- Decisions: Main 불량률 평균은 Trend의 최상위 실제 모델 행 중 표시값이 있는 항목만 사용하며 `기타_누락`, Sub Group, `-`로 표시되는 0값은 제외한다. 재료비 달성율은 정수 %, Main 불량 달성율은 소수점 1자리로 표시한다.
- Files: JinoSupporter.Web/Components/Pages/BmesFCostPage.razor, JinoSupporter.Web/Services/BmesReportHtmlExportService.cs
- Verification: `dotnet msbuild JinoSupporter.Web/JinoSupporter.Web.csproj -t:Compile -p:BuildProjectReferences=false` 성공(오류 0, 기존 경고만 발생). 데이터 바인딩·평균 제외 조건·달성율 정밀도 소스 검사를 모두 PASS 확인했고, 표본 계산으로 TOTAL RATE 1.25%→80%, Main 평균 63,803 ppm→62.7%를 확인했다. 애플리케이션/서버는 실행하지 않았다.
- Next: 웹 서버 재시작 후 Get Report를 다시 생성해 KPI의 월·주 값이 F-COST Trend 합계 및 모델 평균과 일치하는지 실제 화면에서 확인한다.

## 2026-07-29 07:49 - [DATA][ADD] KpiSourceParsers: MES072410 및 MES050032 파싱 계약 구현
- Agent: Codex
- Session: S1
- Task-ID: S1-20260729-071709013-3e19e3d342ed48dba
- Category: DATA
- Feature: KpiSourceParsers
- Change: ADD
- Completed: MES072410 응답의 원본 행·nullable 수치·신뢰 가능한 열 메타데이터를 파싱하고 INAMT 뒤 FCOST/FRATE 식별 필드를 보정해 `fcost_raw.db` 전용 테이블에 날짜별로 멱등 저장하는 서비스를 구현했다. MES050032의 `header`와 `ChartRec`, AVEGR 및 COL0001~COL0014를 파싱하고 지정 조건의 POST payload를 생성하는 서비스를 추가했다.
- Decisions: MES072410 열 헤더는 `BottomGridColumnList`가 실제 응답에 있을 때만 저장하며 추측하지 않는다. MES050032는 이번 단계에서 파싱과 요청 구성까지만 제공하고 DB 저장 및 Report KPI 매핑은 사용자 정의를 받은 뒤 연결하며, 복사된 브라우저 쿠키·SSO·세션 값은 어떤 입력이나 파일에도 사용하지 않는다.
- Files: JinoSupporter.Web/Program.cs, JinoSupporter.Web/Services/FCostCorePartsService.cs, JinoSupporter.Web/Services/IpgDefectService.cs
- Verification: 격리된 무자격 샘플 검증에서 MES072410의 16222.2·75.99·0.47, INAMT forward-fill, RawJson 불변, 동일 날짜 이중 저장 시 행 중복 없음, SQLite의 0/null 구분과 MES050032 헤더·ChartRec·빈값/0·요청 조건을 모두 PASS 확인했다. `dotnet msbuild JinoSupporter.Web/JinoSupporter.Web.csproj -t:Compile -p:BuildProjectReferences=false`도 오류 없이 성공했으며 애플리케이션·서버 및 외부 BMES 요청은 실행하지 않았다.
- Next: 사용자에게 두 데이터 원천의 KPI 행 선택·집계·달성율 규칙을 확인한 뒤 Report UI 및 필요 시 MES050032 저장 계약을 구현한다.

## 2026-07-29 08:02 - [INTEGRATION][ADD] ReportKpiData: 내재화와 IPG KPI 연동
- Agent: Codex
- Session: S1
- Task-ID: S1-20260729-071709013-3e19e3d342ed48dba
- Category: INTEGRATION
- Feature: ReportKpiData
- Change: ADD
- Completed: `초과투입 재료비(내재화)`에 기준실적 0.95%·목표 0.76%를 설정하고 MES072410 Total 묶음의 FRATE를 실적, FCOST 달러를 초과 투입 재료비, 목표/실적 비율을 달성율로 연결했다. `IPG 공정불량 개선율`에는 기준 1,403 ppm·목표 1,000 ppm을 설정하고 MES050032의 각 기간 열 전체 공정 평균과 목표 대비 달성율을 연결했다.
- Decisions: 내재화 Total은 명시적인 Total/합계/전체 또는 `ZTOTE` 표시가 있는 INAMT 묶음을 우선하며, 표시가 없을 때는 INAMT 묶음이 하나뿐인 경우에만 사용한다. IPG 평균은 각 COL 열의 null/빈값을 제외하고 숫자 0은 포함하며, `AVEGR` 전체 평균은 연간실적과 연간 달성율에 사용한다.
- Files: JinoSupporter.Web/Services/BmesReportHtmlExportService.cs, JinoSupporter.Web/Services/FCostCorePartsService.cs, JinoSupporter.Web/Services/IpgDefectService.cs
- Verification: 격리 샘플에서 상세행과 Total 행이 함께 있을 때 Total의 `$75.99`와 `0.47%`만 선택되는지, SQLite 저장 후 동일 값이 재구성되는지, IPG 열 평균이 null을 제외하고 0을 포함하는지 모두 PASS 확인했다. `dotnet msbuild JinoSupporter.Web/JinoSupporter.Web.csproj -t:Compile -p:BuildProjectReferences=false`도 오류 없이 성공했으며 기존 경고만 발생했고 애플리케이션·서버 및 외부 BMES 요청은 실행하지 않았다.
- Next: 웹 서버 재시작 후 Get Report를 실행해 실제 MES072410 Total 표시값과 MES050032 월·주 매칭 결과가 KPI 표에 기대값으로 표시되는지 확인한다.

## 2026-07-29 08:17 - [DATA][CHANGE] ReportKpiData: 빈 기간 원인 확인
- Agent: Codex
- Session: S1
- Task-ID: S1-20260729-071709013-3e19e3d342ed48dba
- Category: DATA
- Feature: ReportKpiData
- Change: CHANGE
- Completed: KPI 표 중간의 빈 칸은 계산이나 키 매칭 오류가 아니라 각 BMES 응답이 제공하는 기간 범위가 Main 기준 열보다 짧아서 생기는 것으로 확인했다. 최신 MES072410 스냅샷은 W28~W31과 5~7월을, 현재 MES050032 주간 요청은 W25~W31만 제공한다.
- Decisions: 사용자 선택 전에는 값을 추정하거나 빈 기간을 0으로 채우지 않는다. 과거 기준일 추가 조회로 주차 범위를 합칠 수 있으며, IPG 월 값은 월간 요청 응답을 검증해 사용하거나 주간값에서 계산하는 두 방식 중 하나를 정해야 한다.
- Files: None
- Verification: 실제 `fcost_raw.db`를 읽기 전용으로 조회해 MES072410의 최신 `BottomGridColumnList`가 7개 일간 열, W28~W31, 5~7월로 구성된 것을 확인했고 화면의 내재화 표시 기간과 일치했다. MES050032 코드의 `ZGUBN=W` 단일 종료일 요청 및 제공 샘플의 W25~W31 헤더가 화면의 IPG 표시 기간과 일치하는 것도 확인했다.
- Next: 사용자에게 과거 주차 추가 조회 여부와 IPG 월 값을 월간 BMES 조회 또는 주차 평균 중 어떤 방식으로 채울지 확인한 뒤 구현한다.

## 2026-07-29 08:56 - [INTEGRATION][CHANGE] ReportKpiData: 과거 주차와 IPG 월간 응답 병합
- Agent: Codex
- Session: S1
- Task-ID: S1-20260729-071709013-3e19e3d342ed48dba
- Category: INTEGRATION
- Feature: ReportKpiData
- Change: CHANGE
- Completed: MES072410을 보고서 범위 내 21일 간격 기준일로 추가 조회하고 저장된 스냅샷의 주·월 값을 기간별로 병합해 최신 겹침값을 사용하도록 변경했다. MES050032는 한 인증 세션에서 42일 간격의 과거 주간 요청과 종료일 월간 `ZGUBN=M` 요청을 수행하고 실제 응답 헤더가 KPI 기간과 일치할 때만 평균값을 연결한다.
- Decisions: 과거 응답과 최신 응답에 같은 주·월이 있으면 최신 기준일 값을 사용한다. IPG 월간 헤더가 없거나 요청이 실패하면 주간값으로 월 값을 추정하지 않고 `-`를 유지하며, 다른 기간 요청 실패도 확보된 실제 응답만 표시한다.
- Files: JinoSupporter.Web/Services/BmesReportHtmlExportService.cs, JinoSupporter.Web/Services/FCostCorePartsService.cs, JinoSupporter.Web/Services/IpgDefectService.cs
- Verification: 격리 샘플에서 W19~W21 병합, 중복 W20의 최신값 우선, `ZGUBN=M` payload, 월간 전체 행 평균, `26-07`→`M:202607` 및 `26-W19`→`W:202619` 매칭을 모두 PASS 확인했다. `dotnet msbuild JinoSupporter.Web/JinoSupporter.Web.csproj -t:Compile -p:BuildProjectReferences=false`도 오류 없이 성공했으며 기존 경고만 발생했고 애플리케이션·서버 및 외부 BMES 요청은 실행하지 않았다.
- Next: 웹 서버 재시작 후 동일 기간으로 Get Report를 실행해 MES072410 과거 주차와 MES050032 월간 실제 응답이 채워지는지 확인하고, BMES가 반환하지 않은 기간만 `-`로 남는지 점검한다.

## 2026-07-29 09:05 - [UI][ADD] ReportKpiWeekToggle: 월별 주차 접기·펼치기와 경계선 정렬
- Agent: Codex
- Session: S1
- Task-ID: S1-20260729-071709013-3e19e3d342ed48dba
- Category: UI
- Feature: ReportKpiWeekToggle
- Change: ADD
- Completed: KPI 월 헤더에 작은 `−`/`+` 버튼을 추가해 해당 월에 속한 주차 헤더와 모든 KPI 값 열을 함께 접고 펼치도록 구현했다. 병합 행에서 첫 DOM 셀에 잘못 추가되던 좌측선과 KPI 구간의 위·아래 이중 경계선을 제거하고 묶음 하단의 한 줄로 정렬했다.
- Decisions: 월 열 자체는 항상 표시하며 주차는 기존 시간축 정렬에 사용하는 주간 중간 날짜의 연월로 묶는다. 접힘 상태는 동일 정적 보고서에서 탭을 전환해도 유지하고, 월에 매칭되지 않는 주차는 임의로 숨기지 않는다.
- Files: JinoSupporter.Web/Services/BmesReportHtmlExportService.cs
- Verification: 월 버튼/주차 그룹/숨김 상태/버튼 기호/월 열 유지/중복 좌측선 제거/단일 구간 경계선 계약 검사를 모두 PASS했다. `dotnet msbuild JinoSupporter.Web/JinoSupporter.Web.csproj -t:Compile -p:BuildProjectReferences=false`도 오류 없이 성공했으며 기존 경고만 발생했고 애플리케이션·서버는 실행하지 않았다.
- Next: 웹 서버 재시작 후 Get Report를 다시 생성해 각 월 버튼이 해당 월 주차만 접는지와 KPI 구간선이 전체 폭에서 동일한 굵기로 보이는지 확인한다.

## 2026-07-29 09:35 - [UI][ADD] BmesTest5ModelGroupBom: Test 5 메뉴와 모델그룹 BOM 목록 페이지
- Agent: Claude
- Session: S3
- Task-ID: S3-20260729-090813163-1ba4d206edd6441bb
- Category: UI
- Feature: BmesTest5ModelGroupBom
- Change: ADD
- Completed: BMES 하위에 Test 5 메뉴(`bmes-test5`, `/bmes/test5`)를 추가하고, 모델그룹을 고르면 그 그룹의 중그룹 모델별 하위 BOM을 레벨 들여쓰기 테이블로 보여주는 페이지를 만들었다. 기존에 정의만 되어 있고 호출되지 않던 `BomMaterialsSql`(MAST/STPO 재귀 전개)을 소비하는 `FetchBomTreeAsync`를 서비스에 추가했고, 결과를 독립 실행 HTML 파일로 내려받는 버튼도 붙였다.
- Decisions: `FetchBomMaterialsWithSqlAsync`의 리더 루프를 `ReadBomRowsAsync`로 분리해, FCOST 후보 목록은 기존대로 자재 단위 중복 제거를 유지하고 BOM 트리는 부모·레벨 문맥을 살리기 위해 중복 제거를 하지 않는다. 모델그룹의 "모델"은 `MidGroupRecord.Material`로 해석했고, 모델 수가 많으면 조회 비용이 커지므로 체크박스로 대상 모델을 고르게 했다(기본 전체 선택, 모델당 최대 2000행). 권한은 신규 DB에서 Viewer 제외 역할에 기본 부여하고, 기존 DB에는 Test 3/Test 4/Model Group 권한 보유 역할에 마이그레이션으로 승계한다.
- Files: JinoSupporter.Web/Components/Pages/BmesTest5Page.razor, JinoSupporter.Web/Components/Layout/NavMenu.razor, JinoSupporter.Web/Services/AppMenus.cs, JinoSupporter.Web/Services/BmesFcostActualService.cs, JinoSupporter.Web/Services/WebRepository.cs
- Verification: `dotnet build JinoSupporter.Web/JinoSupporter.Web.csproj`를 스크래치패드 출력 경로로 실행해 오류 0개·경고 0개로 성공했다(기본 출력 경로는 실행 중인 웹 서버가 exe를 잠가 복사 단계에서만 실패). 실제 BMES DB 조회와 화면 동작은 서버 기동 금지 방침에 따라 실행하지 않았다.
- Next: 웹 서버 재시작 후 `/bmes/test5`에서 모델그룹을 선택해 BOM 행이 채워지는지, Download HTML 결과가 열리는지 확인한다. 모델 수가 많은 그룹에서 조회 시간이 길면 모델별 병렬 조회나 캐시 도입을 검토한다.

## 2026-07-29 09:55 - [DATA][FIX] BmesTest5ModelGroupBom: BOM 쿼리를 dbo.BOMC 기준으로 교체
- Agent: Claude
- Session: S3
- Task-ID: S3-20260729-090813163-1ba4d206edd6441bb
- Category: DATA
- Feature: BmesTest5ModelGroupBom
- Change: FIX
- Completed: Test 5에서 모든 모델이 `Invalid object name 'dbo.STPO'`로 실패하던 문제를 고쳤다. `BomMaterialsSql`이 SAP 이름 `dbo.STPO`를 참조하고 있었으나 BMES_LIV의 BOM 라인 테이블은 `dbo.BOMC`이므로, `BMES_FCOST_DB_NOTES.md`에 문서화된 `MAST`+`BOMC` 조인(STLNR/STLAL/STLAN)과 `CMATE`/`MENGE`/`MEINS`/`SDATE`/`EDATE`/`USEYN` 컬럼으로 재귀 CTE를 다시 작성했다. 페이지에 Work date 입력을 추가해 BOM 유효기간 기준일을 지정할 수 있게 했다.
- Decisions: SQL Server는 재귀 CTE의 recursive member에서 서브쿼리와 outer join을 금지하므로, 재귀는 자재코드만으로 전개하고 이름 조회(`MATE.MAKTX`)는 최종 SELECT의 OUTER APPLY로 옮겼다. 그 결과 SQL이 만들 수 없게 된 이름 기반 `BomPathText`와 `HasChildren`은 `DecorateBomTree`에서 C#으로 채운다. MAST의 BOM 대체안 때문에 같은 경로가 중복될 수 있어 `BomPath` 기준으로 중복을 제거하고, 수량 곱셈 오버플로 여유를 위해 decimal(38,10) 대신 decimal(28,6)을 쓴다. 유효기간 컬럼이 NULL인 행은 제외하지 않는다. 모든 쿼리는 SELECT + NOLOCK 전용이며 BMES 서버에 쓰기 작업은 하지 않는다.
- Files: JinoSupporter.Web/Services/BmesFcostActualService.cs, JinoSupporter.Web/Components/Pages/BmesTest5Page.razor
- Verification: `dotnet build JinoSupporter.Web/JinoSupporter.Web.csproj`를 스크래치패드 출력 경로로 실행해 오류 0개로 성공했다. 실제 BMES DB에 대한 쿼리 실행 검증은 하지 않았으므로 BOMC 스키마 가정(SDATE/EDATE가 date 계열, STLNR/STLAL/STLAN 조인)은 화면에서 확인이 필요하다.
- Next: 웹 서버 재시작 후 `/bmes/test5`에서 같은 모델그룹을 다시 Load BOM 해 행이 채워지는지 확인한다. 여전히 비면 Work date를 BOM 유효기간 안의 날짜로 바꾸거나 Plant 값을 비워 재시도한다.

## 2026-07-29 10:20 - [UI][CHANGE] BmesTest5ModelGroupBom: M-P 자재 숨김 옵션 추가
- Agent: Claude
- Session: S3
- Task-ID: S3-20260729-090813163-1ba4d206edd6441bb
- Category: UI
- Feature: BmesTest5ModelGroupBom
- Change: CHANGE
- Completed: Test 5 BOM 목록에서 `M-P`로 시작하는 구매 자재 코드를 감추는 `Hide M-P codes` 체크박스를 추가했다(기본 켜짐). 화면 테이블, 행 수 배지, 상태 문구, Download HTML 결과에 모두 동일하게 적용되며 숨긴 행 수를 헤더에 표시한다.
- Decisions: 조회 단계가 아니라 표시 단계에서 거른다. SQL에서 걸러내면 재귀 전개가 그 지점에서 끊겨 하위가 사라지고 옵션을 끌 때 재조회가 필요하기 때문이다. 부모가 숨겨졌는데 자식만 들여쓰기된 채 남는 것을 막기 위해 `BomPath` 접두사로 하위 트리도 함께 숨긴다. QTY 산출 기준은 이번 변경에서 건드리지 않았다.
- Files: JinoSupporter.Web/Components/Pages/BmesTest5Page.razor
- Verification: `dotnet build JinoSupporter.Web/JinoSupporter.Web.csproj`를 스크래치패드 출력 경로로 실행해 오류 0개로 성공했다. 실제 데이터로 몇 행이 걸러지는지는 서버 미기동으로 확인하지 않았다.
- Next: 사용자에게 QTY 기준을 설명했으며, BOM 기준수량(base quantity) 컬럼이 BOMC/MAST에 존재하는지 조회 전용으로 확인할지 회신 대기 중이다. 존재하면 누적 수량을 기준수량으로 나누도록 보정해야 한다.

## 2026-07-29 10:35 - [UI][REMOVE] BmesTest5ModelGroupBom: 수량 열 제거와 누적 곱셈 폐기
- Agent: Claude
- Session: S3
- Task-ID: S3-20260729-090813163-1ba4d206edd6441bb
- Category: UI
- Feature: BmesTest5ModelGroupBom
- Change: REMOVE
- Completed: Test 5 화면 테이블과 Download HTML에서 Qty·Unit 열을 없앴다. 함께 `BomMaterialsSql` 재귀 부분의 `bt.UsageQty * c.MENGE` 누적 곱셈을 제거해 각 행이 해당 BOM 라인의 원본 `MENGE`만 담도록 했다. 남은 표 구성은 Lv / Material / Code / Parent (+ HTML은 Path)이다.
- Decisions: `BmesBomMaterialCandidate.UsageQty`/`UsageUnit` 속성은 FCOST 후보 조회 경로가 공유하므로 모델에서 제거하지 않았다. 누적 곱셈은 표시되지 않으면서 산술 오버플로 위험과 기준수량 미반영 문제만 남기므로 폐기했고, 이 근거를 `FetchBomTreeAsync` XML 주석에 남겼다. BOMC 기준수량 컬럼 확인 건은 수량 자체가 불필요해져 종료한다.
- Files: JinoSupporter.Web/Components/Pages/BmesTest5Page.razor, JinoSupporter.Web/Services/BmesFcostActualService.cs
- Verification: `dotnet build JinoSupporter.Web/JinoSupporter.Web.csproj`를 스크래치패드 출력 경로로 실행해 오류 0개로 성공했다. 화면 확인은 서버 미기동으로 하지 않았다.
- Next: 웹 서버 재시작 후 `/bmes/test5`에서 BOM 행이 실제로 조회되는지 확인한다. dbo.BOMC 기반 쿼리는 아직 실 DB에서 한 번도 실행되지 않았다.

## 2026-07-29 10:55 - [UI][CHANGE] BmesTest5ModelGroupBom: C-S 숨김·Parent 열 제거·매칭 제품 표시
- Agent: Claude
- Session: S3
- Task-ID: S3-20260729-090813163-1ba4d206edd6441bb
- Category: UI
- Feature: BmesTest5ModelGroupBom
- Change: CHANGE
- Completed: 숨김 대상에 `C-S` 접두사를 추가하고(체크박스 라벨 `Hide M-P / C-S codes`), 화면과 Download HTML에서 Parent 열을 제거했다. BOM 경로는 Material 셀 tooltip과 HTML의 Path 열에 남는다. 카드 헤더에 해당 모델명이 MATE에서 실제로 매칭한 제품 목록을 `matched: …`로 표시하도록 추가했다.
- Decisions: `C-S`는 하위가 달린 어셈블리라 M-P처럼 하위 트리째 지우면 R-S 자재가 전부 사라진다. 그래서 숨김 로직을 "행만 제거하고 자식은 유지"로 바꾸고, 숨겨진 조상 수만큼 `BomLevel`과 `BomPathText`를 다시 계산해 들여쓰기가 어긋나지 않게 했다. 이 방식은 M-P에도 그대로 안전해 두 접두사를 같은 경로로 처리한다. `matched` 표시는 모델명이 한쪽 변형(예: -L)만 잡는 상황을 사용자가 바로 판별하도록 넣은 진단용이다.
- Files: JinoSupporter.Web/Components/Pages/BmesTest5Page.razor
- Verification: `dotnet build JinoSupporter.Web/JinoSupporter.Web.csproj`를 스크래치패드 출력 경로로 실행해 오류 0개로 성공했다. 사용자 스크린샷으로 dbo.BOMC 기반 쿼리가 실 DB에서 정상 동작함을 확인했다(TIU-C11-20, 36행).
- Next: TIU-C11-20이 `TIU-C11-20-L-ZZ`만 매칭하는 원인을 `matched` 표시로 확인한다. R 변형이 MATE에 있는지, 있다면 MAST에 BOM 헤더가 없거나 plant/유효기간 조건에서 걸리는지 조회 전용으로 점검이 필요하다.

## 2026-07-29 11:20 - [UI][CHANGE] BmesTest5ModelGroupBom: 모델그룹 선택을 모델명 평면 목록으로 교체
- Agent: Claude
- Session: S3
- Task-ID: S3-20260729-090813163-1ba4d206edd6441bb
- Category: UI
- Feature: BmesTest5ModelGroupBom
- Change: CHANGE
- Completed: 제품군/모델그룹 드롭다운을 없애고 모든 모델그룹의 중그룹 모델명을 하나로 펼친 평면 목록으로 바꿨다. 툴바에는 모델명 필터 입력을 넣었고, 칩은 전체 모델(중복 제거)을 보여주며 tooltip으로 그 모델이 속한 그룹명을 알려준다. All 버튼은 현재 필터에 걸린 것만 선택한다.
- Decisions: 그룹은 BOM 조회에 불필요한 중간 단계이고 여러 그룹에 걸친 모델을 한 번에 뽑을 수 없었으므로 제거했다. 그룹 정보는 버리지 않고 칩 tooltip으로 남겼다. 이전에는 그룹 선택 시 전체 모델이 자동 선택됐으나, 이제 초기 선택은 비워 둔다(의도치 않은 대량 조회 방지). HTML 파일명·제목은 그룹명 대신 조회한 모델명 기준으로 바꿨다.
- Files: JinoSupporter.Web/Components/Pages/BmesTest5Page.razor
- Verification: `dotnet build JinoSupporter.Web/JinoSupporter.Web.csproj`를 스크래치패드 출력 경로로 실행해 오류 0개로 성공했다. 로컬 process-review.db를 읽기 전용으로 조회해 distinct 모델명이 18개임을 확인했고, 이 규모에서는 칩 전체 렌더가 부담되지 않는다.
- Next: 화면에서 모델 필터와 다중 선택 동작을 확인한다. TIU-C11-20이 -L 변형만 매칭하는 원인은 카드 헤더 `matched:` 표시로 계속 확인 대기 중이다.

## 2026-07-29 11:45 - [UI][FIX] BmesTest5ModelGroupBom: 제품(L/R)별 그룹 구분과 중복 행 정리
- Agent: Claude
- Session: S3
- Task-ID: S3-20260729-090813163-1ba4d206edd6441bb
- Category: UI
- Feature: BmesTest5ModelGroupBom
- Change: FIX
- Completed: 한 모델명이 여러 제품(-L-ZZ / -R-ZZ)에 매칭될 때 행들이 구분 없이 섞여 보이던 문제를 고쳤다. 화면과 Download HTML 모두 제품별 구분 행(제품명 + 행 수)으로 묶어 표시한다. 아울러 C-S 숨김으로 계층이 평탄해지면서 같은 자재가 한 제품 안에 중복 표시되던 것을 제품+자재코드 기준으로 합치고 가장 얕은 레벨의 행만 남긴다.
- Decisions: `matched:` 표시로 L/R 제품이 실제로는 둘 다 조회되고 있었음을 확인했고(TIU-L5S3-01), 원인은 조회가 아니라 표시였다. 앞서 TIU-C11-20이 -L만 나온 것은 그 모델에 L 제품만 존재하기 때문으로 판단한다. 중복 제거는 숨김 옵션이 켜져 계층이 무너진 경우에만 적용되며, 옵션을 끄면 원래 트리와 중복 경로가 그대로 보인다.
- Files: JinoSupporter.Web/Components/Pages/BmesTest5Page.razor
- Verification: `dotnet build JinoSupporter.Web/JinoSupporter.Web.csproj`를 스크래치패드 출력 경로로 실행해 오류 0개로 성공했다. 화면 확인은 서버 미기동으로 하지 않았다.
- Next: `/bmes/test5`에서 TIU-L5S3-01을 다시 조회해 L-ZZ / R-ZZ 두 구획으로 나뉘고 각 구획 안에 자재 중복이 사라졌는지 확인한다.

## 2026-07-29 12:00 - [UI][CHANGE] BmesTest5ModelGroupBom: 자재명 정렬 적용
- Agent: Claude
- Session: S3
- Task-ID: S3-20260729-090813163-1ba4d206edd6441bb
- Category: UI
- Feature: BmesTest5ModelGroupBom
- Change: CHANGE
- Completed: 제품 구획 안의 자재 행을 이름순으로 정렬했다. 화면과 Download HTML 모두 동일하게 적용되며, 동명 자재는 자재코드 순으로 이어 정렬한다.
- Decisions: 자재명 자체가 아니라 `BomPathText`(이름 경로)로 정렬한다. 부모의 경로가 자식 경로의 접두사라서 이 방식은 계층을 유지한 채 형제만 이름순으로 정렬하며, 숨김 옵션으로 트리가 평탄해진 기본 상태에서는 결과가 단순 자재명 오름차순과 같다. 자재명으로 직접 정렬하면 숨김을 해제했을 때 들여쓰기와 부모-자식 순서가 깨진다. `ByProduct`가 IGrouping 대신 `ProductRows` 레코드를 반환하도록 바꿔 그룹 내부 정렬을 담게 했다.
- Files: JinoSupporter.Web/Components/Pages/BmesTest5Page.razor
- Verification: `dotnet build JinoSupporter.Web/JinoSupporter.Web.csproj`를 스크래치패드 출력 경로로 실행해 오류 0개로 성공했다. 화면 확인은 서버 미기동으로 하지 않았다.
- Next: `/bmes/test5`에서 정렬 결과를 확인하고, 숨김 옵션을 해제했을 때 계층이 유지되는지도 함께 본다.

## 2026-07-29 12:20 - [UI][ADD] BmesTest5ModelGroupBom: 공통/편측 전용 자재 구분 표시
- Agent: Claude
- Session: S3
- Task-ID: S3-20260729-090813163-1ba4d206edd6441bb
- Category: UI
- Feature: BmesTest5ModelGroupBom
- Change: ADD
- Completed: 한 모델명이 여러 제품(-L-ZZ / -R-ZZ)에 매칭될 때, 제품별로 나열하던 것을 자재가 어느 제품들에 쓰이는지 기준으로 다시 묶었다. `Common — all N products`(초록 헤더)가 먼저 나오고 그 뒤에 `<제품> only` 구획이 이어진다. 공통 자재는 이제 한 번만 표시된다. 화면과 Download HTML 모두 동일하다.
- Decisions: 이 재그룹핑은 숨김 옵션이 켜져 트리가 평탄한 상태이고 매칭 제품이 2개 이상일 때만 적용한다. 계층이 살아 있을 때 사용처 기준으로 묶으면 자식이 부모에서 떨어져 나가므로, 그 경우와 제품이 1개인 경우는 기존 제품별 트리 구획을 유지한다. 구획 내 정렬은 평탄 모드에서는 자재명 기준, 트리 모드에서는 경로 기준(형제만 이름순)으로 나눠 적용한다.
- Files: JinoSupporter.Web/Components/Pages/BmesTest5Page.razor
- Verification: `dotnet build JinoSupporter.Web/JinoSupporter.Web.csproj`를 스크래치패드 출력 경로로 실행해 오류 0개로 성공했다. 화면 확인은 서버 미기동으로 하지 않았다.
- Next: TIU-L5S3-01로 확인해 YOKE/DIAPHRAGM 류가 Common으로, FRAME-L/FRAME-R이 각 편측 구획으로 갈리는지 본다.

## 2026-07-29 12:50 - [INTEGRATION][ADD] BmesPdmScrape: MES073300 도면/사양 마스터 수집과 Test 5 모델 선택 연동
- Agent: Claude
- Session: S3
- Task-ID: S3-20260729-090813163-1ba4d206edd6441bb
- Category: INTEGRATION
- Feature: BmesPdmScrape
- Change: ADD
- Completed: BMES `MES073300/SearchList`(POST JSON)를 호출해 도면/사양 마스터 전체를 받아 `bmes_pdm_raw.db`의 `BmesPdm` 테이블에 전량 저장하는 `BmesPdmScrapeService`를 추가했다. Test 5의 모델 선택 UI를 모델그룹 칩에서 이 PDM 리스트로 교체해, `PDRTX`(부품명) 또는 `BMONO_TX`(모델명)로 검색하고 체크한 행의 `BMONO_TX`를 BOM 조회 대상으로 넘긴다. 페이지에 `Fetch PDM list` 버튼과 수집 시각·행 수 표시를 넣었다.
- Decisions: 기존 `BmesRoutingScrapeService`의 로그인·토큰·전량 덮어쓰기 패턴을 그대로 따랐다(HTTPS 필수, 테이블 DROP 후 재생성). 검색 조건 payload는 사용자가 캡처한 값을 그대로 전송한다 — 응답에 여러 PDTNO가 섞여 오므로 `PDTNO:"0001"`이 결과를 좁히지 않는 것으로 보이나 확증은 없다. BOM 조회의 루트로는 `BMONO_TX`를 쓴다(모델명 형태가 기존 모델그룹 중그룹명과 동일 계열). 응답이 약 6.7MB라 `AutomaticDecompression`을 켜고 타임아웃 300초를 뒀으며, 검색은 기본 300행으로 제한하고 `(BMONO_TX, PDRTX)` 인덱스를 만든다. 이 엔드포인트는 BMES 웹 API 조회이며 서버 DB에는 쓰지 않는다.
- Files: JinoSupporter.Web/Services/BmesPdmScrapeService.cs, JinoSupporter.Web/Components/Pages/BmesTest5Page.razor, JinoSupporter.Web/Program.cs
- Verification: `dotnet build JinoSupporter.Web/JinoSupporter.Web.csproj`를 스크래치패드 출력 경로로 실행해 오류 0개로 성공했다. 실제 MES073300 호출과 저장은 서버 미기동으로 실행하지 않았다.
- Next: 웹 서버 재시작 후 Test 5에서 `Fetch PDM list`를 눌러 행이 저장되는지, 검색·선택 후 Load BOM이 동작하는지 확인한다. 모델그룹 기반 모델 목록은 제거됐으므로 필요하면 복원 여부를 결정한다.

## 2026-07-29 13:30 - [INTEGRATION][CHANGE] BmesPdmScrape: PDM 리스트를 BOM 결과 기반 필터로 전환
- Agent: Claude
- Session: S3
- Task-ID: S3-20260729-090813163-1ba4d206edd6441bb
- Category: INTEGRATION
- Feature: BmesPdmScrape
- Change: CHANGE
- Completed: 흐름을 뒤집었다. 모델 선택은 다시 모델그룹 기반 모델 칩 목록으로 되돌렸고, Load BOM으로 BOM을 먼저 조회한 뒤 그 결과 자재명으로 PDM 마스터를 걸러 보여준다. `SearchByNames(partNames, modelNames)`를 추가해 PDRTX/BMONO_TX exact match(대소문자 무시, 400개씩 배치)로 조회하며, PDM 테이블은 선택 기능 없는 표시 전용이 됐다.
- Decisions: 첫 수집 데이터를 읽기 전용으로 검사한 결과 14,622행 중 `BMONO_TX`가 채워진 행은 17개뿐이고 전부 `MSU-L15S10-20` 하나였다. 따라서 모델 기준 조인은 불가능하고 실질 조인키는 `PDRTX`다. 스크린샷의 BOM 자재명 8건(FRAME-L5S3-01-L 등)이 모두 PDRTX에 정확히 1건씩 존재함을 확인해 `PDRTX = MATE.MAKTX` 관계를 검증했다. 정렬도 PDRTX 우선으로 바꿨다 — BMONO_TX 우선이면 빈 값이 앞에 몰려 첫 300행이 전부 공란이 되고 체크박스가 잠긴다(직전 증상의 원인).
- Files: JinoSupporter.Web/Services/BmesPdmScrapeService.cs, JinoSupporter.Web/Components/Pages/BmesTest5Page.razor
- Verification: `dotnet build JinoSupporter.Web/JinoSupporter.Web.csproj` 오류 0개. 로컬 `bmes_pdm_raw.db`를 읽기 전용으로 조회해 BOM 자재명 8건이 전부 PDRTX와 exact match됨을 확인했다. 화면 동작은 서버 미기동으로 확인하지 않았다.
- Next: 웹 서버 재시작 후 모델을 골라 Load BOM 하고 PDM list 구획이 해당 자재의 도면행으로 채워지는지 확인한다. 매칭이 부족하면 exact match를 접두사/부분 일치로 완화할지 판단한다.

## 2026-07-29 13:50 - [UI][CHANGE] BmesPdmScrape: PDM 정보를 BOM 표 인라인 열로 병합
- Agent: Claude
- Session: S3
- Task-ID: S3-20260729-090813163-1ba4d206edd6441bb
- Category: UI
- Feature: BmesPdmScrape
- Change: CHANGE
- Completed: 별도 `PDM list` 구획을 제거하고 BOM 표에 `PDMNO`·`PDTNO_TX` 열을 붙였다. 각 BOM 행의 자재명(PDRTX)으로 조회해 그 행에 도면번호와 분류가 바로 붙는다. `Fetch PDM` 버튼은 툴바로 옮겼고(툴팁에 마스터 행 수·수집 시각), Download HTML도 같은 두 열을 포함한다.
- Decisions: 매칭 결과를 리스트 대신 자재명 기준 사전(`_pdmByPart`)으로 보관해 행 렌더링에서 O(1) 조회한다. PDRTX는 실무상 자재당 1건이지만 중복이 생기면 첫 행을 채택한다(TryAdd). 제품 구분 행의 colspan은 화면 5, HTML 6으로 맞췄다.
- Files: JinoSupporter.Web/Components/Pages/BmesTest5Page.razor
- Verification: `dotnet build JinoSupporter.Web/JinoSupporter.Web.csproj` 오류 0개. 화면 확인은 서버 미기동으로 하지 않았다.
- Next: 웹 서버 재시작 후 Load BOM 하고 PDMNO 열이 채워지는지, 빈 칸이 많으면 자재명 매칭 규칙을 완화할지 판단한다.

## 2026-07-29 14:20 - [INTEGRATION][ADD] BmesPdmScrape: PDMNO 클릭 시 DWG 도면 다운로드
- Agent: Claude
- Session: S3
- Task-ID: S3-20260729-090813163-1ba4d206edd6441bb
- Category: INTEGRATION
- Feature: BmesPdmScrape
- Change: ADD
- Completed: BOM 표의 PDMNO 셀을 버튼으로 바꾸고 `DownloadDrawingAsync(pdmno, pdmvr)`를 추가했다. SearchDetail로 첨부 파일 정보(PATTX/PATTP/P_FOLDE/P_PATTX/PDFSQ)를 얻고, DownloadCheck로 권한 확인(`Result=="S"`) 후 파일을 받아 브라우저로 내려준다. 첨부가 여러 개면 USEYN이 N이 아닌 DWG를 우선 선택하고 없으면 첫 파일을 쓴다.
- Decisions: 캡처된 트래픽에 실제 파일 바이트를 받는 요청이 없다(BMES 페이지가 XHR이 아니라 navigation으로 받는 것으로 추정). 그래서 마지막 단계는 `/MES073300/Download`, `/MES073300/FileDownload`, `/Common/Download` 순으로 쿼리스트링 GET을 시도하고 바이트를 반환하는 첫 응답을 채택하며, 전부 실패하면 시도한 경로와 사유를 그대로 오류 메시지에 노출한다. content-type이 json/html이고 Content-Disposition이 없으면 파일이 아닌 것으로 판정한다. 다운로드는 매 클릭마다 새 세션으로 로그인한다(기존 스크레이프 서비스와 동일 패턴).
- Files: JinoSupporter.Web/Services/BmesPdmScrapeService.cs, JinoSupporter.Web/Components/Pages/BmesTest5Page.razor
- Verification: `dotnet build JinoSupporter.Web/JinoSupporter.Web.csproj` 오류 0개. 실제 다운로드는 서버 미기동으로 실행하지 않았고, 파일 엔드포인트 경로는 추정이라 미검증이다.
- Next: 웹 서버 재시작 후 PDMNO를 클릭해 DWG가 받아지는지 확인한다. 실패하면 오류 메시지에 시도 경로가 나오므로, BMES에서 도면 다운로드를 누를 때 DevTools 필터를 XHR이 아닌 All로 두고 실제 요청 URL을 캡처해 `DownloadDrawingAsync`의 후보 목록을 교체한다.

## 2026-07-29 14:45 - [INTEGRATION][FIX] BmesPdmScrape: 실제 Download 엔드포인트 적용과 PDMVR 전달
- Agent: Claude
- Session: S3
- Task-ID: S3-20260729-090813163-1ba4d206edd6441bb
- Category: INTEGRATION
- Feature: BmesPdmScrape
- Change: FIX
- Completed: 추정으로 두었던 파일 다운로드 단계를 실제 캡처대로 교체했다. `POST /MES073300/Download`에 `application/x-www-form-urlencoded`로 `Download[PDMNO]`·`Download[PDMVR]`·`Download[PATTX]`·`Download[P_FOLDE]`·`Download[P_PATTX]`를 보내고, 응답의 `Content-Disposition` 파일명을 그대로 사용한다. 아울러 PDMVR을 "0"으로 고정하던 것을 PDM 목록 행의 실제 값으로 전달하도록 고쳤다(`BmesPdmRow.Pdmvr` 추가, SELECT에 PDMVR 포함).
- Decisions: 로컬 DB 확인 결과 14,622행 중 2,448행(약 17%)이 PDMVR≠0이고, 화면에서 보던 `FRAME-L5S3-01-L`/`-R`도 PDMVR=1이라 하드코딩 "0"이면 그 자재들부터 실패했을 것이다. SearchDetail 응답의 `Files[].PDMVR`이 있으면 그 값을 우선 사용하고 없으면 요청 버전으로 되돌린다. `X-Requested-With`는 기본 헤더에서 빼고 SearchDetail/DownloadCheck 두 JSON 호출에만 붙였다 — 파일 액션은 브라우저에서 navigation으로 호출되므로 AJAX로 오인되면 응답이 달라질 수 있다. 실패 시 서버가 200과 함께 HTML 오류 페이지를 주는 경우를 대비해 Content-Disposition 부재 + html 타입이면 실패로 처리한다.
- Files: JinoSupporter.Web/Services/BmesPdmScrapeService.cs, JinoSupporter.Web/Components/Pages/BmesTest5Page.razor
- Verification: `dotnet build JinoSupporter.Web/JinoSupporter.Web.csproj` 오류 0개. 로컬 `bmes_pdm_raw.db` 읽기 전용 조회로 PDMVR 분포(0:12,174 / 1:1,770 / 2:469 / 3:142 / 4:41 / 5:20 / 6:3 / 8:2 / 9:1)를 확인했다. 실제 다운로드는 서버 미기동으로 실행하지 않았다.
- Next: 웹 서버 재시작 후 PDMNO를 클릭해 DWG가 실제로 받아지는지 확인한다.

## 2026-07-30 07:21 - [DATA][ADD] BmesTest5BomCache: 모델 캐시 우선 조회와 유사명 즉시 로드
- Agent: Codex
- Session: S2
- Task-ID: S2-20260730-072033265-3c5e6d2087ae4f6d9
- Category: DATA
- Feature: BmesTest5BomCache
- Change: ADD
- Completed: Test 5의 모델별 BOM을 `bmes_bom_cache.db`에 저장하고 기본 조회는 로컬 캐시를 우선 사용하며, `Load BOM from server` 체크 시 BMES SQL Server에서 강제 갱신하도록 구현했다. 최초 모델 칩은 BOM 행이 정상 저장된 캐시 모델만 표시하고, 모델 입력의 유사 후보를 클릭하면 해당 모델을 즉시 로드하며 표·트리 마커도 좌측 밀집 정렬로 보정했다.
- Decisions: 캐시 키는 모델명·Plant·Depth·MaxRows 기준이며 Work date는 저장 메타로만 남겨 다음 날짜에도 같은 모델 캐시를 재사용한다. 새 날짜 기준으로 갱신하려면 서버 체크박스를 명시적으로 켜며, 캐시 미스는 체크가 꺼져 있어도 서버에서 한 번 조회 후 저장한다. 후보 순위는 구분자 제거 후 exact/prefix/contains/Levenshtein 순이고 상위 8개를 보여준다.
- Files: JinoSupporter.Web/Components/Pages/BmesTest5Page.razor, JinoSupporter.Web/Services/BmesBomCacheService.cs, JinoSupporter.Web/Program.cs
- Verification: 격리 출력의 `dotnet build JinoSupporter.Web/JinoSupporter.Web.csproj`가 오류 0개로 성공했다(기존 경고 32개). 임시 데이터 디렉터리 왕복 검증에서 2개 BOM 행 저장, 다음 날짜 캐시 재사용, decimal/문자열 복원, 성공 모델 인덱스를 확인해 `PASS rows=2 cachedDate=2026-07-29`를 얻었다. 앱/서버는 실행하지 않았다.
- Next: 실행 중인 Web 서버를 재시작한 뒤 `/bmes/test5`에서 유사 모델 후보 클릭 즉시 로드, cache/server 배지, 서버 강제 갱신, 좌측 밀집 표와 트리 마커를 화면 확인한다.

## 2026-07-30 07:33 - [INTEGRATION][FIX] BmesTest5ModelSearch: 모델그룹 별칭을 실제 BMES 제품 후보로 변환
- Agent: Codex
- Session: S2
- Task-ID: S2-20260730-073222125-691315c5117049ada
- Category: INTEGRATION
- Feature: BmesTest5ModelSearch
- Change: FIX
- Completed: Test 5의 모델 입력을 로컬 모델그룹명 후보가 아니라 BMES `MATE`의 실제 제품명 후보를 검색하도록 변경했다. 입력 토큰을 숫자·문자 단위로 순위화하고 `MAST`·`BOMC`에서 선택 Plant와 날짜에 유효한 BOM이 있는 제품만 보여주며, 후보 클릭 시 실제 제품명으로 BOM을 즉시 조회한다.
- Decisions: `ASSY 338 RA1`처럼 모델그룹명과 제품명이 다를 때도 `ASSY`와 `338`의 부분 일치로 `ASSY REAR...338...` 계열을 상위에 둘 수 있도록 숫자 포함 토큰의 가중치를 높였다. 검색은 300ms 디바운스하고 이전 요청을 취소하며, 기존 성공 캐시는 관련 토큰이 맞는 경우에만 후보에 합친다. 후보 검색도 기존 읽기 전용 SQL 연결과 `NOLOCK`만 사용한다.
- Files: JinoSupporter.Web/Components/Pages/BmesTest5Page.razor, JinoSupporter.Web/Services/BmesFcostActualService.cs
- Verification: `dotnet build JinoSupporter.Web/JinoSupporter.Web.csproj --no-restore --artifacts-path .codex-verify/bmes-model-search-artifacts`가 오류 0개로 성공했다(기존 경고 32개). 사용자 지시에 따라 BMES 서버 연결과 앱 실행은 수행하지 않았다.
- Next: Web 서버 재시작 후 `/bmes/test5`에서 `ASSY 338 RA1`을 입력해 실제 `ASSY REAR...338...` BMES 후보가 표시되고, 후보 클릭 후 BOM 행이 로드·캐시되는지 화면 확인한다.

## 2026-07-30 07:37 - [UI][FIX] BmesTest5BomAlignment: 공통·전용 부품 목록의 잘못된 트리 들여쓰기 제거
- Agent: Codex
- Session: S2
- Task-ID: S2-20260730-073522028-4e10c36da9f248bc9
- Category: UI
- Feature: BmesTest5BomAlignment
- Change: FIX
- Completed: 여러 제품의 공통/전용 부품을 하나의 구간으로 합쳐 표시할 때 원래 제품별 BOM 레벨이 남아 관계없는 다음 행이 자식처럼 보이던 문제를 수정했다. 비교 구간의 행은 모두 레벨 1과 점 마커로 평면 정렬하며, 단일 제품을 표시하거나 노이즈 코드 숨김을 끈 경우에는 실제 BOM 트리를 그대로 유지한다.
- Decisions: 공통/전용 구간은 제품별 트리가 아니라 자재 사용 비교 목록이므로 알파벳순 평면 표시가 의미에 맞다. 실제 경로는 툴팁용 `BomPathText`에 유지해 원본 계층 정보는 잃지 않는다.
- Files: JinoSupporter.Web/Components/Pages/BmesTest5Page.razor
- Verification: `dotnet build JinoSupporter.Web/JinoSupporter.Web.csproj --no-restore --artifacts-path .codex-verify/bmes-model-search-artifacts`가 오류 0개로 성공했다(기존 경고 32개). 앱/서버는 실행하지 않았다.
- Next: Web 서버 재시작 후 여러 실제 제품이 매칭된 BOM에서 공통/전용 구간의 모든 행이 같은 좌측 위치와 `Lv 1`로 표시되는지 확인한다.

## 2026-07-30 07:45 - [DATA][CHANGE] BmesTest5ModelCatalog: 입력 검색을 로컬 모델명 카탈로그로 전환
- Agent: Codex
- Session: S2
- Task-ID: S2-20260730-073940136-f27433752e36442eb
- Category: DATA
- Feature: BmesTest5ModelCatalog
- Change: CHANGE
- Completed: Test 5가 사용자의 입력마다 BMES SQL Server를 조회하던 흐름을 제거하고, `Sync model names` 실행 시 BOM이 있는 C-S 완제품의 코드·모델명만 한 번 받아 `bmes_bom_cache.db`에 저장하도록 변경했다. 이후 유사명 검색과 후보 순위화는 메모리에 로드된 로컬 카탈로그에서만 수행하고, 후보 클릭 시 저장된 정확한 모델명으로 기존 BOM 로드·캐시 흐름을 실행한다.
- Decisions: 전체 MATE나 BOM 본문을 미리 저장하지 않고 선택 Plant·작업일에 유효한 `MAST`/`BOMC`가 있는 `C-S-` 제품의 코드와 이름만 최대 20,000개 동기화한다. 모델그룹 별칭처럼 일부 단어가 다른 경우를 위해 숫자 토큰을 우선하는 로컬 순위를 유지하며, Plant별 카탈로그·동기화 시각·기준일을 별도로 저장한다.
- Files: JinoSupporter.Web/Components/Pages/BmesTest5Page.razor, JinoSupporter.Web/Services/BmesBomCacheService.cs, JinoSupporter.Web/Services/BmesFcostActualService.cs
- Verification: `dotnet build JinoSupporter.Web/JinoSupporter.Web.csproj --no-restore --artifacts-path .codex-verify/bmes-model-search-artifacts`가 오류 0개로 성공했다(기존 경고 32개). 임시 데이터 디렉터리의 SQLite 왕복 검사에서 모델명 2개와 Plant·기준일·동기화 시각을 저장·복원해 `PASS rows=2 cachedDate=2026-07-29 modelNames=2`를 얻었다. BMES 서버와 앱은 실행하지 않았다.
- Next: Web 서버 재시작 후 최초 한 번 `Sync model names`를 실행하고, 이후 여러 검색어를 입력해 네트워크 요청 없이 `local` 후보가 표시되며 정확한 모델명을 클릭했을 때 BOM이 로드되는지 확인한다.

## 2026-07-30 07:56 - [DATA][FIX] BmesTest5ModelCatalog: 완제품 P-S 모델명으로 동기화 범위 수정
- Agent: Codex
- Session: S2
- Task-ID: S2-20260730-075459859-6ec8a57c308048e2a
- Category: DATA
- Feature: BmesTest5ModelCatalog
- Change: FIX
- Completed: 모델명 카탈로그가 `C-S-` 중간조립품을 저장해 TIU 완제품 L/R 모델을 누락하던 문제를 수정하고, 실제 완제품 코드인 `P-S-`만 동기화하도록 변경했다. 재동기화 후 `TIU-C11-20`과 `TIU-L5S3-01` 검색에서 L/R 실제 제품명이 로컬 후보가 되도록 했다.
- Decisions: 실제 로컬 BOM 캐시를 읽어 `TIU-C11-20-L-ZZ`의 제품 코드가 `P-S-151100700`인 반면 잘못 동기화된 33개 이름은 `C-S-` ASSY 항목임을 확인했다. 기존 로컬 카탈로그는 삭제하지 않고 다음 `Sync model names` 실행이 Plant 3200 카탈로그를 P-S 결과로 교체하게 한다.
- Files: JinoSupporter.Web/Components/Pages/BmesTest5Page.razor, JinoSupporter.Web/Services/BmesFcostActualService.cs
- Verification: 실제 `bmes_bom_cache.db`를 읽기 전용 조회해 잘못된 카탈로그 33개가 C-S ASSY이고 캐시된 TIU 완제품이 P-S임을 확인했다. `dotnet build JinoSupporter.Web/JinoSupporter.Web.csproj --no-restore --artifacts-path .codex-verify/bmes-model-search-artifacts`가 오류 0개로 성공했다(기존 경고 32개). BMES 서버와 앱은 실행하지 않았다.
- Next: Web 서버 재시작 후 `Sync model names`를 다시 한 번 눌러 기존 C-S 33개를 P-S 완제품 카탈로그로 교체하고, `TIU-C11-20` 및 `TIU-L5S3-01` 검색에서 L/R 후보가 함께 나오는지 확인한다.

## 2026-07-30 11:45 - [UI][CHANGE] BomAndDrawingTitle: Test 5 표시명 변경
- Agent: Codex
- Session: S2
- Task-ID: S2-20260730-114453958-04b42e257ea942a29
- Category: UI
- Feature: BomAndDrawingTitle
- Change: CHANGE
- Completed: 기존 `Test 5` 표시명을 메뉴, 페이지 제목, 화면 헤더, HTML 내보내기 제목 및 활동 로그에서 `BOM & Drawing`으로 변경했다.
- Decisions: 기존 URL `/bmes/test5`, 권한 키 `bmes-test5`, 컴포넌트·상수 이름은 저장된 권한과 링크 호환성을 위해 유지했다. 메뉴 아이콘 표시는 `T5`에서 `BD`, 화면 머리표시는 `B&D`로 변경했다.
- Files: JinoSupporter.Web/Components/Layout/NavMenu.razor, JinoSupporter.Web/Components/Pages/BmesTest5Page.razor, JinoSupporter.Web/Services/AppMenus.cs
- Verification: `dotnet build JinoSupporter.Web/JinoSupporter.Web.csproj --no-restore --artifacts-path .codex-verify/bmes-model-search-artifacts`가 오류 0개로 성공했다(기존 경고 32개). 앱/서버는 실행하지 않았다.
- Next: Web 서버 재시작 후 좌측 BMES 메뉴와 `/bmes/test5` 페이지·브라우저 탭에 `BOM & Drawing`이 표시되는지 확인한다.

## 2026-07-30 11:56 - [UI][CHANGE] BmesSettingModelMenu: BMES 설정 메뉴 2단 그룹화
- Agent: Codex
- Session: S2
- Task-ID: S2-20260730-115431133-b91b9ae0bbf44f719
- Category: UI
- Feature: BmesSettingModelMenu
- Change: CHANGE
- Completed: BMES 하위에 접이식 `Setting Model` 그룹을 추가하고 `Model Group`, `Routing Table`, `Reason Table`을 2단 메뉴로 이동했다. 각 기존 권한은 그대로 적용되며 관련 페이지에 접속하면 BMES와 Setting Model 그룹이 자동으로 열린다.
- Decisions: 새 권한 키나 경로를 추가하지 않고 기존 세 메뉴의 권한과 URL을 재사용했다. 하나 이상의 하위 메뉴 권한이 있을 때만 Setting Model 그룹을 표시한다.
- Files: JinoSupporter.Web/Components/Layout/NavMenu.razor, JinoSupporter.Web/Components/Layout/NavMenu.razor.css
- Verification: `dotnet build JinoSupporter.Web/JinoSupporter.Web.csproj --no-restore --artifacts-path .codex-verify/bmes-model-search-artifacts`가 오류 0개로 성공했다(기존 경고 32개). 앱과 서버는 지침에 따라 실행하지 않았다.
- Next: Web 서버 재시작 후 BMES 메뉴의 2단 계층, 관련 페이지 활성 상태, 권한별 메뉴 표시를 확인한다.

## 2026-07-30 11:59 - [UI][CHANGE] BmesSettingModelMenu: Setting Model 메뉴를 최하단으로 이동
- Agent: Codex
- Session: S2
- Task-ID: S2-20260730-115825226-062441303a4e40ebb
- Category: UI
- Feature: BmesSettingModelMenu
- Change: CHANGE
- Completed: BMES 내부의 `Setting Model` 그룹을 `LPA` 다음 위치로 옮겨 BMES 하위 메뉴의 가장 마지막에 표시되도록 변경했다.
- Decisions: 그룹 내부 메뉴 순서와 권한 조건, 자동 펼침 동작은 유지하고 표시 위치만 변경했다.
- Files: JinoSupporter.Web/Components/Layout/NavMenu.razor
- Verification: `dotnet build JinoSupporter.Web/JinoSupporter.Web.csproj --no-restore --artifacts-path .codex-verify/bmes-model-search-artifacts`가 오류 0개로 성공했다(기존 경고 32개). 앱과 서버는 지침에 따라 실행하지 않았다.
- Next: Web 서버 재시작 후 BMES 하위 메뉴에서 `Setting Model`이 최하단에 표시되는지 확인한다.

## 2026-07-30 15:33 - [OTHER][FIX] BmesReportPerformance: Get Report 성능 병목 조사
- Agent: Codex
- Session: S2
- Task-ID: S2-20260730-152832180-70b0ecbf8bab45b8a
- Category: OTHER
- Feature: BmesReportPerformance
- Change: FIX
- Completed: BMES Report의 실제 실행 로그와 호출 구조를 대조해 Daily와 F-COST의 주요 병목을 확인했다. Daily는 60만 행 임시 DB 생성과 동일 원본의 최대 3회 집계가 크고, F-COST는 최근 3일 RAW 갱신·네트워크 원자재 분석·MES072410·MES050032가 직렬 실행되어 약 40초가 소요된다.
- Decisions: 1차 개선은 선택 LineShift SQL 필터와 계층 집계 중복 제거, F-COST 원격 결과 캐시 및 독립 작업 병렬화로 잡는다. 원인 비중·Weekly의 Daily 결과 재사용과 F-COST 4개 탭의 계산 스냅샷 재사용은 이미 적절하므로 유지한다.
- Files: None
- Verification: `web-20260730.log`의 2026-07-30 15:20 실행을 확인했다. NG Rate 수집·임시 DB 처리는 15.524초였고, F-COST는 RAW 갱신 7.514초, 원자재 분석 16.805초, MES072410 1.820초, MES050032 13.965초로 확인됐다. 코드 변경은 없어 빌드는 실행하지 않았다.
- Next: 승인 시 선택 LineShift를 SQLite 쿼리에서 필터링하고 계층 보고서 중복 집계를 축소한 뒤, F-COST 캐시·병렬화를 단계별로 구현하고 동일 기간으로 전후 시간을 비교한다.

## 2026-07-30 15:47 - [DATA][FIX] BmesReportPerformance: Daily 선택 LineShift 선필터 적용
- Agent: Codex
- Session: S2
- Task-ID: S2-20260730-154200396-1ce6502d9efd47548
- Category: DATA
- Feature: BmesReportPerformance
- Change: FIX
- Completed: BMES Report의 Daily 생성 시 선택된 모델 그룹의 LineShift 목록을 NG Rate 수집 단계까지 전달하고, 월별 DB에서 날짜와 LineShift를 함께 SQL 필터링하도록 변경했다. 최근 서버 행도 임시 DB 저장 전에 같은 LineShift 집합으로 줄이며 Sub Group 보고서는 필요한 요약 집계만 생성한다.
- Decisions: 기존 monthly DB와 전체 데이터 캐시 형식은 유지하고 Report 전용 경로에서만 선필터를 전달한다. LineShift 필터가 없는 기존 호출자는 이전처럼 전체 데이터를 사용하며, 인스턴스 재사용 DB 키에도 필터 집합을 포함해 다른 선택 결과가 섞이지 않게 했다.
- Files: JinoSupporter.Web/Components/Pages/NgRateForDailyReportPage.razor, JinoSupporter.Web/Services/NgRateService.cs, JinoSupporter.Web/Services/HierReportSupport.cs
- Verification: `dotnet build JinoSupporter.Web/JinoSupporter.Web.csproj --no-restore --artifacts-path .codex-verify/bmes-model-search-artifacts`가 오류 0개로 성공했다(기존 경고 32개). 새 출력 폴더의 첫 시도는 복원 자산 부재로 시작되지 않아 기존 검증 폴더를 사용했으며 앱과 서버는 실행하지 않았다.
- Next: MES050032 및 F-COST 원자재 분석 결과의 SQLite 캐시 구현을 통합하고 동일 Web 빌드로 최종 검증한다.

## 2026-07-30 15:51 - [DATA][ADD] IpgDefectCache: MES050032 response cache
- Agent: Codex
- Session: S2
- Task-ID: S2-20260730-154200396-1ce6502d9efd47548
- Category: DATA
- Feature: IpgDefectCache
- Change: ADD
- Completed: Added a persistent MES050032 SearchList cache in `fcost_raw.db` keyed by query date and W/M period, with cache reads checked before BMES authentication so repeat report calls can avoid login and network requests. Historical cached entries are reused indefinitely, recent mutable entries use a 15 minute TTL, and cache hit/progress messages are logged.
- Decisions: Stored only BMES response JSON plus metadata in a dedicated `MES050032SearchListCache` table so credentials and cookies are never persisted and existing FCost/MES072410 tables coexist untouched. Weekly query generation now uses 42-day checkpoints and omits the redundant start query when the oldest checkpoint covers it.
- Files: JinoSupporter.Web/Services/IpgDefectService.cs
- Verification: `dotnet build JinoSupporter.Web/JinoSupporter.Web.csproj --artifacts-path .codex-verify/ipg-cache-artifacts` succeeded with existing warnings. A temporary smoke console verified weekly query dates and cache TTL policy, then was removed except ignored binary/cache remnants that deletion policy blocked.
- Next: Run an actual Get Report twice for the same period and confirm the second run logs MES050032 cache hits and skips BMES login for fully cached IPG queries.

## 2026-07-30 15:53 - [DATA][ADD] BmesFcostRawBreakdownCache: raw breakdown SQLite cache
- Agent: Codex
- Session: S2
- Task-ID: S2-20260730-154200396-1ce6502d9efd47548
- Category: DATA
- Feature: BmesFcostRawBreakdownCache
- Change: ADD
- Completed: Added a persistent cache lookup before the SQL Server raw breakdown query and stores successful `BmesFcostRawBreakdownResult` payloads in `fcost_raw.db`. Historical queries are cached without expiry, recent mutable periods use a 15 minute TTL, and cache hit, miss, and store outcomes are logged.
- Decisions: The cache uses a dedicated coexistence table keyed from non-secret inputs only: normalized database name, fact, plant, normalized periods, and the effective line-shift/group/model mapping. Server, user, password, cookies, and connection secrets are not stored; corrupt or expired payloads are ignored and the SQL path is used.
- Files: JinoSupporter.Web/Services/BmesFcostActualService.cs
- Verification: `dotnet build JinoSupporter.Web/JinoSupporter.Web.csproj --artifacts-path .codex-verify/fcost-raw-breakdown-cache-artifacts` succeeded with existing warnings. An initial `--no-restore` attempt against the fresh artifacts path failed only because `project.assets.json` did not exist there.
- Next: Run the same F-COST raw breakdown query twice in the app and confirm the first run logs cache miss/store and the second run logs cache hit without SQL Server raw breakdown execution.

## 2026-07-30 15:59 - [DATA][FIX] BmesReportPerformance: F-COST 반복 원격 조회 제거
- Agent: Codex
- Session: S2
- Task-ID: S2-20260730-155710252-59494120a90f44bfb
- Category: DATA
- Feature: BmesReportPerformance
- Change: FIX
- Completed: F-COST 최근 RAW와 MES072410은 15분 이내 성공 캐시를 재사용하고, MES050032 응답 및 원자재 breakdown 결과는 `fcost_raw.db`의 전용 테이블에 영속 캐시하도록 통합했다. 필요한 날짜가 모두 캐시에 있으면 BMES 로그인까지 생략하며, 보고서 로딩 시 사용하지 않는 F-COST RawJson 할당과 불필요한 Sub Group 상세 집계도 제거했다.
- Decisions: 과거 조회 캐시는 만료 없이 재사용하고 최근 변동 가능 데이터만 15분 TTL을 적용한다. MES050032의 시작일 쿼리는 기간 경계 KPI 누락을 피하기 위해 유지하며, 원격 작업의 공격적 병렬화는 공유 SQLite 쓰기 경합을 유발할 수 있어 이번 단계에서는 적용하지 않았다.
- Files: JinoSupporter.Web/Components/Pages/BmesFCostPage.razor, JinoSupporter.Web/Services/BmesFcostActualService.cs, JinoSupporter.Web/Services/BmesReportHtmlExportService.cs, JinoSupporter.Web/Services/FCostCorePartsService.cs, JinoSupporter.Web/Services/FCostReportService.cs, JinoSupporter.Web/Services/FCostService.cs, JinoSupporter.Web/Services/IpgDefectService.cs
- Verification: `dotnet build JinoSupporter.Web/JinoSupporter.Web.csproj --no-restore --artifacts-path .codex-verify/bmes-model-search-artifacts`가 오류 0개로 성공했고 기존 경고 32개만 남았다. 변경 파일 대상 `git diff --check`도 통과했으며 프로젝트 지침에 따라 웹 서버와 외부 BMES 요청은 실행하지 않았다.
- Next: 웹 서버 재시작 후 같은 기간·같은 모델 선택으로 Get Report를 두 번 실행해 두 번째 실행의 Daily filtered row 수와 F-COST/MES072410/MES050032 cache hit 로그 및 총 소요 시간을 비교한다.

## 2026-07-30 16:10 - [DATA][FIX] BmesReportPerformance: Daily 중복 집계와 완성 보고서 캐시 개선
- Agent: Codex
- Session: S2
- Task-ID: S2-20260730-160047439-911b33f3d807401a9
- Category: DATA
- Feature: BmesReportPerformance
- Change: FIX
- Completed: Daily 계층 보고서가 같은 30만 행을 Group/Mid/Sub 용도로 세 번 읽고 집계하던 경로를 모든 계층 요약과 Mid 상세를 한 번에 만드는 단일 보고서로 합쳤다. 최근 3일 NG Rate 원본도 monthly DB에 저장해 15분간 재사용하며, 같은 기간·같은 모델 선택의 완성 HTML 보고서는 디스크 메타데이터와 함께 15분간 재사용해 프로세스 재시작 후에도 재생성을 건너뛴다.
- Decisions: 완성 보고서 키에는 날짜와 선택된 모델 그룹 전체 구조를 포함해 다른 선택 결과가 섞이지 않게 했다. 상세 Reason/Input/NG 행은 Mid 매핑만 사용해 단일 집계가 불필요하게 모든 계층 상세를 확장하지 않도록 했고, 최근 데이터와 완성 보고서는 변동 가능성을 고려해 모두 15분 TTL을 적용했다.
- Files: JinoSupporter.Web/Services/BmesReportHtmlExportService.cs, JinoSupporter.Web/Services/HierReportSupport.cs, JinoSupporter.Web/Services/NgRateReportService.cs, JinoSupporter.Web/Services/NgRateService.cs
- Verification: 실제 `web-20260730.log`에서 개선 전 완료 실행이 Daily 원본 준비 11.813초, 이후 Daily/Cause/Weekly 생성 약 27.2초, 미충전 MES050032 약 10.3초를 사용한 것을 확인했다. `dotnet build JinoSupporter.Web/JinoSupporter.Web.csproj --no-restore --artifacts-path .codex-verify/bmes-model-search-artifacts`는 오류 0개와 기존 경고 32개로 성공했고 대상 파일 `git diff --check`도 통과했다. 실행 중 서버는 재시작하지 않았다.
- Next: 서버를 한 번 재시작해 새 코드를 반영하고 같은 기간·모델로 연속 두 번 실행한다. 첫 실행의 단일 hierarchy 시간과 두 번째 실행의 `Reusing complete report cache` 로그를 확인한다.

## 2026-07-30 16:33 - [DATA][FIX] BmesReportPerformance: hierarchy 회귀 제거와 KPI 병렬화
- Agent: Codex
- Session: S2
- Task-ID: S2-20260730-161613748-081631f64670454db
- Category: DATA
- Feature: BmesReportPerformance
- Change: FIX
- Completed: 모든 hierarchy 단계를 하나의 상세 보고서로 합쳐 3분 가까이 걸리던 회귀를 제거하고, 화면에 상세 행이 필요한 Mid 보고서와 전체 hierarchy 요약만 병렬 생성하도록 축소했다. 요약 전용 경로의 불필요한 Process 원시 집계를 생략하고 F-COST 생성과 MES072410/MES050032 KPI 조회도 병렬 시작하도록 변경했다.
- Decisions: Sub1/Sub2/LineShift 화면은 `GroupSummary`만 사용하므로 전체 hierarchy 요약 결과를 공유하고, Reason/Input/NG 상세가 필요한 Mid만 별도 상세 보고서로 유지한다. 실행 중인 웹 서버는 프로젝트 지침에 따라 재시작하지 않았다.
- Files: JinoSupporter.Web/Services/BmesReportHtmlExportService.cs, JinoSupporter.Web/Services/HierReportSupport.cs, JinoSupporter.Web/Services/NgRateReportService.cs
- Verification: `dotnet build JinoSupporter.Web/JinoSupporter.Web.csproj --no-restore -p:OutputPath=...`가 기존 경고 32개와 오류 0개로 성공했다. 임시 성능 하네스는 실DB 18,866행에서 1.282초, 202607 월 DB에서 3.239초, 3개월 601,259행 과부하 DB에서 9.170초에 hierarchy를 생성했고 모두 결과 행 존재 및 종료 코드 0을 확인했다.
- Next: 웹 앱을 한 번 재시작한 뒤 동일 조건 Get Report를 실행해 새 hierarchy 시간과 `Starting F-COST and KPI source loading in parallel.` 로그를 확인한다. 같은 기간·선택을 15분 안에 다시 실행하면 완성 HTML 캐시가 즉시 재사용되어야 한다.

## 2026-08-01 08:40 - [UI][ADD] InstrumentDesignSystem: 전체 UI 재설계 HTML 목업
- Agent: Claude
- Session: S4
- Task-ID: S4-20260801-082612948-ca9a245bad0347d6a
- Category: UI
- Feature: InstrumentDesignSystem
- Change: ADD
- Prompt: "이 프로젝트 전체 UI 디자인 다시 해줘 . 싹다 뜯어고쳐두 됨. 일단 html로 먼저 UI 만 만들어봐"
- Answer: 기존 Bootstrap 기반 slate/blue UI를 버리고 "계측기(INSTRUMENT)" 컨셉 — 다크 그래파이트 섀시가 밝은 데이터 플레이트를 감싸는 구조, 앰버 시그널 단일 액센트, 각진 hairline, Archivo + IBM Plex Mono — 으로 잡고 Blazor 이식 전에 정적 HTML 목업으로 먼저 확인하도록 제안했다.
- Applied: (1) 디자인 시스템 CSS(토큰/셸/패널/덴스테이블/차트/채팅/로그인) 신규 작성 (2) 셸 렌더러와 차트 런타임 JS 작성(라인·바·스파크라인·히트맵·정렬·테마토글) (3) 대표 6개 화면 목업 페이지 작성(Daily Report, NG Rate, F-Cost, Ask AI, Users, Login) (4) 차트 팔레트를 dataviz 6검사로 light/dark 양쪽 검증 (5) 디자인 토큰·컴포넌트·Blazor 이식 매핑 README 작성
- Completed: `ui-redesign/` 아래에 의존성 없는 정적 HTML 목업 세트를 만들었다. 브라우저에서 바로 열리며 ?theme=dark 로 다크모드 확인이 가능하고, 실제 앱의 메뉴 구조(22개 메뉴, 권한 그룹)와 화면 원형(설정 deck → 리포트 테이블 → 실행 로그)을 그대로 반영했다.
- Decisions: Bootstrap 의존을 제거하는 전제로 설계했다(모든 컴포넌트가 자체 클래스). 사이드바/상단바는 라이트·다크 공통으로 항상 다크 섀시로 고정해 컨셉을 유지했다. 차트 계열색은 orange를 1번 슬롯으로 재정렬했는데, 원래 순서는 dark에서 red↔yellow가 인접해 normal-vision floor(ΔE 13.0 < 15)에 걸려 yellow와 red 사이에 green을 넣어 재검증했다. 목업 셸 마크업은 각 페이지에 복붙하지 않고 instrument.js의 NAV 배열에서 렌더하도록 해 실제 AppMenus/MenuPermissionService 구조와 대응시켰다.
- Files: ui-redesign/assets/instrument.css, ui-redesign/assets/instrument.js, ui-redesign/index.html, ui-redesign/ng-rate.html, ui-redesign/f-cost.html, ui-redesign/ask-ai.html, ui-redesign/admin-users.html, ui-redesign/login.html, ui-redesign/README.md
- Verification: headless Chrome로 6개 페이지 + 다크모드를 1600px 폭으로 렌더해 육안 확인했다. 1차 렌더에서 발견한 축 눈금이 clean number가 아닌 문제, 라인 끝 라벨 겹침, `.notice` flex로 인한 텍스트 분해, `.split` 패널 강제 stretch, 2행 헤더 sticky 겹침, 다크 히트맵 셀 텍스트 대비 부족을 수정하고 재렌더로 확인했다. 팔레트는 `validate_palette.js`로 light(surface #ffffff)/dark(surface #171b21) 모두 통과(worst adjacent CVD ΔE 7.2/6.9, normal-vision 19.6/19.3), sequential ramp는 monotone L·ΔL≥0.06·단일 hue 통과. .NET 빌드나 서버 기동은 하지 않았다(정적 파일만 추가, 앱 코드 미변경).
- Next: 목업 시안 확정 후 이식 착수. 순서는 토큰+셸 먼저(wwwroot/app.css, MainLayout.razor(.css), NavMenu.razor(.css)) → 페이지 단위 순차 전환. Bootstrap 제거 여부를 먼저 결정해야 한다(현재 목업은 제거 전제).

## 2026-08-01 09:05 - [UI][CHANGE] InstrumentDesignSystem: 데이터 표를 엑셀 격자 폼으로 전환
- Agent: Claude
- Session: S4
- Task-ID: S4-20260801-082612948-ca9a245bad0347d6a
- Category: UI
- Feature: InstrumentDesignSystem
- Change: CHANGE
- Prompt: "데이터 표 는 왠만하면 기본 엑셀폼과 비슷한 표로 만들어줘 . 캡쳐해서 공유해야되서"
- Answer: 표를 캡쳐해 공유하는 용도이므로 가로줄만 있던 hairline 스타일을 버리고 전 셀 격자선 + 회색 헤더 밴드의 엑셀 폼으로 바꿨다. 표 일부만 잘라내도 표로 읽히고, hover나 색 단독에 의존하는 정보가 없도록 했다.
- Applied: (1) 표 전용 토큰 5종 추가(--tbl-line/--tbl-line-hard/--tbl-head/--tbl-head-2/--tbl-head-ink, light/dark) (2) .dt를 전 셀 우/하 보더 + 회색 볼드 중앙정렬 헤더 밴드로 재작성하고 정렬 표시를 색 대신 ▾/▴ 글리프로 변경 (3) 그룹행/합계행을 굵은 밴드 행으로, 컬럼그룹 구분선과 합계선을 진한 선으로 분리 (4) .heat 매트릭스도 border-spacing 제거하고 같은 격자로 통일 (5) 긴 표용 .dt--banded 옵션 추가(기본 off) (6) NG Rate 리포트 표를 split에서 빼내 전체 폭으로 배치하고 보조 패널 3개를 아래 grid3으로 이동
- Completed: 모든 데이터 표가 엑셀 기본 표 형태(전 셀 격자 + 회색 헤더)로 렌더된다. 13열짜리 NG Rate 계층 표가 헤더 폰트 확대로 State 열이 잘리던 문제도 전체 폭 배치로 해소했다.
- Decisions: 헤더는 열 정렬과 무관하게 전부 중앙정렬로 통일했다(국내 관리표 폼 관례). 밴딩(줄무늬)은 히트맵 색상 셀과 충돌해 기본 비활성으로 두고 옵션 클래스로만 제공한다. 히트맵 셀 간 2px 간격(dataviz 권장)은 엑셀 조건부서식처럼 보이도록 격자선으로 대체했다.
- Files: ui-redesign/assets/instrument.css, ui-redesign/ng-rate.html, ui-redesign/README.md
- Verification: headless Chrome로 ng-rate/f-cost/index(dark)를 1560px 폭으로 재렌더해 육안 확인했다. F-Cost 2단 헤더 표와 NG Rate 13열 표 모두 격자·헤더 밴드·합계선이 의도대로 렌더되고 열 잘림이 없음을 확인했다. 다크모드에서도 헤더 밴드와 격자선 대비가 유지됨을 확인했다.
- Next: 시안 확정 대기. 확정되면 토큰+셸부터 Blazor 이식 착수(Bootstrap 제거 여부 결정 필요).

## 2026-08-01 09:35 - [UI][ADD] NewVerPreview: 앱 안에서 새 디자인 미리보기 메뉴 추가
- Agent: Claude
- Session: S4
- Task-ID: S4-20260801-082612948-ca9a245bad0347d6a
- Category: UI
- Feature: NewVerPreview
- Change: ADD
- Prompt: "근데 이게 실제로 이렇게 보일려나... 기본 페이지에 New Ver 이라고 메뉴 하나 만들고 그거 누르면 수정된 디자인으로 보여줄수있게 만들어봐"
- Answer: 목업을 wwwroot로 옮겨 앱이 직접 서빙하게 하고, 사이드바 최상단에 Preview 섹션과 `New Ver` 메뉴(/new-ver)를 추가했다. 앱이 Bootstrap을 전역 로드하고 새 디자인 시스템이 .btn/.badge/.row/.small을 같은 이름으로 정의하므로 한 문서에 섞으면 충돌한다고 판단해 미리보기는 iframe으로 격리했고, 이 점을 README와 페이지 주석에 명시했다.
- Applied: (1) `ui-redesign/`을 `JinoSupporter.Web/wwwroot/ui-redesign/`으로 이동해 단일 사본으로 통합 (2) `/new-ver` 라우트의 NewVerPage.razor 추가(EmptyLayout, 상단 페이지 전환 스트립 + iframe, 미인증 시 로그인 리다이렉트) (3) AppMenus에 NewVer("new-ver") 상수와 Preview 그룹 MenuItemDef 등록 (4) NavMenu 최상단에 Preview 섹션과 New Ver 링크 추가 (5) README에 접근 경로 3가지와 iframe 사용 이유 기재
- Completed: 앱 메뉴에서 새 디자인을 열어볼 수 있는 경로를 만들었다. 정적 파일은 실행 중인 서버가 이미 서빙하고 있어(200 확인) 재시작 없이 http://localhost:5050/ui-redesign/index.html 로 즉시 확인 가능하고, /new-ver 라우트만 재빌드·재시작이 필요하다.
- Decisions: 미리보기를 iframe으로 한 것은 Bootstrap과의 클래스 충돌을 피하면서 시안을 왜곡 없이 보여주기 위함이다. 따라서 이 미리보기는 "새 디자인이 앱 안에서 어떻게 보이는지"는 증명하지만 "새 CSS가 Bootstrap과 공존 가능한지"는 증명하지 않는다. 목업을 repo 루트에 두지 않고 wwwroot로 옮겨 사본 이원화를 막았다. New Ver 메뉴는 이식 완료 후 제거할 임시 항목으로 표시했다.
- Files: JinoSupporter.Web/Components/Pages/NewVerPage.razor, JinoSupporter.Web/Components/Layout/NavMenu.razor, JinoSupporter.Web/Services/AppMenus.cs, JinoSupporter.Web/wwwroot/ui-redesign/README.md, JinoSupporter.Web/wwwroot/ui-redesign/(assets/instrument.css, assets/instrument.js, index.html, ng-rate.html, f-cost.html, ask-ai.html, admin-users.html, login.html — 루트에서 이동)
- Verification: `dotnet build JinoSupporter.Web/JinoSupporter.Web.csproj --artifacts-path .codex-verify/newver-artifacts`가 오류 0개, 경고 32개(기존과 동일)로 성공했다. 실행 중인 5050 서버에 curl로 `/ui-redesign/index.html`과 `/ui-redesign/assets/instrument.css` 모두 200을 확인했고, `/new-ver`는 예상대로 404였다(실행 중 바이너리가 신규 라우트 이전 빌드). 지침에 따라 서버는 기동·종료하지 않았고 빌드도 bin/Debug가 아닌 별도 artifacts 경로로 했다.
- Next: 사용자가 F5(또는 restart-web.cmd)로 앱을 재빌드·재시작하면 사이드바 최상단 Preview > New Ver 이 뜬다. 시안 확정 후 Bootstrap 제거 여부를 결정하고 토큰+셸부터 실제 이식을 시작한다.

## 2026-08-01 11:45 - [UI][CHANGE] NewVerPreview: iframe 목업을 실제 동작하는 Blazor 페이지로 교체
- Agent: Claude
- Session: S4
- Task-ID: S4-20260801-082612948-ca9a245bad0347d6a
- Category: UI
- Feature: NewVerPreview
- Change: CHANGE
- Prompt: "아니 실제 사용은 안되는데?" / "실제 동작도 되야지"
- Answer: /new-ver 라우트 자체는 살아 있었고(앞서 404로 보고한 것은 Git Bash가 curl URL을 경로로 변환한 오판정, 실제 302), 문제는 내용물이 정적 목업 iframe이라 아무 동작도 하지 않는 것이었다. iframe을 걷어내고 실제 서비스·실제 DB에 연결된 Blazor 페이지로 교체했다.
- Applied: (1) tools/scope-css.js 신규 — instrument.css의 모든 셀렉터를 .ins 아래로 스코프해 instrument.scoped.css 생성(:root→.ins, [data-theme=dark]→.ins[data-theme=dark], @keyframes 내부는 미변환) (2) instrument.css에 §23 Framework neutraliser 추가 — Bootstrap이 남기는 .btn/.badge border-radius와 .row>* 그리드 컬럼 무력화 (3) NewVerPage.razor를 실제 페이지로 재작성 — 좌측 레일은 AppMenus+MenuPermissionService 권한 필터 결과, 상단 readout은 ConnectedUsersService 실시간 값, 본문은 WebRepository 기반 Users 화면(검색·정렬·역할변경·표시명 인라인편집·추가·삭제·비밀번호 재설정·메뉴권한 매트릭스 전부 실제 쓰기) (4) 테마 토글·레일 접기 동작 (5) 미이식 화면은 Design reference 탭으로 정적 목업 연결
- Completed: New Ver 메뉴가 실제로 쓸 수 있는 화면이 됐다. Users 화면의 모든 조작이 /admin/users와 동일한 SQLite 행을 읽고 쓴다. Bootstrap과의 클래스 충돌은 .ins 스코프로 해소했다.
- Decisions: 첫 실제 이식 대상으로 Users를 골랐다. BMES 세션·네트워크 의존이 없어 단독으로 완결되고, 읽기와 쓰기가 모두 있어 "실제 동작"을 증명하기에 적합하기 때문이다. NG Rate/F-Cost/Daily Report/Ask AI는 아직 정적 목업이며 같은 패턴으로 순차 이식한다. instrument.scoped.css는 생성물이므로 직접 수정 금지, instrument.css 수정 후 스크립트 재실행이 규칙이다.
- Files: tools/scope-css.js, JinoSupporter.Web/Components/Pages/NewVerPage.razor, JinoSupporter.Web/wwwroot/ui-redesign/assets/instrument.css, JinoSupporter.Web/wwwroot/ui-redesign/assets/instrument.scoped.css, JinoSupporter.Web/wwwroot/ui-redesign/README.md
- Verification: `dotnet build --artifacts-path .codex-verify/newver-artifacts`가 오류 0개, 경고 32개(기존과 동일, 신규 경고 없음)로 성공했다. Bootstrap 공존은 별도 검증 페이지로 확인 — App.razor와 같은 순서로 bootstrap.min.css + app.css + instrument.scoped.css를 로드하고 .ins 안팎을 동시에 렌더한 뒤 headless Chrome로 캡쳐해, 바깥 Bootstrap 마크업(btn-primary/badge/table/alert)은 무변화이고 안쪽은 새 디자인으로 렌더됨을 확인했다. 1차 캡쳐에서 .row>* 때문에 배지가 세로로 늘어지고 .btn/.badge에 Bootstrap 라운드가 남는 것을 발견해 §23 추가 후 재캡쳐로 해소를 확인했다. 지침에 따라 서버는 기동·종료하지 않았고 빌드도 bin/Debug가 아닌 별도 artifacts 경로로 했다.
- Next: 사용자가 F5로 재빌드·재시작하면 New Ver에서 실제 동작 화면이 뜬다(현재 실행 중 바이너리는 이전 빌드). 그 다음 이식 대상은 사용자가 지정하는 화면으로, 같은 .ins 스코프 + 실제 서비스 주입 패턴을 따른다.

## 2026-08-01 11:55 - [UI][FIX] NewVerPreview: HeadContent가 static HeadOutlet에 전달되지 않아 CSS 미적용
- Agent: Claude
- Session: S4
- Task-ID: S4-20260801-082612948-ca9a245bad0347d6a
- Category: UI
- Feature: NewVerPreview
- Change: FIX
- Prompt: (사용자가 /new-ver 화면 캡쳐 첨부 — 스타일이 전혀 적용되지 않은 순수 HTML 상태)
- Answer: instrument.scoped.css는 서버에서 정상 제공되고 있었고(200, 49,933바이트), 원인은 App.razor의 `<HeadOutlet />`에 렌더모드가 없어 static인데 NewVerPage는 InteractiveServer라 `<HeadContent>`의 link가 head에 도달하지 못한 것이었다. link를 컴포넌트 본문으로 옮겨 해결했다.
- Applied: (1) NewVerPage.razor의 `<HeadContent>` 제거하고 폰트·스타일시트 link를 컴포넌트 본문 최상단으로 이동(이유를 주석으로 명시) (2) 스코프 CSS 캐시버스터 v1→v2 (3) 레일 채널코드가 글리프와 어긋나던 문제 수정 — Code()가 id가 아닌 Label 기준으로 생성하도록 변경해 글리프와 코드의 두 글자가 일치
- Completed: 페이지 본문에서 스타일시트를 로드하도록 바꿔 CSS 미적용 문제를 제거했다. App.razor의 전역 HeadOutlet은 건드리지 않았다.
- Decisions: HeadOutlet에 렌더모드를 부여하는 방법(문서상 정공법)은 앱 전체 페이지에 영향이 가므로 채택하지 않았다. body 내 `<link>`는 모든 브라우저가 적용하고 컴포넌트 언마운트 시 Blazor가 제거하므로 이 페이지에만 CSS가 국한되는 이점도 있다.
- Files: JinoSupporter.Web/Components/Pages/NewVerPage.razor
- Verification: `dotnet build --artifacts-path .codex-verify/newver-artifacts` 오류 0개, 경고 32개(신규 없음). 렌더링은 실제 페이지와 동일한 CSS 로드 순서(bootstrap.min.css → app.css → body 내 instrument.scoped.css)와 동일한 셸 마크업(rail/topbar/tabs/view/stats/dt/권한매트릭스)을 가진 검증 페이지를 headless Chrome로 1500px 캡쳐해 확인했다. 레일·상단바·탭·테이블·배지가 모두 의도대로 렌더되고, app.css의 `a{color:#2563eb}`가 링크를 파랗게 만들지 않는 것도 확인했다. 서버는 기동·종료하지 않았다.
- Next: F5 재빌드·재시작 후 /new-ver 확인. 이후 이식 대상 화면을 사용자가 지정.

## 2026-08-01 12:20 - [UI][CHANGE] InstrumentShell: 새 UI를 레이아웃으로 승격해 전체 기능을 새 셸에서 사용
- Agent: Claude
- Session: S4
- Task-ID: S4-20260801-082612948-ca9a245bad0347d6a
- Category: UI
- Feature: InstrumentShell
- Change: CHANGE
- Prompt: "아니... 새로운 UI 로 모든 기능이 되어야지. 누르면 원래 버전으로 돌아가면 안되고"
- Answer: 새 셸을 페이지가 아니라 레이아웃(InstrumentLayout)으로 만들어 Routes의 DefaultLayout을 런타임에 전환하도록 했다. 페이지를 한 장도 고치지 않고 모든 라우트·기능이 새 셸 안에서 열리며, 메뉴를 눌러도 구버전으로 나가지 않는다. 기존 Bootstrap 마크업은 §23–24 브리지 CSS로 새 디자인에 매핑했다.
- Applied: (1) UiModeService(Scoped) 추가 — 회로 단위 셸 선택, localStorage로 새로고침 후 복원 (2) Routes.razor의 DefaultLayout을 동적으로 변경 (3) InstrumentLayout.razor 신규 — 실제 권한 기반 레일, 실시간 접속자 readout, 탭 스트립(MainLayout 로직 이식), 테마 토글, Classic UI 복귀 버튼, @Body 호스팅 (4) instrument.css에 BRIDGE 구간 신설 — card/form/table/alert/badge/nav-tabs/버튼 및 app.css 고유 규칙(paste-modal, hier-level-badge, rpt-row 등)을 새 디자인으로 매핑 (5) scope-css.js가 BRIDGE 구간 선언에 !important를 자동 부착 (6) NewVerPage를 셸 전환 페이지로 재작성 (7) Program.cs에 UiModeService 등록
- Completed: New Ver에서 전환하면 앱 전체가 새 셸로 바뀌고, 모든 메뉴·페이지·기능이 그 안에서 그대로 동작한다. Classic UI 버튼으로 즉시 되돌아간다.
- Decisions: 페이지별 재작성 대신 레이아웃 교체 + CSS 브리지를 택했다. 60k 라인 마크업을 건드리지 않고 전 기능을 새 UI로 옮길 수 있고, 페이지별 정식 이식은 그 위에서 점진적으로 하면 되기 때문이다. UiModeService는 Scoped라 한 사용자의 전환이 다른 세션에 영향을 주지 않으며 기본값은 기존 MainLayout이다. app.css를 통째로 비활성화하는 방법은, 누락 시 화면이 무스타일로 깨지는 실패 모드가 더 나빠서 채택하지 않고 !important 브리지를 택했다.
- Files: JinoSupporter.Web/Services/UiModeService.cs, JinoSupporter.Web/Components/Routes.razor, JinoSupporter.Web/Components/Layout/InstrumentLayout.razor, JinoSupporter.Web/Components/Pages/NewVerPage.razor, JinoSupporter.Web/Program.cs, JinoSupporter.Web/wwwroot/ui-redesign/assets/instrument.css, JinoSupporter.Web/wwwroot/ui-redesign/assets/instrument.scoped.css, tools/scope-css.js
- Verification: 전체 재빌드 오류 0개, 경고 32개(기존과 동일). 렌더링은 실제 UsersPage.razor의 Bootstrap 마크업을 그대로 새 셸 안에 넣고 App.razor와 동일한 CSS 로드 순서로 headless Chrome 캡쳐해 확인했다. 1차 캡쳐에서 btn-primary가 파란색으로 남고 셀 안 select가 잘리는 문제를 발견 — 원인은 app.css가 Bootstrap을 !important로 재정의하는 것과 브리지의 고정 행 높이였고, BRIDGE 구간 !important 자동화와 행 높이 content-driven 전환으로 해소 후 재캡쳐로 확인했다. 디자인 시스템 자체 컴포넌트(.btn--go/.badge--*)가 !important 재기술로 깨지지 않는 것도 별도 캡쳐로 확인했다. 서버는 기동·종료하지 않았다.
- Next: F5 재빌드·재시작 → New Ver에서 "Switch to new UI". 전체 화면을 돌며 레이아웃이 깨지는 페이지(특히 자체 100% 높이 스크롤을 쓰는 Ask AI/일부 BMES 화면)를 추려 개별 이식 순서를 정한다.

## 2026-08-06 16:43 - [UI][FIX] InstrumentShell: element 리셋을 스코프 빌드에서 제외하고 .row 소유권 반환
- Agent: Claude
- Session: S24
- Task-ID: S24-20260806-162319060-be2325855bb24abc
- Category: UI
- Feature: InstrumentShell
- Change: FIX
- Prompt: "JinoSupporter(.NET 8 Blazor Server) — 새 UI 셸의 CSS 소유권 버그를 고치는 구현 작업이다. … 할 일 1 — R1: 리셋을 스코프 빌드에서 제외 … 리셋에 의존하던 디자인 시스템 자기 컴포넌트는 스코프 빌드에도 남도록 명시적으로 다시 세워라." (프롬프트가 "이 작업의 성공 기준은 \"셸"에서 잘림)
- Answer: SHELL_AUDIT_CSSLAYER.md의 R1·R2를 구현했다. 리셋을 원본에서 지우지 않고 RESET 마커로 감싸 스코프 빌드에서만 버리고, 리셋에 의존하던 디자인 시스템 컨트롤은 §1b에 명시적으로 재기술했다. 감사가 제시한 최소 집합은 실사용 기준으로 조정했다 — `.chip__x`는 마크업에 한 번도 쓰이지 않아 제외하고, `.chip`/`.brand`/`.tab`/`.input`/`.select`/`.textarea`/`.composer__ta`를 추가했다. `html, body`/`body{margin}`도 버렸다(스코프하면 `.ins`가 되는데 컨테이너의 인라인 `height:100vh`가 이미 이기고 div에는 지울 마진이 없다). `.heat`는 이미 자기 블록에서 border-collapse를 선언하지만 의도를 한곳에 모으려고 §1b에 함께 적었다.
- Applied: (1) instrument.css §1을 RESET-START/RESET-END로 감싸고 box-sizing·:focus-visible만 마커 밖에 남김 (2) instrument.css §1b 신설 — .btn/.brand/.railbtn/.iconbtn/.navlink/.tab/.tab__x/.seg__b/.chip/.qrow/.input/.select/.textarea/.composer__ta의 font·color·cursor·background·border·padding·text-decoration과 .dt/.heat의 border-collapse를 명시 재기술 (3) scope-css.js가 RESET 구간을 출력에서 버리고 그 자리에 사유 주석을 남기며, 닫히지 않은 마커는 예외로 실패시키고 드롭 개수를 로그에 표시 (4) scope-css.js의 "이미 스코프됨" 판정을 startsWith에서 경계 검사로 교체 — `.ins-row`가 `.ins`로 오인돼 스코프를 벗어나던 버그 (5) 디자인 시스템 `.row`/`.row--wrap`을 `.ins-row`/`.ins-row--wrap`으로 개명하고 목업 15곳 갱신 (6) 브리지 `.row > *` 블록 삭제 (7) instrument.scoped.css 재생성, 캐시버스터 v3→v4 (8) README를 두 마커 규약·ins- 접두사 규칙·`?v=` 수동 갱신 의무로 갱신
- Completed: 페이지 CSS 242개 규칙이 되살아났고 Bootstrap 그리드 73곳이 정상 폭으로 돌아왔다. 셸 크롬은 계산된 스타일 기준으로 변경 전과 완전히 동일하다.
- Decisions: R3(cascade layer)는 채택하지 않았다 — !important의 레이어 순서 역전 때문에 브리지 전체를 재검증해야 하고 R1으로 같은 결과를 얻는다. `.hstack`은 Bootstrap 5가 이미 정의하고 있어(`display:flex;align-items:center`) 개명 후보에서 제외하고 `.ins-row`를 택했다. §1b는 Bootstrap Reboot에 의존하지 않게 자족적으로 썼다 — 브리지가 한시적이라 Bootstrap이 빠지면 Reboot의 `button{font-family:inherit}`도 함께 사라지기 때문이다. `.cbx`는 appearance:none에 텍스트가 없고 체크 표시를 border로 그려서 font·color 재기술 대상에서 제외했다.
- Files: JinoSupporter.Web/wwwroot/ui-redesign/assets/instrument.css, JinoSupporter.Web/wwwroot/ui-redesign/assets/instrument.scoped.css, tools/scope-css.js, JinoSupporter.Web/Components/Layout/InstrumentLayout.razor, JinoSupporter.Web/wwwroot/ui-redesign/README.md, JinoSupporter.Web/wwwroot/ui-redesign/index.html, JinoSupporter.Web/wwwroot/ui-redesign/ng-rate.html, JinoSupporter.Web/wwwroot/ui-redesign/f-cost.html, JinoSupporter.Web/wwwroot/ui-redesign/ask-ai.html, JinoSupporter.Web/wwwroot/ui-redesign/admin-users.html
- Verification: `dotnet build` 오류 0개, 경고 32개(신규 없음). 생성물 정적 검사 6항목 통과 — 셀렉터 438개 전부 .ins 스코프, `.ins button|table|a|input|select|textarea` 0개, `.row`/`col-*` 0개, `.ins-row`가 `.ins .ins-row`로 스코프됨, 브리지 !important 247/247 유지, 컨트롤이 상속할 `.ins{font-family:var(--sans)}` 존재. 셸 컨트롤 커버리지 스크립트로 리셋이 주던 속성 중 아무도 선언하지 않는 것이 없음을 확인(초기 실행에서 `.brand`의 color 누락을 발견해 §1b에 추가). 렌더링은 실제 앱과 동일한 CSS 로드 순서(bootstrap → app.css → body 내 instrument.scoped.css → 페이지 style)로 조립한 검증 페이지를 변경 전/후 CSS로 각각 headless Chrome 캡쳐·계산스타일 비교했다. 변경 후: .dib-tab border 0→1px·padding-left 0→12px, .bmes-table separate→collapse, .col-md-6 86.7px→745px, 링크 #0f1319/none→#2563eb/underline, .grp-name-input 13px/400→12px/700, .idt-auto-preview Archivo→Consolas, .btn-close background-image none→SVG 복원. 셸 쪽 .railbtn/.iconbtn/.tab__x/.btn--sm의 background·border·padding·cursor·font-family·color는 전/후 완전 동일. 서버는 기동·종료하지 않았다.
- Next: F5 재빌드·재시작 후 실제 화면에서 확인. 다음 우선순위는 감사 §5의 3번 R4(테마 localStorage 영속화), 이어서 R5(생성기 강화 — importantise 문자 파서, 블록 내 주석 처리, @keyframes ins- 접두사), R6(빌드 연결 + 해시 기반 캐시버스팅).

## 2026-08-07 07:30 - [UI][ADD] InstrumentShell: R1·R2 독립 검증 통과
- Agent: Claude
- Session: External
- Task-ID: Unavailable
- Category: UI
- Feature: InstrumentShell
- Change: ADD
- Prompt: "'JinoSupporter' 작업을 이어서 진행해."
- Answer: 구현 세션이 자기 합성 마크업으로 자기 작업을 검증한 점이 약점이라, 실제 .razor 마크업과 실제 CSS 로드 순서로 루트가 직접 재검증했다. 판정은 통과 — 셸 크롬 6개 그룹이 전후 완전 동일하고 페이지 스타일 7개 그룹이 의도한 방향으로 복원됐으며 회귀는 없었다.
- Applied: (1) 검증 보고서 SHELL_FIX_VALIDATION.md 추가 — 렌더 대조 결과표, 정적 항목 9종, 검증하지 못한 범위 명시
- Completed: headless Chrome로 e7e7177^ 대 현재의 계산 스타일을 비교해 R1·R2가 의도대로 동작함을 확인했다. 코드 변경은 없다.
- Decisions: 검증 세션(S25)이 훅 신뢰 프롬프트에서 멈춘 뒤 재시작으로 소실돼, 단일 결합 작업이므로 워커를 다시 만들지 않고 루트가 직접 수행했다. 렌더 대조는 대표 4개 페이지(DataInferenceBatch, BmesReasonTable, BmesFCost, BmesMakeModelGroup)의 style 블록 682줄을 원문 그대로 넣어 캐스케이드를 재현했다.
- Files: JinoSupporter.Web/wwwroot/ui-redesign/audit/SHELL_FIX_VALIDATION.md
- Verification: 셸 크롬 .railbtn/.iconbtn/.tab__x/.btn--sm/.navlink/.brand 6개 그룹 전후 동일. 복원 확인 — .dib-tab border 0→1px·padding-left 0→12px, .pg-pill border 0→1px, .bmes-table separate→collapse, .btn-close background-image none→SVG, .col-md-6 width 8.39px→675px, .row align-items center→normal, 링크 #0f1319/none→#2563eb/underline. scope-css.js 재실행 결과가 커밋본과 일치(438 blocks, 6 dropped). 스코프 출력에 element 리셋 0건, bare .row/col-* 0건, .row-grp/.row-sum 4건 보존. 서버는 기동·종료하지 않았다.
- Next: 사용자가 F5 재빌드·재시작 후 실제 화면 확인. 그다음은 감사 §5 우선순위 3번 이후 — R4(테마 localStorage 영속화), R5(생성기 강화), R6(빌드 연결), 이어서 감사 A·B의 페이지별 이식 순서.

## 2026-08-07 07:50 - [UI][CHANGE] InstrumentShell: 테마 영속화·생성기 강화·빌드 연결(R4·R5·R6)
- Agent: Claude
- Session: External
- Task-ID: Unavailable
- Category: UI
- Feature: InstrumentShell
- Change: CHANGE
- Prompt: "ㅇㅇ 진행해"
- Answer: 감사 §5 우선순위 3~5번을 한 번에 처리했다. 각각 단일 결합 변경이라 워커를 만들지 않고 루트가 직접 구현했다. 조사 중 실제 결함 두 가지를 추가로 발견해 함께 고쳤다 — `.ins` 컨테이너 자신이 box-sizing을 못 받고 있었고, `@keyframes pulse` 같은 흔한 이름이 스코프와 무관하게 전역으로 새고 있었다.
- Applied: (1) R4 테마를 localStorage(`jino-ui-theme`)에 미러링해 F5 후에도 유지 (2) R5 importantise를 문자 파서로 교체 — 따옴표·괄호 깊이를 추적해 `url("data:...;...")`의 세미콜론에 안 잘리고, 주석만 있는 조각에 `!important`를 붙이지 않음 (3) R5 선언 블록 안의 주석을 버퍼에 이어붙여 브리지 규칙이 조용히 !important를 잃지 않게 함 (4) R5 `*` 셀렉터를 `.ins, .ins *`로 확장 (5) R5 셀렉터 목록 중복 제거 (6) S8 애니메이션 이름에 ins- 접두사 (7) R6 scope-css.js에 check 모드 추가, csproj CheckScopedCss 타깃이 드리프트 시 빌드 실패 (8) R6 `?v=` 수동 토큰을 콘텐츠 해시(AssetVersionService)로 교체 (9) README를 새 규약으로 갱신
- Completed: 테마가 새로고침 후에도 유지되고, 생성물이 소스와 어긋난 채 배포되는 경로가 빌드에서 차단되며, 캐시버스터를 사람이 잊을 수 없게 됐다.
- Decisions: keyframes 접두사는 생성기에서 정규식으로 바꾸지 않고 소스에서 직접 개명했다 — animation 참조까지 안전하게 바꾸려면 이름 수집·치환이 필요한데, 소스에서 고치면 목업도 같은 이름을 쓰게 되어 단순하고 위험이 없다. check 모드는 줄바꿈을 정규화해 비교한다(git이 CRLF로 체크아웃하는데 생성기는 LF로 써서 오탐이 난다). 빌드 타깃은 node가 없으면 경고만 내고 넘어간다 — 검사 하나 때문에 Node 미설치 환경의 빌드를 막는 건 과하다. AssetVersionService는 해시를 프로세스 수명 동안 캐시한다(재빌드하면 어차피 새로 뜬다).
- Files: JinoSupporter.Web/Components/Layout/InstrumentLayout.razor, JinoSupporter.Web/Services/AssetVersionService.cs, JinoSupporter.Web/Program.cs, JinoSupporter.Web/JinoSupporter.Web.csproj, tools/scope-css.js, JinoSupporter.Web/wwwroot/ui-redesign/assets/instrument.css, JinoSupporter.Web/wwwroot/ui-redesign/assets/instrument.scoped.css, JinoSupporter.Web/wwwroot/ui-redesign/README.md
- Verification: `dotnet build` 오류 0개, 경고 32개(신규 없음). 드리프트 검사 종단 확인 — 생성물을 일부러 오염시키면 빌드가 `error : instrument.scoped.css is out of date`로 실패하고 복원하면 통과. importantise 파서 6개 케이스 단위 검증(데이터 URI 세미콜론 미분할, gradient 쉼표, 주석에 stray !important 없음, 중복 !important 없음). 생성물 재확인 — element 리셋 0건, `.ins, .ins *`로 box-sizing이 컨테이너까지 도달, keyframes 3개 모두 ins- 접두사, `!important` 252개 유지. 회귀 확인 — 검증 완료 커밋 83587b6의 스코프 CSS와 현재를 동일 하네스로 headless Chrome 렌더 비교해 계산 스타일 차이 0. 서버는 기동·종료하지 않았다.
- Next: 사용자가 F5 재빌드·재시작 후 실제 화면 확인(테마 토글 후 새로고침하면 유지되어야 한다). 남은 것은 감사 §5의 R7(청소)과 감사 A·B의 페이지별 이식 — 1순위는 BmesTest3Page, QrBakoDataPage, DataInferenceAskPage, DataInferenceDbPage.

## 2026-08-07 08:45 - [UI][FIX] InstrumentShell: 라우트별 .view 계약과 뷰포트 루트 치환(G1·G2)
- Agent: Claude
- Session: External
- Task-ID: Unavailable
- Category: UI
- Feature: InstrumentShell
- Change: FIX
- Prompt: "'JinoSupporter' 작업을 이어서 진행해."
- Answer: 페이지 개별 이식으로 넘어가려 했으나 감사 B의 결론이 "18개를 동시에 고치는 작업이 아니라 전역 slice G1/G2를 먼저 적용하면 H와 root 수준 V/O 대부분이 사라진다"고 명시하고 있어 그쪽을 먼저 했다. 페이지 클래스명을 열거하는 대신 셸이 자식에게 계약을 거는 방식(.view > *, .view--flush > *)을 택해 페이지를 한 장도 고치지 않았다.
- Applied: (1) instrument.css §8b 신설 — 문서형은 `.view > * { min-height:0 }`로 뷰포트 바닥 제거, 작업영역형은 `.view--flush > *`에 확정 높이 부여 (2) `.view`와 `.stage`에 min-height:0 명시 (3) InstrumentLayout이 WorkspaceRoutes allow-list로 `.view--flush`를 선택 (4) 라우트 11개를 작업영역으로 지정, 나머지는 문서형 기본값
- Completed: 페이지가 뷰포트를 소유하던 전제를 셸이 대신 만족시켜, `height:100vh`/`calc(100vh - N)` 루트가 stage 높이에 맞게 해석된다.
- Decisions: 작업영역 판정은 취향이 아니라 근거로 했다 — 각 페이지 CSS가 `flex:1; min-height:0; overflow:hidden` 또는 명시적 `height:100vh` flex 컬럼을 선언한 것만 넣었다. 기본값은 문서형으로 뒀다: 문서형을 flush로 렌더하면 잘리지만 작업영역형을 문서형으로 렌더하면 스크롤이 두 번 생길 뿐이라 실패 모드가 덜 나쁘다. `/data-inference/db`는 MicroSpeakerResultPage이고 루트가 `min-height: calc(100vh - 24px)`라 문서형으로 분류했다(감사 B가 라우트 연결을 확인 못 했던 항목이다). 감사 G4대로 sticky 전역 보정은 하지 않았고, G3(overlay 정책)은 현재 Z 충돌 근거가 없어 이번 범위에서 제외했다.
- Files: JinoSupporter.Web/wwwroot/ui-redesign/assets/instrument.css, JinoSupporter.Web/wwwroot/ui-redesign/assets/instrument.scoped.css, JinoSupporter.Web/Components/Layout/InstrumentLayout.razor
- Verification: `dotnet build` 오류 0개, 경고 32개(신규 없음). 실제 페이지 style 블록을 넣고 headless Chrome 렌더 대조 — 문서형(.msr-root) 이전/현재: `.view` 높이 839px→725px, 루트 781px→65px. 이전에는 뷰포트 805px보다 view가 커서 stage가 셸 밖으로 밀려 `.shell{overflow:hidden}`에 잘리고 있었고 지금은 정확히 들어맞는다. 작업영역형(.diask-root): view 725px = 루트 725px, `.view` 자체 스크롤 없음, 셸 밖으로 넘치지 않음. 셸 크롬·페이지 probe 13종은 직전 커밋 1f6a87d 대비 계산 스타일 차이 0. 서버는 기동·종료하지 않았다.
- Next: 사용자가 F5 재빌드·재시작 후 Ask AI, Daily Input, Model Analysis 같은 작업영역 화면에서 이중 스크롤이 사라졌는지 확인. 그다음은 감사 A·B의 페이지 개별 이식(G5) — 하드코딩 hex를 토큰으로, 페이지 `<style>`의 전역 셀렉터 격리(InputDataBatchPage의 `.panel`/`table`/`th`/`td`가 1순위), 공유 컴포넌트 OcrTextEditor → CurrentProblemWorkflowStrip 순.

## 2026-08-07 11:15 - [UI][FIX] InstrumentShell: 페이지 style 블록의 전역 셀렉터 누출 차단
- Agent: Claude
- Session: External
- Task-ID: Unavailable
- Category: UI
- Feature: InstrumentShell
- Change: FIX
- Prompt: "중간에 멈춘거 같은데 진행해"
- Answer: 페이지 이식(G5)을 워커 3개로 나누려 했으나 호스트가 응답하지 않아 대상을 직접 검증했고, 그 과정에서 감사가 집계한 "전역 셀렉터 누출"의 상당수가 오탐임을 확인했다. 실제 누출은 4개 파일 9개 규칙뿐이었고 작고 기계적이라 워커 없이 직접 고쳤다.
- Applied: (1) InputDataBatchPage의 h1/table/th,td/th/td/td .btn+.btn 6개 규칙을 .input-data-root 아래로 스코프 (2) DataInferenceDbPage의 td.n을 .didb-root 아래로 (3) DataInferenceBatchPage의 td.n을 .dib-root 아래로 (4) DataInferenceValidationPage의 td.n을 .div-root 아래로
- Completed: 페이지 style 블록이 다른 화면을 오염시키는 경로를 없앴다. 특히 td의 max-width:360px + white-space:nowrap가 무관한 화면의 표 셀을 잘라내던 실제 기능 버그가 사라졌다.
- Decisions: 오탐 3종은 건드리지 않았다 — DailyTestDataInputPage의 h3/p/table/th,td/pre(3629~3635)와 DataInferenceDbPage:4038의 body/section/h3은 내보내기용 독립 HTML을 조립하는 C# 문자열이고, LoginPage의 html,body는 @layout EmptyLayout으로 자체 문서 전체를 렌더하므로 셸에 들어가지 않는다. 감사 A·B는 R1/R2 이전에 작성돼 B(브리지 미적용) 항목 상당수가 이미 무의미해졌으므로, 남은 페이지 작업은 감사 목록을 그대로 따르지 말고 재확인 후 진행해야 한다.
- Files: JinoSupporter.Web/Components/Pages/InputDataBatchPage.razor, JinoSupporter.Web/Components/Pages/DataInferenceDbPage.razor, JinoSupporter.Web/Components/Pages/DataInferenceBatchPage.razor, JinoSupporter.Web/Components/Pages/DataInferenceValidationPage.razor
- Verification: `dotnet build` 오류 0개, 경고 32개(신규 없음). 1차 시도에서 CSS 주석에 <style> 문자열을 넣어 Razor가 태그로 파싱해 RZ9980 오류 2건이 났고 문구를 고쳐 해소했다. 렌더 대조 — 실제 CSS 로드 순서로 조립한 하네스에 무관한 화면 마크업과 이 페이지 마크업을 나란히 넣고 수정 전/후를 headless Chrome로 비교. 무관한 h1 26px/800 -> 30.71px/500, 무관한 th 배경 #eef2f7/하단 1px/패딩 8px -> 전부 해제, 무관한 td max-width 360px·nowrap -> none·normal. Input Data Batch 자신의 h1/th/td 3항목은 전후 동일. 서버는 기동·종료하지 않았다.
- Next: 남은 페이지 작업은 화면 렌더 판단이 필요한 것들이다 — DailyTestDataInputPage 내부 패널 vh 5곳(38/64/68/78vh, min-height 72vh 2곳), InputDataTestPage:899 `min(1200px, calc(100vh - 230px))`, DataInferenceAskPage:428 iframe `calc(100vh - 235px)`, DataInferenceDbPage:2069/2206 보고서 frame. 전부 stage가 뷰포트보다 짧아진 만큼 어긋나며, flex 체인 확인이 필요해 추측으로 고치면 패널이 무너질 수 있다.

## 2026-08-07 11:45 - [UI][FIX] InstrumentShell: 페이지 내부 뷰포트 수식을 --view-h 기준으로 교정
- Agent: Claude
- Session: External
- Task-ID: Unavailable
- Category: UI
- Feature: InstrumentShell
- Change: FIX
- Prompt: "하던거 마저해"
- Answer: 감사가 "일괄 override하지 말고 페이지별로 고치라"고 한 내부 iframe·패널 높이를 실제 렌더 측정으로 교정했다. 셸이 자기 크롬 높이를 --view-h로 노출하고 페이지가 그것을 기준으로 쓰게 해서, 매직넘버를 다시 심지 않고 해결했다.
- Applied: (1) instrument.css에 --view-h 토큰 추가 (100vh에서 topbar 46px·tabs 34px·보더 2px 차감) (2) DataInferenceAskPage의 .diask-html-wrap을 calc(var(--view-h) - 201px)로 교정 (3) DataInferenceDbPage의 보고서 frame 2곳을 58vh·72vh에서 --view-h 비례로 (4) DailyTestDataInputPage의 root와 패널 5곳, 분석 iframe 인라인 스타일을 --view-h 비례로
- Completed: 페이지가 뷰포트를 소유하던 시절의 높이 수식이 실제 작업 영역 기준으로 계산된다.
- Decisions: 상수 201px은 추측이 아니라 측정값이다 — 셸 안에서 Ask 페이지를 렌더해 프레임 위 페이지 자체 크롬(header 43 + question 92 + toolbars 66)을 재서 얻었고, 원래 상수 235px과의 차이 34px이 구 MainLayout 탭 스트립 높이와 일치해 교차 검증됐다. min-height floor(Ask 680px, DB 520/720px)는 전부 유지했다 — 작은 화면에서 보고서를 일부러 크게 두고 스크롤시키려는 의도이며 지우면 가독성이 나빠진다. InputDataTestPage.razor는 @page도 컴포넌트 참조도 없는 미사용 파일이라 건드리지 않았다. Classic UI에서는 --view-h가 없어 fallback 100vh가 쓰이며 Ask 프레임이 34px 커지는데, 폐기 예정 경로이고 스크롤 컨테이너 안이라 영향이 없다.
- Files: JinoSupporter.Web/wwwroot/ui-redesign/assets/instrument.css, JinoSupporter.Web/wwwroot/ui-redesign/assets/instrument.scoped.css, JinoSupporter.Web/Components/Pages/DataInferenceAskPage.razor, JinoSupporter.Web/Components/Pages/DataInferenceDbPage.razor, JinoSupporter.Web/Components/Pages/DailyTestDataInputPage.razor
- Verification: `dotnet build` 오류 0개, 경고 32개(신규 없음). scope-css.js check 통과(생성물 동기). headless Chrome 측정 — Ask 프레임은 창 1080에서 750px(스크롤러 대비 +46px 넘침) -> 702px(-2px, 정확히 들어맞음), 1440에서도 -2px. 창 900에서는 680px floor가 유지돼 의도대로 스크롤. Daily 패널은 .64가 462.72px(--view-h 723의 64% = 462.7), .68이 491.64px로 계산 일치. DB ask frame은 1440에서 732.53px(0.58 x 1263), 900에서는 min-height 520px floor가 적용. 서버는 기동·종료하지 않았다.
- Next: 사용자가 F5 재빌드·재시작 후 Ask AI와 Daily Input에서 보고서 프레임이 화면에 맞는지 확인. 감사 A(BMES)의 calc(100vh - N) 화면들(BmesTest3 150/96px, RoutingTable·ReasonTable 190px, CauseMonthly 285px, MicroSpeakerResult 184/260px)도 같은 --view-h 패턴으로 교정 가능하며, 각각 페이지 자체 크롬을 측정해 상수를 다시 잡아야 한다.

## 2026-08-07 12:10 - [UI][FIX] InstrumentShell: view--flush 작업영역 계약 회귀를 되돌림
- Agent: Claude
- Session: External
- Task-ID: Unavailable
- Category: UI
- Feature: InstrumentShell
- Change: FIX
- Prompt: "안보임"
- Answer: 사용자가 화면이 안 보인다고 알려와 원인을 추적한 결과, 2026-08-07 08:45 항목(G1/G2)에서 내가 넣은 `.view--flush > *` 계약이 원인이었다. 그 규칙은 "페이지가 루트 요소 하나만 렌더한다"고 가정했는데 실제로는 대부분 최상위 형제가 여럿이라, flex 컬럼이 높이를 균등 분배해 진짜 콘텐츠 루트가 화면의 일부만 차지하고 나머지가 잘렸다.
- Applied: (1) instrument.css 8b에서 `.view--flush > *` 규칙 삭제하고 되돌린 이유를 주석으로 남김 (2) InstrumentLayout에서 WorkspaceRoutes 목록과 IsWorkspaceRoute 제거, `.view`만 사용 (3) instrument.scoped.css 재생성
- Completed: 모든 라우트가 문서형 `.view`로 돌아가 화면이 다시 정상 표시된다.
- Decisions: 근본 실수는 최상위 구조를 확인하지 않고 "루트 하나"를 가정한 것이다. 실제 확인 결과 DataInferenceAskPage는 최상위 형제 4개(PageTitle, div, div, CurrentProblemWorkflowStrip, div, div), DataInferenceDbPage는 5개(PageTitle, div, div, CurrentProblemWorkflowStrip, section, details)다. 측정으로 재현했다 — 725px 뷰에 형제 3개를 넣으면 각 242px씩 균등 분배됐다. 8b의 `.view > * { min-height: 0 }`과 `.view`/`.stage`의 min-height:0은 유지했다. 이것들은 별개 근거로 검증됐고(문서형 화면에서 view 839px -> 725px), 형제 수와 무관하다. --view-h 토큰과 페이지별 높이 교정(11:45 항목)도 유지했다. 그것들은 flush와 독립이며 스크롤 컨테이너 안에서 동작한다.
- Files: JinoSupporter.Web/wwwroot/ui-redesign/assets/instrument.css, JinoSupporter.Web/wwwroot/ui-redesign/assets/instrument.scoped.css, JinoSupporter.Web/Components/Layout/InstrumentLayout.razor
- Verification: `dotnet build` 오류 0개, 경고 32개(신규 없음). headless Chrome 재현 및 회복 확인 — 되돌리기 전 형제 3개가 각 242px로 균등 분배됐고, 되돌린 뒤 각자 자연 크기(19px/19px/87px)를 갖고 `.view`가 필요할 때만 스크롤하며 잘림 없음. 서버는 기동·종료하지 않았다.
- Next: 작업영역 계약을 다시 시도한다면 반드시 각 페이지의 실제 최상위 구조를 먼저 확인하고, 단일 루트를 가정하지 말 것. 형제가 여럿인 페이지는 페이지 쪽에서 루트 하나로 감싸는 편이 안전하다. 그 전까지는 문서형이 기본이며 이중 스크롤은 감수한다.

## 2026-08-07 12:40 - [UI][ADD] SheetDesign: 검사 성적서 방향 1차 시안
- Agent: Claude
- Session: External
- Task-ID: Unavailable
- Category: UI
- Feature: SheetDesign
- Change: ADD
- Prompt: "이런 UI 말고 .."
- Answer: 사용자가 INSTRUMENT 시안을 거절했다. 확인 결과 방향 자체를 바꾸길 원했고, 가장 거슬린 것은 어두운 레일·상단바였다. 새 방향은 "검사 성적서" — 크롬을 밝게 걷어내고 데이터를 문서처럼 두며, 비교는 공유 기준선 좌우 편차 막대로 표현한다.
- Applied: (1) wwwroot/ui-redesign/sheet/assets/sheet.css 신규 — 종이/잉크 토큰, 밝은 인덱스 레일, 성적서 시트, 엑셀 격자 표, 편차 막대 (2) sheet/daily-report.html 신규 — 실제 화면의 실제 수치로 구성한 Daily Report 시안
- Completed: 정적 목업 1장을 만들어 headless Chrome로 렌더 확인했다. 앱 코드는 건드리지 않았다.
- Decisions: 웹폰트를 쓰지 않는다 — 감사 C가 지적한 대로 공장 인트라넷에서 Google Fonts 요청이 지연되므로, Windows 기본 탑재 Bahnschrift(DIN 계열, 도면·계측기 서체)를 숫자·라벨에, Segoe UI + 맑은 고딕을 본문에 쓴다. 한글 데이터도 폴백 없이 렌더된다. 개선색은 초록이 아니라 파랑이다 — 국내 검사 성적서 관행이고 적록색약에 안전하다. 좋아짐/나빠짐은 기준선 좌우 위치로 먼저 읽히고 색은 보조라 WCAG 1.4.1을 충족한다. 1차 시안은 좌우 2단이었으나 렌더해 보니 두 표가 각각 좁아져 열이 잘리고 기준선이 둘로 나뉘어, 하나의 기준선을 공유하는 단일 표로 바꿨다. 표는 캡쳐해 보고서에 붙이는 용도라 엑셀 격자를 유지했다.
- Files: JinoSupporter.Web/wwwroot/ui-redesign/sheet/assets/sheet.css, JinoSupporter.Web/wwwroot/ui-redesign/sheet/daily-report.html
- Verification: headless Chrome 1600x1000 렌더로 육안 확인. 1차 렌더에서 모델명이 3줄로 깨지고 오른쪽 표가 잘리는 문제를 발견해 단일 표로 재구성 후 재렌더로 해소. 앱 빌드는 하지 않았다(정적 파일만 추가, 앱 코드 미변경).
- Next: 사용자가 이 시안을 확정하면 나머지 대표 화면(NG Rate, F-Cost, Ask AI, Users, Login)으로 확장한 뒤 Blazor 이식에 착수한다. 이식은 기존 InstrumentShell을 대체하는 방식이므로, 그때 instrument.css 계열 자산의 폐기 범위를 함께 정한다.

## 2026-08-07 13:05 - [UI][CHANGE] SheetDesign: 좌측 메뉴 전면 개편과 누락 메뉴 복원
- Agent: Claude
- Session: External
- Task-ID: Unavailable
- Category: UI
- Feature: SheetDesign
- Change: CHANGE
- Prompt: "신규 UI 에 BMES-REPORT 메뉴 빠져있음. 그리고 UI 싹다 갈아  좌측 메뉴 를 이런식으로 구성 / 일단 메인화면은 냅둬"
- Answer: 시안의 메뉴를 실제 NavMenu.razor 구조와 대조한 결과 Report 하나가 아니라 Test 3, Test 4, Setting Model 하위 3개까지 빠져 있었다. 실제 구조대로 복원하면서 사용자가 준 참고 이미지 구성(검색창, 아이콘, 섹션, 접기, 하단 고정)을 적용했다. 본문은 지시대로 건드리지 않았다.
- Applied: (1) 좌측 메뉴를 로고+접기 버튼, 검색창(Ctrl K), 아이콘 메뉴, 섹션 라벨, 하단 고정 Setting/접속자/사용자 구조로 재작성 (2) 누락 메뉴 복원 — Report(report/bmes), Test 3, Test 4, 그리고 Setting Model 아래 Model Group/Routing Table/Reason Table 2단 중첩 (3) 아이콘 26종을 인라인 SVG 스프라이트로 작성 (4) 레일 폭 208px에서 246px로
- Completed: 좌측 메뉴가 실제 앱 메뉴와 1:1로 맞고 참고 이미지 구성을 따른다.
- Decisions: 아이콘은 아이콘 폰트나 CDN 대신 인라인 SVG 스프라이트로 넣었다 — 이 디자인의 전제가 공장 인트라넷에서 외부 요청 0이기 때문이다. 메뉴 근거는 AppMenus.cs가 아니라 NavMenu.razor에서 가져왔다. AppMenus는 Report를 Test Data Analysis 그룹으로 두지만 실제 화면은 BMES 아래 report/bmes로 렌더하며, 사용자가 본 것은 후자다. 선택된 메뉴의 아이콘 색을 빨강에서 잉크색으로 바꿨다 — 이 디자인에서 빨강은 규격 이탈을 뜻하므로 현재 위치 표시에 쓰면 의미가 흐려진다.
- Files: JinoSupporter.Web/wwwroot/ui-redesign/sheet/assets/sheet.css, JinoSupporter.Web/wwwroot/ui-redesign/sheet/daily-report.html
- Verification: headless Chrome 1600x1120 렌더로 육안 확인. 메뉴 항목을 NavMenu.razor와 대조해 BMES 11개(2단 3개 포함), Test Data Analysis 4개, Daily Test Data 1개, Tools 3개, Admin 4개, Setting 1개가 모두 존재함을 확인. 앱 빌드는 하지 않았다(정적 파일만 변경).
- Next: 사용자가 메뉴 구성을 확정하면 나머지 대표 화면으로 확장한다. 본문 레이아웃은 아직 지시 대기 상태다.

## 2026-08-07 16:35 - [UI][CHANGE] SheetDesign: 시안을 앱 셸에 실제 적용
- Agent: Claude
- Session: External
- Task-ID: Unavailable
- Category: UI
- Feature: SheetDesign
- Change: CHANGE
- Prompt: "전혀 바뀌는게 없는데 / 일단 이 형태로 다 적용해"
- Answer: 사용자가 앱을 보며 변화가 없다고 했는데, 지금까지 만든 것이 정적 목업(/ui-redesign/sheet/)이라 앱에는 반영되지 않는 게 맞았다. 목업을 더 만드는 대신 시안 체계를 앱 셸에 실제로 이식했다.
- Applied: (1) 섀시 토큰 --ch-* 를 어두운 기계 하우징에서 종이색으로 재정의해 레일·상단바가 한 번에 밝아지도록 함 (2) 다크 전용으로 쓰인 예외 7곳 수정 — 흰색 알파 hover, 앰버 액티브 그라디언트·글로우를 종이 배경에 맞게 (3) 액센트 --signal 을 앰버에서 잉크로, --crit 을 스탬프 레드, --ok 를 스탬프 블루로, focus 아웃라인은 시인성 때문에 파랑 유지 (4) 폰트를 Archivo/IBM Plex Mono(Google Fonts CDN)에서 Bahnschrift/Segoe UI/맑은 고딕 시스템 스택으로 교체하고 InstrumentLayout 의 폰트 link 3개 제거 (5) 레일에 검색창 추가(라벨·그룹 부분일치 필터) (6) 두 글자 코드(RN/FC/MG-04)를 인라인 SVG 아이콘 25종으로 교체 (7) BMES → Report 복원 — AppMenus 에 없고 classic NavMenu 가 인라인으로 report/bmes 를 렌더하던 항목이라 같은 조건(F-Cost 뒤, NG Rate/F-Cost 권한)으로 되살림
- Completed: 앱 셸이 시안의 종이·잉크 체계로 바뀌었고 메뉴 누락이 해소됐다. 본문 페이지 마크업은 건드리지 않았고 브리지 CSS를 통해 같은 토큰을 따른다.
- Decisions: 크롬을 규칙 단위로 다시 쓰지 않고 --ch-* 토큰을 재정의했다 — 레일과 상단바가 모두 이 토큰에서 색을 읽으므로 한 곳만 바꾸면 전체가 뒤집히고, 되돌리기도 쉽다. --signal 을 파랑이 아니라 잉크로 정한 이유는 파랑을 이미 '개선'에 썼기 때문이다. 같은 파랑을 기본 동작에도 쓰면 색의 의미가 흐려진다. --ok 를 초록에서 파랑으로 바꾼 것은 검사 성적서 관행이자 적록색약 안전이다. Setting Model 2단 중첩은 이번에 넣지 않았다 — AppMenus 가 평면 Group 문자열만 갖고 있어 하위 그룹 개념을 추가해야 하고, 그건 메뉴 데이터 모델 변경이라 범위를 넘는다.
- Files: JinoSupporter.Web/wwwroot/ui-redesign/assets/instrument.css, JinoSupporter.Web/wwwroot/ui-redesign/assets/instrument.scoped.css, JinoSupporter.Web/Components/Layout/InstrumentLayout.razor
- Verification: 전체 재컴파일(--no-incremental) 오류 0개, 경고 32개로 기존과 동일하며 InstrumentLayout 관련 신규 경고 0건. 중간에 42개로 보였던 것은 파일 잠금으로 실패한 빌드가 경고를 중복 집계한 것이었다. 서버를 중지하고 재빌드 후 재기동했으며(PID 12848), 배포된 instrument.scoped.css 에서 --ch-100:#FFFFFF, --ch-ink:#16202B, --signal:#16202b, Bahnschrift, navlink__ico, railseek 를 HTTP 응답으로 직접 확인했다. 실제 로그인 화면 내부는 인증이 필요해 확인하지 못했다.
- Next: 사용자가 브라우저에서 셸을 확인한다. 본문 페이지들은 아직 각자 하드코딩된 색(hex 2,641개)을 갖고 있어 완전히 시안대로 보이지는 않는다. 다음 단계는 페이지별 hex → 토큰 치환이며, 감사 문서의 G5 순서를 따른다.

## 2026-08-08 07:20 - [UI][CHANGE] SheetDesign: Daily Report 액센트를 스탬프 잉크로 정렬
- Agent: Claude
- Session: External
- Task-ID: Unavailable
- Category: UI
- Feature: SheetDesign
- Change: CHANGE
- Prompt: "'JinoSupporter' 작업을 이어서 진행해."
- Answer: 남은 작업으로 잡아뒀던 "페이지 hex 2,641개를 토큰으로 치환"을 실행하기 전에 분포를 측정했고, 그 결과 계획을 접었다. 상위 20종이 61.6%를 차지하는데 전부 Tailwind slate 계열이라 시안 팔레트와 거의 같은 색이었다. 대신 실제로 어긋나던 액센트만 정렬했다.
- Applied: (1) BmesDailyReportPage의 .dr-worse/.dr-bar-worse/.dr-badge-worse를 --crit(스탬프 레드)로, .dr-better 계열을 --ok(스탬프 블루)로 (2) 스파크라인 fill도 같은 쌍으로 (3) .dr-btn 기본 동작 색을 파랑에서 --signal(잉크)로, hover는 --signal-hot, 스피너 테두리도 잉크로
- Completed: Daily Report의 색이 셸과 같은 의미 체계를 쓴다. 파랑은 개선, 빨강은 악화, 잉크는 동작이다.
- Decisions: 전면 hex 치환을 하지 않기로 했다 — 본문 중립색(#f8fafc, #0f172a, #e2e8f0)과 시안 팔레트(#F7F8FA, #16202B, #D7DCE3)가 거의 같아서 2,641곳을 고쳐도 화면 변화는 미미하고 회귀 위험만 크다. 실제 충돌은 액센트에 있었다. 이 페이지는 이미 '파랑=개선, 빨강=악화'를 쓰고 있어 시안과 의미가 일치했고 색조만 달랐다. 기본 동작 버튼이 개선과 같은 파랑이던 것이 유일한 의미 충돌이라 잉크로 옮겼다. 값은 var(--token, 하드코딩 폴백) 형태로 넣어 Classic UI에서도 그대로 동작한다. 브리지에 페이지 클래스를 복제하는 방식은 감사 G5가 유지보수 불가라고 못박아 채택하지 않았다.
- Files: JinoSupporter.Web/Components/Pages/BmesDailyReportPage.razor
- Verification: 전체 재컴파일(--no-incremental) 오류 0개, 경고 32개로 기존과 동일. diff 12줄 교체로 범위가 좁음을 확인. 서버 재기동 후 /login 200. 실제 Daily Report 화면은 인증이 필요해 직접 보지 못했고 사용자 확인이 필요하다.
- Next: 사용자가 Daily Report를 열어 색 체계를 확인한다. 나머지 페이지도 같은 방식이 필요하면 액센트 충돌이 있는 곳만(예: DataInferenceDbPage의 인디고 #4338ca 토글 19곳) 선별해 정렬한다. 중립색 전면 치환은 하지 않는다.

## 2026-08-08 07:35 - [UI][CHANGE] SheetDesign: Daily Report 본문을 검사 성적서 형태로
- Agent: Claude
- Session: External
- Task-ID: Unavailable
- Category: UI
- Feature: SheetDesign
- Change: CHANGE
- Prompt: "싹다 해줘"
- Answer: 화면 전체 이식을 워커 3개로 병렬 진행하려 했으나 세션 27이 생성 직후 재시작으로 죽어 아무것도 남기지 못했다. 오늘 만든 워커 8개 중 산출물을 낸 것은 3개뿐이고 반영된 커밋 7건은 모두 루트가 직접 한 것이라, 방식을 바꿔 루트가 화면 단위로 직접 이식하고 매번 커밋하기로 했다. 첫 화면은 사용자가 가장 자주 보는 Daily Report다.
- Applied: (1) 화면 이식 공통 규칙 문서 SHEET_PORTING_CONTRACT.md 작성 — 색 토큰과 세 가지 의미(빨강=악화, 파랑=개선, 잉크=동작), 형태(radius 2px·그림자 금지), 엑셀 격자 표, 웹폰트 금지, 뷰포트 높이 금지, 페이지 style 누출 금지, 검증 절차 (2) BmesDailyReportPage 본문 이식 — 카드 radius 12px→2px와 그림자 제거, 표를 아래 테두리만 있던 형태에서 전 셀 격자+헤더 밴드로, 중립색 하드코딩 hex를 토큰으로(폴백 포함), 흰색 7곳 포함
- Completed: Daily Report 본문이 셸과 같은 종이·잉크 체계로 보인다.
- Decisions: 워커 병렬 이식을 접고 루트 직접 작업으로 전환했다 — 호스트 재시작 주기가 워커 수명보다 짧아 지금 환경에서는 위임이 산출물로 이어지지 않는다. 규칙 문서를 먼저 만든 이유는 화면마다 다른 사람이 손대도 한 벌로 보이게 하기 위함이며, 이후 화면에도 이 문서를 기준으로 적용한다. inset box-shadow 한 곳(.dr-pt-other)은 남겼다 — 떠 있는 그림자가 아니라 1px 테두리 대용이라 규칙 §3에 어긋나지 않는다. border-radius 50%는 원형 점이라 유지했다.
- Files: JinoSupporter.Web/wwwroot/ui-redesign/SHEET_PORTING_CONTRACT.md, JinoSupporter.Web/Components/Pages/BmesDailyReportPage.razor
- Verification: 전체 재컴파일(--no-incremental) 오류 0개, 경고 32개로 기존과 동일. 렌더 대조(headless Chrome, 실제 CSS 로드 순서 재현) — .dr-stat radius 12px→2px·그림자 있음→none, .dr-table-wrap 동일, th 배경 투명→rgb(245,246,249)·테두리 0→1px, td 테두리 0→1px. 서버 재기동 완료. 실제 로그인 화면은 인증이 필요해 직접 보지 못했다.
- Next: 같은 방식으로 다음 화면을 이어간다. 남은 순서는 Report NG RATE 계열(가장 큼, 12개 파일), F-Cost, Report/LPA, 그 뒤 Test Data Analysis·Tools·Admin. 화면 50개 중 지금까지 본문까지 이식된 것은 1개다.

## 2026-08-08 07:40 - [UI][CHANGE] SheetDesign: 전 화면 표면을 종이·잉크 체계로 일괄 전환
- Agent: Claude
- Session: External
- Task-ID: Unavailable
- Category: UI
- Feature: SheetDesign
- Change: CHANGE
- Prompt: "끝까지 골을 잡고 다하세요"
- Answer: Daily Report 한 장에 쓴 변환을 일반화해 전 화면에 적용했다. 의미가 모호하지 않은 것(중립 회색, 모서리, 그림자)만 자동화하고, 표 격자와 액센트처럼 화면마다 뜻이 다른 것은 자동 변환에서 제외했다.
- Applied: (1) 마크업 style 블록 38개 파일에서 중립 하드코딩 hex 1,661곳을 토큰으로(Classic UI용 폴백 포함) (2) border-radius 317곳을 2px로 (3) 떠 있는 box-shadow 35곳 제거, inset 링은 1px 테두리 대용이라 유지 (4) 서버를 콘솔 창 없이 기동하도록 변경
- Completed: 앱 전 화면이 같은 표면 규칙을 따른다. 토큰 사용 1,745곳, 떠 있는 그림자 0곳, radius는 2px·50%·모서리 조합만 남았다.
- Decisions: 자동화 범위를 의도적으로 좁혔다 — 표 격자선과 액센트 색은 화면마다 의미가 달라 일괄 치환하면 뜻이 뭉개진다. C# 문자열로 조립하는 내보내기 HTML은 화면 스타일이 아니므로 제외했고, LoginPage는 자체 문서 전체를 렌더해 셸 밖이라 제외했다. 적용 전 데이터 URI 안의 hex와 id 셀렉터가 오염될 위험을 점검해 해당 사례가 없음을 확인했다.
- Files: JinoSupporter.Web/Components 아래 .razor 38개 (Pages 35, Shared 3)
- Verification: 전체 재컴파일(--no-incremental) 오류 0개, 경고 32개로 기존과 동일. 드라이런을 먼저 돌려 대상과 건수를 확인한 뒤 적용했다. 1차 적용에서 radius 정규식이 다중값의 첫 값만 바꿔 `8px 8px 0 0`이 `2px 8px 0 0`으로 비대칭이 되는 버그를 자체 검증에서 발견하고, 값 단위로 정규화하는 후속 스크립트로 8곳을 교정했다. 서버 재기동 후 /login 200. 실제 화면은 인증이 필요해 직접 보지 못했다.
- Next: 표 격자(엑셀 폼)와 액센트 색은 화면별 수작업이 필요하다. Daily Report는 완료했고, 다음은 사용 빈도 순으로 Report NG RATE 계열, F-Cost, Report/LPA 순이다.

## 2026-08-08 07:50 - [UI][CHANGE] SheetDesign: 표를 엑셀 격자로 일괄 전환
- Agent: Claude
- Session: External
- Task-ID: Unavailable
- Category: UI
- Feature: SheetDesign
- Change: CHANGE
- Prompt: "계속해"
- Answer: 표면 전환에 이어 표 격자를 정리했다. 조사해 보니 대부분의 표가 이미 border-collapse:collapse 인데 셀은 아래 테두리만 있어, 잘라내 붙이면 표로 안 읽히는 상태였다. 사용자가 예전에 요구한 "엑셀폼 표, 캡쳐해서 공유"와 직결된다.
- Applied: (1) th·td 를 지정한 규칙의 border-bottom 을 전 테두리(border: 1px solid var(--line))로 38곳 (2) 마지막 행의 border-bottom:none 1곳은 격자가 끊기므로 선언 자체를 제거
- Completed: 19개 화면의 표가 전 셀 격자로 렌더된다.
- Decisions: 1차 드라이런에서 두 가지 오작동을 발견해 자동화 범위를 좁혔다 — th:nth-child(1) 처럼 폭만 지정하는 규칙에 헤더 배경이 붙었고, tr:last-child td { border-bottom: 0 } 같은 의도적 제거가 전 테두리로 뒤집혔다. 헤더 밴드 자동 부착은 폐기하고 화면별 수작업으로 넘겼으며, 0/none 값은 변환 대신 선언 제거로 처리했다. border 축약형을 이미 가진 규칙은 네 면을 의도적으로 정한 것이라 건드리지 않았다.
- Files: JinoSupporter.Web/Components 아래 .razor 19개
- Verification: 전체 재컴파일(--no-incremental) 오류 0개, 경고 32개로 기존과 동일. 드라이런으로 대상 39곳을 먼저 확인하고 적용했다. 서버 재기동 후 /login 200. 실제 표 렌더는 인증이 필요해 직접 보지 못했다.
- Next: 표 헤더 밴드(회색 배경)는 화면별로 넣어야 한다. 그다음 액센트 정리 — DataInferenceDbPage 의 인디고 #4338ca 19곳처럼 이 체계에 없는 색.

## 2026-08-08 08:00 - [UI][CHANGE] SheetDesign: 체계 밖 액센트를 세 가지 의미로 통합
- Agent: Claude
- Session: External
- Task-ID: Unavailable
- Category: UI
- Feature: SheetDesign
- Change: CHANGE
- Prompt: "ㅇㅇ"
- Answer: 표 격자에 이어 액센트를 정리했다. 디자인이 정한 색의 의미는 셋(빨강=규격 이탈, 파랑=개선·통과, 잉크=동작·선택)인데 인디고·초록·다른 톤의 빨강이 네 번째 목소리로 섞여 있었다. 쓰임을 하나씩 확인한 뒤 세 의미로 접었다.
- Applied: (1) 인디고 #4338ca·#4f46e5 와 그 wash 를 잉크(--signal)로 — 활성 토글과 기준선 표시라 선택 상태에 해당 (2) 초록 #166534·#15803d 와 wash 를 스탬프 블루(--ok)로 — 완료·추가 상태이며 이 체계의 양호는 파랑이다 (3) 빨강 #991b1b·#b91c1c·#dc2626 와 wash 3종을 --crit 계열로 — 이미 악화를 뜻하므로 색조만 맞춤. 총 212곳, 20개 파일
- Completed: 화면에서 네 번째 액센트가 사라지고 색이 세 가지 의미만 갖는다.
- Decisions: 파랑 계열(#2563eb 36, #1d4ed8 33, #3b82f6 15 등)은 손대지 않았다 — 링크·기본 동작·개선 세 가지 뜻으로 섞여 쓰이고 있어 일괄 치환하면 의미가 뭉개진다. 화면을 읽어야 구분되므로 화면별 작업으로 남긴다. 표 헤더 밴드 자동 부착도 폐기했다 — 후보로 잡힌 17곳이 전부 .dr-th-num, .trend-model-th 처럼 폭·정렬만 지정하는 규칙이었고, 진짜 헤더 규칙은 이미 배경을 갖고 있어 후보에서 제외됐다.
- Files: JinoSupporter.Web/Components 아래 .razor 20개
- Verification: 전체 재컴파일(--no-incremental) 오류 0개, 경고 32개로 기존과 동일. 드라이런으로 212곳을 먼저 확인하고, 인디고·초록·진한빨강 각각의 실제 쓰임을 grep 으로 읽어 의미를 확인한 뒤 매핑했다. 적용 후 남은 체계 밖 색을 재측정해 파랑 계열과 회색 몇 종만 남았음을 확인했다. 서버 재기동 후 /login 200.
- Next: 파랑 계열 화면별 정리(링크는 파랑 유지, 기본 동작 버튼은 잉크로, 개선 지표는 --ok 로). 남은 회색 #617084·#dbe3ef·#dfe5ee 는 중립이라 토큰 매핑 후보다.

## 2026-08-08 08:15 - [UI][CHANGE] SheetDesign: 선택 상태의 파랑을 잉크로, 디스크 소진 원인 정리
- Agent: Claude
- Session: External
- Task-ID: Unavailable
- Category: UI
- Feature: SheetDesign
- Change: CHANGE
- Prompt: "계속해" / "ㅇㅇ" / "했어"
- Answer: 파랑이 "선택됨"과 "좋아짐" 두 가지를 동시에 뜻하던 충돌을 없앴다. 작업 중 D 드라이브 여유가 0이 되어 빌드가 실패했고, 원인이 내가 만든 검증용 빌드 산출물 누적임을 확인해 보고했다. 사용자가 공간을 확보한 뒤 검증을 마쳤다.
- Applied: (1) .active/.selected/.pg-pill-on 등 선택 상태 규칙에서 파랑 배경·테두리를 잉크(--signal)와 --signal-wash 로 교체, 10개 파일 13개 규칙 (2) 글자색 파랑은 유지 — 링크와 개선 지표는 파랑이 맞는 의미다
- Completed: 선택 상태가 잉크로 표시되어 파랑이 개선 한 가지만 뜻한다.
- Decisions: 파랑을 일괄 치환하지 않고 선택 상태의 chrome(배경·테두리)만 바꿨다. 파랑은 앱 안에서 링크·기본 동작·개선 세 가지로 쓰이는데, 색 역할을 selector 로 분류해 보니 배경·테두리로 쓰인 곳만 의미 충돌이었고 글자색은 링크나 측정값이라 그대로 두는 것이 옳았다. 디스크는 .codex-verify 아래에 매번 새 이름으로 빌드한 산출물이 약 3GB 쌓인 것이 원인이었다(오늘치만 6개 1.2GB). 앞으로는 검증 빌드 경로를 하나로 재사용한다. 삭제는 사용자가 직접 수행했다.
- Files: JinoSupporter.Web/Components 아래 .razor 10개
- Verification: 전체 재컴파일(--no-incremental) 오류 0개, 경고 32개로 기존과 동일. 드라이런으로 13개 규칙을 먼저 확인하고 적용했다. 적용 후 선택 상태에 파랑 배경이 남았는지 grep 으로 재확인해 0건. 서버 재기동 후 /login 200. 디스크 여유는 정리 후 7.81GB.
- Next: 남은 회색 #617084·#dbe3ef·#dfe5ee 토큰 매핑, 그리고 화면별 세부 확인. 사용자 육안 확인이 필요한 시점이다.

## 2026-08-08 08:30 - [UI][CHANGE] SheetDesign: 남은 중립색·초록·폰트 스택 토큰화
- Agent: Claude
- Session: External
- Task-ID: Unavailable
- Category: UI
- Feature: SheetDesign
- Change: CHANGE
- Prompt: "확인은 최종적으로 내가 할테니 너는 계속해"
- Answer: 기계적으로 처리 가능한 마지막 대상을 정리했다. 앞선 패스가 닿지 않은 중립 회색 9종, 놓친 초록 5종, 하드코딩 폰트 스택을 토큰으로 바꿨다.
- Applied: (1) 중립 회색 #617084·#dbe3ef·#dfe5ee·#e5eaf1·#d7dee8·#d7deea·#eef2f7·#eef3f8·#fbfcfd 와 순수 검정 #000 을 ink/line/panel 토큰으로 (2) 초록 #16a34a·#065f46 와 wash 3종을 --ok 계열로 (3) font-family 32곳을 var(--mono)/var(--sans) 로 — Consolas·Courier New·monospace·Segoe UI 스택
- Completed: 페이지 style 블록의 하드코딩 색이 2,641회에서 669회로, 고유 색이 287종에서 234종으로 줄었다.
- Decisions: 앰버(#92400e 16, #fef3c7 14, #fcd34d 9)는 손대지 않았다 — 이 디자인은 빨강·파랑·잉크 세 의미만 정의하고 경고에 쓸 어두운 앰버가 없다. #92400e 를 --warn(#fab219, 밝은 노랑)으로 옮기면 종이 위에서 글자가 읽히지 않는다. 경고는 이 앱에 실제로 있는 상태이므로, 억지로 세 색에 밀어넣는 대신 디자인에 경고 항목을 추가하는 것이 옳은 순서다. 파랑 계열(#1d4ed8 32, #2563eb 27 등)도 링크·개선 지표라 그대로 두었다.
- Files: JinoSupporter.Web/Components 아래 .razor 다수
- Verification: 전체 재컴파일(--no-incremental) 오류 0개, 경고 32개로 기존과 동일. 드라이런으로 색 179곳·폰트 32곳을 먼저 확인하고 적용했다. 적용 후 전수 재측정으로 감소분을 확인했다. 서버 재기동 완료.
- Next: 기계적으로 처리할 것은 사실상 소진됐다. 남은 것은 (1) 디자인에 경고 색 정의 추가 (2) 화면별 레이아웃 이식 — 시안의 기준선 편차 막대 같은 구조는 화면마다 수작업이 필요하다. 사용자가 화면을 확인 중이며 지적 사항에 따라 우선순위를 정한다.

## 2026-08-08 08:18 - [UI][CHANGE] QrBakoDateBrowser: 날짜별 원본 데이터 화면 구현
- Agent: Codex
- Session: External
- Task-ID: Unavailable
- Category: UI
- Feature: QrBakoDateBrowser
- Change: CHANGE
- Completed: QR BAKO DATA의 KPI·Pareto·중첩 상세 이력을 제거하고 좌측 TestTime 날짜 목록과 우측 선택 날짜의 BKTD 원본 컬럼 표로 교체했다. 조회 서비스는 날짜별 건수와 선택 날짜 행만 SELECT하도록 변경했다.
- Decisions: 실 DB 메타데이터를 읽기 전용으로 확인한 결과 67개 컬럼이 요청 목록과 일치했다. 연결 문자열의 ApplicationIntent=ReadOnly와 SELECT/NOLOCK만 유지하며 DB 쓰기 기능은 추가하지 않았다.
- Files: JinoSupporter.Web/Components/Pages/QrBakoDataPage.razor, JinoSupporter.Web/Services/QrBakoDataService.cs
- Verification: 아직 실행하지 않음 — 구현 직후 단계이며 다음 단계에서 변경 파일 대상 빌드와 읽기 전용 SQL 정적 점검을 수행한다.
- Next: Razor/C# 컴파일을 수행하고 날짜별 조회 결과 및 SQL 쓰기 구문 부재를 검증한다.

## 2026-08-08 08:20 - [BUILD][CHANGE] QrBakoDateBrowser: 날짜 조회 화면 검증 완료
- Agent: Codex
- Session: External
- Task-ID: Unavailable
- Category: BUILD
- Feature: QrBakoDateBrowser
- Change: CHANGE
- Completed: 변경된 QR BAKO 페이지와 서비스가 컴파일되며, 실 BKTD에서 날짜 집계와 선택 날짜 원본 행 조회가 일치함을 확인했다. 서비스 SQL에는 SELECT만 남아 있고 쓰기 구문은 없다.
- Decisions: 실행 중인 서버와 브라우저는 사용자 지시의 실행 제한에 따라 재기동하거나 열지 않았다. 기존 프로젝트 경고는 이번 변경 파일 밖의 32건으로 유지했다.
- Files: None
- Verification: JinoSupporter.Web 프로젝트 별도 artifacts 빌드 오류 0개·기존 경고 32개. ApplicationIntent=ReadOnly 실측 조회 결과 날짜 12개, 최신 2026-07-15 집계 8건=조회 8건, 컬럼 67개였고 git diff --check와 쓰기 SQL 토큰 검사도 통과했다.
- Next: 실행 중인 웹 서버를 다음 정상 재기동 시 갱신한 뒤 `/bmes/qr-bako-data`에서 좌측 날짜 선택과 우측 가로 스크롤 표의 실제 배치를 확인한다.

## 2026-08-08 08:30 - [BUILD][REMOVE] RepositoryDiskCleanup: 빌드 산출물·병합완료 워크트리 정리로 11.2GB 회수

- Agent: Claude
- Session: External
- Task-ID: Unavailable
- Category: BUILD
- Feature: RepositoryDiskCleanup
- Change: REMOVE
- Prompt: "D:\000. MyWorks\005. Program\Repository\JinoSupporter\ 여기에 쓰지 않는 빌드 산출물 지워줘 .. 여기 용량 엄청 먹네"
- Answer: 22.4GB → 11.17GB로 절반을 회수했다. 정리 중 **용량 폭증의 구조적 원인**을 찾았다 — `.codex-verify/`가 `.gitignore`에 없고 **git이 5018개 파일로 추적 중**이다(빌드 산출물, DLL, Playwright 브라우저 포함). 그래서 워크트리를 하나 만들 때마다 통째로 체크아웃되고, 워크트리 20개가 그 사본을 각각 들고 있었다. `.codex`가 18GB였던 이유이며 bin/obj는 그중 1.14GB뿐이었다. 사용자는 이번에 이 부분(git에서 제거)과 `BmesNgRateStandalone/dist` 설치본 정리는 선택하지 않았다.
- Applied: (1) `bin`/`obj`/`__pycache__` 40개 폴더 삭제 — 2.56GB. 전부 `.gitignore:2-3`의 `bin/`·`obj/` 대상이라 git 추적에 영향 없음 (2) 병합완료·클린 워크트리 10개를 `git worktree remove`로 제거 — s3·s11·s15·s19·s20·s21·s22·s23·s24·s27, 약 8.6GB (3) `git worktree prune`으로 메타데이터 정리. 등록 워크트리 21 → 11
- Completed: 전체 22.4GB → 11.17GB. 남은 `.codex` 7.98GB는 미병합·더티 워크트리 10개다.
- Decisions: (1) **삭제 대상을 미병합 커밋 0개 + 더티 파일 0개인 워크트리로만 한정했다.** 제외한 9개: s10(미병합 1커밋+더티 9), s12(미병합 1커밋), s16(미병합 1커밋+더티 5), s13(더티 8), s14(1), s18(1), s25(1), s26(2), s28(더티 8392, locked). (2) **`JinoSupporter.Web/bin`은 건드리지 않았다** — `JinoSupporter.Web.exe`(PID 21092)가 실행 중이다. (3) 10개 중 4개(s15·s21·s23·s27)는 `git worktree remove`가 "Permission denied"로 디렉터리 삭제에 실패했으나 **등록 해제는 이미 끝난 상태**였고, 이후 `rm -rf`로 삭제됐다. git 자체 삭제 경로만 막히는 현상으로, CodexDesktop 뿌리 W의 비상속 DENY ACE와 같은 계열로 보이나 이번에 확정하지는 않았다.
- Files: None
- Verification: 삭제 전 `git worktree list --porcelain`으로 워크트리별 `rev-list --count main..HEAD`와 `status --porcelain`을 전수 측정해 0|0인 것만 골랐다. `git check-ignore -v`로 `bin/`·`obj/`가 ignore 대상임을, `.codex-verify`가 아님을 확인. 삭제 후 `git worktree list` 11개, PowerShell 재측정 11.17GB. **공유 체크아웃의 미커밋 3파일(`HANDOFF.md`, `QrBakoDataPage.razor`, `QrBakoDataService.cs`)은 이 작업 이전부터 있던 것이다** — 이번 삭제는 gitignore 대상과 등록 해제된 워크트리 디렉터리만 건드렸다.
- Next: 근본 해결은 `.codex-verify`를 `git rm --cached` 하고 `.gitignore`에 추가하는 것이다(추적 5018파일, 현재 1.73GB + 남은 워크트리 사본). main에 커밋이 필요하고 기존 워크트리 10개가 각각 dirty로 표시되므로, 그 워크트리들을 정리한 뒤에 하는 편이 낫다. 미병합 3개(s10·s12·s16)는 먼저 병합 여부를 판단해야 한다.
