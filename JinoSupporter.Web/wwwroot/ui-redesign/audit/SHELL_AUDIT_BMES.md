# INSTRUMENT 셸 BMES·NG Rate 읽기 전용 감사

감사 기준 커밋은 `542e5954252538f16f5114d548ceb49daaf1e7b0`이다. 지정된 29개 Razor 파일만 감사 대상으로 읽었고, `instrument.css`와 `InstrumentLayout.razor`는 셸 동작을 판단하는 참고 자료로만 읽었다. 런타임 확인, 서버 조작, 빌드, 코드 수정은 하지 않았다.

먼저 모든 대상 파일에 아래 정규식을 동일하게 적용했다.

```text
rg -n -e '100vh|100dvh|position:\s*fixed|position:\s*sticky|height:\s*100%|max-height|overflow' -- <파일>
```

이 문서의 줄번호와 인용문은 이 검색 출력에 실제로 나온 줄만 사용한다. `overflow`가 넓은 패턴이므로 `text-overflow`, `overflow-wrap`, 반응형 해제 규칙인 `max-height: none`도 건수에 포함된다. 따라서 제시된 교차 검증치와 달랐던 파일은 실제 출력 기준으로 다음과 같으며, 값을 맞추기 위해 줄을 버리지 않았다.

| 파일 | 제시치 | 실제 `rg -n` 줄 수 |
|---|---:|---:|
| NgRateReportStyles.razor | 26 | 32 |
| BmesTest3Page.razor | 14 | 17 |
| QrBakoDataPage.razor | 17 | 18 |
| BmesTest5Page.razor | 5 | 8 |
| BmesTest4Page.razor | 4 | 5 |
| ReportPage.razor | 9 | 12 |
| BmesCauseMonthlyReportPage.razor | 3 | 4 |

나머지는 제시치와 일치한다. 위 차이는 부록 원출력으로 그대로 검증할 수 있다.

## 1. 요약 표

코드는 `V`=100vh 계열 가정, `P`=fixed, `H`=height 100% 체인, `S`=sticky 기준/상단 오프셋, `O`=자체 overflow에 의한 중첩 스크롤, `Z`=z-index/stacking context, `B`=§23–24 브리지가 직접 덮지 않는 전용 클래스다. `부분`은 브리지로 색·표면 같은 `B`만 고칠 수 있고 레이아웃은 해당 파일 수정이 필요하다는 뜻이다.

| 라우트/호스트 | 파일 | 근거 건수 | 최고 심각도 | 코드 | 전역 해결 가능 여부 |
|---|---|---:|---|---|---|
| `/report/bmes` | BmesReportPage.razor | 3 | 보기흉함 | B | 예(브리지, 시각만) |
| `/bmes/daily-report` | BmesDailyReportPage.razor | 12 | 보기흉함 | B | 예(브리지, 시각만) |
| `/bmes/f-cost`의 본문 | BmesFCostPage.razor | 8 | 보기흉함 | S, B | 부분(B만) |
| `/bmes/f-cost` | BmesFCost.razor | 0 | 사소 | — | 수정 불필요 |
| `/bmes/lpa` | BmesLpaPage.razor | 2 | 사소 | — | 수정 불필요 |
| `/bmes/setting`, `/setting` | BmesSettingPage.razor | 5 | 사소 | — | 수정 불필요 |
| `/bmes/test3` | BmesTest3Page.razor | 17 | 보기흉함 | V, O, B | 부분(B만) |
| `/bmes/test4` | BmesTest4Page.razor | 5 | 보기흉함 | O, B | 부분(B만) |
| `/bmes/test5` | BmesTest5Page.razor | 8 | 보기흉함 | O, B | 부분(B만) |
| `/bmes/make-model-group` | BmesMakeModelGroupPage.razor | 4 | 보기흉함 | O, B | 부분(B만) |
| `/bmes/routing-table` | BmesRoutingTablePage.razor | 3 | 보기흉함 | V, O, B | 아니오(페이지) |
| `/bmes/reason-table` | BmesReasonTablePage.razor | 3 | 보기흉함 | V, O, B | 아니오(페이지) |
| 직접 라우트 없음 | BmesCauseMonthlyReportPage.razor | 4 | 보기흉함 | V, O, B | 부분(B만) |
| `/bmes/qr-bako-data` | QrBakoDataPage.razor | 18 | 보기흉함 | V, O, B | 부분(B만) |
| `/bmes/worker-status` | WorkerStatusPage.razor | 1 | 사소 | B | 예(브리지, 시각만) |
| `/data-inference/db` | MicroSpeakerResultPage.razor | 7 | 보기흉함 | V, O, B | 부분(B만) |
| `/report` | ReportPage.razor | 12 | 보기흉함 | O, S, B | 부분(B만) |
| `/ng-rate` | NgRatePage.razor | 0 + 공유 | 보기흉함 | S, B(공유) | 부분(B만) |
| `/ng-rate-all` | NgRateAllPage.razor | 0 + 공유 | 보기흉함 | S, B(공유) | 부분(B만) |
| `/ng-rate-by-group` | NgRateByGroupPage.razor | 3 + 공유 | 보기흉함 | S, B | 부분(B만) |
| 직접 라우트 없음 | NgRateForDailyReportPage.razor | 0 + 공유 | 보기흉함 | S, B(공유) | 부분(B만) |
| 직접 라우트 없음 | NgRateForWeeklyReportPage.razor | 4 + 공유 | 보기흉함 | O, S, B | 부분(B만) |
| `/bmes/make-model-group` 내부 | SubGroupNode.razor | 0 | 사소 | — | 수정 불필요 |
| 공유 스타일 | NgRateReportStyles.razor | 32 | 보기흉함 | O, S, B | 부분(B만) |
| 공유 설정 패널 | NgRateSetupPanel.razor | 1 | 사소 | — | 수정 불필요 |
| 공유 내비게이션 | NgRateViewNav.razor | 0 | 사소 | — | 수정 불필요 |
| 공유 그룹 선택기 | NgRateModelGroupPicker.razor | 0 | 사소 | — | 수정 불필요 |
| 공유 단순 선택기 | NgRateSimpleGroupPicker.razor | 0 | 사소 | — | 수정 불필요 |
| 공유 계층 행 | HierSubRows.razor | 0 | 사소 | — | 수정 불필요 |

정적 감사에서 `사용불가`로 확정할 항목은 없었다. `.view` 자체가 세로 스크롤을 제공하므로 주된 실패 모드는 콘텐츠 소실보다는 과도한 높이, 중첩 스크롤, sticky 무효화와 구형 화면 스킨 혼재다.

## 2. 화면별 상세

### BmesReportPage (`/report/bmes`)

- 근거: `BmesReportPage.razor:127` `overflow-x: auto !important;`, `:152` `overflow: visible !important;`, `:153` `text-overflow: clip !important;`.
- 판정: 검색된 세 줄은 가로 표와 셀 텍스트 표시 규칙이며 자체 세로 스크롤이나 100vh 근거가 아니다. 다만 `.bmes-report-tab-content`, `.pivot-wrap` 같은 전용 클래스는 현재 Bootstrap 중심 브리지의 직접 대상이 아니므로 `B`다.
- 증상: 설정 패널과 보고서 컨테이너가 INSTRUMENT 표면·테두리 체계와 섞여 보일 수 있다.
- 심각도: `보기흉함`.
- 수정 구분: **전역**. 브리지에 BMES 보고서/NG Rate 공유 클래스를 추가하면 시각 혼재는 줄일 수 있다. 검색 근거만으로 페이지 레이아웃 수정 필요성은 확정하지 않는다.

### BmesDailyReportPage (`/bmes/daily-report`)

- 근거: `BmesDailyReportPage.razor:498` `overflow-x: auto;`, `:754` `overflow-x: auto;`. `:714` `overflow: hidden;`과 `:718` `height: 100%;`는 고정 높이 트랙 안의 막대다.
- 판정: 가로 스크롤은 넓은 표/로그를 위한 것이고, `height: 100%`는 높이 9px인 부모 트랙에 종속되어 `H` 체인 파손이 아니다. 기능상 셸 충돌 근거는 없지만 `.dr-*` 전용 스킨은 브리지가 덮지 않아 `B`다.
- 증상: 세로 스크롤은 정상이나 Daily 카드와 표가 새 셸과 다른 시각 언어로 남는다.
- 심각도: `보기흉함`.
- 수정 구분: **전역**. `.dr-*`의 표면·테두리·색만 scoped bridge로 재스킨할 수 있다.

### BmesFCostPage (`/bmes/f-cost` 본문)

- 근거: `BmesFCostPage.razor:171`, `:343`, `:440`, `:603`, `:815`의 실제 선언은 모두 `<div style="overflow-x:auto;">`; `:975`는 `position: sticky; top: 0; z-index: 1;`이다.
- 판정: 세로 스크롤은 `.view`가 소유하지만 sticky 헤더의 가장 가까운 overflow 조상은 각 가로 스크롤 래퍼다. 래퍼가 세로로 움직이지 않으므로 헤더가 `.view` 상단에 붙지 않는 `S` 가능성이 높다. `.fcost-*`는 `B`다.
- 증상: 긴 F-Cost 표를 아래로 읽을 때 열 제목이 사라지고, 표 스킨도 새 셸과 혼재한다.
- 심각도: `보기흉함`.
- 수정 구분: **페이지**. 가로 스크롤과 세로 sticky를 함께 만족하도록 표 래퍼 구조를 정해야 한다. 색상만 전역 브리지로 해결 가능하다.

### BmesFCost (`/bmes/f-cost` 라우트 래퍼)

- 근거: 지정 패턴 `0건`.
- 판정/증상: 래퍼 자체에는 감사 대상 선언이 없다. 실제 판정은 자식 BmesFCostPage에 있다.
- 심각도: `사소`.
- 수정 구분: **수정 불필요**.

### BmesLpaPage (`/bmes/lpa`)

- 근거: `BmesLpaPage.razor:100` `.lpa-prog { height: 4px; ... overflow: hidden; }`, `:101` `.lpa-prog-bar { height: 100%; ... }`.
- 판정: 100% 높이는 명시적으로 4px인 진행 막대 부모를 채우므로 `H`가 아니다. 검색 근거상 수정할 셸 충돌이 없다.
- 증상: 없음.
- 심각도: `사소`.
- 수정 구분: **수정 불필요**.

### BmesSettingPage (`/bmes/setting`, `/setting`)

- 근거: `BmesSettingPage.razor:334`의 로그는 `max-height:160px; overflow:auto;`; `:411`, `:558`은 `max-height:460px`인 표 래퍼다. `:713`은 `position: sticky; top: 0; z-index: 2;`, `:716`은 `position: sticky; top: 29px; z-index: 1; background: #fff;`이다.
- 판정: 맥락 확인 결과 411/558의 표와 713/716의 2단 헤더는 `d-none` 블록 안에 있어 현재 화면에 표시되지 않는다. 29px은 같은 로컬 스크롤박스 안에서 첫 헤더 아래에 둘째 검색 행을 붙이는 값이다. 활성 상태인 334의 작은 로그 스크롤은 의도된 국소 패널이다.
- 증상: 현재 표시 경로에서는 확인된 파손 없음. 숨은 표를 다시 활성화하거나 자체 스크롤을 제거할 때에는 `top: 29px` 관계를 함께 보존해야 한다.
- 심각도: `사소`.
- 수정 구분: **수정 불필요**.

### BmesTest3Page (`/bmes/test3`)

- 근거: `BmesTest3Page.razor:417` `overflow: auto;`, `:418` `max-height: calc(100vh - 150px);`; `:576` `max-height: calc(100vh - 96px);`, `:577` `overflow: hidden;`, `:582` `position: sticky;`; `:596` `max-height: calc(100vh - 150px);`, `:597` `overflow: auto;`.
- 판정: 표와 BOM 목록이 셸의 실제 남은 높이가 아닌 viewport에서 상수를 빼 높이를 잡는다(`V`). `.view` 스크롤 안에 표/BOM 스크롤이 다시 생긴다(`O`). sticky BOM 트레이 자체의 `top: 8px`은 `.view` 기준으로 동작할 수 있지만 내부 높이 계산은 여전히 잘못된다. `.test3-*`는 `B`다.
- 증상: 표와 BOM 영역에 별도 세로 스크롤바가 생기고, 탭 유무에 따라 아래쪽이 지나치게 길거나 이중 스크롤된다.
- 심각도: `보기흉함`.
- 수정 구분: **페이지**. viewport calc를 제거하고 페이지를 flex/grid의 `min-height: 0` 체인으로 이식해야 한다.

### BmesTest4Page (`/bmes/test4`)

- 근거: `BmesTest4Page.razor:307` `overflow: auto;`, `:308` `max-height: 360px;`, `:321` `position: sticky;`.
- 판정: 결과 표가 360px 자체 스크롤을 소유해 `.view`와 중첩된다(`O`). sticky `top: 0`은 그 로컬 스크롤박스 안에서는 맞다. `.test4-*`는 `B`다.
- 증상: 큰 결과에서 페이지와 표를 번갈아 스크롤해야 한다.
- 심각도: `보기흉함`.
- 수정 구분: **페이지**. 표 높이를 콘텐츠 흐름에 맡길지, 명시적인 작업영역 스크롤로 유지할지 결정해야 한다.

### BmesTest5Page (`/bmes/test5`)

- 근거: `BmesTest5Page.razor:320`/`:321`은 제안 목록 `max-height: 240px; overflow-y: auto;`, `:413`/`:414`는 모델 칩 `max-height: 160px; overflow-y: auto;`, `:453`/`:454`는 주 표 `max-height: 520px; overflow: auto;`, `:470`은 `position: sticky;`이다.
- 판정: 제안/칩 스크롤은 국소 UI로 타당하지만 주 표의 520px 스크롤은 `.view`와 중첩된다(`O`). 헤더 sticky는 주 표의 로컬 스크롤 기준으로는 정상이다. `.t5-*`는 `B`다.
- 증상: 긴 표에서 바깥 페이지와 안쪽 표의 스크롤 포커스가 갈린다.
- 심각도: `보기흉함`.
- 수정 구분: **페이지**. 주 표만 셸 소유 스크롤에 맞추고, 자동완성/칩의 작은 스크롤은 유지하는 구분이 필요하다.

### BmesMakeModelGroupPage (`/bmes/make-model-group`)

- 근거: `BmesMakeModelGroupPage.razor:66`, `:183`, `:245`는 각각 `max-height:70vh; overflow:auto;`; `:297`은 `position: sticky; top: 0; z-index: 1;`이다.
- 판정: 세 열이 각각 viewport의 70%를 자체 스크롤 높이로 사용한다(`O`). 탭 유무로 바뀌는 실제 `.view` 높이를 반영하지 못한다. sticky 헤더는 각 열 내부에서는 맞지만 세 개의 독립 스크롤을 고착한다. `.mmg-*`는 `B`다.
- 증상: 한 화면에 바깥 스크롤과 최대 세 개의 열 스크롤이 나타나며 열별 위치가 서로 어긋난다.
- 심각도: `보기흉함`.
- 수정 구분: **페이지**. 세 열을 셸의 남은 높이를 받는 하나의 grid/flex 작업영역으로 재구성해야 한다.

### BmesRoutingTablePage (`/bmes/routing-table`)

- 근거: `BmesRoutingTablePage.razor:67` `<div style="overflow-x:auto; max-height:calc(100vh - 190px); overflow-y:auto;">`; `:149`은 `position: sticky; top: 0; z-index: 2;`, `:151`은 `position: sticky; top: 29px; z-index: 1; background: #fff;`이다.
- 판정: 내부 표 높이가 viewport 기준이라 `V`, `.view`와의 이중 세로 스크롤이라 `O`다. 2단 sticky의 0/29px은 현재 같은 내부 스크롤박스 안에서는 일관되므로 단독 `S`로 판정하지 않았다. 자체 `.bmes-table`은 `B`다.
- 증상: 표와 페이지가 따로 스크롤되며 탭이 생기면 가용 높이보다 표가 길어진다.
- 심각도: `보기흉함`.
- 수정 구분: **페이지**. 인라인 높이/overflow를 직접 제거하거나 클래스화하고, 2단 헤더를 함께 이식해야 한다.

### BmesReasonTablePage (`/bmes/reason-table`)

- 근거: `BmesReasonTablePage.razor:73` `<div style="overflow-x:auto; max-height:calc(100vh - 190px); overflow-y:auto;">`; `:287`은 `position: sticky; top: 0; z-index: 2;`, `:289`는 `position: sticky; top: 29px; z-index: 1; background: #fff;`이다.
- 판정: Routing과 같은 `V`/`O` 패턴이다. 29px 둘째 행은 현재 로컬 스크롤박스 기준으로 맞지만, 외부 `.view`로 스크롤을 넘길 때는 첫 행 높이와 함께 다시 정해야 한다. `.bmes-table`은 `B`다.
- 증상: 표 내부와 페이지에 스크롤바가 동시에 생기고 탭 유무에 따라 높이 여백이 달라진다.
- 심각도: `보기흉함`.
- 수정 구분: **페이지**.

### BmesCauseMonthlyReportPage (직접 라우트 없음)

- 근거: `BmesCauseMonthlyReportPage.razor:185` `overflow: auto;`, `:186` `max-height: calc(100vh - 285px);`, `:200` `position: sticky;`, `:287` `max-height: none;`.
- 판정: 데스크톱에서 cause 표가 viewport calc와 자체 스크롤을 사용해 `V`/`O`다. 1200px 이하에서는 `max-height: none`으로 풀리지만 넓은 화면에서 셸 높이와 충돌한다. `.cause-*`는 `B`다.
- 증상: 넓은 화면에서 원인 표가 별도 스크롤되고 셸 탭 높이를 반영하지 못한다.
- 심각도: `보기흉함`.
- 수정 구분: **페이지**.

### QrBakoDataPage (`/bmes/qr-bako-data`)

- 근거: `QrBakoDataPage.razor:326` `min-height: calc(100vh - 1.5rem);`; `:687` `overflow: auto;`, `:688` `max-height: calc(100vh - 28rem);`, `:702` `position: sticky;`; `:940` `max-height: none;`.
- 판정: 페이지 루트와 이력 표가 viewport를 직접 소유한다고 가정한다(`V`). `.view` 안에서 이력 표가 다시 스크롤해 `O`가 된다. 1200px 이하에서는 내부 최대 높이를 해제하지만 루트의 viewport 최소 높이는 남는다. `.qr-*` 전체가 자체 다크 스킨이라 `B`다.
- 증상: 데스크톱에서 바깥 `.view`와 이력 표가 따로 움직이고, 페이지 전체가 실제 콘텐츠 영역보다 길어질 수 있다.
- 심각도: `보기흉함`.
- 수정 구분: **페이지**. 루트 최소 높이와 이력 영역을 셸의 flex/grid 잔여 높이에 연결해야 한다.

### WorkerStatusPage (`/bmes/worker-status`)

- 근거: `WorkerStatusPage.razor:33` `.ws-table-wrap { overflow-x: auto; }`.
- 판정: 가로 스크롤만 있고 높이 제한은 없어 `O`가 아니다. `.ws-*` 전용 카드/표 스킨이 브리지 밖에 남는 `B`만 있다.
- 증상: 기능은 유지되지만 카드와 표가 새 셸과 다른 모서리·표면을 사용한다.
- 심각도: `사소`.
- 수정 구분: **전역**. 시각 브리지로 충분하다.

### MicroSpeakerResultPage (`/data-inference/db`)

- 근거: `MicroSpeakerResultPage.razor:129` `min-height: calc(100vh - 24px);`; `:245` `overflow: hidden;`, `:246` `height: calc(100vh - 184px);`; `:253` `height: 100%;`; `:270` `height: calc(100vh - 260px);`.
- 판정: 루트와 iframe 패널이 모두 viewport에서 고정 상수를 빼 높이를 산출한다(`V`). 패널은 내용을 숨기고 iframe이 그 높이 100%를 채우므로 페이지 전체는 `.view`와 별도의 높이 체계를 갖는다(`O`). `:253`은 명시 높이 부모를 채우므로 독립 `H` 문제는 아니다. `.msr-*`는 `B`다.
- 증상: 탭이 나타나거나 좁은 화면이 되면 iframe 패널이 가용 높이보다 길어져 바깥 스크롤과 큰 빈/잘린 작업영역이 생긴다.
- 심각도: `보기흉함`.
- 수정 구분: **페이지**. iframe 화면을 `display:flex; flex-direction:column; min-height:0`인 작업영역으로 이식해야 한다.

### ReportPage (`/report`)

- 근거: `ReportPage.razor:26` `max-height:220px; overflow-y:auto;`, `:54` `max-height:500px; overflow-y:auto;`, `:136` `max-height:60vh; overflow:auto;`; `:208` `<div style="overflow-x:auto;">`, `:211` `position:sticky; top:0; z-index:1;`; `:315` `position: fixed; ... z-index: 9500;`, `:318` `max-height: 92vh;`.
- 판정: 태그·데이터셋·AI 보고서가 각각 세로 스크롤을 가져 `.view`와 중첩된다(`O`). 측정표의 sticky 행은 가로 overflow 래퍼에 갇혀 `.view` 스크롤 상단에 고정되지 않을 수 있다(`S`). 이미지 뷰어의 fixed/z-index는 실제 사용 요소가 있고 검토 대상이지만, 참고 셸의 `.shell/.stage/.view`에 transform·isolation·z-index stacking context가 없어 정적 구조상 현재 `P`/`Z` 파손은 입증되지 않았다. 대부분 인라인 스타일이라 `B` 성격도 강하다.
- 증상: 왼쪽 목록과 바깥 페이지 스크롤이 분리되고, 긴 측정표의 헤더가 사라진다. 이미지 뷰어는 현재 구조상 전체 화면을 덮을 것으로 판단된다.
- 심각도: `보기흉함`.
- 수정 구분: **페이지**. 주요 좌우 영역의 높이·스크롤 소유권과 표 sticky를 직접 정리해야 한다.

### NgRatePage (`/ng-rate`), NgRateAllPage (`/ng-rate-all`), NgRateForDailyReportPage

- 근거: 세 파일 자체의 지정 패턴은 각각 `0건`. 공유 스타일의 실제 근거는 `NgRateReportStyles.razor:305` `overflow-x: auto;`, `:327` `position: sticky;`, 그리고 `:67`/`:68`의 `max-height: 150px;`/`overflow: auto;`다.
- 판정: 페이지 파일 자체에는 수정 근거가 없지만 세 화면은 NgRateReportStyles를 사용한다. pivot 헤더 sticky는 가로 overflow 조상에 갇혀 `.view` 상단 sticky가 되지 않는 `S` 패턴이고, 공유 전용 클래스는 `B`다. 150px 그룹 목록은 작은 선택 패널의 의도된 국소 스크롤이므로 주 페이지 `O`로 확대 판정하지 않는다.
- 증상: 긴 NG Rate 표를 세로로 읽을 때 헤더가 유지되지 않고 공유 NG Rate 패널이 구형 스킨으로 남는다.
- 심각도: `보기흉함`.
- 수정 구분: **페이지/공유**. 각 0건 페이지가 아니라 NgRateReportStyles와 pivot 래퍼 구조를 고치는 것이 맞다. 시각만 전역 브리지로 해결 가능하다.

### NgRateByGroupPage (`/ng-rate-by-group`)

- 근거: `NgRateByGroupPage.razor:116`, `:266`, `:455`는 모두 `<div style="overflow-x:auto;">`; 공유 `NgRateReportStyles.razor:327`은 `position: sticky;`다.
- 판정: 각 표의 가로 overflow 래퍼가 공유 sticky 헤더의 기준 조상이 되어 `S`가 된다. 페이지와 공유 전용 클래스는 `B`다. 세로 높이 제한은 없어 `O`로 판정하지 않았다.
- 증상: 그룹·모델·Top 10 표를 아래로 스크롤하면 헤더가 `.view` 상단에 남지 않는다.
- 심각도: `보기흉함`.
- 수정 구분: **페이지/공유**. 인라인 래퍼와 공유 sticky 정책을 함께 바꿔야 한다.

### NgRateForWeeklyReportPage

- 근거: `NgRateForWeeklyReportPage.razor:131`, `:264`, `:434`, `:660`은 모두 `<div style="overflow-x:auto;">`. 공유 스타일은 `NgRateReportStyles.razor:754` `max-height: 360px;`, `:755` `overflow: auto;`, `:764` `position: sticky;`, `:789` `overflow-x: auto;`다.
- 판정: 일반 보고서 표는 가로 래퍼 때문에 `S`; weekly 설정 표는 360px 자체 세로 스크롤이라 `O`다. 공유 `.weekly-*`/`.pivot-*`는 `B`다.
- 증상: 일반 표 헤더는 바깥 스크롤에서 사라지고, 설정 표에는 별도 세로 스크롤바가 생긴다.
- 심각도: `보기흉함`.
- 수정 구분: **페이지/공유**.

### SubGroupNode

- 근거: 지정 패턴 `0건`.
- 판정/증상: 감사 대상 선언이 없고 자체 스크롤·viewport 높이 근거가 없다.
- 심각도: `사소`.
- 수정 구분: **수정 불필요**.

### NgRateReportStyles (공유)

- 근거: `NgRateReportStyles.razor:67`/`:68`은 150px 그룹 목록 스크롤, `:88`은 `height: 100% !important;`, `:90`/`:106`은 `max-height: none;`; `:262`와 `:703`은 `position: fixed;`; `:277`/`:717`은 `max-height: 80vh;`; `:305`와 `:327`은 가로 overflow + sticky; `:754`/`:755`/`:764`은 360px 설정 표 + sticky다.
- 판정: `:88`의 100%는 같은 grid row를 채우는 로그 카드이고 자식 로그는 flex/min-height 0이므로 검색 근거상 `H` 파손으로 보지 않는다. `:262`/`:703`의 fixed 오버레이 클래스는 지정 대상 파일에서 대응 마크업을 찾지 못해 현재 `P`/`Z` 장애로 확정하지 않았다. 실제 문제는 pivot 가로 스크롤이 sticky를 가두는 `S`, weekly 설정 표의 `O`, 그리고 광범위한 `.ngrate-*`, `.pivot-*`, `.weekly-*`가 브리지 밖인 `B`다.
- 증상: 여러 NG Rate 화면에 같은 헤더 비고정과 구형 공유 패널 스킨이 반복된다.
- 심각도: `보기흉함`.
- 수정 구분: **페이지/공유**. 기능 레이아웃은 이 파일과 소비자 래퍼를 고쳐야 하고, 색·표면은 전역 브리지로 일괄 개선할 수 있다.

### NgRateSetupPanel 및 나머지 Shared 컴포넌트

- 근거: `NgRateSetupPanel.razor:222` `<div ... class="font-monospace p-2 overflow-auto ngrate-log-pane">`. NgRateViewNav, NgRateModelGroupPicker, NgRateSimpleGroupPicker, HierSubRows는 모두 `0건`.
- 판정: SetupPanel의 overflow 클래스는 NgRateReportStyles에서 flex 로그 영역으로 제한되어 의도된 국소 스크롤이다. 나머지 네 컴포넌트에는 감사 패턴이 없다.
- 증상: 확인된 셸 파손 없음.
- 심각도: `사소`.
- 수정 구분: **수정 불필요**. 공유 스타일 이식 시 회귀 확인만 필요하다.

## 3. 전역으로 해결 가능한 항목 모음

1. **셸 축소 가능성 보강**: `InstrumentLayout`/`instrument.css`에서 `.stage`와 `.view`의 `min-height: 0; min-width: 0;`을 명시하는 것은 안전한 전역 보강이다. 다만 이 조치만으로 페이지 내부 `calc(100vh - ...)`나 `70vh`는 없어지지 않으므로 V/O 항목을 해결했다고 간주하면 안 된다.
2. **공유 시각 브리지 확장**: `.ngrate-*`, `.pivot-table`, `.weekly-form-*`, `.bmes-table`, `.fcost-*`, `.dr-*`, `.ws-*`의 배경·테두리·글꼴·색을 `.ins` scoped bridge에서 묶어 재스킨할 수 있다. 이것은 `B`만 해결하며 `height`, `max-height`, `overflow`, `position`은 전역에서 덮지 않는 것이 안전하다.
3. **NG Rate 한 지점 수정**: 여러 화면의 중복 증상은 NgRateReportStyles 한 파일에 집중되어 있다. pivot sticky 정책과 weekly 설정 스크롤을 그 공유 파일에서 정리하면 `/ng-rate`, `/ng-rate-all`, `/ng-rate-by-group`, Daily/Weekly 컴포넌트에 함께 반영된다. 다만 이는 instrument.css 전역 수정이 아니라 감사 대상 공유 파일 수정이다.
4. **fixed 오버레이는 현 상태 유지**: ReportPage의 z-index 9500 오버레이와 공유 스타일의 fixed 선언에 대해 셸 조상 stacking context 파손 근거가 없다. speculative한 전역 z-index 상향이나 fixed→absolute 변경은 하지 않는다.
5. **인라인 viewport/overflow는 전역 선택자로 잡지 않음**: Routing/Reason의 인라인 `calc(100vh - 190px)`, Report의 여러 인라인 max-height, NgRate 화면의 인라인 가로 래퍼를 `[style*=...]` 전역 선택자로 덮으면 다른 화면까지 오염된다. 이들은 페이지 수정으로 분류한다.

## 4. 이식 우선순위

아래는 페이지 또는 감사 대상 공유 파일을 직접 고쳐야 하는 항목만이다. 모두 정적 기준 최고 심각도는 `보기흉함`이며, 같은 심각도 안에서는 셸 높이 의존성과 영향 범위로 정렬했다.

1. **MicroSpeakerResultPage** — 루트와 iframe 패널에 100vh calc가 중복되고 `overflow: hidden`이 결합한다. iframe 작업영역을 셸 잔여 높이 flex로 바꾸는 것이 최우선이다.
2. **BmesTest3Page** — 표, sticky BOM 트레이, BOM 목록이 각각 viewport 높이와 자체 스크롤을 가진다. 독립 스크롤이 가장 많다.
3. **QrBakoDataPage** — 페이지 루트의 viewport 최소 높이와 이력 표의 viewport 최대 높이가 동시에 존재한다.
4. **BmesRoutingTablePage / BmesReasonTablePage** — 동일한 `calc(100vh - 190px)` 인라인 스크롤 패턴이다. 한 설계로 함께 고치되 2단 sticky의 0/29px 관계를 보존한다.
5. **BmesCauseMonthlyReportPage** — 동일한 viewport 기반 표 스크롤이며 공유 가능한 이식 패턴을 적용할 수 있다.
6. **BmesMakeModelGroupPage** — 세 열 각각의 `70vh` 스크롤을 하나의 셸 종속 작업영역으로 통합한다.
7. **ReportPage** — 태그, 데이터셋, AI 보고서의 세로 스크롤 소유권과 측정표 sticky를 좌우 레이아웃 단위로 재정의한다. fixed 이미지 뷰어는 유지한다.
8. **BmesTest5Page / BmesTest4Page** — 자동완성 같은 작은 국소 스크롤은 보존하고 주 결과 표 스크롤만 셸 정책과 맞춘다.
9. **BmesFCostPage** — 다섯 가로 래퍼와 sticky 헤더의 기준 조상을 정리한다.
10. **NgRateReportStyles + NgRateByGroupPage + NgRateForWeeklyReportPage** — 공유 pivot sticky와 weekly 설정 스크롤을 한 번에 정리한다. NgRatePage/NgRateAll/NgRateForDailyReportPage는 자체 0건이므로 우선 직접 수정하지 않는다.

## 5. 부록: 실제 검색 원본 출력

각 파일에 위 정규식을 단독 실행했을 때의 stdout이다. 파일명은 소제목으로 분리했으며, 코드 블록 내부는 `rg -n`이 출력한 `줄번호:실제 줄`을 그대로 옮겼다. `[0 matches]`는 stdout이 비어 있던 파일을 명시한 것이다.

### JinoSupporter.Web/Components/Pages/BmesReportPage.razor

```text
127:        overflow-x: auto !important;
152:        overflow: visible !important;
153:        text-overflow: clip !important;
```

### JinoSupporter.Web/Components/Pages/BmesDailyReportPage.razor

```text
498:        overflow-x: auto;
569:        overflow: hidden;
570:        text-overflow: ellipsis;
578:        overflow: hidden;
579:        text-overflow: ellipsis;
607:        overflow: hidden;
608:        text-overflow: ellipsis;
617:        overflow: hidden;
618:        text-overflow: ellipsis;
714:        overflow: hidden;
718:        height: 100%;
754:        overflow-x: auto;
```

### JinoSupporter.Web/Components/Pages/BmesFCostPage.razor

```text
171:                <div style="overflow-x:auto;">
343:                    <div style="overflow-x:auto;">
440:            <div style="overflow-x:auto;">
603:            <div style="overflow-x:auto;">
815:            <div style="overflow-x:auto;">
975:        position: sticky; top: 0; z-index: 1;
1042:    .fcost-weekly-table td.weekly-model-td { max-width: 180px; overflow: hidden; text-overflow: ellipsis; }
1150:        max-width: 240px; overflow: hidden; text-overflow: ellipsis;
```

### JinoSupporter.Web/Components/Pages/BmesFCost.razor

```text
[0 matches]
```

### JinoSupporter.Web/Components/Pages/BmesLpaPage.razor

```text
100:    .lpa-prog { height: 4px; background: #dfe5ec; border-radius: 2px; overflow: hidden; }
101:    .lpa-prog-bar { height: 100%; background: #0d6efd; border-radius: 2px; transition: width .15s linear; }
```

### JinoSupporter.Web/Components/Pages/BmesSettingPage.razor

```text
334:                    <pre class="small mb-2" style="background:#f8fafc; border:1px solid #e2e8f0; border-radius:4px; padding:8px; max-height:160px; overflow:auto;">@_fetchLog</pre>
411:            <div style="overflow-x:auto; max-height:460px; overflow-y:auto;">
558:            <div style="overflow-x:auto; max-height:460px; overflow-y:auto;">
713:        position: sticky; top: 0; z-index: 2;
716:        position: sticky; top: 29px; z-index: 1; background: #fff;
```

### JinoSupporter.Web/Components/Pages/BmesTest3Page.razor

```text
352:        max-height: 90px;
390:        overflow: hidden;
396:        height: 100%;
405:        overflow: hidden;
406:        text-overflow: ellipsis;
417:        overflow: auto;
418:        max-height: calc(100vh - 150px);
428:        position: sticky;
521:        overflow: hidden;
522:        text-overflow: ellipsis;
576:        max-height: calc(100vh - 96px);
577:        overflow: hidden;
582:        position: sticky;
596:        max-height: calc(100vh - 150px);
597:        overflow: auto;
633:        .test3-bom-tray { position: static; max-height: none; }
634:        .test3-bom-list { max-height: 260px; }
```

### JinoSupporter.Web/Components/Pages/BmesTest4Page.razor

```text
307:        overflow: auto;
308:        max-height: 360px;
321:        position: sticky;
363:        overflow: hidden;
364:        text-overflow: ellipsis;
```

### JinoSupporter.Web/Components/Pages/BmesTest5Page.razor

```text
320:        max-height: 240px;
321:        overflow-y: auto;
413:        max-height: 160px;
414:        overflow-y: auto;
440:        overflow: hidden;
453:        max-height: 520px;
454:        overflow: auto;
470:        position: sticky;
```

### JinoSupporter.Web/Components/Pages/BmesMakeModelGroupPage.razor

```text
66:            <div style="max-height:70vh; overflow:auto;" class="p-2">
183:            <div style="max-height:70vh; overflow:auto;">
245:            <div style="max-height:70vh; overflow:auto;">
297:        position: sticky; top: 0; z-index: 1;
```

### JinoSupporter.Web/Components/Pages/BmesRoutingTablePage.razor

```text
67:    <div style="overflow-x:auto; max-height:calc(100vh - 190px); overflow-y:auto;">
149:        border-bottom: 1px solid #e2e8f0; position: sticky; top: 0; z-index: 2;
151:    .bmes-search-row { position: sticky; top: 29px; z-index: 1; background: #fff; }
```

### JinoSupporter.Web/Components/Pages/BmesReasonTablePage.razor

```text
73:    <div style="overflow-x:auto; max-height:calc(100vh - 190px); overflow-y:auto;">
287:        border-bottom: 1px solid #e2e8f0; position: sticky; top: 0; z-index: 2;
289:    .bmes-search-row { position: sticky; top: 29px; z-index: 1; background: #fff; }
```

### JinoSupporter.Web/Components/Pages/BmesCauseMonthlyReportPage.razor

```text
185:        overflow: auto;
186:        max-height: calc(100vh - 285px);
200:        position: sticky;
287:            max-height: none;
```

### JinoSupporter.Web/Components/Pages/QrBakoDataPage.razor

```text
326:        min-height: calc(100vh - 1.5rem);
518:        overflow: hidden;
524:        text-overflow: ellipsis;
540:        overflow: hidden;
544:        text-overflow: ellipsis;
617:        overflow: hidden;
623:        text-overflow: ellipsis;
687:        overflow: auto;
688:        max-height: calc(100vh - 28rem);
702:        position: sticky;
716:        overflow: hidden;
720:        text-overflow: ellipsis;
871:        overflow-wrap: anywhere;
881:        overflow: hidden;
885:        text-overflow: ellipsis;
940:            max-height: none;
1240:        overflow: visible;
1242:        text-overflow: clip;
```

### JinoSupporter.Web/Components/Pages/WorkerStatusPage.razor

```text
33:    .ws-table-wrap { overflow-x: auto; }
```

### JinoSupporter.Web/Components/Pages/MicroSpeakerResultPage.razor

```text
129:        min-height: calc(100vh - 24px);
238:        overflow: hidden;
239:        text-overflow: ellipsis;
245:        overflow: hidden;
246:        height: calc(100vh - 184px);
253:        height: 100%;
270:            height: calc(100vh - 260px);
```

### JinoSupporter.Web/Components/Pages/ReportPage.razor

```text
26:            <div class="card-body py-2" style="max-height:220px; overflow-y:auto;">
54:            <div class="list-group list-group-flush" style="max-height:500px; overflow-y:auto;">
72:                            <div style="color:@(isActive ? "#1d4ed8" : isSelected ? "#2563eb" : "#334155"); font-weight:@(isSelected ? "600" : "500"); overflow:hidden; text-overflow:ellipsis; white-space:nowrap;">
136:                    <div class="card-body p-0" style="max-height:60vh; overflow:auto;">
150:                                    style="max-width:200px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;"
208:                            <div style="overflow-x:auto;">
211:                                        <tr style="background:#f8fafc; position:sticky; top:0; z-index:1; box-shadow:0 1px 0 #e2e8f0;">
262:                                                 style="max-width:480px; max-height:360px; object-fit:contain;
290:                                            <span style="font-weight:600; max-width:260px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;" title="@f.FileName">@f.FileName</span>
315:        position: fixed; inset: 0; background: rgba(0,0,0,.85); z-index: 9500;
318:    .rpt-viewer-img  { max-width: 94vw; max-height: 92vh; object-fit: contain; cursor: auto; }
583:            sb.Append($"<img src=\"data:{mediaType};base64,{base64}\" alt=\"{enc}\" style=\"max-width:480px;max-height:360px;object-fit:contain;border:1px solid #e2e8f0;border-radius:8px;\" />");
```

### JinoSupporter.Web/Components/Pages/NgRatePage.razor

```text
[0 matches]
```

### JinoSupporter.Web/Components/Pages/NgRateAllPage.razor

```text
[0 matches]
```

### JinoSupporter.Web/Components/Pages/NgRateByGroupPage.razor

```text
116:            <div style="overflow-x:auto;">
266:            <div style="overflow-x:auto;">
455:                    <div style="overflow-x:auto;">
```

### JinoSupporter.Web/Components/Pages/NgRateForDailyReportPage.razor

```text
[0 matches]
```

### JinoSupporter.Web/Components/Pages/NgRateForWeeklyReportPage.razor

```text
131:                <div style="overflow-x:auto;">
264:            <div style="overflow-x:auto;">
434:            <div style="overflow-x:auto;">
660:                    <div style="overflow-x:auto;">
```

### JinoSupporter.Web/Components/Pages/SubGroupNode.razor

```text
[0 matches]
```

### JinoSupporter.Web/Components/Shared/NgRateReportStyles.razor

```text
67:        max-height: 150px;
68:        overflow: auto;
88:        height: 100% !important;
90:        max-height: none;
106:        max-height: none;
171:        overflow: hidden;
172:        text-overflow: ellipsis;
262:        position: fixed;
277:        max-height: 80vh;
290:        overflow: auto;
302:        overflow-x: hidden;
305:        overflow-x: auto;
327:        position: sticky;
342:        overflow: hidden;
343:        text-overflow: clip;
354:        overflow: hidden;
355:        text-overflow: ellipsis;
601:        overflow: hidden !important;
602:        text-overflow: ellipsis !important;
622:        overflow: visible !important;
623:        text-overflow: clip !important;
642:        overflow: visible !important;
643:        text-overflow: clip !important;
703:        position: fixed;
717:        max-height: 80vh;
720:        overflow: hidden;
733:        overflow: auto;
754:        max-height: 360px;
755:        overflow: auto;
764:        position: sticky;
789:        overflow-x: auto;
818:        overflow-wrap: anywhere;
```

### JinoSupporter.Web/Components/Shared/NgRateSetupPanel.razor

```text
222:                    <div id="@LogElementId" class="font-monospace p-2 overflow-auto ngrate-log-pane">
```

### JinoSupporter.Web/Components/Shared/NgRateViewNav.razor

```text
[0 matches]
```

### JinoSupporter.Web/Components/Shared/NgRateModelGroupPicker.razor

```text
[0 matches]
```

### JinoSupporter.Web/Components/Shared/NgRateSimpleGroupPicker.razor

```text
[0 matches]
```

### JinoSupporter.Web/Components/Shared/HierSubRows.razor

```text
[0 matches]
```
