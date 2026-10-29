# INSTRUMENT 셸 감사 — Data Inference · Input Data · Tools

감사 기준은 `InstrumentLayout.razor:43-44`의 `height:100vh` 셸, `InstrumentLayout.razor:161-163`의 `main.view > @Body`, `instrument.css:159-164`의 셸 고정 높이/클리핑, `instrument.css:462-463`의 `.view` 스크롤·패딩 계약이다. 지정된 18개 페이지와 2개 공유 컴포넌트의 전체 파일을 읽고 정적 CSS/DOM 관계를 대조했다. 실행 중인 서버와 생성물 `instrument.scoped.css`는 건드리지 않았고, 런타임 렌더링은 하지 않았다.

판정 코드는 `V`=뷰포트(`100vh`) 가정, `P`=`position:fixed`, `H`=`height:100%`/flex 높이 체인 단절, `S`=sticky 상단 오프셋, `O`=자체 overflow에 따른 중첩 스크롤, `Z`=z-index/stacking context, `B`=현재 §24 브리지가 다루지 못하는 페이지 고유 클래스다.

`사용불가`는 정적 근거만으로 조작 불가능 또는 콘텐츠 도달 불가가 확정될 때만 사용했다. 이번 범위에는 그 정도로 확정 가능한 항목은 없었다. `보기흉함`은 이중 스크롤, 셸 위를 덮는 오버레이, 불필요한 세로 여백/스크롤, 작업영역 비율 붕괴를 뜻한다.

## 1. 요약 표

| 라우트 | 파일 | 최고 심각도 | 파손 유형 | 전역 해결 가능 여부 |
|---|---|---:|---|---|
| `/data-inference/ask` | `DataInferenceAskPage.razor` | 보기흉함 | H, V, O, P, B | 부분 — 작업영역 계약은 전역, iframe 높이·B는 페이지 |
| `/data-inference/db-legacy` | `DataInferenceDbPage.razor` | 보기흉함 | H, V, O, P, B | 부분 — 작업영역/모달 정책은 전역, 보고서 높이·B는 페이지 |
| `/data-inference/analysis` | `DataInferenceAnalysisPage.razor` | 보기흉함 | H, O, B | 부분 — H/O는 전역, B는 페이지 |
| `/data-inference/batch` | `DataInferenceBatchPage.razor` | 보기흉함 | H, O, P, B | 부분 — H/O/모달 정책은 전역, B는 페이지 |
| `/data-inference/detail` | `DataInferenceDetailPage.razor` | 보기흉함 | V, P, B | 부분 — 루트 높이/모달 정책은 전역, B는 페이지 |
| `/data-inference/validation` | `DataInferenceValidationPage.razor` | 보기흉함 | H, O, P, B | 부분 — H/O/모달 정책은 전역, B는 페이지·공유 컴포넌트 |
| `/data-inference/model-analysis` | `DataInferenceModelAnalysisPage.razor` | 보기흉함 | V, O, P, B | 부분 — 루트 높이/모달 정책은 전역, B는 페이지 |
| `/data-inference-test` | `DataInferenceInputTestPage.razor` | 보기흉함 | H, O, P, B | 부분 — H/O/모달 정책은 전역, B는 페이지 |
| `/data-inference` | `DataInferencePage_Test.razor` | 보기흉함 | H, O, B | 부분 — H/O는 전역, B는 페이지 |
| `/input-data-batch` | `InputDataBatchPage.razor` | 보기흉함 | V, O, B | 부분 — 루트 높이는 전역, 중첩 패널/B는 페이지 |
| 없음 (`@page` 없음) | `InputDataTestPage.razor` | 보기흉함 | V, O, B | 부분 — 호스트가 작업영역 계약을 주면 높이는 전역, B는 페이지 |
| `/daily-test-data/input` | `DailyTestDataInputPage.razor` | 보기흉함 | V, O, B | 부분 — 루트 높이는 전역, 결과 iframe/중첩 패널/B는 페이지 |
| `/data-inference/first-analysis`, `/data-inference/current-problem` | `CurrentProblemAnalysisPage.razor` | 보기흉함 | V, O, B | 부분 — 루트 높이/단일 스크롤은 전역, B는 페이지 |
| `/tools/graph-maker` | `GraphMakerPage.razor` | 보기흉함 | V, O, B | 부분 — 루트 최소높이는 전역, 시트/차트/B는 페이지 |
| `/data-inference/ai-prompts` | `AiPromptPage.razor` | 보기흉함 | V, O, B | 부분 — 루트 최소높이는 전역, 세부 높이식/B는 페이지 |
| `/translate` | `TranslatePage.razor` | 사소 | O, B | 아니오 — 결과 패널과 고유 클래스 직접 수정 |
| `/tools/pc-download` | `PcDownloadPage.razor` | 사소 | B | 아니오 — 페이지 고유 시각 체계 직접 이식 |
| `/admin/test-excel-converter` | `TestExcelConverterPage.razor` | 사소 | O, B | 아니오 — 로그 패널과 고유 클래스 직접 수정 |

라우트 주의: `DataInferenceDbPage.razor:1`은 실제로 `/data-inference/db-legacy`만 선언하지만 셸 메뉴는 `InstrumentLayout.razor:340`에서 `data-inference/db`를 가리킨다. 허용 범위 밖의 라우트 구현은 확인하지 않았으므로 이 파일을 `/data-inference/db` 화면이라고 단정하지 않는다. `InputDataTestPage.razor`는 파일 전체에 `@page`가 없고 렌더 루트가 `InputDataTestPage.razor:14`의 `<div class="idt-root">`에서 시작한다. 어느 호스트가 사용하는지는 범위 밖이라 확인하지 않았다.

## 2. 화면별 상세

### 2.1 Ask AI — `DataInferenceAskPage.razor`

- 근거(H/O): `DataInferenceAskPage.razor:367-369` — `.diask-root { flex: 1; min-height: 0; display: flex; ... overflow: hidden; }`. 그러나 직접 부모 `.view`는 현재 `instrument.css:462`에서 flex 컨테이너가 아니다. 루트의 `flex:1`은 높이를 받지 못하고 `DataInferenceAskPage.razor:401-402`의 내부 `overflow-y:auto`와 셸 `.view` 스크롤이 경쟁한다.
- 근거(V): `DataInferenceAskPage.razor:428` — `min-height: 680px; height: calc(100vh - 235px);`. 셸 topbar·tabs·`.view` 패딩을 빼지 않아 iframe 호스트가 실제 가용 높이보다 커진다. 자식도 `DataInferenceAskPage.razor:433`에서 `height: 100%; min-height: 680px`를 강제한다.
- 근거(P): `DataInferenceAskPage.razor:717` — `position: fixed; inset: 0; ... z-index: 9000;`. 확인 대화상자가 페이지 영역이 아니라 rail/topbar를 포함한 브라우저 전체를 덮는다.
- 근거(B): `DataInferenceAskPage.razor:371-373` — `.diask-header { ... background: #fff; border-bottom: 1px solid #e2e8f0; }`. 현재 브리지는 `instrument.css:1103-1143`의 Bootstrap 카드/폼 클래스만 재매핑하므로 `diask-*`의 하드코딩 색·radius는 그대로다. 공유 `CurrentProblemWorkflowStrip`도 `DataInferenceAskPage.razor:42`에서 들어온다.
- 증상: 답변/이력 영역과 셸에 각각 스크롤이 생기고 HTML 보고서가 과도하게 길며, 삭제 확인은 셸 전체를 가린다.
- 심각도: **보기흉함**.
- 수정 구분: **전역** — 이 라우트를 `view--flush` 작업영역으로 지정하고 `.view` 높이 체인을 복구. **페이지** — iframe을 `100vh`가 아니라 작업영역/콘텐츠 기준으로 재설계하고 `diask-*`를 토큰화. **공유 페이지** — `CurrentProblemWorkflowStrip.razor`의 고유 스타일 이식.

### 2.2 DB Legacy — `DataInferenceDbPage.razor`

- 근거(H/O): `DataInferenceDbPage.razor:2010-2012` — `.didb-root { flex: 1; min-height: 0; ... overflow: hidden; }`; `DataInferenceDbPage.razor:2140` — `.didb-table-wrap { flex: 1; overflow: auto; }`. 비-flex `.view` 아래에서 루트 높이 계약이 끊겨 외부/내부 스크롤이 함께 생긴다.
- 근거(V/O): `DataInferenceDbPage.razor:2067-2070` — 보고서 frame에 `height: min(900px, 58vh); min-height: 520px`; `DataInferenceDbPage.razor:2204-2207` — 다른 frame에 `height: min(1300px, 72vh); min-height: 720px`; `DataInferenceDbPage.razor:3094-3096`과 `3544-3545`에도 별도 세로 스크롤 패널이 있다.
- 근거(P): `DataInferenceDbPage.razor:3610` — 이미지 뷰어 `position: fixed; ... z-index: 9500`; `DataInferenceDbPage.razor:3649` — 삭제 확인 `position: fixed; ... z-index: 9000`. OCR 편집기도 `DataInferenceDbPage.razor:1983`에서 공유 fixed overlay를 호출한다.
- 근거(B): `DataInferenceDbPage.razor:2014-2016` — `.didb-header { ... background: #fff; border-bottom: 1px solid #e2e8f0; }`; `DataInferenceDbPage.razor:2056-2060`의 고유 chip 스타일도 브리지 대상이 아니다.
- 증상: 데이터/AI 보고서/미리보기 각각에 스크롤이 생기고 긴 frame 때문에 셸 스크롤이 추가되며, 뷰어·확인·OCR가 rail까지 덮는다.
- 심각도: **보기흉함**.
- 수정 구분: **전역** — flush 작업영역과 overlay 정책. **페이지** — 두 iframe의 vh/min-height 제거, 다수의 중첩 스크롤 책임 정리, `didb-*` 이식. **공유 페이지** — `OcrTextEditor.razor` 이식.

### 2.3 Aggregated Results — `DataInferenceAnalysisPage.razor`

- 근거(H/O): `DataInferenceAnalysisPage.razor:165-167` — `.dian-root { flex: 1; min-height: 0; ... overflow: hidden; }`; `DataInferenceAnalysisPage.razor:186-192` — 내부 workspace/pane도 `flex:1`, `min-height:0`, `overflow:hidden`을 전제로 한다. 현재 `.view`가 이 높이를 제공하지 않는다.
- 근거(H/O): `DataInferenceAnalysisPage.razor:206` — `.dian-tags-list, .dian-file-list { height: calc(100% - 42px); overflow: auto; }`; `DataInferenceAnalysisPage.razor:261` — `.dian-table-preview { flex: 1; min-height: 0; overflow: auto; }`. 부모 높이가 확정되지 않으면 `100%` 계산과 내부 스크롤 경계가 풀린다.
- 근거(B): `DataInferenceAnalysisPage.razor:169-185` — 헤더/stat이 `#fff`, `#dfe5ee`, pill radius/청색을 직접 지정한다.
- 증상: 3-pane가 셸 가용 높이에 고정되지 않고 목록 또는 Excel pane이 바깥 `.view`와 함께 길어져 작업영역 비율이 무너진다.
- 심각도: **보기흉함**.
- 수정 구분: **전역** — flush flex 높이 체인으로 H/O 복구. **페이지** — `dian-*` 색·panel·pill을 INSTRUMENT 토큰/컴포넌트로 이식.

### 2.4 Batch — `DataInferenceBatchPage.razor`

- 근거(H/O): `DataInferenceBatchPage.razor:360-362` — `.dib-root { flex: 1; min-height: 0; ... overflow: hidden; }`; `DataInferenceBatchPage.razor:388-390` — content/table에 다시 `flex:1`과 `overflow:auto`가 이어진다.
- 근거(P): `DataInferenceBatchPage.razor:504-505` — `.dib-overlay { position: fixed; inset: 0; ... z-index: 9000; }`; modal은 `DataInferenceBatchPage.razor:511`에서 `max-height: 80vh`를 사용한다.
- 근거(B): `DataInferenceBatchPage.razor:374-384` — `.dib-tabs/.dib-tab`이 pill radius와 고정 청색을 직접 선언한다.
- 증상: 표/매핑 패널과 셸이 별도로 스크롤되고 오류 modal이 셸 chrome 전체를 가린다.
- 심각도: **보기흉함**.
- 수정 구분: **전역** — flush flex 계약과 overlay 범위/레이어 정책. **페이지** — `dib-*` tabs/map 패널 이식.

### 2.5 Detail — `DataInferenceDetailPage.razor`

- 근거(V): `DataInferenceDetailPage.razor:557-559` — `.didet-root { min-height: 100vh; ... padding: 20px 24px; }`. `.view` 자체 패딩까지 더해져 항상 실제 stage보다 높은 최소 높이가 된다.
- 근거(P): `DataInferenceDetailPage.razor:925-929` — `.didet-viewer { position: fixed; inset: 0; ... z-index: 9500; }`, 이미지도 `max-height: 92vh`다.
- 근거(B): `DataInferenceDetailPage.razor:562-580` — header/card가 흰 배경, 10px radius, 고정 slate 테두리를 직접 쓴다.
- 증상: 짧은 상세도 불필요한 셸 세로 스크롤이 생기고 이미지 뷰어가 rail/topbar까지 덮는다.
- 심각도: **보기흉함**.
- 수정 구분: **전역** — INSTRUMENT 안에서 `.didet-root`의 viewport 최소높이를 해제하고 공통 overlay 정책 적용. **페이지** — 상세 카드·chip·viewer 콘텐츠 크기를 토큰 기반으로 이식.
- S 비파손 확인: `DataInferenceDetailPage.razor:643`의 `position: sticky; top: 12px`는 바깥 `.view` 스크롤포트 안의 side panel이라 topbar 높이를 더할 대상이 아니다. `top`을 전역 가산하면 오히려 틀어진다.

### 2.6 Validation — `DataInferenceValidationPage.razor`

- 근거(H/O): `DataInferenceValidationPage.razor:299` — `.div-root { flex:1; min-height:0; ... overflow:hidden; }`; `DataInferenceValidationPage.razor:309` — `.div-table-wrap { flex:1; overflow:auto; }`.
- 근거(P): `DataInferenceValidationPage.razor:333-339` — viewer와 close가 각각 `position:fixed`, overlay는 `z-index:9999`; `DataInferenceValidationPage.razor:290`에서 `OcrTextEditor`도 호출한다.
- 근거(B): `DataInferenceValidationPage.razor:300-316` — header/table head가 `div-*` 고유 selector와 고정 배경/색을 사용한다.
- 증상: 검증 표가 셸과 중첩 스크롤되고 이미지/OCR overlay가 셸 전체를 가린다.
- 심각도: **보기흉함**.
- 수정 구분: **전역** — flush flex 계약과 overlay 정책. **페이지** — `div-*` 표/상태 이식. **공유 페이지** — OCR dialog 이식.

### 2.7 Each Model Analysis — `DataInferenceModelAnalysisPage.razor`

- 근거(V/O): `DataInferenceModelAnalysisPage.razor:1216-1224` — `.dima-root { height: 100vh; max-height: 100vh; ... overflow: hidden; }`. 셸 안에서도 브라우저 viewport 전체 높이를 다시 소유한다.
- 근거(O): `DataInferenceModelAnalysisPage.razor:1263-1268`의 body `overflow:hidden`, `DataInferenceModelAnalysisPage.razor:1326-1334`의 group/report `overflow:auto`, `DataInferenceModelAnalysisPage.razor:1586-1588`의 분석 pane `overflow:auto`가 외부 `.view` 스크롤과 겹친다.
- 근거(P): `DataInferenceModelAnalysisPage.razor:1680-1693` — modal backdrop `position:fixed; z-index:1060`, modal은 `height: min(840px, 92vh)`.
- 근거(B): `DataInferenceModelAnalysisPage.razor:1234-1241` — `.dima-header`가 흰 배경과 고정 slate border를 직접 선언한다.
- 증상: 화면 전체가 셸보다 한 topbar+tabs만큼 더 길어 바깥/목록/분석 pane 스크롤을 번갈아 써야 하고 modal은 chrome 위에 뜬다.
- 심각도: **보기흉함**.
- 수정 구분: **전역** — 작업영역 라우트에서 root를 `height:100%`로 치환하고 overlay 정책 통일. **페이지** — 세 pane의 유일한 세로 스크롤 소유자를 정하고 `dima-*` 이식.

### 2.8 Input Test Queue — `DataInferenceInputTestPage.razor`

- 근거(H/O): `DataInferenceInputTestPage.razor:334-350` — `.di-root`와 queue/main pane이 flex+`overflow:hidden`인데 직접 부모 `.view`에서 높이를 받지 못한다. queue는 `DataInferenceInputTestPage.razor:384-386`, 결과는 `DataInferenceInputTestPage.razor:719-721`에서 각각 `overflow-y:auto`다.
- 근거(P): `DataInferenceInputTestPage.razor:439-446` — alias overlay가 `position:fixed; z-index:1500`, modal은 `max-height:85vh`다.
- 근거(B): `DataInferenceInputTestPage.razor:343-362` — queue pane/head/count가 고유 `dit-*`와 고정 white/indigo/slate 스타일을 쓴다.
- 증상: queue와 결과 pane 높이가 stage가 아니라 콘텐츠에 끌려가고, alias modal은 셸 전체를 덮는다.
- 심각도: **보기흉함**.
- 수정 구분: **전역** — flush flex 계약과 overlay 정책. **페이지** — queue/Excel/결과 고유 스타일 이식.

### 2.9 Data Inference Test — `DataInferencePage_Test.razor`

- 근거(H/O): `DataInferencePage_Test.razor:221-223` — `.di-root { flex:1; min-height:0; ... overflow:hidden; }`; `DataInferencePage_Test.razor:343-344` — 결과가 `flex:1; ... overflow-y:auto`다.
- 근거(B): `DataInferencePage_Test.razor:225-236`과 `251-270` — topbar/status/textarea/drop-zone을 고유 selector와 고정 slate/blue/radius로 선언한다.
- 증상: 결과 영역이 stage 높이에 고정되지 않아 셸 스크롤로 풀리고, 시각적으로 Classic UI 카드가 INSTRUMENT 안에 중첩된다.
- 심각도: **보기흉함**.
- 수정 구분: **전역** — flush flex 계약. **페이지** — `di-*` 입력·drop-zone·결과 이식.

### 2.10 Input Data Batch — `InputDataBatchPage.razor`

- 근거(V): `InputDataBatchPage.razor:2196-2200` — `.input-data-root { min-height: calc(100vh - 24px); padding:18px; }`. `.view`의 18/20/40px 패딩과 topbar/tabs를 고려하지 않는다.
- 근거(O): `InputDataBatchPage.razor:2602-2605` — recognized data가 `max-height:420px; overflow:auto`; `InputDataBatchPage.razor:2820-2835` — JSON/log에도 각각 제한 높이와 `overflow:auto`가 있다.
- 근거(B/충돌): `InputDataBatchPage.razor:2231-2236`이 전역명 `.panel`에 radius 8px/padding 14px/흰 배경을 선언한다. 이는 INSTRUMENT의 동명 `.panel`(`instrument.css:480-490`)과 한 요소에 합성되어 둥근 padded panel과 bezel corner tick이 섞인다. 또한 `InputDataBatchPage.razor:2381-2405`는 `table`, `th`, `td`를 페이지 root로 scope하지 않는다.
- 증상: 페이지 최소높이 때문에 바깥 스크롤이 상시 생기고, 결과/JSON/log 안에도 스크롤이 생기며 INSTRUMENT panel과 페이지의 둥근 panel 규칙이 혼합된다.
- 심각도: **보기흉함**.
- 수정 구분: **전역** — INSTRUMENT 안에서 root의 viewport 최소높이를 해제. **페이지** — `.panel/table/th/td/h1`을 `.input-data-root` 아래로 namespace하고 결과 스크롤 정책과 디자인 토큰을 직접 이식. 이 B 항목은 브리지 확장보다 페이지 수정이 안전하다.

### 2.11 Input Data Test component — `InputDataTestPage.razor`

- 근거(V/O): `InputDataTestPage.razor:368-375` — `.idt-root { height: calc(100vh - 24px); ... overflow:hidden; }`; 좌측은 `InputDataTestPage.razor:384-390`에서 `overflow-y:auto`, step은 `InputDataTestPage.razor:626-628`에서 다시 `overflow-y:auto`다.
- 근거(V/O): `InputDataTestPage.razor:897-900` — HTML frame `height: min(1200px, calc(100vh - 230px))`; 작은 화면 media 규칙도 `InputDataTestPage.razor:1084-1088`에서 `min-height: calc(100vh - 24px); overflow:auto`를 유지한다.
- 근거(B): `InputDataTestPage.razor:378-390` — left/step panel이 고정 white/slate/radius를 직접 선언한다.
- 증상: 어느 페이지에 포함되든 viewport 기준 root와 frame이 호스트의 실제 높이를 무시하고 좌측/step/호스트 스크롤이 겹친다.
- 심각도: **보기흉함**.
- 수정 구분: **전역** — 호스트가 이 컴포넌트에 확정 높이를 주는 경우 root를 `height:100%`로 덮을 수 있다. **페이지** — `@page`가 없는 재사용 컴포넌트이므로 최종적으로는 host-size 기반 CSS와 `idt-*` 이식이 필요하다.

### 2.12 Daily Test Data Input — `DailyTestDataInputPage.razor`

- 근거(V/O): `DailyTestDataInputPage.razor:308-315` — `.daily-root { height: calc(100vh - 24px); ... overflow:hidden; }`; 내부 목록은 `DailyTestDataInputPage.razor:333-354`, main은 `DailyTestDataInputPage.razor:405`, 결과는 `DailyTestDataInputPage.razor:630-645`에서 각각 스크롤한다.
- 근거(V): `DailyTestDataInputPage.razor:707-713` — generated HTML `max-height:78vh`, 결과 panel `min-height:72vh`; 생성 iframe 문자열도 `DailyTestDataInputPage.razor:3615`에서 `min-height:72vh`를 직접 넣는다.
- 근거(B): `DailyTestDataInputPage.razor:317-345` — sidebar/item 카드가 고정 white/slate/blue/radius를 직접 선언한다.
- 증상: 셸 스크롤, main 스크롤, history/result 스크롤이 겹치고 생성 보고서가 실제 stage보다 크게 잡힌다.
- 심각도: **보기흉함**.
- 수정 구분: **전역** — root를 flush 작업영역의 `height:100%`로 치환. **페이지** — history/main/result 중 스크롤 소유자를 정하고 생성 iframe의 `72vh` 및 `daily-*` 스타일을 직접 이식.

### 2.13 Current Problem Analysis — `CurrentProblemAnalysisPage.razor`

- 근거(V/O): `CurrentProblemAnalysisPage.razor:164-170` — `.fav-root { height:100vh; ... overflow:hidden; }`; main도 `CurrentProblemAnalysisPage.razor:199-205`에서 overflow를 숨기고 표 wrapper가 `CurrentProblemAnalysisPage.razor:241-243`에서 `overflow:auto`다.
- 근거(B): `CurrentProblemAnalysisPage.razor:174-183` — header가 고정 높이/흰 배경/slate border를 직접 선언한다.
- 증상: stage보다 큰 root를 `.view`가 다시 스크롤하고 그 안에서 표를 다시 스크롤해야 한다.
- 심각도: **보기흉함**.
- 수정 구분: **전역** — 두 라우트를 flush 작업영역으로 지정하고 root를 `height:100%`로 치환. **페이지** — `fav-*` header/table 이식.

### 2.14 Graph Maker — `GraphMakerPage.razor`

- 근거(V): `GraphMakerPage.razor:1220-1224` — `.gm-page { min-height: calc(100vh - 32px); ... padding:18px; }`로 stage가 아닌 viewport를 기준으로 한다.
- 근거(O): `GraphMakerPage.razor:1473-1476` — `.gm-sheet { overflow:auto; max-height:360px; }`가 외부 `.view`와 중첩된다.
- 근거(B): `GraphMakerPage.razor:1262-1269` 이후의 `gm-panel/gm-workspace` 체계와 `GraphMakerPage.razor:1220-1224`의 hard-coded 배경/색은 브리지 클래스가 아니다.
- 증상: 짧은 그래프 화면에도 바깥 스크롤이 생기고 데이터 시트는 별도 스크롤을 요구하며 패널 외형은 Classic 계열로 남는다.
- 심각도: **보기흉함**.
- 수정 구분: **전역** — root viewport 최소높이 해제. **페이지** — chart/table의 스크롤 책임과 `gm-*` 이식.

### 2.15 AI Prompts — `AiPromptPage.razor`

- 근거(V/O): `AiPromptPage.razor:88-93` — `.aip-root { min-height: calc(100vh - 24px); ... }`; 목록은 `AiPromptPage.razor:177-179`, preview는 `AiPromptPage.razor:229-231`, code는 `AiPromptPage.razor:290-294`에서 각각 `100vh` 계산과 overflow를 갖는다.
- 근거(B): `AiPromptPage.razor:125-135` — stat card가 고정 white/slate/radius를 직접 선언한다.
- 증상: 목록·코드와 셸을 따로 스크롤해야 하고 preview의 최소높이가 실제 stage보다 커진다.
- 심각도: **보기흉함**.
- 수정 구분: **전역** — root 최소높이 해제. **페이지** — list/preview/code를 grid의 `minmax(0,1fr)`와 단일 패널 스크롤로 재구성하고 `aip-*` 이식.

### 2.16 Translate — `TranslatePage.razor`

- 근거(O): `TranslatePage.razor:116` — 결과 `.card-body`에 인라인 `overflow-y:auto; max-height:72vh`를 지정해 `.view`와 별도 스크롤을 만든다.
- 근거(B): `TranslatePage.razor:149-173` — `img-*` strip/thumb/clear가 고정 slate/white/radius를 사용하고, `TranslatePage.razor:176-196`의 source/usage panel도 고유 selector다. Bootstrap card 자체만 §24 브리지의 적용을 받는다.
- 증상: 일반 문서형 페이지 안에서 결과 카드만 별도 스크롤되고 이미지/OCR 보조 UI는 새 디자인과 시각적으로 분리된다.
- 심각도: **사소**.
- 수정 구분: **페이지** — 결과 높이는 콘텐츠 또는 stage 기반으로 바꾸고 `img-*`, `translate-*`를 토큰화.

### 2.17 PC Download — `PcDownloadPage.razor`

- 근거(B): `PcDownloadPage.razor:107-120` — `.pcd-card/.pcd-btn`이 8~10px radius, 고정 white/blue/slate 값을 직접 사용한다. `pcd-*`는 Bootstrap bridge 클래스가 아니다.
- 증상: 셸 기능은 깨지지 않지만 다운로드 카드/버튼만 둥근 Classic 스타일로 남아 한 화면 안에 두 디자인 시스템이 보인다.
- 심각도: **사소**.
- 수정 구분: **페이지** — `pcd-*`를 INSTRUMENT panel/button/token으로 이식. 셸 전역 수정으로 의미 있는 해결은 불가능하다.

### 2.18 Test Excel Converter — `TestExcelConverterPage.razor`

- 근거(O): `TestExcelConverterPage.razor:162-164` — 로그 영역이 `overflow-auto`와 `height:380px`를 인라인 지정해 바깥 `.view`와 별도 스크롤한다.
- 근거(B): `TestExcelConverterPage.razor:187-202` — `.tec-drop-zone`이 고정 blue/slate/8px radius를 사용한다. 카드/폼은 bridge되지만 `tec-*`는 그대로다.
- 증상: 기능은 유지되나 로그에 별도 스크롤이 있고 drop zone만 Classic 계열 외형으로 남는다.
- 심각도: **사소**.
- 수정 구분: **페이지** — 로그 패널과 `tec-*` drop zone을 직접 이식.

### 2.19 공유 컴포넌트

#### `OcrTextEditor.razor`

- 소비 근거: `DataInferenceDbPage.razor:1983`, `DataInferenceValidationPage.razor:290`.
- 근거(P/O): `OcrTextEditor.razor:75-82` — `.ocr-overlay { position:fixed; ... z-index:9600 }`, dialog `height:85vh`; `OcrTextEditor.razor:95` — body `overflow:auto`.
- 근거(B): `OcrTextEditor.razor:79-88` — dialog/head가 고정 흰 배경, 10px radius, slate 색을 쓴다.
- 증상: OCR를 열면 두 소비 화면 모두에서 rail/topbar까지 가리고 dialog 내부에 별도 스크롤이 생긴다.
- 심각도: **보기흉함**. 수정 구분: overlay 범위/레이어는 **전역 정책**, dialog 크기·토큰은 **공유 컴포넌트 직접 수정**.

#### `CurrentProblemWorkflowStrip.razor`

- 소비 근거: `DataInferenceAskPage.razor:42`, `DataInferenceDbPage.razor:22`.
- 근거(B): `CurrentProblemWorkflowStrip.razor:67-78` — `.cp-strip`이 고정 margin/padding/radius/white/slate를 사용하고, chip도 `CurrentProblemWorkflowStrip.razor:97-107`에서 별도 고정 스타일이다.
- 증상: Ask/DB 상단에 새 셸과 다른 둥근 흰 strip이 반복 노출된다.
- 심각도: **사소**. 수정 구분: **공유 컴포넌트 직접 수정** 한 번으로 두 화면을 해결한다.

### 2.20 S와 Z의 음성 판정

- **S는 현재 확인되지 않았다.** 표 sticky는 대부분 자기 스크롤 컨테이너 안에 있다. 예: `DataInferenceDbPage.razor:2140`의 `overflow:auto` 안에 `DataInferenceDbPage.razor:2147`의 `position:sticky; top:0`; `GraphMakerPage.razor:1475-1476` 안에 `GraphMakerPage.razor:1508-1510`; `CurrentProblemAnalysisPage.razor:241-243` 안에 `CurrentProblemAnalysisPage.razor:277-279`. 이들은 topbar가 아니라 로컬 표 상단에 붙어야 하므로 전역 topbar offset을 더하면 안 된다.
- **Z 충돌도 현재 정적 근거로 확인되지 않았다.** shell/rail/topbar는 `instrument.css:159-190`, `373-383`에서 별도 z-index나 transform stacking context를 만들지 않는다. 페이지 overlay의 z-index는 1060~9999다. 따라서 P(셸 전체를 덮는 범위)는 있지만 “셸 뒤로 숨는” Z 파손은 입증되지 않았다. 이후 shell에 `isolation`/z-index를 추가한다면 modal layer token과 함께 재검증해야 한다.

## 3. 전역으로 해결 가능한 항목 모음

### G1. `.view`에 두 가지 명시적 계약을 둔다

반복 근거는 `instrument.css:462-463`이다. 현재 레이아웃은 모든 라우트에 스크롤+패딩 문서형 `.view`를 주지만, 이미 `.view--flush`에는 `padding:0; overflow:hidden; display:flex; flex-direction:column`이 정의되어 있다. 다음 구현에서는 `InstrumentLayout.razor:161`이 라우트 성격에 따라 둘 중 하나를 선택해야 한다.

1. **문서형**: `.view`가 유일한 세로 스크롤 소유자. Detail, Input Data Batch, Graph Maker, AI Prompts, Translate, PC Download, Test Excel Converter처럼 콘텐츠가 아래로 흐르는 화면.
2. **작업영역형**: `.view.view--flush`가 스크롤하지 않고 직계 root에 확정 높이를 제공. Ask, DB legacy, Aggregated Analysis, Batch, Validation, Input Test Queue, `/data-inference`, Model Analysis, Daily Input, Current Problem.

구현 시 `.stage`, `.view`에 `min-height:0; min-width:0`도 명시한다. 작업영역형 적용은 CSS class 이름 추측보다 `CurrentRelative()` 기반 route metadata/allow-list가 안전하다. 이 한 단계가 `.diask-root`, `.didb-root`, `.dian-root`, `.dib-root`, `.div-root`, `.di-root`의 반복 H/O를 해결한다.

### G2. INSTRUMENT 안에서 viewport root를 stage root로 치환한다

반복 대상은 다음과 같다.

- `DataInferenceModelAnalysisPage.razor:1217-1218` — `height/max-height:100vh`
- `InputDataTestPage.razor:369`, `DailyTestDataInputPage.razor:309` — `height:calc(100vh - 24px)`
- `CurrentProblemAnalysisPage.razor:165` — `height:100vh`
- `DataInferenceDetailPage.razor:558`, `InputDataBatchPage.razor:2197`, `GraphMakerPage.razor:1221`, `AiPromptPage.razor:89` — viewport 기반 `min-height`

첫 구현 slice에서는 `instrument.css` BRIDGE에 `.ins` 한정 selector로 작업영역 root를 `height:100%; max-height:100%; min-height:0`로, 문서형 root는 `min-height:0`로 덮을 수 있다. Classic UI에는 영향이 없다. 장기적으로는 각 페이지에서 viewport 식을 제거해야 한다.

단, Ask/DB/InputDataTest/Daily의 내부 iframe·보고서 식(`DataInferenceAskPage.razor:428`, `DataInferenceDbPage.razor:2069,2206`, `InputDataTestPage.razor:899`, `DailyTestDataInputPage.razor:3615`)은 의미가 서로 달라 일괄 override하지 말고 페이지별로 고친다.

### G3. overlay의 “셸 전체” 대 “stage 내부” 정책을 먼저 정한다

반복 fixed selector는 Ask 9000, DB 9000/9500, Batch 9000, Detail 9500, Validation 9999, Model 1060, Input Test 1500, OCR 9600으로 제각각이다. 전역 구현 입력은 다음과 같다.

- `--layer-modal`, `--layer-lightbox` 토큰을 정의해 z-index 숫자를 통일한다.
- modal이 앱 전체를 막아야 한다면 `position:fixed`는 유지하되 rail/topbar를 덮는 것이 명시된 정책이어야 한다.
- stage 내부 modal이어야 한다면 `.stage { position:relative; isolation:isolate; }`만 추가하고 끝내지 말 것. fixed를 absolute로 바꾸면 중간 `overflow:hidden` ancestor에 잘릴 수 있으므로 overlay host/portal을 `InstrumentLayout`에 두고 페이지가 그 host를 사용해야 한다.
- 현재는 Z 충돌이 없으므로 z-index만 더 크게 올리는 수정은 하지 않는다.

### G4. sticky offset 전역 보정은 하지 않는다

현재 S 파손 근거가 없고 표 sticky는 로컬 스크롤 컨테이너 기준이다. `top: var(--topbar-h)` 같은 전역 rule은 DB/Graph/Current Problem의 table header를 컨테이너 안에서 아래로 밀어 새 버그를 만든다. 향후 `.view` 자체에 붙는 sticky page header가 발견될 때만 해당 selector에 한정한다.

### G5. B는 “selector 추가”가 아니라 페이지/공유 컴포넌트 이식으로 처리한다

§24는 `instrument.css:1055-1067`이 설명하듯 Bootstrap/app.css 호환 bridge이고, 실제 mapping도 card/form/button/table/alert/badge/nav 중심(`instrument.css:1103-1243`)이다. `diask-*`, `didb-*`, `dian-*`, `dib-*`, `didet-*`, `div-*`, `dima-*`, `dit-*`, `idt-*`, `daily-*`, `fav-*`, `gm-*`, `aip-*`, `pcd-*`, `tec-*`, `ocr-*`, `cp-*`의 수백 개 hard-coded 선언을 `!important` bridge로 복제하면 유지보수 불가능하다.

페이지 이식 공통 규칙은 다음과 같다.

- hard-coded `#fff/#f8fafc/#e2e8f0/#2563eb`와 radius를 INSTRUMENT surface/line/signal 토큰으로 교체한다.
- 페이지 안 `<style>`을 root-prefixed 또는 `.razor.css`로 옮긴다. 특히 `InputDataBatchPage.razor:2231-2405`의 `.panel`, `table`, `th`, `td`, `h1` 전역 selector를 먼저 격리한다.
- panel/list/table의 세로 스크롤은 화면당 주 작업영역 하나를 원칙으로 하고, 작은 log/code/preview만 제한된 보조 스크롤로 남긴다.
- 공유 `OcrTextEditor`와 `CurrentProblemWorkflowStrip`을 먼저 이식하면 DB/Validation, Ask/DB에 각각 재사용된다.

## 4. 이식 우선순위

실사용 telemetry는 이 범위에서 확인할 수 없었다. 따라서 사용빈도는 허용된 참조의 셸 직접 메뉴 노출을 대리 지표로 썼다. 직접 메뉴 근거는 `InstrumentLayout.razor:338-348`이며 Graph Maker, Input Data Batch, Ask, Daily Input, Translate, PC Download, Test Excel Converter가 포함된다. DB 메뉴는 `/data-inference/db`라 이 감사 파일의 `/db-legacy`와 동일 화면인지 불명이다. 아래는 **페이지 직접 수정이 필요한 화면만**의 순서다.

1. **DataInferenceAskPage** — 직접 메뉴, H/V/O/P가 동시에 있고 핵심 AI 결과 iframe이 stage 높이를 무시한다. `CurrentProblemWorkflowStrip`도 함께 이식.
2. **DailyTestDataInputPage** — 직접 메뉴, V/O가 root·history·main·result·생성 iframe에 반복되어 스크롤 비용이 가장 크다.
3. **InputDataBatchPage** — 직접 메뉴, V/O뿐 아니라 전역 `.panel/table/th/td`가 INSTRUMENT CSS와 충돌한다. namespace 수정이 다른 페이지 오염 위험도 낮춘다.
4. **GraphMakerPage** — 직접 메뉴, V/O가 있고 chart+sheet 복합 작업 화면이라 세로 공간 손실의 체감이 크다.
5. **DataInferenceModelAnalysisPage** — 메뉴 직접 노출은 확인 못했지만 `100vh` root와 다중 pane, fixed modal이 결합된 가장 복잡한 작업영역이다.
6. **InputDataTestPage** — 라우트/호스트는 미확인이나 V/O가 root·frame·responsive 규칙까지 중복된다. 재사용 컴포넌트 계약부터 고쳐야 한다.
7. **CurrentProblemAnalysisPage** — 두 route가 있고 `100vh` root+내부 table scroll이 명확하다.
8. **DataInferenceDbPage (`/db-legacy`)** — H/V/O/P와 OCR 공유 modal이 모두 있으나 셸의 `/data-inference/db`와 연결 여부를 먼저 확인해야 한다. 확인되면 1~3순위 수준으로 올린다.
9. **DataInferenceDetailPage** — deep-link 성격으로 보이며 V/P가 있다. 상세 card 수가 많아 B 이식량은 크다.
10. **DataInferenceAnalysisPage** — H/O 3-pane를 전역 계약으로 먼저 살린 뒤 고유 panel/list를 이식한다.
11. **DataInferenceBatchPage** — H/O/P 작업영역. 전역 flush 적용 후 tabs/map 패널 이식.
12. **DataInferenceValidationPage** — H/O/P. `OcrTextEditor`를 먼저 고치면 modal 이슈 일부가 함께 해결된다.
13. **DataInferenceInputTestPage** — queue/results H/O/P. 테스트 route라 직접 메뉴 화면보다 뒤에 둔다.
14. **DataInferencePage_Test** — `/data-inference`지만 셸 메뉴 직접 노출은 확인되지 않았고 H/O/B만 있다.
15. **AiPromptPage** — V/O/B지만 셸 직접 메뉴 노출은 확인되지 않았다.
16. **TranslatePage** — 직접 메뉴지만 기능 파손은 없고 O/B가 국소적이다.
17. **TestExcelConverterPage** — 직접 메뉴라도 admin/test 화면이며 O/B가 log/drop-zone에 한정된다.
18. **PcDownloadPage** — 직접 메뉴지만 B만 있고 기능/스크롤 파손은 없다.

공유 컴포넌트 선행 순서는 **`OcrTextEditor.razor` → `CurrentProblemWorkflowStrip.razor`**다. 전자는 P/O/B를 두 소비 화면에서 제거하고, 후자는 B를 Ask/DB 두 화면에서 한 번에 제거한다.

## 결론

첫 구현은 페이지 18개를 동시에 고치는 작업이 아니다. `InstrumentLayout`의 route-aware 문서형/작업영역형 `.view` 계약(G1), INSTRUMENT 한정 viewport-root 치환(G2), modal 정책 결정(G3)을 한 전역 slice로 적용하면 H와 root 수준 V/O의 대부분을 제거할 수 있다. 그 다음 Ask → Daily → Input Data Batch 순으로 내부 iframe/중첩 overflow와 B를 페이지 직접 이식하는 것이 가장 안전하다. sticky top offset과 z-index 증폭은 이번 근거상 전역 수정 대상이 아니다.
