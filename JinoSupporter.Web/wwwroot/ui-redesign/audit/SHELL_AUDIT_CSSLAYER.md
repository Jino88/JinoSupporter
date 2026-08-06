# SHELL AUDIT — CSS LAYER

INSTRUMENT 셸(`542e595`, "화면 새 UI를 레이아웃으로 승격")의 **CSS 계층 전용 읽기 감사**.
대상은 `.ins` 스코프 설계, `tools/scope-css.js` 생성기, Bootstrap 브릿지, 그리고
새 셸이 호스팅하는 40개 페이지의 인라인 `<style>` 블록이 만드는 캐스케이드다.

| | |
|---|---|
| Branch / Commit | `codex/s21/claude-session-21-12a6abd3` / `542e595` |
| 감사 유형 | 읽기 전용 정적 분석 (빌드·실행·브라우저 렌더 없음) |
| 산출물 | 이 파일 하나. 코드 변경 없음 |
| Agent | Claude (worker session, role: analysis) |

---

## 0. 한 줄 요약

스코핑은 **바깥을 지키는 데는 성공했고, 안을 지키는 데는 실패했다.**
`.ins`는 디자인 시스템이 클래식 셸을 오염시키는 것을 막지만, 새 셸이 레이아웃으로 승격되면서
**모든 페이지가 `.ins` 안으로 들어왔다.** 그래서 `.ins button` / `.ins table` / `.ins a` /
`.ins input` 같은 element 리셋(specificity `(0,1,1)`)이 페이지 로컬 클래스(`(0,1,0)`)를
전부 이겨버린다. 정적으로 확인된 것만 **페이지 CSS 규칙 242개**가 조용히 무력화된다.

---

## 1. 현재 CSS 계층이 조립되는 순서

문서 순서 = 캐스케이드 tie-break 순서. 아래에서 **뒤쪽이 동일 specificity에서 이긴다.**

| # | 위치 | 파일 | 규모 | 스코프 |
|---|---|---|---|---|
| 1 | `<head>` (`App.razor:8`) | `bootstrap/bootstrap.min.css` | 159 KB | 전역 |
| 2 | `<head>` (`App.razor:9`) | `app.css?v=2` | 351줄 / `!important` 117 | 전역 |
| 3 | `<head>` (`App.razor:10`) | `JinoSupporter.Web.styles.css` | 컴포넌트 2개분 | `[b-xxxxx]` 스코프 |
| 4 | `<head>` (`App.razor:11-12`) | Syncfusion `bootstrap5.css` × 2 | — | 전역 |
| 5 | **`<body>`** (`InstrumentLayout.razor:31`) | `ui-redesign/assets/instrument.scoped.css?v=3` | 1,249줄 | `.ins` 하위 |
| 6 | **`<body>`, @Body 내부** | 페이지 `<style>` 블록 40개 | **11,075줄** / `!important` 107 / hex 2,641 | **전역(스코프 없음)** |

핵심 사실 두 가지:

* **5번은 `<head>`가 아니라 `<body>`에 있다.** `InstrumentLayout.razor:31`의 `<link>`는
  컴포넌트 본문에서 렌더되므로 `@Body`보다 앞, 나머지 `<head>` 시트보다 뒤에 놓인다.
* **6번은 스코프가 없다.** `.razor.css`(Blazor scoped CSS)를 쓰는 컴포넌트는
  `MainLayout`, `NavMenu` 둘뿐이다. 나머지 40개 페이지는 `<style>` 블록을 문서에 그대로 주입한다.
  즉 페이지 CSS는 **디자인 시스템보다 뒤에 로드되는 전역 시트**다.

```
bootstrap  <  app.css  <  scoped(.razor.css)  <  syncfusion  <  instrument.scoped  <  page <style>
                                                                   └ .ins 하위       └ 전역, 11,075줄
```

---

## 2. 심각도별 발견 사항

### S1 — `.ins button` 리셋이 페이지 로컬 버튼 스타일 149개를 무력화 (치명, 광범위)

**규칙**: `instrument.css:15` → `instrument.scoped.css:19`

```css
.ins button { background: none; border: 0; padding: 0; cursor: pointer; }
```

**specificity**: `.ins button` = **(0,1,1)** · 페이지 로컬 `.dib-tab` = **(0,1,0)**
→ b가 같고 c가 1 > 0이므로 **리셋이 이긴다.** 문서 순서는 개입하지 못한다(specificity가 다름).

**결과**: `<button class="페이지로컬클래스">`로 만든 모든 커스텀 버튼이 새 셸 안에서
**배경·테두리·패딩을 잃는다.** Bootstrap `.btn`은 브릿지(`instrument.css:1070`)가
`.ins .btn` `(0,2,0) !important`로 다시 세워주지만, 페이지 로컬 버튼에는 그런 구제가 없다.

**정적으로 확인된 영향**: 페이지 규칙 **149개 / 20개 파일**. 대표 사례:

| 페이지 | 클래스 | 잃는 속성 |
|---|---|---|
| `DataInferenceBatchPage.razor:378` | `.dib-tab` (필터 탭 4개) | background, border, padding |
| `AiPromptPage.razor:144,196` | `.aip-summary-card`, `.aip-item` | background, border, padding |
| `BmesMakeModelGroupPage.razor:383` | `.pg-pill` (제품군 필터) | background, border, padding |
| `DataInferenceDbPage.razor` | `.didb-lang` `.didb-src-tab` `.didb-view-btn` 외 5종 | background, border, padding |
| `TranslatePage.razor` | `.img-remove` `.img-clear` | background, border, padding |
| `QrBakoDataPage.razor` | `.qr-btn` | background, border, padding, border-color |
| `BmesFCostPage.razor` | `.trend-expand-btn`, `.breakdown-expand-btn` | background, border, padding |

전체 목록은 §6 검증 로그의 `reset2.js` 출력 참조.

**부수 사례 — Bootstrap `.btn-close` 4개가 보이지 않게 된다.**
`.btn-close`는 × 아이콘을 `background: url(data:image/svg+xml,…)`로 그린다(`bootstrap.min.css`).
`.btn-close` `(0,1,0)` < `.ins button` `(0,1,1)` → `background: none`이 이겨 **아이콘이 사라진다**
(크기·클릭은 유지되므로 "안 보이는 닫기 버튼"이 된다).
`BmesReasonTablePage.razor:19`, `BmesRoutingTablePage.razor:19`,
`BmesSettingPage.razor:26`, `BmesTest3Page.razor:29`.

---

### S2 — `.ins table` 리셋이 페이지 테이블 39개의 `border-collapse`를 뒤집음 (치명, 광범위)

**규칙**: `instrument.css:16` → `instrument.scoped.css:20`

```css
.ins table { border-collapse: separate; border-spacing: 0; }
```

**specificity**: `.ins table` **(0,1,1)** > `.bmes-table` **(0,1,0)**.
Bootstrap Reboot의 `table { border-collapse: collapse }`는 `(0,0,1)`이라 애초에 상대가 안 된다.

**결과**: 페이지가 `border-collapse: collapse`를 선언한 테이블이 새 셸 안에서 **separate로 바뀐다.**
`collapse` 전제로 `td`/`th` 양쪽에 `border: 1px`를 준 테이블은 인접 셀 경계가 **2px로 겹쳐 보이고**,
`border-spacing: 0` 때문에 간격 없이 굵은 격자만 남는다.

**영향**: 페이지 테이블 클래스 **39개 / 15개 파일**. 이 앱의 핵심 화면 전부가 포함된다 —
`.fcost-table` `.fcost-weekly-table` `.fcost-trend-table` `.fcost-breakdown-table`
(`BmesFCostPage.razor:970,1019,1132,1181`), `.bmes-table`(Reason/Routing/Setting 3개 페이지),
`.dr-table`(`BmesDailyReportPage.razor:504`), `.cause-table` `.cause-monthly-table`,
`.mmg-table`, `.didb-*` 12종, `.didet-*`, `.dian-preview-table`, `.diask-action-table`,
`.dib-table`, `.test3-table`, `.test4-table`, `.fav-table`.

**구조적 아이러니**: 앱의 `<table>` 95개 중 Bootstrap `.table` 클래스를 단 **6개**만 쓴다.
브릿지 §24는 그 6개를 위해 20여 줄을 쓰고, 나머지 89개는 브릿지가 손대지 않는다 —
그런데 **리셋 한 줄이 그 89개를 전부 건드린다.** 브릿지는 앱이 거의 안 쓰는 것을 겨냥하고,
리셋은 앱이 실제로 쓰는 것을 조용히 바꾼다.

---

### S3 — 브릿지 `.row > *`가 Bootstrap 그리드 컬럼 73개를 무너뜨림 (높음)

**규칙**: `instrument.css:1094` → `instrument.scoped.css:1091`

```css
.ins .row > * {
  width: auto !important; max-width: none !important; flex-shrink: 1 !important;
  margin-top: 0 !important; padding-left: 0 !important; padding-right: 0 !important;
}
```

`.col-md-6 { width: 50% }`는 `(0,1,0)` non-important → `(0,2,0) !important`에 진다.
**모든 `col-*`가 콘텐츠 폭으로 붕괴한다.** `col-*` 사용 **73회 / 7개 파일**
(`col-md-6` 12, `col-lg-4` 7, `col-md-4` 6, `col-12` 6, `col-lg-3` 5, `col-auto` 5, …) —
`BmesSettingPage`, `BmesFCostPage`, `AdminAiUsagesPage`, `ReportPage`, `TranslatePage`,
`UsersPage`, `AdminDbQueryPage`.

**추가로 음수 마진이 남는다.** 브릿지는 `.row > *`의 padding만 0으로 만들고,
Bootstrap `.row` 자체의 `margin-left/right: calc(var(--bs-gutter-x) * -.5)`는 그대로 둔다
(`.ins .row`(§18, `instrument.css:969`)는 `display/align-items/gap`만 선언).
→ `.row g-3`은 좌우로 **8px씩 컨테이너 밖으로 삐져나온다.**

또 `.ins .row { align-items: center }`는 Bootstrap 기본 `stretch`를 바꾼다.
`align-items-end` 유틸은 `!important`라 살아남지만(`.align-items-end{…!important}`),
유틸이 없는 `class="row g-3"` 21곳은 세로 정렬이 바뀐다.

**근본 원인**: 디자인 시스템이 `.row`라는 **Bootstrap 소유 이름을 자기 flex 유틸로 재사용**했다.
`.ins .row`가 Bootstrap `.row`를 깨뜨렸고, 그 파손을 덮으려고 `.row > *` 브릿지가 생겼고,
그 브릿지가 다시 `col-*`를 깨뜨렸다. 파손의 연쇄다.

---

### S4 — `.ins a` 리셋으로 페이지 링크의 색·밑줄이 사라짐 (높음, 접근성)

`instrument.css:17` → `.ins a { color: inherit; text-decoration: none; }` = `(0,1,1)`.
`app.css:9`의 `a, .btn-link { color: #2563eb }`(`(0,0,1)`)도, 페이지 로컬 `.diask-ds-link`
같은 `(0,1,0)` 규칙도 전부 진다.

**결과**: 새 셸 안에서 **평범한 `<a href>`가 본문 텍스트와 시각적으로 구분되지 않는다.**
색으로도, 밑줄로도. WCAG 1.4.1(색에만 의존하지 말 것) 이전에 **어떤 어포던스도 남지 않는다.**

정적 확인 — 페이지 규칙 16개가 무력화:
`.dr-btn` `.dr-btn-ghost`(`BmesDailyReportPage`), `.diask-ds-link` `.diask-md`(`DataInferenceAskPage`),
`.didb-set-pt` `.didb-file-chip`(`DataInferenceDbPage`), `.didet-file-chip`, `.dima-md`,
`.pcd-btn`(`PcDownloadPage`), `.dib-badge-warn`.

`NewVerPage.razor:81`의 `<a href="/ui-redesign/index.html">`가 바로 이 케이스다 —
새 셸에서 열면 링크로 보이지 않는다.

---

### S5 — `.ins input/select/textarea { font: inherit }`가 입력 폰트 설정 38개를 무력화 (중간)

`instrument.css:14` → `.ins input, .ins select, .ins textarea { font: inherit; color: inherit }` = `(0,1,1)`.
`font` 축약형이라 **font-size / font-family / font-weight / line-height를 한꺼번에 상속으로 되돌린다.**

Bootstrap `.form-control`은 브릿지가 `!important`로 복구하지만(`instrument.css:1126`),
페이지 로컬 입력 클래스는 복구 대상이 아니다. 확인된 손실 **38개**:

* input 21개 — `.grp-name-input` `.mid-name-input`(`BmesMakeModelGroupPage`),
  `.bmes-search-input` `.bmes-input`(Reason/Routing/Setting), `.qr-input`, `.didb-pt-input`, …
* textarea 13개 — `.daily-textarea`(font-family 포함), `.di-text-area`, `.idt-textarea`,
  `.idt-auto-preview`(등폭 폰트 지정이 사라짐), `.diask-question`
* select 4개 — `.didb-ask-select`, `.idt-session-select`, `.qr-input`

특히 `.idt-auto-preview` / `.daily-textarea`는 **등폭 폰트로 정렬을 보여주는 프리뷰**다.
`font: inherit`이 걸리면 정렬이 무너진다.

---

### S6 — 다크 테마가 셸에만 존재한다 (높음, 기능 미완)

`data-theme="dark"`는 `.ins` 컨테이너에 붙고(`InstrumentLayout.razor:43`),
`.ins[data-theme="dark"]`가 토큰을 뒤집는다(`instrument.scoped.css`). 문제는 **페이지가 토큰을 안 쓴다**는 것.

* 페이지 `<style>` 40개 안에 **하드코딩 hex 색상 2,641개** (`#fff`, `#0f172a`, `#e2e8f0` …).
* 이들은 `data-theme`에 반응하지 않는다. 브릿지가 배경을 어둡게 만든 컨테이너 안에서
  페이지가 `color: #0f172a`(거의 검정)를 강제하면 **검정 위 검정**이 된다.
* `app.css`도 전부 라이트 팔레트 하드코딩(`!important` 117개)이며, 브릿지가 덮는 것은
  Bootstrap 계열 셀렉터뿐이다.

**추가로 테마가 저장되지 않는다.** `InstrumentLayout.razor:173`의 `private string _theme = "light"`는
서킷 메모리 필드다. `UiMode`는 localStorage에 미러링되지만(`Routes.razor:32`) **테마는 아니다.**
새로고침·재연결마다 라이트로 돌아간다. `ui-redesign/README.md`는
"the choice is remembered in localStorage"라고 적었지만, 그건 목업의 `instrument.js` 이야기이고
**앱은 `instrument.js`를 로드하지 않는다**(§5 참조).

---

### S7 — 페이지 `<style>` 11,075줄이 디자인 시스템보다 뒤에서, 전역으로 로드된다 (구조적)

40개 파일, 11,075줄, `!important` 107개. 전부 스코프가 없다.
두 방향 모두 위험하다.

**바깥 → 안**: 페이지 CSS는 `instrument.scoped.css`보다 **문서 순서상 뒤**다.
동일 specificity면 페이지가 이긴다. 즉 어떤 페이지가 `.tab`, `.view`, `.stat`, `.panel`처럼
셸 크롬이 쓰는 이름을 정의하면 **그 페이지를 여는 동안 셸 자체가 다시 칠해진다.**

**안 → 바깥**: 디자인 시스템 클래스 272개 중 상당수가 극도로 일반적인 이름이다 —
`.row .panel .view .stage .stack .live .stat .stats .chat .msg .tab .tabs .select .input
.num .code .small .muted .faint .empty .badge .btn .card .table .chip .field .deck .console
.spacer .n .b .t .ok`. 페이지 CSS는 서로 다른 클래스 1,693개를 쓴다.

**현재 실제 충돌 15건** (양쪽에 같은 이름이 정의됨):

| 클래스 | 정의한 페이지 수 | 비고 |
|---|---|---|
| `.active` | 13 | 디자인 시스템 쪽은 `.nav-tabs .nav-link.active`라 실질 충돌 없음 |
| `.n` | 7 | 디자인 시스템 쪽은 `.dt td.n` |
| `.btn` | 6 | 브릿지 `!important`가 페이지를 이김 → **페이지 커스터마이즈가 죽는다** |
| `.ok` | 5 | 디자인 시스템 쪽은 `.console .ok` |
| `.muted` `.empty` | 2 | 디자인 시스템 쪽이 bare `(0,1,0)` → `.ins .muted`가 이김 |
| `.card` `.card-header` `.card-body` | 3 | 브릿지 `!important`가 이김 → 페이지 카드 조정이 죽는다 |
| `.panel` `.select` `.num` `.btn-sm` `.form-control-sm` `.form-select-sm` | 각 1 | `.ins X`가 이김 |

지금은 15건이 대체로 "브릿지가 이긴다"로 귀결되지만, **다음에 추가되는 페이지가
일반적인 클래스 이름 하나만 써도 셸이 깨진다.** 방어 장치가 전혀 없다.

---

### S8 — `@keyframes` 이름이 전역으로 새어나간다 (낮음, 잠재)

`scope-css.js:113`은 `@keyframes` 블록을 통과시키고 **이름은 스코프하지 않는다**(불가능하다).
결과적으로 `pulse`, `blip`, `rise`(`instrument.scoped.css:348,957,1002`), 그리고
`modal-pop`(app.css)이 **한 문서 전역 네임스페이스**를 공유한다.

현재 Bootstrap이 정의하는 것은 `placeholder-glow` `placeholder-wave`
`progress-bar-stripes` `spinner-border` `spinner-grow` 뿐이라 **충돌 없음(확인함)**.
다만 페이지가 `@keyframes pulse`를 정의하면 나중 정의가 문서 전체를 덮는다.
접두사(`ins-pulse`)를 붙이면 영구히 닫힌다.

---

### S9 — 생성기 `tools/scope-css.js`의 취약점 4가지 (낮음~중간, 잠재)

생성기 자체는 **동작이 정확하다**(§6에서 재생성·바이트 비교로 확인). 다만 다음이 잠복해 있다.

**(a) `importantise`의 세미콜론 분할 (`scope-css.js:45-54`)**
선언 블록을 `';'`로 단순 분할한다. BRIDGE 영역에 `url("data:image/svg+xml;base64,…")`나
`content: ";"`가 들어오는 순간 선언이 쪼개져 **깨진 CSS가 생성된다.**
지금 브릿지 영역에는 세미콜론을 품은 값이 없어 우연히 안전하다.

**(b) 선언 블록 내부 주석이 `!important` 적용을 건너뛰게 만든다 (`scope-css.js:95-105`)**
주석을 만나면 `out += buf + comment`로 **버퍼를 즉시 방출**한다.
`}`에서만 `importantise(buf)`가 호출되므로, **주석 앞에 있던 선언들은 `!important` 없이 나간다.**
현재 브릿지 영역에서 블록 내부 주석은 `instrument.css:1095` 한 곳뿐이고 그 앞이 공백뿐이라
피해가 없다. 브릿지 블록 안에 주석을 한 줄 더 넣는 순간 그 위 선언들이 조용히 힘을 잃는다.

**(c) `*` 리셋이 `.ins` 자신에게는 적용되지 않는다**
`*, *::before, *::after` → `.ins *, .ins *::before, .ins *::after`.
`.ins` 컨테이너 자체는 `box-sizing: border-box`를 못 받는다. 현재 `.ins`에 padding/border가
없어 무해하지만, 규칙이 의도대로 완결되지 않았다.

**(d) 중복 셀렉터 `.ins, .ins` (`instrument.scoped.css:16`)**
`html, body { height: 100% }`에서 둘 다 `.ins`로 접히며 `.ins, .ins { height: 100%; }`가 나온다.
동작상 무해한 생성 흔적.

**참고 — 스코프 규칙 자체는 올바르다:**
`:root` → `.ins`, `[data-theme="dark"]` → `.ins[data-theme="dark"]`(플래그가 컨테이너에 있으므로 정확),
`@media` 내부 셀렉터는 스코프하고 `@keyframes` 스텝(`0%`, `from`)은 통과 —
`scope-css.js:107-127`의 스택 처리가 이 구분을 정확히 해낸다.

---

### S10 — 생성물이 빌드에 연결되어 있지 않다 (중간, 프로세스)

`instrument.scoped.css`는 **커밋된 생성 산출물**인데, 재생성을 강제하는 장치가 없다.
`.csproj`에 타깃 없음, npm 스크립트 없음, CI 훅 없음. 규약은 `README.md`와
`scope-css.js:22`의 주석("Re-run after every edit") 뿐이다.

* 현재 상태는 **동기화되어 있다**(§6에서 재생성 후 바이트 비교, 헤더 개행 제외 완전 일치).
* 하지만 `instrument.css`만 고치고 커밋하면 아무도 막지 못하며, 앱은 **낡은 스코프 파일**을 계속 서빙한다.
* 캐시 무효화도 수동이다 — `InstrumentLayout.razor:31`의 `?v=3`을 사람이 올려야 한다
  (`app.css?v=2`, `app.js?v=79`와 같은 방식). Blazor의 `MapStaticAssets` 지문(fingerprint)을
  안 쓰므로 **버전을 안 올리면 사용자는 낡은 CSS를 본다.**

부수: 생성기가 헤더를 LF로 쓰고 본문은 소스의 CRLF를 유지한다.
CRLF 체크아웃에서 재실행하면 헤더 4줄에 개행 차이가 나타난다(내용은 동일).

---

### S11 — 죽은 자산과 문서 드리프트 (낮음, 유지보수)

**클래스 사용률**: `instrument.css`가 정의하는 클래스 **272개 중 146개(54%)가
`Components/**/*.razor` 어디에도 등장하지 않는다.**
`pagehead* panel__* deck* stat__* hero__* dt--banded heat lvl* tip* legend* msg__* chat__log
composer* ctxrail ctxsec* srcrow* qrow thinking seg__b cbx switch__t notice* empty__t
grid2 grid3 leadrow split--wide stack-v tablewrap worksplit plate* …`
— 즉 §8~§19, §22의 대부분이 **앱에서는 아직 죽은 코드**다.
(이 수치는 토큰 일치 기준이며, 반대로 "등장한다"로 분류된 126개 중 일부는
`hero` `panel` `console` `select` 처럼 **페이지 로컬 CSS에 같은 이름이 있어서** 잡힌 것이다.
따라서 실사용률은 46%보다 더 낮다.)

**`§22 Login`(`.auth` `.plate*`, 43줄)은 도달 불가능하다.**
`LoginPage.razor:2`가 `@layout EmptyLayout`이고 자체 `<head>`에 `bootstrap` + `app.css`만 싣는다.
`instrument.scoped.css`는 로그인 화면에 아예 로드되지 않는다.

**`assets/instrument.js`(431줄)는 앱에서 전혀 로드되지 않는다.**
참조는 정적 목업 6개(`index/ng-rate/f-cost/ask-ai/admin-users/login.html`)뿐이다.
목업의 차트 런타임·정렬·툴팁·테마 저장 로직은 Blazor 셸에 존재하지 않는다.

**타깃이 0인 브릿지 셀렉터**: `.card-footer` `.btn-light` `.btn-lg` `.alert-primary`
`.modal-content` `.dropdown-menu` `.dropdown-item` `.accordion-item` `.accordion-button` `.toast`
— 마크업에 한 번도 없다. (`.blazor-error-boundary` `.validation-message` `.rpt-row`는
런타임/서비스 생성 마크업이라 0이어도 정상.)

**README 드리프트** — `wwwroot/ui-redesign/README.md`:

| README 서술 | 현재 사실 |
|---|---|
| "`instrument.scoped.css` — Used by `/new-ver` inside the app" | `InstrumentLayout`이 로드 → **모든 페이지**가 사용 |
| "Section 23 … restates the handful of properties Bootstrap still lands" | 브릿지 §23–24는 **190줄**, Bootstrap 전 영역(카드/폼/버튼/테이블/얼럿/배지/탭/드롭다운)을 `!important`로 재정의 |
| "the choice is remembered in localStorage" (테마) | 목업만 해당. 앱의 `_theme`은 **비영속** (S6) |

README는 스코프 파일 재생성 규약이 적힌 유일한 문서다. 그게 낡으면 §S10 리스크가 커진다.

---

### S12 — 잡다한 관찰 (정보)

* **FOUC**: `<link>`가 `<body>`에 있어(`InstrumentLayout.razor:31`) 셸 마크업 일부가
  스타일 적용 전에 파싱될 수 있다. 폰트도 같은 위치의 Google Fonts `<link>` 2개에 의존한다
  (`InstrumentLayout.razor:27-30`) — 외부 CDN 차단 환경에서는 `--sans`/`--mono` 폴백으로 떨어진다.
* **Syncfusion 테마 시트가 `.ins` 밖에 남는다.** Syncfusion 팝업/다이얼로그는 `body`로
  이동(portal)하므로 `.ins` 토큰을 못 받는다. 현재 `Components/**`에서 `<Sf*>` 사용은 0건이라
  실피해는 없지만, 도입 시 `.ins button` 리셋(S1)이 `.e-btn`(0,1,0)도 이긴다는 점에 유의.
* **정적 목업이 인증 없이 공개된다.** `Program.cs:164-166`은 `UseStaticFiles()`를
  `UseAuthentication()` 앞에 둔다 → `/ui-redesign/index.html`, `/ui-redesign/admin-users.html`
  등이 익명 접근 가능하다. 목업 데이터라 정보 노출 위험은 낮지만, 관리자 화면 구조가 노출된다.
* **`color-scheme`**: `:root { color-scheme: light }` → `.ins { color-scheme: light }`.
  루트가 아닌 요소에 걸리므로 문서 바깥(브라우저 UI, `body` 배경)에는 다크가 반영되지 않는다.
  `.ins`가 `100vh`를 채우므로 대개 안 보이지만, 오버스크롤 영역에서 라이트 배경이 드러난다.

---

## 3. 잘 되어 있는 것 (되돌리지 말 것)

* **스코프 방향이 옳다.** `.ins X` `(0,2,0)`가 Bootstrap `X` `(0,1,0)`를 이기고,
  `.ins` 밖에는 아무 영향이 없다 — 클래식 셸 회귀 위험 없음. 설계 의도대로 동작한다.
* **BRIDGE 영역 마킹**이 훌륭하다. `!important`를 손으로 반복하지 않고
  마커 한 쌍으로 영역을 선언하는 방식(`scope-css.js:42-43`)은 소스 가독성을 지켜준다.
  브릿지가 **한시적**이며 페이지 포팅이 끝나면 통째로 삭제된다는 의도도 주석에 명시돼 있다.
* **`@keyframes` / `@media` 처리**가 정확하다. 스텝 셀렉터를 스코프하지 않고
  조건부 at-rule 내부는 스코프한다 — 흔히 틀리는 지점을 제대로 짚었다.
* **`.dt` 스프레드시트 그리드**의 설계 근거(§14 주석: "스크린샷으로 잘려나가도 살아남게")가
  실제 사용 맥락(보고서 캡처)과 정확히 맞물린다.
* **생성 파일이 소스와 동기 상태다** — §6에서 실측 확인.

---

## 4. 권고 (영향/노력 순)

### R1. Element 리셋을 스코프 빌드에서 제외한다 — S1·S2·S4·S5를 한 번에 해결

리셋(`instrument.css:11-18`)은 **목업 페이지에는 필요하고, 앱에서는 유해하다.**
BRIDGE 마커와 대칭으로 `RESET-START` / `RESET-END` 마커를 추가하고,
`scope-css.js`가 스코프 출력에서 그 영역을 **버리도록** 한다. 필요한 최소만 남긴다:

```css
/* RESET-START — 목업 전용, 스코프 빌드에서 제외 */
button, input, select, textarea { font: inherit; color: inherit; }
button { background: none; border: 0; padding: 0; cursor: pointer; }
table { border-collapse: separate; border-spacing: 0; }
a { color: inherit; text-decoration: none; }
/* RESET-END */

/* 스코프 빌드에도 남는 것: 디자인 시스템 자신의 컴포넌트에만 */
.railbtn, .iconbtn, .tab__x, .seg__b, .chip__x, .qrow, .btn {
  background: none; border: 0; padding: 0; cursor: pointer;
}
.dt, .heat { border-collapse: separate; border-spacing: 0; }
```

`box-sizing: border-box`와 `:focus-visible`은 유지해도 안전하다(전자는 Bootstrap과 동일,
후자는 페이지에 경쟁 규칙이 없음).
**효과**: 페이지 규칙 242개가 되살아난다. 셸 크롬은 자기 클래스를 갖고 있으므로 손실 없음.

### R2. 디자인 시스템이 남의 이름을 쓰지 않게 한다 — S3·S7의 근본 해결

`.row` `.card` `.table` `.badge` `.btn` `.small` `.panel` `.select` `.input` `.num` `.n` `.b` `.t` `.ok`
는 Bootstrap 또는 페이지가 이미 쓰는 이름이다. 최소한 **`.row`부터** 네임스페이스로 옮긴다
(`.ins-row` 또는 `.hstack`). 그러면 브릿지의 `.row > *` 블록을 삭제할 수 있고,
`col-*` 73개와 21곳의 `class="row g-*"` 레이아웃이 원래대로 돌아온다.
이후 신규 클래스는 접두사 규칙(`ins-`)을 세운다.

### R3. 대안/보완 — Cascade layer로 소유권을 명시한다

`@layer bootstrap, app, instrument;`로 계층을 선언하고 페이지 `<style>`은 **레이어 밖**에 두면,
비-important 선언에서 **페이지가 specificity와 무관하게 항상 이긴다.**
S1·S2·S4·S5가 마커 없이 해결된다. 단 `!important`는 레이어 순서가 **역전**되므로
브릿지 `!important`와의 상호작용을 실측 후 도입할 것. R1보다 강력하지만 검증 부담이 크다.

### R4. 테마를 실제로 완성하거나, 범위를 명시한다

* 즉시: `_theme`을 `localStorage`에 미러링한다(`UiModeService`와 동일 패턴,
  `Routes.razor:32`가 이미 그 코드를 갖고 있다).
* 근본: 페이지 CSS의 하드코딩 hex 2,641개가 남는 한 다크는 완성될 수 없다.
  **페이지를 포팅할 때 hex → `var(--…)` 치환을 포팅 완료 조건으로 정의**하고,
  그 전까지는 `Toggle theme` 버튼에 "shell only" 같은 표시를 남기는 편이 정직하다.

### R5. 생성기 강화 (작고 확실한 것들)

* `importantise`를 문자 단위 파서로 바꾼다(괄호/따옴표 깊이 추적). — S9(a)
* 주석을 만나면 버퍼를 비우지 말고 **주석을 버퍼에 이어붙인다.** 그래야 `}`에서
  블록 전체가 `importantise`된다. — S9(b)
* `*` 리셋을 `.ins, .ins *, …`로 확장. — S9(c)
* 셀렉터 중복 제거(`.ins, .ins` → `.ins`). — S9(d)
* `@keyframes` 이름에 `ins-` 접두사. — S8

### R6. 빌드에 연결한다

`.csproj`에 `BeforeTargets="Build"`로 `node tools/scope-css.js` 실행 + 결과가 커밋본과
다르면 **빌드 실패**시키는 검사(또는 CI에서 `git diff --exit-code`)를 넣는다.
동시에 `?v=3` 수동 캐시버스팅을 파일 해시 기반으로 바꾼다.

### R7. 청소

* README를 현재 사실로 갱신(레이아웃 승격, 브릿지 규모, 테마 비영속). — S11
* 타깃 0인 브릿지 셀렉터 10종 삭제.
* `§22 Login`은 목업 전용임을 주석에 명시하거나, `LoginPage`를 실제로 포팅.
* `/ui-redesign/*`를 인증 뒤로 옮길지 결정(또는 배포 산출물에서 제외).

---

## 5. 우선순위 제안

| 순서 | 작업 | 해결되는 항목 | 위험 |
|---|---|---|---|
| 1 | R1 리셋 분리 | S1, S2, S4, S5 (페이지 규칙 242개) | 낮음 — 셸은 자기 클래스로 스타일링됨 |
| 2 | R2 `.row` 개명 + 브릿지 `.row > *` 삭제 | S3 (`col-*` 73개) | 낮음 — 디자인 시스템 내부 사용처만 수정 |
| 3 | R4 테마 영속화 | S6 절반 | 매우 낮음 |
| 4 | R5 생성기 강화 | S8, S9 | 매우 낮음 |
| 5 | R6 빌드 연결 | S10 | 낮음 |
| 6 | R7 문서/청소 | S11, S12 | 없음 |
| — | R3 cascade layer | R1·R2의 상위 대안 | 중간 — `!important` 역전 검증 필요 |

---

## 6. 검증 로그

모두 읽기 전용. 저장소 파일은 **하나도 수정하지 않았다**(이 문서 제외).
생성기 재실행은 `fs.writeFileSync`를 스크래치패드로 리디렉션해 원본을 보호했다.

| # | 수행 | 결과 |
|---|---|---|
| 1 | `node tools/scope-css.js` 재실행(출력 리디렉션) 후 커밋본과 diff | `scoped 439 blocks` · CR 정규화 후 **완전 일치** → 생성물 동기 확인 |
| 2 | `instrument.scoped.css` 스코프 정확성 육안 검증 | `:root`→`.ins`, `[data-theme=dark]`→`.ins[data-theme="dark"]`, `@media` 내부 스코프됨, `@keyframes` 스텝 미스코프 — 모두 정상 |
| 3 | 페이지 `<style>` 블록 정량화 (Python) | 40파일 / 11,075줄 / `!important` 107 / hex 2,641 |
| 4 | 디자인 시스템 ↔ 페이지 클래스명 교집합 (Node) | DS 272개 · 페이지 1,693개 · **충돌 15건** |
| 5 | element 리셋 vs 페이지 규칙 specificity 비교 (Node) | button **149** / table **39** / input 21 / a **16** / textarea 13 / select 4 = **242** |
| 6 | DS 클래스 사용 여부 토큰 스캔 (Node) | 272개 중 **146개 미등장** |
| 7 | `grep` — `col-*` 73회/7파일, `class="row` 21회, `.table` 6회, `<table` 95회, `btn-close` 4곳 | 본문 수치의 근거 |
| 8 | `bootstrap.min.css`에서 `.row` `.row>*` `.col-md-6` `.btn-close` `table` 원문 추출 | specificity 계산의 근거 |
| 9 | `@keyframes` 이름 충돌 확인 | Bootstrap 5종과 DS 3종 **충돌 없음** |
| 10 | `App.razor` / `Program.cs` / `Routes.razor` / `LoginPage.razor` 로드 순서·인증 순서 확인 | §1 표, S12 |

**수행하지 않은 것 (한계)**

* **브라우저 렌더 확인 없음.** 빌드·서버 기동을 하지 않았다(세션 제약). 위 결론은 전부
  specificity와 문서 순서에 근거한 **정적 추론**이며, 계산은 각 항목에 명시했다.
  S1~S5는 CSS 캐스케이드 규칙상 결정적이지만, **실제 화면 확인으로 한 번 더 검증할 것을 권한다** —
  특히 S3(그리드 붕괴)은 페이지마다 체감 정도가 다를 것이다.
* **동적 생성 마크업 미추적.** 서비스가 문자열로 만드는 리포트 HTML
  (`BmesReportHtmlExportService` 등)은 별도 문서로 내보내지므로 셸 CSS의 영향을 받지 않는다고
  보고 분석에서 제외했다.
* **Syncfusion 실사용 없음**으로 판단하고 깊이 보지 않았다(`<Sf*>` 마크업 0건).
* 클래스 사용률(§S11)은 **토큰 일치 휴리스틱**이다. "미사용 146개"는 신뢰할 수 있으나
  "사용 126개"는 페이지 로컬 동명 클래스 때문에 과대 계상돼 있다.
