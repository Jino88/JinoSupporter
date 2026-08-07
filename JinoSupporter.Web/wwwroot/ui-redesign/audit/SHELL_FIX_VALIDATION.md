# SHELL FIX VALIDATION — R1(리셋 분리) · R2(.row 개명)

검증 대상 커밋: `e7e7177` (병합 `a577dcf`). 대조군은 `e7e7177^`의 `instrument.scoped.css`.

구현 세션은 자기가 만든 합성 마크업으로 자기 작업을 검증했다. 이 문서는 **실제 `.razor`
파일에서 뽑은 마크업과 실제 CSS 로드 순서**로 다시 확인한 결과다. 코드는 고치지 않았다.

## 판정

**통과.** 이 변경의 유일한 실패 모드(리셋 제거로 셸 크롬이 Bootstrap/app.css 스타일을
도로 뒤집어쓰는 것)는 발생하지 않았고, 의도한 페이지 스타일 복원은 전부 실측으로 확인됐다.

## 1. 렌더 대조 — 계산 스타일 전/후

`bootstrap.min.css` → `app.css`(head) → `instrument.scoped.css`(body) → 페이지 `<style>`
순서로 조립했다. 이 순서는 `App.razor`와 `InstrumentLayout.razor:31`의 실제 배치를 따른 것이다.
마크업은 `InstrumentLayout.razor`의 rail/topbar/tabs 구조와 아래 페이지들에서 그대로 가져왔다.
페이지 `<style>` 블록 4개(682줄)도 원문 그대로 넣어 실제 캐스케이드를 재현했다.

### 셸 크롬 — 6개 그룹 전부 변화 없음

| 대상 | 확인한 속성 | 결과 |
|---|---|---|
| `.railbtn` | background-color, border-top-width, padding-left, cursor, font-family, color | 전후 동일 |
| `.iconbtn` | 같음 | 전후 동일 |
| `.tab__x` | background-color, border-top-width, padding-left, cursor, font-family | 전후 동일 |
| `.btn--sm` | 같음 + color | 전후 동일 |
| `.navlink` | text-decoration-line, color, cursor | 전후 동일 |
| `.brand` | text-decoration-line, color | 전후 동일 |

### 페이지·프레임워크 스타일 — 7개 그룹 복원

| 대상 | 속성 | 변경 전 → 변경 후 |
|---|---|---|
| `.dib-tab` (`DataInferenceBatchPage`) | border-top-width | `0px` → `1px` |
| | padding-left | `0px` → `12px` |
| | background-color | `rgba(0,0,0,0)` → `rgb(248,250,252)` |
| | font-size | `13px` → `12px` |
| `.pg-pill` (`BmesMakeModelGroupPage`) | border-top-width | `0px` → `1px` |
| | padding-left | `0px` → `8px` |
| | background-color | `rgba(0,0,0,0)` → `rgb(255,255,255)` |
| `.bmes-table` (`BmesReasonTablePage`) | border-collapse | `separate` → `collapse` |
| `.btn-close` (Bootstrap) | background-image | `none` → `url("data:image/svg+xml,…")` |
| `.col-md-6` (Bootstrap) | width | `8.39px` → `675px` |
| | padding-left | `0px` → `7px` |
| `.row` (Bootstrap) | align-items | `center` → `normal` |
| 평범한 `<a>` | color | `rgb(15,19,25)` → `rgb(37,99,235)` |
| | text-decoration-line | `none` → `underline` |

구현 보고서의 수치와 방향이 모두 일치한다. `col-md-6` 폭의 절대값만 다른데(보고서 86.7→745,
여기 8.39→675) 하네스 컨테이너 폭이 달라서이고, 붕괴 정도는 오히려 보고서보다 심했다.

## 2. 정적 항목

| 항목 | 결과 | 근거 |
|---|---|---|
| 생성물이 소스와 일치 | 통과 | `node tools/scope-css.js` 재실행 후 `git diff` 내용 변화 없음. 출력: `scoped 438 blocks (6 dropped with the element reset)` |
| 스코프 출력에 element 리셋 | **0건** | `.ins button\|table\|a\|input\|select\|textarea` 형태 grep 0 |
| 스코프 출력에 bare `.row`/`col-*` | **0건** | 브리지 `.ins .row > *` 삭제 확인 |
| `.row-grp` / `.row-sum` 보존 | 통과 | 스코프 출력에 4건 그대로. 개명에 휩쓸리지 않음 |
| `.ins-row` 정상 스코프 | 통과 | `instrument.scoped.css:998-999` `.ins .ins-row` |
| 캐시버스터 | 통과 | `InstrumentLayout.razor:31` `?v=4` |
| 목업이 리셋 혜택 유지 | 통과 | 목업 6개가 모두 `assets/instrument.css`(원본)를 참조하고, 원본에는 `button`/`table`/`a` 리셋 3개가 그대로 있음 |
| §1b 컨트롤 커버리지 | 통과 | `.brand .btn .cbx .chip .composer__ta .dt .heat .iconbtn .input .navlink .qrow .railbtn .seg__b .select .switch .tab .tab__x .textarea` 재기술됨 |
| 브리지 `!important` 유지 | 통과 | 스코프 출력에 252개 |

## 3. 발견한 회귀나 누락

**없음.** 셸 크롬 6개 그룹이 전부 바이트 수준으로 동일하고, 복원된 7개 그룹은 모두 의도한
방향이다. 조작·과장으로 보이는 서술도 없었다.

## 4. 검증하지 못한 것

- **실제 실행 중인 앱 화면.** 서버 조작이 금지되어 있어(사용자의 F5 디버그 서버가 5050에서
  동작 중) 실제 브라우저 세션에서의 확인은 하지 않았다. F5 재빌드·재시작 후 사용자 확인이 필요하다.
- **모든 페이지.** 렌더 대조는 대표 4개 페이지의 `<style>` 블록과 대표 클래스에 한정했다.
  감사가 집계한 242개 규칙 전체를 하나씩 렌더로 확인하지는 않았다.
- **다크 테마.** `data-theme="light"` 상태만 측정했다. 감사 S6대로 다크는 아직 셸에만 있다.
