# InferenceDataAIService session handoff — 2026-07-22

이 문서는 다음 Codex 세션이 2026-07-22 완료 상태에서 재실행 없이 이어가기 위한
최신 재개 지점이다. 989건 table-first 생성 과정과 code-only 감사의 상세 이력은
`SESSION_HANDOFF_2026-07-21.md`에 보존돼 있다. 충돌하면 이 문서의 현재 상태를
우선한다.

## 1. 최종 목표 완료 상태

단순 키워드 검색 결과를 나열하던 사용자 질의 경로를 질문 시점의 문맥 AI가
대상·조건·지표·비교·시간축을 해석하고 직접 관련 근거만 선택하는 경로로 교체했다.

- 최종 분석 batch: `outputs/table-first/full-989-v8-prompt-v4`
- 최종 history DB: `outputs/table-first-history/history.sqlite`
- DB 규모: workbook 989, Study 3,710, evidence 5,546, term 179
- 전체 batch: 989/989 성공, 실패 0
- 전체 corpus 정적 HTML은 사용자 지시에 따라 생성하지 않았다.
- 기존 `table-first-history-query`는 keyword candidate 검색 진단용으로만 남긴다.
- WPF 사용자 답변은 `table-first-contextual-query`를 호출한다.

사용자용 최종 경로는 다음 두 단계다.

1. 결정적 history 검색으로 최대 30개 Study 후보를 넓게 수집한다.
2. query-time AI가 질문과 직접 연결되는 Study/evidence/fact만 다시 판정한다.

후보에 질문 단어가 있다는 사실은 최종 근거가 아니다.

## 2. 문맥 AI 질의 구현

주요 파일:

- `inference_data_ai_contextual_query.py`
- `inference_data_ai_cli.py`
- `inference_data_ai_table_first_history.py`
- `tests/test_inference_data_ai_contextual_query.py`
- `InferenceDataAIService.Wpf/CanonicalEvidenceClient.cs`
- `InferenceDataAIService.Wpf/MainWindow.CanonicalEvidence.cs`
- `InferenceDataAIService.Wpf/EvidenceHtmlRenderer.cs`

CLI 예시:

```powershell
python inference_data_ai_cli.py table-first-contextual-query `
  --db outputs/table-first-history/history.sqlite `
  --question "VP CD 조립에 따른 Hearing 불량률 추이" `
  --candidate-limit 30 `
  --detail-candidate-limit 18 `
  --max-fact-count 216 `
  --reasoning-effort medium
```

AI 출력은 다음 일반 스키마를 사용한다.

- `answerMode`: `TREND`, `COMPARISON`, `CAUSE`, `SUMMARY`, `LOOKUP`, `OTHER`
- 질문 해석: subject, conditions, metrics, comparison, timeScope
- 근거 상태: `ANSWERED`, `PARTIAL`, `INSUFFICIENT`
- directAnswer, findings, trendRows, limitations, usedStudyIds

숫자 근거 계약:

- 요청 artifact의 정확한 원본 셀 숫자를 `TF-FCT-*` fact로 등록한다.
- 각 fact는 `TF-EVD-*`, Study, sheet/range, coordinate, displayValue,
  rowContext와 연결된다.
- AI는 등록된 `displayValue`를 그대로 사용해야 하며 fact/evidence ID를 함께
  반환해야 한다.
- 등록되지 않은 숫자, 변경한 숫자, 존재하지 않는 근거, fact와 evidence가 서로
  연결되지 않는 응답은 validator가 거절한다.
- `TREND + ANSWERED`에는 최소 2개의 비교 가능한 관측이 필요하다.
- 최종 근거는 최대 10개로 제한한다.
- 인과 효과와 승인된 effect는 생성하지 않는다.

AI 초안이 validator에 걸리면 같은 request로 최대 2회 자동 시도한다. 두 번 모두
실패하면 명령을 crash시키거나 잘못된 수치를 보여주지 않고
`CONTEXTUAL_AI_INSUFFICIENT_EVIDENCE` 안전 답변을 저장한다. 거절된 원문은
`*.ai-response.attempt-N.json`에 감사용으로 보존한다.

Codex 질의 프로세스는 임시 비-Git 디렉터리에서 `--ephemeral`,
`--skip-git-repo-check`, `--sandbox read-only`, JSON output schema로 실행한다.

## 3. 실제 989건 DB 문맥 질의 결과

대표 사용자 질문:

`VP CD 조립에 따른 Hearing 불량률 추이`

- status: `CONTEXTUAL_AI_PARTIAL`
- 후보 30건 중 직접 관련 Study 1건만 채택, 29건 제외
- citation 1개, 연결된 numeric fact 6개
- 조립 조건의 Hearing count는 확인했지만 날짜별 rate 관측이 없어 추이를
  계산하지 않았다.
- 출력:
  `outputs/table-first-contextual-answers/vp-cd-hearing.final-v1.answer.json`

서로 다른 유형의 일반화 실질검증:

### CAUSE / 교란변수

출력: `outputs/table-first-contextual-answers/generalization-gq06.answer.json`

- 후보 30 → 관련 1
- `CONTEXTUAL_AI_INSUFFICIENT_EVIDENCE`
- 금형 온도와 함께 가황제·금형 번호가 바뀌어 단일요인 효과 계산을 거절했다.

### COMPARISON / paired 전후값

출력: `outputs/table-first-contextual-answers/generalization-gq07.answer.json`

- 후보 30 → 관련 1
- `CONTEXTUAL_AI_PARTIAL`
- 같은 sample 열의 검사 전후 coil inner dimension 원값 8개를 연결했다.
- 변화량 자체가 fact로 등록되지 않았으므로 임의 계산하지 않고 전후 원값만
  제시했다.

### LOOKUP / 숫자표 없음

출력: `outputs/table-first-contextual-answers/generalization-gq10.answer.json`

- 후보 30 → 관련 2
- `CONTEXTUAL_AI_ANSWERED`
- `NO_TABULAR_EVIDENCE`를 명시하고 수치 효과 계산과 이미지 추출·분석을 하지
  않았다.

이 검증들은 특정 한글 질문 문자열에 대한 if/else 보정이 아니다. production
코드는 일반 intent schema와 evidence/fact validator만 사용한다.

## 4. WPF UI 전면 재설계

사용자 지시에 따라 기존 WPF를 완전히 새 WorkHub/Obsidian 계열 UI로 교체했다.
WorkHub 프로젝트의 실제 색상과 레이아웃 언어를 기준으로 했으며 WorkHub 자체에는
아직 코드를 이식하지 않았다.

핵심 파일:

- `InferenceDataAIService.Wpf/App.xaml`
- `InferenceDataAIService.Wpf/MainWindow.xaml`
- `InferenceDataAIService.Wpf/MainWindow.xaml.cs`
- `InferenceDataAIService.Wpf/EvidenceHtmlRenderer.cs`

적용 내용:

- 사용자 지정 title bar와 최소화/최대화/닫기
- Obsidian vault 방식 252px 좌측 navigator
- 중앙 workspace와 고정 우측 insight pane
- 하단 status bar와 접을 수 있는 developer console
- WorkHub 호환 팔레트
  - app `#202020`, rail `#242424`, surface `#252525`
  - hover `#303030`, border `#3A3A3A`
  - text `#D6D6D6`, muted `#8E8E8E`
  - accent `#8B5CF6`, soft accent `#3C2A58`
- 공통 Button, navigation, card, TextBox, ComboBox, DataGrid, ContextMenu,
  workspace TabControl 스타일을 `App.xaml` resource로 중앙화
- 기존 7개 기능을 새 navigator에 모두 연결
  - 문맥 AI 질문
  - 저장된 분석
  - Excel 자료
  - DB 적재
  - Excel ↔ DB
  - 사람 검토
  - 개념 사전
- 시작 페이지는 임의 저장 보고서가 아니라 문맥 AI 안내 화면이다.
- 내장 WebBrowser의 새 문맥 답변과 기존 저장 HTML에 WorkHub dark CSS를
  주입한다. 전체 corpus 정적 HTML 생성과는 별개다.

`WorkspaceNavigation_Click`, `WorkspaceTabs_SelectionChanged`, custom chrome
handler를 추가했다. XAML 초기화 도중 selection event가 먼저 발생해 생겼던
`WorkspaceStatusText` null 예외는 guard 후 생성자에서 명시적으로 다시 적용하는
방식으로 해결했다.

WorkHub로 실제 이식할 때는 `App.xaml`의 palette/resource key와
`MainWindow.xaml`의 rail/workspace/insight shell을 우선 옮기면 된다. 현재 내장
결과 뷰는 legacy WPF `WebBrowser`이므로 WorkHub 쪽 기술 선택에 따라 WebView2나
native view로 교체할 수 있다.

## 5. 검증 결과

최종 관련 Python 회귀:

```powershell
python -m unittest `
  tests.test_inference_data_ai_contextual_query `
  tests.test_inference_data_ai_table_first_history `
  tests.test_inference_data_ai_cli
```

- 34/34 PASS
- 문맥 runner 단독 테스트는 7/7 PASS
- 숫자 변조·미등록 fact·가짜 수치 거절
- 자동 재시도 성공과 반복 거절 후 안전 답변 검증

최종 WPF 빌드:

```powershell
dotnet build InferenceDataAIService.Wpf/InferenceDataAIService.Wpf.csproj `
  --no-restore
```

- 경고 0
- 오류 0

런타임 확인:

- 실행 파일:
  `InferenceDataAIService.Wpf/bin/Debug/net9.0-windows/InferenceDataAIService.Wpf.exe`
- 이 handoff 작성 시 PID: `25664`
- title: `Inference Data AI`
- `HasExited=False`, `Responding=True`
- 사용자에게 새 WPF를 보여주기 위해 현재 창을 열어둔 상태다.

다음 세션에서 PID가 없다고 해서 자동 재실행하지 않는다. 사용자가 화면을 다시
띄워 달라고 요청할 때만 실행한다.

## 6. 다음 세션의 재개 원칙

1. 989건 AI 분석이나 history index를 다시 만들지 않는다.
2. 현재 `outputs/table-first-history/history.sqlite`를 그대로 사용한다.
3. UI 수정 전 사용자가 현재 열린 화면에서 지적한 구체적인 가독성·폭·색상 문제를
   먼저 반영한다.
4. 질의 품질 문제는 keyword candidate 수가 아니라 AI가 선택한 Study/evidence/fact와
   directAnswer를 함께 확인한다.
5. 답변 숫자는 반드시 `TF-FCT-*`에서 `TF-EVD-*`와 원본 셀 좌표까지 추적한다.
6. NEEDS_REVIEW 관측값을 승인된 효과나 인과 결론으로 승격하지 않는다.
7. 사용자 지시 없이 전체 corpus HTML, 서버, 새 전체 batch를 실행하지 않는다.
8. 변경 시 가장 좁은 Python test 또는 WPF project build만 실행한다.

## 7. 작업 트리 주의

- `JinoSupporter` 저장소에는 이번 작업 전부터 사용자 소유의 대량 미커밋 변경과
  untracked 파일이 있다.
- 이번 작업에서도 commit, revert, reset, clean을 하지 않았다.
- `InferenceDataAIService.Wpf/CanonicalEvidenceClient.cs`,
  `EvidenceHtmlRenderer.cs`, `MainWindow.CanonicalEvidence.cs`,
  `inference_data_ai_contextual_query.py`,
  `inference_data_ai_table_first_history.py`,
  `TABLE_FIRST_SEMANTIC_PIPELINE.md` 등은 현재 Git에서 untracked로 보일 수 있으나
  최종 기능에 필요하므로 삭제하면 안 된다.
- `outputs/table-first-contextual-answers`의 request, raw AI response, answer JSON,
  Markdown은 실질검증 provenance이므로 유지한다.

