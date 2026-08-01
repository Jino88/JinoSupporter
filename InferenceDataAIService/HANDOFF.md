## 2026-07-22 07:04 - Existing handoff review
- Completed: Located the project root, confirmed that no standard `HANDOFF.md` existed, and reviewed the latest handoff `SESSION_HANDOFF_2026-07-22.md` in UTF-8.
- Decisions: Treat `SESSION_HANDOFF_2026-07-22.md` as the current project state when older handoffs conflict; do not rebuild the 989-workbook analysis or history index without a new user request.
- Files: Added `HANDOFF.md`; no source or existing handoff files were changed.
- Verification: Searched the project tree for `HANDOFF.md`, identified the Git/project boundaries, and successfully read the latest dated handoff; latest recorded checks are Python 34/34 passing and WPF build with 0 warnings/0 errors.
- Next: Await the user's next requested change and resume from `outputs/table-first-history/history.sqlite`; do not launch the WPF app unless explicitly requested.

## 2026-07-22 07:05 - WPF application launch
- Completed: Launched the existing Debug WPF executable at the user's explicit request.
- Decisions: Reused the previously built executable without rebuilding because no code changed and the requested action was only to open the application.
- Files: Updated `HANDOFF.md`; no source or build output files were changed.
- Verification: Process PID 22380 remained active with `HasExited=False`, `Responding=True`, and window title `Inference Data AI`.
- Next: Keep the application open for user review and await specific UI or behavior feedback.

## 2026-07-22 07:12 - Contextual evidence coverage implementation
- Completed: Diagnosed a real WPF query where 30 candidates contained roughly 10 plausible VP/CD-Hearing Studies in the first 12 but the final response exposed only 3; expanded WPF retrieval to 60 candidates/30 detailed candidates and separated all related Study evidence from the at-most-10 core numeric citations.
- Decisions: A Study remains relevant when assembly/material/process conditions and the requested metric are connected within the same Study even if causal sufficiency is weak; causal weakness lowers status/confidence instead of hiding the Study. `usedStudyIds` now represents the exhaustive directly related set (up to 40), while `citations` remains the core claim evidence and `relatedCitations` drives the WPF evidence inventory.
- Files: Changed `inference_data_ai_contextual_query.py`, `InferenceDataAIService.Wpf/CanonicalEvidenceClient.cs`, `InferenceDataAIService.Wpf/MainWindow.CanonicalEvidence.cs`, `InferenceDataAIService.Wpf/EvidenceHtmlRenderer.cs`, `InferenceDataAIService.Wpf/MainWindow.xaml`, and `tests/test_inference_data_ai_contextual_query.py`; updated `HANDOFF.md`.
- Verification: Implementation inspection completed; targeted Python tests, real-query validation, and the narrow WPF project build are pending.
- Next: Run `tests.test_inference_data_ai_contextual_query`, execute the reported question against the existing history DB, then rebuild/relaunch only the WPF project.

## 2026-07-22 07:19 - Contextual evidence coverage verification
- Completed: Verified the coverage fix against the exact user question `VP CD 조립 상태가 Hearing 불량에 끼치는 영향`, built the WPF project, and relaunched the updated application.
- Decisions: Display the exhaustive AI-related Study inventory separately from the core evidence used in numeric claims; the first-stage keyword pool remains a candidate set and does not itself count as relevant evidence.
- Files: Added validation artifacts `outputs/table-first-contextual-answers/vp-cd-hearing-coverage-fix.*`; retained the implementation and test changes from the preceding entry; updated `HANDOFF.md`.
- Verification: `python -m unittest tests.test_inference_data_ai_contextual_query` passed 8/8; real query changed the visible coverage from 3 Studies/3 evidence to 40 related Studies/147 related evidence while preserving 3 core numeric citations and 10 validated facts; `dotnet build InferenceDataAIService.Wpf/InferenceDataAIService.Wpf.csproj --no-restore` passed with 0 warnings/0 errors; relaunched PID 16708 with title `Inference Data AI`, `HasExited=False`, and `Responding=True`.
- Next: Have the user rerun the same question in the open WPF and review the expanded related-evidence grid; refine inclusion precision only if specific false-positive Studies are identified.

## 2026-07-22 07:32 - AI relevance-only query implementation
- Completed: Replaced the WPF question design with a dedicated relevance-only AI path: DB retrieves up to 200 Study candidates, AI selects every report needed for the question, and WPF displays selected report metadata, conditions, metrics, selection reason, and exact source ranges without generating a result answer.
- Decisions: AI may interpret the question and judge document relevance only. It must not interpret numeric results, determine increase/decrease or good/bad, infer effects/causes, rank outcomes, or answer the analytical question. The result schema intentionally has no `directAnswer`, `findings`, `trendRows`, or numeric facts.
- Files: Added `inference_data_ai_relevance_query.py` and `tests/test_inference_data_ai_relevance_query.py`; changed `inference_data_ai_cli.py`, `InferenceDataAIService.Wpf/CanonicalEvidenceClient.cs`, `InferenceDataAIService.Wpf/EvidenceHtmlRenderer.cs`, `InferenceDataAIService.Wpf/MainWindow.CanonicalEvidence.cs`, `InferenceDataAIService.Wpf/MainWindow.xaml`, `InferenceDataAIService.Wpf/MainWindow.xaml.cs`, and `HANDOFF.md`.
- Verification: Implementation and schema tests were added; targeted Python tests, CLI parser check, real relevance query, WPF build, and runtime relaunch are pending.
- Next: Run the relevance/contextual unit tests, fix any failures, execute the user's exact question through `table-first-relevance-query`, then build and relaunch the WPF project.

## 2026-07-22 07:34 - Relevance-only query verification and WPF relaunch
- Completed: Verified the dedicated relevance-only AI query with the user's exact question and relaunched the rebuilt WPF showing a human-scannable related-report table instead of an AI result answer.
- Decisions: Keep AI limited to question interpretation and report/Study relevance selection. The UI exposes report date, file, Study, test groups, captured metrics, relevance reason, and source evidence ranges; humans decide what the report results mean.
- Files: Added validation artifacts under `outputs/table-first-relevance-answers/vp-cd-hearing.*`; retained the relevance backend, CLI, tests, and WPF changes from the preceding entry; updated `HANDOFF.md`.
- Verification: `python -m unittest tests.test_inference_data_ai_relevance_query tests.test_inference_data_ai_contextual_query` passed 13/13; CLI help for `table-first-relevance-query` passed; the exact real query reviewed 200 candidates in 58 seconds and selected 30 related Studies/30 workbooks with 58 source ranges from 3,710 indexed Studies; automated phrase scan found 0 increase/decrease/improvement/degradation/effect/cause/superiority judgments in AI relevance reasons; WPF build passed with 0 warnings/0 errors; relaunched PID 9436 with `HasExited=False`, `Responding=True`, title `Inference Data AI`.
- Next: In the open WPF, run the same question and visually review the 30-report overview and 58 clickable source ranges; adjust table density or relevance precision only from concrete user feedback.

## 2026-07-22 07:43 - Live AI progress and raw source-value implementation
- Completed: Added live Developer console progress every five seconds, automatic console expansion, AI question interpretation and per-report selection-reason logs, and deterministic raw-value extraction for every AI-selected Study.
- Decisions: Preserve the relevance-only AI boundary. Python code, not AI, reads selected Study capture artifacts and displays original condition, metric, Excel-formatted value, cell coordinate, sheet, and range. Include measure columns plus quantity/result headers such as Input, OK, Q'Ty, Total NG, NG rate, Hearing, Noise, and Touch; cap only the per-Study display at 240 values and expose truncation.
- Files: Changed `inference_data_ai_relevance_query.py`, `tests/test_inference_data_ai_relevance_query.py`, `InferenceDataAIService.Wpf/EvidenceHtmlRenderer.cs`, `InferenceDataAIService.Wpf/MainWindow.CanonicalEvidence.cs`, `InferenceDataAIService.Wpf/MainWindow.xaml`, and `HANDOFF.md`; added validation artifacts `outputs/table-first-relevance-answers/vp-cd-hearing.raw-values.answer.*`.
- Verification: Targeted relevance/contextual tests passed 14/14. Rebuilding the previous real AI selection without another AI call attached raw data to all 30 selected Studies: 1,540 source values across 58 evidence ranges, including verified samples `Input 108`, `Hearing Noise 66` and `61.1%`, `Total NG 2`, and `NG Rate 1.7%` with exact cells. WPF build/runtime verification is pending.
- Next: Build only `InferenceDataAIService.Wpf`, relaunch it, then have the user run the same question and inspect live progress plus raw-value tables.

## 2026-07-22 07:44 - Live progress and raw-value WPF verification
- Completed: Built and relaunched the WPF with automatic Developer console expansion, five-second progress heartbeats, four explicit query stages, per-selected-report AI relevance details, and nested raw source-value tables.
- Decisions: During the long AI relevance wait, show elapsed time and keep the button text at `관련 보고서 찾는 중...`. After completion, log the interpreted document need, subjects, conditions, metrics, every selected file/Study/matched aspect/reason, and the full result JSON path. Raw tables remain code-extracted and contain no AI result judgment.
- Files: Added the missing `System.Text.Json` import in `InferenceDataAIService.Wpf/MainWindow.CanonicalEvidence.cs`; retained all changes and artifacts from the preceding entry; updated `HANDOFF.md`.
- Verification: Relevance/contextual tests passed 14/14; real saved AI selection produced raw values for 30/30 Studies and 1,540 total points. The first narrow WPF build caught only the missing JSON namespace; after correction the same project build passed with 0 warnings/0 errors. Relaunched PID 1776 with `HasExited=False`, `Responding=True`, title `Inference Data AI`.
- Next: User should run the question in PID 1776 and verify that the console advances every five seconds and that each report expands into condition/metric/original-value/cell rows; refine density or grouping from the resulting screenshot.

## 2026-07-22 09:51 - Existing relevance history viewer
- Completed: Added `기존 이력 보기` and automatic startup loading of the latest saved relevance result that contains raw source values, so users can inspect prior results without rerunning AI.
- Decisions: Search both `outputs/wpf-evidence` and `outputs/table-first-relevance-answers`, prefer artifacts with `rawDataPointCount > 0`, then choose the newest. Loading history restores the question, selected evidence grid, human-readable report/raw-value HTML, and AI relevance console details without starting Python/Codex.
- Files: Changed `InferenceDataAIService.Wpf/CanonicalEvidenceClient.cs`, `InferenceDataAIService.Wpf/MainWindow.xaml`, `InferenceDataAIService.Wpf/MainWindow.CanonicalEvidence.cs`, `InferenceDataAIService.Wpf/MainWindow.xaml.cs`, and `HANDOFF.md`.
- Verification: Narrow WPF build passed with 0 warnings/0 errors. Relaunched PID 12320 with `HasExited=False`, `Responding=True`, title `Inference Data AI`; the preferred existing artifact contains 30 Studies, 58 evidence ranges, and 1,540 raw values.
- Next: User can review the automatically loaded history or press `기존 이력 보기`; only rerun `관련 보고서 찾기` when a new question or refreshed relevance selection is needed.

## 2026-07-22 07:22 - WorkHub company-glossary integration review
- Completed: Reviewed the read-only `inference_data_ai_term_dictionary.py` adapter and the 199-row legacy dictionary contract for reuse by a new central WorkHub Company Glossary.
- Decisions: Keep the existing `MicroSpeaker_ProductTech_DB/db/term_dictionary.csv` as a fallback, but prefer WorkHub's adapter-compatible export once the glossary feature has created it; no corpus rebuild is required.
- Files: `HANDOFF.md`.
- Verification: Confirmed current adapter only consumes DEFINED aliases and IGNORE terms and that the existing source contains 179 DEFINED, 19 IGNORE, and 1 NEEDS_DEFINITION rows.
- Next: Implement the WorkHub exporter, update the adapter's default path selection, and run the focused term-dictionary/table-first tests.

## 2026-07-22 07:32 - WorkHub company-glossary consumer integration
- Completed: Updated default term-dictionary resolution to prefer `%LOCALAPPDATA%/WorkHub/CompanyGlossary/term_dictionary.csv` when present, while preserving the explicit environment override and legacy repository CSV fallback.
- Decisions: Priority is explicit `INFERENCE_DATA_AI_TERM_DICTIONARY`, then WorkHub central glossary export, then `MicroSpeaker_ProductTech_DB/db/term_dictionary.csv`; no history/corpus data needs migration.
- Files: `inference_data_ai_term_dictionary.py`, `tests/test_inference_data_ai_term_dictionary.py`, `HANDOFF.md`.
- Verification: `python -m unittest tests.test_inference_data_ai_term_dictionary tests.test_inference_data_ai_table_first` passed 35/35, covering override, WorkHub preference, legacy fallback, and existing table-first dictionary behavior.
- Next: Launch WorkHub only when explicitly requested, open Company Glossary once to create/import the central export, then future Inference Data AI runs will consume it automatically.

## 2026-07-22 10:02 - 원본 수치 가로 피벗 UI 구현
- Completed: AI 관련 보고서의 세로형 원본 수치 목록을 원본 Excel과 유사하게 범위별 표로 분리하고, 왼쪽 조건 2열과 오른쪽 지표 열로 펼치는 가로 피벗 렌더링으로 변경했습니다.
- Decisions: AI가 계산하거나 결과를 판정하지 않도록 원본 표시값을 그대로 사용하며, 각 값 아래에 원본 셀 주소를 유지합니다. 지표 열은 Excel 열 순서를 따르고 행은 원본 행 번호 순으로 표시합니다.
- Files: `InferenceDataAIService.Wpf/EvidenceHtmlRenderer.cs`, `HANDOFF.md`.
- Verification: 구현 완료, WPF 프로젝트 범위 빌드 및 기존 이력 화면 확인 예정입니다.
- Next: WPF 프로젝트를 빌드한 뒤 앱을 다시 열어 기존 저장 이력에서 가로 피벗 표 렌더링을 확인합니다.

## 2026-07-22 10:06 - 가로 피벗 UI 검증 및 WPF 재실행
- Completed: WPF 빌드와 기존 저장 이력 데이터 호환성을 검증하고 새 실행 파일로 WPF를 다시 열었습니다.
- Decisions: 새 AI 호출 없이 기존 `vp-cd-hearing.raw-values.answer.json`의 원본 수치를 즉시 가로 피벗으로 렌더링합니다.
- Files: `HANDOFF.md`.
- Verification: `dotnet build .\InferenceDataAIService.Wpf\InferenceDataAIService.Wpf.csproj --no-restore` 성공(경고 0, 오류 0). 저장 이력 30 Study의 원본값 1,540개가 45개 원본 표·196개 행으로 피벗 가능하고 빈 지표가 0개임을 확인했습니다. WPF PID 12944가 응답 중입니다.
- Next: 사용자가 열린 WPF의 기존 이력 화면에서 표 가독성을 확인하고, 필요하면 조건 열 폭이나 지표 표기만 미세 조정합니다.

## 2026-07-22 11:46 - 관련 보고서 화면 가독성 재설계
- Completed: 30개 보고서와 모든 원본표가 한꺼번에 길게 노출되던 화면을 접이식 보고서 카드로 바꾸고, 첫 보고서만 기본으로 펼치도록 재구성했습니다. 원본표는 기준·시험 조건 2열을 고정하고 지표·수치를 크게 표시하도록 정리했습니다.
- Decisions: 파일 설명·근거 ID·Excel 열문자·셀주소의 상시 노출을 제거해 핵심 조건과 수치에 집중합니다. 셀주소는 값에 마우스를 올릴 때만 표시하며 AI 계산이나 결과 판정은 추가하지 않습니다.
- Files: `InferenceDataAIService.Wpf/EvidenceHtmlRenderer.cs`, `HANDOFF.md`.
- Verification: 구현 및 diff 형식 검사 완료. WPF 프로젝트 범위 빌드와 기존 이력 재표시 확인 예정입니다.
- Next: WPF 프로젝트를 빌드하고 실행 중 앱을 교체한 뒤 기존 저장 이력으로 새 접이식 카드·고정열 화면을 확인합니다.

## 2026-07-22 11:47 - 가독성 재설계 빌드 및 실행 확인
- Completed: 가독성 재설계가 적용된 WPF를 빌드하고 다시 실행했습니다. 기존 이력은 새 AI 호출 없이 새 렌더러로 표시됩니다.
- Decisions: 첫 보고서만 기본 펼침, 나머지 보고서는 제목·날짜·원본값 수량만 보이는 접힌 상태를 유지합니다. 원본표는 최대 높이를 두고 표 내부 스크롤을 사용합니다.
- Files: `HANDOFF.md`.
- Verification: `dotnet build .\InferenceDataAIService.Wpf\InferenceDataAIService.Wpf.csproj --no-restore` 성공(경고 0, 오류 0). WPF PID 24088이 응답 중이며 기존 이력 자동 로드 상태입니다.
- Next: 사용자가 열린 화면에서 카드 접기/펼치기와 고정 조건열의 실제 가독성을 확인합니다.

## 2026-07-22 13:17 - 여러 검토 파일 통합 매트릭스 구현
- Completed: AI가 선택한 여러 검토 파일의 원본 행을 하나의 통합표로 취합했습니다. 왼쪽 2열은 검토 파일과 시험 조건, 오른쪽은 검사 수량·OK·전체 NG·Sigma NG·Hearing NG·Noise·Touch·VP/CD 분리 지표군입니다.
- Decisions: 같은 원본 행의 수량과 비율은 같은 지표 셀에 함께 표시하고 원본 헤더·셀주소는 마우스 오버 정보로 보존합니다. 이름 기반의 결정론적 분류만 사용하며 값 계산이나 결과 판단은 하지 않습니다. 파일별 전체 원본표는 아래 접힌 상세보기로 유지합니다.
- Files: `InferenceDataAIService.Wpf/EvidenceHtmlRenderer.cs`, `HANDOFF.md`.
- Verification: 구현 및 diff 형식 검사 완료. WPF 프로젝트 범위 빌드와 기존 이력 통합표 렌더링 확인 예정입니다.
- Next: WPF를 빌드·재실행하고 기존 30 Study 이력에서 통합표 행/열 생성과 앱 응답 상태를 확인합니다.

## 2026-07-22 13:18 - 여러 검토 파일 통합표 검증 및 실행
- Completed: 통합 검토표를 빌드하고 기존 VP+CD/Hearing 저장 이력으로 데이터 구성을 검증한 뒤 WPF를 다시 실행했습니다.
- Decisions: 통합표를 기본 화면으로 사용하고, 파일별 상세 카드는 모두 접힌 보조 화면으로 둡니다.
- Files: `HANDOFF.md`.
- Verification: `dotnet build .\InferenceDataAIService.Wpf\InferenceDataAIService.Wpf.csproj --no-restore` 성공(경고 0, 오류 0). 기존 30 Study 원본에서 통합 대상 823개 값을 179개 시험 조건 행과 8개 지표군으로 구성할 수 있음을 확인했습니다. WPF PID 25124가 응답 중입니다.
- Next: 사용자가 열린 WPF의 통합표에서 파일·조건 2열 고정과 지표 가독성을 확인하고, 필요 시 지표군 명칭/포함 범위를 조정합니다.

## 2026-07-22 13:22 - 결과 표 넓게 보기 전환 구현
- Completed: 좌우 분할 때문에 통합표가 가려지는 문제를 해결하기 위해 Insight Pane 헤더에 `넓게 보기`/`분할 보기` 전환을 추가했습니다. 관련 보고서 조회 또는 기존 이력 로드가 완료되면 왼쪽 질문·근거 영역을 자동으로 접어 결과 표가 전체 작업 폭을 사용합니다.
- Decisions: 조회 진행 중에는 Developer console을 사용할 수 있게 유지하고 결과 표시 완료 시 자동으로 접어 표의 세로 공간도 확보합니다. 다른 좌측 메뉴로 이동하면 분할 보기를 복원합니다.
- Files: `InferenceDataAIService.Wpf/MainWindow.xaml`, `InferenceDataAIService.Wpf/MainWindow.xaml.cs`, `InferenceDataAIService.Wpf/MainWindow.CanonicalEvidence.cs`, `HANDOFF.md`.
- Verification: 구현 및 diff 형식 검사 완료. WPF 프로젝트 범위 빌드와 자동 넓게 보기 실행 확인 예정입니다.
- Next: WPF를 빌드·재실행하여 기존 이력이 전체 폭으로 열리고 `분할 보기` 버튼으로 복귀되는지 확인합니다.

## 2026-07-22 13:24 - 결과 표 넓게 보기 검증 및 실행
- Completed: 실행 중이던 기존 WPF 프로세스를 정확한 실행 경로로 확인해 종료한 뒤 재빌드·재실행하고, 기존 이력의 자동 넓게 보기와 헤더 전환 버튼을 검증했습니다.
- Decisions: 앱은 현재 넓게 보기 상태로 유지해 사용자가 통합표를 즉시 확인할 수 있게 합니다.
- Files: `HANDOFF.md`.
- Verification: 최초 빌드는 PID 25124의 실행 파일 잠금으로 실패했으며 해당 프로세스만 종료 후 동일 명령이 경고 0·오류 0으로 성공했습니다. UI 자동화에서 기존 이력 본문 폭 1,427px와 `분할 보기` 버튼 노출을 확인했고 버튼을 왕복 전환한 뒤 넓게 보기로 복구했습니다. WPF PID 2176이 응답 중입니다.
- Next: 사용자가 전체 폭 통합표의 가독성을 확인하고 필요 시 개별 지표 열 폭만 미세 조정합니다.

## 2026-07-22 13:29 - 통합표 비율 제거 및 조건 열 분리
- Completed: 통합표에서 점유율·비율(%) 원본값을 제외하고 비율만 있던 중복 행도 생성하지 않도록 변경했습니다. 왼쪽 열을 `검토 파일`, `날짜`, `Type`, `Content`로 분리했습니다.
- Decisions: 날짜는 원본 행 조건의 날짜를 우선하고 없을 때 Study 날짜를 사용합니다. Type과 Content는 원본 조건 셀 순서를 그대로 분리하며 값 계산은 하지 않습니다. 네 개의 식별 열은 가로 스크롤 중 고정합니다.
- Files: `InferenceDataAIService.Wpf/EvidenceHtmlRenderer.cs`, `HANDOFF.md`.
- Verification: 구현 및 diff 형식 검사 완료. 기존 저장 이력의 비율 제거 후 행 수 재확인과 WPF 빌드·재실행 예정입니다.
- Next: 실행 중 WPF를 정확히 종료하고 빌드한 뒤 기존 이력에서 날짜·Type·Content 분리와 % 미표시를 확인합니다.

## 2026-07-22 13:32 - 통합표 비율 제거·열 분리 검증 및 실행
- Completed: 통합표의 % 미표시와 날짜·Type·Content 분리 버전을 빌드하고 WPF를 다시 실행했습니다.
- Decisions: 원본 헤더 단위가 `%`로 저장됐더라도 표시값 자체가 건수인 Noise/Touch 값(예: 66, 2)은 유지합니다. 헤더가 rate이거나 표시값에 `%`가 있는 행만 비율로 제외합니다.
- Files: `HANDOFF.md`.
- Verification: WPF 빌드 성공(경고 0, 오류 0). 기존 30 Study 이력에서 통합표 대상은 496개 건수값·118개 조건 행이며 `%` 표시값은 0개입니다. 샘플 2025-04-17/CD lot date/17/3 행에서 Input 108, OK 42, Sigma NG 0, Noise 66, Touch 0, Hearing NG 66이 유지됨을 확인했습니다. WPF PID 16564가 응답 중입니다.
- Next: 사용자가 열린 전체 폭 통합표에서 날짜·Type·Content 열과 건수값만 표시되는지 확인합니다.

## 2026-07-22 13:38 - 동일 Excel 내부 검토표 분리 표시
- Completed: 한 Excel 파일 안의 상·하단 독립 표가 같은 검토로 이어져 보이지 않도록 각 tableId/evidence 범위를 별도 검토 구간으로 분리하고, 구간 시작에 원본 표 제목·파일·시트 범위를 표시하도록 구현했습니다.
- Decisions: 원본 지표명이 `... / No`이면 앞부분을 검토 제목으로 사용하고, 그렇지 않으면 Study의 원본 titles를 표 순서에 맞춰 사용합니다. 표 제목을 계산하거나 해석하지 않고 원본 텍스트와 tableId 경계를 그대로 활용합니다.
- Files: `InferenceDataAIService.Wpf/EvidenceHtmlRenderer.cs`, `HANDOFF.md`.
- Verification: 샘플에서 `RESULT CHECKING FUNCTION`(Sheet1!C22:T28)과 `RESULT CHECK NG SEPARATE AFTER CHECK FUNCTION`(Sheet1!C30:K36)이 서로 다른 source key와 원본 제목을 가짐을 확인했습니다. WPF 빌드·재실행 예정입니다.
- Next: 실행 중 WPF를 종료 후 빌드하고 기존 이력에서 두 검토 구간 제목 행이 별도로 표시되는지 확인합니다.

## 2026-07-22 13:39 - 동일 Excel 내부 검토표 분리 검증 및 실행
- Completed: 원본 tableId별 검토 구간 제목 행이 적용된 WPF를 빌드하고 다시 실행했습니다.
- Decisions: 통합표는 여러 파일을 계속 한 화면에 표시하되, 같은 파일 안에서도 원본 표가 바뀌면 보라색 구분 제목 행을 삽입해 별개 검토임을 명확히 합니다.
- Files: `HANDOFF.md`.
- Verification: WPF 빌드 성공(경고 0, 오류 0). 샘플 파일의 네 원본 범위가 `Test 1 :`(B7:F11), `Test 2 :`(B13:F17), `RESULT CHECKING FUNCTION`(C22:T28), `RESULT CHECK NG SEPARATE AFTER CHECK FUNCTION`(C30:K36)으로 각각 매핑됨을 확인했습니다. WPF PID 6600이 응답 중입니다.
- Next: 사용자가 열린 통합표에서 상·하단 검토가 별도 제목 구간으로 보이는지 확인합니다.

## 2026-07-22 13:47 - 독립 검토표 개별 테이블화 및 NG Rate 복원
- Completed: 같은 파일의 독립 tableId를 한 통합 테이블의 구간이 아니라 완전히 분리된 개별 HTML 표로 렌더링하도록 변경했습니다. 각 표는 원본처럼 `No`, `날짜`, `Type`, `Content` 뒤에 해당 검토에 존재하는 지표만 표시합니다.
- Decisions: 하단 점유율 보조행은 계속 제외하되 원본 헤더가 명시적으로 `NG Rate`인 Sigma·Hearing·전체·VP/CD 분리 불량률은 의미 있는 결과 열로 보존합니다. VP/CD 분리 행과 같은 원본 행의 `NG Rate`는 `VP/CD NG Rate`로 분류합니다.
- Files: `InferenceDataAIService.Wpf/EvidenceHtmlRenderer.cs`, `HANDOFF.md`.
- Verification: 구현 및 diff 형식 검사 완료. 샘플 데이터 분류와 WPF 빌드·재실행 예정입니다.
- Next: 실행 중 WPF를 종료하고 빌드한 뒤 상단/하단 검토표가 서로 다른 표로 표시되고 하단 표에 90.0%, 40.0%, 50.0%, 15.0% VP/CD NG Rate가 복원되는지 확인합니다.

## 2026-07-22 13:48 - 독립 검토표·NG Rate 검증 및 실행
- Completed: 독립 tableId별 개별 표와 명시적 NG Rate 복원이 적용된 WPF를 빌드하고 다시 실행했습니다.
- Decisions: 동일 파일명은 표 제목의 출처 정보로만 표시하며 데이터 표 자체는 tableId마다 독립 생성합니다. 보조 점유율은 제거하고 `NG Rate` 헤더가 있는 값만 별도 Rate 열로 표시합니다.
- Files: `HANDOFF.md`.
- Verification: WPF 빌드 성공(경고 0, 오류 0). 샘플 파일에서 독립 표 4개를 확인했고 하단 `RESULT CHECK NG SEPARATE AFTER CHECK FUNCTION` 표의 No 6/7, 검사수량 30/20/2/20, VP/CD 분리 NG 27/8/1/3, VP/CD NG Rate 90.0%/40.0%/50.0%/15.0%가 모두 보존됨을 확인했습니다. WPF PID 18248이 응답 중입니다.
- Next: 사용자가 열린 화면에서 동일 파일의 독립 검토표가 별도 카드·표로 나뉘고 VP/CD NG Rate 열이 표시되는지 확인합니다.

## 2026-07-22 14:12 - 유사 검토표 날짜 기준 통합 구현
- Completed: 파일·tableId별로 나열하던 표를 검토 유형별 통합표로 변경했습니다. 원본 검토표는 VP/CD 분리, 기능·Hearing, 외관·Vision, Dyne·Tension, Air Leak, 시험 수량, 일반 수량·불량 유형으로 분류하고 동일 유형의 행을 한 표에 모읍니다.
- Decisions: 표 제목에서는 파일명을 제거하고 날짜를 행의 구분 기준으로 유지합니다. 추적성을 잃지 않도록 파일명·시트·범위·원본 행은 각 행의 마우스 도움말에 남기며, 각 원본 표 안의 No 전파는 원본 경계마다 초기화합니다.
- Files: `InferenceDataAIService.Wpf/EvidenceHtmlRenderer.cs`, `HANDOFF.md`.
- Verification: 구현 및 diff 육안 검토 완료. WPF 프로젝트 범위 빌드와 저장 이력 렌더링 검증 예정입니다.
- Next: 실행 중 WPF를 정확한 경로로 종료한 뒤 WPF 프로젝트를 빌드하고, 저장된 30 Study/1,540 원본값에서 파일 수보다 적은 유사 검토 통합표와 날짜별 행이 생성되는지 확인합니다.

## 2026-07-22 14:14 - 유사 검토표 통합 검증 및 WPF 실행
- Completed: 원본 파일 수준의 studyGroup이 같은 파일 안의 서로 다른 표를 오분류하지 않도록 검토 유형 판정은 각 source table의 제목·지표만 사용하게 보정했습니다. 빌드 후 새 WPF를 실행했습니다.
- Decisions: 같은 파일의 `RESULT CHECKING FUNCTION`과 `RESULT CHECK NG SEPARATE...`는 각각 기능·Hearing과 VP/CD 분리 유형으로 들어가며, 다른 파일의 같은 유형 표와는 날짜별 행으로 합쳐집니다. 파일 정보는 표 분리 기준이 아니라 행 도움말의 추적 정보로만 사용합니다.
- Files: `InferenceDataAIService.Wpf/EvidenceHtmlRenderer.cs`, `HANDOFF.md`.
- Verification: `dotnet build .\InferenceDataAIService.Wpf\InferenceDataAIService.Wpf.csproj --no-restore` 성공(경고 0, 오류 0). 저장된 30 Study/1,540 원본값의 source-table 구조 점검에서 41개 표시 대상 원본 표가 6개 주요 유사 검토 유형으로 축약 가능함을 확인했습니다. WPF PID 12216이 응답 중입니다.
- Next: 사용자가 열린 기존 이력 화면에서 파일별 카드 대신 검토 유형별 표가 보이고, 각 행의 날짜와 원본 도움말이 올바른지 확인합니다.

## 2026-07-22 14:23 - Normal 대조군 불량률 열 피벗 구현
- Completed: 원본 조건값이 정확히 `Normal`인 행을 통합표 본문에서 제거하고, 해당 원본 검토표의 명시적 NG Rate를 시험 행의 `Normal 불량률` 열로 옮기는 피벗을 구현했습니다.
- Decisions: 시험 설명 안의 `Normal + 0.05`, `Normal - 0.05`는 대조군으로 오인하지 않고 일반 시험 조건으로 유지합니다. 날짜가 같은 Normal 행이 있으면 그 값을 우선 연결하고, 날짜가 없으면 같은 원본 검토표의 Normal Rate를 사용합니다. 비율은 계산하지 않고 원본에 명시된 Rate만 표시합니다.
- Files: `InferenceDataAIService.Wpf/EvidenceHtmlRenderer.cs`, `HANDOFF.md`.
- Verification: 샘플 원본에서 `F39=Normal`, Input 2,289, Q'Ty NG 86, NG Rate 3.76%를 확인했고 새 로직상 Normal 행은 숨겨지고 3.76%가 같은 검토의 Normal 불량률 열로 이동합니다. WPF 프로젝트 빌드·재실행 예정입니다.
- Next: WPF 프로젝트를 빌드하고 실행 중 앱을 교체한 뒤 기존 이력 화면에서 Normal 행 제거와 3.76% 대조군 열 표시를 확인합니다.

## 2026-07-22 14:24 - Normal 대조군 열 빌드 및 실행 검증
- Completed: Normal 대조군 피벗 버전의 WPF를 빌드하고 다시 실행했습니다.
- Decisions: 통합표에서는 Normal 자체를 시험 행으로 세지 않으며, 원본 대조군 수량·NG 건수는 별도 상세표에 보존하고 질문 관련 통합표에는 원본이 제공한 불량률만 대조 열로 노출합니다.
- Files: `HANDOFF.md`.
- Verification: `dotnet build .\InferenceDataAIService.Wpf\InferenceDataAIService.Wpf.csproj --no-restore` 성공(경고 0, 오류 0). WPF PID 15504가 응답 중입니다.
- Next: 사용자가 기존 이력 화면에서 2,289/86/3.76% Normal 행이 사라지고 관련 시험 행 옆 Normal 불량률 열에 3.76%가 표시되는지 확인합니다.

## 2026-07-23 07:07 - 최근 작업 이력 점검
- Completed: 사용자의 요청에 따라 프로젝트 HANDOFF, 최근 Git 커밋, 현재 작업 트리 상태를 점검하고 2026-07-22의 최신 구현·검증 흐름을 식별했습니다.
- Decisions: 최근 작업 파악에는 2026-07-14 이후 갱신되지 않은 커밋 이력보다 단계별 `HANDOFF.md` 기록을 우선 사용합니다. 현재 프로젝트 변경은 커밋되지 않은 항목이 많으므로 보존합니다.
- Files: `HANDOFF.md`.
- Verification: Git 루트가 `JinoSupporter`임을 확인했고, `InferenceDataAIService` 범위에 추적 파일 변경 10개와 미추적 항목 132개가 있음을 확인했습니다. 최신 HANDOFF 단계는 Normal 대조군 불량률 피벗의 WPF 빌드 성공(경고 0, 오류 0)입니다.
- Next: 사용자가 원하면 최신 UI 작업부터 이어서 점검하거나, 날짜·기능별로 더 상세한 변경 목록을 정리합니다.

## 2026-07-23 07:08 - WPF 애플리케이션 실행
- Completed: 사용자의 명시적 요청에 따라 기존 Debug WPF 실행 파일을 열었습니다.
- Decisions: 소스 변경 없이 애플리케이션 실행만 요청되었으므로 2026-07-22 14:23에 빌드된 기존 실행 파일을 사용하고 재빌드하지 않았습니다.
- Files: `HANDOFF.md`.
- Verification: `InferenceDataAIService.Wpf.exe` PID 3316이 `HasExited=False`, `Responding=True`, 창 제목 `Inference Data AI` 상태임을 확인했습니다.
- Next: 열린 WPF에서 최신 통합표와 Normal 불량률 표시를 확인하고, 문제가 있으면 구체적인 화면 상태를 기준으로 수정합니다.

## 2026-07-23 07:14 - 통합표 밀도와 조건 열 정렬 개선 구현
- Completed: 검사 수량만 있는 원본 표와 행을 통합표에서 제외하고, 같은 검토 유형이라도 실제 지표 구성이 다른 원본 표는 별도 표로 분리했습니다. Excel 병합 셀 때문에 다음 행에서 빠진 날짜·Type은 앞 행 값을 이어받고 현재 조건을 Content에 배치하도록 수정했습니다.
- Decisions: 수량 외에 OK 또는 NG·Rate·Hearing·Noise·Touch 등의 결과 지표가 하나도 없는 표는 요약 가치가 낮아 숨깁니다. Content에는 날짜·Type 다음의 첫 원본 조건만 표시하고 뒤쪽 주석 셀은 열 밀림을 막기 위해 제외합니다.
- Files: `InferenceDataAIService.Wpf/EvidenceHtmlRenderer.cs`, `HANDOFF.md`.
- Verification: 사용자 스크린샷과 저장된 30 Study/1,540 원본값 JSON으로 수량 전용 표와 병합 셀 누락 원인을 재현했으며, 코드 변경을 완료했습니다. WPF 빌드와 저장 이력 HTML 검증은 아직 수행하지 않았습니다.
- Next: WPF 프로젝트를 범위 빌드하고 저장된 relevance 결과를 렌더링하여 수량 전용 표 제거, 지표 프로필별 분리, 날짜·Type 이어받기를 검증합니다.

## 2026-07-23 07:17 - 통합표 실제 이력 렌더링 점검과 제목 보강
- Completed: 저장된 30 Study/1,540 원본값을 수정된 렌더러로 직접 변환해 표 밀도를 점검하고, 분리된 각 표 제목에 실제 표시 지표 목록을 추가했습니다.
- Decisions: 원본 지표 프로필별로 21개 표가 생성되므로 반복되는 검토 유형 제목 옆에 `검사 수량 · 전체 NG · Noise · Touch` 같은 지표 구성을 표시해 표의 차이를 즉시 구분합니다.
- Files: `InferenceDataAIService.Wpf/EvidenceHtmlRenderer.cs`, `HANDOFF.md`; 임시 검증 도구 `.codex-verify/InferenceDataAIRenderCheck`는 검증 후 제거 예정입니다.
- Verification: WPF 범위 빌드가 경고 0·오류 0으로 성공했습니다. 실제 이력 HTML에서 수량 전용 제목은 0개, 통합표 21개, 지표 셀 412개 중 빈 셀 4개(1.0%), 병합 셀 날짜·Type 이어받기 표시는 확인되었습니다.
- Next: 조건 열 검증을 동일 HTML 행 내부 기준으로 강화하고 다시 빌드·렌더링한 뒤 임시 검증 도구를 제거합니다.

## 2026-07-23 07:19 - 통합표 개선 최종 검증
- Completed: 실제 저장 이력의 동일 HTML 행 안에서 병합 셀 날짜·Type·Content 정렬을 검증하고 임시 .NET 9 렌더링 도구를 제거했습니다. WPF 실행 파일도 최종 소스로 다시 빌드했습니다.
- Decisions: 빈 열 최소화를 위해 지표 프로필별 21개 표 구성을 유지하고 각 표에 지표 목록을 표시합니다. 현재 실행 중인 이전 WPF 프로세스는 없으며 새 실행 파일을 자동 실행하지 않습니다.
- Files: `InferenceDataAIService.Wpf/EvidenceHtmlRenderer.cs`, `HANDOFF.md`; `.codex-verify/InferenceDataAIRenderCheck`는 제거했습니다.
- Verification: 최종 WPF 범위 빌드 성공(경고 0, 오류 0). 실제 이력에서 수량 전용 표 0개, 지표 셀 412개 중 빈 셀 4개(1.0%), 원본 행 52가 날짜 `2025-06-09`·Type `S931`·Content `Bonding small inside 0.05`로 같은 행에 정렬됨, 뒤쪽 주석의 Content 유입 0건을 확인했습니다. 새 Debug 실행 파일 존재와 임시 검증 폴더 제거도 확인했습니다.
- Next: 사용자가 명시적으로 요청하면 새 WPF를 실행하고 화면에서 표 분리·열 정렬을 확인합니다.

## 2026-07-23 07:20 - 개선된 WPF 실행
- Completed: 사용자의 명시적 요청에 따라 통합표 밀도와 조건 열 정렬 수정이 포함된 최신 Debug WPF를 실행했습니다.
- Decisions: 2026-07-23 07:18에 최종 빌드된 실행 파일을 그대로 사용했습니다. 첫 실행 시도는 PowerShell 작업 폴더 인자 충돌로 시작되지 않아 경로 계산을 보정한 뒤 다시 실행했습니다.
- Files: `HANDOFF.md`.
- Verification: `InferenceDataAIService.Wpf.exe` PID 10748이 `HasExited=False`, `Responding=True`, 창 제목 `Inference Data AI` 상태임을 확인했습니다.
- Next: 열린 저장 이력 화면에서 수량 전용 표 제거, 지표 프로필별 표 분리, 날짜·Type·Content 정렬을 사용자가 확인합니다.

## 2026-07-23 07:26 - 단일 비교표 화면 재설계 구현
- Completed: 파일별 상세 카드, 파일명, AI 선정 이유, 시험 조건·수집 지표 목록, 상태 요약 등 반복 설명을 relevance 결과 화면에서 제거하고 모든 검토 행을 하나의 비교표로 통합했습니다.
- Decisions: 화면 제목은 `관련 시험 비교표` 하나만 사용하고 그 아래에는 질문에서 추출한 비교군만 간단히 표시합니다. 표는 `검토`, `날짜`, `비교군`, `원본 측정값` 4열이며 측정값은 빈 열 대신 실제 존재하는 지표·값만 한 셀에 태그처럼 표시합니다. 파일·시트·범위는 화면에 노출하지 않고 행 마우스 도움말에만 보존합니다.
- Files: `InferenceDataAIService.Wpf/EvidenceHtmlRenderer.cs`, `HANDOFF.md`.
- Verification: 구현 완료. WPF 범위 빌드와 실제 저장 이력 단일 표 렌더링 검증은 아직 수행하지 않았습니다.
- Next: WPF 프로젝트를 빌드하고 실제 30 Study 이력 HTML에서 제목 1개, 비교군 설명, 단일 표, 파일별 카드 제거를 검증합니다.

## 2026-07-23 07:26 - 단일 비교표 WPF 빌드 검증
- Completed: 실행 중이던 기존 WPF PID 10748을 종료하고 단일 비교표 버전의 WPF 프로젝트를 빌드했습니다.
- Decisions: 첫 빌드는 실행 파일 잠금 때문에 실패했으며 코드 오류는 없었습니다. 잠금 주체가 정확히 `InferenceDataAIService.Wpf` PID 10748임을 확인한 뒤 해당 프로세스만 종료했습니다.
- Files: `HANDOFF.md`.
- Verification: 재빌드 성공(경고 0, 오류 0). 현재 WPF는 종료 상태이며 새 실행 파일이 생성되었습니다.
- Next: 저장된 30 Study 이력을 새 렌더러로 직접 변환해 단일 제목·단일 표와 파일별 카드 제거를 검증합니다.

## 2026-07-23 07:28 - 단일 비교표 실제 이력 검증 완료
- Completed: 저장된 30 Study/1,540 원본값을 최종 렌더러로 변환해 제목·비교군 설명·표 구조와 파일별 카드 제거를 검증하고 임시 검증 도구를 제거했습니다.
- Decisions: 메인 화면에는 파일 단위 구획이나 반복 설명을 다시 노출하지 않습니다. 한 행의 원본 추적 정보는 마우스 도움말에만 유지하며 실제 측정값은 존재하는 지표만 태그형으로 압축 표시합니다.
- Files: `InferenceDataAIService.Wpf/EvidenceHtmlRenderer.cs`, `HANDOFF.md`; `.codex-verify/InferenceDataAISingleTableCheck`는 제거했습니다.
- Verification: 실제 HTML에 `관련 시험 비교표` 제목 1개, 비교군 설명 1개, 전체 table 1개, 비교 행 84개, 측정값 태그 427개가 생성됐습니다. 파일별 상세 카드 0개, 파일별 상세 제목 0개, 수량 전용 표 0개, 기존 상태 요약 0개를 확인했습니다. WPF 빌드는 경고 0·오류 0이며 실행 파일이 존재하고 앱은 현재 종료 상태입니다.
- Next: 사용자가 명시적으로 요청하면 최신 WPF를 실행해 단일 표의 실제 화면 밀도를 확인합니다.

## 2026-07-23 07:54 - 단일 비교표 WPF 실행
- Completed: 사용자의 명시적 요청에 따라 단일 비교표 재설계가 포함된 최신 Debug WPF를 실행했습니다.
- Decisions: 2026-07-23 07:26에 빌드된 검증 완료 실행 파일을 사용했습니다.
- Files: `HANDOFF.md`.
- Verification: `InferenceDataAIService.Wpf.exe` PID 1688이 `HasExited=False`, `Responding=True`, 창 제목 `Inference Data AI` 상태임을 확인했습니다.
- Next: 열린 저장 이력 화면에서 제목 1개·비교군 설명·단일 표의 실제 가독성을 확인합니다.

## 2026-07-23 07:57 - 비교표 Excel 직접 열기 구현
- Completed: 비교군 셀과 모든 원본 측정값 태그를 클릭 가능한 Excel 링크로 변경하고 WPF WebBrowser의 전용 링크 탐색을 가로채 원본 Excel 시트·범위를 여는 처리를 추가했습니다. 행 전체의 긴 원본 파일 툴팁은 제거했습니다.
- Decisions: 단일 값 태그는 정확한 원본 셀을 열고, 여러 값이 한 태그에 묶이면 해당 원본 표 범위를 엽니다. 비교군 셀은 원본 표 전체 범위를 엽니다. 화면에는 파일 경로를 표시하지 않으며 링크 내부의 추적 정보로만 사용합니다.
- Files: `InferenceDataAIService.Wpf/EvidenceHtmlRenderer.cs`, `InferenceDataAIService.Wpf/MainWindow.xaml`, `InferenceDataAIService.Wpf/MainWindow.CanonicalEvidence.cs`, `HANDOFF.md`.
- Verification: 저장 이력 JSON에 각 Study의 `sourcePath`와 원본값의 `sheet`·`range`·`coordinate`가 존재함을 확인하고 링크/탐색 구현을 완료했습니다. WPF 빌드와 링크 구조 검증은 아직 수행하지 않았습니다.
- Next: 실행 중 WPF PID 1688을 종료하고 프로젝트를 빌드한 뒤 실제 이력 HTML의 모든 비교군·측정값 링크와 사용자 지정 URI 파싱을 검증합니다.

## 2026-07-23 07:58 - 비교표 Excel 링크 검증 완료
- Completed: 링크 변수 이름 충돌 1건을 수정하고 WPF를 재빌드한 뒤 실제 저장 이력의 모든 비교군·측정값 링크와 WPF URI 파서를 검증했습니다. 임시 검증 도구는 제거했습니다.
- Decisions: 긴 행 툴팁은 완전히 제거하고 클릭 가능한 비교군·측정값에만 짧은 안내 툴팁을 유지합니다. 원본 경로가 실제 존재할 때만 현재 데이터의 링크가 정상 동작하는 것으로 검증합니다.
- Files: `InferenceDataAIService.Wpf/EvidenceHtmlRenderer.cs`, `InferenceDataAIService.Wpf/MainWindow.xaml`, `InferenceDataAIService.Wpf/MainWindow.CanonicalEvidence.cs`, `HANDOFF.md`; `.codex-verify/InferenceDataAIExcelLinkCheck`는 제거했습니다.
- Verification: WPF 빌드 성공(경고 0, 오류 0). 실제 84개 비교 행에 비교군 링크 84개와 측정값 링크 427개, 합계 511개 전용 Excel 링크가 생성됐습니다. 긴 행 툴팁은 0개이며 첫 링크는 존재하는 원본 파일, `Test 2` 시트, `B21:P27` 범위로 정상 해석됐습니다.
- Next: 최신 WPF를 실행해 사용자가 비교군 또는 측정값을 클릭하고 Excel이 해당 셀·표 범위로 이동하는지 확인합니다.

## 2026-07-23 07:59 - Excel 클릭 버전 WPF 실행
- Completed: 비교표 Excel 직접 열기 기능이 포함된 최신 WPF를 실행했습니다.
- Decisions: 링크 기능 빌드를 위해 종료했던 기존 WPF 대신 2026-07-23 07:57 빌드 실행 파일을 사용했습니다.
- Files: `HANDOFF.md`.
- Verification: `InferenceDataAIService.Wpf.exe` PID 11116이 `HasExited=False`, `Responding=True`, 창 제목 `Inference Data AI` 상태임을 확인했습니다.
- Next: 사용자가 비교군 텍스트나 측정값 태그를 클릭해 Excel이 지정된 시트·셀 또는 표 범위로 열리는지 확인합니다.

## 2026-07-23 11:11 - 원본 비교 세트별 대조군·비교군 구분 구현
- Completed: 평면 행 목록을 원본 Excel 표 단위의 비교 세트로 묶고 각 조건 행에 `대조군`·`비교군` 배지를 표시했습니다. 기존에 숨기던 Normal 행을 대조군으로 복원하고 Normal 값을 비교군 행에 반복 합성하던 로직을 제거했습니다.
- Decisions: 조건에 Normal·Control·Baseline·Reference·Standard·대조·기준이 명시된 행만 대조군으로 판정하며, 근거가 없는 행을 임의로 대조군이라 추측하지 않습니다. 각 세트 머리글에는 대조군·비교군 수를 함께 표시합니다.
- Files: `InferenceDataAIService.Wpf/EvidenceHtmlRenderer.cs`, `HANDOFF.md`.
- Verification: 구현 및 전용 스타일 적용을 완료했으며 WPF 빌드와 실제 저장 이력 렌더링 검증은 아직 수행하지 않았습니다.
- Next: 실행 중인 기존 WPF를 종료하고 WPF 프로젝트를 빌드한 뒤 실제 이력에서 세트 머리글, Normal 대조군 복원, 배지 수, Excel 링크 보존을 검증합니다.

## 2026-07-23 11:14 - 대조군·비교군 렌더링 검증 완료
- Completed: 변경된 WPF를 빌드하고 실제 `vp-cd-hearing` 저장 이력을 렌더링해 비교 세트 묶음과 대조군·비교군 분리를 검증했습니다. 검증용 임시 프로젝트는 결과 확인 후 제거했습니다.
- Decisions: Normal 조건은 실제 대조군 행으로만 표시하고 비교군 행에 별도 Normal 지표를 중복 삽입하지 않습니다. 원본 Excel 열기 링크는 세트·조건·측정값 모두 유지합니다.
- Files: `InferenceDataAIService.Wpf/EvidenceHtmlRenderer.cs`, `HANDOFF.md`; `.codex-verify/InferenceDataAICohortCheck`는 제거했습니다.
- Verification: WPF 빌드 성공(경고 0, 오류 0). 실제 이력에서 비교 세트 38개, 대조군 44행, 비교군 65행, Excel 링크 655개를 확인했습니다. `Normal line`과 Normal 조건 6행은 대조군, TEST 3 조건은 비교군으로 렌더링됐고 긴 행 툴팁 및 중복 Normal 합성 지표는 0개였습니다.
- Next: 최신 WPF를 실행해 실제 화면에서 세트 머리글과 대조군·비교군 배지의 가독성을 확인합니다.

## 2026-07-23 11:15 - 대조군·비교군 버전 WPF 실행
- Completed: 비교 세트별 대조군·비교군 구분이 포함된 최신 Debug WPF를 실행했습니다.
- Decisions: 사용자가 이어서 화면을 확인할 수 있도록 검증 완료 실행 파일을 사용했습니다.
- Files: `HANDOFF.md`.
- Verification: `InferenceDataAIService.Wpf.exe` PID 20464가 `HasExited=False`, `Responding=True`, 창 제목 `Inference Data AI` 상태임을 확인했습니다.
- Next: 열린 저장 이력에서 각 보라색 세트 머리글 아래의 녹색 대조군 배지와 보라색 비교군 배지, 조건별 원본 측정값을 확인합니다.

## 2026-07-23 11:18 - 비교 세트별 측정값 표 구현
- Completed: 하나의 긴 표와 보라색 측정값 태그 나열을 제거하고, 각 원본 비교 세트를 독립 카드·표로 렌더링하도록 변경했습니다. 각 표는 `구분·날짜·조건` 뒤에 해당 세트에 실제 존재하는 지표만 열로 만들고 값을 셀에 정렬합니다.
- Decisions: 세트에 없는 지표 열은 만들지 않아 불필요하게 빈 표가 넓어지지 않게 했습니다. 개별 값 셀과 조건은 기존처럼 원본 Excel로 연결하며, 측정값이 없는 교차 셀만 `—`로 표시합니다.
- Files: `InferenceDataAIService.Wpf/EvidenceHtmlRenderer.cs`, `HANDOFF.md`.
- Verification: 렌더러와 전용 표 스타일 구현을 완료했으며 빌드 및 실제 저장 이력 렌더링은 아직 검증하지 않았습니다.
- Next: 실행 중인 WPF를 종료하고 변경 프로젝트를 빌드한 뒤 세트별 표 38개, 동적 지표 열, 대조군·비교군 행 및 Excel 링크를 검증합니다.

## 2026-07-23 11:19 - 비교 세트별 측정값 표 검증 완료
- Completed: WPF를 빌드하고 실제 `vp-cd-hearing` 저장 이력으로 각 세트의 동적 측정값 표 구조를 검증했습니다. 검증용 임시 프로젝트는 제거했습니다.
- Decisions: 지표명이 열 머리글, 원본값이 표 셀이 되며 각 세트 내부에서만 필요한 지표 열을 표시하는 구조를 확정했습니다. 왼쪽의 구분·날짜·조건 열은 가로 스크롤 중에도 고정합니다.
- Files: `InferenceDataAIService.Wpf/EvidenceHtmlRenderer.cs`, `HANDOFF.md`; `.codex-verify/InferenceDataAIGroupTableCheck`는 제거했습니다.
- Verification: WPF 빌드 성공(경고 0, 오류 0). 실제 이력에서 그룹 카드·표 각각 38개, 대조군 44행, 비교군 65행, 동적 지표 머리글 189개, 지표 셀 516개, 값 링크 508개, 전체 Excel 링크 655개를 확인했습니다. 38개 표 모두 열×행 셀 수가 일치했고 기존 단일 요약표 및 측정값 태그는 0개였습니다.
- Next: 최신 WPF를 실행해 실제 화면에서 세트별 표의 열 정렬과 가로 스크롤 가독성을 확인합니다.

## 2026-07-23 11:20 - 비교 세트별 표 버전 WPF 실행
- Completed: 각 그룹의 측정값을 독립 표로 보여주는 최신 Debug WPF를 실행했습니다.
- Decisions: 검증 완료 실행 파일로 기존 화면을 교체했습니다.
- Files: `HANDOFF.md`.
- Verification: `InferenceDataAIService.Wpf.exe` PID 19336이 `HasExited=False`, `Responding=True`, 창 제목 `Inference Data AI` 상태임을 확인했습니다.
- Next: 열린 저장 이력에서 각 그룹의 지표 열·값 셀 정렬, 왼쪽 고정 열 및 가로 스크롤을 확인합니다.

## 2026-07-23 11:31 - 누락 NG Rate 원본 수식 복원 구현
- Completed: 캡처에서 Excel 수식 셀이 누락된 기능 검사 행에 `OK·전체 NG·전체 NG Rate`를 원본 구성항목으로 복원하는 렌더링 로직을 추가했습니다.
- Decisions: `Input`과 NG 구성항목(Air leak, SPL, THD, SPL+THD, SPL+THD+F0, Noise, Touch)이 숫자로 존재하고 원본 Total NG/Rate가 없을 때만 계산합니다. 원본 표에서 확인한 `Total NG=SUM(I:O)`, `OK=Input-Total NG`, `Total NG Rate=Total NG/Input` 수식을 그대로 사용하며 2자리 백분율로 표시합니다.
- Files: `InferenceDataAIService.Wpf/EvidenceHtmlRenderer.cs`, `HANDOFF.md`.
- Verification: 저장 이력의 해당 행에서 원본 구성값을 확인했고 값 전용 추출본 `C36:S47`에서 수식과 병합 범위 없음까지 확인했습니다. 보호 원본 COM 읽기는 호출 거부로 실패했으며 생성된 숨김 Excel PID 22824·13144만 종료하고 사용자 Excel PID 24260은 유지했습니다. 빌드와 렌더링 검증은 아직 수행하지 않았습니다.
- Next: 기존 WPF PID 19336을 종료하고 WPF 프로젝트를 빌드한 뒤 해당 4행의 전체 NG/Rate가 각각 원본 수식 결과와 일치하는지 검증합니다.

## 2026-07-23 11:33 - 누락 NG Rate 복원 검증 완료
- Completed: WPF를 빌드하고 실제 저장 이력에서 문제의 `RESULT CHECKING FUNCTION OF SPK` 그룹에 `OK·전체 NG·전체 NG Rate` 열이 생성되는지 검증했습니다. 임시 렌더링 검증 프로젝트와 빈 COM 추출 폴더는 제거했습니다.
- Decisions: 계산값은 원본 수식과 구성항목이 확인된 행에만 추가하며, 기존 원본 Total NG/Rate 값이 있으면 덮어쓰지 않습니다.
- Files: `InferenceDataAIService.Wpf/EvidenceHtmlRenderer.cs`, `HANDOFF.md`; `.codex-verify/InferenceDataAINgRateCheck`와 `.codex-verify/ng-rate-source-extract`는 제거했습니다.
- Verification: WPF 빌드 성공(경고 0, 오류 0). 네 행이 각각 `491/OK 482/NG 9/1.83%`, `535/526/9/1.68%`, `598/587/11/1.84%`, `1028/1009/19/1.85%`로 렌더링되어 원본 Excel 수식 결과와 일치했습니다.
- Next: 최신 WPF를 실행해 해당 그룹에서 새 열과 값의 실제 표시를 확인합니다.

## 2026-07-23 11:34 - NG Rate 복원 버전 WPF 실행
- Completed: 누락 수식 복원 기능이 포함된 최신 Debug WPF를 실행했습니다.
- Decisions: 원본 수식 대조와 렌더링 검증을 통과한 실행 파일을 사용했습니다.
- Files: `HANDOFF.md`.
- Verification: `InferenceDataAIService.Wpf.exe` PID 1784가 `HasExited=False`, `Responding=True`, 창 제목 `Inference Data AI` 상태임을 확인했습니다.
- Next: 열린 이력의 해당 기능:Hearing 그룹에서 `OK·전체 NG·전체 NG Rate` 열과 네 행의 백분율을 확인합니다.

## 2026-07-23 12:12 - AI 재학습 필요성 평가
- Completed: 현재 table-first 구조와 최근 오류 유형을 기준으로 모델 재학습 필요성을 평가했습니다.
- Decisions: 현재 병목은 AI 모델 성능보다 결정적 추출·수식 보존·조건 파싱·렌더링 연결이며, 자체 학습 모델이 아닌 프롬프트 기반 AI와 검증 코드 구조이므로 현 단계의 재학습 필요성은 낮습니다. 추출기나 prompt/schema가 바뀌면 영향 workbook 재처리·재인덱싱은 필요할 수 있으나 이는 모델 학습과 구분합니다.
- Files: `HANDOFF.md`.
- Verification: `TABLE_FIRST_SEMANTIC_PIPELINE.md`, `FINAL_GOAL_EXECUTION_PLAN.md`, `README.md`, `inference_data_ai_relevance_query.py`, `inference_data_ai_table_first.py`에서 AI는 표 의미·관련성만 판정하고 숫자·수식·검증은 코드가 담당하는 계약을 확인했습니다.
- Next: 대표 workbook 평가셋에서 표 묶기·대조군 역할·지표 매핑 오류율을 측정한 뒤 prompt 보정 또는 선택적 재분석 범위를 결정합니다.

## 2026-07-23 13:24 - CLI 전 과정 WPF 통합 수준 감사
- Completed: CLI 명령군과 현재 WPF 버튼·클라이언트 호출을 대조해 무CLI 전 과정 가능 여부를 감사했습니다.
- Decisions: 단건 canonical 적재, 근거 조회, 사람 승인, 개념 정규화, Excel drill-down은 WPF에서 가능하지만 완전한 폐쇄 루프는 아닙니다. WPF의 구형 `com-index+analysis_runner`, 신형 `ingest-workbook`, 질문용 `table-first history`가 병존하며 신규 적재 DB와 질문 DB가 자동 동기화되지 않습니다.
- Files: `HANDOFF.md`.
- Verification: `MainWindow.xaml`, `MainWindow.xaml.cs`, `MainWindow.CanonicalEvidence.cs`, `CanonicalEvidenceClient.cs`, `App.xaml.cs`, `inference_data_ai_cli.py`를 확인했습니다. 신규 적재는 `outputs/universal-grid/InputDataFinish.sqlite`를 갱신하지만 질문은 존재하는 `outputs/table-first-history/history.sqlite`를 우선 사용하며, WPF에서 `ingest-corpus`, `table-first-history-index`, contextual query, history validate/acceptance를 직접 실행하는 UI는 없습니다.
- Next: 완전한 WPF 폐쇄 루프를 만들려면 폴더/코퍼스 적재 UI, 적재 후 질문 인덱스 증분 갱신, 질문 모드 선택, 재시도·수식 파생·검증 옵션을 연결하고 구형 분석 경로를 정리합니다.

## 2026-07-23 13:25 - WPF 통합 기준 정정
- Completed: 사용자의 지적에 따라 WPF 전 과정 통합의 기준을 `CLI 제거`가 아니라 `WPF가 CLI를 백그라운드 호출하여 사용자가 터미널 없이 전 과정을 수행`하는 것으로 명확히 했습니다.
- Decisions: 기존 CLI와 Python 파이프라인을 재작성할 필요는 없습니다. 현재 누락은 기술적 처리 능력이 아니라 WPF 버튼·진행 상태·인자·후속 단계 연결이며, CLI 명령들을 순서대로 오케스트레이션하면 완전한 WPF 흐름을 만들 수 있습니다.
- Files: `HANDOFF.md`.
- Verification: 현재 WPF도 `RunPythonAsync`와 `CanonicalEvidenceClient.RunCliAsync`로 `com-index`, `ingest-workbook`, relevance query, review, concept 명령을 이미 백그라운드 호출하는 동일 패턴을 사용합니다.
- Next: WPF에 코퍼스 적재 → 검증 → history 갱신 → 질문 모드 실행을 하나의 작업 흐름으로 연결하고 재개·재시도·로그 UI를 추가합니다.

## 2026-07-23 13:27 - 신규 DRM Excel COM-first 입력 계약 확정
- Completed: 앞으로 추가되는 Excel은 해제본이 없으므로 모든 신규 적재가 DRM 원본의 Excel COM 읽기 전용 추출에서 시작해야 한다는 입력 계약을 반영했습니다.
- Decisions: 신규 흐름은 `DRM 원본 → COM 고정 좌표·병합·수식·표시값 추출 → canonical Capture/AI 분석 → 검증·DB import → History 갱신 → 질문`입니다. COM은 원본을 읽기 전용으로 열고 저장하지 않으며, 병합 covered cell을 제거하거나 좌표를 압축하지 않습니다. 기존 해제본 코퍼스는 현재 기준 자료로 유지합니다.
- Files: `HANDOFF.md`.
- Verification: `excel-com-extractor` 지침과 현재 `com-index`, `ingest-workbook`, 프로젝트 COM extractor를 대조했습니다. 현재 `com-index`는 COM JSON을 Universal DB에 적재하지만 canonical `ingest-workbook`은 DRM-free OpenXML 전용입니다. 기존 프로젝트 COM extractor는 수식·표시값·서식 보존이 불충분하고 `office_blocking_windows(close=True)`로 광범위한 `WM_CLOSE`를 보낼 수 있어 신규 DRM 운영 경로로 그대로 사용하면 안 됩니다.
- Next: 안전한 COM capture 계약을 구현해 canonical ingest가 COM JSON/DB revision에서 이어지게 하고, WPF 신규 파일·폴더 적재가 이 경로를 기본 호출하도록 연결합니다.

## 2026-07-23 13:34 - 안전한 Excel COM Capture v2 구현
- Completed: DRM·정책 보호 Excel을 전용 Excel 인스턴스에서 읽기 전용으로 열고 고정 UsedRange 좌표, 병합 anchor/covered 셀, 수식·캐시값·표시값·서식을 Capture v2 payload로 추출하는 프로젝트 내부 COM backend를 구현했습니다.
- Decisions: 원본 저장은 호출하지 않으며 생성한 workbook/Excel 인스턴스만 닫습니다. 인증창은 생성한 Excel PID 소유 창만 관찰하고, 닫기는 제목·클래스·버튼 문구가 모두 정확히 일치할 때만 해당 버튼에 한해 허용합니다. COM과 OpenXML capture contract를 DB가 provenance를 유지한 채 함께 수용하도록 일반화했습니다.
- Files: `inference_data_ai_com_capture.py`, `inference_data_ai_source_ingest.py`, `HANDOFF.md`.
- Verification: 두 Python 모듈의 `py_compile`과 좌표·행렬·COM contract smoke check가 통과했습니다. 실제 DRM workbook COM 실행은 전체 CLI 연결 뒤 표본 파일로 검증할 예정입니다.
- Next: `ingest-workbook`/`ingest-corpus`에 COM backend·인증 옵션·진행 이벤트를 연결하고 WPF가 기본적으로 이 경로를 호출하게 합니다.

## 2026-07-23 13:38 - COM-first 단건·폴더 CLI 연결
- Completed: `ingest-workbook`과 `ingest-corpus`에 `openxml|com` capture backend, 병합 covered-cell 정책, 숨김 sheet 포함, 안전한 인증창 관찰·정확 일치 닫기 옵션을 연결했습니다. 각 durable workflow 단계는 stderr의 `PROGRESS_JSON` 이벤트와 journal에 RUNNING/COMPLETED/FAILED 상태를 남깁니다.
- Decisions: 기존 해제본은 기본 `openxml` 동작과 기존 run ID를 유지하고, 신규 WPF 경로만 `com`을 명시합니다. COM corpus는 `.xlsx/.xlsm/.xlsb/.xls`를 발견하며 lock file `~$`는 제외합니다. 진행 출력 오류가 실제 적재를 중단시키지 않도록 telemetry callback은 fail-safe입니다.
- Files: `inference_data_ai_workflow.py`, `inference_data_ai_corpus_workflow.py`, `inference_data_ai_cli.py`, `HANDOFF.md`.
- Verification: 관련 4개 모듈 `py_compile`, 두 CLI help 계약, `source_ingest/corpus_workflow/cli/workflow` 단위 테스트 51개가 모두 통과했습니다.
- Next: 적재된 canonical DB를 최신 질문에 즉시 사용하는 질문 모드와 후속 준비 단계를 정리한 뒤, WPF에 단건·폴더 선택, 단계형 진행 표시, 재개·실패 재시도·로그를 연결합니다.

## 2026-07-23 13:46 - WPF DRM Excel 전체 처리 UI 구현
- Completed: WPF에 DRM Excel 단건/폴더 선택, COM 고정 추출부터 AI 분석·근거 검증·DB 반영·질문 준비까지 7단계를 대기/진행/완료/실패 색상으로 보여주는 단계형 화면을 구현했습니다. CLI 진행 이벤트를 실시간 수신하고 journal 기반 실패 재시도, 인증창 관찰 및 정확 일치 버튼 클릭 설정, 질문 화면 바로가기를 연결했습니다.
- Decisions: 신규 WPF 적재는 항상 `--capture-backend com --covered-cell-mode blank`를 사용하며 폴더는 Excel COM 안전성을 위해 workbook worker 1개로 직렬 처리합니다. 질문은 `최신 적재 DB`를 기본으로 사용하고 기존 table-first 전체 이력은 별도 선택 모드로 유지해 신규 자료가 오래된 history DB에 가려지지 않게 했습니다.
- Files: `InferenceDataAIService.Wpf/MainWindow.xaml`, `InferenceDataAIService.Wpf/MainWindow.CanonicalEvidence.cs`, `InferenceDataAIService.Wpf/CanonicalEvidenceClient.cs`, `HANDOFF.md`.
- Verification: WPF Debug 프로젝트 빌드가 경고 0, 오류 0으로 통과했습니다. 사용자 지침에 따라 앱 자체는 실행하지 않았습니다.
- Next: COM payload·DB import 회귀 테스트와 CLI 진행 출력 계약 테스트를 추가하고, 가능한 로컬 Excel 표본으로 read-only COM capture를 실행해 원본 불변과 병합 좌표 보존을 검증합니다.

## 2026-07-23 13:51 - COM-first 전체 처리 회귀 검증 완료
- Completed: COM capture contract DB import, 정확 일치 인증창 안전장치, 고정 좌표 helper, COM corpus 확장자 발견, CLI COM 옵션을 회귀 테스트로 고정했습니다. 폴더 진행 UI는 각 파일의 단계가 반복될 때 현재 파일 기준으로 초기화하고, 폴더 전체 완료 뒤에만 AI 문의 준비를 완료로 표시하도록 보정했습니다.
- Decisions: 인증창 관찰 정보는 최종 CLI JSON을 오염시키지 않도록 stderr 전용 event로 출력합니다. 실제 Excel COM 실행은 데스크톱 앱을 명시적으로 실행하지 말라는 프로젝트 지침에 따라 이번 검증에서는 수행하지 않았고, pure contract·DB·CLI·WPF 검증으로 마감했습니다.
- Files: `tests/test_inference_data_ai_com_capture.py`, `tests/test_inference_data_ai_corpus_workflow.py`, `tests/test_inference_data_ai_cli.py`, `inference_data_ai_com_capture.py`, `InferenceDataAIService.Wpf/MainWindow.CanonicalEvidence.cs`, `HANDOFF.md`.
- Verification: 관련 Python 테스트 55개가 모두 통과했고 WPF Debug 빌드는 경고 0, 오류 0입니다. 앱 및 실제 Excel COM 자동화는 실행하지 않았습니다.
- Next: 사용자가 실행을 요청하면 최신 WPF를 시작하고, 실제 DRM 표본 1개로 7단계 표시·인증창 정보·원본 SHA/mtime 불변·병합 좌표·신규 canonical 질문 결과를 화면에서 확인합니다.

## 2026-07-23 14:19 - 최신 WPF 실행
- Completed: 사용자의 요청으로 COM-first 7단계 전체 처리 UI가 포함된 최신 Debug WPF를 실행했습니다.
- Decisions: 직전 검증에서 통과한 Debug 산출물을 그대로 사용했습니다.
- Files: `HANDOFF.md`.
- Verification: `InferenceDataAIService.Wpf.exe` PID 25596이 `HasExited=False`, `Responding=True`, 창 제목 `Inference Data AI` 상태입니다.
- Next: WPF의 `DRM Excel 전체 처리` 탭에서 실제 DRM 원본을 선택해 7단계 흐름을 확인합니다.

## 2026-07-23 14:25 - 메인 작업 순서와 Excel 끌어놓기 구현
- Completed: 왼쪽 메뉴를 `1 DRM Excel 전체 처리 → 2 Excel↔DB 검수 → 3 사람 검토 승인 → 4 질문 관련 보고서` 순서로 재배치하고, 기존 `Excel 분석 목록`을 별도 ANALYSIS TOOLS로 구분했습니다. 목록 화면에는 단건/배치 사용 순서를 명시했으며 Excel 파일 또는 폴더를 DataGrid로 끌어놓아 재귀 검색·중복 제거 후 추가할 수 있게 했습니다.
- Decisions: 끌어놓기는 `.xlsx/.xlsm/.xlsb/.xls`만 수용하고 `~$` Excel 소유 파일과 접근 불가 하위 폴더는 건너뜁니다. 이 목록은 탐색·단건/배치 분석 도구이고, 신규 DRM 원본의 COM→AI→DB→질문 전체 흐름은 메인 1단계를 사용하도록 안내합니다.
- Files: `InferenceDataAIService.Wpf/MainWindow.xaml`, `InferenceDataAIService.Wpf/MainWindow.xaml.cs`, `HANDOFF.md`.
- Verification: WPF Debug 빌드가 경고 0, 오류 0으로 통과해 XAML drag/drop 이벤트 연결과 C# handler 컴파일을 확인했습니다.
- Next: 최신 WPF를 재실행하고 Excel 파일 또는 폴더를 분석 목록 중앙에 놓았을 때 보라색 drop overlay와 추가 건수 상태가 표시되는지 확인합니다.

## 2026-07-23 14:25 - 끌어놓기 버전 WPF 재실행
- Completed: 작업 순서 재배치와 Excel 파일·폴더 drag/drop이 포함된 최신 Debug WPF를 실행했습니다.
- Decisions: 빌드 통과 산출물을 바로 사용했습니다.
- Files: `HANDOFF.md`.
- Verification: `InferenceDataAIService.Wpf.exe` PID 10592가 `HasExited=False`, `Responding=True`, 창 제목 `Inference Data AI` 상태입니다.
- Next: 사용자가 화면에서 Excel 파일 또는 폴더를 `Excel 분석 목록` 중앙으로 끌어놓아 실제 입력 동작을 확인합니다.

## 2026-07-23 14:30 - COM Capture styleId 충돌 수정
- Completed: 실제 DRM 전체 처리에서 모든 셀이 `styleId=0`으로 기록되어 서로 다른 서식 payload가 충돌하던 원인을 수정했습니다. 정확히 같은 서식 JSON에는 같은 결정적 해시 ID를, 다른 서식에는 다른 ID를 부여하며 빈 서식만 0을 사용합니다.
- Decisions: 기존 결함 Capture와 실패 journal을 재사용하지 않도록 COM Capture 계약과 extractor 버전을 v2.1로 올렸습니다. 재시도 시 새 revision으로 다시 COM 추출합니다.
- Files: `inference_data_ai_com_capture.py`, `inference_data_ai_source_ingest.py`, `tests/test_inference_data_ai_com_capture.py`, `HANDOFF.md`.
- Verification: 대상 모듈 `py_compile`과 관련 Python 단위 테스트 56개가 모두 통과했습니다.
- Next: WPF 오류 표시 보정까지 빌드한 뒤 최신 앱을 재실행하고 동일 DRM 파일로 전체 처리를 다시 실행해 구조 패킷 생성 이후 단계 진입을 확인합니다.

## 2026-07-23 14:30 - WPF 초기 표시와 실패 메시지 정리
- Completed: 앱 시작 직후 이전 질문 이력을 자동 로드하며 WebBrowser가 흰 화면으로 남던 동작을 제거하고, Window Loaded 이후 어두운 홈 화면을 표시하도록 바꿨습니다. 결과 문서 로딩 중에는 어두운 placeholder를 보이며 두 입력 ComboBox의 닫힌 상태 글자를 명시적인 밝은 배경/어두운 전경으로 표시합니다. 처리 실패 시 전체 Traceback 대신 마지막 원인 한 줄만 팝업·단계 카드에 표시하고, 이미 실패한 3단계가 있으면 다음 4단계를 중복 실패 처리하지 않습니다.
- Decisions: 전체 Traceback은 Developer console에만 유지하고 단계 상세는 한 줄 180자로 제한해 UI 높이가 비정상적으로 커지지 않게 했습니다. 저장 이력은 사용자가 질문 관련 보고서를 명시적으로 열 때만 표시합니다.
- Files: `InferenceDataAIService.Wpf/MainWindow.xaml`, `InferenceDataAIService.Wpf/MainWindow.xaml.cs`, `InferenceDataAIService.Wpf/MainWindow.CanonicalEvidence.cs`, `HANDOFF.md`.
- Verification: WPF Debug 프로젝트 빌드가 경고 0, 오류 0으로 통과했습니다.
- Next: 최신 WPF를 실행해 초기 홈 화면·입력 ComboBox와 동일 DRM 파일 재시도 동작을 확인합니다.

## 2026-07-23 14:30 - styleId 수정 버전 WPF 실행
- Completed: COM Capture v2.1 styleId 수정과 WPF 표시 개선이 포함된 최신 Debug 앱을 실행했습니다.
- Decisions: 기존 결함 v2.0 Capture는 재사용하지 않으므로 같은 DRM 파일을 다시 실행하면 COM 추출부터 새 revision으로 진행합니다.
- Files: `HANDOFF.md`.
- Verification: `InferenceDataAIService.Wpf.exe` PID 13628이 `HasExited=False`, `Responding=True`, 창 제목 `Inference Data AI` 상태입니다.
- Next: UI에서 동일 파일을 선택하고 `전체 실행` 또는 `실패 단계 재시도`를 눌러 3단계 구조 패킷 생성이 완료되고 이후 단계로 넘어가는지 확인합니다.

## 2026-07-23 17:11 - DRM 화면의 질문 안내 패널 원인 확인
- Completed: DRM 전체 처리 중 오른쪽에 `질문 관련 보고서` 안내가 표시되는 경로를 조사했습니다.
- Decisions: 오른쪽 `Insight Pane`은 모든 작업 탭이 공유하며, 앱 시작 시 `ShowWorkspaceHome()`이 질문 안내 HTML을 넣습니다. 작업 탭 전환은 왼쪽 `WorkspaceTabs`만 바꾸고 오른쪽 내용을 숨기거나 초기화하지 않아 DRM 처리 화면에도 질문 안내가 잔류합니다. 이는 DRM 단계의 의도된 정보가 아니라 화면 상태 분리가 누락된 UI 문제입니다.
- Files: `HANDOFF.md`.
- Verification: `MainWindow.xaml.cs`의 `ShowWorkspaceHome`, `WorkspaceNavigation_Click`, `ApplyWorkspaceSelection`과 `MainWindow.xaml`의 공용 2열 `WorkspaceResultGrid` 구조를 대조했습니다.
- Next: 수정 요청 시 DRM 처리 중에는 오른쪽 패널을 접거나 단계별 진행/결과 전용으로 전환하고, 7단계 완료 또는 `질문 관련 보고서` 탭에서만 질문 안내·결과를 표시합니다.

## 2026-07-23 17:15 - 4단계 완료 후 진행 중 재표시 원인 확인
- Completed: 단일 파일 처리에서 `4/7 AI 시험·비교군 분석 · 완료` 직후 같은 단계가 다시 `처리 중`으로 표시되는 진행 이벤트 순서를 조사했습니다.
- Decisions: 백엔드는 `LOCATOR`와 `DRAFT`를 별도 순차 단계로 실행하지만 WPF가 둘 다 UI 4단계로 매핑합니다. 따라서 LOCATOR 완료 이벤트가 4단계를 완료로 만든 뒤 DRAFT 시작 이벤트가 같은 카드를 다시 진행 중으로 바꿉니다. 파일 중복 처리는 아니며, UI 집계 상태가 내부 하위 단계를 표현하지 못하는 문제입니다.
- Files: `HANDOFF.md`.
- Verification: 최신 `outputs/incremental-ingest/ingest-run_065791ddac1ab5bf7bf6c100/journal.json`에서 전체 상태 `RUNNING`, 현재 단계 `DRAFT`, `LOCATOR=COMPLETED`, `DRAFT=RUNNING`을 확인했습니다.
- Next: 수정 요청 시 4단계를 `4-1 관련 구간 탐색`과 `4-2 Study/비교군 구성`으로 분리하거나, LOCATOR 완료 시 UI 4단계를 완료 처리하지 않고 DRAFT 완료 후에만 최종 완료로 표시합니다.

## 2026-07-23 17:19 - AI 초안 검증 실패 재시도와 처리 UI 보정
- Completed: 단일 파일 및 폴더의 `실패 단계 재시도`가 CLI의 rejected-draft/source-selection 보정 옵션을 실제로 전달하도록 연결했습니다. LOCATOR 완료 뒤 DRAFT 시작 시 4단계가 완료에서 진행 중으로 되돌아가지 않고 `4-1/2 관련 구간 탐색`과 `4-2/2 Study·시험군·비교군 구성`을 연속 진행으로 표시합니다. DRM 전체 처리 탭에서는 공용 질문 Insight Pane을 숨겨 처리 단계가 전체 너비를 사용합니다. 콘텐츠 커버리지 실패 팝업은 누락 근거와 재시도 안내를 한국어로 표시합니다.
- Decisions: 최초 AI 초안은 기존 정확 예산 요청을 유지하며, 검증 실패 후 사용자가 재시도를 누른 경우에만 현재 rejected 초안을 원본 셀 근거로 1회 보정합니다. 누락 locator source도 같은 명시적 재시도에서만 승격합니다.
- Files: `InferenceDataAIService.Wpf/CanonicalEvidenceClient.cs`, `InferenceDataAIService.Wpf/MainWindow.CanonicalEvidence.cs`, `InferenceDataAIService.Wpf/MainWindow.xaml.cs`, `HANDOFF.md`.
- Verification: WPF Debug 프로젝트 빌드가 경고 0, 오류 0으로 통과했습니다.
- Next: 최신 WPF를 실행해 동일 파일을 다시 선택한 뒤 `실패 단계 재시도`를 눌러 기존 COM·패킷·locator 산출물을 재사용하고 rejected Study 초안 보정부터 완료되는지 확인합니다.

## 2026-07-23 17:19 - 재시도 보정 버전 WPF 실행
- Completed: AI 초안 검증 실패 재시도와 처리 UI 보정이 포함된 최신 Debug WPF를 실행했습니다.
- Decisions: 기존 실패 journal과 rejected manifest는 삭제하지 않고 명시적 재시도 입력으로 사용합니다.
- Files: `HANDOFF.md`.
- Verification: `InferenceDataAIService.Wpf.exe` PID 3832가 `HasExited=False`, `Responding=True`, 창 제목 `Inference Data AI` 상태입니다.
- Next: 동일 파일을 다시 선택하고 `실패 단계 재시도`를 눌러 7개 누락 정량 셀 보정 및 5~7단계 완료를 확인합니다.

## 2026-07-23 17:19 - 재실행 후 실패 journal 자동 감지
- Completed: WPF를 다시 시작한 뒤 동일 Excel을 선택하면 `outputs/incremental-ingest`의 실패 journal과 sourcePath를 대조해 `실패 단계 재시도` 버튼을 자동 복원하도록 구현했습니다.
- Decisions: 다른 파일의 실패 journal, 부분 JSON, 접근 불가 artifact는 무시하며 정확히 같은 정규화 경로의 `FAILED` journal만 재시도로 인정합니다.
- Files: `InferenceDataAIService.Wpf/MainWindow.CanonicalEvidence.cs`, `HANDOFF.md`.
- Verification: WPF Debug 빌드가 경고 0, 오류 0으로 통과했고 최신 앱 PID 19212가 `HasExited=False`, `Responding=True`, 창 제목 `Inference Data AI` 상태로 실행됐습니다.
- Next: DRM 전체 처리 탭에서 동일 Excel을 선택하면 실패 journal 안내와 활성화된 재시도 버튼이 표시되는지 확인하고 재시도를 실행합니다.

## 2026-07-23 17:27 - 다중 파일 처리 순서와 단계별 진행률 검토
- Completed: 폴더 입력에서 여러 Excel의 처리 순서와 WPF 단계별 프로그레스 표시 가능성을 조사했습니다.
- Decisions: 현재 WPF는 corpus CLI에 `workbook-workers=1`을 전달하므로 파일별 end-to-end 방식입니다. 즉 파일 A의 COM→패킷→AI→DB→검증이 끝난 뒤 파일 B를 처리하며, 모든 파일의 한 단계를 일괄 완료한 뒤 다음 단계로 넘어가지 않습니다. 기존 per-workbook 진행 이벤트의 sourcePath와 corpus 전체 처리 대상 수를 결합하면 각 단계에 `완료/실패/전체` 프로그레스 바를 추가할 수 있습니다.
- Files: `HANDOFF.md`.
- Verification: `CanonicalEvidenceClient.IngestCorpusAsync`의 worker 설정과 `inference_data_ai_corpus_workflow.run_corpus_ingest`의 workbook 단위 executor 및 각 `ingest_workbook` 호출을 확인했습니다.
- Next: 구현 요청 시 corpus 시작 이벤트로 이번 실행의 대상 파일 수를 전달하고, WPF가 단계별 고유 sourcePath 완료·실패 집합을 집계해 1~6단계 카드에 진행 막대와 `성공 n / 실패 n / 전체 n`을 표시합니다.

## 2026-07-23 17:29 - 다중 파일 병렬화 안전 경계 검토
- Completed: 다수 파일을 직렬 처리할 때의 병목과 현재 코드가 이미 사용하는 병렬 범위를 확인했습니다.
- Decisions: 파일 내부 LOCATOR와 분할 DRAFT는 각각 기본 최대 3 worker로 병렬화되지만 WPF의 corpus workbook worker는 1입니다. 단순히 workbook worker를 늘리면 여러 Excel COM 인스턴스·DRM 인증창·AI 호출·SQLite 쓰기가 동시에 겹칩니다. 권장 구조는 COM 추출 1 worker, 패킷/AI 분석 2~3 worker의 전역 제한, DB import/검증 1 writer인 bounded pipeline입니다. 이 구조는 Excel/SQLite 안정성을 유지하면서 AI 대기 중 다음 파일 COM을 진행할 수 있습니다.
- Files: `HANDOFF.md`.
- Verification: `CanonicalEvidenceClient`의 `--workbook-workers 1`, corpus의 workbook `ThreadPoolExecutor`, workflow의 locator/draft fragment executor, COM `DispatchEx`, SQLite `busy_timeout=60000`을 확인했습니다.
- Next: 구현 요청 시 단계 큐·전역 AI 동시성 제한·단일 DB writer·파일/단계별 progress telemetry와 재시작 가능한 journal을 함께 설계·구현합니다.

## 2026-07-23 18:42 - Corpus bounded pipeline 및 잘못된 정량 셀 분류 수정
- Completed: 여러 workbook workflow가 겹쳐 실행되면서도 전역 `COM=1`, `PACKET=3`, `AI=3`, `DB=1` semaphore를 지키는 bounded pipeline gate를 구현했습니다. COM 추출과 Capture DB 반영을 분리해 다른 파일의 AI 처리 중 다음 파일 COM 추출이 가능하며, AI permit은 실제 locator/draft/fragment 호출 단위로 공유합니다. corpus 시작·종료 시 전체 대상과 pipeline 설정을 progress event로 내보냅니다. 또한 숨김 행에서 빈 입력만 참조해 계산된 0과 Excel 오류 표시값 `#DIV/0!`을 정량 결과 필수 셀로 잘못 요구하던 coverage 분류를 수정했습니다.
- Decisions: workbook thread는 파이프라인을 채우기 위한 운반자이고 실제 자원 동시성은 단계별 전역 gate가 제한합니다. 숨김 formula는 cached 0이면서 참조 경로에 실제 입력이 전혀 없을 때만 `HIDDEN_FORMULA_WITHOUT_SOURCE_INPUT`으로 제외하며, 보이는 동일 formula는 계속 필수 결과로 유지합니다. Excel 오류는 COM 내부 오류 코드가 숫자여도 정량값으로 취급하지 않습니다.
- Files: `inference_data_ai_workflow.py`, `inference_data_ai_corpus_workflow.py`, `inference_data_ai_content_coverage.py`, `inference_data_ai_cli.py`, `tests/test_inference_data_ai_content_coverage.py`, `tests/test_inference_data_ai_corpus_workflow.py`, `tests/test_inference_data_ai_cli.py`, `HANDOFF.md`.
- Verification: 대상 Python 모듈 `py_compile`과 content coverage/corpus/CLI 테스트 90개가 모두 통과했습니다. 실제 실패 workbook의 rejected manifest를 새 inventory로 다시 검증해 필수 정량 셀 62개, `H23/N23` 명시 제외, `I24:M24` 오류 셀 비정량 처리, 미포함 셀 0개를 확인했습니다.
- Next: WPF가 pipeline worker 옵션을 사용하고 corpus 단계별 파일 진행률을 표시하도록 빌드·검증한 뒤, 기존 실패 journal을 재실행해 AI 재호출 없이 다음 DB 단계로 진입하는지 확인합니다.

## 2026-07-23 18:45 - WPF pipeline 진행률 및 실제 실패 workbook 복구 검증
- Completed: WPF 폴더 실행을 `workbook=4, COM=1, PACKET=3, AI=3, DB=1` 구성으로 연결하고, corpus 대상 수와 파일별 stage event를 집계해 각 단계 카드에 progress bar와 `완료/전체·실패`를 표시하도록 구현했습니다. 동일 파일의 실패 journal이 있으면 사용자가 `전체 실행`을 눌러도 자동으로 repair resume을 선택합니다. 기존 실패 workbook을 실제 재개해 Capture/Packet/Locator를 재사용하고 rejected manifest를 새 coverage 규칙으로 승인한 뒤 DB import와 integrity 검증까지 완료했습니다.
- Decisions: 단계 진행률의 분모는 이번 실행에서 실제 처리할 eligible workbook 수이며 기존 완료 skip은 별도 집계합니다. LOCATOR 완료는 4단계 최종 완료로 세지 않고 DRAFT 완료 시 파일 1건을 완료 처리합니다.
- Files: `InferenceDataAIService.Wpf/CanonicalEvidenceClient.cs`, `InferenceDataAIService.Wpf/MainWindow.xaml`, `InferenceDataAIService.Wpf/MainWindow.CanonicalEvidence.cs`, `tests/test_inference_data_ai_workflow.py`, `HANDOFF.md`.
- Verification: 관련 Python 테스트 116개가 모두 통과했고 WPF Debug 빌드는 경고 0, 오류 0입니다. 실제 journal attempt 3은 `NEEDS_REVIEW`, Study 2건, `integrityOk=true`로 종료됐으며 재개는 약 7초에 완료됐습니다.
- Next: 최신 WPF를 실행하고 다중 파일 폴더로 단계별 progress bar와 COM/AI overlap을 화면에서 확인합니다.

## 2026-07-23 18:45 - Bounded pipeline WPF 실행
- Completed: bounded pipeline, 단계별 파일 progress bar, 실패 journal 자동 복구가 포함된 최신 Debug WPF를 실행했습니다.
- Decisions: 검증 통과한 Debug 산출물을 그대로 사용했습니다.
- Files: `HANDOFF.md`.
- Verification: `InferenceDataAIService.Wpf.exe` PID 21076이 `HasExited=False`, `Responding=True`, 창 제목 `Inference Data AI` 상태입니다.
- Next: `Excel 폴더` 모드에서 다수 파일 폴더를 선택해 각 단계의 `n/전체` 막대와 Developer console의 pipeline 진행을 확인합니다.

## 2026-07-23 18:52 - WPF 드롭다운 다크 테마 고정
- Completed: 비활성화되거나 새로 추가된 ComboBox가 Windows 기본 흰 배경으로 바뀌지 않도록 선택 영역·화살표·팝업·항목을 포함한 공통 다크 템플릿을 구현했습니다.
- Decisions: 개별 화면마다 색을 덮어쓰지 않고 App 공통 스타일에서 활성·포커스·비활성 상태를 일관되게 처리합니다.
- Files: `InferenceDataAIService.Wpf/App.xaml`, `HANDOFF.md`.
- Verification: XAML 구현 완료; 홈 화면 변경을 합친 뒤 WPF 프로젝트 Debug 빌드로 검증할 예정입니다.
- Next: 별도 홈 탭과 시작 선택 상태를 추가한 뒤 좁은 WPF 빌드를 실행합니다.

## 2026-07-23 18:54 - WPF 시작 홈 화면
- Completed: 왼쪽 최상단 홈 메뉴와 4단계 주 작업 바로가기, 분석 도구 바로가기를 갖춘 전용 홈 탭을 추가하고 일반 실행 시 홈이 첫 화면이 되도록 변경했습니다.
- Decisions: 명령행 시작 작업이 있으면 기존처럼 Excel 분석 목록에서 시작하고, 일반 실행만 홈으로 진입합니다. 홈과 DRM 전체 처리 화면에서는 불필요한 Insight pane을 숨깁니다.
- Files: `InferenceDataAIService.Wpf/MainWindow.xaml`, `InferenceDataAIService.Wpf/MainWindow.xaml.cs`, `HANDOFF.md`.
- Verification: `dotnet build InferenceDataAIService.Wpf.csproj -c Debug --no-restore -p:OutputPath=bin\CodexVerify\` 통과(경고 0, 오류 0).
- Next: 기존 동일 프로젝트 WPF를 종료하고 정상 Debug 경로로 빌드한 최신 앱을 다시 실행해 시작 상태를 확인합니다.

## 2026-07-23 18:54 - 홈·드롭다운 수정 버전 WPF 재실행
- Completed: 동일 프로젝트의 기존 Debug WPF만 종료하고 홈 화면과 공통 다크 드롭다운이 반영된 정상 Debug 산출물을 빌드해 다시 실행했습니다.
- Decisions: 다른 경로의 동명 프로세스는 건드리지 않고 정확히 이 프로젝트 Debug 실행 파일만 교체했습니다.
- Files: `HANDOFF.md`.
- Verification: 정상 Debug 빌드 경고 0, 오류 0. 새 프로세스 PID 9316은 `HasExited=False`, `Responding=True`, 창 제목 `Inference Data AI`입니다.
- Next: 홈에서 질문 화면으로 이동해 `최신 적재 DB` 드롭다운의 활성·비활성·펼침 상태가 모두 다크 테마로 보이는지 화면에서 확인합니다.

## 2026-07-23 18:55 - DRM 메뉴 진입 종료 원인 수정
- Completed: 홈에서 1번 DRM 메뉴를 처음 열 때 진행 막대의 읽기 전용 `ProgressValue`에 기본 TwoWay 바인딩이 붙어 발생한 `XamlParseException`을 확인하고 `Mode=OneWay`로 수정했습니다.
- Decisions: Windows Application/.NET Runtime 이벤트 1026의 실제 예외를 기준으로 원인을 확정했으며 드롭다운 템플릿과는 무관합니다.
- Files: `InferenceDataAIService.Wpf/MainWindow.xaml`, `HANDOFF.md`.
- Verification: PID 9316 종료와 .NET Runtime의 `TwoWay ... 읽기 전용 속성 ProgressValue` 스택을 확인했습니다. 수정 후 빌드 및 DRM 탭 런타임 재현 검증이 필요합니다.
- Next: WPF Debug 빌드 후 UI 자동화로 홈의 1번 메뉴를 클릭해 프로세스가 유지되는지 확인하고 최신 앱을 다시 실행합니다.

## 2026-07-23 18:57 - DRM 메뉴 런타임 재검증
- Completed: 수정본을 정상 Debug 경로로 다시 빌드·실행하고 UI 자동화로 홈의 1번 DRM 메뉴를 실제 호출한 뒤 홈으로 복귀했습니다.
- Decisions: 탭은 최초 표시 시에만 내부 DataTemplate을 생성하므로 빌드 성공만으로 잡히지 않는 런타임 바인딩 오류를 실제 메뉴 전환으로 검증했습니다.
- Files: `HANDOFF.md`.
- Verification: 빌드 경고 0, 오류 0. DRM 탭 렌더링 후 PID 23280은 `HasExited=False`, `Responding=True`; 홈 복귀 후에도 정상 응답 중입니다.
- Next: 사용자가 드롭다운 활성·비활성 상태와 다중 파일 처리 흐름을 계속 화면 검증합니다.

## 2026-07-23 18:58 - ComboBox 호버 가림 수정
- Completed: ComboBox 클릭 영역의 내부 ToggleButton이 호버 때 Windows 기본 하늘색 배경을 그려 선택 문구를 가리던 문제를 완전 투명한 전용 템플릿으로 수정했습니다.
- Decisions: 실제 배경·테두리 호버 표현은 바깥 `ComboBorder`만 담당하고 내부 ToggleButton은 입력 처리만 담당합니다.
- Files: `InferenceDataAIService.Wpf/App.xaml`, `HANDOFF.md`.
- Verification: XAML 수정 완료; WPF Debug 빌드와 재실행이 필요합니다.
- Next: 정상 Debug 산출물을 빌드해 동일 프로젝트 WPF를 재실행하고 DRM 화면 진입을 재검증합니다.

## 2026-07-23 18:59 - ComboBox 호버 수정본 실행
- Completed: 호버 가림 수정본을 정상 Debug 경로로 빌드하고 동일 프로젝트 WPF를 교체 실행한 뒤 홈에서 1번 DRM 메뉴까지 자동 진입했습니다.
- Decisions: 내부 ToggleButton은 모든 상태에서 투명하며 ComboBox 외곽만 보라색 테두리 호버를 표시합니다.
- Files: `HANDOFF.md`.
- Verification: WPF Debug 빌드 경고 0, 오류 0. PID 21652는 DRM 메뉴 진입 후 `HasExited=False`, `Responding=True`입니다.
- Next: 사용자가 현재 열린 DRM 화면에서 파일/폴더 선택 드롭다운을 호버해 문구와 화살표가 계속 보이는지 확인합니다.

## 2026-07-23 19:12 - 최신 WPF 실행
- Completed: 사용자의 실행 요청에 따라 홈·드롭다운 호버 수정이 반영된 최신 Debug WPF를 실행했습니다.
- Decisions: 동일 프로젝트 프로세스가 없음을 확인하고 새 프로세스를 한 개만 시작했습니다.
- Files: `HANDOFF.md`.
- Verification: PID 14716은 `HasExited=False`, `Responding=True`, 창 제목 `Inference Data AI`입니다.
- Next: 사용자가 화면에서 홈과 DRM 드롭다운을 확인합니다.

## 2026-07-24 07:12 - 다음 작업 우선순위 조사
- Completed: 최신 handoff, table-first 완료 기록, COM-first incremental ingest 산출물, WPF 진행 흐름, 계획 문서와 작업 트리를 대조해 다음 작업을 확정했습니다.
- Decisions: 활성 경로는 완료된 989건 table-first history와 신규 DRM Excel용 canonical incremental ingest입니다. `outputs/corpus-ingest/full-989-v1`의 오래된 RUNNING/FAILED 상태는 2026-07-20 table-first 전환으로 대체됐으므로 재개하지 않습니다. 다음 우선순위는 소규모 다중 파일 폴더로 bounded pipeline의 실제 화면·journal·DB 종단 검증을 완료하는 것입니다.
- Files: `HANDOFF.md`.
- Verification: 최신 table-first 기록은 989/989 성공·실패 0과 `outputs/table-first-history/history.sqlite` 재사용을 명시합니다. 최신 incremental journal attempt 3은 `NEEDS_REVIEW`, Study 2건, `integrityOk=true`이며 Capture/Packet/Locator 재사용 후 import/verify를 완료했습니다. 관련 116개 테스트 통과와 WPF 빌드 경고 0·오류 0은 직전 handoff에서 확인됐습니다. 이번 조사는 소스 변경이 없어 테스트·빌드·앱 실행을 하지 않았습니다.
- Next: 사용자 요청 시 WPF의 `Excel 폴더` 모드에서 3~5개 대표 파일을 처리해 단계별 `n/전체`, COM=1과 AI overlap, 실패 journal 자동 재개, 최종 canonical DB 무결성 및 `최신 적재 DB` 질문 반영을 확인하고 결과를 재현 가능한 acceptance report로 고정합니다.

## 2026-07-24 07:13 - WPF 실행
- Completed: 사용자의 요청에 따라 최신 Debug WPF 실행 파일을 열었습니다.
- Decisions: 동일 프로젝트의 기존 실행 프로세스가 없음을 확인하고 빌드된 실행 파일을 새 프로세스 한 개로 시작했습니다.
- Files: `HANDOFF.md`.
- Verification: PID 22876은 `HasExited=False`, `Responding=True`, 창 제목 `Inference Data AI`, 유효한 메인 창 핸들 상태입니다.
- Next: 열린 WPF에서 홈과 DRM Excel 처리 화면을 사용하거나, 다음 검증 요청을 진행합니다.

## 2026-07-24 07:34 - 적재 내용 검증 창 및 한글 표시
- Completed: DRM Excel 전체 처리 결과에서 manifest의 Study·시험 조건·관측값·비교·제한사항·원본 셀 근거와 관련 Study를 확인하는 `처리 내용 검증` 창을 추가했습니다. 기존 완료 journal도 선택 파일 기준으로 복원하며, manifest의 영문 설명·제한사항·상태 코드는 저장 데이터를 변경하지 않고 검증 창에서 한글로 표시합니다.
- Decisions: canonical manifest와 DB 원문은 추적성을 위해 그대로 보존하고 표시 계층에서만 한글화합니다. 동적 수량·셀 범위·날짜·식별자는 문장 패턴으로 번역하여 원본 값을 유지하며, 알 수 없는 문장은 임의 해석 없이 원문을 유지합니다.
- Files: `InferenceDataAIService.Wpf/CanonicalEvidenceClient.cs`, `InferenceDataAIService.Wpf/MainWindow.xaml`, `InferenceDataAIService.Wpf/MainWindow.CanonicalEvidence.cs`, `InferenceDataAIService.Wpf/IngestVerificationHtmlRenderer.cs`, `InferenceDataAIService.Wpf/IngestVerificationWindow.cs`, `.codex-verify/IngestVerificationRendererCheck/IngestVerificationRendererCheck.csproj`, `.codex-verify/IngestVerificationRendererCheck/Program.cs`, `HANDOFF.md`.
- Verification: WPF Debug 빌드 경고 0·오류 0. 실제 최신 journal/manifest 렌더러 검사에서 journal 복원, Study 2건, 값 255/320, Excel 근거 링크 83개, NEEDS_REVIEW, 한글 제한사항과 한글 요약을 모두 확인했습니다. 최신 Debug WPF PID 22764는 `Responding=True`, 창 제목 `Inference Data AI`입니다.
- Next: 현재 열린 WPF에서 같은 파일을 선택하고 `처리 내용 검증`을 다시 눌러 제한사항 7개와 Study 요약이 한글로 보이는지 화면 확인합니다. 새 영문 문장 패턴이 남으면 해당 원문을 수집해 표시 번역 규칙을 확장합니다.

## 2026-07-24 13:21 - WPF 프로젝트 시작 및 실행
- Completed: 프로젝트의 전체 `HANDOFF.md`를 확인하고 최신 정상 Debug WPF 실행 파일을 열었습니다.
- Decisions: 소스 변경이나 재빌드 요청이 아니므로 2026-07-24 07:34에 생성된 기존 Debug 산출물을 그대로 사용했으며, 실행 전 동일 프로젝트 프로세스가 없음을 확인했습니다.
- Files: `HANDOFF.md`.
- Verification: `InferenceDataAIService.Wpf.exe` PID 22232가 `HasExited=False`, `Responding=True`, 창 제목 `Inference Data AI`, 유효한 메인 창 핸들 상태입니다.
- Next: 열린 WPF 화면에서 사용자가 확인할 기능이나 수정할 항목을 지정하면 해당 범위부터 조사·구현합니다.

## 2026-07-24 13:27 - 처리 내용 검증 비교표 재구성
- Completed: `처리 내용 검증` HTML을 AI 문의 결과와 유사한 검토별 비교표 중심으로 재구성했습니다. 조건을 행으로, 실제 측정 항목을 열로 배치하고 건수와 대응 Rate를 한 셀에 묶었으며 각 값과 조건에서 원본 Excel 셀을 열 수 있게 유지했습니다.
- Decisions: 분석 ID·Revision·manifest/journal 경로·Workbook 서술 요약·별도 시험군/관측값/비교 정의·제한사항·전체 근거 목록·중복/관련 Study는 기본 검증 화면에서 제거합니다. 관측값이 없는 조건은 숨기고 canonical comparison의 `controlArm`/`comparedArm`을 각각 `기준군`/`비교군`으로 표시합니다.
- Files: `InferenceDataAIService.Wpf/IngestVerificationHtmlRenderer.cs`, `HANDOFF.md`.
- Verification: 구현과 HTML/CSS 구조 변경을 완료했습니다. 실제 manifest 렌더링 검사와 WPF 프로젝트 범위 빌드는 다음 단계에서 수행합니다.
- Next: 최신 실제 manifest를 새 렌더러로 변환해 표 개수·행·열·값·Excel 링크 및 불필요 섹션 제거를 확인한 뒤 WPF Debug 프로젝트를 좁게 빌드합니다.

## 2026-07-24 13:29 - 처리 내용 검증 비교표 검증
- Completed: 최신 실제 canonical manifest를 새 렌더러로 변환해 문의 결과형 비교표 구조와 원본 Excel 링크를 검증하고 정상 Debug 실행 파일을 갱신했습니다.
- Decisions: 검증 화면에는 검토별 표, 기준군·비교군, 조건, 실제 측정값만 유지합니다. 건수와 Rate는 같은 지표 셀에 함께 표시하며 값이 없는 `After 1 day check again` 조건은 숨깁니다. 실행 중이던 이전 WPF가 이미 종료되어 앱 재실행 없이 정상 Debug 경로 빌드만 완료했습니다.
- Files: `InferenceDataAIService.Wpf/IngestVerificationHtmlRenderer.cs`, `.codex-verify/IngestVerificationRendererCheck/Program.cs`, `HANDOFF.md`.
- Verification: WPF Debug 빌드가 경고 0·오류 0으로 통과했습니다. 실제 manifest 렌더링에서 비교표 2개, 원본 Excel 링크 66개, 기계 조립 255·수작업 조립 320 값, 기준군·비교군, 검사 수량·전체 NG·Rate 결합 표시를 확인했습니다. 분석 ID·Revision·경로·Workbook 요약·제한사항·별도 근거·중복/관련 Study 등 제거 대상 11개 섹션은 0건이었습니다.
- Next: 사용자가 최신 WPF를 열고 `처리 내용 검증`에서 두 비교표의 실제 화면 밀도와 가로 스크롤을 확인합니다.

## 2026-07-24 13:31 - 비교표 수정본 WPF 실행
- Completed: 처리 내용 검증 비교표 재구성이 반영된 최신 정상 Debug WPF를 실행했습니다.
- Decisions: 동일 프로젝트 실행 프로세스가 없음을 확인한 뒤 새 프로세스 한 개만 시작했습니다. 첫 실행 시도는 PowerShell 작업 폴더 인자 충돌로 시작되지 않았고, 실행 경로 계산을 보정해 재시도했습니다.
- Files: `HANDOFF.md`.
- Verification: `InferenceDataAIService.Wpf.exe` PID 10720이 `HasExited=False`, `Responding=True`, 창 제목 `Inference Data AI`, 유효한 메인 창 핸들 상태입니다.
- Next: 열린 WPF에서 같은 Excel을 선택하고 `처리 내용 검증`을 눌러 검토별 기준군·비교군 측정값 표를 화면 확인합니다.

## 2026-07-24 13:38 - NEEDS_REVIEW 셀 오류·Excel 호버 연결
- Completed: Study 제한사항을 해당 조건·측정값 셀에 연결해 NEEDS_REVIEW 관련 셀을 노란색으로 강조하고 `!` 표시를 추가했습니다. 값이나 조건에 마우스를 올리면 검토 사유, 실제 Excel 시트·셀 범위, 원본 표시값이 나타나며 클릭 시 같은 원본 셀을 엽니다. Study 상태 배지 호버에는 전체 검토 사유를 표시합니다.
- Decisions: 수량 불일치는 Input 셀, Input/OK/전체 NG 합계 불일치는 해당 조건의 세 셀, 원본 NG 건수 셀 누락은 해당 조건의 외관 NG 셀에 연결합니다. 무작위화·날짜·식별자처럼 특정 측정 셀에 귀속할 수 없는 사유는 Study 상태 배지에만 표시하고 임의 셀 오류로 만들지 않습니다.
- Files: `InferenceDataAIService.Wpf/IngestVerificationHtmlRenderer.cs`, `.codex-verify/IngestVerificationRendererCheck/Program.cs`, `HANDOFF.md`.
- Verification: CodexVerify WPF 빌드가 경고 0·오류 0으로 통과했습니다. 최신 실제 manifest 렌더링에서 비교표 2개와 Excel 링크 66개를 유지하면서 review 셀·마커, `Test!G19`/원본값 255 호버, Input 255/320 수량 불일치, I40:L40 외관 NG 건수 누락, Input 311·OK 297·전체 NG 7 불일치 사유가 모두 연결됨을 확인했습니다.
- Next: 실행 중이던 PID 10720이 종료됐으므로 정상 Debug 경로를 다시 빌드한 뒤, 필요 시 최신 WPF를 열어 실제 브라우저 호버 표시를 화면 확인합니다.

## 2026-07-24 13:38 - NEEDS_REVIEW 호버 버전 Debug 빌드
- Completed: 호버 검토 기능이 포함된 정상 Debug WPF 실행 파일을 갱신했습니다.
- Decisions: 앱은 사용자가 이미 종료한 상태였으므로 재실행하지 않고 빌드까지만 수행했습니다.
- Files: `HANDOFF.md`.
- Verification: `dotnet build InferenceDataAIService.Wpf.csproj -c Debug --no-restore`가 경고 0·오류 0으로 통과했습니다.
- Next: 사용자가 실행을 요청하면 최신 WPF를 열고 `처리 내용 검증`의 노란 셀·느낌표·상태 배지에 마우스를 올려 오류 사유와 Excel 좌표를 확인합니다.

## 2026-07-24 13:44 - NEEDS_REVIEW 호버 버전 WPF 실행
- Completed: 오류 사유·Excel 좌표 호버 기능이 포함된 최신 정상 Debug WPF를 실행했습니다.
- Decisions: 동일 프로젝트 실행 프로세스가 없음을 확인한 뒤 새 프로세스 한 개만 시작했습니다.
- Files: `HANDOFF.md`.
- Verification: `InferenceDataAIService.Wpf.exe` PID 20904가 `HasExited=False`, `Responding=True`, 창 제목 `Inference Data AI`, 유효한 메인 창 핸들 상태입니다.
- Next: 열린 WPF에서 `처리 내용 검증`을 누르고 노란 셀·`!`·NEEDS_REVIEW 배지에 마우스를 올려 오류 사유와 원본 Excel 셀 정보를 확인합니다.

## 2026-07-24 13:47 - 검사 수량 경고 의미 확인
- Completed: 사용자가 제공한 처리 내용 비교표와 원본 Excel 이미지를 대조해 검사 수량 경고의 의미를 확인했습니다.
- Decisions: HTML의 기계 조립 255·수작업 조립 320은 원본 Excel Input 255·320과 정확히 일치하므로 추출 오류가 아닙니다. 현재 `!`는 두 비교 조건의 표본 수가 서로 다른 comparability 주의를 셀 오류처럼 표시해 혼동을 일으킵니다.
- Files: `HANDOFF.md`.
- Verification: 이미지에서 HTML과 Excel의 Input, OK 및 NG BAKO 값이 동일함을 육안 대조했습니다. 특히 비교군 간 Input은 255 대 320으로 서로 다르지만 각 값은 원본과 동일합니다.
- Next: 수정 요청 시 추출/원본 불일치만 `오류`로 표시하고, 표본 수 차이·무작위화 부재 같은 비교 설계 사유는 별도 `비교 주의` 배지로 분리합니다.

## 2026-07-24 13:48 - NEEDS_REVIEW 상태 의미 구분
- Completed: `NEEDS_REVIEW`가 원본 추출 신뢰도 문제와 비교·해석 가능성 문제를 함께 담아 발생한 UI 의미 혼선을 정리했습니다.
- Decisions: 원본 셀과 표시값이 일치하는지는 `추출 검증`, 서로 다른 조건을 공정하게 비교할 수 있는지는 `비교 가능성`으로 분리해야 합니다. 255/320 사례는 원본 셀 근거가 일치하므로 추출은 정상이고, 표본 수 불균형 때문에 효과 비교·집계만 사람 확인이 필요한 상태입니다. AI가 만든 비교 역할·설계 해석 자체는 별도 사람 검토 대상입니다.
- Files: `HANDOFF.md`.
- Verification: 제공된 HTML/Excel 이미지에서 원본 수치 일치를 재확인했고, 현재 manifest의 `verificationStatus=NEEDS_REVIEW`와 comparison의 `validityStatus=NEEDS_REVIEW`가 단일 배지로 합쳐져 있음을 기존 렌더링 구조와 대조했습니다.
- Next: 구현 요청 시 화면 상태를 `원본 추출: 일치/불일치`, `비교 가능성: 확인 필요`, `AI 해석: 검토 필요`로 분리하고 셀 `!`는 실제 원본 불일치 또는 산술 무결성 문제에만 사용합니다.

## 2026-07-24 13:50 - 처리 내용 검증 비교 주의 제거
- Completed: 처리 내용 검증 화면에서 표본 수 차이, 무작위화·매칭 정보 부재, 원본의 대조군 미지정과 같은 비교 설계 주의를 제거했습니다. 이 사유들은 상태 배지 툴팁과 셀 노란색·`!` 표시에도 더 이상 포함되지 않습니다.
- Decisions: 셀 검토 표시는 원본 NG 건수 셀 누락, Input·OK·전체 NG 산술 불일치 등 실제 원본 데이터 무결성 문제에만 사용합니다. 기계 조립 255와 수작업 조립 320은 원본과 일치하므로 일반 값 셀로 표시합니다.
- Files: `InferenceDataAIService.Wpf/IngestVerificationHtmlRenderer.cs`, `.codex-verify/IngestVerificationRendererCheck/Program.cs`, `HANDOFF.md`.
- Verification: 실제 manifest 렌더링 검사에서 비교 주의 문구 3종 제거, G19=255와 G21=320 일반 셀 표시, 실제 산술 불일치와 I40:L40 누락 경고 유지, Excel 링크 66개 유지가 모두 통과했습니다. CodexVerify 및 정상 Debug WPF 빌드는 각각 경고 0·오류 0입니다.
- Next: 사용자가 실행을 요청하면 최신 WPF를 열어 첫 번째 기능 검사표의 255/320 셀에서 노란색·`!`가 제거됐는지 화면 확인합니다.

## 2026-07-24 14:17 - 비교 주의 제거 버전 WPF 실행
- Completed: 비교 설계 주의 표시를 제거하고 실제 데이터 문제만 남긴 최신 정상 Debug WPF를 실행했습니다.
- Decisions: 동일 프로젝트 실행 프로세스가 없음을 확인한 뒤 새 프로세스 한 개만 시작했습니다.
- Files: `HANDOFF.md`.
- Verification: `InferenceDataAIService.Wpf.exe` PID 25308이 `HasExited=False`, `Responding=True`, 창 제목 `Inference Data AI`, 유효한 메인 창 핸들 상태입니다.
- Next: 열린 WPF의 `처리 내용 검증`에서 첫 번째 기능 검사표 255/320 셀의 노란색·`!` 제거와 실제 데이터 오류 셀 경고 유지를 확인합니다.

## 2026-07-24 14:24 - 다중 Excel 동일 파일 스킵 동작 확인
- Completed: WPF 폴더 전체 처리의 corpus inventory, DB 기존 분석 reconciliation, 완료 journal 스킵 및 Capture DB 중복 방지 조건을 확인했습니다.
- Decisions: 동일한 절대 원본 경로와 동일한 SHA-256 내용이 현재 canonical DB 분석에 있으면 `RECONCILED_EXISTING`, 같은 corpus journal에서 완료됐으면 `SKIPPED_COMPLETED`로 처리해 COM·AI 파이프라인을 다시 실행하지 않습니다. 같은 바이트라도 파일명이나 폴더가 달라 절대 경로가 바뀌면 별도 원본으로 처리합니다. 같은 경로의 파일 내용이 바뀌면 새 fingerprint/revision으로 다시 처리합니다.
- Files: `HANDOFF.md`.
- Verification: `inference_data_ai_corpus_workflow.py`의 record ID가 `sourcePath + contentSha256`, DB reconciliation key가 `source_path + content_sha256`, 완료 상태의 `SKIPPED_COMPLETED`, `inference_data_ai_source_ingest.py`의 현재 동일 revision `SKIPPED` 조건을 확인했습니다. 회귀 테스트도 반복 실행 시 `attempted=0`, `skippedCompleted=2`를 검증합니다.
- Next: UI 개선 요청 시 DB reconciliation 건수를 `DB 기존 자료 n개 자동 스킵`으로 별도 표시하고, 동일 내용·다른 경로의 중복 정책을 사용자가 선택할 수 있게 합니다.

## 2026-07-24 14:28 - DRM 해지 접미사 제외 파일명 중복 스킵 구현
- Completed: 신규 DRM 원본 이름과 기존 DB 해지본 이름을 비교할 때 파일명 끝의 `_9~13자리 숫자_clean` 또는 `_clean`을 제거한 정규화 파일명으로 중복을 감지하도록 corpus DB reconciliation을 확장했습니다. WPF 폴더 진행 상태에도 `DB 기존 파일명 n개 자동 스킵`을 표시합니다.
- Decisions: 사용자가 표시한 `_1778470442_clean`은 DRM 해지 과정에서 생성된 접미사로 간주해 제외합니다. `_250119` 같은 기존 날짜·식별 숫자는 뒤에 `_clean`이 직접 결합된 생성 패턴이 아니므로 보존합니다. 정규화 파일명이 현재 canonical DB 분석과 일치하면 원본 경로와 SHA가 달라도 COM·AI를 실행하지 않으며, 결과에 실제 입력 경로와 매칭된 DB 해지본 경로를 모두 기록합니다.
- Files: `inference_data_ai_corpus_workflow.py`, `InferenceDataAIService.Wpf/MainWindow.CanonicalEvidence.cs`, `tests/test_inference_data_ai_corpus_workflow.py`, `HANDOFF.md`.
- Verification: 활성 DB 원본 991건을 조사해 정규화 이름 990개를 확인했고, 예시 해지본이 `00. report test new dry machine with material make press jig 2024.03.26.xlsx`로 정규화됨을 확인했습니다. 구현 회귀 테스트와 WPF 빌드는 다음 단계에서 수행합니다.
- Next: 파일명 정규화·DB reconciliation 단위 테스트, corpus 관련 테스트, Python compile 및 WPF Debug 범위 빌드를 실행합니다.

## 2026-07-24 14:29 - 파일명 중복 스킵 검증
- Completed: DRM 접미사 제거 규칙과 파일명 기반 DB reconciliation을 단위 테스트 및 실제 운영 DB로 검증하고 정상 Debug WPF 실행 파일을 갱신했습니다.
- Decisions: 접미사 없는 신규 원본의 바이트가 기존 해지본과 달라도 정규화 파일명이 같으면 기존 canonical 분석을 재사용합니다. 매칭 결과는 `duplicateMatchKind=NORMALIZED_FILENAME`, 신규 입력 path/hash와 기존 DB 매칭 path/hash를 함께 보존합니다.
- Files: `inference_data_ai_corpus_workflow.py`, `InferenceDataAIService.Wpf/MainWindow.CanonicalEvidence.cs`, `tests/test_inference_data_ai_corpus_workflow.py`, `HANDOFF.md`.
- Verification: `py_compile` 통과, `tests.test_inference_data_ai_corpus_workflow` 11/11 통과, CodexVerify 및 정상 Debug WPF 빌드 경고 0·오류 0입니다. 실제 DB에서 `00. Report TEST ... 2024.03.26.xlsx`가 기존 `...2024.03.26_1778470442_clean.xlsx`와 매칭되어 상태 `COMPLETED`, `NORMALIZED_FILENAME`으로 스킵됨을 확인했습니다.
- Next: 사용자가 실행을 요청하면 최신 WPF를 열고 원본 파일 폴더 전체 처리에서 `DB 기존 파일명 1개 자동 스킵` 표시와 처리 대상 0건을 확인합니다.

## 2026-07-24 14:32 - pending 49개 DB 파일명 대조
- Completed: `D:\000. MyWorks\test\org\pending`의 Excel 49개를 DRM 해지 접미사 제외 파일명으로 현재 canonical DB 활성 원본 및 table-first 전체 이력과 대조했습니다.
- Decisions: 중복 판정은 구현된 `_숫자_clean`/`_clean` 제거 규칙을 사용했습니다. pending에 숫자만 붙은 두 파일은 숫자 단독 접미사까지 제거하는 추가 보수적 대조도 수행했습니다.
- Files: `HANDOFF.md`.
- Verification: pending 49개, 내부 정규화 중복 그룹 0개. canonical DB 활성 원본 991개와 일치 0개, table-first history 989개와 일치 0개입니다. 숫자 단독 접미사까지 제거해 두 DB를 합산 대조해도 일치 0개였습니다.
- Next: 현재 기준으로 pending 49개는 모두 신규 처리 대상입니다. 사용자가 요청하면 최신 WPF로 해당 폴더 전체 처리를 실행합니다.

## 2026-07-24 14:39 - pending 폴더와 DRM 해지 파일 관계 재검증
- Completed: `D:\000. MyWorks\test\org\pending` 49개가 현재 DB 해지본과 매칭되지 않은 원인을 DRM 도구 코드와 실제 폴더 이력으로 재검증했습니다.
- Decisions: 실제 파일명 흐름은 `원본명.xlsx` → `원본명_<숫자>.xlsx` → `원본명_<숫자>_clean.xlsx`가 맞습니다. 다만 `pending`은 정상 처리 원본 보관소가 아니라 사전 스캔에서 도형 수가 100개를 초과한 파일을 DRM 처리 전에 분리하는 폴더이므로, 이 49개에는 대응하는 `_숫자_clean` 파일과 DB 레코드가 없는 것으로 판단합니다.
- Files: `HANDOFF.md` (조사 대상 코드 `External/ExcelDrmCli/excel_drm_clean.py`는 변경 없음).
- Verification: pending 49개를 `result\Org`의 `_숫자` 파일 및 `result\InputDataFinish`의 `_숫자_clean` 파일과 정확 대조해 각각 0건을 확인했습니다. `result\Org` 993개와 pending 49개는 모두 2026-05-08 17:15~17:17에 분리됐고, DRM 도구의 `HEAVY_SHAPE_THRESHOLD = 100` 및 초과 파일의 `pending` 즉시 이동 로직을 확인했습니다. 정상 처리 예시는 `...2024.03.26_1778470442.xlsx` → `...2024.03.26_1778470442_clean.xlsx`로 확인했습니다.
- Next: pending 49개는 DB 중복이 아니라 DRM 사전 분리 파일입니다. 사용자가 요청하면 대용량 도형 파일용 별도 처리 방식으로 해지·적재 가능 여부를 검토합니다.

## 2026-07-24 17:08 - DB 전체 파일명 중복 탐색 확대 및 WPF 실행
- Completed: 신규 엑셀 중복 탐색 범위를 canonical 분석이 있는 파일에서 활성 DB 원본 전체로 확대하고, 최신 WPF를 실행했습니다.
- Decisions: 활성 `source_documents`와 현재 `source_revisions`에 파일이 있으면 canonical 분석 유무와 관계없이 COM·AI 재실행을 건너뜁니다. 분석이 없는 DB 원본 중복은 결과 상태 `EXISTING_SOURCE`로 기록합니다. 파일명 비교 시 끝의 `_clean`과 9~13자리 `_숫자` 접미사를 반복 제거하여 `원본명.xlsx`, `원본명_숫자.xlsx`, `원본명_숫자_clean.xlsx`, `원본명_숫자_숫자_clean.xlsx`를 같은 원본으로 취급합니다.
- Files: `inference_data_ai_corpus_workflow.py`, `tests/test_inference_data_ai_corpus_workflow.py`, `HANDOFF.md`.
- Verification: `py_compile` 통과, corpus workflow 단위 테스트 12/12 통과. 실제 운영 DB에서 활성 현재 원본 991개·canonical 표시 원본 49개를 확인했고, canonical 분석이 없는 기존 `_clean` 파일도 신규 원본명으로 `NORMALIZED_FILENAME`/`EXISTING_SOURCE` 자동 스킵됨을 확인했습니다. 새 정규화 기준 DB 고유 파일명은 990개이며 기존 중복 그룹 1개뿐입니다. pending 49개는 새 기준에서도 DB 매칭 0개입니다. Debug WPF를 PID 24812로 실행했고 창 제목 `Inference Data AI`, 응답 상태 정상입니다.
- Next: 실행 중인 WPF에서 기존 DB 파일과 같은 원본을 추가하면 `DB 기존 파일명 n개 자동 스킵`으로 표시되는지 확인합니다. pending 49개는 DB에 없으므로 스킵되지 않습니다.

## 2026-07-24 17:13 - 읽기 전용 네트워크 Excel 파일명 검색 UI
- Completed: `DRM Excel 전체 처리` 화면의 원본 선택 옆에 `엑셀 검색` 버튼과 결과 표를 추가했습니다. 사용자가 폴더를 지정하면 하위 Excel 파일명을 읽어 현재 DB의 활성 원본 전체와 비교하고 `있음`/`없음`을 표시합니다.
- Decisions: 검색 기능은 처리 경로와 완전히 분리해 선택 경로를 적재 입력란에 넣지 않습니다. 대상 폴더에서는 `.xlsx/.xlsm/.xlsb/.xls` 이름과 상대 하위 폴더명만 열람하며 파일 열기·내용 읽기·해시·복사·이동·수정·삭제·Excel COM을 실행하지 않습니다. 재분석 파일명 규칙과 동일하게 `_clean` 및 반복된 9~13자리 `_숫자` 접미사를 제외합니다. DB는 `SqliteOpenMode.ReadOnly` 및 private cache로만 조회합니다.
- Files: `InferenceDataAIService.Wpf/ExcelFolderSearchService.cs`, `InferenceDataAIService.Wpf/MainWindow.xaml`, `InferenceDataAIService.Wpf/MainWindow.CanonicalEvidence.cs`, `HANDOFF.md`.
- Verification: 별도 `CodexVerifyExcelSearch` 빌드와 정상 Debug 빌드 모두 경고 0·오류 0입니다. 검색 서비스 정적 확인에서 대상 폴더 접근은 `Directory.EnumerateFiles`만 있고 파일 내용/쓰기 API 및 Office/COM 호출이 없으며 DB 연결은 `SqliteOpenMode.ReadOnly`임을 확인했습니다. 최신 WPF를 PID 22516으로 실행했고 창 제목 `Inference Data AI`, 응답 상태 정상입니다.
- Next: 실행 중인 WPF의 `DRM Excel 전체 처리` 화면에서 `엑셀 검색`을 눌러 실제 네트워크 드라이브를 사용자가 선택하고, 요약의 전체·DB 있음·DB 없음 및 결과 표를 확인합니다.

## 2026-07-24 17:24 - 검색 결과 색상·행 삭제 및 원본 보존 로컬 복사 처리
- Completed: 엑셀 검색 결과의 `있음`을 초록색, `없음`을 빨간색으로 표시하고, `있음` 행에 `삭제` 버튼을 추가했습니다. `DB 없음 로컬 복사 후 처리` 버튼도 추가해 DB에 없는 파일만 로컬로 복사한 뒤 기존 전체 처리를 시작합니다.
- Decisions: `삭제`는 `ObservableCollection`의 검색 결과 행만 제거하며 원본 파일과 DB를 변경하지 않습니다. 처리 시 네트워크 원본은 절대 이동·삭제·수정하지 않고 `File.Copy`만 사용합니다. 복사 위치는 DB와 같은 `outputs\universal-grid` 아래 `excel-search-local\<시각_고유값>`이며 하위 상대 경로를 보존합니다. 모든 복사가 끝나고 파일 크기 검증이 통과한 경우에만 입력 경로를 로컬 복사 폴더로 바꿔 전체 처리를 시작합니다.
- Files: `InferenceDataAIService.Wpf/ExcelFolderSearchService.cs`, `InferenceDataAIService.Wpf/ExcelLocalCopyService.cs`, `InferenceDataAIService.Wpf/MainWindow.xaml`, `InferenceDataAIService.Wpf/MainWindow.CanonicalEvidence.cs`, `HANDOFF.md`.
- Verification: 별도 `CodexVerifyExcelSearchCopy` 빌드와 정상 Debug 빌드 모두 경고 0·오류 0입니다. 정적 확인 결과 원본 파일 작업은 `File.Copy(..., overwrite: false)`만 존재하고 `File.Move`, `File.Delete`, `Directory.Move/Delete` 호출은 없습니다. `삭제` 이벤트는 `_excelFolderSearchRows.Remove(row)`만 수행합니다. 최신 Debug WPF를 PID 23396으로 실행했고 창 제목 `Inference Data AI`, 응답 상태 정상입니다.
- Next: 실행 중인 WPF에서 네트워크 폴더를 검색해 색상과 행 삭제를 확인하고, `DB 없음 로컬 복사 후 처리`를 눌렀을 때 로컬 복사 확인창과 `excel-search-local` 경로를 확인합니다.

## 2026-07-24 17:25 - Excel 영구 보관함 분리
- Completed: 검색 결과의 DB 없음 파일을 처리 전에 복사하는 위치를 임시성 `excel-search-local`에서 전용 영구 보관함 `ExcelFileArchive`로 변경했습니다.
- Decisions: 보관함은 DB 파일과 같은 `outputs\universal-grid` 아래에 두며 `ExcelFileArchive\YYYY-MM-DD\HHmmss_fff_<고유값>` 구조로 배치별 보관합니다. 하위 원본 상대 경로를 보존하고, 모든 파일 복사와 크기 검증이 완료된 보관 배치만 전체 처리 입력으로 사용합니다. 네트워크 원본은 계속 `File.Copy`의 읽기 원본일 뿐 이동·삭제·수정하지 않습니다.
- Files: `InferenceDataAIService.Wpf/ExcelLocalCopyService.cs`, `InferenceDataAIService.Wpf/MainWindow.xaml`, `InferenceDataAIService.Wpf/MainWindow.CanonicalEvidence.cs`, `HANDOFF.md`.
- Verification: 별도 `CodexVerifyExcelArchive` 빌드와 정상 Debug 빌드 모두 경고 0·오류 0입니다. 정적 확인에서 `ExcelFileArchive` 경로와 `File.Copy(..., overwrite: false)`만 존재하며 원본 `Move/Delete` 호출이 없음을 확인했습니다. 최신 WPF를 PID 10380으로 실행했고 창 제목 `Inference Data AI`, 응답 상태 정상입니다.
- Next: 사용자가 `DB 없음 보관함 복사 후 처리`를 실행하면 보관함이 실제 생성되고 날짜·배치 하위에 복사본이 만들어지는지 확인합니다.

## 2026-07-24 17:37 - WPF 경로 설정 메뉴 및 동적 적용
- Completed: 좌측 `SYSTEM > 설정` 화면을 추가하고 서비스 폴더, Python/Codex 실행 파일, 현재·이력 DB, 출력 루트, 배치·로그·임시 파일·Excel 보관함·적재·질문·검토·Manifest 폴더와 Benchmark JSON을 찾아보기/직접 입력하여 저장하고 즉시 적용하도록 구현했습니다. 설정은 사용자 LocalAppData의 JSON에 영속화되며 다음 실행 때 자동 로드됩니다.
- Decisions: 네트워크 원본 보존 정책은 유지하고 Excel 복사 보관함만 설정 경로를 사용합니다. 설정 저장 중 실행 작업이 있으면 변경을 막고, 경로 정규화·서비스 CLI·실행 파일을 검증한 뒤 클라이언트를 재생성합니다. 배치 내부의 `batch.json`·`logs` 같은 계약 파일명은 유지하되 그 상위 루트는 설정으로 이동했습니다. 처리 내용 검증용 임시 HTML도 OS 임시 폴더 하드코딩을 제거하고 설정 경로를 사용합니다.
- Files: `InferenceDataAIService.Wpf/AppPathSettings.cs`, `InferenceDataAIService.Wpf/MainWindow.Settings.cs`, `InferenceDataAIService.Wpf/MainWindow.xaml`, `InferenceDataAIService.Wpf/MainWindow.xaml.cs`, `InferenceDataAIService.Wpf/MainWindow.CanonicalEvidence.cs`, `InferenceDataAIService.Wpf/CanonicalEvidenceClient.cs`, `InferenceDataAIService.Wpf/WorkbookComparisonClient.cs`, `InferenceDataAIService.Wpf/IngestVerificationWindow.cs`, `InferenceDataAIService.Wpf/ExcelLocalCopyService.cs`, `InferenceDataAIService.Wpf/StructureScanEngine.cs`, `InferenceDataAIService.Wpf/NumericCaptureEngine.cs`, `InferenceDataAIService.Wpf/NumericReviewEngine.cs`, `InferenceDataAIService.Wpf/NumericRendererEngine.cs`, `InferenceDataAIService.Wpf/TwoLevelGroupAnalysisEngine.cs`, `InferenceDataAIService.Wpf/GroupCatalogRendererEngine.cs`, `tools/WorkbookComparisonAudit/Program.cs`, `tools/WorkbookComparisonAudit/WorkbookComparisonAudit.csproj`, `HANDOFF.md`.
- Verification: WPF 좁은 Debug 빌드 `CodexVerifySettingsFinal`이 경고 0·오류 0으로 통과했고, 연동 감사 도구 `WorkbookComparisonAudit` 좁은 빌드도 경고 0·오류 0으로 통과했습니다. WPF 소스 정적 검색에서 환경별 `outputs` 루트는 기본값 생성부에만 남고 실제 실행부는 설정 객체를 사용하며 `GetTempPath()` 호출과 절대 드라이브 경로 하드코딩이 없음을 확인했습니다. AGENTS 실행 규칙에 따라 이 단계에서는 앱을 새로 실행하지 않았습니다.
- Next: 사용자가 요청하면 최신 Debug WPF를 실행해 `설정` 메뉴에서 경로 변경·저장 후 DB 표시 및 다음 작업 적용을 UI로 확인합니다.

## 2026-07-24 17:38 - 최신 설정 UI WPF 실행
- Completed: 설정 메뉴 변경분을 정상 Debug 출력에 다시 빌드하고 WPF를 실행했습니다.
- Decisions: 실행 중인 이전 프로세스가 없어 별도 종료 작업 없이 최신 실행 파일을 시작했습니다.
- Files: `HANDOFF.md` (애플리케이션 코드는 변경 없음).
- Verification: Debug 빌드 경고 0·오류 0, 실행 프로세스 PID 23456, 창 제목 `Inference Data AI`, `Responding=True`를 확인했습니다.
- Next: 실행된 창의 좌측 `SYSTEM > 설정`에서 원하는 DB·폴더·실행 파일 경로를 지정하고 `저장·즉시 적용`을 누릅니다.

## 2026-07-24 17:45 - Excel DB 단일 통합 폴더 설정
- Completed: 설정 화면의 DB·보관함·배치·로그·임시·검수 등 개별 경로 입력을 제거하고 `Excel DB 통합 폴더` 하나로 축소했습니다. 서비스 폴더와 Python/Codex 실행 파일만 별도 유지하며, 통합 폴더를 바꾸면 모든 DB화 산출물 경로를 즉시 자동 파생합니다.
- Decisions: 기존 JSON의 개별 산출물 경로가 있더라도 앞으로는 `OutputRootDirectory` 하나를 기준으로 DB, Excel 보관함, 처리 결과, 로그, 질문·검토 결과와 Benchmark 경로를 구성합니다. 사용자가 지정한 `D:\000. MyWorks\002. DB`에는 다른 시스템 자료와 `logs`가 이미 있으므로 `D:\000. MyWorks\002. DB\InferenceDataAIService` 전용 통합 폴더로 격리합니다.
- Files: `InferenceDataAIService.Wpf/AppPathSettings.cs`, `InferenceDataAIService.Wpf/MainWindow.Settings.cs`, `InferenceDataAIService.Wpf/MainWindow.xaml`, `HANDOFF.md`.
- Verification: WPF `CodexVerifyUnifiedDbRoot` 빌드와 `WorkbookComparisonAudit` 연동 빌드가 각각 경고 0·오류 0으로 통과했습니다. 기존 `outputs`는 24,269개·12,810,869,724바이트이고 대상 부모는 쓰기 가능함을 확인했습니다.
- Next: PID 4068의 기존 `ingest-corpus`가 DB를 갱신 중이므로 종료까지 보존합니다. 현재 진행 중인 1차 복사 후 종료 시 변경분 재동기화, 파일 수·용량 검증, 기존 `outputs` 정리, 설정 JSON 적용을 순서대로 완료합니다.

## 2026-07-24 18:47 - 통합 DB 폴더 이전 및 기존 로컬 출력 정리
- Completed: 진행 중이던 PID 4068 `ingest-corpus`를 강제 종료하지 않고 최종 상태까지 기다린 뒤, 기존 `outputs` 전체를 `D:\000. MyWorks\002. DB\InferenceDataAIService`로 최종 동기화했습니다. 사용자 설정 JSON의 통합 루트를 새 위치로 지정하고 기존 로컬 `InferenceDataAIService\outputs`는 Windows 휴지통으로 이동했습니다. 네트워크 원본은 건드리지 않았습니다.
- Decisions: 지정 부모 `D:\000. MyWorks\002. DB`에 기존 다른 시스템 DB와 `logs`가 있어 충돌 방지를 위해 `InferenceDataAIService` 전용 하위 폴더를 사용합니다. 직접 영구 삭제 명령은 실행 환경 정책이 막아 복구 가능한 휴지통 이동으로 처리했습니다. 적재 종료 결과 `COMPLETED_WITH_ERRORS`는 완료 5·실패 27이며 journal과 모든 실패 근거도 함께 보존했습니다.
- Files: `C:\Users\jhbyun\AppData\Local\InferenceDataAIService\settings.json`, `D:\000. MyWorks\002. DB\InferenceDataAIService\` 전체 산출물, `HANDOFF.md`; 기존 `D:\000. MyWorks\005. Program\Repository\JinoSupporter\InferenceDataAIService\outputs`는 휴지통으로 이동.
- Verification: 최종 원본/대상 대조에서 양쪽 24,449개·12,825,744,711바이트, 누락 0·크기 불일치 0·추가 0을 확인한 뒤 정리했습니다. 현재 DB 3,287,224,320바이트와 이력 DB 71,790,592바이트의 원본/대상 SHA-256이 각각 일치했고, 새 위치 두 DB의 `PRAGMA quick_check`가 모두 `ok`였습니다. 정리 후 기존 경로 없음, 새 경로 및 두 핵심 DB 존재를 확인했으며 정상 Debug WPF 빌드도 경고 0·오류 0으로 통과했습니다.
- Next: WPF를 실행하면 설정 화면에는 `Excel DB 통합 폴더` 하나가 먼저 표시되고 새 루트가 자동 로드됩니다. 휴지통의 이전 `outputs`는 사용자가 복구 가능하며, 용량을 완전히 비우려면 이후 휴지통을 비웁니다.

## 2026-07-25 07:19 - 최근 작업 이력 점검
- Completed: 사용자의 요청에 따라 프로젝트 `HANDOFF.md`, Git 최근 커밋, 작업 트리 상태와 파일 수정 시각을 대조해 2026-07-24의 최신 작업 범위를 확인했습니다.
- Decisions: 최근 상태는 Git 커밋보다 아직 커밋되지 않은 `HANDOFF.md` 기록과 작업 트리 변경이 더 최신이므로, 17:08~18:47 단계의 네트워크 Excel 검색·중복 스킵, WPF 통합 경로 설정, 산출물 이전을 현재 상태로 봅니다.
- Files: `HANDOFF.md`만 갱신했으며 소스 코드는 변경하지 않았습니다.
- Verification: 최신 HANDOFF 항목 20개와 최근 수정 파일을 확인했고, 마지막 소스 수정은 WPF 설정·경로 관련 파일(17:44), 마지막 완료 기록은 통합 DB 폴더 이전(18:47)임을 확인했습니다.
- Next: 사용자에게 최근 작업 요약과 남은 확인 사항을 전달합니다.

## 2026-07-25 07:21 - 최신 WPF 실행
- Completed: 사용자의 명시적 요청에 따라 최신 정상 Debug WPF 실행 파일을 열었습니다.
- Decisions: 소스 변경이 없고 2026-07-24 17:46 빌드가 최신이므로 재빌드하지 않고 `net9.0-windows` Debug 실행 파일을 사용했습니다.
- Files: `HANDOFF.md`만 갱신했으며 애플리케이션 소스와 빌드 산출물은 변경하지 않았습니다.
- Verification: 프로세스 PID 24044가 종료되지 않았고 `Responding=True`, 창 제목 `Inference Data AI`임을 확인했습니다.
- Next: 사용자가 실행된 WPF에서 통합 DB 경로 및 필요한 화면을 확인하고 추가 요청을 전달합니다.

## 2026-07-25 07:23 - 메인 메뉴 역할과 혼동 원인 분석
- Completed: `DRM Excel 전체 처리`, `Excel ↔ DB 검수`, `사람 검토 승인`, `질문 관련 보고서`의 실제 입력·동작·DB 변경 여부를 WPF 화면과 클라이언트 코드 기준으로 비교했습니다.
- Decisions: 현재 번호는 필수 직렬 단계처럼 보이지만 실제로는 1번이 DB 등록, 2번이 읽기 전용 표본 품질검사, 3번이 보류 비교의 DB 판정 변경, 4번이 DB 검색·활용입니다. 일상 작업과 품질관리 작업을 분리하고 `Excel 가져오기`, `질문으로 자료 찾기`, `원본 값 대조`, `보류 항목 판정`처럼 행동 중심으로 재명명하는 것이 적절합니다.
- Files: `HANDOFF.md`만 갱신했으며 UI 소스는 변경하지 않았습니다.
- Verification: `MainWindow.xaml`의 네 화면 입력·버튼·결과 표, `MainWindow.xaml.cs`의 메뉴 설명, `WorkbookComparisonClient.cs`의 read-only DB 연결, `CanonicalEvidenceClient.cs`의 적재·검색·review-decision CLI 호출을 대조했습니다.
- Next: 사용자 동의 시 번호를 제거하고 `주요 작업`(Excel 가져오기, 질문으로 자료 찾기)과 `품질 관리`(원본 값 대조, 보류 항목 판정)로 메뉴를 재구성하며 각 메뉴에 한 줄 설명을 표시합니다.

## 2026-07-25 07:31 - 신규 Excel 수집 메뉴 분리 및 메뉴 구조 개편
- Completed: 번호 기반 메인 메뉴를 `주요 작업`과 `품질 관리`로 재구성하고, `신규 Excel 수집`을 `DRM Excel 전체 처리` 위에 추가했습니다. 기존 전체 처리 화면의 Excel 검색·DB 비교·보관함 복사 UI를 새 화면으로 옮겼으며, DB 없음 파일 복사 후 자동 전체 처리를 시작하던 연결을 제거했습니다.
- Decisions: 새 수집 화면의 최종 동작은 선택 폴더의 Excel 파일명을 DB와 비교하고 DB에 없는 파일만 `ExcelFileArchive`로 일괄 복사·크기 검증한 뒤 종료하는 것입니다. 네트워크 원본은 보존하고 AI·Excel COM·DB 적재는 실행하지 않습니다. 복사 성공 후 같은 검색 결과의 중복 복사 버튼은 잠급니다.
- Files: `InferenceDataAIService.Wpf/MainWindow.xaml`, `InferenceDataAIService.Wpf/MainWindow.xaml.cs`, `InferenceDataAIService.Wpf/MainWindow.CanonicalEvidence.cs`, `InferenceDataAIService.Wpf/MainWindow.Settings.cs`, `HANDOFF.md`.
- Verification: 좁은 WPF 빌드 `CodexVerifyExcelCollection`이 경고 0·오류 0으로 통과했습니다. 정적 계약 검사에서 새 화면·검색 버튼·복사 버튼은 각각 1개, 이전 `CopyAndProcess` 참조는 0개, 복사 핸들러의 적재 호출은 없고 `ExcelLocalCopyService.CopyAsync`만 호출함을 확인했습니다. 보관 서비스는 `File.Copy` 1곳, 원본 Move/Delete 호출 0곳입니다.
- Next: 현재 실행 중인 PID 24044는 이전 빌드이므로 사용자가 요청하면 종료 후 정상 Debug 출력으로 다시 빌드하고 최신 WPF를 재실행해 새 메뉴 배치와 수집 화면을 확인합니다.

## 2026-07-25 07:32 - 메뉴 개편 최신 WPF 실행
- Completed: 정상 Debug 출력을 최신 소스로 다시 빌드하고 메뉴 개편과 신규 Excel 수집 화면이 포함된 WPF를 실행했습니다.
- Decisions: 이전 Debug 경로에서 실행 중인 프로세스가 없어 강제 종료 없이 최신 실행 파일을 시작했습니다.
- Files: 정상 Debug 빌드 산출물과 `HANDOFF.md`; 애플리케이션 소스는 추가 변경하지 않았습니다.
- Verification: WPF Debug 빌드가 경고 0·오류 0으로 통과했고, PID 19404가 `HasExited=False`, `Responding=True`, 창 제목 `Inference Data AI` 상태임을 확인했습니다.
- Next: 열린 WPF에서 `주요 작업 > 신규 Excel 수집`을 선택해 검색 폴더·로컬 보관함·DB 비교 결과와 `DB 없음 모두 보관함으로 복사` 버튼을 확인합니다.

## 2026-07-25 07:40 - 질문 화면 정리 및 Excel 수집 직접 저장·DB 원본 열기
- Completed: 질문 화면의 세로로 늘어난 실행 버튼 행을 입력·실행·결과 도구·근거 목록으로 재배치하고, 질문 결과가 없을 때 이전 분석 보고서 대신 질문 전용 안내를 표시하도록 결과 패널 상태를 분리했습니다. Excel 수집은 날짜·배치·상대 하위 폴더 없이 지정 보관함 최상위에 직접 저장하도록 바꾸고, DB 있음 행을 클릭하면 검색 파일이 아니라 `source_documents.source_path`의 DB 연결 Excel을 열도록 구현했습니다.
- Decisions: DB 상태는 `DB 있음 · 원본 열기`와 `DB 없음 · 수집 대상`으로 명시합니다. 보관함에는 네트워크 원본을 삭제하지 않고 평면 복사하며, 같은 이름의 다른 파일은 `(2)` 접미사로 충돌을 피하고 동일 파일 재수집은 SHA-256으로 중복 생성하지 않습니다.
- Files: `InferenceDataAIService.Wpf/MainWindow.xaml`, `InferenceDataAIService.Wpf/MainWindow.xaml.cs`, `InferenceDataAIService.Wpf/MainWindow.CanonicalEvidence.cs`, `InferenceDataAIService.Wpf/ExcelFolderSearchService.cs`, `InferenceDataAIService.Wpf/ExcelLocalCopyService.cs`, `.codex-verify/ExcelCollectionBehavior/ExcelCollectionBehavior.csproj`, `.codex-verify/ExcelCollectionBehavior/Program.cs`, `HANDOFF.md`.
- Verification: 첫 WPF 빌드가 제거된 DB 경로 컨트롤 참조 2건을 찾아 `검색 DB` 읽기 전용 행으로 복구한 뒤, 좁은 WPF 빌드 `CodexVerifyQuestionAndCollection`이 경고 0·오류 0으로 통과했습니다. 전용 동작 검증은 DB 정규화 매칭이 연결 `source_path`를 반환하고, 보관함 하위 폴더 0개, 같은 이름 평면 충돌 처리, 반복 수집 중복 방지를 모두 PASS했습니다.
- Next: 실행 중인 이전 PID 19404를 정상 종료하고 정상 Debug 출력을 갱신해 최신 WPF에서 질문 화면과 `DB 있음 · 원본 열기` 동작을 확인합니다.

## 2026-07-25 07:41 - 질문·Excel 수집 개선 WPF 실행
- Completed: 질문 화면 정리, 지정 보관함 직접 저장, DB 있음 원본 열기 변경을 정상 Debug 출력으로 빌드하고 최신 WPF를 실행했습니다.
- Decisions: Debug 경로에서 실행 중인 이전 프로세스가 없어 강제 종료 없이 새 실행 파일을 시작했습니다.
- Files: 정상 Debug 빌드 산출물과 `HANDOFF.md`; 소스는 직전 단계에서 확정했습니다.
- Verification: 정상 Debug 빌드가 경고 0·오류 0으로 통과했고, PID 6684가 `HasExited=False`, `Responding=True`, 창 제목 `Inference Data AI` 상태입니다.
- Next: 열린 WPF의 신규 Excel 수집에서 빨간 `DB 없음 · 수집 대상`과 초록 `DB 있음 · 원본 열기`를 확인하고, DB 있음 버튼을 눌러 DB 연결 Excel 경로가 열리는지 사용자 화면에서 확인합니다.

## 2026-07-25 07:43 - 기존 Excel 보관함 복사본 평면 이동
- Completed: 이전 실행이 `ExcelFileArchive\2026-07-25\073424_605_6442a2e6` 아래에 저장했던 Excel 12개를 사용자의 요청대로 `ExcelFileArchive` 최상위로 이동하고 빈 배치·날짜 폴더를 제거했습니다.
- Decisions: 파일 내용과 이름은 유지했으며 최상위에 기존 파일이 0개여서 이름 충돌 없이 이동했습니다. 네트워크 원본과 DB는 변경하지 않았습니다.
- Files: `D:\000. MyWorks\002. DB\InferenceDataAIService\universal-grid\ExcelFileArchive` 내부 배치 복사본 12개 위치, 빈 하위 폴더 2개, `HANDOFF.md`.
- Verification: 이동 전후 파일 수 12개와 총 154,547,312바이트가 일치합니다. 정리 후 최상위 파일 12개, 하위 파일 0개, 하위 폴더 0개입니다.
- Next: 이후 신규 수집도 코드 변경에 따라 같은 최상위 폴더에 직접 저장되며, 열린 WPF에서 DB 상태 버튼과 평면 저장 결과를 확인합니다.
## 2026-07-25 08:01 - 다중 Excel 검색 폴더 구현
- Completed: 단일 검색 버튼 옆에 다중 폴더 버튼과 모달 창을 추가하고, 저장 이력 불러오기·직접 입력·폴더 찾기·선택 삭제·전체 삭제·확인 시 전체 폴더 검색을 구현했다.
- Decisions: 확인된 폴더 목록 전체를 `%LocalAppData%\InferenceDataAIService\excel-search-folders.json`에 저장하며, 중복 루트와 겹치는 루트에서 발견된 동일 파일은 한 번만 결과에 포함한다. 확인은 검색·DB 비교까지만 수행하고 실제 복사는 기존 `DB 없음 모두 이 폴더로 복사` 동작으로 유지한다.
- Files: `InferenceDataAIService.Wpf/ExcelFolderHistoryStore.cs`, `InferenceDataAIService.Wpf/ExcelFolderSelectionWindow.xaml`, `InferenceDataAIService.Wpf/ExcelFolderSelectionWindow.xaml.cs`, `InferenceDataAIService.Wpf/ExcelFolderSearchService.cs`, `InferenceDataAIService.Wpf/MainWindow.xaml`, `InferenceDataAIService.Wpf/MainWindow.CanonicalEvidence.cs`, `.codex-verify/ExcelCollectionBehavior/Program.cs`, `HANDOFF.md`
- Verification: 구현 직후 `git diff --check` 통과; 컴파일 및 다중 루트 동작 검증은 다음 단계에서 수행한다.
- Next: WPF 좁은 빌드와 ExcelCollectionBehavior 검증을 실행하고 오류가 있으면 수정한다.

## 2026-07-25 08:03 - 다중 Excel 검색 검증
- Completed: WPF XAML·C# 컴파일 오류 2건(창 접근성, KeyEventArgs 모호성)을 수정하고 다중 검색 루트 동작 검증을 완료했다.
- Decisions: 기존 단일 검색 API는 다중 검색 API에 위임해 호환성을 유지하고, 같은 루트가 반복되거나 상위·하위 루트가 겹쳐도 동일 전체 경로 파일은 한 번만 표시한다.
- Files: `InferenceDataAIService.Wpf/ExcelFolderSelectionWindow.xaml.cs`, `.codex-verify/ExcelCollectionBehavior/Program.cs`, `HANDOFF.md`
- Verification: WPF 전용 출력 빌드가 경고 0·오류 0으로 통과했고, `ExcelCollectionBehavior`가 DB 원본 연결·다중 루트 결합·중복 루트 제거·평면 보관함 복사·재복사 중복 방지를 모두 PASS했다.
- Next: 정상 Debug 출력으로 빌드한 뒤 이전 WPF 창을 정상 종료하고 최신 앱을 실행해 창 응답 상태를 확인한다.

## 2026-07-25 08:04 - 다중 검색 폴더 WPF 실행
- Completed: 검증된 다중 검색 폴더 기능을 정상 Debug 산출물에 반영하고 최신 WPF를 실행했다.
- Decisions: 실행 중인 기존 Debug 프로세스가 없어 종료 작업 없이 새 실행 파일을 시작했다. 다중 폴더 팝업의 확인은 검색과 DB 비교를 수행하고, DB 없음 파일 복사는 메인 화면의 기존 복사 버튼에서 명시적으로 수행한다.
- Files: `InferenceDataAIService.Wpf/bin/Debug/net9.0-windows/*`, `HANDOFF.md`
- Verification: 정상 Debug 빌드가 경고 0·오류 0으로 통과했으며 PID 7328이 `HasExited=False`, `Responding=True`, 창 제목 `Inference Data AI` 상태다.
- Next: 열린 WPF의 `신규 Excel 수집`에서 `검색 폴더 여러 개`를 눌러 저장 이력 또는 새 폴더 목록을 확인하고 실제 네트워크 폴더 조합으로 결과를 검토한다.

## 2026-07-25 08:08 - 최신 WPF 다시 열기
- Completed: 사용자의 요청에 따라 최신 Debug WPF를 실행하고 창을 화면 앞으로 복원했다.
- Decisions: 소스 변경이나 재빌드는 하지 않고 기존 검증 완료 실행 파일을 사용했다.
- Files: `HANDOFF.md`; 소스 파일 변경 없음.
- Verification: PID 23624가 `HasExited=False`, `Responding=True`, 메인 창 핸들 2228922, 창 제목 `Inference Data AI` 상태다.
- Next: 사용자가 열린 `신규 Excel 수집` 화면에서 다중 검색 폴더 팝업을 확인한다.

## 2026-07-25 08:10 - 다중 폴더 팝업 암색 테마 수정
- Completed: 파생 Window에 암색 앱 배경이 적용되지 않아 흰색으로 표시되던 다중 검색 폴더 팝업을 암색 배경·테두리·사용자 지정 제목 표시줄로 교체했다.
- Decisions: 메인 WPF의 색상 리소스를 명시적으로 적용하고 Windows 기본 흰색 프레임 대신 메인 창과 일관된 암색 WindowChrome을 사용한다.
- Files: `InferenceDataAIService.Wpf/ExcelFolderSelectionWindow.xaml`, `InferenceDataAIService.Wpf/ExcelFolderSelectionWindow.xaml.cs`, `HANDOFF.md`
- Verification: 다이얼로그 테마 전용 WPF 빌드가 경고 0·오류 0으로 통과했다.
- Next: 실행 중인 이전 WPF를 정상 종료하고 정상 Debug 빌드 후 최신 창을 다시 열어 사용자에게 확인받는다.

## 2026-07-25 08:12 - 암색 팝업 반영 및 WPF 재실행
- Completed: 수정된 암색 다중 폴더 팝업을 정상 Debug 산출물에 반영하고 WPF를 다시 실행했다.
- Decisions: 팝업 기능과 폴더 이력 형식은 변경하지 않고 시각 테마만 메인 앱과 일치시켰다.
- Files: `InferenceDataAIService.Wpf/bin/Debug/net9.0-windows/*`, `HANDOFF.md`
- Verification: 정상 Debug 빌드가 경고 0·오류 0으로 통과했고 `git diff --check`도 통과했다. PID 8960이 `HasExited=False`, `Responding=True`, 창 제목 `Inference Data AI` 상태다.
- Next: 열린 WPF에서 `신규 Excel 수집` → `검색 폴더 여러 개`를 다시 눌러 암색 팝업을 확인한다.

## 2026-07-25 09:23 - WPF 열기
- Completed: 사용자의 요청에 따라 최신 Debug WPF를 실행하고 화면 앞으로 복원했다.
- Decisions: 추가 소스 변경이나 재빌드 없이 최신 검증 완료 실행 파일을 사용했다.
- Files: `HANDOFF.md`; 소스 변경 없음.
- Verification: PID 4992가 `HasExited=False`, `Responding=True`, 창 제목 `Inference Data AI` 상태다.
- Next: 사용자가 최신 암색 다중 폴더 팝업과 수집 흐름을 확인한다.

## 2026-07-25 09:28 - 신규 수집 화면에서 보관함 전체 처리 연결
- Completed: 신규 Excel 수집 화면에 `보관함 N개 전체 처리` 버튼을 추가하고, 클릭 시 별도 Excel 선택 없이 설정된 `ExcelFileArchive`를 기존 corpus 전체 처리 흐름으로 실행하도록 연결했다. 기존 DRM 처리 화면은 `전체 처리 상태·재시도`로 정리하고 처리 대상을 보관함 경로로 고정했다.
- Decisions: 보관함 최상위 Excel 전체를 처리 대상으로 사용한다. 기존 corpus journal 재개 동작을 유지해 완료된 동일 파일은 건너뛰고 새 파일이나 변경된 파일만 처리하며, 비용이 큰 COM·AI 실행 전에는 파일 수와 경로를 보여주는 확인 창을 유지한다.
- Files: `InferenceDataAIService.Wpf/MainWindow.xaml`, `InferenceDataAIService.Wpf/MainWindow.CanonicalEvidence.cs`, `InferenceDataAIService.Wpf/MainWindow.Settings.cs`, `HANDOFF.md`
- Verification: WPF 전용 출력 빌드가 경고 0·오류 0으로 통과했고 corpus 완료 파일 재실행 스킵 단위 테스트가 통과했다. 현재 설정 보관함 최상위에서 Excel 95개가 확인되어 버튼 활성화 조건을 충족한다.
- Next: 실행 중인 이전 WPF를 정상 종료하고 정상 Debug 빌드 후 최신 앱을 다시 열어 `신규 Excel 수집` 화면의 전체 처리 버튼을 확인한다.

## 2026-07-25 09:30 - 보관함 전체 처리 WPF 반영
- Completed: 보관함 전체 처리 변경을 정상 Debug 산출물로 빌드하고 최신 WPF를 실행했다.
- Decisions: 사용자의 지시에 따라 이후 WPF 검증은 빌드와 앱 실행까지만 수행하며, Codex가 버튼 클릭·화면 전환 등 UI 자동 조작 검증을 하지 않는다.
- Files: `InferenceDataAIService.Wpf/bin/Debug/net9.0-windows/*`, `HANDOFF.md`
- Verification: 정상 Debug 빌드가 경고 0·오류 0으로 통과했고 PID 8824에서 최신 WPF를 실행했다. UI 자동 조작 검증은 사용자 요청에 따라 중단했다.
- Next: 사용자가 열린 WPF에서 `신규 Excel 수집`의 `보관함 95개 전체 처리` 버튼과 자동 보관함 경로를 직접 확인한다.

## 2026-07-25 09:32 - 다중 폴더 선택 행 가독성 수정
- Completed: 다중 검색 폴더 팝업에서 선택된 행이 시스템 흰색 배경으로 표시되어 흰 글자가 안 보이던 문제를 보라색 암색 선택 배경과 흰 글자로 수정했다.
- Decisions: 활성·비활성 선택 색상을 모두 팝업 DataGrid 내부에서 명시해 Windows 시스템 테마와 포커스 상태에 관계없이 경로가 보이게 한다. 폴더 목록은 기존대로 `확인하고 검색` 시 저장하고 다음 팝업 실행 때 자동 복원한다.
- Files: `InferenceDataAIService.Wpf/ExcelFolderSelectionWindow.xaml`, `HANDOFF.md`
- Verification: 폴더 선택 색상 전용 WPF 빌드가 경고 0·오류 0으로 통과했다. 사용자 지시에 따라 UI 자동 클릭 검증은 수행하지 않았다.
- Next: 정상 Debug 빌드 후 최신 WPF만 다시 열고 사용자가 선택 행 색상과 이력 복원을 확인한다.

## 2026-07-25 09:33 - 선택 행 색상 WPF 반영
- Completed: 다중 폴더 선택 행 색상 수정본을 정상 Debug로 빌드하고 최신 WPF를 실행했다.
- Decisions: 사용자의 지시대로 앱 실행 뒤 UI 자동 클릭이나 화면 검증은 하지 않았다.
- Files: `InferenceDataAIService.Wpf/bin/Debug/net9.0-windows/*`, `HANDOFF.md`
- Verification: 정상 Debug 빌드가 경고 0·오류 0으로 통과했고 WPF 프로세스 PID 22412를 시작했다.
- Next: 사용자가 `검색 폴더 여러 개` 팝업에서 선택 행의 경로 가독성과 저장된 폴더 이력을 직접 확인한다.

## 2026-07-25 09:52 - Excel COM 사전 분석 단계 분리
- Completed: `신규 Excel 수집 → Excel COM 사전 분석 → 전체 처리` 흐름으로 메뉴를 개편하고, 보관함 Excel의 읽기 전용 COM 캡처·구조 지문 비교·기존/유사/신규/추출 실패 판정·결과 표를 구현했다. 전체 처리는 사전 분석에서 기존 양식으로 통과한 파일 manifest만 받으며 COM 캡처를 재사용한다.
- Decisions: 기존 분석 완료 양식과 유사도 0.82 이상만 `기존 양식`으로 자동 통과시키고, 0.60 이상은 `유사 양식 검토`, 그 미만은 `신규 양식`, COM 오류는 `추출 실패`로 AI 분석에서 보류한다. 현재 보관함 파일 수나 경로가 사전 분석 결과와 다르면 전체 처리를 차단하며, manifest SHA-256 검증으로 사전 분석 뒤 변경된 파일도 거부한다.
- Files: `inference_data_ai_form_preflight.py`, `inference_data_ai_cli.py`, `tests/test_inference_data_ai_form_preflight.py`, `tests/test_inference_data_ai_cli.py`, `InferenceDataAIService.Wpf/AppPathSettings.cs`, `InferenceDataAIService.Wpf/CanonicalEvidenceClient.cs`, `InferenceDataAIService.Wpf/MainWindow.xaml`, `InferenceDataAIService.Wpf/MainWindow.xaml.cs`, `InferenceDataAIService.Wpf/MainWindow.CanonicalEvidence.cs`, `HANDOFF.md`
- Verification: 사전 판정 및 CLI parser 단위 테스트 4건이 통과했고 WPF 프로젝트 `dotnet build --no-restore`가 경고 0·오류 0으로 통과했다. 사용자 지시에 따라 UI 자동 클릭·화면 전환 검증은 수행하지 않았다.
- Next: 최신 WPF를 실행만 한 뒤 사용자가 새 `Excel COM 사전 분석` 메뉴의 구성과 문구를 직접 확인한다.

## 2026-07-25 09:53 - COM 사전 분석 WPF 실행
- Completed: 이전 동일 WPF 프로세스를 종료하고 COM 사전 분석 메뉴가 포함된 최신 Debug WPF를 실행했다.
- Decisions: 사용자가 직접 화면을 확인하도록 앱 실행까지만 수행하고 UI 자동 클릭이나 화면 전환은 하지 않았다.
- Files: `InferenceDataAIService.Wpf/bin/Debug/net9.0-windows/*`, `HANDOFF.md`
- Verification: 최신 WPF 프로세스가 PID 20144로 시작됐다. 자동 UI 검증은 사용자 요청에 따라 의도적으로 생략했다.
- Next: 사용자가 열린 WPF에서 `신규 Excel 수집 → Excel COM 사전 분석 → 전체 처리 상태·재시도` 순서와 판정 결과 표를 직접 확인한다.

## 2026-07-25 09:55 - 사전 분석 전 Excel 목록 표시
- Completed: 사전 분석 결과가 아직 없어도 보관함의 모든 Excel을 결과 표에 즉시 `분석 전` 상태로 표시하고, COM 실행 중인 현재 파일은 `COM 분석 중`으로 바뀌며 해당 행이 보이도록 수정했다. 수정본 WPF를 PID 24424로 다시 실행했다.
- Decisions: 유사도는 판정 전·진행 중·추출 실패 상태에서 `-`로 표시해 실제 판정값 0%와 혼동하지 않게 한다.
- Files: `InferenceDataAIService.Wpf/CanonicalEvidenceClient.cs`, `InferenceDataAIService.Wpf/MainWindow.CanonicalEvidence.cs`, `HANDOFF.md`
- Verification: WPF 프로젝트 좁은 빌드가 경고 0·오류 0으로 통과했고 최신 앱 프로세스가 시작됐다. UI 자동 조작 검증은 사용자 요청에 따라 수행하지 않았다.
- Next: 사용자가 사전 분석 화면에서 보관함 Excel 387개가 `분석 전` 상태로 표시되는지 확인한다.

## 2026-07-25 12:13 - 최신 WPF 재실행
- Completed: 실행 중이던 동일 WPF 프로세스를 종료하고 최신 Debug WPF를 PID 9864로 다시 열었다.
- Decisions: 중복 창을 남기지 않고 최신 빌드 하나만 실행했으며 UI 자동 조작은 하지 않았다.
- Files: `HANDOFF.md`
- Verification: WPF 프로세스 PID 9864 시작을 확인했다.
- Next: 사용자가 열린 WPF 화면을 직접 확인한다.

## 2026-07-25 12:15 - COM 사전 분석 출력 경로 오류 수정
- Completed: WPF 설정 출력 루트가 서비스 소스 폴더 밖에 있을 때 `Output path must stay under` 오류로 사전 분석이 중단되던 문제를 수정했다. WPF가 설정된 출력 루트를 CLI에 명시하고 CLI는 결과 경로가 그 루트 내부인지 검증한다. 수정본 WPF를 PID 24560으로 다시 실행했다.
- Decisions: 임의 경로 저장을 허용하지 않고, 명시된 설정 출력 루트 내부만 허용하는 경계 검증을 유지한다.
- Files: `inference_data_ai_cli.py`, `InferenceDataAIService.Wpf/CanonicalEvidenceClient.cs`, `tests/test_inference_data_ai_cli.py`, `HANDOFF.md`
- Verification: 출력 루트 parser·내부 허용·외부 이탈 차단과 양식 판정 단위 테스트 5건이 통과했고 WPF 빌드가 경고 0·오류 0으로 통과했다.
- Next: 사용자가 `COM 추출·양식 판정 시작`을 직접 재시도한다.

## 2026-07-25 13:21 - 파일별 Excel COM 장애 격리
- Completed: 사전 분석이 53개 완료 후 54번째 파일에서 Python 작업이 종료되고 자동화 Excel PID 2008이 남은 상태를 확인했다. 각 Excel을 별도 Python·Excel 프로세스에서 추출하도록 격리하고, 파일당 300초 제한 시간·전용 Excel PID 정리·파일별 실패 후 계속 진행·판정 즉시 WPF 행 갱신을 구현했다. 실패가 남긴 전용 자동화 Excel PID 2008도 정리하고 최신 WPF를 PID 11980으로 실행했다.
- Decisions: 한 Excel의 COM/네이티브 충돌이나 무응답은 해당 파일만 `COM 추출 실패`로 보류하고 전체 387개 사전 분석과 WPF는 계속 진행한다. 원본은 읽기 전용이며 기존 고정 좌표·병합 셀 보존 추출 계약은 변경하지 않는다.
- Files: `inference_data_ai_com_capture.py`, `inference_data_ai_com_worker.py`, `inference_data_ai_form_preflight.py`, `inference_data_ai_cli.py`, `InferenceDataAIService.Wpf/MainWindow.CanonicalEvidence.cs`, `tests/test_inference_data_ai_form_preflight.py`, `HANDOFF.md`
- Verification: Python 구문 검사와 COM 격리·제한 시간·병합 셀 계약·출력 경로 관련 단위 테스트 12건이 통과했고 WPF 빌드가 경고 0·오류 0으로 통과했다. 실제 Excel COM 자동 실행과 UI 자동 조작은 사용자 요청에 따라 수행하지 않았다.
- Next: 사용자가 사전 분석을 다시 시작한다. 이전 성공 캡처는 DB에서 재사용되며 실패 지점부터 안전하게 계속된다.

## 2026-07-25 13:23 - 사전 분석 JSON 잠금 경합 수정
- Completed: WPF가 부분 결과 `latest.json`을 읽는 순간 Python의 원자 교체가 Windows `PermissionError(WinError 5)`로 실패하는 경합을 수정했다. WPF 읽기 스트림에 삭제 공유를 허용하고 Python은 실행별 고유 임시 파일과 지수형 교체 재시도를 사용한다. 최신 WPF를 PID 18332로 실행했다.
- Decisions: 부분 결과의 원자성은 유지하며, 일시적인 백신·파일 판독 잠금은 최대 약 5초 재시도한 뒤에만 실제 오류로 처리한다.
- Files: `inference_data_ai_form_preflight.py`, `inference_data_ai_com_worker.py`, `InferenceDataAIService.Wpf/CanonicalEvidenceClient.cs`, `tests/test_inference_data_ai_form_preflight.py`, `HANDOFF.md`
- Verification: Python 구문 검사와 잠금 재시도·COM 격리·추출 계약 관련 단위 테스트 13건이 통과했고 WPF 빌드가 경고 0·오류 0으로 통과했다. 실제 COM/UI 자동 실행은 수행하지 않았다.
- Next: 사용자가 사전 분석을 다시 시작한다.

## 2026-07-25 13:26 - 최신 WPF 재실행
- Completed: 실행 중인 동일 WPF를 종료하고 최신 수정본을 PID 3564로 다시 열었다.
- Decisions: 중복 창 없이 최신 빌드 하나만 실행하고 UI 자동 조작은 하지 않았다.
- Files: `HANDOFF.md`
- Verification: WPF 프로세스 PID 3564 시작을 확인했다.
- Next: 사용자가 열린 WPF에서 사전 분석을 직접 실행한다.

## 2026-07-25 13:28 - NumberFormat 격자 오류와 진행 중 스크롤 수정
- Completed: Excel COM이 다중 셀 `NumberFormat`을 단일 값이나 불완전 배열로 반환할 때 고정 격자보다 짧아 `IndexError`가 나던 문제를 수정했다. 스칼라는 전체 좌표에 적용하고 불완전 배열은 좌표를 유지한 채 빈 자리로 보정한다. worker 실패 문구는 전체 Traceback 대신 마지막 핵심 오류만 표시하며, 진행 행 자동 스크롤을 제거해 분석 중 사용자가 표를 자유롭게 스크롤할 수 있게 했다. 최신 WPF를 PID 14268로 실행했다.
- Decisions: 값·수식·서식 배열은 항상 UsedRange의 행×열 크기를 유지하고 병합 셀 placeholder를 제거하지 않는다. 진행 상태 갱신은 행 내용만 변경하며 사용자 스크롤 위치를 강제로 이동하지 않는다.
- Files: `inference_data_ai_com_capture.py`, `inference_data_ai_form_preflight.py`, `InferenceDataAIService.Wpf/MainWindow.CanonicalEvidence.cs`, `tests/test_inference_data_ai_com_capture.py`, `tests/test_inference_data_ai_form_preflight.py`, `HANDOFF.md`
- Verification: Python 구문 검사와 NumberFormat 스칼라·불완전 배열·worker 오류 요약·COM 격리 관련 단위 테스트 12건이 통과했고 WPF 빌드가 경고 0·오류 0으로 통과했다. 실제 COM/UI 자동 실행은 수행하지 않았다.
- Next: 사용자가 사전 분석 재개와 진행 중 수동 스크롤을 확인한다.

## 2026-07-25 13:31 - 분석 중 판정표 입력 잠금 해제
- Completed: 공통 busy 상태가 읽기 전용 판정표까지 `IsEnabled=false`로 설정해 스크롤과 행 선택을 막던 원인을 제거했다. 분석 중에도 표는 활성 상태로 유지하며 실행 버튼만 잠근다. 남아 있던 전용 자동화 Excel PID 16168을 정리하고 최신 WPF를 PID 18100으로 실행했다.
- Decisions: 판정표는 원래 읽기 전용이므로 COM 작업 중 스크롤·행 선택·내용 확인을 허용해도 처리 상태를 변경할 위험이 없다.
- Files: `InferenceDataAIService.Wpf/MainWindow.CanonicalEvidence.cs`, `HANDOFF.md`
- Verification: WPF 프로젝트 빌드가 경고 0·오류 0으로 통과했고 최신 프로세스 시작을 확인했다. UI 자동 조작은 수행하지 않았다.
- Next: 사용자가 사전 분석 중 판정표 스크롤과 행 선택이 가능한지 확인한다.

## 2026-07-25 16:24 - Excel COM 사전 분석 중지 기능
- Completed: 사전 분석 화면에 `추출 중지` 버튼을 추가했다. 중지 요청 시 현재 전용 Python worker와 그 worker가 만든 Excel COM 프로세스만 종료하고, 완료된 파일의 캡처·판정은 `CANCELLED` 부분 결과로 저장한다. 다시 실행하면 DB의 성공 캡처를 재사용한다. 취소 종료 시 SQLite 연결도 즉시 닫도록 보완하고 최신 WPF를 PID 27120으로 실행했다.
- Decisions: 일반 사용자의 Excel 프로세스는 종료하지 않으며, 중지 버튼은 COM 사전 분석 실행 중에만 활성화한다. 중지된 부분 결과는 전체 처리에는 사용할 수 없고 완료된 사전 분석만 기존 양식 manifest를 사용할 수 있다.
- Files: `inference_data_ai_form_preflight.py`, `inference_data_ai_cli.py`, `InferenceDataAIService.Wpf/CanonicalEvidenceClient.cs`, `InferenceDataAIService.Wpf/MainWindow.xaml`, `InferenceDataAIService.Wpf/MainWindow.CanonicalEvidence.cs`, `tests/test_inference_data_ai_form_preflight.py`, `tests/test_inference_data_ai_cli.py`, `HANDOFF.md`
- Verification: 취소 marker·현재 worker 종료·부분 결과·CLI 인자 관련 단위 테스트 10건이 통과했고 WPF 빌드가 경고 0·오류 0으로 통과했다. 실제 COM/UI 자동 실행은 수행하지 않았다.
- Next: 사용자가 분석 중 `추출 중지` 버튼과 재시작 시 캡처 재사용을 확인한다.

## 2026-07-25 18:48 - 구조 상이 양식 후속 처리 검토
- Completed: 최신 사전 분석 결과를 점검해 정상 완료가 아니라 `RUNNING` 상태에서 236/387개 처리 후 프로세스가 사라진 것을 확인했다. 처리분은 기존 양식 0, 유사 검토 111, 신규 보류 110, COM 실패 15개다.
- Decisions: 구조 상이 파일을 221개 개별 AI 분석하지 않고 안정적인 레이아웃 특징으로 양식군을 묶은 뒤, 군별 대표 Excel 1개를 AI로 분석하고 2~3개 표본 검증·사람 승인 후 해당 군 전체에 추출 계약을 적용하는 `양식 검토·등록` 단계가 필요하다. 기존 양식 연결·신규 양식 등록·제외의 3개 판정 경로를 제공해야 한다.
- Files: `HANDOFF.md`
- Verification: `form-preflight/latest.json` 상태·요약·대표 유사도와 실행 프로세스를 읽기 전용으로 확인했으며 현재 Python/Excel 자동화 프로세스가 없고 WPF PID 27120만 실행 중이다.
- Next: 사용자 승인 시 `양식군 검토·등록` 메뉴, 양식군 클러스터링 DB, 대표본 AI 추출 계약 생성·표본 검증·승인 후 재판정 흐름을 구현한다.

## 2026-07-25 18:59 - 중단 지점 복원
- Completed: 마지막 인수인계 이후 작성된 미기록 코드를 대조해 `양식군 검토·등록` 백엔드 초안 작성 도중 작업이 중단된 것을 확인했다. 새 모듈에는 구조 기반 양식군 생성, SQLite 등록부, 대표본 AI 추출 계약, 최대 3개 표본 검증, 신규 등록·기존 연결·제외 판정이 구현돼 있다.
- Decisions: `inference_data_ai_form_registry.py`는 보존하고 이어서 완성한다. 현재는 CLI·WPF·테스트 어디에도 연결되지 않은 독립 초안이므로 사용자 기능으로는 아직 사용할 수 없다.
- Files: 조사 중 소스 변경 없음; `HANDOFF.md`만 갱신.
- Verification: 새 모듈의 수정 시각과 참조 관계를 확인했고 `python -m py_compile inference_data_ai_form_registry.py`가 통과했다. 현재 Python·Excel·WPF 관련 프로세스는 없으며, 기존 사전 분석은 마지막 기록대로 236/387 `RUNNING` 잔여 상태다.
- Next: `inference_data_ai_form_registry.py` 단위 테스트를 먼저 추가한 뒤 CLI 명령에 연결하고, 마지막으로 WPF `양식군 검토·등록` 화면과 승인 후 사전 분석 재판정 흐름을 구현한다.

## 2026-07-25 19:02 - 양식군 기능 연결 구조 점검
- Completed: 중단된 등록 모듈과 COM 사전 분석 보고서 생성, CLI parser/handler, WPF 사전 분석 탭·클라이언트의 실제 연결 지점을 모두 확인했다.
- Decisions: 등록 결정은 DB에만 저장하지 않고 현재 사전 분석 보고서와 통과 manifest를 즉시 재생성한다. 사전 분석 재실행 때도 저장된 양식군 결정을 자동 적용하며, WPF에는 별도 `양식군 검토·등록` 탭을 제공한다.
- Files: 조사 중 소스 변경 없음; `HANDOFF.md`만 갱신.
- Verification: `form_registry`가 CLI·WPF·테스트에서 참조되지 않는 상태와 기존 corpus가 사전 분석 manifest의 파일 경로·SHA-256을 검증하는 경계를 확인했다.
- Next: 백엔드 재판정 함수와 등록 결정 테스트를 구현하고 통과시킨다.

## 2026-07-25 19:05 - 양식군 등록 백엔드 완성
- Completed: 사전 분석 재실행 시 저장된 양식군 결정을 자동 적용하고, 대표본 AI 계약·표본 호환성 검증·사람 승인/제외 후 현재 보고서와 전체 처리 manifest를 즉시 재판정하도록 구현했다. 승인된 신규 양식의 추출 계약도 manifest에 포함한다.
- Decisions: 구조 성장에 안정적인 양식군 ID를 사용하고, 신규 등록은 현재 양식군의 모든 선정 표본이 `PASSED`여야 허용한다. 사람 판정에는 reviewer를 필수로 하며 기존 연결 대상은 실제 분석 완료 양식 서명으로 제한한다.
- Files: `inference_data_ai_form_registry.py`, `inference_data_ai_form_preflight.py`, `tests/test_inference_data_ai_form_registry.py`, `HANDOFF.md`
- Verification: 구문 검사 후 `python -m unittest tests.test_inference_data_ai_form_registry tests.test_inference_data_ai_form_preflight`가 12/12 통과했다. 테스트 중 발견된 Windows SQLite 파일 잠금은 모든 등록부 연결을 명시적으로 닫도록 수정한 뒤 재검증했다.
- Next: 양식군 목록 생성·AI 분석·사람 판정 CLI 명령과 parser 테스트를 추가한다.

## 2026-07-25 19:07 - 양식군 CLI 연결
- Completed: `form-group-review`, `form-family-analyze`, `form-family-decide` 명령을 추가했다. 사람 판정 명령은 DB 등록 후 현재 사전 분석 보고서·manifest·양식군 목록을 한 번에 재생성한다.
- Decisions: WPF 설정 출력 루트 밖으로 결과가 이탈하지 못하도록 기존 경계 검증을 재사용하고, AI 계약은 양식군별 파일로 분리한다.
- Files: `inference_data_ai_cli.py`, `tests/test_inference_data_ai_cli.py`, `HANDOFF.md`
- Verification: CLI 구문 검사와 새 parser/판정 재생성 호출 및 등록 백엔드 테스트 6건이 통과했다.
- Next: WPF 클라이언트와 별도 `양식군 검토·등록` 탭을 연결한다.

## 2026-07-25 19:12 - WPF 양식군 검토·등록 연결
- Completed: 왼쪽 메뉴와 별도 탭에 양식군 목록, 대표본 AI 분석, 판정자·표시 이름·기존 서명·메모 입력, 신규 등록·기존 연결·제외 버튼을 추가했다. 판정 후 사전 분석 표와 전체 처리 가능 수가 즉시 갱신된다.
- Decisions: AI는 제안과 표본 호환성 검증까지만 수행하고 최종 판정은 확인 대화상자를 거친 사람 조작으로 제한한다. 부분 `RUNNING` 보고서도 검토할 수 있지만 전체 처리는 기존대로 `COMPLETED` 보고서만 허용한다.
- Files: `InferenceDataAIService.Wpf/CanonicalEvidenceClient.cs`, `InferenceDataAIService.Wpf/MainWindow.xaml`, `InferenceDataAIService.Wpf/MainWindow.xaml.cs`, `InferenceDataAIService.Wpf/MainWindow.CanonicalEvidence.cs`, `HANDOFF.md`
- Verification: 기존 WPF PID 27120이 정상 Debug EXE를 사용 중이라 첫 빌드는 산출물 복사 단계에서만 실패했다. 사용자 창을 종료하지 않고 `.codex-verify/wpf-form-registry` 별도 출력으로 같은 WPF 프로젝트를 빌드해 경고 0·오류 0을 확인했다. 앱 실행이나 UI 자동 조작은 수행하지 않았다.
- Next: 관련 Python 테스트 전체와 변경 파일 무결성 검사를 실행하고 최종 인수인계를 기록한다.

## 2026-07-25 19:13 - 양식군 검토·등록 최종 검증
- Completed: 중단됐던 양식군 검토·등록 기능을 백엔드, CLI, WPF, 승인 후 사전 분석 재판정까지 완성했다.
- Decisions: 현재 실행 중인 PID 27120은 사용자가 확인 중인 이전 빌드이므로 종료하거나 새 앱을 실행하지 않았다. 새 화면은 다음 정상 빌드·실행부터 표시된다.
- Files: `inference_data_ai_form_registry.py`, `inference_data_ai_form_preflight.py`, `inference_data_ai_cli.py`, `tests/test_inference_data_ai_form_registry.py`, `tests/test_inference_data_ai_cli.py`, `InferenceDataAIService.Wpf/CanonicalEvidenceClient.cs`, `InferenceDataAIService.Wpf/MainWindow.xaml`, `InferenceDataAIService.Wpf/MainWindow.xaml.cs`, `InferenceDataAIService.Wpf/MainWindow.CanonicalEvidence.cs`, `HANDOFF.md`
- Verification: 관련 Python 모듈 테스트 38/38 통과, 새 CLI 3개 도움말 통과, 변경 파일 공백 검사 통과, WPF 별도 출력 빌드 경고 0·오류 0 통과. 실제 Codex AI 호출, Excel COM 실행, UI 자동 조작, 앱 재실행은 수행하지 않았다.
- Next: 사용자가 기존 PID 27120을 닫은 뒤 정상 Debug 빌드·실행을 요청하면 새 `양식군 검토·등록` 화면에서 기존 236개 부분 사전 분석 결과를 불러오고, 대표 양식군부터 AI 분석·사람 판정을 진행한다.

## 2026-07-26 06:07 - 최신 WPF 실행
- Completed: 사용자의 명시적 요청에 따라 `양식군 검토·등록` 변경을 정상 Debug 산출물로 빌드하고 최신 WPF를 열었다.
- Decisions: 기존 WPF 프로세스가 없음을 확인해 종료 작업 없이 최신 실행 파일 하나만 시작했다.
- Files: `InferenceDataAIService.Wpf/bin/Debug/net9.0-windows/*`, `HANDOFF.md`; 소스 변경 없음.
- Verification: WPF 프로젝트 정상 Debug 빌드가 경고 0·오류 0으로 통과했다. PID 27536이 `HasExited=False`, `Responding=True`, 창 제목 `Inference Data AI` 상태다.
- Next: 열린 WPF의 왼쪽 메뉴에서 `양식군 검토·등록`을 선택해 기존 부분 사전 분석 결과를 불러오고 대표 양식군 검토를 시작한다.

## 2026-07-26 06:13 - CLI 전체 처리 실제 상태 점검
- Completed: 설정 파일에서 실제 출력 루트를 복원하고 보관함·DB·사전 분석·corpus journal 상태를 확인했다. 보관함은 387개인데 완료 사전 분석은 386개이며 기존 2, 유사 161, 신규 201, COM 실패 22개다. 검토 대상 362개는 현재 267개 양식군으로 묶인다.
- Decisions: 수백 개 양식군을 수동 클릭하지 않고 사전 분석 재개, 체크포인트 가능한 양식군 AI 분석·자동 판정, manifest 재생성, corpus 실패 재시도를 한 CLI 명령으로 연결한다. 원본 Excel은 계속 읽기 전용으로 유지한다.
- Files: `HANDOFF.md`; 조사 중 소스 변경 없음. 실제 `form-group-review` 결과 `group-review.latest.json`을 설정 출력 루트에 갱신.
- Verification: 실제 DB 3.60GB와 경로 존재, 현재 Python·Excel 자동화 프로세스 없음, WPF PID 27536만 응답 중임을 확인했다. 기존 corpus journal은 완료 5·실패 27의 `COMPLETED_WITH_ERRORS` 상태다.
- Next: 재시작 가능한 `form-pipeline-complete` CLI와 자동 판정 정책·테스트를 구현한다.

## 2026-07-26 06:17 - CLI 전체 자동 파이프라인 구현
- Completed: `form-pipeline-complete` 명령을 추가해 보관함 COM 사전 분석 재개, 양식군 병렬 AI 분석, 결과별 신규 등록·기존 연결·제외 자동 판정, 체크포인트, manifest 재생성, 승인된 corpus 전체 재시도를 한 흐름으로 연결했다.
- Decisions: 표본 검증 실패나 AI `EXCLUDE` 제안은 제외하고, 검증 통과 `LINK_EXISTING`은 실제 기존 서명이 있을 때만 연결하며 나머지 검증 통과군은 신규 등록한다. AI 호출 실패 자동 제외는 명시 옵션으로만 허용하고 원본·캡처는 보존한다.
- Files: `inference_data_ai_form_pipeline.py`, `inference_data_ai_form_registry.py`, `inference_data_ai_cli.py`, `tests/test_inference_data_ai_form_pipeline.py`, `tests/test_inference_data_ai_cli.py`, `HANDOFF.md`
- Verification: 새 모듈 구문 검사와 자동 판정·병렬 처리·CLI parser 관련 테스트 7/7 통과, `form-pipeline-complete --help` 통과, Codex CLI 0.145.0 가용 확인.
- Next: 실제 데이터에서 한 양식군 제한 실행으로 COM 재개·Codex 호출·DB 판정 저장을 검증한 뒤 전체 제한을 해제한다.

## 2026-07-26 06:48 - 실데이터 사전 분석 복구 및 AI 스키마 보정
- Completed: 변경되지 않은 과거 `CAPTURE_FAILED` 22건을 기본적으로 재사용하도록 사전 분석을 보강해 386개 전체 보고서를 26초 내 복구했고, 첫 양식군 실호출에서 확인된 Codex 구조화 출력 스키마 오류를 수정했다. 중단된 COM 실행이 남긴 숨김 Excel 자동화 인스턴스 18개도 실행 시각과 `/automation -Embedding` 조건으로 한정해 종료했다.
- Decisions: 실패 캡처는 `REUSED_FAILED`로 보존하고 명시적 `--retry-failed-captures`에서만 다시 COM 캡처한다. Codex 구조화 출력의 `const`와 `enum` 속성에도 명시적인 문자열 `type`을 둔다.
- Files: `inference_data_ai_form_preflight.py`, `inference_data_ai_form_registry.py`, `inference_data_ai_cli.py`, `tests/test_inference_data_ai_form_preflight.py`, `tests/test_inference_data_ai_form_registry.py`, `HANDOFF.md`
- Verification: 실제 사전 분석 결과가 전체 386·기존 2·유사 161·신규 201·캡처 실패 22로 복구됐고 관련 단위 테스트 15/15가 통과했다. 첫 실호출은 수정 전 스키마의 `schemaVersion`에 `type`이 없다는 HTTP 400을 재현해 원인을 확정했다.
- Next: 수정된 스키마로 양식군 1개 실호출을 다시 수행해 AI 분석·자동 판정·DB 저장까지 검증한 뒤 전체 양식군과 corpus를 연속 실행한다.

## 2026-07-26 06:53 - 실데이터 양식군 1개 종단 검증
- Completed: 구 실행의 잔존 Python·COM 작업을 PID 계보로 제거하고 386개 사전 분석 보고서를 복구한 뒤, `family-42caacbd713e384d`의 Codex 계약 생성·3개 표본 호환성 검증·자동 `REGISTER_NEW` 판정·DB 승인을 끝까지 완료했다.
- Decisions: 경쟁 실행이 최신 보고서를 덮어써 발생한 표본 수 불일치는 원본이나 AI 오류가 아니며, 구 실행 제거 후 저장된 검증 계약을 재사용해 판정했다. 실검증으로 확인된 자동 파이프라인을 잔여 266개 양식군에 그대로 적용한다.
- Files: `D:\000. MyWorks\002. DB\InferenceDataAIService\form-preflight\latest.json`, `group-review.latest.json`, `latest.known-forms.manifest.json`, `contracts\family-42caacbd713e384d.json`, `pipeline.checkpoint.json`, `pipeline.result.json`, `HANDOFF.md`
- Verification: 결과는 오류 0, `APPROVED_NEW` 1개 양식군·14개 파일이며 사전 분석 요약은 전체 386·승인 양식 16·유사 157·신규 191·캡처 실패 22다. 실패 22건은 모두 `REUSED_FAILED`, 나머지 364건은 `REUSED_CAPTURE`임을 확인했고 남은 Python·Excel 자동화 프로세스가 없음을 확인했다.
- Next: `form-pipeline-complete`를 제한 없이 실행해 잔여 266개 양식군을 자동 판정하고 pending 0이 되면 corpus 재시도를 이어서 완료한다.

## 2026-07-26 07:01 - 양식군 배치 처리 CPU 병목 제거
- Completed: 전체 실행 초기에 각 AI 분석과 판정이 동일한 267개 양식군 검토를 매번 다시 구성하는 CPU 병목을 확인하고, 파이프라인이 최초 계산한 양식군 스냅샷을 분석·판정 함수에 전달해 재사용하도록 최적화했다.
- Decisions: 일반 CLI 단일 분석·판정 호출은 기존처럼 최신 보고서에서 양식군을 재구성하고, 한 실행 안에서 이미 검증된 파이프라인 경로만 스냅샷을 사용한다. 따라서 외부 호출의 최신성 검증은 유지하면서 배치 중복 계산만 제거한다.
- Files: `inference_data_ai_form_registry.py`, `inference_data_ai_form_pipeline.py`, `tests/test_inference_data_ai_form_pipeline.py`, `HANDOFF.md`
- Verification: 관련 registry·pipeline 단위 테스트 5/5가 통과했고, 중단 전 두 실행에서 누적 저장된 승인 결정은 DB 체크포인트에 보존됐다. 해당 실행 계보의 Python·Node·Codex 자식 프로세스는 모두 종료됐다.
- Next: 최적화된 코드로 승인 완료 양식군을 건너뛰고 잔여 양식군 전체 실행을 재개해 처리율과 오류율을 확인한다.

## 2026-07-26 07:22 - 전체 양식군 판정 완료 및 공백 파일명 보정
- Completed: 최적화된 전체 실행에서 잔여 249개 양식군을 오류 없이 모두 판정했다. corpus 진입 시 실제 파일명 ` (2).xlsx`의 선행 공백을 manifest 로더가 제거해 `(2).xlsx`로 조회하는 문제를 확인하고, 파일명은 원문 그대로 보존하면서 공백-only 값만 거부하도록 수정했다.
- Decisions: 양식 계약 결과는 누적 244개 신규 승인·5개 기존 연결이며 이전 실행의 승인까지 DB에 보존돼 전체 pending은 해소됐다. 파일 경로 보안 검증과 SHA-256 검증은 유지하고 선행·후행 공백도 합법적인 Windows 파일명으로 취급한다.
- Files: `inference_data_ai_form_pipeline.py`, `tests/test_inference_data_ai_form_pipeline.py`, `D:\000. MyWorks\002. DB\InferenceDataAIService\form-preflight\contracts\*.json`, DB 양식 registry, `HANDOFF.md`
- Verification: 해당 실행의 249/249 판정 완료·오류 0을 체크포인트와 로그에서 확인했다. manifest 361개 경로는 공백을 보존하면 실제 누락 0이며, 관련 pipeline·registry 테스트 6/6가 통과했다.
- Next: 전체 파이프라인을 재개하면 이미 판정된 양식군은 건너뛰고 corpus 단계부터 실행된다. corpus journal 완료와 최종 DB 상태를 끝까지 확인한다.

## 2026-07-26 07:24 - 승인 corpus 산출물 격리
- Completed: corpus 시작 전 기존 journal의 소유 경로 검증 실패를 확인했다. 해당 32건 journal은 현재 DB가 아니라 repository 내부 과거 sourceRoot·databasePath·artifactRoot에 속하므로 그대로 보존하고, 현재 승인 corpus의 artifactRoot를 `incremental-com-corpus\form-approved` 하위로 분리했다.
- Decisions: 서로 다른 원본·DB의 journal이나 중간 산출물을 재사용하지 않는다. 현재 361개 manifest는 전용 하위 디렉터리와 새 `corpus-journal.json`에서 시작하며 이후 같은 작업의 재시작에만 재사용한다.
- Files: `inference_data_ai_form_pipeline.py`, `HANDOFF.md`; 기존 `incremental-com-corpus\corpus-journal.json`은 변경하지 않음.
- Verification: pipeline 단위 테스트 3/3 통과. 과거 중단된 COM 작업의 `/automation -Embedding` Excel 잔존 인스턴스 16개를 한정 종료했고 잔존 0을 확인했다.
- Next: 전체 파이프라인을 재개해 `form-approved\corpus-journal.json` 생성과 361개 처리 진행 상태를 확인한다.

## 2026-07-26 07:38 - 캡처 전용 원본의 의미 분석 재개 보정
- Completed: corpus가 canonical 분석이 없는 DB 원본 캡처도 `COMPLETED`로 조정하던 결함을 수정했다. 이제 공개 canonical 분석이 있는 원본만 완료로 재사용하며, 기존 journal의 `existing-database-source-v1`/`sourceOnlyDuplicate=true` 완료 기록은 자동으로 `PENDING`으로 내려 PACKET→AI→IMPORT→VERIFY 처리를 재개한다.
- Decisions: DB에 source revision이 존재하는 것은 캡처 완료일 뿐 의미 분석 완료로 취급하지 않는다. canonical public analysis ID가 있고 answer-visible verification 상태인 경우에만 corpus 완료를 재조정한다.
- Files: `inference_data_ai_corpus_workflow.py`, `tests/test_inference_data_ai_corpus_workflow.py`, `HANDOFF.md`
- Verification: `python -m unittest tests.test_inference_data_ai_corpus_workflow` 12/12 통과.
- Next: 실제 361개 승인 corpus 파이프라인을 CLI로 재개하고 journal·DB에서 의미 분석 완료 또는 명시적 terminal 상태를 끝까지 확인한다.

## 2026-07-26 07:55 - 전체 분석 방식 WPF 통합
- Completed: WPF의 기존 수동 preflight/양식군 검토 기능은 유지하면서 `form-pipeline-complete`를 직접 호출하는 전체 자동 분석 경로를 추가했다. 한 버튼으로 COM 사전 분석, 신규·유사 양식군 AI 계약 검증·자동 판정, 승인 corpus의 PACKET→AI→DB import→verify, 최신 적재 DB AI 문의 반영까지 진행·요약한다.
- Decisions: 이전의 “기존 양식만 별도 corpus 처리” 버튼을 현재 전체 자동 CLI 파이프라인 진입점으로 바꿨다. 원본과 실패 journal은 보존하며 완료 여부는 pending 양식군, 분석 오류, corpus 실패가 모두 0일 때만 표시한다.
- Files: `InferenceDataAIService.Wpf/CanonicalEvidenceClient.cs`, `InferenceDataAIService.Wpf/MainWindow.CanonicalEvidence.cs`, `InferenceDataAIService.Wpf/MainWindow.xaml`, `HANDOFF.md`
- Verification: `dotnet build InferenceDataAIService.Wpf/InferenceDataAIService.Wpf.csproj --no-restore` 성공, 경고 0·오류 0. 실행 중인 실제 CLI와 충돌하지 않도록 WPF 앱 자체는 시작하지 않음.
- Next: 실행 중인 실제 전체 CLI를 완료하고 canonical DB·AI 질의 포함 여부를 검증한 뒤 필요하면 WPF 결과 표시를 최종 상태에 맞게 조정한다.

## 2026-07-26 08:08 - 통합 변경 회귀 검증
- Completed: 캡처 전용 완료 재큐잉, source count 보정, COM 병합 수식 집계, 양식 전체 파이프라인과 CLI 변경을 함께 묶어 관련 Python 회귀 테스트를 실행했다.
- Decisions: 실제 장기 corpus는 별도 실제 DB에서 계속 실행하고, 회귀 테스트는 임시 fixture DB와 가장 좁은 관련 모듈 집합만 사용했다.
- Files: `HANDOFF.md`; 이 단계의 소스 변경 없음.
- Verification: corpus workflow·workbook workflow·source ingest·COM capture·form pipeline/registry/preflight·CLI 테스트 83/83 통과.
- Next: corpus run 6의 361개 의미 분석을 끝까지 감시하고 실패 항목은 보존된 rejected draft로 retry해 0건이 될 때까지 재개한다.

## 2026-07-26 08:12 - corpus rejected draft 자동 재시도
- Completed: 한 번의 전체 CLI/WPF 실행 안에서 첫 AI draft 검증 실패를 보존된 rejected manifest 기반 repair로 자동 재시도하도록 corpus를 최대 4개 pass로 감쌌다. 각 pass는 완료 항목을 건너뛰고 FAILED 항목만 재개하며 실패 0이면 즉시 종료한다.
- Decisions: 첫 exact budget 요청의 fail-closed 검증은 유지한다. 후속 pass에서만 `repair_rejected_draft`가 저장된 오류와 draft를 사용해 제한된 복구를 수행하며, 4 pass 후에도 실패하면 무한 반복하지 않고 정확한 journal 실패로 남긴다.
- Files: `inference_data_ai_form_pipeline.py`, `tests/test_inference_data_ai_form_pipeline.py`, `HANDOFF.md`
- Verification: form pipeline·corpus workflow 테스트 16/16 통과. 현재 실행 중인 run 6은 시작 시 로드한 이전 코드이므로 종료 후 새 CLI 재실행에서 자동 retry wrapper가 적용된다.
- Next: run 6 종료 후 전체 CLI를 다시 실행해 자동 repair pass를 적용하고 corpus FAILED 0을 확인한다.

## 2026-07-26 08:12 - 실제 corpus 고병렬 안전 재개
- Completed: 3개 AI 슬롯의 처리 시간이 과도해 실제 run 6을 journal checkpoint에서 한정 중지하고, corpus를 workbook 6·AI 6·packet 4·COM 1·DB 1로 조정한 뒤 run 7로 재개했다. 중지 시 생성된 해당 run의 ephemeral Codex 자식만 정확히 종료했고 기존 Codex 세션은 건드리지 않았다.
- Decisions: DB 쓰기는 계속 1개로 직렬화하고 COM도 1개로 유지해 무결성과 Excel 안정성을 보존한다. AI/packet만 병렬화하며 RUNNING 기록은 다음 journal load에서 INTERRUPTED 후 자동 재시도한다.
- Files: `inference_data_ai_form_pipeline.py`, `HANDOFF.md`
- Verification: form pipeline 테스트 4/4 통과. 실제 run 7 journal options에서 workbookWorkers 6, AI 6, PACKET 4, DB 1과 6개 ephemeral 의미 분석 프로세스를 확인했다.
- Next: run 7 및 최대 4개 자동 repair pass의 완료·실패 수를 계속 감시한다.

## 2026-07-26 08:16 - content coverage repair 경로 보정
- Completed: AI draft의 원본 정량·결론 셀 coverage 검증이 runner 반환 뒤에만 실행되어 누락 draft가 repair 입력으로 저장되지 않던 결함을 수정했다. monolithic AI runner의 additional validator에서 canonical claim과 함께 complete content coverage를 검사하므로 실패 출력이 `.rejected` manifest가 되고 다음 pass가 정확한 누락 좌표를 보완할 수 있다.
- Decisions: 기존 import 직전 전체 검증은 이중 안전장치로 유지한다. 실제 run 7은 이전 코드를 로드했으므로 journal checkpoint에서 한정 재시작해 run 8부터 새 validator를 적용했다.
- Files: `inference_data_ai_workflow.py`, `tests/test_inference_data_ai_workflow.py`, `HANDOFF.md`
- Verification: workflow 테스트 13/13 통과. run 7의 pipeline Python과 해당 run이 만든 ephemeral Codex 자식만 종료 후 run 8을 시작했다.
- Next: run 8에서 기존 coverage 실패 2건이 `.rejected` repair로 전환되는지와 전체 처리량을 확인한다.

## 2026-07-26 08:18 - 신규 분석의 AI 문의 검색 포함 검증
- Completed: 실제 corpus에서 완료된 `01.MSM-X526TOP Report test upper material mold 1-2 date 03.27.2025.xlsx`를 대상으로 WPF canonical 질문과 동일한 `evidence-query`를 실행해 신규 Study가 rank 1로 검색되는 것을 확인했다.
- Decisions: 신규 분석은 `NEEDS_REVIEW` 상태에서도 관련 Study·원본 좌표·관측 데이터로 AI 문의 후보에 포함한다. 사람 승인 전에는 안전상 정량 effect 결론만 제외하며 원본 데이터와 citation은 숨기지 않는다.
- Files: `outputs/query-verification.x526top.json`, `HANDOFF.md`
- Verification: 질문 `X526TOP upper material mold 1-2 2025.03.27 시험 데이터` 결과 relevantStudyCount 1, source exclusion 0, publicDataId `DATA-1B94A5938898`, 원본 경로와 VERIFIED EVD 좌표 포함. answerEligibleEffectCount는 미승인 상태이므로 0으로 정상 fail-closed.
- Next: 전체 corpus 완료 후 여러 완료 원본과 최종 DB 집계로 검색 포함을 다시 확인한다.

## 2026-07-26 09:22 - corpus 병렬 처리량 확장
- Completed: 실제 run 8의 65분 처리량과 시스템 사용률을 측정하고, 완료·실패 36건은 journal/DB에 보존한 채 해당 pipeline 프로세스 트리만 종료했다. 16논리 코어 환경에서 workbook/AI 동시 작업을 최대 12개, packet을 최대 8개로 자동 확장했다.
- Decisions: COM과 DB 쓰기는 각각 1개로 직렬화해 Excel·SQLite 안정성을 유지한다. 작업 수는 `os.cpu_count()`에 따라 6~12 범위로 제한해 저사양 환경에서도 일괄 실행이 과도하게 확장되지 않게 한다.
- Files: `inference_data_ai_form_pipeline.py`, `HANDOFF.md`
- Verification: 변경 직후 form pipeline 단위 테스트 4/4 통과. PID 33960 아래의 이번 run 전용 Python/Codex 자식 19개만 확인 후 종료했고 잔존 프로세스 0개를 확인했다.
- Next: 저장된 checkpoint에서 전체 CLI를 재개해 처리량과 자동 repair pass 결과를 확인하고 FAILED 0까지 완료한다.

## 2026-07-26 09:29 - 병렬 분석 SQLite 잠금 해소
- Completed: 12개 병렬 run 9에서 실제 `OperationalError: database is locked`를 재현하고, AI 검증의 장시간 읽기 연결과 canonical import 쓰기가 SQLite 기본 DELETE journal에서 충돌하는 원인을 확인했다. corpus 병렬 시작 전에 canonical DB를 WAL 모드로 전환하는 준비 단계를 추가했다.
- Decisions: DB import/verify gate는 계속 1개로 유지한다. WAL은 장시간 읽기 분석과 단일 쓰기의 공존만 허용하며, 60초 busy timeout과 모드 전환 확인을 fail-closed로 적용한다.
- Files: `inference_data_ai_corpus_workflow.py`, `tests/test_inference_data_ai_corpus_workflow.py`, `HANDOFF.md`
- Verification: corpus workflow 및 form pipeline 단위 테스트 17/17 통과. 테스트 DB가 실제 `journal_mode=wal`로 전환되는 회귀 테스트를 추가했다.
- Next: 실제 DB에서 checkpoint CLI를 다시 재개하고 lock 실패 재발 여부와 전체 처리량을 확인한다.

## 2026-07-26 09:38 - 파일 단위 즉시 AI 복구
- Completed: WAL 적용 run 10에서 DB 잠금 재발 0을 확인했고, coverage 실패가 전체 corpus 1차 순회가 끝난 뒤에야 복구되던 지연을 제거했다. 각 workbook이 저장된 rejected draft를 즉시 사용해 최대 3회 재시도하고 성공한 뒤 다음 파일로 넘어가도록 변경했으며 16코어 환경의 workbook/AI 동시 작업을 16개, packet을 12개로 확장했다.
- Decisions: source fingerprint 변경과 최종 검증은 매 시도 후 그대로 fail-closed이며, COM/DB gate는 1개를 유지한다. 파일 단위 3회 뒤에도 실패한 항목은 기존 corpus pass 재시도에 남겨 완전 복구와 유한 종료를 모두 보장한다.
- Files: `inference_data_ai_corpus_workflow.py`, `inference_data_ai_form_pipeline.py`, `tests/test_inference_data_ai_corpus_workflow.py`, `HANDOFF.md`
- Verification: corpus workflow 및 form pipeline 단위 테스트 18/18 통과. 테스트에서 두 번 실패한 draft가 같은 corpus pass의 세 번째 시도에 완료되고 `workflowRetryAttempts=3`으로 기록되는 것을 검증했다.
- Next: 최종 코드로 checkpoint CLI를 재개하고 선택 361건의 terminal 결과와 FAILED 0을 확인한다.

## 2026-07-26 10:12 - 복합 locator 범위 coverage 오탐 수정
- Completed: 파일별 3회 repair 후에도 동일하게 실패한 `02. MSM-S931B Report check reason NG bako R&B.xlsx`의 locator·source packet·rejected manifest를 대조했다. `condition, measurements, and conclusion` 복합 범위의 조건 열 C4:C9가 source conclusion으로 오인되고, 조건 문구 `max spec`이 인접 수량/결과 셀의 MAX fieldRole로 전파되어 정확한 manifest도 거부되는 원인을 수정했다.
- Decisions: 복합 역할 범위는 위쪽 20행 내 실제 `Result/Conclusion/Decision` 헤더가 있는 열만 narrative conclusion으로 분류한다. MIN/MAX/AVERAGE fieldRole은 4단어를 넘는 조건 서술에서 추론하지 않아 조건명 속 단어가 정량 필드 의미를 덮지 못하게 한다.
- Files: `inference_data_ai_content_coverage.py`, `tests/test_inference_data_ai_content_coverage.py`, `HANDOFF.md`
- Verification: content coverage 및 workflow 테스트 73/73 통과. 실제 실패 manifest를 새 inventory로 재검증해 정량 25/25, 결론 G4:G9 6/6, binding error 0으로 완전 통과했다.
- Next: 메모리에 이전 validator가 로드된 CLI만 checkpoint 재시작하고, 해당 파일의 완료 전환 및 전체 FAILED 0을 확인한다.

## 2026-07-26 10:55 - 끝 공백 Excel sheet 근거 검증 수정
- Completed: `08. MSM-X526,X626 Report test value gauss SPK before input Module date 2025.04.28.xlsx`의 반복 `evidence sheet is not present` 실패를 실제 Capture v2와 117KB rejected manifest로 재현했다. 실제 sheet명 `526TOP `의 끝 공백이 measurement series용 `_text()`에서만 제거되어 exact DB 조회가 `526TOP`으로 실패하던 계약 검증 결함을 수정했다.
- Decisions: sheet 이름은 공백만인 경우에는 계속 거부하지만, Excel/Capture가 저장한 앞뒤 공백은 source identity의 일부로 그대로 evidence checker에 전달한다. 일반 evidence range가 이미 따르던 exact-name 동작과 measurement series를 일치시킨다.
- Files: `inference_data_ai_study_contract.py`, `tests/test_inference_data_ai_study_contract.py`, `HANDOFF.md`
- Verification: study contract 및 workflow 테스트 37/37 통과. 실제 실패 manifest를 canonical revision 1117/capture revision 1047의 DB evidence checker로 재검증해 `REAL_MANIFEST_EVIDENCE_PASS`를 확인했다.
- Next: 이전 계약 모듈을 로드한 현재 CLI만 checkpoint 재시작하고 해당 파일의 완료 전환 및 전체 FAILED 0을 확인한다.

## 2026-07-26 11:39 - 안전성 거부 AI repair 재사용 차단
- Completed: `09.MSM-S931B Report test Bako new programe date 2026.7.08.xlsx`의 3회 실패를 실제 130KB repair artifact까지 대조해, 허용된 rate-pair 경로 밖을 변경하여 거부된 AI 수정본이 다음 resume 입력으로 다시 선택되는 원인을 수정했다.
- Decisions: 안전성 검사에서 거부된 수정본은 진단용으로 계속 보존하되 artifact SHA-256이 연결된 `.repair-rejected.unsafe.json` 표식을 저장하고, 같은 해시의 거부본은 재사용 후보에서 제외한다. 이후 일반 검증 실패가 새 수정본을 기록하면 이전 안전 표식은 제거한다.
- Files: `inference_data_ai_semantic_ai.py`, `tests/test_inference_data_ai_semantic_ai.py`, `HANDOFF.md`
- Verification: reference repair resume, rate-pair deterministic repair 관련 회귀 테스트 3/3 통과. 새 테스트에서 거부본 보존·unsafe 표식 생성·다음 실행의 원본 rejected baseline 재선택을 확인했다.
- Next: 현재 corpus pass 종료 후 기존 실패 artifact에 unsafe 표식을 연결하고 새 코드로 checkpoint를 재개해 해당 파일 및 전체 FAILED 0을 확인한다.

## 2026-07-26 12:10 - 신뢰 repair의 rate-pair 로컬 연속 복구
- Completed: 새 코드 run 14에서 안전 거부본은 정상 배제됐지만, 일반 repair 결과에 남은 다수의 근거 없는 denominator가 항목마다 새 AI 호출을 요구하는 병목을 재현했다. 안전 표식이 없는 신뢰 가능한 repair baseline은 validator가 지목한 numerator/denominator pair만 결정적으로 비우고 연속 재검증하도록 확장했다.
- Decisions: 로컬 수정은 validator 오류에 포함된 정확한 Study/Outcome/Observation 경로의 두 필드만 허용하고 나머지 JSON 동일성을 계속 강제한다. 안전 거부 표식이 붙은 artifact에는 이 경로를 적용하지 않는다.
- Files: `inference_data_ai_semantic_ai.py`, `tests/test_inference_data_ai_semantic_ai.py`, `HANDOFF.md`
- Verification: 신뢰 repair baseline의 무-AI rate-pair 복구, unsafe artifact 배제, 기존 pair 안전성 회귀 테스트 3/3 통과.
- Next: 수정이 로드된 corpus run 15에서 Bako new programme 파일이 COMPLETED로 전환되는지 확인한 뒤 전체 선택 파일을 FAILED 0까지 계속 처리한다.

## 2026-07-26 12:19 - 다중 rate-pair fixpoint 최적화
- Completed: 실제 Bako manifest에서 로컬 복구는 작동했지만 16개 병렬 validator의 GIL 경쟁 아래 pair마다 함수 전체를 재진입해 처리 속도가 느린 것을 확인했다. 신뢰 baseline을 메모리에 유지한 채 연속 unsupported pair를 모두 정리하고 각 안전 투영만 checkpoint에 기록하는 fixpoint로 변경했다.
- Decisions: 각 반복은 이전과 동일하게 validator가 지정한 경로만 수정하며 같은 경로가 재등장하면 fail-closed 한다. 별도 validation 오류로 전환될 때만 일반 repair 경로로 돌아간다.
- Files: `inference_data_ai_semantic_ai.py`, `tests/test_inference_data_ai_semantic_ai.py`, `HANDOFF.md`
- Verification: 두 개의 독립 unsupported denominator가 있는 trusted repair fixture를 AI 호출 0회로 모두 정리하는 테스트를 포함해 관련 회귀 4/4 통과.
- Next: corpus run 16에서 실제 다중 pair 파일의 DRAFT/IMPORT/VERIFY 완료 전환을 확인하고 전체 corpus를 계속 처리한다.

## 2026-07-26 12:35 - DB 기반 rate-pair batch 발견
- Completed: 실제 두 번째 Study에서 pair별 전체 manifest 검증이 반복되는 비용을 제거했다. canonical DB의 각 observation evidence 범위를 한 번씩 조회해 numerator 또는 denominator가 실제 Capture 숫자에 없는 모든 경로를 열거하고, semantic fixpoint에 batch callback으로 전달한다.
- Decisions: batch는 canonical pair 계약을 통과한 완전한 두 값에만 적용하고, 하나라도 근거 숫자에 없을 때 validator와 동일하게 두 필드를 함께 비운다. 첫 validator 경로가 batch에 없거나 잘못된 경로가 오면 fail-closed 한다.
- Files: `inference_data_ai_study_import.py`, `inference_data_ai_workflow.py`, `inference_data_ai_semantic_ai.py`, `tests/test_inference_data_ai_semantic_ai.py`, `HANDOFF.md`
- Verification: 관련 semantic/workflow 회귀 5/5 통과. 실제 Bako rejected manifest와 canonical revision에서 남은 unsupported 경로 55개를 0.006초에 열거했으며 첫/끝 경로를 확인했다.
- Next: batch callback이 로드된 corpus run 17에서 실제 파일을 즉시 DRAFT/IMPORT/VERIFY 완료시키고 전체 corpus를 이어간다.

## 2026-07-26 13:00 - Bako new programme 실제 완료
- Completed: `09.MSM-S931B Report test Bako new programe date 2026.7.08.xlsx`에 batch repair를 실제 적용해 unsupported pair 55개를 로컬 정리하고, 누락된 formula cell 8개는 bounded AI coverage repair로 보완한 뒤 canonical import와 integrity verify까지 완료했다.
- Decisions: 분석 상태는 자동 승인하지 않고 `NEEDS_REVIEW`로 유지하되 corpus 의미 완료는 answer-visible `publicAnalysisId`가 생성된 경우에만 `COMPLETED`로 기록한다.
- Files: 실제 corpus run `ingest-run_9a6d29a7119ecc8869ef5e8d` artifacts, canonical `InputDataFinish.sqlite`, `HANDOFF.md`
- Verification: workbook journal 최종 상태 `NEEDS_REVIEW`, currentStage 공백, corpus journal `COMPLETED=36`, `FAILED=0`, `RUNNING=325`, `PENDING=25` 확인.
- Next: 같은 CLI run 17을 중단하지 않고 나머지 선택 파일을 모두 의미 분석·DB 반영하고 최종 FAILED 0을 확인한다.

## 2026-07-26 13:28 - 복합 outcome 하위 헤더와 공유 sample-size coverage
- Completed: `07. MSM-X526BOTTOM Report test SPK use Dome new supplier date 2025.05.03.xlsx`의 반복 coverage 실패를 실제 source packet·locator·118KB repair manifest로 재현했다. 복합 Total NG/count-rate outcome이 Q13/S13 `NG rate` 하위 헤더를 exact evidence로 보존해도 originalLabel 단일 필드와 다르다는 이유로 미포함 처리되고, `sample_size` outcome의 한 source cell을 여러 defect-category Arm이 공유하면 중복 scalar로 거부되던 원인을 수정했다.
- Decisions: 복합 outcome은 semantic label 셀을 정확히 덮는 evidence의 `sourceText`가 캡처 원문과 정확히 같을 때 하위 헤더 보존으로 인정한다. source 재사용은 명시적 `metricType=sample_size`에만 허용하고 일반 scalar outcome의 injective binding은 유지한다.
- Files: `inference_data_ai_content_coverage.py`, `tests/test_inference_data_ai_content_coverage.py`, `HANDOFF.md`
- Verification: 관련 coverage 회귀 3/3 통과. 실제 repair manifest는 required quantitative 82/82, semantic label 38/38, binding error 0, narrative conclusion 4/4로 `REAL_DOME_COVERAGE_PASS` 확인.
- Next: 새 validator가 로드된 checkpoint CLI에서 Dome supplier 파일을 COMPLETED로 전환하고 전체 FAILED 0까지 계속 처리한다.

## 2026-07-26 13:44 - 유효 repair 체크포인트 무호출 승격
- Completed: 실제 run 18에서 현재 validator를 통과하는 `canonical-study-manifest.repair-rejected.json`도 새 전체 draft AI 호출로 다시 생성하던 재개 결함을 확인하고, source/evidence/additional validator를 모두 통과한 최신 체크포인트를 canonical manifest로 즉시 승격하도록 수정했다.
- Decisions: 체크포인트 이름이 과거의 rejected 상태를 나타내더라도 현재 검증 전체를 통과하면 동일 내용을 모델에 재요청하지 않는다. 안전 거부 SHA 표식과 불일치하는 repair 산출물만 재사용한다는 기존 fail-closed 규칙은 유지한다.
- Files: `inference_data_ai_semantic_ai.py`, `tests/test_inference_data_ai_semantic_ai.py`, `HANDOFF.md`
- Verification: 유효 repair 무호출 승격, trusted rate-pair repair, 복합 outcome 하위 헤더 coverage 회귀 3/3 통과.
- Next: run 19에서 체크포인트를 재개해 유효 기존 draft를 빠르게 승격하고 새로 드러난 개별 coverage 실패를 실제 산출물 기준으로 처리한다.

## 2026-07-26 13:50 - 최신 유효 체크포인트 선택
- Completed: 중단된 AI 호출이 더 최신인 불완전 `.rejected.json`을 남기면 현재 validator를 통과하는 이전 `.repair-rejected.json`이 가려지는 실제 Dome 재개 사례를 수정했다. 최신 후보가 검증에 실패한 경우에만 이전 안전 후보를 최신순으로 검사해 전 validator를 통과한 첫 산출물을 canonical manifest로 승격한다.
- Decisions: 최신 시각 자체보다 현재 source/evidence/content validator 전체 통과를 우선한다. 최신 후보가 유효할 때는 이전 후보를 검사하지 않으며, 안전 거부 표식이 현재 repair SHA와 일치하는 후보는 계속 제외한다.
- Files: `inference_data_ai_semantic_ai.py`, `tests/test_inference_data_ai_semantic_ai.py`, `HANDOFF.md`
- Verification: 최신 invalid/이전 valid 승격, unsafe repair 제외, trusted batch rate-pair repair 검증 호출 수 유지 회귀 4/4 통과.
- Next: run 20에서 Dome의 이전 유효 repair를 즉시 승격해 완료시키고 전체 corpus를 계속 처리한다.

## 2026-07-26 13:56 - Dome 신규 supplier canonical 완료
- Completed: `07. MSM-X526BOTTOM Report test SPK use Dome new supplier date 2025.05.03.xlsx`의 이전 repair를 현재 validator 전체로 재검증·승격하고 canonical DB import 및 integrity verify까지 완료해 corpus 상태를 `COMPLETED`로 전환했다.
- Decisions: AI 생성 분석은 자동 승인하지 않고 `NEEDS_REVIEW`로 유지하되, answer-visible public analysis가 생성되고 integrity가 통과한 경우 corpus 처리는 완료로 기록한다.
- Files: 실제 workbook run `ingest-run_0f5e53827553817e5c063d54` 산출물, canonical `InputDataFinish.sqlite`, `HANDOFF.md`
- Verification: workflow 최종 `NEEDS_REVIEW`, studies 2, `publicAnalysisId=ANALYSIS-0B788D3B7F6E`, `integrityOk=True`; corpus `COMPLETED=37`, `FAILED=0`, `RUNNING=324`, `PENDING=25`.
- Next: run 20을 중단하지 않고 나머지 324개 대상과 PT plating coverage 복구를 끝까지 처리한다.

## 2026-07-26 14:12 - 짧은 원문 추가시험 결론 검증
- Completed: `07.MSM-X626BOTTOM Report test change drying time bond fix PT date 2025.02.20.xlsx`가 C51의 exact source text `Test more`를 올바르게 인용해도 3단어 미만이라는 이유로 SOURCE_CONCLUSION에서 반복 거부되던 false negative를 수정했다.
- Decisions: 일반 `PASSED` 같은 categorical status는 계속 결론으로 거부한다. exact captured evidence·claim 일치가 먼저 증명된 경우에만 명시적 `test more` 지시문을 짧은 source conclusion marker로 인정한다.
- Files: `inference_data_ai_study_import.py`, `tests/test_inference_data_ai_study_import.py`, `HANDOFF.md`
- Verification: 짧은 추가시험 지시문 승인, 반복 PASSED 거부, 기존 source decision import 회귀 3/3 통과; 실제 repair manifest가 `REAL_SHORT_CONCLUSION_PASS`.
- Next: 새 validator로 corpus를 재개해 건조시간 파일을 기존 유효 repair에서 승격하고 남은 대상 분석을 계속한다.

## 2026-07-26 14:13 - C-MG 색상 비교 canonical 완료
- Completed: `10. TIU L5S3-01 - Report test material C-MG difference color (Make sample check tension and send to KR)- date 2026.05.05.xlsx`의 재개 분석, canonical import 및 integrity verify를 완료했다.
- Decisions: AI 분석은 `NEEDS_REVIEW`로 유지하되 public analysis와 검증된 source evidence가 생성된 것을 corpus 완료 기준으로 적용했다.
- Files: 실제 workbook run `ingest-run_9a00f1064a0b814bcd33daa2` 산출물, canonical `InputDataFinish.sqlite`, `HANDOFF.md`
- Verification: `publicAnalysisId=ANALYSIS-5CCE3F0A871E`; corpus `COMPLETED=38`, `FAILED=0`, `RUNNING=323`, `PENDING=25`.
- Next: run 21을 계속 실행해 건조시간 import 및 다음 신규 파일들을 완료한다.

## 2026-07-26 14:20 - 건조시간 PT 시험 canonical 완료
- Completed: `07.MSM-X626BOTTOM Report test change drying time bond fix PT date 2025.02.20.xlsx`의 유효 repair를 승격하고 대용량 canonical DB import 및 전체 integrity verify를 완료했다.
- Decisions: 정확한 원문 결론을 포함한 AI 분석은 자동 승인하지 않고 `NEEDS_REVIEW`로 유지하며, public analysis와 source evidence가 answer-visible인 상태를 완료로 기록했다.
- Files: 실제 workbook run `ingest-run_ab11587f9dd8719c8a135b58` 산출물, canonical `InputDataFinish.sqlite`, `HANDOFF.md`
- Verification: studies 5, `publicAnalysisId=ANALYSIS-4631DDC37D59`, `integrityOk=True`; corpus `COMPLETED=39`, `FAILED=0`, `RUNNING=322`, `PENDING=25`.
- Next: 같은 run 21에서 나머지 322개 대상의 DRAFT·import를 계속 완료한다.

## 2026-07-26 14:29 - 병합 결과 헤더 recordId 충돌 수정
- Completed: 016/025 MSU-L20S15 bond 파일의 동일 deterministic result-table part가 매번 duplicate recordId로 실패하는 원인을 순수 builder에서 재현했다. `Peak`/`Total` 같은 병합 헤더 한 셀이 여러 결과 열을 가리킬 때 각 outcome identity에 해당 열의 첫 값 셀도 포함하도록 수정했다.
- Decisions: 결과 outcome을 헤더 문구만으로 합치지 않고 헤더 cell + 실제 값 column cell의 복합 source identity로 구분한다. payload key만 ID에 넣는 비-source identity 방식은 사용하지 않는다.
- Files: `inference_data_ai_staged_draft_v2.py`, `tests/test_inference_data_ai_staged_draft_v2.py`, `HANDOFF.md`
- Verification: 병합 헤더 고유 identity, 기존 MASK, FO deterministic 회귀 3/3 통과. 실제 016/025 대상 part가 각각 records 174/174 고유 ID로 `validate_fragment_v2` 통과.
- Next: 새 builder로 corpus를 재개해 016/025 파일의 반복 duplicate 실패를 해소하고 다음 stage까지 처리한다.

## 2026-07-26 14:36 - X526 TOP PT plating canonical 완료
- Completed: `10. MSM-X526 TOP report check gauss value when use PT plating 1-3 micromet date 2025.02.19.xlsx`의 반복 coverage repair를 끝내고 canonical import와 integrity verify까지 완료했다.
- Decisions: 원본 source cell coverage를 전부 충족한 분석만 public analysis로 반영하고 자동 승인 없이 `NEEDS_REVIEW`로 유지한다.
- Files: 실제 workbook run `ingest-run_a59fabce744fee86d78c6012` 산출물, canonical `InputDataFinish.sqlite`, `HANDOFF.md`
- Verification: `publicAnalysisId=ANALYSIS-7323FF8BECBD`; corpus `COMPLETED=40`, `FAILED=0`, `RUNNING=321`, `PENDING=25`.
- Next: run 22에서 016/025 staged fragment와 나머지 대상 분석을 계속한다.

## 2026-07-26 14:53 - 분할 결과표 continuation deterministic 처리
- Completed: 016/025 bond 파일의 400셀 continuation part에서 결과표 제목이 primary가 아닌 shared context `RELIABILITY TEST RESULT`에만 남아 projector가 실행되지 않고 AI fragment가 E53:P54/E79:P79 정량 셀을 반복 누락하던 원인을 수정했다.
- Decisions: source segmentation이 보존한 exact shared context도 result-table 판별에 사용하고, 제목 내 독립 단어 `RESULT`/`RESULTS`를 인식한다. 실제 record/evidence 생성은 계속 owned primary cell로 제한한다.
- Files: `inference_data_ai_staged_draft_v2.py`, `tests/test_inference_data_ai_staged_draft_v2.py`, `HANDOFF.md`
- Verification: continuation title, 병합 헤더 identity, 기존 MASK/FO 회귀 4/4 통과. 실제 016 continuation records 291/291 고유, 025 continuation 283/283 고유이며 `validate_fragment_v2` 통과.
- Next: 새 continuation projector가 로드된 corpus run에서 016/025의 반복 numeric binding 실패를 해소하고 전체 분석을 계속한다.

## 2026-07-26 14:54 - X526B-Top Bako/Hearing canonical 완료
- Completed: `100. MSM-X526B-Top Report check bako and hearing class D date 2026.5.19 .xlsx`의 신규 분석, canonical import 및 integrity verify를 첫 corpus 시도에서 완료했다.
- Decisions: 기존과 동일하게 public analysis는 answer-visible하게 반영하되 AI 결과는 `NEEDS_REVIEW`로 유지한다.
- Files: 실제 workbook run `ingest-run_f166d3546d4f44c4ee243dc9` 산출물, canonical `InputDataFinish.sqlite`, `HANDOFF.md`
- Verification: `publicAnalysisId=ANALYSIS-C01CF25647C1`; corpus `COMPLETED=41`, `FAILED=0`, `RUNNING=320`, `PENDING=25`.
- Next: run 23에서 016/025의 나머지 AI fragment와 전체 corpus를 계속 처리한다.

## 2026-07-26 15:00 - continuation outcome label 병합 안정화
- Completed: 025 continuation fragment가 정량 binding을 통과한 뒤 동일 `result_c5_value` outcome의 originalLabel이 part별 첫 행 문구로 달라 병합 충돌하던 원인을 수정했다. shared context의 exact `RELIABILITY TEST RESULT` 셀을 모든 continuation part의 공통 title과 synthetic column label 기준으로 사용한다.
- Decisions: 분할 경계와 무관하게 같은 logical Study·column outcome은 동일 source title 기반 label을 생성한다. shared context는 identity/label 기준으로만 사용하고 측정 evidence는 owned primary cell에 유지한다.
- Files: `inference_data_ai_staged_draft_v2.py`, `tests/test_inference_data_ai_staged_draft_v2.py`, `HANDOFF.md`
- Verification: 관련 deterministic 회귀 4/4 통과. 실제 016/025 각각 result-table deterministic parts 11개, records 1,767개가 개별 검증을 통과했고 cross-part entity label conflict 0.
- Next: 수정된 공통 label projector로 corpus를 재개해 016/025 전체 fragment 병합·import를 완료한다.

## 2026-07-26 15:05 - fragment contract v7 cache 무효화
- Completed: 수정된 projector는 충돌이 없지만 기존 v6 fragment/provenance가 재사용되어 과거 part별 label이 다시 병합되던 문제를 확인하고 fragment contract를 v7로 승격했다.
- Decisions: deterministic projector가 record identity/label 의미를 바꾸면 fragment contract identity도 반드시 변경한다. v6 plan·part·final provenance는 새 run에서 fail-closed로 무효화하고 v7 계획으로 재생성한다.
- Files: `inference_data_ai_staged_draft_v2.py`, `HANDOFF.md`
- Verification: contract version별 plan/part ID binding 및 result-table 회귀 3/3 통과. 실제 016/025의 저장된 v6 plan 모두 live v7에서 `STALE_PLAN_REJECTED`.
- Next: v7 계획으로 corpus를 재개해 오래된 deterministic fragment를 섞지 않고 016/025를 다시 병합한다.

## 2026-07-26 15:13 - 다중 fragment 엔터티 병합 충돌 수정
- Completed: 같은 canonical 엔터티가 여러 source fragment에서 선언될 때 선언별 `entityId`가 달라 병합이 중단되던 문제를 수정했다. `entityId`만 비의미 transport 필드로 제외하고 나머지 엔터티 필드는 계속 엄격하게 충돌 검사한다.
- Decisions: `entityId`는 source-bound `recordId`이며 canonical 엔터티 의미가 아니므로 병합 대상에서 제외한다. 병합 동작 변경을 provenance에 반영하도록 consolidator contract를 v4로 올렸다.
- Files: `inference_data_ai_staged_draft_v2.py`, `tests/test_inference_data_ai_staged_draft_v2.py`, `HANDOFF.md`
- Verification: 신규/기존 병합 단위 테스트 2/2 통과. 실제 실패 산출물 016과 025를 `project_canonical_manifest`로 재투영해 각각 studies 20/21, records 1,786/1,788로 모두 통과했고, 실제 의미 필드 `unit` 충돌은 계속 거부됨을 확인했다.
- Next: consolidator v4로 전체 corpus 파이프라인을 재개해 361개 선택 workbook을 COMPLETED 상태로 만들고 DB·AI 질의·WPF를 최종 검증한다.

## 2026-07-26 15:37 - aggregate identity 및 Before/After coverage 복구
- Completed: `08. MSM-X526,X626...2025.04.28.xlsx`에서 aggregate value range를 identity range로 잘못 넣은 AI 복구 산출물을 결정적으로 교정하고, `(2).xlsx`에서 bare `Before`/`After`가 행의 모든 숫자를 BASELINE/CHANGED 필드로 오분류하던 coverage 문제를 수정했다.
- Decisions: aggregate range는 valueRange 전체 열/행과 정확히 정렬되는 전치 오류만 source identity 축으로 변환한다. `Before`/`After` 단독 라벨은 Arm stage이며 `Before value`/`After value`처럼 명시적 필드 수식어가 있을 때만 fieldRole로 사용한다.
- Files: `inference_data_ai_semantic_ai.py`, `inference_data_ai_content_coverage.py`, `tests/test_inference_data_ai_semantic_ai.py`, `tests/test_inference_data_ai_content_coverage.py`, `HANDOFF.md`
- Verification: 관련 단위 테스트 4/4 통과. 실제 08번 repair-rejected manifest의 9개 series를 교정한 뒤 canonical contract 통과. 실제 `(2).xlsx` coverage는 required 197/197, semantic 89/89, categorical 3/3 통과.
- Next: 전체 corpus 파이프라인을 재개해 새로 드러나는 실패 유형을 계속 제거하고 361개 선택 workbook을 모두 완료한다.

## 2026-07-26 15:40 - 전체 CLI staged 임계값 노출
- Completed: `form-pipeline-complete`에 `--draft-monolithic-max-bytes`를 추가해 복잡한 신규 엑셀을 전체 CLI 실행 안에서 source-complete staged 분석으로 강제할 수 있게 했다.
- Decisions: 기본값 400,000 bytes는 호환성을 유지하고, 현재 전체 재개에서는 80,000 bytes를 사용한다. 88KB prompt의 `10.MSM-S931...2025.06.17.xlsx`는 monolithic AI가 상세 223개 수치를 누락했지만 staged result-table projector는 핵심 `17.6` chunk의 312/312 source cell, records 286을 결정적으로 검증했다.
- Files: `inference_data_ai_form_pipeline.py`, `inference_data_ai_cli.py`, `tests/test_inference_data_ai_form_pipeline.py`, `tests/test_inference_data_ai_cli.py`, `HANDOFF.md`
- Verification: CLI 기본값 및 corpus 옵션 전달 단위 테스트 2/2 통과. 실제 081 run을 staged 계획으로 구성해 9 parts/10 studies를 만들었고 result-table deterministic parts 3개가 모두 `validate_fragment_v2`를 통과했다.
- Next: CLI 임계값 80,000으로 전체 파이프라인을 재개해 081과 유사한 중형 결과표도 source-complete로 처리한다.

## 2026-07-26 15:44 - `(2).xlsx` canonical import 완료
- Completed: `(2).xlsx`의 bare stage fieldRole 수정이 실제 corpus에서 coverage, import, integrity verify를 통과해 신규 canonical 분석으로 등록됐다.
- Decisions: source-backed 분석 상태는 기존 정책대로 `NEEDS_REVIEW`를 유지하되 public analysis와 AI evidence 검색에는 포함한다.
- Files: 실제 workbook run `ingest-run_f03820fe39671ed92bc8ba1b` 산출물, `InputDataFinish.sqlite`, `HANDOFF.md`
- Verification: `publicAnalysisId=ANALYSIS-75ACF0FCCFAD`; corpus `COMPLETED=42`, `FAILED=0`, `RUNNING=319`, `PENDING=25`.
- Next: 08번 aggregate 복구와 081 staged 결과표를 완료시키고 전체 corpus 처리를 계속한다.

## 2026-07-26 15:50 - staged 서술형 비율의 숫자 승격 차단
- Completed: 081 run의 `Bond 201 around enclosure assy ok 0/20` 서술 셀을 numerator/denominator/sampleSize 숫자 근거로 승격해 fragment가 실패하던 문제를 안전한 정규화로 수정했다.
- Decisions: 실제 numeric cell이 없으면 embedded narrative가 보존하는 `valueText`만 유지하고 숫자 필드는 null로 내린다. `1/8 pcs`처럼 전체 셀이 엄격한 count ratio인 경우만 기존 숫자 claim을 유지하며, Arm text sample size도 `10pcs`/`EA`/`samples` 단위가 명시된 경우만 허용한다. validator contract를 v6로 올려 구 fragment를 무효화했다.
- Files: `inference_data_ai_staged_draft_v2.py`, `inference_data_ai_staged_runner_v2.py`, `tests/test_inference_data_ai_staged_draft_v2.py`, `HANDOFF.md`
- Verification: 관련 단위 테스트 2/2 통과. 실제 081 rejected part를 v6 envelope로 재정규화해 records 8 전체 검증 통과, embedded ratio 숫자 필드는 모두 null이고 Arm sampleSize도 null임을 확인했다.
- Next: validator v6와 80KB staged 임계값으로 전체 corpus를 재개한다.

## 2026-07-26 15:56 - staged percentage/rate 수치 binding 보정
- Completed: 016 run의 cached percent formula `I43=0.5625`가 canonical `ratePpm=562500`과 정확히 같은 값인데 fragment coverage가 단위 스케일을 비교하지 않아 누락으로 판단하던 문제를 수정했다.
- Decisions: exact numeric equality 외에 `ratePpm = source × 1,000,000`, percent number format의 `valueNumber = source × 100`만 추가로 허용한다. 표시 반올림값 56.3은 cached 56.25와 동일한 claim으로 인정하지 않는다. validator contract를 v7로 올렸다.
- Files: `inference_data_ai_staged_draft_v2.py`, `tests/test_inference_data_ai_staged_draft_v2.py`, `HANDOFF.md`
- Verification: 관련 단위 테스트 2/2 통과. 실제 016 rejected part를 v7 envelope로 재검증해 records 5 및 `ratePpm=562500` binding 전체 통과.
- Next: validator v7로 전체 corpus를 재개한다.

## 2026-07-26 15:59 - TIU C11-20 AWF canonical 완료
- Completed: `100. TIU C11-20 Report Test AWF machine improve wire offset pad 2026.7.10.xlsx` 신규 분석, canonical import, integrity verify를 v7 corpus에서 완료했다.
- Decisions: public analysis는 source-backed `NEEDS_REVIEW` 정책으로 AI evidence 검색에 포함한다.
- Files: 실제 workbook run `ingest-run_be62266ef6fae886b896016e` 산출물, `InputDataFinish.sqlite`, `HANDOFF.md`
- Verification: `publicAnalysisId=ANALYSIS-6FDBF17DEFDA`; corpus `COMPLETED=43`, `FAILED=0`, `RUNNING=318`, `PENDING=25`.
- Next: v7 corpus의 나머지 318건과 주요 수정 대상 4건을 계속 처리한다.

## 2026-07-26 16:06 - isolated observation Arm 보완
- Completed: 82ff run의 독립 `2. Particle` 결과 row가 outcome과 수치는 만들었지만 Arm key가 없어 fragment 검증에 실패하던 문제를 수정했다.
- Decisions: 같은 logical Study에 Arm 선언이 1개면 그 key를 사용하고, 0개이면 observation evidence 안의 정확한 비수치 source label로 role OTHER의 descriptive Arm을 1개만 생성한다. 여러 Arm이 이미 있거나 source label이 없으면 추론하지 않고 기존 fail-closed 동작을 유지한다. validator contract를 v8로 올렸다.
- Files: `inference_data_ai_staged_draft_v2.py`, `inference_data_ai_staged_runner_v2.py`, `tests/test_inference_data_ai_staged_draft_v2.py`, `HANDOFF.md`
- Verification: 관련 단위 테스트 2/2 통과. 실제 82ff rejected part를 v8로 정규화해 source label `2. Particle`, records 5, 두 observation의 동일 descriptive Arm binding으로 전체 fragment 검증 통과.
- Next: validator v8로 전체 corpus를 재개한다.

## 2026-07-26 16:20 - observation 행 식별자 증거 보강
- Completed: 관찰 결과가 `replicateKey`에 보존한 숫자 행 식별자를 같은 행의 정확한 원본 셀 증거로 자동 연결하여, 실제 X526/X626 조각에서 B4:B13이 미표현 숫자로 탈락하던 문제를 해결했다.
- Decisions: 숫자 원본 텍스트가 `replicateKey`의 선두 토큰과 정확히 일치하고 기존 관찰 증거와 같은 행에 있을 때만 `REPLICATE_IDENTITY` 증거를 추가한다. `1.2`가 `11.2`에 오인 매칭되지 않도록 구분자를 강제하고 validator contract를 v9로 올렸다.
- Files: `inference_data_ai_staged_draft_v2.py`, `inference_data_ai_staged_runner_v2.py`, `tests/test_inference_data_ai_staged_draft_v2.py`, `HANDOFF.md`
- Verification: 관련 단위 테스트 3/3 통과. 실제 `ingest-run_8029314f7be3750302afbbcb` rejected fragment 36 records를 v9 envelope로 재구성해 B4:B13의 10개 숫자 식별자 증거를 추가한 뒤 전체 fragment 검증 통과.
- Next: 실행 중인 v8 corpus 프로세스를 안전하게 종료하고 v9/80KB 설정으로 전체 corpus를 재개한다.

## 2026-07-26 16:45 - staged source binding과 coverage v10 보강
- Completed: v9 실데이터에서 확인된 disposition-only 증거 링크, 누락 disposition, 빈 OUTCOME 필수 필드, 잘못된 series headerRange, 임의 Arm role, 연속 No. 행 식별자와 혼합 outcome/conclusion 영역 오분류를 보강했다.
- Decisions: 모델이 명시한 RECORD_EVIDENCE 링크만 정확한 한 셀 증거로 물질화하고, 미작성 owned cell은 CONTEXT_ONLY로 완결하되 숫자 의미 검증은 기존 fail-closed 검사를 유지한다. 원본 라벨로만 필수 필드를 채우며 미지원 Arm role은 보수적 OTHER로 정규화한다. validator contract를 v10으로 올렸다.
- Files: `inference_data_ai_staged_draft_v2.py`, `inference_data_ai_staged_runner_v2.py`, `inference_data_ai_content_coverage.py`, `inference_data_ai_semantic_ai.py`, `tests/test_inference_data_ai_staged_draft_v2.py`, `tests/test_inference_data_ai_content_coverage.py`, `tests/test_inference_data_ai_semantic_ai.py`, `HANDOFF.md`
- Verification: staged/content/semantic 관련 테스트 151/151 통과. 실제 8029314f와 82ff901b fragment는 선언 증거 추가 후 통과, 123908ae fragment는 headerRange와 464 disposition 완결 후 통과. 실제 bf9389d5 manifest는 required 6/6, 연속 행 식별자 30개 제외, source conclusion M3만 보존해 coverage 통과.
- Next: v10/80KB 전체 corpus CLI를 재개하고 새 실패 유형 또는 완료 증가를 추적한다.

## 2026-07-26 16:48 - Hansol Class-D canonical 완료
- Completed: `100.1 Test X526 Class D Hansol.xlsx`의 신규 분석, canonical import와 integrity verify를 v10 corpus에서 완료했다.
- Decisions: 30개 No. 연속값은 행 식별자로 유지하고, `Total result` 아래 M3 문장만 source conclusion으로 보존하며 공개 분석은 `NEEDS_REVIEW`로 AI 문의 대상에 포함한다.
- Files: 실제 workbook run `ingest-run_bf9389d5233f08bf479372c3` 산출물, `InputDataFinish.sqlite`, `HANDOFF.md`
- Verification: `publicAnalysisId=ANALYSIS-9F95CCB9E4D4`, studies 1, integrityOk true. corpus `COMPLETED=44`, `RUNNING=317`, `PENDING=25`, `FAILED=0`.
- Next: v10 corpus의 나머지 317개와 현재 staged draft 파일을 계속 처리한다.

## 2026-07-26 16:54 - Hansol 신규 데이터 AI 문의 포함 검증
- Completed: 신규 적재된 Hansol Class-D 데이터를 고유 원본 serial `T3005261W1F0231`로 실제 `evidence-query`에 문의해 canonical 데이터와 인용 근거가 함께 반환됨을 확인했다.
- Decisions: 승인되지 않은 효과 계산에는 사용하지 않고 `excludedCandidates`의 source-backed descriptive outcomes로 제공해 review gate를 유지한다.
- Files: `InputDataFinish.sqlite`, `HANDOFF.md`
- Verification: 관련 Study 후보 1개가 정확한 `100.1 Test X526 Class D Hansol.xlsx`와 `ANALYSIS-9F95CCB9E4D4`로 검색됨. descriptive outcomes 8개, arm observation 그룹 41개, 원본 evidence/publicEvidenceId 90개 반환.
- Next: 전체 corpus 완료 뒤 추가 신규 파일 2건에도 같은 AI 문의 검증을 반복한다.

## 2026-07-26 16:59 - 결합 범위 series header v11 보강
- Completed: 실제 08 MSM X526/X626 조각에서 evidence가 H13:H52 결합 범위인데 payload headerRange가 빈 H14를 가리켜 실패한 추가 사례를 해결했다.
- Decisions: 명시 라벨 일치 증거가 없을 때만 valueRange의 같은 열·직전 4행 이내·해당 record가 실제 인용한 비수치 캡처 셀 중 가장 가까운 셀을 headerRange로 사용한다. validator contract를 v11로 올렸다.
- Files: `inference_data_ai_staged_draft_v2.py`, `tests/test_inference_data_ai_staged_draft_v2.py`, `HANDOFF.md`
- Verification: staged 테스트 32/32 통과. 실제 `ingest-run_123908ae40af87c7bd8e179b` v10 rejected part 19 records를 v11로 재구성해 전체 fragment 검증 통과.
- Next: v11/80KB corpus CLI를 재개한다.

## 2026-07-26 17:05 - 이전 fragment current-validator 승격
- Completed: validator contract가 바뀔 때 같은 원본·같은 owned-cell 집합의 이전 fragment를 현재 normalizer와 validator로 다시 검증해 통과한 경우 AI 재호출 없이 현재 part로 승격하는 재개 경로를 추가했다.
- Decisions: revisionUid/contentSha256와 disposition owned-cell 집합이 정확히 같아야 후보로 보고, current envelope scope·source coverage·numeric binding 전체 검증을 통과한 fragment만 승격한다. 새 AI 출력의 provenance identity는 계속 엄격히 검증한다.
- Files: `inference_data_ai_staged_runner_v2.py`, `tests/test_inference_data_ai_staged_draft_v2.py`, `HANDOFF.md`
- Verification: staged 테스트 33/33 통과. 실제 `ingest-run_123908ae40af87c7bd8e179b`에서 v11 part 6개(59, 19, 137, 4, 112, 4 records)가 current validator로 재사용 가능함을 확인했다.
- Next: 동일 v11 CLI를 새 프로세스로 재개해 검증된 checkpoint를 즉시 승격한다.

## 2026-07-26 17:12 - nullable collection canonical projection 보강
- Completed: 재사용된 유효 fragment의 `factorValues`, `observations`, `aggregateOfSeries`, `aggregateReplicateRanges`가 JSON null일 때 canonical projection에서 Python TypeError가 나던 경로를 빈 목록으로 보수 정규화했다.
- Decisions: list가 아닌 collection 값은 의미를 추론하지 않고 빈 목록으로만 투영한다. validator v11은 유지하고 projection 단계의 타입 내구성만 보강했다.
- Files: `inference_data_ai_staged_draft_v2.py`, `HANDOFF.md`
- Verification: staged 테스트 33/33 통과. 실제 08 MSM v11 current fragment 6개, merged records 334개를 canonical studies 10개·measurement series 8개로 projection 통과.
- Next: 동일 v11 CLI를 재개해 08 MSM import/verify와 나머지 corpus를 계속 처리한다.

## 2026-07-26 21:54 - 중단된 Goal 및 corpus 체크포인트 복구 조사
- Completed: 현재 세션에는 활성 Goal이 남아 있지 않음을 확인하고, 최신 `HANDOFF.md`와 실제 외부 corpus journal을 대조해 중단 지점을 복구했다.
- Decisions: 새 목표를 임의 생성하지 않는다. 최신 실행 계약은 승인된 361개 workbook을 v11 validator와 80,000-byte staged 임계값으로 완료한 뒤 canonical DB·AI 질의·WPF를 검증하는 것이다. journal의 `RUNNING`은 실제 writer가 없는 중단 흔적이므로 완료로 간주하지 않는다.
- Files: `HANDOFF.md`
- Verification: 관련 Python/Excel/dotnet writer 프로세스가 없음을 확인했다. 실제 `form-approved/corpus-journal.json`은 44 `COMPLETED`, 2 `FAILED`, 315 stale `RUNNING`, 25 `PENDING`이고 마지막 run `corpus-run_ec4ad4d049636126ac1de1ea`는 2026-07-26 17:06 +07 시작 후 종료 기록 없이 멈춰 있다. 최신 소스 체크포인트는 v11 nullable projection 보강까지 기록되어 있다.
- Next: 사용자가 Goal 재설정·계속 진행을 지시하면 동일 DB/입력/출력 경로에 `form-pipeline-complete --draft-monolithic-max-bytes 80000`을 재개하고, stale `RUNNING` 재조정과 08 MSM import/verify부터 추적한다.

## 2026-07-26 21:56 - 통합 Excel 분석 Goal 재설정
- Completed: 사용자의 재개 지시에 따라 승인 신규 Excel 분석 완료와 기존·신규 canonical 데이터 통합 AI 질의 검증을 활성 Goal로 다시 설정하고 실제 재개 경로를 확정했다.
- Decisions: 기존 데이터와 신규 데이터를 별도 답변 경로로 분리하지 않고 동일 `InputDataFinish.sqlite`의 current public analysis와 evidence query를 사용한다. corpus는 기존 체크포인트를 재사용하고 v11 validator·80,000-byte staged 임계값을 유지한다.
- Files: `HANDOFF.md`
- Verification: 실제 DB `PRAGMA quick_check`가 `ok`이고 관련 writer 프로세스가 0임을 재확인했다. `form-pipeline-complete`의 DB·입력 archive·출력 root와 CLI 옵션을 현재 코드 및 저장 결과에서 복원했다.
- Next: 동일 체크포인트로 전체 CLI를 시작하고 stale `RUNNING` 조정, 08 MSM import, 신규 실패 유형을 순서대로 처리한다.

## 2026-07-27 02:04 - 통합 corpus 재개 4시간 체크포인트
- Completed: v11 validator와 80,000-byte staged 임계값으로 실제 361개 승인 corpus를 재개해 완료 수를 44개에서 87개로 늘렸고, 실패 산출물 22개를 원인과 함께 보존했다.
- Decisions: 장기 실행을 종료하지 않는다. 호출 셸의 4시간 제한 이후에도 주 Python과 현재 Codex 자식이 정상 실행 중이므로 같은 journal에 경쟁 writer를 시작하지 않고 기존 PID 3148을 계속 감시한다.
- Files: 실제 `form-approved/corpus-journal.json`, 신규 workbook별 draft/import 산출물, `InputDataFinish.sqlite`, `HANDOFF.md`
- Verification: journal은 87 `COMPLETED`, 22 `FAILED`, 252 `RUNNING`, 25 `PENDING`이며 2026-07-27 01:33 +07까지 갱신됐다. DB `PRAGMA quick_check`는 `ok`; 재개 후 신규 public analysis가 저장됐고 현재 Codex 하위 작업 2개가 실행 중이다.
- Next: 현재 corpus pass가 끝날 때까지 감시한 뒤 22개 이상 보존된 실패를 유형별로 수정·집중 재시도한다.

## 2026-07-27 02:08 - 비표준 AI enum 보수 정규화
- Completed: 실제 실패 22건 중 9건을 차지한 비표준 `favorableDirection`·`isolationStatus` 출력이 전체 workbook을 탈락시키지 않도록 안전한 정규화를 추가했다.
- Decisions: 허용되지 않은 favorable direction은 결과 방향을 추론하지 않고 `UNKNOWN`, 허용되지 않은 factor isolation은 격리 여부를 추론하지 않고 `UNASSESSED`로 내린다. 승인·인과·효과 상태는 변경하지 않는다.
- Files: `inference_data_ai_semantic_ai.py`, `tests/test_inference_data_ai_semantic_ai.py`, `HANDOFF.md`
- Verification: 해당 정규화 단위 테스트 1/1과 Python 구문 검사가 통과했다. 현재 장기 실행은 수정 전 모듈을 로드했으므로 완료 후 재개 pass부터 적용된다.
- Next: 장기 pass를 계속 감시하면서 나머지 content/staged/conclusion 실패를 일반화 가능한 방식으로 보강한다.

## 2026-07-27 02:11 - 짧은 원본 판정 문구 증거 인정
- Completed: 실제 source-backed 판정 `Can use.`와 `- Follow standard:`가 정확한 셀·sourceText를 인용했는데도 세 단어 미만이라는 이유로 탈락하던 2건을 보강했다.
- Decisions: 반복 상태값 `PASSED` 같은 categorical 셀은 계속 SOURCE_CONCLUSION으로 거부한다. 반면 `Can use`, `Follow standard`, 기존 `Test more`처럼 명시적인 짧은 조치·판정 문구는 정확한 원본 셀과 일치할 때만 source conclusion으로 인정한다.
- Files: `inference_data_ai_study_import.py`, `tests/test_inference_data_ai_study_import.py`, `HANDOFF.md`
- Verification: 세 짧은 명시 판정은 통과하고 `PASSED`는 거부하는 집중 테스트 1/1 및 Python 구문 검사가 통과했다.
- Next: 현재 장기 pass 종료 후 새 enum·source conclusion 보강으로 실패 workbook을 체크포인트 재시도한다.

## 2026-07-27 02:16 - 병합 header placeholder 정규화
- Completed: series `headerRange`가 병합 헤더의 비-anchor 좌표 `G3`를 가리켜 실제 캡처 셀을 찾지 못하던 X626 Top/Bottom 2건을 병합 anchor 범위 `F3:G4`로 결정적으로 정규화했다.
- Decisions: 잘못된 header 범위가 동일 sheet의 실제 merge 범위 내부에 완전히 포함될 때만 가장 작은 포함 merge 범위로 교정한다. 임의의 인접 라벨 추론은 하지 않고 이후 owned/shared scope 검증을 그대로 통과해야 한다.
- Files: `inference_data_ai_staged_draft_v2.py`, `tests/test_inference_data_ai_staged_draft_v2.py`, `HANDOFF.md`
- Verification: 신규 병합 placeholder 테스트 1/1과 Python 구문 검사가 통과했다. 실제 rejected fragment 2개를 정규화해 둘 다 `G3`에서 `F3:G4`로 교정됨을 확인했다.
- Next: 현재 pass 종료 뒤 정규화된 두 fragment를 AI 재호출 없이 current validator로 승격하고 나머지 staged 실패를 처리한다.

## 2026-07-27 02:23 - 다중 Arm 행 series 보수 분할
- Completed: GRAPH 표의 F8:F10 한 series가 STD·Normal·Frame V4 세 원본 Arm 행을 함께 담아 `arm=null`로 탈락하던 사례를 각 원본 행·Arm별 단일 series로 결정적으로 분할했다.
- Decisions: RAW series이고 valueRange와 rowIdentityRange가 같은 행의 단일 열이며 모든 행 라벨이 같은 logical Study의 exact Arm 선언과 일치할 때만 분할한다. 다열·불일치·추정이 필요한 경우는 기존 fail-closed를 유지한다.
- Files: `inference_data_ai_staged_draft_v2.py`, `inference_data_ai_staged_runner_v2.py`, `tests/test_inference_data_ai_staged_draft_v2.py`, `HANDOFF.md`
- Verification: 신규 3-Arm 분할 테스트 1/1과 두 runtime 모듈 구문 검사가 통과했다. 실제 rejected fragment에서 `std/F8/C8`, `normal/F9/C9`, `frame_v4/F10/C10` 세 source-exact series가 생성됨을 확인했다.
- Next: 현재 pass 종료 후 current normalizer로 해당 fragment 전체 scope·numeric binding·projection을 재검증한다.

## 2026-07-27 02:29 - 숫자 Arm identity coverage 보강
- Completed: `1.2`, `2`, `3`처럼 숫자로 저장된 원본 Arm 행 ID가 측정값으로 요구되면서도 Arm 선언으로는 coverage가 인정되지 않아 탈락한 실제 fragment를 보강했다.
- Decisions: exact Arm label이 동일 evidence 셀의 표시 숫자와 경계까지 정확히 일치할 때만 `ARM_NUMERIC_IDENTITY`로 보존한다. 이를 sample size나 outcome 측정값으로 승격하지 않으며 final content coverage와 staged fragment validator에 같은 규칙을 적용한다.
- Files: `inference_data_ai_staged_draft_v2.py`, `inference_data_ai_content_coverage.py`, `tests/test_inference_data_ai_staged_draft_v2.py`, `tests/test_inference_data_ai_content_coverage.py`, `HANDOFF.md`
- Verification: staged exact numeric Arm과 최종 content coverage 집중 테스트 2/2 및 두 runtime 모듈 구문 검사가 통과했다.
- Next: 기존 `B6:B8` 누락 rejected fragment를 새 validator로 재사용하고 남은 numeric 누락 표를 처리한다.

## 2026-07-27 02:36 - 가로 replicate 번호 헤더 분류
- Completed: RAW DATA의 `1..N` 가로 sample 번호가 측정 결과로 오분류되어 I7:M7·R7:X7 등이 누락 처리되던 사례를 source 구조로 식별했다.
- Decisions: 같은 행에서 1부터 연속된 3개 이상 정수이고 인접 AVG/sample 계열 표식이 있으며, 아래 열에 실제 numeric data 또는 header를 참조하는 formula label이 3개 이상 있을 때만 `HORIZONTAL_REPLICATE_IDENTIFIER`로 제외한다. 아래 측정값은 계속 required result로 유지한다.
- Files: `inference_data_ai_content_coverage.py`, `tests/test_inference_data_ai_content_coverage.py`, `HANDOFF.md`
- Verification: 실측 열과 formula-only 빈 sample 열 fixture 2/2 및 구문 검사가 통과했다. 실제 X526 Bottom packet에서 문제 좌표 I7:M7·R7:X7 전부 replicate identity로 판정됨을 확인했다.
- Next: 현재 pass 종료 뒤 해당 rejected fragment를 current coverage로 재사용해 남은 numeric binding 오류가 없는지 확인한다.

## 2026-07-27 02:30 - source-stable 엔터티 alias 병합
- Completed: 같은 원본 셀에서 생성된 엔터티의 모델 key와 설명이 fragment마다 달라 병합이 중단되던 문제를 source-ordered canonical key로 통일하고, 관측·series·비교·Arm factor 참조와 source-derived recordId를 함께 다시 쓰도록 보강했다.
- Decisions: 동일 source-stable recordId의 key와 설명 필드는 첫 원본 순서 값을 유지하되, `entityType`, `unit`, 방향·상태·context kind 같은 핵심 의미 필드의 충돌과 다른 엔터티 key와의 alias 충돌은 계속 fail-closed 한다. 병합 의미 변경에 따라 consolidator contract를 v5로 올렸다.
- Files: `inference_data_ai_staged_draft_v2.py`, `tests/test_inference_data_ai_staged_draft_v2.py`, `HANDOFF.md`
- Verification: 병합 alias·도착 순서 독립성·진짜 단위 충돌 거부 집중 테스트 3/3 및 Python 구문 검사가 통과했다. 실제 실패 run `ingest-run_34e4c80c93659b2fad35f490`의 저장 fragment 5개를 current code로 병합해 57 records·1,745 dispositions가 오류 없이 생성됨을 확인했다.
- Next: 장기 corpus pass가 끝난 뒤 consolidator v5로 해당 workbook과 누적 실패를 재시도하고, 남은 content coverage 실패 유형을 보강한다.

## 2026-07-27 02:48 - 실제 content coverage 실패 6종 보강
- Completed: 숨김 legacy 표의 Excel COM 오류값·오류식 보조 입력, 인용부호가 든 날짜 서식, 번호가 건너뛴 `No.` 열, series 축·observation stratum 라벨, 넓은 표의 병합 field header, 오류 셀 위의 `OK` header를 source 구조에 맞게 분류했다. locator가 결론으로 지정했지만 모델이 conclusion object를 누락한 원문은 covering Study에 exact `SOURCE_CONCLUSION`으로 보존하도록 workflow를 보강했다.
- Decisions: 보이는 값과 정상 숨김 값은 계속 필수이며, 양옆에 숨김 Excel 오류식이 있는 raw helper만 제외한다. source conclusion은 inventory와 Study evidence가 모두 같은 셀을 입증할 때만 원문 그대로 추가하고 AI 문장을 합성하지 않는다.
- Files: `inference_data_ai_content_coverage.py`, `inference_data_ai_workflow.py`, `tests/test_inference_data_ai_content_coverage.py`, `HANDOFF.md`
- Verification: 신규·회귀 집중 테스트 16/16 및 Python 구문 검사가 통과했다. 실제 실패 run 6개(101, 104, 105, 109, 112, 12)의 저장 packet/manifest를 current code로 재검증해 quantitative·semantic·categorical·narrative 누락이 모두 0임을 확인했다.
- Next: 진행 중 corpus pass 종료 후 current workflow로 누적 실패를 재시도하고 새로 드러난 staged numeric binding 실패를 분석한다.

## 2026-07-27 03:01 - GRAPH formula helper numeric 분류
- Completed: GRAPH 표에서 결과가 아닌 THD formula label 입력(200/400/1000), `MATCH`가 반환한 `INDEX` 선택자, 표 배치용 수열 A8:A13이 quantitative result로 요구되던 두 workbook을 구조적으로 분류했다.
- Decisions: 숫자가 바로 아래 formula label에 실제로 렌더링될 때만 `FORMULA_LABEL_INPUT`, numeric `MATCH` 결과를 2개 이상 `INDEX` 식이 소비할 때만 `FORMULA_LOOKUP_INDEX`, 앞쪽 열의 5행 이상 연속 수열에 동일 증가식 2개 이상과 인접 row label 3개 이상이 있을 때만 `FORMULA_LAYOUT_SEQUENCE`로 제외한다. 실제 측정 formula는 계속 required이다.
- Files: `inference_data_ai_content_coverage.py`, `tests/test_inference_data_ai_content_coverage.py`, `HANDOFF.md`
- Verification: 신규 helper 및 기존 구조 분류 집중 테스트 3/3과 구문 검사가 통과했다. 실제 rejected fragment `e2ba...`는 58 records·131 dispositions, `8d0d...`는 74 records·133 dispositions로 current validator를 통과했다.
- Next: 현재 장기 pass가 끝난 뒤 current contract로 전체 실패를 재시도한다.

## 2026-07-27 03:06 - 실행 중 신규 결론 누락 재검증
- Completed: 장기 corpus pass에서 새로 기록된 `Test!B51` source conclusion 누락 1건의 저장 packet과 repair manifest를 현재 content-coverage 코드로 다시 검증해 통과시켰다.
- Decisions: locator가 결론으로 확정한 원문 셀은 현재 `augment_exact_source_conclusions`가 가장 구체적인 covering Study에 exact `SOURCE_CONCLUSION`으로 보충한다. 실행 중인 PID 3148은 수정 전 모듈을 로드했으므로 종료 후 current-code resume pass에서 회수한다.
- Files: `HANDOFF.md`
- Verification: 실제 `ingest-run_554bfe5922d0d10e6bbea5dd`의 3개 chunk·3개 locator와 `canonical-study-manifest.repair-rejected.json`으로 inventory를 재생성했다. 결론 수가 4→5로 보충됐고 complete coverage의 missing 항목은 0이었다.
- Next: PID 3148 종료를 기다리는 동안 기존·신규 Study 통합 질의를 실제 DB로 검증하고, 종료 후 current-code corpus resume pass를 시작한다.

## 2026-07-27 03:17 - 기존·신규 canonical 통합 AI 문의
- Completed: WPF 기본 ‘최신 적재 DB’가 기존/신규를 구분하지 않고 동일 canonical SQLite의 모든 non-STALE Study를 검색함을 확인하고, 외부 운영 output root에 답변을 저장하지 못하던 경로 제한을 DB 범위 안전 경로로 수정했다. 비교 승인 전 자료도 원본 근거가 있는 관측값은 설명용 결과로 함께 렌더링하고, Study 범위 결과는 비교 레코드별로 중복하지 않도록 보강했다.
- Decisions: 정량 효과의 검증 gate는 유지한다. 직접 Observation EVD뿐 아니라 같은 current source revision의 verified Outcome 범위 EVD도 legacy 관측값의 설명용 근거로 허용한다. broad outcome이 파일/analysis/Study 제목에서 매치된 경우 상세 submetric을 숨기지 않는다. 외부 출력은 DB가 속한 알려진 output tree 안만 허용하고 그 밖의 절대 경로는 계속 차단한다.
- Files: `inference_data_ai_cli.py`, `inference_data_ai_query.py`, `inference_data_ai_answer.py`, `tests/test_inference_data_ai_cli.py`, `tests/test_inference_data_ai_query.py`, `tests/test_inference_data_ai_answer.py`, `HANDOFF.md`
- Verification: 관련 query/answer/CLI 집중 테스트 11건과 세 Python 모듈 구문 검사가 통과했다. 운영 DB 질문에서 기존 `DATA-0ADC21566B72`와 신규 `DATA-EF7A58846A92`가 함께 검색·표시됐고, 기존 1 outcome/2 observations와 신규 18 outcomes/54 observations가 149 verified EVD로 연결됐으며 uncited observation은 0이었다. WPF와 동일한 `evidence-answer --db` 호출이 외부 운영 output root에 JSON/Markdown을 정상 저장했다.
- Next: 장기 PID 3148에서 새로 발생한 실패 3건을 current code로 재검증·보강하고, pass 종료 후 current-code resume로 전체 승인 corpus를 terminal 완료한다.

## 2026-07-27 03:29 - 다중 청크 결과표 결정론 복구
- Completed: 같은 시트의 인접 packet chunk 2개가 한 draft part에 포함될 때도 하나의 결과표로 coalesce하고, 겹치는 Study 경계에서는 더 작은 구체적 표 영역을 우선하도록 보강했다. 실제 실패 run 2건의 필수 결과 수치를 모두 source-bound Observation으로 재구성했으며, `TEST data!C1`의 단독 페이지 번호는 결과값이 아닌 구조값으로 분리했다.
- Decisions: 동일 sheet/revision의 중복 없는 chunk만 병합한다. Study 할당 거리 동률에서는 anchor bounding-box 면적이 작은 Study를 선택하고, 필수 수치가 완전히 같은 경계에 걸리면 계속 fail-closed한다. `SHEET_LAYOUT_ORDINAL`은 row 1의 유일한 1~99 정수가 row 2의 merged `TITLE` band 바로 위에 있을 때만 적용한다.
- Files: `inference_data_ai_staged_draft_v2.py`, `inference_data_ai_content_coverage.py`, `tests/test_inference_data_ai_staged_draft_v2.py`, `tests/test_inference_data_ai_content_coverage.py`, `HANDOFF.md`
- Verification: 신규·기존 집중 테스트 4/4 통과. 실제 `ingest-run_967dd43346753616d16e6bc0` part `4f790...`은 484 records/552 dispositions로 복구되어 E18:N18 필수 수치가 Observation에 연결됐고 C1은 관측값에서 제외됐다. `ingest-run_3d6164f057f153a66887b7f0` part `85a178...`은 243 records/309 dispositions로 복구되어 G18:J20 필수 수치가 구체적 cushion Study에 연결됐다.
- Next: Ethanol run `ingest-run_2981a7704b0210dd30578739`의 semantic label 5개를 조사·보강한 뒤 PID 3148 종료 시 current-code corpus resume pass를 실행한다.

## 2026-07-27 03:32 - Ethanol 의미 라벨 coverage 복구
- Completed: Ethanol 96% 보고서의 누락 의미 라벨 5개와 결론 B65를 current code로 모두 보존했다. 정확한 Study 제목 C34/C56은 Study title evidence로, decap diagram의 반복 `Normal` F27/L27/R27은 해당 Arm label과 좁은 Study evidence가 함께 있을 때 semantic coverage로 인정했다.
- Decisions: Study 제목은 sourceText가 정확히 일치하는 Study evidence가 해당 셀을 직접 포함할 때만 인정한다. 반복 Arm 라벨은 동일 label/condition의 Arm이 존재하고, 라벨 문구를 포함하며 면적 256셀 이하인 Study evidence가 해당 셀을 포함할 때만 인정한다. 넓은 포괄 범위는 fail-closed한다.
- Files: `inference_data_ai_content_coverage.py`, `tests/test_inference_data_ai_content_coverage.py`, `HANDOFF.md`
- Verification: 신규·기존 semantic/ordinal 집중 테스트 3/3 통과. 실제 `ingest-run_2981a7704b0210dd30578739` repair manifest를 exact source-conclusion augmentation 후 재검증하여 required numeric 174/174, semantic labels 39/39, narrative conclusions 7/7, missing 0을 확인했다.
- Next: 장기 PID 3148이 종료될 때까지 새 실패를 계속 current code로 재검증하고, 종료 즉시 같은 corpus 명령을 current-code resume로 재실행한다.

## 2026-07-27 03:41 - 결과 섹션 라벨과 source-identical entity 병합
- Completed: `1. Check Process NG rate`처럼 Process와 NG rate를 함께 포함한 결과 섹션 제목이 factor로 오분류되던 문제를 바로잡고, 서로 다른 fragment가 동일 source cell의 같은 Arm을 `old_jig`/`sample_ng_old_jig`처럼 달리 명명해도 한 canonical entity로 병합하도록 보강했다.
- Decisions: outcome 문구가 있고 아래 5행·오른쪽 24열 안에 실제 결과값이 2개 이상이면 factor보다 outcome label을 우선한다. entity는 같은 Study/type과 완전히 같은 source identity cell 집합을 공유하고 key 및 exact label이 서로 포함 관계일 때만 alias로 병합하며, 호환되지 않는 동일-cell entity는 분리 유지한다.
- Files: `inference_data_ai_content_coverage.py`, `inference_data_ai_staged_draft_v2.py`, `tests/test_inference_data_ai_content_coverage.py`, `tests/test_inference_data_ai_staged_draft_v2.py`, `HANDOFF.md`
- Verification: 관련 content coverage 테스트 3/3과 merge 테스트 3/3 통과. 실제 `ingest-run_9d0e3ea5854e470a3adcddc5`는 numeric 95/95·semantic 50/50으로 missing 0, 실제 `ingest-run_0fbfc6787b92ae1489f6c5fe`는 4 fragments를 33 records/1,465 dispositions로 병합하고 3-Study canonical manifest projection까지 통과했다.
- Next: `ingest-run_0c0e39dbce1c19a31f9f2590`의 CI3 및 N7:AB7 SPL 실패를 current deterministic projector/coverage로 복구한다.

## 2026-07-27 03:43 - 빈 formula label SPL 구조 분리
- Completed: SPL raw-data 표의 `CI3 = D7`이 빈 source label을 0으로 캐시한 구조 placeholder인데 결과값으로 요구되던 문제를 해결했다. 이미 보강된 N7 formula-label input 및 S7:AB7 replicate identifier 분류와 합쳐 해당 실패 part의 가짜 필수 수치를 모두 제거했다.
- Decisions: visible 0 formula도 단일 source 참조가 비어 있고, 같은 열 바로 아래 3행 안에 Excel error가 2개 이상이며 4~6행 아래에 AVG label이 있을 때만 `HIDDEN_FORMULA_WITHOUT_SOURCE_INPUT`으로 분류한다. error/AVG 구조가 없으면 계속 required result로 유지한다.
- Files: `inference_data_ai_content_coverage.py`, `tests/test_inference_data_ai_content_coverage.py`, `HANDOFF.md`
- Verification: 신규 및 관련 구조 분류 테스트 3/3 통과. 실제 `ingest-run_0c0e39dbce1c19a31f9f2590` part `ef156...`에서 CI3, N7, S7:AB7 required target이 0이 되었고, 기존 rejected fragment 14 records/365 dispositions가 current `validate_fragment_v2`를 통과했다.
- Next: PID 3148에서 이어서 발생하는 신규 실패를 조사하면서 current-code resume 준비를 유지한다.

## 2026-07-27 03:46 - current-code corpus resume 시작
- Completed: 2026-07-26 22:03에 시작해 구버전 모듈을 계속 사용하던 PID 3148 및 하위 AI 작업을 정확한 process tree로 종료하고, SQLite `quick_check=ok` 확인 후 현재 보강 코드로 corpus resume를 시작했다.
- Decisions: 구버전 pass는 새 validator/projector 수정이 적용되지 않아 남은 220건을 계속 처리하는 대신 안전하게 교체했다. current pass는 기존과 같은 DB/input/output 및 16 AI workers를 사용하고 reviewer를 `Codex Current Code Resume 2026-07-27`로 구분했다.
- Files: `HANDOFF.md`; runtime logs `D:\000. MyWorks\002. DB\InferenceDataAIService\incremental-com-corpus\form-approved\current-code-resume.stdout.log`, `current-code-resume.stderr.log`
- Verification: 교체 전 focused tests 8/8 및 관련 모듈 py_compile 통과, PID 3148 종료 후 DB quick_check 통과. 새 PID 24204가 `corpus-run_3ce643027ead68c22aca4fbb`로 실행 중이며 초기 조정 후 COMPLETED 102, RUNNING 259, PENDING 25(미선택) 상태다.
- Next: PID 24204를 terminal까지 모니터링하고 current-code 신규 실패가 생기면 즉시 재현·보강한다.

## 2026-07-27 03:46 - staged/coverage 전체 회귀 검증
- Completed: 이번 pass에서 변경된 staged draft 병합·결과표 projector·content coverage의 전체 모듈 회귀 테스트를 실행했다.
- Decisions: corpus와 동시에 실행하되 DB를 건드리지 않는 unit test만 사용했다.
- Files: `HANDOFF.md`
- Verification: `tests.test_inference_data_ai_staged_draft_v2`와 `tests.test_inference_data_ai_content_coverage` 전체 117/117 통과.
- Next: corpus 신규 terminal 결과를 감시하면서 semantic import 및 통합 query/answer 모듈 회귀 테스트를 실행한다.

## 2026-07-27 03:47 - canonical import 및 통합 문의 전체 회귀 검증
- Completed: enum normalization, Study import, 기존+신규 canonical 검색, 근거 답변 생성, CLI/WPF 연계 경로의 전체 관련 모듈 테스트를 실행했다.
- Decisions: corpus DB와 충돌하지 않도록 fixture/temp DB를 사용하는 test suite만 실행했다.
- Files: `HANDOFF.md`
- Verification: `semantic_ai`, `study_import`, `query`, `answer`, `cli` 5개 모듈 전체 173/173 통과. corpus는 PID 24204에서 COMPLETED 102/RUNNING 259/PENDING 25 상태로 정상 진행 중이다.
- Next: current corpus 첫 완료/실패 상태 변화를 모니터링한다.

## 2026-07-27 04:01 - 직접 비율식 canonical 투영 보강
- Completed: staged fragment가 분자만 남긴 비율 관측치를 canonical 계약에서 거부하던 문제를 보강했다. 전역 선택 chunk에서 직접 참조식 `numerator/denominator` 또는 `numerator/denominator*1000000`과 양쪽 셀의 정확한 캡처값이 일치할 때만 count pair·sampleSize·근거를 복원하고, 복원이 불가능한 불완전 쌍은 둘 다 비운다. 퍼센트 서식 관측치는 정확한 원시값의 100배 및 화면 문자열로 정규화한다.
- Decisions: 수식이 직접 두 셀을 가리키고 캐시값까지 산술 일치하는 경우에만 복원하며, 간접 계산·복수 후보·불일치 값에는 적용하지 않는다. 모델이 근거 없는 분모를 제안하는 방식은 사용하지 않는다.
- Files: `inference_data_ai_staged_draft_v2.py`, `tests/test_inference_data_ai_staged_draft_v2.py`, `HANDOFF.md`
- Verification: 신규 focused tests 2/2 및 staged 모듈 전체 41/41, py_compile 통과. 실제 `ingest-run_32752d0afc88177b36be8bfe`는 16 studies/57 observations, `ingest-run_08167157d6fc0653f4faac20`는 10 studies/229 observations로 canonical 계약·DB evidence·numeric·factor/arm·comparison·conclusion·complete content coverage를 모두 통과했다.
- Next: PID 24204가 추가로 드러낸 `result_c5_value` source-identity 충돌 두 건, 비수치 measurementSeries 한 건, 근거와 다른 `valueNumber=89.06` 한 건을 current artifacts에서 재현·보강한 뒤 최신 코드로 corpus resume한다.

## 2026-07-27 04:20 - 혼합 시계열 및 직접 비율식 coverage 보강
- Completed: 숫자 시계열 중간의 원문 상태 문자열을 별도 categorical observation으로 보존하고 숫자 구간만 measurement series로 분리했으며, `=+I41/G41` 같은 직접 비율식의 분자·분모가 단순 표시 라벨 입력으로 잘못 제외되던 content coverage 판정을 수정했다.
- Decisions: 직접 두 A1 셀을 나누는 수식과 선택적 `*100`/`*1000000`만 비율식으로 인정한다. 문자열 연결식 등 기존 formula-label 구조 판정은 유지한다.
- Files: `inference_data_ai_staged_draft_v2.py`, `tests/test_inference_data_ai_staged_draft_v2.py`, `inference_data_ai_content_coverage.py`, `tests/test_inference_data_ai_content_coverage.py`, `HANDOFF.md`
- Verification: 직접 비율식 및 기존 문자열 formula-label focused tests 2/2 통과. 실제 `ingest-run_ba0099af89b8b35351a38172`를 최신 projector로 재검증해 3 studies/231 observations/12 numeric series, required 251/251, semantic label 44/44, binding error 0으로 complete content coverage를 통과했다.
- Next: `ingest-run_8029314f7be3750302afbbcb`의 source evidence와 불일치한 `valueNumber=89.06`을 재현하고 정확한 source-backed 정규화를 보강한다.

## 2026-07-27 04:34 - 복합 텍스트 비율 및 의미 근거 완전성 보강
- Completed: 한 셀에 여러 비율이 있는 `NG function 89.06% (Gauss NG 71.88%)` 형식은 outcome 원문 라벨 바로 뒤의 비율만 숫자 근거로 인정하고, 단순 percent→PPM 중복 변환값은 count pair가 없을 때 제거했다. 리터럴 `Normal` Arm은 REFERENCE로 정규화하고, `No` 열의 소수형 행 번호를 구조 식별자로 제외했다. Projector가 registry의 정확한 source-cell 소유권을 이용해 누락된 factor/arm/outcome/factor-level/unit-quantity 의미 엔터티를 원문 근거와 함께 보존하도록 했다.
- Decisions: 복합 텍스트의 모든 숫자를 허용하지 않고 exact outcome label 또는 끝의 `percentage/percent/pct/rate`만 제거한 source-style label과 바로 결합된 `%` 값만 허용한다. 의미 엔터티는 AI 추론이 아니라 inventory 역할과 registry 소유권으로만 추가한다.
- Files: `inference_data_ai_study_import.py`, `tests/test_inference_data_ai_study_import.py`, `inference_data_ai_staged_draft_v2.py`, `tests/test_inference_data_ai_staged_draft_v2.py`, `inference_data_ai_content_coverage.py`, `tests/test_inference_data_ai_content_coverage.py`, `HANDOFF.md`
- Verification: 관련 import/staged/content focused tests 8/8 및 py_compile 통과. 실제 `ingest-run_8029314f7be3750302afbbcb`는 canonical 계약·DB evidence·numeric·factor/arm·comparison·conclusion·complete content coverage를 모두 통과했으며 18 studies/357 observations/28 series, required 424/424, semantic 88/88, binding error 0이다.
- Next: staged/content/import 전체 회귀 테스트를 실행하고, 비정상 종료되어 RUNNING으로 남은 corpus run을 SQLite 무결성 확인 후 최신 코드로 재개한다.

## 2026-07-27 04:35 - 의미 투영 및 import 전체 회귀 검증
- Completed: 이번 보강이 staged projection, complete content coverage, canonical Study import의 기존 동작을 훼손하지 않는지 전체 모듈 단위 테스트로 확인했다.
- Decisions: 운영 corpus writer가 없는 상태에서 fixture/temp DB만 사용하는 세 모듈을 병렬로 검증하고, 재개 전에 운영 SQLite 무결성도 별도로 확인했다.
- Files: `HANDOFF.md`
- Verification: `tests.test_inference_data_ai_staged_draft_v2` + `tests.test_inference_data_ai_content_coverage` 126/126, `tests.test_inference_data_ai_study_import` 53/53 통과. 운영 DB `PRAGMA quick_check` 결과 `ok`.
- Next: stale RUNNING corpus journal을 동일한 승인 corpus 명령으로 resume하여 최신 코드에서 모든 선택 파일을 terminal 상태로 완료한다.

## 2026-07-27 04:38 - 최신 코드 corpus resume 4 시작
- Completed: 구버전 모듈을 메모리에 유지한 채 실행 중이던 PID 24204와 정확한 하위 process tree 10개를 종료하고, SQLite `quick_check=ok` 확인 후 최신 코드로 승인 corpus resume를 시작했다.
- Decisions: PID 24204의 command line이 대상 `form-pipeline-complete`·운영 DB·Excel archive와 모두 일치함을 확인한 뒤에만 종료했다. 일시적 DB lock으로 종료된 resume 2/3은 writer가 아니라 기존 PID를 `python.exe`로만 필터링한 점이 원인이었고, 실제 프로세스명 `python3.11.exe`까지 확인해 중복 writer 없이 재개했다.
- Files: `HANDOFF.md`; runtime logs `D:\000. MyWorks\002. DB\InferenceDataAIService\incremental-com-corpus\form-approved\current-code-resume-4.stdout.log`, `current-code-resume-4.stderr.log`, `current-code-resume-4.pid`
- Verification: 새 PID 30792가 `corpus-run_f73c7552b8e7319c6facf26c`로 실행 중이다. 초기 journal 상태는 COMPLETED 104, RUNNING 59, INTERRUPTED 198, PENDING 25이며 재시도 대상 capture/packet/locator가 진행 중이다.
- Next: PID 30792를 모니터링하며 최신 코드의 신규 실패를 즉시 재현·보강하고, 선택 361건을 COMPLETED/FAILED terminal 상태로 만든다.

## 2026-07-27 04:45 - 반복 section 동일 entity key 병합 보강 및 resume 5
- Completed: 서로 다른 source section이 동일한 canonical entity key를 반복 선언하면 기존 merge가 source identity 차이만으로 거부하던 문제를 수정했다. 동일 Study/type/key 선언은 하나의 canonical record ID로 모으고, 이후 `_merge_entity_payload`가 실제 의미 payload 충돌을 계속 fail-closed 검사한다.
- Decisions: key가 같다는 이유로 payload 충돌을 허용하지 않으며, record ID와 disposition alias만 먼저 통합한 뒤 기존 semantic compatibility gate를 유지한다.
- Files: `inference_data_ai_staged_draft_v2.py`, `tests/test_inference_data_ai_staged_draft_v2.py`, `HANDOFF.md`; runtime logs `current-code-resume-5.stdout.log`, `current-code-resume-5.stderr.log`, `current-code-resume-5.pid`
- Verification: 신규/기존 entity merge focused tests 3/3 통과. 실제 `ingest-run_742fef70853326037e08939f` 17 fragments/1730 records와 `ingest-run_c33141110269f4f7c8763228` 17 fragments/1723 records를 최신 merger로 재현해 모두 통과했다. PID 30792와 하위 tree 종료 후 DB `quick_check=ok`; 최신 코드 PID 37268로 resume 5 시작.
- Next: resume 5에서 두 Reliability 파일이 DRAFT/IMPORT/VERIFY를 통과하는지 확인하고 새 실패 유형을 계속 보강한다.

## 2026-07-27 04:52 - coverage inventory 재사용 최적화 및 resume 6
- Completed: staged canonical projection의 의미 엔터티 보강과 직후 complete-content validation이 동일 workbook coverage inventory를 두 번 계산하던 경로를 하나의 사전 계산 결과 재사용으로 변경했다.
- Decisions: 검증 계약이나 inventory 내용은 바꾸지 않고, staged 최종 merge 경로에만 optional precomputed inventory를 전달한다. 다른 호출자는 기존처럼 projector/validator가 자체 계산하는 호환 경로를 유지한다.
- Files: `inference_data_ai_staged_draft_v2.py`, `inference_data_ai_workflow.py`, `HANDOFF.md`; runtime logs `current-code-resume-6.stdout.log`, `current-code-resume-6.stderr.log`, `current-code-resume-6.pid`
- Verification: staged 전체 47/47, workflow 전체 13/13 통과. resume 5에서는 새 entity 오류 없이 DRAFT 완료가 발생했으며 신규 실패 0건이었다. PID 37268 종료 후 DB `quick_check=ok`; 최적화 코드 PID 36448로 resume 6 시작.
- Next: resume 6 처리 속도와 Reliability/복합비율 실제 성공을 확인하고 terminal 완료까지 모니터링한다.

## 2026-07-27 05:01 - resume 6 첫 end-to-end 완료 확인
- Completed: 최적화된 resume 6에서 `ingest-run_fb3674cfb59ef4cb53f79a12`가 DRAFT→IMPORT→VERIFY 전 단계를 통과해 corpus COMPLETED가 105에서 106으로 증가했다.
- Decisions: writer는 단일 PID 36448로 유지하며, DB import/verify 직렬 구간이 길어도 stage timestamp와 CPU가 계속 전진하는 동안 중단하지 않는다.
- Files: `HANDOFF.md`; runtime journal/logs only
- Verification: corpus 상태 COMPLETED 106, RUNNING 255, PENDING 25, FAILED 0. 실제 회귀 대상 `ingest-run_ba0099af89b8b35351a38172`와 `ingest-run_08167157d6fc0653f4faac20`도 최신 코드에서 DRAFT를 통과하고 IMPORT 단계에 진입했다.
- Next: 두 Reliability 런과 복합 비율 런의 VERIFY 완료 및 새 실패 이벤트를 계속 감시한다.

## 2026-07-27 05:08 - 신규 corpus 및 통합 문의 전체 회귀 검증
- Completed: staged projection, complete content coverage, canonical Study import, workflow, semantic AI, 기존+신규 canonical query, 근거 답변, CLI 연계의 관련 전체 회귀 테스트를 실행했다.
- Decisions: 운영 corpus writer와 충돌하지 않도록 fixture/temp DB를 사용하는 `unittest` 모듈만 검증했으며, 설치되지 않은 `pytest` 대신 프로젝트 기존 검증 방식인 표준 `unittest`를 사용했다.
- Files: `HANDOFF.md`
- Verification: 관련 8개 모듈 전체 314/314 통과. corpus PID 36448은 CPU와 로그 timestamp가 계속 전진하며 신규 실패 이벤트 0건으로 실행 중이다.
- Next: 승인된 361개 workbook이 terminal 상태가 될 때까지 resume 6을 감시하고, 실패가 생기면 해당 실제 artifact를 재현·수정한다.

## 2026-07-27 05:18 - source-backed series header 및 INDEX selector 보강
- Completed: `ingest-run_236cecf74beac6daa70b5af2`의 fragment가 캡처되지 않은 `N8`/`R8:AA8` 등을 series header로 제안해 실패한 문제를 수정했다. 유효 header가 없으면 정확히 참조된 aggregate series header, 단일 열의 cited text header, 또는 focused fragment에 실제 포함된 동일 폭 context row만 사용한다. `INDEX(..., A53)`가 `STD_AVG`/`STD #1` 같은 행 라벨을 만드는 숫자 selector도 구조 입력으로 분류했다.
- Decisions: fragment 범위 밖 셀이나 추정 라벨은 허용하지 않고, 현재 focused chunk에서 source-backed로 검증되는 셀만 대체 header로 사용한다. INDEX selector 제외는 같은 행 오른쪽의 INDEX 수식이 해당 셀을 참조하고 캐시 결과가 텍스트 라벨일 때만 적용한다.
- Files: `inference_data_ai_staged_draft_v2.py`, `inference_data_ai_content_coverage.py`, `tests/test_inference_data_ai_staged_draft_v2.py`, `tests/test_inference_data_ai_content_coverage.py`, `HANDOFF.md`
- Verification: 신규 focused tests 3/3, staged+content 전체 130/130 통과. 실제 rejected fragment는 header를 `B53`, `R4:AA4`, `AL3`, `AN4:AW4`, `BH3` 등 정확한 focused source 범위로 정규화한 뒤 15 records/375 dispositions 전체 `validate_fragment_v2` 통과.
- Next: PID 36448과 정확한 하위 tree를 종료하고 SQLite 무결성 확인 후 수정 코드를 resume하여 해당 workbook의 전체 DRAFT/IMPORT/VERIFY를 확인한다.

## 2026-07-27 05:21 - 최신 코드 corpus resume 7 시작
- Completed: source-backed header/INDEX selector 수정 전 모듈을 유지하던 PID 36448과 정확한 하위 process tree 11개를 종료하고 최신 코드로 corpus resume를 시작했다.
- Decisions: 대상 PID command line이 resume 6 reviewer·운영 DB·Excel archive와 모두 일치함을 확인한 뒤에만 종료했다. 새 writer는 기존과 동일한 16 workers/low 설정을 유지한다.
- Files: `HANDOFF.md`; runtime logs `D:\000. MyWorks\002. DB\InferenceDataAIService\incremental-com-corpus\form-approved\current-code-resume-7.stdout.log`, `current-code-resume-7.stderr.log`, `current-code-resume-7.pid`
- Verification: 종료 대상 PID 11개가 모두 사라졌고 운영 DB `PRAGMA quick_check` 결과 `ok`. 새 PID 18260, corpus `corpus-run_dce16d317d238178549e21e8`가 COMPLETED 107/RUNNING 254/PENDING 25, 신규 실패 0건으로 실행 중이다.
- Next: 실제 실패 대상 `ingest-run_236cecf74beac6daa70b5af2`와 Reliability/복합비율/반복-section 대상들의 DRAFT→IMPORT→VERIFY 완료를 확인한다.

## 2026-07-27 05:24 - resume 7 첫 end-to-end 완료
- Completed: 최신 source-backed header 및 INDEX selector 코드에서 `ingest-run_59b53f04ac254aa99063b78c`가 DRAFT→IMPORT→VERIFY를 모두 통과했다.
- Decisions: 새 실패가 없고 CPU·stage timestamp가 전진하므로 PID 18260을 유지한다.
- Files: `HANDOFF.md`; runtime journal/logs only
- Verification: corpus COMPLETED 108/RUNNING 253/PENDING 25, resume 7 실패 0건. 완료 workflow 상태는 `NEEDS_REVIEW`이며 DB 반영과 VERIFY까지 성공했다.
- Next: 수정 대상 `ingest-run_236cecf74beac6daa70b5af2` 및 지정 회귀 대상들이 terminal 완료될 때까지 계속 감시한다.

## 2026-07-27 05:47 - workflow 중복 source-claim 검증 제거
- Completed: DRAFT에서 같은 in-memory manifest와 resolved revision에 대해 이미 통과한 DB-backed source-claim 검증을 IMPORT에서 다시 수행하지 않도록 내부 전용 `source_claims_prevalidated` 경로를 추가했다.
- Decisions: 직접 호출의 기본값은 `False`로 유지해 기존 fail-closed 동작을 보존했다. workflow에서만 `True`를 전달하며, 이 경우에도 revision 해석과 manifest schema 검증은 다시 수행한다.
- Files: `inference_data_ai_study_import.py`, `inference_data_ai_workflow.py`, `tests/test_inference_data_ai_study_import.py`, `tests/test_inference_data_ai_workflow.py`, `HANDOFF.md`
- Verification: 신규 focused tests 2/2 통과, study import+workflow 관련 모듈 전체 67/67 통과.
- Next: 실행 중인 resume 7을 정확한 process tree 단위로 종료하고 DB quick check 후 최신 코드로 resume 8을 시작해 실제 IMPORT 시간 단축과 corpus 무실패 진행을 확인한다.

## 2026-07-27 05:49 - 최적화 코드 corpus resume 8 시작
- Completed: command line이 resume 7 운영 작업과 일치하는 PID 18260 및 순간적으로 추가된 하위 프로세스를 포함한 정확한 process tree 14개를 종료하고 최신 코드로 corpus resume 8을 시작했다.
- Decisions: AI 요청 하위 프로세스를 포함한 확인된 tree만 종료했으며, 기존 16 workers/low 설정과 동일한 대상 corpus를 유지했다.
- Files: `HANDOFF.md`; runtime logs `D:\000. MyWorks\002. DB\InferenceDataAIService\incremental-com-corpus\form-approved\current-code-resume-8.stdout.log`, `current-code-resume-8.stderr.log`
- Verification: 종료 대상 14개가 모두 사라졌고 운영 DB `PRAGMA quick_check` 결과 `ok`. 새 PID 37472, corpus `corpus-run_24cd001bce2c865940404b56`이 시작됐고 이전 run은 `INTERRUPTED`로 정리됐다. 시작 직후 COMPLETED 109, 신규 실패 0건.
- Next: resume 8에서 첫 대형 manifest의 DRAFT→IMPORT 시간을 측정해 중복 검증 제거 효과를 확인하고, 실패 없이 전체 361개 선택 workbook이 terminal 완료될 때까지 감시한다.

## 2026-07-27 06:07 - per-workbook 3.62GB 전체 DB 검증 제거
- Completed: 각 Excel의 VERIFY마다 `PRAGMA foreign_key_check`, canonical invariant scan, 20개 테이블 전체 COUNT를 실행하던 병목을 분석 범위 검증으로 교체했다. 전체 DB 무결성 검사는 corpus 종료 후 1회만 수행하도록 옮겼다.
- Decisions: IMPORT 연결의 `PRAGMA foreign_keys=ON`이 행 단위 FK를 fail-closed로 보장하므로, 파일별 VERIFY는 방금 반영한 `workbook_analysis_id`의 존재·aggregation invariant·scoped counts를 검사한다. 전체 DB scan 안전망은 제거하지 않고 corpus terminal 경계에 유지한다.
- Files: `inference_data_ai_schema.py`, `inference_data_ai_workflow.py`, `inference_data_ai_corpus_workflow.py`, `tests/test_inference_data_ai_study_import.py`, `tests/test_inference_data_ai_corpus_workflow.py`, `HANDOFF.md`
- Verification: focused 3/3 통과, schema+study import+workflow+corpus 전체 85/85 통과. 운영 DB의 실제 15 studies/181 observations 분석 ID 210에 scoped 검증을 실행해 0.012초, `ok=True`, 정확한 entity counts를 확인했다.
- Next: 현재 구코드 resume 8 process tree를 종료하고 DB quick check 후 최신 VERIFY 최적화 코드로 resume 9를 시작해 첫 IMPORT→VERIFY가 초 단위로 종료되는지 확인한다.

## 2026-07-27 06:09 - scoped VERIFY corpus resume 9 시작
- Completed: resume 8 PID 37472와 그 하위 AI process를 포함한 확인된 tree 11개를 종료하고 scoped VERIFY 코드로 corpus resume 9를 시작했다.
- Decisions: resume 8에서 IMPORT 완료된 두 분석은 DB에 보존해 새 run이 canonical reconciliation하도록 했고, 정확한 command line과 process ancestry가 확인된 프로세스만 종료했다.
- Files: `HANDOFF.md`; runtime logs `D:\000. MyWorks\002. DB\InferenceDataAIService\incremental-com-corpus\form-approved\current-code-resume-9.stdout.log`, `current-code-resume-9.stderr.log`
- Verification: 종료 대상 11개가 모두 사라졌고 DB `PRAGMA quick_check`는 `ok`. 새 PID 4792, corpus `corpus-run_586ecee2a650832f8a07056a`가 시작됐으며 COMPLETED가 109에서 111로 증가했고 신규 실패 0건이다.
- Next: resume 9의 첫 IMPORT 완료 직후 VERIFY elapsed를 측정해 per-workbook 전체 DB scan 제거가 실제 운영 pipeline에서도 초 단위로 적용됐는지 확인한다.

## 2026-07-27 06:46 - 중단된 resume 9 복구 및 scoped VERIFY 운영 확인
- Completed: 중단된 세션의 HANDOFF·journal·runtime log·process 상태를 대조해 resume 9가 더 이상 실행 중이 아니며 stale `RUNNING`으로 남았음을 확인했다. 마지막 완료 workbook의 scoped VERIFY가 운영 pipeline에서 초 단위로 끝난 것도 복구했다.
- Decisions: journal의 `RUNNING`만 신뢰하지 않고 실제 PID·command line·로그 갱신을 함께 확인한다. 새 writer가 없으므로 동일 DB/input/output, 16 AI workers, reviewer `Codex Current Code Resume 2026-07-27`, 80,000-byte staged 임계값 계약으로 최신 코드 resume를 재개한다.
- Files: `HANDOFF.md`
- Verification: 관련 Python writer 0개, resume 9 PID 4792 없음. journal은 COMPLETED 112/RUNNING 249/PENDING 25이며 마지막 `ingest-run_d2550b9e257cd5fb0958b4be`는 IMPORT 완료 후 VERIFY를 2.887초에 완료했다. 운영 DB read-only `PRAGMA quick_check=ok`를 48.344초에 통과했고 resume 9 로그에는 신규 실패·traceback이 없다.
- Next: resume 10을 시작해 stale run을 조정하고 `ingest-run_60477cca5907571d1478cad4`의 PACKET 이후 처리와 신규 terminal 결과를 감시한다.

## 2026-07-27 06:48 - 최신 코드 corpus resume 10 시작
- Completed: writer 부재와 DB 무결성을 확인한 뒤 resume 9와 동일한 DB/input/output, 16 AI workers, low family reasoning, reviewer, 80,000-byte staged 임계값으로 corpus resume 10을 시작했다.
- Decisions: stale lock은 corpus의 PID 검증 경로가 안전하게 회수하도록 두었고, 별도 journal이나 새 artifact root를 만들지 않고 기존 승인 corpus 체크포인트를 재사용했다.
- Files: `HANDOFF.md`; runtime logs `D:\000. MyWorks\002. DB\InferenceDataAIService\incremental-com-corpus\form-approved\current-code-resume-10.stdout.log`, `current-code-resume-10.stderr.log`
- Verification: PID 24100과 lock owner가 일치하고 단일 writer임을 확인했다. 새 corpus `corpus-run_5dc99646b36ca6335a614b8e`가 선택 361건으로 시작됐으며 시작 직후 COMPLETED 112, FAILED 0 상태에서 stale 항목을 재조정하고 CAPTURE/PACKET/DRAFT를 진행 중이다.
- Next: 첫 신규 WORKFLOW terminal 결과와 `ingest-run_60477cca5907571d1478cad4`의 재개 완료를 확인하고, 새 실패가 발생하면 실제 artifact로 재현한다.

## 2026-07-27 06:56 - 과거 989건 대비 현재 corpus 성능 원인 분석
- Completed: 과거 989건 작업과 현재 361건 승인 corpus의 journal·DB·실행 단계를 비교해 체감 시간이 비슷한 원인을 확인했다. 과거 989건은 9분 17초의 OpenXML 구조 캡처가 주 작업이었고 semantic journal은 48건만 COMPLETED, 927건 PENDING이었다. 현재는 112건만 full workflow 완료되어 249건의 AI 의미 분석이 남아 있다.
- Decisions: 현재 resume 10은 재시작 오버헤드와 미완료 DRAFT 재처리를 늘리지 않도록 중단하지 않는다. 즉시 병렬도를 변경하려고 writer를 다시 시작하지 않고, 향후 고정 표본 benchmark에서 AI gate 20/24와 stage-level durable checkpoint를 검증한다. DB writer는 SQLite 안전 때문에 1을 유지한다.
- Files: `HANDOFF.md`
- Verification: 현재 통합 DB는 source documents 1,376건(과거 989 + 현재 archive 계열)이나 workbook analyses는 177건뿐이다. 실제 최근 완료 파일의 단계별 시간은 DRAFT 320.727초, IMPORT 85.063초, PACKET 23.014초, VERIFY 2.887초, LOCATOR 0.465초로 병목이 DRAFT와 IMPORT임을 확인했다. resume 10 PID 24100은 16 logical CPU 환경에서 alive, CPU 누적 증가, 실패 0건이다.
- Next: resume 10의 첫 terminal 완료까지 유지하고, 이후 중단 없이 처리량을 측정해 AI worker 상향보다 먼저 DRAFT checkpoint 재사용과 IMPORT 병목을 코드 단위로 최적화할지 결정한다.

## 2026-07-27 07:05 - corpus AI 사용량 증폭 경로 감사
- Completed: CLI 옵션부터 corpus locator/DRAFT 호출·분할·재시도·캐시 경로를 추적해 높은 AI 사용량의 구체적인 원인을 확인했다. 현재 명령의 `--analysis-workers 16`과 `--reasoning-effort low`는 form-family 검토에만 적용되고, 실제 corpus는 코드에 고정된 AI 16 workers와 `medium` reasoning으로 실행된다.
- Decisions: worker 수만 낮추면 호출 속도만 낮아지고 총 토큰은 거의 줄지 않으므로 우선순위를 reasoning 전달 수정, retry budget 제한, stage별 AI 호출 telemetry, 미완료 queue/checkpoint 보존에 둔다. 80,000-byte staged 임계값 변경은 기존 provenance/cache를 무효화할 수 있어 실행 중인 resume 10에는 적용하지 않는다.
- Files: `HANDOFF.md`
- Verification: 현재 대형 workbook plan은 각각 17~26개 DRAFT part를 가지며, locator batch 실패 시 1회 batch와 최대 6회 singleton 호출, DRAFT JSON retry, workbook당 3회, outer corpus 최대 4 pass로 호출이 증폭될 수 있다. corpus executor가 eligible 249건을 모두 즉시 `RUNNING`으로 표시해 반복 중단 시 동일 앞쪽 작업을 다시 예약하는 구조도 확인했다. 점검 시 PID 24100은 alive였고 별도 pipeline Codex 자식은 없었으나 11개 DRAFT가 진행 중이었다.
- Next: 사용자가 즉시 비용 중단을 우선하면 PID 24100의 정확한 process tree를 체크포인트 보존 상태로 종료한 뒤, corpus reasoning/AI/retry budget CLI를 실제 workflow에 배선하고 낮은 사용량 표본으로 재개한다. 계속 처리를 우선하면 현재 provenance를 유지해 resume 10을 중단 없이 완료시킨다.

## 2026-07-27 07:11 - 과거 989건 AI 완료 사실 정정 및 파이프라인 비교
- Completed: 오래된 canonical corpus journal만으로 과거 989건이 의미 분석 미완료라고 판단한 이전 설명을 정정했다. 실제 최종 경로는 별도 table-first v8 배치이며 989/989 AI 분석 성공, Study 3,710건, evidence 5,546건, history DB 구축과 golden question 10/10 검증까지 완료돼 있다.
- Decisions: 과거 완료는 유효한 AI 분석으로 인정한다. 현재 지연의 원인은 과거 결과가 없어서가 아니라, 추가 archive 387건을 기존 table-first 1-call 경로가 아닌 새로운 source-complete canonical locator/staged-DRAFT/import 경로로 처리한 아키텍처 변경으로 본다.
- Files: `HANDOFF.md`
- Verification: 과거 full table-first 본 실행은 workers 3, reasoning low로 972건 ANALYZED·6건 REUSED·9건 NO_TABLES·2건 FAILED를 약 4시간 32분에 처리했고, 재시도는 987건을 재사용해 2건만 AI 호출했다. 과거 request는 평균 39.3KB, 전체 38MB였다. 현재 생성된 36개 canonical plan은 평균 1.9MB, 평균 24.5 parts, 최대 206 parts로 총 882 parts이며 corpus reasoning은 실제 `medium`이다. 현재 archive 387개와 과거 989개는 exact filename overlap 0이라 결과 자체를 직접 cache-hit할 수는 없지만 기존 table-first 처리 방식을 적용할 수 있는지는 별도 경로로 검토해야 한다.
- Next: 높은 비용의 resume 10을 계속할지 중단할지 사용자 결정을 받은 뒤, 추가 387건을 table-first 방식으로 분석해 기존 history에 병합하는 저비용 경로와 canonical source-complete가 꼭 필요한 예외 파일만 선별하는 경로를 설계한다.

## 2026-07-27 07:15 - 구조 유사도 결과의 추출 재사용 단절 확인
- Completed: 현재 form preflight manifest가 구조·양식군·signature·registry decision을 만들지만 corpus 진입 시 `_manifest_relative_paths`가 경로와 SHA만 남기고 form signature/contract를 버리는 것을 확인했다. downstream workflow에는 form family/signature 사용 지점이 없어 `KNOWN_FORM`도 범용 locator/staged-DRAFT AI를 처음부터 수행한다.
- Decisions: 정상 목표는 과거 989건의 구조+AI 결과에서 form별 extraction recipe를 만들고, 신규 파일은 top structural match의 recipe를 결정적으로 replay한 뒤 검증 실패 부분에만 AI를 호출하는 것이다. 현재 preflight는 승인/제외 gate 역할에 그쳐 이 목표를 구현하지 못한다.
- Files: `HANDOFF.md`
- Verification: 최신 preflight known catalog는 162건뿐이며 과거 989 table-first 구조 catalog 전체가 current registry에 연결되지 않았다. 선택 361건 중 `LINKED_EXISTING`은 8건, `APPROVED_NEW`는 351건이고 268 families/353 signatures로 과분할됐다. corpus 코드는 manifest의 form metadata를 참조하지 않고 selected relative path 361개만 전달한다.
- Next: resume 10 중단 여부를 확정한 뒤, 989 table-first catalog를 form-template registry로 이관하고 high-confidence match는 AI-free recipe replay, low-confidence/coverage gap만 부분 AI 처리하도록 증분 pipeline을 구현한다.

## 2026-07-27 07:24 - 구조 재사용 기반 증분 분석 설계
- Completed: 기존 989건을 구조 template과 실행 가능한 extraction recipe로 전환하고, 신규 파일은 deterministic Top-K 매칭·프로그램 추출·검증을 거치는 증분 pipeline 설계를 작성했다. fingerprint/template/recipe/match/extraction/validation/patch 계약, AI 호출 상한, 캐시 키, fail-closed 정책, 기존 코드 변경 지점과 단계별 전환 기준을 정의했다.
- Decisions: exact/high-confidence 구조는 AI 없이 replay하고, AI는 모호한 template 선택 또는 처음 발견된 variant/family의 recipe patch에만 최대 1회 사용한다. 숫자·수식·통계·단위·근거 셀은 프로그램 소유로 유지하며, 검증 실패 시 전체 workbook DRAFT로 자동 폴백하지 않는다. 이번 단계는 설계만 수행해 실행 중인 resume 10에는 변경을 가하지 않았다.
- Files: `STRUCTURE_REUSE_INCREMENTAL_PIPELINE_DESIGN.md`, `HANDOFF.md`
- Verification: 설계 문서 617줄을 UTF-8로 읽어 필수 8개 계약/운영 섹션이 모두 존재함을 확인했고 `git diff --check -- STRUCTURE_REUSE_INCREMENTAL_PIPELINE_DESIGN.md`를 통과했다. 애플리케이션·서버·pipeline은 실행하지 않았다.
- Next: 승인되면 `fingerprint v2 -> 989개 구조 catalog -> 빈도가 높은 family 1개의 executable recipe -> 과거 replay` 순서의 얇은 수직 절단을 구현하고, 수치·근거 셀 100% 일치와 후속 동일 구조 AI 0회를 검증한다.

## 2026-07-27 07:28 - 구조 재사용 1차 구현 범위 조사
- Completed: Capture v2 JSON의 workbook/sheet/cell/merge/formula/number-format 계약, 기존 form signature·similarity 구현, unittest 패턴과 현재 작업 트리 상태를 확인했다.
- Decisions: 실행 중인 resume 10이 import한 기존 모듈은 수정하지 않고, 1차 수직 절단을 독립 신규 모듈과 합성 Capture v2 fixture 테스트로 구현한다. 기존 v1 preflight 연동과 989개 실제 catalog bootstrap은 이 기반 계약이 검증된 다음 단계로 둔다.
- Files: `HANDOFF.md`
- Verification: Capture v2가 정확한 row/column/coordinate, raw/display/formula, merge range/role, number format과 sheet bounds를 보존함을 코드와 기존 테스트로 확인했다. PID 24100은 계속 실행 중이며 이번 조사에서 pipeline 상태를 변경하지 않았다.
- Next: fingerprint v2, template/recipe 계약, deterministic matcher/executor/validator 모듈과 대표 수직 절단 테스트를 추가한다.

## 2026-07-27 07:31 - fingerprint v2와 실행형 recipe 계약 구현
- Completed: Capture v2에서 숫자값과 source identity를 digest에서 제외하고 anchor 위치, sheet/표 형태, merge 기하, formula pattern, column/number-format 역할을 만드는 deterministic fingerprint v2를 구현했다. 대표 fingerprint를 포함하는 form template 계약과 sheet/anchor/region/axis/field selector를 정적 검증하는 extraction recipe 계약도 추가했다.
- Decisions: 첫 버전은 `exactly-one` sheet, unique anchor, row-repeated region, exact-source-cell evidence만 허용해 잘못된 범용 해석을 막는다. 지원하지 않는 selector는 묵시적으로 처리하지 않고 계약 오류로 실패시킨다.
- Files: `inference_data_ai_structure_fingerprint.py`, `inference_data_ai_extraction_recipe.py`, `HANDOFF.md`
- Verification: 두 신규 Python 모듈을 `python -m py_compile`로 컴파일했고 해당 파일에 `git diff --check`를 통과했다.
- Next: template hard gate/가중치 기반 Top-K matcher, recipe executor와 capture-backed validation을 구현한다.

## 2026-07-27 07:37 - AI-free 구조 재사용 수직 절단 구현
- Completed: 승인 template의 sheet/tabular/required-anchor hard gate와 6개 구조 특징 가중치로 Top-K를 만드는 matcher, recipe의 sheet·anchor·region·header-role을 해석하는 executor, exact Capture v2 evidence와 미사용 정량 셀을 검사하는 fail-closed validator를 구현했다. 합성 Capture fixture에서 값과 source SHA가 달라도 동일 구조가 `EXACT_REUSE`로 선택되어 AI 0회로 값·단위·C4/C5 근거 셀을 추출하는 수직 절단 테스트를 추가했다.
- Decisions: executor는 `EXACT_REUSE` 또는 향후 `AI_CONFIRMED_REUSE` 결정만 허용하며 selected recipe version이 다르면 실행을 거부한다. required anchor가 없으면 `NEW_TEMPLATE_REQUIRED`, evidence 변조 또는 예상 밖 정량 영역은 저장 전 검증 실패로 처리한다.
- Files: `inference_data_ai_recipe_matcher.py`, `inference_data_ai_recipe_executor.py`, `inference_data_ai_recipe_validation.py`, `inference_data_ai_structure_fingerprint.py`, `inference_data_ai_extraction_recipe.py`, `tests/test_inference_data_ai_structure_reuse.py`, `HANDOFF.md`
- Verification: `python -m unittest tests.test_inference_data_ai_structure_reuse -v`의 5개 테스트가 모두 통과했다. 신규 5개 모듈의 `py_compile`과 변경 파일 `git diff --check`도 통과했다.
- Next: 실제 989개 Capture/table-first 자산을 읽는 catalog bootstrap을 구현해 fingerprint 분포와 가장 큰 family를 산출하고, 첫 실제 family의 recipe replay 대상과 비교 기준을 확정한다.

## 2026-07-27 07:44 - 읽기 전용 989건 구조 catalog bootstrap 구현
- Completed: table-first requests의 revision UID/SHA를 Capture v2 DB와 결합하고, SQLite 집계로 전체 셀 값을 Python에 적재하지 않은 채 fingerprint v2를 만드는 경로를 구현했다. exact structure 그룹과 0.90 초기 similarity threshold의 candidate family를 산출하는 AI-free/read-only bootstrap CLI 및 합성 DB 통합 테스트를 추가했다.
- Decisions: 분석·projection artifact가 모두 있는 request만 catalog에 포함하고, revision UID/SHA 불일치는 즉시 실패한다. approximate family는 아직 `CANDIDATE`이며 실제 replay 검증 전 `APPROVED`로 사용하지 않는다. DB는 URI `mode=ro`와 `query_only`로 열고 종료 시 명시적으로 close한다.
- Files: `inference_data_ai_structure_fingerprint.py`, `inference_data_ai_recipe_matcher.py`, `inference_data_ai_template_bootstrap.py`, `tests/test_inference_data_ai_structure_reuse.py`, `HANDOFF.md`
- Verification: 구조 재사용 테스트 7개가 모두 통과했으며, payload와 SQLite 경로의 fingerprint 동일성 및 서로 다른 수치의 2개 request가 AI 0회로 하나의 reusable exact structure가 되는 것을 확인했다. 신규 6개 모듈 `py_compile`과 변경 파일 `git diff --check`를 통과했다.
- Next: 실제 989건에 bootstrap을 실행해 누락 0 여부, exact/candidate family 분포, 가장 큰 family의 대표 구조와 replay 우선순위를 기록한다.

## 2026-07-27 07:46 - 실제 989건 fingerprint catalog 생성
- Completed: 완료된 table-first 989건 전체를 live Capture v2 DB에 revision UID로 연결해 AI 없이 구조 catalog를 생성했다. 989/989 fingerprint 성공, capture 누락 0, analysis/projection 누락 0이며 실행 시간은 약 21초였다.
- Decisions: 현재 결과의 0.90 candidate clustering은 957 families로 과분할됐으므로 template registry로 승인하지 않는다. exact digest는 안전한 동일 구조 확인에 사용하고, 실제 재사용성은 신규 archive를 과거 989 representative 전체에 Top-K 매칭한 점수 분포와 replay 결과로 보정한다. 제목·날짜·샘플명 같은 변동 텍스트를 필수 anchor로 승격하지 않는다.
- Files: runtime artifacts `D:\000. MyWorks\002. DB\InferenceDataAIService\structure-reuse\bootstrap-989-v1\catalog-limit-10.json`, `D:\000. MyWorks\002. DB\InferenceDataAIService\structure-reuse\bootstrap-989-v1\catalog.json`; `HANDOFF.md`
- Verification: catalog summary는 exact structures 973, 중복 exact structures 14/30 files, candidate families 957/60 reusable files, 최대 family 4 files, AI calls 0이다. PID 24100은 계속 실행 중이며 bootstrap은 DB `mode=ro/query_only`만 사용했다.
- Next: 최신 preflight의 신규 승인 capture를 과거 989 structures와 전수 Top-K 비교해 score/component 분포를 구하고, matcher anchor/geometry 가중치와 AI-free/AI-review 경계를 실제 데이터로 보정한다.

## 2026-07-27 07:53 - 989건 table/block 구조 재사용 catalog 구현
- Completed: workbook 전체 조합이 아니라 table-first가 이미 결정론적으로 분리한 5,546개 표를 값 제외 kind layout, 상대 좌표, merge 기하, numeric column role/format/count bucket으로 fingerprint하는 table-level catalog를 구현하고 실제 989건에 실행했다. 과거 AI table type/confidence도 구조별 통계로 결합했다.
- Decisions: workbook fingerprint는 후보 파일 검색에 사용하고, 실제 extraction recipe는 재사용성이 더 높은 table/block 구조를 기본 단위로 둔다. numeric sample 값과 변동 header text는 structure digest에서 제외하며 header token은 후보 의미 비교용으로만 보존한다. TEXT/title block과 정량 extraction block을 분리해 우선순위를 계산한다.
- Files: `inference_data_ai_table_structure_catalog.py`, `tests/test_inference_data_ai_structure_reuse.py`, runtime artifact `D:\000. MyWorks\002. DB\InferenceDataAIService\structure-reuse\bootstrap-989-v1\table-structure-catalog.json`, `HANDOFF.md`
- Verification: 구조 재사용 테스트 9개가 통과했다. 실제 catalog는 5,546 tables/3,195 exact block structures, 재사용 구조 508개에 2,633 tables, 859 workbooks를 포괄했다. 그중 비-TEXT 정량 재사용 구조는 357개/1,284 tables/415 workbooks이고 최대 구조는 79 workbooks에서 반복됐다. AI calls는 0이다.
- Next: 실행 중인 신규 361건 Top-K workbook audit 결과를 회수한 뒤 table-first request를 AI 없이 생성 가능한 신규 파일부터 table/block catalog와 교차 매칭하고, 의미 일관성 1.0인 다빈도 COMPARISON 구조 하나를 첫 recipe replay 대상으로 선정한다.

## 2026-07-27 07:58 - 신규 workbook audit 및 첫 실제 table recipe replay
- Completed: 신규 승인 361건을 과거 989 workbook structures 전체와 전수 Top-K 비교했다. 동시에 과거 14 workbooks/15 tables에서 반복된 function comparison 구조의 metric column 의미를 consensus로 recipe화하고 기존 projection에 replay했다.
- Decisions: 신규 workbook 최고 유사도는 최대 0.7375, 평균 0.4490으로 361건 모두 안전한 전체-workbook reuse 대상이 아니었다. 제품/시트 조합이 달라 workbook threshold를 낮추지 않고, 동일 파일 안의 반복 table/block을 과거 3,195 structures와 교차 매칭하는 방향을 확정했다. 첫 recipe는 exact table fingerprint에서만 실행하며 불일치 시 적용하지 않는다.
- Files: `inference_data_ai_incremental_match_audit.py`, `inference_data_ai_table_recipe_mining.py`, `tests/test_inference_data_ai_structure_reuse.py`; runtime artifacts `D:\000. MyWorks\002. DB\InferenceDataAIService\structure-reuse\bootstrap-989-v1\incremental-match-audit.json`, `...\recipes\table-recipe-9f1774bb07158e602858.json`, `...\replays\table-recipe-9f1774bb07158e602858.replay.json`; `HANDOFF.md`
- Verification: 신규 audit 361/361 완료, AI 0회, score 0.60~0.75는 86건이고 275건은 0.60 미만이었다. 첫 recipe는 COMPARISON table 15/15에서 SPL, THD, SPL+THD, SPL+THD+F0, NOISE, TOUCH, HOHD, TOTAL NG, TOTAL NG RATE의 상대 열을 일치시켰고 code-owned numeric facts와 evidence range도 15/15 통과했다. recipe status는 `VERIFIED_HISTORICAL_REPLAY`다.
- Next: 신규 361 capture에서 table-first request만 AI 없이 생성해 3,195 과거 table structures와 exact/near match를 측정하고, 검증된 첫 recipe가 신규 표에 실제로 적용 가능한지 확인한다.

## 2026-07-27 08:02 - 신규 361건 AI-free table 분해 및 재사용 분포 확인
- Completed: 최신 preflight 승인 361건에서 semantic source packet과 table-first request를 AI 없이 생성하고 과거 3,195 table structures 및 verified recipe와 exact 매칭했다. 이어 성공한 신규 356건 자체를 table/block catalog로 묶어 신규 archive 내부의 반복 구조 분포를 확인했고 설계 문서를 workbook routing + table recipe 계층으로 보정했다.
- Decisions: 과거와 exact 일치한 434개 block은 모두 비정량 metadata/TEXT 형태여서 파라미터 recipe를 즉시 적용하지 않는다. 과거의 정량 block과 exact 일치는 0건이므로 threshold를 낮춰 강행하지 않고 near-match 후보는 AI가 구조 재사용 가능성만 판정하게 한다. 신규 archive 내부에서는 동일 정량 구조를 파일마다 분석하지 않고 structure당 한 번 분석·recipe화한다.
- Files: `inference_data_ai_incremental_table_match.py`, `tests/test_inference_data_ai_structure_reuse.py`, `STRUCTURE_REUSE_INCREMENTAL_PIPELINE_DESIGN.md`; runtime artifacts `D:\000. MyWorks\002. DB\InferenceDataAIService\structure-reuse\incremental-361-table-match-v1\report.json`, `...\requests\*.json`, `...\table-structure-catalog.json`; `HANDOFF.md`
- Verification: 361건 중 356건을 28.6초에 처리하고 1,981 tables를 만들었으며 AI calls 0이다. 과거 exact block match는 434 tables/219 workbooks였지만 정량 match와 첫 verified recipe match는 0이었다. 신규 내부 catalog는 1,981 tables/866 structures이며 반복 정량 구조 126개가 662 tables와 290 workbooks를 포괄했다. 구조 재사용 unittest 10개, 신규 10개 모듈 `py_compile`, 전체 변경 파일 `git diff --check`를 통과했다. 5건은 Capture v2 declared non-empty count와 packet 노출 count가 1~5 cells 달라 fail closed됐다. resume 10 PID 24100은 계속 실행 중이다.
- Next: 신규 반복 정량 구조를 의미 있는 numericColumnCount>=2 중심으로 우선순위화하고, 각 구조의 대표 table과 과거 near candidates만 AI에 한 번 제시해 recipe를 생성한다. 생성된 recipe는 같은 구조 전체에 프로그램 replay하고 실패 구조만 격리한다.

## 2026-07-27 08:46 - 기존 16-worker AI 파이프라인 중단
- Completed: 사용자의 전환 지시에 따라 PID 24100의 `form-pipeline-complete` 실행과 그 자식 프로세스 트리를 확인한 뒤 종료했다. 실행 중이던 16-worker 경로의 Codex 분석 호출도 함께 중단했으며 기존 corpus checkpoint와 산출물은 보존했다.
- Decisions: 기존 파일별 재분석 경로는 재개하지 않는다. 남아 있는 `corpus-journal.json.lock`은 종료 전 생성된 5-byte stale lock일 수 있으나 checkpoint 보존을 위해 삭제하지 않았고, 새 구조 재사용 경로에서는 이 lock을 사용하지 않는다.
- Files: `HANDOFF.md`
- Verification: 정확한 루트 명령행을 확인하고 루트 1개와 자식 37개를 대상으로 종료했다. 종료 후 루트 PID 잔존 0, 기존 대상 PID 잔존 0, 해당 form-approved 경로의 `cmd/node/codex/python` worker 잔존 0을 확인했다.
- Next: 신규 361건의 반복 정량 table 구조를 우선순위화하고, 구조당 대표 표 1개와 과거 Top-K 후보만 담는 bounded recipe proposal을 만든 뒤 첫 구조를 deterministic replay한다.

## 2026-07-27 08:59 - 구조당 1회 recipe 우선순위·계약 경로 구현
- Completed: 신규 반복 정량 table 구조를 AI 없이 우선순위화하고 값이 제거된 대표 표 1개와 과거 Top-3 구조만 AI 입력 후보로 만드는 모듈을 구현했다. strict decision schema, 1회 호출·재시도 0 telemetry, exact fingerprint 전용 recipe compiler, code-owned numeric facts/evidence replay를 함께 추가했다. 실제 361건 catalog에서 우선순위 report를 생성했다.
- Decisions: metric 의미와 상대 열 매핑만 구조당 AI가 1회 결정하며 raw 숫자·통계는 prompt에서 제외한다. 실제 min/max/average/count/sourceRange는 프로그램이 table-first request에서 가져온다. exact fingerprint가 다르거나 열 role이 변하면 실행하지 않고, 고차원/frequency 의심 구조는 AI 호출 전 검토 대상으로 둔다.
- Files: `inference_data_ai_table_recipe_proposal.py`, `tests/test_inference_data_ai_structure_reuse.py`; runtime artifact `D:\000. MyWorks\002. DB\InferenceDataAIService\structure-reuse\incremental-361-table-match-v1\recipe-priority-report.json`; `HANDOFF.md`
- Verification: 구조 재사용 unittest 13개와 신규 모듈 `py_compile`이 통과했다. 실제 report는 94 structures/335 tables를 queue화했고 그중 92개가 proposal-ready, 2개가 사전 검토 대상이며 AI calls 0이다. Top-K feature를 미리 cache해 동일 실제 실행 시간을 초기 99.8초에서 5.7초로 줄였다. 1순위는 `table-structure-ddcef3c12ad8099e3b61`로 10 workbooks/11 tables/7 measure columns다.
- Next: 1순위 구조에만 low reasoning AI 1회를 실행해 semantic recipe를 만들고, 동일 구조 11개 표에 deterministic replay하여 값·통계·evidence가 모두 프로그램에서 추출되는지 확인한다.

## 2026-07-27 09:01 - 첫 신규 구조 1회 판단 및 11-table deterministic replay
- Completed: 1순위 `table-structure-ddcef3c12ad8099e3b61`에 12,382-byte redacted prompt로 low reasoning AI를 정확히 1회 호출해 신규 recipe를 만들었다. 이어 동일 exact 구조 11개 표 전체에 프로그램 replay하여 열 통계뿐 아니라 각 숫자/수식 셀의 좌표, 표시값, number format, count/percent 역할을 추출했다.
- Decisions: 과거 Top-3는 검증 recipe가 없고 similarity 최고 0.6162여서 억지 재사용하지 않고 `NEW_RECIPE`로 결정됐다. 구조 의미는 S931 function lot test의 두 source-authored condition 비교이며 metric은 Bako/Hearing/Hearing 2/Hearing 3/Air leak/Total NG/NG rate 상대 열이다. 같은 열에 건수와 비율 행이 섞이므로 열 평균만 결과로 쓰지 않고 셀별 표시형식과 exact coordinate를 code-owned fact로 보존한다. semantic canonical review 전 상태는 명시적으로 `NEEDS_CANONICAL_REVIEW`다.
- Files: `inference_data_ai_table_recipe_proposal.py`, `tests/test_inference_data_ai_structure_reuse.py`; runtime artifacts `...\decisions\table-structure-ddcef3c12ad8099e3b61.decision.json`, `...\telemetry\table-structure-ddcef3c12ad8099e3b61.ai.json`, `...\recipes\structure-recipe-ddcef3c12ad8099e3b61.json`, `...\replays\structure-recipe-ddcef3c12ad8099e3b61.replay.json`; `HANDOFF.md`
- Verification: AI telemetry는 budget 1, attempted/succeeded 1/1, retry 0, reasoning low, duration 25.1초, prompt 12,382 bytes, output 2,215 bytes다. Replay는 11/11 pass, 77 code-owned column facts, 264 exact cell facts, replay AI calls 0이다. 셀 facts는 NUMBER 132개와 PERCENT 132개로 분리됐고 첫 표에서 `I16=1`/`I17=0.47%`처럼 동일 metric 열의 건수·비율 좌표가 구분됨을 확인했다. unittest 13개와 `py_compile`도 통과했다.
- Next: 이 성공 recipe를 batch registry에 등록하고, exact fingerprint match만 AI 0회로 자동 실행하며 신규 구조는 budget-limited queue에서 structure당 최대 1회만 호출하도록 batch 상태/예산 파일을 구현한다.

## 2026-07-27 09:10 - semantic header 분리 및 AI budget batch 잠금
- Completed: 기하 fingerprint가 같은 11개 표를 metric header 기준으로 교차검사해 서로 다른 의미 구조가 섞여 있음을 발견했다. 앞 단계의 기하-only 11/11 replay 승인을 무효화하고, 상대 MEASURE_VALUE 열별 정규화 header signature를 recipe match key에 추가했다. 기존 AI 1회 판단은 실제로 같은 S931 metric signature인 6개 표에만 다시 replay해 registry에 등록했다. 구조당 최대 1회, retry 0, run AI budget 초과 시 대기하는 batch control도 구현하고 실제 상태 파일을 생성했다.
- Decisions: 기하 digest만 같은 X626/X526 variant에는 S931 recipe를 절대 적용하지 않는다. exact 자동 재사용 조건은 `fingerprintSha256 + semanticHeaderSignature + recipeVersion`이며, 같은 기하라도 metric header 순서/의미가 다르면 별도 recipe structure다. 현재 run budget은 1로 고정해 이미 사용한 1회 외 추가 AI 호출을 허용하지 않는다. semantic 결과는 canonical review 전까지 승인 상태가 아니다.
- Files: `inference_data_ai_table_recipe_proposal.py`, `inference_data_ai_structure_batch_control.py`, `tests/test_inference_data_ai_structure_reuse.py`, `STRUCTURE_REUSE_INCREMENTAL_PIPELINE_DESIGN.md`; runtime artifacts `...\recipe-priority-report.json`, `...\recipe-registry.json`, `...\batch-control.json`, 갱신된 `...\recipes\structure-recipe-ddcef3c12ad8099e3b61.json`과 replay; `HANDOFF.md`
- Verification: 최종 신규 queue는 반복 semantic recipe structures 90개/307 tables/241 workbook references이며 proposal-ready 88개, manual-review 2개다. 첫 registry recipe는 S931 6 tables/5 unique content workbooks에서 6/6 pass, 42 column facts, 144 exact cell facts이고 AI는 최초 1회뿐이다. 실제 resolver는 S931 대표 표에 `EXACT_RECIPE_MATCH`/AI 0, 같은 기하의 X626 표에는 `NO_REGISTERED_RECIPE`를 반환했다. batch budget은 max/consumed/remaining 1/1/0, 87 structures는 budget wait, file-level AI calls 0이다. unittest 15개, 관련 12개 모듈 `py_compile`, 변경 문서·테스트 `git diff --check`를 통과했고 기존/신규 AI process 잔존 0을 확인했다.
- Next: canonical reviewer가 첫 S931 semantic recipe를 승인하거나 수정한다. 승인 후에는 이 recipe exact match를 AI 0회로 유지하고, 다음 상위 recipe structure를 처리할 때만 `max-ai-calls`를 의도적으로 1씩 올려 한 구조 1회 판단→전체 replay→registry 등록 순서를 반복한다.

## 2026-07-27 10:07 - 두 번째 구조 1회 판단 및 fail-closed 격리
- Completed: 사용자 지시에 따라 run AI budget을 1에서 2로 한 칸만 올리고 다음 우선순위 `table-structure-b20ab3829ca756bd7e50`(8 tables/8 workbooks)의 redacted reliability-result 구조를 low reasoning으로 정확히 1회 판단했다. AI 출력은 실행 가능한 metric 열을 하나도 제시하지 못해 strict recipe validator가 거부했고, 값 추출과 replay를 실행하지 않은 채 `QUARANTINED_NO_RETRY`로 batch 상태에 반영했다.
- Decisions: 이 block은 표 제목이 `RELIABILITY TEST RESULT`여도 maker/checker/signature 등 header metadata 성격이 강하며, parameter metric mapping이 없는 결과를 억지 recipe로 승격하지 않는다. 실패한 동일 구조를 재호출하지 않고 별도 검토 대상으로 유지한다. run budget은 max/consumed/remaining 2/2/0으로 다시 잠갔다.
- Files: runtime artifacts `D:\000. MyWorks\002. DB\InferenceDataAIService\structure-reuse\incremental-361-table-match-v1\telemetry\table-structure-b20ab3829ca756bd7e50.ai.json`, 갱신된 `...\batch-control.json`, `...\recipe-registry.json`; `HANDOFF.md`
- Verification: 두 번째 telemetry는 attempted 1, succeeded 0, retry 0, reasoning low, duration 14.9초, prompt 9,687 bytes이며 실패 사유는 `An executable recipe requires at least one metric column.`이다. 전체 새 경로 누계는 AI attempted 2/succeeded 1/failed 1/retry 0이며 관련 process 잔존 0이다. registry는 안전한 S931 recipe 1개/6 tables만 유지하고, batch actions는 exact-ready 1, quarantined-no-retry 1, budget-wait 86, manual-review 2다. unittest 15개, 신규 2개 모듈 `py_compile`, 변경 문서·테스트 `git diff --check`를 통과했다.
- Next: 추가 진행 시 budget을 3으로 한 칸만 올려 현재 다음 proposal-ready 구조 1개를 동일 절차로 처리한다. 격리된 reliability header block은 metric 목적이 명시되기 전까지 재호출·replay하지 않는다.

## 2026-07-27 10:20 - source-owned 비교군 selector 보강
- Completed: 최종 목표 실행을 시작하면서 기존 S931 recipe가 대표 파일의 비교군 label과 title을 다른 파일에도 고정 복사하는 안전성 문제를 발견했다. recipe compiler가 AI의 예시 label을 대표 표의 상대 text cell selector로 변환하고, executor가 각 대상 표에서 실제 title·group label·comparison relation을 프로그램으로 추출하도록 보강했다.
- Decisions: metric 의미/상대 열과 group role만 semantic contract에 남기고 source-authored title과 group label은 파일별 captured preview에서 code-owned로 읽는다. 대표 label이 정확히 하나의 상대 셀로 해석되지 않으면 recipe compile을 실패시킨다. 이 보강 전 S931 replay 결과는 다시 생성해야 registry 안전성이 유지된다.
- Files: `inference_data_ai_table_recipe_proposal.py`, `tests/test_inference_data_ai_structure_reuse.py`, `HANDOFF.md`
- Verification: 서로 같은 구조지만 group label이 `Sample A/B`에서 `After/Before`로 바뀐 fixture에서 recipe가 새 source label과 relation을 추출하고 `CODE_FROM_CAPTURED_TABLE_PREVIEW` authority를 기록함을 확인했다. 구조 재사용 unittest 16개와 신규 모듈 `py_compile`이 통과했다.
- Next: 과거 989건과 신규 queue의 semantic signature overlap 22개를 이용해 일관된 DESCRIPTIVE contract를 AI 0회로 부트스트랩하고, 나머지 고유 signature만 구조당 1회 처리한다.

## 2026-07-27 10:24 - 과거 989 semantic consensus AI-free bootstrap
- Completed: 과거 table-first 989건의 request/analysis를 신규 queue의 metric header signature와 교차해, DESCRIPTIVE type 일관성 0.9 이상·완전 relative metric mapping 2건 이상인 contract를 AI 없이 생성하는 bootstrap을 구현하고 실제 실행했다. 기존 telemetry가 있는 S931/격리 구조는 덮어쓰지 않았으며, 첫 S931 recipe도 source-owned group selector 방식으로 다시 compile/replay했다.
- Decisions: 과거 consensus는 COMPARISON나 type 충돌 signature에 사용하지 않고 일관된 DESCRIPTIVE contract에만 허용한다. historical recipe는 `decisionSource=HISTORICAL_989_CONSENSUS`, `aiCallCount=0`을 명시하고 동일 semantic signature 전체에 deterministic replay가 전부 통과한 경우만 registry에 등록한다.
- Files: `inference_data_ai_historical_semantic_bootstrap.py`, `inference_data_ai_table_recipe_proposal.py`, `inference_data_ai_structure_batch_control.py`, `tests/test_inference_data_ai_structure_reuse.py`; runtime artifacts `...\historical-semantic-contracts.json`, `...\historical-bootstrap-report.json`, 20개 historical decision/recipe/replay, 갱신된 registry/control/S931 replay; `HANDOFF.md`
- Verification: 신규 queue signature 58개 중 과거 overlap 22개를 찾았고, strict 조건으로 8개 semantic contracts를 승인·14개를 거부했다. 승인 contract가 신규 20 structures/82 tables/46 workbook references에서 전부 replay 통과했다. registry는 총 21 recipes/88 tables/51 workbook references이며 그중 historical AI 0회 recipe 20개, bounded AI recipe 1개다. S931 6개 표의 비교군은 각 파일에서 실제 label을 읽어 서로 다른 Test/Normal 문구와 `CODE_FROM_CAPTURED_TABLE_PREVIEW` authority로 저장됨을 확인했다. unittest 18개와 관련 모듈 `py_compile`이 통과했다.
- Next: 남은 66 budget-wait structures를 고유 semantic signature 대표 1회 판단과 같은-signature AI-free 전파로 처리하는 resumable orchestrator를 구현·실행한다.

## 2026-07-27 10:54 - 신규 반복 정량 구조 큐 완주
- Completed: 재개 가능한 구조 단위 orchestrator로 신규 361건 중 반복 정량 큐 90개 구조를 전부 처리해 미해결 0으로 종료했다. 56개 레시피가 193개 표/146개 workbook reference에서 replay 검증을 통과해 등록됐고, 34개 구조는 AI 결정·사전검사·recipe 계약·AI 출력 실패 사유별로 격리했다.
- Decisions: AI는 파일별이 아니라 고유 구조별 최대 1회만 허용하고 실패 재시도는 하지 않았다. 동일 semantic signature는 검증된 recipe만 AI 0회로 전파했으며, replay 또는 계약을 통과하지 못한 구조는 억지 등록하지 않고 fail-closed 격리했다.
- Files: `inference_data_ai_structure_completion.py`; runtime artifacts `D:\000. MyWorks\002. DB\InferenceDataAIService\structure-reuse\incremental-361-table-match-v1\completion-state.json`, `recipe-registry.json`, `recipes\`, `replays\`, `decisions\`, `telemetry\`, `quarantine\`; `HANDOFF.md`
- Verification: completion state `COMPLETED`, 90/90 structures, unresolved 0, registered 56 recipes/193 tables/146 workbook references를 확인했다. AI attempted/succeeded/failed 62/56/6, file-level AI calls 0, retry 0이며 종료 후 관련 completion Python process가 남지 않았다.
- Next: 361개 파일과 1,981개 표 전체를 등록·격리·반복 정량 큐 외 대상으로 합계 대조하는 최종 coverage report를 만들고, 좁은 단위 테스트/컴파일/문서 검증을 수행한다.

## 2026-07-27 10:59 - 신규 361건 전체 coverage 대조
- Completed: workbook/table/queue/registry/telemetry 합계가 하나라도 맞지 않으면 실패하는 최종 coverage 계산기를 구현하고 실제 artifact를 생성했다. 전체 361개 파일은 356개 캡처 성공과 5개 Capture v2 cell-count 불일치 실패로 전부 설명되며, 성공 파일의 1,981개 표는 반복 정량 큐 307개와 큐 밖 1,674개로 닫혔다.
- Decisions: 이번 pass에서 프로그램 추출이 검증된 것은 등록된 193개 표뿐이라고 명시하고, 격리 114개 및 큐 밖 정량 1,021개/비정량 653개는 처리 성공으로 과장하지 않는다. exact token 수는 telemetry에 없으므로 prompt/output byte만 보고하고 token usage는 unavailable로 표시한다.
- Files: `inference_data_ai_incremental_coverage.py`, `inference_data_ai_structure_completion.py`, `tests/test_inference_data_ai_structure_reuse.py`; runtime artifact `D:\000. MyWorks\002. DB\InferenceDataAIService\structure-reuse\incremental-361-table-match-v1\final-coverage-report.json`; `HANDOFF.md`
- Verification: 구조 재사용 unittest 19개와 관련 모듈 `py_compile`이 통과했다. 실제 coverage는 registered/quarantined 193+114=307, queue/outside 307+1,674=1,981, captured/failed 356+5=361을 검증했다. AI 62회, 성공/실패 56/6, retry 0, file-level AI 0, prompt/output 826,844/127,495 bytes, serial duration 합계 20.49분이다.
- Next: 설계 문서에 source-owned selector, Capture v2 full-range replay, historical/propagated AI 0 경로와 실제 최종 지표를 반영한 뒤 diff/프로세스/최종 artifact 검증을 수행한다.

## 2026-07-27 11:01 - 구조 재사용 증분 경로 최종 검증
- Completed: 설계 문서의 대상/상태를 실제 361건 기준으로 바로잡고 source-owned selector, Capture v2 full-range replay, historical consensus/verified propagation AI 0 경로, 최종 수치와 미처리 범위를 기록했다. completion resume 시 이전 `AI_BUDGET_EXHAUSTED` 표시가 실행 중 남는 문제도 `RUNNING` 초기화로 보정했다.
- Decisions: 최초 bootstrap의 AI 평균 0.172회/workbook은 목표 0.1을 충족하지 못했다는 사실을 숨기지 않는다. 등록된 56개 recipe의 후속 exact match가 AI 0회로 상각되며, 큐 밖 정량 1,021개는 다음 확대 범위로 남긴다.
- Files: `STRUCTURE_REUSE_INCREMENTAL_PIPELINE_DESIGN.md`, `inference_data_ai_structure_completion.py`, `inference_data_ai_incremental_coverage.py`, `tests/test_inference_data_ai_structure_reuse.py`, `HANDOFF.md`; runtime `...\final-coverage-report.json`
- Verification: 최종 구조 재사용 unittest 19/19, 관련 5개 모듈 `py_compile`, scoped `git diff --check`, 변경 파일 trailing-whitespace 검사가 통과했다. coverage invariant 8개가 모두 true이고 completion은 `COMPLETED`/unresolved 0이다. 관련 Python/Node/CMD worker 0, 전체 Python process 0을 확인했으며 앱/서버는 실행하지 않았다.
- Next: 현재 목표는 완료됐다. 후속 작업을 시작할 경우 `final-coverage-report.json`의 큐 밖 재사용 가능 정량 355개를 별도 recipe 확대 대상으로 삼고, capture 실패 5개는 Capture v2 cell-count 계약을 먼저 수리한다.

## 2026-07-27 12:06 - 신규 361건 DB 적재 범위 감사
- Completed: 신규 361개 revision을 운영 `InputDataFinish.sqlite`와 직접 대조해 source/capture 적재와 canonical AI 분석 적재를 분리 확인했다. 361/361은 `capture_v2_revisions`와 `source_revisions`에 존재하지만 `workbook_analyses`는 115 revision, `knowledge_studies`는 114 revision/305 studies만 연결돼 있다.
- Decisions: 앞서 완료한 구조 재사용 목표는 recipe 판단·프로그램 replay·registry 생성까지이며 canonical DB import 완료를 의미하지 않는다. report의 5개 failure도 Capture v2 자체 실패가 아니라 captured revision에서 table-first semantic request를 만들 때 선언 cell 수와 노출 cell 수가 달라 실패한 건으로 해석을 바로잡는다.
- Files: `HANDOFF.md`; read-only audited DB `D:\000. MyWorks\002. DB\InferenceDataAIService\universal-grid\InputDataFinish.sqlite`
- Verification: 대상 revision ID 361개가 capture/source bridge 361개와 일치했다. canonical 분석은 114 `NEEDS_REVIEW` + 1 `NO_TABULAR_EVIDENCE`이고, 나머지 246개는 `workbook_analyses`가 없다. 등록 recipe의 193 tables/97 unique files 중 43 files만 기존 canonical 분석이 있고 54 files는 아직 없다. 새 structure registry는 현재 `inference_data_ai_analysis_runner.py`, `inference_data_ai_workflow.py`, WPF 실행 경로에 연결돼 있지 않다.
- Next: 전체 신규 Excel을 AI 질문 DB에 포함하려면 structure replay 결과를 canonical Study/Observation/Evidence 계약으로 변환하는 importer를 구현하고, 54 recipe-covered files부터 idempotent import한 뒤 나머지 정량 큐/격리/semantic request 실패 범위를 처리해야 한다.

## 2026-07-27 13:33 - structure replay canonical importer 구현
- Completed: 등록 recipe replay의 code-owned numeric facts와 source ranges를 canonical workbook analysis, descriptive study, outcome, observation, evidence로 변환하는 importer를 구현했다. 기존 non-stale canonical 분석은 보존하고, 미분석 revision은 recipe import·FAILED table request·EXCLUDED·NEEDS_REVIEW terminal manifest 중 하나로 계획한다.
- Decisions: recipe 결과도 canonical review 전이므로 `NEEDS_REVIEW`로 저장하고 비교/effect를 만들지 않는다. 숫자·min/max/average/sample size와 evidence는 replay 프로그램 결과만 사용한다. terminal 파일은 빈 study로 명시적 상태만 저장해 unsupported numeric claim을 만들지 않는다.
- Files: `inference_data_ai_structure_canonical_import.py`, `tests/test_inference_data_ai_structure_canonical_import.py`, `HANDOFF.md`
- Verification: 임시 운영 schema/Capture v2 fixture에서 동일 import를 2회 실행해 workbook analysis/study/observation/evidence 중복이 없음을 확인했다. 신규 단위 테스트 2/2와 두 파일 `py_compile`이 통과했다.
- Next: 운영 `InputDataFinish.sqlite`를 read-only dry-run하여 361개 action과 모든 실제 replay manifest가 검증되는지 확인하고, 합계가 맞으면 단일 transaction으로 적용한다.

## 2026-07-27 13:34 - 신규 361건 canonical 운영 DB 적재
- Completed: 운영 `InputDataFinish.sqlite`에 structure canonical import를 단일 transaction으로 적용했다. 기존 active canonical 분석 115건은 보존하고 미분석 246건을 적재해 target 361/361이 모두 active `workbook_analyses`를 갖게 했다.
- Decisions: 미분석 246건은 replay 검증 54건, terminal NEEDS_REVIEW 186건, table-request FAILED 5건, EXCLUDED 1건으로 나눴다. replay 54건에만 program-owned observation/evidence를 만들고 terminal 파일에는 숫자 claim을 만들지 않았다.
- Files: `inference_data_ai_structure_canonical_import.py`; runtime artifacts `D:\000. MyWorks\002. DB\InferenceDataAIService\structure-reuse\incremental-361-table-match-v1\canonical-import-result.json`, `canonical-import-v1\manifests\`; modified operational DB `D:\000. MyWorks\002. DB\InferenceDataAIService\universal-grid\InputDataFinish.sqlite`; `HANDOFF.md`
- Verification: importer가 각 신규 analysis에 scoped foreign-key/aggregation integrity 검사를 수행하고 전부 통과한 뒤 commit했다. target DB coverage는 active analysis 115→361, missing 246→0, studies 305→425, observations 11,119→11,714다. AI 호출은 0이다.
- Next: target revision별 terminal/status 합계와 recipe provenance를 read-only로 감사하고, 신규 observation이 canonical evidence query에서 실제 검색되는 smoke query를 수행한다.

## 2026-07-27 13:40 - 361건 canonical DB 및 query 경로 감사
- Completed: 운영 DB의 신규 구조 재사용 반영분과 전체 target 361건을 read-only로 감사하고, 신규 outcome 용어 `Tape separation`을 기존 canonical evidence-query 경로로 조회했다. target 361/361에 활성 analysis가 있고 누락 및 revision별 활성 analysis 중복은 각각 0건이다.
- Decisions: importer 소유 246건은 recipe replay 54, recipe 미확보 NEEDS_REVIEW 186, table-request FAILED 5, EXCLUDED 1로 명시 계상한다. recipe 미확보 terminal 레코드에는 숫자 claim을 만들지 않으며, recipe 54건의 결과도 review 전에는 comparison/effect 답변 후보로 승격하지 않는다.
- Files: runtime query artifact `D:\000. MyWorks\002. DB\InferenceDataAIService\structure-reuse\incremental-361-table-match-v1\canonical-import-v1\query-smoke-new-exact.json`; read-only audited DB `D:\000. MyWorks\002. DB\InferenceDataAIService\universal-grid\InputDataFinish.sqlite`; `HANDOFF.md`
- Verification: importer 소유 레코드는 246 analyses, 120 studies/arms, 595 outcomes/observations, 1,430 distinct evidence 및 links로 연결되었다. query는 `DATA-A6DB8C587B0A`의 `Tape separation NG count/rate`와 6개 근거를 반환했고, NEEDS_REVIEW 정책에 따라 answer-eligible effect는 0이었다. 전체 DB `foreign_key_check`의 66건은 기존 legacy `analysis_review_items` 계열 위반이며 신규 transaction은 FK ON 및 analysis별 integrity 검증을 통과했다.
- Next: importer의 최종 audit 산출물에 AI 0회, active-analysis 유일성, importer entity/evidence 수를 명시하고 설계 문서의 5건을 capture 실패가 아닌 table-request 실패로 정정한 뒤 scoped tests를 실행한다.

## 2026-07-27 13:44 - 최종 canonical 감사 계약 및 coverage 용어 정정
- Completed: canonical importer의 공식 결과 계약에 source/capture 수, table-request 성공/실패, AI 사용량, active analysis 유일성, importer 소유 study/observation/evidence 수를 추가했다. 구조 coverage의 잘못된 `capture 실패 5건` 분류를 `table-request 실패 5건`으로 정정하고 실제 final coverage artifact와 설계 문서를 갱신했다.
- Decisions: 대상 361건은 모두 Capture v2와 canonical source DB에 존재한다. 356/5는 capture 성공/실패가 아니라 semantic table request 성공/실패다. canonical import는 AI 0회, retry 0이며 recipe가 없는 파일은 숫자 claim 없는 terminal 상태로만 계상한다.
- Files: `inference_data_ai_structure_canonical_import.py`, `inference_data_ai_incremental_coverage.py`, `tests/test_inference_data_ai_structure_canonical_import.py`, `tests/test_inference_data_ai_structure_reuse.py`, `STRUCTURE_REUSE_INCREMENTAL_PIPELINE_DESIGN.md`; runtime `D:\000. MyWorks\002. DB\InferenceDataAIService\structure-reuse\incremental-361-table-match-v1\final-coverage-report.json`, `canonical-import-v1\canonical-import-final-audit.json`; `HANDOFF.md`
- Verification: importer tests 2/2, structure reuse/coverage tests 19/19 및 관련 `py_compile`이 통과했다. 실제 final audit는 source/capture 361, active canonical 361, missing 0, multi-active 0, importer 246 analyses/120 studies/595 observations/1,430 evidence, canonical-import AI 0회를 기록한다.
- Next: query artifact와 최종 audit의 핵심 invariant를 자동 검사하고 canonical query 관련 좁은 테스트, diff check를 통과시킨 뒤 goal을 완료 처리한다.

## 2026-07-27 13:46 - 신규 361건 운영 canonical DB goal 최종 검증
- Completed: canonical final audit, query smoke, 구조 재사용/coverage/importer/query/CLI 회귀 테스트와 importer 소유 analysis 전체 무결성 검사를 완료했다. 신규 361건은 운영 DB에서 모두 active canonical analysis로 계상되고 기존 WPF/CLI canonical query 경로에서 신규 recipe 결과를 조회할 수 있다.
- Decisions: 최종 상태는 361건 전부 DB 포함이지만 의미 수준은 구분한다. 기존 canonical 115건은 보존, 신규 recipe-backed 54건은 program-owned observation/evidence가 있는 NEEDS_REVIEW, recipe 미확보 186건은 숫자 claim 없는 NEEDS_REVIEW terminal, semantic table-request 5건은 FAILED terminal, 1건은 EXCLUDED terminal이다.
- Files: `inference_data_ai_structure_canonical_import.py`, `inference_data_ai_incremental_coverage.py`, `tests/test_inference_data_ai_structure_canonical_import.py`, `tests/test_inference_data_ai_structure_reuse.py`, `STRUCTURE_REUSE_INCREMENTAL_PIPELINE_DESIGN.md`, `HANDOFF.md`; operational DB and final artifacts under `D:\000. MyWorks\002. DB\InferenceDataAIService\structure-reuse\incremental-361-table-match-v1`
- Verification: importer 2/2, structure reuse 19/19, query/CLI 52/52가 통과하고 관련 `py_compile`, focused whitespace/diff check가 통과했다. artifact invariant는 canonical 361/361, missing 0, multi-active 0, query `DATA-A6DB8C587B0A`/6 evidence, import AI 0회다. importer 소유 246 analysis는 FK ON read-only 재검사에서 246/246 `ok=true`; 신규 canonical 관련 FK 위반은 0이다. 전체 DB의 legacy FK 위반 66건은 `analysis_conclusions` 12, `analysis_evidence` 48, `analysis_review_items` 6으로 본 변경 범위 밖이다.
- Next: 현재 goal은 완료됐다. 후속으로 결과 품질 범위를 늘리려면 terminal NEEDS_REVIEW 186건과 큐 밖 재사용 가능 정량 355개를 추가 recipe 확대 대상으로 처리하되 현재처럼 구조당 최대 1회, retry 0, 값/통계/evidence program-owned 정책을 유지한다.

## 2026-07-27 19:27 - 신규 batch NEEDS_REVIEW 구성 감사
- Completed: canonical final audit의 361개 action과 운영 DB 상태를 기준으로 NEEDS_REVIEW 대상을 기존 분석, 신규 recipe-backed, 신규 recipe 미확보로 분리했다. 대상 361건 중 NEEDS_REVIEW는 총 354건이며 신규 246건 안에서는 240건이다.
- Decisions: 신규 recipe-backed 54건은 이미 120 studies/595 observations/1,430 evidence가 있으므로 semantic label, unit/percent scale, arm/comparison eligibility만 검토한다. 신규 terminal 186건은 검증 recipe가 없어 숫자 claim이 없으므로 반복 정량 구조 우선순위에 따라 recipe 생성·replay 검증이 먼저다. 기존 canonical 114건은 별도 기존 검토 queue다.
- Files: read-only source `D:\000. MyWorks\002. DB\InferenceDataAIService\structure-reuse\incremental-361-table-match-v1\canonical-import-v1\canonical-import-final-audit.json`; `HANDOFF.md`
- Verification: 신규 terminal 186건의 모델명 기준 분포는 X526 52, X526B 32, TIU 27, X626B 26, X626 25, S931B 6, S931 1, 기타 17로 합계 186이다. 신규 recipe-backed 54건은 기타 15, X526 14, TIU 7, S931B 6, X626 6, X626B 5, X526B 1로 합계 54다.
- Next: 사용자가 검토 진행을 요청하면 54건 semantic-only queue와 186건 recipe-required queue를 분리해 파일별 목록과 우선순위를 만들고, 186건은 재사용 빈도가 높은 구조부터 처리한다.

## 2026-07-27 19:30 - recipe 미확보 186건 해결 경로 정량화
- Completed: terminal NEEDS_REVIEW 186건의 826개 표를 기존 priority/completion/catalog 구조와 다시 매핑해 격리 queue, queue 밖 재사용 정량, 비재사용 정량, 비정량으로 분해했다. 186건 중 138건은 격리 구조 또는 재사용 가능한 반복 구조를 최소 하나 포함하고, 나머지 48건은 재사용 구조가 없는 단발성 정량 파일이다.
- Decisions: 186개 파일별 AI 재분석은 하지 않는다. 먼저 기존 AI 결정이 있으나 selector/contract가 실패한 구조를 code fix와 replay로 복구하고, 다음으로 queue 밖 반복 정량 semantic 구조만 구조당 최대 1회 AI 정책으로 recipe화한다. 단발성 구조는 generic code grammar 또는 명시적 terminal review로 남긴다.
- Files: read-only artifacts `table-structure-catalog.json`, `recipe-priority-report.json`, `completion-state.json`, `canonical-import-v1\canonical-import-final-audit.json`; `HANDOFF.md`
- Verification: 186건의 826개 표는 격리 queue 15 structures/40 tables/30 files, queue 밖 재사용 정량 39 semantic structures/152 tables/123 files, 비재사용 정량 323 structures/343 tables/159 files, 비정량 83 structures/291 tables/113 files로 분류됐다. 격리 15개는 contract failure 8/19 tables/19 files, explicit AI quarantine 4/16/7, AI transport failure 2/4/3, precheck 1/1/1이다. 격리와 queue 밖 재사용 파일의 합집합은 138건이다.
- Next: 실행 요청 시 1순위로 contract failure 8구조를 기존 decision 재사용·AI 0회로 수리하고, 2순위로 39개 반복 semantic 구조를 historical consensus→verified propagation→구조당 최대 1회 AI 순서로 recipe화해 canonical DB를 재적재한다.

## 2026-07-27 19:37 - [DATA][FIX] StructureRecipeRecovery: 중단 지점 및 복구 대상 재확정
- Agent: Codex
- Session: External
- Task-ID: Unavailable
- Category: DATA
- Feature: StructureRecipeRecovery
- Change: FIX
- Completed: 중단된 실행 흔적과 최신 canonical audit를 대조해 전체 completion 재실행이 아니라 기존 decision을 재사용하는 selector 복구가 재개 지점임을 확정했다. terminal 186건과 직접 겹치는 `QUARANTINED_RECIPE_CONTRACT_FAILURE`는 artifact 기준 7개 구조/16개 표/16개 파일이며, 별도 AI validation 실패 2개 중 하나는 `Invalid comparison relation`이다.
- Decisions: 이전 항목의 contract failure 8개/19개 표 표기는 `Invalid comparison relation` 1개 구조/3개 표를 합산한 분류로 보고, 실제 decision 파일이 존재하는 7개 selector 실패부터 AI 0회로 복구한다. `completion-state.json`은 이미 `COMPLETED`이므로 전체 completion runner를 다시 실행하지 않는다.
- Files: None
- Verification: `HANDOFF.md`, `completion-state.json`, `recipe-priority-report.json`, `table-structure-catalog.json`, `canonical-import-final-audit.json`을 읽기 전용으로 교차 대조했고 관련 Python/앱 프로세스가 실행 중이지 않음을 확인했다.
- Next: Capture v2 전체 셀에서 group label selector를 결정적으로 해석하도록 compiler/executor를 보정하고 기존 decision 7개를 개별 replay한다.

## 2026-07-27 19:40 - [DATA][FIX] StructureRecipeRecovery: Capture v2 selector 복구 구현
- Agent: Codex
- Session: External
- Task-ID: Unavailable
- Category: DATA
- Feature: StructureRecipeRecovery
- Change: FIX
- Completed: recipe compiler가 redacted preview에 없는 group label과 반복 label을 Capture v2 전체 셀에서 좌표로 해석하고, token 순서가 다른 복합 source header는 유일한 부분집합 match일 때만 fail-closed로 선택하도록 보정했다. executor도 선택된 group label을 Capture v2에서 읽어 authority를 기록하며 completion의 신규 compile 경로가 같은 resolver를 사용한다.
- Decisions: exact label이 여러 좌표에 있으면 구조 replay가 모든 member를 검증한다는 전제에서 가장 이른 좌표를 결정적으로 사용한다. fuzzy selector는 2개 이상 token인 label의 token 집합을 포함하는 source text가 유일할 때만 허용하고 동률 후보는 계속 격리한다.
- Files: `inference_data_ai_table_recipe_proposal.py`, `inference_data_ai_structure_completion.py`, `tests/test_inference_data_ai_structure_reuse.py`
- Verification: 구조 재사용/importer unittest 22/22와 관련 3개 모듈 `py_compile`이 통과했다. 기존 decision 7개를 별도 `.codex-verify` 경로에서 dry replay해 18/18 member, deterministic facts 142개, cell facts 900개가 전부 통과했고 신규 AI 호출은 0회였다.
- Next: 검증된 7개 recipe/replay를 운영 structure-reuse artifact에 반영해 registry를 재생성한 뒤 canonical importer dry-run과 단일 transaction 재적재를 수행한다.

## 2026-07-27 19:45 - [DATA][FIX] StructureRecipeRecovery: 운영 replay 및 canonical 재적재
- Agent: Codex
- Session: External
- Task-ID: Unavailable
- Category: DATA
- Feature: StructureRecipeRecovery
- Change: FIX
- Completed: 기존 AI decision 7개를 새 호출 없이 운영 recipe/replay로 등록하고 canonical DB를 단일 transaction으로 멱등 재적재했다. 검증 recipe는 56→63개, program extraction 표는 193→211개가 되었고 terminal recipe 미확보 파일은 186→170개로 감소했다.
- Decisions: 이전 8번째 집계였던 `Invalid comparison relation` 구조는 유효 decision artifact가 없어 이번 AI 0회 selector 복구에서 제외했다. 신규 결과는 계속 `NEEDS_REVIEW`로 유지하고 비교/effect answer 후보로 자동 승격하지 않는다.
- Files: `inference_data_ai_structure_completion.py`, `tests/test_inference_data_ai_structure_reuse.py`
- Verification: unittest 23/23, 관련 3개 모듈 `py_compile`, scoped whitespace/diff 검사가 통과했다. 운영 replay 18/18, coverage invariant 8/8, canonical active 361/361·missing 0·multi-active 0, importer integrity 246/246, 신규 AI 0회이며 복구 recipe의 SPL query에서 6개 Study가 모두 evidence와 함께 검색됐다.
- Next: queue 밖 반복 정량 semantic 구조 39개/152개 표를 historical consensus→verified propagation→남은 구조만 1회 AI 순서로 별도 증분 batch 처리한다.

## 2026-07-27 19:50 - [DATA][ADD] SingleMeasureRecipeExpansion: 증분 queue 생성 및 AI 0회 경로 감사
- Agent: Codex
- Session: External
- Task-ID: Unavailable
- Category: DATA
- Feature: SingleMeasureRecipeExpansion
- Change: ADD
- Completed: 기존 2개 이상 measure-column 안전 기준을 명시적 옵션으로 유지하면서 1개 measure-column 반복 표만 추가 허용하는 priority 확장 경로와 baseline 차집합 report를 구현했다. 실제 확장 queue는 31개 구조/325개 표/296 workbook references이며 모두 proposal-ready다.
- Decisions: 기본 동작은 `minimumMeasureColumns=2`로 유지하고 이번 별도 증분 pass에서만 1을 사용한다. 이전 terminal 중심 39 semantic 구조 집계와 실제 실행 가능한 확장 queue가 다르므로, replay 가능한 artifact 기준 31개 구조를 권위 대상으로 사용한다.
- Files: `inference_data_ai_table_recipe_proposal.py`, `tests/test_inference_data_ai_structure_reuse.py`
- Verification: 구조 재사용 unittest 23/23와 관련 `py_compile`이 통과했다. historical 989 consensus를 31개 확장 구조에 대조했으나 동일 semantic signature 관측이 0건이어서 AI 0회 신규 recipe는 없었고, 다음 단계는 기존 검증 recipe propagation 후 남은 구조의 구조당 최대 1회 AI다.
- Next: 기존 90개 outcome state와 합친 121개 queue를 completion runner로 재개해 기존 recipe propagation을 먼저 시도하고 남은 31개 구조를 retry 0 정책으로 처리한다.

## 2026-07-27 20:01 - [DATA][ADD] SingleMeasureRecipeExpansion: source-owned 확장 완료
- Agent: Codex
- Session: External
- Task-ID: Unavailable
- Category: DATA
- Feature: SingleMeasureRecipeExpansion
- Change: ADD
- Completed: 단일-measure 확장 31개 구조를 source-owned header grammar로 처리해 의미 있는 header 3개/11개 표는 AI 0회 recipe로 등록하고 generic `ME` header 28개 구조는 명시 격리했다. 합산 queue는 121/121 resolved이며 canonical DB의 recipe 미확보 terminal은 170→168개로 감소했다.
- Decisions: generic `ME` metadata 열은 parameter metric으로 강제 해석하지 않는다. 초기 확장 시 첫 4개 구조에서 같은 contract 실패가 발생하자 worker를 중지했고, 해당 4회 telemetry는 retry 없이 보존했으며 나머지 24개 generic 구조는 추가 AI 없이 source precheck로 닫았다.
- Files: `inference_data_ai_table_recipe_proposal.py`, `inference_data_ai_single_measure_bootstrap.py`, `inference_data_ai_structure_batch_control.py`, `inference_data_ai_structure_completion.py`, `tests/test_inference_data_ai_structure_reuse.py`
- Verification: 관련 unittest 27/27, 5개 모듈 `py_compile`, scoped whitespace/diff 검사가 통과했다. 최종 registry는 66 recipes/222 tables/171 workbook references, source-owned replay 11/11, coverage invariant 8/8, canonical active 361/361·missing 0·multi-active 0, importer integrity 246/246이다. 총 AI는 66회(성공 56/실패 10), retry 0/file-level 0이며 `Q'ty (pcs)` query에서 신규 2개 Study가 evidence와 함께 검색되고 answer-eligible effect는 0이었다.
- Next: 이번 재개 목표는 완료됐다. 후속 품질 확대 시 decision artifact가 없는 `Invalid comparison relation` 1개 구조를 먼저 별도 수리하고, 그 다음 recipe 미확보 terminal 168건은 단발성 정량/비정량 review queue로 분리한다.

## 2026-07-28 05:30 - [DATA][ADD] SingleMeasureRecipeExpansion: 중단 지점 감사
- Agent: Codex
- Session: External
- Task-ID: Unavailable
- Category: DATA
- Feature: SingleMeasureRecipeExpansion
- Change: ADD
- Completed: 중단 흔적을 대조해 19:53의 `RUNNING` 파일은 보존된 중간 스냅샷이고 실제 큐는 19:58에 121/121 `COMPLETED`, canonical import와 query smoke는 20:00에 완료됐음을 확인했다. 남은 후속 범위는 decision artifact가 없는 `Invalid comparison relation` 구조 1건과 verified recipe가 없는 terminal review 168건이다.
- Decisions: 완료된 전체 큐를 다시 실행하지 않는다. 다음 재개 시 `table-structure-df962eceef16da714d1b`의 비교관계 오류를 3개 표/2개 workbook 범위로 격리한 뒤, 별도로 168건을 단발성 정량/비정형 검토 큐로 분류한다.
- Files: None
- Verification: `HANDOFF.md`, `completion-state.json`, `single-measure-bootstrap-report.json`, `final-coverage-report.json`, canonical import/audit, telemetry와 프로세스 목록을 읽기 전용으로 대조했다. 큐 unresolved 0, canonical active 361/361, missing 0, multi-active 0이며 관련 실행 프로세스는 없었다.
- Next: `table-structure-df962eceef16da714d1b`의 실패 AI 출력/그룹 관계를 재현 가능한 fixture로 격리하고 validator 또는 결정 생성 경로를 보정한 뒤 해당 구조만 replay한다.

## 2026-07-28 05:40 - [DATA][FIX] StructureRecipeRecovery: 비교관계 결정 경로 보강
- Agent: Codex
- Session: External
- Task-ID: Unavailable
- Category: DATA
- Feature: StructureRecipeRecovery
- Change: FIX
- Completed: 비교관계가 선언된 group label의 대소문자나 공백만 다를 때 정확한 원본 label로 결정적으로 정규화하고, 중복 label은 거부하도록 validator를 보강했다. AI 프롬프트에는 relation이 두 개의 서로 다른 group label을 정확히 복사해야 한다는 제약을 추가했으며, 검증에서 거부된 AI 결정은 telemetry에 보존하도록 변경했다.
- Decisions: 의미가 다른 alias는 자동 보정하지 않고 정확 일치 또는 공백 정리·casefold로 유일하게 일치하는 참조만 허용한다. 검증되지 않은 결정 파일은 decision registry에 두지 않고 진단용 telemetry 내부에만 보존한다.
- Files: `inference_data_ai_table_recipe_proposal.py`, `tests/test_inference_data_ai_structure_reuse.py`
- Verification: 신규 회귀 테스트 3/3과 구조 재사용 모듈 테스트 28/28이 통과했다. 변경 Python 파일 `py_compile`과 scoped `git diff --check`도 통과했다.
- Next: `table-structure-df962eceef16da714d1b`에 새 경로로 bounded recovery decision 1회를 실행하고, 성공 시 3개 표를 deterministic replay한다.

## 2026-07-28 05:44 - [DATA][FIX] StructureRecipeRecovery: 비교구조 replay 복구
- Agent: Codex
- Session: External
- Task-ID: Unavailable
- Category: DATA
- Feature: StructureRecipeRecovery
- Change: FIX
- Completed: `table-structure-df962eceef16da714d1b`에 bounded recovery decision 1회를 실행해 X526TOP의 4개 test-bond 조건과 Normal 대조군, 10개 metric 축을 갖는 `NEW_RECIPE`를 생성했다. 3개 표/2개 workbook replay가 모두 통과했으며 registry는 67 recipes/225 tables/173 workbook references로 갱신됐다.
- Decisions: 기존 실패 telemetry는 삭제하지 않고 `telemetry-recovery-archive`로 이동해 보존하고 성공 telemetry만 활성 batch에 사용한다. 121구조 확장 상태는 `recipe-priority-report-min1.json`이 권위 파일이며, resume 시 이 경로를 state inputs에 다시 기록하고 현재 queue 밖의 과거 outcome은 summary에서 제외한다.
- Files: `inference_data_ai_table_recipe_proposal.py`, `inference_data_ai_structure_completion.py`, `tests/test_inference_data_ai_structure_reuse.py`
- Verification: recovery telemetry는 attempted/succeeded 1/1, retry 0이고 replay는 3/3, deterministic facts 30, cell facts 240, replay AI 0이었다. 구조 재사용 테스트 30/30, 관련 `py_compile`, scoped `git diff --check`가 통과했고 completion은 121/121 `COMPLETED`, unresolved 0, active AI attempted/succeeded/failed 66/57/9로 복구됐다.
- Next: canonical importer를 먼저 dry-run하여 terminal no-recipe가 168에서 감소하는지 확인하고, 무결성 검사가 통과하면 단일 transaction으로 운영 DB에 적용한 뒤 query smoke를 실행한다.

## 2026-07-28 05:47 - [DATA][FIX] StructureRecipeRecovery: canonical DB 복구 반영
- Agent: Codex
- Session: External
- Task-ID: Unavailable
- Category: DATA
- Feature: StructureRecipeRecovery
- Change: FIX
- Completed: 복구된 X526TOP 비교구조 recipe를 운영 `InputDataFinish.sqlite`에 단일 transaction으로 반영했다. recipe replay 대상은 72에서 74 workbook으로 늘고 no-verified-recipe terminal은 168에서 166으로 감소했으며, 새 Study 3건과 outcome/observation 30건이 canonical query 경로에 추가됐다.
- Decisions: 새 결과는 계속 `NEEDS_REVIEW`로 유지하고 비교/effect 답변 후보로 자동 승격하지 않는다. 121구조 확장 priority 파일을 coverage와 completion의 권위 입력으로 고정한다.
- Files: `inference_data_ai_table_recipe_proposal.py`, `inference_data_ai_structure_completion.py`, `tests/test_inference_data_ai_structure_reuse.py`
- Verification: dry-run, apply, postcheck가 모두 통과했고 active canonical 361/361, missing 0, multi-active 0이다. query smoke에서 `DATA-74F504F1C24D`, `DATA-1F714B71D094`, `DATA-23BE35F176B4`가 조회됐으며 answer-eligible effect는 0이었다. 구조/importer/query/CLI 테스트 84/84, 관련 `py_compile`, scoped diff check와 coverage invariant 8/8이 통과했다.
- Next: 남은 `NEEDS_REVIEW_NO_VERIFIED_RECIPE` 166건을 단발성 정량표, 비정형 정량표, 명시적 검토 필요 그룹으로 분류하는 재개 가능한 review queue artifact를 생성한다.

## 2026-07-28 05:54 - [DATA][ADD] TerminalReviewQueue: terminal 166건 검토 큐 생성
- Agent: Codex
- Session: External
- Task-ID: Unavailable
- Category: DATA
- Feature: TerminalReviewQueue
- Change: ADD
- Completed: canonical `NEEDS_REVIEW_NO_VERIFIED_RECIPE` 166 workbook을 숫자값이나 추가 AI 없이 표 단위로 구조 큐와 대조해 재개 가능한 human-review artifact를 생성했다. 118 workbook은 반복 정량구조 contract 검토, 48 workbook은 단발성/비재사용 정량 검토로 분류됐으며 전체 747개 표는 반복 격리 140, 단발성 정량 350, 비정량/보조 257로 빠짐없이 분류됐다.
- Decisions: registered recipe가 terminal 큐에 섞이거나 반복구조 outcome이 미결이면 fail-closed `INCOMPLETE`로 처리한다. 모델군은 Bottom/B/BT 표기를 동일한 B 계열로 정규화하고, 우선순위는 contract failure, AI failure, precheck, source precheck, explicit AI quarantine, 단발성 정량 순으로 정한다.
- Files: `inference_data_ai_terminal_review_queue.py`, `tests/test_inference_data_ai_terminal_review_queue.py`
- Verification: 신규 테스트 2/2와 구조 재사용 포함 32/32, 관련 `py_compile`, scoped `git diff --check`가 통과했다. 실제 `terminal-review-queue.json`은 `READY_FOR_HUMAN_REVIEW`, 7개 invariant 전부 true, 반복구조 group 31개, AI 0, numeric-values-read false이다.
- Next: 우선순위 1인 `table-structure-052ac0ae6948e2cbdca4-semantic-8ad80949`의 기존 decision과 replay contract failure를 격리해 AI 0으로 복구 가능한지 확인한다.

## 2026-07-28 06:02 - [DATA][FIX] SelectorRecipeRecovery: 전압 선택자 구조 canonical 복구
- Agent: Codex
- Session: External
- Task-ID: Unavailable
- Category: DATA
- Feature: SelectorRecipeRecovery
- Change: FIX
- Completed: `table-structure-052ac0ae6948e2cbdca4-semantic-8ad80949`의 기존 recipe를 현재 선택자 해석기로 재실행해 4개 workbook/4개 table과 56개 outcome을 복구했다. canonical DB 반영 후 no-verified-recipe terminal은 166건에서 164건으로 줄었고 recipe replay 대상은 74건에서 76건으로 늘었다.
- Decisions: `Normal voltage`, `130% voltage`, `120% voltage`는 각 원본의 `Total NG voltage ...` 헤더에 `SOURCE_TOKEN_SUBSET`으로 결정론적 매핑했으며 추가 AI 호출 없이 기존 decision을 재사용했다. 복구된 Study는 비교 효과로 자동 승격하지 않고 `NEEDS_REVIEW`를 유지했다.
- Files: None
- Verification: replay 4/4, facts 56, cell facts 176, replay AI 0을 확인했다. canonical dry-run/apply/postcheck 후 active canonical 361/361, missing 0, multi-active 0이었고 terminal review queue 164건의 invariant 7/7이 통과했다. DB에서 복구 Study 4건과 outcome 56건을 확인했으며 `DATA-C38C1B124139` 직접 evidence query가 1건을 반환했다. 관련 단위 테스트 32/32와 `py_compile`이 통과했다.
- Next: terminal review queue의 다음 반복 구조 그룹을 같은 방식으로 조사하되, `QUARANTINED_AI_FAILURE_NO_RETRY`는 새 AI 예산 승인 전까지 결정론적 복구 가능성만 우선 확인한다.

## 2026-07-28 06:08 - [DATA][FIX] TerminalReviewQueue: 비측정 메타데이터 오탐 제거
- Agent: Codex
- Session: External
- Task-ID: Unavailable
- Category: DATA
- Feature: TerminalReviewQueue
- Change: FIX
- Completed: source-owned 비측정 격리 decision과 날짜 전용 관리 블록을 terminal review queue에서 수치 계약 검토 대상이 아닌 supporting table로 재분류했다. 실제 artifact에서 source-owned 비측정 118개 표와 날짜 전용 메타데이터 7개 표를 식별해 반복 구조 검토 그룹을 30개에서 5개로 줄였다.
- Decisions: metric column이 없는 명시적 source-owned `QUARANTINE`만 supporting으로 인정하고, 날짜 전용 판정은 모든 숫자 열이 날짜 형식이며 Maker/Checker/Date/Signature 등 관리 라벨이 두 종류 이상일 때만 적용한다. 등록 recipe 충돌과 미결 outcome은 기존처럼 fail-closed를 유지한다.
- Files: `inference_data_ai_terminal_review_queue.py`, `tests/test_inference_data_ai_terminal_review_queue.py`
- Verification: 관련 단위 테스트 34/34와 `py_compile`이 통과했고 실제 CLI 재생성도 성공했다. 새 `terminal-review-queue.json`은 `READY_FOR_HUMAN_REVIEW`, invariant 7/7, terminal 164 workbook/739 table, supporting 378, one-off quantitative 344, repeated quarantined 17, repeated group 5, AI 0이다.
- Next: `table-structure-54c7e8bfe2dc13bfe044`는 고차원 SPL/THD/IMP 원자료라 자동 recipe 범위를 별도로 설계해야 한다. 나머지 4개 반복 그룹은 의미·단위·축이 불충분하다는 HIGH-confidence 명시적 격리이므로 사람 승인 없이 자동 복구하지 않는다.

## 2026-07-30 12:03 - [OTHER][ADD] ResumeContextInvestigation: 재개 대상 확인
- Agent: Codex
- Session: S1
- Task-ID: S1-20260730-120201338-6ff1db371ccd4eea9
- Category: OTHER
- Feature: ResumeContextInvestigation
- Change: ADD
- Completed: 프로젝트 handoff와 작업 문맥, Git 상태를 점검했으나 현재 대화에는 다시 수행할 원 요청이 남아 있지 않음을 확인했다. 점검 중 공유 작업 커밋 `71fbd75`가 생성되어 작업 트리는 깨끗한 상태로 전환됐다.
- Decisions: 원 요청을 추정해 기존 파일이나 새 커밋을 임의로 변경하지 않고 정확한 재작업 범위를 사용자에게 확인한다.
- Files: None
- Verification: `HANDOFF.md`, `WORKING_CONTEXT.md`, 최근 파일 시각, `git status`, `git log`를 읽기 전용으로 확인했으며 현재 `main`은 clean이고 `origin/main`보다 3커밋 앞선 상태다.
- Next: 사용자가 다시 수행할 기능·파일·오류 또는 직전 요청 문장을 알려주면 해당 범위부터 구현하고 검증한다.

## 2026-07-30 15:19 - [DATA][REMOVE] StructureRecipeRecovery: SPL/THD/IMP 원자료 제외
- Agent: Codex
- Session: S1
- Task-ID: S1-20260730-151940479-85d7285cde114437b
- Category: DATA
- Feature: StructureRecipeRecovery
- Change: REMOVE
- Completed: `table-structure-54c7e8bfe2dc13bfe044`의 SPL/THD/IMP 원자료를 후속 분석과 자동 recipe 복구 대상에서 제외했다.
- Decisions: 사용자 지시에 따라 해당 원자료는 구조 확인, recipe 설계, replay, canonical DB 반영을 모두 수행하지 않는다.
- Files: None
- Verification: 범위 제외 결정만 기록했으며 코드와 데이터는 변경하지 않아 검증 명령을 실행하지 않았다.
- Next: 해당 구조를 계속 제외하고, 사용자가 지정하는 다른 작업 대상으로 진행한다.
