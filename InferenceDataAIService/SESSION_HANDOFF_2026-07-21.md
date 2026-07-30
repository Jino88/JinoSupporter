# InferenceDataAIService session handoff — 2026-07-21

이 문서는 다음 Codex 세션이 현재 작업을 재실행 없이 이어가기 위한 재개 지점이다.

## 1. 현재 방향

- 기본 경로는 `TABLE_FIRST_SEMANTIC_PIPELINE.md`의 table-first 방식이다.
- 일반 workbook은 표/주변 텍스트를 코드로 압축하고 AI는 workbook당 1회만 호출한다.
- 숫자, 퍼센트 표시 동치, Min/Max/Average, PPM처럼 원본 값으로 계산 가능한 항목은 코드가 처리한다.
- 불확실한 표 하나가 workbook 전체를 차단하지 않는다.
- 기존의 workbook 전체 셀별 의미화 및 엄격 차단 경로는 기본 경로로 되돌리지 않는다.
- 사용자 지시 없이 WPF 또는 989건 전체 배치를 실행하지 않는다.

## 2. 대표 30건의 저장된 최종 상태

기준 파일:

`outputs/table-first/pilot-30-v3/batch-report.json`

2026-07-21 마지막 결정적 재검사 결과:

- 상태: `ok`
- 30/30 성공, 실패 0
- 기존 analysis 30건 재사용, AI 호출 0회
- 재검사 및 projection 갱신: 19.532초
- table 282개, text block 212개
- confidence: HIGH 196, MEDIUM 85, LOW 1
- raw Min/Max/Average 검산: 65/65 일치
- 검수 권장 workbook: 1건

남은 한 건:

- 파일: `026.MSU-20S15-07 GMI -Result check test bonding VP+ CD improve NG function modul_clean.xlsx`
- 표: `RAW DATA!H202:DW296`, 제목 `IMP`
- 사유: 광폭 IMP RAW DATA의 STD/Normal 의미 경계, 단위 미표기, `#DIV/0!` 및 불완전 trailing block
- 이 항목은 숫자 불일치가 아니라 의미 경계의 LOW confidence다.

`Reliability` workbook의 명시적 RAW DATA 및 부분 수식 파생 상태는 사용자 판단에 따라 별도 사람 검수 대상으로 올리지 않는다.

## 3. 사용자 판단으로 확정된 처리 원칙

- `0.22`와 표시값 `22%`는 같은 값으로 처리한다.
- `Min`, `Max`, `Average`는 raw data가 있으면 코드로 계산·검산하며 학습 대상이 아니다.
- `Check Height of Sample` 같은 표 제목/머리글은 측정값으로 취급하지 않는다.
- 병합 범위 `E33:F33`에 각각 `10, 10`이 있으면 두 원본 값을 보존한다.
- PPM은 원본 `NG`와 `Input`이 있으면 코드 계산을 허용한다.
- 비교군/대조군 또는 단일 지표를 뽑는 실용적 분석이 목표이며, 과도한 검증 때문에 전체 처리가 느려지면 안 된다.
- 검수가 필요할 때는 우선 CLI에서 문제, Excel 실제값, DB 인식값을 직접 제시한다. WPF를 자동으로 띄우지 않는다.

## 4. 발견한 기존 도메인 용어 사전

별도 프로젝트에 기존 마이크로스피커 용어 사전이 있다.

주 파일:

`D:/000. MyWorks/005. Program/Repository/MicroSpeaker_ProductTech_DB/db/term_dictionary.csv`

관련 파일:

- `db/term_dictionary_auto_export.csv`
- `db/term_dictionary_editor.html`
- `db/term_unknown_candidates_final_review.csv`
- `HANDOFF_2026-07-01.md`

확인 내용:

- `term_dictionary.csv`: 199개 용어
- 미확정 후보: 100개
- `F0`, `Fo`, `FO`는 같은 공진주파수로 명시
- `VP-CD`, `VP+CD`, `VP/CD`는 같은 조립체로 명시
- `PPM = NG / Input * 1,000,000`으로 명시
- `VP`, `CD`, `SPL`, `Tension`, `Bonding` 등 현재 workbook에 나오는 용어가 정의돼 있음
- `#REF!` 조각은 분석 제외 대상으로 정의돼 있음

중요한 보류점:

- 사전에는 `SPL` 정의는 있으나 `SPL Average` 별도 정의는 없다.
- 현재 table-first 문서/코드에서 `SPL`과 `SPL Average`를 같은 계열로 정규화한 부분은 사전을 연결하기 전에 다시 검토해야 한다.
- `Fo/F0/FO` 동치는 기존 사전으로 확인됐지만, 다른 약어의 동치를 추정해서 추가하면 안 된다.

## 5. 다음 세션의 첫 작업

1. 이 문서와 `TABLE_FIRST_SEMANTIC_PIPELINE.md`를 읽는다.
2. 기존 `term_dictionary.csv`를 읽기 전용 기준 사전으로 table-first 분석에 연결할 최소 adapter 위치를 찾는다.
3. `SPL`/`SPL Average` 정규화는 원본 표 제목과 축을 보존하도록 재검토한다.
4. 용어 사전에 정의된 alias와 IGNORE 항목만 결정적 규칙으로 적용한다.
5. 좁은 단위 테스트 후 대표 30건의 기존 analysis를 재사용하여 AI 호출 없이 재검사한다.
6. 결과가 유지되는지 확인한 뒤에만 확대 적용을 논의한다.

## 6. 실행 및 작업 트리 상태

- 현재 실행 중인 corpus/table-first 배치나 WPF 프로세스는 없다.
- 대표 30건 request/analysis/projection/batch report는 모두 `outputs/table-first/pilot-30-v3`에 저장되어 있다.
- Git 저장소는 `JinoSupporter` 전체에 사용자 작업을 포함한 대량의 미커밋 변경이 있다.
- 이번 세션에서는 해당 변경들을 commit, revert, clean하지 않았다.
- 이 서비스의 현재 수정/추가 파일과 `outputs/`를 그대로 보존해야 한다.

## 7. 2026-07-21 마지막 사용자 지시 — 주파수별 측정 DATA는 AI 분석 제외

사용자가 다음 규칙을 새 기본 계약으로 확정했다.

- 주파수별로 값이 나열된 측정 원시 DATA는 AI 의미분석 대상에서 제외한다.
- 대표 대상은 `SPL`, `THD`, `IMP`/`IMPEDANCE` 주파수 응답 DATA다.
- Function NG 요약표에서 `SPL`, `THD`, `SPL+THD`, `SPL+THD+F0` 등의
  불량 개수를 비교하는 표는 원시 주파수 DATA가 아니므로 제외 대상이 아니다.
- `SPL Average`, `F0`, Min/Max/Average, NG/Input/PPM처럼 코드로 계산하거나
  검산할 수 있는 요약값은 기존 결정적 코드 경로를 유지한다.
- 제외 대상 원시 DATA는 삭제하지 않는다. 원본 좌표와 값은 보존하되 AI prompt,
  AI table 분류, AI metric/group/comparison 생성에는 전달하지 않는다.
- workbook에 AI 분석 대상 표가 하나도 남지 않으면 AI를 호출하지 않는다.

이번 세션에서는 요구사항과 구현 위치만 조사했고 코드 변경은 시작하지 않았다.
확인한 주요 원시 DATA 예시는 다음과 같다.

- `SPL DATA_(NTI)!A2:AB90`
- `THD DATA_(NTI)!A2:AA89`
- `IMP DATA_(NTI)!A2:Z89`
- `Sheet2!A1:CQ32`처럼 `100.00Hz`, `106.00Hz`, ... 열이 이어지는 측정표
- Reliability workbook의 `SPL RAW DATA`, `THD RAW DATA`,
  `IMPEDANCE RAW DATA` 주파수 곡선 영역
- 기존 최종 검수 대상이던 `RAW DATA!H202:DW296`, 제목 `IMP`도 이 규칙에
  따라 AI 검수 대상에서 빠지는지 결정적 분류기로 확인한다.

다음 세션의 구현 순서:

1. `inference_data_ai_table_first.py`에 주파수 축과 SPL/THD/IMP 계열을 함께
   확인하는 보수적인 원시 측정 DATA 판별기를 추가한다.
2. 판별된 표는 `request["tables"]`에서 빼고 별도 code-owned 제외 inventory에
   보존한다. AI에 전달하는 JSON에는 원시 측정값과 주파수 배열을 포함하지 않는다.
3. 요약 불량률 표와 `SPL Average`/`F0` 요약 행이 잘못 제외되지 않는 회귀
   테스트를 추가한다.
4. 원시 주파수 DATA만 있는 workbook은 기존 `NO_TABLES`/AI 0회 경로를
   사용하도록 검증한다.
5. 좁은 테스트는 우선
   `python -m unittest tests.test_inference_data_ai_table_first`로 실행한다.
6. 대표 30건의 저장된 analysis를 재사용할 수 있는지 확인하고, 가능하면 AI
   재호출 없이 request/projection/batch report를 다시 만든다. 사용자 지시 없이
   WPF나 989건 전체 배치를 실행하지 않는다.

현재 기준 회귀 상태:

- 코드 수정 전 `python -m unittest tests.test_inference_data_ai_table_first`
  실행 결과: 20/20 PASS.
- 앱, WPF, 서버, 989건 배치는 실행하지 않았다.

## 8. 후속 구현 및 대표 30건 실행 완료

같은 날 후속 작업에서 `table-first-builder-v4` 구현과 검증을 완료했다.

- 세로/가로 주파수 축과 원본 SPL/THD/IMP 계열 표기를 함께 요구하는 보수적
  원시 응답 DATA 판별기 추가
- 제외 표는 `codeOwnedExclusions.rawFrequencyResponseTables`에 source packet
  참조 metadata로 보존
- code-owned 제외 inventory는 AI prompt JSON에서 제거
- Function NG 요약표, `SPL Average`, `F0` 비제외 회귀 테스트 추가
- 좁은 테스트: `python -m unittest tests.test_inference_data_ai_table_first`
  결과 25/25 PASS
- 대표 30건 최종 출력:
  `outputs/table-first/pilot-30-v6-frequency-exclusion`
- 30/30 성공, 실패 0, 저장 analysis 30건 재사용, AI 호출 0회
- 7개 workbook의 원시 응답표 17개 제외, AI 대상 표 265개
- 검수 권장 0건

`pilot-30-v5-frequency-exclusion`은 analysis 사전 복사 실패를 발견해 중단한
불완전 실행이므로 기준 결과로 사용하지 않는다. 기준은 v6 디렉터리다.

## 9. 도메인 용어 사전 adapter 연결 완료

후속 작업에서 `table-first-builder-v5`와 읽기 전용 용어 사전 adapter를
구현했다.

- 기준 사전:
  `D:/000. MyWorks/005. Program/Repository/MicroSpeaker_ProductTech_DB/db/term_dictionary.csv`
- `DEFINED` 179개, `IGNORE` 19개, 미확정 1개
- 동일 `normalized_name`의 복수 DEFINED 표기에서 alias 그룹 19개 생성
- `F0`/`FO`, `VP-CD`/`VP+CD`/`VP/CD` 등만 명시 alias로 적용
- IGNORE는 사전 정규화에서만 제외하고 source-authored group/metric은 보존
- `SPL`과 `SPL Average`는 서로 다른 metric key로 유지
- 사전 snapshot은 request에 code-owned로 보존하지만 AI prompt에는 전달하지 않음
- 좁은 테스트 26/26 PASS
- 최종 기준 출력:
  `outputs/table-first/pilot-30-v7-term-dictionary`
- 30/30 성공, analysis 30건 재사용, AI 호출 0회, 실행 7.254초
- v6 대비 table semantic 결과 30/30 동일, 요약 metric 이름·축 5/5 일치

현재 전체 기준 결과는 v7이다. 사용자 지시 없이 WPF, 서버, 989건 전체 배치를
실행하지 않는다.

## 10. 989건 AI-free request 점검 완료

사용자의 명시적 지시로 989건 전체에 대해 AI를 호출하지 않는 code-only 점검을
실행했다. WPF, 앱, 서버는 실행하지 않았다.

- 읽기 전용 Capture v2 DB current revision: 989/989
- semantic packet 출력:
  `outputs/semantic-source-packets/full-989-v1`
- semantic packet 생성: 989/989 성공, 실패 0, 7,787 chunks,
  1,505,559 cells, 85.3초
- 새 CLI: `table-first-request-batch`
  - request 생성·재사용과 code-owned audit만 수행
  - AI 분석 함수는 호출하지 않음
  - NO_TABLES, request 크기, 원시 주파수 응답 제외, 수식 오류,
    aggregate mismatch를 workbook별로 기록
- 좁은 테스트: `python -m unittest tests.test_inference_data_ai_table_first`
  결과 27/27 PASS
- 최종 점검 보고서:
  `outputs/table-first/full-989-v1-request-audit/request-batch-report.json`
- 989/989 성공, 실패 0, 두 번째 실행 request 989/989 byte-for-byte 재사용,
  AI 호출 0회
- table 5,546개, text block 6,251개
- NO_TABLES 9건
  - 빈 workbook 6건
  - 메모/이미지 중심 2건
  - 23,434 cells 전부 raw SPL 표라 code-owned 제외된 workbook 1건
- 원시 SPL/THD/IMP 응답표 제외: 77 workbook, 177개 표, 783,086 cells
  - 상위 제외 표본은 `RAW DATA`, `SPL DATA_(NTI`, `THD DATA_(NTI`,
    `IMP DATA_(NTI` 등 의도한 원시 주파수 응답 영역과 일치
- request 크기: median 30,183 bytes, p95 101,449, p99 212,043,
  최대 459,854; 240,000 bytes 초과 9건
- formula derivation error: 11 workbook, 6,165개
  - 상위 3 workbook이 5,762개를 차지
  - 주요 원인은 cross-sheet 직접 참조, 문자열 연결, `#DIV/0!`, `#REF!`
- aggregate check: MATCH 1,490, MISMATCH 310; mismatch workbook 57건
- 용어 사전: 989/989 `LOADED`, 동일 SHA-256
  `537e330e3b66b12dced13546ef75dc6ad783afa6483de7451426bc2c255706bf`
- 최종 code-only 검토 대상: 중복 포함 83 workbook

다음 우선순위는 AI 전체 분석이 아니다. 먼저 57개 workbook의 aggregate
mismatch 원인을 표본 검증하고, 11개 workbook의 formula unsupported/error를
cached value 사용 또는 지원 범위 확대로 분리한 뒤, 9개 oversized request의
분할/압축 정책을 정한다. 이 세 항목이 정리되기 전에는 980건 AI 분석을
실행하지 않는다.

## 11. Aggregate mismatch 점검 및 builder v6 완료

989건 1차 audit의 57 workbook/310 aggregate mismatch를 원본 셀까지 추적해
범위 선택 규칙을 수정했다.

- builder: `table-first-builder-v6`
- 복수 Min/Max/Avg triplet 독립 분리
- `Total NG`, `Possition`, `Sample` 번호, 공정 basis 열 제외
- 다음 summary row 직전까지 세로 block 경계 제한
- 단일 raw 열의 세로 집계 지원
- 떨어진 숫자 영역으로 넘어가는 aggregate fallback 제거
- 좁은 테스트 30/30 PASS
- 기준 출력:
  `outputs/table-first/full-989-v3-aggregate-audit`
- 989/989 성공, 실패 0, AI 호출 0회
- aggregate MATCH 1,743, MISMATCH 136
- mismatch workbook 57개에서 29개로 감소
- mismatch 310개에서 136개로 56.1% 감소
- 잔여 136개 중 61개는 Min/Max 일치, Average만 원본과 다름
- 나머지 75개는 Min/Max도 달라 자동 승인하지 않고 검수 대상으로 유지

다음 우선순위는 formula derivation error 11 workbook이다. aggregate 잔여 29개는
검수 대상으로 보존하며 source-authored 값을 수정하거나 덮어쓰지 않는다.

## 12. Formula derivation v2 및 builder v7 완료

6,165개 formula error를 전수 분류한 뒤 숫자를 발명하지 않는 범위에서 평가기를
확장했다.

- formula overlay schema: `deterministic-formula-overlay-v2`
- evaluator: `restricted-a1-arithmetic-v2`
- builder: `table-first-builder-v7`
- 같은 workbook의 직접 cross-sheet A1 참조를 정확한 sheet 이름으로만 지원
- 직접 text 참조, 제한된 text 연결식, `IF(AND(...), "OK", "NG")` 형태는
  숫자 오류가 아닌 `NON_NUMERIC`으로 분리
- 존재하지 않는 sheet, cross-sheet range, `#REF!`, `#DIV/0!`, 고급 함수는
  추정하거나 보정하지 않고 fail-closed 유지
- 좁은 회귀 테스트 42/42 PASS
- 기준 출력:
  `outputs/table-first/full-989-v4-formula-audit`
- 최종 보고서:
  `outputs/table-first/full-989-v4-formula-audit/request-batch-report.json`
- 989/989 성공, 실패 0, AI 호출 0회
- 결정성 재검사에서 request 989/989 byte-for-byte 재사용
- 전체 formula 13,995개 중 숫자 파생 10,900개, 비숫자 분류 1,130개,
  error 1,965개
- formula error는 6,165개에서 1,965개로 4,200개(68.1%) 감소
- formula error workbook은 11개에서 9개로 감소
- 잔여 1,965개:
  - 원본 `#DIV/0!` 1,209개
  - `SUMIFS`/`COUNTIFS`/`SUMPRODUCT`/`INDEX`/`MATCH` 계열 618개
  - 실제 없는 sheet 이름 또는 공백이 다른 sheet 참조 90개
  - `#REF!`와 그 의존 수식 31개
  - cross-sheet range 13개
  - 잘못 수식으로 저장된 text 2개와 bare range 산식 2개

잔여 오류는 source 오류이거나 현재의 제한 문법 밖이다. 자동으로 source formula나
sheet 이름을 고치지 않는다. 다음 code-only 우선순위는 240,000 bytes를 넘는
request 10건의 크기 원인 분류와 안전한 분할/압축 정책이다.

## 13. Prompt v4 크기 압축 및 240KB 초과 해소

저장 request 크기와 실제 AI 입력 크기를 분리 계측했다. 저장 request 10건이
240,000 bytes를 넘었지만 반복 template 압축까지 적용한 기존 실제 prompt 기준
초과는 9건이었다.

`table-first-analysis-prompt-v4`에서는 Capture v2와 저장 request를 바꾸지 않고
AI에 보내는 의미 투영만 압축했다.

- `aggregateChecks` 상세 숫자는 status count 요약으로 대체
- numeric column의 Min/Max/Average, count, source range 등 code-owned 상세 제거
- percent 표시 예와 identifier/basis 값은 유지
- `rowLabels`에 그대로 존재하는 preview text 중복 제거
- measure/aggregate 숫자는 preview에서 제거하되 원본 request에는 전부 보존
- 날짜, identifier/basis 숫자, 비중복 문맥, table/row/column ID는 유지
- workbook당 AI 1회 원칙을 유지했으며 분할 호출은 추가하지 않음

좁은 회귀 테스트는 43/43 통과했다. 기준 감사 출력은 다음과 같다.

`outputs/table-first/full-989-v5-prompt-audit/request-batch-report.json`

- 989/989 성공, 실패 0, AI 호출 0회
- 결정성 재검사 request 989/989 byte-for-byte 재사용
- 실제 prompt 총량 39,465,194 → 21,179,656 bytes, 46.3% 감소
- 실제 prompt 최대 436,098 → 215,729 bytes
- 240,000 bytes 초과 prompt 9건 → 0건
- median 17,103, p95 47,200, p99 101,728 bytes
- 저장 request는 감사·provenance 보존 때문에 최대 459,874 bytes와 초과 10건을
  그대로 기록하지만 실제 AI 입력 초과와 구분한다.

aggregate 잔여 29 workbook과 formula error 잔여 9 workbook은 이미 자동 승인하지
않는 진단/검수 대상으로 격리했다. 다음 단계는 AI 전체 실행이 아니라 새 prompt
v4로 실행할 pilot 범위와 비용을 정하는 것이다. 사용자 명시 없이 AI 호출을
실행하지 않는다.

## 14. 사용자용 table-first HTML 리포트 구현

AI 분석 JSON을 사용자가 직접 탐색할 수 있는 정적 HTML로 변환하는 파이프라인을
추가했다.

- 모듈: `inference_data_ai_table_first_html.py`
- CLI: `table-first-html`
- 통합 `index.html`
  - workbook 파일명 검색
  - 분석 상태와 검수 여부 필터
  - workbook/연구/표/지표/비교/검수 건수 요약
- workbook 상세 HTML
  - AI workbook 요약과 메모
  - Study별 시험군×지표 통합표
  - 왼쪽 첫 열에 시험군 역할과 비교 대상을 함께 표시
  - 각 지표 열에 시험군별 실제 값 또는 안전하게 연결된 count/Min/Max/Average 표시
  - percent 원본 scale을 반영한 표시
  - 코드 집계 검산과 제한 사항
  - 원본 표 유형·신뢰도·시험군·지표
- sheet/range와 셀 위치는 원본 JSON에 보존하지만 사용자용 HTML 표에는 표시하지 않음
- 시험군과 지표의 연결을 확정할 수 없는 셀은 다른 통계를 복제하지 않고 `-`로 표시
- 외부 CDN이나 서버가 필요 없는 독립 정적 HTML
- HTML escaping, artifact requestId 일치, 링크 무결성, 결정성 검사 포함

기존 prompt v3 대표 30건 분석으로 실제 HTML을 생성했다.

`outputs/table-first/pilot-30-v7-term-dictionary/html-report/index.html`

- 통합 index 1개, workbook 상세 30개
- 내부 workbook 링크 30/30 유효
- 재실행에서 HTML/manifest 32개 전부 byte-for-byte 재사용
- 관련 좁은 테스트 전체 45/45 PASS
- WPF와 서버는 실행하지 않음. 사용자의 명시적 요청으로 생성된 index를 기본 브라우저에서 열었음

재생성 명령:

```powershell
python inference_data_ai_cli.py table-first-html `
  --batch-dir outputs/table-first/pilot-30-v7-term-dictionary
```

현재 HTML은 기존 30건의 실제 AI 분석 결과를 보여주는 완성된 표시 골격이다.
다음 단계는 prompt v4 실제 AI 5건을 실행해 같은 HTML로 내용 품질까지 확인한 뒤,
통과하면 나머지 975건을 실행하고 980건 통합 HTML을 생성하는 것이다.

## 15. 최종 전체 이력 질의 경로 실행 중

사용자가 최종 목표까지 진행하도록 명시적으로 요청해 위의 이전 중단 조건을
해제했다. `table-first-builder-v7`과 `table-first-analysis-prompt-v4`를 동결하고
대표 5건을 실제 분석한 뒤 전체 989건 배치를 시작했다. 대표 5건은 5/5 성공,
표 47개, HIGH 37개, MEDIUM 10개, LOW 0개, formula error 0개였다.

전체 배치 경로와 재개 명령은 다음과 같다. 성공한 artifact는 재사용되므로 중단 후
같은 명령을 실행해도 완료 건을 AI로 다시 분석하지 않는다.

```powershell
python inference_data_ai_cli.py table-first-batch `
  --packet-dir outputs/semantic-source-packets/full-989-v1 `
  --out-dir outputs/table-first/full-989-v8-prompt-v4 `
  --workers 3 --reasoning-effort low --timeout 900
```

- batch report: `outputs/table-first/full-989-v8-prompt-v4/batch-report.json`
- stdout: `outputs/run-logs/full-989-v8-prompt-v4.resume.stdout.log`
- stderr: `outputs/run-logs/full-989-v8-prompt-v4.resume.stderr.log`
- 이 handoff 작성 시 hidden Python root PID: `25848` (`Get-Process -Id 25848`로 확인)
- 완료 조건: status `ok`, completed/succeeded 989, failed 0

PID가 없고 report가 `running`이면 같은 명령으로 재개한다. 완료된 request/analysis/
projection은 검증 후 재사용되며 `--force`는 사용하지 않는다.

완료 뒤 `table-first-history-index`로
`outputs/table-first-history/history.sqlite`를 만들고,
`table-first-history-acceptance`로 대표 golden question 10개가 필수 원본을 검색하는지
검사한다. 이 검색·답변·수용검증은 저장 JSON만 사용하며 AI를 다시 호출하지 않는다.

사용자의 최신 지시에 따라 전체 corpus 정적 HTML은 생성하지 않는다. 사용자 결과물은
WPF 기존 질문 화면의 이력 검색·근거 답변과 `TF-EVD-*` 원본 상세보기다. WPF 빌드는
검증하지만 앱 자체는 실행하지 않는다.

## 16. 최종 전체 분석·이력 질의 완료

위 섹션 15의 실행 중 상태를 대체한다. 전체 batch는 최종 `ok`, 989/989 성공,
실패 0으로 완료됐다. 첫 순회의 AI 응답 참조 오류 2건은 성공 987건을 그대로
재사용한 뒤 해당 2건만 재호출해 해결했다.

- batch report: `outputs/table-first/full-989-v8-prompt-v4/batch-report.json`
- builder/prompt: `table-first-builder-v7` / `table-first-analysis-prompt-v4`
- 표 없음: 9 workbook
- 추출된 표/Study/근거: 5,546 / 3,710 / 5,546
- 검수 권고: 147 workbook
- formula error가 격리된 workbook: 9개, 1,965건
- 최종 검색 DB: `outputs/table-first-history/history.sqlite`
- DB: workbook 989, Study 3,710, evidence 5,546, term 179
- SQLite quick check `ok`, foreign-key 오류 0
- 수용검증: `outputs/table-first-history/history-acceptance-report.json`
- golden question 10/10 PASS, primary source 15/15 검색, 누락 0
- 대표 답변: `outputs/table-first-history/final-vp-bonding.answer.md`
- 대표 상세: `outputs/table-first-history/final-vp-bonding.detail.json`

검색은 질문마다 관련 workbook 최대 30개를 확보하고 상위 12개를 자세히 표시한 뒤
나머지는 한 줄 원본 목록으로 축약한다. 모든 table-first 결과는 사람 검토 전이므로
승인된 effect는 0으로 유지하고 원본 workbook/sheet/range 근거를 연결한다.

최종 좁은 회귀 테스트는 61/61 PASS, WPF 빌드는 경고 0·오류 0이다. 앱/서버는
실행하지 않았고, 사용자 지시에 따라 전체 corpus HTML도 생성하지 않았다.
