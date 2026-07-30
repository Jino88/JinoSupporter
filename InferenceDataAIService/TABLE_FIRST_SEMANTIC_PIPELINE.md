# Table-first Semantic Pipeline

사용자 확인: 2026-07-20

이 문서는 `InferenceDataAIService`의 기본 분석 경로를 정의한다. 기존의 모든 비어 있지 않은 셀을 AI가 각각 의미화하고, 모든 불확실성을 workbook 전체 차단 사유로 처리하는 방식은 기본 경로로 사용하지 않는다. 그 경로는 정밀 진단이나 특정 실패 건의 보조 도구로만 남긴다.

## 1. 목표

Excel workbook 하나를 다음처럼 처리한다.

1. 코드가 표와 주변 텍스트를 찾고 원본 값과 위치를 보존한다.
2. AI는 workbook당 압축된 입력을 한 번 받아 표의 제목, 비교 대상, 지표, 표 사이 관계만 판정한다.
3. 코드는 AI가 지정한 표와 열을 이용해 원본 값, 비율, Min/Max/Average 등 계산 가능한 값을 확정한다.
4. 유사한 표는 하나의 study로 묶고, 의미가 다른 표는 독립 study로 저장한다.
5. 불확실한 부분은 `NEEDS_REVIEW` 또는 `PARTIAL`로 남기되 다른 정상 표의 저장을 막지 않는다.

최종 산출물은 workbook별 통합 분석과 정확한 Excel 근거 위치이다. 셀별 AI 설명은 산출물이 아니다.

## 2. 표 분류

각 표는 다음 중 하나로 분류한다.

- `COMPARISON`: 둘 이상의 군, 조건, 시점, 재료, 공정 등을 비교하는 표
- `DESCRIPTIVE`: 대조군 없이 단일 지표나 측정 결과를 기록한 표
- `SUPPORTING`: 시험 조건, 규격, 장비, 공정 조건처럼 다른 결과 표를 설명하는 표
- `TEXT`: 표 밖의 결론, 목적, 시험 조건, 주석

`Control`, `Normal`, `Old`, `Before` 같은 이름을 억지로 표준 대조군으로 바꾸지 않는다. 원문 label을 보존하고, 표 안에서 명백한 경우에만 역할을 추가한다.

## 3. AI의 책임

AI는 압축된 workbook inventory를 입력받아 다음 항목만 반환한다.

- `tableId`: 입력에 있던 식별자만 참조
- `title`: 표 제목
- `type`: `COMPARISON`, `DESCRIPTIVE`, `SUPPORTING`, `TEXT`
- `studyGroup`: 같은 목적과 비교 구조를 가진 표를 묶는 식별자
- `groups`: 표에 적힌 원래 군/조건 이름과 선택적 역할
- `metrics`: 측정 지표 이름, 단위, 해당 열 또는 행
- `comparisonRelation`: 무엇과 무엇을 비교하는지
- `textLinks`: 관련 목적, 조건, 결론 텍스트 식별자
- `relatedTableIds`: 유사하지만 합치기 불확실한 표
- `confidence`: `HIGH`, `MEDIUM`, `LOW`
- `limitations`: 불확실하거나 빠진 의미

AI는 다음을 하지 않는다.

- 원본 숫자를 다시 타이핑하거나 계산 결과를 생성하지 않는다.
- 모든 셀에 의미·처분 상태를 붙이지 않는다.
- 근거가 없는 대조군, 단위, 인과관계를 추정하지 않는다.
- 표가 모호하다는 이유로 workbook 전체를 실패 처리하지 않는다.

## 4. 코드의 책임

Python/C# 코드는 다음을 결정적으로 처리한다.

- 원본 셀 값, 표시 값, 수식, 병합 범위, 시트와 주소 보존
- 숫자와 퍼센트 표시의 동치 처리 (`0.22`와 `22%`)
- 동일 raw data에서 Min/Max/Average 및 기타 명시적 통계 계산
- 병합 셀을 펼치지 않고 표 좌표와 header 구조 보존
- 표 제목처럼 보이는 문자열을 raw 측정값과 분리
- AI가 참조한 `tableId`, 열, 행, text id가 실제 입력에 존재하는지 검증
- 모든 결과를 정확한 workbook/sheet/range에 연결

표에 이미 적힌 Min/Max/Average는 원문 표시값으로 보존할 수 있지만 별도 학습 대상이 아니다. raw data가 있으면 코드는 재계산하여 비교한다.

## 5. 표 묶기 규칙

같은 workbook 안에서 다음이 일치하면 하나의 `studyGroup` 후보로 묶는다.

- 시험 목적 또는 결과 제목
- 비교하는 군/조건
- 지표와 단위
- 모델, Lot, 기간, 공정 등 중요한 context

하나라도 중요한 차이가 있거나 판단이 모호하면 독립 study로 둔다. 대신 `relatedTableIds`로 연결할 수 있다. 서로 다른 workbook 사이의 취합은 ingestion 단계가 아니라 검색·related-study 계층에서 수행한다.

## 6. 검증과 상태

기본 원칙은 **넓게 저장하고, 좁게 계산한다**이다.

- `VERIFIED`: 표 구조, 값, 근거 위치가 확정되고 필요한 비교 조건도 충족
- `NEEDS_REVIEW`: 값과 근거는 저장 가능하지만 역할이나 study 묶음이 불확실
- `PARTIAL`: workbook 일부 표만 해석 가능
- `EXCLUDED`: 분석 대상이 아닌 표/텍스트이며 이유가 있음
- `FAILED`: 파일을 열거나 구조를 읽지 못해 원본 확보 자체가 실패

대조군이 없어도 descriptive metric은 저장한다. 다만 비교 효과나 인과 결론만 생성하지 않는다. 단위, 분모, 시점 등이 맞지 않으면 해당 비교 계산만 보류하며 다른 표를 막지 않는다.

## 7. 성능 예산

- 일반 workbook: AI 호출 1회
- 매우 큰 workbook: 최대 2~3회 분할 후 1회 통합, 예외 사유 기록
- AI 입력: 전체 셀 dump 금지. 제목, header, row label, 숫자 열의 구조·개수·샘플·범위, 주변 텍스트만 전달
- workbook 전체를 같은 prompt로 재시도 금지. 실패한 표 또는 응답 형식만 국소 재시도
- AI를 사용하지 않아도 raw 표와 text inventory는 DB에 저장 가능해야 함

운영 측정값으로 workbook별 다음을 기록한다.

- 발견 표 수와 text block 수
- AI 호출 수
- prompt byte 수
- 구조 추출 시간, AI 대기 시간, 코드 계산 시간
- `VERIFIED`, `NEEDS_REVIEW`, `PARTIAL`, `EXCLUDED`, `FAILED` 개수

현재 staged draft의 수십 회 AI 호출과 수 MB 누적 prompt는 이 예산을 충족하지 못하므로 기본 경로에서 제외한다.

## 8. 구현 순서

1. 기존 Capture v2 DB 또는 source packet에서 표 후보와 주변 텍스트를 읽는다.
2. workbook당 `table-first-request-v1` JSON을 생성한다.
3. 응답 schema `table-first-analysis-v1`과 참조 무결성 검증기를 만든다.
4. AI 없이도 테스트 가능한 결정적 projector를 만든다.
5. 기존 canonical Study/Observation/Evidence schema에 `NEEDS_REVIEW` 기본 상태로 연결한다.
6. 대표 30건에서 호출 수, 시간, prompt 크기, 사람이 보는 표 제목·비교군·지표를 확인한다.
7. 기준을 통과하면 989건을 raw inventory 우선, semantic analysis 후속 방식으로 확대한다.

## 9. 1차 완료 기준

- 한 workbook이 정상적으로 하나의 compact request로 표현된다.
- AI 응답에는 입력에 없던 표/열/텍스트 id를 쓸 수 없다.
- 숫자, 퍼센트, Min/Max/Average는 AI 출력에 의존하지 않는다.
- 유사 표는 묶이고 다른 목적의 표는 분리된다.
- 애매한 표 한 개 때문에 workbook 전체가 차단되지 않는다.
- 대표 30건에서 평균 AI 호출 수 1회, 복잡한 파일도 최대 3회 이내이다.
- 기존 방식 대비 semantic 단계 wall time을 최소 3배 줄인다.

## 10. 구현된 1차 경로

2026-07-20 현재 다음 CLI가 연결되어 있다.

```powershell
# AI를 호출하지 않고 기존 lossless packet을 compact table/text request로 변환
python inference_data_ai_cli.py table-first-request `
  --packet <semantic-source-packet.json> `
  --out outputs/table-first/requests/<name>.json

# 정상 workbook은 AI를 정확히 한 번 호출하고 검토 대기 projection 생성
python inference_data_ai_cli.py table-first-analyze `
  --request outputs/table-first/requests/<name>.json `
  --out outputs/table-first/analyses/<name>.json `
  --projection-out outputs/table-first/projections/<name>.json

# packet 디렉터리를 3개씩 병렬 처리하고 완료 산출물을 재사용
python inference_data_ai_cli.py table-first-batch `
  --packet-dir outputs/semantic-source-packets/<packet-set> `
  --out-dir outputs/table-first/<batch-name> `
  --workers 3
```

구현 파일은 `inference_data_ai_table_first.py`이다. 현재 projection은 의도적으로
`NEEDS_REVIEW`이며 canonical query에 바로 노출되지 않는다. 다음 단계는 대표
workbook의 실제 1회 응답 품질을 확인한 뒤 canonical Study/Observation/Evidence
projector에 연결하는 것이다. 기존 corpus runner는 아직 이 경로로 전환하지 않았고,
사용자 지시 없이 989건 처리를 재시작하지 않는다.

대표 30건의 기존 lossless packet에 대한 AI 미호출 전처리 측정:

- 원본 primary cell: 234,287개
- 생성된 table candidate: 245개
- 생성된 text block: 98개
- compact request: 평균 86.1KB, 중앙값은 별도 운영 측정 대상, 최대 413.8KB
- 정상 table이 있는 workbook의 계획 AI 호출: workbook당 1회
- 7.23MB / 15,424셀 대표 packet: 110,400 bytes, 계획 AI 호출 1회

이 수치는 입력 압축과 호출 수 예산 검증값이며 AI 응답 품질 또는 최종 wall time
검증값은 아니다.

최초 실제 단일 호출 확인:

- source: `017.MSU-20S15-07 Result test sample waitting 2 day and check function_clean.xlsx`
- source range: `Function!B3:Q9`
- request: 12,657 bytes
- AI 호출: 1회
- request 생성 후 analysis 파일 기록까지: 약 23.9초
- 결과: 표 제목, `Lot test waiting 2 day` / `Normal` 군, 14개 지표 식별
- 상태: `NEEDS_REVIEW`, canonical query 미노출

한 건의 실측은 새 실행 경로가 동작하고 기존 수백 초짜리 다중 draft 호출을 피한다는
증거이다. 30건 전체의 의미 정확도와 시간 기준을 통과했다는 뜻은 아니다.

## 11. 정확도·속도 보정 결과

2026-07-20 사용자 검수에서 발견된 문제를 `table-first-builder-v3`와
`table-first-analysis-prompt-v3`에 반영했다.

- Excel 퍼센트 format과 rate 수식을 코드에서 사람 표시값으로 정규화
- raw matrix에서 Min/Max/Average를 행 또는 block 단위로 계산하고 원본 집계값 검산
- 캐시가 없는 제한적 A1 수식을 별도 provenance overlay에서 계산
- 번호형 결과 섹션과 빈 행 뒤 새 header를 별도 논리 표로 분리
- 위치, Nozzle 번호, sample 번호, replicate 번호를 비교군으로 만드는 AI 결과 제거
- Min/Max/Average를 독립 metric으로 만드는 AI 결과 제거
- 48개를 넘는 wide numeric column은 `numericSeries`로 묶어 prompt 크기 제한
- 지원 수식은 개별 provenance와 함께 파생하고, 미지원 수식만 원본 상태로 보존
- 반복되는 날짜별 표 template은 대표 의미 표만 AI에 전달하고 결과를 원본 표 순서로 재확장

대표 30건의 AI 미호출 전처리 재측정:

- primary cell: 234,287개
- 전체 request 생성: 16.4초
- request: 평균 72.8KB, 중앙값 41.5KB, 최대 322.9KB
- 이전 보정 중간값 대비 평균 request 47%, 최대 request 54% 감소
- formula 1차 측정: 3건에서 110개 값 파생, 23건 불필요,
  3건 budget skip, 1건 unsupported skip
- 계획 AI 호출: table이 있는 workbook당 1회

정확도 표본 재검수:

- Function: Test/Normal과 결과 지표 및 근거 범위 통과
- Height: 위치 1~6 가짜 비교 제거, `0.22 -> 22.00%`, 50개 raw의
  Min `1.921`, Max `2.02`, Average `1.9801` 원본과 일치
- Tension: 5s->3s/Normal 비교 통과, 10개/8개 raw의 Min/Max/Average 모두 일치
- Bonding: 6개 논리 표 분리, test/reference 비교 통과, 76개 formula 파생,
  Total NG/NG Rate 복원, 4개 raw 통계 행 일치

표본 네 건은 개선 후 엄격 검수 기준을 통과했다. 단일 AI 호출 시간은 단순 표
약 19~28초, Bonding 복합 표 약 75초였다. 이는 확대 적용의 최소 근거이며 30건
전체 의미 정확도를 보장하지는 않는다.

## 12. 대표 30건 전체 실행 및 자동 감사

2026-07-20 `outputs/semantic-source-packets/pilot-30-v2`의 대표 30건을
`outputs/table-first/pilot-30-v3`으로 실제 처리했다.

최종 산출물:

- workbook 30건 성공, 실패 0건
- table이 없는 3건은 AI 없이 `NO_TABLES`
- table이 있는 27건은 workbook당 성공 AI 분석 1회
- 원본 primary cell 234,287개에서 의미 표 279개, text block 212개 생성
- compact request 평균 74.4KB, 최대 331.1KB
- 표 의미 분류: `HIGH` 190개, `MEDIUM` 82개, `LOW` 7개
- 표 유형: `COMPARISON` 99개, `DESCRIPTIVE` 114개,
  `SUPPORTING` 17개, `TEXT` 49개

속도 병목 보정:

- 최초 실행은 날짜별 반복 표가 88개로 그대로 출력된 workbook 1건이
  600초 제한을 초과했다.
- 동일 구조·행 label을 가진 3개 이상의 반복 표를 대표 template으로 압축하여
  해당 AI 출력 대상을 88개에서 39개로 줄였다.
- 해당 workbook은 재실행 전체 173.3초 안에 완료됐다. 실패 시점 600초 대비
  최소 3.4배 빨라졌으며 원본 88개 table id와 sheet/range로 다시 확장됐다.
- 이미 끝난 29건은 재실행하지 않고 재사용했다. 최종 30건 재검사·projection
  갱신은 AI 호출 0회, 19.0초였다.

결정적 정확도 감사:

- analysis의 table/text/axis 참조 무결성: 30건 모두 통과
- projection 재생성 일치: 30건 모두 통과
- raw Min/Max/Average 검산: 65개 중 65개 일치
- 위치·Nozzle·sample·replicate를 비교군으로 오인: 0개
- 중복 group label: 0개
- Min/Max/Average를 독립 metric으로 오인: 0개
- `COMPARISON` 99개는 모두 2개 이상의 선언 group과 명시적 relation 보유

수식은 23건에서 불필요했고, 3건은 110개를 전부 파생했다. 대형 3건은
지원 가능한 4,822개를 부분 파생하고 미지원 5,443개를 원본 수식으로 보존했다.
나머지 1건은 40개 중 안전하게 적용할 수 있는 값이 1개뿐이라 overlay를
적용하지 않고 검수 대상으로 남겼다.

자동 검수 대상으로 남은 workbook은 7건이다. 이 중 7개 table만 `LOW`
confidence이며 나머지 272개는 `HIGH` 또는 `MEDIUM`이다. 사유는 잘못된 숫자
계산이 아니라 불완전 header, `#REF!`, 교차 sheet/미지원 Excel 함수, 단위·조건
경계의 모호성이다. 이 7건은 canonical query에 자동 승인하지 않고 해당 표만
사람이 확인한다. 정확한 sheet/range, limitation, 수식 오류 표본은
`outputs/table-first/pilot-30-v3/batch-report.json`의 `outliers`에 기록한다.

## 13. 2차 사람 검수 보정

2026-07-21에는 앞선 자동 감사에서 남은 7건을 원본 셀, 수식 상태, 요청 축,
분석 축 순서로 CLI에서 직접 대조했다. WPF는 사용하지 않았다.

확인된 문제와 보정:

- 한 행에 붙어 있던 `SPL Average`와 `F0`를 각각 독립 표와 독립 축으로 분리
- IMP sheet 상단의 `Fo` 요약 행을 아래 IMP 주파수 곡선과 분리
- 숫자와 `#REF!`만 남은 2행 조각은 `TEXT`로 강제하고 의미 분석에서 제외
- 실제 상태가 `PARTIALLY_DERIVED`인 분석에서 과거의 "formula skipped" 문구를
  부분 파생 상태에 맞는 문구로 교체
- `SPL`과 `SPL Average`, `Fo`와 `F0`는 동일 원본 지표명 계열로 정규화하여
  중복 metric 생성을 방지

속도 원칙:

- 구조가 바뀐 Reliability와 Air-pressure workbook 2건만 AI를 다시 호출
- 나머지 28건은 기존 분석을 재사용
- 최종 30건 요청·검증·projection 재생성은 AI 호출 없이 약 20초
- 미지원 또는 부분 파생 수식 상태는 진단 정보로 보존하되, 그 상태만으로
  사람 검수 대상으로 올리지 않음. 실제 `LOW` 의미 표나 값 불일치가 있을 때만
  수동 검수 대상으로 올림

보정 후 대표 30건 결과:

- workbook 30/30 성공, 실패 0
- 의미 표 282개, text block 212개
- confidence: `HIGH` 196개, `MEDIUM` 83개, `LOW` 3개
- raw Min/Max/Average 검산 65/65 일치
- Reliability의 SPL Average와 F0 축 분리 확인
- Air-pressure의 Fo와 IMP 축 분리 확인
- `A175:B176`의 `#REF!` 잔여 조각 2건 모두 `TEXT`, metric 0개 확인

추가 사용자 판단에 따라 `NG Rate = NG / Input × 1,000,000`이 반복 일치하면
단위를 `PPM`으로 코드에서 확정한다. Reliability workbook의 명시적 RAW DATA는
원본 보존 대상으로 두되 별도 사람 검수 대상으로 올리지 않는다. 이에 따라
최종 자동 검수 대상은 광폭 IMP RAW DATA의 의미 경계가 남은 workbook 1건이다.
수식 파생 상태는 `batch-report.json`에 계속 기록되지만 단독 검수 사유로
사용하지 않는다.

## 14. 주파수별 원시 측정 DATA의 AI 제외

2026-07-21 사용자 결정에 따라 `table-first-builder-v4`에서 주파수별 SPL,
THD, IMP/IMPEDANCE 원시 응답표를 AI 의미분석 입력에서 제외한다.

결정적 판별 조건:

- `SPL DATA`, `THD DATA`, `IMP DATA`, `SPL RAW DATA`,
  `IMPEDANCE RAW DATA`, `Frequency response [dBSPL]`처럼 원본에 계열 표기가
  있어야 한다.
- 동시에 세로형 주파수 수열 또는 `100.00Hz`, `106.00Hz`, ... 형태의
  가로형 주파수 축이 있어야 한다.
- 두 조건 중 하나만 충족하면 제외하지 않는다. 따라서 Function NG 요약표의
  `SPL`, `THD`, `SPL+THD+F0` 불량 개수와 독립 `SPL Average`, `F0` 요약표는
  기존 table-first 경로에 남는다.

제외 표는 삭제하지 않는다. `codeOwnedExclusions.rawFrequencyResponseTables`에
원본 sheet/range, source table id, 계열, 주파수 축 metadata와 source packet
보존 상태를 기록한다. 실제 좌표와 값은 `semantic-source-packet-v1`에 그대로
남는다. 이 inventory는 AI prompt JSON에서 제거되므로 원시 측정값과 주파수
배열이 AI 분류·metric·group·comparison 생성에 전달되지 않는다.

대표 30건 재검사 결과:

- 기준 출력: `outputs/table-first/pilot-30-v6-frequency-exclusion`
- 30/30 성공, 실패 0
- 기존 analysis 30건 재사용, AI 호출 0회
- 실행 시간 8.193초
- 7개 workbook에서 원시 응답표 17개 제외
- AI 대상 표 265개, text block 212개
- confidence: `HIGH` 186개, `MEDIUM` 79개, `LOW` 0개
- raw Min/Max/Average 검산 65/65 일치
- 검수 권장 workbook 0건
- 이전 검수 대상 `RAW DATA!H202:DW296`의 IMP 응답표가 제외 inventory에 있고
  analysis table에는 없음을 확인

## 15. 검수된 도메인 용어 사전 연결

2026-07-21 `table-first-builder-v5`에서 다음 읽기 전용 사전을 code-owned
adapter로 연결했다.

`D:/000. MyWorks/005. Program/Repository/MicroSpeaker_ProductTech_DB/db/term_dictionary.csv`

적용 경계:

- `DEFINED` 항목 중 동일한 `normalized_name`을 가진 복수 표기만 명시 alias로
  사용한다. 예: `F0`/`FO`, `VP-CD`/`VP+CD`/`VP/CD`.
- `NEEDS_DEFINITION` 항목은 규칙으로 사용하지 않는다.
- `IGNORE`는 용어 alias/정규화 대상에서 제외한다. 원본 표의 group 또는 metric을
  삭제하지 않는다. `SMALL`, `REF`처럼 문맥에 따라 실제 source label일 수 있기
  때문이다.
- `#REF!` spreadsheet 오류 조각은 용어 사전과 별도로 기존 전용 규칙에서
  계속 의미 분석 제외한다.
- 사전 snapshot은 `codeOwnedTermDictionary`에 fingerprint, alias 그룹, IGNORE
  목록으로 저장하고 AI prompt JSON에서는 제거한다.
- `SPL`은 사전에 정의돼 있지만 `SPL Average`는 별도 정의가 없으므로 두 이름을
  같은 metric key로 합치지 않는다. 명시적인 `SPL Average`/`F0` source hint가
  있으면 원본 이름과 축만 유지한다.

CLI에서는 필요할 때 `table-first-request`와 `table-first-batch`의
`--term-dictionary`로 읽기 전용 사전 경로를 명시할 수 있다. 기본값은 위의
형제 프로젝트 사전이다.

대표 30건 재검사 결과:

- 기준 출력: `outputs/table-first/pilot-30-v7-term-dictionary`
- 30/30 성공, 실패 0
- 기존 analysis 30건 재사용, AI 호출 0회
- 실행 시간 7.254초
- 사전 상태: `LOADED` 30/30
- 사전 fingerprint:
  `537e330e3b66b12dced13546ef75dc6ad783afa6483de7451426bc2c255706bf`
- DEFINED 179개, IGNORE 19개, 명시 alias 그룹 19개
- v6 대비 30건의 table semantic 결과 전부 동일
- `SPL Average`/`F0` source metric 표 5개 이름·축 일치
- AI 대상 표 265개, raw 주파수 응답 제외 17개, 검수 권장 0건

## 16. 989건 AI-free request audit

AI 분석을 시작하기 전에 전체 corpus의 table-first 입력을 검사할 수 있도록
`table-first-request-batch` 명령을 추가했다. 이 명령은 request 생성과
code-owned audit만 수행하며 AI 분석 함수를 호출하지 않는다.

```powershell
python inference_data_ai_cli.py table-first-request-batch `
  --packet-dir outputs/semantic-source-packets/full-989-v1 `
  --out-dir outputs/table-first/full-989-v1-request-audit `
  --workers 6 `
  --oversized-request-bytes 240000
```

보고서는 성공/실패, request 재사용, 표·텍스트 수, `NO_TABLES`, 원시 주파수
응답 제외, formula derivation 오류, aggregate mismatch, request 크기 분위수와
임계치 초과 workbook을 기록한다.

2026-07-21 전체 점검 결과:

- semantic packet 989/989 생성 성공, 실패 0, 1,505,559 cells
- request audit 989/989 성공, 실패 0, AI 호출 0회
- 재실행에서 request 989/989 byte-for-byte 재사용
- table 5,546개, text block 6,251개, `NO_TABLES` 9건
- raw SPL/THD/IMP 응답 제외: 77 workbook, 177개 표, 783,086 cells
- request 크기 median 30,183 bytes, p95 101,449, p99 212,043,
  최대 459,854; 240,000 bytes 초과 9건
- formula error 11 workbook, 6,165개
- aggregate check `MATCH` 1,490개, `MISMATCH` 310개; mismatch workbook 57건
- 용어 사전 `LOADED` 989/989
- 최종 보고서:
  `outputs/table-first/full-989-v1-request-audit/request-batch-report.json`

따라서 다음 단계는 980건 AI 호출이 아니다. aggregate mismatch 57건,
formula error 11건, oversized request 9건을 먼저 해소하거나 명시적으로
격리한 뒤 AI 실행 범위를 정한다.

## 17. Aggregate 범위 선택 v6

989건 1차 audit의 aggregate mismatch 310건을 원본 셀까지 역추적한 결과,
기존 로직에 다음 오탐 원인이 있었다.

- 한 표 안의 복수 `Min/Max/Avg` triplet을 하나의 집계 묶음으로 결합
- `Total NG`, `Possition`, `Sample` 번호, 공정시간을 raw 측정값에 포함
- 다음 요약 행까지의 세로 block 경계를 구분하지 않음
- 집계 triplet 사이를 넘어 떨어진 숫자 영역을 raw 범위로 선택

`table-first-builder-v6`에서는 다음과 같이 수정했다.

- source header row별로 aggregate triplet을 독립 분리
- 같은 header 계열의 반복 측정 열과 직접 인접한 raw 영역만 선택
- 명시 basis 열 한 개가 사이에 있을 때만 제한적으로 건너뜀
- 다음 aggregate summary row 직전까지만 세로 block을 계산
- 단일 raw 열의 세로 block 집계 지원
- aggregate header가 실제 source cell에 없는 추정 fallback 제거

회귀 테스트는 30/30 통과했다. AI 없이 전체 989건을 다시 검사한 기준 출력은
다음과 같다.

`outputs/table-first/full-989-v3-aggregate-audit/request-batch-report.json`

- 989/989 성공, 실패 0, AI 호출 0회
- aggregate `MATCH` 1,743개, `MISMATCH` 136개
- mismatch workbook 57개에서 29개로 감소
- 기존 대비 mismatch 310개에서 136개로 56.1% 감소
- 136개 중 61개는 raw Min/Max가 모두 일치하고 source Average만 다름
- 나머지 75개는 Min 또는 Max도 달라 자동 승인하지 않고 사람 검수 유지
- table 5,546개, text block 6,251개, `NO_TABLES` 9건은 기존과 동일

source Average가 `Total NG=0` 열을 분모에 포함하지만 source Min/Max는 해당 열을
제외하는 사례처럼, 61개 average-only mismatch는 단순 반올림이 아니라 원본
계산 정의 차이다. 코드는 원시 측정 열만 사용한 평균을 유지하고 source 값을
덮어쓰지 않는다.

## 18. Formula derivation v2 전수 감사

1차 audit의 formula error 6,165개를 수식 구문과 의존 셀까지 분류했다. 가장 큰
원인은 직접 cross-sheet 참조 2,970개, text 수식, 원본 `#DIV/0!`였다.

`restricted-a1-arithmetic-v2`는 다음만 추가로 허용한다.

- 현재 semantic packet 안에 실제 존재하는 sheet의 직접 A1 셀 참조
- 참조 대상이 숫자이면 기존 provenance 의존성 기록과 함께 숫자 파생
- 직접 text 참조와 제한된 text 연결/판정 수식은 `NON_NUMERIC`으로 분리
- 없는 sheet, range 반환, 깨진 참조, 지원하지 않는 고급 함수는 fail-closed

Capture v2 원본은 수정하지 않으며 파생값은 계속 checksum이 있는 별도 overlay에만
기록한다. 회귀 테스트 42/42 통과 후 AI 없이 989건을 두 번 검사했다.

`outputs/table-first/full-989-v4-formula-audit/request-batch-report.json`

- 989/989 성공, 실패 0, AI 호출 0회
- 두 번째 실행 request 989/989 byte-for-byte 재사용
- builder `table-first-builder-v7`
- formula 13,995개: 숫자 10,900개, 비숫자 1,130개, error 1,965개
- error 6,165개에서 1,965개로 68.1% 감소
- error workbook 11개에서 9개로 감소
- formula 상태: `NOT_NEEDED` 963, `DERIVED` 16,
  `CLASSIFIED_NON_NUMERIC` 1, `PARTIALLY_DERIVED` 8,
  `SKIPPED_UNSUPPORTED` 1

잔여 error는 원본 `#DIV/0!` 1,209개, 한 workbook의 고급 집계/조회 수식
618개, 잘못되거나 오래된 sheet/ref/range 참조 136개, 수식으로 오인된 text
2개다. 이 값들은 자동 보정하지 않고 진단 정보로 유지한다.

## 19. Prompt v4 의미 투영 압축

`table-first-request-batch`가 저장 request 크기뿐 아니라 반복 template 압축과
prompt 의미 투영 이후의 실제 `promptBytes`도 기록하도록 변경했다. 저장 request는
검산과 provenance를 위해 상세 수치를 유지하므로 실제 AI context 크기와 같지 않다.

`table-first-analysis-prompt-v4`는 AI가 계산하면 안 되는 다음 정보를 prompt에서만
제거하거나 요약한다.

- aggregate check의 raw/explicit/calculated 숫자 상세
- 측정 열의 Min/Max/Average, numeric count, source range와 raw 표본
- `rowLabels`에 이미 보존된 preview text 중복
- preview의 measure/aggregate 숫자

table/row/column ID, title 후보, header text, row label, identifier/basis 값, 날짜,
percent 표시 scale, 비중복 preview 문맥은 유지한다. 따라서 AI semantic 출력의
허용 ID와 문맥은 그대로이고, 제거된 숫자는 저장 request와 Capture v2에서 계속
code-owned evidence로 사용할 수 있다.

회귀 테스트 43/43 통과 후 다음 보고서를 두 번 생성했다.

`outputs/table-first/full-989-v5-prompt-audit/request-batch-report.json`

- 989/989 성공, 실패 0, AI 호출 0회
- 재실행 request 989/989 byte-for-byte 재사용
- prompt 총량 39,465,194 bytes에서 21,179,656 bytes로 46.3% 감소
- 최대 prompt 436,098 bytes에서 215,729 bytes로 감소
- 240,000 bytes 초과 prompt 9건에서 0건으로 감소
- prompt median 17,103, p95 47,200, p99 101,728 bytes
- 저장 request 최대 459,874 bytes/초과 10건은 원본 상세 보존 상태로 유지
- workbook당 한 번의 AI 호출 정책을 유지하고 분할 호출은 사용하지 않음

## 20. 사용자용 정적 HTML 리포트

완료된 `table-first-batch-report-v1` 디렉터리에서 통합 index와 workbook 상세
페이지를 생성하는 `table-first-html` 명령을 제공한다. 렌더러는 request,
analysis, projection의 `requestId`가 모두 같은 경우에만 페이지를 만든다.

```powershell
python inference_data_ai_cli.py table-first-html `
  --batch-dir outputs/table-first/pilot-30-v7-term-dictionary
```

기본 출력은 `<batch-dir>/html-report`이다.

- `index.html`: 검색·상태/검수 필터, corpus 요약, workbook 목록
- `workbooks/*.html`: Study별 시험군×지표 통합표, 지표 통계, 검수 사유
- `html-report.json`: source batch와 페이지 수 manifest

sheet/range 같은 기술 근거는 request/projection JSON에는 유지하되 사용자용 HTML의
표와 본문에는 기본 표시하지 않는다. 필요해질 때 별도 상세 보기로 제공할 수 있다.

각 Study는 시험군, 비교 관계, 지표, 코드 통계를 별도 영역으로 나누지 않는다.
첫 열에 시험군을 행으로 배치하고 지표를 열로 배치한다. 행형 원본은 해당 시험군의
실제 값을 표시하고, 열형 원본은 안전하게 연결되는 축의 count/Min/Max/Average를
같은 지표 셀에 표시한다. 비교 대상은 시험군 셀에 함께 표시하며, 안전한 연결을
결정할 수 없는 조합은 다른 시험군의 통계를 복제하지 않고 `-`로 둔다.

모든 페이지는 CSS와 JavaScript를 자체 포함하고 외부 네트워크나 서버를 요구하지
않는다. 원본 숫자를 다시 계산하지 않고 projection의 AI 의미 결과와 Capture 기반
deterministic facts를 구분해 표시한다.

기존 대표 30건으로 index 1개와 상세 30개를 생성했다. 링크 누락 0건,
재실행 byte-for-byte 재사용 32/32, HTML 관련 테스트를 포함한 좁은 회귀 테스트
45/45를 통과했다. 기준 index는 다음과 같다.

`outputs/table-first/pilot-30-v7-term-dictionary/html-report/index.html`

## 21. 전체 이력 검색 DB와 근거 답변

최종 이력 질의 경로에서는 `table-first-builder-v7`과
`table-first-analysis-prompt-v4`를 동결한다. 검색, 날짜 정규화, 답변 문구,
WPF 표시를 바꿀 때 request/analysis/projection AI 작업을 다시 실행하지 않는다.

완료된 batch를 다음 명령으로 별도 SQLite 이력 인덱스로 만든다.

```powershell
python inference_data_ai_cli.py table-first-history-index `
  --batch-dir outputs/table-first/full-989-v8-prompt-v4 `
  --db outputs/table-first-history/history.sqlite
```

대표 golden question 전체가 WPF와 동일한 30개 workbook 검색 범위 안에서 지정된
primary workbook을 검색하고 검토 gate를 유지하는지는 AI 호출 없이 다음 명령으로
검사한다. 답변은 상위 12개만 상세 표시하고 나머지는 한 줄 원본 목록으로 축약한다.

```powershell
python inference_data_ai_cli.py table-first-history-acceptance `
  --db outputs/table-first-history/history.sqlite `
  --manifest pilot/representative-pilot-v1.json `
  --out outputs/table-first-history/history-acceptance-report.json
```

`table-first-history-query`는 다음 정보를 검색한다.

- workbook 요약, Study 제목과 그룹
- 시험군/기준군 label과 role
- 지표명과 단위
- 비교 관계와 제한 사항
- 파일명과 source-authored 날짜 후보
- 용어 사전 alias와 한국어 설명

질의 결과는 workbook과 Study를 시간순으로 정리하고, 각 항목에 안정적인
`TF-EVD-*` 원본 파일·시트·범위 근거를 연결한다. 모든 table-first 결과는
사람 검토 전 의미 투영이므로 검색과 이력 설명에는 보이되 승인된 effect나
인과 결론을 만들지 않는다. 시험군별 수치 연결이 불확실하면 숫자를 추정하거나
다른 열의 통계를 복제하지 않는다.

WPF의 기존 질문 화면은 이력 DB가 있으면 전체 이력 경로를 우선 사용하고,
`TF-EVD-*` 선택 시 저장된 request preview와 정확한 원본 workbook/sheet/range를
표시한다. 전체 corpus 정적 HTML 렌더링은 사용자 요청에 따라 수행하지 않는다.

## 22. 질의 시점 문맥 AI 답변

`table-first-history-query`는 후보 검색과 회귀 점검용 결정론적 명령으로 유지한다.
사용자용 WPF 답변은 `table-first-contextual-query`를 호출한다.

```powershell
python inference_data_ai_cli.py table-first-contextual-query `
  --db outputs/table-first-history/history.sqlite `
  --question "VP CD 조립에 따른 Hearing 불량률 추이" `
  --candidate-limit 30 `
  --detail-candidate-limit 18
```

처리 경계는 다음과 같다.

- 단어·alias 점수는 최대 30개 후보를 수집하는 데만 사용한다.
- 질의 시점 AI가 대상, 조건, 지표, 비교, 시간축과 답변 모드를 분리한다.
- 같은 Study에서 대상·조건·지표의 직접 관계가 확인되지 않는 후보는 제외한다.
- 수치는 `TF-FCT-*`로 등록된 source display 값만 사용할 수 있고, 반드시
  `TF-EVD-*` 원본 파일·시트·범위와 연결한다.
- 추이는 동일 의미·단위의 지표가 비교 가능한 날짜 또는 순서 관측에 2개 이상
  연결될 때만 완전 답변으로 판정한다.
- 근거가 부족하면 관련 파일 목록을 쏟아내지 않고 `PARTIAL` 또는
  `INSUFFICIENT`로 무엇이 부족한지 설명한다.
- WPF는 직접 답변, AI가 이해한 질문, 핵심 판단, 비교 가능한 관측값, 한계,
  최대 10개의 핵심 근거 순서로 표시한다.

2026-07-22 실제 989건 DB에서 대표 질문을 실행한 결과 후보 30건 중 직접 관련
Study 1건만 채택하고 29건을 제외했다. 날짜와 Hearing 전용 불량률이 연결되지 않아
`CONTEXTUAL_AI_PARTIAL`로 추이 판정 불가를 반환했으며, 관련 조건별 관측 건수 6개는
모두 등록된 fact/evidence 연결 검증을 통과했다.
