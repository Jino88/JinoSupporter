# InferenceDataAIService 세션 인수인계 — 2026-07-15

## 현재 완료 상태

WPF의 숫자 검토 배치는 Python·Excel COM 없이 아래 C# 흐름으로 동작한다.

`StructureScanEngine → NumericCaptureEngine → NumericReviewEngine → NumericRendererEngine`

- 구조 스캔은 OpenXML 패키지를 읽어 구조·헤더 후보만 수집한다.
- 숫자 적재는 숫자·날짜·수식 캐시·병합·인접 표 라벨만 배치 전용 SQLite에 보존한다.
- 숫자 검토는 명시된 Test/Normal과 동일 표·동일 날짜 조건만 비교한다.
- HTML은 배치 DB의 숫자 사실만 읽는다. 원본 좌표, 수식, DB 경로, raw sample, fingerprint는 표시하지 않는다.
- 이 WPF 숫자 검토 경로는 `inference_data_ai_numeric_*.py`를 실행하지 않는다. 기존 개별 workbook AI 분석/Excel COM 경로는 별도 기능으로 남아 있다.

## 2026-07-15 추가 구현: `REPEATED_DEFECT_BLOCKS`

여러 개의 `Input` 열이 수평으로 반복되는 불량률 표를 첫 구조 그룹으로 추가했다.

- 구조 스캔 후보: 2행 헤더 창에 `Input`이 둘 이상이고 `Total NG`, `NG Rate`가 있는 경우 `REPEATED_DEFECT_BLOCK_LAYOUT_CANDIDATE`.
- 숫자 표 후보: 같은 서명에 `REPEATED_DEFECT_BLOCK_NUMERIC_TABLE`.
- 전용 사실 테이블: `repeated_defect_block_facts`.
- 전용 비교 테이블: `repeated_block_test_normal_comparisons`.
- 검토 DB 메타데이터 스키마: `numeric-review-db-v2`.
- 전용 HTML: `다중 블록 불량률`, `다중 블록 Test–Normal 비교`.
- 배치 요약: `numeric-structure-groups.json`.

### 안전 계약

- 각 `Input` 열의 다음 `Input` 전까지를 하나의 블록으로 본다.
- 그 구간 안에 명시적인 `Input → Total NG → NG Rate` 열 쌍이 각각 하나일 때만 사실을 저장한다.
- 일반 `NG` 또는 `NG rate`만으로는 `Total NG`를 추정하지 않는다.
- 불완전하거나 다수의 열 쌍이 있는 블록은 해당 표를 `NEEDS_REVIEW`로 남긴다.
- 비교는 반드시 같은 workbook 표·같은 블록·같은 날짜·명시적 Test/Normal에서만 만든다.
- 역할·날짜·숫자 값이 누락되거나 불가능하면 해당 사실은 `NEEDS_REVIEW`이며 자동 비교 근거가 되지 않는다.
- 품질, 출하, 개선, 원인에 대한 서술형 결론은 만들지 않는다.

## 변경 파일

- `InferenceDataAIService.Wpf/StructureScanEngine.cs`
- `InferenceDataAIService.Wpf/NumericCaptureEngine.cs`
- `InferenceDataAIService.Wpf/NumericReviewEngine.cs`
- `InferenceDataAIService.Wpf/NumericRendererEngine.cs`
- `WORK_HISTORY.md` — 7/14 C# 전환 이력 및 7/15 반복 블록 지원 이력

## 검증 완료 및 공백

완료:

```powershell
dotnet build .\InferenceDataAIService.Wpf\InferenceDataAIService.Wpf.csproj --no-restore
```

- 결과: 경고 0, 오류 0.

아직 하지 않음:

- WPF 데스크톱 앱 실행.
- 실제 Excel 배치 실행.
- C# 반복 블록 pairing/comparison 전용 자동 테스트. 현재 C# test project가 없다.

## 다음 세션에서 바로 할 일

1. 대표 반복 블록 Excel 세 종류로 C# 숫자 검토 배치를 실행 검증한다.
   - 유효한 동일 블록·동일 날짜 Test/Normal 비교 1건 이상.
   - Test/Normal이 없어 비교가 0건인 표.
   - 블록 열 쌍이 누락 또는 모호해 `NEEDS_REVIEW`가 되는 표.
2. 각 표에서 DB facts, comparison status, 전용 HTML 표가 안전 계약과 일치하는지 확인한다.
3. 가능하면 위 세 경우를 재현하는 C# fixture test project를 추가한다. 실행 검증이 통과한 뒤에만 다음 구조 그룹(다단 헤더, 세로형 비교, NG breakdown)을 확장한다.
4. 실제 배치 전후 `outputs/universal-grid/InputDataFinish.sqlite` 해시가 바뀌지 않았는지 확인한다. 숫자 검토는 이 DB를 읽거나 변경해서는 안 된다.

## 작업 트리 주의

- 현재 `InferenceDataAIService`에는 이번 작업보다 앞선 WPF·Python·테스트·outputs 변경도 함께 존재한다. 재개 시 이 변경들을 되돌리거나 정리하지 말고, 필요한 범위만 추가 수정한다.
- 원본 Excel은 read-only로만 다루며 저장하지 않는다.
- 숫자 검토 경로에서 Excel COM, Python 실행, 수식 재계산을 다시 추가하지 않는다.
- 앱·데스크톱 UI는 사용자가 명시적으로 요청할 때만 실행한다.

더 자세한 완료 이력과 장기 다음 단계는 `WORK_HISTORY.md` 및 `WORKING_CONTEXT.md`를 참고한다.
