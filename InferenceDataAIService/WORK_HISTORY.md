# 작업 이력

## 2026-07-15 — 반복 불량률 블록 구조 그룹 1차 지원

### 구현 범위

- `REPEATED_DEFECT_BLOCK_LAYOUT_CANDIDATE`를 구조 스캔의 관찰 후보로 추가했다. 2행 헤더 창 안에 Input이 둘 이상이고 `Total NG`, `NG Rate`가 함께 있을 때만 후보로 남긴다.
- 숫자 적재는 같은 엄격한 서명에 `REPEATED_DEFECT_BLOCK_NUMERIC_TABLE`을 부여한다. 원본 Excel·수식·병합 처리 방식은 바꾸지 않았다.
- 숫자 검토는 반복된 Input 구간을 개별 블록으로 분리한다. 각 블록 안에서 명시적인 `Input → Total NG → NG Rate` 열 쌍이 하나일 때만 `repeated_defect_block_facts`에 적재한다.
- 블록에 Test·Normal과 같은 날짜가 모두 명시된 경우에만, 같은 workbook·같은 표·같은 블록·같은 날짜 범위에서 `repeated_block_test_normal_comparisons`를 만든다.
- 역할·날짜·열 쌍이 불명확하거나 누락된 경우는 `NEEDS_REVIEW`로 남긴다. 일반 `NG`를 Total NG로 넓게 해석하지 않고, 자동 품질·출하·인과 결론도 만들지 않는다.
- HTML에는 `다중 블록 불량률`과 `다중 블록 Test–Normal 비교` 전용 표를 추가했다. 좌표, 수식, DB 경로, 원본 샘플은 표시하지 않는다.
- 배치별 `numeric-structure-groups.json`에는 결정적 헤더 서명 기반의 구조 그룹·사실 수 요약을 기록한다.

### 검증

- `dotnet build InferenceDataAIService.Wpf/InferenceDataAIService.Wpf.csproj --no-restore` 성공 (경고 0, 오류 0).
- WPF 데스크톱 앱과 실제 Excel 배치는 이 검증에서 실행하지 않았다.

### 다음 작업

- 대표 반복 블록 Excel 표본으로 facts·동일 블록 Test–Normal 비교·HTML 결과를 실행 검증한다.
- 다단 헤더/세로형 비교/NG breakdown은 별도 구조 그룹으로만 후속 확장한다.

## 2026-07-14 — WPF 숫자 검토 엔진 C# 전환

### 목표

- Excel 원본에서 숫자 표 사실만 배치 DB에 적재한다.
- 불량률은 명시적인 Test와 Normal만 대상으로, 같은 표와 같은 날짜일 때만 비교한다.
- HTML은 숫자 표와 비교 결과만 표시하고, 셀 좌표·수식·DB 경로·원본 근거는 표시하지 않는다.
- 숫자 검토 흐름에서 Excel COM과 Python 실행 의존성을 제거한다.

### 완료한 엔진

| 단계 | C# 엔진 | 역할 |
| --- | --- | --- |
| 구조 스캔 | StructureScanEngine.cs | 읽기 전용 OpenXML 구조·표 후보 스캔 |
| 숫자 적재 | NumericCaptureEngine.cs | 숫자·날짜·수식 캐시·병합·짧은 표 라벨을 numeric-capture.sqlite에 적재 |
| 숫자 검토 | NumericReviewEngine.cs | 불량률/측정 통계 사실과 동일 표·동일 날짜 Test–Normal 비교 생성 |
| HTML 렌더링 | NumericRendererEngine.cs | 배치 DB만 읽어 workbook별 HTML, 색인, renderer manifest 생성 |

MainWindow.xaml.cs의 숫자 검토 실행 경로는 위 C# 엔진 3개(적재·검토·렌더링)를 순서대로 호출한다. inference_data_ai_numeric_capture.py, inference_data_ai_numeric_review.py, inference_data_ai_numeric_renderer.py는 WPF 숫자 검토 실행 경로에서 호출하지 않는다.

### 데이터·비교 규칙

- Excel/COM을 다시 열거나 수식을 재계산하지 않는다.
- Test/Normal 역할은 라벨의 명시적 영문 단어로만 인식한다. Control은 비교 역할이 아니다.
- 비교 키는 반드시 table_id + yyyy-MM-dd이다. 다른 표, 다른 날짜, 다른 workbook 간 비교는 만들지 않는다.
- Input이 0 이하이거나 Total NG가 음수인 사실은 NEEDS_REVIEW로 저장하며, 품질/출하/개선/원인 같은 서술형 판정은 생성하지 않는다.
- HTML에는 source coordinate, formula, raw sample, fingerprint, DB 경로를 노출하지 않는다.
- HTML의 UI 문구는 한국어로 유지하고 Test, Normal, Input, Total NG 등 고유 영문은 원문대로 둔다.

### 검증 결과

- WPF Debug 빌드 성공: 경고 0, 오류 0.
- 실제 Excel 20개 표본:
  - C# 숫자 적재 DB와 Python 기준 DB의 workbook/sheet/병합/숫자/수식/날짜/표 후보/라벨 논리 행 일치.
  - C# 검토 DB와 Python 기준 DB의 table review, defect fact, measurement fact, Test–Normal comparison 논리 행 일치.
  - C# renderer와 Python renderer의 workbook별 보고서 989개 및 index HTML 원시 바이트 일치.
  - renderer summary/manifest는 생성 시각을 제외하고 일치.
- fixture 검증: VALID 및 NO_SAME_DAY_NORMAL 상태를 모두 확인.
- 숫자 엔진 검증 전후 outputs/universal-grid/InputDataFinish.sqlite SHA-256 불변:
  AC024C8E0D854B200EE1F09588E3539AD0678519BAABACC191630FB0A60054C2
- WPF 데스크톱 EXE는 이 검증 단계에서 실행하지 않았다.

### 현재 범위와 다음 작업

- 완료 범위는 WPF의 숫자 검토 배치다.
- 기존 개별 workbook AI 분석/Excel COM 경로는 별도 기능이며, 아직 Python runner를 사용한다.
- 다음 단계는 다양한 Excel 양식의 구조 그룹을 확정하고, 그룹별 숫자 사실 모델과 전용 HTML 표시 규칙을 확장하는 작업이다.
