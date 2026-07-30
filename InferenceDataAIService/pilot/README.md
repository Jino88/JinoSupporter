# Representative Pilot v1

`representative-pilot-v1.json`은 전체 900여 개 이력 Excel을 마이그레이션하기 전에 파이프라인의 위험 범위를 빠르게 검증하는 **커버리지 fixture**다. 2026-07-17 기준으로 기록된 30개 상대 경로는 모두 `D:\000. MyWorks\test\result\InputDataFinish` 아래에 실제 파일로 존재함을 확인했다.

이 목록은 정답 라벨이나 승인된 분석 결과가 아니다. 파일명, 기존 그룹, renderer, 후보 분류에서 얻은 설명은 표본을 다양하게 고르기 위한 힌트일 뿐이다. 각 Study, Factor, Arm, Outcome, Comparison, Effect, Evidence 값과 비교 가능성은 원본 표 셀을 근거로 별도의 추출·검증 절차를 거쳐야 한다.

## 이미지 범위

이미지는 최종 데이터 추출 및 분석 범위가 아니다.

- 이미지 파일, 앵커, OCR, 이미지 속 표나 문구를 추출하지 않는다.
- 이미지 중심·이미지 전용 워크북에서 유효한 표 셀을 찾지 못하면 `NO_TABULAR_EVIDENCE` 또는 `DESCRIPTIVE_ONLY`로 종료한다.
- 이미지에서 factor, outcome, 숫자, 결론, defect taxonomy를 추론하지 않는다.
- 이미지 중심 표본은 정상적인 제외 상태와 무근거 수치 생성 방지를 검증하기 위해서만 유지한다.

## Gate 사용법

각 단계의 검증은 다음 순서를 따른다.

1. **경로 gate**: manifest의 30개 상대 경로가 기준 root 아래에 정확히 한 파일씩 존재해야 한다.
2. **적재 gate**: 30개 모두 성공, 명시적 제외, `NO_TABULAR_EVIDENCE`, `EMPTY_SOURCE` 중 하나의 terminal 상태를 가져야 한다. 조용한 누락은 실패다.
3. **구조 gate**: 표 기반 자료에서 워크북·시트·셀·병합·수식·표시값·number format과 source range lineage가 재현되어야 한다. 이미지는 이 gate에 포함하지 않는다.
4. **canonical gate**: 명확한 2-arm, paired, multi-arm/DOE, 복합 요인, 반복 시점, reference, 빈 파일 패턴을 서로 혼동하지 않아야 한다.
5. **비교 gate**: 검증된 대조군/비교군과 동일 outcome·단위·맥락이 있는 경우에만 Effect를 계산한다. 복합 변경은 confounded로 표시하고 단일 요인 효과로 계산하지 않는다.
6. **근거 gate**: 모든 정량 주장에는 불변 데이터 ID와 원본 sheet/range가 있어야 하며, 그 범위의 셀로 결과를 재계산할 수 있어야 한다.
7. **질의 gate**: 10개 golden question이 관련 Study를 찾고, 계산 가능한 자료와 제외 자료를 구분하며, 누락·중복·근거 없는 수치를 만들지 않아야 한다.

파일럿 통과는 30개 워크북에 대한 사람의 최종 의미 승인을 뜻하지 않는다. 전수 마이그레이션 전에 구조적·계산적 실패 유형이 제어되었다는 의미다. 숫자 기대값, arm 대응, outcome 의미 같은 ground-truth 회귀가 필요할 때는 이 manifest를 바꾸지 말고 별도의 검증 라벨 파일을 버전 관리해 추가한다.

전체 이력 처리는 파일럿 gate가 통과한 뒤 같은 코드와 스키마 버전으로 수행한다. 파일럿용 예외 규칙이나 파일명별 하드코딩은 허용하지 않는다.
