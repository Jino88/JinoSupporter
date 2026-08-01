# Codex Instructions

## Build Rule
- When verifying code changes, run the narrowest relevant build, test, lint, or type-check command for the files or module you changed.
- Do not run broad repository-wide builds or full test suites unless the user explicitly asks for them, or unless there is no narrower reliable verification command.
- Before running any build or test command, identify the changed files and choose a command scoped to those files, their project, or their package.
- If a narrow verification command is unavailable, explain that briefly and run the smallest practical project-level command.
- Do not treat failures from unrelated files or modules as part of your change unless they block verifying the edited code.
- Do not modify, revert, format, or clean up unrelated files while trying to fix build errors.

## Mandatory Verification Rule (필수 검증 수칙)
- **무조건 검증**: 모든 코드 수정, 스크립트 작성, 기능 변경 작업 완료 후에는 반드시 해당 작업에 대한 엄격한 실측 검증(Build / Run / Test / Script verification)을 진행합니다.
- **자동 개선/수정**: 검증 과정에서 문제나 에러가 발견되면 단신 보고로 끝내지 않고 즉시 원인을 파악하여 개선 및 수정 조치까지 완료합니다.
- **결과 증빙**: 단순 "작업 완료" 선언을 금지하고, 검증 실행 결과(로그, 텍스트/JSON output, 빌드 성공 여부 등)를 최종 보고서에 명시합니다.

