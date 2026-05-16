# Codex Instructions

## Build Rule
- When verifying code changes, run the narrowest relevant build, test, lint, or type-check command for the files or module you changed.
- Do not run broad repository-wide builds or full test suites unless the user explicitly asks for them, or unless there is no narrower reliable verification command.
- Before running any build or test command, identify the changed files and choose a command scoped to those files, their project, or their package.
- If a narrow verification command is unavailable, explain that briefly and run the smallest practical project-level command.
- Do not treat failures from unrelated files or modules as part of your change unless they block verifying the edited code.
- Do not modify, revert, format, or clean up unrelated files while trying to fix build errors.

