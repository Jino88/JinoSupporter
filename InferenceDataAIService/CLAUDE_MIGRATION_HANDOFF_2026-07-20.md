# InferenceDataAIService Claude 이관 인수인계

기준일: 2026-07-20 (+07)  
대상 프로젝트: `D:\000. MyWorks\005. Program\Repository\JinoSupporter\InferenceDataAIService`

## 1. 결론

이 프로젝트를 Claude로 이어서 개발할 때의 기본 선택은 다음과 같다.

| 용도 | 모델 | effort | 판단 |
|---|---|---:|---|
| 최초 인수, 아키텍처 판단, provider 이관, 어려운 디버깅 | `claude-opus-4-8` | `xhigh` | 주 개발 모델 |
| 계획이 확정된 일반 구현과 반복 수정 | `claude-sonnet-5` | `high` | 비용·속도 균형 |
| 실제 workbook locator와 semantic draft | `claude-sonnet-5` | `high` | 대표 30 gate 통과 전 기본 모델 |
| 단순 locator 대량 처리 | `claude-haiku-4-5-20251001` | 해당 없음 | Sonnet과 정확도 동등성이 입증된 후에만 |
| 장시간·최고난도 자율 작업 | `claude-fable-5` | `high` 이상 | 기본값으로는 과도함 |

즉, **코드 이관과 첫 복구는 Opus 4.8**, 실제 의미분석 호출은 우선
**Sonnet 5**가 적합하다. Fable 5는 가장 강한 모델이지만 이 프로젝트의
기본 처리 모델로 쓰기에는 비용이 크다. Haiku 4.5는 값싼 locator 후보일
뿐이며, 현재의 다국어·병합 셀·수식·원인/결과 구분을 대표 gate에서
검증하기 전에는 사용하지 않는다.

Claude Code에서는 가능하면 `/model opusplan`을 사용한다. 이 모드는
계획 단계에서 Opus, 구현 단계에서 Sonnet을 사용한다. 중요한 provider
설계나 semantic safety 수정은 `claude-opus-4-8`을 명시하고, 단순 구현은
`claude-sonnet-5`로 전환한다.

## 2. 반드시 구분할 두 가지 이관

### 2.1 개발 세션을 Claude Code로 이관

Claude Code가 이 저장소를 읽고 코드를 수정하는 경로다. 이것만으로도
남은 Python/C# 구현을 이어갈 수 있다.

그러나 현재 애플리케이션 내부의 AI 호출은 여전히 `codex exec`를 직접
실행한다. 따라서 개발 도구만 Claude Code로 바꿔도 B24/B25/B14 같은
실제 semantic AI 작업은 실행되지 않는다.

### 2.2 애플리케이션 AI backend를 Claude로 이관

다음 중 하나를 구현해야 한다.

1. **전환용 Claude Code CLI bridge**
   - Claude Pro/Max 구독 사용량으로 먼저 재개하려는 경우에 적합하다.
   - 현재 subprocess 구조와 가장 비슷해 수정 범위가 작다.
   - Claude Code 사용량 제한과 claude.ai 사용량이 공유된다.
2. **정식 Anthropic Messages API provider**
   - 장기 운영과 947개 대량 처리를 위한 권장 경로다.
   - 호출별 토큰, 비용, request ID, retry를 정확히 기록할 수 있다.
   - Message Batches API와 prompt caching을 사용할 수 있다.

두 경로를 한 함수 안에 조건문으로 섞지 말고 공통 provider 계약 뒤에
각 구현을 둔다. Codex 호환 provider도 당분간 유지해 결과를 비교할 수
있게 한다.

## 3. 현재 프로젝트 상태

### 완료된 기반

- 989개 원본 workbook의 Capture v2가 완료됐다.
- 원본 SHA-256, current Capture revision, SQLite `quick_check`는 정상이다.
- 현재 semantic journal은 실제 실행 기준으로 대략 다음 상태다.
  - 39 `COMPLETED`
  - B24/B25는 프로세스가 없지만 journal에 `RUNNING`
  - B14는 `FAILED`
  - 947 `PENDING`
- staged draft v2, exact prompt hash, source-cell ownership, resume 검증,
  deterministic merge와 fail-closed validator가 구현돼 있다.
- B14용 formula overlay는 구현과 테스트가 끝났지만 실제 AI/import는
  아직 실행하지 않았다.

### 현재 중단 지점

- B24 승인 fragment: 23/28
  - missing: `[2, 3, 20, 27, 28]`
- B25 승인 fragment: 24/29
  - missing: `[2, 3, 11, 21, 29]`
- in-flight fragment는 0이며 live writer도 없는 상태였다.
- B14는 formula-safe staged plan 24 parts, 최대 prompt 357,901 bytes다.
- 대표 30개는 아직 30/30 통과가 아니다.
- answer-visible false-pass 9건이 남아 있다.
  - B05, B06, B11, B12, B17, B19, B23, B27, B30
- B05/B19/B27 Arm identity 보정과 B17 deterministic projector가
  구현 대기 상태다.

### 절대 금지

- 대표 30개 최신 계약 30/30 통과 전에 947개 전체 처리를 시작하지 않는다.
- 원본 Excel을 수정하지 않는다.
- Capture v2를 semantic 결과에 맞추어 변조하지 않는다.
- 기존 accepted fragment를 검증 없이 덮어쓰지 않는다.
- 실패한 model 결과를 validator 우회로 import하지 않는다.
- legacy FK 66건을 새 오류처럼 정리하거나 임의 수정하지 않는다.
- 사용자의 광범위한 dirty worktree를 reset, clean, checkout, 일괄 format하지 않는다.
- 명시적 요청 없이 WPF, Excel, GUI, 서버, preview를 실행하지 않는다.

자세한 작업 기록은 `SESSION_HANDOFF_2026-07-16.md`의 마지막
`사용량 소진에 따른 정지 체크포인트`를 우선 읽는다.

## 4. Claude가 먼저 읽을 파일

다음 순서로 읽고, `outputs` 전체를 무차별 로드하지 않는다.

1. 이 문서
2. `SESSION_HANDOFF_2026-07-16.md`의 마지막 두 섹션
3. `FINAL_GOAL_EXECUTION_PLAN.md`
4. `README.md`의 current semantic checkpoint
5. `inference_data_ai_semantic_ai.py`
6. `inference_data_ai_staged_runner_v2.py`
7. `inference_data_ai_workflow.py`
8. `inference_data_ai_staged_draft_v2.py`
9. 관련 집중 테스트

WPF의 기존 그룹화 AI까지 Claude backend로 바꿀 때만 다음 파일도 읽는다.

- `InferenceDataAIService.Wpf/MainWindow.xaml.cs`
- `InferenceDataAIService.Wpf/TwoLevelGroupAnalysisEngine.cs`

이미 검증된 989개 catalog/render 결과를 단순 점검하거나 다시 렌더링하기
위해 grouping AI를 재호출하지 않는다.

## 5. 현재 Codex 결합 지점

Python의 핵심 결합 지점:

- `inference_data_ai_semantic_ai.py`
  - `_codex_command`
  - `run_codex_locator`
  - `run_codex_locator_batch`
  - `run_codex_study_draft`
- `inference_data_ai_staged_runner_v2.py`
  - `_codex_command`
  - `run_codex_study_fragment_v2`
- `inference_data_ai_workflow.py`
  - 기본 runner dependency
- `inference_data_ai_cli.py`
  - `--model`, `--reasoning-effort`와 ingest 명령 연결

WPF의 결합 지점:

- `InferenceDataAIService.Wpf/MainWindow.xaml.cs`
- `InferenceDataAIService.Wpf/TwoLevelGroupAnalysisEngine.cs`

현재 subprocess 계약은 대략 다음과 같다.

1. prompt를 UTF-8 표준 입력으로 전달한다.
2. JSON schema를 임시 파일로 전달한다.
3. 마지막 model message를 임시 파일로 받는다.
4. JSON 파싱 후 프로젝트 validator를 다시 통과시킨다.
5. 성공한 결과만 target artifact에 기록한다.

Claude provider에서도 1~5의 의미를 그대로 유지해야 한다.

## 6. 권장 provider 설계

먼저 다음과 같은 작은 공통 계약을 만든다.

```python
class SemanticAiProvider(Protocol):
    def generate_structured(
        self,
        *,
        prompt: str,
        schema: dict[str, Any],
        model: str,
        effort: str | None,
        timeout_seconds: int,
        call_kind: str,
    ) -> StructuredAiResult:
        ...
```

`StructuredAiResult`에는 최소한 다음 값을 둔다.

```text
payload
provider
model
effort
request_id
input_tokens
cache_creation_input_tokens
cache_read_input_tokens
output_tokens
latency_ms
stop_reason
```

기존 `run_codex_*` 함수는 즉시 삭제하거나 한꺼번에 rename하지 않는다.
우선 provider를 주입할 수 있게 만들고 기존 이름은 compatibility wrapper로
유지한다. 그래야 기존 테스트와 resume artifact 계약을 작은 단위로
이관할 수 있다.

## 7. 경로 A: Claude Code CLI bridge

Claude Pro/Max 구독으로 가장 빨리 이어가는 경로다. Claude Code 2.1.205
이상이 필요하며 최신 버전을 권장한다.

비대화형 호출의 기준 형태:

```powershell
claude -p `
  --bare `
  --tools "" `
  --strict-mcp-config `
  --disallowedTools "mcp__*" `
  --no-session-persistence `
  --model claude-sonnet-5 `
  --effort high `
  --output-format json `
  --json-schema "<JSON_SCHEMA>"
```

구현 시 prompt는 기존과 같이 stdin으로 전달한다. 응답 최상위 JSON의
`structured_output`만 project validator에 넘긴다.

필수 조건:

- `--tools ""`로 shell/read/edit 같은 built-in 도구를 제거한다.
- `--strict-mcp-config`와 `--disallowedTools "mcp__*"`로 별도 MCP 도구도
  제거한다. `--tools`만으로는 MCP 도구가 제거되지 않는다.
- `--no-session-persistence`로 workbook 간 대화 상태를 공유하지 않는다.
- `--bare`로 `CLAUDE.md`, auto memory, hook, plugin의 프롬프트 오염을 막는다.
- `--output-format json`과 `--json-schema`를 함께 사용한다.
- schema는 shell 문자열로 조립하지 말고 subprocess argument로 직접 전달한다.
- API key나 토큰을 로그, journal, artifact에 기록하지 않는다.
- CLI wrapper JSON 전체가 아니라 `structured_output`을 추출한다.
- CLI가 반환하는 usage/cost metadata가 있으면 journal에 보존한다.

`--json-schema`는 prompt와 달리 command argument로 전달된다. Windows
command-line 길이 제한을 다시 밟지 않도록 각 compact schema의 직렬화
길이를 사전 측정하고, 안전 범위를 넘으면 CLI bridge를 사용하지 말고
Messages API로 전환한다. schema를 잘라내거나 validator를 완화해서
길이를 맞추지 않는다.

구독 사용량을 쓰려면 `claude auth status`로 로그인 방식을 확인한다.
`ANTHROPIC_API_KEY`가 환경 변수에 있으면 Claude Code가 구독 대신 API
과금을 사용할 수 있으므로 반드시 확인한다. Pro/Max 사용량은 Claude
웹·앱·Claude Code가 공유한다.

이 bridge는 전환용이다. 947개 대량 처리의 영구 backend로 고정하기
보다는 provider 계약과 semantic 동등성을 먼저 검증하는 용도로 사용한다.

## 8. 경로 B: Anthropic Messages API

장기 권장 경로다. Python 공식 SDK와 Messages API의 structured outputs를
사용한다.

요청의 핵심 형태:

```python
response = client.messages.create(
    model=model,
    max_tokens=max_tokens,
    messages=[{"role": "user", "content": prompt}],
    output_config={
        "effort": effort,
        "format": {
            "type": "json_schema",
            "schema": schema,
        },
    },
)
payload = json.loads(response.content[0].text)
```

실제 SDK 버전의 타입과 parameter 위치는 구현 시 공식 문서와 설치된 SDK
signature를 다시 확인한다.

### Sonnet 5 주의사항

- model ID는 `claude-sonnet-5`로 고정한다.
- adaptive thinking이 기본 활성화다.
- 수동 `budget_tokens` thinking은 400 오류이므로 사용하지 않는다.
- `temperature`, `top_p`, `top_k`를 비기본값으로 설정하면 400 오류가
  나므로 보내지 않는다.
- `max_tokens`는 thinking과 최종 JSON이 함께 사용하는 hard limit이므로
  기존 출력 크기보다 충분히 크게 잡고 실제 사용량으로 조정한다.
- 새 tokenizer는 같은 텍스트에서 이전 Claude 세대보다 약 30% 많은
  token을 만들 수 있으므로 기존 추정치를 그대로 재사용하지 않는다.

### schema 호환성

Anthropic structured outputs도 지원하는 JSON Schema subset에 제한이 있다.
실제 유료 호출 전에 다음 schema를 전부 로컬 preflight한다.

- locator output
- batch locator output
- study draft output
- staged fragment v2 transport output

기존 staged-v2는 nested object의 `additionalProperties: false` 문제를 이미
한 번 겪었다. schema를 provider별로 임의 완화해 validator를 우회하지
말고, transport schema와 domain payload validation을 분리한 현재 원칙을
유지한다.

### retry

자동 retry는 429, 529, 일시적 5xx, 연결 timeout처럼 전송 계층 오류에만
bounded exponential backoff로 적용한다.

다음 오류는 같은 요청의 무한 retry 대상으로 삼지 않는다.

- schema 자체가 지원되지 않음
- prompt hash 불일치
- source revision/content hash 불일치
- project validator의 의미·근거 오류
- `max_tokens` 설계 부족

의미 검증 실패는 기존 focused repair 또는 fail-closed 경로로 보낸다.

### 비용 제어

- 모든 요청에서 usage를 journal에 기록한다.
- workbook, stage, part index, prompt hash, model, effort별 비용을 남긴다.
- 공통 system/contract 문구는 prompt caching 후보로 분리하되, prompt
  구조 변경 시 prompt version과 exact hash 계약을 함께 올린다.
- 대표 30 통과 뒤 25~50 canary에서 실제 토큰을 측정한다.
- 947개는 동기 API로 곧바로 실행하지 말고 Message Batches API 전환을
  검토한다. Batch는 표준 token 가격의 50%다.
- 실행 전 provider/account 차원의 hard budget과 애플리케이션 누적
  budget을 둘 다 둔다.

## 9. model routing

대표 30 gate까지는 품질 비교를 위해 locator와 draft 모두 Sonnet 5로
통일하는 편이 안전하다.

```text
LOCATOR            -> claude-sonnet-5 / high
MONOLITHIC_DRAFT   -> claude-sonnet-5 / high
STAGED_FRAGMENT_V2 -> claude-sonnet-5 / high
HARD_REPAIR        -> claude-opus-4-8 / high 또는 xhigh
DETERMINISTIC_FIX  -> AI 호출 금지
```

30/30 이후 Haiku locator 후보를 평가한다.

1. 동일한 30개 locator를 Sonnet과 Haiku로 비교한다.
2. exact chunk coverage, candidate recall, conclusion/narrative heading recall,
   false positive를 비교한다.
3. 모든 gate가 Sonnet과 동등할 때만 PENDING locator에 Haiku를 허용한다.
4. draft와 fragment는 계속 Sonnet을 유지한다.

모델을 artifact 중간에서 바꾸면 provider/model/effort를 resume identity와
journal에 포함한다. 기존 accepted fragment를 다른 모델 결과로 자동
재작성하지 않는다.

## 10. 구현 순서

1. read-only로 live process, fragment inventory, journal 상태를 재확인한다.
2. 공통 `SemanticAiProvider`와 result/usage 계약을 추가한다.
3. 기존 Codex provider를 그 계약 뒤로 옮겨 기존 테스트를 먼저 통과시킨다.
4. Claude Code CLI bridge 또는 Anthropic API provider 하나를 구현한다.
5. fake subprocess/fake API 기반 집중 테스트를 추가한다.
6. 실제 artifact를 쓰지 않는 fixture 1건으로 schema·Unicode·usage 파싱을
   검증한다.
7. 사용자의 명시적 승인과 budget 설정 후 실제 AI smoke call 1건을 한다.
8. B24/B25 missing fragment만 resume한다.
9. Arm/B17 deterministic 구현과 false-pass 9건을 복구한다.
10. B14 formula-safe staged run을 수행한다.
11. 대표 30 최신 계약 30/30을 증명한다.
12. DB backup, hash freeze, quick-check, legacy FK66 불변을 확인한다.
13. PENDING947 exact manifest를 만든다.
14. 25~50 canary를 size/formula tier별로 실행한다.
15. canary 실측 비용과 실패율을 사용자에게 보고한 뒤에만 전체 migration을
    승인받는다.

## 11. 검증 규칙

코드 변경 후에는 변경 파일과 직접 관련된 가장 좁은 검증만 실행한다.

Python provider 이관의 우선 테스트:

```powershell
python -m unittest `
  tests.test_inference_data_ai_semantic_ai `
  tests.test_inference_data_ai_staged_draft_v2 `
  tests.test_inference_data_ai_workflow `
  tests.test_inference_data_ai_cli
```

실제 변경 범위가 더 작으면 해당 test module만 실행한다. 관련 Python
파일에는 `python -m py_compile`을 추가할 수 있다.

WPF C# 결합 지점을 변경했을 때만 다음 project build를 실행한다.

```powershell
dotnet build .\InferenceDataAIService.Wpf\InferenceDataAIService.Wpf.csproj --no-restore
```

검증을 위해 앱, WPF, Excel, GUI, 서버를 실행하지 않는다. 실제 Claude
호출은 테스트 명령이 아니라 유료·외부 상태 변경으로 취급하고 사용자의
budget 승인을 받은 뒤 실행한다.

## 12. 인수 직후 Claude Code 시작 방법

Claude Code 버전과 인증을 먼저 읽기 전용으로 확인한다.

```powershell
claude --version
claude auth status
```

초기 인수와 provider 설계:

```powershell
claude --model claude-opus-4-8 --effort xhigh
```

계획 후 일반 구현으로 자동 전환하려면 Claude Code 안에서:

```text
/model opusplan
```

첫 메시지는 다음처럼 준다.

```text
CLAUDE_MIGRATION_HANDOFF_2026-07-20.md를 먼저 전부 읽고,
SESSION_HANDOFF_2026-07-16.md의 마지막
"사용량 소진에 따른 정지 체크포인트"를 확인하라.

지금은 앱, WPF, Excel, 서버, Codex, Claude 유료 호출을 실행하지 말라.
dirty worktree의 기존 변경을 보존하라.

먼저 현재 Codex subprocess 결합 지점을 provider 계약으로 분리하는
최소 변경안을 제시하고, 변경할 파일과 좁은 테스트를 명시하라.
대표 30의 30/30 통과 전에는 PENDING947을 처리하지 말라.
```

## 13. 현재 공식 모델·가격 참고

2026-07-20 기준 Claude API 표준 가격:

| 모델 | 입력 / MTok | 출력 / MTok | 비고 |
|---|---:|---:|---|
| Claude Fable 5 | $10 | $50 | 최고 성능, 장기 agent |
| Claude Opus 4.8 | $5 | $25 | 복잡한 agentic coding 권장 |
| Claude Sonnet 5 | $2 | $10 | 2026-08-31까지 도입 가격 |
| Claude Sonnet 5 | $3 | $15 | 2026-09-01부터 |
| Claude Haiku 4.5 | $1 | $5 | 고속·저비용 |

Batch 가격은 입력·출력 모두 표준의 50%다. 가격과 모델 제공 여부는
실행 직전에 다시 확인한다.

공식 참고:

- [Claude 모델 개요](https://platform.claude.com/docs/en/about-claude/models/overview)
- [Claude 모델 선택](https://platform.claude.com/docs/en/about-claude/models/choosing-a-model)
- [Claude API 가격](https://platform.claude.com/docs/en/about-claude/pricing)
- [Claude Sonnet 5 변경 사항](https://platform.claude.com/docs/en/about-claude/models/whats-new-sonnet-5)
- [Structured outputs](https://platform.claude.com/docs/en/build-with-claude/structured-outputs)
- [Prompt caching](https://platform.claude.com/docs/en/build-with-claude/prompt-caching)
- [Message Batches](https://platform.claude.com/docs/en/build-with-claude/batch-processing)
- [Claude Code model 설정](https://code.claude.com/docs/en/model-config)
- [Claude Code 비대화형 실행](https://code.claude.com/docs/en/headless)
- [Claude Code CLI reference](https://code.claude.com/docs/en/cli-reference)
