# Ask AI / ReviewCase CLI Handoff - 2026-07-07

## 오늘까지 한 것

### 1. 어제 CLI 작업 히스토리 확인

- Codex CLI session:
  `C:\Users\jhbyun\.codex\sessions\2026\07\06\rollout-2026-07-06T06-28-25-019f349c-98f1-7803-adc5-e6c5ad7ed274.jsonl`
- 어제 작업은 웹앱이 아니라 CLI에서 ReviewCase batch/verification을 돌린 것.
- 전체 989개 Excel 기준:
  - verified 13
  - needs_review 917
  - excluded 59
- 당시 확인된 verified file id:
  `219, 330, 442, 477, 553, 711, 721, 775, 802, 854, 864, 893, 936`
- 원본 Excel/DB는 수정하지 않은 작업이었다.

### 2. Ask AI가 verified ReviewCase를 우선 보도록 연결

수정 파일:

- `JinoSupporter.Web/Services/MicroSpeakerAskEvidenceService.cs`
- `AI_PROMPTS/data-inference/ask-ai-cli.md`
- `AI_PROMPTS/data-inference/cli-ask-ai.md`

핵심:

- `REVIEWCASE_AI_DRAFTS/verified/reviewcase_ai_verification_manifest.json`
  와 각 `*.reviewcase-ai-verification.json` / draft JSON을 읽어서
  `microSpeaker.verifiedReviewCases`로 Ask AI evidence pack에 넣었다.
- Ask AI 프롬프트에 verified ReviewCase를 raw row보다 먼저 보라고 명시했다.
- source file link, changedFactors, outcomes, evidenceRows, verification issue/limitation을 함께 넘긴다.

### 3. Ask AI 동시 실행 tmp 충돌 수정

수정 파일:

- `JinoSupporter.Web/Components/Pages/DataInferenceAskPage.razor`
- `AI_PROMPTS/data-inference/ask-ai-cli.md`
- `AI_PROMPTS/data-inference/cli-ask-ai.md`

기존 문제:

- Ask AI를 동시에 2개 돌리면 둘 다 `tmp/ask_request.json`,
  `tmp/ask_result_payload.json` 같은 공유 파일명을 사용해서 한쪽 CLI가 tmp 파일을 못 찾거나 잘못 읽을 수 있었다.

수정:

- 실행마다 `runId`를 만들고,
  `tmp/ask_runs/<runId>/ask_request.json`,
  `tmp/ask_runs/<runId>/ask_result_payload.json`,
  `tmp/ask_runs/<runId>/ask_commit.py`
  를 쓰게 바꿨다.
- 프롬프트 cleanup 규칙도 현재 run 폴더 안의 파일만 지우게 바꿨다.
- 다른 run 폴더나 공유 tmp 파일은 지우지 않도록 명시했다.

### 4. 모델 경계 규칙 추가

수정 파일:

- `JinoSupporter.Web/Services/MicroSpeakerAskEvidenceService.cs`
- `AI_PROMPTS/data-inference/ask-ai-cli.md`
- `AI_PROMPTS/data-inference/cli-ask-ai.md`

문제:

- VP FILM LOT처럼 여러 모델에 있을 수 있는 항목이,
  verified ReviewCase가 있는 1개 모델만 답하고 끝나는 문제가 있었다.

수정:

- evidence pack에 `microSpeaker.modelCoverage`를 추가했다.
- 질문에 걸린 모델을 checklist처럼 보여주고,
  모델별 evidence level을 나누도록 했다.
  - `verified`
  - `fallbackRows`
  - `candidateFilesOnly`
- 프롬프트에 여러 모델이 있으면 첫 verified 모델에서 멈추지 말고
  다른 모델도 표시하라고 명시했다.
- 빌드 확인 완료:
  `dotnet build .\JinoSupporter.Web\JinoSupporter.Web.csproj --no-restore -p:UseAppHost=false -p:OutputPath='.\artifacts\codex-check\ask-ai-model-coverage\'`
  결과: warning 0, error 0

### 5. 화면 테스트에서 확인한 애매한 점

VP FILM LOT 질문:

- 질문: `VP FILM LOT이 function defect에 영향 있었는지 검토해줘`
- 화면상 BRS-201506 하나만 강하게 답이 나왔다.
- 하지만 DB 기준으로 VP FILM LOT 관련 후보는 여러 모델에 있었다.
  예:
  - `BRS-201506`
  - `TIU-C11-20`
  - `L20S15-07 / MSU-L20S15-07`
- 그래서 단순 Ask AI 답변만 믿기보다, ReviewCase를 CLI에서 파일별로 보며 모델명을 확정하는 방식이 필요하다.

FP+VP bonding amount / Magnet Lot 질문:

- 모델 필드와 Magnet/YK evidence 분리가 이전보다 나아진 것으로 보였다.
- 다만 이것도 최종 원인 결론은 모델/조건 mapping 확정 후 보는 게 맞다.

## 방금 방향 정정한 것

사용자 의도:

- 새 파이썬 프로그램을 만들어 자동화하자는 뜻이 아니다.
- Codex CLI에서 Excel 하나씩 직접 읽고,
  ReviewCase와 모델명을 같이 확인하자는 뜻.
- 모델명이 뭔지 모르겠으면 사용자에게 물어보고,
  그 답을 기준으로 매핑 프롬프트/결정 기록을 만들어야 한다.

주의:

- 내가 중간에 `tools/review_reviewcase_cli.py` helper를 만들었지만,
  이 방향은 사용자가 원한 흐름이 아니다.
- 내일 시작 시 이 파일을 삭제할지, 그냥 사용하지 않고 둘지 먼저 정리하면 된다.
- `tools/generate_reviewcase_batch.py`에도 모델 후보 추출 관련 일부 변경이 들어갔다.
  이것도 자동화 방향으로 과할 수 있으니 내일 계속 쓸지 되돌릴지 판단해야 한다.

## 내일 해야 할 일

### 1. 먼저 작업 범위 정리

해야 할 선택:

- 오늘 추가한 `tools/review_reviewcase_cli.py`를 삭제할지 결정.
- `tools/generate_reviewcase_batch.py`에 넣은 모델 후보 추출 변경을 유지할지 되돌릴지 결정.
- Ask AI evidence/modelCoverage/tmp run-scoped 수정은 유지하는 방향.

### 2. CLI에서 Excel 1개씩 ReviewCase 확인

목표:

- 각 Excel에 대해 모델명을 먼저 확정한다.
- 모델명이 명확하지 않으면 사용자에게 짧게 질문한다.
- 그 다음 ReviewCase가 맞는지 확인한다.
- 확정된 내용은 draft/verification 또는 별도 mapping note에 남긴다.

기본 확인 순서:

1. 원본 파일명, `files.models`, `categories`, `term_summary` 확인
2. `sheet_rows`에서 title/purpose/content/result/decision rows 확인
3. 현재 batch draft 확인
4. 현재 AI verification 확인
5. 모델명 확정
6. changed factor / outcome / evidence row가 맞는지 확인
7. 애매하면 사용자에게 질문
8. 확정 내용을 기록

### 3. 우선순위 높은 케이스

VP FILM LOT부터 보는 것이 좋다.

관련 후보로 확인했던 file id:

- `21`: TIU-C11-20 VP new film 2-2 roll high function NG rate
- `286`: TIU-C11-20 쪽 VP new film 관련 파일
- `87`: L20S15-07 / MSU-L20S15-07 쪽 film AEM 관련 파일
- `439`: L20S15-07 / MSU-L20S15-07 쪽 test film 85A improve NG function 관련 파일

이 케이스에서 확인할 것:

- VP FILM LOT이 실제 모델별로 같은 의미인지
- `VP AEM Film`, `Film dry 24h`, `New Film`, `film roll`, `test film 85A`가 같은 축인지 다른 축인지
- function defect와 직접 같은 source/table에서 비교되는지
- process vision 악화 같은 부수 결과를 function defect와 섞지 않을 것

### 4. 모델명 물어보는 기준

사용자에게 물어볼 상황:

- `files.models`가 비어 있음
- 파일명에는 모델 후보가 있는데 DB model이 없음
- `BRS-2015`처럼 패밀리/축약명인지 실제 모델인지 애매함
- 한 파일에 여러 모델이 섞여 있음
- 제목/시트 row와 파일명 모델이 서로 다름

질문 예:

- `file_id 21은 TIU-C11-20으로 확정해도 됩니까?`
- `이 파일의 BRS-2015는 BRS-201506으로 매핑하면 됩니까?`
- `L20S15-07과 MSU-L20S15-07이 같은 모델 그룹으로 봐도 됩니까, 분리해야 합니까?`

### 5. 내일 실행할 때 유용한 명령

DB:

`D:\000. MyWorks\005. Program\Repository\MicroSpeaker_ProductTech_DB\db\InputDataFinish.sqlite`

ReviewCase draft:

`REVIEWCASE_AI_DRAFTS/batch/files/<file_id>.reviewcase-draft.json`

Verified result:

`REVIEWCASE_AI_DRAFTS/verified/files/<file_id>.reviewcase-ai-verification.json`

예시 조회:

```powershell
sqlite3 "D:\000. MyWorks\005. Program\Repository\MicroSpeaker_ProductTech_DB\db\InputDataFinish.sqlite" "select file_id, file_name, models, categories, term_summary from files where file_id=21;"
```

```powershell
sqlite3 "D:\000. MyWorks\005. Program\Repository\MicroSpeaker_ProductTech_DB\db\InputDataFinish.sqlite" "select sheet_name, row_number, row_text from sheet_rows where file_id=21 order by sheet_name, row_number limit 80;"
```

```powershell
Get-Content REVIEWCASE_AI_DRAFTS\batch\files\21.reviewcase-draft.json -TotalCount 220
```

```powershell
Get-Content REVIEWCASE_AI_DRAFTS\verified\files\21.reviewcase-ai-verification.json -TotalCount 160
```

## 운영 주의

- 사용자가 명시적으로 요청하지 않으면 웹앱, dev server, preview server를 실행하지 않는다.
- 빌드는 코드 수정 검증이 필요할 때만 좁게 실행한다.
- 원본 Excel과 MicroSpeaker SQLite는 직접 수정하지 않는다.
- dirty worktree가 많으므로 관련 없는 파일은 건드리지 않는다.
- Ask AI 화면 결과만 보고 확정하지 말고, source row와 모델명을 확인한다.
