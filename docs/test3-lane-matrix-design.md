# Test 3 레인 매트릭스 편집기 설계

## 1. 범위와 현재 구조

대상 화면은 `/bmes/test3`의 `BmesTest3Page.razor`이다. 현재 페이지는 Blazor Server가 데이터·저장·BOM 조회를 담당하고, `OnAfterRenderAsync`에서 `/test3-editor/test3-editor.js`를 동적 import한 뒤 `mount(host, DotNetObjectReference, BuildReactEditorSnapshot())`로 React를 붙인다. 현재 React가 호출하는 JSInvokable은 다음 여섯 개다.

| 메서드 | 현재 역할 |
|---|---|
| `ReactSelectModelAsync(string?, bool)` | 모델/L·R 선택, BOM 로드, 스냅샷 반환 |
| `ReactReloadBomAsync()` | BOM 강제 새로고침 |
| `ReactSaveLayoutAsync(Test3EditorLayoutRequest)` | 레인별 공정 순서·합류 저장 |
| `ReactSaveMaterialAsync(Test3EditorMaterialRequest)` | 자재의 최초 투입 공정과 수량/단위 저장 |
| `ReactClearMaterialAsync(string)` | 자재 매핑 해제 |
| `ReactResetLayoutAsync()` | 선택 모델의 공정 배치 초기화 |

`BuildReactEditorSnapshot`은 `AllModelProcessRows`와 `_processSettings`를 조합한다. 모델마다 저장된 행을 `LaneCode`로 묶고 `ProcessNo`의 숫자를 레인 내 순서로 사용한다. `MergeProcessNo`와 같은 MAIN 레인의 `ProcessNo`를 비교해 `MergeTargetProcessId`를 복원한다. 저장되지 않은 공정은 `AvailableProcesses`로 보낸다. BOM은 `BmesBomCacheService`가 캐시한 `BmesBomMaterialCandidate` 목록에서 공급하고, 모델/좌우 토큰을 기준으로 선택 모델에 귀속한다.

현재 저장소는 질문에 적힌 SQLite 스키마가 아니라 `ProcessMaterialMappingService`가 관리하는 두 JSON 파일이다.

- 공정 배치: `process-material-processes.json`의 `ProcessMaterialProcessRow`.
- 자재 매핑: `process-material-mappings.json`의 `ProcessMaterialMappingRow`.
- 두 파일 모두 서비스의 메모리 캐시와 잠금 아래 `JsonSerializer`로 읽고 쓴다.

공정 저장은 `ProcessNo = {LaneCode}-{순번}`(예: `SUB1-2`), `ReferenceProcessNo = 같은 레인의 바로 앞 순번`으로 기록한다. `LaneCode`는 `MAIN` 또는 `SUB1`~`SUB99`이고, SUB 레인의 마지막 공정만 `MergeProcessNo = MAIN-{대상 순번}`을 갖는다. 과거 `ProcessNo`만 있는 행은 `ResolveLaneCode`가 접두사로 역산한다. `ProcessMaterialProcessRow`는 `Id`, `ModelName`, `ProcessCode`, `ProcessName`, `ProcessNo`, `ReferenceProcessNo`, `LaneCode`, `MergeProcessNo`, `CreatedAt`, `UpdatedAt`을 가진다.

자재 저장은 선택한 최초 투입 공정의 `LinkedProcesses`에 해당하는 실제 공정마다 `ProcessMaterialMappingRow`를 펼쳐 쓴다. 각 행에는 `ModelName`, `ProcessCode`, `ProcessName`, `RawMaterialCode`, `RawMaterialName`, `UsageQty`, `UsageUnit`, `Note`, 시간/ID가 들어간다. `Note`에는 현재 `First input #{공정번호}; applies through final process`가 들어가며, 화면의 이후 공정 칩은 이 펼쳐진 매핑을 계산해 보여준다. 따라서 현재 의미는 “자재를 셀마다 독립 저장”이 아니라 “최초 투입점을 저장하고 이후 공정까지 적용”이다.

현재 UI는 ① 저장된 공정 순서, ② 검색된 모델의 공정, ③ BOM 자재 최초 투입 공정의 3컬럼이다. 공정과 자재를 같은 셀에서 함께 편집할 수 없고, 레인들이 세로로 분리되어 전체 스텝의 가로 관계와 L/R 비교가 어렵다. SUB 산출물을 MAIN의 특정 셀로 끌어오는 상호작용도 레인 헤더의 별도 합류 조작에 의존한다.

## 2. 새 UX 정의

### 2.1 데이터 화면 모델

화면은 선택한 모델의 L/R 두 `SideBoard`를 동시에 렌더링한다. 각 보드는 고정된 좌측 레인 헤더 열과 `MAIN`, `SUB1`, `SUB2`, `SUB3`… 레인 컬럼을 가진다. 컬럼은 가로로 놓고, 스텝은 위에서 아래로 증가한다. 모든 레인 컬럼의 행 높이는 보드의 `maxStepCount`에 맞추며, 빈 셀도 드롭 영역으로 남긴다.

셀은 `{side, laneCode, step}`으로 식별한다. 셀 안에는 다음이 공존할 수 있다.

- 공정 카드: 공정 코드, 공정명, 공정 유형, 순번 배지.
- 자재 칩: `자재코드 · 자재명 · 수량 단위`; 최초 투입 칩은 진한 색, 이후 자동 적용 칩은 옅은 색으로 표시한다.
- 공정 없이 자재만 있는 상태도 허용한다. 다만 저장 시 자재의 최초 투입 대상은 반드시 유효한 공정 카드여야 한다.

누적 표기는 저장하지 않고 계산한다. 같은 레인의 현재 스텝까지의 최초 투입 자재를 `+`로 이어 `MTR1+MTR2+MTR3`처럼 만든다. SUB의 산출물은 `SUB1` 접두사를 붙여 MAIN 합류 셀에서 `MTR1+MTR2+MTR3 +SUB1`로 렌더링한다. 저장에는 최초 투입 공정과 기존 매핑 행만 남긴다.

### 2.2 조작

- 공정 팔레트: 각 보드 오른쪽의 접이식 `검색된 모델의 공정` 서랍. 드래그하거나 `추가` 버튼으로 빈 셀/레인 끝에 배치한다.
- BOM 자재 팔레트: 공정 팔레트 아래의 `BOM 자재` 서랍. 자재를 공정 카드에 드롭하면 해당 공정이 최초 투입점이 된다. 기존 최초 투입점은 칩의 `이동`으로 변경한다.
- SUB 합류: SUB 레인의 헤더에 있는 `MAIN으로 합류` 핸들을 드래그해 같은 보드의 MAIN 공정 셀에 놓는다. 놓인 대상이 `mergeTargetProcessId`가 된다. 또는 핸들 클릭 후 MAIN 셀 클릭으로 키보드/터치 대체 조작을 제공한다. 다른 보드로의 드롭은 거부한다.
- 레인 추가: 보드 헤더의 `+ SUB 레인`으로 빈 다음 번호를 만든다. `SUB1`~`SUB99`만 허용한다. 레인 삭제는 비어 있는 레인에서만 허용하고, 공정/자재가 있으면 먼저 배치 해제를 요구한다.
- 순서 변경: 공정 카드를 같은 레인 안에서 위/아래로 드래그한다. 다른 레인으로 드래그하면 레인과 순번이 동시에 바뀐다.
- L/R 동시 편집: L/R 보드의 저장·BOM 새로고침은 한 번의 작업으로 처리하되, 자재/합류 드롭은 동일 보드 안에서만 허용한다.

팔레트는 폭이 좁으면 오른쪽 서랍으로 접고, 매트릭스는 가로 스크롤한다. 상단에는 `모델`, `공정 검색`, `BOM 자재 검색`, `BOM 다시 불러오기`, `저장 상태`를 둔다. 저장은 조작 후 debounce 자동 저장하되 `저장 중`, `저장 완료`, `저장 실패`를 표시하고, 명시적 `저장` 버튼도 제공한다.

### 2.3 빈 상태와 오류

모델 미선택은 “기준 모델 또는 L/R 모델을 선택하세요.”, 공정 없음은 “오른쪽 공정 팔레트에서 끌어오세요.”, BOM 없음은 “현재 모델의 BOM 자재가 없습니다.”로 표시한다. 저장 충돌/알 수 없는 공정 ID/잘못된 합류 대상은 해당 셀을 오류 테두리로 표시하고 스냅샷의 `statusError`를 함께 보여준다. L/R 불일치 자재 드롭, 중복 공정, MAIN이 아닌 합류 대상은 저장하지 않는다.

### 2.4 ASCII 목업

```text
┌ Test 3 · 레인 매트릭스 ─ 모델 [기준모델________] [검색] [BOM 다시 불러오기] ─┐
│ 상태: 저장 완료                                      [공정 검색] [자재 검색] │
├──────────────────────────────────────────────────────────────────────────────┤
│ LEFT · MODEL-L                         │ RIGHT · MODEL-R                     │
│       MAIN              SUB1       SUB2│       MAIN              SUB1    SUB2│
│  1  ┌──────────────┐  ┌────────┐ ┌────┐│  1  ┌──────────────┐  ┌──────┐ ┌──┐│
│     │MTR1           │  │        │ │    ││     │MTR1           │  │      │ │  ││
│     │[MTR1]         │  │        │ │    ││     │[MTR1]         │  │      │ │  ││
│  2  ├──────────────┤  ├────────┤ ├────┤│  2  ├──────────────┤  ├──────┤ ├──┤│
│     │MTR2 [MTR1+...]│  │MTR4    │ │MTR6││     │MTR2           │  │MTR4  │ │M6││
│     │               │  │[MTR4]  │ │    ││     │               │  │      │ │  ││
│  3  ├──────────────┤  ├────────┤ ├────┤│  3  ├──────────────┤  ├──────┤ ├──┤│
│     │MTR3           │  │MTR4+M5 │ │M6+7││     │MTR3           │  │M4+M5 │ │M6││
│  4  ├──────────────┤  └────────┘ └────┘│  4  ├──────────────┤  └──────┘ └──┘│
│     │MTR1+MTR2+MTR3 +SUB1  ◀─ drop ─SUB1│     │MTR1+MTR2+MTR3 +SUB1       │
│     └──────────────┘  [합류 대상: MAIN 4]│     └──────────────┘             │
├──────────────────────────────┬───────────────────────────────────────────────┤
│ [ + SUB 레인 ]                 │ 공정 팔레트 / BOM 자재 팔레트 (접기 가능)     │
│ SUB3: Process1 → Process2      │ 미배치 공정: [Process3] [Process4]            │
└──────────────────────────────┴───────────────────────────────────────────────┘
```

## 3. Blazor ↔ React 계약

아래 C# DTO를 유일한 기준으로 삼는다. JS interop JSON은 camelCase 키를 사용한다(`PropertyNameCaseInsensitive` 입력은 유지하되 출력 키를 명시적으로 camelCase로 고정한다). `decimal`은 JSON 숫자, nullable 문자열은 `null` 허용이다.

### 3.1 스냅샷 DTO

```csharp
public sealed class Test3EditorSnapshot {
  public List<string> ModelNames { get; init; } = []; // modelNames: string[]
  public string SelectedModel { get; init; } = ""; // selectedModel: string
  public List<Test3EditorSide> Sides { get; init; } = []; // sides: Side[]
  public List<Test3EditorMaterial> Materials { get; init; } = []; // materials: Material[]
  public bool BomBusy { get; init; } // bomBusy: boolean
  public string BomSource { get; init; } = ""; // bomSource: string
  public string BomError { get; init; } = ""; // bomError: string
  public string StatusMessage { get; init; } = ""; // statusMessage: string
  public bool StatusError { get; init; } // statusError: boolean
}
public sealed class Test3EditorSide {
  public string ModelName { get; init; } = ""; // modelName
  public string SideLabel { get; init; } = ""; // sideLabel
  public List<Test3EditorLane> Lanes { get; init; } = []; // lanes
  public List<Test3EditorProcess> AvailableProcesses { get; init; } = []; // availableProcesses
}
public sealed class Test3EditorLane {
  public string LaneCode { get; init; } = "MAIN"; // laneCode
  public string MergeTargetProcessId { get; init; } = ""; // mergeTargetProcessId
  public string MergeTargetLabel { get; init; } = ""; // mergeTargetLabel
  public List<Test3EditorProcess> Processes { get; init; } = []; // processes
}
public sealed class Test3EditorProcess {
  public string Id { get; init; } = ""; // id
  public string ModelName { get; init; } = ""; // modelName
  public string ProcessCode { get; init; } = ""; // processCode
  public string ProcessName { get; init; } = ""; // processName
  public string ProcessType { get; init; } = ""; // processType
  public string LaneCode { get; init; } = ""; // laneCode
  public string ProcessNo { get; init; } = ""; // processNo
  public int Order { get; init; } // order
}
public sealed class Test3EditorMaterial {
  public string Id { get; init; } = ""; // id
  public string ModelName { get; init; } = ""; // modelName
  public string SideLabel { get; init; } = ""; // sideLabel
  public string MaterialCode { get; init; } = ""; // materialCode
  public string MaterialName { get; init; } = ""; // materialName
  public decimal UsageQty { get; init; } // usageQty
  public string UsageUnit { get; init; } = "PC"; // usageUnit
  public string AssignedProcessId { get; init; } = ""; // assignedProcessId
  public string AssignedProcessLabel { get; init; } = ""; // assignedProcessLabel
  public string ScopeLabel { get; init; } = ""; // scopeLabel
}
```

### 3.2 요청 DTO

```csharp
public sealed class Test3EditorLayoutRequest {
  public List<Test3EditorLaneRequest> Lanes { get; init; } = []; // lanes
}
public sealed class Test3EditorLaneRequest {
  public string ModelName { get; init; } = ""; // modelName
  public string LaneCode { get; init; } = "MAIN"; // laneCode
  public string MergeTargetProcessId { get; init; } = ""; // mergeTargetProcessId
  public List<string> ProcessIds { get; init; } = []; // processIds
}
public sealed class Test3EditorMaterialRequest {
  public string MaterialId { get; init; } = ""; // materialId
  public string ProcessId { get; init; } = ""; // processId
  public decimal UsageQty { get; init; } // usageQty
  public string UsageUnit { get; init; } = "PC"; // usageUnit
}
```

`Test3EditorLayoutRequest`는 선택된 L/R의 모든 레인을 한 번에 보내며, `ProcessIds`의 배열 순서가 1부터의 스텝이다. 빈 레인은 `ProcessIds: []`로 보낼 수 있다. `Test3EditorMaterialRequest.ProcessId`는 최초 투입 공정의 `Test3EditorProcess.Id`이고, 해제는 별도 메서드로 한다. DTO에는 계산된 누적 문자열을 보내지 않는다.

### 3.3 정확한 JSInvokable 시그니처

```csharp
[JSInvokable] public Task<Test3EditorSnapshot> ReactSelectModelAsync(string? model, bool force);
[JSInvokable] public Task<Test3EditorSnapshot> ReactReloadBomAsync();
[JSInvokable] public Task<Test3EditorSnapshot> ReactSaveLayoutAsync(Test3EditorLayoutRequest request);
[JSInvokable] public Task<Test3EditorSnapshot> ReactSaveMaterialAsync(Test3EditorMaterialRequest request);
[JSInvokable] public Task<Test3EditorSnapshot> ReactClearMaterialAsync(string materialId);
[JSInvokable] public Task<Test3EditorSnapshot> ReactResetLayoutAsync();
```

React mount 계약은 `mount(HTMLElement host, DotNetObjectReference dotnet, Test3EditorSnapshot initialSnapshot): void`, 해제는 `unmount(HTMLElement host): void`로 고정한다. 모든 메서드는 성공/실패 모두 최신 스냅샷을 반환하고, 실패는 `statusError: true`와 한국어 `statusMessage`로 표현한다.

## 4. 백엔드 저장 매핑

| UI 조작 | `ProcessMaterialProcessRow` 저장 | `ProcessMaterialMappingRow` 저장 |
|---|---|---|
| 공정을 레인 스텝에 배치 | `LaneCode`, `ProcessNo={lane}-{step}`, `ReferenceProcessNo={lane}-{step-1}`; 기존 식별 필드는 유지 | 변경 없음 |
| 공정 이동/순서 변경 | 해당 모델의 행을 삭제 후 새 번호로 Upsert, 미배치 행은 `DeleteProcess` | 기존 자재 매핑은 공정 식별자 기준으로 유지; 필요 시 기존 매핑의 공정 코드가 실제 대상과 일치하는지 검증 |
| SUB 합류 | SUB 마지막 행의 `MergeProcessNo=MAIN-{targetStep}`; 나머지는 빈 문자열 | 변경 없음. 누적/합류 표기는 조회 계산 |
| 자재를 공정 셀에 드롭 | 변경 없음 | 최초 투입 공정의 `ProcessCode/ProcessName`별로 `RawMaterialCode/Name`, `UsageQty`, `UsageUnit`을 Upsert하고 `Note`에 최초 투입/적용 범위를 기록 |
| 자재를 다른 공정으로 이동 | 변경 없음 | 기존 해당 모델·자재 매핑 Delete 후 새 최초점부터 재생성 |
| 자재 해제 | 변경 없음 | 현재 모델·자재에 해당하는 매핑을 모두 Delete |

현 JSON 구조에서는 스텝 셀에 자재를 표시하기 위해 스키마를 추가할 필요가 없다. 최초 투입점은 기존 `ProcessMaterialMappingRow`가 정확히 표현하고, 이후 칩과 `MTR1+MTR2+MTR3`는 `ProcessNo` 순서와 매핑 집합으로 결정론적으로 계산할 수 있다. 이 방식이 기존 데이터를 보존하고 구버전 행도 읽을 수 있는 우선안이다. 단, 향후 “특정 후속 공정에서만 제외” 또는 같은 자재의 다중 최초점이 요구되면 그때 `ProcessMaterialMappingRow`에 `FirstInputProcessNo`(TEXT), `LaneCode`(TEXT)를 추가하고 JSON 역직렬화 기본값을 빈 문자열로 두는 마이그레이션을 별도 도입한다. 현재 요구에는 추가하지 않는다.

현재 코드의 BOM 캐시도 변경하지 않는다. `BmesBomCacheService`의 후보는 `ProductCode`, `MaterialCode`, `MaterialName`, `UsageQty`, `UsageUnit`, 계층/부모 정보를 제공하며, React에는 편집에 필요한 평탄 DTO만 보낸다.

## 5. 레거시 제거 범위

`ReactEditorEnabled == false`가 현재 항상 거짓인 구조이므로, 새 React 매트릭스가 안정화된 뒤 페이지의 해당 블록 전체(상단 제목 중복, Blazor 검색/필터, 3컬럼 마크업, 표 fallback, Blazor drag/drop 마크업)를 삭제한다. 그 블록에만 쓰이는 것으로 정리할 대상은 다음이다.

- 렌더 전용 계산: `BuildProcessLanes`, `BuildMaterialLanes`, `BuildModelSideClass`, `BuildModelSideLabel`, `BuildProcessDragRowClass`, `BuildProcessTitle`, `BuildProcessEffectiveRange`, `BuildProcessSelectLabel`, `BuildMaterialAssignmentTitle`, `BuildMaterialFirstProcessText`, `BuildMaterialScopeText`, `GetMaterialProcessRows`(React가 새 셀 모델을 사용한 뒤), `HasMaterialProcessSelection`, `GetMaterialProcessValue`, `GetMaterialUsageQtyText`, `GetMaterialUsageUnit`, `CanSaveMaterialAssignment`.
- fallback 전용 이벤트/상태: `StartProcessDrag`, `EndProcessDrag`, `DragOverProcessRow`, `ClearDragOverProcessRow`, `DropProcessAtSequenceEnd`, `DropProcessOnSequence`, `DropProcessBackToAvailable`, `AddProcessToSequence`, `RemoveProcessFromSequence`, `_draggedProcess`, `_dragOverProcess`, 그리고 fallback 전용 필터/표 상태.

다음은 삭제하면 안 된다. React snapshot/interop와 저장에 사용된다.

- `ReactEditorEnabled`를 제거하기 전까지는 mount 조건과 호스트 수명 코드를 정리해야 하지만, `_reactEditorHost`, `_reactEditorReference`, `_reactEditorModule`, `OnAfterRenderAsync`, `DisposeAsync`는 유지한다.
- `BuildReactEditorSnapshot`, `SaveReactLayout`, `BuildMaterialAssignments`, `SaveMaterialAssignmentAsync`, `ClearMaterialAssignmentAsync`, `React*Async`, `ResolveLaneCode`, `ResolveLaneOrder`, `NormalizeProcessNo`, `BuildLaneProcessNo`, 모델 선택/ BOM 로드 상태와 `ProcessMaterialMappingService` 호출은 공용/React 경로다.
- `MaterialAssignment`, `_rows`, `_processSettings`, `_modelNames`, `_filterModel`, `_bomBusy`, `_bomError`, `_bomLoadSource`, `_statusMessage`, `_statusError`는 snapshot 구성 또는 저장 검증에 쓰이므로 실제 참조를 확인하기 전에는 삭제하지 않는다.

레거시 제거는 새 React 기능 구현과 같은 커밋에서 하지 말고, 통합 검증 후 별도 정리 커밋으로 한다.

## 6. 후속 구현 분할과 파일 소유권

승인된 체인은 방향이 맞다. 다만 계약을 먼저 고정해야 하므로 아래 4조각으로 진행한다. 각 조각은 표에 없는 파일을 수정하지 않는다.

| 조각 | 책임 | 배타적 파일 범위 |
|---|---|---|
| (A) 계약/설계 기준 | 이 문서의 DTO·JSON 키·저장 규칙을 기준선으로 검토하고 구현 체크리스트 작성 | `docs/test3-lane-matrix-design.md` |
| (B) C#/Blazor 호스트 | DTO 개편, snapshot 생성, JSInvokable 요청 검증/저장, React mount. 필요할 때만 서비스 저장 보정 | `JinoSupporter.Web/Services/Test3EditorContracts.cs`, `JinoSupporter.Web/Components/Pages/BmesTest3Page.razor`, 필요 시 `JinoSupporter.Web/Services/ProcessMaterialMappingService.cs` |
| (C) React UI | 매트릭스 보드, 셀/카드/칩, 팔레트, L/R 제한, drag/keyboard 대체 조작, 한국어 상태/반응형 스타일, 산출물 빌드 설정 | `JinoSupporter.Web/ClientApp/Test3Editor/**`, 필요 시 생성 산출물 `JinoSupporter.Web/wwwroot/test3-editor/**` |
| (D) 통합 정합성 | C# 직렬화 키와 React 접근 키, 저장 왕복, 기존 JSON 보존, 레거시 제거 범위, 좁은 빌드/테스트 결과 점검 및 필요한 수정 | B/C 소유 파일 중 실제 결함이 있는 파일만, 소유자 승인 후 |

(B)와 (C)는 각각 배타적으로 작업한 뒤 (D)가 양쪽을 함께 검증한다. D가 B/C 파일을 고칠 때는 원 소유 조각의 계약 위반 또는 통합 결함으로 한정한다. 루트 세션이 최종 빌드를 실행하며, 구현 세션은 서버를 띄우지 않는다.
