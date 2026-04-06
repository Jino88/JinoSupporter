# DataMaker Load DB Flow

이 문서는 `DataMaker`의 `Load DB` 버튼을 눌렀을 때 현재 코드가 실제로 수행하는 작업을 정리한 것이다.

기준 파일:

- `WorkbenchHost/Modules/DataMaker/MainWindow.xaml.cs`
- `WorkbenchHost/Modules/DataMaker/R6/PreProcessor/clDataProcessor.cs`
- `WorkbenchHost/Modules/DataMaker/R6/PreProcessor/clMakeAccTable.cs`
- `WorkbenchHost/Modules/DataMaker/R6/PreProcessor/clMakeProcTable.cs`

## Entry Point

UI 진입점은 아래 메서드다.

- `CT_BT_LOAD_Click()`
- `LoadProcess()`

흐름:

1. 사용자가 DB 파일을 선택한다.
2. `PathSelectedDB`에 선택 경로를 저장한다.
3. 로그/프로그레스바를 초기화한다.
4. `RunDataProcessing(PathSelectedDB, _processTypeCsvPath, _reasonCsvPath)`를 백그라운드에서 실행한다.
5. 완료 후 `SetReportAvailability(true)`로 `Get Report`를 활성화한다.
6. 이후 `MakeGroupUI()`가 호출되어 그룹 UI를 다시 만든다.

관련 코드:

- `MainWindow.xaml.cs:174`
- `MainWindow.xaml.cs:355`
- `MainWindow.xaml.cs:1196`

## Main Processing

`RunDataProcessing()`은 `clDataProcessor.ProcessData()`를 호출한다.

처리 순서:

1. DB 연결 초기화
2. `Routing` 테이블 적재 및 정규화
3. `Reason` 테이블 적재 및 정규화
4. `ACC` 테이블 생성

관련 코드:

- `MainWindow.xaml.cs:1196`
- `clDataProcessor.cs:42`

## Step 1: Initialize Database

`clDataProcessor.InitializeDatabase()`

수행 내용:

- DB 파일 존재 여부 확인
- `clSQLFileIO` 생성
- `ORG` 테이블 존재 여부 확인
- `ORG` row count 확인

이 단계에서는 실제 데이터 변경은 없다.

관련 코드:

- `clDataProcessor.cs:63`

## Step 2: Load CSV Tables

`clDataProcessor.LoadCsvTables()`

두 개의 기준 테이블을 다시 만든다.

### 2-1. Routing Table

`LoadProcessTypeTable()`

수행 내용:

1. `Routing` 테이블이 이미 있으면 삭제
2. `clMakeTxtTable`로 `_processTypeCsvPath`를 읽어 `Routing` 테이블 생성
3. 생성 후 row count 확인
4. `NormalizeProcessTypeTable()` 실행

`NormalizeProcessTypeTable()` 수행 내용:

1. `Routing` 전체를 `DataTable`로 다시 읽음
2. 각 행의 `ProcessName`, `ProcessType`에 `CONSTANT.Normalize()` 적용
3. `Routing` 테이블 삭제
4. 동일 스키마로 다시 생성
5. 정규화된 `DataTable` 전체를 다시 씀

즉 `Routing`은 현재 매번

- drop
- create
- load all rows
- normalize in memory
- drop
- create
- write all rows

순서로 처리된다.

관련 코드:

- `clDataProcessor.cs:101`
- `clDataProcessor.cs:131`

### 2-2. Reason Table

`LoadReasonTable()`

수행 내용:

1. `Reason` 테이블이 이미 있으면 삭제
2. `clMakeTxtTable`로 `_reasonCsvPath`를 읽어 `Reason` 테이블 생성
3. 생성 후 row count 확인
4. `NormalizeReasonTable()` 실행

`NormalizeReasonTable()` 수행 내용:

1. `Reason` 전체를 `DataTable`로 다시 읽음
2. 각 행의 `processName`, `NgName`에 `CONSTANT.Normalize()` 적용
3. `Reason` 테이블 삭제
4. 동일 스키마로 다시 생성
5. 정규화된 `DataTable` 전체를 다시 씀

즉 `Reason`도 `Routing`과 같은 패턴으로 전체 재작성을 한다.

관련 코드:

- `clDataProcessor.cs:174`
- `clDataProcessor.cs:208`

## Step 3: Create AccessTable

`clDataProcessor.CreateAccessTable()`

실제 작업은 `clMakeAccTable.Run()`에서 수행된다.

전체 순서:

1. 기존 `ACC` 구조 확인
2. 기존 `ACC` 삭제
3. 새 `ACC` 스키마 생성
4. `ORG -> ACC` 데이터 복사
5. 파생 컬럼 채우기
6. `Routing` 기준 `ProcessType` 매핑
7. `Reason` 기준 `Reason` 매핑

관련 코드:

- `clDataProcessor.cs:261`
- `clMakeAccTable.cs:39`

### 3-1. ValidateTableStructure

기존 `ACC`가 있으면 컬럼 목록을 보고 `NGCODE`가 있는지 확인한다.

이 단계는 경고 로그 용도에 가깝다.

관련 코드:

- `clMakeAccTable.cs:72`

### 3-2. Drop Existing ACC

기존 `ACC`가 있으면 삭제한다.

추가 동작:

- 삭제 후 실제로 사라졌는지 다시 확인
- 실패 시 연결을 다시 열고 한 번 더 시도

관련 코드:

- `clMakeAccTable.cs:89`

### 3-3. Create ACC Schema

`clOption.GetAccessTableColumns()` 정의를 사용해 `ACC` 테이블을 생성한다.

관련 코드:

- `clMakeAccTable.cs:126`

### 3-4. Copy ORG -> ACC

현재 `Load DB`에서 가장 무거운 구간 중 하나다.

`CopyDataFromOriginalTable()` 수행 내용:

1. `ORG` 전체를 `DataTable`로 읽음
2. 비어 있는 `ACC`를 `DataTable`로 읽음
3. `SelectRowsForAccTable(orgData)`로 중복 제거 대상 선별
4. 선택된 각 행마다:
   - 새 `DataRow` 생성
   - 공통 컬럼 복사
   - `ProcessName`, `NGName`은 `Normalize` 적용
5. 완성된 `accData` 전체를 SQLite에 다시 씀

현재 구현 특징:

- 메모리에서 모든 행을 다룬다.
- `DataRow`를 행마다 새로 만든다.
- 마지막 저장은 SQLite insert 반복으로 진행된다.

관련 코드:

- `clMakeAccTable.cs:146`

### 3-5. Duplicate Filtering

`SelectRowsForAccTable()`는 `ORG` 데이터 중 중복 후보를 그룹핑한다.

중복 판정 키:

- `PRODUCTION_LINE`
- `PROCESSCODE`
- `PROCESSNAME` 정규화값
- `NGNAME` 정규화값
- `MATERIALNAME`
- `PRODUCT_DATE`
- `SHIFT`

중복 그룹 처리:

- 동일값이면 자동 1행 유지
- `QTYINPUT` 같고 `QTYNG`가 `0 / 비0`이면 비0 자동 선택
- `QTYINPUT` 같고 `QTYNG` 둘 다 비0이면 병합
- 그 외는 사용자 팝업으로 선택

즉 `Load DB` 중 사용자 입력이 발생할 수 있는 경로가 여기에 있다.

관련 코드:

- `clMakeAccTable.cs:180`
- `clMakeAccTable.cs:255`

### 3-6. Set Derived Columns

`SetDerivedColumns()`는 `ACC`에 파생 컬럼을 채운다.

설정 대상 예:

- `MONTH`
- `WEEK`
- `LINE_REMOVE`
- `LINE`
- `LINESHIFT`
- `LR`
- `LRLINE`
- `LR_BUILDING`

실제 계산은 `sql.Processor.SetEmptyColumnsValueInProcTable(OPTION_TABLE_NAME.ACC)`에서 처리된다.

관련 코드:

- `clMakeAccTable.cs:624`

### 3-7. Map ProcessType

`MapProcessType()`는 `ACC`와 `Routing`을 조인해 `ProcessType`을 채운다.

매칭 키:

- `ACC.MaterialName` = `Routing.모델명`
- `ACC.ProcessCode` = `Routing.ProcessCode`
- `ACC.ProcessName` = `Routing.ProcessName`

업데이트 대상:

- `ACC.ProcessType`

실제 조인은 `UpdateTableFromTable()`을 사용한다.

관련 코드:

- `clMakeAccTable.cs:635`
- `clSQLiteDataProcessing.cs:332`

### 3-8. Map Reason

`MapReason()`은 `ACC`와 `Reason`을 조인해 `Reason` 컬럼을 채운다.

매칭 키:

- `ACC.ProcessName` = `Reason.processName`
- `ACC.NGNAME` = `Reason.NgName`

업데이트 대상:

- `ACC.Reason`

현재 추가 작업:

- 매핑 전 `ACC` 전체를 읽어서 빈 `Reason` 수 계산
- `Reason` 전체를 읽어 샘플 로그 출력
- 매핑 실행
- 매핑 후 `ACC` 전체를 다시 읽어서 결과 통계 계산
- 매핑 실패 샘플 조합 출력

즉 실제 업데이트 외에 전체 테이블 재로딩이 여러 번 포함된다.

관련 코드:

- `clMakeAccTable.cs:667`

## After RunDataProcessing

`LoadProcess()`가 끝난 뒤 UI에서는 `MakeGroupUI()`가 호출된다.

`MakeGroupUI()` 내부 후반부에서 추가로 수행되는 것:

1. `ValidateAndRegenerateAccessTable()`
2. `clMakeProcTable.Run(allModels)`

즉 사용자가 체감하는 `Load DB` 이후 준비에는 `ACC` 생성뿐 아니라 `PROC` 재생성도 포함된다.

관련 코드:

- `MainWindow.xaml.cs:355`
- `MainWindow.xaml.cs:1189`

## ValidateAndRegenerateAccessTable

이 메서드는 `ACC`에 `NGCODE` 컬럼이 없으면 `ACC`를 다시 만든다.

현재 정상적인 최신 스키마 DB라면 보통 재생성은 발생하지 않는다.

관련 코드:

- `MainWindow.xaml.cs:1238`

## PROC Table Generation

`clMakeProcTable.Run(allModels)`

목적:

- `ACC`를 바탕으로 `PROC` 테이블 생성/갱신
- 이후 그룹 생성과 리포트 작업의 기반 테이블 준비

즉 `Load DB`의 최종 산출물은 사실상 아래 4개다.

- `Routing`
- `Reason`
- `ACC`
- `PROC`

관련 코드:

- `MainWindow.xaml.cs:1191`
- `WorkbenchHost/Modules/DataMaker/R6/PreProcessor/clMakeProcTable.cs`

## Current Performance Hotspots

현재 코드 기준 병목 후보는 아래와 같다.

### High

- `clMakeAccTable.CopyDataFromOriginalTable()`
  - `ORG` 전체 `LoadTable`
  - `ACC` 전체 `LoadTable`
  - 메모리 `DataRow` 복사
  - 최종 SQLite insert 반복

- `clSQLiteWriter.InsertWithBatchValues()`
  - 트랜잭션은 사용하지만 실제로는 행마다 `ExecuteNonQuery()`

### Medium

- `NormalizeProcessTypeTable()`
  - `Routing` 전체 로드 후 drop/create/write

- `NormalizeReasonTable()`
  - `Reason` 전체 로드 후 drop/create/write

- `MapReason()`
  - 통계/로그 출력을 위해 `ACC` 전체를 전후 두 번 다시 읽음

### Additional

- `UpdateTableFromTable()`
  - temp table 기반이라 구조는 나쁘지 않지만, 테이블 크기가 커지면 인덱스 전략이 중요

## Summary

`Load DB`는 단순 로드가 아니다. 현재는 아래 전체를 수행한다.

1. DB 열기 및 `ORG` 확인
2. `Routing` CSV 적재 및 전체 정규화 재작성
3. `Reason` CSV 적재 및 전체 정규화 재작성
4. `ACC` 삭제 후 전체 재생성
5. `ORG`에서 중복 처리 포함 복사
6. 파생 컬럼 계산
7. `Routing` 조인으로 `ProcessType` 매핑
8. `Reason` 조인으로 `Reason` 매핑
9. `PROC` 테이블 재생성

즉 `Load DB` 시간의 대부분은 파일 선택 자체가 아니라, 전처리와 테이블 재생성에 사용된다.
