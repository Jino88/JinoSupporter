# GraphMaker 구조 정리

## 1. 전체 개요

`GraphMaker`는 `WorkbenchHost/Modules/GraphMaker` 아래에 있는 WPF 기반 그래프 작업 공간입니다.

이 모듈은 하나의 그래프 기능만 있는 구조가 아니라, 여러 종류의 그래프 기능을 한 폴더 아래 묶어 둔 형태입니다.

큰 흐름은 보통 아래처럼 갑니다.

1. `MainWindow`에서 그래프 종류를 선택함
2. 각 그래프용 `View(UserControl)`가 열림
3. 파일을 읽어서 `DataTable` 또는 결과 모델로 변환함
4. 스펙(`Spec`), 상한(`USL`), 하한(`LSL`)을 읽음
5. 계산기(`Calculator`) 또는 View 내부 로직으로 통계/그래프 데이터를 만듦
6. 결과창(`ResultWindow`) 또는 큰 그래프 창에서 보여줌

---

## 2. 최상위 진입점

### 메인 창

- `MainWindow.xaml`
- `MainWindow.xaml.cs`

역할:

- 왼쪽 메뉴를 보여줌
- 사용자가 버튼을 누르면 오른쪽 `ContentArea`에 해당 그래프 `View`를 넣음

현재 연결 구조:

- `SPL Graph Plot` -> `ScatterPlotView`
- `Daily Sampling Scatter` -> `ValuePlotView`
- `SingleX(No) - Multi Y` -> `NoXMultiYView`
- `SingleX(Date) - Multi Y (No Header)` -> `DateNoHeaderMultiYView`
- `Generate Heatmap` -> `HeatMapView`
- `AudioBus Data` -> `AudioBusDataView`

즉, `MainWindow`는 실제 계산을 하지 않고, 그래프 기능별 화면으로 보내주는 허브 역할입니다.

---

## 3. 공통 데이터 모델

### `FileInfo_DailySampling`

정의 위치:

- `ValuePlot/ValuePlotView.xaml.cs`

이 클래스는 GraphMaker 안에서 가장 많이 재사용되는 파일/세션 데이터 모델입니다.

주요 내용:

- 파일 정보
  - `Name`
  - `FilePath`
  - `Delimiter`
  - `HeaderRowNumber`
- 실제 데이터
  - `FullData`
  - `Dates`
  - `SampleNumbers`
- 저장된 설정값
  - 색상
  - X축 모드
  - 표시 모드
- 컬럼별 제한값
  - `SavedColumnLimits`

쉽게 말하면:

- 파일 하나를 읽고
- 그 파일에 대한 그래프 설정과 limit 설정까지 같이 들고 다니는 공용 객체입니다.

---

## 4. 공통 기능 폴더

폴더:

- `Common/`

여기는 여러 그래프 기능이 같이 쓰는 공용 창, 공용 계산기, 공용 헬퍼가 모여 있는 곳입니다.

### 4-1. 공용 결과창 / 공용 상세창

#### `MultiColumnResultWindow`

역할:

- 다중 컬럼 그래프 결과를 카드 형태로 보여줌
- 오른쪽에 전체 불량 요약을 보여줌
- 현재는 추천 `USL/LSL` 값도 오른쪽 요약 영역에 표시함

주로 사용하는 곳:

- `ValuePlotMultiColumnView`
- `DateNoHeaderMultiYView`

#### `MultiColumnLargeDetailWindow`

역할:

- `MultiColumnResultWindow`에서 특정 컬럼을 클릭했을 때 열리는 큰 그래프 상세창

#### `DailySamplingGraphViewerWindow`

역할:

- 일일 샘플링 계열 그래프를 공통 방식으로 보는 창

#### `LargePlotWindowHelper`

역할:

- `OxyPlot.PlotModel`만 있으면 큰 그래프 창을 만들어주는 공용 헬퍼
- 시리즈 보이기/숨기기
- 범례 색상 편집
- 큰 화면용 4분할 보기 구성

### 4-2. 공용 계산기

#### `MultiColumnGraphCalculator`

역할:

- 다중 컬럼 그래프 계산을 담당하는 공용 계산기
- 여러 View가 이 계산기를 재사용함

출력 모델:

- `MultiColumnGraphResult`
- `ColumnGraphResult`
- `OverallGraphResult`

사용하는 곳:

- `ValuePlotMultiColumnView`
- `DateNoHeaderMultiYView`

### 4-3. 공용 보조 창

- `ColumnLimitSettingsWindow`
- `ColumnRenameWindow`
- `LimitValuesWindow`
- `TextInputWindow`

용도:

- 컬럼명 변경
- limit 값 입력
- 간단한 텍스트 입력
- 공통 편집 UI

---

## 5. 주요 그래프 모듈별 구조

## 5-1. ValuePlot

폴더:

- `ValuePlot/`

주요 파일:

- `ValuePlotView.xaml`
- `ValuePlotView.xaml.cs`
- `ValuePlotMultiColumnView.xaml`
- `ValuePlotMultiColumnView.xaml.cs`
- `ValuePlotFileSettingsWindow.xaml.cs`
- `ValuePlotMultiColumnFileSettingsWindow.xaml.cs`

역할:

- Daily Sampling 계열 그래프의 중심 모듈
- 단일 컬럼/다중 컬럼 그래프 둘 다 포함

### `ValuePlotView`

역할:

- 기본 daily sampling scatter/CPK 스타일 그래프
- 파일 로딩, 색상, spec/upper/lower 설정
- 산포도와 CPK 그래프를 직접 만듦

특징:

- `FileInfo_DailySampling`를 핵심 데이터로 사용
- View 내부에서 `PlotModel`을 직접 만드는 편

### `ValuePlotMultiColumnView`

역할:

- 여러 컬럼을 동시에 그리는 Value Plot 전용 화면

흐름:

1. 파일 로딩
2. `DataTable` 생성
3. 상단 preview에 `Upper / Spec / Lower` 행 구성
4. 표시할 컬럼 선택
5. `MultiColumnGraphCalculator.Calculate(...)` 호출
6. `MultiColumnResultWindow` 열기

현재 이 경로에서 연결되는 창:

- `MultiColumnResultWindow`
- `MultiColumnLargeDetailWindow`

---

## 5-2. ScatterPlot

폴더:

- `ScatterPlot/`

주요 파일:

- `ScatterPlotView.xaml`
- `ScatterPlotView.xaml.cs`
- `ScatterFileSettingsWindow.xaml.cs`

역할:

- SPL 그래프 / 산포도 중심 기능

흐름:

1. X, Y 컬럼 선택
2. 필요하면 spec/reference 컬럼 선택
3. 그래프 데이터 생성
4. `GraphViewerWindow`에서 결과 표시

즉:

- ScatterPlot 계열은 `GraphViewerWindow`를 많이 타는 구조입니다.

---

## 5-3. NoXMultiY

폴더:

- `NoXMultiY/`

주요 파일:

- `NoXMultiYView.xaml`
- `NoXMultiYView.xaml.cs`
- `NoXMultiYGraphCalculator.cs`
- `NoXMultiYResultWindow.xaml`
- `NoXMultiYResultWindow.xaml.cs`
- `NoXMultiYLimitRecommendationWindow.xaml`
- `NoXMultiYLimitRecommendationWindow.xaml.cs`

역할:

- X축 없이 여러 Y 컬럼을 비교하는 그래프
- 컬럼 자체를 하나의 영역/카테고리로 보고 분포를 그림

흐름:

1. 파일 로딩
2. 컬럼 limit 확보
3. 선택된 Y 컬럼 수집
4. `NoXMultiYGraphCalculator` 계산
5. `NoXMultiYResultWindow` 출력

계산 결과 모델:

- `NoXMultiYGraphResult`
- `NoXMultiYColumnResult`
- `NoXMultiYPoint`

특징:

- 컬럼별 평균, 표준편차, NG 수, NG율, CPK를 계산
- 컬럼별 정규분포 모양을 겹쳐 그림
- 별도 추천 limit 창도 존재

---

## 5-4. DateNoHeaderMultiY

폴더:

- `DateNoHeaderMultiY/`

주요 파일:

- `DateNoHeaderMultiYView.xaml`
- `DateNoHeaderMultiYView.xaml.cs`

역할:

- X축이 날짜이고, Y쪽 헤더 구조가 일반적이지 않은 파일을 다루는 다중 Y 그래프

흐름:

1. 파일 읽기
2. 선택한 컬럼들로 결합 파일 형태 구성
3. `MultiColumnGraphCalculator` 재사용
4. `MultiColumnResultWindow` 출력

즉:

- 자체 결과창을 따로 만들기보다 `Common`의 공용 결과창을 재사용하는 구조입니다.

---

## 5-5. SingleXSingleY

폴더:

- `SingleXSingleY/`

주요 파일:

- `DailyDataTrendView.xaml`
- `DailyDataTrendView.xaml.cs`
- `DailyDataTrendSetupWindow.xaml.cs`

역할:

- X 하나, Y 하나인 trend 그래프

특징:

- `USL / SPEC / LSL` 편집 UI가 있음
- 단일 값 시리즈 기준으로 trend를 그림
- daily trend 계열의 기본형에 가까움

---

## 5-6. MultiXMultiY

폴더:

- `MultiXMultiY/`

주요 파일:

- `DailyDataTrendExtraView.xaml`
- `DailyDataTrendExtraView.xaml.cs`

역할:

- 여러 X / 여러 Y를 다루는 확장 trend 기능

흐름:

1. 축/컬럼 조합 설정
2. trend 결과 생성
3. `ProcessFlowTrendResultWindow`로 넘김

즉:

- 구조적으로는 `ProcessTrend` 계열 결과창과 연결되는 중간 단계 역할이 큽니다.

---

## 5-7. ProcessTrend

폴더:

- `ProcessTrend/`

주요 파일:

- `ProcessFlowTrendView.xaml`
- `ProcessFlowTrendView.xaml.cs`
- `ProcessFlowTrendResultWindow.xaml`
- `ProcessFlowTrendResultWindow.xaml.cs`
- `ProcessTrendLargeDetailWindow.xaml`
- `ProcessTrendLargeDetailWindow.xaml.cs`
- `ProcessAxisSelectionWindow.xaml.cs`
- `ProcessPairSelectionWindow.xaml.cs`
- `ProcessTrendFileFormatWindow.xaml.cs`
- `ProcessTrendPairChoiceWindow.xaml.cs`
- `ProcessFlowTrendFileSettingsWindow.xaml.cs`

역할:

- GraphMaker 안에서 가장 복잡한 편에 속하는 공정/쌍(pair) 기준 trend 분석 모듈

흐름:

1. 공정 데이터 로딩/정규화
2. 축 선택 또는 pair 선택
3. pair별 결과 생성
4. `ProcessFlowTrendResultWindow`에서 목록 표시
5. 특정 pair 클릭 시 `ProcessTrendLargeDetailWindow` 열기

특징:

- 회귀선 / 추세선
- 상세 통계
- 정규분포 보기
- pair 단위 큰 그래프
- spec target 관련 계산

관련 타입:

- `ProcessPairPlotResult`
- `ProcessTrendComputationCandidate`

---

## 5-8. HeatMap

폴더:

- `HeatMap/`

주요 파일:

- `HeatMapView.xaml`
- `HeatMapView.xaml.cs`
- `HeatMapViewerWindow.xaml`
- `HeatMapViewerWindow.xaml.cs`
- `HeatMapFileSettingsWindow.xaml.cs`

역할:

- Heatmap 생성 전용 모듈

특징:

- 일반 OxyPlot 결과창을 재사용하기보다
- Heatmap 전용 viewer 창으로 보여주는 구조

---

## 5-9. AudioBus

폴더:

- `AudioBus/`

주요 파일:

- `AudioBusDataView.xaml`
- `AudioBusDataView.xaml.cs`

역할:

- AudioBus 관련 데이터 전용 분석 화면

이 모듈은 공용 그래프 기능이라기보다, 특정 도메인용 화면 성격이 더 강합니다.

---

## 5-10. Themes

폴더:

- `Themes/`

주요 파일:

- `ModernTheme.xaml`

역할:

- GraphMaker 전반에서 쓰는 WPF 스타일/브러시/테마 리소스 정의

---

## 6. 공용 그래프 창

### `GraphViewerWindow`

파일:

- `GraphViewerWindow.xaml`
- `GraphViewerWindow.xaml.cs`

역할:

- GraphMaker 안의 범용 그래프 출력창
- 주로 Scatter 계열이나 일반 그래프 표시용

기능:

- 라인 그래프
- 정규분포 그래프
- 평균 그래프
- CPK 그래프
- 큰 창 열기
- spec/reference/limit 라인 출력

즉:

- 공용 그래프 엔진 쪽에 가장 가까운 창입니다.

---

## 7. 구조 패턴으로 보면

GraphMaker 안의 코드는 대체로 아래 3가지 패턴으로 나뉩니다.

### 패턴 A. View -> Calculator -> ResultWindow

예시:

- `ValuePlotMultiColumnView` -> `MultiColumnGraphCalculator` -> `MultiColumnResultWindow`
- `DateNoHeaderMultiYView` -> `MultiColumnGraphCalculator` -> `MultiColumnResultWindow`
- `NoXMultiYView` -> `NoXMultiYGraphCalculator` -> `NoXMultiYResultWindow`

의미:

- 계산 로직이 View 밖의 Calculator로 빠져 있어서 비교적 구조가 명확함

### 패턴 B. View 안에서 PlotModel 직접 생성

예시:

- `ValuePlotView`
- `ScatterPlotView`

의미:

- View 코드가 다소 두꺼운 대신 바로 plot을 만듦

### 패턴 C. View -> 결과 목록창 -> 큰 상세창

예시:

- `ProcessFlowTrendView` -> `ProcessFlowTrendResultWindow` -> `ProcessTrendLargeDetailWindow`

의미:

- 복잡한 분석 결과를 2단계, 3단계 화면으로 나눠서 보여줌

---

## 8. 결과창 계층

GraphMaker 안에는 결과창 스타일이 여러 종류 있습니다.

- 범용 그래프 창
  - `GraphViewerWindow`

- 다중 컬럼 결과창
  - `MultiColumnResultWindow`

- X 없는 다중 Y 결과창
  - `NoXMultiYResultWindow`

- 공정 trend 결과 목록창
  - `ProcessFlowTrendResultWindow`

- 큰 상세창
  - `MultiColumnLargeDetailWindow`
  - `ProcessTrendLargeDetailWindow`

---

## 9. Spec / USL / LSL 처리 방식

여러 모듈에서 공통적으로 비슷한 흐름을 사용합니다.

1. 파일을 `DataTable`로 읽음
2. 컬럼별 limit 저장소를 확보함
3. `Spec`, `Upper`, `Lower`는 문자열로 저장해 둠
4. 실제 계산/그래프 그릴 때 double로 변환함
5. `Upper`와 `Lower`가 둘 다 있을 때만 NG 계산을 함
6. 평균/표준편차와 limit를 바탕으로 CPK 계산을 함

이 패턴이 많이 보이는 파일:

- `Common/MultiColumnGraphCalculator.cs`
- `NoXMultiY/NoXMultiYGraphCalculator.cs`
- `ValuePlot/ValuePlotView.xaml.cs`
- `SingleXSingleY/DailyDataTrendView.xaml.cs`
- `MultiXMultiY/DailyDataTrendExtraView.xaml.cs`

---

## 10. 폴더 맵

```text
GraphMaker/
├─ MainWindow.xaml(.cs)
├─ GraphViewerWindow.xaml(.cs)
├─ App.xaml(.cs)
├─ AudioBus/
├─ Common/
├─ DateNoHeaderMultiY/
├─ HeatMap/
├─ MultiXMultiY/
├─ NoXMultiY/
├─ ProcessTrend/
├─ ScatterPlot/
├─ SingleXSingleY/
├─ Themes/
└─ ValuePlot/
```

---

## 11. 처음 읽을 때 추천 순서

빠르게 구조를 파악하려면 아래 순서가 좋습니다.

1. `MainWindow.xaml.cs`
2. `ValuePlot/ValuePlotView.xaml.cs`
3. `Common/MultiColumnGraphCalculator.cs`
4. `Common/MultiColumnResultWindow.xaml.cs`
5. `NoXMultiY/NoXMultiYView.xaml.cs`
6. `NoXMultiY/NoXMultiYGraphCalculator.cs`
7. `ProcessTrend/ProcessFlowTrendView.xaml.cs`
8. `ProcessTrend/ProcessTrendLargeDetailWindow.xaml.cs`
9. `GraphViewerWindow.xaml.cs`
10. `Common/LargePlotWindowHelper.cs`

---

## 12. 한 줄 요약

`GraphMaker`는 하나의 그래프 프로그램이 아니라,

- Daily Sampling 계열
- Scatter 계열
- Multi-column 계열
- Process Trend 계열
- Heatmap / AudioBus 같은 특화 기능

을 한데 묶은 그래프 작업 모음입니다.

핵심 축은 보통 이렇습니다.

- 진입: `MainWindow`
- 공용 데이터: `FileInfo_DailySampling`
- 공용 계산: `Common/*Calculator`
- 공용 결과창: `GraphViewerWindow`, `MultiColumnResultWindow`
- 고급 분석: `ProcessTrend/*`

---

## 13. 다음에 해볼 수 있는 것

원하면 다음 문서도 이어서 만들 수 있습니다.

1. 화면별로 "어디를 고쳐야 하는지" 정리한 개발자용 문서
2. 각 모듈의 데이터 흐름만 따로 정리한 문서
3. 폴더/창/계산기 관계를 그림처럼 보이는 다이어그램 문서
