# QR BAKO HTML Dashboard Generator AI Prompt

이 파일은 타 AI(Gemini, Claude, ChatGPT 등)에 전달하여 현재 [QrBakoDataPage.razor](file:///D:/000.%20MyWorks/005.%20Program/Repository/JinoSupporter/JinoSupporter.Web/Components/Pages/QrBakoDataPage.razor)와 동일한 기능 및 레이아웃을 가진 **단독 실행형 HTML5 대시보드 페이지**를 생성하도록 지시하는 프롬프트입니다.

---

```markdown
# Role (역할)
당신은 UI/UX 디자인 감각이 뛰어난 시니어 프론트엔드 개발자이자 퍼블리셔입니다. 현대적이고 세련된 디자인 시스템(Inter 폰트, Glassmorphism, 부드러운 애니메이션, HSL 테일러드 컬러칩)을 활용해 단 하나의 파일로 작동하는 프리미엄 HTML5 대시보드를 생성합니다.

# Context (맥락)
사용자는 ASP.NET Core Blazor 서버로 구동되던 "QR BAKO DATA" 모니터링 페이지를 웹 브라우저에서 단독으로 실행하거나 다른 시스템에 빠르게 이식할 수 있도록 **HTML, CSS, JavaScript가 하나로 통합된 단일 파일(.html) 대시보드**로 재구현하고자 합니다.

# Goal (목표)
다음 요구사항을 충족하고 시각적으로 압도적인 고품질 단일 HTML 대시보드 코드를 작성해 주세요. 

---

## 🎨 1. 디자인 시스템 & UX 요구사항 (CSS)
- **최신 폰트**: Google Fonts의 `Outfit` 또는 `Inter`를 웹 폰트로 로드하여 깔끔하고 고급스러운 타이포그래피 구현.
- **색상 팔레트**: 어둡고 차분한 슬레이트 블루 그레이 테마(#0F172A, #1E293B) 또는 세련된 HSL 기반의 라이트/다크 하모니 적용. 평범한 단색 대신 부드러운 그라데이션 포인트 배색 사용.
- **그리드 테이블 스타일링**:
  - `#` 열과 헤더는 Sticky 속성을 주어 스크롤 시에도 고정.
  - 마우스 오버 시 행(Row) 하이라이팅 효과 및 셀 텍스트가 너무 길면 말줄임표(...) 및 툴팁(`title` 속성) 표시.
  - 데이터 영역은 브라우저 높이에 맞춰 반응형 높이 제한(`max-height`) 및 내부 스크롤 처리.
- **컴포넌트 카드**: 그림자 효과(`shadow-sm`), 테두리 라운딩(`border-radius`), 깔끔한 여백 배치를 통해 정보 영역 구획화.

---

## ⚙️ 2. 핵심 기능 요구사항 (JavaScript)
1. **연결 정보 상태 바 (Connection Header)**:
   - 데이터베이스 서버 정보 (`tcdb.server.ip,1430 / TCDB / Read-Only`)를 보여주는 상태 창 제공.
2. **조회 제한수 필터 (Max Rows)**:
   - 숫자 입력창(기본값 1,000, 최대 20,000)을 제공하여 데이터 생성/호출 건수를 제한.
3. **새로고침 시뮬레이션 (Refresh Button)**:
   - 새로고침 클릭 시 버튼이 비활성화되며 로딩 스피너 애니메이션 표시.
   - 0.5초~1초의 비동기 딜레이(Mocking) 후 무작위 샘플 데이터를 렌더링.
   - 조회된 총 행수(Total Rows), 조회 경과 시간(Fetched At), 정렬 방식(Sorted by Test Time DESC) 등을 카드 헤더 배지에 실시간 갱신.
4. **실시간 클라이언트 검색 (Real-time Filter)**:
   - 검색창 입력 시, 전체 컬럼 값을 대상으로 대소문자 구분 없이 실시간 고속 필터링하여 검색어와 매칭되는 행만 화면에 표시.
5. **CSV 엑셀 내보내기 (Export CSV)**:
   - 'Export CSV' 클릭 시 현재 필터링되어 화면에 표시되고 있는 데이터 테이블 내용을 CSV 규격 문자열로 가공.
   - UTF-8 BOM(\uFEFF)을 포함시켜 엑셀에서 열어도 한글이 깨지지 않도록 한 뒤 파일명 `QR_BAKO_DATA_BKTD_YYYYMMDD_HHMMSS.csv`로 다운로드 처리.

---

## 📊 3. 샘플 데이터셋 스키마 (Mock Data)
조회 시 기본으로 적재될 모의 데이터를 코드 내부에 배열 구조로 포함시켜 주세요.
- 컬럼 구성: `#`, `Test Time`, `Barcode`, `Device ID`, `Result (OK/NG)`, `NG Reason`, `Operator` 등 최소 6개 이상의 컬럼.
- 대용량 바이너리나 긴 텍스트 셀은 32바이트 이상일 때 자동으로 `0x4A8C... (1024 bytes)`처럼 말줄임 및 생략 변환(C# FormatCell 로직 모사)되도록 구현.

# Output Format
외부 라이브러리 종속성(jQuery 등) 없이 **순수 HTML5, Vanilla CSS, Vanilla JS**만 사용한 단일 `<DOCTYPE html>` 형식의 완성된 코드 블록만 출력해 주세요. 주석을 친절하게 달아 유지보수가 용이하게 해주세요.
```

---

## 🛠️ 활용 방법
1. 위의 회색 상자(Code Block) 내부 프롬프트 전체를 복사합니다.
2. ChatGPT 또는 Claude와 같은 AI에게 질문 창에 붙여넣습니다.
3. 생성된 코드를 복사하여 `.html` 파일(예: `qr_bako_dashboard.html`)로 저장한 뒤 더블클릭하면 완벽하게 동작하는 대시보드를 바로 사용할 수 있습니다.
