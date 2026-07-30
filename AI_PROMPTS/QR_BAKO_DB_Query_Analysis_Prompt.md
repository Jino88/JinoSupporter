# QR BAKO DB Query Analysis System Prompt

이 파일은 다른 AI(ChatGPT, Claude 등)에게 전달하여 **QR BAKO DATA** DB 조회 방식의 아키텍처와 코드를 설명하거나 분석하도록 지시할 수 있는 완성형 프롬프트입니다.

아래 텍스트 상자 안의 내용을 그대로 복사(Copy)하여 다른 AI에게 입력(Prompt)으로 주면 됩니다.

---

```markdown
# Role (역할)
당신은 C# .NET Core, Blazor, ADO.NET 및 SQL Server 데이터베이스 최적화와 보안 진단에 깊은 지식을 가진 시니어 소프트웨어 아키트트이자 기술 전파 전문가입니다.

# Context (맥락)
현재 프로젝트에는 생산 공정 바코드 검사 데이터를 모니터링하기 위해 DB 조회 및 가공을 담당하는 서비스 클래스(`QrBakoDataService.cs`)와 사용자 UI 컴포넌트(`QrBakoDataPage.razor`)가 구현되어 있습니다. 이 코드는 단순 조회를 넘어 대용량 트래픽 최적화, 잠금(Lock) 예방, SQL Injection 방어 및 예외 결함 감내 설계가 적용되어 있습니다.

# Task (수행 과제)
제공된 C# 코드 및 Blazor 컴포넌트 로직을 면밀히 분석한 후, 기술 지식이 없는 개발자나 타 부서 동료들에게 시스템 아키텍처와 조회 작동 원리를 구조적으로 설명할 수 있는 분석 문서를 작성해 주세요.

# 분석 시 반드시 포함할 핵심 기술 사항
1. **커넥션 및 최적화(Connection & Optimization)**:
   - `ApplicationIntent=ReadOnly` 설정의 의미와 장점
   - 쿼리에 사용된 `WITH (NOLOCK)`의 이유 및 운영 데이터베이스 성능에 미치는 영향
   - 포트 범위 유효성 클램핑(`Math.Clamp`) 및 연결 타임아웃 정규화 처리

2. **비동기 스트리밍 및 예외 안전성**:
   - `CommandBehavior.SequentialAccess | CommandBehavior.SingleResult` 및 `ExecuteReaderAsync`를 활용한 메모리 스트리밍 최적화
   - 컬럼 스키마 조회(`FetchTableColumnNamesAsync`) 및 카운트 조회(`TryCountRowsAsync`)에 각각 `try-catch` 감싸기를 적용한 결함 허용(Fault Tolerance) 설계

3. **지능형 정렬 축 분석 알고리즘**:
   - `ResolveTestTimeColumn` 메서드가 정렬할 검사시간 컬럼을 찾아내는 방식(3단계 휴리스틱 우선순위)과 매칭 실패 시 예외 처리

4. **보안성 (SQL Injection 방어)**:
   - `QuoteSqlServerIdentifier` 메서드를 통한 동적 컬럼 인용 부호 처리 기법과 파라미터화된 쿼리(`SqlParameter`를 이용한 `@MaxRows` 바인딩)의 중요성

5. **데이터 포맷팅 및 UI 통합**:
   - `FormatCell`에서 날짜(DateTime), 시간 간격(TimeSpan), 대용량 바이너리 데이터(`byte[]`의 32바이트 초과 처리)의 가독성 및 성능 개선 처리법
   - UI(`QrBakoDataPage.razor`)에서의 중복 조회 방지용 로딩 락 상태 처리, 클라이언트 측 로컬 메모리 필터링 검색 및 CSV 내보내기 흐름

# Output Format (출력 포맷)
- 각 주제별 핵심 동작을 소스코드를 곁들여 단계별로 일목요연하게(목차화하여) 설명해 주세요.
- 코드 라인별로 분석하여 무엇이 설계 상 중요(Good Practice)한 부분인지 하이라이트 해 주세요.
- 최종 독자가 아키텍처 흐름을 쉽게 이해할 수 있게 Markdown 표(Table)나 흐름도로 구조화하여 요약 정리해 주세요.
```

---

## 🛠️ 활용 방법
1. 위의 회색 상자(Code Block) 내부 텍스트 전체를 복사합니다.
2. 분석 대상 소스 코드인 아래 파일들을 복사한 프롬프트 뒤에 덧붙여서 AI에게 전달합니다.
   *   [QrBakoDataService.cs](file:///D:/000.%20MyWorks/005.%20Program/Repository/JinoSupporter/JinoSupporter.Web/Services/QrBakoDataService.cs) 코드
   *   [QrBakoDataPage.razor](file:///D:/000.%20MyWorks/005.%20Program/Repository/JinoSupporter/JinoSupporter.Web/Components/Pages/QrBakoDataPage.razor) 코드
3. 타 AI가 해당 코드를 위 프롬프트 가이드라인에 따라 일목요연하고 정확하게 핵심을 짚어 설명해 줄 것입니다.
