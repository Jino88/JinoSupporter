using System.Collections.Generic;
using System.Data;

namespace GraphMaker;

/// <summary>
/// 모든 GraphMaker FileInfo 클래스의 공통 기반.
/// 파일 경로, 구분자, 헤더 행 번호, 로드된 데이터 등 공통 속성을 제공합니다.
/// </summary>
public abstract class GraphFileInfoBase
{
    public string Name { get; set; } = string.Empty;
    public string FilePath { get; set; } = string.Empty;
    public DataTable? FullData { get; set; }
    public List<string>? HeaderRow { get; set; }
    public string Delimiter { get; set; } = "\t";
    public int HeaderRowNumber { get; set; } = 1;
}
