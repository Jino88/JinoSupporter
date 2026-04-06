using System;
using System.Collections.Generic;

namespace DataMaker.R6
{
    /// <summary>
    /// 모델 정보 클래스
    /// </summary>
    public class clModelInfo
    {
        public string ModelName { get; set; }
        public string Description { get; set; }

        public clModelInfo()
        {
        }

        public clModelInfo(string modelName)
        {
            ModelName = modelName;
        }

        public clModelInfo(string modelName, string description)
        {
            ModelName = modelName;
            Description = description;
        }

        public override string ToString()
        {
            return ModelName ?? string.Empty;
        }
    }
}

namespace DataMaker.R6
{
    /// <summary>
    /// 전역 옵션 설정 클래스
    /// </summary>
    public static class clOption
    {
        /// <summary>
        /// AccessTable에서 모델 추출 시 사용할 표준 컬럼들
        /// </summary>
        public static readonly string[] StandardColumnsModelExtract = new[]
        {
            CONSTANT.MATERIALNAME.NEW,
            CONSTANT.PROCESSCODE.NEW,
            CONSTANT.PROCESSNAME.NEW
        };

        /// <summary>
        /// AccessTable 컬럼 정의 (컬럼명, 타입)
        /// </summary>
        public static List<(string ColumnName, Type ColumnType)> GetAccessTableColumns()
        {
            return new List<(string, Type)>
            {
                (CONSTANT.PRODUCTION_LINE.NEW, typeof(string)),
                (CONSTANT.PROCESSCODE.NEW, typeof(string)),
                (CONSTANT.PROCESSTYPE.NEW, typeof(string)),
                (CONSTANT.REASON.NEW, typeof(string)),
                (CONSTANT.PROCESSNAME.NEW, typeof(string)),
                (CONSTANT.NGNAME.NEW, typeof(string)),
                (CONSTANT.NGCODE.NEW, typeof(string)),  // ← NGCODE 추가!
                (CONSTANT.QTYINPUT.NEW, typeof(double)),
                (CONSTANT.QTYNG.NEW, typeof(double)),
                (CONSTANT.MATERIALNAME.NEW, typeof(string)),
                (CONSTANT.PRODUCT_DATE.NEW, typeof(string)),
                (CONSTANT.SHIFT.NEW, typeof(string)),
               // (CONSTANT.LINE_REMOVE.NEW, typeof(string)),
                (CONSTANT.ModelName_WithLineShift.NEW, typeof(string)),
                //(CONSTANT.ModelName_MergeLR.NEW, typeof(string)),
                //(CONSTANT.ModelName_MergeLRWithLine.NEW, typeof(string)),
                //(CONSTANT.ModelName_MergeLRBuilding.NEW, typeof(string)),
                (CONSTANT.MONTH.NEW, typeof(int)),
                (CONSTANT.WEEK.NEW, typeof(int))
            };
        }
    
        /// <summary>
        /// ProcessTypeTable (CSV) 컬럼 정의
        /// </summary>
        public static List<string> GetProcessTypeTableColumns()
        {
            return new List<string>
            {
                "모델명",
                "ProcessCode",
                "ProcessName",
                "ProcessType"
            };
        }

        /// <summary>
        /// ReasonTable (CSV) 컬럼 정의
        /// </summary>
        public static List<string> GetReasonTableColumns()
        {
            return new List<string>
            {
                "processName",
                "NgName",
                "Reason"
            };
        }
    }
}
