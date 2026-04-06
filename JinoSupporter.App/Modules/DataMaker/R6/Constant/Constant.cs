using DataMaker.R6.Grouping;
using System.Text.RegularExpressions;

namespace DataMaker.R6
{
    public class STR_MANAGER
    {
        public string ORG, NEW;

        public STR_MANAGER(string _org, string _new)
        {
            ORG = _org;
            NEW = _new;
        }

    }
    public class CONSTANT
    {
        public static string ColumnStringInput = "_INPUT";
        public static string ColumnStringNG = "_NG";
        public static string ColumnStringRATE = "_RATE";

        public static STR_MANAGER PRODUCTION_LINE = new STR_MANAGER("Production Line", "PRODUCTION_LINE");
        public static STR_MANAGER LINE_REMOVE = new STR_MANAGER("LINE_REMOVE", "LINE_REMOVE");

        public static STR_MANAGER MATERIALNAME = new STR_MANAGER("Material Name", "MATERIALNAME");
        public static STR_MANAGER ModelName_WithLine = new STR_MANAGER("Line", "Line");
        public static STR_MANAGER ModelName_MergeLR = new STR_MANAGER("LR", "LR");
        public static STR_MANAGER ModelName_MergeLRWithLine = new STR_MANAGER("LRLine", "LRLine");
        public static STR_MANAGER ModelName_Separator = new STR_MANAGER("ModelName_Separator", "ModelName_Separator");
        public static STR_MANAGER ModelName_WithLineShift = new STR_MANAGER("LineShift", "LineShift");
        public static STR_MANAGER ModelName_MergeLRBuilding = new STR_MANAGER("LR_BUILDING", "LR_BUILDING");
        public static STR_MANAGER PRODUCT_DATE = new STR_MANAGER("Product Date", "PRODUCT_DATE");

        public static STR_MANAGER PROCESSNAME = new STR_MANAGER("Process Name", "PROCESSNAME");
        public static STR_MANAGER NGNAME = new STR_MANAGER("NG Name", "NGNAME");
        public static STR_MANAGER QTYNG = new STR_MANAGER("NG Quantity(OLD)", "QTYNG");
        public static STR_MANAGER QTYINPUT = new STR_MANAGER("Input Quantity(OLD)", "QTYINPUT");
        public static STR_MANAGER WEEK = new STR_MANAGER("WEEK", "WEEK");
        public static STR_MANAGER MONTH = new STR_MANAGER("MONTH", "MONTH");
        public static STR_MANAGER PROCESSTYPE = new STR_MANAGER("PROCESSTYPE", "PROCESSTYPE");
        public static STR_MANAGER ROWID = new STR_MANAGER("ROWID", "ROWID");
        public static STR_MANAGER NGRATE = new STR_MANAGER("NGRATE", "NGRATE");
        public static STR_MANAGER No = new STR_MANAGER("No. ", "No");
        public static STR_MANAGER PROCESSCODE = new STR_MANAGER("Process Code", "PROCESSCODE");
        public static STR_MANAGER NGCODE = new STR_MANAGER("NG Code", "NGCODE");
        public static STR_MANAGER USE = new STR_MANAGER("USE (Y/N)", "USE");
        public static STR_MANAGER INPUTQUANTITY = new STR_MANAGER("Input Quantity", "INPUTQUANTITY");
        public static STR_MANAGER NGQUANTITY = new STR_MANAGER("NG Quantity", "NGQUANTITY");
        public static STR_MANAGER MATERIALCODE = new STR_MANAGER("Material Code", "MATERIALCODE");
        public static STR_MANAGER WORKODER = new STR_MANAGER("작업순서", "WORKODER");
        public static STR_MANAGER PLANT = new STR_MANAGER("Plant", "PLANT");
        public static STR_MANAGER WORKORDER = new STR_MANAGER("Workorder No.", "WORKORDER");
        public static STR_MANAGER ACTIVITYNO = new STR_MANAGER("액티비티 번호", "ACTIVITYNO");
        public static STR_MANAGER LASTREGPERSON = new STR_MANAGER("Last Reg.Person", "LASTREGPERSON");
        public static STR_MANAGER LASTREGDATE = new STR_MANAGER("Last Reg.Date", "LASTREGDATE");
        public static STR_MANAGER SHIFT = new STR_MANAGER("Shift", "Shift");
        public static STR_MANAGER BLANK = new STR_MANAGER("BLANK", "BLANK");
        public static STR_MANAGER REASON = new STR_MANAGER("REASON", "REASON");

        public static readonly Dictionary<string, string> _columnMap = new()
        {
            ["AUFNR"] = WORKORDER.ORG,
            ["WERKS"] = PLANT.ORG,
            ["VERID_TX"] = PRODUCTION_LINE.ORG,
            ["ZSHIF"] = SHIFT.ORG,
            ["WDATE"] = PRODUCT_DATE.ORG,
            ["MATNR"] = MATERIALCODE.ORG,
            ["MAKTX"] = MATERIALNAME.ORG,
            ["KTSCH"] = PROCESSCODE.ORG,
            ["KTSCH_TX"] = PROCESSNAME.ORG,
            ["ZCODE"] = NGCODE.ORG,
            ["ZCODE_TX"] = NGNAME.ORG,
            ["INQTY_O"] = QTYINPUT.ORG,
            ["USEYN"] = USE.ORG,
            ["NGQTY_O"] = QTYNG.ORG,
            ["INQTY"] = INPUTQUANTITY.ORG,
            ["NGQTY"] = NGQUANTITY.ORG,
            ["PLNFL"] = WORKODER.ORG,
            ["VORNR"] = ACTIVITYNO.ORG,
            ["ERNAM"] = LASTREGPERSON.ORG,
            ["ERDAT"] = LASTREGDATE.ORG,
        };
        public readonly static List<string> NeedOrgTableColumn = new List<string>()
        {
            PRODUCTION_LINE.NEW,
            PROCESSCODE.NEW,
            PROCESSNAME.NEW,
            NGNAME.NEW,
            QTYINPUT.NEW,
            QTYNG.NEW,
            MATERIALNAME.NEW,
            PRODUCT_DATE.NEW,

        };
        public static string Normalize(string input)
        {

            if (string.IsNullOrWhiteSpace(input))
                return "";

            string text = input.Replace("\r\n", " ")
                               .Replace("\n", " ")
                               .Replace("\r", " ");

            // 2. 스마트 따옴표 → 일반 따옴표로 변환
            text = text.Replace('‘', '\'')
                       .Replace('’', '\'')
                       .Replace('“', '"')
                       .Replace('”', '"');

            // 3. 따옴표 및 특수문자 제거 (' " ~)
            text = text.Replace("'", " ")
                       .Replace("\"", " ")
                       .Replace("~", " ");

            // 4. 대괄호 처리
            text = text.Replace("[", "")           // [ 제거
                       .Replace("]", "_");         // ] 제거하고 언더스코어 추가

            // 5. + 기호를 공백으로 변환
            text = text.Replace("+", " ");

            // 6. 연속 공백 → 한 칸
            text = Regex.Replace(text, @"\s{2,}", " ");

            // 7. 앞뒤 공백 제거
            return text.Trim();
        }

        /// <summary>
        /// GroupName을 SQL 테이블 이름으로 사용 가능하도록 변환
        /// </summary>
        public static string GetGroupTableName(clModelGroupData group)
        {
            string safeGroupName = SanitizeTableName(group.GroupName);
            return $"Group_{safeGroupName}";
        }

        /// <summary>
        /// 문자열을 SQL 테이블 이름으로 사용 가능하도록 변환
        /// </summary>
        private static string SanitizeTableName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                return "Unnamed";
            }

            // 공백, 특수문자를 언더스코어로 변환
            string safe = Regex.Replace(name, @"[^a-zA-Z0-9가-힣]", "_");

            // 연속된 언더스코어를 하나로
            safe = Regex.Replace(safe, @"_{2,}", "_");

            // 앞뒤 언더스코어 제거
            safe = safe.Trim('_');

            // 빈 문자열이면 기본값
            if (string.IsNullOrEmpty(safe))
            {
                return "Unnamed";
            }

            return safe;
        }

        public readonly static List<KeyValuePair<string, string>> ListMonthName = new List<KeyValuePair<string, string>>()
        {
            new KeyValuePair<string, string>("M1", "Jan"),
            new KeyValuePair<string, string>("M2", "Feb"),
            new KeyValuePair<string, string>("M3", "Mar"),
            new KeyValuePair<string, string>("M4", "Apr"),
            new KeyValuePair<string, string>("M5", "May"),
            new KeyValuePair<string, string>("M6", "Jun"),
            new KeyValuePair<string, string>("M7", "Jul"),
            new KeyValuePair<string, string>("M8", "Aug"),
            new KeyValuePair<string, string>("M9", "Sep"),
            new KeyValuePair<string, string>("M10", "Oct"),
            new KeyValuePair<string, string>("M11", "Nov"),
            new KeyValuePair<string, string>("M12", "Dec")
        };

        public static string GetMonthName(string key)
        {
            var item = ListMonthName.FirstOrDefault(kv => kv.Key == key);
            return item.Equals(default(KeyValuePair<string, string>)) ? key : item.Value;
        }

        public readonly static List<string> DeleteColumn = new List<string>()
        {
            No.NEW,
            USE.NEW,
            INPUTQUANTITY.NEW,
            NGQUANTITY.NEW,
            MATERIALCODE.NEW,
            WORKODER.NEW,
            PLANT.NEW,
            WORKORDER.NEW,
            ACTIVITYNO.NEW,
            LASTREGPERSON.NEW,
            LASTREGDATE.NEW,
            SHIFT.NEW
        };
        public static List<STR_MANAGER> List = new List<STR_MANAGER>()
        {
            PRODUCTION_LINE,
            LINE_REMOVE,
            ModelName_MergeLR,
            ModelName_MergeLRWithLine,
            ModelName_Separator,
            PRODUCT_DATE,
            MATERIALNAME,
            PROCESSNAME,
            NGNAME,
            QTYNG,
            QTYINPUT,
            WEEK,
            MONTH,
            PROCESSTYPE,
            ROWID,
            NGRATE,
            No,
            PROCESSCODE,
            NGCODE,
            USE,
            INPUTQUANTITY,
            NGQUANTITY,
            MATERIALCODE,
            WORKODER,
            PLANT,
            WORKORDER,
            ACTIVITYNO,
            LASTREGPERSON,
            LASTREGDATE,
            SHIFT
        };
        public static List<(string, Type)> ListAccTableColumns = new List<(string, Type)>()
        {
            new(PRODUCTION_LINE.NEW, typeof(string)),
            new(PROCESSCODE.NEW, typeof(string)),
            new(PROCESSTYPE.NEW, typeof(string)),
            new(REASON.NEW, typeof(string)),
            new(PROCESSNAME.NEW, typeof(string)),
            new(NGNAME.NEW, typeof(string)),
            new(NGCODE.NEW, typeof(string)),
            new(QTYINPUT.NEW, typeof(double)),
            new(QTYNG.NEW, typeof(double)),
            new(MATERIALNAME.NEW, typeof(string)),
            new(PRODUCT_DATE.NEW, typeof(string)),
            new(LINE_REMOVE.NEW, typeof(string)),
            new(ModelName_WithLine.NEW, typeof(string)),
            new(ModelName_WithLineShift.NEW, typeof(string)),
            new(ModelName_MergeLR.NEW, typeof(string)),
            new(ModelName_MergeLRWithLine.NEW, typeof(string)),
            new(ModelName_MergeLRBuilding.NEW, typeof(string)),
            new(MONTH.NEW, typeof(int)),
            new(WEEK.NEW, typeof(int))
        };
        public static List<STR_MANAGER> ListSTRManager = new List<STR_MANAGER>()
        {
            No,
            PRODUCTION_LINE,
            PROCESSCODE,
            PROCESSNAME,
            NGCODE,
            NGNAME,
            USE,
            QTYINPUT,
            QTYNG,
            INPUTQUANTITY,
            NGQUANTITY,
            MATERIALCODE,
            MATERIALNAME,
            PLANT,
            WORKORDER,
            SHIFT,
            PRODUCT_DATE,
            WORKODER,
            ACTIVITYNO,
            LASTREGPERSON,
            LASTREGDATE
        };

        public static int Contain_ORG(string str)
        {
            int result = -1;

            for (int i = 0; i < List.Count; i++)
            {
                if (List[i].ORG.Equals(str))
                {
                    result = i;
                }
            }

            return result;
        }

        public static int Contain_NEW(string str)
        {
            int result = -1;

            for (int i = 0; i < List.Count; i++)
            {
                if (List[i].NEW.Equals(str))
                {
                    result = i;
                }
            }

            return result;
        }
        public class OPTION
        {
            public static int nQUERY = 100;
            public static int nQtyWorst = 20;

            // Worst/WorstReason 리포트 ranking 계산 기준 기간 타입
            // "Week" = 주 기준, "Month" = 월 기준
            public static string RankingPeriodType = "Week";

            // Worst/WorstReason 리포트 ranking 계산 시 사용할 주 오프셋
            // 0 = 최신주, 1 = 저번주, 2 = 저저번주, ...
            public static int WorstRankingWeekOffset = 1;

            // Worst/WorstReason 리포트 ranking 계산 시 사용할 월 오프셋
            // 0 = 최신월, 1 = 저번달, 2 = 저저번달, ...
            public static int WorstRankingMonthOffset = 1;

            // Worst 리포트에서 상위 몇 개 NGName을 보여줄지
            // 기본값: 10개
            public static int TopNGCount = 10;

            // WorstProcess 리포트에서 상위 몇 개 ProcessName을 보여줄지
            // 기본값: 10개
            public static int TopProcessCount = 10;

            // WorstReason 리포트에서 각 Reason별로 보여줄 불량 조합 개수
            // 기본값: 5개
            public static int TopDefectsPerReason = 5;
        }

        public static class OPTION_MODEL_SELECT
        {
            public static string MODEL_NAME = CONSTANT.MATERIALNAME.NEW;
            public static string MODEL_NAME_WITH_LINE = CONSTANT.ModelName_WithLine.NEW;
            public static string MODEL_NAME_WITH_LINE_SHIFT = CONSTANT.ModelName_WithLineShift.NEW;
            public static string MODEL_NAME_MERGE_LR = CONSTANT.ModelName_MergeLR.NEW;
            public static string MODEL_NAME_MERGE_LR_WITH_LINE = CONSTANT.ModelName_MergeLRWithLine.NEW;
            public static string MODEL_NAME_BUILDING = CONSTANT.ModelName_MergeLRBuilding.NEW;

        }
        public static class OPTION_ROUTING
        {
            public static string ROUTING_MATERIALNAME = CONSTANT.MATERIALNAME.NEW;
            public static string ROUTING_PROCESSCODE = CONSTANT.PROCESSCODE.NEW;
            public static string ROUTING_PROCESSNAME = CONSTANT.PROCESSNAME.NEW;
            public static string ROUTING_PROCESSTYPE = CONSTANT.PROCESSTYPE.NEW;
        }
        public static class OPTION_TABLE_NAME
        {
            public static string ORG = "OrginalTable";
            public static string ACC = "AccessTable";
            public static string TEMP = "TempTable";
            public static string ROUTING = "RoutingTable";
            public static string Selected = "SelectedTable";
            public static string REASON = "Reason";
            public static string PROC = "ProcTable";
            public static string MISSING_ROUTING = "MissingRouting";
            public static string MISSING_REASON = "MissingReason";

        }


    }
}
