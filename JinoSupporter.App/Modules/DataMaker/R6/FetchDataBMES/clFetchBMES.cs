using DataMaker.Logger;
using System.Data;
using System.Net;
using System.Net.Http;
using System.Text.Json;

namespace DataMaker.R6.FetchDataBMES
{
    public class clFetchBMES
    {
        private readonly HttpClient _client;
        private readonly CookieContainer _cookieContainer = new();
        private const string BaseUrl = "http://bmes.bujeon.com";
        private string LoginID { get; set; }
        private string Password { get; set; }
        private string DateStart { get; set; }
        private string DateEnd { get; set; }

        public clFetchBMES(string loginID, string password, string StartDate, string EndData)
        {
            var handler = new HttpClientHandler()
            {
                UseCookies = true,
                CookieContainer = _cookieContainer
            };
            LoginID = loginID;
            Password = password;
            DateStart = StartDate;
            DateEnd = EndData;

            _client = new HttpClient(handler);
            _client.Timeout = TimeSpan.FromSeconds(300); // 5분 타임아웃
            _client.DefaultRequestHeaders.Add("X-Requested-With", "XMLHttpRequest");
        }

        public async Task<DataTable> RunAsync()
        {
            var token = await GetVerificationTokenAsync();
            if (string.IsNullOrEmpty(token))
            {
                clLogger.Log("Failed to Extract Token");
                return null;
            }

            var loginSuccess = await LoginAsync(token);
            if (!loginSuccess)
            {
                clLogger.Log("Failed to Login");
                return null;
            }

            // Fetch RAW data from both WERKS (3200 and 3220) - NO column transformation yet
            clLogger.Log("Fetching raw data from WERKS=3200...");
            var rawData3200 = await FetchRawDataFromWERKSAsync("3200");

            clLogger.Log("Fetching raw data from WERKS=3220...");
            var rawData3220 = await FetchRawDataFromWERKSAsync("3220");

            // Merge raw data BEFORE column transformation
            DataTable mergedRawData = null;

            if (rawData3200 == null && rawData3220 == null)
            {
                clLogger.Log("Failed to fetch data from both WERKS");
                return null;
            }

            if (rawData3200 == null)
            {
                clLogger.Log("Only WERKS=3220 raw data available");
                mergedRawData = rawData3220;
            }
            else if (rawData3220 == null)
            {
                clLogger.Log("Only WERKS=3200 raw data available");
                mergedRawData = rawData3200;
            }
            else
            {
                // Both exist - merge raw data (same column structure, no duplicates)
                rawData3200.Merge(rawData3220);
                clLogger.Log($"Merged raw data from both WERKS: Total {rawData3200.Rows.Count} rows");
                mergedRawData = rawData3200;
            }

            // NOW perform column transformation ONCE on merged data
            var transformedData = TransformDataTable(mergedRawData);
            return transformedData;
        }

        private async Task<string> GetVerificationTokenAsync()
        {
            var html = await _client.GetStringAsync(BaseUrl);
            var doc = new HtmlAgilityPack.HtmlDocument();
            doc.LoadHtml(html);
            return doc.DocumentNode.SelectSingleNode("//input[@name='__RequestVerificationToken']")?.GetAttributeValue("value", "");
        }

        private async Task<bool> LoginAsync(string token)
        {
            var loginUrl = BaseUrl + "/MES000000/LoginCheck";

            var content = new FormUrlEncodedContent(new Dictionary<string, string>
        {
            {"UserInfo.USRID", LoginID},
            {"UserInfo.PWNO", Password},
            {"UserInfo.LANG", "EN"},
            {"UserInfo.FACCO", "GN"},
            {"UserInfo.STYPE", "P"},
            {"UserInfo.VTYPE", "P"},
            {"__RequestVerificationToken", token}
        });

            var response = await _client.PostAsync(loginUrl, content);
            var body = await response.Content.ReadAsStringAsync();

            clLogger.Log("Login Response: " + response.StatusCode);
            clLogger.Log("Response Content: " + body);

            return body.Contains("\"Result\":\"M\"");
        }

        /// <summary>
        /// Fetch raw data from WERKS without column transformation
        /// </summary>
        private async Task<DataTable> FetchRawDataFromWERKSAsync(string werks)
        {
            var url = BaseUrl + $"/MES020210/SearchList?perPage=&Condition.WERKS={werks}&Condition.SDATE={DateStart}&Condition.EDATE={DateEnd}&Condition.INPYN=N&Condition.USEYN=Y&page=1";
            var response = await _client.GetAsync(url);
            clLogger.Log($"Response Code SearchList (WERKS={werks}): " + response.StatusCode);

            if (!response.IsSuccessStatusCode)
            {
                clLogger.Log("Failed to Request: " + await response.Content.ReadAsStringAsync());
                return null;
            }

            var json = await response.Content.ReadAsStringAsync();
            try
            {
                using var doc = JsonDocument.Parse(json);
                var rows = doc.RootElement.GetProperty("data").GetProperty("contents");

                var table = new DataTable();
                foreach (var prop in rows[0].EnumerateObject())
                    table.Columns.Add(prop.Name);

                clLogger.Log($"Fetched {rows.GetArrayLength()} raw rows from BMES (WERKS={werks})");

                foreach (var item in rows.EnumerateArray())
                {
                    var row = table.NewRow();
                    int colIndex = 0;

                    foreach (var prop in item.EnumerateObject())
                    {
                        var value = prop.Value;

                        // null 체크 및 적절한 타입으로 변환
                        if (value.ValueKind == JsonValueKind.Null)
                        {
                            row[colIndex] = DBNull.Value;
                        }
                        else if (value.ValueKind == JsonValueKind.Number)
                        {
                            // 숫자는 숫자 타입 그대로 저장
                            if (value.TryGetDouble(out double numValue))
                                row[colIndex] = numValue;
                            else
                                row[colIndex] = value.ToString();
                        }
                        else
                        {
                            row[colIndex] = value.ToString();
                        }

                        colIndex++;
                    }

                    table.Rows.Add(row);
                }

                // Return raw table WITHOUT column transformation
                return table;
            }
            catch (Exception ex)
            {
                clLogger.Log("Failed to Parse JSON: " + ex.Message);
                return null;
            }
        }

        /// <summary>
        /// Transform column names and reorder columns (perform ONCE on merged data)
        /// </summary>
        private DataTable TransformDataTable(DataTable table)
        {
            if (table == null)
                return null;

            // Remove VERID column if exists
            if (table.Columns.Contains("VERID"))
                table.Columns.Remove("VERID");

            // First pass: rename columns using _columnMap
            foreach (DataColumn col in table.Columns)
                if (CONSTANT._columnMap.ContainsKey(col.ColumnName))
                    col.ColumnName = CONSTANT._columnMap[col.ColumnName];

            // Create new table with correct column order
            DataTable ColumnNewOrder = new DataTable(table.TableName);

            foreach (STR_MANAGER m in CONSTANT.ListSTRManager)
            {
                ColumnNewOrder.Columns.Add(m.ORG);
            }

            // Merge data into ordered table
            ColumnNewOrder.Merge(table);

            // Second pass: rename columns using ListSTRManager
            foreach (DataColumn col in ColumnNewOrder.Columns)
            {
                string colName = col.ColumnName;

                // 매칭되는 STRManager 객체 찾기
                var match = CONSTANT.ListSTRManager.FirstOrDefault(x => x.ORG == colName);

                // 있으면 이름 변경
                if (match != null)
                {
                    col.ColumnName = match.NEW;
                }
            }

            // QTYNG 값 검증 로깅
            if (ColumnNewOrder.Columns.Contains(CONSTANT.QTYNG.NEW))
                {
                    int zeroCount = 0;
                    int nullCount = 0;

                    foreach (DataRow row in ColumnNewOrder.Rows)
                    {
                        var qtyNgValue = row[CONSTANT.QTYNG.NEW];

                        if (qtyNgValue == DBNull.Value || qtyNgValue == null)
                        {
                            nullCount++;
                            if (nullCount <= 3) // 처음 3개만 로그
                            {
                                clLogger.Log($"Warning: QTYNG is NULL - ProcessName: {row[CONSTANT.PROCESSNAME.NEW]}, " +
                                           $"NgName: {row[CONSTANT.NGNAME.NEW]}, Date: {row[CONSTANT.PRODUCT_DATE.NEW]}");
                            }
                        }
                        else if (qtyNgValue.ToString() == "0" || Convert.ToDouble(qtyNgValue) == 0)
                        {
                            zeroCount++;
                            if (zeroCount <= 3) // 처음 3개만 로그
                            {
                                clLogger.Log($"Warning: QTYNG is 0 - ProcessName: {row[CONSTANT.PROCESSNAME.NEW]}, " +
                                           $"NgName: {row[CONSTANT.NGNAME.NEW]}, Date: {row[CONSTANT.PRODUCT_DATE.NEW]}");
                            }
                        }
                    }

                if (zeroCount > 0 || nullCount > 0)
                {
                    clLogger.Log($"QTYNG validation: {zeroCount} rows with 0, {nullCount} rows with NULL (out of {ColumnNewOrder.Rows.Count} total rows)");
                }
            }

            return ColumnNewOrder;
        }
    }
}
