using DataMaker.Logger;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace DataMaker.R6.FetchDataBMES
{
    /// <summary>
    /// MES020240/SearchList URL에서 Routing 데이터를 가져와 TXT 파일로 저장하는 클래스
    /// </summary>
    public class clFetchRoutingData
    {
        private readonly HttpClient _client;
        private readonly CookieContainer _cookieContainer = new();
        private const string BaseUrl = "http://bmes.bujeon.com";
        private string LoginID { get; set; }
        private string Password { get; set; }

        public clFetchRoutingData(string loginID, string password)
        {
            var handler = new HttpClientHandler()
            {
                UseCookies = true,
                CookieContainer = _cookieContainer
            };
            LoginID = loginID;
            Password = password;

            _client = new HttpClient(handler);
            _client.Timeout = TimeSpan.FromSeconds(300); // 5분 타임아웃
            _client.DefaultRequestHeaders.Add("X-Requested-With", "XMLHttpRequest");
        }

        /// <summary>
        /// 데이터를 가져와서 TXT 파일로 저장
        /// </summary>
        public async Task<bool> FetchAndSaveToTxtAsync(string outputFilePath)
        {
            try
            {
                // 1. 인증 토큰 가져오기
                var token = await GetVerificationTokenAsync();
                if (string.IsNullOrEmpty(token))
                {
                    clLogger.Log("Failed to Extract Token");
                    return false;
                }

                // 2. 로그인
                var loginSuccess = await LoginAsync(token);
                if (!loginSuccess)
                {
                    clLogger.Log("Failed to Login");
                    return false;
                }

                // 3. 데이터 가져오기
                var dataTable = await FetchDataAsync();
                if (dataTable == null || dataTable.Rows.Count == 0)
                {
                    clLogger.Log("No data fetched from server");
                    return false;
                }

                // 4. TXT 파일로 저장
                SaveToTxtFile(dataTable, outputFilePath);
                clLogger.Log($"Successfully saved {dataTable.Rows.Count} rows to {outputFilePath}");
                return true;
            }
            catch (Exception ex)
            {
                clLogger.LogException(ex, "Error in FetchAndSaveToTxtAsync");
                return false;
            }
        }

        /// <summary>
        /// 인증 토큰 가져오기
        /// </summary>
        private async Task<string> GetVerificationTokenAsync()
        {
            var html = await _client.GetStringAsync(BaseUrl);
            var doc = new HtmlAgilityPack.HtmlDocument();
            doc.LoadHtml(html);
            return doc.DocumentNode.SelectSingleNode("//input[@name='__RequestVerificationToken']")?.GetAttributeValue("value", "");
        }

        /// <summary>
        /// 로그인
        /// </summary>
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

            return body.Contains("\"Result\":\"M\"");
        }

        /// <summary>
        /// MES020240 URL에서 데이터 가져오기
        /// </summary>
        private async Task<DataTable> FetchDataAsync()
        {
            // MES020240/SearchList URL 사용 (날짜 범위를 넓게 설정)
            var url = BaseUrl + "/MES020240/SearchList?perPage=&Condition.WERKS=3200&Condition.WPHNO=&Condition.MATNR=&Condition.DATUV_S=1900-01-01&Condition.DATUV_E=2050-12-31&Condition.FDATE=1900-01-01&Condition.TDATE=2050-12-31&Condition.PTYPE=&Condition.BTYPE=&Condition.TTYPE=&page=1";

            clLogger.Log("Fetching data from: " + url);
            var response = await _client.GetAsync(url);
            clLogger.Log("Response Code: " + response.StatusCode);

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

                if (rows.GetArrayLength() == 0)
                {
                    clLogger.Log("No data returned from server");
                    return null;
                }

                // DataTable 생성
                var table = new DataTable();

                // 첫 번째 행에서 컬럼 추출
                foreach (var prop in rows[0].EnumerateObject())
                {
                    table.Columns.Add(prop.Name);
                }

                // 데이터 추가
                foreach (var item in rows.EnumerateArray())
                {
                    var values = item.EnumerateObject().Select(p => (object?)p.Value.ToString()).ToArray();
                    table.Rows.Add(values);
                }

                clLogger.Log($"Fetched {table.Rows.Count} rows with {table.Columns.Count} columns");
                return table;
            }
            catch (Exception ex)
            {
                clLogger.LogException(ex, "Failed to Parse JSON");
                return null;
            }
        }

        /// <summary>
        /// DataTable을 탭 구분자로 TXT 파일에 저장
        /// </summary>
        private void SaveToTxtFile(DataTable dataTable, string outputFilePath)
        {
            using (var writer = new StreamWriter(outputFilePath, false, Encoding.UTF8))
            {
                // 헤더 작성 (컬럼명)
                var columnNames = dataTable.Columns.Cast<DataColumn>().Select(col => col.ColumnName);
                writer.WriteLine(string.Join("\t", columnNames));

                // 데이터 작성
                foreach (DataRow row in dataTable.Rows)
                {
                    var values = row.ItemArray.Select(val => val?.ToString() ?? "");
                    writer.WriteLine(string.Join("\t", values));
                }
            }

            clLogger.Log($"Data saved to: {outputFilePath}");
        }

        /// <summary>
        /// 특정 컬럼만 필터링해서 TXT 파일로 저장
        /// </summary>
        public async Task<bool> FetchAndSaveToTxtAsync(string outputFilePath, string[] columnsToInclude)
        {
            try
            {
                // 1. 인증 토큰 가져오기
                var token = await GetVerificationTokenAsync();
                if (string.IsNullOrEmpty(token))
                {
                    clLogger.Log("Failed to Extract Token");
                    return false;
                }

                // 2. 로그인
                var loginSuccess = await LoginAsync(token);
                if (!loginSuccess)
                {
                    clLogger.Log("Failed to Login");
                    return false;
                }

                // 3. 데이터 가져오기
                var dataTable = await FetchDataAsync();
                if (dataTable == null || dataTable.Rows.Count == 0)
                {
                    clLogger.Log("No data fetched from server");
                    return false;
                }

                // 4. 컬럼 필터링
                var filteredTable = FilterColumns(dataTable, columnsToInclude);

                var map = new Dictionary<string, string>
                {
                    ["Main Assy"] = "MAIN",
                    ["Sub Assy"] = "SUB",
                    ["Function"] = "FUNCTION",
                    ["Final Visual"] = "VISUAL"
                };

                foreach (DataRow v in filteredTable.Rows)
                {
                    string key = (v.Field<string>(3) ?? "").Trim();
                    v.SetField(3, map.TryGetValue(key, out var value) ? value : "");
                }

                // 5. TXT 파일로 저장
                SaveToTxtFile(filteredTable, outputFilePath);
                clLogger.Log($"Successfully saved {filteredTable.Rows.Count} rows with {filteredTable.Columns.Count} columns to {outputFilePath}");
                return true;
            }
            catch (Exception ex)
            {
                clLogger.LogException(ex, "Error in FetchAndSaveToTxtAsync with column filter");
                return false;
            }
        }

        /// <summary>
        /// 특정 컬럼만 포함하는 새 DataTable 생성
        /// </summary>
        private DataTable FilterColumns(DataTable source, string[] columnsToInclude)
        {
            var filtered = new DataTable();

            // 요청된 컬럼만 추가
            foreach (var colName in columnsToInclude)
            {
                if (source.Columns.Contains(colName))
                {
                    filtered.Columns.Add(colName, source.Columns[colName].DataType);
                }
                else
                {
                    clLogger.LogWarning($"Column '{colName}' not found in source data. Skipping.");
                }
            }

            // 데이터 복사
            foreach (DataRow sourceRow in source.Rows)
            {
                var newRow = filtered.NewRow();
                foreach (DataColumn col in filtered.Columns)
                {
                    newRow[col.ColumnName] = sourceRow[col.ColumnName];
                }
                filtered.Rows.Add(newRow);
            }

            return filtered;
        }
    }
}
