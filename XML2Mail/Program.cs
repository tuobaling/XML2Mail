using ClosedXML.Excel;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

class Program
{
    static void Main()
    {
        string basePath = AppDomain.CurrentDomain.BaseDirectory;
        string projectRoot = Directory.GetParent(basePath)       // bin\Debug\net6.0\
                                       ?.Parent                   // XML2Mail\
                                       ?.Parent                   // XML2Mail\
                                       ?.Parent?.FullName         // ✅ ← 這才是專案根目錄
                                       ?? throw new Exception("無法解析專案根目錄");

        string inputPath = Path.Combine(projectRoot, "202505-碩益科技股份有限公司.xlsx");
        string outputPath = Path.Combine(projectRoot, $"{DateTime.Now.ToString("yyyyMMdd")} 處理完成.xlsx");

        using var workbook = new XLWorkbook(inputPath);
        var ws = workbook.Worksheet("用量明細");
        var range = ws.RangeUsed();

        // 取得標題列與資料列
        var headerRow = range.FirstRowUsed();
        var headers = headerRow.Cells().Select(c => c.GetString()).ToList();
        var dataRows = range.RowsUsed().Skip(1); // 跳過表頭

        // 建立輸出檔案
        var newWb = new XLWorkbook();
        var summarySheet = newWb.AddWorksheet("總表");
        summarySheet.Cell(1, 1).Value = "客戶名稱";
        summarySheet.Cell(1, 2).Value = "訂閱名稱";
        summarySheet.Cell(1, 3).Value = "金額欄位";
        summarySheet.Cell(1, 4).Value = "總計";
        int summaryRow = 2;

        // 分組
        var groupedByCustomer = dataRows.GroupBy(r =>
            r.Cell(headers.IndexOf("客戶名稱") + 1).GetString()
        );

        foreach (var customerGroup in groupedByCustomer)
        {
            string customer = customerGroup.Key;

            IEnumerable<IGrouping<string, IXLRangeRow>> subGroups;

            if (customer == "碩益科技股份有限公司")
                subGroups = customerGroup.GroupBy(r => r.Cell(headers.IndexOf("訂閱名稱") + 1).GetString());
            else
                subGroups = new List<IGrouping<string, IXLRangeRow>> { new FakeGroup(customerGroup) };

            foreach (var subGroup in subGroups)
            {
                string subName = subGroup.First().Cell(headers.IndexOf("訂閱名稱") + 1).GetString();
                if (string.IsNullOrWhiteSpace(subName))
                    subName = "全部";
                string sheetName = (customer + "_" + subName).Length > 31
                    ? (customer + "_" + subName).Substring(0, 31)
                    : (customer + "_" + subName);

                var wsNew = newWb.Worksheets.Add(sheetName);
                int currentRow = 1;

                // 表頭
                for (int c = 0; c < headers.Count; c++)
                {
                    wsNew.Cell(currentRow, c + 1).Value = headers[c];
                    wsNew.Cell(currentRow, c + 1).Style.Font.Bold = true;
                    wsNew.Cell(currentRow, c + 1).Style.Font.FontName = "Calibri";
                }
                currentRow++;

                double total = 0;
                string totalColName = customer == "碩益科技股份有限公司" ? "經銷價" : "建議售價";

                foreach (var row in subGroup)
                {
                    for (int c = 0; c < headers.Count; c++)
                    {
                        string colName = headers[c];
                        string value = row.Cell(c + 1).GetString();

                        if (customer == "碩益科技股份有限公司" && colName == "建議售價")
                            value = "";
                        if (customer != "碩益科技股份有限公司" && colName == "經銷價")
                            value = "";

                        wsNew.Cell(currentRow, c + 1).Value = value;
                        wsNew.Cell(currentRow, c + 1).Style.Font.FontName = "Calibri";
                    }

                    var cellVal = row.Cell(headers.IndexOf(totalColName) + 1).GetString();
                    if (double.TryParse(cellVal, NumberStyles.Any, CultureInfo.InvariantCulture, out double parsed))
                        total += parsed;

                    currentRow++;
                }

                // 總計列
                for (int c = 0; c < headers.Count; c++)
                {
                    var colName = headers[c];
                    var cell = wsNew.Cell(currentRow, c + 1);
                    if (colName == totalColName)
                    {
                        cell.Value = total;
                        cell.Style.NumberFormat.Format = "\"NT$\"#,##0.00";
                    }
                    else if (c == 0)
                    {
                        cell.Value = "總計";
                    }

                    cell.Style.Font.Bold = true;
                    cell.Style.Font.FontName = "Calibri";
                }

                // 總表
                summarySheet.Cell(summaryRow, 1).Value = customer;
                summarySheet.Cell(summaryRow, 2).Value = subName;
                summarySheet.Cell(summaryRow, 3).Value = totalColName;
                summarySheet.Cell(summaryRow, 4).Value = total;
                summarySheet.Cell(summaryRow, 4).Style.NumberFormat.Format = "\"NT$\"#,##0.00";
                summaryRow++;
            }
        }

        // 儲存
        Directory.CreateDirectory(Path.GetDirectoryName(outputPath));
        newWb.SaveAs(outputPath);
        Console.WriteLine("✅ 檔案已儲存至：" + outputPath);

        GenrenralMail();
    }

    public static void GenrenralMail()
    {
        string basePath = AppDomain.CurrentDomain.BaseDirectory;
        string projectRoot = Directory.GetParent(basePath)       // bin\Debug\net6.0\
                                       ?.Parent                   // XML2Mail\
                                       ?.Parent                   // XML2Mail\
                                       ?.Parent?.FullName         // ✅ ← 這才是專案根目錄
                                       ?? throw new Exception("無法解析專案根目錄");

        string filePath = Path.Combine(projectRoot, $"{DateTime.Now.ToString("yyyyMMdd")} 處理完成.xlsx");
        var wb = new XLWorkbook(filePath);
        var sheet = wb.Worksheet("總表");

        var rows = sheet.RangeUsed().RowsUsed().Skip(1); // 跳過表頭

        var customerSums = new Dictionary<string, double>();
        var internalSums = new Dictionary<string, double>();

        // 收件人對應表
        var recipientMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "BA Microsoft Azure", "Ben, Victor" },
            { "Soetek BA", "Ben, Victor" },
            { "BASIS Microsoft Azure", "Ken" },
            { "FY Microsoft Azure", "Jimmy, Potter" },
            { "SMB Microsoft Azure", "Momo" },
            { "Soetek AI Microsoft Azure", "Charlie" }
        };

        // 客戶對應表, 要如何維護它
        var customerRecipientMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "美學品牌管理顧問股份有限公司", "Kiki" },
            { "潘朵拉傳藝有限公司", "許小姐" },
            { "台灣保時捷車業股份有限公司", "David" },
            { "車麗屋汽車百貨股份有限公司", "凱元" },  //要特別注意字元編碼 ⾞/車
            { "瑞士商福維克有限公司台灣分公司", "Edison" }, //要特別注意字元編碼 士/⼠
            { "愛貓一生實業股份有限公司", "鄭小姐" },
            { "鮮乳坊_慕渴股份有限公司", "鄭小姐" },
            { "台灣前川股份有限公司", "鄭小姐" }

        };

        foreach (var row in rows)
        {
            string customer = row.Cell(1).GetString().Trim();
            string subscription = row.Cell(2).GetString().Trim();
            string amountType = row.Cell(3).GetString().Trim();
            double amount = row.Cell(4).GetDouble();

            if (recipientMap.ContainsKey(subscription))
            {
                if (!internalSums.ContainsKey(subscription))
                    internalSums[subscription] = 0;
                internalSums[subscription] += amount;
            }

            if (customerRecipientMap.ContainsKey(customer))
            {
                if (!customerSums.ContainsKey(customer))
                    customerSums[customer] = 0;
                customerSums[customer] += amount;
            }
        }

        // 建立信件文字
        var allEmails = new StringBuilder();

        // === 客戶信件 ===
        foreach (var kv in customerSums)
        {
            string recipient = customerRecipientMap.ContainsKey(kv.Key) ? customerRecipientMap[kv.Key] : "";
            string greeting;
            if (recipient.EndsWith("小姐") || recipient.EndsWith("先生"))
                greeting = $"Dear {recipient},";
            else
                greeting = $"Dear {recipient},";

            allEmails.AppendLine($"[{kv.Key.Replace("有限公司", "").Replace("股份", "").Replace("台灣分公司", "")}] Azure 對帳單 {DateTime.Now.AddMonths(-1).ToString("yyyy/MM")}月");
            allEmails.AppendLine(greeting);
            allEmails.AppendLine();
            allEmails.AppendLine($"附件為Azure {DateTime.Now.AddMonths(-1).ToString("yyyy/MM")}月對帳單明細，實際金額 NT${kv.Value:N0} 元(未稅)，");
            allEmails.AppendLine("若對明細內容有任何疑問請再提出；若明細無誤也請回覆確認，謝謝。");
            allEmails.AppendLine();
            allEmails.AppendLine("註：有異議須 3 個工作天內回覆，否則視同確認無誤。");
            allEmails.AppendLine(new string('-', 60));
        }

        // === 內部信件 ===
        // 將 BA Microsoft Azure 與 Soetek BA 合併
        double baTotal = 0;
        foreach (var key in new[] { "BA Microsoft Azure", "Soetek BA" })
        {
            if (internalSums.ContainsKey(key))
                baTotal += internalSums[key];
        }
        if (baTotal > 0)
        {
            allEmails.AppendLine($"[BA] Azure 對帳單 {DateTime.Now.AddMonths(-1).ToString("yyyy/MM")}月");
            allEmails.AppendLine($"Dear Ben, Victor,");
            allEmails.AppendLine();
            allEmails.AppendLine($"附件為Azure {DateTime.Now.AddMonths(-1).ToString("yyyy/MM")}月對帳單明細，實際金額 NT${baTotal:N0} 元(未稅)，");
            allEmails.AppendLine("若對明細內容有任何疑問請再提出，謝謝。");
            allEmails.AppendLine(new string('-', 60));
        }
        // 其他內部信件
        foreach (var kv in internalSums)
        {
            if (kv.Key == "BA Microsoft Azure" || kv.Key == "Soetek BA")
                continue;
            string subscription = kv.Key;
            string recipients = recipientMap.ContainsKey(subscription) ? recipientMap[subscription] : "";

            allEmails.AppendLine($"[{subscription.Replace(" Microsoft Azure", "")}] Azure 對帳單 {DateTime.Now.AddMonths(-1).ToString("yyyy/MM")}月");
            allEmails.AppendLine($"Dear {recipients},");
            allEmails.AppendLine();
            allEmails.AppendLine($"附件為Azure {DateTime.Now.AddMonths(-1).ToString("yyyy/MM")}月對帳單明細，實際金額 NT${kv.Value:N0} 元(未稅)，");
            allEmails.AppendLine("若對明細內容有任何疑問請再提出，謝謝。");
            allEmails.AppendLine(new string('-', 60));
        }

        // === 會計信件 ===
        allEmails.AppendLine($"Azure 對帳單 {DateTime.Now.AddMonths(-1).ToString("yyyy/MM")}月");
        allEmails.AppendLine("Dear Evelyn,");
        allEmails.AppendLine();
        allEmails.AppendLine($"附件是零壹提供 {DateTime.Now.AddMonths(-1).ToString("yyyy/MM")}月的Azure對帳單，已經向客戶確認金額無誤，");
        allEmails.AppendLine("請協助開以下發票向客戶請款：");
        allEmails.AppendLine();

        foreach (var item in customerSums)
            allEmails.AppendLine($"{item.Key,-30} NT$ {item.Value,10:N0}");

        allEmails.AppendLine();
        allEmails.AppendLine("另外公司的費用分攤如下：");

        foreach (var item in internalSums)
        {
            if (item.Key == "BA Microsoft Azure" || item.Key == "Soetek BA")
                continue;
            allEmails.AppendLine($"{item.Key.Replace(" Microsoft Azure", ""),-30} NT$ {item.Value,10:N0}");
        }

        allEmails.AppendLine($"{"BA",-30} NT$ {baTotal,10:N0}");
        allEmails.AppendLine();
        allEmails.AppendLine("以上，再麻煩您，謝謝!");

        // === 輸出檔案 ===
        string outputPath = Path.Combine(projectRoot, $"{DateTime.Now.ToString("yyyyMMdd")} Azure_對帳單_信件總表.txt");
        Directory.CreateDirectory(Path.GetDirectoryName(outputPath));
        File.WriteAllText(outputPath, allEmails.ToString(), Encoding.UTF8);

        Console.WriteLine($"信件已成功輸出至：{outputPath}");
    }

    public class FakeGroup : IGrouping<string, IXLRangeRow>
    {
        private readonly IEnumerable<IXLRangeRow> _rows;
        public FakeGroup(IEnumerable<IXLRangeRow> rows) => _rows = rows;
        public string Key => null;
        public IEnumerator<IXLRangeRow> GetEnumerator() => _rows.GetEnumerator();
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => _rows.GetEnumerator();
    }
}
