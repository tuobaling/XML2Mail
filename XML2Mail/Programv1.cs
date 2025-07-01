using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
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

        string inputPath = Path.Combine(projectRoot, "202504-碩益科技股份有限公司.xlsx");
        string outputPath = Path.Combine(projectRoot, $"{DateTime.Now.ToString("yyyyMMdd")} 處理完成.xlsx");

        string inputTemplate = Path.Combine(projectRoot, @"寄件內文_Azure01.txt");
        string outputTextPath = Path.Combine(projectRoot, @"寄件內文_自動更新.txt");

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

        var wb = new XLWorkbook(outputPath);
        var sheet = wb.Worksheet("總表");

        // 建立客戶名稱 => 金額 字典
        var summary = sheet.RowsUsed()
            .Skip(1)
            .Where(r => !r.Cell(1).IsEmpty() && !r.Cell(4).IsEmpty())
            .GroupBy(r => new {
                Customer = r.Cell(1).GetString().Trim(),
                Subscription = r.Cell(2).GetString().Trim()
            })
            .ToDictionary(
                g => (g.Key.Customer, g.Key.Subscription),
                g => g.Sum(r => r.Cell(4).GetDouble())
            );

        var text = File.ReadAllText(inputTemplate);
        foreach (var (customer, amount) in summary)
        {
            // 建立 NT$ 數字格式
            string formatted = amount.ToString("C0", new CultureInfo("zh-TW"));

            // 嘗試替換內文中的金額
            string pattern = $@"(\[{Regex.Escape(customer.Customer)}_{Regex.Escape(customer.Subscription)}.*?實際金額\s*)(NT\$?|\$)?[\d,]+";
            text = Regex.Replace(text, pattern, $"$1{formatted}", RegexOptions.Multiline);
        }

        File.WriteAllText(outputTextPath, text);
        Console.WriteLine("寄件內容更新完成，檔案已輸出到 資料夾。");
    }

    class FakeGroup : IGrouping<string, IXLRangeRow>
    {
        private readonly IEnumerable<IXLRangeRow> _rows;
        public FakeGroup(IEnumerable<IXLRangeRow> rows) => _rows = rows;
        public string Key => null;
        public IEnumerator<IXLRangeRow> GetEnumerator() => _rows.GetEnumerator();
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => _rows.GetEnumerator();
    }
}
