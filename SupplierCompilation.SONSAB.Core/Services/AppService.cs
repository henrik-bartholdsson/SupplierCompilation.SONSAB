
using Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace SupplierCompilation.SONSAB.Core.Services
{
    public class AppService : IAppService
    {
        private string column = String.Empty;
        private readonly Application excel;

        public AppService()
        {
            excel = new Application();
        }

        public void ProcessFile(string filePath)
        {
            string vatCc = String.Empty;
            string vatNumber = String.Empty;

            if (String.IsNullOrEmpty(column))
            {
                throw new Exception("Error, no column selected.");
            }
            if (String.IsNullOrEmpty(filePath))
            {
                throw new Exception("Error, file not found. Path: " + filePath);
            }

            var workBook = excel.Workbooks.Open(filePath);
            var worksheet = workBook.Worksheets[1];

            var lastUsedRow = worksheet.Cells.Find("*", System.Reflection.Missing.Value,
                               System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                               Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows, Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious,
                               false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

            //73 - lägg till i 74 (+1)
            var lastUsedColumn = 1 + worksheet.Cells.Find("*", System.Reflection.Missing.Value,
                               System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                               Microsoft.Office.Interop.Excel.XlSearchOrder.xlByColumns, Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious,
                               false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Column;

            string columnName = String.Empty;
            while (lastUsedColumn > 0)
            {
                int modulo = (lastUsedColumn - 1) % 26;
                columnName = Convert.ToChar('A' + modulo) + columnName;
                lastUsedColumn = (lastUsedColumn - modulo) / 26;
            }

            var webservice = new WebService();

            Range cells = worksheet.Range[$"{column}1:{column}{lastUsedRow}"];

            Console.Clear();
            for (int i = 2; i < lastUsedRow + 1; i++)
            {
                vatCc = String.Empty;
                vatNumber = String.Empty;

                string cellValue = cells[i].Value.Replace("VAT", "");
                vatNumber = Regex.Replace(cellValue, @"[^0-9]", "");
                vatCc = Regex.Replace(cellValue, @"[^A-Z]", "");

                Console.SetCursorPosition(1, 15);
                Console.Write($"Rad {i} av {lastUsedRow}   ");

                Range cellRange = worksheet.Range[$"{columnName}{i}:{columnName}{i}"];

                if (!String.IsNullOrEmpty(vatNumber) && !String.IsNullOrEmpty(vatCc))
                {
                    var resp = webservice.SendRequest(vatCc, vatNumber).Result;

                    if (resp.IsValid != "false")
                    {
                        cellRange.Value = resp.ContryCode + resp.VatNumber;
                        workBook.Save();
                        continue;
                    }

                    cellRange.Value = new String("invalid VAT");
                    workBook.Save();
                }

                cellRange.Value = new String("no VAT");
                workBook.Save();
            }

            workBook.Close();
        }

        public void SetColumn(string column)
        {
            this.column = column;
        }
    }
}
