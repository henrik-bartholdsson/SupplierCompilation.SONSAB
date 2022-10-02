using Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace SupplierCompilation.SONSAB.Core.Services
{
    public class AppService : IAppService
    {
        private string column = String.Empty;
        private readonly Application excel;
        private readonly WebService _webService;

        public AppService()
        {
            excel = new Application();
            _webService = new WebService();
        }

        public void ProcessFile(string filePath)
        {
            string? vatCc = String.Empty;
            string? vatNumber = String.Empty;

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

            var lastUsedRow = GetLastRow(worksheet);
            string vatColumn = GetLastColumn(worksheet, 1);
            string nameColumn = GetLastColumn(worksheet, 2);
            string addressColumn = GetLastColumn(worksheet, 3);

            Range vatTitle = worksheet.Range[$"{vatColumn}{1}:{vatColumn}{1}"];
            Range nameTitle = worksheet.Range[$"{nameColumn}{1}:{nameColumn}{1}"];
            Range adressTitle = worksheet.Range[$"{addressColumn}{1}:{addressColumn}{1}"];

            vatTitle.Value = "New Vat number";
            nameTitle.Value = "New Company Name";
            adressTitle.Value = "New Address";
            workBook.Save();


            Range cells = worksheet.Range[$"{column}1:{column}{lastUsedRow}"];

            Console.Clear();

            for (int i = 2; i < lastUsedRow + 1; i++)
            {
                string cellValue = cells[i].Value.Replace("VAT", "");
                vatNumber = Regex.Replace(cellValue, @"[^0-9]", "");
                vatCc = Regex.Replace(cellValue, @"[^A-Z]", "");

                Console.SetCursorPosition(1, 15);
                Console.Write($"Rad {i} av {lastUsedRow}   ");

                Range vatCell = worksheet.Range[$"{vatColumn}{i}:{vatColumn}{i}"];
                Range nameCell = worksheet.Range[$"{nameColumn}{i}:{nameColumn}{i}"];
                Range addressCell = worksheet.Range[$"{addressColumn}{i}:{addressColumn}{i}"];

                if (!String.IsNullOrEmpty(vatNumber) && !String.IsNullOrEmpty(vatCc))
                {
                    var resp = _webService.SendRequest(vatCc, vatNumber).Result;

                    if (resp.IsValid != "false")
                    {
                        vatCell.Value = resp.ContryCode + resp.VatNumber;
                        nameCell.Value = resp.Name;
                        addressCell.Value = resp.Address;
                        workBook.Save();
                        continue;
                    }

                    vatCell.Value = new String("invalid VAT");
                    workBook.Save();
                    continue;
                }

                vatCell.Value = new String("no VAT");
                workBook.Save();
            }
        }

        private int GetLastRow(dynamic? worksheet)
        {
            var lastUsedRow = worksheet.Cells.Find("*", System.Reflection.Missing.Value,
                   System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                   Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows, Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious,
                   false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

            return (int)lastUsedRow;
        }
        private string GetLastColumn(dynamic? worksheet, int offset)
        {
            var lastUsedColumn = offset + worksheet.Cells.Find("*", System.Reflection.Missing.Value,
                   System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                   Microsoft.Office.Interop.Excel.XlSearchOrder.xlByColumns, Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious,
                   false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Column;

            string columnName = String.Empty;
            while (lastUsedColumn > 0)
            {
                int modulo = (lastUsedColumn -1) % 26;
                columnName = Convert.ToChar('A' + modulo) + columnName;
                lastUsedColumn = (lastUsedColumn - modulo) / 26;
            }

            return columnName;
        }

        public void SetColumn(string column)
        {
            this.column = column;
        }
    }
}
