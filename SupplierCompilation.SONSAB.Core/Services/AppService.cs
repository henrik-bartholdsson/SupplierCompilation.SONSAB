using Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace SupplierCompilation.SONSAB.Core.Services
{
    public class AppService : IAppService
    {
        private string alternativeCountryCode = String.Empty;
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

            var applicationPath = Directory.GetCurrentDirectory();

            var workBook = excel.Workbooks.Open(applicationPath + "\\" + filePath);
            var worksheet = workBook.Worksheets[1];

            var lastUsedRow = GetLastRow(worksheet);
            string vatColumn = GetLastColumn(worksheet, 1);
            string nameColumn = GetLastColumn(worksheet, 2);
            string addressColumn = GetLastColumn(worksheet, 3);
            string secondAddressColumn = GetLastColumn(worksheet, 4);

            List<string> columns = new List<string>();
            columns.Add(vatColumn);
            columns.Add(nameColumn);
            columns.Add(addressColumn);
            columns.Add(secondAddressColumn);
            SetTitles(workBook, worksheet, columns);

            Range cells = worksheet.Range[$"{column}1:{column}{lastUsedRow}"];



            Console.Clear();
            Console.SetCursorPosition(0, 1);
            Console.WriteLine("========== Sonsab Supplier Compilator 1.0 ==========");

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

                if (String.IsNullOrEmpty(vatNumber) == false)
                {
                    if(String.IsNullOrEmpty(vatCc) == false)
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
                    }

                    if (String.IsNullOrEmpty(alternativeCountryCode) == false)
                    {
                        Range altCountryCode;
                        altCountryCode = worksheet.Range[$"{alternativeCountryCode}{i}:{alternativeCountryCode}{i}"];
                        var altCc = worksheet.Range[$"{alternativeCountryCode}{i}:{alternativeCountryCode}{i}"].Value;

                        if (vatCc != altCc)
                        {
                            var resp = _webService.SendRequest(altCc, vatNumber).Result;

                            if (resp.IsValid != "false")
                            {
                                vatCell.Value = resp.ContryCode + resp.VatNumber;
                                nameCell.Value = resp.Name;
                                addressCell.Value = resp.Address;
                                workBook.Save();
                                continue;
                            }
                        }
                    }

                    vatCell.Value = new String("invalid VAT");
                    workBook.Save();
                    continue;
                }

                vatCell.Value = new String("no VAT");
                workBook.Save();
            }

            workBook.Close();
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
                int modulo = (lastUsedColumn - 1) % 26;
                columnName = Convert.ToChar('A' + modulo) + columnName;
                lastUsedColumn = (lastUsedColumn - modulo) / 26;
            }

            return columnName;
        }

        private void SetTitles(Workbook workBook, dynamic? worksheet, List<string> columns)
        {
            Range vatTitle = worksheet.Range[$"{columns[0]}{1}:{columns[0]}{1}"];
            Range nameTitle = worksheet.Range[$"{columns[1]}{1}:{columns[1]}{1}"];
            Range adressTitle = worksheet.Range[$"{columns[2]}{1}:{columns[2]}{1}"];

            vatTitle.Value = "New Vat number";
            nameTitle.Value = "New Company Name";
            adressTitle.Value = "New Address";
            workBook.Save();
        }

        public void SetColumn(string column)
        {
            this.column = column;
        }

        public void AlternativeCountryCode(string countryCode)
        {
            this.alternativeCountryCode = countryCode;
        }
    }
}
