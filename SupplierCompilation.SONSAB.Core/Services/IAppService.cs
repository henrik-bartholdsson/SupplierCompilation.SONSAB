
using Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace SupplierCompilation.SONSAB.Core.Services
{
    public interface IAppService
    {
        public void ProcessFile(string filePath);
        public void SetColumn(string column);

    }

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
            if(String.IsNullOrEmpty(column))
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

            // Loop
            Range cells = worksheet.Range[$"{column}1:{column}{lastUsedRow}"];
            Console.Clear();
            foreach(var cell in cells.Value)
            {
                Console.WriteLine(cell);
            }
            Console.ReadLine();


        }

        public void SetColumn(string column)
        {
            this.column = column;
        }
    }
}
