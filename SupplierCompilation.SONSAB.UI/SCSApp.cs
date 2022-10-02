using SupplierCompilation.SONSAB.Core.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SupplierCompilation.SONSAB.UI
{
    internal class SCSApp
    {
        AppService _appService;
        public SCSApp()
        {
            _appService = new AppService();
        }

        public void Run()
        {
            while (true)
            {
                PrintMenu();
                var input = Console.ReadKey();

                if (input.Key == ConsoleKey.D1)
                    PrintSecondMenu();

                if (input.Key == ConsoleKey.D2)
                {
                    Console.WriteLine("Not implemented yet... Press Enter to continue");
                    Console.Read();
                }

                if (input.Key == ConsoleKey.F3)
                    return;
            }
        }

        private void PrintMenu()
        {
            Console.Clear();
            Console.SetCursorPosition(0, 1);
            Console.WriteLine("========== Sonsab Supplier Compilator 1.0 ==========");
            Console.SetCursorPosition(1, 3);
            Console.WriteLine("1. Processa Fil");
            Console.SetCursorPosition(1, 5);
            Console.WriteLine("2. Fråga enskilt Vat nr");
            Console.SetCursorPosition(1, 25);
            Console.WriteLine("[F3] för att avsluta");
            Console.SetCursorPosition(1, 12);
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Fil sorce.xlsx hittad");
            Console.ResetColor();
        }

        private void PrintSecondMenu()
        {
            Console.Clear();
            Console.SetCursorPosition(0, 1);
            Console.WriteLine("========== Sonsab Supplier Compilator 1.0 ==========");
            Console.SetCursorPosition(1, 25);
            Console.WriteLine("[F3] för att avsluta");
            Console.SetCursorPosition(1, 3);
            Console.Write("Välj column för Vat: ");
            Console.SetCursorPosition(1, 4);
            Console.WriteLine("Välj alternativ column för landskod (om det finns):");
            Console.SetCursorPosition(21, 3);
            var column = Console.ReadLine().ToUpper();
            Console.SetCursorPosition(21, 3);
            Console.WriteLine(column);
            Console.SetCursorPosition(52, 4);
            var countryCode = Console.ReadLine().ToUpper();
            Console.SetCursorPosition(52, 4);
            Console.WriteLine(countryCode);

            if (string.IsNullOrEmpty(column))
                return;

            try
            {
                int a = int.Parse(column);
                return;
            }
            catch { }

            Console.SetCursorPosition(1, 7);
            Console.WriteLine("Du valde kolumn " + column + " för VAT-nummer");

            if (String.IsNullOrEmpty(countryCode) == false)
            {
                Console.SetCursorPosition(1, 8);
                Console.WriteLine("Du valde kolumn " + countryCode + " för alternativ landskod.");
            }

            var input = Console.ReadKey();

            if (input.Key == ConsoleKey.Enter)
            {
                Console.WriteLine("Running service...");

                _appService.SetColumn(column);
                _appService.AlternativeCountryCode(countryCode);
                try
                {
                    _appService.ProcessFile(@"c:\tmp\FT.lev.xlsx");
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error: " + ex.Message);
                }
                return;
            }

            if (input.Key == ConsoleKey.F3)
                return;
        }
    }
}
