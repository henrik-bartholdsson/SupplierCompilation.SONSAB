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
                    PrintTheredMenu();
                }

                if (input.Key == ConsoleKey.F3)
                    return;

                if (input.Key == ConsoleKey.F6)
                {
                    HelpSection();
                }
            }
        }

        private void PrintMenu()
        {
            Console.Clear();
            Console.SetCursorPosition(0, 1);
            Console.WriteLine("========== Sonsab Supplier Compilator 1.0 ==========");
            Console.SetCursorPosition(1, 3);
            Console.WriteLine("1. Processa Fil VAT");
            Console.SetCursorPosition(1, 5);
            Console.WriteLine("2. Processa Fil Org.nr");
            Console.SetCursorPosition(1, 25);
            Console.WriteLine("[F3] för att avsluta");
            Console.SetCursorPosition(27, 25);
            Console.WriteLine("[F6] för hjälp");
            Console.SetCursorPosition(0, 23);
            Console.WriteLine("=====================================================");
            Console.SetCursorPosition(1, 22);
            if(File.Exists("supplier.xlsx"))
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Fil supplier.xlsx hittad");
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Kan ej hitta fil.");
            }
            Console.ResetColor();
            Console.SetCursorPosition(0, 26);
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

            Console.SetCursorPosition(1, 10);
            Console.WriteLine("Tryck Enter för att starta körningen, eller F3 för att återgå.");

            var input = Console.ReadKey();

            if (input.Key == ConsoleKey.Enter)
            {
                if(!File.Exists("supplier.xlsx"))
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine(" Kan ej hitta fil!");
                    Console.ResetColor();
                    Console.WriteLine(" Tryck enter för att börja om.");
                    Console.ReadKey();
                    return;
                }
                Console.WriteLine("Running service...");

                _appService.SetColumn(column);
                _appService.AlternativeCountryCode(countryCode);
                
                try
                {
                    _appService.ProcessVatFile("supplier.xlsx");
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error: " + ex.Message);
                }
                return;
            }

            if (input.Key == ConsoleKey.F3)
                return;

            if(input.Key == ConsoleKey.F6)
            {
                HelpSection();
            }
        }

        private void PrintTheredMenu()
        {
            Console.Clear();
            Console.SetCursorPosition(0, 1);
            Console.WriteLine("========== Sonsab Supplier Compilator 1.0 ==========");
            Console.SetCursorPosition(1, 25);
            Console.WriteLine("[F3] för att avsluta");
            Console.SetCursorPosition(1, 3);
            Console.Write("Välj column för org.nr: ");
            Console.SetCursorPosition(1, 4);
            Console.WriteLine("Välj alternativ column för landskod (om det finns):");
            Console.SetCursorPosition(25, 3);
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
            Console.WriteLine("Du valde kolumn " + column + " för org.nr");

            if (String.IsNullOrEmpty(countryCode) == false)
            {
                Console.SetCursorPosition(1, 8);
                Console.WriteLine("Du valde kolumn " + countryCode + " för alternativ landskod.");
            }

            Console.SetCursorPosition(1, 10);
            Console.WriteLine("Tryck Enter för att starta körningen, eller F3 för att återgå.");

            var input = Console.ReadKey();

            if (input.Key == ConsoleKey.Enter)
            {
                if (!File.Exists("supplier.xlsx"))
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine(" Kan ej hitta fil!");
                    Console.ResetColor();
                    Console.WriteLine(" Tryck enter för att börja om.");
                    Console.ReadKey();
                    return;
                }
                Console.WriteLine("Running service...");

                _appService.SetColumn(column);
                _appService.AlternativeCountryCode(countryCode);

                try
                {
                    _appService.ProcessOrgNrFile("supplier.xlsx");
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error: " + ex.Message);
                }
                return;
            }

            if (input.Key == ConsoleKey.F3)
                return;

            if (input.Key == ConsoleKey.F6)
            {
                HelpSection();
            }
        }

        private void HelpSection()
        {
            Console.Clear();
            Console.WriteLine();
            Console.WriteLine("==== Hjälpavsnitt ====");
            Console.WriteLine();
            Console.WriteLine(" Excel-filen som ska bearbetas måste ligga i samma mapp som programmet.");
            Console.WriteLine(" Filen måste ha namnet supplier.xlsx");
            Console.WriteLine("");
            Console.WriteLine(" När filen ska bearbetas måste en kolumn som innhåller de troliga VAT-nummren anges.");
            Console.WriteLine("");
            Console.WriteLine(" En alternativ kolum för landskod kan anges, detta kan underlätta vid uppslag mot web-tjänsten.");
            Console.WriteLine("");
            Console.WriteLine("");
            Console.WriteLine(" Resultatet av körningen kommer läggas i de kolumner efter de avslutande.");
            Console.WriteLine("");
            Console.WriteLine("");
            Console.WriteLine("");
            Console.WriteLine(" Tryck på valfri tangent för att återgå...");
            Console.ReadKey();
        }
    }
}
