﻿using SupplierCompilation.SONSAB.Core.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SupplierCompilation.SONSAB.UI
{
    internal class SCSApp
    {
        public SCSApp()
        {
            
        }

        public void Run()
        {
            while(true)
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
            Console.SetCursorPosition(1, 8);
            Console.WriteLine("[F3] för att avsluta");
            Console.SetCursorPosition(1, 12);
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Fil sorce.xlsx hittad");
            Console.ResetColor();
        }

        private void PrintSecondMenu()
        {
            Console.SetCursorPosition(1, 3);
            Console.Write("Välj column för Vat:");
            Console.SetCursorPosition(1, 5);
            Console.WriteLine(new String(' ', Console.BufferWidth));
            Console.SetCursorPosition(21, 3);
            var column = Console.ReadLine();

            if (string.IsNullOrEmpty(column))
                return;

            try
            {
                int a = int.Parse(column);
                return;
            }
            catch { }

            Console.SetCursorPosition(1, 5);
            Console.WriteLine("Du valde kolumn " + column.ToUpper() + ", tryck Enter för att processa filen eller F3 för att avsluta.");

            var input = Console.ReadKey();

            if (input.Key == ConsoleKey.Enter)
            {
                Console.WriteLine("Running service...");
                var serviceOne = new AppService();
                serviceOne.SetColumn(column);
                serviceOne.ProcessFile(@"c:\tmp\FT.xlsx");
                return;
            }

            if (input.Key == ConsoleKey.F3)
                return;
        }
    }
}