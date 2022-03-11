using System;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;

namespace ComputerScienceGrades
{
    class Program
    {
        // секундомер
        static Stopwatch stopwatch = new();
        // переменная для хранения времени работы алгоритма
        static double timeWork;
        static Random random = new();
        // функция для подсчёта "5" в массиве
        static int FindingNumberFiveInArray(int[] array)
        {
            int count = 0;
            for (int i = 0; i < array.Length && count == i; i++)
            {
                count += (int)(array[i] / 5);//проверка на "5"
            }
            return count;
        }
        private static void Stopwatch(int[] array)
        {
            stopwatch.Reset();
            stopwatch.Start();
            FindingNumberFiveInArray(array);
            stopwatch.Stop();
            // время работы алгоритма в тиках
            timeWork += stopwatch.ElapsedTicks;
        }
        // лучший случай заполнения масиива
        static int[] Best(int[] array)
        {
            // Первый элемент не "5"
            for (int i = 1; i < array.Length; i++)
            {
                array[i] = random.Next(1, 5);//заполняем массив случайными числами
            }

            return array;
        }
        // Средний случай заполнения массива
        static int[] Average(int[] array)
        {
            int number = random.Next(1, array.Length);
            for (int i = 0; i < number; i++)
            {
                array[i] = 5;
            }
            for (int i = number; i < array.Length; i++)
            {
                array[i] = random.Next(1, 5);//заполняем массив случайными числами
            }
            return array;
        }
        // Худший вариант заполнения массива
        static int[] Worst(int[] array)
        {
            // Все элементы "5"
            for (int i = 0; i < array.Length; i++)
            {
                array[i] = 5;
            }

            return array;
        }
        static void Main(string[] args)
        {
            try
            {
                // создание excel  файла
                Excel.Application excel = new();
                excel.Visible = true;
                Excel.Workbook workBook = excel.Workbooks.Add(Type.Missing);
                Excel.Worksheet sheet = (Excel.Worksheet)excel.Worksheets.get_Item(1);
                sheet.get_Range("A1", "B1").Columns.EntireColumn.ColumnWidth = 30;
                sheet.Name = "График";
                // создаём массив длинной 20
                int[] pointsArray = new int[20];
                Best(pointsArray); // Холостой прогон
                Console.WriteLine("▬▬▬▬▬▬▬▬▬▬     Количество элементов массива 20    ▬▬▬▬▬▬▬▬▬▬\n" +
                    "----------     Количество испытаний 10 000 000    ----------");
                // обнуляем время перед началом серии экспериментов
                timeWork = 0;
                for (int j = 0; j < 10000000; j++)
                {
                    Stopwatch(Best(pointsArray));
                }
                // выводим время работы алгоритма в милллисекундах
                Console.WriteLine($"Лучший случай:  {timeWork / 10000:0.0}");
                sheet.Cells[1, 1] = String.Format("Лучший случай:");
                sheet.Cells[1, 2] = String.Format($"{timeWork / 10000:0.0}");
                sheet.get_Range("A1", "B1").Font.Color = ColorTranslator.ToOle(Color.Green);
                sheet.get_Range("A1", "B1").Font.Size = 20;
                timeWork = 0;
                for (int j = 0; j < 10000000; j++)
                {
                    Stopwatch(Average(pointsArray));
                }
                Console.WriteLine($"Средний случай: {timeWork / 10000:0.0}");
                sheet.Cells[2, 1] = String.Format("Средний случай:");
                sheet.Cells[2, 2] = String.Format($"{timeWork / 10000:0.0}");
                sheet.get_Range("A2", "B2").Font.Color = ColorTranslator.ToOle(Color.Blue);
                sheet.get_Range("A2", "B2").Font.Size = 20;
                timeWork = 0;
                for (int j = 0; j < 10000000; j++)
                {
                    Stopwatch(Worst(pointsArray));
                }
                Console.WriteLine($"Худший случай:  {timeWork / 10000:0.0}");
                sheet.Cells[3, 1] = String.Format("Худший случай:");
                sheet.Cells[3, 2] = String.Format($"{timeWork / 10000:0.0}");
                sheet.get_Range("A3", "B3").Font.Color = ColorTranslator.ToOle(Color.Orange);
                sheet.get_Range("A3", "B3").Font.Size = 20;
                // сохранение таблицы в папку "Документы" на диск "C"
                excel.Application.ActiveWorkbook.SaveAs("Graphic.xlsx");
                Console.ReadLine();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
