using System;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;
using System.Collections.Generic;

namespace ComputerScienceGrades
{
    class Program
    {       
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
            // секундомер
            Stopwatch stopwatch = new();
            timeWork = 0;
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
                List<int> n = new List<int> { 1000000, 2000000, 3000000, 4000000, 5000000 };
                // количество оиспытаний
                int quantity = 10;
                // создание excel  файла
                Excel.Application excel = new();
                Excel.Workbook workBook = excel.Workbooks.Add(Type.Missing);
                Excel.Worksheet sheet = (Excel.Worksheet)excel.Worksheets.get_Item(1);
                sheet.Name = "График";
                int x = 2, y = 2;
                foreach (int k in n)
                {
                    sheet.Cells[x, 1] = String.Format($"{k}");
                    int[] pointsArray = new int[k];
                    Best(pointsArray); // Холостой прогон
                    Console.WriteLine($"▬▬▬▬▬▬▬▬▬▬     Количество элементов массива {k}    ▬▬▬▬▬▬▬▬▬▬");
                    for (int j = 0; j < quantity; j++)
                    {
                        Stopwatch(Best(pointsArray));
                    }
                    // выводим время работы алгоритма в милллисекундах
                    Console.WriteLine($"Лучший случай:  {timeWork / quantity:0.0}");
                    sheet.Cells[x, y].Formula = $"={timeWork}/{quantity}";
                    y++;
                    for (int j = 0; j < quantity; j++)
                    {
                        Stopwatch(Average(pointsArray));
                    }
                    Console.WriteLine($"Средний случай: {timeWork / quantity:0.0}");
                    sheet.Cells[x, y].Formula = $"={timeWork}/{quantity}";
                    y++;
                    for (int j = 0; j < quantity; j++)
                    {
                        Stopwatch(Worst(pointsArray));
                    }
                    Console.WriteLine($"Худший случай:  {timeWork / quantity:0.0}");
                    sheet.Cells[x, y].Formula = $"={timeWork}/{quantity}";
                    x++; y = 2;
                }
                Excel.Range range = sheet.get_Range("A1", "D6");
                sheet.Cells[1, 2] = String.Format("Лучший случай");
                sheet.Cells[1, 3] = String.Format("Средний случай");
                sheet.Cells[1, 4] = String.Format("Худший случай");
                sheet.get_Range("A1", "D1").Columns.EntireColumn.ColumnWidth = 23;
                range.Cells.Font.Size = 16;
                range.Cells.Font.Name = "Times New Roman";
                range.HorizontalAlignment = Excel.Constants.xlCenter;
                range.VerticalAlignment = Excel.Constants.xlCenter;
                sheet.get_Range("B2", "D6").Cells.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(200, 170, 190));
                sheet.get_Range("B1", "D1").Cells.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(140, 200, 130));
                sheet.get_Range("A2", "A6").Cells.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(230, 170, 170));
                range.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous;
                range.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous;
                range.Borders.get_Item(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlContinuous;
                range.Borders.get_Item(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlContinuous;
                range.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous;
                Excel.ChartObjects chartsobjrcts = (Excel.ChartObjects)sheet.ChartObjects(Type.Missing);
                Excel.ChartObject chartsobjrct = chartsobjrcts.Add(520, 10, 300, 200);
                chartsobjrct.Chart.ChartWizard(sheet.get_Range("A1", "D6"), Excel.XlChartType.xlLine, 2, Excel.XlRowCol.xlColumns, 
                    Type.Missing, -1, true, "Алгоритм", "Длина массива", "Среднее время");
                // сохранение таблицы в папку "Документы" на диск "C"
                excel.Application.DisplayAlerts = false;
                excel.Application.ActiveWorkbook.SaveAs("Graphic.xlsx");
                workBook.Save();
                excel.Visible = true;
                Console.ReadLine();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
