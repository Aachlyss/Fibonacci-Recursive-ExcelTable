using System;
using System.Diagnostics;
using System.IO;
using ClosedXML.Excel;


namespace Fibonacci_Recursive_Timer
{
    class Program
    {
        static void Main()
        {
            Console.WriteLine("Enter the value of N:");
            int n = int.Parse(Console.ReadLine());

            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();

            long result = FibonacciRecursive(n);
            var time = stopwatch.ElapsedMilliseconds;

            stopwatch.Stop();
            Console.WriteLine($"Fibonacci({n}) = {result}");
            Console.WriteLine($"Execution Time: {time} ms");

            using (XLWorkbook workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("FibonacciTimes");
                worksheet.Cell("A1").Value = "N";
                worksheet.Cell("B1").Value = "Execution Time (ms)";

                for (int i = 1; i <= 1; i++)
                {
                    stopwatch.Restart();
                    FibonacciRecursive(i);
                    stopwatch.Stop();

                    worksheet.Cell(i + 1, 1).Value = n;
                    worksheet.Cell(i + 1, 2).Value = time;
                }

                workbook.SaveAs("fibonacci_times.xlsx");
            }
        }

        static long FibonacciRecursive(int n)
        {
            if (n <= 1)
                return n;
            else
                return FibonacciRecursive(n - 1) + FibonacciRecursive(n - 2);
        }
    }
}
