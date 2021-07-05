using SpreadsheetLight;
using System;

var NumberOfRows = 200_000;
var watch = new System.Diagnostics.Stopwatch();
var rand = new Random();

Console.WriteLine("Start of benchmark");
Console.WriteLine("Total Memory: {0}", GC.GetTotalMemory(true));
watch.Start();

var sl = new SLDocument();

watch.Stop();
Console.WriteLine("After new spreadsheet initialization: {0}", watch.Elapsed.ToString());
Console.WriteLine("Total Memory: {0}", GC.GetTotalMemory(true));
watch.Start();

var style = new SLStyle();
style.FormatCode = "dd/mm/yyyy";
// this sets the style for columns 6 through 8 as "dd/mm/yyyy"
// because these columns contain date values.
// Doing it only for the 1st worksheet as an example.
sl.SetColumnStyle(6, 8, style);

watch.Stop();
Console.WriteLine("After setting date format: {0}", watch.Elapsed.ToString());
Console.WriteLine("Total Memory: {0}", GC.GetTotalMemory(true));
watch.Start();

for (var i = 1; i <= NumberOfRows; ++i)
{
    sl.SetCellValue(i, 1, string.Format("R{0}T{1}", i, rand.Next(10)));
    sl.SetCellValue(i, 2, string.Format("R{0}T{1}", i, rand.Next(10)));
    sl.SetCellValue(i, 3, string.Format("R{0}T{1}", i, rand.Next(10)));
    sl.SetCellValue(i, 4, string.Format("R{0}T{1}", i, rand.Next(10)));
    sl.SetCellValue(i, 5, string.Format("R{0}T{1}", i, rand.Next(10)));
    sl.SetCellValue(i, 6, DateTime.Now.AddDays(rand.NextDouble() * 10.0));
    sl.SetCellValue(i, 7, DateTime.Now.AddDays(rand.NextDouble() * 10.0));
    sl.SetCellValue(i, 8, DateTime.Now.AddDays(rand.NextDouble() * 10.0));
    sl.SetCellValue(i, 9, rand.Next(1000));
    sl.SetCellValue(i, 10, rand.Next(2000));
    sl.SetCellValue(i, 11, rand.Next(3000));
    sl.SetCellValue(i, 12, rand.Next(4000));
    sl.SetCellValue(i, 13, rand.Next(5000));
    sl.SetCellValue(i, 14, rand.NextDouble() * 10000.0);
    sl.SetCellValue(i, 15, rand.NextDouble() * 10000.0);
    sl.SetCellValue(i, 16, rand.NextDouble() * 10000.0);
    sl.SetCellValue(i, 17, rand.NextDouble() * 10000.0);
    sl.SetCellValue(i, 18, rand.NextDouble() * 10000.0);
}

watch.Stop();
Console.WriteLine("After writing 1st worksheet: {0}", watch.Elapsed.ToString());
Console.WriteLine("Total Memory: {0}", GC.GetTotalMemory(true));
watch.Start();

sl.SaveAs("BenchmarkWriteCells.xlsx");

watch.Stop();
Console.WriteLine("After saving: {0}", watch.Elapsed.ToString());
Console.WriteLine("Total Memory: {0}", GC.GetTotalMemory(true));

Console.WriteLine("End of program");
Console.ReadLine();