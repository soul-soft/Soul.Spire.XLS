using Spire.Xls;
using Spire.Xls.Core;
using System.Diagnostics;

namespace Soul.Spire.XLS.Test
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var table = new XLSTable(
            [
                new XLSColumn("A")
                {
                    Items = 
                    [
                        new XLSColumn("E")
                        {
                            Width = 10,
                        }, 
                        new XLSColumn("F")
                        {
                            Width = 10,
                        },
                    ]
                },
                new XLSColumn("B"),
                new XLSColumn("C"),
                new XLSColumn("D"),
            ]);
            for (int i = 0; i < 3; i++)
            {
                var row = table.NewRow();
                row["E"] = "fa";
                row["F"] = "fa";
                row["B"] = "fa";
                row["C"] = "fa";
                row["D"] = "fa";
                table.Rows.Add(row);
            }

            var workbook = new Workbook();
            workbook.Worksheets.Clear();
            var worksheet = workbook.CreateEmptySheet();
            var info = worksheet.InsertTable(table, 2,2);
            info.ApplyDefaultStyle();

            workbook.SaveToFile("D:\\ff.xlsx", ExcelVersion.Version2016);
            Process.Start(new ProcessStartInfo
            {
                FileName = "D:\\ff.xlsx",
                UseShellExecute = true,
            });
            Console.WriteLine("Hello, World!");
        }
    }
}
