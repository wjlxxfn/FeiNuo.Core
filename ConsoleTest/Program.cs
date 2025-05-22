using FeiNuo.Core;

namespace ConsoleTest
{
    internal class Program
    {
        static void Main()
        {
            Console.WriteLine("Hello, World!");

            var excel = PoiHelper.CreateExcel(["aa", "bb", "cc"],
                new ExcelColumn<string>("标题1", v => v + "标题1标题1标题1\n标题11标题1标题1标题11", 12, w => w.Wrap()),
                new ExcelColumn<string>("标题2", v => v + "\n22", 12),
                new ExcelColumn<string>("标题3", v => v + "33", 12)
            ); ;
            var stream = File.Open(@$"D:\test{DateTime.Now.ToUnixTimeMilliseconds()}.xlsx", FileMode.Create);
            excel.Workbook.Write(stream);
            Console.WriteLine("End Line");
        }
    }
}
