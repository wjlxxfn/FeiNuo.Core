using FeiNuo.Core;

namespace ConsoleTest
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello, World!");

            var excel = PoiHelper.CreateExcel(out var s);
            s.DefaultStyle.BorderBottom = 0;
            var stream = File.Open(@"D:\wangjialiang\Desktop\test.xlsx", FileMode.Create);
            excel.Workbook.Write(stream);
            Console.WriteLine("End Line");
        }
    }
}
