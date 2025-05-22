using FeiNuo.Core;

namespace ConsoleTest
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello, World!");
            var config = new ExcelConfig("测试.xlsx")
                .AddExcelSheet(new ExcelSheet(new List<ExcelColumn<dynamic>>()
                {
                    new ("姓名", v => v.name, 12,s=>s.Border(0)),
                    new ("日期", v => v.date, 12,s=>s.Format("yyyy-mm-dd").Border(0)),
                    new ("北京", v => v.bb, 22,s=>s.BgColor(26).Border(0)),
                }).AddDescription("测试表", 12)
                .AddMainTitle("maintitle", 10)
                //.AddDataList(new List<dynamic>() {
                //    new { name="name",date=DateTime.Now,bb="sss"},
                //    new { name="nameaaa",date=DateTime.Now,bb="sssfff"},
                //    new { name="nameaccca",date=DateTime.Now,bb="ssddsfff"},
                //})
                );
            var excel = PoiHelper.CreateExcel(config);
            var stream = File.Open(@$"D:\test{DateTime.Now.ToUnixTimeMilliseconds().ToString()}.xlsx", FileMode.Create);
            excel.Workbook.Write(stream);
            Console.WriteLine("End Line");
        }
    }
}
