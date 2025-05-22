using NPOI.SS.UserModel;

namespace FeiNuo.Core.Tests
{
    [TestClass()]
    public class PoiHelperTests
    {
        static IWorkbook GetTestWorkbook()
        {
            string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "PoiTests/PoiTest.xlsx");
            var stream = File.OpenRead(filePath);
            return PoiHelper.CreateWorkbook(stream);
        }

        [TestMethod()]
        public void HelperTest()
        {
            var wb = GetTestWorkbook();
            var sheet = wb.GetSheetAt(0);
            sheet.ForceFormulaRecalculation = true;
            decimal num = 23.26m;
            Assert.AreEqual(num, PoiHelper.GetDecimalValue(PoiHelper.GetCell(sheet, 0, 1))); // 常规
            Assert.AreEqual(num, PoiHelper.GetDecimalValue(PoiHelper.GetCell(sheet, 0, 2))); // 文本
            Assert.AreEqual(num, PoiHelper.GetDecimalValue(PoiHelper.GetCell(sheet, 0, 3))); // 数字
            Assert.AreEqual(num, PoiHelper.GetDecimalValue(PoiHelper.GetCell(sheet, 0, 4))); // 日期
            Assert.ThrowsException<Exception>(() => PoiHelper.GetDecimalValue(PoiHelper.GetCell(sheet, 0, 6)));

            var dt = DateTime.Parse("2025-5-2");
            Assert.AreEqual(dt, PoiHelper.GetDateValue(PoiHelper.GetCell(sheet, 1, 1))); // 常规
            Assert.AreEqual(dt, PoiHelper.GetDateValue(PoiHelper.GetCell(sheet, 1, 2))); // 文本
            Assert.AreEqual(dt, PoiHelper.GetDateValue(PoiHelper.GetCell(sheet, 1, 3))); // 数字
            Assert.AreEqual(dt, PoiHelper.GetDateValue(PoiHelper.GetCell(sheet, 1, 4))); // 日期
            Assert.AreEqual(dt, PoiHelper.GetDateValue(PoiHelper.GetCell(sheet, 1, 5))); // 文本
            Assert.ThrowsException<Exception>(() => PoiHelper.GetDateValue(PoiHelper.GetCell(sheet, 0, 6)));

            num *= 2;
            Assert.AreEqual(num, PoiHelper.GetDecimalValue(PoiHelper.GetCell(sheet, 2, 1)));
            Assert.AreEqual(dt, PoiHelper.GetDateValue(PoiHelper.GetCell(sheet, 2, 2)));

        }
    }
}