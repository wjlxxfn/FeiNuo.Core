namespace FeiNuo.Core.Tests
{
    [TestClass()]
    public class EnumExtensionsTests
    {
        [TestMethod()]
        public void GetDescriptionTest()
        {
            Assert.IsTrue(TestEnum.None.GetDescription() == "None");
            Assert.IsTrue(TestEnum.Test1.GetDescription() == "");
        }
    }

    public enum TestEnum
    {
        [System.ComponentModel.Description("None")]
        None = 1,
        Test1 = 2,
    }
}