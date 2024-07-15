namespace FeiNuo.Core.Tests
{
    [TestClass()]
    public class StringExtensionsTests
    {
        [TestMethod()]
        public void ToUpperFirstTest()
        {
            Assert.IsTrue("a".ToUpperFirst() == "A");
            Assert.IsTrue("".ToUpperFirst() == "");
            Assert.IsTrue("string".ToUpperFirst() == "String");
            Assert.IsTrue("String".ToUpperFirst() == "String");
            Assert.IsTrue(" string".ToUpperFirst() == " string");
            Assert.IsTrue("右string".ToUpperFirst() == "右string");
            Assert.IsTrue("".ToUpperFirst() == "");
            Console.WriteLine("test");
        }

        [TestMethod()]
        public void ToLowerFirstTest()
        {
            Assert.IsTrue("String".ToLowerFirst() == "string");
            Assert.IsTrue("string".ToLowerFirst() == "string");
            Assert.IsTrue(" string".ToLowerFirst() == " string");
            Assert.IsTrue("右string".ToLowerFirst() == "右string");
            Assert.IsTrue("".ToLowerFirst() == "");
        }
    }
}