namespace FeiNuo.CoreTests
{
    [TestClass()]
    public class ConsoleTests
    {
        [TestMethod()]
        public void Test()
        {
            Console.WriteLine("HelloWord");

            Assert.IsTrue(true);
        }
    }
}
