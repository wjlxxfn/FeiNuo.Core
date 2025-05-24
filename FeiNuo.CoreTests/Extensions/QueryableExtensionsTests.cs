

namespace FeiNuo.Core.Tests
{
    [TestClass()]
    public class QueryableExtensionsTests
    {
        [TestMethod()]
        public void OrderByTest()
        {
            var list = new List<TestEntity>()
            {
                new ("1aaa","4bbb"),
                new ("1aaa","2bbb"),
                new ("3aaa","2bbb"),
                new ("4aaa","1bbb"),
            };

            var result = list.AsQueryable().OrderBy("Key").ThenBy("Value").First();
            Assert.AreEqual("1aaa", result.Key);
            Assert.AreEqual("2bbb", result.Value);

            result = list.AsQueryable().OrderBy("Key").ThenBy("Value").First();
            Assert.AreEqual("1aaa", result.Key);
            Assert.AreEqual("4bbb", result.Value);

            Assert.ThrowsException<ArgumentException>(() => list.AsQueryable().OrderBy("KeyNotExists"));

        }
    }


    internal class TestEntity
    {
        public TestEntity(string key, string value)
        {
            Key = key;
            Value = value;
        }

        public string Key { get; set; }
        public string Value { get; set; }
    }
}