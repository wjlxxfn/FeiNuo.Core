namespace FeiNuo.Core
{
    /// <summary>
    /// 异常基类
    /// </summary>
    public class BaseException : Exception
    {
        /// <summary>
        /// 默认构造函数
        /// </summary>
        public BaseException() : base() { }

        /// <summary>
        /// 构造函数：错误消息
        /// </summary>
        public BaseException(string message) : base(message) { }

        /// <summary>
        /// 构造函数：错误消息+异常
        /// </summary>
        public BaseException(string message, Exception innerException) : base(message, innerException) { }
    }
}