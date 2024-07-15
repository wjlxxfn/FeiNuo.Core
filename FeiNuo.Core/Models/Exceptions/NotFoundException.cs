namespace FeiNuo.Core
{
    /// <summary>
    /// 找不到对象异常
    /// </summary>
    /// <param name="message">消息内容</param>
    public class NotFoundException(string message) : BaseException(message)
    {
    }
}