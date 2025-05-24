namespace FeiNuo.Core;

/// <summary>
/// 构造函数
/// </summary>
/// <param name="message">消息内容</param>
/// <param name="messageType">消息类型：info,success,warning,error</param>
public class MessageException(string message, MessageTypeEnum messageType = MessageTypeEnum.Info) : BaseException(message)
{
    /// <summary>
    /// 消息类型
    /// </summary>
    public MessageTypeEnum MessageType { get; set; } = messageType;
}

public class WarningException(string message) : MessageException(message, MessageTypeEnum.Warning)
{

}

public class ErrorException(string message) : MessageException(message, MessageTypeEnum.Error)
{

}