using System.Text.Json.Serialization;

namespace FeiNuo.Core;

/// <summary>
/// 返回给前端的响应消息
/// </summary>
public class MessageResult
{
    /// <summary>
    /// 消息类型：info,success,warning,error
    /// </summary>
    [JsonConverter(typeof(JsonStringEnumConverter))]
    public MessageType Type { get; set; } = MessageType.Info;

    /// <summary>
    /// 消息内容
    /// </summary>
    public string Message { get; set; } = string.Empty;

    /// <summary>
    /// 其它数据
    /// </summary>
    public object? Data { get; set; }

    /// <summary>
    /// 时间戳
    /// </summary>
    public DateTime Timestamp { get; set; } = DateTime.Now;

    ///<summary>
    /// 构造函数
    /// </summary>
    public MessageResult(string message, MessageType type = MessageType.Info, object? data = null)
    {
        Message = message;
        Type = type;
        Data = data;
        Timestamp = DateTime.Now;
    }
}
