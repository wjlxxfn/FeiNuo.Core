using System.ComponentModel;
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
    [Description("消息类型：info,success,warning,error")]
    [JsonConverter(typeof(JsonStringEnumConverter))]
    public MessageTypeEnum Type { get; set; } = MessageTypeEnum.Info;

    /// <summary>
    /// 消息内容
    /// </summary>
    [Description("消息内容")]
    public string Message { get; set; } = string.Empty;

    /// <summary>
    /// 其它数据
    /// </summary>
    [Description("其它数据")]
    public object? Data { get; set; }

    /// <summary>
    /// 时间戳
    /// </summary>
    [Description("时间戳")]
    public DateTime Timestamp { get; set; } = DateTime.Now;

    ///<summary>
    /// 构造函数
    /// </summary>
    public MessageResult(string message, MessageTypeEnum type = MessageTypeEnum.Info, object? data = null)
    {
        Type = type;
        Message = message;
        Data = data;
        Timestamp = DateTime.Now;
    }
}
