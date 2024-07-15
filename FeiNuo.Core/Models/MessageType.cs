using System.ComponentModel;

namespace FeiNuo.Core
{
    /// <summary>
    /// 消息类型
    /// </summary>
    [TypeConverter(typeof(EnumConverter))]
    public enum MessageType
    {
        /// <summary>
        /// 普通信息
        /// </summary>
        [Description("信息")]
        Info,

        /// <summary>
        /// 成功提示
        /// </summary>
        [Description("成功")]
        Success,

        /// <summary>
        /// 警告提示
        /// </summary>
        [Description("警告")]
        Warning,

        /// <summary>
        /// 错误提示
        /// </summary>
        [Description("错误")]
        Error
    }
}
