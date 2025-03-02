using System.ComponentModel;

namespace FeiNuo.Core;

/// <summary>
/// 通用状态枚举
/// </summary>
public enum StatusEnum
{
    /// <summary>
    /// 正常状态
    /// </summary>
    [Description("正常")]
    Normal = 0,

    /// <summary>
    /// 作废状态
    /// </summary>
    [Description("作废")]
    Disabled = 1,

    /// <summary>
    /// 删除状态
    /// </summary>
    [Description("已删除")]
    Deleted = -1,
}
