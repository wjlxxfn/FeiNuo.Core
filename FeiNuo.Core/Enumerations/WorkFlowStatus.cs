using System.ComponentModel;

namespace FeiNuo.Core;

/// <summary>
/// 工作流状态枚举
/// </summary>
public enum WorkFlowStatus
{
    /// <summary>
    /// 草稿状态
    /// </summary>
    [Description("草稿状态")]
    Draft = 0,

    /// <summary>
    /// 已提交
    /// </summary>
    [Description("已提交")]
    Submitted = 1,

    /// <summary>
    /// 审批中
    /// </summary>
    [Description("审批中")]
    InReview = 2,

    /// <summary>
    /// 已批准
    /// </summary>
    [Description("已批准")]
    Approved = 3,

    /// <summary>
    /// 已拒绝
    /// </summary>
    [Description("已拒绝")]
    Rejected = 4,

    /// <summary>
    /// 已取消
    /// </summary>
    [Description("已取消")]
    Cancelled = 5,

    /// <summary>
    /// 已完成
    /// </summary>
    [Description("已完成")]
    Completed = 6,

    /// <summary>
    /// 已归档
    /// </summary>
    [Description("已归档")]
    Archived = 9
}
