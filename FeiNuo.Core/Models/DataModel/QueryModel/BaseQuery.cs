using System.ComponentModel;

namespace FeiNuo.Core;

/// <summary>
/// 基础的查询对象：包括一个时间段和模糊搜索查询
/// </summary>
public class BaseQuery
{
    /// <summary>
    /// 模糊搜索查询条件
    /// </summary>
    [Description("模糊搜索查询条件")]
    public string? Search { get; set; }

    /// <summary>
    /// 查询起始时间
    /// </summary>
    [Description("查询起始时间")]
    public DateOnly? StartDate { get; set; }

    /// <summary>
    /// 查询结束时间
    /// </summary>
    [Description("查询结束时间")]
    public DateOnly? EndDate { get; set; }
}