namespace FeiNuo.Core;

/// <summary>
/// 基础的查询对象：包括一个时间段和模糊搜索查询
/// </summary>
public class BaseQuery
{
    /// <summary>
    /// 模糊搜索查询条件
    /// </summary>
    public string? Search { get; set; }

    /// <summary>
    /// 查询起始时间
    /// </summary>
    public DateOnly? StartDate { get; set; }

    /// <summary>
    /// 查询结束时间
    /// </summary>
    public DateOnly? EndDate { get; set; }
}