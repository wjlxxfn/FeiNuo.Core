using System.ComponentModel;
using System.Text.Json.Serialization;

namespace FeiNuo.Core;

#region 分页对象 Pager
/// <summary>
/// 分页对象：需要页索引，每页记录数，以及排序信息
/// </summary>
public class Pager
{
    /// <summary>
    /// 当前页码
    /// </summary>
    [Description("当前页码")]
    public int PageNo { get; set; }

    /// <summary>
    /// 每页记录数
    /// </summary>
    [Description("每页记录数")]
    public int PageSize { get; set; }

    /// <summary>
    /// 排序
    /// </summary>
    [Description("排序")]
    public List<SortItem> SortBy { get; set; } = [];

    /// <summary>
    /// 构造函数
    /// </summary>
    public Pager() { }

    /// <summary>
    /// 构造函数
    /// </summary>
    public Pager(int pageNo, int pageSize)
    {
        PageNo = pageNo;
        PageSize = pageSize;
    }

    /// <summary>
    /// 返回一个不分页的分页对象
    /// </summary>
    [JsonIgnore]
    public static Pager Unpaged => new() { PageSize = 0 };

    /// <summary>
    /// 是否不分页
    /// </summary>
    public bool IsUnPaged => PageSize <= 0;
}
#endregion


/// <summary>
/// 排序对象
/// </summary>
public class SortItem
{
    /// <summary>
    /// 排序字段
    /// </summary>
    [Description("排序字段")]
    public string SortField { get; set; } = string.Empty;

    /// <summary>
    /// 排序类型
    /// </summary>
    [Description("排序类型")]
    public SortTypeEnum SortType { get; set; } = SortTypeEnum.ASC;

    /// <summary>
    /// 无参构造函数
    /// </summary>
    public SortItem()
    {
    }

    /// <summary>
    /// 构造函数
    /// </summary>
    public SortItem(string sortField, SortTypeEnum sortType = SortTypeEnum.ASC)
    {
        SortField = sortField;
        SortType = sortType;
    }
}

/// <summary>
/// 排序类型枚举
/// </summary>
public enum SortTypeEnum
{
    /// <summary>
    /// 升序
    /// </summary>
    [Description("升序")]
    ASC,

    /// <summary>
    /// 降序
    /// </summary>
    [Description("降序")]
    DESC
}