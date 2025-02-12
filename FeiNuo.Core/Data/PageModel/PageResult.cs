namespace FeiNuo.Core;

/// <summary>
/// 分页数据对象
/// </summary>
/// <typeparam name="TEntity">Dto对象</typeparam>
public class PageResult<TEntity>
{
    /// <summary>
    /// 总记录数
    /// </summary>
    public int TotalCount { get; set; }

    /// <summary>
    /// 分页数据
    /// </summary>
    public IEnumerable<TEntity> DataList { get; set; } = [];

    /// <summary>
    /// 无参构造函数
    /// </summary>
    public PageResult() { }
    /// <summary>
    /// 构造函数
    /// </summary>
    /// <param name="dataList">当前页的数据</param>
    /// <param name="totalCount">数据总条数</param>
    public PageResult(IEnumerable<TEntity> dataList, int totalCount)
    {
        DataList = dataList;
        TotalCount = totalCount;
    }

    /// <summary>
    /// 数据转换
    /// </summary>
    public PageResult<T> Map<T>(Func<TEntity, T> exp)
    {
        var newContent = DataList.Select(exp.Invoke).ToList();
        return new PageResult<T>(newContent, TotalCount);
    }
}