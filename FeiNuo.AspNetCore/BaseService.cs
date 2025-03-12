using Microsoft.EntityFrameworkCore;

namespace FeiNuo.Core;

/// <summary>
/// 服务类顶层接口，实现该接口可自动注入
/// </summary>
public interface IService { }

/// <summary>
/// 服务类基类
/// </summary>
public abstract class BaseService : IService
{

}

/// <summary>
/// 服务类基类：指定数据库上下文和实体类型，提供基础操作
/// </summary>
public abstract class BaseService<T> : BaseService where T : class, new()
{
    #region 构造函数
    private readonly DbContext ctx;
    public BaseService(DbContext ctx)
    {
        this.ctx = ctx;
    }
    #endregion

    #region 查询操作
    /// <summary>
    /// 根据ID查询实体类
    /// </summary>
    protected virtual async Task<T> FindByIdAsync(object id)
    {
        var obj = await ctx.FindAsync<T>(id)
            ?? throw new NotFoundException($"找不到指定数据,Id:{id},Type:{typeof(T)}");
        return obj;
    }

    /// <summary>
    /// 分页查询
    /// </summary>
    /// <param name="query">查询条件</param>
    /// <param name="pager">分页条件</param>
    /// <param name="defaultOrder">默认排序</param>
    /// <returns>分页数据</returns>
    protected async Task<PageResult<T>> FindPagedList(IQueryExpression<T> query, Pager pager, Func<IQueryable<T>, IOrderedQueryable<T>> defaultOrder)
    {
        var querable = ctx.Set<T>().Where(query.GetQueryExpression());
        var orderdQuery = PageHelper.ApplySort(querable, pager, defaultOrder);
        return await PageHelper.FindPagedList(orderdQuery, pager);
    }
    #endregion
}