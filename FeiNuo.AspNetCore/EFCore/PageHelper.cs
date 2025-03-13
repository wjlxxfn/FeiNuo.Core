using Microsoft.EntityFrameworkCore;

namespace FeiNuo.AspNetCore;

/// <summary>
/// 分页辅助类：执行分页查询
/// 使用该类，需引入Microsoft.EntityFrameworkCore
/// </summary>
public class PageHelper
{
    /// <summary>
    /// 根据条件分页查询,会先添加排序条件
    /// </summary>
    /// <param name="query">自定义的查询</param>
    /// <param name="pager">分页条件</param>
    /// <param name="defaultOrder">默认排序</param>
    /// <returns>分页数据</returns>
    public static async Task<PageResult<TEntity>> FindPagedList<TEntity>(IQueryable<TEntity> query, Pager pager, Func<IQueryable<TEntity>, IOrderedQueryable<TEntity>> defaultOrder) where TEntity : class
    {
        var orderdQuery = ApplySort(query, pager, defaultOrder);
        return await FindPagedList(orderdQuery, pager);
    }

    /// <summary>
    /// 执行分页查询
    /// </summary>
    /// <typeparam name="TEntity">实体对象</typeparam>
    /// <param name="query">linq查询(已排序)</param>
    /// <param name="pager">分页参数</param>
    /// <returns>PageResult</returns>
    public static async Task<PageResult<TEntity>> FindPagedList<TEntity>(IOrderedQueryable<TEntity> query, Pager pager) where TEntity : class
    {
        if (pager.PageSize <= 0 || pager.PageSize == int.MaxValue) // 没有分页参数直接查全部
        {
            var lstData = await query.ToListAsync();
            return new PageResult<TEntity>(lstData, lstData.Count);
        }
        // 查询分页数据
        var skipSize = Math.Max(pager.PageNo - 1, 0) * pager.PageSize;
        var pageData = await query.Skip(skipSize).Take(pager.PageSize).ToListAsync();
        // 查询总数量
        int totalCount = skipSize + pageData.Count;
        // 没有数据，直接返回
        if (totalCount == 0) return new PageResult<TEntity>();
        // 可能还有数据，查下总数
        if (pageData.Count == pager.PageSize)
        {
            totalCount = await query.AsNoTracking().CountAsync();
        }
        return new PageResult<TEntity>(pageData, totalCount);
    }


    /// <summary>
    /// 添加排序条件，分页前必须排序，
    /// </summary>
    public static IOrderedQueryable<TEntity> ApplySort<TEntity>(IQueryable<TEntity> query, Pager pager, Func<IQueryable<TEntity>, IOrderedQueryable<TEntity>> defaultOrder) where TEntity : class
    {
        if (pager.SortBy.Count == 0)
        {
            // 没有排序条件添加默认排序
            return defaultOrder(query);
        }

        // 没有排序的话按主键排序
        //if (!pager.IsSorted)
        //{
        //    pager.SortField = ctx.Model.FindEntityType(typeof(TEntity))!.FindPrimaryKey()!.Properties.Select(a => a.PropertyInfo!.Name).ToArray();
        //    pager.SortType = pager.SortField.Select(a => SortTypeEnum.DESC.ToString()).ToArray();
        //}

        // 添加排序
        var orderedQuery = query.OrderBy(pager.SortBy[0].SortField, pager.SortBy[0].SortType == SortTypeEnum.ASC);
        for (var i = 1; i < pager.SortBy.Count; i++)
        {
            orderedQuery = orderedQuery.ThenBy(pager.SortBy[i].SortField, pager.SortBy[i].SortType == SortTypeEnum.ASC);
        }
        return orderedQuery;
    }

    /// <summary>
    /// 内存分页
    /// </summary>
    /// <param name="lstAll">已排序的数据</param>
    /// <param name="pager">分页参数</param>
    /// <returns>PageResult</returns>
    public static PageResult<T> Page<T>(IOrderedEnumerable<T> lstAll, Pager pager)
    {
        var totalCount = lstAll.Count();
        var skipSize = Math.Max(pager.PageNo - 1, 0) * pager.PageSize;
        var pageData = lstAll.Skip(skipSize).Take(pager.PageSize);
        return new PageResult<T>(pageData.ToList(), totalCount);
    }
}
