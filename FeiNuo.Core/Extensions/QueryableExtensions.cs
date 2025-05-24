using System.Linq.Expressions;

namespace FeiNuo.Core;

/// <summary>
/// 查询扩展：添加通过字符串排序的方法
/// </summary>
public static class QueryableExtensions
{
    /// <summary>
    /// 根据属性名添加排序条件
    /// </summary>
    /// <param name="source">IOrderedQueryable</param>
    /// <param name="propertyName">要排序的属性名</param>
    /// <param name="asc">是否升序</param>
    public static IOrderedQueryable<T> OrderBy<T>(this IQueryable<T> source, string propertyName, bool asc = true)
    {
        string methodName = asc ? "OrderBy" : "OrderByDescending";
        return ApplyOrder(source, propertyName, methodName);
    }

    /// <summary>
    /// 根据属性名添加排序条件
    /// </summary>
    /// <param name="source">IOrderedQueryable</param>
    /// <param name="propertyName">要排序的属性名</param>
    /// <param name="asc">是否升序</param>
    public static IOrderedQueryable<T> ThenBy<T>(this IOrderedQueryable<T> source, string propertyName, bool asc = true)
    {
        string methodName = asc ? "ThenBy" : "ThenByDescending";
        return ApplyOrder(source, propertyName, methodName);
    }

    private static IOrderedQueryable<T> ApplyOrder<T>(IQueryable<T> source, string propertyName, string methodName)
    {
        if (string.IsNullOrEmpty(propertyName))
        {
            throw new ArgumentException("属性名必须有值", nameof(propertyName));
        }
        Type entityType = typeof(T);
        // 转成首字母大写
        propertyName = propertyName.ToUpperFirst();
        var propertyInfo = entityType.GetProperty(propertyName);
        if (null == propertyInfo)
        {
            throw new ArgumentException("属性不存在：" + propertyName);
        }
        ParameterExpression parameterExp = Expression.Parameter(entityType, string.Empty);
        MemberExpression propertyExp = Expression.Property(parameterExp, propertyInfo);

        var lambdaExp = Expression.Lambda(propertyExp, parameterExp);
        var methodCall = Expression.Call(typeof(Queryable), methodName,
                                        [entityType, propertyInfo.PropertyType],
                                        source.Expression, lambdaExp);
        return (IOrderedQueryable<T>)source.Provider.CreateQuery<T>(methodCall);
    }
}