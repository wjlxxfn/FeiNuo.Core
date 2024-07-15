using System.Linq.Expressions;

namespace FeiNuo.Core
{

    /// <summary>
    /// 查询扩展：添加通过字符串排序的方法
    /// </summary>
    public static class QueryableExtensions
    {
        /// <summary>
        /// 根据属性名字符串添加排序条件:升序
        /// </summary>
        /// <param name="source">要排序的数据集合</param>
        /// <param name="propertyName">要排序的属性名</param>
        public static IEnumerable<TSource> OrderBy<TSource>(this IEnumerable<TSource> source, string propertyName)
        {
            return ApplyOrder(source.AsQueryable(), propertyName, "OrderBy");
        }

        /// <summary>
        /// 根据属性名字符串添加排序条件:降序
        /// </summary>
        /// <param name="source">要排序的数据集合</param>
        /// <param name="propertyName">要排序的属性名</param>
        public static IEnumerable<TSource> OrderByDescending<TSource>(this IEnumerable<TSource> source, string propertyName)
        {
            return ApplyOrder(source.AsQueryable(), propertyName, "OrderByDescending");
        }

        /// <summary>
        /// 根据属性名添加排序条件：升序
        /// </summary>
        /// <param name="source">要排序的数据集合</param>
        /// <param name="propertyName">要排序的属性名</param>
        public static IEnumerable<TSource> ThenBy<TSource>(this IEnumerable<TSource> source, string propertyName)
        {
            return ApplyOrder(source.AsQueryable(), propertyName, "ThenBy");
        }

        /// <summary>
        /// 根据属性名字符串添加排序条件:降序
        /// </summary>
        /// <param name="source">要排序的数据集合</param>
        /// <param name="propertyName">要排序的属性名</param>
        public static IEnumerable<TSource> ThenByDescending<TSource>(this IEnumerable<TSource> source, string propertyName)
        {
            return ApplyOrder(source.AsQueryable(), propertyName, "ThenByDescending");
        }

        private static IEnumerable<TSource> ApplyOrder<TSource>(this IQueryable<TSource> source, string propertyName, string methodName)
        {
            if (string.IsNullOrEmpty(propertyName))
            {
                throw new ArgumentException("属性名必须有值", nameof(propertyName));
            }
            Type entityType = typeof(TSource);
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
            return source.Provider.CreateQuery<TSource>(methodCall);
        }
    }
}