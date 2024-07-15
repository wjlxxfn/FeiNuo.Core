using System.Linq.Expressions;

namespace FeiNuo.Core
{
    /// <summary>
    /// Linq表达式扩展，添加常用的And/Or等方法
    /// </summary>
    public static class ExpressionUtils
    {
        /// <summary>
        /// 默认True值
        /// </summary>
        public static Expression<Func<T, bool>> True<T>() { return a => true; }

        /// <summary>
        /// 默认False值
        /// </summary>
        public static Expression<Func<T, bool>> False<T>() { return a => false; }

        /// <summary>
        /// And
        /// </summary>
        public static Expression<Func<T, bool>> And<T>(this Expression<Func<T, bool>> left, Expression<Func<T, bool>> right)
        {
            var invokedExpression = Expression.Invoke(right, left.Parameters.Cast<Expression>());
            return Expression.Lambda<Func<T, bool>>(Expression.AndAlso(left.Body, invokedExpression), left.Parameters);
        }

        /// <summary>
        /// Or
        /// </summary>
        public static Expression<Func<T, bool>> Or<T>(this Expression<Func<T, bool>> left, Expression<Func<T, bool>> right)
        {
            var invokedExpression = Expression.Invoke(right, left.Parameters.Cast<Expression>());
            return Expression.Lambda<Func<T, bool>>(Expression.OrElse(left.Body, invokedExpression), left.Parameters);
        }
    }
}