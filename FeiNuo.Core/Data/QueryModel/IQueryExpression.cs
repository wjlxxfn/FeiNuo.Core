using System.Linq.Expressions;

namespace FeiNuo.Core
{
    /// <summary>
    /// 查询接口，返回查询表达式
    /// </summary>
    /// <typeparam name="T">对应的实体类型</typeparam>
    public interface IQueryExpression<T> where T : class
    {
        /// <summary>
        /// 返回拼接好的查询表达式
        /// </summary>
        /// <returns>对应的实体的查询表达式</returns>
        Expression<Func<T, bool>> GetQueryExpression();
    }
}