using System.Linq.Expressions;

namespace FeiNuo.Core
{
    /// <summary>
    /// 普通查询类的基类：实现IQueryExpression，提供添加条件的公共方法
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public abstract class AbstractQuery<T> : BaseQuery, IQueryExpression<T> where T : BaseEntity
    {
        // 表达式集合
        private Expression<Func<T, bool>> _exp = ExpressionUtils.True<T>();

        /// <summary>
        /// 返回拼接好的查询表达式
        /// </summary>
        /// <returns>对应的实体的查询表达式</returns>
        public Expression<Func<T, bool>> GetQueryExpression()
        {
            // 业务类添加查询条件
            MergeQueryExpression();
            return _exp;
        }

        /// <summary>
        /// 添加各项查询条件表达式：调用AddExpression方法拼接表达式
        /// </summary>
        protected abstract void MergeQueryExpression();

        #region 辅助方法
        /// <summary>
        /// 添加查询表达式 AND 连接
        /// </summary>
        /// <param name="exp"></param>
        public AbstractQuery<T> AddExpression(Expression<Func<T, bool>> exp)
        {
            _exp = _exp.And(exp);
            return this;
        }

        /// <summary>
        /// 添加查询表达式
        /// </summary>
        protected void AddExpression(bool hasValue, Expression<Func<T, bool>> exp)
        {
            if (hasValue)
            {
                AddExpression(exp);
            }
        }

        /// <summary>
        /// 添加查询表达式
        /// </summary>
        protected void AddExpression(string? val, Expression<Func<T, bool>> exp)
        {
            AddExpression(!string.IsNullOrWhiteSpace(val), exp);
        }

        /// <summary>
        /// 添加模糊搜索条件
        /// </summary>
        protected void AddSearchExpression(Func<string, Expression<Func<T, bool>>> searchPredicate)
        {
            if (!string.IsNullOrWhiteSpace(Search))
            {
                AddExpression(searchPredicate(Search.Trim()));
            }
        }

        /// <summary>
        /// 添加日期范围查询条件
        /// </summary>
        protected void AddDateExpression(Func<DateTime, Expression<Func<T, bool>>> startDatePredicate, Func<DateTime, Expression<Func<T, bool>>> endDatePredicate)
        {
            if (StartDate.HasValue)
            {
                var start = StartDate.Value.ToDateTime(TimeOnly.MinValue);
                AddExpression(startDatePredicate(start));
            }
            if (EndDate.HasValue)
            {
                var end = EndDate.Value.ToDateTime(TimeOnly.MaxValue);
                AddExpression(endDatePredicate(end));
            }
        }

        /// <summary>
        /// 根据BaseEntity.CreatedTime和BaseQuery.StartDate、EndDate匹配
        /// </summary>
        protected void AddCreatedDateExpression()
        {
            if (StartDate.HasValue)
            {
                var start = StartDate.Value.ToDateTime(TimeOnly.MinValue);
                AddExpression(a => a.CreatedTime >= start);
            }
            if (EndDate.HasValue)
            {
                var end = EndDate.Value.ToDateTime(TimeOnly.MaxValue);
                AddExpression(a => a.CreatedTime <= end);
            }
        }
        #endregion
    }
}