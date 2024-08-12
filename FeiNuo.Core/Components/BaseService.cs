using Microsoft.EntityFrameworkCore;

namespace FeiNuo.Core
{
    /// <summary>
    /// 服务类顶层接口，实现该接口可自动注入
    /// </summary>
    public interface IService { }

    /// <summary>
    /// 服务类基类
    /// </summary>
    public abstract class BaseService : IService { }

    /// <summary>
    /// 服务类基类：指定数据库上下文和实体类型，提供基础操作
    /// </summary>
    public abstract class BaseService<T> : BaseService where T : BaseEntity
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
        protected async Task<T> FindByIdAsync(object id)
        {
            var obj = await ctx.FindAsync<T>(id)
                ?? throw new NotFoundException($"找不到指定数据,Id:{id},Type:{typeof(T)}");
            return obj;
        }
        #endregion

        #region 其它操作
        protected void AppendOperator(BaseEntity entity, string user, bool isNew)
        {
            if (isNew)
            {
                entity.CreatedBy = user;
                entity.CreatedTime = DateTime.Now;
            }
            entity.UpdatedBy = user;
            entity.UpdatedTime = DateTime.Now;
        }
        #endregion
    }
}