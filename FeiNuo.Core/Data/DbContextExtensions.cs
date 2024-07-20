using Dapper;
using Microsoft.EntityFrameworkCore;
using System.Data;

namespace FeiNuo.Core
{

    /// <summary>
    /// 使用dapper扩展原生查询功能
    /// </summary>
    public static class DbContextExtensions
    {
        //TODO 待测试所有方法
        public static async Task<IEnumerable<T>> SqlQueryAsync<T>(this DbContext db, string sql, object? parameters = null)
        {
            SqlMapper.AddTypeHandler(new DateOnlyTypeHandler());
            return await db.Database.GetDbConnection().QueryAsync<T>(sql, parameters);
        }

        public static async Task<IEnumerable<dynamic>> SqlQueryAsync(this DbContext db, string sql, object? parameters = null)
        {
            return await db.Database.GetDbConnection().QueryAsync(sql, parameters);
        }

        public static async Task<IDataReader> SqlQueryReaderAsync(this DbContext db, string sql, object? parameters = null)
        {
            return await db.Database.GetDbConnection().ExecuteReaderAsync(sql, parameters);
        }

        public static async Task<T?> SqlQueryScalarAsync<T>(this DbContext db, string sql, object? parameters = null)
        {
            return await db.Database.GetDbConnection().ExecuteScalarAsync<T>(sql, parameters);
        }

        public static async Task<object?> SqlQueryScalarAsync(this DbContext db, string sql, object? parameters = null)
        {
            return await db.Database.GetDbConnection().ExecuteScalarAsync(sql, parameters);
        }
    }

    /// <summary>
    /// DateOnly类型转换
    /// </summary>
    public class DateOnlyTypeHandler : SqlMapper.TypeHandler<DateOnly>
    {
        public override void SetValue(IDbDataParameter parameter, DateOnly date) => parameter.Value = date.ToDateTime(new TimeOnly(0, 0));

        public override DateOnly Parse(object value) => DateOnly.FromDateTime((DateTime)value);
    }
}
