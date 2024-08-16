using Microsoft.Extensions.Primitives;

namespace FeiNuo.Core
{
    /// <summary>
    /// 通用查询类，不需要定义查询类。
    /// 使用QueryHelpers.ParseQuery把Request.Query转成Dictionary,构建CommonQuery。
    /// 然后在Service中通过CommonQuery手动拼接查询
    /// 适用于条件不多的简单查询
    /// </summary>
    public class CommonQuery
    {
        private readonly Dictionary<string, StringValues> query;
        /// <summary>
        /// 通过Request.Query.ToDictionary()初始化查询条件
        /// </summary>
        public CommonQuery(Dictionary<string, StringValues> query)
        {
            if (query == null) this.query = [];
            else this.query = query;
        }

        #region 根据KEY检查是否有查询条件
        /// <summary>
        /// 检查是否有查询条件，如果有返回StringValues条件，然后自己根据实体类型做处理
        /// </summary>
        public bool HasQuery(string key, out StringValues val)
        {
            return query.TryGetValue(key, out val);
        }

        /// <summary>
        /// 检查是否有指定字符串查询条件，如果有返回trim后的条件
        /// </summary>
        public bool HasStringQuery(string key, out string val)
        {
            val = string.Empty;
            if (query.TryGetValue(key, out var value))
            {
                val = value.ToString().Trim();
                return !string.IsNullOrWhiteSpace(val);
            }
            return false;
        }

        /// <summary>
        /// 检查是否有字符串数组查询条件，如果有返回字符串数组
        /// </summary>
        public bool HasArrayQuery(string key, out string[] val)
        {
            val = [];
            if (query.TryGetValue(key, out var value))
            {
                val = value.ToArray()!;
                return val.Length > 0;
            }
            return false;
        }

        /// <summary>
        /// 检查是否有整形数组查询条件，如果有返回整形数组
        /// </summary>
        public bool HasArrayQuery(string key, out int[] val)
        {

            val = [];
            bool hasQuery = HasArrayQuery(key, out string[] strArray);
            if (hasQuery) val = strArray.Select(int.Parse).ToArray();
            return hasQuery;
        }

        /// <summary>
        /// 检查是否有整形查询条件，如果有返回整形条件
        /// </summary>
        public bool HasIntQuery(string key, out int val)
        {
            val = 0;
            return query.TryGetValue(key, out var v) && int.TryParse(v.ToString(), out val);
        }
        #endregion
    }
}
