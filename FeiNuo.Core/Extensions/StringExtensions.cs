namespace FeiNuo.Core
{
    /// <summary>
    /// 字符串扩展方法
    /// </summary>
    public static class StringExtensions
    {
        #region 首字母大小写
        /// <summary>
        /// 首字母大写
        /// </summary>
        public static string ToUpperFirst(this string str)
        {
            if (string.IsNullOrWhiteSpace(str))
            {
                return str;
            }
            return char.ToUpper(str[0]) + str[1..];
        }

        /// <summary>
        /// 首字母小写
        /// </summary>
        public static string ToLowerFirst(this string str)
        {
            if (string.IsNullOrWhiteSpace(str))
            {
                return str;
            }
            return char.ToLower(str[0]) + str[1..];
        }
        #endregion
    }
}