namespace FeiNuo.Core;

/// <summary>
/// 日期常用方法
/// </summary>
public static class DateTimeExtensions
{
    /// <summary>
    /// 转成日期字符串格式：yyyy-MM-dd ，
    /// 如果是DateTime.MinValue,则返回空字符串
    /// 如果是1900-1-1,则返回空字符串,解决SqlServer里的00值问题
    /// </summary>
    /// <param name="dt">日期</param>
    /// <returns>yyyy-MM-dd</returns>
    public static string ToDateString(this DateTime dt)
    {
        if (dt == DateTime.MinValue) return "";
        var str = dt.ToString("yyyy-MM-dd");
        return str == "1900-01-01" ? "" : str;
    }

    /// <summary>
    /// 转成日期字符串格式：yyyy-MM-dd HH:mm
    /// 如果是DateTime.MinValue,则返回空字符串
    /// 如果是1900-1-1,则返回空字符串,解决SqlServer里的00值问题
    /// </summary>
    /// <param name="dt">日期</param>
    /// <returns>yyyy-MM-dd HH:mm</returns>
    public static string ToDateTimeString(this DateTime dt)
    {
        if (dt == DateTime.MinValue) return "";
        var str = dt.ToString("yyyy-MM-dd HH:mm");
        return str.StartsWith("1900-01-01") ? "" : str;
    }

    /// <summary>
    /// 时间类型转成unix时间戳：秒数
    /// </summary>
    public static long ToUnixTimeSeconds(this DateTime dt)
    {
        return new DateTimeOffset(dt).ToUnixTimeSeconds();
    }

    /// <summary>
    /// 时间类型转成unix时间戳：豪秒数
    /// </summary>
    public static long ToUnixTimeMilliseconds(this DateTime dt)
    {
        return new DateTimeOffset(dt).ToUnixTimeMilliseconds();
    }

}
