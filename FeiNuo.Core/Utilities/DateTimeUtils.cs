namespace FeiNuo.Core.Utilities;

public class DateTimeUtils
{
    /// <summary>
    /// SQL Server中零日期值，主要用于日期字段的默认值。
    /// </summary>
    public static readonly DateTime ZeroSqlDateTime = new(1900, 1, 1);

    /// <summary>
    /// SQL Server中最大日期值，主要用于查询条件：结束日期
    /// </summary>
    public static readonly DateTime MaxSqlDateTime = new(9999, 1, 1);
}
