using NPOI.SS.UserModel;

namespace FeiNuo.Core;

/// <summary>
/// 样式工厂类，根据样式配置ExcelStyle生成新样式，如果已存在样式，则直接返回已存在的样式
/// </summary>
public class StyleFactory
{
    private readonly IWorkbook workbook;
    private readonly Dictionary<string, ICellStyle> CACHED_STYLES = [];

    public StyleFactory(IWorkbook workbook, ExcelStyle? defaultStyle = null)
    {
        this.workbook = workbook;
        // 空样式
        EmptyStyle = workbook.CreateCellStyle();
        // 默认格式
        DefaultStyle = PoiHelper.CreateCellStyle(workbook, defaultStyle ?? new ExcelStyle());
    }

    /// <summary>
    /// 根据配置内容生成单元格样式
    /// <para>注意不同的格式需要重新调用该方法生成，不能直接修改原格式，否则会影响其他内容的格式</para>
    /// <param name="config">样式配置</param>
    /// <param name="source">样式以source为模板，默认为DefaultStyle，如果想完全不要模板，可设置为EmptyStyle</param>
    /// </summary>
    public ICellStyle CreateStyle(ExcelStyle config, ICellStyle? source = null)
    {
        var key = config.StyleKey;
        if (!CACHED_STYLES.TryGetValue(key, out var style))
        {
            style = PoiHelper.CreateCellStyle(workbook, config, source ?? DefaultStyle);
            CACHED_STYLES.Add(key, style);
        }
        return style;
    }

    public ICellStyle GetStyle(ExcelStyle config) => CreateStyle(config);

    /// <summary>
    /// 新建style,不从缓存中取。不要在循环中调用该方法，会产生较多样式，影响性能
    /// </summary>
    public ICellStyle NewStyle(ExcelStyle config, ICellStyle? source = null)
    {
        return PoiHelper.CreateCellStyle(workbook, config, source ?? DefaultStyle);
    }

    #region 预定义常用的样式
    /// <summary>
    /// 空样式
    /// </summary>
    public ICellStyle EmptyStyle { get; private set; }

    /// <summary>
    /// 默认样式
    /// </summary>
    public ICellStyle DefaultStyle { get; private set; }

    /// <summary>
    /// 文本格式：水平居左，垂直居中，格式 @
    /// </summary>
    public ICellStyle TextStyle { get { return CreateStyle(ExcelStyle.NewStyle().Format("@").Align(1, 1)); } }

    /// <summary>
    /// 日期格式：水平居中，垂直居中，格式 yyyy-MM-dd
    /// </summary>
    public ICellStyle DateStyle { get { return CreateStyle(ExcelStyle.NewStyle().Format("yyyy-mm-dd").Align(2, 1)); } }

    /// <summary>
    /// 时间格式：水平居中，垂直居中，格式 yyyy-MM-dd HH:mm
    /// </summary>
    public ICellStyle DateTimeStyle { get { return CreateStyle(new ExcelStyle().Format("yyyy-mm-dd hh:mm").Align(2, 1)); } }

    /// <summary>
    /// 数字格式：水平居中，垂直居中，格式 0.00
    /// </summary>
    public ICellStyle NumberStyle { get { return CreateStyle(new ExcelStyle().Format("0.00").Align(2, 1)); } }

    /// <summary>
    /// 百分比：水平居中，垂直居中，格式 0.00%
    /// </summary>
    public ICellStyle PersentStyle { get { return CreateStyle(new ExcelStyle().Format("0.00%").Align(2, 1)); } }

    /// <summary>
    /// 自动换行样式: 水平居左，垂直居中
    /// </summary>
    public ICellStyle WrapStyle { get { return CreateStyle(new ExcelStyle().Wrap().Align(1, 1)); } }

    #endregion
}
