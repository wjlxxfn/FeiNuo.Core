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
        // 创建默认格式
        DefaultStyle = PoiHelper.CreateCellStyle(workbook, defaultStyle ?? new ExcelStyle());
    }

    /// <summary>
    /// 默认样式
    /// </summary>
    public ICellStyle DefaultStyle { get; private set; }

    /// <summary>
    /// 根本配置内容生成格式
    /// 注意不同的格式需要重新调用该方法生成，不能直接修改原格式，否则会影响其他内容的格式
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
    public ICellStyle CreateNewStyle(ExcelStyle config)
    {
        return PoiHelper.CreateCellStyle(workbook, config, DefaultStyle);
    }

    #region 预定义常用的样式
    /// <summary>
    /// 文本格式：居左，格式 @
    /// </summary>
    public ICellStyle TextStyle { get { return CreateStyle(new() { DataFormat = "@", HorizontalAlignment = (int)HorizontalAlignment.Left }); } }

    /// <summary>
    /// 日期格式：居中，格式 yyyy-MM-dd
    /// </summary>
    public ICellStyle DateStyle { get { return CreateStyle(new() { DataFormat = "yyyy-mm-dd", HorizontalAlignment = (int)HorizontalAlignment.Center }); } }

    /// <summary>
    /// 时间格式：居中，格式 yyyy-MM-dd HH:mm
    /// </summary>
    public ICellStyle DateTimeStyle { get { return CreateStyle(new() { DataFormat = "yyyy-mm-dd hh:mm", HorizontalAlignment = (int)HorizontalAlignment.Center }); } }

    /// <summary>
    /// 数字格式：居中，格式 0.00
    /// </summary>
    public ICellStyle NumbericStyle { get { return CreateStyle(new() { DataFormat = "0.00", HorizontalAlignment = (int)HorizontalAlignment.Right }); } }

    /// <summary>
    /// 百分比：居中，格式 0.00%
    /// </summary>
    public ICellStyle PersentStyle { get { return CreateStyle(new() { DataFormat = "0.00%", HorizontalAlignment = (int)HorizontalAlignment.Right }); } }

    /// <summary>
    /// 自动换行样式: 居左，垂直居中
    /// </summary>
    public ICellStyle WrapStyle { get { return CreateStyle(new() { HorizontalAlignment = (int)HorizontalAlignment.Left, VerticalAlignment = ((int)VerticalAlignment.Center), WrapText = true }); } }

    #endregion
}
