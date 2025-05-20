namespace FeiNuo.Core;

/// <summary>
/// 根据POI的样式做一层转接，用于根据样式做缓存，避免重复创建样式
/// </summary>
public class ExcelStyle
{
    /// <summary>
    /// 边框:POI.BorderStyle 0:None, 1:Thin, 2:Medium, 3:Dashed, 4:Dotted, 5:Thick, 6:Double, 7:Hair
    /// </summary>
    public int? BorderStyle { get; private set; }

    /// <summary>
    /// 水平位置: POI.HorizontalAlignment 1:居左, 2:居中, 3:居右
    /// </summary>
    public int? HorizontalAlignment { get; private set; }

    /// <summary>
    /// 垂直位置: POI.VerticalAlignment  0:居上, 1:居中, 2:居下
    /// </summary>
    public int? VerticalAlignment { get; private set; }

    /// <summary>
    /// 背景颜色:POI.HSSFColor
    /// </summary>
    public short? BackgroundColor { get; private set; }

    /// <summary>
    /// 格式化字符串
    /// </summary>
    public string? DataFormat { get; private set; }

    /// <summary>
    /// 是否换行
    /// </summary>
    public bool? WrapText { get; private set; }

    /// <summary>
    /// 字体名称
    /// </summary>
    public string? FontName { get; private set; }

    /// <summary>
    /// 字体颜色
    /// </summary>
    public short? FontColor { get; private set; }

    /// <summary>
    /// 字体大小
    /// </summary>
    public short? FontSize { get; private set; }

    /// <summary>
    /// 是否粗体
    /// </summary>
    public bool? FontBold { get; private set; }

    #region 链式调用方法
    /// <summary>
    /// 创建新的样式
    /// </summary>
    public static ExcelStyle NewStyle()
    {
        return new ExcelStyle();
    }

    /// <summary>
    /// 边框: 0:None, 1:Thin, 2:Medium, 3:Dashed, 4:Dotted, 5:Thick, 6:Double, 7:Hair
    /// </summary>
    public ExcelStyle Border(int border = 1)
    {
        BorderStyle = border;
        return this;
    }

    /// <summary>
    /// 水平位置:  1:居左, 2:居中, 3:居右
    /// </summary>
    public ExcelStyle HAlign(int hAlign = 2)
    {
        HorizontalAlignment = hAlign;
        return this;
    }

    /// <summary>
    /// 垂直位置:  0:居上, 1:居中, 2:居下
    /// </summary>
    public ExcelStyle VAlign(int vAlign = 1)
    {
        VerticalAlignment = vAlign;
        return this;
    }

    /// <summary>
    /// 水平位置:  1:居左, 2:居中, 3:居右; 垂直位置:  0:居上, 1:居中, 2:居下
    /// </summary>
    public ExcelStyle Align(int hAlign = 2, int vAlign = 1)
    {
        HorizontalAlignment = hAlign;
        VerticalAlignment = vAlign;
        return this;
    }

    /// <summary>
    /// 格式化字符串
    /// </summary>
    public ExcelStyle Format(string format)
    {
        DataFormat = format;
        return this;
    }

    /// <summary>
    /// 是否换行
    /// </summary>
    public ExcelStyle Wrap(bool wrap = true)
    {
        WrapText = wrap;
        return this;
    }

    /// <summary>
    /// 背景颜色
    /// </summary>
    public ExcelStyle BgColor(short bgColor)
    {
        BackgroundColor = bgColor;
        return this;
    }

    /// <summary>
    /// 字体名称
    /// </summary>
    public ExcelStyle FontN(string fontName)
    {
        FontName = fontName;
        return this;
    }

    /// <summary>
    /// 字体颜色
    /// </summary>
    public ExcelStyle FontC(short fontColor)
    {
        FontColor = fontColor;
        return this;
    }

    /// <summary>
    /// 字体大小
    /// </summary>
    public ExcelStyle FontS(short fontSize)
    {
        FontSize = fontSize;
        return this;
    }

    /// <summary>
    /// 是否粗体
    /// </summary>
    public ExcelStyle Bold(bool bold = true)
    {
        FontBold = bold;
        return this;
    }

    /// <summary>
    /// 字体名称、大小、颜色、粗体
    /// </summary>
    public ExcelStyle Font(string? fontName, short? fontSize, short? fontColor, bool? bold)
    {
        FontSize = fontSize;
        return this;
    }

    #region 快捷方法
    /// <summary>
    /// 空样式，没有任何配置
    /// </summary>
    public static ExcelStyle EmptyStyle => new();

    /// <summary>
    /// 边框样式：默认四边带Thin边框
    /// </summary>
    public static ExcelStyle BorderedStyle => NewStyle().Border();

    /// <summary>
    /// 文本格式：水平居左，垂直居中，格式 @
    /// </summary>
    public static ExcelStyle TextStyle => NewStyle().Format("@");

    /// <summary>
    /// 自动换行: 文本格式，水平居左，垂直居中
    /// </summary>
    public static ExcelStyle WrapStyle => NewStyle().Format("@").Wrap();

    /// <summary>
    /// 居中样式: 水平居中，垂直居中
    /// </summary>
    public static ExcelStyle CenterStyle => NewStyle().Align(2, 1);

    /// <summary>
    /// 设置日期格式：format=yyyy-MM-dd，居中
    /// </summary>
    /// <returns></returns>
    public static ExcelStyle DateStye => NewStyle().Format("yyyy-mm-dd").Align(2, 1);

    /// <summary>
    /// 时间格式：水平居中，垂直居中，格式 yyyy-MM-dd HH:mm
    /// </summary>
    public static ExcelStyle DateTimeStyle => NewStyle().Format("yyyy-mm-dd hh:mm").Align(2, 1);

    /// <summary>
    /// 数字格式：水平居中，垂直居中，格式 0.00
    /// </summary>
    public static ExcelStyle NumberStyle => NewStyle().Format("0.00").Align(2, 1);

    /// <summary>
    /// 百分比：水平居中，垂直居中，格式 0.00%
    /// </summary>
    public static ExcelStyle PersentStyle => NewStyle().Format("0.00%").Align(2, 1);
    #endregion

    #endregion

    /// <summary>
    /// 根据各项配置值生成唯一的键值
    /// </summary>
    public string StyleKey
    {
        get
        {
            return $"BD{BorderStyle ?? -1},HA{HorizontalAlignment ?? -1},VA{VerticalAlignment ?? -1}"
            + $"BG{BackgroundColor ?? -1},DF{DataFormat},WP{WrapText ?? false}"
            + $"FN{FontName ?? ""},FC{FontColor ?? -1},FS{FontSize ?? -1},FB{FontBold ?? false}";
        }
    }

    /// <summary>
    /// 是否有设置样式，非默认样式
    /// </summary>
    public bool IsNotEmptyStyle
    {
        get
        {
            return BorderStyle.HasValue || HorizontalAlignment.HasValue || VerticalAlignment.HasValue
                || BackgroundColor.HasValue || (!string.IsNullOrWhiteSpace(DataFormat)) || WrapText.HasValue
                || (!string.IsNullOrWhiteSpace(FontName)) || FontColor.HasValue || FontSize.HasValue || FontBold.HasValue;
        }
    }

    /// <summary>
    /// 是否默认的样式，没有设置任何格式
    /// </summary>
    public bool IsEmptyStyle => !IsNotEmptyStyle;
}
