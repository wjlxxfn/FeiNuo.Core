namespace FeiNuo.Core;

/// <summary>
/// 根据POI的样式做一层转接，用于根据样式做缓存，避免重复创建样式
/// </summary>
public class ExcelStyle
{
    /// <summary>
    /// 边框: 0:None, 1:Thin, 2:Medium, 3:Dashed, 4:Dotted, 5:Thick, 6:Double, 7:Hair
    /// </summary>
    public int? BorderStyle { get; set; }

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
    public int? HorizontalAlignment { get; set; }

    /// <summary>
    /// 水平位置:  1:居左, 2:居中, 3:居右
    /// </summary>
    public ExcelStyle HAlign(int halign = 2)
    {
        HorizontalAlignment = halign;
        return this;
    }

    /// <summary>
    /// 垂直位置:  0:居上, 1:居中, 2:居下
    /// </summary>
    public int? VerticalAlignment { get; set; }

    /// <summary>
    /// 垂直位置:  0:居上, 1:居中, 2:居下
    /// </summary>
    public ExcelStyle VAlign(int valign = 1)
    {
        VerticalAlignment = valign;
        return this;
    }

    /// <summary>
    /// 背景颜色
    /// </summary>
    public short? BackgroundColor { get; set; }

    /// <summary>
    /// 背景颜色
    /// </summary>
    public ExcelStyle BgColor(short bgColor)
    {
        BackgroundColor = bgColor;
        return this;
    }

    /// <summary>
    /// 格式化字符串
    /// </summary>
    public string? DataFormat { get; set; }

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
    public bool? WrapText { get; set; }

    /// <summary>
    /// 是否换行
    /// </summary>
    public ExcelStyle Wrap(bool wrap = true)
    {
        WrapText = wrap;
        return this;
    }

    /// <summary>
    /// 字体名称
    /// </summary>
    public string? FontName { get; set; }

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
    public short? FontColor { get; set; }

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
    public short? FontSize { get; set; }

    /// <summary>
    /// 字体大小
    /// </summary>
    public ExcelStyle FontS(short fontSize)
    {
        FontSize = fontSize;
        return this;
    }

    public ExcelStyle Font(string? fontName, short? fontSize, short? fontColor, bool? bold)
    {
        FontSize = fontSize;
        return this;
    }


    /// <summary>
    /// 是否粗体
    /// </summary>
    public bool? FontBold { get; set; }

    /// <summary>
    /// 是否粗体
    /// </summary>
    public ExcelStyle Bold(bool bold = true)
    {
        FontBold = bold;
        return this;
    }

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
