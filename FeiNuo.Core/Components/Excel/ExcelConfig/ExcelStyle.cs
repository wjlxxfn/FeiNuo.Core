namespace FeiNuo.Core
{
    /// <summary>
    /// 根据POI的样式做一层转接，用于根据样式做缓存，避免重复创建样式
    /// </summary>
    public class ExcelStyle
    {
        /// <summary>
        /// 边框
        /// </summary>
        public int? BorderStyle { get; set; }

        /// <summary>
        /// 水平位置
        /// </summary>
        public int? HorizontalAlignment { get; set; }

        /// <summary>
        /// 垂直位置
        /// </summary>
        public int? VerticalAlignment { get; set; }

        /// <summary>
        /// 背景颜色
        /// </summary>
        public short? BackgroundColor { get; set; }

        /// <summary>
        /// 格式化字符串
        /// </summary>
        public string? DataFormat { get; set; }

        /// <summary>
        /// 是否换行
        /// </summary>
        public bool? WrapText { get; set; }

        /// <summary>
        /// 字体名称
        /// </summary>
        public string? FontName { get; set; }

        /// <summary>
        /// 字体颜色
        /// </summary>
        public short? FontColor { get; set; }

        /// <summary>
        /// 字体大小
        /// </summary>
        public short? FontSize { get; set; }

        /// <summary>
        /// 是否粗体
        /// </summary>
        public bool? FontBold { get; set; }

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
        /// 是否初始值，没设置任何格式
        /// </summary>
        public bool IsEmptyStyle
        {
            get
            {
                return !(
                    BorderStyle.HasValue || HorizontalAlignment.HasValue || VerticalAlignment.HasValue
                    || BackgroundColor.HasValue || (!string.IsNullOrWhiteSpace(DataFormat)) || WrapText.HasValue
                    || (!string.IsNullOrWhiteSpace(FontName)) || FontColor.HasValue || FontSize.HasValue || FontBold.HasValue
                    );
            }
        }
    }
}
