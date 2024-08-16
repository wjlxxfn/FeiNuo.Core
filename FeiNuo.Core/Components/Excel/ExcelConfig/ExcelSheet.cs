namespace FeiNuo.Core
{
    /// <summary>
    /// 工作表配置
    /// </summary>
    public class ExcelSheet
    {
        #region 构造函数
        public ExcelSheet(string sheetName, IEnumerable<ExcelColumn>? columns = null)
        {
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                throw new ArgumentNullException(nameof(sheetName));
            }
            SheetName = sheetName;
            ExcelColumns = columns ?? [];
        }

        public ExcelSheet(string sheetName, IEnumerable<object> dataList, IEnumerable<ExcelColumn> columns) : this(sheetName, columns)
        {
            DataList = dataList ?? [];
        }
        #endregion

        #region 核心字段：SheetName,ExcelColumns，DataList
        /// <summary>
        /// 工作表名
        /// </summary>
        public string SheetName { get; set; }

        /// <summary>
        /// 列配置
        /// </summary>
        public IEnumerable<ExcelColumn> ExcelColumns { get; set; }

        /// <summary>
        /// 数据集合
        /// </summary>
        public IEnumerable<object> DataList { get; set; } = [];
        #endregion

        #region 数据列配置,标题样式
        /// <summary>
        /// 列标题样式：水平居中，字体加粗，加背景色
        /// </summary>
        public ExcelStyle ColumnTitleStyle = new() { HorizontalAlignment = 2, FontBold = true, BackgroundColor = 26 };

        /// <summary>
        /// 在上传Excel数据时，是否效验模板：根据sheet名，标题名必须一致才能继续导入
        /// </summary>
        public bool ValidateImportTemplate { get; set; } = true;
        #endregion

        #region 工作表整体配置：默认列宽，自动计算公式，自动添加边框
        /// <summary>
        /// 默认列宽
        /// </summary>
        public int? DefaultColumnWidth { get; set; }

        /// <summary>
        /// 默认计算公式的值
        /// </summary>
        public bool ForceFormulaRecalculation { get; set; } = true;

        /// <summary>
        /// 使用条件格式对所有数据区别添加边框：边框为BorderStyle.Thin
        /// </summary>
        public bool AddConditionalBorderStyle { get; set; } = true;
        #endregion

        #region 整体说明行，在最前面，可配置样式，行高，合并列的个数
        /// <summary>
        /// 导入说明：如果有添加到第一行, 在主标题的前面
        /// </summary>
        public string Description { get; set; } = string.Empty;

        /// <summary>
        /// 说明的样式：水平居左，自动换行
        /// </summary>
        public ExcelStyle DescriptionStyle = new() { HorizontalAlignment = 1, WrapText = true, };

        /// <summary>
        /// 说明行的高度:默认66
        /// </summary>
        public int DescriptionRowHeight { get; set; } = 66;

        /// <summary>
        /// 说明行合并单元格的数量，默认为列的数量，可通过该参数调整
        /// </summary>
        public int? DescriptionColSpan { get; set; }
        #endregion

        #region 主标题行，在说明行的下面，可配置样式，合并列的个数
        /// <summary>
        /// 主标题，默认在第一行
        /// </summary>
        public string MainTitle { get; set; } = string.Empty;

        /// <summary>
        /// 主标题样式：水平居中，字体加粗，加背景色
        /// </summary>
        public ExcelStyle MainTitleStyle = new() { HorizontalAlignment = 2, FontBold = true, BackgroundColor = 26, };

        /// <summary>
        /// 主标题合并单元格的数量，默认为列的数量，可通过该参数调整
        /// </summary>
        public int? MainTitleColSpan { get; set; }
        #endregion

        #region 其他辅助方法
        /// <summary>
        /// 标题行行号
        /// </summary>
        public int TitleRowIndex
        {
            get
            {
                return (!string.IsNullOrWhiteSpace(Description) ? 1 : 0)
                    + (!string.IsNullOrWhiteSpace(MainTitle) ? 1 : 0);
            }
        }

        /// <summary>
        /// 内容行起始行号
        /// </summary>
        public int DataRowIndex
        {
            get { return TitleRowIndex + ExcelColumns.Max(t => t.RowTitles.Length); }
        }
        #endregion
    }
}
