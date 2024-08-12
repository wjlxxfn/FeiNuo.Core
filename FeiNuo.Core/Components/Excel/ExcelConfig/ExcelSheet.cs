namespace FeiNuo.Core
{
    /// <summary>
    /// 工作表配置
    /// </summary>
    public class ExcelSheet
    {
        public string SheetName { get; set; }
        public ExcelSheet(string sheetName)
        {
            SheetName = sheetName;
        }

        #region 数据列配置,标题样式
        /// <summary>
        /// 列配置
        /// </summary>
        public IEnumerable<ExcelColumn> ExcelColumns { get; set; } = [];

        /// <summary>
        /// 列标题样式：水平居中，字体加粗，加背景色
        /// </summary>
        public ExcelStyle ColumnTitleStyle = new() { HorizontalAlignment = 2, FontBold = true, BackgroundColor = 26 };
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

        /// <summary>
        /// 数据集合
        /// </summary>
        public IEnumerable<object> DataList { get; set; }

    }

    public class ExcelSheet<T> : ExcelSheet where T : class
    {
        /// <summary>
        /// 列配置
        /// </summary>
        public new IEnumerable<ExcelColumn<T>> ExcelColumns
        {
            get { return (IEnumerable<ExcelColumn<T>>)base.ExcelColumns; }
            set { base.ExcelColumns = value; }
        }

        /// <summary>
        /// 数据集合
        /// </summary>
        public new IEnumerable<T> DataList
        {
            get { return (IEnumerable<T>)base.DataList; }
            set { base.DataList = value; }
        }

        public ExcelSheet(string sheetName) : base(sheetName)
        {
        }

        public ExcelSheet(string sheetName, IEnumerable<T> dataList, IEnumerable<ExcelColumn<T>> columns) : base(sheetName)
        {
            DataList = dataList;
            ExcelColumns = columns;
        }
    }
}
