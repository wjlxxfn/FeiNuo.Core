namespace FeiNuo.Core
{
    /// <summary>
    /// 工作表配置
    /// </summary>
    public class ExcelSheet
    {
        public ExcelSheet(string sheetName)
        {
            SheetName = sheetName;
        }

        public string SheetName { get; set; } = null!;

        /// <summary>
        /// 主标题，默认在第一行
        /// </summary>
        public string MainTitle { get; set; } = string.Empty;

        /// <summary>
        /// 主标题合并单元格的数量，默认为列的数量，可通过该参数调整
        /// </summary>
        public int? MainTitleMergedRegion { get; set; }

        /// <summary>
        /// 导入说明：如果有添加到第一行, 在主标题的前面
        /// </summary>
        public string Description { get; set; } = string.Empty;

        /// <summary>
        /// 说明行的高度:默认66
        /// </summary>
        public int DescriptionRowHeight { get; set; } = 66;

        /// <summary>
        /// 入说明合并单元格的数量，默认为列的数量，可通过该参数调整
        /// </summary>
        public int? DescriptionMergedRegion { get; set; }

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

        /// <summary>
        /// 列配置
        /// </summary>
        public List<ExcelColumn> ExcelColumns { get; set; } = [];

        /// <summary>
        /// 开始列索引,默认从0开始
        /// </summary>
        public int StartColIndex { get; set; } = 0;

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
    }
}
