using System.Data;

namespace FeiNuo.Core;

/// <summary>
/// 工作表配置
/// </summary>
public class ExcelSheet
{
    #region 属性定义
    /// <summary>
    /// 工作表名
    /// </summary>
    public string SheetName { get; set; }

    /// <summary>
    /// 列配置
    /// </summary>
    public IEnumerable<ExcelColumn> ExcelColumns { get; set; } = [];

    public IEnumerable<object> DataList = [];

    public ExcelSheet(IEnumerable<ExcelColumn> columns) : this("Sheet1", columns)
    {

    }

    public ExcelSheet(string sheetName, IEnumerable<ExcelColumn> columns)
    {
        if (string.IsNullOrWhiteSpace(sheetName))
        {
            throw new ArgumentNullException(nameof(sheetName));
        }
        SheetName = sheetName;

        ExcelColumns = ConfigColumns([.. columns]);
    }

    private List<ExcelColumn> ConfigColumns(List<ExcelColumn> columns)
    {
        if (columns.Count == 0) throw new Exception("列配置不存在");
        // 检查是否有相同的标题
        var chk = columns.GroupBy(k => k.Title).Where(g => g.Count() > 1).Select(g => g.Key).ToArray();
        if (chk.Length > 0)
        {
            throw new Exception($"【{SheetName}】以下列字段标题重复: {string.Join(", ", chk)}");
        }
        var colIndex = StartColumnIndex;
        foreach (var col in columns)
        {
            col.ColumnIndex = colIndex++;
        }
        return columns;
    }
    #endregion

    #region 核心字段
    /// <summary>
    /// 数据开始的列索引
    /// </summary>
    public int StartColumnIndex { get; private set; } = 0;
    /// <summary>
    /// 最后一列的列索引
    /// </summary>
    public int EndColumnIndex => StartColumnIndex + ColumnCount - 1;
    /// <summary>
    /// 列数
    /// </summary>
    public int ColumnCount => ExcelColumns.Count();
    /// <summary>
    /// 列标题行数
    /// </summary>
    public int ColumnRowCount => ExcelColumns.Max(t => t.RowTitles.Length);

    /// <summary>
    /// 设置数据开始的列索引
    /// </summary>
    /// <param name="colIndex"></param>
    public void SetStartColumnIndex(int colIndex)
    {
        StartColumnIndex = colIndex;
        foreach (var col in ExcelColumns)
        {
            col.ColumnIndex = colIndex++;
        }
    }

    /// <summary>
    /// 设置数据开始的行索引
    /// </summary>
    public int StartRowIndex { get; private set; } = 0;

    /// <summary>
    /// 设置数据开始的行索引
    /// </summary>
    public void SetStartRowIndex(int rowIndex)
    {
        StartRowIndex = rowIndex;
    }
    #endregion

    #region 数据列配置,标题样式
    /// <summary>
    /// 列标题样式：水平居中，字体加粗，加背景色
    /// </summary>
    public ExcelStyle? ColumnTitleStyle { get; set; }

    /// <summary>
    /// 在上传Excel数据时，是否效验模板：根据sheet名，标题名必须一致才能继续导入
    /// </summary>
    public bool ValidateImportTemplate { get; set; } = true;
    #endregion

    #region 工作表整体配置：默认列宽，自动计算公式，自动添加边框
    /// <summary>
    /// 默认列宽
    /// </summary>
    public int? DefaultColumnWidth { get; set; } = 12;

    /// <summary>
    /// 默认计算公式的值
    /// </summary>
    public bool ForceFormulaRecalculation { get; set; } = true;

    /// <summary>
    /// 使用条件格式对所有数据区别添加边框：边框为BorderStyle.Thin
    /// </summary>
    public bool AddConditionalBorderStyle { get; set; } = false;
    #endregion

    #region 整体说明行，在最前面，可配置样式，行高，合并列的个数
    /// <summary>
    /// 导入说明：如果有添加到第一行, 在主标题的前面
    /// </summary>
    public string Description { get; set; } = string.Empty;

    /// <summary>
    /// 说明的样式：水平居左，自动换行
    /// </summary>
    public ExcelStyle? DescriptionStyle { get; set; }

    /// <summary>
    /// 说明行的高度:默认66
    /// </summary>
    public int DescriptionRowHeight { get; set; } = 66;

    /// <summary>
    /// 说明行合并单元格的数量，默认为列的数量，可通过该参数调整
    /// </summary>
    public int? DescriptionColSpan { get; set; }

    /// <summary>
    /// 添加说明行
    /// </summary>
    public ExcelSheet AddDescription(string description, int? colSpan = null, int? rowHeight = null, ExcelStyle? style = null)
    {
        Description = description;
        DescriptionColSpan = colSpan;
        DescriptionRowHeight = rowHeight ?? DescriptionRowHeight;
        DescriptionStyle = style;
        return this;
    }
    #endregion

    #region 主标题行，在说明行的下面，可配置样式，合并列的个数
    /// <summary>
    /// 主标题，默认在第一行
    /// </summary>
    public string MainTitle { get; set; } = string.Empty;

    /// <summary>
    /// 主标题样式：水平居中，字体加粗，加背景色
    /// </summary>
    public ExcelStyle? MainTitleStyle { get; set; }

    /// <summary>
    /// 主标题合并单元格的数量，默认为列的数量，可通过该参数调整
    /// </summary>
    public int? MainTitleColSpan { get; set; }

    /// <summary>
    /// 添加主标题
    /// </summary>
    public ExcelSheet AddMainTitle(string mainTitle, int? colSpan = null, ExcelStyle? style = null)
    {
        MainTitle = mainTitle;
        MainTitleColSpan = colSpan;
        MainTitleStyle = style;
        return this;
    }

    public ExcelSheet AddDataList(IEnumerable<object> dataList)
    {
        DataList = dataList;
        return this;
    }
    #endregion

    #region 其他辅助方法
    /// <summary>
    /// 标题行行号
    /// </summary>
    public int TitleRowIndex
    {
        get
        {
            return StartRowIndex + (!string.IsNullOrWhiteSpace(Description) ? 1 : 0)
                + (!string.IsNullOrWhiteSpace(MainTitle) ? 1 : 0);
        }
    }

    /// <summary>
    /// 内容行起始行号
    /// </summary>
    public int DataRowIndex
    {
        get { return TitleRowIndex + ColumnRowCount; }
    }
    #endregion
}
