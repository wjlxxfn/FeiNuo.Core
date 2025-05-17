using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace FeiNuo.Core;

/// <summary>
/// 封装Workbook和Sheet对象，提供常用的操作方法
/// </summary>
public class PoiExcel
{
    #region 内部变量
    /// <summary>
    /// Excel类型,2007/2003
    /// </summary>
    private readonly ExcelType _excelType = ExcelType.Excel2007;

    /// <summary>
    /// Excel格式工厂类
    /// </summary>
    private readonly StyleFactory _style;

    /// <summary>
    /// Excel工作簿
    /// </summary>
    private readonly IWorkbook _workbook;

    /// <summary>
    /// 当前工作表
    /// </summary>
    private ISheet _sheet;
    #endregion

    #region 构造函数
    /// <summary>
    /// 创建Workbook对象，同时创建一个Sheet
    /// </summary>
    /// <param name="sheetName">默认创建一个Sheet1</param>
    /// <param name="excelType">Excel格式，默认2007</param>
    /// <param name="defaultStyleConfig">配置默认单元格样式</param>
    public PoiExcel(string sheetName = "Sheet1", ExcelType excelType = ExcelType.Excel2007, Action<ExcelStyle>? defaultStyleConfig = null)
    {
        _excelType = excelType;
        _workbook = PoiHelper.CreateWorkbook(IsExcel2007);
        _sheet = _workbook.CreateSheet(sheetName);
        // 初始化样式
        defaultStyleConfig?.Invoke(_defaultStyle);
        _style = new StyleFactory(_workbook, _defaultStyle);
    }

    /// <summary>
    /// 根据文件流创建Workboox
    /// </summary>
    public PoiExcel(Stream stream, Action<ExcelStyle>? defaultStyleConfig = null)
    {
        _workbook = PoiHelper.CreateWorkbook(stream);
        _sheet = _workbook.GetSheetAt(0);
        _excelType = _workbook is XSSFWorkbook ? ExcelType.Excel2003 : ExcelType.Excel2007;
        // 初始化样式
        defaultStyleConfig?.Invoke(_defaultStyle);
        _style = new StyleFactory(_workbook, _defaultStyle);
    }
    #endregion

    #region 添加标题行
    /// <summary>
    /// 添加标题行：支持多行标题，用#分隔
    /// <para>如果多行标题是一样的，需要上下合并的时候可以省略#,但只有所有行都一样才行。也就是每列标题中的#号要么没有，要么#号数量必须一致</para>
    /// <para>如：[A#B,A#C,D,E#F],会生成两行标题：A,A,D,E ； B,C,D,F,第一行AA会合并，第三列DD会合并</para>
    /// </summary>
    public void AddTitleRow(int rowIndex, string titles)
    {
        AddTitleRow(rowIndex, titles.Split(','));
    }
    /// <summary>
    /// 添加标题行：支持多行标题，用#分隔
    /// <para>如果多行标题是一样的，需要上下合并的时候可以省略#,但只有所有行都一样才行。也就是每列标题中的#号要么没有，要么#号数量必须一致</para>
    /// <para>如：[A#B,A#C,D,E#F],会生成两行标题：A,A,D,E ； B,C,D,F,第一行AA会合并，第三列DD会合并</para>
    /// </summary>
    public void AddTitleRow(int rowIndex, params string[] titles)
    {
        AddTitleRow(rowIndex, 0, titles);
    }
    /// <summary>
    /// 添加标题行：支持多行标题，用#分隔
    /// <para>如果多行标题是一样的，需要上下合并的时候可以省略#,但只有所有行都一样才行。也就是每列标题中的#号要么没有，要么#号数量必须一致</para>
    /// <para>如：[A#B,A#C,D,E#F],会生成两行标题：A,A,D,E ； B,C,D,F,第一行AA会合并，第三列DD会合并</para>
    /// </summary>
    public void AddTitleRow(int rowIndex, int startColIndex, params string[] titles)
    {
        var splitTitles = titles.Select(a => a.Trim().Split('#')).ToArray();

        var columnCount = titles.Length;
        var endColIndex = startColIndex + columnCount - 1;
        var titleRowCount = splitTitles.Max(a => a.Length);
        int startRowIndex = rowIndex, endRowIndex = rowIndex + titleRowCount - 1;

        var colIndex = startColIndex;
        foreach (var title in splitTitles)
        {
            if (title.Length != titleRowCount && title.Length != 1)
            {
                throw new MessageException($"标题行数不一致，最大行数为{titleRowCount}");
            }
            for (var i = 0; i < titleRowCount; i++)
            {
                var cell = PoiHelper.GetCell(_sheet, rowIndex + i, colIndex, TitleStyle);
                cell.SetCellValue(title.Length == 1 ? title[0] : title[i]);
            }
            colIndex++;
        }

        if (titleRowCount > 1)
        {
            //合并标题行
            for (var i = 0; i < columnCount; i++)
            {
                PoiHelper.AutoMergeRows(_sheet, startColIndex + i, startRowIndex, endRowIndex);
            }
            // 合并标题列
            for (var i = 0; i < titleRowCount; i++)
            {
                PoiHelper.AutoMergeColumns(_sheet, rowIndex + i, startColIndex, endColIndex);
            }
        }
    }
    #endregion

    #region 添加主标题
    /// <summary>
    /// 添加主标题，如想调整样式，直接修改MainTitleStyle即可
    /// </summary>
    public void AddMainTitle(int rowIndex, string title, int colSpan = 12, int height = 25)
    {
        AddMainTitle(rowIndex, 0, title, colSpan, height);
    }

    /// <summary>
    /// 添加主标题，如想调整样式，直接修改MainTitleStyle即可
    /// </summary>
    public void AddMainTitle(int rowIndex, int startCol, string title, int colSpan = 12, int height = 25)
    {
        var row = PoiHelper.GetRow(_sheet, rowIndex);
        PoiHelper.SetRowHeight(row, height);
        PoiHelper.AddMergedRegion(_sheet, rowIndex, rowIndex, startCol, startCol + colSpan - 1, MainTitleStyle);
        PoiHelper.GetCell(row, startCol).SetCellValue(title);
    }
    #endregion

    #region 添加备注说明
    /// <summary>
    /// 添加备注说明，如想调整样式，直接修改RemarkStyle即可
    /// </summary>
    public void AddRemarkRow(int rowIndex, string title, int colSpan = 12, int height = 60)
    {
        AddRemarkRow(rowIndex, 0, title, colSpan, height);
    }

    /// <summary>
    /// 添加备注说明，如想调整样式，直接修改RemarkStyle即可
    /// </summary>
    public void AddRemarkRow(int rowIndex, int startCol, string title, int colSpan = 12, int height = 60)
    {
        var row = PoiHelper.GetRow(_sheet, rowIndex);
        PoiHelper.SetRowHeight(row, height);
        PoiHelper.AddMergedRegion(_sheet, rowIndex, rowIndex, startCol, startCol + colSpan - 1, RemarkStyle);
        PoiHelper.GetCell(row, startCol).SetCellValue(title);
    }
    #endregion

    #region 添加数据
    /// <summary>
    /// 添加数据行
    /// </summary>
    public void AddDataRow(int rowIndex, params object[] values)
    {
        AddDataRow(rowIndex, 0, values);
    }
    /// <summary>
    /// 添加数据行
    /// </summary>
    public void AddDataRow(int rowIndex, int colIndex, params object[] values)
    {
        SetCellValues(rowIndex, colIndex, values);
    }

    public void SetCellValues(int rowIndex, int startColIndex, params object[] values)
    {
        SetCellValues(rowIndex, startColIndex, DefaultStyle, values);
    }

    public void SetCellValues(int rowIndex, int startColIndex, ICellStyle style, params object[] values)
    {
        var row = PoiHelper.GetRow(_sheet, rowIndex);
        for (var i = 0; i < values.Length; i++)
        {
            PoiHelper.SetCellValue(row, startColIndex + i, values[i], false, style);
        }
    }

    public void SetCellValue(int rowIndex, int colIndex, object value, ICellStyle? style = null)
    {
        var cell = PoiHelper.GetCell(_sheet, rowIndex, colIndex);
        if (style == null && value != null)
        {
            style = (value is DateTime || value is DateOnly) ? DateStyle : DefaultStyle;
        }
        PoiHelper.SetCellValue(cell, value, false, style);
    }

    public void SetCellFormular(int rowIndex, int colIndex, string formular, ICellStyle? style = null)
    {
        var cell = PoiHelper.GetCell(_sheet, rowIndex, colIndex);
        cell.CellStyle = style ?? DefaultStyle;
        cell.SetCellFormula(formular);
    }
    #endregion

    #region 公共方法
    /// <summary>
    /// 是否2007格式
    /// </summary>
    public bool IsExcel2007 => _excelType == ExcelType.Excel2007;

    /// <summary>
    /// contentType
    /// </summary>
    public string ContentType => IsExcel2007 ? "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" : "application/vnd.ms-excel";

    public void CreateSheet(string sheetName)
    {
        _sheet = _workbook.CreateSheet(sheetName);
    }

    /// <summary>
    /// 创建单元格样式
    /// </summary>
    /// <param name="style"></param>
    /// <returns></returns>
    public ICellStyle CreateStyle(ExcelStyle style)
    {
        return _style.CreateStyle(style);
    }

    /// <summary>
    /// 将Excel的列索引转换为列名，列索引从0开始，列名从A开始。如第0列为A，第1列为B...
    /// </summary>
    /// <param name="columnIndex">列索引</param>
    /// <returns>列名，如第0列为A，第1列为B...</returns>
    public static string GetColumnName(int columnIndex)
    {
        return PoiHelper.GetColumnName(columnIndex);
    }

    /// <summary>
    /// 合并单元格式
    /// </summary>
    public void AddMergedRegion(int startRow, int endRow, int startCol, int endCol, ICellStyle? style = null)
    {
        PoiHelper.AddMergedRegion(_sheet, startRow, endRow, startCol, endCol, style);
    }

    /// <summary>
    /// 设置列宽
    /// </summary>
    public void SetColumnWidth(int startColIndex, params int[] width)
    {
        for (var i = 0; i < width.Length; i++)
        {
            PoiHelper.SetColumnWidth(_sheet, startColIndex + i, width[i]);
        }
    }

    /// <summary>
    /// 设置行高
    /// </summary>
    public void SetRowHeight(int rowIndex, int height)
    {
        PoiHelper.SetRowHeight(_sheet, rowIndex, height);
    }

    /// <summary>
    /// 添加边框样式
    /// </summary>
    public void AddConditionalBorderStyle()
    {
        PoiHelper.AddConditionalBorderStyle(_sheet);
    }

    /// <summary>
    /// 获取Excel文件流
    /// </summary>
    /// <returns></returns>
    public byte[] GetExcelBytes()
    {
        return PoiHelper.GetExcelBytes(_workbook, false);
    }
    #endregion

    #region 默认样式
    // 默认有边框，水平居中，垂直居中
    private ExcelStyle _defaultStyle = ExcelStyle.NewStyle().Border().Align(1, 1);

    /// <summary>
    /// 默认单元格样式
    /// </summary>
    public ICellStyle DefaultStyle => _style.DefaultStyle;

    /// <summary>
    /// 空样式，完全不设置样式
    /// </summary>
    public ICellStyle EmptyStyle => _style.EmptyStyle;

    /// <summary>
    /// 标题行样式: 水平居中，垂直居中，背景色26
    /// </summary>
    public ICellStyle TitleStyle => _style.NewStyle(ExcelStyle.NewStyle().HAlign(2).BgColor(26));

    /// <summary>
    /// 主标题样式：水平居中，垂直居中，背景色26，字体13
    /// </summary>
    public ICellStyle MainTitleStyle => _style.NewStyle(ExcelStyle.NewStyle().HAlign(2).BgColor(26).FontS(13));

    /// <summary>
    /// 备注说明样式：水平居左，垂直居中，自动换行
    /// </summary>
    public ICellStyle RemarkStyle => _style.NewStyle(ExcelStyle.NewStyle().Wrap());

    /// <summary>
    ///日期类型单元格样式yyyy-MM-dd
    /// </summary>
    public ICellStyle DateStyle => _style.DateStyle;

    /// <summary>
    ///数字类型单元格样式 0.00,默认不指定，有需要时可以手动指定该样式
    /// </summary>
    public ICellStyle NumberStyle => _style.NumberStyle;

    /// <summary>
    /// 文本样式
    /// </summary>
    public ICellStyle TextStyle => _style.TextStyle;

    /// <summary>
    /// 百分比样式
    /// </summary>
    public ICellStyle PersentStyle => _style.PersentStyle;

    /// <summary>
    /// 自动换行样式
    /// </summary>
    public ICellStyle WrapStyle => _style.WrapStyle;
    #endregion
}
