using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace FeiNuo.Core;

/// <summary>
/// 封装Workbook和Sheet对象，提供常用的操作方法
/// </summary>
public class PoiExcel
{
    /// <summary>
    /// Excel类型,2007/2003
    /// </summary>
    private readonly ExcelType ExcelType = ExcelType.Excel2007;

    /// <summary>
    /// Excel格式工厂类
    /// </summary>
    private readonly StyleFactory _styleFactory;

    /// <summary>
    /// Excel工作簿
    /// </summary>
    private readonly IWorkbook _workbook;

    /// <summary>
    /// 当前工作表
    /// </summary>
    private ISheet _sheet;

    #region 默认样式
    /// <summary>
    /// 默认单元格样式
    /// </summary>
    public ICellStyle DefaultStyle { get; private set; } = null!;

    /// <summary>
    /// 标题行单元格样式
    /// </summary>
    public ICellStyle TitleStyle { get; private set; } = null!;

    /// <summary>
    ///日期类型单元格样式yyyy-MM-dd
    /// </summary>
    public ICellStyle DateStyle { get; private set; } = null!;

    /// <summary>
    ///数字类型单元格样式 0.00,默认不指定，有需要时可以手动指定该样式
    /// </summary>
    public ICellStyle NumberStyle { get; private set; } = null!;

    /// <summary>
    /// 文本样式
    /// </summary>
    public ICellStyle TextStyle { get; private set; } = null!;

    private void InitInternalStyles()
    {
        DefaultStyle = _styleFactory.DefaultStyle;
        TitleStyle = _styleFactory.GetStyle(new ExcelStyle() { HorizontalAlignment = 2, BackgroundColor = (short)26 });
        DateStyle = _styleFactory.DateStyle;
        NumberStyle = _styleFactory.NumbericStyle;
        TextStyle = _styleFactory.GetStyle(new ExcelStyle() { DataFormat = "@" });
    }
    #endregion

    #region 构造函数，创建Workbook对象
    /// <summary>
    /// 创建Workbook对象，同时创建一个Sheet
    /// </summary>
    /// <param name="sheetName">默认创建一个Sheet1</param>
    /// <param name="excelType">Excel格式，默认2007</param>
    public PoiExcel(string sheetName = "Sheet1", ExcelType excelType = ExcelType.Excel2007)
    {
        ExcelType = excelType;
        _workbook = PoiHelper.CreateWorkbook(IsExcel2007);
        _sheet = _workbook.CreateSheet(sheetName);
        // 初始化样式
        _styleFactory = new StyleFactory(_workbook, new ExcelStyle() { BorderStyle = 1, VerticalAlignment = 1, HorizontalAlignment = 1 });
        InitInternalStyles();
    }

    /// <summary>
    /// 根据文件流创建Workboox
    /// </summary>
    public PoiExcel(Stream stream)
    {
        _workbook = PoiHelper.CreateWorkbook(stream);
        _sheet = _workbook.GetSheetAt(0);
        ExcelType = _workbook is XSSFWorkbook ? ExcelType.Excel2003 : ExcelType.Excel2007;
        // 初始化样式
        _styleFactory = new StyleFactory(_workbook);
        InitInternalStyles();
    }
    #endregion

    #region 主标题
    public void AddMainTitle(int rowIndex, string title, int colSpan = 12, ICellStyle? style = null)
    {
        AddMainTitle(rowIndex, 0, title, colSpan, style);
    }
    public void AddMainTitle(int rowIndex, int startCol, string title, int colSpan = 12, ICellStyle? style = null)
    {
        var row = PoiHelper.GetRow(_sheet, rowIndex);
        row.Height = 25 * 20;
        style ??= CreateStyle(new ExcelStyle() { HorizontalAlignment = 2, VerticalAlignment = 1 });
        for (var i = 0; i < colSpan; i++)
        {
            PoiHelper.GetCell(row, startCol + i).CellStyle = style;
        }
        PoiHelper.AddMergedRegion(_sheet, rowIndex, rowIndex, startCol, startCol + colSpan - 1, style);
        var cell = PoiHelper.GetCell(_sheet, rowIndex, startCol);
        cell.SetCellValue(title);
    }
    #endregion

    #region 添加数据
    public void SetCellValues(int rowIndex, int startColIndex, ICellStyle? style, params object[] values)
    {
        style ??= DefaultStyle;
        for (var i = 0; i < values.Length; i++)
        {
            SetCellValue(rowIndex, startColIndex + i, values[i], style);
        }
    }
    public void SetCellValues(int rowIndex, int startColIndex, params object[] values)
    {
        for (var i = 0; i < values.Length; i++)
        {
            SetCellValue(rowIndex, startColIndex + i, values[i]);
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
    public bool IsExcel2007 => ExcelType == ExcelType.Excel2007;

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
        return _styleFactory.GetStyle(style);
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
}
