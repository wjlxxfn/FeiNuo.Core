using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Data;

namespace FeiNuo.Core;

/// <summary>
/// 封装Workbook和Sheet对象，提供常用的操作方法
/// </summary>
public class PoiExcel
{
    #region 构造函数，初始化内部变量
    private readonly IWorkbook _workbook;
    private ISheet _sheet = null!;
    private readonly StyleFactory _style;
    private readonly ExcelType _excelType;
    internal PoiExcel(IWorkbook workbook, ISheet sheet, StyleFactory? styleFactory = null)
    {
        this._workbook = workbook;
        this._sheet = sheet;
        this._style = styleFactory ?? new StyleFactory(workbook);
        _excelType = _workbook is XSSFWorkbook ? ExcelType.Excel2003 : ExcelType.Excel2007;
    }
    #endregion

    #region 公共接口
    /// <summary>
    /// 放开外部使用，以便自由调整
    /// </summary>
    public IWorkbook Workbook => _workbook;

    /// <summary>
    /// 放开外部使用，以便自由调整
    /// </summary>
    public ISheet Sheet => _sheet;

    /// <summary>
    /// 是否2007格式
    /// </summary>
    public bool IsExcel2007 => _excelType == ExcelType.Excel2007;

    /// <summary>
    /// contentType
    /// </summary>
    public string ContentType => IsExcel2007 ? "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" : "application/vnd.ms-excel";

    /// <summary>
    /// 创建Sheet
    /// </summary>
    public ISheet CreateSheet(string sheetName, bool? forceFormulaRecalculation = null)
    {
        _sheet = _workbook.CreateSheet(sheetName);
        if (forceFormulaRecalculation ?? false)
        {
            _sheet.ForceFormulaRecalculation = true;
        }
        return _sheet;
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
    /// 获取Excel文件流
    /// </summary>
    /// <returns></returns>
    public byte[] GetExcelBytes()
    {
        return PoiHelper.GetExcelBytes(_workbook, false);
    }
    #endregion

    #region 链接调用
    /// <summary>
    /// 设置当前Sheet
    /// </summary>
    public PoiExcel ActiveSheet(string sheetName)
    {
        _sheet = _workbook.GetSheet(sheetName);
        return this;
    }

    /// <summary>
    /// 配置Sheet,也可以用PoiExcel.Sheet获取sheet后直接操作
    /// </summary>
    /// <param name="sheetConfig"></param>
    public PoiExcel ConfigSheet(Action<ISheet> sheetConfig)
    {
        sheetConfig(_sheet);
        return this;
    }

    /// <summary>
    /// 合并单元格式
    /// </summary>
    public PoiExcel AddMergedRegion(int startRow, int endRow, int startCol, int endCol, ICellStyle? style = null)
    {
        PoiHelper.AddMergedRegion(_sheet, startRow, endRow, startCol, endCol, style);
        return this;
    }

    /// <summary>
    /// 设置列宽
    /// </summary>
    public PoiExcel SetColumnWidth(int startColIndex, params int[] width)
    {
        for (var i = 0; i < width.Length; i++)
        {
            PoiHelper.SetColumnWidth(_sheet, startColIndex + i, width[i]);
        }
        return this;
    }

    /// <summary>
    /// 设置列宽
    /// </summary>
    public PoiExcel SetColumnHidden(int colIndex, bool hidden = true)
    {
        PoiHelper.SetColumnHidden(_sheet, colIndex, hidden);
        return this;
    }

    /// <summary>
    /// 设置列宽
    /// </summary>
    public PoiExcel SetColumnHidden(params int[] colIndex)
    {
        foreach (var i in colIndex)
        {
            PoiHelper.SetColumnHidden(_sheet, i, true);
        }
        return this;
    }

    /// <summary>
    /// 设置行高
    /// </summary>
    public PoiExcel SetRowHeight(int rowIndex, int height)
    {
        PoiHelper.SetRowHeight(_sheet, rowIndex, height);
        return this;
    }

    /// <summary>
    /// 添加边框条件样式：自动计算有数据的内容区域并设置thin边框
    /// </summary>
    public PoiExcel AddConditionalBorderStyle()
    {
        PoiHelper.AddConditionalBorderStyle(_sheet);
        return this;
    }
    #endregion

    #region 默认样式: 标题，主标题，备注
    /// <summary>
    /// 标题行样式: 水平居中，垂直居中，背景色26, 可在添加标题行前调整
    /// </summary>
    public ExcelStyle TitleStyle => ExcelStyle.NewStyle().Border().Align(2, 1).BgColor(26);

    /// <summary>
    /// 主标题样式：水平居中，垂直居中，背景色26，字体13, 可在添加标题行前调整
    /// </summary>
    public ExcelStyle MainTitleStyle => ExcelStyle.NewStyle().Border().Align(2, 1).BgColor(26).FontS(13);

    /// <summary>
    /// 备注说明样式：水平居左，垂直居中，自动换行, 可在添加标题行前调整
    /// </summary>
    public ExcelStyle RemarkStyle => ExcelStyle.NewStyle().Border().Align(1, 1).Wrap();
    #endregion

    #region 添加标题行
    /// <summary>
    /// 添加标题行：支持多行标题，用#分隔
    /// <para>如果多行标题是一样的，需要上下合并的时候可以省略#,但只有所有行都一样才行。也就是每列标题中的#号要么没有，要么#号数量必须一致</para>
    /// <para>如：[A#B,A#C,D,E#F],会生成两行标题：A,A,D,E ； B,C,D,F,第一行AA会合并，第三列DD会合并</para>
    /// </summary>
    public PoiExcel AddTitleRow(int rowIndex, string titles)
    {
        return AddTitleRow(rowIndex, titles.Split(','));
    }

    /// <summary>
    /// 添加标题行：支持多行标题，用#分隔
    /// <para>如果多行标题是一样的，需要上下合并的时候可以省略#,但只有所有行都一样才行。也就是每列标题中的#号要么没有，要么#号数量必须一致</para>
    /// <para>如：[A#B,A#C,D,E#F],会生成两行标题：A,A,D,E ； B,C,D,F,第一行AA会合并，第三列DD会合并</para>
    /// </summary>
    public PoiExcel AddTitleRow(int rowIndex, params string[] titles)
    {
        return AddTitleRow(rowIndex, 0, titles);
    }

    /// <summary>
    /// 添加标题行：支持多行标题，用#分隔
    /// <para>如果多行标题是一样的，需要上下合并的时候可以省略#,但只有所有行都一样才行。也就是每列标题中的#号要么没有，要么#号数量必须一致</para>
    /// <para>如：[A#B,A#C,D,E#F],会生成两行标题：A,A,D,E ； B,C,D,F,第一行AA会合并，第三列DD会合并</para>
    /// </summary>
    /// <returns>返回标题行数</returns>
    public PoiExcel AddTitleRow(int rowIndex, int colIndex, params string[] titles)
    {
        AddTitleRows(rowIndex, colIndex, titles);
        return this;
    }

    /// <summary>
    /// 添加标题行：支持多行标题，用#分隔
    /// <para>如果多行标题是一样的，需要上下合并的时候可以省略#,但只有所有行都一样才行。也就是每列标题中的#号要么没有，要么#号数量必须一致</para>
    /// <para>如：[A#B,A#C,D,E#F],会生成两行标题：A,A,D,E ； B,C,D,F,第一行AA会合并，第三列DD会合并</para>
    /// </summary>
    /// <returns>返回标题行数</returns>
    public int AddTitleRows(int rowIndex, int colIndex, params string[] titles)
    {
        var splitTitles = titles.Select(a => a.Trim().Split('#')).ToArray();

        var rowCount = splitTitles.Max(a => a.Length);
        int startRowIndex = rowIndex, endRowIndex = startRowIndex + rowCount - 1;
        var columnCount = titles.Length;
        int startColIndex = colIndex, endColIndex = startColIndex + columnCount - 1;

        var style = _style.CreateStyle(TitleStyle);
        foreach (var title in splitTitles)
        {
            if (title.Length != rowCount && title.Length != 1)
            {
                throw new MessageException($"标题行数不一致，最大行数为{rowCount}");
            }
            for (var i = 0; i < rowCount; i++)
            {
                var cell = PoiHelper.GetCell(_sheet, startRowIndex + i, colIndex, style);
                cell.SetCellValue(title.Length == 1 ? title[0] : title[i]);
            }
            colIndex++;
        }

        if (rowCount > 1)
        {
            //合并标题行
            for (var i = 0; i < columnCount; i++)
            {
                PoiHelper.AutoMergeRows(_sheet, startColIndex + i, rowIndex, endRowIndex);
            }
            // 合并标题列
            for (var i = 0; i < rowCount; i++)
            {
                PoiHelper.AutoMergeColumns(_sheet, startRowIndex + i, startColIndex, endColIndex);
            }
        }
        return rowCount;
    }
    #endregion

    #region 添加主标题
    /// <summary>
    /// 添加主标题，如想调整样式，在该用该方法前修改MainTitleStyle即可
    /// </summary>
    public PoiExcel AddMainTitle(int rowIndex, string title, int colSpan = 12, int height = 25)
    {
        return AddMainTitle(rowIndex, 0, title, colSpan, height);
    }

    /// <summary>
    /// 添加主标题，如想调整样式，在该用该方法前修改MainTitleStyle即可
    /// </summary>
    public PoiExcel AddMainTitle(int rowIndex, int startCol, string title, int colSpan = 12, int height = 25)
    {
        var row = PoiHelper.GetRow(_sheet, rowIndex);
        PoiHelper.SetRowHeight(row, height);
        var style = _style.CreateStyle(MainTitleStyle);
        PoiHelper.AddMergedRegion(_sheet, rowIndex, rowIndex, startCol, startCol + colSpan - 1, style);
        PoiHelper.GetCell(row, startCol).SetCellValue(title);
        return this;
    }
    #endregion

    #region 添加备注说明
    /// <summary>
    /// 添加备注说明，如想调整样式，在该用该方法前修改RemarkStyle即可
    /// </summary>
    public PoiExcel AddRemarkRow(int rowIndex, string title, int colSpan = 12, int height = 60)
    {
        return AddRemarkRow(rowIndex, 0, title, colSpan, height);
    }

    /// <summary>
    /// 添加备注说明，如想调整样式，在该用该方法前修改RemarkStyle即可
    /// </summary>
    public PoiExcel AddRemarkRow(int rowIndex, int startCol, string title, int colSpan = 12, int height = 60)
    {
        var row = PoiHelper.GetRow(_sheet, rowIndex);
        PoiHelper.SetRowHeight(row, height);
        var style = _style.CreateStyle(RemarkStyle);
        PoiHelper.AddMergedRegion(_sheet, rowIndex, rowIndex, startCol, startCol + colSpan - 1, style);
        PoiHelper.GetCell(row, startCol).SetCellValue(title);
        return this;
    }
    #endregion

    #region 添加数据
    /// <summary>
    /// 根据数据集合，自动添加标题和数据行
    /// <para>标题取第一个数据对象的属性名,所以第一条数据不能是空</para>
    /// </summary>
    /// <param name="dataList"></param>
    /// <param name="rowIndex">开始行索引,默认0</param>
    /// <param name="includeTitle">是否生成title，默认true</param>
    public PoiExcel AddDataList(IEnumerable<object> dataList, int rowIndex = 0, bool includeTitle = true)
    {
        var lstData = dataList.ToList();
        if (lstData.Count == 0) return this;

        var props = lstData.First().GetType().GetProperties();

        if (includeTitle)
        {
            AddTitleRow(rowIndex++, [.. props.Select(a => a.Name)]);
        }
        foreach (var data in lstData)
        {
            var values = props.Select(a => a.GetValue(data)).ToArray();
            AddDataRow(rowIndex++, values);
        }
        return this;
    }
    public PoiExcel AddDataList<T>(IEnumerable<T> dataList, IEnumerable<ExcelColumn<T>> columns, int startRow = 0, int startCol = 0) where T : class
    {
        var rowIndex = startRow;

        // 标题
        var titles = columns.Select(a => a.Title).ToArray();
        var titleRowCount = titles.Select(a => a.Split('#').Length).Max();
        AddTitleRow(rowIndex, titles);
        // 数据 
        rowIndex += titleRowCount;
        foreach (var data in dataList)
        {
            var values = columns.Select(a => a.ValueGetter(data)).ToArray();
            AddDataRow(rowIndex++, startCol, values);
        }
        // 列配置
        var colIndex = startCol;
        foreach (var column in columns)
        {
            if (column.Width.HasValue)
            {
                SetColumnWidth(colIndex, column.Width.Value);
            }
            if (column.Hidden)
            {
                SetColumnHidden(colIndex);
            }
            colIndex++;
        }
        return this;
    }
    /// <summary>
    /// 根据DataTable 自动添加标题和数据 
    /// <para>标题取ColumnName</para>
    /// </summary>
    /// <param name="dt">数据集</param>
    /// <param name="startRow">开始行索引,默认0</param>
    /// <param name="includeTitle">是否生成title，默认true</param>
    public PoiExcel AddDataTable(DataTable dt, int startRow = 0, bool includeTitle = true)
    {
        var rowIndex = startRow;
        if (includeTitle)
        {
            var titles = "";
            foreach (DataColumn dc in dt.Columns)
            {
                titles += dc.ColumnName + ",";
            }
            AddTitleRow(rowIndex++, titles.TrimEnd(','));
        }
        foreach (DataRow dr in dt.Rows)
        {
            AddDataRow(rowIndex++, dr.ItemArray);
        }
        return this;
    }

    /// <summary>
    /// 添加数据行
    /// </summary>
    public PoiExcel AddDataRow(int rowIndex, params object?[] values)
    {
        return AddDataRow(rowIndex, 0, values);
    }
    /// <summary>
    /// 添加数据行
    /// </summary>
    public PoiExcel AddDataRow(int rowIndex, int startColIndex, params object?[] values)
    {
        return SetCellValues(rowIndex, startColIndex, values);
    }

    public PoiExcel SetCellValues(int rowIndex, int startColIndex, params object?[] values)
    {
        var row = PoiHelper.GetRow(_sheet, rowIndex);
        for (var i = 0; i < values.Length; i++)
        {
            var style = (values[i] is DateTime || values[i] is DateOnly) ? _style.DateStyle : _style.DefaultStyle;
            PoiHelper.SetCellValue(row, startColIndex + i, values[i], false, style);
        }
        return this;
    }

    public PoiExcel SetCellValues(int rowIndex, int startColIndex, ICellStyle style, params object?[] values)
    {
        var row = PoiHelper.GetRow(_sheet, rowIndex);
        for (var i = 0; i < values.Length; i++)
        {
            PoiHelper.SetCellValue(row, startColIndex + i, values[i], false, style);
        }
        return this;
    }

    public PoiExcel SetCellValue(int rowIndex, int colIndex, object? value, ICellStyle? style = null)
    {
        var cell = PoiHelper.GetCell(_sheet, rowIndex, colIndex);
        style ??= (value is DateTime || value is DateOnly) ? _style.DateStyle : _style.DefaultStyle;
        PoiHelper.SetCellValue(cell, value, false, style);
        return this;
    }

    public PoiExcel SetCellFormular(int rowIndex, int colIndex, string formular, ICellStyle? style = null)
    {
        var cell = PoiHelper.GetCell(_sheet, rowIndex, colIndex);
        cell.CellStyle = style ?? _style.DefaultStyle;
        cell.SetCellFormula(formular);
        return this;
    }
    #endregion

}
