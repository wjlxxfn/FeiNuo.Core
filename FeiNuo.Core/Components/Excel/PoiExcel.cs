using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Data;

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
    public PoiExcel(Stream stream, int activeSheetIndex = 0, Action<ExcelStyle>? defaultStyleConfig = null)
    {
        _workbook = PoiHelper.CreateWorkbook(stream);
        _sheet = _workbook.GetSheetAt(activeSheetIndex);
        _excelType = _workbook is XSSFWorkbook ? ExcelType.Excel2003 : ExcelType.Excel2007;
        // 初始化样式
        defaultStyleConfig?.Invoke(_defaultStyle);
        _style = new StyleFactory(_workbook, _defaultStyle);
    }
    #endregion

    #region 静态方法,直接填充数据
    /// <summary>
    /// 创建空的Excel对象，添加默认的Sheet1
    /// </summary>
    public static PoiExcel CreateExcel(string sheetName = "Sheet1", ExcelType excelType = ExcelType.Excel2007, Action<ExcelStyle>? defaultStyleConfig = null)
    {
        return new PoiExcel(sheetName, excelType, defaultStyleConfig);
    }
    /// <summary>
    /// 根据文件流创建Excel对象，当前Sheet默认设置为第一个Sheet
    /// </summary>
    public static PoiExcel CreateExcel(Stream stream, int activeSheetIndex = 0, Action<ExcelStyle>? defaultStyleConfig = null)
    {
        return new PoiExcel(stream, activeSheetIndex, defaultStyleConfig);
    }

    /// <summary>
    /// 根据数据集合构造Excel对象
    /// <para>标题取第一个数据对象的属性名,所以第一条数据不能是空</para>
    /// <para>使用示例: new PoiExcel(users.Select(a=>new {用户名=a.Username,姓名=a.NickName}))</para>
    /// </summary>
    /// <param name="dataList">数据</param>
    public static PoiExcel CreateExcel(IEnumerable<object> dataList)
    {
        var poi = new PoiExcel();
        poi.AddDataList(dataList);
        return poi;
    }

    /// <summary>
    /// 根据数据集合构造Excel对象(多个Sheet,通过sheetName为键加入到Dictionary中)
    /// <para>标题取第一个数据对象的属性名,所以第一条数据不能是空</para>
    /// <para>使用示例： var dict = new Dictionary(); </para>
    /// <para>           dict.Add("用户", users.Select(a => new { 用户名 = a.Username, 姓名 = a.NickName });</para>
    /// <para>           new PoiExcel(dict);</para>
    /// </summary>
    /// <param name="dataMap">Key=SheetName,Value=dataList</param>
    public static PoiExcel CreateExcel(Dictionary<string, IEnumerable<object>> dataMap)
    {
        var poi = new PoiExcel().RemoveSheet1();
        foreach (var data in dataMap)
        {
            poi.CreateSheet(data.Key);
            poi.AddDataList(data.Value);
        }
        return poi;
    }

    /// <summary>
    /// 根据DataTable构建Excel，标题=ColumnName
    /// </summary>
    public static PoiExcel CreateExcel(DataTable dt)
    {
        var poi = new PoiExcel();
        poi.AddDataTable(dt);
        return poi;
    }

    /// <summary>
    /// 根据DataTable构建Excel，SheetName=TableName,标题=ColumnName
    /// </summary>
    public static PoiExcel CreateExcel(DataSet ds)
    {
        var poi = new PoiExcel().RemoveSheet1();
        foreach (DataTable dt in ds.Tables)
        {
            poi.CreateSheet(dt.TableName);
            poi.AddDataTable(dt);
        }
        return poi;
    }

    /// <summary>
    /// 根据数据集和列配置自动生成Excel
    /// <para>注意：列配置的Style不生效，这里只使用DefaultStyle</para>
    /// </summary>
    public static PoiExcel CreateExcel<T>(IEnumerable<T> dataList, IEnumerable<ExcelColumn<T>> columns) where T : class
    {
        return new PoiExcel().AddDataList(dataList, columns);
    }

    /// <summary>
    /// 通过列配置创建Excel对象，不添加数据，每列添加默认格式，主要用于下载导入模板 
    /// </summary>
    public static PoiExcel CreateExcel<T>(IEnumerable<ExcelColumn<T>> columns, int startRowIndex = 0, int startColIndex = 0) where T : class
    {
        var poi = new PoiExcel();
        var titles = columns.Select(a => a.Title).ToArray();
        var titleRowCount = titles.Select(a => a.Split('#').Length).Max();
        poi.AddTitleRow(startRowIndex, startColIndex, titles);

        var colIndex = startColIndex;
        foreach (var column in columns)
        {
            if (column.Width.HasValue)
            {
                poi.SetColumnWidth(colIndex, column.Width.Value);
            }
            if (column.Hidden)
            {
                poi.SetColumnHidden(colIndex);
            }
            if (column.ColumnStyle.IsNotEmptyStyle)
            {
                poi.SetDefaultColumnStyle(colIndex, poi.CreateStyle(column.ColumnStyle));
            }
            colIndex++;
        }
        return poi;
    }
    #endregion

    #region 配置Excel方法
    /// <summary>
    /// 构造函数会默认创建一个Sheet1，有时是不必要的，调用该方法删除默认Sheet.
    /// <para>但删除后需立即手动创建Sheet,否则内部sheet为空会报错.</para>
    /// </summary>
    public PoiExcel RemoveSheet1()
    {
        _workbook.RemoveSheetAt(0);
        _sheet = null!;
        return this;
    }

    /// <summary>
    /// 创建Sheet
    /// </summary>
    public PoiExcel CreateSheet(string sheetName)
    {
        _sheet = _workbook.CreateSheet(sheetName);
        return this;
    }

    /// <summary>
    /// 设置当前Sheet名称
    /// </summary>
    public PoiExcel SetSheetName(string sheetName)
    {
        var index = _workbook.GetSheetIndex(_sheet);
        _workbook.SetSheetName(index, sheetName);
        return this;
    }

    /// <summary>
    /// 设置当前Sheet
    /// </summary>
    public PoiExcel ActiveSheet(int sheetIndex)
    {
        _sheet = _workbook.GetSheetAt(sheetIndex);
        return this;
    }

    /// <summary>
    /// 设置当前Sheet
    /// </summary>
    public PoiExcel ActiveSheet(string sheetName)
    {
        _sheet = _workbook.GetSheet(sheetName);
        return this;
    }

    /// <summary>
    /// 设置默认列宽
    /// </summary>
    public PoiExcel SetDefaultColumnWidth(double width)
    {
        _sheet.DefaultColumnWidth = width;
        return this;
    }

    /// <summary>
    /// 设置列的默认格式
    /// </summary>
    public PoiExcel SetDefaultColumnStyle(int colIndex, ICellStyle cellStyle)
    {
        _sheet.SetDefaultColumnStyle(colIndex, cellStyle);
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
    /// 设置行高
    /// </summary>
    public PoiExcel SetRowHeight(int rowIndex, int height)
    {
        PoiHelper.SetRowHeight(_sheet, rowIndex, height);
        return this;
    }

    /// <summary>
    /// 添加边框样式
    /// </summary>
    public PoiExcel AddConditionalBorderStyle()
    {
        PoiHelper.AddConditionalBorderStyle(_sheet);
        return this;
    }
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
    public PoiExcel AddTitleRow(int rowIndex, int startColIndex, params string[] titles)
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

        return this;
    }
    #endregion

    #region 添加主标题
    /// <summary>
    /// 添加主标题，如想调整样式，直接修改MainTitleStyle即可
    /// </summary>
    public PoiExcel AddMainTitle(int rowIndex, string title, int colSpan = 12, int height = 25)
    {
        return AddMainTitle(rowIndex, 0, title, colSpan, height);
    }

    /// <summary>
    /// 添加主标题，如想调整样式，直接修改MainTitleStyle即可
    /// </summary>
    public PoiExcel AddMainTitle(int rowIndex, int startCol, string title, int colSpan = 12, int height = 25)
    {
        var row = PoiHelper.GetRow(_sheet, rowIndex);
        PoiHelper.SetRowHeight(row, height);
        PoiHelper.AddMergedRegion(_sheet, rowIndex, rowIndex, startCol, startCol + colSpan - 1, MainTitleStyle);
        PoiHelper.GetCell(row, startCol).SetCellValue(title);
        return this;
    }
    #endregion

    #region 添加备注说明
    /// <summary>
    /// 添加备注说明，如想调整样式，直接修改RemarkStyle即可
    /// </summary>
    public PoiExcel AddRemarkRow(int rowIndex, string title, int colSpan = 12, int height = 60)
    {
        return AddRemarkRow(rowIndex, 0, title, colSpan, height);
    }

    /// <summary>
    /// 添加备注说明，如想调整样式，直接修改RemarkStyle即可
    /// </summary>
    public PoiExcel AddRemarkRow(int rowIndex, int startCol, string title, int colSpan = 12, int height = 60)
    {
        var row = PoiHelper.GetRow(_sheet, rowIndex);
        PoiHelper.SetRowHeight(row, height);
        PoiHelper.AddMergedRegion(_sheet, rowIndex, rowIndex, startCol, startCol + colSpan - 1, RemarkStyle);
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
    /// <param name="startRow">开始行索引,默认0</param>
    /// <param name="includeTitle">是否生成title，默认true</param>
    public PoiExcel AddDataList(IEnumerable<object> dataList, int startRow = 0, bool includeTitle = true)
    {
        var lstData = dataList.ToList();
        if (lstData.Count == 0) return this;

        var props = lstData.First().GetType().GetProperties();

        var rowIndex = startRow;
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
        return SetCellValues(rowIndex, startColIndex, DefaultStyle, values);
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
        if (style == null && value != null)
        {
            style = (value is DateTime || value is DateOnly) ? DateStyle : DefaultStyle;
        }
        PoiHelper.SetCellValue(cell, value, false, style);
        return this;
    }

    public PoiExcel SetCellFormular(int rowIndex, int colIndex, string formular, ICellStyle? style = null)
    {
        var cell = PoiHelper.GetCell(_sheet, rowIndex, colIndex);
        cell.CellStyle = style ?? DefaultStyle;
        cell.SetCellFormula(formular);
        return this;
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

    #region 默认样式
    // 默认有边框，水平居中，垂直居中
    private readonly ExcelStyle _defaultStyle = ExcelStyle.NewStyle().Border().Align(1, 1);

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
    public ICellStyle TitleStyle => _style.NewStyle(ExcelStyle.NewStyle().Border().HAlign(2).BgColor(26));

    /// <summary>
    /// 主标题样式：水平居中，垂直居中，背景色26，字体13
    /// </summary>
    public ICellStyle MainTitleStyle => _style.NewStyle(ExcelStyle.NewStyle().Border().HAlign(2).BgColor(26).FontS(13));

    /// <summary>
    /// 备注说明样式：水平居左，垂直居中，自动换行
    /// </summary>
    public ICellStyle RemarkStyle => _style.NewStyle(ExcelStyle.NewStyle().Border().Wrap());

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

    #region 暴漏POI对象，以便外部使用
    public IWorkbook Workbook => _workbook;
    public ISheet Sheet => _sheet;
    #endregion
}
