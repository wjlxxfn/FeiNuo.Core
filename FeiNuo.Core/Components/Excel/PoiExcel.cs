using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Data;

namespace FeiNuo.Core;

/// <summary>
/// 封装Workbook和Sheet对象，提供常用的操作方法,
/// </summary>
public class PoiExcel
{
    #region 构造函数，初始化内部变量
    private readonly IWorkbook _workbook;
    private ISheet _sheet = null!;
    private readonly StyleFactory _style;
    private readonly ExcelType _excelType;
    private int _titleRowCount = 0;
    // 默认样式
    private readonly ExcelStyle TitleStyle = ExcelStyle.NewStyle().Border().Align(2, 1).BgColor(13);
    private readonly ExcelStyle MainTitleStyle = ExcelStyle.NewStyle().Border().Align(2, 1).BgColor(13).FontS(13);
    private readonly ExcelStyle RemarkStyle = ExcelStyle.NewStyle().Border().Align(1, 1).Wrap();

    public PoiExcel(IWorkbook workbook, ISheet? sheet = null, ExcelStyle? defaultStyle = null)
    {
        _workbook = workbook;
        // 如果没有sheet，则创建一个默认的sheet,如果有sheet，则使用第一个sheet
        _sheet = sheet ?? (workbook.NumberOfSheets >= 0 ? workbook.GetSheetAt(0) : workbook.CreateSheet("Sheet1"));
        _style = new StyleFactory(workbook, defaultStyle);
        _excelType = _workbook is XSSFWorkbook ? ExcelType.Excel2007 : ExcelType.Excel2003;
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

    public IRow GetRow(int rowIndex) => PoiHelper.GetRow(_sheet, rowIndex);

    public ICell GetCell(int rowIndex, int colIndex) => PoiHelper.GetCell(_sheet, rowIndex, colIndex);

    /// <summary>
    /// 样式工厂，内有多种常用样式，也可用来创建新样式
    /// </summary>
    public StyleFactory Styles => _style;

    /// <summary>
    /// 标题行的数量，需要在调用AddTitleRow之后才能获取
    /// </summary>
    public int TitleRowCount => _titleRowCount;

    /// <summary>
    /// 是否2007格式
    /// </summary>
    public bool IsExcel2007 => _excelType == ExcelType.Excel2007;

    /// <summary>
    /// contentType
    /// </summary>
    public string ContentType => IsExcel2007 ? "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" : "application/vnd.ms-excel";

    /// <summary>
    /// 文件名，默认为空。内部暂未使用，可自行赋值使用。
    /// </summary>
    public string FileName { get; set; } = string.Empty;
    public PoiExcel SetFileName(string fileName)
    {
        FileName = fileName;
        return this;
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

    #region 链式调用
    /// <summary>
    /// 创建Sheet
    /// </summary>
    public PoiExcel CreateSheet(string sheetName)
    {
        _sheet = _workbook.CreateSheet(sheetName);
        return this;
    }

    /// <summary>
    /// 复制工作表
    /// </summary>
    /// <param name="sheetName">新工作表名</param>
    public PoiExcel CopySheet(string sheetName)
    {
        _sheet = _sheet.CopySheet(sheetName);
        return this;
    }

    /// <summary>
    /// 移除Sheet,默认移除当前Sheet，执行后当前sheet置空
    /// <para>移除后需立即CreateSheet或者ActiveSheet,否则Sheet=null会报错</para>
    /// <para>通常是默认创建PoiExcel对象后已经有sheet1但又不想要时调用该方法删除</para>
    /// </summary>
    /// <returns></returns>
    public PoiExcel RemoveSheet(int? sheetIndex = null)
    {
        _workbook.RemoveSheetAt(sheetIndex ?? _workbook.GetSheetIndex(_sheet));
        _sheet = null!;
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
    /// 设置当前Sheet
    /// </summary>
    public PoiExcel ActiveSheetAt(int sheetIndex)
    {
        _sheet = _workbook.GetSheetAt(sheetIndex);
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
    public PoiExcel SetColumnWidth(int colIndex, params int[] width)
    {
        for (var i = 0; i < width.Length; i++)
        {
            PoiHelper.SetColumnWidth(_sheet, colIndex + i, width[i]);
        }
        return this;
    }

    /// <summary>
    /// 自动列宽
    /// </summary>
    public PoiExcel AutoSizeColumn(int startColIndex, params int[] width)
    {
        for (var i = 0; i < width.Length; i++)
        {
            PoiHelper.AutoSizeColumn(_sheet, startColIndex + i);
        }
        return this;
    }

    /// <summary>
    /// 设置列隐藏
    /// </summary>
    public PoiExcel SetColumnHidden(int colIndex, bool hidden = true)
    {
        PoiHelper.SetColumnHidden(_sheet, colIndex, hidden);
        return this;
    }

    /// <summary>
    /// 设置列隐藏
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

    /// <summary>
    /// 配置默认的标题样式，在添加标题行之前配置有效
    /// <para>默认样式：水平居中，垂直居中，背景色26</para>
    /// </summary>
    public PoiExcel ConfigTitleStyle(Action<ExcelStyle> styleConfig)
    {
        styleConfig.Invoke(TitleStyle);
        return this;
    }

    /// <summary>
    /// 配置默认的主标题样式，在添加主标题行之前配置有效
    /// <para>默认样式：水平居中，垂直居中，背景色26，字体13</para>
    /// </summary>
    public PoiExcel ConfigMainTitleStyle(Action<ExcelStyle> styleConfig)
    {
        styleConfig.Invoke(MainTitleStyle);
        return this;
    }

    /// <summary>
    /// 配置默认的备注行样式，在添加备注行之前配置有效
    /// <para>默认样式：水平居左，垂直居中，自动换行</para>
    /// </summary>
    public PoiExcel ConfigRemarkStyle(Action<ExcelStyle> styleConfig)
    {
        styleConfig.Invoke(RemarkStyle);
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
    /// <returns>返回标题行数</returns>
    public PoiExcel AddTitleRow(int rowIndex, int colIndex, params string[] titles)
    {
        var splitTitles = titles.Select(a => a.Trim().Split('#')).ToArray();
        _titleRowCount = splitTitles.Max(a => a.Length);


        int startRowIndex = rowIndex, endRowIndex = startRowIndex + _titleRowCount - 1;
        var columnCount = titles.Length;
        int startColIndex = colIndex, endColIndex = startColIndex + columnCount - 1;

        var style = _style.CreateStyle(TitleStyle);
        foreach (var title in splitTitles)
        {
            if (title.Length != _titleRowCount && title.Length != 1)
            {
                throw new Exception($"标题行数不一致，最大行数为{_titleRowCount}");
            }
            for (var i = 0; i < _titleRowCount; i++)
            {
                var cell = PoiHelper.GetCell(_sheet, startRowIndex + i, colIndex, style);
                cell.SetCellValue(title.Length == 1 ? title[0] : title[i]);
            }
            colIndex++;
        }

        if (_titleRowCount > 1)
        {
            //合并标题行
            for (var i = 0; i < columnCount; i++)
            {
                PoiHelper.AutoMergeRows(_sheet, startColIndex + i, rowIndex, endRowIndex);
            }
            // 合并标题列
            for (var i = 0; i < _titleRowCount; i++)
            {
                PoiHelper.AutoMergeColumns(_sheet, startRowIndex + i, startColIndex, endColIndex);
            }
        }
        return this;
    }

    /// <summary>
    /// 添加标题行：支持多行标题，用#分隔
    /// <para>如果多行标题是一样的，需要上下合并的时候可以省略#,但只有所有行都一样才行。也就是每列标题中的#号要么没有，要么#号数量必须一致</para>
    /// <para>如：[A#B,A#C,D,E#F],会生成两行标题：A,A,D,E ； B,C,D,F,第一行AA会合并，第三列DD会合并</para>
    /// </summary>
    /// <param name="rowIndex">开始行索引</param>
    /// <param name="colIndex">开始列索引</param>
    /// <param name="columns">列配置</param>
    /// <param name="setDefaultColumnStyle">true，则根据列配置的默认样式设置列默认样式</param>
    /// <returns>返回标题行数</returns>
    public PoiExcel AddTitleRow(int rowIndex, int colIndex, IEnumerable<ExcelColumn> columns, bool setDefaultColumnStyle = false)
    {
        var startColIndex = colIndex;

        #region 配置列的默认属性
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
            if (setDefaultColumnStyle)
            {
                var style = (column.ColumnStyle?.IsNotEmptyStyle ?? false) ? _style.CreateStyle(column.ColumnStyle) : _style.DefaultStyle;
                _sheet.SetDefaultColumnStyle(colIndex, style);
            }
            column.ColumnIndex = colIndex++;
        }
        #endregion

        var titles = columns.Select(a => a.Title).ToArray();
        AddTitleRow(rowIndex, startColIndex, titles);
        return this;
    }
    #endregion

    #region 添加主标题
    /// <summary>
    /// 添加主标题，如想调整样式，在调用该方法前调用ConfigMainTitleStyle
    /// </summary>
    public PoiExcel AddMainTitle(int rowIndex, string title, int colSpan = 10, int height = 25)
    {
        return AddMainTitle(rowIndex, 0, title, colSpan, height);
    }

    /// <summary>
    /// 添加主标题，如想调整样式，在调用该方法前调用ConfigMainTitleStyle
    /// </summary>
    public PoiExcel AddMainTitle(int rowIndex, int colIndex, string title, int colSpan = 10, int height = 25)
    {
        var row = PoiHelper.GetRow(_sheet, rowIndex);
        PoiHelper.SetRowHeight(row, height);
        var style = _style.CreateStyle(MainTitleStyle);
        PoiHelper.AddMergedRegion(_sheet, rowIndex, rowIndex, colIndex, colIndex + colSpan - 1, style);
        PoiHelper.GetCell(row, colIndex).SetCellValue(title);
        return this;
    }
    #endregion

    #region 添加备注说明
    /// <summary>
    /// 添加备注说明，如想调整样式，在调用该方法前调用ConfigRemarkStyle
    /// </summary>
    public PoiExcel AddRemarkRow(int rowIndex, string remark, int colSpan = 10, int height = 60)
    {
        return AddRemarkRow(rowIndex, 0, remark, colSpan, height);
    }

    /// <summary>
    /// 添加备注说明，如想调整样式，在调用该方法前调用ConfigRemarkStyle
    /// </summary>
    public PoiExcel AddRemarkRow(int rowIndex, int colIndex, string remark, int colSpan = 10, int height = 60)
    {
        var row = PoiHelper.GetRow(_sheet, rowIndex);
        PoiHelper.SetRowHeight(row, height);
        var style = _style.CreateStyle(RemarkStyle);
        PoiHelper.AddMergedRegion(_sheet, rowIndex, rowIndex, colIndex, colIndex + colSpan - 1, style);
        PoiHelper.GetCell(row, colIndex).SetCellValue(remark);
        return this;
    }
    #endregion

    #region 添加数据行
    /// <summary>
    /// 根据数据集合，自动添加标题和数据行
    /// <para>标题取第一个数据对象的属性名：第一条数据不能是空</para>
    /// </summary>
    /// <param name="dataList">数据集</param>
    /// <param name="rowIndex">开始行索引</param>
    /// <param name="colIndex">开始列索引</param>
    /// <param name="includeTitle">是否生成title，默认true</param>
    public PoiExcel AddDataList(int rowIndex, int colIndex, IEnumerable<object> dataList, bool includeTitle = true)
    {
        var lstData = dataList.ToList();
        if (lstData.Count == 0) return this;

        var props = lstData.First().GetType().GetProperties();

        if (includeTitle)
        {
            AddTitleRow(rowIndex, colIndex, [.. props.Select(a => a.Name)]);
            rowIndex += _titleRowCount;
        }
        foreach (var data in lstData)
        {
            var values = props.Select(a => a.GetValue(data)).ToArray();
            AddDataRow(rowIndex++, colIndex, values);
        }
        return this;
    }

    /// <summary>
    /// 根据数据集合，自动添加标题和数据行
    /// <para>标题根据列配置取，如果dataList没有数据，则根据列配置设置每列默认格式</para>
    /// </summary>
    public PoiExcel AddDataList(int rowIndex, int colIndex, IEnumerable<object> dataList, IEnumerable<ExcelColumn> columns)
    {
        if (!columns.Any()) return this;
        var hasDataRow = dataList.Any();
        // 标题
        var setDefaultValue = !hasDataRow;// 如果没有数据，则设置默认样式
        AddTitleRow(rowIndex, colIndex, columns, setDefaultValue);
        rowIndex += _titleRowCount;
        if (hasDataRow)
        {
            // 创建每列格式
            var styleMap = columns.ToDictionary(k => k.ColumnIndex, v => (v.ColumnStyle?.IsNotEmptyStyle ?? false) ? _style.CreateStyle(v.ColumnStyle) : null);
            // 添加数据 
            IRow row; ICell cell; ICellStyle? style;
            foreach (var data in dataList)
            {
                row = PoiHelper.GetRow(_sheet, rowIndex++);
                foreach (var c in columns)
                {
                    cell = PoiHelper.GetCell(row, c.ColumnIndex);
                    var value = c.ValueGetter(data);
                    style = styleMap[c.ColumnIndex];
                    if (style == null)
                    {
                        style = value != null && (value is DateTime || value is DateOnly)
                            ? _style.DateStyle
                            : _style.DefaultStyle;
                        styleMap[c.ColumnIndex] = style;
                    }
                    PoiHelper.SetCellValue(cell, value, false, style);
                }
            }
        }
        return this;
    }

    /// <summary>
    /// 根据DataTable 自动添加标题和数据 
    /// <para>标题取ColumnName</para>
    /// </summary>
    /// <param name="dt">数据集</param>
    /// <param name="rowIndex">开始行索引</param>
    /// <param name="colIndex">开始列索引</param>
    /// <param name="includeTitle">是否生成title，默认true</param>
    public PoiExcel AddDataTable(int rowIndex, int colIndex, DataTable dt, bool includeTitle = true)
    {
        if (includeTitle)
        {
            var titles = "";
            foreach (DataColumn dc in dt.Columns)
            {
                titles += dc.ColumnName + ",";
            }
            AddTitleRow(rowIndex, colIndex, titles.TrimEnd(','));
            rowIndex += _titleRowCount;
        }
        foreach (DataRow dr in dt.Rows)
        {
            AddDataRow(rowIndex++, colIndex, dr.ItemArray);
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
    public PoiExcel AddDataRow(int rowIndex, int colIndex, params object?[] values)
    {
        return SetCellValues(rowIndex, colIndex, values);
    }

    /// <summary>
    /// 添加空白行，主要是没内容时也没边框，可以调用该方法设置默认样式，只会处理空的单元格
    /// </summary>
    public PoiExcel AddBlankRow(int rowIndex, int colIndex, int colCount)
    {
        var row = PoiHelper.GetRow(_sheet, rowIndex);
        for (var i = colIndex; i < colIndex + colCount; i++)
        {
            var cell = row.GetCell(i);
            if (cell == null)
            {
                cell = row.CreateCell(i, CellType.Blank);
                cell.CellStyle = _style.DefaultStyle;
            }
        }
        return this;
    }

    public PoiExcel SetCellValues(int rowIndex, int colIndex, params object?[] values)
    {
        var row = PoiHelper.GetRow(_sheet, rowIndex);
        for (var i = 0; i < values.Length; i++)
        {
            var style = (values[i] is DateTime || values[i] is DateOnly) ? _style.DateStyle : _style.DefaultStyle;
            PoiHelper.SetCellValue(row, colIndex + i, values[i], false, style);
        }
        return this;
    }

    public PoiExcel SetCellValues(int rowIndex, int colIndex, ICellStyle style, params object?[] values)
    {
        var row = PoiHelper.GetRow(_sheet, rowIndex);
        for (var i = 0; i < values.Length; i++)
        {
            PoiHelper.SetCellValue(row, colIndex + i, values[i], false, style);
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
