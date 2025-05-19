using NPOI.HSSF.UserModel;
using NPOI.SS.Formula.Eval;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System.Data;
using System.Text;

namespace FeiNuo.Core;

public partial class PoiHelper
{
    #region 创建POI对象
    /// <summary>
    /// 创建工作簿
    /// </summary>
    public static IWorkbook CreateWorkbook(bool xlsx = true)
    {
        return xlsx ? new XSSFWorkbook() : new HSSFWorkbook();
    }

    /// <summary>
    /// 创建工作簿
    /// </summary>
    public static IWorkbook CreateWorkbook(Stream stream)
    {
        return WorkbookFactory.Create(stream);
    }

    /// <summary>
    /// 获取行，不存在的创建一行
    /// </summary>
    /// <param name="sheet"></param>
    /// <param name="rowIndex"></param>
    /// <returns></returns>
    public static IRow GetRow(ISheet sheet, int rowIndex)
    {
        return CellUtil.GetRow(rowIndex, sheet);
    }

    /// <summary>
    /// 获取单元格式，不存在的创建新单元格
    /// </summary>
    public static ICell GetCell(ISheet sheet, int rowIndex, int colIndex, ICellStyle? style = null)
    {
        return GetCell(GetRow(sheet, rowIndex), colIndex, style);
    }

    /// <summary>
    /// 获取单元格式，不存在的创建新单元格
    /// </summary>
    public static ICell GetCell(IRow row, int colIndex, ICellStyle? style = null)
    {
        var cell = CellUtil.GetCell(row, colIndex);
        if (style != null) cell.CellStyle = style;
        return cell;
    }

    /// <summary>
    /// 创建富文本字符串
    /// </summary>
    public static IRichTextString CreateRichTextString(string text, bool xlsx = true)
    {
        return xlsx ? new XSSFRichTextString(text) : new HSSFRichTextString(text);
    }

    /// <summary>
    /// 创建锚点
    /// </summary>
    public static IClientAnchor CreateClientAnchor(int dx1, int dy1, int dx2, int dy2, int col1, int row1, int col2, int row2, bool xlsx = true)
    {
        return xlsx ? new XSSFClientAnchor(dx1, dy1, dx2, dy2, col1, row1, col2, row2) : new HSSFClientAnchor(dx1, dy1, dx2, dy2, col1, row1, col2, row2);
    }
    #endregion

    #region 单元格赋值
    public static void SetCellValues(ISheet sheet, int rowIndex, int startColIndex, params object?[] values)
    {
        var row = GetRow(sheet, rowIndex);
        SetCellValues(row, startColIndex, values);
    }
    public static void SetCellValues(IRow row, int startColIndex, params object?[] values)
    {
        if (values == null || values.Length == 0) return;
        foreach (var val in values)
        {
            SetCellValue(row, startColIndex++, val, false);
        }
    }
    public static void SetCellValues(IRow row, int startColIndex, ICellStyle style, params object?[] values)
    {
        if (values == null || values.Length == 0) return;
        foreach (var val in values)
        {
            SetCellValue(row, startColIndex++, val, false, style);
        }
    }

    public static void SetCellValue(ISheet sheet, int rowIndex, int colIndex, object? value, bool isFormular = false, ICellStyle? style = null)
    {
        var row = GetRow(sheet, rowIndex);
        SetCellValue(row, colIndex, value, isFormular, style);
    }
    public static void SetCellValue(IRow row, int colIndex, object? value, bool isFormular = false, ICellStyle? style = null)
    {
        var cell = GetCell(row, colIndex);
        SetCellValue(cell, value, isFormular, style);
    }
    public static void SetCellValue(ICell cell, object? value, bool isFormular = false, ICellStyle? style = null)
    {
        ArgumentNullException.ThrowIfNull(cell);

        if (style != null) cell.CellStyle = style;

        if (value == null) return;

        if (isFormular)
        {
            cell.SetCellFormula(value.ToString());
            return;
        }

        if (value is int || value is short || value is long || value is decimal || value is double || value is float)
        {
            cell.SetCellValue(Convert.ToDouble(value));
        }
        else if (value is DateOnly dto)
        {
            if (dto != DateOnly.MinValue && dto.ToString("yyyy-MM-dd") != "1900-01-01")
            {
                cell.SetCellValue(dto);
            }
        }
        else if (value is DateTime dt)
        {
            if (dt != DateTime.MinValue && dt.Date.ToString("yyyy-MM-dd") != "1900-01-01")
            {
                cell.SetCellValue(dt);
            }
        }
        else
        {
            cell.SetCellValue(Convert.ToString(value));
        }
    }
    #endregion

    #region 单元格取值
    /// <summary>
    /// 获取单元格的值，空值返回空字符串
    /// </summary>
    /// <param name="cell"></param>
    public static string? GetStringValue(ICell cell)
    {
        if (null == cell || cell.CellType == CellType.Blank)
        {
            return null;
        }

        if (cell.CellType == CellType.String)
        {
            return cell.StringCellValue;
        }
        if (cell.CellType == CellType.Formula)
        {
            if (cell.CachedFormulaResultType == CellType.String) return cell.StringCellValue;
            if (cell.CachedFormulaResultType == CellType.Numeric) return cell.NumericCellValue.ToString();
        }
        return cell.ToString();
    }

    /// <summary>
    /// 获取单元格的值，空值返回空字符串
    /// </summary>
    /// <param name="cell"></param>
    /// <param name="defaultValue"></param>
    public static string GetStringValue(ICell cell, string defaultValue)
    {
        return GetStringValue(cell) ?? defaultValue;
    }

    /// <summary>
    /// 获取单元格的日期值
    /// </summary>
    public static DateTime? GetDateValue(ICell cell)
    {
        if (null == cell || cell.CellType == CellType.Blank)
        {
            return null;
        }

        var cellType = cell.CellType == CellType.Formula ? cell.CachedFormulaResultType : cell.CellType;
        if (cellType == CellType.Numeric)
        {
            return cell.DateCellValue;
        }
        else if (cellType == CellType.String)
        {
            if (DateTime.TryParse(cell.StringCellValue, out var val)) return val;
            else throw new Exception("日期错误: " + Cts(cell));
        }
        else if (DateTime.TryParse(cell.ToString(), out var val))
        {
            return val;
        }
        else throw new Exception("日期错误: " + Cts(cell));
    }

    /// <summary>
    /// 获取单元格的日期值
    /// </summary>
    public static DateTime GetDateValue(ICell cell, DateTime defaultValue)
    {
        return GetDateValue(cell) ?? defaultValue;
    }

    /// <summary>
    /// 获取单元格的Decimal值
    /// </summary>
    /// <param name="cell"></param>
    public static decimal? GetDecimalValue(ICell cell)
    {
        if (null == cell || cell.CellType == CellType.Blank)
        {
            return null;
        }

        var cellType = cell.CellType == CellType.Formula ? cell.CachedFormulaResultType : cell.CellType;
        if (cellType == CellType.Numeric)
        {
            return Convert.ToDecimal(cell.NumericCellValue);
        }
        else if (cellType == CellType.String)
        {
            if (decimal.TryParse(cell.StringCellValue, out var val)) return val;
            else throw new Exception("数值错误: " + Cts(cell));
        }
        else if (decimal.TryParse(cell.ToString(), out var val))
        {
            return val;
        }
        else throw new Exception("数值错误: " + Cts(cell));
    }

    /// <summary>
    /// 获取单元格的Decimal值
    /// </summary>
    /// <param name="cell"></param>
    /// <param name="defaultValue"></param>
    /// <returns></returns>
    public static decimal GetDecimalValue(ICell cell, decimal defaultValue)
    {
        return GetDecimalValue(cell) ?? defaultValue;
    }

    private static string Cts(ICell cell)
    {
        return $"单元格={cell.Address}, 值={cell}";
    }

    /// <summary>
    /// 获取单元格的值
    /// </summary>
    public static T? GetCellValue<T>(ICell cell)
    {
        var value = GetCellValue(cell);
        if (null == value) return (T?)value;
        return (T)Convert.ChangeType(value, typeof(T));
    }

    /// <summary>
    /// 获取单元格的值
    /// </summary>               
    public static IConvertible? GetCellValue(ICell cell)
    {
        if (null == cell || cell.CellType == CellType.Blank)
        {
            return null;
        }
        // 公式的默认已提前计算
        var cellType = cell.CellType == CellType.Formula ? cell.CachedFormulaResultType : cell.CellType;

        switch (cellType)
        {
            case CellType.Boolean:
                return cell.BooleanCellValue;
            case CellType.Error:
                return ErrorEval.GetText(cell.ErrorCellValue);
            case CellType.Numeric:
                if (DateUtil.IsCellDateFormatted(cell))
                {
                    return cell.DateCellValue;
                }
                else
                {
                    return cell.NumericCellValue;
                }
            case CellType.String:
                return cell.StringCellValue.Trim();
            default:
                throw new Exception($"格式错误: " + Cts(cell));
        }
    }
    #endregion

    #region 其他辅助方法
    /// <summary>
    /// 按坐标合并单元格
    /// <para>如果没有传入格式，则合并后的单元格以第一个单元格的格式为准</para>
    /// </summary>
    public static void AddMergedRegion(ISheet sheet, int startRow, int endRow, int startCol, int endCol, ICellStyle? style = null)
    {
        if (style == null)
        {
            style = GetCell(sheet, startRow, startCol).CellStyle;
        }
        for (int i = startRow; i <= endRow; i++)
        {
            var row = GetRow(sheet, i);
            for (int j = startCol; j <= endCol; j++)
            {
                GetCell(row, j).CellStyle = style;
            }
        }
        var cellRange = new CellRangeAddress(startRow, endRow, startCol, endCol);
        sheet.AddMergedRegion(cellRange);
    }

    /// <summary>
    /// 通过条件格式给内容区域的所有单元格添加边框
    /// <para>如果没有执行范围,默认有数据的范围，但需要在添加完数据后调用</para>
    /// </summary>
    /// <param name="sheet">工作表</param>
    /// <param name="borderStyle">边框样式，默认Thin</param>
    /// <param name="range">范围，默认从A1开始，最大列取前三行有内容的最后一列的列索引，最大行取当前sheet的最后一行</param>
    public static void AddConditionalBorderStyle(ISheet sheet, BorderStyle borderStyle = BorderStyle.Thin, CellRangeAddress? range = null)
    {
        var scf = sheet.SheetConditionalFormatting;
        var rule = scf.CreateConditionalFormattingRule("TRUE()");
        var fmt = rule.CreateBorderFormatting();
        fmt.BorderBottom = borderStyle;
        fmt.BorderRight = borderStyle;
        fmt.BorderLeft = borderStyle;
        fmt.BorderTop = borderStyle;

        if (range == null)
        {
            //取前三行的最大列索引
            var colIndex = GetRow(sheet, 0).LastCellNum;
            if (sheet.GetRow(1) != null && sheet.GetRow(1).LastCellNum > colIndex) colIndex = sheet.GetRow(1).LastCellNum;
            if (sheet.GetRow(2) != null && sheet.GetRow(2).LastCellNum > colIndex) colIndex = sheet.GetRow(2).LastCellNum;
            if (colIndex <= 0) colIndex = 1;

            range = CellRangeAddress.ValueOf($"A1:{GetColumnName(colIndex - 1)}{(sheet.LastRowNum + 1)}");
        }
        scf.AddConditionalFormatting([range], rule);
    }

    /// <summary>
    /// 创建单元格样式
    /// </summary>
    internal static ICellStyle CreateCellStyle(IWorkbook workbook, ExcelStyle config, ICellStyle? source = null)
    {
        var style = workbook.CreateCellStyle();
        if (null != source) style.CloneStyleFrom(source);

        // 边框
        if (config.BorderStyle.HasValue)
        {
            if (Enum.TryParse(config.BorderStyle.Value.ToString(), true, out BorderStyle borderStyle))
            {
                style.BorderBottom = borderStyle;
                style.BorderLeft = borderStyle;
                style.BorderRight = borderStyle;
                style.BorderTop = borderStyle;
            }
            else throw new Exception($"边框样式值【{config.BorderStyle.Value}】不是有效的POI样式。");
        }

        // 位置
        if (config.VerticalAlignment.HasValue)
        {
            if (Enum.TryParse(config.VerticalAlignment.Value.ToString(), true, out VerticalAlignment vAlign))
            {
                style.VerticalAlignment = vAlign;
            }
            else throw new Exception($"垂直位置值【{config.VerticalAlignment.Value}】不是有效的POI样式。");
        }
        if (config.HorizontalAlignment.HasValue)
        {
            if (Enum.TryParse(config.HorizontalAlignment.Value.ToString(), true, out HorizontalAlignment hAlign))
            {
                style.Alignment = hAlign;
            }
            else throw new Exception($"水平位置值【{config.HorizontalAlignment.Value}】不是有效的POI样式。");
        }


        // 背景色
        if (config.BackgroundColor.HasValue)
        {
            style.FillForegroundColor = config.BackgroundColor.Value;
            style.FillPattern = FillPattern.SolidForeground;
            style.FillBackgroundColor = config.BackgroundColor.Value;
        }

        // 格式化字符串
        if (!string.IsNullOrWhiteSpace(config.DataFormat))
        {
            var fmt = workbook.CreateDataFormat();
            style.DataFormat = fmt.GetFormat(config.DataFormat);
        }

        // 换行
        if (config.WrapText.HasValue)
        {
            style.WrapText = config.WrapText.Value;
        }

        // 字体
        if (!string.IsNullOrWhiteSpace(config.FontName) || config.FontSize.HasValue || config.FontBold.HasValue || config.FontColor.HasValue)
        {
            var font = workbook.CreateFont();
            if (!string.IsNullOrWhiteSpace(config.FontName)) font.FontName = config.FontName;
            if (config.FontBold.HasValue) font.IsBold = config.FontBold.Value;
            if (config.FontSize.HasValue) font.FontHeightInPoints = config.FontSize.Value;
            if (config.FontColor.HasValue) font.Color = config.FontColor.Value;
            style.SetFont(font);
        }

        return style;
    }

    /// <summary>
    /// 将Excel的列索引转换为列名，列索引从0开始，列名从A开始。如第0列为A，第1列为B...
    /// </summary>
    /// <param name="columnIndex">列索引</param>
    /// <returns>列名，如第0列为A，第1列为B...</returns>
    public static string GetColumnName(int columnIndex)
    {
        columnIndex++;
        int system = 26;
        char[] digArray = new char[100];
        int i = 0;
        while (columnIndex > 0)
        {
            int mod = columnIndex % system;
            if (mod == 0) mod = system;
            digArray[i++] = (char)(mod - 1 + 'A');
            columnIndex = (columnIndex - 1) / 26;
        }
        var sb = new StringBuilder(i);
        for (int j = i - 1; j >= 0; j--)
        {
            sb.Append(digArray[j]);
        }
        return sb.ToString();
    }

    /// <summary>
    /// 判断该行是否有数据
    /// </summary>
    public static bool IsBlankRow(IRow row)
    {
        if (row == null) return true;
        foreach (var cell in row.Cells)
        {
            if (cell != null && cell.CellType != CellType.Blank && !string.IsNullOrWhiteSpace(cell.ToString()))
            {
                return false;
            }
        }
        return true;
    }

    /// <summary>
    /// 指定列，自动合并内容相同的行
    /// </summary>
    /// <param name="sheet"></param>
    /// <param name="colIndex"></param>
    /// <param name="startRow"></param>
    /// <param name="endRow"></param>
    public static void AutoMergeRows(ISheet sheet, int colIndex, int startRow = 0, int endRow = 0)
    {
        ArgumentNullException.ThrowIfNull(sheet);
        if (endRow == 0) endRow = sheet.LastRowNum;

        for (int i = startRow; i <= endRow; i++)
        {
            IRow currentRow = sheet.GetRow(i);
            IRow nextRow = sheet.GetRow(i + 1);

            if (currentRow == null || nextRow == null) continue;

            ICell currentCell = currentRow.GetCell(colIndex);
            ICell nextCell = nextRow.GetCell(colIndex);

            if (currentCell == null || nextCell == null) continue;

            string currentValue = currentCell?.ToString() ?? "";
            string nextValue = nextCell?.ToString() ?? "";

            if (currentValue == nextValue)
            {
                int mergeStartRow = i;
                while (i + 1 <= endRow && currentValue == nextValue)
                {
                    i++;
                    nextRow = sheet.GetRow(i + 1);
                    if (nextRow == null) break;

                    nextCell = nextRow.GetCell(colIndex);
                    if (nextCell == null) break;

                    nextValue = nextCell?.ToString() ?? "";
                }
                int mergeEndRow = i;

                if (mergeStartRow < mergeEndRow)
                {
                    sheet.AddMergedRegion(new CellRangeAddress(mergeStartRow, mergeEndRow, colIndex, colIndex));
                }
            }
        }
    }

    /// <summary>
    /// 根据指定行，自动合并内容相同的列
    /// </summary>
    /// <param name="sheet"></param>
    /// <param name="rowIndex">要合并的行</param>
    /// <param name="startCol">开始列</param>
    /// <param name="endCol">结束列</param>
    /// <exception cref="ArgumentNullException"></exception>
    public static void AutoMergeColumns(ISheet sheet, int rowIndex, int startCol = 0, int endCol = 0)
    {
        ArgumentNullException.ThrowIfNull(sheet);

        IRow row = sheet.GetRow(rowIndex);
        if (row == null) return;

        if (endCol == 0) endCol = row.LastCellNum - 1;

        while (startCol <= endCol)
        {
            ICell currentCell = row.GetCell(startCol);
            if (currentCell == null)
            {
                startCol++;
                continue;
            }

            string currentValue = currentCell?.ToString() ?? "";
            int mergeStartCol = startCol;
            int mergeEndCol = startCol;

            // 查找需要合并的列
            for (int i = startCol + 1; i <= endCol; i++)
            {
                ICell nextCell = row.GetCell(i);
                if (nextCell == null) break;

                string nextValue = nextCell?.ToString() ?? "";
                if (currentValue == nextValue)
                {
                    mergeEndCol = i;
                }
                else
                {
                    break;
                }
            }

            // 合并单元格
            if (mergeStartCol < mergeEndCol)
            {
                sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, mergeStartCol, mergeEndCol));
            }

            startCol = mergeEndCol + 1;
        }
    }

    /// <summary>
    /// 隐藏列
    /// </summary>
    public static void SetColumnHidden(ISheet sheet, int colIndex, bool hidden = true)
    {
        sheet.SetColumnHidden(colIndex, hidden);
    }

    /// <summary>
    /// 设置列宽
    /// </summary>
    public static void SetColumnWidth(ISheet sheet, int colIndex, int width)
    {
        sheet.SetColumnWidth(colIndex, width * 256d);
    }

    /// <summary>
    /// 设置自动列宽
    /// </summary>
    public static void AutoSizeColumn(ISheet sheet, int colIndex)
    {
        sheet.AutoSizeColumn(colIndex);
    }

    /// <summary>
    /// 设置行高
    /// </summary>
    public static void SetRowHeight(ISheet sheet, int rowIndex, int height)
    {
        var row = GetRow(sheet, rowIndex);
        SetRowHeight(row, height);
    }

    /// <summary>
    /// 设置行高
    /// </summary>
    public static void SetRowHeight(IRow row, int height)
    {
        row.Height = (short)(height * 20);
    }

    /// <summary>
    /// 转成字节数组
    /// </summary>
    public static byte[] GetExcelBytes(IWorkbook workboox, bool leaveOpen = false)
    {
        using var ms = new MemoryStream();
        workboox.Write(ms, leaveOpen);
        return ms.ToArray();
    }
    #endregion

    #region 创建PoiExcel
    /// <summary>
    /// 创建空白excel，默认会创建一个Sheet1，如果不需要，创建后调用PoiExcel.RemoveSheet1()删除
    /// </summary>
    public static PoiExcel CreateExcel(ExcelStyle? defaultStyle = null)
    {
        return CreateExcel(out _, defaultStyle);
    }

    /// <summary>
    /// 创建空白excel
    /// </summary>
    public static PoiExcel CreateExcel(out StyleFactory style, ExcelStyle? defaultStyle = null)
    {
        var wb = PoiHelper.CreateWorkbook();
        var sheet = wb.CreateSheet("Sheet1");
        style = new StyleFactory(wb, defaultStyle);
        return new PoiExcel(wb, sheet, style);
    }

    /// <summary>
    /// 根据文件流创建Excel对象，当前Sheet默认设置为第一个Sheet
    /// </summary>
    public static PoiExcel CreateExcel(Stream stream)
    {
        return CreateExcel(stream, out _);
    }

    /// <summary>
    /// 根据文件流创建Excel对象，当前Sheet默认设置为第一个Sheet
    /// </summary>
    public static PoiExcel CreateExcel(Stream stream, out StyleFactory style)
    {
        var wb = PoiHelper.CreateWorkbook(stream);
        var sheet = wb.NumberOfSheets >= 1 ? wb.GetSheetAt(0) : wb.CreateSheet("Sheet1");
        style = new StyleFactory(wb);
        return new PoiExcel(wb, sheet, style);
    }

    /// <summary>
    /// 根据数据集合构造Excel对象
    /// <para>标题取第一个数据对象的属性名：第一条数据不能是空</para>
    /// <para>使用示例: new PoiExcel(users.Select(a=>new {用户名=a.Username,姓名=a.NickName}))</para>
    /// </summary>
    /// <param name="dataList">数据</param>
    public static PoiExcel CreateExcel(IEnumerable<object> dataList)
    {
        return CreateExcel(dataList, out _);
    }

    /// <summary>
    /// 根据数据集合构造Excel对象
    /// <para>标题取第一个数据对象的属性名：第一条数据不能是空</para>
    /// <para>使用示例: new PoiExcel(users.Select(a=>new {用户名=a.Username,姓名=a.NickName}))</para>
    /// </summary>
    public static PoiExcel CreateExcel(IEnumerable<object> dataList, out StyleFactory style)
    {
        var poi = CreateExcel(out style);
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
        var poi = CreateExcel().RemoveSheet1();
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
        var poi = CreateExcel();
        poi.AddDataTable(dt);
        return poi;
    }

    /// <summary>
    /// 根据DataTable构建Excel，SheetName=TableName,标题=ColumnName
    /// </summary>
    public static PoiExcel CreateExcel(DataSet ds)
    {
        var poi = CreateExcel().RemoveSheet1();
        foreach (DataTable dt in ds.Tables)
        {
            poi.CreateSheet(dt.TableName);
            poi.AddDataTable(dt);
        }
        return poi;
    }

    /// <summary>
    /// 根据列配置自动生成Excel,通常是生成导入模板，内部设置各列的默认格式
    /// </summary>
    public static PoiExcel CreateExcel<T>(IEnumerable<ExcelColumn<T>> columns) where T : class
    {
        // 这里用空样式，不然默认样式带边框，列默认导致导出的模板全部是边框
        //TODO 测试这里用默认样式
        var poi = CreateExcel(ExcelStyle.EmptyStyle);
        poi.AddTitleRow(columns, true);
        return poi;
    }

    /// <summary>
    /// 根据数据集和列配置自动生成Excel
    /// </summary>
    public static PoiExcel CreateExcel<T>(IEnumerable<T> dataList, IEnumerable<ExcelColumn<T>> columns) where T : class
    {
        var poi = CreateExcel();
        poi.AddDataList(dataList, columns);
        return poi;
    }
    #endregion
}
