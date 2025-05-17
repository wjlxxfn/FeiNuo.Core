using NPOI.HSSF.UserModel;
using NPOI.SS.Formula.Eval;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System.Data;
using System.Text;

namespace FeiNuo.Core;

public class PoiHelper
{
    #region 转换ExcelConfig
    /// <summary>
    /// 创建工作簿
    /// </summary>
    public static IWorkbook CreateWorkbook(ExcelConfig config, out StyleFactory styles)
    {
        // 检查配置内容
        config.ValidateConfigData();

        // 创建工作簿，生成样式工厂
        var wb = CreateWorkbook(config.IsExcel2007);
        styles = new StyleFactory(wb, config.DefaultStyle);

        // 创建工作表
        foreach (var sheet in config.ExcelSheets)
        {
            CreateWorkSheet(wb, sheet, styles);
        }

        return wb;
    }

    public static IWorkbook FillWorkbook(IWorkbook wb, ExcelConfig config, out StyleFactory styles)
    {
        // 检查配置内容
        config.ValidateConfigData();

        styles = new StyleFactory(wb, config.DefaultStyle);

        // 创建工作表
        foreach (var sheet in config.ExcelSheets)
        {
            CreateWorkSheet(wb, sheet, styles);
        }

        return wb;
    }

    /// <summary>
    /// 创建工作表:添加备注描述，主标题，列标题
    /// </summary>
    public static ISheet CreateWorkSheet(IWorkbook wb, ExcelSheet config, StyleFactory styles)
    {
        var sheet = wb.GetSheet(config.SheetName);
        sheet ??= wb.CreateSheet(config.SheetName);

        var columnCount = config.ColumnCount;
        int rowIndex = config.StartRowIndex;
        IRow row; ICell cell;

        #region 生成描述行
        if (!string.IsNullOrWhiteSpace(config.Description))
        {
            row = GetRow(sheet, rowIndex);
            row.HeightInPoints = (short)(config.DescriptionRowHeight < 0 ? 20 : config.DescriptionRowHeight);
            cell = GetCell(row, config.StartColumnIndex);
            cell.CellStyle = styles.GetStyle(config.DescriptionStyle);
            cell.SetCellValue(config.Description);

            var mergeCount = config.DescriptionColSpan ?? columnCount;
            if (mergeCount > 0)
            {
                sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, config.StartColumnIndex, config.StartColumnIndex + mergeCount - 1));
            }
            rowIndex++;
        }
        #endregion

        #region 生成主标题行
        if (!string.IsNullOrWhiteSpace(config.MainTitle))
        {
            row = GetRow(sheet, rowIndex);
            cell = GetCell(row, config.StartColumnIndex);
            cell.CellStyle = styles.GetStyle(config.MainTitleStyle);
            cell.SetCellValue(config.MainTitle);
            var mergeCount = config.MainTitleColSpan ?? columnCount;
            if (mergeCount > 0)
            {
                sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, config.StartColumnIndex, config.StartColumnIndex + mergeCount - 1));
            }
            rowIndex++;
        }
        #endregion

        #region 生成列标题行
        if (columnCount > 0)
        {
            var titleStyle = styles.GetStyle(config.ColumnTitleStyle);
            var titleRowCount = config.ExcelColumns.Max(a => a.RowTitles.Length);
            int titleRowStartIndex = rowIndex, titleRowEndIndex = rowIndex + titleRowCount - 1;
            foreach (var col in config.ExcelColumns)
            {
                rowIndex = titleRowStartIndex;

                // 设置默认格式
                if (col.ColumnStyle.IsNotEmptyStyle)
                {
                    sheet.SetDefaultColumnStyle(col.ColumnIndex, styles.GetStyle(col.ColumnStyle));
                }

                foreach (var rt in col.RowTitles)
                {
                    cell = GetCell(sheet, rowIndex++, col.ColumnIndex);
                    cell.CellStyle = titleStyle;
                    cell.SetCellValue(rt);
                }

                // 配置列宽，隐藏
                if (col.Width.HasValue) sheet.SetColumnWidth(col.ColumnIndex, col.Width.Value * 256);
                if (col.Hidden) sheet.SetColumnHidden(col.ColumnIndex, true);
            }
            //合并标题行
            for (var i = config.StartColumnIndex; i < config.StartColumnIndex + columnCount; i++)
            {
                AutoMergeRows(sheet, i, titleRowStartIndex, titleRowEndIndex);
            }
            // 合并标题列
            if (titleRowCount > 1)
            {
                for (var i = 0; i < titleRowCount; i++)
                {
                    AutoMergeColumns(sheet, i + config.TitleRowIndex, config.StartColumnIndex, config.StartColumnIndex + columnCount - 1);
                }
            }
            // 重置行索引
            rowIndex = titleRowEndIndex + 1;
        }
        #endregion

        #region 生成数据行
        if (config.DataList.Any())
        {
            foreach (var data in config.DataList)
            {
                row = GetRow(sheet, rowIndex++);
                foreach (var col in config.ExcelColumns)
                {
                    var val = col.ValueGetter?.Invoke(data);
                    var style = col.ColumnStyle.IsNotEmptyStyle
                        ? styles.GetStyle(col.ColumnStyle)
                        : ((val != null && (val is DateOnly || val is DateTime)) ? styles.DateStyle : null);
                    SetCellValue(row, col.ColumnIndex, val, false, style);
                }
            }
        }
        #endregion

        #region 工作表整体配置
        // 设置默认列宽
        if (config.DefaultColumnWidth.HasValue)
        {
            sheet.DefaultColumnWidth = config.DefaultColumnWidth.Value;
        }

        // 自动计算公式
        if (config.ForceFormulaRecalculation) sheet.ForceFormulaRecalculation = true;

        // 自动设置边框

        if (config.AddConditionalBorderStyle)
        {
            var region = new CellRangeAddress(config.StartRowIndex, sheet.LastRowNum, config.StartColumnIndex, config.EndColumnIndex);
            AddConditionalBorderStyle(sheet, range: region);
        }
        #endregion

        return sheet;
    }

    /// <summary>
    /// 根据模板配置，检查excel是否指定模板
    /// </summary>
    public static void ValidateExcelTemplate(IWorkbook wb, ExcelConfig config)
    {
        foreach (var sheet in config.ExcelSheets)
        {
            ValidateExcelTemplate(wb, sheet);
        }
    }

    /// <summary>
    /// 根据模板配置，检查excel是否指定模板
    /// </summary>
    public static void ValidateExcelTemplate(IWorkbook wb, ExcelSheet config)
    {
        if (!config.ValidateImportTemplate) return;
        var sheet = wb.GetSheet(config.SheetName) ?? throw new MessageException($"没有找到名为【{config.SheetName}】的Sheet页");
        if (!config.ExcelColumns.Any()) return;
        var row = GetRow(sheet, config.TitleRowIndex + config.ColumnRowCount - 1);
        foreach (var col in config.ExcelColumns)
        {
            if (GetStringValue(GetCell(row, col.ColumnIndex)) != col.RowTitles.Last())
            {
                throw new Exception($"【{config.SheetName}】模板错误，【{col.ColumnIndex}】列标题应为【{col.Title}】，请重新下载模板。");
            }
        }
    }

    /// <summary>
    /// 从Excel中读取数据，并赋值到ExcelSheet的DataList
    /// </summary>
    public static List<T> GetDataFromExcel<T>(IWorkbook wb, ExcelSheet config) where T : class, new()
    {
        var sheet = wb.GetSheet(config.SheetName) ?? throw new Exception($"找不到Sheet【{config.SheetName}】");
        // 计算公式
        sheet.ForceFormulaRecalculation = true;

        var keyMap = new Dictionary<string, int>();
        var lstData = new List<T>();
        var errMsg = new StringBuilder();

        for (var rowIndex = config.DataRowIndex; rowIndex <= sheet.LastRowNum; rowIndex++)
        {
            var row = GetRow(sheet, rowIndex);
            if (IsBlankRow(row)) continue;

            var data = Activator.CreateInstance<T>() ?? throw new Exception($"无法实例化对象{typeof(T)}");
            var rowMsg = new StringBuilder();
            var keyValue = new StringBuilder();
            foreach (var col in config.ExcelColumns)
            {
                var cell = GetCell(row, col.ColumnIndex);
                var val = GetCellValue(cell);

                try
                {
                    var msg = col.ValueSetter(data, val);
                    if (msg == "_UniqueKey_")
                    {
                        keyValue.Append((val ?? "null") + "|");
                    }
                    else if (!string.IsNullOrWhiteSpace(msg))
                    {
                        rowMsg.Append($"列【{col.RowTitles.Last()}】{msg}；");
                    }
                }
                catch (Exception ex)
                {
                    rowMsg.Append($"列【{col.RowTitles.Last()}】{ex.Message}；");
                }
            }

            if (rowMsg.Length > 0)
            {
                errMsg.AppendLine($"第【{rowIndex}】行：{rowMsg}");
            }
            else
            {
                var key = keyValue.ToString();
                if (!string.IsNullOrEmpty(key))
                {
                    if (keyMap.TryGetValue(key, out int existingRow))
                    {
                        errMsg.AppendLine($"第【{rowIndex}】行：与{existingRow}行重复,重复键值：{key}；");
                    }
                    else
                    {
                        keyMap[key] = rowIndex;
                    }
                }
                lstData.Add(data);
            }
        }

        if (errMsg.Length > 0)
        {
            throw new MessageException($"获取Excel数据出错:<br/>{errMsg}");
        }

        return lstData;
    }

    public static DataSet GetDataSetFromExcel(IWorkbook wb)
    {
        var ds = new DataSet();
        for (int i = 0; i < wb.NumberOfSheets; i++)
        {
            var sheet = wb.GetSheetAt(i);
            var dt = GetDataTableFromSheet(sheet);
            ds.Tables.Add(dt);
        }
        return ds;
    }

    public static DataTable GetDataTableFromSheet(ISheet sheet, int headerRowIndex = 0)
    {
        var dt = new DataTable(sheet.SheetName);
        var rowCount = sheet.LastRowNum + 1;
        for (int j = headerRowIndex; j < rowCount; j++)
        {
            var row = sheet.GetRow(j);
            if (row == null) continue;
            if (j == 0)
            {
                foreach (var cell in row.Cells)
                {
                    dt.Columns.Add(GetStringValue(cell));
                }
            }
            else
            {
                var dr = dt.NewRow();
                foreach (var cell in row.Cells)
                {
                    dr[cell.ColumnIndex] = GetCellValue(cell) ?? DBNull.Value;
                }
                dt.Rows.Add(dr);
            }
        }
        return dt;
    }

    /// <summary>
    /// Poi.SetDefaultColumnStyle后，如果使用POI赋值会把默认样式覆盖，调用该方法重新设置样式
    /// </summary>
    /// <param name="wb"></param>
    /// <param name="config"></param>
    /// <param name="styles"></param>
    public static void ResetColumnStyle(IWorkbook wb, ExcelConfig config, StyleFactory styles)
    {
        foreach (var sheetConfig in config.ExcelSheets)
        {
            var sheet = wb.GetSheet(sheetConfig.SheetName);
            if (sheet == null) continue;

            foreach (var col in sheetConfig.ExcelColumns)
            {
                if (col.ColumnStyle.IsEmptyStyle) continue;

                var cellStyle = styles.GetStyle(col.ColumnStyle);
                for (var r = sheetConfig.DataRowIndex; r <= sheet.LastRowNum; r++)
                {
                    var cell = sheet.GetRow(r)?.GetCell(col.ColumnIndex);
                    if (cell != null) cell.CellStyle = cellStyle;
                }
            }
        }
    }
    #endregion

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
            cell.SetCellValue(dto);
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
        if (cell.CellType == CellType.Numeric)
        {
            return cell.DateCellValue;
        }
        else if (cell.CellType == CellType.Formula)
        {
            if (DateUtil.IsCellDateFormatted(cell)) return cell.DateCellValue;
            else throw new Exception($"日期获取错误，CellType:{cell.CellType}，CellValue:{cell.ToString()}");
        }
        else
        {
            if (DateTime.TryParse(cell.StringCellValue.Trim('\''), out var dt)) return dt;
            else throw new Exception($"日期获取错误，CellType:{cell.CellType}，CellValue:{cell.ToString()}");
        }
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
        if (cell.CellType == CellType.Numeric)
        {
            return Convert.ToDecimal(cell.NumericCellValue);
        }
        if (cell.CellType == CellType.Formula)
        {
            if (cell.CachedFormulaResultType == CellType.Numeric) return Convert.ToDecimal(cell.NumericCellValue);
            if (cell.CachedFormulaResultType == CellType.String && decimal.TryParse(cell.StringCellValue, out var d)) return d;
        }
        if (decimal.TryParse(cell.ToString(), out var result))
        {
            return result;
        }
        else throw new Exception($"数值获取错误，CellType:{cell.CellType}，CellValue:{cell.ToString()}");
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
            case CellType.Unknown:
                throw new Exception("CellType Unknown!");
            default:
                return null;
        }
    }
    #endregion

    #region 其他辅助方法
    /// <summary>
    /// 按坐标合并单元格
    /// </summary>
    public static void AddMergedRegion(ISheet sheet, int startRow, int endRow, int startCol, int endCol, ICellStyle? style = null)
    {
        if (style != null)
        {
            for (int i = startRow; i <= endRow; i++)
            {
                var row = GetRow(sheet, i);
                for (int j = startCol; j <= endCol; j++)
                {
                    GetCell(row, j).CellStyle = style;
                }
            }
        }
        var cellRange = new CellRangeAddress(startRow, endRow, startCol, endCol);
        sheet.AddMergedRegion(cellRange);
    }

    /// <summary>
    /// 通过条件格式给内容区域的所有单元格添加边框,需要在添加完数据后调用
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
    /// 设置列宽
    /// </summary>
    public static void SetColumnWidth(ISheet sheet, int colIndex, int width)
    {
        sheet.SetColumnWidth(colIndex, width * 256d);
    }

    /// <summary>
    /// 设置行高
    /// </summary>
    public static void SetRowHeight(ISheet sheet, int rowIndex, int height)
    {
        var row = GetRow(sheet, rowIndex);
        row.Height = (short)(height * 20);
    }

    /// <summary>
    /// 生成IWorkbook,并转成字节数组
    /// </summary>
    public static byte[] GetExcelBytes(ExcelConfig config)
    {
        var wb = CreateWorkbook(config, out _);
        return GetExcelBytes(wb);
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
}
