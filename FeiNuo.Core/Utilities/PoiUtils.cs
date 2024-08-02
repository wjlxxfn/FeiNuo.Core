using NPOI.HSSF.UserModel;
using NPOI.SS.Formula.Eval;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System.Text;

namespace FeiNuo.Core.Utilities;

public class PoiUtils
{
    #region 创建工作簿，获取单元格
    /// <summary>
    /// 创建工作簿
    /// </summary>
    public static IWorkbook CreateWorkbook(bool xlsx = true)
    {
        return xlsx ? new XSSFWorkbook() : new HSSFWorkbook();
    }

    /// <summary>
    /// 根据文件流创建工作簿
    /// </summary>
    public static IWorkbook CreateWorkbook(Stream stream)
    {
        return WorkbookFactory.Create(stream);
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

    /// <summary>
    /// 创建单元格格式
    /// </summary>
    /// <param name="workbook">POI Workbook对象</param>
    /// <param name="borderStyle">边框格式，默认Thin</param>
    /// <param name="hAlign">水平位置，默认居左</param>
    /// <param name="vAlign">垂直位置，默认居中</param>
    /// <param name="bgColor">背景色</param>
    /// <param name="dataFormat">格式化</param>
    /// <param name="isBold">是否粗体</param>
    /// <param name="fontSize">字体大小</param>
    /// <param name="fontColor">字体颜色</param>
    /// <param name="wrapText">是否换行</param>
    /// <param name="source">源格式，新格式从source复制后在设置新的格式</param>
    /// <returns>ICellStyle</returns>
    public static ICellStyle CreateCellStyle(IWorkbook workbook, BorderStyle? borderStyle = null, HorizontalAlignment? hAlign = null, VerticalAlignment? vAlign = null, short? bgColor = null, string? dataFormat = null, bool? isBold = null, short? fontSize = null, short? fontColor = null, bool? wrapText = null, ICellStyle? source = null)
    {
        var style = workbook.CreateCellStyle();

        if (source != null) style.CloneStyleFrom(source);

        // 边框
        if (borderStyle.HasValue)
        {
            style.BorderBottom = borderStyle.Value;
            style.BorderLeft = borderStyle.Value;
            style.BorderRight = borderStyle.Value;
            style.BorderTop = borderStyle.Value;
        }

        // 位置
        if (vAlign.HasValue) style.VerticalAlignment = vAlign.Value;
        if (hAlign.HasValue) style.Alignment = hAlign.Value;

        // 字体
        if (fontSize.HasValue || isBold.HasValue || fontColor.HasValue)
        {
            var font = workbook.CreateFont();
            if (isBold.HasValue) font.IsBold = isBold.Value;
            if (fontSize.HasValue) font.FontHeightInPoints = fontSize.Value;
            if (fontColor.HasValue) font.Color = fontColor.Value;
            style.SetFont(font);
        }

        // 背景色
        if (bgColor.HasValue)
        {
            style.FillForegroundColor = bgColor.Value;
            style.FillPattern = FillPattern.SolidForeground;
            style.FillBackgroundColor = bgColor.Value;
        }

        // 格式化字符串
        if (!string.IsNullOrEmpty(dataFormat))
        {
            var fmt = workbook.CreateDataFormat();
            style.DataFormat = fmt.GetFormat(dataFormat);
        }

        // 是否换行
        if (wrapText.HasValue)
        {
            style.WrapText = wrapText.Value;
        }
        return style;
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
    public static ICell GetCell(ISheet sheet, int rowIndex, int colIndex)
    {
        return GetCell(GetRow(sheet, rowIndex), colIndex);
    }

    /// <summary>
    /// 获取单元格式，不存在的创建新单元格
    /// </summary>
    public static ICell GetCell(IRow row, int colIndex)
    {
        return CellUtil.GetCell(row, colIndex);
    }
    #endregion

    #region 单元格赋值
    /// <summary>
    /// 单元格赋值
    /// </summary>
    public static void SetCellValue(ICell cell, object value, bool isFormular = false, ICellStyle? cellStyle = null)
    {
        ArgumentNullException.ThrowIfNull(cell);

        if (cellStyle != null) cell.CellStyle = cellStyle;

        if (null == value || string.IsNullOrEmpty(value.ToString())) return;

        if (isFormular)
        {
            cell.SetCellFormula(value.ToString());
        }
        else
        {
            var type = value.GetType();
            if (type.Equals(typeof(decimal)) || type.Equals(typeof(double)) || type.Equals(typeof(float)))
            {
                cell.SetCellValue(Convert.ToDouble(value));
            }
            else if (type.Equals(typeof(int)))
            {
                cell.SetCellValue(Convert.ToInt32(value));
            }
            else if (type.Equals(typeof(DateOnly)))
            {
                cell.SetCellValue((DateOnly)value);
            }
            else if (value.GetType().Equals(typeof(DateTime)))
            {
                var dt = Convert.ToDateTime(value);
                if (dt != DateTime.MinValue && dt.Date != DateTime.Parse("1900-01-01"))
                {
                    cell.SetCellValue(dt);
                }
            }
            else
            {
                cell.SetCellValue(Convert.ToString(value));
            }
        }
    }
    #endregion

    #region 获取单元格的值
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

    #region 其它辅助方法
    /// <summary>
    /// 转成字节数组，方便传输
    /// </summary>
    public static byte[] GetExcelBytes(IWorkbook workboox, bool leaveOpen = false)
    {
        using var ms = new MemoryStream();
        workboox.Write(ms, leaveOpen);
        return ms.ToArray();
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
    #endregion
}
