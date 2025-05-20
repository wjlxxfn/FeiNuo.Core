namespace FeiNuo.Core;

/// <summary>
/// Excel配置类。配置后使用ExcelHelper.CreateWorkBook方法可生成IWorkbook对象
/// </summary>
public class ExcelConfig
{
    #region 属性定义
    /// <summary>
    /// 文件名
    /// </summary>
    public string FileName { get; set; }

    /// <summary>
    /// 工作表配置
    /// </summary>
    public List<ExcelSheet> ExcelSheets { get; set; } = [];

    /// <summary>
    /// 构造函数
    /// </summary>
    public ExcelConfig(string fileName)
    {
        if (string.IsNullOrWhiteSpace(fileName))
        {
            throw new ArgumentNullException(nameof(fileName));
        }

        FileName = fileName;
        var fileType = Path.GetExtension(fileName);
        if (string.IsNullOrWhiteSpace(fileType))
        {
            FileName += ".xlsx";
        }
        else if (!fileType.Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
        {
            throw new Exception("只支持xlsx文件");
        }
    }
    #endregion

    #region 公共方法   
    /// <summary>
    /// 添加工作表
    /// </summary>
    public ExcelConfig AddExcelSheet(ExcelSheet excelSheet)
    {
        if (string.IsNullOrWhiteSpace(excelSheet.SheetName))
        {
            throw new MessageException("Sheet名不能为空");
        }
        if (ExcelSheets.Any(t => t.SheetName == excelSheet.SheetName))
        {
            throw new MessageException($"已存在名为【{excelSheet.SheetName}】的工作表。");
        }
        ExcelSheets.Add(excelSheet);
        return this;
    }

    /// <summary>
    /// 清空所有工作表
    /// </summary>
    public ExcelConfig ClearSheets()
    {
        ExcelSheets.Clear();
        return this;
    }
    #endregion

    #region 重载AddExcelSheet，方便各种场景下的调用
    /// <summary>
    /// 添加工作表
    /// </summary>
    /// <param name="columns">列配置</param>
    /// <param name="sheetConfig"></param>
    /// <returns></returns>
    public ExcelConfig AddExcelSheet(IEnumerable<ExcelColumn> columns, Action<ExcelSheet>? sheetConfig = null)
    {
        return AddExcelSheet("Sheet1", columns, sheetConfig);
    }
    public ExcelConfig AddExcelSheet(string sheetName, IEnumerable<ExcelColumn> columns, Action<ExcelSheet>? sheetConfig = null)
    {
        var excelSheet = new ExcelSheet(sheetName, columns);
        sheetConfig?.Invoke(excelSheet);
        return AddExcelSheet(excelSheet);
    }
    #endregion
}

