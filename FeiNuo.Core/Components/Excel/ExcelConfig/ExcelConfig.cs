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
    /// Excel类型,2007/2003
    /// </summary>
    public ExcelType ExcelType { get; set; } = ExcelType.Excel2007;

    /// <summary>
    /// 工作表配置
    /// </summary>
    public List<ExcelSheet> ExcelSheets { get; set; } = [];

    /// <summary>
    /// 默认样式：水平自动，垂直居中
    /// </summary>
    public ExcelStyle DefaultStyle { get; set; } = new() { HorizontalAlignment = 1, VerticalAlignment = 1 };
    #endregion

    #region 构造函数
    /// <summary>
    /// 空模板的构造函数
    /// </summary>
    public ExcelConfig(string fileName, ExcelType? excelType = null, ExcelStyle? defaultStyle = null)
    {
        if (excelType.HasValue) ExcelType = excelType.Value;
        //如果没有后缀的话加上后缀
        FileName = fileName + (string.IsNullOrWhiteSpace(Path.GetExtension(fileName)) ? (IsExcel2007 ? ".xlsx" : ".xls") : "");
        if (defaultStyle != null) DefaultStyle = defaultStyle;
    }

    /// <summary>
    /// 适用于导出带数据的构造函数，
    /// </summary>
    /// <param name="fileName">文件名</param>
    /// <param name="DataList">要导出的数据</param>
    /// <param name="columns">列配置，使用ExcelColumn&lt;T&gt;对象</param>
    /// <param name="sheetConfig">工作表配置</param>
    public ExcelConfig(string fileName, IEnumerable<object> DataList, IEnumerable<ExcelColumn> columns, Action<ExcelSheet>? sheetConfig = null) : this(fileName)
    {
        AddExcelSheet("Sheet1", DataList, columns, sheetConfig);
    }

    /// <summary>
    /// 适用于导出空的Excel模板,带各列的配置
    /// </summary>
    /// <param name="fileName">文件名</param>
    /// <param name="columns">列配置</param>
    /// <param name="sheetConfig">工作表配置</param>
    public ExcelConfig(string fileName, IEnumerable<ExcelColumn> columns, Action<ExcelSheet>? sheetConfig = null) : this(fileName)
    {
        AddExcelSheet("Sheet1", columns, sheetConfig);
    }
    /// <summary>
    /// 适用于导出空的Excel模板,带各列标题
    /// </summary>
    /// <param name="fileName">文件名</param>
    /// <param name="titles">各列标题</param>
    public ExcelConfig(string fileName, params string[] titles) : this(fileName, titles.Select(a => new ExcelColumn(a)))
    {
    }
    #endregion

    #region 公共方法
    /// <summary>
    /// 验证配置数据是否有不合适的
    /// </summary>
    internal void ValidateConfigData()
    {
        var extension = Path.GetExtension(FileName).ToLower();
        if ((IsExcel2007 && extension != ".xlsx") || (!IsExcel2007 && extension != ".xls"))
        {
            throw new Exception("Excel版本和后缀不匹配");
        }
    }

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
    public ExcelConfig AddExcelSheet(string sheetName, IEnumerable<object> lstData, IEnumerable<ExcelColumn> columns, Action<ExcelSheet>? sheetConfig = null)
    {
        var excelSheet = new ExcelSheet(sheetName, lstData, columns);
        sheetConfig?.Invoke(excelSheet);
        return AddExcelSheet(excelSheet);
    }
    public ExcelConfig AddExcelSheet(string sheetName, params string[] titles)
    {
        return AddExcelSheet(sheetName, titles.Select(a => new ExcelColumn(a)));
    }
    #endregion

    /// <summary>
    /// 是否2007格式
    /// </summary>
    public bool IsExcel2007 { get { return ExcelType == ExcelType.Excel2007; } }

    /// <summary>
    /// contentType
    /// </summary>
    public string ContentType { get { return IsExcel2007 ? "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" : "application/vnd.ms-excel"; } }
    #endregion
}

/// <summary>
/// Excel版本
/// </summary>
public enum ExcelType
{
    Excel2007,
    Excel2003
}