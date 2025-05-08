namespace FeiNuo.Core;

/// <summary>
/// Excel导入配置类
/// </summary>
public class ImportConfig
{
    #region 导入模板
    /// <summary>
    /// 导入模板
    /// </summary>
    public ExcelConfig? ImportTemplate { get; set; }
    public bool ShowTemplate => ImportTemplate != null;
    public string TemplateName => ImportTemplate?.FileName ?? string.Empty;
    #endregion

    #region 导入说明
    public string ImportRemark { get; set; } = string.Empty;
    public bool ShowRemark { get { return !string.IsNullOrWhiteSpace(ImportRemark); } }
    #endregion

    #region 基础数据
    public bool ShowBasicData { get; set; } = false;
    public string BasicDataName { get; set; } = "基础数据.xlsx";
    #endregion

    #region 其他配置
    /// <summary>
    /// 导入所需的角色
    /// </summary>
    public string[] AuthRoles { get; set; } = [];

    /// <summary>
    /// 导入前是否保存excel文件
    /// </summary>
    public bool SaveExcel { get; set; } = false;

    /// <summary>
    /// excel保存路径,默认根路径下Attachment/ExcelImport
    /// 配置路径后，文件会在该路径下创建子路径保存文件 /ImportType/yyyy-mm/guid/filename.
    /// </summary>
    public string SavePath { get; set; } = "Attachment/ExcelImport";
    #endregion
}
