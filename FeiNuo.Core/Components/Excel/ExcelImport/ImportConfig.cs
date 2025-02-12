namespace FeiNuo.Core;

/// <summary>
/// Excel导入配置类
/// </summary>
public class ImportConfig
{
    #region 构造方法
    public ImportConfig(string templateName = "导入模板.xlsx", string? basicDataName = null, string? importRemark = null)
    {
        TemplateName = templateName;
        BasicDataName = basicDataName ?? "";
        ImportRemark = importRemark ?? "";
    }
    #endregion

    #region 导入配置
    /// <summary>
    /// 导入所需的角色
    /// </summary>
    public string[] AuthRoles { get; set; } = [];
    #endregion

    #region 导入模板
    public string TemplateName { get; private set; } = "导入模板.xlsx";
    public bool ShowTemplate { get { return !string.IsNullOrWhiteSpace(TemplateName); } }
    #endregion

    #region 基础数据
    public string BasicDataName { get; set; } = "基础数据.xlsx";
    public bool ShowBasicData { get { return !string.IsNullOrWhiteSpace(BasicDataName); } }
    #endregion

    #region 导入说明
    public string ImportRemark { get; set; } = string.Empty;
    public bool ShowRemark { get { return !string.IsNullOrWhiteSpace(ImportRemark); } }
    #endregion

    #region 保存文件
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

    /// <summary>
    /// 传参用：默认为空，在上传excel时，系统给赋值，避免在执行逻辑时在查一次模板
    /// </summary>
    public ExcelConfig? ImportTemplate { get; internal set; }
}
