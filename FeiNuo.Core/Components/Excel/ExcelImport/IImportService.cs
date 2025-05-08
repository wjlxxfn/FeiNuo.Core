namespace FeiNuo.Core;

/// <summary>
/// Excel导入实现接口
/// </summary>
public interface IImportService
{
    #region 获取导入配置
    /// <summary>
    /// 获取导入类型：不能为空必须唯一
    /// </summary>
    public string GetImportKey();

    /// <summary>
    /// 获取导入配置
    /// </summary>
    public ImportConfig GetImportConfig(Dictionary<string, string> paramMap, LoginUser user);
    #endregion

    #region 下载导入模板
    /// <summary>
    /// 下载导入模板:前端接口直接调用该方法，如需完全自定义，重写该方法即可
    /// </summary>
    public ExcelDownload DownloadTemplate(Dictionary<string, string> paramMap, LoginUser user)
    {
        var config = GetImportConfig(paramMap, user);
        var excel = config.ImportTemplate ?? new ExcelConfig("导入模板.xlsx");
        var bytes = PoiHelper.GetExcelBytes(excel);
        return new ExcelDownload(excel.FileName, excel.ContentType, bytes);
    }
    #endregion

    #region 下载基础数据
    /// <summary>
    /// 下载基础数据
    /// </summary>
    public async Task<ExcelDownload> DownloadBasicData(Dictionary<string, string> paramMap, LoginUser user)
    {
        var excel = await GetImportBasicData(paramMap, user);
        var bytes = PoiHelper.GetExcelBytes(excel);
        return new ExcelDownload(excel.FileName, excel.ContentType, bytes);
    }

    /// <summary>
    /// 下载基础数据
    /// </summary>
    public Task<ExcelConfig> GetImportBasicData(Dictionary<string, string> paramMap, LoginUser user);
    #endregion

    /// <summary>
    /// 执行导入: 默认实现逻辑，保存文件，效验模板        
    /// </summary>
    public Task HandleImport(Stream stream, ImportConfig cfg, Dictionary<string, string> paramMap, LoginUser user);

}

/// <summary>
/// 下载excel的辅助类
/// </summary>
public class ExcelDownload
{
    public string FileName { get; set; }
    public string ContentType { get; set; }
    public byte[] Bytes { get; set; }

    public ExcelDownload(string fileName, string contentType, byte[] bytes)
    {
        FileName = fileName;
        ContentType = contentType;
        Bytes = bytes;
    }
}
