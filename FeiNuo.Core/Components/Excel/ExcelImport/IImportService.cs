using Microsoft.Extensions.Primitives;

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
    public ImportConfig GetImportConfig(Dictionary<string, StringValues> paramMap, LoginUser user);
    #endregion

    /// <summary>
    /// 下载导入模板:前端接口直接调用该方法，如需完全自定义，重写该方法即可
    /// </summary>
    public PoiExcel GetImportTemplate(Dictionary<string, StringValues> paramMap, LoginUser user);

    /// <summary>
    /// 下载基础数据:前端接口直接调用该方法，如需完全自定义，重写该方法即可
    /// </summary>
    public Task<PoiExcel> GetImportBasicData(Dictionary<string, StringValues> paramMap, LoginUser user);

    /// <summary>
    /// 执行导入: 默认实现逻辑，保存文件，效验模板        
    /// </summary>
    public Task HandleImport(Stream stream, ImportConfig cfg, Dictionary<string, StringValues> paramMap, LoginUser user);
}
