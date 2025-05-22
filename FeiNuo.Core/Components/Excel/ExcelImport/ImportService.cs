using Microsoft.Extensions.Primitives;
using NPOI.SS.UserModel;

namespace FeiNuo.Core;

public abstract class BaseImportService<T> : BaseService<T>, IImportService where T : class, new()
{
    /// <summary>
    /// 获取导入类型
    /// </summary>
    public abstract string GetImportKey();

    /// <summary>
    /// 获取导入配置数据
    /// </summary>
    public abstract ImportConfig GetImportConfig(Dictionary<string, StringValues> paramMap, LoginUser user);

    /// <summary>
    /// 下载导入模板:前端接口直接调用该方法，如需完全自定义，重写该方法即可
    /// <para>注意：下载的文件名需要手动配置到PoiExcel.FileName</para>
    /// </summary>
    public virtual PoiExcel GetImportTemplate(Dictionary<string, StringValues> paramMap, LoginUser user)
    {
        var config = GetImportConfig(paramMap, user);
        var template = config.ImportTemplate ?? throw new Exception("没有配置导入模板，需要配置ImportConfig.ImportantTemplate或者重写IImportService.GetImportTemplate方法来获取导入模板");
        template.FileName = config.TemplateName;
        return PoiHelper.CreateExcel(template);
    }

    /// <summary>
    /// 下载基础数据:前端接口直接调用该方法，如需完全自定义，重写该方法即可
    /// <para>注意：下载的文件名需要手动配置到PoiExcel.FileName</para>
    /// </summary>
    public virtual Task<PoiExcel> GetImportBasicData(Dictionary<string, StringValues> paramMap, LoginUser user)
    {
        throw new Exception("未实现下载基础数据的方法，需要重写IImportService.GetImportBasicData方法");
    }

    /// <summary>
    /// 执行导入: 默认实现逻辑，保存文件，效验模板        
    /// </summary>
    public virtual async Task HandleImport(Stream stream, ImportConfig cfg, Dictionary<string, StringValues> paramMap, LoginUser user)
    {
        var workbook = PoiHelper.CreateWorkbook(stream) ?? throw new MessageException("无法识别导入的Excel，请检查文件是否标准Excel文件");

        // 检查模板，sheet数量，列标题
        if (cfg.ShowTemplate && cfg.ImportTemplate != null)
        {
            PoiHelper.ValidateExcelTemplate(workbook, cfg.ImportTemplate);
        }

        // 执行导入
        await HandleImport(workbook, cfg, paramMap, user);
    }

    /// <summary>
    /// 执行导入：默认根据第一个Sheet转成实体类 然后调用 HandleImportAsync(lstData, paramMap, user)
    /// 如果多个Sheet页的需要重写该方法：使用PoiHelper.GetDataFromExcel循环Sheet转成不同实体类即可
    /// </summary>
    public virtual async Task HandleImport(IWorkbook workbook, ImportConfig cfg, Dictionary<string, StringValues> paramMap, LoginUser user)
    {
        if (cfg.ShowTemplate && cfg.ImportTemplate != null)
        {
            var lstData = PoiHelper.GetDataFromExcel<T>(workbook, cfg.ImportTemplate.ExcelSheets[0]);
            await HandleImport(lstData, paramMap, user);
        }
    }

    /// <summary>
    /// 执行导入:最常用的方法，数据已经映射到实体类中
    /// </summary>
    public abstract Task HandleImport(List<T> lstData, Dictionary<string, StringValues> paramMap, LoginUser user);
}
