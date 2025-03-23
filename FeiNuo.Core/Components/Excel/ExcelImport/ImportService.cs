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
    public abstract ImportConfig GetImportConfig(Dictionary<string, string> paramMap, LoginUser user);

    /// <summary>
    /// 下载导入模板
    /// </summary>
    public virtual Task<ExcelConfig> GetExcelTemplateAsync(Dictionary<string, string> paramMap, LoginUser user)
    {
        return Task.FromResult(new ExcelConfig("导入模板.xlsx"));
    }

    /// <summary>
    /// 下载基础数据
    /// </summary>
    public virtual Task<ExcelConfig> GetExcelBasicDataAsync(Dictionary<string, string> paramMap, LoginUser user)
    {
        return Task.FromResult(new ExcelConfig("基础数据.xlsx"));
    }

    /// <summary>
    /// 执行导入: 默认实现逻辑，保存文件，效验模板        
    /// </summary>
    public async Task HandleImportAsync(Stream stream, ImportConfig cfg, Dictionary<string, string> paramMap, LoginUser user)
    {
        var workbook = PoiHelper.CreateWorkbook(stream) ?? throw new MessageException("无法识别导入的Excel，请检查文件是否标准Excel文件");

        // 检查模板，sheet数量，列标题
        if (cfg.ShowTemplate)
        {
            var template = await GetExcelTemplateAsync(paramMap, user);
            cfg.ImportTemplate = template;
            PoiHelper.ValidateExcelTemplate(workbook, template);
        }

        // 执行导入
        await HandleImportAsync(workbook, cfg, paramMap, user);
    }

    /// <summary>
    /// 执行导入：默认根据第一个Sheet转成实体类 然后调用 HandleImportAsync(lstData, paramMap, user)
    /// 如果多个Sheet页的需要重写该方法：使用PoiHelper.GetDataFromExcel循环Sheet转成不同实体类即可
    /// </summary>
    public virtual async Task HandleImportAsync(IWorkbook workbook, ImportConfig cfg, Dictionary<string, string> paramMap, LoginUser user)
    {
        if (cfg.ShowTemplate && cfg.ImportTemplate != null)
        {
            var lstData = PoiHelper.GetDataFromExcel<T>(workbook, cfg.ImportTemplate.ExcelSheets[0]);
            await HandleImportAsync(lstData, paramMap, user);
        }
    }

    /// <summary>
    /// 执行导入:最常用的方法，数据已经映射到实体类中
    /// </summary>
    public abstract Task HandleImportAsync(IEnumerable<T> lstData, Dictionary<string, string> paramMap, LoginUser user);
}
