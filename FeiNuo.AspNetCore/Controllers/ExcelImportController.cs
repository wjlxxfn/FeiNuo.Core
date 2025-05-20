using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.DependencyInjection;

namespace FeiNuo.AspNetCore.Controllers;

[Route("/api/excel/import")]
public class ExcelImportController : BaseController
{
    /// <summary>
    /// 获取导入页面的配置信息，判断是否需要下载基础数据，显示说明文字等
    /// </summary>
    [HttpGet("config")]
    [EndpointSummary("获取Excel上传的配置信息")]
    [EndpointDescription("是否显示说明文字，是否显示下载基础数据，是否显示上传模板")]
    public ActionResult GetImportConfig([FromQuery] string importKey)
    {
        var service = GetExcelImportService(importKey);
        var cfg = service.GetImportConfig(GetRequestParam(), CurrentUser);
        return Ok(new
        {
            cfg.ShowRemark,
            cfg.ImportRemark,

            cfg.ShowBasicData,
            cfg.BasicDataName,

            cfg.ShowTemplate,
            cfg.TemplateName
        });
    }

    /// <summary>
    /// 下载导入模板
    /// </summary>
    [HttpGet("template")]
    [EndpointSummary("下载导入模板")]
    public IActionResult DownloadTemplate([FromQuery] string importKey)
    {
        var service = GetExcelImportService(importKey);
        var excel = service.GetImportTemplate(GetRequestParam(), CurrentUser, out var fileName);
        return File(excel.GetExcelBytes(), excel.ContentType, fileName);
    }

    /// <summary>
    /// 下载基础数据
    /// </summary>
    [HttpGet("basicdata")]
    [EndpointSummary("下载基础数据")]
    public IActionResult DownloadBasicData([FromQuery] string importKey)
    {
        var service = GetExcelImportService(importKey);
        var excel = service.GetImportBasicData(GetRequestParam(), CurrentUser, out var fileName);
        return File(excel.GetExcelBytes(), excel.ContentType, fileName);
    }

    /// <summary>
    /// 上传文件，执行导入
    /// </summary>
    [HttpPost]
    [EndpointSummary("上传文件，执行导入")]
    public async Task<IActionResult> UploadExcel(IFormFile file, [FromForm] string importKey)
    {
        if (file == null || file.Length == 0)
        {
            throw new MessageException("没有选择文件");
        }

        var permittedExtensions = new string[] { ".xls", ".xlsx" };
        var ext = Path.GetExtension(file.FileName).ToLowerInvariant();
        if (string.IsNullOrEmpty(ext) || !permittedExtensions.Contains(ext))
        {
            throw new Exception("只能上传excel文件");
        }

        var param = GetRequestParam();
        var user = CurrentUser;
        var service = GetExcelImportService(importKey);
        var cfg = service.GetImportConfig(param, user);

        // 验证权限
        if (cfg.AuthRoles.Length > 0)
        {
            if (!cfg.AuthRoles.Any(User.IsInRole))
            {
                return Forbid();
            }
        }

        // 保存文件
        var stream = file.OpenReadStream();
        if (cfg.SaveExcel)
        {
            //TODO 文件上传服务
            var dir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, cfg.SavePath, importKey, DateTime.Now.ToString("yyyy-MM"), Guid.NewGuid().ToString());
            Directory.CreateDirectory(dir);
            using var fileStream = System.IO.File.Create(Path.Combine(dir, file.FileName));
            await stream.CopyToAsync(fileStream);
            stream.Position = 0;
        }

        // 执行导入
        await service.HandleImport(stream, cfg, param, user);

        return NoContent();
    }

    #region 辅助方法
    private IImportService GetExcelImportService(string key)
    {
        if (string.IsNullOrWhiteSpace(key))
        {
            throw new MessageException("缺少必须的参数importKey");
        }

        var service = HttpContext.RequestServices.GetServices<IImportService>().SingleOrDefault(a => a.GetImportKey() == key);
        if (null == service)
        {
            throw new MessageException($"找不到【{key}】对应的Excel导入服务类");
        }

        return service;
    }
    #endregion
}
