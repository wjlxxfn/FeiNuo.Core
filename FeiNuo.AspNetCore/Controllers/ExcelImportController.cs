using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.WebUtilities;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Primitives;

namespace FeiNuo.Core
{
    [AllowAnonymous]
    [Route("/api/excel/upload")]
    public class ExcelImportController : BaseController
    {
        /// <summary>
        /// 获取导入页面的配置信息，判断是否需要下载基础数据，显示说明文字等
        /// </summary>
        [HttpGet("config")]
        public async Task<IActionResult> GetImportConfig([FromQuery] string importKey)
        {
            var service = GetExcelImportService(importKey);
            var cfg = await service.GetImportConfigAsync(ParamMap, CurrentUser);
            return Ok(new
            {
                cfg.ShowRemark,
                cfg.ImportRemark,

                cfg.ShowBasicData,
                cfg.BasicDataName,

                cfg.ShowTemplate,
                cfg.ImportTemplate!.FileName
            });
        }

        /// <summary>
        /// 下载导入模板
        /// </summary>
        [HttpGet("template")]
        public async Task<IActionResult> DownloadTemplate([FromQuery] string importKey)
        {
            var service = GetExcelImportService(importKey);
            var file = await service.GetTemplateAsync(ParamMap, CurrentUser);
            return File(file.Bytes, file.ContentType, file.FileName);
        }

        /// <summary>
        /// 下载基础数据
        /// </summary>
        [HttpGet("basicdata")]
        public async Task<IActionResult> DownloadBasicData([FromQuery] string importKey)
        {
            var service = GetExcelImportService(importKey);
            var file = await service.GetBasicDataAsync(ParamMap, CurrentUser);
            return File(file.Bytes, file.ContentType, file.FileName);
        }

        /// <summary>
        /// 上传文件，执行导入
        /// </summary>
        [HttpPost]
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

            var service = GetExcelImportService(importKey);
            // 验证权限
            var authRoles = service.GetAuthRoles();
            if (authRoles.Length > 0)
            {
                if (!authRoles.Any(User.IsInRole))
                {
                    return Unauthorized();
                }
            }
            await service.HandleImportAsync(file.OpenReadStream(), file.FileName, ParamMap, CurrentUser);

            return NoContent();
        }

        #region 辅助方法
        private Dictionary<string, StringValues> ParamMap
        {
            get
            {
                return QueryHelpers.ParseQuery(Request.QueryString.Value);
            }
        }

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
}
