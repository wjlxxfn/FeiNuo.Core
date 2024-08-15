using Microsoft.Extensions.Primitives;
using NPOI.SS.UserModel;

namespace FeiNuo.Core
{
    /// <summary>
    /// Excel导入实现接口
    /// </summary>
    public interface IImportService
    {
        /// <summary>
        /// 导入类型，系统根据类型匹配服务实现类
        /// </summary>
        public string GetImportKey();

        /// <summary>
        /// 导入所需的角色
        /// </summary>
        public string[] GetAuthRoles() { return []; }

        /// <summary>
        /// 获取导入配置
        /// </summary>
        public Task<ImportConfig> GetImportConfigAsync(Dictionary<string, StringValues> paramMap, LoginUser user);

        /// <summary>
        /// 下载导入模板
        /// 默认根据导入配置信息自动生成模板
        /// </summary>
        public async Task<ExcelDownload> GetTemplateAsync(Dictionary<string, StringValues> paramMap, LoginUser user)
        {
            var cfg = await GetImportConfigAsync(paramMap, user);
            var excel = (cfg.ShowTemplate && cfg.ImportTemplate != null) ? cfg.ImportTemplate : new ExcelConfig("导入模板.xlsx");
            var bytes = PoiHelper.GetExcelBytes(excel);
            return new ExcelDownload(excel.FileName, excel.ContentType, bytes);
        }

        /// <summary>
        /// 下载基础数据
        /// </summary>
        public async Task<ExcelDownload> GetBasicDataAsync(Dictionary<string, StringValues> paramMap, LoginUser user)
        {
            var excel = await GetBasicDataExcelAsync(paramMap, user);
            var bytes = PoiHelper.GetExcelBytes(excel);
            return new ExcelDownload(excel.FileName, excel.ContentType, bytes);
        }

        /// <summary>
        /// 下载基础数据
        /// </summary>
        public Task<ExcelConfig> GetBasicDataExcelAsync(Dictionary<string, StringValues> paramMap, LoginUser user);

        /// <summary>
        /// 执行导入: 默认实现逻辑，保存文件，效验模板        
        /// </summary>
        public async Task HandleImportAsync(Stream stream, string fileName, Dictionary<string, StringValues> paramMap, LoginUser user)
        {
            var cfg = await GetImportConfigAsync(paramMap, user);

            // 保存文件
            if (cfg.SaveExcel)
            {
                //TODO 文件上传服务
                var dir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, cfg.SavePath, GetImportKey(), DateTime.Now.ToString("yyyy-MM"), Guid.NewGuid().ToString());
                Directory.CreateDirectory(dir);
                using var fileStream = File.Create(Path.Combine(dir, fileName));
                await stream.CopyToAsync(fileStream);
                stream.Position = 0;
            }

            var workbook = PoiHelper.CreateWorkbook(stream) ?? throw new MessageException("无法识别导入的Excel，请检查文件是否标准Excel文件");

            // 检查模板，sheet数量，列标题
            if (cfg.ShowTemplate && cfg.ImportTemplate != null)
            {
                PoiHelper.ValidateExcelTemplate(workbook, cfg.ImportTemplate);
            }

            // 执行导入
            await HandleImportAsync(workbook, cfg, paramMap, user);
        }

        /// <summary>
        /// 执行导入：默认根据模板配置，读取excel中的数据并存储到DatList中，
        /// </summary>
        public Task HandleImportAsync(IWorkbook workbook, ImportConfig cfg, Dictionary<string, StringValues> paramMap, LoginUser user);

    }
}
