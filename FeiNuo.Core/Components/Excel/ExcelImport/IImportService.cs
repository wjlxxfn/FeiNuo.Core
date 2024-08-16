using Microsoft.Extensions.Primitives;
using NPOI.SS.UserModel;

namespace FeiNuo.Core
{
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

        #region 下载导入模板
        /// <summary>
        /// 下载导入模板:前端接口直接调用该方法，如需完全自定义，重写该方法即可
        /// </summary>
        public async Task<ExcelDownload> DownloadTemplateAsync(Dictionary<string, StringValues> paramMap, LoginUser user)
        {
            var excel = await GetExcelTemplateAsync(paramMap, user);
            var bytes = PoiHelper.GetExcelBytes(excel);
            return new ExcelDownload(excel.FileName, excel.ContentType, bytes);
        }

        /// <summary>
        /// 下载导入模板
        /// </summary>
        public Task<ExcelConfig> GetExcelTemplateAsync(Dictionary<string, StringValues> paramMap, LoginUser user);
        #endregion

        #region 下载基础数据
        /// <summary>
        /// 下载基础数据
        /// </summary>
        public async Task<ExcelDownload> DownloadBasicDataAsync(Dictionary<string, StringValues> paramMap, LoginUser user)
        {
            var excel = await GetExcelBasicDataAsync(paramMap, user);
            var bytes = PoiHelper.GetExcelBytes(excel);
            return new ExcelDownload(excel.FileName, excel.ContentType, bytes);
        }

        /// <summary>
        /// 下载基础数据
        /// </summary>
        public Task<ExcelConfig> GetExcelBasicDataAsync(Dictionary<string, StringValues> paramMap, LoginUser user);
        #endregion

        #region 处理数据导入
        /// <summary>
        /// 执行导入: 默认实现逻辑，保存文件，效验模板        
        /// </summary>
        public async Task HandleImportAsync(Stream stream, ImportConfig cfg, Dictionary<string, StringValues> paramMap, LoginUser user)
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
        /// 执行导入：默认根据模板配置，读取excel中的数据并存储到DatList中，
        /// </summary>
        public Task HandleImportAsync(IWorkbook workbook, ImportConfig cfg, Dictionary<string, StringValues> paramMap, LoginUser user);
        #endregion
    }
}
