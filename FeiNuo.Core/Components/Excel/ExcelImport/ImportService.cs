using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Primitives;
using NPOI.SS.UserModel;

namespace FeiNuo.Core
{
    public abstract class BaseImportService<T> : BaseService<T>, IImportService where T : BaseEntity, new()
    {
        public BaseImportService(DbContext ctx) : base(ctx) { }

        /// <summary>
        /// 获取导入类型
        /// </summary>
        public abstract string GetImportKey();

        /// <summary>
        /// 获取导入配置数据
        /// </summary>
        public abstract Task<ImportConfig> GetImportConfigAsync(Dictionary<string, StringValues> paramMap, LoginUser user);

        /// <summary>
        /// 下载基础数据
        /// </summary>
        public virtual Task<ExcelConfig> GetBasicDataExcelAsync(Dictionary<string, StringValues> paramMap, LoginUser user)
        {
            return Task.FromResult(new ExcelConfig("基础数据.xlsx"));
        }

        /// <summary>
        /// 执行导入：默认根据第一个excel转成实体类 然后调用 HandleImportAsync(lstData, paramMap, user)
        /// 如果多个Sheet页的需要重新该方法，参数实现逻辑，循环转成不同实体类即可
        /// </summary>
        public virtual async Task HandleImportAsync(IWorkbook workbook, ImportConfig cfg, Dictionary<string, StringValues> paramMap, LoginUser user)
        {
            if (cfg.ShowTemplate && cfg.ImportTemplate != null)
            {
                var lstData = PoiHelper.GetDataFromExcel(workbook, (ExcelSheet<T>)cfg.ImportTemplate.ExcelSheets[0]);
                await HandleImportAsync(lstData, paramMap, user);
            }
        }

        /// <summary>
        /// 执行导入
        /// </summary>
        public abstract Task HandleImportAsync(IEnumerable<T> lstData, Dictionary<string, StringValues> paramMap, LoginUser user);
    }
}
