﻿namespace FeiNuo.Core
{
    /// <summary>
    /// Excel配置类。配置后使用ExcelHelper.CreateWorkBook方法可生成IWorkbook对象
    /// </summary>
    public class ExcelConfig
    {
        #region 属性定义
        /// <summary>
        /// 文件名称：导出文件名，导入的模板下载名
        /// </summary>
        public string FileName { get; set; } = null!;

        /// <summary>
        /// Excel类型,2007/2003
        /// </summary>
        public ExcelType ExcelType { get; set; } = ExcelType.Excel2007;

        /// <summary>
        /// 工作表配置
        /// </summary>
        public List<ExcelSheet> ExcelSheets { get; set; } = [];

        /// <summary>
        /// 是否2007格式
        /// </summary>
        public bool IsExcel2007 { get { return ExcelType == ExcelType.Excel2007; } }

        /// <summary>
        /// contentType
        /// </summary>
        public string ContentType { get { return IsExcel2007 ? "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" : "application/vnd.ms-excel"; } }

        #endregion

        #region 公共方法
        public ExcelConfig AddExcelSheet(string sheetName, Action<ExcelSheet>? configExcelSheet = null)
        {
            if (string.IsNullOrWhiteSpace(sheetName)) throw new ArgumentNullException(nameof(sheetName));
            if (ExcelSheets.Any(t => t.SheetName == sheetName))
            {
                throw new MessageException($"已存在名为【{sheetName}】的工作表。");
            }
            var sheet = new ExcelSheet(sheetName);
            configExcelSheet?.Invoke(sheet);
            ExcelSheets.Add(sheet);
            return this;
        }

        /// <summary>
        /// 验证配置数据是否有不合适的
        /// </summary>
        public void ValidateConfigData()
        {
            var extension = Path.GetExtension(FileName).ToLower();
            if ((IsExcel2007 && extension != ".xlsx") || (!IsExcel2007 && extension != ".xls"))
            {
                throw new Exception("Excel版本和后缀不匹配");
            }
            foreach (var sheet in ExcelSheets)
            {
                var columns = sheet.ExcelColumns;
                // 有相同的title报错
                var chk = columns.GroupBy(k => k.Title).Select(a => new { Title = a.Key, Count = a.Count() }).Where(a => a.Count > 1)
                    .Select(a => a.Title).Distinct().ToArray();
                if (chk.Length != 0) throw new MessageException($"【{sheet.SheetName}】以下列字段标题重复:" + string.Join(".", chk));
                var titleRows = columns.Select(a => a.RowTitles.Length).Distinct().Where(a => a > 1);
                if (titleRows.Count() > 1)
                {
                    throw new MessageException($"【{sheet.SheetName}】标题行有多行时必须保证每列都有相同的行数");
                }
            }
        }
        #endregion

        #region 构造函数
        public ExcelConfig(string fileName)
        {
            FileName = fileName + (string.IsNullOrWhiteSpace(Path.GetExtension(fileName)) ? (IsExcel2007 ? ".xlsx" : ".xls") : "");
        }
        public ExcelConfig(string fileName, Action<ExcelSheet> configExcelSheet) : this(fileName)
        {
            AddExcelSheet("Sheet1", configExcelSheet);
        }
        #endregion
    }

    /// <summary>
    /// Excel版本
    /// </summary>
    public enum ExcelType
    {
        Excel2007,
        Excel2003
    }
}