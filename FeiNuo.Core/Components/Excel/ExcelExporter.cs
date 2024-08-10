using NPOI.SS.UserModel;

namespace FeiNuo.Core
{
    public class ExcelExporter
    {
        #region 对外开放属性
        public IWorkbook Workbook { get; private set; }
        public StyleFactory Styles { get; private set; }
        public ISheet? Sheet { get; private set; }
        public string FileName { get; private set; }
        public string ContentType { get; private set; }
        #endregion

        #region 构造方法
        public ExcelExporter(string fileName, ExcelType? excelType = null, ExcelStyle? defaultStyle = null)
        {
            var cfg = new ExcelConfig(fileName);
            if (excelType != null) cfg.ExcelType = excelType.Value;
            if (defaultStyle != null) cfg.DefaultStyle = defaultStyle;

            Workbook = ExcelHelper.CreateWorkbook(cfg, out var styles);
            Styles = styles;

            FileName = cfg.FileName;
            ContentType = cfg.ContentType;
        }
        #endregion

        #region 添加数据

        /// <summary>
        /// 提供数据和列配置生成Excel Sheet，SheetName默认为Sheet1
        /// </summary>
        /// <typeparam name="T">数据类型</typeparam>
        /// <param name="lstData">数据集合</param>
        /// <param name="columns">列配置</param>
        /// <param name="sheetConfig">工作表整体配置</param>
        public ExcelExporter CreateSheet<T>(IEnumerable<T> lstData, IEnumerable<ExcelColumn<T>> columns, Action<ExcelSheet<T>>? sheetConfig = null) where T : class
        {
            return CreateSheet("Sheet1", lstData, columns, sheetConfig);
        }

        /// <summary>
        /// 提供数据和列配置生成Excel Sheet，SheetName默认为Sheet1
        /// </summary>
        /// <typeparam name="T">数据类型</typeparam>
        /// <param name="sheetName">工作表名</param>
        /// <param name="lstData">数据集合</param>
        /// <param name="columns">列配置</param>
        /// <param name="sheetConfig">工作表整体配置</param>
        public ExcelExporter CreateSheet<T>(string sheetName, IEnumerable<T> lstData, IEnumerable<ExcelColumn<T>> columns, Action<ExcelSheet<T>>? sheetConfig = null) where T : class
        {
            var excelSheet = new ExcelSheet<T>(sheetName, lstData, columns);
            sheetConfig?.Invoke(excelSheet);
            return CreateSheet(excelSheet);
        }

        /// <summary>
        /// 根据配置对象生成Excel
        /// </summary>
        internal ExcelExporter CreateSheet<T>(ExcelSheet<T> excelSheet) where T : class
        {
            if (string.IsNullOrWhiteSpace(excelSheet.SheetName))
            {
                throw new MessageException("SheetName不能为空");
            }
            if (Workbook.GetSheetIndex(excelSheet.SheetName) >= 0)
            {
                throw new MessageException($"【{excelSheet.SheetName}】已存在");
            }

            Sheet = ExcelHelper.CreateDataSheet(Workbook, excelSheet, Styles);
            ExcelHelper.AddConditionalBorderStyle(Sheet);
            return this;
        }
        #endregion

        #region 辅助方法
        /// <summary>
        /// POI Workbook 转成字节数组
        /// </summary>
        /// <returns></returns>
        public byte[] GetBytes()
        {
            return ExcelHelper.GetExcelBytes(Workbook);
        }
        #endregion
    }
}
