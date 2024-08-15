namespace FeiNuo.Core
{
    /// <summary>
    /// Excel导入配置类
    /// </summary>
    public class ImportConfig
    {
        #region 构造方法
        /// <summary>
        /// 导入配置构造函数
        /// </summary>
        /// <param name="templateName">如果需要下载模板，则必须指定模板名称</param>
        public ImportConfig(string templateName = "导入模板.xlsx")
        {
            if (!string.IsNullOrWhiteSpace(templateName))
            {
                ImportTemplate = new ExcelConfig(templateName);
            }
            else
            {
                ShowTemplate = false;
            }
        }

        /// <summary>
        /// 方便链式调用
        /// </summary>
        public ImportConfig ConfigImport(Action<ImportConfig> configAction)
        {
            configAction(this);
            return this;
        }
        #endregion

        #region 导入模板
        public bool ShowTemplate { get; private set; } = true;
        public ExcelConfig? ImportTemplate { get; set; }
        #endregion

        #region 基础数据
        public bool ShowBasicData { get; set; } = true;
        public string BasicDataName { get; set; } = "基础数据.xlsx";
        #endregion

        #region 导入说明
        public string ImportRemark { get; set; } = string.Empty;
        public bool ShowRemark { get { return !string.IsNullOrWhiteSpace(ImportRemark); } }
        #endregion

        #region 保存文件
        /// <summary>
        /// 导入前是否保存excel文件
        /// </summary>
        public bool SaveExcel { get; set; } = false;

        /// <summary>
        /// excel保存路径,默认根路径下Attachment/ExcelImport
        /// 配置路径后，文件会在该路径下创建子路径保存文件 /ImportType/yyyy-mm/guid/filename.
        /// </summary>
        public string SavePath { get; set; } = "Attachment/ExcelImport";
        #endregion

        #region 配置模板
        internal ImportConfig AddTemplate<T>(ExcelSheet<T> sheetConfig) where T : class, new()
        {
            if (ImportTemplate == null)
            {
                throw new MessageException("请指定导入模板名称");
            }

            // 必须使用ExcelColumnImp类配置列
            var col = sheetConfig.ExcelColumns.FirstOrDefault();
            if (col == null)
            {
                throw new MessageException($"【{sheetConfig.SheetName}】未配置数据列");
            }
            ImportTemplate.AddExcelSheet(sheetConfig);

            return this;
        }

        /// <summary>
        /// 配置导入模板的Sheet,默认SheetName = Sheet1
        /// </summary>
        /// <typeparam name="T">导入数据对应的实体类</typeparam>
        /// <param name="columns">列配置，必须是ExcelColumnImp类或其子类</param>
        /// <param name="sheetConfig">工作表配置</param>
        public ImportConfig AddTemplate<T>(IEnumerable<ExcelColumn<T>> columns, Action<ExcelSheet<T>>? sheetConfig = null) where T : class, new()
        {
            return AddTemplate("Sheet1", columns, sheetConfig);
        }
        /// <summary>
        /// 配置导入模板的Sheet
        /// </summary>
        /// <typeparam name="T">导入数据对应的实体类</typeparam>
        /// <param name="sheetName">工作表名</param>
        /// <param name="columns">列配置，必须是ExcelColumnImp类或其子类</param>
        /// <param name="sheetConfig">工作表配置</param>
        public ImportConfig AddTemplate<T>(string sheetName, IEnumerable<ExcelColumn<T>> columns, Action<ExcelSheet<T>>? sheetConfig = null) where T : class, new()
        {
            var sheet = new ExcelSheet<T>(sheetName, columns);
            sheetConfig?.Invoke(sheet);
            AddTemplate(sheet);
            return this;
        }
        #endregion
    }
}
