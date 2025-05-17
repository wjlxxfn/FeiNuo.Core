using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System.Data;
using System.Text;

namespace FeiNuo.Core
{
    public partial class PoiHelper
    {
        #region 转换ExcelConfig
        /// <summary>
        /// 创建工作簿
        /// </summary>
        public static IWorkbook CreateWorkbook(ExcelConfig config, out StyleFactory styles)
        {
            // 检查配置内容
            config.ValidateConfigData();

            // 创建工作簿，生成样式工厂
            var wb = CreateWorkbook(config.IsExcel2007);
            styles = new StyleFactory(wb, config.DefaultStyle);

            // 创建工作表
            foreach (var sheet in config.ExcelSheets)
            {
                CreateWorkSheet(wb, sheet, styles);
            }

            return wb;
        }

        public static IWorkbook FillWorkbook(IWorkbook wb, ExcelConfig config, out StyleFactory styles)
        {
            // 检查配置内容
            config.ValidateConfigData();

            styles = new StyleFactory(wb, config.DefaultStyle);

            // 创建工作表
            foreach (var sheet in config.ExcelSheets)
            {
                CreateWorkSheet(wb, sheet, styles);
            }

            return wb;
        }

        /// <summary>
        /// 创建工作表:添加备注描述，主标题，列标题
        /// </summary>
        public static ISheet CreateWorkSheet(IWorkbook wb, ExcelSheet config, StyleFactory styles)
        {
            var sheet = wb.GetSheet(config.SheetName);
            sheet ??= wb.CreateSheet(config.SheetName);

            var columnCount = config.ColumnCount;
            int rowIndex = config.StartRowIndex;
            IRow row; ICell cell;

            #region 生成描述行
            if (!string.IsNullOrWhiteSpace(config.Description))
            {
                row = GetRow(sheet, rowIndex);
                row.HeightInPoints = (short)(config.DescriptionRowHeight < 0 ? 20 : config.DescriptionRowHeight);
                cell = GetCell(row, config.StartColumnIndex);
                cell.CellStyle = styles.CreateStyle(config.DescriptionStyle);
                cell.SetCellValue(config.Description);

                var mergeCount = config.DescriptionColSpan ?? columnCount;
                if (mergeCount > 0)
                {
                    sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, config.StartColumnIndex, config.StartColumnIndex + mergeCount - 1));
                }
                rowIndex++;
            }
            #endregion

            #region 生成主标题行
            if (!string.IsNullOrWhiteSpace(config.MainTitle))
            {
                row = GetRow(sheet, rowIndex);
                cell = GetCell(row, config.StartColumnIndex);
                cell.CellStyle = styles.CreateStyle(config.MainTitleStyle);
                cell.SetCellValue(config.MainTitle);
                var mergeCount = config.MainTitleColSpan ?? columnCount;
                if (mergeCount > 0)
                {
                    sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, config.StartColumnIndex, config.StartColumnIndex + mergeCount - 1));
                }
                rowIndex++;
            }
            #endregion

            #region 生成列标题行
            if (columnCount > 0)
            {
                var titleStyle = styles.CreateStyle(config.ColumnTitleStyle);
                var titleRowCount = config.ExcelColumns.Max(a => a.RowTitles.Length);
                int titleRowStartIndex = rowIndex, titleRowEndIndex = rowIndex + titleRowCount - 1;
                foreach (var col in config.ExcelColumns)
                {
                    rowIndex = titleRowStartIndex;

                    // 设置默认格式
                    if (col.ColumnStyle.IsNotEmptyStyle)
                    {
                        sheet.SetDefaultColumnStyle(col.ColumnIndex, styles.CreateStyle(col.ColumnStyle));
                    }

                    foreach (var rt in col.RowTitles)
                    {
                        cell = GetCell(sheet, rowIndex++, col.ColumnIndex);
                        cell.CellStyle = titleStyle;
                        cell.SetCellValue(rt);
                    }

                    // 配置列宽，隐藏
                    if (col.Width.HasValue) sheet.SetColumnWidth(col.ColumnIndex, col.Width.Value * 256);
                    if (col.Hidden) sheet.SetColumnHidden(col.ColumnIndex, true);
                }
                //合并标题行
                for (var i = config.StartColumnIndex; i < config.StartColumnIndex + columnCount; i++)
                {
                    AutoMergeRows(sheet, i, titleRowStartIndex, titleRowEndIndex);
                }
                // 合并标题列
                if (titleRowCount > 1)
                {
                    for (var i = 0; i < titleRowCount; i++)
                    {
                        AutoMergeColumns(sheet, i + config.TitleRowIndex, config.StartColumnIndex, config.StartColumnIndex + columnCount - 1);
                    }
                }
                // 重置行索引
                rowIndex = titleRowEndIndex + 1;
            }
            #endregion

            #region 生成数据行
            if (config.DataList.Any())
            {
                foreach (var data in config.DataList)
                {
                    row = GetRow(sheet, rowIndex++);
                    foreach (var col in config.ExcelColumns)
                    {
                        var val = col.ValueGetter?.Invoke(data);
                        var style = col.ColumnStyle.IsNotEmptyStyle
                            ? styles.CreateStyle(col.ColumnStyle)
                            : ((val != null && (val is DateOnly || val is DateTime)) ? styles.DateStyle : null);
                        SetCellValue(row, col.ColumnIndex, val, false, style);
                    }
                }
            }
            #endregion

            #region 工作表整体配置
            // 设置默认列宽
            if (config.DefaultColumnWidth.HasValue)
            {
                sheet.DefaultColumnWidth = config.DefaultColumnWidth.Value;
            }

            // 自动计算公式
            if (config.ForceFormulaRecalculation) sheet.ForceFormulaRecalculation = true;

            // 自动设置边框

            if (config.AddConditionalBorderStyle)
            {
                var region = new CellRangeAddress(config.StartRowIndex, sheet.LastRowNum, config.StartColumnIndex, config.EndColumnIndex);
                AddConditionalBorderStyle(sheet, range: region);
            }
            #endregion

            return sheet;
        }

        /// <summary>
        /// 根据模板配置，检查excel是否指定模板
        /// </summary>
        public static void ValidateExcelTemplate(IWorkbook wb, ExcelConfig config)
        {
            foreach (var sheet in config.ExcelSheets)
            {
                ValidateExcelTemplate(wb, sheet);
            }
        }

        /// <summary>
        /// 根据模板配置，检查excel是否指定模板
        /// </summary>
        public static void ValidateExcelTemplate(IWorkbook wb, ExcelSheet config)
        {
            if (!config.ValidateImportTemplate) return;
            var sheet = wb.GetSheet(config.SheetName) ?? throw new MessageException($"没有找到名为【{config.SheetName}】的Sheet页");
            if (!config.ExcelColumns.Any()) return;
            var row = GetRow(sheet, config.TitleRowIndex + config.ColumnRowCount - 1);
            foreach (var col in config.ExcelColumns)
            {
                if (GetStringValue(GetCell(row, col.ColumnIndex)) != col.RowTitles.Last())
                {
                    throw new Exception($"【{config.SheetName}】模板错误，【{col.ColumnIndex}】列标题应为【{col.Title}】，请重新下载模板。");
                }
            }
        }

        /// <summary>
        /// 从Excel中读取数据，并赋值到ExcelSheet的DataList
        /// </summary>
        public static List<T> GetDataFromExcel<T>(IWorkbook wb, ExcelSheet config) where T : class, new()
        {
            var sheet = wb.GetSheet(config.SheetName) ?? throw new Exception($"找不到Sheet【{config.SheetName}】");
            // 计算公式
            sheet.ForceFormulaRecalculation = true;

            var keyMap = new Dictionary<string, int>();
            var lstData = new List<T>();
            var errMsg = new StringBuilder();

            for (var rowIndex = config.DataRowIndex; rowIndex <= sheet.LastRowNum; rowIndex++)
            {
                var row = GetRow(sheet, rowIndex);
                if (IsBlankRow(row)) continue;

                var data = Activator.CreateInstance<T>() ?? throw new Exception($"无法实例化对象{typeof(T)}");
                var rowMsg = new StringBuilder();
                var keyValue = new StringBuilder();
                foreach (var col in config.ExcelColumns)
                {
                    var cell = GetCell(row, col.ColumnIndex);
                    var val = GetCellValue(cell);

                    try
                    {
                        var msg = col.ValueSetter(data, val);
                        if (msg == "_UniqueKey_")
                        {
                            keyValue.Append((val ?? "null") + "|");
                        }
                        else if (!string.IsNullOrWhiteSpace(msg))
                        {
                            rowMsg.Append($"列【{col.RowTitles.Last()}】{msg}；");
                        }
                    }
                    catch (Exception ex)
                    {
                        rowMsg.Append($"列【{col.RowTitles.Last()}】{ex.Message}；");
                    }
                }

                if (rowMsg.Length > 0)
                {
                    errMsg.AppendLine($"第【{rowIndex}】行：{rowMsg}");
                }
                else
                {
                    var key = keyValue.ToString();
                    if (!string.IsNullOrEmpty(key))
                    {
                        if (keyMap.TryGetValue(key, out int existingRow))
                        {
                            errMsg.AppendLine($"第【{rowIndex}】行：与{existingRow}行重复,重复键值：{key}；");
                        }
                        else
                        {
                            keyMap[key] = rowIndex;
                        }
                    }
                    lstData.Add(data);
                }
            }

            if (errMsg.Length > 0)
            {
                throw new MessageException($"获取Excel数据出错:<br/>{errMsg}");
            }

            return lstData;
        }

        public static DataSet GetDataSetFromExcel(IWorkbook wb)
        {
            var ds = new DataSet();
            for (int i = 0; i < wb.NumberOfSheets; i++)
            {
                var sheet = wb.GetSheetAt(i);
                var dt = GetDataTableFromSheet(sheet);
                ds.Tables.Add(dt);
            }
            return ds;
        }

        public static DataTable GetDataTableFromSheet(ISheet sheet, int headerRowIndex = 0)
        {
            var dt = new DataTable(sheet.SheetName);
            var rowCount = sheet.LastRowNum + 1;
            for (int j = headerRowIndex; j < rowCount; j++)
            {
                var row = sheet.GetRow(j);
                if (row == null) continue;
                if (j == 0)
                {
                    foreach (var cell in row.Cells)
                    {
                        dt.Columns.Add(GetStringValue(cell));
                    }
                }
                else
                {
                    var dr = dt.NewRow();
                    foreach (var cell in row.Cells)
                    {
                        dr[cell.ColumnIndex] = GetCellValue(cell) ?? DBNull.Value;
                    }
                    dt.Rows.Add(dr);
                }
            }
            return dt;
        }

        /// <summary>
        /// Poi.SetDefaultColumnStyle后，如果使用POI赋值会把默认样式覆盖，调用该方法重新设置样式
        /// </summary>
        /// <param name="wb"></param>
        /// <param name="config"></param>
        /// <param name="styles"></param>
        public static void ResetColumnStyle(IWorkbook wb, ExcelConfig config, StyleFactory styles)
        {
            foreach (var sheetConfig in config.ExcelSheets)
            {
                var sheet = wb.GetSheet(sheetConfig.SheetName);
                if (sheet == null) continue;

                foreach (var col in sheetConfig.ExcelColumns)
                {
                    if (col.ColumnStyle.IsEmptyStyle) continue;

                    var cellStyle = styles.CreateStyle(col.ColumnStyle);
                    for (var r = sheetConfig.DataRowIndex; r <= sheet.LastRowNum; r++)
                    {
                        var cell = sheet.GetRow(r)?.GetCell(col.ColumnIndex);
                        if (cell != null) cell.CellStyle = cellStyle;
                    }
                }
            }
        }
        #endregion

        /// <summary>
        /// 生成IWorkbook,并转成字节数组
        /// </summary>
        public static byte[] GetExcelBytes(ExcelConfig config)
        {
            var wb = CreateWorkbook(config, out _);
            return GetExcelBytes(wb);
        }
    }
}
