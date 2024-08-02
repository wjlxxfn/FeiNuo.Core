using FeiNuo.Core.Utilities;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace FeiNuo.Core
{
    public class ExcelHelper
    {
        /// <summary>
        /// 创建工作表
        /// </summary>
        public static IWorkbook CreateWorkBook(ExcelConfig config, out StyleFactory styles)
        {
            config.ValidateConfigData();

            var wb = PoiUtils.CreateWorkbook(config.IsExcel2007);
            styles = new StyleFactory(wb);

            foreach (var sheet in config.ExcelSheets)
            {
                CreateWorkSheet(wb, sheet, styles);
            }
            return wb;
        }

        /// <summary>
        /// 创建工作表:添加标题，内容
        /// </summary>
        private static ISheet CreateWorkSheet(IWorkbook wb, ExcelSheet config, StyleFactory styles)
        {
            var sheet = wb.CreateSheet(config.SheetName);

            int rowIndex = 0, colIndex = 0;
            IRow row; ICell cell; ICellStyle style;

            #region 生成描述行
            if (!string.IsNullOrWhiteSpace(config.Description))
            {
                row = PoiUtils.GetRow(sheet, rowIndex);
                row.HeightInPoints = (short)(config.DescriptionRowHeight < 0 ? 20 : config.DescriptionRowHeight);
                cell = PoiUtils.GetCell(row, 0);
                cell.CellStyle = styles.DescriptionStyle;
                cell.SetCellValue(config.Description);

                var mergeCount = config.DescriptionMergedRegion ?? config.ExcelColumns.Count;
                if (mergeCount > 0)
                {
                    sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, config.StartColIndex, config.StartColIndex + mergeCount - 1));
                }
                rowIndex++;
            }
            #endregion

            #region 生成主标题行
            var titleStyle = styles.TitleStyle;
            if (!string.IsNullOrWhiteSpace(config.MainTitle))
            {
                cell = PoiUtils.GetCell(sheet, rowIndex, 0);
                cell.CellStyle = titleStyle;
                cell.SetCellValue(config.MainTitle);
                var mergeCount = config.MainTitleMergedRegion ?? config.ExcelColumns.Count;
                if (mergeCount > 0)
                {
                    sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, config.StartColIndex, config.StartColIndex + mergeCount - 1));
                }
                rowIndex++;
            }
            #endregion

            #region 生成列标题行
            var autoSizeCols = new List<int>(); // 自动列宽
            var hiddenCols = new List<int>();// 需要隐藏
            if (config.ExcelColumns.Count > 0)
            {
                style = styles.TitleStyle;

                var titleRowCount = config.ExcelColumns.Select(a => a.RowTitles.Length).Max();
                int titleRowStartIndex = rowIndex, titleRowEndIndex = rowIndex + titleRowCount - 1;
                colIndex = config.StartColIndex;
                foreach (var col in config.ExcelColumns)
                {
                    var rowTitles = col.RowTitles.ToList();
                    #region 保证每列都有相同的行数，不足的用最后的行标题填充
                    if (rowTitles.Count < titleRowCount)
                    {
                        var last = rowTitles.Last();
                        var addCount = titleRowCount - rowTitles.Count;
                        for (var i = 0; i < addCount; i++)
                        {
                            rowTitles.Add(last);
                        }
                    }
                    #endregion

                    string? lastTitle = null;
                    rowIndex = titleRowStartIndex;
                    var mergeStartRow = titleRowStartIndex;
                    // 写入标题
                    foreach (var curTitle in rowTitles)
                    {
                        cell = PoiUtils.GetCell(sheet, rowIndex, colIndex);
                        cell.CellStyle = style;
                        cell.SetCellValue(curTitle);

                        if (titleRowCount > 1)
                        {
                            // 多行合并,连续相同的标题合并
                            lastTitle ??= curTitle;
                            // 标题不同或者到了最后一行都可能需要合并  性别#性别#别  性别#别#别
                            if (curTitle != lastTitle || titleRowEndIndex == rowIndex)
                            {
                                // 合并结束位置:如果当前行标题和前一行不一致，则取前一行的行号。
                                // 如果已经到了最后一行且还和前一行的标题一致，则取最后一行的行号
                                var mergeEnd = (titleRowEndIndex == rowIndex && curTitle == lastTitle) ? titleRowEndIndex : (rowIndex - 1);
                                if (mergeEnd > mergeStartRow)
                                {
                                    sheet.AddMergedRegion(new CellRangeAddress(mergeStartRow, mergeEnd, colIndex, colIndex));
                                }
                                lastTitle = curTitle;
                                mergeStartRow = rowIndex;
                            }
                        }
                        rowIndex++;
                    }

                    // 配置列宽，隐藏
                    if (col.Width.HasValue)
                    {
                        sheet.SetColumnWidth(colIndex, col.Width.Value * 256);
                    }
                    if (col.AutoSizeColumn) autoSizeCols.Add(colIndex);
                    if (col.Hidden) hiddenCols.Add(colIndex);
                    // 设置默认格式
                    sheet.SetDefaultColumnStyle(colIndex, styles.NumbericStyle);
                    colIndex++;
                }
                // 相邻列相同的合并列
                if (titleRowCount > 1)
                {
                    int titleColEndIndex = config.StartColIndex + config.ExcelColumns.Count - 1;
                    for (int i = 0; i < titleRowCount; i++) // 循环每一行，合并行中的列
                    {
                        rowIndex = titleRowStartIndex + i;
                        var mergeStartCol = config.StartColIndex;
                        string? lastTitle = null;
                        for (var j = 0; j < config.ExcelColumns.Count; j++)
                        {
                            colIndex = config.StartColIndex + j;
                            cell = PoiUtils.GetCell(sheet, rowIndex, colIndex);
                            var curTitle = cell.StringCellValue;
                            lastTitle ??= curTitle;

                            if (curTitle != lastTitle || cell.IsMergedCell || colIndex == titleColEndIndex)
                            {
                                // 合并结束位置:和行合并类似，如果到了最后一行且是需要合并的，索引就在最后一行，否则取当前行的前一行
                                var mergeEnd = (colIndex == titleColEndIndex && curTitle == lastTitle && !cell.IsMergedCell) ? titleColEndIndex : (colIndex - 1);
                                if (mergeEnd > mergeStartCol)
                                {
                                    sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, mergeStartCol, mergeEnd));
                                }
                                lastTitle = curTitle;
                                // 当前单元格已经合并的时候不能做起始合并，需要从下一列开始
                                mergeStartCol = cell.IsMergedCell ? (colIndex + 1) : colIndex;
                            }
                        }
                    }
                }
                // 重置行索引
                rowIndex = titleRowEndIndex + 1;
            }
            #endregion

            #region 添加数据行
            #endregion

            #region 工作表整体配置
            // 设置默认列宽
            if (config.DefaultColumnWidth.HasValue)
            {
                sheet.DefaultColumnWidth = config.DefaultColumnWidth.Value;
            }

            // 设置自动列宽
            foreach (var item in autoSizeCols)
            {
                sheet.AutoSizeColumn(item);
            }
            // 设置隐藏
            foreach (var item in hiddenCols)
            {
                sheet.SetColumnHidden(item, true);
            }

            // 自动计算公式
            if (config.ForceFormulaRecalculation) sheet.ForceFormulaRecalculation = true;

            // 自动设置边框
            if (config.AddConditionalBorderStyle) AddConditionalBorderStyle(sheet);
            #endregion

            return sheet;
        }

        /// <summary>
        /// 通过条件格式给所有单元格添加边框
        /// </summary>
        /// <param name="sheet">工作表</param>
        /// <param name="borderStyle">边框样式，默认Thin</param>
        /// <param name="range">范围，默认从A1开始，最大列取前三行有内容的最后一列的列索引，最大行取当前sheet的最后一行</param>
        public static void AddConditionalBorderStyle(ISheet sheet, BorderStyle borderStyle = BorderStyle.Thin, CellRangeAddress? range = null)
        {
            var scf = sheet.SheetConditionalFormatting;
            var rule = scf.CreateConditionalFormattingRule("TRUE()");
            var fmt = rule.CreateBorderFormatting();
            fmt.BorderBottom = borderStyle;
            fmt.BorderRight = borderStyle;
            fmt.BorderLeft = borderStyle;
            fmt.BorderTop = borderStyle;

            if (range == null)
            {
                //取前三行的最大列索引
                var colIndex = PoiUtils.GetRow(sheet, 0).LastCellNum;
                if (sheet.GetRow(1) != null && sheet.GetRow(1).LastCellNum > colIndex) colIndex = sheet.GetRow(1).LastCellNum;
                if (sheet.GetRow(2) != null && sheet.GetRow(2).LastCellNum > colIndex) colIndex = sheet.GetRow(2).LastCellNum;
                if (colIndex <= 0) colIndex = 1;

                range = CellRangeAddress.ValueOf($"A1:{PoiUtils.GetColumnName(colIndex - 1)}{(sheet.LastRowNum + 1)}");
            }
            scf.AddConditionalFormatting([range], rule);
        }

        /// <summary>
        /// 合并单元格
        /// </summary>
        public static void AddMergedRegion(ISheet sheet, int startRow, int endRow, int startCol, int endCol, ICellStyle? style = null)
        {
            style ??= PoiUtils.GetCell(sheet, startRow, startCol).CellStyle;
            for (int i = startRow; i <= endRow; i++)
            {
                for (int j = startCol; j <= endCol; j++)
                {
                    PoiUtils.GetCell(sheet, i, j).CellStyle = style;
                }
            }
            var cellRange = new CellRangeAddress(startRow, endRow, startCol, endCol);
            sheet.AddMergedRegion(cellRange);
        }

        /// <summary>
        /// 创建样式
        /// </summary>
        public static ICellStyle CreateCellStyle(ExcelStyle config, IWorkbook workbook, ICellStyle? source = null)
        {
            var style = workbook.CreateCellStyle();
            if (null != source) style.CloneStyleFrom(source);

            // 边框
            if (config.BorderStyle.HasValue)
            {
                if (Enum.TryParse(config.BorderStyle.Value.ToString(), true, out BorderStyle borderStyle))
                {
                    style.BorderBottom = borderStyle;
                    style.BorderLeft = borderStyle;
                    style.BorderRight = borderStyle;
                    style.BorderTop = borderStyle;
                }
                else throw new Exception($"边框样式值【{config.BorderStyle.Value}】不是有效的POI样式。");
            }

            // 位置
            if (config.VerticalAlignment.HasValue)
            {
                if (Enum.TryParse(config.VerticalAlignment.Value.ToString(), true, out VerticalAlignment vAlign))
                {
                    style.VerticalAlignment = vAlign;
                }
                else throw new Exception($"垂直位置值【{config.VerticalAlignment.Value}】不是有效的POI样式。");
            }
            if (config.HorizontalAlignment.HasValue)
            {
                if (Enum.TryParse(config.HorizontalAlignment.Value.ToString(), true, out HorizontalAlignment hAlign))
                {
                    style.Alignment = hAlign;
                }
                else throw new Exception($"水平位置值【{config.HorizontalAlignment.Value}】不是有效的POI样式。");
            }


            // 背景色
            if (config.BackgroundColor.HasValue)
            {
                style.FillForegroundColor = config.BackgroundColor.Value;
                style.FillPattern = FillPattern.SolidForeground;
                style.FillBackgroundColor = config.BackgroundColor.Value;
            }

            // 格式化字符串
            if (!string.IsNullOrWhiteSpace(config.DataFormat))
            {
                var fmt = workbook.CreateDataFormat();
                style.DataFormat = fmt.GetFormat(config.DataFormat);
            }

            // 换行
            if (config.WrapText.HasValue)
            {
                style.WrapText = config.WrapText.Value;
            }

            // 字体
            if (!string.IsNullOrWhiteSpace(config.FontName) || config.FontSize.HasValue || config.FontBold.HasValue || config.FontColor.HasValue)
            {
                var font = workbook.CreateFont();
                if (!string.IsNullOrWhiteSpace(config.FontName)) font.FontName = config.FontName;
                if (config.FontBold.HasValue) font.IsBold = config.FontBold.Value;
                if (config.FontSize.HasValue) font.FontHeightInPoints = config.FontSize.Value;
                if (config.FontColor.HasValue) font.Color = config.FontColor.Value;
                style.SetFont(font);
            }

            return style;
        }

        /// <summary>
        /// 生成IWorkbook,转成字节数组
        /// </summary>
        public static byte[] GetExcelBytes(ExcelConfig config)
        {
            var wb = CreateWorkBook(config, out _);
            return PoiUtils.GetExcelBytes(wb);
        }
    }

    #region 样式工厂类
    /// <summary>
    /// 样式工厂类，根据样式配置ExcelStyle生成新样式，如果已存在样式，则直接返回已存在的样式
    /// </summary>
    public class StyleFactory
    {
        private readonly IWorkbook workbook;
        private readonly Dictionary<string, ICellStyle> createdStyles = [];
        private readonly string defaultStyleKey;

        public StyleFactory(IWorkbook workbook)
        {
            this.workbook = workbook;
            // 创建默认格式
            var defaultCfg = new ExcelStyle() { HorizontalAlignment = (int)HorizontalAlignment.Left, VerticalAlignment = (int)VerticalAlignment.Center };
            defaultStyleKey = defaultCfg.StyleKey;
            var defaultStyle = ExcelHelper.CreateCellStyle(defaultCfg, workbook);
            createdStyles.Add(defaultStyleKey, defaultStyle);
        }

        /// <summary>
        /// 根本配置内容生成格式
        /// 注意不同的格式需要重新调用该方法生成，不能直接修改原格式，否则会影响其他内容的格式
        /// </summary>
        public ICellStyle GetStyle(ExcelStyle config)
        {
            var key = config.StyleKey;
            if (!createdStyles.TryGetValue(key, out var style))
            {
                style = ExcelHelper.CreateCellStyle(config, workbook, DefaultStyle);
            }
            return style;
        }

        /// <summary>
        /// 默认格式：常规格式，水平居左，上下居中
        /// </summary>
        public ICellStyle DefaultStyle { get { return createdStyles[defaultStyleKey]; } }

        /// <summary>
        /// 标题格式：继承Default，居中，加背景色
        /// </summary>
        public ICellStyle TitleStyle
        {
            get
            {
                return GetStyle(new()
                {
                    FontBold = true,
                    HorizontalAlignment = (int)HorizontalAlignment.Center,
                    BackgroundColor = HSSFColor.LemonChiffon.Index,
                });
            }
        }

        /// <summary>
        /// 描述格式：继承Default，自动换行
        /// </summary>
        public ICellStyle DescriptionStyle
        {
            get
            {
                return GetStyle(new()
                {
                    HorizontalAlignment = (int)HorizontalAlignment.Left,
                    //BackgroundColor = HSSFColor.LemonChiffon.Index,
                    WrapText = true,
                });
            }
        }

        /// <summary>
        /// 文本格式：，格式 @
        /// </summary>
        public ICellStyle TextStyle { get { return GetStyle(new() { DataFormat = "@" }); } }

        /// <summary>
        /// 日期格式：居中，格式 yyyy-MM-dd
        /// </summary>
        public ICellStyle DateStyle { get { return GetStyle(new() { DataFormat = "yyyy-MM-dd", HorizontalAlignment = (int)HorizontalAlignment.Center }); } }

        /// <summary>
        /// 时间格式：居中，格式 yyyy-MM-dd HH:mm
        /// </summary>
        public ICellStyle DateTimeStyle { get { return GetStyle(new() { DataFormat = "yyyy-MM-dd HH:mm", HorizontalAlignment = (int)HorizontalAlignment.Center }); } }

        /// <summary>
        /// 数字格式：居中，格式 0.00
        /// </summary>
        public ICellStyle NumbericStyle { get { return GetStyle(new() { DataFormat = "0.00", HorizontalAlignment = (int)HorizontalAlignment.Center }); } }

        /// <summary>
        /// 百分比：居中，格式 0.00%
        /// </summary>
        public ICellStyle PersentStyle { get { return GetStyle(new() { DataFormat = "0.00%", HorizontalAlignment = (int)HorizontalAlignment.Center }); } }
    }
    #endregion
}
