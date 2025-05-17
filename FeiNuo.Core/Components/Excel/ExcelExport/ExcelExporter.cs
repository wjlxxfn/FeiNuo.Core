namespace FeiNuo.Core
{
    public class ExcelExporter
    {
        public byte[] ExportData<T>(IEnumerable<T> dataList)
        {
            var poi = new PoiExcel();
            if (dataList.Any())
            {
                int rowIndex = 0;
                var first = dataList.First();
                if (first != null)
                {
                    var propertyInfos = first.GetType().GetProperties();
                    poi.AddTitleRow(rowIndex++, [.. propertyInfos.Select(a => a.Name)]);
                    foreach (var data in dataList)
                    {
                        var values = propertyInfos.Select(a => a.GetValue(data)).ToArray();
                        poi.AddDataRow(rowIndex++, values);
                    }
                }
            }
            return poi.GetExcelBytes();
        }

        //public byte[] ExportData<T>(Dictionary<string, IEnumerable<object>> dataMap)
        //{
        //    var poi = new PoiExcel();
        //    poi.Workbook.RemoveSheetAt(0);
        //    if (dataList.Any())
        //    {
        //        int rowIndex = 0;
        //        var first = dataList.First();
        //        if (first != null)
        //        {
        //            var propertyInfos = first.GetType().GetProperties();
        //            poi.AddTitleRow(rowIndex++, [.. propertyInfos.Select(a => a.Name)]);
        //            foreach (var data in dataList)
        //            {
        //                var values = propertyInfos.Select(a => a.GetValue(data)).ToArray();
        //                poi.AddDataRow(rowIndex++, values);
        //            }
        //        }
        //    }
        //    return poi.GetExcelBytes();
        //}
    }

    /*
     
     using FeiNuo.Core.Utilities;
using NPOI.SS.UserModel;
using System.Data;

namespace FeiNuo.Core.Excel;

/// <summary>
/// Excel导入导出相关辅助类
/// </summary>
public class ExcelHelper
{
    #region 基础数据：把基础数据(DataSet或List)生成Excel

    /// <summary>
    /// 把基础数据(DataSet)生成Excel
    /// </summary>
    public static IWorkbook DataTableToExcel(DataTable dt, string sheetName = "Sheet1", bool xlsx = true)
    {
        var ds = new DataSet();
        ds.Tables.Add(dt);
        return DataSetToExcel(ds, sheetName, xlsx);
    }

    /// <summary>
    /// 根据DataSet内容生成Excel
    /// </summary>
    /// <param name="ds">数据集</param>
    /// <param name="sheetNames">Excel工作表的名，多个以英文逗号分隔，需要和DataTable一一对应，如果不传，默认取DataTable的TableName</param>
    /// <param name="xlsx">是否2007格式</param>
    public static IWorkbook DataSetToExcel(DataSet ds, string sheetNames = "", bool xlsx = true)
    {
        if (ds == null || ds.Tables.Count == 0) return PoiUtils.CreateWorkbook(xlsx);

        if (string.IsNullOrWhiteSpace(sheetNames))
        {
            sheetNames = "";
            foreach (DataTable dt in ds.Tables) sheetNames += dt.TableName + ",";
            sheetNames = sheetNames.Trim(',');
        }
        else if (sheetNames.Split(',').Length != ds.Tables.Count)
        {
            throw new ArgumentException("工作表名称的数量必须和数据表的个数一致");
        }

        var names = sheetNames.Split(',');
        var exporter = new ExcelExporter("", ExcelExporter.DEFAULT_COLUMN_WIDTH, xlsx);
        for (int i = 0; i < ds.Tables.Count; i++)
        {
            exporter.CreateSheet(names[i]);
            var table = ds.Tables[i];
            var titles = new List<string>();
            foreach (DataColumn col in table.Columns)
            {
                titles.Add(col.ColumnName);
            }
            int rowIndex = 0;
            exporter.AddTitleRow(rowIndex++, [.. titles]);
            foreach (DataRow row in table.Rows)
            {
                exporter.AddDataRow(rowIndex++, 0, row.ItemArray!);
            }
        }
        return exporter.Workbook;
    }

    /// <summary>
    /// 根据数据集生成Excel
    /// </summary>
    /// <param name="dataMap">数据集，字典键对应SheetName，</param>
    /// <param name="xlsx">Excel格式</param>
    public static IWorkbook DataListToExcel(Dictionary<string, IEnumerable<object>> dataMap, bool xlsx = true)
    {
        if (dataMap.Count == 0) return PoiUtils.CreateWorkbook(xlsx);

        var exporter = new ExcelExporter("", ExcelExporter.DEFAULT_COLUMN_WIDTH, xlsx);
        foreach (var entry in dataMap)
        {
            exporter.CreateSheet(entry.Key);
            if (!entry.Value.Any()) continue;
            // 反射属性
            var first = entry.Value.First();
            var propertyInfos = first.GetType().GetProperties();
            // 标题
            var titles = propertyInfos.Select(a => a.Name).ToArray();
            exporter.AddTitleRow(0, 0, titles);
            // 内容
            var rowIndex = 1;
            foreach (var data in entry.Value)
            {
                var values = propertyInfos.Select(a => a.GetValue(data)).ToArray();
                exporter.AddDataRow(rowIndex++, 0, values);
            }
        }
        return exporter.Workbook;
    }
    #endregion

    #region 下载模板：根据导入配置生成模板Excel

    /// <summary>
    /// 根据导入配置生成模板Excel
    /// </summary>
    public static IWorkbook CreateTemplateFromImportConfig(ExcelImportConfig cfg)
    {
        var exporter = new ExcelExporter("", xlsx: cfg.IsTemplateXlsx);
        foreach (var sheetConfig in cfg.SheetConfigs)
        {
            int rowIndex = 0, colIndex = 0;
            var sheet = exporter.CreateSheet(sheetConfig.SheetName);

            var tooltipMap = new Dictionary<int, string>();
            var patriarch = sheet.CreateDrawingPatriarch(); // 添加批注用
            foreach (var column in sheetConfig.ColumnConfigs)
            {
                if (!string.IsNullOrEmpty(column.DataFormat)) // 设置数据格式
                {
                    var style = PoiUtils.CreateCellStyle(exporter.Workbook, borderStyle: BorderStyle.None, dataFormat: column.DataFormat);
                    exporter.SetDefaultColumnStyle(colIndex, style);
                }
                if (column.Width > 0) // 设置列的宽度
                {
                    exporter.SetColumnWidth(colIndex, column.Width);
                }
                // 记录批注，需要在生成标题后在填写标注，不然合并单元格会有影响
                if (!string.IsNullOrWhiteSpace(column.ToolTip))
                {
                    tooltipMap.Add(colIndex, column.ToolTip);
                }
                colIndex++;
            }

            // 添加说明行
            if (sheetConfig.HasDescription)
            {
                exporter.AddRemarkRow(rowIndex++, 0, sheetConfig.Description!, sheetConfig.ColumnConfigs.Count);
            }

            // 添加标题行
            rowIndex = exporter.AddTitleRow(rowIndex, 0, sheetConfig.ColumnConfigs.Select(a => a.ColumnTitle).ToArray());

            // 添加批注
            foreach (var entry in tooltipMap.AsEnumerable())
            {
                IClientAnchor anchor = PoiUtils.CreateClientAnchor(0, 0, 0, 0, entry.Key, rowIndex, entry.Key + 2, rowIndex + 5, cfg.IsTemplateXlsx);
                var comment = patriarch.CreateCellComment(anchor);
                comment.String = PoiUtils.CreateRichTextString(entry.Value, cfg.IsTemplateXlsx);
                var cell = sheet.GetRow(rowIndex).GetCell(entry.Key);
                // 如果上下行合并了，把批注加在第一行里
                // 如果有超过两行的，下面几行合并的还是会有问题，只能显示在第一行里。使用场景极少，不处理了
                if (cell.IsMergedCell) cell = sheet.GetRow(sheetConfig.TitleRowStart).GetCell(entry.Key);
                cell.CellComment = comment;
            }
        }

        return exporter.Workbook;
    }
    #endregion

    #region 执行导入：根据导入配置检查上传的Excel模板，并转成相应的数据集DataSet或List
    /// <summary>
    /// 根据上传的excel和导入配置检查模板是否正确
    /// </summary>
    public static void ValidateExcelTemplate(IWorkbook workbook, ExcelImportConfig cfg)
    {
        if (cfg.SheetConfigs.Count == 0) return;

        // 检查sheet数量，至少不能少于配置的数量
        if (workbook.NumberOfSheets < cfg.SheetConfigs.Count)
        {
            throw new Exception("导入模板检查失败，请重新下载最新模板导入！");
        }
        for (int i = 0; i < cfg.SheetConfigs.Count; i++)
        {
            var sheet = workbook.GetSheetAt(0);
            var sheetCfg = cfg.SheetConfigs[i];
            if (!sheetCfg.ValidateTemplate) continue;

            // 最后一行标题行，
            var titleRow = PoiUtils.GetRow(sheet, sheetCfg.TitleRowIndex);
            var rowTitles = string.Join(",", titleRow.Cells.Select(PoiUtils.GetCellValue<string>).ToList());
            var cfgTitles = string.Join(",", sheetCfg.ColumnConfigs.Select(a => a.RowTitles.Last()).ToList());
            if (rowTitles != cfgTitles)
            {
                throw new Exception("导入模板检查失败：列标题不对应，请不要修改列标题或下载最新模板！<br/>  正确模板：" + cfgTitles);
            }
        }
    }

    /// <summary>
    /// excel数据转成DataSet
    /// </summary>
    public static DataSet GetDataSetFromImportExcel(IWorkbook workbook, ExcelImportConfig cfg)
    {
        var ds = new DataSet();
        string errMsg = "", rowMsg, keyValue;
        Dictionary<string, int> keyMap;
        for (int i = 0; i < cfg.SheetConfigs.Count; i++)
        {
            var sheet = workbook.GetSheetAt(i);
            var sheetCfg = cfg.SheetConfigs[i];
            keyMap = [];

            // 计算公式
            sheet.ForceFormulaRecalculation = true;
            var dt = new DataTable(sheet.SheetName);
            #region 创建DataTable
            foreach (var colCfg in sheetCfg.ColumnConfigs)
            {
                dt.Columns.Add(new DataColumn(colCfg.DataField, colCfg.DataType));
            }
            #endregion

            for (var rowIndex = sheetCfg.DataRowIndex; rowIndex <= sheet.LastRowNum; rowIndex++)
            {
                rowMsg = "";
                keyValue = "";
                var row = PoiUtils.GetRow(sheet, rowIndex);
                var rowData = new List<object?>();
                for (var colIndex = 0; colIndex < sheetCfg.ColumnConfigs.Count; colIndex++)
                {
                    var colCfg = sheetCfg.ColumnConfigs[colIndex];
                    var value = PoiUtils.GetCellValue(row.GetCell(colIndex));
                    var parseValue = colCfg.ConvertValueFromExcel(value, out string msg);
                    if (!string.IsNullOrWhiteSpace(msg))
                    {
                        rowMsg += $"列【{colCfg.RowTitles.Last()}】{msg}；";
                    }
                    rowData.Add(parseValue);
                    if (colCfg.UniqueKey)
                    {
                        keyValue += (parseValue ?? "null") + "|";
                    }
                }
                if (rowMsg != "")
                {
                    errMsg += $"第【{rowIndex}】行：" + rowMsg + "<br/>";
                }
                else
                {
                    if (keyValue != "")
                    {
                        if (keyMap.TryGetValue(keyValue, out int value))
                        {
                            errMsg += $"第【{rowIndex + 1}】行：与{value + 1}行重复,重复键值：{keyValue}； <br/>";
                        }
                        else keyMap.Add(keyValue, rowIndex);
                    }
                    dt.Rows.Add([.. rowData]);
                }
            }
            if (errMsg != "")
            {
                throw new MessageException($"获取Excel数据出错【{sheet.SheetName}】:<br/>" + errMsg);
            }
            ds.Tables.Add(dt);
        }
        return ds;
    }

    /// <summary>
    /// excel数据转成集合
    /// </summary>
    public static List<T> GetDataListFromImportExcel<T>(ISheet sheet, ExcelSheetConfig sheetCfg) where T : class, new()
    {
        // 实体类型的属性
        var propMap = typeof(T).GetProperties().ToDictionary(a => a.Name, v => v);

        string errMsg = "", rowMsg, keyValue;
        var keyMap = new Dictionary<string, int>();

        var lstData = new List<T>();
        // 计算公式
        sheet.ForceFormulaRecalculation = true;
        for (var rowIndex = sheetCfg.DataRowIndex; rowIndex <= sheet.LastRowNum; rowIndex++)
        {
            rowMsg = "";
            keyValue = "";

            var row = PoiUtils.GetRow(sheet, rowIndex);
            var rowData = Activator.CreateInstance(typeof(T)) as T;
            for (var colIndex = 0; colIndex < sheetCfg.ColumnConfigs.Count; colIndex++)
            {
                var colCfg = sheetCfg.ColumnConfigs[colIndex];
                var value = PoiUtils.GetCellValue(row.GetCell(colIndex));
                var parseValue = colCfg.ConvertValueFromExcel(value, out string msg);
                if (!string.IsNullOrWhiteSpace(msg))
                {
                    rowMsg += $"列【{colCfg.RowTitles.Last()}】{msg}；";
                }
                else
                {
                    propMap[colCfg.DataField].SetValue(rowData, parseValue);
                }
                if (colCfg.UniqueKey)
                {
                    keyValue += (parseValue ?? "null") + "|";
                }
            }
            if (rowMsg != "")
            {
                errMsg += $"第【{rowIndex}】行：" + rowMsg + "<br/>";
            }
            else if (null != rowData)
            {
                if (keyValue != "")
                {
                    if (keyMap.TryGetValue(keyValue, out int value))
                    {
                        errMsg += $"第【{rowIndex + 1}】行：与{value + 1}行重复,重复键值：{keyValue}； <br/>";
                    }
                    else keyMap.Add(keyValue, rowIndex);
                }
                lstData.Add(rowData);
            }
        }
        if (errMsg != "")
        {
            throw new MessageException($"获取Excel数据出错:<br/>" + errMsg);
        }

        return lstData;
    }
    #endregion
}

     */
}
