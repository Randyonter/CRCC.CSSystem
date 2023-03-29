/****************************************************************************************************************************************************************
 *
 * 版权所有: 2023  All Rights Reserved.
 * 文 件 名:  ExcelHelper
 * 描    述:
 * 创 建 者：  Ltf
 * 创建时间：  2023/3/29 15:35:05
 * 历史更新记录:
 * =============================================================================================
*修改标记
*修改时间：2023/3/29 15:35:05
*修改人： Ltf
*版本号： V1.0.0.0
*描述：
 * =============================================================================================

****************************************************************************************************************************************************************/

using CRCC.SystemTools.Log;
using Spire.Xls;
using Spire.Xls.Core.Spreadsheet.Shapes;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;

namespace CRCC.SystemTools.FileHelper
{
    public static class ExcelHelper
    {
        private static void setCellValue(CellRange range, object value, bool setStyle)
        {
            range.Value = "'" + value;
            if (setStyle)
            {
                range.RowHeight = 20;
                range.HorizontalAlignment = HorizontalAlignType.Center;
                range.VerticalAlignment = VerticalAlignType.Center;
                range.BorderAround(LineStyleType.Thin, Color.Black);
            }
        }

        private static List<Dictionary<int, string>> getSheetData(Worksheet sheet)
        {
            var result = new List<Dictionary<int, string>>();
            if (sheet == null)
                return result;
            var rowcount = sheet.Rows.Length;
            var colcount = sheet.Columns.Length;
            if (sheet.Rows.Length > 10000 || sheet.Columns.Length > 200)
                return result;
            for (int i = 0; i < rowcount; i++)
            {
                var dic = new Dictionary<int, string>();
                for (int j = 0; j < colcount; j++)
                {
                    var value = sheet.Range[i + 1, j + 1].DisplayedText;
                    dic.Add(j, value);
                }
                result.Add(dic);
            }
            return result;
        }

        /// <summary>
        /// 将数据导出到Excle中
        /// </summary>
        /// <param name="values">需要导出的数据，Key为行号(从0开始)，Value为当前行列的数据（其中Value的Key为列号，从0开始。Value的Value为单元格中的值）</param>
        /// <param name="desFile">导出后的文件路径</param>
        /// <param name="sheetIndex">要导出的Sheet页号，从0开始；(默认为0)</param>
        /// <param name="mergeRange">合并单元格的区域，区域行列号都从0开始，默认没有</param>
        /// <param name="templatePath">模板文件全路径（为空时系统自动创建文件），默认空</param>
        /// <returns></returns>
        [Description("将数据导出到Excle中;values:(数据源,Key为行号,从0开始;Value为行数据(key为列号,从0开始;value为单元格值))；sheetIndex:要导出的Sheet页号,从0开始;(默认为0)；templatePath:模板文件全路径（为空时系统自动创建文件）")]
        public static Result<bool> ToExcel(this Dictionary<int, Dictionary<int, string>> values, string desFile, int sheetIndex = 0, List<ExcelRectangle> mergeRange = null, string templatePath = null)
        {
            if (values == null || values.Any() == false)
                return new Result<bool>(false, false, "无可导出的数据！");
            if (string.IsNullOrEmpty(desFile))
                return new Result<bool>(false, false, "文件路径为空！");
            try
            {
                var workbook = new Workbook();//创建EXCEL对象
                if (File.Exists(templatePath))
                    workbook.LoadTemplateFromFile(templatePath, true);
                if (workbook.Worksheets.Count < sheetIndex + 1)
                    return new Result<bool>(false, false, "未找到指定的Sheet页！");
                var sheet = workbook.Worksheets[sheetIndex];
                var columncount = values.Max(s => s.Value.Count);
                for (int i = 1; i <= columncount; i++)
                    sheet.SetColumnWidth(i, 15);
                bool setstyle = string.IsNullOrEmpty(templatePath) || !File.Exists(templatePath);

                #region 填写单元格数据

                foreach (var row in values)
                {
                    if (row.Key < 0 || row.Value == null || row.Value.Count == 0)
                        continue;
                    foreach (var col in row.Value)
                    {
                        if (col.Key < 0)
                            continue;
                        //设置单元格的值
                        setCellValue(sheet.Range[row.Key + 1, col.Key + 1], col.Value, setstyle);
                    }
                }

                #endregion 填写单元格数据

                #region 合并单元格

                if (mergeRange != null && mergeRange.Count > 0)
                {
                    foreach (var mer in mergeRange)
                    {
                        var range = sheet.Range[mer.StartRow, mer.StartColumn, mer.EndRow, mer.EndColumn];
                        range.Merge();
                        range.RowHeight = 20;
                    }
                }

                //移除多余的Sheet页面
                var sheets = workbook.Worksheets.ToList();
                foreach (var item in sheets)
                {
                    if (item.Cells.Length == 0)
                        workbook.Worksheets.Remove(item);
                }

                #endregion 合并单元格

                workbook.SaveToFile(desFile);
                return new Result<bool>(true, true, desFile);
            }
            catch (Exception ex)
            {
                LogHelper.WriteExceptionLog("导出EXCEL数据失败", "导出EXCEL失败：" + ex);
                return new Result<bool>(false, false, ex.Message);
            }
        }

        /// <summary>
        /// 将数据导出到Excle中
        /// </summary>
        /// <param name="values">数据列表（Item为行数据；Item中Dictionary为列数据，其中Key为列名称，Value为单元格数据）</param>
        /// <param name="desFile">导出后的文件路径</param>
        /// <param name="sheetIndex">要导出的Sheet页号，从0开始；(默认为0)</param>
        /// <param name="mergeRange">合并单元格的区域，区域行列号都从0开始，默认没有</param>
        /// <param name="templatePath">模板文件全路径（为空时系统自动创建文件），默认空</param>
        /// <returns></returns>
        [Description("将数据导出到Excle中;values:(数据源,key为列号,从0开始;value为单元格值)；sheetIndex:要导出的Sheet页号,从0开始;(默认为0)；templatePath:模板文件全路径（为空时系统自动创建文件）")]
        public static Result<bool> ToExcel(this List<Dictionary<string, string>> values, string desFile, int sheetIndex = 0, List<ExcelRectangle> mergeRange = null, string templatePath = null)
        {
            if (values == null || values.Any() == false)
                return new Result<bool>(false, false, "无可导出的数据！");
            var values2 = new Dictionary<int, Dictionary<int, string>>();
            var heads = values.Select(s => s.Keys).Merge().Distinct().ToList();
            var columns = heads.ToDictionary(s => s, s => heads.IndexOf(s));
            for (int index = 0; index < values.Count; index++)
            {
                var value = values[index - 1];
                values2.Add(index, value.ToDictionary(s => columns[s.Key], s => s.Value));
            }
            return ToExcel(values2, desFile, sheetIndex, mergeRange, templatePath);
        }

        /// <summary>
        /// 将数据导出到Excle中
        /// </summary>
        /// <param name="values">数据列表（Item为行数据；Item中Dictionary为列数据，其中Key为列名称，Value为单元格数据）</param>
        /// <param name="desFile">导出后的文件路径</param>
        /// <param name="sheetIndex">要导出的Sheet页号，从0开始；(默认为0)</param>
        /// <param name="mergeRange">合并单元格的区域，区域行列号都从0开始，默认没有</param>
        /// <param name="templatePath">模板文件全路径（为空时系统自动创建文件），默认空</param>
        /// <returns></returns>
        [Description("将数据导出到Excle中;values:数据源；sheetIndex:要导出的Sheet页号,从0开始;(默认为0)；templatePath:模板文件全路径（为空时系统自动创建文件）")]
        public static Result<bool> ToExcel(this List<List<string>> values, string desFile, int sheetIndex = 0, List<ExcelRectangle> mergeRange = null, string templatePath = null)
        {
            if (values == null || values.Any() == false)
                return new Result<bool>(false, false, "无可导出的数据！");
            var values2 = new Dictionary<int, Dictionary<int, string>>();
            for (int i = 0; i < values.Count; i++)
            {
                var valuedic = new Dictionary<int, string>();
                var value = values[i];
                for (int j = 0; j < value.Count; j++)
                    valuedic.Add(j, value[j]);
                values2.Add(i, valuedic);
            }
            return ToExcel(values2, desFile, sheetIndex, mergeRange, templatePath);
        }

        /// <summary>
        /// 将数据导出到Excle中
        /// </summary>
        /// <param name="tb">待导出的数据表</param>
        /// <param name="desFile">导出后的文件路径</param>
        /// <param name="includeHead">指示是否要导出DataTable的表头到Exlce</param>
        /// <param name="sheetIndex">要导出的Sheet页号，从0开始；(默认为0)</param>
        /// <param name="mergeRange">合并单元格的区域，区域行列号都从0开始，默认没有</param>
        /// <param name="templatePath">模板文件全路径（为空时系统自动创建文件），默认空</param>
        /// <returns></returns>
        [Description("将数据导出到Excle中;tb:数据源；includeHead:指示是否要导出DataTable的表头到Exlce；sheetIndex:要导出的Sheet页号,从0开始;(默认为0)；templatePath:模板文件全路径（为空时系统自动创建文件）")]
        public static Result<bool> ToExcel(this DataTable tb, string desFile, bool includeHead = true, int sheetIndex = 0, List<ExcelRectangle> mergeRange = null, string templatePath = null)
        {
            if (tb == null || tb.Rows.Count == 0)
                return new Result<bool>(false, false, "无可导出的数据！");
            //所有的列
            var columns = tb.Columns.OfType<DataColumn>().ToDictionary(s => s.ColumnName, s => string.IsNullOrEmpty(s.Caption) ? s.ColumnName : s.Caption);
            var datas = tb.Rows.OfType<DataRow>().Select(s => columns.Select(m => s[m.Key].ToString()).ToList()).ToList();
            if (includeHead)
                datas.Insert(0, columns.Select(s => s.Value).ToList());
            return ToExcel(datas, desFile, sheetIndex, mergeRange, templatePath);
        }

        [Description("按数据源格式导出到Exel中，包括行高列宽,单元格样式等(行列号从1开始)")]
        public static Result<string> ToExcel(this List<ExcelSheet> sheets, string desFile, bool includeHead = true, bool deleteSheet = true)
        {
            if (sheets == null || sheets.Any() == false)
                return new Result<string>(false, null, "无可导出的数据！");
            try
            {
                desFile = string.IsNullOrEmpty(desFile) ? Path.Combine(DataLibrary.TempDirectory, Guid.NewGuid().ToString() + ".xlsx") : desFile;
                var workbook = new Workbook();//创建EXCEL对象
                if (File.Exists(desFile))
                    File.Delete(desFile);
                var xsheets = workbook.Worksheets.OfType<Worksheet>().ToList();
                var newsheets = new List<string>();
                foreach (var sheet in sheets)
                {
                    var sheetname = string.IsNullOrEmpty(sheet.SheetName) ? ("Sheet" + (sheets.IndexOf(sheet) + 1)) : sheet.SheetName;
                    var xsheet = xsheets.FirstOrDefault(s => s.Name == sheetname);
                    if (xsheet == null)
                    {
                        xsheet = workbook.Worksheets.Add(sheetname);
                        xsheets.Add(xsheet);
                    }
                    newsheets.Add(sheetname);
                    var rowindex = 1;

                    #region 导出表头

                    if (includeHead && sheet.Columns.Any())
                    {
                        var colheight = sheet.ColumnHeight <= 0 ? 20 : sheet.ColumnHeight;
                        //设置合并表头的行
                        if (sheet.SpanHeaders.Any())
                        {
                            foreach (var sh in sheet.SpanHeaders)
                            {
                                for (int i = sh.Item2; i <= sh.Item3; i++)
                                    setCellValue(xsheet.Range[rowindex, i], sh.Item1, true);
                                xsheet.Range[rowindex, sh.Item2, rowindex, sh.Item3].Merge();
                            }
                            xsheet.SetRowHeight(rowindex, colheight);
                            rowindex++;
                        }
                        if (sheet.Columns.Any())
                        {
                            foreach (var col in sheet.Columns)
                            {
                                setCellValue(xsheet.Range[rowindex, col.ColumnIndex], col.ColumnText, true);
                                var width = col.ColumnWidth <= 0 ? 10 : col.ColumnWidth;
                                xsheet.SetColumnWidth(col.ColumnIndex, width);
                            }
                            xsheet.SetRowHeight(rowindex, colheight);
                            rowindex++;
                        }
                    }

                    #endregion 导出表头

                    #region 导出表格数据

                    foreach (var row in sheet.Rows)
                    {
                        var height = row.RowHeight <= 0 ? 20 : row.RowHeight;
                        foreach (var cell in row.Cells)
                        {
                            var xcell = xsheet.Range[rowindex, cell.ColumnIndex];
                            setCellValue(xcell, cell.Value, true);
                            xcell.Style.Color = cell.BackColor;
                            xcell.Style.Font.Color = cell.FontColor;
                            xcell.Style.Font.Size = cell.FontSize;
                            if (string.IsNullOrEmpty(cell.CommentText) == false)
                            {
                                var comment = xcell.AddComment();
                                comment.Width = cell.CommentWidth;
                                comment.Height = cell.CommentHeigth;
                                comment.Text = cell.CommentText;
                            }
                        }
                        xsheet.SetRowHeight(rowindex, height);
                        rowindex++;
                    }

                    #endregion 导出表格数据

                    #region 添加合并单元格

                    var headcount = includeHead == false || sheet.Columns.Any() == false ? 0 : sheet.SpanHeaders.Any() ? 2 : 1;
                    if (sheet.MergeRectangles != null && sheet.MergeRectangles.Any())
                    {
                        var mergers = headcount == 0 ? sheet.MergeRectangles.Where(s => s.StartRow > 0).ToList() : sheet.MergeRectangles;
                        foreach (var mer in mergers)
                            xsheet.Range[mer.StartRow + headcount, mer.StartColumn, mer.EndRow + headcount, mer.EndColumn].Merge();
                    }

                    #endregion 添加合并单元格
                }
                if (deleteSheet)
                {
                    foreach (var xs in xsheets)
                    {
                        if (newsheets.Contains(xs.Name) == false)
                            workbook.Worksheets.Remove(xs);
                    }
                }
                workbook.SaveToFile(desFile);
                return new Result<string>(true, desFile);
            }
            catch (Exception ex)
            {
                LogHelper.WriteExceptionLog("根据ExcelSheet导出数据到Excel异常！", "根据ExcelSheet导出数据到Excel异常：" + ex);
                return new Result<string>(false, null, ex.Message);
            }
        }

        [Description("读取Excel文件中指定索引号Sheet页中的数据（sheetIndex,Sheet页索引号，从1开始）")]
        public static Result<List<Dictionary<int, string>>> ReadExcel(string filePath, int sheetIndex)
        {
            if (string.IsNullOrEmpty(filePath) || File.Exists(filePath) == false)
                return new Result<List<Dictionary<int, string>>>(false, null, "文件不存在！");
            try
            {
                var workbook = new Workbook();//创建EXCEL对象
                workbook.LoadFromFile(filePath, false);
                if (workbook.Worksheets.Count < sheetIndex)
                    return new Result<List<Dictionary<int, string>>>(false, null, "未找到指定的Sheet页!");
                var sheet = workbook.Worksheets[sheetIndex - 1];
                var result = getSheetData(sheet);
                return new Result<List<Dictionary<int, string>>>(true, result);
            }
            catch (Exception ex)
            {
                LogHelper.WriteExceptionLog("读取EXCEL异常！", "读取EXCEL异常：" + ex);
                return new Result<List<Dictionary<int, string>>>(false, null, ex.Message);
            }
        }

        [Description("读取EXCEL中所有数据（Key：Sheet页名称；Value：Excel数据）")]
        public static Result<Dictionary<string, List<Dictionary<int, string>>>> ReadExcel(string filePath)
        {
            if (string.IsNullOrEmpty(filePath) || File.Exists(filePath) == false)
                return new Result<Dictionary<string, List<Dictionary<int, string>>>>(false, null, "文件不存在!");
            try
            {
                var result = new Dictionary<string, List<Dictionary<int, string>>>();
                var workbook = new Workbook();//创建EXCEL对象
                workbook.LoadFromFile(filePath, false);
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    if (sheet.Cells.Length == 0 || sheet.Rows.Length > 10000 || sheet.Columns.Length > 200)
                        continue;
                    var name = sheet.Name;
                    var data = getSheetData(sheet);
                    result.Add(name, data);
                }
                return new Result<Dictionary<string, List<Dictionary<int, string>>>>(true, result);
            }
            catch (Exception ex)
            {
                LogHelper.WriteExceptionLog("读取EXCEL异常！", "读取EXCEL异常：" + ex);
                return new Result<Dictionary<string, List<Dictionary<int, string>>>>(false, null, ex.Message);
            }
        }

        [Description("读取EXCEL中所有图片（Key为Sheet页名称，Values为当前Sheet页中所有图片对象（图片所在区域，Tag中存着图片））")]
        public static Result<Dictionary<string, List<ExcelRectangle>>> ReadExcelPictures(string filePath)
        {
            if (string.IsNullOrEmpty(filePath) || File.Exists(filePath) == false)
                return new Result<Dictionary<string, List<ExcelRectangle>>>(false, null, "文件不存在！");
            try
            {
                var result = new Dictionary<string, List<ExcelRectangle>>();
                var workbook = new Workbook();//创建EXCEL对象
                workbook.LoadFromFile(filePath, false);
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    var pictures = sheet.Pictures?.OfType<XlsBitmapShape>().ToList();
                    if (pictures == null || pictures.Count == 0)
                        continue;
                    var name = sheet.Name;
                    var recs = pictures.Select(s => new ExcelRectangle(s.TopRow, s.LeftColumn, s.BottomRow, s.RightColumn) { Tag = s.Picture }).ToList();
                    result.Add(name, recs);
                }
                return new Result<Dictionary<string, List<ExcelRectangle>>>(true, result);
            }
            catch (Exception ex)
            {
                LogHelper.WriteExceptionLog("读取EXCEL中图片异常！", "读取EXCEL中图片异常：" + ex);
                return new Result<Dictionary<string, List<ExcelRectangle>>>(false, null, ex.Message);
            }
        }

        [Description("通过SQL查询的方式读取EXCEL文件（读取性能高，需要安装AccessDatabaseEngine插件）")]
        public static Result<DataSet> ReadExcelToSet(string filePath)
        {
            try
            {
                string strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties='Excel 12.0;HDR=No;IMEX=1;'";
                var con = new OleDbConnection(strConn);
                if (con.State != ConnectionState.Open)
                    con.Open();
                var set = new DataSet();
                var sheettable = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables_Info, null);
                var sheetNames = sheettable.Rows.OfType<DataRow>().Where(s => s["CARDINALITY"].ToString() == "1").Select(s => s["TABLE_NAME"].ToString().Trim('\'')).ToList();
                foreach (var name in sheetNames)
                {
                    if (name.Contains("$") == false)
                        continue;
                    var sheetsql = $"select * from [{name}]";
                    var sheetadapter = new OleDbDataAdapter(sheetsql, con);
                    var dataTable = new DataTable() { TableName = name };
                    sheetadapter.Fill(dataTable);
                    set.Tables.Add(dataTable);
                }
                con.Close();
                return new Result<DataSet>(true, set);
            }
            catch (Exception ex)
            {
                LogHelper.WriteExceptionLog("查询Excel文件异常！", "查询Excel文件异常：" + ex);
                return new Result<DataSet>(false, null, ex.Message);
            }
        }

        [Description("通过SQL查询的方式读取EXCEL文件（读取性能高，需要安装AccessDatabaseEngine插件），Key为Sheet页名称，Value为行数据")]
        public static Result<Dictionary<string, List<Dictionary<int, string>>>> ReadExcel2(string filePath)
        {
            try
            {
                var set = ReadExcelToSet(filePath);
                if (set.IsSuccess == false)
                    return new Result<Dictionary<string, List<Dictionary<int, string>>>>(false, null, set.Message);
                var result = new Dictionary<string, List<Dictionary<int, string>>>();
                foreach (DataTable tb in set.Value.Tables)
                {
                    var sheetname = tb.TableName?.Replace("$", "").TrimStart('\'').TrimEnd('\'');
                    if (string.IsNullOrEmpty(sheetname))
                        continue;
                    var columns = tb.Columns.OfType<DataColumn>().ToList();
                    var colindex = columns.ToDictionary(s => columns.IndexOf(s), s => s.ColumnName);
                    var datas = tb.Rows.OfType<DataRow>().Select(s => colindex.ToDictionary(m => m.Key, m => s[m.Value].ToString())).ToList();
                    result.Add(sheetname, datas);
                }
                return new Result<Dictionary<string, List<Dictionary<int, string>>>>(true, result);
            }
            catch (Exception ex)
            {
                LogHelper.WriteExceptionLog("查询Excel文件异常！", "查询Excel文件异常：" + ex);
                return new Result<Dictionary<string, List<Dictionary<int, string>>>>(false, null, ex.Message);
            }
        }

        [Description("获取Excel的Sheet页名称和索引(Key为索引号，Value为名称)；includeHide:是否读取被隐藏的Sheet页")]
        public static Dictionary<int, string> GetSheetNames(string filePath, bool includeHide = false)
        {
            try
            {
                if (string.IsNullOrEmpty(filePath) || File.Exists(filePath) == false)
                    return new Dictionary<int, string>();
                var workbook = new Workbook();//创建EXCEL对象
                workbook.LoadFromFile(filePath, false);
                var sheets = workbook.Worksheets.ToList();
                return (includeHide ? sheets : sheets.Where(s => s.Visibility == WorksheetVisibility.Visible).ToList()).ToDictionary(s => s.Index, s => s.Name);
            }
            catch (Exception ex)
            {
                LogHelper.WriteExceptionLog("读取Excel文件Sheet页名称异常！", "读取Excel文件Sheet页名称异常：" + ex);
                return new Dictionary<int, string>();
            }
        }

        [Description("读取Excel中所有数据(包括单元格样式)")]
        public static Result<ExcelData> ReadExcelData(string filePath)
        {
            if (string.IsNullOrEmpty(filePath) || File.Exists(filePath) == false)
                return new Result<ExcelData>(false, null, "Excel文件不存在！");
            try
            {
                var exceldata = new ExcelData();
                exceldata.FileName = Path.GetFileName(filePath);
                var workbook = new Workbook();//创建EXCEL对象
                workbook.LoadFromFile(filePath, false);
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    if (sheet.Cells.Length == 0)
                        continue;
                    var excelsheet = new ExcelSheet() { SheetName = sheet.Name, SheetIndex = sheet.Index };
                    excelsheet.IsHide = !(sheet.Visibility == WorksheetVisibility.Visible);
                    excelsheet.SheetIndex = sheet.Index;
                    exceldata.Sheets.Add(excelsheet);
                    var columncount = sheet.Columns.Length;
                    var rowcount = sheet.Rows.Length;
                    for (int j = 0; j < columncount; j++)
                        excelsheet.Columns.Add(new ExcelColumn() { ColumnIndex = j + 1, ColumnWidth = sheet.Columns[j].ColumnWidth, ColumnName = sheet.Columns[j].DisplayedText });
                    for (int i = 0; i < rowcount; i++)
                    {
                        var rowindex = i + 1;
                        var row = new ExcelRow() { RowIndex = rowindex };
                        row.RowHeight = sheet.GetRowHeight(rowindex);
                        excelsheet.Rows.Add(row);
                        for (int j = 0; j < columncount; j++)
                        {
                            var colindex = j + 1;
                            var range = sheet.Range[rowindex, colindex];
                            var value = range.DisplayedText;
                            var cell = new ExcelCell();
                            cell.RowIndex = rowindex;
                            cell.ColumnIndex = colindex;
                            cell.Value = value;
                            if (range.Style != null)
                            {
                                cell.BackColor = range.Style.Color;
                                cell.FontColor = range.Style.Font.Color;
                                cell.FontSize = range.Style.Font.Size;
                            }
                            if (range.HasComment)
                            {
                                cell.CommentText = range.Comment.Text;
                                cell.CommentWidth = range.Comment.Width;
                                cell.CommentHeigth = range.Comment.Height;
                            }
                            row.Cells.Add(cell);
                        }
                    }
                    //读取合并单元格的信息
                    var mergeds = sheet.MergedCells;
                    foreach (var merged in mergeds)
                    {
                        var mergerec = new ExcelRectangle(merged.Row, merged.Column, merged.LastRow, merged.LastColumn);
                        excelsheet.MergeRectangles.Add(mergerec);
                    }
                }
                return new Result<ExcelData>(true, exceldata);
            }
            catch (Exception ex)
            {
                LogHelper.WriteExceptionLog("读取Excel文件异常！", "读取Excel文件异常：" + ex);
                return new Result<ExcelData>(false, null, ex.Message);
            }
        }
    }

    [Description("Excel单元格对象")]
    public class ExcelCell
    {
        [Description("单元格的值")]
        public string Value { get; set; }

        [Description("单元格标记")]
        public object Tag { get; set; }

        [Description("单元格所在行号（从1开始）")]
        public int RowIndex { get; set; } = 1;

        [Description("单元格所在列号（从1开始）")]
        public int ColumnIndex { get; set; } = 1;

        [Description("单元格背景颜色（默认白色）")]
        public Color BackColor { get; set; } = Color.White;

        [Description("单元格文字颜色（默认黑色）")]
        public Color FontColor { get; set; } = Color.Black;

        [Description("单元格文字大小（默认11）")]
        public double FontSize { get; set; } = 11;

        [Description("批注内容")]
        public string CommentText { get; set; }

        [Description("批注宽度")]
        public int CommentWidth { get; set; } = 0;

        [Description("批注高度")]
        public int CommentHeigth { get; set; } = 0;
    }

    [Description("Excel行对象")]
    public class ExcelRow
    {
        [Description("行号（从1开始）")]
        public int RowIndex { get; set; } = 1;

        [Description("行高（默认20）")]
        public double RowHeight { get; set; } = 20;

        [Description("单元格")]
        public List<ExcelCell> Cells { get; private set; } = new List<ExcelCell>();
    }

    [Description("Excel列对象")]
    public class ExcelColumn
    {
        private string _columnText = null;

        [Description("当前列标识")]
        public string ColumnName { get; set; }

        [Description("当前列名称")]
        public string ColumnText
        { get { return _columnText ?? ColumnName; } set { _columnText = value; } }

        [Description("列号（从1开始）")]
        public int ColumnIndex { get; set; } = 1;

        [Description("列高（默认20）")]
        public double ColumnWidth { get; set; } = 10;
    }

    [Description("Excel Sheet页对象")]
    public class ExcelSheet
    {
        [Description("当前Sheet页名称")]
        public string SheetName { get; set; }

        [Description("当前Sheet页索引号")]
        public int SheetIndex { get; set; }

        [Description("当前Sheet页行数")]
        public int RowCount
        { get { return Rows.Count; } }

        [Description("当前Sheet页列数")]
        public int ColumnCount
        { get { return Columns.Count; } }

        [Description("当前Sheet页列数")]
        public int ColumnHeight { get; set; } = 80;

        [Description("指示当前Sheet是否隐藏")]
        public bool IsHide { get; internal set; }

        [Description("当前Sheet页中所有行")]
        public List<ExcelRow> Rows { get; private set; } = new List<ExcelRow>();

        [Description("当前Sheet页中所有行")]
        public List<ExcelColumn> Columns { get; private set; } = new List<ExcelColumn>();

        [Description("当前Sheet页中所有行的合并单元格")]
        public List<ExcelRectangle> MergeRectangles { get; internal set; } = new List<ExcelRectangle>();

        [Description("合并表头的数据(Key:合并后的标题；Value1:起始列号,从1开始；Value2:结束列号,从1开始；)")]
        internal List<Tuple<string, int, int>> SpanHeaders { get; private set; }

        [Description("合并列头的数据(title:合并后的标题；startIndex:起始列号,从1开始；count:被合并的列数；)")]
        public bool AddSpanHeader(string title, int startIndex, int count)
        {
            if (SpanHeaders.Contains(t => t.Item1 == title) || startIndex < 1 || count < 1 || (startIndex + count - 1) > ColumnCount)
                return false;
            SpanHeaders.Add(Tuple.Create(title, startIndex, startIndex + count - 1));
            return true;
        }

        [Description("获取表头的行数")]
        public int GetHeadCount()
        {
            return Columns.Any() == false ? 0 : SpanHeaders.Any() ? 2 : 1;
        }
    }

    [Description("Excel文件对象")]
    public class ExcelData
    {
        [Description("当前Excel文件名称")]
        public string FileName { get; set; }

        [Description("当前Excel文件中所有Sheet页")]
        public List<ExcelSheet> Sheets { get; private set; } = new List<ExcelSheet>();
    }

    #region EXCEL模板参数

    [Description("EXCEL导出模板对象")]
    internal class ExcelTemplate
    {
        /// <summary>
        /// 模板名称
        /// </summary>
        public string TemplateName { get; set; } = "";

        /// <summary>
        /// 模板文件名称（包括扩展名）
        /// </summary>
        public string FileName { get; set; } = "";

        /// <summary>
        /// 导出名称
        /// </summary>
        public string ExportName { get; set; } = "";

        /// <summary>
        /// 模板区域结束行索引
        /// </summary>
        public int EndRow { get; set; } = 1;

        /// <summary>
        /// 模板区域结束列索引
        /// </summary>
        public int EndColumn { get; set; } = 1;

        /// <summary>
        /// 表体（也即行列表数据）起始行索引
        /// </summary>
        public int BodyStartRow { get; set; } = 2;

        /// <summary>
        /// 表体（也即行列表数据）总行数。0表示不需要分页，数据源全部写在同一页。大于0表示数据源多于行数后自动分页
        /// </summary>
        public int BodyRowCount { get; set; } = 0;

        /// <summary>
        /// 指示是否需要显示页码（BodyRowCount大于0时，会根据数据源数量分页）
        /// </summary>
        public bool ShowPage { get; set; } = false;

        /// <summary>
        /// 显示页码时，第一页页码所在的行号
        /// </summary>
        public int PageNumbRow { get; set; } = 0;

        /// <summary>
        /// 显示页码时，第一页页码所在的列号
        /// </summary>
        public int PageNumbColumn { get; set; } = 0;

        /// <summary>
        /// Sheet页
        /// </summary>
        public int SheetIndex { get; set; } = 1;

        internal Dictionary<string, string> ToDictionary()
        {
            var dic = new Dictionary<string, string>();
            dic.Add(nameof(TemplateName), TemplateName);
            dic.Add(nameof(FileName), FileName);
            dic.Add(nameof(ExportName), ExportName);
            dic.Add(nameof(EndRow), EndRow.ToString());
            dic.Add(nameof(EndColumn), EndColumn.ToString());
            dic.Add(nameof(BodyRowCount), BodyRowCount.ToString());
            dic.Add(nameof(BodyStartRow), BodyStartRow.ToString());
            dic.Add(nameof(ShowPage), ShowPage.ToString());
            dic.Add(nameof(PageNumbRow), PageNumbRow.ToString());
            dic.Add(nameof(PageNumbColumn), PageNumbColumn.ToString());
            dic.Add(nameof(SheetIndex), SheetIndex.ToString());
            return dic;
        }

        internal void SetValue(Dictionary<string, string> value)
        {
            if (value == null)
                return;
            if (value.ContainsKey(nameof(TemplateName)))
                TemplateName = value[nameof(TemplateName)];
            if (value.ContainsKey(nameof(FileName)))
                FileName = value[nameof(FileName)];
            if (value.ContainsKey(nameof(ExportName)))
                ExportName = value[nameof(ExportName)];
            if (value.ContainsKey(nameof(EndRow)))
                EndRow = value[nameof(EndRow)].ToInt();
            if (value.ContainsKey(nameof(EndColumn)))
                EndColumn = value[nameof(EndColumn)].ToInt();
            if (value.ContainsKey(nameof(BodyStartRow)))
                BodyStartRow = value[nameof(BodyStartRow)].ToInt();
            if (value.ContainsKey(nameof(BodyRowCount)))
                BodyRowCount = value[nameof(BodyRowCount)].ToInt();
            if (value.ContainsKey(nameof(ShowPage)))
                ShowPage = value[nameof(ShowPage)].ToBool();
            if (value.ContainsKey(nameof(PageNumbRow)))
                PageNumbRow = value[nameof(PageNumbRow)].ToInt();
            if (value.ContainsKey(nameof(PageNumbColumn)))
                PageNumbColumn = value[nameof(PageNumbColumn)].ToInt();
            if (value.ContainsKey(nameof(SheetIndex)))
                SheetIndex = value[nameof(SheetIndex)].ToInt();
        }

        /// <summary>
        /// 单元格的数据映射配置
        /// </summary>
        public List<ExcelDataMap> DataMap { get; set; } = new List<ExcelDataMap>();

        /// <summary>
        /// 合并单元格的列（表示有哪些列需要合并第一个）
        /// </summary>
        public List<ColumnMergeModel> MergeColumn { get; set; } = new List<ColumnMergeModel>();
    }

    [Description("合并列对象，表示哪几个列（ValueColumns）的数据相同时合并哪几个列（MergeColumns）")]
    internal class ColumnMergeModel
    {
        /// <summary>
        /// 要判断值得列
        /// </summary>
        public List<int> ValueColumns { get; set; } = new List<int>();

        /// <summary>
        /// 要合并的列
        /// </summary>
        public List<int> MergeColumns { get; set; } = new List<int>();
    }

    [Description("模板数据映射配置")]
    internal class ExcelDataMap
    {
        [Description("表头对应的数据源名称")]
        public string ValueName { get; set; } = "";

        [Description("指示是否为表头配置")]
        public bool IsHead { get; set; } = false;

        [Description("数据对应行号")]
        public int RowIndex { get; set; } = 1;

        [Description("数据对应的列号")]
        public int ColumnIndex { get; set; } = 1;

        [Description("将当前对象的属性转换到字典中")]
        internal Dictionary<string, string> ToDictionary()
        {
            var dic = new Dictionary<string, string>();
            dic.Add(nameof(ValueName), ValueName);
            dic.Add(nameof(IsHead), IsHead.ToString());
            dic.Add(nameof(RowIndex), RowIndex.ToString());
            dic.Add(nameof(ColumnIndex), ColumnIndex.ToString());
            return dic;
        }

        internal void SetValue(Dictionary<string, string> value)
        {
            if (value == null)
                return;
            if (value.ContainsKey(nameof(ValueName)))
                ValueName = value[nameof(ValueName)];
            if (value.ContainsKey(nameof(IsHead)))
                IsHead = value[nameof(IsHead)].ToBool();
            if (value.ContainsKey(nameof(RowIndex)))
                RowIndex = value[nameof(RowIndex)].ToInt();
            if (value.ContainsKey(nameof(ColumnIndex)))
                ColumnIndex = value[nameof(ColumnIndex)].ToInt();
        }
    }

    [Description("Excel区域对象")]
    public class ExcelRectangle
    {
        public ExcelRectangle(int startRow, int startColumn, int endRow, int endColumn)
        {
            StartRow = startRow;
            EndRow = endRow;
            StartColumn = startColumn;
            EndColumn = endColumn;
        }

        [Description("起始行号（从1开始）")]
        public int StartRow { get; set; } = 1;

        [Description("结束行号（从1开始）")]
        public int EndRow { get; set; } = 1;

        [Description("起始列号（从1开始）")]
        public int StartColumn { get; set; } = 1;

        [Description("结束列号（从1开始）")]
        public int EndColumn { get; set; } = 1;

        [Description("指示当前区域值所表示的区域是否为一个正常的区域（如起始行是否小于等于结束行，起始列是否小于等于结束列）")]
        public bool NormalRange
        { get { return StartColumn > 0 && StartRow > 0 && StartRow <= EndRow && StartColumn <= EndColumn; } }

        [Description("复制生成一个新的区域对象")]
        public ExcelRectangle Copy()
        {
            return new ExcelRectangle(StartRow, StartColumn, EndRow, EndColumn);
        }

        [Description("获取当前区域和另一个区域的交集区域（返回null表示无交集区域）")]
        public ExcelRectangle GetJoinRange(ExcelRectangle range)
        {
            if (range == null)
                return null;
            var joinrange = new ExcelRectangle(Math.Max(StartRow, range.StartRow), Math.Max(StartColumn, range.StartColumn), Math.Min(EndRow, range.EndRow), Math.Min(EndColumn, range.EndColumn));
            return joinrange.NormalRange ? joinrange : null;
        }

        [Description("判断当前区域是否包含某个区域的全部")]
        public bool ContainAll(ExcelRectangle range)
        {
            if (range == null || !NormalRange || !range.NormalRange)
                return false;
            return range.StartColumn >= StartColumn && range.EndColumn <= EndColumn
                && range.StartRow >= StartRow && range.EndRow <= EndRow;
        }

        [Description("判断当前区域是否包含某区域的部分")]
        public bool ContainPart(ExcelRectangle range)
        {
            return GetJoinRange(range) != null;
        }

        [Description("判断当前区域是否包含某个单元格")]
        public bool Contain(int row, int column)
        {
            if (row < 1 || column < 1)
                return false;
            return row >= StartRow && row <= EndRow && column >= StartColumn && column <= EndColumn;
        }

        [Description("标记")]
        public object Tag { get; set; }
    }

    #endregion EXCEL模板参数
}