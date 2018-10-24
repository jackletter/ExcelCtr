using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
using System.Web;
using NPOI;
using NPOI.HPSF;
using NPOI.HSSF;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.POIFS;
using NPOI.SS.Formula.Eval;
using NPOI.SS.UserModel;
using NPOI.Util;
using NPOI.SS;
using NPOI.DDF;
using NPOI.SS.Util;
using System.Collections;
using System.Text.RegularExpressions;
using NPOI.XSSF;
using NPOI.XSSF.UserModel;

namespace ExcelCtr
{
    internal class ExcelHelper
    {
        #region 导出excel

        /// <summary>DataSet导出到Excel的MemoryStream</summary>
        /// <param name="ds">源DataSet</param>
        /// <param name="strHeaderTexts">表格头文本值集合</param>
        /// <param name="sheetCombineColIndexs">每个表格的要垂直合并的列的序号如：{"0,1","2"}表示表1的第0和1列进行合并,表2的第2列进行合并</param>
        /// <returns></returns>
        public static MemoryStream ExportDS(DataSet ds, List<string> strHeaderTexts, List<string> sheetCombineColIndexs)
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            HSSFSheet sheet = null;
            //自动换行单元格样式
            ICellStyle cellStyle_crlf = workbook.CreateCellStyle();
            cellStyle_crlf.WrapText = true;

            #region 右击文件 属性信息

            {
                DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
                dsi.Company = "http://www.jack.com/";
                workbook.DocumentSummaryInformation = dsi;

                SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
                si.Author = "Jack"; //填加xls文件作者信息
                si.ApplicationName = "jackExcel"; //填加xls文件创建程序信息
                si.LastAuthor = "Jack"; //填加xls文件最后保存者信息
                si.Comments = "Jack导出的excel"; //填加xls文件作者信息
                si.Title = "Jack导出的excel"; //填加xls文件标题信息
                si.Subject = "Jack导出的excel"; //填加文件主题信息
                si.CreateDateTime = DateTime.Now;
                workbook.SummaryInformation = si;
            }

            #endregion

            HSSFCellStyle dateStyle = workbook.CreateCellStyle() as HSSFCellStyle;
            HSSFDataFormat format = workbook.CreateDataFormat() as HSSFDataFormat;
            dateStyle.DataFormat = format.GetFormat("yyyy-MM-dd HH:mm:ss.fff");
            for (int i = 0; i < ds.Tables.Count; i++)
            {
                #region 填充单个sheet

                int contentRowStartIndex = 0;//表格内容的内容数据起始行,用于合并同列多行之间的合并
                sheet = workbook.CreateSheet(ds.Tables[i].TableName.StartsWith("Table") ? "Sheet" + (i + 1).ToString() : ds.Tables[i].TableName) as HSSFSheet;
                sheet.DefaultRowHeight = 22 * 20;
                DataTable dtSource = ds.Tables[i];
                //取得列宽
                int[] arrColWidth = new int[dtSource.Columns.Count];//保存列的宽度
                foreach (DataColumn item in dtSource.Columns)
                {
                    //先根据列名的字符串长度初始化所有的列宽
                    arrColWidth[item.Ordinal] = Encoding.GetEncoding(936).GetBytes(item.ColumnName.ToString()).Length;
                }
                for (int ii = 0; ii < dtSource.Rows.Count; ii++)
                {
                    //遍历数据内容,根据每一列的数据最大长度设置列宽
                    for (int j = 0; j < dtSource.Columns.Count; j++)
                    {
                        int intTemp = Encoding.GetEncoding(936).GetBytes(dtSource.Rows[ii][j].ToString()).Length;
                        if (intTemp > arrColWidth[j])
                        {
                            arrColWidth[j] = intTemp;
                        }
                    }
                }

                #region 填充表头,列头,数据内容
                int rowIndex = 0;
                foreach (DataRow row in dtSource.Rows)
                {
                    #region 新建表，填充表头，填充列头，样式
                    if (rowIndex == 0)
                    {

                        #region 表头及样式

                        if (strHeaderTexts != null && strHeaderTexts.Count - 1 >= i)
                        {
                            if (!string.IsNullOrWhiteSpace(strHeaderTexts[i]))
                            {
                                HSSFRow headerRow = sheet.CreateRow(0) as HSSFRow;
                                headerRow.HeightInPoints = 25;
                                ICell cell = headerRow.CreateCell(0);
                                ICellStyle cellStyle = workbook.CreateCellStyle();
                                cellStyle.WrapText = true;
                                cell.CellStyle = cellStyle;
                                cell.SetCellValue(new HSSFRichTextString(strHeaderTexts[i]));

                                HSSFCellStyle headStyle = workbook.CreateCellStyle() as HSSFCellStyle;
                                headStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                                HSSFFont font = workbook.CreateFont() as HSSFFont;
                                font.FontHeightInPoints = 15;
                                font.Boldweight = 700;
                                headStyle.SetFont(font);

                                headerRow.GetCell(0).CellStyle = headStyle;
                                sheet.AddMergedRegion(new Region(0, 0, 0, dtSource.Columns.Count - 1));
                                rowIndex++;
                                //headerRow.Dispose();
                            }
                        }

                        #endregion

                        #region 列头及样式

                        {
                            HSSFRow headerRow = sheet.CreateRow(rowIndex) as HSSFRow;
                            HSSFCellStyle headStyle = workbook.CreateCellStyle() as HSSFCellStyle;
                            headStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                            headStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
                            headStyle.WrapText = true;
                            headerRow.HeightInPoints = 23;
                            HSSFFont font = workbook.CreateFont() as HSSFFont;
                            font.FontHeightInPoints = 11;
                            font.Boldweight = 700;
                            headStyle.SetFont(font);


                            foreach (DataColumn column in dtSource.Columns)
                            {
                                ICell cell = headerRow.CreateCell(column.Ordinal);
                                cell.CellStyle = headStyle;
                                cell.SetCellValue(new HSSFRichTextString(column.ColumnName));

                                //设置列宽，这里我多加了10个字符的长度
                                sheet.SetColumnWidth(column.Ordinal, (arrColWidth[column.Ordinal] + 10) * 256);
                            }

                            rowIndex++;
                            contentRowStartIndex = rowIndex;//记住数据内容的起始行
                            //headerRow.Dispose();
                        }

                        #endregion

                    }

                    #endregion

                    #region 填充内容
                    HSSFRow dataRow = sheet.CreateRow(rowIndex) as HSSFRow;
                    dataRow.HeightInPoints = 22;
                    foreach (DataColumn column in dtSource.Columns)
                    {
                        HSSFCell newCell = dataRow.CreateCell(column.Ordinal) as HSSFCell;

                        string drValue = row[column].ToString();

                        switch (column.DataType.ToString())
                        {
                            case "System.String": //字符串类型
                                newCell.CellStyle = cellStyle_crlf;
                                newCell.SetCellType(CellType.String);
                                newCell.SetCellValue(new HSSFRichTextString(drValue));
                                break;

                            case "System.DateTime": //日期类型
                                DateTime dateV;
                                DateTime.TryParse(drValue, out dateV);
                                newCell.SetCellValue(dateV);

                                newCell.CellStyle = dateStyle; //格式化显示
                                break;
                            case "System.Boolean": //布尔型
                                bool boolV = false;
                                bool.TryParse(drValue, out boolV);
                                newCell.SetCellValue(boolV);
                                break;
                            case "System.Int16": //整型
                            case "System.Int32":
                            case "System.Int64":
                            case "System.Byte":
                                int intV = 0;
                                int.TryParse(drValue, out intV);
                                newCell.SetCellValue(intV);
                                break;
                            case "System.Decimal": //浮点型
                            case "System.Double":
                                double doubV = 0;
                                double.TryParse(drValue, out doubV);
                                newCell.SetCellValue(doubV);
                                break;
                            case "System.DBNull": //空值处理
                                newCell.SetCellValue("");
                                break;
                            default:
                                newCell.SetCellValue("");
                                break;
                        }

                    }

                    #endregion

                    rowIndex++;
                }
                #endregion

                Hashtable ht = new Hashtable();

                #region 同列中多行之间的合并
                if (sheetCombineColIndexs != null && sheetCombineColIndexs.Count > i)
                {
                    List<int> combineColIndexs = new List<int>();
                    string[] strarr = sheetCombineColIndexs[i].Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                    foreach (var item in strarr)
                    {
                        combineColIndexs.Add(int.Parse(item));
                    }
                    for (int j = contentRowStartIndex; j < rowIndex; j++)
                    {
                        for (int jj = 0; jj < combineColIndexs.Count; jj++)
                        {
                            int cloIndex = combineColIndexs[jj];
                            if (j == contentRowStartIndex)
                            {
                                Entry entry = new Entry();
                                entry.startIndex = contentRowStartIndex;
                                object o = GetCellValue(sheet.GetRow(contentRowStartIndex).Cells[cloIndex]);
                                entry.combineValue = o == null ? "" : o.ToString();
                                ht.Add(cloIndex, entry);
                                continue;
                            }
                            object obj = GetCellValue(sheet.GetRow(j).Cells[cloIndex]);
                            string value = obj == null ? "" : obj.ToString();
                            Entry en = (Entry)ht[cloIndex];
                            if (en.combineValue != value)
                            {
                                //如果发生不相等的情况则满足合并条件(最少是2行)就会合并
                                if (en.startIndex + 1 < j)
                                {
                                    sheet.AddMergedRegion(new Region(en.startIndex, cloIndex, j - 1, cloIndex));
                                    ICell cell = sheet.GetRow(en.startIndex).Cells[cloIndex];
                                    ICellStyle cellstyle = workbook.CreateCellStyle();//设置垂直居中格式
                                    cellstyle.VerticalAlignment = VerticalAlignment.Center;//垂直居中
                                    cell.CellStyle = cellstyle;
                                }
                                en.combineValue = value;
                                en.startIndex = j;
                            }
                            else
                            {
                                //如果相等了,再判断是不是最后一行,如果是最后一行也要合并
                                if (j == rowIndex - 1)
                                {
                                    sheet.AddMergedRegion(new Region(en.startIndex, cloIndex, j, cloIndex));
                                    ICell cell = sheet.GetRow(en.startIndex).Cells[cloIndex];
                                    ICellStyle cellstyle = workbook.CreateCellStyle();//设置垂直居中格式
                                    cellstyle.VerticalAlignment = VerticalAlignment.Center;//垂直居中
                                    cell.CellStyle = cellstyle;
                                }
                            }
                        }
                    }
                }
                #endregion

                #endregion

            }
            using (MemoryStream ms = new MemoryStream())
            {
                workbook.Write(ms);
                ms.Flush();
                ms.Position = 0;
                return ms;
            }
        }

        #endregion

        #region 导入excel
        /// <summary>读取excel默认第一行为标头</summary>
        /// <param name="strFileName">excel文档路径</param>
        /// <returns></returns>
        public static DataSet ImportExceltoDs(string strFileName)
        {
            using (FileStream file = new FileStream(strFileName, FileMode.Open, FileAccess.Read))
            {
                return ImportExceltoDs(file);
            }
        }

        /// <summary>从指定excel流中读取excel成为dataset</summary>
        /// <param name="stream"></param>
        /// <returns></returns>
        public static DataSet ImportExceltoDs(Stream stream)
        {
            DataSet ds = new DataSet();
            DataTable dt = null;
            IWorkbook wb;
            wb = WorkbookFactory.Create(stream);
            for (int i = 0; i < wb.NumberOfSheets; i++)
            {
                ISheet sheet = wb.GetSheetAt(i);
                dt = ImportDt(sheet, 0);
                dt.TableName = sheet.SheetName;
                ds.Tables.Add(dt);
            }
            return ds;
        }

        /// <summary>读取excel中指定表名和指定相应列头行的表</summary>
        /// <param name="strFileName"></param>
        /// <param name="sheetNames"></param>
        /// <param name="indexOfColNames"></param>
        /// <returns></returns>
        public static DataSet ImportExceltoDs(string strFileName, List<string> sheetNames, List<int> indexOfColNames)
        {
            using (FileStream file = new FileStream(strFileName, FileMode.Open, FileAccess.Read))
            {
                return ImportExceltoDs(file, sheetNames, indexOfColNames);
            }
        }

        /// <summary>从指定流中读取excel中指定表名和指定相应列头行的表</summary>
        /// <param name="stream"></param>
        /// <param name="sheetNames"></param>
        /// <param name="indexOfColNames"></param>
        /// <returns></returns>
        public static DataSet ImportExceltoDs(Stream stream, List<string> sheetNames, List<int> indexOfColNames)
        {
            DataSet ds = new DataSet();
            DataTable dt = null;
            IWorkbook wb;
            wb = WorkbookFactory.Create(stream);
            for (int i = 0; i < sheetNames.Count; i++)
            {
                ISheet sheet = wb.GetSheet(sheetNames[i]);
                if (sheet != null)
                {
                    dt = ImportDt(sheet, indexOfColNames.Count > i ? indexOfColNames[i] : 0);
                    dt.TableName = sheet.SheetName;
                    ds.Tables.Add(dt);
                }
            }
            return ds;
        }

        /// <summary>读取excel中指定表索引和相应列头行的表</summary>
        /// <param name="strFileName"></param>
        /// <param name="sheetIndexs"></param>
        /// <param name="indexOfColNames"></param>
        /// <returns></returns>
        public static DataSet ImportExceltoDs(string strFileName, List<int> sheetIndexs, List<int> indexOfColNames)
        {
            using (FileStream file = new FileStream(strFileName, FileMode.Open, FileAccess.Read))
            {
                return ImportExceltoDs(file, sheetIndexs, indexOfColNames);
            }
        }

        /// <summary>从指定流中读取excel中指定表索引和相应列头行的表</summary>
        /// <param name="stream"></param>
        /// <param name="sheetIndexs"></param>
        /// <param name="indexOfColNames"></param>
        /// <returns></returns>
        public static DataSet ImportExceltoDs(Stream stream, List<int> sheetIndexs, List<int> indexOfColNames)
        {
            DataSet ds = new DataSet();
            DataTable dt = null;
            IWorkbook wb;
            wb = WorkbookFactory.Create(stream);
            for (int i = 0; i < sheetIndexs.Count; i++)
            {
                ISheet sheet = wb.GetSheetAt(sheetIndexs[i]);
                if (sheet != null)
                {
                    dt = ImportDt(sheet, indexOfColNames.Count > i ? indexOfColNames[i] : 0);
                    dt.TableName = sheet.SheetName;
                    ds.Tables.Add(dt);
                }
            }
            return ds;
        }

        /// <summary>从指定流中读取excel中指定表是否有列头行,和列投行的位置或者数据内容的起始行索引和列索引</summary>
        /// <param name="stream">文件流</param>
        /// <param name="sheetIndexs">sheet索引集合</param>
        /// <param name="hasColNames">每个sheet是否列头行的布尔说明</param>
        /// <param name="dataStartIndex">每个sheet的列头行号或者是数据内容起始的起始行索引或者是起始列索引</param>
        /// <returns></returns>
        public static DataSet ImportExceltoDs(Stream stream, List<int> sheetIndexs, List<bool> hasColNames, List<int[]> dataStartIndex)
        {
            DataSet ds = new DataSet();
            DataTable dt = null;
            IWorkbook wb;
            wb = WorkbookFactory.Create(stream);
            for (int i = 0; i < sheetIndexs.Count; i++)
            {
                ISheet sheet = wb.GetSheetAt(sheetIndexs[i]);
                if (sheet != null)
                {
                    if (hasColNames[i])
                    {
                        dt = ImportDt(sheet, dataStartIndex[i][0]);
                    }
                    else
                    {
                        dt = ImportDt(sheet, dataStartIndex[i][0], dataStartIndex[i][1]);
                    }
                    dt.TableName = sheet.SheetName;
                    ds.Tables.Add(dt);
                }
            }
            return ds;
        }

        /// <summary>从指定流中读取excel中指定表是否有列头行,和列投行的位置或者数据内容的起始行索引和列索引</summary>
        /// <param name="filePath">文件路径</param>
        /// <param name="sheetIndexs">sheet索引集合</param>
        /// <param name="hasColNames">每个sheet是否列头行的布尔说明</param>
        /// <param name="dataStartIndex">每个sheet的列头行号或者是数据内容起始的起始行索引或者是起始列索引</param>
        /// <returns></returns>
        public static DataSet ImportExceltoDs(string strFileName, List<int> sheetIndexs, List<bool> hasColNames, List<int[]> dataStartIndex)
        {
            using (FileStream file = new FileStream(strFileName, FileMode.Open, FileAccess.Read))
            {
                return ImportExceltoDs(file, sheetIndexs, hasColNames, dataStartIndex);
            }
        }

        /// <summary>从指定流中读取excel中指定表是否有列头行,和列投行的位置或者数据内容的起始行索引和列索引</summary>
        /// <param name="stream">文件流</param>
        /// <param name="sheetNames">sheet名字集合</param>
        /// <param name="hasColNames">每个sheet是否列头行的布尔说明</param>
        /// <param name="dataStartIndex">每个sheet的列头行号或者是数据内容起始的起始行索引或者是起始列索引</param>
        /// <returns></returns>
        public static DataSet ImportExceltoDs(Stream stream, List<string> sheetNames, List<bool> hasColNames, List<int[]> dataStartIndex)
        {
            DataSet ds = new DataSet();
            DataTable dt = null;
            IWorkbook wb;
            wb = WorkbookFactory.Create(stream);
            for (int i = 0; i < sheetNames.Count; i++)
            {
                ISheet sheet = wb.GetSheet(sheetNames[i]);
                if (sheet != null)
                {
                    if (hasColNames[i])
                    {
                        dt = ImportDt(sheet, dataStartIndex[i][0]);
                    }
                    else
                    {
                        dt = ImportDt(sheet, dataStartIndex[i][0], dataStartIndex[i][1]);
                    }
                    dt.TableName = sheet.SheetName;
                    ds.Tables.Add(dt);
                }
            }
            return ds;
        }

        /// <summary>从指定流中读取excel中指定表是否有列头行,和列投行的位置或者数据内容的起始行索引和列索引</summary>
        /// <param name="filePath">文件路径</param>
        /// <param name="sheetNames">sheet名字集合</param>
        /// <param name="hasColNames">每个sheet是否列头行的布尔说明</param>
        /// <param name="dataStartIndex">每个sheet的列头行号或者是数据内容起始的起始行索引或者是起始列索引</param>
        /// <returns></returns>
        public static DataSet ImportExceltoDs(string strFileName, List<string> sheetNames, List<bool> hasColNames, List<int[]> dataStartIndex)
        {
            using (FileStream file = new FileStream(strFileName, FileMode.Open, FileAccess.Read))
            {
                return ImportExceltoDs(file, sheetNames, hasColNames, dataStartIndex);
            }
        }

        /// <summary>将指定sheet中的数据读取到datatable中</summary>
        /// <param name="sheet">需要读入的sheet</param>
        /// <param name="HeaderRowIndex">列头所在行号(小于0则第一行视为列头行)</param>
        /// <returns></returns>
        public static DataTable ImportDt(ISheet sheet, int HeaderRowIndex = 0)
        {
            DataTable table = new DataTable();
            IRow headerRow;
            int cellCount;
            try
            {
                #region 生成列名
                //其他行为首行的时候
                headerRow = sheet.GetRow(HeaderRowIndex);
                cellCount = headerRow.LastCellNum;

                for (int i = headerRow.FirstCellNum; i < cellCount; i++)
                {
                    //遍历这一行的单元格的起始索引和结束索引
                    if (headerRow.GetCell(i) == null)
                    {
                        //如果单元格的内容为空,首选(i,i为列索引)生成列名后插入,重复的时候再以(重复列名i,i为列索引)生成列名后插入
                        if (table.Columns.IndexOf(Convert.ToString(i)) > 0)
                        {
                            DataColumn column = new DataColumn(Convert.ToString("重复列名" + i));
                            table.Columns.Add(column);
                        }
                        else
                        {
                            DataColumn column = new DataColumn(Convert.ToString(i));
                            table.Columns.Add(column);
                        }

                    }
                    else if (table.Columns.IndexOf(headerRow.GetCell(i).ToString()) > 0)
                    {
                        //如果出现重复列名,就以(重复列名i,i为索引)格式生成列名后插入
                        DataColumn column = new DataColumn(Convert.ToString("重复列名" + i));
                        table.Columns.Add(column);
                    }
                    else
                    {
                        //一般情况下的直接生成列名后插入到table
                        DataColumn column = new DataColumn(headerRow.GetCell(i).ToString());
                        table.Columns.Add(column);
                    }
                }
                #endregion

                #region 生成数据
                int rowCount = sheet.LastRowNum;
                for (int i = (HeaderRowIndex + 1); i <= sheet.LastRowNum; i++)
                {
                    try
                    {
                        IRow row;
                        if (sheet.GetRow(i) == null)
                        {
                            row = sheet.CreateRow(i);
                        }
                        else
                        {
                            row = sheet.GetRow(i);
                        }

                        DataRow dataRow = table.NewRow();

                        for (int j = row.FirstCellNum; j <= cellCount; j++)
                        {
                            try
                            {
                                if (row.GetCell(j) != null)
                                {
                                    switch (row.GetCell(j).CellType)
                                    {
                                        case CellType.String:
                                            string str = row.GetCell(j).StringCellValue;
                                            if (str != null && str.Length > 0)
                                            {
                                                dataRow[j] = str.ToString();
                                            }
                                            else
                                            {
                                                dataRow[j] = null;
                                            }
                                            break;
                                        case CellType.Numeric:
                                            if (DateUtil.IsCellDateFormatted(row.GetCell(j)))
                                            {
                                                dataRow[j] = DateTime.FromOADate(row.GetCell(j).NumericCellValue).ToString("yyyy-MM-dd HH:mm:ss.fff");
                                            }
                                            else
                                            {
                                                dataRow[j] = Convert.ToDouble(row.GetCell(j).NumericCellValue);
                                            }
                                            break;
                                        case CellType.Boolean:
                                            dataRow[j] = Convert.ToString(row.GetCell(j).BooleanCellValue);
                                            break;
                                        case CellType.Error:
                                            dataRow[j] = ErrorEval.GetText(row.GetCell(j).ErrorCellValue);
                                            break;
                                        case CellType.Formula:
                                            switch (row.GetCell(j).CachedFormulaResultType)
                                            {
                                                case CellType.String:
                                                    string strFORMULA = row.GetCell(j).StringCellValue;
                                                    if (strFORMULA != null && strFORMULA.Length > 0)
                                                    {
                                                        dataRow[j] = strFORMULA.ToString();
                                                    }
                                                    else
                                                    {
                                                        dataRow[j] = null;
                                                    }
                                                    break;
                                                case CellType.Numeric:
                                                    dataRow[j] = Convert.ToString(row.GetCell(j).NumericCellValue);
                                                    break;
                                                case CellType.Boolean:
                                                    dataRow[j] = Convert.ToString(row.GetCell(j).BooleanCellValue);
                                                    break;
                                                case CellType.Error:
                                                    dataRow[j] = ErrorEval.GetText(row.GetCell(j).ErrorCellValue);
                                                    break;
                                                default:
                                                    dataRow[j] = "";
                                                    break;
                                            }
                                            break;
                                        default:
                                            dataRow[j] = "";
                                            break;
                                    }
                                }
                            }
                            catch (Exception exception)
                            {
                                //wl.WriteLogs(exception.ToString());
                                throw exception;
                            }
                        }
                        table.Rows.Add(dataRow);
                    }
                    catch (Exception exception)
                    {
                        throw exception;
                    }
                }
                #endregion
            }
            catch (Exception exception)
            {
                //wl.WriteLogs(exception.ToString());
            }
            return table;
        }

        /// <summary>将指定sheet中的数据读取到datatable中</summary>
        /// <param name="sheet">需要读入的sheet</param>
        /// <param name="dataStartRowIndex">数据内容的起始行索引</param>
        /// <param name="dataStartColIndex">数据内容的起始列索引</param>
        /// <returns></returns>
        public static DataTable ImportDt(ISheet sheet, int dataStartRowIndex, int dataStartColIndex)
        {
            DataTable table = new DataTable();
            try
            {
                #region 生成数据
                int rowCount = sheet.LastRowNum;
                for (int i = dataStartRowIndex; i <= sheet.LastRowNum; i++)
                {
                    try
                    {
                        IRow row;

                        if (sheet.GetRow(i) == null)
                        {
                            row = sheet.CreateRow(i);
                        }
                        else
                        {
                            row = sheet.GetRow(i);
                        }
                        int cellCount = row.LastCellNum;
                        DataRow dataRow = table.NewRow();
                        //处理没有响应列号的情况,因为没有列名,前面也就没有添加列
                        if (table.Columns.Count < cellCount)
                        {
                            int tmp = table.Columns.Count;
                            for (int t = 0; t < cellCount - tmp; t++)
                            {
                                table.Columns.Add(new DataColumn());
                            }
                        }

                        for (int j = row.FirstCellNum; j <= cellCount; j++)
                        {
                            try
                            {
                                if (row.GetCell(j) != null)
                                {
                                    switch (row.GetCell(j).CellType)
                                    {
                                        case CellType.String:
                                            string str = row.GetCell(j).StringCellValue;
                                            if (str != null && str.Length > 0)
                                            {
                                                dataRow[j] = str.ToString();
                                            }
                                            else
                                            {
                                                dataRow[j] = null;
                                            }
                                            break;
                                        case CellType.Numeric:
                                            if (DateUtil.IsCellDateFormatted(row.GetCell(j)))
                                            {
                                                dataRow[j] = DateTime.FromOADate(row.GetCell(j).NumericCellValue).ToString("yyyy-MM-dd HH:mm:ss.fff");
                                            }
                                            else
                                            {
                                                dataRow[j] = Convert.ToDouble(row.GetCell(j).NumericCellValue);
                                            }
                                            break;
                                        case CellType.Boolean:
                                            dataRow[j] = Convert.ToString(row.GetCell(j).BooleanCellValue);
                                            break;
                                        case CellType.Error:
                                            dataRow[j] = ErrorEval.GetText(row.GetCell(j).ErrorCellValue);
                                            break;
                                        case CellType.Formula:
                                            switch (row.GetCell(j).CachedFormulaResultType)
                                            {
                                                case CellType.String:
                                                    string strFORMULA = row.GetCell(j).StringCellValue;
                                                    if (strFORMULA != null && strFORMULA.Length > 0)
                                                    {
                                                        dataRow[j] = strFORMULA.ToString();
                                                    }
                                                    else
                                                    {
                                                        dataRow[j] = null;
                                                    }
                                                    break;
                                                case CellType.Numeric:
                                                    dataRow[j] = Convert.ToString(row.GetCell(j).NumericCellValue);
                                                    break;
                                                case CellType.Boolean:
                                                    dataRow[j] = Convert.ToString(row.GetCell(j).BooleanCellValue);
                                                    break;
                                                case CellType.Error:
                                                    dataRow[j] = ErrorEval.GetText(row.GetCell(j).ErrorCellValue);
                                                    break;
                                                default:
                                                    dataRow[j] = "";
                                                    break;
                                            }
                                            break;
                                        default:
                                            dataRow[j] = "";
                                            break;
                                    }
                                }
                            }
                            catch (Exception exception)
                            {
                                //wl.WriteLogs(exception.ToString());
                                throw exception;
                            }
                        }
                        table.Rows.Add(dataRow);
                    }
                    catch (Exception exception)
                    {
                        throw exception;
                    }
                }
                #endregion
            }
            catch (Exception exception)
            {
                //wl.WriteLogs(exception.ToString());
            }
            return table;
        }

        #endregion

        /// <summary>判断单元格的值是不是数字
        /// </summary>
        /// <param name="value">要进行判断的值</param>
        /// <param name="result">转换成的数字</param>
        /// <returns></returns>
        public static bool isNumeric(String value, out double result)
        {
            Regex rex = new Regex(@"^[-]?\d+[.]?\d*$");
            result = -1;
            if (rex.IsMatch(value))
            {
                result = double.Parse(value);
                return true;
            }
            else
                return false;

        }

        /// <summary>获取单元格的值
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        public static object GetCellValue(ICell cell)
        {
            switch (cell.CellType)
            {
                case CellType.String:
                    string str = cell.StringCellValue;
                    if (str != null && str.Length > 0)
                    {
                        return str.ToString();
                    }
                    else
                    {
                        return null;
                    }
                case CellType.Numeric:
                    if (DateUtil.IsCellDateFormatted(cell))
                    {
                        return DateTime.FromOADate(cell.NumericCellValue).ToString("yyyy-MM-dd HH:mm:ss.fff");
                    }
                    else
                    {
                        return Convert.ToDouble(cell.NumericCellValue);
                    }
                case CellType.Boolean:
                    return Convert.ToString(cell.BooleanCellValue);
                case CellType.Error:
                    return ErrorEval.GetText(cell.ErrorCellValue);
                case CellType.Formula:
                    switch (cell.CachedFormulaResultType)
                    {
                        case CellType.String:
                            string strFORMULA = cell.StringCellValue;
                            if (strFORMULA != null && strFORMULA.Length > 0)
                            {
                                return strFORMULA.ToString();
                            }
                            else
                            {
                                return null;
                            }
                        case CellType.Numeric:
                            return Convert.ToString(cell.NumericCellValue);
                        case CellType.Boolean:
                            return Convert.ToString(cell.BooleanCellValue);
                        case CellType.Error:
                            return ErrorEval.GetText(cell.ErrorCellValue);
                        default:
                            return "";
                    }
                default:
                    return "";
            }
        }

        /// <summary>插入行</summary>
        /// <param name="sheet">要插入行的sheet</param>
        /// <param name="startindex">从这一行的前面插入(这一行开始包括这一行都会被整体向下移动rowcount)</param>
        /// <param name="rowcount">插入的行数</param>
        /// <param name="stylerow">被插入行采用的样式行的索引,注意这个索引行所在的位置应该位于插入起始行之上</param>
        public static void InsertRow(ISheet sheet, int startindex, int rowcount, int styleindex)
        {
            IRow stylerow = sheet.GetRow(styleindex);
            if (sheet.LastRowNum >= startindex)
            {
                //批量移动行
                sheet.ShiftRows(startindex, sheet.LastRowNum, rowcount, true/*是否复制行高*/, false);
            }

            #region 对批量移动后空出的空行插，创建相应的行，并以样式行作为模板设置样式
            for (int i = startindex; i < startindex + rowcount; i++)
            {
                IRow targetRow = null;
                ICell sourceCell = null;
                ICell targetCell = null;

                targetRow = sheet.CreateRow(i);
                targetRow.Height = stylerow.Height;
                targetRow.HeightInPoints = stylerow.HeightInPoints;
                targetRow.ZeroHeight = stylerow.ZeroHeight;

                int mergeindex = -1;
                for (int m = stylerow.FirstCellNum; m < stylerow.LastCellNum; m++)
                {
                    sourceCell = stylerow.GetCell(m);
                    if (sourceCell == null)
                    {
                        if (mergeindex > 0)
                        {
                            sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(i, i, mergeindex, m));
                            mergeindex = -1;
                        }
                        continue;
                    }


                    targetCell = targetRow.CreateCell(m);

                    targetCell.CellStyle = sourceCell.CellStyle;
                    targetCell.SetCellType(sourceCell.CellType);
                    if (sourceCell.IsMergedCell)
                    {
                        if (mergeindex > 0 && m + 1 < stylerow.LastCellNum)
                        {
                            continue;
                        }
                        else if (mergeindex > 0 && m + 1 == stylerow.LastCellNum)
                        {
                            sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(i, i, mergeindex, m));
                            mergeindex = -1;
                        }
                        else
                        {
                            mergeindex = m;
                        }
                    }
                    else
                    {
                        if (mergeindex > 0)
                        {
                            sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(i, i, mergeindex, m));
                            mergeindex = -1;
                        }
                    }
                }
            }

            IRow firstTargetRow = sheet.GetRow(startindex);
            ICell firstSourceCell = null;
            ICell firstTargetCell = null;
            if (rowcount > 0)
            {
                //新添加的行应用样式
                for (int m = stylerow.FirstCellNum; m < stylerow.LastCellNum; m++)
                {
                    firstSourceCell = stylerow.GetCell(m);
                    if (firstSourceCell == null)
                        continue;
                    firstTargetCell = firstTargetRow.CreateCell(m);

                    firstTargetCell.CellStyle = firstSourceCell.CellStyle;
                    firstTargetCell.SetCellType(firstSourceCell.CellType);
                }
            }
            #endregion
        }
    }

    //用于合并垂直单元格时的实体类
    public class Entry
    {
        //记录下当前正在合并的值
        public string combineValue { set; get; }
        //记录下当前合并的起点索引
        public int startIndex { set; get; }
    }
}