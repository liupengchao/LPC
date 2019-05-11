using System;
using System.Data;
using System.IO;
using System.Text;
using System.Web;
using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using NPOI.SS.Formula.Eval;
using System.Collections.Generic;
using NPOI.HSSF.Record;
using System.Web.UI.DataVisualization.Charting;

/// <summary>
/// NPOIHelper 的摘要说明
/// </summary>
/// 

namespace H3.SHHS.Excel
{
    public class NPOIHelper
    {
        public NPOIHelper()
        {
            //
            // TODO: 在此处添加构造函数逻辑
            //
        }

        #region 导出

        /// <summary>
        ///     DataTable导出到Excel文件
        /// </summary>
        /// <param name="dtSource">源DataTable</param>
        /// <param name="strHeaderText">表头文本</param>
        /// <param name="strFileName">保存位置</param>
        public static void Export(DataTable dtSource, string strHeaderText, string strFileName)
        {
            using (MemoryStream ms = Export(dtSource, strHeaderText))
            {
                using (var fs = new FileStream(strFileName, FileMode.Create, FileAccess.Write))
                {
                    byte[] data = ms.ToArray();
                    fs.Write(data, 0, data.Length);
                    fs.Flush();
                }
            }
        }

        /// <summary>
        ///     DataTable导出到Excel的MemoryStream
        /// </summary>
        /// <param name="dtSource">源DataTable</param>
        /// <param name="strHeaderText">表头文本</param>
        protected static MemoryStream Export(DataTable dtSource, string strHeaderText)
        {
            var workbook = new HSSFWorkbook();
            ISheet sheet = workbook.CreateSheet();
            sheet.ForceFormulaRecalculation = true; //强制要求Excel在打开时重新计算的属性
            /*
             * CreateFreezePane()
             * 第一个参数表示要冻结的列数；
             * 第二个参数表示要冻结的行数；
             * 第三个参数表示右边区域可见的首列序号，从1开始计算；
             * 第四个参数表示下边区域可见的首行序号，也是从1开始计算
             */
            sheet.CreateFreezePane(0, 2, 0, 2); //冻结表头与列头
            sheet.DisplayGridlines = false;//显示/隐藏网格线

            #region 右击文件 属性信息

            {
                DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
                dsi.Company = "公司";
                workbook.DocumentSummaryInformation = dsi;

                SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
                si.Author = "作者"; //填加xls文件作者信息
                si.ApplicationName = "程序信息"; //填加xls文件创建程序信息
                si.LastAuthor = "最后保存者"; //填加xls文件最后保存者信息
                si.Comments = "说明信息"; //填加xls文件作者信息
                si.Title = "标题信息"; //填加xls文件标题信息
                si.Subject = "文件主题"; //填加文件主题信息
                si.CreateDateTime = DateTime.Now;
                workbook.SummaryInformation = si;
            }

            #endregion 右击文件 属性信息

            ICellStyle dateStyle = workbook.CreateCellStyle();
            dateStyle.Alignment = HorizontalAlignment.Center;
            IDataFormat format = workbook.CreateDataFormat();
            dateStyle.DataFormat = format.GetFormat("yyyy-mm-dd");
            dateStyle.BorderBottom = BorderStyle.Thin;
            dateStyle.BorderLeft = BorderStyle.Thin;
            dateStyle.BorderRight = BorderStyle.Thin;
            dateStyle.BorderTop = BorderStyle.Thin;
            //取得列宽
            var arrColWidth = new int[dtSource.Columns.Count];
            foreach (DataColumn item in dtSource.Columns)
            {
                arrColWidth[item.Ordinal] = Encoding.GetEncoding(936).GetBytes(item.ColumnName).Length;
            }
            for (int i = 0; i < dtSource.Rows.Count; i++)
            {
                for (int j = 0; j < dtSource.Columns.Count; j++)
                {
                    int intTemp = Encoding.GetEncoding(936).GetBytes(dtSource.Rows[i][j].ToString()).Length;
                    if (intTemp > arrColWidth[j])
                    {
                        if (intTemp > 60)
                        {
                            arrColWidth[j] = 60;
                        }
                        else
                        {
                            arrColWidth[j] = intTemp;
                        }
                    }
                }
            }

            int rowIndex = 0;
            ICellStyle cellStyle = workbook.CreateCellStyle();

            cellStyle.VerticalAlignment = VerticalAlignment.Center;
            cellStyle.Alignment = HorizontalAlignment.Left;
            cellStyle.BorderBottom = BorderStyle.Thin;
            cellStyle.BorderLeft = BorderStyle.Thin;
            cellStyle.BorderRight = BorderStyle.Thin;
            cellStyle.BorderTop = BorderStyle.Thin;
            cellStyle.WrapText = true;

            foreach (DataRow row in dtSource.Rows)
            {
                #region 新建表，填充表头，填充列头，样式

                if (rowIndex == 65535 || rowIndex == 0)
                {
                    if (rowIndex != 0)
                    {
                        sheet = workbook.CreateSheet();
                    }

                    #region 表头及样式

                    {
                        IRow headerRow = sheet.CreateRow(0);
                        headerRow.HeightInPoints = 25;
                        headerRow.CreateCell(0).SetCellValue(strHeaderText);

                        ICellStyle headStyle = workbook.CreateCellStyle();
                        headStyle.Alignment = HorizontalAlignment.Center;
                        headStyle.VerticalAlignment = VerticalAlignment.Center;
                        headStyle.VerticalAlignment = VerticalAlignment.Center;
                        headStyle.BorderBottom = BorderStyle.Thin;
                        headStyle.BorderLeft = BorderStyle.Thin;
                        headStyle.BorderRight = BorderStyle.Thin;
                        headStyle.BorderTop = BorderStyle.Thin;

                        IFont font = workbook.CreateFont();
                        font.FontName = "仿宋";
                        font.FontHeightInPoints = 16;
                        font.Boldweight = 700;
                        headStyle.SetFont(font);

                        headerRow.GetCell(0).CellStyle = headStyle;
                        var vra = new CellRangeAddress(0, 0, 0, dtSource.Columns.Count - 1);
                        sheet.AddMergedRegion(vra);
                    }

                    #endregion 表头及样式

                    #region 列头及样式

                    {
                        IRow headerRow = sheet.CreateRow(1);
                        headerRow.HeightInPoints = (float)21.75;
                        ICellStyle headStyle = workbook.CreateCellStyle();
                        headStyle.Alignment = HorizontalAlignment.Center; // CellHorizontalAlignment.CENTER;
                        headStyle.VerticalAlignment = VerticalAlignment.Center;
                        headStyle.BorderBottom = BorderStyle.Thin;
                        headStyle.BorderLeft = BorderStyle.Thin;
                        headStyle.BorderRight = BorderStyle.Thin;
                        headStyle.BorderTop = BorderStyle.Thin;

                        IFont font = workbook.CreateFont();
                        font.FontHeightInPoints = 12;
                        font.Boldweight = 700;
                        font.FontName = "仿宋";
                        headStyle.SetFont(font);

                        foreach (DataColumn column in dtSource.Columns)
                        {
                            headerRow.CreateCell(column.Ordinal).SetCellValue(column.ColumnName);
                            headerRow.GetCell(column.Ordinal).CellStyle = headStyle;

                            //设置列宽
                            sheet.SetColumnWidth(column.Ordinal, (arrColWidth[column.Ordinal] + 5) * 256);
                        }
                    }

                    #endregion 列头及样式

                    rowIndex = 2;
                }

                #endregion 新建表，填充表头，填充列头，样式

                #region 填充内容

                IRow dataRow = sheet.CreateRow(rowIndex);
                dataRow.HeightInPoints = (float)21.75;
                foreach (DataColumn column in dtSource.Columns)
                {
                    ICell newCell = dataRow.CreateCell(column.Ordinal);
                    //ICellStyle cellStyle = workbook.CreateCellStyle();

                    //cellStyle.VerticalAlignment = VerticalAlignment.CENTER;
                    //cellStyle.Alignment = HorizontalAlignment.LEFT;
                    //cellStyle.BorderBottom = BorderStyle.THIN;
                    //cellStyle.BorderLeft = BorderStyle.THIN;
                    //cellStyle.BorderRight = BorderStyle.THIN;
                    //cellStyle.BorderTop = BorderStyle.THIN;

                    newCell.CellStyle = cellStyle;

                    string drValue = row[column].ToString();

                    switch (column.DataType.ToString())
                    {
                        case "System.String": //字符串类型
                            newCell.SetCellValue(drValue);
                            break;

                        case "System.DateTime": //日期类型
                            if (drValue == "")
                            {
                                newCell.SetCellValue(drValue);
                            }
                            else
                            {
                                DateTime dateV;
                                DateTime.TryParse(drValue, out dateV);
                                newCell.SetCellValue(dateV);
                                newCell.CellStyle = dateStyle; //格式化显示
                            }

                            break;

                        case "System.Boolean": //布尔型
                            bool boolV = false;
                            bool.TryParse(drValue, out boolV);
                            newCell.SetCellValue(boolV);
                            cellStyle.Alignment = HorizontalAlignment.Center;

                            break;

                        case "System.Int16": //整型
                        case "System.Int32":
                        case "System.Byte":
                            int intV = 0;
                            int.TryParse(drValue, out intV);
                            newCell.SetCellValue(intV);
                            cellStyle.Alignment = HorizontalAlignment.Right;

                            break;

                        case "System.Int64":
                            int intV2 = 0;
                            int.TryParse(drValue, out intV2);
                            newCell.SetCellValue(intV2);
                            cellStyle.Alignment = HorizontalAlignment.Center;
                            break;

                        case "System.Decimal": //浮点型
                        case "System.Double":
                            double doubV = 0;
                            double.TryParse(drValue, out doubV);
                            newCell.SetCellValue(doubV);
                            cellStyle.Alignment = HorizontalAlignment.Right;

                            break;

                        case "System.DBNull": //空值处理
                            newCell.SetCellValue("");

                            break;

                        default:
                            newCell.SetCellValue("");
                            break;
                    }
                }

                #endregion 填充内容

                rowIndex++;
            }

            using (var ms = new MemoryStream())
            {
                workbook.Write(ms);
                ms.Flush();
                ms.Position = 0;
                return ms;
            }
        }

        /// <summary>
        ///     DataTable导出到Excel文件
        /// </summary>
        /// <param name="dtSource">源DataTable</param>
        /// <param name="chart1"></param>
        /// <param name="strHeaderText">表头文本</param>
        /// <param name="strFileName">保存位置</param>
        public static void ExportControl(DataTable dtSource, Chart chart1, string strHeaderText, string strFileName)
        {
            using (MemoryStream ms = ExportControl(dtSource, chart1, strHeaderText))
            {
                using (var fs = new FileStream(strFileName, FileMode.Create, FileAccess.Write))
                {
                    byte[] data = ms.ToArray();
                    fs.Write(data, 0, data.Length);
                    fs.Flush();
                }
            }
        }

        /// <summary>
        ///     DataTable导出到Excel文件
        /// </summary>
        /// <param name="dtSource">源DataTable</param>
        /// <param name="chart1"></param>
        /// <param name="strHeaderText">表头文本</param>
        /// <param name="strFileName">文件保存</param>
        /// <param name="top">true则chart在上，false则datatable在上</param>
        public static void ExportControl(DataTable dtSource, Chart chart1, string strHeaderText, string strFileName, bool top)
        {
            if (top)
            {
                using (MemoryStream ms = ExportControl(dtSource, chart1, strHeaderText))
                {
                    using (var fs = new FileStream(strFileName, FileMode.Create, FileAccess.Write))
                    {
                        byte[] data = ms.ToArray();
                        fs.Write(data, 0, data.Length);
                        fs.Flush();
                    }
                }
            }
            else
            {
                using (MemoryStream ms = TopExportControl(dtSource, chart1, strHeaderText))
                {
                    using (var fs = new FileStream(strFileName, FileMode.Create, FileAccess.Write))
                    {
                        byte[] data = ms.ToArray();
                        fs.Write(data, 0, data.Length);
                        fs.Flush();
                    }
                }
            }
        }

        public static MemoryStream TopExportControl(DataTable dtSource, Chart chart1, string strHeaderText)
        {
            var workbook = new HSSFWorkbook();
            ISheet sheet = workbook.CreateSheet();
            sheet.ForceFormulaRecalculation = true; //强制要求Excel在打开时重新计算的属性
            /*
             * CreateFreezePane()
             * 第一个参数表示要冻结的列数；
             * 第二个参数表示要冻结的行数；
             * 第三个参数表示右边区域可见的首列序号，从1开始计算；
             * 第四个参数表示下边区域可见的首行序号，也是从1开始计算
             */
            sheet.CreateFreezePane(0, 1, 0, 1); //冻结表头与列头
            sheet.DisplayGridlines = false;

            #region 右击文件 属性信息

            {
                DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
                dsi.Company = "公司";
                workbook.DocumentSummaryInformation = dsi;

                SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
                si.Author = "作者"; //填加xls文件作者信息
                si.ApplicationName = "程序信息"; //填加xls文件创建程序信息
                si.LastAuthor = "最后保存者"; //填加xls文件最后保存者信息
                si.Comments = "说明信息"; //填加xls文件作者信息
                si.Title = "标题信息"; //填加xls文件标题信息
                si.Subject = "文件主题"; //填加文件主题信息
                si.CreateDateTime = DateTime.Now;
                workbook.SummaryInformation = si;
            }

            #endregion 右击文件 属性信息

            ICellStyle dateStyle = workbook.CreateCellStyle();
            dateStyle.Alignment = HorizontalAlignment.Center;
            IDataFormat format = workbook.CreateDataFormat();
            dateStyle.DataFormat = format.GetFormat("yyyy-mm-dd");
            dateStyle.BorderBottom = BorderStyle.Thin;
            dateStyle.BorderLeft = BorderStyle.Thin;
            dateStyle.BorderRight = BorderStyle.Thin;
            dateStyle.BorderTop = BorderStyle.Thin;

            //取得列宽
            var arrColWidth = new int[dtSource.Columns.Count];
            foreach (DataColumn item in dtSource.Columns)
            {
                arrColWidth[item.Ordinal] = Encoding.GetEncoding(936).GetBytes(item.ColumnName).Length;
            }
            for (int i = 0; i < dtSource.Rows.Count; i++)
            {
                for (int j = 0; j < dtSource.Columns.Count; j++)
                {
                    int intTemp = Encoding.GetEncoding(936).GetBytes(dtSource.Rows[i][j].ToString()).Length;
                    if (intTemp > arrColWidth[j])
                    {
                        if (intTemp > 60)
                        {
                            arrColWidth[j] = 60;
                        }
                        else
                        {
                            arrColWidth[j] = intTemp;
                        }
                    }
                }
            }

            //int widthindex = (int)chart1.Width.Value / 70;
            //int rowIndex = (int)chart1.Height.Value / 16 + 1;
            int rowIndex = 1;
            int temp = rowIndex;

            ICellStyle cellStyle = workbook.CreateCellStyle();
            cellStyle.VerticalAlignment = VerticalAlignment.Center;
            cellStyle.Alignment = HorizontalAlignment.Left;
            cellStyle.BorderBottom = BorderStyle.Thin;
            cellStyle.BorderLeft = BorderStyle.Thin;
            cellStyle.BorderRight = BorderStyle.Thin;
            cellStyle.BorderTop = BorderStyle.Thin;
            cellStyle.WrapText = true;

            foreach (DataRow row in dtSource.Rows)
            {
                #region 新建表，填充表头，填充列头，样式

                if (rowIndex == 65535 || rowIndex == temp)
                {
                    #region 表头及样式

                    {
                        IRow headerRow = sheet.CreateRow(0);

                        headerRow.HeightInPoints = 25;
                        var cellRangeAddress = new CellRangeAddress(0, 0, 0, dtSource.Columns.Count + 1);
                        sheet.AddMergedRegion(cellRangeAddress);
                        headerRow.CreateCell(0).SetCellValue(strHeaderText);

                        ICellStyle headStyle = workbook.CreateCellStyle();
                        headStyle.Alignment = HorizontalAlignment.Center;
                        headStyle.BorderBottom = BorderStyle.Thin;
                        headStyle.BorderLeft = BorderStyle.Thin;
                        headStyle.BorderRight = BorderStyle.Thin;
                        headStyle.BorderTop = BorderStyle.Thin;

                        IFont font = workbook.CreateFont();
                        font.FontHeightInPoints = 20;
                        font.Boldweight = 700;
                        headStyle.SetFont(font);

                        headerRow.GetCell(0).CellStyle = headStyle;
                        var vra = new CellRangeAddress(0, 0, 0, dtSource.Columns.Count - 1);
                        sheet.AddMergedRegion(vra);
                    }

                    #endregion 表头及样式

                    #region 列头及样式

                    {
                        IRow headerRow = sheet.CreateRow(rowIndex);
                        headerRow.HeightInPoints = (float)21.75;
                        ICellStyle headStyle = workbook.CreateCellStyle();
                        headStyle.Alignment = HorizontalAlignment.Center; // CellHorizontalAlignment.CENTER;
                        headStyle.VerticalAlignment = VerticalAlignment.Center;
                        headStyle.BorderBottom = BorderStyle.Thin;
                        headStyle.BorderLeft = BorderStyle.Thin;
                        headStyle.BorderRight = BorderStyle.Thin;
                        headStyle.BorderTop = BorderStyle.Thin;

                        IFont font = workbook.CreateFont();
                        font.FontHeightInPoints = 12;
                        font.Boldweight = 700;
                        font.FontName = "仿宋";
                        headStyle.SetFont(font);

                        foreach (DataColumn column in dtSource.Columns)
                        {
                            headerRow.CreateCell(column.Ordinal).SetCellValue(column.ColumnName);
                            headerRow.GetCell(column.Ordinal).CellStyle = headStyle;

                            //设置列宽
                            sheet.SetColumnWidth(column.Ordinal, (arrColWidth[column.Ordinal] + 5) * 256);
                        }
                    }

                    #endregion 列头及样式

                    rowIndex++;
                }

                #endregion 新建表，填充表头，填充列头，样式

                #region 填充内容

                IRow dataRow = sheet.CreateRow(rowIndex);
                dataRow.HeightInPoints = (float)21.75;

                foreach (DataColumn column in dtSource.Columns)
                {
                    ICell newCell = dataRow.CreateCell(column.Ordinal);

                    newCell.CellStyle = cellStyle;

                    string drValue = row[column].ToString();

                    switch (column.DataType.ToString())
                    {
                        case "System.String": //字符串类型
                            newCell.SetCellValue(drValue);
                            break;

                        case "System.DateTime": //日期类型
                            if (drValue == "")
                            {
                                newCell.SetCellValue(drValue);
                            }
                            else
                            {
                                DateTime dateV;
                                DateTime.TryParse(drValue, out dateV);
                                newCell.SetCellValue(dateV);
                                newCell.CellStyle = dateStyle; //格式化显示
                            }

                            break;

                        case "System.Boolean": //布尔型
                            bool boolV = false;
                            bool.TryParse(drValue, out boolV);
                            newCell.SetCellValue(boolV);
                            cellStyle.Alignment = HorizontalAlignment.Center;

                            break;

                        case "System.Int16": //整型
                        case "System.Int32":
                        case "System.Byte":
                            int intV = 0;
                            int.TryParse(drValue, out intV);
                            newCell.SetCellValue(intV);
                            cellStyle.Alignment = HorizontalAlignment.Right;

                            break;

                        case "System.Int64":
                            int intV2 = 0;
                            int.TryParse(drValue, out intV2);
                            newCell.SetCellValue(intV2);
                            cellStyle.Alignment = HorizontalAlignment.Center;
                            break;

                        case "System.Decimal": //浮点型
                        case "System.Double":
                            double doubV = 0;
                            double.TryParse(drValue, out doubV);
                            newCell.SetCellValue(doubV);
                            cellStyle.Alignment = HorizontalAlignment.Right;

                            break;

                        case "System.DBNull": //空值处理
                            newCell.SetCellValue("");

                            break;

                        default:
                            newCell.SetCellValue("");
                            break;
                    }
                }

                #endregion 填充内容

                rowIndex++;
            }

            string fileName = DateTime.Now.ToString("yyyyMMddhhMMss") + ".jpg";
            string basePath = string.Format("C:\\Program Files\\Authine\\H3\\Portal\\TempImages\\{0}", fileName); //指定图片生成的路径，用来下载，最后要删除该路径下的图片
            chart1.SaveImage(basePath); //导出图片

            byte[] bytes = File.ReadAllBytes(basePath);
            int pictureIdx = workbook.AddPicture(bytes, PictureType.JPEG);

            // Create the drawing patriarch.  This is the top level container for all shapes.
            var patriarch = (HSSFPatriarch)sheet.CreateDrawingPatriarch();

            //add a picture
            var anchor = new HSSFClientAnchor(0, 0, 1023, 0, 0, dtSource.Rows.Count + 2, 0, 0);
            var pict = (HSSFPicture)patriarch.CreatePicture(anchor, pictureIdx);
            pict.Resize();
            //删除服务器上临时文件
            File.Delete(basePath);

            using (var ms = new MemoryStream())
            {
                workbook.Write(ms);
                ms.Flush();
                ms.Position = 0;
                return ms;
            }
        }

        public static MemoryStream ExportControl(DataTable dtSource, Chart chart1, string strHeaderText)
        {
            var workbook = new HSSFWorkbook();
            ISheet sheet = workbook.CreateSheet();
            sheet.ForceFormulaRecalculation = true; //强制要求Excel在打开时重新计算的属性
            /*
             * CreateFreezePane()
             * 第一个参数表示要冻结的列数；
             * 第二个参数表示要冻结的行数；
             * 第三个参数表示右边区域可见的首列序号，从1开始计算；
             * 第四个参数表示下边区域可见的首行序号，也是从1开始计算
             */
            sheet.CreateFreezePane(0, 1, 0, 1); //冻结表头与列头
            sheet.DisplayGridlines = false;//显示/隐藏网格线

            #region 右击文件 属性信息

            {
                DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
                dsi.Company = "公司";
                workbook.DocumentSummaryInformation = dsi;

                SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
                si.Author = "作者"; //填加xls文件作者信息
                si.ApplicationName = "程序信息"; //填加xls文件创建程序信息
                si.LastAuthor = "最后保存者"; //填加xls文件最后保存者信息
                si.Comments = "说明信息"; //填加xls文件作者信息
                si.Title = "标题信息"; //填加xls文件标题信息
                si.Subject = "文件主题"; //填加文件主题信息
                si.CreateDateTime = DateTime.Now;
                workbook.SummaryInformation = si;
            }

            #endregion 右击文件 属性信息

            ICellStyle dateStyle = workbook.CreateCellStyle();
            dateStyle.Alignment = HorizontalAlignment.Center;
            IDataFormat format = workbook.CreateDataFormat();
            dateStyle.DataFormat = format.GetFormat("yyyy-mm-dd");
            dateStyle.BorderBottom = BorderStyle.Thin;
            dateStyle.BorderLeft = BorderStyle.Thin;
            dateStyle.BorderRight = BorderStyle.Thin;
            dateStyle.BorderTop = BorderStyle.Thin;

            //取得列宽
            var arrColWidth = new int[dtSource.Columns.Count];
            foreach (DataColumn item in dtSource.Columns)
            {
                arrColWidth[item.Ordinal] = Encoding.GetEncoding(936).GetBytes(item.ColumnName).Length;
            }
            for (int i = 0; i < dtSource.Rows.Count; i++)
            {
                for (int j = 0; j < dtSource.Columns.Count; j++)
                {
                    int intTemp = Encoding.GetEncoding(936).GetBytes(dtSource.Rows[i][j].ToString()).Length;
                    if (intTemp > arrColWidth[j])
                    {
                        if (intTemp > 60)
                        {
                            arrColWidth[j] = 60;
                        }
                        else
                        {
                            arrColWidth[j] = intTemp;
                        }
                    }
                }
            }

            string fileName = DateTime.Now.ToString("yyyyMMddhhMMss") + ".jpg";
            string basePath = string.Format("C:\\Program Files\\Authine\\H3\\Portal\\TempImages\\{0}", fileName); //指定图片生成的路径，用来下载，最后要删除该路径下的图片
            chart1.SaveImage(basePath); //导出图片

            byte[] bytes = File.ReadAllBytes(basePath);
            int pictureIdx = workbook.AddPicture(bytes, PictureType.JPEG);

            // Create the drawing patriarch.  This is the top level container for all shapes.
            var patriarch = (HSSFPatriarch)sheet.CreateDrawingPatriarch();

            //add a picture
            var anchor = new HSSFClientAnchor(0, 0, 1023, 0, 0, 1, 0, 0);
            var pict = (HSSFPicture)patriarch.CreatePicture(anchor, pictureIdx);
            pict.Resize();
            //删除服务器上临时文件
            File.Delete(basePath);
            int widthindex = (int)chart1.Width.Value / 70;
            int rowIndex = (int)chart1.Height.Value / 16 + 1;
            int temp = rowIndex;

            ICellStyle cellStyle = workbook.CreateCellStyle();
            cellStyle.VerticalAlignment = VerticalAlignment.Center;
            cellStyle.Alignment = HorizontalAlignment.Left;
            cellStyle.BorderBottom = BorderStyle.Thin;
            cellStyle.BorderLeft = BorderStyle.Thin;
            cellStyle.BorderRight = BorderStyle.Thin;
            cellStyle.BorderTop = BorderStyle.Thin;
            cellStyle.WrapText = true;
            foreach (DataRow row in dtSource.Rows)
            {
                #region 新建表，填充表头，填充列头，样式

                if (rowIndex == 65535 || rowIndex == temp)
                {
                    //if (rowIndex != 0)
                    //{
                    //    sheet = workbook.CreateSheet();
                    //}

                    #region 表头及样式

                    {
                        IRow headerRow = sheet.CreateRow(0);

                        headerRow.HeightInPoints = 25;
                        var cellRangeAddress = new CellRangeAddress(0, 0, 0, widthindex + 1);
                        sheet.AddMergedRegion(cellRangeAddress);
                        headerRow.CreateCell(0).SetCellValue(strHeaderText);

                        ICellStyle headStyle = workbook.CreateCellStyle();
                        headStyle.Alignment = HorizontalAlignment.Center;
                        headStyle.BorderBottom = BorderStyle.Thin;
                        headStyle.BorderLeft = BorderStyle.Thin;
                        headStyle.BorderRight = BorderStyle.Thin;
                        headStyle.BorderTop = BorderStyle.Thin;

                        IFont font = workbook.CreateFont();
                        font.FontHeightInPoints = 20;
                        font.Boldweight = 700;
                        headStyle.SetFont(font);

                        headerRow.GetCell(0).CellStyle = headStyle;
                        var vra = new CellRangeAddress(0, 0, 0, dtSource.Columns.Count - 1);
                        sheet.AddMergedRegion(vra);
                    }

                    #endregion 表头及样式

                    #region 列头及样式

                    {
                        IRow headerRow = sheet.CreateRow(rowIndex);
                        headerRow.HeightInPoints = (float)21.75;
                        ICellStyle headStyle = workbook.CreateCellStyle();
                        headStyle.Alignment = HorizontalAlignment.Center; // CellHorizontalAlignment.CENTER;
                        headStyle.VerticalAlignment = VerticalAlignment.Center;
                        headStyle.BorderBottom = BorderStyle.Thin;
                        headStyle.BorderLeft = BorderStyle.Thin;
                        headStyle.BorderRight = BorderStyle.Thin;
                        headStyle.BorderTop = BorderStyle.Thin;

                        IFont font = workbook.CreateFont();
                        font.FontHeightInPoints = 12;
                        font.Boldweight = 700;
                        font.FontName = "仿宋";
                        headStyle.SetFont(font);

                        foreach (DataColumn column in dtSource.Columns)
                        {
                            headerRow.CreateCell(column.Ordinal).SetCellValue(column.ColumnName);
                            headerRow.GetCell(column.Ordinal).CellStyle = headStyle;

                            //设置列宽
                            sheet.SetColumnWidth(column.Ordinal, (arrColWidth[column.Ordinal] + 5) * 256);
                        }
                    }

                    #endregion 列头及样式

                    rowIndex++;
                }

                #endregion 新建表，填充表头，填充列头，样式

                #region 填充内容

                IRow dataRow = sheet.CreateRow(rowIndex);
                dataRow.HeightInPoints = (float)21.75;

                foreach (DataColumn column in dtSource.Columns)
                {
                    ICell newCell = dataRow.CreateCell(column.Ordinal);


                    newCell.CellStyle = cellStyle;

                    string drValue = row[column].ToString();

                    switch (column.DataType.ToString())
                    {
                        case "System.String": //字符串类型
                            newCell.SetCellValue(drValue);
                            break;

                        case "System.DateTime": //日期类型
                            if (drValue == "")
                            {
                                newCell.SetCellValue(drValue);
                            }
                            else
                            {
                                DateTime dateV;
                                DateTime.TryParse(drValue, out dateV);
                                newCell.SetCellValue(dateV);
                                newCell.CellStyle = dateStyle; //格式化显示
                            }

                            break;

                        case "System.Boolean": //布尔型
                            bool boolV = false;
                            bool.TryParse(drValue, out boolV);
                            newCell.SetCellValue(boolV);
                            cellStyle.Alignment = HorizontalAlignment.Center;

                            break;

                        case "System.Int16": //整型
                        case "System.Int32":
                        case "System.Byte":
                            int intV = 0;
                            int.TryParse(drValue, out intV);
                            newCell.SetCellValue(intV);
                            cellStyle.Alignment = HorizontalAlignment.Right;

                            break;

                        case "System.Int64":
                            int intV2 = 0;
                            int.TryParse(drValue, out intV2);
                            newCell.SetCellValue(intV2);
                            cellStyle.Alignment = HorizontalAlignment.Center;
                            break;

                        case "System.Decimal": //浮点型
                        case "System.Double":
                            double doubV = 0;
                            double.TryParse(drValue, out doubV);
                            newCell.SetCellValue(doubV);
                            cellStyle.Alignment = HorizontalAlignment.Right;

                            break;

                        case "System.DBNull": //空值处理
                            newCell.SetCellValue("");

                            break;

                        default:
                            newCell.SetCellValue("");
                            break;
                    }
                }

                #endregion 填充内容

                rowIndex++;
            }

            using (var ms = new MemoryStream())
            {
                workbook.Write(ms);
                ms.Flush();
                ms.Position = 0;
                return ms;
            }
        }

        /// <summary>
        ///     用于Web导出
        /// </summary>
        /// <param name="dtSource">源DataTable</param>
        /// <param name="strHeaderText">表头文本</param>
        /// <param name="strFileName">文件名</param>
        public static void ExportByWeb(DataTable dtSource, string strHeaderText, string strFileName)
        {
            HttpContext curContext = HttpContext.Current;
            string browserType = curContext.Request.Browser.Browser;

            // 设置编码和附件格式
            curContext.Response.ContentType = "application/vnd.ms-excel";
            curContext.Response.ContentEncoding = Encoding.UTF8;
            curContext.Response.Charset = "";
            if (browserType != "Firefox")
            {
                curContext.Response.AppendHeader("Content-Disposition",
                                                 "attachment;filename=" + HttpUtility.UrlEncode(strFileName, Encoding.UTF8));
            }
            else
            {
                curContext.Response.AppendHeader("Content-Disposition",
                                                     "attachment;filename=" + HttpUtility.UrlDecode(strFileName, Encoding.UTF8));
            }
            byte[] by = Export(dtSource, strHeaderText).ToArray();
            curContext.Response.BinaryWrite(by);
            curContext.Response.End();
        }

        /// <summary>
        /// 用于Web导出(同时导出表和图)
        /// </summary>
        /// <param name="dtSource">源DataTable</param>
        /// <param name="chart1"></param>
        /// <param name="strHeaderText">表头文本</param>
        /// <param name="strFileName">文件名</param>
        public static void ExportByWeb(DataTable dtSource, Chart chart1, string strHeaderText, string strFileName, bool top)
        {
            HttpContext curContext = HttpContext.Current;

            // 设置编码和附件格式
            curContext.Response.ContentType = "application/vnd.ms-excel;charset=UTF-8";
            //curContext.Response.ContentEncoding = Encoding.UTF8;
            curContext.Response.ContentEncoding = Encoding.Default;
            curContext.Response.Charset = "";
            curContext.Response.AppendHeader("Content-Disposition",
                                             "attachment;filename=" + HttpUtility.UrlEncode(strFileName, Encoding.UTF8));
            if (top)
            {
                curContext.Response.BinaryWrite(ExportControl(dtSource, chart1, strHeaderText).GetBuffer());
            }
            else
            {
                curContext.Response.BinaryWrite(TopExportControl(dtSource, chart1, strHeaderText).GetBuffer());
            }
            curContext.Response.End();
        }

        /// <summary>
        /// 用于Web导出(同时导出表和图)
        /// </summary>
        /// <param name="dtSource">源DataTable</param>
        /// <param name="chart1"></param>
        /// <param name="strHeaderText">表头文本</param>
        /// <param name="strFileName">文件名</param>
        public static void ExportByWeb(DataTable dtSource, Chart chart1, string strHeaderText, string strFileName)
        {
            HttpContext curContext = HttpContext.Current;

            // 设置编码和附件格式
            curContext.Response.ContentType = "application/vnd.ms-excel";
            //curContext.Response.ContentEncoding = Encoding.UTF8;
            curContext.Response.ContentEncoding = Encoding.Default;
            curContext.Response.Charset = "";
            curContext.Response.AppendHeader("Content-Disposition",
                                             "attachment;filename=" + HttpUtility.UrlEncode(strFileName, Encoding.UTF8));

            curContext.Response.BinaryWrite(ExportControl(dtSource, chart1, strHeaderText).GetBuffer());
            curContext.Response.End();
        }

        #endregion 导出

        /// <summary>
        /// 2.1.1 复制原有sheet的合并单元格到新创建的sheet
        /// </summary>
        /// <param name="fromSheet"></param>
        /// <param name="toSheet"></param>
        public static void MergerRegion(ISheet fromSheet, ISheet toSheet)
        {
            int sheetMergerCount = fromSheet.NumMergedRegions;
            for (int i = 0; i < sheetMergerCount; i++)
            {
                toSheet.AddMergedRegion(fromSheet.GetMergedRegion(i));
            }
        }

        /// <summary>
        /// 2.1.2 复制行
        /// </summary>
        public static void CopyRow(IWorkbook wb, IRow fromRow, IRow toRow, bool copyValueFlag)
        {
            System.Collections.IEnumerator cells = fromRow.GetEnumerator();
            toRow.Height = fromRow.Height;
            while (cells.MoveNext())
            {
                ICell cell = null;
                if (wb is HSSFWorkbook)
                    cell = cells.Current as HSSFCell;
                else
                    cell = cells.Current as NPOI.XSSF.UserModel.XSSFCell;
                ICell newCell = toRow.CreateCell(cell.ColumnIndex);
                CopyCell(wb, cell, newCell, copyValueFlag);
            }
        }

        /// <summary>
        /// 2.1.3 复制单元格
        /// </summary>
        public static void CopyCell(IWorkbook wb, ICell srcCell, ICell distCell, bool copyValueFlag)
        {
            ICellStyle newstyle = wb.CreateCellStyle();
            distCell.CellStyle = newstyle;
            //评论
            if (srcCell.CellComment != null)
            {
                distCell.CellComment = srcCell.CellComment;
            }
            // 不同数据类型处理
            CellType srcCellType = srcCell.CellType;
            distCell.SetCellType(srcCellType);
            if (copyValueFlag)
            {
                if (srcCellType == CellType.Numeric)
                {
                    if (DateUtil.IsCellDateFormatted(srcCell))
                    {
                        //在poi中日期是以double类型表示的，所以要格式化
                        SimpleDateFormat sdf = new SimpleDateFormat("yyyy/M/d");
                        distCell.SetCellValue(sdf.Format(Convert.ToDateTime(srcCell.DateCellValue), new System.Globalization.CultureInfo("en-us")));
                        //distCell.SetCellValue(Convert.ToDateTime(srcCell.DateCellValue));
                    }
                    else
                    {
                        distCell.SetCellValue(srcCell.NumericCellValue);
                    }
                }
                else if (srcCellType == CellType.String)
                {
                    distCell.SetCellValue(srcCell.RichStringCellValue);
                }
                else if (srcCellType == CellType.Blank)
                {
                }
                else if (srcCellType == CellType.Boolean)
                {
                    distCell.SetCellValue(srcCell.BooleanCellValue);
                }
                else if (srcCellType == CellType.Error)
                {
                    distCell.SetCellErrorValue(srcCell.ErrorCellValue);
                }
                else if (srcCellType == CellType.Formula)
                {
                    distCell.SetCellFormula(srcCell.CellFormula);
                }
                else
                {
                }
            }
        }


        /// <summary>
        /// 1.1 将Excel转成CSV文件
        /// </summary>
        /// <param name="builder"></param>
        /// <param name="saveUrlStr">CSV保存路径</param>
        public static void ConvertToCSV(ImportData.IImportData builder, string saveUrlStr, string csvNameStr)
        {
            builder.ExportType = ImportData.DataType.Csv;
            ExportSource.CSVExportSource csv = builder.Export as ExportSource.CSVExportSource;
            csv.FilePath = string.Format("{0}\\{1}", saveUrlStr, csvNameStr);
            builder.Export.Export();
        }

        /// <summary>
        /// 1.2 将CSV文件的数据读取到DataTable中
        /// </summary>
        /// <param name="fileName">CSV文件路径</param>
        /// <returns>返回读取了CSV数据的DataTable</returns>
        public static DataTable OpenCSV(string filePath)
        {
            Encoding encoding = System.Text.UnicodeEncoding.GetEncoding("GB2312");//Common.GetType(filePath); //Encoding.ASCII;//
            DataTable dt = new DataTable();
            FileStream fs = new FileStream(filePath, System.IO.FileMode.Open, System.IO.FileAccess.Read);

            StreamReader sr = new StreamReader(fs, encoding);
            //string fileContent = sr.ReadToEnd();
            //encoding = sr.CurrentEncoding;
            //记录每次读取的一行记录
            string strLine = "";
            //记录每行记录中的各字段内容
            string[] aryLine = null;
            string[] tableHead = null;
            //标示列数
            int columnCount = 0;
            //标示是否是读取的第一行
            bool IsFirst = true;
            //逐行读取CSV中的数据
            while ((strLine = sr.ReadLine()) != null)
            {
                if (IsFirst == true)
                {
                    tableHead = strLine.Split(',');
                    IsFirst = false;
                    columnCount = tableHead.Length;
                    //创建列
                    for (int i = 0; i < columnCount; i++)
                    {
                        DataColumn dc = new DataColumn(tableHead[i].ToString());
                        dt.Columns.Add(dc);
                    }
                    //DataRow dr = dt.NewRow();
                    //for (int i = 0; i < columnCount; i++)
                    //{
                    //    dr[i] = tableHead[i];
                    //}
                    //dt.Rows.Add(dr);
                }
                else
                {
                    aryLine = strLine.Split(',');
                    DataRow dr = dt.NewRow();
                    columnCount = aryLine.Length;
                    for (int j = 0; j < columnCount; j++)
                    {
                        dr[tableHead[j]] = aryLine[j];
                    }
                    dt.Rows.Add(dr);
                }
            }
            sr.Close();
            fs.Close();
            return dt;
        }
        /// <summary>
        /// 1.2 将CSV文件的数据读取到DataTable中maben 
        /// </summary>
        /// <param name="fileName">CSV文件路径</param>
        /// <returns>返回读取了CSV数据的DataTable</returns>
        public static DataTable OpenCSV(string filePath, int columnCount)
        {
            Encoding encoding = System.Text.UnicodeEncoding.GetEncoding("GB2312");//Common.GetType(filePath); //Encoding.ASCII;//
            DataTable dt = new DataTable();
            FileStream fs = new FileStream(filePath, System.IO.FileMode.Open, System.IO.FileAccess.Read);

            StreamReader sr = new StreamReader(fs, encoding);
            //string fileContent = sr.ReadToEnd();
            //encoding = sr.CurrentEncoding;
            //记录每次读取的一行记录
            string strLine = "";
            //记录每行记录中的各字段内容
            string[] aryLine = null;
            string[] tableHead = null;
            //标示列数
            //int columnCount = 0;
            //标示是否是读取的第一行
            bool IsFirst = true;
            //逐行读取CSV中的数据
            while ((strLine = sr.ReadLine()) != null)
            {
                if (IsFirst == true)
                {
                    tableHead = strLine.Split(',');
                    IsFirst = false;
                    //columnCount = tableHead.Length;
                    //创建列
                    for (int i = 0; i < columnCount; i++)
                    {
                        DataColumn dc = new DataColumn(i.ToString());
                        dt.Columns.Add(dc);
                    }
                    DataRow dr = dt.NewRow();
                    for (int i = 0; i < columnCount; i++)
                    {
                        if (i < tableHead.Length)
                        {
                            dr[i] = tableHead[i];
                        }
                        else
                        {
                            dr[i] = "";
                        }
                    }
                    dt.Rows.Add(dr);
                }
                else
                {
                    aryLine = strLine.Split(',');
                    DataRow dr = dt.NewRow();
                    //columnCount = aryLine.Length;
                    DateTime dtime;
                    Double dbs = 0;
                    int innum = 0;
                    for (int j = 0; j < columnCount; j++)
                    {
                        if (j < aryLine.Length)
                        {
                            if (aryLine[j] == "")
                            {
                                dr[j] = aryLine[j];
                            }
                            else if (Int32.TryParse(aryLine[j], out innum))
                            {
                                dr[j] = innum.ToString();
                            }
                            else if (Double.TryParse(aryLine[j], out dbs))
                            {
                                dr[j] = dbs.ToString();
                            }
                            else if (DateTime.TryParse(aryLine[j], out dtime))
                            {
                                dr[j] = dtime.ToString("yyyy/MM/dd");
                            }
                            else
                            {
                                dr[j] = aryLine[j];
                            }

                        }
                        else
                        {
                            dr[j] = "";
                        }
                    }
                    dt.Rows.Add(dr);
                }
            }
            sr.Close();
            fs.Close();
            return dt;
        }

        /// <summary>
        /// 1.3 将DataTable转成CSV文件
        /// </summary>
        /// <param name="saveUrlStr">CSV文件路径</param>
        /// <param name="csvNameStr">CSV文件名</param>
        public static void ConvertTableToCSV(ImportData.IImportData builder, string saveUrlStr, string csvNameStr)
        {
            builder.ExportType = ImportData.DataType.Csv;
            ExportSource.CSVExportSource csv = builder.Export as ExportSource.CSVExportSource;
            csv.FilePath = string.Format("{0}\\{1}", saveUrlStr, csvNameStr);
            builder.Export.Export();
        }
        /// <summary>
        /// 获取Excel中某单元格的值
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="rowindex"></param>
        /// <param name="colindex"></param>
        /// <returns></returns>
        public static ICell GetCellVaule(ISheet sheet, int rowindex, int colindex)
        {
            HSSFRow row = (HSSFRow)sheet.GetRow(rowindex);
            return row.GetCell(colindex);
        }
        /// <summary>
        /// 下载模版时生成下拉框
        /// </summary>
        /// <param name="sheet1">Excel模版的sheet页对象</param>
        /// <param name="PositionAndSource">List中有4个数值，分别代表为 fristRow(从0开始),lastRow(65535),firstCol(从0开始),lastCol(从0开始)决定了哪行哪列拥有数据验证;第二个参数为对应下拉显示数据源</param>
        public static void CellDropDownList(IWorkbook workbook, Dictionary<List<int>, string[]> PositionAndSource, string strFileName)
        {
            if (PositionAndSource != null && PositionAndSource.Count > 0)
            {
                ISheet sheet1 = workbook.GetSheetAt(0);
                ISheet sheet2 = workbook.CreateSheet("ShtDictionary");//保存下拉数据源的sheet
                workbook.SetSheetHidden(1, SheetState.Hidden);
                int cellindex = 0;//下拉数据源在新的sheet中的列索引
                foreach (var dic in PositionAndSource)
                {
                    List<int> listposition = dic.Key;
                    string[] source = dic.Value;
                    if (source != null && source.Length > 0)
                    {
                        //数据源逐列逐行添加
                        for (int j = 1; j <= source.Length; j++)
                        {
                            IRow sheetrow = sheet2.GetRow(j);
                            if (sheetrow == null)
                            {
                                sheetrow = sheet2.CreateRow(j);
                            }
                            ICell sheetcell = sheetrow.GetCell(cellindex);
                            if (sheetcell == null)
                            {
                                sheetcell = sheetrow.CreateCell(cellindex);
                            }
                            sheetcell.SetCellValue(source[j - 1]);
                        }

                        IName range = workbook.CreateName();
                        range.RefersToFormula = GetFormulaRefers("ShtDictionary", cellindex, source.Length + 1);//获取公式范围
                        range.NameName = "dicRange" + cellindex;
                        CellRangeAddressList regions = new CellRangeAddressList(listposition[0], listposition[1], listposition[2], listposition[3]);
                        DVConstraint constraint = DVConstraint.CreateFormulaListConstraint("dicRange" + cellindex);
                        HSSFDataValidation dataValidate = new HSSFDataValidation(regions, constraint);
                        sheet1.AddValidationData(dataValidate);

                        cellindex++;
                    }
                }
                HttpContext curContext = HttpContext.Current;
                string browserType = curContext.Request.Browser.Browser;

                //设置编码和附件格式
                curContext.Response.ContentType = "application/vnd.ms-excel";
                curContext.Response.ContentEncoding = Encoding.UTF8;
                curContext.Response.Charset = "";
                if (browserType != "Firefox")
                {
                    curContext.Response.AppendHeader("Content-Disposition",
                                                     "attachment;filename=" + HttpUtility.UrlEncode(strFileName, Encoding.UTF8));
                }
                else
                {
                    curContext.Response.AppendHeader("Content-Disposition",
                                                         "attachment;filename=" + HttpUtility.UrlDecode(strFileName, Encoding.UTF8));
                }
                using (MemoryStream ms = new MemoryStream())
                {
                    //将工作簿的内容放到内存流中
                    workbook.Write(ms);
                    //将内存流转换成字节数组发送到客户端
                    curContext.Response.BinaryWrite(ms.GetBuffer());
                    curContext.Response.End();
                }
            }


        }
        /// <summary>
        /// 获取Excel下拉引用区域字符串
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="cellindex"></param>
        /// <param name="length"></param>
        /// <returns></returns>
        private static string GetFormulaRefers(string sheetName, int cellindex, int length)
        {
            switch (cellindex)
            {
                case 0:
                    return sheetName + "!$A$2:$A$" + length;
                case 1:
                    return sheetName + "!$B$2:$B$" + length;
                case 2:
                    return sheetName + "!$C$2:$C$" + length;
                case 3:
                    return sheetName + "!$D$2:$D$" + length;
                case 4:
                    return sheetName + "!$E$2:$E$" + length;
                case 5:
                    return sheetName + "!$F$2:$F$" + length;
                case 6:
                    return sheetName + "!$G$2:$G$" + length;
                case 7:
                    return sheetName + "!$H$2:$H$" + length;
                case 8:
                    return sheetName + "!$I$2:$I$" + length;
                case 9:
                    return sheetName + "!$J$2:$J$" + length;
                case 10:
                    return sheetName + "!$K$2:$K$" + length;
                default:
                    return sheetName + "!$A$2:$A$" + length;
            }

        }
        /// <summary>
        ///     用于Web导出
        /// </summary>
        /// <param name="dtSource">源DataTable</param>
        /// <param name="strHeaderText">表头文本</param>
        /// <param name="strFileName">文件名</param>
        public static void ExportByWeb(DataTable dtSource, string strHeaderText, string strFileName, List<CellRangeAddress> cellRanges = null)
        {
            HttpContext curContext = HttpContext.Current;
            string browserType = curContext.Request.Browser.Browser;

            //设置编码和附件格式
            curContext.Response.ContentType = "application/vnd.ms-excel";
            curContext.Response.ContentEncoding = Encoding.UTF8;
            curContext.Response.Charset = "";
            if (browserType != "Firefox")
            {
                curContext.Response.AppendHeader("Content-Disposition",
                                                 "attachment;filename=" + HttpUtility.UrlEncode(strFileName, Encoding.UTF8));
            }
            else
            {
                curContext.Response.AppendHeader("Content-Disposition",
                                                     "attachment;filename=" + HttpUtility.UrlDecode(strFileName, Encoding.UTF8));
            }
            curContext.Response.BinaryWrite(Export(dtSource, strHeaderText, cellRanges).GetBuffer());
            curContext.Response.End();
        }

        /// <summary>
        ///     用于Web导出
        /// </summary>
        /// <param name="dtSource">源DataTable</param>
        /// <param name="strHeaderText">表头文本</param>
        /// <param name="strFileName">文件名</param>
        public static void ExportByWeb(DataSet st, string strHeaderText, string strFileName, bool inOneSheet = false, List<CellRangeAddress> rangeList = null)
        {
            HttpContext curContext = HttpContext.Current;
            string browserType = curContext.Request.Browser.Browser;

            //设置编码和附件格式
            curContext.Response.ContentType = "application/vnd.ms-excel";
            curContext.Response.ContentEncoding = Encoding.UTF8;
            curContext.Response.Charset = "";
            if (browserType != "Firefox")
            {
                curContext.Response.AppendHeader("Content-Disposition",
                                                 "attachment;filename=" + HttpUtility.UrlEncode(strFileName, Encoding.UTF8));
            }
            else
            {
                curContext.Response.AppendHeader("Content-Disposition",
                                                     "attachment;filename=" + HttpUtility.UrlDecode(strFileName, Encoding.UTF8));
            }
            curContext.Response.BinaryWrite(ExportBatch(st, strHeaderText, inOneSheet, rangeList).GetBuffer());
            curContext.Response.End();
        }
        /// <summary>
        ///     DataTable导出到Excel的MemoryStream
        /// </summary>
        /// <param name="dtSource">源DataTable</param>
        /// <param name="strHeaderText">表头文本</param>
        protected static MemoryStream Export(DataTable dtSource, string strHeaderText, List<CellRangeAddress> cellRanges)
        {
            var workbook = new HSSFWorkbook();
            ISheet sheet = workbook.CreateSheet();
            sheet.ForceFormulaRecalculation = true; //强制要求Excel在打开时重新计算的属性
            /*
             * CreateFreezePane()
             * 第一个参数表示要冻结的列数；
             * 第二个参数表示要冻结的行数；
             * 第三个参数表示右边区域可见的首列序号，从1开始计算；
             * 第四个参数表示下边区域可见的首行序号，也是从1开始计算
             */
            sheet.CreateFreezePane(0, 2); //冻结表头与列头
            sheet.DisplayGridlines = false;//显示/隐藏网格线
            #region 右击文件 属性信息

            {
                DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
                dsi.Company = "公司";
                workbook.DocumentSummaryInformation = dsi;

                SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
                si.Author = "作者"; //填加xls文件作者信息
                si.ApplicationName = "程序信息"; //填加xls文件创建程序信息
                si.LastAuthor = "最后保存者"; //填加xls文件最后保存者信息
                si.Comments = "说明信息"; //填加xls文件作者信息
                si.Title = "标题信息"; //填加xls文件标题信息
                si.Subject = "文件主题"; //填加文件主题信息
                si.CreateDateTime = DateTime.Now;
                workbook.SummaryInformation = si;
            }

            #endregion 右击文件 属性信息

            ICellStyle dateStyle = workbook.CreateCellStyle();
            dateStyle.Alignment = HorizontalAlignment.Center;
            IDataFormat format = workbook.CreateDataFormat();
            dateStyle.DataFormat = format.GetFormat("yyyy-mm-dd");
            dateStyle.BorderBottom = BorderStyle.Thin;
            dateStyle.BorderLeft = BorderStyle.Thin;
            dateStyle.BorderRight = BorderStyle.Thin;
            dateStyle.BorderTop = BorderStyle.Thin;
            //取得列宽
            var arrColWidth = new int[dtSource.Columns.Count];
            foreach (DataColumn item in dtSource.Columns)
            {
                arrColWidth[item.Ordinal] = Encoding.GetEncoding(936).GetBytes(item.ColumnName).Length;
            }
            for (int i = 0; i < dtSource.Rows.Count; i++)
            {
                for (int j = 0; j < dtSource.Columns.Count; j++)
                {
                    int intTemp = Encoding.GetEncoding(936).GetBytes(dtSource.Rows[i][j].ToString()).Length;
                    if (intTemp > arrColWidth[j])
                    {
                        if (intTemp > 60)
                        {
                            arrColWidth[j] = 60;
                        }
                        else
                        {
                            arrColWidth[j] = intTemp;
                        }
                    }
                }
            }

            int rowIndex = 0;
            ICellStyle cellStyle = workbook.CreateCellStyle();

            cellStyle.VerticalAlignment = VerticalAlignment.Center;
            cellStyle.Alignment = HorizontalAlignment.Left;
            cellStyle.BorderBottom = BorderStyle.Thin;
            cellStyle.BorderLeft = BorderStyle.Thin;
            cellStyle.BorderRight = BorderStyle.Thin;
            cellStyle.BorderTop = BorderStyle.Thin;
            cellStyle.WrapText = true;

            foreach (DataRow row in dtSource.Rows)
            {
                #region 新建表，填充表头，填充列头，样式

                if (rowIndex == 65535 || rowIndex == 0)
                {
                    if (rowIndex != 0)
                    {
                        sheet = workbook.CreateSheet();
                    }

                    #region 表头及样式
                    int iTmp = 0;
                    {
                        if (!string.IsNullOrWhiteSpace(strHeaderText))
                        {
                            IRow headerRow = sheet.CreateRow(iTmp);
                            headerRow.HeightInPoints = 25;
                            headerRow.CreateCell(0).SetCellValue(strHeaderText);

                            ICellStyle headStyle = workbook.CreateCellStyle();
                            headStyle.Alignment = HorizontalAlignment.Center;
                            headStyle.VerticalAlignment = VerticalAlignment.Center;
                            headStyle.VerticalAlignment = VerticalAlignment.Center;
                            headStyle.BorderBottom = BorderStyle.Thin;
                            headStyle.BorderLeft = BorderStyle.Thin;
                            headStyle.BorderRight = BorderStyle.Thin;
                            headStyle.BorderTop = BorderStyle.Thin;

                            IFont font = workbook.CreateFont();
                            font.FontName = "仿宋";
                            font.FontHeightInPoints = 16;
                            font.Boldweight = 700;
                            headStyle.SetFont(font);

                            headerRow.GetCell(0).CellStyle = headStyle;
                            var vra = new CellRangeAddress(0, 0, 0, dtSource.Columns.Count - 1);
                            sheet.AddMergedRegion(vra);
                            iTmp++;
                        }
                    }

                    #endregion 表头及样式

                    #region 列头及样式

                    {
                        IRow headerRow = sheet.CreateRow(iTmp);
                        headerRow.HeightInPoints = (float)21.75;
                        ICellStyle headStyle = workbook.CreateCellStyle();
                        headStyle.Alignment = HorizontalAlignment.Center; // CellHorizontalAlignment.Center;
                        headStyle.VerticalAlignment = VerticalAlignment.Center;
                        headStyle.BorderBottom = BorderStyle.Thin;
                        headStyle.BorderLeft = BorderStyle.Thin;
                        headStyle.BorderRight = BorderStyle.Thin;
                        headStyle.BorderTop = BorderStyle.Thin;

                        IFont font = workbook.CreateFont();
                        font.FontHeightInPoints = 12;
                        font.Boldweight = 700;
                        font.FontName = "仿宋";
                        headStyle.SetFont(font);

                        foreach (DataColumn column in dtSource.Columns)
                        {
                            headerRow.CreateCell(column.Ordinal).SetCellValue(column.ColumnName);
                            headerRow.GetCell(column.Ordinal).CellStyle = headStyle;

                            //设置列宽
                            sheet.SetColumnWidth(column.Ordinal, (arrColWidth[column.Ordinal] + 5) * 256);
                        }
                        iTmp++;
                    }

                    #endregion 列头及样式

                    rowIndex = iTmp;
                }

                #endregion 新建表，填充表头，填充列头，样式

                #region 填充内容

                IRow dataRow = sheet.CreateRow(rowIndex);
                dataRow.HeightInPoints = (float)21.75;
                foreach (DataColumn column in dtSource.Columns)
                {
                    ICell newCell = dataRow.CreateCell(column.Ordinal);
                    //ICellStyle cellStyle = workbook.CreateCellStyle();

                    //cellStyle.VerticalAlignment = VerticalAlignment.Center;
                    //cellStyle.Alignment = HorizontalAlignment.LEFT;
                    //cellStyle.BorderBottom = BorderStyle.Thin;
                    //cellStyle.BorderLeft = BorderStyle.Thin;
                    //cellStyle.BorderRight = BorderStyle.Thin;
                    //cellStyle.BorderTop = BorderStyle.Thin;

                    newCell.CellStyle = cellStyle;

                    string drValue = row[column].ToString();
                    string type = column.DataType.ToString();
                    DateTime time;
                    double d;
                    if (DateTime.TryParse(drValue, out time)) type = "System.DateTime";
                    if (double.TryParse(drValue, out d)) type = "System.Decimal";
                    switch (type)
                    {
                        case "System.String": //字符串类型
                            newCell.SetCellValue(drValue);
                            break;

                        case "System.DateTime": //日期类型
                            if (drValue == "")
                            {
                                newCell.SetCellValue(drValue);
                            }
                            else
                            {
                                DateTime dateV;
                                DateTime.TryParse(drValue, out dateV);
                                newCell.SetCellValue(dateV);
                                newCell.CellStyle = dateStyle; //格式化显示
                            }

                            break;

                        case "System.Boolean": //布尔型
                            bool boolV = false;
                            bool.TryParse(drValue, out boolV);
                            newCell.SetCellValue(boolV);
                            cellStyle.Alignment = HorizontalAlignment.Center;

                            break;

                        case "System.Int16": //整型
                        case "System.Int32":
                        case "System.Byte":
                            int intV = 0;
                            int.TryParse(drValue, out intV);
                            newCell.SetCellValue(intV);
                            cellStyle.Alignment = HorizontalAlignment.Center;

                            break;

                        case "System.Int64":
                            int intV2 = 0;
                            int.TryParse(drValue, out intV2);
                            newCell.SetCellValue(intV2);
                            cellStyle.Alignment = HorizontalAlignment.Center;
                            break;

                        case "System.Decimal": //浮点型
                        case "System.Double":
                            double doubV = 0;
                            double.TryParse(drValue, out doubV);
                            newCell.SetCellValue(doubV);
                            cellStyle.Alignment = HorizontalAlignment.Center;

                            break;

                        case "System.DBNull": //空值处理
                            newCell.SetCellValue("");

                            break;

                        default:
                            newCell.SetCellValue("");
                            break;
                    }
                }

                #endregion 填充内容

                rowIndex++;
            }
            //合并单元格、冻结及重新设置列宽
            if (cellRanges != null)
            {
                int maxRow = 0;
                List<int> columnWidths = new List<int>();
                foreach (CellRangeAddress r in cellRanges)
                {
                    sheet.AddMergedRegion(r);
                    if (r.LastRow > maxRow) maxRow = r.LastRow;
                }
                //设置列宽
                foreach (DataColumn column in dtSource.Columns)
                {
                    int columnWidth = 0;
                    for (int j = maxRow; j < dtSource.Rows.Count; j++)
                    {
                        DataRow row = dtSource.Rows[j];
                        if (columnWidth < Encoding.GetEncoding(936).GetBytes(row[column] + string.Empty).Length) columnWidth = Encoding.GetEncoding(936).GetBytes(row[column] + string.Empty).Length;
                        //sheet.SetColumnWidth(column.Ordinal, (columnWidth + 5) * 256);
                    }
                    columnWidths.Add(columnWidth);
                }
                foreach (CellRangeAddress r in cellRanges)
                {
                    int startCel = r.FirstColumn;
                    int endCel = r.LastColumn;
                    int childWidth = 0;
                    for (int i = startCel; i <= endCel; i++)
                    {
                        childWidth += columnWidths[i];
                    }
                    if (childWidth < arrColWidth[startCel])
                    {
                        int differWidth = 0;
                        if ((endCel - startCel) > 0)
                        {
                            differWidth = ((arrColWidth[startCel] - childWidth) / (endCel - startCel)) + 1;
                        }
                        else
                        {
                            differWidth = 15;
                        }
                        for (int i = startCel; i <= endCel; i++)
                        {
                            columnWidths[i] += differWidth;
                        }
                    }
                }
                sheet.CreateFreezePane(0, maxRow + 1);                  //冻结行
                foreach (DataColumn column in dtSource.Columns)
                {
                    sheet.SetColumnWidth(column.Ordinal, (columnWidths[column.Ordinal] == 0 ? 15 : (columnWidths[column.Ordinal] + 5)) * 256);
                }
            }
            using (var ms = new MemoryStream())
            {
                workbook.Write(ms);
                ms.Flush();
                ms.Position = 0;
                return ms;
            }
        }

        /// <summary>
        ///     DataTable导出到Excel的MemoryStream
        /// </summary>
        /// <param name="dtSource">源DataTable</param>
        /// <param name="strHeaderText">表头文本</param>
        protected static MemoryStream ExportBatch(DataSet st, string strHeaderText, bool InOneSheet, List<CellRangeAddress> rangeList)
        {
            if (InOneSheet)
            {
                return ExportInOneSheetBatch(st, strHeaderText, rangeList);
            }
            var workbook = new HSSFWorkbook();
            foreach (DataTable dtSource in st.Tables)
            {
                ISheet sheet = workbook.CreateSheet(dtSource.TableName);
                sheet.ForceFormulaRecalculation = true; //强制要求Excel在打开时重新计算的属性
                /*
                 * CreateFreezePane()
                 * 第一个参数表示要冻结的列数；
                 * 第二个参数表示要冻结的行数；
                 * 第三个参数表示右边区域可见的首列序号，从1开始计算；
                 * 第四个参数表示下边区域可见的首行序号，也是从1开始计算
                 */
                sheet.CreateFreezePane(0, 2, 0, 2); //冻结表头与列头
                sheet.DisplayGridlines = false;//显示/隐藏网格线

                #region 右击文件 属性信息

                {
                    DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
                    dsi.Company = "公司";
                    workbook.DocumentSummaryInformation = dsi;

                    SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
                    si.Author = "作者"; //填加xls文件作者信息
                    si.ApplicationName = "程序信息"; //填加xls文件创建程序信息
                    si.LastAuthor = "最后保存者"; //填加xls文件最后保存者信息
                    si.Comments = "说明信息"; //填加xls文件作者信息
                    si.Title = "标题信息"; //填加xls文件标题信息
                    si.Subject = "文件主题"; //填加文件主题信息
                    si.CreateDateTime = DateTime.Now;
                    workbook.SummaryInformation = si;
                }

                #endregion 右击文件 属性信息

                ICellStyle dateStyle = workbook.CreateCellStyle();
                dateStyle.Alignment = HorizontalAlignment.Center;
                IDataFormat format = workbook.CreateDataFormat();
                dateStyle.DataFormat = format.GetFormat("yyyy-mm-dd");
                dateStyle.BorderBottom = BorderStyle.Thin;
                dateStyle.BorderLeft = BorderStyle.Thin;
                dateStyle.BorderRight = BorderStyle.Thin;
                dateStyle.BorderTop = BorderStyle.Thin;
                //取得列宽
                var arrColWidth = new int[dtSource.Columns.Count];
                foreach (DataColumn item in dtSource.Columns)
                {
                    arrColWidth[item.Ordinal] = Encoding.GetEncoding(936).GetBytes(item.ColumnName).Length;
                }
                for (int i = 0; i < dtSource.Rows.Count; i++)
                {
                    for (int j = 0; j < dtSource.Columns.Count; j++)
                    {
                        int intTemp = Encoding.GetEncoding(936).GetBytes(dtSource.Rows[i][j].ToString()).Length;
                        if (intTemp > arrColWidth[j])
                        {
                            if (intTemp > 60)
                            {
                                arrColWidth[j] = 60;
                            }
                            else
                            {
                                arrColWidth[j] = intTemp;
                            }
                        }
                    }
                }

                int rowIndex = 0;
                ICellStyle cellStyle = workbook.CreateCellStyle();

                cellStyle.VerticalAlignment = VerticalAlignment.Center;
                cellStyle.Alignment = HorizontalAlignment.Left;
                cellStyle.BorderBottom = BorderStyle.Thin;
                cellStyle.BorderLeft = BorderStyle.Thin;
                cellStyle.BorderRight = BorderStyle.Thin;
                cellStyle.BorderTop = BorderStyle.Thin;
                cellStyle.WrapText = true;

                foreach (DataRow row in dtSource.Rows)
                {
                    #region 新建表，填充表头，填充列头，样式

                    if (rowIndex == 65535 || rowIndex == 0)
                    {
                        if (rowIndex != 0)
                        {
                            sheet = workbook.CreateSheet(dtSource.TableName);
                        }

                        #region 表头及样式
                        int iTmp = 0;
                        {
                            if (!string.IsNullOrWhiteSpace(strHeaderText))
                            {
                                IRow headerRow = sheet.CreateRow(iTmp);
                                headerRow.HeightInPoints = 25;
                                headerRow.CreateCell(0).SetCellValue(strHeaderText);

                                ICellStyle headStyle = workbook.CreateCellStyle();
                                headStyle.Alignment = HorizontalAlignment.Center;
                                headStyle.VerticalAlignment = VerticalAlignment.Center;
                                headStyle.VerticalAlignment = VerticalAlignment.Center;
                                headStyle.BorderBottom = BorderStyle.Thin;
                                headStyle.BorderLeft = BorderStyle.Thin;
                                headStyle.BorderRight = BorderStyle.Thin;
                                headStyle.BorderTop = BorderStyle.Thin;

                                IFont font = workbook.CreateFont();
                                font.FontName = "仿宋";
                                font.FontHeightInPoints = 16;
                                font.Boldweight = 700;
                                headStyle.SetFont(font);

                                headerRow.GetCell(0).CellStyle = headStyle;
                                var vra = new CellRangeAddress(0, 0, 0, dtSource.Columns.Count - 1);
                                sheet.AddMergedRegion(vra);
                                iTmp++;
                            }
                        }

                        #endregion 表头及样式

                        #region 列头及样式

                        {
                            IRow headerRow = sheet.CreateRow(iTmp);
                            headerRow.HeightInPoints = (float)21.75;
                            ICellStyle headStyle = workbook.CreateCellStyle();
                            headStyle.Alignment = HorizontalAlignment.Center; // CellHorizontalAlignment.Center;
                            headStyle.VerticalAlignment = VerticalAlignment.Center;
                            headStyle.BorderBottom = BorderStyle.Thin;
                            headStyle.BorderLeft = BorderStyle.Thin;
                            headStyle.BorderRight = BorderStyle.Thin;
                            headStyle.BorderTop = BorderStyle.Thin;

                            IFont font = workbook.CreateFont();
                            font.FontHeightInPoints = 12;
                            font.Boldweight = 700;
                            font.FontName = "仿宋";
                            headStyle.SetFont(font);

                            foreach (DataColumn column in dtSource.Columns)
                            {
                                headerRow.CreateCell(column.Ordinal).SetCellValue(column.ColumnName);
                                headerRow.GetCell(column.Ordinal).CellStyle = headStyle;

                                //设置列宽
                                sheet.SetColumnWidth(column.Ordinal, (arrColWidth[column.Ordinal] + 5) * 256);
                            }
                            iTmp++;
                        }

                        #endregion 列头及样式

                        rowIndex = iTmp;
                    }

                    #endregion 新建表，填充表头，填充列头，样式

                    #region 填充内容

                    IRow dataRow = sheet.CreateRow(rowIndex);
                    dataRow.HeightInPoints = (float)21.75;
                    foreach (DataColumn column in dtSource.Columns)
                    {
                        ICell newCell = dataRow.CreateCell(column.Ordinal);
                        //ICellStyle cellStyle = workbook.CreateCellStyle();

                        //cellStyle.VerticalAlignment = VerticalAlignment.Center;
                        //cellStyle.Alignment = HorizontalAlignment.LEFT;
                        //cellStyle.BorderBottom = BorderStyle.Thin;
                        //cellStyle.BorderLeft = BorderStyle.Thin;
                        //cellStyle.BorderRight = BorderStyle.Thin;
                        //cellStyle.BorderTop = BorderStyle.Thin;

                        newCell.CellStyle = cellStyle;

                        string drValue = row[column].ToString();
                        string type = column.DataType.ToString();
                        DateTime time;
                        double d;
                        if (DateTime.TryParse(drValue, out time)) type = "System.DateTime";
                        if (double.TryParse(drValue, out d)) type = "System.Decimal";
                        switch (type)
                        {
                            case "System.String": //字符串类型
                                newCell.SetCellValue(drValue);
                                break;

                            case "System.DateTime": //日期类型
                                if (drValue == "")
                                {
                                    newCell.SetCellValue(drValue);
                                }
                                else
                                {
                                    DateTime dateV;
                                    DateTime.TryParse(drValue, out dateV);
                                    newCell.SetCellValue(dateV);
                                    newCell.CellStyle = dateStyle; //格式化显示
                                }

                                break;

                            case "System.Boolean": //布尔型
                                bool boolV = false;
                                bool.TryParse(drValue, out boolV);
                                newCell.SetCellValue(boolV);
                                cellStyle.Alignment = HorizontalAlignment.Center;

                                break;

                            case "System.Int16": //整型
                            case "System.Int32":
                            case "System.Byte":
                                int intV = 0;
                                int.TryParse(drValue, out intV);
                                newCell.SetCellValue(intV);
                                cellStyle.Alignment = HorizontalAlignment.Center;

                                break;

                            case "System.Int64":
                                int intV2 = 0;
                                int.TryParse(drValue, out intV2);
                                newCell.SetCellValue(intV2);
                                cellStyle.Alignment = HorizontalAlignment.Center;
                                break;

                            case "System.Decimal": //浮点型
                            case "System.Double":
                                double doubV = 0;
                                double.TryParse(drValue, out doubV);
                                newCell.SetCellValue(doubV);
                                cellStyle.Alignment = HorizontalAlignment.Center;

                                break;

                            case "System.DBNull": //空值处理
                                newCell.SetCellValue("");

                                break;

                            default:
                                newCell.SetCellValue("");
                                break;
                        }
                    }

                    #endregion 填充内容

                    rowIndex++;
                }
                if (rangeList != null)
                {
                    foreach (CellRangeAddress r in rangeList)
                    {
                        sheet.AddMergedRegion(r);
                    }
                }
            }
            using (var ms = new MemoryStream())
            {
                workbook.Write(ms);
                ms.Flush();
                ms.Position = 0;
                return ms;
            }
        }


        /// <summary>
        ///     DataTable导出到Excel的MemoryStream
        /// </summary>
        /// <param name="dtSource">源DataTable</param>
        /// <param name="strHeaderText">表头文本</param>
        protected static MemoryStream ExportInOneSheetBatch(DataSet st, string strHeaderText, List<CellRangeAddress> rangeList)
        {
            var workbook = new HSSFWorkbook();
            ISheet sheet = workbook.CreateSheet(st.Tables[0].TableName);
            sheet.ForceFormulaRecalculation = true; //强制要求Excel在打开时重新计算的属性
            /*
             * CreateFreezePane()
             * 第一个参数表示要冻结的列数；
             * 第二个参数表示要冻结的行数；
             * 第三个参数表示右边区域可见的首列序号，从1开始计算；
             * 第四个参数表示下边区域可见的首行序号，也是从1开始计算
             */
            //sheet.CreateFreezePane(0, 2, 0, 2); //冻结表头与列头
            sheet.DisplayGridlines = false;//显示/隐藏网格线

            #region 右击文件 属性信息

            {
                DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
                dsi.Company = "公司";
                workbook.DocumentSummaryInformation = dsi;

                SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
                si.Author = "作者"; //填加xls文件作者信息
                si.ApplicationName = "程序信息"; //填加xls文件创建程序信息
                si.LastAuthor = "最后保存者"; //填加xls文件最后保存者信息
                si.Comments = "说明信息"; //填加xls文件作者信息
                si.Title = "标题信息"; //填加xls文件标题信息
                si.Subject = "文件主题"; //填加文件主题信息
                si.CreateDateTime = DateTime.Now;
                workbook.SummaryInformation = si;
            }

            #endregion 右击文件 属性信息

            ICellStyle dateStyle = workbook.CreateCellStyle();
            dateStyle.Alignment = HorizontalAlignment.Center;
            IDataFormat format = workbook.CreateDataFormat();
            dateStyle.DataFormat = format.GetFormat("yyyy-mm-dd");
            dateStyle.BorderBottom = BorderStyle.Thin;
            dateStyle.BorderLeft = BorderStyle.Thin;
            dateStyle.BorderRight = BorderStyle.Thin;
            dateStyle.BorderTop = BorderStyle.Thin;


            int rowIndex = 0;
            int iTmp = 0;
            foreach (DataTable dtSource in st.Tables)
            {
                //取得列宽
                var arrColWidth = new int[dtSource.Columns.Count];
                foreach (DataColumn item in dtSource.Columns)
                {
                    arrColWidth[item.Ordinal] = Encoding.GetEncoding(936).GetBytes(item.ColumnName).Length;
                }
                for (int i = 0; i < dtSource.Rows.Count; i++)
                {
                    for (int j = 0; j < dtSource.Columns.Count; j++)
                    {
                        int intTemp = Encoding.GetEncoding(936).GetBytes(dtSource.Rows[i][j].ToString()).Length;
                        if (intTemp > arrColWidth[j])
                        {
                            if (intTemp > 60)
                            {
                                arrColWidth[j] = 60;
                            }
                            else
                            {
                                arrColWidth[j] = intTemp;
                            }
                        }
                    }
                }


                ICellStyle cellStyle = workbook.CreateCellStyle();

                cellStyle.VerticalAlignment = VerticalAlignment.Center;
                cellStyle.Alignment = HorizontalAlignment.Left;
                cellStyle.BorderBottom = BorderStyle.Thin;
                cellStyle.BorderLeft = BorderStyle.Thin;
                cellStyle.BorderRight = BorderStyle.Thin;
                cellStyle.BorderTop = BorderStyle.Thin;
                cellStyle.WrapText = true;
                if (iTmp != 0) iTmp++;

                #region 表头及样式
                {
                    if (!string.IsNullOrWhiteSpace(strHeaderText))
                    {
                        IRow headerRow = sheet.CreateRow(iTmp);
                        headerRow.HeightInPoints = 25;
                        headerRow.CreateCell(0).SetCellValue(strHeaderText);

                        ICellStyle headStyle = workbook.CreateCellStyle();
                        headStyle.Alignment = HorizontalAlignment.Center;
                        headStyle.VerticalAlignment = VerticalAlignment.Center;
                        headStyle.VerticalAlignment = VerticalAlignment.Center;
                        headStyle.BorderBottom = BorderStyle.Thin;
                        headStyle.BorderLeft = BorderStyle.Thin;
                        headStyle.BorderRight = BorderStyle.Thin;
                        headStyle.BorderTop = BorderStyle.Thin;

                        IFont font = workbook.CreateFont();
                        font.FontName = "仿宋";
                        font.FontHeightInPoints = 16;
                        font.Boldweight = 700;
                        headStyle.SetFont(font);

                        headerRow.GetCell(0).CellStyle = headStyle;
                        var vra = new CellRangeAddress(iTmp, iTmp, 0, dtSource.Columns.Count - 1);
                        sheet.AddMergedRegion(vra);
                        iTmp++;
                    }
                }

                #endregion 表头及样式

                #region 列头及样式

                {
                    IRow headerRow = sheet.CreateRow(iTmp);
                    headerRow.HeightInPoints = (float)21.75;
                    ICellStyle headStyle = workbook.CreateCellStyle();
                    headStyle.Alignment = HorizontalAlignment.Center; // CellHorizontalAlignment.Center;
                    headStyle.VerticalAlignment = VerticalAlignment.Center;
                    headStyle.BorderBottom = BorderStyle.Thin;
                    headStyle.BorderLeft = BorderStyle.Thin;
                    headStyle.BorderRight = BorderStyle.Thin;
                    headStyle.BorderTop = BorderStyle.Thin;

                    IFont font = workbook.CreateFont();
                    font.FontHeightInPoints = 12;
                    font.Boldweight = 700;
                    font.FontName = "仿宋";
                    headStyle.SetFont(font);

                    foreach (DataColumn column in dtSource.Columns)
                    {
                        headerRow.CreateCell(column.Ordinal).SetCellValue(column.ColumnName);
                        headerRow.GetCell(column.Ordinal).CellStyle = headStyle;

                        //设置列宽
                        sheet.SetColumnWidth(column.Ordinal, (arrColWidth[column.Ordinal] + 5) * 256);
                    }
                    iTmp++;
                }

                #endregion 列头及样式

                foreach (DataRow row in dtSource.Rows)
                {
                    #region 新建表，填充表头，填充列头，样式

                    //if (rowIndex == 0)
                    //{
                    //    sheet = workbook.CreateSheet(dtSource.TableName);
                    //}

                    rowIndex = iTmp;

                    #endregion 新建表，填充表头，填充列头，样式

                    #region 填充内容

                    IRow dataRow = sheet.CreateRow(rowIndex);
                    dataRow.HeightInPoints = (float)21.75;
                    foreach (DataColumn column in dtSource.Columns)
                    {
                        ICell newCell = dataRow.CreateCell(column.Ordinal);
                        //ICellStyle cellStyle = workbook.CreateCellStyle();

                        //cellStyle.VerticalAlignment = VerticalAlignment.Center;
                        //cellStyle.Alignment = HorizontalAlignment.LEFT;
                        //cellStyle.BorderBottom = BorderStyle.Thin;
                        //cellStyle.BorderLeft = BorderStyle.Thin;
                        //cellStyle.BorderRight = BorderStyle.Thin;
                        //cellStyle.BorderTop = BorderStyle.Thin;

                        newCell.CellStyle = cellStyle;

                        string drValue = row[column].ToString();
                        string type = column.DataType.ToString();
                        DateTime time;
                        double d;
                        if (DateTime.TryParse(drValue, out time)) type = "System.DateTime";
                        if (double.TryParse(drValue, out d)) type = "System.Decimal";
                        switch (type)
                        {
                            case "System.String": //字符串类型
                                newCell.SetCellValue(drValue);
                                break;

                            case "System.DateTime": //日期类型
                                if (drValue == "")
                                {
                                    newCell.SetCellValue(drValue);
                                }
                                else
                                {
                                    DateTime dateV;
                                    DateTime.TryParse(drValue, out dateV);
                                    newCell.SetCellValue(dateV);
                                    newCell.CellStyle = dateStyle; //格式化显示
                                }

                                break;

                            case "System.Boolean": //布尔型
                                bool boolV = false;
                                bool.TryParse(drValue, out boolV);
                                newCell.SetCellValue(boolV);
                                cellStyle.Alignment = HorizontalAlignment.Center;

                                break;

                            case "System.Int16": //整型
                            case "System.Int32":
                            case "System.Byte":
                                int intV = 0;
                                int.TryParse(drValue, out intV);
                                newCell.SetCellValue(intV);
                                cellStyle.Alignment = HorizontalAlignment.Center;

                                break;

                            case "System.Int64":
                                int intV2 = 0;
                                int.TryParse(drValue, out intV2);
                                newCell.SetCellValue(intV2);
                                cellStyle.Alignment = HorizontalAlignment.Center;
                                break;

                            case "System.Decimal": //浮点型
                            case "System.Double":
                                double doubV = 0;
                                double.TryParse(drValue, out doubV);
                                newCell.SetCellValue(doubV);
                                cellStyle.Alignment = HorizontalAlignment.Center;

                                break;

                            case "System.DBNull": //空值处理
                                newCell.SetCellValue("");

                                break;

                            default:
                                newCell.SetCellValue("");
                                break;
                        }
                    }
                    iTmp++; rowIndex++;
                    #endregion 填充内容
                }
                if (rangeList != null)
                {
                    foreach (CellRangeAddress r in rangeList)
                    {
                        sheet.AddMergedRegion(r);
                    }
                }
            }
            using (var ms = new MemoryStream())
            {
                workbook.Write(ms);
                ms.Flush();
                ms.Position = 0;
                return ms;
            }
        }


        /// <summary>
        /// 导出EXCEL
        /// </summary>
        /// <param name="sTableNames">所有表名</param>
        /// <param name="strFileName">导出文件名</param>
        public static void ExportExcel(DataSet ds, string strFilePath)
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            for (int iTCount = 0; iTCount < ds.Tables.Count; iTCount++)
            {
                DataTable dt = ds.Tables[iTCount];
                XSSFSheet sheet = (XSSFSheet)workbook.CreateSheet(dt.TableName);

                //填充表头
                XSSFRow dataRow = (XSSFRow)sheet.CreateRow(0);
                foreach (DataColumn column in dt.Columns)
                {
                    ICell rowcell = dataRow.CreateCell(column.Ordinal);
                    rowcell.SetCellValue(column.ColumnName);

                }

                //填充内容
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    dataRow = (XSSFRow)sheet.CreateRow(i + 1);
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        ICell rowcell = dataRow.CreateCell(j);
                        rowcell.SetCellValue(dt.Rows[i][j].ToString());
                    }
                }
            }

            //保存
            using (MemoryStream ms = new MemoryStream())
            {
                using (FileStream fs = new FileStream(strFilePath, FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(fs);
                }
            }

        }

        /// <summary>
        /// 导出Excel,导入到模板中
        /// </summary>
        /// <param name="ds"></param>
        /// <param name="readUrlStr">模板文件地址</param>
        /// <param name="writeUrlStr">保存文件地址</param>
        public static void ExportExcelFirstCreateThenWrite(DataSet ds, string readUrlStr, string writeUrlStr)
        {
            if (File.Exists(writeUrlStr))
            {
                File.Delete(writeUrlStr);
            }
            using (FileStream stream = new FileStream(readUrlStr, FileMode.Open, FileAccess.Read))
            {
                XSSFWorkbook workbook = new XSSFWorkbook(stream);
                ICellStyle cellstyle = (XSSFCellStyle)workbook.CreateCellStyle();
                //创建excel
                for (int iTCount = 0; iTCount < ds.Tables.Count; iTCount++)
                {
                    DataTable dt = ds.Tables[iTCount];
                    XSSFSheet sheet = (XSSFSheet)workbook.GetSheet(dt.TableName);

                    //填充表头
                    XSSFRow dataRow = (XSSFRow)sheet.GetRow(0);
                    foreach (DataColumn column in dt.Columns)
                    {
                        ICell cell = dataRow.GetCell(column.Ordinal);
                        cell.SetCellValue(column.ColumnName);
                    }

                    //填充内容
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        dataRow = (XSSFRow)sheet.GetRow(i + 1);
                        for (int j = 0; j < dt.Columns.Count; j++)
                        {
                            ICell cell = dataRow.GetCell(j);
                            cell.SetCellValue(dt.Rows[i][j] == null ? "" : dt.Rows[i][j].ToString());
                        }
                    }
                    //sheet.ForceFormulaRecalculation = true;
                }

                using (FileStream stream1 = new FileStream(writeUrlStr, FileMode.OpenOrCreate, FileAccess.Write))
                {
                    //保存
                    workbook.Write(stream1);
                }
            }

        }

        /// <summary>
        /// 把Sheet中的数据转换为DataTable
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns></returns>
        private DataTable ExportToDataTable(ISheet sheet)
        {
            DataTable dt = new DataTable();

            //默认，第一行是字段
            IRow headRow = sheet.GetRow(0);

            //设置datatable字段
            for (int i = headRow.FirstCellNum, len = headRow.LastCellNum; i < len; i++)
            {
                dt.Columns.Add(headRow.Cells[i].StringCellValue);
            }
            //遍历数据行
            for (int i = (sheet.FirstRowNum + 1), len = sheet.LastRowNum + 1; i < len; i++)
            {
                IRow tempRow = sheet.GetRow(i);
                DataRow dataRow = dt.NewRow();

                //遍历一行的每一个单元格
                for (int r = 0, j = tempRow.FirstCellNum, len2 = tempRow.LastCellNum; j < len2; j++, r++)
                {

                    ICell cell = tempRow.GetCell(j);

                    if (cell != null)
                    {
                        switch (cell.CellType)
                        {
                            case CellType.String:
                                dataRow[r] = cell.StringCellValue;
                                break;
                            case CellType.Numeric:
                                dataRow[r] = cell.NumericCellValue;
                                break;
                            case CellType.Boolean:
                                dataRow[r] = cell.BooleanCellValue;
                                break;
                            default: dataRow[r] = "ERROR";
                                break;
                        }
                    }
                }
                dt.Rows.Add(dataRow);
            }
            return dt;
        }

        /// <summary>
        /// Excel文件导成Datatable
        /// </summary>
        /// <param name="strFilePath">Excel文件目录地址</param>
        /// <param name="strTableName">Datatable表名</param>
        /// <param name="iSheetIndex">Excel sheet index</param>
        /// <returns></returns>
        public static DataTable XlSToDataTable(string strFilePath, string strTableName, int iSheetIndex)
        {

            string strExtName = Path.GetExtension(strFilePath);

            DataTable dt = new DataTable();
            if (!string.IsNullOrEmpty(strTableName))
            {
                dt.TableName = strTableName;
            }

            if (strExtName.Equals(".xls") || strExtName.Equals(".xlsx"))
            {
                using (FileStream file = new FileStream(strFilePath, FileMode.Open, FileAccess.Read))
                {
                    IWorkbook workbook;
                    if (strExtName.Equals(".xls")) workbook = new HSSFWorkbook(file);
                    else workbook = new XSSFWorkbook(file);
                    ISheet sheet = workbook.GetSheetAt(iSheetIndex);

                    //列头
                    foreach (ICell item in sheet.GetRow(sheet.FirstRowNum).Cells)
                    {
                        dt.Columns.Add(item.ToString(), typeof(string));
                    }

                    //写入内容
                    System.Collections.IEnumerator rows = sheet.GetRowEnumerator();
                    while (rows.MoveNext())
                    {
                        IRow row;
                        if(strExtName.Equals(".xls"))  row=(HSSFRow)rows.Current;
                        else  row = (XSSFRow)rows.Current;
                        if (row.RowNum == sheet.FirstRowNum)
                        {
                            continue;
                        }

                        DataRow dr = dt.NewRow();
                        foreach (ICell item in row.Cells)
                        {
                            switch (item.CellType)
                            {
                                case CellType.Boolean:
                                    dr[item.ColumnIndex] = item.BooleanCellValue;
                                    break;
                                case CellType.Error:
                                    dr[item.ColumnIndex] = ErrorEval.GetText(item.ErrorCellValue);
                                    break;
                                case CellType.Formula:
                                    switch (item.CachedFormulaResultType)
                                    {
                                        case CellType.Boolean:
                                            dr[item.ColumnIndex] = item.BooleanCellValue;
                                            break;
                                        case CellType.Error:
                                            dr[item.ColumnIndex] = ErrorEval.GetText(item.ErrorCellValue);
                                            break;
                                        case CellType.Numeric:
                                            if (DateUtil.IsCellDateFormatted(item))
                                            {
                                                dr[item.ColumnIndex] = item.DateCellValue.ToString("yyyy-MM-dd hh:MM:ss");
                                            }
                                            else
                                            {
                                                dr[item.ColumnIndex] = item.NumericCellValue;
                                            }
                                            break;
                                        case CellType.String:
                                            string str = item.StringCellValue;
                                            if (!string.IsNullOrEmpty(str))
                                            {
                                                dr[item.ColumnIndex] = str.ToString();
                                            }
                                            else
                                            {
                                                dr[item.ColumnIndex] = "";
                                            }
                                            break;
                                        case CellType.Unknown:
                                        case CellType.Blank:
                                        default:
                                            dr[item.ColumnIndex] = string.Empty;
                                            break;
                                    }
                                    break;
                                case CellType.Numeric:
                                    if (DateUtil.IsCellDateFormatted(item))
                                    {
                                        dr[item.ColumnIndex] = item.DateCellValue.ToString("yyyy-MM-dd hh:MM:ss");
                                    }
                                    else
                                    {
                                        dr[item.ColumnIndex] = item.NumericCellValue;
                                    }
                                    break;
                                case CellType.String:
                                    string strValue = item.StringCellValue;
                                    if (!string.IsNullOrEmpty(strValue))
                                    {
                                        dr[item.ColumnIndex] = strValue.ToString();
                                    }
                                    else
                                    {
                                        dr[item.ColumnIndex] = "";
                                    }
                                    break;
                                case CellType.Unknown:
                                case CellType.Blank:
                                default:
                                    dr[item.ColumnIndex] = string.Empty;
                                    break;
                            }
                        }
                        dt.Rows.Add(dr);
                    }
                }
            }

            return dt;
        }

        /// <summary>
        /// Excel某sheet中内容导入到DataTable中
        /// 区分xsl和xslx分别处理
        /// </summary>
        /// <param name="filePath">Excel文件路径,含文件全名</param>
        /// <param name="sheetName">此Excel中sheet名</param>
        /// <returns></returns>
        public static DataTable ExcelSheetImportToDataTable(string filePath, string sheetName)
        {

            DataTable dt = new DataTable();
            if (Path.GetExtension(filePath).ToLower() == ".xls".ToLower())
            {//.xls
                #region .xls文件处理:HSSFWorkbook
                HSSFWorkbook hssfworkbook;
                try
                {
                    using (FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                    {

                        hssfworkbook = new HSSFWorkbook(file);
                    }
                }
                catch (Exception e)
                {
                    throw e;
                }

                ISheet sheet = hssfworkbook.GetSheet(sheetName);
                System.Collections.IEnumerator rows = sheet.GetRowEnumerator();
                HSSFRow headerRow = (HSSFRow)sheet.GetRow(0);

                //一行最后一个方格的编号 即总的列数  
                for (int j = 0; j < (sheet.GetRow(0).LastCellNum); j++)
                {
                    //SET EVERY COLUMN NAME
                    HSSFCell cell = (HSSFCell)headerRow.GetCell(j);

                    dt.Columns.Add(cell.ToString());
                }

                while (rows.MoveNext())
                {
                    IRow row = (HSSFRow)rows.Current;
                    DataRow dr = dt.NewRow();

                    if (row.RowNum == 0) continue;//The firt row is title,no need import

                    for (int i = 0; i < row.LastCellNum; i++)
                    {
                        if (i >= dt.Columns.Count)//cell count>column count,then break //每条记录的单元格数量不能大于表格栏位数量 20140213
                        {
                            break;
                        }

                        ICell cell = row.GetCell(i);

                        if ((i == 0) && (string.IsNullOrEmpty(cell.ToString()) == true))//每行第一个cell为空,break
                        {
                            break;
                        }

                        if (cell == null)
                        {
                            dr[i] = null;
                        }
                        else
                        {
                            dr[i] = cell.ToString();
                        }
                    }

                    dt.Rows.Add(dr);
                }
                #endregion
            }
            else
            {//.xlsx
                #region .xlsx文件处理:XSSFWorkbook
                XSSFWorkbook hssfworkbook;
                try
                {
                    using (FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                    {

                        hssfworkbook = new XSSFWorkbook(file);
                    }
                }
                catch (Exception e)
                {
                    throw e;
                }

                ISheet sheet = hssfworkbook.GetSheet(sheetName);
                System.Collections.IEnumerator rows = sheet.GetRowEnumerator();
                XSSFRow headerRow = (XSSFRow)sheet.GetRow(0);



                //一行最后一个方格的编号 即总的列数  
                for (int j = 0; j < (sheet.GetRow(0).LastCellNum); j++)
                {
                    //SET EVERY COLUMN NAME
                    XSSFCell cell = (XSSFCell)headerRow.GetCell(j);

                    dt.Columns.Add(cell.ToString());

                }

                while (rows.MoveNext())
                {
                    IRow row = (XSSFRow)rows.Current;
                    DataRow dr = dt.NewRow();

                    if (row.RowNum == 0) continue;//The firt row is title,no need import

                    for (int i = 0; i < row.LastCellNum; i++)
                    {
                        if (i >= dt.Columns.Count)//cell count>column count,then break //每条记录的单元格数量不能大于表格栏位数量 20140213
                        {
                            break;
                        }

                        ICell cell = row.GetCell(i);

                        if ((i == 0) && (string.IsNullOrEmpty(cell.ToString()) == true))//每行第一个cell为空,break
                        {
                            break;
                        }

                        if (cell == null)
                        {
                            dr[i] = null;
                        }
                        else
                        {
                            dr[i] = cell.ToString();
                        }
                    }
                    dt.Rows.Add(dr);
                }
                #endregion
            }
            return dt;
        }

    }
}