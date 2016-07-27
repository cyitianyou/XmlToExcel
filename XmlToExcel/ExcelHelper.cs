using Microsoft.Office.Interop.Excel;
using System;
using System.Data;
using System.IO;

namespace XmlToExcel
{
    public class ExcelHelper
    {
        public static string ImportToExcel(DataSet ds, string strFilenamePath,bool flag)
        {
            if (flag)
                return Microsoft_Office_ImportToExcel(ds, strFilenamePath);
            else
                return NPOI_ImportToExcel(ds, strFilenamePath);
        }
        #region Microsoft.Office组件
        /// <summary>
        /// 将数据导入到Excel
        /// </summary>
        /// <param name="ds">
        /// 需要生成Excel的数据源
        /// DataSet->DataTable->TableName为页（Sheet）名字
        /// </param>
        /// <param name="strFilenamePath">生成后文件保存的全路径</param>
        /// <returns></returns>
        public static string Microsoft_Office_ImportToExcel(DataSet ds, string strFilenamePath)
        {

            if (ds.Tables.Count == 0) return "";
            try
            {
                ApplicationClass objApp = new ApplicationClass();
                _Workbook objWorkbook;//工作薄
                _Worksheet objWorksheet;//工作页
                objWorkbook = objApp.Workbooks.Add(true);
                object objMissing = System.Reflection.Missing.Value;
                #region 添加数据
                for (int i = 0; i < ds.Tables.Count; i++)
                {
                    objWorksheet = (_Worksheet)objWorkbook.ActiveSheet;
                    //书签名字为表名
                    objWorksheet.Name = ds.Tables[i].TableName;
                    //正文内容，从第二行开始
                    for (int rows = 0; rows < ds.Tables[i].Rows.Count; rows++)
                    {
                        for (int cols = 0; cols < ds.Tables[i].Columns.Count; cols++)
                        {
                            objApp.Cells[rows + 1, cols + 1] = ds.Tables[i].Rows[rows][cols].ToString();
                        }
                    }
                    //Borders.LineStyle 单元格边框线
                    Range excelRange = objWorksheet.get_Range(objWorksheet.Cells[1, 1], objWorksheet.Cells[ds.Tables[i].Rows.Count, 12]);
                    //单元格边框线类型(线型,虚线型)
                    excelRange.Borders.LineStyle = 1;
                    excelRange.Borders.LineStyle = XlLineStyle.xlContinuous;
                    //指定单元格下边框线粗细,和色彩
                    excelRange.Borders.Weight = XlBorderWeight.xlHairline;
                    excelRange.Borders.ColorIndex = 1;
                    //设置字体在单元格内的对其方式
                    excelRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    //设置单元格的宽度
                    //excelRange.ColumnWidth = 15;
                    //设置单元格的背景色
                    //excelRange.Cells.Interior.Color = System.Drawing.Color.FromArgb(255, 204, 153).ToArgb();
                    // 给单元格加边框
                    excelRange.BorderAround(XlLineStyle.xlDouble, XlBorderWeight.xlThick,
                                                              XlColorIndex.xlColorIndexAutomatic, System.Drawing.Color.FromArgb(0, 0, 0).ToArgb());
                    //自动调整列宽
                    excelRange.EntireColumn.AutoFit();
                    if (i < ds.Tables.Count - 1)
                    {
                        objApp.Sheets.Add(objMissing, objMissing, 1, XlSheetType.xlWorksheet);
                    }
                }
                #endregion
                //objApp.Visible = true;
                //将Excel保存到指定路径
                objWorkbook.SaveAs(
                    strFilenamePath, objMissing, objMissing, objMissing, objMissing,
                    objMissing, XlSaveAsAccessMode.xlNoChange, objMissing,
                    objMissing, objMissing, objMissing, objMissing);
                objWorkbook.Close();
                Close(objApp);
                return strFilenamePath;
            }
            catch (Exception ex)
            {
                string strEXMessage = ex.Message;
                return "";

            }

        }

        /// <summary>
        /// 关闭实例Excel.Application后产生的进程
        /// </summary>
        public static void Close(Microsoft.Office.Interop.Excel.ApplicationClass _xlApp)
        {
            if (_xlApp != null)
            {
                int generation = 0;
                _xlApp.UserControl = false;
                #region 如果您将 DisplayAlerts 属性设置为 False，则系统不会提示您保存任何未保存的数据。
                //如果您将 DisplayAlerts 属性设置为 False，则系统不会提示您保存任何未保存的数据。
                //_xlApp.DisplayAlerts = false;
                //if (_xlWorkbook != null)
                //{
                //    //如果将 Workbook 的 Saved 属性设置为 True，则不管您有没有进行更改，Excel 都不会提示保存它
                //    //_xlWorkbook.Saved = true;
                //    try
                //    {
                //        ////经过实验，这两句写不写都不会影响进程驻留。
                //        ////如果注释掉的话，即使用户手动从界面上关闭了本程序的Excel，也不会影响
                //        //_xlWorkbook.Close(oMissing,oMissing,oMissing);
                //        //_xlWorkbook = null;
                //    }
                //    catch
                //    {
                //        //用户手动从界面上关闭了本程序的Excel窗口
                //    }
                //}
                #endregion
                //即使用户手动从界面上关闭了，但是Excel.Exe进程仍然存在，用_xlApp.Quit()退出也不会出错，用垃圾回收彻底清除
                _xlApp.Quit();

                //System.Runtime.InteropServices.Marshal.ReleaseComObject((object)_xlApp);
                generation = System.GC.GetGeneration(_xlApp);
                _xlApp = null;

                //虽然用了_xlApp.Quit()，但由于是COM，并不能清除驻留在内存在的进程，每实例一次Excel则Excell进程多一个。
                //因此用垃圾回收，建议不要用进程的KILL()方法，否则可能会错杀无辜啊:)。
                System.GC.Collect(generation);
            }
        }
        #endregion
        #region NPOI
        /// <summary>
        /// 将数据导入到Excel
        /// </summary>
        /// <param name="ds">
        /// 需要生成Excel的数据源
        /// DataSet->DataTable->TableName为页（Sheet）名字
        /// </param>
        /// <param name="strFilenamePath">生成后文件保存的全路径</param>
        /// <returns></returns>
        public static string NPOI_ImportToExcel(DataSet ds, string strFilenamePath)
        {
            if (ds.Tables.Count == 0) return "";
            try
            {
                NPOI.HSSF.UserModel.HSSFWorkbook book = new NPOI.HSSF.UserModel.HSSFWorkbook();
                NPOI.SS.UserModel.ICellStyle cellStyle = book.CreateCellStyle();
                //设置单元格上下左右边框线
                cellStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Dotted;
                cellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Dotted;
                cellStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Dotted;
                cellStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Dotted;
                //文字水平和垂直对齐方式
                cellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                cellStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
                //是否换行
                //cellStyle.WrapText = true;
                //缩小字体填充
                //cellStyle.ShrinkToFit = true;
                //左
                NPOI.SS.UserModel.ICellStyle cellStyle1 = book.CreateCellStyle();
                cellStyle1.CloneStyleFrom(cellStyle);
                cellStyle1.BorderLeft = NPOI.SS.UserModel.BorderStyle.Double;
                //左上
                NPOI.SS.UserModel.ICellStyle cellStyle2 = book.CreateCellStyle();
                cellStyle2.CloneStyleFrom(cellStyle1);
                cellStyle2.BorderTop = NPOI.SS.UserModel.BorderStyle.Double;
                //上
                NPOI.SS.UserModel.ICellStyle cellStyle3 = book.CreateCellStyle();
                cellStyle3.CloneStyleFrom(cellStyle);
                cellStyle3.BorderTop = NPOI.SS.UserModel.BorderStyle.Double;
                //右上
                NPOI.SS.UserModel.ICellStyle cellStyle4 = book.CreateCellStyle();
                cellStyle4.CloneStyleFrom(cellStyle3);
                cellStyle4.BorderRight = NPOI.SS.UserModel.BorderStyle.Double;
                //右
                NPOI.SS.UserModel.ICellStyle cellStyle5 = book.CreateCellStyle();
                cellStyle5.CloneStyleFrom(cellStyle);
                cellStyle5.BorderRight = NPOI.SS.UserModel.BorderStyle.Double;
                //右下
                NPOI.SS.UserModel.ICellStyle cellStyle6 = book.CreateCellStyle();
                cellStyle6.CloneStyleFrom(cellStyle5);
                cellStyle6.BorderBottom = NPOI.SS.UserModel.BorderStyle.Double;
                //下
                NPOI.SS.UserModel.ICellStyle cellStyle7 = book.CreateCellStyle();
                cellStyle7.CloneStyleFrom(cellStyle);
                cellStyle7.BorderBottom = NPOI.SS.UserModel.BorderStyle.Double;
                //左下
                NPOI.SS.UserModel.ICellStyle cellStyle8 = book.CreateCellStyle();
                cellStyle8.CloneStyleFrom(cellStyle7);
                cellStyle8.BorderLeft = NPOI.SS.UserModel.BorderStyle.Double;
                #region 添加数据
                for (int i = 0; i < ds.Tables.Count; i++)
                {
                    var dt = ds.Tables[i];
                    NPOI.SS.UserModel.ISheet sheet = book.CreateSheet(dt.TableName);
                    NPOI.SS.UserModel.IRow row = sheet.CreateRow(0);
                    NPOI.SS.UserModel.ICell cell;
                    //正文内容，从第二行开始
                    for (int rows = 0; rows < ds.Tables[i].Rows.Count; rows++)
                    {
                        row = sheet.CreateRow(rows + 1);
                        for (int cols = 0; cols < ds.Tables[i].Columns.Count; cols++)
                        {
                            cell = row.CreateCell(cols);
                            cell.SetCellValue(Convert.ToString(dt.Rows[rows][cols]));
                            cell.CellStyle = cellStyle;
                            //上边框
                            if (rows == 0)
                                cell.CellStyle = cellStyle3;
                            //下边框
                            else if (rows == dt.Rows.Count - 1)
                                cell.CellStyle = cellStyle7;
                            if (cols == 0)
                            {
                                //左上边框
                                if (rows == 0)
                                {
                                    cell.CellStyle = cellStyle2;
                                }
                                //左下边框
                                else if (rows == dt.Rows.Count - 1)
                                {
                                    cell.CellStyle = cellStyle8;
                                }
                                //左边框
                                else
                                {
                                    cell.CellStyle = cellStyle1;
                                }
                            }
                            if (cols == 11)
                            {
                                //右上边框
                                if (rows == 0)
                                {
                                    cell.CellStyle = cellStyle4;
                                }
                                //右下边框
                                else if (rows == dt.Rows.Count - 1)
                                {
                                    cell.CellStyle = cellStyle6;
                                }
                                //右边框
                                else
                                {
                                    cell.CellStyle = cellStyle5;
                                }
                            }
                        }
                    }
                    //调整列宽度
                    for (int columnNum = 0; columnNum <= 11; columnNum++)
                    {
                        sheet.AutoSizeColumn(columnNum);
                        //int columnWidth = sheet.GetColumnWidth(columnNum) / 256;//获取当前列宽度  
                        //for (int rowNum = 1; rowNum <= sheet.LastRowNum; rowNum++)//在这一列上循环行  
                        //{
                        //    row = sheet.GetRow(rowNum);
                        //    cell = row.GetCell(columnNum);
                        //    if (cell == null)
                        //        continue;
                        //    int length = System.Text.Encoding.UTF8.GetBytes(cell.ToString()).Length;//获取当前单元格的内容宽度  
                        //    if (columnWidth < length + 1)
                        //    {
                        //        columnWidth = length + 1;
                        //    }//若当前单元格内容宽度大于列宽，则调整列宽为当前单元格宽度，后面的+1是我人为的将宽度增加一个字符  
                        //}
                        //sheet.SetColumnWidth(columnNum, columnWidth * 256);
                    }
                }
                #endregion
                // 写入到客户端  
                using (System.IO.MemoryStream ms = new System.IO.MemoryStream())
                {
                    book.Write(ms);
                    using (FileStream fs = new FileStream(strFilenamePath, FileMode.Create, FileAccess.Write))
                    {
                        byte[] data = ms.ToArray();
                        fs.Write(data, 0, data.Length);
                        fs.Flush();
                    }
                    book = null;
                }
                return strFilenamePath;
            }
            catch (Exception ex)
            {
                string strEXMessage = ex.Message;
                return "";

            }
        }
        #endregion
    }
}
