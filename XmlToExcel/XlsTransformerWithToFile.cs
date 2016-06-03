using BTulz.ModelsTransformer.DomainModels;
using BTulz.ModelsTransformer.ExcelMaping;
using BTulz.ModelsTransformer.Transformer;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XmlToExcel
{
    class XlsTransformerWithToFile : XlsTransformer
    {
        private string _folderPath;
        private string modelFileName;
        private string configName;
        public string folderPath
        {
            get
            {
                if (_folderPath == null)
                    _folderPath = "";
                return _folderPath;
            }
            set
            {
                _folderPath = value;
                modelFileName = Path.Combine(_folderPath, "datadictionary_Domain_v3.1.xls");
                configName = Path.Combine(_folderPath, "excelmappingEx.xml");
            }
        }
        private string SheetName = @"Model$";
        private System.Data.DataTable _modelTabel;
        private System.Data.DataTable modelTabel
        {
            get
            {
                if(_modelTabel==null)
                {
                    _modelTabel = LoadSheetData(modelFileName, SheetName);
                }
                return _modelTabel;
            }

        }

        private MyModelMapping _myMap;
        private MyModelMapping myMap
        {
            get
            {
                if (_myMap == null)
                {
                    _myMap = new MyModelMapping();
                    _myMap.GetMappingConfig(configName);
                }
                return _myMap;
            }

        }
        public override string ToFile(string outPutFolder)
        {
            DataSet ds = new DataSet();
            foreach (var item in this.DomainModel.BusinessObjects)
            {
                System.Data.DataTable table = ToDataTable(item);
                var tableName = item.Name;
                int i=1;
                while (ds.Tables.Contains(tableName))
                {
                    tableName = tableName + i.ToString();
                    i++;
                }
                table.TableName = tableName;
                ds.Tables.Add(table);
            }
            string filePath = Path.Combine(outPutFolder, string.Format("datadictionary_{0}_v1.1.xlsx", this.DomainModel.Name));
            return ImportToExcel(ds,filePath);
        }

        private System.Data.DataTable ToDataTable(IBusinessObject obj)
        {
            System.Data.DataTable table = modelTabel.Clone();
            myMap.Table = table;
            myMap.GetTableData(this.DomainModel.Tables.FirstOrDefault(c => c.Name == obj.TableName));
            foreach (var item in obj.ChildTables)
            {
                myMap.GetTableData(this.DomainModel.Tables.FirstOrDefault(c => c.Name == item.TableName));
            }
            myMap.GetBusinessObjectData(obj);
            
            return table;

        }

        /// <summary>
        /// 将数据导入到Excel
        /// </summary>
        /// <param name="ds">
        /// 需要生成Excel的数据源
        /// DataSet->DataTable->TableName为页（Sheet）名字
        /// </param>
        /// <param name="strFilenamePath">生成后文件保存的全路径</param>
        /// <returns></returns>
        public string ImportToExcel(DataSet ds, string strFilenamePath)
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
        public void Close(Microsoft.Office.Interop.Excel.ApplicationClass _xlApp)
        {
            if (_xlApp != null)
            {
                int generation = 0;
                _xlApp.UserControl = false;
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
    }
}
