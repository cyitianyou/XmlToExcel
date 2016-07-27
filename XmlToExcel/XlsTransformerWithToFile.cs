using BTulz.ModelsTransformer.DomainModels;
using BTulz.ModelsTransformer.ExcelMaping;
using BTulz.ModelsTransformer.Transformer;

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
        //是否使用微软Office组件,如果否则使用NPOI
        public bool isUseMicrosoftOffice = true;

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
            return ExcelHelper.ImportToExcel(ds, filePath, isUseMicrosoftOffice);
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

        
    }
}
