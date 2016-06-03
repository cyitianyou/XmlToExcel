using BTulz.ModelsTransformer.DomainModels;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Xml;

namespace BTulz.ModelsTransformer.ExcelMaping
{
    public class MyModelMapping
    {
        #region 反编译代码
        private List<ObjectArea> _Arealist = new List<ObjectArea>();
        private ObjectArea _CurrentArea;
        private PropertyInfoManager _pInfoDic = new PropertyInfoManager();
        private int _RowOffset;
        private DataTable _Table;

        private ObjectAreaType GetAreaType(DataRow row)
        {
            foreach (ObjectArea area in this._Arealist)
            {
                if (area.CheckAreaType(row))
                {
                    this._CurrentArea = area;
                    return (ObjectAreaType)Enum.Parse(typeof(ObjectAreaType), area.Name);
                }
            }
            this._CurrentArea = null;
            return ObjectAreaType.NotDefind;
        }

        public void GetMappingConfig(string fileName)
        {
            if (!File.Exists(fileName))
            {
                throw new Exception(string.Format("file was not exists.[{0}]", fileName));
            }
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(fileName);
            this.GetObjectAreas(xmlDoc);
        }


        private void GetObjectAreas(XmlDocument xmlDoc)
        {
            foreach (System.Xml.XmlNode node in xmlDoc.GetElementsByTagName("ObjectArea"))
            {
                ObjectArea area = new ObjectArea();
                this._pInfoDic.GetProperties(area);
                this._pInfoDic.SetModelValue(area, node);
                foreach (System.Xml.XmlNode node2 in node.SelectSingleNode("Conditions"))
                {
                    Condition condition = new Condition();
                    this._pInfoDic.GetProperties(condition);
                    this._pInfoDic.SetModelValue(condition, node2);
                    area.Conditions.Add(condition);
                }
                foreach (System.Xml.XmlNode node3 in node.SelectSingleNode("Propertys"))
                {
                    Property property = new Property();
                    this._pInfoDic.GetProperties(property);
                    this._pInfoDic.SetModelValue(property, node3);
                    area.Propertys.Add(property);
                }
                foreach (System.Xml.XmlNode node4 in node.SelectSingleNode("ValueMaps"))
                {
                    ValueMap map = new ValueMap();
                    this._pInfoDic.GetProperties(map);
                    this._pInfoDic.SetModelValue(map, node4);
                    area.ValueMaps.Add(map);
                }
                this._Arealist.Add(area);
            }
        }

        private void SetModelValues(DataRow row, object obj)
        {
            if (this.CurrentArea != null)
            {
                foreach (Property property in this.CurrentArea.Propertys)
                {
                    this._pInfoDic.SetModelValue(obj, property.Name, this.CurrentArea.ValueMaps.toMapping(row[property.Column].ToString(), true));
                }
            }
        }

        private void SetModelValues(DataTable table, int RowOffset, object obj)
        {
            try
            {
                if (this.CurrentArea != null)
                {
                    foreach (Property property in this.CurrentArea.Propertys)
                    {
                        if ((property.Offset + RowOffset) < table.Rows.Count)
                        {
                            DataRow row = table.Rows[RowOffset + property.Offset];
                            this._pInfoDic.SetModelValue(obj, property.Name, this.CurrentArea.ValueMaps.toMapping(row[property.Column].ToString(), true));
                        }
                    }
                }
            }
            catch (Exception exception)
            {
                throw new Exception("错误：" + table.TableName + "/r/n" + exception.Message);
            }
        }

        public ObjectArea CurrentArea
        {
            get
            {
                return this._CurrentArea;
            }
        }

        public int RowOffset
        {
            get
            {
                return this._RowOffset;
            }
            set
            {
                this._RowOffset = value;
            }
        }

        public DataTable Table
        {
            get
            {
                return this._Table;
            }
            set
            {
                this._Table = value;
            }
        }
        #endregion

        public void GetBusinessObjectData(IBusinessObject ibo)
        {
            if (ibo == null) return;
            this._CurrentArea = this._Arealist.FirstOrDefault(c => c.Name == "BusinessObjectArea");
            GetPropertyValue(ibo);
            foreach (var item in ibo.ChildTables)
            {
                GetChildTableData(item);
            }
        }
        private void GetChildTableData(IChildTable ict)
        {
            if (ict == null) return;
            this._CurrentArea = this._Arealist.FirstOrDefault(c => c.Name == "BusinessObjectChildTableArea");
            GetPropertyValue(ict);
        }
        public void GetTableData(ITable it)
        {
            if (it == null) return;
            this._CurrentArea = this._Arealist.FirstOrDefault(c => c.Name == "TableArea");
            GetPropertyValue(it);
            GetFieldTitleData();
            foreach (var item in it.Fields)
            {
                GetFieldData(item);
            }
            GetEmptyData();
            if (true)
            {

            }
        }

        private void GetFieldTitleData()
        {
            DataRow dr = this.Table.NewRow();
            dr["F1"] = "字段名称";
            dr["F2"] = "字段描述";
            dr["F3"] = "模型属性名";
            dr["F4"] = "系统字段";
            dr["F5"] = "数据类型";
            dr["F7"] = "数据长度";
            dr["F8"] = "默认值";
            dr["F9"] = "可选值";
            dr["F11"] = "连接表";
            dr["F12"] = "备注";
            this.Table.Rows.Add(dr);
            dr = this.Table.NewRow();
            dr["F5"] = "类型";
            dr["F6"] = "结构";
            dr["F9"] = "存储值";
            dr["F10"] = "描述";
            this.Table.Rows.Add(dr);
        }

        private void GetEmptyData()
        {
            DataRow dr = this.Table.NewRow();
            this.Table.Rows.Add(dr);
            dr = this.Table.NewRow();
            this.Table.Rows.Add(dr);
            dr = this.Table.NewRow();
            this.Table.Rows.Add(dr);
        }
        private void GetFieldData(IField ifield)
        {
            if (ifield == null) return;

            this._CurrentArea = this._Arealist.FirstOrDefault(c => c.Name == "FieldArea");
            GetPropertyValue(ifield);
        }

        private void GetPropertyValue(object obj)
        {
            int max = this.CurrentArea.Propertys.Max(c => c.Offset);
            for (int i = 0; i <= max; i++)
            {
                DataRow dr = this.Table.NewRow();
                foreach (var item in this.CurrentArea.Propertys.Where(c => c.Offset == i))
                {
                    Type type = obj.GetType();
                    var property = type.GetProperty(item.Name);
                    if (property != null)
                    {
                        var value = property.GetValue(obj);
                        foreach (var valueMap in this.CurrentArea.ValueMaps)
                        {
                            if (Convert.ToString(value) == valueMap.Target)
                            {
                                value = valueMap.Source;
                                break;
                            }
                        }
                        dr[item.Column] = value;
                    }
                    else
                    {
                        dr[item.Column] = item.Name;
                    }
                }
                this.Table.Rows.Add(dr);
            }


        }

    }
}
