using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Text;
using 水库项目出表.Entity;
using 水库项目出表.Util;

namespace 水库项目出表.Attributes
{
    public class Data
    {
        private Logger log = Logger.GetLogger("Data");
        private DataTable _table;
        public Data(DataTable table)
        {
            _table = table;
        }
        /// <summary>
        /// 数据检查
        /// </summary>
        public void Check()
        {
            log.Info("进入数据检查");
            string[] qsxzAtts = new string[] { "集体", "国有" };
            //1、检查权属性质是否只有集体和国有两种类型的值
            List<string> qsxzExcept = _table.DistinctValues<string>("权属性质", "").Except(qsxzAtts).ToList();
            if (qsxzExcept.Count != 0)
                throw new Exception("权属性质字段含有除国有、集体以外的非法属性");

            //2、选择集体时，检查市州、县、乡镇、村、组是否完整,选择国有时，检查国有权属单位名称是否完整
            foreach (var qsxzAtt in qsxzAtts)
            {
                string filter = "权属性质='" + qsxzAtt + "'";
                if (qsxzAtt == "国有")
                {
                    bool checkGY = _table.DistinctValues<string>("国有权属单位名称", filter).Exists(b => b.Trim() == "");
                    if (checkGY)
                        throw new Exception("权属性质是" + qsxzAtt + "时,国有权属单位名称字段存在空属性");
                }
                else
                {
                    string[] xzqFields = new string[] { "市州", "县", "乡镇", "村", "组" };
                    foreach (var xzq in xzqFields)
                    {
                        bool checkXZQ = _table.DistinctValues<string>(xzq, filter).Exists(b => b.Trim() == "");
                        if (checkXZQ)
                            throw new Exception("权属性质是" + qsxzAtt + "时," + xzq + "字段存在空属性");
                    }
                }
            }

            //3、检查功能分区、使用类型、图斑编号、地类名称、地类代码是否完整
            string[] otherFields = new string[] { "功能分区", "使用类型", "图斑编号", "地类名称", "地类代码" };
            foreach (var fld in otherFields)
            {
                var distinctValues = _table.DistinctValues<object>(fld, "");
                bool checkOtherFields = distinctValues.Exists(b => b.ToString().Trim() == "");
                if (checkOtherFields)
                    throw new Exception(fld + "字段存在空属性");
                if (fld == "使用类型")
                {
                    if (!(distinctValues.Contains("永久") || distinctValues.Contains("临时")))
                        throw new Exception(fld + "字段只能填写永久或者临时");
                }
            }

            //4、检查地类名称、地类代码是否和三调匹配
            List<Landuse> landuses = Tools.ReadLanduses();
            DataTable resultTable = _table.AsDataView().ToTable(true, "地类名称", "地类代码");
            foreach (DataRow dataRow in resultTable.Rows)
            {
                string dlmc = dataRow.Field<string>("地类名称");
                string dldm = dataRow.Field<string>("地类代码");
                var boolExists = landuses.Count(b => b.ThirdType == dlmc && b.Code == dldm) == 0;
                if (boolExists)
                    throw new Exception("地类名称:" + dlmc + ",地类代码:" + dldm + ",不符合三调分类");
            }
        }
    }
}
