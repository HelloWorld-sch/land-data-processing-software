using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using OSGeo.OGR;
using OSGeo.OSR;
using OSGeo.GDAL;

namespace 水库项目出表.Util
{
    public class GisHelper : IDisposable
    {
        static Logger log = Logger.GetLogger("GisHelper");
        private DataSource _dataSource;
        public GisHelper()
        {
            //注册驱动
            Ogr.RegisterAll();
            //Gdal.AllRegister();
            log.Info("gdal驱动注册完成");
        }
        /// <summary>
        /// 打开数据源
        /// </summary>
        /// <param name="vectorFile"></param>
        public void OpenDataSource(string vectorFile)
        {
            try
            {
                DataSource ds = Ogr.Open(vectorFile, 0);
                _dataSource = ds;
                log.Info("打开mdb成功");
            }
            catch (Exception e)
            {
                throw new Exception("打开数据源失败");
            }
        }
        /// <summary>
        /// 获取图层名
        /// </summary>
        /// <returns></returns>
        public List<string> GetLayerNames()
        {
            //获取数据源
            if (_dataSource == null)
                throw new Exception("获取数据源失败");
            int iLayerCount = _dataSource.GetLayerCount();
            if (iLayerCount == 0)
                throw new Exception("没有找到图层");

            //循环获取图层
            List<string> layerNames = new List<string>();
            for (int i = 0; i < iLayerCount; i++)
            {
                Layer oLayer = _dataSource.GetLayerByIndex(i + 1);
                if (oLayer == null) continue;
                layerNames.Add(oLayer.GetName());
                oLayer.Dispose();
            }
            log.Info("获取到" + layerNames.Count + "个可用图层");

            return layerNames;
        }
        /// <summary>
        /// 根据图层名获取属性
        /// </summary>
        /// <param name="layerName"></param>
        /// <returns></returns>
        public DataTable GetData(string layerName)
        {
            log.Info("读取属性表");

            DataTable dataTable = new DataTable();

            Feature oFeature = null;
            Layer oLayer = _dataSource.GetLayerByName(layerName);
            if (oLayer == null)
                throw new Exception("获取图层" + layerName + "失败");

            //获取属性表结构
            log.Info("获取属性表结构");
            string[] standardFields = Tools.GetStandardFields();
            FeatureDefn oDefn = oLayer.GetLayerDefn();
            int iFieldCount = oDefn.GetFieldCount();
            string[] attFields = new string[iFieldCount];
            for (int iAttr = 0; iAttr < iFieldCount; iAttr++)
            {
                FieldDefn oField = oDefn.GetFieldDefn(iAttr);
                string fieldName = oField.GetName();
                attFields[iAttr] = fieldName;
                oField.Dispose();
            }
            string[] resultArray = standardFields.Except(attFields,StringComparer.OrdinalIgnoreCase).ToArray();
            if(resultArray.Length!=0)
                throw new Exception("缺少字段:"+string.Join(",",resultArray));
            foreach (var standardField in standardFields)
            {
                Type fieldType;
                switch (standardField)
                {
                    case "图斑面积":
                    case "田坎面积":
                        fieldType = typeof(int);
                        break;
                    case "图斑编号":
                        fieldType = typeof(int);
                        break;
                    default:
                        fieldType = typeof(string);
                        break;
                }
                DataColumn column = new DataColumn(standardField, fieldType);
                dataTable.Columns.Add(column);
            }

            log.Info("获取每条属性记录");
            oLayer.ResetReading();
            while ((oFeature = oLayer.GetNextFeature()) != null)
            {
                DataRow row = dataTable.NewRow();
                // 获取要素中的属性表内容
                foreach (var standardField in standardFields)
                {
                    object value = null;
                    if (standardField == "图斑面积" || standardField == "田坎面积")
                    {
                        value = oFeature.GetFieldAsInteger(standardField);
                    }
                    else
                    {
                        value = oFeature.GetFieldAsString(standardField);
                    }
                    row[standardField] = value;
                }
                dataTable.Rows.Add(row);
            }

            oDefn.Dispose();
            oLayer.Dispose();
            log.Info("属性表读取完成");
            return dataTable;
        }

        public void Dispose()
        {
            if (_dataSource != null)
                _dataSource.Dispose();
        }
    }
}
