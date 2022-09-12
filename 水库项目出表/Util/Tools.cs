using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Windows.Forms;
using 水库项目出表.Entity;

namespace 水库项目出表.Util
{
    public static class Tools
    {
        /// <summary>
        /// 读取三调土地利用分类
        /// </summary>
        /// <returns></returns>
        public static List<Landuse> ReadLanduses()
        {
            List<Landuse> landuses = new List<Landuse>();
            string tdPath = Path.Combine(Application.StartupPath, "三调三大类.txt");
            try
            {
                string[] lines = File.ReadAllLines(tdPath).Where(b => !string.IsNullOrEmpty(b)).ToArray();
                foreach (var line in lines)
                {
                    string[] arr = line.Split(',');

                    Landuse landuse = new Landuse();
                    landuse.FirstType = arr[0];
                    landuse.SecondType = arr[1];
                    landuse.ThirdType = arr[2];
                    landuse.Code = arr[3];
                    landuses.Add(landuse);
                }
            }
            catch (Exception e)
            {
                throw new Exception("读取文件失败,文件名:" + Path.GetFileName(tdPath));
            }

            return landuses;
        }
        /// <summary>
        /// 读取TreeView上土地利用分类勾选情况
        /// </summary>
        /// <param name="treeView"></param>
        /// <returns></returns>
        public static List<Landuse> GetLandusesFromTreeView(TreeView treeView)
        {
            List<Landuse> srcLanduses = ReadLanduses();

            List<Landuse> landusesTreeView = new List<Landuse>();
            var nodes1 = treeView.Nodes["root"].Nodes;
            foreach (TreeNode node1 in nodes1)
            {
                var nodes2 = node1.Nodes;
                foreach (TreeNode node2 in nodes2)
                {
                    var nodes3 = node2.Nodes;
                    foreach (TreeNode node3 in nodes3)
                    {
                        if (!node3.Checked) continue;
                        Landuse queryLanduse = srcLanduses
                            .FirstOrDefault(b => b.FirstType == node1.Text && b.SecondType == node2.Text && b.ThirdType == node3.Text);
                        string code = queryLanduse == null ? "" : queryLanduse.Code;

                        Landuse landuse = new Landuse();
                        landuse.FirstType = node1.Text;
                        landuse.SecondType = node2.Text;
                        landuse.ThirdType = node3.Text;
                        landuse.Code = code;

                        landusesTreeView.Add(landuse);
                    }
                }

            }

            return landusesTreeView;
        }
        /// <summary>
        /// 定义标准字段
        /// </summary>
        /// <returns></returns>
        public static string[] GetStandardFields()
        {
            string[] standardFields =
            {
                "Shape_Area","县", "乡镇", "村", "组", "地类名称", "地类代码", "权属性质", "功能分区", "户主", "地块编号", "使用类型", "身份证号码", "国有权属单位名称", "图斑编号","市州"
            };
            return standardFields;
        }
        /// <summary>
        /// 获取唯一值
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="dataTable"></param>
        /// <param name="colName"></param>
        /// <param name="filter"></param>
        /// <returns></returns>
        public static List<T> DistinctValues<T>(this DataTable dataTable, string colName,string filter)
        {
            
            var list = dataTable.Select(filter).Select(b => b.Field<T>(colName)).Distinct().ToList();
            return list;
        }

        public static List<XZQ> GetXZQ(DataTable table,string filter)
        {
            List<XZQ> xzqs=new List<XZQ>();
            var xzqdDistinct = table.Select(filter).Select(b => new
            {
                市州 = b.Field<string>("市州"),
                县 = b.Field<string>("县"),
                乡镇=b.Field<string>("乡镇"),
                村 = b.Field<string>("村"),
                组 = b.Field<string>("组")
            }).Distinct().OrderBy(c => c.市州 + c.县+c.乡镇+c.村+c.组);
            char[] invalidFileNameChars = Path.GetInvalidFileNameChars();
            foreach (var xzq in xzqdDistinct)
            {
                string shi = xzq.市州;
                string xian = xzq.县;
                string xz = xzq.乡镇;
                string cun = xzq.村;
                string zu = xzq.组;

                string[] xzqArr = { shi, xian };
                foreach (var xzqStr in xzqArr)
                {
                    if (shi.ToCharArray().Intersect(invalidFileNameChars).Any())
                        throw new Exception(xzqStr + "含有非法字符");
                }

                XZQ xzqEnt=new XZQ();
                xzqEnt.City = shi;
                xzqEnt.County = xian;
                xzqEnt.Town = xz;
                xzqEnt.Village = cun;
                xzqEnt.Group = zu;
                xzqs.Add(xzqEnt);
            }

            return xzqs;
        }

        public static Dictionary<string, int> GetDLWZ()
        {
            Dictionary<string,int> wzDic=new Dictionary<string, int>();

            string dlwzPath = Path.Combine(Application.StartupPath, "地类位置.txt");
            try
            {
                string[] lines = File.ReadAllLines(dlwzPath).Where(b => !string.IsNullOrEmpty(b)).ToArray();
                foreach (var line in lines)
                {
                    string[] arr = line.Split(',');
                    wzDic.Add(arr[0],Convert.ToInt32(arr[1]));
                }
            }
            catch (Exception e)
            {
                throw new Exception("读取文件失败,文件名:" + Path.GetFileName(dlwzPath));
            }

            return wzDic;
        }

        public static List<Sum> GetSumEntiyList()
        {
            List<Sum> sums=new List<Sum>();

            string hzwzPath = Path.Combine(Application.StartupPath, "汇总位置.txt");
            try
            {
                string[] lines = File.ReadAllLines(hzwzPath).Where(b => !string.IsNullOrEmpty(b)).ToArray();
                foreach (var line in lines)
                {
                    string[] arr = line.Split('|');
                    
                    Sum sum=new Sum();
                    sum.Index = Convert.ToInt32(arr[0]);
                    sum.SumColList = arr[1].Split(',').Select(b => Convert.ToInt32(b)).ToList();

                    sums.Add(sum);

                }
            }
            catch (Exception e)
            {
                throw new Exception("读取文件失败,文件名:" + Path.GetFileName(hzwzPath));
            }

            return sums;
        }

    }
}
