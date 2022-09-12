using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Utilities;
using 水库项目出表.Attributes;
using 水库项目出表.Entity;
using 水库项目出表.Util;

namespace 水库项目出表
{
    public partial class Form1 : Form
    {
        private string mdbPath = "";
        Logger log = Logger.GetLogger("Form1");
        GisHelper gis = new GisHelper();
        public Form1()
        {
            InitializeComponent();
        }
        /// <summary>
        /// 打开mdb
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Title = "打开Personal Geodatabase";
            dlg.Filter = "Personal Geodatabase数据(*.mdb)|*.mdb";
            dlg.RestoreDirectory = true;
            if (dlg.ShowDialog() != DialogResult.OK)
                return;

            mdbPath = dlg.FileName;

            //获取图层
            gis.OpenDataSource(mdbPath);
            List<string> layerNames = gis.GetLayerNames();
            comboBox1.Items.Clear();
            foreach (var layerName in layerNames)
            {
                comboBox1.Items.Add(layerName);
            }

            comboBox1.SelectedIndex = 0;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                if (!File.Exists(mdbPath))
                    throw new Exception("mdb文件不存在");
                if(textBox1.Text.Trim()=="")
                    throw new Exception("项目名不能为空");
                //if (textBox2.Text.Trim() == "")
                //    throw new Exception("制表单位不能为空");
                if (comboBox1.Items.Count == 0) return;
                List<Landuse> landusesTreeView = Tools.GetLandusesFromTreeView(treeView1);//读取选择的地类
                if (landusesTreeView.Count == 0)
                    throw new Exception("没有选中三级节点");
                if (!checkBox1.Checked && !checkBox2.Checked && !checkBox3.Checked && !checkBox4.Checked && !checkBox5.Checked)
                    throw new Exception("请选择输出类型");

                string sylx = comboBox2.SelectedItem.ToString();//使用类型

                string layerName = comboBox1.SelectedItem.ToString();
                DataTable table = gis.GetData(layerName);//读取属性表
                if(table.Rows.Count==0)
                    throw new Exception("读取到0条属性记录");

                Data data=new Data(table);
                data.Check();//数据检查
               
                //添加保留面积字段
                DataColumn column1 = new DataColumn("面积亩", typeof(double));
                table.Columns.Add(column1);
                DataColumn column2 = new DataColumn("面积公顷", typeof(double));
                table.Columns.Add(column2);

                //新加字段赋值
                foreach (DataRow row in table.Rows)
                {
                    double area = row.Field<double>("Shape_Area");
                    int tbh = row.Field<int>("图斑编号");
                    double mu=Math.Round(area*0.0015, 2, MidpointRounding.AwayFromZero);
                    if (mu.Equals(0))
                        throw new Exception("图斑编号:" + tbh + ",面积转换成亩并保留两位小数后等于0");
                    double gq = Math.Round(area * 0.0001, 4, MidpointRounding.AwayFromZero);
                    if (gq.Equals(0))
                        throw new Exception("图斑编号:" + tbh + ",面积转换成公顷并保留四位小数后等于0");
                    row["面积亩"] = mu;
                    row["面积公顷"] = gq;
                }
                
                data.Adjustment();//面积平差

                string saveDir = Path.Combine(Application.StartupPath, "数据导出");
                if (!Directory.Exists(saveDir))
                    Directory.CreateDirectory(saveDir);

                //导出数据
                string reservoirName = textBox1.Text.Trim();//水库名称
                string zbdw = textBox2.Text.Trim();//制表单位
                ExportParameter parameter=new ExportParameter();
                parameter.DataSource = table;
                parameter.Landuses = landusesTreeView;
                parameter.ReservoirName = reservoirName;
                parameter.SaveDirectory = saveDir;
                parameter.Unit = zbdw;
                parameter.UseType = sylx;
                Export export = new Export(parameter);
                if(checkBox1.Checked)
                    export.土地调查成果确认表();
                if (checkBox2.Checked)
                    export.图斑量算表();
                if (checkBox3.Checked)
                    export.权属情况汇总表();
                if (checkBox4.Checked)
                    export.土地分类面积汇总表分县();
                if (checkBox5.Checked)
                    export.土地分类面积汇总表分区();
                
                MessageBox.Show("OK");
            }
            catch (Exception ex)
            {
                log.Error(ex.ToString());
                MessageBox.Show(ex.Message);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {

                treeView1.CheckBoxes = true;
                TreeNode node0 = new TreeNode("面积总计");
                node0.Name = "root";
                treeView1.Nodes.Add(node0);

                List<Landuse> landuses = Tools.ReadLanduses();
                foreach (var landuse in landuses)
                {
                    string first = landuse.FirstType;
                    string second = landuse.SecondType;
                    string third = landuse.ThirdType;

                    TreeNode node1 = new TreeNode(first);
                    node1.Name = node1.Text;
                    TreeNode node2 = new TreeNode(second);
                    node2.Name = node2.Text;
                    TreeNode node3 = new TreeNode(third);
                    node3.Name = node3.Text;

                    if (!node0.Nodes.ContainsKey(first))
                    {
                        node2.Nodes.Add(node3);
                        node1.Nodes.Add(node2);
                        node0.Nodes.Add(node1);
                    }
                    else
                    {
                        if (!node0.Nodes[first].Nodes.ContainsKey(second))
                        {
                            node2.Nodes.Add(node3);
                            node0.Nodes[first].Nodes.Add(node2);
                        }
                        else
                        {
                            node0.Nodes[first].Nodes[second].Nodes.Add(node3);
                        }
                    }
                }

                treeView1.ExpandAll();
                treeView1.SelectedNode = node0;

                comboBox2.SelectedIndex = 0;
            }
            catch (Exception exception)
            {
                log.Error(exception.ToString());
                MessageBox.Show(exception.Message);
                this.Close();
            }
        }

        private void treeView1_AfterCheck(object sender, TreeViewEventArgs e)
        {
            if (e.Action != TreeViewAction.Unknown)
            {
                CheckAllChildNodes(e.Node, e.Node.Checked);
                CheckParentNodes(e.Node);
            }
        }
        /// <summary>
        /// 选择所有子节点
        /// </summary>
        /// <param name="treeNode"></param>
        /// <param name="nodeChecked"></param>
        private void CheckAllChildNodes(TreeNode treeNode, bool nodeChecked)
        {

            foreach (TreeNode node in treeNode.Nodes)
            {
                node.Checked = nodeChecked;
                if (node.Nodes.Count > 0)
                {
                    this.CheckAllChildNodes(node, nodeChecked);
                }
            }
        }
        /// <summary>
        /// 根据子节点选择父节点
        /// </summary>
        /// <param name="treeNode"></param>
        private void CheckParentNodes(TreeNode treeNode)
        {
            if (treeNode.Parent != null)
            {
                if (treeNode.Parent.Nodes.Cast<TreeNode>().All(b => !b.Checked))
                {
                    treeNode.Parent.Checked = false;
                }
                else if (treeNode.Parent.Nodes.Cast<TreeNode>().Any(b=>b.Checked))
                {
                    treeNode.Parent.Checked = true;
                }

                CheckParentNodes(treeNode.Parent);
            }
        }
        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            gis.Dispose();
        }
    }
}
