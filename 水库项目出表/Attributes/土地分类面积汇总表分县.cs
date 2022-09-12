using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using 水库项目出表.Entity;
using 水库项目出表.Util;

namespace 水库项目出表.Attributes
{
    public partial class Export
    {
        private DataTable jitiTable = null;
        private DataTable guoyouTable = null;
        public void 土地分类面积汇总表分县()
        {
            string methodName = System.Reflection.MethodBase.GetCurrentMethod().Name;
            try
            {
                log.Info("进入" + methodName);
                string templatePath = Path.Combine(Application.StartupPath, "Template\\土地分类面积汇总表.xlsx");

                string[] selectCodes = _landuses.Select(b => b.Code).ToArray();

                ExcelPackage package = new ExcelPackage(new FileInfo(templatePath));
                ExcelWorksheet worksheet = GetWorksheet(package, "Sheet1");
                worksheet.Workbook.CalcMode = ExcelCalcMode.Automatic;

                //标题
                worksheet.Cells["A1"].Value = worksheet.Cells["A1"].Text.Replace("#水库名#", _reservoirName) + "-分县";
                //隐藏第一列
                worksheet.Column(1).Hidden = true;

                //获取汇总数据
                DataTable tjTable = GetData(selectCodes, false);

                DataTable resultTable = tjTable.Clone();
                (from b in tjTable.AsEnumerable() orderby b.Field<string>("市州"), b.Field<string>("县"), b.Field<string>("权属性质") descending, b.Field<string>("乡镇") descending, b.Field<string>("村") descending, b.Field<string>("组") descending, b.Field<int>("权重") select b).CopyToDataTable(resultTable,LoadOption.OverwriteChanges);

                string dir = Path.Combine(_saveDir, methodName + "-" + _sylx);
                if (!Directory.Exists(dir))
                    Directory.CreateDirectory(dir);
                string saveExcelPath = Path.Combine(dir, methodName + ".xlsx");

                //数据写入excel
                var rows = resultTable.Rows;
                for (int i = 0; i < rows.Count; i++)
                {
                    DataRow row = rows[i];
                    //权属区域
                    worksheet.Cells[6 + i, 1].Value = row["功能分区"];
                    worksheet.Cells[6 + i, 2].Value = row["市州"];
                    worksheet.Cells[6 + i, 3].Value = row["县"];
                    worksheet.Cells[6 + i, 4].Value = row["乡镇"];
                    worksheet.Cells[6 + i, 5].Value = row["村"];
                    worksheet.Cells[6 + i, 6].Value = row["组"];
                    worksheet.Cells[6 + i, 7].Value = row["权属性质"];
                    //各种地类
                    DiLeiToExcel(worksheet, 6 + i, row);
                    //地类汇总
                    DLHJ(worksheet, 6 + i);
                    //合计小计
                    int qz = row.Field<int>("权重");
                    switch (qz)
                    {
                        case 2:
                            worksheet.Cells[6 + i, 5].Value = "村小计";
                            worksheet.Cells[6 + i, 5, 6 + i, 6].Merge = true;
                            break;
                        case 3:
                            worksheet.Cells[6 + i, 4].Value = row.Field<string>("乡镇") + "小计";
                            worksheet.Cells[6 + i, 4, 6 + i, 6].Merge = true;
                            break;
                        case 4:
                            worksheet.Cells[6 + i, 4].Value = row.Field<string>("县") + "集体土地合计";
                            worksheet.Cells[6 + i, 4, 6 + i, 6].Merge = true;
                            break;
                        case 5:
                            worksheet.Cells[6 + i, 4, 6 + i, 6].Merge = true;
                            break;
                        case 6:
                            worksheet.Cells[6 + i, 4].Value = row.Field<string>("县") + "国有土地合计";
                            worksheet.Cells[6 + i, 4, 6 + i, 6].Merge = true;
                            break;
                        case 7:
                            worksheet.Cells[6 + i, 4].Value = row.Field<string>("县") + "土地合计";
                            worksheet.Cells[6 + i, 4, 6 + i, 6].Merge = true;
                            break;
                    }

                }

                //合并功能分区、市州、县、乡镇单元格
                MergeCells(worksheet, 2);
                MergeCells(worksheet, 3);
                MergeCells(worksheet, 4);

                //设边框
                SetBorderStyle(worksheet.Cells[6, 1, worksheet.Dimension.End.Row, worksheet.Dimension.End.Column]);

                package.SaveAs(new FileInfo(saveExcelPath));
                package.Dispose();
            }
            catch (Exception e)
            {
                log.Error("导出" + methodName + "失败," + e.ToString());
                throw ;
            }
        }

        private DataTable GetData(string[] selectCodes,bool fenqu)
        {
            DataTable pivotTable = PivotTable(selectCodes, fenqu);
            jitiTable = pivotTable.Clone();
            pivotTable.Select("权属性质='集体'").CopyToDataTable(jitiTable, LoadOption.OverwriteChanges);
            guoyouTable = pivotTable.Clone();
            pivotTable.Select("权属性质='国有'").CopyToDataTable(guoyouTable, LoadOption.OverwriteChanges);
            //村小计
            DataTable cunTable = CunTable();
            //乡镇小计
            DataTable xzTable = XZTable();
            //县集体土地合计
            DataTable xjtTable = JTHJTable();
            //县国有土地合计
            DataTable xgyTable = GYHJTable();
            //县土地总计
            DataTable xhjTable = XHJTable();
            pivotTable.Merge(cunTable);
            pivotTable.Merge(xzTable);
            pivotTable.Merge(xjtTable);
            pivotTable.Merge(xgyTable);
            pivotTable.Merge(xhjTable);

            return pivotTable;
        }
        /// <summary>
        /// 分组统计，行转列
        /// </summary>
        /// <param name="selectCodes"></param>
        /// <param name="colName"></param>
        /// <returns></returns>
        private DataTable PivotTable(string[] selectCodes,bool fenqu)
        {
            var query = (_table.AsEnumerable().Where(h => selectCodes.Contains(h.Field<string>("地类代码")))
                    .GroupBy(b => new
                    {
                        QSDW = b.Field<string>("权属性质") == "集体"
                            ? new
                            {
                                功能分区 = fenqu ?  b.Field<string>("功能分区"):"",
                                市州 = b.Field<string>("市州"),
                                县 = b.Field<string>("县"),
                                乡镇 = b.Field<string>("乡镇"),
                                村 = b.Field<string>("村"),
                                组 = b.Field<string>("组")
                                
                            }
                            : new
                            {
                                功能分区 = fenqu ?  b.Field<string>("功能分区"):"",
                                市州 = b.Field<string>("市州"),
                                县 = b.Field<string>("县"),
                                乡镇 = b.Field<string>("国有权属单位名称"),
                                村 = b.Field<string>("国有权属单位名称"),
                                组 = b.Field<string>("国有权属单位名称")
                            },
                        QSXZ = b.Field<string>("权属性质")
                    })
                    .Select(g => new
                    {
                        权属单位 = g.Key.QSDW,
                        权属性质 = g.Key.QSXZ,
                        地类列表 = (g.GroupBy(x => x.Field<string>("地类代码"))
                            .Select(y => new
                            {
                                代码 = y.Key,
                                面积 = y.Sum(m => m.Field<double>("面积公顷"))
                            }))
                    }))
                .OrderByDescending(a => a.权属性质)
                .ThenBy(f => f.权属单位.市州)
                .ThenBy(f => f.权属单位.县)
                .ThenBy(f => f.权属单位.乡镇)
                .ThenBy(f => f.权属单位.村)
                .ThenBy(f => f.权属单位.组);

            DataTable table = CreateTable(selectCodes);

            foreach (var q in query)
            {
                DataRow row = table.NewRow();
                row["功能分区"] = q.权属单位.功能分区;
                row["市州"] = q.权属单位.市州;
                row["县"] = q.权属单位.县;
                row["乡镇"] = q.权属单位.乡镇;
                row["村"] = q.权属单位.村;
                row["组"] = q.权属单位.组;
                row["权属性质"] = q.权属性质;
                row["权重"] = q.权属性质=="集体"?1:5;
                foreach (var selectCode in selectCodes)
                {
                    var dl = q.地类列表.FirstOrDefault(b => b.代码 == selectCode);
                    row[selectCode] = dl == null ? 0.0 : dl.面积;
                }
                table.Rows.Add(row);
            }

            return table;
        }
        /// <summary>
        /// 创建datatable
        /// </summary>
        /// <param name="selectCodes"></param>
        /// <returns></returns>
        private DataTable CreateTable(string[] selectCodes)
        {
            DataTable table = new DataTable();
            table.Columns.Add("功能分区", typeof(string));
            table.Columns.Add("市州", typeof(string));
            table.Columns.Add("县", typeof(string));
            table.Columns.Add("乡镇", typeof(string));
            table.Columns.Add("村", typeof(string));
            table.Columns.Add("组", typeof(string));
            table.Columns.Add("权属性质", typeof(string));
            table.Columns.Add("权重", typeof(int));
            foreach (var selectCode in selectCodes)
            {
                DataColumn column=new DataColumn(selectCode,typeof(double));
                column.Caption = "地类编码" + selectCode;
                table.Columns.Add(column);
            }

            return table;
        }
        /// <summary>
        /// 数据汇总
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="xzqFields"></param>
        /// <param name="weight"></param>
        /// <param name="qsxz"></param>
        /// <returns></returns>
        private DataTable SumTable(DataTable dt, string[] xzqFields, int weight,string qsxz)
        {
            DataTable table = dt.Clone();

            List<string> distinctValues = dt.DistinctValues<string>("功能分区", "");
            bool fenqu = distinctValues.TrueForAll(b=>!string.IsNullOrEmpty(b));

            var xzqFldList = xzqFields.ToList();
            if (!fenqu) xzqFldList.Remove("功能分区");
            var qsdwDistinctTable = dt.AsDataView().ToTable(true, xzqFldList.ToArray());
            foreach (DataRow dataRow in qsdwDistinctTable.Rows)
            {
                table.ImportRow(dataRow);
            }

            foreach (DataRow row in table.Rows)
            {
                string[] values = new string[xzqFields.Length];
                for (int i = 0; i < xzqFields.Length; i++)
                {
                    values[i] = row.Field<string>(xzqFields[i]);
                }

                string filter = "";
                if (fenqu)
                {
                    filter = string.Join(" and ", xzqFields.Zip(values, (first, second) => first + "='" + second + "'"));
                }
                else
                {
                    filter = string.Join(" and ", xzqFields.Skip(1).Zip(values.Skip(1), (first, second) => first + "='" + second + "'"));
                }

                foreach (DataColumn column in table.Columns)
                {
                    string colName = column.ColumnName;
                    string caption = column.Caption;
                    if (caption.StartsWith("地类编码"))
                    {
                        var s = dt.Compute("Sum([" + colName + "])", filter);
                        row[colName] = s;
                    }
                }

                row["权属性质"] =qsxz;
                row["权重"] = weight;
            }

            return table;
        }
        /// <summary>
        /// 村小计
        /// </summary>
        /// <returns></returns>
        private DataTable CunTable()
        {
            string[] xzqFields = new string[] {"功能分区", "市州", "县", "乡镇", "村" };
            return SumTable(jitiTable, xzqFields, 2,"集体");
        }
        /// <summary>
        /// 乡镇小计
        /// </summary>
        /// <returns></returns>
        private DataTable XZTable()
        {
            string[] xzqFields = new string[] { "功能分区", "市州", "县", "乡镇" };
            return SumTable(jitiTable, xzqFields, 3,"集体");
        }
        /// <summary>
        /// 县集体土地合计
        /// </summary>
        /// <returns></returns>
        private DataTable JTHJTable()
        {
            string[] xzqFields = new string[] { "功能分区", "市州", "县" };
            return SumTable(jitiTable, xzqFields, 4,"集体");
        }
        /// <summary>
        /// 县国有土地合计
        /// </summary>
        /// <returns></returns>
        private DataTable GYHJTable()
        {
            string[] xzqFields = new string[] { "功能分区", "市州", "县" };
            return SumTable(guoyouTable, xzqFields, 6,"国有");
        }
        /// <summary>
        /// 县土地合计
        /// </summary>
        /// <returns></returns>
        private DataTable XHJTable()
        {
            string[] xzqFields = new string[] { "功能分区", "市州", "县" };
            DataTable jtTable = SumTable(jitiTable, xzqFields, 7,"");
            DataTable gyTable = SumTable(guoyouTable, xzqFields, 7,"");
            jtTable.Merge(gyTable,true);
            return SumTable(jtTable, xzqFields, 7,"");
        }

        private void DLHJ(ExcelWorksheet worksheet, int rowIndex)
        {
            foreach (var sum in _sums)
            {
                int colIndex = sum.Index;
                List<int> sumColList = sum.SumColList;
                var cells = sumColList.Select(b => "R" + rowIndex + "C" + b).ToArray();
                string formula = "=sum(" + string.Join(",", cells) + ")";
                worksheet.Cells[rowIndex, colIndex].FormulaR1C1 = formula;
            }
        }
        private void DiLeiToExcel(ExcelWorksheet worksheet, int rowIndex,DataRow row)
        {
            DataTable dt = row.Table;
            foreach (DataColumn column in dt.Columns)
            {
                string colName = column.ColumnName;
                string caption = column.Caption;
                if (caption.StartsWith("地类编码"))
                {
                    worksheet.Cells[rowIndex, _dlwzDic[colName]].Value = row[colName];
                }
            }
        }

        private void MergeCells(ExcelWorksheet worksheet, int colIndex)
        {
            int startRow = 6;
            for (int j = 1; j < worksheet.Dimension.End.Row; j++)
            {
                var curCell = worksheet.Cells[6 + j, colIndex];
                var preCell = worksheet.Cells[6 + j - 1, colIndex];

                if (curCell.Text != preCell.Text)
                {
                    if (curCell.Merge && preCell.Merge) continue;
                    if (!curCell.Merge && preCell.Merge)
                    {
                        startRow = 6 + j;
                        continue;
                    }

                    int endRow = 6 + j - 1;
                    worksheet.Cells[startRow, colIndex, endRow, colIndex].Merge = true;
                    if (j > 1)
                    {
                        if (curCell.Merge)
                        {
                            j++;
                            startRow = 6 + j;
                        }
                        else
                        {
                            startRow = 6 + j;
                        }
                    }
                }
            }
        }
        private string GetExcelColumnName(int index)
        {
            var columnName = string.Empty;
            while (index > 0)
            {
                var modulo = (index - 1) % 26;
                columnName = Convert.ToChar(65 + modulo) + columnName;
                index = (index - modulo) / 26;
            }
            return columnName;
        }
    }
}
