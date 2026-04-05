using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using 水库项目出表.Entity;
using 水库项目出表.Util;

namespace 水库项目出表.Attributes
{
    public partial class Export
    {
        public void 土地分类面积汇总表分市(string unit)
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
                worksheet.Cells["A1"].Value = worksheet.Cells["A1"].Text.Replace("#水库名#", _reservoirName) + "-分市";
                worksheet.Cells["CR2"].Value = worksheet.Cells["CR2"].Text.Replace("#单位#", unit);
                //隐藏第一列
                worksheet.Column(1).Hidden = true;

                //获取汇总数据
                DataTable tjTable = GetTableData(selectCodes, false, unit);

                DataTable resultTable = tjTable.Clone();
                
                var query = from b in tjTable.AsEnumerable()
                    // 新增逻辑：提取"组"列中的数字并转换为整数用于排序
                    let groupName = b.Field<string>("组")
                    let groupNumber = string.IsNullOrEmpty(groupName) 
                        ? int.MaxValue
                        : _digitRegex.Match(groupName) is Match match && match.Success 
                            ? int.Parse(match.Value) 
                            : int.MaxValue
                    orderby 
                        b.Field<string>("市州"), 
                        b.Field<string>("县") descending, 
                        b.Field<string>("权属性质") descending, 
                        b.Field<string>("乡镇") descending, 
                        b.Field<string>("村") descending, 
                        groupNumber ,  // 按数字排序（注意此处是降序）
                        b.Field<int>("权重")      // 原逻辑保持
                    select b;

                query.CopyToDataTable(resultTable, LoadOption.OverwriteChanges);
                
                
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
                        case 8:
                            worksheet.Cells[6 + i, 3].Value = row.Field<string>("市州") + "集体土地合计";
                            worksheet.Cells[6 + i, 3, 6 + i, 6].Merge = true;
                            break;
                        case 9:
                            worksheet.Cells[6 + i, 3].Value = row.Field<string>("市州") + "国有土地合计";
                            worksheet.Cells[6 + i, 3, 6 + i, 6].Merge = true;
                            break;
                        case 10:
                            worksheet.Cells[6 + i, 3].Value = row.Field<string>("市州") + "土地合计";
                            worksheet.Cells[6 + i, 3, 6 + i, 6].Merge = true;
                            break;
                    }

                }

                //合并功能分区、市州、县、乡镇单元格
                MergeCells(worksheet, 2);
                MergeCells(worksheet, 3);
                MergeCells(worksheet, 4);
                
                FormatAllDecimals(worksheet);

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

        private DataTable GetTableData(string[] selectCodes, bool fenqu, string unit)
        {
            DataTable pivotTable = PivotTable(selectCodes, fenqu, unit);
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
            //市集体土地合计
            DataTable sjtTable = STHJTable();
            //市国有土地合计
            DataTable sgyTable = SYHJTable();
            //市土地总计
            DataTable shjTable = SHJTable();
            pivotTable.Merge(cunTable);
            pivotTable.Merge(xzTable);
            pivotTable.Merge(xjtTable);
            pivotTable.Merge(xgyTable);
            pivotTable.Merge(xhjTable);
            pivotTable.Merge(sjtTable);
            pivotTable.Merge(sgyTable);
            pivotTable.Merge(shjTable);

            return pivotTable;
        }
    }
}
