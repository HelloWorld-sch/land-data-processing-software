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
        public void 土地分类面积汇总表分区(string unit)
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
                worksheet.Cells["A1"].Value = worksheet.Cells["A1"].Text.Replace("#水库名#", _reservoirName) + "-分区";

                //获取汇总数据
                DataTable tjTable = GetData(selectCodes, true, unit);

                DataTable resultTable = tjTable.Clone();
                (from b in tjTable.AsEnumerable() orderby b.Field<string>("功能分区"), b.Field<string>("市州"), b.Field<string>("县"), b.Field<string>("权属性质") descending, b.Field<string>("乡镇") descending, b.Field<string>("村") descending, b.Field<string>("组") descending, b.Field<int>("权重") select b).CopyToDataTable(resultTable, LoadOption.OverwriteChanges);

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
                MergeCells(worksheet, 1);
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
                throw;
            }
        }

    }
}
