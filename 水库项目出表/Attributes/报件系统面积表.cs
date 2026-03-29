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
        public void 报件系统面积表(string unit)
        {
            string methodName = System.Reflection.MethodBase.GetCurrentMethod().Name;
            try
            {
                log.Info("进入" + methodName);
                string templatePath = Path.Combine(Application.StartupPath, "Template\\报件系统面积表.xlsx");

                string[] selectCodes = _landuses.Select(b => b.Code).ToArray();

                ExcelPackage package = new ExcelPackage(new FileInfo(templatePath));
                ExcelWorksheet worksheet = GetWorksheet(package, "Sheet1");
                worksheet.Workbook.CalcMode = ExcelCalcMode.Automatic;

                //获取汇总数据
                DataTable tjTable = PivotTable(selectCodes, false, unit);

                DataTable resultTable = tjTable.Clone();
                (from b in tjTable.AsEnumerable() orderby b.Field<string>("市州"), b.Field<string>("县"), b.Field<string>("权属性质") descending, b.Field<string>("乡镇") descending, b.Field<string>("村") descending, b.Field<string>("组") descending, b.Field<int>("权重") select b).CopyToDataTable(resultTable,LoadOption.OverwriteChanges);

                string dir = Path.Combine(_saveDir, methodName + "-" + _sylx);
                if (!Directory.Exists(dir))
                    Directory.CreateDirectory(dir);
                string saveExcelPath = Path.Combine(dir, methodName + ".xlsx");

                //数据写入excel
                var rows = resultTable.Rows;
                _dlwzDic = Tools.GetDLWZ(ExcelTypeEnum.FilingSystem);
                _sums = Tools.GetSumEntiyList(ExcelTypeEnum.FilingSystem);
                for (int i = 0; i < rows.Count; i++)
                {
                    DataRow row = rows[i];
                    //权属区域
                    worksheet.Cells[3 + i, 1].Value = row["县"];
                    worksheet.Cells[3 + i, 2].Value = row["乡镇"];
                    worksheet.Cells[3 + i, 3].Value = row["村"];
                    worksheet.Cells[3 + i, 4].Value = row["组"];
                    worksheet.Cells[3 + i, 5].Value = row["权属性质"];
                    //各种地类
                    foreach (var item in _dlwzDic)
                    {
                        var columnList = item.Key.Split('|').ToList();
                        object value;
                        if (columnList.Count > 1)
                        {
                            double sum = 0;
                            foreach (var column in columnList)
                            {
                                if(row[column] is double)
                                    sum += Convert.ToDouble(row[column]);
                            }
                            value = sum == 0 ? new DataColumn().DefaultValue : sum;
                        }
                        else
                        {
                            value = row[columnList[0]];
                        }
                        worksheet.Cells[3 + i, item.Value].Value = value;
                    }
                    //地类汇总
                    DLHJ(worksheet, 3 + i);
                }
                
                FormatAllDecimals(worksheet);

                //设边框
                SetBorderStyle(worksheet.Cells[3, 1, worksheet.Dimension.End.Row, worksheet.Dimension.End.Column]);
                package.SaveAs(new FileInfo(saveExcelPath));
                package.Dispose();
            }
            catch (Exception e)
            {
                log.Error("导出" + methodName + "失败," + e.ToString());
                throw ;
            }
        }
    }
}
