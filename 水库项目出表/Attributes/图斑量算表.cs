using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using 水库项目出表.Util;

namespace 水库项目出表.Attributes
{
    public partial class Export
    {
        public void 图斑量算表(string unit)
        {
            string methodName = System.Reflection.MethodBase.GetCurrentMethod().Name;
            try
            {
                log.Info("进入" + methodName);
                string templatePath = Path.Combine(Application.StartupPath, "Template\\图斑量算表.xlsx");

                string[] selectCodes = _landuses.Select(b => b.Code).ToArray();

                var query = (from b in _table.AsEnumerable().Where(b => selectCodes.Contains(b.Field<string>("地类代码")))
                             let jt = b.Field<string>("乡镇") + b.Field<string>("村") + b.Field<string>("组")
                             let qsdw = b.Field<string>("权属性质") == "集体" ? jt : b.Field<string>("国有权属单位名称")
                             select new
                             {
                                 权属单位 = qsdw,
                                 图斑号 = b.Field<int>("图斑编号"),
                                 地类代码 = b.Field<string>("地类代码"),
                                 地类名称 = b.Field<string>("地类名称"),
                                 田坎面积 = b.Field<int>("田坎面积"),
                                 图斑净面积 = b.Field<int>("图斑面积") - b.Field<int>("田坎面积"),
                                 图斑面积 = b.Field<int>("图斑面积"),
                                 权属性质 = b.Field<string>("权属性质"),
                                 功能分区 = b.Field<string>("功能分区")
                             }).OrderBy(d => d.图斑号);


                string dir = Path.Combine(_saveDir, methodName + "-" + _sylx);
                if (!Directory.Exists(dir))
                    Directory.CreateDirectory(dir);
                string saveExcelPath = Path.Combine(dir, methodName + ".xlsx");

                using (ExcelPackage package = new ExcelPackage(new FileInfo(templatePath)))
                {
                    ExcelWorksheet worksheet = GetWorksheet(package, "Sheet1");

                    int i = 0;
                    int startIndex = 4;
                    foreach (var q in query)
                    {
                        int currentRowIndex = i + startIndex;

                        worksheet.InsertRow(currentRowIndex, 1); //插入行
                        var spotArea = GetRound(GetAreaWithUnit(q.图斑净面积, 0, unit));
                        var ridgeArea = GetRound(GetAreaWithUnit(q.田坎面积, 0, unit));
                        var allArea = GetRound(GetAreaWithUnit(q.图斑面积, 0, unit));


                        ExcelRow currentRow = worksheet.Row(currentRowIndex);
                        currentRow.Style.Font.Size = 10; //字体大小
                        currentRow.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        currentRow.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        currentRow.Style.WrapText = true;
                        ExcelRange range = worksheet.Cells[currentRowIndex, 1, currentRowIndex, 9];
                        SetBorderStyle(range);

                        worksheet.Cells[currentRowIndex, 1].Value = q.权属单位; //权属单位
                        worksheet.Cells[currentRowIndex, 2].Value = q.图斑号; //图斑号
                        worksheet.Cells[currentRowIndex, 3].Value = q.地类代码; //地类代码
                        worksheet.Cells[currentRowIndex, 4].Value = q.地类名称; //田坎面积
                        worksheet.Cells[currentRowIndex, 5].Value = ridgeArea; //田坎面积
                        worksheet.Cells[currentRowIndex, 6].Value = spotArea; //图斑净面积
                        worksheet.Cells[currentRowIndex, 7].Value = allArea; //图斑面积
                        worksheet.Cells[currentRowIndex, 8].Value = q.权属性质; //权属性质
                        worksheet.Cells[currentRowIndex, 9].Value = q.功能分区; //功能分区

                        i++;
                    }

                    worksheet.Cells["I2"].Value = worksheet.Cells["I2"].Text.Replace("#单位#", unit);

                    package.SaveAs(new FileInfo(saveExcelPath));
                }
            }
            catch (Exception e)
            {
                log.Error("导出" + methodName + "失败," + e.ToString());
                throw ;
            }
        }
    }
}
