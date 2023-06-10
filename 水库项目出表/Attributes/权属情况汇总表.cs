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
using 水库项目出表.Util;

namespace 水库项目出表.Attributes
{
    public partial class Export
    {
        public void 权属情况汇总表(string unit)
        {
            string methodName = System.Reflection.MethodBase.GetCurrentMethod().Name;
            try
            {
                log.Info("进入" + methodName);
                string templatePath = Path.Combine(Application.StartupPath, "Template\\权属情况汇总表.xlsx");

                string[] selectCodes = _landuses.Select(b => b.Code).ToArray();
                string[] gddm = _landuses.Where(b => b.SecondType == "耕地").Select(c => c.Code).ToArray();


                var xzqList = Tools.GetXZQ(_table, "").Select(a => new { City = a.City, County = a.County }).Distinct();
                foreach (var xzq in xzqList)
                {
                    string city = xzq.City;
                    string county = xzq.County;

                    var query = (from b in _table.Select("市州='" + city + "' and 县='" + county + "'")
                        let jt = b.Field<string>("乡镇") + b.Field<string>("村") + b.Field<string>("组")
                        group b by new
                        {
                            QSDW = b.Field<string>("权属性质") == "集体" ? jt : b.Field<string>("国有权属单位名称"),
                            QSXZ = b.Field<string>("权属性质")
                        }
                        into g
                        select new
                        {
                            土地权利人 = g.Key.QSDW,
                            权属性质 = g.Key.QSXZ,
                            拟占土地面积 = g.Where(c => selectCodes.Contains(c.Field<string>("地类代码"))).Sum(d => d.Field<int>("图斑面积")),
                            耕地 = g.Where(c => gddm.Contains(c.Field<string>("地类代码"))).Sum(d => d.Field<int>("图斑面积")),
                            田坎 = g.Where(c => gddm.Contains(c.Field<string>("地类代码"))).Sum(d => d.Field<int>("田坎面积"))
                        }).OrderByDescending(p => p.权属性质).ThenBy(o => o.土地权利人);

                    string dir = Path.Combine(_saveDir, methodName + "-" + _sylx, city, county);
                    if (!Directory.Exists(dir))
                        Directory.CreateDirectory(dir);
                    string saveExcelPath = Path.Combine(dir,methodName+ ".xlsx");

                    using (ExcelPackage package = new ExcelPackage(new FileInfo(templatePath)))
                    {
                        ExcelWorksheet worksheet = GetWorksheet(package, "Sheet1");
                        worksheet.Workbook.CalcMode = ExcelCalcMode.Automatic;

                        int i = 0;
                        int startIndex = 5;
                        double sumTillArea = 0;
                        double sumLandArea = 0;
                        foreach (var q in query)
                        {
                            int currentRowIndex = i + startIndex;
                            var tillArea = GetRound(GetAreaWithUnit(q.耕地, q.田坎, unit));
                            var landArea = GetRound(GetAreaWithUnit(q.拟占土地面积, 0, unit));
                            sumTillArea += tillArea;
                            sumLandArea += landArea;

                            worksheet.InsertRow(currentRowIndex, 1); //插入行

                            ExcelRow currentRow = worksheet.Row(currentRowIndex);
                            currentRow.Style.Font.Size = 10; //字体大小
                            currentRow.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            currentRow.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            currentRow.Style.WrapText = true;
                            ExcelRange range = worksheet.Cells[currentRowIndex, 1, currentRowIndex, 8];
                            SetBorderStyle(range);

                            worksheet.Cells[currentRowIndex, 1].Value = i + 1; //序号
                            worksheet.Cells[currentRowIndex, 2].Value = q.土地权利人; //土地权利人
                            worksheet.Cells[currentRowIndex, 3].Value = q.权属性质; //权属性质
                            worksheet.Cells[currentRowIndex, 7].Value = landArea; //拟占土地面积
                            if (!tillArea.Equals(0.0))
                                worksheet.Cells[currentRowIndex, 8].Value = tillArea; //耕地
                            i++;
                        }
                        var lastRowIndex = i + startIndex;
                        worksheet.Cells["A2"].Value = worksheet.Cells["A2"].Text.Replace("#水库名#", _reservoirName);
                        worksheet.Cells["F2"].Value = worksheet.Cells["F2"].Text.Replace("#区县名#", county);
                        worksheet.Cells["H2"].Value = worksheet.Cells["H2"].Text.Replace("#单位#", unit);
                        worksheet.Cells["G" + (lastRowIndex + 1)].Value = worksheet.Cells["G" + (lastRowIndex + 1)].Text.Replace("#制表单位#", _zbdw);
                        worksheet.Cells["G" + lastRowIndex].Value = sumLandArea;
                        worksheet.Cells["H" + lastRowIndex].Value= sumTillArea;
                        package.SaveAs(new FileInfo(saveExcelPath));
                    }
                }
            }
            catch (Exception e)
            {
                log.Error("导出" + methodName + "失败," + e.ToString());
                throw ;
            }
        }

        public double GetRound(double value)
        {
            return Math.Round(value, 2, MidpointRounding.AwayFromZero);
        }
    }
}
