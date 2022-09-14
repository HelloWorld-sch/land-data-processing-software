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
        public void 土地调查成果确认表(string unit)
        {
            string methodName = System.Reflection.MethodBase.GetCurrentMethod().Name;
            try
            {
                log.Info("进入" + methodName);
                string templatePath = Path.Combine(Application.StartupPath, "Template\\土地调查成果确认表.xlsx");

                string[] selectCodes = _landuses.Select(b => b.Code).ToArray();

                string[] gddm = _landuses.Where(b => b.SecondType == "耕地").Select(c => c.Code).ToArray();
                string[] lddm = _landuses.Where(b => b.SecondType == "林地").Select(c => c.Code).ToArray();
                string[] yddm = _landuses.Where(b => b.SecondType == "园地").Select(c => c.Code).ToArray();

                var xzqList = Tools.GetXZQ(_table, "权属性质='集体'").Distinct();
                foreach (var xzq in xzqList)
                {
                    string city = xzq.City;
                    string county = xzq.County;

                    string filter = string.Format("市州='{0}' and 县='{1}' and 乡镇='{2}' and 村='{3}' and 组='{4}' and 户主 <>''",
                        city, county, xzq.Town, xzq.Village, xzq.Group);
                    var query = from b in _table.Select(filter).Where(b => selectCodes.Contains(b.Field<string>("地类代码")))
                        group b by b.Field<string>("户主")
                        into g
                        let tbs = g.Select(c => c.Field<string>("地块编号")).Select(k => k.ToString()).ToArray()
                        select new
                        {
                            户主 = g.Key,
                            地块编号 = string.Join(",", tbs),
                            耕地 = g.Where(d => gddm.Contains(d.Field<string>("地类代码")))
                                .Sum(e => e.Field<int>("图斑面积")),
                            林地 = g.Where(d => lddm.Contains(d.Field<string>("地类代码")))
                                .Sum(e => e.Field<int>("图斑面积")),
                            园地 = g.Where(d => yddm.Contains(d.Field<string>("地类代码")))
                                .Sum(e => e.Field<int>("图斑面积")),
                            田坎 = g.Sum(e => e.Field<int>("田坎面积"))
                        };

                    string dir = Path.Combine(_saveDir, methodName + "-" + _sylx, city, county);
                    if (!Directory.Exists(dir))
                        Directory.CreateDirectory(dir);
                    string saveExcelPath = Path.Combine(dir, xzq.Town + xzq.Village + xzq.Group + ".xlsx");

                    using (ExcelPackage package = new ExcelPackage(new FileInfo(templatePath)))
                    {
                        ExcelWorksheet worksheet = GetWorksheet(package, "Sheet1");

                        worksheet.Cells["A1"].Value = worksheet.Cells["A1"].Text.Replace("#水库名#", _reservoirName);
                        worksheet.Cells["A2"].Value = worksheet.Cells["A2"].Text.Replace("#区县名#", xzq.County).
                            Replace("#乡镇名#", xzq.Town).Replace("#村名#", xzq.Village).Replace("#组名#", xzq.Group);
                        worksheet.Cells["D3"].Value = worksheet.Cells["D3"].Text.Replace("#单位#", unit);

                        int i = 0;
                        int startIndex = 5;
                        foreach (var q in query)
                        {
                            int currentRowIndex = i + startIndex;

                            double gdArea = GetRound(GetAreaWithUnit(q.耕地, q.田坎, unit));
                            double ldArea = GetRound(GetAreaWithUnit(q.林地, 0, unit));
                            double ydArea = GetRound(GetAreaWithUnit(q.园地, 0, unit));
                            double sumArea = gdArea + ldArea + ydArea;
                            if (sumArea.Equals(0.0)) continue;

                            worksheet.InsertRow(currentRowIndex, 1);//插入行

                            ExcelRow currentRow = worksheet.Row(currentRowIndex);
                            currentRow.Style.Font.Size = 11;//字体大小
                            currentRow.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            currentRow.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            currentRow.Style.WrapText = true;
                            ExcelRange range = worksheet.Cells[currentRowIndex, 1, currentRowIndex, 9];
                            SetBorderStyle(range);

                            worksheet.Cells[currentRowIndex, 1].Value = i + 1;//序号
                            worksheet.Cells[currentRowIndex, 2].Value = q.户主;//户主
                            worksheet.Cells[currentRowIndex, 3].Value = q.地块编号;//地块编号
                            if (!gdArea.Equals(0.0))
                                worksheet.Cells[currentRowIndex, 4].Value = gdArea;//耕地
                            if (!ldArea.Equals(0.0))
                                worksheet.Cells[currentRowIndex, 5].Value = ldArea;//林地
                            if (!ydArea.Equals(0.0))
                                worksheet.Cells[currentRowIndex, 6].Value = ydArea;//园地
                            if (!sumArea.Equals(0.0))
                                worksheet.Cells[currentRowIndex, 7].Value = sumArea;//小计
                            i++;
                        }

                        ////如果不够28行，则增加
                        //int total = 28;
                        //if (i + 1 < total)
                        //{
                        //    worksheet.InsertRow(i+startIndex, total-i,startIndex);//插入行
                        //}
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

        public double GetAreaWithUnit(double spotArea, double ridgeArea, string unit)
        {
            var value = spotArea - ridgeArea;
            switch (unit)
            {
                case "亩":
                    return value * 0.0015;
                case "公顷":
                    return value * 0.0001;
                default:
                    return value;
            }
        }
    }
}
