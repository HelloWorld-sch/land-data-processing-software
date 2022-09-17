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

                string[] tillCode = _landuses.Where(b => b.SecondType == "耕地").Select(c => c.Code).ToArray();
                string[] woodCode = _landuses.Where(b => b.SecondType == "林地").Select(c => c.Code).ToArray();
                string[] gradenCode = _landuses.Where(b => b.SecondType == "园地").Select(c => c.Code).ToArray();
                string[] reservoirCode = _landuses.Where(b => b.ThirdType == "水库水面").Select(c => c.Code).ToArray();
                string[] pondCode = _landuses.Where(b => b.ThirdType == "坑塘水面").Select(c => c.Code).ToArray();
                string[] ruralCode = _landuses.Where(b => b.ThirdType == "农村道路").Select(c => c.Code).ToArray();
                string[] ditchCode = _landuses.Where(b => b.ThirdType == "沟渠").Select(c => c.Code).ToArray();
                string[] facilityCode = _landuses.Where(b => b.ThirdType == "设施农用地").Select(c => c.Code).ToArray();
                string[] buildCode = _landuses.Where(b => b.FirstType == "建设用地").Select(c => c.Code).ToArray();
                string[] unuseCode = _landuses.Where(b => b.FirstType == "未利用地").Select(c => c.Code).ToArray();


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
                            耕地 = g.Where(d => tillCode.Contains(d.Field<string>("地类代码")))
                                .Sum(e => e.Field<int>("图斑面积")),
                            林地 = g.Where(d => woodCode.Contains(d.Field<string>("地类代码")))
                                .Sum(e => e.Field<int>("图斑面积")),
                            园地 = g.Where(d => gradenCode.Contains(d.Field<string>("地类代码")))
                                .Sum(e => e.Field<int>("图斑面积")),
                            水库水面 = g.Where(d => reservoirCode.Contains(d.Field<string>("地类代码")))
                                .Sum(e => e.Field<int>("图斑面积")),
                            坑塘水面 = g.Where(d => pondCode.Contains(d.Field<string>("地类代码")))
                                .Sum(e => e.Field<int>("图斑面积")),
                            农村道路 = g.Where(d => ruralCode.Contains(d.Field<string>("地类代码")))
                                .Sum(e => e.Field<int>("图斑面积")),
                            沟渠 = g.Where(d => ditchCode.Contains(d.Field<string>("地类代码")))
                                .Sum(e => e.Field<int>("图斑面积")),
                            设施农用地 = g.Where(d => facilityCode.Contains(d.Field<string>("地类代码")))
                                .Sum(e => e.Field<int>("图斑面积")),
                            建设用地 = g.Where(d => buildCode.Contains(d.Field<string>("地类代码")))
                                .Sum(e => e.Field<int>("图斑面积")),
                            未利用地 = g.Where(d => unuseCode.Contains(d.Field<string>("地类代码")))
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

                            double tillArea = GetRound(GetAreaWithUnit(q.耕地, q.田坎, unit));
                            double woodArea = GetRound(GetAreaWithUnit(q.林地, 0, unit));
                            double gradenArea = GetRound(GetAreaWithUnit(q.园地, 0, unit));
                            double reservoirArea = GetRound(GetAreaWithUnit(q.水库水面, 0, unit));
                            double pondArea = GetRound(GetAreaWithUnit(q.坑塘水面, 0, unit));
                            double ruralArea = GetRound(GetAreaWithUnit(q.农村道路, 0, unit));
                            double ditchArea = GetRound(GetAreaWithUnit(q.沟渠, 0, unit));
                            double facilityArea = GetRound(GetAreaWithUnit(q.设施农用地, 0, unit));
                            double buildArea = GetRound(GetAreaWithUnit(q.建设用地, 0, unit));
                            double unuseArea = GetRound(GetAreaWithUnit(q.未利用地, 0, unit));
                            double ridegArea = GetRound(GetAreaWithUnit(q.田坎, 0, unit));
                            double sumArea = tillArea + woodArea + gradenArea + reservoirArea + pondArea + ruralArea
                                + ditchArea + facilityArea + buildArea + unuseArea + ridegArea;
                            if (sumArea.Equals(0.0)) continue;

                            worksheet.InsertRow(currentRowIndex, 1);//插入行

                            ExcelRow currentRow = worksheet.Row(currentRowIndex);
                            currentRow.Style.Font.Size = 11;//字体大小
                            currentRow.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            currentRow.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            currentRow.Style.WrapText = true;
                            ExcelRange range = worksheet.Cells[currentRowIndex, 1, currentRowIndex, 17];
                            SetBorderStyle(range);

                            worksheet.Cells[currentRowIndex, 1].Value = i + 1;//序号
                            worksheet.Cells[currentRowIndex, 2].Value = q.户主;//户主
                            worksheet.Cells[currentRowIndex, 3].Value = q.地块编号;//地块编号
                            if (!tillArea.Equals(0.0))
                                worksheet.Cells[currentRowIndex, 4].Value = tillArea;//耕地
                            if (!woodArea.Equals(0.0))
                                worksheet.Cells[currentRowIndex, 5].Value = woodArea;//林地
                            if (!gradenArea.Equals(0.0))
                                worksheet.Cells[currentRowIndex, 6].Value = gradenArea;//园地
                            if (!reservoirArea.Equals(0.0))
                                worksheet.Cells[currentRowIndex, 7].Value = reservoirArea;//水库水面
                            if (!pondArea.Equals(0.0))
                                worksheet.Cells[currentRowIndex, 8].Value = pondArea;//坑塘水面
                            if (!ruralArea.Equals(0.0))
                                worksheet.Cells[currentRowIndex, 9].Value = ruralArea;//农村道路
                            if (!ditchArea.Equals(0.0))
                                worksheet.Cells[currentRowIndex, 10].Value = ditchArea;//沟渠
                            if (!facilityArea.Equals(0.0))
                                worksheet.Cells[currentRowIndex, 11].Value = facilityArea;//设施农用地
                            if (!ridegArea.Equals(0.0))
                                worksheet.Cells[currentRowIndex, 12].Value = ridegArea;//田坎
                            if (!buildArea.Equals(0.0))
                                worksheet.Cells[currentRowIndex, 13].Value = buildArea;//建设用地
                            if (!unuseArea.Equals(0.0))
                                worksheet.Cells[currentRowIndex, 14].Value = unuseArea;//未利用地
                            if (!sumArea.Equals(0.0))
                                worksheet.Cells[currentRowIndex, 15].Value = sumArea;//小计
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
